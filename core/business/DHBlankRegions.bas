Attribute VB_Name = "DHBlankRegions"
Option Explicit

Private Sub AddRegion(ByVal region As Range, ByVal preview As Boolean, ByVal kind As String, ByVal tableName As String, _
    ByVal result As Collection, ByRef scanned As Double)
    Dim spans As Collection, table As ListObject, span As Variant
    scanned = scanned + region.CountLarge
    If scanned > DH_MAX_CELLS Then Err.Raise vbObjectError + 758, , "实际扫描区域合计超过100万单元格，请分批处理。"
    ' The existing strict planner remains the default used by duplicate deletion.
    Set spans = DH_PlanBlankRows(region, table, preview, kind <> "表格", kind = "整行")
    For Each span In spans
        result.Add Array(span(0), span(1), region.Column, region.Column + region.Columns.Count - 1, kind, tableName)
        If result.Count > DH_MAX_AREAS Then Err.Raise vbObjectError + 760, , "空白记录分布过于碎片化，请分批处理。"
    Next span
End Sub

Private Sub SortSpans(ByRef values() As Variant, ByVal first As Long, ByVal last As Long)
    Dim left As Long, right As Long, pivot As Long, item As Variant, current As Variant
    left = first: right = last: current = values((first + last) \ 2): pivot = current(0)
    Do While left <= right
        Do
            current = values(left)
            If current(0) >= pivot Then Exit Do
            left = left + 1
        Loop
        Do
            current = values(right)
            If current(0) <= pivot Then Exit Do
            right = right - 1
        Loop
        If left <= right Then
            item = values(left): values(left) = values(right): values(right) = item
            left = left + 1: right = right - 1
        End If
    Loop
    If first < right Then SortSpans values, first, right
    If left < last Then SortSpans values, left, last
End Sub

Private Function Ordered(ByVal spans As Collection) As Collection
    Dim result As New Collection, values() As Variant, i As Long
    If spans.Count = 0 Then Set Ordered = result: Exit Function
    ReDim values(1 To spans.Count)
    For i = 1 To spans.Count: values(i) = spans(i): Next i
    SortSpans values, 1, spans.Count
    For i = 1 To spans.Count: result.Add values(i): Next i
    Set Ordered = result
End Function

Public Function DH_BlankPlanMode(ByVal plan As Collection) As String
    Dim span As Variant, cellsUp As Boolean, tables As Boolean, rows As Boolean
    For Each span In plan
        Select Case span(4)
            Case "整行": rows = True
            Case "表格": tables = True
            Case Else: cellsUp = True
        End Select
    Next span
    If rows Then DH_BlankPlanMode = "显式整行选择：仅删除整张工作表中完全空白的行"
    If cellsUp Then DH_BlankPlanMode = "普通区域：仅所选列单元格上移，左右未选列保留；同列下方内容会上移"
    If tables Then
        If Len(DH_BlankPlanMode) > 0 Then DH_BlankPlanMode = DH_BlankPlanMode & vbCrLf
        DH_BlankPlanMode = DH_BlankPlanMode & "超级表：按各表完整记录列清理数据行，保护表头与汇总"
    End If
End Function

' Plan item: firstRow,lastRow,firstColumn,lastColumn,mode,tableName.
' All-table selection is partitioned into ListObjects and safe ordinary strips.
' Every column occupied by a ListObject is reserved, including tables below the
' selected rows: a cell shift must never cut a table further down the worksheet.
Public Function DH_PlanBlankDeletion(ByVal selected As Range, Optional ByVal preview As Boolean = False, _
    Optional ByRef notes As String = "") As Collection
    Dim result As New Collection, ws As Worksheet, wb As Workbook, bounds As Range, part As Range, last As Range
    Dim table As ListObject, blocked() As Boolean, firstCol As Long, lastCol As Long, firstRow As Long, lastRow As Long
    Dim c As Long, first As Long, left As Long, right As Long, tail As Long, scanned As Double, skipped As Long
    Dim wholeColumns As Boolean, entireRows As Boolean
    notes = ""
    If selected Is Nothing Then Err.Raise vbObjectError + 750, , "请先选择要整理的区域。"
    Set ws = selected.Parent: Set wb = ws.Parent
    DH_RequireWorkbook wb
    If ws.ProtectContents Then Err.Raise vbObjectError + 751, , "工作表受保护，无法整理。"
    If Application.ActiveWindow Is Nothing Then Err.Raise vbObjectError + 752, , "没有活动窗口。"
    If Application.ActiveWindow.SelectedSheets.Count <> 1 Then Err.Raise vbObjectError + 753, , "请先取消工作表分组。"
    If selected.Areas.Count <> 1 Then Err.Raise vbObjectError + 754, , "请使用连续矩形选区，可同时覆盖并排表。"
    Set bounds = DH_ContentRange(ws)
    wholeColumns = (selected.Rows.Count = ws.Rows.Count)
    entireRows = (selected.Columns.Count = ws.Columns.Count And Not wholeColumns And ws.ListObjects.Count = 0)
    firstCol = selected.Column: lastCol = firstCol + selected.Columns.Count - 1
    firstRow = selected.Row: lastRow = firstRow + selected.Rows.Count - 1
    If selected.Columns.Count = ws.Columns.Count Then
        If bounds Is Nothing Then
            firstCol = 1: lastCol = 1
        Else
            firstCol = bounds.Column: lastCol = firstCol + bounds.Columns.Count - 1
        End If
        For Each table In ws.ListObjects
            If table.Range.Column < firstCol Then firstCol = table.Range.Column
            tail = table.Range.Column + table.Range.Columns.Count - 1
            If tail > lastCol Then lastCol = tail
        Next table
    End If
    ReDim blocked(1 To ws.Columns.Count)
    For Each table In ws.ListObjects
        left = table.Range.Column: right = left + table.Range.Columns.Count - 1
        For c = left To right: blocked(c) = True: Next c
        Set part = Application.Intersect(selected, table.Range)
        If Not part Is Nothing Then
            If selected.Column <= left And selected.Column + selected.Columns.Count - 1 >= right Then
                AddRegion part, preview, "表格", table.Name, result, scanned
            Else
                skipped = skipped + 1
            End If
        End If
    Next table
    c = firstCol
    Do While c <= lastCol
        If Not blocked(c) Then
            first = c
            Do While c <= lastCol
                If blocked(c) Then Exit Do
                c = c + 1
            Loop
            tail = lastRow
            If wholeColumns Then
                Set part = ws.Range(ws.Cells(firstRow, first), ws.Cells(lastRow, c - 1))
                Set last = part.Find(What:="*", After:=part.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                    SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
                If last Is Nothing Then tail = firstRow - 1 Else tail = last.Row
            End If
            If tail >= firstRow Then
                Set part = ws.Range(ws.Cells(firstRow, first), ws.Cells(tail, c - 1))
                AddRegion part, preview, IIf(entireRows, "整行", "区域"), "", result, scanned
            End If
        Else
            c = c + 1
        End If
    Loop
    If skipped > 0 Then notes = "有" & skipped & "个超级表仅选中部分列，已跳过；请覆盖该表全部记录列后再整理。"
    If ws.ListObjects.Count > 0 Then
        If Len(notes) > 0 Then notes = notes & vbCrLf
        notes = notes & "超级表所在列的表外普通区域保留，防止上移切断该列上方或下方的超级表。"
    End If
    If entireRows Then notes = "整行模式会检查所有有数据的列；旁列非空的行保留。仅整理左表时请框选左表列。"
    Set DH_PlanBlankDeletion = Ordered(result)
End Function
