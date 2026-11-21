Attribute VB_Name = "DHBlankRows"
Option Explicit

' Plan from bulk arrays. No deletion happens here; UI can preview this plan.
Public Function DH_PlanBlankRows(ByVal selected As Range, ByRef table As ListObject, _
    Optional ByVal previewSafeMerges As Boolean = False, Optional ByVal localColumns As Boolean = False, Optional ByVal wholeRowMetadata As Boolean = False) As Collection
    Dim wb As Workbook, ws As Worksheet, bounds As Range, rng As Range, overlap As Range, item As ListObject
    Dim values As Variant, formulas As Variant, merged As Variant, plan As Collection
    Dim firstRow As Long, lastRow As Long, firstCol As Long, lastCol As Long, r As Long, c As Long
    Dim emptyRow As Boolean, runFirst As Long, runLast As Long, absoluteRow As Long
    Dim mergeAreas As Collection, mergeArea As Range, mergedRows As Object, mergeValue As Variant, mergeReason As String
    Dim hasMerges As Boolean, noted As Object, guardMerges As Boolean, span As Variant
    Set plan = New Collection
    Set table = Nothing
    If selected Is Nothing Then Err.Raise vbObjectError + 750, , "请先选择要检查空行的区域。"
    Set ws = selected.Parent: Set wb = ws.Parent
    DH_RequireWorkbook wb
    If ws.ProtectContents Then Err.Raise vbObjectError + 751, , "工作表受保护，不能删除空行。"
    If Application.ActiveWindow Is Nothing Then Err.Raise vbObjectError + 752, , "没有活动窗口。"
    If Application.ActiveWindow.SelectedSheets.Count <> 1 Then Err.Raise vbObjectError + 753, , "请先取消工作表分组。"
    If selected.Areas.Count <> 1 Then Err.Raise vbObjectError + 754, , "删除空行请使用一个连续矩形选区。"

    For Each item In ws.ListObjects
        Set overlap = Application.Intersect(selected, item.Range)
        If Not overlap Is Nothing Then
            If Not table Is Nothing Then Err.Raise vbObjectError + 755, , "一次只能处理一个 Excel 表格。"
            Set table = item
        End If
    Next item
    If Not table Is Nothing Then
        Set overlap = Application.Intersect(selected, table.Range)
        If overlap.CountLarge <> selected.CountLarge Or selected.Column <> table.Range.Column Or _
            selected.Columns.Count <> table.Range.Columns.Count Then
            Err.Raise vbObjectError + 756, , "请在该 Excel 表格内选中全部数据列；表头、汇总行会自动保留。"
        End If
        If table.DataBodyRange Is Nothing Then Set DH_PlanBlankRows = plan: Exit Function
        Set rng = Application.Intersect(selected, table.DataBodyRange)
        If rng Is Nothing Then Set DH_PlanBlankRows = plan: Exit Function
    Else
        Set bounds = DH_ContentRange(ws)
        If bounds Is Nothing Then
            If selected.Rows.Count = ws.Rows.Count Then Set DH_PlanBlankRows = plan: Exit Function
            Set rng = selected
        Else
            firstCol = bounds.Column
            lastCol = bounds.Column + bounds.Columns.Count - 1
            If Not localColumns And (selected.Column > firstCol Or selected.Column + selected.Columns.Count - 1 < lastCol) Then
                Err.Raise vbObjectError + 757, , "请覆盖本工作表全部有数据的列，避免误删选区外的数据。"
            End If
            If localColumns Then firstCol = selected.Column: lastCol = selected.Column + selected.Columns.Count - 1
            firstRow = selected.Row
            lastRow = selected.Row + selected.Rows.Count - 1
            ' Whole-column selections omit Excel's unused million-row tail;
            ' explicitly bounded selections include their selected trailing blanks.
            If selected.Rows.Count = ws.Rows.Count Then lastRow = bounds.Row + bounds.Rows.Count - 1
            If firstRow > lastRow Then Set DH_PlanBlankRows = plan: Exit Function
            Set rng = ws.Range(ws.Cells(firstRow, firstCol), ws.Cells(lastRow, lastCol))
        End If
    End If
    If rng.CountLarge > DH_MAX_CELLS Then Err.Raise vbObjectError + 758, , "删除空行扫描超过100万单元格，请分批处理。"
    firstRow = rng.Row: lastRow = firstRow + rng.Rows.Count - 1
    merged = rng.MergeCells
    hasMerges = IsNull(merged)
    If Not hasMerges Then hasMerges = CBool(merged)
    If hasMerges And Not previewSafeMerges And Not localColumns Then
        ' Merged cells no longer block the whole plan (a title merged across columns
        ' stopped every tidy-up). Rows touching a merge are kept; deleting rows above
        ' a merge would move it, so that case is still refused after planning.
        Set mergeAreas = DH_MergedAreas(rng)
        Set mergedRows = CreateObject("Scripting.Dictionary")
        For Each mergeArea In mergeAreas
            Set overlap = Application.Intersect(mergeArea, rng)
            If Not overlap Is Nothing Then
                For absoluteRow = overlap.Row To overlap.Row + overlap.Rows.Count - 1
                    mergedRows(CStr(absoluteRow)) = True
                Next absoluteRow
            End If
        Next mergeArea
        guardMerges = True
    ElseIf hasMerges And Not localColumns Then
        If Not DH_CanAutoUnmerge(selected, mergeReason) Then Err.Raise vbObjectError + 761, , mergeReason
        Set mergeAreas = DH_MergedAreas(selected)
        Set mergedRows = CreateObject("Scripting.Dictionary")
        For Each mergeArea In mergeAreas
            Set overlap = Application.Intersect(mergeArea, rng)
            If Not overlap Is Nothing Then
                mergeValue = mergeArea.Cells(1, 1).Value2
                If Not IsEmpty(mergeValue) Then
                    For absoluteRow = mergeArea.Row To mergeArea.Row + mergeArea.Rows.Count - 1
                        If absoluteRow >= firstRow And absoluteRow <= lastRow Then mergedRows(CStr(absoluteRow)) = True
                    Next absoluteRow
                End If
            End If
        Next mergeArea
    End If
    If wholeRowMetadata Then Set noted = DH_NotedRows(rng.EntireRow) Else Set noted = DH_NotedRows(rng)
    values = DH_Matrix(rng)
    formulas = DH_Matrix(rng, True)
    For r = 1 To UBound(values, 1)
        emptyRow = True
        For c = 1 To UBound(values, 2)
            If Not IsEmpty(values(r, c)) Then emptyRow = False: Exit For
            If VarType(formulas(r, c)) = vbString Then
                If Len(formulas(r, c)) > 0 Then emptyRow = False: Exit For
            End If
        Next c
        absoluteRow = firstRow + r - 1
        If noted.Exists(CStr(absoluteRow)) Then emptyRow = False
        If Not mergedRows Is Nothing Then
            If mergedRows.Exists(CStr(absoluteRow)) Then emptyRow = False
        End If
        If emptyRow Then
            If runFirst = 0 Then runFirst = absoluteRow
            runLast = absoluteRow
        ElseIf runFirst > 0 Then
            plan.Add Array(runFirst, runLast)
            runFirst = 0
        End If
    Next r
    If runFirst > 0 Then plan.Add Array(runFirst, runLast)
    If plan.Count > DH_MAX_AREAS Then Err.Raise vbObjectError + 760, , "空行分布过于碎片化，请分批删除。"
    If guardMerges And plan.Count > 0 Then
        ' 只拦「真正穿过合并区」的删除；整段在合并区上/下方的删除是安全的
        ' （Excel 会把合并区整体上移），不再因为「上面有合并」就中断整批整理。
        For Each mergeArea In mergeAreas
            For Each span In plan
                If Not (mergeArea.Row + mergeArea.Rows.Count - 1 < span(0) Or mergeArea.Row > span(1)) Then _
                    Err.Raise vbObjectError + 759, "电子羊表格", "要删除的行穿过了合并单元格，请先取消合并。"
            Next span
        Next mergeArea
    End If
    Set DH_PlanBlankRows = plan
End Function

Public Function DH_BlankRowCount(ByVal plan As Collection) As Long
    Dim span As Variant
    For Each span In plan
        DH_BlankRowCount = DH_BlankRowCount + span(1) - span(0) + 1
    Next span
End Function

Public Function DH_DeleteBlankRows(ByVal selected As Range) As Long
    Dim wb As Workbook, ws As Worksheet, plan As Collection, span As Variant, target As Range, original As Range
    Dim state As DH_AppState, frozen As Boolean, backup As String, address As String, sheetName As String
    Dim i As Long, planned As Long, deleted As Long, message As String, code As Long, mode As String, notes As String
    Dim details As New Collection, audit As String, affected As Range
    On Error GoTo Failed
    If selected Is Nothing Then Err.Raise vbObjectError + 750, , "请先选择要检查空行的区域。"
    Set ws = selected.Parent: Set wb = ws.Parent: Set original = DH_OriginalSelection(ws)
    sheetName = ws.Name: address = selected.Address(False, False)
    Set plan = DH_PlanBlankDeletion(selected, False, notes)
    planned = DH_BlankRowCount(plan): mode = DH_BlankPlanMode(plan)
    If planned = 0 Then
        DH_LastSummary = "没有可删除的完全空白记录；无修改，不创建基线。" & notes
        DH_Log "删除空行", wb, sheetName, address, 0, "成功；" & DH_LastSummary
        Exit Function
    End If
    For Each span In plan
        If span(4) = "整行" Then
            Set target = ws.Rows(CStr(span(0)) & ":" & CStr(span(1)))
        Else
            Set target = ws.Range(ws.Cells(span(0), span(2)), ws.Cells(span(1), span(3)))
        End If
        If affected Is Nothing Then Set affected = target Else Set affected = Application.Union(affected, target)
    Next span
    DH_PreflightStructure wb, affected, False
    DH_Log "删除空行", wb, sheetName, address, 0, "开始；计划清理=" & planned & "；" & mode
    DH_Freeze state: frozen = True
    For Each span In plan
        details.Add Array("R" & span(0) & ":" & span(1) & "C" & span(2) & ":" & span(3), "删除空白记录", _
            "区域内无内容无公式；模式=" & span(4) & "；表=" & span(5), "空白", "上移", "")
    Next span
    audit = DH_SaveAudit("删除空行", wb, details)
    backup = DH_Backup(wb)
    ' All regions were preflighted. Descending worksheet rows keep coordinates
    ' stable, including vertically stacked ListObjects.
    For i = plan.Count To 1 Step -1
        span = plan(i)
        If span(4) = "整行" Then
            ws.Rows(CStr(span(0)) & ":" & CStr(span(1))).Delete
        Else
            Set target = ws.Range(ws.Cells(span(0), span(2)), ws.Cells(span(1), span(3)))
            target.Delete Shift:=xlShiftUp
        End If
        deleted = deleted + span(1) - span(0) + 1
    Next i
    DH_LastSummary = "清理空白记录=" & deleted & "（分区域计数）；" & mode & "；" & notes & "；明细=" & audit
    DH_Log "删除空行", wb, sheetName, address, deleted, "成功；" & DH_LastSummary & "；未自动保存原文件", backup
    DH_DeleteBlankRows = deleted
    DH_RestoreFocus ws, original: DH_Restore state
    Exit Function
Failed:
    message = Err.Description: code = Err.Number
    If Len(backup) > 0 Then
        message = message & "；已确认清理=" & deleted & "条"
        If deleted > 0 Then
            message = message & "；结构删除已部分完成，无法原位回滚，请恢复备份：" & backup
        Else
            message = message & "；尚无已确认完成的删除，请核对当前区域；如需恢复可使用基线：" & backup
        End If
    End If
    DH_RestoreFocus ws, original
    If frozen Then DH_Restore state
    On Error Resume Next
    DH_Log "删除空行", wb, sheetName, address, deleted, "失败；已确认清理=" & deleted & "；" & message, backup
    On Error GoTo 0
    Err.Raise code, "电子羊表格", message
End Function
