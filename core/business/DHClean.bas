Attribute VB_Name = "DHClean"
Option Explicit

' Plan only changed vertical runs from the already-read arrays. Unchanged
' constants, including ordinary headers, never enter format/value writeback.
Private Sub AddChangedRuns(ByVal block As Range, ByRef source As Variant, ByRef data As Variant, _
    ByVal targets As Collection, ByVal outputs As Collection)
    Dim r As Long, c As Long, first As Long, k As Long, output As Variant
    For c = 1 To UBound(data, 2)
        r = 1
        Do While r <= UBound(data, 1)
            If StrComp(CStr(data(r, c)), CStr(source(r, c)), vbBinaryCompare) <> 0 Then
                first = r
                Do While r <= UBound(data, 1)
                    If StrComp(CStr(data(r, c)), CStr(source(r, c)), vbBinaryCompare) = 0 Then Exit Do
                    r = r + 1
                Loop
                ReDim output(1 To r - first, 1 To 1)
                For k = first To r - 1: output(k - first + 1, 1) = data(k, c): Next k
                If targets.Count >= DH_MAX_AREAS Then Err.Raise vbObjectError + 737, , "变化区域过于碎片化，请分批清理；尚未修改数据。"
                targets.Add block.Cells(first, c).Resize(r - first, 1)
                outputs.Add output
            Else
                r = r + 1
            End If
        Loop
    Next c
End Sub

Public Function DH_CleanRange(ByVal selectionRange As Range) As Long
    Dim wb As Workbook, ws As Worksheet, bounds As Range, area As Range, clipped As Range, texts As Range, block As Range
    Dim targets As Collection, cleaned As Collection
    Dim source As Variant, data As Variant, merged As Variant, state As DH_AppState, frozen As Boolean
    Dim r As Long, c As Long, i As Long, touched As Long, changed As Long, blockChanges As Long
    Dim backup As String, address As String, sheetName As String, code As Long, message As String, rollbackFailed As Boolean
    Dim totalCells As Double, eligible As Range, candidate As Range, mergeCut As Range, mergeArea As Range
    Dim snapshot As Workbook, snapshotSheet As Worksheet
    Dim details As New Collection, audit As String, blockRow As Long, blockColumn As Long, originalSelection As Range
    On Error GoTo Failed
    If selectionRange Is Nothing Then Err.Raise vbObjectError + 730, , "请先选择要清理的单元格区域。"
    Set ws = selectionRange.Parent: Set wb = ws.Parent
    Set originalSelection = DH_OriginalSelection(ws)
    sheetName = ws.Name: address = selectionRange.Address(False, False)
    DH_RequireWorkbook wb
    If ws.ProtectContents Then Err.Raise vbObjectError + 731, , "工作表受保护，请先取消保护。"
    If Application.ActiveWindow Is Nothing Then Err.Raise vbObjectError + 732, , "没有活动窗口。"
    If Application.ActiveWindow.SelectedSheets.Count <> 1 Then Err.Raise vbObjectError + 733, , "请先取消工作表分组。"
    If selectionRange.Areas.Count > DH_MAX_AREAS Then Err.Raise vbObjectError + 734, , "选区过于碎片化，请分批清理。"
    Set targets = New Collection: Set cleaned = New Collection
    Set bounds = DH_ContentRange(ws)
    If Not bounds Is Nothing Then
        For Each area In selectionRange.Areas
            Set clipped = Application.Intersect(area, bounds)
            If Not clipped Is Nothing Then
                ' Merged cells are skipped and audited instead of rejecting the whole
                ' range (a title merged across columns blocked every tidy-up).
                Set mergeCut = Nothing
                merged = clipped.MergeCells
                If IsNull(merged) Then merged = True
                If merged Then
                    For Each mergeArea In DH_MergedAreas(clipped)
                        If mergeCut Is Nothing Then Set mergeCut = mergeArea Else Set mergeCut = Application.Union(mergeCut, mergeArea)
                        details.Add Array(mergeArea.Address(False, False), "跳过", "合并单元格不清理；拆开合并后可再处理", "", "", "")
                    Next mergeArea
                End If
                If clipped.CountLarge > DH_MAX_CELLS Then Err.Raise vbObjectError + 736, , "单个选区超过100万单元格，请缩小选区。"
                totalCells = totalCells + clipped.CountLarge
                If totalCells > DH_MAX_CELLS Then Err.Raise vbObjectError + 736, , "选区合计超过100万单元格，请分批清理。"
                Set texts = Nothing
                If clipped.CountLarge = 1 Then
                    If Not clipped.HasFormula Then
                        If VarType(clipped.Value2) = vbString Then Set texts = clipped
                    End If
                Else
                    On Error Resume Next
                    Set texts = clipped.SpecialCells(xlCellTypeConstants, xlTextValues)
                    On Error GoTo Failed
                    If Not texts Is Nothing Then Set texts = Application.Intersect(texts, clipped)
                End If
                If Not texts Is Nothing And Not mergeCut Is Nothing Then
                    Set mergeCut = Application.Intersect(mergeCut, clipped)
                    If Not mergeCut Is Nothing Then Set texts = DH_SubtractRange(texts, mergeCut)
                End If
                If Not texts Is Nothing Then
                    If texts.Areas.Count > DH_MAX_AREAS Then Err.Raise vbObjectError + 737, , "文本常量过于碎片化，请分批清理。"
                    Set eligible = DH_WithoutTableHeaders(texts)
                    If eligible Is Nothing Then GoTo NextArea
                    For Each block In eligible.Areas
                        blockRow = block.Row: blockColumn = block.Column
                        source = DH_Matrix(block)
                        data = source: blockChanges = 0
                        For r = 1 To UBound(data, 1)
                            For c = 1 To UBound(data, 2)
                                data(r, c) = DH_CleanText(CStr(source(r, c)))
                                If data(r, c) <> source(r, c) Then
                                    blockChanges = blockChanges + 1
                                    details.Add Array("R" & (blockRow + r - 1) & "C" & (blockColumn + c - 1), "区域清理", "语言安全空格规则", source(r, c), data(r, c), "")
                                End If
                            Next c
                        Next r
                        If blockChanges > 0 Then
                            AddChangedRuns block, source, data, targets, cleaned
                            changed = changed + blockChanges
                        End If
                    Next block
                End If
            End If
NextArea:
        Next area
    End If
    DH_Log "区域清理", wb, sheetName, address, 0, "开始；计划修改=" & changed
    If targets.Count = 0 Then
        DH_Log "区域清理", wb, sheetName, address, 0, "无修改，不创建基线"
        Exit Function
    End If
    ' A baseline is required before writing the nonempty change plan.
    DH_Freeze state: frozen = True
    audit = DH_SaveAudit("区域清理", wb, details)
    backup = DH_Backup(wb)
    If targets.Count > 0 Then
        ' A separate unsaved workbook holds exact value/type/format snapshots.
        ' Copy Destination avoids per-cell format reads, including mixed formats.
        Set snapshot = Application.Workbooks.Add(xlWBATWorksheet)
        Set snapshotSheet = snapshot.Worksheets(1)
        For i = 1 To targets.Count
            Set block = targets(i)
            block.Copy Destination:=snapshotSheet.Range(block.Address)
        Next i
    End If
    For i = 1 To targets.Count
        touched = i
        Set block = targets(i)
        block.NumberFormat = "@"
        block.Value2 = DH_LiteralMatrix(cleaned(i))
    Next i
    DH_Log "区域清理", wb, sheetName, address, changed, "成功；公式及超级表字段名已跳过；未自动保存原文件；明细=" & audit, backup
    If Not snapshot Is Nothing Then snapshot.Close SaveChanges:=False
    Set snapshot = Nothing
    DH_RestoreFocus ws, originalSelection
    DH_CleanRange = changed
    DH_Restore state
    Exit Function
Failed:
    message = Err.Description: code = Err.Number
    ' A write can partially succeed. Include the currently attempted block in rollback.
    On Error Resume Next
    For i = 1 To touched
        Err.Clear
        Set block = targets(i)
        snapshotSheet.Range(block.Address).Copy Destination:=block
        If Err.Number <> 0 Then rollbackFailed = True
    Next i
    If Not snapshot Is Nothing Then snapshot.Close SaveChanges:=False
    DH_RestoreFocus ws, originalSelection
    On Error GoTo 0
    If rollbackFailed Then message = message & "；回滚不完整，请恢复备份：" & backup
    If frozen Then DH_Restore state
    DH_LogFailure "区域清理", wb, sheetName, address, message, backup
    Err.Raise code, "电子羊表格", message
End Function
