Attribute VB_Name = "DHMerges"
Option Explicit

Public Function DH_MergedAreas(ByVal rng As Range) As Collection
    Dim result As Collection, seen As Object, visits As Long, area As Range
    Set result = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    For Each area In rng.Areas: VisitMerged area, result, seen, visits: Next area
    Set DH_MergedAreas = result
End Function

Private Sub VisitMerged(ByVal rng As Range, ByVal result As Collection, ByVal seen As Object, ByRef visits As Long)
    Dim merged As Variant, top As Range, part As Range, firstHalf As Long, key As String
    visits = visits + 1
    If visits > 200000 Then Err.Raise vbObjectError + 780, , "合并结构过于复杂，请缩小范围。"
    merged = rng.MergeCells
    If Not IsNull(merged) Then
        If Not merged Then Exit Sub
    End If
    Set top = rng.Cells(1, 1).MergeArea
    If top.CountLarge > 1 Then
        Set part = Application.Intersect(top, rng)
        If part.CountLarge = rng.CountLarge Then
            key = top.Address
            If Not seen.Exists(key) Then
                seen.Add key, True: result.Add top
                If result.Count > DH_MAX_AREAS Then Err.Raise vbObjectError + 781, , "合并区域超过10000个，请分批处理。"
            End If
            Exit Sub
        End If
    End If
    If rng.Rows.Count >= rng.Columns.Count And rng.Rows.Count > 1 Then
        firstHalf = rng.Rows.Count \ 2
        VisitMerged rng.Resize(firstHalf, rng.Columns.Count), result, seen, visits
        VisitMerged rng.Cells(firstHalf + 1, 1).Resize(rng.Rows.Count - firstHalf, rng.Columns.Count), result, seen, visits
    ElseIf rng.Columns.Count > 1 Then
        firstHalf = rng.Columns.Count \ 2
        VisitMerged rng.Resize(rng.Rows.Count, firstHalf), result, seen, visits
        VisitMerged rng.Cells(1, firstHalf + 1).Resize(rng.Rows.Count, rng.Columns.Count - firstHalf), result, seen, visits
    End If
End Sub

' Preflight the complete selection before a caller offers to unmerge and continue.
' No eligible area is modified when any selected merge cannot be handled safely.
Public Function DH_CanAutoUnmerge(ByVal selected As Range, ByRef reason As String) As Boolean
    Dim merges As Collection, area As Range
    reason = ""
    If selected Is Nothing Then reason = "请先选择单元格区域。": Exit Function
    Set merges = DH_MergedAreas(selected)
    If merges.Count = 0 Then Exit Function
    For Each area In merges
        reason = DH_MergeSkipReason(area, selected)
        If Len(reason) > 0 Then
            reason = "无法安全一键处理 " & area.Address(False, False) & "：" & reason
            Exit Function
        End If
    Next area
    DH_CanAutoUnmerge = True
End Function

Public Function DH_MergeSkipReason(ByVal area As Range, ByVal selected As Range) As String
    Dim inside As Range, values As Variant, r As Long, topValue As Variant
    Set inside = Application.Intersect(area, selected)
    If inside Is Nothing Then
        DH_MergeSkipReason = "部分选中合并区域"
    ElseIf inside.CountLarge <> area.CountLarge Then
        DH_MergeSkipReason = "部分选中合并区域"
    ElseIf area.Columns.Count <> 1 Then
        DH_MergeSkipReason = "横向或多行多列合并不自动填充"
    ElseIf area.Cells(1, 1).HasFormula Then
        DH_MergeSkipReason = "原合并区包含公式"
    Else
        values = DH_Matrix(area)
        For r = 2 To UBound(values, 1)
            If Not IsEmpty(values(r, 1)) Then
                DH_MergeSkipReason = "原合并区其他位置存在非空数据"
                Exit Function
            End If
        Next r
        topValue = values(1, 1)
        If IsError(topValue) Then DH_MergeSkipReason = "原值是错误常量"
    End If
End Function

Public Function DH_UnmergeFill(ByVal selected As Range) As Long
    Dim originalSelection As Range
    Dim ws As Worksheet, wb As Workbook, merges As Collection, targets As Collection, sourceValues As Collection
    Dim details As Collection, area As Range, inside As Range, report As Worksheet, snapshot As Workbook, cache As Worksheet
    Dim values As Variant, filled As Variant, topValue As Variant, formats As Collection, reason As String
    Dim i As Long, r As Long, touched As Long, changed As Long, filledCells As Long, code As Long, message As String
    Dim state As DH_AppState, frozen As Boolean, backup As String, auditPath As String, summary As String, rollbackFailed As Boolean
    On Error GoTo Failed
    DH_RequireSelection selected
    Set ws = selected.Parent: Set wb = ws.Parent
    Set originalSelection = DH_OriginalSelection(ws)
    Set merges = DH_MergedAreas(selected)
    Set targets = New Collection: Set sourceValues = New Collection: Set formats = New Collection: Set details = New Collection
    For Each area In merges
        reason = DH_MergeSkipReason(area, selected)
        If Len(reason) > 0 Then
            details.Add Array(area.Address(False, False), "跳过", reason, "", "", "")
        Else
            values = DH_Matrix(area): topValue = values(1, 1)
            targets.Add area: sourceValues.Add topValue: formats.Add area.Cells(1, 1).NumberFormat
            details.Add Array(area.Address(False, False), "可解除并区内填充", "完整单列纵向合并，无公式", topValue, topValue, "")
        End If
    Next area
    DH_Log "拆分合并", wb, ws.Name, selected.Address(False, False), 0, "开始；安全候选=" & targets.Count
    DH_Freeze state: frozen = True
    auditPath = DH_SaveAudit("拆分合并", wb, details)
    If targets.Count > 0 Then
        backup = DH_Backup(wb)
        Set snapshot = Application.Workbooks.Add(xlWBATWorksheet): Set cache = snapshot.Worksheets(1)
        For Each area In targets: area.Copy Destination:=cache.Range(area.Address): Next area
        For i = 1 To targets.Count
            Set area = targets(i): topValue = sourceValues(i): touched = i
            area.UnMerge
            If Not IsEmpty(topValue) Then
                ReDim filled(1 To area.Rows.Count, 1 To 1)
                For r = 1 To area.Rows.Count: filled(r, 1) = topValue: Next r
                If VarType(topValue) = vbString Then
                    area.NumberFormat = "@": area.Value2 = DH_LiteralMatrix(filled)
                Else
                    area.NumberFormat = formats(i): area.Value2 = filled
                End If
                filledCells = filledCells + area.Rows.Count - 1
            End If
            changed = changed + 1
        Next i
    End If
    summary = "解除合并区域=" & changed & "；填充单元格=" & filledCells & "；跳过区域=" & merges.Count - changed
    Set report = DH_ActionReport(wb, "拆分合并", ws.Name & "!" & selected.Address(False, False), summary, details, backup, auditPath)
    DH_Log "拆分合并", wb, ws.Name, selected.Address(False, False), changed, "成功；" & summary & "；明细=" & auditPath, backup
    If Not snapshot Is Nothing Then snapshot.Close False
    DH_UnmergeFill = changed
    DH_RestoreFocus ws, originalSelection
    DH_Restore state
    Exit Function
Failed:
    message = Err.Description: code = Err.Number
    On Error Resume Next
    For i = 1 To touched
        Err.Clear: Set area = targets(i)
        area.UnMerge
        cache.Range(area.Address).Copy Destination:=area
        If Err.Number <> 0 Then rollbackFailed = True
    Next i
    If Not snapshot Is Nothing Then snapshot.Close False
    If Not report Is Nothing Then report.Range("A1").Value2 = "操作失败：请检查日志并恢复备份"
    On Error GoTo 0
    If rollbackFailed Then message = message & "；回滚不完整，请恢复基线：" & backup
    DH_RestoreFocus ws, originalSelection
    If frozen Then DH_Restore state
    DH_LogFailure "拆分合并", wb, "", "", message, backup
    Err.Raise code, "电子羊表格", message
End Function
