Attribute VB_Name = "DHReconcile"
Option Explicit

Private Function MatchKey(ByVal value As Variant, ByRef valid As Boolean) As String
    valid = False
    If IsEmpty(value) Or IsError(value) Then Exit Function
    If VarType(value) = vbString Then
        If Len(value) = 0 Then Exit Function
    End If
    MatchKey = DH_ValueToken(value)
    valid = True
End Function

Private Function DataColumn(ByVal selected As Range, ByVal hasHeader As Boolean) As Range
    Dim clipped As Range, merges As Collection
    Set clipped = DH_ClipColumn(selected)
    If clipped Is Nothing Then Err.Raise vbObjectError + 830, , "比对列没有数据。"
    DH_RequireSelection clipped
    Set merges = DH_MergedAreas(clipped)
    If merges.Count > 0 Then Err.Raise vbObjectError + 831, , "比对列包含合并，请先安全拆分。"
    If hasHeader Then
        If clipped.Rows.Count < 2 Then Err.Raise vbObjectError + 832, , "比对列只有表头。"
        Set clipped = clipped.Cells(2, 1).Resize(clipped.Rows.Count - 1, 1)
    End If
    Set DataColumn = clipped
End Function

Private Sub CollectMissing(ByVal rng As Range, ByVal values As Variant, ByVal other As Object, ByVal side As Long, _
    ByVal targets As Collection, ByVal sides As Collection, ByVal details As Collection, ByRef missing As Long, ByRef skipped As Long)
    Dim row As Long, first As Long, key As String, valid As Boolean, flags() As Boolean, firstRow As Long, column As Long
    Dim sheetName As String
    ReDim flags(1 To UBound(values, 1))
    firstRow = rng.Row: column = rng.Column: sheetName = rng.Parent.Name
    For row = 1 To UBound(values, 1)
        key = MatchKey(values(row, 1), valid)
        If valid Then
            If Not other.Exists(key) Then
                flags(row) = True: missing = missing + 1
                details.Add Array(sheetName & "!R" & (firstRow + row - 1) & "C" & column, _
                    IIf(side = 1, "主表缺少对应项", "核对表陌生项"), "另一列不存在相同值", values(row, 1), values(row, 1), "")
            End If
        ElseIf IsError(values(row, 1)) Then
            skipped = skipped + 1
            details.Add Array(sheetName & "!R" & (firstRow + row - 1) & "C" & column, "跳过", "错误值不参与匹配", "错误值", "错误值", "")
        End If
    Next row
    row = 1
    Do While row <= UBound(flags)
        If flags(row) Then
            first = row
            Do While row <= UBound(flags)
                If Not flags(row) Then Exit Do
                row = row + 1
            Loop
            targets.Add rng.Cells(first, 1).Resize(row - first, 1): sides.Add side
            If targets.Count > DH_MAX_AREAS Then Err.Raise vbObjectError + 833, , "差异过于碎片化，请分批对账。"
        Else
            row = row + 1
        End If
    Loop
End Sub

Public Function DH_ReconcileColumns(ByVal mainSelection As Range, ByVal checkSelection As Range, _
    Optional ByVal hasHeader As Boolean = False) As Long
    Dim main As Range, other As Range, originalSheet As Worksheet, originalSelection As Range, wb As Workbook, report As Worksheet
    Dim leftValues As Variant, rightValues As Variant, leftSet As Object, rightSet As Object, key As String, valid As Boolean, row As Long
    Dim targets As Collection, sides As Collection, details As Collection, target As Range, i As Long, touched As Long
    Dim mainMissing As Long, otherMissing As Long, skipped As Long, backup As String, auditPath As String, message As String, code As Long
    Dim snapshot As Workbook, cacheLeft As Worksheet, cacheRight As Worksheet, cache As Worksheet, rollbackFailed As Boolean
    Dim state As DH_AppState, frozen As Boolean
    Dim addedRules As New Collection, addedRule As Object, oldRules As New Collection, oldCount As Long
    On Error GoTo Failed
    Set main = DataColumn(mainSelection, hasHeader): Set other = DataColumn(checkSelection, hasHeader)
    Set wb = main.Parent.Parent: Set originalSheet = main.Parent: Set originalSelection = DH_OriginalSelection(originalSheet)
    If Not other.Parent.Parent Is wb Then Err.Raise vbObjectError + 834, , "请在同一工作簿中选择两张工作表。"
    If main.Parent Is other.Parent Then Err.Raise vbObjectError + 835, , "请在不同工作表选择主表和核对表。"
    If main.CountLarge + other.CountLarge > DH_MAX_CELLS Then Err.Raise vbObjectError + 836, , "两列合计超过100万单元格，请分批。"
    leftValues = DH_Matrix(main): rightValues = DH_Matrix(other)
    Set leftSet = CreateObject("Scripting.Dictionary"): leftSet.CompareMode = vbBinaryCompare
    Set rightSet = CreateObject("Scripting.Dictionary"): rightSet.CompareMode = vbBinaryCompare
    For row = 1 To UBound(leftValues, 1)
        key = MatchKey(leftValues(row, 1), valid): If valid Then leftSet(key) = True
    Next row
    For row = 1 To UBound(rightValues, 1)
        key = MatchKey(rightValues(row, 1), valid): If valid Then rightSet(key) = True
    Next row
    Set targets = New Collection: Set sides = New Collection: Set details = New Collection
    CollectMissing main, leftValues, rightSet, 1, targets, sides, details, mainMissing, skipped
    CollectMissing other, rightValues, leftSet, 2, targets, sides, details, otherMissing, skipped
    DH_ValidateMarkNames wb
    oldCount = DH_CountMarks(mainSelection, 1) + DH_CountMarks(checkSelection, 2)
    DH_Freeze state: frozen = True
    auditPath = DH_SaveAudit("两表对账", wb, details)
    If targets.Count > 0 Or oldCount > 0 Then
        backup = DH_Backup(wb)
        DH_TakeMarks mainSelection, 1, oldRules, addedRules
        DH_TakeMarks checkSelection, 2, oldRules, addedRules
        For i = 1 To targets.Count
            Set target = targets(i): touched = i
            Set addedRule = DH_AddMark(target, CLng(sides(i))): addedRules.Add addedRule
        Next i
    End If
    message = "主表独有=" & mainMissing & "；核对表独有=" & otherMissing & "；错误值跳过=" & skipped & "；已刷新选区内旧标记；按底层值/类型及大小写比较存在性，不核对重复次数"
    Set report = DH_ActionReport(wb, "两表对账", main.Parent.Name & "!" & main.Address(False, False) & " vs " & _
        other.Parent.Name & "!" & other.Address(False, False), message, details, backup, auditPath)
    DH_Log "两表对账", wb, main.Parent.Name, main.Address(False, False), mainMissing + otherMissing, "成功；" & message, backup
    If Not snapshot Is Nothing Then snapshot.Close False
    DH_ReconcileColumns = mainMissing + otherMissing
    DH_RestoreFocus originalSheet, originalSelection: DH_Restore state
    Exit Function
Failed:
    message = Err.Description: code = Err.Number
    On Error Resume Next
    For i = addedRules.Count To 1 Step -1
        Err.Clear: addedRules(i).Delete
        If Err.Number <> 0 Then rollbackFailed = True
    Next i
    Err.Clear: DH_RestoreMarks oldRules
    If Err.Number <> 0 Then rollbackFailed = True
    Application.CutCopyMode = False
    If Not snapshot Is Nothing Then snapshot.Close False
    If Not report Is Nothing Then report.Range("A1").Value2 = "对账失败，请查看日志并恢复基线"
    On Error GoTo 0
    If rollbackFailed Then message = message & "；颜色回滚不完整，请恢复基线：" & backup
    DH_RestoreFocus originalSheet, originalSelection
    If frozen Then DH_Restore state
    DH_LogFailure "两表对账", wb, "", "", message, backup
    Err.Raise code, "电子羊表格", message
End Function
