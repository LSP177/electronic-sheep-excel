Attribute VB_Name = "DHDuplicates"
Option Explicit



Public Function DH_DuplicatePlan(ByVal selected As Range, ByVal deleteRows As Boolean, ByVal hasHeader As Boolean, _
    ByRef dataRange As Range, ByRef table As ListObject, ByVal details As Collection, ByRef protected As Long) As Collection
    Dim plan As New Collection, ignored As Collection, bounds As Range, source As Variant, formulas As Variant, mask As Variant
    Dim seen As Object, key As String, r As Long, c As Long, flags() As Boolean, emptyRow As Boolean, hasFormula As Boolean
    Dim first As Long, startRow As Long, reason As String, visible As Range, noted As Object, noteRange As Range
    ' Marking cannot delete rows; retain structural guards without the deletion-only width requirement.
    Set ignored = DH_PlanBlankRows(selected, table, False, Not deleteRows)
    If Not table Is Nothing Then
        If table.DataBodyRange Is Nothing Then Set DH_DuplicatePlan = plan: Exit Function
        Set dataRange = Application.Intersect(selected, table.DataBodyRange)
    Else
        Set bounds = DH_ContentRange(selected.Parent)
        If bounds Is Nothing Then Set DH_DuplicatePlan = plan: Exit Function
        Set dataRange = Application.Intersect(bounds, selected)
        If Not dataRange Is Nothing And hasHeader Then
            If dataRange.Rows.Count < 2 Then Set dataRange = Nothing Else Set dataRange = dataRange.Offset(1).Resize(dataRange.Rows.Count - 1)
        End If
    End If
    If dataRange Is Nothing Then Set DH_DuplicatePlan = plan: Exit Function
    If deleteRows Then
        If selected.Parent.FilterMode Then Err.Raise vbObjectError + 880, , "请先清除筛选后删除重复行，避免误删隐藏记录。"
        ' Some Excel builds return False, rather than Null, for mixed Hidden.
        ' Compare the clipped visible-cell count, also conservatively rejecting
        ' hidden columns. Single-cell SpecialCells may expand, so intersect again.
        On Error Resume Next
        Set visible = dataRange.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        If visible Is Nothing Then Err.Raise vbObjectError + 880, , "选区含隐藏行列，请先全部显示。" & HiddenHint(selected)
        Set visible = Application.Intersect(visible, dataRange)
        If visible Is Nothing Then Err.Raise vbObjectError + 880, , "选区含隐藏行列，请先全部显示。" & HiddenHint(selected)
        If visible.CountLarge <> dataRange.CountLarge Then Err.Raise vbObjectError + 880, , "选区含隐藏行列，请先全部显示。" & HiddenHint(selected)
    End If
    Set noteRange = dataRange
    If deleteRows And table Is Nothing Then Set noteRange = dataRange.EntireRow
    Set noted = DH_NotedRows(noteRange)
    source = DH_Matrix(dataRange): formulas = DH_Matrix(dataRange, True): mask = DH_FormulaMask(dataRange)
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbBinaryCompare
    ReDim flags(1 To UBound(source, 1))
    For r = 1 To UBound(source, 1)
        key = "": emptyRow = True: hasFormula = False
        For c = 1 To UBound(source, 2)
            key = key & DH_ValueToken(source(r, c))
            If mask(r, c) Then
                key = key & "FORMULA:" & DH_ValueToken(formulas(r, c)): hasFormula = True
            End If
            If mask(r, c) Or Not IsEmpty(source(r, c)) Then emptyRow = False
        Next c
        If Not emptyRow Then
            If seen.Exists(key) Then
                reason = "首次出现在行" & seen(key)
                If deleteRows And (hasFormula Or noted.Exists(CStr(dataRange.Row + r - 1))) Then
                    protected = protected + 1
                    details.Add Array("行" & (dataRange.Row + r - 1), "保留", "重复项含公式或备注/批注，保留业务信息；" & reason, "[不记录整行值]", "保留", "")
                Else
                    flags(r) = True
                    details.Add Array("行" & (dataRange.Row + r - 1), IIf(deleteRows, "删除重复行", "标记重复行"), reason, "[不记录整行值]", "保留首次出现行", "")
                End If
            Else
                seen.Add key, dataRange.Row + r - 1
            End If
        End If
    Next r
    r = 1
    Do While r <= UBound(flags)
        If flags(r) Then
            startRow = r
            Do While r <= UBound(flags)
                If Not flags(r) Then Exit Do
                r = r + 1
            Loop
            plan.Add Array(dataRange.Row + startRow - 1, dataRange.Row + r - 2)
            If plan.Count > DH_MAX_AREAS Then Err.Raise vbObjectError + 881, , "重复行过于碎片化，请分批处理。"
        Else
            r = r + 1
        End If
    Loop
    Set DH_DuplicatePlan = plan
End Function

Public Function DH_HandleDuplicates(ByVal selected As Range, ByVal deleteRows As Boolean, Optional ByVal hasHeader As Boolean = True) As Long
    Dim ws As Worksheet, wb As Workbook, original As Range, rng As Range, table As ListObject, plan As Collection, span As Variant
    Dim details As New Collection, rules As New Collection, rule As Object, target As Range, oldRules As New Collection, oldCount As Long
    Dim operation As String, count As Long, changed As Long, protected As Long, i As Long, backup As String, audit As String
    Dim state As DH_AppState, frozen As Boolean, code As Long, message As String, address As String, affected As Range
    On Error GoTo Failed
    If selected Is Nothing Then Err.Raise vbObjectError + 882, , "请先选择整行记录区域。"
    Set ws = selected.Parent: Set wb = ws.Parent: Set original = DH_OriginalSelection(ws)
    address = selected.Address(False, False): operation = IIf(deleteRows, "删除重复行", "高亮重复项")
    Set plan = DH_DuplicatePlan(selected, deleteRows, hasHeader, rng, table, details, protected)
    count = DH_BlankRowCount(plan)
    If deleteRows And count > 0 Then
        For Each span In plan
            Set target = ws.Range(ws.Cells(span(0), rng.Column), ws.Cells(span(1), rng.Column + rng.Columns.Count - 1))
            If table Is Nothing Then Set target = target.EntireRow
            If affected Is Nothing Then Set affected = target Else Set affected = Application.Union(affected, target)
        Next span
        DH_PreflightStructure wb, affected, False
    End If
    If Not deleteRows Then
        DH_ValidateMarkNames wb
        oldCount = DH_CountMarks(selected, 3)
    End If
    audit = DH_SaveAudit(operation, wb, details)
    DH_Freeze state: frozen = True
    If count > 0 Or oldCount > 0 Then
        backup = DH_Backup(wb)
        If Not deleteRows Then DH_TakeMarks selected, 3, oldRules, rules
        For i = plan.Count To 1 Step -1
            span = plan(i)
            Set target = ws.Range(ws.Cells(span(0), rng.Column), ws.Cells(span(1), rng.Column + rng.Columns.Count - 1))
            If deleteRows Then
                If table Is Nothing Then target.EntireRow.Delete Else target.Delete Shift:=xlShiftUp
            Else
                Set rule = DH_AddMark(target, 3): rules.Add rule
            End If
            changed = changed + span(1) - span(0) + 1
        Next i
    End If
    DH_LastSummary = "处理重复行=" & changed & "；保留首次出现；含公式/备注重复行保留=" & protected & "；选区内标记按本次结果刷新；明细=" & audit
    DH_Log operation, wb, ws.Name, address, changed, "成功；" & DH_LastSummary, backup
    DH_HandleDuplicates = changed
    DH_RestoreFocus ws, original: DH_Restore state
    Exit Function
Failed:
    message = Err.Description: code = Err.Number
    If deleteRows And Len(backup) > 0 Then message = message & "；结构删除可能部分完成，确认删除=" & changed & "行，请恢复基线：" & backup
    On Error Resume Next
    For i = rules.Count To 1 Step -1: rules(i).Delete: Next i
    Err.Clear: DH_RestoreMarks oldRules
    If Err.Number <> 0 Then message = message & "；旧标记恢复失败，请恢复基线：" & backup
    On Error GoTo 0
    DH_RestoreFocus ws, original
    If frozen Then DH_Restore state
    DH_LogFailure operation, wb, "", "", message, backup
    Err.Raise code, "电子羊表格", message
End Function


' 告诉用户“隐藏列到底在哪一列”
Private Function HiddenHint(ByVal rng As Range) As String
    Dim ws As Worksheet, ci As Long, s As String
    Set ws = rng.Worksheet
    For ci = rng.Column To rng.Column + rng.Columns.Count - 1
        If ws.Columns(ci).Hidden Then
            If Len(s) > 0 Then s = s & ","
            s = s & Letter4Col(ci)
        End If
    Next ci
    If Len(s) = 0 Then HiddenHint = "" Else HiddenHint = "（隐藏列：" & s & "）"
End Function

Private Function Letter4Col(ByVal idx As Long) As String
    Dim n As Long, s2 As String, r As Long
    n = idx
    Do While n > 0
        r = (n - 1) Mod 26
        s2 = Chr$(65 + r) & s2
        n = (n - 1) \ 26
    Loop
    Letter4Col = s2
End Function
