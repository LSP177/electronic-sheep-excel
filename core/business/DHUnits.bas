Attribute VB_Name = "DHUnits"
Option Explicit

Public Function DH_SplitUnits(ByVal selected As Range) As Long
    Dim ws As Worksheet, wb As Workbook, clipped As Range, body As Range, originalSelection As Range, target As Range, mergeArea As Range
    Dim merges As Collection, targets As Collection, outputs As Collection, details As Collection, report As Worksheet
    Dim values As Variant, masks As Variant, unitData As Variant, flags() As Boolean, numbers() As Double, output As Variant
    Dim row As Long, first As Long, k As Long, i As Long, converted As Long, skipped As Long, column As Long, topRow As Long
    Dim number As Double, unit As String, reason As String, looks As Boolean, accepted As Boolean, header As String, candidate As String
    Dim backup As String, auditPath As String, code As Long, message As String, inserted As Boolean, rollbackFailed As Boolean
    Dim snapshot As Workbook, cache As Worksheet, state As DH_AppState, frozen As Boolean, fso As Object
    Dim table As ListObject
    On Error GoTo Failed
    Set clipped = DH_ClipColumn(selected)
    If clipped Is Nothing Then Err.Raise vbObjectError + 812, , "选区没有内容。"
    DH_RequireSelection clipped
    Set ws = clipped.Parent: Set wb = ws.Parent: Set originalSelection = DH_OriginalSelection(ws)
    column = clipped.Column: topRow = clipped.Row
    If clipped.Rows.Count < 2 Then Err.Raise vbObjectError + 813, , "请包含列标题和至少一行数据。"
    If column = ws.Columns.Count Then Err.Raise vbObjectError + 814, , "最后一列右侧无法插入单位列。"
    If clipped.Cells(1, 1).HasFormula Then Err.Raise vbObjectError + 815, , "列标题不能是公式。"
    If VarType(clipped.Cells(1, 1).Value2) <> vbString Then Err.Raise vbObjectError + 816, , "请从文字列标题开始选择。"
    header = CStr(clipped.Cells(1, 1).Value2)
    accepted = DH_TryNumber(header, number, unit, reason, looks)
    If looks Or Len(header) = 0 Or Len(header) > 200 Then Err.Raise vbObjectError + 817, , "首行应为列标题，请把原列标题包含在选区内。"
    For Each table In ws.ListObjects
        If column >= table.Range.Column And column <= table.Range.Column + table.Range.Columns.Count Then _
            Err.Raise vbObjectError + 818, , "单位拆分暂不处理Excel表格或其相邻边界，请在普通区域副本中操作。"
    Next table
    Set merges = DH_MergedAreas(ws.UsedRange)
    For Each mergeArea In merges
        If Not Application.Intersect(mergeArea, clipped) Is Nothing Then Err.Raise vbObjectError + 819, , "所选列含合并，请先安全拆分。"
        If mergeArea.Column <= column And mergeArea.Column + mergeArea.Columns.Count - 1 > column Then _
            Err.Raise vbObjectError + 820, , "插列位置跨越合并区域，请先调整合并。"
    Next mergeArea
    Set target = ws.Columns(ws.Columns.Count).Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    If Not target Is Nothing Then Err.Raise vbObjectError + 821, , "工作表最后一列存在数据，插列可能挤出数据，已停止。"
    Set body = clipped.Cells(2, 1).Resize(clipped.Rows.Count - 1, 1)
    values = DH_Matrix(body): masks = DH_FormulaMask(body)
    ReDim flags(1 To UBound(values, 1)): ReDim numbers(1 To UBound(values, 1))
    ReDim unitData(1 To UBound(values, 1), 1 To 1)
    Set targets = New Collection: Set outputs = New Collection: Set details = New Collection
    For row = 1 To UBound(values, 1)
        reason = "非带单位文本，保留": unit = "": accepted = False
        If masks(row, 1) Then
            reason = "公式单元格，保留"
        ElseIf VarType(values(row, 1)) = vbString Then
            accepted = DH_TryNumber(CStr(values(row, 1)), number, unit, reason, looks)
            If Len(unit) = 0 Then accepted = False: reason = "无明确支持单位，保留"
        End If
        If accepted Then
            flags(row) = True: numbers(row) = number: unitData(row, 1) = unit: converted = converted + 1
            details.Add Array("R" & (topRow + row) & "C" & column, "可拆分", "明确数值及单位，单位独立保留", values(row, 1), number, unit)
        ElseIf Not IsEmpty(values(row, 1)) Or masks(row, 1) Then
            skipped = skipped + 1
            details.Add Array("R" & (topRow + row) & "C" & column, "保留", reason, values(row, 1), values(row, 1), unit)
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
            ReDim output(1 To row - first, 1 To 1)
            For k = first To row - 1: output(k - first + 1, 1) = numbers(k): Next k
            targets.Add body.Cells(first, 1).Resize(row - first, 1): outputs.Add output
            If targets.Count > DH_MAX_AREAS Then Err.Raise vbObjectError + 822, , "候选过于碎片化，请分批处理。"
        Else
            row = row + 1
        End If
    Loop
    DH_Freeze state: frozen = True
    auditPath = DH_SaveAudit("单位拆分", wb, details)
    If converted > 0 Then
        DH_PreflightStructure wb, ws.Columns(column + 1), True
        backup = DH_Backup(wb)
        Set snapshot = Application.Workbooks.Add(xlWBATWorksheet): Set cache = snapshot.Worksheets(1)
        For Each target In targets: target.Copy Destination:=cache.Range(target.Address): Next target
        ws.Columns(column + 1).Insert Shift:=xlToRight: inserted = True
        candidate = header & "_单位": i = 1
        Do While HeaderExists(ws, topRow, candidate, column + 1)
            i = i + 1: candidate = header & "_单位_" & i
        Loop
        ws.Cells(topRow, column + 1).NumberFormat = "@": ws.Cells(topRow, column + 1).Value2 = "'" & candidate
        With ws.Cells(topRow + 1, column + 1).Resize(UBound(unitData, 1), 1)
            .NumberFormat = "@": .Value2 = DH_LiteralMatrix(unitData)
        End With
        For i = 1 To targets.Count
            Set target = targets(i): target.NumberFormat = "General": target.Value2 = outputs(i)
        Next i
    End If
    Set report = DH_ActionReport(wb, "单位拆分", ws.Name & "!" & clipped.Address(False, False), _
        "转换=" & converted & "；保留=" & skipped & "；混合单位逐格保存在新单位列", details, backup, auditPath)
    DH_Log "单位拆分", wb, ws.Name, clipped.Address(False, False), converted, "成功；明细=" & auditPath, backup
    If Not snapshot Is Nothing Then snapshot.Close False
    DH_SplitUnits = converted
    DH_RestoreFocus ws, originalSelection: DH_Restore state
    Exit Function
Failed:
    message = Err.Description: code = Err.Number
    On Error Resume Next
    If inserted Then
        Err.Clear: ws.Columns(column + 1).Delete
        If Err.Number <> 0 Then rollbackFailed = True
        For Each target In targets
            Err.Clear: cache.Range(target.Address).Copy Destination:=target
            If Err.Number <> 0 Then rollbackFailed = True
        Next target
    End If
    If Not snapshot Is Nothing Then snapshot.Close False
    On Error GoTo 0
    If rollbackFailed Then message = message & "；回滚不完整，请恢复当日基线：" & backup
    DH_RestoreFocus ws, originalSelection
    If frozen Then DH_Restore state
    DH_LogFailure "单位拆分", wb, "", "", message, backup
    Err.Raise code, "电子羊表格", message
End Function

Private Function HeaderExists(ByVal ws As Worksheet, ByVal row As Long, ByVal text As String, ByVal excludedColumn As Long) As Boolean
    Dim values As Variant, c As Long
    values = DH_Matrix(Application.Intersect(ws.Rows(row), ws.UsedRange))
    For c = 1 To UBound(values, 2)
        If VarType(values(1, c)) = vbString Then
            If StrComp(CStr(values(1, c)), text, vbTextCompare) = 0 Then HeaderExists = True: Exit Function
        End If
    Next c
End Function
