Attribute VB_Name = "DHTransform"
Option Explicit

' Reason text for month/day/year slashes; DH_TransformSelection counts it for the result note.
Public Function DH_AmbiguousMdy() As String
    DH_AmbiguousMdy = "月/日/年写法看不出是几月几日，原样保留"
End Function

Public Function DH_TryDate(ByVal text As String, ByRef serial As Double, ByRef reason As String, _
    Optional ByVal date1904 As Boolean = False) As Boolean
    Static pattern As Object, cnPattern As Object, mdyPattern As Object
    Dim body As String, y As Long, m As Long, d As Long, stamp As Date, match As Object
    reason = "不是明确的年-月-日文本": body = Trim$(text)
    If pattern Is Nothing Then
        Set pattern = CreateObject("VBScript.RegExp")
        pattern.Pattern = "^([0-9]{4})([./-])([0-9]{1,2})\2([0-9]{1,2})$"
        ' Year-first Chinese form, e.g. 2026年10月11日 (the trailing 日 may be missing).
        Set cnPattern = CreateObject("VBScript.RegExp")
        cnPattern.Pattern = "^([0-9]{4})年([0-9]{1,2})月([0-9]{1,2})日?$"
        ' Month/day/year with slashes is ambiguous (10/12/2026); it is named, never guessed.
        Set mdyPattern = CreateObject("VBScript.RegExp")
        mdyPattern.Pattern = "^[0-9]{1,2}/[0-9]{1,2}/[0-9]{4}$"
    End If
    If Len(body) = 8 And body Like "########" Then
        y = CLng(Left$(body, 4)): m = CLng(Mid$(body, 5, 2)): d = CLng(Right$(body, 2))
    ElseIf pattern.Test(body) Then
        Set match = pattern.Execute(body)(0)
        y = CLng(match.SubMatches(0)): m = CLng(match.SubMatches(2)): d = CLng(match.SubMatches(3))
    ElseIf cnPattern.Test(body) Then
        Set match = cnPattern.Execute(body)(0)
        y = CLng(match.SubMatches(0)): m = CLng(match.SubMatches(1)): d = CLng(match.SubMatches(2))
    Else
        If mdyPattern.Test(body) Then reason = DH_AmbiguousMdy()
        Exit Function
    End If
    reason = "非法日期或超出1900至2100年范围"
    If y < 1900 Or y > 2100 Or m < 1 Or m > 12 Or d < 1 Or d > 31 Then Exit Function
    stamp = DateSerial(y, m, d)
    If Year(stamp) <> y Or Month(stamp) <> m Or Day(stamp) <> d Then Exit Function
    If date1904 And y < 1904 Then reason = "1904日期系统不处理1904年前日期": Exit Function
    serial = CDbl(stamp)
    If stamp < DateSerial(1900, 3, 1) Then serial = serial - 1
    If date1904 Then serial = CDbl(stamp) - 1462
    reason = "明确日期文本": DH_TryDate = True
End Function

Public Function DH_MaskText(ByVal value As Variant, ByVal mode As Long, ByRef masked As String, ByRef reason As String) As Boolean
    Dim text As String, i As Long, code As Long, surnameLength As Long, prefix As String
    reason = "不符合所选类型，保留原值"
    If IsError(value) Or IsEmpty(value) Then Exit Function
    text = Trim$(CStr(value))
    Select Case mode
        Case 1
            If Len(text) <> 11 Then Exit Function
            If Not text Like "1##########" Then Exit Function
            If Mid$(text, 2, 1) < "3" Or Mid$(text, 2, 1) > "9" Then Exit Function
            masked = Left$(text, 3) & "****" & Right$(text, 4)
        Case 2
            ' Numeric 18-digit IDs may already have lost precision; never infer.
            If VarType(value) <> vbString Or Len(text) <> 18 Then Exit Function
            prefix = Left$(text, 17)
            If Not prefix Like "#################" Then Exit Function
            If Not (Right$(text, 1) Like "#" Or UCase$(Right$(text, 1)) = "X") Then Exit Function
            masked = Left$(text, 6) & "********" & Right$(text, 4)
        Case 3
            If VarType(value) <> vbString Or Len(text) < 2 Or Len(text) > 4 Then Exit Function
            If InStr("|姓名|中文姓名|经办人|联系人|负责人|客户名|员工姓名|客户姓名|", "|" & text & "|") > 0 Then reason = "常见姓名表头保留": Exit Function
            For i = 1 To Len(text)
                code = AscW(Mid$(text, i, 1)): If code < 0 Then code = code + 65536
                If code < &H4E00 Or code > &H9FFF& Then Exit Function
            Next i
            surnameLength = 1
            If InStr("|欧阳|司马|上官|诸葛|夏侯|东方|皇甫|尉迟|公孙|慕容|司徒|司空|令狐|长孙|", "|" & Left$(text, 2) & "|") > 0 Then
                If Len(text) < 3 Then Exit Function
                surnameLength = 2
            End If
            masked = Left$(text, surnameLength) & String$(Len(text) - surnameLength, "*")
        Case Else
            Err.Raise vbObjectError + 860, , "请选择手机号、身份证或姓名模式。"
    End Select
    reason = "按显式所选类型脱敏；不验证身份真实性"
    DH_MaskText = True
End Function

Private Function ConstantBlocks(ByVal selected As Range) As Collection
    Dim result As New Collection, bounds As Range, area As Range, clipped As Range, constants As Range, part As Range
    Set bounds = DH_ContentRange(selected.Parent)
    If Not bounds Is Nothing Then
        For Each area In selected.Areas
            Set clipped = Application.Intersect(area, bounds): Set constants = Nothing
            If Not clipped Is Nothing Then
                If clipped.CountLarge = 1 Then
                    If Not clipped.HasFormula Then Set constants = clipped
                Else
                    On Error Resume Next
                    Set constants = clipped.SpecialCells(xlCellTypeConstants, xlTextValues + xlNumbers)
                    On Error GoTo 0
                    If Not constants Is Nothing Then Set constants = Application.Intersect(constants, clipped)
                End If
                If Not constants Is Nothing Then
                    For Each part In constants.Areas
                        result.Add part
                        If result.Count > DH_MAX_AREAS Then Err.Raise vbObjectError + 861, , "常量过于碎片化，请分批。"
                    Next part
                End If
            End If
        Next area
    End If
    Set ConstantBlocks = result
End Function

Private Function MaskHeaders(ByVal block As Range) As Variant
    Dim protected() As Boolean, table As ListObject, header As Range, overlap As Range, part As Range
    Dim r As Long, c As Long, firstRow As Long, firstColumn As Long
    ReDim protected(1 To block.Rows.Count, 1 To block.Columns.Count)
    firstRow = block.Row: firstColumn = block.Column
    For Each table In block.Parent.ListObjects
        Set header = Nothing
        If table.ShowHeaders Then Set header = table.HeaderRowRange
        If table.ShowTotals Then
            If header Is Nothing Then Set header = table.TotalsRowRange Else Set header = Application.Union(header, table.TotalsRowRange)
        End If
        If Not header Is Nothing Then
            Set overlap = Application.Intersect(block, header)
            If Not overlap Is Nothing Then
                For Each part In overlap.Areas
                    For r = part.Row - firstRow + 1 To part.Row - firstRow + part.Rows.Count
                        For c = part.Column - firstColumn + 1 To part.Column - firstColumn + part.Columns.Count
                            protected(r, c) = True
                        Next c
                    Next r
                Next part
            End If
        End If
    Next table
    MaskHeaders = protected
End Function

' mode=0 dates; 1 phone; 2 ID; 3 explicitly selected Chinese names.
Public Function DH_TransformSelection(ByVal selected As Range, ByVal mode As Long) As Long
    Dim ws As Worksheet, wb As Workbook, original As Range, blocks As Collection, block As Range, merges As Collection
    Dim targets As New Collection, outputs As New Collection, details As New Collection
    Dim data As Variant, values As Variant, flags() As Boolean, part As Variant, serial As Double, masked As String, reason As String
    Dim r As Long, c As Long, first As Long, i As Long, k As Long, touched As Long, changed As Long, skipped As Long, ambiguous As Long
    Dim operation As String, backup As String, audit As String, message As String, code As Long, location As String
    Dim snapshot As Workbook, cache As Worksheet, state As DH_AppState, frozen As Boolean, rollbackFailed As Boolean, accepted As Boolean
    Dim date1904 As Boolean, blockRow As Long, blockColumn As Long
    Dim working As Range, bounds As Range, headerMask As Variant, mergeCut As Range, mergeArea As Range
    On Error GoTo Failed
    If mode < 0 Or mode > 3 Then Err.Raise vbObjectError + 860, , "不支持的处理类型。"
    operation = "日期规范化": If mode > 0 Then operation = "敏感数据脱敏"
    If selected Is Nothing Then Err.Raise vbObjectError + 862, , "请先选择数据区域。"
    Set ws = selected.Parent: Set wb = ws.Parent: Set original = DH_OriginalSelection(ws)
    Set working = selected
    If mode > 0 Then
        ' Validate the context without treating an entire column/sheet's unused
        ' cells as data. Physical row 1 is never a masking candidate.
        DH_RequireSelection selected.Cells(1, 1), False
        If selected.Areas.Count > DH_MAX_AREAS Then Err.Raise vbObjectError + 861, , "选区过于碎片化，请分批。"
        Set bounds = DH_ContentRange(ws)
        Set working = Nothing
        If Not bounds Is Nothing Then Set working = Application.Intersect(selected, bounds, ws.Rows("2:" & ws.Rows.Count))
        If working Is Nothing Then
            DH_LastAuditPath = ""
            DH_LastSummary = "无数据行候选；工作表第1行始终保护，未修改，不创建基线。"
            DH_Log operation, wb, ws.Name, selected.Address(False, False), 0, "成功；" & DH_LastSummary
            Exit Function
        End If
    End If
    DH_RequireSelection working, False
    date1904 = wb.Date1904
    Set merges = DH_MergedAreas(working)
    If merges.Count > 0 Then
        ' Masking must never leave personal data behind silently: keep rejecting.
        If mode > 0 Then Err.Raise vbObjectError + 862, , "选区含合并区域，请先安全拆分。"
        ' Date normalisation skips merged cells and audits them (a title merged
        ' across columns used to block the whole tidy-up).
        Set mergeCut = Nothing
        For Each mergeArea In merges
            If mergeCut Is Nothing Then Set mergeCut = mergeArea Else Set mergeCut = Application.Union(mergeCut, mergeArea)
            details.Add Array(mergeArea.Address(False, False), "跳过", "合并单元格不处理；拆开合并后可再处理", "", "", "")
            skipped = skipped + 1
        Next mergeArea
        Set mergeCut = Application.Intersect(mergeCut, working)
        If Not mergeCut Is Nothing Then Set working = DH_SubtractRange(working, mergeCut)
        If working Is Nothing Then DH_LastSummary = "仅包含合并单元格，未修改": Exit Function
    End If
    Set working = DH_WithoutTableHeaders(working)
    If working Is Nothing Then DH_LastSummary = "仅包含受保护的超级表表头，无修改": Exit Function
    Set blocks = ConstantBlocks(working)
    For Each block In blocks
        blockRow = block.Row: blockColumn = block.Column
        data = DH_Matrix(block): values = data
        If mode > 0 Then headerMask = MaskHeaders(block)
        ReDim flags(1 To UBound(data, 1), 1 To UBound(data, 2))
        For c = 1 To UBound(data, 2)
            For r = 1 To UBound(data, 1)
                accepted = False: reason = "原有数值、空值或非日期文本保留"
                If mode = 0 Then
                    If VarType(data(r, c)) = vbString Then accepted = DH_TryDate(CStr(data(r, c)), serial, reason, date1904)
                    If Not accepted And reason = DH_AmbiguousMdy() Then ambiguous = ambiguous + 1
                    If accepted Then values(r, c) = serial
                Else
                    If headerMask(r, c) Then
                        reason = "超级表表头或汇总行，保护原值"
                    Else
                        accepted = DH_MaskText(data(r, c), mode, masked, reason)
                    End If
                    If accepted Then values(r, c) = masked
                End If
                flags(r, c) = accepted
                location = "R" & (blockRow + r - 1) & "C" & (blockColumn + c - 1)
                If accepted Then
                    changed = changed + 1
                    ' Never persist unmasked personal values, including rejected items.
                    If mode = 0 Then
                        details.Add Array(location, "日期转换", reason, data(r, c), values(r, c), "")
                    Else
                        details.Add Array(location, "脱敏", reason, "[原文不记录]", masked, "")
                    End If
                Else
                    skipped = skipped + 1
                    details.Add Array(location, "保留", reason, "[原文不记录]", "[保持原值]", "")
                End If
            Next r
            r = 1
            Do While r <= UBound(data, 1)
                If flags(r, c) Then
                    first = r
                    Do While r <= UBound(data, 1)
                        If Not flags(r, c) Then Exit Do
                        r = r + 1
                    Loop
                    ReDim part(1 To r - first, 1 To 1)
                    For k = first To r - 1: part(k - first + 1, 1) = values(k, c): Next k
                    targets.Add block.Cells(first, c).Resize(r - first, 1): outputs.Add part
                    If targets.Count > DH_MAX_AREAS Then Err.Raise vbObjectError + 861, , "可处理数据过于碎片化，请分批。"
                Else
                    r = r + 1
                End If
            Loop
        Next c
    Next block
    DH_Freeze state: frozen = True
    audit = DH_SaveAudit(operation, wb, details)
    If targets.Count > 0 Then
        backup = DH_Backup(wb)
        Set snapshot = Application.Workbooks.Add(xlWBATWorksheet): Set cache = snapshot.Worksheets(1)
        For Each block In targets: block.Copy Destination:=cache.Range(block.Address): Next block
        For i = 1 To targets.Count
            Set block = targets(i): touched = i
            If mode = 0 Then
                block.NumberFormat = "yyyy-mm-dd": block.Value2 = outputs(i)
            Else
                block.NumberFormat = "@": block.Value2 = DH_LiteralMatrix(outputs(i))
            End If
        Next i
    End If
    DH_LastSummary = "修改=" & changed & "；保留常量=" & skipped & "；公式不写入；明细=" & audit
    If mode > 0 Then DH_LastSummary = DH_LastSummary & "；第1行及超级表表头/汇总行受保护"
    If ambiguous > 0 Then DH_LastSummary = DH_LastSummary & "；保留 " & ambiguous & " 处：月/日/年写法看不出是几月几日，请手动确认"
    DH_Log operation, wb, ws.Name, selected.Address(False, False), changed, "成功；" & DH_LastSummary, backup
    If Not snapshot Is Nothing Then snapshot.Close False
    DH_TransformSelection = changed
    DH_RestoreFocus ws, original: DH_Restore state
    Exit Function
Failed:
    message = Err.Description: code = Err.Number
    On Error Resume Next
    For i = 1 To touched
        Set block = targets(i): Err.Clear: cache.Range(block.Address).Copy Destination:=block
        If Err.Number <> 0 Then rollbackFailed = True
    Next i
    If Not snapshot Is Nothing Then snapshot.Close False
    On Error GoTo 0
    If rollbackFailed Then message = message & "；回滚不完整，请恢复基线：" & backup
    DH_RestoreFocus ws, original
    If frozen Then DH_Restore state
    DH_LogFailure operation, wb, "", "", message, backup
    Err.Raise code, "电子羊表格", message
End Function
