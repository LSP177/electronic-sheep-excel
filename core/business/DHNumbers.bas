Attribute VB_Name = "DHNumbers"
Option Explicit

' Invariant decimal syntax, conservative ID/date/precision rules; no Excel calls.
Public Function DH_TryNumber(ByVal text As String, ByRef number As Double, ByRef unit As String, _
    ByRef reason As String, ByRef looksNumeric As Boolean) As Boolean
    Static pattern As Object
    Dim body As String, token As Variant, unsigned As String, integerPart As String, digits As String
    Dim dot As Long, y As Long, m As Long, d As Long, dateValue As Date
    unit = "": reason = "": looksNumeric = False
    body = Trim$(Replace$(Replace$(text, ChrW(160), " "), ChrW(12288), " "))
    If Len(body) = 0 Or Len(body) > 128 Then reason = "空文本或文本过长": Exit Function
    If Left$(body, 1) = "=" Then reason = "等号开头的字面文本，保护原值": Exit Function
    For Each token In Array("kg", "元", "件", "箱", "个", "包", "台", "部", "辆", "套", "支", "瓶", "盒", "张", "把", "批")
        If Len(body) > Len(token) Then
            If StrComp(Right$(body, Len(token)), token, vbTextCompare) = 0 Then
                unit = Right$(body, Len(token)): body = Trim$(Left$(body, Len(body) - Len(token))): Exit For
            End If
        End If
    Next token
    If pattern Is Nothing Then
        Set pattern = CreateObject("VBScript.RegExp")
        pattern.Pattern = "^[+-]?([0-9]+(\.[0-9]+)?|\.[0-9]+)$"
        pattern.Global = False
    End If
    If Not pattern.Test(body) Then
        If Len(unit) > 0 Then reason = "混合单位或单位前不是纯数字" Else reason = "日期样式、含糊分隔符或非标准十进制文本"
        Exit Function
    End If
    looksNumeric = True
    unsigned = body
    If Left$(unsigned, 1) = "+" Or Left$(unsigned, 1) = "-" Then unsigned = Mid$(unsigned, 2)
    dot = InStr(unsigned, ".")
    If dot > 0 Then integerPart = Left$(unsigned, dot - 1) Else integerPart = unsigned
    If Len(integerPart) > 1 And Left$(integerPart, 1) = "0" Then reason = "前导零编号，保留文本": Exit Function
    If Len(unit) = 0 And dot = 0 And Len(integerPart) >= 10 Then reason = "10位及以上整数文本，按长编号保护": Exit Function
    If Len(unit) = 0 And dot = 0 And Len(integerPart) = 8 Then
        y = CLng(Left$(integerPart, 4)): m = CLng(Mid$(integerPart, 5, 2)): d = CLng(Right$(integerPart, 2))
        If y >= 1900 And y <= 2100 And m >= 1 And m <= 12 And d >= 1 And d <= 31 Then
            dateValue = DateSerial(y, m, d)
            If Year(dateValue) = y And Month(dateValue) = m And Day(dateValue) = d Then reason = "YYYYMMDD日期样式，保留文本": Exit Function
        End If
    End If
    If Len(unit) = 0 And dot = 5 And Len(unsigned) = 7 Then
        y = CLng(integerPart): m = CLng(Right$(unsigned, 2))
        If y >= 1900 And y <= 2100 And m >= 1 And m <= 12 Then reason = "年.月日期样式，保留文本": Exit Function
    End If
    digits = Replace$(unsigned, ".", "")
    Do While Len(digits) > 1 And Left$(digits, 1) = "0": digits = Mid$(digits, 2): Loop
    If Len(digits) > 15 Then reason = "超过15位有效数字，避免精度损失": Exit Function
    number = Val(body)
    DH_TryNumber = True
End Function

Public Function DH_ObservedUnit(ByVal text As String) As String
    Static pattern As Object, scientific As Object
    Dim body As String, matches As Object
    body = Trim$(Replace$(Replace$(text, ChrW(160), " "), ChrW(12288), " "))
    If Len(body) > 128 Then Exit Function
    If scientific Is Nothing Then
        Set scientific = CreateObject("VBScript.RegExp")
        scientific.Pattern = "^[+-]?([0-9]+(\.[0-9]+)?|\.[0-9]+)[Ee][+-]?[0-9]+$"
    End If
    If scientific.Test(body) Then Exit Function
    If pattern Is Nothing Then
        Set pattern = CreateObject("VBScript.RegExp"): pattern.IgnoreCase = True
        pattern.Pattern = "^[+-]?([0-9]+(\.[0-9]+)?|\.[0-9]+)\s*([A-Za-z\u4e00-\u9fff%].*)$"
    End If
    If pattern.Test(body) Then
        Set matches = pattern.Execute(body)
        DH_ObservedUnit = LCase$(Trim$(matches(0).SubMatches(2)))
    End If
End Function

Private Function ColumnName(ByVal column As Long) As String
    Do While column > 0
        column = column - 1
        ColumnName = Chr$(65 + column Mod 26) & ColumnName
        column = column \ 26
    Loop
End Function

Public Function DH_ConvertNumbers(ByVal selected As Range, Optional ByVal dryRun As Boolean = False, _
    Optional ByVal allowMixed As Boolean = False) As Long
    Dim originalSelection As Range
    Dim ws As Worksheet, wb As Workbook, blocks As Collection, block As Range, records As Collection
    Dim targets As Collection, outputs As Collection, details As Collection, unitKinds As Object, unitCounts As Object, unknownColumns As Object
    Dim source As Variant, numbers As Variant, flags As Variant, reasons As Variant, units As Variant, record As Variant, output As Variant
    Dim r As Long, c As Long, i As Long, k As Long, startRow As Long, touched As Long, converted As Long, skipped As Long
    Dim value As Double, unit As String, reason As String, looks As Boolean, merged As Variant, isMerged As Boolean
    Dim columnKey As String, unitKey As String, location As String, columnLabel As String, backup As String, auditPath As String
    Dim snapshot As Workbook, cache As Worksheet, report As Worksheet, state As DH_AppState, frozen As Boolean
    Dim message As String, code As Long, summary As String, rollbackFailed As Boolean
    Dim blockRow As Long, blockColumn As Long, mixedColumn As Boolean, observed As String
    ' allowMixed (Unit:Force, one run only): convert mixed-unit columns too and say which units were dropped.
    Dim removedUnits As Object, keptColumns As Object, userNote As String
    Dim inMerge() As Boolean, mergeArea As Range, overlap As Range, mr As Long, mc As Long
    On Error GoTo Failed
    DH_RequireSelection selected
    Set ws = selected.Parent: Set wb = ws.Parent
    Set originalSelection = DH_OriginalSelection(ws)
    Set blocks = DH_TextBlocks(selected): Set records = New Collection
    Set targets = New Collection: Set outputs = New Collection: Set details = New Collection
    Set unitKinds = CreateObject("Scripting.Dictionary"): Set unitCounts = CreateObject("Scripting.Dictionary")
    Set unknownColumns = CreateObject("Scripting.Dictionary")
    Set removedUnits = CreateObject("Scripting.Dictionary"): Set keptColumns = CreateObject("Scripting.Dictionary")
    For Each block In blocks
        blockRow = block.Row: blockColumn = block.Column
        source = DH_Matrix(block)
        ReDim numbers(1 To UBound(source, 1), 1 To UBound(source, 2))
        ReDim flags(1 To UBound(source, 1), 1 To UBound(source, 2))
        ReDim reasons(1 To UBound(source, 1), 1 To UBound(source, 2))
        ReDim units(1 To UBound(source, 1), 1 To UBound(source, 2))
        merged = block.MergeCells: isMerged = True
        If Not IsNull(merged) Then isMerged = CBool(merged)
        ' R30-OBS-1: mark only the cells that are themselves merged; the rest of the
        ' block is converted as usual (a merge elsewhere in the block no longer blocks it).
        ReDim inMerge(1 To UBound(source, 1), 1 To UBound(source, 2))
        If isMerged Then
            For Each mergeArea In DH_MergedAreas(block)
                Set overlap = Application.Intersect(mergeArea, block)
                If Not overlap Is Nothing Then
                    For mr = overlap.Row - blockRow + 1 To overlap.Row - blockRow + overlap.Rows.Count
                        For mc = overlap.Column - blockColumn + 1 To overlap.Column - blockColumn + overlap.Columns.Count
                            inMerge(mr, mc) = True
                        Next mc
                    Next mr
                End If
            Next mergeArea
        End If
        For c = 1 To UBound(source, 2)
            columnKey = CStr(blockColumn + c - 1)
            For r = 1 To UBound(source, 1)
                If VarType(source(r, c)) = vbString Then
                    If inMerge(r, c) Then
                        reasons(r, c) = "合并单元格不转换；拆开合并后可再处理"
                    Else
                        flags(r, c) = DH_TryNumber(CStr(source(r, c)), value, unit, reason, looks)
                        numbers(r, c) = value: reasons(r, c) = reason: units(r, c) = unit
                        observed = DH_ObservedUnit(CStr(source(r, c)))
                        If Len(observed) > 0 Then
                            If StrComp(observed, unit, vbTextCompare) <> 0 Then unknownColumns(columnKey) = True
                            unitKey = columnKey & "|" & observed
                            If Not unitKinds.Exists(unitKey) Then
                                unitKinds.Add unitKey, True
                                If Not unitCounts.Exists(columnKey) Then unitCounts.Add columnKey, 0
                                unitCounts(columnKey) = unitCounts(columnKey) + 1
                            End If
                        End If
                    End If
                End If
            Next r
        Next c
        records.Add Array(source, numbers, flags, reasons, units)
    Next block
    For i = 1 To blocks.Count
        Set block = blocks(i): record = records(i)
        blockRow = block.Row: blockColumn = block.Column
        source = record(0): numbers = record(1): flags = record(2): reasons = record(3): units = record(4)
        For c = 1 To UBound(source, 2)
            columnKey = CStr(blockColumn + c - 1): columnLabel = ColumnName(blockColumn + c - 1)
            mixedColumn = unknownColumns.Exists(columnKey)
            If unitCounts.Exists(columnKey) Then mixedColumn = mixedColumn Or (unitCounts(columnKey) > 1)
            For r = 1 To UBound(source, 1)
                If VarType(source(r, c)) = vbString Then
                    If mixedColumn And Not allowMixed Then
                        If flags(r, c) = True Then
                            keptColumns(columnLabel) = True
                            flags(r, c) = False: reasons(r, c) = "同列混合/未知单位，保留该列全部数字候选"
                        End If
                    End If
                    location = columnLabel & CStr(blockRow + r - 1)
                    If flags(r, c) = True Then
                        converted = converted + 1
                        If Len(units(r, c)) > 0 Then removedUnits(CStr(units(r, c))) = True
                        details.Add Array(location, "可转换", "明确十进制数值", source(r, c), numbers(r, c), units(r, c))
                    Else
                        skipped = skipped + 1
                        details.Add Array(location, "保留", reasons(r, c), source(r, c), source(r, c), units(r, c))
                    End If
                End If
            Next r
            r = 1
            Do While r <= UBound(source, 1)
                If flags(r, c) = True Then
                    startRow = r
                    Do While r <= UBound(source, 1)
                        If flags(r, c) <> True Then Exit Do
                        r = r + 1
                    Loop
                    ReDim output(1 To r - startRow, 1 To 1)
                    For k = startRow To r - 1: output(k - startRow + 1, 1) = numbers(k, c): Next k
                    targets.Add block.Cells(startRow, c).Resize(r - startRow, 1)
                    outputs.Add output
                    If targets.Count > DH_MAX_AREAS Then Err.Raise vbObjectError + 790, , "可转换数据过于碎片化，请分批处理。"
                Else
                    r = r + 1
                End If
            Loop
        Next c
    Next i
    If dryRun Then DH_ConvertNumbers = converted: Exit Function
    DH_Log "数字转换", wb, ws.Name, selected.Address(False, False), 0, "开始；安全候选=" & converted
    DH_Freeze state: frozen = True
    auditPath = DH_SaveAudit("数字转换", wb, details)
    If targets.Count > 0 Then
        backup = DH_Backup(wb)
        Set snapshot = Application.Workbooks.Add(xlWBATWorksheet): Set cache = snapshot.Worksheets(1)
        For Each block In targets: block.Copy Destination:=cache.Range(block.Address): Next block
        For i = 1 To targets.Count
            Set block = targets(i): touched = i
            block.NumberFormat = "General": block.Value2 = outputs(i)
        Next i
    End If
    summary = "转换单元格=" & converted & "；保留文本=" & skipped & "；公式/原有数值未写入"
    If keptColumns.Count > 0 Then userNote = "同列单位不统一，整列保留：" & Join(keptColumns.Keys, "、") & " 列"
    If removedUnits.Count > 0 Then
        If Len(userNote) > 0 Then userNote = userNote & "；"
        userNote = userNote & "本次去掉了单位：" & Join(removedUnits.Keys, "/")
    End If
    If Len(userNote) > 0 Then summary = summary & "；" & userNote
    DH_LastSummary = userNote
    Set report = DH_ActionReport(wb, "数字转换", ws.Name & "!" & selected.Address(False, False), summary, details, backup, auditPath)
    DH_Log "数字转换", wb, ws.Name, selected.Address(False, False), converted, "成功；" & summary & "；明细=" & auditPath, backup
    If Not snapshot Is Nothing Then snapshot.Close False
    DH_ConvertNumbers = converted
    DH_RestoreFocus ws, originalSelection
    DH_Restore state
    Exit Function
Failed:
    message = Err.Description: code = Err.Number
    On Error Resume Next
    For i = 1 To touched
        Err.Clear: Set block = targets(i)
        cache.Range(block.Address).Copy Destination:=block
        If Err.Number <> 0 Then rollbackFailed = True
    Next i
    If Not snapshot Is Nothing Then snapshot.Close False
    If Not report Is Nothing Then report.Range("A1").Value2 = "操作失败：请检查日志并恢复备份"
    On Error GoTo 0
    If rollbackFailed Then message = message & "；回滚不完整，请恢复基线：" & backup
    DH_RestoreFocus ws, originalSelection
    If frozen Then DH_Restore state
    DH_LogFailure "数字转换", wb, "", "", message, backup
    Err.Raise code, "电子羊表格", message
End Function
