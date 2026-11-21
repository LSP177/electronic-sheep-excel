Attribute VB_Name = "DHLookup"
Option Explicit

' Cross-sheet exact match and multi-column field completion (v1.0 Patch 2+).
' Reuses DH_Matrix, DH_FormulaMask, DH_ValueToken, DH_WithoutTableHeaders,
' DH_MergedAreas, DH_Backup, DH_SaveAudit, DH_Log and DH_Freeze/DH_Restore.

Private Function ColIdxToLetter(ByVal colIndex As Long) As String
    Dim n As Long
    n = colIndex
    Do While n > 0
        n = n - 1
        ColIdxToLetter = Chr$(65 + (n Mod 26)) & ColIdxToLetter
        n = n \ 26
    Loop
End Function

Public Function DH_ColLetterToIndex(ByVal colLetter As String) As Long
    Dim text As String, i As Long, ch As Long, acc As Long
    text = UCase$(Trim$(colLetter))
    If Len(text) < 1 Or Len(text) > 3 Then Err.Raise vbObjectError + 940, , "列标识无效，请使用A至XFD之间的列字母。"
    For i = 1 To Len(text)
        ch = Asc(Mid$(text, i, 1))
        If ch < 65 Or ch > 90 Then Err.Raise vbObjectError + 940, , "列标识必须为英文字母A至Z。"
        acc = acc * 26 + (ch - 64)
    Next i
    If acc < 1 Or acc > 16384 Then Err.Raise vbObjectError + 940, , "列号超出Excel允许范围(A至XFD)。"
    DH_ColLetterToIndex = acc
End Function

Private Function ParseColIndices(ByVal colSpecCsv As String, ByVal minCount As Long, ByVal maxCount As Long, ByVal fieldLabel As String) As Long()
    Dim raw As String, parts As Variant, i As Long, idx As Long, seen As Object, out() As Long
    raw = Trim$(colSpecCsv)
    If Len(raw) = 0 Then Err.Raise vbObjectError + 940, , fieldLabel & "不能为空。"
    parts = Split(raw, ",")
    If UBound(parts) + 1 < minCount Or UBound(parts) + 1 > maxCount Then
        Err.Raise vbObjectError + 940, , fieldLabel & "列数量必须在 " & minCount & " 至 " & maxCount & " 之间。"
    End If
    Set seen = CreateObject("Scripting.Dictionary")
    ReDim out(0 To UBound(parts))
    For i = 0 To UBound(parts)
        idx = DH_ColLetterToIndex(CStr(parts(i)))
        If seen.Exists(CStr(idx)) Then Err.Raise vbObjectError + 940, , fieldLabel & "中包含重复列：" & ColIdxToLetter(idx)
        seen.Add CStr(idx), True
        out(i) = idx
    Next i
    ParseColIndices = out
End Function

Private Function FindBusinessSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    DH_RequireWorkbook wb
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
            If ws.Visible <> xlSheetVisible Or DH_IsReportSheet(ws) Or DH_IsPreviousReport(ws) Then
                Err.Raise vbObjectError + 941, , "不能使用隐藏工作表或体检报告表作为跨表匹配来源：" & sheetName
            End If
            Set FindBusinessSheet = ws
            Exit Function
        End If
    Next ws
    Err.Raise vbObjectError + 941, , "未找到指定的来源工作表：" & sheetName
End Function

Private Function LookupKeyToken(ByVal rawValue As Variant, ByRef isBlank As Boolean, ByRef isErrorVal As Boolean, ByRef kindTag As String, ByRef normNumText As String) As String
    Dim s As String
    isBlank = False: isErrorVal = False: kindTag = "": normNumText = ""
    If IsError(rawValue) Then
        isErrorVal = True
        LookupKeyToken = "ERR:" & CStr(rawValue)
        Exit Function
    End If
    If IsEmpty(rawValue) Or IsNull(rawValue) Then
        isBlank = True
        Exit Function
    End If
    If VarType(rawValue) = vbString Then
        s = Trim$(CStr(rawValue))
        If Len(s) = 0 Then
            isBlank = True
            Exit Function
        End If
        kindTag = "S"
        LookupKeyToken = "S:" & s
        If IsNumeric(s) Then
            On Error Resume Next
            normNumText = CStr(CDbl(s))
            If Err.Number <> 0 Then normNumText = "": Err.Clear
            On Error GoTo 0
        End If
    ElseIf VarType(rawValue) = vbBoolean Then
        kindTag = "B"
        LookupKeyToken = "B:" & IIf(CBool(rawValue), "1", "0")
    ElseIf IsNumeric(rawValue) Or IsDate(rawValue) Then
        kindTag = "N"
        normNumText = CStr(CDbl(rawValue))
        LookupKeyToken = "N:" & normNumText
    Else
        s = Trim$(CStr(rawValue))
        If Len(s) = 0 Then
            isBlank = True
            Exit Function
        End If
        kindTag = "S"
        LookupKeyToken = "S:" & s
    End If
End Function

' MATCH(, ,0) treats wildcard text and case differently from LookupKeyToken,
' and does not trim source keys. Never write a formula for another source row.
' Static mode keeps the existing typed-key rules; formula mode fails closed.
Private Function FormulaMatchAgrees(ByVal rawKey As Variant, ByVal keys As Range, ByVal expectedRow As Long, Optional ByVal checkFormula As Boolean = True) As Boolean
    Dim matched As Variant
    If Not checkFormula Then FormulaMatchAgrees = True: Exit Function
    matched = Application.Match(rawKey, keys, 0)
    If IsError(matched) Then Exit Function
    FormulaMatchAgrees = (CLng(matched) + keys.Row - 1 = expectedRow)
End Function

Private Function ItemsToArray(ByVal items As Collection) As Variant
    Dim arr() As Variant, i As Long
    If items.Count = 0 Then ItemsToArray = Array(): Exit Function
    ReDim arr(0 To items.Count - 1)
    For i = 1 To items.Count: arr(i - 1) = items(i): Next i
    ItemsToArray = arr
End Function

Private Function ReadCleanHeader(ByVal ws As Worksheet, ByVal rowIdx As Long, ByVal colIdx As Long, ByVal fallbackPrefix As String) As String
    Dim rawVal As Variant, s As String
    rawVal = ws.Cells(rowIdx, colIdx).Value2
    If IsError(rawVal) Then
        ReadCleanHeader = fallbackPrefix & ColIdxToLetter(colIdx)
    ElseIf IsEmpty(rawVal) Or IsNull(rawVal) Then
        ReadCleanHeader = fallbackPrefix & ColIdxToLetter(colIdx)
    Else
        s = Trim$(CStr(rawVal))
        If Len(s) = 0 Then s = fallbackPrefix & ColIdxToLetter(colIdx)
        ReadCleanHeader = s
    End If
End Function

Private Function BuildAutoHeaderNames(ByVal tgtSheet As Worksheet, ByVal tgtHeaderRow As Long, ByVal tgtFirstCol As Long, ByVal tgtLastCol As Long, ByVal srcSheet As Worksheet, ByVal srcHeaderRow As Long, ByRef retCols() As Long) As String()
    Dim existing As Object, c As Long, h As String, i As Long, candidate As String, suffix As Long, out() As String
    Set existing = CreateObject("Scripting.Dictionary")
    existing.CompareMode = 1
    For c = tgtFirstCol To tgtLastCol
        h = Trim$(CStr(tgtSheet.Cells(tgtHeaderRow, c).Text))
        If Len(h) > 0 Then existing(h) = True
    Next c
    ReDim out(0 To UBound(retCols))
    For i = 0 To UBound(retCols)
        candidate = ReadCleanHeader(srcSheet, srcHeaderRow, retCols(i), "补全_")
        If existing.Exists(candidate) Then
            candidate = srcSheet.Name & "_" & candidate
        End If
        If existing.Exists(candidate) Then
            suffix = 2
            Do While existing.Exists(candidate & "_" & CStr(suffix))
                suffix = suffix + 1
            Loop
            candidate = candidate & "_" & CStr(suffix)
        End If
        existing(candidate) = True
        out(i) = candidate
    Next i
    BuildAutoHeaderNames = out
End Function

Private Function PreflightRightAppend(ByVal tgtSheet As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal startCol As Long, ByVal colCount As Long) As String
    Dim endCol As Long, appendRange As Range, lo As ListObject, hit As Range, r As Long, c As Long, cell As Range
    endCol = startCol + colCount - 1
    If startCol < 1 Or endCol > 16384 Then
        PreflightRightAppend = "目标表右侧空间不足（超出 Excel 最大列 XFD），无法自动新增 " & colCount & " 列。"
        Exit Function
    End If
    Set appendRange = tgtSheet.Range(tgtSheet.Cells(firstRow, startCol), tgtSheet.Cells(lastRow, endCol))
    For Each lo In tgtSheet.ListObjects
        Set hit = Nothing
        On Error Resume Next
        Set hit = Application.Intersect(appendRange, lo.Range)
        On Error GoTo 0
        If Not hit Is Nothing Then
            PreflightRightAppend = "目标表右侧区域与 Excel 超级表「" & lo.Name & "」重叠，为保护表结构已中止自动加列，请改用「填入指定列」。"
            Exit Function
        End If
    Next lo
    If tgtSheet.ListObjects.Count > 0 Then
        For Each lo In tgtSheet.ListObjects
            If lo.Range.Row <= lastRow And (lo.Range.Row + lo.Range.Rows.Count - 1) >= firstRow Then
                If lo.Range.Column + lo.Range.Columns.Count - 1 = startCol - 1 Then
                    PreflightRightAppend = "目标区域属于 Excel 超级表「" & lo.Name & "」，在右侧自动新增列可能影响超级表结构，请先转换为普通区域或改用「填入指定列」。"
                    Exit Function
                End If
            End If
        Next lo
    End If
    If Not IsNull(appendRange.MergeCells) Then
        If CBool(appendRange.MergeCells) Then
            PreflightRightAppend = "目标表右侧新增列区域（" & ColIdxToLetter(startCol) & " 至 " & ColIdxToLetter(endCol) & " 列）包含合并单元格，无法安全新增列。"
            Exit Function
        End If
    Else
        PreflightRightAppend = "目标表右侧新增列区域（" & ColIdxToLetter(startCol) & " 至 " & ColIdxToLetter(endCol) & " 列）包含合并单元格，无法安全新增列。"
        Exit Function
    End If
    For r = firstRow To lastRow
        For c = startCol To endCol
            Set cell = tgtSheet.Cells(r, c)
            If cell.HasFormula Or Not IsEmpty(cell.Value2) Then
                PreflightRightAppend = "目标表右侧 " & ColIdxToLetter(c) & r & " 单元格已有内容或公式，为防覆盖已中止自动加列。"
                Exit Function
            End If
        Next c
    Next r
    PreflightRightAppend = ""
End Function

Public Function DH_LookupMeta(ByVal wb As Workbook) As Variant
    Dim ws As Worksheet, bounds As Range, sheets As New Collection, cols As Collection
    Dim c As Long, firstCol As Long, lastCol As Long, headerRow As Long, label As String, rawVal As Variant
    DH_RequireWorkbook wb
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible And Not DH_IsReportSheet(ws) And Not DH_IsPreviousReport(ws) Then
            Set bounds = DH_ContentRange(ws)
            Set cols = New Collection
            If bounds Is Nothing Then
                sheets.Add Array(ws.Name, "", 0, Array())
            Else
                headerRow = bounds.Row
                firstCol = bounds.Column
                lastCol = bounds.Column + bounds.Columns.Count - 1
                If lastCol - firstCol + 1 > 64 Then lastCol = firstCol + 63
                For c = firstCol To lastCol
                    rawVal = ws.Cells(headerRow, c).Value2
                    If IsError(rawVal) Then
                        label = "(错误值)"
                    ElseIf IsEmpty(rawVal) Or IsNull(rawVal) Then
                        label = "(空白)"
                    Else
                        label = Trim$(CStr(rawVal))
                        If Len(label) = 0 Then label = "(空白)"
                        If Len(label) > 32 Then label = Left$(label, 29) & "..."
                    End If
                    cols.Add Array(ColIdxToLetter(c), label)
                Next c
                sheets.Add Array(ws.Name, bounds.Address(False, False), bounds.Rows.Count, ItemsToArray(cols))
            End If
        End If
    Next ws
    DH_LookupMeta = ItemsToArray(sheets)
End Function

Public Sub DH_ParseLookupSpec(ByVal spec As String, ByRef srcSheetName As String, ByRef srcKeyCol As Long, ByRef retColCsv As String, ByRef tgtKeyCol As Long, ByRef tgtWriteSpec As String, ByRef dupRule As String, ByRef writeMode As String)
    Dim parts As Variant, retCols() As Long, writeCols() As Long, i As Long
    parts = Split(spec, ":")
    If UBound(parts) <> 7 Or CStr(parts(0)) <> "Lookup" Then
        Err.Raise vbObjectError + 942, , "跨表匹配参数格式无效。"
    End If
    srcSheetName = Trim$(CStr(parts(1)))
    If Len(srcSheetName) = 0 Or Len(srcSheetName) > 31 Then Err.Raise vbObjectError + 942, , "来源工作表名称无效。"
    srcKeyCol = DH_ColLetterToIndex(CStr(parts(2)))
    retCols = ParseColIndices(CStr(parts(3)), 1, 16, "来源返回列")
    retColCsv = ""
    For i = 0 To UBound(retCols)
        If retCols(i) = srcKeyCol Then Err.Raise vbObjectError + 942, , "来源表的匹配键列与返回列不能相同。"
        If Len(retColCsv) > 0 Then retColCsv = retColCsv & ","
        retColCsv = retColCsv & ColIdxToLetter(retCols(i))
    Next i
    tgtKeyCol = DH_ColLetterToIndex(CStr(parts(4)))
    If StrComp(Trim$(CStr(parts(5))), "AUTO", vbTextCompare) = 0 Then
        tgtWriteSpec = "AUTO"
    Else
        writeCols = ParseColIndices(CStr(parts(5)), 1, 16, "目标写入列")
        If UBound(writeCols) <> UBound(retCols) Then
            Err.Raise vbObjectError + 942, , "目标写入列数量必须与来源返回列数量一致，或设为 AUTO。"
        End If
        tgtWriteSpec = ""
        For i = 0 To UBound(writeCols)
            If writeCols(i) = tgtKeyCol Then Err.Raise vbObjectError + 942, , "目标表的匹配键列与写入结果列不能相同。"
            If Len(tgtWriteSpec) > 0 Then tgtWriteSpec = tgtWriteSpec & ","
            tgtWriteSpec = tgtWriteSpec & ColIdxToLetter(writeCols(i))
        Next i
    End If
    dupRule = LCase$(Trim$(CStr(parts(6))))
    If dupRule <> "abort" And dupRule <> "first" And dupRule <> "skip" Then
        Err.Raise vbObjectError + 942, , "重复键处理规则无效，仅支持 abort / first / skip。"
    End If
    writeMode = LCase$(Trim$(CStr(parts(7))))
    If writeMode <> "static" And writeMode <> "formula" Then
        Err.Raise vbObjectError + 942, , "写入模式无效，仅支持 static（静态值）或 formula（动态公式）。"
    End If
End Sub

Public Function DH_LookupPreview(ByVal tgtSheet As Worksheet, ByVal spec As String) As Variant
    Dim srcSheetName As String, srcKeyCol As Long, retColCsv As String, tgtKeyCol As Long, tgtWriteSpec As String, dupRule As String, writeMode As String
    Dim retCols() As Long, writeCols() As Long, autoHeaders() As String, numCols As Long, isAutoAppend As Boolean
    Dim srcSheet As Worksheet, srcBounds As Range, tgtBounds As Range
    Dim srcFirstRow As Long, srcLastRow As Long, tgtFirstRow As Long, tgtLastRow As Long, tgtFirstCol As Long, tgtLastCol As Long
    Dim keyFirstRow As Object, keyValues As Object, keyCounts As Object, keyConflicting As Object, normNumMap As Object
    Dim r As Long, cIdx As Long, rawKey As Variant, rawRet As Variant, tok As String, retTok As String
    Dim isBlank As Boolean, isErr As Boolean, kindTag As String, normNum As String
    Dim totalTgtRows As Long, matchedRows As Long, unmatchedRows As Long, blankKeyRows As Long
    Dim dupSourceKeys As Long, conflictingDupKeys As Long, typeMismatchRows As Long, protectedTargetRows As Long, sourceErrorCount As Long
    Dim canExecute As Boolean, blockReason As String, rightErr As String, formulaSkippedRows As Long
    Dim samples As New Collection, anomalies As New Collection
    Dim srcRow As Long, tgtCell As Range, sampleText As String, writeSummary As String, colHeaderLabel As String

    DH_RequireWorkbook tgtSheet.Parent
    DH_ParseLookupSpec spec, srcSheetName, srcKeyCol, retColCsv, tgtKeyCol, tgtWriteSpec, dupRule, writeMode
    If StrComp(tgtSheet.Name, srcSheetName, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 943, , "来源工作表与目标工作表不能为同一张表。"
    End If
    Set srcSheet = FindBusinessSheet(tgtSheet.Parent, srcSheetName)
    Set srcBounds = DH_ContentRange(srcSheet)
    Set tgtBounds = DH_ContentRange(tgtSheet)
    If srcBounds Is Nothing Then Err.Raise vbObjectError + 944, , "来源工作表「" & srcSheet.Name & "」为空，无法匹配。"
    If tgtBounds Is Nothing Then Err.Raise vbObjectError + 944, , "当前目标工作表「" & tgtSheet.Name & "」为空，无法补全。"

    retCols = ParseColIndices(retColCsv, 1, 16, "来源返回列")
    numCols = UBound(retCols) + 1

    srcFirstRow = srcBounds.Row
    srcLastRow = srcBounds.Row + srcBounds.Rows.Count - 1
    tgtFirstRow = tgtBounds.Row
    tgtLastRow = tgtBounds.Row + tgtBounds.Rows.Count - 1
    tgtFirstCol = tgtBounds.Column
    tgtLastCol = tgtBounds.Column + tgtBounds.Columns.Count - 1

    If srcLastRow <= srcFirstRow Then Err.Raise vbObjectError + 944, , "来源工作表仅有表头行，无有效数据记录。"
    If tgtLastRow <= tgtFirstRow Then Err.Raise vbObjectError + 944, , "目标工作表仅有表头行，无待补全的数据行。"
    If (tgtLastRow - tgtFirstRow) * numCols > 100000 Then Err.Raise vbObjectError + 944, , "单次跨表补全写入单元格超过10万格上限，请分批处理。"

    isAutoAppend = (StrComp(tgtWriteSpec, "AUTO", vbTextCompare) = 0)
    ReDim writeCols(0 To numCols - 1)
    If isAutoAppend Then
        rightErr = PreflightRightAppend(tgtSheet, tgtFirstRow, tgtLastRow, tgtLastCol + 1, numCols)
        autoHeaders = BuildAutoHeaderNames(tgtSheet, tgtFirstRow, tgtFirstCol, tgtLastCol, srcSheet, srcFirstRow, retCols)
        writeSummary = ""
        For cIdx = 0 To numCols - 1
            writeCols(cIdx) = tgtLastCol + 1 + cIdx
            If Len(writeSummary) > 0 Then writeSummary = writeSummary & ","
            writeSummary = writeSummary & ColIdxToLetter(writeCols(cIdx)) & "(" & autoHeaders(cIdx) & ")"
        Next cIdx
    Else
        writeCols = ParseColIndices(tgtWriteSpec, numCols, numCols, "目标写入列")
        writeSummary = tgtWriteSpec
    End If

    Set keyFirstRow = CreateObject("Scripting.Dictionary")
    Set keyValues = CreateObject("Scripting.Dictionary")
    Set keyCounts = CreateObject("Scripting.Dictionary")
    Set keyConflicting = CreateObject("Scripting.Dictionary")
    Set normNumMap = CreateObject("Scripting.Dictionary")

    For r = srcFirstRow + 1 To srcLastRow
        rawKey = srcSheet.Cells(r, srcKeyCol).Value2
        tok = LookupKeyToken(rawKey, isBlank, isErr, kindTag, normNum)
        If Not isBlank And Not isErr Then
            If Not keyFirstRow.Exists(tok) Then
                keyFirstRow.Add tok, r
                keyCounts.Add tok, 1
                For cIdx = 0 To numCols - 1
                    rawRet = srcSheet.Cells(r, retCols(cIdx)).Value2
                    retTok = DH_ValueToken(rawRet)
                    keyValues.Add tok & "|" & CStr(cIdx), retTok
                Next cIdx
                If Len(normNum) > 0 Then
                    If Not normNumMap.Exists(normNum) Then normNumMap.Add normNum, tok
                End If
            Else
                keyCounts(tok) = CLng(keyCounts(tok)) + 1
                If CLng(keyCounts(tok)) = 2 Then dupSourceKeys = dupSourceKeys + 1
                For cIdx = 0 To numCols - 1
                    rawRet = srcSheet.Cells(r, retCols(cIdx)).Value2
                    retTok = DH_ValueToken(rawRet)
                    If StrComp(CStr(keyValues(tok & "|" & CStr(cIdx))), retTok, vbBinaryCompare) <> 0 Then
                        If Not keyConflicting.Exists(tok) Then
                            keyConflicting.Add tok, r
                            conflictingDupKeys = conflictingDupKeys + 1
                        End If
                    End If
                Next cIdx
            End If
        End If
    Next r

    For r = tgtFirstRow + 1 To tgtLastRow
        totalTgtRows = totalTgtRows + 1
        rawKey = tgtSheet.Cells(r, tgtKeyCol).Value2
        tok = LookupKeyToken(rawKey, isBlank, isErr, kindTag, normNum)

        Dim rowHasWritable As Boolean, rowAllProtected As Boolean
        rowHasWritable = False: rowAllProtected = True
        For cIdx = 0 To numCols - 1
            Set tgtCell = tgtSheet.Cells(r, writeCols(cIdx))
            If CBool(tgtCell.MergeCells) Or tgtCell.HasFormula Or Not IsEmpty(tgtCell.Value2) Then
                protectedTargetRows = protectedTargetRows + 1
                If anomalies.Count < 15 Then
                    anomalies.Add Array(ColIdxToLetter(writeCols(cIdx)) & r, "目标格保护跳过", "目标单元格已有内容、公式或处于合并区域，不覆盖")
                End If
            Else
                rowHasWritable = True
                rowAllProtected = False
            End If
        Next cIdx

        If Not rowAllProtected Then
            If isBlank Then
                blankKeyRows = blankKeyRows + 1
                If anomalies.Count < 15 Then
                    anomalies.Add Array(ColIdxToLetter(tgtKeyCol) & r, "匹配键为空", "目标行匹配键为空白，跳过不补全")
                End If
            ElseIf isErr Then
                unmatchedRows = unmatchedRows + 1
                If anomalies.Count < 15 Then
                    anomalies.Add Array(ColIdxToLetter(tgtKeyCol) & r, "匹配键为错误值", "目标行匹配键包含公式错误值，跳过")
                End If
            ElseIf keyFirstRow.Exists(tok) Then
                If CLng(keyCounts(tok)) > 1 And dupRule = "skip" Then
                    unmatchedRows = unmatchedRows + 1
                    If anomalies.Count < 15 Then
                        anomalies.Add Array(ColIdxToLetter(tgtKeyCol) & r, "重复键按规则跳过", "来源表存在 " & CLng(keyCounts(tok)) & " 条相同键值（行 " & CLng(keyFirstRow(tok)) & " 起）")
                    End If
                ElseIf Not FormulaMatchAgrees(rawKey, srcSheet.Range(srcSheet.Cells(srcFirstRow + 1, srcKeyCol), srcSheet.Cells(srcLastRow, srcKeyCol)), CLng(keyFirstRow(tok)), writeMode = "formula") Then
                    formulaSkippedRows = formulaSkippedRows + 1
                    unmatchedRows = unmatchedRows + 1
                    If anomalies.Count < 15 Then anomalies.Add Array(ColIdxToLetter(tgtKeyCol) & r, "公式匹配规则不一致（跳过）", "Excel MATCH 与静态匹配定位不同；请改用静态值模式，避免大小写、通配符或两端空格造成错配。")
                Else
                    srcRow = CLng(keyFirstRow(tok))
                    sampleText = ""
                    Dim rowWroteAny As Boolean
                    rowWroteAny = False
                    For cIdx = 0 To numCols - 1
                        Set tgtCell = tgtSheet.Cells(r, writeCols(cIdx))
                        If Not (CBool(tgtCell.MergeCells) Or tgtCell.HasFormula Or Not IsEmpty(tgtCell.Value2)) Then
                            rawRet = srcSheet.Cells(srcRow, retCols(cIdx)).Value2
                            If IsError(rawRet) Then
                                sourceErrorCount = sourceErrorCount + 1
                                If anomalies.Count < 15 Then
                                    colHeaderLabel = ReadCleanHeader(srcSheet, srcFirstRow, retCols(cIdx), ColIdxToLetter(retCols(cIdx)))
                                    anomalies.Add Array(ColIdxToLetter(writeCols(cIdx)) & r, "来源含错误值（将跳过）", "来源表「" & srcSheet.Name & "」" & ColIdxToLetter(retCols(cIdx)) & srcRow & "（" & colHeaderLabel & "）为错误值 " & CStr(srcSheet.Cells(srcRow, retCols(cIdx)).Text) & "，保护跳过不写入错误")
                                End If
                            Else
                                matchedRows = matchedRows + 1
                                rowWroteAny = True
                                If Len(sampleText) > 0 Then sampleText = sampleText & " | "
                                If numCols > 1 Then
                                    colHeaderLabel = ReadCleanHeader(srcSheet, srcFirstRow, retCols(cIdx), ColIdxToLetter(retCols(cIdx)))
                                    sampleText = sampleText & colHeaderLabel & "=" & CStr(srcSheet.Cells(srcRow, retCols(cIdx)).Text)
                                Else
                                    sampleText = CStr(srcSheet.Cells(srcRow, retCols(cIdx)).Text)
                                End If
                            End If
                        End If
                    Next cIdx
                    If rowWroteAny And samples.Count < 8 Then
                        samples.Add Array(ColIdxToLetter(writeCols(0)) & IIf(numCols > 1, ".." & ColIdxToLetter(writeCols(numCols - 1)), "") & ":" & r, CStr(tgtSheet.Cells(r, tgtKeyCol).Text), sampleText, srcSheet.Name & "!行" & srcRow)
                    End If
                End If
            Else
                unmatchedRows = unmatchedRows + 1
                If Len(normNum) > 0 And normNumMap.Exists(normNum) Then
                    typeMismatchRows = typeMismatchRows + 1
                    If anomalies.Count < 15 Then
                        anomalies.Add Array(ColIdxToLetter(tgtKeyCol) & r, "类型差异（文本/数字）", "键值「" & CStr(tgtSheet.Cells(r, tgtKeyCol).Text) & "」在两表中存储类型不同（前导零或文本数字差异），未擅自合并")
                    End If
                Else
                    If anomalies.Count < 15 Then
                        anomalies.Add Array(ColIdxToLetter(tgtKeyCol) & r, "来源表无此键值", "键值「" & CStr(tgtSheet.Cells(r, tgtKeyCol).Text) & "」在来源表中不存在")
                    End If
                End If
            End If
        End If
    Next r

    canExecute = True
    blockReason = ""
    If Len(rightErr) > 0 Then
        canExecute = False
        blockReason = rightErr
    ElseIf dupSourceKeys > 0 And dupRule = "abort" Then
        canExecute = False
        blockReason = "来源表匹配键列存在 " & dupSourceKeys & " 个重复键值（其中 " & conflictingDupKeys & " 个对应返回数据不一致）。为避免填错数据，已默认中止；请先清理来源表重复项，或显式选择“取首次出现行 (first)”或“跳过重复键 (skip)”。"
    ElseIf matchedRows = 0 Then
        canExecute = False
        If formulaSkippedRows > 0 Then
            blockReason = "公式匹配规则与静态匹配不一致，已保护跳过；请改用静态值模式（大小写、通配符或键值两端空格）。"
        ElseIf sourceErrorCount > 0 Then
            blockReason = "命中的来源记录均为公式错误值（共 " & sourceErrorCount & " 格），为保护目标表不被错误值污染已全部跳过。"
        Else
            blockReason = "未找到任何可写入的匹配记录（可能键值全部未命中、类型不一致或目标列已有内容受保护）。"
        End If
    End If

    DH_LookupPreview = Array( _
        tgtSheet.Parent.Name, _
        tgtSheet.Name, _
        srcSheet.Name, _
        ColIdxToLetter(srcKeyCol), _
        retColCsv, _
        totalTgtRows, _
        matchedRows, _
        unmatchedRows, _
        blankKeyRows, _
        dupSourceKeys, _
        conflictingDupKeys, _
        typeMismatchRows, _
        protectedTargetRows, _
        ColIdxToLetter(tgtKeyCol), _
        writeSummary, _
        dupRule & "/" & writeMode, _
        canExecute, _
        blockReason, _
        ItemsToArray(samples), _
        ItemsToArray(anomalies), _
        sourceErrorCount _
    )
End Function

Private Function SafeSheetFormulaPrefix(ByVal sheetName As String) As String
    SafeSheetFormulaPrefix = "'" & Replace(sheetName, "'", "''") & "'!"
End Function

Public Function DH_ExecuteLookup(ByVal tgtSheet As Worksheet, ByVal spec As String) As Long
    Dim wb As Workbook, originalSelection As Range
    Dim srcSheetName As String, srcKeyCol As Long, retColCsv As String, tgtKeyCol As Long, tgtWriteSpec As String, dupRule As String, writeMode As String
    Dim retCols() As Long, writeCols() As Long, autoHeaders() As String, numCols As Long, isAutoAppend As Boolean
    Dim srcSheet As Worksheet, srcBounds As Range, tgtBounds As Range
    Dim srcFirstRow As Long, srcLastRow As Long, tgtFirstRow As Long, tgtLastRow As Long, tgtFirstCol As Long, tgtLastCol As Long
    Dim keyFirstRow As Object, keyCounts As Object, normNumMap As Object
    Dim r As Long, cIdx As Long, rawKey As Variant, tok As String, isBlank As Boolean, isErr As Boolean, kindTag As String, normNum As String
    Dim dupSourceKeys As Long, srcRow As Long, tgtCell As Range, srcCell As Range
    Dim state As DH_AppState, frozen As Boolean, details As New Collection
    Dim snapshot As Workbook, cache As Worksheet, writeRange As Range, touched As Boolean, rollbackFailed As Boolean
    Dim backupPath As String, auditPath As String, rightErr As String, targetAddress As String
    Dim changedCells As Long, skippedProtected As Long, skippedUnmatched As Long, skippedDup As Long, skippedBlank As Long, skippedSrcErr As Long
    Dim oldToken As String, newToken As String, formulaExpr As String, srcPrefix As String, srcNumFmt As String, colHeaderLabel As String
    Dim formulaSafe As Boolean, skippedFormula As Long

    Set wb = tgtSheet.Parent
    DH_RequireWorkbook wb
    Set originalSelection = DH_OriginalSelection(tgtSheet)
    DH_ParseLookupSpec spec, srcSheetName, srcKeyCol, retColCsv, tgtKeyCol, tgtWriteSpec, dupRule, writeMode
    If StrComp(tgtSheet.Name, srcSheetName, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 943, , "来源工作表与目标工作表不能为同一张表。"
    End If
    Set srcSheet = FindBusinessSheet(wb, srcSheetName)
    Set srcBounds = DH_ContentRange(srcSheet)
    Set tgtBounds = DH_ContentRange(tgtSheet)
    If srcBounds Is Nothing Or tgtBounds Is Nothing Then
        Err.Raise vbObjectError + 944, , "来源表或目标表为空，无法执行跨表补全。"
    End If

    retCols = ParseColIndices(retColCsv, 1, 16, "来源返回列")
    numCols = UBound(retCols) + 1

    srcFirstRow = srcBounds.Row
    srcLastRow = srcBounds.Row + srcBounds.Rows.Count - 1
    tgtFirstRow = tgtBounds.Row
    tgtLastRow = tgtBounds.Row + tgtBounds.Rows.Count - 1
    tgtFirstCol = tgtBounds.Column
    tgtLastCol = tgtBounds.Column + tgtBounds.Columns.Count - 1

    If srcLastRow <= srcFirstRow Or tgtLastRow <= tgtFirstRow Then
        Err.Raise vbObjectError + 944, , "来源表或目标表缺少数据行，无法执行跨表补全。"
    End If
    If (tgtLastRow - tgtFirstRow) * numCols > 100000 Then
        Err.Raise vbObjectError + 944, , "单次跨表补全写入单元格超过10万格上限。"
    End If

    isAutoAppend = (StrComp(tgtWriteSpec, "AUTO", vbTextCompare) = 0)
    ReDim writeCols(0 To numCols - 1)
    If isAutoAppend Then
        rightErr = PreflightRightAppend(tgtSheet, tgtFirstRow, tgtLastRow, tgtLastCol + 1, numCols)
        If Len(rightErr) > 0 Then Err.Raise vbObjectError + 945, , rightErr
        autoHeaders = BuildAutoHeaderNames(tgtSheet, tgtFirstRow, tgtFirstCol, tgtLastCol, srcSheet, srcFirstRow, retCols)
        For cIdx = 0 To numCols - 1
            writeCols(cIdx) = tgtLastCol + 1 + cIdx
        Next cIdx
    Else
        writeCols = ParseColIndices(tgtWriteSpec, numCols, numCols, "目标写入列")
    End If
    targetAddress = ColIdxToLetter(writeCols(0)) & (tgtFirstRow + 1) & ":" & ColIdxToLetter(writeCols(numCols - 1)) & tgtLastRow

    Set keyFirstRow = CreateObject("Scripting.Dictionary")
    Set keyCounts = CreateObject("Scripting.Dictionary")
    Set normNumMap = CreateObject("Scripting.Dictionary")

    For r = srcFirstRow + 1 To srcLastRow
        rawKey = srcSheet.Cells(r, srcKeyCol).Value2
        tok = LookupKeyToken(rawKey, isBlank, isErr, kindTag, normNum)
        If Not isBlank And Not isErr Then
            If Not keyFirstRow.Exists(tok) Then
                keyFirstRow.Add tok, r
                keyCounts.Add tok, 1
                If Len(normNum) > 0 Then
                    If Not normNumMap.Exists(normNum) Then normNumMap.Add normNum, tok
                End If
            Else
                keyCounts(tok) = CLng(keyCounts(tok)) + 1
                If CLng(keyCounts(tok)) = 2 Then dupSourceKeys = dupSourceKeys + 1
            End If
        End If
    Next r

    If dupSourceKeys > 0 And dupRule = "abort" Then
        Err.Raise vbObjectError + 945, , "来源表存在 " & dupSourceKeys & " 个重复匹配键，当前规则为“中止并提示 (abort)”，未执行写入。"
    End If

    backupPath = DH_Backup(wb)

    On Error GoTo Rollback
    DH_Freeze state: frozen = True
    Set snapshot = Application.Workbooks.Add(xlWBATWorksheet): Set cache = snapshot.Worksheets(1)
    For cIdx = 0 To numCols - 1
        Set writeRange = tgtSheet.Range(tgtSheet.Cells(tgtFirstRow, writeCols(cIdx)), tgtSheet.Cells(tgtLastRow, writeCols(cIdx)))
        writeRange.Copy Destination:=cache.Range(writeRange.Address)
    Next cIdx
    touched = True
    srcPrefix = SafeSheetFormulaPrefix(srcSheet.Name)

    If isAutoAppend Then
        For cIdx = 0 To numCols - 1
            Set tgtCell = tgtSheet.Cells(tgtFirstRow, writeCols(cIdx))
            tgtCell.NumberFormat = "@"
            tgtCell.Value2 = autoHeaders(cIdx)
            details.Add Array(ColIdxToLetter(writeCols(cIdx)) & tgtFirstRow, "新增表头", "右侧自动新增列并写入表头", "", autoHeaders(cIdx), "")
        Next cIdx
    End If

    For r = tgtFirstRow + 1 To tgtLastRow
        rawKey = tgtSheet.Cells(r, tgtKeyCol).Value2
        tok = LookupKeyToken(rawKey, isBlank, isErr, kindTag, normNum)
        formulaSafe = True
        If writeMode = "formula" Then
            If keyFirstRow.Exists(tok) Then formulaSafe = FormulaMatchAgrees(rawKey, srcSheet.Range(srcSheet.Cells(srcFirstRow + 1, srcKeyCol), srcSheet.Cells(srcLastRow, srcKeyCol)), CLng(keyFirstRow(tok)))
        End If

        For cIdx = 0 To numCols - 1
            Set tgtCell = tgtSheet.Cells(r, writeCols(cIdx))
            If CBool(tgtCell.MergeCells) Then
                skippedProtected = skippedProtected + 1
                details.Add Array(tgtCell.Address(False, False), "跳过", "目标单元格位于合并区域保护跳过", DH_ValueToken(tgtCell.Value2), DH_ValueToken(tgtCell.Value2), "")
            ElseIf tgtCell.HasFormula Or Not IsEmpty(tgtCell.Value2) Then
                skippedProtected = skippedProtected + 1
                details.Add Array(tgtCell.Address(False, False), "保留", "目标单元格已有内容或公式保护不覆盖", DH_ValueToken(tgtCell.Value2), DH_ValueToken(tgtCell.Value2), "")
            ElseIf isBlank Then
                If cIdx = 0 Then skippedBlank = skippedBlank + 1
                details.Add Array(tgtCell.Address(False, False), "跳过", "目标行匹配键为空白", "", "", "")
            ElseIf isErr Then
                If cIdx = 0 Then skippedUnmatched = skippedUnmatched + 1
                details.Add Array(tgtCell.Address(False, False), "跳过", "目标行匹配键为公式错误值", "", "", "")
            ElseIf keyFirstRow.Exists(tok) Then
                If CLng(keyCounts(tok)) > 1 And dupRule = "skip" Then
                    If cIdx = 0 Then skippedDup = skippedDup + 1
                    details.Add Array(tgtCell.Address(False, False), "跳过", "来源表匹配键重复，按 skip 规则跳过", "", "", "")
                ElseIf Not formulaSafe Then
                    skippedFormula = skippedFormula + 1
                    details.Add Array(tgtCell.Address(False, False), "跳过", "公式匹配规则不一致，请改用静态值模式（大小写、通配符或键值两端空格）", "", "", "")
                Else
                    srcRow = CLng(keyFirstRow(tok))
                    Set srcCell = srcSheet.Cells(srcRow, retCols(cIdx))
                    oldToken = DH_ValueToken(tgtCell.Value2)
                    If IsError(srcCell.Value2) Then
                        skippedSrcErr = skippedSrcErr + 1
                        colHeaderLabel = ReadCleanHeader(srcSheet, srcFirstRow, retCols(cIdx), ColIdxToLetter(retCols(cIdx)))
                        details.Add Array(tgtCell.Address(False, False), "跳过", "来源表「" & srcSheet.Name & "」" & ColIdxToLetter(retCols(cIdx)) & srcRow & "（" & colHeaderLabel & "）为错误值(" & CStr(srcCell.Text) & ")保护跳过", oldToken, oldToken, "")
                    ElseIf writeMode = "formula" Then
                        formulaExpr = "=INDEX(" & srcPrefix & "$" & ColIdxToLetter(retCols(cIdx)) & "$" & (srcFirstRow + 1) & ":$" & ColIdxToLetter(retCols(cIdx)) & "$" & srcLastRow & _
                                      ",MATCH(" & ColIdxToLetter(tgtKeyCol) & r & "," & srcPrefix & "$" & ColIdxToLetter(srcKeyCol) & "$" & (srcFirstRow + 1) & ":$" & ColIdxToLetter(srcKeyCol) & "$" & srcLastRow & ",0))"
                        If tgtCell.NumberFormat = "@" Then tgtCell.NumberFormat = "General"
                        tgtCell.Formula = formulaExpr
                        If IsDate(srcCell.Value) Then
                            srcNumFmt = CStr(srcCell.NumberFormat)
                            If Len(srcNumFmt) > 0 And srcNumFmt <> "General" And srcNumFmt <> "@" Then
                                tgtCell.NumberFormat = srcNumFmt
                            Else
                                tgtCell.NumberFormat = "yyyy-mm-dd"
                            End If
                        End If
                        newToken = "F:" & formulaExpr
                        changedCells = changedCells + 1
                        details.Add Array(tgtCell.Address(False, False), "公式补全", "写入 INDEX/MATCH 精确匹配公式（来源行 " & srcRow & "）", oldToken, newToken, "")
                    Else
                        If VarType(srcCell.Value2) = vbString Then
                            tgtCell.NumberFormat = "@"
                            tgtCell.Value2 = CStr(srcCell.Value2)
                        ElseIf IsDate(srcCell.Value) Then
                            tgtCell.Value2 = srcCell.Value2
                            srcNumFmt = CStr(srcCell.NumberFormat)
                            If Len(srcNumFmt) > 0 And srcNumFmt <> "General" And srcNumFmt <> "@" Then
                                tgtCell.NumberFormat = srcNumFmt
                            Else
                                tgtCell.NumberFormat = "yyyy-mm-dd"
                            End If
                        Else
                            tgtCell.Value2 = srcCell.Value2
                        End If
                        newToken = DH_ValueToken(tgtCell.Value2)
                        changedCells = changedCells + 1
                        details.Add Array(tgtCell.Address(False, False), "跨表补全", "从「" & srcSheet.Name & "」" & ColIdxToLetter(retCols(cIdx)) & srcRow & " 补全静态值", oldToken, newToken, "")
                    End If
                End If
            Else
                If cIdx = 0 Then skippedUnmatched = skippedUnmatched + 1
                If Len(normNum) > 0 And normNumMap.Exists(normNum) Then
                    details.Add Array(tgtCell.Address(False, False), "跳过", "键值存在文本/数值类型差异（前导零保护未合并）", "", "", "")
                Else
                    details.Add Array(tgtCell.Address(False, False), "跳过", "来源表未找到匹配键值", "", "", "")
                End If
            End If
        Next cIdx
    Next r

    auditPath = DH_SaveAudit("跨表匹配补全", wb, details)
    DH_LastSummary = "跨表补全 " & changedCells & " 格（" & numCols & " 列" & IIf(isAutoAppend, "，右侧新增列", "") & "）；未命中跳过 " & skippedUnmatched & " 行，空键跳过 " & skippedBlank & " 行，重复键跳过 " & skippedDup & " 行，来源错误值跳过 " & skippedSrcErr & " 格，公式规则不一致跳过 " & skippedFormula & " 格，已有内容保护 " & skippedProtected & " 格；基线=" & backupPath & "；明细=" & auditPath
    DH_Log "跨表匹配补全", wb, tgtSheet.Name, targetAddress, changedCells, "成功；" & DH_LastSummary, backupPath
    If Not snapshot Is Nothing Then snapshot.Close False
    DH_ExecuteLookup = changedCells
    DH_RestoreFocus tgtSheet, originalSelection
    DH_Restore state
    Exit Function

Rollback:
    Dim errMsg As String, errNum As Long
    errNum = Err.Number: errMsg = Err.Description
    On Error Resume Next
    If touched And Not cache Is Nothing Then
        For cIdx = 0 To numCols - 1
            Set writeRange = tgtSheet.Range(tgtSheet.Cells(tgtFirstRow, writeCols(cIdx)), tgtSheet.Cells(tgtLastRow, writeCols(cIdx)))
            Err.Clear
            cache.Range(writeRange.Address).Copy Destination:=writeRange
            If Err.Number <> 0 Then rollbackFailed = True
        Next cIdx
    End If
    If Not snapshot Is Nothing Then snapshot.Close False
    On Error GoTo 0
    If rollbackFailed Then errMsg = errMsg & "；回滚不完整，请恢复当日基线：" & backupPath
    DH_RestoreFocus tgtSheet, originalSelection
    If frozen Then DH_Restore state
    DH_LogFailure "跨表匹配补全", wb, tgtSheet.Name, targetAddress, errMsg, backupPath
    Err.Raise IIf(errNum <> 0, errNum, vbObjectError + 946), "电子羊表格", "跨表匹配补全过程中发生异常，已回滚目标列：" & errMsg
End Function
