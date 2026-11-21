Attribute VB_Name = "DHAppend"
Option Explicit

' =============================================================================
' DataHelper 1.0 Modern Patch 3+: Multi-Sheet Vertical Append Engine
' Spec format:
'   Append:<sheetsCsv>:<colMode>:<summaryRowRule>:<sourceTagMode>:<targetSheetName>:<colMapCsv>
' =============================================================================

Private Function ColIdxToLetter(ByVal colIndex As Long) As String
    Dim n As Long
    n = colIndex
    Do While n > 0
        n = n - 1
        ColIdxToLetter = Chr$(65 + (n Mod 26)) & ColIdxToLetter
        n = n \ 26
    Loop
End Function

Private Function ItemsToArray(ByVal items As Collection) As Variant
    Dim arr() As Variant, i As Long
    If items.Count = 0 Then ItemsToArray = Array(): Exit Function
    ReDim arr(0 To items.Count - 1)
    For i = 1 To items.Count: arr(i - 1) = items(i): Next i
    ItemsToArray = arr
End Function

Public Function DH_NormalizeHeaderName(ByVal rawHeader As String) As String
    Dim s As String, i As Long, ch As Long, outStr As String
    s = Trim$(rawHeader)
    If Len(s) = 0 Then
        DH_NormalizeHeaderName = ""
        Exit Function
    End If
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1))
        If ch < 0 Then ch = ch + 65536
        If ch <> 32 And ch <> 12288 And ch <> 160 And ch <> 8203 And ch <> 9 And ch <> 10 And ch <> 13 Then
            outStr = outStr & Mid$(s, i, 1)
        End If
    Next i
    DH_NormalizeHeaderName = LCase$(outStr)
End Function

Public Function DH_ParseColMapDict(ByVal colMapCsv As String) As Object
    Dim d As Object, trimmed As String, pairs() As String, i As Long, p As String, eqPos As Long, leftKey As String, rightVal As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    trimmed = Trim$(colMapCsv)
    If Len(trimmed) = 0 Then
        Set DH_ParseColMapDict = d
        Exit Function
    End If
    pairs = Split(trimmed, ";")
    For i = LBound(pairs) To UBound(pairs)
        p = Trim$(pairs(i))
        If Len(p) > 0 Then
            eqPos = InStr(1, p, "=", vbBinaryCompare)
            If eqPos > 1 And eqPos < Len(p) Then
                leftKey = DH_NormalizeHeaderName(Left$(p, eqPos - 1))
                rightVal = Trim$(Mid$(p, eqPos + 1))
                If Len(leftKey) > 0 And Len(rightVal) > 0 Then
                    d(leftKey) = rightVal
                End If
            End If
        End If
    Next i
    Set DH_ParseColMapDict = d
End Function

Public Function DH_IsSuspectedSummaryRow(ByRef vals As Variant, ByRef fMask As Variant, ByRef fTexts As Variant, ByVal relRow As Long, ByVal colCount As Long, ByRef outReason As String, ByRef outLabelPreview As String) As Boolean
    Dim c As Long, nonBlankCount As Long, firstText As String, hasAggregateFormula As Boolean, aggFormulaSample As String, matchedKeyword As String
    Dim fStr As String, uF As String, v As Variant, t As String, normT As String
    DH_IsSuspectedSummaryRow = False
    outReason = ""
    outLabelPreview = ""

    For c = 1 To colCount
        If CBool(fMask(relRow, c)) Then
            fStr = Trim$(CStr(fTexts(relRow, c)))
            If Len(fStr) > 0 Then
                uF = UCase$(fStr)
                If InStr(1, uF, "SUM(", vbBinaryCompare) > 0 Or _
                   InStr(1, uF, "SUBTOTAL(", vbBinaryCompare) > 0 Or _
                   InStr(1, uF, "AVERAGE(", vbBinaryCompare) > 0 Or _
                   InStr(1, uF, "SUMIF(", vbBinaryCompare) > 0 Or _
                   InStr(1, uF, "SUMIFS(", vbBinaryCompare) > 0 Then
                    hasAggregateFormula = True
                    If Len(aggFormulaSample) = 0 Then aggFormulaSample = fStr
                End If
            End If
        End If

        v = vals(relRow, c)
        If Not IsError(v) And Not IsEmpty(v) And Not IsNull(v) Then
            t = Trim$(CStr(v))
            If Len(t) > 0 Then
                nonBlankCount = nonBlankCount + 1
                If Len(firstText) = 0 Then firstText = t
                If c <= 3 And Len(matchedKeyword) = 0 Then
                    normT = DH_NormalizeHeaderName(t)
                    If Right$(normT, 1) = ":" Or Right$(normT, 1) = "：" Then
                        normT = Left$(normT, Len(normT) - 1)
                    End If
                    Select Case normT
                        Case "合计", "总计", "小计", "本月合计", "本页小计", "本季合计", "本年合计", "累计数", "累计", "汇总", _
                             "total", "subtotal", "grandtotal", "summary"
                            matchedKeyword = t
                    End Select
                End If
            End If
        End If
    Next c

    If Len(matchedKeyword) > 0 Then
        DH_IsSuspectedSummaryRow = True
        outLabelPreview = matchedKeyword
        If hasAggregateFormula Then
            outReason = "行首含汇总词「" & matchedKeyword & "」且含公式 " & aggFormulaSample
        Else
            outReason = "行首标识为汇总词「" & matchedKeyword & "」"
        End If
        Exit Function
    End If

    If hasAggregateFormula And nonBlankCount <= 3 Then
        DH_IsSuspectedSummaryRow = True
        outLabelPreview = IIf(Len(firstText) > 0, firstText, "(空标识汇总行)")
        outReason = "含汇总公式 " & aggFormulaSample & " 且仅含汇总格"
    End If
End Function

Public Sub DH_ParseAppendSpec(ByVal spec As String, _
    ByRef sheetsCsv As String, ByRef colMode As String, ByRef summaryRowRule As String, _
    ByRef sourceTagMode As String, ByRef targetSheetName As String, ByRef colMapCsv As String)
    Dim parts As Variant, offset As Long
    parts = Split(spec, ":")
    If UBound(parts) >= 5 And StrComp(CStr(parts(0)), "Append", vbTextCompare) = 0 Then
        offset = 1
    ElseIf UBound(parts) >= 4 Then
        offset = 0
    Else
        Err.Raise vbObjectError + 950, , "多表纵向追加参数格式无效。"
    End If
    sheetsCsv = Trim$(CStr(parts(offset)))
    colMode = LCase$(Trim$(CStr(parts(offset + 1))))
    summaryRowRule = LCase$(Trim$(CStr(parts(offset + 2))))
    sourceTagMode = LCase$(Trim$(CStr(parts(offset + 3))))
    targetSheetName = Trim$(CStr(parts(offset + 4)))
    If UBound(parts) >= offset + 5 Then
        colMapCsv = Trim$(CStr(parts(offset + 5)))
    Else
        colMapCsv = ""
    End If
    If Len(sheetsCsv) = 0 Then Err.Raise vbObjectError + 950, , "请至少勾选 2 张来源工作表。"
    If colMode <> "union" And colMode <> "common" Then Err.Raise vbObjectError + 950, , "列对齐模式无效，仅支持 union 或 common。"
    If summaryRowRule <> "exclude" And summaryRowRule <> "keep" Then Err.Raise vbObjectError + 950, , "小计/合计行处理规则无效，仅支持 exclude 或 keep。"
    If sourceTagMode <> "sheet" And sourceTagMode <> "sheet_row" And sourceTagMode <> "none" Then Err.Raise vbObjectError + 950, , "来源标记模式无效，仅支持 sheet / sheet_row / none。"
    If Len(targetSheetName) = 0 Then targetSheetName = "多表追加汇总"
    If Len(targetSheetName) > 31 Then Err.Raise vbObjectError + 950, , "汇总工作表名称不能超过 31 个字符。"
End Sub

Public Function DH_AppendPreview(ByVal wb As Workbook, ByVal spec As String) As Variant
    Dim sheetsCsv As String, colMode As String, summaryRowRule As String
    Dim sourceTagMode As String, targetSheetName As String, colMapCsv As String
    Dim sheetNames As Variant, sheetCount As Long, colMapDict As Object
    Dim canonicalOrder As New Collection, canonicalDisplay As Object, canonicalSheetHits As Object
    Dim sheetBreakdown As New Collection, summaryRowSamples As New Collection, anomalies As New Collection
    Dim totalSourceRows As Long, netAppendRows As Long, summaryRowCount As Long, sourceErrorCount As Long
    Dim canExecute As Boolean, blockReason As String, seenSheets As Object
    Dim sIdx As Long, sName As String, ws As Worksheet, contentRng As Range
    Dim rCount As Long, cCount As Long, vals As Variant, fMask As Variant, fTexts As Variant
    Dim sheetColsSeen As Object, c As Long, validHeaderCols As Long, rawH As String, normH As String, mappedDisplay As String
    Dim r As Long, sheetDataRows As Long, sheetSummaryRows As Long, sheetNetRows As Long, rowHasContent As Boolean
    Dim sumReason As String, sumPreview As String
    Dim finalHeaders As New Collection, unionOnlyCount As Long, kIdx As Long, kNorm As String, hits As Long
    Dim resolvedTargetSheet As String

    DH_RequireWorkbook wb
    DH_ParseAppendSpec spec, sheetsCsv, colMode, summaryRowRule, sourceTagMode, targetSheetName, colMapCsv

    sheetNames = Split(sheetsCsv, ",")
    sheetCount = UBound(sheetNames) - LBound(sheetNames) + 1
    If sheetCount < 2 Or sheetCount > 32 Then
        Err.Raise vbObjectError + 951, , "请勾选 2 至 32 张需要纵向追加的工作表。"
    End If

    Set colMapDict = DH_ParseColMapDict(colMapCsv)
    Set canonicalDisplay = CreateObject("Scripting.Dictionary")
    Set canonicalSheetHits = CreateObject("Scripting.Dictionary")
    Set seenSheets = CreateObject("Scripting.Dictionary")
    seenSheets.CompareMode = 1
    canExecute = True
    blockReason = ""

    For sIdx = LBound(sheetNames) To UBound(sheetNames)
        sName = Trim$(CStr(sheetNames(sIdx)))
        If seenSheets.Exists(sName) Then
            canExecute = False
            blockReason = "勾选的工作表列表中存在重复表名「" & sName & "」。"
            Exit For
        End If
        seenSheets(sName) = True

        Set ws = Nothing
        On Error Resume Next
        Set ws = wb.Worksheets(sName)
        On Error GoTo 0
        If ws Is Nothing Then
            canExecute = False
            blockReason = "未找到工作表「" & sName & "」，请刷新列表后重试。"
            Exit For
        End If
        If ws.Visible <> xlSheetVisible Or DH_IsReportSheet(ws) Or DH_IsPreviousReport(ws) Then
            canExecute = False
            blockReason = "工作表「" & sName & "」不是可见的普通业务数据表。"
            Exit For
        End If

        Set contentRng = DH_ContentRange(ws)
        If contentRng Is Nothing Then
            anomalies.Add Array(sName & "!A1", "空工作表", "该工作表无有效数据，追加时将自动跳过")
            sheetBreakdown.Add Array(sName, 0, 0, 0, 0)
        Else
            rCount = contentRng.Rows.Count
            cCount = contentRng.Columns.Count
            If rCount < 2 Then
                anomalies.Add Array(sName & "!1", "仅有表头无数据行", "该表只有第 1 行表头，无可追加明细行")
                sheetBreakdown.Add Array(sName, 0, 0, 0, cCount)
            Else
                vals = DH_Matrix(contentRng)
                fMask = DH_FormulaMask(contentRng)
                fTexts = DH_Matrix(contentRng, True)
                Set sheetColsSeen = CreateObject("Scripting.Dictionary")
                validHeaderCols = 0
                For c = 1 To cCount
                    If IsError(vals(1, c)) Or IsEmpty(vals(1, c)) Or IsNull(vals(1, c)) Then
                        rawH = ""
                    Else
                        rawH = Trim$(CStr(vals(1, c)))
                    End If
                    If Len(rawH) = 0 Then
                        anomalies.Add Array(sName & "!" & ColIdxToLetter(contentRng.Column + c - 1) & contentRng.Row, "表头为空", "首行缺少列名，已忽略该列")
                    Else
                        normH = DH_NormalizeHeaderName(rawH)
                        If colMapDict.Exists(normH) Then
                            mappedDisplay = CStr(colMapDict(normH))
                            normH = DH_NormalizeHeaderName(mappedDisplay)
                        Else
                            mappedDisplay = rawH
                        End If
                        If sheetColsSeen.Exists(normH) Then
                            canExecute = False
                            blockReason = "工作表「" & sName & "」表头行存在重复列名「" & mappedDisplay & "」（或映射后重名），请先修改表头确保唯一。"
                            anomalies.Add Array(sName & "!" & ColIdxToLetter(contentRng.Column + c - 1) & contentRng.Row, "表内列名重复(阻断)", mappedDisplay)
                        Else
                            sheetColsSeen(normH) = c
                            validHeaderCols = validHeaderCols + 1
                            If Not canonicalDisplay.Exists(normH) Then
                                canonicalDisplay(normH) = mappedDisplay
                                canonicalOrder.Add normH
                                canonicalSheetHits(normH) = 1
                            Else
                                canonicalSheetHits(normH) = CLng(canonicalSheetHits(normH)) + 1
                            End If
                        End If
                    End If
                Next c

                sheetDataRows = 0: sheetSummaryRows = 0: sheetNetRows = 0
                For r = 2 To rCount
                    rowHasContent = False
                    For c = 1 To cCount
                        If IsError(vals(r, c)) Then
                            rowHasContent = True
                            sourceErrorCount = sourceErrorCount + 1
                            If anomalies.Count < 25 Then
                                anomalies.Add Array(sName & "!" & ColIdxToLetter(contentRng.Column + c - 1) & (contentRng.Row + r - 1), "来源含错误值(将跳过该格)", CStr(ws.Cells(contentRng.Row + r - 1, contentRng.Column + c - 1).Text))
                            End If
                        ElseIf CBool(fMask(r, c)) Then
                            rowHasContent = True
                        ElseIf Not IsEmpty(vals(r, c)) And Not IsNull(vals(r, c)) Then
                            If Len(Trim$(CStr(vals(r, c)))) > 0 Then rowHasContent = True
                        End If
                    Next c

                    If rowHasContent Then
                        sheetDataRows = sheetDataRows + 1
                        totalSourceRows = totalSourceRows + 1
                        If DH_IsSuspectedSummaryRow(vals, fMask, fTexts, r, cCount, sumReason, sumPreview) Then
                            sheetSummaryRows = sheetSummaryRows + 1
                            summaryRowCount = summaryRowCount + 1
                            If summaryRowSamples.Count < 15 Then
                                summaryRowSamples.Add Array(sName & "!第" & CStr(contentRng.Row + r - 1) & "行", sumPreview, sumReason)
                            End If
                            If summaryRowRule = "keep" Then
                                sheetNetRows = sheetNetRows + 1
                                netAppendRows = netAppendRows + 1
                            End If
                        Else
                            sheetNetRows = sheetNetRows + 1
                            netAppendRows = netAppendRows + 1
                        End If
                    End If
                Next r
                sheetBreakdown.Add Array(sName, sheetDataRows, sheetNetRows, sheetSummaryRows, validHeaderCols)
            End If
        End If
    Next sIdx

    For kIdx = 1 To canonicalOrder.Count
        kNorm = CStr(canonicalOrder(kIdx))
        hits = CLng(canonicalSheetHits(kNorm))
        If hits < sheetCount Then
            unionOnlyCount = unionOnlyCount + 1
            If colMode = "union" Then
                finalHeaders.Add Array(CStr(canonicalDisplay(kNorm)), hits, sheetCount)
                If anomalies.Count < 25 Then
                    anomalies.Add Array("字段「" & CStr(canonicalDisplay(kNorm)) & "」", "仅部分工作表包含(" & CStr(hits) & "/" & CStr(sheetCount) & "张表)", "并集模式下缺失表对应格留空")
                End If
            Else
                If anomalies.Count < 25 Then
                    anomalies.Add Array("字段「" & CStr(canonicalDisplay(kNorm)) & "」", "非共有列已排除(" & CStr(hits) & "/" & CStr(sheetCount) & "张表)", "仅共有列模式下不输出该列")
                End If
            End If
        Else
            finalHeaders.Add Array(CStr(canonicalDisplay(kNorm)), hits, sheetCount)
        End If
    Next kIdx

    If canExecute And finalHeaders.Count = 0 Then
        canExecute = False
        If colMode = "common" Then
            blockReason = "所选工作表之间没有任何同名的共有列，请切换为「保留全部列（并集）」或在更多设置中配置列名映射。"
        Else
            blockReason = "所选工作表未识别到任何有效表头列。"
        End If
    End If
    If canExecute And netAppendRows = 0 Then
        canExecute = False
        blockReason = "所选工作表过滤后没有可追加的明细数据行。"
    End If
    If canExecute And netAppendRows * (finalHeaders.Count + 1) > 100000 Then
        canExecute = False
        blockReason = "预计写入单元格数超过单次 10 万格安全上限，请分批勾选工作表追加。"
    End If

    resolvedTargetSheet = DH_UniqueSheetName(wb, targetSheetName)

    DH_AppendPreview = Array( _
        sheetCount, _
        totalSourceRows, _
        netAppendRows, _
        summaryRowCount, _
        finalHeaders.Count, _
        unionOnlyCount, _
        sourceErrorCount, _
        resolvedTargetSheet, _
        colMode, _
        summaryRowRule, _
        sourceTagMode, _
        canExecute, _
        blockReason, _
        ItemsToArray(finalHeaders), _
        ItemsToArray(summaryRowSamples), _
        ItemsToArray(sheetBreakdown), _
        ItemsToArray(anomalies) _
    )
End Function

Public Function DH_ExecuteAppend(ByVal wb As Workbook, ByVal spec As String) As Long
    Dim sheetsCsv As String, colMode As String, summaryRowRule As String
    Dim sourceTagMode As String, targetSheetName As String, colMapCsv As String
    Dim preview As Variant, canExec As Boolean, blockMsg As String
    Dim backupPath As String, auditPath As String
    Dim sheetNames As Variant, sheetCount As Long, colMapDict As Object
    Dim canonicalOrder As New Collection, canonicalDisplay As Object, canonicalHits As Object
    Dim sIdx As Long, sName As String, wsScan As Worksheet, rngScan As Range, vScan As Variant, cScan As Long
    Dim rawH As String, normH As String, dispH As String
    Dim activeNorms As New Collection, activeDisplays As New Collection, k As Long, kn As String
    Dim hasSourceCol As Boolean, colOffset As Long, totalOutCols As Long, resolvedName As String
    Dim state As DH_AppState, frozen As Boolean, outWs As Worksheet, origActiveSheet As Worksheet, origSelection As Range, details As New Collection
    Dim appendedRows As Long, skippedSummaryRows As Long, skippedErrorCells As Long
    Dim srcColTitle As String, writeRow As Long
    Dim srcWs As Worksheet, srcRng As Range, vals As Variant, fMask As Variant, fTexts As Variant
    Dim srcColForOut() As Long, c As Long, normToSrcCol As Object, rh As String, nh As String
    Dim r As Long, rowHasContent As Boolean, sumReason As String, sumPreview As String, isSum As Boolean
    Dim scIdx As Long, cellVal As Variant, srcCellObj As Range, dstCellObj As Range, srcNumFmt As String

    DH_RequireWorkbook wb, True
    DH_ParseAppendSpec spec, sheetsCsv, colMode, summaryRowRule, sourceTagMode, targetSheetName, colMapCsv

    preview = DH_AppendPreview(wb, spec)
    canExec = CBool(preview(11))
    blockMsg = CStr(preview(12))
    If Not canExec Then
        Err.Raise vbObjectError + 952, , blockMsg
    End If

    If TypeOf Application.ActiveSheet Is Worksheet Then
        Set origActiveSheet = Application.ActiveSheet
        Set origSelection = DH_OriginalSelection(origActiveSheet)
    End If
    backupPath = DH_Backup(wb)

    sheetNames = Split(sheetsCsv, ",")
    sheetCount = UBound(sheetNames) - LBound(sheetNames) + 1
    Set colMapDict = DH_ParseColMapDict(colMapCsv)
    Set canonicalDisplay = CreateObject("Scripting.Dictionary")
    Set canonicalHits = CreateObject("Scripting.Dictionary")

    For sIdx = LBound(sheetNames) To UBound(sheetNames)
        Set wsScan = wb.Worksheets(Trim$(CStr(sheetNames(sIdx))))
        Set rngScan = DH_ContentRange(wsScan)
        If Not rngScan Is Nothing Then
            If rngScan.Rows.Count >= 2 Then
                vScan = DH_Matrix(rngScan)
                For cScan = 1 To rngScan.Columns.Count
                    If Not IsError(vScan(1, cScan)) And Not IsEmpty(vScan(1, cScan)) And Not IsNull(vScan(1, cScan)) Then
                        rawH = Trim$(CStr(vScan(1, cScan)))
                        If Len(rawH) > 0 Then
                            normH = DH_NormalizeHeaderName(rawH)
                            If colMapDict.Exists(normH) Then
                                dispH = CStr(colMapDict(normH))
                                normH = DH_NormalizeHeaderName(dispH)
                            Else
                                dispH = rawH
                            End If
                            If Not canonicalDisplay.Exists(normH) Then
                                canonicalDisplay(normH) = dispH
                                canonicalOrder.Add normH
                                canonicalHits(normH) = 1
                            Else
                                canonicalHits(normH) = CLng(canonicalHits(normH)) + 1
                            End If
                        End If
                    End If
                Next cScan
            End If
        End If
    Next sIdx

    For k = 1 To canonicalOrder.Count
        kn = CStr(canonicalOrder(k))
        If colMode = "union" Or CLng(canonicalHits(kn)) = sheetCount Then
            activeNorms.Add kn
            activeDisplays.Add CStr(canonicalDisplay(kn))
        End If
    Next k

    hasSourceCol = (sourceTagMode <> "none")
    colOffset = IIf(hasSourceCol, 1, 0)
    totalOutCols = activeNorms.Count + colOffset
    resolvedName = DH_UniqueSheetName(wb, targetSheetName)

    On Error GoTo Rollback
    DH_Freeze state: frozen = True

    Set outWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    outWs.Name = resolvedName

    If hasSourceCol Then
        srcColTitle = IIf(sourceTagMode = "sheet_row", "来源工作表与行号", "来源工作表")
        If canonicalDisplay.Exists(DH_NormalizeHeaderName(srcColTitle)) Then
            srcColTitle = srcColTitle & "_DH"
        End If
        outWs.Cells(1, 1).NumberFormat = "@"
        outWs.Cells(1, 1).Value2 = srcColTitle
    End If
    For k = 1 To activeDisplays.Count
        outWs.Cells(1, k + colOffset).NumberFormat = "@"
        outWs.Cells(1, k + colOffset).Value2 = CStr(activeDisplays(k))
    Next k
    outWs.Rows(1).Font.Bold = True

    writeRow = 1
    For sIdx = LBound(sheetNames) To UBound(sheetNames)
        sName = Trim$(CStr(sheetNames(sIdx)))
        Set srcWs = wb.Worksheets(sName)
        Set srcRng = DH_ContentRange(srcWs)
        If Not srcRng Is Nothing Then
            If srcRng.Rows.Count >= 2 Then
                vals = DH_Matrix(srcRng)
                fMask = DH_FormulaMask(srcRng)
                fTexts = DH_Matrix(srcRng, True)
                ReDim srcColForOut(1 To activeNorms.Count)
                Set normToSrcCol = CreateObject("Scripting.Dictionary")
                For c = 1 To srcRng.Columns.Count
                    If Not IsError(vals(1, c)) And Not IsEmpty(vals(1, c)) And Not IsNull(vals(1, c)) Then
                        rh = Trim$(CStr(vals(1, c)))
                        If Len(rh) > 0 Then
                            nh = DH_NormalizeHeaderName(rh)
                            If colMapDict.Exists(nh) Then
                                nh = DH_NormalizeHeaderName(CStr(colMapDict(nh)))
                            End If
                            If Not normToSrcCol.Exists(nh) Then normToSrcCol(nh) = c
                        End If
                    End If
                Next c
                For k = 1 To activeNorms.Count
                    If normToSrcCol.Exists(CStr(activeNorms(k))) Then
                        srcColForOut(k) = CLng(normToSrcCol(CStr(activeNorms(k))))
                    Else
                        srcColForOut(k) = 0
                    End If
                Next k

                For r = 2 To srcRng.Rows.Count
                    rowHasContent = False
                    For c = 1 To srcRng.Columns.Count
                        If CBool(fMask(r, c)) Or IsError(vals(r, c)) Then
                            rowHasContent = True
                            Exit For
                        ElseIf Not IsEmpty(vals(r, c)) And Not IsNull(vals(r, c)) Then
                            If Len(Trim$(CStr(vals(r, c)))) > 0 Then
                                rowHasContent = True
                                Exit For
                            End If
                        End If
                    Next c

                    If rowHasContent Then
                        isSum = DH_IsSuspectedSummaryRow(vals, fMask, fTexts, r, srcRng.Columns.Count, sumReason, sumPreview)
                        If isSum And summaryRowRule = "exclude" Then
                            skippedSummaryRows = skippedSummaryRows + 1
                            details.Add Array(sName & "!第" & CStr(srcRng.Row + r - 1) & "行", "跳过", "排除小计/合计行（" & sumReason & "）", sumPreview, "", "")
                        Else
                            writeRow = writeRow + 1
                            appendedRows = appendedRows + 1
                            If hasSourceCol Then
                                outWs.Cells(writeRow, 1).NumberFormat = "@"
                                If sourceTagMode = "sheet_row" Then
                                    outWs.Cells(writeRow, 1).Value2 = sName & "!" & CStr(srcRng.Row + r - 1)
                                Else
                                    outWs.Cells(writeRow, 1).Value2 = sName
                                End If
                            End If
                            For k = 1 To activeNorms.Count
                                scIdx = srcColForOut(k)
                                If scIdx > 0 Then
                                    cellVal = vals(r, scIdx)
                                    Set srcCellObj = srcRng.Cells(r, scIdx)
                                    Set dstCellObj = outWs.Cells(writeRow, k + colOffset)
                                    If IsError(cellVal) Then
                                        skippedErrorCells = skippedErrorCells + 1
                                        details.Add Array(sName & "!" & srcCellObj.Address(False, False), "跳过", "来源格含错误值(" & CStr(srcCellObj.Text) & ")保护跳过", DH_ValueToken(cellVal), "", "")
                                    ElseIf Not IsEmpty(cellVal) And Not IsNull(cellVal) Then
                                        If VarType(cellVal) = vbString Then
                                            dstCellObj.NumberFormat = "@"
                                            dstCellObj.Value2 = CStr(cellVal)
                                        ElseIf IsDate(srcCellObj.Value) Then
                                            dstCellObj.Value2 = cellVal
                                            srcNumFmt = CStr(srcCellObj.NumberFormat)
                                            If Len(srcNumFmt) > 0 And srcNumFmt <> "General" And srcNumFmt <> "@" Then
                                                dstCellObj.NumberFormat = srcNumFmt
                                            Else
                                                dstCellObj.NumberFormat = "yyyy-mm-dd"
                                            End If
                                        Else
                                            dstCellObj.Value2 = cellVal
                                        End If
                                    End If
                                End If
                            Next k
                            details.Add Array(outWs.Name & "!第" & CStr(writeRow) & "行", "追加明细", "从「" & sName & "」第 " & CStr(srcRng.Row + r - 1) & " 行追加", "", sName, "")
                        End If
                    End If
                Next r
            End If
        End If
    Next sIdx

    auditPath = DH_SaveAudit("多表纵向追加", wb, details)
    DH_LastSummary = "从 " & CStr(sheetCount) & " 张工作表纵向追加 " & CStr(appendedRows) & " 行至新工作表「" & outWs.Name & "」（共 " & CStr(activeNorms.Count) & " 列业务字段" & IIf(hasSourceCol, " + 来源工作表列", "") & "）；排除小计/合计 " & CStr(skippedSummaryRows) & " 行，跳过来源错误值 " & CStr(skippedErrorCells) & " 格；基线=" & backupPath & "；明细=" & auditPath
    DH_Log "多表纵向追加", wb, outWs.Name, "A1:" & ColIdxToLetter(totalOutCols) & CStr(writeRow), appendedRows, "成功；" & DH_LastSummary, backupPath
    DH_RestoreFocus origActiveSheet, origSelection
    DH_Restore state
    DH_ExecuteAppend = appendedRows
    Exit Function

Rollback:
    Dim errMsg As String, errNum As Long
    errNum = Err.Number: errMsg = Err.Description
    On Error Resume Next
    If Not outWs Is Nothing Then
        Application.DisplayAlerts = False
        outWs.Delete
        Application.DisplayAlerts = True
    End If
    DH_RestoreFocus origActiveSheet, origSelection
    If frozen Then DH_Restore state
    On Error GoTo 0
    DH_LogFailure "多表纵向追加", wb, targetSheetName, sheetsCsv, errMsg, backupPath
    Err.Raise IIf(errNum <> 0, errNum, vbObjectError + 954), "电子羊表格", "多表纵向追加过程中发生异常，已清理临时汇总表并回滚：" & errMsg
End Function
