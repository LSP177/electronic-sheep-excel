Attribute VB_Name = "DHUnpivot"
Option Explicit

' =============================================================================
' DataHelper 1.0 Modern Patch 4: Wide-to-Long Table Unpivot Engine (宽表转长表)
' Spec format:
'   Unpivot:<idColsCsv>:<valColsCsv>:<attrHeader>:<valHeader>:<emptyPolicy>:<summaryRule>:<targetSheetName>
'   - idColsCsv: fixed dimension column letters (1..16 cols), e.g. "A,B"
'   - valColsCsv: wide value column letters to unpivot (1..64 cols), e.g. "C,D,E,F"
'   - attrHeader: header name for the unpivoted attribute column (default "属性类别")
'   - valHeader: header name for the unpivoted value column (default "数值")
'   - emptyPolicy: "skip_blank" (default) | "skip_blank_zero" | "keep_all"
'   - summaryRule: "exclude" (default, skip summary rows & summary columns) | "keep"
'   - targetSheetName: desired output worksheet name (default "宽表转长表明细")
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

Private Function ParseColList(ByVal csv As String, ByVal minCount As Long, ByVal maxCount As Long, ByVal label As String) As Long()
    Dim raw As String, parts As Variant, i As Long, idx As Long, seen As Object, out() As Long
    raw = Trim$(csv)
    If Len(raw) = 0 Then Err.Raise vbObjectError + 970, , label & "不能为空。"
    parts = Split(raw, ",")
    If UBound(parts) + 1 < minCount Or UBound(parts) + 1 > maxCount Then
        Err.Raise vbObjectError + 970, , label & "数量必须在 " & minCount & " 至 " & maxCount & " 列之间。"
    End If
    Set seen = CreateObject("Scripting.Dictionary")
    ReDim out(0 To UBound(parts))
    For i = 0 To UBound(parts)
        idx = DH_ColLetterToIndex(CStr(parts(i)))
        If seen.Exists(CStr(idx)) Then Err.Raise vbObjectError + 970, , label & "中包含重复列：" & ColIdxToLetter(idx)
        seen.Add CStr(idx), True
        out(i) = idx
    Next i
    ParseColList = out
End Function

Public Function DH_IsSummaryHeaderTitle(ByVal rawTitle As String) As Boolean
    Dim norm As String
    norm = DH_NormalizeHeaderName(rawTitle)
    If Right$(norm, 1) = ":" Or Right$(norm, 1) = "：" Then
        norm = Left$(norm, Len(norm) - 1)
    End If
    Select Case norm
        Case "合计", "总计", "小计", "行合计", "列合计", "本月合计", "本季合计", "全年合计", "本年合计", "汇总", "累计数", "累计", _
             "total", "subtotal", "grandtotal", "rowtotal", "summary"
            DH_IsSummaryHeaderTitle = True
        Case Else
            DH_IsSummaryHeaderTitle = False
    End Select
End Function

Public Sub DH_ParseUnpivotSpec(ByVal spec As String, _
    ByRef idColsCsv As String, ByRef valColsCsv As String, _
    ByRef attrHeader As String, ByRef valHeader As String, _
    ByRef emptyPolicy As String, ByRef summaryRule As String, _
    ByRef targetSheetName As String)
    Dim parts As Variant, offset As Long
    parts = Split(spec, ":")
    If UBound(parts) = 7 And StrComp(CStr(parts(0)), "Unpivot", vbTextCompare) = 0 Then
        offset = 1
    ElseIf UBound(parts) = 6 Then
        offset = 0
    Else
        Err.Raise vbObjectError + 971, , "宽表转长表参数格式无效。"
    End If
    idColsCsv = UCase$(Trim$(CStr(parts(offset))))
    valColsCsv = UCase$(Trim$(CStr(parts(offset + 1))))
    attrHeader = Trim$(CStr(parts(offset + 2)))
    valHeader = Trim$(CStr(parts(offset + 3)))
    emptyPolicy = LCase$(Trim$(CStr(parts(offset + 4))))
    summaryRule = LCase$(Trim$(CStr(parts(offset + 5))))
    targetSheetName = Trim$(CStr(parts(offset + 6)))

    If Len(attrHeader) = 0 Then attrHeader = "属性类别"
    If Len(valHeader) = 0 Then valHeader = "数值"
    If Len(targetSheetName) = 0 Then targetSheetName = "宽表转长表明细"
    If Len(attrHeader) > 64 Or Len(valHeader) > 64 Then Err.Raise vbObjectError + 971, , "属性列名或数值列名过长。"
    If StrComp(DH_NormalizeHeaderName(attrHeader), DH_NormalizeHeaderName(valHeader), vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 971, , "展开后的「属性列名」与「数值列名」不能同名。"
    End If
    If emptyPolicy <> "skip_blank" And emptyPolicy <> "skip_blank_zero" And emptyPolicy <> "keep_all" Then
        Err.Raise vbObjectError + 971, , "空值处理规则无效，仅支持 skip_blank / skip_blank_zero / keep_all。"
    End If
    If summaryRule <> "exclude" And summaryRule <> "keep" Then
        Err.Raise vbObjectError + 971, , "小计/合计处理规则无效，仅支持 exclude 或 keep。"
    End If
    If Len(targetSheetName) > 31 Then Err.Raise vbObjectError + 971, , "新建工作表名称不能超过 31 个字符。"
End Sub

Public Function DH_UnpivotPreview(ByVal srcSheet As Worksheet, ByVal spec As String) As Variant
    Dim wb As Workbook, bounds As Range
    Dim idColsCsv As String, valColsCsv As String, attrHeader As String, valHeader As String
    Dim emptyPolicy As String, summaryRule As String, targetSheetName As String
    Dim idCols() As Long, valCols() As Long, idSet As Object
    Dim i As Long, j As Long, r As Long, c As Long, firstRow As Long, lastRow As Long, firstCol As Long, lastCol As Long
    Dim vals As Variant, fMask As Variant, fTexts As Variant
    Dim totalSourceRows As Long, netOutputRows As Long, skippedBlankOrZero As Long
    Dim summaryRowCount As Long, summaryColCount As Long, sourceErrorCount As Long
    Dim canExecute As Boolean, blockReason As String, resolvedTargetSheet As String
    Dim samples As New Collection, summarySamples As New Collection, anomalies As New Collection
    Dim valColIsSummary() As Boolean, valColHeaders() As String, idColHeaders() As String
    Dim relRow As Long, relCol As Long, sumReason As String, sumPreview As String, isSumRow As Boolean

    Set wb = srcSheet.Parent
    DH_RequireWorkbook wb
    DH_ParseUnpivotSpec spec, idColsCsv, valColsCsv, attrHeader, valHeader, emptyPolicy, summaryRule, targetSheetName

    idCols = ParseColList(idColsCsv, 1, 16, "左侧固定列")
    valCols = ParseColList(valColsCsv, 1, 64, "要展开的数值列")

    Set idSet = CreateObject("Scripting.Dictionary")
    For i = 0 To UBound(idCols)
        idSet.Add CStr(idCols(i)), True
    Next i
    For j = 0 To UBound(valCols)
        If idSet.Exists(CStr(valCols(j))) Then
            Err.Raise vbObjectError + 972, , "「左侧固定列」与「要展开的列」不能包含同一列（" & ColIdxToLetter(valCols(j)) & " 列）。"
        End If
    Next j

    Set bounds = DH_ContentRange(srcSheet)
    If bounds Is Nothing Then Err.Raise vbObjectError + 973, , "当前工作表为空，没有可转换的数据。"
    If bounds.Rows.Count < 2 Then Err.Raise vbObjectError + 973, , "当前工作表只有表头行，没有可转换的数据行。"

    firstRow = bounds.Row
    lastRow = bounds.Row + bounds.Rows.Count - 1
    firstCol = bounds.Column
    lastCol = bounds.Column + bounds.Columns.Count - 1

    vals = DH_Matrix(bounds)
    fMask = DH_FormulaMask(bounds)
    fTexts = DH_Matrix(bounds, True)

    canExecute = True
    blockReason = ""

    ReDim idColHeaders(0 To UBound(idCols))
    For i = 0 To UBound(idCols)
        If idCols(i) < firstCol Or idCols(i) > lastCol Then
            canExecute = False
            blockReason = "选择的固定列 " & ColIdxToLetter(idCols(i)) & " 超出当前表有效数据范围。"
            idColHeaders(i) = ColIdxToLetter(idCols(i))
        Else
            relCol = idCols(i) - firstCol + 1
            If IsError(vals(1, relCol)) Or IsEmpty(vals(1, relCol)) Or Len(Trim$(CStr(vals(1, relCol)))) = 0 Then
                idColHeaders(i) = "维度_" & ColIdxToLetter(idCols(i))
            Else
                idColHeaders(i) = Trim$(CStr(vals(1, relCol)))
            End If
        End If
    Next i

    ReDim valColHeaders(0 To UBound(valCols))
    ReDim valColIsSummary(0 To UBound(valCols))
    Dim activeValColCount As Long
    For j = 0 To UBound(valCols)
        If valCols(j) < firstCol Or valCols(j) > lastCol Then
            canExecute = False
            blockReason = "选择的展开列 " & ColIdxToLetter(valCols(j)) & " 超出当前表有效数据范围。"
            valColHeaders(j) = ColIdxToLetter(valCols(j))
        Else
            relCol = valCols(j) - firstCol + 1
            If IsError(vals(1, relCol)) Or IsEmpty(vals(1, relCol)) Or Len(Trim$(CStr(vals(1, relCol)))) = 0 Then
                canExecute = False
                blockReason = "要展开的 " & ColIdxToLetter(valCols(j)) & " 列首行缺少表头名称，无法作为属性类别值，请先补齐首行表头。"
                valColHeaders(j) = "(空表头_" & ColIdxToLetter(valCols(j)) & ")"
                anomalies.Add Array(ColIdxToLetter(valCols(j)) & firstRow, "展开列表头为空(阻断)", "展开列必须有表头名称作为转换后的属性值")
            Else
                valColHeaders(j) = Trim$(CStr(vals(1, relCol)))
                If DH_IsSummaryHeaderTitle(valColHeaders(j)) Then
                    valColIsSummary(j) = True
                    summaryColCount = summaryColCount + 1
                    summarySamples.Add Array(ColIdxToLetter(valCols(j)) & "列（表头）", valColHeaders(j), "列标题为汇总词「" & valColHeaders(j) & "」")
                End If
                If Not (valColIsSummary(j) And summaryRule = "exclude") Then
                    activeValColCount = activeValColCount + 1
                End If
            End If
        End If
    Next j

    If canExecute And activeValColCount = 0 Then
        canExecute = False
        blockReason = "剔除汇总列后没有剩余可展开的数值列，请核对勾选的展开列。"
    End If

    If canExecute Then
        For relRow = 2 To bounds.Rows.Count
            r = firstRow + relRow - 1
            Dim rowHasContent As Boolean
            rowHasContent = False
            For c = 1 To bounds.Columns.Count
                If CBool(fMask(relRow, c)) Or IsError(vals(relRow, c)) Then
                    rowHasContent = True: Exit For
                ElseIf Not IsEmpty(vals(relRow, c)) And Not IsNull(vals(relRow, c)) Then
                    If Len(Trim$(CStr(vals(relRow, c)))) > 0 Then rowHasContent = True: Exit For
                End If
            Next c

            If rowHasContent Then
                totalSourceRows = totalSourceRows + 1
                isSumRow = DH_IsSuspectedSummaryRow(vals, fMask, fTexts, relRow, bounds.Columns.Count, sumReason, sumPreview)
                If isSumRow Then
                    summaryRowCount = summaryRowCount + 1
                    If summarySamples.Count < 15 Then
                        summarySamples.Add Array("第" & CStr(r) & "行", sumPreview, sumReason)
                    End If
                End If

                If Not (isSumRow And summaryRule = "exclude") Then
                    Dim idPreview As String
                    idPreview = ""
                    For i = 0 To UBound(idCols)
                        relCol = idCols(i) - firstCol + 1
                        If Len(idPreview) > 0 Then idPreview = idPreview & " · "
                        idPreview = idPreview & CStr(srcSheet.Cells(r, idCols(i)).Text)
                    Next i

                    For j = 0 To UBound(valCols)
                        If Not (valColIsSummary(j) And summaryRule = "exclude") Then
                            relCol = valCols(j) - firstCol + 1
                            Dim vCell As Variant
                            vCell = vals(relRow, relCol)
                            If IsError(vCell) Then
                                sourceErrorCount = sourceErrorCount + 1
                                If anomalies.Count < 25 Then
                                    anomalies.Add Array(ColIdxToLetter(valCols(j)) & r, "来源含错误值(将跳过)", CStr(srcSheet.Cells(r, valCols(j)).Text))
                                End If
                            Else
                                Dim isBlankVal As Boolean, isZeroVal As Boolean
                                isBlankVal = IsEmpty(vCell) Or IsNull(vCell) Or (VarType(vCell) = vbString And Len(Trim$(CStr(vCell))) = 0)
                                isZeroVal = False
                                If Not isBlankVal And IsNumeric(vCell) Then
                                    On Error Resume Next
                                    If CDbl(vCell) = 0# Then isZeroVal = True
                                    On Error GoTo 0
                                End If

                                If isBlankVal And (emptyPolicy = "skip_blank" Or emptyPolicy = "skip_blank_zero") Then
                                    skippedBlankOrZero = skippedBlankOrZero + 1
                                ElseIf isZeroVal And emptyPolicy = "skip_blank_zero" Then
                                    skippedBlankOrZero = skippedBlankOrZero + 1
                                Else
                                    netOutputRows = netOutputRows + 1
                                    If samples.Count < 8 Then
                                        samples.Add Array(idPreview, valColHeaders(j), CStr(srcSheet.Cells(r, valCols(j)).Text))
                                    End If
                                End If
                            End If
                        End If
                    Next j
                End If
            End If
        Next relRow
    End If

    If canExecute And netOutputRows = 0 Then
        canExecute = False
        blockReason = "转换后可输出的明细行数为 0（可能所选数值列全部为空或为错误值）。"
    End If
    If canExecute And netOutputRows * (UBound(idCols) + 3) > 100000 Then
        canExecute = False
        blockReason = "预计生成单元格数超过单次 10 万格安全上限，请减少展开列或分批转换。"
    End If

    resolvedTargetSheet = DH_UniqueSheetName(wb, targetSheetName)

    DH_UnpivotPreview = Array( _
        srcSheet.Name, _
        totalSourceRows, _
        netOutputRows, _
        UBound(idCols) + 1, _
        activeValColCount, _
        summaryRowCount, _
        summaryColCount, _
        skippedBlankOrZero, _
        sourceErrorCount, _
        resolvedTargetSheet, _
        attrHeader, _
        valHeader, _
        emptyPolicy, _
        summaryRule, _
        canExecute, _
        blockReason, _
        ItemsToArray(samples), _
        ItemsToArray(summarySamples), _
        ItemsToArray(anomalies) _
    )
End Function

Public Function DH_ExecuteUnpivot(ByVal srcSheet As Worksheet, ByVal spec As String) As Long
    Dim wb As Workbook, bounds As Range, origSelection As Range
    Dim idColsCsv As String, valColsCsv As String, attrHeader As String, valHeader As String
    Dim emptyPolicy As String, summaryRule As String, targetSheetName As String
    Dim preview As Variant, canExec As Boolean, blockMsg As String
    Dim idCols() As Long, valCols() As Long, idColHeaders() As String, valColHeaders() As String, valColIsSummary() As Boolean
    Dim i As Long, j As Long, r As Long, c As Long, firstRow As Long, firstCol As Long, relRow As Long, relCol As Long
    Dim vals As Variant, fMask As Variant, fTexts As Variant
    Dim backupPath As String, auditPath As String, resolvedName As String
    Dim state As DH_AppState, frozen As Boolean, outWs As Worksheet, details As New Collection
    Dim writeRow As Long, outRows As Long, skippedSummaries As Long, skippedBlanks As Long, skippedErrors As Long
    Dim sumReason As String, sumPreview As String, isSumRow As Boolean, totalOutCols As Long
    Dim srcCellObj As Range, dstCellObj As Range, vCell As Variant, srcNumFmt As String

    Set wb = srcSheet.Parent
    DH_RequireWorkbook wb, True
    DH_ParseUnpivotSpec spec, idColsCsv, valColsCsv, attrHeader, valHeader, emptyPolicy, summaryRule, targetSheetName

    preview = DH_UnpivotPreview(srcSheet, spec)
    canExec = CBool(preview(14))
    blockMsg = CStr(preview(15))
    If Not canExec Then
        Err.Raise vbObjectError + 974, , blockMsg
    End If

    Set origSelection = DH_OriginalSelection(srcSheet)
    idCols = ParseColList(idColsCsv, 1, 16, "左侧固定列")
    valCols = ParseColList(valColsCsv, 1, 64, "要展开的数值列")
    Set bounds = DH_ContentRange(srcSheet)
    firstRow = bounds.Row
    firstCol = bounds.Column
    vals = DH_Matrix(bounds)
    fMask = DH_FormulaMask(bounds)
    fTexts = DH_Matrix(bounds, True)

    ReDim idColHeaders(0 To UBound(idCols))
    For i = 0 To UBound(idCols)
        relCol = idCols(i) - firstCol + 1
        If IsError(vals(1, relCol)) Or IsEmpty(vals(1, relCol)) Or Len(Trim$(CStr(vals(1, relCol)))) = 0 Then
            idColHeaders(i) = "维度_" & ColIdxToLetter(idCols(i))
        Else
            idColHeaders(i) = Trim$(CStr(vals(1, relCol)))
        End If
    Next i

    ReDim valColHeaders(0 To UBound(valCols))
    ReDim valColIsSummary(0 To UBound(valCols))
    For j = 0 To UBound(valCols)
        relCol = valCols(j) - firstCol + 1
        valColHeaders(j) = Trim$(CStr(vals(1, relCol)))
        If DH_IsSummaryHeaderTitle(valColHeaders(j)) Then
            valColIsSummary(j) = True
            If summaryRule = "exclude" Then
                skippedSummaries = skippedSummaries + 1
                details.Add Array(srcSheet.Name & "!" & ColIdxToLetter(valCols(j)) & "列", "跳过", "排除汇总列「" & valColHeaders(j) & "」", valColHeaders(j), "", "")
            End If
        End If
    Next j

    backupPath = DH_Backup(wb)
    resolvedName = DH_UniqueSheetName(wb, targetSheetName)
    totalOutCols = UBound(idCols) + 3

    On Error GoTo Rollback
    DH_Freeze state: frozen = True

    Set outWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    outWs.Name = resolvedName

    ' Write Row 1 headers: idCols + attrHeader + valHeader
    For i = 0 To UBound(idCols)
        outWs.Cells(1, i + 1).NumberFormat = "@"
        outWs.Cells(1, i + 1).Value2 = idColHeaders(i)
    Next i
    outWs.Cells(1, UBound(idCols) + 2).NumberFormat = "@"
    outWs.Cells(1, UBound(idCols) + 2).Value2 = attrHeader
    outWs.Cells(1, UBound(idCols) + 3).NumberFormat = "@"
    outWs.Cells(1, UBound(idCols) + 3).Value2 = valHeader
    outWs.Rows(1).Font.Bold = True

    writeRow = 1
    For relRow = 2 To bounds.Rows.Count
        r = firstRow + relRow - 1
        Dim rowHasContent As Boolean
        rowHasContent = False
        For c = 1 To bounds.Columns.Count
            If CBool(fMask(relRow, c)) Or IsError(vals(relRow, c)) Then
                rowHasContent = True: Exit For
            ElseIf Not IsEmpty(vals(relRow, c)) And Not IsNull(vals(relRow, c)) Then
                If Len(Trim$(CStr(vals(relRow, c)))) > 0 Then rowHasContent = True: Exit For
            End If
        Next c

        If rowHasContent Then
            isSumRow = DH_IsSuspectedSummaryRow(vals, fMask, fTexts, relRow, bounds.Columns.Count, sumReason, sumPreview)
            If isSumRow And summaryRule = "exclude" Then
                skippedSummaries = skippedSummaries + 1
                details.Add Array(srcSheet.Name & "!第" & CStr(r) & "行", "跳过", "排除小计/合计行（" & sumReason & "）", sumPreview, "", "")
            Else
                For j = 0 To UBound(valCols)
                    If Not (valColIsSummary(j) And summaryRule = "exclude") Then
                        relCol = valCols(j) - firstCol + 1
                        vCell = vals(relRow, relCol)
                        Set srcCellObj = srcSheet.Cells(r, valCols(j))
                        If IsError(vCell) Then
                            skippedErrors = skippedErrors + 1
                            details.Add Array(srcSheet.Name & "!" & srcCellObj.Address(False, False), "跳过", "来源格含错误值(" & CStr(srcCellObj.Text) & ")保护跳过", DH_ValueToken(vCell), "", "")
                        Else
                            Dim isBlankVal As Boolean, isZeroVal As Boolean
                            isBlankVal = IsEmpty(vCell) Or IsNull(vCell) Or (VarType(vCell) = vbString And Len(Trim$(CStr(vCell))) = 0)
                            isZeroVal = False
                            If Not isBlankVal And IsNumeric(vCell) Then
                                On Error Resume Next
                                If CDbl(vCell) = 0# Then isZeroVal = True
                                On Error GoTo Rollback
                            End If

                            If isBlankVal And (emptyPolicy = "skip_blank" Or emptyPolicy = "skip_blank_zero") Then
                                skippedBlanks = skippedBlanks + 1
                            ElseIf isZeroVal And emptyPolicy = "skip_blank_zero" Then
                                skippedBlanks = skippedBlanks + 1
                            Else
                                writeRow = writeRow + 1
                                outRows = outRows + 1
                                ' Copy fixed dimension columns
                                For i = 0 To UBound(idCols)
                                    Dim idRelCol As Long, idVal As Variant, idSrcCell As Range, idDstCell As Range
                                    idRelCol = idCols(i) - firstCol + 1
                                    idVal = vals(relRow, idRelCol)
                                    Set idSrcCell = srcSheet.Cells(r, idCols(i))
                                    Set idDstCell = outWs.Cells(writeRow, i + 1)
                                    If Not IsError(idVal) And Not IsEmpty(idVal) And Not IsNull(idVal) Then
                                        If VarType(idVal) = vbString Then
                                            idDstCell.NumberFormat = "@"
                                            idDstCell.Value2 = CStr(idVal)
                                        ElseIf IsDate(idSrcCell.Value) Then
                                            idDstCell.Value2 = idVal
                                            srcNumFmt = CStr(idSrcCell.NumberFormat)
                                            If Len(srcNumFmt) > 0 And srcNumFmt <> "General" And srcNumFmt <> "@" Then
                                                idDstCell.NumberFormat = srcNumFmt
                                            Else
                                                idDstCell.NumberFormat = "yyyy-mm-dd"
                                            End If
                                        Else
                                            idDstCell.Value2 = idVal
                                        End If
                                    End If
                                Next i
                                ' Write Attribute column
                                outWs.Cells(writeRow, UBound(idCols) + 2).NumberFormat = "@"
                                outWs.Cells(writeRow, UBound(idCols) + 2).Value2 = valColHeaders(j)
                                ' Write Value column
                                Set dstCellObj = outWs.Cells(writeRow, UBound(idCols) + 3)
                                If Not isBlankVal Then
                                    If VarType(vCell) = vbString Then
                                        dstCellObj.NumberFormat = "@"
                                        dstCellObj.Value2 = CStr(vCell)
                                    ElseIf IsDate(srcCellObj.Value) Then
                                        dstCellObj.Value2 = vCell
                                        srcNumFmt = CStr(srcCellObj.NumberFormat)
                                        If Len(srcNumFmt) > 0 And srcNumFmt <> "General" And srcNumFmt <> "@" Then
                                            dstCellObj.NumberFormat = srcNumFmt
                                        Else
                                            dstCellObj.NumberFormat = "yyyy-mm-dd"
                                        End If
                                    Else
                                        dstCellObj.Value2 = vCell
                                    End If
                                End If
                                details.Add Array(outWs.Name & "!第" & CStr(writeRow) & "行", "宽表展开", "从「" & srcSheet.Name & "」" & ColIdxToLetter(valCols(j)) & r & " 展开（" & valColHeaders(j) & "）", "", CStr(srcCellObj.Text), "")
                            End If
                        End If
                    End If
                Next j
            End If
        End If
    Next relRow

    auditPath = DH_SaveAudit("宽表转长表", wb, details)
    DH_LastSummary = "将「" & srcSheet.Name & "」展开为新工作表「" & outWs.Name & "」共 " & CStr(outRows) & " 行明细；排除汇总行/列 " & CStr(skippedSummaries) & " 处，过滤空值/零值 " & CStr(skippedBlanks) & " 格，跳过来源错误值 " & CStr(skippedErrors) & " 格；基线=" & backupPath & "；明细=" & auditPath
    DH_Log "宽表转长表", wb, outWs.Name, "A1:" & ColIdxToLetter(totalOutCols) & CStr(writeRow), outRows, "成功；" & DH_LastSummary, backupPath
    DH_RestoreFocus srcSheet, origSelection
    DH_Restore state
    DH_ExecuteUnpivot = outRows
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
    DH_RestoreFocus srcSheet, origSelection
    If frozen Then DH_Restore state
    On Error GoTo 0
    DH_LogFailure "宽表转长表", wb, srcSheet.Name, idColsCsv & "->" & valColsCsv, errMsg, backupPath
    Err.Raise IIf(errNum <> 0, errNum, vbObjectError + 975), "电子羊表格", "宽表转长表过程中发生异常，已清理临时工作表并回滚：" & errMsg
End Function
