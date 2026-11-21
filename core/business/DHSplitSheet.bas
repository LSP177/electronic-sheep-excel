Attribute VB_Name = "DHSplitSheet"
Option Explicit

' =============================================================================
' DataHelper 1.0 Modern Patch 5: Split Table by Key Column into Worksheets (按列拆表分发)
' Spec format:
'   SplitSheet:<splitColLetter>:<blankKeyRule>:<summaryRule>:<sheetPrefix>
' =============================================================================

Private Const DH_MAX_SPLIT_GROUPS As Long = 50

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

Public Function DH_SanitizeSplitSheetBase(ByVal prefix As String, ByVal rawGroupLabel As String) As String
    Dim s As String, i As Long, ch As String, cleaned As String
    s = Trim$(prefix) & Trim$(rawGroupLabel)
    If Len(s) = 0 Then s = "未命名分组"
    Dim code As Long
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        ' AscW 返回有符号 16 位整数：码点 >= U+8000 的汉字会变成负数，
        ' 直接与 32 比较会把「部/门/长/金/高」等常用字误判为控制字符并替换成下划线。
        code = AscW(ch)
        If code < 0 Then code = code + 65536
        If InStr(1, ":\/?*[]'", ch, vbBinaryCompare) > 0 Or code < 32 Then
            cleaned = cleaned & "_"
        Else
            cleaned = cleaned & ch
        End If
    Next i
    cleaned = Trim$(cleaned)
    If Len(cleaned) = 0 Then cleaned = "分组"
    If Len(cleaned) > 25 Then cleaned = Left$(cleaned, 25)
    DH_SanitizeSplitSheetBase = cleaned
End Function

Private Function ReserveUniqueSheetName(ByVal wb As Workbook, ByVal reserved As Object, ByVal baseName As String) As String
    Dim candidate As String, n As Long, ws As Object, taken As Boolean
    candidate = baseName
    n = 1
    Do
        taken = False
        If reserved.Exists(LCase$(candidate)) Then
            taken = True
        Else
            For Each ws In wb.Sheets
                If StrComp(CStr(ws.Name), candidate, vbTextCompare) = 0 Then
                    taken = True
                    Exit For
                End If
            Next ws
        End If
        If Not taken Then
            reserved(LCase$(candidate)) = True
            ReserveUniqueSheetName = candidate
            Exit Function
        End If
        n = n + 1
        candidate = Left$(baseName, 25) & "_" & CStr(n)
    Loop
End Function

Public Sub DH_ParseSplitSheetSpec(ByVal spec As String, _
    ByRef splitColIdx As Long, ByRef splitColLetter As String, _
    ByRef blankKeyRule As String, ByRef summaryRule As String, _
    ByRef sheetPrefix As String)
    Dim parts As Variant, offset As Long
    parts = Split(spec, ":")
    If UBound(parts) = 4 And StrComp(CStr(parts(0)), "SplitSheet", vbTextCompare) = 0 Then
        offset = 1
    ElseIf UBound(parts) = 3 Then
        offset = 0
    Else
        Err.Raise vbObjectError + 985, , "按列拆表参数格式无效。"
    End If
    splitColLetter = UCase$(Trim$(CStr(parts(offset))))
    splitColIdx = DH_ColLetterToIndex(splitColLetter)
    blankKeyRule = LCase$(Trim$(CStr(parts(offset + 1))))
    summaryRule = LCase$(Trim$(CStr(parts(offset + 2))))
    sheetPrefix = Trim$(CStr(parts(offset + 3)))

    If blankKeyRule <> "bucket" And blankKeyRule <> "skip" And blankKeyRule <> "abort" Then
        Err.Raise vbObjectError + 985, , "空白拆分键处理规则无效，仅支持 bucket / skip / abort。"
    End If
    If summaryRule <> "exclude" And summaryRule <> "keep" Then
        Err.Raise vbObjectError + 985, , "小计/合计行处理规则无效，仅支持 exclude 或 keep。"
    End If
    If Len(sheetPrefix) > 12 Or InStr(1, sheetPrefix, ":", vbBinaryCompare) > 0 Or InStr(1, sheetPrefix, "/", vbBinaryCompare) > 0 Or InStr(1, sheetPrefix, "\", vbBinaryCompare) > 0 Or InStr(1, sheetPrefix, "?", vbBinaryCompare) > 0 Or InStr(1, sheetPrefix, "*", vbBinaryCompare) > 0 Or InStr(1, sheetPrefix, "[", vbBinaryCompare) > 0 Or InStr(1, sheetPrefix, "]", vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 985, , "分表名称前缀不能超过 12 个字符且不能包含特殊符号。"
    End If
End Sub

Public Function DH_SplitSheetPreview(ByVal srcSheet As Worksheet, ByVal spec As String) As Variant
    Dim wb As Workbook, bounds As Range
    Dim splitColIdx As Long, splitColLetter As String, blankKeyRule As String, summaryRule As String, sheetPrefix As String
    Dim firstRow As Long, lastRow As Long, firstCol As Long, lastCol As Long, relCol As Long
    Dim vals As Variant, fMask As Variant, fTexts As Variant
    Dim r As Long, c As Long, relRow As Long, rowHasContent As Boolean
    Dim totalSourceRows As Long, netSplitRows As Long, blankKeyRows As Long, summaryRowCount As Long, sourceErrorCount As Long
    Dim canExecute As Boolean, blockReason As String, splitColHeader As String
    Dim groupOrder As New Collection, groupCounts As Object, groupLabels As Object, groupTargetNames As Object, reservedNames As Object
    Dim groupPreviewList As New Collection, summarySamples As New Collection, anomalies As New Collection
    Dim sumReason As String, sumPreview As String, isSumRow As Boolean
    Dim cellVal As Variant, groupKey As String, dispLabel As String, gIdx As Long, gk As String, resolvedName As String

    Set wb = srcSheet.Parent
    DH_RequireWorkbook wb
    DH_ParseSplitSheetSpec spec, splitColIdx, splitColLetter, blankKeyRule, summaryRule, sheetPrefix

    Set bounds = DH_ContentRange(srcSheet)
    If bounds Is Nothing Then Err.Raise vbObjectError + 986, , "当前工作表为空，没有可拆分的数据。"
    If bounds.Rows.Count < 2 Then Err.Raise vbObjectError + 986, , "当前工作表只有表头行，没有可拆分的明细行。"

    firstRow = bounds.Row
    lastRow = bounds.Row + bounds.Rows.Count - 1
    firstCol = bounds.Column
    lastCol = bounds.Column + bounds.Columns.Count - 1

    If splitColIdx < firstCol Or splitColIdx > lastCol Then
        Err.Raise vbObjectError + 986, , "选择的拆分列（" & splitColLetter & " 列）超出当前表的数据区域范围。"
    End If
    relCol = splitColIdx - firstCol + 1

    vals = DH_Matrix(bounds)
    fMask = DH_FormulaMask(bounds)
    fTexts = DH_Matrix(bounds, True)

    If IsError(vals(1, relCol)) Or IsEmpty(vals(1, relCol)) Or Len(Trim$(CStr(vals(1, relCol)))) = 0 Then
        splitColHeader = splitColLetter & "列"
    Else
        splitColHeader = Trim$(CStr(vals(1, relCol)))
    End If

    Set groupCounts = CreateObject("Scripting.Dictionary")
    groupCounts.CompareMode = 1
    Set groupLabels = CreateObject("Scripting.Dictionary")
    groupLabels.CompareMode = 1
    Set groupTargetNames = CreateObject("Scripting.Dictionary")
    groupTargetNames.CompareMode = 1
    Set reservedNames = CreateObject("Scripting.Dictionary")
    reservedNames.CompareMode = 1

    canExecute = True
    blockReason = ""

    For relRow = 2 To bounds.Rows.Count
        r = firstRow + relRow - 1
        rowHasContent = False
        For c = 1 To bounds.Columns.Count
            If IsError(vals(relRow, c)) Then
                rowHasContent = True
                sourceErrorCount = sourceErrorCount + 1
                If anomalies.Count < 25 Then
                    anomalies.Add Array(ColIdxToLetter(firstCol + c - 1) & r, "来源含错误值(将跳过该格)", CStr(srcSheet.Cells(r, firstCol + c - 1).Text))
                End If
            ElseIf CBool(fMask(relRow, c)) Then
                rowHasContent = True
            ElseIf Not IsEmpty(vals(relRow, c)) And Not IsNull(vals(relRow, c)) Then
                If Len(Trim$(CStr(vals(relRow, c)))) > 0 Then rowHasContent = True
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
                cellVal = vals(relRow, relCol)
                If IsError(cellVal) Then
                    anomalies.Add Array(splitColLetter & r, "拆分键为错误值(跳过该行)", CStr(srcSheet.Cells(r, splitColIdx).Text))
                Else
                    dispLabel = ""
                    If Not IsEmpty(cellVal) And Not IsNull(cellVal) Then
                        dispLabel = Trim$(CStr(srcSheet.Cells(r, splitColIdx).Text))
                    End If

                    If Len(dispLabel) = 0 Then
                        blankKeyRows = blankKeyRows + 1
                        If blankKeyRule = "abort" Then
                            canExecute = False
                            blockReason = "第 " & CStr(r) & " 行的「" & splitColHeader & "」为空，当前规则设为「遇到空值先停下处理」，请补齐后再拆，或切换为「归入（未填写）分表」。"
                            If anomalies.Count < 25 Then
                                anomalies.Add Array(splitColLetter & r, "拆分键为空(阻断)", "请补齐该行分类值或修改空值处理规则")
                            End If
                        ElseIf blankKeyRule = "skip" Then
                            If anomalies.Count < 25 Then
                                anomalies.Add Array(splitColLetter & r, "拆分键为空(按规则跳过)", "该行不进入任何分表")
                            End If
                        Else
                            groupKey = "B:"
                            dispLabel = "（未填写）"
                            If Not groupCounts.Exists(groupKey) Then
                                groupOrder.Add groupKey
                                groupCounts(groupKey) = 1
                                groupLabels(groupKey) = dispLabel
                            Else
                                groupCounts(groupKey) = CLng(groupCounts(groupKey)) + 1
                            End If
                            netSplitRows = netSplitRows + 1
                        End If
                    Else
                        groupKey = "V:" & LCase$(dispLabel)
                        If Not groupCounts.Exists(groupKey) Then
                            groupOrder.Add groupKey
                            groupCounts(groupKey) = 1
                            groupLabels(groupKey) = dispLabel
                        Else
                            groupCounts(groupKey) = CLng(groupCounts(groupKey)) + 1
                        End If
                        netSplitRows = netSplitRows + 1
                    End If
                End If
            End If
        End If
    Next relRow

    For gIdx = 1 To groupOrder.Count
        gk = CStr(groupOrder(gIdx))
        resolvedName = ReserveUniqueSheetName(wb, reservedNames, DH_SanitizeSplitSheetBase(sheetPrefix, CStr(groupLabels(gk))))
        groupTargetNames(gk) = resolvedName
        If gIdx <= 20 Then
            groupPreviewList.Add Array(CStr(groupLabels(gk)), resolvedName, CLng(groupCounts(gk)))
        End If
    Next gIdx

    If canExecute And groupOrder.Count = 0 Then
        canExecute = False
        blockReason = "过滤后没有任何可拆分的有效分组数据。"
    ElseIf canExecute And groupOrder.Count > DH_MAX_SPLIT_GROUPS Then
        canExecute = False
        blockReason = "「" & splitColHeader & "」列包含 " & CStr(groupOrder.Count) & " 个不同分组，超过单次拆表安全上限（" & CStr(DH_MAX_SPLIT_GROUPS) & " 张工作表）。请确认是否误选了订单号/手机号等唯一编号列。"
    ElseIf canExecute And netSplitRows * bounds.Columns.Count > 100000 Then
        canExecute = False
        blockReason = "预计写入单元格总数超过单次 10 万格安全上限，请分批处理。"
    End If

    DH_SplitSheetPreview = Array( _
        srcSheet.Name, _
        splitColLetter, _
        splitColHeader, _
        totalSourceRows, _
        netSplitRows, _
        groupOrder.Count, _
        blankKeyRows, _
        summaryRowCount, _
        sourceErrorCount, _
        bounds.Columns.Count, _
        blankKeyRule, _
        summaryRule, _
        sheetPrefix, _
        canExecute, _
        blockReason, _
        ItemsToArray(groupPreviewList), _
        ItemsToArray(summarySamples), _
        ItemsToArray(anomalies) _
    )
End Function

Public Function DH_ExecuteSplitSheet(ByVal srcSheet As Worksheet, ByVal spec As String) As Long
    Dim wb As Workbook, bounds As Range, origSelection As Range
    Dim splitColIdx As Long, splitColLetter As String, blankKeyRule As String, summaryRule As String, sheetPrefix As String
    Dim preview As Variant, canExec As Boolean, blockMsg As String, splitColHeader As String
    Dim firstRow As Long, lastRow As Long, firstCol As Long, lastCol As Long, relCol As Long, colCount As Long
    Dim vals As Variant, fMask As Variant, fTexts As Variant, headerTexts() As String
    Dim groupOrder As New Collection, groupRows As Object, groupLabels As Object, reservedNames As Object
    Dim r As Long, c As Long, relRow As Long, rowHasContent As Boolean
    Dim sumReason As String, sumPreview As String, isSumRow As Boolean
    Dim cellVal As Variant, dispLabel As String, groupKey As String, rColl As Collection
    Dim backupPath As String, auditPath As String
    Dim state As DH_AppState, frozen As Boolean, createdSheets As New Collection, details As New Collection
    Dim gIdx As Long, gk As String, outSheetName As String, outWs As Worksheet, writeRow As Long, itemRow As Variant
    Dim srcCellObj As Range, dstCellObj As Range, vCell As Variant, srcNumFmt As String
    Dim writtenRows As Long, skippedSummaryRows As Long, skippedBlankRows As Long, skippedErrorCells As Long, firstCreatedName As String, lastCreatedName As String

    Set wb = srcSheet.Parent
    DH_RequireWorkbook wb, True
    DH_ParseSplitSheetSpec spec, splitColIdx, splitColLetter, blankKeyRule, summaryRule, sheetPrefix

    preview = DH_SplitSheetPreview(srcSheet, spec)
    canExec = CBool(preview(13))
    blockMsg = CStr(preview(14))
    splitColHeader = CStr(preview(2))
    If Not canExec Then
        Err.Raise vbObjectError + 987, , blockMsg
    End If

    Set origSelection = DH_OriginalSelection(srcSheet)
    Set bounds = DH_ContentRange(srcSheet)
    firstRow = bounds.Row
    lastRow = bounds.Row + bounds.Rows.Count - 1
    firstCol = bounds.Column
    lastCol = bounds.Column + bounds.Columns.Count - 1
    colCount = bounds.Columns.Count
    relCol = splitColIdx - firstCol + 1

    vals = DH_Matrix(bounds)
    fMask = DH_FormulaMask(bounds)
    fTexts = DH_Matrix(bounds, True)

    ReDim headerTexts(1 To colCount)
    For c = 1 To colCount
        If IsError(vals(1, c)) Or IsEmpty(vals(1, c)) Or IsNull(vals(1, c)) Then
            headerTexts(c) = ""
        Else
            headerTexts(c) = Trim$(CStr(vals(1, c)))
        End If
    Next c

    Set groupRows = CreateObject("Scripting.Dictionary")
    groupRows.CompareMode = 1
    Set groupLabels = CreateObject("Scripting.Dictionary")
    groupLabels.CompareMode = 1
    Set reservedNames = CreateObject("Scripting.Dictionary")
    reservedNames.CompareMode = 1

    For relRow = 2 To bounds.Rows.Count
        r = firstRow + relRow - 1
        rowHasContent = False
        For c = 1 To colCount
            If CBool(fMask(relRow, c)) Or IsError(vals(relRow, c)) Then
                rowHasContent = True: Exit For
            ElseIf Not IsEmpty(vals(relRow, c)) And Not IsNull(vals(relRow, c)) Then
                If Len(Trim$(CStr(vals(relRow, c)))) > 0 Then rowHasContent = True: Exit For
            End If
        Next c

        If rowHasContent Then
            isSumRow = DH_IsSuspectedSummaryRow(vals, fMask, fTexts, relRow, colCount, sumReason, sumPreview)
            If isSumRow And summaryRule = "exclude" Then
                skippedSummaryRows = skippedSummaryRows + 1
                details.Add Array(srcSheet.Name & "!第" & CStr(r) & "行", "跳过", "排除小计/合计行（" & sumReason & "）", sumPreview, "", "")
            Else
                cellVal = vals(relRow, relCol)
                If IsError(cellVal) Then
                    skippedErrorCells = skippedErrorCells + 1
                    details.Add Array(srcSheet.Name & "!" & splitColLetter & r, "跳过", "拆分键单元格为错误值，跳过该行", DH_ValueToken(cellVal), "", "")
                Else
                    dispLabel = ""
                    If Not IsEmpty(cellVal) And Not IsNull(cellVal) Then
                        dispLabel = Trim$(CStr(srcSheet.Cells(r, splitColIdx).Text))
                    End If
                    If Len(dispLabel) = 0 Then
                        If blankKeyRule = "skip" Then
                            skippedBlankRows = skippedBlankRows + 1
                            details.Add Array(srcSheet.Name & "!" & splitColLetter & r, "跳过", "拆分键为空白，按 skip 规则跳过", "", "", "")
                        Else
                            groupKey = "B:"
                            dispLabel = "（未填写）"
                            If Not groupRows.Exists(groupKey) Then
                                Set rColl = New Collection
                                groupRows.Add groupKey, rColl
                                groupLabels.Add groupKey, dispLabel
                                groupOrder.Add groupKey
                            End If
                            groupRows(groupKey).Add relRow
                        End If
                    Else
                        groupKey = "V:" & LCase$(dispLabel)
                        If Not groupRows.Exists(groupKey) Then
                            Set rColl = New Collection
                            groupRows.Add groupKey, rColl
                            groupLabels.Add groupKey, dispLabel
                            groupOrder.Add groupKey
                        End If
                        groupRows(groupKey).Add relRow
                    End If
                End If
            End If
        End If
    Next relRow

    backupPath = DH_Backup(wb)

    On Error GoTo Rollback
    DH_Freeze state: frozen = True

    For gIdx = 1 To groupOrder.Count
        gk = CStr(groupOrder(gIdx))
        outSheetName = ReserveUniqueSheetName(wb, reservedNames, DH_SanitizeSplitSheetBase(sheetPrefix, CStr(groupLabels(gk))))
        Set outWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        outWs.Name = outSheetName
        createdSheets.Add outWs
        If gIdx = 1 Then firstCreatedName = outSheetName
        lastCreatedName = outSheetName

        ' Write headers in Row 1
        For c = 1 To colCount
            outWs.Cells(1, c).NumberFormat = "@"
            outWs.Cells(1, c).Value2 = headerTexts(c)
        Next c
        outWs.Rows(1).Font.Bold = True

        writeRow = 1
        Set rColl = groupRows(gk)
        For Each itemRow In rColl
            relRow = CLng(itemRow)
            r = firstRow + relRow - 1
            writeRow = writeRow + 1
            writtenRows = writtenRows + 1
            For c = 1 To colCount
                vCell = vals(relRow, c)
                Set srcCellObj = srcSheet.Cells(r, firstCol + c - 1)
                Set dstCellObj = outWs.Cells(writeRow, c)
                If IsError(vCell) Then
                    skippedErrorCells = skippedErrorCells + 1
                    details.Add Array(srcSheet.Name & "!" & srcCellObj.Address(False, False), "跳过", "来源格含错误值(" & CStr(srcCellObj.Text) & ")保护跳过", DH_ValueToken(vCell), "", "")
                ElseIf Not IsEmpty(vCell) And Not IsNull(vCell) Then
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
            Next c
        Next itemRow
        details.Add Array(outSheetName & "!A1:" & ColIdxToLetter(colCount) & CStr(writeRow), "新建分表", "按「" & splitColHeader & "=" & CStr(groupLabels(gk)) & "」拆分生成 " & CStr(rColl.Count) & " 行明细", CStr(groupLabels(gk)), outSheetName, "")
    Next gIdx

    auditPath = DH_SaveAudit("按列拆表", wb, details)
    DH_LastSummary = "按「" & splitColHeader & "」将「" & srcSheet.Name & "」拆分为 " & CStr(createdSheets.Count) & " 张工作表（" & firstCreatedName & IIf(createdSheets.Count > 1, " .. " & lastCreatedName, "") & "），共分发 " & CStr(writtenRows) & " 行明细；排除小计/合计 " & CStr(skippedSummaryRows) & " 行，跳过空键 " & CStr(skippedBlankRows) & " 行，跳过来源错误值 " & CStr(skippedErrorCells) & " 格；基线=" & backupPath & "；明细=" & auditPath
    DH_Log "按列拆表", wb, srcSheet.Name, bounds.Address(False, False), writtenRows, "成功；" & DH_LastSummary, backupPath
    DH_RestoreFocus srcSheet, origSelection
    DH_Restore state
    DH_ExecuteSplitSheet = writtenRows
    Exit Function

Rollback:
    Dim errMsg As String, errNum As Long, wsDel As Object
    errNum = Err.Number: errMsg = Err.Description
    On Error Resume Next
    Application.DisplayAlerts = False
    For Each wsDel In createdSheets
        wsDel.Delete
    Next wsDel
    Application.DisplayAlerts = True
    DH_RestoreFocus srcSheet, origSelection
    If frozen Then DH_Restore state
    On Error GoTo 0
    DH_LogFailure "按列拆表", wb, srcSheet.Name, splitColLetter, errMsg, backupPath
    Err.Raise IIf(errNum <> 0, errNum, vbObjectError + 988), "电子羊表格", "按列拆表过程中发生异常，已自动清理全部临时分表并回滚：" & errMsg
End Function
