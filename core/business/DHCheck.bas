Attribute VB_Name = "DHCheck"
Option Explicit

Public Type DH_HealthStats
    Address As String
    Characters As Long
    DirtyCells As Long
    Duplicates As Long
    BlankRows As Long
    Merges As Long
    MergedCells As Double
    NumericTexts As Long
    Eligible As Long
    ProtectedNumbers As Long
    UnitTexts As Long
    Formulas As Long
    RiskCells As Long
    IfErrorHints As Long
    LookupHints As Long
    HardcodedHints As Long
    DateCandidates As Long
End Type



Private Sub Detail(ByVal details As Collection, ByVal kind As String, ByVal ws As Worksheet, ByVal address As String, ByVal note As String)
    If details.Count < DH_MAX_DETAILS Then details.Add Array(kind, ws.Name, address, note)
End Sub

Public Sub DH_ScanWorksheet(ByVal ws As Worksheet, ByRef stats As DH_HealthStats, ByVal details As Collection)
    Dim cleared As DH_HealthStats, rng As Range, values As Variant, formulas As Variant, mask As Variant
    Dim seen As Object, merges As Collection, area As Range, r As Long, c As Long, firstRow As Long, firstColumn As Long
    Dim rowEmpty As Boolean, key As String, count As Long, risk As Long, location As String
    Dim dateBlocked() As Boolean, header As Range, overlap As Range, part As Range, table As ListObject, rr As Long, cc As Long, serial As Double
    Dim value As Double, unit As String, reason As String, looks As Boolean, accepted As Boolean
    Dim txt As String
    Dim colUnits As Object, colMixed As Object, colIdx As Long, obsUnit As String
    Dim probeVal As Double, probeUnit As String, probeReason As String, probeLooks As Boolean
    stats = cleared
    Set rng = DH_ContentRange(ws)
    stats.Address = "空工作表"
    If Not rng Is Nothing Then
        If rng.CountLarge > DH_MAX_CELLS Then Err.Raise vbObjectError + 721, , "数据矩形超过100万单元格，请拆分后体检。"
        stats.Address = rng.Address(False, False)
    End If
    Set merges = DH_MergedAreas(ws.UsedRange)
    ' 只统计“拆得动的”合并（纵向单列）；横向/二维合并会按规则跳过，不计入待处理
    stats.Merges = 0
    For Each area In merges
        stats.MergedCells = stats.MergedCells + area.CountLarge
        If area.Columns.Count = 1 And area.Rows.Count > 1 Then stats.Merges = stats.Merges + 1
    Next area
    If rng Is Nothing Then Exit Sub
    firstRow = rng.Row: firstColumn = rng.Column
    values = DH_Matrix(rng): formulas = DH_Matrix(rng, True): mask = DH_FormulaMask(rng)
    ReDim dateBlocked(1 To UBound(values, 1), 1 To UBound(values, 2))
    For Each table In ws.ListObjects
        Set header = Nothing
        If table.ShowHeaders Then Set header = table.HeaderRowRange
        If Not header Is Nothing Then
            Set overlap = Application.Intersect(rng, header)
            If Not overlap Is Nothing Then
                For Each part In overlap.Areas
                    For rr = part.Row - firstRow + 1 To part.Row - firstRow + part.Rows.Count
                        For cc = part.Column - firstColumn + 1 To part.Column - firstColumn + part.Columns.Count
                            dateBlocked(rr, cc) = True
                        Next cc
                    Next rr
                Next part
            End If
        End If
    Next table
    For Each area In merges
        Set overlap = Application.Intersect(rng, area)
        If Not overlap Is Nothing Then
            For rr = overlap.Row - firstRow + 1 To overlap.Row - firstRow + overlap.Rows.Count
                For cc = overlap.Column - firstColumn + 1 To overlap.Column - firstColumn + overlap.Columns.Count
                    dateBlocked(rr, cc) = True
                Next cc
            Next rr
        End If
    Next area
    ' 数字候选要按“列”判混单位：和 DH_ConvertNumbers 同一口径，避免清完还报同样数量
    Set colUnits = CreateObject("Scripting.Dictionary")
    Set colMixed = CreateObject("Scripting.Dictionary")
    For c = 1 To UBound(values, 2)
        colIdx = firstColumn + c - 1
        If Not colUnits.Exists(CStr(colIdx)) Then colUnits.Add CStr(colIdx), CreateObject("Scripting.Dictionary")
        For r = 1 To UBound(values, 1)
            If VarType(values(r, c)) = vbString Then
                If Not mask(r, c) Then
                    If DH_TryNumber(CStr(values(r, c)), probeVal, probeUnit, probeReason, probeLooks) Then
                        If probeLooks Then
                            obsUnit = DH_ObservedUnit(CStr(values(r, c)))
                            If Len(obsUnit) > 0 Then
                                If StrComp(obsUnit, probeUnit, vbTextCompare) <> 0 Then colMixed(CStr(colIdx)) = True
                                If Not colUnits(CStr(colIdx)).Exists(obsUnit) Then colUnits(CStr(colIdx)).Add obsUnit, True
                            End If
                        End If
                    End If
                End If
            End If
        Next r
        If colUnits(CStr(colIdx)).count > 1 Then colMixed(CStr(colIdx)) = True
    Next c
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbBinaryCompare
    For r = 1 To UBound(values, 1)
        rowEmpty = True: key = ""
        For c = 1 To UBound(values, 2)
            location = "行" & (firstRow + r - 1) & " 列" & (firstColumn + c - 1)
            If Not IsEmpty(values(r, c)) Or mask(r, c) Then rowEmpty = False
            key = key & DH_ValueToken(values(r, c))
            If mask(r, c) Then key = key & "FORMULA:" & DH_ValueToken(formulas(r, c))
            If mask(r, c) Then
                stats.Formulas = stats.Formulas + 1
                risk = DH_FormulaRiskMask(CStr(formulas(r, c)))
                If risk > 0 Then
                    stats.RiskCells = stats.RiskCells + 1
                    If (risk And 1) <> 0 Then stats.IfErrorHints = stats.IfErrorHints + 1
                    If (risk And 2) <> 0 Then stats.LookupHints = stats.LookupHints + 1
                    If (risk And 4) <> 0 Then stats.HardcodedHints = stats.HardcodedHints + 1
                    Detail details, "公式安全提示", ws, location, DH_FormulaRiskText(risk)
                End If
            End If
            If VarType(values(r, c)) = vbString Then
                If Not mask(r, c) Then
                    If Not dateBlocked(r, c) Then
                        If DH_TryDate(CStr(values(r, c)), serial, reason, ws.Parent.Date1904) Then stats.DateCandidates = stats.DateCandidates + 1
                    End If
                    accepted = DH_TryNumber(CStr(values(r, c)), value, unit, reason, looks)
                    If looks Then
                        ' 只把“点一下真会被改”的计入待处理；受保护（前导零/长编号等）单独计数
                        If accepted Then
                            If colMixed.Exists(CStr(firstColumn + c - 1)) Then
                                stats.ProtectedNumbers = stats.ProtectedNumbers + 1
                            Else
                                stats.NumericTexts = stats.NumericTexts + 1
                                stats.Eligible = stats.Eligible + 1
                            End If
                        Else
                            stats.ProtectedNumbers = stats.ProtectedNumbers + 1
                        End If
                        If Len(unit) > 0 Then stats.UnitTexts = stats.UnitTexts + 1
                    End If
                End If
                txt = CStr(values(r, c))
                count = DH_EdgeSpaceCount(txt)
                If count > 0 Then
                    stats.Characters = stats.Characters + count: stats.DirtyCells = stats.DirtyCells + 1
                    Detail details, "首尾空格", ws, location, count & "个首尾空白字符（中间的分隔空格不算）"
                End If
            End If
        Next c
        If rowEmpty Then
            stats.BlankRows = stats.BlankRows + 1
            Detail details, "空白行", ws, CStr(firstRow + r - 1), "范围内无值、无公式；若含备注/批注，删除时保护"
        ElseIf seen.Exists(key) Then
            stats.Duplicates = stats.Duplicates + 1
            Detail details, "重复记录", ws, CStr(firstRow + r - 1), "首次出现在行" & seen(key)
        Else
            seen.Add key, firstRow + r - 1
        End If
    Next r
    ' 用执行器的同一函数算「真能转换」的数量，保证显示数与执行数一致
    On Error Resume Next
    Dim exactNum As Long
    exactNum = DH_ConvertNumbers(rng, True)
    If Err.Number = 0 Then
        stats.ProtectedNumbers = stats.ProtectedNumbers + (stats.NumericTexts - exactNum)
        If stats.ProtectedNumbers < 0 Then stats.ProtectedNumbers = 0
        stats.NumericTexts = exactNum
        stats.Eligible = exactNum
    End If
    Err.Clear
    On Error GoTo 0
    ' 隐藏列不改变重复计数；由面板显示「先取消隐藏才能删」
End Sub

Private Sub AddStats(ByRef total As DH_HealthStats, ByRef part As DH_HealthStats)
    total.Characters = total.Characters + part.Characters: total.DirtyCells = total.DirtyCells + part.DirtyCells
    total.Duplicates = total.Duplicates + part.Duplicates: total.BlankRows = total.BlankRows + part.BlankRows
    total.Merges = total.Merges + part.Merges: total.MergedCells = total.MergedCells + part.MergedCells
    total.NumericTexts = total.NumericTexts + part.NumericTexts: total.Eligible = total.Eligible + part.Eligible
    total.ProtectedNumbers = total.ProtectedNumbers + part.ProtectedNumbers: total.UnitTexts = total.UnitTexts + part.UnitTexts
    total.Formulas = total.Formulas + part.Formulas: total.RiskCells = total.RiskCells + part.RiskCells
    total.IfErrorHints = total.IfErrorHints + part.IfErrorHints: total.LookupHints = total.LookupHints + part.LookupHints
    total.HardcodedHints = total.HardcodedHints + part.HardcodedHints
    total.DateCandidates = total.DateCandidates + part.DateCandidates
End Sub

Private Function SheetSummary(ByVal ws As Worksheet, ByRef stats As DH_HealthStats) As Variant
    SheetSummary = Array(ws.Name, "完成；" & stats.Address, "字符=" & stats.Characters & "；合并=" & stats.Merges & _
        "；伪数字=" & stats.NumericTexts & "；空行=" & stats.BlankRows & "；日期候选=" & stats.DateCandidates, "公式=" & stats.Formulas & "；风险提示单元格=" & stats.RiskCells)
End Function

Private Function TryScan(ByVal ws As Worksheet, ByRef stats As DH_HealthStats, ByVal details As Collection, ByRef message As String) As Boolean
    On Error GoTo Failed
    DH_ScanWorksheet ws, stats, details
    TryScan = True: Exit Function
Failed:
    message = Err.Description
End Function

Private Function HealthReport(ByVal wb As Workbook, ByVal title As String, ByVal scope As String, ByRef stats As DH_HealthStats, _
    ByVal summaries As Collection, ByVal details As Collection, ByVal completed As Long, ByVal skipped As Long) As Worksheet
    Dim ws As Worksheet, data As Variant, item As Variant, i As Long, j As Long, detailRow As Long, totalRows As Long, baseline As String
    detailRow = 22 + summaries.Count: totalRows = detailRow + details.Count
    ReDim data(1 To totalRows, 1 To 4)
    data(1, 1) = title
    data(2, 1) = "工作簿": data(2, 2) = wb.Name
    data(3, 1) = "工作表": data(3, 2) = scope
    data(4, 1) = "扫描时间": data(4, 2) = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    baseline = DH_ExistingBaseline(wb)
    data(5, 1) = "当日基线路径": data(5, 2) = IIf(Len(baseline) > 0, baseline, "尚未建立；只读体检不创建备份")
    data(6, 1) = "数据范围": data(6, 2) = stats.Address
    data(7, 1) = "数据与公式健康安全看板"
    data(8, 1) = "隐形字符数量": data(8, 2) = stats.Characters: data(8, 3) = "涉及单元格": data(8, 4) = stats.DirtyCells
    data(9, 1) = "合并区域数量": data(9, 2) = stats.Merges: data(9, 3) = "合并涉及单元格": data(9, 4) = stats.MergedCells
    data(10, 1) = "伪文本数字数量": data(10, 2) = stats.NumericTexts: data(10, 3) = "语法可转换候选": data(10, 4) = stats.Eligible
    data(11, 1) = "空白行数量": data(11, 2) = stats.BlankRows: data(11, 3) = "重复记录数量": data(11, 4) = stats.Duplicates
    data(12, 1) = "公式风险提示单元格": data(12, 2) = stats.RiskCells: data(12, 3) = "公式总数": data(12, 4) = stats.Formulas
    data(13, 1) = "IFERROR吞错提示": data(13, 2) = stats.IfErrorHints: data(13, 3) = "VLOOKUP固定列号": data(13, 4) = stats.LookupHints
    data(14, 1) = "算术硬编码数字提示": data(14, 2) = stats.HardcodedHints: data(14, 3) = "受保护数字文本": data(14, 4) = stats.ProtectedNumbers
    data(15, 1) = "带单位数字文本": data(15, 2) = stats.UnitTexts: data(15, 3) = "明确日期文本候选": data(15, 4) = stats.DateCandidates
    data(16, 1) = "公式提示只供人工复核，不代表公式必错；不修改公式，不计算外部链接。"
    data(17, 1) = "完成工作表": data(17, 2) = completed: data(17, 3) = "未完成工作表": data(17, 4) = skipped
    data(18, 1) = "只读检查，不自动保存。空白行按无值/公式计数；含备注/批注的空白行删除时保护，空白行数量不等于可删除数。字符数含合法英文空格；数字仅标准十进制/支持单位；重复按值、类型、大小写及公式文本判等；删重另保护公式/备注。"
    data(19, 1) = "逐表扫描结果"
    data(20, 1) = "工作表": data(20, 2) = "状态与范围": data(20, 3) = "数据检查": data(20, 4) = "公式检查"
    For i = 1 To summaries.Count
        item = summaries(i)
        For j = 0 To 3: data(20 + i, j + 1) = item(j): Next j
    Next i
    data(detailRow, 1) = "问题类型": data(detailRow, 2) = "工作表": data(detailRow, 3) = "位置": data(detailRow, 4) = "说明（最多1000条）"
    For i = 1 To details.Count
        item = details(i)
        For j = 0 To 3: data(detailRow + i, j + 1) = item(j): Next j
    Next i
    Set ws = DH_HealthSheet(wb)
    ws.Range("A1").Resize(totalRows, 4).Value2 = DH_LiteralMatrix(data)
    DH_ReportStyle ws, totalRows
    ws.Columns("A").ColumnWidth = 28: ws.Columns("B").ColumnWidth = 25
    ws.Columns("C").ColumnWidth = 27: ws.Columns("D").ColumnWidth = 45
    ws.Range("A1:D1").Merge
    For i = 2 To 6: ws.Range("B" & i & ":D" & i).Merge: ws.Rows(i).RowHeight = 25: Next i
    ws.Rows(5).RowHeight = 42
    ws.Range("A7:D7").Merge: ws.Range("A7:D7").Interior.Color = RGB(31, 78, 120)
    ws.Range("A7").Font.Color = vbWhite: ws.Range("A7").Font.Bold = True: ws.Rows(7).RowHeight = 28
    ws.Range("B8:B15").Font.Bold = True: ws.Range("D8:D15").Font.Bold = True
    ws.Range("A8:D15").Rows.AutoFit
    ws.Range("A16:D16").Merge: ws.Rows(16).RowHeight = 32
    ws.Range("A18:D18").Merge: ws.Rows(18).RowHeight = 60
    ws.Range("A19:D20").Font.Bold = True: ws.Range("A" & detailRow & ":D" & detailRow).Font.Bold = True
    DH_BackupLink ws.Range("B5"), wb
    Set HealthReport = ws
End Function

Public Function DH_CheckWorksheet(ByVal ws As Worksheet) As String
    Dim wb As Workbook, report As Worksheet, selected As Range, stats As DH_HealthStats, details As Collection, summaries As Collection
    Dim state As DH_AppState, frozen As Boolean, message As String, code As Long
    On Error GoTo Failed
    If ws Is Nothing Then Err.Raise vbObjectError + 720, , "请选择工作表。"
    Set wb = ws.Parent: DH_RequireWorkbook wb, True
    If DH_IsReportSheet(ws) Or DH_IsPreviousReport(ws) Then Err.Raise vbObjectError + 852, , "请回到业务表后执行体检；当前表是插件报告或旧报告留存。"
    DH_ReportPreflight wb
    Set selected = DH_OriginalSelection(ws)
    Set details = New Collection: Set summaries = New Collection
    DH_Freeze state: frozen = True
    DH_ScanWorksheet ws, stats, details
    summaries.Add SheetSummary(ws, stats)
    Set report = HealthReport(wb, "数据体检报告", ws.Name, stats, summaries, details, 1, 0)
    DH_Log "表格体检", wb, ws.Name, stats.Address, 0, "成功；只读检查；公式提示=" & stats.RiskCells
    DH_CheckWorksheet = report.Name
    DH_RestoreFocus ws, selected: DH_Restore state
    Exit Function
Failed:
    message = Err.Description: code = Err.Number
    DH_RestoreFocus ws, selected
    If frozen Then DH_Restore state
    DH_LogFailure "表格体检", wb, "", "", message
    Err.Raise code, "电子羊表格", message
End Function

Public Function DH_CheckWorkbook(ByVal wb As Workbook) As String
    Dim originalSheet As Worksheet, selected As Range, ws As Worksheet, report As Worksheet, stats As DH_HealthStats, total As DH_HealthStats
    Dim details As Collection, summaries As Collection, completed As Long, skipped As Long, message As String, code As Long
    Dim state As DH_AppState, frozen As Boolean
    On Error GoTo Failed
    DH_RequireWorkbook wb, True
    DH_ReportPreflight wb
    If TypeOf Application.ActiveSheet Is Worksheet Then Set originalSheet = Application.ActiveSheet
    If Not originalSheet Is Nothing Then Set selected = DH_OriginalSelection(originalSheet)
    Set details = New Collection: Set summaries = New Collection
    DH_Freeze state: frozen = True
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            If Not DH_IsReportSheet(ws) And Not DH_IsPreviousReport(ws) Then
                If TryScan(ws, stats, details, message) Then
                    AddStats total, stats: completed = completed + 1: summaries.Add SheetSummary(ws, stats)
                Else
                    skipped = skipped + 1: summaries.Add Array(ws.Name, "未完成", message, "汇总不包含此表")
                End If
            End If
        End If
    Next ws
    total.Address = "已完成表的范围与跳过原因见逐表结果；隐藏表、插件报告及可确认的旧报告留存不参与"
    Set report = HealthReport(wb, "全簿健康总览报告", "全部可见业务工作表", total, summaries, details, completed, skipped)
    DH_Log "全簿体检", wb, "", "可见业务工作表", 0, "成功生成总览；完成=" & completed & "；未完成=" & skipped
    DH_CheckWorkbook = report.Name
    DH_RestoreFocus originalSheet, selected: DH_Restore state
    Exit Function
Failed:
    message = Err.Description: code = Err.Number
    DH_RestoreFocus originalSheet, selected
    If frozen Then DH_Restore state
    DH_LogFailure "全簿体检", wb, "", "", message
    Err.Raise code, "电子羊表格", message
End Function

' 只数“首尾”空白（半角/全角空格、制表符）；中间的分隔空格不算，避免永远清不完
Public Function DH_EdgeSpaceCount(ByVal txt As String) As Long
    Dim i As Long, n As Long, code As Long, total As Long
    total = Len(txt)
    If total = 0 Then DH_EdgeSpaceCount = 0: Exit Function
    For i = 1 To total
        code = AscW(Mid$(txt, i, 1))
        If code = 32 Or code = 12288 Or code = 9 Then n = n + 1 Else Exit For
    Next i
    If n >= total Then DH_EdgeSpaceCount = total: Exit Function
    For i = total To 1 Step -1
        code = AscW(Mid$(txt, i, 1))
        If code = 32 Or code = 12288 Or code = 9 Then n = n + 1 Else Exit For
    Next i
    DH_EdgeSpaceCount = n
End Function


' 内容区里是否存在隐藏的行或列（删重复会被安全拒绝）
Private Function HasHiddenIn(ByVal rng As Range) As Boolean
    Dim ws As Worksheet, i As Long
    Set ws = rng.Worksheet
    On Error Resume Next
    For i = rng.Row To rng.Row + rng.Rows.Count - 1
        If ws.Rows(i).Hidden Then HasHiddenIn = True: Exit Function
    Next i
    For i = rng.Column To rng.Column + rng.Columns.Count - 1
        If ws.Columns(i).Hidden Then HasHiddenIn = True: Exit Function
    Next i
    On Error GoTo 0
End Function
