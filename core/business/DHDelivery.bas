Attribute VB_Name = "DHDelivery"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal format As Long, ByVal memory As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal flags As Long, ByVal bytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal memory As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal memory As LongPtr) As Long
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal memory As LongPtr) As LongPtr
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal destination As LongPtr, ByVal source As LongPtr, ByVal length As LongPtr)
#End If

Public Sub DH_CopyReceipt(ByVal text As String)
    Dim memory As LongPtr, pointer As LongPtr, bytes As LongPtr, opened As Boolean, transferred As Boolean
    Dim message As String, code As Long
    On Error GoTo Failed
    If Len(text) = 0 Then Err.Raise vbObjectError + 890, , "没有可复制的文案。"
    bytes = LenB(text): memory = GlobalAlloc(&H42, bytes + 2)
    If memory = 0 Then Err.Raise vbObjectError + 891, , "无法分配剪贴板内存。"
    pointer = GlobalLock(memory)
    If pointer = 0 Then Err.Raise vbObjectError + 891, , "无法锁定剪贴板内存。"
    CopyMemory pointer, StrPtr(text), bytes
    GlobalUnlock memory
    If OpenClipboard(Application.hwnd) = 0 Then Err.Raise vbObjectError + 892, , "剪贴板正被其他程序占用，请稍后重试或在文本框内复制。"
    opened = True
    If EmptyClipboard() = 0 Then Err.Raise vbObjectError + 892, , "无法准备剪贴板。"
    If SetClipboardData(13, memory) = 0 Then Err.Raise vbObjectError + 892, , "系统拒绝写入剪贴板。"
    transferred = True
    CloseClipboard
    Exit Sub
Failed:
    message = Err.Description: code = Err.Number
    If opened Then CloseClipboard
    If memory <> 0 And Not transferred Then GlobalFree memory
    Err.Raise code, "复制交差小票", message
End Sub

Public Function DH_BuildReceipt(ByVal wb As Workbook) As String
    Dim lines As Variant, fields As Variant, counts As Object, runs As Object, i As Long, key As Variant, result As String, actions As Long
    DH_RequireWorkbook wb
    Set counts = CreateObject("Scripting.Dictionary"): Set runs = CreateObject("Scripting.Dictionary")
    lines = Split(DH_ReadLog(), vbCrLf)
    For i = 1 To UBound(lines)
        fields = Split(lines(i), vbTab)
        If UBound(fields) = 7 Then
            If StrComp(fields(2), wb.FullName, vbTextCompare) = 0 And Left$(fields(6), 2) = "成功" Then
                Select Case fields(1)
                    Case "区域清理", "删除空行", "拆分合并", "数字转换", _
                         "单位拆分", "两表对账", "日期规范化", "高亮重复项", _
                         "删除重复行", "敏感数据脱敏", "一键交表礼仪", "清除临时标记", _
                         "跨表匹配补全", "多表纵向追加", "宽表转长表", "按列拆表"
                        If IsNumeric(fields(5)) Then
                            If CLng(fields(5)) > 0 Then
                                If Not counts.Exists(fields(1)) Then counts.Add fields(1), 0: runs.Add fields(1), 0
                                counts(fields(1)) = counts(fields(1)) + CLng(fields(5))
                                runs(fields(1)) = runs(fields(1)) + 1: actions = actions + 1
                            End If
                        End If
                End Select
            End If
        End If
    Next i
    result = "老师您好，已完成《" & wb.Name & "》的本次整理，汇报如下：" & vbCrLf & _
        "统计日期：" & Format$(Date, "yyyy-mm-dd") & "（以本机今日成功操作日志为准）" & vbCrLf
    If actions = 0 Then result = result & "今日尚无可汇总的成功修改记录。" & vbCrLf
    For Each key In counts.Keys
        result = result & "· " & key & "：" & runs(key) & "次，累计数量" & counts(key) & QuantityUnit(CStr(key)) & "。" & vbCrLf
    Next key
    result = result & "说明：以上为累计操作量，重复操作可能重复计数，并非唯一单元格数或全量质检结论。" & vbCrLf & _
        "文件未由插件自动保存，请复核后保存并发送；公式风险和业务结果仍需人工确认。" & vbCrLf & _
        "基线备份保留原始数据；脱敏仅针对所选区域，不代表全文件或历史备份已经去敏。"
    DH_BuildReceipt = result
End Function

Private Function QuantityUnit(ByVal operation As String) As String
    Select Case operation
        Case "删除空行": QuantityUnit = "条空白记录（分区域计数）"
        Case "删除重复行", "高亮重复项", "多表纵向追加", "宽表转长表", "按列拆表": QuantityUnit = "行"
        Case "拆分合并": QuantityUnit = "个合并区域"
        Case "一键交表礼仪": QuantityUnit = "张次工作表"
        Case "清除临时标记": QuantityUnit = "项（报告/规则/旧底色）"
        Case Else: QuantityUnit = "个单元格"
    End Select
End Function

Private Function TryDeliverySheet(ByVal ws As Worksheet, ByRef message As String) As Boolean
    On Error GoTo Failed
    DH_DeliveryEtiquette ws
    TryDeliverySheet = True
    Exit Function
Failed:
    message = Err.Description
End Function

Public Function DH_DeliveryWorkbook(ByVal wb As Workbook) As Long
    Dim original As Worksheet, selected As Range, ws As Worksheet, targets As New Collection, details As New Collection
    Dim state As DH_AppState, frozen As Boolean, completed As Long, failed As Long, skipped As Long
    Dim backup As String, audit As String, message As String, code As Long, reason As String
    On Error GoTo Failed
    DH_RequireWorkbook wb
    If Application.ActiveWindow Is Nothing Then Err.Raise vbObjectError + 893, , "没有活动窗口。"
    If Application.ActiveWindow.SelectedSheets.Count <> 1 Then Err.Raise vbObjectError + 893, , "请先取消工作表分组。"
    If TypeOf Application.ActiveSheet Is Worksheet Then
        If Application.ActiveSheet.Parent Is wb Then
            Set original = Application.ActiveSheet: Set selected = DH_OriginalSelection(original)
        End If
    End If
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            If Not DH_IsReportSheet(ws) And Not DH_IsPreviousReport(ws) Then targets.Add ws
        End If
    Next ws
    skipped = wb.Sheets.Count - targets.Count
    If targets.Count = 0 Then
        DH_LastSummary = "没有可处理的可见业务表；隐藏表、插件报告和图表工作表跳过。"
        DH_Log "交表礼仪汇总", wb, "", "全簿可见业务表", 0, "成功；无修改；" & DH_LastSummary
        Exit Function
    End If
    ' View settings only: keep local view rollback, never create a data baseline.
    DH_Freeze state: frozen = True
    For Each ws In targets
        reason = ""
        If TryDeliverySheet(ws, reason) Then
            completed = completed + 1
            details.Add Array(ws.Name, "交表礼仪", "成功", "视图设置", "A1；100%；冻结首行；网格线开启", "")
        Else
            failed = failed + 1
            details.Add Array(ws.Name, "交表礼仪", "失败；" & reason, "视图设置", "本表已尝试恢复；其他表继续处理", "")
        End If
    Next ws
    audit = DH_SaveAudit("交表礼仪汇总", wb, details)
    DH_LastSummary = "已整理=" & completed & "张；跳过=" & skipped & "张；失败=" & failed & "张；明细=" & audit
    ' Individual successful sheet records feed the receipt. The batch record has
    ' a separate operation name to avoid counting the same sheets a second time.
    DH_Log "交表礼仪汇总", wb, "", "全簿可见业务表", completed, _
        IIf(failed = 0, "成功；", "部分完成；") & DH_LastSummary & "；未自动保存", backup
    If Not original Is Nothing Then
        If DH_IsReportSheet(original) Then
            DH_RestoreFocus original, selected
        Else
            original.Activate
        End If
    End If
    DH_DeliveryWorkbook = completed
    DH_Restore state
    Exit Function
Failed:
    message = Err.Description: code = Err.Number
    DH_RestoreFocus original, selected
    If frozen Then DH_Restore state
    If completed > 0 Then message = message & "；已整理" & completed & "张，其他已完成视图设置保留，可手工调整视图；本操作不创建数据基线。"
    DH_LogFailure "交表礼仪汇总", wb, "", "", message, backup
    Err.Raise code, "电子羊表格", message
End Function

Public Sub DH_DeliveryEtiquette(ByVal ws As Worksheet)
    Dim wb As Workbook, window As Window, state As DH_AppState, frozen As Boolean, backup As String
    Dim oldZoom As Variant, oldGrid As Boolean, oldFreeze As Boolean, oldRow As Long, oldColumn As Long, oldScrollRow As Long, oldScrollColumn As Long
    Dim original As Range, message As String, code As Long
    On Error GoTo Failed
    If ws Is Nothing Then Err.Raise vbObjectError + 893, , "请选择业务工作表。"
    Set wb = ws.Parent: DH_RequireWorkbook wb
    If DH_IsReportSheet(ws) Then Err.Raise vbObjectError + 893, , "请回到业务表后使用交表礼仪。"
    If Application.ActiveWindow Is Nothing Then Err.Raise vbObjectError + 893, , "没有活动窗口。"
    If Application.ActiveWindow.SelectedSheets.Count <> 1 Then Err.Raise vbObjectError + 893, , "请先取消工作表分组。"
    ws.Activate: Set original = DH_OriginalSelection(ws): Set window = Application.ActiveWindow
    oldZoom = window.Zoom: oldGrid = window.DisplayGridlines: oldFreeze = window.FreezePanes
    oldRow = window.SplitRow: oldColumn = window.SplitColumn: oldScrollRow = window.ScrollRow: oldScrollColumn = window.ScrollColumn
    ' No values, formulas or worksheet structure are changed by view etiquette.
    DH_Freeze state: frozen = True
    window.FreezePanes = False: window.SplitRow = 1: window.SplitColumn = 0: window.FreezePanes = True
    window.Zoom = 100: window.DisplayGridlines = True
    window.ScrollRow = 1: window.ScrollColumn = 1
    ws.Range("A1").Select
    DH_Log "一键交表礼仪", wb, ws.Name, "A1；缩放100%；冻结首行；网格线开启", 1, "成功；仅当前窗口当前表；仅视图不创建基线；未自动保存", backup
    DH_Restore state
    Exit Sub
Failed:
    message = Err.Description: code = Err.Number
    On Error Resume Next
    If frozen Then
        window.FreezePanes = False: window.SplitRow = oldRow: window.SplitColumn = oldColumn: window.FreezePanes = oldFreeze
        window.Zoom = oldZoom: window.DisplayGridlines = oldGrid
        window.ScrollRow = oldScrollRow: window.ScrollColumn = oldScrollColumn
        DH_RestoreFocus ws, original: DH_Restore state
    End If
    On Error GoTo 0
    DH_LogFailure "一键交表礼仪", wb, "", "", message, backup
    Err.Raise code, "电子羊表格", message
End Sub
