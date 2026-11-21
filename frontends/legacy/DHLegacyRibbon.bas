Attribute VB_Name = "DHLegacyRibbon"
Option Explicit
Private licenseRibbon As IRibbonUI
Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal owner As LongPtr, ByVal text As LongPtr, ByVal title As LongPtr, ByVal flags As Long) As Long

Private Function DH1_Message(ByVal text As String, Optional ByVal flags As Long = vbOKOnly, Optional ByVal title As String = "") As Long
    If Len(title) = 0 Then title = "电子羊表格"
    DH1_Message = MessageBoxW(Application.Hwnd, StrPtr(text), StrPtr(title), flags)
End Function

' Presentation/parameter adapter only. Every write uses the same locked plan
' and final execution boundary as the Task Pane. No cleaning algorithm here.
Public Function DH1_ProductVersion() As String
    DH1_ProductVersion = "1.0.0"
End Function

Public Function DH1_Health() As String
    DH1_Health = "DataHelper/1.0.0/shared-engine"
End Function

Public Sub DH1_LegacyVisible(ByVal control As IRibbonControl, ByRef visible)
    Dim mode As String
    On Error Resume Next
    mode = CreateObject("WScript.Shell").RegRead("HKCU\Software\DataHelper\1.0\Frontend")
    On Error GoTo 0
    If Len(Environ$("DH1_FRONTEND")) > 0 Then mode = Environ$("DH1_FRONTEND")
    visible = (mode <> "taskpane")
End Sub

Private Function CallEngine(ByVal command As String, Optional ByVal argument As Variant) As Variant
    Dim envelope As Variant
    If IsMissing(argument) Then envelope = DHPoc2_Invoke(command) Else envelope = DHPoc2_Invoke(command, argument)
    If CStr(envelope(0)) <> "ok" Then Err.Raise CLng(envelope(1)), "电子羊表格", CStr(envelope(2))
    CallEngine = envelope(1)
End Function

Private Sub RequireBusinessSheet(ByVal ws As Worksheet)
    DH_RequireWorkbook ws.Parent
    If DH_IsReportSheet(ws) Or DH_IsPreviousReport(ws) Then Err.Raise 5, , "请选择业务工作表。"
End Sub

Private Function PickColumn(ByVal prompt As String) As Range
    Dim picked As Range
    On Error Resume Next
    Set picked = Application.InputBox(prompt, "电子羊表格", Type:=8)
    On Error GoTo 0
    If picked Is Nothing Then Exit Function
    If picked.Areas.Count <> 1 Or picked.Columns.Count <> 1 Then Err.Raise 5, , "请选择单列中的连续区域。"
    Set PickColumn = picked
End Function

Private Function ColumnLetter(ByVal column As Long) As String
    ColumnLetter = Split(Cells(1, column).Address, "$")(1)
End Function

Private Function LookupAction(ByVal targetSheet As Worksheet) As String
    Dim sourceKey As Range, sourceValue As Range, targetKey As Range, targetValue As Range
    Set sourceKey = PickColumn("请选择来源表匹配键列；取消则不执行。")
    If sourceKey Is Nothing Then Exit Function
    Set sourceValue = PickColumn("请选择同一来源表的返回字段列。")
    If sourceValue Is Nothing Then Exit Function
    Set targetKey = PickColumn("请选择目标业务表的匹配键列。")
    If targetKey Is Nothing Then Exit Function
    Set targetValue = PickColumn("请选择目标业务表的结果列；已有内容和公式不会覆盖。")
    If targetValue Is Nothing Then Exit Function
    If Not sourceKey.Parent Is sourceValue.Parent Then Err.Raise 5, , "来源键与返回字段必须在同一张表。"
    If Not sourceKey.Parent.Parent Is targetSheet.Parent Then Err.Raise 5, , "当前共享接口支持同一工作簿内的两表补全。"
    If Not targetKey.Parent Is targetSheet Or Not targetValue.Parent Is targetSheet Then Err.Raise 5, , "目标列必须属于最初的业务表。"
    LookupAction = "Lookup:" & sourceKey.Parent.Name & ":" & ColumnLetter(sourceKey.Column) & ":" & ColumnLetter(sourceValue.Column) & ":" & ColumnLetter(targetKey.Column) & ":" & ColumnLetter(targetValue.Column) & ":abort:static"
End Function

Private Function ReconcileAction(ByVal targetSheet As Worksheet) As String
    Dim other As Range
    Set other = PickColumn("请选择同一工作簿内核对表的比对列（不含表头）；当前选区为主表比对数据。")
    If other Is Nothing Then Exit Function
    If Not other.Parent.Parent Is targetSheet.Parent Then Err.Raise 5, , "当前共享接口支持同一工作簿内核对。"
    ReconcileAction = "Reconcile:" & other.Parent.Name & ":" & ColumnLetter(other.Column) & ":false:" & CStr(other.Row) & "," & CStr(other.Row + other.Rows.Count - 1)
End Function

Private Function PromptField(ByVal prompt As String, ByVal defaultValue As String, ByRef cancelled As Boolean) As String
    Dim answer As Variant
    answer = Application.InputBox(prompt, "电子羊表格参数", defaultValue, Type:=2)
    If VarType(answer) = vbBoolean Then
        cancelled = True
    Else
        PromptField = CStr(answer)
    End If
End Function

' Parameter collection only. The shared preview/prepare boundary validates every field.
Private Function NewTableAction(ByVal kind As String, ByVal ws As Worksheet) As String
    Dim prompts As Variant, defaults As Variant, fields() As String, i As Long, cancelled As Boolean
    Dim names As String, firstTwo As String, sheet As Worksheet, n As Long
    Select Case kind
        Case "Append"
            For Each sheet In ws.Parent.Worksheets
                If sheet.Visible = xlSheetVisible And Not DH_IsReportSheet(sheet) And Not DH_IsPreviousReport(sheet) Then
                    names = names & sheet.Name & "、": n = n + 1
                    If n <= 2 Then
                        If Len(firstTwo) > 0 Then firstTwo = firstTwo & ","
                        firstTwo = firstTwo & sheet.Name
                    End If
                End If
            Next sheet
            prompts = Array("来源表名，以英文逗号分隔（至少2张）：" & names, "列模式：union保留全部列 / common仅共同列。不同名列默认不会合并。", "小计行：exclude排除 / keep保留", "来源标记：sheet表名 / sheet_row表名行号 / none无", "新建汇总表名称（原表不改动）", "同义列显式映射，可留空。例：产品编码=商品编号;单价=售价；不自动猜测同义词")
            defaults = Array(firstTwo, "union", "exclude", "sheet", "多表追加汇总", "")
        Case "Unpivot"
            prompts = Array("固定维度列，英文逗号分隔（如A,B）", "需要展开的指标列（如C,D,E）", "新表属性列表头", "新表数值列表头", "空值规则：skip_blank / skip_blank_zero / keep_all", "小计行：exclude排除 / keep保留", "新建明细表名称（原表不改动）")
            defaults = Array("A", "B,C", "项目", "数值", "skip_blank", "exclude", "宽表转长表明细")
        Case "SplitSheet"
            prompts = Array("用于分组的列（如B）", "空值分组：bucket单独一组 / skip跳过 / abort阻止", "小计行：exclude排除 / keep保留", "新分表名前缀（原表不改动）")
            defaults = Array("A", "bucket", "exclude", "分表_")
        Case Else: Err.Raise 5, , "未知参数类型。"
    End Select
    ReDim fields(0 To UBound(prompts))
    For i = 0 To UBound(prompts)
        fields(i) = PromptField(CStr(prompts(i)), CStr(defaults(i)), cancelled)
        If cancelled Then Exit Function
    Next i
    NewTableAction = kind & ":" & Join(fields, ":")
End Function

' Fixed read-only preview adapter; indices are the published existing wire contract.
Public Function DH1_PreviewAction(ByVal action As String) As Variant
    Select Case Split(action, ":")(0)
        Case "Lookup": DH1_PreviewAction = CallEngine("DHPoc2_LookupPreview", action)
        Case "Append": DH1_PreviewAction = CallEngine("DHPoc2_AppendPreview", action)
        Case "Unpivot": DH1_PreviewAction = CallEngine("DHPoc2_UnpivotPreview", action)
        Case "SplitSheet": DH1_PreviewAction = CallEngine("DHPoc2_SplitSheetPreview", action)
        Case Else: Err.Raise 5, , "未支持的预览；没有执行。"
    End Select
End Function

Private Function PreviewText(ByVal action As String, ByVal data As Variant) As String
    Select Case Split(action, ":")(0)
        Case "Lookup"
            If Not CBool(data(16)) Then Err.Raise 5, , CStr(data(17))
            PreviewText = "跨表补全预检通过；已有内容、公式和错误值按保护规则处理。"
        Case "Append"
            If Not CBool(data(11)) Then Err.Raise 5, , CStr(data(12))
            PreviewText = "预计追加 " & CStr(data(2)) & " 行；列名不一致时请复核显式映射。"
        Case "Unpivot"
            If Not CBool(data(14)) Then Err.Raise 5, , CStr(data(15))
            PreviewText = "宽转长预检通过；新建明细表，原表保留。"
        Case "SplitSheet"
            If Not CBool(data(13)) Then Err.Raise 5, , CStr(data(14))
            PreviewText = "按列拆表预检通过；新建分表，原表保留。"
    End Select
End Function

Public Function DH1_RunAction(ByVal action As String) As Variant
    Dim ws As Worksheet, target As Range, scan As Variant, plan As Variant, confirm As Variant, result As Variant, preview As Variant, previewNote As String, kind As String
    If Not TypeOf Application.ActiveSheet Is Worksheet Or Not TypeOf Application.Selection Is Range Then Err.Raise 5, , "请先选择业务表中的数据区域。"
    Set ws = Application.ActiveSheet: Set target = Application.Selection
    RequireBusinessSheet ws
    DHL_RequirePlan action
    Select Case action
        Case "Space", "Unmerge", "Unit", "Split", "Date", "BlankRows", "Mark", "Dedup", "Transform:Phone", "Transform:ID", "Transform:Name", "Delivery", "ClearTemporary"
        Case "Lookup"
            action = LookupAction(ws)
            If Len(action) = 0 Then DHPoc2_Invalidate "legacy parameter selection cancelled": DH1_RunAction = Array("CANCELLED"): Exit Function
            Set target = DH_ContentRange(ws)
            If target Is Nothing Then Err.Raise 5, , "目标业务表没有数据。"
        Case "Append", "Unpivot", "SplitSheet"
            action = NewTableAction(action, ws)
            If Len(action) = 0 Then DHPoc2_Invalidate "legacy parameter selection cancelled": DH1_RunAction = Array("CANCELLED"): Exit Function
            Set target = DH_ContentRange(ws)
            If target Is Nothing Then Err.Raise 5, , "当前业务表没有数据。"
        Case "Reconcile"
            action = ReconcileAction(ws)
            If Len(action) = 0 Then DHPoc2_Invalidate "legacy parameter selection cancelled": DH1_RunAction = Array("CANCELLED"): Exit Function
        Case Else
            kind = Split(action, ":")(0)
            If kind <> "Append" And kind <> "Unpivot" And kind <> "SplitSheet" And kind <> "Lookup" Then Err.Raise 5, , "未支持的操作；未执行。"
            Set target = DH_ContentRange(ws)
            If target Is Nothing Then Err.Raise 5, , "当前业务表没有数据。"
    End Select
    ws.Activate: target.Select
    scan = CallEngine("DHPoc2_Scan")
    kind = Split(action, ":")(0)
    If kind = "Lookup" Or kind = "Append" Or kind = "Unpivot" Or kind = "SplitSheet" Then
        preview = DH1_PreviewAction(action)
        previewNote = PreviewText(action, preview) & vbCrLf
    End If
    plan = CallEngine("DHPoc2_Prepare", action)
    confirm = CallEngine("DHPoc2_ConfirmationTarget")
    If DH1_Message(previewNote & "即将处理：" & CStr(confirm(1)) & vbCrLf & CStr(confirm(2)) & " / " & CStr(confirm(3)) & vbCrLf & "操作：" & CStr(confirm(4)) & vbCrLf & "修改前核验当日基线；基线不是逐步撤销。是否继续？", vbYesNo + vbExclamation + vbDefaultButton2, "电子羊表格 · 写入确认") <> vbYes Then
        DHPoc2_Invalidate "legacy native confirmation cancelled"
        DH1_RunAction = Array("CANCELLED")
        Exit Function
    End If
    result = CallEngine("DHPoc2_Confirm", CStr(confirm(0)))
    DH1_RunAction = CallEngine("DHPoc2_Execute")
End Function

Public Sub DH1_OnAction(ByVal control As IRibbonControl)
    Dim action As String, result As Variant, rows As Variant, row As Variant, message As String, failure As String
    On Error GoTo Failed
    Select Case control.ID
        Case "dh1Space": action = "Space"
        Case "dh1Unmerge": action = "Unmerge"
        Case "dh1Unit": action = "Unit"
        Case "dh1Split": action = "Split"
        Case "dh1Date": action = "Date"
        Case "dh1Blank": action = "BlankRows"
        Case "dh1Mark": action = "Mark"
        Case "dh1Dedup": action = "Dedup"
        Case "dh1Phone": action = "Transform:Phone"
        Case "dh1ID": action = "Transform:ID"
        Case "dh1Name": action = "Transform:Name"
        Case "dh1Lookup": action = "Lookup"
        Case "dh1Append": action = "Append"
        Case "dh1Unpivot": action = "Unpivot"
        Case "dh1SplitSheet": action = "SplitSheet"
        Case "dh1Reconcile": action = "Reconcile"
        Case "dh1Delivery": action = "Delivery"
        Case "dh1Clear": action = "ClearTemporary"
        Case Else: Err.Raise 5, , "未知按钮。"
    End Select
    result = DH1_RunAction(action)
    If CStr(result(0)) = "CANCELLED" Then Exit Sub
    If CStr(result(0)) <> "COMPLETED" Then
        DH1_Message "操作未全部完成。请检查已处理数据、审计和当日基线；不要直接重复执行。", vbExclamation, "电子羊表格": Exit Sub
    End If
    DH1_RefreshLicense
    rows = result(3)
    For Each row In rows: message = message & CStr(row(0)) & "：" & CStr(row(1)) & DHUI_CountUnit(CStr(row(0))) & vbCrLf: Next row
    If Len(CStr(result(5))) > 0 Then message = message & CStr(result(5)) & vbCrLf
    DH1_Message message & "文件未自动保存；详细结果已写入本地审计。", vbInformation, "电子羊表格"
    Exit Sub
Failed:
    failure = Err.Description
    DHPoc2_Invalidate "legacy command failed"
    DH1_Message failure, vbExclamation, "电子羊表格"
End Sub

Public Sub DH1_OnCheck(ByVal control As IRibbonControl)
    Dim data As Variant, rows As Variant, message As String
    On Error GoTo Failed
    If control.ID = "dh1Workbook" Then
        rows = CallEngine("DHPoc2_Workbook")
        For Each data In rows: message = message & CStr(data(1)) & "：隐形字符" & CStr(data(3)) & "，合并区域" & CStr(data(6)) & "，伪数字" & CStr(data(5)) & "，日期候选" & CStr(data(12)) & vbCrLf: Next data
    Else
        data = CallEngine("DHPoc2_Scan")
        message = CStr(data(1)) & " / " & CStr(data(2)) & vbCrLf & "隐形字符：" & CStr(data(3)) & vbCrLf & "合并区域：" & CStr(data(6)) & vbCrLf & "伪数字：" & CStr(data(5)) & vbCrLf & "空行：" & CStr(data(8)) & vbCrLf & "重复记录：" & CStr(data(9)) & vbCrLf & "公式风险：" & CStr(data(11)) & vbCrLf & "明确日期文本候选：" & CStr(data(12))
    End If
    DH1_Message message, vbInformation, "电子羊表格 · 体检"
    Exit Sub
Failed:
    DH1_Message Err.Description, vbExclamation, "电子羊表格"
End Sub

Public Sub DH1_OnBackup(ByVal control As IRibbonControl)
    DH_OpenBackupFolder
End Sub

Public Sub DH1_OnReport(ByVal control As IRibbonControl)
    Dim reportName As String
    On Error GoTo Failed
    If Not TypeOf Application.ActiveSheet Is Worksheet Then Err.Raise 5, , "请选择业务工作表。"
    RequireBusinessSheet Application.ActiveSheet
    reportName = DH_CheckWorksheet(Application.ActiveSheet)
    DHPoc2_Invalidate "legacy report refreshed"
    DH1_Message "报告已在后台生成：" & reportName, vbInformation, "电子羊表格"
    Exit Sub
Failed:
    DH1_Message Err.Description, vbExclamation, "电子羊表格"
End Sub

Public Sub DH1_OnReceipt(ByVal control As IRibbonControl)
    On Error GoTo Failed
    DHReceiptForm.Caption = "电子羊表格 · 交差小票"
    DHReceiptForm.Show
    Exit Sub
Failed:
    DH1_Message Err.Description, vbExclamation, "电子羊表格"
End Sub

Public Sub DH1_OnLoad(ByVal ribbon As IRibbonUI)
    Set licenseRibbon = ribbon
End Sub

Public Sub DH1_RefreshLicense()
    On Error Resume Next
    If Not licenseRibbon Is Nothing Then licenseRibbon.Invalidate
End Sub

Public Sub DH1_LicenseLabel(ByVal control As IRibbonControl, ByRef label)
    On Error GoTo Unavailable
    label = DHL_StatusText()
    Exit Sub
Unavailable:
    label = "状态暂不可用"
End Sub

Public Sub DH1_OnLicense(ByVal control As IRibbonControl)
    DH1_Message "电子羊表格 1.0 开源版（MIT License）" & vbCrLf & _
                "全部 25 个表格整理、核对、脱敏与安全基线备份模块已完全解锁，无任何试用次数限制。", vbInformation, "电子羊表格"
    DH1_RefreshLicense
End Sub
