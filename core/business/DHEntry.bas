Attribute VB_Name = "DHEntry"
Option Explicit
Private busy As Boolean

Private Function DH_ErrorText(ByVal message As String) As String
    DH_ErrorText = message
    If Len(DH_LastLogError) > 0 Then DH_ErrorText = message & vbCrLf & DH_LastLogError
End Function

Public Sub DH_OnDates(ByVal control As IRibbonControl)
    DH_RunV04 1
End Sub
Public Sub DH_OnMask(ByVal control As IRibbonControl)
    DH_RunV04 2
End Sub
Public Sub DH_OnDuplicateMark(ByVal control As IRibbonControl)
    DH_RunV04 3
End Sub
Public Sub DH_OnDuplicateDelete(ByVal control As IRibbonControl)
    DH_RunV04 4
End Sub
Public Sub DH_OnClearTemporary(ByVal control As IRibbonControl)
    DH_RunV04 5
End Sub
Public Sub DH_OnEtiquette(ByVal control As IRibbonControl)
    DH_RunV04 6
End Sub
Public Sub DH_OnReceipt(ByVal control As IRibbonControl)
    DH_RunV04 7
End Sub

Private Sub DH_RunV04(ByVal action As Long)
    Dim wb As Workbook, ws As Worksheet, selected As Range, count As Long, mode As Variant, operation As String
    Dim header As VbMsgBoxResult, table As ListObject, rng As Range, plan As Collection, details As New Collection, protected As Long
    Dim message As String
    If busy Then Exit Sub
    busy = True
    On Error GoTo Failed
    Set wb = Application.ActiveWorkbook: DH_RequireWorkbook wb
    If Not TypeOf Application.ActiveSheet Is Worksheet Then Err.Raise vbObjectError + 900, , "请选择业务工作表。"
    Set ws = Application.ActiveSheet
    Select Case action
        Case 1, 2, 3, 4
            If Not TypeOf Application.Selection Is Range Then Err.Raise vbObjectError + 900, , "请先选择数据区域。"
            Set selected = Application.Selection
            If action = 1 Then
                DH_RequireEntrySelection selected
            Else
                DH_RequireEntrySelection selected, True, action = 2
            End If
            If action = 1 Then
                operation = "日期规范化"
                If MsgBox(DH_SelectionTarget(selected) & "仅将明确的年在前日期文本转为真实日期，显示YYYY-MM-DD。" & vbCrLf & _
                    "8位数字文本可能是编号，请只选择日期列；原有数值和公式保留。" & vbCrLf & _
                    "先确保当日基线，不能逐步撤销。继续？", vbOKCancel + vbQuestion + vbDefaultButton2, operation) <> vbOK Then GoTo Cancelled
                count = DH_TransformSelection(selected, 0)
            ElseIf action = 2 Then
                operation = "敏感数据脱敏"
                mode = Application.InputBox(DH_SelectionTarget(selected) & "请输入模式：1=11位手机号；2=18位身份证文本；3=中文姓名。" & vbCrLf & _
                    "姓名仅保留姓氏，其他字全部打星；姓名模式只用于已确认的姓名列。", operation, 1, Type:=1)
                If VarType(mode) = vbBoolean Then GoTo Cancelled
                If mode <> Fix(mode) Or mode < 1 Or mode > 3 Then Err.Raise vbObjectError + 901, , "只支持模式1、2、3。"
                If MsgBox(DH_SelectionTarget(selected) & "仅对选区中符合所选类型的常量脱敏，保留文本类型；公式跳过。" & vbCrLf & _
                    "第1行及超级表表头/汇总行自动保护；仅处理数据行。" & vbCrLf & _
                    "当日基线仍保留原始数据，其他区域和旧日志不会自动去敏。请复核后交付。继续？", _
                    vbOKCancel + vbQuestion + vbDefaultButton2, operation) <> vbOK Then GoTo Cancelled
                count = DH_TransformSelection(selected, CLng(mode))
            Else
                operation = IIf(action = 4, "删除重复行", "高亮重复项")
                header = MsgBox(DH_SelectionTarget(selected) & "请覆盖全部数据列。记录须值、类型、大小写及公式一致才视为完全重复。" & vbCrLf & _
                    "普通选区首行是否为表头？是=跳过首行；否=全部参与；取消=终止。Excel表格表头自动排除。", _
                    vbYesNoCancel + vbQuestion + vbDefaultButton3, operation)
                If header = vbCancel Then GoTo Cancelled
                Set plan = DH_DuplicatePlan(selected, action = 4, header = vbYes, rng, table, details, protected)
                count = DH_BlankRowCount(plan)
                message = "计划处理 " & count & " 行重复记录，保留首次出现。"
                If action = 4 Then message = message & vbCrLf & "记录将删除，下方数据上移；含公式/备注/批注的重复行保留；直接依赖待删记录的公式会阻止删除。请先取消筛选和隐藏行。"
                If MsgBox(DH_SelectionTarget(selected) & message & vbCrLf & "先确保当日基线，不能逐步撤销。继续？", _
                    vbOKCancel + vbQuestion + vbDefaultButton2, operation) <> vbOK Then GoTo Cancelled
                count = DH_HandleDuplicates(selected, action = 4, header = vbYes)
            End If
            MsgBox operation & "完成，数量=" & count & "；明细已存入本地日志。" & vbCrLf & DH_LastSummary, vbInformation, "电子羊表格"
        Case 5
            operation = "清除临时标记"
            DH_RequireWorkbook wb, True
            If MsgBox(DH_WorkbookTarget(wb, "全簿插件报告及对账/重复项标记") & _
                "删除插件体检及旧临时报告，并移除插件对账/重复项标记。" & vbCrLf & _
                "用户自建同名表和普通底色保留；旧版底色仅按可确认历史明细清除。" & vbCrLf & _
                "先确保当日基线，不能逐步撤销。继续？", vbOKCancel + vbQuestion + vbDefaultButton2, operation) <> vbOK Then GoTo Cancelled
            count = DH_ClearTemporary(wb)
            MsgBox "临时标记清除完成。" & vbCrLf & DH_LastSummary & vbCrLf & "原文件未自动保存。", vbInformation, "电子羊表格"
        Case 6
            operation = "一键交表礼仪"
            count = DH_DeliveryWorkbook(wb)
            MsgBox "工作簿《" & wb.Name & "》全簿交表礼仪完成。" & vbCrLf & DH_LastSummary & vbCrLf & _
                "各业务表回A1、缩放100%、冻结首行并显示网格线；请复核失败项并保存。", vbInformation, "电子羊表格"
        Case 7
            operation = "生成交差小票文案"
            DHReceiptForm.Show
            DH_Log operation, wb, ws.Name, "", 0, "成功；显示本地文案，不自动发送"
    End Select
    busy = False
    Exit Sub
Cancelled:
    DH_Log operation, wb, ws.Name, "", 0, "用户取消"
    busy = False
    Exit Sub
Failed:
    message = Err.Description
    DH_LogFailure operation, wb, "", "", message
    MsgBox DH_ErrorText(message), vbExclamation, "电子羊表格"
    busy = False
End Sub

Public Sub DH_OnCheck(ByVal control As IRibbonControl)
    DH_CheckActive
End Sub
Public Sub DH_OnCheckWorkbook(ByVal control As IRibbonControl)
    DH_CheckAll
End Sub
Public Sub DH_OnSplitUnits(ByVal control As IRibbonControl)
    DH_SplitUnitsSelection
End Sub
Public Sub DH_OnReconcile(ByVal control As IRibbonControl)
    DH_ReconcileSelection
End Sub
Public Sub DH_OnClean(ByVal control As IRibbonControl)
    DH_CleanSelection
End Sub
Public Sub DH_OnReport(ByVal control As IRibbonControl)
    DH_ReportActive
End Sub
Public Sub DH_OnDeleteBlankRows(ByVal control As IRibbonControl)
    DH_DeleteBlankSelection
End Sub
Public Sub DH_OnUnmerge(ByVal control As IRibbonControl)
    DH_RunNewAction False
End Sub
Public Sub DH_OnNumbers(ByVal control As IRibbonControl)
    DH_RunNewAction True
End Sub
Public Sub DH_OnOpenBackup(ByVal control As IRibbonControl)
    DH_OpenBackupFolder
End Sub

Private Function DH_ConfirmMergeAndContinue(ByVal selected As Range, ByVal operation As String, _
    ByRef hadMerges As Boolean, ByRef coreStarted As Boolean, Optional ByVal plannedDeleteCount As Long = -1, _
    Optional ByVal plannedMode As String = "") As Boolean
    Dim merges As Collection, remaining As Collection, reason As String, prompt As String, changed As Long, choice As VbMsgBoxResult
    Set merges = DH_MergedAreas(selected)
    hadMerges = (merges.Count > 0)
    If Not hadMerges Then
        DH_ConfirmMergeAndContinue = True
        Exit Function
    End If
    DH_RequireSelection selected
    If Not DH_CanAutoUnmerge(selected, reason) Then
        Err.Raise vbObjectError + 761, "电子羊表格", reason & "。请使用“拆分合并”查看明细，或调整选区。"
    End If
    prompt = DH_SelectionTarget(selected) & "检测到选区包含合并单元格，直接操作可能导致数据错位。" & vbCrLf & _
        "是否立即为您“拆分合并并自动向下填满”，然后继续" & operation & "？" & vbCrLf & _
        "仅完整、单列纵向且无公式或内容冲突的合并区域可一键处理。" & vbCrLf & _
        "首次破坏操作前建立当日基线；不能逐步撤销。"
    If operation = "删除空行" Then prompt = prompt & vbCrLf & plannedMode & vbCrLf & "清理范围内的后方单元格会上移，公式引用可能调整。"
    If operation = "删除空行" Then
        If plannedDeleteCount >= 0 Then prompt = prompt & vbCrLf & "当前预估 " & plannedDeleteCount & " 条空白记录；拆分后重新计算。"
        prompt = prompt & vbCrLf & "是：先拆分，再进入独立的删除确认。" & vbCrLf & _
            "否：仅拆分，暂不删除空行。" & vbCrLf & "取消：两步均不执行。"
        choice = MsgBox(prompt, vbYesNoCancel + vbQuestion + vbDefaultButton3, operation)
    Else
        choice = MsgBox(prompt, vbYesNo + vbQuestion + vbDefaultButton2, operation)
    End If
    If choice = vbCancel Or (choice = vbNo And operation <> "删除空行") Then
        DH_Log operation, selected.Parent.Parent, selected.Parent.Name, selected.Address(False, False), 0, "用户取消一键拆分"
        Exit Function
    End If
    coreStarted = True
    changed = DH_UnmergeFill(selected)
    Set remaining = DH_MergedAreas(selected)
    If remaining.Count > 0 Then
        Err.Raise vbObjectError + 762, "电子羊表格", "仍有合并区域未解除，已停止继续" & operation & "。跳过原因见本地明细日志。"
    End If
    If choice = vbNo Then
        DH_Log operation, selected.Parent.Parent, selected.Parent.Name, selected.Address(False, False), 0, "用户选择仅拆分填充；未执行删除"
        MsgBox "已拆分并填充 " & changed & " 个合并区域；未执行删除空行。原文件未自动保存。", vbInformation, "电子羊表格"
        Exit Function
    End If
    DH_ConfirmMergeAndContinue = True
End Function

Public Sub DH_OpenBackupFolder()
    Dim wb As Workbook, root As String, folder As String
    Dim message As String, explorer As String, processId As Double
    If busy Then Exit Sub
    busy = True
    On Error GoTo Failed
    root = DH_BackupRoot()
    folder = root
    DH_EnsureFolder root
    Set wb = Application.ActiveWorkbook
    If Not wb Is Nothing Then
        If wb Is ThisWorkbook Or wb.IsAddin Then Set wb = Nothing
    End If
    folder = DH_BackupFolderPath(wb)
    explorer = Environ$("WINDIR") & "\explorer.exe"
    processId = Shell(Chr$(34) & explorer & Chr$(34) & " " & Chr$(34) & folder & Chr$(34), vbNormalFocus)
    If StrComp(folder, root, vbTextCompare) = 0 Then
        message = "已打开备份总目录。当前工作簿今天尚无基线备份；首次破坏性操作时会自动创建。"
    Else
        message = "已打开当前工作簿今天的备份目录。"
    End If
    On Error Resume Next
    If Not wb Is Nothing Then DH_Log "打开备份", wb, "", "", 0, "已调用资源管理器；" & folder
    On Error GoTo Failed
    If StrComp(folder, root, vbTextCompare) = 0 Then MsgBox message, vbInformation, "电子羊表格"
    busy = False
    Exit Sub
Failed:
    message = Err.Description
    If Not wb Is Nothing Then DH_LogFailure "打开备份", wb, "", "", message
    MsgBox "无法打开备份目录。请在资源管理器地址栏粘贴：" & vbCrLf & folder & vbCrLf & message, vbExclamation, "电子羊表格"
    busy = False
End Sub

Private Sub DH_RunNewAction(ByVal numbers As Boolean)
    Dim selected As Range, changes As Long, operation As String, message As String, coreStarted As Boolean
    If busy Then Exit Sub
    busy = True
    operation = IIf(numbers, "数字转换", "拆分合并")
    On Error GoTo Failed
    If Not TypeOf Application.Selection Is Range Then Err.Raise vbObjectError + 741, , "请先选择单元格区域。"
    Set selected = Application.Selection
    DH_RequireSelection selected
    If numbers Then
        message = "仅转换明确十进制及常见中文单位；保护前导零、长编号、日期和公式。混合单位保留。"
    Else
        message = "仅解除完整选中的单列纵向合并，并在原区域内部填充。公式和不安全形状跳过。"
    End If
    If MsgBox(DH_SelectionTarget(selected) & message & vbCrLf & "首次破坏操作前建立当日基线；当天后续复用，不能逐步撤销。继续？", _
        vbOKCancel + vbQuestion + vbDefaultButton2, operation) <> vbOK Then
        DH_Log operation, selected.Parent.Parent, selected.Parent.Name, selected.Address(False, False), 0, "用户取消"
        busy = False: Exit Sub
    End If
    coreStarted = True
    If numbers Then changes = DH_ConvertNumbers(selected) Else changes = DH_UnmergeFill(selected)
    MsgBox operation & "完成，处理数量=" & changes & "。明细已存入本地日志；原文件未自动保存。" & vbCrLf & DH_LastSummary, vbInformation, "电子羊表格"
    busy = False: Exit Sub
Failed:
    message = Err.Description
    If Not coreStarted Then DH_LogFailure operation, ActiveWorkbook, "", "", message
    MsgBox DH_ErrorText(message), vbExclamation, "电子羊表格"
    busy = False
End Sub

Public Sub DH_DeleteBlankSelection()
    Dim selected As Range, table As ListObject, plan As Collection, merges As Collection, count As Long, deleted As Long
    Dim mode As String, message As String, coreStarted As Boolean, hadMerges As Boolean, proceed As Boolean, notes As String
    If busy Then Exit Sub
    busy = True
    On Error GoTo Failed
    If Not TypeOf Application.Selection Is Range Then Err.Raise vbObjectError + 741, , "请先选择单元格区域。"
    Set selected = Application.Selection
    ' The blank-row planner partitions parallel tables and validates actual scan
    ' sizes; do not impose the whole sheet's bounding rectangle here.
    DH_RequireSelection selected.Cells(1, 1)
    Set merges = DH_MergedAreas(selected)
    hadMerges = (merges.Count > 0)
    If hadMerges Then
        Set plan = DH_PlanBlankDeletion(selected, True, notes)
        count = DH_BlankRowCount(plan)
        mode = DH_BlankPlanMode(plan) & vbCrLf & notes
    End If
    proceed = DH_ConfirmMergeAndContinue(selected, "删除空行", hadMerges, coreStarted, count, mode)
    If Not proceed Then busy = False: Exit Sub
    Set plan = DH_PlanBlankDeletion(selected, False, notes)
    count = DH_BlankRowCount(plan)
    If count = 0 Then
        If hadMerges Then
            DH_Log "删除空行", selected.Parent.Parent, selected.Parent.Name, selected.Address(False, False), 0, "已拆分合并；没有完全空白行"
            MsgBox "合并区域已安全拆分并填充；没有可清理的空白记录。含空格或公式的记录保留。" & vbCrLf & notes, vbInformation, "电子羊表格"
        Else
            DH_Log "删除空行", selected.Parent.Parent, selected.Parent.Name, selected.Address(False, False), 0, "无完全空白行，未修改"
            MsgBox "选区内没有可清理的完全空白记录。含空格或公式的记录保留。" & vbCrLf & notes, vbInformation, "电子羊表格"
        End If
        busy = False: Exit Sub
    End If
    mode = DH_BlankPlanMode(plan)
    message = ""
    If hadMerges Then message = "本次已完成拆分填充；现在取消仅取消删除，拆分结果保留。" & vbCrLf
    If MsgBox(DH_SelectionTarget(selected) & message & "将清理 " & count & " 条完全空白记录（分区域计数）。" & vbCrLf & mode & vbCrLf & notes & vbCrLf & _
        "修改前确保当日基线；不能逐步撤销；公式引用可能调整。继续？", vbOKCancel + vbQuestion + vbDefaultButton2, "安全删除空行") <> vbOK Then
        DH_Log "删除空行", selected.Parent.Parent, selected.Parent.Name, selected.Address(False, False), 0, "用户取消删除；未删除记录"
        busy = False: Exit Sub
    End If
    coreStarted = True
    deleted = DH_DeleteBlankRows(selected)
    MsgBox "已确保当日基线，清理 " & deleted & " 条空白记录（分区域计数）。" & vbCrLf & DH_LastSummary & vbCrLf & "原文件尚未自动保存。", vbInformation, "电子羊表格"
    busy = False: Exit Sub
Failed:
    message = Err.Description
    If Not coreStarted Then DH_LogFailure "删除空行", ActiveWorkbook, "", "", message
    MsgBox DH_ErrorText(message), vbExclamation, "电子羊表格"
    busy = False
End Sub

Public Sub DH_CheckActive()
    Dim reportName As String, coreStarted As Boolean, message As String
    If busy Then Exit Sub
    busy = True
    On Error GoTo Failed
    If Not TypeOf Application.ActiveSheet Is Worksheet Then Err.Raise vbObjectError + 740, , "请选择数据工作表。"
    If DH_IsReportSheet(Application.ActiveSheet) Or DH_IsPreviousReport(Application.ActiveSheet) Then _
        Err.Raise vbObjectError + 852, , "请先回到业务表，再执行体检。"
    If Not DH_ResolvePreviousReport(Application.ActiveWorkbook) Then busy = False: Exit Sub
    coreStarted = True
    reportName = DH_CheckWorksheet(Application.ActiveSheet)
    MsgBox "体检完成。只读扫描不创建备份；若先改名留存旧报告，该步骤已单独确保基线。报告：" & reportName, vbInformation, "电子羊表格"
    busy = False
    Exit Sub
Failed:
    message = Err.Description
    If Not coreStarted Then DH_LogFailure "表格体检", ActiveWorkbook, "", "", message
    MsgBox DH_ErrorText(message), vbExclamation, "电子羊表格"
    busy = False
End Sub

Public Sub DH_CleanSelection()
    Dim changes As Long, message As String, coreStarted As Boolean, hadMerges As Boolean, proceed As Boolean, selected As Range
    If busy Then Exit Sub
    busy = True
    On Error GoTo Failed
    If Not TypeOf Application.Selection Is Range Then Err.Raise vbObjectError + 741, , "请先选择单元格区域。"
    Set selected = Application.Selection
    DH_RequireEntrySelection selected, True
    proceed = DH_ConfirmMergeAndContinue(selected, "区域清理", hadMerges, coreStarted)
    If Not proceed Then busy = False: Exit Sub
    If Not hadMerges Then
        If MsgBox(DH_SelectionTarget(selected) & "将清理首尾和相邻中文间空格；英文/数字/混合词分隔保留一个空格。公式跳过。" & vbCrLf & _
            "修改前确保当日基线；同日复用，不能逐步撤销。继续清理？", vbOKCancel + vbQuestion + vbDefaultButton2, "区域清理") <> vbOK Then
            DH_Log "区域清理", selected.Parent.Parent, selected.Parent.Name, selected.Address(False, False), 0, "用户取消"
            busy = False
            Exit Sub
        End If
    End If
    coreStarted = True
    changes = DH_CleanRange(selected)
    If changes = 0 And hadMerges Then
        message = "合并区域已安全拆分并填充，已确保当日基线；文本无需进一步清理。"
    ElseIf changes = 0 Then
        message = "清理完成，无需修改，不触发备份。"
    Else
        message = "清理完成，已确保当日基线（同日复用）。修改 " & changes & " 个单元格。"
    End If
    MsgBox message & vbCrLf & "原文件尚未自动保存。", vbInformation, "电子羊表格"
    busy = False
    Exit Sub
Failed:
    message = Err.Description
    If Not coreStarted Then DH_LogFailure "区域清理", ActiveWorkbook, "", "", message
    MsgBox DH_ErrorText(message), vbExclamation, "电子羊表格"
    busy = False
End Sub

Public Sub DH_ReportActive()
    DH_RunV04 7
End Sub

Public Sub DH_CheckAll()
    Dim name As String
    If busy Then Exit Sub
    busy = True
    On Error GoTo Failed
    If Not DH_ResolvePreviousReport(Application.ActiveWorkbook) Then busy = False: Exit Sub
    name = DH_CheckWorkbook(Application.ActiveWorkbook)
    MsgBox "全簿体检完成，报告已在后台生成：" & name & vbCrLf & "请查看总览中的未完成工作表数量。", vbInformation, "电子羊表格"
    busy = False: Exit Sub
Failed:
    MsgBox DH_ErrorText(Err.Description), vbExclamation, "电子羊表格"
    busy = False
End Sub

Public Sub DH_SplitUnitsSelection()
    Dim selected As Range, count As Long
    If busy Then Exit Sub
    busy = True
    On Error GoTo Failed
    If Not TypeOf Application.Selection Is Range Then Err.Raise vbObjectError + 810, , "请先选择包含表头的单列区域。"
    Set selected = DH_ClipColumn(Application.Selection)
    If selected Is Nothing Then Err.Raise vbObjectError + 812, , "选区没有内容。"
    DH_RequireEntrySelection selected
    If MsgBox(DH_SelectionTarget(selected) & "将在所选列右侧插入整列，原列提取数值，新列保留单位。" & vbCrLf & _
        "请包含文字表头。右侧单元格会平移；将预检查询列号、动态引用和名称依赖，遇风险停止。原有公式/不明确数据保留。" & vbCrLf & _
        "操作前确保当日基线，不能逐步撤销。继续？", vbOKCancel + vbQuestion + vbDefaultButton2, "单位拆分(新列)") <> vbOK Then
        DH_Log "单位拆分", selected.Parent.Parent, selected.Parent.Name, selected.Address(False, False), 0, "用户取消"
        busy = False: Exit Sub
    End If
    count = DH_SplitUnits(selected)
    MsgBox "单位拆分完成，转换=" & count & "。明细已存入本地日志。", vbInformation, "电子羊表格"
    busy = False: Exit Sub
Failed:
    MsgBox DH_ErrorText(Err.Description), vbExclamation, "电子羊表格"
    busy = False
End Sub

Private Function AskColumn(ByVal prompt As String) As Range
    On Error Resume Next
    Set AskColumn = Application.InputBox(prompt, "两表对账", Type:=8)
    On Error GoTo 0
End Function

Public Sub DH_ReconcileSelection()
    Dim originalSheet As Worksheet, originalSelection As Range, wb As Workbook, main As Range, other As Range
    Dim choice As VbMsgBoxResult, count As Long, message As String
    If busy Then Exit Sub
    busy = True
    On Error GoTo Failed
    Set wb = Application.ActiveWorkbook: DH_RequireWorkbook wb, True
    If Not TypeOf Application.ActiveSheet Is Worksheet Then Err.Raise vbObjectError + 837, , "请选择业务工作表。"
    Set originalSheet = Application.ActiveSheet: Set originalSelection = DH_OriginalSelection(originalSheet)
    Set main = AskColumn("请选择主表的一列比对区域。稍后选择同一工作簿内另一张表的一列。")
    If main Is Nothing Then GoTo Cancelled
    Set other = AskColumn("请选择核对表的一列比对区域。支持切换工作表选择。")
    If other Is Nothing Then GoTo Cancelled
    Set main = DH_ClipColumn(main): Set other = DH_ClipColumn(other)
    If main Is Nothing Or other Is Nothing Then Err.Raise vbObjectError + 830, , "比对列没有数据。"
    DH_RequireEntrySelection main
    DH_RequireEntrySelection other
    If Not (main.Parent.Parent Is wb) Or Not (other.Parent.Parent Is wb) Then _
        Err.Raise vbObjectError + 834, , "请选择同一工作簿内的两列。"
    choice = MsgBox(DH_SelectionTarget(main, "主表") & DH_SelectionTarget(other, "核对表") & _
        "将把主表独有项标为浅绿，核对表独有项标为浅黄，使用可清除标记保留原底色。" & vbCrLf & _
        "只核对是否存在，按底层值、类型和大小写匹配；文本1与数字1视为不同，不核对重复次数；修改前确保当日基线。" & vbCrLf & _
        "两个选区的首个非空行是否都是表头？是=跳过首行；否=全部参与；取消=终止。", vbYesNoCancel + vbQuestion + vbDefaultButton3, "两表对账")
    If choice = vbCancel Then GoTo Cancelled
    count = DH_ReconcileColumns(main, other, choice = vbYes)
    DH_RestoreFocus originalSheet, originalSelection
    MsgBox "两表对账完成，标记差异单元格=" & count & "。明细已存入本地日志；清除临时标记可移除本版标色。", vbInformation, "电子羊表格"
    busy = False: Exit Sub
Cancelled:
    DH_Log "两表对账", wb, "", "", 0, "用户取消"
    DH_RestoreFocus originalSheet, originalSelection
    busy = False: Exit Sub
Failed:
    message = Err.Description
    DH_RestoreFocus originalSheet, originalSelection
    MsgBox DH_ErrorText(message), vbExclamation, "电子羊表格"
    busy = False
End Sub

' Safe build probe: proves the imported VBA project can compile and execute.
Public Function DH_Version() As String
    DH_Version = "0.4.5"
End Function
Public Function DH_BuildId() As String
    DH_BuildId = "0.4.5"
End Function

