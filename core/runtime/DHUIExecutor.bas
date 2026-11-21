Attribute VB_Name = "DHUIExecutor"
Option Explicit

' Thin adapter only. The audited business functions retain all data rules,
' baseline handling, rollback, local logging and structural preflight.
Public Function DHUI_Apply(ByVal action As String, ByVal target As Range) As Long
    Dim parts As Variant, otherSheet As Worksheet, otherRange As Range, otherCol As Long, hasHeader As Boolean, ws As Worksheet
    DHPoc2_RequireExecuting target
    If Left$(action, 7) = "Lookup:" Then
        DHUI_Apply = DH_ExecuteLookup(target.Parent, action)
        Exit Function
    End If
    If Left$(action, 7) = "Append:" Then
        DHUI_Apply = DH_ExecuteAppend(target.Parent.Parent, action)
        Exit Function
    End If
    If Left$(action, 8) = "Unpivot:" Then
        DHUI_Apply = DH_ExecuteUnpivot(target.Parent, action)
        Exit Function
    End If
    If Left$(action, 11) = "SplitSheet:" Then
        DHUI_Apply = DH_ExecuteSplitSheet(target.Parent, action)
        Exit Function
    End If
    If Left$(action, 10) = "Reconcile:" Then
        parts = Split(action, ":")
        If UBound(parts) <> 3 And UBound(parts) <> 4 Then Err.Raise vbObjectError + 961, , "两表对账参数格式无效。"
        Set otherSheet = Nothing
        For Each ws In target.Parent.Parent.Worksheets
            If StrComp(ws.Name, CStr(parts(1)), vbTextCompare) = 0 Then Set otherSheet = ws: Exit For
        Next ws
        If otherSheet Is Nothing Then Err.Raise vbObjectError + 961, , "未找到核对工作表：" & CStr(parts(1))
        otherCol = DH_ColLetterToIndex(CStr(parts(2)))
        hasHeader = (StrComp(CStr(parts(3)), "true", vbTextCompare) = 0)
        Set otherRange = DHUI_ReconcileRange(action, target.Parent.Parent)
        DHUI_Apply = DH_ReconcileColumns(target, otherRange, hasHeader)
        Exit Function
    End If
    Select Case action
        Case "Space": DHUI_Apply = DH_CleanRange(target)
        Case "Unmerge": DHUI_Apply = DH_UnmergeFill(target)
        Case "Unit": DHUI_Apply = DH_ConvertNumbers(target)
        Case "Unit:Force": DHUI_Apply = DH_ConvertNumbers(target, False, True)
        Case "Split": DHUI_Apply = DH_SplitUnits(target)
        Case "Date": DHUI_Apply = DH_TransformSelection(target, 0)
        Case "BlankRows": DHUI_Apply = DH_DeleteBlankRows(target)
        Case "Mark": DHUI_Apply = DH_HandleDuplicates(target, False, True)
        Case "Dedup": DHUI_Apply = DH_HandleDuplicates(target, True, True)
        Case "Phone": DHUI_Apply = DH_TransformSelection(target, 1)
        Case "ID": DHUI_Apply = DH_TransformSelection(target, 2)
        Case "Name": DHUI_Apply = DH_TransformSelection(target, 3)
        Case "Delivery": DHUI_Apply = DH_DeliveryWorkbook(target.Parent.Parent)
        Case "ClearTemporary": DHUI_Apply = DH_ClearTemporary(target.Parent.Parent)
        Case Else: Err.Raise vbObjectError + 961, , "未知处理类型，未执行。"
    End Select
End Function


' Legacy bounded selection extension: Reconcile:sheet:column:header:firstRow,lastRow.
' Existing four-field modern commands keep their whole-column semantics.
' Only decimal row bounds are accepted; never evaluate an address or VBA expression.
Public Function DHUI_ReconcileRange(ByVal action As String, ByVal wb As Workbook) As Range
    Dim parts As Variant, limits As Variant, ws As Worksheet, col As Long
    Dim first As Long, last As Long, digits As String, i As Long, j As Long, ch As String, found As Boolean
    parts = Split(action, ":")
    If UBound(parts) <> 3 And UBound(parts) <> 4 Then Err.Raise 5, , "Invalid Reconcile parameter count"
    If CStr(parts(0)) <> "Reconcile" Then Err.Raise 5, , "Invalid Reconcile command"
    If CStr(parts(3)) <> "true" And CStr(parts(3)) <> "false" Then Err.Raise 5, , "Invalid Reconcile header"
    col = DH_ColLetterToIndex(CStr(parts(2)))
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, CStr(parts(1)), vbTextCompare) = 0 Then found = True: Exit For
    Next ws
    If Not found Then Err.Raise 5, , "Reference worksheet not found"
    If UBound(parts) = 3 Then
        Set DHUI_ReconcileRange = ws.Columns(col)
        Exit Function
    End If
    limits = Split(CStr(parts(4)), ",")
    If UBound(limits) <> 1 Then Err.Raise 5, , "Invalid Reconcile row bounds"
    For i = 0 To 1
        digits = CStr(limits(i))
        If Len(digits) = 0 Or Len(digits) > 7 Then Err.Raise 5, , "Invalid Reconcile row bounds"
        For j = 1 To Len(digits)
            ch = Mid$(digits, j, 1)
            If ch < "0" Or ch > "9" Then Err.Raise 5, , "Invalid Reconcile row bounds"
        Next j
    Next i
    first = CLng(limits(0)): last = CLng(limits(1))
    If first < 1 Or last < first Or last > ws.Rows.Count Then Err.Raise 5, , "Invalid Reconcile row bounds"
    Set DHUI_ReconcileRange = ws.Range(ws.Cells(first, col), ws.Cells(last, col))
End Function

' Single action registry: display name, consequence shown in the native confirmation,
' batch rank and whether the action must run alone.
Public Function DHUI_ActionInfo(ByVal action As String) As Variant
    Dim key As String, lkParts As Variant, retCount As Long, lkDesc As String
    Dim apParts As Variant, apSheetCount As Long, apDesc As String
    Dim upParts As Variant, upIdCount As Long, upValCount As Long, upDesc As String

    Dim spParts As Variant, spDesc As String
    key = action
    If Left$(key, 7) = "Lookup:" Then
        lkParts = Split(key, ":")
        If UBound(lkParts) = 7 Then
            retCount = UBound(Split(CStr(lkParts(3)), ",")) + 1
            If StrComp(CStr(lkParts(5)), "AUTO", vbTextCompare) = 0 Then
                lkDesc = "按指定匹配键从「" & CStr(lkParts(1)) & "」在目标表右侧自动新增 " & retCount & " 列补全数据；已有内容、公式或来源错误值保护跳过。"
            Else
                lkDesc = "按指定匹配键从「" & CStr(lkParts(1)) & "」跨表补全 " & retCount & " 列至目标列（" & CStr(lkParts(5)) & "）；已有内容、公式或来源错误值保护跳过。"
            End If
        Else
            lkDesc = "按显式指定匹配键跨表补全目标列；已有内容、公式或来源错误值保护跳过。"
        End If
        DHUI_ActionInfo = Array("跨表匹配补全", lkDesc, 0, True)
        Exit Function
    End If
    If Left$(key, 7) = "Append:" Then
        apParts = Split(key, ":")
        If UBound(apParts) >= 5 Then
            apSheetCount = UBound(Split(CStr(apParts(1)), ",")) + 1
            apDesc = "将所选 " & apSheetCount & " 张工作表按表头对齐纵向追加至新工作表「" & CStr(apParts(5)) & "」；原工作表不改动，" & IIf(StrComp(CStr(apParts(3)), "exclude", vbTextCompare) = 0, "自动排除已确认的小计/合计行并记录审计。", "保留全部数据行。")
        Else
            apDesc = "将多张工作表按列名对齐纵向追加至新建汇总表，原表数据不改动。"
        End If
        DHUI_ActionInfo = Array("多表纵向追加", apDesc, 0, True)
        Exit Function
    End If
    If Left$(key, 8) = "Unpivot:" Then
        upParts = Split(key, ":")
        If UBound(upParts) = 7 Then
            upIdCount = UBound(Split(CStr(upParts(1)), ",")) + 1
            upValCount = UBound(Split(CStr(upParts(2)), ",")) + 1
            upDesc = "保留固定维度列（" & CStr(upParts(1)) & " 共 " & upIdCount & " 列），将 " & upValCount & " 个横向指标列（" & CStr(upParts(2)) & "）逆透视展开至新工作表「" & CStr(upParts(7)) & "」；原工作表不改动。"
        Else
            upDesc = "将当前二维交叉表逆透视转换为一维标准明细表并写入新建工作表，原表数据不改动。"
        End If
        DHUI_ActionInfo = Array("宽表转长表", upDesc, 0, True)
        Exit Function
    End If
    If Left$(key, 11) = "SplitSheet:" Then
        spParts = Split(key, ":")
        If UBound(spParts) = 4 Then
            spDesc = "按当前表 " & CStr(spParts(1)) & " 列的不同取值自动拆分为多张独立工作表；原工作表与已有工作表不改动，" & IIf(StrComp(CStr(spParts(3)), "exclude", vbTextCompare) = 0, "自动排除已确认的小计/合计行。", "保留全部数据行。")
        Else
            spDesc = "按指定列的不同取值自动拆分为多张独立工作表，原表数据不改动。"
        End If
        DHUI_ActionInfo = Array("按列拆表", spDesc, 0, True)
        Exit Function
    End If
    If Left$(key, 10) = "Reconcile:" Then key = "Reconcile"
    If Left$(key, 10) = "Transform:" Then key = Mid$(key, 11)
    Select Case key
        Case "Unmerge": DHUI_ActionInfo = Array("拆分合并", "只解除完整纵向单列合并并向下填值。", 0, False)
        Case "Space": DHUI_ActionInfo = Array("区域清理", "只清理变化文本；保留英文分隔、前导零和公式；合并格跳过。", 1, False)
        Case "Unit": DHUI_ActionInfo = Array("数字转换", "原地转纯数值，合法单位剥离并记日志。", 1, False)
        Case "Unit:Force": DHUI_ActionInfo = Array("数字转换（同列多单位也转）", "同列单位不统一也转换：会去掉单位文字只剩数字（件/台/箱等信息会丢），只影响这一次。", 1, False)
        Case "Date": DHUI_ActionInfo = Array("日期规范化", "明确日期文本转真实日期；歧义原值保留。", 1, False)
        Case "Mark": DHUI_ActionInfo = Array("高亮重复项", "只添加可清除标记，保留原底色及内容。", 1, False)
        Case "BlankRows": DHUI_ActionInfo = Array("删除空行", "仅选区列下方数据上移；全列/全表按原规则分区。", 2, False)
        Case "Dedup": DHUI_ActionInfo = Array("删除重复行", "删除整行或表记录，保留首次项；必须覆盖全部记录列。", 3, False)
        Case "Phone", "ID", "Name": DHUI_ActionInfo = Array("敏感数据脱敏", "所选常量原地打码；基线仍含原文，表头和公式跳过。", 4, False)
        Case "Split": DHUI_ActionInfo = Array("单位拆分", "右侧插入整列，原列转数值；右侧数据右移。", 0, True)
        Case "Delivery": DHUI_ActionInfo = Array("一键交表礼仪", "全簿可见业务表回A1、缩放100%、冻结首行并开启网格线。", 0, True)
        Case "ClearTemporary": DHUI_ActionInfo = Array("清除临时标记", "清除插件生成的体检报告与临时条件标记；保留用户表。", 0, True)
        Case "Lookup": DHUI_ActionInfo = Array("跨表匹配补全", "按显式指定匹配键跨表补全目标列；已有内容、公式或来源错误值保护跳过。", 0, True)
        Case "Append": DHUI_ActionInfo = Array("多表纵向追加", "将多张工作表按列名对齐纵向追加至新建汇总表，原表数据不改动。", 0, True)
        Case "Unpivot": DHUI_ActionInfo = Array("宽表转长表", "将二维交叉表逆透视转换为一维标准明细表并写入新建工作表，原表数据不改动。", 0, True)

        Case "SplitSheet": DHUI_ActionInfo = Array("按列拆表", "按指定列的不同取值自动拆分为多张独立工作表，原表数据不改动。", 0, True)
        Case "Reconcile": DHUI_ActionInfo = Array("两表对账", "使用可清除条件格式标出主表与核对表独有项，保留原底色。", 0, True)
        Case Else: DHUI_ActionInfo = Empty
    End Select
End Function

Public Function DHUI_ActionName(ByVal action As String) As String
    Dim info As Variant
    info = DHUI_ActionInfo(action)
    If Not IsEmpty(info) Then DHUI_ActionName = CStr(info(0))
End Function

Public Function DHUI_Consequence(ByVal action As String) As String
    Dim info As Variant
    info = DHUI_ActionInfo(action)
    If Not IsEmpty(info) Then DHUI_Consequence = CStr(info(1))
End Function

' Shared result unit; presentation adapters must not infer rows from a cell count.
Public Function DHUI_CountUnit(ByVal action As String) As String
    Dim key As String
    key = Split(action, ":")(0)
    Select Case key
        Case "Append", "Unpivot", "SplitSheet", "BlankRows", "Dedup", "Mark": DHUI_CountUnit = "行"
        Case "Unmerge": DHUI_CountUnit = "合并区域"
        Case "Delivery": DHUI_CountUnit = "张次工作表"
        Case "ClearTemporary": DHUI_CountUnit = "项"
        Case "Reconcile": DHUI_CountUnit = "处"
        Case Else: DHUI_CountUnit = "格"
    End Select
End Function
