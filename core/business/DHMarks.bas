Attribute VB_Name = "DHMarks"
Option Explicit
Private Const OWNER As String = "ExcelDataHelper.v0.4.marker"

Private Function MarkName(ByVal kind As Long) As String
    Select Case kind
        Case 1: MarkName = "_DH04_RECON_LEFT"
        Case 2: MarkName = "_DH04_RECON_RIGHT"
        Case 3: MarkName = "_DH04_DUPLICATE"
        Case Else: Err.Raise vbObjectError + 870, , "未知标记类型。"
    End Select
End Function

Private Function ShortName(ByVal text As String) As String
    Dim cut As Long
    cut = InStrRev(text, "!")
    ShortName = Mid$(text, cut + 1)
End Function

Private Function AnyScopeName(ByVal wb As Workbook, ByVal text As String) As Boolean
    Dim item As Name
    For Each item In wb.Names
        If StrComp(ShortName(item.Name), text, vbTextCompare) = 0 Then AnyScopeName = True: Exit Function
    Next item
End Function

Private Function GlobalMark(ByVal wb As Workbook, ByVal kind As Long) As Name
    Dim item As Name
    For Each item In wb.Names
        If StrComp(ShortName(item.Name), MarkName(kind), vbTextCompare) = 0 Then
            If InStr(item.Name, "!") > 0 Then Set GlobalMark = Nothing: Exit Function
            If item.Comment = OWNER And item.RefersTo = "=TRUE" Then Set GlobalMark = item
        End If
    Next item
End Function

Public Sub DH_ValidateMarkNames(ByVal wb As Workbook)
    Dim i As Long, item As Name
    For Each item In wb.Names
        For i = 1 To 3
            If StrComp(ShortName(item.Name), MarkName(i), vbTextCompare) = 0 Then
                If InStr(item.Name, "!") > 0 Or item.Comment <> OWNER Or item.RefersTo <> "=TRUE" Then _
                    Err.Raise vbObjectError + 871, , "标记名称存在用户或工作表局部作用域冲突，原名称保留，已停止：" & item.Name
            End If
        Next i
    Next item
End Sub

Public Function DH_AddMark(ByVal target As Range, ByVal kind As Long) As Object
    Dim wb As Workbook, item As Name, rule As Object, color As Long
    Dim message As String, code As Long
    On Error GoTo Failed
    Set wb = target.Parent.Parent
    DH_ValidateMarkNames wb
    Set item = GlobalMark(wb, kind)
    If item Is Nothing Then
        Set item = wb.Names.Add(Name:=MarkName(kind), RefersTo:="=TRUE", Visible:=False)
        item.Comment = OWNER
    End If
    color = RGB(226, 239, 218)
    If kind = 2 Then color = RGB(255, 242, 204)
    If kind = 3 Then color = RGB(255, 199, 206)
    Set rule = target.FormatConditions.Add(Type:=xlExpression, Formula1:="=" & MarkName(kind))
    rule.Interior.Pattern = xlSolid: rule.Interior.Color = color
    rule.StopIfTrue = False: rule.SetFirstPriority
    Set DH_AddMark = rule
    Exit Function
Failed:
    message = Err.Description: code = Err.Number
    On Error Resume Next
    If Not rule Is Nothing Then rule.Delete
    On Error GoTo 0
    Err.Raise code, "插件标记", message
End Function

Public Function DH_IsManagedRule(ByVal rule As Object, ByVal wb As Workbook) As Boolean
    Dim i As Long, item As Name
    On Error GoTo Unowned
    If rule.Type <> xlExpression Then Exit Function
    For i = 1 To 3
        If rule.Formula1 = "=" & MarkName(i) Then
            DH_ValidateMarkNames wb
            Set item = GlobalMark(wb, i)
            If Not item Is Nothing Then DH_IsManagedRule = True
            Exit Function
        End If
    Next i
Unowned:
End Function

Public Function DH_CountMarks(ByVal selected As Range, ByVal kind As Long) As Long
    Dim rule As Object
    For Each rule In selected.Parent.Cells.FormatConditions
        If DH_IsManagedRule(rule, selected.Parent.Parent) Then
            If rule.Formula1 = "=" & MarkName(kind) Then
                If Not Application.Intersect(rule.AppliesTo, selected) Is Nothing Then DH_CountMarks = DH_CountMarks + 1
            End If
        End If
    Next rule
End Function

' Refresh only the current operation's selection/kind; retain marks outside it.
Public Sub DH_TakeMarks(ByVal selected As Range, ByVal kind As Long, ByVal snapshots As Collection, ByVal added As Collection)
    Dim rule As Object, applies As Range, remaining As Range, created As Object, candidates As New Collection
    For Each rule In selected.Parent.Cells.FormatConditions
        If DH_IsManagedRule(rule, selected.Parent.Parent) Then
            If rule.Formula1 = "=" & MarkName(kind) Then
                If Not Application.Intersect(rule.AppliesTo, selected) Is Nothing Then candidates.Add rule
            End If
        End If
    Next rule
    For Each rule In candidates
        Set applies = rule.AppliesTo: Set remaining = DH_SubtractRange(applies, selected)
        rule.Delete: snapshots.Add Array(applies, kind)
        If Not remaining Is Nothing Then
            Set created = DH_AddMark(remaining, kind): added.Add created
        End If
    Next rule
End Sub

Public Sub DH_RestoreMarks(ByVal snapshots As Collection)
    Dim item As Variant, created As Object, target As Range
    For Each item In snapshots
        Set target = item(0): Set created = DH_AddMark(target, CLng(item(1)))
    Next item
End Sub

' Retain proof from a successfully generated legacy reconciliation report before
' health-report consolidation removes that report. Hidden names contain paths,
' never source data, and are removed by ClearTemporary.
Private Function LegacyProof(ByVal wb As Workbook, Optional ByVal remember As Boolean = False) As Object
    Dim paths As Object, saved As Object, item As Name, ws As Worksheet, value As Variant, path As String, fso As Object
    Dim root As String, key As String, i As Long, candidate As String, existing As Name
    Set paths = CreateObject("Scripting.Dictionary"): paths.CompareMode = vbTextCompare
    Set saved = CreateObject("Scripting.Dictionary"): saved.CompareMode = vbTextCompare
    Set fso = CreateObject("Scripting.FileSystemObject")
    root = Environ$("LOCALAPPDATA") & "\ExcelDataHelper\details\": key = DH_WorkbookKey(wb) & "_"
    For Each item In wb.Names
        If item.Comment = OWNER & ".legacy" Then
            value = item.RefersTo
            If Left$(value, 2) = "=""" And Right$(value, 1) = """" Then
                path = Replace$(Mid$(value, 3, Len(value) - 3), """""", """")
                paths(path) = True: saved(path) = True
            End If
        End If
    Next item
    For Each ws In wb.Worksheets
        If DH_IsReportSheet(ws) Then
            If ws.Range("A1").Value2 = "数据助手：两表对账" And ws.Range("A2").Value2 = "执行时间" Then
                value = ws.Range("B6").Value2
                If VarType(value) = vbString Then paths(CStr(value)) = True
            End If
        End If
    Next ws
    For Each value In paths.Keys
        path = fso.GetAbsolutePathName(CStr(value))
        If Left$(path, Len(root)) <> root Or Left$(fso.GetFileName(path), Len(key)) <> key Then
            paths.Remove CStr(value)
        ElseIf remember And Not saved.Exists(CStr(value)) Then
            Do
                i = i + 1: candidate = "_DH04_LEGACY_" & i: Set existing = Nothing
                On Error Resume Next
                Set existing = wb.Names(candidate)
                On Error GoTo 0
            Loop Until Not AnyScopeName(wb, candidate)
            Set item = wb.Names.Add(Name:=candidate, RefersTo:="=""" & Replace$(CStr(value), """", """""") & """", Visible:=False)
            item.Comment = OWNER & ".legacy"
        End If
    Next value
    Set LegacyProof = paths
End Function

Public Sub DH_RememberLegacyMarks(ByVal wb As Workbook)
    Dim paths As Object
    Set paths = LegacyProof(wb, True)
End Sub

' v0.3 fills had no embedded ownership. Require a legacy success-report link as
' well as this saved path's audit address and exact color. Plans alone are not proof.
Private Function LegacyTargets(ByVal wb As Workbook) As Collection
    Dim targets As New Collection, fso As Object, root As String, day As Object, file As Object, stream As Object
    Dim text As String, lines As Variant, fields As Variant, i As Long, cut As Long, location As String, ws As Worksheet, target As Range
    Dim pattern As Object, match As Object, seen As Object, color As Long, key As String, fileCount As Long
    Dim proof As Object
    Set proof = LegacyProof(wb)
    If proof.Count = 0 Then Set LegacyTargets = targets: Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject"): Set seen = CreateObject("Scripting.Dictionary")
    root = Environ$("LOCALAPPDATA") & "\ExcelDataHelper\details"
    If Not fso.FolderExists(root) Then Set LegacyTargets = targets: Exit Function
    Set pattern = CreateObject("VBScript.RegExp"): pattern.Pattern = "^R([0-9]+)C([0-9]+)$"
    For Each day In fso.GetFolder(root).SubFolders
        For Each file In day.Files
            If proof.Exists(file.Path) Then
                fileCount = fileCount + 1
                If fileCount > 2000 Then Err.Raise vbObjectError + 872, , "历史明细过多，请先归档后清除临时标记。"
                Set stream = CreateObject("ADODB.Stream"): stream.Type = 2: stream.Charset = "utf-8": stream.Open
                stream.LoadFromFile file.Path: text = stream.ReadText: stream.Close
                lines = Split(text, vbCrLf)
                If InStr(lines(0), "两表对账") > 0 And InStr(lines(0), "v0.4") = 0 Then
                    For i = 2 To UBound(lines)
                        fields = Split(lines(i), vbTab)
                        If UBound(fields) = 5 Then
                            color = 0
                            If fields(1) = "主表缺少对应项" Then color = RGB(226, 239, 218)
                            If fields(1) = "核对表陌生项" Then color = RGB(255, 242, 204)
                            If color <> 0 Then
                                cut = InStrRev(fields(0), "!")
                                If cut > 1 Then
                                    Set ws = Nothing: Set target = Nothing
                                    On Error Resume Next
                                    Set ws = wb.Worksheets(Left$(fields(0), cut - 1))
                                    On Error GoTo 0
                                    location = Mid$(fields(0), cut + 1)
                                    If Not ws Is Nothing And pattern.Test(location) Then
                                        Set match = pattern.Execute(location)(0)
                                        On Error Resume Next
                                        Set target = ws.Cells(CLng(match.SubMatches(0)), CLng(match.SubMatches(1)))
                                        On Error GoTo 0
                                        If Not target Is Nothing Then
                                            key = ws.CodeName & "!" & target.Address
                                            If Not seen.Exists(key) Then
                                                If target.Interior.Pattern = xlSolid And target.Interior.Color = color Then
                                                    targets.Add target: seen.Add key, True
                                                    If targets.Count > DH_MAX_AREAS Then Err.Raise vbObjectError + 872, , "旧版标记过多，请分批归档处理。"
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next i
                End If
            End If
        Next file
    Next day
    Set LegacyTargets = targets
End Function

Public Function DH_ClearTemporary(ByVal wb As Workbook) As Long
    Dim original As Worksheet, selected As Range, ws As Worksheet, business As Worksheet, rule As Object, target As Range
    Dim reports As New Collection, rules As New Collection, legacy As Collection, details As New Collection, item As Variant
    Dim i As Long, reportCount As Long, ruleCount As Long, backup As String, audit As String, message As String, code As Long
    Dim state As DH_AppState, frozen As Boolean, oldAlerts As Boolean
    On Error GoTo Failed
    DH_RequireWorkbook wb, True
    DH_ValidateMarkNames wb
    oldAlerts = Application.DisplayAlerts
    If TypeOf Application.ActiveSheet Is Worksheet Then Set original = Application.ActiveSheet
    If Not original Is Nothing Then Set selected = DH_OriginalSelection(original)
    Set business = DH_FirstBusinessSheet(wb)
    If business Is Nothing Then Err.Raise vbObjectError + 873, , "必须保留至少一张可见业务表，未清除。"
    If Not original Is Nothing Then
        If DH_IsReportSheet(original) Then Set original = business: Set selected = Nothing
    End If
    For Each ws In wb.Worksheets
        If DH_IsReportSheet(ws) Then
            If ws.ProtectContents Then Err.Raise vbObjectError + 874, , "报告受保护：" & ws.Name
            reports.Add ws
        End If
        For i = 1 To ws.Cells.FormatConditions.Count
            Set rule = ws.Cells.FormatConditions(i)
            If DH_IsManagedRule(rule, wb) Then
                If ws.ProtectContents Then Err.Raise vbObjectError + 874, , "标记所在工作表受保护：" & ws.Name
                rules.Add rule
            End If
        Next i
    Next ws
    Set legacy = LegacyTargets(wb)
    For Each target In legacy
        If target.Parent.ProtectContents Then Err.Raise vbObjectError + 874, , "旧标记所在工作表受保护。"
    Next target
    DH_LastSummary = "报告=" & reports.Count & "；插件标记规则=" & rules.Count & "；有历史明细佐证的旧底色=" & legacy.Count
    details.Add Array("工作簿", "清除临时标记", DH_LastSummary, "不记录单元格值", "移除报告及插件标记", "")
    audit = DH_SaveAudit("清除临时标记", wb, details)
    If reports.Count + rules.Count + legacy.Count = 0 Then
        DH_Log "清除临时标记", wb, "", "", 0, "成功；无修改": Exit Function
    End If
    DH_Freeze state: frozen = True
    backup = DH_Backup(wb)
    business.Activate
    For i = rules.Count To 1 Step -1: rules(i).Delete: ruleCount = ruleCount + 1: Next i
    For Each target In legacy: target.Interior.Pattern = xlNone: Next target
    oldAlerts = Application.DisplayAlerts: Application.DisplayAlerts = False
    For i = reports.Count To 1 Step -1: reports(i).Delete: reportCount = reportCount + 1: Next i
    For i = wb.Names.Count To 1 Step -1
        If wb.Names(i).Comment = OWNER & ".legacy" Then wb.Names(i).Delete
    Next i
    Application.DisplayAlerts = oldAlerts
    DH_ClearTemporary = reportCount + ruleCount + legacy.Count
    DH_Log "清除临时标记", wb, "", "", DH_ClearTemporary, "成功；" & DH_LastSummary & "；旧版底色仅清除，不推断此前原底色", backup
    If Not original Is Nothing Then
        If DH_IsReportSheet(business) Then Set original = business
    End If
    On Error Resume Next
    DH_RestoreFocus original, selected
    If Err.Number <> 0 Then business.Activate
    On Error GoTo 0
    DH_Restore state
    Exit Function
Failed:
    message = Err.Description: code = Err.Number
    If frozen Then
        Application.DisplayAlerts = oldAlerts
        DH_Restore state
        message = message & "；清除可能已部分完成，请恢复基线：" & backup
    End If
    DH_RestoreFocus business
    DH_LogFailure "清除临时标记", wb, "", "", message, backup
    Err.Raise code, "电子羊表格", message
End Function
