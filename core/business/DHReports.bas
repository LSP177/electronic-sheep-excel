Attribute VB_Name = "DHReports"
Option Explicit

' A previous-path marker is evidence for an explicit retain/rename prompt only.
' It NEVER grants permission to clear, overwrite or delete this worksheet.
Public Function DH_IsPreviousReport(ByVal ws As Worksheet) As Boolean
    Dim item As Object, fields As Variant, marker As String, markers As Long, keyPattern As Object
    On Error GoTo NotPrevious
    For Each item In ws.CustomProperties
        If item.Name = "ExcelDataHelperReport" Then marker = CStr(item.Value): markers = markers + 1
    Next item
    If markers <> 1 Then Exit Function
    fields = Split(marker, "|")
    If UBound(fields) <> 2 Then Exit Function
    ' v043 is the persisted ownership schema, retained in v0.4.4 for compatibility.
    If fields(0) <> "ExcelDataHelper.Report.v043" Or fields(2) <> ws.CodeName Then Exit Function
    If fields(1) = DH_WorkbookKey(ws.Parent) Then Exit Function
    Set keyPattern = CreateObject("VBScript.RegExp")
    keyPattern.Pattern = "^[0-9A-F]{1,8}-[0-9A-F]{1,8}$"
    If Not keyPattern.Test(CStr(fields(1))) Then Exit Function
    If VarType(ws.Range("A1").Value2) <> vbString Then Exit Function
    If ws.Range("A1").Value2 <> "数据体检报告" And ws.Range("A1").Value2 <> "全簿健康总览报告" Then Exit Function
    If ws.Range("A2").Value2 <> "工作簿" Or ws.Range("A7").Value2 <> "数据与公式健康安全看板" Then Exit Function
    DH_IsPreviousReport = True
NotPrevious:
End Function

Private Function RetainedName(ByVal wb As Workbook) As String
    Dim base As String, candidate As String, index As Long, item As Object, taken As Boolean
    base = "体检报告_留存_" & Format$(Now, "yyyymmdd_hhnnss")
    candidate = base
    Do
        taken = False
        For Each item In wb.Sheets
            If StrComp(item.Name, candidate, vbTextCompare) = 0 Then taken = True: Exit For
        Next item
        If Not taken Then RetainedName = candidate: Exit Function
        index = index + 1: candidate = Left$(base, 25) & "_" & CStr(index)
    Loop
End Function

' Only the Ribbon/UI path asks; core checks remain noninteractive and fail closed.
Public Function DH_ResolvePreviousReport(ByVal wb As Workbook) As Boolean
    Dim item As Object, previous As Worksheet, newName As String, backup As String
    Dim message As String, code As Long, renamed As Boolean
    On Error GoTo Failed
    DH_RequireWorkbook wb, True
    For Each item In wb.Sheets
        If item.Name = "数据体检报告" Then
            If TypeOf item Is Worksheet Then
                If DH_IsPreviousReport(item) Then Set previous = item
            End If
            Exit For
        End If
    Next item
    If previous Is Nothing Then DH_ResolvePreviousReport = True: Exit Function
    If previous.ProtectContents Then Err.Raise vbObjectError + 851, , "旧报告受保护，请先取消保护；内容已保留。"
    newName = RetainedName(wb)
    If MsgBox(DH_WorkbookTarget(wb, "工作表“" & previous.Name & "”（仅改名留存，不修改单元格）") & _
        "检测到另存为或移动文件后留下的旧体检报告。" & vbCrLf & _
        "是否将“数据体检报告”改名为“" & newName & "”，保留全部内容后继续体检？" & vbCrLf & _
        "不会接管、清空或删除旧表。改名前先建立/验证当前文件的当日基线；取消不改动。", _
        vbYesNo + vbQuestion + vbDefaultButton2, "旧报告改名留存") <> vbYes Then
        DH_Log "旧报告留存", wb, previous.Name, "", 0, "用户取消；未改名、未创建基线"
        Exit Function
    End If
    DH_RequireWorkbook wb, True
    If previous.Name <> "数据体检报告" Or Not DH_IsPreviousReport(previous) Then _
        Err.Raise vbObjectError + 850, , "旧报告状态已变化，已停止改名。"
    backup = DH_Backup(wb)
    previous.Name = newName: renamed = True
    DH_Log "旧报告留存", wb, newName, "", 1, "成功；仅改名留存，旧标记未接管；未自动保存", backup
    DH_ResolvePreviousReport = True
    Exit Function
Failed:
    message = Err.Description: code = Err.Number
    If renamed Then message = message & "；旧报告已留存为“" & newName & "”，内容保留。"
    DH_LogFailure "旧报告留存", wb, "", "", message, backup
    Err.Raise code, "电子羊表格", message
End Function

Public Sub DH_ReportPreflight(ByVal wb As Workbook)
    Dim ws As Worksheet, item As Object
    DH_RequireWorkbook wb, True
    For Each item In wb.Sheets
        If item.Name = "数据体检报告" Then
            If Not TypeOf item Is Worksheet Then Err.Raise vbObjectError + 850, , "数据体检报告名称被非工作表占用，请先更名。"
            If Not DH_IsReportSheet(item) Then Err.Raise vbObjectError + 850, , "存在归属未确认的同名表（可能是旧报告或业务表），已保留；请先手动改名后再体检。"
        End If
    Next item
    For Each ws In wb.Worksheets
        If DH_IsReportSheet(ws) Then
            If ws.ProtectContents Then Err.Raise vbObjectError + 851, , "插件报告受保护，请先取消保护：" & ws.Name
        End If
    Next ws
End Sub

Public Function DH_HealthSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet, target As Worksheet, oldAlerts As Boolean, message As String, code As Long, i As Long
    DH_ReportPreflight wb
    DH_RememberLegacyMarks wb
    For Each ws In wb.Worksheets
        If ws.Name = "数据体检报告" Then Set target = ws
    Next ws
    If target Is Nothing Then
        Set target = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        target.Name = "数据体检报告": DH_TagReport target
    End If
    oldAlerts = Application.DisplayAlerts
    On Error GoTo Failed
    Application.DisplayAlerts = False
    ' Only plugin-owned/recognized derived reports may be removed.
    For i = wb.Worksheets.Count To 1 Step -1
        Set ws = wb.Worksheets(i)
        If Not ws Is target Then
            If DH_IsReportSheet(ws) Then ws.Delete
        End If
    Next i
    Application.DisplayAlerts = oldAlerts
    target.Visible = xlSheetVisible
    target.UsedRange.UnMerge: target.UsedRange.Clear
    If target.Hyperlinks.Count > 0 Then target.Hyperlinks.Delete
    Set DH_HealthSheet = target
    Exit Function
Failed:
    message = Err.Description: code = Err.Number
    Application.DisplayAlerts = oldAlerts
    Err.Raise code, "体检报告", message
End Function

Public Function DH_FirstBusinessSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            If Not DH_IsReportSheet(ws) And Not DH_IsPreviousReport(ws) Then Set DH_FirstBusinessSheet = ws: Exit Function
        End If
    Next ws
End Function
