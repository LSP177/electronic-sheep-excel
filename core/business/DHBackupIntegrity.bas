Attribute VB_Name = "DHBackupIntegrity"
Option Explicit

' Inspect only the password-presence flag; never request/read/store a password.
' Keep this outside the backup-error wrapper so the user sees the real limit.
Public Sub DH_RequireBackupCompatible(ByVal wb As Workbook)
    If wb Is Nothing Then Err.Raise vbObjectError + 701, , "请先打开数据工作簿。"
    If wb.HasPassword Then Err.Raise vbObjectError + 919, , _
        "该文件设置了打开密码，暂不支持在其上执行需要基线备份的数据修改操作。请保留原文件，在您有权使用的无打开密码本地副本中整理；插件不会索取或移除密码。"
End Sub

' First verification opens the copy with macros/events/links disabled.
' Subsequent reuse requires the same SHA256 tied to this full source path.
Public Sub DH_VerifyBaseline(ByVal wb As Workbook, ByVal path As String, Optional ByVal fresh As Boolean = False)
    Dim fso As Object, fingerprint As String, digest As String, expected As String, handle As Object, copy As Workbook
    Dim original As Workbook, selected As Range, security As Long, events As Boolean, alerts As Boolean, savedState As Boolean
    Dim message As String, code As Long, mutexHandle As LongPtr
    DH_RequireBackupCompatible wb
    On Error GoTo Failed
    Set fso = CreateObject("Scripting.FileSystemObject")
    fingerprint = fso.GetParentFolderName(fso.GetParentFolderName(DH_DailyBackupPath(wb))) & "\" & Format$(Date, "yyyymmdd") & ".sha256"
    mutexHandle = DH_Lock("Baseline_" & DH_WorkbookKey(wb))
    digest = DH_FileHash(path)
    expected = LCase$(fso.GetAbsolutePathName(wb.FullName)) & vbCrLf & digest
    If Not fresh And fso.FileExists(fingerprint) Then
        Set handle = fso.OpenTextFile(fingerprint, 1, False, -1)
        message = handle.ReadAll: handle.Close: Set handle = Nothing
        If StrComp(message, expected, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 915, , "当日基线校验不符，可能损坏或被替换；原基线保留，已阻止修改。"
    Else
        Set original = Application.ActiveWorkbook
        If TypeOf Application.Selection Is Range Then Set selected = Application.Selection
        security = Application.AutomationSecurity: events = Application.EnableEvents: alerts = Application.DisplayAlerts: savedState = True
        Application.AutomationSecurity = 3: Application.EnableEvents = False: Application.DisplayAlerts = False
        Set copy = Application.Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, Password:="DH_NO_PASSWORD", _
            WriteResPassword:="DH_NO_PASSWORD", IgnoreReadOnlyRecommended:=True, AddToMru:=False, CorruptLoad:=xlNormalLoad)
        If copy.Sheets.Count = 0 Or copy.FileFormat <> wb.FileFormat Then Err.Raise vbObjectError + 915, , "当日基线不是可用的原格式工作簿。"
        copy.Close False: Set copy = Nothing
        Application.AutomationSecurity = security: Application.EnableEvents = events: Application.DisplayAlerts = alerts
        If Not original Is Nothing Then original.Activate
        If Not selected Is Nothing Then selected.Select
        savedState = False
        If DH_FileHash(path) <> digest Then Err.Raise vbObjectError + 915, , "验证期间基线发生变化，已停止。"
        Set handle = fso.CreateTextFile(fingerprint, True, True): handle.Write expected: handle.Close: Set handle = Nothing
    End If
    DH_Unlock mutexHandle
    Exit Sub
Failed:
    message = Err.Description: code = Err.Number
    On Error Resume Next
    If Not handle Is Nothing Then handle.Close
    If Not copy Is Nothing Then copy.Close False
    If savedState Then
        Application.AutomationSecurity = security: Application.EnableEvents = events: Application.DisplayAlerts = alerts
        If Not original Is Nothing Then original.Activate
        If Not selected Is Nothing Then selected.Select
    End If
    DH_Unlock mutexHandle
    On Error GoTo 0
    Err.Raise code, "基线完整性", "无法确认当日基线可恢复，已停止破坏操作：" & message
End Sub
