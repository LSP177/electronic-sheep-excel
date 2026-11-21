Attribute VB_Name = "DHLog"
Option Explicit
Public DH_LastLogError As String

Public Function DH_LogPath() As String
    Dim base As String
    base = Environ$("LOCALAPPDATA") & "\ExcelDataHelper\logs"
    DH_EnsureFolder base
    DH_LogPath = base & "\" & Format$(Date, "yyyymmdd") & ".tsv"
End Function

Private Function DH_Field(ByVal text As String) As String
    Dim lead As String
    text = Replace$(Replace$(Replace$(text, vbTab, " "), vbCr, " "), vbLf, " ")
    lead = LTrim$(text)
    If Len(lead) > 0 Then
        If InStr("=+-@", Left$(lead, 1)) > 0 Or Left$(lead, 1) = Chr$(34) Then text = "'" & text
    End If
    DH_Field = text
End Function

Public Function DH_ReadLog() As String
    Dim fso As Object, stream As Object, path As String, mutexHandle As LongPtr, message As String, code As Long
    On Error GoTo Failed
    mutexHandle = DH_Lock("OperationLog")
    path = DH_LogPath(): Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(path) Then
        Set stream = CreateObject("ADODB.Stream"): stream.Type = 2: stream.Charset = "utf-8": stream.Open
        stream.LoadFromFile path: DH_ReadLog = stream.ReadText: stream.Close
    End If
    DH_Unlock mutexHandle
    Exit Function
Failed:
    message = Err.Description: code = Err.Number
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    DH_Unlock mutexHandle
    On Error GoTo 0
    Err.Raise code, "读取日志", message
End Function

Public Sub DH_Log(ByVal operation As String, ByVal wb As Workbook, ByVal sheetName As String, _
    ByVal address As String, ByVal changes As Long, ByVal result As String, Optional ByVal backup As String = "")
    Dim text As String, bookName As String, stream As Object, bytes() As Byte, path As String
    Dim file As LongPtr, opened As Boolean, mutexHandle As LongPtr, message As String, code As Long
    On Error GoTo Failed
    mutexHandle = DH_Lock("OperationLog")
    path = DH_LogPath()
    If Not wb Is Nothing Then bookName = wb.FullName
    text = Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & DH_Field(operation) & vbTab & DH_Field(bookName) & vbTab & _
        DH_Field(sheetName) & vbTab & DH_Field(address) & vbTab & CStr(changes) & vbTab & DH_Field(result) & vbTab & DH_Field(backup) & vbCrLf
    ' Win32 Unicode path (N-3); exclusive like the former Open ... Lock Read Write.
    file = DH_FileOpen(path, True): opened = True
    If DH_FileSize(file) = 0 Then text ="执行时间" & vbTab & "操作" & vbTab & "文件" & vbTab & "工作表" & vbTab & "处理范围" & vbTab & "修改数量" & vbTab & "结果" & vbTab & "备份" & vbCrLf & text
    Set stream = CreateObject("ADODB.Stream"): stream.Type = 2: stream.Charset = "utf-8": stream.Open: stream.WriteText text
    stream.Position = 0: stream.Type = 1: stream.Position = 3: bytes = stream.Read: stream.Close
    DH_FileAppend file, bytes
    DH_FileClose file: opened = False
    DH_LastLogError = ""
    DH_Unlock mutexHandle
    Exit Sub
Failed:
    message = Err.Description: code = Err.Number
    DH_LastLogError = "日志写入失败：" & message
    On Error Resume Next
    If opened Then DH_FileClose file
    If Not stream Is Nothing Then stream.Close
    DH_Unlock mutexHandle
    On Error GoTo 0
    Err.Raise code, "操作日志", DH_LastLogError
End Sub

Public Sub DH_LogFailure(ByVal operation As String, ByVal wb As Workbook, ByVal sheetName As String, _
    ByVal address As String, ByVal message As String, Optional ByVal backup As String = "")
    ' Keep the business error and expose a separate log failure status.
    On Error Resume Next
    Err.Clear
    DH_Log operation, wb, sheetName, address, 0, "失败：" & message, backup
    If Err.Number <> 0 Then DH_LastLogError = "失败记录未能保存：" & Err.Description
    On Error GoTo 0
End Sub

Public Function DH_GenerateLogReport(ByVal wb As Workbook) As String
    DH_RequireWorkbook wb
    DH_Log "生成报告", wb, "", "", 0, "成功；本地TSV；不创建工作表"
    DH_GenerateLogReport = DH_LogPath()
End Function
