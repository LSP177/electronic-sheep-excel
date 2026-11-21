Attribute VB_Name = "DHPlatform"
Option Explicit
Private Declare PtrSafe Function CreateFileW Lib "kernel32" (ByVal path As LongPtr, ByVal access As Long, ByVal sharing As Long, ByVal security As LongPtr, ByVal creation As Long, ByVal flags As Long, ByVal template As LongPtr) As LongPtr
Private Declare PtrSafe Function ReadFile Lib "kernel32" (ByVal file As LongPtr, ByRef data As Byte, ByVal requested As Long, ByRef received As Long, ByVal overlapped As LongPtr) As Long
Private Declare PtrSafe Function WriteFile Lib "kernel32" (ByVal file As LongPtr, ByRef data As Byte, ByVal requested As Long, ByRef written As Long, ByVal overlapped As LongPtr) As Long
Private Declare PtrSafe Function GetFileSizeEx Lib "kernel32" (ByVal file As LongPtr, ByRef size As Currency) As Long
Private Declare PtrSafe Function SetFilePointerEx Lib "kernel32" (ByVal file As LongPtr, ByVal distance As Currency, ByRef position As Currency, ByVal method As Long) As Long

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As LongPtr, ByVal source As LongPtr, ByVal size As LongPtr)
Private Declare PtrSafe Function CreateMutexW Lib "kernel32" (ByVal attributes As LongPtr, ByVal initialOwner As Long, ByVal name As LongPtr) As LongPtr
Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal handle As LongPtr, ByVal milliseconds As Long) As Long
Private Declare PtrSafe Function ReleaseMutex Lib "kernel32" (ByVal handle As LongPtr) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal handle As LongPtr) As Long
Private Declare PtrSafe Function CryptAcquireContextW Lib "advapi32" (ByRef provider As LongPtr, ByVal container As LongPtr, ByVal providerName As LongPtr, ByVal providerType As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptCreateHash Lib "advapi32" (ByVal provider As LongPtr, ByVal algorithm As Long, ByVal key As LongPtr, ByVal flags As Long, ByRef hash As LongPtr) As Long
Private Declare PtrSafe Function CryptHashData Lib "advapi32" (ByVal hash As LongPtr, ByRef data As Byte, ByVal count As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptGetHashParam Lib "advapi32" (ByVal hash As LongPtr, ByVal parameter As Long, ByRef data As Byte, ByRef count As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptDestroyHash Lib "advapi32" (ByVal hash As LongPtr) As Long
Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32" (ByVal provider As LongPtr, ByVal flags As Long) As Long

' IEEE bytes avoid the rounding performed by CStr(Double). Normalize +/- zero.
Public Function DH_ValueToken(ByVal value As Variant) As String
    Dim text As String, number As Double, bytes(0 To 7) As Byte, i As Long
    If VarType(value) = vbDouble Or VarType(value) = vbSingle Then
        number = CDbl(value)
        If number = 0 Then number = 0#
        CopyMemory VarPtr(bytes(0)), VarPtr(number), 8
        For i = 0 To 7: text = text & Right$("0" & Hex$(bytes(i)), 2): Next i
    ElseIf Not IsEmpty(value) Then
        text = CStr(value)
    End If
    DH_ValueToken = CStr(VarType(value)) & ":" & Len(text) & ":" & text & ";"
End Function

Public Function DH_Lock(ByVal key As String) As LongPtr
    Dim handle As LongPtr, result As Long, name As String
    name = "Local\ExcelDataHelper_" & key
    handle = CreateMutexW(0, 0, StrPtr(name))
    If handle = 0 Then Err.Raise vbObjectError + 910, , "无法建立电子羊表格进程锁，已停止。"
    result = WaitForSingleObject(handle, 10000)
    If result <> 0 And result <> &H80 Then
        CloseHandle handle
        Err.Raise vbObjectError + 911, , "其他Excel进程正在写入，请稍后重试；本次未取得锁。"
    End If
    DH_Lock = handle
End Function

Public Sub DH_Unlock(ByVal handle As LongPtr)
    If handle <> 0 Then ReleaseMutex handle: CloseHandle handle
End Sub


Public Function DH_FileHash(ByVal path As String) As String
    Dim provider As LongPtr, hash As LongPtr, file As LongPtr, opened As Boolean
    Dim bytes() As Byte, digest(0 To 31) As Byte, length As Long, remaining As Long, i As Long
    Dim message As String, code As Long
    On Error GoTo Failed
    If CryptAcquireContextW(provider, 0, 0, 24, &HF0000000) = 0 Then Err.Raise vbObjectError + 912, , "无法初始化SHA256。"
    If CryptCreateHash(provider, &H800C&, 0, 0, hash) = 0 Then Err.Raise vbObjectError + 912, , "无法创建SHA256。"
    ' Unicode path and equivalent read sharing: deny writers/deletion while hashing.
    file = CreateFileW(StrPtr(path), &H80000000, 1, 0, 3, &H80, 0)
    If file = -1 Then Err.Raise vbObjectError + 912, , "无法打开文件进行SHA256校验（Win32 " & Err.LastDllError & "）"
    opened = True
    ReDim bytes(0 To 1048575)
    Do
        If ReadFile(file, bytes(0), 1048576, length, 0) = 0 Then Err.Raise vbObjectError + 912, , "SHA256校验读取文件失败（Win32 " & Err.LastDllError & "）"
        If length = 0 Then Exit Do
        If CryptHashData(hash, bytes(0), length, 0) = 0 Then Err.Raise vbObjectError + 912, , "SHA256校验更新失败（Win32 " & Err.LastDllError & "）"
    Loop
    CloseHandle file: opened = False
    length = 32
    If CryptGetHashParam(hash, 2, digest(0), length, 0) = 0 Then Err.Raise vbObjectError + 912, , "SHA256计算失败。"
    For i = 0 To 31: DH_FileHash = DH_FileHash & Right$("0" & Hex$(digest(i)), 2): Next i
    CryptDestroyHash hash: CryptReleaseContext provider, 0
    Exit Function
Failed:
    message = Err.Description: code = Err.Number
    On Error Resume Next
    If opened Then CloseHandle file
    If hash <> 0 Then CryptDestroyHash hash
    If provider <> 0 Then CryptReleaseContext provider, 0
    On Error GoTo 0
    Err.Raise code, "文件校验", message
End Function


' N-3: Unicode-safe file access. VBA's Open statement converts paths to the
' system ANSI code page, so names outside it failed with error 52.
' writing=False: read, others may read/write (same as Open ... Shared).
' writing=True: append target, opened always, no sharing (same as Lock Read Write).
Public Function DH_FileOpen(ByVal path As String, ByVal writing As Boolean) As LongPtr
    Dim handle As LongPtr
    If writing Then
        handle = CreateFileW(StrPtr(path), &H40000000, 0, 0, 4, &H80, 0)
    Else
        handle = CreateFileW(StrPtr(path), &H80000000, 7, 0, 3, &H80, 0)
    End If
    If handle = -1 Then Err.Raise vbObjectError + 913, , "无法打开文件（Win32 " & Err.LastDllError & ")"
    DH_FileOpen = handle
End Function

Public Function DH_FileSize(ByVal handle As LongPtr) As Double
    Dim size As Currency
    If GetFileSizeEx(handle, size) = 0 Then Err.Raise vbObjectError + 913, , "无法读取文件大小（Win32 " & Err.LastDllError & ")"
    DH_FileSize = CDbl(size) * 10000#
End Function

' Reads up to UBound+1 bytes into an already sized buffer; returns bytes read.
Public Function DH_FileRead(ByVal handle As LongPtr, ByRef bytes() As Byte) As Long
    Dim received As Long
    If ReadFile(handle, bytes(0), UBound(bytes) + 1, received, 0) = 0 Then Err.Raise vbObjectError + 913, , "无法读取文件（Win32 " & Err.LastDllError & ")"
    DH_FileRead = received
End Function

Public Sub DH_FileRewind(ByVal handle As LongPtr)
    Dim position As Currency
    If SetFilePointerEx(handle, 0, position, 0) = 0 Then Err.Raise vbObjectError + 913, , "无法定位文件（Win32 " & Err.LastDllError & ")"
End Sub

Public Sub DH_FileAppend(ByVal handle As LongPtr, ByRef bytes() As Byte)
    Dim position As Currency, written As Long
    If SetFilePointerEx(handle, 0, position, 2) = 0 Then Err.Raise vbObjectError + 913, , "无法定位文件（Win32 " & Err.LastDllError & ")"
    If WriteFile(handle, bytes(0), UBound(bytes) + 1, written, 0) = 0 Then Err.Raise vbObjectError + 913, , "无法写入文件（Win32 " & Err.LastDllError & ")"
    If written <> UBound(bytes) + 1 Then Err.Raise vbObjectError + 913, , "文件写入不完整"
End Sub

Public Sub DH_FileClose(ByRef handle As LongPtr)
    Dim owned As LongPtr
    owned = handle: handle = 0
    If owned <> 0 And owned <> -1 Then CloseHandle owned
End Sub
