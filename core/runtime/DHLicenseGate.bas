Attribute VB_Name = "DHLicenseGate"
Option Explicit

' Open Source Edition (MIT License): all features unlocked with no usage limits.

Public Function DHL_CanonicalActions(ByVal actions As String) As String
    Dim item As Variant, key As String, result As String
    For Each item In Split(actions, "|")
        key = CStr(item)
        If Left$(key, 10) <> "Transform:" Then key = Split(key, ":")(0)
        If Len(result) > 0 Then result = result & "|"
        result = result & key
    Next item
    DHL_CanonicalActions = result
End Function

Public Sub DHL_RequirePlan(ByVal actions As String)
    ' No-op in Open Source Edition: all actions are unconditionally permitted.
End Sub

Public Function DHL_StatusText() As String
    DHL_StatusText = "开源免费版 (MIT)"
End Function

Public Function DHL_Status() As Variant
    DHL_Status = Array("LICENSED", 0, 0, False, "开源免费版 (MIT License)", False, "OSS")
End Function

Public Function DHL_Activate(ByVal code As String) As Variant
    DHL_Activate = Array(True, "开源免费版已解锁全部功能，无需激活码", DHL_Status())
End Function

Public Function DHL_AfterExecution(ByVal completed As Collection) As Variant
    DHL_AfterExecution = DHL_Status()
End Function
