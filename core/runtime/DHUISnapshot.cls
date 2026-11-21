VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "DHUISnapshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private book As Workbook, sheet As Worksheet, source As Range
Private fullName As String, address As String, structure As String
Private values As Variant, formulas As Variant, masks As Variant

Private Function StructureOf(ByVal ws As Worksheet) As String
    Dim area As Range, table As ListObject, result As String, merges As Collection
    Set merges = DH_MergedAreas(ws.UsedRange)
    For Each area In merges: result = result & area.Address & ";": Next area
    For Each table In ws.ListObjects: result = result & table.Name & ":" & table.Range.Address & ";": Next table
    StructureOf = result
End Function

Public Sub Capture(ByVal ws As Worksheet)
    Set sheet = ws: Set book = ws.Parent
    DH_RequireWorkbook book
    Set source = DH_ContentRange(ws)
    fullName = book.FullName: address = ""
    If Not source Is Nothing Then
        If source.CountLarge > DH_MAX_CELLS Then Err.Raise vbObjectError + 960, , "数据区超过100万格，请缩小数据后体检。"
        address = source.Address
        values = DH_Matrix(source): formulas = DH_Matrix(source, True): masks = DH_FormulaMask(source)
    End If
    structure = StructureOf(ws)
End Sub

Public Function Matches(ByVal requireActive As Boolean) As Boolean
    Dim current As Range, v As Variant, f As Variant, m As Variant, r As Long, c As Long
    On Error GoTo Invalid
    If book Is Nothing Or sheet Is Nothing Then Exit Function
    If StrComp(book.FullName, fullName, vbBinaryCompare) <> 0 Then Exit Function
    If requireActive Then
        If Not Application.ActiveWorkbook Is book Then Exit Function
        If Not Application.ActiveSheet Is sheet Then Exit Function
    End If
    Set current = DH_ContentRange(sheet)
    If current Is Nothing Then
        If Len(address) <> 0 Then Exit Function
    Else
        If current.Address <> address Then Exit Function
        v = DH_Matrix(current): f = DH_Matrix(current, True): m = DH_FormulaMask(current)
        For r = 1 To UBound(v, 1)
            For c = 1 To UBound(v, 2)
                If DH_ValueToken(v(r, c)) <> DH_ValueToken(values(r, c)) Then Exit Function
                If DH_ValueToken(f(r, c)) <> DH_ValueToken(formulas(r, c)) Then Exit Function
                If m(r, c) <> masks(r, c) Then Exit Function
            Next c
        Next r
    End If
    If StructureOf(sheet) <> structure Then Exit Function
    Matches = True
Invalid:
End Function

Public Property Get TargetSheet() As Worksheet
    Set TargetSheet = sheet
End Property
