Attribute VB_Name = "DHSafety"
Option Explicit

Private Sub AddNote(ByVal note As Object, ByVal region As Range, ByVal rows As Object)
    Dim cell As Range
    Set cell = note.Parent
    If Not Application.Intersect(cell, region) Is Nothing Then rows(CStr(cell.Row)) = True
End Sub

Private Function SavedThreadedParts(ByVal wb As Workbook) As Boolean
    Dim file As LongPtr, opened As Boolean, bytes() As Byte, length As Long, remaining As Double, received As Long
    Dim message As String, code As Long, needle() As Byte, i As Long, matched As Long, value As Long, stage As String
    On Error GoTo Failed
    ' The binary .xls format cannot contain OOXML threaded-comment parts.
    If wb.FileFormat = xlExcel8 Then Exit Function
    ' Win32 Unicode path (N-3): VBA Open failed on names outside the ANSI code page.
    stage = "open": file = DH_FileOpen(wb.FullName, False): opened = True
    ReDim bytes(0 To 1): received = DH_FileRead(file, bytes)
    If received < 2 Then Err.Raise vbObjectError + 918, , "本机不支持线程批注API，且文件不是可检查的OOXML，无法确认批注安全。"
    If bytes(0) <> 80 Or bytes(1) <> 75 Then Err.Raise vbObjectError + 918, , "本机不支持线程批注API，且文件不是可检查的OOXML，无法确认批注安全。"
    DH_FileRewind file: remaining = DH_FileSize(file)
    needle = StrConv("threadedcomment", vbFromUnicode)
    Do While remaining > 0
        stage = "read": If remaining > 1048576 Then length = 1048576 Else length = CLng(remaining)
        ReDim bytes(0 To length - 1): received = DH_FileRead(file, bytes)
        If received = 0 Then Exit Do
        If received < length Then ReDim Preserve bytes(0 To received - 1): length = received
        stage = "scan"
        For i = 0 To UBound(bytes)
            value = bytes(i): If value >= 65 And value <= 90 Then value = value + 32
            If value = needle(matched) Then
                matched = matched + 1
                If matched > UBound(needle) Then SavedThreadedParts = True: Exit For
            Else
                matched = 0: If value = needle(0) Then matched = 1
            End If
        Next i
        If SavedThreadedParts Then Exit Do
        remaining = remaining - length
    Loop
    DH_FileClose file
    Exit Function
Failed:
    code = Err.Number: message = Err.Description
    On Error Resume Next
    If opened Then DH_FileClose file
    On Error GoTo 0
    Err.Raise code, "批注兼容检查", "批注文件检查失败(" & stage & ")：" & message
End Function

' Metadata collections, never a per-cell read over UsedRange.
Public Function DH_NotedRows(ByVal region As Range) As Object
    Dim rows As Object, note As Object, threaded As Object, code As Long, message As String
    Set rows = CreateObject("Scripting.Dictionary")
    For Each note In region.Parent.Comments: AddNote note, region, rows: Next note
    On Error Resume Next
    Err.Clear
    Set threaded = region.Parent.CommentsThreaded
    code = Err.Number: message = Err.Description
    On Error GoTo 0
    If code <> 0 Then
        If code = 438 Or code = 445 Or (code = 1004 And (InStr(message, "不支持") > 0 Or InStr(1, message, "support", vbTextCompare) > 0)) Then
            Set threaded = Nothing
            If SavedThreadedParts(region.Parent.Parent) Then Err.Raise vbObjectError + 918, , "此文件含线程批注，但本机Excel不支持读取其对象模型；已停止删除，请用支持的Excel保留/检查批注。"
        Else
            Err.Raise code, "批注检查", "无法完成线程批注检查：" & message
        End If
    End If
    If Not threaded Is Nothing Then
        For Each note In threaded: AddNote note, region, rows: Next note
    End If
    Set DH_NotedRows = rows
End Function

Private Sub KeepPiece(ByVal parts As Collection, ByVal ws As Worksheet, ByVal row1 As Long, ByVal col1 As Long, ByVal row2 As Long, ByVal col2 As Long)
    If row1 <= row2 And col1 <= col2 Then parts.Add ws.Range(ws.Cells(row1, col1), ws.Cells(row2, col2))
End Sub

Public Function DH_SubtractRange(ByVal source As Range, ByVal cut As Range) As Range
    Dim parts As New Collection, nextParts As Collection, part As Range, area As Range, overlap As Range, result As Range
    Dim bottom As Long, right As Long
    For Each part In source.Areas: parts.Add part: Next part
    For Each area In cut.Areas
        Set nextParts = New Collection
        For Each part In parts
            Set overlap = Application.Intersect(part, area)
            If overlap Is Nothing Then
                nextParts.Add part
            Else
                bottom = part.Row + part.Rows.Count - 1: right = part.Column + part.Columns.Count - 1
                KeepPiece nextParts, part.Parent, part.Row, part.Column, overlap.Row - 1, right
                KeepPiece nextParts, part.Parent, overlap.Row + overlap.Rows.Count, part.Column, bottom, right
                KeepPiece nextParts, part.Parent, overlap.Row, part.Column, overlap.Row + overlap.Rows.Count - 1, overlap.Column - 1
                KeepPiece nextParts, part.Parent, overlap.Row, overlap.Column + overlap.Columns.Count, overlap.Row + overlap.Rows.Count - 1, right
            End If
        Next part
        If nextParts.Count > DH_MAX_AREAS Then Err.Raise vbObjectError + 917, , "保留选区外标记过于碎片化，未清除。"
        Set parts = nextParts
    Next area
    For Each part In parts
        If result Is Nothing Then Set result = part Else Set result = Application.Union(result, part)
    Next part
    Set DH_SubtractRange = result
End Function

' Removing headers from the WRITE range also prevents Excel's automatic table
' header renaming from changing fields outside a cleanup plan.
Public Function DH_WithoutTableHeaders(ByVal region As Range) As Range
    Dim parts As New Collection, nextParts As Collection, part As Range, table As ListObject, overlap As Range
    Dim bottom As Long, right As Long, result As Range
    If region Is Nothing Then Exit Function
    For Each part In region.Areas: parts.Add part: Next part
    For Each table In region.Parent.ListObjects
        If table.ShowHeaders Then
            Set nextParts = New Collection
            For Each part In parts
                Set overlap = Application.Intersect(part, table.HeaderRowRange)
                If overlap Is Nothing Then
                    nextParts.Add part
                Else
                    bottom = part.Row + part.Rows.Count - 1: right = part.Column + part.Columns.Count - 1
                    KeepPiece nextParts, part.Parent, part.Row, part.Column, overlap.Row - 1, right
                    KeepPiece nextParts, part.Parent, overlap.Row + overlap.Rows.Count, part.Column, bottom, right
                    KeepPiece nextParts, part.Parent, overlap.Row, part.Column, overlap.Row + overlap.Rows.Count - 1, overlap.Column - 1
                    KeepPiece nextParts, part.Parent, overlap.Row, overlap.Column + overlap.Columns.Count, overlap.Row + overlap.Rows.Count - 1, right
                End If
            Next part
            Set parts = nextParts
        End If
    Next table
    For Each part In parts
        If result Is Nothing Then Set result = part Else Set result = Application.Union(result, part)
    Next part
    Set DH_WithoutTableHeaders = result
End Function

Private Function WithoutStrings(ByVal formula As String) As String
    Dim i As Long, quoted As Boolean, token As String
    For i = 1 To Len(formula)
        token = Mid$(formula, i, 1)
        If token = Chr$(34) Then
            quoted = Not quoted: WithoutStrings = WithoutStrings & " "
        ElseIf quoted Then
            WithoutStrings = WithoutStrings & " "
        Else
            WithoutStrings = WithoutStrings & token
        End If
    Next i
End Function

Private Function FullyRemoved(ByVal reference As Range, ByVal changed As Range) As Boolean
    Dim overlap As Range
    If Not reference.Parent Is changed.Parent Then Exit Function
    Set overlap = Application.Intersect(reference, changed)
    If Not overlap Is Nothing Then FullyRemoved = (overlap.CountLarge = reference.CountLarge)
End Function

Private Function MayShiftReference(ByVal reference As Range, ByVal changed As Range) As Boolean
    Dim area As Range
    If Not reference.Parent Is changed.Parent Then Exit Function
    For Each area In changed.Areas
        If reference.Column <= area.Column + area.Columns.Count - 1 And reference.Column + reference.Columns.Count - 1 >= area.Column Then
            If reference.Row + reference.Rows.Count - 1 >= area.Row Then MayShiftReference = True: Exit Function
        End If
    Next area
End Function

Private Function FormulaProblem(ByVal formula As String, ByVal origin As Worksheet, ByVal changed As Range, ByVal insertColumn As Boolean, ByVal names As Object) As String
    Static refs As Object, dynamic As Object, positional As Object, rowWithoutArgument As Object
    Dim code As String, riskCode As String, item As Object, sheetName As String, address As String, reference As Range, isRelevant As Boolean, ending As Long, key As Variant
    If refs Is Nothing Then
        Set refs = CreateObject("VBScript.RegExp"): refs.Global = True: refs.IgnoreCase = True
        refs.Pattern = "(^|[^A-Za-z0-9_.])(?:(?:'((?:[^']|'')+)'|([A-Za-z_\u4e00-\u9fff][A-Za-z0-9_.\u4e00-\u9fff]*))!)?(\$?[A-Z]{1,3}\$?[1-9][0-9]*(:\$?[A-Z]{1,3}\$?[1-9][0-9]*)?|\$?[1-9][0-9]*:\$?[1-9][0-9]*|\$?[A-Z]{1,3}:\$?[A-Z]{1,3})"
        Set dynamic = CreateObject("VBScript.RegExp"): dynamic.IgnoreCase = True
        dynamic.Pattern = "\b(INDIRECT|OFFSET)\s*\("
        Set positional = CreateObject("VBScript.RegExp"): positional.IgnoreCase = True
        positional.Pattern = "\b(VLOOKUP|HLOOKUP|INDEX|CHOOSECOLS|CHOOSEROWS|ADDRESS|COLUMN|ROW)\s*\("
        Set rowWithoutArgument = CreateObject("VBScript.RegExp")
        rowWithoutArgument.Global = True: rowWithoutArgument.IgnoreCase = True
        rowWithoutArgument.Pattern = "(^|[^A-Za-z0-9_.!])ROW\s*\(\s*\)"
    End If
    code = WithoutStrings(formula)
    ' Dynamic/text references cannot be proved safe by native reference adjustment.
    If dynamic.Test(code) Then FormulaProblem = "存在INDIRECT/OFFSET动态引用，不能证明结构修改安全": Exit Function
    If InStr(code, ":") > 0 And InStr(code, "!") > InStr(code, ":") Then
        FormulaProblem = "存在跨表范围/复合引用，结构依赖无法安全确定": Exit Function
    End If
    isRelevant = (insertColumn And origin Is changed.Parent)
    If Not isRelevant Then
        For Each key In names.Keys
            If InStr(1, code, CStr(key), vbTextCompare) > 0 Then isRelevant = True: Exit For
        Next key
    End If
    For Each item In refs.Execute(code)
        ending = item.FirstIndex + item.Length + 1
        If Mid$(code, ending, 1) <> "(" Then
            sheetName = item.SubMatches(1)
            If Len(sheetName) > 0 Then sheetName = Replace$(sheetName, "''", "'") Else sheetName = item.SubMatches(2)
            If Len(sheetName) = 0 Then sheetName = origin.Name
            If StrComp(sheetName, changed.Parent.Name, vbTextCompare) = 0 Then
                If insertColumn Then isRelevant = True
                If Not insertColumn Then
                    address = item.SubMatches(3): Set reference = Nothing
                    On Error Resume Next
                    Set reference = changed.Parent.Range(address)
                    On Error GoTo 0
                    If Not reference Is Nothing Then
                        If MayShiftReference(reference, changed) Then isRelevant = True
                        If FullyRemoved(reference, changed) Then FormulaProblem = "公式直接引用待删除单元格，会失去引用：" & address: Exit Function
                    End If
                End If
            End If
        End If
    Next item
    riskCode = code
    ' A column insertion cannot change the row of a no-argument ROW(). This
    ' exemption never applies to row deletion, ROW(reference), COLUMN or lookup.
    If insertColumn Then riskCode = rowWithoutArgument.Replace(riskCode, "$1")
    If isRelevant And positional.Test(riskCode) Then FormulaProblem = "存在查询/位置参数公式，结构修改可能返回不同业务字段或位置"
End Function

' Conservative guard, not a full Excel formula parser. Reject known dependency
' risks instead of silently changing semantics. No automatic formula rewriting.
Public Sub DH_PreflightStructure(ByVal wb As Workbook, ByVal changed As Range, ByVal insertColumn As Boolean)
    Dim ws As Worksheet, formulas As Range, area As Range, values As Variant, r As Long, c As Long, reason As String
    Dim total As Double, item As Name, reference As Range, relevantNames As Object, name As String
    Set relevantNames = CreateObject("Scripting.Dictionary")
    For Each item In wb.Names
        Set reference = Nothing
        On Error Resume Next
        Set reference = item.RefersToRange
        On Error GoTo 0
        If Not reference Is Nothing Then
            If reference.Parent Is changed.Parent Then
                If insertColumn Or MayShiftReference(reference, changed) Then
                    name = Mid$(item.Name, InStrRev(item.Name, "!") + 1): relevantNames(name) = True
                End If
            End If
        End If
    Next item
    For Each ws In wb.Worksheets
        Set formulas = Nothing
        On Error Resume Next
        Set formulas = ws.Cells.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
        If Not formulas Is Nothing Then
            total = total + formulas.CountLarge
            If total > DH_MAX_CELLS Then Err.Raise vbObjectError + 913, , "公式依赖预检超过100万格，结构修改已停止。"
            For Each area In formulas.Areas
                values = DH_Matrix(area, True)
                For r = 1 To UBound(values, 1)
                    For c = 1 To UBound(values, 2)
                        reason = FormulaProblem(CStr(values(r, c)), ws, changed, insertColumn, relevantNames)
                        If Len(reason) > 0 Then Err.Raise vbObjectError + 914, , reason & "；位置=" & ws.Name & "!" & area.Cells(r, c).Address(False, False) & _
                            "。原表未执行结构修改。可将待处理数据复制为值到新建空白工作簿，在副本中整理，再人工核对；不要直接覆盖原表公式。"
                    Next c
                Next r
            Next area
        End If
    Next ws
    For Each item In wb.Names
        reason = FormulaProblem(item.RefersTo, changed.Parent, changed, insertColumn, relevantNames)
        If Len(reason) > 0 Then Err.Raise vbObjectError + 914, , "名称依赖风险：" & item.Name & "；" & reason
        If Not insertColumn Then
            Set reference = Nothing
            On Error Resume Next
            Set reference = item.RefersToRange
            On Error GoTo 0
            If Not reference Is Nothing Then
                If FullyRemoved(reference, changed) Then Err.Raise vbObjectError + 914, , "名称引用待删除区域，已停止：" & item.Name
            End If
        End If
    Next item
End Sub


