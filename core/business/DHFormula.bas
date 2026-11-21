Attribute VB_Name = "DHFormula"
Option Explicit

' Heuristic review hints only. Never evaluate or rewrite a formula.
Public Function DH_FormulaRiskMask(ByVal formula As String) As Long
    Static ifErrorPattern As Object, lookupPattern As Object, numberPattern As Object
    Dim code As String, matches As Object, item As Object, position As Long, indexText As String
    If ifErrorPattern Is Nothing Then
        Set ifErrorPattern = CreateObject("VBScript.RegExp")
        ifErrorPattern.Pattern = "\bIFERROR\s*\(": ifErrorPattern.IgnoreCase = True
        Set lookupPattern = CreateObject("VBScript.RegExp")
        lookupPattern.Pattern = "\bVLOOKUP\s*\(": lookupPattern.IgnoreCase = True: lookupPattern.Global = True
        Set numberPattern = CreateObject("VBScript.RegExp")
        numberPattern.Pattern = "(^=|[*/+^\-])\s*[+\-]?([0-9]+(\.[0-9]*)?|\.[0-9]+)([Ee][+\-]?[0-9]+)?(?![A-Za-z0-9_.!:])"
    End If
    code = StripQuoted(formula)
    If ifErrorPattern.Test(code) Then DH_FormulaRiskMask = 1
    Set matches = lookupPattern.Execute(code)
    For Each item In matches
        position = item.FirstIndex + item.Length + 1
        indexText = ThirdArgument(code, position)
        If Len(indexText) > 0 Then
            If indexText Like String$(Len(indexText), "#") Then
                DH_FormulaRiskMask = DH_FormulaRiskMask Or 2: Exit For
            End If
        End If
    Next item
    If numberPattern.Test(code) Then DH_FormulaRiskMask = DH_FormulaRiskMask Or 4
End Function

Private Function StripQuoted(ByVal source As String) As String
    Dim i As Long, token As String, quote As String, result As String
    i = 1
    Do While i <= Len(source)
        token = Mid$(source, i, 1)
        If token = Chr$(34) Or token = "'" Then
            quote = token: result = result & "?": i = i + 1
            Do While i <= Len(source)
                If Mid$(source, i, 1) = quote Then
                    If Mid$(source, i + 1, 1) = quote Then
                        i = i + 2
                    Else
                        i = i + 1: Exit Do
                    End If
                Else
                    i = i + 1
                End If
            Loop
        Else
            result = result & token: i = i + 1
        End If
    Loop
    StripQuoted = UCase$(result)
End Function

Private Function ThirdArgument(ByVal code As String, ByVal start As Long) As String
    Dim i As Long, depth As Long, braces As Long, argument As Long, token As String, text As String
    argument = 1
    For i = start To Len(code)
        token = Mid$(code, i, 1)
        If token = "(" Then depth = depth + 1
        If token = "{" Then braces = braces + 1
        If token = "}" Then braces = braces - 1
        If token = ")" Then
            If depth = 0 Then Exit For
            depth = depth - 1
        End If
        If (token = "," Or token = ";") And depth = 0 And braces = 0 Then
            If argument = 3 Then Exit For
            argument = argument + 1
        ElseIf argument = 3 Then
            text = text & token
        End If
    Next i
    If argument = 3 Then ThirdArgument = Trim$(text)
End Function

Public Function DH_FormulaRiskText(ByVal mask As Long) As String
    Dim result As String
    If (mask And 1) <> 0 Then result = "IFERROR可能隐藏错误；"
    If (mask And 2) <> 0 Then result = result & "VLOOKUP固定列号，插列后需复核；"
    If (mask And 4) <> 0 Then result = result & "算术中存在硬编码数字，建议核对是否应参数化；"
    DH_FormulaRiskText = result & "仅为人工复核提示，不自动判错或修改公式"
End Function
