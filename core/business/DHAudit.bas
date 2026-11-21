Attribute VB_Name = "DHAudit"
Option Explicit

Public Sub DH_RequireSelection(ByVal selected As Range, Optional ByVal needReport As Boolean = False)
    If selected Is Nothing Then Err.Raise vbObjectError + 770, , "请先选择单元格区域。"
    DH_RequireWorkbook selected.Parent.Parent, needReport
    If selected.Parent.ProtectContents Then Err.Raise vbObjectError + 771, , "工作表受保护，已停止操作。"
    If Application.ActiveWindow Is Nothing Then Err.Raise vbObjectError + 772, , "没有活动窗口。"
    If Application.ActiveWindow.SelectedSheets.Count <> 1 Then Err.Raise vbObjectError + 773, , "请先取消工作表分组。"
    If selected.CountLarge > DH_MAX_CELLS Or selected.Areas.Count > DH_MAX_AREAS Then _
        Err.Raise vbObjectError + 774, , "选区超过100万单元格或过于碎片化，请分批处理。"
End Sub

' Read-only entry checks before asking for consent. Preserve existing cleanup
' and masking support for whole-column selections by bounding actual content.
Public Sub DH_RequireEntrySelection(ByVal selected As Range, Optional ByVal boundToContent As Boolean = False, _
    Optional ByVal skipFirstRow As Boolean = False)
    Dim bounds As Range, area As Range, clipped As Range, total As Double
    If Not boundToContent Then
        DH_RequireSelection selected
        Exit Sub
    End If
    If selected Is Nothing Then Err.Raise vbObjectError + 770, , "请先选择单元格区域。"
    DH_RequireSelection selected.Cells(1, 1)
    If selected.Areas.Count > DH_MAX_AREAS Then Err.Raise vbObjectError + 774, , "选区过于碎片化，请分批处理。"
    Set bounds = DH_ContentRange(selected.Parent)
    If bounds Is Nothing Then Exit Sub
    If skipFirstRow Then Set bounds = Application.Intersect(bounds, selected.Parent.Rows("2:" & selected.Parent.Rows.Count))
    If bounds Is Nothing Then Exit Sub
    For Each area In selected.Areas
        Set clipped = Application.Intersect(area, bounds)
        If Not clipped Is Nothing Then total = total + clipped.CountLarge
        If total > DH_MAX_CELLS Then Err.Raise vbObjectError + 774, , "实际数据区域超过100万单元格，请分批处理。"
    Next area
End Sub

Public Function DH_TextBlocks(ByVal selected As Range) As Collection
    Dim blocks As Collection, bounds As Range, area As Range, clipped As Range, texts As Range, part As Range
    Set blocks = New Collection
    Set bounds = DH_ContentRange(selected.Parent)
    If bounds Is Nothing Then Set DH_TextBlocks = blocks: Exit Function
    For Each area In selected.Areas
        Set clipped = Application.Intersect(area, bounds)
        If Not clipped Is Nothing Then
            Set texts = Nothing
            If clipped.CountLarge = 1 Then
                If Not clipped.HasFormula Then
                    If VarType(clipped.Value2) = vbString Then Set texts = clipped
                End If
            Else
                On Error Resume Next
                Set texts = clipped.SpecialCells(xlCellTypeConstants, xlTextValues)
                On Error GoTo 0
                If Not texts Is Nothing Then Set texts = Application.Intersect(clipped, texts)
            End If
            If Not texts Is Nothing Then
                Set texts = DH_WithoutTableHeaders(texts)
                If texts Is Nothing Then GoTo NextTextArea
                For Each part In texts.Areas
                    blocks.Add part
                    If blocks.Count > DH_MAX_AREAS Then Err.Raise vbObjectError + 775, , "文本常量过于碎片化，请分批处理。"
                Next part
            End If
        End If
NextTextArea:
    Next area
    Set DH_TextBlocks = blocks
End Function

Private Function AuditField(ByVal value As Variant) As String
    Dim text As String
    If IsError(value) Then text = "错误值" Else text = CStr(value)
    text = Replace$(text, "\", "\\")
    text = Replace$(text, vbTab, "\t")
    text = Replace$(text, vbCr, "\r")
    text = Replace$(text, vbLf, "\n")
    If Len(text) > 0 Then
        If InStr("=+-@", Left$(text, 1)) > 0 Then text = "'" & text
    End If
    AuditField = text
End Function

Public Function DH_SaveAudit(ByVal operation As String, ByVal wb As Workbook, ByVal details As Collection) As String
    Dim folder As String, path As String, fso As Object, stream As Object, lines() As String
    Dim i As Long, j As Long, item As Variant, fields(0 To 5) As String
    folder = Environ$("LOCALAPPDATA") & "\ExcelDataHelper\details\" & Format$(Date, "yyyymmdd")
    DH_EnsureFolder folder
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = folder & "\" & DH_WorkbookKey(wb) & "_" & Format$(Now, "hhnnss") & "_" & fso.GetTempName() & ".tsv"
    ReDim lines(0 To details.Count + 1)
    lines(0) = "计划明细(v0.4)：" & AuditField(operation) & "；最终结果以操作日志的成功/失败记录为准"
    lines(1) = "位置" & vbTab & "计划" & vbTab & "原因" & vbTab & "原值" & vbTab & "目标值" & vbTab & "剥离单位"
    i = 2
    ' Collection.Item(index) is linear; enumerate once for large audits.
    For Each item In details
        For j = 0 To 5: fields(j) = AuditField(item(j)): Next j
        lines(i) = Join(fields, vbTab)
        i = i + 1
    Next item
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.WriteText Join(lines, vbCrLf)
    stream.SaveToFile path, 1
    stream.Close
    DH_SaveAudit = path
    DH_LastAuditPath = path
End Function

Public Function DH_ActionReport(ByVal wb As Workbook, ByVal operation As String, ByVal address As String, _
    ByVal summary As String, ByVal details As Collection, ByVal backup As String, ByVal auditPath As String) As Worksheet
    ' v0.4 compatibility seam: callers retain the interface, never create sheets.
    ' Full details were already persisted by DH_SaveAudit before mutation.
    Set DH_ActionReport = Nothing
    DH_LastSummary = summary
End Function

Public Function DH_FormulaMask(ByVal rng As Range) As Variant
    Dim mask() As Boolean, formulas As Range, part As Range, r As Long, c As Long
    ReDim mask(1 To rng.Rows.Count, 1 To rng.Columns.Count)
    If rng.CountLarge = 1 Then
        mask(1, 1) = rng.HasFormula
    Else
        On Error Resume Next
        Set formulas = rng.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
        If Not formulas Is Nothing Then Set formulas = Application.Intersect(formulas, rng)
        If Not formulas Is Nothing Then
            For Each part In formulas.Areas
                For r = part.Row To part.Row + part.Rows.Count - 1
                    For c = part.Column To part.Column + part.Columns.Count - 1
                        mask(r - rng.Row + 1, c - rng.Column + 1) = True
                    Next c
                Next r
            Next part
        End If
    End If
    DH_FormulaMask = mask
End Function
