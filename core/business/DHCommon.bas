Attribute VB_Name = "DHCommon"
Option Explicit

Public Const DH_MAX_CELLS As Long = 1000000
Public Const DH_MAX_AREAS As Long = 10000
Public Const DH_MAX_DETAILS As Long = 1000
Public DH_LastAuditPath As String
Public DH_LastSummary As String
Public Type DH_AppState
    Screen As Boolean
    Events As Boolean
    Calculation As XlCalculation
    Status As Variant
End Type

Public Sub DH_Freeze(ByRef state As DH_AppState)
    state.Screen = Application.ScreenUpdating
    state.Events = Application.EnableEvents
    state.Calculation = Application.Calculation
    state.Status = Application.StatusBar
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "电子羊表格正在处理，请稍候…"
    Exit Sub
Failed:
    Dim code As Long, message As String
    code = Err.Number: message = Err.Description
    DH_Restore state
    Err.Raise code, "应用状态", message
End Sub

Public Sub DH_Restore(ByRef state As DH_AppState)
    On Error Resume Next
    Application.Calculation = state.Calculation
    Application.EnableEvents = state.Events
    Application.ScreenUpdating = state.Screen
    Application.StatusBar = state.Status
    On Error GoTo 0
End Sub

Public Sub DH_RequireWorkbook(ByVal wb As Workbook, Optional ByVal needSheet As Boolean = False)
    If wb Is Nothing Then Err.Raise vbObjectError + 701, , "请先打开数据工作簿。"
    If wb Is ThisWorkbook Or wb.IsAddin Then Err.Raise vbObjectError + 702, , "不能处理插件自身。"
    If Len(wb.Path) = 0 Then Err.Raise vbObjectError + 703, , "文件尚未保存，请先保存后再操作。"
    If InStr(1, wb.Path, "://", vbTextCompare) > 0 Then Err.Raise vbObjectError + 704, , "请先将文件保存到本地目录。"
    If wb.ReadOnly Then Err.Raise vbObjectError + 705, , "文件为只读，请先保存可编辑副本。"
    If needSheet And wb.ProtectStructure Then Err.Raise vbObjectError + 706, , "工作簿结构受保护，无法创建报告。"
    Select Case wb.FileFormat
        Case xlOpenXMLWorkbook, xlOpenXMLWorkbookMacroEnabled, xlExcel12, xlExcel8, 61
        Case Else
            Err.Raise vbObjectError + 707, , "请先另存为 .xlsx、.xlsm、.xlsb 或 .xls 工作簿。"
    End Select
End Sub

Public Function DH_Backup(ByVal wb As Workbook) As String
    Dim fso As Object, folder As String, target As String, pending As String, lockPath As String
    Dim manifest As String, handle As Object, owned As Boolean, message As String, code As Long, pendingCreated As Boolean
    DH_RequireWorkbook wb
    DH_RequireBackupCompatible wb
    On Error GoTo Failed
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = DH_BackupRoot() & "\" & DH_WorkbookKey(wb)
    DH_EnsureFolder folder
    manifest = folder & "\source.txt"
    If fso.FileExists(manifest) Then
        Set handle = fso.OpenTextFile(manifest, 1, False, -1)
        message = handle.ReadAll: handle.Close
        If StrComp(message, fso.GetAbsolutePathName(wb.FullName), vbTextCompare) <> 0 Then _
            Err.Raise vbObjectError + 708, , "备份标识冲突，已阻止修改。"
    Else
        Set handle = fso.CreateTextFile(manifest, False, True)
        handle.Write fso.GetAbsolutePathName(wb.FullName): handle.Close
    End If
    target = DH_DailyBackupPath(wb)
    DH_EnsureFolder fso.GetParentFolderName(target)
    If fso.FileExists(target) Then
        If fso.GetFile(target).Size = 0 Then Err.Raise vbObjectError + 709, , "当日基线为空，已阻止修改。"
        DH_VerifyBaseline wb, target
        DH_Backup = target: Exit Function
    End If
    lockPath = target & ".lock"
    Set handle = fso.CreateTextFile(lockPath, False, True)
    owned = True: handle.Write wb.FullName: handle.Close
    pending = fso.GetParentFolderName(target) & "\baseline.pending." & fso.GetExtensionName(wb.Name)
    If fso.FileExists(pending) Then Err.Raise vbObjectError + 710, , "发现未完成备份，请检查备份目录后再试。"
    pendingCreated = True
    wb.SaveCopyAs pending
    If Not fso.FileExists(pending) Then Err.Raise vbObjectError + 711, , "备份未生成，已阻止修改。"
    If fso.GetFile(pending).Size = 0 Then Err.Raise vbObjectError + 712, , "备份为空，已阻止修改。"
    DH_VerifyBaseline wb, pending, True
    fso.MoveFile pending, target
    fso.DeleteFile lockPath
    owned = False
    DH_Backup = target
    Exit Function
Failed:
    message = Err.Description: code = Err.Number
    On Error Resume Next
    If owned Then
        If Not handle Is Nothing Then handle.Close
        If pendingCreated Then
            If fso.FileExists(pending) Then fso.DeleteFile pending
        End If
        If fso.FileExists(lockPath) Then fso.DeleteFile lockPath
    End If
    On Error GoTo 0
    Err.Raise code, "电子羊表格备份", "备份失败，未允许破坏性操作：" & message
End Function

Public Sub DH_EnsureFolder(ByVal path As String)
    Dim fso As Object, parent As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(path) Then Exit Sub
    parent = fso.GetParentFolderName(path)
    If Len(parent) = 0 Then Err.Raise vbObjectError + 713, , "无法建立目录：" & path
    If Not fso.FolderExists(parent) Then DH_EnsureFolder parent
    fso.CreateFolder path
End Sub

Public Function DH_BackupRoot() As String
    DH_BackupRoot = Environ$("LOCALAPPDATA") & "\ExcelDataHelper\backups"
End Function

Public Function DH_WorkbookKey(ByVal wb As Workbook) As String
    Dim path As String, a As Double, b As Double, n As Long, code As Long
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = LCase$(fso.GetAbsolutePathName(wb.FullName))
    a = 5381#: b = 7919#
    For n = 1 To Len(path)
        code = AscW(Mid$(path, n, 1)): If code < 0 Then code = code + 65536
        a = a * 31# + code: a = a - Fix(a / 2147483647#) * 2147483647#
        b = b * 131# + code: b = b - Fix(b / 2147483629#) * 2147483629#
    Next n
    DH_WorkbookKey = Hex$(CLng(a)) & "-" & Hex$(CLng(b))
End Function

Public Function DH_DailyBackupPath(ByVal wb As Workbook) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    DH_DailyBackupPath = DH_BackupRoot() & "\" & DH_WorkbookKey(wb) & "\" & Format$(Date, "yyyymmdd") & _
        "\" & Left$(fso.GetBaseName(wb.Name), 70) & "_当日基线." & fso.GetExtensionName(wb.Name)
End Function

Public Function DH_BackupFolderPath(ByVal wb As Workbook) As String
    Dim fso As Object, candidate As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    DH_BackupFolderPath = DH_BackupRoot()
    If wb Is Nothing Then Exit Function
    If wb Is ThisWorkbook Or wb.IsAddin Or Len(wb.Path) = 0 Then Exit Function
    If InStr(1, wb.Path, "://", vbTextCompare) > 0 Then Exit Function
    candidate = fso.GetParentFolderName(DH_DailyBackupPath(wb))
    If fso.FolderExists(candidate) Then DH_BackupFolderPath = candidate
End Function

Public Function DH_OriginalSelection(ByVal ws As Worksheet) As Range
    On Error Resume Next
    If TypeOf Application.Selection Is Range Then
        If Application.Selection.Parent Is ws Then Set DH_OriginalSelection = Application.Selection
    End If
    On Error GoTo 0
End Function

Public Function DH_ClipColumn(ByVal selected As Range) As Range
    Dim first As Range, last As Range
    If selected Is Nothing Then Err.Raise vbObjectError + 810, , "请先选择一列。"
    If selected.Areas.Count <> 1 Or selected.Columns.Count <> 1 Then Err.Raise vbObjectError + 811, , "请使用单列、连续的选区。"
    Set first = selected.Find(What:="*", After:=selected.Cells(selected.Rows.Count, 1), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    If first Is Nothing Then Exit Function
    Set last = selected.Find(What:="*", After:=selected.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    Set DH_ClipColumn = selected.Parent.Range(first, last)
End Function

Public Sub DH_RestoreFocus(ByVal ws As Worksheet, Optional ByVal selected As Range = Nothing)
    On Error Resume Next
    If Not ws Is Nothing Then
        ws.Activate
        If Not selected Is Nothing Then selected.Select
    End If
    On Error GoTo 0
End Sub

' Confirmation labels use captured Range/Workbook objects, never a later
' Selection lookup. More than three areas are explicitly summarized.
Public Function DH_SelectionTarget(ByVal selected As Range, Optional ByVal label As String = "当前将处理") As String
    Dim address As String
    If selected Is Nothing Then Err.Raise vbObjectError + 770, , "请先选择单元格区域。"
    If selected.Areas.Count <= 3 Then
        address = selected.Address(False, False)
    Else
        address = selected.Areas(1).Address(False, False) & " 等共" & selected.Areas.Count & "个区域（仅展示首区，请核对Excel选区）"
    End If
    DH_SelectionTarget = label & "：工作簿《" & selected.Parent.Parent.Name & "》" & vbCrLf & _
        "'" & Replace$(selected.Parent.Name, "'", "''") & "'!" & address & vbCrLf
End Function

Public Function DH_WorkbookTarget(ByVal wb As Workbook, ByVal scope As String) As String
    If wb Is Nothing Then Err.Raise vbObjectError + 701, , "请先打开数据工作簿。"
    DH_WorkbookTarget = "当前将处理：工作簿《" & wb.Name & "》" & vbCrLf & "范围：" & scope & vbCrLf
End Function

Private Function ReportOwner(ByVal ws As Worksheet) As String
    ReportOwner = "ExcelDataHelper.Report.v043|" & DH_WorkbookKey(ws.Parent) & "|" & ws.CodeName
End Function

Public Sub DH_TagReport(ByVal ws As Worksheet)
    Dim item As Object
    For Each item In ws.CustomProperties
        If item.Name = "ExcelDataHelperReport" Then
            If CStr(item.Value) = ReportOwner(ws) Then Exit Sub
            Err.Raise vbObjectError + 916, , "报告归属属性冲突，原工作表保留。"
        End If
    Next item
    ws.CustomProperties.Add "ExcelDataHelperReport", ReportOwner(ws)
End Sub

Public Function DH_IsReportSheet(ByVal ws As Worksheet) As Boolean
    Dim item As Object
    For Each item In ws.CustomProperties
        If item.Name = "ExcelDataHelperReport" Then
            DH_IsReportSheet = (CStr(item.Value) = ReportOwner(ws))
            Exit Function
        End If
    Next item
    ' Titles, names and legacy value=1 never authorize clearing/deleting a sheet.
End Function

Public Function DH_ExistingBaseline(ByVal wb As Workbook) As String
    Dim fso As Object, path As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = DH_DailyBackupPath(wb)
    If fso.FileExists(path) Then DH_ExistingBaseline = path
End Function

Public Sub DH_BackupLink(ByVal target As Range, ByVal wb As Workbook)
    Dim fso As Object, folder As String, label As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = DH_BackupFolderPath(wb)
    If Not fso.FolderExists(folder) Then Exit Sub
    label = CStr(target.Value2)
    target.Parent.Hyperlinks.Add Anchor:=target, Address:=folder & "\", TextToDisplay:=label, ScreenTip:="打开备份目录"
End Sub

' Explicit Find arguments avoid inheriting the user's last Find-dialog settings.
Public Function DH_ContentRange(ByVal ws As Worksheet) As Range
    Dim firstRow As Range, lastRow As Range, firstCol As Range, lastCol As Range
    Set firstRow = ws.Cells.Find(What:="*", After:=ws.Cells(ws.Rows.Count, ws.Columns.Count), _
        LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    If firstRow Is Nothing Then Exit Function
    Set lastRow = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    Set firstCol = ws.Cells.Find(What:="*", After:=ws.Cells(ws.Rows.Count, ws.Columns.Count), _
        LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    Set lastCol = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    Set DH_ContentRange = ws.Range(ws.Cells(firstRow.Row, firstCol.Column), ws.Cells(lastRow.Row, lastCol.Column))
End Function

Public Function DH_Matrix(ByVal rng As Range, Optional ByVal formulas As Boolean = False) As Variant
    Dim data As Variant, oneCellData(1 To 1, 1 To 1) As Variant
    If formulas Then data = rng.Formula Else data = rng.Value2
    If rng.CountLarge = 1 Then
        oneCellData(1, 1) = data
        DH_Matrix = oneCellData
    Else
        DH_Matrix = data
    End If
End Function

Public Function DH_CleanText(ByVal value As String) As String
    Dim words As Variant, i As Long
    value = Replace$(value, ChrW(12288), " ")
    value = Replace$(value, ChrW(160), " ")
    value = Replace$(value, vbTab, " ")
    value = Replace$(value, vbCr, " ")
    value = Replace$(value, vbLf, " ")
    value = Trim$(value)
    Do While InStr(value, "  ") > 0: value = Replace$(value, "  ", " "): Loop
    If Len(value) = 0 Then Exit Function
    words = Split(value, " ")
    For i = 1 To UBound(words)
        If Not (DH_IsChinese(Right$(words(i - 1), 1)) And DH_IsChinese(Left$(words(i), 1))) Then words(i) = " " & words(i)
    Next i
    DH_CleanText = Join(words, vbNullString)
End Function

Private Function DH_IsChinese(ByVal character As String) As Boolean
    Dim code As Long
    If Len(character) = 0 Then Exit Function
    code = AscW(character): If code < 0 Then code = code + 65536
    DH_IsChinese = (code >= &H3400 And code <= &H4DBF) Or (code >= &H4E00 And code <= &H9FFF&) Or (code >= &HF900& And code <= &HFAFF&)
End Function

Public Function DH_CharacterCount(ByVal value As String) As Long
    Dim length As Long
    length = Len(value)
    value = Replace$(value, " ", vbNullString)
    value = Replace$(value, ChrW(12288), vbNullString)
    value = Replace$(value, ChrW(160), vbNullString)
    value = Replace$(value, vbTab, vbNullString)
    value = Replace$(value, vbCr, vbNullString)
    value = Replace$(value, vbLf, vbNullString)
    DH_CharacterCount = length - Len(value)
End Function

' Apostrophe is Excel's literal-text escape. It is not included in Value2.
' Escape ALL nonempty strings in blocks, including unchanged numeric-looking text.
Public Function DH_LiteralMatrix(ByVal data As Variant) As Variant
    Dim r As Long, c As Long
    For r = LBound(data, 1) To UBound(data, 1)
        For c = LBound(data, 2) To UBound(data, 2)
            If VarType(data(r, c)) = vbString Then
                If Len(data(r, c)) > 0 Then data(r, c) = "'" & data(r, c)
            End If
        Next c
    Next r
    DH_LiteralMatrix = data
End Function

Public Function DH_UniqueSheetName(ByVal wb As Workbook, ByVal base As String) As String
    Dim candidate As String, n As Long, existing As Object, found As Boolean
    candidate = base
    Do
        found = False
        For Each existing In wb.Sheets
            If StrComp(existing.Name, candidate, vbTextCompare) = 0 Then found = True: Exit For
        Next existing
        If Not found Then Exit Do
        n = n + 1
        candidate = Left$(base, 25) & "_" & CStr(n)
    Loop
    DH_UniqueSheetName = candidate
End Function

Public Sub DH_ReportStyle(ByVal ws As Worksheet, ByVal lastRow As Long)
    With ws.Range("A1:D" & lastRow)
        .Font.Name = "微软雅黑"
        .Font.Size = 10
        .VerticalAlignment = xlCenter
        .RowHeight = 22
    End With
    With ws.Range("A1:D1")
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 120)
        .RowHeight = 30
    End With
    ws.Columns("A").ColumnWidth = 24
    ws.Columns("B").ColumnWidth = 48
    ws.Columns("C").ColumnWidth = 26
    ws.Columns("D").ColumnWidth = 55
    ws.Range("A2:D" & lastRow).WrapText = True
    ws.Range("A2:D" & lastRow).Rows.AutoFit
End Sub
