Attribute VB_Name = "DHPoc2Bridge"
Option Explicit

' DataHelper 1.0 VBA command & safety bridge. Business algorithms remain unchanged.
Private sourceBook As Workbook, sourceSheet As Worksheet, sourceRange As Range
Private sourceGuard As DHUISnapshot, auxSheet As Worksheet, auxGuard As DHUISnapshot

Private auxSheets As Collection, auxGuards As Collection
Private events As DHPoc2Events
Private sessionId As String, expires As Date, confirmed As Boolean, invalidReason As String
Private counter As Long
Private planReady As Boolean
Private autoNote As String, executing As Boolean
Private lockedActions As String

' Only this boundary is called by the VSTO command bridge. Never let VBA's
' unhandled error dialog escape into an in-process Office callback.
Public Function DHPoc2_Invoke(ByVal operation As String, Optional ByVal argument As Variant) As Variant
    Dim result As Variant, failureCode As Long, failureMessage As String
    On Error GoTo Failed
    Select Case operation
        Case "DHPoc2_LicenseStatus": result = DHL_Status()
        Case "DHPoc2_LicenseActivate": result = DHL_Activate(argument)
        Case "DHPoc2_Context": result = DHPoc2_Context()
        Case "DHPoc2_Candidate": result = DHPoc2_Candidate()
        Case "DHPoc2_Scan": result = DHPoc2_Scan()
        Case "DHPoc2_Workbook": result = DHPoc2_Workbook()
        Case "DHPoc2_SelectRange": result = DHPoc2_SelectRange(CStr(argument))
        Case "DHPoc2_LookupMeta": result = DHPoc2_LookupMeta()
        Case "DHPoc2_LookupPreview": result = DHPoc2_LookupPreview(CStr(argument))

        Case "DHPoc2_AppendPreview": result = DHPoc2_AppendPreview(CStr(argument))

        Case "DHPoc2_UnpivotPreview": result = DHPoc2_UnpivotPreview(CStr(argument))

        Case "DHPoc2_SplitSheetPreview": result = DHPoc2_SplitSheetPreview(CStr(argument))
        Case "DHPoc2_Prepare": result = DHPoc2_Prepare(CStr(argument))
        Case "DHPoc2_ConfirmationTarget": result = DHPoc2_ConfirmationTarget()
        Case "DHPoc2_Confirm": result = DHPoc2_Confirm(CStr(argument))
        Case "DHPoc2_Execute": If Not IsMissing(argument) Then Err.Raise 5, , "Execute takes no client parameters"
            result = DHPoc2_Execute()
        Case "DHPoc2_Invalidate": DHPoc2_Invalidate CStr(argument): result = "INVALIDATED"
        Case Else: Err.Raise vbObjectError + 983, , "Bridge operation not allowed"
    End Select
    DHPoc2_Invoke = Array("ok", result)
    Exit Function
Failed:
    failureCode = Err.number: failureMessage = Err.Description
    If operation = "DHPoc2_Prepare" Or operation = "DHPoc2_Confirm" Or operation = "DHPoc2_Execute" Then
        executing = False
        DHPoc2_Invalidate "final validation or authorisation failed"
    End If
    DHPoc2_Invoke = Array("error", failureCode, failureMessage)
End Function

Public Sub DHPoc2_Invalidate(Optional ByVal reason As String = "context changed")
    If executing Then Exit Sub
    lockedActions = ""
    confirmed = False
    planReady = False
    Set auxSheet = Nothing: Set auxGuard = Nothing: Set auxSheets = Nothing: Set auxGuards = Nothing
    invalidReason = reason
End Sub

Public Sub DHPoc2_DataChanged(ByVal changedSheet As Object)
    If sourceSheet Is Nothing Then Exit Sub
    If changedSheet Is sourceSheet Then DHPoc2_Invalidate "source data changed"
    If Not auxSheet Is Nothing Then
        If changedSheet Is auxSheet Then DHPoc2_Invalidate "lookup/reconcile reference sheet changed"
    End If
    If Not auxSheets Is Nothing Then
        Dim sObj As Object
        For Each sObj In auxSheets
            If changedSheet Is sObj Then DHPoc2_Invalidate "append source worksheet changed": Exit Sub
        Next sObj
    End If
End Sub

Public Function DHPoc2_SelectionEvent(ByVal changedSheet As Object, ByVal target As Range) As Boolean
    Dim actual As Range
    On Error GoTo Changed
    If sourceSheet Is Nothing Or sourceRange Is Nothing Then Exit Function
    If TypeOf Application.Selection Is Range Then
        Set actual = Application.Selection
        If actual.parent Is sourceSheet Then
            If actual.address = sourceRange.address Then Exit Function
        End If
    End If
Changed:
    DHPoc2_Invalidate "selection changed"
    DHPoc2_SelectionEvent = True
End Function

Private Function ActiveDataSheet() As Worksheet
    If Not TypeOf Application.ActiveSheet Is Worksheet Then Err.Raise vbObjectError + 980, , "Worksheet required"
    Set ActiveDataSheet = Application.ActiveSheet
    DH_RequireWorkbook ActiveDataSheet.parent
    If DH_IsReportSheet(ActiveDataSheet) Or DH_IsPreviousReport(ActiveDataSheet) Then Err.Raise vbObjectError + 980, , "Business sheet required"
End Function

Public Function DHPoc2_Context() As Variant
    Dim ws As Worksheet, selected As Range
    Set ws = ActiveDataSheet()
    If Not TypeOf Application.Selection Is Range Then Err.Raise 5, , "Range required"
    Set selected = Application.Selection
    DHPoc2_Context = Array(ws.parent.name, ws.name, selected.address(False, False), selected.rows.count, selected.Columns.count, DH1_ProductVersion(), ws.parent.FullName)
End Function

Public Function DHPoc2_Candidate() As Variant
    Dim ws As Worksheet, bounds As Range
    Set ws = ActiveDataSheet(): Set bounds = DH_ContentRange(ws)
    DHPoc2_Invalidate "candidate changed"
    If bounds Is Nothing Then DHPoc2_Candidate = Array(ws.parent.name, ws.name, "", 0, 0, ws.parent.FullName): Exit Function
    bounds.Select
    DHPoc2_Candidate = Array(ws.parent.name, ws.name, bounds.address(False, False), bounds.rows.count, bounds.Columns.count, ws.parent.FullName)
End Function

Private Function StatsData(ByVal ws As Worksheet, ByRef stats As DH_HealthStats) As Variant
    StatsData = Array(ws.parent.name, ws.name, stats.address, stats.Characters, stats.DirtyCells, stats.NumericTexts, _
        stats.merges, stats.MergedCells, stats.BlankRows, stats.Duplicates, stats.formulas, stats.RiskCells, stats.DateCandidates, stats.ProtectedNumbers, HiddenColsOf(ws))
End Function

Public Function DHPoc2_Scan() As Variant
    Dim ws As Worksheet, stats As DH_HealthStats, details As New Collection
    Set ws = ActiveDataSheet()
    DH_ScanWorksheet ws, stats, details
    CaptureScan ws
    DHPoc2_Scan = StatsData(ws, stats)
End Function

Public Function DHPoc2_Workbook() As Variant
    Dim Wb As Workbook, ws As Worksheet, rows As New Collection, data() As Variant, n As Long
    Dim stats As DH_HealthStats, details As Collection
    Set Wb = Application.ActiveWorkbook: DH_RequireWorkbook Wb
    For Each ws In Wb.Worksheets
        If ws.visible = xlSheetVisible And Not DH_IsReportSheet(ws) And Not DH_IsPreviousReport(ws) Then
            Set details = New Collection
            DH_ScanWorksheet ws, stats, details
            rows.Add StatsData(ws, stats)
        End If
    Next ws
    If rows.count = 0 Then DHPoc2_Workbook = Array(): Exit Function
    ReDim data(0 To rows.count - 1)
    For n = 1 To rows.count: data(n - 1) = rows(n): Next n
    DHPoc2_Workbook = data
End Function

Public Function DHPoc2_LookupMeta() As Variant
    Dim ws As Worksheet
    Set ws = ActiveDataSheet()
    DHPoc2_LookupMeta = DH_LookupMeta(ws.parent)
End Function

Public Function DHPoc2_LookupPreview(ByVal spec As String) As Variant
    Dim ws As Worksheet
    Set ws = ActiveDataSheet()
    DHPoc2_LookupPreview = DH_LookupPreview(ws, spec)
End Function

Public Function DHPoc2_AppendPreview(ByVal spec As String) As Variant
    Dim ws As Worksheet
    Set ws = ActiveDataSheet()
    DHPoc2_AppendPreview = DH_AppendPreview(ws.parent, spec)
End Function

Public Function DHPoc2_UnpivotPreview(ByVal spec As String) As Variant
    Dim ws As Worksheet
    Set ws = ActiveDataSheet()
    DHPoc2_UnpivotPreview = DH_UnpivotPreview(ws, spec)
End Function

Public Function DHPoc2_SplitSheetPreview(ByVal spec As String) As Variant
    Dim ws As Worksheet
    Set ws = ActiveDataSheet()
    DHPoc2_SplitSheetPreview = DH_SplitSheetPreview(ws, spec)
End Function

Private Function CellData(ByVal value As Variant) As Variant
    If IsError(value) Then
        CellData = Array("error", DH_ValueToken(value))
    ElseIf IsNull(value) Then
        CellData = Array("null", "")
    ElseIf IsEmpty(value) Then
        CellData = Array("empty", "")
    Else
        CellData = Array("value", value)
    End If
End Function

Private Sub CaptureScan(ByVal ws As Worksheet)
    If Not TypeOf Application.Selection Is Range Then Err.Raise 5, , "Range required"
    Set sourceRange = Application.Selection
    ' Capture is read-only: bound validation to real content, keeping the original
    ' selection identity. Every write still validates its final locked range.
    DH_RequireEntrySelection sourceRange, True
    Set sourceSheet = ws: Set sourceBook = ws.parent
    Set sourceGuard = New DHUISnapshot: sourceGuard.Capture ws
    Set auxSheet = Nothing: Set auxGuard = Nothing: Set auxSheets = Nothing: Set auxGuards = Nothing
    If events Is Nothing Then Set events = New DHPoc2Events: Set events.App = Application
    counter = counter + 1
    sessionId = format$(Now, "yyyymmddhhnnss") & "-" & CStr(counter)
    expires = DateAdd("s", 300, Now): confirmed = False: invalidReason = ""
    planReady = False
End Sub

Private Sub BindAuxiliarySheet(ByVal actions As String)
    Dim parts As Variant, sheetName As String, ws As Worksheet, sNames As Variant, i As Long, g As DHUISnapshot
    Set auxSheet = Nothing: Set auxGuard = Nothing: Set auxSheets = Nothing: Set auxGuards = Nothing
    If left$(actions, 7) = "Lookup:" Or left$(actions, 10) = "Reconcile:" Then
        parts = Split(actions, ":")
        If UBound(parts) >= 1 Then sheetName = CStr(parts(1))
        For Each ws In sourceBook.Worksheets
            If StrComp(ws.name, sheetName, vbTextCompare) = 0 Then Set auxSheet = ws: Exit For
        Next ws
        If auxSheet Is Nothing Then Err.Raise 5, , "Reference worksheet not found: " & sheetName
        If Not auxSheet Is sourceSheet Then
            Set auxGuard = New DHUISnapshot
            auxGuard.Capture auxSheet
        End If
    ElseIf left$(actions, 7) = "Append:" Then
        parts = Split(actions, ":")
        If UBound(parts) >= 1 Then
            sNames = Split(CStr(parts(1)), ",")
            Set auxSheets = New Collection
            Set auxGuards = New Collection
            For i = LBound(sNames) To UBound(sNames)
                sheetName = Trim$(CStr(sNames(i)))
                Set ws = Nothing
                For Each ws In sourceBook.Worksheets
                    If StrComp(ws.name, sheetName, vbTextCompare) = 0 Then Exit For
                Next ws
                If ws Is Nothing Then Err.Raise 5, , "Reference worksheet not found: " & sheetName
                If Not ws Is sourceSheet Then
                    auxSheets.Add ws
                    Set g = New DHUISnapshot
                    g.Capture ws
                    auxGuards.Add g
                End If
            Next i
        End If
    End If
End Sub

Public Function DHPoc2_Prepare(ByVal actions As String) As Variant
    RequireCurrent
    ValidateActions actions
    DHL_RequirePlan actions
    autoNote = ""
    If Left$(actions, 10) = "Reconcile:" Then
        If sourceRange.Areas.Count <> 1 Or sourceRange.Columns.Count <> 1 Then _
            Err.Raise 5, , "两表对账请使用单列、连续的选区；可点击「帮我选好」后重新体检。"
        Dim clipped As Range, shown As Range, p2 As Variant, hasHdr As Boolean
        Set clipped = DH_ClipColumn(sourceRange)
        If clipped Is Nothing Then Err.Raise 5, , "所选列没有数据区，请重新选择比对列。"
        DH_RequireSelection clipped
        If clipped.CountLarge > 100000 Then Err.Raise 5, , "选区太大：对账数据区超过10万格，请分批处理。"
        p2 = Split(actions, ":")
        hasHdr = (CStr(p2(3)) = "true")
        Set shown = clipped
        If hasHdr Then
            If clipped.Rows.Count < 2 Then Err.Raise 5, , "比对列只有表头，请重新选择数据区。"
            Set shown = clipped.Offset(1).Resize(clipped.Rows.Count - 1)
        End If
        ' Keep the header in the locked range: DH_ReconcileColumns skips it once.
        ' Set the bound object before Select so selection events see the same range.
        Set sourceRange = clipped
        clipped.Select
        autoNote = "已按 " & shown.Address(False, False) & " 自动取数据区"
    ElseIf actions = "Transform:Name" Or actions = "Transform:Phone" Or actions = "Transform:ID" Then
        ' Masking already clips constants and protects headers in DH_TransformSelection.
        ' Bound the bridge plan first so whole-column selections reach that engine.
        If sourceRange.Areas.Count <> 1 Then Err.Raise 5, , "脱敏请使用连续的选区。"
        Dim maskBounds As Range, maskRange As Range
        If sourceRange.Columns.Count = 1 Then
            Set maskRange = DH_ClipColumn(sourceRange)
        Else
            Set maskBounds = DH_ContentRange(sourceSheet)
            If Not maskBounds Is Nothing Then Set maskRange = Application.Intersect(sourceRange, maskBounds)
        End If
        If maskRange Is Nothing Then Err.Raise 5, , "所选区域没有可脱敏的数据，请重新选择。"
        DH_RequireSelection maskRange
        If maskRange.CountLarge > 100000 Then Err.Raise 5, , "脱敏数据区超过10万格，请分批处理。"
        Set sourceRange = maskRange
        maskRange.Select
        autoNote = "已按 " & maskRange.Address(False, False) & " 自动取数据区；表头和公式按原规则跳过"
    End If
    DH_RequireSelection sourceRange
    If sourceRange.Areas.Count <> 1 Then Err.Raise 5, , "One contiguous target required"
    If sourceRange.CountLarge > 100000 Then Err.Raise 5, , "选区太大：请只选这一列的数据区（例如 A2:A100）"

    BindAuxiliarySheet actions
    counter = counter + 1
    sessionId = format$(Now, "yyyymmddhhnnss") & "-" & CStr(counter)
    lockedActions = actions: confirmed = False: planReady = True
    DHPoc2_Prepare = Array("PREPARED_WRITE", sessionId, sourceBook.fullName, sourceSheet.CodeName, sourceSheet.name, sourceRange.address(False, False), lockedActions)
End Function

Private Sub RequireCurrent()
    If sourceGuard Is Nothing Or Len(sessionId) = 0 Then Err.Raise vbObjectError + 981, , "No scan session"
    If Len(invalidReason) > 0 Then Err.Raise vbObjectError + 981, , "Stale session: " & invalidReason
    If Now > expires Then Err.Raise vbObjectError + 981, , "Expired scan session"
    If Not sourceGuard.Matches(True) Then Err.Raise vbObjectError + 981, , "Workbook/sheet/snapshot changed"
    If Not auxGuard Is Nothing Then
        If Not auxGuard.Matches(False) Then Err.Raise vbObjectError + 981, , "Reference worksheet snapshot changed"
    End If
    If Not auxGuards Is Nothing Then
        Dim gItem As DHUISnapshot
        For Each gItem In auxGuards
            If Not gItem.Matches(False) Then Err.Raise vbObjectError + 981, , "Reference worksheet snapshot changed"
        Next gItem
    End If
    If Not TypeOf Application.Selection Is Range Then Err.Raise vbObjectError + 981, , "Range changed"
    If Not Application.Selection.parent Is sourceSheet Then Err.Raise vbObjectError + 981, , "Sheet changed"
    If Application.Selection.address <> sourceRange.address Then Err.Raise vbObjectError + 981, , "Range changed"
End Sub

' Called from the isolated native test driver, never inferred from HTML confirmed=true.
Public Function DHPoc2_ConfirmationTarget() As Variant
    RequireCurrent
    If Not planReady Then Err.Raise vbObjectError + 982, , "Prepare plan first"
    ' Item 5 (R15) is a readable plan for native dialogs; items 0-4 are unchanged.
    DHPoc2_ConfirmationTarget = Array(sessionId, sourceBook.fullName, sourceSheet.name, sourceRange.address(False, False), lockedActions, DescribePlan(lockedActions))
End Function

Private Function DescribePlan(ByVal actions As String) As String
    Dim parts As Variant, i As Long, action As String, text As String
    parts = Split(actions, "|")
    For i = 0 To UBound(parts)
        action = CStr(parts(i))
        If left$(action, 10) = "Transform:" Then action = Mid$(action, 11)
        text = text & CStr(i + 1) & ". " & DHUI_ActionName(action) & vbLf & "   " & DHUI_Consequence(action) & vbLf
    Next i
    DescribePlan = text
End Function

Public Function DHPoc2_Confirm(ByVal expectedSession As String) As String
    RequireCurrent
    If Not planReady Then Err.Raise vbObjectError + 982, , "Prepare plan first"
    If StrComp(expectedSession, sessionId, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 982, , "Confirmation session mismatch"
    confirmed = True
    DHPoc2_Confirm = "CONFIRMED_BOUND_WRITE_PLAN"
End Function

' Batch rule (R15): one confirmation may run several cleanup actions. Row
' deletion runs after cell-level cleanup so it sees the cleaned values; every
' business function re-plans from live data and deletes bottom-up itself.
' This is the only authoritative copy of the rule; ranks and "alone" come from
' DHUI_ActionInfo. The C# host checks shape only; the page only gives hints.
Private Sub ValidateActions(ByVal actions As String)
    Dim parts As Variant, item As Variant, seen As Object, key As String, i As Long, subParts As Variant
    Dim rank As Long, lastRank As Long, info As Variant
    parts = Split(actions, "|"): Set seen = CreateObject("Scripting.Dictionary")
    If UBound(parts) > 5 Or Len(actions) = 0 Then Err.Raise 5, , "1 to 6 actions required"
    lastRank = -1
    For i = 0 To UBound(parts)
        item = parts(i): key = CStr(item)
        ' Only exact registered spellings; parameterised actions keep a fixed field count.
        If key = "Unit:Force" Then
            ' One-run override of the mixed-unit guard; same slot as Unit (cannot run both).
            key = "Unit"
        ElseIf left$(key, 7) = "Lookup:" Then
            subParts = Split(key, ":")
            If UBound(subParts) <> 7 Then Err.Raise 5, , "Invalid Lookup action parameters"
            key = "Lookup"
        ElseIf left$(key, 7) = "Append:" Then
            subParts = Split(key, ":")
            If UBound(subParts) <> 6 Then Err.Raise 5, , "Invalid Append action parameters"
            key = "Append"
        ElseIf left$(key, 8) = "Unpivot:" Then
            subParts = Split(key, ":")
            If UBound(subParts) <> 7 Then Err.Raise 5, , "Invalid Unpivot action parameters"
            key = "Unpivot"
        ElseIf left$(key, 11) = "SplitSheet:" Then
            subParts = Split(key, ":")
            If UBound(subParts) <> 4 Then Err.Raise 5, , "Invalid SplitSheet action parameters"
            key = "SplitSheet"
        ElseIf left$(key, 10) = "Reconcile:" Then
            subParts = Split(key, ":")
            Dim checkedRange As Range
            Set checkedRange = DHUI_ReconcileRange(CStr(item), sourceBook)
            key = "Reconcile"
        ElseIf left$(key, 10) = "Transform:" Then
            If key <> "Transform:Phone" And key <> "Transform:ID" And key <> "Transform:Name" Then Err.Raise 5, , "Action disabled or unknown"
            key = "Transform"
        ElseIf key = "Phone" Or key = "ID" Or key = "Name" Or key = "Lookup" Or key = "Append" Or key = "Unpivot" Or key = "SplitSheet" Or key = "Reconcile" Then
            Err.Raise 5, , "Action disabled or unknown"
        End If
        info = DHUI_ActionInfo(CStr(item))
        If IsEmpty(info) Then Err.Raise 5, , "Action disabled or unknown"
        rank = CLng(info(2))
        If CBool(info(3)) And UBound(parts) <> 0 Then Err.Raise 5, , "Standalone action cannot be batched"
        If seen.Exists(key) Then Err.Raise 5, , "Duplicate action"
        seen.Add key, True
        If key = "Unmerge" And i <> 0 Then Err.Raise 5, , "Unmerge must be first"
        If rank < lastRank Then Err.Raise 5, , "Batch order must be Unmerge, cell cleanup, BlankRows, Dedup, Masking"
        lastRank = rank
    Next i
    If seen.Exists("Mark") And seen.Exists("Dedup") Then Err.Raise 5, , "Mark and Dedup cannot be combined"
    If seen.Exists("Transform") And (seen.Exists("BlankRows") Or seen.Exists("Dedup")) Then _
        Err.Raise 5, , "Masking cannot be combined with row deletion"
End Sub

Private Function RangeAlive(ByVal target As Range) As Boolean
    Dim probe As String
    If target Is Nothing Then Exit Function
    On Error Resume Next
    probe = target.address
    RangeAlive = (Err.number = 0)
    Err.Clear
End Function

Public Sub DHPoc2_RequireExecuting(ByVal target As Range)
    If Not executing Or Not planReady Then Err.Raise vbObjectError + 990, , "No executing bound plan"
    If Not target.parent Is sourceSheet Then Err.Raise 5, , "Execution sheet changed"
    If target.address <> sourceRange.address Then Err.Raise 5, , "Execution range changed"
    If Not sourceGuard.Matches(True) Then Err.Raise 5, , "Execution snapshot changed"
    If Not auxGuard Is Nothing Then
        If Not auxGuard.Matches(False) Then Err.Raise 5, , "Execution reference sheet snapshot changed"
    End If
    If Not auxGuards Is Nothing Then
        Dim gExec As DHUISnapshot
        For Each gExec In auxGuards
            If Not gExec.Matches(False) Then Err.Raise 5, , "Execution reference sheet snapshot changed"
        Next gExec
    End If
    DH_RequireSelection target
End Sub

Private Function RowsArray(ByVal rows As Collection) As Variant
    Dim result() As Variant, i As Long
    If rows.count = 0 Then RowsArray = Array(): Exit Function
    ReDim result(0 To rows.count - 1)
    For i = 1 To rows.count: result(i - 1) = rows(i): Next i
    RowsArray = result
End Function

Private Function AuditSkips(ByVal path As String) As Variant
    Dim stream As Object, lines As Variant, line As Variant, fields As Variant, reasons As New Collection, n As Long
    If Len(path) = 0 Then AuditSkips = Array(0, Array()): Exit Function
    Set stream = CreateObject("ADODB.Stream"): stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.LoadFromFile path: lines = Split(stream.ReadText, vbCrLf): stream.Close
    For Each line In lines
        fields = Split(CStr(line), vbTab)
        If UBound(fields) >= 2 Then
            If fields(1) = "保留" Or fields(1) = "跳过" Then
                n = n + 1
                If reasons.count < 200 Then reasons.Add Array(fields(0), fields(2))
            End If
        End If
    Next line
    AuditSkips = Array(n, RowsArray(reasons))
End Function

' 只取「真改了什么」：审计里除“保留/跳过”之外的行，给面板展示改动清单
Private Function AuditChanges(ByVal path As String) As Variant
    Dim stream As Object, lines As Variant, line As Variant, fields As Variant, items As New Collection, n As Long
    If Len(path) = 0 Then AuditChanges = Array(0, Array()): Exit Function
    Set stream = CreateObject("ADODB.Stream"): stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.LoadFromFile path: lines = Split(stream.ReadText, vbCrLf): stream.Close
    For Each line In lines
        fields = Split(CStr(line), vbTab)
        If UBound(fields) >= 2 Then
            If fields(1) <> "保留" And fields(1) <> "跳过" And fields(1) <> "类型" And fields(1) <> "类别" _
               And Len(CStr(fields(0))) > 0 And CStr(fields(0)) <> "位置" Then
                n = n + 1
                If items.count < 300 Then items.Add Array(FriendlyAddr(CStr(fields(0))), fields(1), CStr(fields(2)))
            End If
        End If
    Next line
    AuditChanges = Array(n, RowsArray(items))
End Function

Private Function ColLetter(ByVal idx As Long) As String
    Dim n As Long, s As String, r As Long
    n = idx
    Do While n > 0
        r = (n - 1) Mod 26
        s = Chr$(65 + r) & s
        n = (n - 1) \ 26
    Loop
    ColLetter = s
End Function

' 当前工作表被隐藏的列（字母逗号分隔），用来告诉用户“去哪儿找”
Public Function DHPoc2_HiddenColumns() As String
    Dim ws As Worksheet, i As Long, list As String
    On Error Resume Next
    Set ws = ActiveDataSheet()
    On Error GoTo 0
    If ws Is Nothing Then DHPoc2_HiddenColumns = "": Exit Function
    For i = 1 To ws.UsedRange.Columns.Count + ws.UsedRange.Column - 1
        If ws.Columns(i).Hidden Then
            If Len(list) > 0 Then list = list & ","
            list = list & ColLetter(i)
        End If
    Next i
    DHPoc2_HiddenColumns = list
End Function

Public Function DHPoc2_Execute() As Variant
    Dim parts As Variant, item As Variant, action As String, count As Long, completed As New Collection
    Dim backup As String, audit As String, phase As String, message As String, code As Long, fso As Object
    Dim licenseFinished As Boolean, licenseState As Variant
    Dim started As Boolean, skips As Variant, targetName As String, logError As String, targetAddress As String
    Dim skipped As New Collection, skipCode As Long, skipMsg As String
    RequireCurrent
    If Not planReady Or Not confirmed Or Len(lockedActions) = 0 Then Err.Raise vbObjectError + 991, , "Native confirmation required"
    ValidateActions lockedActions
    DH_RequireSelection sourceRange
    DHL_RequirePlan lockedActions
    ' Row deletion can remove every cell of the locked range (R15-RT-02); keep
    ' the planned address as text so logging never touches a dead Range.
    targetAddress = sourceRange.address(False, False)
    ' Consume authorisation before any backup/log attempt; failures cannot replay it.
    confirmed = False: executing = True
    On Error GoTo Failed
    phase = "intent-log"
    DH_LastLogError = ""
    DH_Log "TaskPane执行", sourceBook, sourceSheet.name, targetAddress, 0, "开始；会话=" & sessionId & "；动作=" & lockedActions
    phase = "backup": backup = DH_Backup(sourceBook)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(backup) Then Err.Raise 5, , "Baseline missing"
    DH_VerifyBaseline sourceBook, backup
    parts = Split(lockedActions, "|")
    For Each item In parts
        phase = "operation:" & CStr(item)
        If Not RangeAlive(sourceRange) Then Err.Raise vbObjectError + 992, , "Locked range was removed by a previous step"
        DHPoc2_RequireExecuting sourceRange
        action = CStr(item)
        Select Case action
            Case "Transform:Phone": action = "Phone"
            Case "Transform:ID": action = "ID"
            Case "Transform:Name": action = "Name"
        End Select
        DH_LastAuditPath = "": DH_LastSummary = "": started = True
        On Error Resume Next
        Err.Clear
        count = DHUI_Apply(action, sourceRange)
        If Err.Number <> 0 Then
            skipCode = Err.Number: skipMsg = Err.Description
            Err.Clear
            On Error GoTo Failed
            ' 这一步做不了就跳过，整批继续；结果里如实写原因
            skipped.Add Array(CStr(item), skipMsg)
            If Not RangeAlive(sourceRange) Then Err.Raise vbObjectError + 992, , "Locked range was removed by a previous step"
            started = False
            GoTo NextAction
        End If
        On Error GoTo Failed
        audit = DH_LastAuditPath
        phase = "audit-check:" & CStr(item)
        If count > 0 And Len(audit) = 0 Then Err.Raise 5, , "Operation returned changes without audit path"
        If Len(audit) > 0 Then
            If Not fso.FileExists(audit) Then Err.Raise 5, , "Audit file missing after operation"
        End If
        If Len(DH_LastLogError) > 0 Then Err.Raise 5, , DH_LastLogError
        skips = AuditSkips(audit)
        Dim changes As Variant
        changes = AuditChanges(audit)
        Dim sum2 As String
        sum2 = DH_LastSummary
        If Len(autoNote) > 0 Then sum2 = sum2 & "；" & autoNote
        completed.Add Array(CStr(item), count, sum2, audit, skips, DHUI_CountUnit(CStr(item)), changes)
        sourceGuard.Capture sourceSheet
        If Not auxGuard Is Nothing Then auxGuard.Capture auxSheet
        started = False
NextAction:
    Next item
    phase = "completion-log"
    DH_Log "TaskPane执行", sourceBook, sourceSheet.name, targetAddress, 0, "完成；会话=" & sessionId & "；各动作数量见明细", backup
    licenseFinished = True
    licenseState = DHL_AfterExecution(completed)
    Dim finishNote As String
    finishNote = "文件未自动保存；修改数单位见每项结果末尾。"
    If skipped.Count > 0 Then
        If completed.Count = 0 Then
            finishNote = "没有成功完成的步骤；失败步骤可能已有局部修改，请检查明细及基线。"
        Else
            finishNote = "部分步骤未完成；已完成项见明细。失败步骤不保证零修改，文件未自动保存。"
        End If
    End If
    DHPoc2_Execute = Array("COMPLETED", sessionId, backup, RowsArray(completed), DH_LogPath(), finishNote, "", "", RowsArray(skipped), licenseState)
    executing = False: DHPoc2_Invalidate "execution consumed"
    Exit Function
Failed:
    code = Err.number: message = Err.Description
    logError = DH_LastLogError
    ' Leave the active handler first: inside it, On Error Resume Next does not
    ' trap, and a second error escaped as a bare envelope (R15-RT-02).
    Resume LogFailure
LogFailure:
    On Error Resume Next
    DH_LogFailure "TaskPane执行", sourceBook, sourceSheet.name, targetAddress, "阶段=" & phase & ("；") & message, backup
    If Len(DH_LastLogError) > 0 Then logError = DH_LastLogError
    On Error GoTo 0
    executing = False: DHPoc2_Invalidate "execution failed and consumed"
    If Not licenseFinished Then
        licenseFinished = True
        licenseState = DHL_AfterExecution(completed)
    End If
    DHPoc2_Execute = Array("FAILED", phase, backup, RowsArray(completed), code, message, logError, "已完成动作不保证整体回滚；失败动作可能部分写入，请检查日志和基线", RowsArray(skipped), licenseState)
End Function

' 审计里可能是 R4C10 / 行4 列10，统一翻成用户看得懂的 A1 写法（J4）
Private Function FriendlyAddr(ByVal raw As String) As String
    Dim s As String, pC As Long, rr As Long, cc As Long, tmp As String
    s = Trim$(raw)
    If Len(s) = 0 Then FriendlyAddr = s: Exit Function
    If Left$(s, 1) = "R" Then
        pC = InStr(2, s, "C")
        If pC > 2 Then
            tmp = Mid$(s, 2, pC - 2)
            If IsNumeric(tmp) Then
                rr = CLng(tmp)
                tmp = Mid$(s, pC + 1)
                If IsNumeric(tmp) Then
                    cc = CLng(tmp)
                    On Error Resume Next
                    FriendlyAddr = Cells(rr, cc).Address(False, False)
                    On Error GoTo 0
                    If Len(FriendlyAddr) > 0 Then Exit Function
                End If
            End If
        End If
    End If
    If Left$(s, 1) = "行" Then
        pC = InStr(1, s, "列")
        If pC > 1 Then
            tmp = Mid$(s, 2, pC - 2)
            If IsNumeric(tmp) Then
                rr = CLng(tmp)
                tmp = Mid$(s, pC + 1)
                If IsNumeric(tmp) Then
                    cc = CLng(tmp)
                    On Error Resume Next
                    FriendlyAddr = Cells(rr, cc).Address(False, False)
                    On Error GoTo 0
                    If Len(FriendlyAddr) > 0 Then Exit Function
                End If
            End If
        End If
    End If
    FriendlyAddr = s
End Function

' 当前表内容区里被隐藏的列字母（逗号分隔）
Private Function HiddenColsOf(ByVal ws As Worksheet) As String
    Dim rng As Range, i As Long, s As String
    On Error Resume Next
    Set rng = DH_ContentRange(ws)
    On Error GoTo 0
    If rng Is Nothing Then HiddenColsOf = "": Exit Function
    For i = rng.Column To rng.Column + rng.Columns.Count - 1
        If ws.Columns(i).Hidden Then
            If Len(s) > 0 Then s = s & ","
            s = s & ColLetter(i)
        End If
    Next i
    HiddenColsOf = s
End Function

' 面板点“跳到行N”时用：只接受 A1 地址，选中并滚动到可见
Public Function DHPoc2_SelectRange(ByVal addr As String) As String
    Dim ws As Worksheet, target As Range
    Set ws = ActiveDataSheet()
    If ws Is Nothing Then Err.Raise 5, , "没有活动工作表"
    Set target = ws.Range(addr)
    ws.Activate
    target.Worksheet.Activate
    Application.Goto target, True
    DHPoc2_SelectRange = "SELECTED"
End Function
