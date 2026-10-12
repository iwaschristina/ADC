Attribute VB_Name = "x__UTIL_ADC_NO_EDIT_REFERENCE"
' ================================================================================================================================================================================
' ChatGPT Rewrite Tracking Header
' ================================================================================================================================================================================
' ChatGPT Previous Version ID: CGPT-NONE
' ChatGPT Baseline Version ID: CGPT-NONE
' ChatGPT Version ID: CGPT-7LGA-HT2J-PYQW-EG8S-Y9QT
' ChatGPT Version Timestamp: Monday 8/24/26 1:21 PM
' Purpose: Track ChatGPT rewrites and prevent stale-baseline edits, rollback, regressions, and no-op replacements.
' Tracking: Previous Version ID identifies the version preceding this rewrite; Baseline Version ID identifies the Version ID in the exact code actually edited. For a normal rewrite, both must equal the baseline Version ID. Version ID must be new and Timestamp must use America/New_York and dddd m/d/yy h:nn AM/PM. During an ordinary rewrite, update only the four tracking-value lines unless the user explicitly requests a header-schema revision.
' Version ID Generation: Use format CGPT-XXXX-XXXX-XXXX-XXXX-XXXX and alphabet 23456789ABCDEFGHJKLMNPQRSTUVWXYZ. Generate 20 random characters using one random byte per character; accept 0-247, reject 248-255, and use accepted byte Mod 31 as the alphabet index. Format as five groups of four and prefix with CGPT-.
' Baseline and Rewrite Rules:
'   1. Use the complete latest code provided by the user or, if none is newer, the latest complete version delivered by ChatGPT in this conversation. Treat user-provided code as authoritative, including manual changes made after its Version ID was assigned.
'   2. Do not replace, reconstruct, supplement, or compare the approved baseline with an earlier version having the same Version ID. Unless explicitly authorized, do not inspect, consult, compare, copy from, diff, or otherwise reference any older baseline for any purpose.
'   3. Evaluate fixes, speed improvements, behavior, design, and other preservation requirements only from the exact baseline and explicit user instructions. Report historical requirements that cannot be established from those sources.
'   4. Before editing, read and record the baseline Version ID. If it does not match the expected current version, stop and report a baseline mismatch.
'   5. Write the baseline Version ID into both Previous Version ID and Baseline Version ID, then generate a new Version ID and localized timestamp. If the baseline has no valid Version ID, use CGPT-NONE for both. Do not generate tracking values unless the delivered code contains an actual change.
'   6. For multi-module rewrites, establish each module's baseline and tracking values independently; never copy tracking values between modules.
'   7. If the requested change produces no difference, do not update tracking values or deliver a replacement; report that the baseline already satisfies the request.
' Rewrite Coding Rules:
'   1. Organize module-level constants, variables, object references, types, and related declarations into logical groups. Comment each declaration individually or provide an accurate group comment covering every declaration in the group.
'   2. Give every database object and table field created by the delivered code a detailed Description property covering purpose, meaning, ownership, expected contents, relevant behavior, and applicable constraints.
'   3. Keep comments with the code they describe and move them with moved, reorganized, or replaced code. Remove code or comments made dead, unreachable, obsolete, duplicate, unused, stale, misleading, inaccurate, or inapplicable by the current rewrite. Report pre-existing violations encountered during the work, but do not correct them unless authorized or required to complete the change safely.
'   4. Preserve all fixes and verified speed improvements represented by the exact baseline; do not regress performance, correctness, behavior, ownership, ordering, rollback, transaction, or other established invariants.
'   5. Do not add unnecessary abstraction, named aliases, forwarding procedures, compatibility layers, pass-through helpers, or wrappers. Retain or add them only when they serve a necessary purpose that cannot be achieved more directly without harming correctness or maintainability.
'   6. Report apparent violations encountered while reviewing or editing the applicable code. Do not perform an unrequested complete-module audit or silently correct unrelated violations unless authorized or required for a safe change.
' Preservation and Validation: Preserve required module-level and procedure-level Attribute lines unless the requested change explicitly requires otherwise. The visible header-first layout does not authorize deleting, relocating, or modifying exported VBA Attribute lines. Check the delivered code against the request, report completed static checks and comparisons, and clearly distinguish them from compilation, execution, runtime, performance, or regression testing not actually performed in an available environment.
' Required Output Check:
'   Verify that the exact approved baseline, including user changes, was used and no unauthorized older baseline was referenced.
'   Verify Previous and Baseline Version IDs identify the exact baseline, are equal for a normal rewrite, and use CGPT-NONE only for an unversioned baseline.
'   Verify the new Version ID is unique, the timestamp is current, and each rewritten module has independently derived tracking values.
'   Verify requested changes are present; removed features were not reintroduced; prior required behavior, fixes, and speed improvements were not regressed.
'   Verify affected declarations are organized and commented; created database objects and fields have detailed descriptions; applicable comments remain with their code; code and comments made obsolete by the rewrite were removed; and encountered pre-existing violations were reported rather than silently corrected.
'   Verify no unnecessary abstraction, aliases, forwarding procedures, compatibility layers, pass-through helpers, or wrappers were introduced.
'   Verify required exported Attribute lines were preserved and validation claims accurately distinguish completed checks from checks not performed.
'   Verify the output is not an accidental duplicate of the baseline or another prior version.
' Delivery: Save each rewritten .bas or .cls file under a new filename containing its new Version ID. Reopen and read that exact file and verify its header contains the Version ID and timestamp reported to the user. Do not reuse an old download filename or provide a download link to an unverified file.
' Formatting:
'   Do not add, rename, remove, reorder, or improvise managed-header categories without an explicit schema revision.
'   Keep ordinary comments with their code or in an approved editable managed-header field; do not convert them into new header categories.
'   Visible layout: ChatGPT header, one blank line, managed module header, one blank line, Option statements, one blank line, declarations, one blank line, procedures.
'   Exported Attribute lines remain in required import positions. Use one blank line between procedures, none inside procedures, and no trailing blank lines. Keep this header above the managed module header. Treat generated managed-header metadata as utility-owned unless explicit correction is requested.
'   Preserve editable managed-header documentation and merge required updates instead of deleting existing content.
'   Do not claim generated metadata is current unless regenerated, verified, or unaffected by the code change.
'   Do not modify these instructions, rules, checks, or formatting requirements without an explicit rewrite-tracking header schema revision.
' ================================================================================================================================================================================

' ================================================================================================================================================================================
' MODULE HEADER
' ================================================================================================================================================================================
' STANDARD MODULE: NO EDIT Reference Only Module For Non-ADC Namespaced Modules Referenced by ADC Namespaced Modules
' VB_Name: x__UTIL_ADC_NO_EDIT_REFERENCE
' Export Group: ADC baseline reference/export support
' Entry Points: None
' SHARED CODE: None
' CLASSES USED: None
' DYNAMIC STRING REFERENCES: None
' LastScanTimestamp: Monday 8/24/26 1:19 PM
' Purpose: Preserve commented, non-executable reference copies of separately existing modules whose actual module names do not share the ADC export prefix, so the prefix-based ADC export can carry this reference material with the ADC baseline.
' Author:
' Support:
' Code Category: Reference-only ADC baseline export support
' Requirements: Keep all embedded reference code commented. The actual modules must continue to exist separately in the database under the exact names shown in their embedded module headers; this module must not replace, rename, or become the runtime copy of either module.
' Reference Dependencies: Form_ADC Config - Both; x__SHARED_CODE. The actual modules exist separately under these names; the copies below are reference snapshots only.
' DB Object Dependencies: None for this reference-only wrapper; dependencies shown inside the commented reference code belong to the actual modules represented there.
' Usage: Include x__UTIL_ADC_NO_EDIT_REFERENCE in ADC baseline exports and use the commented contents only when reviewing, auditing, or tracing ADC behavior that interacts with the separately named actual modules.
' Version:
' Notes: This wrapper intentionally uses an ADC-exportable module name because the export utility selects modules by module-name prefix. The referenced modules keep their actual database names, so their code is duplicated here only as commented reference material for baseline completeness.
' Core Rules Enforced Here: DO NOT EDIT. Reference only; preserve the actual module identities shown below; keep all embedded module code commented and non-executable.
' Design Map: Commented reference section for Form_ADC Config - Both, followed by a commented reference section for x__SHARED_CODE.
' Configuration: None
' Behavior: No executable behavior. This module contains comments only.
' Code Status: Reference-only export-support module; not an executable substitute for the separately existing modules represented below.
' Todo: None
' Reporting Bugs:
' Module Level Bugs:
' Public Subs: None
' Public Functions: None
' Public Properties: None
' ================================================================================================================================================================================

' ================================================================================================================================================================================
' FORM MODULE: Parent Form (ADC Config - Both), Form module for ADC Config - Both
' VB_Name: Form_ADC Macro Catalog
'
'Private Const REVIEW_STATUS_NEEDS_REVIEW As String = "NeedsReview"
'Private Const REVIEW_STATUS_REVIEWED As String = "Reviewed"
'Private Const REVIEW_STATUS_DEFERRED As String = "Deferred"
'Private Const REVIEW_STATUS_IGNORE As String = "Ignore"
'Private Const REVIEW_STATUS_BLOCKED As String = "Blocked"
'
'Private mSync As x__FORM_SUBFORM_SYNC
'Private mSyncReady As Boolean
'Private mSuppressGeneratedChildCurrentSync As Boolean
'Private mSuppressGeneratedRuntimePopups As Boolean
'Private mGeneratedRuntimeErrorCount As Long
'Private mLastGeneratedRuntimeProcedure As String
'Private mLastGeneratedRuntimeErrorNumber As Long
'Private mLastGeneratedRuntimeErrorSource As String
'Private mLastGeneratedRuntimeErrorDescription As String
'Private mLastGeneratedRuntimeErrorDetail As String
'Private Const DS_SUBFORM_CONTROL_NAME As String = "DSForm"
'Private Const SINGLE_SUBFORM_CONTROL_NAME As String = "SingleForm"
'Private Const KEY_FIELD_NAME As String = "ID"
'Private Const ROLE_DS As String = "DS"
'Private Const ROLE_SINGLE As String = "Single"
'Private Const HEALTH_CHECK_TEMPVAR_NAME As String = "GeneratedFormSetBuilderHealthCheck"
'
'Private Sub Form_Load()
'    On Error GoTo ErrHandler
'    FitGeneratedPopupToSavedDesign
'    InitSubformSync
'    RefreshDatasheetForCurrentTab
'    Exit Sub
'ErrHandler:
'    ShowGeneratedRuntimeError "Form_Load", Err.Number, Err.Source, Err.Description
'End Sub
'
'Private Sub FitGeneratedPopupToSavedDesign()
'    On Error GoTo ErrHandler
'    Dim requiredInsideHeight As Long
'    requiredInsideHeight = VisibleGeneratedSectionHeight(acHeader) + VisibleGeneratedSectionHeight(acDetail) + VisibleGeneratedSectionHeight(acFooter)
'    If requiredInsideHeight <= 0 Then Err.Raise vbObjectError + 20, Me.Name & ".FitGeneratedPopupToSavedDesign", "The generated popup has no positive visible section height."
'    Me.InsideHeight = requiredInsideHeight
'    Exit Sub
'ErrHandler:
'    Err.Raise Err.Number, Me.Name & ".FitGeneratedPopupToSavedDesign", Err.Description
'End Sub
'
'Private Function VisibleGeneratedSectionHeight(ByVal sectionIndex As Long) As Long
'    On Error GoTo SectionUnavailable
'    If Me.Section(sectionIndex).Visible Then VisibleGeneratedSectionHeight = CLng(Me.Section(sectionIndex).Height)
'    Exit Function
'SectionUnavailable:
'    Err.Clear
'End Function
'
'Private Sub Form_Unload(Cancel As Integer)
'    On Error Resume Next
'    mSyncReady = False
'    If Not mSync Is Nothing Then mSync.Clear
'    Set mSync = Nothing
'End Sub
'
'Private Sub InitSubformSync()
'    On Error GoTo ErrHandler
'    If mSyncReady Then Exit Sub
'    mSyncReady = False
'    If Not mSync Is Nothing Then mSync.Clear
'    Set mSync = New x__FORM_SUBFORM_SYNC
'    mSync.Init Me, DS_SUBFORM_CONTROL_NAME, SINGLE_SUBFORM_CONTROL_NAME, KEY_FIELD_NAME
'    mSyncReady = True
'    Exit Sub
'ErrHandler:
'    mSyncReady = False
'    Set mSync = Nothing
'    Err.Raise Err.Number, Me.Name & ".InitSubformSync", Err.Description
'End Sub
'
'' Intended use: Shows only datasheet fields whose bound SingleForm controls are on the currently selected tab page.
'' Maintenance: No field lists or tab captions are stored here. Move a bound SingleForm control to another tab page and the datasheet grouping follows automatically.
'Public Sub RefreshDatasheetForCurrentTab()
'    On Error GoTo ErrHandler
'    Dim SingleForm As Access.Form
'    Dim DSForm As Access.Form
'    Dim ctl As Access.control
'    Dim dsCtl As Access.control
'    Dim tabCtl As Access.TabControl
'    Dim activePage As Access.Page
'    Dim visibleFields As Object
'    Dim fieldName As String
'    Dim tabControlCount As Long
'    Set SingleForm = Me.Controls(SINGLE_SUBFORM_CONTROL_NAME).Form
'    Set DSForm = Me.Controls(DS_SUBFORM_CONTROL_NAME).Form
'    If SingleForm Is Nothing Then Err.Raise vbObjectError + 41, Me.Name & ".RefreshDatasheetForCurrentTab", "The SingleForm subform is unavailable."
'    If DSForm Is Nothing Then Err.Raise vbObjectError + 42, Me.Name & ".RefreshDatasheetForCurrentTab", "The DSForm subform is unavailable."
'    For Each ctl In SingleForm.Controls
'        If ctl.controlType = acTabCtl Then
'            tabControlCount = tabControlCount + 1
'            If tabControlCount = 1 Then Set tabCtl = ctl
'        End If
'    Next ctl
'    If tabControlCount <> 1 Then Err.Raise vbObjectError + 43, Me.Name & ".RefreshDatasheetForCurrentTab", "Expected exactly one tab control on SingleForm but found " & CStr(tabControlCount) & "."
'    Set activePage = tabCtl.pages(tabCtl.value)
'    Set visibleFields = CreateObject("Scripting.Dictionary")
'    visibleFields.CompareMode = vbTextCompare
'    For Each ctl In SingleForm.Controls
'        Select Case ctl.controlType
'            Case acTextBox, acComboBox, acCheckBox, acListBox, acOptionGroup, acOptionButton, acToggleButton
'                If StrComp(ctl.Parent.Name, activePage.Name, vbBinaryCompare) = 0 Then
'                    fieldName = GeneratedBoundFieldName(ctl)
'                    If Len(fieldName) > 0 Then visibleFields(fieldName) = True
'                End If
'        End Select
'    Next ctl
'    If visibleFields.count = 0 Then Err.Raise vbObjectError + 44, Me.Name & ".RefreshDatasheetForCurrentTab", "The active SingleForm tab page '" & activePage.Caption & "' has no bound field controls."
'    For Each dsCtl In DSForm.Controls
'        Select Case dsCtl.controlType
'            Case acTextBox, acComboBox, acCheckBox, acListBox, acOptionGroup, acOptionButton, acToggleButton
'                fieldName = GeneratedBoundFieldName(dsCtl)
'                If Len(fieldName) > 0 Then CallByName dsCtl, "ColumnHidden", VbLet, Not visibleFields.Exists(fieldName)
'        End Select
'    Next dsCtl
'    Exit Sub
'ErrHandler:
'    ShowGeneratedRuntimeError "RefreshDatasheetForCurrentTab", Err.Number, Err.Source, Err.Description
'End Sub
'
'Private Function GeneratedBoundFieldName(ByVal ctl As Access.control) As String
'    On Error GoTo ExitHere
'    Dim sourceText As String
'    sourceText = Trim$(CStr(CallByName(ctl, "ControlSource", VbGet)))
'    If Len(sourceText) = 0 Then Exit Function
'    If Left$(sourceText, 1) = "=" Then Exit Function
'    If Left$(sourceText, 1) = "[" And Right$(sourceText, 1) = "]" And Len(sourceText) > 2 Then sourceText = Mid$(sourceText, 2, Len(sourceText) - 2)
'    GeneratedBoundFieldName = sourceText
'ExitHere:
'End Function
'
'' Intended use: Returns True after parent/subform synchronization has initialized successfully.
'Public Function GeneratedSyncReady() As Boolean
'    GeneratedSyncReady = (mSyncReady And Not mSync Is Nothing)
'End Function
'
'' Intended use: Tells generated child Current events to remain idle only while the parent health check positions records.
'Public Function GeneratedChildCurrentSyncSuppressed() As Boolean
'    GeneratedChildCurrentSyncSuppressed = mSuppressGeneratedChildCurrentSync
'End Function
'
'' Intended use: Runs deterministic outcome-based parent, child, bidirectional-sync, and requery health checks and returns blank on success.
'Public Function GeneratedRuntimeHealthCheck() As String
'    On Error GoTo ErrHandler
'    Dim childStatus As String
'    Dim DSForm As Form
'    Dim SingleForm As Form
'    Dim dsRecordset As Object
'    Dim firstKey As Variant
'    Dim secondKey As Variant
'    Dim originalDSKey As Variant
'    Dim hasSecondRecord As Boolean
'    Dim hasOriginalDSKey As Boolean
'    If Len(GeneratedRuntimeHealthStatus()) > 0 Then
'        GeneratedRuntimeHealthCheck = GeneratedRuntimeHealthStatus()
'        Exit Function
'    End If
'    If Not GeneratedSyncReady() Then InitSubformSync
'    If Not GeneratedSyncReady() Then Err.Raise vbObjectError + 1, Me.Name & ".GeneratedRuntimeHealthCheck", "Synchronization did not become ready."
'    Set SingleForm = Me.Controls(SINGLE_SUBFORM_CONTROL_NAME).Form
'    Set DSForm = Me.Controls(DS_SUBFORM_CONTROL_NAME).Form
'    childStatus = CStr(CallByName(SingleForm, "GeneratedRuntimeHealthStatus", VbMethod))
'    If Len(childStatus) > 0 Then Err.Raise vbObjectError + 2, Me.Name & ".GeneratedRuntimeHealthCheck", "SingleForm reported a runtime error: " & childStatus
'    childStatus = CStr(CallByName(DSForm, "GeneratedRuntimeHealthStatus", VbMethod))
'    If Len(childStatus) > 0 Then Err.Raise vbObjectError + 3, Me.Name & ".GeneratedRuntimeHealthCheck", "DSForm reported a runtime error: " & childStatus
'    hasOriginalDSKey = TryGeneratedFormCurrentKey(DSForm, originalDSKey)
'    Set dsRecordset = DSForm.RecordsetClone
'    If dsRecordset Is Nothing Then Err.Raise vbObjectError + 4, Me.Name & ".GeneratedRuntimeHealthCheck", "DSForm did not expose a RecordsetClone."
'    If dsRecordset.BOF And dsRecordset.EOF Then
'        dsRecordset.Close
'        Set dsRecordset = Nothing
'        RequerySyncedSubforms False, ROLE_DS
'    Else
'        dsRecordset.MoveFirst
'        firstKey = dsRecordset.fields(KEY_FIELD_NAME).value
'        dsRecordset.MoveNext
'        If Not dsRecordset.EOF Then
'            secondKey = dsRecordset.fields(KEY_FIELD_NAME).value
'            hasSecondRecord = True
'        End If
'        dsRecordset.Close
'        Set dsRecordset = Nothing
'        PositionGeneratedFormByKeyForHealth DSForm, firstKey
'        SyncFromDS
'        AssertGeneratedFormCurrentKey SingleForm, firstKey, "DS-to-Single synchronization"
'        RequerySyncedSubforms True, ROLE_DS
'        AssertGeneratedFormCurrentKey DSForm, firstKey, "DS-source requery preservation on DSForm"
'        AssertGeneratedFormCurrentKey SingleForm, firstKey, "DS-source requery preservation on SingleForm"
'        If hasSecondRecord Then
'            PositionGeneratedFormByKeyForHealth SingleForm, secondKey
'            SyncFromSingle
'            AssertGeneratedFormCurrentKey DSForm, secondKey, "Single-to-DS synchronization"
'            RequerySyncedSubforms True, ROLE_SINGLE
'            AssertGeneratedFormCurrentKey SingleForm, secondKey, "Single-source requery preservation on SingleForm"
'            AssertGeneratedFormCurrentKey DSForm, secondKey, "Single-source requery preservation on DSForm"
'        Else
'            PositionGeneratedFormByKeyForHealth SingleForm, firstKey
'            SyncFromSingle
'            AssertGeneratedFormCurrentKey DSForm, firstKey, "Single-to-DS synchronization for one-record source"
'        End If
'        If hasOriginalDSKey Then
'            PositionGeneratedFormByKeyForHealth DSForm, originalDSKey
'            SyncFromDS
'            AssertGeneratedFormCurrentKey SingleForm, originalDSKey, "Health-check original-record restoration"
'        End If
'    End If
'    If Len(GeneratedRuntimeHealthStatus()) > 0 Then Err.Raise vbObjectError + 5, Me.Name & ".GeneratedRuntimeHealthCheck", GeneratedRuntimeHealthStatus()
'    childStatus = CStr(CallByName(SingleForm, "GeneratedRuntimeHealthStatus", VbMethod))
'    If Len(childStatus) > 0 Then Err.Raise vbObjectError + 6, Me.Name & ".GeneratedRuntimeHealthCheck", "SingleForm reported a runtime error after synchronization: " & childStatus
'    childStatus = CStr(CallByName(DSForm, "GeneratedRuntimeHealthStatus", VbMethod))
'    If Len(childStatus) > 0 Then Err.Raise vbObjectError + 7, Me.Name & ".GeneratedRuntimeHealthCheck", "DSForm reported a runtime error after synchronization: " & childStatus
'    GeneratedRuntimeHealthCheck = vbNullString
'    Exit Function
'ErrHandler:
'    Dim errNumber As Long
'    Dim errSource As String
'    Dim errDescription As String
'    errNumber = Err.Number
'    errSource = Err.Source
'    errDescription = Err.Description
'    On Error Resume Next
'    mSuppressGeneratedChildCurrentSync = False
'    If Not dsRecordset Is Nothing Then dsRecordset.Close
'    Set dsRecordset = Nothing
'    GeneratedRuntimeHealthCheck = "Source: " & errSource & vbCrLf & "Error " & CStr(errNumber) & ": " & errDescription
'End Function
'
'Private Sub PositionGeneratedFormByKeyForHealth(ByVal targetForm As Form, ByVal keyValue As Variant)
'    On Error GoTo ErrHandler
'    mSuppressGeneratedChildCurrentSync = True
'    PositionGeneratedFormByKey targetForm, keyValue
'    mSuppressGeneratedChildCurrentSync = False
'    Exit Sub
'ErrHandler:
'    Dim errNumber As Long
'    Dim errDescription As String
'    errNumber = Err.Number
'    errDescription = Err.Description
'    mSuppressGeneratedChildCurrentSync = False
'    Err.Raise errNumber, Me.Name & ".PositionGeneratedFormByKeyForHealth", errDescription
'End Sub
'
'Private Function TryGeneratedFormCurrentKey(ByVal targetForm As Form, ByRef keyValue As Variant) As Boolean
'    On Error GoTo ExitHere
'    If targetForm Is Nothing Then Exit Function
'    If targetForm.newRecord Then Exit Function
'    If targetForm.Recordset.BOF And targetForm.Recordset.EOF Then Exit Function
'    keyValue = targetForm.Recordset.fields(KEY_FIELD_NAME).value
'    If IsNull(keyValue) Then Exit Function
'    TryGeneratedFormCurrentKey = True
'ExitHere:
'End Function
'
'Private Sub PositionGeneratedFormByKey(ByVal targetForm As Form, ByVal keyValue As Variant)
'    On Error GoTo ErrHandler
'    Dim rs As Object
'    Dim found As Boolean
'    Set rs = targetForm.RecordsetClone
'    If rs Is Nothing Then Err.Raise vbObjectError + 8, Me.Name & ".PositionGeneratedFormByKey", "Target form did not expose a RecordsetClone."
'    If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
'    Do While Not rs.EOF
'        If GeneratedKeysEqual(rs.fields(KEY_FIELD_NAME).value, keyValue) Then
'            targetForm.Bookmark = rs.Bookmark
'            found = True
'            Exit Do
'        End If
'        rs.MoveNext
'    Loop
'    rs.Close
'    Set rs = Nothing
'    If Not found Then Err.Raise vbObjectError + 9, Me.Name & ".PositionGeneratedFormByKey", "Could not locate key value '" & SafeGeneratedKeyText(keyValue) & "' in form '" & targetForm.Name & "'."
'    Exit Sub
'ErrHandler:
'    Dim errNumber As Long
'    Dim errDescription As String
'    errNumber = Err.Number
'    errDescription = Err.Description
'    On Error Resume Next
'    If Not rs Is Nothing Then rs.Close
'    Set rs = Nothing
'    On Error GoTo 0
'    Err.Raise errNumber, Me.Name & ".PositionGeneratedFormByKey", errDescription
'End Sub
'
'Private Sub AssertGeneratedFormCurrentKey(ByVal targetForm As Form, ByVal expectedKey As Variant, ByVal checkName As String)
'    Dim actualKey As Variant
'    If Not TryGeneratedFormCurrentKey(targetForm, actualKey) Then Err.Raise vbObjectError + 10, Me.Name & ".AssertGeneratedFormCurrentKey", checkName & " failed because form '" & targetForm.Name & "' has no current key."
'    If Not GeneratedKeysEqual(actualKey, expectedKey) Then Err.Raise vbObjectError + 11, Me.Name & ".AssertGeneratedFormCurrentKey", checkName & " failed for form '" & targetForm.Name & "'. Expected key '" & SafeGeneratedKeyText(expectedKey) & "'; actual key '" & SafeGeneratedKeyText(actualKey) & "'."
'End Sub
'
'Private Function GeneratedKeysEqual(ByVal leftValue As Variant, ByVal rightValue As Variant) As Boolean
'    On Error GoTo CompareText
'    If IsNull(leftValue) Or IsNull(rightValue) Then
'        GeneratedKeysEqual = (IsNull(leftValue) And IsNull(rightValue))
'    Else
'        GeneratedKeysEqual = (leftValue = rightValue)
'    End If
'    Exit Function
'CompareText:
'    Err.Clear
'    GeneratedKeysEqual = (StrComp(CStr(leftValue), CStr(rightValue), vbBinaryCompare) = 0)
'End Function
'
'Private Function SafeGeneratedKeyText(ByVal keyValue As Variant) As String
'    On Error GoTo ExitHere
'    If IsNull(keyValue) Then
'        SafeGeneratedKeyText = "Null"
'    ElseIf IsEmpty(keyValue) Then
'        SafeGeneratedKeyText = "Empty"
'    Else
'        SafeGeneratedKeyText = CStr(keyValue)
'    End If
'ExitHere:
'End Function
'
'' Intended use: Synchronizes the SingleForm subform to the current DSForm record. Control highlighting is handled by ControlEntered.
'Public Sub SyncFromDS(Optional ByVal controlName As String = vbNullString)
'    On Error GoTo ErrHandler
'    If mSync Is Nothing Then InitSubformSync
'    If mSync Is Nothing Then Exit Sub
'    mSync.SyncFromDS
'    Exit Sub
'ErrHandler:
'    ShowGeneratedRuntimeError "SyncFromDS", Err.Number, Err.Source, Err.Description, "ControlName: " & controlName
'End Sub
'
'' Intended use: Synchronizes the DSForm subform to the current SingleForm record. Control highlighting is handled by ControlEntered.
'Public Sub SyncFromSingle(Optional ByVal controlName As String = vbNullString)
'    On Error GoTo ErrHandler
'    If mSync Is Nothing Then InitSubformSync
'    If mSync Is Nothing Then Exit Sub
'    mSync.SyncFromSingle
'    Exit Sub
'ErrHandler:
'    ShowGeneratedRuntimeError "SyncFromSingle", Err.Number, Err.Source, Err.Description, "ControlName: " & controlName
'End Sub
'
'' Intended use: Requeries both synchronized subforms while optionally preserving the source current record.
'Public Sub RequerySyncedSubforms(Optional ByVal KeepCurrentRecord As Boolean = True, Optional ByVal SourceRole As String = ROLE_DS)
'    On Error GoTo ErrHandler
'    If mSync Is Nothing Then InitSubformSync
'    If mSync Is Nothing Then Exit Sub
'    mSync.RequeryBoth KeepCurrentRecord, SourceRole
'    Exit Sub
'ErrHandler:
'    ShowGeneratedRuntimeError "RequerySyncedSubforms", Err.Number, Err.Source, Err.Description, "KeepCurrentRecord: " & CStr(KeepCurrentRecord) & "; SourceRole: " & SourceRole
'End Sub
'
'' Intended use: Returns retained generated runtime error details for automated or interactive diagnostics; blank means no error has been recorded.
'Public Function GeneratedRuntimeHealthStatus() As String
'    If mGeneratedRuntimeErrorCount = 0 Then Exit Function
'    GeneratedRuntimeHealthStatus = "Runtime error count: " & CStr(mGeneratedRuntimeErrorCount) & vbCrLf & "Procedure: " & mLastGeneratedRuntimeProcedure & vbCrLf & "Source: " & mLastGeneratedRuntimeErrorSource & vbCrLf & "Error " & CStr(mLastGeneratedRuntimeErrorNumber) & ": " & mLastGeneratedRuntimeErrorDescription
'    If Len(mLastGeneratedRuntimeErrorDetail) > 0 Then GeneratedRuntimeHealthStatus = GeneratedRuntimeHealthStatus & vbCrLf & mLastGeneratedRuntimeErrorDetail
'End Function
'
'Private Function GeneratedHealthCheckModeActive() As Boolean
'    On Error GoTo ExitHere
'    GeneratedHealthCheckModeActive = CBool(TempVars(HEALTH_CHECK_TEMPVAR_NAME).value)
'ExitHere:
'    Err.Clear
'End Function
'
'Private Sub ShowGeneratedRuntimeError(ByVal ProcedureName As String, ByVal errNumber As Long, ByVal errSource As String, ByVal errDescription As String, Optional ByVal detailText As String = vbNullString)
'    On Error Resume Next
'    Dim messageText As String
'    Dim response As VbMsgBoxResult
'    mGeneratedRuntimeErrorCount = mGeneratedRuntimeErrorCount + 1
'    mLastGeneratedRuntimeProcedure = ProcedureName
'    mLastGeneratedRuntimeErrorNumber = errNumber
'    mLastGeneratedRuntimeErrorSource = errSource
'    mLastGeneratedRuntimeErrorDescription = errDescription
'    mLastGeneratedRuntimeErrorDetail = detailText
'    messageText = "Generated form runtime error." & vbCrLf & vbCrLf & "Form: " & Me.Name & vbCrLf & "Procedure: " & ProcedureName & vbCrLf & "Error " & CStr(errNumber) & ": " & errDescription
'    If Len(errSource) > 0 Then messageText = messageText & vbCrLf & "Source: " & errSource
'    If Len(detailText) > 0 Then messageText = messageText & vbCrLf & vbCrLf & detailText
'    Debug.Print Me.Name & ": " & ProcedureName & " failed. Source='" & errSource & "'; Error " & CStr(errNumber) & ": " & errDescription & IIf(Len(detailText) > 0, "; " & detailText, vbNullString)
'    SysCmd acSysCmdSetStatus, Me.Name & " failed in " & ProcedureName & ". Error " & CStr(errNumber) & ": " & errDescription
'    If GeneratedHealthCheckModeActive() Then Exit Sub
'    If mSuppressGeneratedRuntimePopups Then Exit Sub
'    messageText = messageText & vbCrLf & vbCrLf & "Choose Yes to suppress further generated-form runtime popups for this open form. Diagnostics will continue in the Immediate Window and status bar."
'    response = MsgBox(messageText, vbCritical + vbYesNo + vbDefaultButton2, "Generated Form Runtime Error")
'    If response = vbYes Then
'        mSuppressGeneratedRuntimePopups = True
'        Debug.Print Me.Name & ": Further generated runtime error popups suppressed for this open form. Diagnostics continue in Immediate Window and status bar."
'        SysCmd acSysCmdSetStatus, Me.Name & " generated runtime popups suppressed for this open form."
'    End If
'End Sub
'
'Private Sub cmdRunCurrentConfig_Click()
'    ' Runs the configured target represented by the current AOC_Config record.
'    On Error GoTo ErrHandler
'    Dim frm As Access.Form
'    Dim ConfigID As Long
'    ' Save and validate the synchronized child forms, then return the bound SingleForm.
'    Set frm = PrepareCurrentConfigAction(ConfigID)
'    If frm Is Nothing Then Exit Sub
'    ' Run the current configuration without forcing its confirmation workflow.
'    RunConfiguredMacroTarget CStr(frm!MacroKey)
'    ' Refresh both child forms and return them to the same configuration record when it remains visible.
'    RefreshDSFormToConfigID ConfigID
'    Exit Sub
'ErrHandler:
'    ShowGeneratedRuntimeError "cmdRunCurrentConfig_Click", Err.Number, Err.Source, Err.Description, "ConfigID: " & CStr(ConfigID)
'End Sub
'
'Private Sub cmdRunCurrentConfigConfirmed_Click()
'    ' Runs the configured target represented by the current record through its confirmation workflow.
'    On Error GoTo ErrHandler
'    Dim frm As Access.Form
'    Dim ConfigID As Long
'    ' Save and validate the synchronized child forms, then return the bound SingleForm.
'    Set frm = PrepareCurrentConfigAction(ConfigID)
'    If frm Is Nothing Then Exit Sub
'    ' Run the current configuration through the confirmed-action procedure.
'    RunConfirmedConfiguredMacroTarget CStr(frm!MacroKey)
'    ' Refresh both child forms and return them to the same configuration record when it remains visible.
'    RefreshDSFormToConfigID ConfigID
'    Exit Sub
'ErrHandler:
'    ShowGeneratedRuntimeError "cmdRunCurrentConfigConfirmed_Click", Err.Number, Err.Source, Err.Description, "ConfigID: " & CStr(ConfigID)
'End Sub
'
'Private Sub cmdGenerateCurrentConfig_Click()
'    ' Generates or regenerates the Access macro for the current configuration record.
'    On Error GoTo ErrHandler
'    Dim frm As Access.Form
'    Dim ConfigID As Long
'    ' Save and validate the synchronized child forms, then return the bound SingleForm.
'    Set frm = PrepareCurrentConfigAction(ConfigID)
'    If frm Is Nothing Then Exit Sub
'    ' Generate the macro represented by the current configuration record.
'    GenerateConfiguredMacro ConfigID, True
'    ' Refresh both child forms and return them to the same configuration record when it remains visible.
'    RefreshDSFormToConfigID ConfigID
'    Exit Sub
'ErrHandler:
'    ShowGeneratedRuntimeError "cmdGenerateCurrentConfig_Click", Err.Number, Err.Source, Err.Description, "ConfigID: " & CStr(ConfigID)
'End Sub
'
'Private Sub cmdValidateCurrentConfig_Click()
'    ' Validates the procedure call and related metadata for the current configuration record.
'    On Error GoTo ErrHandler
'    Dim frm As Access.Form
'    Dim ConfigID As Long
'    ' Save and validate the synchronized child forms, then return the bound SingleForm.
'    Set frm = PrepareCurrentConfigAction(ConfigID)
'    If frm Is Nothing Then Exit Sub
'    ' Resolve and validate the current configuration through the current single-record target authority.
'    ResolveConfiguredTarget ConfigID, True
'    ' Refresh both child forms and return them to the same configuration record when it remains visible.
'    RefreshDSFormToConfigID ConfigID
'    Exit Sub
'ErrHandler:
'    ShowGeneratedRuntimeError "cmdValidateCurrentConfig_Click", Err.Number, Err.Source, Err.Description, "ConfigID: " & CStr(ConfigID)
'End Sub
'
'Private Sub cmdResolveCurrentConfig_Click()
'    ' Resolves and stores the database, project, module, procedure, and related target metadata.
'    On Error GoTo ErrHandler
'    Dim frm As Access.Form
'    Dim ConfigID As Long
'    ' Save and validate the synchronized child forms, then return the bound SingleForm.
'    Set frm = PrepareCurrentConfigAction(ConfigID)
'    If frm Is Nothing Then Exit Sub
'    ' Resolve and store the target represented by the current configuration record.
'    ResolveConfiguredTarget ConfigID, True
'    ' Refresh both child forms and return them to the same configuration record when it remains visible.
'    RefreshDSFormToConfigID ConfigID
'    Exit Sub
'ErrHandler:
'    ShowGeneratedRuntimeError "cmdResolveCurrentConfig_Click", Err.Number, Err.Source, Err.Description, "ConfigID: " & CStr(ConfigID)
'End Sub
'
'Private Sub cmdToggleCurrentActive_Click()
'    ' Reverses the Active value for the current configuration record.
'    On Error GoTo ErrHandler
'    Dim frm As Access.Form
'    Dim ConfigID As Long
'    ' Save and validate the synchronized child forms, then return the bound SingleForm.
'    Set frm = PrepareCurrentConfigAction(ConfigID)
'    If frm Is Nothing Then Exit Sub
'    ' Change Active from True to False or from False to True.
'    frm!Active = Not CBool(Nz(frm!Active, False))
'    If frm.Dirty Then frm.Dirty = False
'    ' Refresh both child forms and return them to the same configuration record when it remains visible.
'    RefreshDSFormToConfigID ConfigID
'    Exit Sub
'ErrHandler:
'    ShowGeneratedRuntimeError "cmdToggleCurrentActive_Click", Err.Number, Err.Source, Err.Description, "ConfigID: " & CStr(ConfigID)
'End Sub
'
'Private Sub cmdToggleCurrentGenerateByDefault_Click()
'    ' Reverses the GenerateByDefault value for the current configuration record.
'    On Error GoTo ErrHandler
'    Dim frm As Access.Form
'    Dim ConfigID As Long
'    ' Save and validate the synchronized child forms, then return the bound SingleForm.
'    Set frm = PrepareCurrentConfigAction(ConfigID)
'    If frm Is Nothing Then Exit Sub
'    ' Change GenerateByDefault from True to False or from False to True.
'    frm!GenerateByDefault = Not CBool(Nz(frm!GenerateByDefault, False))
'    If frm.Dirty Then frm.Dirty = False
'    ' Refresh both child forms and return them to the same configuration record when it remains visible.
'    RefreshDSFormToConfigID ConfigID
'    Exit Sub
'ErrHandler:
'    ShowGeneratedRuntimeError "cmdToggleCurrentGenerateByDefault_Click", Err.Number, Err.Source, Err.Description, "ConfigID: " & CStr(ConfigID)
'End Sub
'
'Private Sub cmdAcceptCurrentCandidate_Click()
'    ' Accepts the current candidate and promotes it into configured status.
'    On Error GoTo ErrHandler
'    Dim frm As Access.Form
'    Dim ConfigID As Long
'    ' Save and validate the synchronized child forms, then return the bound SingleForm.
'    Set frm = PrepareCurrentConfigAction(ConfigID)
'    If frm Is Nothing Then Exit Sub
'    ' Accept the candidate represented by the current record.
'    AcceptCandidate ConfigID, True
'    ' Refresh both child forms and return them to the resulting record when the same ID remains visible.
'    RefreshDSFormToConfigID ConfigID
'    Exit Sub
'ErrHandler:
'    ShowGeneratedRuntimeError "cmdAcceptCurrentCandidate_Click", Err.Number, Err.Source, Err.Description, "ConfigID: " & CStr(ConfigID)
'End Sub
'
'Private Sub cmdIgnoreCurrentCandidate_Click()
'    ' Marks the current candidate as ignored so it is not accepted for generation.
'    On Error GoTo ErrHandler
'    Dim frm As Access.Form
'    Dim ConfigID As Long
'    ' Save and validate the synchronized child forms, then return the bound SingleForm.
'    Set frm = PrepareCurrentConfigAction(ConfigID)
'    If frm Is Nothing Then Exit Sub
'    ' Ignore the candidate represented by the current record.
'    frm.Recordset.Edit
'    frm.Recordset.fields("CandidateStatus").value = CANDIDATE_STATUS_REJECTED
'    frm.Recordset.fields("SelectedForAcceptance").value = False
'    frm.Recordset.fields("ManualReviewStatus").value = MANUAL_REVIEW_STATUS_IGNORE
'    frm.Recordset.fields("RejectedOn").value = Now
'    frm.Recordset.Update
'    ' Refresh both child forms and return them to the same record when it remains in the current view.
'    RefreshDSFormToConfigID ConfigID
'    Exit Sub
'ErrHandler:
'    ShowGeneratedRuntimeError "cmdIgnoreCurrentCandidate_Click", Err.Number, Err.Source, Err.Description, "ConfigID: " & CStr(ConfigID)
'End Sub
'
'Private Sub cmdShowCurrentRecordSummary_Click()
'    ' Displays a summary of the current configuration or candidate record without changing it.
'    On Error GoTo ErrHandler
'    Dim frm As Access.Form
'    Dim ConfigID As Long
'    ' Save and validate the synchronized child forms, then return the bound SingleForm.
'    Set frm = PrepareCurrentConfigAction(ConfigID)
'    If frm Is Nothing Then Exit Sub
'    ' Display the current record summary without changing its review state.
'    ShowCurrentRecordSummary frm
'    Exit Sub
'ErrHandler:
'    ShowGeneratedRuntimeError "cmdShowCurrentRecordSummary_Click", Err.Number, Err.Source, Err.Description, "ConfigID: " & CStr(ConfigID)
'End Sub
'
'Private Sub cmdSetReviewNeedsReview_Click()
'    ' Marks the current record as still requiring manual review.
'    On Error GoTo ErrHandler
'    SetCurrentManualReviewStatus REVIEW_STATUS_NEEDS_REVIEW, "cmdSetReviewNeedsReview_Click"
'    Exit Sub
'ErrHandler:
'    ShowGeneratedRuntimeError "cmdSetReviewNeedsReview_Click", Err.Number, Err.Source, Err.Description
'End Sub
'
'Private Sub cmdSetReviewReviewed_Click()
'    ' Marks the current record as manually reviewed.
'    On Error GoTo ErrHandler
'    SetCurrentManualReviewStatus REVIEW_STATUS_REVIEWED, "cmdSetReviewReviewed_Click"
'    Exit Sub
'ErrHandler:
'    ShowGeneratedRuntimeError "cmdSetReviewReviewed_Click", Err.Number, Err.Source, Err.Description
'End Sub
'
'Private Sub cmdSetReviewDeferred_Click()
'    ' Marks the current record as deferred for review at a later time.
'    On Error GoTo ErrHandler
'    SetCurrentManualReviewStatus REVIEW_STATUS_DEFERRED, "cmdSetReviewDeferred_Click"
'    Exit Sub
'ErrHandler:
'    ShowGeneratedRuntimeError "cmdSetReviewDeferred_Click", Err.Number, Err.Source, Err.Description
'End Sub
'
'Private Sub cmdSetReviewIgnore_Click()
'    ' Marks the current record as intentionally ignored during manual review.
'    On Error GoTo ErrHandler
'    SetCurrentManualReviewStatus REVIEW_STATUS_IGNORE, "cmdSetReviewIgnore_Click"
'    Exit Sub
'ErrHandler:
'    ShowGeneratedRuntimeError "cmdSetReviewIgnore_Click", Err.Number, Err.Source, Err.Description
'End Sub
'
'Private Sub cmdSetReviewBlocked_Click()
'    ' Marks the current record as blocked from completing manual review.
'    On Error GoTo ErrHandler
'    SetCurrentManualReviewStatus REVIEW_STATUS_BLOCKED, "cmdSetReviewBlocked_Click"
'    Exit Sub
'ErrHandler:
'    ShowGeneratedRuntimeError "cmdSetReviewBlocked_Click", Err.Number, Err.Source, Err.Description
'End Sub
'
'Private Sub SetCurrentManualReviewStatus(ByVal reviewStatus As String, ByVal callerProcedure As String)
'    ' Applies one supported manual-review status through the bound SingleForm and refreshes both views.
'    On Error GoTo ErrHandler
'    Dim frm As Access.Form
'    Dim ConfigID As Long
'    Set frm = PrepareCurrentConfigAction(ConfigID)
'    If frm Is Nothing Then Exit Sub
'    frm.Recordset.Edit
'    frm.Recordset.fields("ManualReviewStatus").value = reviewStatus
'    frm.Recordset.fields("ManualReviewUpdatedOn").value = Now
'    If StrComp(reviewStatus, REVIEW_STATUS_REVIEWED, vbTextCompare) = 0 Then frm.Recordset.fields("ManualReviewedOn").value = Now
'    frm.Recordset.Update
'    RefreshDSFormToConfigID ConfigID
'    Exit Sub
'ErrHandler:
'    Err.Raise Err.Number, Me.Name & ".SetCurrentManualReviewStatus", "Could not apply manual-review status '" & reviewStatus & "' for caller '" & callerProcedure & "'." & vbCrLf & Err.Description
'End Sub
'
'Private Sub ShowCurrentRecordSummary(ByVal frm As Access.Form)
'    ' Displays the current bound Config/candidate row without changing review or execution state.
'    On Error GoTo ErrHandler
'    Dim summaryText As String
'    If frm Is Nothing Then Err.Raise vbObjectError + 40, Me.Name & ".ShowCurrentRecordSummary", "Current form was not supplied."
'    summaryText = "ID: " & CStr(Nz(frm!ID, 0)) & vbCrLf & _
'        "RowKind: " & CStr(Nz(frm!RowKind, vbNullString)) & vbCrLf & _
'        "MacroKey: " & CStr(Nz(frm!MacroKey, vbNullString)) & vbCrLf & _
'        "CandidateMacroKey: " & CStr(Nz(frm!CandidateMacroKey, vbNullString)) & vbCrLf & _
'        "MacroTitle: " & CStr(Nz(frm!MacroTitle, vbNullString)) & vbCrLf & _
'        "MacroAction: " & CStr(Nz(frm!MacroAction, vbNullString)) & vbCrLf & _
'        "ConfigurationStatus: " & CStr(Nz(frm!ConfigurationStatus, vbNullString)) & vbCrLf & _
'        "CandidateStatus: " & CStr(Nz(frm!CandidateStatus, vbNullString)) & vbCrLf & _
'        "ManualReviewStatus: " & CStr(Nz(frm!ManualReviewStatus, vbNullString)) & vbCrLf & _
'        "TargetProcedureName: " & CStr(Nz(frm!TargetProcedureName, vbNullString)) & vbCrLf & _
'        "TargetObjectName: " & CStr(Nz(frm!TargetObjectName, vbNullString))
'    MsgBox summaryText, vbInformation, "ADC Config Record"
'    Exit Sub
'ErrHandler:
'    Err.Raise Err.Number, Me.Name & ".ShowCurrentRecordSummary", "Could not display the current configuration/candidate record summary." & vbCrLf & Err.Description
'End Sub
'
'Private Function PrepareCurrentConfigAction(ByRef ConfigID As Long) As Access.Form
'    ' Returns the bound SingleForm only after both synchronized views are saved, refreshed, and verified on the same ID.
'    On Error GoTo ErrHandler
'    Dim DSForm As Access.Form
'    Dim SingleForm As Access.Form
'    Dim currentKey As Variant
'    If Not GeneratedSyncReady() Then InitSubformSync
'    If Not GeneratedSyncReady() Then Err.Raise vbObjectError + 30, Me.Name & ".PrepareCurrentConfigAction", "Subform synchronization is not ready."
'    Set SingleForm = Me.Controls(SINGLE_SUBFORM_CONTROL_NAME).Form
'    Set DSForm = Me.Controls(DS_SUBFORM_CONTROL_NAME).Form
'    If SingleForm Is Nothing Then Err.Raise vbObjectError + 31, Me.Name & ".PrepareCurrentConfigAction", "The SingleForm subform is unavailable."
'    If DSForm Is Nothing Then Err.Raise vbObjectError + 32, Me.Name & ".PrepareCurrentConfigAction", "The DSForm subform is unavailable."
'    If SingleForm.Dirty And DSForm.Dirty Then Err.Raise vbObjectError + 33, Me.Name & ".PrepareCurrentConfigAction", "Both synchronized child forms contain unsaved edits. Save or cancel one edit before running a record action so one version cannot overwrite the other."
'    SaveGeneratedFormIfDirty DSForm, ROLE_DS
'    SaveGeneratedFormIfDirty SingleForm, ROLE_SINGLE
'    If Not TryGeneratedFormCurrentKey(SingleForm, currentKey) Then
'        If Not TryGeneratedFormCurrentKey(DSForm, currentKey) Then Err.Raise vbObjectError + 34, Me.Name & ".PrepareCurrentConfigAction", "No current AOC_Config record is available."
'        SyncFromDS
'        If Not TryGeneratedFormCurrentKey(SingleForm, currentKey) Then Err.Raise vbObjectError + 35, Me.Name & ".PrepareCurrentConfigAction", "SingleForm did not synchronize to the current DSForm record."
'    End If
'    If IsNull(currentKey) Then Err.Raise vbObjectError + 36, Me.Name & ".PrepareCurrentConfigAction", "The current configuration ID is Null."
'    ConfigID = CLng(currentKey)
'    RequerySyncedSubforms True, ROLE_SINGLE
'    AssertGeneratedFormCurrentKey SingleForm, ConfigID, "button-action SingleForm preparation"
'    AssertGeneratedFormCurrentKey DSForm, ConfigID, "button-action DSForm preparation"
'    Set PrepareCurrentConfigAction = SingleForm
'    Exit Function
'ErrHandler:
'    Set PrepareCurrentConfigAction = Nothing
'    Err.Raise Err.Number, Me.Name & ".PrepareCurrentConfigAction", "Could not prepare the current synchronized configuration record for a parent-form action." & vbCrLf & Err.Description
'End Function
'
'Private Sub SaveGeneratedFormIfDirty(ByVal targetForm As Access.Form, ByVal SourceRole As String)
'    ' Commits pending edits on one generated child form before a button action reads the record through SingleForm.
'    On Error GoTo ErrHandler
'    If targetForm Is Nothing Then Err.Raise vbObjectError + 37, Me.Name & ".SaveGeneratedFormIfDirty", "Target form was not supplied."
'    If targetForm.Dirty Then targetForm.Dirty = False
'    Exit Sub
'ErrHandler:
'    Err.Raise Err.Number, Me.Name & ".SaveGeneratedFormIfDirty", "Could not save pending edits on generated child form '" & SafeGeneratedFormName(targetForm) & "' for source role '" & SourceRole & "'." & vbCrLf & Err.Description
'End Sub
'
'Private Sub RefreshDSFormToConfigID(ByVal ConfigID As Long)
'    ' Requeries both generated child forms and restores the requested configuration ID when it remains visible.
'    On Error GoTo ErrHandler
'    Dim DSForm As Access.Form
'    Dim SingleForm As Access.Form
'    Dim foundInDS As Boolean
'    Dim foundInSingle As Boolean
'    If Not GeneratedSyncReady() Then InitSubformSync
'    If Not GeneratedSyncReady() Then Err.Raise vbObjectError + 38, Me.Name & ".RefreshDSFormToConfigID", "Subform synchronization is not ready."
'    Set SingleForm = Me.Controls(SINGLE_SUBFORM_CONTROL_NAME).Form
'    Set DSForm = Me.Controls(DS_SUBFORM_CONTROL_NAME).Form
'    RequerySyncedSubforms False, ROLE_SINGLE
'    mSuppressGeneratedChildCurrentSync = True
'    foundInSingle = TryPositionGeneratedFormByKey(SingleForm, ConfigID)
'    foundInDS = TryPositionGeneratedFormByKey(DSForm, ConfigID)
'    mSuppressGeneratedChildCurrentSync = False
'    If foundInSingle Then
'        SyncFromSingle
'    ElseIf foundInDS Then
'        SyncFromDS
'    Else
'        Debug.Print Me.Name & ": Refreshed synchronized child forms, but ID " & CStr(ConfigID) & " is no longer visible in either current record source."
'        SysCmd acSysCmdSetStatus, Me.Name & " refreshed; configuration ID " & CStr(ConfigID) & " is no longer visible in the current view."
'    End If
'    Exit Sub
'ErrHandler:
'    mSuppressGeneratedChildCurrentSync = False
'    Err.Raise Err.Number, Me.Name & ".RefreshDSFormToConfigID", "Could not refresh synchronized child forms after action for configuration ID " & CStr(ConfigID) & "." & vbCrLf & Err.Description
'End Sub
'
'Private Function SafeGeneratedFormName(ByVal targetForm As Access.Form) As String
'    ' Returns a diagnostic form name without replacing the original error when a form reference is unavailable.
'    On Error GoTo ExitHere
'    If targetForm Is Nothing Then
'        SafeGeneratedFormName = "(none)"
'    Else
'        SafeGeneratedFormName = targetForm.Name
'    End If
'ExitHere:
'    If Len(SafeGeneratedFormName) = 0 Then SafeGeneratedFormName = "(none)"
'End Function
'
'Private Function TryPositionGeneratedFormByKey(ByVal targetForm As Form, ByVal keyValue As Variant) As Boolean
'    On Error GoTo ErrHandler
'    Dim rs As Object
'    If targetForm Is Nothing Then Err.Raise vbObjectError + 8, Me.Name & ".TryPositionGeneratedFormByKey", "Target form was not supplied."
'    Set rs = targetForm.RecordsetClone
'    If rs Is Nothing Then Err.Raise vbObjectError + 39, Me.Name & ".TryPositionGeneratedFormByKey", "Target form did not expose a RecordsetClone."
'    If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
'    Do While Not rs.EOF
'        If GeneratedKeysEqual(rs.fields(KEY_FIELD_NAME).value, keyValue) Then
'            targetForm.Bookmark = rs.Bookmark
'            TryPositionGeneratedFormByKey = True
'            Exit Do
'        End If
'        rs.MoveNext
'    Loop
'    rs.Close
'    Set rs = Nothing
'    Exit Function
'ErrHandler:
'    Dim errNumber As Long
'    Dim errDescription As String
'    errNumber = Err.Number
'    errDescription = Err.Description
'    On Error Resume Next
'    If Not rs Is Nothing Then rs.Close
'    Set rs = Nothing
'    On Error GoTo 0
'    Err.Raise errNumber, Me.Name & ".TryPositionGeneratedFormByKey", errDescription
'End Function
'
'Private Sub ADC1Btn_Click()
'    ' Run without confirmation.
'    cmdRunCurrentConfig_Click
'End Sub
'
'Private Sub ADC2Btn_Click()
'    ' Run through the confirmation workflow.
'    cmdRunCurrentConfigConfirmed_Click
'End Sub
'
'Private Sub ADC3Btn_Click()
'    ' Generate or regenerate the current configuration.
'    cmdGenerateCurrentConfig_Click
'End Sub
'
'Private Sub ADC4Btn_Click()
'    ' Validate the current configuration.
'    cmdValidateCurrentConfig_Click
'End Sub
'
'Private Sub ADC5Btn_Click()
'    ' Resolve the current configuration's target.
'    cmdResolveCurrentConfig_Click
'End Sub
'
'Private Sub ADC6Btn_Click()
'    ' Toggle the current record's Active value.
'    cmdToggleCurrentActive_Click
'End Sub
'
'Private Sub ADC7Btn_Click()
'    ' Toggle the current record's Generate By Default value.
'    cmdToggleCurrentGenerateByDefault_Click
'End Sub
'
'Private Sub ADC8Btn_Click()
'    ' Accept the current candidate.
'    cmdAcceptCurrentCandidate_Click
'End Sub
'
'Private Sub ADC9Btn_Click()
'    ' Ignore the current candidate.
'    cmdIgnoreCurrentCandidate_Click
'End Sub
'
'Private Sub ADC10Btn_Click()
'    ' Display the current record summary.
'    cmdShowCurrentRecordSummary_Click
'End Sub
'
'Private Sub ADC11Btn_Click()
'    ' Mark the current record as needing review.
'    cmdSetReviewNeedsReview_Click
'End Sub
'
'Private Sub ADC12Btn_Click()
'    ' Mark the current record as reviewed.
'    cmdSetReviewReviewed_Click
'End Sub
'
'Private Sub ADC13Btn_Click()
'    ' Defer review of the current record.
'    cmdSetReviewDeferred_Click
'End Sub
'
'Private Sub ADC14Btn_Click()
'    ' Ignore the current record during review.
'    cmdSetReviewIgnore_Click
'End Sub
'
'Private Sub ADC15Btn_Click()
'    ' Mark the current record as blocked.
'    cmdSetReviewBlocked_Click
'End Sub

' ================================================================================================================================================================================
' STANDARD MODULE: Shared Code
' VB_Name: x__SHARED_CODE

'Option Compare Database
'Option Explicit
'
'' ================================================================================================================================================================================
'' GENERAL MODULE-LEVEL DECLARATIONS
'' ================================================================================================================================================================================
'' ================================================================================================================================================================================
'' DAO MODULE-LEVEL DECLARATIONS
'' ================================================================================================================================================================================
''Access linked-table Connect strings include this marker before the linked database path.
''Used by SharedDao_IsAccessLinkedTable, SharedDao_LinkedTableDataPath, SharedDao_AccessDatabasePathFromConnect, and SharedDao_SingleLinkedDatabasePath.
'Private Const SHARED_DAO_ACCESS_CONNECT_DATABASE_MARKER As String = ";DATABASE="
'Private Const SHARED_DAO_ERROR_PROPERTY_NOT_FOUND As Long = 3270
'Private Const SHARED_ADODB_STREAM_TYPE_TEXT As Long = 2
'Private Const SHARED_ADODB_STREAM_READ_ALL As Long = -1
'Private Const SHARED_ADODB_STREAM_CHARSET_UTF8 As String = "utf-8"
'
'' ================================================================================================================================================================================
'' FORM DESIGN MODULE-LEVEL DECLARATIONS
'' ================================================================================================================================================================================
'Private Const SHARED_FORM_DESIGN_RIGHT_LIMIT As Long = 31000
'Private Const SHARED_FORM_DESIGN_BOTTOM_LIMIT As Long = 31000
'Private Const SHARED_FORM_DESIGN_LOCKED_BACK As Long = 15592941
'Private mFormDesignReplacementBackups As Object
'
'' ================================================================================================================================================================================
'' FORMS MODULE-LEVEL DECLARATIONS
'' ================================================================================================================================================================================
'Private Const SHARED_FORM_ZOOM_MIN_FONT As Integer = 6
'Private Const SHARED_FORM_ZOOM_MAX_FONT As Integer = 48
'Private mFormZoomGeometry As Object
'Private mFormZoomFonts As Object
'Private mFormZoomLevels As Object
'
'' ================================================================================================================================================================================
'' TREEVIEW MODULE-LEVEL DECLARATIONS
'' ================================================================================================================================================================================
'Public Const TVW_RELATIONSHIP_CHILD As Long = 4
'Public Const TVW_LABEL_EDIT_AUTOMATIC As Long = 0
'Public Const TVW_LABEL_EDIT_MANUAL As Long = 1
'
'' ================================================================================================================================================================================
''GENERAL SHARED PROCEDURES
'' ================================================================================================================================================================================
'Public Function SharedText_DoubleQuote(ByVal value As String) As String
'    'Wraps a string in double quotes and escapes embedded double quotes.
'    'Use when passing paths or arguments to Shell commands.
'    SharedText_DoubleQuote = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
'End Function
'
'Public Function SharedText_Nz(ByVal v As Variant) As String
'    'Returns an empty string for Null, otherwise returns CStr(value).
'    'Use when building display text or string comparisons that should treat Null as blank.
'    If IsNull(v) Then
'        SharedText_Nz = vbNullString
'    Else
'        SharedText_Nz = CStr(v)
'    End If
'End Function
'
'Public Function SharedText_HasText(ByVal v As Variant) As Boolean
'    'Returns True when a value contains non-whitespace text.
'    'Consolidated old names: local TV DATA GENERATOR MODULE HasText.
'    SharedText_HasText = (Len(Trim$(SharedText_Nz(v))) > 0)
'End Function
'
'Public Function SharedDao_EscapeSqlText(ByVal value As String) As String
'    'Escapes apostrophes for insertion into a quoted Access SQL text literal.
'    'Use inside larger SQL expressions when the surrounding single quotes are added elsewhere.
'    SharedDao_EscapeSqlText = Replace(Nz(value, vbNullString), "'", "''")
'End Function
'
'Public Function SharedDao_AccessBooleanLiteral(ByVal value As Boolean) As String
'    'Returns the Access SQL Boolean literal True or False.
'    'Use when building SQL where a Boolean literal is needed.
'    If value Then
'        SharedDao_AccessBooleanLiteral = "True"
'    Else
'        SharedDao_AccessBooleanLiteral = "False"
'    End If
'End Function
'
'Public Function SharedText_AfterFirstColon(ByVal sourceText As String) As String
'    'Intended use: Returns the text after the first colon for parsers that consume labeled diagnostic or metadata lines.
'    Dim colonPos As Long
'    colonPos = InStr(1, sourceText, ":", vbBinaryCompare)
'    If colonPos = 0 Then
'        SharedText_AfterFirstColon = vbNullString
'    Else
'        SharedText_AfterFirstColon = Mid$(sourceText, colonPos + 1)
'    End If
'End Function
'
'Public Function SharedVba_PublicProcedureNameFromLine(ByVal lineText As String) As String
'    'Intended use: Extracts a Public procedure name from one cleaned VBA declaration line for header and reference scanning.
'    Dim s As String
'    s = SharedText_CollapseWhitespace(Trim$(lineText))
'    If Len(s) = 0 Then Exit Function
'    If SharedText_StartsWithWord(s, "Attribute") Then Exit Function
'    If SharedText_StartsWithWord(s, "Option") Then Exit Function
'    If SharedText_StartsWithWord(s, "Private") Then Exit Function
'    If SharedText_StartsWithWord(s, "Declare") Then Exit Function
'    If SharedText_StartsWithWord(s, "Public") Then
'        s = Trim$(Mid$(s, Len("Public") + 1))
'    End If
'    If SharedText_StartsWithWord(s, "Friend") Then
'        s = Trim$(Mid$(s, Len("Friend") + 1))
'    End If
'    If SharedText_StartsWithWord(s, "Static") Then
'        s = Trim$(Mid$(s, Len("Static") + 1))
'    End If
'    If SharedText_StartsWithWord(s, "Declare") Then Exit Function
'    If SharedText_StartsWithWord(s, "Sub") Then
'        SharedVba_PublicProcedureNameFromLine = SharedVba_IdentifierAfterKeyword(s, "Sub")
'        Exit Function
'    End If
'    If SharedText_StartsWithWord(s, "Function") Then
'        SharedVba_PublicProcedureNameFromLine = SharedVba_IdentifierAfterKeyword(s, "Function")
'        Exit Function
'    End If
'    If SharedText_StartsWithWord(s, "Property") Then
'        s = Trim$(Mid$(s, Len("Property") + 1))
'        If SharedText_StartsWithWord(s, "Get") Then
'            SharedVba_PublicProcedureNameFromLine = SharedVba_IdentifierAfterKeyword(s, "Get")
'        ElseIf SharedText_StartsWithWord(s, "Let") Then
'            SharedVba_PublicProcedureNameFromLine = SharedVba_IdentifierAfterKeyword(s, "Let")
'        ElseIf SharedText_StartsWithWord(s, "Set") Then
'            SharedVba_PublicProcedureNameFromLine = SharedVba_IdentifierAfterKeyword(s, "Set")
'        End If
'        Exit Function
'    End If
'End Function
'
'Public Function SharedVba_IdentifierAfterKeyword(ByVal lineText As String, ByVal keywordText As String) As String
'    'Intended use: Extracts the identifier following a VBA keyword while parsing declarations.
'    Dim s As String
'    Dim i As Long
'    Dim ch As String
'    Dim resultText As String
'    s = Trim$(Mid$(lineText, Len(keywordText) + 1))
'    For i = 1 To Len(s)
'        ch = Mid$(s, i, 1)
'        If SharedVba_IsIdentifierChar(ch) Then
'            resultText = resultText & ch
'        Else
'            Exit For
'        End If
'    Next i
'    SharedVba_IdentifierAfterKeyword = resultText
'End Function
'
'Public Function SharedVba_CleanCodeLine(ByVal sourceLine As String) As String
'    'Intended use: Normalizes one VBA source line before declaration or token analysis.
'    Dim s As String
'    s = SharedVba_StripStringsAndComments(sourceLine)
'    s = Replace(s, vbTab, " ")
'    s = SharedText_CollapseWhitespace(s)
'    If SharedText_StartsWithWord(Trim$(s), "Rem") Then
'        SharedVba_CleanCodeLine = vbNullString
'    Else
'        SharedVba_CleanCodeLine = s
'    End If
'End Function
'
'Public Function SharedVba_StripStringsAndComments(ByVal sourceLine As String) As String
'    'Intended use: Removes string-literal and comment content before safe VBA token scanning.
'    Dim i As Long
'    Dim ch As String
'    Dim resultText As String
'    Dim inString As Boolean
'    For i = 1 To Len(sourceLine)
'        ch = Mid$(sourceLine, i, 1)
'        If inString Then
'            If ch = """" Then
'                If i < Len(sourceLine) Then
'                    If Mid$(sourceLine, i + 1, 1) = """" Then
'                        i = i + 1
'                    Else
'                        inString = False
'                    End If
'                Else
'                    inString = False
'                End If
'            End If
'            resultText = resultText & " "
'        Else
'            If ch = """" Then
'                inString = True
'                resultText = resultText & " "
'            ElseIf ch = "'" Then
'                Exit For
'            Else
'                resultText = resultText & ch
'            End If
'        End If
'    Next i
'    SharedVba_StripStringsAndComments = resultText
'End Function
'
'Public Function SharedVba_ContainsIdentifierToken(ByVal sourceTextUpper As String, ByVal tokenText As String) As Boolean
'    'Intended use: Determines whether cleaned VBA source contains a complete identifier token rather than a substring.
'    Dim tokenUpper As String
'    Dim startPos As Long
'    Dim foundPos As Long
'    Dim beforeChar As String
'    Dim afterChar As String
'    Dim afterPos As Long
'    tokenUpper = UCase$(tokenText)
'    startPos = 1
'    Do
'        foundPos = InStr(startPos, sourceTextUpper, tokenUpper, vbBinaryCompare)
'        If foundPos = 0 Then
'            SharedVba_ContainsIdentifierToken = False
'            Exit Function
'        End If
'        If foundPos > 1 Then
'            beforeChar = Mid$(sourceTextUpper, foundPos - 1, 1)
'        Else
'            beforeChar = vbNullString
'        End If
'        afterPos = foundPos + Len(tokenUpper)
'        If afterPos <= Len(sourceTextUpper) Then
'            afterChar = Mid$(sourceTextUpper, afterPos, 1)
'        Else
'            afterChar = vbNullString
'        End If
'        If Not SharedVba_IsIdentifierChar(beforeChar) And Not SharedVba_IsIdentifierChar(afterChar) Then
'            SharedVba_ContainsIdentifierToken = True
'            Exit Function
'        End If
'        startPos = foundPos + Len(tokenUpper)
'    Loop
'End Function
'
'Public Function SharedCollection_JoinSortedDictionaryKeys(ByVal sourceDictionary As Object) As String
'    'Intended use: Returns dictionary keys in sorted display order for deterministic diagnostics and generated text.
'    Dim arr() As String
'    Dim keyValue As Variant
'    Dim i As Long
'    If sourceDictionary.count = 0 Then
'        SharedCollection_JoinSortedDictionaryKeys = "None"
'        Exit Function
'    End If
'    ReDim arr(0 To sourceDictionary.count - 1)
'    i = 0
'    For Each keyValue In sourceDictionary.keys
'        arr(i) = CStr(keyValue)
'        i = i + 1
'    Next keyValue
'    SharedArray_SortStringArray arr
'    SharedCollection_JoinSortedDictionaryKeys = Join(arr, ", ")
'End Function
'
'Public Sub SharedArray_SortStringArray(ByRef arr() As String)
'    'Intended use: Sorts a String array in place when deterministic output ordering is required.
'    Dim i As Long
'    Dim j As Long
'    Dim tempValue As String
'    For i = LBound(arr) To UBound(arr) - 1
'        For j = i + 1 To UBound(arr)
'            If StrComp(arr(j), arr(i), vbTextCompare) < 0 Then
'                tempValue = arr(i)
'                arr(i) = arr(j)
'                arr(j) = tempValue
'            End If
'        Next j
'    Next i
'End Sub
'
'Public Function SharedText_CollapseWhitespace(ByVal sourceText As String) As String
'    'Intended use: Collapses repeated whitespace for normalized comparison or display text.
'    Dim resultText As String
'    resultText = sourceText
'    Do While InStr(1, resultText, "  ", vbBinaryCompare) > 0
'        resultText = Replace(resultText, "  ", " ")
'    Loop
'    SharedText_CollapseWhitespace = Trim$(resultText)
'End Function
'
'Public Function SharedText_StartsWithText(ByVal sourceText As String, ByVal prefixText As String) As Boolean
'    'Intended use: Determines whether text begins with a specified prefix using case-insensitive comparison.
'    SharedText_StartsWithText = (StrComp(Left$(sourceText, Len(prefixText)), prefixText, vbTextCompare) = 0)
'End Function
'
'Public Function SharedText_StartsWithWord(ByVal sourceText As String, ByVal wordText As String) As Boolean
'    'Intended use: Determines whether text begins with a complete word rather than a partial identifier.
'    Dim nextChar As String
'    If Len(sourceText) < Len(wordText) Then
'        SharedText_StartsWithWord = False
'        Exit Function
'    End If
'    If StrComp(Left$(sourceText, Len(wordText)), wordText, vbTextCompare) <> 0 Then
'        SharedText_StartsWithWord = False
'        Exit Function
'    End If
'    If Len(sourceText) = Len(wordText) Then
'        SharedText_StartsWithWord = True
'        Exit Function
'    End If
'    nextChar = Mid$(sourceText, Len(wordText) + 1, 1)
'    SharedText_StartsWithWord = Not SharedVba_IsIdentifierChar(nextChar)
'End Function
'
'Public Function SharedVba_IsIdentifierChar(ByVal value As String) As Boolean
'    'Intended use: Determines whether one character is valid inside a VBA identifier token.
'    Dim codePoint As Long
'    If Len(value) = 0 Then
'        SharedVba_IsIdentifierChar = False
'        Exit Function
'    End If
'    codePoint = AscW(Left$(value, 1))
'    Select Case codePoint
'        Case 48 To 57, 65 To 90, 97 To 122, 95
'            SharedVba_IsIdentifierChar = True
'        Case Else
'            SharedVba_IsIdentifierChar = False
'    End Select
'End Function
'
'Public Function SharedDao_SqlLiteralFromValue(ByVal value As Variant) As String
'    'Converts a Variant to an Access SQL literal using locale-invariant numeric formatting.
'    'Consolidated old names: TVSettings_SqlValue, local TV NSS IMPORT MODULE SqlValue.
'    If IsNull(value) Then
'        SharedDao_SqlLiteralFromValue = "Null"
'    ElseIf VarType(value) = vbBoolean Then
'        SharedDao_SqlLiteralFromValue = IIf(CBool(value), "True", "False")
'    ElseIf VarType(value) = vbDate Then
'        SharedDao_SqlLiteralFromValue = "#" & Format$(CDate(value), "yyyy-mm-dd hh:nn:ss") & "#"
'    ElseIf IsNumeric(value) And VarType(value) <> vbString Then
'        SharedDao_SqlLiteralFromValue = SharedDao_InvariantNumericLiteral(value)
'    Else
'        SharedDao_SqlLiteralFromValue = SharedDao_SqlQuote(value)
'    End If
'End Function
'
'Private Function SharedDao_InvariantNumericLiteral(ByVal value As Variant) As String
'    'Returns a numeric literal with a period decimal separator regardless of the Windows locale.
'    On Error GoTo ErrHandler
'    SharedDao_InvariantNumericLiteral = Trim$(Str(value))
'    If Len(SharedDao_InvariantNumericLiteral) = 0 Then Err.Raise vbObjectError + 7630, "SharedDao_InvariantNumericLiteral", "The numeric SQL literal is blank."
'    SharedDao_InvariantNumericLiteral = Replace(SharedDao_InvariantNumericLiteral, "D", "E", 1, -1, vbTextCompare)
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedDao_InvariantNumericLiteral", "Unable to format a locale-invariant numeric SQL literal." & vbCrLf & "Value type: " & typeName(value) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'
'Public Function SharedCollection_KeyExists(ByVal values As Collection, ByVal key As String) As Boolean
'    'Returns True when a keyed Collection contains the requested key, including when the stored item is an object.
'    'Use for duplicate checks without coercing object-valued Collection items through a default property.
'    On Error GoTo NotFound
'    Dim ScalarValue As Variant
'    Dim objectValue As Object
'    If values Is Nothing Then Exit Function
'    If IsObject(values.item(key)) Then
'        Set objectValue = values.item(key)
'    Else
'        ScalarValue = values.item(key)
'    End If
'    SharedCollection_KeyExists = True
'    Exit Function
'NotFound:
'    SharedCollection_KeyExists = False
'End Function
'
'Public Function SharedError_FormatDetails(ByVal ProcedureName As String, Optional ByVal contextText As String = vbNullString) As String
'    'Builds a standard multi-line error message from the current Err object.
'    'Use inside an error handler before Err is cleared or overwritten.
'    SharedError_FormatDetails = ProcedureName & " failed:" & vbCrLf
'    If Len(contextText) > 0 Then SharedError_FormatDetails = SharedError_FormatDetails & contextText & vbCrLf
'    SharedError_FormatDetails = SharedError_FormatDetails & "Err.Number: " & CStr(Err.Number) & vbCrLf & "Err.Source: " & Err.Source & vbCrLf & "Err.Description: " & Err.Description
'End Function
'
'Public Function SharedPath_FullPath(ByVal filePath As String) As String
'    'Returns the FileSystemObject absolute path for a file path.
'    'Use before comparing paths that may be relative, differently cased, or differently formatted.
'    On Error GoTo ErrHandler
'    SharedPath_FullPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(filePath)
'    Exit Function
'ErrHandler:
'    SharedPath_FullPath = filePath
'End Function
'
'Private Function SharedPath_FullPathStrict(ByVal filePath As String, ByVal callerName As String) As String
'    'Resolves one path for identity or target verification and raises when canonicalization cannot be completed.
'    On Error GoTo ErrHandler
'    Dim resolvedPath As String
'    filePath = Trim$(filePath)
'    callerName = Trim$(callerName)
'    If Len(callerName) = 0 Then callerName = "SharedPath_FullPathStrict"
'    If Len(filePath) = 0 Then Err.Raise vbObjectError + 8000, callerName, "The path is blank and cannot be resolved for an exact comparison."
'    resolvedPath = Trim$(CStr(CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(filePath)))
'    If Len(resolvedPath) = 0 Then Err.Raise vbObjectError + 8001, callerName, "The resolved absolute path is blank." & vbCrLf & "Input path: " & filePath
'    SharedPath_FullPathStrict = resolvedPath
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedPath_FullPathStrict", "Unable to resolve a path required for an exact identity comparison." & vbCrLf & "Caller: " & callerName & vbCrLf & "Input path: " & filePath & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'
'Public Function SharedPath_UniqueFilePath(ByVal filePath As String) As String
'    'Returns a non-existing file path by appending _1, _2, etc. when needed.
'    'Use before creating backup/export/temp files that should not overwrite existing files.
'    Dim fso As Object
'    Dim folderPath As String
'    Dim baseName As String
'    Dim extensionName As String
'    Dim testPath As String
'    Dim n As Long
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    folderPath = fso.GetParentFolderName(filePath)
'    baseName = fso.GetBaseName(filePath)
'    extensionName = fso.GetExtensionName(filePath)
'    testPath = filePath
'    Do While Len(Dir$(testPath)) > 0
'        n = n + 1
'        If Len(folderPath) > 0 Then
'            If Len(extensionName) > 0 Then
'                testPath = folderPath & "\" & baseName & "_" & CStr(n) & "." & extensionName
'            Else
'                testPath = folderPath & "\" & baseName & "_" & CStr(n)
'            End If
'        Else
'            If Len(extensionName) > 0 Then
'                testPath = baseName & "_" & CStr(n) & "." & extensionName
'            Else
'                testPath = baseName & "_" & CStr(n)
'            End If
'        End If
'    Loop
'    SharedPath_UniqueFilePath = testPath
'End Function
'
'Public Function SharedPath_IsCurrentProjectFile(ByVal dataPath As String) As Boolean
'    'Returns True when a path points to the currently open Access project file.
'    'Use as a fail-closed safety check before treating a selected ACCDB as an external data file.
'    SharedPath_IsCurrentProjectFile = (StrComp(SharedPath_FullPathStrict(dataPath, "SharedPath_IsCurrentProjectFile"), SharedPath_FullPathStrict(CurrentProject.FullName, "SharedPath_IsCurrentProjectFile"), vbTextCompare) = 0)
'End Function
'
'Public Function SharedPath_FolderExists(ByVal folderPath As String) As Boolean
'    'Returns True when the supplied path exists and is a folder.
'    'Consolidated old names: local TV NSS EXPORT MODULE FolderExists.
'    On Error Resume Next
'    SharedPath_FolderExists = ((GetAttr(folderPath) And vbDirectory) = vbDirectory)
'    Err.Clear
'    On Error GoTo 0
'End Function
'
'Public Function SharedPath_Combine(ByVal folderPath As String, ByVal fileName As String) As String
'    'Combines a folder path and file name using a single path separator.
'    'Consolidated old names: local TV NSS EXPORT MODULE CombinePath.
'    If Right$(folderPath, 1) = "\" Or Right$(folderPath, 1) = "/" Then
'        SharedPath_Combine = folderPath & fileName
'    Else
'        SharedPath_Combine = folderPath & "\" & fileName
'    End If
'End Function
'
'Public Sub SharedFile_WriteUtf8NoBom(ByVal filePath As String, ByVal lines As Collection)
'    'Writes a Collection of lines to a staged UTF-8 file without a BOM, verifies replacement, and restores the prior file on pre-commit failure.
'    'Consolidated old names: local TV NSS EXPORT MODULE WriteUtf8NoBomFile.
'    On Error GoTo ErrHandler
'    Dim fso As Object
'    Dim textStream As Object
'    Dim binaryStream As Object
'    Dim lineValue As Variant
'    Dim tempFilePath As String
'    Dim backupFilePath As String
'    Dim candidateNumber As Long
'    Dim originalMoved As Boolean
'    Dim replacementMoved As Boolean
'    Dim replacementVerified As Boolean
'    Dim expectedSize As Variant
'    Dim savedErrorNumber As Long
'    Dim savedErrorSource As String
'    Dim savedErrorDescription As String
'    Dim cleanupErrorNumber As Long
'    Dim cleanupErrorSource As String
'    Dim cleanupErrorDescription As String
'    Dim cleanupFailureText As String
'    Dim inspectionErrorNumber As Long
'    Dim inspectionErrorSource As String
'    Dim inspectionErrorDescription As String
'    Dim targetVerificationErrorNumber As Long
'    Dim targetVerificationErrorSource As String
'    Dim targetVerificationErrorDescription As String
'    Dim targetSizeVerificationErrorNumber As Long
'    Dim targetSizeVerificationErrorSource As String
'    Dim targetSizeVerificationErrorDescription As String
'    Dim finalExists As Boolean
'    Dim finalSizeMatches As Boolean
'    Dim tempExists As Boolean
'    Dim backupExists As Boolean
'    If Len(Trim$(filePath)) = 0 Then Err.Raise vbObjectError + 7700, "SharedFile_WriteUtf8NoBom", "The target file path is blank."
'    If lines Is Nothing Then Err.Raise vbObjectError + 7701, "SharedFile_WriteUtf8NoBom", "The lines Collection is Nothing."
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    tempFilePath = filePath & ".sharedu8.tmp"
'    Do While fso.fileExists(tempFilePath)
'        candidateNumber = candidateNumber + 1
'        tempFilePath = filePath & ".sharedu8." & CStr(candidateNumber) & ".tmp"
'    Loop
'    Set textStream = CreateObject("ADODB.Stream")
'    textStream.Type = 2
'    textStream.charset = "utf-8"
'    textStream.Open
'    For Each lineValue In lines
'        textStream.WriteText CStr(lineValue), 1
'    Next lineValue
'    Set binaryStream = CreateObject("ADODB.Stream")
'    binaryStream.Type = 1
'    binaryStream.Open
'    If textStream.Size >= 3 Then
'        textStream.position = 3
'    Else
'        textStream.position = 0
'    End If
'    textStream.CopyTo binaryStream
'    binaryStream.SaveToFile tempFilePath, 2
'    binaryStream.Close
'    Set binaryStream = Nothing
'    textStream.Close
'    Set textStream = Nothing
'    If Not fso.fileExists(tempFilePath) Then Err.Raise vbObjectError + 7702, "SharedFile_WriteUtf8NoBom", "The staged UTF-8 file was not created."
'    expectedSize = fso.GetFile(tempFilePath).Size
'    If fso.fileExists(filePath) Then
'        candidateNumber = 0
'        backupFilePath = filePath & ".sharedu8.backup"
'        Do While fso.fileExists(backupFilePath)
'            candidateNumber = candidateNumber + 1
'            backupFilePath = filePath & ".sharedu8.backup." & CStr(candidateNumber)
'        Loop
'        fso.MoveFile filePath, backupFilePath
'        originalMoved = True
'        If fso.fileExists(filePath) Or Not fso.fileExists(backupFilePath) Then Err.Raise vbObjectError + 7703, "SharedFile_WriteUtf8NoBom", "The existing target file could not be staged for replacement."
'    End If
'    fso.MoveFile tempFilePath, filePath
'    replacementMoved = True
'    If Not fso.fileExists(filePath) Then Err.Raise vbObjectError + 7704, "SharedFile_WriteUtf8NoBom", "The replacement file does not exist after the move."
'    If fso.fileExists(tempFilePath) Then Err.Raise vbObjectError + 7705, "SharedFile_WriteUtf8NoBom", "The staged file still exists after replacement."
'    If fso.GetFile(filePath).Size <> expectedSize Then Err.Raise vbObjectError + 7706, "SharedFile_WriteUtf8NoBom", "The replacement file size does not match the staged file size."
'    replacementVerified = True
'    If originalMoved Then
'        On Error Resume Next
'        Err.Clear
'        fso.DeleteFile backupFilePath, False
'        cleanupErrorNumber = Err.Number
'        cleanupErrorSource = Err.Source
'        cleanupErrorDescription = Err.Description
'        Err.Clear
'        backupExists = fso.fileExists(backupFilePath)
'        inspectionErrorNumber = Err.Number
'        inspectionErrorSource = Err.Source
'        inspectionErrorDescription = Err.Description
'        On Error GoTo ErrHandler
'        If cleanupErrorNumber <> 0 Or inspectionErrorNumber <> 0 Or backupExists Then GoTo CommittedCleanupFailure
'    End If
'    Exit Sub
'ErrHandler:
'    savedErrorNumber = Err.Number
'    savedErrorSource = Err.Source
'    savedErrorDescription = Err.Description
'    On Error Resume Next
'    If Not binaryStream Is Nothing Then binaryStream.Close
'    Set binaryStream = Nothing
'    If Not textStream Is Nothing Then textStream.Close
'    Set textStream = Nothing
'    If replacementVerified Then
'        If cleanupErrorNumber = 0 Then
'            cleanupErrorNumber = savedErrorNumber
'            cleanupErrorSource = savedErrorSource
'            cleanupErrorDescription = savedErrorDescription
'        End If
'        GoTo CommittedCleanupFailure
'    End If
'    If Not fso Is Nothing Then
'        If originalMoved Then
'            If fso.fileExists(filePath) Then fso.DeleteFile filePath, False
'            If Not fso.fileExists(filePath) And fso.fileExists(backupFilePath) Then fso.MoveFile backupFilePath, filePath
'        ElseIf replacementMoved Then
'            If fso.fileExists(filePath) Then fso.DeleteFile filePath, False
'        End If
'        If Len(tempFilePath) > 0 Then
'            If fso.fileExists(tempFilePath) Then fso.DeleteFile tempFilePath, False
'        End If
'        finalExists = fso.fileExists(filePath)
'        If Len(tempFilePath) > 0 Then tempExists = fso.fileExists(tempFilePath)
'        If Len(backupFilePath) > 0 Then backupExists = fso.fileExists(backupFilePath)
'    End If
'    cleanupErrorNumber = Err.Number
'    cleanupErrorDescription = Err.Description
'    On Error GoTo 0
'    Err.Raise savedErrorNumber, "SharedFile_WriteUtf8NoBom", "Could not atomically write UTF-8 without BOM to '" & filePath & "'." & vbCrLf & "Original source: " & savedErrorSource & vbCrLf & "Original description: " & savedErrorDescription & vbCrLf & "Final exists: " & CStr(finalExists) & vbCrLf & "Staged file exists: " & CStr(tempExists) & vbCrLf & "Prior-file staging copy exists: " & CStr(backupExists) & IIf(cleanupErrorNumber <> 0, vbCrLf & "Cleanup error " & CStr(cleanupErrorNumber) & ": " & cleanupErrorDescription, vbNullString)
'CommittedCleanupFailure:
'    On Error Resume Next
'    Err.Clear
'    finalExists = fso.fileExists(filePath)
'    targetVerificationErrorNumber = Err.Number
'    targetVerificationErrorSource = Err.Source
'    targetVerificationErrorDescription = Err.Description
'    If targetVerificationErrorNumber = 0 And finalExists Then
'        Err.Clear
'        finalSizeMatches = (fso.GetFile(filePath).Size = expectedSize)
'        targetSizeVerificationErrorNumber = Err.Number
'        targetSizeVerificationErrorSource = Err.Source
'        targetSizeVerificationErrorDescription = Err.Description
'    End If
'    Err.Clear
'    If Len(tempFilePath) > 0 Then tempExists = fso.fileExists(tempFilePath)
'    If Err.Number <> 0 Then cleanupFailureText = SharedText_AppendNonBlankLine(cleanupFailureText, "Staged-file verification error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'    Err.Clear
'    If Len(backupFilePath) > 0 Then backupExists = fso.fileExists(backupFilePath)
'    If Err.Number <> 0 Then cleanupFailureText = SharedText_AppendNonBlankLine(cleanupFailureText, "Backup verification error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'    On Error GoTo 0
'    If cleanupErrorNumber <> 0 Then cleanupFailureText = SharedText_AppendNonBlankLine(cleanupFailureText, "Backup cleanup error " & CStr(cleanupErrorNumber) & " from " & cleanupErrorSource & ": " & cleanupErrorDescription)
'    If inspectionErrorNumber <> 0 Then cleanupFailureText = SharedText_AppendNonBlankLine(cleanupFailureText, "Initial backup verification error " & CStr(inspectionErrorNumber) & " from " & inspectionErrorSource & ": " & inspectionErrorDescription)
'    If targetVerificationErrorNumber <> 0 Then cleanupFailureText = SharedText_AppendNonBlankLine(cleanupFailureText, "Target existence verification error " & CStr(targetVerificationErrorNumber) & " from " & targetVerificationErrorSource & ": " & targetVerificationErrorDescription)
'    If targetSizeVerificationErrorNumber <> 0 Then cleanupFailureText = SharedText_AppendNonBlankLine(cleanupFailureText, "Target size verification error " & CStr(targetSizeVerificationErrorNumber) & " from " & targetSizeVerificationErrorSource & ": " & targetSizeVerificationErrorDescription)
'    cleanupFailureText = SharedText_AppendNonBlankLine(cleanupFailureText, "Prior-file staging-copy path: " & backupFilePath)
'    cleanupFailureText = SharedText_AppendNonBlankLine(cleanupFailureText, "Prior-file staging copy exists: " & CStr(backupExists))
'    cleanupFailureText = SharedText_AppendNonBlankLine(cleanupFailureText, "Reverified target exists: " & CStr(finalExists))
'    cleanupFailureText = SharedText_AppendNonBlankLine(cleanupFailureText, "Reverified target size matches staged file: " & CStr(finalSizeMatches))
'    cleanupFailureText = SharedText_AppendNonBlankLine(cleanupFailureText, "Staged file exists: " & CStr(tempExists))
'    If targetVerificationErrorNumber <> 0 Or targetSizeVerificationErrorNumber <> 0 Or Not finalExists Or Not finalSizeMatches Then Err.Raise vbObjectError + 7720, "SharedFile_WriteUtf8NoBom", "The UTF-8 replacement had committed, but the target could not be reverified after prior-file staging-copy cleanup failed. The material output state is uncertain; do not retry or roll back automatically." & vbCrLf & "Target: " & filePath & vbCrLf & cleanupFailureText
'    On Error Resume Next
'    Debug.Print "SharedFile_WriteUtf8NoBom warning: replacement committed and reverified; residual prior-file cleanup was preserved without failing the write." & vbCrLf & "Target: " & filePath & vbCrLf & cleanupFailureText
'    On Error GoTo 0
'    Exit Sub
'End Sub
'
'Public Function SharedValue_CoerceBoolean(ByVal v As Variant) As Boolean
'    'Coerces recognized Boolean representations and raises for unrecognized nonblank values.
'    'Consolidated old names: CoerceImportBoolean, IsTrueish.
'    Dim s As String
'    If IsNull(v) Then Exit Function
'    If VarType(v) = vbBoolean Then
'        SharedValue_CoerceBoolean = CBool(v)
'        Exit Function
'    End If
'    If IsNumeric(v) Then
'        SharedValue_CoerceBoolean = (CDbl(v) <> 0)
'        Exit Function
'    End If
'    s = LCase$(Trim$(CStr(v)))
'    If Len(s) = 0 Then Exit Function
'    Select Case s
'        Case "true", "yes", "y", "1", "-1", "on"
'            SharedValue_CoerceBoolean = True
'        Case "false", "no", "n", "0", "off"
'            SharedValue_CoerceBoolean = False
'        Case Else
'            Err.Raise vbObjectError + 7740, "SharedValue_CoerceBoolean", "Unrecognized Boolean value: " & CStr(v)
'    End Select
'End Function
'
'Public Function SharedValue_CoerceForDaoFieldType(ByVal v As Variant, ByVal targetType As Long) As Variant
'    'Coerces an imported value for assignment to a DAO field type.
'    'Consolidated old names: CoerceImportValue.
'    If IsNull(v) Then
'        SharedValue_CoerceForDaoFieldType = Null
'    ElseIf VarType(v) = vbString And Len(v) = 0 Then
'        SharedValue_CoerceForDaoFieldType = Null
'    ElseIf targetType = dbBoolean Then
'        SharedValue_CoerceForDaoFieldType = SharedValue_CoerceBoolean(v)
'    Else
'        SharedValue_CoerceForDaoFieldType = v
'    End If
'End Function
'
'Public Function SharedText_DelimitedNameListContains(ByVal nameValue As String, ByVal delimitedNames As String, Optional ByVal delimiter As String = ";") As Boolean
'    'Returns True when a delimited list contains a name using case-insensitive comparison.
'    'Consolidated old names: NameIsExcluded.
'    Dim v As Variant
'    If Len(delimiter) = 0 Then delimiter = ";"
'    For Each v In Split(delimitedNames, delimiter)
'        If StrComp(nameValue, Trim$(CStr(v)), vbTextCompare) = 0 Then
'            SharedText_DelimitedNameListContains = True
'            Exit Function
'        End If
'    Next v
'End Function
'
'Public Sub SharedDao_CopyMatchingRecordsetFields(ByVal src As DAO.Recordset, ByVal dst As DAO.Recordset, Optional ByVal excludedNames As String = vbNullString, Optional ByVal skipAutoNumberFields As Boolean = True)
'    'Copies matching fields from one recordset row to another, excluding semicolon-listed names and optionally AutoNumber destination fields.
'    'Consolidated old names: CopyImportFields.
'    Dim fld As DAO.Field
'    Dim srcName As String
'    Dim dstName As String
'    Dim v As Variant
'    For Each fld In dst.fields
'        If (skipAutoNumberFields = False Or SharedDao_IsAutoNumberField(fld) = False) Then
'            If Not SharedText_DelimitedNameListContains(fld.Name, excludedNames, ";") Then
'                If SharedDao_RecordsetFieldExists(src, fld.Name) Then
'                    srcName = fld.Name
'                    dstName = fld.Name
'                    v = src.fields(srcName).value
'                    On Error GoTo FieldError
'                    dst.fields(dstName).value = SharedValue_NormalizeAssignment(v)
'                    On Error GoTo 0
'                End If
'            End If
'        End If
'    Next fld
'    Exit Sub
'FieldError:
'    Dim savedErrNumber As Long
'    Dim savedErrDescription As String
'    Dim copyErrorText As String
'    savedErrNumber = Err.Number
'    savedErrDescription = Err.Description
'    copyErrorText = "Could not copy source field '" & srcName & "' to destination field '" & dstName & "'." & vbCrLf
'    copyErrorText = copyErrorText & "Value: " & Nz(v, "<Null>") & vbCrLf
'    copyErrorText = copyErrorText & "Destination DAO type: " & CStr(fld.Type) & vbCrLf
'    copyErrorText = copyErrorText & "Destination size: " & CStr(fld.Size) & vbCrLf
'    copyErrorText = copyErrorText & savedErrDescription
'    Err.Raise savedErrNumber, "SharedDao_CopyMatchingRecordsetFields", copyErrorText
'End Sub
'
'Public Sub SharedDao_SetMappedRecordsetField(ByVal src As DAO.Recordset, ByVal dst As DAO.Recordset, ByVal srcName As String, ByVal dstName As String)
'    'Copies one source recordset field to one destination recordset field, normalizing blank text to Null.
'    'Consolidated old names: SetMappedField.
'    Dim v As Variant
'    If SharedDao_RecordsetFieldExists(src, srcName) And SharedDao_RecordsetFieldExists(dst, dstName) Then
'        v = src.fields(srcName).value
'        On Error GoTo FieldError
'        dst.fields(dstName).value = SharedValue_NormalizeAssignment(v)
'        On Error GoTo 0
'    End If
'    Exit Sub
'FieldError:
'    Dim savedErrNumber As Long
'    Dim savedErrDescription As String
'    Dim mappedErrorText As String
'    savedErrNumber = Err.Number
'    savedErrDescription = Err.Description
'    mappedErrorText = "Could not copy source field '" & srcName & "' to destination field '" & dstName & "'." & vbCrLf
'    mappedErrorText = mappedErrorText & "Value: " & Nz(v, "<Null>") & vbCrLf
'    mappedErrorText = mappedErrorText & "Destination DAO type: " & CStr(dst.fields(dstName).Type) & vbCrLf
'    mappedErrorText = mappedErrorText & "Destination size: " & CStr(dst.fields(dstName).Size) & vbCrLf
'    mappedErrorText = mappedErrorText & savedErrDescription
'    Err.Raise savedErrNumber, "SharedDao_SetMappedRecordsetField", mappedErrorText
'End Sub
'
'Public Sub SharedCollection_AddValueToDictionaryCollection(ByVal dictionaryObject As Object, ByVal collectionKey As String, ByVal value As Variant)
'    'Adds a value to a Collection stored under a key in a Scripting.Dictionary-like object.
'    'Consolidated old names: AddChildNode.
'    Dim values As Collection
'    If Not dictionaryObject.Exists(collectionKey) Then
'        Set values = New Collection
'        dictionaryObject.Add collectionKey, values
'    Else
'        Set values = dictionaryObject(collectionKey)
'    End If
'    values.Add value
'End Sub
'
'Public Function SharedValue_NormalizeAssignment(ByVal v As Variant) As Variant
'    'Returns Null for Null or whitespace-only text; otherwise returns the original value.
'    'Consolidated old names: NormalizeImportAssignmentValue, NullIfBlank.
'    Dim s As String
'    If IsNull(v) Then
'        SharedValue_NormalizeAssignment = Null
'    ElseIf VarType(v) = vbString Then
'        s = CStr(v)
'        If Len(Trim$(s)) = 0 Then
'            SharedValue_NormalizeAssignment = Null
'        Else
'            SharedValue_NormalizeAssignment = v
'        End If
'    Else
'        SharedValue_NormalizeAssignment = v
'    End If
'End Function
'
'Public Function SharedValue_IsBlankOrMissing(ByVal v As Variant) As Boolean
'    'Returns True when a value is Null or a string that trims to blank.
'    'Consolidated old names: IsImportMissing.
'    If IsNull(v) Then
'        SharedValue_IsBlankOrMissing = True
'    ElseIf VarType(v) = vbString Then
'        SharedValue_IsBlankOrMissing = (Len(Trim$(CStr(v))) = 0)
'    End If
'End Function
'
'Public Function SharedPath_BaseFileName(ByVal filePath As String) As String
'    'Returns the file name portion of a path.
'    'Consolidated old names: local TV NSS IMPORT MODULE BaseFileName.
'    Dim pBackslash As Long
'    Dim pSlash As Long
'    Dim p As Long
'    pBackslash = InStrRev(filePath, "\")
'    pSlash = InStrRev(filePath, "/")
'    If pBackslash > pSlash Then p = pBackslash Else p = pSlash
'    If p > 0 Then
'        SharedPath_BaseFileName = Mid$(filePath, p + 1)
'    Else
'        SharedPath_BaseFileName = filePath
'    End If
'End Function
'
'Public Function SharedFile_ReadAllText(ByVal filePath As String) As String
'    'Reads an entire text file into a String.
'    'Consolidated old names: local TV NSS IMPORT MODULE ReadAllText.
'    Dim fso As Object
'    Dim ts As Object
'    If Dir$(filePath, vbNormal) = vbNullString Then Exit Function
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    Set ts = fso.OpenTextFile(filePath, 1, False)
'    SharedFile_ReadAllText = ts.ReadAll
'    ts.Close
'End Function
'
'Public Function SharedText_RemoveBom(ByVal value As String) As String
'    'Removes UTF BOM characters commonly found at the start of text files.
'    'Consolidated old names: local TV NSS IMPORT MODULE RemoveBom.
'    value = Replace(value, Chr$(239) & Chr$(187) & Chr$(191), vbNullString)
'    value = Replace(value, ChrW$(&HFEFF), vbNullString)
'    SharedText_RemoveBom = value
'End Function
'
'Public Sub SharedPath_EnsureFolder(ByVal folderPath As String)
'    'Creates every missing folder in the supplied path and verifies that the final folder exists.
'    'Use for backup, export, import, and setup paths whose parent hierarchy may not exist yet.
'    On Error GoTo ErrHandler
'    Dim fso As Object
'    Dim requestedPath As String
'    requestedPath = Trim$(folderPath)
'    If Len(requestedPath) = 0 Then Exit Sub
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    requestedPath = fso.GetAbsolutePathName(requestedPath)
'    SharedPath_EnsureFolderCore fso, requestedPath, folderPath
'    Set fso = Nothing
'    Exit Sub
'ErrHandler:
'    Dim savedErrorNumber As Long
'    Dim savedErrorSource As String
'    Dim savedErrorDescription As String
'    savedErrorNumber = Err.Number
'    savedErrorSource = Err.Source
'    savedErrorDescription = Err.Description
'    Set fso = Nothing
'    Err.Raise savedErrorNumber, "SharedPath_EnsureFolder", "Could not ensure folder path." & vbCrLf & "Requested path: " & folderPath & vbCrLf & "Resolved path: " & IIf(Len(requestedPath) = 0, "(not resolved)", requestedPath) & vbCrLf & "Original source: " & savedErrorSource & vbCrLf & "Description: " & savedErrorDescription
'End Sub
'
'Private Sub SharedPath_EnsureFolderCore(ByVal fso As Object, ByVal requestedPath As String, ByVal originalPath As String)
'    On Error GoTo ErrHandler
'    Dim parentPath As String
'    If fso Is Nothing Then Err.Raise vbObjectError + 2100, "SharedPath_EnsureFolderCore", "No FileSystemObject was supplied."
'    requestedPath = Trim$(requestedPath)
'    If Len(requestedPath) = 0 Then Err.Raise vbObjectError + 2101, "SharedPath_EnsureFolderCore", "The resolved folder path was blank."
'    If fso.FolderExists(requestedPath) Then Exit Sub
'    parentPath = fso.GetParentFolderName(requestedPath)
'    If Len(parentPath) = 0 Or StrComp(parentPath, requestedPath, vbTextCompare) = 0 Then Err.Raise vbObjectError + 2102, "SharedPath_EnsureFolderCore", "The folder hierarchy could not be resolved."
'    If Not fso.FolderExists(parentPath) Then SharedPath_EnsureFolderCore fso, parentPath, originalPath
'    fso.CreateFolder requestedPath
'    If Not fso.FolderExists(requestedPath) Then Err.Raise vbObjectError + 2103, "SharedPath_EnsureFolderCore", "The folder was not present after the create request."
'    Exit Sub
'ErrHandler:
'    Dim savedErrorNumber As Long
'    Dim savedErrorSource As String
'    Dim savedErrorDescription As String
'    savedErrorNumber = Err.Number
'    savedErrorSource = Err.Source
'    savedErrorDescription = Err.Description
'    Err.Raise savedErrorNumber, "SharedPath_EnsureFolderCore", "Folder-tree creation failed." & vbCrLf & "Original requested path: " & originalPath & vbCrLf & "Current path: " & IIf(Len(requestedPath) = 0, "(blank)", requestedPath) & vbCrLf & "Parent path: " & IIf(Len(parentPath) = 0, "(not resolved)", parentPath) & vbCrLf & "Original source: " & savedErrorSource & vbCrLf & "Description: " & savedErrorDescription
'End Sub
'
'Public Sub SharedFile_DeleteIfExists(ByVal filePath As String)
'    'Deletes one exact file when it exists and verifies that deletion completed.
'    'Consolidated old names: local TV NSS IMPORT MODULE DeleteFileIfExists.
'    On Error GoTo ErrHandler
'    Dim fso As Object
'    Dim savedErrorNumber As Long
'    Dim savedErrorDescription As String
'    If Len(Trim$(filePath)) = 0 Then Exit Sub
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    If Not fso.fileExists(filePath) Then Exit Sub
'    fso.DeleteFile filePath, False
'    If fso.fileExists(filePath) Then Err.Raise vbObjectError + 3202, "SharedFile_DeleteIfExists", "The file still exists after the delete request."
'    Exit Sub
'ErrHandler:
'    savedErrorNumber = Err.Number
'    savedErrorDescription = Err.Description
'    Err.Raise savedErrorNumber, "SharedFile_DeleteIfExists", "Could not delete and verify file '" & filePath & "'. Original error " & savedErrorNumber & ": " & savedErrorDescription
'End Sub
'
'Public Function SharedText_PowerShellSingleQuote(ByVal value As String) As String
'    'Wraps a value in PowerShell single quotes and escapes embedded single quotes.
'    'Consolidated old names: PSQ.
'    SharedText_PowerShellSingleQuote = "'" & Replace(value, "'", "''") & "'"
'End Function
'
'Public Sub SharedProcess_TerminateByID(ByVal processID As Long)
'    'Terminates one Windows process by process ID through WMI and verifies its disappearance through a bounded retry loop.
'    'Consolidated old names: local TV NSS IMPORT MODULE TerminateProcessByID.
'    On Error GoTo ErrHandler
'    Dim wmi As Object
'    Dim procs As Object
'    Dim proc As Object
'    Dim terminationResult As Long
'    Dim processMatched As Boolean
'    Dim processStillExists As Boolean
'    Dim verificationDeadline As Date
'    Dim verificationAttempt As Long
'    If processID <= 0 Then Exit Sub
'    Set wmi = GetObject("winmgmts:\.\root\cimv2")
'    Set procs = wmi.ExecQuery("SELECT * FROM Win32_Process WHERE ProcessId=" & CStr(processID))
'    For Each proc In procs
'        processMatched = True
'        terminationResult = CLng(proc.Terminate())
'        If terminationResult <> 0 Then Err.Raise vbObjectError + 7708, "SharedProcess_TerminateByID", "WMI returned a nonzero process-termination result: " & CStr(terminationResult)
'    Next proc
'    If Not processMatched Then Err.Raise vbObjectError + 7709, "SharedProcess_TerminateByID", "No process matched the requested process ID."
'    verificationDeadline = DateAdd("s", 3, Now)
'    Do
'        verificationAttempt = verificationAttempt + 1
'        processStillExists = False
'        Set procs = wmi.ExecQuery("SELECT * FROM Win32_Process WHERE ProcessId=" & CStr(processID))
'        For Each proc In procs
'            processStillExists = True
'            Exit For
'        Next proc
'        If Not processStillExists Then Exit Do
'        If Now >= verificationDeadline Then Exit Do
'        DoEvents
'    Loop
'    If processStillExists Then Err.Raise vbObjectError + 7710, "SharedProcess_TerminateByID", "The process still exists after WMI reported successful termination and bounded verification completed."
'    Exit Sub
'ErrHandler:
'    Dim savedErrorNumber As Long
'    Dim savedErrorSource As String
'    Dim savedErrorDescription As String
'    savedErrorNumber = Err.Number
'    savedErrorSource = Err.Source
'    savedErrorDescription = Err.Description
'    Err.Raise savedErrorNumber, "SharedProcess_TerminateByID", "Could not terminate and verify the Windows process." & vbCrLf & "Process ID: " & CStr(processID) & vbCrLf & "WMI result: " & CStr(terminationResult) & vbCrLf & "Verification attempts: " & CStr(verificationAttempt) & vbCrLf & "Original source: " & savedErrorSource & vbCrLf & "Description: " & savedErrorDescription
'End Sub
'
'Public Sub SharedFile_WriteText(ByVal filePath As String, ByVal textValue As String)
'    'Writes text through a same-folder temporary file, verifies the replacement, and restores the prior file on pre-commit failure.
'    'Consolidated old names: local TV NSS IMPORT MODULE WriteTextFile.
'    On Error GoTo ErrHandler
'    Dim fso As Object
'    Dim ts As Object
'    Dim tempFilePath As String
'    Dim backupFilePath As String
'    Dim candidateNumber As Long
'    Dim originalMoved As Boolean
'    Dim replacementMoved As Boolean
'    Dim replacementVerified As Boolean
'    Dim expectedSize As Variant
'    Dim savedErrorNumber As Long
'    Dim savedErrorSource As String
'    Dim savedErrorDescription As String
'    Dim cleanupErrorNumber As Long
'    Dim cleanupErrorSource As String
'    Dim cleanupErrorDescription As String
'    Dim cleanupFailureText As String
'    Dim inspectionErrorNumber As Long
'    Dim inspectionErrorSource As String
'    Dim inspectionErrorDescription As String
'    Dim targetVerificationErrorNumber As Long
'    Dim targetVerificationErrorSource As String
'    Dim targetVerificationErrorDescription As String
'    Dim targetSizeVerificationErrorNumber As Long
'    Dim targetSizeVerificationErrorSource As String
'    Dim targetSizeVerificationErrorDescription As String
'    Dim finalExists As Boolean
'    Dim finalSizeMatches As Boolean
'    Dim tempExists As Boolean
'    Dim backupExists As Boolean
'    If Len(Trim$(filePath)) = 0 Then Err.Raise vbObjectError + 3203, "SharedFile_WriteText", "The target file path is blank."
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    tempFilePath = filePath & ".sharedwrite.tmp"
'    Do While fso.fileExists(tempFilePath)
'        candidateNumber = candidateNumber + 1
'        tempFilePath = filePath & ".sharedwrite." & CStr(candidateNumber) & ".tmp"
'    Loop
'    Set ts = fso.CreateTextFile(tempFilePath, False, False)
'    ts.Write textValue
'    ts.Close
'    Set ts = Nothing
'    If Not fso.fileExists(tempFilePath) Then Err.Raise vbObjectError + 3204, "SharedFile_WriteText", "The staged text file was not created."
'    expectedSize = fso.GetFile(tempFilePath).Size
'    If fso.fileExists(filePath) Then
'        candidateNumber = 0
'        backupFilePath = filePath & ".sharedwrite.backup"
'        Do While fso.fileExists(backupFilePath)
'            candidateNumber = candidateNumber + 1
'            backupFilePath = filePath & ".sharedwrite.backup." & CStr(candidateNumber)
'        Loop
'        fso.MoveFile filePath, backupFilePath
'        originalMoved = True
'        If fso.fileExists(filePath) Or Not fso.fileExists(backupFilePath) Then Err.Raise vbObjectError + 3205, "SharedFile_WriteText", "The existing target file could not be staged for replacement."
'    End If
'    fso.MoveFile tempFilePath, filePath
'    replacementMoved = True
'    If Not fso.fileExists(filePath) Then Err.Raise vbObjectError + 3206, "SharedFile_WriteText", "The replacement file does not exist after the move."
'    If fso.fileExists(tempFilePath) Then Err.Raise vbObjectError + 3207, "SharedFile_WriteText", "The staged file still exists after replacement."
'    If fso.GetFile(filePath).Size <> expectedSize Then Err.Raise vbObjectError + 3208, "SharedFile_WriteText", "The replacement file size does not match the staged file size."
'    replacementVerified = True
'    If originalMoved Then
'        On Error Resume Next
'        Err.Clear
'        fso.DeleteFile backupFilePath, False
'        cleanupErrorNumber = Err.Number
'        cleanupErrorSource = Err.Source
'        cleanupErrorDescription = Err.Description
'        Err.Clear
'        backupExists = fso.fileExists(backupFilePath)
'        inspectionErrorNumber = Err.Number
'        inspectionErrorSource = Err.Source
'        inspectionErrorDescription = Err.Description
'        On Error GoTo ErrHandler
'        If cleanupErrorNumber <> 0 Or inspectionErrorNumber <> 0 Or backupExists Then GoTo CommittedCleanupFailure
'    End If
'    Exit Sub
'ErrHandler:
'    savedErrorNumber = Err.Number
'    savedErrorSource = Err.Source
'    savedErrorDescription = Err.Description
'    On Error Resume Next
'    If Not ts Is Nothing Then ts.Close
'    Set ts = Nothing
'    If replacementVerified Then
'        If cleanupErrorNumber = 0 Then
'            cleanupErrorNumber = savedErrorNumber
'            cleanupErrorSource = savedErrorSource
'            cleanupErrorDescription = savedErrorDescription
'        End If
'        GoTo CommittedCleanupFailure
'    End If
'    If Not fso Is Nothing Then
'        If originalMoved Then
'            If fso.fileExists(filePath) Then fso.DeleteFile filePath, False
'            If Not fso.fileExists(filePath) And fso.fileExists(backupFilePath) Then fso.MoveFile backupFilePath, filePath
'        ElseIf replacementMoved Then
'            If fso.fileExists(filePath) Then fso.DeleteFile filePath, False
'        End If
'        If Len(tempFilePath) > 0 Then
'            If fso.fileExists(tempFilePath) Then fso.DeleteFile tempFilePath, False
'        End If
'        finalExists = fso.fileExists(filePath)
'        If Len(tempFilePath) > 0 Then tempExists = fso.fileExists(tempFilePath)
'        If Len(backupFilePath) > 0 Then backupExists = fso.fileExists(backupFilePath)
'    End If
'    cleanupErrorNumber = Err.Number
'    cleanupErrorDescription = Err.Description
'    On Error GoTo 0
'    Err.Raise savedErrorNumber, "SharedFile_WriteText", "Could not atomically write text to '" & filePath & "'. Original error " & savedErrorNumber & " from " & savedErrorSource & ": " & savedErrorDescription & ". Final exists=" & CStr(finalExists) & "; staged file exists=" & CStr(tempExists) & "; prior-file staging copy exists=" & CStr(backupExists) & IIf(cleanupErrorNumber <> 0, ". Cleanup error " & cleanupErrorNumber & ": " & cleanupErrorDescription, vbNullString)
'CommittedCleanupFailure:
'    On Error Resume Next
'    Err.Clear
'    finalExists = fso.fileExists(filePath)
'    targetVerificationErrorNumber = Err.Number
'    targetVerificationErrorSource = Err.Source
'    targetVerificationErrorDescription = Err.Description
'    If targetVerificationErrorNumber = 0 And finalExists Then
'        Err.Clear
'        finalSizeMatches = (fso.GetFile(filePath).Size = expectedSize)
'        targetSizeVerificationErrorNumber = Err.Number
'        targetSizeVerificationErrorSource = Err.Source
'        targetSizeVerificationErrorDescription = Err.Description
'    End If
'    Err.Clear
'    If Len(tempFilePath) > 0 Then tempExists = fso.fileExists(tempFilePath)
'    If Err.Number <> 0 Then cleanupFailureText = SharedText_AppendNonBlankLine(cleanupFailureText, "Staged-file verification error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'    Err.Clear
'    If Len(backupFilePath) > 0 Then backupExists = fso.fileExists(backupFilePath)
'    If Err.Number <> 0 Then cleanupFailureText = SharedText_AppendNonBlankLine(cleanupFailureText, "Backup verification error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'    On Error GoTo 0
'    If cleanupErrorNumber <> 0 Then cleanupFailureText = SharedText_AppendNonBlankLine(cleanupFailureText, "Backup cleanup error " & CStr(cleanupErrorNumber) & " from " & cleanupErrorSource & ": " & cleanupErrorDescription)
'    If inspectionErrorNumber <> 0 Then cleanupFailureText = SharedText_AppendNonBlankLine(cleanupFailureText, "Initial backup verification error " & CStr(inspectionErrorNumber) & " from " & inspectionErrorSource & ": " & inspectionErrorDescription)
'    If targetVerificationErrorNumber <> 0 Then cleanupFailureText = SharedText_AppendNonBlankLine(cleanupFailureText, "Target existence verification error " & CStr(targetVerificationErrorNumber) & " from " & targetVerificationErrorSource & ": " & targetVerificationErrorDescription)
'    If targetSizeVerificationErrorNumber <> 0 Then cleanupFailureText = SharedText_AppendNonBlankLine(cleanupFailureText, "Target size verification error " & CStr(targetSizeVerificationErrorNumber) & " from " & targetSizeVerificationErrorSource & ": " & targetSizeVerificationErrorDescription)
'    cleanupFailureText = SharedText_AppendNonBlankLine(cleanupFailureText, "Prior-file staging-copy path: " & backupFilePath)
'    cleanupFailureText = SharedText_AppendNonBlankLine(cleanupFailureText, "Prior-file staging copy exists: " & CStr(backupExists))
'    cleanupFailureText = SharedText_AppendNonBlankLine(cleanupFailureText, "Reverified target exists: " & CStr(finalExists))
'    cleanupFailureText = SharedText_AppendNonBlankLine(cleanupFailureText, "Reverified target size matches staged file: " & CStr(finalSizeMatches))
'    cleanupFailureText = SharedText_AppendNonBlankLine(cleanupFailureText, "Staged file exists: " & CStr(tempExists))
'    If targetVerificationErrorNumber <> 0 Or targetSizeVerificationErrorNumber <> 0 Or Not finalExists Or Not finalSizeMatches Then Err.Raise vbObjectError + 3210, "SharedFile_WriteText", "The text replacement had committed, but the target could not be reverified after prior-file staging-copy cleanup failed. The material output state is uncertain; do not retry or roll back automatically." & vbCrLf & "Target: " & filePath & vbCrLf & cleanupFailureText
'    On Error Resume Next
'    Debug.Print "SharedFile_WriteText warning: replacement committed and reverified; residual prior-file cleanup was preserved without failing the write." & vbCrLf & "Target: " & filePath & vbCrLf & cleanupFailureText
'    On Error GoTo 0
'    Exit Sub
'End Sub
'
'Public Function SharedValue_ToDisplayText(ByVal value As Variant) As String
'    'Converts a Variant value to safe display/export text.
'    'Consolidated old names: PropertyValueToText.
'    If IsObject(value) Then
'        SharedValue_ToDisplayText = "[Object: " & typeName(value) & "]"
'    ElseIf IsNull(value) Then
'        SharedValue_ToDisplayText = "[Null]"
'    ElseIf IsEmpty(value) Then
'        SharedValue_ToDisplayText = "[Empty]"
'    ElseIf IsArray(value) Then
'        SharedValue_ToDisplayText = "[Array]"
'    Else
'        SharedValue_ToDisplayText = CStr(value)
'    End If
'End Function
'
'Public Function SharedValue_TypeToText(ByVal value As Variant) As String
'    'Returns a display/export type name for a Variant value.
'    'Consolidated old names: PropertyTypeToText.
'    If IsObject(value) Then
'        SharedValue_TypeToText = typeName(value)
'    ElseIf IsArray(value) Then
'        SharedValue_TypeToText = "Array"
'    Else
'        SharedValue_TypeToText = typeName(value)
'    End If
'End Function
'
'Public Function SharedValue_NzLong(ByVal value As Variant, Optional ByVal DefaultValue As Long = 0) As Long
'    'Intended use: Returns a Long value with a caller-supplied default when the source is Null or blank.
'    If IsNull(value) Then
'        SharedValue_NzLong = DefaultValue
'    ElseIf Len(Trim$(CStr(value))) = 0 Then
'        SharedValue_NzLong = DefaultValue
'    ElseIf IsNumeric(value) Then
'        SharedValue_NzLong = CLng(value)
'    Else
'        SharedValue_NzLong = DefaultValue
'    End If
'End Function
'
'Public Function SharedText_AppendNonBlankLine(ByVal existingText As String, ByVal lineText As String) As String
'    'Intended use: Appends a nonblank line to existing text while preserving clean line separation.
'    If Len(Trim$(lineText)) = 0 Then
'        SharedText_AppendNonBlankLine = existingText
'    ElseIf Len(Trim$(existingText)) = 0 Then
'        SharedText_AppendNonBlankLine = lineText
'    Else
'        SharedText_AppendNonBlankLine = existingText & vbCrLf & lineText
'    End If
'End Function
'
'Public Function SharedCollection_Join(ByVal values As Collection, Optional ByVal delimiter As String = ", ") As String
'    'Intended use: Joins Collection items into deterministic delimiter-separated text for reports or generated output.
'    Dim item As Variant
'    Dim resultText As String
'    If values Is Nothing Then Exit Function
'    For Each item In values
'        If Len(resultText) > 0 Then resultText = resultText & delimiter
'        resultText = resultText & CStr(item)
'    Next item
'    SharedCollection_Join = resultText
'End Function
'
'Public Function SharedName_SafeAccessName(ByVal value As String, Optional ByVal firstCharacterPrefix As String = "Db_", Optional ByVal fallbackName As String = vbNullString, Optional ByVal maxLength As Long = 0) As String
'    'Intended use: Returns a sanitized Access identifier with caller-controlled prefix, fallback, and length limits.
'    Dim i As Long
'    Dim ch As String
'    Dim resultText As String
'    value = Trim$(value)
'    For i = 1 To Len(value)
'        ch = Mid$(value, i, 1)
'        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Or ch = "_" Then
'            resultText = resultText & ch
'        Else
'            resultText = resultText & "_"
'        End If
'    Next i
'    Do While InStr(1, resultText, "__", vbBinaryCompare) > 0
'        resultText = Replace(resultText, "__", "_")
'    Loop
'    Do While Len(resultText) > 0 And Left$(resultText, 1) = "_"
'        resultText = Mid$(resultText, 2)
'    Loop
'    Do While Len(resultText) > 0 And Right$(resultText, 1) = "_"
'        resultText = Left$(resultText, Len(resultText) - 1)
'    Loop
'    If Len(resultText) = 0 And Len(fallbackName) > 0 Then resultText = fallbackName
'    If Len(resultText) > 0 Then
'        ch = Left$(resultText, 1)
'        If Not ((ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z")) Then resultText = firstCharacterPrefix & resultText
'    End If
'    If maxLength > 0 Then resultText = Left$(resultText, maxLength)
'    SharedName_SafeAccessName = resultText
'End Function
'
'Public Function SharedName_SafeAccessObjectName(ByVal value As String, Optional ByVal maxLength As Long = 0, Optional ByVal firstCharacterPrefix As String = "DB_", Optional ByVal fallbackName As String = vbNullString) As String
'    'Intended use: Returns a sanitized Access database-object name within the requested length limit.
'    SharedName_SafeAccessObjectName = SharedName_SafeAccessName(value, firstCharacterPrefix, fallbackName, maxLength)
'End Function
'
'Public Function SharedName_SafeAccessObjectNamePart(ByVal valueText As String, Optional ByVal fallbackName As String = "ADC", Optional ByVal firstCharacterPrefix As String = "Db_") As String
'    'Intended use: Returns a sanitized component suitable for composing a larger Access object name.
'    SharedName_SafeAccessObjectNamePart = SharedName_SafeAccessName(valueText, firstCharacterPrefix, fallbackName, 0)
'End Function
'
'Public Function SharedName_SafeFileName(ByVal value As String, Optional ByVal fallbackName As String = "AccessObject", Optional ByVal maxLength As Long = 120) As String
'    'Intended use: Returns a Windows-filesystem-safe filename component with fallback and maximum-length enforcement.
'    Dim resultText As String
'    resultText = SharedName_SanitizeFileNameCore(value)
'    If Len(resultText) = 0 Then resultText = SharedName_SanitizeFileNameCore(fallbackName)
'    If Len(resultText) = 0 Then resultText = "AccessObject"
'    If SharedName_WindowsFileBaseIsReserved(resultText) Then resultText = "_" & resultText
'    If maxLength > 0 Then resultText = Left$(resultText, maxLength)
'    resultText = SharedName_SanitizeFileNameCore(resultText)
'    If Len(resultText) = 0 Then resultText = "AccessObject"
'    If SharedName_WindowsFileBaseIsReserved(resultText) Then resultText = "_" & resultText
'    If maxLength > 0 Then resultText = Left$(resultText, maxLength)
'    resultText = SharedName_SanitizeFileNameCore(resultText)
'    If Len(resultText) = 0 Then resultText = "_"
'    SharedName_SafeFileName = resultText
'End Function
'
'Private Function SharedName_SanitizeFileNameCore(ByVal value As String) As String
'    'Removes Windows-invalid filename characters, control characters, repeated underscores, and trailing spaces or periods.
'    Dim invalidChars As Variant
'    Dim ch As Variant
'    Dim resultText As String
'    Dim characterCode As Long
'    resultText = Trim$(value)
'    invalidChars = Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")
'    For Each ch In invalidChars
'        resultText = Replace(resultText, CStr(ch), "_")
'    Next ch
'    For characterCode = 0 To 31
'        resultText = Replace(resultText, Chr$(characterCode), "_")
'    Next characterCode
'    Do While InStr(1, resultText, "__", vbBinaryCompare) > 0
'        resultText = Replace(resultText, "__", "_")
'    Loop
'    Do While Len(resultText) > 0 And (Right$(resultText, 1) = " " Or Right$(resultText, 1) = ".")
'        resultText = Left$(resultText, Len(resultText) - 1)
'    Loop
'    SharedName_SanitizeFileNameCore = resultText
'End Function
'
'Private Function SharedName_WindowsFileBaseIsReserved(ByVal fileNameText As String) As Boolean
'    'Returns whether the filename base is a reserved Windows device name, including when an extension is present.
'    Dim baseName As String
'    Dim dotPosition As Long
'    baseName = UCase$(Trim$(fileNameText))
'    dotPosition = InStr(1, baseName, ".", vbBinaryCompare)
'    If dotPosition > 0 Then baseName = Left$(baseName, dotPosition - 1)
'    Do While Len(baseName) > 0 And (Right$(baseName, 1) = " " Or Right$(baseName, 1) = ".")
'        baseName = Left$(baseName, Len(baseName) - 1)
'    Loop
'    Select Case baseName
'        Case "CON", "PRN", "AUX", "NUL", "CLOCK$", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
'            SharedName_WindowsFileBaseIsReserved = True
'    End Select
'End Function
'
'' ================================================================================================================================================================================
'' CSV
'' ================================================================================================================================================================================
'Public Function SharedCsv_TextToBoolean(ByVal csvValue As String) As Boolean
'    'Converts strict CSV Boolean text to Boolean and raises an error for invalid values.
'    'Consolidated old names: local TV NSS IMPORT MODULE CsvTextToBoolean.
'    Dim normalized As String
'    normalized = LCase$(Trim$(csvValue))
'    Select Case normalized
'        Case "true", "yes", "y", "1", "-1", "on"
'            SharedCsv_TextToBoolean = True
'        Case "false", "no", "n", "0", "off"
'            SharedCsv_TextToBoolean = False
'        Case Else
'            Err.Raise vbObjectError + 3201, "SharedCsv_TextToBoolean", "Invalid Boolean CSV value for Yes/No field: " & csvValue
'    End Select
'End Function
'
'Public Function SharedCsv_ReadRecord(ByVal ts As Object, ByRef physicalLineNum As Long, ByVal csvPath As String) As String
'    'Reads one logical CSV record from a text stream, including embedded line breaks inside quoted fields.
'    'Consolidated old names: local TV NSS IMPORT MODULE ReadCsvRecord.
'    Dim txt As String
'    If ts.AtEndOfStream Then Exit Function
'    txt = ts.ReadLine
'    physicalLineNum = physicalLineNum + 1
'    Do While Not SharedCsv_RecordIsComplete(txt)
'        If ts.AtEndOfStream Then Err.Raise vbObjectError + 3101, "SharedCsv_ReadRecord", "Malformed CSV record. Quoted field was not closed." & vbCrLf & "CSV File: " & csvPath & vbCrLf & "Physical Row: " & physicalLineNum
'        txt = txt & vbCrLf & ts.ReadLine
'        physicalLineNum = physicalLineNum + 1
'    Loop
'    SharedCsv_ReadRecord = txt
'End Function
'
'Public Function SharedCsv_RecordIsComplete(ByVal txt As String) As Boolean
'    'Returns True when a CSV record has balanced non-escaped quotes.
'    'Consolidated old names: local TV NSS IMPORT MODULE CsvRecordIsComplete.
'    Dim inQuotes As Boolean
'    Dim i As Long
'    Dim ch As String
'    For i = 1 To Len(txt)
'        ch = Mid$(txt, i, 1)
'        If ch = Chr$(34) Then
'            If inQuotes And i < Len(txt) And Mid$(txt, i + 1, 1) = Chr$(34) Then
'                i = i + 1
'            Else
'                inQuotes = Not inQuotes
'            End If
'        End If
'    Next i
'    SharedCsv_RecordIsComplete = Not inQuotes
'End Function
'
'Public Function SharedCsv_SplitLine(ByVal txt As String) As String()
'    'Splits one CSV record and rejects quotes outside a valid quoted-field position.
'    'Consolidated old names: local TV NSS IMPORT MODULE SplitCsvLine.
'    Dim quotedFields() As Boolean
'    SharedCsv_SplitLine = SharedCsv_SplitLineDetailed(txt, quotedFields)
'End Function
'
'Private Function SharedCsv_SplitLineDetailed(ByVal txt As String, ByRef quotedFields() As Boolean) As String()
'    'Splits one CSV record while retaining whether each field was explicitly quoted.
'    Dim results() As String
'    Dim currentValue As String
'    Dim inQuotes As Boolean
'    Dim fieldStarted As Boolean
'    Dim quoteClosed As Boolean
'    Dim currentFieldQuoted As Boolean
'    Dim i As Long
'    Dim ch As String
'    ReDim results(0)
'    ReDim quotedFields(0)
'    For i = 1 To Len(txt)
'        ch = Mid$(txt, i, 1)
'        If inQuotes Then
'            If ch = Chr$(34) Then
'                If i < Len(txt) And Mid$(txt, i + 1, 1) = Chr$(34) Then
'                    currentValue = currentValue & Chr$(34)
'                    i = i + 1
'                Else
'                    inQuotes = False
'                    quoteClosed = True
'                End If
'            Else
'                currentValue = currentValue & ch
'            End If
'        Else
'            Select Case ch
'                Case Chr$(34)
'                    If fieldStarted Or quoteClosed Then Err.Raise vbObjectError + 7741, "SharedCsv_SplitLineDetailed", "A quote appeared outside the beginning of a quoted CSV field at character " & CStr(i) & "."
'                    inQuotes = True
'                    fieldStarted = True
'                    currentFieldQuoted = True
'                Case ","
'                    results(UBound(results)) = currentValue
'                    quotedFields(UBound(quotedFields)) = currentFieldQuoted
'                    ReDim Preserve results(UBound(results) + 1)
'                    ReDim Preserve quotedFields(UBound(quotedFields) + 1)
'                    currentValue = vbNullString
'                    fieldStarted = False
'                    quoteClosed = False
'                    currentFieldQuoted = False
'                Case Else
'                    If quoteClosed Then Err.Raise vbObjectError + 7742, "SharedCsv_SplitLineDetailed", "Unexpected text appeared after a closing CSV quote at character " & CStr(i) & "."
'                    currentValue = currentValue & ch
'                    fieldStarted = True
'            End Select
'        End If
'    Next i
'    If inQuotes Then Err.Raise vbObjectError + 7743, "SharedCsv_SplitLineDetailed", "The CSV record ended before a quoted field was closed."
'    results(UBound(results)) = currentValue
'    quotedFields(UBound(quotedFields)) = currentFieldQuoted
'    SharedCsv_SplitLineDetailed = results
'End Function
'
'Public Sub SharedCsv_AssignValueToField(ByVal targetField As DAO.Field, ByVal csvValue As String)
'    'Assigns one unquoted CSV value using explicit DAO-type conversion and raises for unsupported field types.
'    'Use SharedCsv_ImportToTable when quoted-empty text must remain distinct from an unquoted Null field.
'    If targetField Is Nothing Then Err.Raise vbObjectError + 7744, "SharedCsv_AssignValueToField", "The target DAO.Field is Nothing."
'    If (targetField.attributes And dbAutoIncrField) <> 0 Then Err.Raise vbObjectError + 7745, "SharedCsv_AssignValueToField", "CSV assignment to an AutoNumber field is not supported." & vbCrLf & "Field: " & targetField.Name
'    targetField.value = SharedCsv_ValueForDaoFieldType(csvValue, targetField.Type, targetField.Name, False)
'End Sub
'
'Private Sub SharedCsv_AssignParsedValueToField(ByVal targetField As DAO.Field, ByVal csvValue As String, ByVal valueWasQuoted As Boolean)
'    'Assigns one parsed CSV field while preserving the distinction between quoted empty text and unquoted Null.
'    If targetField Is Nothing Then Err.Raise vbObjectError + 7793, "SharedCsv_AssignParsedValueToField", "The target DAO.Field is Nothing."
'    If (targetField.attributes And dbAutoIncrField) <> 0 Then Err.Raise vbObjectError + 7794, "SharedCsv_AssignParsedValueToField", "CSV assignment to an AutoNumber field is not supported." & vbCrLf & "Field: " & targetField.Name
'    targetField.value = SharedCsv_ValueForDaoFieldType(csvValue, targetField.Type, targetField.Name, valueWasQuoted)
'End Sub
'
'Private Function SharedCsv_ValueForDaoFieldType(ByVal csvValue As String, ByVal fieldType As Long, ByVal fieldName As String, Optional ByVal valueWasQuoted As Boolean = False) As Variant
'    'Converts one nonbinary CSV value according to a DAO field type using invariant numeric and ISO date parsing.
'    If Len(csvValue) = 0 Then
'        If valueWasQuoted And (fieldType = dbText Or fieldType = dbMemo Or fieldType = 18) Then
'            SharedCsv_ValueForDaoFieldType = vbNullString
'        Else
'            SharedCsv_ValueForDaoFieldType = Null
'        End If
'        Exit Function
'    End If
'    Select Case fieldType
'        Case dbBoolean
'            SharedCsv_ValueForDaoFieldType = SharedCsv_TextToBoolean(csvValue)
'        Case dbByte
'            SharedCsv_ValueForDaoFieldType = CByte(SharedCsv_IntegerValue(csvValue, 0#, 255#, "Byte"))
'        Case dbInteger
'            SharedCsv_ValueForDaoFieldType = CInt(SharedCsv_IntegerValue(csvValue, -32768#, 32767#, "Integer"))
'        Case dbLong
'            SharedCsv_ValueForDaoFieldType = CLng(SharedCsv_IntegerValue(csvValue, -2147483648#, 2147483647#, "Long"))
'        Case dbSingle
'            SharedCsv_ValueForDaoFieldType = CSng(SharedCsv_InvariantNumber(csvValue))
'        Case dbDouble
'            SharedCsv_ValueForDaoFieldType = CDbl(SharedCsv_InvariantNumber(csvValue))
'        Case dbCurrency
'            SharedCsv_ValueForDaoFieldType = SharedCsv_InvariantCurrency(csvValue)
'        Case dbDecimal
'            SharedCsv_ValueForDaoFieldType = SharedCsv_InvariantDecimal(csvValue)
'        Case dbDate
'            SharedCsv_ValueForDaoFieldType = SharedCsv_IsoDateValue(csvValue)
'        Case dbText, dbMemo, 18
'            SharedCsv_ValueForDaoFieldType = csvValue
'        Case dbGUID
'            SharedCsv_ValueForDaoFieldType = SharedCsv_NormalizedGuid(csvValue)
'        Case Else
'            Err.Raise vbObjectError + 7773, "SharedCsv_ValueForDaoFieldType", "CSV conversion is not supported for this DAO field type." & vbCrLf & "Field: " & fieldName & vbCrLf & "DAO field type: " & CStr(fieldType)
'    End Select
'End Function
'
'Private Function SharedCsv_InvariantNumber(ByVal valueText As String) As Double
'    'Parses a complete invariant numeric token into Double after validating every character.
'    Dim normalized As String
'    normalized = SharedCsv_ValidateInvariantNumber(valueText)
'    SharedCsv_InvariantNumber = Val(normalized)
'End Function
'
'Private Function SharedCsv_ValidateInvariantNumber(ByVal valueText As String) As String
'    'Validates an invariant numeric token with optional decimal and exponent notation.
'    Dim s As String
'    Dim i As Long
'    Dim ch As String
'    Dim digitSeen As Boolean
'    Dim decimalSeen As Boolean
'    Dim exponentSeen As Boolean
'    Dim exponentDigitSeen As Boolean
'    s = Trim$(valueText)
'    If Len(s) = 0 Then Err.Raise vbObjectError + 7774, "SharedCsv_ValidateInvariantNumber", "The numeric CSV value is blank."
'    For i = 1 To Len(s)
'        ch = Mid$(s, i, 1)
'        If ch >= "0" And ch <= "9" Then
'            digitSeen = True
'            If exponentSeen Then exponentDigitSeen = True
'        ElseIf (ch = "+" Or ch = "-") And (i = 1 Or (exponentSeen And i > 1 And (Mid$(s, i - 1, 1) = "e" Or Mid$(s, i - 1, 1) = "E"))) Then
'        ElseIf ch = "." And decimalSeen = False And exponentSeen = False Then
'            decimalSeen = True
'        ElseIf (ch = "e" Or ch = "E") And exponentSeen = False And digitSeen Then
'            exponentSeen = True
'        Else
'            Err.Raise vbObjectError + 7775, "SharedCsv_ValidateInvariantNumber", "Invalid invariant numeric CSV value: " & valueText
'        End If
'    Next i
'    If digitSeen = False Then Err.Raise vbObjectError + 7776, "SharedCsv_ValidateInvariantNumber", "The numeric CSV value contains no digits: " & valueText
'    If exponentSeen And exponentDigitSeen = False Then Err.Raise vbObjectError + 7777, "SharedCsv_ValidateInvariantNumber", "The numeric CSV exponent contains no digits: " & valueText
'    SharedCsv_ValidateInvariantNumber = s
'End Function
'
'Private Function SharedCsv_InvariantDecimal(ByVal valueText As String) As Variant
'    'Parses an invariant decimal token without an intermediate Double conversion.
'    On Error GoTo ErrHandler
'    Dim s As String
'    Dim mantissaText As String
'    Dim digitsText As String
'    Dim exponentText As String
'    Dim exponentValue As Long
'    Dim decimalPosition As Long
'    Dim scaleValue As Long
'    Dim negativeValue As Boolean
'    Dim separatorPosition As Long
'    Dim exponentPosition As Long
'    Dim i As Long
'    Dim resultValue As Variant
'    s = SharedCsv_ValidateInvariantNumber(valueText)
'    If Left$(s, 1) = "-" Or Left$(s, 1) = "+" Then
'        negativeValue = (Left$(s, 1) = "-")
'        s = Mid$(s, 2)
'    End If
'    exponentPosition = InStr(1, s, "e", vbTextCompare)
'    If exponentPosition > 0 Then
'        mantissaText = Left$(s, exponentPosition - 1)
'        exponentText = Mid$(s, exponentPosition + 1)
'        exponentValue = CLng(exponentText)
'    Else
'        mantissaText = s
'    End If
'    separatorPosition = InStr(1, mantissaText, ".", vbBinaryCompare)
'    If separatorPosition > 0 Then
'        decimalPosition = separatorPosition - 1
'        digitsText = Replace(mantissaText, ".", vbNullString)
'    Else
'        decimalPosition = Len(mantissaText)
'        digitsText = mantissaText
'    End If
'    decimalPosition = decimalPosition + exponentValue
'    Do While Len(digitsText) > 1 And Left$(digitsText, 1) = "0"
'        digitsText = Mid$(digitsText, 2)
'        decimalPosition = decimalPosition - 1
'    Loop
'    If Replace(digitsText, "0", vbNullString) = vbNullString Then
'        SharedCsv_InvariantDecimal = CDec(0)
'        Exit Function
'    End If
'    scaleValue = Len(digitsText) - decimalPosition
'    Do While scaleValue > 0 And Right$(digitsText, 1) = "0"
'        digitsText = Left$(digitsText, Len(digitsText) - 1)
'        scaleValue = scaleValue - 1
'    Loop
'    If Len(digitsText) > 29 Then Err.Raise vbObjectError + 7795, "SharedCsv_InvariantDecimal", "The CSV Decimal value contains more than 29 significant digits: " & valueText
'    If scaleValue > 28 Then Err.Raise vbObjectError + 7796, "SharedCsv_InvariantDecimal", "The CSV Decimal value requires more than 28 decimal places: " & valueText
'    If decimalPosition > 29 Then Err.Raise vbObjectError + 7797, "SharedCsv_InvariantDecimal", "The CSV Decimal value exceeds the supported integer-digit range: " & valueText
'    resultValue = CDec(0)
'    For i = 1 To Len(digitsText)
'        resultValue = resultValue * CDec(10) + CDec(Asc(Mid$(digitsText, i, 1)) - Asc("0"))
'    Next i
'    If scaleValue > 0 Then
'        For i = 1 To scaleValue
'            resultValue = resultValue / CDec(10)
'        Next i
'    ElseIf scaleValue < 0 Then
'        For i = 1 To -scaleValue
'            resultValue = resultValue * CDec(10)
'        Next i
'    End If
'    If negativeValue Then resultValue = -resultValue
'    SharedCsv_InvariantDecimal = resultValue
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedCsv_InvariantDecimal", "Could not convert the invariant CSV value to Decimal." & vbCrLf & "Value: " & valueText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'
'Private Function SharedCsv_InvariantCurrency(ByVal valueText As String) As Currency
'    'Parses an invariant Currency token exactly and rejects values that would require rounding.
'    On Error GoTo ErrHandler
'    Dim decimalValue As Variant
'    Dim currencyValue As Currency
'    decimalValue = SharedCsv_InvariantDecimal(valueText)
'    currencyValue = CCur(decimalValue)
'    If CDec(currencyValue) <> decimalValue Then Err.Raise vbObjectError + 7798, "SharedCsv_InvariantCurrency", "The CSV Currency value cannot be represented exactly with four decimal places: " & valueText
'    SharedCsv_InvariantCurrency = currencyValue
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedCsv_InvariantCurrency", "Could not convert the invariant CSV value to Currency." & vbCrLf & "Value: " & valueText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'
'Private Function SharedCsv_IntegerValue(ByVal valueText As String, ByVal minimumValue As Double, ByVal maximumValue As Double, ByVal typeText As String) As Double
'    'Parses and range-checks an invariant integer CSV value.
'    Dim numericValue As Double
'    numericValue = SharedCsv_InvariantNumber(valueText)
'    If numericValue <> Fix(numericValue) Then Err.Raise vbObjectError + 7778, "SharedCsv_IntegerValue", "The CSV value is not a whole " & typeText & " value: " & valueText
'    If numericValue < minimumValue Or numericValue > maximumValue Then Err.Raise vbObjectError + 7779, "SharedCsv_IntegerValue", "The CSV " & typeText & " value is out of range: " & valueText
'    SharedCsv_IntegerValue = numericValue
'End Function
'
'Private Function SharedCsv_IsoDateValue(ByVal valueText As String) As Date
'    'Parses yyyy-mm-dd with optional HH:NN or HH:NN:SS time using locale-independent components.
'    Dim normalized As String
'    Dim dateTimeParts As Variant
'    Dim dateParts As Variant
'    Dim timeParts As Variant
'    Dim yearValue As Long
'    Dim monthValue As Long
'    Dim dayValue As Long
'    Dim hourValue As Long
'    Dim minuteValue As Long
'    Dim secondValue As Long
'    Dim resultValue As Date
'    normalized = SharedText_CollapseWhitespace(Replace(Trim$(valueText), "T", " "))
'    dateTimeParts = Split(normalized, " ")
'    If UBound(dateTimeParts) > 1 Then Err.Raise vbObjectError + 7780, "SharedCsv_IsoDateValue", "Invalid ISO CSV date/time value: " & valueText
'    dateParts = Split(CStr(dateTimeParts(0)), "-")
'    If UBound(dateParts) <> 2 Then Err.Raise vbObjectError + 7781, "SharedCsv_IsoDateValue", "CSV dates must use yyyy-mm-dd format: " & valueText
'    yearValue = SharedCsv_UnsignedLong(CStr(dateParts(0)), "year")
'    monthValue = SharedCsv_UnsignedLong(CStr(dateParts(1)), "month")
'    dayValue = SharedCsv_UnsignedLong(CStr(dateParts(2)), "day")
'    If yearValue < 100 Or yearValue > 9999 Then Err.Raise vbObjectError + 7782, "SharedCsv_IsoDateValue", "CSV date year is out of range: " & valueText
'    If monthValue < 1 Or monthValue > 12 Then Err.Raise vbObjectError + 7783, "SharedCsv_IsoDateValue", "CSV date month is out of range: " & valueText
'    If dayValue < 1 Or dayValue > 31 Then Err.Raise vbObjectError + 7784, "SharedCsv_IsoDateValue", "CSV date day is out of range: " & valueText
'    resultValue = DateSerial(yearValue, monthValue, dayValue)
'    If Year(resultValue) <> yearValue Or Month(resultValue) <> monthValue Or Day(resultValue) <> dayValue Then Err.Raise vbObjectError + 7785, "SharedCsv_IsoDateValue", "CSV date is not a valid calendar date: " & valueText
'    If UBound(dateTimeParts) = 1 Then
'        timeParts = Split(CStr(dateTimeParts(1)), ":")
'        If UBound(timeParts) < 1 Or UBound(timeParts) > 2 Then Err.Raise vbObjectError + 7786, "SharedCsv_IsoDateValue", "CSV time must use HH:NN or HH:NN:SS format: " & valueText
'        hourValue = SharedCsv_UnsignedLong(CStr(timeParts(0)), "hour")
'        minuteValue = SharedCsv_UnsignedLong(CStr(timeParts(1)), "minute")
'        If UBound(timeParts) = 2 Then secondValue = SharedCsv_UnsignedLong(CStr(timeParts(2)), "second")
'        If hourValue > 23 Or minuteValue > 59 Or secondValue > 59 Then Err.Raise vbObjectError + 7787, "SharedCsv_IsoDateValue", "CSV time component is out of range: " & valueText
'        resultValue = resultValue + TimeSerial(hourValue, minuteValue, secondValue)
'    End If
'    SharedCsv_IsoDateValue = resultValue
'End Function
'
'Private Function SharedCsv_UnsignedLong(ByVal valueText As String, ByVal componentText As String) As Long
'    'Parses a required unsigned decimal component without locale-dependent conversion.
'    Dim i As Long
'    Dim ch As String
'    If Len(valueText) = 0 Then Err.Raise vbObjectError + 7788, "SharedCsv_UnsignedLong", "The CSV " & componentText & " component is blank."
'    For i = 1 To Len(valueText)
'        ch = Mid$(valueText, i, 1)
'        If ch < "0" Or ch > "9" Then Err.Raise vbObjectError + 7789, "SharedCsv_UnsignedLong", "The CSV " & componentText & " component is not numeric: " & valueText
'    Next i
'    SharedCsv_UnsignedLong = CLng(valueText)
'End Function
'
'Private Function SharedCsv_NormalizedGuid(ByVal valueText As String) As String
'    'Validates a canonical GUID token and returns it enclosed in braces for DAO assignment.
'    Dim coreText As String
'    Dim i As Long
'    Dim ch As String
'    coreText = Trim$(valueText)
'    If Left$(coreText, 1) = "{" And Right$(coreText, 1) = "}" Then coreText = Mid$(coreText, 2, Len(coreText) - 2)
'    If Len(coreText) <> 36 Then Err.Raise vbObjectError + 7790, "SharedCsv_NormalizedGuid", "The CSV GUID must contain 36 canonical characters: " & valueText
'    For i = 1 To Len(coreText)
'        ch = Mid$(coreText, i, 1)
'        If i = 9 Or i = 14 Or i = 19 Or i = 24 Then
'            If ch <> "-" Then Err.Raise vbObjectError + 7791, "SharedCsv_NormalizedGuid", "The CSV GUID has an invalid separator position: " & valueText
'        ElseIf InStr(1, "0123456789abcdefABCDEF", ch, vbBinaryCompare) = 0 Then
'            Err.Raise vbObjectError + 7792, "SharedCsv_NormalizedGuid", "The CSV GUID contains an invalid hexadecimal character: " & valueText
'        End If
'    Next i
'    SharedCsv_NormalizedGuid = "{" & coreText & "}"
'End Function
'
'Private Function SharedCsv_NormalizeLineEndings(ByVal sourceText As String) As String
'    'Normalizes CRLF, LF-only, and CR-only input to LF before physical-line parsing.
'    sourceText = Replace(sourceText, vbCrLf, vbLf)
'    SharedCsv_NormalizeLineEndings = Replace(sourceText, vbCr, vbLf)
'End Function
'
'Private Function SharedCsv_ReadUtf8Record(ByVal csvText As String, ByRef nextPosition As Long, ByRef physicalLineNum As Long, ByVal csvPath As String) As String
'    'Reads one logical CSV record from normalized UTF-8 text while preserving embedded quoted line breaks.
'    Dim txt As String
'    Dim lineText As String
'    Dim separatorPosition As Long
'    Dim physicalLinesInRecord As Long
'    If nextPosition < 1 Then nextPosition = 1
'    If nextPosition > Len(csvText) Then Exit Function
'    Do
'        separatorPosition = InStr(nextPosition, csvText, vbLf, vbBinaryCompare)
'        If separatorPosition = 0 Then
'            lineText = Mid$(csvText, nextPosition)
'            nextPosition = Len(csvText) + 1
'        Else
'            lineText = Mid$(csvText, nextPosition, separatorPosition - nextPosition)
'            nextPosition = separatorPosition + 1
'        End If
'        physicalLineNum = physicalLineNum + 1
'        physicalLinesInRecord = physicalLinesInRecord + 1
'        If physicalLinesInRecord = 1 Then
'            txt = lineText
'        Else
'            txt = txt & vbCrLf & lineText
'        End If
'        If SharedCsv_RecordIsComplete(txt) Then Exit Do
'        If nextPosition > Len(csvText) Then Err.Raise vbObjectError + 7799, "SharedCsv_ReadUtf8Record", "Malformed CSV record. Quoted field was not closed." & vbCrLf & "CSV File: " & csvPath & vbCrLf & "Physical Row: " & physicalLineNum
'    Loop
'    SharedCsv_ReadUtf8Record = txt
'End Function
'
'Private Sub SharedCsv_ReportImportProgress(ByVal tableName As String, ByVal recordNum As Long, ByVal physicalLineNum As Long)
'    'Reports bounded importer-owned diagnostics without overwriting or clearing a caller-owned Access status bar.
'    If recordNum <> 1 And recordNum Mod 250 <> 0 Then Exit Sub
'    On Error Resume Next
'    Debug.Print "CSV import to " & tableName & ": logical record " & Format$(recordNum, "#,##0") & "; physical line " & Format$(physicalLineNum, "#,##0")
'    DoEvents
'    On Error GoTo 0
'End Sub
'
'Public Sub SharedCsv_ImportToTable(ByVal csvPath As String, ByVal tableName As String, Optional ByVal logPath As String = vbNullString)
'    'Imports UTF-8 CSV into a DAO table, preserves quoted empty text, and rolls back the complete import on failure.
'    'Consolidated old names: local TV NSS IMPORT MODULE ImportCsvToTable.
'    On Error GoTo ErrHandler
'    Dim db As DAO.Database
'    Dim rs As DAO.Recordset
'    Dim ws As DAO.workspace
'    Dim ts As Object
'    Dim csvText As String
'    Dim nextPosition As Long
'    Dim headerRecord As String
'    Dim dataRecord As String
'    Dim headers() As String
'    Dim values() As String
'    Dim quotedValues() As Boolean
'    Dim fieldMap As Object
'    Dim headerMap As Object
'    Dim fld As DAO.Field
'    Dim i As Long
'    Dim expectedColumnCount As Long
'    Dim actualColumnCount As Long
'    Dim physicalLineNum As Long
'    Dim recordNum As Long
'    Dim fieldName As String
'    Dim actualFieldName As String
'    Dim fieldValue As String
'    Dim fieldType As Long
'    Dim stageText As String
'    Dim fieldAssignmentActive As Boolean
'    Dim transactionStarted As Boolean
'    Dim savedErrorNumber As Long
'    Dim savedErrorSource As String
'    Dim savedErrorDescription As String
'    Dim cancelErrorNumber As Long
'    Dim cancelErrorDescription As String
'    Dim rollbackErrorNumber As Long
'    Dim rollbackErrorDescription As String
'    Dim recordsetCloseErrorNumber As Long
'    Dim recordsetCloseErrorDescription As String
'    Dim streamCloseErrorNumber As Long
'    Dim streamCloseErrorDescription As String
'    Dim reportErrorNumber As Long
'    Dim reportErrorDescription As String
'    Dim rollbackStatusText As String
'    Dim failureText As String
'    Dim importCommitted As Boolean
'    Dim cleanupFailureText As String
'    csvPath = Trim$(csvPath)
'    If Len(csvPath) = 0 Then Err.Raise vbObjectError + 7711, "SharedCsv_ImportToTable", "The CSV path is blank."
'    If Len(tableName) = 0 Then Err.Raise vbObjectError + 7712, "SharedCsv_ImportToTable", "The target table name is blank."
'    stageText = "opening target table"
'    Set db = currentDb
'    Set ws = DBEngine.Workspaces(0)
'    Set rs = db.OpenRecordset(tableName, dbOpenDynaset, dbAppendOnly)
'    Set fieldMap = CreateObject("Scripting.Dictionary")
'    fieldMap.CompareMode = vbTextCompare
'    For Each fld In rs.fields
'        fieldMap(fld.Name) = fld.Name
'    Next fld
'    stageText = "opening UTF-8 CSV file"
'    Set ts = CreateObject("ADODB.Stream")
'    ts.Type = SHARED_ADODB_STREAM_TYPE_TEXT
'    ts.charset = SHARED_ADODB_STREAM_CHARSET_UTF8
'    ts.Open
'    ts.LoadFromFile csvPath
'    csvText = CStr(ts.ReadText(SHARED_ADODB_STREAM_READ_ALL))
'    csvText = SharedCsv_NormalizeLineEndings(csvText)
'    If Len(csvText) = 0 Then GoTo CleanExit
'    nextPosition = 1
'    physicalLineNum = 0
'    recordNum = 0
'    stageText = "reading CSV header"
'    headerRecord = SharedText_RemoveBom(SharedCsv_ReadUtf8Record(csvText, nextPosition, physicalLineNum, csvPath))
'    headers = SharedCsv_SplitLine(headerRecord)
'    expectedColumnCount = UBound(headers) - LBound(headers) + 1
'    Set headerMap = CreateObject("Scripting.Dictionary")
'    headerMap.CompareMode = vbTextCompare
'    stageText = "validating CSV header"
'    For i = LBound(headers) To UBound(headers)
'        fieldName = SharedText_RemoveBom(Trim$(headers(i)))
'        If Len(fieldName) = 0 Then Err.Raise vbObjectError + 7746, "SharedCsv_ImportToTable", "The CSV header contains a blank field name." & vbCrLf & "Column: " & CStr(i + 1)
'        If headerMap.Exists(fieldName) Then Err.Raise vbObjectError + 7747, "SharedCsv_ImportToTable", "The CSV header contains a duplicate field name." & vbCrLf & "Field: " & fieldName & vbCrLf & "Column: " & CStr(i + 1)
'        headerMap.Add fieldName, True
'        If fieldMap.Exists(fieldName) = False Then Err.Raise vbObjectError + 7748, "SharedCsv_ImportToTable", "The CSV header references a field that does not exist in the target table." & vbCrLf & "Field: " & fieldName & vbCrLf & "Table: " & tableName
'        headers(i) = CStr(fieldMap(fieldName))
'    Next i
'    ws.BeginTrans
'    transactionStarted = True
'    Do While nextPosition <= Len(csvText)
'        stageText = "reading logical CSV record"
'        dataRecord = SharedCsv_ReadUtf8Record(csvText, nextPosition, physicalLineNum, csvPath)
'        recordNum = recordNum + 1
'        SharedCsv_ReportImportProgress tableName, recordNum, physicalLineNum
'        values = SharedCsv_SplitLineDetailed(dataRecord, quotedValues)
'        actualColumnCount = UBound(values) - LBound(values) + 1
'        If actualColumnCount <> expectedColumnCount Then Err.Raise vbObjectError + 7749, "SharedCsv_ImportToTable", "The CSV record column count does not match the header." & vbCrLf & "Expected columns: " & CStr(expectedColumnCount) & vbCrLf & "Actual columns: " & CStr(actualColumnCount)
'        stageText = "starting target record"
'        rs.AddNew
'        For i = LBound(headers) To UBound(headers)
'            fieldName = headers(i)
'            actualFieldName = fieldName
'            fieldValue = values(i)
'            fieldType = rs.fields(actualFieldName).Type
'            stageText = "assigning CSV field"
'            fieldAssignmentActive = True
'            SharedCsv_AssignParsedValueToField rs.fields(actualFieldName), fieldValue, quotedValues(i)
'            fieldAssignmentActive = False
'        Next i
'        stageText = "committing target record"
'        rs.Update
'    Loop
'    stageText = "committing CSV import transaction"
'    ws.CommitTrans
'    transactionStarted = False
'    importCommitted = True
'    Debug.Print "CSV import to " & tableName & " completed: " & Format$(recordNum, "#,##0") & " logical record(s)."
'CleanExit:
'    On Error Resume Next
'    If Not rs Is Nothing Then
'        Err.Clear
'        rs.Close
'        recordsetCloseErrorNumber = Err.Number
'        recordsetCloseErrorDescription = Err.Description
'    End If
'    If Not ts Is Nothing Then
'        Err.Clear
'        ts.Close
'        streamCloseErrorNumber = Err.Number
'        streamCloseErrorDescription = Err.Description
'    End If
'    On Error GoTo 0
'    Set headerMap = Nothing
'    Set fieldMap = Nothing
'    Set rs = Nothing
'    Set ts = Nothing
'    Set ws = Nothing
'    Set db = Nothing
'    If recordsetCloseErrorNumber <> 0 Then cleanupFailureText = SharedText_AppendNonBlankLine(cleanupFailureText, "Recordset-close error " & CStr(recordsetCloseErrorNumber) & ": " & recordsetCloseErrorDescription)
'    If streamCloseErrorNumber <> 0 Then cleanupFailureText = SharedText_AppendNonBlankLine(cleanupFailureText, "CSV-stream close error " & CStr(streamCloseErrorNumber) & ": " & streamCloseErrorDescription)
'    If Len(cleanupFailureText) > 0 Then
'        If importCommitted Then Err.Raise vbObjectError + 8423, "SharedCsv_ImportToTable", "The CSV import transaction committed successfully, but resource cleanup failed. Do not rerun the import solely because of this cleanup error." & vbCrLf & cleanupFailureText
'        Err.Raise vbObjectError + 8424, "SharedCsv_ImportToTable", "The CSV import did not start, but resource cleanup failed." & vbCrLf & cleanupFailureText
'    End If
'    Exit Sub
'ErrHandler:
'    savedErrorNumber = Err.Number
'    savedErrorSource = Err.Source
'    savedErrorDescription = Err.Description
'    rollbackStatusText = "No import transaction had started."
'    On Error Resume Next
'    If Not rs Is Nothing Then
'        Err.Clear
'        If rs.EditMode <> dbEditNone Then rs.CancelUpdate
'        cancelErrorNumber = Err.Number
'        cancelErrorDescription = Err.Description
'    End If
'    If transactionStarted And Not ws Is Nothing Then
'        Err.Clear
'        ws.Rollback
'        rollbackErrorNumber = Err.Number
'        rollbackErrorDescription = Err.Description
'        If rollbackErrorNumber = 0 Then
'            transactionStarted = False
'            rollbackStatusText = "The import transaction was rolled back."
'        Else
'            rollbackStatusText = "The import transaction rollback failed."
'        End If
'    End If
'    If Not rs Is Nothing Then
'        Err.Clear
'        rs.Close
'        recordsetCloseErrorNumber = Err.Number
'        recordsetCloseErrorDescription = Err.Description
'    End If
'    If Not ts Is Nothing Then
'        Err.Clear
'        ts.Close
'        streamCloseErrorNumber = Err.Number
'        streamCloseErrorDescription = Err.Description
'    End If
'    On Error GoTo 0
'    If fieldAssignmentActive Then
'        On Error Resume Next
'        SharedCsv_ShowImportFailure tableName, csvPath, recordNum, physicalLineNum, i + 1, fieldName, actualFieldName, fieldType, fieldValue, savedErrorNumber, savedErrorDescription, logPath
'        reportErrorNumber = Err.Number
'        reportErrorDescription = Err.Description
'        On Error GoTo 0
'    End If
'    Set headerMap = Nothing
'    Set fieldMap = Nothing
'    Set rs = Nothing
'    Set ts = Nothing
'    Set ws = Nothing
'    Set db = Nothing
'    failureText = "CSV import failed." & vbCrLf & rollbackStatusText
'    failureText = failureText & vbCrLf & "Stage: " & stageText
'    failureText = failureText & vbCrLf & "Target table: " & tableName
'    failureText = failureText & vbCrLf & "CSV file: " & csvPath
'    failureText = failureText & vbCrLf & "Logical record: " & CStr(recordNum)
'    failureText = failureText & vbCrLf & "Physical line: " & CStr(physicalLineNum)
'    failureText = failureText & vbCrLf & "CSV field: " & fieldName
'    failureText = failureText & vbCrLf & "Table field: " & actualFieldName
'    failureText = failureText & vbCrLf & "Original source: " & savedErrorSource
'    failureText = failureText & vbCrLf & "Description: " & savedErrorDescription
'    If cancelErrorNumber <> 0 Then failureText = failureText & vbCrLf & "CancelUpdate error " & CStr(cancelErrorNumber) & ": " & cancelErrorDescription
'    If rollbackErrorNumber <> 0 Then failureText = failureText & vbCrLf & "Rollback error " & CStr(rollbackErrorNumber) & ": " & rollbackErrorDescription
'    If recordsetCloseErrorNumber <> 0 Then failureText = failureText & vbCrLf & "Recordset-close error " & CStr(recordsetCloseErrorNumber) & ": " & recordsetCloseErrorDescription
'    If streamCloseErrorNumber <> 0 Then failureText = failureText & vbCrLf & "CSV-stream close error " & CStr(streamCloseErrorNumber) & ": " & streamCloseErrorDescription
'    If reportErrorNumber <> 0 Then failureText = failureText & vbCrLf & "Failure-reporting error " & CStr(reportErrorNumber) & ": " & reportErrorDescription
'    Err.Raise savedErrorNumber, "SharedCsv_ImportToTable", failureText
'End Sub
'Public Sub SharedCsv_ShowImportFailure(ByVal tableName As String, ByVal csvPath As String, ByVal recordNum As Long, ByVal physicalLineNum As Long, ByVal ColumnIndex As Long, ByVal csvFieldName As String, ByVal tableFieldName As String, ByVal tableFieldType As Long, ByVal csvValue As String, ByVal errorNumber As Long, ByVal errorDescription As String, Optional ByVal logPath As String = vbNullString)
'    'Displays a detailed CSV import field-assignment failure message.
'    'Consolidated old names: local TV NSS IMPORT MODULE ShowCsvImportFailure.
'    Dim msg As String
'    msg = "CSV import failed while assigning a field." & vbCrLf & vbCrLf
'    msg = msg & "Target table: " & tableName & vbCrLf
'    msg = msg & "CSV file: " & csvPath & vbCrLf
'    msg = msg & "Logical record: " & CStr(recordNum) & vbCrLf
'    msg = msg & "Physical line: " & CStr(physicalLineNum) & vbCrLf
'    msg = msg & "Column index: " & CStr(ColumnIndex) & vbCrLf
'    msg = msg & "CSV field: " & csvFieldName & vbCrLf
'    msg = msg & "Table field: " & tableFieldName & vbCrLf
'    msg = msg & "DAO field type: " & CStr(tableFieldType) & vbCrLf
'    msg = msg & "CSV value: " & csvValue & vbCrLf & vbCrLf
'    msg = msg & "Error " & CStr(errorNumber) & ": " & errorDescription
'    SharedMessage_ShowFailureDialogWithOpenLog msg, "Import Failure", logPath
'End Sub
'
'' ================================================================================================================================================================================
'' DAO
'' ================================================================================================================================================================================
'Public Function SharedDao_TableExists(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
'    'Returns True when the supplied DAO database contains the named table.
'    'Technical inspection failures are raised; use SharedDao_DatabaseTableExistsSafe only when False-on-error behavior is intentional.
'    If db Is Nothing Then Err.Raise 91, "SharedDao_TableExists", "The supplied DAO.Database is Nothing."
'    SharedDao_TableExists = SharedDao_DatabaseTableExists(db, tableName)
'End Function
'
'Public Function SharedDao_QueryExists(ByVal db As DAO.Database, ByVal queryName As String) As Boolean
'    'Returns True when the supplied DAO database contains the named query.
'    'Technical inspection failures are raised so callers do not mistake an unreadable collection for a missing query.
'    If db Is Nothing Then Err.Raise 91, "SharedDao_QueryExists", "The supplied DAO.Database is Nothing."
'    SharedDao_QueryExists = SharedDao_DatabaseQueryExists(db, queryName)
'End Function
'
'Public Function SharedDao_FieldExists(ByVal tdf As DAO.tableDef, ByVal fieldName As String) As Boolean
'    'Returns True when the supplied DAO TableDef contains the named field.
'    'Technical inspection failures are raised; use SharedDao_TableDefFieldExistsSafe only when False-on-error behavior is intentional.
'    If tdf Is Nothing Then Err.Raise 91, "SharedDao_FieldExists", "The supplied DAO.TableDef is Nothing."
'    SharedDao_FieldExists = SharedDao_TableDefFieldExists(tdf, fieldName)
'End Function
'
'Public Function SharedDao_TableFieldExists(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String) As Boolean
'    'Returns True when the supplied DAO database contains the table and that table contains the named field.
'    'Technical inspection failures are raised; use SharedDao_DatabaseTableFieldExistsSafe only when False-on-error behavior is intentional.
'    If db Is Nothing Then Err.Raise 91, "SharedDao_TableFieldExists", "The supplied DAO.Database is Nothing."
'    If SharedDao_DatabaseTableExists(db, tableName) = False Then Exit Function
'    SharedDao_TableFieldExists = SharedDao_TableDefFieldExists(db.TableDefs(tableName), fieldName)
'End Function
'
'Public Function SharedDao_IndexExists(ByVal tdf As DAO.tableDef, ByVal indexName As String) As Boolean
'    'Returns True when the supplied DAO TableDef contains the named index.
'    'Technical inspection failures are raised; use SharedDao_TableDefIndexExistsSafe only when False-on-error behavior is intentional.
'    If tdf Is Nothing Then Err.Raise 91, "SharedDao_IndexExists", "The supplied DAO.TableDef is Nothing."
'    SharedDao_IndexExists = SharedDao_TableDefIndexExists(tdf, indexName)
'End Function
'
'Public Function SharedDao_FieldDefinitionMatches(ByVal fld As DAO.Field, ByVal expectedType As DAO.DataTypeEnum, Optional ByVal expectedSize As Long = 0, Optional ByVal mustBeAutoNumber As Boolean = False, Optional ByVal allowLargerTextSize As Boolean = False) As Boolean
'    'Intended use: Determines whether an existing DAO field matches the requested type, size, AutoNumber, and text-size contract.
'    Dim fieldIsAutoNumber As Boolean
'    If fld Is Nothing Then Exit Function
'    fieldIsAutoNumber = ((fld.attributes And dbAutoIncrField) <> 0)
'    If mustBeAutoNumber Then
'        SharedDao_FieldDefinitionMatches = (fld.Type = dbLong And fieldIsAutoNumber)
'        Exit Function
'    End If
'    If fld.Type <> expectedType Then Exit Function
'    If expectedType = dbText And expectedSize > 0 Then
'        If allowLargerTextSize Then
'            If fld.Size < expectedSize Then Exit Function
'        Else
'            If fld.Size <> expectedSize Then Exit Function
'        End If
'    End If
'    If fieldIsAutoNumber Then Exit Function
'    SharedDao_FieldDefinitionMatches = True
'End Function
'
'Public Function SharedDao_IndexMatchesFields(ByVal idx As DAO.index, ByVal fieldNames As Variant, ByVal uniqueIndex As Boolean, ByVal primaryIndex As Boolean) As Boolean
'    'Intended use: Determines whether a DAO index matches the requested fields, uniqueness, and primary-key settings.
'    Dim i As Long
'    Dim expectedCount As Long
'    If idx Is Nothing Then Exit Function
'    If CBool(idx.Unique) <> CBool(uniqueIndex) Then Exit Function
'    If CBool(idx.Primary) <> CBool(primaryIndex) Then Exit Function
'    expectedCount = UBound(fieldNames) - LBound(fieldNames) + 1
'    If idx.fields.count <> expectedCount Then Exit Function
'    For i = LBound(fieldNames) To UBound(fieldNames)
'        If StrComp(idx.fields(i - LBound(fieldNames)).Name, CStr(fieldNames(i)), vbTextCompare) <> 0 Then Exit Function
'    Next i
'    SharedDao_IndexMatchesFields = True
'End Function
'
'Private Function SharedValue_VariantsMatchExact(ByVal firstValue As Variant, ByVal secondValue As Variant) As Boolean
'    'Compares scalar Variant values without collapsing Null, Empty, type, date, Boolean, numeric, or text distinctions.
'    On Error GoTo ValuesDoNotMatch
'    If VarType(firstValue) <> VarType(secondValue) Then Exit Function
'    If IsNull(firstValue) Then
'        SharedValue_VariantsMatchExact = IsNull(secondValue)
'        Exit Function
'    End If
'    If IsEmpty(firstValue) Then
'        SharedValue_VariantsMatchExact = IsEmpty(secondValue)
'        Exit Function
'    End If
'    Select Case VarType(firstValue)
'        Case vbString
'            SharedValue_VariantsMatchExact = (StrComp(CStr(firstValue), CStr(secondValue), vbBinaryCompare) = 0)
'        Case vbDate
'            SharedValue_VariantsMatchExact = (CDbl(firstValue) = CDbl(secondValue))
'        Case vbBoolean
'            SharedValue_VariantsMatchExact = (CBool(firstValue) = CBool(secondValue))
'        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
'            SharedValue_VariantsMatchExact = (firstValue = secondValue)
'        Case vbError
'            SharedValue_VariantsMatchExact = (CLng(firstValue) = CLng(secondValue))
'        Case Else
'            SharedValue_VariantsMatchExact = (StrComp(CStr(firstValue), CStr(secondValue), vbBinaryCompare) = 0)
'    End Select
'    Exit Function
'ValuesDoNotMatch:
'    Err.Clear
'End Function
'
'Public Function SharedDao_EnsurePropertyValue(ByVal targetObject As Object, ByVal propertyName As String, ByVal propertyType As DAO.DataTypeEnum, ByVal propertyValue As Variant, Optional ByVal callerName As String = "SharedDao_EnsurePropertyValue") As Boolean
'    'Ensures and verifies one DAO property, restoring its exact prior existence, DAO type, Variant subtype, and value when a write or verification fails.
'    On Error GoTo ErrHandler
'    Dim prp As DAO.Property
'    Dim originalValue As Variant
'    Dim verifiedValue As Variant
'    Dim originalType As Long
'    Dim verifiedType As Long
'    Dim originalExists As Boolean
'    Dim mutationStarted As Boolean
'    Dim lookupErrNumber As Long
'    Dim lookupErrSource As String
'    Dim lookupErrDescription As String
'    If targetObject Is Nothing Then Err.Raise vbObjectError + 8400, callerName, "The DAO property target is Nothing."
'    propertyName = Trim$(propertyName)
'    If Len(propertyName) = 0 Then Err.Raise vbObjectError + 8401, callerName, "The DAO property name is blank."
'    On Error Resume Next
'    Err.Clear
'    originalValue = targetObject.Properties(propertyName).value
'    lookupErrNumber = Err.Number
'    lookupErrSource = Err.Source
'    lookupErrDescription = Err.Description
'    If lookupErrNumber = 0 Then originalType = CLng(targetObject.Properties(propertyName).Type)
'    Err.Clear
'    On Error GoTo ErrHandler
'    If lookupErrNumber = 0 Then
'        originalExists = True
'        If originalType <> CLng(propertyType) Then Err.Raise vbObjectError + 8402, callerName, "The existing DAO property type does not match the requested type."
'        If SharedValue_VariantsMatchExact(originalValue, propertyValue) = False Then
'            mutationStarted = True
'            targetObject.Properties(propertyName).value = propertyValue
'            SharedDao_EnsurePropertyValue = True
'        End If
'    ElseIf lookupErrNumber = SHARED_DAO_ERROR_PROPERTY_NOT_FOUND Then
'        mutationStarted = True
'        Set prp = targetObject.CreateProperty(propertyName, propertyType, propertyValue)
'        targetObject.Properties.Append prp
'        SharedDao_EnsurePropertyValue = True
'    Else
'        Err.Raise lookupErrNumber, lookupErrSource, lookupErrDescription
'    End If
'    verifiedValue = targetObject.Properties(propertyName).value
'    verifiedType = CLng(targetObject.Properties(propertyName).Type)
'    If verifiedType <> CLng(propertyType) Then Err.Raise vbObjectError + 8403, callerName, "The DAO property type did not match after the ensure operation."
'    If SharedValue_VariantsMatchExact(verifiedValue, propertyValue) = False Then Err.Raise vbObjectError + 8435, callerName, "The DAO property Variant subtype or value did not match after the ensure operation."
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim rollbackErrNumber As Long
'    Dim rollbackErrDescription As String
'    Dim rollbackText As String
'    Dim currentExists As Boolean
'    Dim currentValue As Variant
'    Dim currentType As Long
'    Dim propertyLookupErrNumber As Long
'    Dim propertyLookupErrSource As String
'    Dim propertyLookupErrDescription As String
'    Dim errText As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    If mutationStarted Then
'        On Error Resume Next
'        Err.Clear
'        If originalExists Then
'            targetObject.Properties(propertyName).value = originalValue
'            If Err.Number = 0 Then currentValue = targetObject.Properties(propertyName).value
'            If Err.Number = 0 Then currentType = CLng(targetObject.Properties(propertyName).Type)
'            If Err.Number = 0 Then
'                If currentType <> originalType Then Err.Raise vbObjectError + 8436, callerName, "The original DAO property type was not preserved during rollback."
'            End If
'            If Err.Number = 0 Then
'                If SharedValue_VariantsMatchExact(currentValue, originalValue) = False Then Err.Raise vbObjectError + 8437, callerName, "The original DAO property Variant subtype or value was not restored."
'            End If
'        Else
'            currentExists = False
'            Err.Clear
'            currentValue = targetObject.Properties(propertyName).value
'            propertyLookupErrNumber = Err.Number
'            propertyLookupErrSource = Err.Source
'            propertyLookupErrDescription = Err.Description
'            Err.Clear
'            If propertyLookupErrNumber = 0 Then
'                currentExists = True
'            ElseIf propertyLookupErrNumber <> SHARED_DAO_ERROR_PROPERTY_NOT_FOUND Then
'                Err.Raise propertyLookupErrNumber, propertyLookupErrSource, propertyLookupErrDescription
'            End If
'            If currentExists Then
'                targetObject.Properties.Delete propertyName
'            End If
'            If Err.Number = 0 Then
'                currentExists = False
'                Err.Clear
'                currentValue = targetObject.Properties(propertyName).value
'                propertyLookupErrNumber = Err.Number
'                propertyLookupErrSource = Err.Source
'                propertyLookupErrDescription = Err.Description
'                Err.Clear
'                If propertyLookupErrNumber = 0 Then
'                    currentExists = True
'                ElseIf propertyLookupErrNumber <> SHARED_DAO_ERROR_PROPERTY_NOT_FOUND Then
'                    Err.Raise propertyLookupErrNumber, propertyLookupErrSource, propertyLookupErrDescription
'                End If
'            End If
'            If Err.Number = 0 Then
'                If currentExists Then Err.Raise vbObjectError + 8438, callerName, "The newly created DAO property remained after rollback."
'            End If
'        End If
'        rollbackErrNumber = Err.Number
'        rollbackErrDescription = Err.Description
'        On Error GoTo 0
'        If rollbackErrNumber = 0 Then
'            rollbackText = "The original DAO property state was restored and verified."
'        Else
'            rollbackText = "Rollback error " & CStr(rollbackErrNumber) & ": " & rollbackErrDescription
'        End If
'    End If
'    errText = "Unable to ensure and verify the DAO property value." & vbCrLf
'    errText = errText & "Caller: " & callerName & vbCrLf
'    errText = errText & "Property: " & propertyName & vbCrLf
'    errText = errText & "Property type: " & CStr(propertyType) & vbCrLf
'    errText = errText & "Requested value: " & SharedValue_ToDisplayText(propertyValue) & vbCrLf
'    errText = errText & "Requested VarType: " & CStr(VarType(propertyValue)) & vbCrLf
'    errText = errText & "Property previously existed: " & CStr(originalExists) & vbCrLf
'    errText = errText & "Original error: " & CStr(savedErrNumber) & vbCrLf
'    errText = errText & "Original source: " & savedErrSource & vbCrLf
'    errText = errText & "Original description: " & savedErrDescription
'    If Len(rollbackText) > 0 Then errText = errText & vbCrLf & "Rollback: " & rollbackText
'    Err.Raise savedErrNumber, callerName, errText
'End Function
'
'Public Sub SharedDao_GetPropertyState(ByVal targetObject As Object, ByVal propertyName As String, ByRef propertyExists As Boolean, ByRef propertyValue As Variant)
'    'Captures whether a DAO property exists and its value while distinguishing a missing property from an inspection failure.
'    On Error GoTo ErrHandler
'    Dim lookupErrNumber As Long
'    Dim lookupErrSource As String
'    Dim lookupErrDescription As String
'    If targetObject Is Nothing Then Err.Raise vbObjectError + 7410, "SharedDao_GetPropertyState", "The target object is Nothing."
'    propertyExists = False
'    propertyValue = Null
'    On Error Resume Next
'    Err.Clear
'    propertyValue = targetObject.Properties(propertyName).value
'    lookupErrNumber = Err.Number
'    lookupErrSource = Err.Source
'    lookupErrDescription = Err.Description
'    Err.Clear
'    On Error GoTo ErrHandler
'    If lookupErrNumber = 0 Then
'        propertyExists = True
'    ElseIf lookupErrNumber <> SHARED_DAO_ERROR_PROPERTY_NOT_FOUND Then
'        Err.Raise lookupErrNumber, lookupErrSource, lookupErrDescription
'    End If
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedDao_GetPropertyState", "Unable to inspect the DAO property state." & vbCrLf & "Property: " & propertyName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Sub
'
'Public Sub SharedDao_RestorePropertyState(ByVal targetObject As Object, ByVal propertyName As String, ByVal propertyType As DAO.DataTypeEnum, ByVal propertyExisted As Boolean, ByVal propertyValue As Variant)
'    'Restores one DAO property to its captured present/value or absent state and verifies its exact DAO type, Variant subtype, and value.
'    On Error GoTo ErrHandler
'    Dim currentExists As Boolean
'    Dim currentValue As Variant
'    If targetObject Is Nothing Then Err.Raise vbObjectError + 7411, "SharedDao_RestorePropertyState", "The target object is Nothing."
'    If propertyExisted Then
'        SharedDao_EnsurePropertyValue targetObject, propertyName, propertyType, propertyValue, "SharedDao_RestorePropertyState"
'        SharedDao_GetPropertyState targetObject, propertyName, currentExists, currentValue
'        If currentExists = False Then Err.Raise vbObjectError + 7412, "SharedDao_RestorePropertyState", "The restored property is missing."
'        If CLng(targetObject.Properties(propertyName).Type) <> CLng(propertyType) Then Err.Raise vbObjectError + 7413, "SharedDao_RestorePropertyState", "The restored property DAO type does not match the captured type."
'        If SharedValue_VariantsMatchExact(currentValue, propertyValue) = False Then Err.Raise vbObjectError + 7414, "SharedDao_RestorePropertyState", "The restored property Variant subtype or value does not match the captured value."
'    Else
'        SharedDao_GetPropertyState targetObject, propertyName, currentExists, currentValue
'        If currentExists Then targetObject.Properties.Delete propertyName
'        SharedDao_GetPropertyState targetObject, propertyName, currentExists, currentValue
'        If currentExists Then Err.Raise vbObjectError + 7415, "SharedDao_RestorePropertyState", "The property still exists after absent-state restoration."
'    End If
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedDao_RestorePropertyState", "Unable to restore the DAO property state." & vbCrLf & "Property: " & propertyName & vbCrLf & "Captured existence: " & CStr(propertyExisted) & vbCrLf & "Captured value: " & SharedValue_ToDisplayText(propertyValue) & vbCrLf & "Captured VarType: " & CStr(VarType(propertyValue)) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Sub
'
'Public Function SharedDao_EnsureTableExists(ByVal db As DAO.Database, ByVal tableName As String, ByVal tableDescription As String, Optional ByVal idFieldName As String = "ID", Optional ByVal createAutoNumberID As Boolean = True, Optional ByRef tableCreated As Boolean = False) As Boolean
'    'Intended use: Creates a missing DAO table and creates or verifies its requested Long ID field and description, rolling back mutations if verification fails.
'    On Error GoTo ErrHandler
'    Dim tdf As DAO.tableDef
'    Dim fld As DAO.Field
'    Dim changed As Boolean
'    Dim tableExistedBefore As Boolean
'    Dim idFieldExistedBefore As Boolean
'    Dim descriptionStateCaptured As Boolean
'    Dim descriptionExistedBefore As Boolean
'    Dim descriptionValueBefore As Variant
'    Dim preStateCaptured As Boolean
'    tableCreated = False
'    If db Is Nothing Then Err.Raise vbObjectError + 7400, "SharedDao_EnsureTableExists", "Database object is not set while ensuring table [" & tableName & "]."
'    idFieldName = Trim$(idFieldName)
'    If Len(tableName) = 0 Then Err.Raise vbObjectError + 7403, "SharedDao_EnsureTableExists", "The table name is blank."
'    db.TableDefs.Refresh
'    tableExistedBefore = SharedDao_DatabaseTableExists(db, tableName)
'    If tableExistedBefore Then
'        Set tdf = db.TableDefs(tableName)
'        tdf.fields.Refresh
'        If Len(idFieldName) > 0 Then idFieldExistedBefore = SharedDao_TableDefFieldExists(tdf, idFieldName)
'        If Len(tableDescription) > 0 Then
'            SharedDao_GetPropertyState tdf, "Description", descriptionExistedBefore, descriptionValueBefore
'            descriptionStateCaptured = True
'        End If
'        preStateCaptured = True
'    Else
'        preStateCaptured = True
'        Set tdf = db.CreateTableDef(tableName)
'        If Len(idFieldName) > 0 Then
'            Set fld = tdf.CreateField(idFieldName, dbLong)
'            If createAutoNumberID Then fld.attributes = fld.attributes Or dbAutoIncrField
'            tdf.fields.Append fld
'        End If
'        db.TableDefs.Append tdf
'        db.TableDefs.Refresh
'        tableCreated = True
'        changed = True
'        Set tdf = db.TableDefs(tableName)
'    End If
'    If Len(idFieldName) > 0 Then If SharedDao_EnsureDatabaseTableField(db, tableName, idFieldName, dbLong, 0, createAutoNumberID, False, vbNullString, False) Then changed = True
'    Set tdf = db.TableDefs(tableName)
'    If Len(idFieldName) > 0 Then
'        tdf.fields.Refresh
'        If SharedDao_TableDefFieldExists(tdf, idFieldName) = False Then Err.Raise vbObjectError + 7404, "SharedDao_EnsureTableExists", "The requested ID field is missing after table verification."
'        Set fld = tdf.fields(idFieldName)
'        If SharedDao_FieldDefinitionMatches(fld, dbLong, 0, createAutoNumberID, False) = False Then Err.Raise vbObjectError + 7405, "SharedDao_EnsureTableExists", "The requested ID field is incompatible after table verification." & vbCrLf & "Expected: " & SharedDao_FieldDefinitionText(dbLong, 0, createAutoNumberID) & vbCrLf & "Actual: " & SharedDao_ExistingFieldDefinitionText(fld)
'    End If
'    If Len(tableDescription) > 0 Then If SharedDao_EnsurePropertyValue(tdf, "Description", dbText, tableDescription, "SharedDao_EnsureTableExists") Then changed = True
'    SharedDao_EnsureTableExists = changed
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim cleanupText As String
'    Dim cleanupErrNumber As Long
'    Dim cleanupErrDescription As String
'    Dim tableStillExists As Boolean
'    Dim fieldStillExists As Boolean
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If Not db Is Nothing And preStateCaptured Then
'        Err.Clear
'        db.TableDefs.Refresh
'        cleanupErrNumber = Err.Number
'        cleanupErrDescription = Err.Description
'        If cleanupErrNumber <> 0 Then
'            cleanupText = cleanupText & vbCrLf & "Cleanup table inspection error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription
'        Else
'            Err.Clear
'            tableStillExists = SharedDao_DatabaseTableExists(db, tableName)
'            cleanupErrNumber = Err.Number
'            cleanupErrDescription = Err.Description
'            If cleanupErrNumber <> 0 Then
'                cleanupText = cleanupText & vbCrLf & "Cleanup table-existence error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription
'            ElseIf tableExistedBefore = False Then
'                If tableStillExists Then
'                    Err.Clear
'                    db.TableDefs.Delete tableName
'                    db.TableDefs.Refresh
'                    cleanupErrNumber = Err.Number
'                    cleanupErrDescription = Err.Description
'                    If cleanupErrNumber <> 0 Then
'                        cleanupText = cleanupText & vbCrLf & "Created-table cleanup error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription
'                    Else
'                        Err.Clear
'                        tableStillExists = SharedDao_DatabaseTableExists(db, tableName)
'                        cleanupErrNumber = Err.Number
'                        cleanupErrDescription = Err.Description
'                        If cleanupErrNumber <> 0 Then
'                            cleanupText = cleanupText & vbCrLf & "Created-table cleanup verification error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription
'                        ElseIf tableStillExists Then
'                            cleanupText = cleanupText & vbCrLf & "Created-table cleanup failed verification: the table still exists."
'                        Else
'                            tableCreated = False
'                            cleanupText = cleanupText & vbCrLf & "Created table was removed and verified."
'                        End If
'                    End If
'                Else
'                    tableCreated = False
'                    cleanupText = cleanupText & vbCrLf & "Created table was not present during cleanup."
'                End If
'            ElseIf tableStillExists Then
'                Set tdf = db.TableDefs(tableName)
'                If Len(idFieldName) > 0 And idFieldExistedBefore = False Then
'                    Err.Clear
'                    tdf.fields.Refresh
'                    fieldStillExists = SharedDao_TableDefFieldExists(tdf, idFieldName)
'                    cleanupErrNumber = Err.Number
'                    cleanupErrDescription = Err.Description
'                    If cleanupErrNumber <> 0 Then
'                        cleanupText = cleanupText & vbCrLf & "Created-ID-field inspection error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription
'                    ElseIf fieldStillExists Then
'                        Err.Clear
'                        tdf.fields.Delete idFieldName
'                        tdf.fields.Refresh
'                        cleanupErrNumber = Err.Number
'                        cleanupErrDescription = Err.Description
'                        If cleanupErrNumber <> 0 Then
'                            cleanupText = cleanupText & vbCrLf & "Created-ID-field cleanup error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription
'                        Else
'                            fieldStillExists = SharedDao_TableDefFieldExists(tdf, idFieldName)
'                            cleanupErrNumber = Err.Number
'                            cleanupErrDescription = Err.Description
'                            If cleanupErrNumber <> 0 Then
'                                cleanupText = cleanupText & vbCrLf & "Created-ID-field cleanup verification error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription
'                            ElseIf fieldStillExists Then
'                                cleanupText = cleanupText & vbCrLf & "Created-ID-field cleanup failed verification: the field still exists."
'                            Else
'                                cleanupText = cleanupText & vbCrLf & "Created ID field was removed and verified."
'                            End If
'                        End If
'                    End If
'                End If
'                If descriptionStateCaptured Then
'                    Err.Clear
'                    SharedDao_RestorePropertyState tdf, "Description", dbText, descriptionExistedBefore, descriptionValueBefore
'                    cleanupErrNumber = Err.Number
'                    cleanupErrDescription = Err.Description
'                    If cleanupErrNumber <> 0 Then
'                        cleanupText = cleanupText & vbCrLf & "Table-description restore error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription
'                    Else
'                        cleanupText = cleanupText & vbCrLf & "Table description was restored and verified."
'                    End If
'                End If
'            Else
'                cleanupText = cleanupText & vbCrLf & "Pre-existing table is missing during cleanup; its prior state could not be restored."
'            End If
'        End If
'    ElseIf Not db Is Nothing Then
'        cleanupText = cleanupText & vbCrLf & "No schema cleanup was attempted because the pre-operation table state was not captured conclusively."
'    End If
'    Set fld = Nothing
'    Set tdf = Nothing
'    On Error GoTo 0
'    Err.Raise savedErrNumber, "SharedDao_EnsureTableExists", "Unable to create or verify table." & vbCrLf & "Table: " & tableName & vbCrLf & "Requested ID field: " & idFieldName & vbCrLf & "Requested AutoNumber: " & CStr(createAutoNumberID) & vbCrLf & "Database: " & SharedDao_SafeDatabaseName(db) & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription & cleanupText
'End Function
'Public Function SharedDao_EnsureField(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String, ByVal fieldType As DAO.DataTypeEnum, Optional ByVal fieldSize As Long = 0, Optional ByVal autoNumber As Boolean = False, Optional ByVal allowZeroLength As Boolean = False, Optional ByVal descriptionText As String = vbNullString, Optional ByVal allowLargerTextSize As Boolean = False) As Boolean
'    'Public shared API for callers that need to create or verify one DAO field without duplicating schema-repair logic.
'    On Error GoTo ErrHandler
'    SharedDao_EnsureField = SharedDao_EnsureFieldCore(db, tableName, fieldName, fieldType, fieldSize, autoNumber, allowZeroLength, descriptionText, allowLargerTextSize, "SharedDao_EnsureField", vbObjectError + 7401, vbObjectError + 7402)
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedDao_EnsureField", "Unable to create or verify field." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName & vbCrLf & "Expected: " & SharedDao_FieldDefinitionText(fieldType, fieldSize, autoNumber) & vbCrLf & "Database: " & SharedDao_SafeDatabaseName(db) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'
'Public Function SharedDao_EnsureSingleFieldIndex(ByVal db As DAO.Database, ByVal tableName As String, ByVal indexName As String, ByVal fieldName As String, ByVal uniqueIndex As Boolean, ByVal primaryIndex As Boolean, Optional ByVal replaceIfDifferent As Boolean = False) As Boolean
'    'Intended use: Creates or repairs one named single-field index and verifies the resulting definition.
'    SharedDao_EnsureSingleFieldIndex = SharedDao_EnsureCompositeIndex(db, tableName, indexName, Array(fieldName), uniqueIndex, primaryIndex, replaceIfDifferent)
'End Function
'
'Public Function SharedDao_EnsureCompositeIndex(ByVal db As DAO.Database, ByVal tableName As String, ByVal indexName As String, ByVal fieldNames As Variant, ByVal uniqueIndex As Boolean, ByVal primaryIndex As Boolean, Optional ByVal replaceIfDifferent As Boolean = False) As Boolean
'    'Ensures an exact composite index and restores the complete prior definition if replacement fails.
'    On Error GoTo ErrHandler
'    Dim tdf As DAO.tableDef
'    Dim expectedIdx As DAO.index
'    Dim expectedFld As DAO.Field
'    Dim existingIdx As DAO.index
'    Dim expectedDefinition As Object
'    Dim originalDefinition As Object
'    Dim normalizedFieldNames() As String
'    Dim i As Long
'    Dim requestedFieldName As String
'    Dim existingIndexCaptured As Boolean
'    Dim existingIndexDeleted As Boolean
'    Dim newIndexAppendAttempted As Boolean
'    Dim newIndexAppended As Boolean
'    If db Is Nothing Then Err.Raise vbObjectError + 7713, "SharedDao_EnsureCompositeIndex", "The DAO.Database argument is Nothing."
'    indexName = Trim$(indexName)
'    If Len(tableName) = 0 Then Err.Raise vbObjectError + 8056, "SharedDao_EnsureCompositeIndex", "The table name is blank."
'    If Len(indexName) = 0 Then Err.Raise vbObjectError + 8057, "SharedDao_EnsureCompositeIndex", "The index name is blank."
'    If SharedDao_DatabaseTableExists(db, tableName) = False Then Err.Raise vbObjectError + 7954, "SharedDao_EnsureCompositeIndex", "Table is missing before index repair. Table: " & tableName & "; Index: " & indexName
'    If Not IsArray(fieldNames) Then Err.Raise vbObjectError + 7714, "SharedDao_EnsureCompositeIndex", "The requested index field list is not an array."
'    If UBound(fieldNames) < LBound(fieldNames) Then Err.Raise vbObjectError + 7715, "SharedDao_EnsureCompositeIndex", "The requested index field list is empty."
'    ReDim normalizedFieldNames(LBound(fieldNames) To UBound(fieldNames))
'    Set tdf = db.TableDefs(tableName)
'    tdf.fields.Refresh
'    Set expectedIdx = tdf.CreateIndex(indexName)
'    expectedIdx.Unique = uniqueIndex
'    expectedIdx.Primary = primaryIndex
'    For i = LBound(fieldNames) To UBound(fieldNames)
'        requestedFieldName = Trim$(CStr(fieldNames(i)))
'        If Len(requestedFieldName) = 0 Then Err.Raise vbObjectError + 7716, "SharedDao_EnsureCompositeIndex", "A requested index field name is blank."
'        If SharedDao_TableDefFieldExists(tdf, requestedFieldName) = False Then Err.Raise vbObjectError + 7717, "SharedDao_EnsureCompositeIndex", "A requested index field is missing." & vbCrLf & "Table: " & tableName & vbCrLf & "Index: " & indexName & vbCrLf & "Field: " & requestedFieldName
'        normalizedFieldNames(i) = requestedFieldName
'        Set expectedFld = expectedIdx.CreateField(requestedFieldName)
'        expectedFld.attributes = 0
'        expectedIdx.fields.Append expectedFld
'    Next i
'    Set expectedDefinition = SharedDao_CaptureIndexDefinition(expectedIdx)
'    tdf.Indexes.Refresh
'    If SharedDao_TableDefIndexExists(tdf, indexName) Then
'        Set existingIdx = tdf.Indexes(indexName)
'        If SharedDao_IndexDefinitionMatches(existingIdx, expectedDefinition) Then
'            SharedDao_EnsureCompositeIndex = True
'            Exit Function
'        End If
'        If replaceIfDifferent = False Then Err.Raise vbObjectError + 7955, "SharedDao_EnsureCompositeIndex", "Existing index has an incompatible definition." & vbCrLf & "Table: " & tableName & vbCrLf & "Index: " & indexName & vbCrLf & "Expected fields: " & SharedDao_JoinVariantArray(normalizedFieldNames, ", ")
'        Set originalDefinition = SharedDao_CaptureIndexDefinition(existingIdx)
'        existingIndexCaptured = True
'        tdf.Indexes.Delete indexName
'        tdf.Indexes.Refresh
'        If SharedDao_TableDefIndexExists(tdf, indexName) Then Err.Raise vbObjectError + 8002, "SharedDao_EnsureCompositeIndex", "The incompatible index still exists after deletion." & vbCrLf & "Table: " & tableName & vbCrLf & "Index: " & indexName
'        existingIndexDeleted = True
'    End If
'    newIndexAppendAttempted = True
'    SharedDao_AppendIndexDefinition tdf, expectedDefinition
'    newIndexAppended = True
'    tdf.Indexes.Refresh
'    If SharedDao_TableDefIndexExists(tdf, indexName) = False Then Err.Raise vbObjectError + 7718, "SharedDao_EnsureCompositeIndex", "The index was not present after append."
'    If SharedDao_IndexDefinitionMatches(tdf.Indexes(indexName), expectedDefinition) = False Then Err.Raise vbObjectError + 7719, "SharedDao_EnsureCompositeIndex", "The appended index did not match the exact requested definition."
'    SharedDao_EnsureCompositeIndex = True
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim restoreErrNumber As Long
'    Dim restoreErrDescription As String
'    Dim cleanupErrNumber As Long
'    Dim cleanupErrDescription As String
'    Dim verificationErrNumber As Long
'    Dim verificationErrDescription As String
'    Dim indexStillExists As Boolean
'    Dim recoveryText As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    If existingIndexCaptured Then
'        On Error Resume Next
'        Err.Clear
'        SharedDao_RestoreIndexDefinitionExact tdf, originalDefinition
'        restoreErrNumber = Err.Number
'        restoreErrDescription = Err.Description
'        On Error GoTo 0
'        If restoreErrNumber <> 0 Then
'            recoveryText = vbCrLf & "Exact original-index restoration error " & CStr(restoreErrNumber) & ": " & restoreErrDescription
'        Else
'            recoveryText = vbCrLf & "The captured original index definition was restored and verified."
'        End If
'    ElseIf newIndexAppendAttempted And Not tdf Is Nothing Then
'        On Error Resume Next
'        Err.Clear
'        tdf.Indexes.Refresh
'        cleanupErrNumber = Err.Number
'        cleanupErrDescription = Err.Description
'        If cleanupErrNumber = 0 Then
'            Err.Clear
'            indexStillExists = SharedDao_TableDefIndexExists(tdf, indexName)
'            cleanupErrNumber = Err.Number
'            cleanupErrDescription = Err.Description
'        End If
'        If cleanupErrNumber = 0 And indexStillExists Then
'            Err.Clear
'            tdf.Indexes.Delete indexName
'            tdf.Indexes.Refresh
'            cleanupErrNumber = Err.Number
'            cleanupErrDescription = Err.Description
'        End If
'        If cleanupErrNumber = 0 Then
'            Err.Clear
'            indexStillExists = SharedDao_TableDefIndexExists(tdf, indexName)
'            verificationErrNumber = Err.Number
'            verificationErrDescription = Err.Description
'        End If
'        On Error GoTo 0
'        If cleanupErrNumber <> 0 Then
'            recoveryText = vbCrLf & "Partial-index cleanup error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription
'        ElseIf verificationErrNumber <> 0 Then
'            recoveryText = vbCrLf & "Partial-index cleanup verification error " & CStr(verificationErrNumber) & ": " & verificationErrDescription
'        ElseIf indexStillExists Then
'            recoveryText = vbCrLf & "Partial-index cleanup failed verification: the index still exists."
'        Else
'            recoveryText = vbCrLf & "Any partially appended new index was removed and verified absent."
'        End If
'    Else
'        recoveryText = vbCrLf & "No index cleanup was attempted because this invocation did not begin a persisted index mutation."
'    End If
'    recoveryText = recoveryText & vbCrLf & "Mutation state: existing captured=" & CStr(existingIndexCaptured) & "; existing deleted=" & CStr(existingIndexDeleted) & "; new append attempted=" & CStr(newIndexAppendAttempted) & "; new appended=" & CStr(newIndexAppended)
'    Err.Raise savedErrNumber, "SharedDao_EnsureCompositeIndex", "Unable to create or verify index." & vbCrLf & "Table: " & tableName & vbCrLf & "Index: " & indexName & vbCrLf & "Database: " & SharedDao_SafeDatabaseName(db) & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription & recoveryText
'End Function
'Public Function SharedDao_EnsureLookupField(ByVal tdf As DAO.tableDef, ByVal fieldName As String, ByVal lookupTableName As String, ByVal keyFieldName As String, ByVal displayFieldName As String, ByVal columnWidthsText As String, Optional ByVal orderByText As String = "[SortOrder]") As Boolean
'    'Intended use: Atomically creates or repairs Access lookup-display properties for an existing DAO field and restores the complete prior property set on failure.
'    On Error GoTo ErrHandler
'    Dim fld As DAO.Field
'    Dim rowSourceText As String
'    Dim changed As Boolean
'    Dim mutationStarted As Boolean
'    Dim propertyNames(0 To 6) As String
'    Dim propertyTypes(0 To 6) As DAO.DataTypeEnum
'    Dim requestedValues(0 To 6) As Variant
'    Dim originalExists(0 To 6) As Boolean
'    Dim originalValues(0 To 6) As Variant
'    Dim originalTypes(0 To 6) As DAO.DataTypeEnum
'    Dim failedPropertyName As String
'    Dim i As Long
'    If tdf Is Nothing Then Err.Raise vbObjectError + 7956, "SharedDao_EnsureLookupField", "The DAO.TableDef argument is Nothing while ensuring lookup metadata."
'    tdf.fields.Refresh
'    If SharedDao_TableDefFieldExists(tdf, fieldName) = False Then Err.Raise vbObjectError + 7406, "SharedDao_EnsureLookupField", "The target field is missing while ensuring lookup metadata." & vbCrLf & "Table: " & SharedDao_SafeTableDefName(tdf) & vbCrLf & "Field: " & fieldName
'    rowSourceText = "SELECT " & SharedDao_BracketName(keyFieldName) & ", " & SharedDao_BracketName(displayFieldName) & " FROM " & SharedDao_BracketName(lookupTableName)
'    If Len(Trim$(orderByText)) > 0 Then rowSourceText = rowSourceText & " ORDER BY " & orderByText & ", " & SharedDao_BracketName(keyFieldName)
'    rowSourceText = rowSourceText & ";"
'    Set fld = tdf.fields(fieldName)
'    propertyNames(0) = "DisplayControl": propertyTypes(0) = dbInteger: requestedValues(0) = 111
'    propertyNames(1) = "RowSourceType": propertyTypes(1) = dbText: requestedValues(1) = "Table/Query"
'    propertyNames(2) = "RowSource": propertyTypes(2) = dbText: requestedValues(2) = rowSourceText
'    propertyNames(3) = "BoundColumn": propertyTypes(3) = dbInteger: requestedValues(3) = 1
'    propertyNames(4) = "ColumnCount": propertyTypes(4) = dbInteger: requestedValues(4) = 2
'    propertyNames(5) = "ColumnWidths": propertyTypes(5) = dbText: requestedValues(5) = columnWidthsText
'    propertyNames(6) = "LimitToList": propertyTypes(6) = dbBoolean: requestedValues(6) = True
'    For i = 0 To 6
'        SharedDao_GetPropertyState fld, propertyNames(i), originalExists(i), originalValues(i)
'        If originalExists(i) Then originalTypes(i) = fld.Properties(propertyNames(i)).Type Else originalTypes(i) = propertyTypes(i)
'    Next i
'    mutationStarted = True
'    For i = 0 To 6
'        failedPropertyName = propertyNames(i)
'        changed = SharedDao_EnsurePropertyValue(fld, propertyNames(i), propertyTypes(i), requestedValues(i), "SharedDao_EnsureLookupField") Or changed
'    Next i
'    SharedDao_EnsureLookupField = changed
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim restoreErrNumber As Long
'    Dim restoreErrDescription As String
'    Dim rollbackText As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    If mutationStarted And Not fld Is Nothing Then
'        For i = 6 To 0 Step -1
'            On Error Resume Next
'            Err.Clear
'            SharedDao_RestorePropertyState fld, propertyNames(i), originalTypes(i), originalExists(i), originalValues(i)
'            restoreErrNumber = Err.Number
'            restoreErrDescription = Err.Description
'            On Error GoTo 0
'            If restoreErrNumber <> 0 Then rollbackText = SharedText_AppendNonBlankLine(rollbackText, propertyNames(i) & " restore error " & CStr(restoreErrNumber) & ": " & restoreErrDescription)
'        Next i
'    End If
'    If Len(rollbackText) = 0 And mutationStarted Then rollbackText = "The complete original lookup-property set was restored and verified."
'    Err.Raise savedErrNumber, "SharedDao_EnsureLookupField", "Unable to create or verify lookup field metadata." & vbCrLf & "Table: " & SharedDao_SafeTableDefName(tdf) & vbCrLf & "Field: " & fieldName & vbCrLf & "LookupTable: " & lookupTableName & vbCrLf & "Failed property: " & failedPropertyName & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription & IIf(Len(rollbackText) > 0, vbCrLf & "Rollback: " & rollbackText, vbNullString)
'End Function
'
'Public Function SharedDao_FieldDefinitionText(ByVal fieldType As DAO.DataTypeEnum, Optional ByVal fieldSize As Long = 0, Optional ByVal autoNumber As Boolean = False) As String
'    'Intended use: Returns a readable expected DAO field definition for diagnostics and mismatch reports.
'    If autoNumber Then
'        SharedDao_FieldDefinitionText = "AutoNumber Long"
'    ElseIf fieldType = dbText Then
'        SharedDao_FieldDefinitionText = "Text(" & CStr(fieldSize) & ")"
'    Else
'        SharedDao_FieldDefinitionText = SharedDao_DataTypeText(fieldType)
'    End If
'End Function
'
'Public Function SharedDao_ExistingFieldDefinitionText(ByVal fld As DAO.Field) As String
'    'Intended use: Returns a readable description of an existing DAO field for schema diagnostics.
'    Dim autoNumberText As String
'    If (fld.attributes And dbAutoIncrField) <> 0 Then
'        autoNumberText = "; AutoNumber=True"
'    Else
'        autoNumberText = "; AutoNumber=False"
'    End If
'    If fld.Type = dbText Then
'        SharedDao_ExistingFieldDefinitionText = "Text(" & CStr(fld.Size) & ")" & autoNumberText
'    Else
'        SharedDao_ExistingFieldDefinitionText = SharedDao_DataTypeText(fld.Type) & autoNumberText
'    End If
'End Function
'
'Public Function SharedDao_DataTypeText(ByVal fieldType As DAO.DataTypeEnum) As String
'    'Intended use: Returns the display name for a DAO data-type constant.
'    Select Case fieldType
'        Case dbBoolean
'            SharedDao_DataTypeText = "Boolean"
'        Case dbByte
'            SharedDao_DataTypeText = "Byte"
'        Case dbInteger
'            SharedDao_DataTypeText = "Integer"
'        Case dbLong
'            SharedDao_DataTypeText = "Long"
'        Case dbCurrency
'            SharedDao_DataTypeText = "Currency"
'        Case dbSingle
'            SharedDao_DataTypeText = "Single"
'        Case dbDouble
'            SharedDao_DataTypeText = "Double"
'        Case dbDate
'            SharedDao_DataTypeText = "Date"
'        Case dbText
'            SharedDao_DataTypeText = "Text"
'        Case dbMemo
'            SharedDao_DataTypeText = "Memo"
'        Case Else
'            SharedDao_DataTypeText = "DAO type " & CStr(fieldType)
'    End Select
'End Function
'
'Public Function SharedDao_FieldExistsInSource(ByVal sourceName As String, ByVal fieldName As String) As Boolean
'    'Returns whether a field is exposed by a table or query in CurrentDb.
'    'Use SharedDao_DatabaseFieldExistsInSource when the database is already known.
'    Dim db As DAO.Database
'    Set db = currentDb
'    SharedDao_FieldExistsInSource = SharedDao_DatabaseFieldExistsInSource(db, sourceName, fieldName)
'    Set db = Nothing
'End Function
'
'Private Function SharedDao_SourceFieldMap(ByVal db As DAO.Database, ByVal sourceName As String, ByVal callerName As String) As Object
'    'Builds one case-insensitive field-name map for a table or saved query so callers can reuse a single metadata inspection.
'    On Error GoTo ErrHandler
'    Dim fieldMap As Object
'    Dim tdf As DAO.tableDef
'    Dim qdf As DAO.queryDef
'    Dim fld As DAO.Field
'    If db Is Nothing Then Err.Raise 91, callerName, "The supplied DAO.Database is Nothing."
'    Set fieldMap = CreateObject("Scripting.Dictionary")
'    fieldMap.CompareMode = vbTextCompare
'    For Each tdf In db.TableDefs
'        If StrComp(tdf.Name, sourceName, vbTextCompare) = 0 Then
'            tdf.fields.Refresh
'            For Each fld In tdf.fields
'                If fieldMap.Exists(fld.Name) = False Then fieldMap.Add fld.Name, fld.Name
'            Next fld
'            Set SharedDao_SourceFieldMap = fieldMap
'            Exit Function
'        End If
'    Next tdf
'    For Each qdf In db.QueryDefs
'        If StrComp(qdf.Name, sourceName, vbTextCompare) = 0 Then
'            qdf.fields.Refresh
'            For Each fld In qdf.fields
'                If fieldMap.Exists(fld.Name) = False Then fieldMap.Add fld.Name, fld.Name
'            Next fld
'            Set SharedDao_SourceFieldMap = fieldMap
'            Exit Function
'        End If
'    Next qdf
'    Err.Raise 3265, callerName, "The source table or query was not found: " & sourceName
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, callerName, "Unable to inspect source fields." & vbCrLf & "Database: " & SharedDao_SafeDatabaseName(db) & vbCrLf & "Source: " & sourceName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Original description: " & savedErrDescription
'End Function
'
'Private Function SharedDao_FieldNameCollection(ByVal fieldValues As Variant) As Collection
'    'Materializes caller-provided field names once while preserving order and duplicates.
'    Dim fieldNames As Collection
'    Dim item As Variant
'    Set fieldNames = New Collection
'    For Each item In fieldValues
'        fieldNames.Add CStr(item)
'    Next item
'    Set SharedDao_FieldNameCollection = fieldNames
'End Function
'
'Private Function SharedDao_ExistingFieldsFromMap(ByVal fieldMap As Object, ByVal preferredFields As Variant) As Collection
'    'Returns requested fields present in an already-inspected source-field map.
'    Dim result As Collection
'    Dim item As Variant
'    Set result = New Collection
'    For Each item In preferredFields
'        If fieldMap.Exists(CStr(item)) Then result.Add CStr(item)
'    Next item
'    Set SharedDao_ExistingFieldsFromMap = result
'End Function
'
'Private Function SharedDao_OrderByTextFromMap(ByVal fieldMap As Object, ByVal preferredFields As Variant) As String
'    'Builds an ORDER BY list from an already-inspected source-field map.
'    Dim item As Variant
'    Dim resultText As String
'    For Each item In preferredFields
'        If fieldMap.Exists(CStr(item)) Then
'            If Len(resultText) > 0 Then resultText = resultText & ", "
'            resultText = resultText & SharedDao_BracketName(CStr(item))
'        End If
'    Next item
'    SharedDao_OrderByTextFromMap = resultText
'End Function
'
'Private Function SharedDao_WhereTextHasUnterminatedBracket(ByVal whereText As String) As Boolean
'    'Returns True when a bracketed identifier is not closed, honoring escaped closing brackets.
'    Dim startPos As Long
'    Dim scanPos As Long
'    Dim bracketClosed As Boolean
'    startPos = InStr(1, whereText, "[", vbBinaryCompare)
'    Do While startPos > 0
'        scanPos = startPos + 1
'        bracketClosed = False
'        Do While scanPos <= Len(whereText)
'            If Mid$(whereText, scanPos, 1) = "]" Then
'                If scanPos < Len(whereText) And Mid$(whereText, scanPos + 1, 1) = "]" Then
'                    scanPos = scanPos + 2
'                Else
'                    bracketClosed = True
'                    Exit Do
'                End If
'            Else
'                scanPos = scanPos + 1
'            End If
'        Loop
'        If bracketClosed = False Then
'            SharedDao_WhereTextHasUnterminatedBracket = True
'            Exit Function
'        End If
'        startPos = InStr(scanPos + 1, whereText, "[", vbBinaryCompare)
'    Loop
'End Function
'
'Private Function SharedDao_WhereTextReferencesMissingFieldInMap(ByVal fieldMap As Object, ByVal whereText As String) As Boolean
'    'Checks bracketed WHERE identifiers against an already-inspected source-field map.
'    Dim startPos As Long
'    Dim scanPos As Long
'    Dim fieldName As String
'    Dim bracketClosed As Boolean
'    startPos = InStr(1, whereText, "[", vbBinaryCompare)
'    Do While startPos > 0
'        scanPos = startPos + 1
'        fieldName = vbNullString
'        bracketClosed = False
'        Do While scanPos <= Len(whereText)
'            If Mid$(whereText, scanPos, 1) = "]" Then
'                If scanPos < Len(whereText) And Mid$(whereText, scanPos + 1, 1) = "]" Then
'                    fieldName = fieldName & "]"
'                    scanPos = scanPos + 2
'                Else
'                    bracketClosed = True
'                    Exit Do
'                End If
'            Else
'                fieldName = fieldName & Mid$(whereText, scanPos, 1)
'                scanPos = scanPos + 1
'            End If
'        Loop
'        If bracketClosed = False Then
'            SharedDao_WhereTextReferencesMissingFieldInMap = True
'            Exit Function
'        End If
'        If Len(fieldName) > 0 Then
'            If fieldMap.Exists(fieldName) = False Then
'                SharedDao_WhereTextReferencesMissingFieldInMap = True
'                Exit Function
'            End If
'        End If
'        startPos = InStr(scanPos + 1, whereText, "[", vbBinaryCompare)
'    Loop
'End Function
'
'Public Function SharedDao_DatabaseFieldExistsInSource(ByVal db As DAO.Database, ByVal sourceName As String, ByVal fieldName As String) As Boolean
'    'Intended use: Returns whether a field is exposed by a table or saved query in the supplied database.
'    'A confirmed missing field returns False; source-resolution and collection-inspection failures are preserved and raised.
'    On Error GoTo ErrHandler
'    Dim fieldMap As Object
'    If db Is Nothing Then Err.Raise 91, "SharedDao_DatabaseFieldExistsInSource", "The supplied DAO.Database is Nothing."
'    Set fieldMap = SharedDao_SourceFieldMap(db, sourceName, "SharedDao_DatabaseFieldExistsInSource")
'    SharedDao_DatabaseFieldExistsInSource = fieldMap.Exists(fieldName)
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedDao_DatabaseFieldExistsInSource", "Unable to inspect a source field." & vbCrLf & "Database: " & SharedDao_SafeDatabaseName(db) & vbCrLf & "Source: " & sourceName & vbCrLf & "Field: " & fieldName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Original description: " & savedErrDescription
'End Function
'
'Public Function SharedDao_ExistingFields(ByVal sourceName As String, ByVal preferredFields As Variant) As Collection
'    'Returns preferred fields that exist in a CurrentDb table or query.
'    'Use SharedDao_DatabaseExistingFields when the database is already known.
'    Dim db As DAO.Database
'    Set db = currentDb
'    Set SharedDao_ExistingFields = SharedDao_DatabaseExistingFields(db, sourceName, preferredFields)
'    Set db = Nothing
'End Function
'
'Public Function SharedDao_DatabaseExistingFields(ByVal db As DAO.Database, ByVal sourceName As String, ByVal preferredFields As Variant) As Collection
'    'Intended use: Returns preferred fields that exist in a table or query after one source metadata inspection.
'    Dim requestedFields As Collection
'    Dim fieldMap As Object
'    Set requestedFields = SharedDao_FieldNameCollection(preferredFields)
'    If requestedFields.count = 0 Then
'        Set SharedDao_DatabaseExistingFields = requestedFields
'        Exit Function
'    End If
'    Set fieldMap = SharedDao_SourceFieldMap(db, sourceName, "SharedDao_DatabaseExistingFields")
'    Set SharedDao_DatabaseExistingFields = SharedDao_ExistingFieldsFromMap(fieldMap, requestedFields)
'End Function
'
'Public Function SharedDao_ExistingOrderByText(ByVal sourceName As String, ByVal preferredFields As Variant) As String
'    'Builds an ORDER BY field list from fields exposed by a CurrentDb table or query.
'    'Use SharedDao_DatabaseExistingOrderByText when the database is already known.
'    Dim db As DAO.Database
'    Set db = currentDb
'    SharedDao_ExistingOrderByText = SharedDao_DatabaseExistingOrderByText(db, sourceName, preferredFields)
'    Set db = Nothing
'End Function
'
'Public Function SharedDao_DatabaseExistingOrderByText(ByVal db As DAO.Database, ByVal sourceName As String, ByVal preferredFields As Variant) As String
'    'Intended use: Builds an ORDER BY field list after one source metadata inspection.
'    Dim requestedFields As Collection
'    Dim fieldMap As Object
'    Set requestedFields = SharedDao_FieldNameCollection(preferredFields)
'    If requestedFields.count = 0 Then Exit Function
'    Set fieldMap = SharedDao_SourceFieldMap(db, sourceName, "SharedDao_DatabaseExistingOrderByText")
'    SharedDao_DatabaseExistingOrderByText = SharedDao_OrderByTextFromMap(fieldMap, requestedFields)
'End Function
'
'Public Function SharedDao_BuildSelectSql(ByVal sourceName As String, ByVal preferredFields As Variant, ByVal whereText As String, ByVal orderFields As Variant, Optional ByRef omittedWhereReason As String = vbNullString) As String
'    'Builds a defensive SELECT statement against a CurrentDb table or query.
'    'Use SharedDao_DatabaseBuildSelectSql when the database is already known.
'    Dim db As DAO.Database
'    Set db = currentDb
'    SharedDao_BuildSelectSql = SharedDao_DatabaseBuildSelectSql(db, sourceName, preferredFields, whereText, orderFields, omittedWhereReason)
'    Set db = Nothing
'End Function
'
'Public Function SharedDao_DatabaseBuildSelectSql(ByVal db As DAO.Database, ByVal sourceName As String, ByVal preferredFields As Variant, ByVal whereText As String, ByVal orderFields As Variant, Optional ByRef omittedWhereReason As String = vbNullString) As String
'    'Intended use: Builds a defensive SELECT statement from one reusable source-field inspection.
'    'A requested WHERE clause is never removed to produce broader SQL; missing bracketed fields return blank SQL and a failure reason.
'    On Error GoTo ErrHandler
'    Dim fieldMap As Object
'    Dim requestedSelectFields As Collection
'    Dim requestedOrderFields As Collection
'    Dim fields As Collection
'    Dim fieldName As Variant
'    Dim selectText As String
'    Dim orderText As String
'    Dim cleanWhereText As String
'    omittedWhereReason = vbNullString
'    If db Is Nothing Then Err.Raise 91, "SharedDao_DatabaseBuildSelectSql", "The supplied DAO.Database is Nothing."
'    Set requestedSelectFields = SharedDao_FieldNameCollection(preferredFields)
'    If requestedSelectFields.count = 0 Then
'        omittedWhereReason = "SELECT not built because source [" & sourceName & "] exposes none of the requested SELECT fields."
'        Exit Function
'    End If
'    Set fieldMap = SharedDao_SourceFieldMap(db, sourceName, "SharedDao_DatabaseBuildSelectSql")
'    Set fields = SharedDao_ExistingFieldsFromMap(fieldMap, requestedSelectFields)
'    If fields.count = 0 Then
'        omittedWhereReason = "SELECT not built because source [" & sourceName & "] exposes none of the requested SELECT fields."
'        Exit Function
'    End If
'    For Each fieldName In fields
'        If Len(selectText) > 0 Then selectText = selectText & ", "
'        selectText = selectText & SharedDao_BracketName(CStr(fieldName))
'    Next fieldName
'    Set requestedOrderFields = SharedDao_FieldNameCollection(orderFields)
'    orderText = SharedDao_OrderByTextFromMap(fieldMap, requestedOrderFields)
'    cleanWhereText = Trim$(whereText)
'    If Len(cleanWhereText) > 0 Then
'        If SharedDao_WhereTextReferencesMissingFieldInMap(fieldMap, cleanWhereText) Then
'            omittedWhereReason = "SELECT not built because source [" & sourceName & "] does not expose every bracketed WHERE field or the WHERE clause contains an unterminated bracketed identifier. WHERE: " & cleanWhereText
'            Exit Function
'        End If
'    End If
'    SharedDao_DatabaseBuildSelectSql = "SELECT " & selectText & " FROM " & SharedDao_BracketName(sourceName)
'    If Len(cleanWhereText) > 0 Then SharedDao_DatabaseBuildSelectSql = SharedDao_DatabaseBuildSelectSql & " WHERE " & cleanWhereText
'    If Len(orderText) > 0 Then SharedDao_DatabaseBuildSelectSql = SharedDao_DatabaseBuildSelectSql & " ORDER BY " & orderText
'    SharedDao_DatabaseBuildSelectSql = SharedDao_DatabaseBuildSelectSql & ";"
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    omittedWhereReason = "SELECT not built because source or field inspection failed. Error " & CStr(savedErrNumber) & ": " & savedErrDescription
'    Err.Raise savedErrNumber, "SharedDao_DatabaseBuildSelectSql", "Unable to build the defensive SELECT statement." & vbCrLf & "Database: " & SharedDao_SafeDatabaseName(db) & vbCrLf & "Source: " & sourceName & vbCrLf & "WHERE: " & whereText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Original description: " & savedErrDescription
'End Function
'
'Public Function SharedDao_WhereTextReferencesMissingField(ByVal sourceName As String, ByVal whereText As String) As Boolean
'    'Returns whether a bracketed WHERE field is missing from a CurrentDb table or query.
'    'Use SharedDao_DatabaseWhereTextReferencesMissingField when the database is already known.
'    Dim db As DAO.Database
'    Set db = currentDb
'    SharedDao_WhereTextReferencesMissingField = SharedDao_DatabaseWhereTextReferencesMissingField(db, sourceName, whereText)
'    Set db = Nothing
'End Function
'
'Public Function SharedDao_DatabaseWhereTextReferencesMissingField(ByVal db As DAO.Database, ByVal sourceName As String, ByVal whereText As String) As Boolean
'    'Intended use: Returns whether a bracketed WHERE identifier is missing after one source metadata inspection.
'    'An unterminated bracketed identifier is treated as invalid so callers fail closed rather than returning broader SQL.
'    Dim fieldMap As Object
'    If db Is Nothing Then Err.Raise 91, "SharedDao_DatabaseWhereTextReferencesMissingField", "The supplied DAO.Database is Nothing."
'    If InStr(1, whereText, "[", vbBinaryCompare) = 0 Then Exit Function
'    If SharedDao_WhereTextHasUnterminatedBracket(whereText) Then
'        SharedDao_DatabaseWhereTextReferencesMissingField = True
'        Exit Function
'    End If
'    Set fieldMap = SharedDao_SourceFieldMap(db, sourceName, "SharedDao_DatabaseWhereTextReferencesMissingField")
'    SharedDao_DatabaseWhereTextReferencesMissingField = SharedDao_WhereTextReferencesMissingFieldInMap(fieldMap, whereText)
'End Function
'
'Public Function SharedDao_AccessObjectExists(ByVal db As DAO.Database, ByVal objectType As AcObjectType, ByVal objectName As String) As Boolean
'    'Returns whether an Access object exists in the supplied database.
'    'Active-project collections are consulted only when the supplied database is the active project; technical inspection failures are preserved and raised.
'    On Error GoTo ErrHandler
'    Dim typeID As Long
'    Dim rs As DAO.Recordset
'    Dim sqlText As String
'    Dim obj As accessObject
'    If db Is Nothing Then Err.Raise 91, "SharedDao_AccessObjectExists", "The supplied DAO.Database is Nothing."
'    If Len(objectName) = 0 Then Exit Function
'    Select Case objectType
'        Case acTable
'            SharedDao_AccessObjectExists = SharedDao_DatabaseTableExists(db, objectName)
'            If SharedDao_AccessObjectExists Then Exit Function
'        Case acQuery
'            SharedDao_AccessObjectExists = SharedDao_DatabaseQueryExists(db, objectName)
'            If SharedDao_AccessObjectExists Then Exit Function
'    End Select
'    typeID = SharedDao_MSysObjectTypeForAccessObjectType(objectType)
'    If typeID = 0 Then Exit Function
'    sqlText = "SELECT TOP 1 [ID] FROM [MSysObjects] WHERE UCase([Name])=" & SharedDao_SqlQuote(UCase$(objectName))
'    sqlText = sqlText & " AND [Type]=" & CStr(typeID) & " ORDER BY [ID];"
'    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
'    If Not rs.EOF Then
'        SharedDao_AccessObjectExists = True
'        GoTo CleanExit
'    End If
'    rs.Close
'    Set rs = Nothing
'    If SharedDao_IsCurrentProjectDatabase(db) = False Then GoTo CleanExit
'    Select Case objectType
'        Case acForm
'            For Each obj In Application.CurrentProject.AllForms
'                If StrComp(obj.Name, objectName, vbTextCompare) = 0 Then
'                    SharedDao_AccessObjectExists = True
'                    Exit For
'                End If
'            Next obj
'        Case acReport
'            For Each obj In Application.CurrentProject.AllReports
'                If StrComp(obj.Name, objectName, vbTextCompare) = 0 Then
'                    SharedDao_AccessObjectExists = True
'                    Exit For
'                End If
'            Next obj
'        Case acMacro
'            For Each obj In Application.CurrentProject.AllMacros
'                If StrComp(obj.Name, objectName, vbTextCompare) = 0 Then
'                    SharedDao_AccessObjectExists = True
'                    Exit For
'                End If
'            Next obj
'    End Select
'CleanExit:
'    On Error Resume Next
'    If Not rs Is Nothing Then rs.Close
'    Set rs = Nothing
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If Not rs Is Nothing Then rs.Close
'    Set rs = Nothing
'    On Error GoTo 0
'    Err.Raise savedErrNumber, "SharedDao_AccessObjectExists", "Unable to inspect the requested Access object." & vbCrLf & "Database: " & SharedDao_SafeDatabaseName(db) & vbCrLf & "Object type: " & CStr(objectType) & vbCrLf & "Object name: " & objectName & vbCrLf & "SQL: " & sqlText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Original description: " & savedErrDescription
'End Function
'Private Function SharedDao_IsCurrentProjectDatabase(ByVal db As DAO.Database) As Boolean
'    If db Is Nothing Then Err.Raise 91, "SharedDao_IsCurrentProjectDatabase", "The supplied DAO.Database is Nothing."
'    SharedDao_IsCurrentProjectDatabase = (StrComp(Trim$(db.Name), Trim$(Application.CurrentProject.FullName), vbTextCompare) = 0)
'End Function
'
'Public Function SharedDao_MSysObjectTypeForAccessObjectType(ByVal objectType As AcObjectType) As Long
'    'Intended use: Maps an Access object type to the corresponding MSysObjects type code.
'    Select Case objectType
'        Case acTable
'            SharedDao_MSysObjectTypeForAccessObjectType = 1
'        Case acQuery
'            SharedDao_MSysObjectTypeForAccessObjectType = 5
'        Case acForm
'            SharedDao_MSysObjectTypeForAccessObjectType = -32768
'        Case acReport
'            SharedDao_MSysObjectTypeForAccessObjectType = -32764
'        Case acMacro
'            SharedDao_MSysObjectTypeForAccessObjectType = -32766
'        Case acModule
'            SharedDao_MSysObjectTypeForAccessObjectType = -32761
'        Case Else
'            SharedDao_MSysObjectTypeForAccessObjectType = 0
'    End Select
'End Function
'
'Public Function SharedDao_BracketName(ByVal objectName As String) As String
'    'Intended use: Returns an Access SQL identifier escaped and enclosed in brackets.
'    SharedDao_BracketName = "[" & Replace(objectName, "]", "]]") & "]"
'End Function
'
'Public Function SharedDao_SqlQuote(ByVal value As Variant) As String
'    'Intended use: Returns a text value quoted and escaped for Access SQL.
'    If IsNull(value) Then
'        SharedDao_SqlQuote = "Null"
'    Else
'        SharedDao_SqlQuote = "'" & Replace(CStr(value), "'", "''") & "'"
'    End If
'End Function
'
'Public Function SharedDao_JoinVariantArray(ByVal values As Variant, Optional ByVal delimiter As String = ", ") As String
'    'Intended use: Joins a Variant array into delimiter-separated diagnostic or SQL text.
'    Dim i As Long
'    Dim resultText As String
'    For i = LBound(values) To UBound(values)
'        If Len(resultText) > 0 Then resultText = resultText & delimiter
'        resultText = resultText & CStr(values(i))
'    Next i
'    SharedDao_JoinVariantArray = resultText
'End Function
'
'Public Function SharedDao_SafeDatabaseName(ByVal db As DAO.Database) As String
'    'Intended use: Returns a diagnostic database name without allowing a name-read failure to replace the primary error.
'    On Error GoTo ErrHandler
'    If db Is Nothing Then
'        SharedDao_SafeDatabaseName = "(database object is not set)"
'    Else
'        SharedDao_SafeDatabaseName = db.Name
'    End If
'    Exit Function
'ErrHandler:
'    SharedDao_SafeDatabaseName = "(database name unavailable)"
'End Function
'
'Public Function SharedDao_SafeTableDefName(ByVal tdf As DAO.tableDef) As String
'    'Intended use: Returns a diagnostic TableDef name without allowing a name-read failure to replace the primary error.
'    On Error GoTo ErrHandler
'    If tdf Is Nothing Then
'        SharedDao_SafeTableDefName = "(tabledef object is not set)"
'    Else
'        SharedDao_SafeTableDefName = tdf.Name
'    End If
'    Exit Function
'ErrHandler:
'    SharedDao_SafeTableDefName = "(tabledef name unavailable)"
'End Function
'
'Public Function SharedDao_DeleteRecordByID(ByVal tableName As String, ByVal recordID As Long, Optional ByVal idFieldName As String = "ID") As Boolean
'    'Intended use: Transactionally deletes one CurrentDb record and commits only when no more than one row matched.
'    On Error GoTo ErrHandler
'    Dim db As DAO.Database
'    Dim wrk As DAO.workspace
'    Dim sqlText As String
'    Dim affectedRows As Long
'    Dim transactionStarted As Boolean
'    idFieldName = Trim$(idFieldName)
'    If Len(tableName) = 0 Then Exit Function
'    If Len(idFieldName) = 0 Then Exit Function
'    If recordID <= 0 Then Exit Function
'    Set db = currentDb
'    If SharedDao_DatabaseTableExists(db, tableName) = False Then Exit Function
'    If SharedDao_TableFieldExists(db, tableName, idFieldName) = False Then Err.Raise vbObjectError + 7750, "SharedDao_DeleteRecordByID", "The identifier field does not exist in the target table."
'    Set wrk = DBEngine.Workspaces(0)
'    sqlText = "DELETE FROM " & SharedDao_BracketName(tableName) & " WHERE " & SharedDao_BracketName(idFieldName) & "=" & SharedDao_InvariantNumericLiteral(recordID) & ";"
'    wrk.BeginTrans
'    transactionStarted = True
'    db.Execute sqlText, dbFailOnError
'    affectedRows = db.RecordsAffected
'    If affectedRows > 1 Then Err.Raise vbObjectError + 7631, "SharedDao_DeleteRecordByID", "The delete affected more than one row; the identifier field is not unique for this value."
'    wrk.CommitTrans
'    transactionStarted = False
'    SharedDao_DeleteRecordByID = (affectedRows = 1)
'    Set wrk = Nothing
'    Set db = Nothing
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim rollbackErrNumber As Long
'    Dim rollbackErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    Err.Clear
'    If transactionStarted And Not wrk Is Nothing Then wrk.Rollback
'    rollbackErrNumber = Err.Number
'    rollbackErrDescription = Err.Description
'    Set wrk = Nothing
'    Set db = Nothing
'    On Error GoTo 0
'    If rollbackErrNumber <> 0 Then savedErrDescription = savedErrDescription & vbCrLf & "Rollback error " & CStr(rollbackErrNumber) & ": " & rollbackErrDescription
'    Err.Raise savedErrNumber, "SharedDao_DeleteRecordByID", "Unable to delete or verify the requested record." & vbCrLf & "Table: " & tableName & vbCrLf & "ID field: " & idFieldName & vbCrLf & "Record ID: " & CStr(recordID) & vbCrLf & "Affected rows before rollback: " & CStr(affectedRows) & vbCrLf & "SQL: " & sqlText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'Public Function SharedDao_AddRecord(ByVal tableName As String, ByVal fieldValues As Object, Optional ByVal idFieldName As String = "ID") As Long
'    'Transactionally adds one keyed record, closes the verified recordset before commit, and cannot report a committed insert as an ordinary failure.
'    On Error GoTo ErrHandler
'    Dim db As DAO.Database
'    Dim wrk As DAO.workspace
'    Dim rs As DAO.Recordset
'    Dim fieldName As Variant
'    Dim currentFieldName As String
'    Dim stageText As String
'    Dim transactionStarted As Boolean
'    idFieldName = Trim$(idFieldName)
'    If Len(tableName) = 0 Then Exit Function
'    If fieldValues Is Nothing Then Exit Function
'    stageText = "opening target table"
'    Set db = currentDb
'    If SharedDao_DatabaseTableExists(db, tableName) = False Then Err.Raise vbObjectError + 7751, "SharedDao_AddRecord", "The target table does not exist."
'    If Len(idFieldName) > 0 Then
'        If SharedDao_TableFieldExists(db, tableName, idFieldName) = False Then Err.Raise vbObjectError + 7752, "SharedDao_AddRecord", "The requested identifier field does not exist in the target table."
'    End If
'    Set wrk = DBEngine.Workspaces(0)
'    Set rs = db.OpenRecordset(tableName, dbOpenDynaset)
'    wrk.BeginTrans
'    transactionStarted = True
'    stageText = "starting new record"
'    rs.AddNew
'    For Each fieldName In fieldValues.keys
'        currentFieldName = CStr(fieldName)
'        stageText = "assigning field"
'        rs.fields(currentFieldName).value = fieldValues(fieldName)
'    Next fieldName
'    stageText = "updating record"
'    rs.Update
'    stageText = "reading inserted record"
'    rs.Bookmark = rs.LastModified
'    If Len(idFieldName) > 0 Then
'        If IsNull(rs.fields(idFieldName).value) Then Err.Raise vbObjectError + 7753, "SharedDao_AddRecord", "The inserted record identifier is Null."
'        SharedDao_AddRecord = CLng(rs.fields(idFieldName).value)
'        If SharedDao_AddRecord <= 0 Then Err.Raise vbObjectError + 7754, "SharedDao_AddRecord", "The inserted record identifier is not positive."
'    End If
'    stageText = "closing the verified recordset before commit"
'    rs.Close
'    Set rs = Nothing
'    stageText = "committing insert transaction"
'    wrk.CommitTrans
'    transactionStarted = False
'    Set wrk = Nothing
'    Set db = Nothing
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim cancelErrNumber As Long
'    Dim cancelErrDescription As String
'    Dim rollbackErrNumber As Long
'    Dim rollbackErrDescription As String
'    Dim closeErrNumber As Long
'    Dim closeErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If Not rs Is Nothing Then
'        Err.Clear
'        If rs.EditMode <> dbEditNone Then rs.CancelUpdate
'        cancelErrNumber = Err.Number
'        cancelErrDescription = Err.Description
'    End If
'    Err.Clear
'    If transactionStarted And Not wrk Is Nothing Then wrk.Rollback
'    rollbackErrNumber = Err.Number
'    rollbackErrDescription = Err.Description
'    If Not rs Is Nothing Then
'        Err.Clear
'        rs.Close
'        closeErrNumber = Err.Number
'        closeErrDescription = Err.Description
'    End If
'    Set rs = Nothing
'    Set wrk = Nothing
'    Set db = Nothing
'    On Error GoTo 0
'    If cancelErrNumber <> 0 Then savedErrDescription = savedErrDescription & vbCrLf & "CancelUpdate error " & CStr(cancelErrNumber) & ": " & cancelErrDescription
'    If rollbackErrNumber <> 0 Then savedErrDescription = savedErrDescription & vbCrLf & "Rollback error " & CStr(rollbackErrNumber) & ": " & rollbackErrDescription
'    If closeErrNumber <> 0 Then savedErrDescription = savedErrDescription & vbCrLf & "Recordset-close error " & CStr(closeErrNumber) & ": " & closeErrDescription
'    Err.Raise savedErrNumber, "SharedDao_AddRecord", "Could not add the DAO record." & vbCrLf & "Table: " & tableName & vbCrLf & "ID field: " & idFieldName & vbCrLf & "Stage: " & stageText & vbCrLf & "Field: " & currentFieldName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'Public Function SharedDao_CurrentDbQueryExists(ByVal queryName As String) As Boolean
'    'Returns True when a query exists in CurrentDb by case-insensitive name comparison.
'    'Technical inspection failures are raised so callers do not mistake an unreadable collection for a missing query.
'    Dim db As DAO.Database
'    Set db = currentDb
'    SharedDao_CurrentDbQueryExists = SharedDao_DatabaseQueryExists(db, queryName)
'    Set db = Nothing
'End Function
'
'Public Function SharedDao_DatabaseQueryExists(ByVal db As DAO.Database, ByVal queryName As String) As Boolean
'    'Returns True when a query exists in the supplied DAO.Database.
'    'Use when checking a database object other than CurrentDb or when case-insensitive matching is preferred.
'    Dim qdf As DAO.queryDef
'    For Each qdf In db.QueryDefs
'        If StrComp(qdf.Name, queryName, vbTextCompare) = 0 Then
'            SharedDao_DatabaseQueryExists = True
'            Exit Function
'        End If
'    Next qdf
'End Function
'
'Public Function SharedDao_SqlLiteralForFieldValue(ByVal v As Variant, ByVal fld As DAO.Field) As String
'    'Converts a value to the correct Access SQL literal for the supplied DAO field type.
'    'Use when copying or comparing rows where text/date/number/Boolean formatting must match DAO field metadata.
'    If IsNull(v) Then
'        SharedDao_SqlLiteralForFieldValue = "Null"
'        Exit Function
'    End If
'    Select Case fld.Type
'        Case dbBoolean
'            If CBool(v) Then
'                SharedDao_SqlLiteralForFieldValue = "True"
'            Else
'                SharedDao_SqlLiteralForFieldValue = "False"
'            End If
'        Case dbByte, dbInteger, dbLong, dbSingle, dbDouble, dbCurrency, dbDecimal, dbNumeric, dbBigInt
'            SharedDao_SqlLiteralForFieldValue = SharedDao_InvariantNumericLiteral(v)
'        Case dbDate, 22, 23
'            SharedDao_SqlLiteralForFieldValue = "#" & Format$(CDate(v), "yyyy-mm-dd hh:nn:ss") & "#"
'        Case dbGUID
'            SharedDao_SqlLiteralForFieldValue = "'" & Replace(CStr(v), "'", "''") & "'"
'        Case Else
'            SharedDao_SqlLiteralForFieldValue = "'" & Replace(CStr(v), "'", "''") & "'"
'    End Select
'End Function
'
'Public Function SharedDao_CurrentDbTableExists(ByVal tableName As String) As Boolean
'    'Returns True when a table exists in CurrentDb.
'    'Consolidated old names: TVSettings_TableExists, local TV NSS IMPORT MODULE TableExists.
'    SharedDao_CurrentDbTableExists = SharedDao_DatabaseTableExists(currentDb, tableName)
'End Function
'
'Public Function SharedDao_DatabaseTableExists(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
'    'Returns True when the supplied DAO.Database contains a table.
'    'Consolidated old names: TVSettings_TableExistsInDb, local TV DEV DEBUG MODULE TableExists(db, tableName).
'    Dim tdf As DAO.tableDef
'    For Each tdf In db.TableDefs
'        If StrComp(tdf.Name, tableName, vbTextCompare) = 0 Then
'            SharedDao_DatabaseTableExists = True
'            Exit Function
'        End If
'    Next tdf
'End Function
'
'Public Function SharedDao_CurrentDbTableFieldExists(ByVal tableName As String, ByVal fieldName As String) As Boolean
'    'Returns True when a CurrentDb table contains a field.
'    'Consolidated old names: TVSettings_TableHasField, FieldExistsInTable.
'    If SharedDao_CurrentDbTableExists(tableName) = False Then Exit Function
'    SharedDao_CurrentDbTableFieldExists = SharedDao_TableDefFieldExists(currentDb.TableDefs(tableName), fieldName)
'End Function
'
'Public Function SharedDao_TableDefFieldExists(ByVal tdf As DAO.tableDef, ByVal fieldName As String) As Boolean
'    'Returns True when the supplied DAO.TableDef contains a field.
'    'Consolidated old names: TVSettings_FieldExistsInTableDef.
'    Dim fld As DAO.Field
'    For Each fld In tdf.fields
'        If StrComp(fld.Name, fieldName, vbTextCompare) = 0 Then
'            SharedDao_TableDefFieldExists = True
'            Exit Function
'        End If
'    Next fld
'End Function
'
'Public Function SharedDao_RecordsetFieldExists(ByVal rs As DAO.Recordset, ByVal fieldName As String) As Boolean
'    'Returns True when the supplied DAO.Recordset contains a field.
'    'Consolidated old names: TVSettings_RecordsetHasField, FieldExistsInRecordset, RecordsetHasFieldVBA.
'    Dim fld As DAO.Field
'    For Each fld In rs.fields
'        If StrComp(fld.Name, fieldName, vbTextCompare) = 0 Then
'            SharedDao_RecordsetFieldExists = True
'            Exit Function
'        End If
'    Next fld
'End Function
'
'Public Function SharedDao_RecordsetFieldValueOrNull(ByVal rs As DAO.Recordset, ByVal fieldName As String) As Variant
'    'Returns a recordset field value, or Null when the field is not present.
'    'Consolidated old names: FieldValueVBA.
'    If SharedDao_RecordsetFieldExists(rs, fieldName) Then
'        SharedDao_RecordsetFieldValueOrNull = rs.fields(fieldName).value
'    Else
'        SharedDao_RecordsetFieldValueOrNull = Null
'    End If
'End Function
'
'Public Function SharedDao_TableDefIndexExists(ByVal tdf As DAO.tableDef, ByVal indexName As String) As Boolean
'    'Returns True when the supplied DAO.TableDef contains an index.
'    'Consolidated old names: TVSettings_IndexExistsInTableDef.
'    Dim idx As DAO.index
'    For Each idx In tdf.Indexes
'        If StrComp(idx.Name, indexName, vbTextCompare) = 0 Then
'            SharedDao_TableDefIndexExists = True
'            Exit Function
'        End If
'    Next idx
'End Function
'
'Public Function SharedDao_RelationExists(ByVal db As DAO.Database, ByVal relName As String) As Boolean
'    'Returns True when the supplied DAO.Database contains a relationship.
'    'Consolidated old names: TVSettings_RelationExists.
'    Dim rel As DAO.Relation
'    For Each rel In db.Relations
'        If StrComp(rel.Name, relName, vbTextCompare) = 0 Then
'            SharedDao_RelationExists = True
'            Exit Function
'        End If
'    Next rel
'End Function
'
'Public Function SharedDao_GetTableDef(ByVal db As DAO.Database, ByVal tableName As String) As DAO.tableDef
'    'Returns a DAO.TableDef by name from the supplied database.
'    'Use when the caller already validated existence or wants DAO to raise an error if missing.
'    db.TableDefs.Refresh
'    Set SharedDao_GetTableDef = db.TableDefs(tableName)
'End Function
'
'Public Function SharedDao_IsAccessLinkedTable(ByVal tdf As DAO.tableDef) As Boolean
'    'Returns True when a TableDef has an Access linked-table Connect string.
'    'Use before relinking, deleting, or inspecting linked-table targets.
'    If Len(tdf.Connect) = 0 Then Exit Function
'    If InStr(1, tdf.Connect, SHARED_DAO_ACCESS_CONNECT_DATABASE_MARKER, vbTextCompare) = 0 Then Exit Function
'    SharedDao_IsAccessLinkedTable = True
'End Function
'
'Public Function SharedDao_AccessDatabasePathFromConnect(ByVal connectText As String) As String
'    'Extracts the linked database path from an Access linked-table Connect string.
'    'Consolidated old names: TVSettings_DatabasePathFromConnect.
'    Dim p As Long
'    p = InStr(1, connectText, SHARED_DAO_ACCESS_CONNECT_DATABASE_MARKER, vbTextCompare)
'    If p > 0 Then SharedDao_AccessDatabasePathFromConnect = Mid$(connectText, p + Len(SHARED_DAO_ACCESS_CONNECT_DATABASE_MARKER))
'End Function
'
'Public Function SharedDao_LinkedTableDataPath(ByVal tdf As DAO.tableDef) As String
'    'Extracts the linked database path from an Access linked TableDef Connect string.
'    'Use after SharedDao_IsAccessLinkedTable confirms the table is an Access linked table.
'    SharedDao_LinkedTableDataPath = SharedDao_AccessDatabasePathFromConnect(tdf.Connect)
'End Function
'
'Public Function SharedDao_SingleLinkedDatabasePath() As String
'    'Returns the single linked Access data database path used by CurrentDb.
'    'Raises an error when more than one linked Access data database is found.
'    'Consolidated old names: TVSettings_SingleLinkedDatabasePath.
'    Dim tdf As DAO.tableDef
'    Dim paths As Collection
'    Dim pathText As String
'    Dim item As Variant
'    Set paths = New Collection
'    For Each tdf In currentDb.TableDefs
'        If SharedDao_IsAccessLinkedTable(tdf) Then
'            pathText = SharedDao_LinkedTableDataPath(tdf)
'            If Len(pathText) > 0 Then
'                On Error Resume Next
'                paths.Add pathText, LCase$(pathText)
'                On Error GoTo 0
'            End If
'        End If
'    Next tdf
'    If paths.count = 0 Then
'        SharedDao_SingleLinkedDatabasePath = vbNullString
'    ElseIf paths.count = 1 Then
'        SharedDao_SingleLinkedDatabasePath = CStr(paths(1))
'    Else
'        pathText = vbNullString
'        For Each item In paths
'            pathText = pathText & CStr(item) & vbCrLf
'        Next item
'        Err.Raise vbObjectError + 7301, "SharedDao_SingleLinkedDatabasePath", "More than one linked data database was found. Maintenance requires exactly one linked data database." & vbCrLf & vbCrLf & pathText
'    End If
'End Function
'
'Public Sub SharedDao_LinkAccessTableIfMissing(ByVal tableName As String, ByVal dataDbPath As String, Optional ByVal sourceTableName As String = vbNullString)
'    'Intended use: Creates a missing CurrentDb Access-table link or verifies that an existing table name already targets the exact requested database and source table.
'    On Error GoTo ErrHandler
'    Dim db As DAO.Database
'    Dim tdf As DAO.tableDef
'    Dim requestedPath As String
'    Dim tableExistedBefore As Boolean
'    Dim tableStateCaptured As Boolean
'    dataDbPath = Trim$(dataDbPath)
'    If Len(tableName) = 0 Then Err.Raise vbObjectError + 7302, "SharedDao_LinkAccessTableIfMissing", "The linked table name is blank."
'    If Len(dataDbPath) = 0 Then Err.Raise vbObjectError + 7303, "SharedDao_LinkAccessTableIfMissing", "The linked database path is blank." & vbCrLf & "Table: " & tableName
'    If Len(sourceTableName) = 0 Then sourceTableName = tableName
'    requestedPath = SharedPath_FullPathStrict(dataDbPath, "SharedDao_LinkAccessTableIfMissing")
'    Set db = currentDb
'    db.TableDefs.Refresh
'    tableExistedBefore = SharedDao_DatabaseTableExists(db, tableName)
'    tableStateCaptured = True
'    If tableExistedBefore Then
'        Set tdf = db.TableDefs(tableName)
'        SharedDao_VerifyAccessTableLink tdf, requestedPath, sourceTableName
'        Exit Sub
'    End If
'    Set tdf = db.CreateTableDef(tableName)
'    tdf.Connect = SHARED_DAO_ACCESS_CONNECT_DATABASE_MARKER & requestedPath
'    tdf.sourceTableName = sourceTableName
'    db.TableDefs.Append tdf
'    db.TableDefs.Refresh
'    Set tdf = db.TableDefs(tableName)
'    SharedDao_VerifyAccessTableLink tdf, requestedPath, sourceTableName
'    Set tdf = Nothing
'    Set db = Nothing
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim cleanupText As String
'    Dim cleanupErrNumber As Long
'    Dim cleanupErrDescription As String
'    Dim tableExistsAfterFailure As Boolean
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If tableStateCaptured And tableExistedBefore = False And Not db Is Nothing Then
'        Err.Clear
'        db.TableDefs.Refresh
'        cleanupErrNumber = Err.Number
'        cleanupErrDescription = Err.Description
'        If cleanupErrNumber <> 0 Then
'            cleanupText = vbCrLf & "Failed-link cleanup inspection error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription
'        Else
'            Err.Clear
'            tableExistsAfterFailure = SharedDao_DatabaseTableExists(db, tableName)
'            cleanupErrNumber = Err.Number
'            cleanupErrDescription = Err.Description
'            If cleanupErrNumber <> 0 Then
'                cleanupText = vbCrLf & "Failed-link cleanup existence error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription
'            ElseIf tableExistsAfterFailure Then
'                Err.Clear
'                db.TableDefs.Delete tableName
'                db.TableDefs.Refresh
'                cleanupErrNumber = Err.Number
'                cleanupErrDescription = Err.Description
'                If cleanupErrNumber <> 0 Then
'                    cleanupText = vbCrLf & "Failed-link cleanup deletion error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription
'                Else
'                    Err.Clear
'                    tableExistsAfterFailure = SharedDao_DatabaseTableExists(db, tableName)
'                    cleanupErrNumber = Err.Number
'                    cleanupErrDescription = Err.Description
'                    If cleanupErrNumber <> 0 Then
'                        cleanupText = vbCrLf & "Failed-link cleanup verification error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription
'                    ElseIf tableExistsAfterFailure Then
'                        cleanupText = vbCrLf & "Failed-link cleanup failed verification: the TableDef still exists."
'                    Else
'                        cleanupText = vbCrLf & "Failed link was removed and its absence was verified."
'                    End If
'                End If
'            Else
'                cleanupText = vbCrLf & "No failed link remained after the error."
'            End If
'        End If
'    ElseIf Not db Is Nothing And tableStateCaptured = False Then
'        cleanupText = vbCrLf & "No failed-link cleanup was attempted because the pre-operation table state was not captured conclusively."
'    End If
'    Set tdf = Nothing
'    Set db = Nothing
'    On Error GoTo 0
'    If Len(requestedPath) = 0 Then requestedPath = dataDbPath
'    Err.Raise savedErrNumber, "SharedDao_LinkAccessTableIfMissing", "Unable to create or verify the Access linked table." & vbCrLf & "Table: " & tableName & vbCrLf & "Requested database: " & requestedPath & vbCrLf & "Requested source table: " & sourceTableName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Original description: " & savedErrDescription & cleanupText
'End Sub
'Private Sub SharedDao_VerifyAccessTableLink(ByVal tdf As DAO.tableDef, ByVal requestedPath As String, ByVal requestedSourceTableName As String)
'    'Verifies that one existing TableDef is an Access link to the exact requested database path and source table.
'    Dim actualPath As String
'    Dim actualSourceTableName As String
'    If tdf Is Nothing Then Err.Raise vbObjectError + 7304, "SharedDao_VerifyAccessTableLink", "The linked TableDef is Nothing."
'    If SharedDao_IsAccessLinkedTable(tdf) = False Then Err.Raise vbObjectError + 7305, "SharedDao_VerifyAccessTableLink", "An object with the requested linked-table name exists but is not an Access linked table." & vbCrLf & "Table: " & tdf.Name & vbCrLf & "Connect: " & tdf.Connect
'    actualPath = SharedPath_FullPathStrict(SharedDao_LinkedTableDataPath(tdf), "SharedDao_VerifyAccessTableLink")
'    actualSourceTableName = Trim$(tdf.sourceTableName)
'    requestedPath = SharedPath_FullPathStrict(requestedPath, "SharedDao_VerifyAccessTableLink")
'    If StrComp(actualPath, requestedPath, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 7306, "SharedDao_VerifyAccessTableLink", "The existing linked table points to the wrong database." & vbCrLf & "Table: " & tdf.Name & vbCrLf & "Requested database: " & requestedPath & vbCrLf & "Actual database: " & actualPath
'    If StrComp(actualSourceTableName, requestedSourceTableName, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 7307, "SharedDao_VerifyAccessTableLink", "The existing linked table points to the wrong source table." & vbCrLf & "Table: " & tdf.Name & vbCrLf & "Requested source table: " & requestedSourceTableName & vbCrLf & "Actual source table: " & actualSourceTableName
'End Sub
'
'Public Function SharedDao_EquivalentIndexExists(ByVal targetTdf As DAO.tableDef, ByVal sourceIdx As DAO.index) As Boolean
'    'Returns True when a table has an index structurally equivalent to the source index.
'    'Use when index names may differ but field order and critical properties must match.
'    Dim idx As DAO.index
'    For Each idx In targetTdf.Indexes
'        If SharedDao_IndexSignaturesMatch(idx, sourceIdx) Then
'            SharedDao_EquivalentIndexExists = True
'            Exit Function
'        End If
'    Next idx
'End Function
'
'Public Function SharedDao_IndexSignaturesMatch(ByVal idxA As DAO.index, ByVal idxB As DAO.index) As Boolean
'    'Compares two DAO indexes by field order and critical index properties.
'    'Use for schema validation or copy routines where index structure matters more than object name.
'    Dim i As Long
'    If idxA.fields.count <> idxB.fields.count Then Exit Function
'    If idxA.Primary <> idxB.Primary Then Exit Function
'    If idxA.Unique <> idxB.Unique Then Exit Function
'    If idxA.Required <> idxB.Required Then Exit Function
'    If idxA.ignoreNulls <> idxB.ignoreNulls Then Exit Function
'    For i = 0 To idxA.fields.count - 1
'        If StrComp(idxA.fields(i).Name, idxB.fields(i).Name, vbTextCompare) <> 0 Then Exit Function
'        If idxA.fields(i).attributes <> idxB.fields(i).attributes Then Exit Function
'    Next i
'    SharedDao_IndexSignaturesMatch = True
'End Function
'
'Public Function SharedDao_IndexContainsField(ByVal idx As DAO.index, ByVal fieldName As String) As Boolean
'    'Returns True when a DAO index includes a field by name.
'    'Consolidated old names: TVSettings_IndexContainsField, local TV NSS IMPORT MODULE IndexContainsField.
'    Dim fld As DAO.Field
'    For Each fld In idx.fields
'        If StrComp(fld.Name, fieldName, vbTextCompare) = 0 Then
'            SharedDao_IndexContainsField = True
'            Exit Function
'        End If
'    Next fld
'End Function
'
'Public Sub SharedDao_RemoveIndexesContainingField(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String, ByRef reportText As String)
'    'Intended use: Removes every non-foreign index containing a field, verifies each deletion, and exactly restores every captured definition if any mutation fails.
'    On Error GoTo ErrHandler
'    Dim tdf As DAO.tableDef
'    Dim idx As DAO.index
'    Dim Definition As Object
'    Dim definitions As Collection
'    Dim i As Long
'    Dim idxName As String
'    Set definitions = New Collection
'    If db Is Nothing Then Err.Raise vbObjectError + 7310, "SharedDao_RemoveIndexesContainingField", "The DAO.Database argument is Nothing."
'    fieldName = Trim$(fieldName)
'    If Len(tableName) = 0 Or Len(fieldName) = 0 Then Err.Raise vbObjectError + 7311, "SharedDao_RemoveIndexesContainingField", "The table name and field name are required."
'    Set tdf = db.TableDefs(tableName)
'    tdf.Indexes.Refresh
'    For Each idx In tdf.Indexes
'        If SharedDao_IndexContainsField(idx, fieldName) Then
'            If idx.Foreign Then Err.Raise vbObjectError + 7312, "SharedDao_RemoveIndexesContainingField", "A foreign index contains the field. Remove or repair the relationship before dropping the field." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName & vbCrLf & "Index: " & idx.Name
'            Set Definition = SharedDao_CaptureIndexDefinition(idx)
'            definitions.Add Definition
'        End If
'    Next idx
'    For i = 1 To definitions.count
'        Set Definition = definitions(i)
'        idxName = CStr(Definition("Name"))
'        tdf.Indexes.Delete idxName
'        tdf.Indexes.Refresh
'        If SharedDao_TableDefIndexExists(tdf, idxName) Then Err.Raise vbObjectError + 7313, "SharedDao_RemoveIndexesContainingField", "An index still exists after deletion." & vbCrLf & "Table: " & tableName & vbCrLf & "Index: " & idxName
'    Next i
'    For i = 1 To definitions.count
'        Set Definition = definitions(i)
'        reportText = reportText & tableName & "." & CStr(Definition("Name")) & ": index removed before dropping extra field " & fieldName & "." & vbCrLf
'    Next i
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim restoredText As String
'    Dim restoreFailureText As String
'    Dim verificationText As String
'    Dim restoreErrNumber As Long
'    Dim restoreErrDescription As String
'    Dim inspectionErrNumber As Long
'    Dim inspectionErrDescription As String
'    Dim IndexExists As Boolean
'    Dim indexMatches As Boolean
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If Not tdf Is Nothing And Not definitions Is Nothing Then
'        For i = 1 To definitions.count
'            Set Definition = definitions(i)
'            idxName = CStr(Definition("Name"))
'            Err.Clear
'            SharedDao_RestoreIndexDefinitionExact tdf, Definition
'            restoreErrNumber = Err.Number
'            restoreErrDescription = Err.Description
'            If restoreErrNumber = 0 Then
'                restoredText = restoredText & IIf(Len(restoredText) > 0, ", ", vbNullString) & idxName
'            Else
'                restoreFailureText = restoreFailureText & vbCrLf & "Index " & idxName & " restore error " & CStr(restoreErrNumber) & ": " & restoreErrDescription
'            End If
'        Next i
'        Err.Clear
'        tdf.Indexes.Refresh
'        inspectionErrNumber = Err.Number
'        inspectionErrDescription = Err.Description
'        If inspectionErrNumber <> 0 Then
'            verificationText = verificationText & vbCrLf & "Complete post-restore collection refresh error " & CStr(inspectionErrNumber) & ": " & inspectionErrDescription
'        Else
'            For i = 1 To definitions.count
'                Set Definition = definitions(i)
'                idxName = CStr(Definition("Name"))
'                Err.Clear
'                IndexExists = SharedDao_TableDefIndexExists(tdf, idxName)
'                inspectionErrNumber = Err.Number
'                inspectionErrDescription = Err.Description
'                If inspectionErrNumber <> 0 Then
'                    verificationText = verificationText & vbCrLf & "Index " & idxName & " post-restore inspection error " & CStr(inspectionErrNumber) & ": " & inspectionErrDescription
'                ElseIf IndexExists = False Then
'                    verificationText = verificationText & vbCrLf & "Index " & idxName & " is missing after restoration."
'                Else
'                    Err.Clear
'                    indexMatches = SharedDao_IndexDefinitionMatches(tdf.Indexes(idxName), Definition)
'                    inspectionErrNumber = Err.Number
'                    inspectionErrDescription = Err.Description
'                    If inspectionErrNumber <> 0 Then
'                        verificationText = verificationText & vbCrLf & "Index " & idxName & " post-restore comparison error " & CStr(inspectionErrNumber) & ": " & inspectionErrDescription
'                    ElseIf indexMatches = False Then
'                        verificationText = verificationText & vbCrLf & "Index " & idxName & " remains malformed after restoration."
'                    End If
'                End If
'            Next i
'        End If
'    End If
'    On Error GoTo 0
'    If Len(restoredText) > 0 Then restoredText = vbCrLf & "Restored and verified indexes: " & restoredText
'    If Len(restoreFailureText) > 0 Then restoreFailureText = vbCrLf & "Restore failures:" & restoreFailureText
'    If Len(verificationText) > 0 Then verificationText = vbCrLf & "Post-restore verification failures:" & verificationText
'    Err.Raise savedErrNumber, "SharedDao_RemoveIndexesContainingField", "Unable to remove indexes containing the field. Exact restoration was attempted for every captured index." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Original description: " & savedErrDescription & restoredText & restoreFailureText & verificationText
'End Sub
'Private Function SharedDao_CaptureIndexDefinition(ByVal idx As DAO.index) As Object
'    'Captures the mutable DAO index definition needed for exact restoration after a failed multi-index removal.
'    Dim Definition As Object
'    Dim fieldNames() As String
'    Dim fieldAttributes() As Long
'    Dim i As Long
'    If idx Is Nothing Then Err.Raise vbObjectError + 7314, "SharedDao_CaptureIndexDefinition", "The DAO.Index argument is Nothing."
'    If idx.fields.count = 0 Then Err.Raise vbObjectError + 7315, "SharedDao_CaptureIndexDefinition", "The index has no fields." & vbCrLf & "Index: " & idx.Name
'    ReDim fieldNames(0 To idx.fields.count - 1)
'    ReDim fieldAttributes(0 To idx.fields.count - 1)
'    For i = 0 To idx.fields.count - 1
'        fieldNames(i) = idx.fields(i).Name
'        fieldAttributes(i) = idx.fields(i).attributes
'    Next i
'    Set Definition = CreateObject("Scripting.Dictionary")
'    Definition.Add "Name", idx.Name
'    Definition.Add "Primary", idx.Primary
'    Definition.Add "Unique", idx.Unique
'    Definition.Add "Required", idx.Required
'    Definition.Add "IgnoreNulls", idx.ignoreNulls
'    Definition.Add "FieldNames", fieldNames
'    Definition.Add "FieldAttributes", fieldAttributes
'    Set SharedDao_CaptureIndexDefinition = Definition
'End Function
'
'Private Sub SharedDao_RestoreIndexDefinitionExact(ByVal tdf As DAO.tableDef, ByVal Definition As Object)
'    'Restores one captured index exactly, replacing an existing malformed copy and verifying the final definition.
'    On Error GoTo ErrHandler
'    Dim idxName As String
'    If tdf Is Nothing Then Err.Raise vbObjectError + 7318, "SharedDao_RestoreIndexDefinitionExact", "The DAO.TableDef argument is Nothing."
'    idxName = CStr(Definition("Name"))
'    tdf.Indexes.Refresh
'    If SharedDao_TableDefIndexExists(tdf, idxName) Then
'        If SharedDao_IndexDefinitionMatches(tdf.Indexes(idxName), Definition) Then Exit Sub
'        tdf.Indexes.Delete idxName
'        tdf.Indexes.Refresh
'        If SharedDao_TableDefIndexExists(tdf, idxName) Then Err.Raise vbObjectError + 7319, "SharedDao_RestoreIndexDefinitionExact", "The malformed index still exists after deletion."
'    End If
'    SharedDao_AppendIndexDefinition tdf, Definition
'    tdf.Indexes.Refresh
'    If SharedDao_TableDefIndexExists(tdf, idxName) = False Then Err.Raise vbObjectError + 7320, "SharedDao_RestoreIndexDefinitionExact", "The index is missing after exact restoration."
'    If SharedDao_IndexDefinitionMatches(tdf.Indexes(idxName), Definition) = False Then Err.Raise vbObjectError + 7321, "SharedDao_RestoreIndexDefinitionExact", "The index does not match its captured definition after exact restoration."
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedDao_RestoreIndexDefinitionExact", "Unable to restore the captured index definition exactly." & vbCrLf & "Table: " & SharedDao_SafeTableDefName(tdf) & vbCrLf & "Index: " & idxName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Sub
'
'Private Sub SharedDao_AppendIndexDefinition(ByVal tdf As DAO.tableDef, ByVal Definition As Object)
'    'Restores one captured index definition, removes any failed replacement, and verifies the resulting name, properties, field order, and field attributes.
'    On Error GoTo ErrHandler
'    Dim newIdx As DAO.index
'    Dim newFld As DAO.Field
'    Dim fieldNames As Variant
'    Dim fieldAttributes As Variant
'    Dim i As Long
'    Dim idxName As String
'    Dim appendAttempted As Boolean
'    idxName = CStr(Definition("Name"))
'    fieldNames = Definition("FieldNames")
'    fieldAttributes = Definition("FieldAttributes")
'    Set newIdx = tdf.CreateIndex(idxName)
'    newIdx.Primary = CBool(Definition("Primary"))
'    newIdx.Unique = CBool(Definition("Unique"))
'    newIdx.Required = CBool(Definition("Required"))
'    newIdx.ignoreNulls = CBool(Definition("IgnoreNulls"))
'    For i = LBound(fieldNames) To UBound(fieldNames)
'        Set newFld = newIdx.CreateField(CStr(fieldNames(i)))
'        newFld.attributes = CLng(fieldAttributes(i))
'        newIdx.fields.Append newFld
'    Next i
'    appendAttempted = True
'    tdf.Indexes.Append newIdx
'    tdf.Indexes.Refresh
'    If SharedDao_TableDefIndexExists(tdf, idxName) = False Then Err.Raise vbObjectError + 7316, "SharedDao_AppendIndexDefinition", "The restored index is missing after append." & vbCrLf & "Table: " & tdf.Name & vbCrLf & "Index: " & idxName
'    If SharedDao_IndexDefinitionMatches(tdf.Indexes(idxName), Definition) = False Then Err.Raise vbObjectError + 7317, "SharedDao_AppendIndexDefinition", "The restored index does not match its captured definition." & vbCrLf & "Table: " & tdf.Name & vbCrLf & "Index: " & idxName
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim cleanupText As String
'    Dim cleanupErrNumber As Long
'    Dim cleanupErrDescription As String
'    Dim replacementExists As Boolean
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If appendAttempted And Not tdf Is Nothing Then
'        Err.Clear
'        tdf.Indexes.Refresh
'        replacementExists = SharedDao_TableDefIndexExists(tdf, idxName)
'        cleanupErrNumber = Err.Number
'        cleanupErrDescription = Err.Description
'        If cleanupErrNumber <> 0 Then
'            cleanupText = vbCrLf & "Failed replacement inspection error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription
'        ElseIf replacementExists Then
'            Err.Clear
'            tdf.Indexes.Delete idxName
'            tdf.Indexes.Refresh
'            cleanupErrNumber = Err.Number
'            cleanupErrDescription = Err.Description
'            If cleanupErrNumber <> 0 Then
'                cleanupText = vbCrLf & "Failed replacement deletion error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription
'            Else
'                replacementExists = SharedDao_TableDefIndexExists(tdf, idxName)
'                cleanupErrNumber = Err.Number
'                cleanupErrDescription = Err.Description
'                If cleanupErrNumber <> 0 Then
'                    cleanupText = vbCrLf & "Failed replacement deletion verification error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription
'                ElseIf replacementExists Then
'                    cleanupText = vbCrLf & "Failed replacement deletion did not remove the malformed index."
'                Else
'                    cleanupText = vbCrLf & "Failed replacement index was removed and verified absent."
'                End If
'            End If
'        End If
'    End If
'    Set newFld = Nothing
'    Set newIdx = Nothing
'    On Error GoTo 0
'    Err.Raise savedErrNumber, "SharedDao_AppendIndexDefinition", "Unable to append and verify the captured index definition." & vbCrLf & "Table: " & SharedDao_SafeTableDefName(tdf) & vbCrLf & "Index: " & idxName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & cleanupText
'End Sub
'
'Private Function SharedDao_IndexDefinitionMatches(ByVal idx As DAO.index, ByVal Definition As Object) As Boolean
'    'Returns True only when an index exactly matches a previously captured restorable definition.
'    Dim fieldNames As Variant
'    Dim fieldAttributes As Variant
'    Dim i As Long
'    fieldNames = Definition("FieldNames")
'    fieldAttributes = Definition("FieldAttributes")
'    If StrComp(idx.Name, CStr(Definition("Name")), vbTextCompare) <> 0 Then Exit Function
'    If idx.Primary <> CBool(Definition("Primary")) Then Exit Function
'    If idx.Unique <> CBool(Definition("Unique")) Then Exit Function
'    If idx.Required <> CBool(Definition("Required")) Then Exit Function
'    If idx.ignoreNulls <> CBool(Definition("IgnoreNulls")) Then Exit Function
'    If idx.fields.count <> UBound(fieldNames) - LBound(fieldNames) + 1 Then Exit Function
'    For i = LBound(fieldNames) To UBound(fieldNames)
'        If StrComp(idx.fields(i - LBound(fieldNames)).Name, CStr(fieldNames(i)), vbTextCompare) <> 0 Then Exit Function
'        If idx.fields(i - LBound(fieldNames)).attributes <> CLng(fieldAttributes(i)) Then Exit Function
'    Next i
'    SharedDao_IndexDefinitionMatches = True
'End Function
'
'Public Function SharedDao_FieldTypeUsesExplicitSize(ByVal fieldType As Long) As Boolean
'    'Returns True for field types that need an explicit Size argument when created.
'    'Use when copying DAO fields from one TableDef to another.
'    Select Case fieldType
'        Case dbText, dbBinary
'            SharedDao_FieldTypeUsesExplicitSize = True
'    End Select
'End Function
'
'Public Sub SharedDao_CopyFieldPropertyIfSupported(ByVal targetFld As DAO.Field, ByVal sourceFld As DAO.Field, ByVal propertyName As String, Optional ByVal mustCopy As Boolean = False)
'    'Copies a supported DAO field property from source field to target field.
'    'Consolidated old names: TVSettings_CopyFieldProperties.
'    On Error GoTo ErrHandler
'    Select Case propertyName
'        Case "Attributes"
'            targetFld.attributes = sourceFld.attributes
'        Case "AllowZeroLength"
'            targetFld.allowZeroLength = sourceFld.allowZeroLength
'        Case "Required"
'            targetFld.Required = sourceFld.Required
'        Case "DefaultValue"
'            targetFld.DefaultValue = sourceFld.DefaultValue
'        Case "ValidationRule"
'            targetFld.validationRule = sourceFld.validationRule
'        Case "ValidationText"
'            targetFld.validationText = sourceFld.validationText
'        Case "OrdinalPosition"
'            targetFld.OrdinalPosition = sourceFld.OrdinalPosition
'    End Select
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrDescription = Err.Description
'    If mustCopy Then Err.Raise savedErrNumber, "SharedDao_CopyFieldPropertyIfSupported", "Could not copy critical field property '" & propertyName & "' for field '" & sourceFld.Name & "': " & savedErrDescription
'    Err.Clear
'End Sub
'
'Public Sub SharedDao_AppendCopiedField(ByVal targetTdf As DAO.tableDef, ByVal sourceFld As DAO.Field)
'    'Intended use: Creates, appends, and verifies a copy of a DAO field in a target TableDef.
'    'Consolidated old names: TVSettings_CopyFieldForTableDef, TVSettings_CopyFieldProperties.
'    On Error GoTo ErrHandler
'    Dim newFld As DAO.Field
'    Dim persistedFld As DAO.Field
'    Dim fieldStateCaptured As Boolean
'    Dim fieldExistedBefore As Boolean
'    Dim appendAttempted As Boolean
'    Dim fieldAppended As Boolean
'    Dim expectedAutoNumber As Boolean
'    Dim targetName As String
'    Dim fieldName As String
'    If targetTdf Is Nothing Then Err.Raise vbObjectError + 8003, "SharedDao_AppendCopiedField", "The target DAO.TableDef argument is Nothing."
'    If sourceFld Is Nothing Then Err.Raise vbObjectError + 8004, "SharedDao_AppendCopiedField", "The source DAO.Field argument is Nothing."
'    targetName = targetTdf.Name
'    fieldName = sourceFld.Name
'    fieldExistedBefore = SharedDao_TableDefFieldExists(targetTdf, fieldName)
'    fieldStateCaptured = True
'    If SharedDao_FieldTypeUsesExplicitSize(sourceFld.Type) Then
'        Set newFld = targetTdf.CreateField(fieldName, sourceFld.Type, sourceFld.Size)
'    Else
'        Set newFld = targetTdf.CreateField(fieldName, sourceFld.Type)
'    End If
'    SharedDao_CopyFieldPropertyIfSupported newFld, sourceFld, "Attributes", True
'    If sourceFld.Type = dbText Or sourceFld.Type = dbMemo Then SharedDao_CopyFieldPropertyIfSupported newFld, sourceFld, "AllowZeroLength", True
'    SharedDao_CopyFieldPropertyIfSupported newFld, sourceFld, "Required", True
'    If Len(Nz(sourceFld.DefaultValue, vbNullString)) > 0 Then SharedDao_CopyFieldPropertyIfSupported newFld, sourceFld, "DefaultValue", True
'    If Len(Nz(sourceFld.validationRule, vbNullString)) > 0 Then SharedDao_CopyFieldPropertyIfSupported newFld, sourceFld, "ValidationRule", True
'    If Len(Nz(sourceFld.validationText, vbNullString)) > 0 Then SharedDao_CopyFieldPropertyIfSupported newFld, sourceFld, "ValidationText", True
'    SharedDao_CopyFieldPropertyIfSupported newFld, sourceFld, "OrdinalPosition", False
'    appendAttempted = True
'    targetTdf.fields.Append newFld
'    fieldAppended = True
'    targetTdf.fields.Refresh
'    If SharedDao_TableDefFieldExists(targetTdf, fieldName) = False Then Err.Raise vbObjectError + 8005, "SharedDao_AppendCopiedField", "The copied field is missing after append and refresh."
'    Set persistedFld = targetTdf.fields(fieldName)
'    expectedAutoNumber = ((sourceFld.attributes And dbAutoIncrField) <> 0)
'    If SharedDao_FieldDefinitionMatches(persistedFld, sourceFld.Type, sourceFld.Size, expectedAutoNumber, False) = False Then Err.Raise vbObjectError + 8006, "SharedDao_AppendCopiedField", "The copied field definition does not match the source field."
'    If persistedFld.attributes <> sourceFld.attributes Then Err.Raise vbObjectError + 8007, "SharedDao_AppendCopiedField", "The copied field Attributes value does not match the source field."
'    If sourceFld.Type = dbText Or sourceFld.Type = dbMemo Then If persistedFld.allowZeroLength <> sourceFld.allowZeroLength Then Err.Raise vbObjectError + 8008, "SharedDao_AppendCopiedField", "The copied field AllowZeroLength value does not match the source field."
'    If persistedFld.Required <> sourceFld.Required Then Err.Raise vbObjectError + 8009, "SharedDao_AppendCopiedField", "The copied field Required value does not match the source field."
'    If Nz(persistedFld.DefaultValue, vbNullString) <> Nz(sourceFld.DefaultValue, vbNullString) Then Err.Raise vbObjectError + 8010, "SharedDao_AppendCopiedField", "The copied field DefaultValue does not match the source field."
'    If Nz(persistedFld.validationRule, vbNullString) <> Nz(sourceFld.validationRule, vbNullString) Then Err.Raise vbObjectError + 8011, "SharedDao_AppendCopiedField", "The copied field ValidationRule does not match the source field."
'    If Nz(persistedFld.validationText, vbNullString) <> Nz(sourceFld.validationText, vbNullString) Then Err.Raise vbObjectError + 8012, "SharedDao_AppendCopiedField", "The copied field ValidationText does not match the source field."
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim cleanupText As String
'    Dim cleanupErrNumber As Long
'    Dim cleanupErrDescription As String
'    Dim fieldStillExists As Boolean
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If appendAttempted And fieldStateCaptured And fieldExistedBefore = False And Not targetTdf Is Nothing And Len(fieldName) > 0 Then
'        Err.Clear
'        fieldStillExists = SharedDao_TableDefFieldExists(targetTdf, fieldName)
'        cleanupErrNumber = Err.Number
'        cleanupErrDescription = Err.Description
'        If cleanupErrNumber = 0 And fieldStillExists Then
'            Err.Clear
'            targetTdf.fields.Delete fieldName
'            targetTdf.fields.Refresh
'            cleanupErrNumber = Err.Number
'            cleanupErrDescription = Err.Description
'        End If
'        If cleanupErrNumber = 0 Then
'            Err.Clear
'            fieldStillExists = SharedDao_TableDefFieldExists(targetTdf, fieldName)
'            cleanupErrNumber = Err.Number
'            cleanupErrDescription = Err.Description
'        End If
'        If cleanupErrNumber <> 0 Then
'            cleanupText = vbCrLf & "Failed-field cleanup or verification error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription
'        ElseIf fieldStillExists Then
'            cleanupText = vbCrLf & "Failed-field cleanup did not remove the field."
'        Else
'            cleanupText = vbCrLf & "Any field persisted by the failed append was removed and verified absent."
'        End If
'    ElseIf appendAttempted And fieldStateCaptured And fieldExistedBefore Then
'        cleanupText = vbCrLf & "No field cleanup was attempted because the target field existed before this invocation."
'    Else
'        cleanupText = vbCrLf & "No field cleanup was attempted because no owned append mutation began."
'    End If
'    On Error GoTo 0
'    cleanupText = cleanupText & vbCrLf & "Mutation state: prior state captured=" & CStr(fieldStateCaptured) & "; existed before=" & CStr(fieldExistedBefore) & "; append attempted=" & CStr(appendAttempted) & "; append returned=" & CStr(fieldAppended)
'    Err.Raise savedErrNumber, "SharedDao_AppendCopiedField", "Could not copy and verify the DAO field." & vbCrLf & "Target table: " & targetName & vbCrLf & "Field: " & fieldName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & cleanupText
'End Sub
'
'Public Sub SharedDao_AppendCopiedIndex(ByVal targetTdf As DAO.tableDef, ByVal sourceIdx As DAO.index)
'    'Intended use: Creates, appends, and exactly verifies a copy of a DAO index in a target TableDef.
'    'Consolidated old names: TVSettings_CopyOneIndex.
'    On Error GoTo ErrHandler
'    Dim Definition As Object
'    Dim targetName As String
'    Dim sourceIndexName As String
'    If targetTdf Is Nothing Then Err.Raise vbObjectError + 8013, "SharedDao_AppendCopiedIndex", "The target DAO.TableDef argument is Nothing."
'    If sourceIdx Is Nothing Then Err.Raise vbObjectError + 8014, "SharedDao_AppendCopiedIndex", "The source DAO.Index argument is Nothing."
'    targetName = targetTdf.Name
'    sourceIndexName = sourceIdx.Name
'    targetTdf.Indexes.Refresh
'    If SharedDao_TableDefIndexExists(targetTdf, sourceIndexName) Then
'        If SharedDao_IndexSignaturesMatch(targetTdf.Indexes(sourceIndexName), sourceIdx) Then Exit Sub
'        If SharedDao_EquivalentIndexExists(targetTdf, sourceIdx) Then Exit Sub
'        Err.Raise vbObjectError + 8015, "SharedDao_AppendCopiedIndex", "An incompatible target index already uses the source index name."
'    End If
'    Set Definition = SharedDao_CaptureIndexDefinition(sourceIdx)
'    SharedDao_AppendIndexDefinition targetTdf, Definition
'    targetTdf.Indexes.Refresh
'    If SharedDao_EquivalentIndexExists(targetTdf, sourceIdx) = False Then Err.Raise vbObjectError + 8016, "SharedDao_AppendCopiedIndex", "No exactly equivalent target index exists after the copy operation."
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If Len(targetName) = 0 And Not targetTdf Is Nothing Then targetName = targetTdf.Name
'    If Len(sourceIndexName) = 0 And Not sourceIdx Is Nothing Then sourceIndexName = sourceIdx.Name
'    On Error GoTo 0
'    Err.Raise savedErrNumber, "SharedDao_AppendCopiedIndex", "Could not copy and exactly verify the index." & vbCrLf & "Target table: " & targetName & vbCrLf & "Index: " & sourceIndexName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Sub
'
'Public Sub SharedDao_CopyTemplateIndexes(ByVal templateTdf As DAO.tableDef, ByVal targetTdf As DAO.tableDef)
'    'Intended use: Copies every non-foreign template index and verifies an exact equivalent in the target TableDef.
'    'Use after creating a target table from a template table.
'    On Error GoTo ErrHandler
'    Dim idx As DAO.index
'    If templateTdf Is Nothing Then Err.Raise vbObjectError + 8017, "SharedDao_CopyTemplateIndexes", "The template DAO.TableDef argument is Nothing."
'    If targetTdf Is Nothing Then Err.Raise vbObjectError + 8018, "SharedDao_CopyTemplateIndexes", "The target DAO.TableDef argument is Nothing."
'    templateTdf.Indexes.Refresh
'    targetTdf.Indexes.Refresh
'    For Each idx In templateTdf.Indexes
'        If Not idx.Foreign Then SharedDao_AppendCopiedIndex targetTdf, idx
'    Next idx
'    targetTdf.Indexes.Refresh
'    For Each idx In templateTdf.Indexes
'        If Not idx.Foreign Then If SharedDao_EquivalentIndexExists(targetTdf, idx) = False Then Err.Raise vbObjectError + 8019, "SharedDao_CopyTemplateIndexes", "A template index has no exactly equivalent target index after copying." & vbCrLf & "Index: " & idx.Name
'    Next idx
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedDao_CopyTemplateIndexes", "Could not copy and verify all template indexes." & vbCrLf & "Template table: " & SharedDao_SafeTableDefName(templateTdf) & vbCrLf & "Target table: " & SharedDao_SafeTableDefName(targetTdf) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Sub
'
'Public Sub SharedDao_AddMissingIndexesFromTemplate(ByVal templateTdf As DAO.tableDef, ByVal targetTdf As DAO.tableDef)
'    'Adds template indexes that do not already have an equivalent target index.
'    'Use during schema repair when existing equivalent indexes should be preserved.
'    Dim idx As DAO.index
'    For Each idx In templateTdf.Indexes
'        If Not idx.Foreign Then
'            If Not SharedDao_EquivalentIndexExists(targetTdf, idx) Then SharedDao_AppendCopiedIndex targetTdf, idx
'        End If
'    Next idx
'    targetTdf.Indexes.Refresh
'End Sub
'
'Public Sub SharedDao_CopyTemplateTableDef(ByVal templateTdf As DAO.tableDef, ByVal dbData As DAO.Database, ByVal linkedTableName As String)
'    'Creates and verifies a new target table from template fields and removes only the table created by this call if completion fails.
'    'Use before copying indexes or template data into the new table.
'    'Consolidated old names: TVSettings_CreateRuntimeTableFromTemplate field-copy loop.
'    On Error GoTo ErrHandler
'    Dim newTdf As DAO.tableDef
'    Dim persistedTdf As DAO.tableDef
'    Dim fld As DAO.Field
'    Dim copiedFld As DAO.Field
'    Dim targetExistedBefore As Boolean
'    Dim tableAppended As Boolean
'    Dim expectedAutoNumber As Boolean
'    If templateTdf Is Nothing Then Err.Raise vbObjectError + 7930, "SharedDao_CopyTemplateTableDef", "The template TableDef argument is Nothing."
'    If dbData Is Nothing Then Err.Raise vbObjectError + 7931, "SharedDao_CopyTemplateTableDef", "The target DAO.Database argument is Nothing."
'    linkedTableName = Trim$(linkedTableName)
'    If Len(linkedTableName) = 0 Then Err.Raise vbObjectError + 7932, "SharedDao_CopyTemplateTableDef", "The target table name is blank."
'    dbData.TableDefs.Refresh
'    targetExistedBefore = SharedDao_DatabaseTableExists(dbData, linkedTableName)
'    If targetExistedBefore Then Err.Raise vbObjectError + 7933, "SharedDao_CopyTemplateTableDef", "The target table already exists and will not be replaced."
'    templateTdf.fields.Refresh
'    Set newTdf = dbData.CreateTableDef(linkedTableName)
'    For Each fld In templateTdf.fields
'        SharedDao_AppendCopiedField newTdf, fld
'    Next fld
'    dbData.TableDefs.Append newTdf
'    tableAppended = True
'    dbData.TableDefs.Refresh
'    If SharedDao_DatabaseTableExists(dbData, linkedTableName) = False Then Err.Raise vbObjectError + 7934, "SharedDao_CopyTemplateTableDef", "The target table was not present after append and refresh."
'    Set persistedTdf = dbData.TableDefs(linkedTableName)
'    persistedTdf.fields.Refresh
'    If persistedTdf.fields.count <> templateTdf.fields.count Then Err.Raise vbObjectError + 7935, "SharedDao_CopyTemplateTableDef", "The copied table field count does not match the template." & vbCrLf & "Expected fields: " & CStr(templateTdf.fields.count) & vbCrLf & "Actual fields: " & CStr(persistedTdf.fields.count)
'    For Each fld In templateTdf.fields
'        If SharedDao_TableDefFieldExists(persistedTdf, fld.Name) = False Then Err.Raise vbObjectError + 7936, "SharedDao_CopyTemplateTableDef", "A template field is missing from the copied table." & vbCrLf & "Field: " & fld.Name
'        Set copiedFld = persistedTdf.fields(fld.Name)
'        expectedAutoNumber = ((fld.attributes And dbAutoIncrField) <> 0)
'        If SharedDao_FieldDefinitionMatches(copiedFld, fld.Type, fld.Size, expectedAutoNumber, False) = False Then Err.Raise vbObjectError + 7937, "SharedDao_CopyTemplateTableDef", "A copied field definition does not match the template." & vbCrLf & "Field: " & fld.Name & vbCrLf & "Expected: " & SharedDao_ExistingFieldDefinitionText(fld) & vbCrLf & "Actual: " & SharedDao_ExistingFieldDefinitionText(copiedFld)
'        If fld.Type = dbText Or fld.Type = dbMemo Then If copiedFld.allowZeroLength <> fld.allowZeroLength Then Err.Raise vbObjectError + 7938, "SharedDao_CopyTemplateTableDef", "A copied text field AllowZeroLength value does not match the template." & vbCrLf & "Field: " & fld.Name
'    Next fld
'    Set copiedFld = Nothing
'    Set persistedTdf = Nothing
'    Set newTdf = Nothing
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim cleanupText As String
'    Dim cleanupErrNumber As Long
'    Dim cleanupErrDescription As String
'    Dim tableStillExists As Boolean
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If tableAppended And Not targetExistedBefore And Not dbData Is Nothing Then
'        Err.Clear
'        dbData.TableDefs.Refresh
'        tableStillExists = SharedDao_DatabaseTableExists(dbData, linkedTableName)
'        cleanupErrNumber = Err.Number
'        cleanupErrDescription = Err.Description
'        If cleanupErrNumber <> 0 Then
'            cleanupText = cleanupText & vbCrLf & "Created-table inspection error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription
'        ElseIf tableStillExists Then
'            Err.Clear
'            dbData.TableDefs.Delete linkedTableName
'            dbData.TableDefs.Refresh
'            cleanupErrNumber = Err.Number
'            cleanupErrDescription = Err.Description
'            If cleanupErrNumber <> 0 Then
'                cleanupText = cleanupText & vbCrLf & "Created-table cleanup error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription
'            Else
'                tableStillExists = SharedDao_DatabaseTableExists(dbData, linkedTableName)
'                cleanupErrNumber = Err.Number
'                cleanupErrDescription = Err.Description
'                If cleanupErrNumber <> 0 Then
'                    cleanupText = cleanupText & vbCrLf & "Created-table cleanup verification error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription
'                ElseIf tableStillExists Then
'                    cleanupText = cleanupText & vbCrLf & "Created-table cleanup failed verification: the table still exists."
'                Else
'                    cleanupText = cleanupText & vbCrLf & "Created table was removed and verified."
'                End If
'            End If
'        Else
'            cleanupText = cleanupText & vbCrLf & "Created table was already absent during cleanup."
'        End If
'    End If
'    Set copiedFld = Nothing
'    Set persistedTdf = Nothing
'    Set newTdf = Nothing
'    On Error GoTo 0
'    Err.Raise savedErrNumber, "SharedDao_CopyTemplateTableDef", "Could not create and verify the copied template table." & vbCrLf & "Target table: " & linkedTableName & vbCrLf & "Target database: " & SharedDao_SafeDatabaseName(dbData) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & cleanupText
'End Sub
'
'Public Function SharedDao_UniqueBackupTableName(ByVal db As DAO.Database, ByVal tableName As String, ByVal backupPrefix As String) As String
'    'Intended use: Returns a sanitized unique backup-table name that never exceeds the Access 64-character object-name limit.
'    'Consolidated old names: TVSettings_UniqueBackupTableName.
'    Dim combinedBase As String
'    Dim safeBase As String
'    Dim basePart As String
'    Dim suffixText As String
'    Dim candidate As String
'    Dim timestampText As String
'    Dim availableLength As Long
'    Dim n As Long
'    Dim preserveLeadingUnderscore As Boolean
'    If db Is Nothing Then Err.Raise vbObjectError + 8020, "SharedDao_UniqueBackupTableName", "The DAO.Database argument is Nothing."
'    If Len(tableName) = 0 Then Err.Raise vbObjectError + 8021, "SharedDao_UniqueBackupTableName", "The source table name is blank."
'    combinedBase = backupPrefix & tableName
'    preserveLeadingUnderscore = (Left$(combinedBase, 1) = "_")
'    safeBase = SharedName_SafeAccessObjectName(combinedBase, 0, "B_", "Backup")
'    If preserveLeadingUnderscore Then safeBase = "_" & safeBase
'    timestampText = Format$(Now, "yyyymmdd_hhnnss")
'    Do
'        If n = 0 Then
'            suffixText = "_" & timestampText
'        Else
'            suffixText = "_" & timestampText & "_" & CStr(n)
'        End If
'        availableLength = 64 - Len(suffixText)
'        If availableLength < 1 Then Err.Raise vbObjectError + 8022, "SharedDao_UniqueBackupTableName", "The backup suffix leaves no room for a valid Access table name."
'        basePart = Left$(safeBase, availableLength)
'        Do While Len(basePart) > 1 And Right$(basePart, 1) = "_"
'            basePart = Left$(basePart, Len(basePart) - 1)
'        Loop
'        If Len(basePart) = 0 Then basePart = "B"
'        candidate = basePart & suffixText
'        If Len(candidate) > 64 Then Err.Raise vbObjectError + 8023, "SharedDao_UniqueBackupTableName", "The generated backup table name exceeds 64 characters." & vbCrLf & "Candidate: " & candidate
'        If SharedDao_DatabaseTableExists(db, candidate) = False Then Exit Do
'        n = n + 1
'    Loop
'    SharedDao_UniqueBackupTableName = candidate
'End Function
'Public Function SharedDao_BackupTable(ByVal db As DAO.Database, ByVal tableName As String, ByVal backupPrefix As String, ByRef reportText As String, Optional ByVal reasonText As String = vbNullString) As String
'    'Creates a verified same-database data-and-index backup; relationships and Access lookup UI metadata are not copied.
'    'Consolidated old names: TVSettings_BackupRuntimeTable.
'    On Error GoTo ErrHandler
'    Dim wrk As DAO.workspace
'    Dim sourceTdf As DAO.tableDef
'    Dim backupTdf As DAO.tableDef
'    Dim backupName As String
'    Dim sqlText As String
'    Dim transactionStarted As Boolean
'    If db Is Nothing Then Err.Raise vbObjectError + 7755, "SharedDao_BackupTable", "The DAO.Database argument is Nothing."
'    If Len(tableName) = 0 Then Err.Raise vbObjectError + 7756, "SharedDao_BackupTable", "The source table name is blank."
'    If SharedDao_DatabaseTableExists(db, tableName) = False Then Err.Raise vbObjectError + 7757, "SharedDao_BackupTable", "The source table does not exist." & vbCrLf & "Table: " & tableName
'    backupName = SharedDao_UniqueBackupTableName(db, tableName, backupPrefix)
'    sqlText = "SELECT * INTO " & SharedDao_BracketName(backupName) & " FROM " & SharedDao_BracketName(tableName) & ";"
'    Set wrk = DBEngine.Workspaces(0)
'    wrk.BeginTrans
'    transactionStarted = True
'    db.Execute sqlText, dbFailOnError
'    db.TableDefs.Refresh
'    If SharedDao_DatabaseTableExists(db, backupName) = False Then Err.Raise vbObjectError + 7758, "SharedDao_BackupTable", "The backup table was not present after SELECT INTO."
'    Set sourceTdf = db.TableDefs(tableName)
'    Set backupTdf = db.TableDefs(backupName)
'    SharedDao_CopyTemplateIndexes sourceTdf, backupTdf
'    backupTdf.Indexes.Refresh
'    wrk.CommitTrans
'    transactionStarted = False
'    If Len(reasonText) > 0 Then
'        reportText = reportText & tableName & ": data-and-index backup created " & reasonText & " as " & backupName & "; relationships and lookup UI metadata were not copied." & vbCrLf
'    Else
'        reportText = reportText & tableName & ": data-and-index backup created as " & backupName & "; relationships and lookup UI metadata were not copied." & vbCrLf
'    End If
'    SharedDao_BackupTable = backupName
'    Set backupTdf = Nothing
'    Set sourceTdf = Nothing
'    Set wrk = Nothing
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim rollbackErrNumber As Long
'    Dim rollbackErrDescription As String
'    Dim cleanupRefreshErrNumber As Long
'    Dim cleanupRefreshErrDescription As String
'    Dim cleanupInspectErrNumber As Long
'    Dim cleanupInspectErrDescription As String
'    Dim cleanupDeleteErrNumber As Long
'    Dim cleanupDeleteErrDescription As String
'    Dim cleanupVerifyErrNumber As Long
'    Dim cleanupVerifyErrDescription As String
'    Dim backupExistsAfterFailure As Boolean
'    Dim backupExistsAfterCleanup As Boolean
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    Err.Clear
'    If transactionStarted And Not wrk Is Nothing Then wrk.Rollback
'    rollbackErrNumber = Err.Number
'    rollbackErrDescription = Err.Description
'    If Len(backupName) > 0 And Not db Is Nothing Then
'        Err.Clear
'        db.TableDefs.Refresh
'        cleanupRefreshErrNumber = Err.Number
'        cleanupRefreshErrDescription = Err.Description
'        If cleanupRefreshErrNumber = 0 Then
'            Err.Clear
'            backupExistsAfterFailure = SharedDao_DatabaseTableExists(db, backupName)
'            cleanupInspectErrNumber = Err.Number
'            cleanupInspectErrDescription = Err.Description
'            If cleanupInspectErrNumber = 0 And backupExistsAfterFailure Then
'                Err.Clear
'                db.TableDefs.Delete backupName
'                cleanupDeleteErrNumber = Err.Number
'                cleanupDeleteErrDescription = Err.Description
'                Err.Clear
'                db.TableDefs.Refresh
'                If cleanupDeleteErrNumber = 0 And Err.Number = 0 Then
'                    backupExistsAfterCleanup = SharedDao_DatabaseTableExists(db, backupName)
'                    cleanupVerifyErrNumber = Err.Number
'                    cleanupVerifyErrDescription = Err.Description
'                    If cleanupVerifyErrNumber = 0 And backupExistsAfterCleanup Then
'                        cleanupVerifyErrNumber = vbObjectError + 7800
'                        cleanupVerifyErrDescription = "The incomplete backup table remained after cleanup."
'                    End If
'                ElseIf cleanupDeleteErrNumber = 0 Then
'                    cleanupVerifyErrNumber = Err.Number
'                    cleanupVerifyErrDescription = Err.Description
'                End If
'            End If
'        End If
'    End If
'    Set backupTdf = Nothing
'    Set sourceTdf = Nothing
'    Set wrk = Nothing
'    On Error GoTo 0
'    If rollbackErrNumber <> 0 Then savedErrDescription = savedErrDescription & vbCrLf & "Rollback error " & CStr(rollbackErrNumber) & ": " & rollbackErrDescription
'    If cleanupRefreshErrNumber <> 0 Then savedErrDescription = savedErrDescription & vbCrLf & "Incomplete-backup refresh error " & CStr(cleanupRefreshErrNumber) & ": " & cleanupRefreshErrDescription
'    If cleanupInspectErrNumber <> 0 Then savedErrDescription = savedErrDescription & vbCrLf & "Incomplete-backup inspection error " & CStr(cleanupInspectErrNumber) & ": " & cleanupInspectErrDescription
'    If cleanupDeleteErrNumber <> 0 Then savedErrDescription = savedErrDescription & vbCrLf & "Incomplete-backup deletion error " & CStr(cleanupDeleteErrNumber) & ": " & cleanupDeleteErrDescription
'    If cleanupVerifyErrNumber <> 0 Then savedErrDescription = savedErrDescription & vbCrLf & "Incomplete-backup verification error " & CStr(cleanupVerifyErrNumber) & ": " & cleanupVerifyErrDescription
'    Err.Raise savedErrNumber, "SharedDao_BackupTable", "Could not create the same-database table backup." & vbCrLf & "Database: " & SharedDao_SafeDatabaseName(db) & vbCrLf & "Source table: " & tableName & vbCrLf & "Backup table: " & backupName & vbCrLf & "SQL: " & sqlText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'Public Function SharedDao_PrimaryKeyFields(ByVal db As DAO.Database, ByVal tableName As String) As Collection
'    'Returns primary-key field names for a table in key order.
'    'Consolidated old names: TVSettings_PrimaryKeyFields.
'    Dim result As Collection
'    Dim idx As DAO.index
'    Dim fld As DAO.Field
'    Set result = New Collection
'    If SharedDao_DatabaseTableExists(db, tableName) = False Then
'        Set SharedDao_PrimaryKeyFields = result
'        Exit Function
'    End If
'    For Each idx In db.TableDefs(tableName).Indexes
'        If idx.Primary Then
'            For Each fld In idx.fields
'                result.Add fld.Name
'            Next fld
'            Exit For
'        End If
'    Next idx
'    Set SharedDao_PrimaryKeyFields = result
'End Function
'
'Public Function SharedDao_RecordPrimaryKeyWhere(ByVal rs As DAO.Recordset, ByVal pkFields As Collection) As String
'    'Builds a WHERE expression from a recordset row and a Collection of primary-key field names.
'    'Consolidated old names: TVSettings_RecordPrimaryKeyWhere.
'    Dim i As Long
'    Dim fieldName As String
'    For i = 1 To pkFields.count
'        fieldName = CStr(pkFields(i))
'        If Len(SharedDao_RecordPrimaryKeyWhere) > 0 Then SharedDao_RecordPrimaryKeyWhere = SharedDao_RecordPrimaryKeyWhere & " AND "
'        SharedDao_RecordPrimaryKeyWhere = SharedDao_RecordPrimaryKeyWhere & SharedDao_BracketName(fieldName) & "=" & SharedDao_SqlLiteralFromValue(rs.fields(fieldName).value)
'    Next i
'End Function
'
'Public Function SharedDao_PrimaryKeyValueExists(ByVal db As DAO.Database, ByVal tableName As String, ByVal keyValues As Variant) As Boolean
'    'Intended use: Fail-closed primary-key existence validation for migration, repair, insert, delete, and other mutation decisions.
'    On Error GoTo ErrHandler
'    Dim pkFields As Collection
'    Dim rs As DAO.Recordset
'    Dim whereText As String
'    Dim i As Long
'    If db Is Nothing Then Err.Raise vbObjectError + 7620, "SharedDao_PrimaryKeyValueExists", "The DAO.Database argument is Nothing."
'    If Not IsArray(keyValues) Then Err.Raise vbObjectError + 7621, "SharedDao_PrimaryKeyValueExists", "The keyValues argument must be an array."
'    Set pkFields = SharedDao_PrimaryKeyFields(db, tableName)
'    If pkFields.count = 0 Then Err.Raise vbObjectError + 7622, "SharedDao_PrimaryKeyValueExists", "The table does not have a primary key or could not be resolved." & vbCrLf & "Table: " & tableName
'    If pkFields.count <> UBound(keyValues) - LBound(keyValues) + 1 Then Err.Raise vbObjectError + 7623, "SharedDao_PrimaryKeyValueExists", "The supplied primary-key value count does not match the table primary key." & vbCrLf & "Table: " & tableName & vbCrLf & "Expected values: " & CStr(pkFields.count) & vbCrLf & "Supplied values: " & CStr(UBound(keyValues) - LBound(keyValues) + 1)
'    For i = 1 To pkFields.count
'        If Len(whereText) > 0 Then whereText = whereText & " AND "
'        whereText = whereText & SharedDao_BracketName(CStr(pkFields(i))) & "=" & SharedDao_SqlLiteralFromValue(keyValues(LBound(keyValues) + i - 1))
'    Next i
'    Set rs = db.OpenRecordset("SELECT * FROM " & SharedDao_BracketName(tableName) & " WHERE " & whereText, dbOpenSnapshot)
'    SharedDao_PrimaryKeyValueExists = Not rs.EOF
'    rs.Close
'    Set rs = Nothing
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If Not rs Is Nothing Then rs.Close
'    Set rs = Nothing
'    On Error GoTo 0
'    Err.Raise savedErrNumber, "SharedDao_PrimaryKeyValueExists", "Unable to determine whether the primary-key value exists." & vbCrLf & "Database: " & SharedDao_SafeDatabaseName(db) & vbCrLf & "Table: " & tableName & vbCrLf & "Criteria: " & whereText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'
'Public Function SharedDao_PrimaryKeyValueExistsSafe(ByVal db As DAO.Database, ByVal tableName As String, ByVal keyValues As Variant) As Boolean
'    'Intended use: Convenience-only nonraising primary-key existence check where an indeterminate result may safely be treated as False.
'    On Error GoTo ErrHandler
'    SharedDao_PrimaryKeyValueExistsSafe = SharedDao_PrimaryKeyValueExists(db, tableName, keyValues)
'    Exit Function
'ErrHandler:
'    SharedDao_PrimaryKeyValueExistsSafe = False
'End Function
'
'Public Function SharedDao_IsAutoNumberField(ByVal fld As DAO.Field) As Boolean
'    'Returns True when a DAO field is an AutoNumber field.
'    'Consolidated old names: local TV NSS EXPORT MODULE IsAutoNumberField.
'    SharedDao_IsAutoNumberField = ((fld.attributes And dbAutoIncrField) <> 0)
'End Function
'
'Public Sub SharedDao_CurrentDbDropTableIfExists(ByVal tableName As String)
'    'Intended use: Deletes one caller-authorized CurrentDb table after saving and closing only that affected open table.
'    On Error GoTo ErrHandler
'    Dim db As DAO.Database
'    If Len(tableName) = 0 Then Err.Raise vbObjectError + 7624, "SharedDao_CurrentDbDropTableIfExists", "The table name is blank."
'    Set db = currentDb
'    db.TableDefs.Refresh
'    If Not SharedDao_DatabaseTableExists(db, tableName) Then
'        Set db = Nothing
'        Exit Sub
'    End If
'    If SharedAccess_PrepareOpenObjectForDesignChange(acTable, tableName, "deleting the caller-authorized table", True) = False Then Err.Raise vbObjectError + 7759, "SharedDao_CurrentDbDropTableIfExists", "The affected open table was not prepared for deletion."
'    DBEngine.Idle dbRefreshCache
'    db.TableDefs.Delete tableName
'    db.TableDefs.Refresh
'    DBEngine.Idle dbRefreshCache
'    If SharedDao_DatabaseTableExists(db, tableName) Then Err.Raise vbObjectError + 7625, "SharedDao_CurrentDbDropTableIfExists", "The table remained present after the delete operation." & vbCrLf & "Table: " & tableName
'    Set db = Nothing
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Set db = Nothing
'    Err.Raise savedErrNumber, "SharedDao_CurrentDbDropTableIfExists", "Unable to prepare, delete, or verify deletion of the caller-authorized CurrentDb table." & vbCrLf & "Table: " & tableName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Sub
'Public Function SharedDao_IndexFieldsExistInTable(ByVal tableName As String, ByVal idx As DAO.index) As Boolean
'    'Returns True when every field in an index exists in a CurrentDb table.
'    'Consolidated old names: local TV NSS IMPORT MODULE IndexFieldsExistInTable.
'    Dim fld As DAO.Field
'    For Each fld In idx.fields
'        If Not SharedDao_CurrentDbTableFieldExists(tableName, fld.Name) Then Exit Function
'    Next fld
'    SharedDao_IndexFieldsExistInTable = True
'End Function
'
'Public Function SharedDao_IndexFieldsExistInRecordset(ByVal rs As DAO.Recordset, ByVal idx As DAO.index) As Boolean
'    'Returns True when every field in an index exists in a recordset.
'    'Consolidated old names: local TV NSS IMPORT MODULE IndexFieldsExistInRecordset.
'    Dim fld As DAO.Field
'    For Each fld In idx.fields
'        If Not SharedDao_RecordsetFieldExists(rs, fld.Name) Then Exit Function
'    Next fld
'    SharedDao_IndexFieldsExistInRecordset = True
'End Function
'
'Public Function SharedDao_IndexRowShouldBeSkippedForNulls(ByVal rs As DAO.Recordset, ByVal idx As DAO.index) As Boolean
'    'Returns True when an IgnoreNulls index should skip the current recordset row.
'    'Consolidated old names: local TV NSS IMPORT MODULE IndexRowShouldBeSkippedForNulls.
'    Dim fld As DAO.Field
'    If Not idx.ignoreNulls Then Exit Function
'    For Each fld In idx.fields
'        If IsNull(rs.fields(fld.Name).value) Then
'            SharedDao_IndexRowShouldBeSkippedForNulls = True
'            Exit Function
'        End If
'    Next fld
'End Function
'
'Public Function SharedDao_CriteriaFromIndexFields(ByVal rs As DAO.Recordset, ByVal idx As DAO.index) As String
'    'Builds an Access WHERE criteria expression from the fields in an index and the current recordset row.
'    'Consolidated old names: local TV NSS IMPORT MODULE CriteriaFromIndexFields.
'    Dim fld As DAO.Field
'    Dim s As String
'    For Each fld In idx.fields
'        If Len(s) > 0 Then s = s & " AND "
'        If IsNull(rs.fields(fld.Name).value) Then
'            s = s & SharedDao_BracketName(fld.Name) & " Is Null"
'        Else
'            s = s & SharedDao_BracketName(fld.Name) & "=" & SharedDao_SqlLiteralFromValue(rs.fields(fld.Name).value)
'        End If
'    Next fld
'    SharedDao_CriteriaFromIndexFields = s
'End Function
'
'Public Function SharedDao_IndexFieldList(ByVal idx As DAO.index) As String
'    'Returns a comma-separated list of field names in an index.
'    'Consolidated old names: local TV NSS IMPORT MODULE IndexFieldList.
'    Dim fld As DAO.Field
'    Dim s As String
'    For Each fld In idx.fields
'        If Len(s) > 0 Then s = s & ", "
'        s = s & fld.Name
'    Next fld
'    SharedDao_IndexFieldList = s
'End Function
'
'Public Function SharedDao_IndexValueList(ByVal rs As DAO.Recordset, ByVal idx As DAO.index) As String
'    'Returns a semicolon-separated field=value list for index fields in the current recordset row.
'    'Consolidated old names: local TV NSS IMPORT MODULE IndexValueList.
'    Dim fld As DAO.Field
'    Dim s As String
'    For Each fld In idx.fields
'        If Len(s) > 0 Then s = s & "; "
'        s = s & fld.Name & "=" & Nz(rs.fields(fld.Name).value, "<Null>")
'    Next fld
'    SharedDao_IndexValueList = s
'End Function
'
'Private Sub SharedDao_AppendFieldVerified(ByVal tdf As DAO.tableDef, ByVal fieldName As String, ByVal fieldType As DAO.DataTypeEnum, ByVal fieldSize As Long, ByVal typeDisplayName As String, ByVal callerName As String, ByVal nothingErrorValue As Long, ByVal blankErrorValue As Long, ByVal invalidSizeErrorValue As Long, ByVal missingErrorValue As Long, ByVal mismatchErrorValue As Long)
'    'Creates, appends, verifies, and conditionally cleans up one caller-owned DAO field mutation.
'    On Error GoTo ErrHandler
'    Dim fld As DAO.Field
'    Dim fieldStateCaptured As Boolean
'    Dim fieldExistedBefore As Boolean
'    Dim appendAttempted As Boolean
'    Dim appended As Boolean
'    If tdf Is Nothing Then Err.Raise nothingErrorValue, callerName, "The DAO.TableDef argument is Nothing."
'    fieldName = Trim$(fieldName)
'    If Len(fieldName) = 0 Then Err.Raise blankErrorValue, callerName, "The field name is blank."
'    If fieldType = dbText Then
'        If fieldSize < 1 Or fieldSize > 255 Then Err.Raise invalidSizeErrorValue, callerName, "The Text field size must be between 1 and 255." & vbCrLf & "Field: " & fieldName & vbCrLf & "Size: " & CStr(fieldSize)
'    End If
'    fieldExistedBefore = SharedDao_TableDefFieldExists(tdf, fieldName)
'    fieldStateCaptured = True
'    If fieldType = dbText Then
'        Set fld = tdf.CreateField(fieldName, fieldType, fieldSize)
'    Else
'        Set fld = tdf.CreateField(fieldName, fieldType)
'    End If
'    appendAttempted = True
'    tdf.fields.Append fld
'    appended = True
'    tdf.fields.Refresh
'    If SharedDao_TableDefFieldExists(tdf, fieldName) = False Then Err.Raise missingErrorValue, callerName, "The " & typeDisplayName & " field is missing after append and refresh."
'    If SharedDao_FieldDefinitionMatches(tdf.fields(fieldName), fieldType, fieldSize, False, False) = False Then Err.Raise mismatchErrorValue, callerName, "The appended " & typeDisplayName & " field definition does not match the request."
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim cleanupText As String
'    Dim cleanupErrNumber As Long
'    Dim cleanupErrDescription As String
'    Dim fieldStillExists As Boolean
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If appendAttempted And fieldStateCaptured And fieldExistedBefore = False And Not tdf Is Nothing And Len(fieldName) > 0 Then
'        Err.Clear
'        fieldStillExists = SharedDao_TableDefFieldExists(tdf, fieldName)
'        cleanupErrNumber = Err.Number
'        cleanupErrDescription = Err.Description
'        If cleanupErrNumber = 0 And fieldStillExists Then
'            Err.Clear
'            tdf.fields.Delete fieldName
'            tdf.fields.Refresh
'            cleanupErrNumber = Err.Number
'            cleanupErrDescription = Err.Description
'        End If
'        If cleanupErrNumber = 0 Then
'            Err.Clear
'            fieldStillExists = SharedDao_TableDefFieldExists(tdf, fieldName)
'            cleanupErrNumber = Err.Number
'            cleanupErrDescription = Err.Description
'        End If
'        If cleanupErrNumber <> 0 Then
'            cleanupText = vbCrLf & "Cleanup or verification error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription
'        ElseIf fieldStillExists Then
'            cleanupText = vbCrLf & "Cleanup did not remove the failed field."
'        Else
'            cleanupText = vbCrLf & "Any field persisted by the failed append was removed and verified absent."
'        End If
'    ElseIf appendAttempted And fieldStateCaptured And fieldExistedBefore Then
'        cleanupText = vbCrLf & "No field cleanup was attempted because the target field existed before this invocation."
'    Else
'        cleanupText = vbCrLf & "No field cleanup was attempted because no owned append mutation began."
'    End If
'    On Error GoTo 0
'    cleanupText = cleanupText & vbCrLf & "Mutation state: prior state captured=" & CStr(fieldStateCaptured) & "; existed before=" & CStr(fieldExistedBefore) & "; append attempted=" & CStr(appendAttempted) & "; append returned=" & CStr(appended)
'    Err.Raise savedErrNumber, callerName, "Could not append and verify the " & typeDisplayName & " field." & vbCrLf & "Table: " & SharedDao_SafeTableDefName(tdf) & vbCrLf & "Field: " & fieldName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & cleanupText
'End Sub
'
'Public Sub SharedDao_AppendTextField(ByVal tdf As DAO.tableDef, ByVal fieldName As String, ByVal fieldSize As Long)
'    'Intended use: Creates, appends, and verifies a Text field in a DAO TableDef.
'    'Consolidated old names: AddTextField.
'    SharedDao_AppendFieldVerified tdf, fieldName, dbText, fieldSize, "Text", "SharedDao_AppendTextField", vbObjectError + 8024, vbObjectError + 8025, vbObjectError + 8026, vbObjectError + 8027, vbObjectError + 8028
'End Sub
'
'Public Sub SharedDao_AppendLongTextField(ByVal tdf As DAO.tableDef, ByVal fieldName As String)
'    'Intended use: Creates, appends, and verifies a Long Text field in a DAO TableDef.
'    'Consolidated old names: AddLongTextField.
'    SharedDao_AppendFieldVerified tdf, fieldName, dbMemo, 0, "Long Text", "SharedDao_AppendLongTextField", vbObjectError + 8029, vbObjectError + 8030, 0, vbObjectError + 8031, vbObjectError + 8032
'End Sub
'
'Public Sub SharedDao_AppendYesNoField(ByVal tdf As DAO.tableDef, ByVal fieldName As String)
'    'Intended use: Creates, appends, and verifies a Yes/No field in a DAO TableDef.
'    'Consolidated old names: AddYesNoField.
'    SharedDao_AppendFieldVerified tdf, fieldName, dbBoolean, 0, "Yes/No", "SharedDao_AppendYesNoField", vbObjectError + 8033, vbObjectError + 8034, 0, vbObjectError + 8035, vbObjectError + 8036
'End Sub
'
'Public Function SharedDao_NextSequenceLong(ByVal sequenceTableName As String, ByVal sequenceNameField As String, ByVal sequenceValueField As String, ByVal sequenceName As String) As Long
'    'Atomically increments one Long sequence row and returns the value that was allocated before the increment.
'    'Consolidated old names: AssignNextRef sequence-table logic.
'    On Error GoTo ErrorHandler
'    Dim db As DAO.Database
'    Dim wrk As DAO.workspace
'    Dim rs As DAO.Recordset
'    Dim updateSqlText As String
'    Dim selectSqlText As String
'    Dim affectedRows As Long
'    Dim updatedValue As Long
'    Dim transactionStarted As Boolean
'    sequenceTableName = Trim$(sequenceTableName)
'    sequenceNameField = Trim$(sequenceNameField)
'    sequenceValueField = Trim$(sequenceValueField)
'    sequenceName = Trim$(sequenceName)
'    If Len(sequenceTableName) = 0 Or Len(sequenceNameField) = 0 Or Len(sequenceValueField) = 0 Or Len(sequenceName) = 0 Then Err.Raise vbObjectError + 7760, "SharedDao_NextSequenceLong", "Sequence table, field, and name arguments must not be blank."
'    Set wrk = DBEngine.Workspaces(0)
'    Set db = currentDb
'    updateSqlText = "UPDATE " & SharedDao_BracketName(sequenceTableName) & " SET " & SharedDao_BracketName(sequenceValueField) & "=" & SharedDao_BracketName(sequenceValueField) & "+1 WHERE " & SharedDao_BracketName(sequenceNameField) & "=" & SharedDao_SqlQuote(sequenceName) & " AND " & SharedDao_BracketName(sequenceValueField) & " Is Not Null AND " & SharedDao_BracketName(sequenceValueField) & "<2147483647;"
'    selectSqlText = "SELECT " & SharedDao_BracketName(sequenceValueField) & " FROM " & SharedDao_BracketName(sequenceTableName) & " WHERE " & SharedDao_BracketName(sequenceNameField) & "=" & SharedDao_SqlQuote(sequenceName) & ";"
'    wrk.BeginTrans
'    transactionStarted = True
'    db.Execute updateSqlText, dbFailOnError
'    affectedRows = db.RecordsAffected
'    If affectedRows <> 1 Then Err.Raise vbObjectError + 7761, "SharedDao_NextSequenceLong", "The atomic sequence update did not affect exactly one row. The sequence may be missing, duplicated, Null, or exhausted." & vbCrLf & "Affected rows: " & CStr(affectedRows)
'    Set rs = db.OpenRecordset(selectSqlText, dbOpenSnapshot)
'    If rs.EOF Then Err.Raise vbObjectError + 7762, "SharedDao_NextSequenceLong", "The updated sequence row could not be read back."
'    updatedValue = CLng(rs.fields(sequenceValueField).value)
'    rs.MoveNext
'    If Not rs.EOF Then Err.Raise vbObjectError + 7763, "SharedDao_NextSequenceLong", "More than one sequence row matched after the update."
'    If updatedValue <= -2147483647 Then Err.Raise vbObjectError + 7764, "SharedDao_NextSequenceLong", "The allocated sequence value is outside the supported positive Long range."
'    SharedDao_NextSequenceLong = updatedValue - 1
'    If SharedDao_NextSequenceLong < 0 Then Err.Raise vbObjectError + 7765, "SharedDao_NextSequenceLong", "The allocated sequence value is negative."
'    rs.Close
'    Set rs = Nothing
'    wrk.CommitTrans
'    transactionStarted = False
'    Set db = Nothing
'    Set wrk = Nothing
'    Exit Function
'ErrorHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim cleanupErrNumber As Long
'    Dim cleanupErrDescription As String
'    Dim rollbackErrNumber As Long
'    Dim rollbackErrDescription As String
'    Dim finalDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    Err.Clear
'    If Not rs Is Nothing Then rs.Close
'    cleanupErrNumber = Err.Number
'    cleanupErrDescription = Err.Description
'    Err.Clear
'    If transactionStarted And Not wrk Is Nothing Then wrk.Rollback
'    rollbackErrNumber = Err.Number
'    rollbackErrDescription = Err.Description
'    Set rs = Nothing
'    Set db = Nothing
'    Set wrk = Nothing
'    On Error GoTo 0
'    finalDescription = savedErrDescription
'    If cleanupErrNumber <> 0 Then finalDescription = finalDescription & vbCrLf & "Recordset cleanup error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription
'    If rollbackErrNumber <> 0 Then finalDescription = finalDescription & vbCrLf & "Rollback error " & CStr(rollbackErrNumber) & ": " & rollbackErrDescription
'    Err.Raise savedErrNumber, "SharedDao_NextSequenceLong", "Unable to allocate the next sequence value." & vbCrLf & "Sequence table: " & sequenceTableName & vbCrLf & "Sequence name field: " & sequenceNameField & vbCrLf & "Sequence value field: " & sequenceValueField & vbCrLf & "Sequence name: " & sequenceName & vbCrLf & "Affected rows before rollback: " & CStr(affectedRows) & vbCrLf & "Update SQL: " & updateSqlText & vbCrLf & "Read-back SQL: " & selectSqlText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & finalDescription
'End Function
'
'Public Function SharedDao_TryDatabaseTableExists(ByVal db As DAO.Database, ByVal tableName As String, ByRef tableExists As Boolean, ByRef errorNumber As Long, ByRef errorDescription As String) As Boolean
'    'Intended use: Distinguishes a successful table-existence inspection from a technical inspection failure.
'    On Error GoTo ErrHandler
'    Dim savedErrorNumber As Long
'    Dim savedErrorDescription As String
'    Dim databaseName As String
'    tableExists = False
'    errorNumber = 0
'    errorDescription = vbNullString
'    If db Is Nothing Then Err.Raise vbObjectError + 7951, "SharedDao_TryDatabaseTableExists", "The DAO.Database reference is Nothing while inspecting table '" & tableName & "'."
'    db.TableDefs.Refresh
'    tableExists = SharedDao_DatabaseTableExists(db, tableName)
'    SharedDao_TryDatabaseTableExists = True
'    Exit Function
'ErrHandler:
'    savedErrorNumber = Err.Number
'    savedErrorDescription = Err.Description
'    databaseName = "[Unavailable]"
'    On Error Resume Next
'    databaseName = SharedDao_SafeDatabaseName(db)
'    On Error GoTo 0
'    errorNumber = savedErrorNumber
'    errorDescription = "Could not inspect table existence. Database='" & databaseName & "', Table='" & tableName & "'. Error " & savedErrorNumber & ": " & savedErrorDescription
'    tableExists = False
'    SharedDao_TryDatabaseTableExists = False
'End Function
'
'Public Function SharedDao_DatabaseTableExistsSafe(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
'    'Intended use: Provides a nonraising table-existence check for optional logic where an inspection failure may be treated as False.
'    On Error GoTo ErrHandler
'    If db Is Nothing Then Exit Function
'    db.TableDefs.Refresh
'    SharedDao_DatabaseTableExistsSafe = SharedDao_DatabaseTableExists(db, tableName)
'    Exit Function
'ErrHandler:
'    SharedDao_DatabaseTableExistsSafe = False
'End Function
'
'Public Function SharedDao_TableDefFieldExistsSafe(ByVal tdf As DAO.tableDef, ByVal fieldName As String) As Boolean
'    'Intended use: Provides a nonraising TableDef field-existence check for optional logic where an inspection failure may be treated as False.
'    On Error GoTo ErrHandler
'    If tdf Is Nothing Then Exit Function
'    tdf.fields.Refresh
'    SharedDao_TableDefFieldExistsSafe = SharedDao_TableDefFieldExists(tdf, fieldName)
'    Exit Function
'ErrHandler:
'    SharedDao_TableDefFieldExistsSafe = False
'End Function
'
'Public Function SharedDao_DatabaseTableFieldExistsSafe(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String) As Boolean
'    'Intended use: Provides a nonraising database table-field check for optional logic where an inspection failure may be treated as False.
'    On Error GoTo ErrHandler
'    If SharedDao_DatabaseTableExistsSafe(db, tableName) = False Then Exit Function
'    SharedDao_DatabaseTableFieldExistsSafe = SharedDao_TableDefFieldExistsSafe(db.TableDefs(tableName), fieldName)
'    Exit Function
'ErrHandler:
'    SharedDao_DatabaseTableFieldExistsSafe = False
'End Function
'
'Public Function SharedDao_TableDefIndexExistsSafe(ByVal tdf As DAO.tableDef, ByVal indexName As String) As Boolean
'    'Intended use: Provides a nonraising TableDef index-existence check for optional logic where an inspection failure may be treated as False.
'    On Error GoTo ErrHandler
'    If tdf Is Nothing Then Exit Function
'    tdf.Indexes.Refresh
'    SharedDao_TableDefIndexExistsSafe = SharedDao_TableDefIndexExists(tdf, indexName)
'    Exit Function
'ErrHandler:
'    SharedDao_TableDefIndexExistsSafe = False
'End Function
'
'Public Function SharedDao_TableDefSingleFieldIndexMatches(ByVal tdf As DAO.tableDef, ByVal indexName As String, ByVal fieldName As String, ByVal uniqueIndex As Boolean, ByVal primaryIndex As Boolean) As Boolean
'    'Intended use: Fail-closed schema validation of one named single-field index before repair or migration decisions.
'    If tdf Is Nothing Then Err.Raise vbObjectError + 7627, "SharedDao_TableDefSingleFieldIndexMatches", "The DAO.TableDef argument is Nothing."
'    tdf.Indexes.Refresh
'    If Not SharedDao_TableDefIndexExists(tdf, indexName) Then Exit Function
'    SharedDao_TableDefSingleFieldIndexMatches = SharedDao_IndexMatchesFields(tdf.Indexes(indexName), Array(fieldName), uniqueIndex, primaryIndex)
'End Function
'
'Public Function SharedDao_EnsureDatabaseTableField(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String, ByVal fieldType As DAO.DataTypeEnum, Optional ByVal fieldSize As Long = 0, Optional ByVal autoNumber As Boolean = False, Optional ByVal allowZeroLength As Boolean = False, Optional ByVal descriptionText As String = vbNullString, Optional ByVal allowLargerTextSize As Boolean = False) As Boolean
'    'Public shared API for database-explicit callers that need to create or verify one DAO table field.
'    On Error GoTo ErrHandler
'    SharedDao_EnsureDatabaseTableField = SharedDao_EnsureFieldCore(db, tableName, fieldName, fieldType, fieldSize, autoNumber, allowZeroLength, descriptionText, allowLargerTextSize, "SharedDao_EnsureDatabaseTableField", vbObjectError + 7322, vbObjectError + 7323)
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedDao_EnsureDatabaseTableField", "Unable to create or verify table field." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName & vbCrLf & "Expected: " & SharedDao_FieldDefinitionText(fieldType, fieldSize, autoNumber) & vbCrLf & "Database: " & SharedDao_SafeDatabaseName(db) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'
'Private Function SharedDao_EnsureFieldCore(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String, ByVal fieldType As DAO.DataTypeEnum, ByVal fieldSize As Long, ByVal autoNumber As Boolean, ByVal allowZeroLength As Boolean, ByVal descriptionText As String, ByVal allowLargerTextSize As Boolean, ByVal callerName As String, ByVal missingTableErrorNumber As Long, ByVal mismatchErrorNumber As Long) As Boolean
'    'Implements common field repair, verifies the result, and restores the captured pre-operation field state if later work fails.
'    On Error GoTo ErrHandler
'    Dim tdf As DAO.tableDef
'    Dim fld As DAO.Field
'    Dim changed As Boolean
'    Dim preStateCaptured As Boolean
'    Dim fieldExistedBefore As Boolean
'    Dim fieldAppended As Boolean
'    Dim allowZeroLengthCaptured As Boolean
'    Dim allowZeroLengthBefore As Boolean
'    Dim descriptionStateCaptured As Boolean
'    Dim descriptionExistedBefore As Boolean
'    Dim descriptionValueBefore As Variant
'    If db Is Nothing Then Err.Raise vbObjectError + 7802, callerName, "The DAO.Database argument is Nothing."
'    If SharedDao_DatabaseTableExists(db, tableName) = False Then Err.Raise missingTableErrorNumber, callerName, "Table is missing before field repair." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName
'    Set tdf = db.TableDefs(tableName)
'    tdf.fields.Refresh
'    fieldExistedBefore = SharedDao_TableDefFieldExists(tdf, fieldName)
'    If fieldExistedBefore Then
'        Set fld = tdf.fields(fieldName)
'        If fieldType = dbText Then
'            allowZeroLengthBefore = fld.allowZeroLength
'            allowZeroLengthCaptured = True
'        End If
'        If Len(descriptionText) > 0 Then
'            SharedDao_GetPropertyState fld, "Description", descriptionExistedBefore, descriptionValueBefore
'            descriptionStateCaptured = True
'        End If
'    End If
'    preStateCaptured = True
'    If fieldExistedBefore Then
'        If SharedDao_FieldDefinitionMatches(fld, fieldType, fieldSize, autoNumber, allowLargerTextSize) = False Then Err.Raise mismatchErrorNumber, callerName, "Existing field has an incompatible definition." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName & vbCrLf & "Expected: " & SharedDao_FieldDefinitionText(fieldType, fieldSize, autoNumber) & vbCrLf & "Actual: " & SharedDao_ExistingFieldDefinitionText(fld)
'        If fieldType = dbText Then
'            If fld.allowZeroLength <> allowZeroLength Then
'                fld.allowZeroLength = allowZeroLength
'                tdf.fields.Refresh
'                Set fld = tdf.fields(fieldName)
'                If fld.allowZeroLength <> allowZeroLength Then Err.Raise mismatchErrorNumber, callerName, "Existing text field AllowZeroLength could not be repaired and verified." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName & vbCrLf & "Expected AllowZeroLength: " & CStr(allowZeroLength) & vbCrLf & "Actual AllowZeroLength: " & CStr(fld.allowZeroLength)
'                changed = True
'            End If
'        End If
'        If Len(descriptionText) > 0 Then If SharedDao_EnsurePropertyValue(fld, "Description", dbText, descriptionText, callerName) Then changed = True
'        SharedDao_EnsureFieldCore = changed
'        Exit Function
'    End If
'    If fieldType = dbText And fieldSize > 0 Then
'        Set fld = tdf.CreateField(fieldName, fieldType, fieldSize)
'        fld.allowZeroLength = allowZeroLength
'    Else
'        Set fld = tdf.CreateField(fieldName, fieldType)
'    End If
'    If autoNumber Then fld.attributes = fld.attributes Or dbAutoIncrField
'    tdf.fields.Append fld
'    fieldAppended = True
'    tdf.fields.Refresh
'    Set fld = tdf.fields(fieldName)
'    If SharedDao_FieldDefinitionMatches(fld, fieldType, fieldSize, autoNumber, allowLargerTextSize) = False Then Err.Raise mismatchErrorNumber, callerName, "New field was appended but its definition could not be verified." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName & vbCrLf & "Expected: " & SharedDao_FieldDefinitionText(fieldType, fieldSize, autoNumber) & vbCrLf & "Actual: " & SharedDao_ExistingFieldDefinitionText(fld)
'    If fieldType = dbText Then If fld.allowZeroLength <> allowZeroLength Then Err.Raise mismatchErrorNumber, callerName, "New text field AllowZeroLength does not match the requested setting." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName & vbCrLf & "Expected AllowZeroLength: " & CStr(allowZeroLength) & vbCrLf & "Actual AllowZeroLength: " & CStr(fld.allowZeroLength)
'    If Len(descriptionText) > 0 Then SharedDao_EnsurePropertyValue fld, "Description", dbText, descriptionText, callerName
'    SharedDao_EnsureFieldCore = True
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim cleanupText As String
'    Dim cleanupErrNumber As Long
'    Dim cleanupErrDescription As String
'    Dim fieldStillExists As Boolean
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If Not db Is Nothing And preStateCaptured Then
'        Err.Clear
'        db.TableDefs.Refresh
'        Set tdf = db.TableDefs(tableName)
'        tdf.fields.Refresh
'        fieldStillExists = SharedDao_TableDefFieldExists(tdf, fieldName)
'        cleanupErrNumber = Err.Number
'        cleanupErrDescription = Err.Description
'        If cleanupErrNumber <> 0 Then
'            cleanupText = cleanupText & vbCrLf & "Field cleanup inspection error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription
'        ElseIf fieldExistedBefore = False Then
'            If fieldStillExists And fieldAppended Then
'                Err.Clear
'                tdf.fields.Delete fieldName
'                tdf.fields.Refresh
'                cleanupErrNumber = Err.Number
'                cleanupErrDescription = Err.Description
'                If cleanupErrNumber <> 0 Then
'                    cleanupText = cleanupText & vbCrLf & "Created-field cleanup error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription
'                Else
'                    fieldStillExists = SharedDao_TableDefFieldExists(tdf, fieldName)
'                    cleanupErrNumber = Err.Number
'                    cleanupErrDescription = Err.Description
'                    If cleanupErrNumber <> 0 Then
'                        cleanupText = cleanupText & vbCrLf & "Created-field cleanup verification error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription
'                    ElseIf fieldStillExists Then
'                        cleanupText = cleanupText & vbCrLf & "Created-field cleanup failed verification: the field still exists."
'                    Else
'                        cleanupText = cleanupText & vbCrLf & "Created field was removed and verified."
'                    End If
'                End If
'            ElseIf fieldStillExists Then
'                cleanupText = cleanupText & vbCrLf & "A field not present in the captured pre-state exists, but this call did not confirm that it appended the field; it was not deleted."
'            End If
'        ElseIf fieldStillExists Then
'            Set fld = tdf.fields(fieldName)
'            If allowZeroLengthCaptured Then
'                Err.Clear
'                fld.allowZeroLength = allowZeroLengthBefore
'                tdf.fields.Refresh
'                Set fld = tdf.fields(fieldName)
'                cleanupErrNumber = Err.Number
'                cleanupErrDescription = Err.Description
'                If cleanupErrNumber <> 0 Then
'                    cleanupText = cleanupText & vbCrLf & "AllowZeroLength restore error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription
'                ElseIf fld.allowZeroLength <> allowZeroLengthBefore Then
'                    cleanupText = cleanupText & vbCrLf & "AllowZeroLength restore failed verification."
'                Else
'                    cleanupText = cleanupText & vbCrLf & "AllowZeroLength was restored and verified."
'                End If
'            End If
'            If descriptionStateCaptured Then
'                Err.Clear
'                SharedDao_RestorePropertyState fld, "Description", dbText, descriptionExistedBefore, descriptionValueBefore
'                cleanupErrNumber = Err.Number
'                cleanupErrDescription = Err.Description
'                If cleanupErrNumber <> 0 Then
'                    cleanupText = cleanupText & vbCrLf & "Description restore error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription
'                Else
'                    cleanupText = cleanupText & vbCrLf & "Description was restored and verified."
'                End If
'            End If
'        Else
'            cleanupText = cleanupText & vbCrLf & "The pre-existing field is missing during cleanup; its prior state could not be restored."
'        End If
'    ElseIf Not db Is Nothing Then
'        cleanupText = cleanupText & vbCrLf & "No field cleanup was attempted because the pre-operation field state was not captured conclusively."
'    End If
'    Set fld = Nothing
'    Set tdf = Nothing
'    On Error GoTo 0
'    Err.Raise savedErrNumber, callerName, "Unable to create or verify the field with rollback-safe cleanup." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName & vbCrLf & "Expected: " & SharedDao_FieldDefinitionText(fieldType, fieldSize, autoNumber) & vbCrLf & "Database: " & SharedDao_SafeDatabaseName(db) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & cleanupText
'End Function
'' ================================================================================================================================================================================
'' FORM DESIGN
'' ================================================================================================================================================================================
'Public Function SharedAccess_ObjectIsOpen(ByVal objectType As AcObjectType, ByVal objectName As String) As Boolean
'    'Intended use: Convenience-only nonraising open-state check for optional UI behavior where an indeterminate result may safely be treated as False.
'    On Error GoTo ErrHandler
'    SharedAccess_ObjectIsOpen = SharedAccess_ObjectIsOpenStrict(objectType, objectName)
'    Exit Function
'ErrHandler:
'    SharedAccess_ObjectIsOpen = False
'End Function
'
'Public Function SharedAccess_ObjectIsOpenStrict(ByVal objectType As AcObjectType, ByVal objectName As String) As Boolean
'    'Intended use: Fail-closed open-state inspection before design changes, replacement, deletion, save, close, or closure verification.
'    On Error GoTo ErrHandler
'    If Len(objectName) = 0 Then Err.Raise vbObjectError + 7626, "SharedAccess_ObjectIsOpenStrict", "The object name is blank."
'    SharedAccess_ObjectIsOpenStrict = ((SysCmd(acSysCmdGetObjectState, objectType, objectName) And acObjStateOpen) <> 0)
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedAccess_ObjectIsOpenStrict", "Unable to determine whether the Access object is open." & vbCrLf & "Object type: " & CStr(objectType) & vbCrLf & "Object: " & objectName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'Public Function SharedAccess_PrepareOpenObjectsForDesignChange(ByVal objectType As AcObjectType, ByVal objectNames As Collection, ByVal actionText As String, Optional ByVal confirmationAlreadyProvided As Boolean = False) As Boolean
'    'Identifies affected open objects, optionally confirms, saves and closes them, verifies closure, and restores already prepared objects if a later preparation step fails.
'    On Error GoTo ErrHandler
'    Dim openNames As Collection
'    Dim restoreStates As Collection
'    Dim currentState As Object
'    Dim item As Variant
'    Dim objectName As String
'    Dim promptText As String
'    Dim response As VbMsgBoxResult
'    Set openNames = New Collection
'    Set restoreStates = New Collection
'    If objectNames Is Nothing Then
'        SharedAccess_PrepareOpenObjectsForDesignChange = True
'        Exit Function
'    End If
'    For Each item In objectNames
'        objectName = CStr(item)
'        If Len(objectName) > 0 Then
'            If SharedAccess_ObjectIsOpenStrict(objectType, objectName) Then
'                On Error Resume Next
'                openNames.Add objectName, UCase$(objectName)
'                On Error GoTo ErrHandler
'            End If
'        End If
'    Next item
'    If openNames.count = 0 Then
'        SharedAccess_PrepareOpenObjectsForDesignChange = True
'        Exit Function
'    End If
'    If confirmationAlreadyProvided = False Then
'        promptText = "The following affected object(s) are open and must be saved and closed before " & Trim$(actionText) & ":" & vbCrLf & vbCrLf & SharedCollection_Join(openNames, vbCrLf) & vbCrLf & vbCrLf & "Save pending changes and close only these affected objects?"
'        response = MsgBox(promptText, vbQuestion Or vbYesNo Or vbDefaultButton2, "Prepare Open Access Objects")
'        If response <> vbYes Then Exit Function
'    End If
'    For Each item In openNames
'        objectName = CStr(item)
'        Set currentState = SharedAccess_CaptureOpenObjectRestoreState(objectType, objectName)
'        restoreStates.Add currentState
'        If objectType = acForm Then SharedAccess_SaveDirtyFormRecord objectName
'        DoCmd.Save objectType, objectName
'        DoCmd.Close objectType, objectName, acSaveYes
'        If SharedAccess_ObjectIsOpenStrict(objectType, objectName) Then Err.Raise vbObjectError + 7612, "SharedAccess_PrepareOpenObjectsForDesignChange", "The object remained open after save and close." & vbCrLf & "Object: " & objectName
'        Set currentState = Nothing
'    Next item
'    SharedAccess_PrepareOpenObjectsForDesignChange = True
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim restorationText As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    restorationText = SharedAccess_RestorePreparedOpenObjects(objectType, restoreStates)
'    Err.Raise savedErrNumber, "SharedAccess_PrepareOpenObjectsForDesignChange", "Could not prepare affected open Access objects for a design change." & vbCrLf & "Action: " & actionText & vbCrLf & "Object type: " & CStr(objectType) & vbCrLf & "Object: " & objectName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & IIf(Len(restorationText) > 0, vbCrLf & "Application-state restoration:" & vbCrLf & restorationText, vbNullString)
'End Function
'Private Function SharedAccess_CaptureOpenObjectRestoreState(ByVal objectType As AcObjectType, ByVal objectName As String) As Object
'    'Captures the minimum state needed to reopen one object if a later design-preparation step fails.
'    Dim state As Object
'    Dim currentViewValue As Long
'    Set state = CreateObject("Scripting.Dictionary")
'    state.CompareMode = vbTextCompare
'    state.Add "Name", objectName
'    currentViewValue = -1
'    On Error Resume Next
'    If objectType = acForm Then currentViewValue = forms(objectName).CurrentView
'    If objectType = acReport Then currentViewValue = reports(objectName).CurrentView
'    On Error GoTo 0
'    state.Add "CurrentView", currentViewValue
'    Set SharedAccess_CaptureOpenObjectRestoreState = state
'End Function
'
'Private Function SharedAccess_RestorePreparedOpenObjects(ByVal objectType As AcObjectType, ByVal restoreStates As Collection) As String
'    'Best-effort restoration used only after design preparation fails; reports every reopen or verification failure.
'    Dim i As Long
'    Dim state As Object
'    Dim objectName As String
'    Dim currentViewValue As Long
'    Dim detailText As String
'    Dim objectIsOpen As Boolean
'    If restoreStates Is Nothing Then Exit Function
'    For i = restoreStates.count To 1 Step -1
'        Set state = restoreStates(i)
'        objectName = CStr(state("Name"))
'        currentViewValue = CLng(state("CurrentView"))
'        On Error Resume Next
'        Err.Clear
'        objectIsOpen = SharedAccess_ObjectIsOpenStrict(objectType, objectName)
'        If Err.Number <> 0 Then
'            detailText = SharedText_AppendNonBlankLine(detailText, objectName & " open-state inspection failed with error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'        ElseIf objectIsOpen Then
'            detailText = SharedText_AppendNonBlankLine(detailText, objectName & " was already open; no reopen was required.")
'        Else
'            Err.Clear
'            SharedAccess_ReopenPreparedObject objectType, objectName, currentViewValue
'            If Err.Number <> 0 Then
'                detailText = SharedText_AppendNonBlankLine(detailText, objectName & " reopen failed with error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'            Else
'                Err.Clear
'                objectIsOpen = SharedAccess_ObjectIsOpenStrict(objectType, objectName)
'                If Err.Number <> 0 Then
'                    detailText = SharedText_AppendNonBlankLine(detailText, objectName & " reopen verification failed with error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'                ElseIf objectIsOpen = False Then
'                    detailText = SharedText_AppendNonBlankLine(detailText, objectName & " remained closed after the reopen attempt.")
'                Else
'                    detailText = SharedText_AppendNonBlankLine(detailText, objectName & " was reopened and verified.")
'                End If
'            End If
'        End If
'        On Error GoTo 0
'    Next i
'    SharedAccess_RestorePreparedOpenObjects = detailText
'End Function
'
'Private Sub SharedAccess_ReopenPreparedObject(ByVal objectType As AcObjectType, ByVal objectName As String, ByVal currentViewValue As Long)
'    'Reopens one previously open object without executing a macro as part of restoration.
'    Select Case objectType
'        Case acForm
'            Select Case currentViewValue
'                Case 0
'                    DoCmd.OpenForm objectName, acDesign
'                Case 2
'                    DoCmd.OpenForm objectName, acFormDS
'                Case 7
'                    DoCmd.OpenForm objectName, acLayout
'                Case Else
'                    DoCmd.OpenForm objectName, acNormal
'            End Select
'        Case acReport
'            Select Case currentViewValue
'                Case 0
'                    DoCmd.OpenReport objectName, acViewDesign
'                Case 7
'                    DoCmd.OpenReport objectName, acViewLayout
'                Case 6
'                    DoCmd.OpenReport objectName, acViewReport
'                Case Else
'                    DoCmd.OpenReport objectName, acViewPreview
'            End Select
'        Case acQuery
'            DoCmd.OpenQuery objectName, acViewNormal
'        Case acTable
'            DoCmd.OpenTable objectName, acViewNormal
'        Case acModule
'            DoCmd.OpenModule objectName
'        Case acMacro
'            DoCmd.SelectObject acMacro, objectName, True
'            DoCmd.RunCommand acCmdDesignView
'        Case Else
'            Err.Raise vbObjectError + 8462, "SharedAccess_ReopenPreparedObject", "The object type does not have a safe supported reopen path." & vbCrLf & "Object type: " & CStr(objectType) & vbCrLf & "Object: " & objectName
'    End Select
'End Sub
'
'Private Sub SharedAccess_SaveDirtyFormRecord(ByVal formName As String)
'    'Saves a pending record edit before an affected open form is saved and closed for a design change.
'    On Error GoTo ErrHandler
'    If CurrentProject.AllForms(formName).IsLoaded Then
'        If forms(formName).Dirty Then forms(formName).Dirty = False
'    End If
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedAccess_SaveDirtyFormRecord", "Could not save the pending form record before design preparation." & vbCrLf & "Form: " & formName & vbCrLf & "Description: " & savedErrDescription
'End Sub
'
'Public Function SharedAccess_PrepareOpenObjectForDesignChange(ByVal objectType As AcObjectType, ByVal objectName As String, ByVal actionText As String, Optional ByVal confirmationAlreadyProvided As Boolean = False) As Boolean
'    'Convenience wrapper for preparing one affected open Access object before a design change.
'    Dim names As New Collection
'    names.Add objectName
'    SharedAccess_PrepareOpenObjectForDesignChange = SharedAccess_PrepareOpenObjectsForDesignChange(objectType, names, actionText, confirmationAlreadyProvided)
'End Function
'
'Public Function SharedFormDesign_PrepareGeneratedForm(ByVal formName As String, ByVal replaceExisting As Boolean, Optional ByVal requiredPrefix As String = vbNullString) As Boolean
'    'Prepares a generated form replacement by durably recording operation ownership before any generated target can be created.
'    On Error GoTo ErrHandler
'    Dim replacementKey As String
'    Dim recoveryName As String
'    Dim rollbackComplete As Boolean
'    Dim rollbackText As String
'    Dim replacementState As Object
'    Dim hadOriginal As Boolean
'    Dim markerName As String
'    If Len(formName) = 0 Then Err.Raise vbObjectError + 8463, "SharedFormDesign_PrepareGeneratedForm", "The generated form name is blank."
'    If Len(requiredPrefix) > 0 Then
'        If Left$(formName, Len(requiredPrefix)) <> requiredPrefix Then Err.Raise vbObjectError + 7601, "SharedFormDesign_PrepareGeneratedForm", "Refusing to create or replace form because the name does not start with the required generated-form prefix." & vbCrLf & "Form: " & formName & vbCrLf & "Required prefix: " & requiredPrefix
'    End If
'    SharedFormDesign_EnsureReplacementStore
'    replacementKey = UCase$(formName)
'    recoveryName = SharedFormDesign_RecoveryFormName(formName)
'    markerName = SharedFormDesign_RecoveryMarkerName(formName)
'    If mFormDesignReplacementBackups.Exists(replacementKey) Or SharedFormDesign_FormExistsStrict(recoveryName) Or SharedForm_TempVarExistsStrict(markerName) Then
'        rollbackText = SharedFormDesign_RollbackPendingReplacement(formName, rollbackComplete)
'        If rollbackComplete = False Then Err.Raise vbObjectError + 8464, "SharedFormDesign_PrepareGeneratedForm", "A prior generated-form replacement is still incomplete and could not be restored safely." & vbCrLf & "Form: " & formName & IIf(Len(rollbackText) > 0, vbCrLf & rollbackText, vbNullString)
'    End If
'    hadOriginal = SharedFormDesign_FormExistsStrict(formName)
'    If hadOriginal And replaceExisting = False Then
'        SharedFormDesign_PrepareGeneratedForm = False
'        Exit Function
'    End If
'    If hadOriginal Then
'        If SharedAccess_ObjectIsOpenStrict(acForm, formName) Then Err.Raise vbObjectError + 7611, "SharedFormDesign_PrepareGeneratedForm", "The generated form is still open. The owning workflow must save, close, and verify affected open forms before replacement." & vbCrLf & "Form: " & formName
'    End If
'    Set replacementState = CreateObject("Scripting.Dictionary")
'    replacementState.CompareMode = vbTextCompare
'    replacementState.Add "RecoveryName", recoveryName
'    replacementState.Add "HadOriginal", hadOriginal
'    replacementState.Add "OriginalCopied", False
'    replacementState.Add "OriginalDeleted", False
'    replacementState.Add "ReplacementSavedAndClosed", False
'    mFormDesignReplacementBackups.Add replacementKey, replacementState
'    If hadOriginal Then
'        SharedForm_SetTempVarTextStrict markerName, SharedFormDesign_RecoveryMarkerValue(formName, hadOriginal, "S")
'    Else
'        SharedForm_SetTempVarTextStrict markerName, SharedFormDesign_RecoveryMarkerValue(formName, hadOriginal, "P")
'    End If
'    If hadOriginal Then
'        DoCmd.CopyObject , recoveryName, acForm, formName
'        replacementState("OriginalCopied") = True
'        If SharedFormDesign_FormExistsStrict(formName) = False Then Err.Raise vbObjectError + 8465, "SharedFormDesign_PrepareGeneratedForm", "The authoritative form disappeared while its recovery copy was being created." & vbCrLf & "Form: " & formName & vbCrLf & "Recovery form: " & recoveryName
'        If SharedFormDesign_FormExistsStrict(recoveryName) = False Then Err.Raise vbObjectError + 8466, "SharedFormDesign_PrepareGeneratedForm", "The recovery copy is missing after Access reported that the copy completed." & vbCrLf & "Form: " & formName & vbCrLf & "Recovery form: " & recoveryName
'        DoCmd.DeleteObject acForm, formName
'        replacementState("OriginalDeleted") = True
'        If SharedFormDesign_FormExistsStrict(formName) Then Err.Raise vbObjectError + 8474, "SharedFormDesign_PrepareGeneratedForm", "The authoritative form remained after replacement preparation attempted to remove it." & vbCrLf & "Form: " & formName & vbCrLf & "Recovery form: " & recoveryName
'        If SharedFormDesign_FormExistsStrict(recoveryName) = False Then Err.Raise vbObjectError + 8475, "SharedFormDesign_PrepareGeneratedForm", "The recovery copy disappeared during replacement preparation." & vbCrLf & "Form: " & formName & vbCrLf & "Recovery form: " & recoveryName
'        SharedForm_SetTempVarTextStrict markerName, SharedFormDesign_RecoveryMarkerValue(formName, hadOriginal, "P")
'    End If
'    SharedFormDesign_PrepareGeneratedForm = True
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    If Len(formName) > 0 Then rollbackText = SharedFormDesign_RollbackPendingReplacement(formName, rollbackComplete)
'    Err.Raise savedErrNumber, "SharedFormDesign_PrepareGeneratedForm", "Could not prepare the generated form replacement safely." & vbCrLf & "Form: " & formName & vbCrLf & "Recovery form: " & recoveryName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & IIf(Len(rollbackText) > 0, vbCrLf & "Rollback details:" & vbCrLf & rollbackText, vbNullString)
'End Function
'Private Sub SharedFormDesign_EnsureReplacementStore()
'    'Creates the private generated-form replacement map used for same-execution cleanup while recovery identity remains reconstructable from the target name.
'    If mFormDesignReplacementBackups Is Nothing Then
'        Set mFormDesignReplacementBackups = CreateObject("Scripting.Dictionary")
'        mFormDesignReplacementBackups.CompareMode = vbTextCompare
'    End If
'End Sub
'
'Private Function SharedFormDesign_StableRecoveryToken(ByVal valueText As String) As String
'    'Builds a deterministic eight-character token so a recovery form can be rediscovered after a VBA project reset.
'    Dim i As Long
'    Dim charCode As Long
'    Dim hashValue As Double
'    Dim nextValue As Double
'    hashValue = 5381#
'    For i = 1 To Len(valueText)
'        charCode = AscW(Mid$(valueText, i, 1))
'        If charCode < 0 Then charCode = charCode + 65536
'        nextValue = (hashValue * 33#) + CDbl(charCode)
'        hashValue = nextValue - (Int(nextValue / 2147483647#) * 2147483647#)
'    Next i
'    SharedFormDesign_StableRecoveryToken = Right$("00000000" & Hex$(CLng(hashValue)), 8)
'End Function
'
'Private Function SharedFormDesign_RecoveryFormName(ByVal formName As String) As String
'    'Returns the deterministic reserved recovery-form name for one authoritative generated form.
'    Dim prefixText As String
'    Dim tokenText As String
'    Dim safePart As String
'    Dim maxPartLength As Long
'    prefixText = "SharedFormRecovery_"
'    tokenText = SharedFormDesign_StableRecoveryToken(UCase$(formName))
'    safePart = SharedName_SafeAccessObjectNamePart(formName, "Form", "F_")
'    maxPartLength = 64 - Len(prefixText) - Len(tokenText) - 1
'    If Len(safePart) > maxPartLength Then safePart = Left$(safePart, maxPartLength)
'    SharedFormDesign_RecoveryFormName = prefixText & safePart & "_" & tokenText
'End Function
'Private Function SharedFormDesign_RecoveryMarkerName(ByVal formName As String) As String
'    'Returns the deterministic TempVar name that persists ownership of one generated-form operation across a VBA reset.
'    SharedFormDesign_RecoveryMarkerName = "SharedFormRecoveryMarker_" & SharedFormDesign_StableRecoveryToken(UCase$(formName))
'End Function
'Private Function SharedFormDesign_RecoveryMarkerValue(ByVal formName As String, ByVal hadOriginal As Boolean, ByVal phaseText As String) As String
'    'Encodes marker version, operation phase, prior-form state, and exact target identity without restricting valid form-name characters.
'    Dim originalText As String
'    If phaseText <> "S" And phaseText <> "P" And phaseText <> "C" Then Err.Raise vbObjectError + 8484, "SharedFormDesign_RecoveryMarkerValue", "The generated-form recovery phase is invalid." & vbCrLf & "Phase: " & phaseText & vbCrLf & "Form: " & formName
'    If hadOriginal Then originalText = "1" Else originalText = "0"
'    SharedFormDesign_RecoveryMarkerValue = "V2|" & phaseText & "|" & originalText & "|" & formName
'End Function
'
'Private Function SharedFormDesign_TryParseRecoveryMarker(ByVal markerValue As String, ByRef markerFormName As String, ByRef markerHadOriginal As Boolean, ByRef markerPhase As String) As Boolean
'    'Parses the durable marker; legacy plain-name markers are recognized as pending replacements that had an original form.
'    Dim phaseEnd As Long
'    Dim originalEnd As Long
'    Dim originalText As String
'    markerFormName = vbNullString
'    markerHadOriginal = False
'    markerPhase = vbNullString
'    If Left$(markerValue, 3) = "V2|" Then
'        phaseEnd = InStr(4, markerValue, "|", vbBinaryCompare)
'        If phaseEnd = 0 Then Exit Function
'        originalEnd = InStr(phaseEnd + 1, markerValue, "|", vbBinaryCompare)
'        If originalEnd = 0 Then Exit Function
'        markerPhase = Mid$(markerValue, 4, phaseEnd - 4)
'        originalText = Mid$(markerValue, phaseEnd + 1, originalEnd - phaseEnd - 1)
'        markerFormName = Mid$(markerValue, originalEnd + 1)
'        If markerPhase <> "S" And markerPhase <> "P" And markerPhase <> "C" Then Exit Function
'        If originalText <> "0" And originalText <> "1" Then Exit Function
'        If Len(markerFormName) = 0 Then Exit Function
'        markerHadOriginal = (originalText = "1")
'        SharedFormDesign_TryParseRecoveryMarker = True
'    ElseIf Len(markerValue) > 0 Then
'        markerFormName = markerValue
'        markerHadOriginal = True
'        markerPhase = "P"
'        SharedFormDesign_TryParseRecoveryMarker = True
'    End If
'End Function
'
'Private Function SharedFormDesign_RollbackPendingReplacement(ByVal formName As String, ByRef rollbackComplete As Boolean) As String
'    'Uses durable ownership and phase state to restore a prior form or remove an incomplete first-time generated form.
'    Dim replacementKey As String
'    Dim replacementState As Object
'    Dim recoveryName As String
'    Dim markerName As String
'    Dim markerValue As String
'    Dim markerFormName As String
'    Dim markerExists As Boolean
'    Dim markerInspectionKnown As Boolean
'    Dim markerParsed As Boolean
'    Dim markerMatches As Boolean
'    Dim markerHadOriginal As Boolean
'    Dim markerPhase As String
'    Dim hadOriginal As Boolean
'    Dim hasTrackedState As Boolean
'    Dim replacementSavedAndClosed As Boolean
'    Dim replacementCommitStateKnown As Boolean
'    Dim ownershipKnown As Boolean
'    Dim stateConflict As Boolean
'    Dim targetExists As Boolean
'    Dim recoveryExists As Boolean
'    Dim targetInspectionKnown As Boolean
'    Dim recoveryInspectionKnown As Boolean
'    Dim detailText As String
'    rollbackComplete = False
'    If Len(formName) = 0 Then
'        SharedFormDesign_RollbackPendingReplacement = "Rollback could not start because the target form name is blank."
'        Exit Function
'    End If
'    SharedFormDesign_EnsureReplacementStore
'    replacementKey = UCase$(formName)
'    recoveryName = SharedFormDesign_RecoveryFormName(formName)
'    markerName = SharedFormDesign_RecoveryMarkerName(formName)
'    hasTrackedState = mFormDesignReplacementBackups.Exists(replacementKey)
'    If hasTrackedState Then
'        Set replacementState = mFormDesignReplacementBackups(replacementKey)
'        hadOriginal = CBool(replacementState("HadOriginal"))
'        ownershipKnown = True
'        If replacementState.Exists("ReplacementSavedAndClosed") = False Then
'            stateConflict = True
'            detailText = SharedText_AppendNonBlankLine(detailText, "The active replacement state does not contain a committed-state value.")
'        ElseIf VarType(replacementState("ReplacementSavedAndClosed")) <> vbBoolean Then
'            stateConflict = True
'            detailText = SharedText_AppendNonBlankLine(detailText, "The active replacement committed-state value is not Boolean.")
'        Else
'            replacementSavedAndClosed = CBool(replacementState("ReplacementSavedAndClosed"))
'            replacementCommitStateKnown = True
'        End If
'    End If
'    On Error Resume Next
'    Err.Clear
'    targetExists = SharedFormDesign_FormExistsStrict(formName)
'    targetInspectionKnown = (Err.Number = 0)
'    If targetInspectionKnown = False Then detailText = SharedText_AppendNonBlankLine(detailText, "Target inspection error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'    Err.Clear
'    recoveryExists = SharedFormDesign_FormExistsStrict(recoveryName)
'    recoveryInspectionKnown = (Err.Number = 0)
'    If recoveryInspectionKnown = False Then detailText = SharedText_AppendNonBlankLine(detailText, "Recovery inspection error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'    Err.Clear
'    markerExists = SharedForm_TempVarExistsStrict(markerName)
'    markerInspectionKnown = (Err.Number = 0)
'    If markerInspectionKnown = False Then
'        detailText = SharedText_AppendNonBlankLine(detailText, "Recovery marker inspection error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'    ElseIf markerExists Then
'        Err.Clear
'        markerValue = SharedForm_TempVarTextStrict(markerName)
'        markerInspectionKnown = (Err.Number = 0)
'        If markerInspectionKnown = False Then
'            detailText = SharedText_AppendNonBlankLine(detailText, "Recovery marker read error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'        Else
'            markerParsed = SharedFormDesign_TryParseRecoveryMarker(markerValue, markerFormName, markerHadOriginal, markerPhase)
'            markerMatches = markerParsed And (StrComp(markerFormName, formName, vbTextCompare) = 0)
'            If markerParsed = False Then
'                detailText = SharedText_AppendNonBlankLine(detailText, "The persisted recovery marker has an unrecognized format.")
'            ElseIf markerMatches = False Then
'                detailText = SharedText_AppendNonBlankLine(detailText, "The persisted recovery marker identifies a different target.")
'            ElseIf ownershipKnown And hadOriginal <> markerHadOriginal Then
'                stateConflict = True
'                detailText = SharedText_AppendNonBlankLine(detailText, "The persisted prior-form state conflicts with the active replacement state.")
'            Else
'                hadOriginal = markerHadOriginal
'                ownershipKnown = True
'            End If
'        End If
'    End If
'    If stateConflict = False And hasTrackedState And replacementCommitStateKnown And markerInspectionKnown And markerExists And markerMatches Then
'        If replacementSavedAndClosed Then
'            If markerPhase <> "C" Then
'                stateConflict = True
'                detailText = SharedText_AppendNonBlankLine(detailText, "The active replacement state records a committed form, but the durable marker records a non-committed phase.")
'            End If
'        ElseIf markerPhase = "C" Then
'            stateConflict = True
'            detailText = SharedText_AppendNonBlankLine(detailText, "The durable marker records a committed form, but the active replacement state records an uncommitted form.")
'        End If
'    End If
'    If targetInspectionKnown And recoveryInspectionKnown And markerInspectionKnown Then
'        If stateConflict Or (markerExists And markerMatches = False) Then
'            detailText = SharedText_AppendNonBlankLine(detailText, "Recovery was not attempted because durable ownership could not be verified consistently.")
'        ElseIf (markerExists And markerPhase = "C") Or (hasTrackedState And replacementCommitStateKnown And replacementSavedAndClosed) Then
'            If hadOriginal = False And recoveryExists Then
'                detailText = SharedText_AppendNonBlankLine(detailText, "Committed cleanup was not attempted because a recovery copy exists for an operation recorded as having no prior form.")
'            ElseIf targetExists = False Then
'                detailText = SharedText_AppendNonBlankLine(detailText, "Committed generated-form cleanup was not attempted because the verified replacement target is now missing.")
'            Else
'                If recoveryExists Then
'                    Err.Clear
'                    DoCmd.DeleteObject acForm, recoveryName
'                    If Err.Number <> 0 Then detailText = SharedText_AppendNonBlankLine(detailText, "Committed recovery-copy cleanup error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'                End If
'                Err.Clear
'                recoveryExists = SharedFormDesign_FormExistsStrict(recoveryName)
'                recoveryInspectionKnown = (Err.Number = 0)
'                If recoveryInspectionKnown = False Then detailText = SharedText_AppendNonBlankLine(detailText, "Committed recovery-copy verification error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'                If recoveryInspectionKnown And recoveryExists = False Then
'                    Err.Clear
'                    SharedForm_RemoveTempVarIfPresentStrict markerName
'                    markerExists = SharedForm_TempVarExistsStrict(markerName)
'                    markerInspectionKnown = (Err.Number = 0)
'                    If markerInspectionKnown = False Then
'                        detailText = SharedText_AppendNonBlankLine(detailText, "Committed marker cleanup verification error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'                    ElseIf markerExists Then
'                        detailText = SharedText_AppendNonBlankLine(detailText, "The committed recovery marker remained after cleanup.")
'                    Else
'                        rollbackComplete = True
'                        If mFormDesignReplacementBackups.Exists(replacementKey) Then mFormDesignReplacementBackups.Remove replacementKey
'                        detailText = SharedText_AppendNonBlankLine(detailText, "Committed generated-form cleanup was completed without rolling back the verified replacement.")
'                    End If
'                ElseIf recoveryInspectionKnown Then
'                    detailText = SharedText_AppendNonBlankLine(detailText, "The retained recovery copy remained after committed cleanup.")
'                End If
'            End If
'        ElseIf markerExists And markerPhase = "S" And ownershipKnown And hadOriginal And recoveryExists = False Then
'            If targetExists Then
'                Err.Clear
'                SharedForm_RemoveTempVarIfPresentStrict markerName
'                markerExists = SharedForm_TempVarExistsStrict(markerName)
'                markerInspectionKnown = (Err.Number = 0)
'                If markerInspectionKnown = False Then
'                    detailText = SharedText_AppendNonBlankLine(detailText, "Starting-state marker cleanup verification error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'                ElseIf markerExists Then
'                    detailText = SharedText_AppendNonBlankLine(detailText, "The starting-state recovery marker remained after cleanup.")
'                Else
'                    rollbackComplete = True
'                    If mFormDesignReplacementBackups.Exists(replacementKey) Then mFormDesignReplacementBackups.Remove replacementKey
'                    detailText = SharedText_AppendNonBlankLine(detailText, "The original form remained authoritative and the interrupted pre-copy state was cleared.")
'                End If
'            Else
'                detailText = SharedText_AppendNonBlankLine(detailText, "The original form is missing before a recovery copy was established; automatic cleanup is unsafe.")
'            End If
'        ElseIf recoveryExists Then
'            If ownershipKnown = False Then
'                detailText = SharedText_AppendNonBlankLine(detailText, "Recovery was not attempted because the deterministic recovery object has no verified ownership state.")
'            ElseIf hadOriginal = False Then
'                detailText = SharedText_AppendNonBlankLine(detailText, "Recovery was not attempted because a recovery copy exists for an operation recorded as having no prior form.")
'            Else
'                Err.Clear
'                SharedFormDesign_CloseFormIfOpen formName, acSaveNo
'                If Err.Number <> 0 Then detailText = SharedText_AppendNonBlankLine(detailText, "Incomplete target close error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'                Err.Clear
'                targetExists = SharedFormDesign_FormExistsStrict(formName)
'                targetInspectionKnown = (Err.Number = 0)
'                If targetInspectionKnown = False Then
'                    detailText = SharedText_AppendNonBlankLine(detailText, "Target pre-deletion inspection error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'                ElseIf targetExists Then
'                    Err.Clear
'                    DoCmd.DeleteObject acForm, formName
'                    If Err.Number <> 0 Then detailText = SharedText_AppendNonBlankLine(detailText, "Incomplete target deletion error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'                End If
'                Err.Clear
'                targetExists = SharedFormDesign_FormExistsStrict(formName)
'                targetInspectionKnown = (Err.Number = 0)
'                If targetInspectionKnown = False Then
'                    detailText = SharedText_AppendNonBlankLine(detailText, "Target deletion verification error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'                ElseIf targetExists Then
'                    detailText = SharedText_AppendNonBlankLine(detailText, "The incomplete target remained after rollback cleanup: " & formName)
'                Else
'                    Err.Clear
'                    DoCmd.CopyObject , formName, acForm, recoveryName
'                    If Err.Number <> 0 Then detailText = SharedText_AppendNonBlankLine(detailText, "Original-form restore-copy error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'                End If
'                Err.Clear
'                targetExists = SharedFormDesign_FormExistsStrict(formName)
'                targetInspectionKnown = (Err.Number = 0)
'                If targetInspectionKnown = False Then detailText = SharedText_AppendNonBlankLine(detailText, "Restored target verification error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'                Err.Clear
'                recoveryExists = SharedFormDesign_FormExistsStrict(recoveryName)
'                recoveryInspectionKnown = (Err.Number = 0)
'                If recoveryInspectionKnown = False Then detailText = SharedText_AppendNonBlankLine(detailText, "Recovery pre-cleanup verification error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'                If targetInspectionKnown And recoveryInspectionKnown And targetExists And recoveryExists Then
'                    Err.Clear
'                    DoCmd.DeleteObject acForm, recoveryName
'                    If Err.Number <> 0 Then detailText = SharedText_AppendNonBlankLine(detailText, "Recovery cleanup error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'                End If
'                Err.Clear
'                targetExists = SharedFormDesign_FormExistsStrict(formName)
'                targetInspectionKnown = (Err.Number = 0)
'                If targetInspectionKnown = False Then detailText = SharedText_AppendNonBlankLine(detailText, "Final target verification error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'                Err.Clear
'                recoveryExists = SharedFormDesign_FormExistsStrict(recoveryName)
'                recoveryInspectionKnown = (Err.Number = 0)
'                If recoveryInspectionKnown = False Then detailText = SharedText_AppendNonBlankLine(detailText, "Final recovery verification error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'                If targetInspectionKnown And recoveryInspectionKnown And targetExists And recoveryExists = False Then
'                    Err.Clear
'                    SharedForm_RemoveTempVarIfPresentStrict markerName
'                    markerExists = SharedForm_TempVarExistsStrict(markerName)
'                    markerInspectionKnown = (Err.Number = 0)
'                    If markerInspectionKnown = False Then
'                        detailText = SharedText_AppendNonBlankLine(detailText, "Final recovery marker verification error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'                    ElseIf markerExists Then
'                        detailText = SharedText_AppendNonBlankLine(detailText, "The recovery marker remained after the prior form was restored.")
'                    Else
'                        rollbackComplete = True
'                        If mFormDesignReplacementBackups.Exists(replacementKey) Then mFormDesignReplacementBackups.Remove replacementKey
'                        detailText = SharedText_AppendNonBlankLine(detailText, "The prior form was restored from its recovery copy and verified under its authoritative name.")
'                    End If
'                Else
'                    detailText = SharedText_AppendNonBlankLine(detailText, "Rollback remains incomplete. Target exists: " & CStr(targetExists) & "; recovery exists: " & CStr(recoveryExists) & ".")
'                End If
'            End If
'        ElseIf ownershipKnown And hadOriginal = False Then
'            Err.Clear
'            SharedFormDesign_CloseFormIfOpen formName, acSaveNo
'            If Err.Number <> 0 Then detailText = SharedText_AppendNonBlankLine(detailText, "New target close error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'            Err.Clear
'            targetExists = SharedFormDesign_FormExistsStrict(formName)
'            targetInspectionKnown = (Err.Number = 0)
'            If targetInspectionKnown And targetExists Then
'                Err.Clear
'                DoCmd.DeleteObject acForm, formName
'                If Err.Number <> 0 Then detailText = SharedText_AppendNonBlankLine(detailText, "New target deletion error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'            End If
'            Err.Clear
'            targetExists = SharedFormDesign_FormExistsStrict(formName)
'            targetInspectionKnown = (Err.Number = 0)
'            If targetInspectionKnown And targetExists = False Then
'                Err.Clear
'                SharedForm_RemoveTempVarIfPresentStrict markerName
'                markerExists = SharedForm_TempVarExistsStrict(markerName)
'                markerInspectionKnown = (Err.Number = 0)
'                If markerInspectionKnown = False Then
'                    detailText = SharedText_AppendNonBlankLine(detailText, "New-target marker cleanup verification error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'                ElseIf markerExists Then
'                    detailText = SharedText_AppendNonBlankLine(detailText, "The new-target ownership marker remained after rollback cleanup.")
'                Else
'                    rollbackComplete = True
'                    If mFormDesignReplacementBackups.Exists(replacementKey) Then mFormDesignReplacementBackups.Remove replacementKey
'                    detailText = SharedText_AppendNonBlankLine(detailText, "The incomplete newly generated form was removed and verified absent.")
'                End If
'            ElseIf targetInspectionKnown Then
'                detailText = SharedText_AppendNonBlankLine(detailText, "The incomplete newly generated form remained after rollback cleanup.")
'            Else
'                detailText = SharedText_AppendNonBlankLine(detailText, "Rollback remains incomplete because the new target state could not be verified.")
'            End If
'        ElseIf ownershipKnown And hadOriginal Then
'            If markerExists And markerPhase = "P" Then
'                detailText = SharedText_AppendNonBlankLine(detailText, "The pending replacement records a prior form, but its recovery copy is missing; automatic cleanup is unsafe.")
'            ElseIf targetExists Then
'                Err.Clear
'                SharedForm_RemoveTempVarIfPresentStrict markerName
'                markerExists = SharedForm_TempVarExistsStrict(markerName)
'                markerInspectionKnown = (Err.Number = 0)
'                If markerInspectionKnown = False Then
'                    detailText = SharedText_AppendNonBlankLine(detailText, "Pre-copy marker cleanup verification error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'                ElseIf markerExists Then
'                    detailText = SharedText_AppendNonBlankLine(detailText, "The pre-copy recovery marker remained after cleanup.")
'                Else
'                    rollbackComplete = True
'                    If mFormDesignReplacementBackups.Exists(replacementKey) Then mFormDesignReplacementBackups.Remove replacementKey
'                    detailText = SharedText_AppendNonBlankLine(detailText, "The original form remained under its authoritative name and the pending pre-copy state was cleared.")
'                End If
'            Else
'                detailText = SharedText_AppendNonBlankLine(detailText, "The recorded original form and its recovery copy are both missing; automatic restoration is unsafe.")
'            End If
'        ElseIf recoveryExists Then
'            detailText = SharedText_AppendNonBlankLine(detailText, "Recovery was not attempted because the deterministic recovery object has no persisted ownership marker.")
'        Else
'            rollbackComplete = True
'        End If
'    Else
'        detailText = SharedText_AppendNonBlankLine(detailText, "Rollback remains incomplete because form-state inspection was not conclusive.")
'    End If
'    On Error GoTo 0
'    SharedFormDesign_RollbackPendingReplacement = detailText
'End Function
'Public Sub SharedFormDesign_AbortGeneratedFormReplacement(ByVal formName As String)
'    'Intended use: Call from every generated-form builder error handler to remove an incomplete target and restore any retained prior form.
'    Dim rollbackComplete As Boolean
'    Dim rollbackText As String
'    If Len(formName) = 0 Then Err.Raise vbObjectError + 8476, "SharedFormDesign_AbortGeneratedFormReplacement", "The generated form name is blank."
'    rollbackText = SharedFormDesign_RollbackPendingReplacement(formName, rollbackComplete)
'    If rollbackComplete = False Then Err.Raise vbObjectError + 8477, "SharedFormDesign_AbortGeneratedFormReplacement", "The generated-form replacement could not be aborted safely." & vbCrLf & "Form: " & formName & IIf(Len(rollbackText) > 0, vbCrLf & rollbackText, vbNullString)
'End Sub
'Private Function SharedFormDesign_AbortPendingReplacementForError(ByVal formName As String) As String
'    'Best-effort error-handler adapter that preserves the triggering error while reporting generated-form rollback results.
'    On Error GoTo ErrHandler
'    Dim rollbackComplete As Boolean
'    Dim rollbackText As String
'    If Len(formName) = 0 Then Exit Function
'    rollbackText = SharedFormDesign_RollbackPendingReplacement(formName, rollbackComplete)
'    If rollbackComplete = False Then
'        SharedFormDesign_AbortPendingReplacementForError = "Automatic generated-form abort did not complete." & IIf(Len(rollbackText) > 0, vbCrLf & rollbackText, vbNullString)
'    ElseIf Len(rollbackText) > 0 Then
'        SharedFormDesign_AbortPendingReplacementForError = rollbackText
'    End If
'    Exit Function
'ErrHandler:
'    SharedFormDesign_AbortPendingReplacementForError = "Automatic generated-form abort error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
'End Function
'Private Function SharedFormDesign_FormNameFromControl(ByVal ctl As Access.control) As String
'    'Returns the containing form name for a generated control or page when Access exposes a traversable parent chain.
'    On Error Resume Next
'    Dim currentObject As Object
'    Dim parentObject As Object
'    Dim depth As Long
'    Set currentObject = ctl
'    For depth = 1 To 12
'        If currentObject Is Nothing Then Exit For
'        If TypeOf currentObject Is Access.Form Then
'            SharedFormDesign_FormNameFromControl = CStr(currentObject.Name)
'            Exit Function
'        End If
'        Set parentObject = Nothing
'        Err.Clear
'        Set parentObject = currentObject.Parent
'        If Err.Number <> 0 Then Exit For
'        Set currentObject = parentObject
'    Next depth
'End Function
'
'Private Sub SharedFormDesign_CommitPendingReplacement(ByVal formName As String, ByRef recoveryName As String)
'    'Durably marks the verified replacement committed before deleting retained recovery state, preventing cleanup failures from causing a later false rollback.
'    Dim replacementKey As String
'    Dim replacementState As Object
'    Dim recoveryExists As Boolean
'    Dim markerName As String
'    Dim markerExists As Boolean
'    Dim markerValue As String
'    Dim markerFormName As String
'    Dim markerHadOriginal As Boolean
'    Dim markerPhase As String
'    Dim markerParsed As Boolean
'    Dim hasTrackedState As Boolean
'    Dim hadOriginal As Boolean
'    Dim ownershipKnown As Boolean
'    SharedFormDesign_EnsureReplacementStore
'    replacementKey = UCase$(formName)
'    recoveryName = SharedFormDesign_RecoveryFormName(formName)
'    markerName = SharedFormDesign_RecoveryMarkerName(formName)
'    hasTrackedState = mFormDesignReplacementBackups.Exists(replacementKey)
'    If hasTrackedState Then
'        Set replacementState = mFormDesignReplacementBackups(replacementKey)
'        hadOriginal = CBool(replacementState("HadOriginal"))
'        ownershipKnown = True
'    End If
'    If SharedAccess_ObjectIsOpenStrict(acForm, formName) Then Err.Raise vbObjectError + 8468, "SharedFormDesign_CommitPendingReplacement", "The replacement form is still open at commit." & vbCrLf & "Form: " & formName
'    If SharedFormDesign_FormExistsStrict(formName) = False Then Err.Raise vbObjectError + 8469, "SharedFormDesign_CommitPendingReplacement", "The replacement form is missing at commit." & vbCrLf & "Form: " & formName
'    recoveryExists = SharedFormDesign_FormExistsStrict(recoveryName)
'    markerExists = SharedForm_TempVarExistsStrict(markerName)
'    If markerExists Then
'        markerValue = SharedForm_TempVarTextStrict(markerName)
'        markerParsed = SharedFormDesign_TryParseRecoveryMarker(markerValue, markerFormName, markerHadOriginal, markerPhase)
'        If markerParsed = False Then Err.Raise vbObjectError + 8480, "SharedFormDesign_CommitPendingReplacement", "The persisted recovery marker has an unrecognized format; commit was stopped without deleting recovery state." & vbCrLf & "Form: " & formName & vbCrLf & "Recovery form: " & recoveryName
'        If StrComp(markerFormName, formName, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 8478, "SharedFormDesign_CommitPendingReplacement", "The persisted recovery marker identifies a different target; commit was stopped without deleting the recovery form." & vbCrLf & "Form: " & formName & vbCrLf & "Recovery form: " & recoveryName
'        If ownershipKnown And hadOriginal <> markerHadOriginal Then Err.Raise vbObjectError + 8481, "SharedFormDesign_CommitPendingReplacement", "The persisted prior-form state conflicts with the active replacement state; commit was stopped without deleting recovery state." & vbCrLf & "Form: " & formName & vbCrLf & "Recovery form: " & recoveryName
'        hadOriginal = markerHadOriginal
'        ownershipKnown = True
'    ElseIf recoveryExists Or hasTrackedState Then
'        Err.Raise vbObjectError + 8479, "SharedFormDesign_CommitPendingReplacement", "The pending generated-form operation has no persisted ownership marker; commit was stopped without deleting recovery state." & vbCrLf & "Form: " & formName & vbCrLf & "Recovery form: " & recoveryName
'    End If
'    If ownershipKnown = False Then Exit Sub
'    If markerPhase = "S" Then Err.Raise vbObjectError + 8485, "SharedFormDesign_CommitPendingReplacement", "The replacement cannot be committed because durable recovery state still records the pre-copy preparation phase." & vbCrLf & "Form: " & formName & vbCrLf & "Recovery form: " & recoveryName
'    If hadOriginal And recoveryExists = False Then Err.Raise vbObjectError + 8482, "SharedFormDesign_CommitPendingReplacement", "The operation records a prior form, but its retained recovery copy is missing; commit was stopped." & vbCrLf & "Form: " & formName & vbCrLf & "Recovery form: " & recoveryName
'    If hadOriginal = False And recoveryExists Then Err.Raise vbObjectError + 8483, "SharedFormDesign_CommitPendingReplacement", "A recovery copy exists for an operation recorded as having no prior form; commit was stopped without deleting either object." & vbCrLf & "Form: " & formName & vbCrLf & "Recovery form: " & recoveryName
'    SharedForm_SetTempVarTextStrict markerName, SharedFormDesign_RecoveryMarkerValue(formName, hadOriginal, "C")
'    If hasTrackedState Then replacementState("ReplacementSavedAndClosed") = True
'    If recoveryExists Then
'        DoCmd.DeleteObject acForm, recoveryName
'        If SharedFormDesign_FormExistsStrict(recoveryName) Then Err.Raise vbObjectError + 8471, "SharedFormDesign_CommitPendingReplacement", "The retained prior-form copy remained after commit cleanup." & vbCrLf & "Form: " & formName & vbCrLf & "Recovery form: " & recoveryName
'        If SharedFormDesign_FormExistsStrict(formName) = False Then Err.Raise vbObjectError + 8472, "SharedFormDesign_CommitPendingReplacement", "The replacement form is missing after commit cleanup." & vbCrLf & "Form: " & formName
'    End If
'    SharedForm_RemoveTempVarIfPresentStrict markerName
'    If mFormDesignReplacementBackups.Exists(replacementKey) Then mFormDesignReplacementBackups.Remove replacementKey
'End Sub
'Public Function SharedFormDesign_CreateBlankFormForDesign(ByVal formName As String) As Access.Form
'    'Creates one new blank generated form, removes operation-created artifacts on failure, and restores any retained prior form.
'    On Error GoTo ErrHandler
'    Dim frm As Access.Form
'    Dim tempName As String
'    Dim tempCreated As Boolean
'    Dim targetCreated As Boolean
'    If Len(formName) = 0 Then Err.Raise vbObjectError + 7766, "SharedFormDesign_CreateBlankFormForDesign", "The target form name is blank."
'    If SharedFormDesign_FormExistsStrict(formName) Then Err.Raise vbObjectError + 7767, "SharedFormDesign_CreateBlankFormForDesign", "The target form already exists." & vbCrLf & "Form: " & formName
'    Set frm = CreateForm
'    tempName = frm.Name
'    tempCreated = True
'    DoCmd.Save acForm, tempName
'    Set frm = Nothing
'    SharedFormDesign_CloseFormIfOpen tempName, acSaveYes
'    DoCmd.Rename formName, acForm, tempName
'    tempCreated = False
'    targetCreated = True
'    DoCmd.OpenForm formName, acDesign
'    If SharedAccess_ObjectIsOpenStrict(acForm, formName) = False Then Err.Raise vbObjectError + 7768, "SharedFormDesign_CreateBlankFormForDesign", "The generated form did not open in Design view."
'    Set SharedFormDesign_CreateBlankFormForDesign = forms(formName)
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim cleanupErrorText As String
'    Dim rollbackText As String
'    Dim rollbackComplete As Boolean
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Set frm = Nothing
'    If targetCreated Then cleanupErrorText = SharedText_AppendNonBlankLine(cleanupErrorText, SharedFormDesign_CleanupGeneratedFormArtifact(formName, "Target form"))
'    If tempCreated And Len(tempName) > 0 Then cleanupErrorText = SharedText_AppendNonBlankLine(cleanupErrorText, SharedFormDesign_CleanupGeneratedFormArtifact(tempName, "Temporary form"))
'    rollbackText = SharedFormDesign_RollbackPendingReplacement(formName, rollbackComplete)
'    cleanupErrorText = SharedText_AppendNonBlankLine(cleanupErrorText, rollbackText)
'    Err.Raise savedErrNumber, "SharedFormDesign_CreateBlankFormForDesign", "Could not create blank generated form." & vbCrLf & "Target form: " & formName & vbCrLf & "Temporary form: " & tempName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & IIf(Len(cleanupErrorText) > 0, vbCrLf & "Cleanup and rollback details:" & vbCrLf & cleanupErrorText, vbNullString)
'End Function
'Private Function SharedFormDesign_CleanupGeneratedFormArtifact(ByVal formName As String, ByVal artifactText As String) As String
'    'Attempts close, deletion, and strict final-state verification for one operation-created generated form artifact.
'    On Error GoTo UnexpectedError
'    Dim detailText As String
'    Dim FormExists As Boolean
'    On Error Resume Next
'    SharedFormDesign_CloseFormIfOpen formName, acSaveNo
'    If Err.Number <> 0 Then detailText = SharedText_AppendNonBlankLine(detailText, artifactText & " close error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'    Err.Clear
'    FormExists = SharedFormDesign_FormExistsStrict(formName)
'    If Err.Number <> 0 Then
'        detailText = SharedText_AppendNonBlankLine(detailText, artifactText & " inspection error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'    ElseIf FormExists Then
'        Err.Clear
'        DoCmd.DeleteObject acForm, formName
'        If Err.Number <> 0 Then detailText = SharedText_AppendNonBlankLine(detailText, artifactText & " deletion error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'        Err.Clear
'        FormExists = SharedFormDesign_FormExistsStrict(formName)
'        If Err.Number <> 0 Then
'            detailText = SharedText_AppendNonBlankLine(detailText, artifactText & " verification error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
'        ElseIf FormExists Then
'            detailText = SharedText_AppendNonBlankLine(detailText, artifactText & " remained after cleanup: " & formName)
'        End If
'    End If
'    On Error GoTo 0
'    SharedFormDesign_CleanupGeneratedFormArtifact = detailText
'    Exit Function
'UnexpectedError:
'    SharedFormDesign_CleanupGeneratedFormArtifact = artifactText & " cleanup helper error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
'End Function
'
'Public Sub SharedFormDesign_SaveAndCloseDesignForm(ByVal formName As String)
'    'Intended use: Saves one generated design form, closes and verifies it, then commits any pending replacement by deleting the retained prior-form copy.
'    On Error GoTo ErrHandler
'    Dim targetSavedAndClosed As Boolean
'    Dim recoveryName As String
'    If Len(formName) = 0 Then Err.Raise vbObjectError + 7769, "SharedFormDesign_SaveAndCloseDesignForm", "The form name is blank."
'    DoCmd.Save acForm, formName
'    SharedFormDesign_CloseFormIfOpen formName, acSaveYes
'    If SharedAccess_ObjectIsOpenStrict(acForm, formName) Then Err.Raise vbObjectError + 7770, "SharedFormDesign_SaveAndCloseDesignForm", "The generated form remained open after save and close."
'    If SharedFormDesign_FormExistsStrict(formName) = False Then Err.Raise vbObjectError + 8473, "SharedFormDesign_SaveAndCloseDesignForm", "The generated form is missing after save and close." & vbCrLf & "Form: " & formName
'    targetSavedAndClosed = True
'    SharedFormDesign_CommitPendingReplacement formName, recoveryName
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim rollbackText As String
'    Dim rollbackComplete As Boolean
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    If targetSavedAndClosed = False Then rollbackText = SharedFormDesign_RollbackPendingReplacement(formName, rollbackComplete)
'    Err.Raise savedErrNumber, "SharedFormDesign_SaveAndCloseDesignForm", "Could not save, close, verify, and commit the generated form." & vbCrLf & "Form: " & formName & vbCrLf & "Recovery form: " & recoveryName & vbCrLf & "Replacement saved and closed: " & CStr(targetSavedAndClosed) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & IIf(Len(rollbackText) > 0, vbCrLf & "Rollback details:" & vbCrLf & rollbackText, vbNullString)
'End Sub
'Public Function SharedFormDesign_CreateControl(ByVal formName As String, ByVal ctlType As AcControlType, ByVal sectionConstant As AcSection, ByVal parentName As String, ByVal columnName As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal widthValue As Long, ByVal heightValue As Long, Optional ByVal actionText As String = vbNullString, Optional ByVal detailText As String = vbNullString) As Access.control
'    'Public shared API for generated-form builders that need a validated Access control-creation boundary.
'    On Error GoTo ErrHandler
'    SharedFormDesign_ValidateControlBounds leftPos, topPos, widthValue, heightValue
'    SharedFormDesign_ValidateParentNameForCreateControl formName, parentName
'    Set SharedFormDesign_CreateControl = CreateControl(formName, ctlType, sectionConstant, parentName, columnName, leftPos, topPos, widthValue, heightValue)
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim abortText As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    abortText = SharedFormDesign_AbortPendingReplacementForError(formName)
'    Err.Raise savedErrNumber, "SharedFormDesign_CreateControl", SharedFormDesign_ControlErrorMessage("Could not create form control.", formName, ctlType, sectionConstant, parentName, columnName, leftPos, topPos, widthValue, heightValue, actionText, detailText, savedErrNumber, savedErrSource, savedErrDescription) & IIf(Len(abortText) > 0, vbCrLf & "Automatic abort details:" & vbCrLf & abortText, vbNullString)
'End Function
'
'Public Function SharedFormDesign_AddBoundControl(ByVal formName As String, ByVal sourceName As String, ByVal fieldName As String, ByVal ctlType As AcControlType, ByVal sectionConstant As AcSection, ByVal parentName As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal widthValue As Long, ByVal heightValue As Long, ByVal controlPrefix As String) As Access.control
'    'Public shared API for generated-form builders that need one named bound control.
'    On Error GoTo ErrHandler
'    Dim ctl As Access.control
'    Dim ctlName As String
'    Dim detailText As String
'    ctlName = SharedFormDesign_SafeControlName(controlPrefix & "_" & fieldName)
'    detailText = "Field: " & fieldName & "; Source: " & sourceName & "; ControlName: " & ctlName
'    Set ctl = SharedFormDesign_CreateControl(formName, ctlType, sectionConstant, parentName, vbNullString, leftPos, topPos, widthValue, heightValue, "Create bound control", detailText)
'    ctl.Name = ctlName
'    ctl.controlSource = SharedDao_BracketName(fieldName)
'    ctl.statusBarText = fieldName
'    Set SharedFormDesign_AddBoundControl = ctl
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim abortText As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    abortText = SharedFormDesign_AbortPendingReplacementForError(formName)
'    Err.Raise savedErrNumber, "SharedFormDesign_AddBoundControl", "Could not add the bound form control." & vbCrLf & "Form: " & formName & vbCrLf & "Source: " & sourceName & vbCrLf & "Field: " & fieldName & vbCrLf & "Control prefix: " & controlPrefix & vbCrLf & "Control type: " & SharedFormDesign_ControlTypeDisplayName(ctlType) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & IIf(Len(abortText) > 0, vbCrLf & "Automatic abort details:" & vbCrLf & abortText, vbNullString)
'End Function
'
'Public Function SharedFormDesign_AddPlainLabel(ByVal formName As String, ByVal sectionConstant As AcSection, ByVal parentName As String, ByVal captionText As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal widthValue As Long, ByVal heightValue As Long, Optional ByVal fontSize As Long = 9, Optional ByVal fontBold As Boolean = False) As Access.control
'    'Public shared API for generated-form builders that need one plain label.
'    On Error GoTo ErrHandler
'    Dim ctl As Access.control
'    Dim detailText As String
'    detailText = "Caption: " & Left$(captionText, 120)
'    Set ctl = SharedFormDesign_CreateControl(formName, acLabel, sectionConstant, parentName, vbNullString, leftPos, topPos, widthValue, heightValue, "Create label", detailText)
'    ctl.Caption = captionText
'    ctl.fontSize = fontSize
'    ctl.fontBold = fontBold
'    Set SharedFormDesign_AddPlainLabel = ctl
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim abortText As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    abortText = SharedFormDesign_AbortPendingReplacementForError(formName)
'    Err.Raise savedErrNumber, "SharedFormDesign_AddPlainLabel", "Could not add the form label." & vbCrLf & "Form: " & formName & vbCrLf & "Caption: " & Left$(captionText, 120) & vbCrLf & "Section: " & SharedFormDesign_SectionDisplayName(sectionConstant) & vbCrLf & "Parent/Page: " & parentName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & IIf(Len(abortText) > 0, vbCrLf & "Automatic abort details:" & vbCrLf & abortText, vbNullString)
'End Function
'
'Public Function SharedFormDesign_AddTitleLabel(ByVal formName As String, ByVal sectionConstant As AcSection, ByVal parentName As String, ByVal captionText As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal widthValue As Long, ByVal heightValue As Long) As Access.control
'    'Intended use: Adds a consistently styled title label while generating a form in Design view.
'    Set SharedFormDesign_AddTitleLabel = SharedFormDesign_AddPlainLabel(formName, sectionConstant, parentName, captionText, leftPos, topPos, widthValue, heightValue, 14, True)
'End Function
'
'Public Function SharedFormDesign_AddSectionLabel(ByVal formName As String, ByVal sectionConstant As AcSection, ByVal parentName As String, ByVal captionText As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal widthValue As Long, ByVal heightValue As Long) As Access.control
'    'Intended use: Adds a consistently styled section label while generating a form in Design view.
'    Set SharedFormDesign_AddSectionLabel = SharedFormDesign_AddPlainLabel(formName, sectionConstant, parentName, captionText, leftPos, topPos, widthValue, heightValue, 11, True)
'End Function
'
'Public Function SharedFormDesign_AddButton(ByVal formName As String, ByVal sectionConstant As AcSection, ByVal parentName As String, ByVal captionText As String, ByVal onClickExpression As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal widthValue As Long, ByVal heightValue As Long, ByVal tooltipText As String) As Access.control
'    'Public shared API for generated-form builders that need one configured command button.
'    On Error GoTo ErrHandler
'    Dim ctl As Access.control
'    Dim detailText As String
'    detailText = "Caption: " & captionText & "; OnClick: " & onClickExpression
'    Set ctl = SharedFormDesign_CreateControl(formName, acCommandButton, sectionConstant, parentName, vbNullString, leftPos, topPos, widthValue, heightValue, "Create command button", detailText)
'    ctl.Caption = captionText
'    ctl.OnClick = onClickExpression
'    ctl.ControlTipText = tooltipText
'    Set SharedFormDesign_AddButton = ctl
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim abortText As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    abortText = SharedFormDesign_AbortPendingReplacementForError(formName)
'    Err.Raise savedErrNumber, "SharedFormDesign_AddButton", "Could not add the command button." & vbCrLf & "Form: " & formName & vbCrLf & "Caption: " & captionText & vbCrLf & "OnClick: " & onClickExpression & vbCrLf & "Parent/Page: " & parentName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & IIf(Len(abortText) > 0, vbCrLf & "Automatic abort details:" & vbCrLf & abortText, vbNullString)
'End Function
'
'Public Function SharedFormDesign_AddTabControl(ByVal formName As String, ByVal controlName As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal widthValue As Long, ByVal heightValue As Long) As Access.control
'    'Public shared API for generated-form builders that need one named tab control.
'    On Error GoTo ErrHandler
'    Dim ctl As Access.control
'    Dim safeName As String
'    Dim detailText As String
'    safeName = SharedFormDesign_SafeControlName(controlName)
'    detailText = "ControlName: " & safeName
'    Set ctl = SharedFormDesign_CreateControl(formName, acTabCtl, acDetail, vbNullString, vbNullString, leftPos, topPos, widthValue, heightValue, "Create tab control", detailText)
'    ctl.Name = safeName
'    Set SharedFormDesign_AddTabControl = ctl
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim abortText As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    abortText = SharedFormDesign_AbortPendingReplacementForError(formName)
'    Err.Raise savedErrNumber, "SharedFormDesign_AddTabControl", "Could not add the tab control." & vbCrLf & "Form: " & formName & vbCrLf & "Requested control name: " & controlName & vbCrLf & "Safe control name: " & safeName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & IIf(Len(abortText) > 0, vbCrLf & "Automatic abort details:" & vbCrLf & abortText, vbNullString)
'End Function
'
'Public Function SharedFormDesign_AddSubformControl(ByVal formName As String, ByVal sectionConstant As AcSection, ByVal parentName As String, ByVal controlName As String, ByVal sourceFormName As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal widthValue As Long, ByVal heightValue As Long) As Access.control
'    'Public shared API for generated-form builders that need one configured subform control.
'    On Error GoTo ErrHandler
'    Dim ctl As Access.control
'    Dim safeName As String
'    Dim detailText As String
'    safeName = SharedFormDesign_SafeControlName(controlName)
'    detailText = "ControlName: " & safeName & "; SourceObject: Form." & sourceFormName
'    Set ctl = SharedFormDesign_CreateControl(formName, acSubform, sectionConstant, parentName, vbNullString, leftPos, topPos, widthValue, heightValue, "Create subform control", detailText)
'    ctl.Name = safeName
'    ctl.sourceObject = "Form." & sourceFormName
'    ctl.LinkChildFields = vbNullString
'    ctl.LinkMasterFields = vbNullString
'    ctl.BorderStyle = 1
'    Set SharedFormDesign_AddSubformControl = ctl
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim abortText As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    abortText = SharedFormDesign_AbortPendingReplacementForError(formName)
'    Err.Raise savedErrNumber, "SharedFormDesign_AddSubformControl", "Could not add the subform control." & vbCrLf & "Form: " & formName & vbCrLf & "Control name: " & controlName & vbCrLf & "Source form: " & sourceFormName & vbCrLf & "Parent/Page: " & parentName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & IIf(Len(abortText) > 0, vbCrLf & "Automatic abort details:" & vbCrLf & abortText, vbNullString)
'End Function
'
'Public Function SharedFormDesign_ConfigureTabPage(ByVal tabCtl As Access.control, ByVal pageIndex As Long, ByVal pageName As String, ByVal captionText As String) As Access.Page
'    'Intended use: Names and captions one tab page and returns the verified page object to the form generator.
'    On Error GoTo ErrHandler
'    Dim pg As Access.Page
'    If pageIndex < 0 Then Err.Raise vbObjectError + 7602, "SharedFormDesign_ConfigureTabPage", "Invalid tab page index: " & CStr(pageIndex)
'    Do While tabCtl.pages.count <= pageIndex
'        Set pg = tabCtl.pages.Add
'    Loop
'    Set pg = tabCtl.pages(pageIndex)
'    pg.Name = SharedFormDesign_SafeControlName(pageName)
'    pg.Caption = captionText
'    Set SharedFormDesign_ConfigureTabPage = pg
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim abortFormName As String
'    Dim abortText As String
'    Dim tabControlName As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    abortFormName = SharedFormDesign_FormNameFromControl(tabCtl)
'    tabControlName = SharedFormDesign_SafeControlName(SharedFormDesign_ControlName(tabCtl))
'    abortText = SharedFormDesign_AbortPendingReplacementForError(abortFormName)
'    Err.Raise savedErrNumber, "SharedFormDesign_ConfigureTabPage", "Could not configure tab page." & vbCrLf & "TabControl: " & tabControlName & vbCrLf & "PageIndex: " & CStr(pageIndex) & vbCrLf & "PageName: " & pageName & vbCrLf & "Caption: " & captionText & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription & IIf(Len(abortText) > 0, vbCrLf & "Automatic abort details:" & vbCrLf & abortText, vbNullString)
'End Function
'
'Public Sub SharedFormDesign_ConfigureTwoColumnLookup(ByVal ctl As Access.control, ByVal sqlText As String)
'    'Intended use: Configures a generated lookup control with the supplied two-column RowSource SQL and aborts pending generation if configuration fails.
'    On Error GoTo ErrHandler
'    ctl.rowSourceType = "Table/Query"
'    ctl.rowSource = sqlText
'    ctl.LimitToList = True
'    ctl.columnCount = 2
'    ctl.BoundColumn = 1
'    ctl.columnWidths = "0;2.5in"
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim abortFormName As String
'    Dim abortText As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    abortFormName = SharedFormDesign_FormNameFromControl(ctl)
'    abortText = SharedFormDesign_AbortPendingReplacementForError(abortFormName)
'    Err.Raise savedErrNumber, "SharedFormDesign_ConfigureTwoColumnLookup", "Could not configure the generated two-column lookup." & vbCrLf & "Form: " & abortFormName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & IIf(Len(abortText) > 0, vbCrLf & "Automatic abort details:" & vbCrLf & abortText, vbNullString)
'End Sub
'
'Public Sub SharedFormDesign_ConfigureMultilineTextBox(ByVal ctl As Access.control)
'    'Intended use: Configures a generated text box for multiline entry and vertical scrolling, aborting pending generation on a material property failure.
'    On Error GoTo ErrHandler
'    ctl.EnterKeyBehavior = True
'    ctl.ScrollBars = 2
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim abortFormName As String
'    Dim abortText As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    abortFormName = SharedFormDesign_FormNameFromControl(ctl)
'    abortText = SharedFormDesign_AbortPendingReplacementForError(abortFormName)
'    Err.Raise savedErrNumber, "SharedFormDesign_ConfigureMultilineTextBox", "Could not configure the generated multiline text box." & vbCrLf & "Form: " & abortFormName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & IIf(Len(abortText) > 0, vbCrLf & "Automatic abort details:" & vbCrLf & abortText, vbNullString)
'End Sub
'
'Public Sub SharedFormDesign_ApplyLockedControlStyle(ByVal ctl As Access.control, Optional ByVal backColorValue As Long = SHARED_FORM_DESIGN_LOCKED_BACK)
'    'Intended use: Applies the shared enabled-but-locked style to a generated control, skipping unsupported background-color styling for selector controls, and aborts pending generation if a required property cannot
'    'be applied.
'    On Error GoTo ErrHandler
'    Dim styleStage As String
'    Dim controlName As String
'    Dim controlTypeValue As Long
'    Dim controlTypeText As String
'    Dim skipBackColor As Boolean
'    If ctl Is Nothing Then Err.Raise 91, "SharedFormDesign_ApplyLockedControlStyle", "The generated control is Nothing."
'    controlName = ctl.Name
'    controlTypeValue = CLng(ctl.controlType)
'    controlTypeText = SharedFormDesign_ControlTypeDisplayName(ctl.controlType)
'    skipBackColor = (controlTypeValue = acCheckBox Or controlTypeValue = acOptionButton)
'    styleStage = "Set Locked=True"
'    ctl.Locked = True
'    If ctl.Locked = False Then Err.Raise vbObjectError + 8600, "SharedFormDesign_ApplyLockedControlStyle", "The generated control did not retain Locked=True."
'    styleStage = "Set Enabled=True"
'    ctl.enabled = True
'    If ctl.enabled = False Then Err.Raise vbObjectError + 8601, "SharedFormDesign_ApplyLockedControlStyle", "The generated control did not retain Enabled=True."
'    If skipBackColor = False Then
'        styleStage = "Set BackColor=" & CStr(backColorValue)
'        ctl.BackColor = backColorValue
'        If ctl.BackColor <> backColorValue Then Err.Raise vbObjectError + 8602, "SharedFormDesign_ApplyLockedControlStyle", "The generated control did not retain the requested BackColor."
'    End If
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim abortFormName As String
'    Dim abortText As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    If Len(controlName) = 0 Then controlName = "(unavailable)"
'    If Len(controlTypeText) = 0 Then controlTypeText = "(unavailable)"
'    If Len(styleStage) = 0 Then styleStage = "Initialize locked-control style"
'    abortFormName = SharedFormDesign_FormNameFromControl(ctl)
'    abortText = SharedFormDesign_AbortPendingReplacementForError(abortFormName)
'    Err.Raise savedErrNumber, "SharedFormDesign_ApplyLockedControlStyle", "Could not apply the required locked style to the generated control." & vbCrLf & "Form: " & abortFormName & vbCrLf & "Control: " & controlName & vbCrLf & "Control type: " & controlTypeText & " (" & CStr(controlTypeValue) & ")" & vbCrLf & "Style stage: " & styleStage & vbCrLf & "Requested Locked: True" & vbCrLf & "Requested Enabled: True" & vbCrLf & "Requested BackColor: " & CStr(backColorValue) & vbCrLf & "BackColor intentionally skipped for this control type: " & CStr(skipBackColor) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Original description: " & savedErrDescription & IIf(Len(abortText) > 0, vbCrLf & "Automatic abort details:" & vbCrLf & abortText, vbNullString)
'End Sub
'
'Public Sub SharedFormDesign_ValidateControlBounds(ByVal leftPos As Long, ByVal topPos As Long, ByVal widthValue As Long, ByVal heightValue As Long)
'    'Intended use: Validates generated-control coordinates and Access design-surface limits before control creation.
'    If leftPos < 0 Or topPos < 0 Or widthValue <= 0 Or heightValue <= 0 Then
'        Err.Raise vbObjectError + 7603, "SharedFormDesign_ValidateControlBounds", "Invalid form-control bounds." & vbCrLf & "Left: " & CStr(leftPos) & vbCrLf & "Top: " & CStr(topPos) & vbCrLf & "Width: " & CStr(widthValue) & vbCrLf & "Height: " & CStr(heightValue)
'    End If
'    If leftPos + widthValue > SHARED_FORM_DESIGN_RIGHT_LIMIT Then
'        Err.Raise vbObjectError + 7604, "SharedFormDesign_ValidateControlBounds", "Form-control right edge exceeds Access design limit." & vbCrLf & "Left: " & CStr(leftPos) & vbCrLf & "Width: " & CStr(widthValue) & vbCrLf & "RightEdge: " & CStr(leftPos + widthValue) & vbCrLf & "Limit: " & CStr(SHARED_FORM_DESIGN_RIGHT_LIMIT)
'    End If
'    If topPos + heightValue > SHARED_FORM_DESIGN_BOTTOM_LIMIT Then
'        Err.Raise vbObjectError + 7605, "SharedFormDesign_ValidateControlBounds", "Form-control bottom edge exceeds Access design limit." & vbCrLf & "Top: " & CStr(topPos) & vbCrLf & "Height: " & CStr(heightValue) & vbCrLf & "BottomEdge: " & CStr(topPos + heightValue) & vbCrLf & "Limit: " & CStr(SHARED_FORM_DESIGN_BOTTOM_LIMIT)
'    End If
'End Sub
'
'Public Sub SharedFormDesign_ValidateParentNameForCreateControl(ByVal formName As String, ByVal parentName As String)
'    'Intended use: Validates a generated control parent or tab page before mutation and raises with form-specific evidence when inspection fails.
'    If Len(parentName) = 0 Then Exit Sub
'    If SharedFormDesign_FormParentOrPageExistsStrict(formName, parentName) = False Then Err.Raise vbObjectError + 7606, "SharedFormDesign_ValidateParentNameForCreateControl", "Parent control/page was not found on form." & vbCrLf & "Form: " & formName & vbCrLf & "Parent/Page: " & parentName
'End Sub
'
'Public Function SharedFormDesign_FormParentOrPageExists(ByVal formName As String, ByVal parentName As String) As Boolean
'    'Intended use: Provides a nonraising existence check for optional UI logic; generation and mutation paths must use the strict validator.
'    On Error GoTo ErrHandler
'    SharedFormDesign_FormParentOrPageExists = SharedFormDesign_FormParentOrPageExistsStrict(formName, parentName)
'    Exit Function
'ErrHandler:
'    SharedFormDesign_FormParentOrPageExists = False
'End Function
'
'Private Function SharedFormDesign_FormParentOrPageExistsStrict(ByVal formName As String, ByVal parentName As String) As Boolean
'    'Inspects an open form for a control or tab page and preserves technical inspection failures.
'    On Error GoTo ErrHandler
'    Dim frm As Access.Form
'    Dim ctl As Access.control
'    Dim pg As Access.Page
'    If Len(parentName) = 0 Then
'        SharedFormDesign_FormParentOrPageExistsStrict = True
'        Exit Function
'    End If
'    Set frm = forms(formName)
'    For Each ctl In frm.Controls
'        If StrComp(ctl.Name, parentName, vbTextCompare) = 0 Then
'            SharedFormDesign_FormParentOrPageExistsStrict = True
'            Exit Function
'        End If
'        If ctl.controlType = acTabCtl Then
'            For Each pg In ctl.pages
'                If StrComp(pg.Name, parentName, vbTextCompare) = 0 Then
'                    SharedFormDesign_FormParentOrPageExistsStrict = True
'                    Exit Function
'                End If
'            Next pg
'        End If
'    Next ctl
'    SharedFormDesign_FormParentOrPageExistsStrict = False
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedFormDesign_FormParentOrPageExistsStrict", "Unable to inspect the generated-form parent or tab-page collection." & vbCrLf & "Form: " & formName & vbCrLf & "Parent/Page: " & parentName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'
'Public Function SharedFormDesign_ControlTypeDisplayName(ByVal ctlType As AcControlType) As String
'    'Intended use: Returns a readable Access control-type name for generation diagnostics.
'    Select Case ctlType
'        Case acLabel
'            SharedFormDesign_ControlTypeDisplayName = "Label"
'        Case acTextBox
'            SharedFormDesign_ControlTypeDisplayName = "TextBox"
'        Case acComboBox
'            SharedFormDesign_ControlTypeDisplayName = "ComboBox"
'        Case acCheckBox
'            SharedFormDesign_ControlTypeDisplayName = "CheckBox"
'        Case acCommandButton
'            SharedFormDesign_ControlTypeDisplayName = "CommandButton"
'        Case acTabCtl
'            SharedFormDesign_ControlTypeDisplayName = "TabControl"
'        Case acSubform
'            SharedFormDesign_ControlTypeDisplayName = "Subform"
'        Case Else
'            SharedFormDesign_ControlTypeDisplayName = "AcControlType " & CStr(ctlType)
'    End Select
'End Function
'
'Public Function SharedFormDesign_SectionDisplayName(ByVal sectionConstant As AcSection) As String
'    'Intended use: Returns a readable Access form-section name for generation diagnostics.
'    Select Case sectionConstant
'        Case acDetail
'            SharedFormDesign_SectionDisplayName = "Detail"
'        Case acHeader
'            SharedFormDesign_SectionDisplayName = "Header"
'        Case acFooter
'            SharedFormDesign_SectionDisplayName = "Footer"
'        Case Else
'            SharedFormDesign_SectionDisplayName = "AcSection " & CStr(sectionConstant)
'    End Select
'End Function
'
'Public Function SharedFormDesign_SafeControlName(ByVal value As String) As String
'    'Intended use: Returns a valid generated control name derived from caller-provided text.
'    Dim i As Long
'    Dim ch As String
'    Dim resultText As String
'    value = Trim$(value)
'    For i = 1 To Len(value)
'        ch = Mid$(value, i, 1)
'        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Or ch = "_" Then
'            resultText = resultText & ch
'        Else
'            resultText = resultText & "_"
'        End If
'    Next i
'    Do While InStr(1, resultText, "__", vbBinaryCompare) > 0
'        resultText = Replace(resultText, "__", "_")
'    Loop
'    Do While Len(resultText) > 0 And Left$(resultText, 1) = "_"
'        resultText = Mid$(resultText, 2)
'    Loop
'    If Len(resultText) = 0 Then resultText = "Ctl"
'    ch = Left$(resultText, 1)
'    If Not ((ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z")) Then resultText = "Ctl_" & resultText
'    SharedFormDesign_SafeControlName = Left$(resultText, 64)
'End Function
'
'Public Function SharedFormDesign_FormExists(ByVal formName As String) As Boolean
'    'Public best-effort existence probe for non-destructive display logic; mutation paths use the strict private probe.
'    On Error GoTo NotFound
'    SharedFormDesign_FormExists = SharedFormDesign_FormExistsStrict(formName)
'    Exit Function
'NotFound:
'    SharedFormDesign_FormExists = False
'End Function
'
'Private Function SharedFormDesign_FormExistsStrict(ByVal formName As String) As Boolean
'    'Returns form existence while raising any Access collection-inspection failure.
'    On Error GoTo ErrHandler
'    Dim formObject As Object
'    If Len(formName) = 0 Then Err.Raise vbObjectError + 7801, "SharedFormDesign_FormExistsStrict", "The form name is blank."
'    For Each formObject In CurrentProject.AllForms
'        If StrComp(CStr(formObject.Name), formName, vbTextCompare) = 0 Then
'            SharedFormDesign_FormExistsStrict = True
'            Exit Function
'        End If
'    Next formObject
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedFormDesign_FormExistsStrict", "Could not inspect whether the Access form exists." & vbCrLf & "Form: " & formName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'Public Sub SharedFormDesign_CloseFormIfOpen(ByVal formName As String, Optional ByVal saveMode As AcCloseSave = acSaveNo)
'    'Closes one loaded form and verifies closure; errors are raised instead of being silently discarded.
'    On Error GoTo ErrHandler
'    If SharedAccess_ObjectIsOpenStrict(acForm, formName) Then
'        DoCmd.Close acForm, formName, saveMode
'        If SharedAccess_ObjectIsOpenStrict(acForm, formName) Then Err.Raise vbObjectError + 7610, "SharedFormDesign_CloseFormIfOpen", "The form remained open after the close command."
'    End If
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedFormDesign_CloseFormIfOpen", "Could not close the form." & vbCrLf & "Form: " & formName & vbCrLf & "Save mode: " & CStr(saveMode) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Sub
'
'Private Function SharedFormDesign_ControlErrorMessage(ByVal leadText As String, ByVal formName As String, ByVal ctlType As AcControlType, ByVal sectionConstant As AcSection, ByVal parentName As String, ByVal columnName As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal widthValue As Long, ByVal heightValue As Long, ByVal actionText As String, ByVal detailText As String, ByVal savedErrNumber As Long, ByVal savedErrSource As String, ByVal savedErrDescription As String) As String
'    'Builds control-creation diagnostics exclusively from captured error evidence.
'    SharedFormDesign_ControlErrorMessage = leadText & vbCrLf
'    SharedFormDesign_ControlErrorMessage = SharedFormDesign_ControlErrorMessage & "Form: " & formName & vbCrLf
'    SharedFormDesign_ControlErrorMessage = SharedFormDesign_ControlErrorMessage & "Action: " & actionText & vbCrLf
'    SharedFormDesign_ControlErrorMessage = SharedFormDesign_ControlErrorMessage & "Control type: " & SharedFormDesign_ControlTypeDisplayName(ctlType) & vbCrLf
'    SharedFormDesign_ControlErrorMessage = SharedFormDesign_ControlErrorMessage & "Section: " & SharedFormDesign_SectionDisplayName(sectionConstant) & vbCrLf
'    SharedFormDesign_ControlErrorMessage = SharedFormDesign_ControlErrorMessage & "Parent/Page: " & parentName & vbCrLf
'    SharedFormDesign_ControlErrorMessage = SharedFormDesign_ControlErrorMessage & "ColumnName: " & columnName & vbCrLf
'    SharedFormDesign_ControlErrorMessage = SharedFormDesign_ControlErrorMessage & "Bounds: Left=" & CStr(leftPos) & "; Top=" & CStr(topPos) & "; Width=" & CStr(widthValue) & "; Height=" & CStr(heightValue) & vbCrLf
'    SharedFormDesign_ControlErrorMessage = SharedFormDesign_ControlErrorMessage & "Detail: " & detailText & vbCrLf
'    SharedFormDesign_ControlErrorMessage = SharedFormDesign_ControlErrorMessage & "Err.Number: " & CStr(savedErrNumber) & vbCrLf
'    SharedFormDesign_ControlErrorMessage = SharedFormDesign_ControlErrorMessage & "Err.Source: " & savedErrSource & vbCrLf
'    SharedFormDesign_ControlErrorMessage = SharedFormDesign_ControlErrorMessage & "Err.Description: " & savedErrDescription
'End Function
'
'Private Function SharedFormDesign_ControlName(ByVal ctl As Access.control) As String
'    On Error GoTo ErrHandler
'    SharedFormDesign_ControlName = ctl.Name
'    Exit Function
'ErrHandler:
'    SharedFormDesign_ControlName = "(control name unavailable)"
'End Function
'
'' ================================================================================================================================================================================
'' FORMS
'' ================================================================================================================================================================================
'
'Private Function SharedForm_DistinctBoundFieldNames(ByVal frm As Form, ByVal rs As DAO.Recordset, ByVal visibleOnly As Boolean, ByVal honorDatasheetHidden As Boolean) As Collection
'    'Collects distinct bound field names once for bulk form operations.
'    Dim fieldNames As Collection
'    Dim seenNames As Object
'    Dim ctl As control
'    Dim fieldName As String
'    Dim includeControl As Boolean
'    Dim datasheetView As Boolean
'    Set fieldNames = New Collection
'    Set seenNames = CreateObject("Scripting.Dictionary")
'    seenNames.CompareMode = vbTextCompare
'    If honorDatasheetHidden Then datasheetView = (frm.CurrentView = acCurViewDatasheet)
'    For Each ctl In frm.Controls
'        If ctl.controlType = acTextBox Or ctl.controlType = acComboBox Or ctl.controlType = acCheckBox Then
'            fieldName = CStr(Nz(ctl.controlSource, vbNullString))
'            includeControl = (Len(fieldName) > 0)
'            If includeControl Then includeControl = (Left$(fieldName, 1) <> "=")
'            If includeControl And visibleOnly Then includeControl = ctl.Visible
'            If includeControl And honorDatasheetHidden And datasheetView Then includeControl = (ctl.ColumnHidden = False)
'            If includeControl Then includeControl = SharedDao_RecordsetHasFieldSafe(rs, fieldName)
'            If includeControl Then
'                If seenNames.Exists(fieldName) = False Then
'                    seenNames.Add fieldName, fieldName
'                    fieldNames.Add fieldName
'                End If
'            End If
'        End If
'    Next ctl
'    Set SharedForm_DistinctBoundFieldNames = fieldNames
'End Function
'
'Private Function SharedForm_ValueHasData(ByVal value As Variant) As Boolean
'    'Applies the shared non-Null and nonblank-text definition used by column-data scans.
'    If IsNull(value) Then Exit Function
'    If VarType(value) = vbString Then
'        SharedForm_ValueHasData = (Len(Trim$(CStr(value))) > 0)
'    Else
'        SharedForm_ValueHasData = True
'    End If
'End Function
'
'Private Function SharedForm_RecordsetFieldDataMap(ByVal rs As DAO.Recordset, ByVal fieldNames As Collection) As Object
'    'Scans one recordset pass and records whether each requested field contains data.
'    Dim dataMap As Object
'    Dim fieldName As Variant
'    Dim unresolvedCount As Long
'    Set dataMap = CreateObject("Scripting.Dictionary")
'    dataMap.CompareMode = vbTextCompare
'    For Each fieldName In fieldNames
'        If dataMap.Exists(CStr(fieldName)) = False Then dataMap.Add CStr(fieldName), False
'    Next fieldName
'    unresolvedCount = dataMap.count
'    If unresolvedCount = 0 Then
'        Set SharedForm_RecordsetFieldDataMap = dataMap
'        Exit Function
'    End If
'    If rs.BOF And rs.EOF Then
'        Set SharedForm_RecordsetFieldDataMap = dataMap
'        Exit Function
'    End If
'    rs.MoveFirst
'    Do Until rs.EOF Or unresolvedCount = 0
'        For Each fieldName In fieldNames
'            If dataMap(CStr(fieldName)) = False Then
'                If SharedForm_ValueHasData(rs.fields(CStr(fieldName)).value) Then
'                    dataMap(CStr(fieldName)) = True
'                    unresolvedCount = unresolvedCount - 1
'                End If
'            End If
'        Next fieldName
'        If unresolvedCount > 0 Then rs.MoveNext
'    Loop
'    Set SharedForm_RecordsetFieldDataMap = dataMap
'End Function
'
'Public Sub SharedForm_HideNoDataColumns(ByVal frm As Form)
'    'Intended use: Hides bound datasheet columns with no data by scanning the current recordset only once.
'    On Error GoTo ErrHandler
'    Dim rs As DAO.Recordset
'    Dim fieldNames As Collection
'    Dim dataMap As Object
'    Dim ctl As control
'    Dim fieldName As String
'    Dim currentControlName As String
'    Dim formNameText As String
'    formNameText = "(Nothing)"
'    If Not frm Is Nothing Then formNameText = frm.Name
'    If frm Is Nothing Then Err.Raise 91, "SharedForm_HideNoDataColumns", "The Access.Form argument is Nothing."
'    Set rs = frm.RecordsetClone
'    Set fieldNames = SharedForm_DistinctBoundFieldNames(frm, rs, False, False)
'    Set dataMap = SharedForm_RecordsetFieldDataMap(rs, fieldNames)
'    For Each ctl In frm.Controls
'        currentControlName = ctl.Name
'        If ctl.controlType = acTextBox Or ctl.controlType = acComboBox Or ctl.controlType = acCheckBox Then
'            fieldName = CStr(Nz(ctl.controlSource, vbNullString))
'            If dataMap.Exists(fieldName) Then ctl.ColumnHidden = Not CBool(dataMap(fieldName))
'        End If
'    Next ctl
'    rs.Close
'    Set rs = Nothing
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim cleanupErrNumber As Long
'    Dim cleanupErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If Not rs Is Nothing Then rs.Close
'    cleanupErrNumber = Err.Number
'    cleanupErrDescription = Err.Description
'    Set rs = Nothing
'    Set fieldNames = Nothing
'    Set dataMap = Nothing
'    On Error GoTo 0
'    Err.Raise savedErrNumber, "SharedForm_HideNoDataColumns", "Could not evaluate and hide no-data columns." & vbCrLf & "Form: " & formNameText & vbCrLf & "Control: " & currentControlName & vbCrLf & "Field: " & fieldName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & IIf(cleanupErrNumber <> 0, vbCrLf & "Recordset cleanup error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription, vbNullString)
'End Sub
'
'Public Function SharedForm_ColumnHasData(ByVal rs As DAO.Recordset, ByVal fieldName As String) As Boolean
'    'Intended use: Determines whether a recordset field contains any non-Null value without changing the caller record position.
'    On Error GoTo ErrHandler
'    Dim scanRs As DAO.Recordset
'    If rs Is Nothing Then Err.Raise 91, "SharedForm_ColumnHasData", "The DAO.Recordset argument is Nothing."
'    Set scanRs = rs.Clone
'    If scanRs.BOF And scanRs.EOF Then GoTo ExitProcedure
'    scanRs.MoveFirst
'    Do Until scanRs.EOF
'        If SharedForm_ValueHasData(scanRs.fields(fieldName).value) Then
'            SharedForm_ColumnHasData = True
'            Exit Do
'        End If
'        scanRs.MoveNext
'    Loop
'ExitProcedure:
'    If Not scanRs Is Nothing Then scanRs.Close
'    Set scanRs = Nothing
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim cleanupErrNumber As Long
'    Dim cleanupErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If Not scanRs Is Nothing Then scanRs.Close
'    cleanupErrNumber = Err.Number
'    cleanupErrDescription = Err.Description
'    Set scanRs = Nothing
'    On Error GoTo 0
'    Err.Raise savedErrNumber, "SharedForm_ColumnHasData", "Could not inspect recordset field data." & vbCrLf & "Field: " & fieldName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & IIf(cleanupErrNumber <> 0, vbCrLf & "Clone cleanup error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription, vbNullString)
'End Function
'
'Public Sub SharedForm_ResetDatasheetColumns(ByVal frm As Form)
'    'Intended use: Restores datasheet columns to visible default ordering and sizing.
'    Dim ctl As control
'    For Each ctl In frm.Controls
'        If ctl.controlType = acTextBox Or ctl.controlType = acComboBox Or ctl.controlType = acCheckBox Then
'            ctl.ColumnHidden = False
'        End If
'    Next ctl
'End Sub
'
'Public Sub SharedForm_SearchVisibleFields(ByVal frm As Form, ByVal searchText As String)
'    'Intended use: Selects the first record whose currently visible bound fields contain caller-provided text.
'    On Error GoTo ErrHandler
'    Dim rs As DAO.Recordset
'    Dim fieldNames As Collection
'    Dim fieldName As Variant
'    Dim found As Boolean
'    Dim currentFieldName As String
'    Dim currentRecordNumber As Long
'    Dim formNameText As String
'    formNameText = "(Nothing)"
'    If Not frm Is Nothing Then formNameText = frm.Name
'    If frm Is Nothing Then Err.Raise vbObjectError + 8058, "SharedForm_SearchVisibleFields", "The Access.Form argument is Nothing."
'    Set fieldNames = SharedForm_VisibleBoundFieldNames(frm)
'    If fieldNames.count = 0 Then
'        MsgBox "No visible bound fields are available to search.", vbExclamation, "Search"
'        GoTo ExitProcedure
'    End If
'    Set rs = frm.RecordsetClone
'    If rs.BOF And rs.EOF Then GoTo NotFound
'    rs.MoveFirst
'    currentRecordNumber = 1
'    Do Until rs.EOF
'        For Each fieldName In fieldNames
'            currentFieldName = CStr(fieldName)
'            If InStr(1, Nz(rs.fields(currentFieldName).value, ""), searchText, vbTextCompare) > 0 Then
'                frm.Bookmark = rs.Bookmark
'                found = True
'                Exit For
'            End If
'        Next fieldName
'        If found Then Exit Do
'        rs.MoveNext
'        currentRecordNumber = currentRecordNumber + 1
'    Loop
'    If Not found Then GoTo NotFound
'ExitProcedure:
'    If Not rs Is Nothing Then
'        rs.Close
'        Set rs = Nothing
'    End If
'    Set fieldNames = Nothing
'    Exit Sub
'NotFound:
'    MsgBox "No visible field contains: " & searchText, vbInformation, "Search"
'    GoTo ExitProcedure
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim cleanupErrNumber As Long
'    Dim cleanupErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If Not rs Is Nothing Then rs.Close
'    cleanupErrNumber = Err.Number
'    cleanupErrDescription = Err.Description
'    Set rs = Nothing
'    Set fieldNames = Nothing
'    On Error GoTo 0
'    Err.Raise savedErrNumber, "SharedForm_SearchVisibleFields", "Could not search the visible bound fields." & vbCrLf & "Form: " & formNameText & vbCrLf & "Search text: " & searchText & vbCrLf & "Record number: " & CStr(currentRecordNumber) & vbCrLf & "Field: " & currentFieldName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & IIf(cleanupErrNumber <> 0, vbCrLf & "Recordset cleanup error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription, vbNullString)
'End Sub
'
'Public Function SharedForm_VisibleBoundFieldNames(ByVal frm As Form) As Collection
'    'Intended use: Returns distinct bound field names represented by visible form controls.
'    On Error GoTo ErrHandler
'    Dim fieldNames As Collection
'    Dim rs As DAO.Recordset
'    Dim formNameText As String
'    formNameText = "(Nothing)"
'    If Not frm Is Nothing Then formNameText = frm.Name
'    If frm Is Nothing Then Err.Raise 91, "SharedForm_VisibleBoundFieldNames", "The Access.Form argument is Nothing."
'    Set rs = frm.RecordsetClone
'    Set fieldNames = SharedForm_DistinctBoundFieldNames(frm, rs, True, True)
'    rs.Close
'    Set rs = Nothing
'    Set SharedForm_VisibleBoundFieldNames = fieldNames
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim cleanupErrNumber As Long
'    Dim cleanupErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If Not rs Is Nothing Then rs.Close
'    cleanupErrNumber = Err.Number
'    cleanupErrDescription = Err.Description
'    Set rs = Nothing
'    Set fieldNames = Nothing
'    On Error GoTo 0
'    Err.Raise savedErrNumber, "SharedForm_VisibleBoundFieldNames", "Could not collect distinct visible bound fields." & vbCrLf & "Form: " & formNameText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & IIf(cleanupErrNumber <> 0, vbCrLf & "Recordset cleanup error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription, vbNullString)
'End Function
'
'Public Function SharedDao_RecordsetHasFieldSafe(ByVal rs As DAO.Recordset, ByVal fieldName As String) As Boolean
'    'Intended use: Provides a nonraising recordset-field check for optional form and display logic.
'    Dim fld As DAO.Field
'    On Error Resume Next
'    Set fld = rs.fields(fieldName)
'    SharedDao_RecordsetHasFieldSafe = (Err.Number = 0 And Not fld Is Nothing)
'    Set fld = Nothing
'    Err.Clear
'    On Error GoTo 0
'End Function
'
'Private Function SharedForm_FilterSortExpressionErrorIsExpected(ByVal errorNumber As Long) As Boolean
'    'Classifies only known Access expression errors as invalid saved filter or OrderBy state.
'    Select Case errorNumber
'        Case 2101, 2465, 3061, 3075, 3122
'            SharedForm_FilterSortExpressionErrorIsExpected = True
'    End Select
'End Function
'
'Private Function SharedForm_CaptureTempVarTextState(ByVal varName As String) As Object
'    'Captures one optional TempVar name, existence state, exact Variant value, and Variant subtype for rollback.
'    On Error GoTo ErrHandler
'    Dim state As Object
'    Dim existed As Boolean
'    Dim originalValue As Variant
'    existed = SharedForm_TempVarExistsStrict(varName)
'    Set state = CreateObject("Scripting.Dictionary")
'    state.Add "Name", varName
'    state.Add "Existed", existed
'    If existed Then
'        originalValue = TempVars(varName).value
'        state.Add "Value", originalValue
'        state.Add "ValueType", VarType(originalValue)
'    Else
'        state.Add "Value", Empty
'        state.Add "ValueType", vbEmpty
'    End If
'    Set SharedForm_CaptureTempVarTextState = state
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedForm_CaptureTempVarTextState", "Could not capture the TempVar rollback state." & vbCrLf & "TempVar: " & varName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'
'Private Sub SharedForm_WriteTempVarTextRawStrict(ByVal varName As String, ByVal valueText As Variant)
'    'Writes one scalar Variant TempVar value without owning rollback; callers must capture prior state first.
'    On Error GoTo ErrHandler
'    Dim actualValue As Variant
'    Dim expectedType As Long
'    expectedType = VarType(valueText)
'    SharedForm_RemoveTempVarIfPresentStrict varName
'    TempVars.Add varName, valueText
'    If SharedForm_TempVarExistsStrict(varName) = False Then Err.Raise vbObjectError + 8038, "SharedForm_WriteTempVarTextRawStrict", "The TempVar is missing after it was added." & vbCrLf & "TempVar: " & varName
'    actualValue = TempVars(varName).value
'    If VarType(actualValue) <> expectedType Then Err.Raise vbObjectError + 8039, "SharedForm_WriteTempVarTextRawStrict", "The TempVar Variant subtype does not match after it was added." & vbCrLf & "TempVar: " & varName & vbCrLf & "Expected VarType: " & CStr(expectedType) & vbCrLf & "Actual VarType: " & CStr(VarType(actualValue))
'    If IsNull(valueText) Then
'        If IsNull(actualValue) = False Then Err.Raise vbObjectError + 8059, "SharedForm_WriteTempVarTextRawStrict", "The TempVar Null value does not match after it was added." & vbCrLf & "TempVar: " & varName
'    ElseIf IsEmpty(valueText) Then
'        If IsEmpty(actualValue) = False Then Err.Raise vbObjectError + 8060, "SharedForm_WriteTempVarTextRawStrict", "The TempVar Empty value does not match after it was added." & vbCrLf & "TempVar: " & varName
'    ElseIf StrComp(CStr(actualValue), CStr(valueText), vbBinaryCompare) <> 0 Then
'        Err.Raise vbObjectError + 8061, "SharedForm_WriteTempVarTextRawStrict", "The TempVar value does not match after it was added." & vbCrLf & "TempVar: " & varName
'    End If
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedForm_WriteTempVarTextRawStrict", "Could not write and verify the TempVar." & vbCrLf & "TempVar: " & varName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Sub
'
'Private Sub SharedForm_RestoreTempVarTextState(ByVal state As Object)
'    'Restores one captured TempVar state exactly and verifies existence, Variant subtype, and value.
'    On Error GoTo ErrHandler
'    Dim varName As String
'    Dim expectedValue As Variant
'    Dim actualValue As Variant
'    Dim expectedExists As Boolean
'    Dim expectedType As Long
'    If state Is Nothing Then Err.Raise vbObjectError + 8045, "SharedForm_RestoreTempVarTextState", "The captured TempVar state is Nothing."
'    varName = CStr(state("Name"))
'    expectedExists = CBool(state("Existed"))
'    expectedValue = state("Value")
'    expectedType = CLng(state("ValueType"))
'    If expectedExists Then
'        SharedForm_WriteTempVarTextRawStrict varName, expectedValue
'        If SharedForm_TempVarExistsStrict(varName) = False Then Err.Raise vbObjectError + 8046, "SharedForm_RestoreTempVarTextState", "The restored TempVar is missing." & vbCrLf & "TempVar: " & varName
'        actualValue = TempVars(varName).value
'        If VarType(actualValue) <> expectedType Then Err.Raise vbObjectError + 8047, "SharedForm_RestoreTempVarTextState", "The restored TempVar Variant subtype does not match the captured subtype." & vbCrLf & "TempVar: " & varName
'        If IsNull(expectedValue) Then
'            If IsNull(actualValue) = False Then Err.Raise vbObjectError + 8062, "SharedForm_RestoreTempVarTextState", "The restored TempVar Null value does not match the captured value." & vbCrLf & "TempVar: " & varName
'        ElseIf IsEmpty(expectedValue) Then
'            If IsEmpty(actualValue) = False Then Err.Raise vbObjectError + 8063, "SharedForm_RestoreTempVarTextState", "The restored TempVar Empty value does not match the captured value." & vbCrLf & "TempVar: " & varName
'        ElseIf StrComp(CStr(actualValue), CStr(expectedValue), vbBinaryCompare) <> 0 Then
'            Err.Raise vbObjectError + 8064, "SharedForm_RestoreTempVarTextState", "The restored TempVar value does not match the captured value." & vbCrLf & "TempVar: " & varName
'        End If
'    Else
'        SharedForm_RemoveTempVarIfPresentStrict varName
'        If SharedForm_TempVarExistsStrict(varName) Then Err.Raise vbObjectError + 8048, "SharedForm_RestoreTempVarTextState", "A TempVar that was originally absent still exists after rollback." & vbCrLf & "TempVar: " & varName
'    End If
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedForm_RestoreTempVarTextState", "Could not restore and verify the captured TempVar state." & vbCrLf & "TempVar: " & varName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Sub
'
'Private Function SharedForm_RestoreTempVarStatesBestEffort(ByVal states As Collection) As String
'    'Restores every captured TempVar state in reverse order and returns complete rollback evidence without raising.
'    On Error GoTo ReportingFailure
'    Dim i As Long
'    Dim state As Object
'    Dim restoreErrNumber As Long
'    Dim restoreErrDescription As String
'    Dim reportText As String
'    Dim varName As String
'    If states Is Nothing Then
'        SharedForm_RestoreTempVarStatesBestEffort = "Rollback state collection was unavailable."
'        Exit Function
'    End If
'    For i = states.count To 1 Step -1
'        Set state = Nothing
'        varName = "(unavailable)"
'        restoreErrNumber = 0
'        restoreErrDescription = vbNullString
'        On Error Resume Next
'        Err.Clear
'        Set state = states(i)
'        If Err.Number = 0 Then varName = CStr(state("Name"))
'        If Err.Number = 0 Then SharedForm_RestoreTempVarTextState state
'        restoreErrNumber = Err.Number
'        restoreErrDescription = Err.Description
'        On Error GoTo ReportingFailure
'        If Len(reportText) > 0 Then reportText = reportText & vbCrLf
'        If restoreErrNumber <> 0 Then
'            reportText = reportText & "TempVar rollback failed for " & varName & " with error " & CStr(restoreErrNumber) & ": " & restoreErrDescription
'        Else
'            reportText = reportText & "TempVar rollback restored and verified " & varName & "."
'        End If
'    Next i
'    SharedForm_RestoreTempVarStatesBestEffort = reportText
'    Exit Function
'ReportingFailure:
'    If Len(reportText) > 0 Then reportText = reportText & vbCrLf
'    reportText = reportText & "Rollback reporting failed with error " & CStr(Err.Number) & ": " & Err.Description
'    SharedForm_RestoreTempVarStatesBestEffort = reportText
'    Err.Clear
'End Function
'
'Private Sub SharedForm_RemoveTempVarPairStrict(ByVal firstVarName As String, ByVal secondVarName As String)
'    'Removes a related TempVar pair atomically and restores both original states if either removal fails.
'    On Error GoTo ErrHandler
'    Dim states As Collection
'    Dim mutationStarted As Boolean
'    Set states = New Collection
'    states.Add SharedForm_CaptureTempVarTextState(firstVarName)
'    states.Add SharedForm_CaptureTempVarTextState(secondVarName)
'    mutationStarted = True
'    SharedForm_RemoveTempVarIfPresentStrict firstVarName
'    SharedForm_RemoveTempVarIfPresentStrict secondVarName
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim rollbackText As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    If mutationStarted Then rollbackText = SharedForm_RestoreTempVarStatesBestEffort(states)
'    If Len(rollbackText) > 0 Then rollbackText = vbCrLf & "Rollback results:" & vbCrLf & rollbackText
'    Err.Raise savedErrNumber, "SharedForm_RemoveTempVarPairStrict", "Could not remove the related TempVar pair atomically." & vbCrLf & "First TempVar: " & firstVarName & vbCrLf & "Second TempVar: " & secondVarName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & rollbackText
'End Sub
'
'Private Sub SharedForm_RestoreOrderByStateStrict(ByVal frm As Access.Form, ByVal orderByText As String, ByVal orderByOnValue As Boolean)
'    'Restores and verifies one form's OrderBy expression and enabled state.
'    On Error GoTo ErrHandler
'    Dim formNameText As String
'    formNameText = "(Nothing)"
'    If Not frm Is Nothing Then formNameText = frm.Name
'    frm.OrderByOn = False
'    frm.OrderBy = orderByText
'    frm.OrderByOn = orderByOnValue
'    If StrComp(Nz(frm.OrderBy, vbNullString), orderByText, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 8049, "SharedForm_RestoreOrderByStateStrict", "The OrderBy expression does not match after restoration."
'    If CBool(frm.OrderByOn) <> orderByOnValue Then Err.Raise vbObjectError + 8050, "SharedForm_RestoreOrderByStateStrict", "The OrderByOn value does not match after restoration."
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedForm_RestoreOrderByStateStrict", "Could not restore and verify the live OrderBy state." & vbCrLf & "Form: " & formNameText & vbCrLf & "OrderBy: " & orderByText & vbCrLf & "OrderByOn: " & CStr(orderByOnValue) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Sub
'
'Private Sub SharedForm_RestoreLiveFilterSortState(ByVal frm As Access.Form, ByVal filterText As String, ByVal filterOnValue As Boolean, ByVal orderByText As String, ByVal orderByOnValue As Boolean)
'    'Restores and verifies the complete live filter and sort state after a failed apply operation.
'    On Error GoTo ErrHandler
'    Dim formNameText As String
'    formNameText = "(Nothing)"
'    If Not frm Is Nothing Then formNameText = frm.Name
'    frm.FilterOn = False
'    frm.Filter = filterText
'    frm.FilterOn = filterOnValue
'    frm.OrderByOn = False
'    frm.OrderBy = orderByText
'    frm.OrderByOn = orderByOnValue
'    If StrComp(Nz(frm.Filter, vbNullString), filterText, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 8051, "SharedForm_RestoreLiveFilterSortState", "The Filter expression does not match after restoration."
'    If CBool(frm.FilterOn) <> filterOnValue Then Err.Raise vbObjectError + 8052, "SharedForm_RestoreLiveFilterSortState", "The FilterOn value does not match after restoration."
'    If StrComp(Nz(frm.OrderBy, vbNullString), orderByText, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 8053, "SharedForm_RestoreLiveFilterSortState", "The OrderBy expression does not match after restoration."
'    If CBool(frm.OrderByOn) <> orderByOnValue Then Err.Raise vbObjectError + 8054, "SharedForm_RestoreLiveFilterSortState", "The OrderByOn value does not match after restoration."
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedForm_RestoreLiveFilterSortState", "Could not restore and verify the complete live filter and sort state." & vbCrLf & "Form: " & formNameText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Sub
'
'Private Function SharedForm_OrderByIsValid(ByVal frm As Access.Form, ByVal orderByText As String) As Boolean
'    'Tests one saved OrderBy expression, restores the original live state and record position, returns False only for known expression errors, and raises technical failures.
'    On Error GoTo ValidationFailure
'    Dim positionRs As DAO.Recordset
'    Dim originalOrderBy As String
'    Dim originalOrderByOn As Boolean
'    Dim originalBookmark As Variant
'    Dim stageText As String
'    Dim stateCaptured As Boolean
'    Dim bookmarkCaptured As Boolean
'    Dim validationSucceeded As Boolean
'    Dim validationErrNumber As Long
'    Dim validationErrSource As String
'    Dim validationErrDescription As String
'    Dim validationErrorExpected As Boolean
'    Dim restoreErrNumber As Long
'    Dim restoreErrDescription As String
'    Dim formNameText As String
'    formNameText = "(Nothing)"
'    If Not frm Is Nothing Then formNameText = frm.Name
'    If frm Is Nothing Then Err.Raise vbObjectError + 8055, "SharedForm_OrderByIsValid", "The Access.Form argument is Nothing."
'    If Len(Trim$(orderByText)) = 0 Then Exit Function
'    stageText = "capturing the original OrderBy state"
'    originalOrderBy = Nz(frm.OrderBy, vbNullString)
'    originalOrderByOn = CBool(frm.OrderByOn)
'    stateCaptured = True
'    stageText = "inspecting the current record position"
'    Set positionRs = frm.RecordsetClone
'    If Not (positionRs.BOF And positionRs.EOF) And frm.newRecord = False Then
'        originalBookmark = frm.Bookmark
'        bookmarkCaptured = True
'    End If
'    positionRs.Close
'    Set positionRs = Nothing
'    stageText = "assigning the candidate OrderBy expression"
'    frm.OrderBy = orderByText
'    stageText = "enabling the candidate OrderBy expression"
'    frm.OrderByOn = True
'    validationSucceeded = True
'    GoTo RestoreState
'ValidationFailure:
'    validationErrNumber = Err.Number
'    validationErrSource = Err.Source
'    validationErrDescription = Err.Description
'    validationErrorExpected = stateCaptured And (stageText = "assigning the candidate OrderBy expression" Or stageText = "enabling the candidate OrderBy expression") And SharedForm_FilterSortExpressionErrorIsExpected(validationErrNumber)
'    On Error Resume Next
'    If Not positionRs Is Nothing Then positionRs.Close
'    Set positionRs = Nothing
'    On Error GoTo 0
'RestoreState:
'    If stateCaptured Then
'        On Error Resume Next
'        Err.Clear
'        SharedForm_RestoreOrderByStateStrict frm, originalOrderBy, originalOrderByOn
'        If Err.Number = 0 And bookmarkCaptured Then frm.Bookmark = originalBookmark
'        restoreErrNumber = Err.Number
'        restoreErrDescription = Err.Description
'        On Error GoTo 0
'    End If
'    If restoreErrNumber <> 0 Then Err.Raise restoreErrNumber, "SharedForm_OrderByIsValid", "OrderBy validation could not restore the original live state and record position." & vbCrLf & "Form: " & formNameText & vbCrLf & "Candidate OrderBy: " & orderByText & vbCrLf & "Validation error number: " & CStr(validationErrNumber) & vbCrLf & "Validation error source: " & validationErrSource & vbCrLf & "Validation error description: " & validationErrDescription & vbCrLf & "Restoration error: " & restoreErrDescription
'    If validationErrNumber <> 0 Then
'        If validationErrorExpected Then
'            SharedForm_OrderByIsValid = False
'            Exit Function
'        End If
'        Err.Raise validationErrNumber, "SharedForm_OrderByIsValid", "A technical failure occurred while validating the saved OrderBy expression." & vbCrLf & "Form: " & formNameText & vbCrLf & "Stage: " & stageText & vbCrLf & "Candidate OrderBy: " & orderByText & vbCrLf & "Original source: " & validationErrSource & vbCrLf & "Description: " & validationErrDescription
'    End If
'    SharedForm_OrderByIsValid = validationSucceeded
'End Function
'
'Private Function SharedForm_TempVarExistsStrict(ByVal varName As String) As Boolean
'    'Returns whether one TempVar exists and raises if the collection cannot be inspected.
'    On Error GoTo ErrHandler
'    Dim tempItem As Object
'    For Each tempItem In TempVars
'        If StrComp(CStr(tempItem.Name), varName, vbBinaryCompare) = 0 Then
'            SharedForm_TempVarExistsStrict = True
'            Exit Function
'        End If
'    Next tempItem
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedForm_TempVarExistsStrict", "Could not inspect the TempVars collection." & vbCrLf & "TempVar: " & varName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'
'Private Function SharedForm_TempVarTextStrict(ByVal varName As String) As String
'    'Returns one optional TempVar as text while preserving technical collection failures.
'    On Error GoTo ErrHandler
'    If SharedForm_TempVarExistsStrict(varName) Then SharedForm_TempVarTextStrict = CStr(TempVars(varName))
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedForm_TempVarTextStrict", "Could not read the optional TempVar." & vbCrLf & "TempVar: " & varName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'
'Private Sub SharedForm_RemoveTempVarIfPresentStrict(ByVal varName As String)
'    'Removes one TempVar when present and verifies its absence.
'    On Error GoTo ErrHandler
'    If SharedForm_TempVarExistsStrict(varName) Then TempVars.Remove varName
'    If SharedForm_TempVarExistsStrict(varName) Then Err.Raise vbObjectError + 8037, "SharedForm_RemoveTempVarIfPresentStrict", "The TempVar still exists after removal." & vbCrLf & "TempVar: " & varName
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedForm_RemoveTempVarIfPresentStrict", "Could not remove and verify the TempVar." & vbCrLf & "TempVar: " & varName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Sub
'
'Private Sub SharedForm_SetTempVarTextStrict(ByVal varName As String, ByVal valueText As String)
'    'Replaces one TempVar with text, verifies it, and restores the original state if replacement fails.
'    On Error GoTo ErrHandler
'    Dim states As Collection
'    Dim mutationStarted As Boolean
'    Set states = New Collection
'    states.Add SharedForm_CaptureTempVarTextState(varName)
'    mutationStarted = True
'    SharedForm_WriteTempVarTextRawStrict varName, valueText
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim rollbackText As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    If mutationStarted Then rollbackText = SharedForm_RestoreTempVarStatesBestEffort(states)
'    If Len(rollbackText) > 0 Then rollbackText = vbCrLf & "Rollback results:" & vbCrLf & rollbackText
'    Err.Raise savedErrNumber, "SharedForm_SetTempVarTextStrict", "Could not replace and verify the TempVar." & vbCrLf & "TempVar: " & varName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & rollbackText
'End Sub
'
'Private Sub SharedForm_ClearFilterSortStateAtomic(ByVal frm As Access.Form, ByVal clearFilter As Boolean, ByVal clearOrderBy As Boolean)
'    'Clears requested live and saved form state atomically and restores the complete snapshot if any mutation or verification fails.
'    On Error GoTo ErrHandler
'    Dim originalFilter As String
'    Dim originalFilterOn As Boolean
'    Dim originalOrderBy As String
'    Dim originalOrderByOn As Boolean
'    Dim originalStates As Collection
'    Dim filterKey As String
'    Dim filterOnKey As String
'    Dim orderByKey As String
'    Dim orderByOnKey As String
'    Dim stageText As String
'    Dim formNameText As String
'    Dim mutationStarted As Boolean
'    formNameText = "(Nothing)"
'    If Not frm Is Nothing Then formNameText = frm.Name
'    If frm Is Nothing Then Err.Raise vbObjectError + 8065, "SharedForm_ClearFilterSortStateAtomic", "The Access.Form argument is Nothing."
'    filterKey = SharedForm_FilterKey(frm)
'    filterOnKey = SharedForm_FilterOnKey(frm)
'    orderByKey = SharedForm_OrderByKey(frm)
'    orderByOnKey = SharedForm_OrderByOnKey(frm)
'    stageText = "capturing the complete live state"
'    originalFilter = Nz(frm.Filter, vbNullString)
'    originalFilterOn = CBool(frm.FilterOn)
'    originalOrderBy = Nz(frm.OrderBy, vbNullString)
'    originalOrderByOn = CBool(frm.OrderByOn)
'    stageText = "capturing the complete saved state"
'    Set originalStates = New Collection
'    originalStates.Add SharedForm_CaptureTempVarTextState(filterKey)
'    originalStates.Add SharedForm_CaptureTempVarTextState(filterOnKey)
'    originalStates.Add SharedForm_CaptureTempVarTextState(orderByKey)
'    originalStates.Add SharedForm_CaptureTempVarTextState(orderByOnKey)
'    mutationStarted = True
'    If clearFilter Then
'        stageText = "clearing the live filter"
'        frm.FilterOn = False
'        frm.Filter = vbNullString
'        stageText = "removing the saved filter pair"
'        SharedForm_RemoveTempVarIfPresentStrict filterKey
'        SharedForm_RemoveTempVarIfPresentStrict filterOnKey
'    End If
'    If clearOrderBy Then
'        stageText = "clearing the live OrderBy"
'        frm.OrderByOn = False
'        frm.OrderBy = vbNullString
'        stageText = "removing the saved OrderBy pair"
'        SharedForm_RemoveTempVarIfPresentStrict orderByKey
'        SharedForm_RemoveTempVarIfPresentStrict orderByOnKey
'    End If
'    stageText = "verifying the cleared state"
'    If clearFilter Then
'        If CBool(frm.FilterOn) Then Err.Raise vbObjectError + 8066, "SharedForm_ClearFilterSortStateAtomic", "FilterOn remained enabled after clearing."
'        If Len(Nz(frm.Filter, vbNullString)) > 0 Then Err.Raise vbObjectError + 8067, "SharedForm_ClearFilterSortStateAtomic", "The Filter expression remained populated after clearing."
'        If SharedForm_TempVarExistsStrict(filterKey) Or SharedForm_TempVarExistsStrict(filterOnKey) Then Err.Raise vbObjectError + 8068, "SharedForm_ClearFilterSortStateAtomic", "Saved Filter TempVars remained after clearing."
'    End If
'    If clearOrderBy Then
'        If CBool(frm.OrderByOn) Then Err.Raise vbObjectError + 8069, "SharedForm_ClearFilterSortStateAtomic", "OrderByOn remained enabled after clearing."
'        If Len(Nz(frm.OrderBy, vbNullString)) > 0 Then Err.Raise vbObjectError + 8070, "SharedForm_ClearFilterSortStateAtomic", "The OrderBy expression remained populated after clearing."
'        If SharedForm_TempVarExistsStrict(orderByKey) Or SharedForm_TempVarExistsStrict(orderByOnKey) Then Err.Raise vbObjectError + 8071, "SharedForm_ClearFilterSortStateAtomic", "Saved OrderBy TempVars remained after clearing."
'    End If
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim liveRollbackErrNumber As Long
'    Dim liveRollbackErrDescription As String
'    Dim tempVarRollbackText As String
'    Dim rollbackText As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    If mutationStarted Then
'        On Error Resume Next
'        Err.Clear
'        SharedForm_RestoreLiveFilterSortState frm, originalFilter, originalFilterOn, originalOrderBy, originalOrderByOn
'        liveRollbackErrNumber = Err.Number
'        liveRollbackErrDescription = Err.Description
'        On Error GoTo 0
'        tempVarRollbackText = SharedForm_RestoreTempVarStatesBestEffort(originalStates)
'        If liveRollbackErrNumber <> 0 Then
'            rollbackText = vbCrLf & "Live-state rollback error " & CStr(liveRollbackErrNumber) & ": " & liveRollbackErrDescription
'        Else
'            rollbackText = vbCrLf & "The complete live state was restored and verified."
'        End If
'        If Len(tempVarRollbackText) > 0 Then rollbackText = rollbackText & vbCrLf & "Saved-state rollback results:" & vbCrLf & tempVarRollbackText
'    Else
'        rollbackText = vbCrLf & "No rollback was required because no mutation began."
'    End If
'    Err.Raise savedErrNumber, "SharedForm_ClearFilterSortStateAtomic", "Could not clear the requested form filter and sort state atomically." & vbCrLf & "Form: " & formNameText & vbCrLf & "Clear filter: " & CStr(clearFilter) & vbCrLf & "Clear OrderBy: " & CStr(clearOrderBy) & vbCrLf & "Stage: " & stageText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & rollbackText
'End Sub
'
'Private Sub SharedForm_ResetOrderBy(ByVal frm As Access.Form)
'    'Strictly clears active and saved OrderBy state for one form through the atomic state helper.
'    SharedForm_ClearFilterSortStateAtomic frm, False, True
'End Sub
'
'Public Sub SharedForm_SaveFilterAndSort(ByVal frm As Access.Form)
'    'Intended use: Persists the current form filter and sort state atomically and raises on technical persistence failures.
'    On Error GoTo ErrHandler
'    Dim currentFilter As String
'    Dim currentFilterOn As String
'    Dim currentOrderBy As String
'    Dim currentOrderByOn As String
'    Dim savedOrderBy As String
'    Dim savedOrderByOn As String
'    Dim stageText As String
'    Dim formNameText As String
'    Dim originalStates As Collection
'    Dim writesStarted As Boolean
'    Dim filterKey As String
'    Dim filterOnKey As String
'    Dim orderByKey As String
'    Dim orderByOnKey As String
'    formNameText = "(Nothing)"
'    If Not frm Is Nothing Then formNameText = frm.Name
'    If frm Is Nothing Then Err.Raise vbObjectError + 8041, "SharedForm_SaveFilterAndSort", "The Access.Form argument is Nothing."
'    filterKey = SharedForm_FilterKey(frm)
'    filterOnKey = SharedForm_FilterOnKey(frm)
'    orderByKey = SharedForm_OrderByKey(frm)
'    orderByOnKey = SharedForm_OrderByOnKey(frm)
'    stageText = "reading the current filter"
'    currentFilter = Nz(frm.Filter, vbNullString)
'    currentFilterOn = CStr(frm.FilterOn)
'    stageText = "reading the current order"
'    currentOrderBy = Nz(frm.OrderBy, vbNullString)
'    currentOrderByOn = CStr(frm.OrderByOn)
'    stageText = "reading previously saved order state"
'    savedOrderBy = SharedForm_TempVarTextStrict(orderByKey)
'    savedOrderByOn = SharedForm_TempVarTextStrict(orderByOnKey)
'    If Len(currentOrderBy) = 0 And currentOrderByOn = "False" And Len(savedOrderBy) > 0 And savedOrderByOn = "True" Then
'        stageText = "validating the previously saved order"
'        If SharedForm_OrderByIsValid(frm, savedOrderBy) Then
'            currentOrderBy = savedOrderBy
'            currentOrderByOn = savedOrderByOn
'        Else
'            currentOrderBy = vbNullString
'            currentOrderByOn = "False"
'        End If
'    End If
'    stageText = "capturing the original saved filter and sort state"
'    Set originalStates = New Collection
'    originalStates.Add SharedForm_CaptureTempVarTextState(filterKey)
'    originalStates.Add SharedForm_CaptureTempVarTextState(filterOnKey)
'    originalStates.Add SharedForm_CaptureTempVarTextState(orderByKey)
'    originalStates.Add SharedForm_CaptureTempVarTextState(orderByOnKey)
'    writesStarted = True
'    stageText = "saving the filter expression"
'    SharedForm_SetTempVarTextStrict filterKey, currentFilter
'    stageText = "saving the filter enabled flag"
'    SharedForm_SetTempVarTextStrict filterOnKey, currentFilterOn
'    stageText = "saving the order expression"
'    SharedForm_SetTempVarTextStrict orderByKey, currentOrderBy
'    stageText = "saving the order enabled flag"
'    SharedForm_SetTempVarTextStrict orderByOnKey, currentOrderByOn
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim rollbackText As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    If writesStarted Then rollbackText = SharedForm_RestoreTempVarStatesBestEffort(originalStates)
'    If Len(rollbackText) > 0 Then rollbackText = vbCrLf & "Atomic rollback results:" & vbCrLf & rollbackText
'    Err.Raise savedErrNumber, "SharedForm_SaveFilterAndSort", "Could not persist the form filter and sort state atomically." & vbCrLf & "Form: " & formNameText & vbCrLf & "Stage: " & stageText & vbCrLf & "Filter: " & currentFilter & vbCrLf & "FilterOn: " & currentFilterOn & vbCrLf & "OrderBy: " & currentOrderBy & vbCrLf & "OrderByOn: " & currentOrderByOn & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & rollbackText
'End Sub
'
'Public Sub SharedForm_RestoreFilterAndSort(ByVal frm As Access.Form)
'    'Intended use: Restores each complete saved filter or sort pair exactly, including disabled and blank states, while preserving live and saved state on technical failures.
'    On Error GoTo ErrHandler
'    Dim filterState As Object
'    Dim filterOnState As Object
'    Dim orderByState As Object
'    Dim orderByOnState As Object
'    Dim savedFilter As String
'    Dim savedFilterOn As Boolean
'    Dim savedOrderBy As String
'    Dim savedOrderByOn As Boolean
'    Dim originalFilter As String
'    Dim originalFilterOn As Boolean
'    Dim originalOrderBy As String
'    Dim originalOrderByOn As Boolean
'    Dim stageText As String
'    Dim formNameText As String
'    Dim filterKey As String
'    Dim filterOnKey As String
'    Dim orderByKey As String
'    Dim orderByOnKey As String
'    Dim filterPairExists As Boolean
'    Dim orderByPairExists As Boolean
'    Dim liveStateCaptured As Boolean
'    Dim mutationStarted As Boolean
'    Dim failureArea As Long
'    formNameText = "(Nothing)"
'    If Not frm Is Nothing Then formNameText = frm.Name
'    If frm Is Nothing Then Err.Raise vbObjectError + 8042, "SharedForm_RestoreFilterAndSort", "The Access.Form argument is Nothing."
'    filterKey = SharedForm_FilterKey(frm)
'    filterOnKey = SharedForm_FilterOnKey(frm)
'    orderByKey = SharedForm_OrderByKey(frm)
'    orderByOnKey = SharedForm_OrderByOnKey(frm)
'    stageText = "capturing the saved filter state"
'    Set filterState = SharedForm_CaptureTempVarTextState(filterKey)
'    Set filterOnState = SharedForm_CaptureTempVarTextState(filterOnKey)
'    If CBool(filterState("Existed")) <> CBool(filterOnState("Existed")) Then Err.Raise vbObjectError + 8072, "SharedForm_RestoreFilterAndSort", "The saved Filter state is incomplete because only one member of the pair exists."
'    filterPairExists = CBool(filterState("Existed"))
'    If filterPairExists Then
'        savedFilter = CStr(filterState("Value"))
'        savedFilterOn = SharedValue_CoerceBoolean(filterOnState("Value"))
'        If savedFilterOn And Len(savedFilter) = 0 Then Err.Raise vbObjectError + 8073, "SharedForm_RestoreFilterAndSort", "The saved Filter state enables a blank expression."
'    End If
'    stageText = "capturing the saved OrderBy state"
'    Set orderByState = SharedForm_CaptureTempVarTextState(orderByKey)
'    Set orderByOnState = SharedForm_CaptureTempVarTextState(orderByOnKey)
'    If CBool(orderByState("Existed")) <> CBool(orderByOnState("Existed")) Then Err.Raise vbObjectError + 8074, "SharedForm_RestoreFilterAndSort", "The saved OrderBy state is incomplete because only one member of the pair exists."
'    orderByPairExists = CBool(orderByState("Existed"))
'    If orderByPairExists Then
'        savedOrderBy = CStr(orderByState("Value"))
'        savedOrderByOn = SharedValue_CoerceBoolean(orderByOnState("Value"))
'        If savedOrderByOn And Len(savedOrderBy) = 0 Then Err.Raise vbObjectError + 8075, "SharedForm_RestoreFilterAndSort", "The saved OrderBy state enables a blank expression."
'    End If
'    stageText = "capturing the original live filter and sort state"
'    originalFilter = Nz(frm.Filter, vbNullString)
'    originalFilterOn = CBool(frm.FilterOn)
'    originalOrderBy = Nz(frm.OrderBy, vbNullString)
'    originalOrderByOn = CBool(frm.OrderByOn)
'    liveStateCaptured = True
'    If filterPairExists Then
'        failureArea = 1
'        mutationStarted = True
'        stageText = "applying the complete saved filter state"
'        frm.FilterOn = False
'        frm.Filter = savedFilter
'        frm.FilterOn = savedFilterOn
'        If StrComp(Nz(frm.Filter, vbNullString), savedFilter, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 8076, "SharedForm_RestoreFilterAndSort", "The Filter expression does not match the saved state after application."
'        If CBool(frm.FilterOn) <> savedFilterOn Then Err.Raise vbObjectError + 8077, "SharedForm_RestoreFilterAndSort", "FilterOn does not match the saved state after application."
'        failureArea = 0
'    End If
'    If orderByPairExists Then
'        failureArea = 2
'        mutationStarted = True
'        stageText = "applying the complete saved OrderBy state"
'        frm.OrderByOn = False
'        frm.OrderBy = savedOrderBy
'        frm.OrderByOn = savedOrderByOn
'        If StrComp(Nz(frm.OrderBy, vbNullString), savedOrderBy, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 8078, "SharedForm_RestoreFilterAndSort", "The OrderBy expression does not match the saved state after application."
'        If CBool(frm.OrderByOn) <> savedOrderByOn Then Err.Raise vbObjectError + 8079, "SharedForm_RestoreFilterAndSort", "OrderByOn does not match the saved state after application."
'        failureArea = 0
'    End If
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim rollbackErrNumber As Long
'    Dim rollbackErrDescription As String
'    Dim cleanupErrNumber As Long
'    Dim cleanupErrDescription As String
'    Dim invalidSavedExpression As Boolean
'    Dim recoveryText As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    invalidSavedExpression = failureArea > 0 And SharedForm_FilterSortExpressionErrorIsExpected(savedErrNumber)
'    If mutationStarted And liveStateCaptured Then
'        On Error Resume Next
'        Err.Clear
'        SharedForm_RestoreLiveFilterSortState frm, originalFilter, originalFilterOn, originalOrderBy, originalOrderByOn
'        rollbackErrNumber = Err.Number
'        rollbackErrDescription = Err.Description
'        On Error GoTo 0
'        If rollbackErrNumber <> 0 Then
'            recoveryText = vbCrLf & "Live-state rollback error " & CStr(rollbackErrNumber) & ": " & rollbackErrDescription
'        Else
'            recoveryText = vbCrLf & "The original live filter and sort state was restored and verified."
'        End If
'    Else
'        recoveryText = vbCrLf & "No live-state rollback was required because no live mutation began."
'    End If
'    If invalidSavedExpression Then
'        On Error Resume Next
'        Err.Clear
'        If failureArea = 1 Then
'            SharedForm_RemoveTempVarPairStrict filterKey, filterOnKey
'        ElseIf failureArea = 2 Then
'            SharedForm_RemoveTempVarPairStrict orderByKey, orderByOnKey
'        End If
'        cleanupErrNumber = Err.Number
'        cleanupErrDescription = Err.Description
'        On Error GoTo 0
'        If cleanupErrNumber <> 0 Then
'            recoveryText = recoveryText & vbCrLf & "Invalid saved-state cleanup error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription
'        ElseIf failureArea = 1 Then
'            recoveryText = recoveryText & vbCrLf & "Only the invalid saved filter state was removed; unrelated saved state was preserved."
'        ElseIf failureArea = 2 Then
'            recoveryText = recoveryText & vbCrLf & "Only the invalid saved OrderBy state was removed; unrelated saved state was preserved."
'        End If
'    Else
'        recoveryText = recoveryText & vbCrLf & "Saved TempVar state was preserved because the failure was technical, incomplete, or occurred outside expression application."
'    End If
'    Err.Raise savedErrNumber, "SharedForm_RestoreFilterAndSort", "Could not restore the form filter and sort state." & vbCrLf & "Form: " & formNameText & vbCrLf & "Stage: " & stageText & vbCrLf & "Filter pair exists: " & CStr(filterPairExists) & vbCrLf & "Saved filter: " & savedFilter & vbCrLf & "Saved FilterOn: " & CStr(savedFilterOn) & vbCrLf & "OrderBy pair exists: " & CStr(orderByPairExists) & vbCrLf & "Saved OrderBy: " & savedOrderBy & vbCrLf & "Saved OrderByOn: " & CStr(savedOrderByOn) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & recoveryText
'End Sub
'
'Public Sub SharedForm_ClearFiltersAndSorts(ByVal frm As Access.Form)
'    'Intended use: Atomically clears active and saved form filtering and ordering without closing the form.
'    SharedForm_ClearFilterSortStateAtomic frm, True, True
'End Sub
'
'Private Sub SharedForm_ResetFilter(ByVal frm As Access.Form)
'    'Strictly clears active and saved Filter state for one form through the atomic state helper.
'    SharedForm_ClearFilterSortStateAtomic frm, True, False
'End Sub
'
'Private Sub SharedForm_RemoveSavedFilterAndSortTempVars(ByVal frm As Access.Form)
'    'Removes all four saved-state TempVars atomically and restores their exact Variant states if any removal fails.
'    On Error GoTo ErrHandler
'    Dim states As Collection
'    Dim filterKey As String
'    Dim filterOnKey As String
'    Dim orderByKey As String
'    Dim orderByOnKey As String
'    Dim mutationStarted As Boolean
'    Dim formNameText As String
'    formNameText = "(Nothing)"
'    If Not frm Is Nothing Then formNameText = frm.Name
'    If frm Is Nothing Then Err.Raise vbObjectError + 8044, "SharedForm_RemoveSavedFilterAndSortTempVars", "The Access.Form argument is Nothing."
'    filterKey = SharedForm_FilterKey(frm)
'    filterOnKey = SharedForm_FilterOnKey(frm)
'    orderByKey = SharedForm_OrderByKey(frm)
'    orderByOnKey = SharedForm_OrderByOnKey(frm)
'    Set states = New Collection
'    states.Add SharedForm_CaptureTempVarTextState(filterKey)
'    states.Add SharedForm_CaptureTempVarTextState(filterOnKey)
'    states.Add SharedForm_CaptureTempVarTextState(orderByKey)
'    states.Add SharedForm_CaptureTempVarTextState(orderByOnKey)
'    mutationStarted = True
'    SharedForm_RemoveTempVarIfPresentStrict filterKey
'    SharedForm_RemoveTempVarIfPresentStrict filterOnKey
'    SharedForm_RemoveTempVarIfPresentStrict orderByKey
'    SharedForm_RemoveTempVarIfPresentStrict orderByOnKey
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim rollbackText As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    If mutationStarted Then rollbackText = SharedForm_RestoreTempVarStatesBestEffort(states)
'    If Len(rollbackText) > 0 Then rollbackText = vbCrLf & "Rollback results:" & vbCrLf & rollbackText
'    Err.Raise savedErrNumber, "SharedForm_RemoveSavedFilterAndSortTempVars", "Could not remove the four saved filter and sort TempVars atomically." & vbCrLf & "Form: " & formNameText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & rollbackText
'End Sub
'
'Private Sub SharedForm_PrepareDesignChange(ByVal frm As Access.Form, ByVal actionText As String, ByRef formName As String, ByRef hostFormName As String)
'    'Saves the target record and filter state, then strictly saves and closes the top-level host before a design change.
'    formName = frm.Name
'    hostFormName = SharedForm_TopLevelHostFormName(frm)
'    SharedForm_SaveFilterAndSort frm
'    If frm.Dirty Then frm.Dirty = False
'    If SharedAccess_PrepareOpenObjectForDesignChange(acForm, hostFormName, actionText, True) = False Then Err.Raise vbObjectError + 7733, "SharedForm_PrepareDesignChange", "The host form was not prepared for the requested design change."
'End Sub
'
'Private Function SharedForm_RecoverAfterDesignChangeFailure(ByVal formName As String, ByVal hostFormName As String, ByVal designFormOpened As Boolean, ByVal hostFormClosed As Boolean) As String
'    'Best-effort recovery for a failed form design operation; returns closure, inspection, reopening, and verification details without replacing the primary error.
'    Dim recoveryText As String
'    Dim objectIsOpen As Boolean
'    Dim operationErrNumber As Long
'    Dim operationErrDescription As String
'    If designFormOpened Then
'        On Error Resume Next
'        Err.Clear
'        objectIsOpen = SharedAccess_ObjectIsOpenStrict(acForm, formName)
'        operationErrNumber = Err.Number
'        operationErrDescription = Err.Description
'        On Error GoTo 0
'        If operationErrNumber <> 0 Then
'            recoveryText = SharedText_AppendNonBlankLine(recoveryText, "Could not inspect whether the target Design-view form is open: " & CStr(operationErrNumber) & " - " & operationErrDescription)
'        ElseIf objectIsOpen Then
'            On Error Resume Next
'            Err.Clear
'            DoCmd.Close acForm, formName, acSaveNo
'            operationErrNumber = Err.Number
'            operationErrDescription = Err.Description
'            On Error GoTo 0
'            If operationErrNumber <> 0 Then
'                recoveryText = SharedText_AppendNonBlankLine(recoveryText, "Could not close the target Design-view form without saving: " & CStr(operationErrNumber) & " - " & operationErrDescription)
'            Else
'                On Error Resume Next
'                Err.Clear
'                objectIsOpen = SharedAccess_ObjectIsOpenStrict(acForm, formName)
'                operationErrNumber = Err.Number
'                operationErrDescription = Err.Description
'                On Error GoTo 0
'                If operationErrNumber <> 0 Then
'                    recoveryText = SharedText_AppendNonBlankLine(recoveryText, "Could not verify target-form closure: " & CStr(operationErrNumber) & " - " & operationErrDescription)
'                ElseIf objectIsOpen Then
'                    recoveryText = SharedText_AppendNonBlankLine(recoveryText, "The target Design-view form remained open after the recovery close command.")
'                End If
'            End If
'        End If
'    End If
'    If hostFormClosed Then
'        On Error Resume Next
'        Err.Clear
'        objectIsOpen = SharedAccess_ObjectIsOpenStrict(acForm, hostFormName)
'        operationErrNumber = Err.Number
'        operationErrDescription = Err.Description
'        On Error GoTo 0
'        If operationErrNumber <> 0 Then
'            recoveryText = SharedText_AppendNonBlankLine(recoveryText, "Could not inspect whether the host form is open: " & CStr(operationErrNumber) & " - " & operationErrDescription)
'        ElseIf objectIsOpen = False Then
'            On Error Resume Next
'            Err.Clear
'            DoCmd.OpenForm hostFormName
'            operationErrNumber = Err.Number
'            operationErrDescription = Err.Description
'            On Error GoTo 0
'            If operationErrNumber <> 0 Then
'                recoveryText = SharedText_AppendNonBlankLine(recoveryText, "Could not reopen the host form: " & CStr(operationErrNumber) & " - " & operationErrDescription)
'            Else
'                On Error Resume Next
'                Err.Clear
'                objectIsOpen = SharedAccess_ObjectIsOpenStrict(acForm, hostFormName)
'                operationErrNumber = Err.Number
'                operationErrDescription = Err.Description
'                On Error GoTo 0
'                If operationErrNumber <> 0 Then
'                    recoveryText = SharedText_AppendNonBlankLine(recoveryText, "Could not verify that the host form reopened: " & CStr(operationErrNumber) & " - " & operationErrDescription)
'                ElseIf objectIsOpen = False Then
'                    recoveryText = SharedText_AppendNonBlankLine(recoveryText, "The host form remained closed after the recovery open command.")
'                End If
'            End If
'        End If
'    End If
'    SharedForm_RecoverAfterDesignChangeFailure = recoveryText
'End Function
'
'Public Sub SharedForm_LoadView(ByVal frm As Access.Form)
'    'Intended use: Applies saved form-view state during form loading.
'    SharedForm_RestoreFilterAndSort frm
'    If Not frm.Recordset.EOF Then
'        frm.Recordset.MoveFirst
'    End If
'    frm.SelLeft = 2
'End Sub
'
'Public Sub SharedForm_ClearFiltersSortsAndRefresh(ByVal frm As Access.Form)
'    'Clears the supplied form's live and saved filters and sorts atomically, saves its record, and reopens only its top-level host form.
'    On Error GoTo ErrHandler
'    Dim formName As String
'    Dim hostFormName As String
'    Dim stageText As String
'    Dim hostFormClosed As Boolean
'    formName = frm.Name
'    hostFormName = SharedForm_TopLevelHostFormName(frm)
'    stageText = "clearing live and saved filter and sort state atomically"
'    SharedForm_ClearFiltersAndSorts frm
'    stageText = "saving the current record"
'    If frm.Dirty Then frm.Dirty = False
'    stageText = "saving and closing the host form"
'    hostFormClosed = True
'    If SharedAccess_PrepareOpenObjectForDesignChange(acForm, hostFormName, "clearing filters and sorts", True) = False Then Err.Raise vbObjectError + 7734, "SharedForm_ClearFiltersSortsAndRefresh", "The host form was not prepared for refresh."
'    stageText = "reopening the host form"
'    DoCmd.OpenForm hostFormName
'    hostFormClosed = False
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim recoveryText As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    recoveryText = SharedForm_RecoverAfterDesignChangeFailure(formName, hostFormName, False, hostFormClosed)
'    Err.Raise savedErrNumber, "SharedForm_ClearFiltersSortsAndRefresh", "Could not clear filters and sorts and refresh the requested form." & vbCrLf & "Form: " & formName & vbCrLf & "Host form: " & hostFormName & vbCrLf & "Stage: " & stageText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & IIf(Len(recoveryText) > 0, vbCrLf & "Recovery: " & recoveryText, vbNullString)
'End Sub
'
'Private Function SharedForm_ObjectPropertyExists(ByVal targetObject As Object, ByVal propertyName As String) As Boolean
'    'Returns whether an Access object exposes a named property and raises if the property collection cannot be inspected.
'    On Error GoTo ErrHandler
'    Dim prp As Object
'    If targetObject Is Nothing Then Err.Raise vbObjectError + 8415, "SharedForm_ObjectPropertyExists", "The property target is Nothing."
'    For Each prp In targetObject.Properties
'        If StrComp(CStr(prp.Name), propertyName, vbTextCompare) = 0 Then
'            SharedForm_ObjectPropertyExists = True
'            Exit Function
'        End If
'    Next prp
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedForm_ObjectPropertyExists", "Could not inspect the Access object property collection." & vbCrLf & "Property: " & propertyName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'
'Private Sub SharedForm_SetLongPropertyVerified(ByVal targetObject As Object, ByVal propertyName As String, ByVal propertyValue As Long, ByVal contextText As String)
'    'Writes one numeric Access property and rereads it to prevent false-success reporting.
'    On Error GoTo ErrHandler
'    Dim actualValue As Long
'    targetObject.Properties(propertyName).value = propertyValue
'    actualValue = CLng(targetObject.Properties(propertyName).value)
'    If actualValue <> propertyValue Then Err.Raise vbObjectError + 8416, "SharedForm_SetLongPropertyVerified", "The Access property value did not match after assignment."
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedForm_SetLongPropertyVerified", "Could not write and verify an Access property." & vbCrLf & "Context: " & contextText & vbCrLf & "Property: " & propertyName & vbCrLf & "Requested value: " & CStr(propertyValue) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Sub
'
'Private Sub SharedForm_VerifyLongPropertyValue(ByVal targetObject As Object, ByVal propertyName As String, ByVal expectedValue As Long, ByVal contextText As String)
'    'Rereads one numeric Access property during final validation and raises when later mutations changed the expected result.
'    On Error GoTo ErrHandler
'    Dim actualValue As Long
'    actualValue = CLng(targetObject.Properties(propertyName).value)
'    If actualValue <> expectedValue Then Err.Raise vbObjectError + 8426, "SharedForm_VerifyLongPropertyValue", "The final Access property value does not match the expected value."
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedForm_VerifyLongPropertyValue", "Could not complete final Access property validation." & vbCrLf & "Context: " & contextText & vbCrLf & "Property: " & propertyName & vbCrLf & "Expected value: " & CStr(expectedValue) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Sub
'
'Private Function SharedForm_ChangeControlFontsVerified(ByVal targetForm As Access.Form, ByVal sizeDelta As Integer, ByRef currentControlName As String, ByRef expectedFontSizes As Object) As Long
'    'Changes every supported control FontSize property, records each requested persisted value, and verifies each live write before save.
'    Dim ctl As Access.control
'    Dim oldSize As Long
'    Dim newSize As Long
'    If expectedFontSizes Is Nothing Then Set expectedFontSizes = CreateObject("Scripting.Dictionary")
'    For Each ctl In targetForm.Controls
'        currentControlName = ctl.Name
'        If SharedForm_ObjectPropertyExists(ctl, "FontSize") Then
'            oldSize = CLng(ctl.Properties("FontSize").value)
'            newSize = oldSize + CLng(sizeDelta)
'            If newSize < SHARED_FORM_ZOOM_MIN_FONT Then newSize = SHARED_FORM_ZOOM_MIN_FONT
'            If newSize > SHARED_FORM_ZOOM_MAX_FONT Then newSize = SHARED_FORM_ZOOM_MAX_FONT
'            If newSize <> oldSize Then
'                SharedForm_SetLongPropertyVerified ctl, "FontSize", newSize, "Control " & currentControlName
'                expectedFontSizes.Add currentControlName, newSize
'                SharedForm_ChangeControlFontsVerified = SharedForm_ChangeControlFontsVerified + 1
'            End If
'        End If
'    Next ctl
'End Function
'
'Private Sub SharedForm_VerifyExpectedControlFontSizes(ByVal targetForm As Access.Form, ByVal expectedFontSizes As Object, ByRef currentControlName As String)
'    'Rereads every changed control FontSize from a reopened Design-view form to verify persistence after save.
'    Dim controlNameValue As Variant
'    Dim ctl As Access.control
'    If expectedFontSizes Is Nothing Then Exit Sub
'    For Each controlNameValue In expectedFontSizes.keys
'        currentControlName = CStr(controlNameValue)
'        Set ctl = targetForm.Controls(currentControlName)
'        SharedForm_VerifyLongPropertyValue ctl, "FontSize", CLng(expectedFontSizes(controlNameValue)), "Persisted control " & currentControlName
'    Next controlNameValue
'End Sub
'
'Public Sub SharedForm_ReopenInDefaultView(ByVal frm As Access.Form, ByVal defaultView As AcDefView)
'    'Changes DefaultView, saves it, reopens the form in Design view to verify persistence, and distinguishes uncertain save outcomes from verified commits.
'    On Error GoTo ErrHandler
'    Dim formName As String
'    Dim hostFormName As String
'    Dim stageText As String
'    Dim hostFormClosed As Boolean
'    Dim designFormOpened As Boolean
'    Dim saveAttempted As Boolean
'    Dim saveCompleted As Boolean
'    Dim mutationVerified As Boolean
'    stageText = "preparing the host form"
'    hostFormClosed = True
'    SharedForm_PrepareDesignChange frm, "changing the target form's default view", formName, hostFormName
'    stageText = "opening the target form in Design view"
'    DoCmd.OpenForm formName, acDesign, , , , acHidden
'    designFormOpened = True
'    stageText = "setting and verifying DefaultView"
'    forms(formName).defaultView = defaultView
'    If forms(formName).defaultView <> defaultView Then Err.Raise vbObjectError + 8417, "SharedForm_ReopenInDefaultView", "DefaultView did not match after assignment."
'    stageText = "saving and closing the target Design-view form"
'    saveAttempted = True
'    DoCmd.Close acForm, formName, acSaveYes
'    designFormOpened = False
'    saveCompleted = True
'    stageText = "reopening the target form to verify persisted DefaultView"
'    DoCmd.OpenForm formName, acDesign, , , , acHidden
'    designFormOpened = True
'    If forms(formName).defaultView <> defaultView Then Err.Raise vbObjectError + 8444, "SharedForm_ReopenInDefaultView", "The persisted DefaultView does not match the requested value."
'    DoCmd.Close acForm, formName, acSaveNo
'    designFormOpened = False
'    mutationVerified = True
'    stageText = "reopening the host form"
'    DoCmd.OpenForm hostFormName
'    hostFormClosed = False
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim recoveryText As String
'    Dim outcomeText As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    recoveryText = SharedForm_RecoverAfterDesignChangeFailure(formName, hostFormName, designFormOpened, hostFormClosed)
'    If saveAttempted And saveCompleted = False Then
'        outcomeText = "The save command failed and the persisted DefaultView outcome is uncertain. Inspect the form before rerunning this additive or replacement operation."
'    ElseIf saveCompleted And mutationVerified = False Then
'        outcomeText = "The save command completed, but persisted DefaultView verification failed. The saved form state is uncertain; inspect it before rerunning."
'    ElseIf mutationVerified Then
'        outcomeText = "The requested DefaultView change was saved and verified, but host-form restoration failed. Do not rerun the view change solely because of this error."
'    Else
'        outcomeText = "The requested DefaultView change was not saved; recovery discarded any open unsaved design changes."
'    End If
'    Err.Raise savedErrNumber, "SharedForm_ReopenInDefaultView", outcomeText & vbCrLf & "Form: " & formName & vbCrLf & "Host form: " & hostFormName & vbCrLf & "Stage: " & stageText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & IIf(Len(recoveryText) > 0, vbCrLf & "Recovery: " & recoveryText, vbNullString)
'End Sub
'
'Public Sub SharedForm_ChangeControlFontSizes(ByVal frm As Access.Form, ByVal sizeDelta As Integer)
'    'Changes control font sizes in one save boundary, verifies persisted values, and invalidates any zoom baseline affected by the committed font change.
'    On Error GoTo ErrHandler
'    Dim formName As String
'    Dim hostFormName As String
'    Dim stageText As String
'    Dim hostFormClosed As Boolean
'    Dim designFormOpened As Boolean
'    Dim saveAttempted As Boolean
'    Dim saveCompleted As Boolean
'    Dim mutationVerified As Boolean
'    Dim cacheInvalidated As Boolean
'    Dim changedCount As Long
'    Dim currentControlName As String
'    Dim expectedFontSizes As Object
'    stageText = "preparing the host form"
'    hostFormClosed = True
'    SharedForm_PrepareDesignChange frm, "changing control font sizes", formName, hostFormName
'    stageText = "opening the target form in Design view"
'    DoCmd.OpenForm formName, acDesign, , , , acHidden
'    designFormOpened = True
'    Set expectedFontSizes = CreateObject("Scripting.Dictionary")
'    stageText = "changing and verifying supported control font sizes"
'    changedCount = SharedForm_ChangeControlFontsVerified(forms(formName), sizeDelta, currentControlName, expectedFontSizes)
'    If changedCount > 0 Then
'        stageText = "saving and closing the target Design-view form"
'        saveAttempted = True
'        DoCmd.Close acForm, formName, acSaveYes
'        designFormOpened = False
'        saveCompleted = True
'        stageText = "reopening the target form to verify persisted control font sizes"
'        DoCmd.OpenForm formName, acDesign, , , , acHidden
'        designFormOpened = True
'        SharedForm_VerifyExpectedControlFontSizes forms(formName), expectedFontSizes, currentControlName
'        DoCmd.Close acForm, formName, acSaveNo
'        designFormOpened = False
'        mutationVerified = True
'        stageText = "invalidating the zoom baseline after the committed control-font change"
'        SharedForm_RemoveZoomBaselineForForm formName
'        cacheInvalidated = True
'    Else
'        stageText = "closing the unchanged target Design-view form"
'        DoCmd.Close acForm, formName, acSaveNo
'        designFormOpened = False
'    End If
'    stageText = "reopening the host form"
'    DoCmd.OpenForm hostFormName
'    hostFormClosed = False
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim recoveryText As String
'    Dim cacheRecoveryText As String
'    Dim cacheRecoveryErrNumber As Long
'    Dim cacheRecoveryErrDescription As String
'    Dim outcomeText As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    If changedCount > 0 And saveAttempted And cacheInvalidated = False Then
'        On Error Resume Next
'        Err.Clear
'        SharedForm_RemoveZoomBaselineForForm formName
'        cacheRecoveryErrNumber = Err.Number
'        cacheRecoveryErrDescription = Err.Description
'        cacheInvalidated = (cacheRecoveryErrNumber = 0)
'        On Error GoTo 0
'    End If
'    If cacheInvalidated Then cacheRecoveryText = "The zoom baseline was invalidated because control-font state was committed or may have been committed."
'    If cacheRecoveryErrNumber <> 0 Then cacheRecoveryText = "Zoom-baseline invalidation failed with error " & CStr(cacheRecoveryErrNumber) & ": " & cacheRecoveryErrDescription
'    recoveryText = SharedForm_RecoverAfterDesignChangeFailure(formName, hostFormName, designFormOpened, hostFormClosed)
'    If saveAttempted And saveCompleted = False Then
'        outcomeText = "The save command failed and the persisted control font-size outcome is uncertain. Inspect the form before rerunning the additive size change."
'    ElseIf saveCompleted And mutationVerified = False Then
'        outcomeText = "The save command completed, but persisted control font-size verification failed. Inspect the saved form before rerunning."
'    ElseIf mutationVerified And cacheInvalidated = False Then
'        outcomeText = "The control font-size changes were saved and verified, but the zoom baseline could not be invalidated. Clear the baseline before another scaling operation."
'    ElseIf mutationVerified Then
'        outcomeText = "The control font-size changes were saved and verified and the zoom baseline was invalidated. A later recovery-stage failure occurred; do not rerun the size change solely because of this error."
'    Else
'        outcomeText = "The control font-size changes were not saved; recovery discarded any open unsaved design changes."
'    End If
'    Err.Raise savedErrNumber, "SharedForm_ChangeControlFontSizes", outcomeText & vbCrLf & "Changed controls: " & CStr(changedCount) & vbCrLf & "Form: " & formName & vbCrLf & "Host form: " & hostFormName & vbCrLf & "Stage: " & stageText & vbCrLf & "Control: " & currentControlName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & IIf(Len(cacheRecoveryText) > 0, vbCrLf & "Cache recovery: " & cacheRecoveryText, vbNullString) & IIf(Len(recoveryText) > 0, vbCrLf & "Recovery: " & recoveryText, vbNullString)
'End Sub
'
'Private Function SharedForm_TopLevelHostFormName(ByVal frm As Access.Form) As String
'    'Returns the confirmed top-level host form and raises when parent inspection fails instead of treating an inspection error as a top-level boundary.
'    On Error GoTo ErrHandler
'    Dim currentForm As Access.Form
'    Dim parentObject As Object
'    Dim parentForm As Access.Form
'    Dim partialPath As String
'    If frm Is Nothing Then Err.Raise vbObjectError + 8404, "SharedForm_TopLevelHostFormName", "The Access.Form argument is Nothing."
'    Set currentForm = frm
'    partialPath = currentForm.Name
'    Do
'        Set parentObject = currentForm.Parent
'        If parentObject Is Nothing Then Err.Raise vbObjectError + 8405, "SharedForm_TopLevelHostFormName", "The form Parent property returned Nothing."
'        If StrComp(typeName(parentObject), "SubForm", vbTextCompare) <> 0 Then Exit Do
'        Set parentForm = parentObject.Parent
'        If parentForm Is Nothing Then Err.Raise vbObjectError + 8406, "SharedForm_TopLevelHostFormName", "The subform control Parent property returned Nothing."
'        partialPath = parentForm.Name & "__" & parentObject.Name & "__" & partialPath
'        Set currentForm = parentForm
'    Loop
'    SharedForm_TopLevelHostFormName = currentForm.Name
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedForm_TopLevelHostFormName", "Could not derive the top-level host form." & vbCrLf & "Partial path: " & partialPath & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'
'Public Sub SharedForm_ChangeDatasheetFontSize(ByVal frm As Access.Form, ByVal sizeDelta As Integer)
'    'Changes DatasheetFontHeight and verifies the persisted value after reopening the saved form.
'    On Error GoTo ErrHandler
'    Dim formName As String
'    Dim hostFormName As String
'    Dim oldSize As Long
'    Dim newSize As Long
'    Dim stageText As String
'    Dim hostFormClosed As Boolean
'    Dim designFormOpened As Boolean
'    Dim saveAttempted As Boolean
'    Dim saveCompleted As Boolean
'    Dim mutationVerified As Boolean
'    Dim changed As Boolean
'    stageText = "preparing the host form"
'    hostFormClosed = True
'    SharedForm_PrepareDesignChange frm, "changing the datasheet font size", formName, hostFormName
'    stageText = "opening the target form in Design view"
'    DoCmd.OpenForm formName, acDesign, , , , acHidden
'    designFormOpened = True
'    stageText = "changing and verifying DatasheetFontHeight"
'    If SharedForm_ObjectPropertyExists(forms(formName), "DatasheetFontHeight") = False Then Err.Raise vbObjectError + 8418, "SharedForm_ChangeDatasheetFontSize", "The target form does not expose DatasheetFontHeight."
'    oldSize = CLng(forms(formName).Properties("DatasheetFontHeight").value)
'    newSize = oldSize + CLng(sizeDelta)
'    If newSize < SHARED_FORM_ZOOM_MIN_FONT Then newSize = SHARED_FORM_ZOOM_MIN_FONT
'    If newSize > SHARED_FORM_ZOOM_MAX_FONT Then newSize = SHARED_FORM_ZOOM_MAX_FONT
'    If newSize <> oldSize Then
'        SharedForm_SetLongPropertyVerified forms(formName), "DatasheetFontHeight", newSize, "Form " & formName
'        changed = True
'    End If
'    If changed Then
'        stageText = "saving and closing the target Design-view form"
'        saveAttempted = True
'        DoCmd.Close acForm, formName, acSaveYes
'        designFormOpened = False
'        saveCompleted = True
'        stageText = "reopening the target form to verify persisted DatasheetFontHeight"
'        DoCmd.OpenForm formName, acDesign, , , , acHidden
'        designFormOpened = True
'        SharedForm_VerifyLongPropertyValue forms(formName), "DatasheetFontHeight", newSize, "Persisted form " & formName
'        DoCmd.Close acForm, formName, acSaveNo
'        designFormOpened = False
'        mutationVerified = True
'    Else
'        stageText = "closing the unchanged target Design-view form"
'        DoCmd.Close acForm, formName, acSaveNo
'        designFormOpened = False
'    End If
'    stageText = "reopening the host form"
'    DoCmd.OpenForm hostFormName
'    hostFormClosed = False
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim recoveryText As String
'    Dim outcomeText As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    recoveryText = SharedForm_RecoverAfterDesignChangeFailure(formName, hostFormName, designFormOpened, hostFormClosed)
'    If saveAttempted And saveCompleted = False Then
'        outcomeText = "The save command failed and the persisted datasheet font-size outcome is uncertain. Inspect the form before rerunning the additive size change."
'    ElseIf saveCompleted And mutationVerified = False Then
'        outcomeText = "The save command completed, but persisted DatasheetFontHeight verification failed. Inspect the saved form before rerunning."
'    ElseIf mutationVerified Then
'        outcomeText = "The datasheet font-size change was saved and verified, but host-form restoration failed. Do not rerun the size change solely because of this error."
'    Else
'        outcomeText = "The datasheet font-size change was not saved; recovery discarded any open unsaved design changes."
'    End If
'    Err.Raise savedErrNumber, "SharedForm_ChangeDatasheetFontSize", outcomeText & vbCrLf & "Form: " & formName & vbCrLf & "Host form: " & hostFormName & vbCrLf & "Stage: " & stageText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & IIf(Len(recoveryText) > 0, vbCrLf & "Recovery: " & recoveryText, vbNullString)
'End Sub
'
'Public Sub SharedForm_IncreaseDatasheetFontSize(ByVal frm As Access.Form)
'    'Intended use: Increases a form datasheet font size by the shared increment.
'    SharedForm_ChangeDatasheetFontSize frm, 1
'End Sub
'
'Public Sub SharedForm_DecreaseDatasheetFontSize(ByVal frm As Access.Form)
'    'Intended use: Decreases a form datasheet font size by the shared increment.
'    SharedForm_ChangeDatasheetFontSize frm, -1
'End Sub
'
'Public Sub SharedForm_RefreshAll(ByVal frm As Access.Form)
'    'Saves, refreshes, and requeries the supplied form without relying on the active Access object.
'    On Error GoTo ErrHandler
'    Dim formName As String
'    Dim stageText As String
'    formName = frm.Name
'    stageText = "saving filter and sort state"
'    SharedForm_SaveFilterAndSort frm
'    stageText = "saving the current record"
'    If frm.Dirty Then frm.Dirty = False
'    stageText = "refreshing the supplied form"
'    frm.Refresh
'    stageText = "requerying the supplied form"
'    frm.Requery
'    stageText = "restoring filter and sort state"
'    SharedForm_RestoreFilterAndSort frm
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedForm_RefreshAll", "Could not refresh and requery the supplied form." & vbCrLf & "Form: " & formName & vbCrLf & "Stage: " & stageText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Sub
'
'Private Sub SharedForm_RestorePersistentFilterSortMetadata(ByVal formName As String, ByVal hostFormName As String, ByRef designFormOpened As Boolean, ByVal originalFilter As String, ByVal originalFilterOn As Boolean, ByVal originalOrderBy As String, ByVal originalOrderByOn As Boolean)
'    'Restores and verifies the persistent filter and sort properties after an uncommitted metadata-removal failure, without silently discarding unrelated host-form changes.
'    On Error GoTo ErrHandler
'    Dim stageText As String
'    Dim objectIsOpen As Boolean
'    stageText = "discarding the failed target Design-view mutation"
'    If designFormOpened Then
'        objectIsOpen = SharedAccess_ObjectIsOpenStrict(acForm, formName)
'        If objectIsOpen Then DoCmd.Close acForm, formName, acSaveNo
'        If SharedAccess_ObjectIsOpenStrict(acForm, formName) Then Err.Raise vbObjectError + 8500, "SharedForm_RestorePersistentFilterSortMetadata", "The failed target Design-view form remained open."
'        designFormOpened = False
'    End If
'    stageText = "safely closing the host form for metadata restoration"
'    If SharedAccess_ObjectIsOpenStrict(acForm, hostFormName) Then
'        If SharedAccess_PrepareOpenObjectForDesignChange(acForm, hostFormName, "restoring filter and sort metadata after a failed removal", True) = False Then Err.Raise vbObjectError + 8501, "SharedForm_RestorePersistentFilterSortMetadata", "The host form was not prepared for metadata restoration."
'    End If
'    stageText = "safely closing a separately open target form"
'    If StrComp(formName, hostFormName, vbTextCompare) <> 0 Then
'        If SharedAccess_ObjectIsOpenStrict(acForm, formName) Then
'            If SharedAccess_PrepareOpenObjectForDesignChange(acForm, formName, "restoring filter and sort metadata after a failed removal", True) = False Then Err.Raise vbObjectError + 8502, "SharedForm_RestorePersistentFilterSortMetadata", "The target form was not prepared for metadata restoration."
'        End If
'    End If
'    stageText = "opening the target form to restore persistent metadata"
'    DoCmd.OpenForm formName, acDesign, , , , acHidden
'    designFormOpened = True
'    SharedForm_RestoreLiveFilterSortState forms(formName), originalFilter, originalFilterOn, originalOrderBy, originalOrderByOn
'    stageText = "saving the restored persistent metadata"
'    DoCmd.Close acForm, formName, acSaveYes
'    designFormOpened = False
'    stageText = "reopening the target form to verify restored persistent metadata"
'    DoCmd.OpenForm formName, acDesign, , , , acHidden
'    designFormOpened = True
'    If StrComp(Nz(forms(formName).Filter, vbNullString), originalFilter, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 8503, "SharedForm_RestorePersistentFilterSortMetadata", "The restored persistent Filter expression does not match the original state."
'    If CBool(forms(formName).FilterOn) <> originalFilterOn Then Err.Raise vbObjectError + 8504, "SharedForm_RestorePersistentFilterSortMetadata", "The restored persistent FilterOn value does not match the original state."
'    If StrComp(Nz(forms(formName).OrderBy, vbNullString), originalOrderBy, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 8505, "SharedForm_RestorePersistentFilterSortMetadata", "The restored persistent OrderBy expression does not match the original state."
'    If CBool(forms(formName).OrderByOn) <> originalOrderByOn Then Err.Raise vbObjectError + 8506, "SharedForm_RestorePersistentFilterSortMetadata", "The restored persistent OrderByOn value does not match the original state."
'    DoCmd.Close acForm, formName, acSaveNo
'    designFormOpened = False
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim cleanupErrNumber As Long
'    Dim cleanupErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    Err.Clear
'    If designFormOpened Then
'        DoCmd.Close acForm, formName, acSaveNo
'        cleanupErrNumber = Err.Number
'        cleanupErrDescription = Err.Description
'        If cleanupErrNumber = 0 Then
'            Err.Clear
'            If SharedAccess_ObjectIsOpenStrict(acForm, formName) = False And Err.Number = 0 Then designFormOpened = False
'            If Err.Number <> 0 Then
'                cleanupErrNumber = Err.Number
'                cleanupErrDescription = Err.Description
'            End If
'        End If
'    End If
'    On Error GoTo 0
'    Err.Raise savedErrNumber, "SharedForm_RestorePersistentFilterSortMetadata", "Could not restore and verify the original persistent filter and sort metadata." & vbCrLf & "Form: " & formName & vbCrLf & "Host form: " & hostFormName & vbCrLf & "Stage: " & stageText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & IIf(cleanupErrNumber <> 0, vbCrLf & "Recovery-close error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription, vbNullString)
'End Sub
'
'Public Sub SharedForm_RemoveMetadata(ByVal frm As Access.Form)
'    'Removes persistent form filter and sort metadata as a verified operation and restores the exact original state after any pre-commit failure.
'    On Error GoTo ErrHandler
'    Dim formName As String
'    Dim hostFormName As String
'    Dim ctl As Access.control
'    Dim originalTempVarStates As Collection
'    Dim originalFilter As String
'    Dim originalFilterOn As Boolean
'    Dim originalOrderBy As String
'    Dim originalOrderByOn As Boolean
'    Dim stageText As String
'    Dim hostFormClosed As Boolean
'    Dim designFormOpened As Boolean
'    Dim stateMutationStarted As Boolean
'    Dim designPreparationStarted As Boolean
'    Dim saveAttempted As Boolean
'    Dim saveCompleted As Boolean
'    Dim metadataCommitted As Boolean
'    formName = frm.Name
'    hostFormName = SharedForm_TopLevelHostFormName(frm)
'    stageText = "capturing the original live filter and sort state"
'    originalFilter = Nz(frm.Filter, vbNullString)
'    originalFilterOn = CBool(frm.FilterOn)
'    originalOrderBy = Nz(frm.OrderBy, vbNullString)
'    originalOrderByOn = CBool(frm.OrderByOn)
'    stageText = "capturing the original saved filter and sort state"
'    Set originalTempVarStates = New Collection
'    originalTempVarStates.Add SharedForm_CaptureTempVarTextState(SharedForm_FilterKey(frm))
'    originalTempVarStates.Add SharedForm_CaptureTempVarTextState(SharedForm_FilterOnKey(frm))
'    originalTempVarStates.Add SharedForm_CaptureTempVarTextState(SharedForm_OrderByKey(frm))
'    originalTempVarStates.Add SharedForm_CaptureTempVarTextState(SharedForm_OrderByOnKey(frm))
'    stageText = "saving the current record"
'    If frm.Dirty Then frm.Dirty = False
'    stageText = "clearing runtime and saved filter and sort metadata atomically"
'    stateMutationStarted = True
'    SharedForm_ClearFilterSortStateAtomic frm, True, True
'    For Each ctl In frm.Controls
'        On Error Resume Next
'        ctl.ControlTipText = ctl.ControlTipText
'        Err.Clear
'        On Error GoTo ErrHandler
'    Next ctl
'    stageText = "saving and closing the host form"
'    designPreparationStarted = True
'    hostFormClosed = True
'    If SharedAccess_PrepareOpenObjectForDesignChange(acForm, hostFormName, "removing persistent form filter and sort metadata", True) = False Then Err.Raise vbObjectError + 7735, "SharedForm_RemoveMetadata", "The host form was not prepared for metadata removal."
'    stageText = "opening the target form in Design view"
'    DoCmd.OpenForm formName, acDesign, , , , acHidden
'    designFormOpened = True
'    SharedForm_RestoreLiveFilterSortState forms(formName), vbNullString, False, vbNullString, False
'    stageText = "saving and closing the target Design-view form"
'    saveAttempted = True
'    DoCmd.Close acForm, formName, acSaveYes
'    designFormOpened = False
'    saveCompleted = True
'    stageText = "reopening the target form to verify persistent metadata removal"
'    DoCmd.OpenForm formName, acDesign, , , , acHidden
'    designFormOpened = True
'    If Len(Nz(forms(formName).Filter, vbNullString)) > 0 Then Err.Raise vbObjectError + 8507, "SharedForm_RemoveMetadata", "The persistent Filter expression remained populated after metadata removal."
'    If CBool(forms(formName).FilterOn) Then Err.Raise vbObjectError + 8508, "SharedForm_RemoveMetadata", "Persistent FilterOn remained enabled after metadata removal."
'    If Len(Nz(forms(formName).OrderBy, vbNullString)) > 0 Then Err.Raise vbObjectError + 8509, "SharedForm_RemoveMetadata", "The persistent OrderBy expression remained populated after metadata removal."
'    If CBool(forms(formName).OrderByOn) Then Err.Raise vbObjectError + 8510, "SharedForm_RemoveMetadata", "Persistent OrderByOn remained enabled after metadata removal."
'    DoCmd.Close acForm, formName, acSaveNo
'    designFormOpened = False
'    metadataCommitted = True
'    stageText = "reopening the host form"
'    DoCmd.OpenForm hostFormName
'    If SharedAccess_ObjectIsOpenStrict(acForm, hostFormName) = False Then Err.Raise vbObjectError + 8511, "SharedForm_RemoveMetadata", "The host form remained closed after the reopen command."
'    hostFormClosed = False
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim recoveryText As String
'    Dim rollbackText As String
'    Dim tempVarRollbackText As String
'    Dim rollbackErrNumber As Long
'    Dim rollbackErrDescription As String
'    Dim liveRollbackErrNumber As Long
'    Dim liveRollbackErrDescription As String
'    Dim hostInspectionErrNumber As Long
'    Dim hostInspectionErrDescription As String
'    Dim hostIsOpen As Boolean
'    Dim directRollback As Boolean
'    Dim outcomeText As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    If stateMutationStarted And metadataCommitted = False Then
'        On Error Resume Next
'        Err.Clear
'        hostIsOpen = SharedAccess_ObjectIsOpenStrict(acForm, hostFormName)
'        hostInspectionErrNumber = Err.Number
'        hostInspectionErrDescription = Err.Description
'        On Error GoTo 0
'        directRollback = (hostInspectionErrNumber = 0 And hostIsOpen And designFormOpened = False And designPreparationStarted = False And saveAttempted = False)
'        If directRollback Then
'            tempVarRollbackText = SharedForm_RestoreTempVarStatesBestEffort(originalTempVarStates)
'            On Error Resume Next
'            Err.Clear
'            SharedForm_RestoreLiveFilterSortState frm, originalFilter, originalFilterOn, originalOrderBy, originalOrderByOn
'            liveRollbackErrNumber = Err.Number
'            liveRollbackErrDescription = Err.Description
'            On Error GoTo 0
'            If liveRollbackErrNumber = 0 And InStr(1, tempVarRollbackText, "failed", vbTextCompare) = 0 And InStr(1, tempVarRollbackText, "unavailable", vbTextCompare) = 0 Then
'                rollbackText = "The original live and saved filter and sort state was restored and verified."
'            Else
'                rollbackText = "The direct state rollback was incomplete."
'                If liveRollbackErrNumber <> 0 Then rollbackText = rollbackText & vbCrLf & "Live-state rollback error " & CStr(liveRollbackErrNumber) & ": " & liveRollbackErrDescription
'            End If
'        Else
'            On Error Resume Next
'            Err.Clear
'            SharedForm_RestorePersistentFilterSortMetadata formName, hostFormName, designFormOpened, originalFilter, originalFilterOn, originalOrderBy, originalOrderByOn
'            rollbackErrNumber = Err.Number
'            rollbackErrDescription = Err.Description
'            On Error GoTo 0
'            tempVarRollbackText = SharedForm_RestoreTempVarStatesBestEffort(originalTempVarStates)
'            If rollbackErrNumber = 0 And InStr(1, tempVarRollbackText, "failed", vbTextCompare) = 0 And InStr(1, tempVarRollbackText, "unavailable", vbTextCompare) = 0 Then
'                rollbackText = "The original persistent and saved filter and sort state was restored and verified."
'            Else
'                rollbackText = "The persistent-state rollback was incomplete."
'                If rollbackErrNumber <> 0 Then rollbackText = rollbackText & vbCrLf & "Persistent rollback error " & CStr(rollbackErrNumber) & ": " & rollbackErrDescription
'            End If
'        End If
'        If hostInspectionErrNumber <> 0 Then rollbackText = rollbackText & vbCrLf & "Host-open inspection error " & CStr(hostInspectionErrNumber) & ": " & hostInspectionErrDescription
'        If Len(tempVarRollbackText) > 0 Then rollbackText = rollbackText & vbCrLf & "Saved-state rollback results:" & vbCrLf & tempVarRollbackText
'        recoveryText = SharedForm_RecoverAfterDesignChangeFailure(formName, hostFormName, designFormOpened, True)
'        outcomeText = "The metadata removal did not commit. Recovery attempted to restore the exact original state before reporting the failure."
'    ElseIf metadataCommitted Then
'        recoveryText = SharedForm_RecoverAfterDesignChangeFailure(formName, hostFormName, designFormOpened, hostFormClosed)
'        outcomeText = "The persistent metadata removal was saved and reverified. A later host-restoration failure occurred; do not rerun the removal solely because of this error."
'    Else
'        recoveryText = SharedForm_RecoverAfterDesignChangeFailure(formName, hostFormName, designFormOpened, hostFormClosed)
'        outcomeText = "The metadata-removal mutation did not begin; no filter or sort rollback was required."
'    End If
'    Err.Raise savedErrNumber, "SharedForm_RemoveMetadata", outcomeText & vbCrLf & "Form: " & formName & vbCrLf & "Host form: " & hostFormName & vbCrLf & "Stage: " & stageText & vbCrLf & "Save attempted: " & CStr(saveAttempted) & vbCrLf & "Save completed: " & CStr(saveCompleted) & vbCrLf & "Metadata committed: " & CStr(metadataCommitted) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & IIf(Len(rollbackText) > 0, vbCrLf & "Rollback: " & rollbackText, vbNullString) & IIf(Len(recoveryText) > 0, vbCrLf & "Recovery: " & recoveryText, vbNullString)
'End Sub
'
'Public Function SharedForm_TempVarText(ByVal varName As String) As String
'    'Intended use: Returns a TempVar value as text and returns blank when the TempVar is unavailable.
'    On Error GoTo MissingTempVar
'    SharedForm_TempVarText = CStr(TempVars(varName))
'    Exit Function
'MissingTempVar:
'    SharedForm_TempVarText = vbNullString
'End Function
'
'Private Function SharedForm_StateScopeKey(ByVal frm As Access.Form) As String
'    'Returns a stable top-level-host and subform-control path and raises when parent inspection fails.
'    On Error GoTo ErrHandler
'    Dim currentForm As Access.Form
'    Dim parentObject As Object
'    Dim parentForm As Access.Form
'    Dim scopeText As String
'    Dim formNameText As String
'    formNameText = "(Nothing)"
'    If Not frm Is Nothing Then formNameText = frm.Name
'    If frm Is Nothing Then Err.Raise vbObjectError + 8080, "SharedForm_StateScopeKey", "The Access.Form argument is Nothing."
'    Set currentForm = frm
'    scopeText = currentForm.Name
'    Do
'        Set parentObject = currentForm.Parent
'        If parentObject Is Nothing Then Err.Raise vbObjectError + 8407, "SharedForm_StateScopeKey", "The form Parent property returned Nothing."
'        If StrComp(typeName(parentObject), "SubForm", vbTextCompare) <> 0 Then Exit Do
'        Set parentForm = parentObject.Parent
'        If parentForm Is Nothing Then Err.Raise vbObjectError + 8408, "SharedForm_StateScopeKey", "The subform control Parent property returned Nothing."
'        scopeText = parentForm.Name & "__" & parentObject.Name & "__" & scopeText
'        Set currentForm = parentForm
'    Loop
'    SharedForm_StateScopeKey = scopeText
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedForm_StateScopeKey", "Could not derive the stable host-qualified form-state key." & vbCrLf & "Form: " & formNameText & vbCrLf & "Partial scope: " & scopeText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'
'Private Function SharedForm_FilterKey(ByVal frm As Access.Form) As String
'    SharedForm_FilterKey = SharedForm_StateScopeKey(frm) & "_FormFilter"
'End Function
'
'Private Function SharedForm_FilterOnKey(ByVal frm As Access.Form) As String
'    SharedForm_FilterOnKey = SharedForm_StateScopeKey(frm) & "_FormFilterOn"
'End Function
'
'Private Function SharedForm_OrderByKey(ByVal frm As Access.Form) As String
'    SharedForm_OrderByKey = SharedForm_StateScopeKey(frm) & "_FormOrderBy"
'End Function
'
'Private Function SharedForm_OrderByOnKey(ByVal frm As Access.Form) As String
'    SharedForm_OrderByOnKey = SharedForm_StateScopeKey(frm) & "_FormOrderByOn"
'End Function
'
'Public Sub SharedForm_LoadViewMaximized(ByVal frm As Access.Form)
'    'Intended use: Applies saved view state and maximizes the loaded form.
'        DoCmd.Maximize
'        SharedForm_LoadView frm
'End Sub
'
'Public Sub SharedForm_QueueFilterRebuild(ByRef pendingFilterRebuild As Boolean, ByRef savedTimerInterval As Long, ByVal frm As Access.Form)
'    'Intended use: Atomically queues deferred filter rebuilding, including rollback of timer, caller variables, and saved filter/sort TempVars.
'    On Error GoTo ErrHandler
'    Dim originalPending As Boolean
'    Dim originalSavedTimerInterval As Long
'    Dim originalTimerInterval As Long
'    Dim timerChanged As Boolean
'    Dim tempVarStates As Collection
'    Dim tempVarStatesCaptured As Boolean
'    Dim tempVarRollbackText As String
'    Dim filterKey As String
'    Dim filterOnKey As String
'    Dim orderByKey As String
'    Dim orderByOnKey As String
'    If frm Is Nothing Then Err.Raise vbObjectError + 8409, "SharedForm_QueueFilterRebuild", "The Access.Form argument is Nothing."
'    If pendingFilterRebuild Then Exit Sub
'    originalPending = pendingFilterRebuild
'    originalSavedTimerInterval = savedTimerInterval
'    originalTimerInterval = frm.TimerInterval
'    filterKey = SharedForm_FilterKey(frm)
'    filterOnKey = SharedForm_FilterOnKey(frm)
'    orderByKey = SharedForm_OrderByKey(frm)
'    orderByOnKey = SharedForm_OrderByOnKey(frm)
'    Set tempVarStates = New Collection
'    tempVarStates.Add SharedForm_CaptureTempVarTextState(filterKey)
'    tempVarStates.Add SharedForm_CaptureTempVarTextState(filterOnKey)
'    tempVarStates.Add SharedForm_CaptureTempVarTextState(orderByKey)
'    tempVarStates.Add SharedForm_CaptureTempVarTextState(orderByOnKey)
'    tempVarStatesCaptured = True
'    frm.TimerInterval = 1
'    timerChanged = True
'    If frm.TimerInterval <> 1 Then Err.Raise vbObjectError + 8410, "SharedForm_QueueFilterRebuild", "The temporary TimerInterval was not applied."
'    SharedForm_SaveFilterAndSort frm
'    savedTimerInterval = originalTimerInterval
'    pendingFilterRebuild = True
'    If pendingFilterRebuild = False Then Err.Raise vbObjectError + 8434, "SharedForm_QueueFilterRebuild", "The pending filter-rebuild flag was not published."
'    If savedTimerInterval <> originalTimerInterval Then Err.Raise vbObjectError + 8439, "SharedForm_QueueFilterRebuild", "The original TimerInterval was not published to the caller state."
'    If frm.TimerInterval <> 1 Then Err.Raise vbObjectError + 8440, "SharedForm_QueueFilterRebuild", "The queued temporary TimerInterval did not remain active."
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim restoreErrNumber As Long
'    Dim restoreErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    pendingFilterRebuild = originalPending
'    savedTimerInterval = originalSavedTimerInterval
'    If tempVarStatesCaptured Then tempVarRollbackText = SharedForm_RestoreTempVarStatesBestEffort(tempVarStates)
'    If timerChanged Then
'        On Error Resume Next
'        Err.Clear
'        frm.TimerInterval = originalTimerInterval
'        If Err.Number = 0 Then
'            If frm.TimerInterval <> originalTimerInterval Then Err.Raise vbObjectError + 8411, "SharedForm_QueueFilterRebuild", "TimerInterval restoration did not verify."
'        End If
'        restoreErrNumber = Err.Number
'        restoreErrDescription = Err.Description
'        On Error GoTo 0
'    End If
'    Err.Raise savedErrNumber, "SharedForm_QueueFilterRebuild", "Could not queue the filter rebuild without changing caller, timer, or saved filter/sort state." & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & IIf(Len(tempVarRollbackText) > 0, vbCrLf & "TempVar rollback:" & vbCrLf & tempVarRollbackText, vbNullString) & IIf(restoreErrNumber <> 0, vbCrLf & "Timer restoration error " & CStr(restoreErrNumber) & ": " & restoreErrDescription, vbNullString)
'End Sub
'
'Public Sub SharedForm_Timer(ByRef pendingFilterRebuild As Boolean, ByRef savedTimerInterval As Long, ByVal frm As Access.Form)
'    'Intended use: Processes one queued filter rebuild after first restoring the prior timer; failures cancel automatic retry so no rapid timer loop remains.
'    On Error GoTo ErrHandler
'    Dim requestedRestoreInterval As Long
'    If frm Is Nothing Then Err.Raise vbObjectError + 8412, "SharedForm_Timer", "The Access.Form argument is Nothing."
'    requestedRestoreInterval = savedTimerInterval
'    If pendingFilterRebuild = False Then
'        If frm.TimerInterval = 1 Then
'            frm.TimerInterval = requestedRestoreInterval
'            If frm.TimerInterval <> requestedRestoreInterval Then Err.Raise vbObjectError + 8413, "SharedForm_Timer", "The saved TimerInterval was not restored while clearing stray timer state."
'        End If
'        Exit Sub
'    End If
'    frm.TimerInterval = requestedRestoreInterval
'    If frm.TimerInterval <> requestedRestoreInterval Then Err.Raise vbObjectError + 8441, "SharedForm_Timer", "The saved TimerInterval was not restored before queued work began."
'    SharedForm_SaveFilterAndSort frm
'    pendingFilterRebuild = False
'    If pendingFilterRebuild Then Err.Raise vbObjectError + 8442, "SharedForm_Timer", "The completed pending state did not clear."
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim recoveryErrNumber As Long
'    Dim recoveryErrDescription As String
'    Dim fallbackErrNumber As Long
'    Dim fallbackErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    pendingFilterRebuild = False
'    On Error Resume Next
'    Err.Clear
'    frm.TimerInterval = requestedRestoreInterval
'    If Err.Number = 0 Then
'        If frm.TimerInterval <> requestedRestoreInterval Then Err.Raise vbObjectError + 8414, "SharedForm_Timer", "TimerInterval recovery did not verify."
'    End If
'    recoveryErrNumber = Err.Number
'    recoveryErrDescription = Err.Description
'    If recoveryErrNumber <> 0 Then
'        Err.Clear
'        frm.TimerInterval = 0
'        If Err.Number = 0 Then
'            If frm.TimerInterval <> 0 Then Err.Raise vbObjectError + 8443, "SharedForm_Timer", "The zero-interval fallback did not stop the timer."
'        End If
'        fallbackErrNumber = Err.Number
'        fallbackErrDescription = Err.Description
'    End If
'    On Error GoTo 0
'    Err.Raise savedErrNumber, "SharedForm_Timer", "Could not process the queued filter rebuild. The pending request was canceled and no automatic retry remains." & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & IIf(recoveryErrNumber <> 0, vbCrLf & "Saved-timer restoration error " & CStr(recoveryErrNumber) & ": " & recoveryErrDescription, vbNullString) & IIf(fallbackErrNumber <> 0, vbCrLf & "Timer-stop fallback error " & CStr(fallbackErrNumber) & ": " & fallbackErrDescription, vbNullString)
'End Sub
'
'Public Sub SharedForm_Close(ByVal frm As Access.Form, Optional ByVal saveFilterAndSort As Boolean = True)
'    'Intended use: Saves optional filter/sort state and closes the supplied form instance.
'    If saveFilterAndSort Then
'        SharedForm_SaveFilterAndSort frm
'    End If
'    DoCmd.Close acForm, frm.Name, acSaveNo
'End Sub
'
'Public Sub SharedForm_Maximize(ByVal frm As Access.Form)
'    'Intended use: Maximizes the supplied form window.
'    DoCmd.SelectObject acForm, frm.Name, False
'    DoCmd.Maximize
'End Sub
'
'Public Sub SharedForm_Minimize(ByVal frm As Access.Form)
'    'Intended use: Minimizes the supplied form window.
'    DoCmd.SelectObject acForm, frm.Name, False
'    DoCmd.Minimize
'End Sub
'
'Public Sub SharedForm_ShowDatasheetView(ByVal frm As Access.Form)
'    'Intended use: Reopens or switches the supplied form to Datasheet view.
'    SharedForm_ReopenInDefaultView frm, acDefViewDatasheet
'End Sub
'
'Public Sub SharedForm_ShowSplitView(ByVal frm As Access.Form)
'    'Intended use: Reopens or switches the supplied form to Split Form view.
'    SharedForm_ReopenInDefaultView frm, acDefViewSplitForm
'End Sub
'
'Public Sub SharedForm_ShowContinuousView(ByVal frm As Access.Form)
'    'Intended use: Reopens or switches the supplied form to Continuous Forms view.
'    SharedForm_ReopenInDefaultView frm, acDefViewContinuous
'End Sub
'
'Public Sub SharedForm_ShowSingleView(ByVal frm As Access.Form)
'    'Intended use: Reopens or switches the supplied form to Single Form view.
'    SharedForm_ReopenInDefaultView frm, acDefViewSingle
'End Sub
'
'Public Sub SharedForm_DebugSources(ByVal frm As Access.Form)
'    'Intended use: Prints form, control, and record-source details for manual troubleshooting.
'    Dim ctl As Access.control
'    Debug.Print "FORM: "; frm.Name
'    Debug.Print "RecordSource: "; frm.recordSource
'    Debug.Print "Filter: "; frm.Filter
'    Debug.Print "OrderBy: "; frm.OrderBy
'    For Each ctl In frm.Controls
'        On Error Resume Next
'        Debug.Print ctl.Name; " ControlSource="; ctl.controlSource
'        Debug.Print ctl.Name; " RowSource="; ctl.rowSource
'        On Error GoTo 0
'    Next ctl
'End Sub
'
'Public Sub SharedForm_ChangeAllFontSizes(ByVal frm As Access.Form, ByVal sizeDelta As Integer)
'    'Intended use: Changes control and datasheet font sizes in one save boundary, verifies all persisted values, and invalidates the zoom baseline only when control fonts changed.
'    On Error GoTo ErrHandler
'    Dim formName As String
'    Dim hostFormName As String
'    Dim stageText As String
'    Dim hostFormClosed As Boolean
'    Dim designFormOpened As Boolean
'    Dim saveAttempted As Boolean
'    Dim saveCompleted As Boolean
'    Dim mutationVerified As Boolean
'    Dim cacheInvalidated As Boolean
'    Dim changedCount As Long
'    Dim controlChangedCount As Long
'    Dim oldDatasheetSize As Long
'    Dim newDatasheetSize As Long
'    Dim datasheetChanged As Boolean
'    Dim currentControlName As String
'    Dim expectedFontSizes As Object
'    stageText = "preparing the host form"
'    hostFormClosed = True
'    SharedForm_PrepareDesignChange frm, "changing all form font sizes", formName, hostFormName
'    stageText = "opening the target form in Design view"
'    DoCmd.OpenForm formName, acDesign, , , , acHidden
'    designFormOpened = True
'    Set expectedFontSizes = CreateObject("Scripting.Dictionary")
'    stageText = "changing and verifying supported control font sizes"
'    controlChangedCount = SharedForm_ChangeControlFontsVerified(forms(formName), sizeDelta, currentControlName, expectedFontSizes)
'    changedCount = controlChangedCount
'    stageText = "changing and verifying DatasheetFontHeight"
'    If SharedForm_ObjectPropertyExists(forms(formName), "DatasheetFontHeight") = False Then Err.Raise vbObjectError + 8425, "SharedForm_ChangeAllFontSizes", "The target form does not expose DatasheetFontHeight."
'    oldDatasheetSize = CLng(forms(formName).Properties("DatasheetFontHeight").value)
'    newDatasheetSize = oldDatasheetSize + CLng(sizeDelta)
'    If newDatasheetSize < SHARED_FORM_ZOOM_MIN_FONT Then newDatasheetSize = SHARED_FORM_ZOOM_MIN_FONT
'    If newDatasheetSize > SHARED_FORM_ZOOM_MAX_FONT Then newDatasheetSize = SHARED_FORM_ZOOM_MAX_FONT
'    If newDatasheetSize <> oldDatasheetSize Then
'        SharedForm_SetLongPropertyVerified forms(formName), "DatasheetFontHeight", newDatasheetSize, "Form " & formName
'        datasheetChanged = True
'        changedCount = changedCount + 1
'    End If
'    If changedCount > 0 Then
'        stageText = "saving and closing the target Design-view form"
'        saveAttempted = True
'        DoCmd.Close acForm, formName, acSaveYes
'        designFormOpened = False
'        saveCompleted = True
'        stageText = "reopening the target form to verify all persisted font sizes"
'        DoCmd.OpenForm formName, acDesign, , , , acHidden
'        designFormOpened = True
'        SharedForm_VerifyExpectedControlFontSizes forms(formName), expectedFontSizes, currentControlName
'        If datasheetChanged Then SharedForm_VerifyLongPropertyValue forms(formName), "DatasheetFontHeight", newDatasheetSize, "Persisted form " & formName
'        DoCmd.Close acForm, formName, acSaveNo
'        designFormOpened = False
'        mutationVerified = True
'        If controlChangedCount > 0 Then
'            stageText = "invalidating the zoom baseline after committed control-font changes"
'            SharedForm_RemoveZoomBaselineForForm formName
'            cacheInvalidated = True
'        End If
'    Else
'        stageText = "closing the unchanged target Design-view form"
'        DoCmd.Close acForm, formName, acSaveNo
'        designFormOpened = False
'    End If
'    stageText = "reopening the host form"
'    DoCmd.OpenForm hostFormName
'    hostFormClosed = False
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim recoveryText As String
'    Dim cacheRecoveryText As String
'    Dim cacheRecoveryErrNumber As Long
'    Dim cacheRecoveryErrDescription As String
'    Dim outcomeText As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    If controlChangedCount > 0 And saveAttempted And cacheInvalidated = False Then
'        On Error Resume Next
'        Err.Clear
'        SharedForm_RemoveZoomBaselineForForm formName
'        cacheRecoveryErrNumber = Err.Number
'        cacheRecoveryErrDescription = Err.Description
'        cacheInvalidated = (cacheRecoveryErrNumber = 0)
'        On Error GoTo 0
'    End If
'    If cacheInvalidated Then cacheRecoveryText = "The zoom baseline was invalidated because control-font state was committed or may have been committed."
'    If cacheRecoveryErrNumber <> 0 Then cacheRecoveryText = "Zoom-baseline invalidation failed with error " & CStr(cacheRecoveryErrNumber) & ": " & cacheRecoveryErrDescription
'    recoveryText = SharedForm_RecoverAfterDesignChangeFailure(formName, hostFormName, designFormOpened, hostFormClosed)
'    If saveAttempted And saveCompleted = False Then
'        outcomeText = "The combined save command failed and the persisted font-size outcome is uncertain. Inspect the form before rerunning the additive size change."
'    ElseIf saveCompleted And mutationVerified = False Then
'        outcomeText = "The combined save completed, but persisted font-size verification failed. Inspect the saved form before rerunning."
'    ElseIf mutationVerified And controlChangedCount > 0 And cacheInvalidated = False Then
'        outcomeText = "The combined font-size changes were saved and verified, but the zoom baseline could not be invalidated. Clear the baseline before another scaling operation."
'    ElseIf mutationVerified And controlChangedCount > 0 Then
'        outcomeText = "The combined font-size changes were saved and verified and the zoom baseline was invalidated. A later recovery-stage failure occurred; do not rerun the size change solely because of this error."
'    ElseIf mutationVerified Then
'        outcomeText = "The datasheet-only font-size change was saved and verified, but host-form restoration failed. Do not rerun the size change solely because of this error."
'    Else
'        outcomeText = "The combined font-size changes were not saved; recovery discarded any open unsaved design changes."
'    End If
'    Err.Raise savedErrNumber, "SharedForm_ChangeAllFontSizes", outcomeText & vbCrLf & "Changed properties: " & CStr(changedCount) & vbCrLf & "Changed control fonts: " & CStr(controlChangedCount) & vbCrLf & "Form: " & formName & vbCrLf & "Host form: " & hostFormName & vbCrLf & "Stage: " & stageText & vbCrLf & "Control: " & currentControlName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & IIf(Len(cacheRecoveryText) > 0, vbCrLf & "Cache recovery: " & cacheRecoveryText, vbNullString) & IIf(Len(recoveryText) > 0, vbCrLf & "Recovery: " & recoveryText, vbNullString)
'End Sub
'
'Public Sub SharedForm_ChangeControlFontSizesSingleForm(ByVal frm As Access.Form, ByVal sizeDelta As Integer)
'    'Intended use: Retains the single-form public entry point while using the verified shared control-font implementation.
'    SharedForm_ChangeControlFontSizes frm, sizeDelta
'End Sub
'
'Public Sub SharedForm_IncreaseControlFontSizes(ByVal frm As Access.Form)
'    'Intended use: Increases supported control font sizes on the supplied form.
'    SharedForm_ChangeControlFontSizes frm, 1
'End Sub
'
'Public Sub SharedForm_DecreaseControlFontSizes(ByVal frm As Access.Form)
'    'Intended use: Decreases supported control font sizes on the supplied form.
'    SharedForm_ChangeControlFontSizes frm, -1
'End Sub
'
'Private Sub SharedForm_EnsureZoomStores()
'    If mFormZoomGeometry Is Nothing Then Set mFormZoomGeometry = CreateObject("Scripting.Dictionary")
'    If mFormZoomFonts Is Nothing Then Set mFormZoomFonts = CreateObject("Scripting.Dictionary")
'    If mFormZoomLevels Is Nothing Then Set mFormZoomLevels = CreateObject("Scripting.Dictionary")
'End Sub
'
'Private Function SharedForm_ZoomKey(ByVal formName As String, ByVal itemName As String, ByVal propName As String) As String
'    SharedForm_ZoomKey = formName & "|" & itemName & "|" & propName
'End Function
'
'Private Function SharedForm_ZoomGeometryExists(ByVal formName As String, ByVal itemName As String, ByVal propName As String) As Boolean
'    SharedForm_EnsureZoomStores
'    SharedForm_ZoomGeometryExists = mFormZoomGeometry.Exists(SharedForm_ZoomKey(formName, itemName, propName))
'End Function
'
'Private Function SharedForm_ZoomFontExists(ByVal formName As String, ByVal itemName As String) As Boolean
'    SharedForm_EnsureZoomStores
'    SharedForm_ZoomFontExists = mFormZoomFonts.Exists(SharedForm_ZoomKey(formName, itemName, "FontSize"))
'End Function
'
'Private Function SharedForm_ZoomGeometryValue(ByVal formName As String, ByVal itemName As String, ByVal propName As String) As Double
'    SharedForm_EnsureZoomStores
'    SharedForm_ZoomGeometryValue = CDbl(mFormZoomGeometry(SharedForm_ZoomKey(formName, itemName, propName)))
'End Function
'
'Private Function SharedForm_ZoomFontValue(ByVal formName As String, ByVal itemName As String) As Integer
'    SharedForm_EnsureZoomStores
'    SharedForm_ZoomFontValue = CInt(mFormZoomFonts(SharedForm_ZoomKey(formName, itemName, "FontSize")))
'End Function
'
'Private Function SharedForm_ZoomLevelValue(ByVal formName As String) As Long
'    SharedForm_EnsureZoomStores
'    If mFormZoomLevels.Exists(formName) Then
'        SharedForm_ZoomLevelValue = CLng(mFormZoomLevels(formName))
'    Else
'        SharedForm_ZoomLevelValue = 0
'    End If
'End Function
'
'Private Sub SharedForm_SetZoomLevelValue(ByVal formName As String, ByVal zoomLevel As Long)
'    SharedForm_EnsureZoomStores
'    If mFormZoomLevels.Exists(formName) Then
'        mFormZoomLevels(formName) = zoomLevel
'    Else
'        mFormZoomLevels.Add formName, zoomLevel
'    End If
'    If CLng(mFormZoomLevels(formName)) <> zoomLevel Then Err.Raise vbObjectError + 8419, "SharedForm_SetZoomLevelValue", "The zoom-level cache did not verify after assignment."
'End Sub
'
'Private Sub SharedForm_SetZoomGeometryValue(ByVal formName As String, ByVal itemName As String, ByVal propName As String, ByVal propValue As Variant)
'    Dim keyText As String
'    SharedForm_EnsureZoomStores
'    keyText = SharedForm_ZoomKey(formName, itemName, propName)
'    If mFormZoomGeometry.Exists(keyText) Then
'        mFormZoomGeometry(keyText) = CDbl(propValue)
'    Else
'        mFormZoomGeometry.Add keyText, CDbl(propValue)
'    End If
'End Sub
'
'Private Sub SharedForm_SetZoomFontValue(ByVal formName As String, ByVal itemName As String, ByVal fontSizeValue As Variant)
'    Dim keyText As String
'    SharedForm_EnsureZoomStores
'    keyText = SharedForm_ZoomKey(formName, itemName, "FontSize")
'    If mFormZoomFonts.Exists(keyText) Then
'        mFormZoomFonts(keyText) = CInt(fontSizeValue)
'    Else
'        mFormZoomFonts.Add keyText, CInt(fontSizeValue)
'    End If
'End Sub
'
'Private Sub SharedForm_RemoveZoomBaselineForForm(ByVal formName As String)
'    'Removes every geometry, font, and level entry for one form so a failed or explicitly cleared baseline cannot remain partially visible.
'    Dim keyValue As Variant
'    Dim keysToRemove As Collection
'    SharedForm_EnsureZoomStores
'    Set keysToRemove = New Collection
'    For Each keyValue In mFormZoomGeometry.keys
'        If Left$(CStr(keyValue), Len(formName) + 1) = formName & "|" Then keysToRemove.Add CStr(keyValue)
'    Next keyValue
'    For Each keyValue In keysToRemove
'        mFormZoomGeometry.Remove CStr(keyValue)
'    Next keyValue
'    Set keysToRemove = New Collection
'    For Each keyValue In mFormZoomFonts.keys
'        If Left$(CStr(keyValue), Len(formName) + 1) = formName & "|" Then keysToRemove.Add CStr(keyValue)
'    Next keyValue
'    For Each keyValue In keysToRemove
'        mFormZoomFonts.Remove CStr(keyValue)
'    Next keyValue
'    If mFormZoomLevels.Exists(formName) Then mFormZoomLevels.Remove formName
'End Sub
'
'Private Sub SharedForm_CaptureSectionHeightIfPresent(ByVal frm As Access.Form, ByVal sectionNumber As AcSection, ByVal itemName As String, ByVal geometryValues As Object)
'    'Captures an optional form section height, treating only the normal missing-section error as absence.
'    Dim sectionHeight As Long
'    Dim inspectionErrNumber As Long
'    Dim inspectionErrSource As String
'    Dim inspectionErrDescription As String
'    On Error Resume Next
'    Err.Clear
'    sectionHeight = frm.Section(sectionNumber).Height
'    inspectionErrNumber = Err.Number
'    inspectionErrSource = Err.Source
'    inspectionErrDescription = Err.Description
'    Err.Clear
'    On Error GoTo 0
'    If inspectionErrNumber = 0 Then
'        geometryValues.Add SharedForm_ZoomKey(frm.Name, itemName, "Height"), CDbl(sectionHeight)
'    ElseIf inspectionErrNumber <> 2462 Then
'        Err.Raise inspectionErrNumber, inspectionErrSource, inspectionErrDescription
'    End If
'End Sub
'
'Private Sub SharedForm_ValidateZoomBaselineCoverage(ByVal formName As String)
'    'Validates that the current Design-view form exactly matches the cached baseline at its recorded zoom level and that no controls or sections are missing or obsolete.
'    On Error GoTo ErrHandler
'    Dim frm As Access.Form
'    Dim ctl As Access.control
'    Dim expectedGeometryKeys As Object
'    Dim expectedFontKeys As Object
'    Dim currentSectionValues As Object
'    Dim keyValue As Variant
'    Dim keyText As String
'    Dim prefixText As String
'    Dim currentZoomLevel As Long
'    Dim layoutScale As Double
'    Dim expectedValue As Long
'    Dim currentControlName As String
'    SharedForm_EnsureZoomStores
'    If mFormZoomLevels.Exists(formName) = False Then Err.Raise vbObjectError + 8445, "SharedForm_ValidateZoomBaselineCoverage", "The zoom-level baseline is missing."
'    Set frm = forms(formName)
'    Set expectedGeometryKeys = CreateObject("Scripting.Dictionary")
'    Set expectedFontKeys = CreateObject("Scripting.Dictionary")
'    Set currentSectionValues = CreateObject("Scripting.Dictionary")
'    currentZoomLevel = CLng(mFormZoomLevels(formName))
'    layoutScale = SharedForm_GetZoomScale(formName, currentZoomLevel)
'    SharedForm_CaptureSectionHeightIfPresent frm, acHeader, "_SECTION_HEADER", currentSectionValues
'    SharedForm_CaptureSectionHeightIfPresent frm, acDetail, "_SECTION_DETAIL", currentSectionValues
'    SharedForm_CaptureSectionHeightIfPresent frm, acFooter, "_SECTION_FOOTER", currentSectionValues
'    For Each keyValue In currentSectionValues.keys
'        keyText = CStr(keyValue)
'        expectedGeometryKeys.Add keyText, True
'        If mFormZoomGeometry.Exists(keyText) = False Then Err.Raise vbObjectError + 8446, "SharedForm_ValidateZoomBaselineCoverage", "A current form section is missing from the zoom baseline: " & keyText
'        expectedValue = CLng(CDbl(mFormZoomGeometry(keyText)) * layoutScale)
'        If CLng(currentSectionValues(keyText)) <> expectedValue Then Err.Raise vbObjectError + 8447, "SharedForm_ValidateZoomBaselineCoverage", "A current form section no longer matches the recorded zoom baseline: " & keyText
'    Next keyValue
'    For Each ctl In frm.Controls
'        currentControlName = ctl.Name
'        keyText = SharedForm_ZoomKey(formName, ctl.Name, "Left")
'        expectedGeometryKeys.Add keyText, True
'        If mFormZoomGeometry.Exists(keyText) = False Then Err.Raise vbObjectError + 8448, "SharedForm_ValidateZoomBaselineCoverage", "A current control is missing its Left baseline: " & ctl.Name
'        If CLng(ctl.Left) <> CLng(CDbl(mFormZoomGeometry(keyText)) * layoutScale) Then Err.Raise vbObjectError + 8449, "SharedForm_ValidateZoomBaselineCoverage", "A current control Left value no longer matches the recorded zoom baseline: " & ctl.Name
'        keyText = SharedForm_ZoomKey(formName, ctl.Name, "Top")
'        expectedGeometryKeys.Add keyText, True
'        If mFormZoomGeometry.Exists(keyText) = False Then Err.Raise vbObjectError + 8450, "SharedForm_ValidateZoomBaselineCoverage", "A current control is missing its Top baseline: " & ctl.Name
'        If CLng(ctl.Top) <> CLng(CDbl(mFormZoomGeometry(keyText)) * layoutScale) Then Err.Raise vbObjectError + 8451, "SharedForm_ValidateZoomBaselineCoverage", "A current control Top value no longer matches the recorded zoom baseline: " & ctl.Name
'        keyText = SharedForm_ZoomKey(formName, ctl.Name, "Width")
'        expectedGeometryKeys.Add keyText, True
'        If mFormZoomGeometry.Exists(keyText) = False Then Err.Raise vbObjectError + 8452, "SharedForm_ValidateZoomBaselineCoverage", "A current control is missing its Width baseline: " & ctl.Name
'        If CLng(ctl.Width) <> CLng(CDbl(mFormZoomGeometry(keyText)) * layoutScale) Then Err.Raise vbObjectError + 8453, "SharedForm_ValidateZoomBaselineCoverage", "A current control Width value no longer matches the recorded zoom baseline: " & ctl.Name
'        keyText = SharedForm_ZoomKey(formName, ctl.Name, "Height")
'        expectedGeometryKeys.Add keyText, True
'        If mFormZoomGeometry.Exists(keyText) = False Then Err.Raise vbObjectError + 8454, "SharedForm_ValidateZoomBaselineCoverage", "A current control is missing its Height baseline: " & ctl.Name
'        If CLng(ctl.Height) <> CLng(CDbl(mFormZoomGeometry(keyText)) * layoutScale) Then Err.Raise vbObjectError + 8455, "SharedForm_ValidateZoomBaselineCoverage", "A current control Height value no longer matches the recorded zoom baseline: " & ctl.Name
'        If SharedForm_ObjectPropertyExists(ctl, "FontSize") Then
'            keyText = SharedForm_ZoomKey(formName, ctl.Name, "FontSize")
'            expectedFontKeys.Add keyText, True
'            If mFormZoomFonts.Exists(keyText) = False Then Err.Raise vbObjectError + 8456, "SharedForm_ValidateZoomBaselineCoverage", "A current control is missing its FontSize baseline: " & ctl.Name
'            expectedValue = CInt(mFormZoomFonts(keyText)) + currentZoomLevel
'            If expectedValue < SHARED_FORM_ZOOM_MIN_FONT Then expectedValue = SHARED_FORM_ZOOM_MIN_FONT
'            If expectedValue > SHARED_FORM_ZOOM_MAX_FONT Then expectedValue = SHARED_FORM_ZOOM_MAX_FONT
'            If CLng(ctl.Properties("FontSize").value) <> expectedValue Then Err.Raise vbObjectError + 8457, "SharedForm_ValidateZoomBaselineCoverage", "A current control FontSize no longer matches the recorded zoom baseline: " & ctl.Name
'        End If
'    Next ctl
'    prefixText = formName & "|"
'    For Each keyValue In mFormZoomGeometry.keys
'        keyText = CStr(keyValue)
'        If Left$(keyText, Len(prefixText)) = prefixText Then
'            If expectedGeometryKeys.Exists(keyText) = False Then Err.Raise vbObjectError + 8458, "SharedForm_ValidateZoomBaselineCoverage", "The zoom baseline contains obsolete geometry for a removed control or section: " & keyText
'        End If
'    Next keyValue
'    For Each keyValue In mFormZoomFonts.keys
'        keyText = CStr(keyValue)
'        If Left$(keyText, Len(prefixText)) = prefixText Then
'            If expectedFontKeys.Exists(keyText) = False Then Err.Raise vbObjectError + 8459, "SharedForm_ValidateZoomBaselineCoverage", "The zoom baseline contains obsolete font data for a removed or changed control: " & keyText
'        End If
'    Next keyValue
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedForm_ValidateZoomBaselineCoverage", "The existing zoom baseline is stale or incomplete. Clear the scaling baseline before applying another scale operation." & vbCrLf & "Form: " & formName & vbCrLf & "Control: " & currentControlName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Sub
'
'Private Sub SharedForm_CaptureZoomBaseline(ByVal formName As String)
'    'Builds a complete temporary zoom baseline, validates an existing baseline before reuse, and publishes new state only after every required value is captured.
'    On Error GoTo ErrHandler
'    Dim ctl As Access.control
'    Dim frm As Access.Form
'    Dim geometryValues As Object
'    Dim fontValues As Object
'    Dim keyValue As Variant
'    Dim currentControlName As String
'    Dim publicationStarted As Boolean
'    SharedForm_EnsureZoomStores
'    If mFormZoomLevels.Exists(formName) Then
'        SharedForm_ValidateZoomBaselineCoverage formName
'        Exit Sub
'    End If
'    Set frm = forms(formName)
'    Set geometryValues = CreateObject("Scripting.Dictionary")
'    Set fontValues = CreateObject("Scripting.Dictionary")
'    SharedForm_CaptureSectionHeightIfPresent frm, acHeader, "_SECTION_HEADER", geometryValues
'    SharedForm_CaptureSectionHeightIfPresent frm, acDetail, "_SECTION_DETAIL", geometryValues
'    SharedForm_CaptureSectionHeightIfPresent frm, acFooter, "_SECTION_FOOTER", geometryValues
'    For Each ctl In frm.Controls
'        currentControlName = ctl.Name
'        geometryValues.Add SharedForm_ZoomKey(formName, ctl.Name, "Left"), CDbl(ctl.Left)
'        geometryValues.Add SharedForm_ZoomKey(formName, ctl.Name, "Top"), CDbl(ctl.Top)
'        geometryValues.Add SharedForm_ZoomKey(formName, ctl.Name, "Width"), CDbl(ctl.Width)
'        geometryValues.Add SharedForm_ZoomKey(formName, ctl.Name, "Height"), CDbl(ctl.Height)
'        If SharedForm_ObjectPropertyExists(ctl, "FontSize") Then fontValues.Add SharedForm_ZoomKey(formName, ctl.Name, "FontSize"), CInt(ctl.Properties("FontSize").value)
'    Next ctl
'    publicationStarted = True
'    SharedForm_RemoveZoomBaselineForForm formName
'    For Each keyValue In geometryValues.keys
'        mFormZoomGeometry.Add CStr(keyValue), CDbl(geometryValues(keyValue))
'    Next keyValue
'    For Each keyValue In fontValues.keys
'        mFormZoomFonts.Add CStr(keyValue), CInt(fontValues(keyValue))
'    Next keyValue
'    mFormZoomLevels.Add formName, 0
'    For Each keyValue In geometryValues.keys
'        If mFormZoomGeometry.Exists(CStr(keyValue)) = False Then Err.Raise vbObjectError + 8427, "SharedForm_CaptureZoomBaseline", "A published geometry baseline key is missing."
'        If CDbl(mFormZoomGeometry(CStr(keyValue))) <> CDbl(geometryValues(keyValue)) Then Err.Raise vbObjectError + 8428, "SharedForm_CaptureZoomBaseline", "A published geometry baseline value does not match."
'    Next keyValue
'    For Each keyValue In fontValues.keys
'        If mFormZoomFonts.Exists(CStr(keyValue)) = False Then Err.Raise vbObjectError + 8429, "SharedForm_CaptureZoomBaseline", "A published font baseline key is missing."
'        If CInt(mFormZoomFonts(CStr(keyValue))) <> CInt(fontValues(keyValue)) Then Err.Raise vbObjectError + 8430, "SharedForm_CaptureZoomBaseline", "A published font baseline value does not match."
'    Next keyValue
'    If mFormZoomLevels.Exists(formName) = False Then Err.Raise vbObjectError + 8431, "SharedForm_CaptureZoomBaseline", "The published zoom level is missing."
'    If CLng(mFormZoomLevels(formName)) <> 0 Then Err.Raise vbObjectError + 8432, "SharedForm_CaptureZoomBaseline", "The published initial zoom level is not zero."
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    If publicationStarted Then
'        On Error Resume Next
'        SharedForm_RemoveZoomBaselineForForm formName
'        On Error GoTo 0
'    End If
'    Err.Raise savedErrNumber, "SharedForm_CaptureZoomBaseline", "Could not capture, validate, or publish a complete form zoom baseline." & vbCrLf & "Form: " & formName & vbCrLf & "Control: " & currentControlName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Sub
'
'Private Function SharedForm_CanonicalZoomLevel(ByVal formName As String, ByVal requestedZoomLevel As Long) As Long
'    'Returns the single canonical checkpoint for the effective clamped control-font state represented by the requested level.
'    Dim ctl As Access.control
'    Dim baseFontSize As Long
'    Dim lowerCandidate As Long
'    Dim upperCandidate As Long
'    Dim lowerBound As Long
'    Dim upperBound As Long
'    Dim boundsInitialized As Boolean
'    For Each ctl In forms(formName).Controls
'        If SharedForm_ZoomFontExists(formName, ctl.Name) Then
'            baseFontSize = CLng(SharedForm_ZoomFontValue(formName, ctl.Name))
'            lowerCandidate = CLng(SHARED_FORM_ZOOM_MIN_FONT) - baseFontSize
'            upperCandidate = CLng(SHARED_FORM_ZOOM_MAX_FONT) - baseFontSize
'            If boundsInitialized = False Then
'                lowerBound = lowerCandidate
'                upperBound = upperCandidate
'                boundsInitialized = True
'            Else
'                If lowerCandidate < lowerBound Then lowerBound = lowerCandidate
'                If upperCandidate > upperBound Then upperBound = upperCandidate
'            End If
'        End If
'    Next ctl
'    If boundsInitialized = False Then
'        SharedForm_CanonicalZoomLevel = 0
'    ElseIf requestedZoomLevel < lowerBound Then
'        SharedForm_CanonicalZoomLevel = lowerBound
'    ElseIf requestedZoomLevel > upperBound Then
'        SharedForm_CanonicalZoomLevel = upperBound
'    Else
'        SharedForm_CanonicalZoomLevel = requestedZoomLevel
'    End If
'End Function
'
'Private Function SharedForm_GetZoomScale(ByVal formName As String, ByVal zoomLevel As Long) As Double
'    Dim ctl As Access.control
'    Dim baseFontSize As Integer
'    Dim targetFontSize As Integer
'    Dim scaleTotal As Double
'    Dim scaleCount As Long
'    For Each ctl In forms(formName).Controls
'        If SharedForm_ZoomFontExists(formName, ctl.Name) Then
'            baseFontSize = SharedForm_ZoomFontValue(formName, ctl.Name)
'            targetFontSize = baseFontSize + zoomLevel
'            If targetFontSize < SHARED_FORM_ZOOM_MIN_FONT Then targetFontSize = SHARED_FORM_ZOOM_MIN_FONT
'            If targetFontSize > SHARED_FORM_ZOOM_MAX_FONT Then targetFontSize = SHARED_FORM_ZOOM_MAX_FONT
'            If baseFontSize > 0 Then
'                scaleTotal = scaleTotal + (CDbl(targetFontSize) / CDbl(baseFontSize))
'                scaleCount = scaleCount + 1
'            End If
'        End If
'    Next ctl
'    If scaleCount = 0 Then
'        SharedForm_GetZoomScale = 1
'    Else
'        SharedForm_GetZoomScale = scaleTotal / CDbl(scaleCount)
'    End If
'End Function
'
'Private Sub SharedForm_ApplyOrVerifyScaledState(ByVal formName As String, ByVal targetZoomLevel As Long, ByVal applyValues As Boolean, ByRef currentControlName As String)
'    'Applies or rereads the complete scaled section, geometry, and font state using the authoritative cached baseline.
'    Dim ctl As Access.control
'    Dim layoutScale As Double
'    Dim baseFontSize As Long
'    Dim targetFontSize As Long
'    Dim targetValue As Long
'    layoutScale = SharedForm_GetZoomScale(formName, targetZoomLevel)
'    If SharedForm_ZoomGeometryExists(formName, "_SECTION_HEADER", "Height") Then
'        targetValue = CLng(SharedForm_ZoomGeometryValue(formName, "_SECTION_HEADER", "Height") * layoutScale)
'        If applyValues Then
'            SharedForm_SetLongPropertyVerified forms(formName).Section(acHeader), "Height", targetValue, "Form header section"
'        Else
'            SharedForm_VerifyLongPropertyValue forms(formName).Section(acHeader), "Height", targetValue, "Persisted form header section"
'        End If
'    End If
'    If SharedForm_ZoomGeometryExists(formName, "_SECTION_DETAIL", "Height") Then
'        targetValue = CLng(SharedForm_ZoomGeometryValue(formName, "_SECTION_DETAIL", "Height") * layoutScale)
'        If applyValues Then
'            SharedForm_SetLongPropertyVerified forms(formName).Section(acDetail), "Height", targetValue, "Form detail section"
'        Else
'            SharedForm_VerifyLongPropertyValue forms(formName).Section(acDetail), "Height", targetValue, "Persisted form detail section"
'        End If
'    End If
'    If SharedForm_ZoomGeometryExists(formName, "_SECTION_FOOTER", "Height") Then
'        targetValue = CLng(SharedForm_ZoomGeometryValue(formName, "_SECTION_FOOTER", "Height") * layoutScale)
'        If applyValues Then
'            SharedForm_SetLongPropertyVerified forms(formName).Section(acFooter), "Height", targetValue, "Form footer section"
'        Else
'            SharedForm_VerifyLongPropertyValue forms(formName).Section(acFooter), "Height", targetValue, "Persisted form footer section"
'        End If
'    End If
'    For Each ctl In forms(formName).Controls
'        currentControlName = ctl.Name
'        If SharedForm_ZoomGeometryExists(formName, ctl.Name, "Left") Then
'            targetValue = CLng(SharedForm_ZoomGeometryValue(formName, ctl.Name, "Left") * layoutScale)
'            If applyValues Then
'                SharedForm_SetLongPropertyVerified ctl, "Left", targetValue, "Control " & ctl.Name
'            Else
'                SharedForm_VerifyLongPropertyValue ctl, "Left", targetValue, "Persisted control " & ctl.Name
'            End If
'        End If
'        If SharedForm_ZoomGeometryExists(formName, ctl.Name, "Top") Then
'            targetValue = CLng(SharedForm_ZoomGeometryValue(formName, ctl.Name, "Top") * layoutScale)
'            If applyValues Then
'                SharedForm_SetLongPropertyVerified ctl, "Top", targetValue, "Control " & ctl.Name
'            Else
'                SharedForm_VerifyLongPropertyValue ctl, "Top", targetValue, "Persisted control " & ctl.Name
'            End If
'        End If
'        If SharedForm_ZoomGeometryExists(formName, ctl.Name, "Width") Then
'            targetValue = CLng(SharedForm_ZoomGeometryValue(formName, ctl.Name, "Width") * layoutScale)
'            If applyValues Then
'                SharedForm_SetLongPropertyVerified ctl, "Width", targetValue, "Control " & ctl.Name
'            Else
'                SharedForm_VerifyLongPropertyValue ctl, "Width", targetValue, "Persisted control " & ctl.Name
'            End If
'        End If
'        If SharedForm_ZoomGeometryExists(formName, ctl.Name, "Height") Then
'            targetValue = CLng(SharedForm_ZoomGeometryValue(formName, ctl.Name, "Height") * layoutScale)
'            If applyValues Then
'                SharedForm_SetLongPropertyVerified ctl, "Height", targetValue, "Control " & ctl.Name
'            Else
'                SharedForm_VerifyLongPropertyValue ctl, "Height", targetValue, "Persisted control " & ctl.Name
'            End If
'        End If
'        If SharedForm_ZoomFontExists(formName, ctl.Name) Then
'            baseFontSize = SharedForm_ZoomFontValue(formName, ctl.Name)
'            targetFontSize = baseFontSize + targetZoomLevel
'            If targetFontSize < SHARED_FORM_ZOOM_MIN_FONT Then targetFontSize = SHARED_FORM_ZOOM_MIN_FONT
'            If targetFontSize > SHARED_FORM_ZOOM_MAX_FONT Then targetFontSize = SHARED_FORM_ZOOM_MAX_FONT
'            If applyValues Then
'                SharedForm_SetLongPropertyVerified ctl, "FontSize", targetFontSize, "Control " & ctl.Name
'            Else
'                SharedForm_VerifyLongPropertyValue ctl, "FontSize", targetFontSize, "Persisted control " & ctl.Name
'            End If
'        End If
'    Next ctl
'End Sub
'
'Public Sub SharedForm_ScaleControlFontsAndLayout(ByVal frm As Access.Form, ByVal sizeDelta As Integer)
'    'Scales from the canonical current zoom checkpoint and prevents checkpoint drift when all effective font values are clamped.
'    SharedForm_SetControlFontsAndLayoutZoom frm, CLng(sizeDelta), False
'End Sub
'
'Private Sub SharedForm_SetControlFontsAndLayoutZoom(ByVal frm As Access.Form, ByVal levelValue As Long, ByVal absoluteTarget As Boolean)
'    'Applies an incremental or absolute canonical zoom target, verifies persisted state, and publishes only the effective checkpoint represented by the saved form. Absolute targets require an existing authoritative
'    'baseline; incremental targets may capture one.
'    On Error GoTo ErrHandler
'    Dim formName As String
'    Dim hostFormName As String
'    Dim priorZoomLevel As Long
'    Dim canonicalPriorZoomLevel As Long
'    Dim requestedZoomLevel As Long
'    Dim targetZoomLevel As Long
'    Dim stageText As String
'    Dim hostFormClosed As Boolean
'    Dim designFormOpened As Boolean
'    Dim saveAttempted As Boolean
'    Dim saveCompleted As Boolean
'    Dim formStateChanged As Boolean
'    Dim formStateVerified As Boolean
'    Dim baselineExisted As Boolean
'    Dim cacheStateCaptured As Boolean
'    Dim cacheMutationStarted As Boolean
'    Dim cachePublished As Boolean
'    Dim currentControlName As String
'    If frm Is Nothing Then Err.Raise vbObjectError + 8460, "SharedForm_ScaleControlFontsAndLayout", "The Access.Form argument is Nothing."
'    formName = frm.Name
'    SharedForm_EnsureZoomStores
'    baselineExisted = mFormZoomLevels.Exists(formName)
'    If absoluteTarget And baselineExisted = False Then
'        On Error GoTo 0
'        Err.Raise vbObjectError + 8461, "SharedForm_ResetAndScaleControlFonts", "Cannot restore zoom baseline level zero because no authoritative zoom baseline exists for form '" & formName & "'. The baseline may have been cleared, invalidated by another control-font operation, or lost after a VBA project reset. The current form state was not adopted as a replacement baseline."
'    End If
'    If baselineExisted Then priorZoomLevel = CLng(mFormZoomLevels(formName))
'    cacheStateCaptured = True
'    stageText = "preparing the host form"
'    hostFormClosed = True
'    SharedForm_PrepareDesignChange frm, "scaling control fonts and layout", formName, hostFormName
'    stageText = "opening the target form in Design view"
'    DoCmd.OpenForm formName, acDesign, , , , acHidden
'    designFormOpened = True
'    If baselineExisted Then
'        stageText = "validating existing zoom baseline coverage and freshness"
'        SharedForm_ValidateZoomBaselineCoverage formName
'    Else
'        stageText = "capturing the complete zoom baseline"
'        cacheMutationStarted = True
'        SharedForm_CaptureZoomBaseline formName
'        priorZoomLevel = 0
'    End If
'    canonicalPriorZoomLevel = SharedForm_CanonicalZoomLevel(formName, priorZoomLevel)
'    If absoluteTarget Then
'        requestedZoomLevel = levelValue
'    Else
'        requestedZoomLevel = canonicalPriorZoomLevel + levelValue
'    End If
'    targetZoomLevel = SharedForm_CanonicalZoomLevel(formName, requestedZoomLevel)
'    If targetZoomLevel <> canonicalPriorZoomLevel Then
'        formStateChanged = True
'        stageText = "applying and verifying the requested canonical scaled state"
'        SharedForm_ApplyOrVerifyScaledState formName, targetZoomLevel, True, currentControlName
'        stageText = "saving and closing the target Design-view form"
'        saveAttempted = True
'        DoCmd.Close acForm, formName, acSaveYes
'        designFormOpened = False
'        saveCompleted = True
'        stageText = "reopening the target form to verify persisted canonical scaled state"
'        DoCmd.OpenForm formName, acDesign, , , , acHidden
'        designFormOpened = True
'        SharedForm_ApplyOrVerifyScaledState formName, targetZoomLevel, False, currentControlName
'        DoCmd.Close acForm, formName, acSaveNo
'        designFormOpened = False
'        formStateVerified = True
'    Else
'        stageText = "verifying the already-current canonical scaled state"
'        SharedForm_ApplyOrVerifyScaledState formName, targetZoomLevel, False, currentControlName
'        DoCmd.Close acForm, formName, acSaveNo
'        designFormOpened = False
'        formStateVerified = True
'    End If
'    stageText = "publishing the verified canonical zoom level"
'    cacheMutationStarted = True
'    SharedForm_SetZoomLevelValue formName, targetZoomLevel
'    cachePublished = True
'    stageText = "reopening the host form"
'    DoCmd.OpenForm hostFormName
'    hostFormClosed = False
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim recoveryText As String
'    Dim cacheRecoveryText As String
'    Dim cacheRecoveryErrNumber As Long
'    Dim cacheRecoveryErrDescription As String
'    Dim outcomeText As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    If cacheStateCaptured Then
'        On Error Resume Next
'        Err.Clear
'        If saveAttempted And cachePublished = False Then
'            SharedForm_RemoveZoomBaselineForForm formName
'            cacheRecoveryText = "The zoom cache was invalidated because the saved form outcome is uncertain or unverified."
'        ElseIf saveAttempted = False And cacheMutationStarted And cachePublished = False Then
'            If baselineExisted Then
'                SharedForm_SetZoomLevelValue formName, priorZoomLevel
'                cacheRecoveryText = "The prior zoom level was restored."
'            Else
'                SharedForm_RemoveZoomBaselineForForm formName
'                cacheRecoveryText = "The newly captured zoom baseline was removed."
'            End If
'        End If
'        cacheRecoveryErrNumber = Err.Number
'        cacheRecoveryErrDescription = Err.Description
'        On Error GoTo 0
'    End If
'    recoveryText = SharedForm_RecoverAfterDesignChangeFailure(formName, hostFormName, designFormOpened, hostFormClosed)
'    If cacheRecoveryErrNumber <> 0 Then cacheRecoveryText = SharedText_AppendNonBlankLine(cacheRecoveryText, "Zoom-cache recovery error " & CStr(cacheRecoveryErrNumber) & ": " & cacheRecoveryErrDescription)
'    If saveAttempted And saveCompleted = False Then
'        outcomeText = "The scale save command failed and the persisted form outcome is uncertain. The zoom cache was invalidated; inspect the form before rerunning."
'    ElseIf saveCompleted And formStateVerified = False Then
'        outcomeText = "The scale save completed, but persisted geometry or font verification failed. The zoom cache was invalidated; inspect the saved form before rerunning."
'    ElseIf formStateVerified And cachePublished = False Then
'        outcomeText = "The canonical scaled form state was verified, but its canonical zoom checkpoint could not be published. Clear and recapture the cache before another scale."
'    ElseIf formStateVerified And cachePublished And formStateChanged Then
'        outcomeText = "The scaled form was saved and verified and its canonical zoom checkpoint was published, but host-form restoration failed. Do not rerun the scale solely because of this error."
'    ElseIf formStateVerified And cachePublished Then
'        outcomeText = "The form already matched the requested effective canonical zoom state and the canonical checkpoint was published, but host-form restoration failed. Do not rerun the scale solely because of this error."
'    Else
'        outcomeText = "The scale operation was not saved; recovery discarded any open unsaved design changes and restored or removed only cache state changed by this invocation."
'    End If
'    Err.Raise savedErrNumber, "SharedForm_ScaleControlFontsAndLayout", outcomeText & vbCrLf & "Prior zoom level: " & CStr(priorZoomLevel) & vbCrLf & "Canonical prior zoom level: " & CStr(canonicalPriorZoomLevel) & vbCrLf & "Requested zoom level: " & CStr(requestedZoomLevel) & vbCrLf & "Target zoom level: " & CStr(targetZoomLevel) & vbCrLf & "Form: " & formName & vbCrLf & "Host form: " & hostFormName & vbCrLf & "Stage: " & stageText & vbCrLf & "Control: " & currentControlName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & IIf(Len(cacheRecoveryText) > 0, vbCrLf & "Cache recovery: " & cacheRecoveryText, vbNullString) & IIf(Len(recoveryText) > 0, vbCrLf & "Form recovery: " & recoveryText, vbNullString)
'End Sub
'
'Public Sub SharedForm_IncreaseAndScaleControlFonts(ByVal frm As Access.Form)
'    'Intended use: Increases supported control fonts and proportionally scales the form layout.
'    SharedForm_ScaleControlFontsAndLayout frm, 1
'End Sub
'
'Public Sub SharedForm_DecreaseAndScaleControlFonts(ByVal frm As Access.Form)
'    'Intended use: Decreases supported control fonts and proportionally scales the form layout.
'    SharedForm_ScaleControlFontsAndLayout frm, -1
'End Sub
'
'Public Sub SharedForm_ClearScaleControlFontsBaseline(ByVal frm As Access.Form)
'    'Intended use: Clears all stored font, geometry, and level baseline metadata before a new scaling baseline is captured.
'    If frm Is Nothing Then Err.Raise vbObjectError + 8420, "SharedForm_ClearScaleControlFontsBaseline", "The Access.Form argument is Nothing."
'    SharedForm_RemoveZoomBaselineForForm frm.Name
'End Sub
'
'Public Sub SharedForm_MoveTopLeft(ByVal frm As Access.Form, Optional ByVal leftPos As Long = 0, Optional ByVal topPos As Long = 0, Optional ByVal formWidth As Long = -1, Optional ByVal FormHeight As Long = -1)
'    'Intended use: Best-effort optional UI helper that moves and optionally resizes a form window; callers must not use it as an authoritative success boundary.
'    On Error Resume Next
'    DoCmd.SelectObject acForm, frm.Name, False
'    DoCmd.Restore
'    If formWidth = -1 Then formWidth = frm.windowWidth
'    If FormHeight = -1 Then FormHeight = frm.windowHeight
'    DoCmd.MoveSize leftPos, topPos, formWidth, FormHeight
'    On Error GoTo 0
'End Sub
'
'Public Sub SharedForm_ResetAndScaleControlFonts(ByVal frm As Access.Form)
'    'Intended use: Restores authoritative absolute zoom level zero; fails without changing the form when no authoritative baseline exists.
'    SharedForm_SetControlFontsAndLayoutZoom frm, 0, True
'End Sub
'
'Public Function SharedForm_IsLoadedAsSubform(ByVal frm As Access.Form) As Boolean
'    'Returns True when a form has a Parent form, which means it is currently loaded as a subform.
'    'Use before calling Me.Parent methods from code that can run as a main form or subform.
'    On Error GoTo NotSubform
'    Dim parentName As String
'    parentName = frm.Parent.Name
'    SharedForm_IsLoadedAsSubform = True
'    Exit Function
'NotSubform:
'    Err.Clear
'End Function
'
'Public Sub SharedForm_CloseIfOpen(ByVal formName As String, Optional ByVal saveMode As AcCloseSave = acSaveNo)
'    'Best-effort optional cleanup helper that closes an Access form if open; callers requiring verified closure must use the strict design-change helpers.
'    'Consolidated old names: CloseProgressFormIfOpen.
'    On Error Resume Next
'    If SysCmd(acSysCmdGetObjectState, acForm, formName) <> 0 Then DoCmd.Close acForm, formName, saveMode
'    On Error GoTo 0
'End Sub
'
'Public Sub SharedForm_GoToRecordID(ByVal frm As Access.Form, ByVal recordID As Variant, Optional ByVal idFieldName As String = "ID")
'    'Moves an Access form to the first record matching idFieldName=recordID.
'    'Consolidated old names: GoToID.
'    Dim rs As DAO.Recordset
'    If frm Is Nothing Then Exit Sub
'    If IsNull(recordID) Then Exit Sub
'    If Not IsNumeric(recordID) Then Exit Sub
'    Set rs = frm.RecordsetClone
'    rs.FindFirst SharedDao_BracketName(idFieldName) & "=" & CLng(recordID)
'    If Not rs.NoMatch Then frm.Bookmark = rs.Bookmark
'    rs.Close
'    Set rs = Nothing
'End Sub
'
'Public Sub SharedForm_SelectComboText(ByVal comboControl As Access.control)
'    'Selects all visible text in a combo box control.
'    'Consolidated old names: ffwFile_Enter combo selection logic, ffwMenu_Enter combo selection logic.
'    On Error Resume Next
'    If comboControl Is Nothing Then Exit Sub
'    If comboControl.controlType <> acComboBox Then Exit Sub
'    comboControl.SelStart = 0
'    comboControl.SelLength = Len(Nz(comboControl.text, vbNullString))
'    On Error GoTo 0
'End Sub
'
'Public Sub SharedForm_SelectActiveComboText()
'    'Selects all visible text in the active control when the active control is a combo box.
'    'Consolidated old names: Form_Timer active combo selection logic.
'    On Error Resume Next
'    If Screen.ActiveControl.controlType = acComboBox Then SharedForm_SelectComboText Screen.ActiveControl
'    On Error GoTo 0
'End Sub
'
'' ================================================================================================================================================================================
'' NAVPANE
'' ================================================================================================================================================================================
'Public Sub SharedNavPane_AddQueryToCustomGroup(ByVal groupName As String, ByVal queryName As String)
'    'Adds a saved query to one unambiguous existing custom Navigation Pane group through a transactional, verified shortcut replacement.
'    On Error GoTo ErrHandler
'    Const ACCESS_QUERY_TYPE As Long = 5
'    Dim db As DAO.Database
'    Dim groupID As Long
'    Dim objectID As Long
'    Set db = currentDb
'    groupID = SharedNavPane_ExistingCustomGroupID(db, groupName)
'    If groupID = 0 Then Err.Raise vbObjectError + 7111, "SharedNavPane_AddQueryToCustomGroup", "Custom Navigation Pane group not found: " & groupName
'    objectID = SharedNavPane_ObjectID(db, queryName, ACCESS_QUERY_TYPE)
'    If objectID = 0 Then Err.Raise vbObjectError + 7112, "SharedNavPane_AddQueryToCustomGroup", "Query not found in MSysObjects: " & queryName
'    SharedNavPane_ReplaceGroupObjectShortcutTransactional db, groupID, objectID, queryName, ACCESS_QUERY_TYPE
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedNavPane_AddQueryToCustomGroup", "Could not add the query to the custom Navigation Pane group." & vbCrLf & "Group: " & groupName & vbCrLf & "Query: " & queryName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Sub
'
'Public Sub SharedNavPane_AddTableToCustomGroup(ByVal groupName As String, ByVal tableName As String)
'    'Adds a local table to one unambiguous existing custom Navigation Pane group through a transactional, verified shortcut replacement.
'    'Consolidated old names: local TV NSS IMPORT MODULE AddTableToCustomNavGroup.
'    On Error GoTo ErrHandler
'    Const ACCESS_LOCAL_TABLE_TYPE As Long = 1
'    Dim db As DAO.Database
'    Dim groupID As Long
'    Dim objectID As Long
'    Set db = currentDb
'    groupID = SharedNavPane_ExistingCustomGroupID(db, groupName)
'    If groupID = 0 Then Err.Raise vbObjectError + 7101, "SharedNavPane_AddTableToCustomGroup", "Custom Navigation Pane group not found: " & groupName
'    objectID = SharedNavPane_ObjectID(db, tableName, ACCESS_LOCAL_TABLE_TYPE)
'    If objectID = 0 Then Err.Raise vbObjectError + 7102, "SharedNavPane_AddTableToCustomGroup", "Table not found in MSysObjects: " & tableName
'    SharedNavPane_ReplaceGroupObjectShortcutTransactional db, groupID, objectID, tableName, ACCESS_LOCAL_TABLE_TYPE
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedNavPane_AddTableToCustomGroup", "Could not add the table to the custom Navigation Pane group." & vbCrLf & "Group: " & groupName & vbCrLf & "Table: " & tableName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Sub
'
'Public Function SharedNavPane_GroupID(ByVal db As DAO.Database, ByVal groupName As String) As Long
'    'Returns the Navigation Pane custom group ID for a group name.
'    'Consolidated old names: local TV NSS IMPORT MODULE NavGroupID.
'    Dim rs As DAO.Recordset
'    Dim sqlText As String
'    sqlText = "SELECT [Id] FROM [MSysNavPaneGroups] WHERE [Name]=" & SharedDao_SqlQuote(groupName) & " AND [Name] Not Like 'Unassigned*'"
'    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
'    If Not rs.EOF Then SharedNavPane_GroupID = CLng(rs.fields("Id").value)
'    rs.Close
'End Function
'
'Public Function SharedNavPane_ObjectID(ByVal db As DAO.Database, ByVal objectName As String, ByVal objectType As Long) As Long
'    'Returns the MSysObjects ID for an object name/type pair.
'    'Consolidated old names: local TV NSS IMPORT MODULE NavObjectID.
'    Dim rs As DAO.Recordset
'    Dim sqlText As String
'    sqlText = "SELECT [Id] FROM [MSysObjects] WHERE [Name]=" & SharedDao_SqlQuote(objectName) & " AND [Type]=" & CStr(objectType)
'    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
'    If Not rs.EOF Then SharedNavPane_ObjectID = CLng(rs.fields("Id").value)
'    rs.Close
'End Function
'
'Public Sub SharedNavPane_EnsureObjectID(ByVal db As DAO.Database, ByVal objectID As Long, ByVal objectName As String, ByVal objectType As Long)
'    'Ensures MSysNavPaneObjectIDs contains the current Navigation Pane object mapping inside a verified transaction.
'    'Consolidated old names: local TV NSS IMPORT MODULE EnsureNavPaneObjectID.
'    On Error GoTo ErrHandler
'    Dim wrk As DAO.workspace
'    Dim transactionStarted As Boolean
'    If db Is Nothing Then Err.Raise vbObjectError + 7640, "SharedNavPane_EnsureObjectID", "The DAO.Database argument is Nothing."
'    Set wrk = DBEngine.Workspaces(0)
'    wrk.BeginTrans
'    transactionStarted = True
'    SharedNavPane_EnsureObjectIDCore db, objectID, objectName, objectType
'    wrk.CommitTrans
'    transactionStarted = False
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim rollbackErrNumber As Long
'    Dim rollbackErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    Err.Clear
'    If transactionStarted And Not wrk Is Nothing Then wrk.Rollback
'    rollbackErrNumber = Err.Number
'    rollbackErrDescription = Err.Description
'    Set wrk = Nothing
'    On Error GoTo 0
'    If rollbackErrNumber <> 0 Then savedErrDescription = savedErrDescription & vbCrLf & "Rollback error " & CStr(rollbackErrNumber) & ": " & rollbackErrDescription
'    Err.Raise savedErrNumber, "SharedNavPane_EnsureObjectID", "Could not ensure the Navigation Pane object-ID mapping." & vbCrLf & "Object: " & objectName & vbCrLf & "Object ID: " & CStr(objectID) & vbCrLf & "Object type: " & CStr(objectType) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Sub
'
'Private Sub SharedNavPane_EnsureObjectIDCore(ByVal db As DAO.Database, ByVal objectID As Long, ByVal objectName As String, ByVal objectType As Long)
'    'Performs one object-ID mapping mutation for an owning transaction and verifies the exact persisted mapping.
'    On Error GoTo ErrHandler
'    Dim rs As DAO.Recordset
'    Dim sqlText As String
'    Dim verifySqlText As String
'    If db Is Nothing Then Err.Raise vbObjectError + 7641, "SharedNavPane_EnsureObjectIDCore", "The DAO.Database argument is Nothing."
'    If objectID = 0 Then Err.Raise vbObjectError + 7642, "SharedNavPane_EnsureObjectIDCore", "ObjectID must not be zero."
'    If Len(objectName) = 0 Then Err.Raise vbObjectError + 7643, "SharedNavPane_EnsureObjectIDCore", "Object name is blank."
'    sqlText = "DELETE FROM [MSysNavPaneObjectIDs] WHERE [Name]=" & SharedDao_SqlQuote(objectName) & " AND [Type]=" & CStr(objectType) & " AND [Id]<>" & CStr(objectID)
'    db.Execute sqlText, dbFailOnError
'    sqlText = "SELECT * FROM [MSysNavPaneObjectIDs] WHERE [Id]=" & CStr(objectID)
'    Set rs = db.OpenRecordset(sqlText, dbOpenDynaset)
'    If rs.EOF Then
'        rs.AddNew
'        rs.fields("Id").value = objectID
'        rs.fields("Name").value = objectName
'        rs.fields("Type").value = objectType
'        rs.Update
'    Else
'        rs.Edit
'        rs.fields("Name").value = objectName
'        rs.fields("Type").value = objectType
'        rs.Update
'    End If
'    rs.Close
'    Set rs = Nothing
'    verifySqlText = "SELECT Count(*) AS MatchCount FROM [MSysNavPaneObjectIDs] WHERE [Id]=" & CStr(objectID) & " AND [Name]=" & SharedDao_SqlQuote(objectName) & " AND [Type]=" & CStr(objectType)
'    Set rs = db.OpenRecordset(verifySqlText, dbOpenSnapshot)
'    If rs.EOF Or CLng(rs.fields("MatchCount").value) <> 1 Then Err.Raise vbObjectError + 7644, "SharedNavPane_EnsureObjectIDCore", "The exact Navigation Pane object-ID mapping could not be verified after the write."
'    rs.Close
'    Set rs = Nothing
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If Not rs Is Nothing Then rs.Close
'    Set rs = Nothing
'    On Error GoTo 0
'    Err.Raise savedErrNumber, "SharedNavPane_EnsureObjectIDCore", "Unable to write or verify the Navigation Pane object-ID mapping." & vbCrLf & "Object: " & objectName & vbCrLf & "Object ID: " & CStr(objectID) & vbCrLf & "Object type: " & CStr(objectType) & vbCrLf & "Mutation SQL: " & sqlText & vbCrLf & "Verification SQL: " & verifySqlText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Sub
'Private Sub SharedNavPane_ReplaceGroupObjectShortcutTransactional(ByVal db As DAO.Database, ByVal groupID As Long, ByVal objectID As Long, ByVal objectName As String, ByVal objectType As Long)
'    'Replaces one legacy Navigation Pane shortcut atomically and verifies the final exact membership before commit.
'    On Error GoTo ErrHandler
'    Dim wrk As DAO.workspace
'    Dim transactionStarted As Boolean
'    If db Is Nothing Then Err.Raise vbObjectError + 7645, "SharedNavPane_ReplaceGroupObjectShortcutTransactional", "The DAO.Database argument is Nothing."
'    Set wrk = DBEngine.Workspaces(0)
'    wrk.BeginTrans
'    transactionStarted = True
'    SharedNavPane_EnsureObjectIDCore db, objectID, objectName, objectType
'    SharedNavPane_DeleteGroupObjectShortcuts db, groupID, objectID, objectName
'    If Not SharedNavPane_InsertGroupObjectShortcutWriteOnlyStrict(db, groupID, objectID, objectName, objectType, True, True) Then Err.Raise vbObjectError + 7646, "SharedNavPane_ReplaceGroupObjectShortcutTransactional", "The Navigation Pane shortcut insert did not report success."
'    If Not SharedNavPane_GroupObjectShortcutExistsStrict(db, groupID, objectID) Then Err.Raise vbObjectError + 7647, "SharedNavPane_ReplaceGroupObjectShortcutTransactional", "The exact Navigation Pane shortcut membership could not be verified after insertion."
'    wrk.CommitTrans
'    transactionStarted = False
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim rollbackErrNumber As Long
'    Dim rollbackErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    Err.Clear
'    If transactionStarted And Not wrk Is Nothing Then wrk.Rollback
'    rollbackErrNumber = Err.Number
'    rollbackErrDescription = Err.Description
'    Set wrk = Nothing
'    On Error GoTo 0
'    If rollbackErrNumber <> 0 Then savedErrDescription = savedErrDescription & vbCrLf & "Rollback error " & CStr(rollbackErrNumber) & ": " & rollbackErrDescription
'    Err.Raise savedErrNumber, "SharedNavPane_ReplaceGroupObjectShortcutTransactional", "Could not transactionally replace the Navigation Pane shortcut." & vbCrLf & "Object: " & objectName & vbCrLf & "Group ID: " & CStr(groupID) & vbCrLf & "Object ID: " & CStr(objectID) & vbCrLf & "Object type: " & CStr(objectType) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Sub
'
'Public Function SharedNavPane_NextGroupPosition(ByVal db As DAO.Database, ByVal groupID As Long) As Long
'    'Returns the next Navigation Pane group position for a group.
'    'Consolidated old names: local TV NSS IMPORT MODULE NextNavGroupPosition.
'    Dim rs As DAO.Recordset
'    Dim sqlText As String
'    sqlText = "SELECT Max([Position]) AS MaxPosition FROM [MSysNavPaneGroupToObjects] WHERE [GroupID]=" & CStr(groupID)
'    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
'    If rs.EOF Or IsNull(rs.fields("MaxPosition").value) Then
'        SharedNavPane_NextGroupPosition = 0
'    Else
'        SharedNavPane_NextGroupPosition = CLng(rs.fields("MaxPosition").value) + 1
'    End If
'    rs.Close
'End Function
'
'Public Function SharedNavPane_GroupsFieldExists(ByVal db As DAO.Database, ByVal fieldName As String) As Boolean
'    'Intended use: Determines whether the Navigation Pane groups system table exposes a requested field.
'    SharedNavPane_GroupsFieldExists = SharedDao_DatabaseTableFieldExistsSafe(db, "MSysNavPaneGroups", fieldName)
'End Function
'
'Private Function SharedNavPane_GroupsFieldExistsStrict(ByVal db As DAO.Database, ByVal fieldName As String) As Boolean
'    'Returns confirmed field existence and raises when the Navigation Pane groups table cannot be inspected.
'    On Error GoTo ErrHandler
'    If db Is Nothing Then Err.Raise vbObjectError + 7901, "SharedNavPane_GroupsFieldExistsStrict", "The DAO.Database argument is Nothing."
'    fieldName = Trim$(fieldName)
'    If Len(fieldName) = 0 Then Err.Raise vbObjectError + 7902, "SharedNavPane_GroupsFieldExistsStrict", "The field name is blank."
'    If SharedDao_DatabaseTableExists(db, "MSysNavPaneGroups") = False Then Err.Raise vbObjectError + 7903, "SharedNavPane_GroupsFieldExistsStrict", "MSysNavPaneGroups is not available."
'    SharedNavPane_GroupsFieldExistsStrict = SharedDao_TableFieldExists(db, "MSysNavPaneGroups", fieldName)
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedNavPane_GroupsFieldExistsStrict", "Could not inspect the Navigation Pane groups schema." & vbCrLf & "Database: " & SharedDao_SafeDatabaseName(db) & vbCrLf & "Field: " & fieldName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'
'Public Function SharedNavPane_GroupToObjectsFieldExists(ByVal db As DAO.Database, ByVal fieldName As String) As Boolean
'    'Intended use: Determines whether the Navigation Pane group-to-objects system table exposes a requested field.
'    SharedNavPane_GroupToObjectsFieldExists = SharedDao_DatabaseTableFieldExistsSafe(db, "MSysNavPaneGroupToObjects", fieldName)
'End Function
'
'Private Function SharedNavPane_GroupToObjectsFieldExistsStrict(ByVal db As DAO.Database, ByVal fieldName As String) As Boolean
'    'Returns confirmed field existence and raises when the Navigation Pane system table cannot be inspected.
'    On Error GoTo ErrHandler
'    If db Is Nothing Then Err.Raise vbObjectError + 7904, "SharedNavPane_GroupToObjectsFieldExistsStrict", "The DAO.Database argument is Nothing."
'    fieldName = Trim$(fieldName)
'    If Len(fieldName) = 0 Then Err.Raise vbObjectError + 7905, "SharedNavPane_GroupToObjectsFieldExistsStrict", "The field name is blank."
'    If SharedDao_DatabaseTableExists(db, "MSysNavPaneGroupToObjects") = False Then Err.Raise vbObjectError + 7906, "SharedNavPane_GroupToObjectsFieldExistsStrict", "MSysNavPaneGroupToObjects is not available."
'    SharedNavPane_GroupToObjectsFieldExistsStrict = SharedDao_TableFieldExists(db, "MSysNavPaneGroupToObjects", fieldName)
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedNavPane_GroupToObjectsFieldExistsStrict", "Could not inspect the Navigation Pane shortcut schema." & vbCrLf & "Field: " & fieldName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'
'Public Function SharedNavPane_GroupsCategoryFieldName(ByVal db As DAO.Database) As String
'    'Returns the supported category-ID field name and raises when the Navigation Pane groups schema cannot be inspected.
'    On Error GoTo ErrHandler
'    If SharedNavPane_GroupsFieldExistsStrict(db, "GroupCategoryID") Then
'        SharedNavPane_GroupsCategoryFieldName = "GroupCategoryID"
'    ElseIf SharedNavPane_GroupsFieldExistsStrict(db, "CategoryID") Then
'        SharedNavPane_GroupsCategoryFieldName = "CategoryID"
'    End If
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedNavPane_GroupsCategoryFieldName", "Could not resolve the Navigation Pane groups category field." & vbCrLf & "Database: " & SharedDao_SafeDatabaseName(db) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'
'Public Function SharedNavPane_CustomCategoryID(ByVal db As DAO.Database) As Long
'    'Returns the unique custom Navigation Pane category ID, zero when genuinely absent, and raises on ambiguity or inspection failure.
'    On Error GoTo ErrHandler
'    Dim rs As DAO.Recordset
'    Dim sqlText As String
'    Dim matchCount As Long
'    If db Is Nothing Then Err.Raise vbObjectError + 7803, "SharedNavPane_CustomCategoryID", "The DAO.Database argument is Nothing."
'    sqlText = "SELECT [ID], [Name] FROM [MSysNavPaneGroupCategories] WHERE UCase(Trim([Name])) IN ('CUSTOM','CUSTOM GROUPS','CUSTOM CATEGORY') ORDER BY [ID];"
'    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
'    Do While rs.EOF = False
'        matchCount = matchCount + 1
'        If matchCount = 1 Then SharedNavPane_CustomCategoryID = CLng(rs.fields("ID").value)
'        rs.MoveNext
'    Loop
'    If matchCount > 1 Then Err.Raise vbObjectError + 7804, "SharedNavPane_CustomCategoryID", "More than one recognized custom Navigation Pane category was found."
'CleanExit:
'    On Error Resume Next
'    If Not rs Is Nothing Then rs.Close
'    Set rs = Nothing
'    On Error GoTo 0
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If Not rs Is Nothing Then rs.Close
'    Set rs = Nothing
'    On Error GoTo 0
'    Err.Raise savedErrNumber, "SharedNavPane_CustomCategoryID", "Could not resolve the custom Navigation Pane category." & vbCrLf & "Database: " & SharedDao_SafeDatabaseName(db) & vbCrLf & "Matches found: " & CStr(matchCount) & vbCrLf & "SQL: " & sqlText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'
'Public Function SharedNavPane_ExistingAnyGroupID(ByVal db As DAO.Database, ByVal groupName As String) As Long
'    'Returns an existing Navigation Pane group ID from any category, zero when genuinely absent, and raises on inspection failure.
'    On Error GoTo ErrHandler
'    Dim rs As DAO.Recordset
'    Dim sqlText As String
'    Dim normalizedGroupName As String
'    If db Is Nothing Then Err.Raise vbObjectError + 7805, "SharedNavPane_ExistingAnyGroupID", "The DAO.Database argument is Nothing."
'    If Len(groupName) = 0 Then Exit Function
'    normalizedGroupName = UCase$(groupName)
'    sqlText = "SELECT TOP 1 [ID] FROM [MSysNavPaneGroups] WHERE UCase(Nz([Name],''))=" & SharedDao_SqlQuote(normalizedGroupName) & " ORDER BY [ID];"
'    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
'    If Not rs.EOF Then SharedNavPane_ExistingAnyGroupID = CLng(rs.fields("ID").value)
'CleanExit:
'    On Error Resume Next
'    If Not rs Is Nothing Then rs.Close
'    Set rs = Nothing
'    On Error GoTo 0
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If Not rs Is Nothing Then rs.Close
'    Set rs = Nothing
'    On Error GoTo 0
'    Err.Raise savedErrNumber, "SharedNavPane_ExistingAnyGroupID", "Could not inspect Navigation Pane groups." & vbCrLf & "Database: " & SharedDao_SafeDatabaseName(db) & vbCrLf & "Group: " & groupName & vbCrLf & "SQL: " & sqlText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'Public Function SharedNavPane_ExistingCustomGroupID(ByVal db As DAO.Database, ByVal groupName As String) As Long
'    'Returns the unique matching custom Navigation Pane group ID, zero when genuinely absent, and raises on ambiguity or inspection failure.
'    On Error GoTo ErrHandler
'    Dim rs As DAO.Recordset
'    Dim categoryID As Long
'    Dim sqlText As String
'    Dim normalizedGroupName As String
'    Dim categoryFieldName As String
'    Dim matchCount As Long
'    If db Is Nothing Then Err.Raise vbObjectError + 7806, "SharedNavPane_ExistingCustomGroupID", "The DAO.Database argument is Nothing."
'    If Len(groupName) = 0 Then Exit Function
'    normalizedGroupName = UCase$(groupName)
'    categoryID = SharedNavPane_CustomCategoryID(db)
'    If categoryID <= 0 Then Exit Function
'    categoryFieldName = SharedNavPane_GroupsCategoryFieldName(db)
'    If Len(categoryFieldName) = 0 Then Err.Raise vbObjectError + 7807, "SharedNavPane_ExistingCustomGroupID", "No supported category-ID field exists in MSysNavPaneGroups."
'    sqlText = "SELECT [ID] FROM [MSysNavPaneGroups] WHERE [" & categoryFieldName & "]=" & CStr(categoryID)
'    sqlText = sqlText & " AND UCase(Nz([Name],''))=" & SharedDao_SqlQuote(normalizedGroupName) & " ORDER BY [ID];"
'    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
'    Do While rs.EOF = False
'        matchCount = matchCount + 1
'        If matchCount = 1 Then SharedNavPane_ExistingCustomGroupID = CLng(rs.fields("ID").value)
'        rs.MoveNext
'    Loop
'    If matchCount > 1 Then Err.Raise vbObjectError + 7808, "SharedNavPane_ExistingCustomGroupID", "More than one matching custom Navigation Pane group was found."
'CleanExit:
'    On Error Resume Next
'    If Not rs Is Nothing Then rs.Close
'    Set rs = Nothing
'    On Error GoTo 0
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If Not rs Is Nothing Then rs.Close
'    Set rs = Nothing
'    On Error GoTo 0
'    Err.Raise savedErrNumber, "SharedNavPane_ExistingCustomGroupID", "Could not resolve the custom Navigation Pane group." & vbCrLf & "Database: " & SharedDao_SafeDatabaseName(db) & vbCrLf & "Group: " & groupName & vbCrLf & "Category ID: " & CStr(categoryID) & vbCrLf & "Category field: " & categoryFieldName & vbCrLf & "Matches found: " & CStr(matchCount) & vbCrLf & "SQL: " & sqlText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'Public Function SharedNavPane_MSysObjectID(ByVal db As DAO.Database, ByVal objectName As String, ByVal objectType As AcObjectType) As Long
'    'Returns the exact-type MSysObjects ID, zero when genuinely absent, and raises on inspection failure.
'    SharedNavPane_MSysObjectID = SharedNavPane_MSysObjectIDStrict(db, objectName, objectType)
'End Function
'
'Private Function SharedNavPane_MSysObjectIDStrict(ByVal db As DAO.Database, ByVal objectName As String, ByVal objectType As AcObjectType) As Long
'    'Resolves only MSysObjects types belonging to the requested Access object type and preserves technical inspection failures.
'    On Error GoTo ErrHandler
'    Dim rs As DAO.Recordset
'    Dim sqlText As String
'    Dim typeCriteria As String
'    Dim normalizedObjectName As String
'    If db Is Nothing Then Err.Raise vbObjectError + 7730, "SharedNavPane_MSysObjectIDStrict", "The DAO.Database argument is Nothing."
'    typeCriteria = SharedNavPane_MSysObjectTypeCriteria(objectType)
'    If Len(typeCriteria) = 0 Then Exit Function
'    If Len(objectName) = 0 Then Exit Function
'    normalizedObjectName = UCase$(objectName)
'    sqlText = "SELECT TOP 1 [ID] FROM [MSysObjects] WHERE UCase([Name])=" & SharedDao_SqlQuote(normalizedObjectName) & " AND " & typeCriteria & " ORDER BY [ID];"
'    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
'    If Not rs.EOF Then SharedNavPane_MSysObjectIDStrict = CLng(rs.fields("ID").value)
'CleanExit:
'    On Error Resume Next
'    If Not rs Is Nothing Then rs.Close
'    Set rs = Nothing
'    On Error GoTo 0
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If Not rs Is Nothing Then rs.Close
'    Set rs = Nothing
'    On Error GoTo 0
'    Err.Raise savedErrNumber, "SharedNavPane_MSysObjectIDStrict", "Could not resolve the requested-type Navigation Pane object ID." & vbCrLf & "Database: " & SharedDao_SafeDatabaseName(db) & vbCrLf & "Object: " & objectName & vbCrLf & "Access object type: " & CStr(objectType) & vbCrLf & "MSysObjects criteria: " & typeCriteria & vbCrLf & "SQL: " & sqlText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'Private Function SharedNavPane_MSysObjectTypeCriteria(ByVal objectType As AcObjectType) As String
'    'Returns the MSysObjects type predicate for one Access object class without admitting same-name objects from another class.
'    Select Case objectType
'        Case acTable
'            SharedNavPane_MSysObjectTypeCriteria = "[Type] IN (1,4,6)"
'        Case Else
'            If SharedDao_MSysObjectTypeForAccessObjectType(objectType) <> 0 Then SharedNavPane_MSysObjectTypeCriteria = "[Type]=" & CStr(SharedDao_MSysObjectTypeForAccessObjectType(objectType))
'    End Select
'End Function
'
'Public Sub SharedNavPane_EnsureObjectIDRow(ByVal objectName As String, Optional ByVal objectType As AcObjectType = acForm, Optional ByVal tempPrefix As String = "Shared_NavPane_")
'    'Refreshes the Access object-ID row through a unique temporary export and deletes only the file created by this call.
'    On Error GoTo ErrHandler
'    Dim fso As Object
'    Dim tempFolderPath As String
'    Dim tempFilePath As String
'    Dim tempFileCreated As Boolean
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim cleanupErrNumber As Long
'    Dim cleanupErrDescription As String
'    If Len(objectName) = 0 Then Err.Raise vbObjectError + 7726, "SharedNavPane_EnsureObjectIDRow", "The Access object name is blank."
'    Select Case objectType
'        Case acMacro, acQuery, acForm, acReport
'        Case Else
'            Exit Sub
'    End Select
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    tempFolderPath = Environ$("Temp")
'    If Len(Trim$(tempFolderPath)) = 0 Then Err.Raise vbObjectError + 7727, "SharedNavPane_EnsureObjectIDRow", "The Windows temporary-folder path is unavailable."
'    If Not fso.FolderExists(tempFolderPath) Then Err.Raise vbObjectError + 7728, "SharedNavPane_EnsureObjectIDRow", "The Windows temporary folder does not exist: " & tempFolderPath
'    tempFilePath = SharedPath_UniqueFilePath(SharedPath_Combine(tempFolderPath, tempPrefix & SharedName_SafeFileName(objectName) & ".txt"))
'    Application.SaveAsText objectType, objectName, tempFilePath
'    tempFileCreated = fso.fileExists(tempFilePath)
'    If Not tempFileCreated Then Err.Raise vbObjectError + 7729, "SharedNavPane_EnsureObjectIDRow", "The temporary object export was not created."
'    SharedFile_DeleteIfExists tempFilePath
'    tempFileCreated = False
'    Exit Sub
'ErrHandler:
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If tempFileCreated And Not fso Is Nothing Then
'        If fso.fileExists(tempFilePath) Then fso.DeleteFile tempFilePath, False
'    End If
'    cleanupErrNumber = Err.Number
'    cleanupErrDescription = Err.Description
'    On Error GoTo 0
'    Err.Raise savedErrNumber, "SharedNavPane_EnsureObjectIDRow", "Could not refresh the Navigation Pane object-ID row." & vbCrLf & "Object: " & objectName & vbCrLf & "Object type: " & CStr(objectType) & vbCrLf & "Temporary file: " & tempFilePath & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & IIf(cleanupErrNumber <> 0, vbCrLf & "Temporary-file cleanup error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription, vbNullString)
'End Sub
'Public Function SharedNavPane_GroupObjectShortcutExists(ByVal db As DAO.Database, ByVal groupID As Long, ByVal objectID As Long) As Boolean
'    'Intended use: Provides a nonraising optional check for an existing Navigation Pane group shortcut; authoritative mutation and rollback paths must use SharedNavPane_GroupObjectShortcutExistsStrict.
'    On Error GoTo ErrHandler
'    Dim rs As DAO.Recordset
'    Dim sqlText As String
'    If groupID <= 0 Or objectID = 0 Then Exit Function
'    sqlText = "SELECT Count(*) AS Cnt FROM [MSysNavPaneGroupToObjects] WHERE [GroupID]=" & CStr(groupID)
'    sqlText = sqlText & " AND [ObjectID]=" & CStr(objectID) & ";"
'    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
'    If Not rs.EOF Then SharedNavPane_GroupObjectShortcutExists = (CLng(rs!Cnt) > 0)
'CleanExit:
'    On Error Resume Next
'    If Not rs Is Nothing Then rs.Close
'    Exit Function
'ErrHandler:
'    SharedNavPane_GroupObjectShortcutExists = False
'    Resume CleanExit
'End Function
'
'Public Function SharedNavPane_NextGroupObjectPosition(ByVal db As DAO.Database, ByVal groupID As Long) As Long
'    'Returns the next available display position, zero when Position is genuinely unsupported, and raises on inspection failure.
'    On Error GoTo ErrHandler
'    Dim rs As DAO.Recordset
'    Dim sqlText As String
'    If db Is Nothing Then Err.Raise vbObjectError + 7809, "SharedNavPane_NextGroupObjectPosition", "The DAO.Database argument is Nothing."
'    If groupID <= 0 Then Err.Raise vbObjectError + 7810, "SharedNavPane_NextGroupObjectPosition", "GroupID must be greater than zero."
'    If SharedNavPane_GroupToObjectsFieldExistsStrict(db, "Position") = False Then Exit Function
'    sqlText = "SELECT Max([Position]) AS MaxPosition FROM [MSysNavPaneGroupToObjects] WHERE [GroupID]=" & CStr(groupID) & ";"
'    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
'    If rs.EOF Or IsNull(rs.fields("MaxPosition").value) Then
'        SharedNavPane_NextGroupObjectPosition = 1
'    Else
'        SharedNavPane_NextGroupObjectPosition = CLng(rs.fields("MaxPosition").value) + 1
'    End If
'CleanExit:
'    On Error Resume Next
'    If Not rs Is Nothing Then rs.Close
'    Set rs = Nothing
'    On Error GoTo 0
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If Not rs Is Nothing Then rs.Close
'    Set rs = Nothing
'    On Error GoTo 0
'    Err.Raise savedErrNumber, "SharedNavPane_NextGroupObjectPosition", "Could not resolve the next Navigation Pane shortcut position." & vbCrLf & "Database: " & SharedDao_SafeDatabaseName(db) & vbCrLf & "GroupID: " & CStr(groupID) & vbCrLf & "SQL: " & sqlText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'
'Public Function SharedNavPane_GroupObjectShortcutExistsStrict(ByVal db As DAO.Database, ByVal groupID As Long, ByVal objectID As Long) As Boolean
'    'Returns whether an exact Navigation Pane membership exists and raises instead of converting DAO failures to False.
'    On Error GoTo ErrHandler
'    Dim rs As DAO.Recordset
'    Dim sqlText As String
'    If db Is Nothing Then Err.Raise vbObjectError + 7907, "SharedNavPane_GroupObjectShortcutExistsStrict", "The DAO.Database argument is Nothing."
'    If groupID <= 0 Then Err.Raise vbObjectError + 7908, "SharedNavPane_GroupObjectShortcutExistsStrict", "GroupID must be greater than zero."
'    If objectID = 0 Then Err.Raise vbObjectError + 7909, "SharedNavPane_GroupObjectShortcutExistsStrict", "ObjectID must not be zero."
'    sqlText = "SELECT Count(*) AS Cnt FROM [MSysNavPaneGroupToObjects] WHERE [GroupID]=" & CStr(groupID) & " AND [ObjectID]=" & CStr(objectID) & ";"
'    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
'    If rs.EOF = False Then SharedNavPane_GroupObjectShortcutExistsStrict = (CLng(rs.fields("Cnt").value) > 0)
'    rs.Close
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If Not rs Is Nothing Then rs.Close
'    On Error GoTo 0
'    Err.Raise savedErrNumber, "SharedNavPane_GroupObjectShortcutExistsStrict", "Could not query an exact Navigation Pane membership." & vbCrLf & "GroupID: " & CStr(groupID) & vbCrLf & "ObjectID: " & CStr(objectID) & vbCrLf & "SQL: " & sqlText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
'End Function
'
'Private Function SharedNavPane_NextGroupObjectPositionStrict(ByVal db As DAO.Database, ByVal groupID As Long) As Long
'    'Returns the next group shortcut position and raises on schema or DAO inspection failures.
'    On Error GoTo ErrHandler
'    Dim rs As DAO.Recordset
'    Dim sqlText As String
'    If SharedNavPane_GroupToObjectsFieldExistsStrict(db, "Position") = False Then Exit Function
'    sqlText = "SELECT Max([Position]) AS MaxPosition FROM [MSysNavPaneGroupToObjects] WHERE [GroupID]=" & CStr(groupID) & ";"
'    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
'    If rs.EOF Then
'        SharedNavPane_NextGroupObjectPositionStrict = 1
'    Else
'        SharedNavPane_NextGroupObjectPositionStrict = SharedValue_NzLong(rs.fields("MaxPosition").value, 0) + 1
'    End If
'    rs.Close
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If Not rs Is Nothing Then rs.Close
'    On Error GoTo 0
'    Err.Raise savedErrNumber, "SharedNavPane_NextGroupObjectPositionStrict", "Could not resolve the next Navigation Pane shortcut position." & vbCrLf & "GroupID: " & CStr(groupID) & vbCrLf & "SQL: " & sqlText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'
'Private Sub SharedNavPane_InsertGroupObjectShortcutStrictCore(ByVal db As DAO.Database, ByVal groupID As Long, ByVal objectID As Long, ByVal objectName As String, ByVal typeValue As Long, ByVal includeTypeAndFlags As Boolean, ByVal includePosition As Boolean, ByRef useName As Boolean, ByRef useType As Boolean, ByRef useFlags As Boolean, ByRef usePosition As Boolean, ByRef expectedPosition As Long, ByRef sqlText As String)
'    'Builds and executes one schema-confirmed shortcut insert and returns the exact fields and position written.
'    Dim fieldList As String
'    Dim valueList As String
'    fieldList = "[GroupID], [ObjectID]"
'    valueList = CStr(groupID) & ", " & CStr(objectID)
'    useName = SharedNavPane_GroupToObjectsFieldExistsStrict(db, "Name")
'    If includeTypeAndFlags Then
'        useType = SharedNavPane_GroupToObjectsFieldExistsStrict(db, "Type")
'        useFlags = SharedNavPane_GroupToObjectsFieldExistsStrict(db, "Flags")
'    End If
'    If includePosition Then usePosition = SharedNavPane_GroupToObjectsFieldExistsStrict(db, "Position")
'    If useName Then
'        fieldList = fieldList & ", [Name]"
'        valueList = valueList & ", " & SharedDao_SqlQuote(objectName)
'    End If
'    If useType Then
'        fieldList = fieldList & ", [Type]"
'        valueList = valueList & ", " & CStr(typeValue)
'    End If
'    If useFlags Then
'        fieldList = fieldList & ", [Flags]"
'        valueList = valueList & ", 0"
'    End If
'    If usePosition Then
'        expectedPosition = SharedNavPane_NextGroupObjectPositionStrict(db, groupID)
'        fieldList = fieldList & ", [Position]"
'        valueList = valueList & ", " & CStr(expectedPosition)
'    End If
'    sqlText = "INSERT INTO [MSysNavPaneGroupToObjects] (" & fieldList & ") VALUES (" & valueList & ");"
'    db.Execute sqlText, dbFailOnError
'    If db.RecordsAffected <> 1 Then Err.Raise vbObjectError + 7910, "SharedNavPane_InsertGroupObjectShortcutStrictCore", "The shortcut insert did not affect exactly one row." & vbCrLf & "Affected rows: " & CStr(db.RecordsAffected)
'End Sub
'
'Private Function SharedNavPane_GroupObjectShortcutValuesMatchStrict(ByVal db As DAO.Database, ByVal groupID As Long, ByVal objectID As Long, ByVal objectName As String, ByVal typeValue As Long, ByVal useName As Boolean, ByVal useType As Boolean, ByVal useFlags As Boolean, ByVal usePosition As Boolean, ByVal expectedPosition As Long) As Boolean
'    'Verifies the unique inserted shortcut row and every optional field included in the write.
'    On Error GoTo ErrHandler
'    Dim rs As DAO.Recordset
'    Dim sqlText As String
'    sqlText = "SELECT * FROM [MSysNavPaneGroupToObjects] WHERE [GroupID]=" & CStr(groupID) & " AND [ObjectID]=" & CStr(objectID) & ";"
'    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
'    If rs.EOF Then GoTo CleanExit
'    If useName Then
'        If StrComp(Nz(rs.fields("Name").value, vbNullString), objectName, vbBinaryCompare) <> 0 Then GoTo CleanExit
'    End If
'    If useType Then
'        If SharedValue_NzLong(rs.fields("Type").value, 0) <> typeValue Then GoTo CleanExit
'    End If
'    If useFlags Then
'        If SharedValue_NzLong(rs.fields("Flags").value, -1) <> 0 Then GoTo CleanExit
'    End If
'    If usePosition Then
'        If IsNull(rs.fields("Position").value) Then GoTo CleanExit
'        If CLng(rs.fields("Position").value) <> expectedPosition Then GoTo CleanExit
'    End If
'    rs.MoveNext
'    If Not rs.EOF Then Err.Raise vbObjectError + 7911, "SharedNavPane_GroupObjectShortcutValuesMatchStrict", "More than one shortcut row matched the inserted GroupID and ObjectID."
'    SharedNavPane_GroupObjectShortcutValuesMatchStrict = True
'CleanExit:
'    On Error Resume Next
'    If Not rs Is Nothing Then rs.Close
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If Not rs Is Nothing Then rs.Close
'    On Error GoTo 0
'    Err.Raise savedErrNumber, "SharedNavPane_GroupObjectShortcutValuesMatchStrict", "Could not verify the inserted Navigation Pane shortcut values." & vbCrLf & "GroupID: " & CStr(groupID) & vbCrLf & "ObjectID: " & CStr(objectID) & vbCrLf & "SQL: " & sqlText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'
'Public Function SharedNavPane_InsertGroupObjectShortcutWriteOnlyStrict(ByVal db As DAO.Database, ByVal groupID As Long, ByVal objectID As Long, ByVal objectName As String, ByVal typeValue As Long, ByVal includeTypeAndFlags As Boolean, ByVal includePosition As Boolean) As Boolean
'    'Transactionally inserts one Navigation Pane shortcut and verifies every field included in the write.
'    On Error GoTo ErrHandler
'    Dim wrk As DAO.workspace
'    Dim useName As Boolean
'    Dim useType As Boolean
'    Dim useFlags As Boolean
'    Dim usePosition As Boolean
'    Dim expectedPosition As Long
'    Dim sqlText As String
'    Dim transactionStarted As Boolean
'    If db Is Nothing Then Err.Raise vbObjectError + 7912, "SharedNavPane_InsertGroupObjectShortcutWriteOnlyStrict", "The DAO.Database argument is Nothing."
'    If groupID <= 0 Then Err.Raise vbObjectError + 7913, "SharedNavPane_InsertGroupObjectShortcutWriteOnlyStrict", "GroupID must be greater than zero."
'    If objectID = 0 Then Err.Raise vbObjectError + 7914, "SharedNavPane_InsertGroupObjectShortcutWriteOnlyStrict", "ObjectID must not be zero."
'    Set wrk = DBEngine.Workspaces(0)
'    wrk.BeginTrans
'    transactionStarted = True
'    SharedNavPane_InsertGroupObjectShortcutStrictCore db, groupID, objectID, objectName, typeValue, includeTypeAndFlags, includePosition, useName, useType, useFlags, usePosition, expectedPosition, sqlText
'    If SharedNavPane_GroupObjectShortcutValuesMatchStrict(db, groupID, objectID, objectName, typeValue, useName, useType, useFlags, usePosition, expectedPosition) = False Then Err.Raise vbObjectError + 7915, "SharedNavPane_InsertGroupObjectShortcutWriteOnlyStrict", "The inserted shortcut did not match every field included in the write."
'    wrk.CommitTrans
'    transactionStarted = False
'    SharedNavPane_InsertGroupObjectShortcutWriteOnlyStrict = True
'    Set wrk = Nothing
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim rollbackErrNumber As Long
'    Dim rollbackErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    Err.Clear
'    If transactionStarted And Not wrk Is Nothing Then wrk.Rollback
'    rollbackErrNumber = Err.Number
'    rollbackErrDescription = Err.Description
'    Set wrk = Nothing
'    On Error GoTo 0
'    If rollbackErrNumber <> 0 Then savedErrDescription = savedErrDescription & vbCrLf & "Rollback error " & CStr(rollbackErrNumber) & ": " & rollbackErrDescription
'    Err.Raise savedErrNumber, "SharedNavPane_InsertGroupObjectShortcutWriteOnlyStrict", "Could not insert and verify the Navigation Pane shortcut row." & vbCrLf & "Object: " & objectName & vbCrLf & "GroupID: " & CStr(groupID) & vbCrLf & "ObjectID: " & CStr(objectID) & vbCrLf & "SQL: " & sqlText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'
'Public Function SharedNavPane_DeleteGroupObjectShortcutsStrict(ByVal db As DAO.Database, ByVal groupID As Long, ByVal objectID As Long, ByVal objectName As String) As Long
'    'Transactionally deletes matching shortcut rows, verifies removal, and raises on schema or DAO failures.
'    On Error GoTo ErrHandler
'    Dim wrk As DAO.workspace
'    Dim rs As DAO.Recordset
'    Dim sqlText As String
'    Dim verifySqlText As String
'    Dim matchCriteria As String
'    Dim cleanedObjectName As String
'    Dim affectedRows As Long
'    Dim transactionStarted As Boolean
'    If db Is Nothing Then Err.Raise vbObjectError + 7916, "SharedNavPane_DeleteGroupObjectShortcutsStrict", "The DAO.Database argument is Nothing."
'    cleanedObjectName = objectName
'    If groupID <= 0 Then Err.Raise vbObjectError + 7917, "SharedNavPane_DeleteGroupObjectShortcutsStrict", "GroupID must be greater than zero."
'    If objectID = 0 And Len(cleanedObjectName) = 0 Then Err.Raise vbObjectError + 7918, "SharedNavPane_DeleteGroupObjectShortcutsStrict", "ObjectID or objectName is required."
'    If objectID <> 0 Then
'        matchCriteria = "[ObjectID]=" & CStr(objectID)
'    ElseIf Len(cleanedObjectName) > 0 Then
'        If SharedNavPane_GroupToObjectsFieldExistsStrict(db, "Name") Then matchCriteria = "UCase(Nz([Name],''))=" & SharedDao_SqlQuote(UCase$(cleanedObjectName))
'    End If
'    If Len(matchCriteria) = 0 Then Err.Raise vbObjectError + 7919, "SharedNavPane_DeleteGroupObjectShortcutsStrict", "No valid shortcut deletion criteria could be built."
'    sqlText = "DELETE FROM [MSysNavPaneGroupToObjects] WHERE [GroupID]=" & CStr(groupID) & " AND (" & matchCriteria & ");"
'    verifySqlText = "SELECT Count(*) AS RemainingCount FROM [MSysNavPaneGroupToObjects] WHERE [GroupID]=" & CStr(groupID) & " AND (" & matchCriteria & ");"
'    Set wrk = DBEngine.Workspaces(0)
'    wrk.BeginTrans
'    transactionStarted = True
'    db.Execute sqlText, dbFailOnError
'    affectedRows = db.RecordsAffected
'    Set rs = db.OpenRecordset(verifySqlText, dbOpenSnapshot)
'    If CLng(rs.fields("RemainingCount").value) <> 0 Then Err.Raise vbObjectError + 7920, "SharedNavPane_DeleteGroupObjectShortcutsStrict", "Matching Navigation Pane shortcut rows remained after deletion."
'    rs.Close
'    Set rs = Nothing
'    wrk.CommitTrans
'    transactionStarted = False
'    SharedNavPane_DeleteGroupObjectShortcutsStrict = affectedRows
'    Set wrk = Nothing
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim rollbackErrNumber As Long
'    Dim rollbackErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If Not rs Is Nothing Then rs.Close
'    Err.Clear
'    If transactionStarted And Not wrk Is Nothing Then wrk.Rollback
'    rollbackErrNumber = Err.Number
'    rollbackErrDescription = Err.Description
'    Set rs = Nothing
'    Set wrk = Nothing
'    On Error GoTo 0
'    If rollbackErrNumber <> 0 Then savedErrDescription = savedErrDescription & vbCrLf & "Rollback error " & CStr(rollbackErrNumber) & ": " & rollbackErrDescription
'    Err.Raise savedErrNumber, "SharedNavPane_DeleteGroupObjectShortcutsStrict", "Could not delete and verify the Navigation Pane shortcut row." & vbCrLf & "Object: " & objectName & vbCrLf & "GroupID: " & CStr(groupID) & vbCrLf & "ObjectID: " & CStr(objectID) & vbCrLf & "Delete SQL: " & sqlText & vbCrLf & "Verification SQL: " & verifySqlText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'Public Function SharedNavPane_InsertGroupObjectShortcut(ByVal db As DAO.Database, ByVal groupID As Long, ByVal objectID As Long, ByVal objectName As String, ByVal typeValue As Long, ByVal includeTypeAndFlags As Boolean, ByVal includePosition As Boolean, ByRef detailText As String) As Boolean
'    'Intended use: Nonraising adapter over the transactional strict Navigation Pane shortcut insert.
'    On Error GoTo ErrHandler
'    detailText = vbNullString
'    SharedNavPane_InsertGroupObjectShortcut = SharedNavPane_InsertGroupObjectShortcutWriteOnlyStrict(db, groupID, objectID, objectName, typeValue, includeTypeAndFlags, includePosition)
'    Exit Function
'ErrHandler:
'    detailText = "Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
'    SharedNavPane_InsertGroupObjectShortcut = False
'End Function
'
'Public Sub SharedNavPane_DeleteGroupObjectShortcuts(ByVal db As DAO.Database, ByVal groupID As Long, ByVal objectID As Long, ByVal objectName As String)
'    'Intended use: Transactionally deletes shortcuts matching the object ID or confirmed Name field and verifies that none remain.
'    On Error GoTo ErrHandler
'    Dim wrk As DAO.workspace
'    Dim rs As DAO.Recordset
'    Dim sqlText As String
'    Dim verifySqlText As String
'    Dim matchCriteria As String
'    Dim cleanedObjectName As String
'    Dim transactionStarted As Boolean
'    cleanedObjectName = objectName
'    If db Is Nothing Then Err.Raise vbObjectError + 7632, "SharedNavPane_DeleteGroupObjectShortcuts", "The DAO.Database argument is Nothing."
'    If groupID <= 0 Then Err.Raise vbObjectError + 7633, "SharedNavPane_DeleteGroupObjectShortcuts", "The Navigation Pane group ID must be positive."
'    If objectID = 0 And Len(cleanedObjectName) = 0 Then Err.Raise vbObjectError + 7634, "SharedNavPane_DeleteGroupObjectShortcuts", "Both the object ID and object name are blank or zero."
'    If objectID <> 0 Then matchCriteria = "[ObjectID]=" & SharedDao_InvariantNumericLiteral(objectID)
'    If Len(cleanedObjectName) > 0 Then
'        If SharedNavPane_GroupToObjectsFieldExistsStrict(db, "Name") Then
'            If Len(matchCriteria) > 0 Then matchCriteria = matchCriteria & " OR "
'            matchCriteria = matchCriteria & "UCase(Nz([Name],''))=" & SharedDao_SqlQuote(UCase$(cleanedObjectName))
'        End If
'    End If
'    If Len(matchCriteria) = 0 Then Err.Raise vbObjectError + 7635, "SharedNavPane_DeleteGroupObjectShortcuts", "No usable shortcut-match criteria could be built."
'    sqlText = "DELETE FROM [MSysNavPaneGroupToObjects] WHERE [GroupID]=" & SharedDao_InvariantNumericLiteral(groupID) & " AND (" & matchCriteria & ");"
'    verifySqlText = "SELECT Count(*) AS RemainingCount FROM [MSysNavPaneGroupToObjects] WHERE [GroupID]=" & SharedDao_InvariantNumericLiteral(groupID) & " AND (" & matchCriteria & ");"
'    Set wrk = DBEngine.Workspaces(0)
'    wrk.BeginTrans
'    transactionStarted = True
'    db.Execute sqlText, dbFailOnError
'    Set rs = db.OpenRecordset(verifySqlText, dbOpenSnapshot)
'    If CLng(rs.fields("RemainingCount").value) <> 0 Then Err.Raise vbObjectError + 7636, "SharedNavPane_DeleteGroupObjectShortcuts", "Matching Navigation Pane shortcut rows remained after deletion."
'    rs.Close
'    Set rs = Nothing
'    wrk.CommitTrans
'    transactionStarted = False
'    Set wrk = Nothing
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim rollbackErrNumber As Long
'    Dim rollbackErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If Not rs Is Nothing Then rs.Close
'    Err.Clear
'    If transactionStarted And Not wrk Is Nothing Then wrk.Rollback
'    rollbackErrNumber = Err.Number
'    rollbackErrDescription = Err.Description
'    Set rs = Nothing
'    Set wrk = Nothing
'    On Error GoTo 0
'    If rollbackErrNumber <> 0 Then savedErrDescription = savedErrDescription & vbCrLf & "Rollback error " & CStr(rollbackErrNumber) & ": " & rollbackErrDescription
'    Err.Raise savedErrNumber, "SharedNavPane_DeleteGroupObjectShortcuts", "Unable to delete or verify Navigation Pane shortcuts." & vbCrLf & "Group ID: " & CStr(groupID) & vbCrLf & "Object ID: " & CStr(objectID) & vbCrLf & "Object name: " & cleanedObjectName & vbCrLf & "Delete SQL: " & sqlText & vbCrLf & "Verification SQL: " & verifySqlText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Sub
'Public Function SharedNavPane_AddAccessObjectToExistingCustomGroup(ByVal db As DAO.Database, ByVal objectName As String, ByVal objectType As AcObjectType, ByVal groupID As Long, ByRef detailText As String, Optional ByVal forceRefreshShortcut As Boolean = False, Optional ByVal ensureObjectIDRow As Boolean = True) As Boolean
'    'Adds one Access object to an existing custom Navigation Pane group and preserves the prior shortcut if every forced replacement attempt fails.
'    On Error GoTo ErrHandler
'    Dim wrk As DAO.workspace
'    Dim objectID As Long
'    Dim objectTypeID As Long
'    Dim firstDetailText As String
'    Dim secondDetailText As String
'    Dim thirdDetailText As String
'    Dim deleteSqlText As String
'    Dim verifyDeleteSqlText As String
'    Dim transactionStarted As Boolean
'    If db Is Nothing Then Err.Raise vbObjectError + 7731, "SharedNavPane_AddAccessObjectToExistingCustomGroup", "The DAO.Database argument is Nothing."
'    If groupID <= 0 Then Err.Raise vbObjectError + 7732, "SharedNavPane_AddAccessObjectToExistingCustomGroup", "The Navigation Pane group ID must be positive."
'    objectTypeID = SharedDao_MSysObjectTypeForAccessObjectType(objectType)
'    If objectTypeID = 0 Then
'        detailText = "Unsupported Access object type. Object: " & objectName & "; Access object type: " & CStr(objectType)
'        Exit Function
'    End If
'    objectID = SharedNavPane_MSysObjectIDStrict(db, objectName, objectType)
'    If objectID = 0 And ensureObjectIDRow Then
'        SharedNavPane_EnsureObjectIDRow objectName, objectType
'        objectID = SharedNavPane_MSysObjectIDStrict(db, objectName, objectType)
'    End If
'    If objectID = 0 Then
'        detailText = "Could not resolve object ID in MSysObjects. Object: " & objectName & "; MSysObjects.Type: " & CStr(objectTypeID)
'        Exit Function
'    End If
'    If forceRefreshShortcut Then
'        Set wrk = DBEngine.Workspaces(0)
'        wrk.BeginTrans
'        transactionStarted = True
'        SharedNavPane_DeleteExactShortcutRowsCore db, groupID, objectID, deleteSqlText, verifyDeleteSqlText
'        If SharedNavPane_TryStrictShortcutInsertWithinTransaction(db, groupID, objectID, objectName, objectTypeID, True, True, firstDetailText) Then GoTo CommitForcedReplacement
'        SharedNavPane_DeleteExactShortcutRowsCore db, groupID, objectID, deleteSqlText, verifyDeleteSqlText
'        If SharedNavPane_TryStrictShortcutInsertWithinTransaction(db, groupID, objectID, objectName, 0, True, True, secondDetailText) Then GoTo CommitForcedReplacement
'        SharedNavPane_DeleteExactShortcutRowsCore db, groupID, objectID, deleteSqlText, verifyDeleteSqlText
'        If SharedNavPane_TryStrictShortcutInsertWithinTransaction(db, groupID, objectID, objectName, objectTypeID, False, True, thirdDetailText) Then GoTo CommitForcedReplacement
'        wrk.Rollback
'        transactionStarted = False
'        Set wrk = Nothing
'        detailText = "All strict shortcut insert attempts failed; the pre-operation shortcut state was rolled back." & vbCrLf & "Object: " & objectName & vbCrLf & "ObjectID: " & CStr(objectID) & vbCrLf & "GroupID: " & CStr(groupID) & vbCrLf & "First attempt: " & firstDetailText & vbCrLf & "Second attempt: " & secondDetailText & vbCrLf & "Third attempt: " & thirdDetailText
'        Exit Function
'    End If
'    If SharedNavPane_GroupObjectShortcutExistsStrict(db, groupID, objectID) Then
'        SharedNavPane_AddAccessObjectToExistingCustomGroup = True
'        Exit Function
'    End If
'    If SharedNavPane_TryStrictShortcutInsert(db, groupID, objectID, objectName, objectTypeID, True, True, firstDetailText) Then
'        SharedNavPane_AddAccessObjectToExistingCustomGroup = True
'        Exit Function
'    End If
'    If SharedNavPane_GroupObjectShortcutExistsStrict(db, groupID, objectID) Then
'        SharedNavPane_AddAccessObjectToExistingCustomGroup = True
'        Exit Function
'    End If
'    If SharedNavPane_TryStrictShortcutInsert(db, groupID, objectID, objectName, 0, True, True, secondDetailText) Then
'        SharedNavPane_AddAccessObjectToExistingCustomGroup = True
'        Exit Function
'    End If
'    If SharedNavPane_GroupObjectShortcutExistsStrict(db, groupID, objectID) Then
'        SharedNavPane_AddAccessObjectToExistingCustomGroup = True
'        Exit Function
'    End If
'    If SharedNavPane_TryStrictShortcutInsert(db, groupID, objectID, objectName, objectTypeID, False, True, thirdDetailText) Then
'        SharedNavPane_AddAccessObjectToExistingCustomGroup = True
'        Exit Function
'    End If
'    If SharedNavPane_GroupObjectShortcutExistsStrict(db, groupID, objectID) Then
'        SharedNavPane_AddAccessObjectToExistingCustomGroup = True
'    Else
'        detailText = "All strict shortcut insert attempts failed." & vbCrLf & "Object: " & objectName & vbCrLf & "ObjectID: " & CStr(objectID) & vbCrLf & "GroupID: " & CStr(groupID) & vbCrLf & "First attempt: " & firstDetailText & vbCrLf & "Second attempt: " & secondDetailText & vbCrLf & "Third attempt: " & thirdDetailText
'    End If
'    Exit Function
'CommitForcedReplacement:
'    wrk.CommitTrans
'    transactionStarted = False
'    Set wrk = Nothing
'    SharedNavPane_AddAccessObjectToExistingCustomGroup = True
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim rollbackErrNumber As Long
'    Dim rollbackErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If transactionStarted And Not wrk Is Nothing Then wrk.Rollback
'    rollbackErrNumber = Err.Number
'    rollbackErrDescription = Err.Description
'    Set wrk = Nothing
'    On Error GoTo 0
'    detailText = "Navigation Pane grouping error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription
'    If rollbackErrNumber <> 0 Then detailText = detailText & vbCrLf & "Rollback error " & CStr(rollbackErrNumber) & ": " & rollbackErrDescription
'    SharedNavPane_AddAccessObjectToExistingCustomGroup = False
'End Function
'Private Sub SharedNavPane_DeleteExactShortcutRowsCore(ByVal db As DAO.Database, ByVal groupID As Long, ByVal objectID As Long, ByRef sqlText As String, ByRef verifySqlText As String)
'    'Deletes and verifies one exact group/object shortcut membership inside the caller-owned transaction.
'    On Error GoTo ErrHandler
'    Dim rs As DAO.Recordset
'    sqlText = "DELETE FROM [MSysNavPaneGroupToObjects] WHERE [GroupID]=" & CStr(groupID) & " AND [ObjectID]=" & CStr(objectID) & ";"
'    verifySqlText = "SELECT Count(*) AS RemainingCount FROM [MSysNavPaneGroupToObjects] WHERE [GroupID]=" & CStr(groupID) & " AND [ObjectID]=" & CStr(objectID) & ";"
'    db.Execute sqlText, dbFailOnError
'    Set rs = db.OpenRecordset(verifySqlText, dbOpenSnapshot)
'    If CLng(rs.fields("RemainingCount").value) <> 0 Then Err.Raise vbObjectError + 7921, "SharedNavPane_DeleteExactShortcutRowsCore", "Exact shortcut rows remained after deletion."
'    rs.Close
'    Set rs = Nothing
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If Not rs Is Nothing Then rs.Close
'    Set rs = Nothing
'    On Error GoTo 0
'    Err.Raise savedErrNumber, "SharedNavPane_DeleteExactShortcutRowsCore", "Could not delete and verify the exact shortcut rows inside the caller-owned transaction." & vbCrLf & "GroupID: " & CStr(groupID) & vbCrLf & "ObjectID: " & CStr(objectID) & vbCrLf & "Delete SQL: " & sqlText & vbCrLf & "Verification SQL: " & verifySqlText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Sub
'
'Private Function SharedNavPane_TryStrictShortcutInsertWithinTransaction(ByVal db As DAO.Database, ByVal groupID As Long, ByVal objectID As Long, ByVal objectName As String, ByVal typeValue As Long, ByVal includeTypeAndFlags As Boolean, ByVal includePosition As Boolean, ByRef detailText As String) As Boolean
'    'Attempts and verifies one shortcut insert without opening or committing a transaction because the caller owns the transaction boundary.
'    On Error GoTo ErrHandler
'    Dim useName As Boolean
'    Dim useType As Boolean
'    Dim useFlags As Boolean
'    Dim usePosition As Boolean
'    Dim expectedPosition As Long
'    Dim sqlText As String
'    detailText = vbNullString
'    SharedNavPane_InsertGroupObjectShortcutStrictCore db, groupID, objectID, objectName, typeValue, includeTypeAndFlags, includePosition, useName, useType, useFlags, usePosition, expectedPosition, sqlText
'    If SharedNavPane_GroupObjectShortcutValuesMatchStrict(db, groupID, objectID, objectName, typeValue, useName, useType, useFlags, usePosition, expectedPosition) = False Then Err.Raise vbObjectError + 7922, "SharedNavPane_TryStrictShortcutInsertWithinTransaction", "The inserted shortcut did not match every field included in the write."
'    SharedNavPane_TryStrictShortcutInsertWithinTransaction = True
'    Exit Function
'ErrHandler:
'    detailText = "Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
'End Function
'
'Private Function SharedNavPane_TryStrictShortcutInsert(ByVal db As DAO.Database, ByVal groupID As Long, ByVal objectID As Long, ByVal objectName As String, ByVal typeValue As Long, ByVal includeTypeAndFlags As Boolean, ByVal includePosition As Boolean, ByRef detailText As String) As Boolean
'    'Attempts one transactional strict Navigation Pane shortcut insert and returns detailed failure text.
'    On Error GoTo ErrHandler
'    detailText = vbNullString
'    SharedNavPane_TryStrictShortcutInsert = SharedNavPane_InsertGroupObjectShortcutWriteOnlyStrict(db, groupID, objectID, objectName, typeValue, includeTypeAndFlags, includePosition)
'    Exit Function
'ErrHandler:
'    detailText = "Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
'End Function
'' ================================================================================================================================================================================
'' NSS
'' ================================================================================================================================================================================
'Public Function NSS_PropPair(ByVal propName As String, ByVal propValue As Variant) As String
'    'Builds a name=value property fragment when propValue has text.
'    'Consolidated old names: local TV DATA GENERATOR MODULE PropPair.
'    If SharedText_HasText(propValue) Then NSS_PropPair = propName & "=" & SharedText_Nz(propValue)
'End Function
'
'Public Function NSS_JoinProps(ParamArray parts() As Variant) As String
'    'Joins nonblank property/text fragments with one space.
'    'Consolidated old names: local TV DATA GENERATOR MODULE JoinProps.
'    Dim i As Long
'    Dim s As String
'    Dim p As String
'    For i = LBound(parts) To UBound(parts)
'        p = Trim$(SharedText_Nz(parts(i)))
'        If Len(p) > 0 Then s = s & IIf(Len(s) > 0, " ", vbNullString) & p
'    Next i
'    NSS_JoinProps = Trim$(s)
'End Function
'
'Public Function NSS_PropValueFromProps(ByVal propsText As Variant, ByVal propName As String) As String
'    'Extracts a property value from a space-delimited property string.
'    'Consolidated old names: local TV DATA GENERATOR MODULE PropValueFromProps.
'    Dim s As String
'    Dim target As String
'    Dim startPos As Long
'    Dim i As Long
'    Dim ch As String
'    Dim quoteChar As String
'    Dim inQuote As Boolean
'    s = SharedText_Nz(propsText)
'    target = propName & "="
'    startPos = InStr(1, s, target, vbTextCompare)
'    If startPos = 0 Then
'        NSS_PropValueFromProps = vbNullString
'        Exit Function
'    End If
'    i = startPos + Len(target)
'    Do While i <= Len(s)
'        ch = Mid$(s, i, 1)
'        If inQuote Then
'            NSS_PropValueFromProps = NSS_PropValueFromProps & ch
'            If ch = quoteChar Then inQuote = False
'        ElseIf ch = "'" Or ch = """" Then
'            inQuote = True
'            quoteChar = ch
'            NSS_PropValueFromProps = NSS_PropValueFromProps & ch
'        ElseIf ch = " " Then
'            Exit Function
'        Else
'            NSS_PropValueFromProps = NSS_PropValueFromProps & ch
'        End If
'        i = i + 1
'    Loop
'End Function
'
'Public Function NSS_RemovePropFromProps(ByVal propsText As Variant, ByVal propName As String) As String
'    'Removes a name=value fragment from a space-delimited property string.
'    'Consolidated old names: local TV DATA GENERATOR MODULE RemovePropFromProps.
'    Dim parts() As String
'    Dim i As Long
'    Dim p As String
'    Dim s As String
'    parts = Split(Trim$(SharedText_Nz(propsText)), " ")
'    For i = LBound(parts) To UBound(parts)
'        p = Trim$(parts(i))
'        If Len(p) > 0 Then
'            If StrComp(Left$(p, Len(propName) + 1), propName & "=", vbTextCompare) <> 0 Then
'                s = s & IIf(Len(s) > 0, " ", vbNullString) & p
'            End If
'        End If
'    Next i
'    NSS_RemovePropFromProps = Trim$(s)
'End Function
'
'Public Function NSS_PropsForEntry(ByVal entryText As String, ByVal propsText As Variant) As String
'    'Returns properties appropriate for an NSS entry.
'    'Consolidated old names: local TV DATA GENERATOR MODULE PropsForEntry.
'    If LCase$(Trim$(entryText)) = "modify" Then
'        NSS_PropsForEntry = SharedText_Nz(propsText)
'    Else
'        NSS_PropsForEntry = NSS_RemovePropFromProps(propsText, "menu")
'    End If
'End Function
'
'Public Sub NSS_AppendDelimitedValue(ByRef resultText As String, ByVal valueText As Variant, Optional ByVal delimiter As String = ", ")
'    'Appends a nonblank value to a delimited string.
'    'Consolidated old names: AppendCommaValue, AppendSelect.
'    If Not SharedText_HasText(valueText) Then Exit Sub
'    If Len(resultText) > 0 Then resultText = resultText & delimiter
'    resultText = resultText & SharedText_Nz(valueText)
'End Sub
'
'Public Sub NSS_AppendTokenIfTruthy(ByRef resultText As String, ByVal flagValue As Variant, ByVal tokenText As String, Optional ByVal delimiter As String = "|")
'    'Appends tokenText to a delimited token string when flagValue is truthy.
'    'Consolidated old names: AddTypePart.
'    If SharedValue_CoerceBoolean(flagValue) Then
'        If Len(resultText) > 0 Then resultText = resultText & delimiter
'        resultText = resultText & tokenText
'    End If
'End Sub
'
'Public Sub NSS_AppendRecordsetTokenIfTruthy(ByRef resultText As String, ByVal rs As DAO.Recordset, ByVal fieldName As String, ByVal tokenText As String, Optional ByVal delimiter As String = "|")
'    'Appends tokenText when a recordset field exists and is truthy.
'    'Consolidated old names: AddTypeFieldPart.
'    If Not SharedDao_RecordsetFieldExists(rs, fieldName) Then Exit Sub
'    NSS_AppendTokenIfTruthy resultText, rs.fields(fieldName).value, tokenText, delimiter
'End Sub
'
'Public Function NSS_CleanNodeKey(ByVal nodeKey As Variant) As String
'    'Normalizes an NSS TreeView node key by trimming whitespace and removing ":".
'    'Consolidated old names: local TV APP FORM MODULE NSSCleanNodeKey.
'    NSS_CleanNodeKey = Replace(Trim$(CStr(Nz(nodeKey, vbNullString))), ":", vbNullString)
'End Function
'
'Public Function NSS_NodePrefix(ByVal nodeKey As Variant) As String
'    'Returns the first character of a normalized NSS node key.
'    'Consolidated old names: local TV APP FORM MODULE NSSNodePrefix.
'    Dim cleanKey As String
'    cleanKey = NSS_CleanNodeKey(nodeKey)
'    If Len(cleanKey) > 0 Then NSS_NodePrefix = UCase$(Left$(cleanKey, 1))
'End Function
'
'Public Function NSS_NodeID(ByVal nodeKey As Variant) As Long
'    'Returns the numeric ID portion of a normalized NSS node key.
'    'Consolidated old names: local TV APP FORM MODULE NSSNodeID.
'    Dim cleanKey As String
'    Dim idText As String
'    cleanKey = NSS_CleanNodeKey(nodeKey)
'    If Len(cleanKey) < 2 Then Exit Function
'    idText = Mid$(cleanKey, 2)
'    If Not IsNumeric(idText) Then Exit Function
'    NSS_NodeID = CLng(idText)
'End Function
'
'Public Function NSS_NodeIsFile(ByVal nodeKey As Variant) As Boolean
'    'Returns True when an NSS node key identifies a file node.
'    'Consolidated old names: local TV APP FORM MODULE NSSNodeIsFile.
'    NSS_NodeIsFile = (NSS_NodePrefix(nodeKey) = "F")
'End Function
'
'Public Function NSS_NodeIsMenu(ByVal nodeKey As Variant) As Boolean
'    'Returns True when an NSS node key identifies a menu node.
'    'Consolidated old names: local TV APP FORM MODULE NSSNodeIsMenu.
'    NSS_NodeIsMenu = (NSS_NodePrefix(nodeKey) = "M")
'End Function
'
'Public Function NSS_NodeIsItem(ByVal nodeKey As Variant) As Boolean
'    'Returns True when an NSS node key identifies an item node.
'    'Consolidated old names: local TV APP FORM MODULE NSSNodeIsItem.
'    NSS_NodeIsItem = (NSS_NodePrefix(nodeKey) = "I")
'End Function
'
'Public Function NSS_NodeTypeText(ByVal nodeKey As Variant) As String
'    'Returns the data table name for an NSS node key prefix.
'    'Consolidated old names: local TV APP FORM MODULE NSSNodeTypeText.
'    Select Case NSS_NodePrefix(nodeKey)
'        Case "F"
'            NSS_NodeTypeText = "file"
'        Case "M"
'            NSS_NodeTypeText = "menus"
'        Case "I"
'            NSS_NodeTypeText = "items"
'        Case Else
'            NSS_NodeTypeText = vbNullString
'    End Select
'End Function
'
'Public Function NSS_TreeTag(ByVal nodeKey As Variant) As String
'    'Builds the TreeView node Tag value as tableName|ID.
'    'Consolidated old names: local TV APP FORM MODULE NSSTreeTag.
'    NSS_TreeTag = NSS_NodeTypeText(nodeKey) & "|" & CStr(NSS_NodeID(nodeKey))
'End Function
'
'Public Function NSS_NodeParseNo(ByVal nodeKey As Variant) As Boolean
'    'Returns True when a menu/item NSS node has Parse=False.
'    'Consolidated old names: local TV APP FORM MODULE NSSNodeParseNo.
'    Dim idValue As Long
'    idValue = NSS_NodeID(nodeKey)
'    If idValue <= 0 Then Exit Function
'    Select Case NSS_NodePrefix(nodeKey)
'        Case "M"
'            NSS_NodeParseNo = (Nz(DLookup("[Parse]", "[menus]", "[ID]=" & CStr(idValue)), True) = False)
'        Case "I"
'            NSS_NodeParseNo = (Nz(DLookup("[Parse]", "[items]", "[ID]=" & CStr(idValue)), True) = False)
'    End Select
'End Function
'
'Public Function TVW_NodeTextEquals(ByVal Node As Object, ByVal CompareText As String) As Boolean
'    'Returns True when a TreeView node text matches compareText case-insensitively.
'    'Consolidated old names: IsImportsNssNode, ShellTVW_IsResourcesFileNode.
'    If Node Is Nothing Then Exit Function
'    TVW_NodeTextEquals = (StrComp(CStr(Node.text & vbNullString), CompareText, vbTextCompare) = 0)
'End Function
'
'Public Function NSS_NodeFileIDText(ByVal Node As Object) As String
'    'Returns the file ID text for a file TreeView node.
'    'Consolidated old names: ShellTVW_GetNodeFileValue.
'    NSS_NodeFileIDText = vbNullString
'    If Node Is Nothing Then Exit Function
'    If NSS_NodeIsFile(Node.key) Then NSS_NodeFileIDText = CStr(NSS_NodeID(Node.key))
'End Function
'
'Public Sub NSS_ApplyTreeNodeParseVisuals(ByVal Node As Object, ByVal ParseNo As Boolean)
'    'Applies the standard visual style for parsed/unparsed NSS TreeView nodes.
'    'Consolidated old names: ShellTVW_ApplyTreeNodeVisuals.
'    On Error Resume Next
'    Node.Bold = ParseNo
'    If ParseNo Then
'        Node.ForeColor = 8421504
'    Else
'        Node.ForeColor = 0
'    End If
'    On Error GoTo 0
'End Sub
'
'' ================================================================================================================================================================================
'' RUN and UI
'' ================================================================================================================================================================================
'Public Sub SharedFile_OpenTextFileInNotepad(ByVal filePath As String)
'    'Opens a text file in Notepad.
'    'Use for logs or generated text files after first confirming the file exists.
'    On Error GoTo ErrHandler
'    If Len(filePath) = 0 Then Exit Sub
'    If Dir$(filePath, vbNormal) = vbNullString Then
'        MsgBox "Log file not found:" & vbCrLf & filePath, vbExclamation, "Open Log"
'        Exit Sub
'    End If
'    shell "notepad.exe " & SharedText_DoubleQuote(filePath), vbNormalFocus
'    Exit Sub
'ErrHandler:
'    MsgBox "Could not open log file." & vbCrLf & vbCrLf & "Log:" & vbCrLf & filePath & vbCrLf & vbCrLf & "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Open Log"
'End Sub
'
'Public Sub SharedMessage_ShowFailureDialogWithOpenLog(ByVal messageText As String, ByVal dialogTitle As String, Optional ByVal logPath As String = vbNullString)
'    'Shows a failure message and optionally prompts to open a related log file.
'    'Consolidated old names: local TV NSS IMPORT MODULE ShowFailureDialogWithOpenLog.
'    Dim response As VbMsgBoxResult
'    If Len(logPath) > 0 And Dir$(logPath, vbNormal) <> vbNullString Then
'        response = MsgBox(messageText & vbCrLf & vbCrLf & "Open the log?", vbCritical + vbYesNo, dialogTitle)
'        If response = vbYes Then SharedFile_OpenTextFileInNotepad logPath
'    Else
'        MsgBox messageText, vbCritical, dialogTitle
'    End If
'End Sub
'
'Public Function SharedRun_ApplicationRunWithArgs(ByVal ProcedureName As String, ByVal args As Collection) As Variant
'    'Intended use: Invokes a dynamically named Public procedure or function with arguments supplied in a Collection.
'    Dim argCount As Long
'    If args Is Nothing Then
'        argCount = 0
'    Else
'        argCount = args.count
'    End If
'    Select Case argCount
'        Case 0
'            SharedRun_ApplicationRunWithArgs = Application.Run(ProcedureName)
'        Case 1
'            SharedRun_ApplicationRunWithArgs = Application.Run(ProcedureName, args(1))
'        Case 2
'            SharedRun_ApplicationRunWithArgs = Application.Run(ProcedureName, args(1), args(2))
'        Case 3
'            SharedRun_ApplicationRunWithArgs = Application.Run(ProcedureName, args(1), args(2), args(3))
'        Case 4
'            SharedRun_ApplicationRunWithArgs = Application.Run(ProcedureName, args(1), args(2), args(3), args(4))
'        Case 5
'            SharedRun_ApplicationRunWithArgs = Application.Run(ProcedureName, args(1), args(2), args(3), args(4), args(5))
'        Case 6
'            SharedRun_ApplicationRunWithArgs = Application.Run(ProcedureName, args(1), args(2), args(3), args(4), args(5), args(6))
'        Case 7
'            SharedRun_ApplicationRunWithArgs = Application.Run(ProcedureName, args(1), args(2), args(3), args(4), args(5), args(6), args(7))
'        Case 8
'            SharedRun_ApplicationRunWithArgs = Application.Run(ProcedureName, args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8))
'        Case 9
'            SharedRun_ApplicationRunWithArgs = Application.Run(ProcedureName, args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9))
'        Case 10
'            SharedRun_ApplicationRunWithArgs = Application.Run(ProcedureName, args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10))
'        Case Else
'            Err.Raise vbObjectError + 7952, "SharedRun_ApplicationRunWithArgs", "Application.Run supports up to 10 collected arguments in this helper." & vbCrLf & "Procedure: " & ProcedureName & vbCrLf & "Argument count: " & CStr(argCount)
'    End Select
'End Function
'
'Public Function SharedRun_TryPublicProcedure0(ByVal ProcedureName As String, ByRef errNumber As Long, ByRef errDescription As String) As Boolean
'    'Intended use: Attempts a dynamically named zero-argument Public procedure and returns captured failure details without raising.
'    On Error GoTo ErrHandler
'    errNumber = 0
'    errDescription = vbNullString
'    Application.Run ProcedureName
'    SharedRun_TryPublicProcedure0 = True
'    Exit Function
'ErrHandler:
'    errNumber = Err.Number
'    errDescription = Err.Description
'    SharedRun_TryPublicProcedure0 = False
'End Function
'
'Public Function SharedRun_TryPublicProcedure1(ByVal ProcedureName As String, ByVal arg1 As Variant, ByRef errNumber As Long, ByRef errDescription As String) As Boolean
'    'Intended use: Attempts a dynamically named one-argument Public procedure and returns captured failure details without raising.
'    On Error GoTo ErrHandler
'    errNumber = 0
'    errDescription = vbNullString
'    Application.Run ProcedureName, arg1
'    SharedRun_TryPublicProcedure1 = True
'    Exit Function
'ErrHandler:
'    errNumber = Err.Number
'    errDescription = Err.Description
'    SharedRun_TryPublicProcedure1 = False
'End Function
'
'Public Function SharedRun_RequiredPublicStringFunction(ByVal functionName As String, Optional ByVal callerName As String = "SharedRun_RequiredPublicStringFunction", Optional ByVal expectedModuleName As String = vbNullString, Optional ByVal currentStep As String = vbNullString, Optional ByVal dependencyNote As String = vbNullString) As String
'    'Intended use: Runs a required dynamically named Public String function and raises with dependency context when it fails or returns blank.
'    On Error GoTo ErrHandler
'    SharedRun_RequiredPublicStringFunction = Trim$(CStr(Application.Run(functionName)))
'    If Len(SharedRun_RequiredPublicStringFunction) = 0 Then
'        Err.Raise vbObjectError + 7953, callerName, "Required public function returned a blank value." & vbCrLf & "Function: " & functionName & vbCrLf & "Current step: " & currentStep
'    End If
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim errText As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    errText = "Required public function failed." & vbCrLf
'    errText = errText & "Function: " & functionName & vbCrLf
'    If Len(expectedModuleName) > 0 Then errText = errText & "Expected module: " & expectedModuleName & vbCrLf
'    If Len(currentStep) > 0 Then errText = errText & "Current step: " & currentStep & vbCrLf
'    If Len(dependencyNote) > 0 Then errText = errText & dependencyNote & vbCrLf
'    errText = errText & "Original error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription
'    Err.Raise savedErrNumber, callerName, errText
'End Function
'
'Public Sub SharedProgress_Begin(ByRef progressActive As Boolean, ByRef currentStep As Long, ByRef maxSteps As Long, ByVal captionText As String, Optional ByVal statusPrefix As String = vbNullString)
'    'Intended use: Starts best-effort shared progress and resets caller state if meter initialization fails.
'    Dim errorNumber As Long
'    Dim errorSource As String
'    Dim errorDescription As String
'    Call SharedProgress_TryBegin(progressActive, currentStep, maxSteps, captionText, statusPrefix, errorNumber, errorSource, errorDescription)
'End Sub
'
'Public Sub SharedProgress_Update(ByRef progressActive As Boolean, ByRef currentStep As Long, ByRef maxSteps As Long, ByVal statusText As String, Optional ByVal statusPrefix As String = vbNullString)
'    'Intended use: Advances an active shared progress meter and updates its status text without raising display errors.
'    On Error Resume Next
'    statusText = Trim$(statusText)
'    If Len(statusText) = 0 Then Exit Sub
'    If progressActive = False Then Exit Sub
'    currentStep = currentStep + 1
'    If currentStep > maxSteps Then currentStep = maxSteps
'    SysCmd acSysCmdUpdateMeter, currentStep
'    SysCmd acSysCmdSetStatus, statusPrefix & statusText
'    DoEvents
'    On Error GoTo 0
'End Sub
'
'Public Sub SharedProgress_Pulse(ByVal progressActive As Boolean, ByVal statusText As String, Optional ByVal statusPrefix As String = vbNullString)
'    'Intended use: Updates status text for an active shared progress operation without advancing its meter.
'    On Error Resume Next
'    statusText = Trim$(statusText)
'    If Len(statusText) = 0 Then Exit Sub
'    If progressActive = False Then Exit Sub
'    SysCmd acSysCmdSetStatus, statusPrefix & statusText
'    DoEvents
'    On Error GoTo 0
'End Sub
'
'Public Sub SharedProgress_PulseEvery(ByVal progressActive As Boolean, ByVal currentCount As Long, ByVal intervalCount As Long, ByVal statusText As String, Optional ByVal statusPrefix As String = vbNullString)
'    'Intended use: Updates shared progress status only at the caller-specified item interval.
'    If intervalCount < 1 Then intervalCount = 1
'    If currentCount <= 0 Then Exit Sub
'    If currentCount Mod intervalCount <> 0 Then Exit Sub
'    SharedProgress_Pulse progressActive, statusText & " " & CStr(currentCount), statusPrefix
'End Sub
'
'Public Sub SharedProgress_Complete(ByRef progressActive As Boolean, ByRef currentStep As Long, ByRef maxSteps As Long, ByVal completeStatus As String, Optional ByVal clearAfterComplete As Boolean = True)
'    'Intended use: Completes and normalizes caller-owned shared progress state after successful work.
'    On Error Resume Next
'    If progressActive Then SysCmd acSysCmdUpdateMeter, maxSteps
'    SysCmd acSysCmdSetStatus, completeStatus
'    DoEvents
'    SharedProgress_End progressActive, currentStep, maxSteps, clearAfterComplete
'    On Error GoTo 0
'End Sub
'
'Public Sub SharedProgress_Fail(ByRef progressActive As Boolean, ByRef currentStep As Long, ByRef maxSteps As Long, ByVal failedStatus As String)
'    'Intended use: Removes the active meter, normalizes caller state, and leaves a failure status message visible.
'    On Error Resume Next
'    If progressActive Then SysCmd acSysCmdRemoveMeter
'    progressActive = False
'    currentStep = 0
'    maxSteps = 0
'    SysCmd acSysCmdSetStatus, failedStatus
'    DoEvents
'    On Error GoTo 0
'End Sub
'
'Public Sub SharedProgress_End(ByRef progressActive As Boolean, ByRef currentStep As Long, ByRef maxSteps As Long, Optional ByVal clearStatus As Boolean = True)
'    'Intended use: Ends caller-owned shared progress and optionally clears the Access status bar.
'    On Error Resume Next
'    If progressActive Then SysCmd acSysCmdRemoveMeter
'    progressActive = False
'    currentStep = 0
'    maxSteps = 0
'    If clearStatus Then SysCmd acSysCmdClearStatus
'    On Error GoTo 0
'End Sub
'
'' Public safe progress adapter.
'' Intended use: execute SharedProgress_Begin without allowing a status-bar failure to replace the caller's primary operation result.
'Public Function SharedProgress_TryBegin(ByRef progressActive As Boolean, ByRef currentStep As Long, ByRef maxSteps As Long, ByVal captionText As String, ByVal statusPrefix As String, ByRef errorNumber As Long, ByRef errorSource As String, ByRef errorDescription As String) As Boolean
'    'Intent: Starts shared progress without raising and returns captured error details to callers that own warning/reporting policy.
'    On Error GoTo ErrHandler
'    errorNumber = 0
'    errorSource = vbNullString
'    errorDescription = vbNullString
'    If maxSteps < 1 Then maxSteps = 1
'    currentStep = 0
'    progressActive = True
'    SysCmd acSysCmdInitMeter, captionText, maxSteps
'    SysCmd acSysCmdUpdateMeter, currentStep
'    SysCmd acSysCmdSetStatus, statusPrefix & captionText
'    DoEvents
'    SharedProgress_TryBegin = True
'    Exit Function
'ErrHandler:
'    errorNumber = Err.Number
'    errorSource = Err.Source
'    errorDescription = Err.Description
'    SharedProgress_NormalizeAfterFailure progressActive, currentStep, maxSteps, errorDescription
'End Function
'
'' Public safe progress adapter.
'' Intended use: execute SharedProgress_Update while returning display failures to the caller for module-specific diagnostics.
'Public Function SharedProgress_TryUpdate(ByRef progressActive As Boolean, ByRef currentStep As Long, ByRef maxSteps As Long, ByVal statusText As String, ByVal statusPrefix As String, ByRef errorNumber As Long, ByRef errorSource As String, ByRef errorDescription As String) As Boolean
'    'Intent: Advances shared progress without raising and returns captured error details to callers that own warning/reporting policy.
'    On Error GoTo ErrHandler
'    errorNumber = 0
'    errorSource = vbNullString
'    errorDescription = vbNullString
'    statusText = Trim$(statusText)
'    If Len(statusText) = 0 Then SharedProgress_TryUpdate = True: Exit Function
'    If progressActive = False Then SharedProgress_TryUpdate = True: Exit Function
'    currentStep = currentStep + 1
'    If currentStep > maxSteps Then currentStep = maxSteps
'    SysCmd acSysCmdUpdateMeter, currentStep
'    SysCmd acSysCmdSetStatus, statusPrefix & statusText
'    DoEvents
'    SharedProgress_TryUpdate = True
'    Exit Function
'ErrHandler:
'    errorNumber = Err.Number
'    errorSource = Err.Source
'    errorDescription = Err.Description
'    SharedProgress_NormalizeAfterFailure progressActive, currentStep, maxSteps, errorDescription
'End Function
'
'' Public safe progress adapter.
'' Intended use: execute SharedProgress_Pulse while returning display failures to the caller for module-specific diagnostics.
'Public Function SharedProgress_TryPulse(ByVal progressActive As Boolean, ByVal statusText As String, ByVal statusPrefix As String, ByRef errorNumber As Long, ByRef errorSource As String, ByRef errorDescription As String) As Boolean
'    'Intent: Updates shared progress status without raising and returns captured error details to callers that own warning/reporting policy.
'    On Error GoTo ErrHandler
'    errorNumber = 0
'    errorSource = vbNullString
'    errorDescription = vbNullString
'    statusText = Trim$(statusText)
'    If Len(statusText) = 0 Then SharedProgress_TryPulse = True: Exit Function
'    If progressActive = False Then SharedProgress_TryPulse = True: Exit Function
'    SysCmd acSysCmdSetStatus, statusPrefix & statusText
'    DoEvents
'    SharedProgress_TryPulse = True
'    Exit Function
'ErrHandler:
'    errorNumber = Err.Number
'    errorSource = Err.Source
'    errorDescription = Err.Description
'End Function
'
'' Public safe progress adapter.
'' Intended use: execute SharedProgress_PulseEvery while returning display failures to the caller for module-specific diagnostics.
'Public Function SharedProgress_TryPulseEvery(ByVal progressActive As Boolean, ByVal currentCount As Long, ByVal intervalCount As Long, ByVal statusText As String, ByVal statusPrefix As String, ByRef errorNumber As Long, ByRef errorSource As String, ByRef errorDescription As String) As Boolean
'    'Intent: Pulses shared progress at the requested interval without raising and returns captured error details.
'    On Error GoTo ErrHandler
'    errorNumber = 0
'    errorSource = vbNullString
'    errorDescription = vbNullString
'    If intervalCount < 1 Then intervalCount = 1
'    If currentCount <= 0 Then SharedProgress_TryPulseEvery = True: Exit Function
'    If currentCount Mod intervalCount <> 0 Then SharedProgress_TryPulseEvery = True: Exit Function
'    SharedProgress_TryPulseEvery = SharedProgress_TryPulse(progressActive, statusText & " " & CStr(currentCount), statusPrefix, errorNumber, errorSource, errorDescription)
'    Exit Function
'ErrHandler:
'    errorNumber = Err.Number
'    errorSource = Err.Source
'    errorDescription = Err.Description
'End Function
'
'' Public safe progress adapter.
'' Intended use: execute SharedProgress_Complete while returning display failures to the caller for module-specific diagnostics.
'Public Function SharedProgress_TryComplete(ByRef progressActive As Boolean, ByRef currentStep As Long, ByRef maxSteps As Long, ByVal completeStatus As String, ByVal clearAfterComplete As Boolean, ByRef errorNumber As Long, ByRef errorSource As String, ByRef errorDescription As String) As Boolean
'    'Intent: Completes shared progress without raising and returns captured error details to callers that own warning/reporting policy.
'    On Error GoTo ErrHandler
'    errorNumber = 0
'    errorSource = vbNullString
'    errorDescription = vbNullString
'    If progressActive Then SysCmd acSysCmdUpdateMeter, maxSteps
'    SysCmd acSysCmdSetStatus, completeStatus
'    DoEvents
'    If progressActive Then SysCmd acSysCmdRemoveMeter
'    progressActive = False
'    currentStep = 0
'    maxSteps = 0
'    If clearAfterComplete Then SysCmd acSysCmdClearStatus
'    SharedProgress_TryComplete = True
'    Exit Function
'ErrHandler:
'    errorNumber = Err.Number
'    errorSource = Err.Source
'    errorDescription = Err.Description
'    SharedProgress_NormalizeAfterFailure progressActive, currentStep, maxSteps, errorDescription
'End Function
'
'' Public safe progress adapter.
'' Intended use: execute SharedProgress_Fail while returning display failures to the caller for module-specific diagnostics.
'Public Function SharedProgress_TryFail(ByRef progressActive As Boolean, ByRef currentStep As Long, ByRef maxSteps As Long, ByVal failedStatus As String, ByRef errorNumber As Long, ByRef errorSource As String, ByRef errorDescription As String) As Boolean
'    'Intent: Marks shared progress failed without raising and returns captured error details to callers that own warning/reporting policy.
'    On Error GoTo ErrHandler
'    errorNumber = 0
'    errorSource = vbNullString
'    errorDescription = vbNullString
'    If progressActive Then SysCmd acSysCmdRemoveMeter
'    progressActive = False
'    currentStep = 0
'    maxSteps = 0
'    SysCmd acSysCmdSetStatus, failedStatus
'    DoEvents
'    SharedProgress_TryFail = True
'    Exit Function
'ErrHandler:
'    errorNumber = Err.Number
'    errorSource = Err.Source
'    errorDescription = Err.Description
'    SharedProgress_NormalizeAfterFailure progressActive, currentStep, maxSteps, errorDescription
'End Function
'
'' Public safe progress adapter.
'' Intended use: execute SharedProgress_End while returning display failures to the caller for module-specific diagnostics.
'Public Function SharedProgress_TryEnd(ByRef progressActive As Boolean, ByRef currentStep As Long, ByRef maxSteps As Long, ByVal clearStatus As Boolean, ByRef errorNumber As Long, ByRef errorSource As String, ByRef errorDescription As String) As Boolean
'    'Intent: Ends shared progress without raising and returns captured error details to callers that own warning/reporting policy.
'    On Error GoTo ErrHandler
'    errorNumber = 0
'    errorSource = vbNullString
'    errorDescription = vbNullString
'    If progressActive Then SysCmd acSysCmdRemoveMeter
'    progressActive = False
'    currentStep = 0
'    maxSteps = 0
'    If clearStatus Then SysCmd acSysCmdClearStatus
'    SharedProgress_TryEnd = True
'    Exit Function
'ErrHandler:
'    errorNumber = Err.Number
'    errorSource = Err.Source
'    errorDescription = Err.Description
'    SharedProgress_NormalizeAfterFailure progressActive, currentStep, maxSteps, errorDescription
'End Function
'
'Private Sub SharedProgress_NormalizeAfterFailure(ByRef progressActive As Boolean, ByRef currentStep As Long, ByRef maxSteps As Long, ByRef errorDescription As String)
'    'Removes any partially initialized meter and resets caller state while retaining the original display failure.
'    Dim cleanupErrorNumber As Long
'    Dim cleanupErrorDescription As String
'    On Error Resume Next
'    If progressActive Then SysCmd acSysCmdRemoveMeter
'    cleanupErrorNumber = Err.Number
'    cleanupErrorDescription = Err.Description
'    progressActive = False
'    currentStep = 0
'    maxSteps = 0
'    Err.Clear
'    On Error GoTo 0
'    If cleanupErrorNumber <> 0 Then errorDescription = errorDescription & vbCrLf & "Progress cleanup error " & CStr(cleanupErrorNumber) & ": " & cleanupErrorDescription
'End Sub
'
'Public Function SharedMessage_SplitTextForMsgBox(ByVal promptText As String, Optional ByVal maxPageLength As Long = 3000) As Collection
'    'Intended use: Splits long message text into readable pages that respect the requested maximum length.
'    Dim pages As Collection
'    Dim remainingText As String
'    Dim pageText As String
'    Dim breakPos As Long
'    Dim searchText As String
'    Set pages = New Collection
'    If maxPageLength < 250 Then maxPageLength = 250
'    remainingText = CStr(promptText)
'    Do While Len(remainingText) > maxPageLength
'        searchText = Left$(remainingText, maxPageLength)
'        breakPos = InStrRev(searchText, vbCrLf)
'        If breakPos < (maxPageLength \ 2) Then breakPos = InStrRev(searchText, vbLf)
'        If breakPos < (maxPageLength \ 2) Then breakPos = InStrRev(searchText, " ")
'        If breakPos < (maxPageLength \ 2) Then breakPos = maxPageLength
'        pageText = Left$(remainingText, breakPos)
'        pages.Add SharedMessage_RTrimText(pageText)
'        remainingText = SharedMessage_TrimLeadingSeparators(Mid$(remainingText, breakPos + 1))
'    Loop
'    If Len(remainingText) > 0 Then pages.Add remainingText
'    If pages.count = 0 Then pages.Add vbNullString
'    Set SharedMessage_SplitTextForMsgBox = pages
'End Function
'
'Public Sub SharedMessage_ShowLong(ByVal promptText As String, ByVal iconStyle As VbMsgBoxStyle, ByVal TitleText As String, Optional ByVal maxPageLength As Long = 3000, Optional ByVal titlePageOpenText As String = " (", Optional ByVal titlePageMiddleText As String = " of ", Optional ByVal titlePageCloseText As String = ")")
'    'Intended use: Displays long text as a sequence of paged message boxes.
'    On Error GoTo ErrHandler
'    Dim pages As Collection
'    Dim pageIndex As Long
'    Dim pageTitle As String
'    Set pages = SharedMessage_SplitTextForMsgBox(promptText, maxPageLength)
'    For pageIndex = 1 To pages.count
'        pageTitle = TitleText
'        If pages.count > 1 Then pageTitle = TitleText & titlePageOpenText & CStr(pageIndex) & titlePageMiddleText & CStr(pages.count) & titlePageCloseText
'        MsgBox CStr(pages(pageIndex)), vbOKOnly Or iconStyle, pageTitle
'    Next pageIndex
'    Exit Sub
'ErrHandler:
'    MsgBox Left$(promptText, maxPageLength), vbOKOnly Or iconStyle, TitleText
'End Sub
'
'Public Function SharedMessage_ConfirmLong(ByVal promptText As String, ByVal TitleText As String, ByVal finalQuestionText As String, Optional ByVal maxPageLength As Long = 3000, Optional ByVal trailingQuestionText As String = vbNullString) As VbMsgBoxResult
'    'Intended use: Displays paged explanatory text and returns the final confirmation response.
'    On Error GoTo ErrHandler
'    Dim detailText As String
'    If maxPageLength < 250 Then maxPageLength = 250
'    If Len(promptText) <= maxPageLength Then
'        SharedMessage_ConfirmLong = MsgBox(promptText, vbExclamation Or vbYesNo Or vbDefaultButton2, TitleText)
'    Else
'        detailText = promptText
'        If Len(trailingQuestionText) > 0 Then detailText = SharedMessage_RemoveTrailingText(detailText, trailingQuestionText)
'        SharedMessage_ShowLong detailText, vbExclamation, TitleText & " Details", maxPageLength
'        SharedMessage_ConfirmLong = MsgBox(finalQuestionText, vbExclamation Or vbYesNo Or vbDefaultButton2, TitleText)
'    End If
'    Exit Function
'ErrHandler:
'    SharedMessage_ConfirmLong = MsgBox(finalQuestionText, vbExclamation Or vbYesNo Or vbDefaultButton2, TitleText)
'End Function
'
'Public Function SharedMessage_RemoveTrailingText(ByVal promptText As String, ByVal trailingText As String) As String
'    'Intended use: Removes a known trailing prompt fragment before message repaging.
'    Dim resultText As String
'    resultText = SharedMessage_RTrimText(promptText)
'    If Len(trailingText) > 0 Then
'        If Right$(resultText, Len(trailingText)) = trailingText Then resultText = Left$(resultText, Len(resultText) - Len(trailingText))
'    End If
'    SharedMessage_RemoveTrailingText = SharedMessage_RTrimText(resultText)
'End Function
'
'Public Function SharedMessage_RTrimText(ByVal valueText As String) As String
'    'Intended use: Removes trailing whitespace and line separators from message text.
'    valueText = RTrim$(valueText)
'    Do While Len(valueText) > 0
'        If Right$(valueText, 2) = vbCrLf Then
'            valueText = Left$(valueText, Len(valueText) - 2)
'        ElseIf Right$(valueText, 1) = vbCr Or Right$(valueText, 1) = vbLf Or Right$(valueText, 1) = vbTab Or Right$(valueText, 1) = " " Then
'            valueText = Left$(valueText, Len(valueText) - 1)
'        Else
'            Exit Do
'        End If
'    Loop
'    SharedMessage_RTrimText = valueText
'End Function
'
'Public Function SharedMessage_TrimLeadingSeparators(ByVal valueText As String) As String
'    'Intended use: Removes leading message separators before pagination or display.
'    Do While Len(valueText) > 0
'        If Left$(valueText, 2) = vbCrLf Then
'            valueText = Mid$(valueText, 3)
'        ElseIf Left$(valueText, 1) = vbCr Or Left$(valueText, 1) = vbLf Or Left$(valueText, 1) = vbTab Or Left$(valueText, 1) = " " Then
'            valueText = Mid$(valueText, 2)
'        Else
'            Exit Do
'        End If
'    Loop
'    SharedMessage_TrimLeadingSeparators = valueText
'End Function
'
'Public Function SharedMessage_OneLine(ByVal valueText As String, Optional ByVal maxLength As Long = 700, Optional ByVal ellipsisText As String = "...") As String
'    'Intended use: Returns a whitespace-normalized single-line message with optional truncation.
'    valueText = Replace(CStr(valueText), vbCrLf, " ")
'    valueText = Replace(valueText, vbCr, " ")
'    valueText = Replace(valueText, vbLf, " ")
'    valueText = Replace(valueText, vbTab, " ")
'    Do While InStr(1, valueText, "  ", vbBinaryCompare) > 0
'        valueText = Replace(valueText, "  ", " ")
'    Loop
'    valueText = Trim$(valueText)
'    If maxLength < 50 Then maxLength = 50
'    If Len(valueText) > maxLength Then
'        SharedMessage_OneLine = Left$(valueText, maxLength - Len(ellipsisText)) & ellipsisText
'    Else
'        SharedMessage_OneLine = valueText
'    End If
'End Function
'
'Public Function SharedRun_PublicProcedure0WithMessage(ByVal ProcedureName As String, ByVal actionCaption As String, ByVal dialogTitle As String) As Boolean
'    'Intended use: Runs a zero-argument Public procedure and presents a concise success or failure message.
'    Dim errNumber As Long
'    Dim errDescription As String
'    If SharedRun_TryPublicProcedure0(ProcedureName, errNumber, errDescription) Then
'        SharedRun_PublicProcedure0WithMessage = True
'    Else
'        MsgBox "Could not run [" & actionCaption & "]." & vbCrLf & vbCrLf & "Error " & CStr(errNumber) & ": " & errDescription, vbExclamation, dialogTitle
'        SharedRun_PublicProcedure0WithMessage = False
'    End If
'End Function
'
'Public Function SharedRun_PublicProcedure1WithMessage(ByVal ProcedureName As String, ByVal arg1 As Variant, ByVal actionCaption As String, ByVal dialogTitle As String) As Boolean
'    'Intended use: Runs a one-argument Public procedure and presents a concise success or failure message.
'    Dim errNumber As Long
'    Dim errDescription As String
'    If SharedRun_TryPublicProcedure1(ProcedureName, arg1, errNumber, errDescription) Then
'        SharedRun_PublicProcedure1WithMessage = True
'    Else
'        MsgBox "Could not run [" & actionCaption & "]." & vbCrLf & vbCrLf & "Error " & CStr(errNumber) & ": " & errDescription, vbExclamation, dialogTitle
'        SharedRun_PublicProcedure1WithMessage = False
'    End If
'End Function
'
'' ================================================================================================================================================================================
'' TREEVIEW
'' ================================================================================================================================================================================
'Public Function TVW_NodeExists(ByVal tv As Object, ByVal nodeKey As String) As Boolean
'    'Intent: Checks whether a TreeView node key currently exists before selection, restoration, or state-recovery operations.
'    On Error GoTo ExitFunction
'    Dim n As Object
'    If tv Is Nothing Then Exit Function
'    If Len(nodeKey) = 0 Then Exit Function
'    Set n = tv.Nodes(nodeKey)
'    TVW_NodeExists = Not n Is Nothing
'ExitFunction:
'End Function
'
'Public Sub TVW_EnableDirectNodeTextEditing(ByVal tv As Object)
'    'Intended use: Enables direct label editing for the supplied TreeView control.
'    If tv Is Nothing Then Exit Sub
'    tv.LabelEdit = TVW_LABEL_EDIT_AUTOMATIC
'End Sub
'
'Public Sub TVW_DisableDirectNodeTextEditing(ByVal tv As Object)
'    'Intended use: Disables direct label editing for the supplied TreeView control.
'    If tv Is Nothing Then Exit Sub
'    tv.LabelEdit = TVW_LABEL_EDIT_MANUAL
'End Sub
'
'Public Sub TVW_ConfigureTreeView(ByVal tv As Object, Optional ByVal allowDirectTextEdit As Boolean = True, Optional ByVal oleDragMode As Long = 1, Optional ByVal oleDropMode As Long = 1)
'    'Intended use: Applies the shared editing and drag/drop configuration to a TreeView control.
'    If tv Is Nothing Then Exit Sub
'    tv.oleDragMode = oleDragMode
'    tv.oleDropMode = oleDropMode
'    If allowDirectTextEdit Then
'        tv.LabelEdit = TVW_LABEL_EDIT_AUTOMATIC
'    Else
'        tv.LabelEdit = TVW_LABEL_EDIT_MANUAL
'    End If
'End Sub
'
'Public Function TVW_GetExpandedNodeKeys(ByVal tv As Object) As Object
'    'Intended use: Returns a dictionary containing the keys of currently expanded TreeView nodes.
'    Dim expandedKeys As Object
'    Dim n As Object
'    Set expandedKeys = CreateObject("Scripting.Dictionary")
'    If tv Is Nothing Then
'        Set TVW_GetExpandedNodeKeys = expandedKeys
'        Exit Function
'    End If
'    For Each n In tv.Nodes
'        If n.expanded Then
'            If Not expandedKeys.Exists(CStr(n.key)) Then expandedKeys.Add CStr(n.key), True
'        End If
'    Next n
'    Set TVW_GetExpandedNodeKeys = expandedKeys
'End Function
'
'Public Sub TVW_RestoreExpandedNodeKeys(ByVal tv As Object, ByVal expandedKeys As Object)
'    'Intended use: Re-expands TreeView nodes whose keys were captured before a refresh.
'    Dim key As Variant
'    If tv Is Nothing Then Exit Sub
'    If expandedKeys Is Nothing Then Exit Sub
'    For Each key In expandedKeys.keys
'        If TVW_NodeExists(tv, CStr(key)) Then tv.Nodes(CStr(key)).expanded = True
'    Next key
'End Sub
'
'Public Sub TVW_ExpandNodeAndParents(ByVal Node As Object)
'    'Intended use: Expands a TreeView node and each of its ancestors.
'    Dim n As Object
'    If Node Is Nothing Then Exit Sub
'    Node.expanded = True
'    Set n = Node.Parent
'    Do While Not n Is Nothing
'        n.expanded = True
'        Set n = n.Parent
'    Loop
'End Sub
'
'Public Sub TVW_SelectNode(ByVal tv As Object, ByVal nodeKey As String, Optional ByVal expandParents As Boolean = True, Optional ByVal ensureVisible As Boolean = True)
'    'Intended use: Selects a TreeView node by key and optionally expands ancestors and ensures visibility.
'    If tv Is Nothing Then Exit Sub
'    If Not TVW_NodeExists(tv, nodeKey) Then Exit Sub
'    If expandParents Then TVW_ExpandNodeAndParents tv.Nodes(nodeKey)
'    tv.Nodes(nodeKey).selected = True
'    If ensureVisible Then tv.Nodes(nodeKey).ensureVisible
'End Sub
'
'Public Function TVW_SelectedNodeKey(ByVal tv As Object) As String
'    'Intended use: Returns the selected TreeView node key, or blank when no node is selected.
'    If tv Is Nothing Then Exit Function
'    If tv.selectedItem Is Nothing Then Exit Function
'    TVW_SelectedNodeKey = CStr(tv.selectedItem.key)
'End Function
'
'Public Function TVW_SelectedParentNodeKey(ByVal tv As Object) As String
'    'Intended use: Returns the selected node parent key, or blank when no parent is available.
'    If tv Is Nothing Then Exit Function
'    If tv.selectedItem Is Nothing Then Exit Function
'    If tv.selectedItem.Parent Is Nothing Then Exit Function
'    TVW_SelectedParentNodeKey = CStr(tv.selectedItem.Parent.key)
'End Function
'
'Public Function TVW_GetCurrentTreeState(ByVal tv As Object, Optional ByVal preferredKey As String = vbNullString) As Object
'    'Intended use: Captures selected and expanded TreeView state for restoration after a reload.
'    Dim state As Object
'    Dim selectedKey As String
'    Dim fallbackKey As String
'    Set state = CreateObject("Scripting.Dictionary")
'    state.Add "ExpandedKeys", TVW_GetExpandedNodeKeys(tv)
'    If Not tv Is Nothing Then
'        If Not tv.selectedItem Is Nothing Then
'            selectedKey = CStr(tv.selectedItem.key)
'            If Not tv.selectedItem.Parent Is Nothing Then fallbackKey = CStr(tv.selectedItem.Parent.key)
'        End If
'    End If
'    If preferredKey <> vbNullString Then selectedKey = preferredKey
'    state.Add "SelectedKey", selectedKey
'    state.Add "FallbackKey", fallbackKey
'    Set TVW_GetCurrentTreeState = state
'End Function
'
'Public Function TVW_BeginRefresh(ByVal tv As Object, Optional ByVal preferredKey As String = vbNullString) As Object
'    'Intended use: Captures TreeView state before caller-owned node rebuilding begins.
'    Set TVW_BeginRefresh = TVW_GetCurrentTreeState(tv, preferredKey)
'End Function
'
'Public Sub TVW_EndRefresh(ByVal tv As Object, ByVal treeState As Object, Optional ByVal expandSelectedParents As Boolean = True, Optional ByVal ensureSelectedVisible As Boolean = True)
'    'Intended use: Restores TreeView state after caller-owned node rebuilding completes.
'    TVW_RestoreTreeState tv, treeState, expandSelectedParents, ensureSelectedVisible
'End Sub
'
'Public Sub TVW_RestoreTreeState(ByVal tv As Object, ByVal treeState As Object, Optional ByVal expandSelectedParents As Boolean = True, Optional ByVal ensureSelectedVisible As Boolean = True)
'    'Intended use: Restores selected and expanded nodes from a previously captured state object.
'    Dim selectedKey As String
'    Dim fallbackKey As String
'    If tv Is Nothing Then Exit Sub
'    If treeState Is Nothing Then Exit Sub
'    If treeState.Exists("ExpandedKeys") Then TVW_RestoreExpandedNodeKeys tv, treeState("ExpandedKeys")
'    If treeState.Exists("SelectedKey") Then selectedKey = CStr(treeState("SelectedKey") & "")
'    If treeState.Exists("FallbackKey") Then fallbackKey = CStr(treeState("FallbackKey") & "")
'    If selectedKey <> vbNullString Then
'        If TVW_NodeExists(tv, selectedKey) Then
'            TVW_SelectNode tv, selectedKey, expandSelectedParents, ensureSelectedVisible
'            Exit Sub
'        End If
'    End If
'    If fallbackKey <> vbNullString Then
'        If TVW_NodeExists(tv, fallbackKey) Then TVW_SelectNode tv, fallbackKey, True, True
'    End If
'End Sub
'
'Public Sub TVW_ReloadTreePreserveState(ByVal tv As Object, ByVal treeState As Object)
'    'Intended use: Reloads caller-populated TreeView content while preserving selection and expansion state.
'    TVW_EndRefresh tv, treeState, True, True
'End Sub
'
'Public Function TVW_IsDescendantNode(ByVal possibleAncestor As Object, ByVal possibleDescendant As Object) As Boolean
'    'Intended use: Determines whether one TreeView node is a descendant of another.
'    Dim n As Object
'    If possibleAncestor Is Nothing Then Exit Function
'    If possibleDescendant Is Nothing Then Exit Function
'    Set n = possibleDescendant.Parent
'    Do While Not n Is Nothing
'        If CStr(n.key) = CStr(possibleAncestor.key) Then
'            TVW_IsDescendantNode = True
'            Exit Function
'        End If
'        Set n = n.Parent
'    Loop
'End Function
'
'Public Function TVW_NodeTagPart(ByVal Node As Object, ByVal partIndex As Long, Optional ByVal delimiter As String = "|") As String
'    'Intended use: Returns one delimiter-separated component from a TreeView node Tag.
'    Dim parts() As String
'    If Node Is Nothing Then Exit Function
'    parts = Split(CStr(Node.Tag & ""), delimiter)
'    If partIndex < LBound(parts) Then Exit Function
'    If partIndex > UBound(parts) Then Exit Function
'    TVW_NodeTagPart = CStr(parts(partIndex) & "")
'End Function
'
'Public Function TVW_GetNodeFileKey(ByVal Node As Object) As String
'    'Intended use: Returns the file-node key represented by a TreeView node.
'    TVW_GetNodeFileKey = TVW_NodeTagPart(Node, 0, "|")
'End Function
'
'Public Function TVW_GetNodeMenuKey(ByVal Node As Object) As String
'    'Intended use: Returns the menu-node key represented by a TreeView node.
'    Dim TagValue As String
'    TagValue = TVW_NodeTagPart(Node, 1, "|")
'    If TagValue = "" Then
'        If Not Node Is Nothing Then
'            If Left$(CStr(Node.key), 2) = "M:" Then TagValue = Mid$(CStr(Node.key), 3)
'        End If
'    End If
'    TVW_GetNodeMenuKey = TagValue
'End Function
'
'Public Sub TVW_SaveTreeState(ByVal tv As Object, ByVal appName As String, Optional ByVal sectionName As String = "TreeView", Optional ByVal selectedSettingName As String = "SelectedNode", Optional ByVal expandedSettingName As String = "ExpandedNodes")
'    'Intended use: Persists selected and expanded TreeView keys in the requested application settings section.
'    Dim selectedKey As String
'    Dim expandedText As String
'    Dim n As Object
'    If Not tv Is Nothing Then
'        If Not tv.selectedItem Is Nothing Then selectedKey = CStr(tv.selectedItem.key)
'        For Each n In tv.Nodes
'            If n.expanded Then expandedText = expandedText & CStr(n.key) & ";"
'        Next n
'    End If
'    SaveSetting appName, sectionName, selectedSettingName, selectedKey
'    SaveSetting appName, sectionName, expandedSettingName, expandedText
'End Sub
'
'Public Sub TVW_RestoreSavedExpandedNodes(ByVal tv As Object, ByVal appName As String, Optional ByVal sectionName As String = "TreeView", Optional ByVal expandedSettingName As String = "ExpandedNodes")
'    'Intended use: Restores expanded TreeView nodes from persisted application settings.
'    Dim expandedText As String
'    Dim parts() As String
'    Dim i As Long
'    Dim nodeKey As String
'    If tv Is Nothing Then Exit Sub
'    expandedText = GetSetting(appName, sectionName, expandedSettingName, "")
'    If expandedText = "" Then Exit Sub
'    parts = Split(expandedText, ";")
'    For i = LBound(parts) To UBound(parts)
'        nodeKey = CStr(parts(i) & "")
'        If nodeKey <> "" Then
'            If TVW_NodeExists(tv, nodeKey) Then tv.Nodes(nodeKey).expanded = True
'        End If
'    Next i
'End Sub
'
'Public Function TVW_LoadSavedSelectedNodeKey(ByVal appName As String, Optional ByVal sectionName As String = "TreeView", Optional ByVal selectedSettingName As String = "SelectedNode") As String
'    'Intended use: Returns the selected TreeView node key stored in application settings.
'    TVW_LoadSavedSelectedNodeKey = GetSetting(appName, sectionName, selectedSettingName, "")
'End Function
'
'Public Sub TVW_RestoreSavedTreeState(ByVal tv As Object, ByVal appName As String, Optional ByVal sectionName As String = "TreeView", Optional ByVal selectedSettingName As String = "SelectedNode", Optional ByVal expandedSettingName As String = "ExpandedNodes")
'    'Intended use: Restores persisted TreeView selection and expansion state.
'    Dim selectedKey As String
'    If tv Is Nothing Then Exit Sub
'    TVW_RestoreSavedExpandedNodes tv, appName, sectionName, expandedSettingName
'    selectedKey = TVW_LoadSavedSelectedNodeKey(appName, sectionName, selectedSettingName)
'    If selectedKey <> vbNullString Then TVW_SelectNode tv, selectedKey, True, True
'End Sub
'
'Public Sub TVW_ConfigureDragDrop(ByVal tv As Object, Optional ByVal oleDragMode As Long = 1, Optional ByVal oleDropMode As Long = 1)
'    'Intended use: Applies caller-specified OLE drag and drop modes to a TreeView control.
'    If tv Is Nothing Then Exit Sub
'    tv.oleDragMode = oleDragMode
'    tv.oleDropMode = oleDropMode
'End Sub
'
'Public Function TVW_NodeID(ByVal Node As Object, Optional ByVal delimiter As String = "|", Optional ByVal partIndex As Long = -1) As Long
'    'Intended use: Returns the numeric ID component encoded in a TreeView node Tag.
'    Dim valueText As String
'    If Node Is Nothing Then Exit Function
'    If partIndex >= 0 Then
'        valueText = TVW_NodeTagPart(Node, partIndex, delimiter)
'    Else
'        valueText = CStr(Node.key & "")
'        If InStr(valueText, ":") > 0 Then valueText = Mid$(valueText, InStrRev(valueText, ":") + 1)
'        If Len(valueText) > 1 Then
'            If Not IsNumeric(valueText) Then valueText = Mid$(valueText, 2)
'        End If
'    End If
'    If IsNumeric(valueText) Then TVW_NodeID = CLng(valueText)
'End Function
'
'' DONT DELETE, MAY BE NEEDED
''Public Function GetNodeFileValue(ByVal Node As Object) As String
''    Dim parts() As String
''    parts = Split(CStr(Node.Tag & ""), "|")
''    If UBound(parts) >= 2 Then
''        GetNodeFileValue = parts(2)
''    End If
''End Function
''
''Public Function IsResourcesFileNode(ByVal Node As Object) As Boolean
''    IsResourcesFileNode = (StrComp(GetNodeFileName(Node), "resources.nss", vbTextCompare) = 0)
''End Function
''
''Public Function GetNodeFileName(ByVal Node As Object) As String
''    Dim fileMenuID As String
''    Dim fileNameValue As Variant
''    fileMenuID = GetNodeFileKey(Node)
''    If fileMenuID <> "" Then
''        fileNameValue = DLookup("[menu]", "[qMenus]", "[ID]=" & CLng(fileMenuID))
''        If Not IsNull(fileNameValue) Then
''            GetNodeFileName = CStr(fileNameValue)
''            Exit Function
''        End If
''    End If
''    GetNodeFileName = CStr(Node.text & "")
''End Function
''Public Function HasProperty(ByVal obj As Object, ByVal memberName As String) As Boolean
''    On Error Resume Next
''    CallByName obj, memberName, VbMethod, 0
''    HasProperty = (Err.Number = 0)
''    Err.Clear
''    On Error GoTo 0
''End Function
''Public Sub ApplyTreeNodeVisuals(ByVal n As Object, ByVal ParseNo As Boolean)
''    On Error Resume Next
''    If ParseNo Then
''        n.ForeColor = vbRed
''    Else
''        n.ForeColor = vbWindowText
''    End If
''    Err.Clear
''    On Error GoTo 0
''End Sub
''
''Public Function GetNodeFileKey(ByVal Node As Object) As String
''    Dim parts() As String
''    parts = Split(CStr(Node.Tag & ""), "|")
''    If UBound(parts) >= 0 Then
''        GetNodeFileKey = parts(0)
''    End If
''End Function
''
''Public Function GetNodeMenuKey(ByVal Node As Object) As String
''    Dim parts() As String
''    parts = Split(CStr(Node.Tag & ""), "|")
''    If UBound(parts) >= 1 Then
''        GetNodeMenuKey = parts(1)
''    End If
''End Function
''
''Public Function IsDescendantNode(ByVal sourceNode As Object, ByVal targetNode As Object) As Boolean
''    Dim n As Object
''    Set n = targetNode.Parent
''    Do While Not n Is Nothing
''        If n.key = sourceNode.key Then
''            IsDescendantNode = True
''            Exit Function
''        End If
''        Set n = n.Parent
''    Loop
''End Function
''
''Public Function GetExpandedNodeKeys(ByVal tv As Object) As Object
''    Dim expandedKeys As Object
''    Dim n As Object
''    Set expandedKeys = CreateObject("Scripting.Dictionary")
''    For Each n In tv.Nodes
''        If n.expanded Then
''            expandedKeys.Add n.key, True
''        End If
''    Next n
''    Set GetExpandedNodeKeys = expandedKeys
''End Function
''
''Public Sub RestoreExpandedNodeKeys(ByVal tv As Object, ByVal expandedKeys As Object)
''    Dim key As Variant
''    For Each key In expandedKeys.Keys
''        If NodeExists(tv, CStr(key)) Then
''            tv.Nodes(CStr(key)).expanded = True
''        End If
''    Next key
''End Sub
''
''Public Function NodeExists(ByVal tv As Object, ByVal nodeKey As String) As Boolean
''    Dim n As Object
''    On Error Resume Next
''    Err.Clear
''    Set n = tv.Nodes(nodeKey)
''    NodeExists = (Err.Number = 0 And Not n Is Nothing)
''    Set n = Nothing
''    Err.Clear
''    On Error GoTo 0
''End Function
'
'' ================================================================================================================================================================================
'' Sort
'' ================================================================================================================================================================================
'Public Sub SharedForm_SizeDatasheetColumnsToFieldTitles(ByVal frm As Access.Form, Optional ByVal paddingTwips As Long = 360, Optional ByVal minimumWidthTwips As Long = 420, Optional ByVal maximumWidthTwips As Long = 9000, Optional ByVal saveFormLayout As Boolean = True, Optional ByVal includeHiddenColumns As Boolean = False)
'    'Sizes the datasheet columns for an already-open form so each column is wide enough to show its field title.
'    '
'    'Intended use:
'    '  - Call from a datasheet form's Load/Open event.
'    '  - Call from a split form's Load/Open event.
'    '  - Call against a datasheet subform by passing the subform's Form object, for example:
'    '       SharedForm_SizeDatasheetColumnsToFieldTitles Me.MySubformControl.Form
'    '
'    'What this changes:
'    '  - Sets each eligible control's ColumnWidth property.
'    '  - Only affects datasheet-style column widths. It does not resize normal single-form controls.
'    '  - Saves the form layout by default so the new widths persist after the form closes.
'    '
'    'How the field title is chosen:
'    '  1. Attached label Caption, when the bound control has an attached label.
'    '  2. Control Caption property, when available.
'    '  3. DAO field Caption from the form RecordsetClone, when available.
'    '  4. ControlSource / control name fallback.
'    '
'    'Important behavior:
'    '  - This sizes columns to the title/header only, not to the longest data value.
'    '  - Hidden columns are skipped by default so a hidden field stays hidden and untouched.
'    '  - Set includeHiddenColumns=True only when you want hidden columns to receive a saved width too.
'    On Error GoTo ErrHandler
'    Dim ctl As Access.control
'    Dim rs As DAO.Recordset
'    Dim TitleText As String
'    Dim widthTwips As Long
'    Dim fontSizePoints As Long
'    Dim contextText As String
'    fontSizePoints = SharedDatasheet_FormDatasheetFontHeight(frm, 11)
'    On Error Resume Next
'    Err.Clear
'    Set rs = frm.RecordsetClone
'    Err.Clear
'    On Error GoTo ErrHandler
'    For Each ctl In frm.Controls
'        If SharedDatasheet_ControlCanHaveColumnWidth(ctl) Then
'            If includeHiddenColumns Or Not SharedDatasheet_ControlColumnHidden(ctl) Then
'                TitleText = SharedDatasheet_FormControlTitle(ctl, rs)
'                If Len(TitleText) > 0 Then
'                    widthTwips = SharedDatasheet_TitleWidthTwips(TitleText, fontSizePoints, paddingTwips, minimumWidthTwips, maximumWidthTwips)
'                    ctl.ColumnWidth = widthTwips
'                End If
'            End If
'        End If
'    Next ctl
'    If saveFormLayout Then SharedDatasheet_SaveFormLayout frm
'    If Not rs Is Nothing Then rs.Close
'    Set rs = Nothing
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim closeErrNumber As Long
'    Dim closeErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If Not rs Is Nothing Then
'        Err.Clear
'        rs.Close
'        closeErrNumber = Err.Number
'        closeErrDescription = Err.Description
'    End If
'    Set rs = Nothing
'    On Error GoTo 0
'    contextText = "Could not size form datasheet columns to field titles." & vbCrLf
'    contextText = contextText & "Form: " & SharedDatasheet_SafeFormName(frm) & vbCrLf
'    contextText = contextText & "Err.Number: " & CStr(savedErrNumber) & vbCrLf
'    contextText = contextText & "Err.Source: " & savedErrSource & vbCrLf
'    contextText = contextText & "Err.Description: " & savedErrDescription
'    If closeErrNumber <> 0 Then contextText = contextText & vbCrLf & "Recordset-close error " & CStr(closeErrNumber) & ": " & closeErrDescription
'    Err.Raise savedErrNumber, "SharedForm_SizeDatasheetColumnsToFieldTitles", contextText
'End Sub
'
'Public Sub SharedTable_SizeDatasheetColumnsToFieldTitles(ByVal tableName As String, Optional ByVal paddingTwips As Long = 360, Optional ByVal minimumWidthTwips As Long = 420, Optional ByVal maximumWidthTwips As Long = 9000, Optional ByVal fontSizePoints As Long = 11)
'    'Sizes saved datasheet column widths for a table so each field title is fully visible.
'    'Intended use:
'    '  - Call once after creating/repairing a table.
'    '  - Call manually when table field captions have changed.
'    '  - Useful for linked tables too, as long as Access allows the local field display property to be saved.
'    'What this changes:
'    '  - Sets or creates each DAO Field object's ColumnWidth property.
'    '  - Does not open the table.
'    '  - Does not alter field definitions, data, indexes, relationships, or captions.
'    'How the field title is chosen:
'    '  1. Field Caption property, when present.
'    '  2. Actual field name fallback.
'    'Important behavior:
'    '  - If the table is already open, Access may not visually refresh the widths until the table is closed and reopened.
'    '  - This sizes to the title/header only, not to data values.
'    On Error GoTo ErrHandler
'    Dim db As DAO.Database
'    Dim tdf As DAO.tableDef
'    Dim fld As DAO.Field
'    Dim TitleText As String
'    Dim widthTwips As Long
'    Dim contextText As String
'    Set db = currentDb
'    If Not SharedDao_TableExists(db, tableName) Then Err.Raise 3265, "SharedTable_SizeDatasheetColumnsToFieldTitles", "Table was not found: " & tableName
'    Set tdf = db.TableDefs(tableName)
'    For Each fld In tdf.fields
'        TitleText = SharedDatasheet_DaoFieldTitle(fld)
'        If Len(TitleText) > 0 Then
'            widthTwips = SharedDatasheet_TitleWidthTwips(TitleText, fontSizePoints, paddingTwips, minimumWidthTwips, maximumWidthTwips)
'            SharedDatasheet_SetDaoFieldColumnWidth fld, widthTwips
'        End If
'    Next fld
'    db.TableDefs.Refresh
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    contextText = "Could not size table datasheet columns to field titles." & vbCrLf
'    contextText = contextText & "Table: " & tableName & vbCrLf
'    contextText = contextText & "Err.Number: " & CStr(savedErrNumber) & vbCrLf
'    contextText = contextText & "savedErrSource: " & savedErrSource & vbCrLf
'    contextText = contextText & "savedErrDescription: " & savedErrDescription
'    Err.Raise savedErrNumber, "SharedTable_SizeDatasheetColumnsToFieldTitles", contextText
'End Sub
'
'Public Sub SharedQuery_SizeDatasheetColumnsToFieldTitles(ByVal queryName As String, Optional ByVal paddingTwips As Long = 360, Optional ByVal minimumWidthTwips As Long = 420, Optional ByVal maximumWidthTwips As Long = 9000, Optional ByVal fontSizePoints As Long = 11)
'    'Sizes saved datasheet column widths for a saved query so each output field title is fully visible.
'    'Intended use:
'    '  - Call after creating or modifying a saved query.
'    '  - Call manually when query output aliases/captions have changed.
'    'What this changes:
'    '  - Sets or creates each query output Field object's ColumnWidth property.
'    '  - Does not change the SQL text.
'    '  - Does not open the query.
'    'How the field title is chosen:
'    '  1. Query output field Caption property, when present.
'    '  2. Query output field name/alias fallback.
'    'Important behavior:
'    '  - If the query is already open, Access may not visually refresh the widths until the query is closed and reopened.
'    '  - Parameter queries may fail if Access cannot resolve the output Fields collection without parameter values.
'    '  - This sizes to the title/header only, not to data values.
'    On Error GoTo ErrHandler
'    Dim db As DAO.Database
'    Dim qdf As DAO.queryDef
'    Dim fld As DAO.Field
'    Dim TitleText As String
'    Dim widthTwips As Long
'    Dim contextText As String
'    Set db = currentDb
'    If Not SharedDao_QueryExists(db, queryName) Then Err.Raise 3265, "SharedQuery_SizeDatasheetColumnsToFieldTitles", "Query was not found: " & queryName
'    Set qdf = db.QueryDefs(queryName)
'    qdf.fields.Refresh
'    For Each fld In qdf.fields
'        TitleText = SharedDatasheet_DaoFieldTitle(fld)
'        If Len(TitleText) > 0 Then
'            widthTwips = SharedDatasheet_TitleWidthTwips(TitleText, fontSizePoints, paddingTwips, minimumWidthTwips, maximumWidthTwips)
'            SharedDatasheet_SetDaoFieldColumnWidth fld, widthTwips
'        End If
'    Next fld
'    db.QueryDefs.Refresh
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    contextText = "Could not size query datasheet columns to field titles." & vbCrLf
'    contextText = contextText & "Query: " & queryName & vbCrLf
'    contextText = contextText & "Err.Number: " & CStr(savedErrNumber) & vbCrLf
'    contextText = contextText & "savedErrSource: " & savedErrSource & vbCrLf
'    contextText = contextText & "savedErrDescription: " & savedErrDescription
'    Err.Raise savedErrNumber, "SharedQuery_SizeDatasheetColumnsToFieldTitles", contextText
'End Sub
'
'Private Function SharedDatasheet_FormControlTitle(ByVal ctl As Access.control, ByVal rs As DAO.Recordset) As String
'    'Returns the best available title for one form datasheet column.
'    'Why this helper exists:
'    '  - Form datasheet headings can come from different places depending on how the form/control was built.
'    '  - Some generated controls have attached labels.
'    '  - Some controls rely on the bound DAO field Caption.
'    '  - Some controls have only a ControlSource or control name.
'    Dim TitleText As String
'    Dim fieldName As String
'    TitleText = SharedDatasheet_AttachedLabelCaption(ctl)
'    If Len(TitleText) = 0 Then TitleText = SharedDatasheet_PropertyText(ctl, "Caption")
'    fieldName = SharedDatasheet_ControlFieldName(ctl)
'    If Len(TitleText) = 0 And Len(fieldName) > 0 Then
'        If Not rs Is Nothing Then
'            If SharedDao_RecordsetHasFieldSafe(rs, fieldName) Then TitleText = SharedDatasheet_DaoFieldTitle(rs.fields(fieldName))
'        End If
'    End If
'    If Len(TitleText) = 0 Then TitleText = fieldName
'    If Len(TitleText) = 0 Then TitleText = SharedDatasheet_SafeControlName(ctl)
'    SharedDatasheet_FormControlTitle = Trim$(TitleText)
'End Function
'
'Private Function SharedDatasheet_AttachedLabelCaption(ByVal ctl As Access.control) As String
'    'Returns the caption from the first label attached to a form control.
'    'Why errors are suppressed:
'    '  - Not every Access control exposes a child Controls collection.
'    '  - Some controls can have ColumnWidth but no attached label.
'    '  - Missing label support should not stop sizing the rest of the columns.
'    On Error GoTo ErrHandler
'    Dim childCtl As Access.control
'    For Each childCtl In ctl.Controls
'        If childCtl.controlType = acLabel Then
'            SharedDatasheet_AttachedLabelCaption = Trim$(CStr(Nz(childCtl.Caption, vbNullString)))
'            Exit Function
'        End If
'    Next childCtl
'    Exit Function
'ErrHandler:
'    SharedDatasheet_AttachedLabelCaption = vbNullString
'End Function
'
'Private Function SharedDatasheet_DaoFieldTitle(ByVal fld As DAO.Field) As String
'    'Returns the user-facing title for a DAO field.
'    'DAO field display rule:
'    '  - Access stores the optional field heading in the Caption property.
'    '  - If Caption is missing or blank, Access falls back to the field name.
'    Dim TitleText As String
'    TitleText = SharedDatasheet_PropertyText(fld, "Caption")
'    If Len(TitleText) = 0 Then TitleText = fld.Name
'    SharedDatasheet_DaoFieldTitle = Trim$(TitleText)
'End Function
'
'Private Sub SharedDatasheet_SetDaoFieldColumnWidth(ByVal fld As DAO.Field, ByVal widthTwips As Long)
'    'Sets and verifies the saved datasheet ColumnWidth for one DAO field, creating the optional property only when DAO reports that it is missing.
'    On Error GoTo ErrHandler
'    Dim assignmentErrNumber As Long
'    Dim assignmentErrSource As String
'    Dim assignmentErrDescription As String
'    Dim fieldName As String
'    Dim requestedWidth As Integer
'    If fld Is Nothing Then Err.Raise vbObjectError + 7958, "SharedDatasheet_SetDaoFieldColumnWidth", "The DAO.Field argument is Nothing."
'    fieldName = fld.Name
'    requestedWidth = CInt(widthTwips)
'    On Error Resume Next
'    Err.Clear
'    fld.Properties("ColumnWidth").value = requestedWidth
'    assignmentErrNumber = Err.Number
'    assignmentErrSource = Err.Source
'    assignmentErrDescription = Err.Description
'    Err.Clear
'    On Error GoTo ErrHandler
'    If assignmentErrNumber = SHARED_DAO_ERROR_PROPERTY_NOT_FOUND Then
'        fld.Properties.Append fld.CreateProperty("ColumnWidth", dbInteger, requestedWidth)
'    ElseIf assignmentErrNumber <> 0 Then
'        Err.Raise assignmentErrNumber, assignmentErrSource, assignmentErrDescription
'    End If
'    If CLng(fld.Properties("ColumnWidth").value) <> CLng(requestedWidth) Then Err.Raise vbObjectError + 7959, "SharedDatasheet_SetDaoFieldColumnWidth", "The ColumnWidth property does not match the requested value after assignment."
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedDatasheet_SetDaoFieldColumnWidth", "Unable to set and verify the DAO field ColumnWidth property." & vbCrLf & "Field: " & fieldName & vbCrLf & "Requested width: " & CStr(widthTwips) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Original description: " & savedErrDescription
'End Sub
'
'Private Function SharedDatasheet_TitleWidthTwips(ByVal TitleText As String, ByVal fontSizePoints As Long, ByVal paddingTwips As Long, ByVal minimumWidthTwips As Long, ByVal maximumWidthTwips As Long) As Long
'    'Estimates the column width needed to display a title/header.
'    'Why this is estimated instead of using Access auto-fit:
'    '  - Access can auto-fit to data values, but the requested behavior is title-only sizing.
'    '  - Title-only sizing keeps columns compact when data values are much longer than the field title.
'    'Padding/min/max behavior:
'    '  - paddingTwips leaves breathing room after the title text.
'    '  - minimumWidthTwips prevents very short titles like "ID" from becoming unusably narrow.
'    '  - maximumWidthTwips prevents long titles from creating oversized datasheet columns.
'    Dim i As Long
'    Dim cleanTitle As String
'    Dim widthTwips As Long
'    cleanTitle = Replace(TitleText, vbCr, " ")
'    cleanTitle = Replace(cleanTitle, vbLf, " ")
'    cleanTitle = Replace(cleanTitle, vbTab, " ")
'    If fontSizePoints <= 0 Then fontSizePoints = 11
'    For i = 1 To Len(cleanTitle)
'        widthTwips = widthTwips + SharedDatasheet_TitleCharacterWidthTwips(Mid$(cleanTitle, i, 1))
'    Next i
'    widthTwips = ((widthTwips * fontSizePoints) \ 11) + paddingTwips
'    If widthTwips < minimumWidthTwips Then widthTwips = minimumWidthTwips
'    If maximumWidthTwips > 0 And widthTwips > maximumWidthTwips Then widthTwips = maximumWidthTwips
'    If widthTwips > 30000 Then widthTwips = 30000
'    SharedDatasheet_TitleWidthTwips = widthTwips
'End Function
'
'Private Function SharedDatasheet_TitleCharacterWidthTwips(ByVal ch As String) As Long
'    'Returns an approximate 11-point datasheet-header character width in twips.
'    'Why characters are weighted:
'    '  - A title like "iiii" should not receive the same width as "WWWW".
'    '  - This does not need print-perfect measurement; it only needs a practical datasheet column width.
'    Select Case ch
'        Case "i", "l", "I", "j", "t", "f", ".", ",", "'", "`", ":", ";", "!", "|"
'            SharedDatasheet_TitleCharacterWidthTwips = 55
'        Case " ", "-", "_", "(", ")", "[", "]", "{", "}", "/", "\"
'            SharedDatasheet_TitleCharacterWidthTwips = 75
'        Case "m", "w", "M", "W", "@", "%", "&", "#"
'            SharedDatasheet_TitleCharacterWidthTwips = 145
'        Case "A" To "Z"
'            SharedDatasheet_TitleCharacterWidthTwips = 115
'        Case Else
'            SharedDatasheet_TitleCharacterWidthTwips = 95
'    End Select
'End Function
'
'Private Function SharedDatasheet_ControlCanHaveColumnWidth(ByVal ctl As Access.control) As Boolean
'    'Returns True for normal bound controls that appear as editable/readable datasheet columns.
'    'Included:
'    '  - Text boxes for text/number/date/etc. fields.
'    '  - Combo boxes for lookup-style fields.
'    '  - Check boxes for Yes/No fields.
'    'Excluded:
'    '  - Labels, buttons, subform controls, tab controls, option groups, images, and other non-field layout controls.
'    Select Case ctl.controlType
'        Case acTextBox, acComboBox, acCheckBox
'            SharedDatasheet_ControlCanHaveColumnWidth = True
'    End Select
'End Function
'
'Private Function SharedDatasheet_ControlColumnHidden(ByVal ctl As Access.control) As Boolean
'    'Safely returns whether a datasheet column is hidden.
'    'Unsupported controls are treated as not hidden because this helper is only used after control-type filtering.
'    On Error GoTo ErrHandler
'    SharedDatasheet_ControlColumnHidden = CBool(ctl.ColumnHidden)
'    Exit Function
'ErrHandler:
'    SharedDatasheet_ControlColumnHidden = False
'End Function
'
'Private Function SharedDatasheet_ControlFieldName(ByVal ctl As Access.control) As String
'    'Returns the bound DAO field name for a form control.
'    'Rules:
'    '  - Blank ControlSource means the control is unbound.
'    '  - ControlSource starting with "=" means the control is calculated, not a direct DAO field.
'    '  - Bracketed field names like [FieldName] are normalized to FieldName.
'    'This helper intentionally does not try to parse expressions.
'    Dim sourceText As String
'    sourceText = Trim$(SharedDatasheet_PropertyText(ctl, "ControlSource"))
'    If Len(sourceText) = 0 Then Exit Function
'    If Left$(sourceText, 1) = "=" Then Exit Function
'    If Left$(sourceText, 1) = "[" And Right$(sourceText, 1) = "]" Then sourceText = Mid$(sourceText, 2, Len(sourceText) - 2)
'    SharedDatasheet_ControlFieldName = sourceText
'End Function
'
'Private Function SharedDatasheet_PropertyText(ByVal targetObject As Object, ByVal propertyName As String) As String
'    'Reads a property from an Access object or DAO object and returns it as trimmed text.
'    'Why missing properties return blank:
'    '  - Caption, ColumnWidth, and similar display properties are optional in several Access/DAO objects.
'    '  - The caller usually has a fallback source, so missing property errors should not stop the process.
'    On Error GoTo ErrHandler
'    SharedDatasheet_PropertyText = Trim$(CStr(Nz(targetObject.Properties(propertyName).value, vbNullString)))
'    Exit Function
'ErrHandler:
'    SharedDatasheet_PropertyText = vbNullString
'End Function
'
'Private Function SharedDatasheet_FormDatasheetFontHeight(ByVal frm As Access.Form, ByVal defaultFontSizePoints As Long) As Long
'    'Returns the datasheet font size used for width estimation.
'    'Access exposes the saved datasheet font size through the form's DatasheetFontHeight property when available.
'    'If Access does not expose it for the form/view, this helper returns the supplied default.
'    On Error GoTo ErrHandler
'    SharedDatasheet_FormDatasheetFontHeight = CLng(frm.Properties("DatasheetFontHeight"))
'    If SharedDatasheet_FormDatasheetFontHeight <= 0 Then SharedDatasheet_FormDatasheetFontHeight = defaultFontSizePoints
'    Exit Function
'ErrHandler:
'    SharedDatasheet_FormDatasheetFontHeight = defaultFontSizePoints
'End Function
'
'Private Sub SharedDatasheet_SaveFormLayout(ByVal frm As Access.Form)
'    'Saves runtime datasheet ColumnWidth changes and raises with both direct-form and host-form evidence when persistence fails.
'    On Error GoTo ErrHandler
'    Dim formName As String
'    Dim hostFormName As String
'    Dim directErrNumber As Long
'    Dim directErrSource As String
'    Dim directErrDescription As String
'    Dim hostErrNumber As Long
'    Dim hostErrSource As String
'    Dim hostErrDescription As String
'    formName = frm.Name
'    On Error Resume Next
'    Err.Clear
'    DoCmd.Save acForm, formName
'    directErrNumber = Err.Number
'    directErrSource = Err.Source
'    directErrDescription = Err.Description
'    Err.Clear
'    On Error GoTo ErrHandler
'    If directErrNumber = 0 Then Exit Sub
'    On Error Resume Next
'    Err.Clear
'    hostFormName = SharedForm_TopLevelHostFormName(frm)
'    hostErrNumber = Err.Number
'    hostErrSource = Err.Source
'    hostErrDescription = Err.Description
'    Err.Clear
'    If hostErrNumber = 0 Then
'        If Len(hostFormName) = 0 Then
'            hostErrNumber = vbObjectError + 7324
'            hostErrSource = "SharedDatasheet_SaveFormLayout"
'            hostErrDescription = "The top-level host form name is blank."
'        ElseIf StrComp(hostFormName, formName, vbTextCompare) = 0 Then
'            hostErrNumber = directErrNumber
'            hostErrSource = directErrSource
'            hostErrDescription = "No different top-level host form is available. " & directErrDescription
'        Else
'            DoCmd.Save acForm, hostFormName
'            hostErrNumber = Err.Number
'            hostErrSource = Err.Source
'            hostErrDescription = Err.Description
'        End If
'    End If
'    Err.Clear
'    On Error GoTo ErrHandler
'    If hostErrNumber = 0 Then Exit Sub
'    Err.Raise vbObjectError + 7960, "SharedDatasheet_SaveFormLayout", "The datasheet layout could not be persisted." & vbCrLf & "Form: " & formName & vbCrLf & "Top-level host form: " & hostFormName & vbCrLf & "Direct-save error " & CStr(directErrNumber) & " from " & directErrSource & ": " & directErrDescription & vbCrLf & "Host-save error " & CStr(hostErrNumber) & " from " & hostErrSource & ": " & hostErrDescription
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Err.Raise savedErrNumber, "SharedDatasheet_SaveFormLayout", "Unable to save the form datasheet layout." & vbCrLf & "Form: " & formName & vbCrLf & "Top-level host form: " & hostFormName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Original description: " & savedErrDescription
'End Sub
'
'Private Function SharedDatasheet_SafeFormName(ByVal frm As Access.Form) As String
'    'Returns a form name for error messages without allowing diagnostic text building to raise another error.
'    On Error GoTo ErrHandler
'    SharedDatasheet_SafeFormName = frm.Name
'    Exit Function
'ErrHandler:
'    SharedDatasheet_SafeFormName = "(form name unavailable)"
'End Function
'
'Private Function SharedDatasheet_SafeControlName(ByVal ctl As Access.control) As String
'    'Returns a control name for fallback titles and error messages.
'    '
'    'If the name is unavailable, blank is safer than failing inside a fallback helper.
'    On Error GoTo ErrHandler
'    SharedDatasheet_SafeControlName = ctl.Name
'    Exit Function
'ErrHandler:
'    SharedDatasheet_SafeControlName = vbNullString
'End Function
'
'' ================================================================================================================================================================================
'' =============================================================================
'' Access development-option helpers
'' Purpose: These procedures help verify and restore important Access/VBA development settings after an Office reinstall, profile reset, or machine change.
'' Main settings covered:
''   1. Name AutoCorrect settings
''      These are database-specific. Run these procedures in each ACCDB you care about. For development work where objects are copied/renamed manually, all three Name AutoCorrect settings should normally be False.
''   2. VBA error trapping
''      "Break on Unhandled Errors" is normally the safest development default. Access stores that setting as numeric value 2.
''   3. Confirmation prompts
''      These help prevent accidental destructive edits, deletes, and action query execution.
'' Notes:
''   - Trusted Locations are not handled here because they are Trust Center/Office profile settings, not normal per-database Access options.
''   - After changing "Track Name AutoCorrect Info", close and reopen thedatabase so Access can fully drop/rebuild its internal name-map state.
'' =============================================================================
'
'' -------------------------------------------------------------------------
'' PrintAccessDevelopmentOptions
'' Public diagnostic procedure.
'' Use this when:
''   - You want to verify the current database's Name AutoCorrect settings.
''   - You want to confirm the current VBA error trapping mode.
''   - You want to confirm confirmation prompts are enabled.
'' Output: Prints the current option values to the Immediate Window.
'' Expected preferred values for this user's development workflow:
''   Track Name AutoCorrect Info  = False
''   Perform Name AutoCorrect     = False
''   Log Name AutoCorrect Changes = False
''   Error Trapping               = 2
''   Confirm Record Changes       = True
''   Confirm Document Deletions   = True
''   Confirm Action Queries       = True
'' -------------------------------------------------------------------------
'Public Sub PrintAccessDevelopmentOptions()
'    'Intended use: Prints the current database development-option state and identifies the exact option that could not be read.
'    On Error GoTo ErrHandler
'    Dim currentStep As String
'    currentStep = "Print report heading"
'    Debug.Print String$(80, "-")
'    Debug.Print "Access development options"
'    Debug.Print "Database: " & CurrentProject.FullName
'    Debug.Print "Printed:  " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
'    Debug.Print String$(80, "-")
'    Debug.Print "Name AutoCorrect options:"
'    currentStep = "Read Track Name AutoCorrect Info"
'    Debug.Print "  Track Name AutoCorrect Info:  "; Application.GetOption("Track Name AutoCorrect Info")
'    currentStep = "Read Perform Name AutoCorrect"
'    Debug.Print "  Perform Name AutoCorrect:     "; Application.GetOption("Perform Name AutoCorrect")
'    currentStep = "Read Log Name AutoCorrect Changes"
'    Debug.Print "  Log Name AutoCorrect Changes: "; Application.GetOption("Log Name AutoCorrect Changes")
'    Debug.Print vbNullString
'    Debug.Print "VBA error trapping:"
'    currentStep = "Read Error Trapping"
'    Debug.Print "  Error Trapping:               "; Application.GetOption("Error Trapping"); "  (2 = Break on Unhandled Errors)"
'    Debug.Print vbNullString
'    Debug.Print "Confirmation prompts:"
'    currentStep = "Read Confirm Record Changes"
'    Debug.Print "  Confirm Record Changes:       "; Application.GetOption("Confirm Record Changes")
'    currentStep = "Read Confirm Document Deletions"
'    Debug.Print "  Confirm Document Deletions:   "; Application.GetOption("Confirm Document Deletions")
'    currentStep = "Read Confirm Action Queries"
'    Debug.Print "  Confirm Action Queries:       "; Application.GetOption("Confirm Action Queries")
'    Debug.Print String$(80, "-")
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    Debug.Print "PrintAccessDevelopmentOptions failed."
'    Debug.Print "Step: " & currentStep
'    Debug.Print "Err.Number: " & CStr(savedErrNumber)
'    Debug.Print "Err.Source: " & savedErrSource
'    Debug.Print "Err.Description: " & savedErrDescription
'    Debug.Print String$(80, "-")
'    Err.Raise savedErrNumber, "PrintAccessDevelopmentOptions", "Unable to print all Access development options." & vbCrLf & "Step: " & currentStep & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Sub
'
'' -------------------------------------------------------------------------
'' ApplyAccessDevelopmentOptions
'' Public repair/apply procedure.
'' Use this when:
''   - Office was reinstalled.
''   - Access was replaced with another channel/version, such as LTSC.
''   - A project database appears to have Name AutoCorrect enabled again.
''   - You want to restore the preferred development defaults in the currently open database.
'' Important:
''   Name AutoCorrect options are database-specific. This procedure changes the options for the currently open database, not every ACCDB on disk.
'' Recommended follow-up:
''   After running this procedure, close and reopen the database. Then run PrintAccessDevelopmentOptions again to verify the final state.
'' -------------------------------------------------------------------------
'Public Sub ApplyAccessDevelopmentOptions()
'    'Intended use: Applies and verifies preferred Access development options, then fully validates restoration or reports that application state is uncertain after rollback failure.
'    On Error GoTo ErrHandler
'    Dim optionNames(0 To 6) As String
'    Dim requestedValues(0 To 6) As Variant
'    Dim originalValues(0 To 6) As Variant
'    Dim currentStep As String
'    Dim i As Long
'    Dim mutationStarted As Boolean
'    optionNames(0) = "Track Name AutoCorrect Info": requestedValues(0) = False
'    optionNames(1) = "Perform Name AutoCorrect": requestedValues(1) = False
'    optionNames(2) = "Log Name AutoCorrect Changes": requestedValues(2) = False
'    optionNames(3) = "Error Trapping": requestedValues(3) = 2
'    optionNames(4) = "Confirm Record Changes": requestedValues(4) = True
'    optionNames(5) = "Confirm Document Deletions": requestedValues(5) = True
'    optionNames(6) = "Confirm Action Queries": requestedValues(6) = True
'    currentStep = "capturing current Access option values"
'    For i = 0 To 6
'        originalValues(i) = Application.GetOption(optionNames(i))
'    Next i
'    mutationStarted = True
'    For i = 0 To 6
'        currentStep = "setting and verifying " & optionNames(i)
'        Application.SetOption optionNames(i), requestedValues(i)
'        If CStr(Application.GetOption(optionNames(i))) <> CStr(requestedValues(i)) Then Err.Raise vbObjectError + 8421, "ApplyAccessDevelopmentOptions", "The Access option did not match after assignment: " & optionNames(i)
'    Next i
'    currentStep = "performing final Access option validation"
'    For i = 0 To 6
'        If CStr(Application.GetOption(optionNames(i))) <> CStr(requestedValues(i)) Then Err.Raise vbObjectError + 8433, "ApplyAccessDevelopmentOptions", "An Access option changed before final validation completed: " & optionNames(i)
'    Next i
'    Debug.Print String$(80, "-")
'    Debug.Print "Applied and verified Access development options for:"
'    Debug.Print CurrentProject.FullName
'    Debug.Print "Applied: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
'    Debug.Print "Close and reopen this database, then run PrintAccessDevelopmentOptions."
'    Debug.Print String$(80, "-")
'    Exit Sub
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    Dim restoreErrNumber As Long
'    Dim restoreErrDescription As String
'    Dim validationErrNumber As Long
'    Dim validationErrDescription As String
'    Dim restoreOperationText As String
'    Dim finalValidationText As String
'    Dim rollbackText As String
'    Dim rollbackVerified As Boolean
'    Dim currentValue As Variant
'    Dim outcomeText As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    If mutationStarted Then
'        For i = 0 To 6
'            restoreErrNumber = 0
'            restoreErrDescription = vbNullString
'            On Error Resume Next
'            Err.Clear
'            Application.SetOption optionNames(i), originalValues(i)
'            restoreErrNumber = Err.Number
'            restoreErrDescription = Err.Description
'            On Error GoTo 0
'            If restoreErrNumber <> 0 Then restoreOperationText = SharedText_AppendNonBlankLine(restoreOperationText, optionNames(i) & " restore command error " & CStr(restoreErrNumber) & ": " & restoreErrDescription)
'        Next i
'        For i = 0 To 6
'            validationErrNumber = 0
'            validationErrDescription = vbNullString
'            On Error Resume Next
'            Err.Clear
'            currentValue = Application.GetOption(optionNames(i))
'            validationErrNumber = Err.Number
'            validationErrDescription = Err.Description
'            On Error GoTo 0
'            If validationErrNumber <> 0 Then
'                finalValidationText = SharedText_AppendNonBlankLine(finalValidationText, optionNames(i) & " final validation error " & CStr(validationErrNumber) & ": " & validationErrDescription)
'            ElseIf CStr(currentValue) <> CStr(originalValues(i)) Then
'                finalValidationText = SharedText_AppendNonBlankLine(finalValidationText, optionNames(i) & " final value does not match the captured original value.")
'            End If
'        Next i
'        rollbackVerified = (Len(finalValidationText) = 0)
'    End If
'    If mutationStarted Then
'        If rollbackVerified Then
'            outcomeText = "The requested option update failed, but the complete original Access option state was restored and verified by a final full reread."
'        Else
'            outcomeText = "The requested option update failed and final rollback validation did not complete successfully. The current Access option state is uncertain."
'        End If
'    Else
'        outcomeText = "The requested option update failed before any option mutation began."
'    End If
'    If Len(restoreOperationText) > 0 Then rollbackText = SharedText_AppendNonBlankLine(rollbackText, "Restore command diagnostics:" & vbCrLf & restoreOperationText)
'    If Len(finalValidationText) > 0 Then rollbackText = SharedText_AppendNonBlankLine(rollbackText, "Final validation failures:" & vbCrLf & finalValidationText)
'    If mutationStarted And rollbackVerified Then rollbackText = SharedText_AppendNonBlankLine(rollbackText, "Final result: every captured original option value was reread and verified.")
'    Debug.Print "ApplyAccessDevelopmentOptions failed."
'    Debug.Print "Step: " & currentStep
'    Debug.Print "Err.Number: " & CStr(savedErrNumber)
'    Debug.Print "Err.Source: " & savedErrSource
'    Debug.Print "Err.Description: " & savedErrDescription
'    If Len(rollbackText) > 0 Then Debug.Print "Rollback: " & rollbackText
'    Debug.Print String$(80, "-")
'    Err.Raise savedErrNumber, "ApplyAccessDevelopmentOptions", outcomeText & vbCrLf & "Failed step: " & currentStep & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & IIf(Len(rollbackText) > 0, vbCrLf & "Rollback: " & rollbackText, vbNullString)
'End Sub
'' ================================================================================================================================================================================
''
'' ================================================================================================================================================================================
'
'Public Function ReadTextFileUtf8(ByVal filePath As String) As String
'    'Reads one UTF-8 text file and preserves the original read error through stream cleanup.
'    On Error GoTo ErrHandler
'    Dim stream As Object
'    Set stream = CreateObject("ADODB.Stream")
'    With stream
'        .Type = 2
'        .charset = "utf-8"
'        .Open
'        .LoadFromFile filePath
'        ReadTextFileUtf8 = .ReadText(-1)
'        .Close
'    End With
'    Set stream = Nothing
'    Exit Function
'ErrHandler:
'    Dim savedErrNumber As Long
'    Dim savedErrSource As String
'    Dim savedErrDescription As String
'    savedErrNumber = Err.Number
'    savedErrSource = Err.Source
'    savedErrDescription = Err.Description
'    On Error Resume Next
'    If Not stream Is Nothing Then stream.Close
'    Set stream = Nothing
'    On Error GoTo 0
'    Err.Raise savedErrNumber, "ReadTextFileUtf8", "Unable to read the UTF-8 text file." & vbCrLf & "File: " & filePath & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
'End Function
'
'Public Function OpenForm(ByVal formName As String) As Boolean
'    'Opens the specified form in its default view.
'    On Error GoTo ErrHandler
'    If Len(formName) = 0 Then Err.Raise vbObjectError + 1, "OpenForm", "The form name is blank."
'    DoCmd.OpenForm formName, acNormal
'    OpenForm = True
'    Exit Function
'ErrHandler:
'    Debug.Print "OpenForm failed." & vbCrLf & _
'                "Form: " & formName & vbCrLf & _
'                "Error " & CStr(Err.Number) & ": " & Err.Description
'    MsgBox "Could not open the form '" & formName & "'." & vbCrLf & vbCrLf & _
'           "Error " & CStr(Err.Number) & ": " & Err.Description, _
'           vbExclamation, "Open Form"
'End Function
'
'Public Function OpenFormAndCloseCurrent(ByVal formName As String) As Boolean
'    'Closes the current form, then opens the specified form in its default view.
'    On Error GoTo ErrHandler
'    Dim currentFormName As String
'    If Len(Trim$(formName)) = 0 Then
'        Err.Raise vbObjectError + 1, _
'                  "OpenFormAndCloseCurrent", _
'                  "The form name is blank."
'    End If
'    'Capture the current form name before closing it.
'    currentFormName = Screen.ActiveForm.Name
'    'Do not close and reopen the same form.
'    If StrComp(currentFormName, formName, vbTextCompare) = 0 Then
'        OpenFormAndCloseCurrent = True
'        Exit Function
'    End If
'    'Close the current form before opening the destination form.
'    DoCmd.Close acForm, currentFormName, acSaveNo
'    DoCmd.OpenForm formName, acNormal
'    OpenFormAndCloseCurrent = True
'    Exit Function
'ErrHandler:
'    Debug.Print "OpenFormAndCloseCurrent failed." & vbCrLf & _
'                "Current form: " & currentFormName & vbCrLf & _
'                "Destination form: " & formName & vbCrLf & _
'                "Error " & CStr(Err.Number) & ": " & Err.Description
'    MsgBox "Could not close the current form and open the form '" & _
'           formName & "'." & vbCrLf & vbCrLf & _
'           "Error " & CStr(Err.Number) & ": " & Err.Description, _
'           vbExclamation, _
'           "Close Current Form and Open Form"
'End Function
'
'Public Function NzDateOrZero(ByVal value As Variant) As Date
'    If IsDate(value) Then
'        NzDateOrZero = CDate(value)
'    Else
'        NzDateOrZero = 0
'    End If
'End Function
'
'' ================================================================================================================================================================================
'' ADC TEMP SORT LATER
'' ================================================================================================================================================================================
'
''Intended use: Normalizes repeated spaces
'Public Function Shared_String_CompactSpaces(ByVal value As String) As String
'    Dim resultText As String
'    resultText = Trim$(value)
'    Do While InStr(1, resultText, "  ", vbBinaryCompare) > 0
'        resultText = Replace(resultText, "  ", " ")
'    Loop
'    Shared_String_CompactSpaces = resultText
'End Function
'
''Consolidated old names: local TV NSS EXPORT MODULE NullIfBlank.
''Returns Null when a value is Null or trims to blank; otherwise returns the original value.
'Public Function Shared_String_ConvertToNullIfBlank(ByVal v As Variant) As Variant
'    Shared_String_ConvertToNullIfBlank = SharedValue_NormalizeAssignment(v)
'End Function
