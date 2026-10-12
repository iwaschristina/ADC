Attribute VB_Name = "x__UTIL_ADC_RENAME_PREVIEW_TOOL"
' ================================================================================================================================================================================
' ChatGPT Rewrite Tracking Header
' ================================================================================================================================================================================
' ChatGPT Previous Version ID: CGPT-HMNN-CMEN-A4C9-3NC5-79CU
' ChatGPT Baseline Version ID: CGPT-HMNN-CMEN-A4C9-3NC5-79CU
' ChatGPT Version ID: CGPT-VBV5-Q9TW-A8MT-2XAU-ESQB
' ChatGPT Version Timestamp: Friday 9/4/26 9:43 AM
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
' STANDARD MODULE:
' VB_Name: x__UTIL_ADC_RENAME_PREVIEW_TOOL
' Export Group:
' Entry Points:
'     Run_PreviewProposedObjectNames() -> Variant
'     Run_ApplyProposedObjectNames() -> Variant
' SHARED CODE: None
' CLASSES USED: None
' DYNAMIC STRING REFERENCES: Event property string reference: ctl.OnClick = "=RenamePreview_RequestCancel()"
' LastScanTimestamp: Tuesday 9/1/26 1:08 PM
' Purpose:
' Author:
' Support:
' Code Category:
' Requirements:
' Reference Dependencies:
' DB Object Dependencies:
' Usage:
' Version:
' Notes:
' Core Rules Enforced Here:
' Design Map:
' Configuration:
' Behavior:
' Code Status:
' Todo:
' Reporting Bugs:
' Module Level Bugs:
' Public Subs:
'     ApplyProposedObjectNames(Optional previewOnly As Boolean = False) -> None
' Public Functions:
'     Run_PreviewProposedObjectNames() -> Variant
'     Run_ApplyProposedObjectNames() -> Variant
'     RenamePreview_RequestCancel() -> Boolean
' Public Properties: None
' ================================================================================================================================================================================

Option Compare Database
Option Explicit

' Standalone post-setup rename preview/live utility.
' The module intentionally sits outside the x__UTIL_ADC_ baseline namespace so Module Guard does not treat it as a baseline module.
' It reads the Rename workbook through the Access Database Engine, previews or renames mapped table/query fields and existing top-level Access objects, and does not alter ADC setup code.
' Field ownership is resolved by Parent ID. Query output aliases are applied before table fields, and top-level objects are renamed last.
' Rows whose Status is Delete, and fields whose parent Status is Delete, are skipped.
' Calculated field renames may internally replace the field definition but never treat Delete status as an instruction to delete it.

Private Const PROGRESS_FORM As String = "ADC Rename Preview Progress"
Private Const UI_REFRESH_SECONDS As Double = 0.15
Private Const MAX_VISIBLE_LOG_CHARS As Long = 24000
Private Const TEMP_PREFIX As String = "zzADC_RenameTmp_"
Private Const LOG_FILE_NAME As String = "ADC_RenamePreview.log"
Private Const LOG_ORDINARY_TABLE_FIELD_DIAGNOSTICS As Boolean = False

Private mLogFileNo As Integer
Private mLogPath As String
Private mLogOpen As Boolean
Private mProgressText As String
Private mLastUiTimer As Double
Private mCancelRequested As Boolean
Private mProgressOpen As Boolean
Private mSessionTag As String
Private mTempSerial As Long

Public Function Run_PreviewProposedObjectNames()
    ApplyProposedObjectNames True
End Function

Public Function Run_ApplyProposedObjectNames()
    ApplyProposedObjectNames False
End Function

Public Sub ApplyProposedObjectNames(Optional ByVal previewOnly As Boolean = False)
    Dim workbookPath As String
    Dim plan As Collection
    Dim fieldPlan As Collection
    Dim queryFieldOps As Collection
    Dim stats As Object
    Dim hardIssues As Collection
    Dim inventories As Object
    Dim journal As Collection
    Dim tempNames As Collection
    Dim verifyIssues As Collection
    Dim nacTrack As Variant
    Dim nacPerform As Variant
    Dim nacLog As Variant
    Dim nacKnown As Boolean
    Dim Summary As String
    Dim reply As VbMsgBoxResult
    Dim errNumber As Long
    Dim errDescription As String
    Dim errSource As String
    Dim rollbackFailures As Long
    Dim rollbackMessage As String
    Dim journalHasChanges As Boolean
    Dim verifyOK As Boolean
    Dim runTitle As String
    On Error GoTo FatalError
    If previewOnly Then
        runTitle = "ADC Rename Preview"
    Else
        runTitle = "ADC Rename Live"
    End If
    workbookPath = PickRenameWorkbook()
    If Len(workbookPath) = 0 Then Exit Sub
    mSessionTag = Format$(Now, "yyyymmdd_hhnnss")
    mCancelRequested = False
    mProgressText = vbNullString
    mLastUiTimer = 0
    mTempSerial = 0
    StartRunLog workbookPath, previewOnly
    RebuildProgressForm previewOnly, runTitle
    OpenProgressForm
    ProgressAppend "Rename workbook: " & workbookPath
    If previewOnly Then
        ProgressAppend "MODE: PREVIEW ONLY - no database objects or fields will be renamed."
    Else
        ProgressAppend "MODE: LIVE RENAME - database objects and fields may be renamed after preflight."
    End If
    ProgressUpdate "Loading rename workbook", 0, "Reading Rename Proposed and Rename Review...", True
    Set plan = New Collection
    Set fieldPlan = New Collection
    Set queryFieldOps = New Collection
    Set stats = CreateTextDictionary()
    Set hardIssues = New Collection
    Set journal = New Collection
    Set tempNames = New Collection
    LoadRenamePlan workbookPath, plan, fieldPlan, stats, hardIssues
    If mCancelRequested Then
        ProgressAppend "Cancelled before preflight."
        GoTo CleanExit
    End If
    ProgressUpdate "Inventory", 8, "Reading current Access object and field inventory...", True
    Set inventories = BuildInventories()
    PreflightPlan plan, inventories, stats, hardIssues
    PreflightFieldPlan fieldPlan, inventories, queryFieldOps, stats, hardIssues
    PreflightQueryObjectReferenceRewrites plan, hardIssues
    PreflightQueryFieldReferenceRewrites fieldPlan, hardIssues
    PreflightTableFieldReferenceRewrites fieldPlan, queryFieldOps, hardIssues
    If mCancelRequested Then
        ProgressAppend "Cancelled during preflight."
        GoTo CleanExit
    End If
    GetNameAutoCorrectState nacTrack, nacPerform, nacLog, nacKnown
    PreflightCalculatedFieldNameAutoCorrect stats, hardIssues, nacTrack, nacPerform, nacKnown
    LogPreflightDetails stats, hardIssues, nacTrack, nacPerform, nacLog, nacKnown
    LogMissingItems plan, fieldPlan
    If hardIssues.count > 0 Then
        ProgressAppend "PRE-FLIGHT FAILED. No database objects or fields were renamed."
        ProgressUpdate "Preflight failed", 100, FirstIssuesText(hardIssues, 8), True
        MsgBox BuildPreflightFailureMessage(stats, hardIssues), vbCritical + vbOKOnly, runTitle
        GoTo CleanExit
    End If
    If previewOnly Then
        Summary = BuildConfirmationSummary(stats, nacTrack, nacPerform, nacLog, nacKnown, previewOnly)
        ProgressAppend "PREVIEW COMPLETE. No database objects or fields were renamed."
        ProgressUpdate "Preview complete", 100, _
            "Pending objects: " & Format$(StatValue(stats, "Pending"), "#,##0") & _
            " | Pending fields: " & Format$(StatValue(stats, "FieldPending"), "#,##0") & _
            " | Missing objects: " & Format$(StatValue(stats, "Missing"), "#,##0") & _
            " | Missing fields: " & Format$(StatValue(stats, "FieldMissing"), "#,##0"), True
        MsgBox Summary, vbInformation + vbOKOnly, runTitle
        GoTo CleanExit
    End If
    ProgressUpdate "Preflight passed", 12, "Preflight checks complete.", True
    If nacKnown Then
        If (Not CBool(nacTrack)) Or (Not CBool(nacPerform)) Then
            reply = MsgBox( _
                "Name AutoCorrect tracking and/or Perform Name AutoCorrect is OFF." & vbCrLf & vbCrLf & _
                "This utility will NOT change those settings automatically." & vbCrLf & _
                "Access therefore may not propagate renamed field or object references through dependent objects." & _
                vbCrLf & vbCrLf & "Continue with the rename anyway?", _
                vbExclamation + vbYesNo + vbDefaultButton2, runTitle)
            If reply <> vbYes Then
                ProgressAppend "Cancelled because Name AutoCorrect is not fully enabled."
                GoTo CleanExit
            End If
        End If
    End If
    ProgressAppend "Applying query output field renames..."
    ApplyQueryFieldOps queryFieldOps, fieldPlan, journal, stats, hardIssues
    ValidateAllLocalQueryDefs "QUERY OUTPUT PHASE", hardIssues
    If hardIssues.count > 0 Then
        ForEachIssueToLog hardIssues, "QUERY REWRITE ISSUE: "
        ProgressUpdate "Query rewrite failures", 28, FirstIssuesText(hardIssues, 8), True
        Err.Raise vbObjectError + 7174, "ApplyProposedObjectNames", _
            CStr(hardIssues.count) & " query rewrite issue(s) were found. Table fields and top-level objects were not renamed."
    End If
    ProgressAppend "Applying table field renames..."
    ApplyTableFieldPlan fieldPlan, journal, stats, hardIssues
    If hardIssues.count > 0 Then
        ForEachIssueToLog hardIssues, "TABLE FIELD REWRITE ISSUE: "
        ProgressUpdate "Table field rewrite failures", 55, FirstIssuesText(hardIssues, 8), True
        Err.Raise vbObjectError + 7175, "ApplyProposedObjectNames", _
            CStr(hardIssues.count) & " table field dependency rewrite issue(s) were found. Top-level objects were not renamed."
    End If
    ProgressAppend "Applying object renames..."
    ApplyPlan plan, inventories, journal, tempNames, stats
    ValidateAllLocalQueryDefs "OBJECT RENAME PHASE", hardIssues
    If hardIssues.count > 0 Then
        ForEachIssueToLog hardIssues, "OBJECT RENAME QUERY ISSUE: "
        ProgressUpdate "Object rename query failures", 95, FirstIssuesText(hardIssues, 8), True
        Err.Raise vbObjectError + 7177, "ApplyProposedObjectNames", _
            CStr(hardIssues.count) & " query issue(s) remained after object renames."
    End If
    ProgressUpdate "Verification", 96, "Re-reading Access object and field inventory...", True
    Set verifyIssues = New Collection
    verifyOK = VerifyPlan(plan, tempNames, verifyIssues)
    If Not VerifyFieldPlan(fieldPlan, verifyIssues) Then verifyOK = False
    If Not verifyOK Then
        ForEachIssueToLog verifyIssues, "VERIFY FAIL: "
        Err.Raise vbObjectError + 7101, "ApplyProposedObjectNames", _
            "Final verification failed. Rollback will be offered."
    End If
    ProgressAppend "FINAL VERIFICATION: PASS"
    ProgressUpdate "Complete", 100, BuildCompletionText(stats), True
    MsgBox BuildCompletionMessage(stats), vbInformation + vbOKOnly, runTitle
    GoTo CleanExit
FatalError:
    errNumber = Err.Number
    errDescription = Err.Description
    errSource = Err.Source
    On Error Resume Next
    ProgressAppend "ERROR " & CStr(errNumber) & ": " & errDescription & _
        IIf(Len(errSource) > 0, " | Source: " & errSource, vbNullString)
    LogADOProviderErrors "PROVIDER ERROR"
    rollbackMessage = "No rollback-restorable state was recorded; nothing was rolled back."
    journalHasChanges = False
    If Not journal Is Nothing Then
        journalHasChanges = (journal.count > 0)
        If journalHasChanges Then
            reply = MsgBox( _
                "The rename run failed after rollback-restorable state was recorded." & vbCrLf & vbCrLf & _
                "Rollback the recorded changes?" & vbCrLf & vbCrLf & _
                "Yes = rollback now." & vbCrLf & _
                "No = leave this database in its current partial rename state for inspection.", _
                vbExclamation + vbYesNo + vbDefaultButton1, runTitle)
            If reply = vbYes Then
                ProgressAppend "Rollback selected by user."
                ProgressUpdate "Rollback", 0, "Restoring completed physical rename steps...", True
                rollbackFailures = RollbackJournal(journal)
                If rollbackFailures = 0 Then
                    ProgressAppend "ROLLBACK COMPLETE"
                    ProgressUpdate "Rollback complete", 100, "All recorded rename steps were reversed.", True
                    rollbackMessage = "All recorded rename steps were rolled back."
                Else
                    ProgressAppend "ROLLBACK INCOMPLETE. Failures: " & CStr(rollbackFailures)
                    ProgressUpdate "Rollback incomplete", 100, _
                        CStr(rollbackFailures) & " rollback step(s) failed. See the run log.", True
                    rollbackMessage = "Rollback was incomplete; inspect the log before doing anything else."
                End If
            Else
                ProgressAppend "ROLLBACK SKIPPED BY USER"
                ProgressUpdate "Rollback skipped", 100, "Database left in its current partial rename state.", True
                rollbackMessage = "Rollback was skipped by user; the database remains in its partial rename state."
            End If
        End If
    End If
    MsgBox IIf(previewOnly, "Rename preview stopped.", "Live rename stopped.") & vbCrLf & vbCrLf & _
        "Error " & CStr(errNumber) & ": " & errDescription & vbCrLf & _
        IIf(Len(errSource) > 0, "Source: " & errSource & vbCrLf, vbNullString) & vbCrLf & _
        rollbackMessage & vbCrLf & vbCrLf & _
        "Log: " & mLogPath, vbCritical + vbOKOnly, runTitle
CleanExit:
    On Error Resume Next
    CloseProgressForm
    DeleteProgressForm
    CloseRunLog
End Sub

Public Function RenamePreview_RequestCancel() As Boolean
    mCancelRequested = True
    RenamePreview_RequestCancel = True
End Function

Private Function PickRenameWorkbook() As String
    Const msoFileDialogFilePicker As Long = 3
    Dim fd As Object

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select ADC rename workbook"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel workbooks", "*.xlsm;*.xlsx;*.xlsb"
        .InitialFileName = CurrentProject.Path & "\"
        If .Show <> -1 Then Exit Function
        PickRenameWorkbook = CStr(.SelectedItems(1))
    End With
End Function

Private Sub LoadRenamePlan( _
    ByVal workbookPath As String, _
    ByRef plan As Collection, _
    ByRef fieldPlan As Collection, _
    ByVal stats As Object, _
    ByRef hardIssues As Collection)
    Dim workbookDb As Object
    Dim rsReview As Object
    Dim rsProposed As Object
    Dim proposedData As Variant
    Dim reviewData As Variant
    Dim proposedHeaders As Object
    Dim reviewHeaders As Object
    Dim effectiveCurrentByRow As Object
    Dim effectiveProposedByRow As Object
    Dim proposedByRow As Object
    Dim rowMeta As Object
    Dim parentMeta As Object
    Dim i As Long
    Dim proposedRowCount As Long
    Dim reviewRowCount As Long
    Dim rowID As String
    Dim parentID As String
    Dim currentName As String
    Dim proposedName As String
    Dim typeName As String
    Dim statusText As String
    Dim canonicalTypeName As String
    Dim parentType As String
    Dim parentStatus As String
    Dim parentOld As String
    Dim parentNew As String
    Dim item As Object
    Dim invalidReason As String
    Dim savedErr As Long
    Dim savedSource As String
    Dim savedDesc As String
    On Error GoTo LoadFail
    ProgressAppend "Opening rename workbook through Access Database Engine..."
    Set workbookDb = OpenWorkbookReadDatabase(workbookPath)
    Set rsReview = workbookDb.OpenRecordset("SELECT * FROM [Rename Review$]", 4)
    Set reviewHeaders = RecordsetFieldMap(rsReview)
    RequireHeader reviewHeaders, "Effective Current Name", "Rename Review"
    RequireHeader reviewHeaders, "Effective Proposed Name", "Rename Review"
    RequireHeader reviewHeaders, "Row ID", "Rename Review"
    Set effectiveCurrentByRow = CreateTextDictionary()
    Set effectiveProposedByRow = CreateTextDictionary()
    If Not rsReview.EOF Then
        rsReview.MoveLast
        reviewRowCount = rsReview.RecordCount
        rsReview.MoveFirst
        reviewData = rsReview.GetRows(reviewRowCount)
        For i = LBound(reviewData, 2) To UBound(reviewData, 2)
            rowID = CellText(reviewData(CLng(reviewHeaders("Row ID")), i))
            If Len(rowID) > 0 Then
                currentName = CellText(reviewData(CLng(reviewHeaders("Effective Current Name")), i))
                proposedName = CellText(reviewData(CLng(reviewHeaders("Effective Proposed Name")), i))
                If effectiveCurrentByRow.Exists(rowID) Then
                    AddIssue hardIssues, "Duplicate Row ID in Rename Review: " & rowID
                Else
                    effectiveCurrentByRow.Add rowID, currentName
                    effectiveProposedByRow.Add rowID, proposedName
                End If
            End If
        Next i
    End If
    rsReview.Close
    Set rsReview = Nothing
    Set rsProposed = workbookDb.OpenRecordset("SELECT * FROM [Rename Proposed$]", 4)
    Set proposedHeaders = RecordsetFieldMap(rsProposed)
    RequireHeader proposedHeaders, "Row ID", "Rename Proposed"
    RequireHeader proposedHeaders, "Parent ID", "Rename Proposed"
    RequireHeader proposedHeaders, "Type", "Rename Proposed"
    RequireHeader proposedHeaders, "Status", "Rename Proposed"
    If rsProposed.EOF Then
        StatSet stats, "WorkbookRows", 0
        AddIssue hardIssues, "Rename Proposed has no data rows."
        GoTo LoadDone
    End If
    rsProposed.MoveLast
    proposedRowCount = rsProposed.RecordCount
    rsProposed.MoveFirst
    proposedData = rsProposed.GetRows(proposedRowCount)
    proposedRowCount = UBound(proposedData, 2) - LBound(proposedData, 2) + 1
    StatSet stats, "WorkbookRows", proposedRowCount
    Set proposedByRow = CreateTextDictionary()
    For i = LBound(proposedData, 2) To UBound(proposedData, 2)
        rowID = CellText(proposedData(CLng(proposedHeaders("Row ID")), i))
        If Len(rowID) > 0 Then
            If proposedByRow.Exists(rowID) Then
                AddIssue hardIssues, "Duplicate Row ID in Rename Proposed: " & rowID
            Else
                Set rowMeta = CreateTextDictionary()
                rowMeta.Add "ParentID", CellText(proposedData(CLng(proposedHeaders("Parent ID")), i))
                rowMeta.Add "Type", CanonicalType(CellText(proposedData(CLng(proposedHeaders("Type")), i)))
                rowMeta.Add "Status", CellText(proposedData(CLng(proposedHeaders("Status")), i))
                proposedByRow.Add rowID, rowMeta
            End If
        End If
    Next i
    For i = LBound(proposedData, 2) To UBound(proposedData, 2)
        rowID = CellText(proposedData(CLng(proposedHeaders("Row ID")), i))
        parentID = CellText(proposedData(CLng(proposedHeaders("Parent ID")), i))
        typeName = CellText(proposedData(CLng(proposedHeaders("Type")), i))
        statusText = CellText(proposedData(CLng(proposedHeaders("Status")), i))
        canonicalTypeName = CanonicalType(typeName)
        If StrComp(canonicalTypeName, "Field", vbTextCompare) = 0 Then
            If StrComp(statusText, "Delete", vbTextCompare) = 0 Then
                StatInc stats, "FieldDeleteRowsSkipped"
                GoTo NextProposed
            End If
            If Len(rowID) > 0 Then
                If effectiveProposedByRow.Exists(rowID) Then
                    proposedName = CStr(effectiveProposedByRow(rowID))
                    If Len(Trim$(proposedName)) = 0 Then
                        StatInc stats, "BlankProposedConflicts"
                        AddIssue hardIssues, "Supported Field row has a blank Effective Proposed Name in Rename Review at Rename Proposed data row " & _
                            CStr(i - LBound(proposedData, 2) + 2) & "."
                        GoTo NextProposed
                    End If
                End If
            End If
            If Len(rowID) = 0 Then
                AddIssue hardIssues, "Field row has a blank Row ID at Rename Proposed data row " & _
                    CStr(i - LBound(proposedData, 2) + 2) & "."
                GoTo NextProposed
            End If
            If Not effectiveCurrentByRow.Exists(rowID) Or Not effectiveProposedByRow.Exists(rowID) Then
                AddIssue hardIssues, "Field Row ID " & rowID & " exists in Rename Proposed but not Rename Review."
                GoTo NextProposed
            End If
            currentName = CStr(effectiveCurrentByRow(rowID))
            proposedName = CStr(effectiveProposedByRow(rowID))
            If Len(currentName) = 0 Then
                AddIssue hardIssues, "Field Row ID " & rowID & " has a blank Effective Current Name in Rename Review."
                GoTo NextProposed
            End If
            If Len(parentID) = 0 Or Not proposedByRow.Exists(parentID) Then
                AddIssue hardIssues, "Field Row ID " & rowID & " has an unresolved Parent ID '" & parentID & "'."
                GoTo NextProposed
            End If
            Set parentMeta = proposedByRow(parentID)
            parentType = CStr(parentMeta("Type"))
            parentStatus = CStr(parentMeta("Status"))
            If StrComp(parentStatus, "Delete", vbTextCompare) = 0 Then
                StatInc stats, "FieldParentDeleteRowsSkipped"
                GoTo NextProposed
            End If
            If StrComp(parentType, "Table", vbTextCompare) <> 0 And _
               StrComp(parentType, "Query", vbTextCompare) <> 0 Then
                AddIssue hardIssues, "Field Row ID " & rowID & " has unsupported parent type '" & parentType & "'."
                GoTo NextProposed
            End If
            If Not effectiveCurrentByRow.Exists(parentID) Or Not effectiveProposedByRow.Exists(parentID) Then
                AddIssue hardIssues, "Field Row ID " & rowID & " parent Row ID " & parentID & " is missing from Rename Review."
                GoTo NextProposed
            End If
            parentOld = CStr(effectiveCurrentByRow(parentID))
            parentNew = CStr(effectiveProposedByRow(parentID))
            invalidReason = AccessNameProblem(proposedName)
            If Len(invalidReason) > 0 Then
                AddIssue hardIssues, rowID & " proposed field name '" & proposedName & "' is invalid: " & invalidReason
                GoTo NextProposed
            End If
            Set item = CreateTextDictionary()
            item.Add "RowID", rowID
            item.Add "Type", "Field"
            item.Add "Status", statusText
            item.Add "Old", currentName
            item.Add "New", proposedName
            item.Add "WorkOld", currentName
            item.Add "Action", vbNullString
            item.Add "ParentRowID", parentID
            item.Add "ParentType", parentType
            item.Add "ParentOld", parentOld
            item.Add "ParentNew", parentNew
            item.Add "ParentActual", vbNullString
            fieldPlan.Add item
            StatInc stats, "FieldMappings"
            If StrComp(parentType, "Table", vbTextCompare) = 0 Then
                StatInc stats, "TableFieldMappings"
            Else
                StatInc stats, "QueryFieldMappings"
            End If
            GoTo NextProposed
        End If
        If Not IsSupportedTopLevelType(canonicalTypeName) Then
            If Len(typeName) > 0 Then StatInc stats, "UnsupportedRowsSkipped"
            GoTo NextProposed
        End If
        If StrComp(statusText, "Delete", vbTextCompare) = 0 Then
            StatInc stats, "DeleteRowsSkipped"
            GoTo NextProposed
        End If
        If Len(rowID) > 0 Then
            If effectiveProposedByRow.Exists(rowID) Then
                proposedName = CStr(effectiveProposedByRow(rowID))
                If Len(Trim$(proposedName)) = 0 Then
                    StatInc stats, "BlankProposedConflicts"
                    AddIssue hardIssues, "Supported " & canonicalTypeName & _
                        " row has a blank Effective Proposed Name in Rename Review at Rename Proposed data row " & _
                        CStr(i - LBound(proposedData, 2) + 2) & "."
                    GoTo NextProposed
                End If
            End If
        End If
        If StrComp(statusText, "Decided", vbTextCompare) = 0 Then
            StatInc stats, "DecidedRowsIncluded"
        ElseIf StrComp(statusText, "Review", vbTextCompare) = 0 Or _
               StrComp(statusText, "Rvw", vbTextCompare) = 0 Then
            StatInc stats, "ReviewRowsIncluded"
        Else
            StatInc stats, "OtherStatusRowsIncluded"
        End If
        If Len(rowID) = 0 Then
            AddIssue hardIssues, "Supported top-level row has a blank Row ID at Rename Proposed data row " & _
                CStr(i - LBound(proposedData, 2) + 2) & "."
            GoTo NextProposed
        End If
        If Not effectiveCurrentByRow.Exists(rowID) Or Not effectiveProposedByRow.Exists(rowID) Then
            AddIssue hardIssues, "Row ID " & rowID & " exists in Rename Proposed but not Rename Review."
            GoTo NextProposed
        End If
        currentName = CStr(effectiveCurrentByRow(rowID))
        proposedName = CStr(effectiveProposedByRow(rowID))
        If Len(currentName) = 0 Then
            AddIssue hardIssues, "Row ID " & rowID & " has a blank Effective Current Name in Rename Review."
            GoTo NextProposed
        End If
        invalidReason = AccessNameProblem(proposedName)
        If Len(invalidReason) > 0 Then
            AddIssue hardIssues, rowID & " proposed name '" & proposedName & "' is invalid: " & invalidReason
            GoTo NextProposed
        End If
        Set item = CreateTextDictionary()
        item.Add "RowID", rowID
        item.Add "Type", canonicalTypeName
        item.Add "Status", statusText
        item.Add "Old", currentName
        item.Add "New", proposedName
        item.Add "WorkOld", currentName
        item.Add "Action", vbNullString
        plan.Add item
        StatInc stats, "SupportedMappings"
NextProposed:
        If ((i - LBound(proposedData, 2) + 1) Mod 250) = 0 Then
            ProgressUpdate "Loading rename workbook", _
                2 + (5# * (i - LBound(proposedData, 2) + 1) / proposedRowCount), _
                "Reading proposed rows: " & Format$(i - LBound(proposedData, 2) + 1, "#,##0") & " / " & _
                Format$(proposedRowCount, "#,##0"), False
        End If
    Next i
    ProgressAppend "Rename workbook loaded through Access Database Engine. Rows: " & _
        Format$(proposedRowCount, "#,##0")
LoadDone:
    On Error Resume Next
    If Not rsProposed Is Nothing Then rsProposed.Close
    If Not rsReview Is Nothing Then rsReview.Close
    If Not workbookDb Is Nothing Then workbookDb.Close
    Set rsProposed = Nothing
    Set rsReview = Nothing
    Set workbookDb = Nothing
    Exit Sub
LoadFail:
    savedErr = Err.Number
    savedSource = Err.Source
    savedDesc = Err.Description
    On Error Resume Next
    If Not rsProposed Is Nothing Then rsProposed.Close
    If Not rsReview Is Nothing Then rsReview.Close
    If Not workbookDb Is Nothing Then workbookDb.Close
    Set rsProposed = Nothing
    Set rsReview = Nothing
    Set workbookDb = Nothing
    On Error GoTo 0
    If Len(savedSource) > 0 Then savedDesc = savedDesc & " [Source: " & savedSource & "]"
    Err.Raise savedErr, "LoadRenamePlan", savedDesc
End Sub
Private Function OpenWorkbookReadDatabase(ByVal workbookPath As String) As Object
    Dim connectCandidates As Variant
    Dim connectText As Variant
    Dim workbookDb As Object
    Dim errorsText As String
    Dim attemptErr As Long
    Dim attemptDesc As String

    connectCandidates = WorkbookConnectCandidates(workbookPath)

    For Each connectText In connectCandidates
        On Error Resume Next
        Set workbookDb = DBEngine.OpenDatabase(workbookPath, False, True, CStr(connectText))
        attemptErr = Err.Number
        attemptDesc = Err.Description
        Err.Clear
        On Error GoTo 0

        If attemptErr = 0 And Not workbookDb Is Nothing Then
            On Error Resume Next
            ProgressAppend "Workbook ACE connection selected: " & CStr(connectText)
            Err.Clear
            On Error GoTo 0
            Set OpenWorkbookReadDatabase = workbookDb
            Exit Function
        End If

        If Len(errorsText) > 0 Then errorsText = errorsText & " | "
        errorsText = errorsText & CStr(connectText) & ": " & CStr(attemptErr) & " - " & attemptDesc
        Set workbookDb = Nothing
    Next connectText

    Err.Raise vbObjectError + 7010, "OpenWorkbookReadDatabase", _
        "Could not open the selected workbook read-only through the Access Database Engine. " & errorsText
End Function

Private Function WorkbookConnectCandidates(ByVal workbookPath As String) As Variant
    Dim extensionText As String

    extensionText = LCase$(Mid$(workbookPath, InStrRev(workbookPath, ".") + 1))

    Select Case extensionText
        Case "xlsx"
            WorkbookConnectCandidates = Array( _
                "Excel 12.0 Xml;HDR=Yes;IMEX=1;", _
                "Excel 12.0;HDR=Yes;IMEX=1;")
        Case "xlsm"
            WorkbookConnectCandidates = Array( _
                "Excel 12.0 Macro;HDR=Yes;IMEX=1;", _
                "Excel 12.0;HDR=Yes;IMEX=1;", _
                "Excel 12.0 Xml;HDR=Yes;IMEX=1;")
        Case "xlsb"
            WorkbookConnectCandidates = Array( _
                "Excel 12.0;HDR=Yes;IMEX=1;", _
                "Excel 12.0 Macro;HDR=Yes;IMEX=1;")
        Case Else
            Err.Raise vbObjectError + 7011, "WorkbookConnectCandidates", _
                "Unsupported workbook extension: ." & extensionText
    End Select
End Function

Private Function RecordsetFieldMap(ByVal rs As Object) As Object
    Dim fields As Object
    Dim i As Long
    Dim fieldName As String

    Set fields = CreateTextDictionary()
    For i = 0 To rs.fields.count - 1
        fieldName = Trim$(CStr(rs.fields(i).Name))
        If Not fields.Exists(fieldName) Then fields.Add fieldName, i
    Next i

    Set RecordsetFieldMap = fields
End Function

Private Sub RequireHeader(ByVal headers As Object, ByVal headerName As String, ByVal sheetName As String)
    Dim key As Variant
    Dim availableColumns As String
    If Not headers.Exists(headerName) Then
        For Each key In headers.keys
            If Len(availableColumns) > 0 Then availableColumns = availableColumns & " | "
            availableColumns = availableColumns & CStr(key)
        Next key
        If Len(availableColumns) = 0 Then availableColumns = "<none>"
        Err.Raise vbObjectError + 7001, "RequireHeader", _
            "Required column '" & headerName & "' was not found on sheet '" & sheetName & "'. ACE header map columns: " & availableColumns
    End If
End Sub

Private Function CellText(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        CellText = vbNullString
    Else
        CellText = CStr(value)
    End If
End Function

Private Function CanonicalType(ByVal typeName As String) As String
    Select Case LCase$(Trim$(typeName))
        Case "table": CanonicalType = "Table"
        Case "query": CanonicalType = "Query"
        Case "form": CanonicalType = "Form"
        Case "report": CanonicalType = "Report"
        Case "macro": CanonicalType = "Macro"
        Case "module": CanonicalType = "Module"
        Case "field": CanonicalType = "Field"
        Case Else: CanonicalType = Trim$(typeName)
    End Select
End Function

Private Function IsSupportedTopLevelType(ByVal typeName As String) As Boolean
    Select Case typeName
        Case "Table", "Query", "Form", "Report", "Macro", "Module"
            IsSupportedTopLevelType = True
    End Select
End Function

Private Function AccessNameProblem(ByVal objectName As String) As String
    Dim i As Long
    Dim ch As String

    If Len(objectName) = 0 Then
        AccessNameProblem = "blank name"
        Exit Function
    End If

    If Len(objectName) > 64 Then
        AccessNameProblem = "longer than 64 characters"
        Exit Function
    End If

    If Left$(objectName, 1) = " " Then
        AccessNameProblem = "begins with a space"
        Exit Function
    End If

    If InStr(1, objectName, ".", vbBinaryCompare) > 0 Or _
       InStr(1, objectName, "!", vbBinaryCompare) > 0 Or _
       InStr(1, objectName, "`", vbBinaryCompare) > 0 Or _
       InStr(1, objectName, "[", vbBinaryCompare) > 0 Or _
       InStr(1, objectName, "]", vbBinaryCompare) > 0 Then
        AccessNameProblem = "contains a character Access does not allow in object names (. ! ` [ ])"
        Exit Function
    End If

    For i = 1 To Len(objectName)
        ch = Mid$(objectName, i, 1)
        If AscW(ch) >= 0 And AscW(ch) <= 31 Then
            AccessNameProblem = "contains an ASCII control character"
            Exit Function
        End If
    Next i
End Function

Private Function BuildInventories() As Object
    Dim inventories As Object
    Dim inv As Object
    Dim db As Object
    Dim obj As Object

    Set inventories = CreateTextDictionary()
    Set db = currentDb

    Set inv = CreateTextDictionary()
    For Each obj In db.TableDefs
        AddInventoryName inv, CStr(obj.Name)
    Next obj
    inventories.Add "Table", inv

    Set inv = CreateTextDictionary()
    For Each obj In db.QueryDefs
        AddInventoryName inv, CStr(obj.Name)
    Next obj
    inventories.Add "Query", inv

    Set inv = CreateTextDictionary()
    For Each obj In CurrentProject.AllForms
        AddInventoryName inv, CStr(obj.Name)
    Next obj
    inventories.Add "Form", inv

    Set inv = CreateTextDictionary()
    For Each obj In CurrentProject.AllReports
        AddInventoryName inv, CStr(obj.Name)
    Next obj
    inventories.Add "Report", inv

    Set inv = CreateTextDictionary()
    For Each obj In CurrentProject.AllMacros
        AddInventoryName inv, CStr(obj.Name)
    Next obj
    inventories.Add "Macro", inv

    Set inv = CreateTextDictionary()
    For Each obj In CurrentProject.AllModules
        AddInventoryName inv, CStr(obj.Name)
    Next obj
    inventories.Add "Module", inv

    Set BuildInventories = inventories
End Function

Private Sub AddInventoryName(ByVal inv As Object, ByVal objectName As String)
    If inv.Exists(objectName) Then
        inv(objectName) = objectName
    Else
        inv.Add objectName, objectName
    End If
End Sub

Private Sub PreflightPlan( _
    ByVal plan As Collection, _
    ByVal inventories As Object, _
    ByVal stats As Object, _
    ByRef hardIssues As Collection)

    Dim sourceOwners As Object
    Dim targetOwners As Object
    Dim item As Object
    Dim targetItem As Object
    Dim i As Long
    Dim sourceKey As String
    Dim targetKey As String
    Dim typeName As String
    Dim oldName As String
    Dim newName As String
    Dim inv As Object
    Dim oldExists As Boolean
    Dim newExists As Boolean
    Dim sameKey As Boolean

    Set sourceOwners = CreateTextDictionary()
    Set targetOwners = CreateTextDictionary()

    For i = 1 To plan.count
        Set item = plan(i)
        sourceKey = NamespaceNameKey(CStr(item("Type")), CStr(item("Old")))
        targetKey = NamespaceNameKey(CStr(item("Type")), CStr(item("New")))

        If sourceOwners.Exists(sourceKey) Then
            AddIssue hardIssues, "Duplicate Current Name mapping: " & item("Type") & " '" & item("Old") & "'."
        Else
            sourceOwners.Add sourceKey, item
        End If

        If targetOwners.Exists(targetKey) Then
            AddIssue hardIssues, "Duplicate Proposed Name mapping: " & item("Type") & " '" & item("New") & "'."
        Else
            targetOwners.Add targetKey, item
        End If
    Next i

    For i = 1 To plan.count
        Set item = plan(i)
        typeName = CStr(item("Type"))
        oldName = CStr(item("Old"))
        newName = CStr(item("New"))
        Set inv = inventories(typeName)

        oldExists = inv.Exists(oldName)
        newExists = inv.Exists(newName)
        sameKey = (StrComp(oldName, newName, vbTextCompare) = 0)

        If sameKey Then
            If oldExists Then
                If StrComp(InventoryActualName(inv, oldName), newName, vbBinaryCompare) = 0 Then
                    item("Action") = "Unchanged"
                    StatInc stats, "Unchanged"
                Else
                    item("Action") = "Pending"
                    StatInc stats, "Pending"
                    StatInc stats, "CaseOnlyPending"
                End If
            Else
                item("Action") = "Missing"
                StatInc stats, "Missing"
            End If

        ElseIf Not oldExists Then
            If newExists Then
                item("Action") = "Already"
                StatInc stats, "Already"
            ElseIf NameExistsInNamespace(inventories, typeName, newName) Then
                item("Action") = "Conflict"
                StatInc stats, "Conflicts"
                AddIssue hardIssues, typeName & " target '" & newName & _
                    "' is occupied by the other table/query object type while source '" & oldName & "' is missing."
            Else
                item("Action") = "Missing"
                StatInc stats, "Missing"
            End If

        ElseIf Not NameExistsInNamespace(inventories, typeName, newName) Then
            item("Action") = "Pending"
            StatInc stats, "Pending"

        Else
            targetKey = NamespaceNameKey(typeName, newName)
            If sourceOwners.Exists(targetKey) Then
                Set targetItem = sourceOwners(targetKey)
                If StrComp(CStr(targetItem("Old")), CStr(targetItem("New")), vbTextCompare) <> 0 And _
                   InventoryHasExpectedSource(inventories, targetItem) Then
                    item("Action") = "Pending"
                    StatInc stats, "Pending"
                    StatInc stats, "ChainMembers"
                Else
                    item("Action") = "Conflict"
                    StatInc stats, "Conflicts"
                    AddIssue hardIssues, typeName & " target '" & newName & _
                        "' already exists and is not a verified moving source. Source: '" & oldName & "'."
                End If
            Else
                item("Action") = "Conflict"
                StatInc stats, "Conflicts"
                AddIssue hardIssues, typeName & " target '" & newName & _
                    "' already exists. Source: '" & oldName & "'."
            End If
        End If
    Next i
End Sub


Private Sub PreflightFieldPlan( _
    ByVal fieldPlan As Collection, _
    ByVal inventories As Object, _
    ByRef queryFieldOps As Collection, _
    ByVal stats As Object, _
    ByRef hardIssues As Collection)
    Dim fieldInventories As Object
    Dim fieldInventoryAvailable As Object
    Dim sourceOwners As Object
    Dim targetOwners As Object
    Dim item As Object
    Dim targetItem As Object
    Dim inv As Object
    Dim i As Long
    Dim parentKey As String
    Dim sourceKey As String
    Dim targetKey As String
    Dim parentType As String
    Dim parentActual As String
    Dim oldName As String
    Dim newName As String
    Dim oldExists As Boolean
    Dim newExists As Boolean
    Dim sameKey As Boolean
    Dim problem As String
    Set fieldInventories = CreateTextDictionary()
    Set fieldInventoryAvailable = CreateTextDictionary()
    Set sourceOwners = CreateTextDictionary()
    Set targetOwners = CreateTextDictionary()
    For i = 1 To fieldPlan.count
        Set item = fieldPlan(i)
        parentType = CStr(item("ParentType"))
        parentActual = ResolveFieldParentActualName(item, inventories)
        item("ParentActual") = parentActual
        If Len(parentActual) = 0 Then
            item("Action") = "MissingParent"
            StatInc stats, "FieldMissing"
        Else
            parentKey = FieldParentKey(parentType, parentActual)
            If Not fieldInventoryAvailable.Exists(parentKey) Then
                Set inv = Nothing
                problem = vbNullString
                If TryBuildFieldInventory(parentType, parentActual, inv, problem) Then
                    fieldInventories.Add parentKey, inv
                    fieldInventoryAvailable.Add parentKey, True
                Else
                    fieldInventoryAvailable.Add parentKey, False
                    AddIssue hardIssues, parentType & " '" & parentActual & "' field inventory could not be read: " & problem
                End If
            End If
            If Not CBool(fieldInventoryAvailable(parentKey)) Then
                item("Action") = "Conflict"
                StatInc stats, "FieldConflicts"
            Else
                sourceKey = FieldNamespaceKey(parentType, parentActual, CStr(item("Old")))
                targetKey = FieldNamespaceKey(parentType, parentActual, CStr(item("New")))
                If sourceOwners.Exists(sourceKey) Then
                    AddIssue hardIssues, "Duplicate Current Field mapping: " & parentType & " '" & parentActual & _
                        "'.'" & item("Old") & "'."
                Else
                    sourceOwners.Add sourceKey, item
                End If
                If targetOwners.Exists(targetKey) Then
                    AddIssue hardIssues, "Duplicate Proposed Field mapping: " & parentType & " '" & parentActual & _
                        "'.'" & item("New") & "'."
                Else
                    targetOwners.Add targetKey, item
                End If
            End If
        End If
        If (i Mod 100) = 0 And fieldPlan.count > 0 Then
            ProgressUpdate "Preflight fields", 8# + (i / fieldPlan.count), _
                "Reading field inventories: " & Format$(i, "#,##0") & " / " & Format$(fieldPlan.count, "#,##0"), False
        End If
    Next i
    For i = 1 To fieldPlan.count
        Set item = fieldPlan(i)
        If Len(CStr(item("Action"))) = 0 Then
            parentType = CStr(item("ParentType"))
            parentActual = CStr(item("ParentActual"))
            parentKey = FieldParentKey(parentType, parentActual)
            Set inv = fieldInventories(parentKey)
            oldName = CStr(item("Old"))
            newName = CStr(item("New"))
            oldExists = inv.Exists(oldName)
            newExists = inv.Exists(newName)
            sameKey = (StrComp(oldName, newName, vbTextCompare) = 0)
            If sameKey Then
                If oldExists Then
                    If StrComp(InventoryActualName(inv, oldName), newName, vbBinaryCompare) = 0 Then
                        item("Action") = "Unchanged"
                        StatInc stats, "FieldUnchanged"
                    Else
                        item("Action") = "Pending"
                        StatInc stats, "FieldPending"
                        StatInc stats, "FieldCaseOnlyPending"
                    End If
                Else
                    item("Action") = "Missing"
                    StatInc stats, "FieldMissing"
                End If
            ElseIf Not oldExists Then
                If newExists Then
                    item("Action") = "Already"
                    StatInc stats, "FieldAlready"
                Else
                    item("Action") = "Missing"
                    StatInc stats, "FieldMissing"
                End If
            ElseIf Not newExists Then
                item("Action") = "Pending"
                StatInc stats, "FieldPending"
            Else
                targetKey = FieldNamespaceKey(parentType, parentActual, newName)
                If sourceOwners.Exists(targetKey) Then
                    Set targetItem = sourceOwners(targetKey)
                    If StrComp(CStr(targetItem("Old")), CStr(targetItem("New")), vbTextCompare) <> 0 And _
                       inv.Exists(CStr(targetItem("Old"))) Then
                        item("Action") = "Pending"
                        StatInc stats, "FieldPending"
                        StatInc stats, "FieldChainMembers"
                    Else
                        item("Action") = "Conflict"
                        StatInc stats, "FieldConflicts"
                        AddIssue hardIssues, parentType & " '" & parentActual & "' field target '" & newName & _
                            "' already exists and is not a verified moving source. Source: '" & oldName & "'."
                    End If
                Else
                    item("Action") = "Conflict"
                    StatInc stats, "FieldConflicts"
                    AddIssue hardIssues, parentType & " '" & parentActual & "' field target '" & newName & _
                        "' already exists. Source: '" & oldName & "'."
                End If
            End If
        End If
        If (i Mod 250) = 0 And fieldPlan.count > 0 Then
            ProgressUpdate "Preflight fields", 9# + (i / fieldPlan.count), _
                "Classifying field mappings: " & Format$(i, "#,##0") & " / " & Format$(fieldPlan.count, "#,##0"), False
        End If
    Next i
    PreflightCalculatedTableFields fieldPlan, stats, hardIssues
    problem = vbNullString
    If HasPendingFieldType(fieldPlan, "Table") Then
        If Not CanOpenADOXCatalog(problem) Then
            AddIssue hardIssues, "Table field renames require ADOX, but the catalog could not be opened: " & problem
        End If
    End If
    PrepareQueryFieldOps fieldPlan, queryFieldOps, hardIssues
End Sub

Private Function HasPendingFieldType(ByVal fieldPlan As Collection, ByVal parentType As String) As Boolean
    Dim item As Object
    Dim i As Long
    For i = 1 To fieldPlan.count
        Set item = fieldPlan(i)
        If StrComp(CStr(item("ParentType")), parentType, vbTextCompare) = 0 And _
           StrComp(CStr(item("Action")), "Pending", vbTextCompare) = 0 Then
            HasPendingFieldType = True
            Exit Function
        End If
    Next i
End Function

Private Function CanOpenADOXCatalog(ByRef problem As String) As Boolean
    Dim cat As Object
    On Error GoTo OpenFail
    Set cat = CreateObject("ADOX.Catalog")
    Set cat.ActiveConnection = CurrentProject.Connection
    cat.tables.Refresh
    CanOpenADOXCatalog = True
    Exit Function
OpenFail:
    problem = "Error " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function ResolveFieldParentActualName(ByVal item As Object, ByVal inventories As Object) As String
    Dim inv As Object
    Dim parentType As String
    Dim parentOld As String
    Dim parentNew As String
    parentType = CStr(item("ParentType"))
    parentOld = CStr(item("ParentOld"))
    parentNew = CStr(item("ParentNew"))
    Set inv = inventories(parentType)
    If inv.Exists(parentOld) Then
        ResolveFieldParentActualName = InventoryActualName(inv, parentOld)
    ElseIf inv.Exists(parentNew) Then
        ResolveFieldParentActualName = InventoryActualName(inv, parentNew)
    End If
End Function

Private Function TryBuildFieldInventory( _
    ByVal parentType As String, _
    ByVal parentName As String, _
    ByRef inv As Object, _
    ByRef problem As String, _
    Optional ByVal refreshQueryFields As Boolean = False) As Boolean
    Dim db As Object
    Dim parentObject As Object
    Dim fld As Object
    On Error GoTo BuildFail
    Set inv = CreateTextDictionary()
    Set db = currentDb
    If StrComp(parentType, "Table", vbTextCompare) = 0 Then
        Set parentObject = db.TableDefs(parentName)
    ElseIf StrComp(parentType, "Query", vbTextCompare) = 0 Then
        Set parentObject = db.QueryDefs(parentName)
        If refreshQueryFields Then parentObject.fields.Refresh
    Else
        problem = "Unsupported field parent type: " & parentType
        Exit Function
    End If
    For Each fld In parentObject.fields
        AddInventoryName inv, CStr(fld.Name)
    Next fld
    TryBuildFieldInventory = True
    Exit Function
BuildFail:
    problem = "Error " & CStr(Err.Number) & ": " & Err.Description
    Set inv = Nothing
End Function

Private Function FieldParentKey(ByVal parentType As String, ByVal parentName As String) As String
    FieldParentKey = parentType & Chr$(30) & parentName
End Function

Private Function FieldNamespaceKey( _
    ByVal parentType As String, _
    ByVal parentName As String, _
    ByVal fieldName As String) As String
    FieldNamespaceKey = parentType & Chr$(30) & parentName & Chr$(30) & fieldName
End Function

Private Sub PrepareQueryFieldOps( _
    ByVal fieldPlan As Collection, _
    ByRef queryFieldOps As Collection, _
    ByRef hardIssues As Collection)
    Dim prepared As Object
    Dim item As Object
    Dim op As Object
    Dim i As Long
    Dim queryName As String
    Dim oldSql As String
    Dim newSql As String
    Dim problem As String
    Set prepared = CreateTextDictionary()
    For i = 1 To fieldPlan.count
        Set item = fieldPlan(i)
        If StrComp(CStr(item("ParentType")), "Query", vbTextCompare) = 0 And _
           StrComp(CStr(item("Action")), "Pending", vbTextCompare) = 0 Then
            queryName = CStr(item("ParentActual"))
            If Not prepared.Exists(queryName) Then
                prepared.Add queryName, True
                oldSql = vbNullString
                newSql = vbNullString
                problem = vbNullString
                ProgressUpdate "Preflight query fields", 10# + (2# * i / fieldPlan.count), _
                    "Validating query output rewrite: " & queryName, False
                If BuildQueryFieldRewrite(queryName, fieldPlan, oldSql, newSql, problem) Then
                    Set op = CreateTextDictionary()
                    op.Add "Query", queryName
                    op.Add "OldSQL", oldSql
                    op.Add "NewSQL", newSql
                    queryFieldOps.Add op
                Else
                    AddIssue hardIssues, "Query '" & queryName & "' output fields cannot be renamed safely: " & problem
                End If
            End If
        End If
    Next i
End Sub

Private Function BuildQueryFieldRewrite( _
    ByVal queryName As String, _
    ByVal fieldPlan As Collection, _
    ByRef oldSql As String, _
    ByRef newSql As String, _
    ByRef problem As String) As Boolean
    ' Local state for source-query inspection, pending alias mapping, rewritten-SQL validation, and expected output names.
    Dim db As Object
    Dim qdf As Object
    Dim testQdf As Object
    Dim renameMap As Object
    Dim renameByOrdinal As Object
    Dim expectedNames As Object
    Dim item As Object
    Dim i As Long
    Dim currentFieldName As String
    Dim expectedName As String
    Dim matched As Long
    On Error GoTo RewriteFail
    Set db = currentDb
    Set qdf = db.QueryDefs(queryName)
    qdf.fields.Refresh
    If Len(CStr(qdf.Connect)) > 0 Then
        problem = "Pass-through query output aliases are not rewritten automatically."
        Exit Function
    End If
    Select Case CLng(qdf.Type)
        Case dbQSelect, dbQCrosstab, dbQSetOperation
        Case Else
            problem = "Query type " & CStr(qdf.Type) & " does not expose a safely rewritable output SELECT list."
            Exit Function
    End Select
    oldSql = CStr(qdf.sql)
    Set renameMap = CreateTextDictionary()
    For i = 1 To fieldPlan.count
        Set item = fieldPlan(i)
        If StrComp(CStr(item("ParentType")), "Query", vbTextCompare) = 0 And _
           StrComp(CStr(item("ParentActual")), queryName, vbTextCompare) = 0 And _
           StrComp(CStr(item("Action")), "Pending", vbTextCompare) = 0 Then
            renameMap.Add CStr(item("Old")), CStr(item("New"))
        End If
    Next i
    If renameMap.count = 0 Then
        problem = "No pending query field mappings were found."
        Exit Function
    End If
    Set renameByOrdinal = CreateTextDictionary()
    Set expectedNames = CreateTextDictionary()
    For i = 0 To qdf.fields.count - 1
        currentFieldName = CStr(qdf.fields(i).Name)
        expectedName = currentFieldName
        If renameMap.Exists(currentFieldName) Then
            expectedName = CStr(renameMap(currentFieldName))
            renameByOrdinal.Add CStr(i), expectedName
            matched = matched + 1
        End If
        expectedNames.Add CStr(i), expectedName
    Next i
    If matched <> renameMap.count Then
        problem = "Only " & CStr(matched) & " of " & CStr(renameMap.count) & " pending output fields matched QueryDef fields."
        Exit Function
    End If
    If CLng(qdf.Type) = dbQSetOperation Then
        If Not RewriteSetOperationSql(oldSql, renameByOrdinal, renameMap, qdf.fields.count, newSql, problem) Then Exit Function
    Else
        If Not RewriteSelectStatementByOrdinal( _
            oldSql, vbNullString, renameByOrdinal, qdf.fields.count, newSql, problem) Then Exit Function
    End If
    Set testQdf = db.CreateQueryDef(vbNullString, newSql)
    testQdf.fields.Refresh
    If testQdf.fields.count <> qdf.fields.count Then
        problem = "Rewritten SQL changes the query output field count from " & CStr(qdf.fields.count) & _
            " to " & CStr(testQdf.fields.count) & "."
        Exit Function
    End If
    For i = 0 To testQdf.fields.count - 1
        expectedName = CStr(expectedNames(CStr(i)))
        If StrComp(CStr(testQdf.fields(i).Name), expectedName, vbBinaryCompare) <> 0 Then
            problem = "Rewritten SQL output " & CStr(i + 1) & " is '" & CStr(testQdf.fields(i).Name) & _
                "' instead of expected '" & expectedName & "'."
            Exit Function
        End If
    Next i
    BuildQueryFieldRewrite = True
    Exit Function
RewriteFail:
    problem = "Error " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function RewriteSelectStatementByOrdinal( _
    ByVal sqlText As String, _
    ByVal parameterPrefix As String, _
    ByVal renameByOrdinal As Object, _
    ByVal expectedFieldCount As Long, _
    ByRef rewrittenSql As String, _
    ByRef problem As String) As Boolean
    ' Local state for SELECT-list parsing, wildcard expansion, ordinal alias application, and SQL reconstruction.
    Dim selectItems As Collection
    Dim expandedItems As Collection
    Dim rewrittenItems As Collection
    Dim listStart As Long
    Dim listEnd As Long
    Dim i As Long
    If Not FindSelectListBounds(sqlText, listStart, listEnd, problem) Then Exit Function
    Set selectItems = SplitTopLevelSqlList(Mid$(sqlText, listStart, listEnd - listStart))
    If Not ExpandSelectItemsForStatement( _
        sqlText, parameterPrefix, listStart, listEnd, selectItems, expectedFieldCount, expandedItems, problem) Then Exit Function
    Set rewrittenItems = New Collection
    For i = 1 To expandedItems.count
        If renameByOrdinal.Exists(CStr(i - 1)) Then
            rewrittenItems.Add SetSelectItemAlias(CStr(expandedItems(i)), CStr(renameByOrdinal(CStr(i - 1))))
        Else
            rewrittenItems.Add CStr(expandedItems(i))
        End If
    Next i
    rewrittenSql = Left$(sqlText, listStart - 1) & JoinSqlItems(rewrittenItems) & Mid$(sqlText, listEnd)
    RewriteSelectStatementByOrdinal = True
End Function

Private Function ExpandSelectItemsForStatement( _
    ByVal sqlText As String, _
    ByVal parameterPrefix As String, _
    ByVal listStart As Long, _
    ByVal listEnd As Long, _
    ByVal selectItems As Collection, _
    ByVal expectedFieldCount As Long, _
    ByRef expandedItems As Collection, _
    ByRef problem As String) As Boolean
    ' Local state for wildcard probe QueryDefs, expanded output expressions, and output-count validation.
    Dim db As Object
    Dim probeQdf As Object
    Dim itemText As String
    Dim probeBaseSql As String
    Dim probeSql As String
    Dim qualifier As String
    Dim i As Long
    Dim j As Long
    On Error GoTo ExpandFail
    Set expandedItems = New Collection
    probeBaseSql = RemoveTopLevelOrderBy(sqlText)
    For i = 1 To selectItems.count
        itemText = CStr(selectItems(i))
        If IsWildcardSelectItem(itemText) Then
            qualifier = WildcardSelectQualifier(itemText)
            probeSql = Left$(probeBaseSql, listStart - 1) & itemText & " " & Mid$(probeBaseSql, listEnd)
            Set db = currentDb
            Set probeQdf = db.CreateQueryDef(vbNullString, parameterPrefix & probeSql)
            probeQdf.fields.Refresh
            If probeQdf.fields.count = 0 Then
                problem = "Wildcard SELECT item '" & Trim$(itemText) & "' produced no output fields."
                Exit Function
            End If
            For j = 0 To probeQdf.fields.count - 1
                expandedItems.Add BuildWildcardFieldReference(qualifier, CStr(probeQdf.fields(j).Name))
            Next j
            Set probeQdf = Nothing
        Else
            expandedItems.Add itemText
        End If
    Next i
    If expandedItems.count <> expectedFieldCount And selectItems.count = 1 Then
        If Not TryExpandSingleProjectionBySourceMetadata( _
            sqlText, parameterPrefix, expectedFieldCount, expandedItems, problem) Then Exit Function
    End If
    If expandedItems.count <> expectedFieldCount Then
        problem = "Expanded SELECT output item count (" & CStr(expandedItems.count) & _
            ") does not match QueryDef field count (" & CStr(expectedFieldCount) & _
            "). This query shape cannot be rewritten safely by output ordinal."
        Exit Function
    End If
    ExpandSelectItemsForStatement = True
    Exit Function
ExpandFail:
    problem = "Error " & CStr(Err.Number) & " while expanding SELECT outputs: " & Err.Description
End Function

Private Function IsWildcardSelectItem(ByVal selectItem As String) As Boolean
    ' Local state for normalized SELECT-item text.
    Dim itemText As String
    itemText = Trim$(selectItem)
    If itemText = "*" Then
        IsWildcardSelectItem = True
    ElseIf Len(itemText) > 2 Then
        IsWildcardSelectItem = (Right$(itemText, 2) = ".*")
    End If
End Function

Private Function WildcardSelectQualifier(ByVal selectItem As String) As String
    ' Local state for normalized SELECT-item text.
    Dim itemText As String
    itemText = Trim$(selectItem)
    If Len(itemText) > 2 And Right$(itemText, 2) = ".*" Then
        WildcardSelectQualifier = RTrim$(Left$(itemText, Len(itemText) - 2))
    End If
End Function

Private Function BuildWildcardFieldReference(ByVal qualifier As String, ByVal fieldName As String) As String
    If Len(qualifier) > 0 Then
        BuildWildcardFieldReference = qualifier & ".[" & Replace(fieldName, "]", "]]") & "]"
    Else
        BuildWildcardFieldReference = "[" & Replace(fieldName, "]", "]]") & "]"
    End If
End Function

Private Function TryExpandSingleProjectionBySourceMetadata( _
    ByVal sqlText As String, _
    ByVal parameterPrefix As String, _
    ByVal expectedFieldCount As Long, _
    ByRef expandedItems As Collection, _
    ByRef problem As String) As Boolean
    ' Local state for probing a single SELECT projection that Access expands into multiple source-backed output fields.
    Dim db As Object
    Dim probeQdf As Object
    Dim probeField As Object
    Dim sourceKeys As Object
    Dim rebuiltItems As New Collection
    Dim probeSql As String
    Dim sourceTable As String
    Dim sourceField As String
    Dim sourceKey As String
    Dim transformPos As Long
    Dim selectPos As Long
    Dim i As Long
    On Error GoTo ExpandFail
    transformPos = FindTopLevelKeyword(sqlText, "TRANSFORM", 1)
    selectPos = FindTopLevelKeyword(sqlText, "SELECT", 1)
    If transformPos > 0 And selectPos > transformPos Then
        problem = "Crosstab-generated output columns cannot be rewritten safely by source-field metadata."
        Exit Function
    End If
    probeSql = RemoveTopLevelOrderBy(sqlText)
    Set db = currentDb
    Set probeQdf = db.CreateQueryDef(vbNullString, parameterPrefix & probeSql)
    probeQdf.fields.Refresh
    If probeQdf.fields.count <> expectedFieldCount Then
        problem = "The single SELECT projection exposes " & CStr(probeQdf.fields.count) & _
            " fields instead of the expected " & CStr(expectedFieldCount) & "."
        Exit Function
    End If
    Set sourceKeys = CreateTextDictionary()
    For i = 0 To probeQdf.fields.count - 1
        Set probeField = probeQdf.fields(i)
        sourceTable = vbNullString
        sourceField = vbNullString
        If Not QueryFieldSourceMetadata(probeField, sourceTable, sourceField) Then
            problem = "The single SELECT projection expands to multiple fields, but output field '" & _
                CStr(probeField.Name) & "' does not expose source-field metadata required for a safe explicit rewrite."
            Exit Function
        End If
        sourceKey = sourceTable & Chr$(0) & sourceField
        If sourceKeys.Exists(sourceKey) Then
            problem = "The single SELECT projection expands to repeated source field '" & sourceField & _
                "'. The projection cannot be rewritten safely by source-field metadata."
            Exit Function
        End If
        sourceKeys.Add sourceKey, True
        rebuiltItems.Add BuildSourceFieldReference(sourceTable, sourceField)
    Next i
    Set expandedItems = rebuiltItems
    TryExpandSingleProjectionBySourceMetadata = True
    Exit Function
ExpandFail:
    problem = "Error " & CStr(Err.Number) & " while expanding a multi-field SELECT projection: " & Err.Description
End Function

Private Function QueryFieldSourceMetadata( _
    ByVal fieldObject As Object, _
    ByRef sourceTable As String, _
    ByRef sourceField As String) As Boolean
    ' Local state for DAO QueryDef Field.SourceTable and Field.SourceField metadata.
    Dim value As Variant
    On Error GoTo MetadataFail
    value = fieldObject.sourceTable
    If Not IsNull(value) Then sourceTable = CStr(value)
    value = fieldObject.sourceField
    If Not IsNull(value) Then sourceField = CStr(value)
    QueryFieldSourceMetadata = (Len(sourceField) > 0)
    Exit Function
MetadataFail:
    sourceTable = vbNullString
    sourceField = vbNullString
End Function

Private Function BuildSourceFieldReference(ByVal sourceTable As String, ByVal sourceField As String) As String
    If Len(sourceTable) > 0 Then
        BuildSourceFieldReference = "[" & Replace(sourceTable, "]", "]]") & "].[" & _
            Replace(sourceField, "]", "]]") & "]"
    Else
        BuildSourceFieldReference = "[" & Replace(sourceField, "]", "]]") & "]"
    End If
End Function

Private Function RewriteTopLevelOrderByOutputNames(ByVal sqlText As String, ByVal renameMap As Object) As String
    ' Local state for locating the final top-level ORDER BY and updating renamed UNION output identifiers.
    Dim orderPos As Long
    Dim byPos As Long
    Dim clauseStart As Long
    orderPos = FindTopLevelKeyword(sqlText, "ORDER", 1)
    Do While orderPos > 0
        byPos = SkipSqlSpaces(sqlText, orderPos + 5)
        If KeywordAt(sqlText, byPos, "BY") Then
            clauseStart = SkipSqlSpaces(sqlText, byPos + 2)
            RewriteTopLevelOrderByOutputNames = Left$(sqlText, clauseStart - 1) & _
                RewriteOrderByOutputIdentifiers(Mid$(sqlText, clauseStart), renameMap)
            Exit Function
        End If
        orderPos = FindTopLevelKeyword(sqlText, "ORDER", orderPos + 5)
    Loop
    RewriteTopLevelOrderByOutputNames = sqlText
End Function

Private Function RewriteOrderByOutputIdentifiers(ByVal orderText As String, ByVal renameMap As Object) As String
    ' Local state for preserving quoted SQL while replacing exact bracketed or bare output identifiers in ORDER BY.
    Dim result As String
    Dim token As String
    Dim originalToken As String
    Dim ch As String
    Dim prevChar As String
    Dim nextChar As String
    Dim i As Long
    Dim j As Long
    Dim nextPos As Long
    Do While i < Len(orderText)
        i = i + 1
        ch = Mid$(orderText, i, 1)
        If ch = "'" Or ch = """" Then
            j = i + 1
            Do While j <= Len(orderText)
                If Mid$(orderText, j, 1) = ch Then
                    If j < Len(orderText) And Mid$(orderText, j + 1, 1) = ch Then
                        j = j + 2
                    Else
                        Exit Do
                    End If
                Else
                    j = j + 1
                End If
            Loop
            If j > Len(orderText) Then j = Len(orderText)
            result = result & Mid$(orderText, i, j - i + 1)
            i = j
        ElseIf ch = "[" Then
            token = vbNullString
            j = i + 1
            Do While j <= Len(orderText)
                If Mid$(orderText, j, 1) = "]" Then
                    If j < Len(orderText) And Mid$(orderText, j + 1, 1) = "]" Then
                        token = token & "]"
                        j = j + 2
                    Else
                        Exit Do
                    End If
                Else
                    token = token & Mid$(orderText, j, 1)
                    j = j + 1
                End If
            Loop
            If j <= Len(orderText) Then
                originalToken = Mid$(orderText, i, j - i + 1)
                If i > 1 Then prevChar = Mid$(orderText, i - 1, 1) Else prevChar = vbNullString
                If j < Len(orderText) Then nextChar = Mid$(orderText, j + 1, 1) Else nextChar = vbNullString
                If renameMap.Exists(token) And prevChar <> "." And nextChar <> "." Then
                    result = result & "[" & Replace(CStr(renameMap(token)), "]", "]]") & "]"
                Else
                    result = result & originalToken
                End If
                i = j
            Else
                result = result & Mid$(orderText, i)
                Exit Do
            End If
        ElseIf IsSqlWordChar(ch) Then
            j = i
            Do While j <= Len(orderText) And IsSqlWordChar(Mid$(orderText, j, 1))
                j = j + 1
            Loop
            token = Mid$(orderText, i, j - i)
            If i > 1 Then prevChar = Mid$(orderText, i - 1, 1) Else prevChar = vbNullString
            If j <= Len(orderText) Then nextChar = Mid$(orderText, j, 1) Else nextChar = vbNullString
            nextPos = SkipSqlSpaces(orderText, j)
            If renameMap.Exists(token) And prevChar <> "." And nextChar <> "." And _
               StrComp(token, "ASC", vbTextCompare) <> 0 And StrComp(token, "DESC", vbTextCompare) <> 0 And _
               (nextPos > Len(orderText) Or Mid$(orderText, nextPos, 1) <> "(") Then
                result = result & "[" & Replace(CStr(renameMap(token)), "]", "]]") & "]"
            Else
                result = result & token
            End If
            i = j - 1
        Else
            result = result & ch
        End If
    Loop
    RewriteOrderByOutputIdentifiers = result
End Function

Private Function RewriteSetOperationSql( _
    ByVal sqlText As String, _
    ByVal renameByOrdinal As Object, _
    ByVal renameMap As Object, _
    ByVal expectedFieldCount As Long, _
    ByRef rewrittenSql As String, _
    ByRef problem As String) As Boolean
    ' Local state for UNION branch parsing, branch-level validation, rewriting, and SQL reconstruction.
    Dim branches As Collection
    Dim separators As Collection
    Dim rewrittenBranches As Collection
    Dim parameterPrefix As String
    Dim branchSql As String
    Dim branchProbeSql As String
    Dim rewrittenBranch As String
    Dim db As Object
    Dim probeQdf As Object
    Dim i As Long
    On Error GoTo SetFail
    If Not SplitSetOperationSql(sqlText, parameterPrefix, branches, separators, problem) Then Exit Function
    Set rewrittenBranches = New Collection
    For i = 1 To branches.count
        branchSql = CStr(branches(i))
        branchProbeSql = RemoveTopLevelOrderBy(branchSql)
        Set db = currentDb
        Set probeQdf = db.CreateQueryDef(vbNullString, parameterPrefix & branchProbeSql)
        probeQdf.fields.Refresh
        If probeQdf.fields.count <> expectedFieldCount Then
            problem = "UNION branch " & CStr(i) & " exposes " & CStr(probeQdf.fields.count) & _
                " fields instead of the expected " & CStr(expectedFieldCount) & "."
            Exit Function
        End If
        Set probeQdf = Nothing
        rewrittenBranch = vbNullString
        If Not RewriteSelectStatementByOrdinal( _
            branchSql, parameterPrefix, renameByOrdinal, expectedFieldCount, rewrittenBranch, problem) Then
            problem = "UNION branch " & CStr(i) & ": " & problem
            Exit Function
        End If
        rewrittenBranches.Add rewrittenBranch
    Next i
    rewrittenSql = parameterPrefix
    For i = 1 To rewrittenBranches.count
        rewrittenSql = rewrittenSql & CStr(rewrittenBranches(i))
        If i <= separators.count Then rewrittenSql = rewrittenSql & CStr(separators(i))
    Next i
    rewrittenSql = RewriteTopLevelOrderByOutputNames(rewrittenSql, renameMap)
    RewriteSetOperationSql = True
    Exit Function
SetFail:
    problem = "Error " & CStr(Err.Number) & " while rewriting UNION query: " & Err.Description
End Function

Private Function SplitSetOperationSql( _
    ByVal sqlText As String, _
    ByRef parameterPrefix As String, _
    ByRef branches As Collection, _
    ByRef separators As Collection, _
    ByRef problem As String) As Boolean
    ' Local state for locating top-level UNION branches and preserving their original separators.
    Dim firstSelect As Long
    Dim currentStart As Long
    Dim unionPos As Long
    Dim nextSelect As Long
    firstSelect = FindTopLevelKeyword(sqlText, "SELECT", 1)
    If firstSelect = 0 Then
        problem = "No top-level SELECT clause was found in the set-operation query."
        Exit Function
    End If
    parameterPrefix = Left$(sqlText, firstSelect - 1)
    Set branches = New Collection
    Set separators = New Collection
    currentStart = firstSelect
    Do
        unionPos = FindTopLevelKeyword(sqlText, "UNION", currentStart)
        If unionPos = 0 Then
            branches.Add Mid$(sqlText, currentStart)
            Exit Do
        End If
        nextSelect = FindTopLevelKeyword(sqlText, "SELECT", unionPos + 5)
        If nextSelect = 0 Then
            problem = "A top-level UNION is not followed by a SELECT clause."
            Exit Function
        End If
        branches.Add Mid$(sqlText, currentStart, unionPos - currentStart)
        separators.Add Mid$(sqlText, unionPos, nextSelect - unionPos)
        currentStart = nextSelect
    Loop
    If branches.count < 2 Then
        problem = "The set-operation query did not contain multiple top-level SELECT branches."
        Exit Function
    End If
    SplitSetOperationSql = True
End Function

Private Function RemoveTopLevelOrderBy(ByVal sqlText As String) As String
    ' Local state for locating a top-level ORDER BY clause without touching nested expressions.
    Dim orderPos As Long
    Dim byPos As Long
    orderPos = FindTopLevelKeyword(sqlText, "ORDER", 1)
    Do While orderPos > 0
        byPos = SkipSqlSpaces(sqlText, orderPos + 5)
        If KeywordAt(sqlText, byPos, "BY") Then
            RemoveTopLevelOrderBy = RTrim$(Left$(sqlText, orderPos - 1))
            Exit Function
        End If
        orderPos = FindTopLevelKeyword(sqlText, "ORDER", orderPos + 5)
    Loop
    RemoveTopLevelOrderBy = sqlText
End Function

Private Function FindSelectListBounds( _
    ByVal sqlText As String, _
    ByRef listStart As Long, _
    ByRef listEnd As Long, _
    ByRef problem As String) As Boolean
    Dim selectPos As Long
    Dim fromPos As Long
    Dim p As Long
    selectPos = FindTopLevelKeyword(sqlText, "SELECT", 1)
    If selectPos = 0 Then
        problem = "No top-level SELECT clause was found."
        Exit Function
    End If
    p = SkipSqlSpaces(sqlText, selectPos + 6)
    If KeywordAt(sqlText, p, "DISTINCTROW") Then
        p = SkipSqlSpaces(sqlText, p + 11)
    ElseIf KeywordAt(sqlText, p, "DISTINCT") Then
        p = SkipSqlSpaces(sqlText, p + 8)
    End If
    If KeywordAt(sqlText, p, "TOP") Then
        p = SkipSqlSpaces(sqlText, p + 3)
        If Mid$(sqlText, p, 1) = "(" Then
            p = PositionAfterMatchingParen(sqlText, p)
            If p = 0 Then
                problem = "TOP expression has unbalanced parentheses."
                Exit Function
            End If
        Else
            Do While p <= Len(sqlText) And Not IsSqlSpace(Mid$(sqlText, p, 1))
                p = p + 1
            Loop
        End If
        p = SkipSqlSpaces(sqlText, p)
        If KeywordAt(sqlText, p, "PERCENT") Then p = SkipSqlSpaces(sqlText, p + 7)
    End If
    fromPos = FindTopLevelKeyword(sqlText, "FROM", p)
    If fromPos = 0 Then
        problem = "The top-level SELECT clause has no FROM clause; this query shape is not rewritten automatically."
        Exit Function
    End If
    listStart = p
    listEnd = fromPos
    If listEnd <= listStart Then
        problem = "The SELECT output list is empty."
        Exit Function
    End If
    FindSelectListBounds = True
End Function

Private Function FindTopLevelKeyword(ByVal sqlText As String, ByVal keyword As String, ByVal startPos As Long) As Long
    Dim i As Long
    Dim depth As Long
    Dim ch As String
    Dim inBracket As Boolean
    Dim inSingle As Boolean
    Dim inDouble As Boolean
    For i = startPos To Len(sqlText)
        ch = Mid$(sqlText, i, 1)
        If inBracket Then
            If ch = "]" Then
                If i < Len(sqlText) And Mid$(sqlText, i + 1, 1) = "]" Then
                    i = i + 1
                Else
                    inBracket = False
                End If
            End If
        ElseIf inSingle Then
            If ch = "'" Then
                If i < Len(sqlText) And Mid$(sqlText, i + 1, 1) = "'" Then
                    i = i + 1
                Else
                    inSingle = False
                End If
            End If
        ElseIf inDouble Then
            If ch = """" Then
                If i < Len(sqlText) And Mid$(sqlText, i + 1, 1) = """" Then
                    i = i + 1
                Else
                    inDouble = False
                End If
            End If
        Else
            Select Case ch
                Case "[": inBracket = True
                Case "'": inSingle = True
                Case """": inDouble = True
                Case "(": depth = depth + 1
                Case ")": If depth > 0 Then depth = depth - 1
                Case Else
                    If depth = 0 And KeywordAt(sqlText, i, keyword) Then
                        FindTopLevelKeyword = i
                        Exit Function
                    End If
            End Select
        End If
    Next i
End Function

Private Function KeywordAt(ByVal sqlText As String, ByVal startPos As Long, ByVal keyword As String) As Boolean
    Dim beforeChar As String
    Dim afterChar As String
    If startPos < 1 Or startPos + Len(keyword) - 1 > Len(sqlText) Then Exit Function
    If StrComp(Mid$(sqlText, startPos, Len(keyword)), keyword, vbTextCompare) <> 0 Then Exit Function
    If startPos > 1 Then beforeChar = Mid$(sqlText, startPos - 1, 1)
    If startPos + Len(keyword) <= Len(sqlText) Then afterChar = Mid$(sqlText, startPos + Len(keyword), 1)
    If Len(beforeChar) > 0 Then
        If IsSqlWordChar(beforeChar) Then Exit Function
    End If
    If Len(afterChar) > 0 Then
        If IsSqlWordChar(afterChar) Then Exit Function
    End If
    KeywordAt = True
End Function

Private Function IsSqlWordChar(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    IsSqlWordChar = (ch Like "[A-Za-z0-9_]")
End Function

Private Function IsSqlSpace(ByVal ch As String) As Boolean
    IsSqlSpace = (ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf)
End Function

Private Function SkipSqlSpaces(ByVal sqlText As String, ByVal startPos As Long) As Long
    Dim p As Long
    p = startPos
    Do While p <= Len(sqlText) And IsSqlSpace(Mid$(sqlText, p, 1))
        p = p + 1
    Loop
    SkipSqlSpaces = p
End Function

Private Function PositionAfterMatchingParen(ByVal sqlText As String, ByVal openPos As Long) As Long
    Dim i As Long
    Dim depth As Long
    Dim ch As String
    Dim inBracket As Boolean
    Dim inSingle As Boolean
    Dim inDouble As Boolean
    For i = openPos To Len(sqlText)
        ch = Mid$(sqlText, i, 1)
        If inBracket Then
            If ch = "]" Then
                If i < Len(sqlText) And Mid$(sqlText, i + 1, 1) = "]" Then
                    i = i + 1
                Else
                    inBracket = False
                End If
            End If
        ElseIf inSingle Then
            If ch = "'" Then
                If i < Len(sqlText) And Mid$(sqlText, i + 1, 1) = "'" Then
                    i = i + 1
                Else
                    inSingle = False
                End If
            End If
        ElseIf inDouble Then
            If ch = """" Then
                If i < Len(sqlText) And Mid$(sqlText, i + 1, 1) = """" Then
                    i = i + 1
                Else
                    inDouble = False
                End If
            End If
        Else
            Select Case ch
                Case "[": inBracket = True
                Case "'": inSingle = True
                Case """": inDouble = True
                Case "(": depth = depth + 1
                Case ")"
                    depth = depth - 1
                    If depth = 0 Then
                        PositionAfterMatchingParen = i + 1
                        Exit Function
                    End If
            End Select
        End If
    Next i
End Function

Private Function SplitTopLevelSqlList(ByVal listText As String) As Collection
    Dim parts As New Collection
    Dim i As Long
    Dim startPos As Long
    Dim depth As Long
    Dim ch As String
    Dim inBracket As Boolean
    Dim inSingle As Boolean
    Dim inDouble As Boolean
    startPos = 1
    For i = 1 To Len(listText)
        ch = Mid$(listText, i, 1)
        If inBracket Then
            If ch = "]" Then
                If i < Len(listText) And Mid$(listText, i + 1, 1) = "]" Then
                    i = i + 1
                Else
                    inBracket = False
                End If
            End If
        ElseIf inSingle Then
            If ch = "'" Then
                If i < Len(listText) And Mid$(listText, i + 1, 1) = "'" Then
                    i = i + 1
                Else
                    inSingle = False
                End If
            End If
        ElseIf inDouble Then
            If ch = """" Then
                If i < Len(listText) And Mid$(listText, i + 1, 1) = """" Then
                    i = i + 1
                Else
                    inDouble = False
                End If
            End If
        Else
            Select Case ch
                Case "[": inBracket = True
                Case "'": inSingle = True
                Case """": inDouble = True
                Case "(": depth = depth + 1
                Case ")": If depth > 0 Then depth = depth - 1
                Case ","
                    If depth = 0 Then
                        parts.Add Mid$(listText, startPos, i - startPos)
                        startPos = i + 1
                    End If
            End Select
        End If
    Next i
    parts.Add Mid$(listText, startPos)
    Set SplitTopLevelSqlList = parts
End Function

Private Function SetSelectItemAlias(ByVal selectItem As String, ByVal newName As String) As String
    Dim asPos As Long
    Dim expressionText As String
    asPos = FindTopLevelKeyword(selectItem, "AS", 1)
    If asPos > 0 Then
        expressionText = RTrim$(Left$(selectItem, asPos - 1))
    Else
        expressionText = RTrim$(selectItem)
    End If
    SetSelectItemAlias = expressionText & " AS [" & Replace(newName, "]", "]]") & "]"
End Function

Private Function JoinSqlItems(ByVal items As Collection) As String
    Dim i As Long
    Dim result As String
    For i = 1 To items.count
        If i > 1 Then result = result & ","
        result = result & CStr(items(i))
    Next i
    If Len(result) > 0 Then
        If Not IsSqlSpace(Right$(result, 1)) Then result = result & " "
    End If
    JoinSqlItems = result
End Function

Private Sub ApplyQueryFieldOps( _
    ByVal queryFieldOps As Collection, _
    ByVal fieldPlan As Collection, _
    ByRef journal As Collection, _
    ByVal stats As Object, _
    ByRef hardIssues As Collection)
    ' Apply query-output rewrites cumulatively, collect every independently testable failure, and leave rollback journal coverage for every attempted SQL change.
    Dim db As Object
    Dim qdf As Object
    Dim op As Object
    Dim item As Object
    Dim renameMap As Object
    Dim unavailableQueries As Object
    Dim queryNames As New Collection
    Dim i As Long
    Dim j As Long
    Dim queryName As String
    Dim oldSql As String
    Dim newSql As String
    Dim problem As String
    Dim rootFailure As String
    Dim errNumber As Long
    Dim errDescription As String
    Dim changed As Long
    Dim failures As Long
    Dim blockedSources As Long
    Dim blockedDownstreamPaths As Long
    Dim dependentsTested As Long
    If queryFieldOps.count = 0 Then
        ProgressUpdate "Query output fields", 28, "No query output field rewrites are required.", True
        Exit Sub
    End If
    Set db = currentDb
    Set unavailableQueries = CreateTextDictionary()
    For Each qdf In db.QueryDefs
        If Left$(CStr(qdf.Name), 1) <> "~" And Len(CStr(qdf.Connect)) = 0 Then queryNames.Add CStr(qdf.Name)
    Next qdf
    For i = 1 To queryFieldOps.count
        If mCancelRequested Then Err.Raise vbObjectError + 7102, "ApplyQueryFieldOps", "Cancelled by user."
        Set op = queryFieldOps(i)
        queryName = CStr(op("Query"))
        If unavailableQueries.Exists(queryName) Then
            blockedSources = blockedSources + 1
            rootFailure = CStr(unavailableQueries(queryName))
            ProgressAppend "QUERY REWRITE BLOCKED SOURCE: " & queryName & " | Root failure: " & rootFailure
            GoTo NextQueryFieldOp
        End If
        Set renameMap = PendingQueryFieldRenameMap(queryName, fieldPlan)
        If renameMap.count = 0 Then
            failures = failures + 1
            AddIssue hardIssues, "QUERY REWRITE FAIL: Query '" & queryName & "' has no pending output-field mappings at apply time."
            GoTo NextQueryFieldOp
        End If
        CloseObjectIfOpen "Query", queryName
        Set qdf = db.QueryDefs(queryName)
        oldSql = CStr(qdf.sql)
        newSql = CStr(op("NewSQL"))
        problem = vbNullString
        If Not RebaseQueryFieldRewriteSqlForCompletedSources(newSql, fieldPlan, newSql, problem) Then
            failures = failures + 1
            AddIssue hardIssues, "QUERY REWRITE FAIL: Query '" & queryName & _
                "' could not preserve already-applied source-query field renames: " & problem
            GoTo NextQueryFieldOp
        End If
        ProgressAppend "ATTEMPT Query output field rewrite: " & queryName
        ClearADOProviderErrors
        journal.Add Array("QuerySQL", queryName, oldSql, newSql)
        If Not TryAssignAndRefreshQuerySql(qdf, newSql, errNumber, errDescription) Then
            failures = failures + 1
            RestoreQuerySqlAfterFailedAttempt qdf, oldSql, queryName
            AddIssue hardIssues, "QUERY REWRITE FAIL: Query '" & queryName & _
                "' is invalid after its planned output-field rewrite: Error " & CStr(errNumber) & ": " & errDescription
            GoTo NextQueryFieldOp
        End If
        changed = 0
        For j = 1 To fieldPlan.count
            Set item = fieldPlan(j)
            If StrComp(CStr(item("ParentType")), "Query", vbTextCompare) = 0 And _
               StrComp(CStr(item("ParentActual")), queryName, vbTextCompare) = 0 And _
               StrComp(CStr(item("Action")), "Pending", vbTextCompare) = 0 Then
                item("Action") = "Done"
                changed = changed + 1
                StatInc stats, "FieldRenamed"
                StatInc stats, "QueryFieldRenamed"
                ProgressAppend "RENAMED Query field: " & queryName & "." & item("Old") & " -> " & item("New")
            End If
        Next j
        RewriteDependentQueryFieldReferences queryName, renameMap, journal, queryNames, unavailableQueries, _
            hardIssues, failures, dependentsTested
        ProgressUpdate "Query output fields", 12# + (16# * i / queryFieldOps.count), _
            Format$(i, "#,##0") & " / " & Format$(queryFieldOps.count, "#,##0") & _
            " queries | " & queryName & " | " & CStr(changed) & " field(s)", False
NextQueryFieldOp:
    Next i
    blockedDownstreamPaths = 0
    For Each qdf In db.QueryDefs
        If unavailableQueries.Exists(CStr(qdf.Name)) Then
            If StrComp(CStr(qdf.Name), CStr(unavailableQueries(CStr(qdf.Name))), vbTextCompare) <> 0 Then _
                blockedDownstreamPaths = blockedDownstreamPaths + 1
        End If
    Next qdf
    ProgressAppend "QUERY REWRITE EXHAUSTIVE SUMMARY | Planned source rewrites: " & CStr(queryFieldOps.count) & _
        " | Dependent states tested: " & CStr(dependentsTested) & " | Failures: " & CStr(failures) & _
        " | Blocked source rewrites: " & CStr(blockedSources) & " | Blocked downstream paths: " & CStr(blockedDownstreamPaths) & _
        " | Unavailable query paths: " & CStr(unavailableQueries.count)
End Sub

Private Function PendingQueryFieldRenameMap(ByVal queryName As String, ByVal fieldPlan As Collection) As Object
    Set PendingQueryFieldRenameMap = QueryFieldRenameMapForAction(queryName, fieldPlan, "Pending")
End Function

Private Function RebaseQueryFieldRewriteSqlForCompletedSources( _
    ByVal plannedSql As String, _
    ByVal fieldPlan As Collection, _
    ByRef rebasedSql As String, _
    ByRef problem As String) As Boolean
    ' Preserve source-query field renames already completed earlier in this run without rebuilding the validated output-alias plan.
    Dim processed As Object
    Dim renameMap As Object
    Dim item As Object
    Dim sourceQueryName As String
    Dim workingSql As String
    Dim nextSql As String
    Dim changed As Boolean
    Dim sourceReferenced As Boolean
    Dim i As Long
    Set processed = CreateTextDictionary()
    workingSql = plannedSql
    For i = 1 To fieldPlan.count
        Set item = fieldPlan(i)
        If StrComp(CStr(item("ParentType")), "Query", vbTextCompare) = 0 And _
           StrComp(CStr(item("Action")), "Done", vbTextCompare) = 0 Then
            sourceQueryName = CStr(item("ParentActual"))
            If Not processed.Exists(sourceQueryName) Then
                processed.Add sourceQueryName, True
                Set renameMap = QueryFieldRenameMapForAction(sourceQueryName, fieldPlan, "Done")
                If InStr(1, workingSql, sourceQueryName, vbTextCompare) > 0 Then
                    nextSql = vbNullString
                    changed = False
                    sourceReferenced = False
                    problem = vbNullString
                    If Not RewriteQueryFieldReferenceMapSql( _
                        workingSql, sourceQueryName, renameMap, nextSql, changed, sourceReferenced, problem) Then Exit Function
                    workingSql = nextSql
                End If
            End If
        End If
    Next i
    rebasedSql = workingSql
    RebaseQueryFieldRewriteSqlForCompletedSources = True
End Function

Private Function QueryFieldRenameMapForAction( _
    ByVal queryName As String, _
    ByVal fieldPlan As Collection, _
    ByVal actionText As String) As Object
    Dim renameMap As Object
    Dim item As Object
    Dim i As Long
    Set renameMap = CreateTextDictionary()
    For i = 1 To fieldPlan.count
        Set item = fieldPlan(i)
        If StrComp(CStr(item("ParentType")), "Query", vbTextCompare) = 0 And _
           StrComp(CStr(item("ParentActual")), queryName, vbTextCompare) = 0 And _
           StrComp(CStr(item("Action")), actionText, vbTextCompare) = 0 Then
            renameMap.Add CStr(item("Old")), CStr(item("New"))
        End If
    Next i
    Set QueryFieldRenameMapForAction = renameMap
End Function

Private Function TryAssignAndRefreshQuerySql( _
    ByVal qdf As Object, _
    ByVal sqlText As String, _
    ByRef errNumber As Long, _
    ByRef errDescription As String) As Boolean
    On Error Resume Next
    Err.Clear
    qdf.sql = sqlText
    If Err.Number = 0 Then qdf.fields.Refresh
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    TryAssignAndRefreshQuerySql = (errNumber = 0)
End Function

Private Function TryRefreshQueryDef( _
    ByVal qdf As Object, _
    ByRef errNumber As Long, _
    ByRef errDescription As String) As Boolean
    On Error Resume Next
    Err.Clear
    qdf.fields.Refresh
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    TryRefreshQueryDef = (errNumber = 0)
End Function

Private Sub ValidateAllLocalQueryDefs(ByVal phaseLabel As String, ByRef issues As Collection)
    ' Independently refresh every local QueryDef so an unrecognized dependency cannot survive a completed SQL-affecting phase.
    Dim db As Object
    Dim qdf As Object
    Dim errNumber As Long
    Dim errDescription As String
    Dim tested As Long
    Dim failures As Long
    Set db = currentDb
    For Each qdf In db.QueryDefs
        If Left$(CStr(qdf.Name), 1) <> "~" And Len(CStr(qdf.Connect)) = 0 Then
            tested = tested + 1
            If Not TryRefreshQueryDef(qdf, errNumber, errDescription) Then
                failures = failures + 1
                AddIssue issues, phaseLabel & " QUERY VALIDATION FAIL: Query '" & CStr(qdf.Name) & _
                    "' is invalid: Error " & CStr(errNumber) & ": " & errDescription
            End If
        End If
    Next qdf
    ProgressAppend phaseLabel & " QUERY VALIDATION SUMMARY | Tested: " & CStr(tested) & " | Failures: " & CStr(failures)
End Sub

Private Sub RestoreQuerySqlAfterFailedAttempt(ByVal qdf As Object, ByVal oldSql As String, ByVal queryName As String)
    Dim errNumber As Long
    Dim errDescription As String
    On Error Resume Next
    Err.Clear
    qdf.sql = oldSql
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If errNumber <> 0 Then
        Err.Raise errNumber, "ApplyQueryFieldOps", _
            "Query rewrite failure could not restore query '" & queryName & "' before continuing the exhaustive pass: " & errDescription
    End If
End Sub

Private Sub PropagateUnavailableQueryPaths( _
    ByVal queryNames As Collection, _
    ByVal unavailableQueries As Object, _
    ByRef hardIssues As Collection, _
    ByRef failures As Long)
    ' Mark transitive dependents unavailable so the exhaustive pass continues only through query paths independent of known failures.
    Dim db As Object
    Dim qdf As Object
    Dim key As Variant
    Dim queryName As String
    Dim querySql As String
    Dim problem As String
    Dim added As Boolean
    Dim i As Long
    Set db = currentDb
    Do
        added = False
        For i = 1 To queryNames.count
            queryName = CStr(queryNames(i))
            If Not unavailableQueries.Exists(queryName) Then
                Set qdf = db.QueryDefs(queryName)
                querySql = CStr(qdf.sql)
                For Each key In unavailableQueries.keys
                    If InStr(1, querySql, CStr(key), vbTextCompare) > 0 Then
                        problem = vbNullString
                        If QuerySqlReferencesSourceQuery(querySql, CStr(key), problem) Then
                            unavailableQueries.Add queryName, CStr(unavailableQueries(CStr(key)))
                            ProgressAppend "QUERY REWRITE BLOCKED PATH: " & queryName & _
                                " | Root failure: " & CStr(unavailableQueries(CStr(key)))
                            added = True
                            Exit For
                        ElseIf Len(problem) > 0 Then
                            failures = failures + 1
                            unavailableQueries.Add queryName, queryName
                            AddIssue hardIssues, "QUERY REWRITE FAIL: Query '" & queryName & _
                                "' dependency ownership could not be determined while isolating failed query path '" & _
                                CStr(key) & "': " & problem
                            added = True
                            Exit For
                        End If
                    End If
                Next key
            End If
        Next i
    Loop While added
End Sub

Private Function QuerySqlReferencesSourceQuery( _
    ByVal sqlText As String, _
    ByVal sourceQueryName As String, _
    ByRef problem As String) As Boolean
    ' Resolve a local QueryDef dependency with the same source parser used by field-reference rewriting.
    Dim branches As Collection
    Dim separators As Collection
    Dim tokens As Collection
    Dim sourceQualifiers As Object
    Dim sourceTokenIndexes As Object
    Dim aliasTokenIndexes As Object
    Dim parameterPrefix As String
    Dim branchSql As String
    Dim bodySql As String
    Dim sourceFound As Boolean
    Dim soleSource As Boolean
    Dim unionPos As Long
    Dim i As Long
    problem = vbNullString
    bodySql = RemoveTopLevelOrderBy(sqlText)
    unionPos = FindTopLevelKeyword(bodySql, "UNION", 1)
    If unionPos = 0 Then
        If Not TokenizeSqlObjectReferenceText(bodySql, tokens, problem) Then Exit Function
        If Not CollectSqlSourceQualifiers( _
            tokens, sourceQueryName, sourceQualifiers, sourceTokenIndexes, aliasTokenIndexes, sourceFound, soleSource, problem) Then Exit Function
        QuerySqlReferencesSourceQuery = sourceFound
        Exit Function
    End If
    If Not SplitSetOperationSql(bodySql, parameterPrefix, branches, separators, problem) Then Exit Function
    For i = 1 To branches.count
        branchSql = CStr(branches(i))
        Set tokens = Nothing
        sourceFound = False
        soleSource = False
        If Not TokenizeSqlObjectReferenceText(branchSql, tokens, problem) Then Exit Function
        If Not CollectSqlSourceQualifiers( _
            tokens, sourceQueryName, sourceQualifiers, sourceTokenIndexes, aliasTokenIndexes, sourceFound, soleSource, problem) Then Exit Function
        If sourceFound Then
            QuerySqlReferencesSourceQuery = True
            Exit Function
        End If
    Next i
End Function

Private Function CollectSourceOutputNamesToPreserve( _
    ByVal sqlText As String, _
    ByVal sourceQueryName As String, _
    ByVal renameMap As Object, _
    ByRef outputNames As Object, _
    ByRef problem As String) As Boolean
    ' Identify unaliased source fields whose dependent output names must remain stable while the source field reference changes.
    Dim tokens As Collection
    Dim itemTokens As Collection
    Dim sourceQualifiers As Object
    Dim sourceTokenIndexes As Object
    Dim aliasTokenIndexes As Object
    Dim selectItems As Collection
    Dim tokenData As Variant
    Dim middleData As Variant
    Dim lastData As Variant
    Dim itemText As String
    Dim qualifierText As String
    Dim fieldName As String
    Dim listStart As Long
    Dim listEnd As Long
    Dim i As Long
    Dim sourceFound As Boolean
    Dim soleSource As Boolean
    Set outputNames = CreateTextDictionary()
    If Not TokenizeSqlObjectReferenceText(sqlText, tokens, problem) Then Exit Function
    If Not CollectSqlSourceQualifiers( _
        tokens, sourceQueryName, sourceQualifiers, sourceTokenIndexes, aliasTokenIndexes, sourceFound, soleSource, problem) Then Exit Function
    If Not sourceFound Then
        CollectSourceOutputNamesToPreserve = True
        Exit Function
    End If
    If Not FindSelectListBounds(sqlText, listStart, listEnd, problem) Then Exit Function
    Set selectItems = SplitTopLevelSqlList(Mid$(sqlText, listStart, listEnd - listStart))
    For i = 1 To selectItems.count
        itemText = Trim$(CStr(selectItems(i)))
        fieldName = vbNullString
        If FindTopLevelKeyword(itemText, "AS", 1) = 0 Then
            If IsWildcardSelectItem(itemText) Then
                Set itemTokens = Nothing
                If Not TokenizeSqlObjectReferenceText(itemText, itemTokens, problem) Then Exit Function
                If itemTokens.count = 1 Then
                    tokenData = itemTokens(1)
                    If StrComp(CStr(tokenData(0)), "Symbol", vbBinaryCompare) = 0 And CStr(tokenData(3)) = "*" Then
                        problem = "A source-query wildcard would implicitly rename dependent output fields; expand or alias the wildcard before running the rename."
                        Exit Function
                    End If
                ElseIf itemTokens.count = 3 Then
                    tokenData = itemTokens(1)
                    middleData = itemTokens(2)
                    lastData = itemTokens(3)
                    If StrComp(CStr(tokenData(0)), "Identifier", vbBinaryCompare) = 0 And _
                       StrComp(CStr(middleData(0)), "Symbol", vbBinaryCompare) = 0 And CStr(middleData(3)) = "." And _
                       StrComp(CStr(lastData(0)), "Symbol", vbBinaryCompare) = 0 And CStr(lastData(3)) = "*" Then
                        qualifierText = CStr(tokenData(3))
                        If sourceQualifiers.Exists(qualifierText) Then
                            problem = "A source-query wildcard would implicitly rename dependent output fields; expand or alias the wildcard before running the rename."
                            Exit Function
                        End If
                    End If
                End If
            Else
                Set itemTokens = Nothing
                If Not TokenizeSqlObjectReferenceText(itemText, itemTokens, problem) Then Exit Function
                If itemTokens.count = 1 Then
                    tokenData = itemTokens(1)
                    If soleSource And StrComp(CStr(tokenData(0)), "Identifier", vbBinaryCompare) = 0 Then fieldName = CStr(tokenData(3))
                ElseIf itemTokens.count = 3 Then
                    tokenData = itemTokens(1)
                    middleData = itemTokens(2)
                    lastData = itemTokens(3)
                    If StrComp(CStr(tokenData(0)), "Identifier", vbBinaryCompare) = 0 And _
                       StrComp(CStr(middleData(0)), "Symbol", vbBinaryCompare) = 0 And _
                       (CStr(middleData(3)) = "." Or CStr(middleData(3)) = "!") And _
                       StrComp(CStr(lastData(0)), "Identifier", vbBinaryCompare) = 0 Then
                        qualifierText = CStr(tokenData(3))
                        If sourceQualifiers.Exists(qualifierText) Then fieldName = CStr(lastData(3))
                    End If
                End If
            End If
        End If
        If Len(fieldName) > 0 And renameMap.Exists(fieldName) Then
            If Not outputNames.Exists(fieldName) Then outputNames.Add fieldName, True
        End If
    Next i
    CollectSourceOutputNamesToPreserve = True
End Function

Private Function PreserveDependentSelectOutputNames( _
    ByVal originalSql As String, _
    ByVal rewrittenSql As String, _
    ByVal sourceQueryName As String, _
    ByVal renameMap As Object, _
    ByRef preservedSql As String, _
    ByRef problem As String) As Boolean
    ' Alias rewritten simple source fields back to the dependent query's pre-existing output names.
    Dim outputNames As Object
    Dim originalTokens As Collection
    Dim sourceQualifiers As Object
    Dim sourceTokenIndexes As Object
    Dim aliasTokenIndexes As Object
    Dim originalItems As Collection
    Dim rewrittenItems As Collection
    Dim preservedItems As New Collection
    Dim itemTokens As Collection
    Dim tokenData As Variant
    Dim middleData As Variant
    Dim lastData As Variant
    Dim originalItem As String
    Dim rewrittenItem As String
    Dim qualifierText As String
    Dim fieldName As String
    Dim originalListStart As Long
    Dim originalListEnd As Long
    Dim rewrittenListStart As Long
    Dim rewrittenListEnd As Long
    Dim i As Long
    Dim sourceFound As Boolean
    Dim soleSource As Boolean
    preservedSql = rewrittenSql
    If Not CollectSourceOutputNamesToPreserve(originalSql, sourceQueryName, renameMap, outputNames, problem) Then Exit Function
    If outputNames.count = 0 Then
        PreserveDependentSelectOutputNames = True
        Exit Function
    End If
    If Not TokenizeSqlObjectReferenceText(originalSql, originalTokens, problem) Then Exit Function
    If Not CollectSqlSourceQualifiers( _
        originalTokens, sourceQueryName, sourceQualifiers, sourceTokenIndexes, aliasTokenIndexes, sourceFound, soleSource, problem) Then Exit Function
    If Not FindSelectListBounds(originalSql, originalListStart, originalListEnd, problem) Then Exit Function
    If Not FindSelectListBounds(rewrittenSql, rewrittenListStart, rewrittenListEnd, problem) Then Exit Function
    Set originalItems = SplitTopLevelSqlList(Mid$(originalSql, originalListStart, originalListEnd - originalListStart))
    Set rewrittenItems = SplitTopLevelSqlList(Mid$(rewrittenSql, rewrittenListStart, rewrittenListEnd - rewrittenListStart))
    If originalItems.count <> rewrittenItems.count Then
        problem = "Source-reference rewriting changed the SELECT item count while preserving dependent output names."
        Exit Function
    End If
    For i = 1 To originalItems.count
        originalItem = Trim$(CStr(originalItems(i)))
        rewrittenItem = CStr(rewrittenItems(i))
        fieldName = vbNullString
        If FindTopLevelKeyword(originalItem, "AS", 1) = 0 And Not IsWildcardSelectItem(originalItem) Then
            Set itemTokens = Nothing
            If Not TokenizeSqlObjectReferenceText(originalItem, itemTokens, problem) Then Exit Function
            If itemTokens.count = 1 Then
                tokenData = itemTokens(1)
                If soleSource And StrComp(CStr(tokenData(0)), "Identifier", vbBinaryCompare) = 0 Then fieldName = CStr(tokenData(3))
            ElseIf itemTokens.count = 3 Then
                tokenData = itemTokens(1)
                middleData = itemTokens(2)
                lastData = itemTokens(3)
                If StrComp(CStr(tokenData(0)), "Identifier", vbBinaryCompare) = 0 And _
                   StrComp(CStr(middleData(0)), "Symbol", vbBinaryCompare) = 0 And _
                   (CStr(middleData(3)) = "." Or CStr(middleData(3)) = "!") And _
                   StrComp(CStr(lastData(0)), "Identifier", vbBinaryCompare) = 0 Then
                    qualifierText = CStr(tokenData(3))
                    If sourceQualifiers.Exists(qualifierText) Then fieldName = CStr(lastData(3))
                End If
            End If
        End If
        If Len(fieldName) > 0 And outputNames.Exists(fieldName) Then
            preservedItems.Add SetSelectItemAlias(rewrittenItem, fieldName)
        Else
            preservedItems.Add rewrittenItem
        End If
    Next i
    preservedSql = Left$(rewrittenSql, rewrittenListStart - 1) & JoinSqlItems(preservedItems) & Mid$(rewrittenSql, rewrittenListEnd)
    PreserveDependentSelectOutputNames = True
End Function

Private Sub PreflightQueryFieldReferenceRewrites(ByVal fieldPlan As Collection, ByRef hardIssues As Collection)
    ' Prove that every pending query-output rename can be propagated through saved local QueryDef SQL before live mutation.
    Dim db As Object
    Dim qdf As Object
    Dim sourceQueries As Object
    Dim renameMap As Object
    Dim item As Object
    Dim key As Variant
    Dim sourceQueryName As String
    Dim oldSql As String
    Dim newSql As String
    Dim problem As String
    Dim changed As Boolean
    Dim sourceReferenced As Boolean
    Dim i As Long
    Set sourceQueries = CreateTextDictionary()
    For i = 1 To fieldPlan.count
        Set item = fieldPlan(i)
        If StrComp(CStr(item("ParentType")), "Query", vbTextCompare) = 0 And _
           StrComp(CStr(item("Action")), "Pending", vbTextCompare) = 0 Then
            sourceQueryName = CStr(item("ParentActual"))
            If Not sourceQueries.Exists(sourceQueryName) Then sourceQueries.Add sourceQueryName, True
        End If
    Next i
    If sourceQueries.count = 0 Then Exit Sub
    Set db = currentDb
    For Each key In sourceQueries.keys
        sourceQueryName = CStr(key)
        Set renameMap = PendingQueryFieldRenameMap(sourceQueryName, fieldPlan)
        For Each qdf In db.QueryDefs
            If Left$(CStr(qdf.Name), 1) <> "~" And Len(CStr(qdf.Connect)) = 0 And _
               StrComp(CStr(qdf.Name), sourceQueryName, vbTextCompare) <> 0 Then
                oldSql = CStr(qdf.sql)
                If InStr(1, oldSql, sourceQueryName, vbTextCompare) > 0 Then
                    newSql = vbNullString
                    changed = False
                    sourceReferenced = False
                    problem = vbNullString
                    If Not RewriteQueryFieldReferenceMapSql( _
                        oldSql, sourceQueryName, renameMap, newSql, changed, sourceReferenced, problem) Then
                        AddIssue hardIssues, "Query '" & CStr(qdf.Name) & "' cannot safely track output-field renames from query '" & _
                            sourceQueryName & "': " & problem
                    End If
                End If
            End If
        Next qdf
    Next key
End Sub

Private Sub PreflightTableFieldReferenceRewrites( _
    ByVal fieldPlan As Collection, _
    ByVal queryFieldOps As Collection, _
    ByRef hardIssues As Collection)
    ' Prove table-field dependency rewrites against each query's planned post-output-rewrite SQL rather than its earlier saved SQL.
    Dim db As Object
    Dim qdf As Object
    Dim sourceTables As Object
    Dim projectedQuerySql As Object
    Dim renameMap As Object
    Dim op As Object
    Dim item As Object
    Dim key As Variant
    Dim tableName As String
    Dim queryName As String
    Dim currentSql As String
    Dim newSql As String
    Dim problem As String
    Dim changed As Boolean
    Dim sourceReferenced As Boolean
    Dim i As Long
    Set sourceTables = CreateTextDictionary()
    For i = 1 To fieldPlan.count
        Set item = fieldPlan(i)
        If StrComp(CStr(item("ParentType")), "Table", vbTextCompare) = 0 And _
           StrComp(CStr(item("Action")), "Pending", vbTextCompare) = 0 Then
            tableName = CStr(item("ParentActual"))
            If Not sourceTables.Exists(tableName) Then sourceTables.Add tableName, True
        End If
    Next i
    If sourceTables.count = 0 Then Exit Sub
    Set projectedQuerySql = CreateTextDictionary()
    For i = 1 To queryFieldOps.count
        Set op = queryFieldOps(i)
        projectedQuerySql.Add CStr(op("Query")), CStr(op("NewSQL"))
    Next i
    Set db = currentDb
    For Each qdf In db.QueryDefs
        If Left$(CStr(qdf.Name), 1) <> "~" And Len(CStr(qdf.Connect)) = 0 Then
            queryName = CStr(qdf.Name)
            If projectedQuerySql.Exists(queryName) Then
                currentSql = CStr(projectedQuerySql(queryName))
            Else
                currentSql = CStr(qdf.sql)
            End If
            For Each key In sourceTables.keys
                tableName = CStr(key)
                If InStr(1, currentSql, tableName, vbTextCompare) > 0 Then
                    Set renameMap = PendingTableFieldRenameMap(tableName, fieldPlan)
                    newSql = vbNullString
                    changed = False
                    sourceReferenced = False
                    problem = vbNullString
                    If Not RewriteQueryFieldReferenceMapSql( _
                        currentSql, tableName, renameMap, newSql, changed, sourceReferenced, problem) Then
                        AddIssue hardIssues, "Query '" & queryName & "' cannot safely track field renames from table '" & _
                            tableName & "': " & problem
                        Exit For
                    End If
                    If sourceReferenced Then currentSql = newSql
                End If
            Next key
        End If
    Next qdf
End Sub

Private Function PendingTableFieldRenameMap(ByVal tableName As String, ByVal fieldPlan As Collection) As Object
    Dim renameMap As Object
    Dim item As Object
    Dim i As Long
    Set renameMap = CreateTextDictionary()
    For i = 1 To fieldPlan.count
        Set item = fieldPlan(i)
        If StrComp(CStr(item("ParentType")), "Table", vbTextCompare) = 0 And _
           StrComp(CStr(item("ParentActual")), tableName, vbTextCompare) = 0 And _
           StrComp(CStr(item("Action")), "Pending", vbTextCompare) = 0 Then
            renameMap(CStr(item("Old"))) = CStr(item("New"))
        End If
    Next i
    Set PendingTableFieldRenameMap = renameMap
End Function

Private Sub RewriteDependentQueryFieldReferences( _
    ByVal sourceQueryName As String, _
    ByVal renameMap As Object, _
    ByRef journal As Collection, _
    ByVal queryNames As Collection, _
    ByVal unavailableQueries As Object, _
    ByRef hardIssues As Collection, _
    ByRef failures As Long, _
    ByRef dependentsTested As Long)
    ' Update every independently testable direct dependent, preserving its output contract and collecting failures instead of stopping at the first one.
    Dim db As Object
    Dim qdf As Object
    Dim queryName As String
    Dim oldSql As String
    Dim newSql As String
    Dim problem As String
    Dim changed As Boolean
    Dim sourceReferenced As Boolean
    Dim errNumber As Long
    Dim errDescription As String
    Dim i As Long
    Set db = currentDb
    For i = 1 To queryNames.count
        queryName = CStr(queryNames(i))
        If StrComp(queryName, sourceQueryName, vbTextCompare) <> 0 And _
           Not unavailableQueries.Exists(queryName) Then
            Set qdf = db.QueryDefs(queryName)
            oldSql = CStr(qdf.sql)
            If InStr(1, oldSql, sourceQueryName, vbTextCompare) > 0 Then
                newSql = vbNullString
                changed = False
                sourceReferenced = False
                problem = vbNullString
                If Not RewriteQueryFieldReferenceMapSql( _
                    oldSql, sourceQueryName, renameMap, newSql, changed, sourceReferenced, problem) Then
                    failures = failures + 1
                    unavailableQueries.Add queryName, queryName
                    AddIssue hardIssues, "QUERY REWRITE FAIL: Query '" & queryName & _
                        "' cannot safely track output-field renames from query '" & sourceQueryName & "': " & problem
                    PropagateUnavailableQueryPaths queryNames, unavailableQueries, hardIssues, failures
                ElseIf sourceReferenced Then
                    dependentsTested = dependentsTested + 1
                    If changed Then
                        CloseObjectIfOpen "Query", queryName
                        journal.Add Array("QueryFieldRefSQL", queryName, oldSql, newSql, sourceQueryName)
                        If TryAssignAndRefreshQuerySql(qdf, newSql, errNumber, errDescription) Then
                            ProgressAppend "UPDATED Query field references: " & queryName & " | Source query: " & sourceQueryName
                        Else
                            failures = failures + 1
                            RestoreQuerySqlAfterFailedAttempt qdf, oldSql, queryName
                            unavailableQueries.Add queryName, queryName
                            AddIssue hardIssues, "QUERY REWRITE FAIL: Dependent query '" & queryName & _
                                "' is invalid after output-field renames from query '" & sourceQueryName & _
                                "': Error " & CStr(errNumber) & ": " & errDescription
                            PropagateUnavailableQueryPaths queryNames, unavailableQueries, hardIssues, failures
                        End If
                    ElseIf Not TryRefreshQueryDef(qdf, errNumber, errDescription) Then
                        failures = failures + 1
                        unavailableQueries.Add queryName, queryName
                        AddIssue hardIssues, "QUERY REWRITE FAIL: Dependent query '" & queryName & _
                            "' becomes invalid after output-field renames from query '" & sourceQueryName & _
                            "' even though no SQL text rewrite is required: Error " & CStr(errNumber) & ": " & errDescription
                        PropagateUnavailableQueryPaths queryNames, unavailableQueries, hardIssues, failures
                    End If
                End If
            End If
        End If
    Next i
End Sub

Private Function RewriteQueryFieldReferenceMapSql( _
    ByVal sqlText As String, _
    ByVal sourceQueryName As String, _
    ByVal renameMap As Object, _
    ByRef rewrittenSql As String, _
    ByRef changed As Boolean, _
    ByRef sourceReferenced As Boolean, _
    ByRef problem As String) As Boolean
    ' Rewrite source references while preserving each dependent query's existing output names until its own output-rename step.
    Dim branches As Collection
    Dim separators As Collection
    Dim rewrittenBranches As Collection
    Dim sqlTokens As Collection
    Dim parameterPrefix As String
    Dim branchSql As String
    Dim rewrittenBranch As String
    Dim preservedBranch As String
    Dim setBodySql As String
    Dim orderBySuffix As String
    Dim branchChanged As Boolean
    Dim branchSourceReferenced As Boolean
    Dim unionPos As Long
    Dim orderByStart As Long
    Dim i As Long
    changed = False
    sourceReferenced = False
    rewrittenSql = sqlText
    If renameMap Is Nothing Then
        problem = "Field rename map is unavailable."
        Exit Function
    End If
    If renameMap.count = 0 Then
        RewriteQueryFieldReferenceMapSql = True
        Exit Function
    End If
    unionPos = FindTopLevelKeyword(sqlText, "UNION", 1)
    If unionPos = 0 Then
        If Not RewriteQueryFieldReferenceMapBranch( _
            sqlText, sourceQueryName, renameMap, rewrittenSql, changed, sourceReferenced, problem) Then Exit Function
        If changed And FindTopLevelKeyword(sqlText, "SELECT", 1) > 0 Then
            preservedBranch = vbNullString
            If Not PreserveDependentSelectOutputNames( _
                sqlText, rewrittenSql, sourceQueryName, renameMap, preservedBranch, problem) Then Exit Function
            rewrittenSql = preservedBranch
        End If
        RewriteQueryFieldReferenceMapSql = True
        Exit Function
    End If
    If Not TokenizeSqlObjectReferenceText(sqlText, sqlTokens, problem) Then Exit Function
    orderByStart = TopLevelOrderByStartPosition(sqlTokens)
    If orderByStart > 0 Then
        setBodySql = Left$(sqlText, orderByStart - 1)
        orderBySuffix = Mid$(sqlText, orderByStart)
    Else
        setBodySql = sqlText
        orderBySuffix = vbNullString
    End If
    If Not SplitSetOperationSql(setBodySql, parameterPrefix, branches, separators, problem) Then Exit Function
    Set rewrittenBranches = New Collection
    For i = 1 To branches.count
        branchSql = CStr(branches(i))
        rewrittenBranch = vbNullString
        branchChanged = False
        branchSourceReferenced = False
        If Not RewriteQueryFieldReferenceMapBranch( _
            branchSql, sourceQueryName, renameMap, rewrittenBranch, branchChanged, branchSourceReferenced, problem) Then
            problem = "UNION branch " & CStr(i) & ": " & problem
            Exit Function
        End If
        If branchChanged Then
            preservedBranch = vbNullString
            If Not PreserveDependentSelectOutputNames( _
                branchSql, rewrittenBranch, sourceQueryName, renameMap, preservedBranch, problem) Then
                problem = "UNION branch " & CStr(i) & ": " & problem
                Exit Function
            End If
            rewrittenBranch = preservedBranch
            changed = True
        End If
        rewrittenBranches.Add rewrittenBranch
        If branchSourceReferenced Then sourceReferenced = True
    Next i
    rewrittenSql = parameterPrefix
    For i = 1 To rewrittenBranches.count
        rewrittenSql = rewrittenSql & CStr(rewrittenBranches(i))
        If i <= separators.count Then rewrittenSql = rewrittenSql & CStr(separators(i))
    Next i
    rewrittenSql = rewrittenSql & orderBySuffix
    RewriteQueryFieldReferenceMapSql = True
End Function

Private Function RewriteQueryFieldReferenceMapBranch( _
    ByVal sqlText As String, _
    ByVal sourceQueryName As String, _
    ByVal renameMap As Object, _
    ByRef rewrittenSql As String, _
    ByRef changed As Boolean, _
    ByRef sourceReferenced As Boolean, _
    ByRef problem As String) As Boolean
    ' Replace only references proven to belong to the named source query; literals, aliases, and other source qualifiers are preserved.
    Dim tokens As Collection
    Dim sourceQualifiers As Object
    Dim sourceTokenIndexes As Object
    Dim aliasTokenIndexes As Object
    Dim replacements As Object
    Dim selectAliasNames As Object
    Dim preservedOutputNames As Object
    Dim key As Variant
    Dim tokenData As Variant
    Dim qualifierData As Variant
    Dim previousData As Variant
    Dim nextData As Variant
    Dim tokenText As String
    Dim qualifierText As String
    Dim newFieldName As String
    Dim result As String
    Dim bodyStart As Long
    Dim orderByStart As Long
    Dim cursor As Long
    Dim tokenIndex As Long
    Dim startPos As Long
    Dim endPos As Long
    Dim sourceFound As Boolean
    Dim soleSource As Boolean
    changed = False
    sourceReferenced = False
    rewrittenSql = sqlText
    If Not TokenizeSqlObjectReferenceText(sqlText, tokens, problem) Then Exit Function
    If Not CollectSqlSourceQualifiers( _
        tokens, sourceQueryName, sourceQualifiers, sourceTokenIndexes, aliasTokenIndexes, sourceFound, soleSource, problem) Then Exit Function
    If Not sourceFound Then
        RewriteQueryFieldReferenceMapBranch = True
        Exit Function
    End If
    sourceReferenced = True
    bodyStart = SqlStatementBodyStartPosition(sqlText)
    orderByStart = TopLevelOrderByStartPosition(tokens)
    Set selectAliasNames = CollectTopLevelSelectAliasNames(tokens)
    Set preservedOutputNames = CreateTextDictionary()
    If FindTopLevelKeyword(sqlText, "SELECT", 1) > 0 Then
        If Not CollectSourceOutputNamesToPreserve(sqlText, sourceQueryName, renameMap, preservedOutputNames, problem) Then Exit Function
        For Each key In preservedOutputNames.keys
            If Not selectAliasNames.Exists(CStr(key)) Then selectAliasNames.Add CStr(key), True
        Next key
    End If
    Set replacements = CreateTextDictionary()
    For tokenIndex = 1 To tokens.count
        tokenData = tokens(tokenIndex)
        If StrComp(CStr(tokenData(0)), "Identifier", vbBinaryCompare) = 0 Then
            tokenText = CStr(tokenData(3))
            If renameMap.Exists(tokenText) And CLng(tokenData(1)) >= bodyStart Then
                If Not sourceTokenIndexes.Exists(CStr(tokenIndex)) And Not aliasTokenIndexes.Exists(CStr(tokenIndex)) Then
                    newFieldName = CStr(renameMap(tokenText))
                    If tokenIndex >= 3 Then
                        previousData = tokens(tokenIndex - 1)
                        qualifierData = tokens(tokenIndex - 2)
                        If StrComp(CStr(previousData(0)), "Symbol", vbBinaryCompare) = 0 And _
                           (CStr(previousData(3)) = "." Or CStr(previousData(3)) = "!") And _
                           StrComp(CStr(qualifierData(0)), "Identifier", vbBinaryCompare) = 0 Then
                            qualifierText = CStr(qualifierData(3))
                            If sourceQualifiers.Exists(qualifierText) Then replacements.Add CStr(tokenIndex), newFieldName
                            GoTo NextFieldToken
                        End If
                    End If
                    If tokenIndex < tokens.count Then
                        nextData = tokens(tokenIndex + 1)
                        If StrComp(CStr(nextData(0)), "Symbol", vbBinaryCompare) = 0 Then
                            If CStr(nextData(3)) = "." Or CStr(nextData(3)) = "!" Or CStr(nextData(3)) = "(" Then GoTo NextFieldToken
                        End If
                    End If
                    If tokenIndex > 1 Then
                        previousData = tokens(tokenIndex - 1)
                        If SqlTokenIsKeyword(previousData, "AS") Then GoTo NextFieldToken
                    End If
                    If soleSource Then
                        If orderByStart > 0 And CLng(tokenData(1)) >= orderByStart Then
                            If selectAliasNames.Exists(tokenText) Then GoTo NextFieldToken
                        End If
                        replacements.Add CStr(tokenIndex), newFieldName
                    Else
                        If orderByStart > 0 And CLng(tokenData(1)) >= orderByStart Then
                            If selectAliasNames.Exists(tokenText) Then GoTo NextFieldToken
                        End If
                        problem = "Unqualified field reference '" & tokenText & "' cannot be attributed safely to source query '" & _
                            sourceQueryName & "' in a multi-source SQL branch. Qualify the field in the saved query before renaming it."
                        Exit Function
                    End If
                End If
            End If
        End If
NextFieldToken:
    Next tokenIndex
    If replacements.count = 0 Then
        RewriteQueryFieldReferenceMapBranch = True
        Exit Function
    End If
    cursor = 1
    For tokenIndex = 1 To tokens.count
        If replacements.Exists(CStr(tokenIndex)) Then
            tokenData = tokens(tokenIndex)
            startPos = CLng(tokenData(1))
            endPos = CLng(tokenData(2))
            result = result & Mid$(sqlText, cursor, startPos - cursor) & "[" & _
                Replace(CStr(replacements(CStr(tokenIndex))), "]", "]]") & "]"
            cursor = endPos + 1
        End If
    Next tokenIndex
    result = result & Mid$(sqlText, cursor)
    rewrittenSql = result
    changed = True
    RewriteQueryFieldReferenceMapBranch = True
End Function

Private Function CollectTopLevelSelectAliasNames(ByVal tokens As Collection) As Object
    ' Preserve explicit output aliases when the same text is later used by a top-level ORDER BY clause.
    Dim aliases As Object
    Dim tokenData As Variant
    Dim nextData As Variant
    Dim depth As Long
    Dim i As Long
    Dim inSelect As Boolean
    Set aliases = CreateTextDictionary()
    For i = 1 To tokens.count
        tokenData = tokens(i)
        If StrComp(CStr(tokenData(0)), "Symbol", vbBinaryCompare) = 0 Then
            If CStr(tokenData(3)) = "(" Then
                depth = depth + 1
            ElseIf CStr(tokenData(3)) = ")" Then
                If depth > 0 Then depth = depth - 1
            End If
        ElseIf depth = 0 And StrComp(CStr(tokenData(0)), "Identifier", vbBinaryCompare) = 0 Then
            If SqlTokenIsKeyword(tokenData, "SELECT") Then
                inSelect = True
            ElseIf inSelect And SqlTokenIsKeyword(tokenData, "FROM") Then
                Exit For
            ElseIf inSelect And SqlTokenIsKeyword(tokenData, "AS") Then
                If i < tokens.count Then
                    nextData = tokens(i + 1)
                    If StrComp(CStr(nextData(0)), "Identifier", vbBinaryCompare) = 0 Then
                        If Not aliases.Exists(CStr(nextData(3))) Then aliases.Add CStr(nextData(3)), True
                    End If
                End If
            End If
        End If
    Next i
    Set CollectTopLevelSelectAliasNames = aliases
End Function

Private Function TopLevelOrderByStartPosition(ByVal tokens As Collection) As Long
    Dim tokenData As Variant
    Dim nextData As Variant
    Dim depth As Long
    Dim i As Long
    For i = 1 To tokens.count - 1
        tokenData = tokens(i)
        If StrComp(CStr(tokenData(0)), "Symbol", vbBinaryCompare) = 0 Then
            If CStr(tokenData(3)) = "(" Then
                depth = depth + 1
            ElseIf CStr(tokenData(3)) = ")" Then
                If depth > 0 Then depth = depth - 1
            End If
        ElseIf depth = 0 And SqlTokenIsKeyword(tokenData, "ORDER") Then
            nextData = tokens(i + 1)
            If SqlTokenIsKeyword(nextData, "BY") Then
                TopLevelOrderByStartPosition = CLng(tokenData(1))
                Exit Function
            End If
        End If
    Next i
End Function

Private Function CollectSqlSourceQualifiers( _
    ByVal tokens As Collection, _
    ByVal sourceQueryName As String, _
    ByRef sourceQualifiers As Object, _
    ByRef sourceTokenIndexes As Object, _
    ByRef aliasTokenIndexes As Object, _
    ByRef sourceFound As Boolean, _
    ByRef soleSource As Boolean, _
    ByRef problem As String) As Boolean
    ' Collect aliases within one UNION branch; duplicate aliases owned by different sources are rejected as ambiguous.
    Dim inFrom(0 To 255) As Boolean
    Dim expectSource(0 To 255) As Boolean
    Dim qualifierOwners As Object
    Dim tokenData As Variant
    Dim aliasName As String
    Dim tokenText As String
    Dim tokenType As String
    Dim symbolText As String
    Dim qualifierText As String
    Dim ownerText As String
    Dim depth As Long
    Dim i As Long
    Dim aliasIndex As Long
    Dim localSourceCount As Long
    Dim isExternal As Boolean
    Set sourceQualifiers = CreateTextDictionary()
    Set sourceTokenIndexes = CreateTextDictionary()
    Set aliasTokenIndexes = CreateTextDictionary()
    Set qualifierOwners = CreateTextDictionary()
    For i = 1 To tokens.count
        tokenData = tokens(i)
        tokenType = CStr(tokenData(0))
        If StrComp(tokenType, "Symbol", vbBinaryCompare) = 0 Then
            symbolText = CStr(tokenData(3))
            If symbolText = "(" Then
                If depth = 255 Then
                    problem = "SQL nesting exceeds the supported field-reference rewrite depth."
                    Exit Function
                End If
                If expectSource(depth) Then
                    expectSource(depth) = False
                    depth = depth + 1
                    expectSource(depth) = True
                    inFrom(depth) = True
                Else
                    depth = depth + 1
                End If
            ElseIf symbolText = ")" Then
                If depth > 0 Then
                    inFrom(depth) = False
                    expectSource(depth) = False
                    depth = depth - 1
                End If
            ElseIf symbolText = "," Then
                If inFrom(depth) Then expectSource(depth) = True
            End If
        ElseIf StrComp(tokenType, "Identifier", vbBinaryCompare) = 0 Then
            tokenText = CStr(tokenData(3))
            If expectSource(depth) Then
                If SqlTokenIsKeyword(tokenData, "SELECT") Or SqlTokenIsKeyword(tokenData, "TRANSFORM") Then
                    expectSource(depth) = False
                    inFrom(depth) = False
                Else
                    aliasName = vbNullString
                    aliasIndex = 0
                    isExternal = False
                    GetSqlSourceAliasDetails tokens, i, aliasName, aliasIndex, isExternal
                    If Not isExternal Then
                        localSourceCount = localSourceCount + 1
                        sourceTokenIndexes.Add CStr(i), True
                        If aliasIndex > 0 Then aliasTokenIndexes.Add CStr(aliasIndex), True
                        If Len(aliasName) > 0 Then
                            qualifierText = aliasName
                        Else
                            qualifierText = tokenText
                        End If
                        If qualifierOwners.Exists(qualifierText) Then
                            ownerText = CStr(qualifierOwners(qualifierText))
                            If StrComp(ownerText, tokenText, vbTextCompare) <> 0 Then
                                problem = "Source qualifier '" & qualifierText & "' is owned by both '" & ownerText & _
                                    "' and '" & tokenText & "' within one SQL branch."
                                Exit Function
                            End If
                        Else
                            qualifierOwners.Add qualifierText, tokenText
                        End If
                        If StrComp(tokenText, sourceQueryName, vbTextCompare) = 0 Then
                            sourceFound = True
                            If Not sourceQualifiers.Exists(qualifierText) Then sourceQualifiers.Add qualifierText, True
                        End If
                    End If
                    expectSource(depth) = False
                    inFrom(depth) = True
                End If
            ElseIf SqlTokenIsKeyword(tokenData, "FROM") Or SqlTokenIsKeyword(tokenData, "JOIN") Or _
                   SqlTokenIsKeyword(tokenData, "UPDATE") Or SqlTokenIsKeyword(tokenData, "INTO") Or _
                   SqlTokenIsKeyword(tokenData, "TABLE") Then
                expectSource(depth) = True
                inFrom(depth) = True
            ElseIf SqlTokenEndsSourceClause(tokenData) Then
                expectSource(depth) = False
                inFrom(depth) = False
            End If
        End If
    Next i
    soleSource = (sourceFound And localSourceCount = 1)
    CollectSqlSourceQualifiers = True
End Function

Private Sub GetSqlSourceAliasDetails( _
    ByVal tokens As Collection, _
    ByVal sourceIndex As Long, _
    ByRef aliasName As String, _
    ByRef aliasIndex As Long, _
    ByRef isExternal As Boolean)
    Dim tokenData As Variant
    Dim nextData As Variant
    Dim j As Long
    aliasName = vbNullString
    aliasIndex = 0
    isExternal = False
    j = sourceIndex + 1
    If j > tokens.count Then Exit Sub
    tokenData = tokens(j)
    If SqlTokenIsKeyword(tokenData, "IN") Then
        isExternal = True
        Exit Sub
    End If
    If SqlTokenIsKeyword(tokenData, "AS") Then
        If j + 1 <= tokens.count Then
            nextData = tokens(j + 1)
            If StrComp(CStr(nextData(0)), "Identifier", vbBinaryCompare) = 0 Then
                aliasName = CStr(nextData(3))
                aliasIndex = j + 1
                j = j + 2
            End If
        End If
    ElseIf StrComp(CStr(tokenData(0)), "Identifier", vbBinaryCompare) = 0 Then
        If CBool(tokenData(4)) Or Not SqlSourceAliasReserved(CStr(tokenData(3))) Then
            aliasName = CStr(tokenData(3))
            aliasIndex = j
            j = j + 1
        End If
    End If
    If j <= tokens.count Then
        tokenData = tokens(j)
        If SqlTokenIsKeyword(tokenData, "IN") Then isExternal = True
    End If
End Sub

Private Function SqlStatementBodyStartPosition(ByVal sqlText As String) As Long
    Dim p As Long
    Dim candidate As Long
    Dim keywords As Variant
    Dim keyword As Variant
    keywords = Array("TRANSFORM", "SELECT", "UPDATE", "INSERT", "DELETE")
    For Each keyword In keywords
        candidate = FindTopLevelKeyword(sqlText, CStr(keyword), 1)
        If candidate > 0 Then
            If p = 0 Or candidate < p Then p = candidate
        End If
    Next keyword
    If p = 0 Then p = 1
    SqlStatementBodyStartPosition = p
End Function

Private Sub PreflightQueryObjectReferenceRewrites(ByVal plan As Collection, ByRef hardIssues As Collection)
    ' Prove QueryDef source-reference rewrites are unambiguous before any physical object or field rename begins.
    Dim db As Object
    Dim qdf As Object
    Dim item As Object
    Dim i As Long
    Dim typeName As String
    Dim oldName As String
    Dim newName As String
    Dim oldSql As String
    Dim rewrittenSql As String
    Dim changed As Boolean
    Dim problem As String
    Set db = currentDb
    For i = 1 To plan.count
        Set item = plan(i)
        If StrComp(CStr(item("Action")), "Pending", vbTextCompare) = 0 Then
            typeName = CStr(item("Type"))
            If StrComp(typeName, "Table", vbTextCompare) = 0 Or StrComp(typeName, "Query", vbTextCompare) = 0 Then
                oldName = CStr(item("WorkOld"))
                newName = CStr(item("New"))
                For Each qdf In db.QueryDefs
                    If Left$(CStr(qdf.Name), 1) <> "~" And Len(CStr(qdf.Connect)) = 0 Then
                        oldSql = CStr(qdf.sql)
                        If InStr(1, oldSql, oldName, vbTextCompare) > 0 Then
                            rewrittenSql = vbNullString
                            changed = False
                            problem = vbNullString
                            If Not RewriteQueryObjectReferenceSql(oldSql, oldName, newName, rewrittenSql, changed, problem) Then
                                AddIssue hardIssues, "Query '" & CStr(qdf.Name) & "' cannot safely track object rename '" & oldName & _
                                    "' -> '" & newName & "': " & problem
                            End If
                        End If
                    End If
                Next qdf
            End If
        End If
    Next i
End Sub

Private Sub RewriteQueryReferencesAfterObjectRename( _
    ByVal sourceType As String, _
    ByVal oldName As String, _
    ByVal newName As String, _
    ByRef journal As Collection)
    ' Rewrite every affected saved local QueryDef before validating any of them so dependent-query refreshes see a consistent source graph.
    Dim db As Object
    Dim qdf As Object
    Dim queryNames As New Collection
    Dim validateNames As New Collection
    Dim queryName As String
    Dim oldSql As String
    Dim newSql As String
    Dim changed As Boolean
    Dim problem As String
    Dim refreshErr As Long
    Dim refreshDescription As String
    Dim directReferenceStillPresent As Boolean
    Dim i As Long
    Set db = currentDb
    For Each qdf In db.QueryDefs
        If Left$(CStr(qdf.Name), 1) <> "~" Then queryNames.Add CStr(qdf.Name)
    Next qdf
    For i = 1 To queryNames.count
        queryName = CStr(queryNames(i))
        Set qdf = db.QueryDefs(queryName)
        If Len(CStr(qdf.Connect)) = 0 Then
            oldSql = CStr(qdf.sql)
            If InStr(1, oldSql, oldName, vbTextCompare) > 0 Then
                validateNames.Add queryName
                newSql = vbNullString
                changed = False
                problem = vbNullString
                If Not RewriteQueryObjectReferenceSql(oldSql, oldName, newName, newSql, changed, problem) Then
                    Err.Raise vbObjectError + 7161, "RewriteQueryReferencesAfterObjectRename", _
                        "Query '" & queryName & "' cannot safely track object rename '" & oldName & "' -> '" & newName & "': " & problem
                End If
                If changed Then
                    CloseObjectIfOpen "Query", queryName
                    journal.Add Array("QuerySourceSQL", queryName, oldSql, newSql, sourceType, oldName, newName)
                    qdf.sql = newSql
                    ProgressAppend "UPDATED Query source reference: " & queryName & " | " & oldName & " -> " & newName
                End If
            End If
        End If
    Next i
    For i = 1 To validateNames.count
        queryName = CStr(validateNames(i))
        Set qdf = db.QueryDefs(queryName)
        refreshErr = 0
        refreshDescription = vbNullString
        On Error Resume Next
        qdf.fields.Refresh
        refreshErr = Err.Number
        refreshDescription = Err.Description
        Err.Clear
        On Error GoTo 0
        If refreshErr <> 0 Then
            If refreshErr = 3078 And InStr(1, refreshDescription, "'" & oldName & "'", vbTextCompare) > 0 Then
                directReferenceStillPresent = (InStr(1, CStr(qdf.sql), oldName, vbTextCompare) > 0)
                If directReferenceStillPresent Then
                    Err.Raise refreshErr, "RewriteQueryReferencesAfterObjectRename", _
                        "Query '" & queryName & "' still contains the renamed input text '" & oldName & _
                        "' after all affected QueryDefs were rewritten for object rename '" & oldName & "' -> '" & newName & "': " & refreshDescription
                Else
                    Err.Raise refreshErr, "RewriteQueryReferencesAfterObjectRename", _
                        "Query '" & queryName & "' still resolves the renamed input '" & oldName & _
                        "' through its dependency graph after all affected QueryDefs were rewritten for object rename '" & oldName & "' -> '" & newName & "': " & refreshDescription
                End If
            Else
                Err.Raise refreshErr, "RewriteQueryReferencesAfterObjectRename", _
                    "Query '" & queryName & "' is invalid after object rename '" & oldName & "' -> '" & newName & _
                    "': " & refreshDescription
            End If
        End If
    Next i
End Sub

Private Function RewriteQueryObjectReferenceSql( _
    ByVal sqlText As String, _
    ByVal oldName As String, _
    ByVal newName As String, _
    ByRef rewrittenSql As String, _
    ByRef changed As Boolean, _
    ByRef problem As String) As Boolean
    ' Replace only local source-object declarations and their unaliased qualifiers; quoted literals and external IN sources are preserved.
    Dim tokens As Collection
    Dim sourceIndexes As Object
    Dim qualifierIndexes As Object
    Dim replacements As Object
    Dim tokenData As Variant
    Dim key As Variant
    Dim result As String
    Dim cursor As Long
    Dim tokenIndex As Long
    Dim startPos As Long
    Dim endPos As Long
    changed = False
    rewrittenSql = sqlText
    If Len(oldName) = 0 Or Len(newName) = 0 Then
        problem = "Object rename contains a blank source or target name."
        Exit Function
    End If
    If Not TokenizeSqlObjectReferenceText(sqlText, tokens, problem) Then Exit Function
    If Not AnalyzeSqlObjectReferenceTokens(tokens, oldName, newName, sourceIndexes, qualifierIndexes, problem) Then Exit Function
    Set replacements = CreateTextDictionary()
    For Each key In sourceIndexes.keys
        replacements.Add CStr(key), True
    Next key
    For Each key In qualifierIndexes.keys
        If Not replacements.Exists(CStr(key)) Then replacements.Add CStr(key), True
    Next key
    If replacements.count = 0 Then
        RewriteQueryObjectReferenceSql = True
        Exit Function
    End If
    cursor = 1
    For tokenIndex = 1 To tokens.count
        If replacements.Exists(CStr(tokenIndex)) Then
            tokenData = tokens(tokenIndex)
            startPos = CLng(tokenData(1))
            endPos = CLng(tokenData(2))
            result = result & Mid$(sqlText, cursor, startPos - cursor) & "[" & Replace(newName, "]", "]]") & "]"
            cursor = endPos + 1
        End If
    Next tokenIndex
    result = result & Mid$(sqlText, cursor)
    rewrittenSql = result
    changed = True
    RewriteQueryObjectReferenceSql = True
End Function

Private Function TokenizeSqlObjectReferenceText( _
    ByVal sqlText As String, _
    ByRef tokens As Collection, _
    ByRef problem As String) As Boolean
    ' Tokenize identifiers and punctuation while treating quoted text and date literals as opaque content.
    Dim ch As String
    Dim quoteChar As String
    Dim token As String
    Dim i As Long
    Dim j As Long
    Set tokens = New Collection
    i = 1
    Do While i <= Len(sqlText)
        ch = Mid$(sqlText, i, 1)
        If IsSqlSpace(ch) Then
            i = i + 1
        ElseIf ch = "'" Or ch = """" Then
            quoteChar = ch
            j = i + 1
            Do While j <= Len(sqlText)
                If Mid$(sqlText, j, 1) = quoteChar Then
                    If j < Len(sqlText) And Mid$(sqlText, j + 1, 1) = quoteChar Then
                        j = j + 2
                    Else
                        Exit Do
                    End If
                Else
                    j = j + 1
                End If
            Loop
            If j > Len(sqlText) Then
                problem = "SQL contains an unterminated quoted literal."
                Exit Function
            End If
            tokens.Add Array("Literal", i, j, Mid$(sqlText, i, j - i + 1), False)
            i = j + 1
        ElseIf ch = "#" Then
            j = i + 1
            Do While j <= Len(sqlText) And Mid$(sqlText, j, 1) <> "#"
                j = j + 1
            Loop
            If j > Len(sqlText) Then
                problem = "SQL contains an unterminated date literal."
                Exit Function
            End If
            tokens.Add Array("Literal", i, j, Mid$(sqlText, i, j - i + 1), False)
            i = j + 1
        ElseIf ch = "[" Then
            token = vbNullString
            j = i + 1
            Do While j <= Len(sqlText)
                If Mid$(sqlText, j, 1) = "]" Then
                    If j < Len(sqlText) And Mid$(sqlText, j + 1, 1) = "]" Then
                        token = token & "]"
                        j = j + 2
                    Else
                        Exit Do
                    End If
                Else
                    token = token & Mid$(sqlText, j, 1)
                    j = j + 1
                End If
            Loop
            If j > Len(sqlText) Then
                problem = "SQL contains an unterminated bracketed identifier."
                Exit Function
            End If
            tokens.Add Array("Identifier", i, j, token, True)
            i = j + 1
        ElseIf IsSqlWordChar(ch) Then
            j = i
            Do While j <= Len(sqlText) And IsSqlWordChar(Mid$(sqlText, j, 1))
                j = j + 1
            Loop
            tokens.Add Array("Identifier", i, j - 1, Mid$(sqlText, i, j - i), False)
            i = j
        Else
            tokens.Add Array("Symbol", i, i, ch, False)
            i = i + 1
        End If
    Loop
    TokenizeSqlObjectReferenceText = True
End Function

Private Function AnalyzeSqlObjectReferenceTokens( _
    ByVal tokens As Collection, _
    ByVal oldName As String, _
    ByVal newName As String, _
    ByRef sourceIndexes As Object, _
    ByRef qualifierIndexes As Object, _
    ByRef problem As String) As Boolean
    ' Identify exact local source declarations and determine whether bare source-name qualifiers can be rewritten without alias ambiguity.
    Dim inFrom(0 To 255) As Boolean
    Dim expectSource(0 To 255) As Boolean
    Dim sourceParen(0 To 255) As Boolean
    Dim tokenData As Variant
    Dim aliasName As String
    Dim tokenText As String
    Dim tokenType As String
    Dim symbolText As String
    Dim depth As Long
    Dim i As Long
    Dim noAliasSource As Boolean
    Dim oldNameUsedAsAlias As Boolean
    Dim newNameUsedAsAlias As Boolean
    Dim isExternal As Boolean
    Set sourceIndexes = CreateTextDictionary()
    Set qualifierIndexes = CreateTextDictionary()
    AddSqlDdlObjectReferenceIndexes tokens, oldName, sourceIndexes
    For i = 1 To tokens.count
        tokenData = tokens(i)
        tokenType = CStr(tokenData(0))
        If StrComp(tokenType, "Symbol", vbBinaryCompare) = 0 Then
            symbolText = CStr(tokenData(3))
            If symbolText = "(" Then
                If depth = 255 Then
                    problem = "SQL nesting exceeds the supported reference-rewrite depth."
                    Exit Function
                End If
                If expectSource(depth) Then
                    expectSource(depth) = False
                    depth = depth + 1
                    expectSource(depth) = True
                    inFrom(depth) = True
                    sourceParen(depth) = True
                Else
                    depth = depth + 1
                End If
            ElseIf symbolText = ")" Then
                If depth > 0 Then
                    If sourceParen(depth) Then
                        aliasName = SqlAliasAfterSourceExpression(tokens, i)
                        If Len(aliasName) > 0 Then
                            If StrComp(aliasName, oldName, vbTextCompare) = 0 Then oldNameUsedAsAlias = True
                            If StrComp(aliasName, newName, vbTextCompare) = 0 Then newNameUsedAsAlias = True
                        End If
                    End If
                    inFrom(depth) = False
                    expectSource(depth) = False
                    sourceParen(depth) = False
                    depth = depth - 1
                End If
            ElseIf symbolText = "," Then
                If inFrom(depth) Then expectSource(depth) = True
            End If
        ElseIf StrComp(tokenType, "Identifier", vbBinaryCompare) = 0 Then
            tokenText = CStr(tokenData(3))
            If expectSource(depth) Then
                If SqlTokenIsKeyword(tokenData, "SELECT") Or SqlTokenIsKeyword(tokenData, "TRANSFORM") Then
                    expectSource(depth) = False
                    inFrom(depth) = False
                Else
                    aliasName = vbNullString
                    isExternal = False
                    GetSqlSourceAliasAndExternal tokens, i, aliasName, isExternal
                    If Len(aliasName) > 0 Then
                        If StrComp(aliasName, oldName, vbTextCompare) = 0 Then oldNameUsedAsAlias = True
                        If StrComp(aliasName, newName, vbTextCompare) = 0 Then newNameUsedAsAlias = True
                    End If
                    If Not isExternal Then
                        If StrComp(tokenText, oldName, vbTextCompare) = 0 Then
                            sourceIndexes.Add CStr(i), True
                            If Len(aliasName) = 0 Then noAliasSource = True
                        End If
                    End If
                    expectSource(depth) = False
                    inFrom(depth) = True
                End If
            ElseIf SqlTokenIsKeyword(tokenData, "FROM") Or SqlTokenIsKeyword(tokenData, "JOIN") Or _
                   SqlTokenIsKeyword(tokenData, "UPDATE") Or SqlTokenIsKeyword(tokenData, "INTO") Or _
                   SqlTokenIsKeyword(tokenData, "TABLE") Then
                expectSource(depth) = True
                inFrom(depth) = True
            ElseIf SqlTokenEndsSourceClause(tokenData) Then
                expectSource(depth) = False
                inFrom(depth) = False
            End If
        End If
    Next i
    If noAliasSource And oldNameUsedAsAlias Then
        problem = "Object name '" & oldName & "' is also used as a source alias while an unaliased source of that name exists; qualified references are ambiguous."
        Exit Function
    End If
    If noAliasSource And newNameUsedAsAlias Then
        problem = "Target object name '" & newName & "' is already used as a source alias while unaliased source '" & oldName & "' is being renamed; rewritten qualifiers would be ambiguous."
        Exit Function
    End If
    If noAliasSource Then
        For i = 1 To tokens.count - 1
            tokenData = tokens(i)
            If StrComp(CStr(tokenData(0)), "Identifier", vbBinaryCompare) = 0 Then
                If StrComp(CStr(tokenData(3)), oldName, vbTextCompare) = 0 Then
                    If Not sourceIndexes.Exists(CStr(i)) Then
                        If SqlNextTokenIsQualifier(tokens, i) Then qualifierIndexes.Add CStr(i), True
                    End If
                End If
            End If
        Next i
    End If
    AnalyzeSqlObjectReferenceTokens = True
End Function

Private Sub GetSqlSourceAliasAndExternal( _
    ByVal tokens As Collection, _
    ByVal sourceIndex As Long, _
    ByRef aliasName As String, _
    ByRef isExternal As Boolean)
    ' Inspect tokens immediately following one source declaration for an optional alias and Access IN external-database clause.
    Dim tokenData As Variant
    Dim nextData As Variant
    Dim j As Long
    aliasName = vbNullString
    isExternal = False
    j = sourceIndex + 1
    If j > tokens.count Then Exit Sub
    tokenData = tokens(j)
    If SqlTokenIsKeyword(tokenData, "IN") Then
        isExternal = True
        Exit Sub
    End If
    If SqlTokenIsKeyword(tokenData, "AS") Then
        If j + 1 <= tokens.count Then
            nextData = tokens(j + 1)
            If StrComp(CStr(nextData(0)), "Identifier", vbBinaryCompare) = 0 Then
                aliasName = CStr(nextData(3))
                j = j + 2
            End If
        End If
    ElseIf StrComp(CStr(tokenData(0)), "Identifier", vbBinaryCompare) = 0 Then
        If CBool(tokenData(4)) Or Not SqlSourceAliasReserved(CStr(tokenData(3))) Then
            aliasName = CStr(tokenData(3))
            j = j + 1
        End If
    End If
    If j <= tokens.count Then
        tokenData = tokens(j)
        If SqlTokenIsKeyword(tokenData, "IN") Then isExternal = True
    End If
End Sub

Private Function SqlAliasAfterSourceExpression(ByVal tokens As Collection, ByVal closeParenIndex As Long) As String
    ' Return an explicit or implicit alias immediately following a parenthesized source expression.
    Dim tokenData As Variant
    Dim nextData As Variant
    Dim j As Long
    j = closeParenIndex + 1
    If j > tokens.count Then Exit Function
    tokenData = tokens(j)
    If SqlTokenIsKeyword(tokenData, "AS") Then
        If j + 1 <= tokens.count Then
            nextData = tokens(j + 1)
            If StrComp(CStr(nextData(0)), "Identifier", vbBinaryCompare) = 0 Then SqlAliasAfterSourceExpression = CStr(nextData(3))
        End If
    ElseIf StrComp(CStr(tokenData(0)), "Identifier", vbBinaryCompare) = 0 Then
        If CBool(tokenData(4)) Or Not SqlSourceAliasReserved(CStr(tokenData(3))) Then
            SqlAliasAfterSourceExpression = CStr(tokenData(3))
        End If
    End If
End Function

Private Sub AddSqlDdlObjectReferenceIndexes( _
    ByVal tokens As Collection, _
    ByVal oldName As String, _
    ByRef sourceIndexes As Object)
    ' Add exact table references used by Access DDL clauses that are outside ordinary FROM/JOIN source syntax.
    Dim tokenData As Variant
    Dim nextData As Variant
    Dim tokenText As String
    Dim i As Long
    Dim depth As Long
    Dim indexState As Long
    For i = 1 To tokens.count - 1
        tokenData = tokens(i)
        If SqlTokenIsKeyword(tokenData, "REFERENCES") Then
            nextData = tokens(i + 1)
            If StrComp(CStr(nextData(0)), "Identifier", vbBinaryCompare) = 0 Then
                If StrComp(CStr(nextData(3)), oldName, vbTextCompare) = 0 Then
                    If Not sourceIndexes.Exists(CStr(i + 1)) Then sourceIndexes.Add CStr(i + 1), True
                End If
            End If
        End If
    Next i
    For i = 1 To tokens.count
        tokenData = tokens(i)
        If StrComp(CStr(tokenData(0)), "Symbol", vbBinaryCompare) = 0 Then
            If CStr(tokenData(3)) = "(" Then
                depth = depth + 1
            ElseIf CStr(tokenData(3)) = ")" Then
                If depth > 0 Then depth = depth - 1
            End If
        ElseIf depth = 0 And StrComp(CStr(tokenData(0)), "Identifier", vbBinaryCompare) = 0 Then
            tokenText = CStr(tokenData(3))
            Select Case indexState
                Case 0
                    If SqlTokenIsKeyword(tokenData, "CREATE") Then
                        indexState = 1
                    ElseIf SqlTokenIsKeyword(tokenData, "DROP") Then
                        indexState = 2
                    End If
                Case 1
                    If SqlTokenIsKeyword(tokenData, "UNIQUE") Then
                        indexState = 1
                    ElseIf SqlTokenIsKeyword(tokenData, "INDEX") Then
                        indexState = 3
                    Else
                        indexState = 0
                    End If
                Case 2
                    If SqlTokenIsKeyword(tokenData, "INDEX") Then
                        indexState = 3
                    Else
                        indexState = 0
                    End If
                Case 3
                    indexState = 4
                Case 4
                    If SqlTokenIsKeyword(tokenData, "ON") Then
                        indexState = 5
                    Else
                        indexState = 0
                    End If
                Case 5
                    If StrComp(tokenText, oldName, vbTextCompare) = 0 Then
                        If Not sourceIndexes.Exists(CStr(i)) Then sourceIndexes.Add CStr(i), True
                    End If
                    indexState = 0
            End Select
        End If
    Next i
End Sub

Private Function SqlTokenIsKeyword(ByVal tokenData As Variant, ByVal keyword As String) As Boolean
    If StrComp(CStr(tokenData(0)), "Identifier", vbBinaryCompare) <> 0 Then Exit Function
    If CBool(tokenData(4)) Then Exit Function
    SqlTokenIsKeyword = (StrComp(CStr(tokenData(3)), keyword, vbTextCompare) = 0)
End Function

Private Function SqlTokenEndsSourceClause(ByVal tokenData As Variant) As Boolean
    If SqlTokenIsKeyword(tokenData, "WHERE") Or SqlTokenIsKeyword(tokenData, "GROUP") Or _
       SqlTokenIsKeyword(tokenData, "HAVING") Or SqlTokenIsKeyword(tokenData, "ORDER") Or SqlTokenIsKeyword(tokenData, "ON") Or _
       SqlTokenIsKeyword(tokenData, "PIVOT") Or SqlTokenIsKeyword(tokenData, "UNION") Or _
       SqlTokenIsKeyword(tokenData, "SET") Or SqlTokenIsKeyword(tokenData, "VALUES") Or _
       SqlTokenIsKeyword(tokenData, "SELECT") Or SqlTokenIsKeyword(tokenData, "ADD") Or _
       SqlTokenIsKeyword(tokenData, "ALTER") Or SqlTokenIsKeyword(tokenData, "DROP") Or _
       SqlTokenIsKeyword(tokenData, "CONSTRAINT") Then
        SqlTokenEndsSourceClause = True
    End If
End Function

Private Function SqlSourceAliasReserved(ByVal tokenText As String) As Boolean
    Select Case UCase$(tokenText)
        Case "AS", "INNER", "LEFT", "RIGHT", "FULL", "OUTER", "CROSS", "JOIN", "ON", "WHERE", "GROUP", "ORDER", _
             "HAVING", "PIVOT", "UNION", "SET", "VALUES", "FROM", "INTO", "UPDATE", "INSERT", "DELETE", "SELECT", _
             "TRANSFORM", "IN", "WITH", "OWNERACCESS", "OPTION", "TABLE", "ADD", "ALTER", "DROP", "CONSTRAINT", _
             "REFERENCES", "CREATE", "INDEX", "UNIQUE"
            SqlSourceAliasReserved = True
    End Select
End Function

Private Function SqlNextTokenIsQualifier(ByVal tokens As Collection, ByVal tokenIndex As Long) As Boolean
    Dim tokenData As Variant
    If tokenIndex + 1 > tokens.count Then Exit Function
    tokenData = tokens(tokenIndex + 1)
    If StrComp(CStr(tokenData(0)), "Symbol", vbBinaryCompare) = 0 Then
        SqlNextTokenIsQualifier = (CStr(tokenData(3)) = "." Or CStr(tokenData(3)) = "!")
    End If
End Function

Private Sub PreflightCalculatedTableFields( _
    ByVal fieldPlan As Collection, _
    ByVal stats As Object, _
    ByRef hardIssues As Collection)
    ' Identify calculated table fields before live execution and reject schema shapes that cannot be migrated without losing constraints.
    Dim item As Object
    Dim i As Long
    Dim expressionText As String
    Dim problem As String
    For i = 1 To fieldPlan.count
        Set item = fieldPlan(i)
        If StrComp(CStr(item("ParentType")), "Table", vbTextCompare) = 0 And _
           StrComp(CStr(item("Action")), "Pending", vbTextCompare) = 0 Then
            expressionText = vbNullString
            If TryGetCalculatedFieldExpression(CStr(item("ParentActual")), CStr(item("Old")), expressionText) Then
                If item.Exists("Calculated") Then
                    item("Calculated") = True
                Else
                    item.Add "Calculated", True
                End If
                StatInc stats, "CalculatedTableFieldPending"
                problem = CalculatedFieldMigrationProblem(CStr(item("ParentActual")), CStr(item("Old")))
                If Len(problem) > 0 Then
                    item("Action") = "Conflict"
                    StatInc stats, "FieldConflicts"
                    AddIssue hardIssues, problem
                End If
            End If
        End If
    Next i
End Sub

Private Function CalculatedFieldMigrationProblem(ByVal tableName As String, ByVal fieldName As String) As String
    ' Deleting/recreating a calculated field is not safe when the field owns an index/relationship or belongs to a linked table.
    Dim db As Object
    Dim tdf As Object
    Dim idx As Object
    Dim idxField As Object
    Dim rel As Object
    Dim relField As Object
    On Error GoTo CheckFail
    Set db = currentDb
    Set tdf = db.TableDefs(tableName)
    If Not CBool(tdf.Updatable) Then
        CalculatedFieldMigrationProblem = "Calculated field '" & tableName & "." & fieldName & _
            "' is in a non-updatable table and cannot be migrated safely by this utility."
        Exit Function
    End If
    If Len(CStr(tdf.Connect)) > 0 Then
        CalculatedFieldMigrationProblem = "Calculated field '" & tableName & "." & fieldName & _
            "' is in a linked table and cannot be migrated safely by this utility."
        Exit Function
    End If
    For Each idx In tdf.Indexes
        For Each idxField In idx.fields
            If StrComp(CStr(idxField.Name), fieldName, vbTextCompare) = 0 Then
                CalculatedFieldMigrationProblem = "Calculated field '" & tableName & "." & fieldName & _
                    "' participates in index '" & CStr(idx.Name) & "' and cannot be delete/recreated safely."
                Exit Function
            End If
        Next idxField
    Next idx
    For Each rel In db.Relations
        For Each relField In rel.fields
            If StrComp(CStr(rel.Table), tableName, vbTextCompare) = 0 Then
                If StrComp(CStr(relField.Name), fieldName, vbTextCompare) = 0 Then
                    CalculatedFieldMigrationProblem = "Calculated field '" & tableName & "." & fieldName & _
                        "' participates in relationship '" & CStr(rel.Name) & "' and cannot be delete/recreated safely."
                    Exit Function
                End If
            End If
            If StrComp(CStr(rel.foreignTable), tableName, vbTextCompare) = 0 Then
                If StrComp(CStr(relField.ForeignName), fieldName, vbTextCompare) = 0 Then
                    CalculatedFieldMigrationProblem = "Calculated field '" & tableName & "." & fieldName & _
                        "' participates in relationship '" & CStr(rel.Name) & "' and cannot be delete/recreated safely."
                    Exit Function
                End If
            End If
        Next relField
    Next rel
    Exit Function
CheckFail:
    CalculatedFieldMigrationProblem = "Calculated field '" & tableName & "." & fieldName & _
        "' could not be inspected for safe migration: Error " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Sub PreflightCalculatedFieldNameAutoCorrect( _
    ByVal stats As Object, _
    ByRef hardIssues As Collection, _
    ByVal nacTrack As Variant, _
    ByVal nacPerform As Variant, _
    ByVal nacKnown As Boolean)
    ' The ordinary placeholder rename must be a real Name AutoCorrect event so dependent Access objects follow the calculated-field rename.
    If StatValue(stats, "CalculatedTableFieldPending") = 0 Then Exit Sub
    If Not nacKnown Then
        AddIssue hardIssues, "Calculated table field renames require readable Name AutoCorrect state with Track and Perform enabled."
    ElseIf (Not CBool(nacTrack)) Or (Not CBool(nacPerform)) Then
        AddIssue hardIssues, "Calculated table field renames require Track Name AutoCorrect Info and Perform Name AutoCorrect to be ON."
    End If
End Sub

Private Function CaptureLocalQuerySqlSnapshot() As Object
    ' Capture the exact post-query-output/pre-table-field QueryDef SQL state so rollback never depends on Name AutoCorrect reversing its own edits.
    Dim snapshot As Object
    Dim db As Object
    Dim qdf As Object
    Set snapshot = CreateTextDictionary()
    Set db = currentDb
    For Each qdf In db.QueryDefs
        If Left$(CStr(qdf.Name), 1) <> "~" And Len(CStr(qdf.Connect)) = 0 Then
            snapshot.Add CStr(qdf.Name), CStr(qdf.sql)
        End If
    Next qdf
    Set CaptureLocalQuerySqlSnapshot = snapshot
End Function

Private Function BuildAppliedTableFieldRenameMaps( _
    ByVal journal As Collection, _
    ByVal firstIndex As Long, _
    ByVal lastIndex As Long) As Object
    ' Collapse the actual physical field-rename history, including cycle-break names, into source-table old/intermediate-to-final maps.
    Dim maps As Object
    Dim renameMap As Object
    Dim stepData As Variant
    Dim tableName As String
    Dim oldName As String
    Dim newName As String
    Dim finalName As String
    Dim i As Long
    Set maps = CreateTextDictionary()
    For i = lastIndex To firstIndex Step -1
        stepData = journal(i)
        If StrComp(CStr(stepData(0)), "TableField", vbTextCompare) = 0 Then
            tableName = CStr(stepData(1))
            If maps.Exists(tableName) Then
                Set renameMap = maps(tableName)
            Else
                Set renameMap = CreateTextDictionary()
                maps.Add tableName, renameMap
            End If
            oldName = CStr(stepData(2))
            newName = CStr(stepData(3))
            finalName = newName
            If renameMap.Exists(newName) Then finalName = CStr(renameMap(newName))
            renameMap(oldName) = finalName
        End If
    Next i
    Set BuildAppliedTableFieldRenameMaps = maps
End Function

Private Sub RestoreQuerySqlAfterFailedTableFieldAttempt( _
    ByVal qdf As Object, _
    ByVal oldSql As String, _
    ByVal queryName As String)
    Dim errNumber As Long
    Dim errDescription As String
    On Error Resume Next
    Err.Clear
    qdf.sql = oldSql
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If errNumber <> 0 Then
        Err.Raise errNumber, "RewriteDependentTableFieldReferences", _
            "Table-field dependency rewrite failure could not restore query '" & queryName & "': " & errDescription
    End If
End Sub

Private Sub RewriteDependentTableFieldReferences( _
    ByRef journal As Collection, _
    ByVal tableJournalStart As Long, _
    ByVal tableJournalEnd As Long, _
    ByVal querySqlBefore As Object, _
    ByRef hardIssues As Collection)
    ' Rebuild the entire local QueryDef target graph from pre-table snapshots, then validate only after the target graph is stable.
    Dim db As Object
    Dim qdf As Object
    Dim tableMaps As Object
    Dim renameMap As Object
    Dim unavailableQueries As Object
    Dim targetSqlByQuery As Object
    Dim sourceReferencedByQuery As Object
    Dim loggedAssignments As Object
    Dim queryNames As New Collection
    Dim key As Variant
    Dim tableName As String
    Dim queryName As String
    Dim originalSql As String
    Dim currentSql As String
    Dim rebuiltSql As String
    Dim newSql As String
    Dim problem As String
    Dim changed As Boolean
    Dim sourceReferenced As Boolean
    Dim anySourceReferenced As Boolean
    Dim errNumber As Long
    Dim errDescription As String
    Dim failures As Long
    Dim dependentsRebuilt As Long
    Dim snapshotStatesRestored As Long
    Dim blockedDownstreamPaths As Long
    Dim assignmentsThisPass As Long
    Dim mismatches As Long
    Dim maxPasses As Long
    Dim stabilizationPasses As Long
    Dim pass As Long
    Dim i As Long
    Set tableMaps = BuildAppliedTableFieldRenameMaps(journal, tableJournalStart, tableJournalEnd)
    If tableMaps.count = 0 Then Exit Sub
    Set db = currentDb
    Set unavailableQueries = CreateTextDictionary()
    Set targetSqlByQuery = CreateTextDictionary()
    Set sourceReferencedByQuery = CreateTextDictionary()
    Set loggedAssignments = CreateTextDictionary()
    For Each qdf In db.QueryDefs
        If Left$(CStr(qdf.Name), 1) <> "~" And Len(CStr(qdf.Connect)) = 0 Then queryNames.Add CStr(qdf.Name)
    Next qdf
    For i = 1 To queryNames.count
        queryName = CStr(queryNames(i))
        If Not unavailableQueries.Exists(queryName) Then
            If Not querySqlBefore.Exists(queryName) Then
                Err.Raise vbObjectError + 7176, "RewriteDependentTableFieldReferences", _
                    "Pre-table-field QueryDef SQL snapshot is missing query '" & queryName & "'."
            End If
            originalSql = CStr(querySqlBefore(queryName))
            rebuiltSql = originalSql
            anySourceReferenced = False
            For Each key In tableMaps.keys
                tableName = CStr(key)
                If InStr(1, rebuiltSql, tableName, vbTextCompare) > 0 Then
                    Set renameMap = tableMaps(tableName)
                    newSql = vbNullString
                    changed = False
                    sourceReferenced = False
                    problem = vbNullString
                    If Not RewriteQueryFieldReferenceMapSql( _
                        rebuiltSql, tableName, renameMap, newSql, changed, sourceReferenced, problem) Then
                        failures = failures + 1
                        unavailableQueries.Add queryName, queryName
                        AddIssue hardIssues, "TABLE FIELD REWRITE FAIL: Query '" & queryName & _
                            "' cannot safely track field renames from table '" & tableName & "': " & problem
                        PropagateUnavailableQueryPaths queryNames, unavailableQueries, hardIssues, failures
                        Exit For
                    End If
                    If sourceReferenced Then
                        dependentsRebuilt = dependentsRebuilt + 1
                        anySourceReferenced = True
                        rebuiltSql = newSql
                    End If
                End If
            Next key
            If Not unavailableQueries.Exists(queryName) Then
                targetSqlByQuery.Add queryName, rebuiltSql
                sourceReferencedByQuery.Add queryName, anySourceReferenced
            End If
        End If
    Next i
    maxPasses = queryNames.count + 1
    For pass = 1 To maxPasses
        stabilizationPasses = pass
        assignmentsThisPass = 0
        For i = 1 To queryNames.count
            queryName = CStr(queryNames(i))
            If targetSqlByQuery.Exists(queryName) And Not unavailableQueries.Exists(queryName) Then
                Set qdf = db.QueryDefs(queryName)
                currentSql = CStr(qdf.sql)
                rebuiltSql = CStr(targetSqlByQuery(queryName))
                If StrComp(currentSql, rebuiltSql, vbBinaryCompare) <> 0 Then
                    CloseObjectIfOpen "Query", queryName
                    On Error Resume Next
                    Err.Clear
                    qdf.sql = rebuiltSql
                    errNumber = Err.Number
                    errDescription = Err.Description
                    Err.Clear
                    On Error GoTo 0
                    If errNumber = 0 Then
                        assignmentsThisPass = assignmentsThisPass + 1
                        If Not loggedAssignments.Exists(queryName) Then
                            loggedAssignments.Add queryName, True
                            If CBool(sourceReferencedByQuery(queryName)) Then
                                ProgressAppend "REBUILT Query table-field references from snapshot: " & queryName
                            Else
                                snapshotStatesRestored = snapshotStatesRestored + 1
                                ProgressAppend "RESTORED Query SQL changed only by Name AutoCorrect: " & queryName
                            End If
                        End If
                    Else
                        failures = failures + 1
                        RestoreQuerySqlAfterFailedTableFieldAttempt qdf, currentSql, queryName
                        unavailableQueries.Add queryName, queryName
                        If CBool(sourceReferencedByQuery(queryName)) Then
                            AddIssue hardIssues, "TABLE FIELD REWRITE FAIL: Dependent query '" & queryName & _
                                "' could not accept its reconstructed SQL: Error " & CStr(errNumber) & ": " & errDescription
                        Else
                            AddIssue hardIssues, "TABLE FIELD SNAPSHOT RESTORE FAIL: Query '" & queryName & _
                                "' could not restore its pre-table SQL after Name AutoCorrect changes: Error " & _
                                CStr(errNumber) & ": " & errDescription
                        End If
                        PropagateUnavailableQueryPaths queryNames, unavailableQueries, hardIssues, failures
                    End If
                End If
            End If
        Next i
        mismatches = 0
        For i = 1 To queryNames.count
            queryName = CStr(queryNames(i))
            If targetSqlByQuery.Exists(queryName) And Not unavailableQueries.Exists(queryName) Then
                Set qdf = db.QueryDefs(queryName)
                If StrComp(CStr(qdf.sql), CStr(targetSqlByQuery(queryName)), vbBinaryCompare) <> 0 Then _
                    mismatches = mismatches + 1
            End If
        Next i
        If mismatches = 0 Then Exit For
        If assignmentsThisPass = 0 Then Exit For
    Next pass
    If mismatches > 0 Then
        For i = 1 To queryNames.count
            queryName = CStr(queryNames(i))
            If targetSqlByQuery.Exists(queryName) And Not unavailableQueries.Exists(queryName) Then
                Set qdf = db.QueryDefs(queryName)
                If StrComp(CStr(qdf.sql), CStr(targetSqlByQuery(queryName)), vbBinaryCompare) <> 0 Then
                    failures = failures + 1
                    unavailableQueries.Add queryName, queryName
                    AddIssue hardIssues, "TABLE FIELD REWRITE FAIL: Query '" & queryName & _
                        "' did not stabilize at its reconstructed SQL after " & CStr(maxPasses) & " assignment pass(es)."
                    PropagateUnavailableQueryPaths queryNames, unavailableQueries, hardIssues, failures
                End If
            End If
        Next i
    End If
    blockedDownstreamPaths = 0
    For Each qdf In db.QueryDefs
        If unavailableQueries.Exists(CStr(qdf.Name)) Then
            If StrComp(CStr(qdf.Name), CStr(unavailableQueries(CStr(qdf.Name))), vbTextCompare) <> 0 Then _
                blockedDownstreamPaths = blockedDownstreamPaths + 1
        End If
    Next qdf
    ProgressAppend "TABLE FIELD REWRITE EXHAUSTIVE SUMMARY | Source tables: " & CStr(tableMaps.count) & _
        " | Direct source states rebuilt: " & CStr(dependentsRebuilt) & " | Failures: " & CStr(failures) & _
        " | Name AutoCorrect-only SQL states restored: " & CStr(snapshotStatesRestored) & _
        " | Stabilization passes: " & CStr(stabilizationPasses) & _
        " | Blocked downstream paths: " & CStr(blockedDownstreamPaths) & _
        " | Unavailable query paths: " & CStr(unavailableQueries.count)
End Sub

Private Sub ApplyTableFieldPlan( _
    ByVal fieldPlan As Collection, _
    ByRef journal As Collection, _
    ByVal stats As Object, _
    ByRef hardIssues As Collection)
    Dim pending As Object
    Dim reserved As Object
    Dim processedTables As Object
    Dim querySqlBefore As Object
    Dim cat As Object
    Dim conn As Object
    Dim item As Object
    Dim keys As Variant
    Dim key As Variant
    Dim targetKey As String
    Dim oldKey As String
    Dim tableName As String
    Dim workOld As String
    Dim newName As String
    Dim tempName As String
    Dim i As Long
    Dim initialPending As Long
    Dim completed As Long
    Dim didFinalRename As Boolean
    Dim rebuiltCalculatedField As Boolean
    Dim tableJournalStart As Long
    Dim tableJournalEnd As Long
    Set pending = CreateTextDictionary()
    Set reserved = BuildReservedFieldNameSet(fieldPlan)
    For i = 1 To fieldPlan.count
        Set item = fieldPlan(i)
        If StrComp(CStr(item("ParentType")), "Table", vbTextCompare) = 0 And _
           StrComp(CStr(item("Action")), "Pending", vbTextCompare) = 0 Then
            pending.Add FieldNamespaceKey("Table", CStr(item("ParentActual")), CStr(item("WorkOld"))), item
        End If
    Next i
    initialPending = pending.count
    If initialPending = 0 Then
        ProgressUpdate "Table fields", 55, "No table field renames are required.", True
        Exit Sub
    End If
    Set conn = CurrentProject.Connection
    Set cat = CreateObject("ADOX.Catalog")
    Set cat.ActiveConnection = conn
    cat.tables.Refresh
    Set querySqlBefore = CaptureLocalQuerySqlSnapshot()
    For Each key In querySqlBefore.keys
        journal.Add Array("TableFieldQuerySnapshot", CStr(key), CStr(querySqlBefore(key)))
    Next key
    ProgressAppend "TABLE FIELD QUERY SNAPSHOTS JOURNALED: " & CStr(querySqlBefore.count)
    tableJournalStart = journal.count + 1
    Set processedTables = CreateTextDictionary()
    keys = pending.keys
    For Each key In keys
        If pending.Exists(CStr(key)) Then
            Set item = pending(CStr(key))
            tableName = CStr(item("ParentActual"))
            If Not processedTables.Exists(tableName) Then
                processedTables.Add tableName, True
                ApplyCalculatedFieldTablePlan tableName, pending, reserved, journal, stats, initialPending, completed, cat, conn
            End If
        End If
    Next key
    Do While pending.count > 0
        If mCancelRequested Then Err.Raise vbObjectError + 7103, "ApplyTableFieldPlan", "Cancelled by user."
        didFinalRename = False
        keys = pending.keys
        For Each key In keys
            Set item = pending(CStr(key))
            tableName = CStr(item("ParentActual"))
            workOld = CStr(item("WorkOld"))
            newName = CStr(item("New"))
            targetKey = FieldNamespaceKey("Table", tableName, newName)
            If Not pending.Exists(targetKey) Then
                RenameTableField cat, conn, tableName, workOld, newName, rebuiltCalculatedField
                journal.Add Array("TableField", tableName, workOld, newName)
                If rebuiltCalculatedField Then
                    ' DAO delete/recreate of calculated fields requires the caller's reused ADOX catalog to be reopened.
                    ProgressAppend "ADOX CATALOG REOPEN AFTER CALCULATED FIELD REBUILD: " & _
                        tableName & "." & workOld & " -> " & newName
                    Set cat = Nothing
                    Set cat = CreateObject("ADOX.Catalog")
                    Set cat.ActiveConnection = conn
                    cat.tables.Refresh
                    cat.tables(tableName).Columns.Refresh
                    ProgressAppend "ADOX CATALOG REOPEN COMPLETE: " & tableName
                End If
                pending.Remove CStr(key)
                item("WorkOld") = newName
                item("Action") = "Done"
                completed = completed + 1
                StatInc stats, "FieldRenamed"
                StatInc stats, "TableFieldRenamed"
                ProgressAppend "RENAMED Table field: " & tableName & "." & workOld & " -> " & newName
                ProgressUpdate "Table fields", 28# + (27# * completed / initialPending), _
                    Format$(completed, "#,##0") & " / " & Format$(initialPending, "#,##0") & _
                    " | " & tableName & "." & workOld & " -> " & newName, False
                didFinalRename = True
                Exit For
            End If
        Next key
        If Not didFinalRename Then
            keys = pending.keys
            key = keys(LBound(keys))
            Set item = pending(CStr(key))
            tableName = CStr(item("ParentActual"))
            workOld = CStr(item("WorkOld"))
            tempName = NextTemporaryFieldName(tableName, cat, reserved)
            RenameTableField cat, conn, tableName, workOld, tempName, rebuiltCalculatedField
            journal.Add Array("TableField", tableName, workOld, tempName)
            If rebuiltCalculatedField Then
                ' DAO delete/recreate of calculated fields requires the caller's reused ADOX catalog to be reopened.
                ProgressAppend "ADOX CATALOG REOPEN AFTER CALCULATED FIELD REBUILD: " & _
                    tableName & "." & workOld & " -> " & tempName
                Set cat = Nothing
                Set cat = CreateObject("ADOX.Catalog")
                Set cat.ActiveConnection = conn
                cat.tables.Refresh
                cat.tables(tableName).Columns.Refresh
                ProgressAppend "ADOX CATALOG REOPEN COMPLETE: " & tableName
            End If
            oldKey = CStr(key)
            pending.Remove oldKey
            item("WorkOld") = tempName
            pending.Add FieldNamespaceKey("Table", tableName, tempName), item
            reserved.Add FieldNamespaceKey("Table", tableName, tempName), True
            StatInc stats, "FieldTemporaryRenames"
            ProgressAppend "FIELD CYCLE BREAK: " & tableName & "." & workOld & " -> " & tempName
            ProgressUpdate "Resolving field rename cycle", 28# + (27# * completed / initialPending), _
                tableName & "." & workOld & " -> " & tempName, True
        End If
    Loop
    tableJournalEnd = journal.count
    If tableJournalEnd >= tableJournalStart Then
        RewriteDependentTableFieldReferences journal, tableJournalStart, tableJournalEnd, querySqlBefore, hardIssues
    End If
    ValidateAllLocalQueryDefs "TABLE FIELD PHASE", hardIssues
End Sub

Private Function BuildReservedFieldNameSet(ByVal fieldPlan As Collection) As Object
    Dim reserved As Object
    Dim item As Object
    Dim i As Long
    Dim key As String
    Set reserved = CreateTextDictionary()
    For i = 1 To fieldPlan.count
        Set item = fieldPlan(i)
        If StrComp(CStr(item("ParentType")), "Table", vbTextCompare) = 0 And Len(CStr(item("ParentActual"))) > 0 Then
            key = FieldNamespaceKey("Table", CStr(item("ParentActual")), CStr(item("Old")))
            If Not reserved.Exists(key) Then reserved.Add key, True
            key = FieldNamespaceKey("Table", CStr(item("ParentActual")), CStr(item("New")))
            If Not reserved.Exists(key) Then reserved.Add key, True
        End If
    Next i
    Set BuildReservedFieldNameSet = reserved
End Function

Private Function NextTemporaryFieldName( _
    ByVal tableName As String, _
    ByVal cat As Object, _
    ByVal reserved As Object) As String
    Dim candidate As String
    Dim key As String
    Do
        mTempSerial = mTempSerial + 1
        candidate = TEMP_PREFIX & "F_" & mSessionTag & "_" & Format$(mTempSerial, "0000")
        key = FieldNamespaceKey("Table", tableName, candidate)
    Loop While ADOXColumnExists(cat, tableName, candidate) Or reserved.Exists(key)
    NextTemporaryFieldName = candidate
End Function

Private Function ADOXColumnExists(ByVal cat As Object, ByVal tableName As String, ByVal fieldName As String) As Boolean
    Dim col As Object
    On Error Resume Next
    Set col = cat.tables(tableName).Columns(fieldName)
    ADOXColumnExists = (Err.Number = 0 And Not col Is Nothing)
    Err.Clear
    On Error GoTo 0
End Function

Private Sub RenameTableField( _
    ByVal cat As Object, _
    ByVal conn As Object, _
    ByVal tableName As String, _
    ByVal oldName As String, _
    ByVal newName As String, _
    ByRef rebuiltCalculatedField As Boolean)
    Dim col As Object
    Dim errNumber As Long
    Dim errDescription As String
    Dim errSource As String
    Dim diagnosticText As String
    Dim calculatedExpression As String
    Dim rollbackProblem As String
    Dim isCalculated As Boolean
    Dim renameAttempted As Boolean
    Dim rollbackAttempted As Boolean
    rebuiltCalculatedField = False
    On Error GoTo RenameFail
    diagnosticText = TableFieldDiagnosticText(tableName, oldName)
    ProgressAppend "ATTEMPT Table field: " & tableName & "." & oldName & " -> " & newName & _
        IIf(Len(diagnosticText) > 0, " | " & diagnosticText, vbNullString)
    isCalculated = TryGetCalculatedFieldExpression(tableName, oldName, calculatedExpression)
    If isCalculated Then
        RenameCalculatedTableField cat, conn, tableName, oldName, newName, calculatedExpression
        rebuiltCalculatedField = True
        Exit Sub
    End If
    CloseObjectIfOpen "Table", tableName
    cat.tables(tableName).Columns.Refresh
    Set col = cat.tables(tableName).Columns(oldName)
    If LOG_ORDINARY_TABLE_FIELD_DIAGNOSTICS Then LogOrdinaryTableFieldDiagnostics tableName, oldName, newName, col
    On Error Resume Next
    conn.Errors.Clear
    Err.Clear
    On Error GoTo RenameFail
    renameAttempted = True
    col.Name = newName
    cat.tables(tableName).Columns.Refresh
    Exit Sub
RenameFail:
    errNumber = Err.Number
    errDescription = Err.Description
    errSource = Err.Source
    On Error Resume Next
    If Not isCalculated Then LogADOProviderErrorsFromConnection conn, "TABLE FIELD PROVIDER ERROR"
    On Error GoTo 0
    If Not isCalculated And renameAttempted Then
        rollbackProblem = vbNullString
        rollbackAttempted = False
        If TryRollbackPartialOrdinaryTableFieldRename(conn, tableName, oldName, newName, rollbackAttempted, rollbackProblem) Then
            If rollbackAttempted Then
                ProgressAppend "TABLE FIELD PARTIAL RENAME SELF-ROLLBACK COMPLETE: " & tableName & "." & newName & " -> " & oldName
            End If
        ElseIf rollbackAttempted Then
            ProgressAppend "TABLE FIELD PARTIAL RENAME SELF-ROLLBACK FAILED: " & tableName & "." & newName & " -> " & oldName & _
                IIf(Len(rollbackProblem) > 0, " | " & rollbackProblem, vbNullString)
            If Len(rollbackProblem) > 0 Then errDescription = errDescription & " Partial-rename self-rollback failed: " & rollbackProblem
        End If
    End If
    On Error Resume Next
    ProgressAppend "TABLE FIELD RENAME FAILED: " & tableName & "." & oldName & " -> " & newName & _
        " | Error " & CStr(errNumber) & ": " & errDescription & _
        IIf(Len(errSource) > 0, " | Source: " & errSource, vbNullString)
    On Error GoTo 0
    If Len(errSource) = 0 Then errSource = "RenameTableField"
    Err.Raise errNumber, errSource, errDescription
End Sub

Private Function ApplyCalculatedFieldTablePlan( _
    ByVal tableName As String, _
    ByVal pending As Object, _
    ByVal reserved As Object, _
    ByRef journal As Collection, _
    ByVal stats As Object, _
    ByVal initialPending As Long, _
    ByRef completed As Long, _
    ByRef cat As Object, _
    ByVal conn As Object) As Boolean
    Dim definitions As Collection
    Dim Definition As Collection
    Dim tableKeys As Collection
    Dim tableItems As Collection
    Dim ordinaryPending As Object
    Dim localJournal As Collection
    Dim db As Object
    Dim tdf As Object
    Dim item As Object
    Dim keys As Variant
    Dim key As Variant
    Dim localKey As Variant
    Dim stepData As Variant
    Dim targetKey As String
    Dim oldKey As String
    Dim calculatedOriginalName As String
    Dim calculatedFinalName As String
    Dim workOld As String
    Dim newName As String
    Dim tempName As String
    Dim originalExpression As String
    Dim workingExpression As String
    Dim nextExpression As String
    Dim probeExpression As String
    Dim migrationProblem As String
    Dim propertyProblem As String
    Dim diagnosticText As String
    Dim errDescription As String
    Dim errSource As String
    Dim rollbackProblem As String
    Dim rollbackDetail As String
    Dim rebuiltCalculatedField As Boolean
    Dim suspensionStarted As Boolean
    Dim calculatedNameEventApplied As Boolean
    Dim rollbackFailed As Boolean
    Dim didFinalRename As Boolean
    Dim localTemporaryRenames As Long
    Dim errNumber As Long
    Dim rollbackErr As Long
    Dim preflightErr As Long
    Dim i As Long
    On Error Resume Next
    Err.Clear
    On Error GoTo BatchFail
    Set definitions = CaptureTableCalculatedFieldDefinitions(tableName)
    preflightErr = Err.Number
    Err.Clear
    On Error GoTo BatchFail
    If preflightErr <> 0 Or definitions Is Nothing Then Exit Function
    If definitions.count = 0 Then Exit Function
    If definitions.count > 1 Then Exit Function
    Set Definition = definitions(1)
    calculatedOriginalName = CStr(Definition("Name"))
    calculatedFinalName = calculatedOriginalName
    originalExpression = CStr(Definition("Expression"))
    workingExpression = originalExpression
    Set tableKeys = New Collection
    Set tableItems = New Collection
    Set ordinaryPending = CreateTextDictionary()
    Set localJournal = New Collection
    keys = pending.keys
    For Each key In keys
        If pending.Exists(CStr(key)) Then
            Set item = pending(CStr(key))
            If StrComp(CStr(item("ParentActual")), tableName, vbTextCompare) = 0 Then
                tableKeys.Add CStr(key)
                tableItems.Add item
                If StrComp(CStr(item("WorkOld")), calculatedOriginalName, vbTextCompare) = 0 Then
                    calculatedFinalName = CStr(item("New"))
                Else
                    ordinaryPending.Add CStr(key), item
                End If
            End If
        End If
    Next key
    migrationProblem = CalculatedFieldMigrationProblem(tableName, calculatedOriginalName)
    If Len(migrationProblem) > 0 Then Exit Function
    keys = ordinaryPending.keys
    For Each key In keys
        Set item = ordinaryPending(CStr(key))
        workOld = CStr(item("WorkOld"))
        newName = CStr(item("New"))
        If StrComp(workOld, calculatedOriginalName, vbTextCompare) = 0 Or _
           StrComp(newName, calculatedOriginalName, vbTextCompare) = 0 Or _
           StrComp(workOld, calculatedFinalName, vbTextCompare) = 0 Or _
           StrComp(newName, calculatedFinalName, vbTextCompare) = 0 Then
            Exit Function
        End If
        propertyProblem = CalculatedFieldStoredPropertyReferenceProblem(tableName, Definition, workOld)
        If Len(propertyProblem) > 0 Then Exit Function
        On Error Resume Next
        Err.Clear
        probeExpression = RewriteCalculatedFieldReference(originalExpression, workOld, "__ADC_CALCULATED_FIELD_REFERENCE_PROBE__")
        preflightErr = Err.Number
        Err.Clear
        On Error GoTo BatchFail
        If preflightErr <> 0 Then Exit Function
    Next key
    If StrComp(calculatedOriginalName, calculatedFinalName, vbBinaryCompare) <> 0 Then
        diagnosticText = TableFieldDiagnosticText(tableName, calculatedOriginalName)
        ProgressAppend "ATTEMPT Table field: " & tableName & "." & calculatedOriginalName & " -> " & calculatedFinalName & _
            IIf(Len(diagnosticText) > 0, " | " & diagnosticText, vbNullString)
    End If
    ProgressAppend "TABLE CALCULATED FIELD SUSPEND BEGIN: " & tableName & "." & calculatedOriginalName & _
        " | Pending ordinary fields=" & CStr(ordinaryPending.count)
    CloseObjectIfOpen "Table", tableName
    Set db = currentDb
    Set tdf = db.TableDefs(tableName)
    suspensionStarted = True
    DeleteCapturedCalculatedFields tdf, definitions
    EnsureSuspendedCalculatedFieldPlaceholder tableName, Definition, calculatedOriginalName
    Set cat = Nothing
    Set cat = CreateObject("ADOX.Catalog")
    Set cat.ActiveConnection = conn
    cat.tables.Refresh
    cat.tables(tableName).Columns.Refresh
    Do While ordinaryPending.count > 0
        If mCancelRequested Then Err.Raise vbObjectError + 7103, "ApplyCalculatedFieldTablePlan", "Cancelled by user."
        didFinalRename = False
        keys = ordinaryPending.keys
        For Each localKey In keys
            Set item = ordinaryPending(CStr(localKey))
            workOld = CStr(item("WorkOld"))
            newName = CStr(item("New"))
            targetKey = FieldNamespaceKey("Table", tableName, newName)
            If Not ordinaryPending.Exists(targetKey) Then
                nextExpression = RewriteCalculatedFieldReference(workingExpression, workOld, newName)
                RenameTableField cat, conn, tableName, workOld, newName, rebuiltCalculatedField
                If rebuiltCalculatedField Then Err.Raise vbObjectError + 7164, "ApplyCalculatedFieldTablePlan", _
                    "Calculated field unexpectedly existed during table-scoped ordinary-field rename."
                localJournal.Add Array("TableField", tableName, workOld, newName)
                workingExpression = nextExpression
                ordinaryPending.Remove CStr(localKey)
                item("WorkOld") = newName
                didFinalRename = True
                Exit For
            End If
        Next localKey
        If Not didFinalRename Then
            keys = ordinaryPending.keys
            localKey = keys(LBound(keys))
            Set item = ordinaryPending(CStr(localKey))
            workOld = CStr(item("WorkOld"))
            tempName = NextTemporaryFieldName(tableName, cat, reserved)
            nextExpression = RewriteCalculatedFieldReference(workingExpression, workOld, tempName)
            RenameTableField cat, conn, tableName, workOld, tempName, rebuiltCalculatedField
            If rebuiltCalculatedField Then Err.Raise vbObjectError + 7164, "ApplyCalculatedFieldTablePlan", _
                "Calculated field unexpectedly existed during table-scoped cycle break."
            localJournal.Add Array("TableField", tableName, workOld, tempName)
            workingExpression = nextExpression
            oldKey = CStr(localKey)
            ordinaryPending.Remove oldKey
            item("WorkOld") = tempName
            ordinaryPending.Add FieldNamespaceKey("Table", tableName, tempName), item
            reserved.Add FieldNamespaceKey("Table", tableName, tempName), True
            localTemporaryRenames = localTemporaryRenames + 1
            ProgressAppend "FIELD CYCLE BREAK: " & tableName & "." & workOld & " -> " & tempName
        End If
    Loop
    RestoreSuspendedCalculatedField tableName, conn, cat, Definition, calculatedOriginalName, calculatedFinalName, _
        workingExpression, calculatedNameEventApplied
    If StrComp(calculatedOriginalName, calculatedFinalName, vbBinaryCompare) <> 0 Then
        journal.Add Array("TableField", tableName, calculatedOriginalName, calculatedFinalName)
    End If
    For i = 1 To localJournal.count
        stepData = localJournal(i)
        journal.Add stepData
    Next i
    suspensionStarted = False
    For i = 1 To tableKeys.count
        If pending.Exists(CStr(tableKeys(i))) Then pending.Remove CStr(tableKeys(i))
    Next i
    For i = 1 To tableItems.count
        Set item = tableItems(i)
        workOld = CStr(item("Old"))
        newName = CStr(item("New"))
        item("WorkOld") = newName
        item("Action") = "Done"
        completed = completed + 1
        StatInc stats, "FieldRenamed"
        StatInc stats, "TableFieldRenamed"
        ProgressAppend "RENAMED Table field: " & tableName & "." & workOld & " -> " & newName
        ProgressUpdate "Table fields", 28# + (27# * completed / initialPending), _
            Format$(completed, "#,##0") & " / " & Format$(initialPending, "#,##0") & _
            " | " & tableName & "." & workOld & " -> " & newName, False
    Next i
    For i = 1 To localTemporaryRenames
        StatInc stats, "FieldTemporaryRenames"
    Next i
    ProgressAppend "TABLE CALCULATED FIELD SUSPEND COMPLETE: " & tableName & "." & calculatedFinalName
    ApplyCalculatedFieldTablePlan = True
    Exit Function
BatchFail:
    errNumber = Err.Number
    errDescription = Err.Description
    errSource = Err.Source
    On Error Resume Next
    ProgressAppend "TABLE CALCULATED FIELD BATCH FAILED: " & tableName & " | Error " & CStr(errNumber) & ": " & errDescription & _
        IIf(Len(errSource) > 0, " | Source: " & errSource, vbNullString)
    On Error GoTo 0
    If suspensionStarted Then
        rollbackProblem = vbNullString
        If calculatedNameEventApplied Then
            workOld = calculatedFinalName
        Else
            workOld = calculatedOriginalName
        End If
        On Error Resume Next
        Set db = currentDb
        Set tdf = db.TableDefs(tableName)
        If StrComp(calculatedOriginalName, calculatedFinalName, vbTextCompare) <> 0 Then
            If StrComp(workOld, calculatedOriginalName, vbTextCompare) <> 0 And DAOFieldExists(tdf, calculatedOriginalName) Then
                tdf.fields.Delete calculatedOriginalName
            End If
            If StrComp(workOld, calculatedFinalName, vbTextCompare) <> 0 And DAOFieldExists(tdf, calculatedFinalName) Then
                tdf.fields.Delete calculatedFinalName
            End If
            tdf.fields.Refresh
        End If
        EnsureSuspendedCalculatedFieldPlaceholder tableName, Definition, workOld
        Set cat = Nothing
        Set cat = CreateObject("ADOX.Catalog")
        Set cat.ActiveConnection = conn
        cat.tables.Refresh
        cat.tables(tableName).Columns.Refresh
        Err.Clear
        On Error GoTo 0
        For i = localJournal.count To 1 Step -1
            stepData = localJournal(i)
            rollbackErr = 0
            On Error Resume Next
            Err.Clear
            nextExpression = RewriteCalculatedFieldReference(workingExpression, CStr(stepData(3)), CStr(stepData(2)))
            rollbackErr = Err.Number
            rollbackDetail = Err.Description
            Err.Clear
            On Error GoTo 0
            If rollbackErr = 0 Then
                On Error Resume Next
                Err.Clear
                RenameTableField cat, conn, tableName, CStr(stepData(3)), CStr(stepData(2)), rebuiltCalculatedField
                rollbackErr = Err.Number
                rollbackDetail = Err.Description
                Err.Clear
                On Error GoTo 0
            End If
            If rollbackErr = 0 Then
                workingExpression = nextExpression
            Else
                rollbackFailed = True
                If Len(rollbackProblem) > 0 Then rollbackProblem = rollbackProblem & " | "
                rollbackProblem = rollbackProblem & CStr(stepData(3)) & " -> " & CStr(stepData(2)) & _
                    " error " & CStr(rollbackErr) & ": " & rollbackDetail
            End If
        Next i
        calculatedNameEventApplied = False
        On Error Resume Next
        Err.Clear
        RestoreSuspendedCalculatedField tableName, conn, cat, Definition, workOld, calculatedOriginalName, _
            workingExpression, calculatedNameEventApplied
        rollbackErr = Err.Number
        rollbackDetail = Err.Description
        Err.Clear
        On Error GoTo 0
        If rollbackErr <> 0 Then
            rollbackFailed = True
            If Len(rollbackProblem) > 0 Then rollbackProblem = rollbackProblem & " | "
            rollbackProblem = rollbackProblem & "Calculated-field restore error " & CStr(rollbackErr) & ": " & rollbackDetail
        End If
        For i = 1 To tableItems.count
            Set item = tableItems(i)
            item("WorkOld") = CStr(item("Old"))
            item("Action") = "Pending"
        Next i
    End If
    If rollbackFailed Then errDescription = errDescription & " Table-scoped self-rollback failed: " & rollbackProblem
    If Len(errSource) = 0 Then errSource = "ApplyCalculatedFieldTablePlan"
    Err.Raise errNumber, errSource, errDescription
End Function

Private Function TryRollbackPartialOrdinaryTableFieldRename( _
    ByVal conn As Object, _
    ByVal tableName As String, _
    ByVal oldName As String, _
    ByVal newName As String, _
    ByRef attempted As Boolean, _
    ByRef problem As String) As Boolean
    Dim db As Object
    Dim tdf As Object
    Dim rollbackCat As Object
    Dim col As Object
    Dim oldFieldExists As Boolean
    Dim newFieldExists As Boolean
    Dim caseOnlyRename As Boolean
    Dim actualFieldName As String
    Dim fieldAtOldState As Boolean
    Dim fieldAtNewState As Boolean
    Dim renameErr As Long
    Dim renameDesc As String
    On Error GoTo RollbackFail
    attempted = False
    caseOnlyRename = (StrComp(oldName, newName, vbTextCompare) = 0 And _
        StrComp(oldName, newName, vbBinaryCompare) <> 0)
    Set db = currentDb
    Set tdf = db.TableDefs(tableName)
    If caseOnlyRename Then
        If DAOFieldExists(tdf, oldName) Then
            actualFieldName = CStr(tdf.fields(oldName).Name)
            fieldAtOldState = (StrComp(actualFieldName, oldName, vbBinaryCompare) = 0)
            fieldAtNewState = (StrComp(actualFieldName, newName, vbBinaryCompare) = 0)
        End If
    Else
        oldFieldExists = DAOFieldExists(tdf, oldName)
        newFieldExists = DAOFieldExists(tdf, newName)
        fieldAtOldState = oldFieldExists And Not newFieldExists
        fieldAtNewState = newFieldExists And Not oldFieldExists
    End If
    If fieldAtOldState Then
        TryRollbackPartialOrdinaryTableFieldRename = True
        Exit Function
    End If
    attempted = True
    If Not fieldAtNewState Then
        problem = "Ordinary field is in an ambiguous physical state after the provider error."
        Exit Function
    End If
    Set rollbackCat = CreateObject("ADOX.Catalog")
    Set rollbackCat.ActiveConnection = conn
    rollbackCat.tables.Refresh
    rollbackCat.tables(tableName).Columns.Refresh
    Set col = rollbackCat.tables(tableName).Columns(newName)
    On Error Resume Next
    conn.Errors.Clear
    Err.Clear
    On Error GoTo RollbackFail
    col.Name = oldName
    Set db = currentDb
    Set tdf = db.TableDefs(tableName)
    If caseOnlyRename Then
        If DAOFieldExists(tdf, oldName) Then
            actualFieldName = CStr(tdf.fields(oldName).Name)
            fieldAtOldState = (StrComp(actualFieldName, oldName, vbBinaryCompare) = 0)
        Else
            fieldAtOldState = False
        End If
    Else
        fieldAtOldState = DAOFieldExists(tdf, oldName) And Not DAOFieldExists(tdf, newName)
    End If
    If Not fieldAtOldState Then
        problem = "Provider reverse rename did not restore the expected original physical field state."
        Exit Function
    End If
    TryRollbackPartialOrdinaryTableFieldRename = True
    Exit Function
RollbackFail:
    renameErr = Err.Number
    renameDesc = Err.Description
    On Error Resume Next
    LogADOProviderErrorsFromConnection conn, "TABLE FIELD PARTIAL RENAME SELF-ROLLBACK PROVIDER ERROR"
    On Error GoTo 0
    problem = "Error " & CStr(renameErr) & ": " & renameDesc
End Function

Private Function CaptureTableCalculatedFieldDefinitions(ByVal tableName As String) As Collection
    Dim definitions As Collection
    Dim db As Object
    Dim tdf As Object
    Dim fld As Object
    Dim prop As Object
    Dim Definition As Collection
    Dim propertyValues As Collection
    Dim expressionValue As Variant
    Dim expressionText As String
    Dim fieldName As String
    Dim fieldType As Long
    Dim fieldSize As Long
    Dim ordinalPosition As Long
    Dim readErr As Long
    Dim readSource As String
    Dim readDesc As String
    Dim i As Long
    Set definitions = New Collection
    Set db = currentDb
    Set tdf = db.TableDefs(tableName)
    For i = 0 To tdf.fields.count - 1
        Set fld = tdf.fields(i)
        fieldName = CStr(fld.Name)
        Set prop = Nothing
        readErr = 0
        readSource = vbNullString
        readDesc = vbNullString
        On Error Resume Next
        Err.Clear
        Set prop = fld.properties("Expression")
        readErr = Err.Number
        readSource = Err.Source
        readDesc = Err.Description
        Err.Clear
        On Error GoTo 0
        If readErr <> 0 Then
            If readErr <> 3270 And readErr <> 3265 Then
                Err.Raise readErr, "CaptureTableCalculatedFieldDefinitions", _
                    "Could not inspect Expression for field '" & tableName & "." & fieldName & _
                    "'. Original source: '" & readSource & "'. " & readDesc
            End If
        ElseIf Not prop Is Nothing Then
            On Error Resume Next
            Err.Clear
            expressionValue = prop.value
            readErr = Err.Number
            readSource = Err.Source
            readDesc = Err.Description
            Err.Clear
            On Error GoTo 0
            If readErr <> 0 Then
                Err.Raise readErr, "CaptureTableCalculatedFieldDefinitions", _
                    "Could not read Expression for field '" & tableName & "." & fieldName & _
                    "'. Original source: '" & readSource & "'. " & readDesc
            End If
            If IsNull(expressionValue) Or IsEmpty(expressionValue) Then
                expressionText = vbNullString
            Else
                expressionText = CStr(expressionValue)
            End If
            If Len(expressionText) > 0 Then
                CaptureCalculatedFieldDefinition tableName, fieldName, fld, fieldType, fieldSize, ordinalPosition, propertyValues
                Set Definition = New Collection
                Definition.Add fieldName, "Name"
                Definition.Add fieldType, "Type"
                Definition.Add fieldSize, "Size"
                Definition.Add ordinalPosition, "OrdinalPosition"
                Definition.Add expressionText, "Expression"
                Definition.Add propertyValues, "Properties"
                definitions.Add Definition
            End If
        End If
    Next i
    Set CaptureTableCalculatedFieldDefinitions = definitions
End Function

Private Sub DeleteCapturedCalculatedFields(ByVal tdf As Object, ByVal definitions As Collection)
    Dim Definition As Collection
    Dim i As Long
    For i = definitions.count To 1 Step -1
        Set Definition = definitions(i)
        If DAOFieldExists(tdf, CStr(Definition("Name"))) Then tdf.fields.Delete CStr(Definition("Name"))
    Next i
    tdf.fields.Refresh
End Sub

Private Sub EnsureSuspendedCalculatedFieldPlaceholder( _
    ByVal tableName As String, _
    ByVal Definition As Collection, _
    ByVal fieldName As String)
    Dim db As Object
    Dim tdf As Object
    Dim fld As Object
    Dim expressionText As String
    Dim isCalculated As Boolean
    Set db = currentDb
    Set tdf = db.TableDefs(tableName)
    If DAOFieldExists(tdf, fieldName) Then
        isCalculated = TryGetCalculatedFieldExpression(tableName, fieldName, expressionText)
        If isCalculated Then
            tdf.fields.Delete fieldName
            tdf.fields.Refresh
        Else
            Set fld = tdf.fields(fieldName)
            If CLng(fld.Type) <> CLng(Definition("Type")) Or CLng(fld.Size) <> CLng(Definition("Size")) Then
                Err.Raise vbObjectError + 7168, "EnsureSuspendedCalculatedFieldPlaceholder", _
                    "Existing calculated-field placeholder has unexpected type or size: '" & tableName & "." & fieldName & "'."
            End If
            Exit Sub
        End If
    End If
    AppendOrdinaryPlaceholderField tdf, fieldName, CLng(Definition("Type")), CLng(Definition("Size")), _
        CLng(Definition("OrdinalPosition"))
End Sub

Private Sub RestoreSuspendedCalculatedField( _
    ByVal tableName As String, _
    ByVal conn As Object, _
    ByRef cat As Object, _
    ByVal Definition As Collection, _
    ByVal sourceCalculatedName As String, _
    ByVal targetCalculatedName As String, _
    ByVal expressionText As String, _
    ByRef calculatedNameEventApplied As Boolean)
    Dim db As Object
    Dim tdf As Object
    Dim propertyValues As Collection
    Dim rebuiltCalculatedField As Boolean
    calculatedNameEventApplied = False
    Set propertyValues = Definition("Properties")
    EnsureSuspendedCalculatedFieldPlaceholder tableName, Definition, sourceCalculatedName
    Set db = currentDb
    Set tdf = db.TableDefs(tableName)
    If StrComp(sourceCalculatedName, targetCalculatedName, vbTextCompare) <> 0 Then
        If DAOFieldExists(tdf, targetCalculatedName) Then
            Err.Raise vbObjectError + 7165, "RestoreSuspendedCalculatedField", _
                "Calculated-field target name is not free before restoration: '" & tableName & "." & targetCalculatedName & "'."
        End If
    End If
    If StrComp(sourceCalculatedName, targetCalculatedName, vbBinaryCompare) <> 0 Then
        Set cat = Nothing
        Set cat = CreateObject("ADOX.Catalog")
        Set cat.ActiveConnection = conn
        cat.tables.Refresh
        cat.tables(tableName).Columns.Refresh
        RenameTableField cat, conn, tableName, sourceCalculatedName, targetCalculatedName, rebuiltCalculatedField
        If rebuiltCalculatedField Then Err.Raise vbObjectError + 7166, "RestoreSuspendedCalculatedField", _
            "Temporary calculated-field rename placeholder was unexpectedly treated as a calculated field."
        calculatedNameEventApplied = True
        Set db = currentDb
        Set tdf = db.TableDefs(tableName)
    End If
    tdf.fields.Delete targetCalculatedName
    tdf.fields.Refresh
    AppendCalculatedField tdf, targetCalculatedName, CLng(Definition("Type")), CLng(Definition("Size")), _
        expressionText, CLng(Definition("OrdinalPosition")), propertyValues
    VerifyCalculatedFieldDefinition tableName, targetCalculatedName, CLng(Definition("Type")), CLng(Definition("Size")), _
        CLng(Definition("OrdinalPosition")), expressionText, propertyValues
    Set cat = Nothing
    Set cat = CreateObject("ADOX.Catalog")
    Set cat.ActiveConnection = conn
    cat.tables.Refresh
    cat.tables(tableName).Columns.Refresh
End Sub

Private Function RewriteCalculatedFieldReference( _
    ByVal expressionText As String, _
    ByVal oldName As String, _
    ByVal newName As String) As String
    ' Rewrite exact calculated-field references while preserving quoted/date literals, qualifier names, and function names.
    Dim result As String
    Dim token As String
    Dim originalToken As String
    Dim ch As String
    Dim quoteChar As String
    Dim i As Long
    Dim j As Long
    Dim nextPos As Long
    Do While i < Len(expressionText)
        i = i + 1
        ch = Mid$(expressionText, i, 1)
        If ch = "'" Or ch = """" Then
            quoteChar = ch
            j = i + 1
            Do While j <= Len(expressionText)
                If Mid$(expressionText, j, 1) = quoteChar Then
                    If j < Len(expressionText) And Mid$(expressionText, j + 1, 1) = quoteChar Then
                        j = j + 2
                    Else
                        Exit Do
                    End If
                Else
                    j = j + 1
                End If
            Loop
            If j > Len(expressionText) Then j = Len(expressionText)
            originalToken = Mid$(expressionText, i, j - i + 1)
            If InStr(1, originalToken, oldName, vbTextCompare) > 0 Then
                Err.Raise vbObjectError + 7160, "RewriteCalculatedFieldReference", _
                    "Calculated expression contains '" & oldName & _
                    "' inside a quoted literal. The recovery cannot determine whether that text is an embedded field reference."
            End If
            result = result & originalToken
            i = j
        ElseIf ch = "#" Then
            j = i + 1
            Do While j <= Len(expressionText) And Mid$(expressionText, j, 1) <> "#"
                j = j + 1
            Loop
            If j > Len(expressionText) Then j = Len(expressionText)
            result = result & Mid$(expressionText, i, j - i + 1)
            i = j
        ElseIf ch = "[" Then
            token = vbNullString
            j = i + 1
            Do While j <= Len(expressionText)
                If Mid$(expressionText, j, 1) = "]" Then
                    If j < Len(expressionText) And Mid$(expressionText, j + 1, 1) = "]" Then
                        token = token & "]"
                        j = j + 2
                    Else
                        Exit Do
                    End If
                Else
                    token = token & Mid$(expressionText, j, 1)
                    j = j + 1
                End If
            Loop
            If j <= Len(expressionText) Then
                originalToken = Mid$(expressionText, i, j - i + 1)
                nextPos = SkipSqlSpaces(expressionText, j + 1)
                If StrComp(token, oldName, vbTextCompare) = 0 And _
                   (nextPos > Len(expressionText) Or (Mid$(expressionText, nextPos, 1) <> "." And Mid$(expressionText, nextPos, 1) <> "!")) Then
                    result = result & "[" & Replace(newName, "]", "]]") & "]"
                Else
                    result = result & originalToken
                End If
                i = j
            Else
                result = result & Mid$(expressionText, i)
                Exit Do
            End If
        ElseIf IsSqlWordChar(ch) Then
            j = i
            Do While j <= Len(expressionText) And IsSqlWordChar(Mid$(expressionText, j, 1))
                j = j + 1
            Loop
            token = Mid$(expressionText, i, j - i)
            nextPos = SkipSqlSpaces(expressionText, j)
            If StrComp(token, oldName, vbTextCompare) = 0 Then
                If nextPos <= Len(expressionText) And _
                   (Mid$(expressionText, nextPos, 1) = "." Or Mid$(expressionText, nextPos, 1) = "!") Then
                    result = result & token
                ElseIf nextPos <= Len(expressionText) And Mid$(expressionText, nextPos, 1) = "(" Then
                    result = result & token
                Else
                    Err.Raise vbObjectError + 7157, "RewriteCalculatedFieldReference", _
                        "Calculated expression contains ambiguous bare identifier '" & oldName & _
                        "'. The recovery requires that field reference to be bracketed before it can be rewritten safely."
                End If
            Else
                result = result & token
            End If
            i = j - 1
        Else
            result = result & ch
        End If
    Loop
    RewriteCalculatedFieldReference = result
End Function

Private Function CalculatedFieldStoredPropertyReferenceProblem( _
    ByVal tableName As String, _
    ByVal Definition As Collection, _
    ByVal fieldName As String) As String
    ' These stored properties can contain expressions or SQL. Do not delete/recreate the calculated field when they may require a field-reference rewrite.
    Dim propertyValues As Collection
    Dim entry As Variant
    Dim propertyName As String
    Dim propertyText As String
    Set propertyValues = Definition("Properties")
    For Each entry In propertyValues
        propertyName = LCase$(CStr(entry(0)))
        Select Case propertyName
            Case "defaultvalue", "validationrule", "rowsource"
                If Not IsNull(entry(2)) And Not IsEmpty(entry(2)) Then
                    propertyText = CStr(entry(2))
                    If InStr(1, propertyText, fieldName, vbTextCompare) > 0 Then
                        CalculatedFieldStoredPropertyReferenceProblem = _
                            "Calculated field '" & tableName & "." & CStr(Definition("Name")) & _
                            "' property '" & CStr(entry(0)) & "' contains '" & fieldName & _
                            "'. The recovery cannot rewrite that stored property safely."
                        Exit Function
                    End If
                End If
        End Select
    Next entry
End Function

Private Sub RenameCalculatedTableField( _
    ByVal cat As Object, _
    ByVal conn As Object, _
    ByVal tableName As String, _
    ByVal oldName As String, _
    ByVal newName As String, _
    ByVal expressionText As String)
    ' The placeholder rename is the real Access field rename event; the calculated definition is restored under the final name afterward.
    Dim db As Object
    Dim tdf As Object
    Dim col As Object
    Dim propertyValues As Collection
    Dim fieldType As Long
    Dim fieldSize As Long
    Dim ordinalPosition As Long
    Dim stage As Long
    Dim errNumber As Long
    Dim errDescription As String
    Dim errSource As String
    Dim restoreProblem As String
    On Error GoTo MigrationFail
    CloseObjectIfOpen "Table", tableName
    Set db = currentDb
    Set tdf = db.TableDefs(tableName)
    CaptureCalculatedFieldDefinition tableName, oldName, tdf.fields(oldName), fieldType, fieldSize, ordinalPosition, propertyValues
    ProgressAppend "CALCULATED FIELD MIGRATION: " & tableName & "." & oldName & " -> " & newName
    tdf.fields.Delete oldName
    stage = 1
    tdf.fields.Refresh
    AppendOrdinaryPlaceholderField tdf, oldName, fieldType, fieldSize, ordinalPosition
    stage = 2
    cat.tables.Refresh
    cat.tables(tableName).Columns.Refresh
    Set col = cat.tables(tableName).Columns(oldName)
    On Error Resume Next
    conn.Errors.Clear
    Err.Clear
    On Error GoTo MigrationFail
    col.Name = newName
    stage = 3
    cat.tables(tableName).Columns.Refresh
    Set db = currentDb
    Set tdf = db.TableDefs(tableName)
    tdf.fields.Delete newName
    tdf.fields.Refresh
    stage = 4
    AppendCalculatedField tdf, newName, fieldType, fieldSize, expressionText, ordinalPosition, propertyValues
    stage = 5
    VerifyCalculatedFieldDefinition tableName, newName, fieldType, fieldSize, ordinalPosition, expressionText, propertyValues
    ProgressAppend "CALCULATED FIELD MIGRATION COMPLETE: " & tableName & "." & oldName & " -> " & newName
    Exit Sub
MigrationFail:
    errNumber = Err.Number
    errDescription = Err.Description
    errSource = Err.Source
    On Error Resume Next
    LogADOProviderErrorsFromConnection conn, "CALCULATED FIELD PROVIDER ERROR"
    ProgressAppend "CALCULATED FIELD MIGRATION FAILED: " & tableName & "." & oldName & " -> " & newName & _
        " | Stage=" & CStr(stage) & " | Error " & CStr(errNumber) & ": " & errDescription & _
        IIf(Len(errSource) > 0, " | Source: " & errSource, vbNullString)
    On Error GoTo 0
    restoreProblem = vbNullString
    If RestoreFailedCalculatedFieldMigration(cat, conn, tableName, oldName, newName, fieldType, fieldSize, _
        expressionText, ordinalPosition, propertyValues, stage, restoreProblem) Then
        ProgressAppend "CALCULATED FIELD SELF-ROLLBACK COMPLETE: " & tableName & "." & newName & " -> " & oldName
    Else
        ProgressAppend "CALCULATED FIELD SELF-ROLLBACK FAILED: " & tableName & "." & newName & " -> " & oldName & _
            IIf(Len(restoreProblem) > 0, " | " & restoreProblem, vbNullString)
        errDescription = errDescription & " Calculated-field self-rollback failed: " & restoreProblem
    End If
    If Len(errSource) = 0 Then errSource = "RenameCalculatedTableField"
    Err.Raise errNumber, errSource, errDescription
End Sub

Private Sub CaptureCalculatedFieldDefinition( _
    ByVal tableName As String, _
    ByVal fieldName As String, _
    ByVal fld As Object, _
    ByRef fieldType As Long, _
    ByRef fieldSize As Long, _
    ByRef ordinalPosition As Long, _
    ByRef propertyValues As Collection)
    ' Preserve the physical definition and editable presentation properties that would otherwise be lost on delete/recreate.
    Dim names As Variant
    Dim propertyName As Variant
    Dim prop As Object
    Dim propertyType As Long
    Dim propertyValue As Variant
    Dim readErr As Long
    Dim readSource As String
    Dim readDesc As String
    fieldType = CLng(fld.Type)
    fieldSize = CLng(fld.Size)
    ordinalPosition = CLng(fld.ordinalPosition)
    Set propertyValues = New Collection
    names = Array("Attributes", "Required", "AllowZeroLength", "DefaultValue", "ValidationRule", "ValidationText", "AppendOnly", _
        "Description", "Caption", "Format", "DecimalPlaces", "DisplayControl", "ColumnWidth", "ColumnOrder", "ColumnHidden", _
        "TextAlign", "InputMask", "TextFormat", "ShowDatePicker", "UnicodeCompression", "IMEMode", "IMESentenceMode", _
        "RowSourceType", "RowSource", "BoundColumn", "ColumnCount", "ColumnHeads", "ColumnWidths", "ListRows", "ListWidth", _
        "LimitToList", "AllowValueListEdits", "ListItemsEditForm", "ShowOnlyRowSourceValues")
    For Each propertyName In names
        Set prop = Nothing
        readErr = 0
        readSource = vbNullString
        readDesc = vbNullString
        On Error Resume Next
        Err.Clear
        Set prop = fld.properties(CStr(propertyName))
        readErr = Err.Number
        readSource = Err.Source
        readDesc = Err.Description
        Err.Clear
        On Error GoTo 0
        If readErr <> 0 Then
            If readErr <> 3270 And readErr <> 3265 Then
                logLine "CALCULATED FIELD PROPERTY CAPTURE FAILED: Table=" & tableName & " | Field=" & fieldName & _
                    " | Property=" & CStr(propertyName) & " | Read=Property object | Err.Number=" & CStr(readErr) & _
                    " | Err.Source=" & readSource & " | Err.Description=" & readDesc
                Err.Raise readErr, "CaptureCalculatedFieldDefinition", _
                    "Could not inspect existing calculated field property '" & tableName & "." & fieldName & "." & _
                    CStr(propertyName) & "'. Original source: '" & readSource & "'. " & readDesc
            End If
        ElseIf Not prop Is Nothing Then
            On Error Resume Next
            Err.Clear
            propertyType = CLng(prop.Type)
            readErr = Err.Number
            readSource = Err.Source
            readDesc = Err.Description
            Err.Clear
            On Error GoTo 0
            If readErr <> 0 Then
                logLine "CALCULATED FIELD PROPERTY CAPTURE FAILED: Table=" & tableName & " | Field=" & fieldName & _
                    " | Property=" & CStr(propertyName) & " | Read=Type | Err.Number=" & CStr(readErr) & _
                    " | Err.Source=" & readSource & " | Err.Description=" & readDesc
                Err.Raise readErr, "CaptureCalculatedFieldDefinition", _
                    "Could not read Type for existing calculated field property '" & tableName & "." & fieldName & "." & _
                    CStr(propertyName) & "'. Original source: '" & readSource & "'. " & readDesc
            End If
            On Error Resume Next
            Err.Clear
            propertyValue = prop.value
            readErr = Err.Number
            readSource = Err.Source
            readDesc = Err.Description
            Err.Clear
            On Error GoTo 0
            If readErr <> 0 Then
                logLine "CALCULATED FIELD PROPERTY CAPTURE FAILED: Table=" & tableName & " | Field=" & fieldName & _
                    " | Property=" & CStr(propertyName) & " | Read=Value | Err.Number=" & CStr(readErr) & _
                    " | Err.Source=" & readSource & " | Err.Description=" & readDesc
                Err.Raise readErr, "CaptureCalculatedFieldDefinition", _
                    "Could not read Value for existing calculated field property '" & tableName & "." & fieldName & "." & _
                    CStr(propertyName) & "'. Original source: '" & readSource & "'. " & readDesc
            End If
            propertyValues.Add Array(CStr(propertyName), propertyType, propertyValue)
        End If
    Next propertyName
End Sub

Private Sub AppendOrdinaryPlaceholderField( _
    ByVal tdf As Object, _
    ByVal fieldName As String, _
    ByVal fieldType As Long, _
    ByVal fieldSize As Long, _
    ByVal ordinalPosition As Long)
    Dim fld As Object
    If fieldSize > 0 Then
        Set fld = tdf.CreateField(fieldName, fieldType, fieldSize)
    Else
        Set fld = tdf.CreateField(fieldName, fieldType)
    End If
    tdf.fields.Append fld
    tdf.fields.Refresh
    tdf.fields(fieldName).ordinalPosition = ordinalPosition
    tdf.fields.Refresh
End Sub

Private Sub AppendCalculatedField( _
    ByVal tdf As Object, _
    ByVal fieldName As String, _
    ByVal fieldType As Long, _
    ByVal fieldSize As Long, _
    ByVal expressionText As String, _
    ByVal ordinalPosition As Long, _
    ByVal propertyValues As Collection)
    Dim fld As Object
    Dim prop As Object
    Dim entry As Variant
    Dim currentValue As Variant
    Dim i As Long
    Dim restoreErr As Long
    Dim restoreDesc As String
    If fieldSize > 0 Then
        Set fld = tdf.CreateField(fieldName, fieldType, fieldSize)
    Else
        Set fld = tdf.CreateField(fieldName, fieldType)
    End If
    fld.expression = expressionText
    tdf.fields.Append fld
    tdf.fields.Refresh
    Set fld = tdf.fields(fieldName)
    fld.ordinalPosition = ordinalPosition
    For i = 1 To propertyValues.count
        entry = propertyValues(i)
        Set prop = Nothing
        restoreErr = 0
        restoreDesc = vbNullString
        On Error Resume Next
        Err.Clear
        Set prop = fld.properties(CStr(entry(0)))
        If Err.Number = 0 And Not prop Is Nothing Then
            currentValue = prop.value
            If Err.Number <> 0 Then
                Err.Clear
                prop.value = entry(2)
                restoreErr = Err.Number
                restoreDesc = Err.Description
            ElseIf Not CalculatedFieldPropertyValuesEqual(currentValue, entry(2)) Then
                Err.Clear
                prop.value = entry(2)
                restoreErr = Err.Number
                restoreDesc = Err.Description
            End If
        ElseIf IsCalculatedFieldBuiltInProperty(CStr(entry(0))) Then
            restoreErr = Err.Number
            restoreDesc = Err.Description
            If restoreErr = 0 Then
                restoreErr = vbObjectError + 7148
                restoreDesc = "Built-in calculated field property is unavailable."
            End If
        Else
            Err.Clear
            Set prop = fld.CreateProperty(CStr(entry(0)), CLng(entry(1)), entry(2))
            fld.properties.Append prop
            restoreErr = Err.Number
            restoreDesc = Err.Description
        End If
        Err.Clear
        On Error GoTo 0
        If restoreErr <> 0 Then
            Err.Raise restoreErr, "AppendCalculatedField", _
                "Could not restore calculated field property '" & CStr(entry(0)) & "': " & restoreDesc
        End If
    Next i
    tdf.fields.Refresh
End Sub

Private Function IsCalculatedFieldBuiltInProperty(ByVal propertyName As String) As Boolean
    Select Case LCase$(propertyName)
        Case "attributes", "required", "allowzerolength", "defaultvalue", "validationrule", "validationtext", "appendonly"
            IsCalculatedFieldBuiltInProperty = True
    End Select
End Function

Private Function CalculatedFieldPropertyValuesEqual(ByVal actualValue As Variant, ByVal expectedValue As Variant) As Boolean
    If IsNull(expectedValue) Then
        CalculatedFieldPropertyValuesEqual = IsNull(actualValue)
    ElseIf IsNull(actualValue) Then
        CalculatedFieldPropertyValuesEqual = False
    Else
        CalculatedFieldPropertyValuesEqual = (StrComp(CStr(actualValue), CStr(expectedValue), vbBinaryCompare) = 0)
    End If
End Function

Private Sub VerifyCalculatedFieldDefinition( _
    ByVal tableName As String, _
    ByVal fieldName As String, _
    ByVal expectedType As Long, _
    ByVal expectedSize As Long, _
    ByVal expectedOrdinalPosition As Long, _
    ByVal expectedExpression As String, _
    ByVal propertyValues As Collection)
    Dim db As Object
    Dim fld As Object
    Dim prop As Object
    Dim entry As Variant
    Dim expectedValue As Variant
    Dim actualValue As Variant
    Dim i As Long
    Dim propertyName As String
    Dim actualExpression As String
    Set db = currentDb
    Set fld = db.TableDefs(tableName).fields(fieldName)
    If CLng(fld.Type) <> expectedType Then
        Err.Raise vbObjectError + 7140, "VerifyCalculatedFieldDefinition", _
            "Calculated field type changed while renaming '" & tableName & "." & fieldName & "'."
    End If
    If expectedSize > 0 Then
        If CLng(fld.Size) <> expectedSize Then
            Err.Raise vbObjectError + 7141, "VerifyCalculatedFieldDefinition", _
                "Calculated field size changed while renaming '" & tableName & "." & fieldName & "'."
        End If
    End If
    If CLng(fld.ordinalPosition) <> expectedOrdinalPosition Then
        Err.Raise vbObjectError + 7144, "VerifyCalculatedFieldDefinition", _
            "Calculated field ordinal position changed while renaming '" & tableName & "." & fieldName & "'."
    End If
    If Not TryGetCalculatedFieldExpression(tableName, fieldName, actualExpression) Then
        Err.Raise vbObjectError + 7142, "VerifyCalculatedFieldDefinition", _
            "Calculated field expression is missing after renaming '" & tableName & "." & fieldName & "'."
    End If
    If StrComp(actualExpression, expectedExpression, vbBinaryCompare) <> 0 Then
        Err.Raise vbObjectError + 7143, "VerifyCalculatedFieldDefinition", _
            "Calculated field expression changed while renaming '" & tableName & "." & fieldName & "'."
    End If
    For i = 1 To propertyValues.count
        entry = propertyValues(i)
        propertyName = CStr(entry(0))
        Set prop = Nothing
        On Error Resume Next
        Err.Clear
        Set prop = fld.properties(propertyName)
        If Err.Number <> 0 Or prop Is Nothing Then
            Err.Clear
            On Error GoTo 0
            Err.Raise vbObjectError + 7145, "VerifyCalculatedFieldDefinition", _
                "Calculated field property is missing after renaming '" & tableName & "." & fieldName & "': '" & propertyName & "'."
        End If
        expectedValue = entry(2)
        actualValue = prop.value
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            Err.Raise vbObjectError + 7146, "VerifyCalculatedFieldDefinition", _
                "Calculated field property could not be read after renaming '" & tableName & "." & fieldName & "': '" & propertyName & "'."
        End If
        Err.Clear
        On Error GoTo 0
        If Not CalculatedFieldPropertyValuesEqual(actualValue, expectedValue) Then
            Err.Raise vbObjectError + 7147, "VerifyCalculatedFieldDefinition", _
                "Calculated field property changed while renaming '" & tableName & "." & fieldName & "': '" & propertyName & "'."
        End If
    Next i
End Sub

Private Function RestoreFailedCalculatedFieldMigration( _
    ByVal cat As Object, _
    ByVal conn As Object, _
    ByVal tableName As String, _
    ByVal oldName As String, _
    ByVal newName As String, _
    ByVal fieldType As Long, _
    ByVal fieldSize As Long, _
    ByVal expressionText As String, _
    ByVal ordinalPosition As Long, _
    ByVal propertyValues As Collection, _
    ByVal stage As Long, _
    ByRef problem As String) As Boolean
    ' Restore both schema and Name AutoCorrect direction when a calculated migration fails before its journal step can be recorded.
    Dim db As Object
    Dim tdf As Object
    Dim col As Object
    On Error GoTo RestoreFail
    If stage = 0 Then
        RestoreFailedCalculatedFieldMigration = True
        Exit Function
    End If
    CloseObjectIfOpen "Table", tableName
    Set db = currentDb
    Set tdf = db.TableDefs(tableName)
    If stage >= 3 Then
        If DAOFieldExists(tdf, oldName) Then tdf.fields.Delete oldName
        If DAOFieldExists(tdf, newName) Then tdf.fields.Delete newName
        tdf.fields.Refresh
        AppendOrdinaryPlaceholderField tdf, newName, fieldType, fieldSize, ordinalPosition
        cat.tables.Refresh
        cat.tables(tableName).Columns.Refresh
        Set col = cat.tables(tableName).Columns(newName)
        On Error Resume Next
        conn.Errors.Clear
        Err.Clear
        On Error GoTo RestoreFail
        col.Name = oldName
        cat.tables(tableName).Columns.Refresh
        Set db = currentDb
        Set tdf = db.TableDefs(tableName)
        tdf.fields.Delete oldName
        tdf.fields.Refresh
        AppendCalculatedField tdf, oldName, fieldType, fieldSize, expressionText, ordinalPosition, propertyValues
    Else
        If DAOFieldExists(tdf, oldName) Then tdf.fields.Delete oldName
        tdf.fields.Refresh
        AppendCalculatedField tdf, oldName, fieldType, fieldSize, expressionText, ordinalPosition, propertyValues
    End If
    VerifyCalculatedFieldDefinition tableName, oldName, fieldType, fieldSize, ordinalPosition, expressionText, propertyValues
    RestoreFailedCalculatedFieldMigration = True
    Exit Function
RestoreFail:
    problem = "Error " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function DAOFieldExists(ByVal tdf As Object, ByVal fieldName As String) As Boolean
    Dim fld As Object
    On Error Resume Next
    Set fld = tdf.fields(fieldName)
    DAOFieldExists = (Err.Number = 0 And Not fld Is Nothing)
    Err.Clear
    On Error GoTo 0
End Function

Private Function TryGetCalculatedFieldExpression( _
    ByVal tableName As String, _
    ByVal fieldName As String, _
    ByRef expressionText As String) As Boolean
    Dim db As Object
    Dim fld As Object
    Dim value As Variant
    expressionText = vbNullString
    On Error GoTo NotCalculated
    Set db = currentDb
    Set fld = db.TableDefs(tableName).fields(fieldName)
    On Error Resume Next
    Err.Clear
    value = fld.properties("Expression").value
    If Err.Number = 0 Then
        If Not IsNull(value) Then expressionText = CStr(value)
    End If
    Err.Clear
    On Error GoTo 0
    TryGetCalculatedFieldExpression = (Len(expressionText) > 0)
    Exit Function
NotCalculated:
    expressionText = vbNullString
    Err.Clear
End Function

Private Function TableFieldDiagnosticText(ByVal tableName As String, ByVal fieldName As String) As String
    Dim db As Object
    Dim fld As Object
    Dim value As Variant
    Dim text As String
    On Error GoTo MetadataFail
    Set db = currentDb
    Set fld = db.TableDefs(tableName).fields(fieldName)
    text = "DAO Type=" & CStr(fld.Type) & " | Attributes=" & CStr(fld.Attributes) & " | Size=" & CStr(fld.Size)
    On Error Resume Next
    value = fld.Required
    If Err.Number = 0 Then text = text & " | Required=" & CStr(CBool(value))
    Err.Clear
    value = fld.allowZeroLength
    If Err.Number = 0 Then text = text & " | AllowZeroLength=" & CStr(CBool(value))
    Err.Clear
    value = fld.properties("Expression").value
    If Err.Number = 0 Then
        If Not IsNull(value) Then
            If Len(CStr(value)) > 0 Then text = text & " | Expression=" & CStr(value)
        End If
    End If
    Err.Clear
    value = fld.DefaultValue
    If Err.Number = 0 Then
        If Not IsNull(value) Then
            If Len(CStr(value)) > 0 Then text = text & " | DefaultValue=" & CStr(value)
        End If
    End If
    Err.Clear
    value = fld.validationRule
    If Err.Number = 0 Then
        If Not IsNull(value) Then
            If Len(CStr(value)) > 0 Then text = text & " | ValidationRule=" & CStr(value)
        End If
    End If
    On Error GoTo 0
    TableFieldDiagnosticText = text
    Exit Function
MetadataFail:
    TableFieldDiagnosticText = "Field metadata unavailable: Error " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Sub LogOrdinaryTableFieldDiagnostics( _
    ByVal tableName As String, _
    ByVal fieldName As String, _
    ByVal newName As String, _
    ByVal col As Object)
    Dim db As Object
    Dim tdf As Object
    Dim fld As Object
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    logLine "TABLE FIELD DIAGNOSTICS BEGIN: " & tableName & "." & fieldName & " -> " & newName
    On Error Resume Next
    Err.Clear
    Set db = currentDb
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If errNumber <> 0 Or db Is Nothing Then
        logLine "  DAO Database unavailable: " & DiagnosticErrorText(errNumber, errSource, errDescription)
    Else
        On Error Resume Next
        Err.Clear
        Set tdf = db.TableDefs(tableName)
        errNumber = Err.Number
        errSource = Err.Source
        errDescription = Err.Description
        Err.Clear
        On Error GoTo 0
        If errNumber <> 0 Or tdf Is Nothing Then
            logLine "  DAO TableDef unavailable: " & DiagnosticErrorText(errNumber, errSource, errDescription)
        Else
            On Error Resume Next
            Err.Clear
            Set fld = tdf.fields(fieldName)
            errNumber = Err.Number
            errSource = Err.Source
            errDescription = Err.Description
            Err.Clear
            On Error GoTo 0
            If errNumber <> 0 Or fld Is Nothing Then
                logLine "  DAO Field unavailable: " & DiagnosticErrorText(errNumber, errSource, errDescription)
            Else
                LogDAOFieldDiagnostics fld
            End If
            LogTableDefDiagnostics tdf, fieldName
        End If
        LogRelationDiagnostics db, tableName, fieldName
    End If
    If col Is Nothing Then
        logLine "  ADOX Column unavailable."
    Else
        LogADOXColumnDiagnostics col
    End If
    logLine "TABLE FIELD DIAGNOSTICS END: " & tableName & "." & fieldName & " -> " & newName
End Sub

Private Sub LogDAOFieldDiagnostics(ByVal fld As Object)
    Dim value As Variant
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim properties As Object
    logLine "  DAO Field core BEGIN"
    On Error Resume Next
    Err.Clear
    value = fld.Type
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If errNumber = 0 Then
        logLine "    Type=" & DiagnosticValueText(value)
    Else
        logLine "    Type=<unreadable: " & DiagnosticErrorText(errNumber, errSource, errDescription) & ">"
    End If
    On Error Resume Next
    Err.Clear
    value = fld.Size
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If errNumber = 0 Then
        logLine "    Size=" & DiagnosticValueText(value)
    Else
        logLine "    Size=<unreadable: " & DiagnosticErrorText(errNumber, errSource, errDescription) & ">"
    End If
    On Error Resume Next
    Err.Clear
    value = fld.Attributes
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If errNumber = 0 Then
        logLine "    Attributes=" & DiagnosticValueText(value)
    Else
        logLine "    Attributes=<unreadable: " & DiagnosticErrorText(errNumber, errSource, errDescription) & ">"
    End If
    On Error Resume Next
    Err.Clear
    value = fld.ordinalPosition
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If errNumber = 0 Then
        logLine "    OrdinalPosition=" & DiagnosticValueText(value)
    Else
        logLine "    OrdinalPosition=<unreadable: " & DiagnosticErrorText(errNumber, errSource, errDescription) & ">"
    End If
    logLine "  DAO Field core END"
    Set properties = Nothing
    On Error Resume Next
    Err.Clear
    Set properties = fld.properties
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If errNumber <> 0 Or properties Is Nothing Then
        logLine "  DAO Field.Properties unavailable: " & DiagnosticErrorText(errNumber, errSource, errDescription)
        Exit Sub
    End If
    LogDiagnosticPropertyCollection properties, "  DAO Field.Properties"
    LogNamedDiagnosticProperties properties, "  DAO Field selected metadata", _
        Array("DisplayControl", "RowSourceType", "RowSource", "BoundColumn", "ColumnCount", "ColumnWidths", _
            "LimitToList", "AllowValueListEdits", "ListItemsEditForm", "ShowOnlyRowSourceValues", "Description", _
            "Caption", "Format", "InputMask", "DefaultValue", "ValidationRule", "ValidationText", _
            "UnicodeCompression", "Required", "AllowZeroLength", "AppendOnly")
End Sub

Private Sub LogTableDefDiagnostics(ByVal tdf As Object, ByVal fieldName As String)
    Dim value As Variant
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim idx As Object
    Dim idxField As Object
    Dim indexCount As Long
    Dim fieldCount As Long
    Dim i As Long
    Dim j As Long
    Dim indexName As String
    Dim indexFieldName As String
    Dim fieldList As String
    Dim containsField As Boolean
    Dim foundIndex As Boolean
    logLine "  DAO TableDef metadata BEGIN"
    On Error Resume Next
    Err.Clear
    value = tdf.Attributes
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If errNumber = 0 Then
        logLine "    Attributes=" & DiagnosticValueText(value)
    Else
        logLine "    Attributes=<unreadable: " & DiagnosticErrorText(errNumber, errSource, errDescription) & ">"
    End If
    On Error Resume Next
    Err.Clear
    value = tdf.Connect
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If errNumber = 0 Then
        logLine "    Connect=" & DiagnosticValueText(value)
    Else
        logLine "    Connect=<unreadable: " & DiagnosticErrorText(errNumber, errSource, errDescription) & ">"
    End If
    On Error Resume Next
    Err.Clear
    value = tdf.sourceTableName
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If errNumber = 0 Then
        logLine "    SourceTableName=" & DiagnosticValueText(value)
    Else
        logLine "    SourceTableName=<unreadable: " & DiagnosticErrorText(errNumber, errSource, errDescription) & ">"
    End If
    logLine "  DAO TableDef metadata END"
    On Error Resume Next
    Err.Clear
    indexCount = tdf.Indexes.count
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If errNumber <> 0 Then
        logLine "  DAO indexes unavailable: " & DiagnosticErrorText(errNumber, errSource, errDescription)
        Exit Sub
    End If
    logLine "  DAO indexes containing field BEGIN"
    For i = 0 To indexCount - 1
        Set idx = Nothing
        On Error Resume Next
        Err.Clear
        Set idx = tdf.Indexes(i)
        errNumber = Err.Number
        errSource = Err.Source
        errDescription = Err.Description
        Err.Clear
        On Error GoTo 0
        If errNumber <> 0 Or idx Is Nothing Then
            logLine "    Index #" & CStr(i) & " unavailable: " & DiagnosticErrorText(errNumber, errSource, errDescription)
        Else
            indexName = "<unreadable>"
            On Error Resume Next
            Err.Clear
            indexName = CStr(idx.Name)
            errNumber = Err.Number
            errSource = Err.Source
            errDescription = Err.Description
            Err.Clear
            On Error GoTo 0
            If errNumber <> 0 Then
                indexName = "<unreadable: " & DiagnosticErrorText(errNumber, errSource, errDescription) & ">"
            End If
            fieldList = vbNullString
            containsField = False
            On Error Resume Next
            Err.Clear
            fieldCount = idx.fields.count
            errNumber = Err.Number
            errSource = Err.Source
            errDescription = Err.Description
            Err.Clear
            On Error GoTo 0
            If errNumber <> 0 Then
                fieldList = "<unreadable field collection: " & DiagnosticErrorText(errNumber, errSource, errDescription) & ">"
            Else
                For j = 0 To fieldCount - 1
                    Set idxField = Nothing
                    On Error Resume Next
                    Err.Clear
                    Set idxField = idx.fields(j)
                    errNumber = Err.Number
                    errSource = Err.Source
                    errDescription = Err.Description
                    Err.Clear
                    On Error GoTo 0
                    If errNumber <> 0 Or idxField Is Nothing Then
                        indexFieldName = "<unreadable: " & DiagnosticErrorText(errNumber, errSource, errDescription) & ">"
                    Else
                        On Error Resume Next
                        Err.Clear
                        indexFieldName = CStr(idxField.Name)
                        errNumber = Err.Number
                        errSource = Err.Source
                        errDescription = Err.Description
                        Err.Clear
                        On Error GoTo 0
                        If errNumber <> 0 Then
                            indexFieldName = "<unreadable: " & DiagnosticErrorText(errNumber, errSource, errDescription) & ">"
                        ElseIf StrComp(indexFieldName, fieldName, vbTextCompare) = 0 Then
                            containsField = True
                        End If
                    End If
                    If Len(fieldList) > 0 Then fieldList = fieldList & ", "
                    fieldList = fieldList & indexFieldName
                Next j
            End If
            If containsField Then
                foundIndex = True
                logLine "    Index=" & indexName
                On Error Resume Next
                Err.Clear
                value = idx.Primary
                errNumber = Err.Number
                errSource = Err.Source
                errDescription = Err.Description
                Err.Clear
                On Error GoTo 0
                If errNumber = 0 Then
                    logLine "      Primary=" & DiagnosticValueText(value)
                Else
                    logLine "      Primary=<unreadable: " & DiagnosticErrorText(errNumber, errSource, errDescription) & ">"
                End If
                On Error Resume Next
                Err.Clear
                value = idx.Unique
                errNumber = Err.Number
                errSource = Err.Source
                errDescription = Err.Description
                Err.Clear
                On Error GoTo 0
                If errNumber = 0 Then
                    logLine "      Unique=" & DiagnosticValueText(value)
                Else
                    logLine "      Unique=<unreadable: " & DiagnosticErrorText(errNumber, errSource, errDescription) & ">"
                End If
                On Error Resume Next
                Err.Clear
                value = idx.Required
                errNumber = Err.Number
                errSource = Err.Source
                errDescription = Err.Description
                Err.Clear
                On Error GoTo 0
                If errNumber = 0 Then
                    logLine "      Required=" & DiagnosticValueText(value)
                Else
                    logLine "      Required=<unreadable: " & DiagnosticErrorText(errNumber, errSource, errDescription) & ">"
                End If
                On Error Resume Next
                Err.Clear
                value = idx.IgnoreNulls
                errNumber = Err.Number
                errSource = Err.Source
                errDescription = Err.Description
                Err.Clear
                On Error GoTo 0
                If errNumber = 0 Then
                    logLine "      IgnoreNulls=" & DiagnosticValueText(value)
                Else
                    logLine "      IgnoreNulls=<unreadable: " & DiagnosticErrorText(errNumber, errSource, errDescription) & ">"
                End If
                logLine "      Fields=" & fieldList
            End If
        End If
    Next i
    If Not foundIndex Then logLine "    <none>"
    logLine "  DAO indexes containing field END"
End Sub

Private Sub LogRelationDiagnostics(ByVal db As Object, ByVal tableName As String, ByVal fieldName As String)
    Dim rel As Object
    Dim relField As Object
    Dim value As Variant
    Dim relationCount As Long
    Dim fieldCount As Long
    Dim i As Long
    Dim j As Long
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim relationName As String
    Dim localTable As String
    Dim foreignTable As String
    Dim localField As String
    Dim foreignField As String
    Dim mappings As String
    Dim containsField As Boolean
    Dim foundRelation As Boolean
    On Error Resume Next
    Err.Clear
    relationCount = db.Relations.count
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If errNumber <> 0 Then
        logLine "  DAO relations unavailable: " & DiagnosticErrorText(errNumber, errSource, errDescription)
        Exit Sub
    End If
    logLine "  DAO relations involving field BEGIN"
    For i = 0 To relationCount - 1
        Set rel = Nothing
        On Error Resume Next
        Err.Clear
        Set rel = db.Relations(i)
        errNumber = Err.Number
        errSource = Err.Source
        errDescription = Err.Description
        Err.Clear
        On Error GoTo 0
        If errNumber <> 0 Or rel Is Nothing Then
            logLine "    Relation #" & CStr(i) & " unavailable: " & DiagnosticErrorText(errNumber, errSource, errDescription)
        Else
            relationName = "<unreadable>"
            localTable = vbNullString
            foreignTable = vbNullString
            On Error Resume Next
            Err.Clear
            relationName = CStr(rel.Name)
            errNumber = Err.Number
            errSource = Err.Source
            errDescription = Err.Description
            Err.Clear
            On Error GoTo 0
            If errNumber <> 0 Then
                relationName = "<unreadable: " & DiagnosticErrorText(errNumber, errSource, errDescription) & ">"
            End If
            On Error Resume Next
            Err.Clear
            localTable = CStr(rel.Table)
            errNumber = Err.Number
            errSource = Err.Source
            errDescription = Err.Description
            Err.Clear
            On Error GoTo 0
            If errNumber <> 0 Then
                logLine "    Relation=" & relationName & " Table unreadable: " & _
                    DiagnosticErrorText(errNumber, errSource, errDescription)
                localTable = vbNullString
            End If
            On Error Resume Next
            Err.Clear
            foreignTable = CStr(rel.foreignTable)
            errNumber = Err.Number
            errSource = Err.Source
            errDescription = Err.Description
            Err.Clear
            On Error GoTo 0
            If errNumber <> 0 Then
                logLine "    Relation=" & relationName & " ForeignTable unreadable: " & _
                    DiagnosticErrorText(errNumber, errSource, errDescription)
                foreignTable = vbNullString
            End If
            mappings = vbNullString
            containsField = False
            On Error Resume Next
            Err.Clear
            fieldCount = rel.fields.count
            errNumber = Err.Number
            errSource = Err.Source
            errDescription = Err.Description
            Err.Clear
            On Error GoTo 0
            If errNumber <> 0 Then
                mappings = "<unreadable field collection: " & DiagnosticErrorText(errNumber, errSource, errDescription) & ">"
            Else
                For j = 0 To fieldCount - 1
                    Set relField = Nothing
                    On Error Resume Next
                    Err.Clear
                    Set relField = rel.fields(j)
                    errNumber = Err.Number
                    errSource = Err.Source
                    errDescription = Err.Description
                    Err.Clear
                    On Error GoTo 0
                    If errNumber <> 0 Or relField Is Nothing Then
                        localField = "<unreadable: " & DiagnosticErrorText(errNumber, errSource, errDescription) & ">"
                        foreignField = "<unreadable>"
                    Else
                        On Error Resume Next
                        Err.Clear
                        localField = CStr(relField.Name)
                        errNumber = Err.Number
                        errSource = Err.Source
                        errDescription = Err.Description
                        Err.Clear
                        On Error GoTo 0
                        If errNumber <> 0 Then
                            localField = "<unreadable: " & DiagnosticErrorText(errNumber, errSource, errDescription) & ">"
                        End If
                        On Error Resume Next
                        Err.Clear
                        foreignField = CStr(relField.ForeignName)
                        errNumber = Err.Number
                        errSource = Err.Source
                        errDescription = Err.Description
                        Err.Clear
                        On Error GoTo 0
                        If errNumber <> 0 Then
                            foreignField = "<unreadable: " & DiagnosticErrorText(errNumber, errSource, errDescription) & ">"
                        End If
                    End If
                    If Len(mappings) > 0 Then mappings = mappings & "; "
                    mappings = mappings & localField & " -> " & foreignField
                    If StrComp(localTable, tableName, vbTextCompare) = 0 And _
                       StrComp(localField, fieldName, vbTextCompare) = 0 Then containsField = True
                    If StrComp(foreignTable, tableName, vbTextCompare) = 0 And _
                       StrComp(foreignField, fieldName, vbTextCompare) = 0 Then containsField = True
                Next j
            End If
            If containsField Then
                foundRelation = True
                logLine "    Relation=" & relationName
                logLine "      Table=" & DiagnosticValueText(localTable)
                logLine "      ForeignTable=" & DiagnosticValueText(foreignTable)
                On Error Resume Next
                Err.Clear
                value = rel.Attributes
                errNumber = Err.Number
                errSource = Err.Source
                errDescription = Err.Description
                Err.Clear
                On Error GoTo 0
                If errNumber = 0 Then
                    logLine "      Attributes=" & DiagnosticValueText(value)
                Else
                    logLine "      Attributes=<unreadable: " & DiagnosticErrorText(errNumber, errSource, errDescription) & ">"
                End If
                logLine "      Fields=" & mappings
            End If
        End If
    Next i
    If Not foundRelation Then logLine "    <none>"
    logLine "  DAO relations involving field END"
End Sub

Private Sub LogADOXColumnDiagnostics(ByVal col As Object)
    Dim value As Variant
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim properties As Object
    logLine "  ADOX Column core BEGIN"
    On Error Resume Next
    Err.Clear
    value = col.Name
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If errNumber = 0 Then
        logLine "    Name=" & DiagnosticValueText(value)
    Else
        logLine "    Name=<unreadable: " & DiagnosticErrorText(errNumber, errSource, errDescription) & ">"
    End If
    On Error Resume Next
    Err.Clear
    value = col.Type
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If errNumber = 0 Then
        logLine "    Type=" & DiagnosticValueText(value)
    Else
        logLine "    Type=<unreadable: " & DiagnosticErrorText(errNumber, errSource, errDescription) & ">"
    End If
    On Error Resume Next
    Err.Clear
    value = col.DefinedSize
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If errNumber = 0 Then
        logLine "    DefinedSize=" & DiagnosticValueText(value)
    Else
        logLine "    DefinedSize=<unreadable: " & DiagnosticErrorText(errNumber, errSource, errDescription) & ">"
    End If
    On Error Resume Next
    Err.Clear
    value = col.Attributes
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If errNumber = 0 Then
        logLine "    Attributes=" & DiagnosticValueText(value)
    Else
        logLine "    Attributes=<unreadable: " & DiagnosticErrorText(errNumber, errSource, errDescription) & ">"
    End If
    logLine "  ADOX Column core END"
    Set properties = Nothing
    On Error Resume Next
    Err.Clear
    Set properties = col.properties
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If errNumber <> 0 Or properties Is Nothing Then
        logLine "  ADOX Column.Properties unavailable: " & DiagnosticErrorText(errNumber, errSource, errDescription)
        Exit Sub
    End If
    LogDiagnosticPropertyCollection properties, "  ADOX Column.Properties"
End Sub

Private Sub LogDiagnosticPropertyCollection(ByVal properties As Object, ByVal labelText As String)
    Dim prop As Object
    Dim propertyName As String
    Dim propertyType As Variant
    Dim propertyValue As Variant
    Dim propertyCount As Long
    Dim itemNumber As Long
    Dim itemErrNumber As Long
    Dim itemErrSource As String
    Dim itemErrDescription As String
    Dim nameErrNumber As Long
    Dim nameErrSource As String
    Dim nameErrDescription As String
    Dim typeErrNumber As Long
    Dim typeErrSource As String
    Dim typeErrDescription As String
    Dim valueErrNumber As Long
    Dim valueErrSource As String
    Dim valueErrDescription As String
    Dim typeText As String
    Dim valueText As String
    On Error Resume Next
    Err.Clear
    propertyCount = properties.count
    itemErrNumber = Err.Number
    itemErrSource = Err.Source
    itemErrDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If itemErrNumber <> 0 Then
        logLine labelText & " count unavailable: " & DiagnosticErrorText(itemErrNumber, itemErrSource, itemErrDescription)
        Exit Sub
    End If
    logLine labelText & " BEGIN"
    For itemNumber = 0 To propertyCount - 1
        Set prop = Nothing
        On Error Resume Next
        Err.Clear
        Set prop = properties(itemNumber)
        itemErrNumber = Err.Number
        itemErrSource = Err.Source
        itemErrDescription = Err.Description
        Err.Clear
        On Error GoTo 0
        If itemErrNumber <> 0 Or prop Is Nothing Then
            logLine labelText & " #" & CStr(itemNumber) & " unavailable: " & _
                DiagnosticErrorText(itemErrNumber, itemErrSource, itemErrDescription)
        Else
            propertyName = vbNullString
            propertyType = Empty
            propertyValue = Empty
            On Error Resume Next
            Err.Clear
            propertyName = CStr(prop.Name)
            nameErrNumber = Err.Number
            nameErrSource = Err.Source
            nameErrDescription = Err.Description
            Err.Clear
            On Error GoTo 0
            On Error Resume Next
            Err.Clear
            propertyType = prop.Type
            typeErrNumber = Err.Number
            typeErrSource = Err.Source
            typeErrDescription = Err.Description
            Err.Clear
            On Error GoTo 0
            On Error Resume Next
            Err.Clear
            propertyValue = prop.value
            valueErrNumber = Err.Number
            valueErrSource = Err.Source
            valueErrDescription = Err.Description
            Err.Clear
            On Error GoTo 0
            If nameErrNumber <> 0 Then
                propertyName = "<unreadable: " & DiagnosticErrorText(nameErrNumber, nameErrSource, nameErrDescription) & ">"
            End If
            If typeErrNumber = 0 Then
                typeText = DiagnosticValueText(propertyType)
            Else
                typeText = "<unreadable: " & DiagnosticErrorText(typeErrNumber, typeErrSource, typeErrDescription) & ">"
            End If
            If valueErrNumber = 0 Then
                valueText = DiagnosticValueText(propertyValue)
            Else
                valueText = "<unreadable: " & DiagnosticErrorText(valueErrNumber, valueErrSource, valueErrDescription) & ">"
            End If
            logLine labelText & " #" & CStr(itemNumber) & ": Name=" & propertyName & _
                " | Type=" & typeText & " | Value=" & valueText
        End If
    Next itemNumber
    logLine labelText & " END"
End Sub

Private Sub LogNamedDiagnosticProperties( _
    ByVal properties As Object, _
    ByVal labelText As String, _
    ByVal names As Variant)
    Dim propertyName As Variant
    Dim prop As Object
    Dim propertyType As Variant
    Dim propertyValue As Variant
    Dim lookupErrNumber As Long
    Dim lookupErrSource As String
    Dim lookupErrDescription As String
    Dim typeErrNumber As Long
    Dim typeErrSource As String
    Dim typeErrDescription As String
    Dim valueErrNumber As Long
    Dim valueErrSource As String
    Dim valueErrDescription As String
    Dim typeText As String
    Dim valueText As String
    logLine labelText & " BEGIN"
    For Each propertyName In names
        Set prop = Nothing
        On Error Resume Next
        Err.Clear
        Set prop = properties(CStr(propertyName))
        lookupErrNumber = Err.Number
        lookupErrSource = Err.Source
        lookupErrDescription = Err.Description
        Err.Clear
        On Error GoTo 0
        If lookupErrNumber = 3270 Or lookupErrNumber = 3265 Then
            logLine labelText & ": Name=" & CStr(propertyName) & " | <not present>"
        ElseIf lookupErrNumber <> 0 Then
            logLine labelText & ": Name=" & CStr(propertyName) & " | Lookup=<unreadable: " & _
                DiagnosticErrorText(lookupErrNumber, lookupErrSource, lookupErrDescription) & ">"
        ElseIf prop Is Nothing Then
            logLine labelText & ": Name=" & CStr(propertyName) & " | <not present>"
        Else
            propertyType = Empty
            propertyValue = Empty
            On Error Resume Next
            Err.Clear
            propertyType = prop.Type
            typeErrNumber = Err.Number
            typeErrSource = Err.Source
            typeErrDescription = Err.Description
            Err.Clear
            On Error GoTo 0
            On Error Resume Next
            Err.Clear
            propertyValue = prop.value
            valueErrNumber = Err.Number
            valueErrSource = Err.Source
            valueErrDescription = Err.Description
            Err.Clear
            On Error GoTo 0
            If typeErrNumber = 0 Then
                typeText = DiagnosticValueText(propertyType)
            Else
                typeText = "<unreadable: " & DiagnosticErrorText(typeErrNumber, typeErrSource, typeErrDescription) & ">"
            End If
            If valueErrNumber = 0 Then
                valueText = DiagnosticValueText(propertyValue)
            Else
                valueText = "<unreadable: " & DiagnosticErrorText(valueErrNumber, valueErrSource, valueErrDescription) & ">"
            End If
            logLine labelText & ": Name=" & CStr(propertyName) & " | Type=" & typeText & " | Value=" & valueText
        End If
    Next propertyName
    logLine labelText & " END"
End Sub

Private Function DiagnosticValueText(ByVal value As Variant) As String
    Dim text As String
    On Error GoTo RenderFail
    If IsObject(value) Then
        DiagnosticValueText = "<Object " & typeName(value) & ">"
    ElseIf IsArray(value) Then
        DiagnosticValueText = "<Array VarType=" & CStr(VarType(value)) & ">"
    ElseIf IsNull(value) Then
        DiagnosticValueText = "<Null>"
    ElseIf IsEmpty(value) Then
        DiagnosticValueText = "<Empty>"
    ElseIf IsError(value) Then
        DiagnosticValueText = "<Error value>"
    ElseIf VarType(value) = vbString Then
        text = CStr(value)
        text = Replace(text, vbCr, "\r")
        text = Replace(text, vbLf, "\n")
        text = Replace(text, vbTab, "\t")
        DiagnosticValueText = """" & text & """"
    ElseIf VarType(value) = vbDate Then
        DiagnosticValueText = Format$(CDate(value), "yyyy-mm-dd hh:nn:ss")
    Else
        DiagnosticValueText = CStr(value)
    End If
    Exit Function
RenderFail:
    DiagnosticValueText = "<render failed: Error " & CStr(Err.Number) & " | Source=" & Err.Source & _
        " | Description=" & Err.Description & ">"
    Err.Clear
End Function

Private Function DiagnosticErrorText( _
    ByVal errNumber As Long, _
    ByVal errSource As String, _
    ByVal errDescription As String) As String
    If Len(errSource) = 0 Then errSource = "<empty>"
    If Len(errDescription) = 0 Then errDescription = "<empty>"
    DiagnosticErrorText = "Error " & CStr(errNumber) & " | Source=" & errSource & " | Description=" & errDescription
End Function

Private Sub ClearADOProviderErrors()
    Dim conn As Object
    On Error Resume Next
    Set conn = CurrentProject.Connection
    If Not conn Is Nothing Then conn.Errors.Clear
End Sub

Private Sub LogADOProviderErrors(ByVal prefix As String)
    Dim conn As Object
    On Error Resume Next
    Set conn = CurrentProject.Connection
    If conn Is Nothing Then
        logLine prefix & ": CurrentProject.Connection is unavailable."
        Exit Sub
    End If
    LogADOProviderErrorsFromConnection conn, prefix
End Sub

Private Sub LogADOProviderErrorsFromConnection(ByVal conn As Object, ByVal prefix As String)
    Dim providerError As Object
    Dim lineText As String
    On Error Resume Next
    If conn Is Nothing Then
        logLine prefix & ": connection is unavailable."
        Exit Sub
    End If
    If conn.Errors.count = 0 Then
        logLine prefix & ": no provider-specific errors were exposed."
        Exit Sub
    End If
    For Each providerError In conn.Errors
        lineText = prefix & ": Number=" & CStr(providerError.Number) & _
            " | NativeError=" & CStr(providerError.NativeError) & _
            " | SQLState=" & CStr(providerError.SQLState) & _
            " | Source=" & CStr(providerError.Source) & _
            " | Description=" & CStr(providerError.Description)
        logLine lineText
    Next providerError
End Sub

Private Function VerifyFieldPlan(ByVal fieldPlan As Collection, ByRef issues As Collection) As Boolean
    Dim inventories As Object
    Dim fieldInventories As Object
    Dim item As Object
    Dim inv As Object
    Dim i As Long
    Dim parentType As String
    Dim parentName As String
    Dim parentKey As String
    Dim oldName As String
    Dim newName As String
    Dim actionText As String
    Dim problem As String
    Dim ok As Boolean
    ok = True
    Set inventories = BuildInventories()
    Set fieldInventories = CreateTextDictionary()
    For i = 1 To fieldPlan.count
        Set item = fieldPlan(i)
        actionText = CStr(item("Action"))
        If actionText <> "Missing" And actionText <> "MissingParent" Then
            parentType = CStr(item("ParentType"))
            If inventories(parentType).Exists(CStr(item("ParentNew"))) Then
                parentName = InventoryActualName(inventories(parentType), CStr(item("ParentNew")))
            ElseIf inventories(parentType).Exists(CStr(item("ParentOld"))) Then
                parentName = InventoryActualName(inventories(parentType), CStr(item("ParentOld")))
            Else
                AddIssue issues, parentType & " field parent is missing during verification: '" & CStr(item("ParentNew")) & "'."
                ok = False
                GoTo NextField
            End If
            parentKey = FieldParentKey(parentType, parentName)
            If Not fieldInventories.Exists(parentKey) Then
                Set inv = Nothing
                problem = vbNullString
                If Not TryBuildFieldInventory(parentType, parentName, inv, problem, True) Then
                    AddIssue issues, parentType & " '" & parentName & "' fields could not be verified: " & problem
                    ok = False
                    GoTo NextField
                End If
                fieldInventories.Add parentKey, inv
            Else
                Set inv = fieldInventories(parentKey)
            End If
            oldName = CStr(item("Old"))
            newName = CStr(item("New"))
            Select Case actionText
                Case "Done", "Already"
                    If Not inv.Exists(newName) Then
                        AddIssue issues, parentType & " '" & parentName & "' expected field is missing: '" & newName & "'."
                        ok = False
                    ElseIf StrComp(InventoryActualName(inv, newName), newName, vbBinaryCompare) <> 0 Then
                        AddIssue issues, parentType & " '" & parentName & "' field casing does not match: expected '" & newName & "'."
                        ok = False
                    End If
                    If StrComp(oldName, newName, vbTextCompare) <> 0 Then
                        If inv.Exists(oldName) Then
                            AddIssue issues, parentType & " '" & parentName & "' old field still exists: '" & oldName & "'."
                            ok = False
                        End If
                    End If
                    If item.Exists("Calculated") Then
                        If CBool(item("Calculated")) Then
                            problem = vbNullString
                            If Not TryGetCalculatedFieldExpression(parentName, newName, problem) Then
                                AddIssue issues, "Table '" & parentName & "' renamed calculated field is no longer calculated: '" & newName & "'."
                                ok = False
                            End If
                        End If
                    End If
                Case "Unchanged"
                    If Not inv.Exists(newName) Then
                        AddIssue issues, parentType & " '" & parentName & "' unchanged field is missing: '" & newName & "'."
                        ok = False
                    End If
                Case "Conflict"
                    AddIssue issues, parentType & " '" & parentName & "' field remained in conflict: '" & oldName & "'."
                    ok = False
                Case Else
                    AddIssue issues, parentType & " '" & parentName & "' has unexpected field action '" & actionText & "'."
                    ok = False
            End Select
        End If
NextField:
    Next i
    VerifyFieldPlan = ok
End Function

Private Sub ApplyPlan( _
    ByVal plan As Collection, _
    ByVal inventories As Object, _
    ByRef journal As Collection, _
    ByRef tempNames As Collection, _
    ByVal stats As Object)

    Dim pending As Object
    Dim reserved As Object
    Dim item As Object
    Dim keys As Variant
    Dim key As Variant
    Dim targetKey As String
    Dim oldKey As String
    Dim typeName As String
    Dim workOld As String
    Dim newName As String
    Dim tempName As String
    Dim i As Long
    Dim initialPending As Long
    Dim completed As Long
    Dim didFinalRename As Boolean
    Dim pct As Double

    Set pending = CreateTextDictionary()
    Set reserved = BuildReservedNameSet(plan)

    For i = 1 To plan.count
        Set item = plan(i)
        If StrComp(CStr(item("Action")), "Pending", vbTextCompare) = 0 Then
            pending.Add NamespaceNameKey(CStr(item("Type")), CStr(item("WorkOld"))), item
        End If
    Next i

    initialPending = pending.count
    If initialPending = 0 Then
        ProgressAppend "No physical renames are required."
        ProgressUpdate "Applying object renames", 95, "Everything is already renamed, unchanged, or missing.", True
        Exit Sub
    End If

    Do While pending.count > 0
        If mCancelRequested Then
            Err.Raise vbObjectError + 7100, "ApplyPlan", "Cancelled by user."
        End If

        didFinalRename = False
        keys = pending.keys

        For Each key In keys
            Set item = pending(CStr(key))
            typeName = CStr(item("Type"))
            workOld = CStr(item("WorkOld"))
            newName = CStr(item("New"))
            targetKey = NamespaceNameKey(typeName, newName)

            If Not pending.Exists(targetKey) Then
                ProgressAppend "ATTEMPT Object rename: " & typeName & " | " & workOld & " -> " & newName
                ClearADOProviderErrors
                RenameOne typeName, workOld, newName
                journal.Add Array(typeName, workOld, newName)
                UpdateInventory inventories(typeName), workOld, newName
                If StrComp(typeName, "Table", vbTextCompare) = 0 Or StrComp(typeName, "Query", vbTextCompare) = 0 Then
                    RewriteQueryReferencesAfterObjectRename typeName, workOld, newName, journal
                End If

                pending.Remove CStr(key)
                item("WorkOld") = newName
                item("Action") = "Done"
                completed = completed + 1
                StatInc stats, "Renamed"

                ProgressAppend "RENAMED " & typeName & ": " & workOld & " -> " & newName
                pct = 55# + (39# * completed / initialPending)
                ProgressUpdate "Applying object renames", pct, _
                    Format$(completed, "#,##0") & " / " & Format$(initialPending, "#,##0") & _
                    " | " & typeName & vbCrLf & workOld & " -> " & newName, False

                didFinalRename = True
                Exit For
            End If
        Next key

        If Not didFinalRename Then
            keys = pending.keys
            key = keys(LBound(keys))
            Set item = pending(CStr(key))
            typeName = CStr(item("Type"))
            workOld = CStr(item("WorkOld"))

            tempName = NextTemporaryName(typeName, inventories, reserved)
            ProgressAppend "ATTEMPT Object cycle-break rename: " & typeName & " | " & workOld & " -> " & tempName
            ClearADOProviderErrors
            RenameOne typeName, workOld, tempName
            journal.Add Array(typeName, workOld, tempName)
            tempNames.Add Array(typeName, tempName)
            UpdateInventory inventories(typeName), workOld, tempName
            If StrComp(typeName, "Table", vbTextCompare) = 0 Or StrComp(typeName, "Query", vbTextCompare) = 0 Then
                RewriteQueryReferencesAfterObjectRename typeName, workOld, tempName, journal
            End If

            oldKey = CStr(key)
            pending.Remove oldKey
            item("WorkOld") = tempName
            pending.Add NamespaceNameKey(typeName, tempName), item
            reserved.Add NamespaceNameKey(typeName, tempName), True

            StatInc stats, "TemporaryRenames"
            ProgressAppend "CYCLE BREAK " & typeName & ": " & workOld & " -> " & tempName
            ProgressUpdate "Resolving object rename cycle", 55# + (39# * completed / initialPending), _
                typeName & vbCrLf & workOld & " -> " & tempName, True
        End If
    Loop

    ProgressUpdate "Applying object renames", 94, _
        Format$(completed, "#,##0") & " physical target rename(s) complete.", True
End Sub

Private Function BuildReservedNameSet(ByVal plan As Collection) As Object
    Dim reserved As Object
    Dim item As Object
    Dim i As Long
    Dim key As String

    Set reserved = CreateTextDictionary()
    For i = 1 To plan.count
        Set item = plan(i)
        key = NamespaceNameKey(CStr(item("Type")), CStr(item("Old")))
        If Not reserved.Exists(key) Then reserved.Add key, True
        key = NamespaceNameKey(CStr(item("Type")), CStr(item("New")))
        If Not reserved.Exists(key) Then reserved.Add key, True
    Next i
    Set BuildReservedNameSet = reserved
End Function

Private Function NextTemporaryName(ByVal typeName As String, ByVal inventories As Object, ByVal reserved As Object) As String
    Dim candidate As String
    Dim key As String

    Do
        mTempSerial = mTempSerial + 1
        candidate = TEMP_PREFIX & mSessionTag & "_" & Format$(mTempSerial, "0000")
        key = NamespaceNameKey(typeName, candidate)
    Loop While NameExistsInNamespace(inventories, typeName, candidate) Or reserved.Exists(key)

    NextTemporaryName = candidate
End Function

Private Sub RenameOne(ByVal typeName As String, ByVal oldName As String, ByVal newName As String)
    CloseObjectIfOpen typeName, oldName
    CloseObjectIfOpen typeName, newName
    DoCmd.Rename newName, AccessObjectType(typeName), oldName
End Sub

Private Sub CloseObjectIfOpen(ByVal typeName As String, ByVal objectName As String)
    If Len(objectName) = 0 Then Exit Sub
    If StrComp(typeName, "Form", vbTextCompare) = 0 And _
       StrComp(objectName, PROGRESS_FORM, vbTextCompare) = 0 Then Exit Sub

    On Error Resume Next
    DoCmd.Close AccessObjectType(typeName), objectName, acSaveNo
    On Error GoTo 0
End Sub

Private Function AccessObjectType(ByVal typeName As String) As Long
    Select Case typeName
        Case "Table": AccessObjectType = acTable
        Case "Query": AccessObjectType = acQuery
        Case "Form": AccessObjectType = acForm
        Case "Report": AccessObjectType = acReport
        Case "Macro": AccessObjectType = acMacro
        Case "Module": AccessObjectType = acModule
        Case Else
            Err.Raise vbObjectError + 7002, "AccessObjectType", "Unsupported Access object type: " & typeName
    End Select
End Function

Private Sub UpdateInventory(ByVal inv As Object, ByVal oldName As String, ByVal newName As String)
    If inv.Exists(oldName) Then inv.Remove oldName
    If inv.Exists(newName) Then
        inv(newName) = newName
    Else
        inv.Add newName, newName
    End If
End Sub

Private Function VerifyPlan( _
    ByVal plan As Collection, _
    ByVal tempNames As Collection, _
    ByRef issues As Collection) As Boolean

    Dim inventories As Object
    Dim item As Object
    Dim inv As Object
    Dim i As Long
    Dim typeName As String
    Dim oldName As String
    Dim newName As String
    Dim actionText As String
    Dim stepData As Variant

    Set inventories = BuildInventories()

    For i = 1 To plan.count
        Set item = plan(i)
        typeName = CStr(item("Type"))
        oldName = CStr(item("Old"))
        newName = CStr(item("New"))
        actionText = CStr(item("Action"))
        Set inv = inventories(typeName)

        Select Case actionText
            Case "Done", "Already"
                If Not inv.Exists(newName) Then
                    AddIssue issues, typeName & " expected target is missing: '" & newName & "'."
                ElseIf StrComp(InventoryActualName(inv, newName), newName, vbBinaryCompare) <> 0 Then
                    AddIssue issues, typeName & " target casing does not match: expected '" & newName & _
                        "', found '" & InventoryActualName(inv, newName) & "'."
                End If

                If StrComp(oldName, newName, vbTextCompare) <> 0 Then
                    If inv.Exists(oldName) Then
                        AddIssue issues, typeName & " old name still exists after rename: '" & oldName & "'."
                    End If
                End If

            Case "Unchanged"
                If Not inv.Exists(newName) Then
                    AddIssue issues, typeName & " unchanged object is missing: '" & newName & "'."
                End If

            Case "Missing"
                ' Missing objects were warnings during preflight and are not verification failures.

            Case Else
                AddIssue issues, typeName & " has unexpected final action state '" & actionText & _
                    "' for '" & oldName & "'."
        End Select
    Next i

    For i = 1 To tempNames.count
        stepData = tempNames(i)
        typeName = CStr(stepData(0))
        newName = CStr(stepData(1))
        Set inv = inventories(typeName)
        If inv.Exists(newName) Then
            AddIssue issues, "Temporary rename object still exists: " & typeName & " '" & newName & "'."
        End If
    Next i

    VerifyPlan = (issues.count = 0)
End Function

Private Function RollbackJournal(ByVal journal As Collection) As Long
    Dim i As Long
    Dim firstIndex As Long
    Dim stepData As Variant
    Dim previousStepData As Variant
    Dim failCount As Long
    Dim tableName As String
    Dim objectRolledBack As Boolean
    On Error GoTo UnexpectedRollbackError
    i = journal.count
    Do While i >= 1
        stepData = journal(i)
        If StrComp(CStr(stepData(0)), "QuerySourceSQL", vbTextCompare) = 0 Then
            firstIndex = i
            Do While firstIndex > 1
                previousStepData = journal(firstIndex - 1)
                If StrComp(CStr(previousStepData(0)), "QuerySourceSQL", vbTextCompare) <> 0 Then Exit Do
                If StrComp(CStr(previousStepData(4)), CStr(stepData(4)), vbTextCompare) <> 0 Or _
                   StrComp(CStr(previousStepData(5)), CStr(stepData(5)), vbTextCompare) <> 0 Or _
                   StrComp(CStr(previousStepData(6)), CStr(stepData(6)), vbTextCompare) <> 0 Then Exit Do
                firstIndex = firstIndex - 1
            Loop
            If mProgressOpen Then
                ProgressUpdate "Rollback", 100# * (journal.count - i) / journal.count, _
                    "Object/query references" & vbCrLf & CStr(stepData(5)) & " -> " & CStr(stepData(6)), False
            End If
            objectRolledBack = False
            If Not RollbackObjectQueryReferenceBatch(journal, firstIndex, i, objectRolledBack) Then failCount = failCount + 1
            If objectRolledBack Then
                i = firstIndex - 2
            Else
                i = firstIndex - 1
            End If
        ElseIf StrComp(CStr(stepData(0)), "TableFieldQuerySnapshot", vbTextCompare) = 0 Then
            firstIndex = i
            Do While firstIndex > 1
                previousStepData = journal(firstIndex - 1)
                If StrComp(CStr(previousStepData(0)), "TableFieldQuerySnapshot", vbTextCompare) <> 0 Then Exit Do
                firstIndex = firstIndex - 1
            Loop
            If mProgressOpen Then
                ProgressUpdate "Rollback", 100# * (journal.count - i) / journal.count, _
                    "Query pre-table-field SQL snapshots", False
            End If
            If Not RollbackTableFieldQuerySnapshotBatch(journal, firstIndex, i) Then failCount = failCount + 1
            i = firstIndex - 1
        ElseIf StrComp(CStr(stepData(0)), "QueryFieldRefSQL", vbTextCompare) = 0 Then
            firstIndex = i
            Do While firstIndex > 1
                previousStepData = journal(firstIndex - 1)
                If StrComp(CStr(previousStepData(0)), "QueryFieldRefSQL", vbTextCompare) <> 0 Then Exit Do
                If StrComp(CStr(previousStepData(4)), CStr(stepData(4)), vbTextCompare) <> 0 Then Exit Do
                firstIndex = firstIndex - 1
            Loop
            If mProgressOpen Then
                ProgressUpdate "Rollback", 100# * (journal.count - i) / journal.count, _
                    "Query output dependencies" & vbCrLf & CStr(stepData(4)), False
            End If
            objectRolledBack = False
            If Not RollbackQueryFieldReferenceBatch(journal, firstIndex, i, objectRolledBack) Then failCount = failCount + 1
            If objectRolledBack Then
                i = firstIndex - 2
            Else
                i = firstIndex - 1
            End If
        ElseIf StrComp(CStr(stepData(0)), "TableField", vbTextCompare) = 0 Then
            tableName = CStr(stepData(1))
            firstIndex = i
            Do While firstIndex > 1
                previousStepData = journal(firstIndex - 1)
                If StrComp(CStr(previousStepData(0)), "TableField", vbTextCompare) <> 0 Then Exit Do
                If StrComp(CStr(previousStepData(1)), tableName, vbTextCompare) <> 0 Then Exit Do
                firstIndex = firstIndex - 1
            Loop
            If mProgressOpen Then
                ProgressUpdate "Rollback", 100# * (journal.count - i) / journal.count, _
                    "Table fields" & vbCrLf & tableName, False
            End If
            If Not RollbackTableFieldJournalBatch(journal, firstIndex, i) Then failCount = failCount + 1
            i = firstIndex - 1
        Else
            If mProgressOpen Then
                ProgressUpdate "Rollback", 100# * (journal.count - i) / journal.count, RollbackStepText(stepData), False
            End If
            If Not RollbackJournalStep(stepData) Then failCount = failCount + 1
            i = i - 1
        End If
    Loop
    RollbackJournal = failCount
    Exit Function
UnexpectedRollbackError:
    logLine "ROLLBACK ERROR " & CStr(Err.Number) & ": " & Err.Description
    RollbackJournal = failCount + 1
End Function

Private Function RollbackQueryFieldReferenceBatch( _
    ByVal journal As Collection, _
    ByVal firstIndex As Long, _
    ByVal lastIndex As Long, _
    ByRef sourceRolledBack As Boolean) As Boolean
    ' Restore the source query's old output aliases first, then restore dependent QueryDef SQL that referenced them.
    Dim sourceIndex As Long
    Dim sourceStep As Variant
    Dim stepData As Variant
    Dim sourceQueryName As String
    Dim i As Long
    Dim failCount As Long
    sourceRolledBack = False
    sourceIndex = firstIndex - 1
    If sourceIndex < 1 Then
        logLine "ROLLBACK FAILED Query field-reference group | Missing preceding QuerySQL journal step."
        Exit Function
    End If
    stepData = journal(firstIndex)
    sourceQueryName = CStr(stepData(4))
    sourceStep = journal(sourceIndex)
    If StrComp(CStr(sourceStep(0)), "QuerySQL", vbTextCompare) <> 0 Or _
       StrComp(CStr(sourceStep(1)), sourceQueryName, vbTextCompare) <> 0 Then
        logLine "ROLLBACK FAILED Query field-reference group | Preceding QuerySQL step does not match source query '" & _
            sourceQueryName & "'."
        Exit Function
    End If
    If Not RollbackQuerySqlStep(sourceStep) Then Exit Function
    sourceRolledBack = True
    For i = lastIndex To firstIndex Step -1
        stepData = journal(i)
        If Not RollbackQueryFieldRefSqlStep(stepData) Then failCount = failCount + 1
    Next i
    RollbackQueryFieldReferenceBatch = (failCount = 0)
End Function

Private Function RollbackQuerySqlStep(ByVal stepData As Variant) As Boolean
    Dim db As Object
    Dim qdf As Object
    On Error GoTo RollbackFail
    CloseObjectIfOpen "Query", CStr(stepData(1))
    Set db = currentDb
    Set qdf = db.QueryDefs(CStr(stepData(1)))
    qdf.sql = CStr(stepData(2))
    logLine "ROLLBACK Query fields: " & CStr(stepData(1))
    RollbackQuerySqlStep = True
    Exit Function
RollbackFail:
    logLine "ROLLBACK FAILED Query fields: " & CStr(stepData(1)) & " | Error " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function RollbackQueryFieldRefSqlStep(ByVal stepData As Variant) As Boolean
    Dim db As Object
    Dim qdf As Object
    On Error GoTo RollbackFail
    CloseObjectIfOpen "Query", CStr(stepData(1))
    Set db = currentDb
    Set qdf = db.QueryDefs(CStr(stepData(1)))
    qdf.sql = CStr(stepData(2))
    logLine "ROLLBACK Query field references: " & CStr(stepData(1)) & " | Source query: " & CStr(stepData(4))
    RollbackQueryFieldRefSqlStep = True
    Exit Function
RollbackFail:
    logLine "ROLLBACK FAILED Query field references: " & CStr(stepData(1)) & " | Error " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function RollbackTableFieldJournalBatch( _
    ByVal journal As Collection, _
    ByVal firstIndex As Long, _
    ByVal lastIndex As Long) As Boolean
    Dim definitions As Collection
    Dim Definition As Collection
    Dim db As Object
    Dim tdf As Object
    Dim cat As Object
    Dim conn As Object
    Dim stepData As Variant
    Dim tableName As String
    Dim currentCalculatedName As String
    Dim originalCalculatedName As String
    Dim currentExpression As String
    Dim workingExpression As String
    Dim nextExpression As String
    Dim probeExpression As String
    Dim propertyProblem As String
    Dim errDescription As String
    Dim errSource As String
    Dim restoreProblem As String
    Dim recoverySourceName As String
    Dim rebuiltCalculatedField As Boolean
    Dim calculatedNameEventApplied As Boolean
    Dim suspensionStarted As Boolean
    Dim stepFailed As Boolean
    Dim calculatedStepIndex As Long
    Dim errNumber As Long
    Dim restoreErr As Long
    Dim i As Long
    stepData = journal(firstIndex)
    tableName = CStr(stepData(1))
    On Error GoTo BatchFail
    Set definitions = CaptureTableCalculatedFieldDefinitions(tableName)
    If definitions.count = 0 Then
        For i = lastIndex To firstIndex Step -1
            stepData = journal(i)
            If Not RollbackJournalStep(stepData) Then stepFailed = True
        Next i
        RollbackTableFieldJournalBatch = Not stepFailed
        Exit Function
    End If
    If definitions.count > 1 Then
        Err.Raise vbObjectError + 7167, "RollbackTableFieldJournalBatch", _
            "Table '" & tableName & "' contains multiple calculated fields; grouped rollback cannot prove dependency-safe restoration order."
    End If
    Set Definition = definitions(1)
    currentCalculatedName = CStr(Definition("Name"))
    originalCalculatedName = currentCalculatedName
    currentExpression = CStr(Definition("Expression"))
    workingExpression = currentExpression
    For i = firstIndex To lastIndex
        stepData = journal(i)
        If StrComp(CStr(stepData(3)), currentCalculatedName, vbTextCompare) = 0 Then
            calculatedStepIndex = i
            originalCalculatedName = CStr(stepData(2))
            Exit For
        End If
    Next i
    probeExpression = currentExpression
    For i = lastIndex To firstIndex Step -1
        If i <> calculatedStepIndex Then
            stepData = journal(i)
            propertyProblem = CalculatedFieldStoredPropertyReferenceProblem(tableName, Definition, CStr(stepData(3)))
            If Len(propertyProblem) > 0 Then Err.Raise vbObjectError + 7159, "RollbackTableFieldJournalBatch", propertyProblem
            probeExpression = RewriteCalculatedFieldReference(probeExpression, CStr(stepData(3)), CStr(stepData(2)))
        End If
    Next i
    Set conn = CurrentProject.Connection
    CloseObjectIfOpen "Table", tableName
    Set db = currentDb
    Set tdf = db.TableDefs(tableName)
    suspensionStarted = True
    DeleteCapturedCalculatedFields tdf, definitions
    EnsureSuspendedCalculatedFieldPlaceholder tableName, Definition, currentCalculatedName
    Set cat = CreateObject("ADOX.Catalog")
    Set cat.ActiveConnection = conn
    cat.tables.Refresh
    cat.tables(tableName).Columns.Refresh
    For i = lastIndex To firstIndex Step -1
        If i <> calculatedStepIndex Then
            stepData = journal(i)
            nextExpression = RewriteCalculatedFieldReference(workingExpression, CStr(stepData(3)), CStr(stepData(2)))
            RenameTableField cat, conn, tableName, CStr(stepData(3)), CStr(stepData(2)), rebuiltCalculatedField
            If rebuiltCalculatedField Then Err.Raise vbObjectError + 7164, "RollbackTableFieldJournalBatch", _
                "Calculated field unexpectedly existed during grouped ordinary-field rollback."
            workingExpression = nextExpression
            logLine "ROLLBACK Table field: " & tableName & "." & CStr(stepData(3)) & " -> " & CStr(stepData(2))
        End If
    Next i
    RestoreSuspendedCalculatedField tableName, conn, cat, Definition, currentCalculatedName, originalCalculatedName, _
        workingExpression, calculatedNameEventApplied
    If calculatedStepIndex > 0 Then
        logLine "ROLLBACK Table field: " & tableName & "." & currentCalculatedName & " -> " & originalCalculatedName
    End If
    RollbackTableFieldJournalBatch = True
    Exit Function
BatchFail:
    errNumber = Err.Number
    errDescription = Err.Description
    errSource = Err.Source
    If suspensionStarted Then
        On Error Resume Next
        If calculatedNameEventApplied Then
            recoverySourceName = originalCalculatedName
        Else
            recoverySourceName = currentCalculatedName
        End If
        Set db = currentDb
        Set tdf = db.TableDefs(tableName)
        If StrComp(currentCalculatedName, originalCalculatedName, vbTextCompare) <> 0 Then
            If StrComp(recoverySourceName, currentCalculatedName, vbTextCompare) <> 0 And DAOFieldExists(tdf, currentCalculatedName) Then
                tdf.fields.Delete currentCalculatedName
            End If
            If StrComp(recoverySourceName, originalCalculatedName, vbTextCompare) <> 0 And DAOFieldExists(tdf, originalCalculatedName) Then
                tdf.fields.Delete originalCalculatedName
            End If
            tdf.fields.Refresh
        End If
        EnsureSuspendedCalculatedFieldPlaceholder tableName, Definition, recoverySourceName
        Set cat = Nothing
        Set cat = CreateObject("ADOX.Catalog")
        Set cat.ActiveConnection = conn
        cat.tables.Refresh
        cat.tables(tableName).Columns.Refresh
        Err.Clear
        calculatedNameEventApplied = False
        RestoreSuspendedCalculatedField tableName, conn, cat, Definition, recoverySourceName, currentCalculatedName, _
            workingExpression, calculatedNameEventApplied
        restoreErr = Err.Number
        restoreProblem = Err.Description
        Err.Clear
        On Error GoTo 0
        If restoreErr <> 0 Then errDescription = errDescription & " Calculated-field recovery after rollback failure also failed: " & restoreProblem
    End If
    logLine "ROLLBACK FAILED Table fields" & vbCrLf & tableName & " | Error " & CStr(errNumber) & ": " & errDescription & _
        IIf(Len(errSource) > 0, " | Source: " & errSource, vbNullString)
End Function

Private Function RollbackObjectQueryReferenceBatch( _
    ByVal journal As Collection, _
    ByVal firstIndex As Long, _
    ByVal lastIndex As Long, _
    ByRef objectRolledBack As Boolean) As Boolean
    ' Reverse one physical table/query rename first, then restore every QueryDef SQL change that tracked that exact physical rename.
    Dim objectIndex As Long
    Dim objectStep As Variant
    Dim stepData As Variant
    Dim i As Long
    Dim failCount As Long
    objectRolledBack = False
    objectIndex = firstIndex - 1
    If objectIndex < 1 Then
        logLine "ROLLBACK FAILED Query source SQL group | Missing preceding object rename journal step."
        Exit Function
    End If
    stepData = journal(firstIndex)
    objectStep = journal(objectIndex)
    If StrComp(CStr(objectStep(0)), CStr(stepData(4)), vbTextCompare) <> 0 Or _
       StrComp(CStr(objectStep(1)), CStr(stepData(5)), vbTextCompare) <> 0 Or _
       StrComp(CStr(objectStep(2)), CStr(stepData(6)), vbTextCompare) <> 0 Then
        logLine "ROLLBACK FAILED Query source SQL group | Preceding object rename journal step does not match source rename '" & _
            CStr(stepData(5)) & "' -> '" & CStr(stepData(6)) & "'."
        Exit Function
    End If
    If Not RollbackOne(CStr(objectStep(0)), CStr(objectStep(1)), CStr(objectStep(2))) Then Exit Function
    objectRolledBack = True
    For i = lastIndex To firstIndex Step -1
        stepData = journal(i)
        If Not RollbackQuerySourceSqlStep(stepData) Then failCount = failCount + 1
    Next i
    RollbackObjectQueryReferenceBatch = (failCount = 0)
End Function

Private Function RollbackQuerySourceSqlStep(ByVal stepData As Variant) As Boolean
    ' Restore one QueryDef SQL text after its source object has first been returned to the pre-rename physical name.
    Dim db As Object
    Dim qdf As Object
    Dim queryName As String
    On Error GoTo RollbackFail
    queryName = CStr(stepData(1))
    If StrComp(CStr(stepData(4)), "Query", vbTextCompare) = 0 And _
       StrComp(queryName, CStr(stepData(6)), vbTextCompare) = 0 Then
        queryName = CStr(stepData(5))
    End If
    CloseObjectIfOpen "Query", queryName
    Set db = currentDb
    Set qdf = db.QueryDefs(queryName)
    qdf.sql = CStr(stepData(2))
    logLine "ROLLBACK Query source SQL: " & queryName & " | " & CStr(stepData(6)) & " -> " & CStr(stepData(5))
    RollbackQuerySourceSqlStep = True
    Exit Function
RollbackFail:
    logLine "ROLLBACK FAILED Query source SQL: " & queryName & " | Error " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function RollbackTableFieldQuerySnapshotBatch( _
    ByVal journal As Collection, _
    ByVal firstIndex As Long, _
    ByVal lastIndex As Long) As Boolean
    ' Restore all pre-table QueryDef SQL text before refreshing any one query so rollback does not depend on QueryDef dependency order.
    Dim db As Object
    Dim qdf As Object
    Dim assigned As Object
    Dim stepData As Variant
    Dim queryName As String
    Dim errNumber As Long
    Dim errDescription As String
    Dim failed As Boolean
    Dim i As Long
    Set db = currentDb
    Set assigned = CreateTextDictionary()
    For i = firstIndex To lastIndex
        stepData = journal(i)
        queryName = CStr(stepData(1))
        On Error Resume Next
        Err.Clear
        CloseObjectIfOpen "Query", queryName
        Err.Clear
        Set qdf = Nothing
        Set qdf = db.QueryDefs(queryName)
        qdf.sql = CStr(stepData(2))
        errNumber = Err.Number
        errDescription = Err.Description
        Err.Clear
        On Error GoTo 0
        If errNumber = 0 Then
            assigned.Add queryName, True
        Else
            failed = True
            logLine "ROLLBACK FAILED Query pre-table-field SQL snapshot assignment: " & queryName & _
                " | Error " & CStr(errNumber) & ": " & errDescription
        End If
    Next i
    For i = firstIndex To lastIndex
        stepData = journal(i)
        queryName = CStr(stepData(1))
        If assigned.Exists(queryName) Then
            On Error Resume Next
            Err.Clear
            Set qdf = Nothing
            Set qdf = db.QueryDefs(queryName)
            qdf.fields.Refresh
            errNumber = Err.Number
            errDescription = Err.Description
            Err.Clear
            On Error GoTo 0
            If errNumber = 0 Then
                logLine "ROLLBACK Query pre-table-field SQL snapshot: " & queryName
            Else
                failed = True
                logLine "ROLLBACK FAILED Query pre-table-field SQL snapshot validation: " & queryName & _
                    " | Error " & CStr(errNumber) & ": " & errDescription
            End If
        End If
    Next i
    RollbackTableFieldQuerySnapshotBatch = Not failed
End Function

Private Function RollbackJournalStep(ByVal stepData As Variant) As Boolean
    Dim stepType As String
    Dim cat As Object
    Dim conn As Object
    Dim rebuiltCalculatedField As Boolean
    On Error GoTo RollbackFail
    stepType = CStr(stepData(0))
    If StrComp(stepType, "QuerySQL", vbTextCompare) = 0 Then
        If Not RollbackQuerySqlStep(stepData) Then Exit Function
    ElseIf StrComp(stepType, "QueryFieldRefSQL", vbTextCompare) = 0 Then
        logLine "ROLLBACK FAILED Query field-reference SQL step was encountered outside its source-query rollback group: " & CStr(stepData(1))
        Exit Function
    ElseIf StrComp(stepType, "QuerySourceSQL", vbTextCompare) = 0 Then
        logLine "ROLLBACK FAILED Query source SQL step was encountered outside its object-rename rollback group: " & CStr(stepData(1))
        Exit Function
    ElseIf StrComp(stepType, "TableFieldQuerySnapshot", vbTextCompare) = 0 Then
        logLine "ROLLBACK FAILED Query pre-table-field SQL snapshot step was encountered outside its rollback batch: " & CStr(stepData(1))
        Exit Function
    ElseIf StrComp(stepType, "TableField", vbTextCompare) = 0 Then
        Set conn = CurrentProject.Connection
        Set cat = CreateObject("ADOX.Catalog")
        Set cat.ActiveConnection = conn
        cat.tables.Refresh
        RenameTableField cat, conn, CStr(stepData(1)), CStr(stepData(3)), CStr(stepData(2)), rebuiltCalculatedField
        logLine "ROLLBACK Table field: " & CStr(stepData(1)) & "." & CStr(stepData(3)) & " -> " & CStr(stepData(2))
    Else
        If Not RollbackOne(CStr(stepData(0)), CStr(stepData(1)), CStr(stepData(2))) Then Exit Function
    End If
    RollbackJournalStep = True
    Exit Function
RollbackFail:
    logLine "ROLLBACK FAILED " & RollbackStepText(stepData) & " | Error " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function RollbackStepText(ByVal stepData As Variant) As String
    Dim stepType As String
    stepType = CStr(stepData(0))
    If StrComp(stepType, "QuerySQL", vbTextCompare) = 0 Then
        RollbackStepText = "Query fields" & vbCrLf & CStr(stepData(1))
    ElseIf StrComp(stepType, "QueryFieldRefSQL", vbTextCompare) = 0 Then
        RollbackStepText = "Query field references" & vbCrLf & CStr(stepData(1)) & " | Source: " & CStr(stepData(4))
    ElseIf StrComp(stepType, "QuerySourceSQL", vbTextCompare) = 0 Then
        RollbackStepText = "Query source SQL" & vbCrLf & CStr(stepData(1))
    ElseIf StrComp(stepType, "TableFieldQuerySnapshot", vbTextCompare) = 0 Then
        RollbackStepText = "Query pre-table-field SQL snapshot" & vbCrLf & CStr(stepData(1))
    ElseIf StrComp(stepType, "TableField", vbTextCompare) = 0 Then
        RollbackStepText = "Table field" & vbCrLf & CStr(stepData(1)) & "." & CStr(stepData(3)) & " -> " & CStr(stepData(2))
    Else
        RollbackStepText = stepType & vbCrLf & CStr(stepData(2)) & " -> " & CStr(stepData(1))
    End If
End Function

Private Function RollbackOne(ByVal typeName As String, ByVal oldName As String, ByVal newName As String) As Boolean
    On Error GoTo RollbackFail
    CloseObjectIfOpen typeName, newName
    CloseObjectIfOpen typeName, oldName
    DoCmd.Rename oldName, AccessObjectType(typeName), newName
    logLine "ROLLBACK " & typeName & ": " & newName & " -> " & oldName
    RollbackOne = True
    Exit Function

RollbackFail:
    logLine "ROLLBACK FAILED " & typeName & ": " & newName & " -> " & oldName & _
        " | Error " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function InventoryActualName(ByVal inv As Object, ByVal lookupName As String) As String
    If inv.Exists(lookupName) Then InventoryActualName = CStr(inv(lookupName))
End Function

Private Function InventoryHasExpectedSource(ByVal inventories As Object, ByVal item As Object) As Boolean
    InventoryHasExpectedSource = inventories(CStr(item("Type"))).Exists(CStr(item("Old")))
End Function

Private Function NameExistsInNamespace( _
    ByVal inventories As Object, _
    ByVal typeName As String, _
    ByVal objectName As String) As Boolean

    If StrComp(typeName, "Table", vbTextCompare) = 0 Or _
       StrComp(typeName, "Query", vbTextCompare) = 0 Then
        NameExistsInNamespace = inventories("Table").Exists(objectName) Or _
            inventories("Query").Exists(objectName)
    Else
        NameExistsInNamespace = inventories(typeName).Exists(objectName)
    End If
End Function

Private Function NamespaceNameKey(ByVal typeName As String, ByVal objectName As String) As String
    If StrComp(typeName, "Table", vbTextCompare) = 0 Or _
       StrComp(typeName, "Query", vbTextCompare) = 0 Then
        NamespaceNameKey = "TableOrQuery" & Chr$(30) & objectName
    Else
        NamespaceNameKey = typeName & Chr$(30) & objectName
    End If
End Function

Private Function CreateTextDictionary() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set CreateTextDictionary = d
End Function

Private Sub AddIssue(ByRef issues As Collection, ByVal issueText As String)
    issues.Add issueText
End Sub

Private Sub StatSet(ByVal stats As Object, ByVal statName As String, ByVal value As Long)
    If stats.Exists(statName) Then
        stats(statName) = value
    Else
        stats.Add statName, value
    End If
End Sub

Private Sub StatInc(ByVal stats As Object, ByVal statName As String, Optional ByVal amount As Long = 1)
    If stats.Exists(statName) Then
        stats(statName) = CLng(stats(statName)) + amount
    Else
        stats.Add statName, amount
    End If
End Sub

Private Function StatValue(ByVal stats As Object, ByVal statName As String) As Long
    If stats.Exists(statName) Then StatValue = CLng(stats(statName))
End Function

Private Sub GetNameAutoCorrectState( _
    ByRef trackInfo As Variant, _
    ByRef performCorrect As Variant, _
    ByRef logChanges As Variant, _
    ByRef known As Boolean)

    On Error GoTo UnknownState
    trackInfo = CBool(Application.GetOption("Track Name AutoCorrect Info"))
    performCorrect = CBool(Application.GetOption("Perform Name AutoCorrect"))
    logChanges = CBool(Application.GetOption("Log Name AutoCorrect Changes"))
    known = True
    Exit Sub

UnknownState:
    known = False
    trackInfo = Empty
    performCorrect = Empty
    logChanges = Empty
    Err.Clear
End Sub

Private Sub LogPreflightDetails( _
    ByVal stats As Object, _
    ByVal hardIssues As Collection, _
    ByVal nacTrack As Variant, _
    ByVal nacPerform As Variant, _
    ByVal nacLog As Variant, _
    ByVal nacKnown As Boolean)
    logLine "PRE-FLIGHT SUMMARY"
    logLine "Workbook rows: " & Format$(StatValue(stats, "WorkbookRows"), "#,##0")
    logLine "Supported top-level mappings: " & Format$(StatValue(stats, "SupportedMappings"), "#,##0")
    logLine "Supported field mappings: " & Format$(StatValue(stats, "FieldMappings"), "#,##0")
    logLine "  Table fields: " & Format$(StatValue(stats, "TableFieldMappings"), "#,##0")
    logLine "  Query output fields: " & Format$(StatValue(stats, "QueryFieldMappings"), "#,##0")
    logLine "Decided top-level included: " & Format$(StatValue(stats, "DecidedRowsIncluded"), "#,##0")
    logLine "Review top-level included: " & Format$(StatValue(stats, "ReviewRowsIncluded"), "#,##0")
    logLine "Top-level Delete skipped: " & Format$(StatValue(stats, "DeleteRowsSkipped"), "#,##0")
    logLine "Field Delete skipped: " & Format$(StatValue(stats, "FieldDeleteRowsSkipped"), "#,##0")
    logLine "Blank proposed rename conflicts: " & Format$(StatValue(stats, "BlankProposedConflicts"), "#,##0")
    logLine "Fields under Delete parent skipped: " & Format$(StatValue(stats, "FieldParentDeleteRowsSkipped"), "#,##0")
    logLine "Object pending: " & Format$(StatValue(stats, "Pending"), "#,##0")
    logLine "Object already renamed: " & Format$(StatValue(stats, "Already"), "#,##0")
    logLine "Object unchanged: " & Format$(StatValue(stats, "Unchanged"), "#,##0")
    logLine "Object missing: " & Format$(StatValue(stats, "Missing"), "#,##0")
    logLine "Field pending: " & Format$(StatValue(stats, "FieldPending"), "#,##0")
    logLine "  Calculated table fields pending: " & Format$(StatValue(stats, "CalculatedTableFieldPending"), "#,##0")
    logLine "Field already renamed: " & Format$(StatValue(stats, "FieldAlready"), "#,##0")
    logLine "Field unchanged: " & Format$(StatValue(stats, "FieldUnchanged"), "#,##0")
    logLine "Field missing: " & Format$(StatValue(stats, "FieldMissing"), "#,##0")
    logLine "Conflicts: " & Format$(hardIssues.count, "#,##0")
    If nacKnown Then
        logLine "Name AutoCorrect - Track: " & OnOffText(CBool(nacTrack))
        logLine "Name AutoCorrect - Perform: " & OnOffText(CBool(nacPerform))
        logLine "Name AutoCorrect - Log changes: " & OnOffText(CBool(nacLog))
    Else
        logLine "Name AutoCorrect state: UNKNOWN"
    End If
    ForEachIssueToLog hardIssues, "PRE-FLIGHT ISSUE: "
End Sub
Private Sub LogMissingItems(ByVal plan As Collection, ByVal fieldPlan As Collection)
    Dim item As Object
    Dim i As Long
    Dim actionText As String
    Dim parentName As String
    For i = 1 To plan.count
        Set item = plan(i)
        If StrComp(CStr(item("Action")), "Missing", vbTextCompare) = 0 Then
            logLine "MISSING/NOT BUILT OBJECT: " & CStr(item("Type")) & " | Current='" & CStr(item("Old")) & _
                "' | Proposed='" & CStr(item("New")) & "'"
        End If
    Next i
    For i = 1 To fieldPlan.count
        Set item = fieldPlan(i)
        actionText = CStr(item("Action"))
        If StrComp(actionText, "Missing", vbTextCompare) = 0 Or _
           StrComp(actionText, "MissingParent", vbTextCompare) = 0 Then
            parentName = CStr(item("ParentActual"))
            If Len(parentName) = 0 Then parentName = CStr(item("ParentOld"))
            logLine "MISSING/NOT BUILT FIELD: " & CStr(item("ParentType")) & " | Object='" & parentName & _
                "' | Field='" & CStr(item("Old")) & "' | Proposed='" & CStr(item("New")) & "'" & _
                IIf(StrComp(actionText, "MissingParent", vbTextCompare) = 0, " | Parent object missing/not built", vbNullString)
        End If
    Next i
End Sub
Private Function BuildConfirmationSummary( _
    ByVal stats As Object, _
    ByVal nacTrack As Variant, _
    ByVal nacPerform As Variant, _
    ByVal nacLog As Variant, _
    ByVal nacKnown As Boolean, _
    ByVal previewOnly As Boolean) As String
    Dim s As String
    s = "PRE-FLIGHT PASS" & vbCrLf & vbCrLf & _
        "Supported top-level mappings: " & Format$(StatValue(stats, "SupportedMappings"), "#,##0") & vbCrLf & _
        "  Decided included: " & Format$(StatValue(stats, "DecidedRowsIncluded"), "#,##0") & vbCrLf & _
        "  Review included: " & Format$(StatValue(stats, "ReviewRowsIncluded"), "#,##0") & vbCrLf & _
        "Supported field mappings: " & Format$(StatValue(stats, "FieldMappings"), "#,##0") & vbCrLf & _
        "  Table fields: " & Format$(StatValue(stats, "TableFieldMappings"), "#,##0") & vbCrLf & _
        "  Query output fields: " & Format$(StatValue(stats, "QueryFieldMappings"), "#,##0") & vbCrLf & _
        "Top-level Delete rows skipped: " & Format$(StatValue(stats, "DeleteRowsSkipped"), "#,##0") & vbCrLf & _
        "Field Delete rows skipped: " & Format$(StatValue(stats, "FieldDeleteRowsSkipped"), "#,##0") & vbCrLf & _
        "Blank proposed rename conflicts: " & Format$(StatValue(stats, "BlankProposedConflicts"), "#,##0") & vbCrLf & _
        "Fields under Delete parent skipped: " & Format$(StatValue(stats, "FieldParentDeleteRowsSkipped"), "#,##0") & vbCrLf & vbCrLf & _
        "Object renames pending: " & Format$(StatValue(stats, "Pending"), "#,##0") & vbCrLf & _
        "Field renames pending: " & Format$(StatValue(stats, "FieldPending"), "#,##0") & vbCrLf & _
        "  Calculated table fields: " & Format$(StatValue(stats, "CalculatedTableFieldPending"), "#,##0") & vbCrLf & _
        "Already renamed (objects / fields): " & Format$(StatValue(stats, "Already"), "#,##0") & " / " & _
            Format$(StatValue(stats, "FieldAlready"), "#,##0") & vbCrLf & _
        "Unchanged (objects / fields): " & Format$(StatValue(stats, "Unchanged"), "#,##0") & " / " & _
            Format$(StatValue(stats, "FieldUnchanged"), "#,##0") & vbCrLf & _
        "Missing/not built (objects / fields): " & Format$(StatValue(stats, "Missing"), "#,##0") & " / " & _
            Format$(StatValue(stats, "FieldMissing"), "#,##0") & vbCrLf & _
        "Hard conflicts: 0" & vbCrLf & vbCrLf
    If nacKnown Then
        s = s & "Name AutoCorrect:" & vbCrLf & _
            "  Track info: " & OnOffText(CBool(nacTrack)) & vbCrLf & _
            "  Perform: " & OnOffText(CBool(nacPerform)) & vbCrLf & _
            "  Log changes: " & OnOffText(CBool(nacLog)) & vbCrLf & vbCrLf
    Else
        s = s & "Name AutoCorrect state could not be read." & vbCrLf & vbCrLf
    End If
    s = s & "Execution order:" & vbCrLf & _
        "  1. Query output field aliases" & vbCrLf & _
        "  2. Table fields" & vbCrLf & _
        "  3. Top-level objects" & vbCrLf & vbCrLf & _
        "Status Delete is skip-only; this utility does not delete mapped objects or fields." & vbCrLf & _
        "Calculated field renames use an internal placeholder migration and require Name AutoCorrect Track + Perform ON." & vbCrLf & _
        "The utility does not change Name AutoCorrect settings." & vbCrLf & vbCrLf
    If previewOnly Then
        s = s & "PREVIEW ONLY - no database objects or fields were renamed." & vbCrLf & _
            "Missing/not-built objects and fields are listed in the run log."
    End If
    BuildConfirmationSummary = s
End Function
Private Function BuildPreflightFailureMessage(ByVal stats As Object, ByVal issues As Collection) As String
    BuildPreflightFailureMessage = "PRE-FLIGHT FAIL - no database objects or fields were renamed." & vbCrLf & vbCrLf & _
        "Hard issues: " & Format$(issues.count, "#,##0") & vbCrLf & vbCrLf & _
        FirstIssuesText(issues, 10) & vbCrLf & vbCrLf & _
        "Full log: " & mLogPath
End Function
Private Function BuildCompletionText(ByVal stats As Object) As String
    BuildCompletionText = _
        "Objects renamed: " & Format$(StatValue(stats, "Renamed"), "#,##0") & _
        " | Fields renamed: " & Format$(StatValue(stats, "FieldRenamed"), "#,##0") & _
        " | Missing objects: " & Format$(StatValue(stats, "Missing"), "#,##0") & _
        " | Missing fields: " & Format$(StatValue(stats, "FieldMissing"), "#,##0")
End Function
Private Function BuildCompletionMessage(ByVal stats As Object) As String
    BuildCompletionMessage = "FINAL VERIFICATION: PASS" & vbCrLf & vbCrLf & _
        "Objects renamed: " & Format$(StatValue(stats, "Renamed"), "#,##0") & vbCrLf & _
        "Table fields renamed: " & Format$(StatValue(stats, "TableFieldRenamed"), "#,##0") & vbCrLf & _
        "Query output fields renamed: " & Format$(StatValue(stats, "QueryFieldRenamed"), "#,##0") & vbCrLf & _
        "Already renamed (objects / fields): " & Format$(StatValue(stats, "Already"), "#,##0") & " / " & _
            Format$(StatValue(stats, "FieldAlready"), "#,##0") & vbCrLf & _
        "Unchanged (objects / fields): " & Format$(StatValue(stats, "Unchanged"), "#,##0") & " / " & _
            Format$(StatValue(stats, "FieldUnchanged"), "#,##0") & vbCrLf & _
        "Missing/not built (objects / fields): " & Format$(StatValue(stats, "Missing"), "#,##0") & " / " & _
            Format$(StatValue(stats, "FieldMissing"), "#,##0") & vbCrLf & _
        "Temporary object cycle renames: " & Format$(StatValue(stats, "TemporaryRenames"), "#,##0") & vbCrLf & _
        "Temporary field cycle renames: " & Format$(StatValue(stats, "FieldTemporaryRenames"), "#,##0") & vbCrLf & vbCrLf & _
        "Full run log:" & vbCrLf & mLogPath
End Function
Private Function FirstIssuesText(ByVal issues As Collection, ByVal maxCount As Long) As String
    Dim i As Long
    Dim n As Long
    Dim s As String

    n = issues.count
    If n > maxCount Then n = maxCount

    For i = 1 To n
        If Len(s) > 0 Then s = s & vbCrLf
        s = s & CStr(i) & ". " & CStr(issues(i))
    Next i

    If issues.count > maxCount Then
        s = s & vbCrLf & "... plus " & CStr(issues.count - maxCount) & " more."
    End If

    FirstIssuesText = s
End Function

Private Sub ForEachIssueToLog(ByVal issues As Collection, ByVal prefix As String)
    Dim i As Long
    For i = 1 To issues.count
        logLine prefix & CStr(issues(i))
    Next i
End Sub

Private Function OnOffText(ByVal value As Boolean) As String
    If value Then
        OnOffText = "ON"
    Else
        OnOffText = "OFF"
    End If
End Function

Private Sub StartRunLog(ByVal workbookPath As String, ByVal previewOnly As Boolean)
    CloseRunLog
    On Error GoTo UseTempFolder

    mLogPath = CurrentProject.Path & "\" & LOG_FILE_NAME
    mLogFileNo = FreeFile
    Open mLogPath For Output As #mLogFileNo
    mLogOpen = True
    GoTo LogHeader

UseTempFolder:
    Err.Clear
    mLogPath = Environ$("TEMP") & "\" & LOG_FILE_NAME
    mLogFileNo = FreeFile
    Open mLogPath For Output As #mLogFileNo
    mLogOpen = True

LogHeader:
    logLine String$(78, "=")
    If previewOnly Then
        logLine "ADC RENAME PREVIEW"
    Else
        logLine "ADC RENAME LIVE"
    End If
    logLine "Database: " & CurrentProject.FullName
    logLine "Rename workbook: " & workbookPath
    logLine String$(78, "=")
End Sub

Private Sub CloseRunLog()
    On Error Resume Next
    If mLogOpen Then Close #mLogFileNo
    mLogOpen = False
End Sub

Private Sub logLine(ByVal text As String)
    If Not mLogOpen Then Exit Sub
    Print #mLogFileNo, Format$(Now, "hh:nn:ss") & " | " & text
End Sub

Private Sub ProgressAppend(ByVal text As String)
    logLine text
    mProgressText = Format$(Now, "hh:nn:ss") & " | " & text & vbCrLf & mProgressText
    If Len(mProgressText) > MAX_VISIBLE_LOG_CHARS Then
        mProgressText = Left$(mProgressText, MAX_VISIBLE_LOG_CHARS)
    End If
End Sub

Private Sub RebuildProgressForm(ByVal previewOnly As Boolean, ByVal runTitle As String)
    Dim frm As Form
    Dim ctl As control
    Dim tempName As String

    DeleteProgressFormIfExists PROGRESS_FORM
    On Error GoTo BuildFail

    Set frm = CreateForm
    tempName = frm.Name

    With frm
        .Caption = runTitle
        .PopUp = True
        .Modal = False
        .AutoCenter = True
        .NavigationButtons = False
        .RecordSelectors = False
        .DividingLines = False
        .ScrollBars = 0
        .BorderStyle = 3
        .MinMaxButtons = 0
        .Width = 11200
        .Section(acDetail).Height = 6650
    End With

    Set ctl = CreateControl(tempName, acLabel, acDetail, vbNullString, vbNullString, 300, 180, 10400, 360)
    ctl.Name = "lblPhase"
    ctl.Caption = "Starting..."
    ctl.fontSize = 14
    ctl.fontBold = True

    Set ctl = CreateControl(tempName, acLabel, acDetail, vbNullString, vbNullString, 300, 650, 10400, 300)
    ctl.Name = "lblCount"
    ctl.Caption = vbNullString
    ctl.fontSize = 10

    Set ctl = CreateControl(tempName, acLabel, acDetail, vbNullString, vbNullString, 300, 1020, 9800, 300)
    ctl.Name = "lblBarBack"
    ctl.Caption = vbNullString
    ctl.BackStyle = 1
    ctl.BackColor = RGB(220, 220, 220)
    ctl.BorderStyle = 1

    Set ctl = CreateControl(tempName, acLabel, acDetail, vbNullString, vbNullString, 300, 1020, 10, 300)
    ctl.Name = "lblBar"
    ctl.Caption = vbNullString
    ctl.BackStyle = 1
    ctl.BackColor = RGB(70, 120, 190)
    ctl.BorderStyle = 0

    Set ctl = CreateControl(tempName, acLabel, acDetail, vbNullString, vbNullString, 10250, 1020, 700, 300)
    ctl.Name = "lblPct"
    ctl.Caption = "0%"
    ctl.TextAlign = 3

    Set ctl = CreateControl(tempName, acTextBox, acDetail, vbNullString, vbNullString, 300, 1450, 10400, 850)
    ctl.Name = "txtCurrent"
    ctl.controlSource = vbNullString
    ctl.Locked = True
    ctl.TabStop = False
    ctl.ScrollBars = 2
    ctl.FontName = "Courier New"
    ctl.fontSize = 9

    Set ctl = CreateControl(tempName, acTextBox, acDetail, vbNullString, vbNullString, 300, 2450, 10400, 3400)
    ctl.Name = "txtOutput"
    ctl.controlSource = vbNullString
    ctl.Locked = True
    ctl.TabStop = False
    ctl.ScrollBars = 2
    ctl.FontName = "Courier New"
    ctl.fontSize = 8

    Set ctl = CreateControl(tempName, acCommandButton, acDetail, vbNullString, vbNullString, 8850, 6020, 1850, 420)
    ctl.Name = "cmdCancel"
    If previewOnly Then
        ctl.Caption = "Cancel"
    Else
        ctl.Caption = "Cancel + Roll Back"
    End If
    ctl.OnClick = "=RenamePreview_RequestCancel()"

    DoCmd.Save acForm, tempName
    DoCmd.Close acForm, tempName, acSaveYes
    DoCmd.Rename PROGRESS_FORM, acForm, tempName
    Exit Sub

BuildFail:
    Err.Raise Err.Number, "RebuildProgressForm", Err.Description
End Sub

Private Sub OpenProgressForm()
    DoCmd.OpenForm PROGRESS_FORM, acNormal
    mProgressOpen = True
    ProgressUpdate "Starting", 0, vbNullString, True
End Sub

Private Sub CloseProgressForm()
    On Error Resume Next
    If mProgressOpen Then DoCmd.Close acForm, PROGRESS_FORM, acSaveNo
    mProgressOpen = False
End Sub

Private Sub DeleteProgressForm()
    DeleteProgressFormIfExists PROGRESS_FORM
End Sub

Private Sub DeleteProgressFormIfExists(ByVal formName As String)
    Dim ao As accessObject

    On Error GoTo CleanFail

    For Each ao In CurrentProject.AllForms
        If StrComp(ao.Name, formName, vbTextCompare) = 0 Then
            If ao.IsLoaded Then DoCmd.Close acForm, formName, acSaveNo
            DoCmd.DeleteObject acForm, formName
            Exit For
        End If
    Next ao

CleanExit:
    Exit Sub

CleanFail:
    ' Cleanup must never mask the result of the rename operation.
    Err.Clear
    Resume CleanExit
End Sub

Private Sub ProgressUpdate( _
    ByVal phaseText As String, _
    ByVal percentComplete As Double, _
    ByVal currentText As String, _
    Optional ByVal forceRefresh As Boolean = False)

    Dim frm As Form
    Dim barWidth As Long

    If Not mProgressOpen Then Exit Sub
    If percentComplete < 0 Then percentComplete = 0
    If percentComplete > 100 Then percentComplete = 100

    If Not forceRefresh Then
        If SecondsSince(mLastUiTimer) < UI_REFRESH_SECONDS Then Exit Sub
    End If
    mLastUiTimer = Timer

    On Error GoTo UiFail
    Set frm = forms(PROGRESS_FORM)
    frm.Controls("lblPhase").Caption = phaseText
    frm.Controls("lblPct").Caption = Format$(percentComplete, "0") & "%"
    frm.Controls("lblCount").Caption = BuildLiveCountCaption()
    frm.Controls("txtCurrent").value = currentText
    frm.Controls("txtOutput").value = mProgressText

    barWidth = CLng(9800# * percentComplete / 100#)
    If barWidth < 10 Then barWidth = 10
    frm.Controls("lblBar").Width = barWidth
    frm.Repaint
    DoEvents

    If mCancelRequested Then
        frm.Controls("cmdCancel").Caption = "Cancelling..."
        frm.Controls("cmdCancel").enabled = False
    End If
    Exit Sub

UiFail:
    logLine "Progress UI warning " & CStr(Err.Number) & ": " & Err.Description
    Err.Clear
End Sub

Private Function BuildLiveCountCaption() As String
    BuildLiveCountCaption = "Run log: " & mLogPath
End Function

Private Function SecondsSince(ByVal priorTimer As Double) As Double
    Dim nowTimer As Double
    nowTimer = Timer
    If priorTimer <= 0 Then
        SecondsSince = 86400#
    ElseIf nowTimer >= priorTimer Then
        SecondsSince = nowTimer - priorTimer
    Else
        SecondsSince = (86400# - priorTimer) + nowTimer
    End If
End Function
