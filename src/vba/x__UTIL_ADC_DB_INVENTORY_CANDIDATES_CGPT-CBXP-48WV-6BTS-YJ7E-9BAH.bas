Attribute VB_Name = "x__UTIL_ADC_DB_INVENTORY_CANDIDATES"
' ================================================================================================================================================================================
' ChatGPT Rewrite Tracking Header
' ================================================================================================================================================================================
' ChatGPT Previous Version ID: CGPT-D722-7LVG-UMW2-KJJG-ALPN
' ChatGPT Baseline Version ID: CGPT-D722-7LVG-UMW2-KJJG-ALPN
' ChatGPT Version ID: CGPT-CBXP-48WV-6BTS-YJ7E-9BAH
' ChatGPT Version Timestamp: Friday 8/28/26 11:52 AM
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
' STANDARD MODULE: Access Development Console DB Inventory Candidates
' VB_Name: x__UTIL_ADC_DB_INVENTORY_CANDIDATES
' Entry Points:
'     ScanCandidatesForRouting
'     AcceptCandidate
'     AcceptSelectedRunnerMacroCandidates
'     ReviewStaleCandidates
'     MarkStaleCandidates
' Purpose:
'     Owns DB Inventory candidate discovery, scan-state maintenance, acceptance, stale/review evaluation, and candidate-marking recovery. It consumes shared canonical/default/derivation authority from
'     x__UTIL_ADC_DEFAULTS and does not own physical MC_Records or MC_Candidates schema creation.
' Author: Christina Shonts
' Support:
'     Access Development Console DB Inventory candidate discovery, acceptance, review, stale marking, and recovery workflows.
' Code Category: Access Development Console DB Inventory candidate operational authority.
' Requirements:
'     Microsoft Access, DAO, trusted access to the VBA project object model for procedure discovery, and the companion ADC modules.
' Reference Dependencies:
'     Access object model, DAO, VBE project model, x__UTIL_ADC_DEFAULTS shared derivation/string helpers, x__UTIL_ADC_MACRO_RUNTIME configured-row readiness/lookups, and companion ADC modules.
' DB Object Dependencies:
'     Uses the current MC_Candidates candidate storage created/repaired by x__UTIL_ADC_TABLES; this module does not own physical schema creation.
' Usage:
'     Core and generated UI routes candidate scan, acceptance, stale-review, stale-marking, and recovery commands here.
' Version:
'     Preserves actual Access object names exactly during candidate skip and hidden/system classification.
'     Stores discovery/review state in MC_Candidates and changes acceptance from in-place row promotion to transactional creation of a linked MC_Records configuration while preserving Setup-phase ownership.
'     Routes Managed Object Catalog references to x__UTIL_ADC_MOC and physical lifecycle references to x__UTIL_ADC_MOC_LIFECYCLE after the responsibility split; runtime and Setup-phase behavior are unchanged.
'     Created by extracting the candidate subsystem from x__UTIL_ADC_CONFIG without changing Setup-phase ownership or candidate workflow sequencing.
' Notes:
'     DB_DEFINITIONS remains authoritative for physical schema definitions; TABLES remains the A2 physical create/repair/migration authority.
'     Editable defaults and shared derivation/string manipulation remain in DEFAULTS.
' Core Rules Enforced Here:
'     Candidate discovery and lifecycle behavior is centralized here; shared derivation/string helpers are not duplicated; configured-macro generation/execution remains in x__UTIL_ADC_MACRO_RUNTIME.
' Design Map:
'     candidate discovery -> candidate row maintenance -> candidate acceptance -> stale/review evaluation -> stale marking/recovery.
' Configuration:
'     Current DEFAULTS, DB_DEFINITIONS, TABLES, MOC, and MACRO_RUNTIME authorities remain authoritative for their respective defaults, schema definitions, physical repair, managed-object identity,
'     and configured-target behavior.
' Behavior:
'     Preserves candidate scan progress/heartbeat behavior, transaction and rollback semantics, row-level error handling, and command error propagation.
'     Acceptance now inserts a separate configuration and links the candidate to it.
' Code Status: Active canonical DB Inventory candidate authority.
' Todo: None.
' Reporting Bugs: None currently recorded in this header.
' Module Level Bugs: None currently recorded in this header.
' Public Subs:
'     RetryConfigCandidateMarkingRecovery(Optional ByVal showMessage As Boolean = True) -> None
'     ScanCandidatesForRouting(ByVal scanCodeDb As Boolean, ByVal scanHostDb As Boolean, ByVal resultConfigDb As DAO.Database, ByVal resultConfigDatabaseRole As String, ByVal progressCaption As String, Optional
'     ByVal showMessage As Boolean = True) -> None
'     AcceptCandidate(ByVal CandidateID As Long, Optional ByVal showMessage As Boolean = True) -> None
'     AcceptSelectedRunnerMacroCandidates(Optional ByVal showMessage As Boolean = True) -> None
'     MarkStaleCandidates(Optional ByVal showMessage As Boolean = True, Optional ByVal confirmStatusChange As Boolean = True) -> None
' Public Functions:
'     ConfigCandidateMarkingRecoveryBlocked(Optional ByVal durableRecoverySnapshot As Variant) -> Boolean
'     ReviewStaleCandidates(ByVal showMessage As Boolean, ByVal writeHealthFindings As Boolean) -> Long
'     CandidateStaleReason(ByRef cfg As TCandidateRecord, ByVal latestBatchID As String) -> String
'     CandidateReviewReason(ByRef cfg As TCandidateRecord, ByVal latestBatchID As String) -> String
'     LatestCandidateScanBatchID(ByVal CandidateOwnerDatabaseRole As String) -> String
' Public Properties: None
' ================================================================================================================================================================================

Option Compare Database
Option Explicit

' Candidate discovery limits and row-level error contracts.
Private Const CANDIDATE_SCAN_ROW_ERRORS_ERROR_NUMBER As Long = vbObjectError + 5458
Private Const CANDIDATE_ACCEPT_ROW_ERRORS_ERROR_NUMBER As Long = vbObjectError + 9197

' Candidate-marking transaction recovery authority.
Private Const ERR_CONFIG_CANDIDATE_MARKING_RECOVERY_BLOCKED As Long = vbObjectError + 5952
Private Const ERR_CANDIDATE_MARKING_ROLLBACK_FAILED As Long = vbObjectError + 5953
' Preserve the durable key value so unresolved markers written by earlier builds remain recoverable.
Private Const CONFIG_CANDIDATE_MARKING_RECOVERY_KEY As String = "ADC.Review.CandidateMarkingRecoveryV1"
Private mConfigCandidateMarkingRecoveryBlocked As Boolean
Private mConfigCandidateMarkingRecoveryOperationName As String
Private mConfigCandidateMarkingRecoveryDetail As String
Private mConfigCandidateMarkingRecoveryWorkspace As DAO.workspace
Private mConfigCandidateMarkingRecoveryDatabase As DAO.Database

' ===============================================================================================================================================================================
' DB INVENTORY CANDIDATE PROCEDURES
' ===============================================================================================================================================================================

' ===============================================================================================================================================================================
' CANDIDATE RECOVERY PROCEDURES
' ===============================================================================================================================================================================

' Purpose: Returns whether candidate-marking transaction recovery is unresolved, including a durable blocker from a prior VBA reset.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   durableRecoverySnapshot (Input Optional; Variant): Snapshot of durable recovery state captured before or during the operation and used for validation, rollback, or recovery decisions.
' Returns: Boolean: True when config candidate marking recovery is blocked by the condition tracked by this routine; False when work may proceed.
Public Function ConfigCandidateMarkingRecoveryBlocked(Optional ByVal durableRecoverySnapshot As Variant) As Boolean
    'Returns whether candidate-marking transaction recovery is unresolved, including a durable blocker from a prior VBA reset.
    On Error GoTo InspectionFailed
    Dim durableOperation As String
    Dim durableDetail As String
    Dim durablePayload As String
    Dim recoverySnapshot As Object
    Dim markerExists As Boolean
    If mConfigCandidateMarkingRecoveryBlocked Then ConfigCandidateMarkingRecoveryBlocked = True: Exit Function
    If isMissing(durableRecoverySnapshot) Then
        markerExists = CoreDurableRecoveryMarkerExists(CONFIG_CANDIDATE_MARKING_RECOVERY_KEY, durableOperation, durableDetail, durablePayload)
    Else
        If IsObject(durableRecoverySnapshot) = False Then Err.Raise vbObjectError + 6148, "ConfigCandidateMarkingRecoveryBlocked", "The supplied durable recovery snapshot is not an object."
        Set recoverySnapshot = durableRecoverySnapshot
        markerExists = CoreDurableRecoveryMarkerExistsInSnapshot(recoverySnapshot, CONFIG_CANDIDATE_MARKING_RECOVERY_KEY, durableOperation, durableDetail, durablePayload)
    End If
    If markerExists Then
        mConfigCandidateMarkingRecoveryBlocked = True
        ConfigCandidateMarkingRecoveryBlocked = True
    End If
    Exit Function
InspectionFailed:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If savedErrNumber = 0 Then savedErrNumber = vbObjectError + 6147
    If Len(savedErrSource) = 0 Then savedErrSource = "ConfigCandidateMarkingRecoveryBlocked"
    Err.Raise savedErrNumber, "ConfigCandidateMarkingRecoveryBlocked", "Could not inspect the durable Config/Candidate candidate-marking recovery marker. No unresolved recovery state was inferred or latched from this inspection failure." & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Function

Private Sub RequireNoUnfinishedConfigCandidateMarkingRecovery(ByVal requestedOperationName As String)
    'Blocks candidate marking and all coordinated top-level work while its transaction boundary is unresolved.
    requestedOperationName = Trim$(requestedOperationName)
    If Len(requestedOperationName) = 0 Then requestedOperationName = "Access Development Console Config/Candidate operation"
    If ConfigCandidateMarkingRecoveryBlocked() = False Then Exit Sub
    Err.Raise ERR_CONFIG_CANDIDATE_MARKING_RECOVERY_BLOCKED, "RequireNoUnfinishedConfigCandidateMarkingRecovery", "A Access Development Console operation cannot start because stale-candidate marking transaction recovery is unresolved." & vbCrLf & "Requested operation: " & requestedOperationName & vbCrLf & "Recovery operation: " & IIf(Len(mConfigCandidateMarkingRecoveryOperationName) = 0, "(not recorded)", mConfigCandidateMarkingRecoveryOperationName) & vbCrLf & "Recovery detail:" & vbCrLf & IIf(Len(mConfigCandidateMarkingRecoveryDetail) = 0, "(not recorded)", mConfigCandidateMarkingRecoveryDetail) & vbCrLf & "Run RetryConfigCandidateMarkingRecovery before continuing."
End Sub

Private Sub RegisterConfigCandidateMarkingRecoveryFailure(ByVal operationName As String, ByVal detailText As String, ByVal recoveryWs As DAO.workspace, ByVal recoveryDb As DAO.Database)
    On Error GoTo PersistErr
    mConfigCandidateMarkingRecoveryBlocked = True
    mConfigCandidateMarkingRecoveryOperationName = Trim$(operationName)
    mConfigCandidateMarkingRecoveryDetail = Trim$(detailText)
    Set mConfigCandidateMarkingRecoveryWorkspace = recoveryWs
    Set mConfigCandidateMarkingRecoveryDatabase = recoveryDb
    PersistDurableRecoveryMarker CONFIG_CANDIDATE_MARKING_RECOVERY_KEY, mConfigCandidateMarkingRecoveryOperationName, mConfigCandidateMarkingRecoveryDetail
    Exit Sub
PersistErr:
    mConfigCandidateMarkingRecoveryDetail = mConfigCandidateMarkingRecoveryDetail & vbCrLf & "Durable recovery persistence failed. Error " & CStr(Err.Number) & ": " & Err.Description
    Err.Raise ERR_CANDIDATE_MARKING_ROLLBACK_FAILED, "RegisterConfigCandidateMarkingRecoveryFailure", mConfigCandidateMarkingRecoveryDetail
End Sub

' Purpose: Performs the public retry config candidate marking recovery operation using the routine's existing ADC workflow and side effects.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine shows its user-facing completion, status, warning, or error message in addition to performing the underlying work. Default: True.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub RetryConfigCandidateMarkingRecovery(Optional ByVal showMessage As Boolean = True)
    'Retries the retained dedicated candidate-marking workspace or verifies that a restarted session released it, then clears the durable blocker separately.
    On Error GoTo ErrHandler
    Dim rollbackErrNumber As Long
    Dim rollbackErrDescription As String
    Dim sourceDb As DAO.Database
    Dim verifyWs As DAO.workspace
    Dim verifyDb As DAO.Database
    Dim verifyRs As DAO.Recordset
    Dim resolutionText As String
    If ConfigCandidateMarkingRecoveryBlocked() = False Then
        If showMessage Then CoreShowMessage "No unfinished Access Development Console candidate-marking recovery exists.", vbInformation, "Access Development Console Candidate Marking Recovery", "RetryConfigCandidateMarkingRecovery", "Config"
        Exit Sub
    End If
    If Not mConfigCandidateMarkingRecoveryWorkspace Is Nothing Then
        On Error Resume Next
        Err.Clear
        mConfigCandidateMarkingRecoveryWorkspace.Rollback
        rollbackErrNumber = Err.Number
        rollbackErrDescription = Err.Description
        On Error GoTo ErrHandler
        If rollbackErrNumber <> 0 And rollbackErrNumber <> 3034 Then Err.Raise ERR_CANDIDATE_MARKING_ROLLBACK_FAILED, "RetryConfigCandidateMarkingRecovery", "The retained candidate-marking workspace still could not be rolled back. Error " & CStr(rollbackErrNumber) & ": " & rollbackErrDescription
        If rollbackErrNumber = 3034 Then resolutionText = "The retained candidate-marking workspace had no live transaction remaining." Else resolutionText = "The retained candidate-marking workspace transaction was rolled back successfully."
    Else
        resolutionText = "The original candidate-marking workspace is no longer retained. No default-workspace rollback was attempted."
    End If
    On Error Resume Next
    If Not mConfigCandidateMarkingRecoveryDatabase Is Nothing Then mConfigCandidateMarkingRecoveryDatabase.Close
    If Not mConfigCandidateMarkingRecoveryWorkspace Is Nothing Then mConfigCandidateMarkingRecoveryWorkspace.Close
    Set mConfigCandidateMarkingRecoveryDatabase = Nothing
    Set mConfigCandidateMarkingRecoveryWorkspace = Nothing
    On Error GoTo ErrHandler
    Set sourceDb = ADCDb()
    Set verifyWs = DBEngine.CreateWorkspace("ADCConfigCandidateMarkingRecoveryVerify_" & Format$(Now, "hhnnss") & "_" & Format$(CLng(Timer * 1000), "000000000"), "Admin", vbNullString, dbUseJet)
    Set verifyDb = verifyWs.OpenDatabase(sourceDb.Name)
    Set verifyRs = verifyDb.OpenRecordset("SELECT TOP 1 " & SQL_FIELD_ID & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "];", dbOpenSnapshot)
    verifyRs.Close
    Set verifyRs = Nothing
    verifyDb.Close
    Set verifyDb = Nothing
    verifyWs.Close
    Set verifyWs = Nothing
    resolutionText = resolutionText & vbCrLf & "CodeDb and the MC_Records table were reopened and read successfully through a new independent workspace."
    mConfigCandidateMarkingRecoveryOperationName = "Candidate marking recovery marker cleanup"
    mConfigCandidateMarkingRecoveryDetail = resolutionText & vbCrLf & "Transaction resolution is complete; durable blocker cleanup is pending."
    PersistDurableRecoveryMarker CONFIG_CANDIDATE_MARKING_RECOVERY_KEY, mConfigCandidateMarkingRecoveryOperationName, mConfigCandidateMarkingRecoveryDetail
    DeleteDurableRecoveryMarker CONFIG_CANDIDATE_MARKING_RECOVERY_KEY
    mConfigCandidateMarkingRecoveryBlocked = False
    mConfigCandidateMarkingRecoveryOperationName = vbNullString
    mConfigCandidateMarkingRecoveryDetail = vbNullString
    Set mConfigCandidateMarkingRecoveryDatabase = Nothing
    Set mConfigCandidateMarkingRecoveryWorkspace = Nothing
    If showMessage Then CoreShowMessage "Access Development Console candidate-marking transaction recovery completed and was verified." & vbCrLf & vbCrLf & resolutionText, vbInformation, "Access Development Console Candidate Marking Recovery", "RetryConfigCandidateMarkingRecovery", "Config"
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not verifyRs Is Nothing Then verifyRs.Close
    If Not verifyDb Is Nothing Then verifyDb.Close
    If Not verifyWs Is Nothing Then verifyWs.Close
    PersistDurableRecoveryMarker CONFIG_CANDIDATE_MARKING_RECOVERY_KEY, mConfigCandidateMarkingRecoveryOperationName, savedErrDescription
    On Error GoTo 0
    mConfigCandidateMarkingRecoveryBlocked = True
    If showMessage Then
        CoreShowErrorMessage "Access Development Console candidate-marking recovery remains unresolved." & vbCrLf & vbCrLf & "Error " & CStr(savedErrNumber) & ": " & savedErrDescription, vbCritical, "Access Development Console Candidate Marking Recovery", "RetryConfigCandidateMarkingRecovery", "Core"
        Exit Sub
    End If
    Err.Raise ERR_CANDIDATE_MARKING_ROLLBACK_FAILED, "RetryConfigCandidateMarkingRecovery", savedErrDescription
End Sub

' ===============================================================================================================================================================================
' CANDIDATE SCAN / DISCOVERY PROCEDURES
' ===============================================================================================================================================================================

' Purpose: Scans candidates for routing for scan code DB, scan host DB.
' Public role: Public module entry point. This procedure is listed in the module header as an intentional externally callable ADC entry point and may be invoked directly by Access/VBA callers.
' Parameters:
'   scanCodeDb (Input; Boolean): When True, enables or requests code DB for this call; when False, the routine follows its normal non-forced/default path.
'   scanHostDb (Input; Boolean): When True, enables or requests host DB for this call; when False, the routine follows its normal non-forced/default path.
'   resultConfigDb (Input; DAO.Database): DAO.Database reference for the database context that this operation reads or updates.
'   resultConfigDatabaseRole (Input; String): Role value selecting which result config database ownership or database context applies to this operation.
'   progressCaption (Input; String): Progress caption text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub ScanCandidatesForRouting(ByVal scanCodeDb As Boolean, ByVal scanHostDb As Boolean, ByVal resultConfigDb As DAO.Database, ByVal resultConfigDatabaseRole As String, ByVal progressCaption As String, Optional ByVal showMessage As Boolean = True)
    On Error GoTo ErrHandler
    Dim progressStarted As Boolean
    resultConfigDatabaseRole = NormalizedConfigOwnerDatabaseRole(resultConfigDatabaseRole)
    If resultConfigDb Is Nothing Then Err.Raise vbObjectError + 5450, "ScanCandidatesForRouting", "ResultConfigDb reference is Nothing."
    If Len(resultConfigDatabaseRole) = 0 Then Err.Raise vbObjectError + 5451, "ScanCandidatesForRouting", "ResultConfigDb role is not valid."
    If scanCodeDb = False And scanHostDb = False Then Err.Raise vbObjectError + 5452, "ScanCandidatesForRouting", "At least one ScanTargetDb role must be selected."
    progressStarted = CoreBeginOperationProgressIfNeeded(progressCaption, 2)
    If progressStarted Then CoreUpdateSetupProgress "Candidate scan: prepare " & resultConfigDatabaseRole & " MC_Candidates candidate table" Else CorePulseProgress "Candidate scan: prepare " & resultConfigDatabaseRole & " MC_Candidates candidate table"
    RepairResultConfigDbForScan resultConfigDatabaseRole
    If SetupControlCancellationRequested() Then RequireSafeWorkContinue "ScanCandidatesForRouting", "Candidate scan foundation prepared"
    If progressStarted Then CoreUpdateSetupProgress "Candidate scan: inspect selected database roles for VBA procedures and Access objects" Else CorePulseProgress "Candidate scan: inspect selected database roles for VBA procedures and Access objects"
    ScanCandidatesCore scanCodeDb, scanHostDb, resultConfigDb, resultConfigDatabaseRole, showMessage
    If progressStarted Then CoreCompleteSetupProgress
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If progressStarted Then CoreFailSetupProgress
    If savedErrNumber = CANDIDATE_SCAN_ROW_ERRORS_ERROR_NUMBER Then Err.Raise savedErrNumber, savedErrSource, savedErrDescription
    Err.Raise savedErrNumber, "ScanCandidatesForRouting", "Could not scan macro candidates." & vbCrLf & "ResultConfigDb role: " & resultConfigDatabaseRole & vbCrLf & "ResultConfigDb: " & SafeDatabaseNameForError(resultConfigDb) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub ScanCandidatesCore(ByVal scanCodeDb As Boolean, ByVal scanHostDb As Boolean, ByVal resultConfigDb As DAO.Database, ByVal resultConfigDatabaseRole As String, Optional ByVal showMessage As Boolean = True)
    On Error GoTo ErrHandler
    CoreResetOperationNotes
    Dim rs As DAO.Recordset
    Dim systemMacroCatalogRows As Collection
    Dim systemMacroWrapperNames As Object
    Dim scanCache As Object
    Dim foundCount As Long
    Dim procedureCandidateCount As Long
    Dim objectCandidateCount As Long
    Dim objectTableCount As Long
    Dim objectQueryCount As Long
    Dim objectFormCount As Long
    Dim objectReportCount As Long
    Dim objectMacroCount As Long
    Dim skippedObjectCount As Long
    Dim skippedModuleCount As Long
    Dim discoveryErrorCount As Long
    Dim skippedBuiltInMacroTargetCount As Long
    Dim procedureSourceSummary As String
    Dim objectSourceSummary As String
    Dim vbeAccessDetail As String
    Dim scanBatchID As String
    Dim discoveryMessageLines As Collection
    Dim discoveryMessageText As String
    Dim timingText As String
    Dim scanStartedTimer As Single
    Dim phaseStartedTimer As Single
    Dim prerequisiteElapsedSeconds As Double
    Dim stalePreparationElapsedSeconds As Double
    Dim cachePreloadElapsedSeconds As Double
    Dim procedureScanElapsedSeconds As Double
    Dim objectScanElapsedSeconds As Double
    Dim tableScanElapsedSeconds As Double
    Dim queryScanElapsedSeconds As Double
    Dim queryNameFilterElapsedSeconds As Double
    Dim queryHiddenSystemFilterElapsedSeconds As Double
    Dim queryCandidateProcessingElapsedSeconds As Double
    Dim formScanElapsedSeconds As Double
    Dim reportScanElapsedSeconds As Double
    Dim macroScanElapsedSeconds As Double
    Dim finalizationElapsedSeconds As Double
    Dim totalElapsedSeconds As Double
    scanStartedTimer = Timer
    phaseStartedTimer = Timer
    If resultConfigDb Is Nothing Then Err.Raise vbObjectError + 5401, "ScanCandidatesCore", "ResultConfigDb reference is Nothing."
    If Len(resultConfigDatabaseRole) = 0 Then Err.Raise vbObjectError + 5401, "ScanCandidatesCore", "ResultConfigDb role is not valid."
    If CanAccessVBEProjectModel(vbeAccessDetail) = False Then Err.Raise vbObjectError + 5402, "ScanCandidatesCore", "VBE ACCESS ERROR: candidate discovery requires access to the VBA project object model." & vbCrLf & vbeAccessDetail
    CorePulseProgress "Candidate scan | preparing System Macro metadata"
    PumpSetupResponsiveness True
    If SetupControlCancellationRequested() Then RequireSafeWorkContinue "ScanCandidatesCore", "Candidate scan phase boundary"
    Set systemMacroCatalogRows = SystemMacroCatalog()
    CorePulseProgress "Candidate scan | System Macro catalog ready | rows=" & CStr(systemMacroCatalogRows.count)
    Set systemMacroWrapperNames = BuildSystemMacroWrapperNameDictionary(systemMacroCatalogRows)
    CorePulseProgress "Candidate scan | wrapper exclusions ready | result role=" & resultConfigDatabaseRole & " | CodeDb=" & CStr(scanCodeDb) & " | HostDb=" & CStr(scanHostDb)
    prerequisiteElapsedSeconds = ElapsedSecondsFromTimer(phaseStartedTimer)
    scanBatchID = Format$(Now, "yyyymmdd_hhnnss") & "_" & resultConfigDatabaseRole
    phaseStartedTimer = Timer
    MarkPriorScanCandidateRowsStaleInDb resultConfigDb, scanBatchID, scanCodeDb, scanHostDb
    CorePulseProgress "Candidate scan | prior candidate rows marked stale | batch=" & scanBatchID
    Set rs = resultConfigDb.OpenRecordset(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates"), dbOpenDynaset)
    stalePreparationElapsedSeconds = ElapsedSecondsFromTimer(phaseStartedTimer)
    phaseStartedTimer = Timer
    CorePulseProgress "Candidate scan | configuration identity cache loading"
    PumpSetupResponsiveness True
    If SetupControlCancellationRequested() Then RequireSafeWorkContinue "ScanCandidatesCore", "Candidate scan phase boundary"
    Set scanCache = BuildCandidateScanCache(resultConfigDb, rs, systemMacroCatalogRows)
    CorePulseProgress "Candidate scan | configuration identity cache ready | beginning source discovery"
    cachePreloadElapsedSeconds = ElapsedSecondsFromTimer(phaseStartedTimer)
    If scanHostDb Then
        ScanCandidateTargetDb DatabaseForRole(CONFIG_OWNER_DATABASE_ROLE_HOST_DB), CONFIG_OWNER_DATABASE_ROLE_HOST_DB, rs, scanCache, scanBatchID, systemMacroWrapperNames, procedureCandidateCount, objectCandidateCount, objectTableCount, objectQueryCount, objectFormCount, objectReportCount, objectMacroCount, skippedObjectCount, skippedModuleCount, skippedBuiltInMacroTargetCount, procedureSourceSummary, objectSourceSummary, procedureScanElapsedSeconds, objectScanElapsedSeconds, tableScanElapsedSeconds, queryScanElapsedSeconds, queryNameFilterElapsedSeconds, queryHiddenSystemFilterElapsedSeconds, queryCandidateProcessingElapsedSeconds, formScanElapsedSeconds, reportScanElapsedSeconds, macroScanElapsedSeconds
        If SetupControlCancellationRequested() Then RequireSafeWorkContinue "ScanCandidatesCore", "HostDb candidate scan complete"
    End If
    If scanCodeDb Then
        ScanCandidateTargetDb codeDb, CONFIG_OWNER_DATABASE_ROLE_CODE_DB, rs, scanCache, scanBatchID, systemMacroWrapperNames, procedureCandidateCount, objectCandidateCount, objectTableCount, objectQueryCount, objectFormCount, objectReportCount, objectMacroCount, skippedObjectCount, skippedModuleCount, skippedBuiltInMacroTargetCount, procedureSourceSummary, objectSourceSummary, procedureScanElapsedSeconds, objectScanElapsedSeconds, tableScanElapsedSeconds, queryScanElapsedSeconds, queryNameFilterElapsedSeconds, queryHiddenSystemFilterElapsedSeconds, queryCandidateProcessingElapsedSeconds, formScanElapsedSeconds, reportScanElapsedSeconds, macroScanElapsedSeconds
        If SetupControlCancellationRequested() Then RequireSafeWorkContinue "ScanCandidatesCore", "CodeDb candidate scan complete"
    End If
    phaseStartedTimer = Timer
    foundCount = procedureCandidateCount + objectCandidateCount
    discoveryErrorCount = CandidateDiscoveryErrorRowCountInDb(resultConfigDb)
    CorePulseProgress "Candidate scan | finalizing | total candidates=" & CStr(foundCount) & " | procedures=" & CStr(procedureCandidateCount) & " | objects=" & CStr(objectCandidateCount) & " | skipped objects=" & CStr(skippedObjectCount) & " | errors=" & CStr(discoveryErrorCount)
    InvalidateDuplicateReviewCache
    finalizationElapsedSeconds = ElapsedSecondsFromTimer(phaseStartedTimer)
    totalElapsedSeconds = ElapsedSecondsFromTimer(scanStartedTimer)
    timingText = CandidateScanTimingSummary(prerequisiteElapsedSeconds, stalePreparationElapsedSeconds, cachePreloadElapsedSeconds, procedureScanElapsedSeconds, objectScanElapsedSeconds, tableScanElapsedSeconds, queryScanElapsedSeconds, queryNameFilterElapsedSeconds, queryHiddenSystemFilterElapsedSeconds, queryCandidateProcessingElapsedSeconds, formScanElapsedSeconds, reportScanElapsedSeconds, macroScanElapsedSeconds, finalizationElapsedSeconds, totalElapsedSeconds)
    ReportCandidateScanTiming timingText
    If showMessage Or discoveryErrorCount > 0 Then
        ' INTENTIONAL ORDERED VIEW: discoveryMessageLines preserves the deterministic sequence required by ScanCandidatesCore; keyed-only or unordered storage would not preserve the consumer semantics.
        Set discoveryMessageLines = New Collection
        discoveryMessageLines.Add IIf(discoveryErrorCount > 0, "Macro candidate discovery completed with errors.", "Macro candidate discovery complete.")
        discoveryMessageLines.Add ""
        discoveryMessageLines.Add "ResultConfigDb role: " & resultConfigDatabaseRole
        discoveryMessageLines.Add "ResultConfigDb: " & SharedDao_SafeDatabaseName(resultConfigDb)
        discoveryMessageLines.Add "Procedure ScanTargetDb roles: " & EmptyToPlaceholder(procedureSourceSummary, "None")
        discoveryMessageLines.Add "Access-object ScanTargetDb roles: " & EmptyToPlaceholder(objectSourceSummary, "None")
        discoveryMessageLines.Add "MC_Candidates table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates")
        discoveryMessageLines.Add "Candidate rows inserted/refreshed: " & CStr(foundCount)
        discoveryMessageLines.Add "Procedure candidate rows inserted/refreshed: " & CStr(procedureCandidateCount)
        discoveryMessageLines.Add "Object candidate rows inserted/refreshed: " & CStr(objectCandidateCount)
        discoveryMessageLines.Add " - Tables: " & CStr(objectTableCount)
        discoveryMessageLines.Add " - Queries: " & CStr(objectQueryCount)
        discoveryMessageLines.Add " - Forms: " & CStr(objectFormCount)
        discoveryMessageLines.Add " - Reports: " & CStr(objectReportCount)
        discoveryMessageLines.Add " - Existing macros: " & CStr(objectMacroCount)
        discoveryMessageLines.Add "Objects skipped as system/hidden/utility/generated: " & CStr(skippedObjectCount)
        discoveryMessageLines.Add "Built-in system macro target procedures skipped: " & CStr(skippedBuiltInMacroTargetCount)
        discoveryMessageLines.Add "Discovery error rows logged: " & CStr(discoveryErrorCount)
        discoveryMessageLines.Add "Non-standard modules skipped: " & CStr(skippedModuleCount)
        discoveryMessageLines.Add ""
        discoveryMessageLines.Add timingText
        discoveryMessageLines.Add ""
        discoveryMessageLines.Add "Candidate/discovery rows are stored in MC_Candidates. Review the candidate table before accepting selected rows into MC_Records."
        discoveryMessageLines.Add "Object candidates use MacroAction values OpenTable, OpenQuery, OpenForm, OpenReport, or RunMacro and store the target object name in TargetObjectName in MC_Candidates until acceptance creates the MC_Records configuration."
        discoveryMessageLines.Add "Candidate rows include CandidateOwnerDatabaseRole and default ConfigOwnerDatabaseRole values."
        discoveryMessageLines.Add "Set SelectedForAcceptance=True on rows you want to accept, then run " & SystemMacroCatalogTitleByKey(SYS_MACRO_KEY_ACCEPT_SELECTED_RUNNER_MACRO_CANDIDATES) & "."
        discoveryMessageLines.Add "Discovery errors, if any, are visible as CandidateStatus='ScanError' rows in MC_Candidates."
        discoveryMessageText = SharedCollection_Join(discoveryMessageLines, vbCrLf)
        discoveryMessageText = discoveryMessageText & CoreOperationSummaryForMessage()
        discoveryMessageText = discoveryMessageText & NavigationPaneGroupingSummaryForMessage()
        If discoveryErrorCount > 0 Then
            LogTraceMessage TRACE_LEVEL_ERROR, "CandidateDiscovery", "ScanCandidatesCore", "Macro candidate discovery completed with error row(s).", discoveryMessageText, "Complete candidate discovery", CANDIDATE_SCAN_ROW_ERRORS_ERROR_NUMBER, "ScanCandidatesCore", discoveryMessageText, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates"), "Table", "CodeDb", "CandidateDiscovery"
            If showMessage Then
                CoreShowErrorMessage discoveryMessageText, vbCritical, "Scan Macro Candidates Completed With Errors", "ScanCandidatesCore", "CandidateDiscovery"
            Else
                Err.Raise CANDIDATE_SCAN_ROW_ERRORS_ERROR_NUMBER, "ScanCandidatesCore", discoveryMessageText
            End If
        ElseIf showMessage Then
            CoreShowMessage discoveryMessageText, vbInformation, "Scan Macro Candidates"
        End If
    End If
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If savedErrNumber = CANDIDATE_SCAN_ROW_ERRORS_ERROR_NUMBER Then Err.Raise savedErrNumber, savedErrSource, savedErrDescription
    Err.Raise savedErrNumber, "ScanCandidatesCore", "Could not discover access development console candidates." & vbCrLf & "ResultConfigDb role: " & resultConfigDatabaseRole & vbCrLf & "ResultConfigDb: " & SafeDatabaseNameForError(resultConfigDb) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Function CandidateScanTimingSummary(ByVal prerequisiteElapsedSeconds As Double, ByVal stalePreparationElapsedSeconds As Double, ByVal cachePreloadElapsedSeconds As Double, ByVal procedureScanElapsedSeconds As Double, ByVal objectScanElapsedSeconds As Double, ByVal tableScanElapsedSeconds As Double, ByVal queryScanElapsedSeconds As Double, ByVal queryNameFilterElapsedSeconds As Double, ByVal queryHiddenSystemFilterElapsedSeconds As Double, ByVal queryCandidateProcessingElapsedSeconds As Double, ByVal formScanElapsedSeconds As Double, ByVal reportScanElapsedSeconds As Double, ByVal macroScanElapsedSeconds As Double, ByVal finalizationElapsedSeconds As Double, ByVal totalElapsedSeconds As Double) As String
    On Error GoTo ErrHandler
    Dim lines As Collection
    ' INTENTIONAL ORDERED VIEW: lines preserves the deterministic sequence required by CandidateScanTimingSummary; keyed-only or unordered storage would not preserve the consumer semantics.
    Set lines = New Collection
    lines.Add "Scan timing:"
    lines.Add " - Prerequisites and wrapper cache: " & FormattedElapsedSeconds(prerequisiteElapsedSeconds)
    lines.Add " - Stale-row preparation and Config recordset open: " & FormattedElapsedSeconds(stalePreparationElapsedSeconds)
    lines.Add " - Config identity, bookmark, field, and macro-title cache preload: " & FormattedElapsedSeconds(cachePreloadElapsedSeconds)
    lines.Add " - Procedure scanning and candidate processing: " & FormattedElapsedSeconds(procedureScanElapsedSeconds)
    lines.Add " - Access-object scanning and candidate processing: " & FormattedElapsedSeconds(objectScanElapsedSeconds)
    lines.Add "   - Tables: " & FormattedElapsedSeconds(tableScanElapsedSeconds)
    lines.Add "   - Queries: " & FormattedElapsedSeconds(queryScanElapsedSeconds)
    lines.Add "     - Name exclusion: " & FormattedElapsedSeconds(queryNameFilterElapsedSeconds)
    lines.Add "     - Hidden/system classification: " & FormattedElapsedSeconds(queryHiddenSystemFilterElapsedSeconds)
    lines.Add "     - Candidate comparison/update: " & FormattedElapsedSeconds(queryCandidateProcessingElapsedSeconds)
    lines.Add "   - Forms: " & FormattedElapsedSeconds(formScanElapsedSeconds)
    lines.Add "   - Reports: " & FormattedElapsedSeconds(reportScanElapsedSeconds)
    lines.Add "   - Macros: " & FormattedElapsedSeconds(macroScanElapsedSeconds)
    lines.Add " - Final counts and cache invalidation: " & FormattedElapsedSeconds(finalizationElapsedSeconds)
    lines.Add " - Total: " & FormattedElapsedSeconds(totalElapsedSeconds)
    CandidateScanTimingSummary = SharedCollection_Join(lines, vbCrLf)
    Exit Function
ErrHandler:
    CandidateScanTimingSummary = "Scan timing could not be formatted. Err.Description: " & Err.Description
End Function

Private Sub ReportCandidateScanTiming(ByVal timingText As String)
    On Error Resume Next
    Debug.Print "Access Development Console candidate scan timing: " & Replace(timingText, vbCrLf, "; ")
    LogTraceMessage "Info", "Scan", "ScanCandidatesCore", "Candidate scan phase timing.", timingText
    On Error GoTo 0
End Sub

Private Function BuildCandidateScanCache(ByVal resultConfigDb As DAO.Database, ByVal rsCandidate As DAO.Recordset, ByVal systemMacroCatalogRows As Collection) As Object
    On Error GoTo ErrHandler
    Dim cache As Object
    Dim rsConfig As DAO.Recordset
    Dim fieldIndex As Long
    Dim rowID As Long
    Dim ScanSignature As String
    Dim configOwnerRole As String
    Dim TargetProcedureName As String
    Dim TargetProcedureModuleName As String
    Dim MacroAction As String
    Dim TargetObjectName As String
    Dim MacroTitle As String
    Dim originalBookmark As Variant
    Dim hadOriginalBookmark As Boolean
    Dim rowData As Variant
    ' INTENTIONAL KEYED LOOKUP: cache provides keyed membership/lookup over candidate rows and configuration identities already read by this scan; it is a derived access path, not a second authority.
    Set cache = CreateObject("Scripting.Dictionary")
    cache.CompareMode = vbTextCompare
    If resultConfigDb Is Nothing Then Err.Raise vbObjectError + 5460, "BuildCandidateScanCache", "ResultConfigDb reference is Nothing."
    If rsCandidate Is Nothing Then Err.Raise vbObjectError + 5461, "BuildCandidateScanCache", "Candidate recordset reference is Nothing."
    For fieldIndex = 0 To rsCandidate.fields.count - 1
        PumpSetupResponsiveness
        CandidateScanCacheAddIfMissing cache, CandidateScanCacheKey("FIELD", rsCandidate.fields(fieldIndex).Name), fieldIndex
    Next fieldIndex
    If Not (rsCandidate.BOF And rsCandidate.EOF) Then
        originalBookmark = rsCandidate.Bookmark
        hadOriginalBookmark = True
        rsCandidate.MoveFirst
        Do While Not rsCandidate.EOF
            PumpSetupResponsiveness
            rowID = SharedValue_NzLong(rsCandidate.fields(FIELD_ID).value)
            ScanSignature = SharedText_Nz(rsCandidate.fields(FIELD_SCAN_SIGNATURE).value)
            If rowID > 0 And Len(Trim$(ScanSignature)) > 0 Then
                If StrComp(SharedText_Nz(rsCandidate.fields(FIELD_CANDIDATE_STATUS).value), CANDIDATE_STATUS_SCAN_ERROR, vbTextCompare) = 0 Then
                    CandidateScanCacheAddIfMissing cache, CandidateScanCacheKey("ROW", CANDIDATE_STATUS_SCAN_ERROR, ScanSignature), rsCandidate.Bookmark
                Else
                    CandidateScanCacheAddIfMissing cache, CandidateScanCacheKey("ROW", ROW_KIND_CANDIDATE, ScanSignature), rsCandidate.Bookmark
                End If
            End If
            rsCandidate.MoveNext
        Loop
    End If
    Set rsConfig = resultConfigDb.OpenRecordset("SELECT [" & FIELD_CONFIG_OWNER_DATABASE_ROLE & "]," & SQL_FIELD_TARGET_PROCEDURE_NAME & "," & SQL_FIELD_TARGET_PROCEDURE_MODULE_NAME & "," & SQL_FIELD_MACRO_ACTION & "," & SQL_FIELD_TARGET_OBJECT_NAME & "," & SQL_FIELD_MACRO_TITLE & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "]", dbOpenSnapshot)
    Do While Not rsConfig.EOF
        PumpSetupResponsiveness
        configOwnerRole = SharedText_Nz(rsConfig.fields(FIELD_CONFIG_OWNER_DATABASE_ROLE).value)
        TargetProcedureName = SharedText_Nz(rsConfig.fields(FIELD_TARGET_PROCEDURE_NAME).value)
        TargetProcedureModuleName = SharedText_Nz(rsConfig.fields(FIELD_TARGET_PROCEDURE_MODULE_NAME).value)
        MacroAction = SharedText_Nz(rsConfig.fields(FIELD_MACRO_ACTION).value)
        TargetObjectName = SharedText_Nz(rsConfig.fields(FIELD_TARGET_OBJECT_NAME).value)
        MacroTitle = SharedText_Nz(rsConfig.fields(FIELD_MACRO_TITLE).value)
        If Len(Trim$(TargetProcedureName)) > 0 Then CandidateScanCacheAddIfMissing cache, CandidateScanCacheKey("CFGPROC", configOwnerRole, TargetProcedureName, TargetProcedureModuleName), True
        If Len(Trim$(MacroAction)) > 0 And Len(Trim$(TargetObjectName)) > 0 Then CandidateScanCacheAddIfMissing cache, CandidateScanExactIdentityCacheKey("CFGOBJ", Trim$(MacroAction), TargetObjectName), True
        If Len(Trim$(MacroTitle)) > 0 Then CandidateScanCacheAddIfMissing cache, CandidateScanExactIdentityCacheKey("CFGMACROTITLE", MacroTitle), True
        rsConfig.MoveNext
    Loop
    rsConfig.Close
    Set rsConfig = Nothing
    For Each rowData In systemMacroCatalogRows
        PumpSetupResponsiveness
        MacroTitle = SystemMacroCatalogTitle(rowData)
        If Len(Trim$(MacroTitle)) > 0 Then CandidateScanCacheAddIfMissing cache, CandidateScanExactIdentityCacheKey("SYSTEMMACROTITLE", MacroTitle), True
    Next rowData
    If hadOriginalBookmark Then rsCandidate.Bookmark = originalBookmark
    Set BuildCandidateScanCache = cache
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not rsConfig Is Nothing Then rsConfig.Close
    If hadOriginalBookmark Then rsCandidate.Bookmark = originalBookmark
    On Error GoTo 0
    Err.Raise savedErrNumber, "BuildCandidateScanCache", "Could not preload Access Development Console candidate rows and configuration identity exclusions." & vbCrLf & "ResultConfigDb: " & SafeDatabaseNameForError(resultConfigDb) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function
Private Sub CandidateScanCacheAddIfMissing(ByVal scanCache As Object, ByVal cacheKey As String, ByVal cacheValue As Variant)
    If scanCache Is Nothing Then Exit Sub
    If Len(cacheKey) = 0 Then Exit Sub
    If scanCache.Exists(cacheKey) = False Then scanCache.Add cacheKey, cacheValue
End Sub

Private Function CandidateScanExactIdentityCacheKey(ByVal categoryText As String, ByVal part1 As String, Optional ByVal part2 As String = vbNullString, Optional ByVal part3 As String = vbNullString) As String
    Dim delimiterText As String
    delimiterText = Chr$(30)
    CandidateScanExactIdentityCacheKey = LCase$(categoryText) & delimiterText & LCase$(part1) & delimiterText & LCase$(part2) & delimiterText & LCase$(part3)
End Function

Private Function CandidateScanCacheHasConfiguredProcedure(ByVal scanCache As Object, ByVal configOwnerRole As String, ByVal TargetProcedureName As String, ByVal TargetProcedureModuleName As String) As Boolean
    If scanCache Is Nothing Then Exit Function
    CandidateScanCacheHasConfiguredProcedure = scanCache.Exists(CandidateScanCacheKey("CFGPROC", configOwnerRole, TargetProcedureName, TargetProcedureModuleName))
End Function

Private Function CandidateScanCacheHasConfiguredAccessObject(ByVal scanCache As Object, ByVal MacroAction As String, ByVal TargetObjectName As String) As Boolean
    If scanCache Is Nothing Then Exit Function
    CandidateScanCacheHasConfiguredAccessObject = scanCache.Exists(CandidateScanExactIdentityCacheKey("CFGOBJ", Trim$(MacroAction), TargetObjectName))
End Function

Private Sub ScanCandidateTargetDb(ByVal scanTargetDb As DAO.Database, ByVal CandidateOwnerDatabaseRole As String, ByVal rs As DAO.Recordset, ByVal scanCache As Object, ByVal scanBatchID As String, ByVal systemMacroWrapperNames As Object, ByRef procedureCandidateCount As Long, ByRef objectCandidateCount As Long, ByRef objectTableCount As Long, ByRef objectQueryCount As Long, ByRef objectFormCount As Long, ByRef objectReportCount As Long, ByRef objectMacroCount As Long, ByRef skippedObjectCount As Long, ByRef skippedModuleCount As Long, ByRef skippedBuiltInMacroTargetCount As Long, ByRef procedureSourceSummary As String, ByRef objectSourceSummary As String, ByRef procedureScanElapsedSeconds As Double, ByRef objectScanElapsedSeconds As Double, _
ByRef tableScanElapsedSeconds As Double, ByRef queryScanElapsedSeconds As Double, ByRef queryNameFilterElapsedSeconds As Double, ByRef queryHiddenSystemFilterElapsedSeconds As Double, ByRef queryCandidateProcessingElapsedSeconds As Double, ByRef formScanElapsedSeconds As Double, ByRef reportScanElapsedSeconds As Double, ByRef macroScanElapsedSeconds As Double)
    On Error GoTo ErrHandler
    Dim vbProject As Object
    Dim phaseStartedTimer As Single
    If scanTargetDb Is Nothing Then Err.Raise vbObjectError + 5403, "ScanCandidateTargetDb", "ScanTargetDb reference is Nothing."
    If Len(CandidateOwnerDatabaseRole) = 0 Then Err.Raise vbObjectError + 5404, "ScanCandidateTargetDb", "ScanTargetDb role is not valid."
    phaseStartedTimer = Timer
    Set vbProject = VBProjectForConfigOwnerDatabaseRole(CandidateOwnerDatabaseRole)
    If vbProject Is Nothing Then
        AddCandidateDiscoveryErrorRow rs, Nothing, Nothing, CandidateOwnerDatabaseRole, scanTargetDb.Name, "Candidate discovery could not resolve the ScanTargetDb VBProject. Procedure scanning was skipped for this ScanTargetDb role.", False, False, vbNullString, vbNullString, vbNullString, vbNullString, scanBatchID, scanCache
        CoreAddOperationNote "Procedure scan skipped for " & CandidateOwnerDatabaseRole & " because its VBProject could not be resolved."
    Else
        CorePulseProgress "Candidate scan [" & CandidateOwnerDatabaseRole & "] | VBA procedures | project=" & SafeObjectText(vbProject, "Name") & " | candidates=" & CStr(procedureCandidateCount)
        AppendScanSourceSummary procedureSourceSummary, CandidateOwnerDatabaseRole
        AddCandidatesFromProject vbProject, CandidateOwnerDatabaseRole, scanTargetDb.Name, rs, scanCache, systemMacroWrapperNames, procedureCandidateCount, skippedModuleCount, skippedBuiltInMacroTargetCount, scanBatchID
    End If
    procedureScanElapsedSeconds = procedureScanElapsedSeconds + ElapsedSecondsFromTimer(phaseStartedTimer)
    phaseStartedTimer = Timer
    CorePulseProgress "Candidate scan [" & CandidateOwnerDatabaseRole & "] | Access objects | procedures=" & CStr(procedureCandidateCount) & " | object candidates=" & CStr(objectCandidateCount) & " | skipped objects=" & CStr(skippedObjectCount)
    AppendScanSourceSummary objectSourceSummary, CandidateOwnerDatabaseRole
    objectCandidateCount = objectCandidateCount + AddObjectCandidatesFromScanTargetDb(scanTargetDb, CandidateOwnerDatabaseRole, rs, scanCache, scanBatchID, objectTableCount, objectQueryCount, objectFormCount, objectReportCount, objectMacroCount, skippedObjectCount, tableScanElapsedSeconds, queryScanElapsedSeconds, queryNameFilterElapsedSeconds, queryHiddenSystemFilterElapsedSeconds, queryCandidateProcessingElapsedSeconds, formScanElapsedSeconds, reportScanElapsedSeconds, macroScanElapsedSeconds)
    CorePulseProgress "Candidate scan [" & CandidateOwnerDatabaseRole & "] complete | procedures=" & CStr(procedureCandidateCount) & " | objects=" & CStr(objectCandidateCount) & " | skipped objects=" & CStr(skippedObjectCount) & " | skipped modules=" & CStr(skippedModuleCount)
    objectScanElapsedSeconds = objectScanElapsedSeconds + ElapsedSecondsFromTimer(phaseStartedTimer)
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ScanCandidateTargetDb", "Could not scan Access Development Console candidates for ScanTargetDb." & vbCrLf & "ScanTargetDb role: " & CandidateOwnerDatabaseRole & vbCrLf & "ScanTargetDb: " & SafeDatabaseNameForError(scanTargetDb) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub AddCandidatesFromProject(ByVal vbProject As Object, ByVal CandidateOwnerDatabaseRole As String, ByVal scanTargetDbName As String, ByVal rs As DAO.Recordset, ByVal scanCache As Object, ByVal systemMacroWrapperNames As Object, ByRef scanWideProcedureCandidateCount As Long, ByRef skippedModuleCount As Long, ByRef skippedBuiltInMacroTargetCount As Long, ByVal scanBatchID As String)
    On Error GoTo ErrHandler
    Dim vbComponent As Object
    Dim componentIndex As Long
    Dim componentCount As Long
    Dim standardModuleCount As Long
    Dim lastModuleStatusTimer As Single
    componentCount = vbProject.VBComponents.count
    lastModuleStatusTimer = Timer
    For Each vbComponent In vbProject.VBComponents
        PumpSetupResponsiveness
        componentIndex = componentIndex + 1
        If vbComponent.Type = VBEXT_CT_STD_MODULE Then
            standardModuleCount = standardModuleCount + 1
            If standardModuleCount = 1 Or ElapsedSecondsFromTimer(lastModuleStatusTimer) >= 1# Then
                CorePulseProgress "Candidate scan [" & CandidateOwnerDatabaseRole & "] | module " & CStr(componentIndex) & "/" & CStr(componentCount) & " | " & SafeObjectText(vbComponent, "Name") & " | procedure candidates=" & CStr(scanWideProcedureCandidateCount) & " | non-standard skipped=" & CStr(skippedModuleCount)
                lastModuleStatusTimer = Timer
            End If
            AddCandidatesFromStandardModule vbProject, vbComponent, rs, scanCache, CandidateOwnerDatabaseRole, scanTargetDbName, systemMacroWrapperNames, scanWideProcedureCandidateCount, skippedBuiltInMacroTargetCount, scanBatchID
        Else
            skippedModuleCount = skippedModuleCount + 1
        End If
    Next vbComponent
    Exit Sub
ErrHandler:
    AddCandidateDiscoveryErrorRow rs, vbProject, Nothing, CandidateOwnerDatabaseRole, scanTargetDbName, "Candidate discovery could not inspect this VBProject." & vbCrLf & "Err.Number: " & CStr(Err.Number) & vbCrLf & "Err.Description: " & Err.Description, False, False, vbNullString, vbNullString, vbNullString, vbNullString, scanBatchID, scanCache
End Sub

Private Sub AddCandidatesFromStandardModule(ByVal vbProject As Object, ByVal vbComponent As Object, ByVal rs As DAO.Recordset, ByVal scanCache As Object, ByVal CandidateOwnerDatabaseRole As String, ByVal scanTargetDbName As String, ByVal systemMacroWrapperNames As Object, ByRef scanWideProcedureCandidateCount As Long, ByRef skippedBuiltInMacroTargetCount As Long, ByVal scanBatchID As String)
    On Error GoTo ErrHandler
    Dim codeModule As Object
    Dim lineIndex As Long
    Dim procName As String
    Dim startLine As Long
    Dim countLines As Long
    Dim declarationText As String
    Dim procedureKind As String
    Dim visibilityText As String
    Dim argumentSignature As String
    Dim RequiredArgumentCount As Long
    Dim OptionalArgumentCount As Long
    Dim ExpectedReturnType As String
    Dim ModuleHeaderFound As Boolean
    Dim ModuleHeaderMalformed As Boolean
    Dim ModuleHeaderTitle As String
    Dim ModuleHeaderPurpose As String
    Dim moduleHeaderEntryPointsRaw As String
    Dim ModuleHeaderScanNotes As String
    Dim moduleTopText As String
    Dim isCoreUtilityModule As Boolean
    Dim isGeneratedUiUtilityModule As Boolean
    Dim exactEntryPointNames As Object
    Dim allEntryPointIdentifiers As Object
    Dim errorText As String
    Set codeModule = vbComponent.codeModule
    moduleTopText = CodeModuleTopText(codeModule)
    isCoreUtilityModule = CandidateComponentIsThisUtility(vbComponent, moduleTopText)
    isGeneratedUiUtilityModule = CandidateComponentIsGeneratedUiUtility(vbComponent, moduleTopText)
    ReadModuleHeaderMetadata moduleTopText, ModuleHeaderFound, ModuleHeaderMalformed, ModuleHeaderTitle, ModuleHeaderPurpose, moduleHeaderEntryPointsRaw, ModuleHeaderScanNotes
    BuildEntryPointNameDictionaries moduleHeaderEntryPointsRaw, exactEntryPointNames, allEntryPointIdentifiers
    lineIndex = codeModule.CountOfDeclarationLines + 1
    If lineIndex < 1 Then lineIndex = 1
    Do While lineIndex <= codeModule.CountOfLines
        PumpSetupResponsiveness
        procName = ProcedureNameAtCodeLine(codeModule, lineIndex)
        If Len(procName) = 0 Then
            lineIndex = lineIndex + 1
        Else
            startLine = ProcedureStartLineOrDefault(codeModule, procName, lineIndex)
            countLines = ProcedureCountLinesOrDefault(codeModule, procName)
            declarationText = ProcedureDeclarationText(codeModule, startLine)
            If CandidateProcedureDeclarationInfo(declarationText, procName, procedureKind, visibilityText, argumentSignature, RequiredArgumentCount, OptionalArgumentCount, ExpectedReturnType) Then
                If CandidateIsBuiltInSystemMacroTarget(isCoreUtilityModule, isGeneratedUiUtilityModule, systemMacroWrapperNames, procName) Then
                    skippedBuiltInMacroTargetCount = skippedBuiltInMacroTargetCount + 1
                ElseIf AddCandidateRow(rs, scanCache, vbProject, vbComponent, CandidateOwnerDatabaseRole, scanTargetDbName, procName, procedureKind, visibilityText, declarationText, argumentSignature, RequiredArgumentCount, OptionalArgumentCount, ExpectedReturnType, ModuleHeaderFound, ModuleHeaderMalformed, ModuleHeaderTitle, ModuleHeaderPurpose, moduleHeaderEntryPointsRaw, ModuleHeaderScanNotes, exactEntryPointNames, allEntryPointIdentifiers, scanBatchID) Then
                    scanWideProcedureCandidateCount = scanWideProcedureCandidateCount + 1
                    PulseProgressEvery scanWideProcedureCandidateCount, 50, "Candidate scan [" & CandidateOwnerDatabaseRole & "] | procedure " & SafeObjectText(vbComponent, "Name") & "." & procName & " | candidates=" & CStr(scanWideProcedureCandidateCount) & " | built-in targets skipped=" & CStr(skippedBuiltInMacroTargetCount), 0, "CandidateProcedureScan", "CandidateDiscovery"
                End If
            End If
            If countLines > 0 Then
                lineIndex = startLine + countLines
            Else
                lineIndex = lineIndex + 1
            End If
        End If
    Loop
    Exit Sub
ErrHandler:
    errorText = "Candidate discovery could not inspect this standard module." & vbCrLf
    errorText = errorText & "The module was skipped so the discovery run could continue." & vbCrLf
    errorText = errorText & "Err.Number: " & CStr(Err.Number) & vbCrLf
    errorText = errorText & "Err.Description: " & Err.Description
    AddCandidateDiscoveryErrorRow rs, vbProject, vbComponent, CandidateOwnerDatabaseRole, scanTargetDbName, errorText, ModuleHeaderFound, ModuleHeaderMalformed, ModuleHeaderTitle, ModuleHeaderPurpose, moduleHeaderEntryPointsRaw, ModuleHeaderScanNotes, scanBatchID, scanCache
End Sub

Private Function AddCandidateRow(ByVal rs As DAO.Recordset, ByVal scanCache As Object, ByVal vbProject As Object, ByVal vbComponent As Object, ByVal CandidateOwnerDatabaseRole As String, ByVal scanTargetDbName As String, ByVal procName As String, ByVal procedureKind As String, ByVal visibilityText As String, ByVal declarationText As String, ByVal argumentSignature As String, ByVal RequiredArgumentCount As Long, ByVal OptionalArgumentCount As Long, ByVal ExpectedReturnType As String, ByVal ModuleHeaderFound As Boolean, ByVal ModuleHeaderMalformed As Boolean, ByVal ModuleHeaderTitle As String, ByVal ModuleHeaderPurpose As String, ByVal moduleHeaderEntryPointsRaw As String, ByVal ModuleHeaderScanNotes As String, ByVal exactEntryPointNames As Object, ByVal allEntryPointIdentifiers As Object, ByVal scanBatchID As String) As Boolean
    On Error GoTo ErrHandler
    Dim candidateMacroKeyText As String
    Dim ScanSignature As String
    Dim ResultHandling As String
    Dim descriptionText As String
    Dim EntryPointMatchConfidence As String
    Dim listedInEntryPoints As Boolean
    Dim foundExistingCandidate As Boolean
    Dim cachedBookmark As Variant
    EntryPointMatchConfidence = EntryPointMatchConfidenceForProcedure(procName, moduleHeaderEntryPointsRaw, ModuleHeaderFound, ModuleHeaderMalformed, exactEntryPointNames, allEntryPointIdentifiers)
    listedInEntryPoints = (StrComp(EntryPointMatchConfidence, "Exact", vbTextCompare) = 0)
    candidateMacroKeyText = CandidateMacroKey(CandidateOwnerDatabaseRole, vbComponent.Name, procName)
    ScanSignature = CandidateScanSignature(CandidateOwnerDatabaseRole, scanTargetDbName, vbComponent.Name, procName, procedureKind, argumentSignature)
    If CandidateScanCacheHasConfiguredProcedure(scanCache, CandidateOwnerDatabaseRole, procName, vbComponent.Name) Then
        AddCandidateRow = False
        Exit Function
    End If
    If StrComp(procedureKind, TARGET_PROCEDURE_KIND_SUB, vbTextCompare) = 0 Then
        ResultHandling = RESULT_HANDLING_IGNORE
    Else
        ResultHandling = RESULT_HANDLING_AUTO
    End If
    descriptionText = "Candidate discovered from " & visibilityText & " " & procedureKind & " in standard module " & vbComponent.Name & ". Required arguments: " & CStr(RequiredArgumentCount) & "; optional arguments: " & CStr(OptionalArgumentCount) & ". Module header found: " & CStr(ModuleHeaderFound) & "; Entry Points match: " & EntryPointMatchConfidence & "."
    cachedBookmark = CachedScanRowBookmark(scanCache, ROW_KIND_CANDIDATE, ScanSignature)
    foundExistingCandidate = PositionCachedScanRow(rs, cachedBookmark, ROW_KIND_CANDIDATE, ScanSignature)
    If foundExistingCandidate Then
        rs.Edit
    Else
        rs.AddNew
    End If
    rs.fields(FIELD_CANDIDATE_STATUS) = CANDIDATE_STATUS_NEW
    rs.fields(FIELD_MACRO_KEY) = Null
    rs.fields(FIELD_CANDIDATE_MACRO_KEY) = candidateMacroKeyText
    rs.fields(FIELD_SCAN_SIGNATURE) = ScanSignature
    rs.fields(FIELD_LAST_SCAN_BATCH_ID) = scanBatchID
    rs.fields(FIELD_LAST_SCANNED_ON) = Now
    rs.fields(FIELD_REJECTED_ON) = Null
    rs.fields(FIELD_MACRO_TITLE) = vbComponent.Name & " - " & procName
    rs.fields(FIELD_MACRO_ACTION) = ACTION_RUNCODE
    SetScanFieldIfAvailable rs, scanCache, FIELD_CONFIG_OWNER_DATABASE_ROLE, CandidateOwnerDatabaseRole
    rs.fields(FIELD_TARGET_PROCEDURE_NAME) = procName
    rs.fields(FIELD_TARGET_PROCEDURE_KIND) = procedureKind
    rs.fields(FIELD_TARGET_ARGUMENTS) = Null
    rs.fields(FIELD_RESULT_HANDLING) = ResultHandling
    rs.fields(FIELD_EXPECTED_RETURN_TYPE) = Shared_String_ConvertToNullIfBlank(ExpectedReturnType)
    SetScanFieldIfAvailable rs, scanCache, FIELD_CANDIDATE_OWNER_DATABASE_ROLE, CandidateOwnerDatabaseRole
    rs.fields(FIELD_CANDIDATE_SCAN_TARGET_DB_NAME) = DatabaseBaseNameFromPath(scanTargetDbName)
    SetScanFieldIfAvailable rs, scanCache, FIELD_CANDIDATE_SCAN_TARGET_DB_PATH, scanTargetDbName
    rs.fields(FIELD_CANDIDATE_VB_PROJECT_NAME) = SafeObjectText(vbProject, "Name")
    rs.fields(FIELD_CANDIDATE_MODULE_NAME) = vbComponent.Name
    SetScanFieldIfAvailable rs, scanCache, FIELD_GENERATION_TARGET_DB_NAME, DatabaseBaseNameFromPath(DatabaseForRole(CandidateOwnerDatabaseRole).Name)
    SetScanFieldIfAvailable rs, scanCache, FIELD_GENERATION_TARGET_DB_PATH, DatabasePathForRole(CandidateOwnerDatabaseRole)
    SetScanFieldIfAvailable rs, scanCache, FIELD_TARGET_REFERENCE_NAME, IIf(StrComp(CandidateOwnerDatabaseRole, CONFIG_OWNER_DATABASE_ROLE_CODE_DB, vbTextCompare) = 0, SafeObjectText(vbProject, "Name"), Null)
    SetScanFieldIfAvailable rs, scanCache, FIELD_TARGET_VB_PROJECT_NAME, SafeObjectText(vbProject, "Name")
    SetScanFieldIfAvailable rs, scanCache, FIELD_TARGET_PROCEDURE_MODULE_NAME, vbComponent.Name
    rs.fields(FIELD_MODULE_HEADER_FOUND) = ModuleHeaderFound
    rs.fields(FIELD_MODULE_HEADER_MALFORMED) = ModuleHeaderMalformed
    rs.fields(FIELD_MODULE_HEADER_TITLE) = Shared_String_ConvertToNullIfBlank(ModuleHeaderTitle)
    rs.fields(FIELD_MODULE_HEADER_PURPOSE) = Shared_String_ConvertToNullIfBlank(ModuleHeaderPurpose)
    rs.fields(FIELD_MODULE_HEADER_ENTRY_POINTS) = Shared_String_ConvertToNullIfBlank(moduleHeaderEntryPointsRaw)
    rs.fields(FIELD_PROCEDURE_LISTED_IN_ENTRY_POINTS) = listedInEntryPoints
    rs.fields(FIELD_ENTRY_POINT_MATCH_CONFIDENCE) = Shared_String_ConvertToNullIfBlank(EntryPointMatchConfidence)
    rs.fields(FIELD_MODULE_HEADER_SCAN_NOTES) = Shared_String_ConvertToNullIfBlank(ModuleHeaderScanNotes)
    rs.fields(FIELD_CANDIDATE_PROCEDURE_DECLARATION) = Shared_String_ConvertToNullIfBlank(declarationText)
    rs.fields(FIELD_CANDIDATE_ARGUMENT_SIGNATURE) = Shared_String_ConvertToNullIfBlank(argumentSignature)
    rs.fields(FIELD_HAS_REQUIRED_ARGUMENTS) = (RequiredArgumentCount > 0)
    rs.fields(FIELD_REQUIRED_ARGUMENT_COUNT) = RequiredArgumentCount
    rs.fields(FIELD_OPTIONAL_ARGUMENT_COUNT) = OptionalArgumentCount
    rs.fields(FIELD_DISCOVERED_ON) = Now
    rs.fields(FIELD_TARGET_OBJECT_NAME) = Null
    SetScanFieldIfAvailable rs, scanCache, FIELD_NAVIGATION_PANE_GROUP, BuildNavigationPaneGroupName(DEFAULT_GENERATED_MACROS_NAVIGATION_PANE_GROUP_LOGICAL_NAME, False)
    rs.fields(FIELD_CONFIRMATION_MODE) = CONFIRM_ALWAYS
    rs.fields(FIELD_ACTION_DESCRIPTION) = Shared_String_ConvertToNullIfBlank(descriptionText)
    rs.fields(FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER) = CandidateSortOrder(procedureKind, RequiredArgumentCount, OptionalArgumentCount)
    rs.fields(FIELD_SELECTED_FOR_ACCEPTANCE) = False
    rs.fields(FIELD_ACCEPTANCE_STATUS) = Null
    rs.fields(FIELD_ACCEPTED_ON) = Null
    rs.fields(FIELD_ACCEPTED_CONFIGURATION_ID) = Null
    rs.fields(FIELD_ACCEPTANCE_NOTES) = Null
    rs.Update
    RememberCachedScanRow rs, scanCache, ROW_KIND_CANDIDATE, ScanSignature, foundExistingCandidate
    AddCandidateRow = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "AddCandidateRow", "Could not add or refresh access development console candidate row." & vbCrLf & "Procedure: " & procName & vbCrLf & "Module: " & SafeObjectNameForError(vbComponent) & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Sub AddCandidateDiscoveryErrorRow(ByVal rs As DAO.Recordset, ByVal vbProject As Object, ByVal vbComponent As Object, ByVal CandidateOwnerDatabaseRole As String, ByVal scanTargetDbName As String, ByVal errorDetailText As String, ByVal ModuleHeaderFound As Boolean, ByVal ModuleHeaderMalformed As Boolean, ByVal ModuleHeaderTitle As String, ByVal ModuleHeaderPurpose As String, ByVal moduleHeaderEntryPointsRaw As String, ByVal ModuleHeaderScanNotes As String, ByVal scanBatchID As String, ByVal scanCache As Object)
    On Error GoTo ErrHandler
    Dim ModuleName As String
    Dim candidateMacroKeyText As String
    Dim ScanSignature As String
    Dim notesText As String
    Dim foundExistingError As Boolean
    Dim cachedBookmark As Variant
    If vbComponent Is Nothing Then
        ModuleName = "(project)"
    Else
        ModuleName = vbComponent.Name
    End If
    candidateMacroKeyText = CandidateDiscoveryErrorMacroKey(CandidateOwnerDatabaseRole, ModuleName)
    ScanSignature = CandidateScanSignature(CandidateOwnerDatabaseRole, scanTargetDbName, ModuleName, "(DiscoveryError)", CANDIDATE_STATUS_SCAN_ERROR, vbNullString)
    notesText = errorDetailText
    cachedBookmark = CachedScanRowBookmark(scanCache, CANDIDATE_STATUS_SCAN_ERROR, ScanSignature)
    foundExistingError = PositionCachedScanRow(rs, cachedBookmark, CANDIDATE_STATUS_SCAN_ERROR, ScanSignature)
    If foundExistingError Then
        rs.Edit
    Else
        rs.AddNew
    End If
    rs.fields(FIELD_CANDIDATE_STATUS) = CANDIDATE_STATUS_SCAN_ERROR
    rs.fields(FIELD_MACRO_KEY) = Null
    rs.fields(FIELD_CANDIDATE_MACRO_KEY) = candidateMacroKeyText
    rs.fields(FIELD_SCAN_SIGNATURE) = ScanSignature
    rs.fields(FIELD_LAST_SCAN_BATCH_ID) = scanBatchID
    rs.fields(FIELD_LAST_SCANNED_ON) = Now
    rs.fields(FIELD_REJECTED_ON) = Null
    rs.fields(FIELD_MACRO_TITLE) = "DISCOVERY ERROR - " & ModuleName
    rs.fields(FIELD_MACRO_ACTION) = Null
    SetScanFieldIfAvailable rs, scanCache, FIELD_CONFIG_OWNER_DATABASE_ROLE, CandidateOwnerDatabaseRole
    rs.fields(FIELD_TARGET_PROCEDURE_NAME) = Null
    rs.fields(FIELD_TARGET_PROCEDURE_KIND) = Null
    rs.fields(FIELD_TARGET_ARGUMENTS) = Null
    rs.fields(FIELD_RESULT_HANDLING) = Null
    rs.fields(FIELD_EXPECTED_RETURN_TYPE) = Null
    SetScanFieldIfAvailable rs, scanCache, FIELD_CANDIDATE_OWNER_DATABASE_ROLE, CandidateOwnerDatabaseRole
    rs.fields(FIELD_CANDIDATE_SCAN_TARGET_DB_NAME) = DatabaseBaseNameFromPath(scanTargetDbName)
    SetScanFieldIfAvailable rs, scanCache, FIELD_CANDIDATE_SCAN_TARGET_DB_PATH, scanTargetDbName
    rs.fields(FIELD_CANDIDATE_VB_PROJECT_NAME) = SafeObjectText(vbProject, "Name")
    rs.fields(FIELD_CANDIDATE_MODULE_NAME) = ModuleName
    SetScanFieldIfAvailable rs, scanCache, FIELD_GENERATION_TARGET_DB_NAME, DatabaseBaseNameFromPath(DatabaseForRole(CandidateOwnerDatabaseRole).Name)
    SetScanFieldIfAvailable rs, scanCache, FIELD_GENERATION_TARGET_DB_PATH, DatabasePathForRole(CandidateOwnerDatabaseRole)
    SetScanFieldIfAvailable rs, scanCache, FIELD_TARGET_REFERENCE_NAME, IIf(StrComp(CandidateOwnerDatabaseRole, CONFIG_OWNER_DATABASE_ROLE_CODE_DB, vbTextCompare) = 0, SafeObjectText(vbProject, "Name"), Null)
    SetScanFieldIfAvailable rs, scanCache, FIELD_TARGET_VB_PROJECT_NAME, SafeObjectText(vbProject, "Name")
    SetScanFieldIfAvailable rs, scanCache, FIELD_TARGET_PROCEDURE_MODULE_NAME, IIf(ModuleName = "(project)", Null, ModuleName)
    rs.fields(FIELD_MODULE_HEADER_FOUND) = ModuleHeaderFound
    rs.fields(FIELD_MODULE_HEADER_MALFORMED) = ModuleHeaderMalformed
    rs.fields(FIELD_MODULE_HEADER_TITLE) = Shared_String_ConvertToNullIfBlank(ModuleHeaderTitle)
    rs.fields(FIELD_MODULE_HEADER_PURPOSE) = Shared_String_ConvertToNullIfBlank(ModuleHeaderPurpose)
    rs.fields(FIELD_MODULE_HEADER_ENTRY_POINTS) = Shared_String_ConvertToNullIfBlank(moduleHeaderEntryPointsRaw)
    rs.fields(FIELD_PROCEDURE_LISTED_IN_ENTRY_POINTS) = False
    rs.fields(FIELD_ENTRY_POINT_MATCH_CONFIDENCE) = Null
    rs.fields(FIELD_MODULE_HEADER_SCAN_NOTES) = Shared_String_ConvertToNullIfBlank(ModuleHeaderScanNotes)
    rs.fields(FIELD_CANDIDATE_PROCEDURE_DECLARATION) = Null
    rs.fields(FIELD_CANDIDATE_ARGUMENT_SIGNATURE) = Null
    rs.fields(FIELD_HAS_REQUIRED_ARGUMENTS) = False
    rs.fields(FIELD_REQUIRED_ARGUMENT_COUNT) = 0
    rs.fields(FIELD_OPTIONAL_ARGUMENT_COUNT) = 0
    rs.fields(FIELD_DISCOVERED_ON) = Now
    rs.fields(FIELD_CANDIDATE_NOTES) = Shared_String_ConvertToNullIfBlank(notesText)
    rs.fields(FIELD_TARGET_OBJECT_NAME) = Null
    SetScanFieldIfAvailable rs, scanCache, FIELD_NAVIGATION_PANE_GROUP, BuildNavigationPaneGroupName(DEFAULT_GENERATED_MACROS_NAVIGATION_PANE_GROUP_LOGICAL_NAME, False)
    rs.fields(FIELD_CONFIRMATION_MODE) = CONFIRM_NEVER
    rs.fields(FIELD_ACTION_DESCRIPTION) = "Candidate discovery could not inspect this standard module. Review CandidateNotes."
    rs.fields(FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER) = 900000
    rs.fields(FIELD_SELECTED_FOR_ACCEPTANCE) = False
    rs.fields(FIELD_ACCEPTANCE_STATUS) = Null
    rs.fields(FIELD_ACCEPTED_ON) = Null
    rs.fields(FIELD_ACCEPTED_CONFIGURATION_ID) = Null
    rs.fields(FIELD_ACCEPTANCE_NOTES) = Null
    rs.Update
    RememberCachedScanRow rs, scanCache, CANDIDATE_STATUS_SCAN_ERROR, ScanSignature, foundExistingError
    Exit Sub
ErrHandler:
    'Do not stop candidate discovery because the error-log row itself failed.
End Sub

Private Function CandidateDiscoveryErrorRowCountInDb(ByVal resultConfigDb As DAO.Database) As Long
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    If resultConfigDb Is Nothing Then Exit Function
    If SharedDao_DatabaseTableExistsSafe(resultConfigDb, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates")) = False Then Exit Function
    Set rs = resultConfigDb.OpenRecordset("SELECT Count(*) AS ErrorCount FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates") & "] WHERE " & SQL_FIELD_CANDIDATE_STATUS & "=" & SharedDao_SqlQuote(CANDIDATE_STATUS_SCAN_ERROR), dbOpenSnapshot)
    CandidateDiscoveryErrorRowCountInDb = CLng(rs!errorCount)
ExitHere:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Exit Function
ErrHandler:
    CandidateDiscoveryErrorRowCountInDb = 0
    Resume ExitHere
End Function
Private Function SafeDatabaseNameForError(ByVal db As DAO.Database) As String
    On Error GoTo Fallback
    SafeDatabaseNameForError = SharedDao_SafeDatabaseName(db)
    Exit Function
Fallback:
    SafeDatabaseNameForError = "(database name unavailable: " & Err.Description & ")"
End Function

Private Function SafeObjectNameForError(ByVal targetObject As Object) As String
    On Error GoTo Fallback
    SafeObjectNameForError = targetObject.Name
    Exit Function
Fallback:
    SafeObjectNameForError = "(object name unavailable: " & Err.Description & ")"
End Function

Private Function CandidateIsBuiltInSystemMacroTarget(ByVal isCoreUtilityModule As Boolean, ByVal isGeneratedUiUtilityModule As Boolean, ByVal systemMacroWrapperNames As Object, ByVal procName As String) As Boolean
    Dim procKey As String
    Dim isCachedWrapper As Boolean
    procKey = LCase$(Trim$(procName))
    isCachedWrapper = WrapperDictionaryContains(systemMacroWrapperNames, procKey)
    If isCoreUtilityModule Then
        If isCachedWrapper Then
            CandidateIsBuiltInSystemMacroTarget = True
            Exit Function
        End If
    End If
    If isGeneratedUiUtilityModule Then
        If isCachedWrapper Then
            CandidateIsBuiltInSystemMacroTarget = True
            Exit Function
        End If
        If Left$(procKey, 8) = "chatgpt_" Then
            CandidateIsBuiltInSystemMacroTarget = True
            Exit Function
        End If
        Select Case procKey
            Case "buildgeneratedforms", "deletegeneratedforms", "buildribbon", "ribbonopenmanagedobject"
                CandidateIsBuiltInSystemMacroTarget = True
                Exit Function
        End Select
    End If
End Function

Private Function BuildSystemMacroWrapperNameDictionary(ByVal systemMacroCatalogRows As Collection) As Object
    'Consumes the one catalog materialized by ScanCandidatesCore; do not rebuild canonical System Macro metadata here.
    On Error GoTo ErrHandler
    Dim names As Object
    Dim rowData As Variant
    Dim codeDbRef As DAO.Database
    Dim rs As DAO.Recordset
    Dim stageText As String
    Dim savedErrDescription As String
    ' INTENTIONAL KEYED LOOKUP: names provides keyed membership/lookup over source rows or state already read by BuildSystemMacroWrapperNameDictionary; it is a derived access path, not a second authority.
    Set names = CreateObject("Scripting.Dictionary")
    names.CompareMode = 1
    stageText = "system macro catalog"
    For Each rowData In systemMacroCatalogRows
        PumpSetupResponsiveness
        AddWrapperReferenceToDictionary names, SystemMacroCatalogRunCodeFunctionName(rowData)
    Next rowData
    stageText = "managed-object action payloads"
    If CodeDbTableExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectActionPayloads")) Then
        Set codeDbRef = ADCDb()
        Set rs = codeDbRef.OpenRecordset("SELECT " & SQL_FIELD_RUN_CODE_WRAPPER_NAME & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectActionPayloads") & "] WHERE Len(Nz(" & SQL_FIELD_RUN_CODE_WRAPPER_NAME & ",''))>0;", dbOpenSnapshot)
        Do While rs.EOF = False
            PumpSetupResponsiveness
            AddWrapperReferenceToDictionary names, CStr(rs.fields(FIELD_RUN_CODE_WRAPPER_NAME).value)
            rs.MoveNext
        Loop
    End If
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set BuildSystemMacroWrapperNameDictionary = names
    Exit Function
ErrHandler:
    savedErrDescription = Err.Description
    On Error Resume Next
    CoreAddOperationNote "Candidate scan could not completely preload built-in wrapper names from " & stageText & ". Cached names already loaded remain active. Err.Description: " & savedErrDescription
    If Not rs Is Nothing Then rs.Close
    Set BuildSystemMacroWrapperNameDictionary = names
    Exit Function
End Function

Private Sub RepairResultConfigDbForScan(ByVal resultConfigDatabaseRole As String)
    Dim resultConfigDb As DAO.Database
    If StrComp(resultConfigDatabaseRole, CONFIG_OWNER_DATABASE_ROLE_CODE_DB, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 5453, "RepairResultConfigDbForScan", "Candidate storage is CodeDb-owned. Scan HostDb or CodeDb sources into the CodeDb MC_Candidates table."
    EnsureSettingsReadyForRead
    Set resultConfigDb = ADCDb()
    RequireResultConfigTableReadyForScan resultConfigDb, "RepairResultConfigDbForScan"
End Sub
Private Function ConfigActiveMacroTitleExists(ByVal macroTitleText As String) As Boolean
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    If Len(macroTitleText) = 0 Then Exit Function
    Set rs = ADCDb().OpenRecordset("SELECT " & SQL_FIELD_ID & ", " & SQL_FIELD_MACRO_TITLE & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "] WHERE " & SQL_FIELD_ACTIVE & "<>0", dbOpenSnapshot)
    Do While Not rs.EOF
        PumpSetupResponsiveness
        If StrComp(SharedText_Nz(rs.fields(FIELD_MACRO_TITLE).value), macroTitleText, vbTextCompare) = 0 Then
            ConfigActiveMacroTitleExists = True
            Exit Do
        End If
        rs.MoveNext
    Loop
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
    Err.Raise savedErrNumber, "ConfigActiveMacroTitleExists", "Could not determine whether an active configured macro title already exists." & vbCrLf & "Macro title: " & macroTitleText & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function ExistingMacroObjectShouldBeSkipped(ByVal macroName As String, ByVal scanCache As Object) As Boolean
    If scanCache Is Nothing Then Exit Function
    If scanCache.Exists(CandidateScanExactIdentityCacheKey("SYSTEMMACROTITLE", macroName)) Then
        ExistingMacroObjectShouldBeSkipped = True
        Exit Function
    End If
    If scanCache.Exists(CandidateScanExactIdentityCacheKey("CFGMACROTITLE", macroName)) Then ExistingMacroObjectShouldBeSkipped = True
End Function

Private Sub RequireResultConfigTableReadyForScan(ByVal resultConfigDb As DAO.Database, ByVal sourceName As String)
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim tdf As DAO.tableDef
    Dim fieldNames As Variant
    Dim fieldName As Variant
    Dim candidateTableName As String
    Dim configTableName As String
    If resultConfigDb Is Nothing Then Err.Raise vbObjectError + 5480, sourceName, "ResultConfigDb reference is Nothing."
    candidateTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates")
    configTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")
    If SharedDao_DatabaseTableExistsSafe(resultConfigDb, candidateTableName) = False Then Err.Raise vbObjectError + 5481, sourceName, "Access Development Console MC_Candidates table is missing. Run ADC Setup before candidate scanning." & vbCrLf & "MC_Candidates table: " & candidateTableName
    If SharedDao_DatabaseTableExistsSafe(resultConfigDb, configTableName) = False Then Err.Raise vbObjectError + 5483, sourceName, "Access Development Console MC_Records configuration table is missing. Run ADC Setup before candidate scanning." & vbCrLf & "MC_Records table: " & configTableName
    resultConfigDb.TableDefs.Refresh
    Set tdf = resultConfigDb.TableDefs(candidateTableName)
    Set rs = resultConfigDb.OpenRecordset("SELECT TOP 1 * FROM [" & candidateTableName & "]", dbOpenSnapshot)
    rs.Close
    Set rs = Nothing
    fieldNames = Split(CandidateFieldOrderText(), "|")
    For Each fieldName In fieldNames
        PumpSetupResponsiveness
        If SharedDao_TableDefFieldExistsSafe(tdf, CStr(fieldName)) = False Then Err.Raise vbObjectError + 5482, sourceName, "Access Development Console MC_Candidates table is missing a required current-design field. Run ADC Setup before candidate scanning." & vbCrLf & "MC_Candidates table: " & candidateTableName & vbCrLf & "Missing field: " & CStr(fieldName)
    Next fieldName
    Set tdf = resultConfigDb.TableDefs(configTableName)
    fieldNames = Array(FIELD_ID, FIELD_ACTIVE, FIELD_MACRO_KEY, FIELD_MACRO_TITLE, FIELD_CONFIG_OWNER_DATABASE_ROLE, FIELD_TARGET_PROCEDURE_NAME, FIELD_TARGET_PROCEDURE_MODULE_NAME, FIELD_TARGET_OBJECT_NAME)
    For Each fieldName In fieldNames
        PumpSetupResponsiveness
        If SharedDao_TableDefFieldExistsSafe(tdf, CStr(fieldName)) = False Then Err.Raise vbObjectError + 5484, sourceName, "Access Development Console MC_Records table is missing a configuration field required by candidate duplicate-exclusion checks. Run ADC Setup before candidate scanning." & vbCrLf & "MC_Records table: " & configTableName & vbCrLf & "Missing field: " & CStr(fieldName)
    Next fieldName
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
    Err.Raise savedErrNumber, "RequireResultConfigTableReadyForScan", "Could not verify the ResultConfigDb MC_Candidates table for candidate scanning." & vbCrLf & "ResultConfigDb: " & SafeDatabaseNameForError(resultConfigDb) & vbCrLf & "MC_Candidates table: " & candidateTableName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub
Private Function CachedScanRowBookmark(ByVal scanCache As Object, ByVal RowKind As String, ByVal ScanSignature As String) As Variant
    Dim cacheKey As String
    If scanCache Is Nothing Then Exit Function
    cacheKey = CandidateScanCacheKey("ROW", RowKind, ScanSignature)
    If scanCache.Exists(cacheKey) Then CachedScanRowBookmark = scanCache(cacheKey)
End Function

Private Function PositionCachedScanRow(ByVal rs As DAO.Recordset, ByVal cachedBookmark As Variant, ByVal expectedRowKind As String, ByVal expectedScanSignature As String) As Boolean
    If IsEmpty(cachedBookmark) Then Exit Function
    rs.Bookmark = cachedBookmark
    If StrComp(SharedText_Nz(rs.fields(FIELD_SCAN_SIGNATURE).value), expectedScanSignature, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 5464, "PositionCachedScanRow", "The cached candidate scan signature changed during candidate scanning. Row ID: " & CStr(SharedValue_NzLong(rs.fields(FIELD_ID).value))
    If StrComp(expectedRowKind, CANDIDATE_STATUS_SCAN_ERROR, vbTextCompare) = 0 Then
        If StrComp(SharedText_Nz(rs.fields(FIELD_CANDIDATE_STATUS).value), CANDIDATE_STATUS_SCAN_ERROR, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 5463, "PositionCachedScanRow", "The cached discovery-error candidate changed status during candidate scanning. Row ID: " & CStr(SharedValue_NzLong(rs.fields(FIELD_ID).value))
    End If
    PositionCachedScanRow = True
End Function
Private Sub RememberCachedScanRow(ByVal rs As DAO.Recordset, ByVal scanCache As Object, ByVal RowKind As String, ByVal ScanSignature As String, ByVal updatedExistingRow As Boolean)
    Dim rowID As Long
    Dim cacheKey As String
    If updatedExistingRow = False Then rs.Bookmark = rs.LastModified
    rowID = SharedValue_NzLong(rs.fields(FIELD_ID).value)
    If rowID <= 0 Then Err.Raise vbObjectError + 5465, "RememberCachedScanRow", "The inserted or updated candidate row does not have a valid ID."
    cacheKey = CandidateScanCacheKey("ROW", RowKind, ScanSignature)
    If scanCache.Exists(cacheKey) Then
        scanCache(cacheKey) = rs.Bookmark
    Else
        scanCache.Add cacheKey, rs.Bookmark
    End If
End Sub

Private Sub SetScanFieldIfAvailable(ByVal rs As DAO.Recordset, ByVal scanCache As Object, ByVal fieldName As String, ByVal fieldValue As Variant)
    Dim cacheKey As String
    Dim fieldOrdinal As Long
    cacheKey = CandidateScanCacheKey("FIELD", fieldName)
    If scanCache.Exists(cacheKey) = False Then Exit Sub
    fieldOrdinal = CLng(scanCache(cacheKey))
    rs.fields(fieldOrdinal).value = fieldValue
End Sub

Private Sub SetScanFieldIfChanged(ByVal rs As DAO.Recordset, ByVal scanCache As Object, ByVal fieldName As String, ByVal fieldValue As Variant, ByRef rowChanged As Boolean, Optional ByVal fieldRequired As Boolean = True)
    Dim cacheKey As String
    Dim fieldOrdinal As Long
    cacheKey = CandidateScanCacheKey("FIELD", fieldName)
    If scanCache.Exists(cacheKey) = False Then
        If fieldRequired Then Err.Raise vbObjectError + 5466, "SetScanFieldIfChanged", "The required MC_Candidates field is unavailable during candidate scanning. Field: " & fieldName
        Exit Sub
    End If
    fieldOrdinal = CLng(scanCache(cacheKey))
    If ScanValuesEqual(rs.fields(fieldOrdinal).value, fieldValue) Then Exit Sub
    rs.fields(fieldOrdinal).value = fieldValue
    rowChanged = True
End Sub

Private Function ScanValuesEqual(ByVal currentValue As Variant, ByVal expectedValue As Variant) As Boolean
    On Error GoTo CompareAsText
    If IsNull(currentValue) And IsNull(expectedValue) Then
        ScanValuesEqual = True
        Exit Function
    End If
    If IsNull(currentValue) Xor IsNull(expectedValue) Then Exit Function
    If VarType(currentValue) = vbDate Or VarType(expectedValue) = vbDate Then
        ScanValuesEqual = (CDbl(CDate(currentValue)) = CDbl(CDate(expectedValue)))
        Exit Function
    End If
    If VarType(currentValue) = vbBoolean Or VarType(expectedValue) = vbBoolean Then
        ScanValuesEqual = (CBool(currentValue) = CBool(expectedValue))
        Exit Function
    End If
    If IsNumeric(currentValue) And IsNumeric(expectedValue) Then
        ScanValuesEqual = (CDbl(currentValue) = CDbl(expectedValue))
        Exit Function
    End If
CompareAsText:
    Err.Clear
    ScanValuesEqual = (StrComp(CStr(currentValue), CStr(expectedValue), vbBinaryCompare) = 0)
End Function

Private Sub AppendScanSourceSummary(ByRef summaryText As String, ByVal roleText As String)
    If Len(roleText) = 0 Then Exit Sub
    If InStr(1, ", " & summaryText & ", ", ", " & roleText & ", ", vbTextCompare) > 0 Then Exit Sub
    If Len(Trim$(summaryText)) > 0 Then summaryText = summaryText & ", "
    summaryText = summaryText & roleText
End Sub

Private Function CandidateComponentIsGeneratedUiUtility(ByVal vbComponent As Object, ByVal moduleTopText As String) As Boolean
    On Error GoTo ErrHandler
    Dim componentName As String
    componentName = LCase$(Trim$(SafeObjectText(vbComponent, "Name")))
    If componentName = "x__util_macro_creator_ui" Then
        CandidateComponentIsGeneratedUiUtility = True
        Exit Function
    End If
    If InStr(1, moduleTopText, "standard module: access development console ui utility", vbTextCompare) > 0 Then
        CandidateComponentIsGeneratedUiUtility = True
        Exit Function
    End If
    If InStr(1, moduleTopText, "vb_name: x__util_macro_creator_ui", vbTextCompare) > 0 Then CandidateComponentIsGeneratedUiUtility = True
    Exit Function
ErrHandler:
    CandidateComponentIsGeneratedUiUtility = False
End Function

Private Function CandidateComponentIsThisUtility(ByVal vbComponent As Object, ByVal moduleTopText As String) As Boolean
    On Error GoTo ErrHandler
    Dim componentName As String
    componentName = LCase$(Trim$(SafeObjectText(vbComponent, "Name")))
    If componentName = "x__util_macro_creator" Then
        CandidateComponentIsThisUtility = True
        Exit Function
    End If
    If InStr(1, moduleTopText, "standard module: access development console utility", vbTextCompare) > 0 Then
        CandidateComponentIsThisUtility = True
        Exit Function
    End If
    If InStr(1, moduleTopText, "vb_name: x__util_macro_creator", vbTextCompare) > 0 Then CandidateComponentIsThisUtility = True
    Exit Function
ErrHandler:
    CandidateComponentIsThisUtility = False
End Function

Private Function AddObjectCandidatesFromScanTargetDb(ByVal scanTargetDb As DAO.Database, ByVal CandidateOwnerDatabaseRole As String, ByVal rs As DAO.Recordset, ByVal scanCache As Object, ByVal scanBatchID As String, ByRef tableCandidateCount As Long, ByRef queryCandidateCount As Long, ByRef formCandidateCount As Long, ByRef reportCandidateCount As Long, ByRef macroCandidateCount As Long, ByRef skippedObjectCount As Long, ByRef tableScanElapsedSeconds As Double, ByRef queryScanElapsedSeconds As Double, ByRef queryNameFilterElapsedSeconds As Double, ByRef queryHiddenSystemFilterElapsedSeconds As Double, ByRef queryCandidateProcessingElapsedSeconds As Double, ByRef formScanElapsedSeconds As Double, ByRef reportScanElapsedSeconds As Double, ByRef macroScanElapsedSeconds As Double) As Long
    On Error GoTo ErrHandler
    Dim tdf As DAO.tableDef
    Dim qdf As DAO.queryDef
    Dim doc As DAO.Document
    Dim addedCount As Long
    Dim categoryStartedTimer As Single
    Dim operationStartedTimer As Single
    Dim queryNameSkippable As Boolean
    Dim queryHiddenOrSystem As Boolean
    scanTargetDb.TableDefs.Refresh
    scanTargetDb.QueryDefs.Refresh
    CorePulseProgress "Candidate scan [" & CandidateOwnerDatabaseRole & "] | tables | source objects=" & CStr(scanTargetDb.TableDefs.count) & " | candidates=" & CStr(addedCount) & " | skipped=" & CStr(skippedObjectCount)
    categoryStartedTimer = Timer
    For Each tdf In scanTargetDb.TableDefs
        PumpSetupResponsiveness
        If DaoObjectIsHiddenOrSystem(tdf.Attributes) Or CandidateObjectNameIsSkippable(tdf.Name) Then
            skippedObjectCount = skippedObjectCount + 1
        ElseIf AddAccessObjectCandidateRow(rs, scanCache, tdf.Name, ADCAT_OBJECT_TYPE_TABLE, ACTION_OPENTABLE, CandidateOwnerDatabaseRole, scanTargetDb.Name, scanTargetDb, scanBatchID) Then
            tableCandidateCount = tableCandidateCount + 1
            addedCount = addedCount + 1
            PulseProgressEvery addedCount, 25, "Candidate scan [" & CandidateOwnerDatabaseRole & "] | table " & tdf.Name & " | object candidates=" & CStr(addedCount) & " | skipped=" & CStr(skippedObjectCount), 0, "CandidateObjectScan", "CandidateDiscovery"
        End If
    Next tdf
    tableScanElapsedSeconds = tableScanElapsedSeconds + ElapsedSecondsFromTimer(categoryStartedTimer)
    CorePulseProgress "Candidate scan [" & CandidateOwnerDatabaseRole & "] | queries | source objects=" & CStr(scanTargetDb.QueryDefs.count) & " | candidates=" & CStr(addedCount) & " | skipped=" & CStr(skippedObjectCount)
    categoryStartedTimer = Timer
    For Each qdf In scanTargetDb.QueryDefs
        PumpSetupResponsiveness
        operationStartedTimer = Timer
        queryNameSkippable = CandidateObjectNameIsSkippable(qdf.Name)
        queryNameFilterElapsedSeconds = queryNameFilterElapsedSeconds + ElapsedSecondsFromTimer(operationStartedTimer)
        If queryNameSkippable Then
            skippedObjectCount = skippedObjectCount + 1
        Else
            operationStartedTimer = Timer
            queryHiddenOrSystem = QueryDefIsHiddenOrSystem(qdf, scanTargetDb)
            queryHiddenSystemFilterElapsedSeconds = queryHiddenSystemFilterElapsedSeconds + ElapsedSecondsFromTimer(operationStartedTimer)
            If queryHiddenOrSystem Then
                skippedObjectCount = skippedObjectCount + 1
            Else
                operationStartedTimer = Timer
                If AddAccessObjectCandidateRow(rs, scanCache, qdf.Name, ADCAT_OBJECT_TYPE_QUERY, ACTION_OPENQUERY, CandidateOwnerDatabaseRole, scanTargetDb.Name, scanTargetDb, scanBatchID) Then
                    queryCandidateCount = queryCandidateCount + 1
                    addedCount = addedCount + 1
                    PulseProgressEvery addedCount, 25, "Candidate scan [" & CandidateOwnerDatabaseRole & "] | query " & qdf.Name & " | object candidates=" & CStr(addedCount) & " | skipped=" & CStr(skippedObjectCount), 0, "CandidateObjectScan", "CandidateDiscovery"
                End If
                queryCandidateProcessingElapsedSeconds = queryCandidateProcessingElapsedSeconds + ElapsedSecondsFromTimer(operationStartedTimer)
            End If
        End If
    Next qdf
    queryScanElapsedSeconds = queryScanElapsedSeconds + ElapsedSecondsFromTimer(categoryStartedTimer)
    CorePulseProgress "Candidate scan [" & CandidateOwnerDatabaseRole & "] | forms | source objects=" & CStr(scanTargetDb.Containers("Forms").Documents.count) & " | candidates=" & CStr(addedCount) & " | skipped=" & CStr(skippedObjectCount)
    categoryStartedTimer = Timer
    For Each doc In scanTargetDb.Containers("Forms").Documents
        PumpSetupResponsiveness
        If DaoDocumentIsHiddenOrSystem(doc) Or CandidateObjectNameIsSkippable(doc.Name) Then
            skippedObjectCount = skippedObjectCount + 1
        ElseIf AddAccessObjectCandidateRow(rs, scanCache, doc.Name, ADCAT_OBJECT_TYPE_FORM, ACTION_OPENFORM, CandidateOwnerDatabaseRole, scanTargetDb.Name, scanTargetDb, scanBatchID) Then
            formCandidateCount = formCandidateCount + 1
            addedCount = addedCount + 1
            PulseProgressEvery addedCount, 25, "Candidate scan [" & CandidateOwnerDatabaseRole & "] | form " & doc.Name & " | object candidates=" & CStr(addedCount) & " | skipped=" & CStr(skippedObjectCount), 0, "CandidateObjectScan", "CandidateDiscovery"
        End If
    Next doc
    formScanElapsedSeconds = formScanElapsedSeconds + ElapsedSecondsFromTimer(categoryStartedTimer)
    CorePulseProgress "Candidate scan [" & CandidateOwnerDatabaseRole & "] | reports | source objects=" & CStr(scanTargetDb.Containers("Reports").Documents.count) & " | candidates=" & CStr(addedCount) & " | skipped=" & CStr(skippedObjectCount)
    categoryStartedTimer = Timer
    For Each doc In scanTargetDb.Containers("Reports").Documents
        PumpSetupResponsiveness
        If DaoDocumentIsHiddenOrSystem(doc) Or CandidateObjectNameIsSkippable(doc.Name) Then
            skippedObjectCount = skippedObjectCount + 1
        ElseIf AddAccessObjectCandidateRow(rs, scanCache, doc.Name, ADCAT_OBJECT_TYPE_REPORT, ACTION_OPENREPORT, CandidateOwnerDatabaseRole, scanTargetDb.Name, scanTargetDb, scanBatchID) Then
            reportCandidateCount = reportCandidateCount + 1
            addedCount = addedCount + 1
            PulseProgressEvery addedCount, 25, "Candidate scan [" & CandidateOwnerDatabaseRole & "] | report " & doc.Name & " | object candidates=" & CStr(addedCount) & " | skipped=" & CStr(skippedObjectCount), 0, "CandidateObjectScan", "CandidateDiscovery"
        End If
    Next doc
    reportScanElapsedSeconds = reportScanElapsedSeconds + ElapsedSecondsFromTimer(categoryStartedTimer)
    CorePulseProgress "Candidate scan [" & CandidateOwnerDatabaseRole & "] | macros | source objects=" & CStr(scanTargetDb.Containers("Scripts").Documents.count) & " | candidates=" & CStr(addedCount) & " | skipped=" & CStr(skippedObjectCount)
    categoryStartedTimer = Timer
    For Each doc In scanTargetDb.Containers("Scripts").Documents
        PumpSetupResponsiveness
        If DaoDocumentIsHiddenOrSystem(doc) Or CandidateObjectNameIsSkippable(doc.Name) Or ExistingMacroObjectShouldBeSkipped(doc.Name, scanCache) Then
            skippedObjectCount = skippedObjectCount + 1
        ElseIf AddAccessObjectCandidateRow(rs, scanCache, doc.Name, ADCAT_OBJECT_TYPE_MACRO, ACTION_RUNMACRO, CandidateOwnerDatabaseRole, scanTargetDb.Name, scanTargetDb, scanBatchID) Then
            macroCandidateCount = macroCandidateCount + 1
            addedCount = addedCount + 1
            PulseProgressEvery addedCount, 25, "Candidate scan [" & CandidateOwnerDatabaseRole & "] | macro " & doc.Name & " | object candidates=" & CStr(addedCount) & " | skipped=" & CStr(skippedObjectCount), 0, "CandidateObjectScan", "CandidateDiscovery"
        End If
    Next doc
    macroScanElapsedSeconds = macroScanElapsedSeconds + ElapsedSecondsFromTimer(categoryStartedTimer)
    CorePulseProgress "Candidate scan [" & CandidateOwnerDatabaseRole & "] | Access objects complete | candidates=" & CStr(addedCount) & " | tables=" & CStr(tableCandidateCount) & " | queries=" & CStr(queryCandidateCount) & " | forms=" & CStr(formCandidateCount) & " | reports=" & CStr(reportCandidateCount) & " | macros=" & CStr(macroCandidateCount) & " | skipped=" & CStr(skippedObjectCount)
    AddObjectCandidatesFromScanTargetDb = addedCount
    Exit Function
ErrHandler:
    CoreAddOperationNote "Access-object candidate scan stopped early for " & CandidateOwnerDatabaseRole & ". Err.Number: " & CStr(Err.Number) & "; Err.Description: " & SharedMessage_OneLine(Err.Description, 250)
    AddObjectCandidatesFromScanTargetDb = addedCount
End Function

Private Function CandidateObjectNameIsSkippable(ByVal objectName As String) As Boolean
    Dim nameText As String
    nameText = objectName
    If Len(nameText) = 0 Then
        CandidateObjectNameIsSkippable = True
        Exit Function
    End If
    If Left$(nameText, 1) = "~" Then
        CandidateObjectNameIsSkippable = True
        Exit Function
    End If
    If Left$(nameText, 4) = "MSys" Then
        CandidateObjectNameIsSkippable = True
        Exit Function
    End If
    If StrComp(nameText, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config"), vbTextCompare) = 0 Or StrComp(nameText, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates"), vbTextCompare) = 0 Or StrComp(nameText, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID), vbTextCompare) = 0 Or StrComp(nameText, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReview"), vbTextCompare) = 0 Then CandidateObjectNameIsSkippable = True
End Function

Private Function AddAccessObjectCandidateRow(ByVal rs As DAO.Recordset, ByVal scanCache As Object, ByVal objectName As String, ByVal objectKind As String, ByVal MacroAction As String, ByVal CandidateOwnerDatabaseRole As String, ByVal scanTargetDbName As String, ByVal scanTargetDb As DAO.Database, ByVal scanBatchID As String) As Boolean
    On Error GoTo ErrHandler
    Dim candidateMacroKeyText As String
    Dim ScanSignature As String
    Dim foundExistingCandidate As Boolean
    Dim cachedBookmark As Variant
    Dim descriptionText As String
    Dim macroTitleText As String
    Dim rowChanged As Boolean
    Dim ScanTimestamp As Date
    If Len(objectName) = 0 Or Len(objectKind) = 0 Or Len(MacroAction) = 0 Then Exit Function
    candidateMacroKeyText = CandidateAccessObjectMacroKey(CandidateOwnerDatabaseRole, objectKind, objectName)
    ScanSignature = AccessObjectCandidateScanSignature(CandidateOwnerDatabaseRole, scanTargetDbName, objectKind, objectName, MacroAction)
    macroTitleText = AccessObjectCandidateMacroTitle(objectKind, objectName)
    If CandidateScanCacheHasConfiguredAccessObject(scanCache, MacroAction, objectName) Then Exit Function
    descriptionText = "Candidate discovered from existing visible Access " & LCase$(objectKind) & " object [" & objectName & "]. Logical macro action: " & MacroAction & "."
    ScanTimestamp = Now
    cachedBookmark = CachedScanRowBookmark(scanCache, ROW_KIND_CANDIDATE, ScanSignature)
    foundExistingCandidate = PositionCachedScanRow(rs, cachedBookmark, ROW_KIND_CANDIDATE, ScanSignature)
    If foundExistingCandidate Then
        rs.Edit
    Else
        rs.AddNew
    End If
    SetScanFieldIfChanged rs, scanCache, FIELD_CANDIDATE_STATUS, CANDIDATE_STATUS_NEW, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_MACRO_KEY, Null, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_CANDIDATE_MACRO_KEY, candidateMacroKeyText, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_SCAN_SIGNATURE, ScanSignature, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_LAST_SCAN_BATCH_ID, scanBatchID, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_LAST_SCANNED_ON, ScanTimestamp, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_REJECTED_ON, Null, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_MACRO_TITLE, macroTitleText, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_MACRO_ACTION, MacroAction, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_CONFIG_OWNER_DATABASE_ROLE, CandidateOwnerDatabaseRole, rowChanged, False
    SetScanFieldIfChanged rs, scanCache, FIELD_TARGET_PROCEDURE_NAME, Null, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_TARGET_PROCEDURE_KIND, Null, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_TARGET_ARGUMENTS, Null, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_RESULT_HANDLING, Null, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_EXPECTED_RETURN_TYPE, Null, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_CANDIDATE_OWNER_DATABASE_ROLE, CandidateOwnerDatabaseRole, rowChanged, False
    SetScanFieldIfChanged rs, scanCache, FIELD_CANDIDATE_SCAN_TARGET_DB_NAME, DatabaseBaseNameFromPath(scanTargetDbName), rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_CANDIDATE_SCAN_TARGET_DB_PATH, scanTargetDbName, rowChanged, False
    SetScanFieldIfChanged rs, scanCache, FIELD_CANDIDATE_VB_PROJECT_NAME, Null, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_CANDIDATE_MODULE_NAME, objectKind, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_GENERATION_TARGET_DB_NAME, DatabaseBaseNameFromPath(scanTargetDb.Name), rowChanged, False
    SetScanFieldIfChanged rs, scanCache, FIELD_GENERATION_TARGET_DB_PATH, scanTargetDb.Name, rowChanged, False
    SetScanFieldIfChanged rs, scanCache, FIELD_TARGET_REFERENCE_NAME, Null, rowChanged, False
    SetScanFieldIfChanged rs, scanCache, FIELD_TARGET_VB_PROJECT_NAME, Null, rowChanged, False
    SetScanFieldIfChanged rs, scanCache, FIELD_TARGET_PROCEDURE_MODULE_NAME, Null, rowChanged, False
    SetScanFieldIfChanged rs, scanCache, FIELD_MODULE_HEADER_FOUND, False, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_MODULE_HEADER_MALFORMED, False, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_MODULE_HEADER_TITLE, Null, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_MODULE_HEADER_PURPOSE, Null, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_MODULE_HEADER_ENTRY_POINTS, Null, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_PROCEDURE_LISTED_IN_ENTRY_POINTS, False, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_ENTRY_POINT_MATCH_CONFIDENCE, Null, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_MODULE_HEADER_SCAN_NOTES, Null, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_CANDIDATE_PROCEDURE_DECLARATION, Null, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_CANDIDATE_ARGUMENT_SIGNATURE, Null, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_HAS_REQUIRED_ARGUMENTS, False, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_REQUIRED_ARGUMENT_COUNT, 0, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_OPTIONAL_ARGUMENT_COUNT, 0, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_DISCOVERED_ON, ScanTimestamp, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_TARGET_OBJECT_NAME, objectName, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_NAVIGATION_PANE_GROUP, BuildNavigationPaneGroupName(DEFAULT_GENERATED_MACROS_NAVIGATION_PANE_GROUP_LOGICAL_NAME, False), rowChanged, False
    SetScanFieldIfChanged rs, scanCache, FIELD_CONFIRMATION_MODE, IIf(StrComp(MacroAction, ACTION_OPENFORM, vbTextCompare) = 0, CONFIRM_NEVER, CONFIRM_ALWAYS), rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_ACTION_DESCRIPTION, Shared_String_ConvertToNullIfBlank(descriptionText), rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER, AccessObjectCandidateSortOrder(objectKind), rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_SELECTED_FOR_ACCEPTANCE, False, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_ACCEPTANCE_STATUS, Null, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_ACCEPTED_ON, Null, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_ACCEPTED_CONFIGURATION_ID, Null, rowChanged
    SetScanFieldIfChanged rs, scanCache, FIELD_ACCEPTANCE_NOTES, Null, rowChanged
    rs.Update
    RememberCachedScanRow rs, scanCache, ROW_KIND_CANDIDATE, ScanSignature, foundExistingCandidate
    AddAccessObjectCandidateRow = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "AddAccessObjectCandidateRow", "Could not add or refresh Access object access development console candidate row." & vbCrLf & "Object kind: " & objectKind & vbCrLf & "Object name: " & objectName & vbCrLf & "MacroAction: " & MacroAction & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Sub MarkPriorScanCandidateRowsStaleInDb(ByVal resultConfigDb As DAO.Database, ByVal scanBatchID As String, ByVal scanCodeDb As Boolean, ByVal scanHostDb As Boolean)
    'The current ScanBatchID is authoritative for candidate freshness in the scanned database-role scope; prior-scan candidate rows are marked stale before current results are written.
    On Error GoTo ErrHandler
    Dim roleFilterText As String
    Dim sqlText As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    If resultConfigDb Is Nothing Then Err.Raise vbObjectError + 5405, "MarkPriorScanCandidateRowsStaleInDb", "ResultConfigDb reference is Nothing."
    If scanCodeDb = False And scanHostDb = False Then Err.Raise vbObjectError + 5406, "MarkPriorScanCandidateRowsStaleInDb", "At least one scanned database role is required."
    If scanCodeDb Xor scanHostDb Then
        If scanCodeDb Then
            roleFilterText = " AND Nz([" & FIELD_CANDIDATE_OWNER_DATABASE_ROLE & "],'')=" & SharedDao_SqlQuote(CONFIG_OWNER_DATABASE_ROLE_CODE_DB)
        Else
            roleFilterText = " AND Nz([" & FIELD_CANDIDATE_OWNER_DATABASE_ROLE & "],'')=" & SharedDao_SqlQuote(CONFIG_OWNER_DATABASE_ROLE_HOST_DB)
        End If
    End If
    sqlText = "UPDATE [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates") & "] SET " & SQL_FIELD_CANDIDATE_STATUS & "=" & SharedDao_SqlQuote(CANDIDATE_STATUS_STALE) & " WHERE 1=1" & roleFilterText & " AND Nz(" & SQL_FIELD_LAST_SCAN_BATCH_ID & ",'')<>" & SharedDao_SqlQuote(scanBatchID)
    resultConfigDb.Execute sqlText, dbFailOnError
    InvalidateDuplicateReviewCache
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "MarkPriorScanCandidateRowsStaleInDb", "Could not mark prior-scan candidate rows stale before the scan." & vbCrLf & "ResultConfigDb: " & SafeDatabaseNameForError(resultConfigDb) & vbCrLf & "MC_Candidates table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates") & vbCrLf & "Scan CodeDb: " & CStr(scanCodeDb) & vbCrLf & "Scan HostDb: " & CStr(scanHostDb) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub
Private Sub AddWrapperReferenceToDictionary(ByVal names As Object, ByVal runCodeReference As String)
    Dim ProcedureName As String
    If names Is Nothing Then Exit Sub
    runCodeReference = Trim$(runCodeReference)
    If Len(runCodeReference) = 0 Then Exit Sub
    ProcedureName = ProcedureNameFromRunCodeExpression(runCodeReference)
    If Len(ProcedureName) = 0 Then ProcedureName = runCodeReference
    ProcedureName = LCase$(Trim$(ProcedureName))
    If Len(ProcedureName) = 0 Then Exit Sub
    If names.Exists(ProcedureName) = False Then names.Add ProcedureName, True
End Sub

Private Function WrapperDictionaryContains(ByVal names As Object, ByVal procName As String) As Boolean
    On Error GoTo ExitHere
    If names Is Nothing Then Exit Function
    If Len(procName) = 0 Then Exit Function
    WrapperDictionaryContains = names.Exists(procName)
ExitHere:
End Function

Private Function DaoObjectIsHiddenOrSystem(ByVal attributesValue As Long) As Boolean
    On Error GoTo ErrHandler
    If (attributesValue And dbSystemObject) <> 0 Then
        DaoObjectIsHiddenOrSystem = True
        Exit Function
    End If
    If (attributesValue And dbHiddenObject) <> 0 Then
        DaoObjectIsHiddenOrSystem = True
        Exit Function
    End If
    Exit Function
ErrHandler:
    DaoObjectIsHiddenOrSystem = False
End Function

Private Function DaoDocumentIsHiddenOrSystem(ByVal documentObject As DAO.Document) As Boolean
    On Error GoTo ErrHandler
    Dim nameText As String
    nameText = documentObject.Name
    If Len(nameText) = 0 Then
        DaoDocumentIsHiddenOrSystem = True
        Exit Function
    End If
    If Left$(nameText, 1) = "~" Then
        DaoDocumentIsHiddenOrSystem = True
        Exit Function
    End If
    If Left$(nameText, 4) = "MSys" Then
        DaoDocumentIsHiddenOrSystem = True
        Exit Function
    End If
    DaoDocumentIsHiddenOrSystem = DaoPropertiesIndicateHiddenOrSystem(documentObject.properties)
    Exit Function
ErrHandler:
    DaoDocumentIsHiddenOrSystem = False
End Function

Private Function DaoPropertiesIndicateHiddenOrSystem(ByVal objectProperties As DAO.properties) As Boolean
    On Error GoTo ErrHandler
    Dim propertyObject As DAO.Property
    Dim propertyName As String
    For Each propertyObject In objectProperties
        PumpSetupResponsiveness
        propertyName = LCase$(Trim$(propertyObject.Name))
        If propertyName = "hidden" Or propertyName = "system" Then
            If IsNull(propertyObject.value) = False Then
                If CBool(propertyObject.value) Then
                    DaoPropertiesIndicateHiddenOrSystem = True
                    Exit Function
                End If
            End If
        End If
    Next propertyObject
    Exit Function
ErrHandler:
    DaoPropertiesIndicateHiddenOrSystem = False
End Function

Private Function QueryDefIsHiddenOrSystem(ByVal qdf As DAO.queryDef, ByVal queryDb As DAO.Database) As Boolean
    On Error GoTo PropertyFallback
    Dim queryName As String
    queryName = qdf.Name
    If Len(queryName) = 0 Then
        QueryDefIsHiddenOrSystem = True
        Exit Function
    End If
    If Left$(queryName, 1) = "~" Or Left$(queryName, 4) = "MSys" Then
        QueryDefIsHiddenOrSystem = True
        Exit Function
    End If
    If Not queryDb Is Nothing Then
        If DatabaseIsCurrentProject(queryDb) Then
            QueryDefIsHiddenOrSystem = Application.GetHiddenAttribute(acQuery, queryName)
            Exit Function
        End If
    End If
PropertyFallback:
    Err.Clear
    On Error GoTo ErrHandler
    QueryDefIsHiddenOrSystem = DaoPropertiesIndicateHiddenOrSystem(qdf.properties)
    Exit Function
ErrHandler:
    QueryDefIsHiddenOrSystem = False
End Function

' ===============================================================================================================================================================================
' CANDIDATE ACCEPTANCE PROCEDURES
' ===============================================================================================================================================================================

' Purpose: Determines whether accept meets the discovery/classification rules required to be processed as a candidate.
' Public role: Form-code/event-expression public entry point. Access form code or an event property can invoke this routine by name; its Public visibility is part of that UI contract.
' Parameters:
'   CandidateID (Input; Long): Primary-key ID of the candidate row being reviewed/accepted/rejected.
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub AcceptCandidate(ByVal CandidateID As Long, Optional ByVal showMessage As Boolean = True)
    'Form-facing single-candidate entry point. The current acceptance core remains the sole candidate-to-configuration creation authority.
    On Error GoTo ErrHandler
    Dim progressStarted As Boolean
    If CandidateID <= 0 Then Err.Raise vbObjectError + 9195, "AcceptCandidate", "CandidateID must be greater than zero."
    progressStarted = CoreBeginOperationProgressIfNeeded("Accept " & DEFAULT_UTILITY_SHORT_NAME & " candidate", 2)
    CoreUpdateSetupProgress "Check default Navigation Pane group"
    RequireDefaultCodeDbGeneratedMacrosNavigationPaneGroupForRoleContext "AcceptCandidate", False
    CoreUpdateSetupProgress "Accept candidate ID " & CStr(CandidateID)
    AcceptSelectedRunnerMacroCandidatesCore showMessage, CandidateID
    If progressStarted Then CoreCompleteSetupProgress
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If progressStarted Then CoreFailSetupProgress
    If savedErrNumber = CANDIDATE_ACCEPT_ROW_ERRORS_ERROR_NUMBER Then
        If showMessage Then
            CoreShowErrorMessage savedErrDescription, vbCritical, "Accept Access Development Console Candidate Completed With Errors", "AcceptCandidate", "CandidateAcceptance"
            Exit Sub
        End If
        Err.Raise savedErrNumber, savedErrSource, savedErrDescription
    End If
    If showMessage Then
        CoreShowErrorMessage "Could not accept the requested Access Development Console candidate." & vbCrLf & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, vbCritical, "Accept Access Development Console Candidate", "AcceptCandidate", "Core"
        Exit Sub
    End If
    Err.Raise savedErrNumber, "AcceptCandidate", "Could not accept the requested Access Development Console candidate." & vbCrLf & "CandidateID: " & CStr(CandidateID) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

' Purpose: Performs accept selected runner macro candidates by coordinating `CoreBeginOperationProgressIfNeeded` and `CoreUpdateSetupProgress` while preserving this module's existing validation and error semantics.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub AcceptSelectedRunnerMacroCandidates(Optional ByVal showMessage As Boolean = True)
    On Error GoTo ErrHandler
    Dim progressStarted As Boolean
    progressStarted = CoreBeginOperationProgressIfNeeded("Accept selected " & DEFAULT_UTILITY_SHORT_NAME & " candidates", 2)
    CoreUpdateSetupProgress "Check default Navigation Pane group"
    RequireDefaultCodeDbGeneratedMacrosNavigationPaneGroupForRoleContext "AcceptSelectedRunnerMacroCandidates", False
    CoreUpdateSetupProgress "Accept selected candidates"
    AcceptSelectedRunnerMacroCandidatesCore showMessage
    If progressStarted Then CoreCompleteSetupProgress
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If progressStarted Then CoreFailSetupProgress
    If savedErrNumber = CANDIDATE_ACCEPT_ROW_ERRORS_ERROR_NUMBER Then Err.Raise savedErrNumber, savedErrSource, savedErrDescription
    Err.Raise savedErrNumber, "AcceptSelectedRunnerMacroCandidates", "Could not accept selected Runner Macro candidates." & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub AcceptSelectedRunnerMacroCandidatesCore(Optional ByVal showMessage As Boolean = True, Optional ByVal CandidateID As Long = 0)
    On Error GoTo ErrHandler
    CoreResetOperationNotes
    Dim db As DAO.Database
    Dim ws As DAO.workspace
    Dim rsCandidate As DAO.Recordset
    Dim rsConfig As DAO.Recordset
    Dim cfg As TCandidateRecord
    Dim selectedCount As Long
    Dim acceptedCount As Long
    Dim acceptedNeedsArgumentsCount As Long
    Dim duplicateCount As Long
    Dim invalidCount As Long
    Dim alreadyAcceptedCount As Long
    Dim errorCount As Long
    Dim validationProblem As String
    Dim duplicateReason As String
    Dim AcceptanceStatus As String
    Dim AcceptanceNotes As String
    Dim currentID As Long
    Dim newConfigurationID As Long
    Dim activeValue As Boolean
    Dim rowErrNumber As Long
    Dim rowErrSource As String
    Dim rowErrDescription As String
    Dim summaryText As String
    Dim transactionStarted As Boolean
    Dim copyField As Variant
    Dim copyFields As Variant
    RequireResultConfigTableReadyForScan ADCDb(), "AcceptSelectedRunnerMacroCandidatesCore"
    RequireConfigTableForUtilityWorkflow "AcceptSelectedRunnerMacroCandidatesCore"
    Set db = ADCDb()
    Set ws = DBEngine.Workspaces(0)
    If CandidateID > 0 Then
        Set rsCandidate = db.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates") & "] WHERE " & SQL_FIELD_ID & "=" & CStr(CandidateID), dbOpenDynaset)
        If rsCandidate.EOF Then Err.Raise vbObjectError + 9196, "AcceptSelectedRunnerMacroCandidatesCore", "No MC_Candidates row exists for the requested CandidateID." & vbCrLf & "CandidateID: " & CStr(CandidateID)
    Else
        Set rsCandidate = db.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates") & "] WHERE " & SQL_FIELD_SELECTED_FOR_ACCEPTANCE & "<>0 ORDER BY Nz(" & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER & ",0), " & SQL_FIELD_MACRO_TITLE & ", " & SQL_FIELD_CANDIDATE_MACRO_KEY, dbOpenDynaset)
    End If
    copyFields = Array(FIELD_MACRO_KEY, FIELD_MACRO_TITLE, FIELD_MACRO_ACTION, FIELD_MACRO_CATEGORY_KEY, FIELD_MACRO_SUBCATEGORY_KEY, FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER, FIELD_CONFIG_OWNER_DATABASE_ROLE, FIELD_TARGET_PROCEDURE_NAME, FIELD_TARGET_PROCEDURE_MODULE_NAME, FIELD_TARGET_PROCEDURE_KIND, FIELD_TARGET_ARGUMENTS, FIELD_RESULT_HANDLING, FIELD_EXPECTED_RETURN_TYPE, FIELD_REQUIRED_ARGUMENT_COUNT, FIELD_OPTIONAL_ARGUMENT_COUNT, FIELD_TARGET_OBJECT_NAME, FIELD_TARGET_REFERENCE_NAME, FIELD_TARGET_VB_PROJECT_NAME, FIELD_GENERATION_TARGET_DB_NAME, FIELD_GENERATION_TARGET_DB_PATH, FIELD_NAVIGATION_PANE_GROUP, FIELD_CONFIRMATION_MODE, FIELD_ACTION_DESCRIPTION, FIELD_LAUNCHER_TOOLTIP, FIELD_OBJECT_DESCRIPTION, FIELD_SOURCE_NOTES, FIELD_REVIEW_PRIORITY, FIELD_MANUAL_REVIEW_STATUS, FIELD_MANUAL_REVIEW_NOTES)
    Do While Not rsCandidate.EOF
        PumpSetupResponsiveness
        selectedCount = selectedCount + 1
        currentID = SharedValue_NzLong(rsCandidate.fields(FIELD_ID).value)
        On Error GoTo RowErrHandler
        CandidateRecordFromRecordset rsCandidate, cfg
        PulseProgressEvery selectedCount, 0, "Accepting selected candidate row ID " & CStr(currentID)
        If CandidateRowAlreadyAccepted(rsCandidate) Then
            alreadyAcceptedCount = alreadyAcceptedCount + 1
            MarkCandidateAcceptResult rsCandidate, ACCEPTANCE_STATUS_SKIPPED_ALREADY_ACCEPTED, "Skipped because this candidate already links to an accepted configuration.", cfg.AcceptedConfigurationID, True
            GoTo NextCandidateRow
        End If
        validationProblem = SelectedForAcceptanceValidationProblem(cfg)
        If Len(validationProblem) > 0 Then
            invalidCount = invalidCount + 1
            MarkCandidateAcceptResult rsCandidate, ACCEPTANCE_STATUS_INVALID, validationProblem, 0, True
            GoTo NextCandidateRow
        End If
        duplicateReason = SelectedForAcceptanceDuplicateReason(cfg)
        If Len(duplicateReason) > 0 Then
            duplicateCount = duplicateCount + 1
            MarkCandidateAcceptResult rsCandidate, ACCEPTANCE_STATUS_DUPLICATE, duplicateReason, 0, True
            GoTo NextCandidateRow
        End If
        If cfg.HasRequiredArguments Or cfg.RequiredArgumentCount > 0 Then
            AcceptanceStatus = ACCEPTANCE_STATUS_ACCEPTED_NEEDS_ARGUMENTS
            AcceptanceNotes = "Accepted into config but left inactive because the target procedure has required arguments. Fill TargetArguments, then set Active/SelectedForRunnerMacroGeneration as appropriate."
            activeValue = False
        Else
            AcceptanceStatus = ACCEPTANCE_STATUS_ACCEPTED
            AcceptanceNotes = "Accepted into config and selected for Runner Macro generation."
            activeValue = True
        End If
        If Len(NormalizedConfigOwnerDatabaseRole(cfg.ConfigOwnerDatabaseRole)) = 0 Then
            cfg.ConfigOwnerDatabaseRole = NormalizedConfigOwnerDatabaseRole(cfg.CandidateOwnerDatabaseRole)
        End If
        If Len(cfg.ConfigOwnerDatabaseRole) = 0 Then
            AcceptanceStatus = ACCEPTANCE_STATUS_INVALID
            AcceptanceNotes = AcceptanceNotes & vbCrLf & "ConfigOwnerDatabaseRole could not be resolved from CandidateOwnerDatabaseRole. Manually set ConfigOwnerDatabaseRole before accepting."
            activeValue = False
            invalidCount = invalidCount + 1
            MarkCandidateAcceptResult rsCandidate, AcceptanceStatus, AcceptanceNotes, 0, True
            GoTo NextCandidateRow
        End If
        ws.BeginTrans
        transactionStarted = True
        Set rsConfig = db.OpenRecordset(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config"), dbOpenDynaset)
        rsConfig.AddNew
        For Each copyField In copyFields
            SetRecordsetFieldIfExists rsConfig, CStr(copyField), FieldValueIfExists(rsCandidate, CStr(copyField))
        Next copyField
        SetRecordsetFieldIfExists rsConfig, FIELD_MACRO_KEY, cfg.MacroKey
        SetRecordsetFieldIfExists rsConfig, FIELD_CONFIG_OWNER_DATABASE_ROLE, cfg.ConfigOwnerDatabaseRole
        SetRecordsetFieldIfExists rsConfig, FIELD_ACTIVE, activeValue
        SetRecordsetFieldIfExists rsConfig, FIELD_SELECTED_FOR_RUNNER_MACRO_GENERATION, activeValue
        SetRecordsetFieldIfExists rsConfig, FIELD_CONFIGURATION_STATUS, IIf(activeValue, CONFIGURATION_STATUS_ACTIVE, CONFIGURATION_STATUS_NEEDS_REVIEW)
        SetRecordsetFieldIfExists rsConfig, FIELD_CONFIGURED_ON, Now
        rsConfig.Update
        rsConfig.Bookmark = rsConfig.LastModified
        newConfigurationID = SharedValue_NzLong(rsConfig.fields(FIELD_ID).value)
        rsConfig.Close
        Set rsConfig = Nothing
        If newConfigurationID <= 0 Then Err.Raise vbObjectError + 9197, "AcceptSelectedRunnerMacroCandidatesCore", "The accepted configuration was inserted without a valid ID."
        MarkCandidateAcceptResult rsCandidate, AcceptanceStatus, AcceptanceNotes, newConfigurationID, True
        ws.CommitTrans
        transactionStarted = False
        InvalidateDuplicateReviewCache
        If StrComp(AcceptanceStatus, ACCEPTANCE_STATUS_ACCEPTED, vbTextCompare) = 0 Then
            acceptedCount = acceptedCount + 1
        Else
            acceptedNeedsArgumentsCount = acceptedNeedsArgumentsCount + 1
        End If
NextCandidateRow:
        On Error GoTo ErrHandler
        rsCandidate.MoveNext
    Loop
    summaryText = AcceptSelectedRunnerMacroCandidatesSummaryText(selectedCount, acceptedCount, acceptedNeedsArgumentsCount, duplicateCount, invalidCount, alreadyAcceptedCount, errorCount)
    If errorCount > 0 Then
        LogTraceMessage TRACE_LEVEL_ERROR, "CandidateAcceptance", "AcceptSelectedRunnerMacroCandidatesCore", "Candidate acceptance completed with row-level error(s).", summaryText, "Complete candidate acceptance", CANDIDATE_ACCEPT_ROW_ERRORS_ERROR_NUMBER, "AcceptSelectedRunnerMacroCandidatesCore", summaryText, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates"), "Table", "CodeDb", "CandidateAcceptance"
        If showMessage Then
            CoreShowErrorMessage summaryText, vbCritical, "Accept Selected Runner Macro Candidates Completed With Errors", "AcceptSelectedRunnerMacroCandidatesCore", "CandidateAcceptance"
        Else
            Err.Raise CANDIDATE_ACCEPT_ROW_ERRORS_ERROR_NUMBER, "AcceptSelectedRunnerMacroCandidatesCore", summaryText
        End If
    ElseIf showMessage Then
        CoreShowMessage summaryText, vbInformation, "Accept Selected Runner Macro Candidates"
    End If
    On Error Resume Next
    If Not rsConfig Is Nothing Then rsConfig.Close
    If Not rsCandidate Is Nothing Then rsCandidate.Close
    Exit Sub
RowErrHandler:
    errorCount = errorCount + 1
    rowErrNumber = Err.Number
    rowErrSource = Err.Source
    rowErrDescription = Err.Description
    On Error Resume Next
    If Not rsConfig Is Nothing Then rsConfig.Close
    If transactionStarted Then ws.Rollback
    transactionStarted = False
    LogTraceMessage "Error", "CandidateAcceptance", "AcceptSelectedRunnerMacroCandidatesCore", "Candidate acceptance failed for one row.", "Candidate row ID: " & CStr(currentID) & vbCrLf & "Err.Number: " & CStr(rowErrNumber) & vbCrLf & "Err.Source: " & rowErrSource & vbCrLf & "Err.Description: " & rowErrDescription, vbNullString, rowErrNumber, rowErrSource, rowErrDescription, CStr(currentID), "ADC_Candidates", "CodeDb", "CandidateAcceptance"
    MarkCandidateAcceptResult rsCandidate, ACCEPTANCE_STATUS_ERROR, "Acceptance failed for this candidate row. Err.Number: " & CStr(rowErrNumber) & "; Err.Source: " & rowErrSource & "; Err.Description: " & rowErrDescription, 0, False
    Err.Clear
    On Error GoTo ErrHandler
    Resume NextCandidateRow
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not rsConfig Is Nothing Then rsConfig.Close
    If transactionStarted Then ws.Rollback
    If Not rsCandidate Is Nothing Then rsCandidate.Close
    On Error GoTo 0
    If savedErrNumber = CANDIDATE_ACCEPT_ROW_ERRORS_ERROR_NUMBER Then Err.Raise savedErrNumber, savedErrSource, savedErrDescription
    Err.Raise savedErrNumber, "AcceptSelectedRunnerMacroCandidatesCore", "Could not accept selected Runner Macro candidates." & vbCrLf & "Candidate table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates") & vbCrLf & "Configuration table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & vbCrLf & "Err.Description: " & savedErrDescription
End Sub
Private Function AcceptSelectedRunnerMacroCandidatesSummaryText(ByVal selectedCount As Long, ByVal acceptedCount As Long, ByVal acceptedNeedsArgumentsCount As Long, ByVal duplicateCount As Long, ByVal invalidCount As Long, ByVal alreadyAcceptedCount As Long, ByVal errorCount As Long) As String
    AcceptSelectedRunnerMacroCandidatesSummaryText = IIf(errorCount > 0, "Selected macro candidate acceptance completed with row-level errors.", "Selected macro candidate acceptance complete.") & vbCrLf & vbCrLf & "Candidate table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates") & vbCrLf & "Configuration table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")
    AcceptSelectedRunnerMacroCandidatesSummaryText = AcceptSelectedRunnerMacroCandidatesSummaryText & vbCrLf & vbCrLf & "Selected candidate rows processed: " & CStr(selectedCount) & vbCrLf & "Accepted and enabled: " & CStr(acceptedCount) & vbCrLf & "Accepted but needs arguments: " & CStr(acceptedNeedsArgumentsCount) & vbCrLf & "Skipped duplicates: " & CStr(duplicateCount) & vbCrLf & "Invalid rows: " & CStr(invalidCount)
    AcceptSelectedRunnerMacroCandidatesSummaryText = AcceptSelectedRunnerMacroCandidatesSummaryText & vbCrLf & "Already accepted rows skipped: " & CStr(alreadyAcceptedCount) & vbCrLf & "Row-level errors: " & CStr(errorCount) & vbCrLf & vbCrLf & "Review CandidateStatus, ConfigOwnerDatabaseRole, AcceptanceStatus, AcceptedConfigurationID, and AcceptanceNotes in MC_Candidates for candidate details."
    AcceptSelectedRunnerMacroCandidatesSummaryText = AcceptSelectedRunnerMacroCandidatesSummaryText & CoreOperationSummaryForMessage() & NavigationPaneGroupingSummaryForMessage()
End Function
Private Function CandidateRowAlreadyAccepted(ByVal rsCandidate As DAO.Recordset) As Boolean
    Dim statusText As String
    If SharedValue_NzLong(FieldValueIfExists(rsCandidate, FIELD_ACCEPTED_CONFIGURATION_ID)) > 0 Then
        CandidateRowAlreadyAccepted = True
        Exit Function
    End If
    statusText = SharedText_Nz(FieldValueIfExists(rsCandidate, FIELD_ACCEPTANCE_STATUS))
    Select Case UCase$(Trim$(statusText))
        Case UCase$(ACCEPTANCE_STATUS_ACCEPTED), UCase$(ACCEPTANCE_STATUS_ACCEPTED_NEEDS_ARGUMENTS)
            CandidateRowAlreadyAccepted = True
    End Select
End Function
Private Function SelectedForAcceptanceValidationProblem(ByRef cfg As TCandidateRecord) As String
    If Len(Trim$(cfg.MacroKey)) = 0 Then
        cfg.MacroKey = cfg.CandidateMacroKey
    End If
    If Len(Trim$(cfg.MacroKey)) = 0 Then
        SelectedForAcceptanceValidationProblem = "Skipped because CandidateMacroKey/MacroKey is blank."
        Exit Function
    End If
    If Len(Trim$(cfg.MacroTitle)) = 0 Then
        SelectedForAcceptanceValidationProblem = "Skipped because MacroTitle is blank."
        Exit Function
    End If
    If Len(cfg.MacroTitle) > MAX_ACCESS_OBJECT_NAME_LENGTH Then
        SelectedForAcceptanceValidationProblem = "Skipped because MacroTitle exceeds the Access object-name limit. Length: " & CStr(Len(cfg.MacroTitle)) & "; maximum: " & CStr(MAX_ACCESS_OBJECT_NAME_LENGTH) & "."
        Exit Function
    End If
    Select Case UCase$(Trim$(cfg.MacroAction))
        Case UCase$(ACTION_RUNCODE)
            If Len(Trim$(cfg.TargetProcedureName)) = 0 Then
                SelectedForAcceptanceValidationProblem = "Skipped because RunCode candidate has no TargetProcedureName."
            End If
        Case UCase$(ACTION_OPENFORM)
            If Len(Trim$(cfg.TargetObjectName)) = 0 Then
                SelectedForAcceptanceValidationProblem = "Skipped because OpenForm candidate has no TargetObjectName."
            End If
        Case Else
            SelectedForAcceptanceValidationProblem = "Skipped because MacroAction is not importable: " & cfg.MacroAction
    End Select
End Function

Private Function SelectedForAcceptanceDuplicateReason(ByRef cfg As TCandidateRecord) As String
    If ConfigMacroKeyExists(cfg.MacroKey) Then
        SelectedForAcceptanceDuplicateReason = "Skipped because config already contains MacroKey: " & cfg.MacroKey
        Exit Function
    End If
    If ConfigActiveMacroTitleExists(cfg.MacroTitle) Then
        SelectedForAcceptanceDuplicateReason = "Skipped because an active config row already generates the same Access macro object name: " & cfg.MacroTitle
        Exit Function
    End If
    If ConfigTargetCallExists(cfg.ConfigOwnerDatabaseRole, cfg.MacroAction, cfg.TargetProcedureName, cfg.TargetArguments, cfg.TargetObjectName, cfg.ConfirmationMode) Then
        SelectedForAcceptanceDuplicateReason = "Skipped because config already contains the same action-aware target signature for this database role."
        Exit Function
    End If
End Function

Private Sub MarkCandidateAcceptResult(ByVal rsCandidate As DAO.Recordset, ByVal statusText As String, ByVal notesText As String, Optional ByVal AcceptedConfigurationID As Long = 0, Optional ByVal clearSelection As Boolean = True)
    On Error GoTo ErrHandler
    rsCandidate.Edit
    rsCandidate.fields(FIELD_ACCEPTANCE_STATUS).value = statusText
    rsCandidate.fields(FIELD_ACCEPTANCE_NOTES).value = Shared_String_ConvertToNullIfBlank(notesText)
    Select Case UCase$(statusText)
        Case UCase$(ACCEPTANCE_STATUS_ACCEPTED), UCase$(ACCEPTANCE_STATUS_ACCEPTED_NEEDS_ARGUMENTS)
            rsCandidate.fields(FIELD_CANDIDATE_STATUS).value = CANDIDATE_STATUS_ACCEPTED
        Case UCase$(ACCEPTANCE_STATUS_DUPLICATE)
            rsCandidate.fields(FIELD_CANDIDATE_STATUS).value = CANDIDATE_STATUS_DUPLICATE
        Case UCase$(ACCEPTANCE_STATUS_INVALID), UCase$(ACCEPTANCE_STATUS_ERROR)
            rsCandidate.fields(FIELD_CANDIDATE_STATUS).value = CANDIDATE_STATUS_INVALID
    End Select
    If AcceptedConfigurationID > 0 Then
        rsCandidate.fields(FIELD_ACCEPTED_CONFIGURATION_ID).value = AcceptedConfigurationID
        rsCandidate.fields(FIELD_ACCEPTED_ON).value = Now
    End If
    If clearSelection Then rsCandidate.fields(FIELD_SELECTED_FOR_ACCEPTANCE).value = False
    rsCandidate.Update
    InvalidateDuplicateReviewCache
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "MarkCandidateAcceptResult", "Could not update candidate acceptance status." & vbCrLf & "Status: " & statusText & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Function ConfigTargetCallExists(ByVal ConfigOwnerDatabaseRole As String, ByVal MacroAction As String, ByVal TargetProcedureName As String, ByVal TargetArguments As String, ByVal TargetObjectName As String, ByVal ConfirmationMode As String) As Boolean
    On Error GoTo ErrHandler
    Dim targetSignature As String
    targetSignature = TargetCallSignature(ConfigOwnerDatabaseRole, MacroAction, TargetProcedureName, TargetArguments, TargetObjectName, ConfirmationMode)
    If Len(targetSignature) = 0 Then Exit Function
    ConfigTargetCallExists = (CountTargetSignatureRows(targetSignature) > 0)
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ConfigTargetCallExists", "Could not determine whether a configured action-aware target already exists." & vbCrLf & "ConfigOwnerDatabaseRole: " & ConfigOwnerDatabaseRole & vbCrLf & "MacroAction: " & MacroAction & vbCrLf & "TargetProcedureName: " & TargetProcedureName & vbCrLf & "TargetObjectName: " & TargetObjectName & vbCrLf & "TargetArguments: " & TargetArguments & vbCrLf & "ConfirmationMode: " & ConfirmationMode & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

' ===============================================================================================================================================================================
' STALE CANDIDATE REVIEW PROCEDURES
' ===============================================================================================================================================================================

' Purpose: Performs review stale candidates by coordinating `RequireConfigTableForUtilityWorkflow` and `ADCDb` while preserving this module's existing validation and error semantics.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the underlying
'   operation.
'   writeHealthFindings (Input; Boolean): Boolean state describing whether write health findings applies to this invocation; True selects the named condition/path and False selects its complementary/default
'   path.
' Returns: Long: Returns the number of review stale candidates items matched, processed, selected, deleted, or reviewed by this operation; zero represents no qualifying items.
Public Function ReviewStaleCandidates(ByVal showMessage As Boolean, ByVal writeHealthFindings As Boolean) As Long
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim cfg As TCandidateRecord
    Dim reviewedCount As Long
    Dim staleCount As Long
    Dim reviewCount As Long
    Dim invalidRoleCount As Long
    Dim latestBatchID As String
    Dim staleReason As String
    Dim reviewReason As String
    Dim healthFindings As Collection
    RequireConfigTableForUtilityWorkflow "ReviewStaleCandidates"
    Set db = ADCDb()
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: healthFindings holds planned/recovery work or independently observed validation evidence for ReviewStaleCandidates; keep it distinct from the source authority.
    If writeHealthFindings Then Set healthFindings = New Collection
    Set rs = db.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates") & "] ORDER BY " & SQL_FIELD_ID, dbOpenSnapshot)
    Do While Not rs.EOF
        PumpSetupResponsiveness
        reviewedCount = reviewedCount + 1
        PulseProgressEvery reviewedCount, 25, "Candidate rows reviewed"
        CandidateRecordFromRecordset rs, cfg
        latestBatchID = LatestCandidateScanBatchID(NormalizedConfigOwnerDatabaseRole(cfg.CandidateOwnerDatabaseRole))
        staleReason = CandidateStaleReason(cfg, latestBatchID)
        reviewReason = CandidateReviewReason(cfg, latestBatchID)
        If Len(NormalizedConfigOwnerDatabaseRole(cfg.ConfigOwnerDatabaseRole)) = 0 Or Len(NormalizedConfigOwnerDatabaseRole(cfg.CandidateOwnerDatabaseRole)) = 0 Then invalidRoleCount = invalidRoleCount + 1
        If Len(staleReason) > 0 Then
            staleCount = staleCount + 1
            If writeHealthFindings Then
                healthFindings.Add Array("CANDIDATE_STALE_REVIEW", "CandidateRow", cfg.CandidateMacroKey, cfg.TargetProcedureName, cfg.MacroKey, cfg.MacroTitle, cfg.AcceptedConfigurationID, vbNullString, ROW_KIND_CANDIDATE, vbNullString, cfg.CandidateStatus, cfg.ConfigOwnerDatabaseRole, cfg.CandidateOwnerDatabaseRole, cfg.TargetProcedureName, cfg.TargetProcedureModuleName, vbNullString, vbNullString, vbNullString, "Candidate appears stale. Reason: " & staleReason, "Run Mark Stale Candidates if this review is correct.")
            End If
        ElseIf Len(reviewReason) > 0 Then
            reviewCount = reviewCount + 1
            If writeHealthFindings Then
                healthFindings.Add Array("CANDIDATE_STALE_REVIEW_NEEDED", "CandidateRow", cfg.CandidateMacroKey, cfg.TargetProcedureName, cfg.MacroKey, cfg.MacroTitle, cfg.AcceptedConfigurationID, vbNullString, ROW_KIND_CANDIDATE, vbNullString, cfg.CandidateStatus, cfg.ConfigOwnerDatabaseRole, cfg.CandidateOwnerDatabaseRole, cfg.TargetProcedureName, cfg.TargetProcedureModuleName, vbNullString, vbNullString, vbNullString, "Candidate needs review before stale marking. Reason: " & reviewReason, "Review metadata before marking stale.")
            End If
        End If
        rs.MoveNext
    Loop
    If writeHealthFindings Then PersistStaleCandidateReviewHealthFindings healthFindings
    If showMessage Then
        OpenReviewQuery ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Config_StaleCandidates"), "ReviewStaleCandidates"
        CoreShowMessage "Stale candidate review complete." & vbCrLf & vbCrLf & "Candidates reviewed: " & CStr(reviewedCount) & vbCrLf & "Candidates likely stale: " & CStr(staleCount) & vbCrLf & "Candidates needing review: " & CStr(reviewCount) & vbCrLf & "Candidates with invalid role metadata: " & CStr(invalidRoleCount) & vbCrLf & vbCrLf & "No candidate rows were modified.", vbInformation, "Review Stale Candidates"
    End If
    ReviewStaleCandidates = staleCount
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Exit Function
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrSource As String
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrSource = Err.Source
    originalErrDescription = Err.Description
    If showMessage Then
        CoreShowErrorMessage "Could not review stale candidates." & vbCrLf & vbCrLf & "Err.Number: " & CStr(originalErrNumber) & vbCrLf & "Err.Source: " & originalErrSource & vbCrLf & "Err.Description: " & originalErrDescription, vbCritical, "Review Stale Candidates", "ReviewStaleCandidates", "Review"
        Exit Function
    End If
    Err.Raise originalErrNumber, "ReviewStaleCandidates", "Could not review stale candidates." & vbCrLf & "Err.Description: " & originalErrDescription
End Function

' Purpose: Marks stale candidates for confirm status change.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
'   confirmStatusChange (Input Optional; Boolean): When True, enables or requests status change for this call; when False, the routine follows its normal non-forced/default path. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub MarkStaleCandidates(Optional ByVal showMessage As Boolean = True, Optional ByVal confirmStatusChange As Boolean = True)
    RequireNoUnfinishedConfigCandidateMarkingRecovery "MarkStaleCandidates"
    On Error GoTo ErrHandler
    Dim progressStarted As Boolean
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    progressStarted = CoreBeginOperationProgressIfNeeded("Mark stale " & DEFAULT_UTILITY_SHORT_NAME & " candidates", 1)
    CoreUpdateSetupProgress "Mark stale candidates"
    If MarkStaleCandidatesCore(showMessage, confirmStatusChange) = False Then
        If SystemCommandWasCancelled() Then
            If progressStarted Then EndSetupProgress True
            Exit Sub
        End If
        Err.Raise vbObjectError + 5620, "MarkStaleCandidates", "Stale-candidate marking did not complete successfully." & vbCrLf & "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB & vbCrLf & "MC_Candidates table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates") & vbCrLf & "This workflow only updates candidate status fields; review candidate rows for row-level marking details."
    End If
    If progressStarted Then CoreCompleteSetupProgress
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If progressStarted Then CoreFailSetupProgress
    If showMessage Then
        CoreShowErrorMessage "Could not mark stale ADC candidate rows." & vbCrLf & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, vbCritical, "Mark Stale Candidates", "MarkStaleCandidates", "Review"
        Exit Sub
    End If
    Err.Raise savedErrNumber, "MarkStaleCandidates", "Could not mark stale ADC candidate rows." & vbCrLf & "Original Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Function MarkStaleCandidatesCore(Optional ByVal showMessage As Boolean = True, Optional ByVal confirmStatusChange As Boolean = True) As Boolean
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim cfg As TCandidateRecord
    Dim staleCount As Long
    Dim reviewedCount As Long
    Dim markedCount As Long
    Dim skippedCount As Long
    Dim reviewCount As Long
    Dim latestBatchID As String
    Dim staleReason As String
    Dim backupCount As Long
    Dim backupStatus As String
    Dim response As VbMsgBoxResult
    Dim transactionStarted As Boolean
    Dim ws As DAO.workspace
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim backupErrNumber As Long
    Dim backupErrSource As String
    Dim backupErrDescription As String
    Dim currentID As Long
    staleCount = ReviewStaleCandidates(False, True)
    If staleCount = 0 Then
        If showMessage Then
            CoreShowMessage "No stale candidate rows were found. No rows were modified.", vbInformation, "Mark Stale Candidates"
        End If
        MarkStaleCandidatesCore = True
        Exit Function
    End If
    If confirmStatusChange Then
        response = CoreShowMessageResult("Mark " & CStr(staleCount) & " stale candidate row(s)?" & vbCrLf & vbCrLf & "This workflow does not delete candidate rows, config rows, forms, or generated macro objects." & vbCrLf & "It only writes CandidateStatus, SelectedForAcceptance, AcceptanceStatus, and AcceptanceNotes on MC_Candidates rows." & vbCrLf & vbCrLf & "Continue?", vbExclamation Or vbYesNo Or vbDefaultButton2, "Mark Stale Candidates")
        If response <> vbYes Then
            MarkSystemCommandCancelled
            MarkStaleCandidatesCore = False
            Exit Function
        End If
    End If
    On Error Resume Next
    backupCount = BackupAllUserDataCore(False)
    backupStatus = "Backup tables created: " & CStr(backupCount)
    If Err.Number <> 0 Then
        backupErrNumber = Err.Number
        backupErrSource = Err.Source
        backupErrDescription = Err.Description
        Err.Clear
        backupStatus = "Backup failed." & vbCrLf & "Err.Number: " & CStr(backupErrNumber) & vbCrLf & "Err.Source: " & backupErrSource & vbCrLf & "Err.Description: " & backupErrDescription
        LogTraceMessage TRACE_LEVEL_ERROR, "CandidateMarking", "MarkStaleCandidatesCore", "Candidate marking backup failed.", backupStatus, "Back up ADC data tables before stale-candidate marking", backupErrNumber, backupErrSource, backupErrDescription, vbNullString, "Backup", "CodeDb", "CandidateMarking"
        CoreShowErrorMessage backupStatus, vbCritical, "Backup Failed", "MarkStaleCandidatesCore", "CandidateMarking"
        If CoreShowMessageResult("Continue without a backup?", vbQuestion Or vbYesNo Or vbDefaultButton2, "Continue Without Backup?", "MarkStaleCandidatesCore", "CandidateMarking") <> vbYes Then
            MarkSystemCommandCancelled
            MarkStaleCandidatesCore = False
            Exit Function
        End If
    End If
    On Error GoTo ErrHandler
    RequireConfigTableForUtilityWorkflow "MarkStaleCandidatesCore"
    Set ws = DBEngine.CreateWorkspace("ADCCandidateMarking_" & Format$(Now, "hhnnss") & "_" & Format$(CLng(Timer * 1000), "000000000"), "Admin", vbNullString, dbUseJet)
    Set db = ws.OpenDatabase(ADCDb().Name)
    ws.BeginTrans
    transactionStarted = True
    Set rs = db.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates") & "] ORDER BY " & SQL_FIELD_ID, dbOpenDynaset)
    Do While Not rs.EOF
        PumpSetupResponsiveness
        reviewedCount = reviewedCount + 1
        CandidateRecordFromRecordset rs, cfg
        currentID = SharedValue_NzLong(rs.fields(FIELD_ID).value)
        PulseProgressEvery reviewedCount, 0, "Reviewing stale candidate row ID " & CStr(currentID)
        latestBatchID = LatestCandidateScanBatchID(NormalizedConfigOwnerDatabaseRole(cfg.CandidateOwnerDatabaseRole))
        staleReason = CandidateStaleReason(cfg, latestBatchID)
        If Len(staleReason) > 0 And StrComp(cfg.CandidateStatus, CANDIDATE_STATUS_STALE, vbTextCompare) <> 0 Then
            rs.Edit
            SetRecordsetFieldIfExists rs, FIELD_CANDIDATE_STATUS, CANDIDATE_STATUS_STALE
            SetRecordsetFieldIfExists rs, FIELD_SELECTED_FOR_ACCEPTANCE, False
            SetRecordsetFieldIfExists rs, FIELD_ACCEPTANCE_STATUS, ACCEPTANCE_STATUS_INVALID
            SetRecordsetFieldIfExists rs, FIELD_ACCEPTANCE_NOTES, AppendedNote(cfg.AcceptanceNotes, "Marked stale after target review. Reason: " & staleReason & ".")
            rs.Update
            markedCount = markedCount + 1
        ElseIf Len(CandidateReviewReason(cfg, latestBatchID)) > 0 Then
            reviewCount = reviewCount + 1
        Else
            skippedCount = skippedCount + 1
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    ws.CommitTrans
    transactionStarted = False
    db.Close
    ws.Close
    Set db = Nothing
    Set ws = Nothing
    InvalidateDuplicateReviewCache
    If showMessage Then
        OpenReviewQuery ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Config_StaleCandidates"), "MarkStaleCandidatesCore"
        CoreShowMessage "Stale-candidate marking complete." & vbCrLf & vbCrLf & "Candidates reviewed: " & CStr(reviewedCount) & vbCrLf & "Candidates marked stale: " & CStr(markedCount) & vbCrLf & "Candidates skipped: " & CStr(skippedCount) & vbCrLf & "Candidates needing manual review: " & CStr(reviewCount) & vbCrLf & "Backup status: " & backupStatus, vbInformation, "Mark Stale Candidates"
    End If
    MarkStaleCandidatesCore = True
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Exit Function
ErrHandler:
    Dim rollbackErrNumber As Long
    Dim rollbackErrSource As String
    Dim rollbackErrDescription As String
    Dim rollbackStatus As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If transactionStarted Then
        On Error Resume Next
        Err.Clear
        ws.Rollback
        rollbackErrNumber = Err.Number
        rollbackErrSource = Err.Source
        rollbackErrDescription = Err.Description
        On Error GoTo 0
        If rollbackErrNumber = 0 Then
            transactionStarted = False
            rollbackStatus = "Transaction rollback completed without error."
        Else
            rollbackStatus = "Transaction rollback failed. Err.Number: " & CStr(rollbackErrNumber) & "; Err.Source: " & rollbackErrSource & "; Err.Description: " & rollbackErrDescription
        End If
    Else
        rollbackStatus = "No transaction was active when the failure was handled."
    End If
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
    MarkStaleCandidatesCore = False
    If rollbackErrNumber <> 0 Then
        RegisterConfigCandidateMarkingRecoveryFailure "Stale-candidate marking", rollbackStatus & vbCrLf & "Original operation error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription, ws, db
        On Error Resume Next
        LogTraceMessage "Error", "CandidateMarking", "MarkStaleCandidatesCore", "Stale-candidate marking failed and transaction rollback also failed.", "Candidate row ID: " & CStr(currentID) & vbCrLf & "Reason: " & staleReason & vbCrLf & "Backup status: " & backupStatus & vbCrLf & rollbackStatus & vbCrLf & "Original Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Original Err.Source: " & savedErrSource & vbCrLf & "Original Err.Description: " & savedErrDescription, vbNullString, rollbackErrNumber, rollbackErrSource, rollbackErrDescription, CStr(currentID), "ADC_Candidates", "CodeDb", "CandidateMarking"
        On Error GoTo 0
        Err.Raise ERR_CANDIDATE_MARKING_ROLLBACK_FAILED, "MarkStaleCandidatesCore", "Could not recover safely after stale-candidate marking failed because transaction rollback also failed. The dedicated workspace and durable recovery marker were retained, and later Access Development Console work is blocked." & vbCrLf & "Candidate ID: " & CStr(currentID) & vbCrLf & "Reason: " & staleReason & vbCrLf & "Backup status: " & backupStatus & vbCrLf & rollbackStatus & vbCrLf & "Original operation Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Original operation Err.Source: " & savedErrSource & vbCrLf & "Original operation Err.Description: " & savedErrDescription
    End If
    On Error Resume Next
    If Not db Is Nothing Then db.Close
    If Not ws Is Nothing Then ws.Close
    Set db = Nothing
    Set ws = Nothing
    LogTraceMessage "Error", "CandidateMarking", "MarkStaleCandidatesCore", "Stale-candidate marking failed and transaction rollback completed.", "Candidate row ID: " & CStr(currentID) & vbCrLf & "Reason: " & staleReason & vbCrLf & "Backup status: " & backupStatus & vbCrLf & rollbackStatus & vbCrLf & "Original Err.Source: " & savedErrSource & vbCrLf & "Original Err.Description: " & savedErrDescription, vbNullString, savedErrNumber, savedErrSource, savedErrDescription, CStr(currentID), "ADC_Candidates", "CodeDb", "CandidateMarking"
    On Error GoTo 0
    Err.Raise savedErrNumber, "MarkStaleCandidatesCore", "Could not complete stale-candidate marking. Transaction rollback completed without error." & vbCrLf & "Candidate ID: " & CStr(currentID) & vbCrLf & "Reason: " & staleReason & vbCrLf & "Backup status: " & backupStatus & vbCrLf & "Original Err.Source: " & savedErrSource & vbCrLf & "Original Err.Description: " & savedErrDescription
End Function

' Purpose: Checks whether current state and policy permit the caller to didate stale reason.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   cfg (Input; TCandidateRecord): ADC candidate record containing persisted discovery, proposed configuration, and acceptance-state values used by the operation.
'   latestBatchID (Input; String): Identifier of the most recent candidate-scan batch whose rows should be reviewed or compared.
' Returns: String: Returns the resolved candidate stale reason text in the canonical representation consumed by this module's caller; vbNullString represents no applicable value where the existing contract
'   permits it.
Public Function CandidateStaleReason(ByRef cfg As TCandidateRecord, ByVal latestBatchID As String) As String
    If Len(Trim$(cfg.TargetProcedureName)) = 0 Then
        CandidateStaleReason = "TargetProcedureName is blank"
    ElseIf Len(NormalizedConfigOwnerDatabaseRole(cfg.ConfigOwnerDatabaseRole)) = 0 Then
        CandidateStaleReason = "ConfigOwnerDatabaseRole is blank or invalid"
    ElseIf Len(NormalizedConfigOwnerDatabaseRole(cfg.CandidateOwnerDatabaseRole)) = 0 Then
        CandidateStaleReason = "CandidateOwnerDatabaseRole is blank or invalid"
    ElseIf CandidateStatusIsDirectlyStale(cfg.CandidateStatus) Then
        CandidateStaleReason = "CandidateStatus is " & cfg.CandidateStatus
    ElseIf CandidateRowAlreadyConfigured(cfg) Then
        CandidateStaleReason = "same candidate is already accepted/configured"
    ElseIf cfg.HasRequiredArguments And Len(Trim$(cfg.TargetArguments)) = 0 Then
        CandidateStaleReason = "candidate has required arguments and no TargetArguments plan"
    ElseIf Len(Trim$(latestBatchID)) > 0 And Len(Trim$(cfg.LastScanBatchID)) > 0 And StrComp(cfg.LastScanBatchID, latestBatchID, vbBinaryCompare) <> 0 And Len(Trim$(cfg.ScanSignature)) > 0 Then
        CandidateStaleReason = "candidate scan signature is not in the latest scan batch for " & cfg.CandidateOwnerDatabaseRole
    End If
End Function

' Purpose: Checks whether current state and policy permit the caller to didate review reason.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   cfg (Input; TCandidateRecord): ADC candidate record containing persisted discovery, proposed configuration, and acceptance-state values used by the operation.
'   latestBatchID (Input; String): Identifier of the most recent candidate-scan batch whose rows should be reviewed or compared.
' Returns: String: Returns the resolved candidate review reason text in the canonical representation consumed by this module's caller; vbNullString represents no applicable value where the existing contract
'   permits it.
Public Function CandidateReviewReason(ByRef cfg As TCandidateRecord, ByVal latestBatchID As String) As String
    If Len(Trim$(cfg.LastScanBatchID)) = 0 Then
        CandidateReviewReason = "missing LastScanBatchID"
    ElseIf cfg.LastScannedOn = 0 Then
        CandidateReviewReason = "missing LastScannedOn"
    ElseIf Len(Trim$(cfg.ScanSignature)) = 0 Then
        CandidateReviewReason = "missing ScanSignature"
    ElseIf Len(Trim$(latestBatchID)) = 0 Then
        CandidateReviewReason = "no latest scan batch could be resolved for CandidateOwnerDatabaseRole"
    End If
End Function

Private Function CandidateStatusIsDirectlyStale(ByVal CandidateStatus As String) As Boolean
    Select Case LCase$(Trim$(CandidateStatus))
        Case LCase$(CANDIDATE_STATUS_STALE), LCase$(CANDIDATE_STATUS_SCAN_ERROR), LCase$(CANDIDATE_STATUS_INVALID), LCase$(CANDIDATE_STATUS_DUPLICATE), LCase$(CANDIDATE_STATUS_CONFIGURED)
            CandidateStatusIsDirectlyStale = True
    End Select
End Function

Private Function CandidateRowAlreadyConfigured(ByRef cfg As TCandidateRecord) As Boolean
    If Len(Trim$(cfg.CandidateMacroKey)) > 0 Then
        If ConfigMacroKeyExists(cfg.CandidateMacroKey) Then
            CandidateRowAlreadyConfigured = True
            Exit Function
        End If
    End If
    CandidateRowAlreadyConfigured = ConfigTargetCallExists(cfg.ConfigOwnerDatabaseRole, cfg.MacroAction, cfg.TargetProcedureName, cfg.TargetArguments, cfg.TargetObjectName, cfg.ConfirmationMode)
End Function

' Purpose: Performs latest candidate scan batch ID by coordinating `NormalizedConfigOwnerDatabaseRole` and `ManagedObjectPhysicalName` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   CandidateOwnerDatabaseRole (Input; String): Identifies the database role in which the candidate procedure or object is being evaluated.
' Returns: String: Returns the persisted/current identifier for latest candidate scan batch ID; vbNullString indicates that no qualifying run/item is currently available.
Public Function LatestCandidateScanBatchID(ByVal CandidateOwnerDatabaseRole As String) As String
    On Error GoTo ErrHandler
    Dim normalizedRole As String
    Dim rs As DAO.Recordset
    Dim sqlText As String
    normalizedRole = NormalizedConfigOwnerDatabaseRole(CandidateOwnerDatabaseRole)
    If Len(normalizedRole) = 0 Then Exit Function
    sqlText = "SELECT TOP 1 " & SQL_FIELD_LAST_SCAN_BATCH_ID & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates") & "] WHERE " & SQL_FIELD_CANDIDATE_OWNER_DATABASE_ROLE & "=" & SharedDao_SqlQuote(normalizedRole) & " AND Len(Nz(" & SQL_FIELD_LAST_SCAN_BATCH_ID & ",''))>0 ORDER BY " & SQL_FIELD_LAST_SCANNED_ON & " DESC, " & SQL_FIELD_ID & " DESC"
    Set rs = ADCDb().OpenRecordset(sqlText, dbOpenSnapshot)
    If Not rs.EOF Then LatestCandidateScanBatchID = CStr(rs.fields(FIELD_LAST_SCAN_BATCH_ID))
    rs.Close
    Set rs = Nothing
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
    Err.Raise savedErrNumber, "LatestCandidateScanBatchID", "Could not establish the latest candidate scan batch." & vbCrLf & "Candidate owner database role: " & IIf(Len(normalizedRole) = 0, "(invalid or blank)", normalizedRole) & vbCrLf & "SQL: " & sqlText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Function
