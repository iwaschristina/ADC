Attribute VB_Name = "x__UTIL_ADC_UI"
' ================================================================================================================================================================================
' ChatGPT Rewrite Tracking Header
' ================================================================================================================================================================================
' ChatGPT Previous Version ID: CGPT-FQCH-TMRD-QWDP-TSU4-KN6J
' ChatGPT Baseline Version ID: CGPT-FQCH-TMRD-QWDP-TSU4-KN6J
' ChatGPT Version ID: CGPT-WUCV-T29E-8MB2-C4SL-YABH
' ChatGPT Version Timestamp: Tuesday 9/1/26 12:44 PM
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
' STANDARD MODULE: Access Development Console UI Utility
' VB_Name: x__UTIL_ADC_UI
' Entry Points:
'     BeginSetupControl
'     BeginScanProgressControl
'     EndSetupControl
'     UIRecoveryBlockedForSetupPreflight
'     CancelSetupPreparedWork
'     RetryUIRecovery
'     ChatGPT_OpenFormByName
'     ChatGPT_OpenQueryByName
'     ChatGPT_RunProcedure
'     ChatGPT_FilterCurrentSelection
'     ChatGPT_OpenConfigurationDetailFromList
' SHARED CODE:
'     SharedDao_AccessObjectExists(ByVal db As DAO.Database, ByVal objectType As AcObjectType, ByVal objectName As String) -> Boolean
'     SharedProgress_TryEnd(ByRef progressActive As Boolean, ByRef currentStep As Long, ByRef maxSteps As Long, ByVal clearStatus As Boolean, ByRef errorNumber As Long, ByRef errorSource As String, ByRef
'     errorDescription As String) -> Boolean
'     SharedProgress_TryPulse(ByVal progressActive As Boolean, ByVal statusText As String, ByVal statusPrefix As String, ByRef errorNumber As Long, ByRef errorSource As String, ByRef errorDescription As
'     String) -> Boolean
' CLASSES USED: None
' DYNAMIC STRING REFERENCES:
'     Dynamic-reference scan stopped after 2000 logical line(s) to prevent a long-running scan; review this module manually if needed.
' LastScanTimestamp: Saturday 8/15/26 8:48 AM
' Purpose:
'     Owns Setup Control and scan-progress presentation, Pause/Cancel UI checkpoints, shared UI status/progress rendering, UI recovery state, and generic generated-form runtime callbacks used to open, filter,
'     or invoke existing ADC surfaces.
'     Specialized generated-form construction, Help UI, Add-In Review UI, and Ribbon lifecycle/runtime behavior are owned by x__UTIL_ADC_UI_FORMS, x__UTIL_ADC_UI_HELP, x__UTIL_ADC_UI_REVIEW, and
'     x__UTIL_ADC_UI_RIBBON respectively.
' Author: Christina Shonts
' Support: Access Development Console Setup Control, progress/status presentation, UI recovery, and generic runtime callback support.
' Code Category: Access Development Console V6 shared UI runtime/progress utility.
' Requirements:
'     Microsoft Access form/runtime support is required for Setup Control, scan-progress forms, and generic form/query callbacks.
'     DAO and the Access object model are required for UI recovery verification and managed-object existence/open-state checks.
' Reference Dependencies:
'     Access object model, DAO, Core progress/recovery authorities, Defaults naming authorities, and the dedicated UI owner modules listed above.
'     Setup Control may delegate generated-form signature work to x__UTIL_ADC_UI_FORMS without assuming D2 ownership of the broader generated-form build.
' DB Object Dependencies:
'     Generated Setup Control and scan-progress form identities come from the canonical managed naming/defaults authorities.
'     Durable UI recovery records are stored through Core-owned recovery storage while this module owns UI recovery keys, payload semantics, and in-memory recovery state.
' Usage:
'     Use BeginSetupControl/EndSetupControl for full Setup presentation and BeginScanProgressControl for standalone scan presentation.
'     Use RetryUIRecovery only to resolve UI recovery work after a failed UI mutation or interrupted transaction.
'     Generated forms use the public ChatGPT_* callbacks here for generic open/filter/run behavior; domain-specific callbacks remain with their dedicated UI owner modules.
' Version:
'     Shifts the standard-Ribbon Setup lifecycle to H1-H3 and deferred Navigation Pane placement to I3 after insertion of dependency Phase F.
'     Adds Setup/scan control, UI recovery, progress/status, durable recovery, and generic callback procedure sections without changing UI behavior.
'     Narrows module-internal UI helpers while retaining generated Access callbacks and other externally required public entry points.
'     Repaired the FormDisplayNameCalc Long Text generated-form sort path by passing the shared DAO SELECT builder an array-shaped no-order contract before appending the canonical sortable ORDER BY expression.
'     Documented the intentional pre-D bootstrap materialization of Setup Control and Report Error forms and corrected stale generated-form preflight/Help proof labels from F1 to D1.
'     Moved the standard-Ribbon Setup lifecycle to G1-G3, renamed generated-form same-session proof ownership to D2, and made Setup D2 defer all physical Navigation Pane grouping to H3.
'     Managed-header dependency metadata updated because x__UTIL_ADC_DEFAULTS is now a standard module; executable behavior unchanged.
'     Reused the G2 operation's lazily materialized System Macro definitions collection for final canonical group counts when available.
'     Carried each continuous-list field's already-resolved generated lookup metadata into combo configuration so one form-build operation no longer resolves the same lookup or static value list twice.
'     Continuous-list forms now keep conventional header captions only when one physical Detail row fits; wrapped records place each caption in Detail directly above its control using the same authoritative wrap
'     calculation.
'     Routed ordinary runtime popups and shared System Command cancellation state through Core authorities.
'     Inlined the Ribbon progress-scope trim expression into RibbonStatusPartPrefixText and removed the unnecessary private micro-helper without changing status behavior.
'     Removed UI-owned durable-recovery database/storage helpers and routed independent recovery database opening and readiness through Core while retaining UI payload/recovery semantics.
'     Consolidated generated-form, Report Error, Help, Add-In Review, and generic UI callbacks here while preserving Setup progress presentation authority.
'     Expanded Setup Control status-label height and shifted the elapsed/action controls plus detail height equally so legitimate wrapped Setup status text has additional vertical room without overlap.
'     Kept the running Setup Control popup non-modal so D2-generated temporary Design-view forms can become active for native Access stacked-layout operations while pause/cancel/status behavior remains available.
'     Routed nested generated-form progress through the shared progress dispatcher so Setup Control receives detailed status, heartbeat registration, and forced event pumps instead of status-bar-only updates.
'     Corrected UtilityLazyLoadSelectedTab and consolidated generated-form callbacks on the shared DEFAULT_UI_MESSAGE_TITLE authority.
'     Delegated nested parent/child progress composition to Core.
'     Removed duplicate UI process-run state and lifecycle; UI progress, Ribbon, diagnostics, and generated-form workflows now consume the Core tracker directly.
'     Centralized physical Access status-bar clearing in ClearVisibleStatusText so external modules no longer bypass the UI status authority.
'     Removed duplicate system-command cancellation and Ribbon group-prefix policies; UI now consumes COMMON and DEFAULTS authorities directly.
'     Removed duplicate UI Managed Object and ampersand Ribbon substitutions; generation and callbacks now consume the DEFAULTS-owned rules directly without changing prior context-specific spacing behavior.
'     Exposed the existing generated UI_Frm_Registry registration implementation as the sole cross-module generated-form catalog writer for D4 fallback repair and reports its actual write count.
'     Current UI documentation describes direct managed-object launch behavior.
'     Ribbon generation now rejects duplicate tab/group/button CustomIDs before persistence, normalizes emitted Ribbon identities, and collapses duplicate USysRibbons rows for the same CustomRibbonID to one
'     authoritative row.
'     Ribbon generation now enforces one Ribbon item per managed object, one authoritative Ribbon group per item, and one unique explicitly assigned permanent tag per emitted button and persisted target row.
'     Ribbon button and group XML IDs now derive from those permanent tags rather than traversal/group counters, so insertion, deletion, reordering, or button movement cannot renumber surviving controls.
'     Ribbon target tags are explicit permanent values owned by Defaults rather than rebuild-order numbers or calculated object hashes, so rebuilds cannot silently assign a different tag to an existing item.
'     Every generated Ribbon tab, group, and button now receives a unique explicitly authored permanent tag; missing assignments fail closed instead of allowing AccessUI or runtime code to invent a replacement tag.
'     Generated Help now points Terminology, Report Error, and Add-Ins Review to their direct managed-object targets and documents only the remaining work-performing Add-Ins commands.
'     Added the canonical Reference Manager OnOpen contract to centralized generated-form build and validation so direct form opens own reference preparation after every Setup rebuild.
'     Integrated the managed Report Error dialog into the Setup-generated form build/validation path and updated Help to describe Submit Report/Cancel behavior and physical trace-log preservation.
' Notes:
'     Setup Control and scan-progress forms are bootstrap UI surfaces that may be created before D2 so Setup can present status and cancellation without transferring general generated-form ownership into this
'     module.
'     General generated-form construction and validation remain in x__UTIL_ADC_UI_FORMS; Help, Add-In Review, and Ribbon responsibilities remain in their dedicated modules.
'     UI recovery can coordinate with specialized owners, but this module does not absorb their domain logic or Setup-phase responsibilities.
' Core Rules Enforced Here:
'     Reporting ownership rule: UI action entry points own one diagnostic popup; callbacks/shared helpers log/raise upward and must not show duplicate popups.
'     Setup Control/progress presentation remains UI-owned while Setup phase mutation authority remains with the module responsible for that phase.
'     Generic callbacks open/filter/invoke existing managed surfaces and do not take ownership of the target domain's build, repair, or persistence logic.
' Design Map:
'     BeginSetupControl, BeginScanProgressControl, SetupControlUpdateStatus, and EndSetupControl own visible Setup/scan progress presentation and Pause/Cancel interaction.
'     UIRecoveryBlockedForSetupPreflight, durable UI recovery helpers, and RetryUIRecovery own the shared UI recovery gate while delegating specialized recovery actions to the responsible module.
'     ChatGPT_OpenFormByName, ChatGPT_OpenQueryByName, ChatGPT_RunProcedure, ChatGPT_FilterCurrentSelection, and ChatGPT_OpenConfigurationDetailFromList provide generic generated-form runtime callbacks.
' Configuration:
'     Defaults supplies managed Setup Control/scan-progress identities and common UI text. Core supplies progress composition, durable recovery storage, and shared message presentation.
'     Specialized UI configuration remains with x__UTIL_ADC_UI_FORMS, x__UTIL_ADC_UI_HELP, x__UTIL_ADC_UI_REVIEW, and x__UTIL_ADC_UI_RIBBON.
' Behavior:
'     Setup Control may materialize its own bootstrap forms before D2, but this module does not perform the general D2 generated-form build or later Ribbon/Help/Add-In Review work.
'     Generic callbacks act on already-defined targets and preserve the target module's validation, ownership, and error semantics.
'     UI recovery fails closed when in-memory or durable recovery state is unresolved.
' Code Status:
'     Active shared UI runtime module for Setup Control, progress/status presentation, UI recovery, and generic generated-form callbacks. Specialized UI responsibilities are consolidated in their dedicated modules.
' Todo: None currently recorded in this header.
' Reporting Bugs: None currently recorded in this header.
' Module Level Bugs: None currently recorded in this header.
' Public Subs:
'     BeginSetupControl(Optional ByVal initialStatusText As String = "Setup is starting.") -> None
'     BeginScanProgressControl(Optional ByVal initialStatusText As String = "Scan is starting.", Optional ByVal modalHeadingText As String = vbNullString, Optional ByVal modalLeadText As String = "Scan is
'     running.", Optional ByVal allowCancel As Boolean = True) -> None
'     EndSetupControl() -> None
'     SetupControlUpdateStatus(ByVal statusText As String, Optional ByVal updateStatusBar As Boolean = True, Optional ByVal forceDisplay As Boolean = False, Optional ByVal pumpAfterDisplay As Boolean = True) ->
'     None
'     ResetSetupControlPerformanceDiagnostics() -> None
'     EnsureSetupControlForm(ByVal formName As String, ByVal expectedTag As String, ByVal runningForm As Boolean, Optional ByVal scanForm As Boolean = False) -> None
'     RegisterUIRecoveryFailure(ByVal operationName As String, ByVal detailText As String, Optional ByVal recoveryWorkspace As DAO.workspace = Nothing, Optional ByVal recoveryState As Collection = Nothing,
'     Optional ByVal recoveryType As String = vbNullString, Optional ByVal recoveryPayload As String = vbNullString) -> None
'     RequireNoUnfinishedUIRecovery(ByVal requestedOperationName As String) -> None
'     SetUIStatusBarText(ByVal statusText As String) -> None
'     BeginChatGPTProgress(ByVal captionText As String, ByVal maxValue As Long) -> None
'     EndChatGPTProgress() -> None
'     FailChatGPTProgress(ByVal statusText As String) -> None
'     UpdateChatGPTProgress(ByVal statusText As String) -> None
'     SetVisibleStatusText(ByVal statusText As String, Optional ByVal forceDisplay As Boolean = True, Optional ByVal pumpAfterDisplay As Boolean = False) -> None
'     ClearVisibleStatusText() -> None
'     PrintUIStatus(ByVal sourceProcedure As String, ByVal statusText As String, Optional ByVal objectName As String = vbNullString, Optional ByVal contextName As String = "UI") -> None
'     DeleteUIDurableRecoveryRecord() -> None
'     PersistUIDurableRecoveryRecord(ByVal operationName As String, ByVal detailText As String, ByVal recoveryType As String, ByVal recoveryPayload As String, ByVal recoveryStatus As String) -> None
'     PrepareUIDurableRecovery(ByVal operationName As String, ByVal detailText As String, ByVal recoveryType As String, ByVal recoveryPayload As String) -> None
'     RefreshNavigationPaneAfterGrouping() -> None
'     CompleteUITraceRunIfOwned(ByVal traceRunOwnedByProcedure As Boolean, ByVal completeMessage As String, Optional ByVal clearStatus As Boolean = True) -> None
'     RetryUIRecovery(Optional ByVal showMessage As Boolean = True) -> None
' Public Functions:
'     SetupControlPump(Optional ByVal forcePump As Boolean = False, Optional ByVal suppressHeartbeat As Boolean = False) -> Boolean
'     SetupControlPerformanceSummary() -> String
'     SetupControlHeartbeat(Optional ByVal showMessage As Boolean = True) -> Boolean
'     SetupControlCancellationRequested() -> Boolean
'     SetupControlPause() -> Boolean
'     SetupControlCancel() -> Boolean
'     SetupControlResume() -> Boolean
'     SetupControlWorkCheckpoint(ByVal checkpointText As String, Optional ByVal workerRoutine As String = vbNullString, Optional ByVal logicalItem As String = vbNullString) -> Boolean
'     SetupControlCheckpoint(ByVal checkpointText As String) -> Boolean
'     CancelSetupPreparedWork(ByRef resultText As String) -> Boolean
'     UIRecoveryBlockedForSetupPreflight(Optional ByRef detailText As String = vbNullString, Optional ByVal durableRecoverySnapshot As Variant) -> Boolean
'     UIRecoveryBackupReadable(ByVal backupPath As String, ByRef failureText As String) -> Boolean
'     UIRecoveryBlocked() -> Boolean
'     UniqueRecoveryFormName(ByVal prefixText As String) -> String
'     ClearUIRecoveryFailure(ByVal operationName As String, ByRef failureDetail As String) -> Boolean
'     ResolveUIRecoveryWorkspaceTransaction(ByVal recoveryContext As String) -> String
'     StatusBarTextWithCurrentSetupElapsed(ByVal statusText As String) -> String
'     UIRenderSharedProgressActivity(ByVal progressActive As Boolean, ByVal semanticStatusText As String, ByVal statusBarText As String, ByVal statusPrefix As String, ByVal forceDisplay As Boolean, ByRef
'     errNumber As Long, ByRef errSource As String, ByRef errDescription As String) -> Boolean
'     UIEndSharedProgress(ByRef progressActive As Boolean, ByRef currentStep As Long, ByRef maxSteps As Long, ByVal clearStatus As Boolean, ByRef errNumber As Long, ByRef errSource As String, ByRef
'     errDescription As String) -> Boolean
'     UITextMismatchDetail(ByVal existingText As String, ByVal generatedText As String) -> String
'     UIVisibleSuccessText(ByVal statusText As String) -> String
'     UIScreenContextForMessage() -> String
'     TryParseUIDurableBoolean(ByVal sourceText As String, ByRef parsedValue As Boolean) -> Boolean
'     UIDurableRecoveryRecordExists(ByRef operationName As String, ByRef detailText As String, ByRef recoveryType As String, ByRef recoveryPayload As String, ByRef recoveryStatus As String) -> Boolean
'     BeginUITraceRunIfNeeded(ByVal contextName As String) -> Boolean
'     ChatGPTSameDbOnlyMessage(ByVal actionText As String) -> String
'     UIVariantForMessage(Optional ByVal value As Variant) -> String
'     ChatGPT_OpenFormByName(ByVal formName As String) -> Boolean
'     ChatGPT_ClearCurrentFormFilter() -> Boolean
'     CurrentFormRecordText(ByVal frm As Access.Form, ByVal fieldName As String) -> String
'     LooksLikeSql(ByVal valueText As String) -> Boolean
'     ChatGPT_OpenQueryByName(ByVal queryName As String) -> Boolean
'     ChatGPT_RunProcedure(ByVal ProcedureName As String) -> Boolean
'     ChatGPT_FilterCurrentSelection() -> Boolean
'     ChatGPT_OpenConfigurationDetailFromList(Optional ByVal rowID As Variant) -> Boolean
' Public Properties: None
' ================================================================================================================================================================================

Option Compare Database
Option Explicit

' ===============================================================================================================================================================================
' MODULE-LEVEL DECLARATIONS
' ===============================================================================================================================================================================

' Windows structures used by read-only Access minimized-state inspection and Setup Control form window handling.
Public Type TReportErrorRect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type TSetupWindowPoint
    X As Long
    Y As Long
End Type
Private Type TSetupWindowPlacement
    Length As Long
    flags As Long
    showCommand As Long
    MinimumPosition As TSetupWindowPoint
    MaximumPosition As TSetupWindowPoint
    NormalPosition As TReportErrorRect
End Type

' Windows API declarations used by read-only Access minimized-state inspection and Setup Control form window handling.
Public Declare PtrSafe Sub ReportErrorSleep Lib "kernel32" Alias "Sleep" (ByVal milliseconds As Long)
Private Declare PtrSafe Function SetupGetWindowPlacement Lib "user32" Alias "GetWindowPlacement" (ByVal windowHandle As LongPtr, ByRef placement As TSetupWindowPlacement) As Long
Private Declare PtrSafe Function SetupSetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal windowHandle As LongPtr, ByVal insertAfter As LongPtr, ByVal windowLeft As Long, ByVal windowTop As Long, ByVal windowWidth As Long, ByVal windowHeight As Long, ByVal flags As Long) As Long

' Setup controller forms, heartbeat, window constants, and runtime state.
Public Const SETUP_CONTROL_RUNNING_FORM_TAG As String = "ADCSetupControlRunningV1"
Public Const SETUP_CONTROL_PAUSED_FORM_TAG As String = "ADCSetupControlPausedV1"
Public Const SCAN_PROGRESS_FORM_TAG As String = "ADCScanProgressV1"
Private Const SETUP_CONTROL_STATUS_CONTROL As String = "lblSetupStatus"
Private Const SETUP_CONTROL_ELAPSED_CONTROL As String = "lblSetupElapsed"
Private Const SETUP_WINDOW_SHOW_MINIMIZED As Long = 2
Private Const SETUP_WINDOW_MINIMIZE As Long = 6
Private Const SETUP_WINDOW_SHOW_MIN_NO_ACTIVE As Long = 7
Private Const SETUP_WINDOW_FORCE_MINIMIZE As Long = 11
Private Const SETUP_WINDOW_POS_NO_SIZE As Long = &H1
Private Const SETUP_WINDOW_POS_NO_MOVE As Long = &H2
Private Const SETUP_WINDOW_POS_NO_ACTIVATE As Long = &H10
Private Const SETUP_WINDOW_POS_SHOW_WINDOW As Long = &H40
Private Const SETUP_CONTROL_PUMP_REQUEST_CHECK_INTERVAL As Long = 8
Private Const SETUP_CONTROL_SLOW_CHECKPOINT_SECONDS As Double = 1#
Public mSetupControlActive As Boolean
Public mSetupControlScanMode As Boolean
Private mSetupControlModalHeadingText As String
Private mSetupControlModalLeadText As String
Private mSetupControlCancelEnabled As Boolean
Private mSetupControlPauseRequested As Boolean
Private mSetupControlCancelRequested As Boolean
Private mSetupControlResumeRequested As Boolean
Private mSetupControlCancelRequestedTimer As Single
Private mSetupControlCancelObservedTimer As Single
Private mSetupControlCancelObserved As Boolean
Private mSetupControlCancelObservedWorker As String
Private mSetupControlCancelObservedItem As String
Public mSetupControlStatusText As String
Private mSetupControlStartedTimer As Single
Private mSetupControlLastHeartbeatTimer As Single
Private mSetupControlLastPumpTimer As Single
Private mSetupControlPumpActive As Boolean
Private mSetupControlCheckpointActive As Boolean
Private mSetupControlCheckpointDisplayActive As Boolean
Private mSetupControlPaused As Boolean
' D2 can intentionally keep a stale Setup Control surface closed while its planned physical replacement/rollback boundary is active; event pumping continues, but automatic controller reopening and Pause transitions
' are deferred until the boundary ends.
Public mSetupControlVisibilitySuspended As Boolean
Private mSetupControlLastDisplayedStatusText As String
Private mSetupControlUpdateRequestCount As Long
Private mSetupControlDisplayDispatchCount As Long
Private mSetupControlHeartbeatDispatchCount As Long
Private mSetupControlDuplicateDisplaySkipCount As Long
Private mSetupControlPumpRequestCount As Long
Private mSetupControlPumpCheckCounter As Long
Private mSetupControlPumpCoalescedRequestCount As Long
Private mSetupControlDoEventsDispatchCount As Long
Private mSetupControlLastDoEventsTimer As Single
Private mSetupControlMaximumDoEventsGapSeconds As Double
Private mSetupControlLastWindowEnforcementTimer As Single

' Shared UI child-progress adapter state.
Private mChatGPTProgressActive As Boolean
Private mChatGPTProgressStep As Long
Private mChatGPTProgressMax As Long
Private mChatGPTProgressScopeLabel As String
Private mChatGPTProgressNestedUnderParent As Boolean

' Status-bar update diagnostics.
Private mUIStatusBarFailureCount As Long

' UI recovery errors, durable markers, payload, workspace, and form-reference rollback state.
Public Const ERR_UI_RECOVERY_BLOCKED As Long = vbObjectError + 7210

Public Const ERR_UI_DURABLE_RECOVERY_FAILED As Long = vbObjectError + 7218

Private Const UI_RECOVERY_KEY As String = "ADC.UIRecoveryStateV1"
Private Const UI_RECOVERY_STATUS_PREPARED As String = "Prepared"
Public Const UI_RECOVERY_TYPE_GENERATED_FORMS As String = "GeneratedForms"
Public Const UI_RECOVERY_TYPE_RIBBON_TRANSACTION As String = "RibbonTransaction"

' UI recovery runtime state.
Public mUIRecoveryBlocked As Boolean
Public mUIRecoveryOperationName As String
Public mUIRecoveryDetail As String
Public mUIRecoveryType As String
Public mUIRecoveryPayload As String
Public mUIDurableRecoveryOwned As Boolean
Public mUIRecoveryWorkspace As DAO.workspace

' ===============================================================================================================================================================================
' PROCEDURES
' ===============================================================================================================================================================================

' =============================================================================================================================================================================
' SETUP / SCAN CONTROL AND BOOTSTRAP UI PROCEDURES
' =============================================================================================================================================================================

Private Function SetupControlQualifiedStatusText(ByVal statusText As String) As String
    On Error GoTo UseOriginal
    Dim qualifiedText As String
    statusText = Trim$(statusText)
    If Len(statusText) = 0 Then Exit Function
    qualifiedText = CoreQualifyProgressStatusText(statusText)
    SetupControlQualifiedStatusText = qualifiedText
    Exit Function
UseOriginal:
    Err.Clear
    SetupControlQualifiedStatusText = statusText
End Function

Private Function SetupHostWindowIsMinimized() As Boolean
    On Error GoTo SafeExit
    Dim windowHandle As LongPtr
    Dim placement As TSetupWindowPlacement
    windowHandle = Application.hWndAccessApp
    If windowHandle = 0 Then Exit Function
    placement.Length = Len(placement)
    If SetupGetWindowPlacement(windowHandle, placement) = 0 Then Exit Function
    Select Case placement.showCommand
        Case SETUP_WINDOW_SHOW_MINIMIZED, SETUP_WINDOW_MINIMIZE, SETUP_WINDOW_SHOW_MIN_NO_ACTIVE, SETUP_WINDOW_FORCE_MINIMIZE
            SetupHostWindowIsMinimized = True
    End Select
SafeExit:
    Err.Clear
End Function

Private Sub LogSetupControlStartupStage(ByVal messageText As String, ByVal detailText As String, Optional ByVal traceLevel As String = "Info")
    If Len(Trim$(detailText)) = 0 Then detailText = messageText
    LogTraceMessage traceLevel, "SetupControl", "BeginSetupControl", messageText, detailText, "Setup Preflight", 0, vbNullString, vbNullString, vbNullString, "SetupControl", "CodeDb", "UI"
    FlushTraceLogFile
End Sub

Private Sub LogSetupControlCheckpointStage(ByVal messageText As String, ByVal checkpointText As String, Optional ByVal detailText As String = vbNullString, Optional ByVal traceLevel As String = "Info")
    On Error Resume Next
    If Len(Trim$(detailText)) > 0 Then detailText = "Checkpoint: " & checkpointText & " | " & detailText Else detailText = "Checkpoint: " & checkpointText
    LogTraceMessage traceLevel, "SetupControl", "SetupControlCheckpoint", messageText, detailText, checkpointText, 0, vbNullString, vbNullString, vbNullString, "SetupControl", "CodeDb", "Performance"
    FlushTraceLogFile
    On Error GoTo 0
End Sub

' Purpose: Initializes setup control, coordinating `LogSetupControlStartupStage` and `ManagedObjectPhysicalName` as the implementation requires.
' Public role: Public module entry point. This procedure is listed in the module header as an intentional externally callable ADC entry point and may be invoked directly by Access/VBA callers.
' Parameters:
'   initialStatusText (Input Optional; String): Initial status text text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine. Default: "Setup is
'     starting.".
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub BeginSetupControl(Optional ByVal initialStatusText As String = "Setup is starting.")
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    On Error GoTo ErrHandler
    LogSetupControlStartupStage "Existing Setup Control cleanup starting.", "Next call: EndSetupControl"
    EndSetupControl
    LogSetupControlStartupStage "Existing Setup Control cleanup completed.", "Next call: EnsureSetupControlForms"
    LogSetupControlStartupStage "Setup Control form verification starting.", "Running form: " & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_RUNNING) & vbCrLf & "Paused form: " & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_PAUSED)
    ' PRE-D BOOTSTRAP EXCEPTION: these canonical generated forms must exist before numbered Setup Parts begin so pause/cancel/progress control is available from Preflight onward. D1-D4 retain ordinary generated-form
    ' lifecycle ownership; this early materialization is limited to the Setup controller surfaces.
    EnsureSetupControlForms
    LogSetupControlStartupStage "Setup Control form verification completed.", "Initializing Setup Control state."
    InitializeProgressControlState initialStatusText, False
    LogSetupControlStartupStage "Setup Control state initialization completed.", "Next call: OpenSetupControlRunningForm"
    LogSetupControlStartupStage "Running Setup Control form open starting.", "Form: " & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_RUNNING)
    OpenSetupControlRunningForm
    LogSetupControlStartupStage "Running Setup Control form open completed.", "Next call: SetupControlUpdateStatus"
    LogSetupControlStartupStage "Initial Setup Control status display starting.", "Status: " & mSetupControlStatusText
    SetupControlUpdateStatus mSetupControlStatusText, True, True
    LogSetupControlStartupStage "Initial Setup Control status display completed.", "Setup Control startup completed successfully."
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    LogSetupControlStartupStage "Setup Control startup failed.", "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, TRACE_LEVEL_ERROR
    EndSetupControl
    On Error GoTo 0
    Err.Raise savedErrNumber, "BeginSetupControl", "Could not start the setup pause/cancel controller." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

' Purpose: Initializes scan progress control, coordinating `InitializeProgressControlState` and `SetupControlUpdateStatus` as the implementation requires.
' Public role: Public module entry point. This procedure is listed in the module header as an intentional externally callable ADC entry point and may be invoked directly by Access/VBA callers.
' Parameters:
'   initialStatusText (Input Optional; String): Initial status text text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine. Default: "Scan is
'     starting.".
'   modalHeadingText (Input Optional; String): Input modal heading text for begin scan progress control; is passed to `InitializeProgressControlState`. Default: vbNullString.
'   modalLeadText (Input Optional; String): Input modal lead text for begin scan progress control; is passed to `InitializeProgressControlState`. Default: "Scan is running.".
'   allowCancel (Input Optional; Boolean): When True, enables or requests cancel for this call; when False, the routine follows its normal non-forced/default path. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub BeginScanProgressControl(Optional ByVal initialStatusText As String = "Scan is starting.", Optional ByVal modalHeadingText As String = vbNullString, Optional ByVal modalLeadText As String = "Scan is running.", Optional ByVal allowCancel As Boolean = True)
    'Shared modal progress-controller entry point. Default arguments preserve scan behavior; operational callers may supply non-scan wording and suppress cancellation.
    On Error GoTo ErrHandler
    EndSetupControl
    EnsureScanProgressForm
    InitializeProgressControlState initialStatusText, True, modalHeadingText, modalLeadText, allowCancel
    OpenScanProgressForm
    SetupControlUpdateStatus mSetupControlStatusText, True, True
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    EndSetupControl
    On Error GoTo 0
    Err.Raise savedErrNumber, "BeginScanProgressControl", "Could not start the modal progress controller." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub InitializeProgressControlState(ByVal initialStatusText As String, ByVal scanMode As Boolean, Optional ByVal modalHeadingText As String = vbNullString, Optional ByVal modalLeadText As String = vbNullString, Optional ByVal cancelEnabled As Boolean = True)
    mSetupControlActive = True
    mSetupControlScanMode = scanMode
    mSetupControlModalHeadingText = Trim$(modalHeadingText)
    mSetupControlModalLeadText = Trim$(modalLeadText)
    If scanMode And Len(mSetupControlModalLeadText) = 0 Then mSetupControlModalLeadText = "Scan is running."
    mSetupControlCancelEnabled = cancelEnabled
    mSetupControlPauseRequested = False
    mSetupControlCancelRequested = False
    mSetupControlResumeRequested = False
    mSetupControlCancelRequestedTimer = 0
    mSetupControlCancelObservedTimer = 0
    mSetupControlCancelObserved = False
    mSetupControlCancelObservedWorker = vbNullString
    mSetupControlCancelObservedItem = vbNullString
    mSetupControlStatusText = Trim$(initialStatusText)
    mSetupControlStartedTimer = Timer
    mSetupControlLastHeartbeatTimer = mSetupControlStartedTimer
    mSetupControlLastPumpTimer = mSetupControlStartedTimer
    mSetupControlLastDoEventsTimer = mSetupControlStartedTimer
    mSetupControlPumpActive = False
    mSetupControlCheckpointActive = False
    mSetupControlCheckpointDisplayActive = False
    mSetupControlPaused = False
    mSetupControlVisibilitySuspended = False
    mSetupControlLastWindowEnforcementTimer = 0
    mSetupControlLastDisplayedStatusText = vbNullString
    mSetupControlUpdateRequestCount = 0
    mSetupControlDisplayDispatchCount = 0
    mSetupControlHeartbeatDispatchCount = 0
    mSetupControlDuplicateDisplaySkipCount = 0
    mSetupControlPumpRequestCount = 0
    mSetupControlPumpCheckCounter = 0
    mSetupControlPumpCoalescedRequestCount = 0
    mSetupControlDoEventsDispatchCount = 0
    mSetupControlMaximumDoEventsGapSeconds = 0
End Sub

' Purpose: Closes setup control, coordinating `ManagedObjectPhysicalName` and `ElapsedSecondsFromTimer` as the implementation requires.
' Public role: Public module entry point. This procedure is listed in the module header as an intentional externally callable ADC entry point and may be invoked directly by Access/VBA callers.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub EndSetupControl()
    Dim finalDoEventsGapSeconds As Double
    On Error Resume Next
    SetupControlTraceCancellationCompleted
    CloseSetupControlForm ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_RUNNING)
    CloseSetupControlForm ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_PAUSED)
    CloseSetupControlForm ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SCAN_PROGRESS)
    mSetupControlActive = False
    mSetupControlScanMode = False
    mSetupControlModalHeadingText = vbNullString
    mSetupControlModalLeadText = vbNullString
    mSetupControlCancelEnabled = False
    mSetupControlPauseRequested = False
    mSetupControlCancelRequested = False
    mSetupControlResumeRequested = False
    mSetupControlCancelRequestedTimer = 0
    mSetupControlCancelObservedTimer = 0
    mSetupControlCancelObserved = False
    mSetupControlCancelObservedWorker = vbNullString
    mSetupControlCancelObservedItem = vbNullString
    mSetupControlStatusText = vbNullString
    finalDoEventsGapSeconds = ElapsedSecondsFromTimer(mSetupControlLastDoEventsTimer)
    If finalDoEventsGapSeconds > mSetupControlMaximumDoEventsGapSeconds Then mSetupControlMaximumDoEventsGapSeconds = finalDoEventsGapSeconds
    mSetupControlStartedTimer = 0
    mSetupControlLastHeartbeatTimer = 0
    mSetupControlLastPumpTimer = 0
    mSetupControlLastDoEventsTimer = 0
    mSetupControlPumpActive = False
    mSetupControlPumpCheckCounter = 0
    mSetupControlCheckpointActive = False
    mSetupControlCheckpointDisplayActive = False
    mSetupControlPaused = False
    mSetupControlVisibilitySuspended = False
    mSetupControlLastWindowEnforcementTimer = 0
    mSetupControlLastDisplayedStatusText = vbNullString
    ClearVisibleStatusText
    On Error GoTo 0
End Sub

' Purpose: Performs setup control update status by coordinating `SetupControlQualifiedStatusText` and `End` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   statusText (Input; String): Status/progress message supplied for the current operation; the routine forwards it to the applicable UI/progress and trace/reporting path.
'   updateStatusBar (Input Optional; Boolean): True updates status bar as part of the operation; False preserves the existing value/state. Default: True.
'   forceDisplay (Input Optional; Boolean): When True, enables or requests display for this call; when False, the routine follows its normal non-forced/default path. Default: False.
'   pumpAfterDisplay (Input Optional; Boolean): When True, enables or requests after display for this call; when False, the routine follows its normal non-forced/default path. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub SetupControlUpdateStatus(ByVal statusText As String, Optional ByVal updateStatusBar As Boolean = True, Optional ByVal forceDisplay As Boolean = False, Optional ByVal pumpAfterDisplay As Boolean = True)
    On Error Resume Next
    If mSetupControlActive = False Then Exit Sub
    statusText = SetupControlQualifiedStatusText(statusText)
    If Len(statusText) > 0 Then mSetupControlStatusText = statusText
    mSetupControlUpdateRequestCount = mSetupControlUpdateRequestCount + 1
    If mSetupControlCheckpointActive And mSetupControlPumpActive Then
        mSetupControlDuplicateDisplaySkipCount = mSetupControlDuplicateDisplaySkipCount + 1
        Exit Sub
    End If
    If forceDisplay = False Then
        If StrComp(mSetupControlStatusText, mSetupControlLastDisplayedStatusText, vbBinaryCompare) = 0 Then
            mSetupControlDuplicateDisplaySkipCount = mSetupControlDuplicateDisplaySkipCount + 1
            If pumpAfterDisplay Then SetupControlPump
            Exit Sub
        End If
    End If
    EnsureSetupControlCurrentFormVisible
    If updateStatusBar Then SetUIStatusBarText mSetupControlStatusText
    If mSetupControlScanMode Then
        UpdateSetupControlFormStatus ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SCAN_PROGRESS), SetupControlVisibleStatus(False)
    Else
        UpdateSetupControlFormStatus ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_RUNNING), SetupControlVisibleStatus(False)
        UpdateSetupControlFormStatus ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_PAUSED), SetupControlVisibleStatus(True)
    End If
    mSetupControlLastDisplayedStatusText = mSetupControlStatusText
    mSetupControlDisplayDispatchCount = mSetupControlDisplayDispatchCount + 1
    If pumpAfterDisplay Then SetupControlPump
    On Error GoTo 0
End Sub

' Purpose: Performs setup control pump by coordinating `End` and `ElapsedSecondsFromTimer` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   forcePump (Input Optional; Boolean): When True, enables or requests pump for this call; when False, the routine follows its normal non-forced/default path. Default: False.
'   suppressHeartbeat (Input Optional; Boolean): True suppresses heartbeat; False leaves the normal behavior enabled. Default: False.
' Returns: Boolean: True on the function's affirmative/success branch for setup control pump; False on its explicit negative, unavailable, rejected, or handled-failure branch.
Public Function SetupControlPump(Optional ByVal forcePump As Boolean = False, Optional ByVal suppressHeartbeat As Boolean = False) As Boolean
    'Coalesces direct hot-loop requests before timer, heartbeat, and form work. Time-gated common callers and explicit safe checkpoints force this boundary.
    On Error GoTo ErrHandler
    Dim doEventsGapSeconds As Double
    mSetupControlPumpRequestCount = mSetupControlPumpRequestCount + 1
    If mSetupControlActive = False Then
        SetupControlPump = True
        Exit Function
    End If
    If forcePump = False Then
        mSetupControlPumpCheckCounter = mSetupControlPumpCheckCounter + 1
        If mSetupControlPumpCheckCounter < SETUP_CONTROL_PUMP_REQUEST_CHECK_INTERVAL Then
            mSetupControlPumpCoalescedRequestCount = mSetupControlPumpCoalescedRequestCount + 1
            SetupControlPump = True
            Exit Function
        End If
        mSetupControlPumpCheckCounter = 0
        If ElapsedSecondsFromTimer(mSetupControlLastPumpTimer) < DEFAULT_SETUP_EVENT_PUMP_INTERVAL_SECONDS Then
            SetupControlPump = True
            Exit Function
        End If
    Else
        mSetupControlPumpCheckCounter = 0
    End If
    If mSetupControlPumpActive Then
        SetupControlPump = True
        Exit Function
    End If
    mSetupControlPumpActive = True
    doEventsGapSeconds = ElapsedSecondsFromTimer(mSetupControlLastDoEventsTimer)
    If doEventsGapSeconds > mSetupControlMaximumDoEventsGapSeconds Then mSetupControlMaximumDoEventsGapSeconds = doEventsGapSeconds
    mSetupControlLastPumpTimer = Timer
    mSetupControlLastDoEventsTimer = mSetupControlLastPumpTimer
    mSetupControlDoEventsDispatchCount = mSetupControlDoEventsDispatchCount + 1
    DoEvents
    EnsureSetupControlCurrentFormVisible
    If suppressHeartbeat = False Then SetupControlHeartbeat False
    SetupControlPump = True
    mSetupControlPumpActive = False
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    mSetupControlPumpActive = False
    Err.Raise savedErrNumber, "SetupControlPump", "The Setup Control event pump failed." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

' Purpose: Clears run-scoped setup display and event-pump measurements before a new trace run.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub ResetSetupControlPerformanceDiagnostics()
    'Clears run-scoped setup display and event-pump measurements before a new trace run.
    mSetupControlUpdateRequestCount = 0
    mSetupControlDisplayDispatchCount = 0
    mSetupControlHeartbeatDispatchCount = 0
    mSetupControlDuplicateDisplaySkipCount = 0
    mSetupControlPumpRequestCount = 0
    mSetupControlPumpCheckCounter = 0
    mSetupControlPumpCoalescedRequestCount = 0
    mSetupControlDoEventsDispatchCount = 0
    mSetupControlLastDoEventsTimer = 0
    mSetupControlMaximumDoEventsGapSeconds = 0
End Sub

' Purpose: Resolves and returns setup control performance summary from the supplied inputs and the authoritative ADC definitions/state used by this module.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: String: Returns formatted setup control performance summary text intended for tracing, diagnostics, or user-facing reporting; vbNullString is returned when there is nothing applicable to report.
Public Function SetupControlPerformanceSummary() As String
    SetupControlPerformanceSummary = _
        "UI status requests=" & CStr(mSetupControlUpdateRequestCount) & _
        "; UI status dispatches=" & CStr(mSetupControlDisplayDispatchCount) & _
        "; heartbeat dispatches=" & CStr(mSetupControlHeartbeatDispatchCount) & _
        "; duplicate display skips=" & CStr(mSetupControlDuplicateDisplaySkipCount) & _
        "; pump requests=" & CStr(mSetupControlPumpRequestCount) & _
        "; coalesced hot-loop requests=" & CStr(mSetupControlPumpCoalescedRequestCount) & _
        "; DoEvents dispatches=" & CStr(mSetupControlDoEventsDispatchCount) & _
        "; maximum DoEvents gap seconds=" & Format$(mSetupControlMaximumDoEventsGapSeconds, "0.000")
End Function

' Purpose: Processes the Setup Control form heartbeat callback so the running setup can pump UI state and observe pause/cancel requests.
' Public role: Form-code/event-expression public entry point. Access form code or an event property can invoke this routine by name; its Public visibility is part of that UI contract.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
' Returns: Boolean: True when the Setup Control heartbeat is processed successfully for the active setup; False when no active controllable setup exists or the callback handles an error.
Public Function SetupControlHeartbeat(Optional ByVal showMessage As Boolean = True) As Boolean
    On Error GoTo ErrHandler
    Dim heartbeatText As String
    Dim setupElapsedSeconds As Double
    Dim latestRegisteredStatusText As String
    If mSetupControlActive = False Then
        SetupControlHeartbeat = True
        Exit Function
    End If
    If ElapsedSecondsFromTimer(mSetupControlLastHeartbeatTimer) < DEFAULT_SETUP_HEARTBEAT_INTERVAL_SECONDS Then
        SetupControlHeartbeat = True
        Exit Function
    End If
    EnsureSetupControlCurrentFormVisible
    latestRegisteredStatusText = SharedLatestProgressStatusText()
    If Len(Trim$(latestRegisteredStatusText)) > 0 Then mSetupControlStatusText = SetupControlQualifiedStatusText(latestRegisteredStatusText)
    mSetupControlLastHeartbeatTimer = Timer
    setupElapsedSeconds = ElapsedSecondsFromTimer(mSetupControlStartedTimer)
    heartbeatText = SetupControlHeartbeatStatusText(setupElapsedSeconds)
    SetUIStatusBarText heartbeatText
    If mSetupControlScanMode Then
        UpdateSetupControlFormStatus ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SCAN_PROGRESS), SetupControlVisibleStatus(False)
    Else
        UpdateSetupControlFormStatus ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_RUNNING), SetupControlVisibleStatus(False)
        UpdateSetupControlFormStatus ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_PAUSED), SetupControlVisibleStatus(True)
    End If
    mSetupControlHeartbeatDispatchCount = mSetupControlHeartbeatDispatchCount + 1
    SetupControlHeartbeat = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If showMessage Then
        CoreShowErrorMessage "The Setup Control heartbeat could not update the running Setup display." & vbCrLf & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, vbCritical, DEFAULT_UI_MESSAGE_TITLE, "SetupControlHeartbeat", "SetupControl"
        SetupControlHeartbeat = False
        Exit Function
    End If
    Err.Raise savedErrNumber, "SetupControlHeartbeat", "The Setup Control heartbeat could not update the running Setup display." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function SetupControlHeartbeatStatusText(ByVal setupElapsedSeconds As Double) As String
    Dim statusText As String
    Dim heartbeatStateText As String
    statusText = mSetupControlStatusText
    If Len(statusText) = 0 Then
        If mSetupControlScanMode Then
            statusText = IIf(Len(mSetupControlModalLeadText) > 0, mSetupControlModalLeadText, "Scan is running.")
        Else
            statusText = "Setup is running."
        End If
    End If
    statusText = SetupControlQualifiedStatusText(statusText)
    If mSetupControlCancelRequested Then
        heartbeatStateText = "Cancel pending at the next safe checkpoint"
    ElseIf mSetupControlScanMode Then
        heartbeatStateText = IIf(mSetupControlCancelEnabled, "Scan still working", "Operation still working")
    ElseIf mSetupControlPauseRequested Then
        heartbeatStateText = "Pause pending at the next safe checkpoint"
    ElseIf CurrentProject.AllForms(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_PAUSED)).IsLoaded Then
        heartbeatStateText = "Setup paused"
    Else
        heartbeatStateText = "Setup still working"
    End If
    SetupControlHeartbeatStatusText = statusText & " | " & heartbeatStateText & " | " & SetupControlElapsedText(setupElapsedSeconds) & " elapsed"
End Function

Private Function SetupControlElapsedText(ByVal elapsedSeconds As Double) As String
    Dim totalSeconds As Long
    Dim elapsedMinutes As Long
    Dim remainingSeconds As Long
    totalSeconds = CLng(elapsedSeconds)
    elapsedMinutes = totalSeconds \ 60
    remainingSeconds = totalSeconds Mod 60
    If elapsedMinutes > 0 Then
        SetupControlElapsedText = CStr(elapsedMinutes) & "m " & Format$(remainingSeconds, "00") & "s"
    Else
        SetupControlElapsedText = CStr(remainingSeconds) & "s"
    End If
End Function

' Purpose: Reports whether setup control cancellation has been requested in the current control/state record.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: Boolean: True when setup control cancellation has been requested in current state; False otherwise.
'   path.
Public Function SetupControlCancellationRequested() As Boolean
    SetupControlCancellationRequested = (mSetupControlActive And mSetupControlCancelRequested)
End Function

' Purpose: Records a pause request from the Setup Control form when setup is active and pausing is permitted.
' Public role: Form-code/event-expression public entry point. Access form code or an event property can invoke this routine by name; its Public visibility is part of that UI contract.
' Parameters: None.
' Returns: Boolean: True when a pause request is accepted and recorded for the active setup; False when setup is inactive/ineligible or the callback handles an error.
Public Function SetupControlPause() As Boolean
    On Error GoTo ErrHandler
    If mSetupControlActive = False Or mSetupControlScanMode Then Exit Function
    mSetupControlPauseRequested = True
    mSetupControlResumeRequested = False
    SetupControlUpdateStatus mSetupControlStatusText, True, True
    SetupControlPause = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    CoreShowErrorMessage "The Setup Control could not process the Pause request." & vbCrLf & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, vbCritical, DEFAULT_UI_MESSAGE_TITLE, "SetupControlPause", "SetupControl"
    SetupControlPause = False
End Function

' Purpose: Records a cancellation request from the Setup Control form when setup is active and cancellation is permitted.
' Public role: Form-code/event-expression public entry point. Access form code or an event property can invoke this routine by name; its Public visibility is part of that UI contract.
' Parameters: None.
' Returns: Boolean: True when a cancellation request is accepted and recorded for the active setup; False when setup is inactive/ineligible or the callback handles an error.
Public Function SetupControlCancel() As Boolean
    On Error GoTo ErrHandler
    If mSetupControlActive = False Or mSetupControlCancelEnabled = False Then Exit Function
    If mSetupControlCancelRequested = False Then
        mSetupControlCancelRequested = True
        mSetupControlCancelRequestedTimer = Timer
        LogTraceMessage TRACE_LEVEL_WARNING, "Cancellation", "SetupControlCancel", "Cancellation requested by the user.", "Current work: " & IIf(Len(mSetupControlStatusText) = 0, "(not recorded)", mSetupControlStatusText), "Cancel requested", 0, vbNullString, vbNullString, vbNullString, "SetupControl", "CodeDb", "Cancellation"
    End If
    mSetupControlPauseRequested = False
    mSetupControlResumeRequested = False
    SetupControlUpdateStatus mSetupControlStatusText, True, True
    SetupControlCancel = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    CoreShowErrorMessage "The Setup Control could not process the Cancel request." & vbCrLf & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, vbCritical, DEFAULT_UI_MESSAGE_TITLE, "SetupControlCancel", "SetupControl"
    SetupControlCancel = False
End Function

' Purpose: Records a resume request from the Setup Control form and clears the pending pause request.
' Public role: Form-code/event-expression public entry point. Access form code or an event property can invoke this routine by name; its Public visibility is part of that UI contract.
' Parameters: None.
' Returns: Boolean: True when a resume request is accepted, recorded, and the pending pause request is cleared; False when setup is inactive/ineligible or the callback handles an error.
Public Function SetupControlResume() As Boolean
    On Error GoTo ErrHandler
    If mSetupControlActive = False Or mSetupControlScanMode Then Exit Function
    mSetupControlResumeRequested = True
    mSetupControlPauseRequested = False
    SetupControlResume = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    CoreShowErrorMessage "The Setup Control could not process the Resume request." & vbCrLf & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, vbCritical, DEFAULT_UI_MESSAGE_TITLE, "SetupControlResume", "SetupControl"
    SetupControlResume = False
End Function

' Purpose: Performs setup control work checkpoint by coordinating `End` and `SetupControlPump` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   checkpointText (Input; String): Input checkpoint text for setup control work checkpoint; is passed to `SetupControlTraceCancellationObserved`; is passed to `SetupControlCheckpoint`.
'   workerRoutine (Input Optional; String): Name of the worker procedure being invoked/tracked; used in dispatch, tracing, status, and error context. Default: vbNullString.
'   logicalItem (Input Optional; String): Stable logical item label associated with the worker operation for progress, tracing, or user-facing diagnostics. Default: vbNullString.
' Returns: Boolean result of setup control work checkpoint; True and False correspond to the affirmative and negative branches explicitly implemented by this function.
Public Function SetupControlWorkCheckpoint(ByVal checkpointText As String, Optional ByVal workerRoutine As String = vbNullString, Optional ByVal logicalItem As String = vbNullString) As Boolean
    'Lightweight safe-work checkpoint: pump pending UI events, honor cancel immediately, and enter the existing pause controller only when pause was actually requested.
    On Error GoTo ErrHandler
    If mSetupControlActive = False Then
        SetupControlWorkCheckpoint = True
        Exit Function
    End If
    SetupControlPump True
    If mSetupControlCancelRequested Then
        SetupControlTraceCancellationObserved checkpointText, workerRoutine, logicalItem
        Exit Function
    End If
    If mSetupControlPauseRequested And mSetupControlVisibilitySuspended Then
        SetupControlWorkCheckpoint = True
        Exit Function
    End If
    If mSetupControlPauseRequested Then
        SetupControlWorkCheckpoint = SetupControlCheckpoint(checkpointText)
        If SetupControlWorkCheckpoint = False And mSetupControlCancelRequested Then SetupControlTraceCancellationObserved checkpointText, workerRoutine, logicalItem
        Exit Function
    End If
    SetupControlWorkCheckpoint = True
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "SetupControlWorkCheckpoint", "The setup/repair pause-cancel controller failed at an internal safe-work checkpoint." & vbCrLf & "Checkpoint: " & checkpointText & vbCrLf & "Worker: " & workerRoutine & vbCrLf & "Item: " & logicalItem & vbCrLf & "Err.Description: " & Err.Description
End Function

Private Sub SetupControlTraceCancellationObserved(ByVal checkpointText As String, ByVal workerRoutine As String, ByVal logicalItem As String)
    If mSetupControlCancelRequested = False Then Exit Sub
    If mSetupControlCancelObserved Then Exit Sub
    mSetupControlCancelObserved = True
    mSetupControlCancelObservedTimer = Timer
    mSetupControlCancelObservedWorker = Trim$(workerRoutine)
    mSetupControlCancelObservedItem = Trim$(logicalItem)
    LogTraceMessage TRACE_LEVEL_WARNING, "Cancellation", "SetupControlWorkCheckpoint", "Cancellation observed at a safe internal work boundary.", "Checkpoint: " & IIf(Len(Trim$(checkpointText)) = 0, "(not recorded)", checkpointText) & vbCrLf & "Worker: " & IIf(Len(mSetupControlCancelObservedWorker) = 0, "(not recorded)", mSetupControlCancelObservedWorker) & vbCrLf & "Logical item: " & IIf(Len(mSetupControlCancelObservedItem) = 0, "(not recorded)", mSetupControlCancelObservedItem) & vbCrLf & "Request-to-observation latency: " & Format$(ElapsedSecondsFromTimer(mSetupControlCancelRequestedTimer), "0.000") & " seconds", "Cancel observed", 0, vbNullString, vbNullString, vbNullString, "SetupControl", "CodeDb", "Cancellation"
End Sub

Private Sub SetupControlTraceCancellationCompleted()
    Dim detailText As String
    If mSetupControlCancelRequested = False Then Exit Sub
    detailText = "Request-to-controller-exit latency: " & Format$(ElapsedSecondsFromTimer(mSetupControlCancelRequestedTimer), "0.000") & " seconds"
    If mSetupControlCancelObserved Then
        detailText = detailText & vbCrLf & "Observation-to-controller-exit latency: " & Format$(ElapsedSecondsFromTimer(mSetupControlCancelObservedTimer), "0.000") & " seconds"
        detailText = detailText & vbCrLf & "Observed worker: " & IIf(Len(mSetupControlCancelObservedWorker) = 0, "(not recorded)", mSetupControlCancelObservedWorker)
        detailText = detailText & vbCrLf & "Observed item: " & IIf(Len(mSetupControlCancelObservedItem) = 0, "(not recorded)", mSetupControlCancelObservedItem)
    Else
        detailText = detailText & vbCrLf & "Cancellation was not observed by a worker checkpoint before the controller ended."
    End If
    LogTraceMessage TRACE_LEVEL_WARNING, "Cancellation", "EndSetupControl", "Cancellation controller cleanup completed.", detailText, "Cancel completed", 0, vbNullString, vbNullString, vbNullString, "SetupControl", "CodeDb", "Cancellation"
End Sub

' Purpose: Performs setup control checkpoint by coordinating `End` and `LogSetupControlCheckpointStage` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   checkpointText (Input; String): Input checkpoint text for setup control checkpoint; is passed to `LogSetupControlCheckpointStage`; is assigned to `If Len(checkpointText) > 0 Then mSetupControlStatusText`.
' Returns: Boolean result of setup control checkpoint; True and False correspond to the affirmative and negative branches explicitly implemented by this function.
Public Function SetupControlCheckpoint(ByVal checkpointText As String) As Boolean
    On Error GoTo ErrHandler
    Dim checkpointTimer As Single
    Dim stageTimer As Single
    Dim stageElapsedSeconds As Double
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    If mSetupControlActive = False Then
        SetupControlCheckpoint = True
        Exit Function
    End If
    If mSetupControlCheckpointActive Then
        SetupControlCheckpoint = Not mSetupControlCancelRequested
        Exit Function
    End If
    checkpointText = Trim$(checkpointText)
    checkpointTimer = Timer
    mSetupControlCheckpointActive = True
    LogSetupControlCheckpointStage "Setup Control checkpoint starting.", checkpointText, "The part identity is recorded before display and event-pump work."
    If Len(checkpointText) > 0 Then mSetupControlStatusText = checkpointText
    stageTimer = Timer
    mSetupControlCheckpointDisplayActive = True
    SetupControlUpdateStatus mSetupControlStatusText, True, False, False
    mSetupControlCheckpointDisplayActive = False
    stageElapsedSeconds = ElapsedSecondsFromTimer(stageTimer)
    If stageElapsedSeconds >= SETUP_CONTROL_SLOW_CHECKPOINT_SECONDS Then LogSetupControlCheckpointStage "Setup Control checkpoint status display was slow.", checkpointText, "Elapsed: " & FormattedElapsedSeconds(stageElapsedSeconds), TRACE_LEVEL_WARNING
    stageTimer = Timer
    SetupControlPump True, True
    stageElapsedSeconds = ElapsedSecondsFromTimer(stageTimer)
    If stageElapsedSeconds >= SETUP_CONTROL_SLOW_CHECKPOINT_SECONDS Then LogSetupControlCheckpointStage "Setup Control checkpoint event pump was slow.", checkpointText, "Elapsed: " & FormattedElapsedSeconds(stageElapsedSeconds), TRACE_LEVEL_WARNING
    If mSetupControlCancelRequested Then
        SetupControlTraceCancellationObserved checkpointText, "SetupControlCheckpoint", vbNullString
        GoTo ExitPoint
    End If
    'A Pause request remains pending while D2 owns a Setup Control form for physical replacement; entering the paused-form loop here could reopen the exact form that rollback/build requires to stay closed.
    If mSetupControlPauseRequested And mSetupControlVisibilitySuspended Then
        SetupControlCheckpoint = True
        GoTo ExitPoint
    End If
    If mSetupControlPauseRequested Then
        mSetupControlPauseRequested = False
        mSetupControlResumeRequested = False
        OpenSetupControlPausedForm
        Do While mSetupControlActive And mSetupControlResumeRequested = False And mSetupControlCancelRequested = False
            SetupControlPump True
            UpdateSetupControlFormStatus ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_PAUSED), SetupControlVisibleStatus(True)
            ReportErrorSleep 50
        Loop
        CloseSetupControlForm ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_PAUSED)
        If mSetupControlCancelRequested Then
            SetupControlTraceCancellationObserved checkpointText, "SetupControlCheckpoint", "Paused"
            GoTo ExitPoint
        End If
        mSetupControlResumeRequested = False
        OpenSetupControlRunningForm
        SetupControlUpdateStatus mSetupControlStatusText, True, False, False
    End If
    SetupControlCheckpoint = True
ExitPoint:
    mSetupControlCheckpointDisplayActive = False
    mSetupControlCheckpointActive = False
    LogSetupControlCheckpointStage "Setup Control checkpoint completed.", checkpointText, "Result: " & IIf(SetupControlCheckpoint, "continue", "cancel") & " | Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(checkpointTimer))
    Exit Function
ErrHandler:
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    mSetupControlCheckpointDisplayActive = False
    mSetupControlCheckpointActive = False
    LogSetupControlCheckpointStage "Setup Control checkpoint failed.", checkpointText, "Err.Number: " & CStr(savedErrNumber) & " | Err.Description: " & savedErrDescription, TRACE_LEVEL_ERROR
    Err.Raise savedErrNumber, "SetupControlCheckpoint", "The setup pause/cancel controller failed at a safe checkpoint." & vbCrLf & "Checkpoint: " & checkpointText & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function SetupControlVisibleStatus(ByVal pausedForm As Boolean) As String
    Dim statusText As String
    Dim modalLeadText As String
    statusText = mSetupControlStatusText
    modalLeadText = mSetupControlModalLeadText
    If Len(modalLeadText) = 0 Then modalLeadText = "Scan is running."
    If Len(statusText) = 0 Then statusText = IIf(mSetupControlScanMode, modalLeadText, "Setup is running.")
    If mSetupControlScanMode Then
        If mSetupControlCancelRequested Then
            SetupControlVisibleStatus = "Cancel requested. The scan will stop at the next safe-work checkpoint." & vbCrLf & vbCrLf & statusText
        Else
            SetupControlVisibleStatus = modalLeadText & vbCrLf & vbCrLf & statusText
        End If
    ElseIf mSetupControlCancelRequested Then
        SetupControlVisibleStatus = "Cancel requested. Setup will stop at the next rollback-safe checkpoint." & vbCrLf & vbCrLf & statusText
    ElseIf mSetupControlPauseRequested Then
        SetupControlVisibleStatus = "Pause requested. Setup will pause at the next safe checkpoint." & vbCrLf & vbCrLf & statusText
    ElseIf pausedForm Then
        SetupControlVisibleStatus = "Setup is paused at a safe checkpoint." & vbCrLf & vbCrLf & statusText
    Else
        SetupControlVisibleStatus = "Setup is running." & vbCrLf & vbCrLf & statusText
    End If
End Function

Private Sub EnsureSetupControlForms()
    ' PRE-D BOOTSTRAP EXCEPTION: Setup Control and Setup Paused are canonical generated forms, but Setup must be controllable before D1/D2 can run. This procedure may therefore create/repair only these controller
    ' forms during Preflight; D2 later verifies/converges them with the complete generated-form domain.
    On Error GoTo ErrHandler
    Dim namingSnapshotOwned As Boolean
    Dim generatedFormDefinition As Variant
    InvalidateGeneratedFormSignatureCurrentCache
    LogSetupControlStartupStage "Setup Control naming snapshot starting.", "The existing naming snapshot authority will cache Settings and MAC physical-name resolution while generated-form definitions are initialized."
    If NamingSnapshotIsReady() Then
        LogSetupControlStartupStage "Setup Control naming snapshot reused.", "An already-active naming snapshot remains owned by its caller; Setup Control will borrow it without invalidating it."
    Else
        BeginNamingSnapshot
        namingSnapshotOwned = True
        LogSetupControlStartupStage "Setup Control naming snapshot completed.", "Next stage: initialize the existing generated-form definition authority once."
    End If
    LogSetupControlStartupStage "Generated-form definition initialization starting.", "Authority: GeneratedFormCatalogDefinition"
    generatedFormDefinition = GeneratedFormCatalogDefinition(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_RUNNING))
    If IsArray(generatedFormDefinition) = False Then Err.Raise vbObjectError + 9199, "EnsureSetupControlForms", "The authoritative generated-form definition was not found." & vbCrLf & "Form: " & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_RUNNING)
    LogSetupControlStartupStage "Generated-form definition initialization completed.", "The existing static generated-form definition collection is ready for both Setup Control forms."
    LogSetupControlStartupStage "Running Setup Control form verification starting.", "Form: " & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_RUNNING)
    EnsureSetupControlForm ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_RUNNING), SETUP_CONTROL_RUNNING_FORM_TAG, True
    LogSetupControlStartupStage "Running Setup Control form verification completed.", "Next stage: stamp the running-form build signature."
    LogSetupControlStartupStage "Running Setup Control form signature stamp starting.", "Form: " & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_RUNNING)
    StampGeneratedFormBuildSignature ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_RUNNING)
    LogSetupControlStartupStage "Running Setup Control form signature stamp completed.", "Next stage: verify the paused Setup Control form."
    LogSetupControlStartupStage "Paused Setup Control form verification starting.", "Form: " & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_PAUSED)
    EnsureSetupControlForm ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_PAUSED), SETUP_CONTROL_PAUSED_FORM_TAG, False
    LogSetupControlStartupStage "Paused Setup Control form verification completed.", "Next stage: stamp the paused-form build signature."
    LogSetupControlStartupStage "Paused Setup Control form signature stamp starting.", "Form: " & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_PAUSED)
    StampGeneratedFormBuildSignature ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_PAUSED)
    LogSetupControlStartupStage "Paused Setup Control form signature stamp completed.", "Both Setup Control forms are verified and stamped."
    If namingSnapshotOwned Then
        LogSetupControlStartupStage "Setup Control naming snapshot cleanup starting.", "Snapshot ownership: EnsureSetupControlForms"
        InvalidateNamingSnapshot
        namingSnapshotOwned = False
        LogSetupControlStartupStage "Setup Control naming snapshot cleanup completed.", "Setup Control form verification is returning to BeginSetupControl."
    Else
        LogSetupControlStartupStage "Setup Control naming snapshot cleanup skipped.", "The naming snapshot is caller-owned and remains active."
    End If
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If namingSnapshotOwned Then InvalidateNamingSnapshot
    On Error GoTo 0
    Err.Raise savedErrNumber, "EnsureSetupControlForms", "Could not create or verify the Setup Control forms." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub EnsureScanProgressForm()
    On Error GoTo ErrHandler
    Dim namingSnapshotOwned As Boolean
    Dim generatedFormDefinition As Variant
    InvalidateGeneratedFormSignatureCurrentCache
    If NamingSnapshotIsReady() = False Then
        BeginNamingSnapshot
        namingSnapshotOwned = True
    End If
    generatedFormDefinition = GeneratedFormCatalogDefinition(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SCAN_PROGRESS))
    If IsArray(generatedFormDefinition) = False Then Err.Raise vbObjectError + 9200, "EnsureScanProgressForm", "The authoritative generated-form definition was not found." & vbCrLf & "Form: " & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SCAN_PROGRESS)
    EnsureSetupControlForm ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SCAN_PROGRESS), SCAN_PROGRESS_FORM_TAG, True, True
    StampGeneratedFormBuildSignature ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SCAN_PROGRESS)
    If namingSnapshotOwned Then InvalidateNamingSnapshot
    namingSnapshotOwned = False
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If namingSnapshotOwned Then InvalidateNamingSnapshot
    On Error GoTo 0
    Err.Raise savedErrNumber, "EnsureScanProgressForm", "Could not create or verify the Scan Progress form." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Function SetupControlFormDesignDefinition(ByVal expectedTag As String, ByVal runningForm As Boolean, Optional ByVal scanForm As Boolean = False) As Variant
    'Canonical form-level progress-controller design consumed by create and repair paths.
    'Setup controllers stay popup but non-modal so generated Design-view work can become active; the shared operation progress controller is modal because its scan and selected operational work do not require
    'Design-view activation.
    SetupControlFormDesignDefinition = Array( _
        IIf(scanForm, "Scan Progress", IIf(runningForm, "Setup Running", "Setup Paused")), expectedTag, 7200, 3550, 0, False, False, False, 0, True, True, True, scanForm, 3, 0, False, _
        DEFAULT_SETUP_HEARTBEAT_TIMER_INTERVAL_MS, "=SetupControlHeartbeat()")
End Function

Private Function SetupControlControlDefinitions(ByVal runningForm As Boolean, Optional ByVal scanForm As Boolean = False) As Collection
    'Canonical control-level progress-controller design consumed by create, required-control verification, and repair.
    ' CANONICAL AUTHORITY: rows owns the in-memory definition set constructed by SetupControlControlDefinitions; downstream consumers derive views from this set rather than create a competing authority.
    Dim rows As New Collection
    rows.Add Array("Heading", vbNullString, acLabel, DEFAULT_UI_FORM_LEFT_MARGIN, DEFAULT_UI_FORM_TOP_MARGIN, 6500, 350, IIf(scanForm, DEFAULT_UTILITY_SHORT_NAME & " scan is running.", IIf(runningForm, DEFAULT_UTILITY_SHORT_NAME & " setup is running.", DEFAULT_UTILITY_SHORT_NAME & " setup is paused.")), DEFAULT_UI_LARGE_LABEL_FONT_SIZE, True, vbNullString, -1, False, True)
    rows.Add Array("Status", SETUP_CONTROL_STATUS_CONTROL, acLabel, DEFAULT_UI_FORM_LEFT_MARGIN, DEFAULT_UI_FORM_TOP_MARGIN + 350 + DEFAULT_UI_SECTION_TITLE_CONTENT_GAP, 6500, 1600, IIf(scanForm, "Scan is running.", IIf(runningForm, "Setup is running.", "Setup is paused at a safe checkpoint.")), DEFAULT_UI_READING_FONT_SIZE, False, vbNullString, -1, False, True)
    rows.Add Array("Elapsed", SETUP_CONTROL_ELAPSED_CONTROL, acLabel, DEFAULT_UI_FORM_LEFT_MARGIN, DEFAULT_UI_FORM_TOP_MARGIN + 350 + DEFAULT_UI_SECTION_TITLE_CONTENT_GAP + 1600 + DEFAULT_UI_ROW_GAP, 3000, 300, "Elapsed: 0s", DEFAULT_UI_READING_FONT_SIZE, True, vbNullString, -1, False, True)
    If scanForm Then
        rows.Add Array("Cancel", "cmdCancel", acCommandButton, 5650, DEFAULT_UI_FORM_TOP_MARGIN + 350 + DEFAULT_UI_SECTION_TITLE_CONTENT_GAP + 1600 + DEFAULT_UI_ROW_GAP + 300 + DEFAULT_UI_BUTTON_ROW_GAP, 1200, 450, "Cancel", DEFAULT_UI_FONT_SIZE, False, "=SetupControlCancel()", 0, False, False)
    ElseIf runningForm Then
        rows.Add Array("PrimaryAction", "cmdPause", acCommandButton, 4300, DEFAULT_UI_FORM_TOP_MARGIN + 350 + DEFAULT_UI_SECTION_TITLE_CONTENT_GAP + 1600 + DEFAULT_UI_ROW_GAP + 300 + DEFAULT_UI_BUTTON_ROW_GAP, 1200, 450, "Pause", DEFAULT_UI_FONT_SIZE, False, "=SetupControlPause()", 0, False, True)
        rows.Add Array("Cancel", "cmdCancel", acCommandButton, 5650, DEFAULT_UI_FORM_TOP_MARGIN + 350 + DEFAULT_UI_SECTION_TITLE_CONTENT_GAP + 1600 + DEFAULT_UI_ROW_GAP + 300 + DEFAULT_UI_BUTTON_ROW_GAP, 1200, 450, "Cancel", DEFAULT_UI_FONT_SIZE, False, "=SetupControlCancel()", 1, False, True)
    Else
        rows.Add Array("PrimaryAction", "cmdResume", acCommandButton, 4300, DEFAULT_UI_FORM_TOP_MARGIN + 350 + DEFAULT_UI_SECTION_TITLE_CONTENT_GAP + 1600 + DEFAULT_UI_ROW_GAP + 300 + DEFAULT_UI_BUTTON_ROW_GAP, 1200, 450, "Resume", DEFAULT_UI_FONT_SIZE, False, "=SetupControlResume()", 0, True, True)
        rows.Add Array("Cancel", "cmdCancel", acCommandButton, 5650, DEFAULT_UI_FORM_TOP_MARGIN + 350 + DEFAULT_UI_SECTION_TITLE_CONTENT_GAP + 1600 + DEFAULT_UI_ROW_GAP + 300 + DEFAULT_UI_BUTTON_ROW_GAP, 1200, 450, "Cancel", DEFAULT_UI_FONT_SIZE, False, "=SetupControlCancel()", 1, False, True)
    End If
    Set SetupControlControlDefinitions = rows
End Function

Private Sub ApplySetupControlFormDefinition(ByVal frm As Form, ByVal formDefinition As Variant)
    frm.Caption = CStr(formDefinition(0))
    frm.Tag = CStr(formDefinition(1))
    frm.Width = CLng(formDefinition(2))
    frm.Section(acDetail).Height = CLng(formDefinition(3))
    frm.defaultView = CLng(formDefinition(4))
    frm.RecordSelectors = CBool(formDefinition(5))
    frm.NavigationButtons = CBool(formDefinition(6))
    frm.DividingLines = CBool(formDefinition(7))
    frm.ScrollBars = CLng(formDefinition(8))
    frm.AutoCenter = CBool(formDefinition(9))
    frm.AutoResize = CBool(formDefinition(10))
    frm.PopUp = CBool(formDefinition(11))
    frm.Modal = CBool(formDefinition(12))
    frm.BorderStyle = CLng(formDefinition(13))
    frm.MinMaxButtons = CLng(formDefinition(14))
    frm.CloseButton = CBool(formDefinition(15))
    frm.TimerInterval = CLng(formDefinition(16))
    frm.OnTimer = CStr(formDefinition(17))
End Sub

Private Function SetupControlControlForDefinition(ByVal frm As Form, ByVal controlDefinition As Variant) As control
    On Error GoTo NotFound
    Dim ctl As control
    Dim controlName As String
    controlName = CStr(controlDefinition(1))
    If Len(controlName) > 0 Then
        Set SetupControlControlForDefinition = frm.Controls(controlName)
        Exit Function
    End If
    For Each ctl In frm.Controls
        If ctl.controlType = CLng(controlDefinition(2)) Then
            If StrComp(ctl.Name, SETUP_CONTROL_STATUS_CONTROL, vbTextCompare) <> 0 And StrComp(ctl.Name, SETUP_CONTROL_ELAPSED_CONTROL, vbTextCompare) <> 0 Then
                Set SetupControlControlForDefinition = ctl
                Exit Function
            End If
        End If
    Next ctl
    Exit Function
NotFound:
    Err.Clear
End Function

Private Sub ApplySetupControlControlDefinition(ByVal ctl As control, ByVal controlDefinition As Variant)
    ctl.Left = CLng(controlDefinition(3))
    ctl.Top = CLng(controlDefinition(4))
    ctl.Width = CLng(controlDefinition(5))
    ctl.Height = CLng(controlDefinition(6))
    ctl.Caption = CStr(controlDefinition(7))
    ctl.FontName = DEFAULT_UI_FONT_NAME
    ctl.fontSize = CLng(controlDefinition(8))
    ctl.fontBold = CBool(controlDefinition(9))
    ctl.Visible = CBool(controlDefinition(13))
    If ctl.controlType = acLabel Then ctl.TextAlign = 1
    If ctl.controlType = acCommandButton Then
        ctl.OnClick = CStr(controlDefinition(10))
        If CLng(controlDefinition(11)) >= 0 Then ctl.TabIndex = CLng(controlDefinition(11))
        ctl.Default = CBool(controlDefinition(12))
    End If
End Sub

' Purpose: Ensures setup control form for form name, expected tag.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   formName (Input; String): Physical Access form name identifying the form to create, open, configure, validate, or inspect.
'   expectedTag (Input; String): Exact expected Ribbon/control tag used to validate that the generated UI element is wired to the intended managed command.
'   runningForm (Input; Boolean): Boolean state describing whether running form applies to this invocation; True selects the named condition/path and False selects its complementary/default path.
'   scanForm (Input Optional; Boolean): When True, enables or requests form for this call; when False, the routine follows its normal non-forced/default path. Default: False.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub EnsureSetupControlForm(ByVal formName As String, ByVal expectedTag As String, ByVal runningForm As Boolean, Optional ByVal scanForm As Boolean = False)
    On Error GoTo ErrHandler
    Dim temporaryFormName As String
    Dim frm As Form
    Dim ctl As control
    Dim controlDefinition As Variant
    Dim controlDefinitions As Collection
    Set controlDefinitions = SetupControlControlDefinitions(runningForm, scanForm)
    If SharedDao_AccessObjectExists(GetHostDb(), acForm, formName) Then
        DoCmd.OpenForm formName, acDesign, , , , acHidden
        If StrComp(CStr(Nz(forms(formName).Tag, vbNullString)), expectedTag, vbBinaryCompare) <> 0 Then
            DoCmd.Close acForm, formName, acSaveNo
            Err.Raise vbObjectError + 9140, "EnsureSetupControlForm", "A required progress-controller form name is already used by a form that is not owned by this workflow." & vbCrLf & "Form: " & formName
        End If
        If SetupControlFormHasRequiredControls(forms(formName), controlDefinitions) Then
            RepairSetupControlFormDesign forms(formName), expectedTag, runningForm, scanForm, controlDefinitions
            DoCmd.Close acForm, formName, acSaveYes
            Application.SetHiddenAttribute acForm, formName, True
            CompleteManagedSetupControlFormDescriptionIfAllowed formName
            Exit Sub
        End If
        DoCmd.Close acForm, formName, acSaveNo
        DoCmd.DeleteObject acForm, formName
    End If
    Set frm = CreateForm
    temporaryFormName = frm.Name
    With frm.DefaultControl(acLabel)
        .FontName = DEFAULT_UI_FONT_NAME
        .fontSize = DEFAULT_UI_FONT_SIZE
        .TextAlign = 1
    End With
    With frm.DefaultControl(acCommandButton)
        .FontName = DEFAULT_UI_FONT_NAME
        .fontSize = DEFAULT_UI_FONT_SIZE
    End With
    ApplySetupControlFormDefinition frm, SetupControlFormDesignDefinition(expectedTag, runningForm, scanForm)
    For Each controlDefinition In controlDefinitions
        Set ctl = CreateControl(temporaryFormName, CLng(controlDefinition(2)), acDetail, vbNullString, vbNullString, CLng(controlDefinition(3)), CLng(controlDefinition(4)), CLng(controlDefinition(5)), CLng(controlDefinition(6)))
        If Len(CStr(controlDefinition(1))) > 0 Then ctl.Name = CStr(controlDefinition(1))
        ApplySetupControlControlDefinition ctl, controlDefinition
    Next controlDefinition
    DoCmd.Close acForm, temporaryFormName, acSaveYes
    DoCmd.Rename formName, acForm, temporaryFormName
    Application.SetHiddenAttribute acForm, formName, True
    CompleteManagedSetupControlFormDescriptionIfAllowed formName
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Len(temporaryFormName) > 0 Then DoCmd.Close acForm, temporaryFormName, acSaveNo
    On Error GoTo 0
    Err.Raise savedErrNumber, "EnsureSetupControlForm", "Could not create or verify a progress-controller form." & vbCrLf & "Form: " & formName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub CompleteManagedSetupControlFormDescriptionIfAllowed(ByVal formName As String)
    'Progress-controller forms may be bootstrapped before or outside numbered Setup. During D2/D3 the central Description guard keeps this completion silent so G1 remains the numbered-Setup owner.
    Dim definitionData As Variant
    Dim descriptionChanged As Variant
    If SetupDomainMutationWritesAllowedForCurrentPart(SETUP_DOMAIN_G1_DESCRIPTIONS) = False Then Exit Sub
    If TryGetManagedObjectDefinition(formName, ADCAT_OBJECT_TYPE_FORM, definitionData) = False Then Exit Sub
    descriptionChanged = False
    If ApplyAndVerifyManagedObjectDescriptionForStandaloneRepair(formName, ADCAT_OBJECT_TYPE_FORM, descriptionChanged) = False Then Err.Raise vbObjectError + 10129, "CompleteManagedSetupControlFormDescriptionIfAllowed", "Managed progress-controller form Description reconciliation returned False." & vbCrLf & "Form: " & formName
End Sub

Private Function SetupControlFormHasRequiredControls(ByVal frm As Form, ByVal controlDefinitions As Collection) As Boolean
    On Error GoTo NotComplete
    Dim controlDefinition As Variant
    Dim ctl As control
    For Each controlDefinition In controlDefinitions
        Set ctl = SetupControlControlForDefinition(frm, controlDefinition)
        If ctl Is Nothing Then Exit Function
        If ctl.controlType <> CLng(controlDefinition(2)) Then Exit Function
        Set ctl = Nothing
    Next controlDefinition
    SetupControlFormHasRequiredControls = True
    Exit Function
NotComplete:
    Err.Clear
End Function

Private Sub RepairSetupControlFormDesign(ByVal frm As Form, ByVal expectedTag As String, ByVal runningForm As Boolean, ByVal scanForm As Boolean, ByVal controlDefinitions As Collection)
    Dim controlDefinition As Variant
    Dim ctl As control
    ApplySetupControlFormDefinition frm, SetupControlFormDesignDefinition(expectedTag, runningForm, scanForm)
    For Each controlDefinition In controlDefinitions
        Set ctl = SetupControlControlForDefinition(frm, controlDefinition)
        If ctl Is Nothing Then Err.Raise vbObjectError + 9141, "RepairSetupControlFormDesign", "A canonical progress-controller form control is missing during repair." & vbCrLf & "Form: " & frm.Name & vbCrLf & "Control role: " & CStr(controlDefinition(0))
        ApplySetupControlControlDefinition ctl, controlDefinition
        Set ctl = Nothing
    Next controlDefinition
End Sub

Private Sub EnsureSetupControlCurrentFormVisible()
    On Error GoTo SafeExit
    Dim currentFormName As String
    If mSetupControlActive = False Then Exit Sub
    If mSetupControlVisibilitySuspended Then Exit Sub
    If SetupHostWindowIsMinimized Then Exit Sub
    If mSetupControlScanMode Then
        If CurrentProject.AllForms(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_RUNNING)).IsLoaded Then DoCmd.Close acForm, ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_RUNNING), acSaveNo
        If CurrentProject.AllForms(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_PAUSED)).IsLoaded Then DoCmd.Close acForm, ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_PAUSED), acSaveNo
        currentFormName = ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SCAN_PROGRESS)
        If CurrentProject.AllForms(currentFormName).IsLoaded = False Then OpenScanProgressForm
    ElseIf mSetupControlPaused Then
        If CurrentProject.AllForms(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_RUNNING)).IsLoaded Then DoCmd.Close acForm, ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_RUNNING), acSaveNo
        currentFormName = ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_PAUSED)
        If CurrentProject.AllForms(currentFormName).IsLoaded = False Then OpenSetupControlPausedForm
    Else
        If CurrentProject.AllForms(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_PAUSED)).IsLoaded Then DoCmd.Close acForm, ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_PAUSED), acSaveNo
        currentFormName = ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_RUNNING)
        If CurrentProject.AllForms(currentFormName).IsLoaded = False Then OpenSetupControlRunningForm
    End If
    If CurrentProject.AllForms(currentFormName).IsLoaded = False Then Exit Sub
    If forms(currentFormName).Visible = False Then
        ShowSetupControlFormWindow currentFormName
    ElseIf mSetupControlLastWindowEnforcementTimer = 0 Or ElapsedSecondsFromTimer(mSetupControlLastWindowEnforcementTimer) >= DEFAULT_SETUP_CONTROLLER_WINDOW_ENFORCEMENT_INTERVAL_SECONDS Then
        ReinforceSetupControlFormWindow currentFormName
    End If
SafeExit:
    Err.Clear
End Sub

Private Sub ShowSetupControlFormWindow(ByVal formName As String)
    On Error Resume Next
    If SetupHostWindowIsMinimized Then Exit Sub
    If CurrentProject.AllForms(formName).IsLoaded = False Then Exit Sub
    forms(formName).Visible = True
    ReinforceSetupControlFormWindow formName
    forms(formName).Repaint
    On Error GoTo 0
End Sub

Private Sub ReinforceSetupControlFormWindow(ByVal formName As String)
    On Error Resume Next
    Dim formWindowHandle As LongPtr
    If SetupHostWindowIsMinimized Then Exit Sub
    If CurrentProject.AllForms(formName).IsLoaded = False Then Exit Sub
    formWindowHandle = CLngPtr(forms(formName).hWnd)
    If formWindowHandle <> 0 Then SetupSetWindowPos formWindowHandle, 0, 0, 0, 0, 0, SETUP_WINDOW_POS_NO_MOVE Or SETUP_WINDOW_POS_NO_SIZE Or SETUP_WINDOW_POS_NO_ACTIVATE Or SETUP_WINDOW_POS_SHOW_WINDOW
    mSetupControlLastWindowEnforcementTimer = Timer
    On Error GoTo 0
End Sub

Private Sub OpenScanProgressForm()
    Dim ctl As control
    Dim headingText As String
    mSetupControlPaused = False
    If CurrentProject.AllForms(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_RUNNING)).IsLoaded Then DoCmd.Close acForm, ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_RUNNING), acSaveNo
    If CurrentProject.AllForms(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_PAUSED)).IsLoaded Then DoCmd.Close acForm, ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_PAUSED), acSaveNo
    ' The canonical Scan Progress design keeps cmdCancel hidden and non-default so the modal form cannot focus it before operation-specific cancellation is configured.
    If CurrentProject.AllForms(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SCAN_PROGRESS)).IsLoaded = False Then DoCmd.OpenForm ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SCAN_PROGRESS), acNormal, , , , acHidden
    headingText = mSetupControlModalHeadingText
    If Len(headingText) = 0 Then headingText = DEFAULT_UTILITY_SHORT_NAME & " scan is running."
    forms(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SCAN_PROGRESS)).Caption = IIf(Len(mSetupControlModalHeadingText) > 0, mSetupControlModalHeadingText, "Scan Progress")
    For Each ctl In forms(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SCAN_PROGRESS)).Controls
        If ctl.controlType = acLabel Then
            If StrComp(ctl.Name, SETUP_CONTROL_STATUS_CONTROL, vbTextCompare) <> 0 And StrComp(ctl.Name, SETUP_CONTROL_ELAPSED_CONTROL, vbTextCompare) <> 0 Then
                ctl.Caption = headingText
                Exit For
            End If
        End If
    Next ctl
    With forms(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SCAN_PROGRESS)).Controls("cmdCancel")
        .enabled = mSetupControlCancelEnabled
        .Default = mSetupControlCancelEnabled
        .Visible = mSetupControlCancelEnabled
    End With
    ShowSetupControlFormWindow ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SCAN_PROGRESS)
End Sub

Private Sub OpenSetupControlRunningForm()
    mSetupControlPaused = False
    If CurrentProject.AllForms(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_PAUSED)).IsLoaded Then DoCmd.Close acForm, ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_PAUSED), acSaveNo
    If CurrentProject.AllForms(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_RUNNING)).IsLoaded = False Then DoCmd.OpenForm ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_RUNNING), acNormal
    ShowSetupControlFormWindow ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_RUNNING)
End Sub

Private Sub OpenSetupControlPausedForm()
    mSetupControlPaused = True
    If CurrentProject.AllForms(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_RUNNING)).IsLoaded Then DoCmd.Close acForm, ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_RUNNING), acSaveNo
    If CurrentProject.AllForms(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_PAUSED)).IsLoaded = False Then DoCmd.OpenForm ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_PAUSED), acNormal
    ShowSetupControlFormWindow ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_PAUSED)
    UpdateSetupControlFormStatus ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_PAUSED), SetupControlVisibleStatus(True)
End Sub

Private Sub CloseSetupControlForm(ByVal formName As String)
    If CurrentProject.AllForms(formName).IsLoaded Then DoCmd.Close acForm, formName, acSaveNo
End Sub

Private Sub UpdateSetupControlFormStatus(ByVal formName As String, ByVal statusText As String)
    Dim statusControl As control
    Dim elapsedControl As control
    Dim elapsedText As String
    Dim statusChanged As Boolean
    Dim elapsedChanged As Boolean
    Dim repaintTimer As Single
    If SetupHostWindowIsMinimized Then Exit Sub
    If CurrentProject.AllForms(formName).IsLoaded = False Then Exit Sub
    Set statusControl = forms(formName).Controls(SETUP_CONTROL_STATUS_CONTROL)
    Set elapsedControl = forms(formName).Controls(SETUP_CONTROL_ELAPSED_CONTROL)
    elapsedText = "Elapsed: " & SetupControlElapsedText(ElapsedSecondsFromTimer(mSetupControlStartedTimer))
    statusChanged = (StrComp(CStr(Nz(statusControl.Caption, vbNullString)), statusText, vbBinaryCompare) <> 0)
    elapsedChanged = (StrComp(CStr(Nz(elapsedControl.Caption, vbNullString)), elapsedText, vbBinaryCompare) <> 0)
    If statusChanged = False And elapsedChanged = False Then Exit Sub
    If statusChanged Then statusControl.Caption = statusText
    If elapsedChanged Then elapsedControl.Caption = elapsedText
    If mSetupControlCheckpointDisplayActive Then
        repaintTimer = Timer
        LogSetupControlCheckpointStage "Setup Control checkpoint form repaint starting.", mSetupControlStatusText, "Form: " & formName
        forms(formName).Repaint
        LogSetupControlCheckpointStage "Setup Control checkpoint form repaint completed.", mSetupControlStatusText, "Form: " & formName & " | Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(repaintTimer))
    Else
        forms(formName).Repaint
    End If
End Sub

' =============================================================================================================================================================================
' UI RECOVERY PROCEDURES
' =============================================================================================================================================================================

' Purpose: Checks whether current state and policy permit the caller to cel setup prepared work.
' Public role: Public module entry point. This procedure is listed in the module header as an intentional externally callable ADC entry point and may be invoked directly by Access/VBA callers.
' Parameters:
'   resultText (Output; String): Receives the result text produced by the operation so the caller can inspect the outcome without re-running the work.
' Returns: Boolean: True when the supplied/current context can cel setup prepared work; False otherwise.
Public Function CancelSetupPreparedWork(ByRef resultText As String) As Boolean
    On Error GoTo ErrHandler
    Dim ribbonRollbackDetail As String
    Dim recoveryClearDetail As String
    Dim recoveryClearFailed As Boolean
    resultText = vbNullString
    If mRibbonPhaseTransactionStarted Then
        On Error Resume Next
        mRibbonPhaseWorkspace.Rollback
        If Err.Number <> 0 Then
            ribbonRollbackDetail = "H2 Ribbon rollback failed. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
            Err.Clear
        Else
            mRibbonPhaseTransactionStarted = False
            ribbonRollbackDetail = "H2 staged Ribbon transaction was rolled back."
            recoveryClearFailed = (ClearUIRecoveryFailure("Ribbon transaction rollback", recoveryClearDetail) = False)
            ResetRibbonPhaseState
        End If
        On Error GoTo ErrHandler
        If Len(recoveryClearDetail) > 0 Then ribbonRollbackDetail = ribbonRollbackDetail & vbCrLf & recoveryClearDetail
    ElseIf mRibbonPhasePrepared Then
        ResetRibbonPhaseState
        ribbonRollbackDetail = "Uncommitted H1/H2 Ribbon plan state was discarded."
    End If
    EndSetupControl
    resultText = "Setup Control was closed without changing the Access application window state."
    If Len(ribbonRollbackDetail) > 0 Then resultText = resultText & vbCrLf & ribbonRollbackDetail
    If recoveryClearFailed Then Exit Function
    If InStr(1, ribbonRollbackDetail, "failed", vbTextCompare) > 0 Then Exit Function
    CancelSetupPreparedWork = True
    Exit Function
ErrHandler:
    resultText = "Setup prepared-work cleanup failed. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description & IIf(Len(ribbonRollbackDetail) = 0, vbNullString, vbCrLf & ribbonRollbackDetail)
End Function

' Purpose: Performs UI recovery blocked for setup preflight by coordinating `End` and `isMissing` while preserving this module's existing validation and error semantics.
' Public role: Public module entry point. This procedure is listed in the module header as an intentional externally callable ADC entry point and may be invoked directly by Access/VBA callers.
' Parameters:
'   detailText (Output Optional; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller. Default: vbNullString.
'   durableRecoverySnapshot (Input Optional; Variant): Snapshot of durable recovery state captured before or during the operation and used for validation, rollback, or recovery decisions.
' Returns: Boolean: True when UI recovery is blocked by the prerequisite/recovery condition checked here; False when processing may proceed.
Public Function UIRecoveryBlockedForSetupPreflight(Optional ByRef detailText As String = vbNullString, Optional ByVal durableRecoverySnapshot As Variant) As Boolean
    'Read-only Setup Preflight gate. It detects unfinished UI/ribbon recovery through the setup recovery store but does not resolve, commit, or roll back it.
    On Error GoTo FailClosed
    Dim durableOperation As String
    Dim durableDetail As String
    Dim durablePayload As String
    Dim recoverySnapshot As Object
    Dim ribbonMarkerExists As Boolean
    Dim uiMarkerExists As Boolean
    detailText = vbNullString
    If mUIRecoveryBlocked Then
        detailText = "In-memory UI recovery gate is blocked." & IIf(Len(Trim$(mUIRecoveryOperationName)) = 0, vbNullString, vbCrLf & "Operation: " & mUIRecoveryOperationName)
        UIRecoveryBlockedForSetupPreflight = True
        Exit Function
    End If
    If isMissing(durableRecoverySnapshot) Then
        ribbonMarkerExists = CoreDurableRecoveryMarkerExists(UI_RIBBON_COMMIT_MARKER_KEY, durableOperation, durableDetail, durablePayload)
    Else
        If IsObject(durableRecoverySnapshot) = False Then Err.Raise vbObjectError + 7214, "UIRecoveryBlockedForSetupPreflight", "The supplied durable recovery snapshot is not an object."
        Set recoverySnapshot = durableRecoverySnapshot
        ribbonMarkerExists = CoreDurableRecoveryMarkerExistsInSnapshot(recoverySnapshot, UI_RIBBON_COMMIT_MARKER_KEY, durableOperation, durableDetail, durablePayload)
    End If
    If ribbonMarkerExists Then
        detailText = "A pending ribbon commit marker exists. Resolve it with the dedicated recovery path before running Setup." & IIf(Len(durableOperation) = 0, vbNullString, vbCrLf & "Operation: " & durableOperation) & IIf(Len(durableDetail) = 0, vbNullString, vbCrLf & durableDetail)
        UIRecoveryBlockedForSetupPreflight = True
        Exit Function
    End If
    durableOperation = vbNullString
    durableDetail = vbNullString
    durablePayload = vbNullString
    If isMissing(durableRecoverySnapshot) Then
        uiMarkerExists = CoreDurableRecoveryMarkerExists(UI_RECOVERY_KEY, durableOperation, durableDetail, durablePayload)
    Else
        uiMarkerExists = CoreDurableRecoveryMarkerExistsInSnapshot(recoverySnapshot, UI_RECOVERY_KEY, durableOperation, durableDetail, durablePayload)
    End If
    If uiMarkerExists Then
        detailText = "Durable UI recovery state exists." & IIf(Len(durableOperation) = 0, vbNullString, vbCrLf & "Operation: " & durableOperation) & IIf(Len(durableDetail) = 0, vbNullString, vbCrLf & durableDetail)
        UIRecoveryBlockedForSetupPreflight = True
    End If
    Exit Function
FailClosed:
    detailText = "UI recovery state could not be inspected safely. Error " & CStr(Err.Number) & ": " & Err.Description
    UIRecoveryBlockedForSetupPreflight = True
End Function

' Purpose: Validates whether UI recovery backup can be read and used by the recovery/preflight operation.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   backupPath (Input; String): Filesystem or database path identifying the backup path location used by the operation.
'   failureText (Output; String): Receives the resolved failure text produced from this procedure's validated inputs and current ADC state.
' Returns: Boolean: True when the recovery backup required by UI preflight exists and can be read/validated; False when it is absent, unreadable, or unusable.
Public Function UIRecoveryBackupReadable(ByVal backupPath As String, ByRef failureText As String) As Boolean
    On Error GoTo ErrHandler
    backupPath = Trim$(backupPath)
    If Len(backupPath) = 0 Then Err.Raise vbObjectError + 7213, "UIRecoveryBackupReadable", "The backup path is blank."
    If Len(Dir$(backupPath)) = 0 Then Err.Raise vbObjectError + 7214, "UIRecoveryBackupReadable", "The backup file does not exist."
    If fileLen(backupPath) <= 0 Then Err.Raise vbObjectError + 7215, "UIRecoveryBackupReadable", "The backup file is empty."
    UIRecoveryBackupReadable = True
    Exit Function
ErrHandler:
    failureText = "Backup validation failed. Path: " & backupPath & "; Error " & CStr(Err.Number) & ": " & Err.Description
End Function

' Purpose: Determines whether UI recovery is blocked by an unresolved prerequisite, recovery state, or policy condition.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: Boolean: True when UI recovery is blocked by the condition tracked by this routine; False when work may proceed.
Public Function UIRecoveryBlocked() As Boolean
    'Intended use: provides a non-raising cross-module recovery gate that includes durable CodeDb recovery state.
    On Error GoTo FailClosed
    Dim durableOperation As String
    Dim durableDetail As String
    Dim durableType As String
    Dim durablePayload As String
    Dim durableStatus As String
    Dim preparedOwnedByCurrentContext As Boolean
    Dim ignoredRecoveryOperationName As String
    Dim ignoredRecoveryDetailText As String
    Dim ignoredRecoveryPayloadText As String
    If CoreDurableRecoveryMarkerExists(UI_RIBBON_COMMIT_MARKER_KEY, ignoredRecoveryOperationName, ignoredRecoveryDetailText, ignoredRecoveryPayloadText) Then ResolvePendingRibbonCommitMarker
    If mUIRecoveryBlocked Then UIRecoveryBlocked = True: Exit Function
    If UIDurableRecoveryRecordExists(durableOperation, durableDetail, durableType, durablePayload, durableStatus) Then
        preparedOwnedByCurrentContext = (StrComp(durableStatus, UI_RECOVERY_STATUS_PREPARED, vbTextCompare) = 0)
        preparedOwnedByCurrentContext = preparedOwnedByCurrentContext And mUIDurableRecoveryOwned
        preparedOwnedByCurrentContext = preparedOwnedByCurrentContext And (mUIRecoveryBlocked = False)
        preparedOwnedByCurrentContext = preparedOwnedByCurrentContext And (StrComp(Trim$(durableOperation), Trim$(mUIRecoveryOperationName), vbTextCompare) = 0)
        preparedOwnedByCurrentContext = preparedOwnedByCurrentContext And (StrComp(Trim$(durableType), Trim$(mUIRecoveryType), vbTextCompare) = 0)
        preparedOwnedByCurrentContext = preparedOwnedByCurrentContext And (StrComp(durableDetail, mUIRecoveryDetail, vbBinaryCompare) = 0)
        preparedOwnedByCurrentContext = preparedOwnedByCurrentContext And (StrComp(durablePayload, mUIRecoveryPayload, vbBinaryCompare) = 0)
        If preparedOwnedByCurrentContext Then Exit Function
        mUIRecoveryBlocked = True
        mUIRecoveryOperationName = durableOperation
        mUIRecoveryDetail = durableDetail
        mUIRecoveryType = durableType
        mUIRecoveryPayload = durablePayload
        mUIDurableRecoveryOwned = False
        UIRecoveryBlocked = True
    End If
    Exit Function
FailClosed:
    mUIRecoveryBlocked = True
    mUIRecoveryOperationName = "Durable recovery-state inspection"
    mUIRecoveryDetail = "The durable UI recovery state could not be inspected. Error " & CStr(Err.Number) & ": " & Err.Description
    mUIRecoveryType = vbNullString
    mUIRecoveryPayload = vbNullString
    mUIDurableRecoveryOwned = False
    UIRecoveryBlocked = True
End Function

Private Function UIRecoveryTypeForOperation(ByVal operationName As String) As String
    'Returns the authoritative recovery type for a known UI recovery operation.
    operationName = Trim$(operationName)
    If StrComp(operationName, "Form-reference restoration", vbTextCompare) = 0 Then
    ElseIf StrComp(operationName, "Generated-form restoration", vbTextCompare) = 0 Then
        UIRecoveryTypeForOperation = UI_RECOVERY_TYPE_GENERATED_FORMS
    ElseIf InStr(1, operationName, "Ribbon", vbTextCompare) > 0 Then
        UIRecoveryTypeForOperation = UI_RECOVERY_TYPE_RIBBON_TRANSACTION
    End If
End Function

Private Function UISerializeCoreRecoveryPayload(ByVal recoveryType As String, ByVal recoveryPayload As String, ByVal recoveryStatus As String) As String
    UISerializeCoreRecoveryPayload = "Version=1" & vbCrLf & _
        "RecoveryStatus=" & UIDurableEncode(recoveryStatus) & vbCrLf & _
        "RecoveryType=" & UIDurableEncode(recoveryType) & vbCrLf & _
        "Payload=" & UIDurableEncode(recoveryPayload)
End Function

' Purpose: Resolves unique recovery form name from the supplied/current ADC state, using `SharedDao_AccessObjectExists` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   prefixText (Input; String): Input prefix text for unique recovery form name; is passed to `candidateName`.
' Returns: String: Returns the resolved name produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function UniqueRecoveryFormName(ByVal prefixText As String) As String
    Dim candidateName As String
    Dim attempt As Long
    Do
        SetupControlPump
        attempt = attempt + 1
        candidateName = Left$(prefixText & Format$(Now, "hhnnss") & Format$(CLng(Timer * 1000), "000000000") & CStr(attempt), 60)
    Loop While SharedDao_AccessObjectExists(GetHostDb(), acForm, candidateName)
    UniqueRecoveryFormName = candidateName
End Function


' Purpose: Clears a matching recovery blocker, verifies absence, and restores equivalent durable protection after post-delete verification failure.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   operationName (Input; String): Logical operation name written to progress, trace, warning, or error context for this call.
'   failureDetail (Output; String): Receives failure detail used for caller diagnostics, verification output, or user-facing reporting.
' Returns: Boolean: True on the function's affirmative/success branch for clear UI recovery failure; False on its explicit negative, unavailable, rejected, or handled-failure branch.
Public Function ClearUIRecoveryFailure(ByVal operationName As String, ByRef failureDetail As String) As Boolean
    'Clears a matching recovery blocker, verifies absence, and restores equivalent durable protection after post-delete verification failure.
    On Error GoTo ErrHandler
    Dim durableOperation As String
    Dim durableDetail As String
    Dim durableType As String
    Dim durablePayload As String
    Dim durableStatus As String
    Dim originalOperation As String
    Dim originalDetail As String
    Dim originalType As String
    Dim originalPayload As String
    Dim checkOperation As String
    Dim checkDetail As String
    Dim checkType As String
    Dim checkPayload As String
    Dim checkStatus As String
    Dim durableExists As Boolean
    Dim deletionCompleted As Boolean
    Dim restoreErrNumber As Long
    Dim restoreErrSource As String
    Dim restoreErrDescription As String
    durableExists = UIDurableRecoveryRecordExists(durableOperation, durableDetail, durableType, durablePayload, durableStatus)
    If durableExists Then
        originalOperation = durableOperation
        originalDetail = durableDetail
        originalType = durableType
        originalPayload = durablePayload
        If Len(Trim$(operationName)) > 0 Then
            If StrComp(Trim$(operationName), durableOperation, vbTextCompare) <> 0 Then
                failureDetail = "The durable recovery blocker belongs to a different operation and was not cleared." & vbCrLf & "Requested operation: " & operationName & vbCrLf & "Durable operation: " & durableOperation
                Exit Function
            End If
        End If
    ElseIf mUIRecoveryBlocked = False Then
        ClearUIRecoveryFailure = True
        Exit Function
    End If
    If Len(Trim$(operationName)) > 0 And Len(mUIRecoveryOperationName) > 0 Then
        If StrComp(Trim$(operationName), mUIRecoveryOperationName, vbTextCompare) <> 0 Then
            failureDetail = "The in-memory recovery blocker belongs to a different operation and was not cleared." & vbCrLf & "Requested operation: " & operationName & vbCrLf & "In-memory operation: " & mUIRecoveryOperationName
            Exit Function
        End If
    End If
    If durableExists Then
        DeleteUIDurableRecoveryRecord
        deletionCompleted = True
        If UIDurableRecoveryRecordExists(checkOperation, checkDetail, checkType, checkPayload, checkStatus) Then
            mUIRecoveryBlocked = True
            mUIRecoveryOperationName = checkOperation
            mUIRecoveryDetail = checkDetail
            mUIRecoveryType = checkType
            mUIRecoveryPayload = checkPayload
            mUIDurableRecoveryOwned = False
            failureDetail = "The durable recovery blocker still exists after verified deletion." & vbCrLf & "Operation: " & checkOperation & vbCrLf & "Recovery type: " & checkType & vbCrLf & "Status: " & checkStatus & vbCrLf & "Detail: " & checkDetail
            Exit Function
        End If
    End If
    mUIRecoveryBlocked = False
    mUIRecoveryOperationName = vbNullString
    mUIRecoveryDetail = vbNullString
    mUIRecoveryType = vbNullString
    mUIRecoveryPayload = vbNullString
    mUIDurableRecoveryOwned = False
    Set mUIRecoveryWorkspace = Nothing
    ClearUIRecoveryFailure = True
    Exit Function
RestoreBlockedRecord:
    On Error Resume Next
    If Len(Trim$(originalOperation)) = 0 Then originalOperation = IIf(Len(Trim$(operationName)) = 0, "UI recovery verification", operationName)
    If Len(Trim$(originalDetail)) = 0 Then originalDetail = failureDetail
    If Len(Trim$(originalType)) = 0 Then originalType = UIRecoveryTypeForOperation(originalOperation)
    Err.Clear
    PersistUIDurableRecoveryRecord originalOperation, originalDetail & vbCrLf & "Partial cleanup detail: " & failureDetail, originalType, originalPayload, RECOVERY_STATUS_BLOCKED
    restoreErrNumber = Err.Number
    restoreErrSource = Err.Source
    restoreErrDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    mUIRecoveryBlocked = True
    mUIRecoveryOperationName = originalOperation
    mUIRecoveryDetail = originalDetail & vbCrLf & "Partial cleanup detail: " & failureDetail
    mUIRecoveryType = originalType
    mUIRecoveryPayload = originalPayload
    mUIDurableRecoveryOwned = False
    If restoreErrNumber <> 0 Then
        failureDetail = failureDetail & vbCrLf & "Equivalent durable blocked-recovery restoration also failed. Error " & CStr(restoreErrNumber) & " from " & restoreErrSource & ": " & restoreErrDescription
    End If
    Exit Function
ErrHandler:
    failureDetail = "Could not clear and verify the UI recovery blocker." & vbCrLf & "Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    If deletionCompleted Then GoTo RestoreBlockedRecord
End Function

' Purpose: Registers UI recovery failure for operation name, recovery workspace.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   operationName (Input; String): Logical operation name written to progress, trace, warning, or error context for this call.
'   detailText (Input; String): Receives or supplies diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
'   recoveryWorkspace (Input Optional; DAO.workspace): DAO.Workspace used for the recovery transaction so transaction ownership/commit/rollback occurs in the caller-selected workspace; Nothing uses the normal
'   workspace path. Default: Nothing.
'   recoveryState (Input Optional; Collection): State data describing the current or prior recovery condition used for recovery, rollback, UI, or lifecycle decisions. Default: Nothing.
'   recoveryType (Input Optional; String): Recovery-marker/type token identifying which durable recovery payload/state is being written, read, or cleared. Default: vbNullString.
'   recoveryPayload (Input Optional; String): Serialized or structured payload carrying the recovery recovery/operation data that must be persisted or reapplied. Default: vbNullString.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub RegisterUIRecoveryFailure(ByVal operationName As String, ByVal detailText As String, Optional ByVal recoveryWorkspace As DAO.workspace = Nothing, Optional ByVal recoveryState As Collection = Nothing, Optional ByVal recoveryType As String = vbNullString, Optional ByVal recoveryPayload As String = vbNullString)
    On Error GoTo PersistErr
    Dim requestedOperation As String
    Dim requestedDetail As String
    Dim requestedType As String
    Dim existingOperation As String
    Dim existingDetail As String
    Dim existingType As String
    Dim existingPayload As String
    Dim existingStatus As String
    Dim existingRecord As Boolean
    Dim allowedReplacement As Boolean
    Dim exactOwnedPreparedContext As Boolean
    Dim sameBlockedRecovery As Boolean
    requestedOperation = operationName
    requestedDetail = Trim$(detailText)
    requestedType = recoveryType
    If Len(requestedType) = 0 Then requestedType = UIRecoveryTypeForOperation(requestedOperation)
    If Len(requestedOperation) = 0 Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "RegisterUIRecoveryFailure", "A UI recovery failure cannot be registered with a blank operation name."
    If Len(requestedType) = 0 Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "RegisterUIRecoveryFailure", "A UI recovery failure could not resolve a supported recovery type." & vbCrLf & "Operation: " & requestedOperation
    If Len(recoveryPayload) = 0 Then
        If StrComp(requestedType, UI_RECOVERY_TYPE_GENERATED_FORMS, vbTextCompare) = 0 And Not recoveryState Is Nothing Then recoveryPayload = SerializeGeneratedFormRecoveryState(recoveryState)
    End If
    existingRecord = UIDurableRecoveryRecordExists(existingOperation, existingDetail, existingType, existingPayload, existingStatus)
    If existingRecord Then
        exactOwnedPreparedContext = mUIDurableRecoveryOwned
        exactOwnedPreparedContext = exactOwnedPreparedContext And (mUIRecoveryBlocked = False)
        exactOwnedPreparedContext = exactOwnedPreparedContext And (StrComp(existingStatus, UI_RECOVERY_STATUS_PREPARED, vbTextCompare) = 0)
        exactOwnedPreparedContext = exactOwnedPreparedContext And (StrComp(Trim$(existingOperation), Trim$(mUIRecoveryOperationName), vbTextCompare) = 0)
        exactOwnedPreparedContext = exactOwnedPreparedContext And (StrComp(Trim$(existingOperation), requestedOperation, vbTextCompare) = 0)
        exactOwnedPreparedContext = exactOwnedPreparedContext And (StrComp(Trim$(existingType), Trim$(mUIRecoveryType), vbTextCompare) = 0)
        exactOwnedPreparedContext = exactOwnedPreparedContext And (StrComp(Trim$(existingType), requestedType, vbTextCompare) = 0)
        exactOwnedPreparedContext = exactOwnedPreparedContext And (StrComp(existingDetail, mUIRecoveryDetail, vbBinaryCompare) = 0)
        exactOwnedPreparedContext = exactOwnedPreparedContext And (StrComp(existingPayload, mUIRecoveryPayload, vbBinaryCompare) = 0)
        exactOwnedPreparedContext = exactOwnedPreparedContext And (StrComp(existingPayload, recoveryPayload, vbBinaryCompare) = 0)
        sameBlockedRecovery = (StrComp(existingStatus, RECOVERY_STATUS_BLOCKED, vbTextCompare) = 0)
        sameBlockedRecovery = sameBlockedRecovery And (StrComp(Trim$(existingOperation), requestedOperation, vbTextCompare) = 0)
        sameBlockedRecovery = sameBlockedRecovery And (StrComp(Trim$(existingType), requestedType, vbTextCompare) = 0)
        If sameBlockedRecovery Then
            If Len(recoveryPayload) = 0 Then
                recoveryPayload = existingPayload
            ElseIf StrComp(existingPayload, recoveryPayload, vbBinaryCompare) <> 0 Then
                sameBlockedRecovery = False
            End If
        End If
        allowedReplacement = exactOwnedPreparedContext Or sameBlockedRecovery
        If allowedReplacement = False Then
            Err.Raise ERR_UI_RECOVERY_BLOCKED, "RegisterUIRecoveryFailure", "A different or unowned durable UI recovery record already exists. It was not overwritten by the new recovery failure." & vbCrLf & "Requested operation: " & requestedOperation & vbCrLf & "Requested recovery type: " & requestedType & vbCrLf & "Existing operation: " & IIf(Len(Trim$(existingOperation)) = 0, "(blank)", existingOperation) & vbCrLf & "Existing recovery type: " & IIf(Len(Trim$(existingType)) = 0, "(blank)", existingType) & vbCrLf & "Existing status: " & IIf(Len(Trim$(existingStatus)) = 0, "(blank)", existingStatus) & vbCrLf & "Existing detail: " & existingDetail & vbCrLf & "Requested payload matches existing: " & CStr(StrComp(existingPayload, recoveryPayload, vbBinaryCompare) = 0)
        End If
    End If
    PersistUIDurableRecoveryRecord requestedOperation, requestedDetail, requestedType, recoveryPayload, RECOVERY_STATUS_BLOCKED
    mUIRecoveryBlocked = True
    mUIRecoveryOperationName = requestedOperation
    mUIRecoveryDetail = requestedDetail
    mUIRecoveryType = requestedType
    mUIRecoveryPayload = recoveryPayload
    mUIDurableRecoveryOwned = False
    If Not recoveryWorkspace Is Nothing Then Set mUIRecoveryWorkspace = recoveryWorkspace
    Exit Sub
PersistErr:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If existingRecord Then
        mUIRecoveryBlocked = True
        mUIRecoveryOperationName = existingOperation
        mUIRecoveryDetail = existingDetail
        mUIRecoveryType = existingType
        mUIRecoveryPayload = existingPayload
        mUIDurableRecoveryOwned = False
    Else
        mUIRecoveryBlocked = True
        mUIRecoveryOperationName = requestedOperation
        mUIRecoveryDetail = requestedDetail & vbCrLf & "Durable recovery-state persistence failed. Error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription
        mUIRecoveryType = requestedType
        mUIRecoveryPayload = recoveryPayload
        mUIDurableRecoveryOwned = False
    End If
    Err.Raise ERR_UI_DURABLE_RECOVERY_FAILED, "RegisterUIRecoveryFailure", IIf(existingRecord, "The existing durable recovery record was preserved. The new recovery failure could not be registered. ", vbNullString) & "Error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription
End Sub

' Purpose: Requires no unfinished UI recovery for requested operation name.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   requestedOperationName (Input; String): Operation-name token identifying the destructive/maintenance action for confirmation, diagnostics, and dispatch.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub RequireNoUnfinishedUIRecovery(ByVal requestedOperationName As String)
    'Intended use: blocks UI, setup, reconciliation, generated-object, Ribbon, and managed-object naming work while any Core, Catalog, or UI recovery remains unresolved.
    requestedOperationName = Trim$(requestedOperationName)
    If Len(requestedOperationName) = 0 Then requestedOperationName = DEFAULT_UTILITY_LONG_NAME & " UI operation"
    If ConfigCandidateMarkingRecoveryBlocked() Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "RequireNoUnfinishedUIRecovery", "A " & DEFAULT_UTILITY_LONG_NAME & " UI operation cannot start because Config/Candidate marking retains unresolved transaction recovery." & vbCrLf & "Requested operation: " & requestedOperationName
    If CatalogRecoveryBlocked() Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "RequireNoUnfinishedUIRecovery", "A " & DEFAULT_UTILITY_LONG_NAME & " UI operation cannot start because the Catalog retains unresolved transaction or snapshot recovery state." & vbCrLf & "Requested operation: " & requestedOperationName
    If UIRecoveryBlocked() = False Then Exit Sub
    Err.Raise ERR_UI_RECOVERY_BLOCKED, "RequireNoUnfinishedUIRecovery", "A " & DEFAULT_UTILITY_LONG_NAME & " UI operation cannot start because an earlier recovery boundary is unresolved." & vbCrLf & "Requested operation: " & requestedOperationName & vbCrLf & "Recovery operation: " & IIf(Len(mUIRecoveryOperationName) = 0, "(not recorded)", mUIRecoveryOperationName) & vbCrLf & "Recovery detail:" & vbCrLf & IIf(Len(mUIRecoveryDetail) = 0, "(not recorded)", mUIRecoveryDetail) & vbCrLf & "Do not continue setup, reconciliation, form generation, Ribbon, or managed-object naming work until recovery is retried and verified."
End Sub

Private Function ExecutePendingUIRecovery() As String
    'Executes and verifies the authoritative UI durable-recovery operation selected from the UI-owned recovery protocol.
    On Error GoTo ErrHandler
    Dim operationName As String
    Dim detailText As String
    Dim recoveryType As String
    Dim recoveryPayload As String
    Dim recoveryStatus As String
    Dim checkOperation As String
    Dim checkDetail As String
    Dim checkType As String
    Dim checkPayload As String
    Dim checkStatus As String
    Dim recoveryState As Collection
    Dim resultText As String
    Dim clearFailureDetail As String
    Dim ignoredInitialRecoveryOperationName As String
    Dim ignoredInitialRecoveryDetailText As String
    Dim ignoredInitialRecoveryPayloadText As String
    Dim ignoredRemainingRecoveryOperationName As String
    Dim ignoredRemainingRecoveryDetailText As String
    Dim ignoredRemainingRecoveryPayloadText As String
    If CoreDurableRecoveryMarkerExists(UI_RIBBON_COMMIT_MARKER_KEY, ignoredInitialRecoveryOperationName, ignoredInitialRecoveryDetailText, ignoredInitialRecoveryPayloadText) Then ResolvePendingRibbonCommitMarker
    If UIDurableRecoveryRecordExists(operationName, detailText, recoveryType, recoveryPayload, recoveryStatus) = False Then
        If mUIRecoveryBlocked = False Then
            ExecutePendingUIRecovery = "No unfinished Access Development Console UI recovery record exists."
            Exit Function
        End If
        operationName = mUIRecoveryOperationName
        recoveryType = mUIRecoveryType
        recoveryPayload = mUIRecoveryPayload
        If Len(Trim$(recoveryType)) = 0 Then recoveryType = UIRecoveryTypeForOperation(operationName)
        If Len(Trim$(recoveryType)) = 0 Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "ExecutePendingUIRecovery", "The in-memory UI recovery blocker has no recognized recovery type." & vbCrLf & "Operation: " & IIf(Len(Trim$(operationName)) = 0, "(blank)", operationName) & vbCrLf & "Detail: " & mUIRecoveryDetail
    Else
        mUIRecoveryType = recoveryType
        mUIRecoveryPayload = recoveryPayload
    End If
    Select Case UCase$(Trim$(recoveryType))
        Case UCase$(UI_RECOVERY_TYPE_GENERATED_FORMS)
            resultText = ResolveUIRecoveryWorkspaceTransaction("generated-form catalog recovery")
            Set recoveryState = ParseGeneratedFormRecoveryState(recoveryPayload)
            resultText = resultText & vbCrLf & RestoreGeneratedFormRecoveryState(recoveryState)
        Case UCase$(UI_RECOVERY_TYPE_RIBBON_TRANSACTION)
            If CoreDurableRecoveryMarkerExists(UI_RIBBON_COMMIT_MARKER_KEY, ignoredRemainingRecoveryOperationName, ignoredRemainingRecoveryDetailText, ignoredRemainingRecoveryPayloadText) Then
                ResolvePendingRibbonCommitMarker
                If UIDurableRecoveryRecordExists(checkOperation, checkDetail, checkType, checkPayload, checkStatus) = False Then
                    resultText = "Committed ribbon state was verified and pending post-commit recovery was completed."
                Else
                    operationName = checkOperation
                    resultText = RetryRibbonTransactionRecovery()
                    If ClearUIRecoveryFailure(operationName, clearFailureDetail) = False Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "ExecutePendingUIRecovery", clearFailureDetail
                End If
            Else
                resultText = RetryRibbonTransactionRecovery()
                If ClearUIRecoveryFailure(operationName, clearFailureDetail) = False Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "ExecutePendingUIRecovery", clearFailureDetail
            End If
        Case Else
            Err.Raise ERR_UI_RECOVERY_BLOCKED, "ExecutePendingUIRecovery", "The UI recovery record has an unsupported recovery type: " & recoveryType & vbCrLf & "Operation: " & operationName
    End Select
    If UIDurableRecoveryRecordExists(checkOperation, checkDetail, checkType, checkPayload, checkStatus) Then
        mUIRecoveryBlocked = True
        mUIRecoveryOperationName = checkOperation
        mUIRecoveryDetail = checkDetail
        mUIRecoveryType = checkType
        mUIRecoveryPayload = checkPayload
        mUIDurableRecoveryOwned = False
        Err.Raise ERR_UI_RECOVERY_BLOCKED, "ExecutePendingUIRecovery", "Recovery work completed, but a durable recovery blocker still exists." & vbCrLf & "Operation: " & checkOperation & vbCrLf & "Recovery type: " & checkType & vbCrLf & "Status: " & checkStatus & vbCrLf & "Detail: " & checkDetail
    End If
    If mUIRecoveryBlocked Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "ExecutePendingUIRecovery", "Recovery work completed, but the in-memory UI recovery gate remains blocked." & vbCrLf & "Operation: " & mUIRecoveryOperationName & vbCrLf & "Recovery type: " & mUIRecoveryType & vbCrLf & "Detail: " & mUIRecoveryDetail
    ExecutePendingUIRecovery = resultText
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, IIf(Len(Trim$(savedErrSource)) = 0, "ExecutePendingUIRecovery", savedErrSource), savedErrDescription
End Function

' Purpose: Resolves UI recovery workspace transaction for recovery context.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   recoveryContext (Input; String): Recovery context text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
' Returns: String: Returns the resolved resolve UI recovery workspace transaction text in the canonical representation consumed by this module's caller; vbNullString represents no applicable value where the
'   existing contract permits it.
Public Function ResolveUIRecoveryWorkspaceTransaction(ByVal recoveryContext As String) As String
    On Error GoTo ErrHandler
    Dim rollbackErrNumber As Long
    Dim rollbackErrDescription As String
    Dim verifyWs As DAO.workspace
    Dim verifyDb As DAO.Database
    If Not mUIRecoveryWorkspace Is Nothing Then
        On Error Resume Next
        Err.Clear
        mUIRecoveryWorkspace.Rollback
        rollbackErrNumber = Err.Number
        rollbackErrDescription = Err.Description
        On Error GoTo ErrHandler
        If rollbackErrNumber <> 0 And rollbackErrNumber <> 3034 Then Err.Raise ERR_RIBBON_ROLLBACK_FAILED, "ResolveUIRecoveryWorkspaceTransaction", "The retained originating " & recoveryContext & " workspace still could not be rolled back. Error " & CStr(rollbackErrNumber) & ": " & rollbackErrDescription
        Set mUIRecoveryWorkspace = Nothing
        If rollbackErrNumber = 3034 Then ResolveUIRecoveryWorkspaceTransaction = "The retained originating " & recoveryContext & " workspace had no live transaction remaining." Else ResolveUIRecoveryWorkspaceTransaction = "The retained originating " & recoveryContext & " transaction was rolled back successfully."
    Else
        ResolveUIRecoveryWorkspaceTransaction = "The originating " & recoveryContext & " workspace is no longer retained. No default-workspace rollback was attempted."
    End If
    OpenCoreRecoveryDatabase verifyWs, verifyDb
    verifyDb.TableDefs.Refresh
    verifyDb.Close
    Set verifyDb = Nothing
    verifyWs.Close
    Set verifyWs = Nothing
    ResolveUIRecoveryWorkspaceTransaction = ResolveUIRecoveryWorkspaceTransaction & vbCrLf & "CodeDb was reopened successfully through a new independent workspace before recovery-specific verification continued."
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not verifyDb Is Nothing Then verifyDb.Close
    If Not verifyWs Is Nothing Then verifyWs.Close
    On Error GoTo 0
    Err.Raise savedErrNumber, "ResolveUIRecoveryWorkspaceTransaction", savedErrDescription
End Function

' =============================================================================================================================================================================
' SHARED PROGRESS / STATUS RENDERING PROCEDURES
' =============================================================================================================================================================================

' Purpose: Resolves status bar text with current setup elapsed from the supplied/current ADC state, using `End` and `SetupControlElapsedText` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   statusText (Input; String): Status/progress message supplied for the current operation; the routine forwards it to the applicable UI/progress and trace/reporting path.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function StatusBarTextWithCurrentSetupElapsed(ByVal statusText As String) As String
    'Returns status-bar text using the same setup-start timer as the running and paused progress forms.
    statusText = Trim$(statusText)
    If Len(statusText) = 0 Then Exit Function
    If mSetupControlActive = False Or mSetupControlStartedTimer = 0 Then
        StatusBarTextWithCurrentSetupElapsed = statusText
        Exit Function
    End If
    If InStr(1, statusText, "Elapsed:", vbTextCompare) > 0 Or InStr(1, statusText, " elapsed", vbTextCompare) > 0 Then
        StatusBarTextWithCurrentSetupElapsed = statusText
    Else
        StatusBarTextWithCurrentSetupElapsed = statusText & " | Elapsed: " & SetupControlElapsedText(ElapsedSecondsFromTimer(mSetupControlStartedTimer))
    End If
End Function

' Purpose: Sets UI status bar text for status text.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   statusText (Input; String): Status/progress message supplied for the current operation; the routine forwards it to the applicable UI/progress and trace/reporting path.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub SetUIStatusBarText(ByVal statusText As String)
    On Error Resume Next
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    statusText = Trim$(statusText)
    If Len(statusText) = 0 Then Exit Sub
    statusText = StatusBarTextWithCurrentSetupElapsed(statusText)
    Err.Clear
    Application.SysCmd acSysCmdSetStatus, ApplyDisplayBranding(DEFAULT_OPERATION_PROGRESS_STATUS_PREFIX & StatusBarProgressText(statusText))
    If Err.Number <> 0 Then
        savedErrNumber = Err.Number
        savedErrSource = Err.Source
        savedErrDescription = Err.Description
        mUIStatusBarFailureCount = mUIStatusBarFailureCount + 1
        If mUIStatusBarFailureCount = 1 Or (mUIStatusBarFailureCount Mod DEFAULT_UI_STATUS_BAR_FAILURE_TRACE_INTERVAL) = 0 Then
            LogTraceMessage TRACE_LEVEL_WARNING, "StatusBar", "SetUIStatusBarText", "Status-bar update failed; UI work continues.", "Status text: " & statusText & vbCrLf & "Failure count this session: " & CStr(mUIStatusBarFailureCount) & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, "Set Access status bar", savedErrNumber, savedErrSource, savedErrDescription, vbNullString, "StatusBar", "CodeDb", "UI"
        End If
        Err.Clear
    End If
End Sub

Private Sub AdvanceChatGPTProgressStep()
    If mChatGPTProgressMax <= 0 Then Exit Sub
    mChatGPTProgressStep = mChatGPTProgressStep + 1
    If mChatGPTProgressStep < 1 Then mChatGPTProgressStep = 1
    If mChatGPTProgressStep > mChatGPTProgressMax Then mChatGPTProgressStep = mChatGPTProgressMax
End Sub

' Purpose: Initializes chat GPT progress, coordinating `ChatGPTProgressScopeFromCaption` and `CoreParentProgressActive` as the implementation requires.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   captionText (Input; String): Caption text text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   maxValue (Input; Long): Upper numeric bound enforced by the operation when validating or constraining the supplied value.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub BeginChatGPTProgress(ByVal captionText As String, ByVal maxValue As Long)
    BeginTrackedProcess captionText
    mChatGPTProgressStep = 0
    mChatGPTProgressMax = maxValue
    mChatGPTProgressScopeLabel = ChatGPTProgressScopeFromCaption(captionText)
    mChatGPTProgressNestedUnderParent = CoreParentProgressActive()
    If mChatGPTProgressNestedUnderParent Then
        mChatGPTProgressActive = False
        TrySharedProgressUpdate mChatGPTProgressActive, mChatGPTProgressStep, mChatGPTProgressMax, CoreParentNestedProgressText(captionText, mChatGPTProgressScopeLabel), DEFAULT_OPERATION_PROGRESS_STATUS_PREFIX, "BeginChatGPTProgress", "UI", "UI"
    Else
        TrySharedProgressBegin mChatGPTProgressActive, mChatGPTProgressStep, mChatGPTProgressMax, captionText, DEFAULT_OPERATION_PROGRESS_STATUS_PREFIX, "BeginChatGPTProgress", "UI", "UI"
    End If
    LogTraceMessage TRACE_LEVEL_INFO, "Progress", "BeginChatGPTProgress", captionText, "Scope: " & mChatGPTProgressScopeLabel & vbCrLf & "Max steps: " & CStr(maxValue) & vbCrLf & "Nested under parent progress: " & CStr(mChatGPTProgressNestedUnderParent), captionText, 0, vbNullString, vbNullString, vbNullString, vbNullString, "CodeDb", "UI"
End Sub

Private Function ChatGPTProgressPartPrefixText() As String
    Dim scopeText As String
    scopeText = mChatGPTProgressScopeLabel
    If Len(scopeText) > 0 Then ChatGPTProgressPartPrefixText = scopeText & " "
End Function

Private Function ChatGPTProgressPartText(ByVal statusText As String, Optional ByVal advanceStep As Boolean = False) As String
    Dim partNumber As Long
    statusText = Trim$(statusText)
    If Len(statusText) = 0 Then Exit Function
    If (mChatGPTProgressActive = False And mChatGPTProgressNestedUnderParent = False) Or mChatGPTProgressMax <= 0 Then
        ChatGPTProgressPartText = statusText
        Exit Function
    End If
    partNumber = mChatGPTProgressStep
    If advanceStep Then partNumber = partNumber + 1
    If partNumber < 1 Then partNumber = 1
    If partNumber > mChatGPTProgressMax Then partNumber = mChatGPTProgressMax
    ChatGPTProgressPartText = ChatGPTProgressPartPrefixText() & "Part " & CStr(partNumber) & " of " & CStr(mChatGPTProgressMax) & " > " & statusText
End Function

Private Function ChatGPTProgressScopeFromCaption(ByVal captionText As String) As String
    Dim captionKey As String
    captionKey = captionText
    If InStr(1, captionKey, "deleting", vbTextCompare) > 0 Then
        ChatGPTProgressScopeFromCaption = "Form Delete"
    ElseIf InStr(1, captionKey, "form", vbTextCompare) > 0 Then
        ChatGPTProgressScopeFromCaption = "Generated Forms"
    ElseIf Len(Trim$(captionText)) > 0 Then
        ChatGPTProgressScopeFromCaption = Trim$(captionText)
    Else
        ChatGPTProgressScopeFromCaption = "UI"
    End If
End Function

Private Function ChatGPTProgressVisibleText(ByVal statusText As String) As String
    If mChatGPTProgressNestedUnderParent And CoreParentProgressActive() Then
        ChatGPTProgressVisibleText = CoreParentNestedProgressText(statusText, mChatGPTProgressScopeLabel)
    Else
        ChatGPTProgressVisibleText = ChatGPTProgressPartText(statusText, False)
    End If
End Function

' Purpose: Closes chat GPT progress, coordinating `LogTraceMessage` and `End` as the implementation requires.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub EndChatGPTProgress()
    LogTraceMessage TRACE_LEVEL_INFO, "Progress", "EndChatGPTProgress", "ChatGPT " & DEFAULT_UTILITY_LONG_NAME & " UI progress ended.", vbNullString, vbNullString, 0, vbNullString, vbNullString, vbNullString, vbNullString, "CodeDb", "UI"
    If mChatGPTProgressNestedUnderParent = False Then TrySharedProgressEnd mChatGPTProgressActive, mChatGPTProgressStep, mChatGPTProgressMax, True, "EndChatGPTProgress", "UI", "UI"
    mChatGPTProgressNestedUnderParent = False
    mChatGPTProgressScopeLabel = vbNullString
    EndTrackedProcess
End Sub

' Purpose: Performs the public fail chat gpt progress operation using the routine's existing ADC workflow and side effects.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   statusText (Input; String): Status/progress message supplied for the current operation; the routine forwards it to the applicable UI/progress and trace/reporting path.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub FailChatGPTProgress(ByVal statusText As String)
    Dim partStatusText As String
    Dim visibleStatusText As String
    Dim traceStepText As String
    Dim buildForm As String
    Dim buildDetail As String
    If Len(Trim$(statusText)) = 0 Then statusText = "ChatGPT " & DEFAULT_UTILITY_SHORT_NAME & " UI process failed. Review the Immediate Window trace."
    partStatusText = ChatGPTProgressPartText(statusText, False)
    If mChatGPTProgressNestedUnderParent And CoreParentProgressActive() Then
        visibleStatusText = ChatGPTProgressVisibleText(statusText)
        traceStepText = visibleStatusText
        TrySharedProgressFail mChatGPTProgressActive, mChatGPTProgressStep, mChatGPTProgressMax, visibleStatusText, "FailChatGPTProgress", "UI", "UI"
    Else
        visibleStatusText = partStatusText
        traceStepText = partStatusText
        TrySharedProgressFail mChatGPTProgressActive, mChatGPTProgressStep, mChatGPTProgressMax, visibleStatusText, "FailChatGPTProgress", "UI", "UI"
    End If
    GetChatGPTBuildProgressContext buildForm, buildDetail
    LogTraceMessage TRACE_LEVEL_STATUS, "Progress", "FailChatGPTProgress", visibleStatusText, buildDetail, traceStepText, 0, vbNullString, vbNullString, buildForm, "Form", "CodeDb", "UI"
    mChatGPTProgressNestedUnderParent = False
    mChatGPTProgressScopeLabel = vbNullString
    CaptureTrackedProcessSummary
    EndTrackedProcess
End Sub

' Purpose: Updates chat gpt progress for status text.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   statusText (Input; String): Status/progress message supplied for the current operation; the routine forwards it to the applicable UI/progress and trace/reporting path.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub UpdateChatGPTProgress(ByVal statusText As String)
    Dim partStatusText As String
    Dim visibleStatusText As String
    Dim traceStepText As String
    Dim buildForm As String
    Dim buildDetail As String
    partStatusText = ChatGPTProgressPartText(statusText, True)
    If mChatGPTProgressNestedUnderParent And CoreParentProgressActive() Then
        AdvanceChatGPTProgressStep
        visibleStatusText = ChatGPTProgressVisibleText(statusText)
        traceStepText = visibleStatusText
        TrySharedProgressUpdate mChatGPTProgressActive, mChatGPTProgressStep, mChatGPTProgressMax, visibleStatusText, DEFAULT_OPERATION_PROGRESS_STATUS_PREFIX, "UpdateChatGPTProgress", "UI", "UI"
    Else
        visibleStatusText = partStatusText
        traceStepText = partStatusText
        TrySharedProgressUpdate mChatGPTProgressActive, mChatGPTProgressStep, mChatGPTProgressMax, visibleStatusText, DEFAULT_OPERATION_PROGRESS_STATUS_PREFIX, "UpdateChatGPTProgress", "UI", "UI"
    End If
    GetChatGPTBuildProgressContext buildForm, buildDetail
    LogTraceMessage TRACE_LEVEL_STATUS, "Progress", "UpdateChatGPTProgress", visibleStatusText, buildDetail, traceStepText, 0, vbNullString, vbNullString, buildForm, "Form", "CodeDb", "UI"
End Sub

' Purpose: Performs UI render shared progress activity by coordinating `SharedProgress_TryPulse` and `SetupControlUpdateStatus` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   progressActive (Input; Boolean): Indicates whether the caller already owns an active progress scope; True lets this routine update/reuse it, while False lets the routine avoid parent-progress assumptions.
'   semanticStatusText (Input; String): Semantic status text text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   statusBarText (Input; String): Status bar text text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   statusPrefix (Input; String): Prefix text prepended to or matched against the status content handled by this routine.
'   forceDisplay (Input; Boolean): When True, enables or requests display for this call; when False, the routine follows its normal non-forced/default path.
'   errNumber (Input/output; Long): Numeric identifier or ordinal for the err item processed or reported by this operation.
'   errSource (Input/output; String): Error source text carried into the returned/raised diagnostic; blank allows the procedure to supply its own source when coded to do so.
'   errDescription (Input/output; String): Error description text carried into the returned/raised diagnostic; blank allows the procedure to supply its own description when coded to do so.
' Returns: Boolean: True when UI text mismatch detail completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or not-applicable path.
Public Function UIRenderSharedProgressActivity(ByVal progressActive As Boolean, ByVal semanticStatusText As String, ByVal statusBarText As String, ByVal statusPrefix As String, ByVal forceDisplay As Boolean, ByRef errNumber As Long, ByRef errSource As String, ByRef errDescription As String) As Boolean
    'Physical shared-progress presentation authority: status bar plus active Setup Control surfaces.
    Dim uiErrNumber As Long
    Dim uiErrSource As String
    Dim uiErrDescription As String
    UIRenderSharedProgressActivity = SharedProgress_TryPulse(progressActive, statusBarText, statusPrefix, errNumber, errSource, errDescription)
    On Error Resume Next
    Call SetupControlUpdateStatus(semanticStatusText, False, forceDisplay, False)
    uiErrNumber = Err.Number
    uiErrSource = Err.Source
    uiErrDescription = Err.Description
    Err.Clear
    Call SetupControlPump(forceDisplay)
    If uiErrNumber <> 0 And errNumber = 0 Then
        errNumber = uiErrNumber
        errSource = uiErrSource
        errDescription = uiErrDescription
        UIRenderSharedProgressActivity = False
    End If
    On Error GoTo 0
End Function

' Purpose: Performs UI end shared progress by coordinating `SharedProgress_TryEnd` and `End` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   progressActive (Input/output; Boolean): Supplies whether the shared progress scope is active and receives False after the progress UI/state is ended or normalized.
'   currentStep (Input/output; Long): Supplies the caller's current progress step and receives 0 after the shared progress UI/state is ended or normalized.
'   maxSteps (Input/output; Long): Supplies the caller's current progress maximum and receives 0 after the shared progress UI/state is ended or normalized.
'   clearStatus (Input; Boolean): When True, enables or requests status for this call; when False, the routine follows its normal non-forced/default path.
'   errNumber (Output; Long): Receives 0 when the shared progress end/reset succeeds or the captured UI/status-bar error number when it does not.
'   errSource (Output; String): Receives vbNullString on success or the captured UI/status-bar error source when the shared progress end/reset fails.
'   errDescription (Output; String): Receives vbNullString on success or the captured UI/status-bar error description, including cleanup detail when applicable, on failure.
' Returns: Boolean: True when the shared progress end/reset completed without a display error; False when the adapter captured a UI/status-bar failure in the documented ByRef error outputs.
Public Function UIEndSharedProgress(ByRef progressActive As Boolean, ByRef currentStep As Long, ByRef maxSteps As Long, ByVal clearStatus As Boolean, ByRef errNumber As Long, ByRef errSource As String, ByRef errDescription As String) As Boolean
    ' Wrapper purpose: UI-owned facade that performs the physical shared-progress close/reset operation.
    ' Why retained: Retained to keep presentation lifecycle ownership in UI so Core can end progress without depending on the lower-level shared progress implementation.
    'Physical shared-progress close/reset authority.
    UIEndSharedProgress = SharedProgress_TryEnd(progressActive, currentStep, maxSteps, clearStatus, errNumber, errSource, errDescription)
End Function

' Purpose: Sets visible status text for status text, force display.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   statusText (Input; String): Status/progress message supplied for the current operation; the routine forwards it to the applicable UI/progress and trace/reporting path.
'   forceDisplay (Input Optional; Boolean): When True, enables or requests display for this call; when False, the routine follows its normal non-forced/default path. Default: True.
'   pumpAfterDisplay (Input Optional; Boolean): When True, enables or requests after display for this call; when False, the routine follows its normal non-forced/default path. Default: False.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub SetVisibleStatusText(ByVal statusText As String, Optional ByVal forceDisplay As Boolean = True, Optional ByVal pumpAfterDisplay As Boolean = False)
    'One status authority: active Setup updates the status bar and running/paused progress form from the same elapsed-time source; non-Setup callers update only the status bar.
    statusText = Trim$(statusText)
    If Len(statusText) = 0 Then Exit Sub
    If mSetupControlActive Then
        SetupControlUpdateStatus statusText, True, forceDisplay, pumpAfterDisplay
    Else
        SetUIStatusBarText statusText
    End If
End Sub

' Purpose: Performs the public clear visible status text operation using the routine's existing ADC workflow and side effects.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub ClearVisibleStatusText()
    'Single physical clear authority for the Access status bar; Setup controller state is preserved and may repaint its current status on the next heartbeat.
    On Error Resume Next
    Application.SysCmd acSysCmdClearStatus
    Err.Clear
    On Error GoTo 0
End Sub

' Purpose: Resolves UI text mismatch detail from the supplied/current ADC state, using `FirstTextDifferencePosition` and `TextComparisonExcerpt` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   existingText (Input; String): Input existing text for UI text mismatch detail; is passed to `FirstTextDifferencePosition`; is passed to `CStr`.
'   generatedText (Input; String): Input generated text for UI text mismatch detail; is passed to `FirstTextDifferencePosition`; is passed to `CStr`.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function UITextMismatchDetail(ByVal existingText As String, ByVal generatedText As String) As String
    Dim differencePosition As Long
    differencePosition = FirstTextDifferencePosition(existingText, generatedText)
    UITextMismatchDetail = "Stored length/hash: " & CStr(Len(existingText)) & "/" & AdlerTextHash(existingText)
    UITextMismatchDetail = UITextMismatchDetail & vbCrLf & "Generated length/hash: " & CStr(Len(generatedText)) & "/" & AdlerTextHash(generatedText)
    UITextMismatchDetail = UITextMismatchDetail & vbCrLf & "First differing character position: " & CStr(differencePosition)
    UITextMismatchDetail = UITextMismatchDetail & vbCrLf & "Stored excerpt: " & TextComparisonExcerpt(existingText, differencePosition)
    UITextMismatchDetail = UITextMismatchDetail & vbCrLf & "Generated excerpt: " & TextComparisonExcerpt(generatedText, differencePosition)
End Function

' Purpose: Resolves UI visible success text from the supplied/current ADC state, using `RibbonStatusPartText` and `ChatGPTProgressPartText` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   statusText (Input; String): Status/progress message supplied for the current operation; the routine forwards it to the applicable UI/progress and trace/reporting path.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function UIVisibleSuccessText(ByVal statusText As String) As String
    statusText = Trim$(statusText)
    If Len(statusText) = 0 Then Exit Function
    If mRibbonProgressActive Then
        UIVisibleSuccessText = RibbonStatusPartText(statusText, False)
    Else
        UIVisibleSuccessText = ChatGPTProgressPartText(statusText, False)
    End If
End Function

' Purpose: Resolves UI screen context for message from the supplied/current ADC state, using `Unavailable` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function UIScreenContextForMessage() As String
    On Error Resume Next
    Dim detailText As String
    Dim activeFormName As String
    Dim ActiveControlName As String
    Dim previousControlName As String
    activeFormName = "(none)"
    ActiveControlName = "(none)"
    previousControlName = "(none)"
    Err.Clear
    activeFormName = Screen.ActiveForm.Name
    If Err.Number <> 0 Then activeFormName = "Unavailable (" & Err.Description & ")"
    Err.Clear
    ActiveControlName = Screen.ActiveControl.Name
    If Err.Number <> 0 Then ActiveControlName = "Unavailable (" & Err.Description & ")"
    Err.Clear
    previousControlName = Screen.previousControl.Name
    If Err.Number <> 0 Then previousControlName = "Unavailable (" & Err.Description & ")"
    detailText = "Active form: " & activeFormName & vbCrLf & "Active control: " & ActiveControlName & vbCrLf & "Previous control: " & previousControlName
    UIScreenContextForMessage = vbCrLf & "Screen context:" & vbCrLf & detailText
End Function


' Purpose: Writes UI status, coordinating `ChatGPTProgressVisibleText` and `SetUIStatusBarText` as the implementation requires.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   sourceProcedure (Input; String): Procedure name recorded as the source of trace, error, status, or diagnostic information emitted by this call.
'   statusText (Input; String): Status/progress message supplied for the current operation; the routine forwards it to the applicable UI/progress and trace/reporting path.
'   objectName (Input Optional; String): Physical Access object name identifying the database object targeted by the operation. Default: vbNullString.
'   contextName (Input Optional; String): Context label recorded with tracing/progress/errors so the caller can distinguish the logical operation producing the message. Default: "UI".
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub PrintUIStatus(ByVal sourceProcedure As String, ByVal statusText As String, Optional ByVal objectName As String = vbNullString, Optional ByVal contextName As String = "UI")
    On Error Resume Next
    statusText = Trim$(statusText)
    If Len(statusText) = 0 Then Exit Sub
    statusText = ChatGPTProgressVisibleText(statusText)
    SetUIStatusBarText statusText
    LogTraceMessage TRACE_LEVEL_STATUS, "UI", sourceProcedure, statusText, vbNullString, "UI status", 0, vbNullString, vbNullString, objectName, "Form", "CodeDb", contextName
    If mSetupControlActive Then
        SetupControlPump
    Else
        DoEvents
    End If
End Sub

' =============================================================================================================================================================================
' DURABLE UI RECOVERY STORAGE PROCEDURES
' =============================================================================================================================================================================

' Purpose: Attempts to perform parse UI durable boolean for source text.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   sourceText (Input; String): Input source text for try parse UI durable boolean; is tested by If StrComp(sourceText, "True", vbTextCompare) = 0; is tested by ElseIf StrComp(sourceText, "False", vbTextCompare) =
'     0.
'   parsedValue (Output; Boolean): Set True when the procedure determines that parsed value applies; set False when that condition is not met.
' Returns: Boolean: True when the attempt to parse UI durable boolean completes successfully; False for the routine's defined normal unsuccessful/not-applicable outcome.
Public Function TryParseUIDurableBoolean(ByVal sourceText As String, ByRef parsedValue As Boolean) As Boolean
    sourceText = Trim$(sourceText)
    If StrComp(sourceText, "True", vbTextCompare) = 0 Then
        parsedValue = True
        TryParseUIDurableBoolean = True
    ElseIf StrComp(sourceText, "False", vbTextCompare) = 0 Then
        parsedValue = False
        TryParseUIDurableBoolean = True
    End If
End Function

' Purpose: Reads one UI marker from the current Core durable-recovery store and restores the UI-specific typed payload.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   operationName (Output; String): Receives the resolved operation name produced from this procedure's validated inputs and current ADC state.
'   detailText (Output; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
'   recoveryType (Output; String): Receives the recovery type returned by `UIDurableNoteValue` after this procedure completes its validation/resolution work.
'   recoveryPayload (Output; String): Receives the recovery payload returned by `UIDurableNoteValue` after this procedure completes its validation/resolution work.
'   recoveryStatus (Output; String): Receives the recovery status returned by `UIDurableNoteValue` after this procedure completes its validation/resolution work.
' Returns: Boolean: True when UI durable recovery record exists in the inspected context; False when it does not.
Public Function UIDurableRecoveryRecordExists(ByRef operationName As String, ByRef detailText As String, ByRef recoveryType As String, ByRef recoveryPayload As String, ByRef recoveryStatus As String) As Boolean
    'Reads one UI marker from the current Core durable-recovery store and restores the UI-specific typed payload.
    On Error GoTo ErrHandler
    Dim corePayload As String
    Dim notesVersion As String
    operationName = vbNullString
    detailText = vbNullString
    recoveryType = vbNullString
    recoveryPayload = vbNullString
    recoveryStatus = vbNullString
    If CoreDurableRecoveryMarkerExists(UI_RECOVERY_KEY, operationName, detailText, corePayload) = False Then Exit Function
    notesVersion = UIDurableNoteValue(corePayload, "Version")
    recoveryStatus = UIDurableNoteValue(corePayload, "RecoveryStatus")
    recoveryType = UIDurableNoteValue(corePayload, "RecoveryType")
    recoveryPayload = UIDurableNoteValue(corePayload, "Payload")
    If StrComp(notesVersion, "1", vbBinaryCompare) <> 0 Then Err.Raise ERR_UI_DURABLE_RECOVERY_FAILED, "UIDurableRecoveryRecordExists", "The durable UI recovery marker has an unsupported payload version: " & IIf(Len(notesVersion) = 0, "(blank)", notesVersion)
    If StrComp(recoveryStatus, RECOVERY_STATUS_BLOCKED, vbTextCompare) <> 0 And StrComp(recoveryStatus, UI_RECOVERY_STATUS_PREPARED, vbTextCompare) <> 0 Then Err.Raise ERR_UI_DURABLE_RECOVERY_FAILED, "UIDurableRecoveryRecordExists", "The durable UI recovery marker has an unsupported UI status: " & IIf(Len(recoveryStatus) = 0, "(blank)", recoveryStatus)
    UIDurableRecoveryRecordExists = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, IIf(Len(Trim$(savedErrSource)) = 0, "UIDurableRecoveryRecordExists", savedErrSource), savedErrDescription
End Function

' Purpose: Deletes the current Core-backed UI marker and verifies that it is gone.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub DeleteUIDurableRecoveryRecord()
    'Deletes the current Core-backed UI marker and verifies that it is gone.
    On Error GoTo ErrHandler
    Dim operationName As String
    Dim detailText As String
    Dim recoveryType As String
    Dim recoveryPayload As String
    Dim recoveryStatus As String
    DeleteDurableRecoveryMarker UI_RECOVERY_KEY
    If UIDurableRecoveryRecordExists(operationName, detailText, recoveryType, recoveryPayload, recoveryStatus) Then Err.Raise ERR_UI_DURABLE_RECOVERY_FAILED, "DeleteUIDurableRecoveryRecord", "The durable UI recovery marker still exists after deletion verification."
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, IIf(Len(Trim$(savedErrSource)) = 0, "DeleteUIDurableRecoveryRecord", savedErrSource), savedErrDescription
End Sub

' Purpose: Persists the remaining UI recovery state through the current Core durable-recovery store.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   operationName (Input; String): Logical operation name written to progress, trace, warning, or error context for this call.
'   detailText (Input; String): Receives or supplies diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
'   recoveryType (Input; String): Recovery-marker/type token identifying which durable recovery payload/state is being written, read, or cleared.
'   recoveryPayload (Input; String): Serialized or structured payload carrying the recovery recovery/operation data that must be persisted or reapplied.
'   recoveryStatus (Input; String): Status value describing the recovery lifecycle/review/result state consumed or emitted by the routine.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub PersistUIDurableRecoveryRecord(ByVal operationName As String, ByVal detailText As String, ByVal recoveryType As String, ByVal recoveryPayload As String, ByVal recoveryStatus As String)
    'Persists the remaining UI recovery state through the current Core durable-recovery store.
    On Error GoTo ErrHandler
    Dim actualOperation As String
    Dim actualDetail As String
    Dim actualType As String
    Dim actualPayload As String
    Dim actualStatus As String
    Dim serializedPayload As String
    recoveryType = Trim$(recoveryType)
    If StrComp(recoveryStatus, RECOVERY_STATUS_BLOCKED, vbTextCompare) <> 0 And StrComp(recoveryStatus, UI_RECOVERY_STATUS_PREPARED, vbTextCompare) <> 0 Then Err.Raise ERR_UI_DURABLE_RECOVERY_FAILED, "PersistUIDurableRecoveryRecord", "Unsupported UI recovery status: " & recoveryStatus
    serializedPayload = UISerializeCoreRecoveryPayload(recoveryType, recoveryPayload, recoveryStatus)
    PersistDurableRecoveryMarker UI_RECOVERY_KEY, operationName, detailText, serializedPayload
    If UIDurableRecoveryRecordExists(actualOperation, actualDetail, actualType, actualPayload, actualStatus) = False Then Err.Raise ERR_UI_DURABLE_RECOVERY_FAILED, "PersistUIDurableRecoveryRecord", "The durable UI recovery marker was absent during write verification."
    If StrComp(actualOperation, operationName, vbBinaryCompare) <> 0 Or StrComp(actualDetail, detailText, vbBinaryCompare) <> 0 Or StrComp(actualType, recoveryType, vbBinaryCompare) <> 0 Or StrComp(actualPayload, recoveryPayload, vbBinaryCompare) <> 0 Or StrComp(actualStatus, recoveryStatus, vbBinaryCompare) <> 0 Then Err.Raise ERR_UI_DURABLE_RECOVERY_FAILED, "PersistUIDurableRecoveryRecord", "The durable UI recovery marker did not match the requested state after write verification."
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, IIf(Len(Trim$(savedErrSource)) = 0, "PersistUIDurableRecoveryRecord", savedErrSource), savedErrDescription
End Sub

' Purpose: Prepares UI durable recovery for operation name, recovery type.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   operationName (Input; String): Logical operation name written to progress, trace, warning, or error context for this call.
'   detailText (Input; String): Receives or supplies diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
'   recoveryType (Input; String): Recovery-marker/type token identifying which durable recovery payload/state is being written, read, or cleared.
'   recoveryPayload (Input; String): Serialized or structured payload carrying the recovery recovery/operation data that must be persisted or reapplied.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub PrepareUIDurableRecovery(ByVal operationName As String, ByVal detailText As String, ByVal recoveryType As String, ByVal recoveryPayload As String)
    On Error GoTo ErrHandler
    Dim previousBlocked As Boolean
    Dim previousOperation As String
    Dim previousDetail As String
    Dim previousType As String
    Dim previousPayload As String
    Dim previousOwned As Boolean
    Dim preparedOperation As String
    Dim preparedDetail As String
    Dim preparedType As String
    Dim existingOperation As String
    Dim existingDetail As String
    Dim existingType As String
    Dim existingPayload As String
    Dim existingStatus As String
    Dim existingRecord As Boolean
    Dim sameOwnedPreparationContext As Boolean
    previousBlocked = mUIRecoveryBlocked
    previousOperation = mUIRecoveryOperationName
    previousDetail = mUIRecoveryDetail
    previousType = mUIRecoveryType
    previousPayload = mUIRecoveryPayload
    previousOwned = mUIDurableRecoveryOwned
    preparedOperation = Trim$(operationName)
    preparedDetail = Trim$(detailText)
    preparedType = Trim$(recoveryType)
    If Len(preparedType) = 0 Then preparedType = UIRecoveryTypeForOperation(preparedOperation)
    If Len(preparedOperation) = 0 Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "PrepareUIDurableRecovery", "A durable UI recovery preparation cannot use a blank operation name."
    If Len(preparedType) = 0 Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "PrepareUIDurableRecovery", "A durable UI recovery preparation could not resolve a supported recovery type." & vbCrLf & "Operation: " & preparedOperation
    EnsureCoreRecoveryStorageReady
    existingRecord = UIDurableRecoveryRecordExists(existingOperation, existingDetail, existingType, existingPayload, existingStatus)
    If existingRecord Then
        sameOwnedPreparationContext = mUIDurableRecoveryOwned
        sameOwnedPreparationContext = sameOwnedPreparationContext And (mUIRecoveryBlocked = False)
        sameOwnedPreparationContext = sameOwnedPreparationContext And (StrComp(existingStatus, UI_RECOVERY_STATUS_PREPARED, vbTextCompare) = 0)
        sameOwnedPreparationContext = sameOwnedPreparationContext And (StrComp(Trim$(existingOperation), Trim$(mUIRecoveryOperationName), vbTextCompare) = 0)
        sameOwnedPreparationContext = sameOwnedPreparationContext And (StrComp(Trim$(existingOperation), preparedOperation, vbTextCompare) = 0)
        sameOwnedPreparationContext = sameOwnedPreparationContext And (StrComp(Trim$(existingType), Trim$(mUIRecoveryType), vbTextCompare) = 0)
        sameOwnedPreparationContext = sameOwnedPreparationContext And (StrComp(Trim$(existingType), preparedType, vbTextCompare) = 0)
        sameOwnedPreparationContext = sameOwnedPreparationContext And (StrComp(existingDetail, mUIRecoveryDetail, vbBinaryCompare) = 0)
        sameOwnedPreparationContext = sameOwnedPreparationContext And (StrComp(existingDetail, preparedDetail, vbBinaryCompare) = 0)
        sameOwnedPreparationContext = sameOwnedPreparationContext And (StrComp(existingPayload, mUIRecoveryPayload, vbBinaryCompare) = 0)
        sameOwnedPreparationContext = sameOwnedPreparationContext And (StrComp(existingPayload, recoveryPayload, vbBinaryCompare) = 0)
        If sameOwnedPreparationContext = False Then
            Err.Raise ERR_UI_RECOVERY_BLOCKED, "PrepareUIDurableRecovery", "A durable UI recovery record already exists and is not owned by this exact in-memory preparation context. The existing record was not modified." & vbCrLf & "Requested operation: " & preparedOperation & vbCrLf & "Requested recovery type: " & preparedType & vbCrLf & "Existing operation: " & IIf(Len(Trim$(existingOperation)) = 0, "(blank)", existingOperation) & vbCrLf & "Existing recovery type: " & IIf(Len(Trim$(existingType)) = 0, "(blank)", existingType) & vbCrLf & "Existing status: " & IIf(Len(Trim$(existingStatus)) = 0, "(blank)", existingStatus) & vbCrLf & "Existing detail: " & existingDetail & vbCrLf & "Requested detail matches existing: " & CStr(StrComp(existingDetail, preparedDetail, vbBinaryCompare) = 0) & vbCrLf & "Requested payload matches existing: " & CStr(StrComp(existingPayload, recoveryPayload, vbBinaryCompare) = 0)
        End If
    End If
    PersistUIDurableRecoveryRecord preparedOperation, preparedDetail, preparedType, recoveryPayload, UI_RECOVERY_STATUS_PREPARED
    mUIRecoveryBlocked = False
    mUIRecoveryOperationName = preparedOperation
    mUIRecoveryDetail = preparedDetail
    mUIRecoveryType = preparedType
    mUIRecoveryPayload = recoveryPayload
    mUIDurableRecoveryOwned = True
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    mUIRecoveryBlocked = previousBlocked
    mUIRecoveryOperationName = previousOperation
    mUIRecoveryDetail = previousDetail
    mUIRecoveryType = previousType
    mUIRecoveryPayload = previousPayload
    mUIDurableRecoveryOwned = previousOwned
    Err.Raise savedErrNumber, IIf(Len(Trim$(savedErrSource)) = 0, "PrepareUIDurableRecovery", savedErrSource), savedErrDescription
End Sub

' =============================================================================================================================================================================
' GENERIC GENERATED-FORM / RUNTIME CALLBACK PROCEDURES
' =============================================================================================================================================================================

' Purpose: Refreshes navigation pane after grouping.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub RefreshNavigationPaneAfterGrouping()
    On Error Resume Next
    Application.RefreshDatabaseWindow
    On Error GoTo 0
End Sub

' Purpose: Initializes UI trace run if needed, coordinating `RequireNoUnfinishedUIRecovery` and `TraceRunActive` as the implementation requires.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   contextName (Input; String): Context label recorded with tracing/progress/errors so the caller can distinguish the logical operation producing the message.
' Returns: Boolean: True on the function's affirmative/success branch for begin UI trace run if needed; False on its explicit negative, unavailable, rejected, or handled-failure branch.
Public Function BeginUITraceRunIfNeeded(ByVal contextName As String) As Boolean
    On Error GoTo ErrHandler
    RequireNoUnfinishedUIRecovery contextName
    If TraceRunActive() = False Then
        StartTraceRunForSetup contextName
        BeginUITraceRunIfNeeded = True
    End If
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "BeginUITraceRunIfNeeded", Err.Description
End Function

' Purpose: Finalizes UI trace run if owned, coordinating `End` as the implementation requires.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   traceRunOwnedByProcedure (Input; Boolean): Indicates whether this procedure started the active trace run; True makes it responsible for completing that run, while False leaves the caller-owned run open.
'   completeMessage (Input; String): Complete message text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   clearStatus (Input Optional; Boolean): When True, enables or requests status for this call; when False, the routine follows its normal non-forced/default path. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub CompleteUITraceRunIfOwned(ByVal traceRunOwnedByProcedure As Boolean, ByVal completeMessage As String, Optional ByVal clearStatus As Boolean = True)
    On Error Resume Next
    If traceRunOwnedByProcedure Then CompleteTraceRun completeMessage, clearStatus
End Sub

' Purpose: Resolves chat GPT same DB only message from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   actionText (Input; String): Input action text for chat GPT same DB only message; is passed to `ChatGPTSameDbOnlyMessage`.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function ChatGPTSameDbOnlyMessage(ByVal actionText As String) As String
    ChatGPTSameDbOnlyMessage = actionText & " is same-database only for now." & vbCrLf & vbCrLf
    ChatGPTSameDbOnlyMessage = ChatGPTSameDbOnlyMessage & "CodeDb: " & codeDb.Name & vbCrLf
    ChatGPTSameDbOnlyMessage = ChatGPTSameDbOnlyMessage & "HostDb: " & GetHostDb().Name & vbCrLf & vbCrLf
    ChatGPTSameDbOnlyMessage = ChatGPTSameDbOnlyMessage & "No ribbon rows or CustomRibbonID changes were made because the " & DEFAULT_UTILITY_LONG_NAME & " CodeDb and HostDb database files are different."
End Function

' Purpose: Resolves UI variant for message from the supplied/current ADC state, using `isMissing` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   value (Input Optional; Variant): Variant input representing value; the routine accepts the caller-provided value in the form required by its existing validation and conversion logic.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function UIVariantForMessage(Optional ByVal value As Variant) As String
    On Error GoTo ErrHandler
    If isMissing(value) Then
        UIVariantForMessage = "(missing)"
    ElseIf IsNull(value) Then
        UIVariantForMessage = "(Null)"
    Else
        UIVariantForMessage = CStr(value)
    End If
    Exit Function
ErrHandler:
    UIVariantForMessage = "(unavailable: " & Err.Description & ")"
End Function

Private Function UIFormExistsForMessage(ByVal formName As String) As String
    On Error GoTo ErrHandler
    UIFormExistsForMessage = CStr(SharedDao_AccessObjectExists(GetHostDb(), acForm, formName))
    Exit Function
ErrHandler:
    UIFormExistsForMessage = "Unknown (" & Err.Description & ")"
End Function

' Purpose: Opens a generated or configured Access form by name from generated UI buttons or form expressions.
' Public role: Form-code/event-expression entry point. Generated or managed Access forms can invoke this Public routine from form code or an event expression.
' Parameters:
'   formName (Input; String): Physical Access form name identifying the form to create, open, configure, validate, or inspect.
' Returns: Boolean: True when the condition represented by RetryUIRecovery is present under the routine's documented inputs/current state; False when that condition is absent.
Public Function ChatGPT_OpenFormByName(ByVal formName As String) As Boolean
    ' Wrapper purpose: Generated-UI/form-expression callback that ensures the requested form exists, builds generated forms when necessary, and then opens the form.
    ' Why retained: Retained because Access control expressions require a stable public Function callback rather than direct access to UI's private build/open implementation.
    'Opens a generated or configured Access form by name from generated UI buttons or form expressions.
    On Error GoTo ErrHandler
    If SharedDao_AccessObjectExists(GetHostDb(), acForm, formName) = False Then BuildGeneratedForms False, False
    DoCmd.OpenForm formName, acNormal
    ChatGPT_OpenFormByName = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim messageText As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    messageText = "Could not open form [" & formName & "]." & vbCrLf & vbCrLf & "Procedure: ChatGPT_OpenFormByName" & vbCrLf & "Form: " & formName & vbCrLf & "Form exists now: " & UIFormExistsForMessage(formName) & vbCrLf & "Error " & CStr(savedErrNumber) & ": " & savedErrDescription & TableDiagnosticsLocationForMessage()
    LogTraceMessage "Error", "UI", "ChatGPT_OpenFormByName", "Could not open generated form by name.", messageText, vbNullString, savedErrNumber, savedErrSource, savedErrDescription, formName, "Form", "CodeDb", "OpenForm"
    CoreShowErrorMessage messageText, vbCritical, DEFAULT_UI_MESSAGE_TITLE, "ChatGPT_OpenFormByName", "UI"
    ChatGPT_OpenFormByName = False
End Function

' Purpose: Executes the chat gpt clear current form filter managed-form/UI callback using the routine's existing workflow and reports results through its established return/side-effect contract.
' Public role: Form-code/event-expression entry point. Generated or managed Access forms can invoke this Public routine from form code or an event expression.
' Parameters: None.
' Returns: Boolean: True when the chat gpt clear current form filter condition holds under the routine's existing validation rules; False when it does not.
Public Function ChatGPT_ClearCurrentFormFilter() As Boolean
    ' Wrapper purpose: Generated-form callback that clears filter and ordering state on the active form and returns Access-compatible Boolean success.
    ' Why retained: Retained because Access control expressions require a stable public Function callback while the underlying UI operations remain implementation details.
    'Intent: Clears the active generated-form filter and returns whether the operation succeeded.
    On Error GoTo ErrHandler
    Dim messageText As String
    If Not Screen.ActiveForm Is Nothing Then
        Screen.ActiveForm.FilterOn = False
        Screen.ActiveForm.OrderByOn = False
    End If
    ChatGPT_ClearCurrentFormFilter = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    messageText = "Could not clear the currently open Access form filter/sort." & vbCrLf & vbCrLf & "Procedure: ChatGPT_ClearCurrentFormFilter" & UIScreenContextForMessage() & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description:" & vbCrLf & savedErrDescription & TableDiagnosticsLocationForMessage()
    LogTraceMessage "Error", "UI", "ChatGPT_ClearCurrentFormFilter", "Could not clear the currently open Access form filter/sort.", messageText, vbNullString, savedErrNumber, savedErrSource, savedErrDescription, "ClearFilter", "Command", "CodeDb", "GeneratedFormCommand"
    CoreShowErrorMessage messageText, vbCritical, DEFAULT_UI_MESSAGE_TITLE, "ChatGPT_ClearCurrentFormFilter", "UI"
    ChatGPT_ClearCurrentFormFilter = False
End Function

' Purpose: Resolves current form record text from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   frm (Input; Access.Form): Open Access Form instance whose controls, properties, or current record are validated or manipulated.
'   fieldName (Input; String): Canonical or physical name identifying the field targeted by this operation.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function CurrentFormRecordText(ByVal frm As Access.Form, ByVal fieldName As String) As String
    On Error GoTo MissingField
    If frm Is Nothing Then Exit Function
    CurrentFormRecordText = CStr(Nz(frm.Recordset.fields(fieldName).value, vbNullString))
    Exit Function
MissingField:
    Err.Clear
End Function

' Purpose: Performs the public retry UI recovery operation using the routine's existing ADC workflow and side effects.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine shows its user-facing completion, status, warning, or error message in addition to performing the underlying work. Default: True.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub RetryUIRecovery(Optional ByVal showMessage As Boolean = True)
    'UI recovery retry orchestration: decide whether recovery is pending, invoke the authoritative UI recovery engine, and report the verified result.
    On Error GoTo ErrHandler
    Dim resultText As String
    If UIRecoveryBlocked() = False Then
        If showMessage Then CoreShowMessage "No unfinished Access Development Console UI recovery record exists.", vbInformation, DEFAULT_UI_MESSAGE_TITLE, "RetryUIRecovery", "UI"
        Exit Sub
    End If
    resultText = ExecutePendingUIRecovery()
    If UIRecoveryBlocked() Then Err.Raise vbObjectError + 7210, "RetryUIRecovery", "Recovery work returned without clearing the authoritative UI recovery blocker."
    If showMessage Then CoreShowMessage "Access Development Console UI recovery completed and was verified." & vbCrLf & vbCrLf & resultText, vbInformation, DEFAULT_UI_MESSAGE_TITLE, "RetryUIRecovery", "UI"
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    If showMessage Then
        CoreShowErrorMessage "Access Development Console UI recovery remains unresolved." & vbCrLf & vbCrLf & "Error " & CStr(savedErrNumber) & ": " & savedErrDescription, vbCritical, DEFAULT_UI_MESSAGE_TITLE, "RetryUIRecovery", "UI"
        Exit Sub
    End If
    Err.Raise savedErrNumber, "RetryUIRecovery", savedErrDescription
End Sub

' Purpose: Determines whether valueText begins with a supported SQL command prefix (SELECT, TRANSFORM, or EXEC) after trimming and case normalization.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   valueText (Input; String): Input value text for looks like SQL; is passed to `upperText`.
' Returns: Boolean: True when valueText begins with SELECT, TRANSFORM, or EXEC after trimming/case normalization; False otherwise.
Public Function LooksLikeSql(ByVal valueText As String) As Boolean
    Dim upperText As String
    upperText = UCase$(Trim$(valueText))
    LooksLikeSql = (Left$(upperText, 7) = "SELECT " Or Left$(upperText, 10) = "TRANSFORM " Or Left$(upperText, 5) = "EXEC ")
End Function

' Purpose: Opens the named saved Access query from the generated UI entry point and reports whether the open operation succeeded.
' Public role: Form-code/event-expression entry point. Generated or managed Access forms can invoke this Public routine from form code or an event expression.
' Parameters:
'   queryName (Input; String): Physical Access query name identifying the saved query the routine creates, validates, reads, or updates.
' Returns: Boolean: True when chat gpt open query by name completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or not-applicable path.
Public Function ChatGPT_OpenQueryByName(ByVal queryName As String) As Boolean
    ' Wrapper purpose: Generated-UI callback that validates a saved query name/existence and opens the query read-only.
    ' Why retained: Retained because Access control expressions require a stable public Function callback that normalizes validation, opening, and errors.
    On Error GoTo ErrHandler
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    If Len(Trim$(queryName)) = 0 Then Err.Raise vbObjectError + 6741, "ChatGPT_OpenQueryByName", "The query name is blank."
    If SharedDao_AccessObjectExists(GetHostDb(), acQuery, queryName) = False Then Err.Raise vbObjectError + 6742, "ChatGPT_OpenQueryByName", "The requested saved query does not exist: " & queryName
    DoCmd.OpenQuery queryName, acViewNormal, acReadOnly
    ChatGPT_OpenQueryByName = True
    Exit Function
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    LogTraceMessage "Error", "UI", "ChatGPT_OpenQueryByName", "A generated-form query-open action could not complete.", "Query: " & queryName & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, "Open generated-form query", savedErrNumber, savedErrSource, savedErrDescription, queryName, "Query", "CodeDb", "GeneratedFormCommand"
    CoreShowErrorMessage "Could not open query [" & queryName & "]." & vbCrLf & vbCrLf & savedErrDescription, vbCritical, DEFAULT_UI_MESSAGE_TITLE, "ChatGPT_OpenQueryByName", "UI"
End Function

' Purpose: Resolves and invokes the named public ADC procedure from the generated UI entry point, preserving the routine’s existing argument and error-handling rules.
' Public role: Form-code/event-expression entry point. Generated or managed Access forms can invoke this Public routine from form code or an event expression.
' Parameters:
'   ProcedureName (Input; String): Identifies the VBA procedure being searched, validated, described, or invoked by its procedure name.
' Returns: Boolean: True when chat gpt open configuration detail from list completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or
'   not-applicable path.
Public Function ChatGPT_RunProcedure(ByVal ProcedureName As String) As Boolean
    ' Wrapper purpose: Generated-UI callback that invokes a named ADC procedure while handling active-form state and normalizing callback errors.
    ' Why retained: Retained because generated Access controls need a stable public Function contract instead of invoking private procedure-dispatch machinery directly.
    On Error GoTo ErrHandler
    Dim frm As Access.Form
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    If Len(Trim$(ProcedureName)) = 0 Then Err.Raise vbObjectError + 6740, "ChatGPT_RunProcedure", "The Access Development Console procedure name is blank."
    On Error Resume Next
    Set frm = Screen.ActiveControl.Parent
    If Not frm Is Nothing Then If frm.Dirty Then frm.Dirty = False
    Err.Clear
    On Error GoTo ErrHandler
    Application.Run ProcedureName, False
    If Not frm Is Nothing Then frm.Requery
    ChatGPT_RunProcedure = True
    Exit Function
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    LogTraceMessage "Error", "UI", "ChatGPT_RunProcedure", "A generated-form Access Development Console action could not complete.", "Procedure: " & ProcedureName & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, "Generated form action", savedErrNumber, savedErrSource, savedErrDescription, ProcedureName, "Procedure", "CodeDb", "GeneratedFormCommand"
    CoreShowErrorMessage "Could not run Access Development Console procedure [" & ProcedureName & "]." & vbCrLf & vbCrLf & savedErrDescription, vbCritical, DEFAULT_UI_MESSAGE_TITLE, "ChatGPT_RunProcedure", "UI"
End Function

' Purpose: Executes the chat gpt filter current selection managed-form/UI callback using the routine's existing workflow and reports results through its established return/side-effect contract.
' Public role: Form-code/event-expression entry point. Generated or managed Access forms can invoke this Public routine from form code or an event expression.
' Parameters: None.
' Returns: Boolean: True when the chat gpt filter current selection condition holds under the routine's existing validation rules; False when it does not.
Public Function ChatGPT_FilterCurrentSelection() As Boolean
    ' Wrapper purpose: Generated-form callback that applies the current selection as a filter and returns whether Access applied the filter successfully.
    ' Why retained: Retained because Access control expressions require a stable public Function callback around the command operation and its success/error contract.
    'Intent: Applies the current generated-form selection as a filter and returns whether the filter was applied.
    On Error GoTo ErrHandler
    Dim previousControl As Access.control
    Dim messageText As String
    On Error Resume Next
    Set previousControl = Screen.previousControl
    If Not previousControl Is Nothing Then previousControl.SetFocus
    On Error GoTo ErrHandler
    DoCmd.RunCommand acCmdFilterBySelection
    ChatGPT_FilterCurrentSelection = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    messageText = "Could not apply Filter By Selection. Select a datasheet or form value first, then run the filter action." & vbCrLf & vbCrLf & "Procedure: ChatGPT_FilterCurrentSelection" & UIScreenContextForMessage() & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description:" & vbCrLf & savedErrDescription & TableDiagnosticsLocationForMessage()
    LogTraceMessage "Error", "UI", "ChatGPT_FilterCurrentSelection", "Could not apply Filter By Selection.", messageText, vbNullString, savedErrNumber, savedErrSource, savedErrDescription, "FilterBySelection", "Command", "CodeDb", "GeneratedFormCommand"
    CoreShowErrorMessage messageText, vbCritical, DEFAULT_UI_MESSAGE_TITLE, "ChatGPT_FilterCurrentSelection", "UI"
    ChatGPT_FilterCurrentSelection = False
End Function

' Purpose: Opens the configuration-detail form for the row selected in the generated configuration list.
' Public role: Form-code/event-expression entry point. Generated or managed Access forms can invoke this Public routine from form code or an event expression.
' Parameters:
'   rowID (Input Optional; Variant): Persistent row identifier selecting the catalog, review, configuration, or history record affected by the operation.
' Returns: Boolean: True when chat gpt open configuration detail from list completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or
'   not-applicable path.
Public Function ChatGPT_OpenConfigurationDetailFromList(Optional ByVal rowID As Variant) As Boolean
    ' Wrapper purpose: Generated-list callback that ensures the configuration detail form exists and opens either the general detail surface or the selected row.
    ' Why retained: Retained because Access control expressions require a stable public Function callback around generated-form preparation and row navigation.
    'Intent: Opens the selected configuration row in the generated detail surface and returns whether navigation succeeded.
    On Error GoTo ErrHandler
    Dim messageText As String
    If SharedDao_AccessObjectExists(GetHostDb(), acForm, ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Configuration_Detail")) = False Then BuildGeneratedForms False, False
    If isMissing(rowID) Or IsNull(rowID) Then
        DoCmd.OpenForm ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Configuration_Detail"), acNormal
    Else
        DoCmd.OpenForm ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Configuration_Detail"), acNormal, , SQL_FIELD_ID & "=" & CStr(CLng(rowID))
    End If
    ChatGPT_OpenConfigurationDetailFromList = True
    Exit Function
ErrHandler:
    messageText = "Could not open the ChatGPT_ configuration detail form." & vbCrLf & vbCrLf & "Procedure: ChatGPT_OpenConfigurationDetailFromList" & vbCrLf & "Form: " & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Configuration_Detail") & vbCrLf & "Row ID: " & UIVariantForMessage(rowID) & vbCrLf & "Form exists now: " & UIFormExistsForMessage(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Configuration_Detail")) & vbCrLf & "Error " & CStr(Err.Number) & ": " & Err.Description & TableDiagnosticsLocationForMessage()
    LogTraceMessage "Error", "UI", "ChatGPT_OpenConfigurationDetailFromList", "Could not open generated configuration detail form.", messageText, vbNullString, Err.Number, Err.Source, Err.Description, ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Configuration_Detail"), "Form", "CodeDb", "OpenForm"
    CoreShowErrorMessage messageText, vbCritical, DEFAULT_UI_MESSAGE_TITLE, "ChatGPT_OpenConfigurationDetailFromList", "UI"
    ChatGPT_OpenConfigurationDetailFromList = False
End Function
