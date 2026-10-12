Attribute VB_Name = "x__UTIL_ADC_MACRO_RUNTIME"
' ================================================================================================================================================================================
' ChatGPT Rewrite Tracking Header
' ================================================================================================================================================================================
' ChatGPT Previous Version ID: CGPT-P2N4-JEQG-V8KD-36SJ-MCG2
' ChatGPT Baseline Version ID: CGPT-P2N4-JEQG-V8KD-36SJ-MCG2
' ChatGPT Version ID: CGPT-JCMB-X27Q-H6X9-29WD-KM9A
' ChatGPT Version Timestamp: Tuesday 9/1/26 11:32 AM
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
' STANDARD MODULE: Access Development Console Macro Runtime
' VB_Name: x__UTIL_ADC_MACRO_RUNTIME
' Entry Points:
'     GenerateConfiguredMacro
'     ResolveConfiguredTarget
'     RunConfiguredMacroTarget
'     RunConfirmedConfiguredMacroTarget
' SHARED CODE:
'     SharedAccess_PrepareOpenObjectForDesignChange(ByVal objectType As AcObjectType, ByVal objectName As String, ByVal actionText As String, Optional ByVal confirmationAlreadyProvided As Boolean = False) -> Boolean
'     SharedAccess_PrepareOpenObjectsForDesignChange(ByVal objectType As AcObjectType, ByVal objectNames As Collection, ByVal actionText As String, Optional ByVal confirmationAlreadyProvided As Boolean = False) ->
'     Boolean
'     SharedCollection_Join(ByVal values As Collection, Optional ByVal delimiter As String = ", ") -> String
'     SharedCollection_KeyExists(ByVal values As Collection, ByVal key As String) -> Boolean
'     SharedDao_DatabaseTableExistsSafe(ByVal db As DAO.Database, ByVal tableName As String) -> Boolean
'     SharedDao_SafeDatabaseName(ByVal db As DAO.Database) -> String
'     SharedDao_SqlQuote(ByVal value As Variant) -> String
'     SharedDao_TableDefFieldExistsSafe(ByVal tdf As DAO.tableDef, ByVal fieldName As String) -> Boolean
'     SharedError_FormatDetails(ByVal ProcedureName As String, Optional ByVal contextText As String = vbNullString) -> String
'     SharedMessage_OneLine(ByVal valueText As String, Optional ByVal maxLength As Long = 700, Optional ByVal ellipsisText As String = "...") -> String
'     SharedName_SafeAccessObjectName(ByVal value As String, Optional ByVal maxLength As Long = 0, Optional ByVal firstCharacterPrefix As String = "DB_", Optional ByVal fallbackName As String = vbNullString) ->
'     String
'     SharedName_SafeFileName(ByVal value As String, Optional ByVal fallbackName As String = "AccessObject", Optional ByVal maxLength As Long = 120) -> String
'     SharedRun_ApplicationRunWithArgs(ByVal ProcedureName As String, ByVal args As Collection) -> Variant
'     SharedText_AfterFirstColon(ByVal sourceText As String) -> String
'     SharedText_AppendNonBlankLine(ByVal existingText As String, ByVal lineText As String) -> String
'     SharedText_DoubleQuote(ByVal value As String) -> String
'     SharedText_Nz(ByVal v As Variant) -> String
'     Shared_String_ConvertToNullIfBlank(ByVal v As Variant) -> Variant
'     SharedValue_NzLong(ByVal value As Variant, Optional ByVal defaultValue As Long = 0) -> Long
' CLASSES USED: None
' DYNAMIC STRING REFERENCES:
'     RunCode string reference: messageText = "Macro procedure validation complete." & vbCrLf & vbCrLf & "Configured but missing macro object: " & CStr(configuredMissingMacroEntries.count) & vbCrLf & "Missing: " &
'     CStr(missingEntries.count) & vbCrLf & "Unresolved: " & CStr(unresolvedEntries.count) & vbCrLf & "Ambiguous: " & CStr(ambiguousEntries.count) & vbCrLf & "VBE Access Error: " & CStr(vbeAccessErrorEntries.count)
'     & vbCrLf & "Valid: " & CStr(validEntries.count) & vbCrLf & "Non-RunCode: " & CStr(nonRunCodeEntries.count) & vbCrLf & "Export Error: " & CStr(exportErrorEntries.count) & vbCrLf & vbCrLf & "Report:" & vbCrLf &
'     reportPath
'     RunCode string reference: nonRunCodeEntries.Add MacroValidationReportEntry(MacroTitle, "NON-RUNCODE", "", "", "No RunCode or OpenForm action was found in the exported macro text for this Active=True configured
'     macro object." & vbCrLf & "MacroKey: " & cfg.MacroKey & vbCrLf & "Configured action: " & cfg.MacroAction)
'     RunCode string reference: If InStr(1, lineText, "Action", vbTextCompare) > 0 And InStr(1, lineText, """RunCode""", vbTextCompare) > 0 Then
'     Dynamic-reference scan stopped after 2000 logical line(s) to prevent a long-running scan; review this module manually if needed.
' LastScanTimestamp: Saturday 8/15/26 8:43 AM
' Purpose:
'     Owns the post-candidate configured-macro lifecycle: MacroCatalog_Records readiness and lookup, configured-target resolution and validation, macro generation and execution, and Custom Macro Launcher
'     integration that shares generated-macro rollback state.
' Author: Christina Shonts
' Support:
'     Access Development Console configured-macro, configured-target, and Custom Macro Launcher runtime workflows.
' Code Category: Access Development Console configured-macro operational authority.
' Requirements:
'     Microsoft Access, DAO, trusted access to the VBA project object model for configured-target procedure inspection, and the companion Access Development Console modules.
' Reference Dependencies:
'     Access object model, DAO, VBE project model where configured-target procedure inspection requires it, and companion ADC modules.
' DB Object Dependencies:
'     Uses the current ADC configuration and Custom Macro Launcher review storage created/repaired by their canonical storage owners; candidate/discovery storage is owned separately by
'     MC_Candidates, and this module does not own physical table schema creation.
' Usage:
'     Core System Command dispatch routes configured-macro generation/execution commands here; DB Inventory candidate commands are owned by
'     Generated forms and Access macros may call the intentional public single-record/runtime entry points directly.
' Version:
'     Shifts current managed Access macro Description persistence ownership from F1 to G1 after insertion of Setup dependency Phase F.
'     Uses exact physical macro names when classifying ADC-owned System Macros at runtime.
'     Uses configured TargetObjectName exactly for object actions instead of trimming the physical Access object name before execution.
'     Consolidates macro export parsing, RunCode target parsing, configured-target counting/display, and macro escaping helpers into Macro Runtime without changing runtime behavior.
'     Renames x__UTIL_ADC_CONFIG to x__UTIL_ADC_MACRO_RUNTIME and adopts the shared TMacroRecord/MacroRecordFromRecordset names; configured-macro runtime behavior and Setup-phase ownership are unchanged.
'     Removes the direct Trim$ mutation from managed Access macro Description persistence; F1 ownership and the existing property-size policy are unchanged.
'     Routes Managed Object Catalog references to x__UTIL_ADC_MOC and physical lifecycle references to x__UTIL_ADC_MOC_LIFECYCLE after the responsibility split; runtime and Setup-phase behavior are unchanged.
'     Uses the canonical CodeDb role constant directly and narrows module-internal procedure/constant visibility.
'     Added the authoritative operation-scoped public-callable procedure lookup used by strict D4 System Macro wrapper validation, eliminating one full VBProject search per wrapper while preserving duplicate,
'     signature, module-type, visibility, and fallback-text semantics.
'     Managed-header dependency metadata updated because x__UTIL_ADC_DEFAULTS is now a standard module; executable behavior unchanged.
'     Routed HostDb resolution, Settings persistence, System Command cancellation state, and ordinary recovery popups through their owning authorities.
'     Removed stale-preview HealthReview run/persistence ownership and generic Settings metadata fallback; Config retains stale-candidate analysis, delegates findings to Review, and fails if authoritative Settings
'     readiness cannot supply the required row.
'     Candidate discovery/acceptance ownership moved to x__UTIL_ADC_DB_INVENTORY_CANDIDATES; this module retains Config operational readiness/lookup/resolution/validation, configured macro generation/execution,
'     and coherent Custom Macro Launcher integration.
' Notes:
'     Physical MacroCatalog_Records table/schema authority remains in  Physical Configuration Editor form construction remains solely in x__FORM_SUBFORM_SET.
'     Managed-object identity remains in x__UTIL_ADC_MOC; review/health findings remain in x__UTIL_ADC_REVIEW; physical Setup Control/status presentation remains in
' Core Rules Enforced Here:
'     One post-candidate configured-macro authority; no compatibility wrappers, duplicate parsers, second physical MacroCatalog_Records schema owner, or second Configuration Editor builder authority.
' Design Map:
'     Config lookup/readiness -> target resolution/validation -> configured macro generation -> configured target execution; Custom Macro Launcher generation remains colocated because it shares configured-macro
'     rollback/generation state.
' Configuration:
'     Current DEFAULTS, Settings, and table/catalog authorities remain authoritative for their respective stored definitions and schemas.
' Behavior:
'     Preserves existing public form/macro-facing contracts, generated-macro safety/rollback behavior, configured-target validation, and command error propagation to Core.
' Code Status: Active canonical configured-macro runtime authority.
' Todo:
'     Regenerate managed-header public/shared/dynamic reference inventories after import if the normal ADC header-maintenance workflow requires it.
' Reporting Bugs: None currently recorded in this header.
' Module Level Bugs: None currently recorded in this header.
' Public Subs:
'     RequireConfigTableForUtilityWorkflow(ByVal sourceName As String) -> None
'     GenerateSelectedRunnerMacrosForActiveRoute(ByVal callerName As String, ByVal operationCaption As String) -> None
'     GenerateAllConfiguredRunnerMacrosForActiveRoute(ByVal callerName As String, ByVal operationCaption As String) -> None
'     GeneratePromptedRunnerMacrosForActiveRoute(ByVal callerName As String, ByVal operationCaption As String, ByVal completionMessageTitle As String) -> None
'     ValidateRunnerMacroTargets() -> None
'     BeginPublicCallableProcedureLookupScope(ByVal ConfigOwnerDatabaseRole As String) -> None
'     EndPublicCallableProcedureLookupScope() -> None
'     GetPublicCallableProcedureLookupScopeStats(ByRef distinctProcedureNamesIndexed As Long, ByRef vbComponentsEnumerated As Long, ByRef completeProjectEnumerationPasses As Long, ByRef fallbackSourceTextScans
'     As Long) -> None
'     ClearVBEProcedureSearchError() -> None
'     DisableCustomMacroLauncherAutomation(Optional ByVal showMessage As Boolean = True) -> None
'     EnableCustomMacroLauncherAutomation(ByVal showMessage As Boolean) -> None
'     BuildRepairCustomMacroLaunchers(ByVal showMessage As Boolean) -> None
'     BuildUpdateCustomMacroLauncherButtons(ByVal showMessage As Boolean) -> None
'     ReviewCustomMacroLauncherButtons(ByVal showMessage As Boolean) -> None
'     CleanStaleCustomMacroLauncherButtons(ByVal showMessage As Boolean) -> None
'     GenerateConfiguredMacro(ByVal ConfigID As Long, Optional ByVal showMessage As Boolean = True) -> None
'     SetMacroGenerationCompletionDeferral(ByVal deferCompletion As Boolean, Optional ByVal showPendingCompletion As Boolean = False) -> None
'     DeleteSelectedDuplicateRows(Optional ByVal showMessage As Boolean = True) -> None
'     ResolveConfiguredTarget(ByVal ConfigID As Long, Optional ByVal showMessage As Boolean = True) -> None
'     ResolveRunnerMacroTargets(Optional ByVal showMessage As Boolean = True) -> None
' Public Functions:
'     MacroExportLineNameValue(ByVal lineText As String) -> String
'     RunCodeTargetsFromMacroText(ByVal macroText As String) -> Collection
'     ConfiguredTargetArgumentCount(ByVal argumentsText As String) -> Long
'     EscapeMacroText(ByVal value As String) -> String
'     FormatDisplayValue(ByVal value As Variant) -> String
'     ConfigMacroKeyExists(ByVal MacroKey As String) -> Boolean
'     TryExportMacroText(ByVal MacroTitle As String, ByRef macroText As String, ByRef errorText As String) -> Boolean
'     VBEProcedureSearchErrorOccurred() -> Boolean
'     VBEProcedureSearchErrorText() -> String
'     VBEProcedureSearchAmbiguityOccurred() -> Boolean
'     VBEProcedureSearchAmbiguityText() -> String
'     RunConfirmedConfiguredMacroTarget(ByVal MacroKey As String, Optional ByVal showMessage As Boolean = True) -> Boolean
'     RunConfiguredMacroTarget(ByVal MacroKey As String, Optional ByVal showMessage As Boolean = True) -> Boolean
'     PublicCallableProcedureInfoInStandardModule(ByVal ProcedureName As String, ByVal ConfigOwnerDatabaseRole As String, ByRef procedureKind As String, ByRef procedureLocation As String, ByRef argumentSignature
'     As String, ByRef RequiredArgumentCount As Long, ByRef OptionalArgumentCount As Long, ByRef hasParamArray As Boolean, ByRef ExpectedReturnType As String) -> Boolean
'     OpenFormTargetsFromMacroText(ByVal macroText As String) -> Collection
'     CanAccessVBEProjectModel(ByRef detailText As String) -> Boolean
'     ExpectedGeneratedMacroObjectNameForConfig(ByRef cfg As TMacroRecord) -> String
'     GeneratedRunCodeMacroRequiresConfirmation(ByRef cfg As TMacroRecord) -> Boolean
'     GeneratedResultIndicatesProblem(ByVal resultText As String) -> Boolean
'     ConfiguredObjectActionTargetName(ByRef cfg As TMacroRecord) -> String
'     FormatErrorWithContext(ByVal summaryText As String, Optional ByVal contextText As String = vbNullString) -> String
' Public Properties: None
' ================================================================================================================================================================================

' Public API Notes:
'     GenerateConfiguredMacro and ResolveConfiguredTarget are intentional external Config contracts; RunConfiguredMacroTarget and RunConfirmedConfiguredMacroTarget remain generated-macro/runtime contracts.
'     Candidate-facing contracts are owned by

Option Compare Database
Option Explicit

' Callable-procedure search diagnostics and authoritative operation-scoped lookup.
Private mLastVBEProcedureSearchError As String
Private mLastVBEProcedureSearchAmbiguity As String
' INTENTIONAL AUTHORITATIVE REUSABLE CACHE:
' MACRO_RUNTIME owns callable-procedure discovery. This lookup is a single stored projection for one explicit validation operation; Begin invalidates it, first use builds it lazily, all same-role consumers
' reuse it, and End discards it.
Private mPublicCallableProcedureLookupScopeActive As Boolean
Private mPublicCallableProcedureLookupRole As String
Private mPublicCallableProcedureLookupTreatOptionPrivateAsNotCallable As Boolean
Private mPublicCallableProcedureLookupBuilt As Boolean
Private mPublicCallableProcedureLookup As Object
Private mPublicCallableProcedureLookupBuildErrorText As String
Private mPublicCallableProcedureLookupDistinctNameCount As Long
Private mPublicCallableProcedureLookupVBComponentsEnumerated As Long
Private mPublicCallableProcedureLookupProjectEnumerationPasses As Long
Private mPublicCallableProcedureLookupFallbackSourceTextScans As Long

' Configured-macro replacement error contract.
Private Const MACRO_REPLACEMENT_RECOVERY_FAILURE_ERROR_NUMBER As Long = vbObjectError + 5951

' Launcher-control generation transaction state. This remains one coherent subsystem with configured-macro generation.
Private mLauncherGenerationTrackingActive As Boolean
' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION:
' These run-owned Collections track generated configuration identities and exact macro/form rollback payloads for one launcher-generation transaction. They separate completed generation work from recovery evidence,
' preserve rollback/processing order, and are cleared only by the existing completion, rollback, or explicit-acceptance paths; callers do not receive the mutable Collections.
Private mLauncherGeneratedConfigIDs As Collection
Private mGeneratedMacroRollbackEntries As Collection
Private mLauncherFormRollbackEntries As Collection
Private mGeneratedMacroRollbackRecoveryRequired As Boolean

' Detailed generation completion is deferred only while the operational-command layer owns the shared modal progress dialog.
Private mDeferMacroGenerationCompletionMessage As Boolean
Private mDeferredMacroGenerationCompletionPending As Boolean
Private mDeferredMacroGenerationSummaryTitle As String
Private mDeferredMacroGenerationGeneratedCount As Long
Private mDeferredMacroGenerationSkippedCount As Long
Private mDeferredMacroGenerationErrorEntries As Collection
Private mDeferredMacroGenerationMessageTitle As String
Private mDeferredMacroGenerationExtraDetailText As String

' ===============================================================================================================================================================================
' MACRO RUNTIME PROCEDURES
' ===============================================================================================================================================================================


' ===============================================================================================================================================================================
' CONFIGURED TARGET VALIDATION / EXECUTION PROCEDURES
' ===============================================================================================================================================================================

Private Function SafeHostDbNameForError() As String
    On Error GoTo Fallback
    SafeHostDbNameForError = GetHostDb().Name
    Exit Function
Fallback:
    SafeHostDbNameForError = "(HostDb unavailable: " & Err.Description & ")"
End Function

Private Function SafeCodeDbNameForError() As String
    On Error GoTo Fallback
    SafeCodeDbNameForError = codeDb.Name
    Exit Function
Fallback:
    SafeCodeDbNameForError = "(CodeDb unavailable: " & Err.Description & ")"
End Function

Private Function SafeEffectiveTargetCallTextForError(ByRef cfg As TMacroRecord) As String
    On Error GoTo Fallback
    SafeEffectiveTargetCallTextForError = EffectiveTargetCallText(cfg)
    Exit Function
Fallback:
    SafeEffectiveTargetCallTextForError = "(target call unavailable: " & Err.Description & ")"
End Function

Private Function SystemMacroDetailsForMessage(ByVal systemMacroDetailsText As String) As String
    If Len(Trim$(systemMacroDetailsText)) = 0 Then Exit Function
    SystemMacroDetailsForMessage = vbCrLf & vbCrLf & "System macro details:" & vbCrLf & systemMacroDetailsText
End Function

Private Function OwnedSystemMacroTitle(ByVal macroName As String) As Boolean
    Dim rowData As Variant
    For Each rowData In SystemMacroCatalog()
        PumpSetupResponsiveness
        If StrComp(macroName, SystemMacroCatalogTitle(rowData), vbTextCompare) = 0 Then
            OwnedSystemMacroTitle = True
            Exit Function
        End If
    Next rowData
End Function


Private Function EffectiveTargetCallText(ByRef cfg As TMacroRecord) As String
    EffectiveTargetCallText = ProcedureCallText(QualifiedRunProcedureName(cfg), cfg.TargetArguments)
End Function

Private Function DatabaseFolderPath() As String
    Dim dbPath As String
    Dim slashPos As Long
    dbPath = Trim$(ADCDb().Name)
    slashPos = InStrRev(dbPath, "\")
    If slashPos > 1 Then
        DatabaseFolderPath = Left$(dbPath, slashPos - 1)
    Else
        DatabaseFolderPath = Application.CurrentProject.Path
    End If
End Function

Private Function OppositeConfigOwnerDatabaseRole(ByVal ConfigOwnerDatabaseRole As String) As String
    Select Case LCase$(NormalizedConfigOwnerDatabaseRole(ConfigOwnerDatabaseRole))
        Case LCase$(CONFIG_OWNER_DATABASE_ROLE_CODE_DB)
            OppositeConfigOwnerDatabaseRole = CONFIG_OWNER_DATABASE_ROLE_HOST_DB
        Case Else
            OppositeConfigOwnerDatabaseRole = CONFIG_OWNER_DATABASE_ROLE_CODE_DB
    End Select
End Function

' Purpose: Generates selected runner macros for active route for caller name, operation caption.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   callerName (Input; String): Procedure/caller name recorded for diagnostics and ownership decisions so nested operations can be attributed correctly.
'   operationCaption (Input; String): Operation caption text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub GenerateSelectedRunnerMacrosForActiveRoute(ByVal callerName As String, ByVal operationCaption As String)
    On Error GoTo ErrHandler
    Dim progressStarted As Boolean
    If Len(callerName) = 0 Then callerName = "GenerateSelectedRunnerMacrosForActiveRoute"
    If Len(operationCaption) = 0 Then operationCaption = "Generate Selected Runner Macros"
    progressStarted = CoreBeginOperationProgressIfNeeded(operationCaption, 2)
    CoreUpdateSetupProgress "Check default Navigation Pane group"
    RequireDefaultCodeDbGeneratedMacrosNavigationPaneGroupForRoleContext callerName, False
    CoreUpdateSetupProgress "Generate selected Runner Macros"
    GenerateSelectedRunnerMacrosCore
    If progressStarted Then CoreCompleteSetupProgress
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    If progressStarted Then CoreFailSetupProgress
    Err.Raise savedErrNumber, callerName, "Could not generate selected Runner Macros." & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

' Purpose: Generates all configured runner macros for active route for caller name, operation caption.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   callerName (Input; String): Procedure/caller name recorded for diagnostics and ownership decisions so nested operations can be attributed correctly.
'   operationCaption (Input; String): Operation caption text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub GenerateAllConfiguredRunnerMacrosForActiveRoute(ByVal callerName As String, ByVal operationCaption As String)
    On Error GoTo ErrHandler
    Dim progressStarted As Boolean
    If Len(callerName) = 0 Then callerName = "GenerateAllConfiguredRunnerMacrosForActiveRoute"
    If Len(operationCaption) = 0 Then operationCaption = "Generate All Configured Runner Macros"
    progressStarted = CoreBeginOperationProgressIfNeeded(operationCaption, 2)
    CoreUpdateSetupProgress "Check default Navigation Pane group"
    RequireDefaultCodeDbGeneratedMacrosNavigationPaneGroupForRoleContext callerName, False
    CoreUpdateSetupProgress "Generate all configured Runner Macros"
    GenerateAllConfiguredRunnerMacrosCore
    If progressStarted Then CoreCompleteSetupProgress
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    If progressStarted Then CoreFailSetupProgress
    Err.Raise savedErrNumber, callerName, "Could not generate all configured Runner Macros." & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

' Purpose: Generates prompted runner macros for active route for caller name, operation caption.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   callerName (Input; String): Procedure/caller name recorded for diagnostics and ownership decisions so nested operations can be attributed correctly.
'   operationCaption (Input; String): Operation caption text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   completionMessageTitle (Input; String): Completion message title text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub GeneratePromptedRunnerMacrosForActiveRoute(ByVal callerName As String, ByVal operationCaption As String, ByVal completionMessageTitle As String)
    On Error GoTo ErrHandler
    Dim progressStarted As Boolean
    If Len(callerName) = 0 Then callerName = "GeneratePromptedRunnerMacrosForActiveRoute"
    If Len(operationCaption) = 0 Then operationCaption = "Generate Prompted Runner Macros"
    If Len(completionMessageTitle) = 0 Then completionMessageTitle = "Generate Prompted Runner Macros"
    progressStarted = CoreBeginOperationProgressIfNeeded(operationCaption, 2)
    CoreUpdateSetupProgress "Check default Navigation Pane group"
    RequireDefaultCodeDbGeneratedMacrosNavigationPaneGroupForRoleContext callerName, False
    CoreUpdateSetupProgress "Prompt for and generate configured macros"
    GeneratePromptedRunnerMacrosCore completionMessageTitle, callerName
    If SystemCommandWasCancelled() Then
        If progressStarted Then EndSetupProgress True
        Exit Sub
    End If
    If progressStarted Then CoreCompleteSetupProgress
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    If progressStarted Then CoreFailSetupProgress
    Err.Raise savedErrNumber, callerName, "Could not generate prompted Runner Macros." & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Function ConfiguredMacroGenerationSqlFilter() As String
    ConfiguredMacroGenerationSqlFilter = "AND (Nz(" & SQL_FIELD_CONFIGURATION_STATUS & ",'')='' OR " & SQL_FIELD_CONFIGURATION_STATUS & "=" & SharedDao_SqlQuote(CONFIGURATION_STATUS_ACTIVE) & ") " & "AND [" & FIELD_CONFIG_OWNER_DATABASE_ROLE & "] In (" & SharedDao_SqlQuote(CONFIG_OWNER_DATABASE_ROLE_HOST_DB) & "," & SharedDao_SqlQuote(CONFIG_OWNER_DATABASE_ROLE_CODE_DB) & ") "
End Function

' Purpose: Validates the current configured Runner Macro targets against the required CodeDb context and Navigation Pane prerequisites, reporting progress and surfacing validation failures.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub ValidateRunnerMacroTargets()
    On Error GoTo ErrHandler
    Dim progressStarted As Boolean
    progressStarted = CoreBeginOperationProgressIfNeeded("Validate Runner Macro Targets", 2)
    CoreUpdateSetupProgress "Check default Navigation Pane group"
    RequireDefaultCodeDbGeneratedMacrosNavigationPaneGroupForRoleContext "ValidateRunnerMacroTargets", False
    CoreUpdateSetupProgress "Validate Runner Macro targets"
    ValidateRunnerMacroTargetsCore
    If progressStarted Then CoreCompleteSetupProgress
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    If progressStarted Then CoreFailSetupProgress
    Err.Raise savedErrNumber, "ValidateRunnerMacroTargets", "Could not validate Runner Macro targets." & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

' Purpose: Checks whether config macro key exists in the database, collection, or metadata scope supplied to this function.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   MacroKey (Input; String): Canonical key identifying the macro definition, setting, or managed item to use.
' Returns: Boolean: True when config macro key exists in the inspected context; False when it does not.
Public Function ConfigMacroKeyExists(ByVal MacroKey As String) As Boolean
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    If Len(Trim$(MacroKey)) = 0 Then Exit Function
    Set rs = ADCDb().OpenRecordset("SELECT TOP 1 " & SQL_FIELD_ID & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "] WHERE " & SQL_FIELD_MACRO_KEY & "=" & SharedDao_SqlQuote(MacroKey), dbOpenSnapshot)
    ConfigMacroKeyExists = (rs.EOF = False)
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
    Err.Raise savedErrNumber, "ConfigMacroKeyExists", "Could not determine whether a configured MacroKey already exists." & vbCrLf & "MacroKey: " & MacroKey & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Sub ValidateRunnerMacroTargetsCore()
    On Error GoTo ErrHandler
    Dim configuredMissingMacroEntries As Collection
    Dim missingEntries As Collection
    Dim unresolvedEntries As Collection
    Dim ambiguousEntries As Collection
    Dim vbeAccessErrorEntries As Collection
    Dim validEntries As Collection
    Dim nonRunCodeEntries As Collection
    Dim exportErrorEntries As Collection
    Dim reportText As String
    Dim reportPath As String
    Dim response As VbMsgBoxResult
    Dim messageText As String
    RequireConfigTableForUtilityWorkflow "ValidateRunnerMacroTargetsCore"
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: configuredMissingMacroEntries, missingEntries, unresolvedEntries, ambiguousEntries, vbeAccessErrorEntries, validEntries, nonRunCodeEntries, exportErrorEntries holds
    ' planned/recovery work or independently observed validation evidence for ValidateRunnerMacroTargetsCore; keep it distinct from the source authority.
    Set configuredMissingMacroEntries = New Collection
    Set missingEntries = New Collection
    Set unresolvedEntries = New Collection
    Set ambiguousEntries = New Collection
    Set vbeAccessErrorEntries = New Collection
    Set validEntries = New Collection
    Set nonRunCodeEntries = New Collection
    Set exportErrorEntries = New Collection
    BuildRunnerMacroTargetValidationCollections configuredMissingMacroEntries, missingEntries, unresolvedEntries, ambiguousEntries, vbeAccessErrorEntries, validEntries, nonRunCodeEntries, exportErrorEntries
    reportText = RunnerMacroTargetValidationReport(configuredMissingMacroEntries, missingEntries, unresolvedEntries, ambiguousEntries, vbeAccessErrorEntries, validEntries, nonRunCodeEntries, exportErrorEntries)
    reportPath = RunnerMacroTargetValidationReportPath
    WriteTextFile reportPath, reportText
    messageText = "Runner Macro target validation complete." & vbCrLf & vbCrLf & "Configured but missing macro object: " & CStr(configuredMissingMacroEntries.count) & vbCrLf & "Missing: " & CStr(missingEntries.count) & vbCrLf & "Unresolved: " & CStr(unresolvedEntries.count) & vbCrLf & "Ambiguous: " & CStr(ambiguousEntries.count) & vbCrLf & "VBE Access Error: " & CStr(vbeAccessErrorEntries.count) & vbCrLf & "Valid: " & CStr(validEntries.count) & vbCrLf & "Non-RunCode: " & CStr(nonRunCodeEntries.count) & vbCrLf & "Export Error: " & CStr(exportErrorEntries.count) & vbCrLf & vbCrLf & "Report:" & vbCrLf & reportPath
    If vbeAccessErrorEntries.count > 0 Or exportErrorEntries.count > 0 Then
        LogTraceMessage TRACE_LEVEL_ERROR, "Validation", "ValidateRunnerMacroTargetsCore", "Runner Macro target validation completed with operational error(s).", messageText, "Complete Runner Macro target validation", 0, vbNullString, vbNullString, reportPath, "File", "CodeDb", "Validation"
        CoreShowErrorMessage messageText, vbCritical, "Validate Runner Macro Targets Completed With Errors", "ValidateRunnerMacroTargetsCore", "Validation"
        response = CoreShowMessageResult("Would you like to open the validation report now?" & vbCrLf & vbCrLf & reportPath, vbQuestion Or vbYesNo Or vbDefaultButton2, "Open Validation Report", "ValidateRunnerMacroTargetsCore", "Validation")
    Else
        response = CoreShowMessageResult(messageText & vbCrLf & vbCrLf & "Would you like to open the report now?", vbInformation Or vbYesNo, "Validate Runner Macro Targets", "ValidateRunnerMacroTargetsCore", "Validation")
    End If
    If response = vbYes Then OpenTextReport reportPath
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ValidateRunnerMacroTargets", "Could not validate Runner Macro targets." & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub AddConfiguredMacroObjectValidationEntries(ByVal MacroTitle As String, ByRef cfg As TMacroRecord, ByRef missingEntries As Collection, ByRef unresolvedEntries As Collection, ByRef ambiguousEntries As Collection, ByRef vbeAccessErrorEntries As Collection, ByRef validEntries As Collection, ByRef nonRunCodeEntries As Collection, ByRef exportErrorEntries As Collection)
    On Error GoTo ErrHandler
    Dim macroText As String
    Dim exportError As String
    Dim runCodeTargets As Collection
    Dim openFormTargets As Collection
    Dim runCodeExpression As Variant
    Dim formName As Variant
    Dim statusText As String
    Dim resolvedProcedureName As String
    Dim detailText As String
    Dim entryText As String
    If TryExportMacroText(MacroTitle, macroText, exportError) = False Then
        exportErrorEntries.Add MacroValidationReportEntry(MacroTitle, "EXPORT ERROR", "", "", exportError)
        Exit Sub
    End If
    Set runCodeTargets = RunCodeTargetsFromMacroText(macroText)
    Set openFormTargets = OpenFormTargetsFromMacroText(macroText)
    If runCodeTargets.count = 0 And openFormTargets.count = 0 Then
        nonRunCodeEntries.Add MacroValidationReportEntry(MacroTitle, "NON-RUNCODE", "", "", "No RunCode or OpenForm action was found in the exported macro text for this Active=True configured macro object." & vbCrLf & "MacroKey: " & cfg.MacroKey & vbCrLf & "Configured action: " & cfg.MacroAction)
    End If
    For Each runCodeExpression In runCodeTargets
        PumpSetupResponsiveness
        resolvedProcedureName = vbNullString
        detailText = vbNullString
        statusText = RunCodeTargetValidationStatus(CStr(runCodeExpression), resolvedProcedureName, detailText)
        entryText = MacroValidationReportEntry(MacroTitle, statusText, CStr(runCodeExpression), resolvedProcedureName, "MacroKey: " & cfg.MacroKey & vbCrLf & detailText)
        AddMacroValidationEntryByStatus statusText, entryText, missingEntries, unresolvedEntries, ambiguousEntries, vbeAccessErrorEntries, validEntries, nonRunCodeEntries, exportErrorEntries
        UpdateValidationStatus cfg.ID, ValidationResultFromTargetStatus(statusText), detailText
    Next runCodeExpression
    For Each formName In openFormTargets
        PumpSetupResponsiveness
        resolvedProcedureName = CStr(formName)
        If AccessObjectExistsInCurrentDatabase(acForm, CStr(formName)) Then
            statusText = "VALID"
            detailText = "OpenForm target exists in HostDb AllForms collection."
        Else
            statusText = "MISSING"
            detailText = "Missing OpenForm target: " & CStr(formName)
        End If
        entryText = MacroValidationReportEntry(MacroTitle, statusText, "OpenForm: " & CStr(formName), resolvedProcedureName, "MacroKey: " & cfg.MacroKey & vbCrLf & detailText)
        AddMacroValidationEntryByStatus statusText, entryText, missingEntries, unresolvedEntries, ambiguousEntries, vbeAccessErrorEntries, validEntries, nonRunCodeEntries, exportErrorEntries
        UpdateValidationStatus cfg.ID, ValidationResultFromTargetStatus(statusText), detailText
    Next formName
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    exportErrorEntries.Add MacroValidationReportEntry(MacroTitle, "EXPORT ERROR", "", "", "Validation failed while inspecting this configured macro object." & vbCrLf & "MacroKey: " & cfg.MacroKey & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription)
End Sub

Private Sub AddConfiguredButMissingMacroObjectEntries(ByRef configuredMissingMacroEntries As Collection)
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim cfg As TMacroRecord
    Dim macroTitleText As String
    Dim processedCount As Long
    If CodeDbTableExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")) = False Then Exit Sub
    Set rs = ADCDb().OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "] WHERE " & SQL_FIELD_ACTIVE & "<>0 ORDER BY Nz(" & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER & ",0), " & SQL_FIELD_MACRO_TITLE & ", " & SQL_FIELD_MACRO_KEY, dbOpenSnapshot)
    Do While Not rs.EOF
        PumpSetupResponsiveness
        processedCount = processedCount + 1
        PulseProgressEvery processedCount, 0, "Configured macro objects checked"
        MacroRecordFromRecordset rs, cfg
        macroTitleText = cfg.MacroTitle
        If Len(Trim$(macroTitleText)) = 0 Then
            configuredMissingMacroEntries.Add MacroValidationReportEntry(SharedText_Nz(cfg.MacroKey), "CONFIGURED BUT MISSING MACRO OBJECT", "(configured row)", "(blank MacroTitle)", "Active config row has a blank MacroTitle, so no macro object can be matched.")
        ElseIf AccessObjectExistsInCurrentDatabase(acMacro, macroTitleText) = False Then
            configuredMissingMacroEntries.Add MacroValidationReportEntry(macroTitleText, "CONFIGURED BUT MISSING MACRO OBJECT", "(configured row)", EffectiveConfiguredTargetSummary(cfg), "Active config row exists but no same-named Access macro object exists in HostDb AllMacros collection." & vbCrLf & "MacroKey: " & cfg.MacroKey & vbCrLf & "MacroAction: " & cfg.MacroAction)
        End If
        rs.MoveNext
    Loop
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
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
    Err.Raise savedErrNumber, "AddConfiguredButMissingMacroObjectEntries", "Could not scan active config rows for missing macro objects." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

' Purpose: Attempts to perform export macro text for macro title.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   MacroTitle (Input; String): Title text used as the macro title shown, stored, compared, or resolved by this operation.
'   macroText (Output; String): Receives the resolved macro text produced from this procedure's validated inputs and current ADC state.
'   errorText (Output; String): Receives the resolved error text produced from this procedure's validated inputs and current ADC state.
' Returns: Boolean: True when the attempt to export macro text completes successfully; False for the routine's defined normal unsuccessful/not-applicable outcome.
Public Function TryExportMacroText(ByVal MacroTitle As String, ByRef macroText As String, ByRef errorText As String) As Boolean
    On Error GoTo ErrHandler
    Dim tempPath As String
    Dim fileNum As Integer
    tempPath = Environ$("TEMP") & "\ExportMacro_" & SharedName_SafeFileName(MacroTitle, ADCAT_OBJECT_TYPE_MACRO, 120) & "_" & Format$(Now, "yyyymmdd_hhnnss") & ".txt"
    Application.SaveAsText acMacro, MacroTitle, tempPath
    fileNum = FreeFile
    Open tempPath For Binary Access Read As #fileNum
    If LOF(fileNum) > 0 Then
        macroText = Space$(LOF(fileNum))
        Get #fileNum, , macroText
    Else
        macroText = vbNullString
    End If
    Close #fileNum
    fileNum = 0
    TryExportMacroText = True
ExitHere:
    On Error Resume Next
    If fileNum <> 0 Then Close #fileNum
    If Len(tempPath) > 0 Then
        If Len(Dir$(tempPath)) > 0 Then Kill tempPath
    End If
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    errorText = "Could not export macro text." & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
    TryExportMacroText = False
    Resume ExitHere
End Function



Private Function ExpectedProcedureKindForValidation(ByVal TargetProcedureKind As String) As String
    Select Case LCase$(Trim$(TargetProcedureKind))
        Case LCase$(TARGET_PROCEDURE_KIND_FUNCTION)
            ExpectedProcedureKindForValidation = TARGET_PROCEDURE_KIND_FUNCTION
        Case LCase$(TARGET_PROCEDURE_KIND_SUB)
            ExpectedProcedureKindForValidation = TARGET_PROCEDURE_KIND_SUB
        Case Else
            ExpectedProcedureKindForValidation = vbNullString
    End Select
End Function

Private Function ValidateConfiguredTargetArgumentCompatibility(ByRef cfg As TMacroRecord, ByRef detailText As String) As String
    On Error GoTo ErrHandler
    Dim procedureKind As String
    Dim procedureLocation As String
    Dim argumentSignature As String
    Dim RequiredArgumentCount As Long
    Dim OptionalArgumentCount As Long
    Dim hasParamArray As Boolean
    Dim ExpectedReturnType As String
    Dim actualArgumentCount As Long
    Dim maximumArgumentCount As Long
    Dim roleText As String
    roleText = cfg.ConfigOwnerDatabaseRole
    ClearVBEProcedureSearchError
    If PublicCallableProcedureInfoInStandardModule(cfg.TargetProcedureName, roleText, procedureKind, procedureLocation, argumentSignature, RequiredArgumentCount, OptionalArgumentCount, hasParamArray, ExpectedReturnType) = False Then
        If VBEProcedureSearchAmbiguityOccurred() Then
            detailText = "Live argument validation found duplicate callable declarations for the target procedure." & vbCrLf & VBEProcedureSearchAmbiguityText()
            ValidateConfiguredTargetArgumentCompatibility = "AMBIGUOUS"
        ElseIf VBEProcedureSearchErrorOccurred() Then
            detailText = "Live argument validation could not inspect the target declaration." & vbCrLf & VBEProcedureSearchErrorText()
            ValidateConfiguredTargetArgumentCompatibility = "VBE ACCESS ERROR"
        Else
            detailText = "Live argument validation could not re-read the target declaration after existence validation." & vbCrLf & "TargetProcedureName: " & cfg.TargetProcedureName & vbCrLf & "ConfigOwnerDatabaseRole: " & cfg.ConfigOwnerDatabaseRole
            ValidateConfiguredTargetArgumentCompatibility = "MISSING"
        End If
        Exit Function
    End If
    actualArgumentCount = ConfiguredTargetArgumentCount(cfg.TargetArguments)
    maximumArgumentCount = RequiredArgumentCount + OptionalArgumentCount
    If actualArgumentCount < RequiredArgumentCount Then
        detailText = "Live argument validation failed. Too few TargetArguments are configured." & vbCrLf & LiveArgumentValidationSummary(actualArgumentCount, RequiredArgumentCount, OptionalArgumentCount, hasParamArray, argumentSignature, procedureLocation)
        ValidateConfiguredTargetArgumentCompatibility = "UNRESOLVED"
        Exit Function
    End If
    If hasParamArray = False Then
        If actualArgumentCount > maximumArgumentCount Then
            detailText = "Live argument validation failed. Too many TargetArguments are configured." & vbCrLf & LiveArgumentValidationSummary(actualArgumentCount, RequiredArgumentCount, OptionalArgumentCount, hasParamArray, argumentSignature, procedureLocation)
            ValidateConfiguredTargetArgumentCompatibility = "UNRESOLVED"
            Exit Function
        End If
    End If
    detailText = "Live argument validation passed." & vbCrLf & LiveArgumentValidationSummary(actualArgumentCount, RequiredArgumentCount, OptionalArgumentCount, hasParamArray, argumentSignature, procedureLocation)
    ValidateConfiguredTargetArgumentCompatibility = "VALID"
    Exit Function
ErrHandler:
    detailText = "Live argument validation failed with an error." & vbCrLf & "TargetProcedureName: " & cfg.TargetProcedureName & vbCrLf & "Err.Description: " & Err.Description
    ValidateConfiguredTargetArgumentCompatibility = "UNRESOLVED"
End Function


Private Function LiveArgumentValidationSummary(ByVal actualArgumentCount As Long, ByVal RequiredArgumentCount As Long, ByVal OptionalArgumentCount As Long, ByVal hasParamArray As Boolean, ByVal argumentSignature As String, ByVal procedureLocation As String) As String
    LiveArgumentValidationSummary = "Configured TargetArguments count: " & CStr(actualArgumentCount) & vbCrLf & "Live required argument count: " & CStr(RequiredArgumentCount) & vbCrLf & "Live optional argument count: " & CStr(OptionalArgumentCount) & vbCrLf & "Live ParamArray: " & IIf(hasParamArray, "Yes", "No") & vbCrLf & "Live declaration arguments: " & IIf(Len(Trim$(argumentSignature)) = 0, "(none)", argumentSignature) & vbCrLf & "Location: " & procedureLocation
End Function


' Purpose: Starts the explicit lifetime for MACRO_RUNTIME's authoritative reusable callable-procedure projection.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   ConfigOwnerDatabaseRole (Input; String): Identifies the database role whose configuration is authoritative for this operation, such as the CodeDb or HostDb role defined by ADC configuration.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub BeginPublicCallableProcedureLookupScope(ByVal ConfigOwnerDatabaseRole As String)
    'Starts the explicit lifetime for MACRO_RUNTIME's authoritative reusable callable-procedure projection.
    Dim roleText As String
    roleText = NormalizedConfigOwnerDatabaseRole(ConfigOwnerDatabaseRole)
    If Len(roleText) = 0 Then Err.Raise vbObjectError + 5460, "BeginPublicCallableProcedureLookupScope", "ConfigOwnerDatabaseRole is blank or unsupported."
    ResetPublicCallableProcedureLookupScope
    mPublicCallableProcedureLookupScopeActive = True
    mPublicCallableProcedureLookupRole = roleText
End Sub

' Purpose: Ends the explicit callable-procedure lookup lifetime and discards all derived VBE state.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub EndPublicCallableProcedureLookupScope()
    ' Wrapper purpose: Public lifecycle endpoint that ends the callable-procedure lookup scope and discards its derived VBE state.
    ' Why retained: Retained so external owners can close CONFIG-owned private lookup state without calling MACRO_RUNTIME's internal reset implementation directly.
    'Ends the explicit callable-procedure lookup lifetime and discards all derived VBE state.
    ResetPublicCallableProcedureLookupScope
End Sub

' Purpose: Returns public callable procedure lookup scope stats.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   distinctProcedureNamesIndexed (Output; Long): Receives the resolved distinct procedure names indexed produced from this procedure's validated inputs and current ADC state.
'   vbComponentsEnumerated (Output; Long): Receives the resolved vb components enumerated produced from this procedure's validated inputs and current ADC state.
'   completeProjectEnumerationPasses (Output; Long): Receives the resolved complete project enumeration passes produced from this procedure's validated inputs and current ADC state.
'   fallbackSourceTextScans (Output; Long): Receives the resolved fallback source text scans produced from this procedure's validated inputs and current ADC state.
' Returns: None (Sub). Results are returned through the documented ByRef parameters and/or the routine's database, UI, or module-state side effects.
Public Sub GetPublicCallableProcedureLookupScopeStats(ByRef distinctProcedureNamesIndexed As Long, ByRef vbComponentsEnumerated As Long, ByRef completeProjectEnumerationPasses As Long, ByRef fallbackSourceTextScans As Long)
    'Reports coarse operation metrics without exposing or duplicating the authoritative lookup itself.
    distinctProcedureNamesIndexed = mPublicCallableProcedureLookupDistinctNameCount
    vbComponentsEnumerated = mPublicCallableProcedureLookupVBComponentsEnumerated
    completeProjectEnumerationPasses = mPublicCallableProcedureLookupProjectEnumerationPasses
    fallbackSourceTextScans = mPublicCallableProcedureLookupFallbackSourceTextScans
End Sub

Private Sub ResetPublicCallableProcedureLookupScope()
    Set mPublicCallableProcedureLookup = Nothing
    mPublicCallableProcedureLookupScopeActive = False
    mPublicCallableProcedureLookupRole = vbNullString
    mPublicCallableProcedureLookupTreatOptionPrivateAsNotCallable = False
    mPublicCallableProcedureLookupBuilt = False
    mPublicCallableProcedureLookupBuildErrorText = vbNullString
    mPublicCallableProcedureLookupDistinctNameCount = 0
    mPublicCallableProcedureLookupVBComponentsEnumerated = 0
    mPublicCallableProcedureLookupProjectEnumerationPasses = 0
    mPublicCallableProcedureLookupFallbackSourceTextScans = 0
End Sub

Private Sub EnsurePublicCallableProcedureLookup(ByVal vbProject As Object, ByVal treatOptionPrivateAsNotCallable As Boolean)
    On Error GoTo ErrHandler
    Dim vbComponent As Object
    Dim moduleErrorText As String
    Dim seenProcedureLocations As Object
    If mPublicCallableProcedureLookupBuilt Then
        If mPublicCallableProcedureLookupTreatOptionPrivateAsNotCallable <> treatOptionPrivateAsNotCallable Then Err.Raise vbObjectError + 5461, "EnsurePublicCallableProcedureLookup", "The active callable-procedure lookup scope cannot change Option Private callability semantics after construction."
        Exit Sub
    End If
    If vbProject Is Nothing Then Err.Raise vbObjectError + 5462, "EnsurePublicCallableProcedureLookup", "Could not resolve a VBProject for the active callable-procedure lookup scope."
    Set mPublicCallableProcedureLookup = CreateObject("Scripting.Dictionary")
    mPublicCallableProcedureLookup.CompareMode = vbTextCompare
    Set seenProcedureLocations = CreateObject("Scripting.Dictionary")
    seenProcedureLocations.CompareMode = vbTextCompare
    mPublicCallableProcedureLookupTreatOptionPrivateAsNotCallable = treatOptionPrivateAsNotCallable
    mPublicCallableProcedureLookupProjectEnumerationPasses = 1
    For Each vbComponent In vbProject.VBComponents
        PumpSetupResponsiveness
        mPublicCallableProcedureLookupVBComponentsEnumerated = mPublicCallableProcedureLookupVBComponentsEnumerated + 1
        If vbComponent.Type = VBEXT_CT_STD_MODULE Then
            moduleErrorText = vbNullString
            AddPublicCallableProcedureLookupFromStandardModule vbProject, vbComponent, treatOptionPrivateAsNotCallable, seenProcedureLocations, moduleErrorText
            If Len(Trim$(moduleErrorText)) > 0 Then
                If Len(mPublicCallableProcedureLookupBuildErrorText) > 0 Then mPublicCallableProcedureLookupBuildErrorText = mPublicCallableProcedureLookupBuildErrorText & vbCrLf
                mPublicCallableProcedureLookupBuildErrorText = mPublicCallableProcedureLookupBuildErrorText & moduleErrorText
            End If
        End If
    Next vbComponent
    mPublicCallableProcedureLookupDistinctNameCount = mPublicCallableProcedureLookup.count
    mPublicCallableProcedureLookupBuilt = True
    Exit Sub
ErrHandler:
    If Len(mPublicCallableProcedureLookupBuildErrorText) > 0 Then mPublicCallableProcedureLookupBuildErrorText = mPublicCallableProcedureLookupBuildErrorText & vbCrLf
    mPublicCallableProcedureLookupBuildErrorText = mPublicCallableProcedureLookupBuildErrorText & "Could not construct the operation-scoped callable-procedure lookup. Project: " & SafeObjectText(vbProject, "Name") & ". Err.Description: " & Err.Description
    mPublicCallableProcedureLookupBuilt = True
End Sub

Private Sub AddPublicCallableProcedureLookupFromStandardModule(ByVal vbProject As Object, ByVal vbComponent As Object, ByVal treatOptionPrivateAsNotCallable As Boolean, ByVal seenProcedureLocations As Object, ByRef moduleErrorText As String)
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
    Set codeModule = vbComponent.codeModule
    If treatOptionPrivateAsNotCallable Then
        If ModuleHasOptionPrivateModule(codeModule) Then Exit Sub
    End If
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
            procedureKind = vbNullString
            visibilityText = vbNullString
            argumentSignature = vbNullString
            RequiredArgumentCount = 0
            OptionalArgumentCount = 0
            ExpectedReturnType = vbNullString
            If CandidateProcedureDeclarationInfo(declarationText, procName, procedureKind, visibilityText, argumentSignature, RequiredArgumentCount, OptionalArgumentCount, ExpectedReturnType) Then
                AddPublicCallableProcedureLookupEntry vbProject, vbComponent, procName, startLine, False, procedureKind, argumentSignature, RequiredArgumentCount, OptionalArgumentCount, ProcedureSignatureHasParamArray(argumentSignature), ExpectedReturnType, seenProcedureLocations
            End If
            If countLines > 0 Then
                lineIndex = startLine + countLines
            Else
                lineIndex = lineIndex + 1
            End If
        End If
    Loop
    'One source-text fallback scan per standard module is folded into the same VBComponents pass. It only adds declarations the VBE procedure API did not expose, preserving the old fallback semantics without N
    'project rescans.
    mPublicCallableProcedureLookupFallbackSourceTextScans = mPublicCallableProcedureLookupFallbackSourceTextScans + 1
    For lineIndex = 1 To codeModule.CountOfLines
        PumpSetupResponsiveness
        declarationText = Trim$(codeModule.lines(lineIndex, 1))
        procName = PublicCallableProcedureNameFromDeclaration(declarationText)
        If Len(procName) > 0 Then
            declarationText = ProcedureDeclarationTextFromKnownLine(codeModule, lineIndex)
            procedureKind = vbNullString
            visibilityText = vbNullString
            argumentSignature = vbNullString
            RequiredArgumentCount = 0
            OptionalArgumentCount = 0
            ExpectedReturnType = vbNullString
            If CandidateProcedureDeclarationInfo(declarationText, procName, procedureKind, visibilityText, argumentSignature, RequiredArgumentCount, OptionalArgumentCount, ExpectedReturnType) Then
                AddPublicCallableProcedureLookupEntry vbProject, vbComponent, procName, lineIndex, True, procedureKind, argumentSignature, RequiredArgumentCount, OptionalArgumentCount, ProcedureSignatureHasParamArray(argumentSignature), ExpectedReturnType, seenProcedureLocations
            End If
        End If
    Next lineIndex
    Exit Sub
ErrHandler:
    moduleErrorText = "Could not inspect standard module while constructing the callable-procedure lookup. Project: " & SafeObjectText(vbProject, "Name") & ". Module: " & SafeObjectText(vbComponent, "Name") & ". Err.Description: " & Err.Description
End Sub


Private Sub AddPublicCallableProcedureLookupEntry(ByVal vbProject As Object, ByVal vbComponent As Object, ByVal procName As String, ByVal startLine As Long, ByVal fromTextScan As Boolean, ByVal procedureKind As String, ByVal argumentSignature As String, ByVal RequiredArgumentCount As Long, ByVal OptionalArgumentCount As Long, ByVal hasParamArray As Boolean, ByVal ExpectedReturnType As String, ByVal seenProcedureLocations As Object)
    Dim procedureKey As String
    Dim locationKey As String
    Dim procedureLocation As String
    Dim matchBuckets As Collection
    Dim primaryMatches As Collection
    Dim fallbackMatches As Collection
    procedureKey = LCase$(Trim$(procName))
    If Len(procedureKey) = 0 Then Exit Sub
    locationKey = LCase$(SafeObjectText(vbComponent, "Name")) & "|" & procedureKey
    If seenProcedureLocations.Exists(locationKey) Then Exit Sub
    seenProcedureLocations.Add locationKey, True
    procedureLocation = SafeObjectText(vbProject, "Name") & "." & SafeObjectText(vbComponent, "Name") & ":" & CStr(startLine)
    If fromTextScan Then procedureLocation = procedureLocation & " (text scan)"
    If mPublicCallableProcedureLookup.Exists(procedureKey) Then
        Set matchBuckets = mPublicCallableProcedureLookup(procedureKey)
    Else
        Set matchBuckets = New Collection
        Set primaryMatches = New Collection
        Set fallbackMatches = New Collection
        matchBuckets.Add primaryMatches
        matchBuckets.Add fallbackMatches
        mPublicCallableProcedureLookup.Add procedureKey, matchBuckets
    End If
    If primaryMatches Is Nothing Then Set primaryMatches = matchBuckets(1)
    If fallbackMatches Is Nothing Then Set fallbackMatches = matchBuckets(2)
    If fromTextScan Then
        fallbackMatches.Add Array(procedureKind, procedureLocation, argumentSignature, RequiredArgumentCount, OptionalArgumentCount, hasParamArray, ExpectedReturnType)
    Else
        primaryMatches.Add Array(procedureKind, procedureLocation, argumentSignature, RequiredArgumentCount, OptionalArgumentCount, hasParamArray, ExpectedReturnType)
    End If
End Sub

Private Function PublicCallableProcedureInfoFromLookup(ByVal vbProject As Object, ByVal ProcedureName As String, ByRef procedureKind As String, ByRef procedureLocation As String, ByRef argumentSignature As String, ByRef RequiredArgumentCount As Long, ByRef OptionalArgumentCount As Long, ByRef hasParamArray As Boolean, ByRef ExpectedReturnType As String) As Boolean
    Dim procedureKey As String
    Dim matchBuckets As Collection
    Dim primaryMatches As Collection
    Dim fallbackMatches As Collection
    Dim matches As Collection
    Dim matchData As Variant
    Dim item As Variant
    Dim matchText As String
    If Len(Trim$(mPublicCallableProcedureLookupBuildErrorText)) > 0 Then AppendVBEProcedureSearchError mPublicCallableProcedureLookupBuildErrorText
    If mPublicCallableProcedureLookup Is Nothing Then Exit Function
    procedureKey = LCase$(Trim$(ProcedureName))
    If mPublicCallableProcedureLookup.Exists(procedureKey) = False Then Exit Function
    Set matchBuckets = mPublicCallableProcedureLookup(procedureKey)
    Set primaryMatches = matchBuckets(1)
    Set fallbackMatches = matchBuckets(2)
    If primaryMatches.count > 0 Then
        Set matches = primaryMatches
    Else
        Set matches = fallbackMatches
    End If
    If matches.count = 1 Then
        matchData = matches(1)
        procedureKind = CStr(matchData(0))
        procedureLocation = CStr(matchData(1))
        argumentSignature = CStr(matchData(2))
        RequiredArgumentCount = CLng(matchData(3))
        OptionalArgumentCount = CLng(matchData(4))
        hasParamArray = CBool(matchData(5))
        ExpectedReturnType = CStr(matchData(6))
        PublicCallableProcedureInfoFromLookup = True
    ElseIf matches.count > 1 Then
        For Each item In matches
            matchText = matchText & " - " & CStr(item(0)) & " at " & CStr(item(1)) & IIf(Len(Trim$(CStr(item(2)))) = 0, " with no arguments", " with arguments (" & CStr(item(2)) & ")") & vbCrLf
        Next item
        AppendVBEProcedureSearchAmbiguity "ProcedureName: " & ProcedureName & vbCrLf & "VBProject: " & SafeObjectText(vbProject, "Name") & vbCrLf & "Matches found in the same VBProject: " & CStr(matches.count) & vbCrLf & matchText & "Application.Run cannot safely disambiguate same-named public procedures. Rename duplicates or hide non-target procedures."
    End If
End Function

' Purpose: Performs the public clear vbe procedure search error operation using the routine's existing ADC workflow and side effects.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub ClearVBEProcedureSearchError()
    'Intended use: Runs Clear VBE Procedure Search Error for direct, generated-macro, form-button, or required cross-module callers.
    mLastVBEProcedureSearchError = vbNullString
    mLastVBEProcedureSearchAmbiguity = vbNullString
End Sub

Private Sub AppendVBEProcedureSearchError(ByVal detailText As String)
    If Len(Trim$(detailText)) = 0 Then Exit Sub
    If Len(Trim$(mLastVBEProcedureSearchError)) = 0 Then
        mLastVBEProcedureSearchError = detailText
    Else
        mLastVBEProcedureSearchError = mLastVBEProcedureSearchError & vbCrLf & detailText
    End If
End Sub

' Purpose: Performs VBE procedure search error occurred by coordinating `End` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: Boolean: True when VBE procedure search error occurred in the operation/state tracked by this routine; False when it did not.
Public Function VBEProcedureSearchErrorOccurred() As Boolean
    'Intended use: Returns the VBE Procedure Search Error Occurred result for direct, query, form, or required cross-module callers.
    VBEProcedureSearchErrorOccurred = (Len(Trim$(mLastVBEProcedureSearchError)) > 0)
End Function

' Purpose: Resolves VBE procedure search error text from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function VBEProcedureSearchErrorText() As String
    'Intended use: Returns the VBE Procedure Search Error Text result for direct, query, form, or required cross-module callers.
    If Len(Trim$(mLastVBEProcedureSearchError)) = 0 Then
        VBEProcedureSearchErrorText = "VBE access failed while inspecting callable procedures."
    Else
        VBEProcedureSearchErrorText = "VBE ACCESS ERROR while inspecting callable procedures:" & vbCrLf & mLastVBEProcedureSearchError
    End If
End Function

Private Sub AppendVBEProcedureSearchAmbiguity(ByVal detailText As String)
    If Len(Trim$(detailText)) = 0 Then Exit Sub
    If Len(Trim$(mLastVBEProcedureSearchAmbiguity)) = 0 Then
        mLastVBEProcedureSearchAmbiguity = detailText
    Else
        mLastVBEProcedureSearchAmbiguity = mLastVBEProcedureSearchAmbiguity & vbCrLf & detailText
    End If
End Sub

' Purpose: Performs VBE procedure search ambiguity occurred by coordinating `End` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: Boolean: True when VBE procedure search ambiguity occurred in the operation/state tracked by this routine; False when it did not.
Public Function VBEProcedureSearchAmbiguityOccurred() As Boolean
    'Intended use: Returns the VBE Procedure Search Ambiguity Occurred result for direct, query, form, or required cross-module callers.
    VBEProcedureSearchAmbiguityOccurred = (Len(Trim$(mLastVBEProcedureSearchAmbiguity)) > 0)
End Function

' Purpose: Resolves VBE procedure search ambiguity text from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function VBEProcedureSearchAmbiguityText() As String
    'Intended use: Returns the VBE Procedure Search Ambiguity Text result for direct, query, form, or required cross-module callers.
    If Len(Trim$(mLastVBEProcedureSearchAmbiguity)) = 0 Then
        VBEProcedureSearchAmbiguityText = "Duplicate callable procedure declarations were found."
    Else
        VBEProcedureSearchAmbiguityText = "AMBIGUOUS callable procedure declarations:" & vbCrLf & mLastVBEProcedureSearchAmbiguity
    End If
End Function

Private Function QualifiedRunProcedureName(ByRef cfg As TMacroRecord) As String
    On Error GoTo ErrHandler
    Dim vbProject As Object
    Dim roleText As String
    Dim procText As String
    roleText = NormalizedConfigOwnerDatabaseRole(cfg.ConfigOwnerDatabaseRole)
    procText = Trim$(cfg.TargetProcedureName)
    If Len(procText) = 0 Then Exit Function
    Select Case LCase$(roleText)
        Case LCase$(CONFIG_OWNER_DATABASE_ROLE_CODE_DB)
            Set vbProject = VBProjectForConfigOwnerDatabaseRole(roleText)
            If Not vbProject Is Nothing Then
                QualifiedRunProcedureName = SafeObjectText(vbProject, "Name") & "." & procText
            Else
                QualifiedRunProcedureName = procText
            End If
        Case Else
            'HostDb targets intentionally remain unqualified so Access resolves them in the open project.
            QualifiedRunProcedureName = procText
    End Select
    Exit Function
ErrHandler:
    QualifiedRunProcedureName = Trim$(cfg.TargetProcedureName)
End Function


Private Function ProcedureDeclarationTextFromKnownLine(ByVal codeModule As Object, ByVal startLine As Long) As String
    On Error GoTo ErrHandler
    Dim lineIndex As Long
    Dim maxLine As Long
    Dim declarationText As String
    maxLine = startLine + 25
    If maxLine > codeModule.CountOfLines Then maxLine = codeModule.CountOfLines
    declarationText = CodeLineWithoutTrailingComment(codeModule.lines(startLine, 1))
    lineIndex = startLine
    Do While VbaLineContinues(declarationText) And lineIndex < maxLine
        PumpSetupResponsiveness
        declarationText = RemoveVbaLineContinuation(declarationText)
        lineIndex = lineIndex + 1
        declarationText = declarationText & " " & CodeLineWithoutTrailingComment(codeModule.lines(lineIndex, 1))
    Loop
    ProcedureDeclarationTextFromKnownLine = declarationText
    Exit Function
ErrHandler:
    ProcedureDeclarationTextFromKnownLine = vbNullString
End Function


' Purpose: Runs confirmed configured macro target for macro key.
' Public role: Form-code/event-expression public entry point. Access form code or an event property can invoke this routine by name; its Public visibility is part of that UI contract.
' Parameters:
'   MacroKey (Input; String): Canonical key identifying the macro definition, setting, or managed item to use.
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
' Returns: Boolean result of run confirmed configured macro target; True and False correspond to the affirmative and negative branches explicitly implemented by this function.
'   path.
Public Function RunConfirmedConfiguredMacroTarget(ByVal MacroKey As String, Optional ByVal showMessage As Boolean = True) As Boolean
    'Intended use: Returns the Run Confirmed Configured Macro Target result for direct, query, form, or required cross-module callers.
    On Error GoTo ErrHandler
    Dim traceRunOwnedByProcedure As Boolean
    If TraceRunActive() = False Then
        StartTraceRun "ConfirmedConfiguredMacroTarget"
        traceRunOwnedByProcedure = True
    End If
    CorePulseProgress "Confirmed configured macro target started: " & MacroKey
    DoEvents
    LogTraceMessage "Info", "GeneratedMacro", "RunConfirmedConfiguredMacroTarget", "Confirmed configured macro target started.", "MacroKey: " & MacroKey, MacroKey, 0, vbNullString, vbNullString, MacroKey, "GeneratedMacro", "CodeDb", "ConfiguredMacroTarget"
    RunConfirmedConfiguredMacroTarget = RunConfirmedConfiguredMacroTargetCore(MacroKey)
    If RunConfirmedConfiguredMacroTarget Then
        CorePulseProgress "Confirmed configured macro target complete: " & MacroKey
        DoEvents
        LogTraceMessage "Info", "GeneratedMacro", "RunConfirmedConfiguredMacroTarget", "Confirmed configured macro target completed.", "MacroKey: " & MacroKey, MacroKey, 0, vbNullString, vbNullString, MacroKey, "GeneratedMacro", "CodeDb", "ConfiguredMacroTarget"
    Else
        CorePulseProgress "Confirmed configured macro target did not run or returned False: " & MacroKey
        DoEvents
        LogTraceMessage "Info", "GeneratedMacro", "RunConfirmedConfiguredMacroTarget", "Confirmed configured macro target did not run or returned False.", "MacroKey: " & MacroKey, MacroKey, 0, vbNullString, vbNullString, MacroKey, "GeneratedMacro", "CodeDb", "ConfiguredMacroTarget"
    End If
    If traceRunOwnedByProcedure Then CompleteTraceRun "Confirmed configured macro target complete: " & MacroKey
    Exit Function
ErrHandler:
    Dim messageText As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    CorePulseProgress "Confirmed configured macro target failed: " & MacroKey
    DoEvents
    messageText = SharedError_FormatDetails("RunConfirmedConfiguredMacroTarget") & vbCrLf & vbCrLf & "MacroKey: " & MacroKey & TableDiagnosticsLocationForMessage()
    LogTraceMessage "Error", "GeneratedMacro", "RunConfirmedConfiguredMacroTarget", "Confirmed configured macro target failed.", messageText, MacroKey, savedErrNumber, savedErrSource, savedErrDescription, MacroKey, "GeneratedMacro", "CodeDb", "ConfiguredMacroTarget"
    If showMessage Then CoreShowErrorMessage messageText, vbCritical, "Confirmed Macro Action", "RunConfirmedConfiguredMacroTarget", "Core"
    If traceRunOwnedByProcedure Then CompleteTraceRun "Confirmed configured macro target failed: " & MacroKey, False
    RunConfirmedConfiguredMacroTarget = False
End Function

Private Function RunConfirmedConfiguredMacroTargetCore(ByVal MacroKey As String) As Boolean
    On Error GoTo ErrHandler
    Dim cfg As TMacroRecord
    Dim response As VbMsgBoxResult
    Dim promptText As String
    If LoadMacroConfigByKey(MacroKey, cfg) = False Then
        Err.Raise vbObjectError + 5320, "RunConfirmedConfiguredMacroTargetCore", "No macro config row was found for MacroKey: " & MacroKey & vbCrLf & "MC_Records table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & vbCrLf & "The macro will not run until an active config row with this MacroKey exists."
    End If
    promptText = "Would you like to run this macro?" & vbCrLf & vbCrLf & "Description: " & EffectiveActionDescription(cfg) & vbCrLf & "Macro: " & cfg.MacroTitle & vbCrLf & "Action: " & cfg.MacroAction & vbCrLf & "ConfigOwnerDatabaseRole: " & cfg.ConfigOwnerDatabaseRole
    If StrComp(cfg.MacroAction, ACTION_RUNCODE, vbTextCompare) = 0 Then
        promptText = promptText & vbCrLf & "Target: " & EffectiveTargetCallText(cfg)
    ElseIf ActionIsAccessObjectAction(cfg.MacroAction) Then
        promptText = promptText & vbCrLf & "Object: " & ConfiguredObjectActionTargetName(cfg)
    End If
    response = CoreShowMessageResult(promptText, vbQuestion Or vbYesNo Or vbDefaultButton2, "Confirm Macro Action")
    If response <> vbYes Then
        RunConfirmedConfiguredMacroTargetCore = False
        Exit Function
    End If
    RunConfirmedConfiguredMacroTargetCore = RunConfiguredMacroTargetCore(MacroKey)
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "RunConfirmedConfiguredMacroTargetCore", "Could not run confirmed configured macro target." & vbCrLf & vbCrLf & "MacroKey: " & MacroKey & vbCrLf & "Err.Description: " & savedErrDescription
End Function

' Purpose: Runs configured macro target for macro key.
' Public role: Form-code/event-expression public entry point. Access form code or an event property can invoke this routine by name; its Public visibility is part of that UI contract.
' Parameters:
'   MacroKey (Input; String): Canonical key identifying the macro definition, setting, or managed item to use.
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
' Returns: Boolean result of run configured macro target; True and False correspond to the affirmative and negative branches explicitly implemented by this function.
Public Function RunConfiguredMacroTarget(ByVal MacroKey As String, Optional ByVal showMessage As Boolean = True) As Boolean
    'Intended use: Returns the Run Configured Macro Target result for direct, query, form, or required cross-module callers.
    On Error GoTo ErrHandler
    Dim traceRunOwnedByProcedure As Boolean
    If TraceRunActive() = False Then
        StartTraceRun "ConfiguredMacroTarget"
        traceRunOwnedByProcedure = True
    End If
    CorePulseProgress "Configured macro target started: " & MacroKey
    DoEvents
    LogTraceMessage "Info", "GeneratedMacro", "RunConfiguredMacroTarget", "Configured macro target started.", "MacroKey: " & MacroKey, MacroKey, 0, vbNullString, vbNullString, MacroKey, "GeneratedMacro", "CodeDb", "ConfiguredMacroTarget"
    RunConfiguredMacroTarget = RunConfiguredMacroTargetCore(MacroKey)
    If RunConfiguredMacroTarget Then
        CorePulseProgress "Configured macro target complete: " & MacroKey
        DoEvents
        LogTraceMessage "Info", "GeneratedMacro", "RunConfiguredMacroTarget", "Configured macro target completed.", "MacroKey: " & MacroKey, MacroKey, 0, vbNullString, vbNullString, MacroKey, "GeneratedMacro", "CodeDb", "ConfiguredMacroTarget"
    Else
        CorePulseProgress "Configured macro target returned False: " & MacroKey
        DoEvents
        LogTraceMessage "Info", "GeneratedMacro", "RunConfiguredMacroTarget", "Configured macro target returned False.", "MacroKey: " & MacroKey, MacroKey, 0, vbNullString, vbNullString, MacroKey, "GeneratedMacro", "CodeDb", "ConfiguredMacroTarget"
    End If
    If traceRunOwnedByProcedure Then CompleteTraceRun "Configured macro target complete: " & MacroKey
    Exit Function
ErrHandler:
    Dim messageText As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    CorePulseProgress "Configured macro target failed: " & MacroKey
    DoEvents
    messageText = SharedError_FormatDetails("RunConfiguredMacroTarget") & vbCrLf & vbCrLf & "MacroKey: " & MacroKey & TableDiagnosticsLocationForMessage()
    LogTraceMessage "Error", "GeneratedMacro", "RunConfiguredMacroTarget", "Configured macro target failed.", messageText, MacroKey, savedErrNumber, savedErrSource, savedErrDescription, MacroKey, "GeneratedMacro", "CodeDb", "ConfiguredMacroTarget"
    If showMessage Then CoreShowErrorMessage messageText, vbCritical, "Run Configured Macro Target", "RunConfiguredMacroTarget", "Core"
    If traceRunOwnedByProcedure Then CompleteTraceRun "Configured macro target failed: " & MacroKey, False
    RunConfiguredMacroTarget = False
End Function

Private Function RunConfiguredMacroTargetCore(ByVal MacroKey As String) As Boolean
    On Error GoTo ErrHandler
    Dim cfg As TMacroRecord
    If LoadMacroConfigByKey(MacroKey, cfg) = False Then
        Err.Raise vbObjectError + 5300, "RunConfiguredMacroTargetCore", "No macro config row was found for MacroKey: " & MacroKey & vbCrLf & "MC_Records table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & vbCrLf & "The macro will not run until an active config row with this MacroKey exists."
    End If
    If cfg.Active = False Then
        Err.Raise vbObjectError + 5301, "RunConfiguredMacroTargetCore", "The macro config row is inactive." & vbCrLf & "MacroKey: " & MacroKey
    End If
    If ConfigurationStatusBlocksGeneration(cfg.ConfigurationStatus) Then
        Err.Raise vbObjectError + 5303, "RunConfiguredMacroTargetCore", "ConfigurationStatus prevents runtime execution." & vbCrLf & "MacroKey: " & MacroKey & vbCrLf & "ConfigurationStatus: " & cfg.ConfigurationStatus
    End If
    If IsValidConfigOwnerDatabaseRole(cfg.ConfigOwnerDatabaseRole) = False Then
        Err.Raise vbObjectError + 5304, "RunConfiguredMacroTargetCore", "ConfigOwnerDatabaseRole is blank or unsupported; the target will not be run." & vbCrLf & "MacroKey: " & MacroKey & vbCrLf & "ConfigOwnerDatabaseRole: " & IIf(Len(Trim$(cfg.ConfigOwnerDatabaseRole)) = 0, "(blank)", cfg.ConfigOwnerDatabaseRole)
    End If
    Select Case UCase$(Trim$(cfg.MacroAction))
        Case UCase$(ACTION_RUNCODE)
            RunConfiguredMacroTargetCore = RunGeneratedRunCodeTarget(cfg)
        Case UCase$(ACTION_OPENFORM), UCase$(ACTION_OPENTABLE), UCase$(ACTION_OPENQUERY), UCase$(ACTION_OPENREPORT), UCase$(ACTION_RUNMACRO)
            RunConfiguredMacroTargetCore = RunConfiguredAccessObjectTarget(cfg)
        Case Else
            Err.Raise vbObjectError + 5302, "RunConfiguredMacroTargetCore", "Unsupported MacroAction: " & cfg.MacroAction & vbCrLf & "MacroKey: " & cfg.MacroKey & vbCrLf & "MacroTitle: " & cfg.MacroTitle & vbCrLf & "Configuration row ID: " & CStr(cfg.ID) & vbCrLf & "MacroAction must be RunCode, OpenForm, OpenTable, OpenQuery, OpenReport, or RunMacro."
    End Select
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "RunConfiguredMacroTargetCore", "Could not run configured macro target." & vbCrLf & vbCrLf & "MacroKey: " & MacroKey & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function RunConfiguredAccessObjectTarget(ByRef cfg As TMacroRecord) As Boolean
    On Error GoTo ErrHandler
    Dim targetName As String
    Dim reportOpenTimer As Single
    Dim reportOpenElapsedSeconds As Double
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    targetName = ConfiguredObjectActionTargetName(cfg)
    If Len(targetName) = 0 Then
        Err.Raise vbObjectError + 5306, "RunConfiguredAccessObjectTarget", "The configured Access object target name is blank. Populate TargetObjectName." & vbCrLf & "MacroKey: " & cfg.MacroKey & vbCrLf & "MacroAction: " & cfg.MacroAction
    End If
    If StrComp(NormalizedConfigOwnerDatabaseRole(cfg.ConfigOwnerDatabaseRole), CONFIG_OWNER_DATABASE_ROLE_CODE_DB, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 5305, "RunConfiguredAccessObjectTarget", cfg.MacroAction & " targets with ConfigOwnerDatabaseRole=CodeDb are not directly supported. Use a HostDb-owned target object and ConfigOwnerDatabaseRole=HostDb." & vbCrLf & "MacroKey: " & cfg.MacroKey & vbCrLf & "Object: " & targetName
    End If
    Select Case UCase$(Trim$(cfg.MacroAction))
        Case UCase$(ACTION_OPENFORM)
            DoCmd.OpenForm targetName
        Case UCase$(ACTION_OPENTABLE)
            DoCmd.OpenTable targetName
        Case UCase$(ACTION_OPENQUERY)
            DoCmd.OpenQuery targetName
        Case UCase$(ACTION_OPENREPORT)
            On Error Resume Next
            LogTraceMessage TRACE_LEVEL_INFO, "ReportOpen", "RunConfiguredAccessObjectTarget", "Configured report open dispatch starting.", "MacroKey: " & cfg.MacroKey & vbCrLf & "Report: " & targetName & vbCrLf & "Requested view: Print Preview (acViewPreview)" & vbCrLf & "ConfigOwnerDatabaseRole: " & cfg.ConfigOwnerDatabaseRole, "Run configured report", 0, vbNullString, vbNullString, targetName, "Report", "HostDb", "Performance"
            Err.Clear
            On Error GoTo ErrHandler
            reportOpenTimer = Timer
            DoCmd.OpenReport targetName, acViewPreview
            reportOpenElapsedSeconds = ElapsedSecondsFromTimer(reportOpenTimer)
            On Error Resume Next
            LogTraceMessage TRACE_LEVEL_INFO, "ReportOpen", "RunConfiguredAccessObjectTarget", "Configured report open returned successfully.", "MacroKey: " & cfg.MacroKey & vbCrLf & "Report: " & targetName & vbCrLf & "Requested view: Print Preview (acViewPreview)" & vbCrLf & "DoCmd.OpenReport elapsed: " & FormattedElapsedSeconds(reportOpenElapsedSeconds), "Run configured report", 0, vbNullString, vbNullString, targetName, "Report", "HostDb", "Performance"
            Err.Clear
            On Error GoTo ErrHandler
        Case UCase$(ACTION_RUNMACRO)
            DoCmd.RunMacro targetName
        Case Else
            Err.Raise vbObjectError + 5307, "RunConfiguredAccessObjectTarget", "Unsupported Access object MacroAction: " & cfg.MacroAction & vbCrLf & "MacroKey: " & cfg.MacroKey & vbCrLf & "MacroTitle: " & cfg.MacroTitle & vbCrLf & "Target object: " & IIf(Len(Trim$(cfg.TargetObjectName)) = 0, "(blank)", cfg.TargetObjectName) & vbCrLf & "Supported object actions: OpenForm, OpenTable, OpenQuery, OpenReport, RunMacro."
    End Select
    RunConfiguredAccessObjectTarget = True
    Exit Function
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If StrComp(Trim$(cfg.MacroAction), ACTION_OPENREPORT, vbTextCompare) = 0 Then
        reportOpenElapsedSeconds = 0
        If reportOpenTimer <> 0 Then reportOpenElapsedSeconds = ElapsedSecondsFromTimer(reportOpenTimer)
        On Error Resume Next
        LogTraceMessage TRACE_LEVEL_ERROR, "ReportOpen", "RunConfiguredAccessObjectTarget", "Configured report open failed.", "MacroKey: " & cfg.MacroKey & vbCrLf & "Report: " & targetName & vbCrLf & "Requested view: Print Preview (acViewPreview)" & vbCrLf & "DoCmd.OpenReport elapsed before failure: " & FormattedElapsedSeconds(reportOpenElapsedSeconds) & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, "Run configured report", savedErrNumber, savedErrSource, savedErrDescription, targetName, "Report", "HostDb", "Performance"
        Err.Clear
        On Error GoTo 0
    End If
    Err.Raise savedErrNumber, savedErrSource, savedErrDescription
End Function

Private Function RunGeneratedRunCodeTarget(ByRef cfg As TMacroRecord) As Boolean
    On Error GoTo ErrHandler
    RunGeneratedRunCodeTarget = RunStructuredProcedureTarget(cfg)
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "RunGeneratedRunCodeTarget", "Could not run configured RunCode target." & vbCrLf & "RunCode target: " & SafeEffectiveTargetCallTextForError(cfg) & vbCrLf & "ConfigOwnerDatabaseRole: " & cfg.ConfigOwnerDatabaseRole & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function RunStructuredProcedureTarget(ByRef cfg As TMacroRecord) As Boolean
    On Error GoTo ErrHandler
    Dim args As Collection
    Dim resultValue As Variant
    Dim runProcedureName As String
    If Len(Trim$(cfg.TargetProcedureName)) = 0 Then
        Err.Raise vbObjectError + 5310, "RunStructuredProcedureTarget", "TargetProcedureName is blank." & vbCrLf & "MacroKey: " & cfg.MacroKey & vbCrLf & "MacroTitle: " & cfg.MacroTitle & vbCrLf & "ConfigOwnerDatabaseRole: " & cfg.ConfigOwnerDatabaseRole & vbCrLf & "Populate TargetProcedureName before running this RunCode target."
    End If
    If IsValidConfigOwnerDatabaseRole(cfg.ConfigOwnerDatabaseRole) = False Then
        Err.Raise vbObjectError + 5312, "RunStructuredProcedureTarget", "ConfigOwnerDatabaseRole is blank or unsupported." & vbCrLf & "ConfigOwnerDatabaseRole: " & IIf(Len(Trim$(cfg.ConfigOwnerDatabaseRole)) = 0, "(blank)", cfg.ConfigOwnerDatabaseRole)
    End If
    EnsureCallableProcedureConfigOwnerDatabaseRole cfg.TargetProcedureName, cfg.ConfigOwnerDatabaseRole, cfg.TargetProcedureKind, cfg.TargetProcedureModuleName
    Set args = ParseProcedureArguments(cfg.TargetArguments)
    runProcedureName = QualifiedRunProcedureName(cfg)
    resultValue = SharedRun_ApplicationRunWithArgs(runProcedureName, args)
    RunStructuredProcedureTarget = ResultVariantToBooleanByHandling(resultValue, cfg.ResultHandling, True)
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "RunStructuredProcedureTarget", "Could not run structured procedure target." & vbCrLf & "TargetProcedureName: " & cfg.TargetProcedureName & vbCrLf & "TargetProcedureKind: " & cfg.TargetProcedureKind & vbCrLf & "ConfigOwnerDatabaseRole: " & cfg.ConfigOwnerDatabaseRole & vbCrLf & "TargetProcedureModuleName: " & IIf(Len(Trim$(cfg.TargetProcedureModuleName)) = 0, "(not specified)", cfg.TargetProcedureModuleName) & vbCrLf & "TargetArguments: " & cfg.TargetArguments & vbCrLf & "HostDb path: " & SafeHostDbNameForError() & vbCrLf & "CodeDb path: " & SafeCodeDbNameForError() & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function EffectiveActionDescription(ByRef cfg As TMacroRecord) As String
    Dim phraseText As String
    Dim dashPos As Long
    phraseText = Trim$(cfg.ActionDescription)
    If Len(phraseText) = 0 Then
        phraseText = Trim$(cfg.MacroTitle)
        dashPos = InStr(1, phraseText, " - ", vbBinaryCompare)
        If dashPos > 0 Then phraseText = Trim$(Mid$(phraseText, dashPos + 3))
    End If
    If Len(phraseText) = 0 Then phraseText = EffectiveTargetCallText(cfg)
    If Len(phraseText) = 0 Then phraseText = "Run the configured macro action"
    If Right$(phraseText, 1) <> "." Then phraseText = phraseText & "."
    EffectiveActionDescription = phraseText
End Function

Private Function TargetMetadataWarningText(ByRef cfg As TMacroRecord) As String
    On Error GoTo ErrHandler
    Dim currentPath As String
    Dim warningText As String
    If Len(Trim$(cfg.GenerationTargetDbPath)) = 0 Then Exit Function
    currentPath = DatabasePathForRole(cfg.ConfigOwnerDatabaseRole)
    If Len(Trim$(currentPath)) = 0 Then Exit Function
    If StrComp(NormalizedDatabasePath(cfg.GenerationTargetDbPath), NormalizedDatabasePath(currentPath), vbTextCompare) <> 0 Then
        warningText = vbCrLf & vbCrLf & "Warning: stored generation target metadata differs from the currently resolved database path." & vbCrLf & "ConfigOwnerDatabaseRole: " & cfg.ConfigOwnerDatabaseRole & vbCrLf & "Stored GenerationTargetDbPath: " & cfg.GenerationTargetDbPath & vbCrLf & "Resolved path: " & currentPath & vbCrLf & "ConfigOwnerDatabaseRole remains authoritative for routing."
    End If
    TargetMetadataWarningText = warningText
    Exit Function
ErrHandler:
    TargetMetadataWarningText = vbNullString
End Function

Private Function ActionIsAccessObjectAction(ByVal MacroAction As String) As Boolean
    Select Case UCase$(Trim$(MacroAction))
        Case UCase$(ACTION_OPENFORM), UCase$(ACTION_OPENTABLE), UCase$(ACTION_OPENQUERY), UCase$(ACTION_OPENREPORT), UCase$(ACTION_RUNMACRO)
            ActionIsAccessObjectAction = True
    End Select
End Function

Private Function ResultVariantToBoolean(ByVal resultValue As Variant, ByVal DefaultValue As Boolean) As Boolean
    If IsEmpty(resultValue) Then
        ResultVariantToBoolean = DefaultValue
    ElseIf IsNull(resultValue) Then
        ResultVariantToBoolean = DefaultValue
    ElseIf VarType(resultValue) = vbBoolean Then
        ResultVariantToBoolean = CBool(resultValue)
    Else
        ResultVariantToBoolean = DefaultValue
    End If
End Function

Private Function ResultVariantToBooleanByHandling(ByVal resultValue As Variant, ByVal ResultHandling As String, ByVal DefaultValue As Boolean) As Boolean
    Select Case LCase$(Trim$(ResultHandling))
        Case "", LCase$(RESULT_HANDLING_AUTO)
            ResultVariantToBooleanByHandling = ResultVariantToBoolean(resultValue, DefaultValue)
        Case LCase$(RESULT_HANDLING_IGNORE)
            ResultVariantToBooleanByHandling = True
        Case LCase$(RESULT_HANDLING_BOOLEAN), LCase$(RESULT_HANDLING_REQUIRE_TRUE)
            If IsEmpty(resultValue) Or IsNull(resultValue) Then
                ResultVariantToBooleanByHandling = False
            Else
                ResultVariantToBooleanByHandling = CBool(resultValue)
            End If
        Case Else
            ResultVariantToBooleanByHandling = ResultVariantToBoolean(resultValue, DefaultValue)
    End Select
End Function

Private Sub UpdateValidationStatus(ByVal configID As Long, ByVal resultText As String, ByVal notesText As String)
    On Error GoTo ErrHandler
    If configID <= 0 Then Exit Sub
    ADCDb().Execute "UPDATE [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "] SET " & SQL_FIELD_LAST_VALIDATED_ON & "=Now(), " & SQL_FIELD_LAST_VALIDATION_RESULT & "=" & SharedDao_SqlQuote(Left$(resultText, 30)) & ", " & SQL_FIELD_LAST_VALIDATION_NOTES & "=" & IIf(Len(Trim$(notesText)) = 0, "Null", SharedDao_SqlQuote(notesText)) & " " & "WHERE " & SQL_FIELD_ID & "=" & CStr(configID), dbFailOnError
    InvalidateDuplicateReviewCache
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "UpdateValidationStatus", "Could not persist macro-validation status for configuration row ID " & CStr(configID) & "." & vbCrLf & "Requested result: " & resultText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Sub

Private Function RunCodeTargetValidationStatus(ByVal runCodeExpression As String, ByRef resolvedProcedureName As String, ByRef detailText As String) As String
    On Error GoTo ErrHandler
    Dim directProcedureName As String
    Dim directStatus As String
    Dim directDetail As String
    Dim directResolvedProcedureName As String
    Dim MacroKey As String
    Dim cfg As TMacroRecord
    Dim targetCallText As String
    Dim chainText As String
    Dim dispatcherWorkflow As String
    directProcedureName = ProcedureNameFromRunCodeExpression(runCodeExpression)
    If Len(directProcedureName) = 0 Then
        resolvedProcedureName = "(not resolved)"
        detailText = "The RunCode action argument is blank or could not be parsed."
        RunCodeTargetValidationStatus = "UNRESOLVED"
        Exit Function
    End If
    Select Case LCase$(directProcedureName)
        Case "runconfiguredmacrotarget", "runconfirmedconfiguredmacrotarget"
            directStatus = ValidateCallableRunCodeExpression(directProcedureName & "()", directResolvedProcedureName, directDetail, vbNullString)
            If directStatus <> "VALID" Then
                resolvedProcedureName = directProcedureName
                detailText = "The macro calls a table-driven dispatcher, but the dispatcher procedure itself is not safely callable." & vbCrLf & directDetail
                RunCodeTargetValidationStatus = directStatus
                Exit Function
            End If
            MacroKey = FirstStringArgumentFromFunctionCall(runCodeExpression)
            If Len(MacroKey) = 0 Then
                resolvedProcedureName = directProcedureName
                detailText = "The table-driven dispatcher call did not contain a resolvable MacroKey string argument."
                RunCodeTargetValidationStatus = "UNRESOLVED"
                Exit Function
            End If
            dispatcherWorkflow = CONFIG_OWNER_DATABASE_ROLE_CODE_DB
            If LoadMacroConfigByKey(MacroKey, cfg) = False Then
                resolvedProcedureName = directProcedureName & "(" & MacroKey & ")"
                detailText = "The generated macro calls the table-driven dispatcher, but no config row exists for MacroKey in the ConfigDb: " & MacroKey
                RunCodeTargetValidationStatus = "UNRESOLVED"
                Exit Function
            End If
            chainText = "Dispatcher found: " & directProcedureName & vbCrLf & "Dispatcher availability: " & directDetail & vbCrLf & "Dispatcher workflow: " & dispatcherWorkflow & vbCrLf & "MacroKey: " & MacroKey & vbCrLf & "Configured macro title: " & cfg.MacroTitle & vbCrLf & "Configured action: " & cfg.MacroAction & vbCrLf
            If ActionIsAccessObjectAction(cfg.MacroAction) Then
                RunCodeTargetValidationStatus = ValidateConfiguredAccessObjectTarget(cfg, resolvedProcedureName, detailText)
                detailText = chainText & "Configured object target: " & ConfiguredObjectActionTargetName(cfg) & vbCrLf & detailText
                Exit Function
            End If
            If StrComp(cfg.MacroAction, ACTION_RUNCODE, vbTextCompare) <> 0 Then
                resolvedProcedureName = cfg.MacroAction
                detailText = chainText & "Unsupported configured MacroAction."
                RunCodeTargetValidationStatus = "UNRESOLVED"
                Exit Function
            End If
            targetCallText = EffectiveTargetCallText(cfg)
            RunCodeTargetValidationStatus = ValidateConfiguredRunCodeTarget(cfg, resolvedProcedureName, detailText)
            detailText = chainText & "Configured RunCode target: " & targetCallText & vbCrLf & detailText
            Exit Function
        Case "runconfiguredmacrotarget" & Chr$(95) & "int", "runconfirmedconfiguredmacrotarget" & Chr$(95) & "int"
            resolvedProcedureName = directProcedureName
            detailText = "Existing macro object still calls a removed wrapper. Regenerate or delete this macro object."
            RunCodeTargetValidationStatus = "UNRESOLVED"
            Exit Function
        Case "runconfirmedmacrotarget"
            resolvedProcedureName = directProcedureName
            detailText = "RunConfirmedMacroTarget expression handling is not supported. Use RunConfirmedConfiguredMacroTarget with a MacroKey."
            RunCodeTargetValidationStatus = "UNRESOLVED"
            Exit Function
        Case Else
            RunCodeTargetValidationStatus = ValidateCallableRunCodeExpression(runCodeExpression, resolvedProcedureName, detailText, vbNullString)
            Exit Function
    End Select
ErrHandler:
    resolvedProcedureName = directProcedureName
    detailText = "Validation error while resolving RunCode target." & vbCrLf & "Err.Description: " & Err.Description
    RunCodeTargetValidationStatus = "UNRESOLVED"
End Function

Private Function ValidateConfiguredRunCodeTarget(ByRef cfg As TMacroRecord, ByRef resolvedProcedureName As String, ByRef detailText As String) As String
    Dim targetCallText As String
    Dim expectedProcedureKind As String
    Dim statusText As String
    Dim argumentStatus As String
    Dim argumentDetail As String
    targetCallText = EffectiveTargetCallText(cfg)
    expectedProcedureKind = ExpectedProcedureKindForValidation(cfg.TargetProcedureKind)
    If Len(Trim$(cfg.TargetProcedureName)) = 0 Then
        resolvedProcedureName = "(not configured)"
        detailText = "Configured RunCode row has no TargetProcedureName. TargetExpression is not supported."
        ValidateConfiguredRunCodeTarget = "UNRESOLVED"
        Exit Function
    End If
    statusText = ValidateCallableRunCodeExpression(targetCallText, resolvedProcedureName, detailText, cfg.ConfigOwnerDatabaseRole, expectedProcedureKind)
    If StrComp(statusText, "VALID", vbTextCompare) = 0 Then
        argumentStatus = ValidateConfiguredTargetArgumentCompatibility(cfg, argumentDetail)
        If StrComp(argumentStatus, "VALID", vbTextCompare) <> 0 Then
            detailText = detailText & vbCrLf & vbCrLf & argumentDetail
            ValidateConfiguredRunCodeTarget = argumentStatus
            Exit Function
        End If
        detailText = detailText & vbCrLf & vbCrLf & argumentDetail & TargetMetadataWarningText(cfg)
    End If
    ValidateConfiguredRunCodeTarget = statusText
    detailText = detailText & vbCrLf & "Structured target fields used." & vbCrLf & "TargetProcedureName: " & cfg.TargetProcedureName & vbCrLf & "TargetProcedureKind: " & IIf(Len(Trim$(cfg.TargetProcedureKind)) = 0, TARGET_PROCEDURE_KIND_AUTO, cfg.TargetProcedureKind) & vbCrLf & "TargetArguments: " & IIf(Len(Trim$(cfg.TargetArguments)) = 0, "(none)", cfg.TargetArguments) & vbCrLf & "ResultHandling: " & IIf(Len(Trim$(cfg.ResultHandling)) = 0, RESULT_HANDLING_AUTO, cfg.ResultHandling) & vbCrLf & "ExpectedReturnType: " & IIf(Len(Trim$(cfg.ExpectedReturnType)) = 0, "(not specified)", cfg.ExpectedReturnType)
End Function

Private Function ValidateCallableRunCodeExpression(ByVal runCodeExpression As String, ByRef resolvedProcedureName As String, ByRef detailText As String, Optional ByVal ConfigOwnerDatabaseRole As String = vbNullString, Optional ByVal expectedProcedureKind As String = vbNullString) As String
    Dim ProcedureName As String
    Dim localKind As String
    Dim localLocation As String
    Dim referencedKind As String
    Dim referencedLocation As String
    Dim foundLocal As Boolean
    Dim foundReferenced As Boolean
    Dim roleText As String
    Dim vbeAccessDetail As String
    Dim ignoredLocalArgumentSignature As String
    Dim ignoredLocalRequiredArgumentCount As Long
    Dim ignoredLocalOptionalArgumentCount As Long
    Dim ignoredLocalHasParamArray As Boolean
    Dim ignoredLocalExpectedReturnType As String
    Dim ignoredReferencedArgumentSignature As String
    Dim ignoredReferencedRequiredArgumentCount As Long
    Dim ignoredReferencedOptionalArgumentCount As Long
    Dim ignoredReferencedHasParamArray As Boolean
    Dim ignoredReferencedExpectedReturnType As String
    If CanAccessVBEProjectModel(vbeAccessDetail) = False Then
        resolvedProcedureName = "(VBE unavailable)"
        detailText = vbeAccessDetail
        ValidateCallableRunCodeExpression = "VBE ACCESS ERROR"
        Exit Function
    End If
    ClearVBEProcedureSearchError
    ProcedureName = ProcedureNameFromRunCodeExpression(runCodeExpression)
    If Len(ProcedureName) = 0 Then
        resolvedProcedureName = "(not resolved)"
        detailText = "The effective RunCode target could not be parsed."
        ValidateCallableRunCodeExpression = "UNRESOLVED"
        Exit Function
    End If
    roleText = NormalizedConfigOwnerDatabaseRole(ConfigOwnerDatabaseRole)
    If Len(roleText) = 0 Then
        foundLocal = PublicCallableProcedureInfoInStandardModule(ProcedureName, CONFIG_OWNER_DATABASE_ROLE_HOST_DB, localKind, localLocation, ignoredLocalArgumentSignature, ignoredLocalRequiredArgumentCount, ignoredLocalOptionalArgumentCount, ignoredLocalHasParamArray, ignoredLocalExpectedReturnType)
        If VBEProcedureSearchAmbiguityOccurred() Then
            resolvedProcedureName = ProcedureName
            detailText = "Direct RunCode target is ambiguous within the HostDb MC_Records-owner role." & vbCrLf & VBEProcedureSearchAmbiguityText()
            ValidateCallableRunCodeExpression = "AMBIGUOUS"
            Exit Function
        End If
        If CodeDbIsHostDb() = False Then
            foundReferenced = PublicCallableProcedureInfoInStandardModule(ProcedureName, CONFIG_OWNER_DATABASE_ROLE_CODE_DB, referencedKind, referencedLocation, ignoredReferencedArgumentSignature, ignoredReferencedRequiredArgumentCount, ignoredReferencedOptionalArgumentCount, ignoredReferencedHasParamArray, ignoredReferencedExpectedReturnType)
            If VBEProcedureSearchAmbiguityOccurred() Then
                resolvedProcedureName = ProcedureName
                detailText = "Direct RunCode target is ambiguous within the CodeDb MC_Records-owner role." & vbCrLf & VBEProcedureSearchAmbiguityText()
                ValidateCallableRunCodeExpression = "AMBIGUOUS"
                Exit Function
            End If
        End If
        If VBEProcedureSearchErrorOccurred() Then
            resolvedProcedureName = ProcedureName
            detailText = "VBE access failed before validation could complete the HostDb/CodeDb ambiguity check." & vbCrLf & VBEProcedureSearchErrorText()
            ValidateCallableRunCodeExpression = "VBE ACCESS ERROR"
            Exit Function
        End If
        If foundLocal And foundReferenced Then
            resolvedProcedureName = ProcedureName
            detailText = "Direct RunCode target is ambiguous across HostDb and CodeDb MC_Records-owner roles." & vbCrLf & "HostDb location: " & localLocation & vbCrLf & "CodeDb location: " & referencedLocation & vbCrLf & "Use table-driven config with an explicit ConfigOwnerDatabaseRole or rename one procedure."
            ValidateCallableRunCodeExpression = "AMBIGUOUS"
            Exit Function
        End If
        If foundLocal Then
            If ProcedureKindMatchesExpected(localKind, expectedProcedureKind) = False Then
                resolvedProcedureName = ProcedureName
                detailText = "Direct RunCode target was found in the HostDb MC_Records-owner role, but the procedure kind does not match." & vbCrLf & "Expected kind: " & ExpectedProcedureKindForValidation(expectedProcedureKind) & vbCrLf & "Found kind: " & localKind & vbCrLf & "Location: " & localLocation
                ValidateCallableRunCodeExpression = "MISSING"
                Exit Function
            End If
            resolvedProcedureName = ProcedureName
            detailText = "Direct RunCode target found in the HostDb MC_Records-owner role: " & ProcedureName & " (" & localKind & "; " & localLocation & ")"
            ValidateCallableRunCodeExpression = "VALID"
            Exit Function
        End If
        If foundReferenced Then
            If ProcedureKindMatchesExpected(referencedKind, expectedProcedureKind) = False Then
                resolvedProcedureName = ProcedureName
                detailText = "Direct RunCode target was found in the CodeDb MC_Records-owner role, but the procedure kind does not match." & vbCrLf & "Expected kind: " & ExpectedProcedureKindForValidation(expectedProcedureKind) & vbCrLf & "Found kind: " & referencedKind & vbCrLf & "Location: " & referencedLocation
                ValidateCallableRunCodeExpression = "MISSING"
                Exit Function
            End If
            resolvedProcedureName = ProcedureName
            detailText = "Direct RunCode target found in the CodeDb MC_Records-owner role: " & ProcedureName & " (" & referencedKind & "; " & referencedLocation & ")"
            ValidateCallableRunCodeExpression = "VALID"
            Exit Function
        End If
        resolvedProcedureName = ProcedureName
        detailText = "Missing direct RunCode target procedure: " & ProcedureName & vbCrLf & "Config-owner roles searched: HostDb" & IIf(CodeDbIsHostDb(), "", " and CodeDb") & vbCrLf & "Expected: Public or implicit-public Function/Sub in a standard module."
        ValidateCallableRunCodeExpression = "MISSING"
        Exit Function
    End If
    ValidateCallableRunCodeExpression = ValidateCallableProcedureInExplicitConfigOwnerDatabaseRole(ProcedureName, roleText, expectedProcedureKind, resolvedProcedureName, detailText)
End Function

Private Sub BuildRunnerMacroTargetValidationCollections(ByRef configuredMissingMacroEntries As Collection, ByRef missingEntries As Collection, ByRef unresolvedEntries As Collection, ByRef ambiguousEntries As Collection, ByRef vbeAccessErrorEntries As Collection, ByRef validEntries As Collection, ByRef nonRunCodeEntries As Collection, ByRef exportErrorEntries As Collection)
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim cfg As TMacroRecord
    Dim macroTitleText As String
    Dim processedCount As Long
    AddConfiguredButMissingMacroObjectEntries configuredMissingMacroEntries
    If CodeDbTableExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")) = False Then Exit Sub
    Set rs = ADCDb().OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "] WHERE " & SQL_FIELD_ACTIVE & "<>0 ORDER BY Nz(" & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER & ",0), " & SQL_FIELD_MACRO_TITLE & ", " & SQL_FIELD_MACRO_KEY, dbOpenSnapshot)
    Do While Not rs.EOF
        PumpSetupResponsiveness
        processedCount = processedCount + 1
        PulseProgressEvery processedCount, 0, "Generated macro objects validated"
        MacroRecordFromRecordset rs, cfg
        macroTitleText = cfg.MacroTitle
        If Len(Trim$(macroTitleText)) > 0 Then
            If AccessObjectExistsInCurrentDatabase(acMacro, macroTitleText) Then
                AddConfiguredMacroObjectValidationEntries macroTitleText, cfg, missingEntries, unresolvedEntries, ambiguousEntries, vbeAccessErrorEntries, validEntries, nonRunCodeEntries, exportErrorEntries
            End If
        End If
        rs.MoveNext
    Loop
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
    Err.Raise savedErrNumber, "BuildRunnerMacroTargetValidationCollections", "Could not build Runner Macro target validation collections." & vbCrLf & "MC_Records table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & vbCrLf & "Validation source: active/generated macro config rows and existing Access macro objects." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub AddMacroValidationEntryByStatus(ByVal statusText As String, ByVal entryText As String, ByRef missingEntries As Collection, ByRef unresolvedEntries As Collection, ByRef ambiguousEntries As Collection, ByRef vbeAccessErrorEntries As Collection, ByRef validEntries As Collection, ByRef nonRunCodeEntries As Collection, ByRef exportErrorEntries As Collection)
    Select Case UCase$(statusText)
        Case "MISSING"
            missingEntries.Add entryText
        Case "UNRESOLVED"
            unresolvedEntries.Add entryText
        Case "AMBIGUOUS"
            ambiguousEntries.Add entryText
        Case "VBE ACCESS ERROR"
            vbeAccessErrorEntries.Add entryText
        Case "VALID"
            validEntries.Add entryText
        Case "NON-RUNCODE"
            nonRunCodeEntries.Add entryText
        Case "EXPORT ERROR"
            exportErrorEntries.Add entryText
        Case Else
            unresolvedEntries.Add entryText
    End Select
End Sub

Private Function RunnerMacroTargetValidationReport(ByVal configuredMissingMacroEntries As Collection, ByVal missingEntries As Collection, ByVal unresolvedEntries As Collection, ByVal ambiguousEntries As Collection, ByVal vbeAccessErrorEntries As Collection, ByVal validEntries As Collection, ByVal nonRunCodeEntries As Collection, ByVal exportErrorEntries As Collection) As String
    Dim lines As Collection
    ' INTENTIONAL ORDERED VIEW: lines preserves the deterministic sequence required by RunnerMacroTargetValidationReport; keyed-only or unordered storage would not preserve the consumer semantics.
    Set lines = New Collection
    lines.Add "Runner Macro Target Validation Report"
    lines.Add "Generated: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    lines.Add "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB
    lines.Add "MC_Records table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")
    lines.Add "Scope: Active=True configured macro objects only; unrelated Access macros are not scanned."
    lines.Add String$(72, "=")
    lines.Add ""
    AppendReportSection lines, "CONFIGURED BUT MISSING MACRO OBJECT", configuredMissingMacroEntries
    AppendReportSection lines, "MISSING", missingEntries
    AppendReportSection lines, "UNRESOLVED", unresolvedEntries
    AppendReportSection lines, "AMBIGUOUS", ambiguousEntries
    AppendReportSection lines, "VBE ACCESS ERROR", vbeAccessErrorEntries
    AppendReportSection lines, "VALID", validEntries
    AppendReportSection lines, "NON-RUNCODE", nonRunCodeEntries
    AppendReportSection lines, "EXPORT ERROR", exportErrorEntries
    RunnerMacroTargetValidationReport = SharedCollection_Join(lines, vbCrLf)
End Function

Private Function MacroValidationReportEntry(ByVal MacroTitle As String, ByVal statusText As String, ByVal runCodeExpression As String, ByVal resolvedProcedureName As String, ByVal detailText As String) As String
    MacroValidationReportEntry = "Macro: " & MacroTitle & vbCrLf & "Status: " & statusText & vbCrLf & "Action/expression: " & IIf(Len(runCodeExpression) = 0, "(none)", runCodeExpression) & vbCrLf & "Resolved target: " & IIf(Len(resolvedProcedureName) = 0, "(none)", resolvedProcedureName) & vbCrLf & "Detail: " & IIf(Len(detailText) = 0, "(none)", detailText) & vbCrLf
End Function

Private Function RunnerMacroTargetValidationReportPath() As String
    Dim folderPath As String
    folderPath = DatabaseFolderPath()
    If Len(Trim$(folderPath)) = 0 Then folderPath = Environ$("TEMP")
    RunnerMacroTargetValidationReportPath = folderPath & "\RunnerMacroTargetValidation_" & Format$(Now, "yyyymmdd_hhnnss") & ".txt"
End Function

Private Sub OpenTextReport(ByVal reportPath As String)
    On Error GoTo ErrHandler
    Application.FollowHyperlink reportPath
    Exit Sub
ErrHandler:
    CoreShowErrorMessage "The report was created, but Access could not open it automatically." & vbCrLf & vbCrLf & "Report path: " & reportPath & vbCrLf & "Err.Description: " & Err.Description, vbCritical, "Open Report", "OpenTextReport", "Core"
End Sub

Private Function ValidateCallableProcedureInExplicitConfigOwnerDatabaseRole(ByVal ProcedureName As String, ByVal ConfigOwnerDatabaseRole As String, ByVal expectedProcedureKind As String, ByRef resolvedProcedureName As String, ByRef detailText As String) As String
    Dim procedureKind As String
    Dim procedureLocation As String
    Dim oppositeProcedureKind As String
    Dim oppositeProcedureLocation As String
    Dim oppositeRoleText As String
    Dim ignoredConfiguredArgumentSignature As String
    Dim ignoredConfiguredRequiredArgumentCount As Long
    Dim ignoredConfiguredOptionalArgumentCount As Long
    Dim ignoredConfiguredHasParamArray As Boolean
    Dim ignoredConfiguredExpectedReturnType As String
    Dim ignoredOppositeArgumentSignature As String
    Dim ignoredOppositeRequiredArgumentCount As Long
    Dim ignoredOppositeOptionalArgumentCount As Long
    Dim ignoredOppositeHasParamArray As Boolean
    Dim ignoredOppositeExpectedReturnType As String
    ClearVBEProcedureSearchError
    If PublicCallableProcedureInfoInStandardModule(ProcedureName, ConfigOwnerDatabaseRole, procedureKind, procedureLocation, ignoredConfiguredArgumentSignature, ignoredConfiguredRequiredArgumentCount, ignoredConfiguredOptionalArgumentCount, ignoredConfiguredHasParamArray, ignoredConfiguredExpectedReturnType) = False Then
        If VBEProcedureSearchAmbiguityOccurred() Then
            resolvedProcedureName = ProcedureName
            detailText = "Effective RunCode target is ambiguous inside the configured ConfigOwnerDatabaseRole." & vbCrLf & "ConfigOwnerDatabaseRole: " & ConfigOwnerDatabaseRole & vbCrLf & VBEProcedureSearchAmbiguityText()
            ValidateCallableProcedureInExplicitConfigOwnerDatabaseRole = "AMBIGUOUS"
            Exit Function
        End If
        If VBEProcedureSearchErrorOccurred() Then
            resolvedProcedureName = ProcedureName
            detailText = VBEProcedureSearchErrorText()
            ValidateCallableProcedureInExplicitConfigOwnerDatabaseRole = "VBE ACCESS ERROR"
            Exit Function
        End If
        resolvedProcedureName = ProcedureName
        detailText = "Missing effective RunCode target procedure: " & ProcedureName & vbCrLf & "ConfigOwnerDatabaseRole: " & ConfigOwnerDatabaseRole & vbCrLf & "Expected: Public or implicit-public Function/Sub in a standard module."
        ValidateCallableProcedureInExplicitConfigOwnerDatabaseRole = "MISSING"
        Exit Function
    End If
    If ProcedureKindMatchesExpected(procedureKind, expectedProcedureKind) = False Then
        resolvedProcedureName = ProcedureName
        detailText = "Effective RunCode target was found, but the procedure kind does not match the config row." & vbCrLf & "Expected kind: " & ExpectedProcedureKindForValidation(expectedProcedureKind) & vbCrLf & "Found kind: " & procedureKind & vbCrLf & "ConfigOwnerDatabaseRole: " & ConfigOwnerDatabaseRole & vbCrLf & "Location: " & procedureLocation
        ValidateCallableProcedureInExplicitConfigOwnerDatabaseRole = "MISSING"
        Exit Function
    End If
    If CodeDbIsHostDb() = False Then
        ClearVBEProcedureSearchError
        oppositeRoleText = OppositeConfigOwnerDatabaseRole(ConfigOwnerDatabaseRole)
        If PublicCallableProcedureInfoInStandardModule(ProcedureName, oppositeRoleText, oppositeProcedureKind, oppositeProcedureLocation, ignoredOppositeArgumentSignature, ignoredOppositeRequiredArgumentCount, ignoredOppositeOptionalArgumentCount, ignoredOppositeHasParamArray, ignoredOppositeExpectedReturnType) Then
            resolvedProcedureName = ProcedureName
            detailText = "Effective RunCode target is ambiguous across HostDb and CodeDb MC_Records-owner roles." & vbCrLf & "Configured config-owner role searched: " & ConfigOwnerDatabaseRole & vbCrLf & "Configured target location: " & procedureLocation & vbCrLf & "Other target-role location: " & oppositeProcedureLocation & vbCrLf & "Rename one procedure or change ConfigOwnerDatabaseRole so the generated macro cannot resolve the wrong target."
            ValidateCallableProcedureInExplicitConfigOwnerDatabaseRole = "AMBIGUOUS"
            Exit Function
        End If
        If VBEProcedureSearchAmbiguityOccurred() Then
            resolvedProcedureName = ProcedureName
            detailText = "Opposite target-role ambiguity check found duplicate callable procedures." & vbCrLf & "Opposite config-owner role searched: " & oppositeRoleText & vbCrLf & VBEProcedureSearchAmbiguityText()
            ValidateCallableProcedureInExplicitConfigOwnerDatabaseRole = "AMBIGUOUS"
            Exit Function
        End If
        If VBEProcedureSearchErrorOccurred() Then
            resolvedProcedureName = ProcedureName
            detailText = "VBE access failed before validation could complete the opposite target-role ambiguity check." & vbCrLf & VBEProcedureSearchErrorText()
            ValidateCallableProcedureInExplicitConfigOwnerDatabaseRole = "VBE ACCESS ERROR"
            Exit Function
        End If
    End If
    resolvedProcedureName = ProcedureName
    detailText = "Effective RunCode target found: " & ProcedureName & " (" & procedureKind & "; " & procedureLocation & ")" & vbCrLf & "ConfigOwnerDatabaseRole: " & ConfigOwnerDatabaseRole
    ValidateCallableProcedureInExplicitConfigOwnerDatabaseRole = "VALID"
End Function

Private Function ProcedureKindMatchesExpected(ByVal actualProcedureKind As String, ByVal expectedProcedureKind As String) As Boolean
    Dim expectedKind As String
    expectedKind = ExpectedProcedureKindForValidation(expectedProcedureKind)
    If Len(Trim$(expectedKind)) = 0 Then
        ProcedureKindMatchesExpected = True
    Else
        ProcedureKindMatchesExpected = (StrComp(actualProcedureKind, expectedKind, vbTextCompare) = 0)
    End If
End Function

' Purpose: Searches the selected configuration-owner VBA project for a callable public procedure in a standard module and returns its discovered callable metadata through ByRef outputs.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   ProcedureName (Input; String): Identifies the VBA procedure being searched, validated, described, or invoked by its procedure name.
'   ConfigOwnerDatabaseRole (Input; String): Identifies the database role whose configuration is authoritative for this operation, such as the CodeDb or HostDb role defined by ADC configuration.
'   procedureKind (Output; String): Receives the discovered VBA procedure kind (for example Sub or Function) for the target procedure.
'   procedureLocation (Output; String): Receives the discovered module/project location of the target procedure so the caller can report or validate where it was found.
'   argumentSignature (Output; String): Receives the discovered argument signature for the target procedure, including its declared parameter list.
'   RequiredArgumentCount (Output; Long): Receives the number of required arguments declared by the target procedure.
'   OptionalArgumentCount (Output; Long): Receives the number of optional arguments declared by the target procedure.
'   hasParamArray (Output; Boolean): Receives whether the target procedure declares a ParamArray argument.
'   ExpectedReturnType (Output; String): Receives the return type discovered or expected for the target callable procedure.
' Returns: Boolean: True when a matching callable public procedure is found and its metadata outputs are populated; False when no qualifying procedure is found.
Public Function PublicCallableProcedureInfoInStandardModule(ByVal ProcedureName As String, ByVal ConfigOwnerDatabaseRole As String, ByRef procedureKind As String, ByRef procedureLocation As String, ByRef argumentSignature As String, ByRef RequiredArgumentCount As Long, ByRef OptionalArgumentCount As Long, ByRef hasParamArray As Boolean, ByRef ExpectedReturnType As String) As Boolean
    'Intended use: Returns the Public Callable Procedure Info In Standard Module result for direct, query, form, or required cross-module callers.
    On Error GoTo ErrHandler
    Dim vbProject As Object
    Dim roleText As String
    Dim treatOptionPrivateAsNotCallable As Boolean
    If Len(Trim$(ProcedureName)) = 0 Then Exit Function
    roleText = NormalizedConfigOwnerDatabaseRole(ConfigOwnerDatabaseRole)
    If Len(roleText) = 0 Then roleText = CONFIG_OWNER_DATABASE_ROLE_HOST_DB
    treatOptionPrivateAsNotCallable = (StrComp(roleText, CONFIG_OWNER_DATABASE_ROLE_CODE_DB, vbTextCompare) = 0 And CodeDbIsHostDb() = False)
    Set vbProject = VBProjectForConfigOwnerDatabaseRole(roleText)
    If mPublicCallableProcedureLookupScopeActive And StrComp(roleText, mPublicCallableProcedureLookupRole, vbTextCompare) = 0 Then
        EnsurePublicCallableProcedureLookup vbProject, treatOptionPrivateAsNotCallable
        PublicCallableProcedureInfoInStandardModule = PublicCallableProcedureInfoFromLookup(vbProject, ProcedureName, procedureKind, procedureLocation, argumentSignature, RequiredArgumentCount, OptionalArgumentCount, hasParamArray, ExpectedReturnType)
    Else
        PublicCallableProcedureInfoInStandardModule = PublicCallableProcedureInfoInVBProject(vbProject, ProcedureName, treatOptionPrivateAsNotCallable, procedureKind, procedureLocation, argumentSignature, RequiredArgumentCount, OptionalArgumentCount, hasParamArray, ExpectedReturnType)
    End If
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "PublicCallableProcedureInfoInStandardModule", "Could not search standard modules for a callable procedure." & vbCrLf & "ProcedureName: " & ProcedureName & vbCrLf & "ConfigOwnerDatabaseRole: " & ConfigOwnerDatabaseRole & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Sub EnsureCallableProcedureConfigOwnerDatabaseRole(ByVal ProcedureName As String, ByVal ConfigOwnerDatabaseRole As String, ByVal expectedProcedureKind As String, Optional ByVal TargetProcedureModuleName As String = vbNullString)
    On Error GoTo ErrHandler
    Dim roleText As String
    Dim procedureKind As String
    Dim procedureLocation As String
    Dim expectedKind As String
    Dim ignoredArgumentSignature As String
    Dim ignoredRequiredArgumentCount As Long
    Dim ignoredOptionalArgumentCount As Long
    Dim ignoredHasParamArray As Boolean
    Dim ignoredExpectedReturnType As String
    ProcedureName = Trim$(ProcedureName)
    roleText = NormalizedConfigOwnerDatabaseRole(ConfigOwnerDatabaseRole)
    expectedKind = ExpectedProcedureKindForValidation(expectedProcedureKind)
    If Len(ProcedureName) = 0 Then
        Err.Raise vbObjectError + 5312, "EnsureCallableProcedureConfigOwnerDatabaseRole", "ProcedureName is blank." & vbCrLf & "ConfigOwnerDatabaseRole: " & IIf(Len(Trim$(ConfigOwnerDatabaseRole)) = 0, "(blank)", ConfigOwnerDatabaseRole) & vbCrLf & "Expected kind: " & IIf(Len(Trim$(expectedKind)) = 0, "(blank)", expectedKind) & vbCrLf & "TargetProcedureModuleName: " & IIf(Len(Trim$(TargetProcedureModuleName)) = 0, "(blank)", TargetProcedureModuleName) & vbCrLf & "Pass the configured TargetProcedureName before validating the config-owner role."
    End If
    If Len(roleText) = 0 Then
        Err.Raise vbObjectError + 5313, "EnsureCallableProcedureConfigOwnerDatabaseRole", "ConfigOwnerDatabaseRole is blank or unsupported." & vbCrLf & "ProcedureName: " & ProcedureName & vbCrLf & "ConfigOwnerDatabaseRole: " & IIf(Len(Trim$(ConfigOwnerDatabaseRole)) = 0, "(blank)", ConfigOwnerDatabaseRole)
    End If
    ClearVBEProcedureSearchError
    If PublicCallableProcedureInfoInStandardModule(ProcedureName, roleText, procedureKind, procedureLocation, ignoredArgumentSignature, ignoredRequiredArgumentCount, ignoredOptionalArgumentCount, ignoredHasParamArray, ignoredExpectedReturnType) = False Then
        If VBEProcedureSearchAmbiguityOccurred() Then
            Err.Raise vbObjectError + 5316, "EnsureCallableProcedureConfigOwnerDatabaseRole", "Target procedure name is ambiguous inside the configured ConfigOwnerDatabaseRole." & vbCrLf & "ProcedureName: " & ProcedureName & vbCrLf & "ConfigOwnerDatabaseRole: " & roleText & vbCrLf & VBEProcedureSearchAmbiguityText()
        End If
        If VBEProcedureSearchErrorOccurred() Then
            Err.Raise vbObjectError + 5317, "EnsureCallableProcedureConfigOwnerDatabaseRole", "VBE access failed while checking the configured ConfigOwnerDatabaseRole." & vbCrLf & "ProcedureName: " & ProcedureName & vbCrLf & "ConfigOwnerDatabaseRole: " & roleText & vbCrLf & VBEProcedureSearchErrorText()
        End If
        Err.Raise vbObjectError + 5314, "EnsureCallableProcedureConfigOwnerDatabaseRole", "Target procedure was not found in the configured ConfigOwnerDatabaseRole." & vbCrLf & "ProcedureName: " & ProcedureName & vbCrLf & "ConfigOwnerDatabaseRole: " & roleText & vbCrLf & "HostDb path: " & GetHostDb().Name & vbCrLf & "CodeDb path: " & codeDb.Name
    End If
    If Len(expectedKind) > 0 Then
        If StrComp(procedureKind, expectedKind, vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + 5315, "EnsureCallableProcedureConfigOwnerDatabaseRole", "Target procedure kind does not match TargetProcedureKind." & vbCrLf & "ProcedureName: " & ProcedureName & vbCrLf & "ConfigOwnerDatabaseRole: " & roleText & vbCrLf & "Expected kind: " & expectedKind & vbCrLf & "Found kind: " & procedureKind & vbCrLf & "Location: " & procedureLocation
        End If
    End If
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "EnsureCallableProcedureConfigOwnerDatabaseRole", "Could not verify callable procedure target metadata." & vbCrLf & "ProcedureName: " & ProcedureName & vbCrLf & "ConfigOwnerDatabaseRole: " & ConfigOwnerDatabaseRole & vbCrLf & "Expected kind: " & expectedKind & vbCrLf & "TargetProcedureModuleName: " & TargetProcedureModuleName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Function PublicCallableProcedureInfoInVBProject(ByVal vbProject As Object, ByVal ProcedureName As String, ByVal treatOptionPrivateAsNotCallable As Boolean, ByRef procedureKind As String, ByRef procedureLocation As String, ByRef argumentSignature As String, ByRef RequiredArgumentCount As Long, ByRef OptionalArgumentCount As Long, ByRef hasParamArray As Boolean, ByRef ExpectedReturnType As String) As Boolean
    On Error GoTo ErrHandler
    Dim vbComponent As Object
    Dim moduleErrors As Collection
    Dim moduleErrorText As String
    Dim item As Variant
    Dim matchCount As Long
    Dim matchLocations As Collection
    Dim matchText As String
    If vbProject Is Nothing Then
        AppendVBEProcedureSearchError "Could not resolve a VBProject while searching for procedure: " & ProcedureName
        Exit Function
    End If
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: moduleErrors, matchLocations holds planned/recovery work or independently observed validation evidence for PublicCallableProcedureInfoInVBProject; keep it distinct
    ' from the source authority.
    Set moduleErrors = New Collection
    Set matchLocations = New Collection
    For Each vbComponent In vbProject.VBComponents
        PumpSetupResponsiveness
        If vbComponent.Type = VBEXT_CT_STD_MODULE Then
            moduleErrorText = vbNullString
            If PublicCallableProcedureInfoInVBComponent(vbProject, vbComponent, ProcedureName, treatOptionPrivateAsNotCallable, procedureKind, procedureLocation, argumentSignature, RequiredArgumentCount, OptionalArgumentCount, hasParamArray, ExpectedReturnType, moduleErrorText) Then
                matchCount = matchCount + 1
                matchLocations.Add procedureKind & " at " & procedureLocation & IIf(Len(Trim$(argumentSignature)) = 0, " with no arguments", " with arguments (" & argumentSignature & ")")
            End If
            If Len(Trim$(moduleErrorText)) > 0 Then moduleErrors.Add moduleErrorText
        End If
    Next vbComponent
    For Each item In moduleErrors
        PumpSetupResponsiveness
        AppendVBEProcedureSearchError CStr(item)
    Next item
    If matchCount = 1 Then
        PublicCallableProcedureInfoInVBProject = True
    ElseIf matchCount > 1 Then
        For Each item In matchLocations
            PumpSetupResponsiveness
            matchText = matchText & " - " & CStr(item) & vbCrLf
        Next item
        AppendVBEProcedureSearchAmbiguity "ProcedureName: " & ProcedureName & vbCrLf & "VBProject: " & SafeObjectText(vbProject, "Name") & vbCrLf & "Matches found in the same VBProject: " & CStr(matchCount) & vbCrLf & matchText & "Application.Run cannot safely disambiguate same-named public procedures. Rename duplicates or hide non-target procedures."
    Else
        PublicCallableProcedureInfoInVBProject = PublicCallableProcedureTextScanInfoInVBProject(vbProject, ProcedureName, treatOptionPrivateAsNotCallable, procedureKind, procedureLocation, argumentSignature, RequiredArgumentCount, OptionalArgumentCount, hasParamArray, ExpectedReturnType)
    End If
    Exit Function
ErrHandler:
    AppendVBEProcedureSearchError "Could not inspect VBProject while searching for procedure [" & ProcedureName & "]. " & "Project: " & SafeObjectText(vbProject, "Name") & ". Err.Description: " & Err.Description
    PublicCallableProcedureInfoInVBProject = False
End Function

Private Function PublicCallableProcedureTextScanInfoInVBProject(ByVal vbProject As Object, ByVal ProcedureName As String, ByVal treatOptionPrivateAsNotCallable As Boolean, ByRef procedureKind As String, ByRef procedureLocation As String, ByRef argumentSignature As String, ByRef RequiredArgumentCount As Long, ByRef OptionalArgumentCount As Long, ByRef hasParamArray As Boolean, ByRef ExpectedReturnType As String) As Boolean
    On Error GoTo ErrHandler
    Dim vbComponent As Object
    Dim moduleErrors As Collection
    Dim moduleErrorText As String
    Dim item As Variant
    Dim matchCount As Long
    Dim matchLocations As Collection
    Dim matchText As String
    If vbProject Is Nothing Then Exit Function
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: moduleErrors, matchLocations holds planned/recovery work or independently observed validation evidence for PublicCallableProcedureTextScanInfoInVBProject; keep it
    ' distinct from the source authority.
    Set moduleErrors = New Collection
    Set matchLocations = New Collection
    For Each vbComponent In vbProject.VBComponents
        PumpSetupResponsiveness
        If vbComponent.Type = VBEXT_CT_STD_MODULE Then
            moduleErrorText = vbNullString
            If PublicCallableProcedureTextScanInfoInVBComponent(vbProject, vbComponent, ProcedureName, treatOptionPrivateAsNotCallable, procedureKind, procedureLocation, argumentSignature, RequiredArgumentCount, OptionalArgumentCount, hasParamArray, ExpectedReturnType, moduleErrorText) Then
                matchCount = matchCount + 1
                matchLocations.Add procedureKind & " at " & procedureLocation & IIf(Len(Trim$(argumentSignature)) = 0, " with no arguments", " with arguments (" & argumentSignature & ")")
            End If
            If Len(Trim$(moduleErrorText)) > 0 Then moduleErrors.Add moduleErrorText
        End If
    Next vbComponent
    For Each item In moduleErrors
        PumpSetupResponsiveness
        AppendVBEProcedureSearchError CStr(item)
    Next item
    If matchCount = 1 Then
        PublicCallableProcedureTextScanInfoInVBProject = True
    ElseIf matchCount > 1 Then
        For Each item In matchLocations
            PumpSetupResponsiveness
            matchText = matchText & " - " & CStr(item) & vbCrLf
        Next item
        AppendVBEProcedureSearchAmbiguity "ProcedureName: " & ProcedureName & vbCrLf & "VBProject: " & SafeObjectText(vbProject, "Name") & vbCrLf & "Matches found in the same VBProject: " & CStr(matchCount) & vbCrLf & matchText & "Application.Run cannot safely disambiguate same-named public procedures. Rename duplicates or hide non-target procedures."
    End If
    Exit Function
ErrHandler:
    AppendVBEProcedureSearchError "Fallback text scan failed while checking procedure: " & ProcedureName & ". Project: " & SafeObjectText(vbProject, "Name") & ". Description: " & Err.Description
    PublicCallableProcedureTextScanInfoInVBProject = False
End Function

Private Function PublicCallableProcedureTextScanInfoInVBComponent(ByVal vbProject As Object, ByVal vbComponent As Object, ByVal ProcedureName As String, ByVal treatOptionPrivateAsNotCallable As Boolean, ByRef procedureKind As String, ByRef procedureLocation As String, ByRef argumentSignature As String, ByRef RequiredArgumentCount As Long, ByRef OptionalArgumentCount As Long, ByRef hasParamArray As Boolean, ByRef ExpectedReturnType As String, ByRef moduleErrorText As String) As Boolean
    On Error GoTo ErrHandler
    Dim codeModule As Object
    Dim lineIndex As Long
    Dim lineText As String
    Dim declarationText As String
    Dim parsedKind As String
    Dim parsedVisibility As String
    Dim parsedSignature As String
    Dim parsedRequiredCount As Long
    Dim parsedOptionalCount As Long
    Dim parsedReturnType As String
    Set codeModule = vbComponent.codeModule
    If treatOptionPrivateAsNotCallable Then
        If ModuleHasOptionPrivateModule(codeModule) Then Exit Function
    End If
    For lineIndex = 1 To codeModule.CountOfLines
        PumpSetupResponsiveness
        lineText = Trim$(codeModule.lines(lineIndex, 1))
        If CodeLineMayDeclareProcedure(lineText, ProcedureName) Then
            declarationText = ProcedureDeclarationTextFromKnownLine(codeModule, lineIndex)
            If CandidateProcedureDeclarationInfo(declarationText, ProcedureName, parsedKind, parsedVisibility, parsedSignature, parsedRequiredCount, parsedOptionalCount, parsedReturnType) Then
                procedureKind = parsedKind
                argumentSignature = parsedSignature
                RequiredArgumentCount = parsedRequiredCount
                OptionalArgumentCount = parsedOptionalCount
                ExpectedReturnType = parsedReturnType
                hasParamArray = ProcedureSignatureHasParamArray(argumentSignature)
                procedureLocation = SafeObjectText(vbProject, "Name") & "." & vbComponent.Name & ":" & CStr(lineIndex) & " (text scan)"
                PublicCallableProcedureTextScanInfoInVBComponent = True
                Exit Function
            End If
        End If
    Next lineIndex
    Exit Function
ErrHandler:
    moduleErrorText = "Fallback text scan failed while checking procedure: " & ProcedureName & ". Project: " & SafeObjectText(vbProject, "Name") & ". Module: " & SafeObjectText(vbComponent, "Name") & ". Description: " & Err.Description
    PublicCallableProcedureTextScanInfoInVBComponent = False
End Function

Private Function PublicCallableProcedureInfoInVBComponent(ByVal vbProject As Object, ByVal vbComponent As Object, ByVal ProcedureName As String, ByVal treatOptionPrivateAsNotCallable As Boolean, ByRef procedureKind As String, ByRef procedureLocation As String, ByRef argumentSignature As String, ByRef RequiredArgumentCount As Long, ByRef OptionalArgumentCount As Long, ByRef hasParamArray As Boolean, ByRef ExpectedReturnType As String, ByRef moduleErrorText As String) As Boolean
    On Error GoTo ErrHandler
    Dim codeModule As Object
    Dim lineIndex As Long
    Dim procName As String
    Dim startLine As Long
    Dim countLines As Long
    Dim declarationText As String
    Dim visibilityText As String
    Set codeModule = vbComponent.codeModule
    If treatOptionPrivateAsNotCallable Then
        If ModuleHasOptionPrivateModule(codeModule) Then Exit Function
    End If
    lineIndex = 1
    Do While lineIndex <= codeModule.CountOfLines
        PumpSetupResponsiveness
        procName = ProcedureNameAtCodeLine(codeModule, lineIndex)
        If Len(procName) = 0 Then
            lineIndex = lineIndex + 1
        Else
            startLine = ProcedureStartLineOrDefault(codeModule, procName, lineIndex)
            countLines = ProcedureCountLinesOrDefault(codeModule, procName)
            If StrComp(procName, ProcedureName, vbTextCompare) = 0 Then
                declarationText = ProcedureDeclarationText(codeModule, startLine)
                If CandidateProcedureDeclarationInfo(declarationText, procName, procedureKind, visibilityText, argumentSignature, RequiredArgumentCount, OptionalArgumentCount, ExpectedReturnType) Then
                    hasParamArray = ProcedureSignatureHasParamArray(argumentSignature)
                    procedureLocation = SafeObjectText(vbProject, "Name") & "." & vbComponent.Name & ":" & CStr(startLine)
                    PublicCallableProcedureInfoInVBComponent = True
                    Exit Function
                End If
            End If
            If countLines > 0 Then
                lineIndex = startLine + countLines
            Else
                lineIndex = lineIndex + 1
            End If
        End If
    Loop
    Exit Function
ErrHandler:
    moduleErrorText = "Could not inspect standard module while searching for procedure [" & ProcedureName & "]. " & "Project: " & SafeObjectText(vbProject, "Name") & ". Module: " & SafeObjectText(vbComponent, "Name") & ". Err.Description: " & Err.Description
    PublicCallableProcedureInfoInVBComponent = False
End Function

Private Function EffectiveConfiguredTargetSummary(ByRef cfg As TMacroRecord) As String
    On Error GoTo ErrHandler
    If StrComp(cfg.MacroAction, ACTION_RUNCODE, vbTextCompare) = 0 Then
        EffectiveConfiguredTargetSummary = EffectiveTargetCallText(cfg)
    ElseIf ActionIsAccessObjectAction(cfg.MacroAction) Then
        EffectiveConfiguredTargetSummary = cfg.MacroAction & ": " & ConfiguredObjectActionTargetName(cfg)
    Else
        EffectiveConfiguredTargetSummary = cfg.MacroAction
    End If
    Exit Function
ErrHandler:
    EffectiveConfiguredTargetSummary = "(target summary could not be resolved: " & Err.Description & ")"
End Function

Private Sub AppendReportSection(ByRef lines As Collection, ByVal sectionTitle As String, ByVal entries As Collection)
    Dim entry As Variant
    lines.Add sectionTitle & " (" & CStr(entries.count) & ")"
    lines.Add String$(72, "-")
    If entries.count = 0 Then
        lines.Add "(none)"
    Else
        For Each entry In entries
            PumpSetupResponsiveness
            lines.Add CStr(entry)
        Next entry
    End If
    lines.Add ""
End Sub

' Purpose: Opens form targets from macro text for macro text.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   macroText (Input; String): Input macro text for open form targets from macro text; is passed to `Split`.
' Returns: Collection: Returns the collection of open form from macro text targets parsed from the supplied macro/source text, with each qualifying target represented once according to the routine's existing
'   parser rules.
Public Function OpenFormTargetsFromMacroText(ByVal macroText As String) As Collection
    Dim targets As Collection
    Dim lines() As String
    Dim lineIndex As Long
    Dim lineText As String
    Dim expectingFirstArgument As Boolean
    Dim valueText As String
    ' INTENTIONAL FILTERED PARTITION: targets is a consumer-required subset/projection derived from the source enumerated by OpenFormTargetsFromMacroText; its membership is not independently authoritative.
    Set targets = New Collection
    lines = Split(macroText, vbCrLf)
    For lineIndex = LBound(lines) To UBound(lines)
        PumpSetupResponsiveness
        lineText = Trim$(lines(lineIndex))
        If InStr(1, lineText, "Action", vbTextCompare) > 0 And InStr(1, lineText, """OpenForm""", vbTextCompare) > 0 Then
            expectingFirstArgument = True
        ElseIf expectingFirstArgument And Left$(lineText, 8) = "Argument" Then
            valueText = MacroExportLineNameValue(lineText)
            If Len(Trim$(valueText)) > 0 Then targets.Add valueText
            expectingFirstArgument = False
        ElseIf Left$(lineText, 6) = "Action" Then
            expectingFirstArgument = False
        End If
    Next lineIndex
    Set OpenFormTargetsFromMacroText = targets
End Function

' Purpose: Checks whether current state and policy permit the caller to access VBE project model.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   detailText (Output; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
' Returns: Boolean: True when the supplied/current context can access VBE project model; False otherwise.
Public Function CanAccessVBEProjectModel(ByRef detailText As String) As Boolean
    'Intended use: Returns the Can Access VBE Project Model result for direct, query, form, or required cross-module callers.
    On Error GoTo ErrHandler
    Dim projectCount As Long
    Dim activeProjectName As String
    projectCount = Application.vbe.VBProjects.count
    If projectCount > 0 Then
        activeProjectName = Application.vbe.ActiveVBProject.Name
    End If
    CanAccessVBEProjectModel = True
    Exit Function
ErrHandler:
    detailText = "VBE ACCESS ERROR: Access could not inspect the VBA project object model." & vbCrLf & "This usually means Trust Center > Macro Settings > 'Trust access to the VBA project object model' is disabled, " & "or the target VBA project is locked/protected." & vbCrLf & "Err.Description: " & Err.Description
    CanAccessVBEProjectModel = False
End Function

Private Function ConfigurationStatusBlocksGeneration(ByVal ConfigurationStatus As String) As Boolean
    Select Case LCase$(Trim$(ConfigurationStatus))
        Case "", LCase$(CONFIGURATION_STATUS_ACTIVE)
            ConfigurationStatusBlocksGeneration = False
        Case Else
            ConfigurationStatusBlocksGeneration = True
    End Select
End Function

Private Function GeneratedResultFromGenerationError(ByVal errorText As String) As String
    If InStr(1, errorText, FIELD_CONFIG_OWNER_DATABASE_ROLE, vbTextCompare) > 0 Then
        GeneratedResultFromGenerationError = GENERATED_RESULT_NEEDS_REVIEW
    ElseIf InStr(1, errorText, "ConfigurationStatus", vbTextCompare) > 0 Then
        GeneratedResultFromGenerationError = GENERATED_RESULT_SKIPPED
    Else
        GeneratedResultFromGenerationError = GENERATED_RESULT_FAILED
    End If
End Function

Private Sub RetainFailedMacroReplacementRecovery(ByVal MacroTitle As String, ByVal backupPath As String)
    On Error GoTo ErrHandler
    Dim entryData As Variant
    If mGeneratedMacroRollbackEntries Is Nothing Then Set mGeneratedMacroRollbackEntries = New Collection
    entryData = Array(MacroTitle, True, backupPath)
    mGeneratedMacroRollbackEntries.Add entryData
    mGeneratedMacroRollbackRecoveryRequired = True
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    mGeneratedMacroRollbackRecoveryRequired = True
    Err.Raise savedErrNumber, "RetainFailedMacroReplacementRecovery", "Could not retain structured recovery state after destructive macro replacement recovery failed." & vbCrLf & "Macro: " & MacroTitle & vbCrLf & "Backup: " & IIf(Len(Trim$(backupPath)) = 0, "(none)", backupPath) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Sub

' ===============================================================================================================================================================================
' CUSTOM MACRO LAUNCHER PROCEDURES
' ===============================================================================================================================================================================

Private Function LauncherFormNameForDatabaseRole(ByVal ConfigOwnerDatabaseRole As String) As String
    Dim normalizedRole As String
    normalizedRole = NormalizedConfigOwnerDatabaseRole(ConfigOwnerDatabaseRole)
    Select Case LCase$(normalizedRole)
        Case LCase$(CONFIG_OWNER_DATABASE_ROLE_CODE_DB)
            LauncherFormNameForDatabaseRole = CANON_PERSISTENT_FORM_NAME_CODEDB_CUSTOM_MACROS
        Case LCase$(CONFIG_OWNER_DATABASE_ROLE_HOST_DB)
            LauncherFormNameForDatabaseRole = CANON_PERSISTENT_FORM_NAME_HOSTDB_CUSTOM_MACROS
        Case Else
            Err.Raise vbObjectError + 5411, "LauncherFormNameForDatabaseRole", "Unknown ConfigOwnerDatabaseRole value for the canonical persistent launcher form." & vbCrLf & FIELD_CONFIG_OWNER_DATABASE_ROLE & ": " & ConfigOwnerDatabaseRole
    End Select
End Function

Private Function EnsureLauncherFormReady(Optional ByVal requestedFormName As String = vbNullString) As Boolean
    On Error GoTo ErrHandler
    Dim formName As String
    formName = requestedFormName
    If Len(formName) = 0 Then formName = LauncherFormNameForDatabaseRole(CONFIG_OWNER_DATABASE_ROLE_HOST_DB)
    EnsureLauncherFormReady = LauncherFormExistsStrict(formName, "operational launcher readiness")
    If EnsureLauncherFormReady Then Exit Function
    Err.Raise vbObjectError + 5902, "EnsureLauncherFormReady", "The required persistent generated launcher form is missing. Setup owns creation of this CodeDb UI shell; operational launcher workflows only verify and use it." & vbCrLf & "CodeDb-owned launcher form: " & formName & vbCrLf & "HostDb owns only MC_Records and generated macro objects."
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "EnsureLauncherFormReady", "Could not verify CodeDb-owned launcher form." & vbCrLf & "Form: " & formName & vbCrLf & "Err.Description: " & originalErrDescription
End Function

Private Function OpenLauncherDesignForm(ByVal formName As String, ByVal askIfOpen As Boolean) As Boolean
    On Error GoTo ErrHandler
    Dim response As VbMsgBoxResult
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    RequireSameDatabaseUICommand "OpenLauncherDesignForm", "Opening a CodeDb-owned Access Development Console launcher form in design view"
    If AccessFormIsOpen(formName) Then
        If askIfOpen Then
            response = CoreShowMessageResult("The launcher form is already open." & vbCrLf & vbCrLf & "Close it so the design can be updated?" & vbCrLf & vbCrLf & "Default answer is No.", vbExclamation Or vbYesNo Or vbDefaultButton2, "Launcher Form Already Open")
            If response <> vbYes Then Exit Function
        End If
        DoCmd.Close acForm, formName, acSavePrompt
    End If
    DoCmd.OpenForm formName, acDesign, , , , acHidden
    OpenLauncherDesignForm = True
    Exit Function
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "OpenLauncherDesignForm", "Could not open launcher form in design view." & vbCrLf & "Form: " & formName & vbCrLf & "Original Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Sub CloseLauncherDesignForm(ByVal formName As String, ByVal saveChanges As Boolean)
    On Error GoTo ErrHandler
    If AccessFormIsOpen(formName) = False Then Exit Sub
    If saveChanges Then
        DoCmd.Close acForm, formName, acSaveYes
    Else
        DoCmd.Close acForm, formName, acSaveNo
        CoreAddOperationNote "Launcher form design changes were closed without saving by user choice."
    End If
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "CloseLauncherDesignForm", "Could not close launcher form design." & vbCrLf & "Form: " & formName & vbCrLf & "Save changes: " & CStr(saveChanges) & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

Private Function AccessFormIsOpen(ByVal formName As String) As Boolean
    On Error GoTo ErrHandler
    AccessFormIsOpen = (SysCmd(acSysCmdGetObjectState, acForm, formName) <> 0)
    Exit Function
ErrHandler:
    AccessFormIsOpen = False
End Function

Private Function ControlExistsOnOpenForm(ByVal formName As String, ByVal controlName As String) As Boolean
    On Error GoTo ErrHandler
    Dim ctl As control
    Set ctl = forms(formName).Controls(controlName)
    ControlExistsOnOpenForm = True
    Exit Function
ErrHandler:
    ControlExistsOnOpenForm = False
End Function

Private Function LauncherFormExistsStrict(ByVal formName As String, ByVal stageText As String) As Boolean
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim doc As DAO.Document
    Dim containerObject As DAO.Container
    Dim objectLookupStarted As Boolean
    If Len(formName) = 0 Then Err.Raise vbObjectError + 5940, "LauncherFormExistsStrict", "A strict launcher-form inspection received a blank form name." & vbCrLf & "Stage: " & stageText
    Set db = ADCDb()
    Set containerObject = db.Containers("Forms")
    containerObject.Documents.Refresh
    objectLookupStarted = True
    Set doc = containerObject.Documents(formName)
    LauncherFormExistsStrict = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If savedErrNumber = 3265 And objectLookupStarted Then Exit Function
    Err.Raise savedErrNumber, "LauncherFormExistsStrict", "Could not authoritatively inspect a launcher form during rollback." & vbCrLf & "Database: " & SharedDao_SafeDatabaseName(ADCDb()) & vbCrLf & "Form: " & formName & vbCrLf & "Stage: " & stageText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Function

Private Sub VerifyLauncherFormLoadable(ByVal formName As String, ByVal stageText As String)
    On Error GoTo ErrHandler
    Dim tempPath As String
    Dim failureText As String
    tempPath = Environ$("TEMP") & "\ADC_LauncherFormRollbackVerify_" & SharedName_SafeFileName(formName, ADCAT_OBJECT_TYPE_MACRO, 120) & "_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & Format$(Timer * 1000, "000000") & ".txt"
    Application.SaveAsText acForm, formName, tempPath
    If GeneratedMacroRollbackBackupReadable(tempPath, failureText) = False Then Err.Raise vbObjectError + 5941, "VerifyLauncherFormLoadable", failureText
    On Error Resume Next
    If Len(tempPath) > 0 Then
        If Len(Dir$(tempPath)) > 0 Then Kill tempPath
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
    If Len(tempPath) > 0 Then
        If Len(Dir$(tempPath)) > 0 Then Kill tempPath
    End If
    On Error GoTo 0
    Err.Raise savedErrNumber, "VerifyLauncherFormLoadable", "The restored launcher form failed loadability verification." & vbCrLf & "Form: " & formName & vbCrLf & "Stage: " & stageText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Sub

Private Sub SnapshotLauncherForms(ByVal formNames As Collection)
    On Error GoTo ErrHandler
    Dim item As Variant
    Dim formName As String
    Dim backupPath As String
    Dim failureText As String
    Dim entryData As Variant
    If formNames Is Nothing Then Exit Sub
    If mLauncherFormRollbackEntries Is Nothing Then Set mLauncherFormRollbackEntries = New Collection
    For Each item In formNames
        PumpSetupResponsiveness
        formName = Trim$(CStr(item))
        If Len(formName) = 0 Then Err.Raise vbObjectError + 5942, "SnapshotLauncherForms", "An affected launcher form name was blank."
        If LauncherFormExistsStrict(formName, "before launcher-form snapshot") = False Then Err.Raise vbObjectError + 5943, "SnapshotLauncherForms", "An affected launcher form did not exist before snapshot." & vbCrLf & "Form: " & formName
        backupPath = Environ$("TEMP") & "\ADC_LauncherFormRollback_" & SharedName_SafeFileName(formName, ADCAT_OBJECT_TYPE_MACRO, 120) & "_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & Format$(Timer * 1000, "000000") & ".txt"
        Application.SaveAsText acForm, formName, backupPath
        If GeneratedMacroRollbackBackupReadable(backupPath, failureText) = False Then Err.Raise vbObjectError + 5944, "SnapshotLauncherForms", failureText
        entryData = Array(formName, True, backupPath)
        mLauncherFormRollbackEntries.Add entryData
    Next item
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If GeneratedMacroRollbackEntryCount() > 0 Then mGeneratedMacroRollbackRecoveryRequired = True
    Err.Raise savedErrNumber, "SnapshotLauncherForms", "Could not snapshot all affected launcher forms before mutation." & vbCrLf & "Form: " & formName & vbCrLf & "Backup: " & IIf(Len(backupPath) = 0, "(not created)", backupPath) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Sub

Private Function LauncherRollbackTemporaryFormName(ByVal roleText As String) As String
    'Operation-scoped noncanonical name used only while restoring one persistent launcher without deleting its canonical shell identity.
    Dim attemptIndex As Long
    Dim candidateName As String
    If Len(roleText) = 0 Then roleText = "Form"
    For attemptIndex = 1 To 100
        candidateName = "zADCLRB_" & roleText & "_" & Format$(CLng(Timer * 1000) Mod 1000000, "000000") & "_" & Format$(attemptIndex, "00")
        If LauncherFormExistsStrict(candidateName, "while reserving a rollback swap name") = False Then
            LauncherRollbackTemporaryFormName = candidateName
            Exit Function
        End If
    Next attemptIndex
    Err.Raise vbObjectError + 5970, "LauncherRollbackTemporaryFormName", "Could not reserve a temporary launcher-form rollback name."
End Function

Private Function RestorePersistentLauncherFormFromBackup(ByVal formName As String, ByVal backupPath As String, ByRef failureText As String) As Boolean
    'Restores the pre-mutation design through a verified rename/load swap. The canonical persistent launcher is never deleted as a rollback action.
    On Error GoTo ErrHandler
    Dim changedFormName As String
    Dim failedRestoreName As String
    Dim changedFormRenamed As Boolean
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    failureText = vbNullString
    If ManagedPersistentGeneratedFormIsDefined(formName) = False Then Err.Raise vbObjectError + 5971, "RestorePersistentLauncherFormFromBackup", "The requested launcher form is not classified as PersistentGenerated." & vbCrLf & "Form: " & formName
    If LauncherFormExistsStrict(formName, "before persistent launcher rollback swap") = False Then Err.Raise vbObjectError + 5972, "RestorePersistentLauncherFormFromBackup", "The persistent launcher form is missing before rollback restoration." & vbCrLf & "Form: " & formName
    changedFormName = LauncherRollbackTemporaryFormName("Changed")
    DoCmd.Rename changedFormName, acForm, formName
    changedFormRenamed = True
    Application.LoadFromText acForm, formName, backupPath
    If LauncherFormExistsStrict(formName, "after loading persistent launcher rollback backup") = False Then Err.Raise vbObjectError + 5973, "RestorePersistentLauncherFormFromBackup", "The restored persistent launcher form is missing after LoadFromText." & vbCrLf & "Form: " & formName
    VerifyLauncherFormLoadable formName, "verifying persistent launcher rollback restoration"
    If LauncherFormExistsStrict(changedFormName, "before removing the superseded changed launcher copy") Then DoCmd.DeleteObject acForm, changedFormName
    changedFormRenamed = False
    RestorePersistentLauncherFormFromBackup = True
    Exit Function
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If changedFormRenamed Then
        If LauncherFormExistsStrict(formName, "while recovering a failed persistent launcher rollback swap") Then
            failedRestoreName = LauncherRollbackTemporaryFormName("Failed")
            DoCmd.Rename failedRestoreName, acForm, formName
        End If
        If LauncherFormExistsStrict(changedFormName, "while restoring the preserved changed launcher after rollback failure") Then
            If LauncherFormExistsStrict(formName, "before restoring the preserved changed launcher to its canonical name") = False Then DoCmd.Rename formName, acForm, changedFormName
        End If
        If Len(failedRestoreName) > 0 Then
            If LauncherFormExistsStrict(failedRestoreName, "while cleaning a failed restored launcher copy") Then DoCmd.DeleteObject acForm, failedRestoreName
        End If
    End If
    On Error GoTo 0
    failureText = "Persistent launcher rollback swap failed. Form: " & formName & "; Backup: " & backupPath & "; Error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription
End Function

Private Function RollbackLauncherFormsFromRun() As String
    On Error GoTo ErrHandler
    Dim itemIndex As Long
    Dim entryData As Variant
    Dim formName As String
    Dim targetExisted As Boolean
    Dim backupPath As String
    Dim restoredCount As Long
    Dim deletedCount As Long
    Dim failedCount As Long
    Dim detailsText As String
    Dim objectWasPresent As Boolean
    Dim backupFailureText As String
    Dim entryErrNumber As Long
    Dim entryErrSource As String
    Dim entryErrDescription As String
    If mLauncherFormRollbackEntries Is Nothing Then
        RollbackLauncherFormsFromRun = "Launcher-form rollback had no tracked forms."
        Exit Function
    End If
    For itemIndex = mLauncherFormRollbackEntries.count To 1 Step -1
        PumpSetupResponsiveness
        entryData = mLauncherFormRollbackEntries(itemIndex)
        formName = vbNullString
        targetExisted = False
        backupPath = vbNullString
        If IsArray(entryData) Then
            If UBound(entryData) >= 0 Then formName = Trim$(CStr(entryData(0)))
            If UBound(entryData) >= 1 Then targetExisted = CBool(entryData(1))
            If UBound(entryData) >= 2 Then backupPath = Trim$(CStr(entryData(2)))
        End If
        If Len(formName) = 0 Then
            failedCount = failedCount + 1
            detailsText = SharedText_AppendNonBlankLine(detailsText, "Launcher-form rollback entry had a blank form name. Backup: " & IIf(Len(backupPath) = 0, "(none)", backupPath))
        Else
            On Error GoTo EntryFailed
            objectWasPresent = LauncherFormExistsStrict(formName, "before rolling back a launcher form")
            If targetExisted Then
                backupFailureText = vbNullString
                If GeneratedMacroRollbackBackupReadable(backupPath, backupFailureText) = False Then
                    failedCount = failedCount + 1
                    detailsText = SharedText_AppendNonBlankLine(detailsText, "Preserved current launcher form [" & formName & "] because its required rollback backup was unavailable. Backup: " & IIf(Len(backupPath) = 0, "(none)", backupPath) & "; " & backupFailureText)
                Else
                    If ManagedPersistentGeneratedFormIsDefined(formName) Then
                        If RestorePersistentLauncherFormFromBackup(formName, backupPath, backupFailureText) = False Then Err.Raise vbObjectError + 5945, "RollbackLauncherFormsFromRun", backupFailureText
                        restoredCount = restoredCount + 1
                        LogTraceMessage TRACE_LEVEL_INFO, "CustomMacroLauncher", "RollbackLauncherFormsFromRun", "Persistent generated launcher form restored without canonical deletion.", "Form: " & formName & vbCrLf & "CatalogOwner: " & DEFAULT_CATALOG_OWNER_PERSISTENT_GENERATED, "Rollback launcher forms", 0, vbNullString, vbNullString, formName, "Form", "CodeDb", "PersistentGenerated"
                    Else
                        If objectWasPresent Then DoCmd.DeleteObject acForm, formName
                        If LauncherFormExistsStrict(formName, "after deleting the changed launcher form") Then Err.Raise vbObjectError + 5945, "RollbackLauncherFormsFromRun", "The changed launcher form remained present after rollback deletion. Form: " & formName
                        Application.LoadFromText acForm, formName, backupPath
                        If LauncherFormExistsStrict(formName, "after restoring the original launcher form") = False Then Err.Raise vbObjectError + 5946, "RollbackLauncherFormsFromRun", "The original launcher form was not present after rollback restoration. Form: " & formName
                        VerifyLauncherFormLoadable formName, "verifying the restored original launcher form"
                        restoredCount = restoredCount + 1
                    End If
                End If
            Else
                If ManagedPersistentGeneratedFormIsDefined(formName) Then
                    detailsText = SharedText_AppendNonBlankLine(detailsText, "Persistent generated launcher shell preserved during rollback: " & formName)
                    LogTraceMessage TRACE_LEVEL_INFO, "CustomMacroLauncher", "RollbackLauncherFormsFromRun", "Persistent generated launcher form preserved during rollback.", "Form: " & formName & vbCrLf & "CatalogOwner: " & DEFAULT_CATALOG_OWNER_PERSISTENT_GENERATED, "Rollback launcher forms", 0, vbNullString, vbNullString, formName, "Form", "CodeDb", "PersistentGenerated"
                Else
                    If objectWasPresent Then DoCmd.DeleteObject acForm, formName
                    If LauncherFormExistsStrict(formName, "after deleting a newly created launcher form") Then Err.Raise vbObjectError + 5947, "RollbackLauncherFormsFromRun", "A newly created launcher form remained present after rollback deletion. Form: " & formName
                    If objectWasPresent Then deletedCount = deletedCount + 1
                End If
            End If
            On Error GoTo ErrHandler
        End If
ContinueLoop:
    Next itemIndex
    RollbackLauncherFormsFromRun = "Launcher-form rollback attempted." & vbCrLf & "New launcher forms deleted: " & CStr(deletedCount) & vbCrLf & "Launcher forms restored: " & CStr(restoredCount) & vbCrLf & "Rollback failures: " & CStr(failedCount)
    If Len(Trim$(detailsText)) > 0 Then RollbackLauncherFormsFromRun = RollbackLauncherFormsFromRun & vbCrLf & detailsText
    Exit Function
EntryFailed:
    entryErrNumber = Err.Number
    entryErrSource = Err.Source
    entryErrDescription = Err.Description
    failedCount = failedCount + 1
    detailsText = SharedText_AppendNonBlankLine(detailsText, "Launcher-form rollback failed for [" & formName & "]. Backup: " & IIf(Len(backupPath) = 0, "(none - object was newly created)", backupPath) & "; Error " & CStr(entryErrNumber) & " from " & entryErrSource & ": " & entryErrDescription)
    Err.Clear
    On Error GoTo ErrHandler
    Resume ContinueLoop
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    mGeneratedMacroRollbackRecoveryRequired = True
    RollbackLauncherFormsFromRun = "Launcher-form rollback failed before completion. Error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription & vbCrLf & GeneratedMacroRollbackPathsForMessage()
End Function

Private Sub ReviewCustomMacroLauncherOpenForm(ByVal formName As String, ByVal batchID As String, Optional ByVal ConfigOwnerDatabaseRole As String = vbNullString)
    On Error GoTo ErrHandler
    Dim ctl As control
    Dim cfg As TMacroRecord
    Dim rs As DAO.Recordset
    Dim matched As control
    Dim conflictText As String
    Dim expectedCaption As String
    Dim expectedOnClick As String
    Dim expectedTooltip As String
    Dim tagText As String
    Dim seenButtons As Object
    Dim keyText As String
    ' INTENTIONAL KEYED LOOKUP: seenButtons provides keyed membership/lookup over source rows or state already read by ReviewCustomMacroLauncherOpenForm; it is a derived access path, not a second authority.
    Set seenButtons = CreateObject("Scripting.Dictionary")
    If CustomMacroLauncherAvailable(formName) = False Then
        InsertCustomMacroLauncherReviewFinding batchID, "Error", "MissingFooter", 0, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, "Form footer is missing or unavailable. Form: " & formName
        Exit Sub
    End If
    If ControlExistsOnOpenForm(formName, LauncherAnchorControlName()) = False Then
        InsertCustomMacroLauncherReviewFinding batchID, HEALTH_SEVERITY_WARNING, "MissingAnchor", 0, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, "Launcher append anchor is missing. Form: " & formName
    End If
    Set rs = ADCDb().OpenRecordset(LauncherBackfillSql(ConfigOwnerDatabaseRole), dbOpenSnapshot)
    Do While Not rs.EOF
        PumpSetupResponsiveness
        MacroRecordFromRecordset rs, cfg
        If LauncherConfigRowQualifiesForButton(cfg) Then
            conflictText = vbNullString
            Set matched = FindLauncherButton(formName, cfg.ID, cfg.MacroKey, conflictText)
            expectedCaption = LauncherButtonCaption(cfg)
            expectedOnClick = LauncherExpectedOnClick(cfg.MacroKey)
            expectedTooltip = cfg.LauncherTooltip
            If Len(Trim$(conflictText)) > 0 Then
                InsertCustomMacroLauncherReviewFinding batchID, "Error", "NameConflict", cfg.ID, cfg.MacroKey, cfg.MacroTitle, cfg.GeneratedMacroObjectName, LauncherButtonName(cfg.ID), vbNullString, expectedCaption, expectedOnClick, vbNullString, expectedTooltip, vbNullString, conflictText
            ElseIf matched Is Nothing Then
                InsertCustomMacroLauncherReviewFinding batchID, HEALTH_SEVERITY_WARNING, "MissingButton", cfg.ID, cfg.MacroKey, cfg.MacroTitle, cfg.GeneratedMacroObjectName, LauncherButtonName(cfg.ID), vbNullString, expectedCaption, expectedOnClick, vbNullString, expectedTooltip, vbNullString, vbNullString
            Else
                If StrComp(SharedText_Nz(matched.Caption), expectedCaption, vbBinaryCompare) <> 0 Then InsertCustomMacroLauncherReviewFinding batchID, HEALTH_SEVERITY_INFO, "StaleButtonCaption", cfg.ID, cfg.MacroKey, cfg.MacroTitle, cfg.GeneratedMacroObjectName, matched.Name, SharedText_Nz(matched.Caption), expectedCaption, expectedOnClick, SharedText_Nz(matched.OnClick), expectedTooltip, SharedText_Nz(matched.ControlTipText), vbNullString
                If StrComp(SharedText_Nz(matched.OnClick), expectedOnClick, vbBinaryCompare) <> 0 Then InsertCustomMacroLauncherReviewFinding batchID, HEALTH_SEVERITY_WARNING, "StaleButtonAction", cfg.ID, cfg.MacroKey, cfg.MacroTitle, cfg.GeneratedMacroObjectName, matched.Name, SharedText_Nz(matched.Caption), expectedCaption, expectedOnClick, SharedText_Nz(matched.OnClick), expectedTooltip, SharedText_Nz(matched.ControlTipText), vbNullString
                If StrComp(SharedText_Nz(matched.ControlTipText), expectedTooltip, vbBinaryCompare) <> 0 Then InsertCustomMacroLauncherReviewFinding batchID, HEALTH_SEVERITY_INFO, "StaleButtonTooltip", cfg.ID, cfg.MacroKey, cfg.MacroTitle, cfg.GeneratedMacroObjectName, matched.Name, SharedText_Nz(matched.Caption), expectedCaption, expectedOnClick, SharedText_Nz(matched.OnClick), expectedTooltip, SharedText_Nz(matched.ControlTipText), vbNullString
                If StrComp(SharedText_Nz(matched.Tag), LauncherButtonTag(cfg), vbBinaryCompare) <> 0 Then InsertCustomMacroLauncherReviewFinding batchID, HEALTH_SEVERITY_INFO, "StaleButtonTag", cfg.ID, cfg.MacroKey, cfg.MacroTitle, cfg.GeneratedMacroObjectName, matched.Name, SharedText_Nz(matched.Caption), expectedCaption, expectedOnClick, SharedText_Nz(matched.OnClick), expectedTooltip, SharedText_Nz(matched.ControlTipText), vbNullString
            End If
        End If
        rs.MoveNext
    Loop
    For Each ctl In forms(formName).Controls
        PumpSetupResponsiveness
        If ctl.controlType = acCommandButton Then
            tagText = SharedText_Nz(ctl.Tag)
            If (StrComp(Left$(tagText, Len("ADCLauncher" & MODULE_TAG_DELIMITER)), "ADCLauncher" & MODULE_TAG_DELIMITER, vbTextCompare) = 0) Then
                keyText = LauncherButtonIdentityKeyFromTag(tagText)
                If Len(keyText) = 0 Then
                    InsertCustomMacroLauncherReviewFinding batchID, HEALTH_SEVERITY_WARNING, "MalformedTag", 0, vbNullString, vbNullString, vbNullString, ctl.Name, SharedText_Nz(ctl.Caption), vbNullString, vbNullString, SharedText_Nz(ctl.OnClick), vbNullString, SharedText_Nz(ctl.ControlTipText), vbNullString
                Else
                    If seenButtons.Exists(keyText) Then
                        InsertCustomMacroLauncherReviewFinding batchID, HEALTH_SEVERITY_WARNING, "DuplicateButton", CLng(Val(LauncherTagValue(tagText, "ConfigID"))), LauncherTagValue(tagText, "MacroKey"), vbNullString, LauncherTagValue(tagText, "GeneratedMacroObjectName"), ctl.Name, SharedText_Nz(ctl.Caption), vbNullString, vbNullString, SharedText_Nz(ctl.OnClick), vbNullString, SharedText_Nz(ctl.ControlTipText), vbNullString
                    Else
                        seenButtons.Add keyText, ctl.Name
                    End If
                    If TaggedButtonMapsToActiveGeneratedConfig(tagText, formName) = False Then
                        InsertCustomMacroLauncherReviewFinding batchID, HEALTH_SEVERITY_WARNING, "UnmatchedButton", CLng(Val(LauncherTagValue(tagText, "ConfigID"))), LauncherTagValue(tagText, "MacroKey"), vbNullString, LauncherTagValue(tagText, "GeneratedMacroObjectName"), ctl.Name, SharedText_Nz(ctl.Caption), vbNullString, vbNullString, SharedText_Nz(ctl.OnClick), vbNullString, SharedText_Nz(ctl.ControlTipText), vbNullString
                    End If
                End If
            End If
        End If
    Next ctl
ExitHere:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Exit Sub
ErrHandler:
    InsertCustomMacroLauncherReviewFinding batchID, "Error", "ReviewError", 0, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, "Review failed: " & Err.Description
    Resume ExitHere
End Sub

Private Sub CreateOrReplaceOpenFormMacro(ByRef cfg As TMacroRecord)
    On Error GoTo ErrHandler
    Dim macroText As String
    EnsureValidAccessMacroObjectName cfg.MacroTitle, "CreateOrReplaceOpenFormMacro"
    macroText = OpenFormMacroText(cfg.TargetObjectName)
    CreateOrReplaceMacroFromText cfg.MacroTitle, macroText, cfg.navigationPaneGroupName, "CreateOrReplaceOpenFormMacro", ConfigObjectDescription(cfg)
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "CreateOrReplaceOpenFormMacro", "Could not create Access OpenForm macro." & vbCrLf & "MacroKey: " & cfg.MacroKey & vbCrLf & "MacroTitle: " & cfg.MacroTitle & vbCrLf & "TargetObjectName: " & cfg.TargetObjectName & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

' Purpose: Resolves and returns escape macro text from the supplied inputs and the authoritative ADC definitions/state used by this module.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   value (Input; String): Input value for escape macro text; is passed to `Replace`.
' Returns: String: Returns the supplied text escaped/quoted for the escape macro text context so it can be embedded safely in the generated expression/macro text.
Public Function EscapeMacroText(ByVal value As String) As String
    'Escapes embedded double quotes for Access macro text by doubling them.
    EscapeMacroText = Replace(value, """", """""")
End Function

Private Function OpenFormMacroText(ByVal formName As String) As String
    OpenFormMacroText = "Version =196611" & vbCrLf & "ColumnsShown =0" & vbCrLf & "Begin" & vbCrLf & "    Action =""OpenForm""" & vbCrLf & "    Argument =""" & EscapeMacroText(formName) & """" & vbCrLf & "    Argument =""0""" & vbCrLf & "    Argument =""""" & vbCrLf & "    Argument =""""" & vbCrLf & "    Argument =""""" & vbCrLf & "    Argument =""0""" & vbCrLf & "End" & vbCrLf
End Function

' Purpose: Disables custom macro launcher automation, coordinating `CoreBeginOperationProgressIfNeeded` and `CoreUpdateSetupProgress` as the implementation requires.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub DisableCustomMacroLauncherAutomation(Optional ByVal showMessage As Boolean = True)
    On Error GoTo ErrHandler
    Dim progressStarted As Boolean
    progressStarted = CoreBeginOperationProgressIfNeeded("Disable " & DEFAULT_UTILITY_SHORT_NAME & " Custom Macro Launcher automation", 2)
    CoreUpdateSetupProgress "Verify launcher table readiness and foundation"
    RequireConfigTableForUtilityWorkflow "DisableCustomMacroLauncherAutomation"
    RepairLauncherFoundation False
    CoreUpdateSetupProgress "Disable Custom Macro Launcher automation setting"
    SetSettingRawText SETTING_ENABLE_CUSTOM_MACRO_LAUNCHER_AUTOMATION, "False"
    ReportCoreActionSuccess "DisableCustomMacroLauncherAutomation", "CustomMacroLauncher", "Custom Macro Launcher automation disabled successfully.", "Custom Macro Launcher automation has been disabled." & vbCrLf & vbCrLf & "Existing launcher buttons were not deleted." & vbCrLf & "Accepted pre-alpha warning fields were not cleared.", "Disable Custom Macro Launcher automation", SETTING_ENABLE_CUSTOM_MACRO_LAUNCHER_AUTOMATION, "Setting", "CodeDb", "CustomMacroLauncher", showMessage, "Disable Custom Macro Launcher Automation"
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
    If showMessage Then
        CoreShowErrorMessage "Could not disable Custom Macro Launcher automation." & vbCrLf & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, vbCritical, "Disable Custom Macro Launcher Automation", "DisableCustomMacroLauncherAutomation", "Review"
        Exit Sub
    End If
    Err.Raise savedErrNumber, "DisableCustomMacroLauncherAutomation", "Could not disable Custom Macro Launcher automation." & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Function LauncherAutomationActive() As Boolean
    On Error GoTo ErrHandler
    LauncherAutomationActive = False
    If SettingBool(SETTING_ENABLE_CUSTOM_MACRO_LAUNCHER_AUTOMATION, True) = False Then Exit Function
    If SettingBool(SETTING_CUSTOM_MACRO_LAUNCHER_WARNING_ACCEPTED, True) = False Then Exit Function
    If SettingBool(SETTING_ADD_GENERATED_MACROS_TO_CUSTOM_MACRO_LAUNCHER, True) = False Then Exit Function
    LauncherAutomationActive = True
    Exit Function
ErrHandler:
    LauncherAutomationActive = False
End Function

' Purpose: Enables custom macro launcher automation, coordinating `CoreBeginOperationProgressIfNeeded` and `CoreUpdateSetupProgress` as the implementation requires.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the underlying
'   operation.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub EnableCustomMacroLauncherAutomation(ByVal showMessage As Boolean)
    On Error GoTo ErrHandler
    Dim typedText As String
    Dim progressStarted As Boolean
    progressStarted = CoreBeginOperationProgressIfNeeded("Enable " & DEFAULT_UTILITY_SHORT_NAME & " Custom Macro Launcher automation", 2)
    CoreUpdateSetupProgress "Verify launcher table readiness and foundation"
    RequireConfigTableForUtilityWorkflow "EnableCustomMacroLauncherAutomation"
    RepairLauncherFoundation False
    CoreUpdateSetupProgress "Prompt for Custom Macro Launcher automation confirmation"
    typedText = InputBox(LauncherPreAlphaWarningText(), "Enable Custom Macro Launcher Automation")
    If StrComp(Trim$(typedText), "I Understand", vbBinaryCompare) = 0 Then
        SetSettingRawText SETTING_ENABLE_CUSTOM_MACRO_LAUNCHER_AUTOMATION, "True"
        SetSettingRawText SETTING_CUSTOM_MACRO_LAUNCHER_WARNING_ACCEPTED, "True"
        SetSettingRawText SETTING_CUSTOM_MACRO_LAUNCHER_WARNING_ACCEPTED_ON, Format$(Now, "yyyy-mm-dd hh:nn:ss")
        ReportCoreActionSuccess "EnableCustomMacroLauncherAutomation", "CustomMacroLauncher", "Custom Macro Launcher automation enabled successfully.", "Custom Macro Launcher automation has been enabled." & vbCrLf & vbCrLf & "This feature is pre-alpha and intended only for your own development database.", "Enable Custom Macro Launcher automation", SETTING_ENABLE_CUSTOM_MACRO_LAUNCHER_AUTOMATION, "Setting", "CodeDb", "CustomMacroLauncher", showMessage, "Enable Custom Macro Launcher Automation"
        If progressStarted Then CoreCompleteSetupProgress
    Else
        SetSettingRawText SETTING_ENABLE_CUSTOM_MACRO_LAUNCHER_AUTOMATION, "False"
        If progressStarted Then CoreCompleteSetupProgress
        CorePulseProgress "Custom Macro Launcher automation was not enabled. Typed confirmation did not match."
        CoreAddOperationNote "Warning: Custom Macro Launcher automation was not enabled because the typed confirmation did not exactly match."
        LogTraceMessage "Warning", "CustomMacroLauncher", "EnableCustomMacroLauncherAutomation", "Custom Macro Launcher automation was not enabled.", "Typed confirmation did not exactly match: I Understand", "Enable Custom Macro Launcher automation", 0, vbNullString, vbNullString, SETTING_ENABLE_CUSTOM_MACRO_LAUNCHER_AUTOMATION, "Setting", "CodeDb", "CustomMacroLauncher"
        If showMessage Then CoreShowMessage "Custom Macro Launcher automation was not enabled." & vbCrLf & vbCrLf & "The typed confirmation did not exactly match: I Understand", vbExclamation, "Enable Custom Macro Launcher Automation"
    End If
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If progressStarted Then CoreFailSetupProgress
    If showMessage Then
        CoreShowErrorMessage "Could not enable Custom Macro Launcher automation." & vbCrLf & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, vbCritical, "Enable Custom Macro Launcher Automation", "EnableCustomMacroLauncherAutomation", "Review"
        Exit Sub
    End If
    Err.Raise savedErrNumber, "EnableCustomMacroLauncherAutomation", "Could not enable Custom Macro Launcher automation." & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

' Purpose: Builds repair custom macro launchers.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the underlying
'   operation.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub BuildRepairCustomMacroLaunchers(ByVal showMessage As Boolean)
    On Error GoTo ErrHandler
    Dim formName As String
    Dim saveResponse As VbMsgBoxResult
    Dim roleItem As Variant
    Dim preparedForms As Collection
    Dim progressStarted As Boolean
    ' INTENTIONAL FILTERED PARTITION: preparedForms is a consumer-required subset/projection derived from the source enumerated by BuildRepairCustomMacroLaunchers; its membership is not independently authoritative.
    Set preparedForms = New Collection
    progressStarted = CoreBeginOperationProgressIfNeeded("Prepare " & DEFAULT_UTILITY_SHORT_NAME & " Custom Macro Launcher", 1)
    CoreResetOperationNotes
    CoreUpdateSetupProgress "Verify launcher table readiness and foundation"
    RequireConfigTableForUtilityWorkflow "BuildRepairCustomMacroLaunchers"
    RepairLauncherFoundation False
    For Each roleItem In LauncherDatabaseRoles()
        PumpSetupResponsiveness
        formName = LauncherFormNameForDatabaseRole(CStr(roleItem))
        CorePulseProgress "Build/Repair Custom Macro Launchers: " & formName
        If EnsureLauncherFormReady(formName) = False Then
            Err.Raise vbObjectError + 5901, "BuildRepairCustomMacroLaunchers", "Launcher form could not be prepared." & vbCrLf & "Form: " & formName
        End If
        If OpenLauncherDesignForm(formName, True) = False Then
            Err.Raise vbObjectError + 5902, "BuildRepairCustomMacroLaunchers", "Launcher form is open and was not closed for design updates." & vbCrLf & "Form: " & formName
        End If
        EnsureCustomMacroLauncherAndAnchorCore formName
        saveResponse = CoreShowMessageResult("Custom Macro Launcher build/repair changes are ready." & vbCrLf & vbCrLf & "Role: " & LauncherRoleLabel(CStr(roleItem)) & vbCrLf & "Form: " & formName & vbCrLf & vbCrLf & "Save changes to " & formName & "?", vbQuestion Or vbYesNo Or vbDefaultButton1, "Build/Repair Custom Macro Launchers")
        CloseLauncherDesignForm formName, (saveResponse = vbYes)
        preparedForms.Add LauncherRoleLabel(CStr(roleItem)) & ": " & formName
    Next roleItem
    ReportCoreActionSuccess "BuildRepairCustomMacroLaunchers", "CustomMacroLauncher", "Custom Macro Launcher build/repair completed successfully.", "Custom Macro Launcher build/repair complete." & vbCrLf & vbCrLf & "Prepared forms:" & vbCrLf & SharedCollection_Join(preparedForms, vbCrLf) & CoreOperationSummaryForMessage(), "Build/Repair Custom Macro Launchers", formName, "Form", "CodeDb", "CustomMacroLauncher", showMessage, "Build/Repair Custom Macro Launchers"
    If progressStarted Then CoreCompleteSetupProgress
    Exit Sub
ErrHandler:
    Dim handlerErrNumber As Long
    Dim handlerErrSource As String
    Dim handlerErrDescription As String
    handlerErrNumber = Err.Number
    handlerErrSource = Err.Source
    handlerErrDescription = Err.Description
    If progressStarted Then CoreFailSetupProgress
    On Error Resume Next
    If Len(formName) > 0 Then
        If AccessFormIsOpen(formName) Then DoCmd.Close acForm, formName, acSaveNo
    End If
    On Error GoTo 0
    If showMessage Then
        CoreShowErrorMessage "Could not build/repair Custom Macro Launchers." & vbCrLf & vbCrLf & "Err.Number: " & CStr(handlerErrNumber) & vbCrLf & "Err.Source: " & handlerErrSource & vbCrLf & "Err.Description: " & handlerErrDescription, vbCritical, "Build/Repair Custom Macro Launchers", "BuildRepairCustomMacroLaunchers", "Review"
        Exit Sub
    End If
    Err.Raise handlerErrNumber, "BuildRepairCustomMacroLaunchers", "Could not build/repair Custom Macro Launchers." & vbCrLf & "Form: " & formName & vbCrLf & "Err.Description: " & handlerErrDescription
End Sub

Private Function LauncherRoleLabel(ByVal ConfigOwnerDatabaseRole As String) As String
    If StrComp(NormalizedConfigOwnerDatabaseRole(ConfigOwnerDatabaseRole), CONFIG_OWNER_DATABASE_ROLE_CODE_DB, vbTextCompare) = 0 Then
        LauncherRoleLabel = "code database"
    Else
        LauncherRoleLabel = "referencing database"
    End If
End Function

Private Function LauncherAnchorControlName() As String
    LauncherAnchorControlName = SettingText(SETTING_MACRO_LAUNCHER_APPEND_ANCHOR_CONTROL_NAME, DEFAULT_MACRO_LAUNCHER_APPEND_ANCHOR_CONTROL_NAME)
    LauncherAnchorControlName = Trim$(LauncherAnchorControlName)
    If Len(LauncherAnchorControlName) = 0 Then LauncherAnchorControlName = DEFAULT_MACRO_LAUNCHER_APPEND_ANCHOR_CONTROL_NAME
End Function

Private Sub EnsureCustomMacroLauncherAndAnchorCore(ByVal formName As String)
    On Error GoTo ErrHandler
    Dim frm As Form
    Dim ctl As control
    Dim anchorName As String
    Set frm = forms(formName)
    anchorName = LauncherAnchorControlName()
    If CustomMacroLauncherAvailable(formName) = False Then
        On Error Resume Next
        DoCmd.RunCommand acCmdFormHdrFtr
        Err.Clear
        On Error GoTo ErrHandler
    End If
    frm.Section(acFooter).Visible = True
    If frm.Section(acFooter).Height < DEFAULT_MACRO_LAUNCHER_BUTTON_HEIGHT + DEFAULT_MACRO_LAUNCHER_BUTTON_GAP_Y Then
        frm.Section(acFooter).Height = DEFAULT_MACRO_LAUNCHER_BUTTON_HEIGHT + DEFAULT_MACRO_LAUNCHER_BUTTON_GAP_Y
    End If
    If ControlExistsOnOpenForm(formName, anchorName) = False Then
        Set ctl = CreateControl(formName, acLabel, acFooter, , vbNullString, 0, 0, 1800, 240)
        ctl.TextAlign = 1
        ctl.Name = anchorName
        ctl.Caption = "Macro launcher append anchor"
        ctl.Visible = False
        ctl.Tag = "ADCLauncherAppendAnchor"
        CoreAddOperationNote "Created hidden launcher append anchor [" & anchorName & "] on form [" & formName & "]."
    Else
        Set ctl = forms(formName).Controls(anchorName)
        ctl.Visible = False
        ctl.Tag = "ADCLauncherAppendAnchor"
    End If
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "EnsureCustomMacroLauncherAndAnchorCore", "Could not ensure Custom Macro Launcher layout and anchor." & vbCrLf & "Form: " & formName & vbCrLf & "Anchor: " & anchorName & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

Private Function CustomMacroLauncherAvailable(ByVal formName As String) As Boolean
    On Error GoTo ErrHandler
    Dim heightValue As Long
    heightValue = forms(formName).Section(acFooter).Height
    CustomMacroLauncherAvailable = True
    Exit Function
ErrHandler:
    CustomMacroLauncherAvailable = False
End Function

Private Function LauncherConfigRowQualifiesForButton(ByRef cfg As TMacroRecord) As Boolean
    If cfg.ID <= 0 Then Exit Function
    If cfg.Active = False Then Exit Function
    If StrComp(SharedText_Nz(cfg.ConfigurationStatus), CONFIGURATION_STATUS_ACTIVE, vbTextCompare) <> 0 Then Exit Function
    If Len(Trim$(cfg.GeneratedMacroObjectName)) = 0 Then Exit Function
    If StrComp(cfg.MacroTitle, CONFIRM_AUTOEXEC, vbTextCompare) = 0 Then Exit Function
    If StrComp(cfg.MacroKey, CONFIRM_AUTOEXEC, vbTextCompare) = 0 Then Exit Function
    If OwnedSystemMacroTitle(cfg.MacroTitle) Then Exit Function
    If OwnedSystemMacroTitle(cfg.GeneratedMacroObjectName) Then Exit Function
    If AccessObjectExistsInCurrentDatabase(acMacro, cfg.GeneratedMacroObjectName) = False Then Exit Function
    LauncherConfigRowQualifiesForButton = True
End Function

Private Function ApplyLauncherButtonFromConfig(ByVal formName As String, ByRef cfg As TMacroRecord, ByRef changedKind As String, ByRef conflictText As String) As Boolean
    On Error GoTo ErrHandler
    Dim ctl As control
    Dim buttonName As String
    Dim buttonCaption As String
    Dim expectedOnClick As String
    Dim expectedTag As String
    Dim leftValue As Long
    Dim topValue As Long
    Dim widthValue As Long
    Dim heightValue As Long
    buttonName = LauncherButtonName(cfg.ID)
    buttonCaption = LauncherButtonCaption(cfg)
    expectedOnClick = LauncherExpectedOnClick(cfg.MacroKey)
    expectedTag = LauncherButtonTag(cfg)
    Set ctl = FindLauncherButton(formName, cfg.ID, cfg.MacroKey, conflictText)
    If Len(Trim$(conflictText)) > 0 Then Exit Function
    If ctl Is Nothing Then
        If ControlExistsOnOpenForm(formName, buttonName) Then
            conflictText = "Control name conflict. Existing control [" & buttonName & "] is not a launcher button for ConfigID " & CStr(cfg.ID) & "."
            Exit Function
        End If
        widthValue = LauncherButtonWidth()
        heightValue = LauncherButtonHeight()
        LauncherNextButtonPosition formName, leftValue, topValue
        EnsureCustomMacroLauncherHeight formName, topValue, heightValue
        Set ctl = CreateControl(formName, acCommandButton, acFooter, , vbNullString, leftValue, topValue, widthValue, heightValue)
        ctl.Name = buttonName
        ctl.Caption = buttonCaption
        ctl.OnClick = expectedOnClick
        ctl.ControlTipText = cfg.LauncherTooltip
        ctl.Tag = expectedTag
        ctl.enabled = True
        ctl.Visible = True
        On Error Resume Next
        ctl.Transparent = False
        ctl.TabIndex = LauncherNextTabIndex(formName)
        On Error GoTo ErrHandler
        changedKind = "Added"
    Else
        ctl.Caption = buttonCaption
        ctl.OnClick = expectedOnClick
        ctl.ControlTipText = cfg.LauncherTooltip
        ctl.Tag = expectedTag
        changedKind = "Updated"
    End If
    ApplyLauncherButtonFromConfig = True
    Exit Function
ErrHandler:
    conflictText = "Button add/update failed for ConfigID " & CStr(cfg.ID) & ", MacroKey [" & cfg.MacroKey & "]. Err.Description: " & Err.Description
End Function

Private Function FindLauncherButton(ByVal formName As String, ByVal configID As Long, ByVal MacroKey As String, ByRef conflictText As String) As control
    On Error GoTo ErrHandler
    Dim ctl As control
    Dim tagConfigID As Long
    Dim tagMacroKey As String
    Dim expectedName As String
    expectedName = LauncherButtonName(configID)
    For Each ctl In forms(formName).Controls
        PumpSetupResponsiveness
        If StrComp(ctl.Name, expectedName, vbTextCompare) = 0 Then
            If ControlIsLauncherButton(ctl) Then
                tagConfigID = CLng(Val(LauncherTagValue(SharedText_Nz(ctl.Tag), "ConfigID")))
                If tagConfigID = configID Or tagConfigID = 0 Then
                    Set FindLauncherButton = ctl
                    Exit Function
                End If
            End If
            conflictText = "Control name conflict. [" & expectedName & "] exists but belongs to another control/config row."
            Exit Function
        End If
    Next ctl
    For Each ctl In forms(formName).Controls
        PumpSetupResponsiveness
        If ControlIsLauncherButton(ctl) Then
            tagConfigID = CLng(Val(LauncherTagValue(SharedText_Nz(ctl.Tag), "ConfigID")))
            If tagConfigID = configID Then
                Set FindLauncherButton = ctl
                Exit Function
            End If
        End If
    Next ctl
    If Len(Trim$(MacroKey)) > 0 Then
        For Each ctl In forms(formName).Controls
            PumpSetupResponsiveness
            If ControlIsLauncherButton(ctl) Then
                tagMacroKey = LauncherTagValue(SharedText_Nz(ctl.Tag), "MacroKey")
                If StrComp(tagMacroKey, MacroKey, vbTextCompare) = 0 Then
                    Set FindLauncherButton = ctl
                    Exit Function
                End If
            End If
        Next ctl
    End If
    Exit Function
ErrHandler:
    conflictText = "Could not match launcher button for ConfigID " & CStr(configID) & ". Err.Description: " & Err.Description
End Function

Private Function LauncherButtonName(ByVal configID As Long) As String
    LauncherButtonName = "cmdMRLauncher_" & CStr(configID)
End Function

Private Function LauncherButtonCaption(ByRef cfg As TMacroRecord) As String
    LauncherButtonCaption = Trim$(cfg.MacroTitle)
    If Len(LauncherButtonCaption) = 0 Then LauncherButtonCaption = Trim$(cfg.GeneratedMacroObjectName)
End Function

Private Function LauncherExpectedOnClick(ByVal MacroKey As String) As String
    If SettingBool(SETTING_MACRO_LAUNCHER_USE_CONFIRMED_RUNNER, True) Then
        LauncherExpectedOnClick = "=RunConfirmedConfiguredMacroTarget(" & SharedText_DoubleQuote(MacroKey) & ")"
    Else
        LauncherExpectedOnClick = "=RunConfiguredMacroTarget(" & SharedText_DoubleQuote(MacroKey) & ")"
    End If
End Function

Private Function LauncherButtonTag(ByRef cfg As TMacroRecord) As String
    LauncherButtonTag = "ADCLauncher" & MODULE_TAG_DELIMITER & "ConfigID=" & CStr(cfg.ID) & MODULE_TAG_DELIMITER & "MacroKey=" & cfg.MacroKey & MODULE_TAG_DELIMITER & "ConfigOwnerDatabaseRole=" & cfg.ConfigOwnerDatabaseRole & MODULE_TAG_DELIMITER & "GeneratedMacroObjectName=" & cfg.GeneratedMacroObjectName
End Function

Private Function LauncherTagValue(ByVal tagText As String, ByVal keyName As String) As String
    Dim parts As Variant
    Dim part As Variant
    Dim segmentText As String
    Dim equalPos As Long
    parts = Split(tagText, MODULE_TAG_DELIMITER)
    For Each part In parts
        PumpSetupResponsiveness
        segmentText = CStr(part)
        equalPos = InStr(1, segmentText, "=", vbBinaryCompare)
        If equalPos > 1 Then
            If StrComp(Left$(segmentText, equalPos - 1), keyName, vbTextCompare) = 0 Then
                LauncherTagValue = Mid$(segmentText, equalPos + 1)
                Exit Function
            End If
        End If
    Next part
End Function

Private Sub LauncherNextButtonPosition(ByVal formName As String, ByRef leftValue As Long, ByRef topValue As Long)
    On Error GoTo ErrHandler
    Dim anchor As control
    Dim ctl As control
    Dim taggedCount As Long
    Dim buttonsPerRow As Long
    Dim buttonWidth As Long
    Dim buttonHeight As Long
    Dim gapX As Long
    Dim gapY As Long
    Set anchor = forms(formName).Controls(LauncherAnchorControlName())
    For Each ctl In forms(formName).Controls
        PumpSetupResponsiveness
        If ControlIsLauncherButton(ctl) Then taggedCount = taggedCount + 1
    Next ctl
    buttonsPerRow = LauncherButtonsPerRow()
    buttonWidth = LauncherButtonWidth()
    buttonHeight = LauncherButtonHeight()
    gapX = LauncherButtonGapX()
    gapY = LauncherButtonGapY()
    leftValue = anchor.Left + ((taggedCount Mod buttonsPerRow) * (buttonWidth + gapX))
    topValue = anchor.Top + ((taggedCount \ buttonsPerRow) * (buttonHeight + gapY))
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "LauncherNextButtonPosition", "Could not determine the next Custom Macro Launcher button position." & vbCrLf & "Form: " & formName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Sub

Private Function LauncherNextTabIndex(ByVal formName As String) As Long
    On Error GoTo ErrHandler
    Dim ctl As control
    Dim maxTabIndex As Long
    For Each ctl In forms(formName).Controls
        PumpSetupResponsiveness
        If ControlIsLauncherButton(ctl) Then
            If ctl.TabIndex >= maxTabIndex Then maxTabIndex = ctl.TabIndex + 1
        End If
    Next ctl
    LauncherNextTabIndex = maxTabIndex
    Exit Function
ErrHandler:
    LauncherNextTabIndex = 0
End Function

Private Sub EnsureCustomMacroLauncherHeight(ByVal formName As String, ByVal topValue As Long, ByVal heightValue As Long)
    On Error GoTo ErrHandler
    Dim requiredHeight As Long
    requiredHeight = topValue + heightValue + LauncherButtonGapY()
    If forms(formName).Section(acFooter).Height < requiredHeight Then
        forms(formName).Section(acFooter).Height = requiredHeight
    End If
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "EnsureCustomMacroLauncherHeight", "Could not ensure the Custom Macro Launcher was tall enough for the generated button." & vbCrLf & "Form: " & formName & vbCrLf & "Required footer height: " & CStr(requiredHeight) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Sub

Private Function LauncherButtonWidth() As Long
    LauncherButtonWidth = SettingLong(SETTING_MACRO_LAUNCHER_BUTTON_WIDTH, DEFAULT_MACRO_LAUNCHER_BUTTON_WIDTH)
    If LauncherButtonWidth < 100 Then
        Err.Raise vbObjectError + 5415, "LauncherButtonWidth", "SETTING_MACRO_LAUNCHER_BUTTON_WIDTH setting is below the allowed minimum." & vbCrLf & "SettingKey: " & SETTING_MACRO_LAUNCHER_BUTTON_WIDTH & vbCrLf & "Stored value: " & CStr(LauncherButtonWidth) & vbCrLf & "Minimum: " & CStr(100)
    End If
End Function

Private Function LauncherButtonHeight() As Long
    LauncherButtonHeight = SettingLong(SETTING_MACRO_LAUNCHER_BUTTON_HEIGHT, DEFAULT_MACRO_LAUNCHER_BUTTON_HEIGHT)
    If LauncherButtonHeight < 100 Then
        Err.Raise vbObjectError + 5415, "LauncherButtonHeight", "SETTING_MACRO_LAUNCHER_BUTTON_HEIGHT setting is below the allowed minimum." & vbCrLf & "SettingKey: " & SETTING_MACRO_LAUNCHER_BUTTON_HEIGHT & vbCrLf & "Stored value: " & CStr(LauncherButtonHeight) & vbCrLf & "Minimum: " & CStr(100)
    End If
End Function

Private Function LauncherButtonGapX() As Long
    LauncherButtonGapX = SettingLong(SETTING_MACRO_LAUNCHER_BUTTON_GAP_X, DEFAULT_MACRO_LAUNCHER_BUTTON_GAP_X)
    If LauncherButtonGapX < 0 Then
        Err.Raise vbObjectError + 5415, "LauncherButtonGapX", "SETTING_MACRO_LAUNCHER_BUTTON_GAP_X setting is below the allowed minimum." & vbCrLf & "SettingKey: " & SETTING_MACRO_LAUNCHER_BUTTON_GAP_X & vbCrLf & "Stored value: " & CStr(LauncherButtonGapX) & vbCrLf & "Minimum: " & CStr(0)
    End If
End Function

Private Function LauncherButtonGapY() As Long
    LauncherButtonGapY = SettingLong(SETTING_MACRO_LAUNCHER_BUTTON_GAP_Y, DEFAULT_MACRO_LAUNCHER_BUTTON_GAP_Y)
    If LauncherButtonGapY < 0 Then
        Err.Raise vbObjectError + 5415, "LauncherButtonGapY", "SETTING_MACRO_LAUNCHER_BUTTON_GAP_Y setting is below the allowed minimum." & vbCrLf & "SettingKey: " & SETTING_MACRO_LAUNCHER_BUTTON_GAP_Y & vbCrLf & "Stored value: " & CStr(LauncherButtonGapY) & vbCrLf & "Minimum: " & CStr(0)
    End If
End Function

Private Function LauncherButtonsPerRow() As Long
    LauncherButtonsPerRow = SettingLong(SETTING_MACRO_LAUNCHER_BUTTONS_PER_ROW, DEFAULT_MACRO_LAUNCHER_BUTTONS_PER_ROW)
    If LauncherButtonsPerRow < 1 Then
        Err.Raise vbObjectError + 5415, "LauncherButtonsPerRow", "SETTING_MACRO_LAUNCHER_BUTTONS_PER_ROW setting is below the allowed minimum." & vbCrLf & "SettingKey: " & SETTING_MACRO_LAUNCHER_BUTTONS_PER_ROW & vbCrLf & "Stored value: " & CStr(LauncherButtonsPerRow) & vbCrLf & "Minimum: " & CStr(1)
    End If
End Function

Private Sub BeginLauncherGenerationTracking(ByVal firstRowOnly As Boolean)
    On Error GoTo ErrHandler
    Dim automationExpected As Boolean
    ResolveGeneratedMacroRollbackRecovery "starting launcher generation tracking"
    mLauncherGenerationTrackingActive = False
    Set mLauncherGeneratedConfigIDs = Nothing
    Set mGeneratedMacroRollbackEntries = Nothing
    Set mLauncherFormRollbackEntries = Nothing
    mGeneratedMacroRollbackRecoveryRequired = False
    If firstRowOnly Then Exit Sub
    automationExpected = LauncherAutomationActive()
    If automationExpected = False Then Exit Sub
    Set mLauncherGeneratedConfigIDs = New Collection
    Set mGeneratedMacroRollbackEntries = New Collection
    RepairLauncherFoundation False
    mLauncherGenerationTrackingActive = True
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    mLauncherGenerationTrackingActive = False
    If GeneratedMacroRollbackEntryCount() > 0 Then mGeneratedMacroRollbackRecoveryRequired = True
    LogTraceMessage "Error", "CustomMacroLauncher", "BeginLauncherGenerationTracking", "Custom Macro Launcher generation tracking could not start.", "Automation expected: " & CStr(automationExpected) & vbCrLf & "First row only: " & CStr(firstRowOnly) & vbCrLf & "Retained rollback entries: " & CStr(GeneratedMacroRollbackEntryCount()) & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, "Start Custom Macro Launcher tracking", savedErrNumber, savedErrSource, savedErrDescription, vbNullString, "CustomMacroLauncher", "CodeDb", "CustomMacroLauncher"
    If automationExpected Or GeneratedMacroRollbackEntryCount() > 0 Then Err.Raise savedErrNumber, "BeginLauncherGenerationTracking", "Custom Macro Launcher generation tracking could not start safely." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription & vbCrLf & GeneratedMacroRollbackPathsForMessage()
End Sub

Private Sub TrackLauncherGeneratedConfig(ByRef cfg As TMacroRecord)
    On Error Resume Next
    If mLauncherGenerationTrackingActive = False Then Exit Sub
    If mLauncherGeneratedConfigIDs Is Nothing Then Set mLauncherGeneratedConfigIDs = New Collection
    If cfg.ID <= 0 Then Exit Sub
    If StrComp(cfg.MacroTitle, CONFIRM_AUTOEXEC, vbTextCompare) = 0 Then Exit Sub
    If StrComp(cfg.MacroKey, CONFIRM_AUTOEXEC, vbTextCompare) = 0 Then Exit Sub
    If OwnedSystemMacroTitle(cfg.MacroTitle) Then Exit Sub
    mLauncherGeneratedConfigIDs.Add CStr(cfg.ID), CStr(cfg.ID)
End Sub

Private Function LauncherButtonIdentityKeyFromTag(ByVal tagText As String) As String
    Dim configID As Long
    Dim MacroKey As String
    configID = CLng(Val(LauncherTagValue(tagText, "ConfigID")))
    MacroKey = LauncherTagValue(tagText, "MacroKey")
    If configID > 0 Then
        LauncherButtonIdentityKeyFromTag = "ID=" & CStr(configID)
    ElseIf Len(Trim$(MacroKey)) > 0 Then
        LauncherButtonIdentityKeyFromTag = "KEY=" & LCase$(Trim$(MacroKey))
    End If
End Function

Private Sub InsertCustomMacroLauncherReviewFinding(ByVal batchID As String, ByVal severityText As String, ByVal issueTypeText As String, ByVal configID As Long, ByVal MacroKey As String, ByVal MacroTitle As String, ByVal GeneratedMacroObjectName As String, ByVal controlName As String, ByVal controlCaption As String, ByVal expectedCaption As String, ByVal expectedOnClick As String, ByVal actualOnClick As String, ByVal expectedTooltip As String, ByVal actualTooltip As String, ByVal findingNotes As String)
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Set rs = ADCDb().OpenRecordset(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LauncherControlsReview"), dbOpenDynaset)
    rs.AddNew
    SetRecordsetFieldIfExists rs, FIELD_REVIEW_BATCH_ID, batchID
    SetRecordsetFieldIfExists rs, FIELD_REVIEWED_ON, Now
    SetRecordsetFieldIfExists rs, FIELD_SEVERITY, severityText
    SetRecordsetFieldIfExists rs, FIELD_ISSUE_TYPE, issueTypeText
    SetRecordsetFieldIfExists rs, FIELD_CONFIG_ID, configID
    SetRecordsetFieldIfExists rs, FIELD_MACRO_KEY, MacroKey
    SetRecordsetFieldIfExists rs, FIELD_MACRO_TITLE, MacroTitle
    SetRecordsetFieldIfExists rs, FIELD_GENERATED_MACRO_OBJECT_NAME, GeneratedMacroObjectName
    SetRecordsetFieldIfExists rs, FIELD_CONTROL_NAME, controlName
    SetRecordsetFieldIfExists rs, FIELD_CONTROL_CAPTION, controlCaption
    SetRecordsetFieldIfExists rs, FIELD_EXPECTED_CAPTION, expectedCaption
    SetRecordsetFieldIfExists rs, FIELD_EXPECTED_ON_CLICK, expectedOnClick
    SetRecordsetFieldIfExists rs, FIELD_ACTUAL_ON_CLICK, actualOnClick
    SetRecordsetFieldIfExists rs, FIELD_EXPECTED_TOOLTIP, expectedTooltip
    SetRecordsetFieldIfExists rs, FIELD_ACTUAL_TOOLTIP, actualTooltip
    SetRecordsetFieldIfExists rs, FIELD_FINDING_NOTES, findingNotes
    rs.Update
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
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
    Err.Raise savedErrNumber, "InsertCustomMacroLauncherReviewFinding", "Could not save a Custom Macro Launcher review finding." & vbCrLf & "ReviewBatchID: " & batchID & vbCrLf & "IssueType: " & issueTypeText & vbCrLf & "ConfigID: " & CStr(configID) & vbCrLf & "ControlName: " & controlName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Function LauncherKeeperNameFromRank(ByVal rankText As String) As String
    Dim parts As Variant
    parts = Split(rankText, MODULE_TAG_DELIMITER)
    If UBound(parts) >= 2 Then LauncherKeeperNameFromRank = CStr(parts(2))
End Function

Private Function LauncherPreAlphaWarningText() As String
    LauncherPreAlphaWarningText = "PRE-ALPHA FEATURE WARNING" & vbCrLf & vbCrLf & "Macro Custom Macro Launcher automation is not release ready or tested." & vbCrLf & vbCrLf & "This feature opens and modifies the design of the Access Development Console Macros form." & vbCrLf & "It can add, update, delete, resize, and save controls in the form footer." & vbCrLf & vbCrLf & "Use this only in your own development database." & vbCrLf & vbCrLf & "Do not enable this in a released/shared utility build unless it has been reviewed and tested." & vbCrLf & vbCrLf & "To enable this feature, type:" & vbCrLf & vbCrLf & "I Understand"
End Function

Private Function CustomMacroLauncherReviewSummaryText(ByVal batchID As String) As String
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim summaryText As String
    sqlText = "SELECT " & SQL_FIELD_ISSUE_TYPE & ", Count(*) AS IssueCount FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LauncherControlsReview") & "] WHERE " & SQL_FIELD_REVIEW_BATCH_ID & "=" & SharedDao_SqlQuote(batchID) & " GROUP BY " & SQL_FIELD_ISSUE_TYPE & " ORDER BY " & SQL_FIELD_ISSUE_TYPE & ";"
    Set rs = ADCDb().OpenRecordset(sqlText, dbOpenSnapshot)
    If rs.EOF Then
        CustomMacroLauncherReviewSummaryText = "No Custom Macro Launcher issues were found."
    Else
        summaryText = "Issue counts by type:"
        Do While Not rs.EOF
            PumpSetupResponsiveness
            summaryText = summaryText & vbCrLf & " - " & SharedText_Nz(rs!issueType) & ": " & CStr(CLng(rs!issueCount))
            rs.MoveNext
        Loop
        CustomMacroLauncherReviewSummaryText = summaryText
    End If
ExitHere:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Exit Function
ErrHandler:
    CustomMacroLauncherReviewSummaryText = "Review summary could not be built. Err.Description: " & Err.Description
    Resume ExitHere
End Function

Private Sub WriteTextFile(ByVal filePath As String, ByVal contents As String)
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, contents
    Close #fileNum
End Sub

' Purpose: Builds update custom macro launcher buttons.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the underlying
'   operation.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub BuildUpdateCustomMacroLauncherButtons(ByVal showMessage As Boolean)
    On Error GoTo ErrHandler
    Dim formName As String
    Dim addCount As Long
    Dim updateCount As Long
    Dim totalAddCount As Long
    Dim totalUpdateCount As Long
    Dim failureText As String
    Dim saveResponse As VbMsgBoxResult
    Dim enableResponse As VbMsgBoxResult
    Dim sqlText As String
    Dim roleItem As Variant
    Dim summaryLines As Collection
    Dim progressStarted As Boolean
    ' INTENTIONAL ORDERED VIEW: summaryLines preserves the deterministic sequence required by BuildUpdateCustomMacroLauncherButtons; keyed-only or unordered storage would not preserve the consumer semantics.
    Set summaryLines = New Collection
    progressStarted = CoreBeginOperationProgressIfNeeded("Backfill " & DEFAULT_UTILITY_SHORT_NAME & " Custom Macro Launcher Buttons", 1)
    CoreResetOperationNotes
    CoreUpdateSetupProgress "Verify launcher table readiness and foundation"
    RequireConfigTableForUtilityWorkflow "BuildUpdateCustomMacroLauncherButtons"
    RepairLauncherFoundation False
    If LauncherAutomationActive() = False Then
        enableResponse = CoreShowMessageResult("Custom Macro Launcher automation is disabled." & vbCrLf & vbCrLf & "Enable it now before backfill?" & vbCrLf & vbCrLf & "Default answer is No.", vbQuestion Or vbYesNo Or vbDefaultButton2, "Build/Update Custom Macro Launcher Buttons")
        If enableResponse <> vbYes Then
            CoreAddOperationNote "Warning: Custom Macro Launcher Button build/update skipped because Custom Macro Launcher automation was not enabled."
            CorePulseProgress "Custom Macro Launcher Button build/update skipped. Automation was not enabled."
            LogTraceMessage "Warning", "CustomMacroLauncher", "BuildUpdateCustomMacroLauncherButtons", "Custom Macro Launcher Button build/update skipped.", "Custom Macro Launcher automation was disabled and the user did not enable it.", "Build/Update Custom Macro Launcher Buttons", 0, vbNullString, vbNullString, SETTING_ENABLE_CUSTOM_MACRO_LAUNCHER_AUTOMATION, "Setting", "CodeDb", "CustomMacroLauncher"
            If progressStarted Then CoreCompleteSetupProgress
            Exit Sub
        End If
        EnableCustomMacroLauncherAutomation False
        If LauncherAutomationActive() = False Then
            CoreAddOperationNote "Warning: Custom Macro Launcher Button build/update skipped because Custom Macro Launcher automation remained disabled after the enable prompt."
            CorePulseProgress "Custom Macro Launcher Button build/update skipped. Automation remained disabled."
            LogTraceMessage "Warning", "CustomMacroLauncher", "BuildUpdateCustomMacroLauncherButtons", "Custom Macro Launcher Button build/update skipped.", "Custom Macro Launcher automation remained disabled after the enable prompt.", "Build/Update Custom Macro Launcher Buttons", 0, vbNullString, vbNullString, SETTING_ENABLE_CUSTOM_MACRO_LAUNCHER_AUTOMATION, "Setting", "CodeDb", "CustomMacroLauncher"
            If progressStarted Then CoreCompleteSetupProgress
            Exit Sub
        End If
    End If
    For Each roleItem In LauncherDatabaseRoles()
        PumpSetupResponsiveness
        CorePulseProgress "Check Custom Macro Launcher Button build/update rows for " & LauncherRoleLabel(CStr(roleItem))
        sqlText = LauncherBackfillSql(CStr(roleItem))
        If SqlReturnsRows(sqlText) Then
            addCount = 0
            updateCount = 0
            failureText = vbNullString
            formName = LauncherFormNameForDatabaseRole(CStr(roleItem))
            CorePulseProgress "Build/Update Custom Macro Launcher Buttons: " & formName
            If EnsureLauncherFormReady(formName) = False Then
                Err.Raise vbObjectError + 5903, "BuildUpdateCustomMacroLauncherButtons", "Launcher form could not be prepared." & vbCrLf & "Form: " & formName
            End If
            If OpenLauncherDesignForm(formName, True) = False Then
                Err.Raise vbObjectError + 5904, "BuildUpdateCustomMacroLauncherButtons", "Launcher form is open and was not closed for design updates." & vbCrLf & "Form: " & formName
            End If
            EnsureCustomMacroLauncherAndAnchorCore formName
            ApplyLauncherButtonsFromSql formName, sqlText, addCount, updateCount, failureText
            If Len(Trim$(failureText)) > 0 Then
                Err.Raise vbObjectError + 5905, "BuildUpdateCustomMacroLauncherButtons", failureText
            End If
            saveResponse = CoreShowMessageResult("Custom Macro Launcher changes are ready." & vbCrLf & vbCrLf & "Role: " & LauncherRoleLabel(CStr(roleItem)) & vbCrLf & "Form: " & formName & vbCrLf & "Buttons added: " & CStr(addCount) & vbCrLf & "Buttons updated: " & CStr(updateCount) & vbCrLf & vbCrLf & "Save changes to " & formName & "?", vbQuestion Or vbYesNo Or vbDefaultButton1, "Build/Update Custom Macro Launcher Buttons")
            CloseLauncherDesignForm formName, (saveResponse = vbYes)
            totalAddCount = totalAddCount + addCount
            totalUpdateCount = totalUpdateCount + updateCount
            summaryLines.Add LauncherRoleLabel(CStr(roleItem)) & " [" & formName & "]: added " & CStr(addCount) & ", updated " & CStr(updateCount)
        End If
    Next roleItem
    If summaryLines.count = 0 Then summaryLines.Add "No active generated config rows qualified for launcher backfill."
    ReportCoreActionSuccess "BuildUpdateCustomMacroLauncherButtons", "CustomMacroLauncher", "Custom Macro Launcher Button build/update completed successfully.", "Custom Macro Launcher Button build/update complete." & vbCrLf & vbCrLf & "Buttons added: " & CStr(totalAddCount) & vbCrLf & "Buttons updated: " & CStr(totalUpdateCount) & vbCrLf & vbCrLf & SharedCollection_Join(summaryLines, vbCrLf) & CoreOperationSummaryForMessage(), "Build/Update Custom Macro Launcher Buttons", formName, "Form", "CodeDb", "CustomMacroLauncher", showMessage, "Build/Update Custom Macro Launcher Buttons"
    If progressStarted Then CoreCompleteSetupProgress
    Exit Sub
ErrHandler:
    Dim handlerErrNumber As Long
    Dim handlerErrSource As String
    Dim handlerErrDescription As String
    handlerErrNumber = Err.Number
    handlerErrSource = Err.Source
    handlerErrDescription = Err.Description
    If progressStarted Then CoreFailSetupProgress
    On Error Resume Next
    If Len(formName) > 0 Then
        If AccessFormIsOpen(formName) Then DoCmd.Close acForm, formName, acSaveNo
    End If
    On Error GoTo 0
    If showMessage Then
        CoreShowErrorMessage "Could not build/update Custom Macro Launcher Buttons." & vbCrLf & vbCrLf & "Err.Number: " & CStr(handlerErrNumber) & vbCrLf & "Err.Source: " & handlerErrSource & vbCrLf & "Err.Description: " & handlerErrDescription, vbCritical, "Build/Update Custom Macro Launcher Buttons", "BuildUpdateCustomMacroLauncherButtons", "Review"
        Exit Sub
    End If
    Err.Raise handlerErrNumber, "BuildUpdateCustomMacroLauncherButtons", "Could not build/update Custom Macro Launcher Buttons." & vbCrLf & "Form: " & formName & vbCrLf & "Err.Description: " & handlerErrDescription
End Sub

' Purpose: Performs the public review custom macro launcher buttons operation using the routine's existing ADC workflow and side effects.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input; Boolean): Controls whether the routine shows its user-facing completion, status, warning, or error message in addition to performing the underlying work.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub ReviewCustomMacroLauncherButtons(ByVal showMessage As Boolean)
    On Error GoTo ErrHandler
    Dim formName As String
    Dim batchID As String
    Dim openedForReview As Boolean
    Dim issueSummary As String
    Dim roleItem As Variant
    Dim progressStarted As Boolean
    progressStarted = CoreBeginOperationProgressIfNeeded("Review " & DEFAULT_UTILITY_SHORT_NAME & " Custom Macro Launcher Buttons", 1)
    CoreUpdateSetupProgress "Verify launcher table readiness and foundation"
    RequireConfigTableForUtilityWorkflow "ReviewCustomMacroLauncherButtons"
    RepairLauncherFoundation False
    batchID = Format$(Now, "yyyymmddhhnnss") & "_" & Format$(Timer * 1000, "000000")
    ADCDb().Execute "DELETE FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LauncherControlsReview") & "]", dbFailOnError
    For Each roleItem In LauncherDatabaseRoles()
        PumpSetupResponsiveness
        formName = LauncherFormNameForDatabaseRole(CStr(roleItem))
        CorePulseProgress "Review Custom Macro Launcher Buttons: " & formName
        openedForReview = False
        If LauncherFormExistsStrict(formName, "reviewing Custom Macro Launcher Buttons") = False Then
            InsertCustomMacroLauncherReviewFinding batchID, "Error", "MissingFooter", 0, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, "Launcher form is missing for " & LauncherRoleLabel(CStr(roleItem)) & ": " & formName
        Else
            If AccessFormIsOpen(formName) = False Then
                DoCmd.OpenForm formName, acDesign, , , , acHidden
                openedForReview = True
            End If
            ReviewCustomMacroLauncherOpenForm formName, batchID, CStr(roleItem)
            If openedForReview Then DoCmd.Close acForm, formName, acSaveNo
            openedForReview = False
        End If
    Next roleItem
    issueSummary = CustomMacroLauncherReviewSummaryText(batchID)
    ReportCoreActionSuccess "ReviewCustomMacroLauncherButtons", "CustomMacroLauncher", "Custom Macro Launcher review completed successfully.", "Custom Macro Launcher review complete." & vbCrLf & vbCrLf & "Review table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LauncherControlsReview") & vbCrLf & "Review query: " & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_LauncherControls_Review") & vbCrLf & vbCrLf & issueSummary, "Review Custom Macro Launcher Buttons", ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LauncherControlsReview"), "Table", "CodeDb", "CustomMacroLauncher", showMessage, "Review Custom Macro Launcher Buttons"
    If progressStarted Then CoreCompleteSetupProgress
    Exit Sub
ErrHandler:
    Dim handlerErrNumber As Long
    Dim handlerErrSource As String
    Dim handlerErrDescription As String
    handlerErrNumber = Err.Number
    handlerErrSource = Err.Source
    handlerErrDescription = Err.Description
    If progressStarted Then CoreFailSetupProgress
    On Error Resume Next
    If openedForReview Then DoCmd.Close acForm, formName, acSaveNo
    On Error GoTo 0
    If showMessage Then
        CoreShowErrorMessage "Could not review Custom Macro Launcher Buttons." & vbCrLf & vbCrLf & "Err.Number: " & CStr(handlerErrNumber) & vbCrLf & "Err.Source: " & handlerErrSource & vbCrLf & "Err.Description: " & handlerErrDescription, vbCritical, "Review Custom Macro Launcher Buttons", "ReviewCustomMacroLauncherButtons", "Review"
        Exit Sub
    End If
    Err.Raise handlerErrNumber, "ReviewCustomMacroLauncherButtons", "Could not review Custom Macro Launcher Buttons." & vbCrLf & "Form: " & formName & vbCrLf & "Err.Description: " & handlerErrDescription
End Sub

' Purpose: Performs the public clean stale custom macro launcher buttons operation using the routine's existing ADC workflow and side effects.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input; Boolean): Controls whether the routine shows its user-facing completion, status, warning, or error message in addition to performing the underlying work.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub CleanStaleCustomMacroLauncherButtons(ByVal showMessage As Boolean)
    On Error GoTo ErrHandler
    Dim formName As String
    Dim staleButtons As Object
    Dim duplicateButtons As Object
    Dim staleCount As Long
    Dim duplicateCount As Long
    Dim deletedCount As Long
    Dim totalStaleCount As Long
    Dim totalDuplicateCount As Long
    Dim totalDeletedCount As Long
    Dim response As VbMsgBoxResult
    Dim saveNeeded As Boolean
    Dim roleItem As Variant
    Dim summaryLines As Collection
    Dim progressStarted As Boolean
    ' INTENTIONAL ORDERED VIEW: summaryLines preserves the deterministic sequence required by CleanStaleCustomMacroLauncherButtons; keyed-only or unordered storage would not preserve the consumer semantics.
    Set summaryLines = New Collection
    progressStarted = CoreBeginOperationProgressIfNeeded("Clean stale " & DEFAULT_UTILITY_SHORT_NAME & " Custom Macro Launcher Buttons", 1)
    CoreUpdateSetupProgress "Verify launcher table readiness and foundation"
    RequireConfigTableForUtilityWorkflow "CleanStaleCustomMacroLauncherButtons"
    RepairLauncherFoundation False
    For Each roleItem In LauncherDatabaseRoles()
        PumpSetupResponsiveness
        formName = LauncherFormNameForDatabaseRole(CStr(roleItem))
        CorePulseProgress "Clean Custom Macro Launcher Buttons: " & formName
        If LauncherFormExistsStrict(formName, "cleaning stale Custom Macro Launcher Buttons") = False Then
            summaryLines.Add LauncherRoleLabel(CStr(roleItem)) & " [" & formName & "]: form missing; no cleanup run."
        Else
            If OpenLauncherDesignForm(formName, True) = False Then
                Err.Raise vbObjectError + 5906, "CleanStaleCustomMacroLauncherButtons", "Launcher form is open and was not closed for design updates." & vbCrLf & "Form: " & formName
            End If
            ' INTENTIONAL KEYED LOOKUP: staleButtons, duplicateButtons provides keyed membership/lookup over source rows or state already read by CleanStaleCustomMacroLauncherButtons; it is a derived access path,
            ' not a second authority.
            Set staleButtons = CreateObject("Scripting.Dictionary")
            Set duplicateButtons = CreateObject("Scripting.Dictionary")
            deletedCount = 0
            saveNeeded = False
            CollectLauncherCleanupTargets formName, staleButtons, duplicateButtons
            staleCount = staleButtons.count
            duplicateCount = duplicateButtons.count
            If staleCount > 0 Then
                response = CoreShowMessageResult("Stale tagged launcher buttons found: " & CStr(staleCount) & vbCrLf & vbCrLf & "Role: " & LauncherRoleLabel(CStr(roleItem)) & vbCrLf & "Form: " & formName & vbCrLf & vbCrLf & "Delete these stale tagged launcher buttons?" & vbCrLf & vbCrLf & "Untagged controls, non-launcher controls, and malformed tagged controls will not be deleted.", vbExclamation Or vbYesNo Or vbDefaultButton2, "Clean Stale Custom Macro Launcher Buttons")
                If response = vbYes Then
                    deletedCount = deletedCount + DeleteLauncherControlsByDictionary(formName, staleButtons)
                    saveNeeded = True
                End If
            End If
            If duplicateCount > 0 Then
                response = CoreShowMessageResult("Duplicate tagged launcher buttons found: " & CStr(duplicateCount) & vbCrLf & vbCrLf & "Role: " & LauncherRoleLabel(CStr(roleItem)) & vbCrLf & "Form: " & formName & vbCrLf & vbCrLf & "Select duplicate for deletion tagged launcher buttons?" & vbCrLf & vbCrLf & "The first top-left matching button is kept.", vbExclamation Or vbYesNo Or vbDefaultButton2, "Clean Duplicate Custom Macro Launcher Buttons")
                If response = vbYes Then
                    deletedCount = deletedCount + DeleteLauncherControlsByDictionary(formName, duplicateButtons)
                    saveNeeded = True
                End If
            End If
            CloseLauncherDesignForm formName, saveNeeded
            totalStaleCount = totalStaleCount + staleCount
            totalDuplicateCount = totalDuplicateCount + duplicateCount
            totalDeletedCount = totalDeletedCount + deletedCount
            summaryLines.Add LauncherRoleLabel(CStr(roleItem)) & " [" & formName & "]: stale " & CStr(staleCount) & ", duplicate " & CStr(duplicateCount) & ", deleted " & CStr(deletedCount)
        End If
    Next roleItem
    ReportCoreActionSuccess "CleanStaleCustomMacroLauncherButtons", "CustomMacroLauncher", "Custom Macro Launcher cleanup completed successfully.", "Custom Macro Launcher cleanup complete." & vbCrLf & vbCrLf & "Stale buttons found: " & CStr(totalStaleCount) & vbCrLf & "Duplicate buttons found: " & CStr(totalDuplicateCount) & vbCrLf & "Buttons deleted: " & CStr(totalDeletedCount) & vbCrLf & vbCrLf & SharedCollection_Join(summaryLines, vbCrLf), "Clean Custom Macro Launcher Buttons", formName, "Form", "CodeDb", "CustomMacroLauncher", showMessage, "Clean Stale Custom Macro Launcher Buttons"
    If progressStarted Then CoreCompleteSetupProgress
    Exit Sub
ErrHandler:
    Dim handlerErrNumber As Long
    Dim handlerErrSource As String
    Dim handlerErrDescription As String
    handlerErrNumber = Err.Number
    handlerErrSource = Err.Source
    handlerErrDescription = Err.Description
    If progressStarted Then CoreFailSetupProgress
    On Error Resume Next
    If Len(formName) > 0 Then
        If AccessFormIsOpen(formName) Then DoCmd.Close acForm, formName, acSaveNo
    End If
    On Error GoTo 0
    If showMessage Then
        CoreShowErrorMessage "Could not clean Custom Macro Launcher Buttons." & vbCrLf & vbCrLf & "Err.Number: " & CStr(handlerErrNumber) & vbCrLf & "Err.Source: " & handlerErrSource & vbCrLf & "Err.Description: " & handlerErrDescription, vbCritical, "Clean Stale Custom Macro Launcher Buttons", "CleanStaleCustomMacroLauncherButtons", "Review"
        Exit Sub
    End If
    Err.Raise handlerErrNumber, "CleanStaleCustomMacroLauncherButtons", "Could not clean Custom Macro Launcher Buttons." & vbCrLf & "Form: " & formName & vbCrLf & "Err.Description: " & handlerErrDescription
End Sub

Private Function LauncherDatabaseRoles() As Variant
    LauncherDatabaseRoles = Array(CONFIG_OWNER_DATABASE_ROLE_CODE_DB, CONFIG_OWNER_DATABASE_ROLE_HOST_DB)
End Function

Private Function ResolvedConfigOwnerDatabaseRoleForConfig(ByRef cfg As TMacroRecord) As String
    ResolvedConfigOwnerDatabaseRoleForConfig = NormalizedConfigOwnerDatabaseRole(cfg.ConfigOwnerDatabaseRole)
End Function

Private Function SqlReturnsRows(ByVal sqlText As String) As Boolean
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Set rs = ADCDb().OpenRecordset(sqlText, dbOpenSnapshot)
    SqlReturnsRows = (rs.EOF = False)
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
    Err.Raise savedErrNumber, "SqlReturnsRows", "Could not determine whether the supplied launcher query returns rows." & vbCrLf & "SQL: " & sqlText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Function

Private Function LauncherBackfillSql(Optional ByVal ConfigOwnerDatabaseRole As String = vbNullString) As String
    Dim sqlText As String
    Dim normalizedRole As String
    normalizedRole = NormalizedConfigOwnerDatabaseRole(ConfigOwnerDatabaseRole)
    sqlText = ""
    sqlText = sqlText & "SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "] "
    sqlText = sqlText & "WHERE " & SQL_FIELD_ACTIVE & "<>0 "
    sqlText = sqlText & "AND Nz(" & SQL_FIELD_CONFIGURATION_STATUS & "," & SharedDao_SqlQuote(CONFIGURATION_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(CONFIGURATION_STATUS_ACTIVE) & " "
    sqlText = sqlText & "AND Len(Nz(" & SQL_FIELD_GENERATED_MACRO_OBJECT_NAME & ",''))>0 "
    sqlText = sqlText & "AND UCase(Nz(" & SQL_FIELD_MACRO_TITLE & ",''))<>'AUTOEXEC' "
    sqlText = sqlText & "AND UCase(Nz(" & SQL_FIELD_MACRO_KEY & ",''))<>'AUTOEXEC' "
    If Len(normalizedRole) > 0 Then
        sqlText = sqlText & "AND [" & FIELD_CONFIG_OWNER_DATABASE_ROLE & "]=" & SharedDao_SqlQuote(normalizedRole) & " "
    End If
    sqlText = sqlText & "ORDER BY Nz(" & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER & ",0), " & SQL_FIELD_MACRO_TITLE & ", " & SQL_FIELD_MACRO_KEY & ", " & SQL_FIELD_ID & ";"
    LauncherBackfillSql = sqlText
End Function

Private Sub ApplyLauncherButtonsFromSql(ByVal formName As String, ByVal sqlText As String, ByRef addCount As Long, ByRef updateCount As Long, ByRef failureText As String)
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim cfg As TMacroRecord
    Dim changedKind As String
    Dim conflictText As String
    Dim reviewedCount As Long
    Set rs = ADCDb().OpenRecordset(sqlText, dbOpenSnapshot)
    Do While Not rs.EOF
        PumpSetupResponsiveness
        reviewedCount = reviewedCount + 1
        PulseProgressEvery reviewedCount, 25, "Custom Macro Launcher config rows processed"
        MacroRecordFromRecordset rs, cfg
        If LauncherConfigRowQualifiesForButton(cfg) Then
            changedKind = vbNullString
            conflictText = vbNullString
            If ApplyLauncherButtonFromConfig(formName, cfg, changedKind, conflictText) Then
                Select Case changedKind
                    Case "Added"
                        addCount = addCount + 1
                    Case "Updated"
                        updateCount = updateCount + 1
                End Select
            End If
            If Len(Trim$(conflictText)) > 0 Then
                failureText = SharedText_AppendNonBlankLine(failureText, conflictText)
            End If
        End If
        rs.MoveNext
    Loop
ExitHere:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Exit Sub
ErrHandler:
    failureText = SharedText_AppendNonBlankLine(failureText, "Launcher button add/update failed. Err.Description: " & Err.Description)
    Resume ExitHere
End Sub

Private Function TrackedLauncherConfigSql(Optional ByVal ConfigOwnerDatabaseRole As String = vbNullString) As String
    Dim item As Variant
    Dim idList As String
    Dim sqlText As String
    Dim normalizedRole As String
    normalizedRole = NormalizedConfigOwnerDatabaseRole(ConfigOwnerDatabaseRole)
    If mLauncherGeneratedConfigIDs Is Nothing Then
        TrackedLauncherConfigSql = "SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "] WHERE 1=0;"
        Exit Function
    End If
    For Each item In mLauncherGeneratedConfigIDs
        PumpSetupResponsiveness
        If Len(idList) > 0 Then idList = idList & ","
        idList = idList & CStr(CLng(Val(CStr(item))))
    Next item
    If Len(idList) = 0 Then
        TrackedLauncherConfigSql = "SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "] WHERE 1=0;"
        Exit Function
    End If
    sqlText = ""
    sqlText = sqlText & "SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "] "
    sqlText = sqlText & "WHERE " & SQL_FIELD_ID & " IN (" & idList & ") "
    sqlText = sqlText & "AND " & SQL_FIELD_LAST_GENERATED_RESULT & "=" & SharedDao_SqlQuote(GENERATED_RESULT_GENERATED) & " "
    If Len(normalizedRole) > 0 Then
        sqlText = sqlText & "AND [" & FIELD_CONFIG_OWNER_DATABASE_ROLE & "]=" & SharedDao_SqlQuote(normalizedRole) & " "
    End If
    sqlText = sqlText & "ORDER BY Nz(" & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER & ",0), " & SQL_FIELD_MACRO_TITLE & ", " & SQL_FIELD_MACRO_KEY & ", " & SQL_FIELD_ID & ";"
    TrackedLauncherConfigSql = sqlText
End Function

Private Function ControlIsLauncherButton(ByVal ctl As control) As Boolean
    On Error GoTo ErrHandler
    If ctl.controlType <> acCommandButton Then Exit Function
    If StrComp(Left$(SharedText_Nz(ctl.Tag), Len("ADCLauncher" & MODULE_TAG_DELIMITER)), "ADCLauncher" & MODULE_TAG_DELIMITER, vbTextCompare) <> 0 Then Exit Function
    ControlIsLauncherButton = True
    Exit Function
ErrHandler:
    ControlIsLauncherButton = False
End Function

' ===============================================================================================================================================================================
' GENERATED MACRO GENERATION / ROLLBACK PROCEDURES
' ===============================================================================================================================================================================

Private Function GeneratedMacroRollbackEntryCount() As Long
    If Not mGeneratedMacroRollbackEntries Is Nothing Then GeneratedMacroRollbackEntryCount = mGeneratedMacroRollbackEntries.count
    If Not mLauncherFormRollbackEntries Is Nothing Then GeneratedMacroRollbackEntryCount = GeneratedMacroRollbackEntryCount + mLauncherFormRollbackEntries.count
End Function

Private Function GeneratedMacroRollbackPathsForMessage() As String
    Dim item As Variant
    Dim entryData As Variant
    Dim objectName As String
    Dim backupPath As String
    If Not mGeneratedMacroRollbackEntries Is Nothing Then
        For Each item In mGeneratedMacroRollbackEntries
            PumpSetupResponsiveness
            entryData = item
            objectName = vbNullString
            backupPath = vbNullString
            If IsArray(entryData) Then
                If UBound(entryData) >= 0 Then objectName = Trim$(CStr(entryData(0)))
                If UBound(entryData) >= 2 Then backupPath = Trim$(CStr(entryData(2)))
            End If
            GeneratedMacroRollbackPathsForMessage = SharedText_AppendNonBlankLine(GeneratedMacroRollbackPathsForMessage, "Macro: " & IIf(Len(objectName) = 0, "(blank)", objectName) & "; Backup: " & IIf(Len(backupPath) = 0, "(none - object was newly created)", backupPath))
        Next item
    End If
    If Not mLauncherFormRollbackEntries Is Nothing Then
        For Each item In mLauncherFormRollbackEntries
            PumpSetupResponsiveness
            entryData = item
            objectName = vbNullString
            backupPath = vbNullString
            If IsArray(entryData) Then
                If UBound(entryData) >= 0 Then objectName = Trim$(CStr(entryData(0)))
                If UBound(entryData) >= 2 Then backupPath = Trim$(CStr(entryData(2)))
            End If
            GeneratedMacroRollbackPathsForMessage = SharedText_AppendNonBlankLine(GeneratedMacroRollbackPathsForMessage, "Form: " & IIf(Len(objectName) = 0, "(blank)", objectName) & "; Backup: " & IIf(Len(backupPath) = 0, "(none - object was newly created)", backupPath))
        Next item
    End If
End Function

Private Sub ResolveGeneratedMacroRollbackRecovery(ByVal operationName As String)
    On Error GoTo ErrHandler
    Dim entryCount As Long
    Dim response As VbMsgBoxResult
    Dim discardResponse As VbMsgBoxResult
    Dim rollbackSummary As String
    Dim cleanupWarnings As String
    entryCount = GeneratedMacroRollbackEntryCount()
    If mGeneratedMacroRollbackRecoveryRequired = False And entryCount = 0 Then Exit Sub
    If PopupMessagesSuppressedForCurrentRun() Then
        response = vbCancel
    Else
        response = CoreShowMessageResult("An earlier generated-macro or launcher-form rollback did not finish cleanly." & vbCrLf & vbCrLf & "Operation requesting a new boundary: " & operationName & vbCrLf & "Retained rollback entries: " & CStr(entryCount) & vbCrLf & vbCrLf & GeneratedMacroRollbackPathsForMessage() & vbCrLf & vbCrLf & "Yes = retry verified rollback now" & vbCrLf & "No = explicitly accept the current macros and launcher forms" & vbCrLf & "Cancel = leave recovery state unchanged and stop", vbExclamation Or vbYesNoCancel Or vbDefaultButton3, "Access Development Console Generated-Object Recovery")
    End If
    If response = vbYes Then
        rollbackSummary = RollbackGeneratedOperationFromRun()
        If RollbackSummaryHasFailures(rollbackSummary) Then
            mGeneratedMacroRollbackRecoveryRequired = True
            Err.Raise vbObjectError + 5930, "ResolveGeneratedMacroRollbackRecovery", "The retained generated-object rollback still has failures." & vbCrLf & rollbackSummary & vbCrLf & GeneratedMacroRollbackPathsForMessage()
        End If
        cleanupWarnings = CleanupGeneratedRollbackFiles()
        If Len(cleanupWarnings) > 0 Then CoreAddOperationNote cleanupWarnings
        Set mGeneratedMacroRollbackEntries = Nothing
        Set mLauncherFormRollbackEntries = Nothing
        mGeneratedMacroRollbackRecoveryRequired = False
        Exit Sub
    End If
    If response = vbNo Then
        If PopupMessagesSuppressedForCurrentRun() Then
            discardResponse = vbNo
        Else
            discardResponse = CoreShowMessageResult("Discard the retained launcher-generation recovery state?" & vbCrLf & vbCrLf & "This accepts the current macro objects and launcher forms as authoritative and deletes the retained rollback backup files." & vbCrLf & vbCrLf & GeneratedMacroRollbackPathsForMessage(), vbExclamation Or vbYesNo Or vbDefaultButton2, "Confirm Recovery-State Discard")
        End If
        If discardResponse <> vbYes Then Err.Raise vbObjectError + 5931, "ResolveGeneratedMacroRollbackRecovery", "Generated-object recovery-state discard was not explicitly confirmed. Recovery state and backup files were retained."
        cleanupWarnings = CleanupGeneratedRollbackFiles()
        If Len(cleanupWarnings) > 0 Then CoreAddOperationNote cleanupWarnings
        Set mGeneratedMacroRollbackEntries = Nothing
        Set mLauncherFormRollbackEntries = Nothing
        mGeneratedMacroRollbackRecoveryRequired = False
        Exit Sub
    End If
    Err.Raise vbObjectError + 5932, "ResolveGeneratedMacroRollbackRecovery", "A new operation was cancelled because generated-object rollback recovery remains unresolved. Suppressed popups never authorize recovery-state discard."
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If GeneratedMacroRollbackEntryCount() > 0 Then mGeneratedMacroRollbackRecoveryRequired = True
    Err.Raise savedErrNumber, "ResolveGeneratedMacroRollbackRecovery", "Could not resolve retained generated-object rollback recovery before a new operation." & vbCrLf & "Operation: " & operationName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & vbCrLf & GeneratedMacroRollbackPathsForMessage()
End Sub

Private Function GeneratedMacroObjectExistsStrict(ByVal MacroTitle As String, ByVal stageText As String) As Boolean
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim doc As DAO.Document
    Dim containerObject As DAO.Container
    Dim objectLookupStarted As Boolean
    If Len(MacroTitle) = 0 Then Err.Raise vbObjectError + 5933, "GeneratedMacroObjectExistsStrict", "A strict generated-macro inspection received a blank macro title." & vbCrLf & "Stage: " & stageText
    Set db = GetHostDb()
    Set containerObject = db.Containers(MACRO_DAO_CONTAINER_NAME)
    containerObject.Documents.Refresh
    objectLookupStarted = True
    Set doc = containerObject.Documents(MacroTitle)
    GeneratedMacroObjectExistsStrict = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If savedErrNumber = 3265 And objectLookupStarted Then Exit Function
    Err.Raise savedErrNumber, "GeneratedMacroObjectExistsStrict", "Could not authoritatively inspect a generated macro during rollback. The error was not treated as macro absence." & vbCrLf & "Database: " & SharedDao_SafeDatabaseName(GetHostDb()) & vbCrLf & "Macro: " & MacroTitle & vbCrLf & "Stage: " & stageText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Function

Private Sub VerifyGeneratedMacroLoadable(ByVal MacroTitle As String, ByVal stageText As String)
    On Error GoTo ErrHandler
    Dim tempPath As String
    Dim failureText As String
    tempPath = Environ$("TEMP") & "\ADC_GeneratedMacroRollbackVerify_" & SharedName_SafeFileName(MacroTitle, ADCAT_OBJECT_TYPE_MACRO, 120) & "_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & Format$(Timer * 1000, "000000") & ".txt"
    Application.SaveAsText acMacro, MacroTitle, tempPath
    If GeneratedMacroRollbackBackupReadable(tempPath, failureText) = False Then Err.Raise vbObjectError + 5936, "VerifyGeneratedMacroLoadable", failureText
    On Error Resume Next
    If Len(tempPath) > 0 Then
        If Len(Dir$(tempPath)) > 0 Then Kill tempPath
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
    If Len(tempPath) > 0 Then
        If Len(Dir$(tempPath)) > 0 Then Kill tempPath
    End If
    On Error GoTo 0
    Err.Raise savedErrNumber, "VerifyGeneratedMacroLoadable", "The restored generated macro failed loadability verification." & vbCrLf & "Macro: " & MacroTitle & vbCrLf & "Stage: " & stageText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Sub

Private Sub TrackGeneratedMacroRollback(ByVal MacroTitle As String, ByVal targetExisted As Boolean, ByVal backupPath As String)
    On Error GoTo ErrHandler
    Dim entryData As Variant
    If mLauncherGenerationTrackingActive = False Then Exit Sub
    If mGeneratedMacroRollbackEntries Is Nothing Then Set mGeneratedMacroRollbackEntries = New Collection
    entryData = Array(MacroTitle, targetExisted, backupPath)
    mGeneratedMacroRollbackEntries.Add entryData
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    mGeneratedMacroRollbackRecoveryRequired = True
    Err.Raise savedErrNumber, "TrackGeneratedMacroRollback", "Could not retain generated-macro rollback metadata." & vbCrLf & "Macro: " & MacroTitle & vbCrLf & "Backup: " & IIf(Len(Trim$(backupPath)) = 0, "(none)", backupPath) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Sub

Private Function RollbackGeneratedMacroObjectsFromRun() As String
    On Error GoTo ErrHandler
    Dim item As Variant
    Dim entryData As Variant
    Dim MacroTitle As String
    Dim targetExisted As Boolean
    Dim backupPath As String
    Dim restoredCount As Long
    Dim deletedCount As Long
    Dim failedCount As Long
    Dim detailsText As String
    Dim rollbackCount As Long
    Dim objectWasPresent As Boolean
    Dim backupFailureText As String
    Dim entryErrNumber As Long
    Dim entryErrSource As String
    Dim entryErrDescription As String
    If mGeneratedMacroRollbackEntries Is Nothing Then
        RollbackGeneratedMacroObjectsFromRun = "Generated-macro rollback had no tracked macro objects."
        Exit Function
    End If
    For Each item In mGeneratedMacroRollbackEntries
        PumpSetupResponsiveness
        rollbackCount = rollbackCount + 1
        PulseProgressEvery rollbackCount, 0, "Generated macro rollback entries processed"
        entryData = item
        MacroTitle = vbNullString
        targetExisted = False
        backupPath = vbNullString
        If IsArray(entryData) Then
            If UBound(entryData) >= 0 Then MacroTitle = Trim$(CStr(entryData(0)))
            If UBound(entryData) >= 1 Then targetExisted = CBool(entryData(1))
            If UBound(entryData) >= 2 Then backupPath = Trim$(CStr(entryData(2)))
        End If
        If Len(MacroTitle) = 0 Then
            failedCount = failedCount + 1
            detailsText = SharedText_AppendNonBlankLine(detailsText, "Generated-macro rollback entry had a blank macro title. Backup: " & IIf(Len(backupPath) = 0, "(none)", backupPath))
        Else
            On Error GoTo EntryFailed
            objectWasPresent = GeneratedMacroObjectExistsStrict(MacroTitle, "before rolling back a generated macro")
            If targetExisted Then
                backupFailureText = vbNullString
                If GeneratedMacroRollbackBackupReadable(backupPath, backupFailureText) = False Then
                    failedCount = failedCount + 1
                    detailsText = SharedText_AppendNonBlankLine(detailsText, "Preserved current macro [" & MacroTitle & "] because its required rollback backup was unavailable. Backup: " & IIf(Len(backupPath) = 0, "(none)", backupPath) & "; " & backupFailureText)
                Else
                    If objectWasPresent Then DoCmd.DeleteObject acMacro, MacroTitle
                    If GeneratedMacroObjectExistsStrict(MacroTitle, "after deleting the replacement generated macro") Then Err.Raise vbObjectError + 5925, "RollbackGeneratedMacroObjectsFromRun", "The current macro remained present after rollback deletion. Macro: " & MacroTitle
                    Application.LoadFromText acMacro, MacroTitle, backupPath
                    If GeneratedMacroObjectExistsStrict(MacroTitle, "after restoring the original generated macro") = False Then Err.Raise vbObjectError + 5926, "RollbackGeneratedMacroObjectsFromRun", "The original macro was not present after rollback restoration. Macro: " & MacroTitle
                    VerifyGeneratedMacroLoadable MacroTitle, "verifying the restored original generated macro"
                    restoredCount = restoredCount + 1
                End If
            Else
                If objectWasPresent Then DoCmd.DeleteObject acMacro, MacroTitle
                If GeneratedMacroObjectExistsStrict(MacroTitle, "after deleting a newly generated macro") Then Err.Raise vbObjectError + 5927, "RollbackGeneratedMacroObjectsFromRun", "A newly generated macro remained present after rollback deletion. Macro: " & MacroTitle
                If objectWasPresent Then deletedCount = deletedCount + 1
            End If
            On Error GoTo ErrHandler
        End If
ContinueLoop:
    Next item
    RollbackGeneratedMacroObjectsFromRun = "Generated-macro rollback attempted." & vbCrLf & "New macro objects deleted: " & CStr(deletedCount) & vbCrLf & "Replaced macro objects restored: " & CStr(restoredCount) & vbCrLf & "Rollback failures: " & CStr(failedCount)
    If Len(Trim$(detailsText)) > 0 Then RollbackGeneratedMacroObjectsFromRun = RollbackGeneratedMacroObjectsFromRun & vbCrLf & detailsText
    Exit Function
EntryFailed:
    entryErrNumber = Err.Number
    entryErrSource = Err.Source
    entryErrDescription = Err.Description
    failedCount = failedCount + 1
    detailsText = SharedText_AppendNonBlankLine(detailsText, "Generated-macro rollback failed for [" & MacroTitle & "]. Backup: " & IIf(Len(backupPath) = 0, "(none - object was newly created)", backupPath) & "; Error " & CStr(entryErrNumber) & " from " & entryErrSource & ": " & entryErrDescription)
    Err.Clear
    On Error GoTo ErrHandler
    Resume ContinueLoop
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    mGeneratedMacroRollbackRecoveryRequired = True
    RollbackGeneratedMacroObjectsFromRun = "Generated-macro rollback failed before completion. Error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription & vbCrLf & GeneratedMacroRollbackPathsForMessage()
End Function

' Purpose: Resolves and returns expected generated macro object name for config from the supplied inputs and the authoritative ADC definitions/state used by this module.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   cfg (Input; TMacroRecord): ADC macro record containing the persisted accepted/current configuration values used by the operation; the routine reads it and may populate or update its fields when passed ByRef.
' Returns: String: Returns the physical/logical name resolved for expected generated macro object name for config under the module's naming/definition rules; vbNullString indicates that no applicable name can
'   be resolved where supported.
Public Function ExpectedGeneratedMacroObjectNameForConfig(ByRef cfg As TMacroRecord) As String
    If Len(Trim$(cfg.GeneratedMacroObjectName)) > 0 Then
        ExpectedGeneratedMacroObjectNameForConfig = cfg.GeneratedMacroObjectName
    Else
        ExpectedGeneratedMacroObjectNameForConfig = cfg.MacroTitle
    End If
End Function

' Purpose: Generates configured macro for config ID.
' Public role: Form-code/event-expression public entry point. Access form code or an event property can invoke this routine by name; its Public visibility is part of that UI contract.
' Parameters:
'   ConfigID (Input; Long): Primary-key ID of the configuration row targeted by the generated UI/action.
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub GenerateConfiguredMacro(ByVal configID As Long, Optional ByVal showMessage As Boolean = True)
    'Form-facing single-record generation entry point. Generation still executes through the current shared generation core.
    On Error GoTo ErrHandler
    Dim progressStarted As Boolean
    Dim generatedCount As Long
    Dim skippedCount As Long
    Dim sqlText As String
    Dim generationErrorEntries As Collection
    Dim cancelRequested As Boolean
    If configID <= 0 Then Err.Raise vbObjectError + 9192, "GenerateConfiguredMacro", "ConfigID must be greater than zero."
    progressStarted = CoreBeginOperationProgressIfNeeded("Generate " & DEFAULT_UTILITY_SHORT_NAME & " configured macro", 2)
    CoreUpdateSetupProgress "Check default Navigation Pane group"
    RequireDefaultCodeDbGeneratedMacrosNavigationPaneGroupForRoleContext "GenerateConfiguredMacro", False
    CoreUpdateSetupProgress "Generate configured macro ID " & CStr(configID)
    CoreResetOperationNotes
    InitializeNavigationPaneGroupingNotes
    RequireConfigTableForUtilityWorkflow "GenerateConfiguredMacro"
    sqlText = "SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "] WHERE " & SQL_FIELD_ID & "=" & CStr(configID)
    PreflightGeneratedMacroTitlesForSql sqlText
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: generationErrorEntries holds planned/recovery work or independently observed validation evidence for GenerateConfiguredMacro; keep it distinct from the source
    ' authority.
    Set generationErrorEntries = New Collection
    GenerateMacrosFromSql sqlText, False, generatedCount, skippedCount, generationErrorEntries, cancelRequested, True
    If generatedCount = 0 And skippedCount = 0 Then Err.Raise vbObjectError + 9193, "GenerateConfiguredMacro", "No configuration row exists for the requested ConfigID." & vbCrLf & "ConfigID: " & CStr(configID)
    If showMessage Then ShowMacroGenerationCompletionMessage "Configured macro generation complete.", generatedCount, skippedCount, generationErrorEntries, "Generate Configured Macro", "Configuration row ID: " & CStr(configID)
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
    If showMessage Then
        CoreShowErrorMessage "Could not generate the requested Access Development Console configured macro." & vbCrLf & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, vbCritical, "Generate Configured Macro", "GenerateConfiguredMacro", "Review"
        Exit Sub
    End If
    Err.Raise savedErrNumber, "GenerateConfiguredMacro", "Could not generate the requested Access Development Console configured macro." & vbCrLf & "ConfigID: " & CStr(configID) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub GenerateSelectedRunnerMacrosCore()
    On Error GoTo ErrHandler
    CoreResetOperationNotes
    Dim generatedCount As Long
    Dim skippedCount As Long
    Dim sqlText As String
    Dim generationErrorEntries As Collection
    Dim cancelRequested As Boolean
    InitializeNavigationPaneGroupingNotes
    RequireConfigTableForUtilityWorkflow "GenerateSelectedRunnerMacrosCore"
    EnsureNoActiveAutoExecConfigInconsistencies
    sqlText = "SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "] " & "WHERE " & SQL_FIELD_ACTIVE & "<>0 AND " & SQL_FIELD_SELECTED_FOR_RUNNER_MACRO_GENERATION & "<>0 " & ConfiguredMacroGenerationSqlFilter() & "AND UCase(Nz(" & SQL_FIELD_MACRO_TITLE & ",''))<>'AUTOEXEC' " & "AND UCase(Nz(" & SQL_FIELD_MACRO_KEY & ",''))<>'AUTOEXEC' " & "ORDER BY Nz(" & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER & ",0), " & SQL_FIELD_MACRO_TITLE & ", " & SQL_FIELD_MACRO_KEY
    PreflightGeneratedMacroTitlesForSql sqlText
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: generationErrorEntries holds planned/recovery work or independently observed validation evidence for GenerateSelectedRunnerMacrosCore; keep it distinct from the
    ' source authority.
    Set generationErrorEntries = New Collection
    GenerateMacrosFromSql sqlText, False, generatedCount, skippedCount, generationErrorEntries, cancelRequested
    ShowOrDeferMacroGenerationCompletionMessage "Macro generation complete.", generatedCount, skippedCount, generationErrorEntries, "Generate Selected Runner Macros", vbNullString
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    originalErrNumber = Err.Number
    Err.Raise originalErrNumber, "GenerateSelectedRunnerMacrosCore", FormatErrorWithContext("Could not generate selected Runner Macros.", "MC_Records table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & vbCrLf & "Selection: active configuration rows with SelectedForRunnerMacroGeneration=True; AutoExec excluded." & vbCrLf & "Generated count before failure: " & CStr(generatedCount) & vbCrLf & "Skipped count before failure: " & CStr(skippedCount))
End Sub

Private Sub GenerateAllConfiguredRunnerMacrosCore()
    On Error GoTo ErrHandler
    CoreResetOperationNotes
    Dim generatedCount As Long
    Dim skippedCount As Long
    Dim sqlText As String
    Dim generationErrorEntries As Collection
    Dim cancelRequested As Boolean
    InitializeNavigationPaneGroupingNotes
    RequireConfigTableForUtilityWorkflow "GenerateAllConfiguredRunnerMacrosCore"
    EnsureNoActiveAutoExecConfigInconsistencies
    sqlText = "SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "] " & "WHERE " & SQL_FIELD_ACTIVE & "<>0 " & ConfiguredMacroGenerationSqlFilter() & "AND UCase(Nz(" & SQL_FIELD_MACRO_TITLE & ",''))<>'AUTOEXEC' " & "AND UCase(Nz(" & SQL_FIELD_MACRO_KEY & ",''))<>'AUTOEXEC' " & "ORDER BY Nz(" & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER & ",0), " & SQL_FIELD_MACRO_TITLE & ", " & SQL_FIELD_MACRO_KEY
    PreflightGeneratedMacroTitlesForSql sqlText
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: generationErrorEntries holds planned/recovery work or independently observed validation evidence for GenerateAllConfiguredRunnerMacrosCore; keep it distinct from
    ' the source authority.
    Set generationErrorEntries = New Collection
    GenerateMacrosFromSql sqlText, False, generatedCount, skippedCount, generationErrorEntries, cancelRequested
    ShowOrDeferMacroGenerationCompletionMessage "All Active=True configured Runner Macro generation complete.", generatedCount, skippedCount, generationErrorEntries, "Generate All Configured Runner Macros", "AutoExec: excluded. Use GenerateAutoExecMacro for AutoExec."
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    originalErrNumber = Err.Number
    Err.Raise originalErrNumber, "GenerateAllConfiguredRunnerMacrosCore", FormatErrorWithContext("Could not generate all configured Runner Macros.", "MC_Records table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & vbCrLf & "Selection: all active configuration rows; AutoExec excluded." & vbCrLf & "Generated count before failure: " & CStr(generatedCount) & vbCrLf & "Skipped count before failure: " & CStr(skippedCount))
End Sub

Private Sub GeneratePromptedRunnerMacrosCore(ByVal completionMessageTitle As String, ByVal callerName As String)
    On Error GoTo ErrHandler
    If Len(completionMessageTitle) = 0 Then completionMessageTitle = "Generate Prompted Runner Macros"
    If Len(callerName) = 0 Then callerName = "GeneratePromptedRunnerMacrosCore"
    CoreResetOperationNotes
    Dim generatedCount As Long
    Dim skippedCount As Long
    Dim sqlText As String
    Dim generationErrorEntries As Collection
    Dim cancelRequested As Boolean
    InitializeNavigationPaneGroupingNotes
    RequireConfigTableForUtilityWorkflow "GeneratePromptedRunnerMacrosCore"
    EnsureNoActiveAutoExecConfigInconsistencies
    sqlText = "SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "] " & "WHERE " & SQL_FIELD_ACTIVE & "<>0 " & ConfiguredMacroGenerationSqlFilter() & "AND UCase(Nz(" & SQL_FIELD_MACRO_TITLE & ",''))<>'AUTOEXEC' " & "AND UCase(Nz(" & SQL_FIELD_MACRO_KEY & ",''))<>'AUTOEXEC' " & "ORDER BY Nz(" & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER & ",0), " & SQL_FIELD_MACRO_TITLE & ", " & SQL_FIELD_MACRO_KEY
    PreflightGeneratedMacroTitlesForSql sqlText
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: generationErrorEntries holds planned/recovery work or independently observed validation evidence for GeneratePromptedRunnerMacrosCore; keep it distinct from the
    ' source authority.
    Set generationErrorEntries = New Collection
    GenerateMacrosFromSql sqlText, True, generatedCount, skippedCount, generationErrorEntries, cancelRequested
    If cancelRequested Then
        CoreAddOperationNote "Prompted Runner Macro generation was cancelled after partial processing. Generated: " & CStr(generatedCount) & "; skipped: " & CStr(skippedCount) & ". Custom Macro Launcher post-processing was not run."
        LogTraceMessage "Info", "MacroGeneration", "GeneratePromptedRunnerMacrosCore", "Prompted Runner Macro generation cancelled after partial processing.", "Generated before cancellation: " & CStr(generatedCount) & vbCrLf & "Skipped before cancellation: " & CStr(skippedCount), "Prompted Runner Macro generation cancelled", 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config"), "Table", "CodeDb", "MacroGeneration"
        Exit Sub
    End If
    ShowMacroGenerationCompletionMessage "Prompted Runner Macro generation complete.", generatedCount, skippedCount, generationErrorEntries, completionMessageTitle, "AutoExec: excluded. Use GenerateAutoExecMacro for AutoExec."
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    originalErrNumber = Err.Number
    Err.Raise originalErrNumber, callerName, FormatErrorWithContext("Could not generate prompted Runner Macros.", "MC_Records table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & vbCrLf & "Selection: prompted active configuration rows; AutoExec excluded." & vbCrLf & "Generated count before failure: " & CStr(generatedCount) & vbCrLf & "Skipped count before failure: " & CStr(skippedCount))
End Sub

' Purpose: Resolves format display value from the supplied/current ADC state, using `SharedText_Nz` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   value (Input; Variant): Variant input representing value; the routine accepts the caller-provided value in the form required by its existing validation and conversion logic.
' Returns: String: Returns the format display value text value. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function FormatDisplayValue(ByVal value As Variant) As String
    'Formats a Variant for display, returning the literal placeholder (blank) for empty text.
    Dim textValue As String
    textValue = SharedText_Nz(value)
    If Len(Trim$(textValue)) = 0 Then
        FormatDisplayValue = "(blank)"
    Else
        FormatDisplayValue = textValue
    End If
End Function

Private Sub PreflightGeneratedMacroTitlesForSql(ByVal sqlText As String)
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim titleRows As Object
    Dim titleCounts As Object
    Dim lines As Collection
    Dim macroTitleText As String
    Dim keyText As String
    Dim rowText As String
    Dim item As Variant
    Dim rowID As Long
    Dim MacroKey As String
    Dim MacroTitle As String
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: titleRows, titleCounts, lines holds planned/recovery work or independently observed validation evidence for PreflightGeneratedMacroTitlesForSql; keep it distinct
    ' from the source authority.
    Set titleRows = CreateObject("Scripting.Dictionary")
    Set titleCounts = CreateObject("Scripting.Dictionary")
    titleRows.CompareMode = 1
    titleCounts.CompareMode = 1
    Set lines = New Collection
    Set rs = ADCDb().OpenRecordset(sqlText, dbOpenSnapshot)
    Do While Not rs.EOF
        PumpSetupResponsiveness
        rowID = SharedValue_NzLong(FieldValueIfExists(rs, FIELD_ID))
        MacroKey = SharedText_Nz(FieldValueIfExists(rs, FIELD_MACRO_KEY))
        MacroTitle = SharedText_Nz(FieldValueIfExists(rs, FIELD_MACRO_TITLE))
        macroTitleText = MacroTitle
        If Len(Trim$(macroTitleText)) = 0 Then
            lines.Add "Selected active config row has a blank MacroTitle." & vbCrLf & "  ID: " & CStr(rowID) & vbCrLf & "  MacroKey: " & FormatDisplayValue(MacroKey)
        ElseIf Len(macroTitleText) > MAX_ACCESS_OBJECT_NAME_LENGTH Then
            lines.Add "Selected active config row has an overlength MacroTitle." & vbCrLf & "  ID: " & CStr(rowID) & vbCrLf & "  MacroKey: " & FormatDisplayValue(MacroKey) & vbCrLf & "  MacroTitle: " & FormatDisplayValue(MacroTitle) & vbCrLf & "  Length: " & CStr(Len(macroTitleText)) & vbCrLf & "  Maximum: " & CStr(MAX_ACCESS_OBJECT_NAME_LENGTH)
        Else
            keyText = macroTitleText
            rowText = "  - ID: " & CStr(rowID) & "; MacroKey: " & FormatDisplayValue(MacroKey) & "; MacroTitle: " & FormatDisplayValue(MacroTitle)
            If titleRows.Exists(keyText) Then
                titleRows(keyText) = titleRows(keyText) & vbCrLf & rowText
                titleCounts(keyText) = CLng(titleCounts(keyText)) + 1
            Else
                titleRows.Add keyText, rowText
                titleCounts.Add keyText, 1
            End If
        End If
        rs.MoveNext
    Loop
    For Each item In titleCounts.keys
        PumpSetupResponsiveness
        If CLng(titleCounts(item)) > 1 Then
            lines.Add "Duplicate generated macro object name detected: " & CStr(item) & vbCrLf & titleRows(item)
        End If
    Next item
    If lines.count > 0 Then
        Err.Raise vbObjectError + 5222, "PreflightGeneratedMacroTitlesForSql", "MacroTitle preflight failed. Generation stopped before replacing any macro objects." & vbCrLf & vbCrLf & SharedCollection_Join(lines, vbCrLf & vbCrLf) & vbCrLf & vbCrLf & "MacroTitle must be nonblank, no longer than " & CStr(MAX_ACCESS_OBJECT_NAME_LENGTH) & " characters, and unique among the selected Active=True rows."
    End If
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Exit Sub
ErrHandler:
    Dim handlerErrNumber As Long
    Dim handlerErrDescription As String
    handlerErrNumber = Err.Number
    handlerErrDescription = Err.Description
    If handlerErrNumber < 0 Then
        On Error Resume Next
        If Not rs Is Nothing Then rs.Close
        On Error GoTo 0
        Err.Raise handlerErrNumber, "PreflightGeneratedMacroTitlesForSql", FormatErrorWithContext("Could not run generated MacroTitle preflight.", "MC_Records table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & vbCrLf & "Preflight: validate selected MacroTitle values before any macro object is created or replaced.")
    Else
        Err.Raise handlerErrNumber, "PreflightGeneratedMacroTitlesForSql", "Could not run generated MacroTitle preflight." & vbCrLf & "MC_Records table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & vbCrLf & "Err.Number: " & CStr(handlerErrNumber) & vbCrLf & "Err.Description: " & handlerErrDescription
    End If
End Sub

Private Sub GenerateMacrosFromSql(ByVal sqlText As String, ByVal prompted As Boolean, ByRef generatedCount As Long, ByRef skippedCount As Long, ByRef generationErrorEntries As Collection, ByRef cancelRequested As Boolean, Optional ByVal firstRowOnly As Boolean = False)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim cfg As TMacroRecord
    Dim response As VbMsgBoxResult
    Dim processedCount As Long
    Dim totalCount As Long
    Dim currentMacroDisplay As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim cleanupWarnings As String
    cancelRequested = False
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: generationErrorEntries holds planned/recovery work or independently observed validation evidence for GenerateMacrosFromSql; keep it distinct from the source
    ' authority.
    If generationErrorEntries Is Nothing Then Set generationErrorEntries = New Collection
    BeginLauncherGenerationTracking firstRowOnly
    Set db = ADCDb()
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveLast
        totalCount = rs.RecordCount
        rs.MoveFirst
        If firstRowOnly And totalCount > 1 Then totalCount = 1
    End If
    Do While Not rs.EOF
        PumpSetupResponsiveness
        processedCount = processedCount + 1
        MacroRecordFromRecordset rs, cfg
        currentMacroDisplay = Trim$(cfg.MacroTitle)
        If Len(currentMacroDisplay) = 0 Then currentMacroDisplay = Trim$(cfg.MacroKey)
        If Len(currentMacroDisplay) = 0 Then currentMacroDisplay = "Config ID " & CStr(cfg.ID)
        If totalCount > 0 Then
            CorePulseProgress "Macro " & CStr(processedCount) & " of " & CStr(totalCount) & " > " & currentMacroDisplay
        Else
            CorePulseProgress "Macro " & CStr(processedCount) & " > " & currentMacroDisplay
        End If
        If prompted Then
            response = PromptForMacroGeneration(cfg)
            Select Case response
                Case vbYes
                    If GenerateMacroFromConfig(cfg, generationErrorEntries) Then
                        generatedCount = generatedCount + 1
                    Else
                        skippedCount = skippedCount + 1
                    End If
                Case vbNo
                    skippedCount = skippedCount + 1
                Case vbCancel
                    cancelRequested = True
                    MarkSystemCommandCancelled
                    Exit Do
            End Select
        Else
            If GenerateMacroFromConfig(cfg, generationErrorEntries) Then
                generatedCount = generatedCount + 1
            Else
                skippedCount = skippedCount + 1
            End If
        End If
        If firstRowOnly Then Exit Do
        rs.MoveNext
    Loop
    If cancelRequested Then
        cleanupWarnings = CleanupGeneratedRollbackFiles()
        If Len(cleanupWarnings) > 0 Then CoreAddOperationNote cleanupWarnings
        Set mGeneratedMacroRollbackEntries = Nothing
        Set mLauncherFormRollbackEntries = Nothing
        mGeneratedMacroRollbackRecoveryRequired = False
        mLauncherGenerationTrackingActive = False
    Else
        CorePulseProgress "Generated " & CStr(generatedCount) & " macros > updating launcher tracking"
        ProcessLauncherGenerationTracking
    End If
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    mLauncherGenerationTrackingActive = False
    If GeneratedMacroRollbackEntryCount() > 0 Then mGeneratedMacroRollbackRecoveryRequired = True
    AddMacroGenerationSqlErrorEntry generationErrorEntries, "GenerateMacrosFromSql", savedErrNumber, savedErrDescription, sqlText
    Err.Raise savedErrNumber, "GenerateMacrosFromSql", FormatErrorWithContext("Could not generate macro objects from the selected config-row SQL.", "MC_Records table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & vbCrLf & "Prompted: " & CStr(prompted) & vbCrLf & "First row only: " & CStr(firstRowOnly) & vbCrLf & "Generated count before failure: " & CStr(generatedCount) & vbCrLf & "Skipped count before failure: " & CStr(skippedCount) & vbCrLf & "Retained rollback entries: " & CStr(GeneratedMacroRollbackEntryCount()) & vbCrLf & GeneratedMacroRollbackPathsForMessage() & vbCrLf & "Original Err.Source: " & savedErrSource & vbCrLf & "Original Err.Description: " & savedErrDescription & vbCrLf & "SQL: " & SharedMessage_OneLine(sqlText))
End Sub

Private Function PromptForMacroGeneration(ByRef cfg As TMacroRecord) As VbMsgBoxResult
    Dim promptText As String
    promptText = "Create or replace this Access macro object?" & vbCrLf & vbCrLf & "MacroTitle: " & cfg.MacroTitle & vbCrLf & "MacroKey: " & cfg.MacroKey & vbCrLf & "Action: " & cfg.MacroAction & vbCrLf & "ConfigOwnerDatabaseRole: " & cfg.ConfigOwnerDatabaseRole & vbCrLf & "Description: " & EffectiveActionDescription(cfg)
    If StrComp(cfg.MacroAction, ACTION_RUNCODE, vbTextCompare) = 0 Then
        promptText = promptText & vbCrLf & "Target: " & EffectiveTargetCallText(cfg)
    ElseIf ActionIsAccessObjectAction(cfg.MacroAction) Then
        promptText = promptText & vbCrLf & "Object: " & ConfiguredObjectActionTargetName(cfg)
    End If
    PromptForMacroGeneration = CoreShowMessageResult(promptText, vbQuestion Or vbYesNoCancel Or vbDefaultButton2, "Create Macro")
End Function

Private Function GenerateMacroFromConfig(ByRef cfg As TMacroRecord, ByRef generationErrorEntries As Collection) As Boolean
    On Error GoTo ErrHandler
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: generationErrorEntries holds planned/recovery work or independently observed validation evidence for GenerateMacroFromConfig; keep it distinct from the source
    ' authority.
    If generationErrorEntries Is Nothing Then Set generationErrorEntries = New Collection
    ValidateConfigClassification cfg
    If ConfigurationStatusBlocksGeneration(cfg.ConfigurationStatus) Then
        UpdateMacroGenerationStatus cfg.ID, GENERATED_RESULT_SKIPPED, "ConfigurationStatus prevents generation: " & cfg.ConfigurationStatus, vbNullString
        GenerateMacroFromConfig = False
        Exit Function
    End If
    ValidateMacroConfig cfg
    Select Case UCase$(Trim$(cfg.MacroAction))
        Case UCase$(ACTION_RUNCODE)
            CreateOrReplaceRunCodeMacro cfg
            UpdateMacroGenerationStatus cfg.ID, GENERATED_RESULT_GENERATED, "Generated RunCode macro.", cfg.MacroTitle
            GenerateMacroFromConfig = True
        Case UCase$(ACTION_OPENFORM)
            CreateOrReplaceOpenFormMacro cfg
            UpdateMacroGenerationStatus cfg.ID, GENERATED_RESULT_GENERATED, "Generated OpenForm macro.", cfg.MacroTitle
            GenerateMacroFromConfig = True
        Case UCase$(ACTION_OPENTABLE), UCase$(ACTION_OPENQUERY), UCase$(ACTION_OPENREPORT), UCase$(ACTION_RUNMACRO)
            'Object-action rows use the same generated RunCode dispatcher shell so confirmation behavior and runtime diagnostics stay centralized.
            CreateOrReplaceRunCodeMacro cfg
            UpdateMacroGenerationStatus cfg.ID, GENERATED_RESULT_GENERATED, "Generated " & cfg.MacroAction & " dispatcher macro.", cfg.MacroTitle
            GenerateMacroFromConfig = True
        Case Else
            Err.Raise vbObjectError + 5200, "GenerateMacroFromConfig", "Unsupported MacroAction: " & cfg.MacroAction & vbCrLf & "MacroKey: " & cfg.MacroKey & vbCrLf & "MacroTitle: " & cfg.MacroTitle & vbCrLf & "Configuration row ID: " & CStr(cfg.ID) & vbCrLf & "MacroAction must be RunCode, OpenForm, OpenTable, OpenQuery, OpenReport, or RunMacro."
    End Select
    If GenerateMacroFromConfig Then TrackLauncherGeneratedConfig cfg
    Exit Function
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    UpdateMacroGenerationStatus cfg.ID, GeneratedResultFromGenerationError(savedErrDescription), Left$(savedErrDescription, 2000), cfg.MacroTitle
    AddMacroGenerationErrorEntry generationErrorEntries, cfg, "GenerateMacroFromConfig", savedErrNumber, "Original source: " & savedErrSource & vbCrLf & savedErrDescription
    On Error GoTo 0
    If savedErrNumber = MACRO_REPLACEMENT_RECOVERY_FAILURE_ERROR_NUMBER Then Err.Raise savedErrNumber, "GenerateMacroFromConfig", "Configured macro generation stopped because destructive replacement recovery failed. The original macro may be missing or unusable, and retained recovery state must be resolved before another top-level operation." & vbCrLf & "Configuration row ID: " & CStr(cfg.ID) & vbCrLf & "MacroKey: " & cfg.MacroKey & vbCrLf & "MacroTitle: " & cfg.MacroTitle & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & vbCrLf & GeneratedMacroRollbackPathsForMessage()
    GenerateMacroFromConfig = False
End Function

Private Sub ShowOrDeferMacroGenerationCompletionMessage(ByVal summaryTitle As String, ByVal generatedCount As Long, ByVal skippedCount As Long, ByVal generationErrorEntries As Collection, ByVal messageTitle As String, ByVal extraDetailText As String)
    If mDeferMacroGenerationCompletionMessage Then
        mDeferredMacroGenerationCompletionPending = True
        mDeferredMacroGenerationSummaryTitle = summaryTitle
        mDeferredMacroGenerationGeneratedCount = generatedCount
        mDeferredMacroGenerationSkippedCount = skippedCount
        Set mDeferredMacroGenerationErrorEntries = generationErrorEntries
        mDeferredMacroGenerationMessageTitle = messageTitle
        mDeferredMacroGenerationExtraDetailText = extraDetailText
        Exit Sub
    End If
    ShowMacroGenerationCompletionMessage summaryTitle, generatedCount, skippedCount, generationErrorEntries, messageTitle, extraDetailText
End Sub

' Purpose: Sets macro generation completion deferral for defer completion, show pending completion.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   deferCompletion (Input; Boolean): Boolean state describing whether defer completion applies to this invocation; True selects the named condition/path and False selects its complementary/default path.
'   showPendingCompletion (Input Optional; Boolean): True displays pending completion; False suppresses that user-facing presentation without changing the underlying work. Default: False.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub SetMacroGenerationCompletionDeferral(ByVal deferCompletion As Boolean, Optional ByVal showPendingCompletion As Boolean = False)
    On Error GoTo ErrHandler
    If deferCompletion Then
        ClearDeferredMacroGenerationCompletionState
        mDeferMacroGenerationCompletionMessage = True
        Exit Sub
    End If
    mDeferMacroGenerationCompletionMessage = False
    If showPendingCompletion And mDeferredMacroGenerationCompletionPending Then
        ShowMacroGenerationCompletionMessage mDeferredMacroGenerationSummaryTitle, mDeferredMacroGenerationGeneratedCount, mDeferredMacroGenerationSkippedCount, mDeferredMacroGenerationErrorEntries, mDeferredMacroGenerationMessageTitle, mDeferredMacroGenerationExtraDetailText
    End If
    ClearDeferredMacroGenerationCompletionState
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    ClearDeferredMacroGenerationCompletionState
    Err.Raise savedErrNumber, "SetMacroGenerationCompletionDeferral", "Could not finalize deferred macro-generation completion reporting after the modal progress dialog closed." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub ClearDeferredMacroGenerationCompletionState()
    mDeferMacroGenerationCompletionMessage = False
    mDeferredMacroGenerationCompletionPending = False
    mDeferredMacroGenerationSummaryTitle = vbNullString
    mDeferredMacroGenerationGeneratedCount = 0
    mDeferredMacroGenerationSkippedCount = 0
    Set mDeferredMacroGenerationErrorEntries = Nothing
    mDeferredMacroGenerationMessageTitle = vbNullString
    mDeferredMacroGenerationExtraDetailText = vbNullString
End Sub

Private Sub ShowMacroGenerationCompletionMessage(ByVal summaryTitle As String, ByVal generatedCount As Long, ByVal skippedCount As Long, ByVal generationErrorEntries As Collection, ByVal messageTitle As String, ByVal extraDetailText As String, Optional ByVal systemMacroGeneratedCount As Long = 0, Optional ByVal systemMacroSkippedCount As Long = 0, Optional ByVal systemMacroDetailsText As String = vbNullString)
    On Error GoTo ErrHandler
    Dim messageText As String
    Dim reportPath As String
    Dim response As VbMsgBoxResult
    Dim errorCount As Long
    If Not generationErrorEntries Is Nothing Then errorCount = generationErrorEntries.count
    messageText = summaryTitle & vbCrLf & vbCrLf & "Configured macros generated/replaced: " & CStr(generatedCount) & vbCrLf & "Configured macros skipped: " & CStr(skippedCount) & vbCrLf & "System macros installed/refreshed: " & CStr(systemMacroGeneratedCount) & vbCrLf & "System macros skipped: " & CStr(systemMacroSkippedCount) & vbCrLf & "Errors: " & CStr(errorCount)
    messageText = messageText & SystemMacroDetailsForMessage(systemMacroDetailsText)
    If Len(Trim$(extraDetailText)) > 0 Then
        messageText = messageText & vbCrLf & extraDetailText
    End If
    messageText = messageText & CoreOperationSummaryForMessage()
    messageText = messageText & NavigationPaneGroupingSummaryForMessage()
    If errorCount > 0 Then
        reportPath = MacroGenerationErrorReportPath()
        WriteTextFile reportPath, MacroGenerationErrorReport(generationErrorEntries, summaryTitle)
        messageText = messageText & vbCrLf & vbCrLf & "Detailed generation error report:" & vbCrLf & reportPath
        CorePulseProgress "Macro generation completed with row-level error(s)."
        LogTraceMessage "Error", "MacroGeneration", "ShowMacroGenerationCompletionMessage", "Macro generation completed with row-level error(s).", messageText, "Report generation completion", 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config"), "Table", "CodeDb", "MacroGeneration"
        CoreShowErrorMessage messageText, vbCritical, messageTitle, "ShowMacroGenerationCompletionMessage", "MacroGeneration"
        response = CoreShowMessageResult("Open the detailed generation error report now?" & vbCrLf & vbCrLf & reportPath, vbQuestion Or vbYesNo Or vbDefaultButton2, "Open Macro Generation Error Report", "ShowMacroGenerationCompletionMessage", "MacroGeneration")
        If response = vbYes Then OpenTextReport reportPath
    Else
        ReportCoreActionSuccess "ShowMacroGenerationCompletionMessage", "MacroGeneration", "Macro generation completed successfully.", messageText, "Report generation completion", ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config"), "Table", "CodeDb", "MacroGeneration", True, messageTitle
    End If
    Exit Sub
ErrHandler:
    messageText = summaryTitle & vbCrLf & vbCrLf & "Configured macros generated/replaced: " & CStr(generatedCount) & vbCrLf & "Configured macros skipped: " & CStr(skippedCount) & vbCrLf & "System macros installed/refreshed: " & CStr(systemMacroGeneratedCount) & vbCrLf & "System macros skipped: " & CStr(systemMacroSkippedCount) & vbCrLf & "Errors: " & CStr(errorCount) & vbCrLf & vbCrLf & "The completion summary could not write/open the detailed error report." & vbCrLf & "Err.Description: " & Err.Description
    LogTraceMessage "Error", "MacroGeneration", "ShowMacroGenerationCompletionMessage", "Macro generation completion reporting failed.", messageText, "Report generation completion", Err.Number, Err.Source, Err.Description, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config"), "Table", "CodeDb", "MacroGeneration"
    CoreShowErrorMessage messageText, vbCritical, messageTitle, "ShowMacroGenerationCompletionMessage", "MacroGeneration"
End Sub

Private Sub AddMacroGenerationSqlErrorEntry(ByRef generationErrorEntries As Collection, ByVal sourceName As String, ByVal errNumber As Long, ByVal errDescription As String, ByVal sqlText As String)
    Dim entryText As String
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: generationErrorEntries holds planned/recovery work or independently observed validation evidence for AddMacroGenerationSqlErrorEntry; keep it distinct from the
    ' source authority.
    If generationErrorEntries Is Nothing Then Set generationErrorEntries = New Collection
    entryText = "Generation SQL/runtime failure" & vbCrLf & String$(72, "-") & vbCrLf & "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB & vbCrLf & "MC_Records table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & vbCrLf & "Source: " & sourceName & vbCrLf & "Err.Number: " & CStr(errNumber) & vbCrLf & "Err.Description: " & errDescription & vbCrLf & IIf(True, "SQL:" & vbCrLf & sqlText & vbCrLf, vbNullString)
    generationErrorEntries.Add entryText
    LogTraceMessage "Error", "MacroGeneration", sourceName, "Generation SQL/runtime failure.", entryText, "Collect generation SQL error", errNumber, sourceName, errDescription, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config"), "Table", "CodeDb", "MacroGeneration"
End Sub

Private Sub AddMacroGenerationErrorEntry(ByRef generationErrorEntries As Collection, ByRef cfg As TMacroRecord, ByVal sourceName As String, ByVal errNumber As Long, ByVal errDescription As String)
    Dim entryText As String
    Dim entryLines As Collection
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: generationErrorEntries holds planned/recovery work or independently observed validation evidence for AddMacroGenerationErrorEntry; keep it distinct from the source
    ' authority.
    ' INTENTIONAL ORDERED VIEW: entryLines preserves the deterministic sequence required by AddMacroGenerationErrorEntry; keyed-only or unordered storage would not preserve the consumer semantics.
    If generationErrorEntries Is Nothing Then Set generationErrorEntries = New Collection
    Set entryLines = New Collection
    entryLines.Add "Macro generation failure"
    entryLines.Add String$(72, "-")
    entryLines.Add "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB
    entryLines.Add "MC_Records table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")
    entryLines.Add "MacroKey: " & cfg.MacroKey
    entryLines.Add "MacroTitle: " & cfg.MacroTitle
    entryLines.Add "MacroAction: " & cfg.MacroAction
    entryLines.Add "ConfigOwnerDatabaseRole: " & cfg.ConfigOwnerDatabaseRole
    entryLines.Add "TargetProcedureName: " & IIf(Len(Trim$(cfg.TargetProcedureName)) = 0, "(blank)", cfg.TargetProcedureName)
    entryLines.Add "TargetProcedureKind: " & IIf(Len(Trim$(cfg.TargetProcedureKind)) = 0, TARGET_PROCEDURE_KIND_AUTO, cfg.TargetProcedureKind)
    entryLines.Add "TargetArguments: " & IIf(Len(Trim$(cfg.TargetArguments)) = 0, "(none)", cfg.TargetArguments)
    entryLines.Add "TargetObjectName: " & IIf(Len(Trim$(cfg.TargetObjectName)) = 0, "(blank)", cfg.TargetObjectName)
    entryLines.Add "Source: " & sourceName
    entryLines.Add "Err.Number: " & CStr(errNumber)
    entryLines.Add "Err.Description: " & errDescription
    entryText = SharedCollection_Join(entryLines, vbCrLf)
    entryText = entryText & vbCrLf & vbCrLf & "Diagnostic notes:" & vbCrLf & "- Generation validates the configured target before replacing a macro object." & vbCrLf & "- RunCode targets must resolve to exactly one public or implicit-public Function/Sub in the configured ConfigOwnerDatabaseRole." & vbCrLf & "- Access object actions must resolve to a visible HostDb Access object; CodeDb object actions are intentionally unsupported." & vbCrLf & "- If this is an ambiguity failure, rename duplicate public procedures or change ConfigOwnerDatabaseRole." & vbCrLf & "- If this is a replacement failure, the replacement routine attempts to restore the original macro object before returning this error."
    generationErrorEntries.Add entryText & vbCrLf
    LogTraceMessage "Error", "MacroGeneration", sourceName, "Macro generation row failed.", entryText, "Generate macro from config", errNumber, sourceName, errDescription, cfg.MacroTitle, "Macro", cfg.ConfigOwnerDatabaseRole, "MacroGeneration"
End Sub

Private Function MacroGenerationErrorReport(ByVal generationErrorEntries As Collection, ByVal summaryTitle As String) As String
    Dim lines As Collection
    Dim item As Variant
    Dim itemNumber As Long
    ' INTENTIONAL ORDERED VIEW: lines preserves the deterministic sequence required by MacroGenerationErrorReport; keyed-only or unordered storage would not preserve the consumer semantics.
    Set lines = New Collection
    lines.Add "Access Development Console Generation Error Report"
    lines.Add "Generated: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    lines.Add "Summary: " & summaryTitle
    lines.Add "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB
    lines.Add "MC_Records table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")
    lines.Add "HostDb: " & GetHostDb().Name
    lines.Add "CodeDb: " & codeDb.Name
    lines.Add "Detailed messages enabled: " & CStr(True)
    lines.Add String$(72, "=")
    lines.Add ""
    If generationErrorEntries Is Nothing Then
        lines.Add "No generation errors were collected."
    ElseIf generationErrorEntries.count = 0 Then
        lines.Add "No generation errors were collected."
    Else
        For Each item In generationErrorEntries
            PumpSetupResponsiveness
            itemNumber = itemNumber + 1
            lines.Add "ERROR " & CStr(itemNumber)
            lines.Add CStr(item)
            lines.Add ""
        Next item
    End If
    MacroGenerationErrorReport = SharedCollection_Join(lines, vbCrLf)
End Function

Private Function MacroGenerationErrorReportPath() As String
    Dim db As DAO.Database
    Set db = ADCDb()
    MacroGenerationErrorReportPath = Environ$("TEMP") & "\GenerationErrors_" & SharedName_SafeFileName(DatabaseBaseNameFromPath(db.Name), ADCAT_OBJECT_TYPE_MACRO, 120) & "_" & Format$(Now, "yyyymmdd_hhnnss") & ".txt"
End Function

Private Sub UpdateMacroGenerationStatus(ByVal configID As Long, ByVal statusText As String, Optional ByVal notesText As String = vbNullString, Optional ByVal generatedObjectName As String = vbNullString)
    On Error GoTo ErrHandler
    If configID <= 0 Then Exit Sub
    ADCDb().Execute "UPDATE [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "] SET " & SQL_FIELD_GENERATED_MACRO_OBJECT_NAME & "=" & IIf(Len(Trim$(generatedObjectName)) = 0, "Null", SharedDao_SqlQuote(generatedObjectName)) & ", " & SQL_FIELD_LAST_GENERATED_ON & "=Now(), " & SQL_FIELD_LAST_GENERATED_RESULT & "=" & SharedDao_SqlQuote(Left$(statusText, 30)) & ", " & SQL_FIELD_LAST_GENERATED_NOTES & "=" & IIf(Len(Trim$(notesText)) = 0, "Null", SharedDao_SqlQuote(notesText)) & " " & "WHERE " & SQL_FIELD_ID & "=" & CStr(configID), dbFailOnError
    InvalidateDuplicateReviewCache
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "UpdateMacroGenerationStatus", "Could not persist macro-generation status for configuration row ID " & CStr(configID) & "." & vbCrLf & "Requested result: " & statusText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Sub

' Purpose: Performs generated run code macro requires confirmation by coordinating `StrComp` and `End` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   cfg (Input; TMacroRecord): ADC macro record containing the persisted accepted/current configuration values used by the operation; the routine reads it and may populate or update its fields when passed ByRef.
' Returns: Boolean: True when the generated RunCode macro requires an explicit confirmation step under its effective action policy; False when it may run without that confirmation.
Public Function GeneratedRunCodeMacroRequiresConfirmation(ByRef cfg As TMacroRecord) As Boolean
    Dim modeText As String
    If StrComp(Trim$(cfg.MacroTitle), CONFIRM_AUTOEXEC, vbTextCompare) = 0 Then Exit Function
    modeText = LCase$(Trim$(cfg.ConfirmationMode))
    Select Case modeText
        Case "", LCase$(CONFIRM_AUTO)
            GeneratedRunCodeMacroRequiresConfirmation = True
        Case "yes", "prompt", "confirm", LCase$(CONFIRM_ALWAYS)
            GeneratedRunCodeMacroRequiresConfirmation = True
        Case "no", "none", "false", LCase$(CONFIRM_NEVER), LCase$(CONFIRM_AUTOEXEC)
            GeneratedRunCodeMacroRequiresConfirmation = False
        Case Else
            GeneratedRunCodeMacroRequiresConfirmation = True
    End Select
End Function

Private Sub RestoreOriginalAfterFailedReplacement(ByVal MacroTitle As String, ByVal backupPath As String)
    On Error GoTo ErrHandler
    Dim backupFailureText As String
    If GeneratedMacroRollbackBackupReadable(backupPath, backupFailureText) = False Then Err.Raise vbObjectError + 5952, "RestoreOriginalAfterFailedReplacement", backupFailureText
    If GeneratedMacroObjectExistsStrict(MacroTitle, "before restoring the original after failed replacement") Then DoCmd.DeleteObject acMacro, MacroTitle
    If GeneratedMacroObjectExistsStrict(MacroTitle, "after deleting a partial failed replacement") Then Err.Raise vbObjectError + 5953, "RestoreOriginalAfterFailedReplacement", "A partial replacement macro remained present after deletion." & vbCrLf & "Macro: " & MacroTitle
    Application.LoadFromText acMacro, MacroTitle, backupPath
    If GeneratedMacroObjectExistsStrict(MacroTitle, "after restoring the original following failed replacement") = False Then Err.Raise vbObjectError + 5954, "RestoreOriginalAfterFailedReplacement", "The original macro was not present after restoration." & vbCrLf & "Macro: " & MacroTitle
    VerifyGeneratedMacroLoadable MacroTitle, "verifying the original restored after failed replacement"
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "RestoreOriginalAfterFailedReplacement", "Could not restore and verify the original macro after replacement failed." & vbCrLf & "Macro: " & MacroTitle & vbCrLf & "Backup: " & IIf(Len(Trim$(backupPath)) = 0, "(none)", backupPath) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Sub

Private Function ShouldKeepRollbackBackup(ByVal finalLoadSucceeded As Boolean, ByVal targetExisted As Boolean, ByVal backupPath As String) As Boolean
    If finalLoadSucceeded = False Then Exit Function
    If targetExisted = False Then Exit Function
    If Len(Trim$(backupPath)) = 0 Then Exit Function
    If mLauncherGenerationTrackingActive Then
        ShouldKeepRollbackBackup = True
        Exit Function
    End If
    If SetupRollbackTrackingActive() Then
        ShouldKeepRollbackBackup = True
        Exit Function
    End If
End Function

Private Function GeneratedMacroRollbackBackupReadable(ByVal backupPath As String, ByRef failureText As String) As Boolean
    On Error GoTo ErrHandler
    Dim fileNum As Integer
    backupPath = Trim$(backupPath)
    failureText = vbNullString
    If Len(backupPath) = 0 Then
        failureText = "Backup path was blank."
        Exit Function
    End If
    If Len(Dir$(backupPath)) = 0 Then
        failureText = "Backup file was not found. Path: " & backupPath
        Exit Function
    End If
    fileNum = FreeFile
    Open backupPath For Binary Access Read As #fileNum
    Close #fileNum
    fileNum = 0
    GeneratedMacroRollbackBackupReadable = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If fileNum <> 0 Then Close #fileNum
    On Error GoTo 0
    failureText = "Backup file could not be read. Path: " & backupPath & "; Error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription
End Function

Private Sub ProcessLauncherGenerationTracking()
    On Error GoTo ErrHandler
    Dim formName As String
    Dim addCount As Long
    Dim updateCount As Long
    Dim failureText As String
    Dim approvalResponse As VbMsgBoxResult
    Dim sqlText As String
    Dim roleItem As Variant
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim affectedForms As Collection
    Dim affectedRoles As Collection
    Dim cleanupWarnings As String
    If mLauncherGenerationTrackingActive = False Then Exit Sub
    RequireSameDatabaseUICommand "ProcessLauncherGenerationTracking", "Applying the atomic CodeDb launcher-form update after macro generation"
    If mLauncherGeneratedConfigIDs Is Nothing Then
        cleanupWarnings = CleanupGeneratedRollbackFiles()
        If Len(cleanupWarnings) > 0 Then CoreAddOperationNote cleanupWarnings
        Set mGeneratedMacroRollbackEntries = Nothing
        Set mLauncherFormRollbackEntries = Nothing
        mGeneratedMacroRollbackRecoveryRequired = False
        mLauncherGenerationTrackingActive = False
        Exit Sub
    End If
    If mLauncherGeneratedConfigIDs.count = 0 Then
        cleanupWarnings = CleanupGeneratedRollbackFiles()
        If Len(cleanupWarnings) > 0 Then CoreAddOperationNote cleanupWarnings
        Set mGeneratedMacroRollbackEntries = Nothing
        Set mLauncherFormRollbackEntries = Nothing
        mGeneratedMacroRollbackRecoveryRequired = False
        mLauncherGenerationTrackingActive = False
        Exit Sub
    End If
    ' INTENTIONAL FILTERED PARTITION: affectedForms, affectedRoles is a consumer-required subset/projection derived from the source enumerated by ProcessLauncherGenerationTracking; its membership is not
    ' independently authoritative.
    Set affectedForms = New Collection
    Set affectedRoles = New Collection
    For Each roleItem In LauncherDatabaseRoles()
        PumpSetupResponsiveness
        sqlText = TrackedLauncherConfigSql(CStr(roleItem))
        If SqlReturnsRows(sqlText) Then
            formName = LauncherFormNameForDatabaseRole(CStr(roleItem))
            If EnsureLauncherFormReady(formName) = False Then Err.Raise vbObjectError + 5907, "ProcessLauncherGenerationTracking", "Launcher form could not be prepared." & vbCrLf & "Form: " & formName
            On Error Resume Next
            affectedForms.Add formName, UCase$(formName)
            Err.Clear
            On Error GoTo ErrHandler
            affectedRoles.Add CStr(roleItem), UCase$(CStr(roleItem))
        End If
    Next roleItem
    If affectedForms.count = 0 Then
        cleanupWarnings = CleanupGeneratedRollbackFiles()
        If Len(cleanupWarnings) > 0 Then CoreAddOperationNote cleanupWarnings
        Set mGeneratedMacroRollbackEntries = Nothing
        Set mLauncherFormRollbackEntries = Nothing
        mGeneratedMacroRollbackRecoveryRequired = False
        mLauncherGenerationTrackingActive = False
        Exit Sub
    End If
    If PopupMessagesSuppressedForCurrentRun() Then
        approvalResponse = vbNo
    Else
        approvalResponse = CoreShowMessageResult("Macro generation is complete and launcher forms are ready for one atomic update." & vbCrLf & vbCrLf & "Affected forms:" & vbCrLf & SharedCollection_Join(affectedForms, vbCrLf) & vbCrLf & vbCrLf & "Any open affected form will be saved and closed. Every affected form will be snapshotted before the first launcher design change. Continue and save the full launcher-form batch?", vbQuestion Or vbYesNo Or vbDefaultButton2, "Custom Macro Launcher Automation")
    End If
    If approvalResponse <> vbYes Then Err.Raise vbObjectError + 5908, "ProcessLauncherGenerationTracking", "Launcher-form batch update was not approved. Generated macro and launcher-form recovery state will be retained or rolled back through the failure workflow."
    If SharedAccess_PrepareOpenObjectsForDesignChange(acForm, affectedForms, "applying the atomic Access Development Console launcher-form update", True) = False Then Err.Raise vbObjectError + 5909, "ProcessLauncherGenerationTracking", "Affected launcher forms could not be prepared for the atomic update."
    SnapshotLauncherForms affectedForms
    For Each roleItem In affectedRoles
        PumpSetupResponsiveness
        sqlText = TrackedLauncherConfigSql(CStr(roleItem))
        If SqlReturnsRows(sqlText) = False Then Err.Raise vbObjectError + 5950, "ProcessLauncherGenerationTracking", "Tracked launcher rows disappeared after the affected launcher-form batch was approved. Recovery state will be retained or rolled back through the failure workflow." & vbCrLf & "Role: " & LauncherRoleLabel(CStr(roleItem))
        addCount = 0
        updateCount = 0
        failureText = vbNullString
        formName = LauncherFormNameForDatabaseRole(CStr(roleItem))
        If OpenLauncherDesignForm(formName, False) = False Then Err.Raise vbObjectError + 5948, "ProcessLauncherGenerationTracking", "Launcher form could not be opened for the approved atomic update." & vbCrLf & "Form: " & formName
        EnsureCustomMacroLauncherAndAnchorCore formName
        ApplyLauncherButtonsFromSql formName, sqlText, addCount, updateCount, failureText
        If Len(Trim$(failureText)) > 0 Then Err.Raise vbObjectError + 5949, "ProcessLauncherGenerationTracking", failureText
        CloseLauncherDesignForm formName, True
    Next roleItem
    cleanupWarnings = CleanupGeneratedRollbackFiles()
    If Len(cleanupWarnings) > 0 Then CoreAddOperationNote cleanupWarnings
    Set mGeneratedMacroRollbackEntries = Nothing
    Set mLauncherFormRollbackEntries = Nothing
    mGeneratedMacroRollbackRecoveryRequired = False
    mLauncherGenerationTrackingActive = False
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Len(formName) > 0 Then
        If AccessFormIsOpen(formName) Then DoCmd.Close acForm, formName, acSaveNo
    End If
    On Error GoTo 0
    failureText = "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
    HandleLauncherFailureAfterGeneration failureText
End Sub

Private Sub HandleLauncherFailureAfterGeneration(ByVal failureText As String)
    On Error GoTo ErrHandler
    Dim response As VbMsgBoxResult
    Dim acceptResponse As VbMsgBoxResult
    Dim rollbackSummary As String
    Dim cleanupWarnings As String
    If PopupMessagesSuppressedForCurrentRun() Then
        response = vbCancel
    Else
        response = CoreShowMessageResult("Macro generation completed, but the atomic launcher-form update failed." & vbCrLf & vbCrLf & "Generated macros and one or more launcher forms may have changed." & vbCrLf & vbCrLf & "Yes = attempt verified rollback of both macros and launcher forms" & vbCrLf & "No = request explicit acceptance of the current partial state" & vbCrLf & "Cancel = retain all recovery state and backup files", vbExclamation Or vbYesNoCancel Or vbDefaultButton3, "Custom Macro Launcher Automation Failed")
    End If
    If response = vbYes Then
        rollbackSummary = RollbackGeneratedOperationFromRun()
        If RollbackSummaryHasFailures(rollbackSummary) Then
            mGeneratedMacroRollbackRecoveryRequired = True
            rollbackSummary = rollbackSummary & vbCrLf & "Rollback backup files and recovery entries were preserved. A later top-level operation will require retry or explicit discard before continuing."
        Else
            cleanupWarnings = CleanupGeneratedRollbackFiles()
            If Len(cleanupWarnings) > 0 Then rollbackSummary = rollbackSummary & vbCrLf & cleanupWarnings
            Set mGeneratedMacroRollbackEntries = Nothing
            Set mLauncherFormRollbackEntries = Nothing
            mGeneratedMacroRollbackRecoveryRequired = False
        End If
    ElseIf response = vbNo Then
        If PopupMessagesSuppressedForCurrentRun() Then
            acceptResponse = vbNo
        Else
            acceptResponse = CoreShowMessageResult("Explicitly accept the current generated macros and launcher forms?" & vbCrLf & vbCrLf & "This abandons rollback, deletes retained backup files, and treats the current partial state as authoritative." & vbCrLf & vbCrLf & GeneratedMacroRollbackPathsForMessage(), vbExclamation Or vbYesNo Or vbDefaultButton2, "Confirm Partial-State Acceptance")
        End If
        If acceptResponse = vbYes Then
            cleanupWarnings = CleanupGeneratedRollbackFiles()
            rollbackSummary = "Rollback was explicitly declined and the current macro/form state was explicitly accepted."
            If Len(cleanupWarnings) > 0 Then rollbackSummary = rollbackSummary & vbCrLf & cleanupWarnings
            Set mGeneratedMacroRollbackEntries = Nothing
            Set mLauncherFormRollbackEntries = Nothing
            mGeneratedMacroRollbackRecoveryRequired = False
        Else
            mGeneratedMacroRollbackRecoveryRequired = (GeneratedMacroRollbackEntryCount() > 0)
            rollbackSummary = "Partial-state acceptance was not explicitly confirmed. Recovery entries and backup files were retained."
        End If
    Else
        mGeneratedMacroRollbackRecoveryRequired = (GeneratedMacroRollbackEntryCount() > 0)
        rollbackSummary = "Rollback choice was cancelled or suppressed. Recovery entries and backup files were retained. Popup suppression never authorizes rollback-state discard."
    End If
    mLauncherGenerationTrackingActive = False
    On Error GoTo 0
    Err.Raise vbObjectError + 5910, "HandleLauncherFailureAfterGeneration", "Macro generation completed, but Custom Macro Launcher automation failed." & vbCrLf & vbCrLf & "Launcher failure:" & vbCrLf & failureText & vbCrLf & vbCrLf & rollbackSummary & vbCrLf & GeneratedMacroRollbackPathsForMessage()
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrSource As String
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrSource = Err.Source
    originalErrDescription = Err.Description
    mLauncherGenerationTrackingActive = False
    If GeneratedMacroRollbackEntryCount() > 0 Then mGeneratedMacroRollbackRecoveryRequired = True
    Err.Raise originalErrNumber, "HandleLauncherFailureAfterGeneration", "Custom Macro Launcher automation failed after macro generation. Rollback handling also failed or reported failure. Rollback recovery state and backup files were preserved." & vbCrLf & "Failure: " & failureText & vbCrLf & "Original source: " & originalErrSource & vbCrLf & "Err.Description: " & originalErrDescription & vbCrLf & GeneratedMacroRollbackPathsForMessage()
End Sub

Private Function RollbackSummaryHasFailures(ByVal rollbackSummary As String) As Boolean
    Dim lines As Variant
    Dim lineItem As Variant
    Dim lineText As String
    Dim colonPos As Long
    rollbackSummary = LCase$(rollbackSummary)
    lines = Split(rollbackSummary, vbCrLf)
    For Each lineItem In lines
        PumpSetupResponsiveness
        lineText = Trim$(CStr(lineItem))
        If Left$(lineText, Len("rollback failures:")) = "rollback failures:" Then
            colonPos = InStr(1, lineText, ":", vbBinaryCompare)
            If colonPos > 0 Then
                If CLng(Val(Trim$(Mid$(lineText, colonPos + 1)))) > 0 Then
                    RollbackSummaryHasFailures = True
                    Exit Function
                End If
            End If
        End If
    Next lineItem
    If InStr(1, rollbackSummary, "rollback failed", vbTextCompare) > 0 Then RollbackSummaryHasFailures = True
    If InStr(1, rollbackSummary, "restore failed", vbTextCompare) > 0 Then RollbackSummaryHasFailures = True
End Function

Private Function RollbackGeneratedOperationFromRun() As String
    Dim macroSummary As String
    Dim formSummary As String
    macroSummary = RollbackGeneratedMacroObjectsFromRun()
    formSummary = RollbackLauncherFormsFromRun()
    RollbackGeneratedOperationFromRun = macroSummary & vbCrLf & formSummary
    mGeneratedMacroRollbackRecoveryRequired = RollbackSummaryHasFailures(RollbackGeneratedOperationFromRun)
End Function

' Purpose: Performs generated result indicates problem by coordinating `End` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   resultText (Input; String): Receives the result text produced by the operation so the caller can inspect the outcome without re-running the work.
' Returns: Boolean: True when the generated result represents a warning, failure, invalid state, or other problem condition; False when the result is non-problematic.
Public Function GeneratedResultIndicatesProblem(ByVal resultText As String) As Boolean
    Select Case LCase$(Trim$(resultText))
        Case LCase$(GENERATED_RESULT_FAILED), LCase$("Invalid"), LCase$(GENERATED_RESULT_NEEDS_REVIEW)
            GeneratedResultIndicatesProblem = True
    End Select
End Function

Private Function CleanupGeneratedRollbackFiles() As String
    Dim item As Variant
    Dim entryData As Variant
    Dim backupPath As String
    Dim failedPaths As String
    If Not mGeneratedMacroRollbackEntries Is Nothing Then
        For Each item In mGeneratedMacroRollbackEntries
            PumpSetupResponsiveness
            entryData = item
            backupPath = vbNullString
            If IsArray(entryData) Then
                If UBound(entryData) >= 2 Then backupPath = CStr(entryData(2))
            End If
            DeleteRollbackBackupFile backupPath, failedPaths
        Next item
    End If
    If Not mLauncherFormRollbackEntries Is Nothing Then
        For Each item In mLauncherFormRollbackEntries
            PumpSetupResponsiveness
            entryData = item
            backupPath = vbNullString
            If IsArray(entryData) Then
                If UBound(entryData) >= 2 Then backupPath = CStr(entryData(2))
            End If
            DeleteRollbackBackupFile backupPath, failedPaths
        Next item
    End If
    If Len(failedPaths) > 0 Then CleanupGeneratedRollbackFiles = "Launcher-generation backup cleanup completed with warnings. The following files could not be removed:" & vbCrLf & failedPaths
End Function

Private Function TaggedButtonMapsToActiveGeneratedConfig(ByVal tagText As String, Optional ByVal expectedFormName As String = vbNullString) As Boolean
    On Error GoTo ErrHandler
    Dim configID As Long
    Dim MacroKey As String
    Dim sqlText As String
    Dim rs As DAO.Recordset
    Dim cfg As TMacroRecord
    configID = CLng(Val(LauncherTagValue(tagText, "ConfigID")))
    MacroKey = LauncherTagValue(tagText, "MacroKey")
    sqlText = "SELECT TOP 1 * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "] WHERE " & SQL_FIELD_ACTIVE & "<>0 AND Nz(" & SQL_FIELD_CONFIGURATION_STATUS & "," & SharedDao_SqlQuote(CONFIGURATION_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(CONFIGURATION_STATUS_ACTIVE) & " "
    If configID > 0 Then
        sqlText = sqlText & "AND " & SQL_FIELD_ID & "=" & CStr(configID)
    ElseIf Len(Trim$(MacroKey)) > 0 Then
        sqlText = sqlText & "AND " & SQL_FIELD_MACRO_KEY & "=" & SharedDao_SqlQuote(MacroKey)
    Else
        Exit Function
    End If
    Set rs = ADCDb().OpenRecordset(sqlText, dbOpenSnapshot)
    If rs.EOF = False Then
        MacroRecordFromRecordset rs, cfg
        If LauncherConfigRowQualifiesForButton(cfg) Then
            If Len(Trim$(expectedFormName)) = 0 Then
                TaggedButtonMapsToActiveGeneratedConfig = True
            Else
                TaggedButtonMapsToActiveGeneratedConfig = (StrComp(LauncherFormNameForDatabaseRole(cfg.ConfigOwnerDatabaseRole), expectedFormName, vbTextCompare) = 0)
            End If
        End If
    End If
ExitHere:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Exit Function
ErrHandler:
    TaggedButtonMapsToActiveGeneratedConfig = False
    Resume ExitHere
End Function

Private Sub CollectLauncherCleanupTargets(ByVal formName As String, ByVal staleButtons As Object, ByVal duplicateButtons As Object)
    On Error GoTo ErrHandler
    Dim ctl As control
    Dim tagText As String
    Dim keyText As String
    Dim keepers As Object
    Dim currentBest As String
    Dim currentRank As String
    Dim currentName As String
    ' INTENTIONAL KEYED LOOKUP: keepers provides keyed membership/lookup over source rows or state already read by CollectLauncherCleanupTargets; it is a derived access path, not a second authority.
    Set keepers = CreateObject("Scripting.Dictionary")
    For Each ctl In forms(formName).Controls
        PumpSetupResponsiveness
        If ctl.controlType = acCommandButton Then
            tagText = SharedText_Nz(ctl.Tag)
            If (StrComp(Left$(tagText, Len("ADCLauncher" & MODULE_TAG_DELIMITER)), "ADCLauncher" & MODULE_TAG_DELIMITER, vbTextCompare) = 0) Then
                keyText = LauncherButtonIdentityKeyFromTag(tagText)
                If Len(keyText) > 0 Then
                    If TaggedButtonMapsToActiveGeneratedConfig(tagText, formName) = False Then
                        If staleButtons.Exists(ctl.Name) = False Then staleButtons.Add ctl.Name, ctl.Name
                    End If
                    currentRank = Format$(CLng(ctl.Top), "0000000000") & MODULE_TAG_DELIMITER & Format$(CLng(ctl.Left), "0000000000") & MODULE_TAG_DELIMITER & ctl.Name
                    If keepers.Exists(keyText) = False Then
                        keepers.Add keyText, currentRank
                    Else
                        currentBest = CStr(keepers(keyText))
                        If StrComp(currentRank, currentBest, vbTextCompare) < 0 Then keepers(keyText) = currentRank
                    End If
                End If
            End If
        End If
    Next ctl
    For Each ctl In forms(formName).Controls
        PumpSetupResponsiveness
        If ctl.controlType = acCommandButton Then
            tagText = SharedText_Nz(ctl.Tag)
            If (StrComp(Left$(tagText, Len("ADCLauncher" & MODULE_TAG_DELIMITER)), "ADCLauncher" & MODULE_TAG_DELIMITER, vbTextCompare) = 0) Then
                keyText = LauncherButtonIdentityKeyFromTag(tagText)
                If Len(keyText) > 0 Then
                    currentName = LauncherKeeperNameFromRank(CStr(keepers(keyText)))
                    If StrComp(ctl.Name, currentName, vbTextCompare) <> 0 Then
                        If duplicateButtons.Exists(ctl.Name) = False Then duplicateButtons.Add ctl.Name, ctl.Name
                    End If
                End If
            End If
        End If
    Next ctl
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "CollectLauncherCleanupTargets", "Could not collect stale/duplicate launcher cleanup targets." & vbCrLf & "Form: " & formName & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

Private Function DeleteLauncherControlsByDictionary(ByVal formName As String, ByVal controlNames As Object) As Long
    On Error GoTo ErrHandler
    Dim key As Variant
    Dim deletedCount As Long
    For Each key In controlNames.keys
        PumpSetupResponsiveness
        If ControlExistsOnOpenForm(formName, CStr(key)) Then
            DeleteControl formName, CStr(key)
            deletedCount = deletedCount + 1
        End If
    Next key
    DeleteLauncherControlsByDictionary = deletedCount
    Exit Function
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "DeleteLauncherControlsByDictionary", "Could not delete launcher controls." & vbCrLf & "Form: " & formName & vbCrLf & "Err.Description: " & originalErrDescription
End Function

Private Sub ValidateMacroConfig(ByRef cfg As TMacroRecord)
    If Len(Trim$(cfg.MacroKey)) = 0 Then
        Err.Raise vbObjectError + 5201, "ValidateMacroConfig", "MacroKey is blank." & vbCrLf & "Configuration row ID: " & CStr(cfg.ID) & vbCrLf & "MacroTitle: " & IIf(Len(Trim$(cfg.MacroTitle)) = 0, "(blank)", cfg.MacroTitle) & vbCrLf & "MacroAction: " & IIf(Len(Trim$(cfg.MacroAction)) = 0, "(blank)", cfg.MacroAction) & vbCrLf & "Populate MacroKey before generation or runtime execution."
    End If
    If ConfigurationStatusBlocksGeneration(cfg.ConfigurationStatus) Then
        Err.Raise vbObjectError + 5230, "ValidateMacroConfig", "ConfigurationStatus prevents generation or runtime execution." & vbCrLf & "MacroKey: " & cfg.MacroKey & vbCrLf & "ConfigurationStatus: " & cfg.ConfigurationStatus
    End If
    If IsValidConfigOwnerDatabaseRole(cfg.ConfigOwnerDatabaseRole) = False Then
        Err.Raise vbObjectError + 5231, "ValidateMacroConfig", "ConfigOwnerDatabaseRole is blank or unsupported. Manually set HostDb or CodeDb before generation/runtime execution." & vbCrLf & "MacroKey: " & cfg.MacroKey & vbCrLf & "ConfigOwnerDatabaseRole: " & IIf(Len(Trim$(cfg.ConfigOwnerDatabaseRole)) = 0, "(blank)", cfg.ConfigOwnerDatabaseRole)
    End If
    If Len(Trim$(cfg.MacroTitle)) = 0 Then
        Err.Raise vbObjectError + 5202, "ValidateMacroConfig", "MacroTitle is blank for MacroKey: " & cfg.MacroKey & vbCrLf & "Configuration row ID: " & CStr(cfg.ID) & vbCrLf & "MacroAction: " & cfg.MacroAction & vbCrLf & "Populate MacroTitle with the Access macro object name to create."
    End If
    EnsureValidAccessMacroObjectName cfg.MacroTitle, "ValidateMacroConfig"
    If StrComp(cfg.MacroKey, CONFIRM_AUTOEXEC, vbTextCompare) = 0 And StrComp(cfg.MacroTitle, CONFIRM_AUTOEXEC, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 5209, "ValidateMacroConfig", "MacroKey is AutoExec, but MacroTitle is not AutoExec." & vbCrLf & "Access startup runs only the macro object named AutoExec." & vbCrLf & "MacroKey: " & cfg.MacroKey & vbCrLf & "MacroTitle: " & cfg.MacroTitle
    End If
    Select Case UCase$(Trim$(cfg.MacroAction))
        Case UCase$(ACTION_RUNCODE)
            If Len(Trim$(cfg.TargetProcedureName)) = 0 Then
                Err.Raise vbObjectError + 5203, "ValidateMacroConfig", "RunCode target procedure is blank for macro: " & cfg.MacroTitle & vbCrLf & "Populate TargetProcedureName. TargetExpression is not supported."
            End If
            ValidateTargetProcedureMetadata cfg
            EnsureConfiguredTargetCanGenerate cfg
        Case UCase$(ACTION_OPENFORM), UCase$(ACTION_OPENTABLE), UCase$(ACTION_OPENQUERY), UCase$(ACTION_OPENREPORT), UCase$(ACTION_RUNMACRO)
            If Len(ConfiguredObjectActionTargetName(cfg)) = 0 Then
                Err.Raise vbObjectError + 5204, "ValidateMacroConfig", cfg.MacroAction & " target object name is blank for macro: " & cfg.MacroTitle & vbCrLf & "Populate TargetObjectName on the accepted MC_Records configuration. Candidate proposal/evidence values are stored separately in MC_Candidates before acceptance."
            End If
            EnsureConfiguredTargetCanGenerate cfg
        Case Else
            Err.Raise vbObjectError + 5205, "ValidateMacroConfig", "MacroAction must be RunCode, OpenForm, OpenTable, OpenQuery, OpenReport, or RunMacro." & vbCrLf & "MacroKey: " & cfg.MacroKey & vbCrLf & "MacroAction: " & cfg.MacroAction
    End Select
End Sub

Private Function LoadMacroConfigByKey(ByVal MacroKey As String, ByRef cfg As TMacroRecord) As Boolean
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    RequireConfigTableForUtilityWorkflow "LoadMacroConfigByKey"
    Set db = ADCDb()
    Set rs = db.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "] WHERE " & SQL_FIELD_MACRO_KEY & "=" & SharedDao_SqlQuote(MacroKey), dbOpenSnapshot)
    If Not rs.EOF Then
        MacroRecordFromRecordset rs, cfg
        LoadMacroConfigByKey = True
    End If
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Exit Function
ErrHandler:
    Dim originalErrNumber As Long
    originalErrNumber = Err.Number
    Err.Raise originalErrNumber, "LoadMacroConfigByKey", FormatErrorWithContext("Could not load a macro-creator config row by MacroKey.", "MC_Records table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & vbCrLf & "MacroKey: " & MacroKey & vbCrLf & "Expected storage: MC_Records configuration row")
End Function

Private Sub SetAccessMacroObjectDescription(ByVal MacroTitle As String, ByVal descriptionText As String)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim doc As DAO.Document
    Dim physicalDescription As String
    Dim wasBounded As Boolean
    Dim authoritativeCharacterLength As Long
    Dim authoritativeByteLength As Long
    Dim writtenCharacterLength As Long
    Dim writtenByteLength As Long
    If Len(MacroTitle) = 0 Then Exit Sub
    If Len(descriptionText) = 0 Then Exit Sub
    physicalDescription = ManagedDescriptionExpectedPhysicalText(descriptionText)
    authoritativeCharacterLength = Len(descriptionText)
    authoritativeByteLength = LenB(descriptionText)
    writtenCharacterLength = Len(physicalDescription)
    writtenByteLength = LenB(physicalDescription)
    wasBounded = (StrComp(physicalDescription, descriptionText, vbBinaryCompare) <> 0)
    Set db = GetHostDb()
    Set doc = db.Containers(MACRO_DAO_CONTAINER_NAME).Documents(MacroTitle)
    If TextPropertyMatches(doc, DAO_DESCRIPTION_PROPERTY_NAME, physicalDescription) Then Exit Sub
    If SetupDomainMutationWritesAllowedForCurrentPart(SETUP_DOMAIN_G1_DESCRIPTIONS) = False Then Err.Raise vbObjectError + 9625, "SetAccessMacroObjectDescription", "Managed DAO macro Description mutation is owned exclusively by Setup G1 during numbered Setup." & vbCrLf & "Macro: " & MacroTitle
    If EnsureDaoPropertyValue(doc, DAO_DESCRIPTION_PROPERTY_NAME, dbText, physicalDescription, "SetAccessMacroObjectDescription") = False Then Err.Raise vbObjectError + 6454, "SetAccessMacroObjectDescription", "The macro Description property writer returned False after a required change."
    If wasBounded Then CoreAddOperationNote "Applied a deterministic property-safe macro Description representation." & vbCrLf & "MacroTitle: " & MacroTitle & vbCrLf & "Authoritative characters: " & CStr(authoritativeCharacterLength) & vbCrLf & "Authoritative bytes: " & CStr(authoritativeByteLength) & vbCrLf & "Written characters: " & CStr(writtenCharacterLength) & vbCrLf & "Written bytes: " & CStr(writtenByteLength)
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim daoErrorsText As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    daoErrorsText = ReviewDBEngineErrorsText()
    CoreAddOperationNote "Could not apply required Access macro Description metadata." & vbCrLf & "MacroTitle: " & MacroTitle & vbCrLf & "Property: " & DAO_DESCRIPTION_PROPERTY_NAME & vbCrLf & "Authoritative characters: " & CStr(authoritativeCharacterLength) & vbCrLf & "Authoritative bytes: " & CStr(authoritativeByteLength) & vbCrLf & "Written characters: " & CStr(writtenCharacterLength) & vbCrLf & "Written bytes: " & CStr(writtenByteLength) & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription & IIf(Len(daoErrorsText) > 0, vbCrLf & "DBEngine.Errors:" & vbCrLf & daoErrorsText, vbNullString)
    Err.Raise savedErrNumber, "SetAccessMacroObjectDescription", FormatErrorWithContext("Could not apply the required Access macro Description metadata.", "MacroTitle: " & MacroTitle & vbCrLf & "Property: " & DAO_DESCRIPTION_PROPERTY_NAME & vbCrLf & "Original Err.Source: " & savedErrSource & vbCrLf & "Original Err.Description: " & savedErrDescription & IIf(Len(daoErrorsText) > 0, vbCrLf & "DBEngine.Errors:" & vbCrLf & daoErrorsText, vbNullString))
End Sub

Private Sub CreateOrReplaceRunCodeMacro(ByRef cfg As TMacroRecord)
    On Error GoTo ErrHandler
    Dim macroText As String
    EnsureValidAccessMacroObjectName cfg.MacroTitle, "CreateOrReplaceRunCodeMacro"
    macroText = RunCodeMacroText(cfg)
    CreateOrReplaceMacroFromText cfg.MacroTitle, macroText, cfg.navigationPaneGroupName, "CreateOrReplaceRunCodeMacro", ConfigObjectDescription(cfg)
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "CreateOrReplaceRunCodeMacro", "Could not create Access RunCode macro." & vbCrLf & "MacroKey: " & cfg.MacroKey & vbCrLf & "MacroTitle: " & cfg.MacroTitle & vbCrLf & "RunCode target: " & EffectiveTargetCallText(cfg) & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

Private Sub CreateOrReplaceMacroFromText(ByVal MacroTitle As String, ByVal macroText As String, ByVal navPaneGroup As String, ByVal sourceName As String, Optional ByVal descriptionText As String = vbNullString, Optional ByVal forceRefreshNavPaneShortcut As Boolean = False, Optional ByVal skipNavigationPaneGrouping As Boolean = False)
    On Error GoTo ErrHandler
    Dim tempPath As String
    Dim originalBackupPath As String
    Dim backupDeleteFailures As String
    Dim fileNum As Integer
    Dim targetExisted As Boolean
    Dim targetDeleted As Boolean
    Dim finalLoadSucceeded As Boolean
    Dim postCommitWarningText As String
    EnsureValidAccessMacroObjectName MacroTitle, sourceName
    tempPath = Environ$("TEMP") & "\AccessMacro_" & SharedName_SafeFileName(MacroTitle, ADCAT_OBJECT_TYPE_MACRO, 120) & "_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & Format$(Timer * 1000, "000000") & ".txt"
    fileNum = FreeFile
    Open tempPath For Output As #fileNum
    Print #fileNum, macroText
    Close #fileNum
    fileNum = 0
    targetExisted = AccessObjectExistsInCurrentDatabase(acMacro, MacroTitle)
    If targetExisted Then
        If SharedAccess_PrepareOpenObjectForDesignChange(acMacro, MacroTitle, "replacing generated macro [" & MacroTitle & "]") = False Then
            MarkSystemCommandCancelled
            Err.Raise vbObjectError + 6289, "CreateOrReplaceMacroFromText", "Macro replacement was cancelled because the affected open macro was not approved for save and close."
        End If
        originalBackupPath = Environ$("TEMP") & "\AccessMacro_Backup_" & SharedName_SafeFileName(MacroTitle, ADCAT_OBJECT_TYPE_MACRO, 120) & "_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & Format$(Timer * 1000, "000000") & ".txt"
        'Back up the existing macro as text before touching it. This is safer than renaming Access macro objects during setup, because DoCmd.Rename can raise run-time error 5 for some macro-object names/states.
        Application.SaveAsText acMacro, MacroTitle, originalBackupPath
        DoCmd.DeleteObject acMacro, MacroTitle
        targetDeleted = True
    End If
    'Load the replacement directly under the final macro object name.
    'The original object has already been backed up to text, so the ErrHandler can restore it.
    Application.LoadFromText acMacro, MacroTitle, tempPath
    finalLoadSucceeded = True
    TrackGeneratedMacroRollback MacroTitle, targetExisted, originalBackupPath
    TrackSetupAccessObjectRollback SETUP_ROLLBACK_RESTORE_MACRO_ACTION, MacroTitle, targetExisted, originalBackupPath
    On Error Resume Next
    SetAccessMacroObjectDescription MacroTitle, descriptionText
    If Err.Number <> 0 Then
        postCommitWarningText = "Description update failed. Error " & CStr(Err.Number) & ": " & Err.Description
        Err.Clear
    End If
    If skipNavigationPaneGrouping = False Then
        AddMacroToConfiguredNavigationPaneGroup MacroTitle, navPaneGroup, forceRefreshNavPaneShortcut
        If Err.Number <> 0 Then
            If Len(postCommitWarningText) > 0 Then postCommitWarningText = postCommitWarningText & vbCrLf
            postCommitWarningText = postCommitWarningText & "Navigation Pane grouping failed. Error " & CStr(Err.Number) & ": " & Err.Description
            Err.Clear
        End If
    End If
    If Len(postCommitWarningText) > 0 Then
        Debug.Print "Access Development Console post-commit metadata error for [" & MacroTitle & "]:" & vbCrLf & postCommitWarningText
        CoreAddOperationNote "Macro [" & MacroTitle & "] was replaced successfully, but required post-commit metadata work failed: " & Replace(postCommitWarningText, vbCrLf, " | ")
        LogTraceMessage "Error", "Generation", sourceName, "Generated macro replacement committed, but required post-commit metadata work failed.", postCommitWarningText, "Apply macro description and Navigation Pane grouping", 0, vbNullString, vbNullString, MacroTitle, ADCAT_OBJECT_TYPE_MACRO, "CodeDb", "Generation"
        On Error GoTo ErrHandler
        Err.Raise vbObjectError + 6455, sourceName, "Access macro replacement committed, but required post-commit metadata work failed." & vbCrLf & "MacroTitle: " & MacroTitle & vbCrLf & postCommitWarningText
    End If
    On Error GoTo ErrHandler
    On Error Resume Next
    If fileNum <> 0 Then Close #fileNum
    If Len(tempPath) > 0 Then
        If Len(Dir$(tempPath)) > 0 Then Kill tempPath
    End If
    If Len(originalBackupPath) > 0 Then
        If ShouldKeepRollbackBackup(finalLoadSucceeded, targetExisted, originalBackupPath) = False Then
            DeleteRollbackBackupFile originalBackupPath, backupDeleteFailures
        End If
    End If
    Exit Sub
ErrHandler:
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim rollbackDetail As String
    Dim rollbackRestoredSuccessfully As Boolean
    Dim rollbackErrNumber As Long
    Dim rollbackErrSource As String
    Dim rollbackErrDescription As String
    Dim retentionErrNumber As Long
    Dim retentionErrSource As String
    Dim retentionErrDescription As String
    Dim recoveryEntryRetained As Boolean
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    On Error Resume Next
    If fileNum <> 0 Then Close #fileNum
    On Error GoTo 0
    If finalLoadSucceeded = False And targetExisted Then
        On Error Resume Next
        RestoreOriginalAfterFailedReplacement MacroTitle, originalBackupPath
        rollbackErrNumber = Err.Number
        rollbackErrSource = Err.Source
        rollbackErrDescription = Err.Description
        Err.Clear
        On Error GoTo 0
        If rollbackErrNumber = 0 Then
            rollbackRestoredSuccessfully = True
            rollbackDetail = "Rollback restored and verified the original macro object from text backup."
        Else
            rollbackDetail = "Rollback failed while restoring or verifying the original macro object. Rollback Err.Number: " & CStr(rollbackErrNumber) & "; Rollback Err.Source: " & rollbackErrSource & "; Rollback Err.Description: " & rollbackErrDescription
            On Error Resume Next
            RetainFailedMacroReplacementRecovery MacroTitle, originalBackupPath
            retentionErrNumber = Err.Number
            retentionErrSource = Err.Source
            retentionErrDescription = Err.Description
            recoveryEntryRetained = (retentionErrNumber = 0)
            Err.Clear
            On Error GoTo 0
            If recoveryEntryRetained Then
                rollbackDetail = rollbackDetail & vbCrLf & "Structured recovery state was retained for the macro and backup path."
            Else
                mGeneratedMacroRollbackRecoveryRequired = True
                rollbackDetail = rollbackDetail & vbCrLf & "Structured recovery-state registration also failed. Retention Err.Number: " & CStr(retentionErrNumber) & "; Retention Err.Source: " & retentionErrSource & "; Retention Err.Description: " & retentionErrDescription
            End If
        End If
    ElseIf finalLoadSucceeded Then
        rollbackDetail = "Final replacement had already succeeded before a later operation failed."
    Else
        rollbackDetail = "No original macro object existed before replacement, so no rollback target was needed."
    End If
    On Error Resume Next
    If Len(tempPath) > 0 Then
        If Len(Dir$(tempPath)) > 0 Then Kill tempPath
    End If
    If Len(originalBackupPath) > 0 Then
        If Len(Dir$(originalBackupPath)) > 0 Then
            If rollbackErrNumber <> 0 Then
                rollbackDetail = rollbackDetail & vbCrLf & "Original backup was preserved for retained recovery: " & originalBackupPath
            ElseIf finalLoadSucceeded Then
                If ShouldKeepRollbackBackup(finalLoadSucceeded, targetExisted, originalBackupPath) = False Then DeleteRollbackBackupFile originalBackupPath, backupDeleteFailures
            ElseIf rollbackRestoredSuccessfully Then
                DeleteRollbackBackupFile originalBackupPath, backupDeleteFailures
            End If
        End If
    End If
    If Len(backupDeleteFailures) > 0 Then rollbackDetail = rollbackDetail & vbCrLf & "Rollback backup cleanup warning(s):" & vbCrLf & backupDeleteFailures
    On Error GoTo 0
    If rollbackErrNumber <> 0 Then
        On Error Resume Next
        CoreAddOperationNote "Destructive macro replacement recovery failed for [" & MacroTitle & "]. Recovery state retained: " & CStr(recoveryEntryRetained) & ". Backup: " & IIf(Len(Trim$(originalBackupPath)) = 0, "(none)", originalBackupPath)
        LogTraceMessage "Error", "Generation", sourceName, "Macro replacement failed and the original macro could not be restored safely.", rollbackDetail & vbCrLf & "Original replacement Err.Number: " & CStr(errNumber) & vbCrLf & "Original replacement Err.Source: " & errSource & vbCrLf & "Original replacement Err.Description: " & errDescription, "Restore original macro after failed replacement", rollbackErrNumber, rollbackErrSource, rollbackErrDescription, MacroTitle, ADCAT_OBJECT_TYPE_MACRO, "CodeDb", "Generation"
        On Error GoTo 0
        Err.Raise MACRO_REPLACEMENT_RECOVERY_FAILURE_ERROR_NUMBER, sourceName, "Destructive macro replacement recovery failed. The original macro may be missing or unusable, the current batch must stop, and retained recovery state must be resolved before another top-level operation." & vbCrLf & "MacroTitle: " & MacroTitle & vbCrLf & "Original backup path: " & IIf(Len(Trim$(originalBackupPath)) = 0, "(none)", originalBackupPath) & vbCrLf & "Structured recovery entry retained: " & CStr(recoveryEntryRetained) & vbCrLf & "Rollback Err.Number: " & CStr(rollbackErrNumber) & vbCrLf & "Rollback Err.Source: " & rollbackErrSource & vbCrLf & "Rollback Err.Description: " & rollbackErrDescription & vbCrLf & "Original replacement Err.Number: " & CStr(errNumber) & vbCrLf & "Original replacement Err.Source: " & errSource & vbCrLf & "Original replacement Err.Description: " & errDescription & vbCrLf & _
        GeneratedMacroRollbackPathsForMessage()
    End If
    Err.Raise errNumber, sourceName, "Could not safely create or replace Access macro object." & vbCrLf & "MacroTitle: " & MacroTitle & vbCrLf & "Replacement text path: " & IIf(Len(Trim$(tempPath)) = 0, "(none)", tempPath) & vbCrLf & "Original backup path: " & IIf(Len(Trim$(originalBackupPath)) = 0, "(none)", originalBackupPath) & vbCrLf & "Target existed before replacement: " & CStr(targetExisted) & vbCrLf & "Target deleted before final load: " & CStr(targetDeleted) & vbCrLf & "Final load succeeded: " & CStr(finalLoadSucceeded) & vbCrLf & "Rollback status: " & rollbackDetail & vbCrLf & "Original Err.Number: " & CStr(errNumber) & vbCrLf & "Original Err.Source: " & errSource & vbCrLf & "Original Err.Description: " & errDescription
End Sub

Private Function RunCodeMacroText(ByRef cfg As TMacroRecord) As String
    Dim runCodeArgument As String
    Dim runDispatcher As String
    Dim confirmedDispatcher As String
    runDispatcher = "RunConfiguredMacroTarget"
    confirmedDispatcher = "RunConfirmedConfiguredMacroTarget"
    If GeneratedRunCodeMacroRequiresConfirmation(cfg) Then
        runCodeArgument = confirmedDispatcher & "(" & SharedText_DoubleQuote(cfg.MacroKey) & ")"
    Else
        runCodeArgument = runDispatcher & "(" & SharedText_DoubleQuote(cfg.MacroKey) & ")"
    End If
    RunCodeMacroText = "Version =196611" & vbCrLf & "ColumnsShown =0" & vbCrLf & "Begin" & vbCrLf & "    Action =""RunCode""" & vbCrLf & "    Argument =""" & EscapeMacroText(runCodeArgument) & """" & vbCrLf & "End" & vbCrLf
End Function


Private Sub EnsureValidAccessMacroObjectName(ByVal objectName As String, ByVal sourceName As String)
    If Len(objectName) = 0 Then
        Err.Raise vbObjectError + 5600, sourceName, "Access macro object name is blank." & vbCrLf & "Source: " & sourceName & vbCrLf & "Access macro object names must be nonblank before generation or replacement."
    End If
    If Len(objectName) > MAX_ACCESS_OBJECT_NAME_LENGTH Then
        Err.Raise vbObjectError + 5601, sourceName, "Access macro object name exceeds the limit." & vbCrLf & "Name: " & objectName & vbCrLf & "Limit: " & CStr(MAX_ACCESS_OBJECT_NAME_LENGTH)
    End If
End Sub

' ===============================================================================================================================================================================
' CONFIGURATION ROW CLEANUP / TARGET-RESOLUTION PROCEDURES
' ===============================================================================================================================================================================

Private Function LoadConfigByID(ByVal rowID As Long, ByRef cfg As TMacroRecord) As Boolean
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    If rowID <= 0 Then Exit Function
    Set rs = ADCDb().OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "] WHERE " & SQL_FIELD_ID & "=" & CStr(rowID), dbOpenSnapshot)
    If rs.EOF Then GoTo ExitHere
    MacroRecordFromRecordset rs, cfg
    LoadConfigByID = True
ExitHere:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Exit Function
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
    Err.Raise savedErrNumber, "LoadConfigByID", "Could not load Access Development Console configuration row ID " & CStr(rowID) & "." & vbCrLf & "Table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Sub DeleteConfigRowByID(ByVal rowID As Long)
    On Error GoTo ErrHandler
    If rowID <= 0 Then Exit Sub
    ADCDb().Execute "DELETE FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "] WHERE " & SQL_FIELD_ID & "=" & CStr(rowID), dbFailOnError
    InvalidateDuplicateReviewCache
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "DeleteConfigRowByID", "Could not delete macro-creator config row." & vbCrLf & "ID: " & CStr(rowID) & vbCrLf & "Table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

Private Function ConfigObjectDescription(ByRef cfg As TMacroRecord) As String
    Dim tooltipText As String
    tooltipText = Trim$(cfg.ObjectDescription)
    If Len(tooltipText) = 0 Then tooltipText = Trim$(cfg.ActionDescription)
    If Len(tooltipText) = 0 Then tooltipText = Trim$(cfg.LauncherTooltip)
    If Len(tooltipText) = 0 Then
        tooltipText = "Generated Access Development Console shortcut for configured row [" & cfg.MacroKey & "]."
        If Len(Trim$(cfg.MacroAction)) > 0 Then tooltipText = tooltipText & " Macro action: " & cfg.MacroAction & "."
        If Len(Trim$(cfg.ConfigOwnerDatabaseRole)) > 0 Then tooltipText = tooltipText & " ConfigOwnerDatabaseRole: " & cfg.ConfigOwnerDatabaseRole & "."
        If Len(Trim$(cfg.TargetProcedureName)) > 0 Then tooltipText = tooltipText & " Target procedure: " & cfg.TargetProcedureName & "."
        If Len(Trim$(cfg.TargetObjectName)) > 0 Then tooltipText = tooltipText & " Target object: " & cfg.TargetObjectName & "."
        tooltipText = tooltipText & " Generated macro objects route through Access Development Console runtime validation, confirmation, and target-resolution logic."
    End If
    ConfigObjectDescription = tooltipText
End Function

' Purpose: Deletes user-selected duplicate configuration rows through the managed duplicate-review workflow, including confirmation, backup/rollback handling, progress accounting, and result reporting.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub DeleteSelectedDuplicateRows(Optional ByVal showMessage As Boolean = True)
    On Error GoTo ErrHandler
    Dim reviewMode As VbMsgBoxResult
    Dim selectedCount As Long
    Dim riskyGroupCount As Long
    Dim backupCount As Long
    Dim deletedCount As Long
    Dim skippedCount As Long
    Dim remainingCount As Long
    Dim cancelRequested As Boolean
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    RequireDefaultCodeDbGeneratedMacrosNavigationPaneGroupForRoleContext "DeleteSelectedDuplicateRows", False
    RequireConfigTableForUtilityWorkflow "DeleteSelectedDuplicateRows"
    RepairDuplicateReviewQueryForSetup
    InvalidateDuplicateReviewCache
    selectedCount = CountDuplicateRowsBySelectedState(True)
    If selectedCount = 0 Then
        If showMessage Then
            CoreShowMessage "No duplicate-review rows are currently selected for deletion." & vbCrLf & vbCrLf & "Open " & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Config_DuplicateRows") & " and check SelectedForDuplicateDeletion for rows you want to remove.", vbInformation, "Delete Selected Duplicate Rows"
        End If
        Exit Sub
    End If
    reviewMode = CoreShowMessageResult("You are about to permanently delete rows selected with SelectedForDuplicateDeletion=True from the Runner Macro MC_Records table." & vbCrLf & vbCrLf & "Only rows that still appear in " & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Config_DuplicateRows") & " are eligible." & vbCrLf & "A backup will be created before any row is deleted." & vbCrLf & vbCrLf & "Rows currently selected for duplicate deletion: " & CStr(selectedCount) & vbCrLf & vbCrLf & "Yes = Review each selected duplicate before deleting." & vbCrLf & "No = Skip item-by-item review and delete all selected duplicate rows." & vbCrLf & "Cancel = Cancel duplicate deletion entirely.", vbExclamation Or vbYesNoCancel Or vbDefaultButton3, "Delete Selected Duplicate Rows")
    If reviewMode = vbCancel Then
        MarkSystemCommandCancelled
        Exit Sub
    End If
    riskyGroupCount = CountAllRowsSelectedDuplicateGroups()
    If riskyGroupCount > 0 Then
        If ConfirmAllRowsSelectedDuplicateGroups(riskyGroupCount) = False Then
            MarkSystemCommandCancelled
            Exit Sub
        End If
    End If
    backupCount = BackupAllUserDataCore(False)
    If reviewMode = vbYes Then
        DeleteSelectedDuplicateRowsWithReview deletedCount, skippedCount, remainingCount, cancelRequested
    Else
        deletedCount = DeleteSelectedDuplicateRowsWithoutReview()
        remainingCount = CountDuplicateRowsBySelectedState(True)
    End If
    If cancelRequested Then
        MarkSystemCommandCancelled
        CoreAddOperationNote "Duplicate deletion stopped during item-by-item review after partial work. Rows deleted: " & CStr(deletedCount) & "; rows skipped: " & CStr(skippedCount) & "; selected rows remaining: " & CStr(remainingCount) & "."
        If showMessage Then CoreShowMessage "Duplicate deletion was cancelled during item-by-item review." & vbCrLf & vbCrLf & "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB & vbCrLf & "MC_Records table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & vbCrLf & "Selected duplicate rows at start: " & CStr(selectedCount) & vbCrLf & "Backup table copies created: " & CStr(backupCount) & vbCrLf & "Rows deleted before cancellation: " & CStr(deletedCount) & vbCrLf & "Rows skipped during review: " & CStr(skippedCount) & vbCrLf & "Selected rows still remaining: " & CStr(remainingCount), vbExclamation, "Duplicate Deletion Cancelled"
    ElseIf showMessage Then
        CoreShowMessage "Duplicate deletion complete." & vbCrLf & vbCrLf & "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB & vbCrLf & "MC_Records table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & vbCrLf & "Selected duplicate rows at start: " & CStr(selectedCount) & vbCrLf & "Backup table copies created: " & CStr(backupCount) & vbCrLf & "Rows deleted: " & CStr(deletedCount) & vbCrLf & "Rows skipped during review: " & CStr(skippedCount) & vbCrLf & "Selected rows still remaining: " & CStr(remainingCount), vbInformation, "Delete Selected Duplicate Rows"
    End If
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If showMessage Then
        CoreShowErrorMessage "Could not delete selected duplicate rows." & vbCrLf & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, vbCritical, "Delete Selected Duplicate Rows", "DeleteSelectedDuplicateRows", "Review"
        Exit Sub
    End If
    Err.Raise savedErrNumber, "DeleteSelectedDuplicateRows", "Could not delete selected duplicate rows." & vbCrLf & "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB & vbCrLf & "MC_Records table: " & SafeTableName("ADC_Config") & vbCrLf & "Query: " & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Config_DuplicateRows") & vbCrLf & "Original Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub DeleteSelectedDuplicateRowsWithReview(ByRef deletedCount As Long, ByRef skippedCount As Long, ByRef remainingCount As Long, ByRef cancelRequested As Boolean)
    On Error GoTo ErrHandler
    Dim ids As Collection
    Dim item As Variant
    Dim rowID As Long
    Dim reviewedCount As Long
    Dim response As VbMsgBoxResult
    Set ids = SelectedDuplicateRowIDsSnapshot()
    For Each item In ids
        PumpSetupResponsiveness
        rowID = CLng(item)
        reviewedCount = reviewedCount + 1
        PulseProgressEvery reviewedCount, 0, "Reviewing duplicate row ID " & CStr(rowID)
        If SelectedRowStillEligibleForDuplicateDelete(rowID) Then
            response = ConfirmDeleteOneDuplicateRow(rowID)
            Select Case response
                Case vbYes
                    DeleteConfigRowByID rowID
                    deletedCount = deletedCount + 1
                Case vbNo
                    skippedCount = skippedCount + 1
                Case vbCancel
                    cancelRequested = True
                    Exit For
            End Select
        Else
            skippedCount = skippedCount + 1
        End If
    Next item
    remainingCount = CountDuplicateRowsBySelectedState(True)
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "DeleteSelectedDuplicateRowsWithReview", "Could not complete one-at-a-time duplicate deletion review." & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

Private Function DeleteSelectedDuplicateRowsWithoutReview() As Long
    On Error GoTo ErrHandler
    Dim ids As Collection
    Dim item As Variant
    Dim rowID As Long
    Set ids = SelectedDuplicateRowIDsSnapshot()
    For Each item In ids
        PumpSetupResponsiveness
        rowID = CLng(item)
        If SelectedRowStillEligibleForDuplicateDelete(rowID) Then
            DeleteConfigRowByID rowID
            DeleteSelectedDuplicateRowsWithoutReview = DeleteSelectedDuplicateRowsWithoutReview + 1
        End If
    Next item
    Exit Function
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "DeleteSelectedDuplicateRowsWithoutReview", "Could not delete selected duplicate rows without item-by-item review." & vbCrLf & "Err.Description: " & originalErrDescription
End Function

Private Function SelectedDuplicateRowIDsSnapshot() As Collection
    On Error GoTo ErrHandler
    Dim ids As Collection
    Dim rs As DAO.Recordset
    Dim cfg As TMacroRecord
    ' INTENTIONAL OPERATION SNAPSHOT: ids captures one coherent source state for this invocation of SelectedDuplicateRowIDsSnapshot; validity ends with the operation rather than becoming a live data authority.
    Set ids = New Collection
    Set rs = ADCDb().OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "] ORDER BY " & SQL_FIELD_ID, dbOpenSnapshot)
    Do While Not rs.EOF
        PumpSetupResponsiveness
        MacroRecordFromRecordset rs, cfg
        If cfg.SelectedForDuplicateDeletion Then
            If Len(DuplicateIssueTypeForQuery(cfg.ID)) > 0 Then ids.Add cfg.ID
        End If
        rs.MoveNext
    Loop
ExitHere:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set SelectedDuplicateRowIDsSnapshot = ids
    Exit Function
ErrHandler:
    ' INTENTIONAL OPERATION SNAPSHOT: ids captures one coherent source state for this invocation of SelectedDuplicateRowIDsSnapshot; validity ends with the operation rather than becoming a live data authority.
    Set ids = New Collection
    Resume ExitHere
End Function

Private Function SelectedRowStillEligibleForDuplicateDelete(ByVal rowID As Long) As Boolean
    Dim cfg As TMacroRecord
    If LoadConfigByID(rowID, cfg) = False Then Exit Function
    If cfg.SelectedForDuplicateDeletion = False Then Exit Function
    If Len(DuplicateIssueTypeForQuery(cfg.ID)) = 0 Then Exit Function
    SelectedRowStillEligibleForDuplicateDelete = True
End Function

Private Function ConfirmAllRowsSelectedDuplicateGroups(ByVal riskyGroupCount As Long) As Boolean
    Dim response As VbMsgBoxResult
    response = CoreShowMessageResult("ATTENTION: One or more duplicate groups have every row selected for deletion." & vbCrLf & vbCrLf & "This means both the entry and its duplicate row(s) will be deleted." & vbCrLf & "That may remove the entire configured macro/candidate group, not just the extra copy." & vbCrLf & vbCrLf & "Duplicate groups with every row selected: " & CStr(riskyGroupCount) & vbCrLf & vbCrLf & "Continue with duplicate deletion?", vbCritical Or vbYesNo Or vbDefaultButton2, "Attention: Entire Duplicate Groups Selected")
    ConfirmAllRowsSelectedDuplicateGroups = (response = vbYes)
End Function

Private Function CountAllRowsSelectedDuplicateGroups() As Long
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim cfg As TMacroRecord
    Dim groupKey As String
    Dim seenGroups As Collection
    ' INTENTIONAL FILTERED PARTITION: seenGroups is a consumer-required subset/projection derived from the source enumerated by CountAllRowsSelectedDuplicateGroups; its membership is not independently authoritative.
    Set seenGroups = New Collection
    Set rs = ADCDb().OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "]", dbOpenSnapshot)
    Do While Not rs.EOF
        PumpSetupResponsiveness
        MacroRecordFromRecordset rs, cfg
        groupKey = DuplicateGroupKeyForQuery(cfg.ID)
        If Len(groupKey) > 0 Then
            If SharedCollection_KeyExists(seenGroups, groupKey) = False Then
                MarkCollectionTextKey seenGroups, groupKey
                If DuplicateAllRowsSelectedForQuery(cfg.ID) Then CountAllRowsSelectedDuplicateGroups = CountAllRowsSelectedDuplicateGroups + 1
            End If
        End If
        rs.MoveNext
    Loop
ExitHere:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Exit Function
ErrHandler:
    CountAllRowsSelectedDuplicateGroups = 0
    Resume ExitHere
End Function

Private Function ConfirmDeleteOneDuplicateRow(ByVal rowID As Long) As VbMsgBoxResult
    Dim cfg As TMacroRecord
    Dim issueType As String
    Dim groupKey As String
    Dim groupCount As Long
    Dim selectedCount As Long
    If LoadConfigByID(rowID, cfg) = False Then
        ConfirmDeleteOneDuplicateRow = vbNo
        Exit Function
    End If
    issueType = DuplicateIssueTypeForQuery(cfg.ID)
    groupKey = DuplicateGroupKeyForQuery(cfg.ID)
    groupCount = DuplicateGroupCountForQuery(rowID)
    selectedCount = DuplicateSelectedCountForQuery(rowID)
    ConfirmDeleteOneDuplicateRow = CoreShowMessageResult("Delete this selected duplicate configuration row?" & vbCrLf & vbCrLf & "Duplicate issue: " & issueType & vbCrLf & "Duplicate group: " & groupKey & vbCrLf & "Rows in group: " & CStr(groupCount) & vbCrLf & "Rows selected for delete: " & CStr(selectedCount) & vbCrLf & vbCrLf & "ID: " & CStr(cfg.ID) & vbCrLf & "MacroKey: " & cfg.MacroKey & vbCrLf & "MacroTitle: " & cfg.MacroTitle & vbCrLf & "TargetProcedureName: " & cfg.TargetProcedureName & vbCrLf & vbCrLf & "Yes = Delete this row." & vbCrLf & "No = Skip this row." & vbCrLf & "Cancel = Stop duplicate deletion now.", vbExclamation Or vbYesNoCancel Or vbDefaultButton2, "Review Selected Duplicate Row")
End Function

' Purpose: Resolves macro export line name value from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   lineText (Input; String): Input line text for macro export line name value; is passed to `InStr`; is passed to `valueText`.
' Returns: String: Returns the resolved name produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function MacroExportLineNameValue(ByVal lineText As String) As String
    'Extracts and unescapes the value portion after the first equals sign in an exported Access macro line.
    Dim equalPos As Long
    Dim valueText As String
    equalPos = InStr(1, lineText, "=", vbBinaryCompare)
    If equalPos = 0 Then Exit Function
    valueText = Trim$(Mid$(lineText, equalPos + 1))
    If Len(valueText) >= 2 Then
        If Left$(valueText, 1) = """" And Right$(valueText, 1) = """" Then
            valueText = Mid$(valueText, 2, Len(valueText) - 2)
            valueText = Replace(valueText, """""", """")
        End If
    End If
    MacroExportLineNameValue = valueText
End Function

' Purpose: Parses exported Access macro text and returns the ordered RunCode target expressions.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   macroText (Input; String): Input macro text for run code targets from macro text; is passed to `Split`.
' Returns: Collection: Returns the collection of run code from macro text targets parsed from the supplied macro/source text, with each qualifying target represented once according to the routine's existing
'   parser rules.
Public Function RunCodeTargetsFromMacroText(ByVal macroText As String) As Collection
    'Parses exported Access macro text and returns the ordered RunCode target expressions.
    Dim targets As Collection
    Dim lines() As String
    Dim lineIndex As Long
    Dim lineText As String
    Dim expectingArgument As Boolean
    Dim valueText As String
    ' INTENTIONAL FILTERED PARTITION: targets is a consumer-required subset/projection derived from the source enumerated by RunCodeTargetsFromMacroText; its membership is not independently authoritative.
    Set targets = New Collection
    lines = Split(macroText, vbCrLf)
    For lineIndex = LBound(lines) To UBound(lines)
        PumpSetupResponsiveness
        lineText = Trim$(lines(lineIndex))
        If InStr(1, lineText, "Action", vbTextCompare) > 0 And InStr(1, lineText, """RunCode""", vbTextCompare) > 0 Then
            expectingArgument = True
        ElseIf expectingArgument And Left$(lineText, 8) = "Argument" Then
            valueText = MacroExportLineNameValue(lineText)
            If Len(Trim$(valueText)) > 0 Then targets.Add valueText
            expectingArgument = False
        ElseIf Left$(lineText, 6) = "Action" Then
            expectingArgument = False
        End If
    Next lineIndex
    Set RunCodeTargetsFromMacroText = targets
End Function

' Purpose: Resolves and returns configured target argument count from the supplied inputs and the authoritative ADC definitions/state used by this module.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   argumentsText (Input; String): Input arguments text for configured target argument count; is tested by If Len(Trim$(argumentsText)) = 0; is passed to `SplitTopLevelCommaText`.
' Returns: Long: Returns the number of configured target argument count items matched, processed, selected, deleted, or reviewed by this operation; zero represents no qualifying items.
Public Function ConfiguredTargetArgumentCount(ByVal argumentsText As String) As Long
    'Counts top-level configured target arguments without splitting nested expressions or quoted strings.
    Dim args As Collection
    If Len(Trim$(argumentsText)) = 0 Then Exit Function
    Set args = SplitTopLevelCommaText(argumentsText)
    ConfiguredTargetArgumentCount = args.count
End Function

Public Sub RequireConfigTableForUtilityWorkflow(ByVal sourceName As String)
    Dim fieldName As Variant
    EnsureSettingsReadyForRead
    If CodeDbTableExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")) = False Then
        Err.Raise vbObjectError + 5605, sourceName, "Access Development Console MC_Records table is missing. Run ADC Setup before this workflow." & vbCrLf & "MC_Records table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")
    End If
    If TableCanOpen(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")) = False Then
        Err.Raise vbObjectError + 5606, sourceName, "Access Development Console MC_Records table exists but cannot be opened. Run ADC Setup before this workflow." & vbCrLf & "MC_Records table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")
    End If
    For Each fieldName In Split(ConfigFieldOrderText(), "|")
        PumpSetupResponsiveness
        If TableFieldExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config"), CStr(fieldName)) = False Then
            Err.Raise vbObjectError + 5607, sourceName, "Access Development Console MC_Records table is missing a required current-design field. Run ADC Setup before this workflow." & vbCrLf & "MC_Records table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & vbCrLf & "Missing field: " & CStr(fieldName)
        End If
    Next fieldName
End Sub

' Purpose: Resolves configured target for config ID.
' Public role: Form-code/event-expression public entry point. Access form code or an event property can invoke this routine by name; its Public visibility is part of that UI contract.
' Parameters:
'   ConfigID (Input; Long): Primary-key ID of the configuration row targeted by the generated UI/action.
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub ResolveConfiguredTarget(ByVal configID As Long, Optional ByVal showMessage As Boolean = True)
    'Form-facing single-record target-resolution entry point. Resolution remains owned by ResolveRunnerMacroTargetsCore.
    On Error GoTo ErrHandler
    Dim progressStarted As Boolean
    If configID <= 0 Then Err.Raise vbObjectError + 9191, "ResolveConfiguredTarget", "ConfigID must be greater than zero."
    progressStarted = CoreBeginOperationProgressIfNeeded("Resolve " & DEFAULT_UTILITY_SHORT_NAME & " configured target", 1)
    CoreUpdateSetupProgress "Resolve configured target ID " & CStr(configID)
    If ResolveRunnerMacroTargetsCore(showMessage, configID) = False Then Err.Raise vbObjectError + 5610, "ResolveConfiguredTarget", "Runner Macro target resolution did not complete successfully." & vbCrLf & "ConfigID: " & CStr(configID)
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
    If showMessage Then
        CoreShowErrorMessage "Could not resolve the requested Access Development Console configured target." & vbCrLf & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, vbCritical, "Resolve Configured Target", "ResolveConfiguredTarget", "Review"
        Exit Sub
    End If
    Err.Raise savedErrNumber, "ResolveConfiguredTarget", "Could not resolve the requested Access Development Console configured target." & vbCrLf & "ConfigID: " & CStr(configID) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

' Purpose: Resolves configured Runner Macro targets for active configuration rows, records validation results in MC_Records, and coordinates progress plus optional user-facing error reporting.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub ResolveRunnerMacroTargets(Optional ByVal showMessage As Boolean = True)
    On Error GoTo ErrHandler
    Dim progressStarted As Boolean
    progressStarted = CoreBeginOperationProgressIfNeeded("Resolve " & DEFAULT_UTILITY_SHORT_NAME & " targets", 1)
    CoreUpdateSetupProgress "Resolve configured macro targets"
    If ResolveRunnerMacroTargetsCore(showMessage) = False Then
        Err.Raise vbObjectError + 5610, "ResolveRunnerMacroTargets", "Runner Macro target resolution did not complete successfully." & vbCrLf & "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB & vbCrLf & "MC_Records table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & vbCrLf & "Review ConfigOwnerDatabaseRole, LastValidationResult, and LastValidationNotes on affected config rows."
    End If
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
    If showMessage Then
        CoreShowErrorMessage "Could not resolve Runner Macro targets." & vbCrLf & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, vbCritical, "Resolve Runner Macro Targets", "ResolveRunnerMacroTargets", "Review"
        Exit Sub
    End If
    Err.Raise savedErrNumber, "ResolveRunnerMacroTargets", "Could not resolve Runner Macro targets." & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Function ResolveRunnerMacroTargetsCore(Optional ByVal showMessage As Boolean = True, Optional ByVal configID As Long = 0) As Boolean
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim cfg As TMacroRecord
    Dim reviewedCount As Long
    Dim resolvedCount As Long
    Dim invalidCount As Long
    Dim skippedCount As Long
    Dim statusText As String
    Dim validationResultText As String
    Dim resolvedTarget As String
    Dim detailText As String
    Dim roleText As String
    Dim procedureKind As String
    Dim procedureLocation As String
    Dim argumentSignature As String
    Dim RequiredArgumentCount As Long
    Dim OptionalArgumentCount As Long
    Dim hasParamArray As Boolean
    Dim ExpectedReturnType As String
    Dim vbProject As Object
    Dim currentID As Long
    RequireConfigTableForUtilityWorkflow "ResolveRunnerMacroTargetsCore"
    Set db = ADCDb()
    If configID > 0 Then
        Set rs = db.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "] WHERE " & SQL_FIELD_ID & "=" & CStr(configID), dbOpenDynaset)
        If rs.EOF Then Err.Raise vbObjectError + 9194, "ResolveRunnerMacroTargetsCore", "No configuration row exists for the requested ConfigID." & vbCrLf & "ConfigID: " & CStr(configID)
    Else
        Set rs = db.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "] ORDER BY " & SQL_FIELD_ID, dbOpenDynaset)
    End If
    Do While Not rs.EOF
        PumpSetupResponsiveness
        MacroRecordFromRecordset rs, cfg
        currentID = SharedValue_NzLong(rs.fields(FIELD_ID).value)
        If configID <= 0 And (cfg.Active = False Or ConfigurationStatusBlocksGeneration(cfg.ConfigurationStatus)) Then
            skippedCount = skippedCount + 1
        Else
            reviewedCount = reviewedCount + 1
            PulseProgressEvery reviewedCount, 0, "Resolving configured target row ID " & CStr(currentID)
            roleText = ResolvedConfigOwnerDatabaseRoleForConfig(cfg)
            detailText = vbNullString
            resolvedTarget = vbNullString
            statusText = vbNullString
            procedureKind = vbNullString
            procedureLocation = vbNullString
            argumentSignature = vbNullString
            RequiredArgumentCount = 0
            OptionalArgumentCount = 0
            hasParamArray = False
            ExpectedReturnType = vbNullString
            If Len(roleText) = 0 Then
                statusText = VALIDATION_RESULT_INVALID
                detailText = "ConfigOwnerDatabaseRole is blank or invalid."
            Else
                cfg.ConfigOwnerDatabaseRole = roleText
                Select Case LCase$(Trim$(cfg.MacroAction))
                    Case LCase$(ACTION_RUNCODE)
                        statusText = ValidateConfiguredRunCodeTarget(cfg, resolvedTarget, detailText)
                        If PublicCallableProcedureInfoInStandardModule(cfg.TargetProcedureName, roleText, procedureKind, procedureLocation, argumentSignature, RequiredArgumentCount, OptionalArgumentCount, hasParamArray, ExpectedReturnType) Then
                            cfg.TargetProcedureKind = procedureKind
                            cfg.TargetProcedureModuleName = procedureLocation
                            cfg.RequiredArgumentCount = RequiredArgumentCount
                            cfg.OptionalArgumentCount = OptionalArgumentCount
                            cfg.ExpectedReturnType = ExpectedReturnType
                        End If
                    Case LCase$(ACTION_OPENFORM), LCase$(ACTION_OPENTABLE), LCase$(ACTION_OPENQUERY), LCase$(ACTION_OPENREPORT), LCase$(ACTION_RUNMACRO)
                        statusText = ValidateConfiguredAccessObjectTarget(cfg, resolvedTarget, detailText)
                    Case Else
                        statusText = VALIDATION_RESULT_INVALID
                        detailText = "Unsupported MacroAction: " & cfg.MacroAction
                End Select
            End If
            validationResultText = ValidationResultFromTargetStatus(statusText)
            rs.Edit
            SetRecordsetFieldIfExists rs, FIELD_CONFIG_OWNER_DATABASE_ROLE, roleText
            If Len(roleText) > 0 Then
                SetRecordsetFieldIfExists rs, FIELD_GENERATION_TARGET_DB_PATH, DatabasePathForRole(roleText)
                SetRecordsetFieldIfExists rs, FIELD_GENERATION_TARGET_DB_NAME, DatabaseBaseNameFromPath(DatabaseForRole(roleText).Name)
                Set vbProject = VBProjectForConfigOwnerDatabaseRole(roleText)
                If Not vbProject Is Nothing Then SetRecordsetFieldIfExists rs, FIELD_TARGET_VB_PROJECT_NAME, SafeObjectText(vbProject, "Name")
            End If
            If Len(cfg.TargetProcedureModuleName) > 0 Then SetRecordsetFieldIfExists rs, FIELD_TARGET_PROCEDURE_MODULE_NAME, cfg.TargetProcedureModuleName
            If Len(cfg.TargetProcedureKind) > 0 Then SetRecordsetFieldIfExists rs, FIELD_TARGET_PROCEDURE_KIND, cfg.TargetProcedureKind
            If cfg.RequiredArgumentCount > 0 Then SetRecordsetFieldIfExists rs, FIELD_REQUIRED_ARGUMENT_COUNT, cfg.RequiredArgumentCount
            If cfg.OptionalArgumentCount > 0 Then SetRecordsetFieldIfExists rs, FIELD_OPTIONAL_ARGUMENT_COUNT, cfg.OptionalArgumentCount
            If Len(cfg.ExpectedReturnType) > 0 Then SetRecordsetFieldIfExists rs, FIELD_EXPECTED_RETURN_TYPE, cfg.ExpectedReturnType
            SetRecordsetFieldIfExists rs, FIELD_LAST_VALIDATED_ON, Now
            SetRecordsetFieldIfExists rs, FIELD_LAST_VALIDATION_RESULT, validationResultText
            SetRecordsetFieldIfExists rs, FIELD_LAST_VALIDATION_NOTES, detailText
            If StrComp(validationResultText, VALIDATION_RESULT_VALID, vbTextCompare) = 0 Then
                If StrComp(cfg.ConfigurationStatus, CONFIGURATION_STATUS_NEEDS_REVIEW, vbTextCompare) = 0 Or StrComp(cfg.ConfigurationStatus, CONFIGURATION_STATUS_INVALID, vbTextCompare) = 0 Then
                    SetRecordsetFieldIfExists rs, FIELD_CONFIGURATION_STATUS, CONFIGURATION_STATUS_ACTIVE
                End If
                resolvedCount = resolvedCount + 1
            Else
                SetRecordsetFieldIfExists rs, FIELD_CONFIGURATION_STATUS, CONFIGURATION_STATUS_NEEDS_REVIEW
                invalidCount = invalidCount + 1
            End If
            rs.Update
        End If
        rs.MoveNext
    Loop
    InvalidateDuplicateReviewCache
    If showMessage Then
        OpenReviewQuery ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Config_TargetResolutionAudit"), "ResolveRunnerMacroTargetsCore"
        CoreShowMessage "Access Development Console target resolution complete." & vbCrLf & vbCrLf & "Rows reviewed: " & CStr(reviewedCount) & vbCrLf & "Rows resolved/valid: " & CStr(resolvedCount) & vbCrLf & "Rows needing review: " & CStr(invalidCount) & vbCrLf & "Rows skipped: " & CStr(skippedCount) & vbCrLf & vbCrLf & "No macros were generated or replaced.", vbInformation, "Resolve Runner Macro Targets"
    End If
    ResolveRunnerMacroTargetsCore = True
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Exit Function
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    ResolveRunnerMacroTargetsCore = False
    Err.Raise originalErrNumber, "ResolveRunnerMacroTargetsCore", "Could not complete target resolution." & vbCrLf & "Err.Description: " & originalErrDescription
End Function

Private Sub ValidateConfigClassification(ByRef cfg As TMacroRecord)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim storedCategoryKey As String
    If Len(Trim$(cfg.MacroSubcategoryKey)) > 0 And Len(Trim$(cfg.MacroCategoryKey)) = 0 Then Err.Raise vbObjectError + 9132, "ValidateConfigClassification", "MacroSubcategoryKey cannot be populated when MacroCategoryKey is blank." & vbCrLf & "MacroKey: " & cfg.MacroKey
    If Len(Trim$(cfg.MacroCategoryKey)) = 0 Then Exit Sub
    Set db = ADCDb()
    sqlText = "SELECT TOP 1 " & SQL_FIELD_MACRO_CATEGORY_KEY & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_MacroCategories") & "] WHERE " & SQL_FIELD_MACRO_CATEGORY_KEY & "=" & SharedDao_SqlQuote(cfg.MacroCategoryKey) & " AND Nz(" & SQL_FIELD_ACTIVE & ",True)<>False;"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    If rs.EOF Then Err.Raise vbObjectError + 9133, "ValidateConfigClassification", "MacroCategoryKey does not match an active MC_Cats row." & vbCrLf & "MacroKey: " & cfg.MacroKey & vbCrLf & "MacroCategoryKey: " & cfg.MacroCategoryKey
    rs.Close
    Set rs = Nothing
    If Len(Trim$(cfg.MacroSubcategoryKey)) = 0 Then Exit Sub
    sqlText = "SELECT TOP 1 " & SQL_FIELD_MACRO_CATEGORY_KEY & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_MacroSubcategories") & "] WHERE " & SQL_FIELD_MACRO_SUBCATEGORY_KEY & "=" & SharedDao_SqlQuote(cfg.MacroSubcategoryKey) & " AND Nz(" & SQL_FIELD_ACTIVE & ",True)<>False;"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    If rs.EOF Then Err.Raise vbObjectError + 9134, "ValidateConfigClassification", "MacroSubcategoryKey does not match an active MC_Cats_SubCats row." & vbCrLf & "MacroKey: " & cfg.MacroKey & vbCrLf & "MacroSubcategoryKey: " & cfg.MacroSubcategoryKey
    storedCategoryKey = Trim$(CStr(Nz(rs.fields(FIELD_MACRO_CATEGORY_KEY).value, vbNullString)))
    If StrComp(storedCategoryKey, Trim$(cfg.MacroCategoryKey), vbTextCompare) <> 0 Then Err.Raise vbObjectError + 9135, "ValidateConfigClassification", "The selected macro subcategory belongs to a different category." & vbCrLf & "MacroKey: " & cfg.MacroKey & vbCrLf & "Config category: " & cfg.MacroCategoryKey & vbCrLf & "Subcategory parent: " & storedCategoryKey
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
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
    Err.Raise savedErrNumber, "ValidateConfigClassification", "Generated-macro category/subcategory validation failed." & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Sub

Private Function ValidateConfiguredAccessObjectTarget(ByRef cfg As TMacroRecord, ByRef resolvedTarget As String, ByRef detailText As String) As String
    Dim objectType As AcObjectType
    Dim actionText As String
    actionText = Trim$(cfg.MacroAction)
    resolvedTarget = ConfiguredObjectActionTargetName(cfg)
    If Len(resolvedTarget) = 0 Then
        detailText = "Configured " & actionText & " row has no target object name. Populate TargetObjectName."
        ValidateConfiguredAccessObjectTarget = "UNRESOLVED"
        Exit Function
    End If
    If StrComp(NormalizedConfigOwnerDatabaseRole(cfg.ConfigOwnerDatabaseRole), CONFIG_OWNER_DATABASE_ROLE_CODE_DB, vbTextCompare) = 0 Then
        detailText = actionText & " targets must exist in the HostDb Access project. ConfigOwnerDatabaseRole=CodeDb is not directly supported for Access object-action rows."
        ValidateConfiguredAccessObjectTarget = "UNRESOLVED"
        Exit Function
    End If
    Select Case UCase$(actionText)
        Case UCase$(ACTION_OPENFORM)
            objectType = acForm
        Case UCase$(ACTION_OPENTABLE)
            objectType = acTable
        Case UCase$(ACTION_OPENQUERY)
            objectType = acQuery
        Case UCase$(ACTION_OPENREPORT)
            objectType = acReport
        Case UCase$(ACTION_RUNMACRO)
            objectType = acMacro
        Case Else
            detailText = "Unsupported Access object MacroAction: " & actionText
            ValidateConfiguredAccessObjectTarget = "UNRESOLVED"
            Exit Function
    End Select
    If AccessObjectExistsInCurrentDatabase(objectType, resolvedTarget) Then
        detailText = "Configured " & actionText & " target exists in the HostDb Access project: " & resolvedTarget
        ValidateConfiguredAccessObjectTarget = "VALID"
    Else
        detailText = "Missing configured " & actionText & " target: " & resolvedTarget
        ValidateConfiguredAccessObjectTarget = "MISSING"
    End If
End Function


Private Sub EnsureNoActiveAutoExecConfigInconsistencies()
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim lines As Collection
    Dim MacroKey As String
    Dim MacroTitle As String
    Dim macroTitleText As String
    If CodeDbTableExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")) = False Then Exit Sub
    ' INTENTIONAL ORDERED VIEW: lines preserves the deterministic sequence required by EnsureNoActiveAutoExecConfigInconsistencies; keyed-only or unordered storage would not preserve the consumer semantics.
    Set lines = New Collection
    Set rs = ADCDb().OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "] " & "WHERE " & SQL_FIELD_ACTIVE & "<>0 AND " & "(UCase(Nz(" & SQL_FIELD_MACRO_KEY & ",''))='AUTOEXEC' OR UCase(Nz(" & SQL_FIELD_MACRO_TITLE & ",''))='AUTOEXEC') " & "ORDER BY Nz(" & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER & ",0), " & SQL_FIELD_MACRO_TITLE & ", " & SQL_FIELD_MACRO_KEY, dbOpenSnapshot)
    Do While Not rs.EOF
        PumpSetupResponsiveness
        MacroKey = SharedText_Nz(rs.fields(FIELD_MACRO_KEY).value)
        MacroTitle = SharedText_Nz(rs.fields(FIELD_MACRO_TITLE).value)
        macroTitleText = MacroTitle
        If StrComp(MacroKey, CONFIRM_AUTOEXEC, vbTextCompare) = 0 And StrComp(macroTitleText, CONFIRM_AUTOEXEC, vbTextCompare) <> 0 Then
            lines.Add "Active row has MacroKey=AutoExec but MacroTitle is not AutoExec." & vbCrLf & "  MacroKey: " & MacroKey & vbCrLf & "  MacroTitle: " & MacroTitle
        End If
        rs.MoveNext
    Loop
    If lines.count > 0 Then
        Err.Raise vbObjectError + 5221, "EnsureNoActiveAutoExecConfigInconsistencies", "AutoExec config preflight failed." & vbCrLf & vbCrLf & SharedCollection_Join(lines, vbCrLf & vbCrLf) & vbCrLf & vbCrLf & "AutoExec rows must generate an Access macro object named AutoExec. Correct the active config rows before generation."
    End If
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Exit Sub
ErrHandler:
    Dim handlerErrNumber As Long
    Dim handlerErrDescription As String
    handlerErrNumber = Err.Number
    handlerErrDescription = Err.Description
    If handlerErrNumber < 0 Then
        On Error Resume Next
        If Not rs Is Nothing Then rs.Close
        On Error GoTo 0
        Err.Raise handlerErrNumber, "EnsureNoActiveAutoExecConfigInconsistencies", FormatErrorWithContext("Could not run AutoExec config preflight.", "MC_Records table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & vbCrLf & "Preflight: validate active AutoExec config consistency before generation.")
    Else
        Err.Raise handlerErrNumber, "EnsureNoActiveAutoExecConfigInconsistencies", "Could not run AutoExec config preflight." & vbCrLf & "MC_Records table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & vbCrLf & "Err.Description: " & handlerErrDescription
    End If
End Sub

Private Sub EnsureConfiguredTargetCanGenerate(ByRef cfg As TMacroRecord)
    Dim statusText As String
    Dim resolvedTarget As String
    Dim detailText As String
    Select Case UCase$(Trim$(cfg.MacroAction))
        Case UCase$(ACTION_RUNCODE)
            statusText = ValidateConfiguredRunCodeTarget(cfg, resolvedTarget, detailText)
        Case UCase$(ACTION_OPENFORM), UCase$(ACTION_OPENTABLE), UCase$(ACTION_OPENQUERY), UCase$(ACTION_OPENREPORT), UCase$(ACTION_RUNMACRO)
            statusText = ValidateConfiguredAccessObjectTarget(cfg, resolvedTarget, detailText)
        Case Else
            statusText = "UNRESOLVED"
            resolvedTarget = cfg.MacroAction
            detailText = "Unsupported MacroAction."
    End Select
    If StrComp(statusText, "VALID", vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 5211, "EnsureConfiguredTargetCanGenerate", "Configured macro target is not valid. Macro generation stopped before creating/replacing the macro object." & vbCrLf & vbCrLf & "MacroKey: " & cfg.MacroKey & vbCrLf & "MacroTitle: " & cfg.MacroTitle & vbCrLf & "MacroAction: " & cfg.MacroAction & vbCrLf & "Status: " & statusText & vbCrLf & "Resolved target: " & IIf(Len(Trim$(resolvedTarget)) = 0, "(none)", resolvedTarget) & vbCrLf & "Detail: " & IIf(Len(Trim$(detailText)) = 0, "(none)", detailText)
    End If
End Sub

' Purpose: Resolves and returns configured object action target name from the supplied inputs and the authoritative ADC definitions/state used by this module.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   cfg (Input; TMacroRecord): ADC macro record containing the persisted accepted/current configuration values used by the operation; the routine reads it and may populate or update its fields when passed ByRef.
' Returns: String: Returns the physical/logical name resolved for configured object action target name under the module's naming/definition rules; vbNullString indicates that no applicable name can be resolved
'   where supported.
Public Function ConfiguredObjectActionTargetName(ByRef cfg As TMacroRecord) As String
    ConfiguredObjectActionTargetName = cfg.TargetObjectName
End Function

Private Sub ValidateTargetProcedureMetadata(ByRef cfg As TMacroRecord)
    Select Case LCase$(Trim$(cfg.TargetProcedureKind))
        Case "", LCase$(TARGET_PROCEDURE_KIND_AUTO), LCase$(TARGET_PROCEDURE_KIND_FUNCTION), LCase$(TARGET_PROCEDURE_KIND_SUB)
            'Valid.
        Case Else
            Err.Raise vbObjectError + 5206, "ValidateTargetProcedureMetadata", "TargetProcedureKind must be Auto, Function, or Sub." & vbCrLf & "MacroKey: " & cfg.MacroKey & vbCrLf & "TargetProcedureKind: " & cfg.TargetProcedureKind
    End Select
    Select Case LCase$(Trim$(cfg.ResultHandling))
        Case "", LCase$(RESULT_HANDLING_AUTO), LCase$(RESULT_HANDLING_IGNORE), LCase$(RESULT_HANDLING_BOOLEAN), LCase$(RESULT_HANDLING_REQUIRE_TRUE)
            'Valid.
        Case Else
            Err.Raise vbObjectError + 5207, "ValidateTargetProcedureMetadata", "ResultHandling must be Auto, Ignore, Boolean, or RequireTrue." & vbCrLf & "MacroKey: " & cfg.MacroKey & vbCrLf & "ResultHandling: " & cfg.ResultHandling
    End Select
    If cfg.RequiredArgumentCount > 0 And Len(Trim$(cfg.TargetArguments)) = 0 Then
        Err.Raise vbObjectError + 5210, "ValidateTargetProcedureMetadata", "This macro target was discovered with required arguments, but TargetArguments is blank." & vbCrLf & "MacroKey: " & cfg.MacroKey & vbCrLf & "TargetProcedureName: " & cfg.TargetProcedureName & vbCrLf & "RequiredArgumentCount: " & CStr(cfg.RequiredArgumentCount) & vbCrLf & "Populate TargetArguments before generating this macro."
    End If
End Sub


' Purpose: Builds a diagnostic message from caller-supplied summary/context plus the current VBA Err number, source, and description, omitting optional context when it is blank.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   summaryText (Input; String): Summary text text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   contextText (Input Optional; String): Context text text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine. Default: vbNullString.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function FormatErrorWithContext(ByVal summaryText As String, Optional ByVal contextText As String = vbNullString) As String
    'Intended use: Returns the Format Access Development Console Error With Context result for direct, query, form, or required cross-module callers.
    Dim resultText As String
    resultText = summaryText
    If Len(Trim$(contextText)) > 0 Then resultText = resultText & vbCrLf & contextText
    resultText = resultText & vbCrLf & "Err.Number: " & CStr(Err.Number) & vbCrLf & "Err.Source: " & Err.Source & vbCrLf & "Err.Description: " & Err.Description
    FormatErrorWithContext = resultText
End Function
