Attribute VB_Name = "x__UTIL_ADC_Core"
' ================================================================================================================================================================================
' ChatGPT Rewrite Tracking Header
' ================================================================================================================================================================================
' ChatGPT Previous Version ID: CGPT-2XQE-SMQR-7BPA-8SYL-QLFB
' ChatGPT Baseline Version ID: CGPT-2XQE-SMQR-7BPA-8SYL-QLFB
' ChatGPT Version ID: CGPT-39HE-U2R9-SH4J-XWVJ-RRUU
' ChatGPT Version Timestamp: Tuesday 9/1/26 3:40 PM
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
' STANDARD MODULE: ADC Utility Core
' VB_Name: x__UTIL_ADC_Core
' Entry Points:
'     RunFullSetup
' SHARED CODE:
'     SharedCollection_Join(ByVal values As Collection, Optional ByVal delimiter As String = ", ") -> String
'     SharedCollection_KeyExists(ByVal values As Collection, ByVal key As String) -> Boolean
'     SharedDao_DatabaseTableExistsSafe(ByVal db As DAO.Database, ByVal tableName As String) -> Boolean
'     SharedDao_DatabaseTableFieldExistsSafe(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String) -> Boolean
'     SharedDao_SafeDatabaseName(ByVal db As DAO.Database) -> String
'     SharedDao_SqlQuote(ByVal value As Variant) -> String
'     SharedDao_TableExists(ByVal db As DAO.Database, ByVal tableName As String) -> Boolean
'     SharedMessage_OneLine(ByVal valueText As String, Optional ByVal maxLength As Long = 700, Optional ByVal ellipsisText As String = "...") -> String
'     SharedText_AppendNonBlankLine(ByVal existingText As String, ByVal lineText As String) -> String
' CLASSES USED: None
' DYNAMIC STRING REFERENCES:
'     Dynamic-reference scan stopped after 2000 logical line(s) to prevent a long-running scan; review this module manually if needed.
' LastScanTimestamp: Saturday 8/15/26 8:40 AM
' Purpose:
'     Owns the active Access Development Console Setup Preflight/Core infrastructure used to authorize and run the resequenced Setup through Phase J, including setup lifecycle, phase-group rollback checkpoints,
'     rollback/cancellation, physical System Macro and Help Document Macro verification/materialization, Navigation Pane rollback/proof support, managed-description finalization, final J1 completion-state
'     persistence, and the public system-command runtime bridge.
' Author: Christina Shonts
' Support:
'     Access Development Console utility for CodeDb maintenance and HostDb macro-generation workflows; not a broad HostDb environment copier.
' Code Category:
'     Access Development Console V6 core utility; Setup lifecycle/orchestration, recovery, semantic progress, System Command dispatch, and shared command result state.
' Requirements:
'     Microsoft Access with DAO available.
'     Trusted access to the VBA project object model is required for candidate scans and procedure validation paths that inspect code.
'     CodeDb must own the full Access Development Console utility environment and live managed support tables.
'     HostDb work must be routed through explicit database references and must remain limited to HostDb MacroCatalog_Records plus generated macro objects.
' Reference Dependencies:
'     Access object model and DAO are required.
'     VBE/VBProject objects are used through late-bound inspection paths where candidate scanning or procedure validation requires code metadata.
'     Shared helper modules provide SQL quoting, DAO inspection/repair, Navigation Pane grouping, and progress/status reporting.
' DB Object Dependencies:
'     CodeDb-owned tables include MacroCatalog_Records, Settings, MacroCatalog_Cats, MacroCatalog_Cats_SubCats, MacroCatalog_Sys, UI_Ribbon_Generated_BtnTargets, Help_Docs, UI_Groups, UI_Frm_Lookup_Types,
'     UI_Frm_Registry, MAC_Cats, MAC, MAC_ActionPayloads and MAC_Deps.
'     HostDb-owned data for this utility is limited to MacroCatalog_Records in the HostDb plus generated Access macro objects created from that HostDb configuration.
'     Review/query/form support objects are CodeDb-owned unless an explicit generated HostDb macro workflow creates generated macro objects in HostDb.
' Usage:
'     Use RunFullSetup for CodeDb setup and repair of the Access Development Console utility environment.
'     Use table-driven scan command routing when discovering candidates so ScanTargetDb and ResultConfigDb intent is explicit.
'     Accept candidate rows before generation; generation reads accepted/configured rows from MacroCatalog_Records in the resolved ConfigDb.
'     Use table-driven generation command routing for CodeDb or HostDb generation instead of relying on role-ambiguous database behavior.
'     Use OpenADC* procedures for review surfaces; Open procedures should not perform broad setup or generation work.
' Version:
'     Finalizes a Core-owned Setup trace when Module Guard state resolution is cancelled during unnumbered Setup Preflight, preventing cancellation from leaving the active trace identity and retained
'     physical trace handle open; numbered Setup-phase ownership is unchanged.
'     Adds Setup Phase F dependency rollback/domain support and shifts existing description, Ribbon, Navigation Pane, and final Setup ownership to G, H, I, and J.
'     Reads the Core build identity from its commented rewrite header through Module Guard instead of embedding an executable module Version ID; Setup behavior and phase ownership are unchanged.
'     Adds durable Delete All Objects terminal-path checkpoints and physical trace-file flushes around deferred Macro 65 handoff, deletion, completion, and RunCode return.
'     Diagnostic-only changes do not alter deletion behavior or Setup-phase ownership.
'     Moves Phase C macro parsing and Core recovery serialization helpers into Core beside their owning phase and durable-recovery workflows; behavior and phase ownership are unchanged.
'     Corrects stale C3/I1 documentation so it reflects fresh read-only current-state verification; executable behavior and Setup-phase ownership are unchanged.
'     Routes Managed Object Catalog references to x__UTIL_ADC_MOC and physical lifecycle references to x__UTIL_ADC_MOC_LIFECYCLE after the responsibility split; runtime and Setup-phase behavior are unchanged.
'     Suppresses the nested RepairTables standalone Description pass inside Build/Repair All Tables and Queries because the Core coordinator already performs the final complete managed table/query reconciliation.
'     Ensures standalone recovery storage and managed query rollback reapply canonical managed descriptions; Build/Repair All Tables and Queries performs scoped standalone description reconciliation
'     while numbered Setup still defers normal writes to F1.
'     Removes the constant-return WorkflowName layer, adopts the typed Module Guard registration contract, and narrows module-internal visibility while retaining required callbacks.
'     Phase C now keeps C1 as the sole pre-mutation whole-set physical macro evidence construction point, retains one operation-scoped physical evidence snapshot through C2/C3, validates only actual mutation/delete
'     targets with targeted payload rereads, and derives the final C3 exact-state proof without a second whole-set SaveAsText payload pass.
'     Build/Repair Help Document Macros reuses the authoritative D4 complete-catalog finalizer after Help catalog/physical reconciliation before H Ribbon, then applies I Navigation Pane placement, so Ribbon consumes
'     a current persisted catalog proof.
'     Full Setup lifecycle ownership is A Tables, Queries, and DB Inventory; B Managed Object Catalog; C System Macro and Help Document Macro; D Forms; E Reports; F Managed Object Dependencies;
'     G Object and Field Descriptions; H Ribbon; I Navigation Pane Configuration; and J Final Setup Verification, with G1-only Setup description-write gates.
'     Scan All DB Inventory now reuses operation-scoped scan finalization controls so shared inventories and dependent links/correlations are completed only once after their authoritative inputs are current.
'     Managed-header dependency metadata updated because x__UTIL_ADC_DEFAULTS is now a standard module; executable behavior unchanged.
'     Consolidated runtime HostDb resolution and the shared System Command result-state lifecycle under Core.
'     Consolidated reusable durable-recovery workspace/database opening and storage-readiness lifecycle under Core so UI no longer owns duplicate low-level recovery database plumbing.
'     Removed Core's direct Access status-bar meter reset so ordinary physical progress/status presentation remains UI-owned while Core retains semantic top-level progress-state reset behavior.
'     Consolidated the temporary command split into the canonical System Command dispatcher and routes operational domains directly to their permanent authorities.
'     Consolidated nested parent progress text and rollback-backup deletion under Core ownership.
'     Exposed the existing Core process-summary capture/read operations so UI workflows consume the same process state instead of maintaining a second tracker.
'     Centralized result-returning popup display and suppression policy beside the existing Core popup authority.
'     Removed Core's duplicate Setup Control active-state mirror; Core now queries the UI-owned authority.
'     Core owns runtime database-context resolution, including role-to-database/path and Config-owner VBProject resolution; Common owns stateless database-role/path normalization. Catalog does not own global
'     database context.
' Notes:
'     MC_Records is the accepted/current macro-configuration authority; MC_Candidates separately stores candidate/discovery evidence and acceptance state, linked by AcceptedConfigurationID after acceptance.
'     CodeDb and HostDb may each have a table named MC_Records because they are different database files; code must choose the database through explicit DAO.Database references.
'     Settings is CodeDb-owned. HostDb does not own Settings in the current V6 design.
'     Default rows are maintained directly in live CodeDb-owned managed tables; there is no second active table family for the current design.
'     Generated configured macro objects call RunConfiguredMacroTarget or RunConfirmedConfiguredMacroTarget with MacroKey; generated system macro shortcuts call RunSystemCommand with SystemMacroKey.
'     Fixed review-open, CodeDb scan, CodeDb generation, stale-preview, duplicate-delete, and Custom Macro Launcher system commands are dispatched directly instead of passing through one-procedure wrapper entry
'     points.
'     Module Guard mismatch Cancel exits Setup immediately; Yes validates and transactionally registers the exact installed managed-module set the user explicitly approves instead of rerunning the strict
'     embedded-build registrar.
'     A1-J1 setup-part physical macros use Part-ID names (_A1.._J1) and dedicated SetupParts Navigation Pane metadata; Phase I is the sole Setup authority for physical Navigation Pane group creation and
'     managed-object placement.
'     Candidate discovery records managed module-header metadata when recognized headers exist and flags editable Entry Points entries as metadata only. System macro rows are table-driven through MacroCatalog_Sys
'     and represent CodeDb utility shortcuts, not HostDb generated macro output.
'     Help-document shortcut macros are generated from Help_Docs, with descriptions supplied as Access object Description text.
'     Managed Object Catalog support routes utility-owned object import/generation/grouping through MAC and related catalog tables.
'     Long setup, generation, and review summaries are split or routed through status/progress helpers so Access message boxes are not oversized.
'     Custom Macro Launcher automation remains a development-build utility workflow and must stay isolated from ordinary generated macro output.
' Core Rules Enforced Here:
'     Reporting ownership rule: user-invoked, generated-macro, and form-button entry points own one diagnostic popup; shared helpers log/raise upward and must not show duplicate popups.
'     Setup is not generation. Setup repairs required tables/support objects; generation creates or replaces generated macro objects from config rows.
'     CodeDb setup may repair the full Access Development Console utility environment; HostDb setup must remain narrow.
'     Full through-phase Setup commits A1-C3, D1-G1, H1-H3, and I1-I3 as separate durable rollback groups; a later-group failure must not delete or restore an earlier verified group.
'     Generated macro replacement backs up the target, deletes/replaces it, and attempts restore if final replacement fails.
'     OpenForm generated macro objects do not ask for confirmation; RunCode generated macro objects ask unless the config row disables confirmation.
'     Generated macro titles are validated before generation; overlength names fail and duplicate authoritative names are rejected without truncation.
'     Managed-object Navigation Pane grouping adds shortcuts without deleting existing shortcut rows.
'     Backup object Navigation Pane grouping is exclusive: ADC backup and recovery objects are assigned only to Access Development Console - Backup.
'     Custom Macro Launcher cleanup deletes only confirmed stale or duplicate tagged buttons; untagged or malformed controls are not deleted.
'     Candidate scan, target resolution, validation, and generation should continue collecting row-level findings instead of stopping at the first non-critical row issue.
' Design Map:
'     RunFullSetup sequences CodeDb-owned table repair, settings/default row repair, system macro catalog repair, managed-object repair, support query/form refresh, generated utility-form rebuild, ribbon refresh,
'     and grouping while avoiding repeated table/review foundation runs when the UI builder is called from setup/repair.
'     RepairADCCodeDb runs a faster after-first-run update path that repairs tables/catalogs/review support, refreshes generated system macro shortcuts/descriptions, and verifies or creates missing generated utility
'     forms without forcing broad managed-object/form replacement.
'     MacroCatalog_Records in the resolved ConfigDb contains accepted configuration rows and candidate/discovery rows.
'     Scan command routing selects the intended ScanTargetDb and ResultConfigDb pattern before writing candidate rows.
'     Generation command routing selects the intended ConfigDb and GenerationTargetDb pattern before creating macro objects.
'     Review helpers update or open health review, target resolution, stale-candidate, generation-status, duplicate-row, manual-review, settings, candidate, and config review surfaces.
'     Form-button helpers operate on the current form record and delegate to shared row-level workflow routines; table grouping classifies table names inline instead of depending on public IsADC*TableName wrappers.
'     Managed Object Catalog routines validate catalog rows, import or generate managed utility objects, group them, and maintain installed-object/dependency records.
'     RunSystemCommand is the keyed Boolean RunCode dispatcher for Access macro buttons and generated system macro shortcuts.
' Configuration:
'     MacroCatalog_Records controls generated macro title, target type, target procedure/form/object, target arguments, confirmation behavior, grouping, candidate state, and generation status.
'     Settings controls CodeDb utility behavior such as detailed errors, setup popups, Navigation Pane groups, launcher form names, form replacement, and warning toggles.
'     MacroCatalog_Sys controls CodeDb utility shortcut numbering, calculated titles, title overrides, descriptions, and direct RunCode command targets.
'     Help_Docs controls generated help-document shortcut macros and tooltips.
'     UI_Groups supplies configured custom group names used by settings, config rows, catalog rows, help documents, and form records; it also stores editable ribbon tab/group fields used by the custom ribbon
'     builder.
'     Backup objects are grouped only under Access Development Console - Backup by default.
'     MAC and related catalog tables control managed utility object import/generation/grouping, verification, conflict policy, backup policy, and failure behavior.
'     Exploding Kittens warning settings are CodeDb settings for destructive-potential warnings; they do not create a separate warning workflow by themselves.
' Behavior:
'     User-invoked setup clears and prints trace/status output, repairs CodeDb-owned utility objects, and reports final success/failure status.
'     Full Setup replaces the active rollback journal after C3, G1, H3, and I3 so only the currently executing phase group remains provisional.
'     Role-specific generation routing generates only accepted Active=True configured rows in the selected ConfigDb/GenerationTargetDb path.
'     Bulk generation collects row-level errors and reports them together instead of showing one message per failed row.
'     Candidate scan can discover public/implicit-public standard-module procedures and visible Access objects depending on the selected scan route.
'     Review/open procedures are intended to open review/configuration surfaces and should not silently perform destructive cleanup.
'     Progress helpers are used for potentially long-running setup, scan, review, generation, and cleanup workflows.
' Code Status:
'     Active core orchestration module. Public table-maintenance bodies remain owned by x__UTIL_ADC_TABLES; this module retains core orchestration, scan/review, generation, runtime dispatch, Navigation Pane
'     grouping, and form-button support.
' Todo:
' Reporting Bugs: None currently recorded in this header.
' Module Level Bugs: None currently recorded in this header.
' Public Subs:
'     CoreResetOperationNotes() -> None
'     CoreAddOperationNote(ByVal noteText As String) -> None
'     BeginOperationDetailTraceSuppression() -> None
'     EndOperationDetailTraceSuppression() -> None
'     CaptureTrackedProcessSummary() -> None
'     BeginTrackedProcess(ByVal processName As String) -> None
'     EndTrackedProcess() -> None
'     CoreResetSharedProgressDiagnostics() -> None
'     TrySharedProgressBegin(ByRef progressActive As Boolean, ByRef progressStep As Long, ByRef progressMax As Long, ByVal captionText As String, ByVal statusPrefix As String, ByVal sourceProcedure As String,
'     ByVal contextName As String, Optional ByVal componentName As String = vbNullString) -> None
'     TrySharedProgressUpdate(ByRef progressActive As Boolean, ByRef progressStep As Long, ByRef progressMax As Long, ByVal statusText As String, ByVal statusPrefix As String, ByVal sourceProcedure As String,
'     ByVal contextName As String, Optional ByVal componentName As String = vbNullString) -> None
'     TrySharedProgressFail(ByRef progressActive As Boolean, ByRef progressStep As Long, ByRef progressMax As Long, ByVal statusText As String, ByVal sourceProcedure As String, ByVal contextName As String,
'     Optional ByVal componentName As String = vbNullString) -> None
'     TrySharedProgressEnd(ByRef progressActive As Boolean, ByRef progressStep As Long, ByRef progressMax As Long, ByVal clearStatus As Boolean, ByVal sourceProcedure As String, ByVal contextName As String,
'     Optional ByVal componentName As String = vbNullString) -> None
'     CorePulseProgress(ByVal statusText As String) -> None
'     PulseProgressEvery(ByVal currentCount As Long, ByVal intervalCount As Long, ByVal statusText As String, Optional ByVal totalCount As Long = 0, Optional ByVal sourceProcedure As String =
'     "PulseProgressEvery", Optional ByVal contextName As String = "Core") -> None
'     CoreBeginProgressChildDetail(ByVal childScopeLabel As String) -> None
'     CoreEndProgressChildDetail() -> None
'     CoreBeginStandaloneSetupPartRun(ByVal partIdentity As String, ByRef moduleGuardVerificationDetails As String, ByRef traceRunOwnedByProcedure As Boolean, Optional ByVal operationIdentity As String =
'     vbNullString, Optional ByVal progressPartCount As Long = 1) -> None
'     CoreCompleteStandaloneSetupPartRun(ByVal partIdentity As String, Optional ByVal timingText As String = vbNullString, Optional ByVal showMessage As Boolean = False, Optional ByVal traceRunOwnedByProcedure
'     As Boolean = False, Optional ByVal operationIdentity As String = vbNullString) -> None
'     CoreHandleStandaloneSetupPartFailure(ByVal partIdentity As String, ByVal errNumber As Long, ByVal errSource As String, ByVal errDescription As String, Optional ByVal traceRunOwnedByProcedure As Boolean =
'     False, Optional ByVal operationIdentity As String = vbNullString) -> None
'     CoreUpdateSetupProgress(ByVal stepText As String) -> None
'     CoreCompleteSetupProgress() -> None
'     CoreFailSetupProgress() -> None
'     EndSetupProgress(Optional ByVal clearStatus As Boolean = True) -> None
'     CoreBeginSetupPhasePart(ByVal phasePartIdentity As String, ByVal phaseLabel As String, ByVal phasePartNumber As Long, ByVal phasePartCount As Long, ByVal stepText As String) -> None
'     RequireSetupDomainMutationAllowed(ByVal domainIdentity As String, ByVal sourceProcedure As String) -> None
'     RunFullSetup(Optional ByVal showMessage As Boolean = True) -> None
'     TrackSetupQueryRollback(ByVal queryName As String, ByVal targetExisted As Boolean, ByVal previousSql As String) -> None
'     TrackSetupCreatedTableRollback(ByVal tableName As String) -> None
'     TrackSetupAccessObjectRollback(ByVal actionText As String, ByVal objectName As String, ByVal targetExisted As Boolean, ByVal backupPath As String) -> None
'     PhaseCVerifyPhysicalSystemHelpMacroState(ByRef detailText As String, Optional ByVal synchronizationScope As HelpDocumentSynchronizationScope = HelpDocumentScopeSetupSystemAndAllActive, Optional ByVal
'     helpDocumentMutationsOnly As Boolean = False) -> None
'     PhaseCBuildRequiredSystemHelpMacros(ByRef detailText As String) -> None
'     PhaseCFinalizePhysicalSystemHelpMacroState(ByRef detailText As String) -> None
'     CoreBeginSetupNamingSnapshot() -> None
'     CoreMarkSetupNamingSnapshotDirty(Optional ByVal sourceDetail As String = vbNullString, Optional ByVal namingPolicyChanged As Boolean = False) -> None
'     CoreEnsureSetupNamingSnapshotCurrent() -> None
'     CorePulseModuleGuardPreflight(ByVal statusText As String) -> None
'     ReportBlockingPromptWait(ByVal sourceProcedure As String, ByVal promptReason As String) -> None
'     ReportBlockingPromptResult(ByVal sourceProcedure As String, ByVal promptReason As String, ByVal approved As Boolean) -> None
'     PersistDurableRecoveryMarker(ByVal recoveryKey As String, ByVal operationName As String, ByVal detailText As String, Optional ByVal payloadText As String = vbNullString) -> None
'     DeleteDurableRecoveryMarker(ByVal recoveryKey As String) -> None
'     ClearSettingsRuntimeStateForCurrentContext() -> None
'     EnsureSettingsReadyForRead() -> None
'     MarkSettingsReadyForCurrentContext() -> None
'     RequireSafeWorkContinue(ByVal workerRoutine As String, Optional ByVal logicalItem As String = vbNullString) -> None
'     CoreShowMessage(ByVal promptText As String, ByVal iconStyle As VbMsgBoxStyle, ByVal TitleText As String, Optional ByVal sourceProcedure As String = "ShowADCMessage", Optional ByVal sourceArea As String =
'     "Core") -> None
'     CoreShowErrorMessage(ByVal promptText As String, ByVal iconStyle As VbMsgBoxStyle, ByVal TitleText As String, Optional ByVal sourceProcedure As String = "ShowADCErrorMessage", Optional ByVal sourceArea As
'     String = "Core") -> None
'     ReportCoreActionSuccess( ByVal sourceProcedure As String, ByVal traceCategory As String, ByVal successMessage As String, Optional ByVal popupMessage As String = vbNullString, Optional ByVal stepName As
'     String = vbNullString, Optional ByVal objectName As String = vbNullString, Optional ByVal objectType As String = vbNullString, Optional ByVal databaseRole As String = "CodeDb", Optional ByVal contextName
'     As String = vbNullString, Optional ByVal showMessage As Boolean = False, Optional ByVal popupTitle As String = "ADC", Optional ByVal popupIcon As VbMsgBoxStyle = vbInformation, Optional ByVal
'     applyRunWarnings As Boolean = True, Optional ByVal reportingArea As String = "Core", Optional ByVal detectMessageWarnings As Boolean = False, Optional ByVal visibleStatusMessage As String = vbNullString)
'     -> None
'     CoreSetTableRepairWarningResult(ByVal hadWarnings As Boolean) -> None
'     ClearTableRepairWarningResult() -> None
'     AddTableRepairWarningToParent() -> None
'     CoreCommitSetupRollbackGroupBoundary(ByVal completedGroup As String, ByVal nextGroup As String) -> None
'     DeleteRollbackBackupFile(ByVal backupPath As String, ByRef failedPaths As String) -> None
'     CoreMarkInvokingMacro65DeletionDeferred(ByVal macroName As String, Optional ByVal completionKind As String = "Normal") -> None
'     CoreConfigureDeferredMacro65NormalCompletion(ByVal showMessage As Boolean, ByVal completionMessage As String, Optional ByVal completionTitle As String = "Delete All ADC Objects") -> None
'     CoreCancelDeferredMacro65DeletionRequest() -> None
'     CoreDeferredMacro65DeleteTimerProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal idEvent As LongPtr, ByVal dwTime As Long) -> None
'     CoreDeferredMacro65DeleteTimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long) -> None
'     OpenCoreRecoveryDatabase(ByRef recoveryWs As DAO.workspace, ByRef recoveryDb As DAO.Database) -> None
'     EnsureCoreRecoveryStorageReady() -> None
'     TrackSetupNavPaneShortcutRollback(ByVal groupID As Long, ByVal objectID As Long) -> None
'     TrackSetupNavPaneShortcutRestoreRollback(ByVal groupID As Long, ByVal objectID As Long, ByVal priorName As Variant, ByVal priorType As Variant, ByVal priorFlags As Variant, ByVal priorPosition As Variant)
'     -> None
'     TrackSetupNavPaneObjectIDRollback(ByVal objectID As Long, ByVal objectName As String, ByVal objectType As Long) -> None
'     TrackSetupNavPaneObjectIDRestoreRollback(ByVal objectID As Long, ByVal repairedObjectName As String, ByVal repairedObjectType As Long, ByVal priorObjectName As String, ByVal priorObjectType As Long) ->
'     None
'     TrackSetupDeletedNavPaneObjectIDRestoreRollback(ByVal objectID As Long, ByVal objectName As String, ByVal objectType As Long) -> None
'     TrackSetupNavPaneGroupRollback(ByVal groupID As Long, ByVal groupName As String) -> None
'     CoreStageFinalSetupStateForJ1(ByVal finalProofToken As String, ByVal moduleGuardStateToken As String, ByVal detailText As String) -> None
'     CoreBeginSetupRunThroughPhaseJ(ByVal moduleGuardVerificationDetails As String) -> None
'     CoreCompleteSetupRunThroughPhaseJ(Optional ByVal timingText As String = vbNullString, Optional ByVal showMessage As Boolean = False) -> None
'     CoreHandleSetupRunThroughPhaseJFailure(ByVal errNumber As Long, ByVal errSource As String, ByVal errDescription As String) -> None
'     MarkRibbonRestartRequiredForCurrentRun() -> None
'     MarkSystemCommandCancelled() -> None
'     RequireSameDatabaseUICommand(ByVal sourceProcedure As String, ByVal actionText As String) -> None
'     CoreScheduleDeferredMacro65Deletion(ByVal completeOwnedTrace As Boolean, ByVal traceSuccessMessage As String, ByVal traceFailureMessage As String) -> None
' Public Functions:
'     PartCExtractRunCodeArgument(ByVal macroText As String) -> String
'     CoreRecoveryEncode(ByVal sourceText As String) -> String
'     CoreRecoveryDecode(ByVal sourceText As String) -> String
'     CoreRecoveryNoteValue(ByVal notesText As String, ByVal keyName As String) -> String
'     Run_RunFullSetup() -> Boolean
'     ADCDb() -> DAO.Database
'     GetHostDb() -> DAO.Database
'     DatabaseForRole(ByVal ConfigOwnerDatabaseRole As String) -> DAO.Database
'     DatabasePathForRole(ByVal ConfigOwnerDatabaseRole As String) -> String
'     VBProjectForConfigOwnerDatabaseRole(ByVal ConfigOwnerDatabaseRole As String) -> Object
'     CodeDbIsHostDb() -> Boolean
'     DatabaseIsCurrentProject(ByVal db As DAO.Database) -> Boolean
'     CoreOperationHasWarnings() -> Boolean
'     CoreOperationSummaryForMessage() -> String
'     ProcessRunSummaryForMessage(ByVal sourceProcedure As String) -> String
'     FailedSetupPhaseContextForMessage() -> String
'     CoreParentProgressActive() -> Boolean
'     CoreParentNestedProgressText(ByVal statusText As String, Optional ByVal childScopeText As String = vbNullString) -> String
'     CoreQualifyProgressStatusText(ByVal statusText As String) -> String
'     CoreVisibleProgressText(ByVal statusText As String) -> String
'     CoreProgressPerformanceDiagnosticsSummary() -> String
'     SharedLatestProgressStatusText() -> String
'     CoreAppendProcessRunSummary(ByVal messageText As String, ByVal sourceProcedure As String) -> String
'     CoreBeginOperationProgressIfNeeded(ByVal progressTitle As String, Optional ByVal maxSteps As Long = 0, Optional ByVal progressScopeLabel As String = vbNullString, Optional ByVal forceNewTopLevelOperation
'     As Boolean = False) -> Boolean
'     A2PhysicalSchemaWritesAllowedForCurrentSetupPhase() -> Boolean
'     B1CatalogStructureWritesAllowedForCurrentSetupPhase() -> Boolean
'     SetupMutationWritesAllowedForCurrentPart(ByVal expectedPartIdentity As String) -> Boolean
'     SetupDomainOwnerPart(ByVal domainIdentity As String) -> String
'     SetupDomainMutationWritesAllowedForCurrentPart(ByVal domainIdentity As String) -> Boolean
'     PhaseC3CurrentForSetup(ByRef detailText As String) -> Boolean
'     PrepareSystemHelpMacroDeletionPlan(Optional ByRef plannedCountResult As Long = 0) -> Collection
'     DeleteSystemHelpMacros(Optional ByVal showMessage As Boolean = True, Optional ByVal confirmDestructiveAction As Boolean = True, Optional ByRef plannedCountResult As Long = 0, Optional ByRef
'     attemptedCountResult As Long = 0, Optional ByRef deletedCountResult As Long = 0, Optional ByRef skippedCountResult As Long = 0, Optional ByRef failedCountResult As Long = 0, Optional ByVal
'     preparedDeletionPlan As Variant) -> Boolean
'     CoreDurableRecoveryMarkerExistsInSnapshot(ByVal recoverySnapshot As Object, ByVal recoveryKey As String, ByRef operationName As String, ByRef detailText As String, ByRef payloadText As String) -> Boolean
'     CoreDurableRecoveryMarkerExists(ByVal recoveryKey As String, ByRef operationName As String, ByRef detailText As String, ByRef payloadText As String) -> Boolean
'     CoreSetupHelpDocumentSynchronizationScope() -> HelpDocumentSynchronizationScope
'     IsCancellationError(ByVal errorNumber As Long) -> Boolean
'     SafeDatabaseName() -> String
'     SafeTableName(ByVal tableRole As String) -> String
'     CoreShowMessageResult(ByVal promptText As String, ByVal iconStyle As VbMsgBoxStyle, ByVal TitleText As String, Optional ByVal sourceProcedure As String = "ShowADCMessageResult", Optional ByVal sourceArea
'     As String = "Core") -> VbMsgBoxResult
'     SetupRollbackTrackingActive() -> Boolean
'     RunSystemCommand(ByVal systemMacroKey As String) -> Boolean
'     CoreInvokingMacro65ObjectName() -> String
'     CoreShouldDeferInvokingMacro65Deletion(ByVal macroName As String) -> Boolean
'     CoreDeferredMacro65DeletionRequired() -> Boolean
'     CoreDeferredMacro65DeletionPending() -> Boolean
'     CoreDeferredMacro65DeletionCompletionKind() -> String
'     CoreIsDeferredInvokingMacro65Name(ByVal macroName As String) -> Boolean
'     TrackSetupManagedObjectRowRollback(ByVal managedObjectID As Long, ByVal priorRowExisted As Boolean, ByVal rollbackSql As String) -> Boolean
'     TrackSetupManagedObjectDependenciesRollback(ByVal rollbackSql As String) -> None
'     ValidationResultFromTargetStatus(ByVal statusText As String) -> String
'     EntryPointErrorMessage(ByVal headlineText As String, ByVal ProcedureName As String, ByVal errNumber As Long, ByVal errSource As String, ByVal errDescription As String, Optional ByVal detailText As String =
'     vbNullString) -> String
'     SystemCommandWasCancelled() -> Boolean
' Public Properties: None
' ================================================================================================================================================================================
Option Compare Database
Option Explicit

' ===============================================================================================================================================================================
' MODULE-LEVEL DECLARATIONS
' ===============================================================================================================================================================================
' System-macro inventory properties.
Public Const SYSTEM_MACRO_SET_FINGERPRINT_PROPERTY As String = "ADCSystemMacroSetFingerprint"
Private Const SYSTEM_MACRO_INSTALLED_INVENTORY_TOKEN_PROPERTY As String = "ADCSystemMacroInstalledInventoryToken"
Private Const SYSTEM_MACRO_EXPECTED_COUNT_PROPERTY As String = "ADCSystemMacroExpectedCount"
' Phase C operation-scoped physical evidence item contract.
Private Const PART_C_EVIDENCE_NAME_INDEX As Long = 0
Private Const PART_C_EVIDENCE_EXISTED_INDEX As Long = 1
Private Const PART_C_EVIDENCE_RUNCODE_INDEX As Long = 2
Private Const PART_C_EVIDENCE_EXPECTED_INDEX As Long = 3
Private Const PART_C_EVIDENCE_STALE_OWNED_INDEX As Long = 4
' Shared System Command result-state identity and values.
Private Const TEMPVAR_SYSTEM_COMMAND_RESULT As String = "ADCSystemCommandResult"
Private Const SYSTEM_COMMAND_RESULT_CANCELLED As String = "Cancelled"
' Explicit System Macro #65 deletion scopes.
Public Enum ObjectDeletionScope
    DELETE_SCOPE_NOT_SELECTED = -1
    DELETE_SCOPE_CANCEL = 0
    DELETE_SCOPE_GENERATED_FORMS = 2
    DELETE_SCOPE_ALL = 3
    DELETE_SCOPE_TABLES = 4
    DELETE_SCOPE_QUERIES = 5
    DELETE_SCOPE_MACROS = 6
    DELETE_SCOPE_REPORTS = 7
End Enum
' Shared object limits, catalog types, and statuses.
Public Const ADCAT_STATUS_ACTIVE As String = "Active"
Public Const ADCAT_OBJECT_TYPE_MACRO As String = "Macro"
Public Const ADCAT_OBJECT_TYPE_FORM As String = "Form"
Public Const ADCAT_OBJECT_TYPE_QUERY As String = "Query"
Public Const ADCAT_OBJECT_TYPE_TABLE As String = "Table"
Public Const ADCAT_OBJECT_TYPE_REPORT As String = "Report"
Public Const ADCAT_STATUS_INACTIVE As String = "Inactive"
Public Const ADCAT_STATUS_NEEDS_REVIEW As String = "NeedsReview"
Public Const ADCAT_STATUS_BLOCKED As String = "Blocked"
' Managed database-structure identity strings are centralized in x__UTIL_ADC_DB_DEFINITIONS.
' VBE constants and run-scoped temporary-variable names.
Public Const VBEXT_CT_STD_MODULE As Long = 1
Public Const VBEXT_PK_PROC As Long = 0
Private Const TEMPVAR_TABLE_REPAIR_WARNINGS As String = "ADCLastTableRepairHadWarnings"
' Settings keys and parent-progress temporary variables.
Public Const SETTING_ENABLE_MODULE_UPDATE_GUARD As String = "EnableModuleUpdateGuard"
Public Const SETTING_GENERATE_SETUP_POPUP_MESSAGES As String = "GenerateSetupPopupMessages"
Public Const SETTING_ENABLE_CUSTOM_MACRO_LAUNCHER_AUTOMATION As String = "EnableCustomMacroLauncherAutomation"
Public Const SETTING_CUSTOM_MACRO_LAUNCHER_WARNING_ACCEPTED As String = "CustomMacroLauncherPreAlphaWarningAccepted"
Public Const SETTING_CUSTOM_MACRO_LAUNCHER_WARNING_ACCEPTED_ON As String = "CustomMacroLauncherPreAlphaWarningAcceptedOn"
Public Const SETTING_ADD_GENERATED_MACROS_TO_CUSTOM_MACRO_LAUNCHER As String = "AddGeneratedMacrosToCustomMacroLauncher"
Public Const SETTING_MACRO_LAUNCHER_APPEND_ANCHOR_CONTROL_NAME As String = "MacroLauncherAppendAnchorControlName"
Public Const SETTING_MACRO_LAUNCHER_BUTTON_WIDTH As String = "MacroLauncherButtonWidth"
Public Const SETTING_MACRO_LAUNCHER_BUTTON_HEIGHT As String = "MacroLauncherButtonHeight"
Public Const SETTING_MACRO_LAUNCHER_BUTTON_GAP_X As String = "MacroLauncherButtonGapX"
Public Const SETTING_MACRO_LAUNCHER_BUTTON_GAP_Y As String = "MacroLauncherButtonGapY"
Public Const SETTING_MACRO_LAUNCHER_BUTTONS_PER_ROW As String = "MacroLauncherButtonsPerRow"
Public Const SETTING_MACRO_LAUNCHER_USE_CONFIRMED_RUNNER As String = "MacroLauncherUseConfirmedRunner"
Private Const PROGRESS_TEMPVAR_ACTIVE As String = "ADCParentProgressActive"
Private Const PROGRESS_TEMPVAR_SCOPE As String = "ADCParentProgressScope"
Private Const PROGRESS_TEMPVAR_STEP As String = "ADCParentProgressStep"
Private Const PROGRESS_TEMPVAR_MAX As String = "ADCParentProgressMax"
Public Const SETTING_DATA_TYPE_STRING As String = "String"
Public Const SETTING_DATA_TYPE_BOOLEAN As String = "Boolean"
Public Const SETTING_DATA_TYPE_LONG As String = "Long"
Public Const SETTING_DATA_TYPE_DATE As String = "Date"
Public Const SETTING_DATA_TYPE_DOUBLE As String = "Double"
Public Const SETTING_DATA_TYPE_CURRENCY As String = "Currency"
Public Const SETTING_DIRECT_FALLBACK_KEY_PREFIX As String = "SettingKey: "
Public Const SETTING_IMMEDIATE_TRACE_VERBOSITY As String = "ImmediateTraceVerbosity"
Public Const SETTING_LOG_POPUP_LIFECYCLE_TO_IMMEDIATE As String = "LogPopupLifecycleToImmediate"
Public Const SETTING_TRACE_LOG_FILE_MAX_LINE_LENGTH As String = "TraceLogFileMaxLineLength"
' Macro execution, confirmation, procedure, result, data-type, and return-type values.
Public Const ACTION_RUNCODE As String = "RunCode"
Public Const ACTION_OPENFORM As String = "OpenForm"
Public Const ACTION_OPENTABLE As String = "OpenTable"
Public Const ACTION_OPENQUERY As String = "OpenQuery"
Public Const ACTION_OPENREPORT As String = "OpenReport"
Public Const ACTION_RUNMACRO As String = "RunMacro"
Public Const CONFIRM_AUTO As String = "Auto"
Public Const CONFIRM_ALWAYS As String = "Always"
Public Const CONFIRM_NEVER As String = "Never"
Public Const CONFIRM_AUTOEXEC As String = "AutoExec"
Public Const TARGET_PROCEDURE_KIND_AUTO As String = "Auto"
Public Const TARGET_PROCEDURE_KIND_FUNCTION As String = "Function"
Public Const TARGET_PROCEDURE_KIND_SUB As String = "Sub"
Public Const RESULT_HANDLING_AUTO As String = "Auto"
Public Const RESULT_HANDLING_IGNORE As String = "Ignore"
Public Const RESULT_HANDLING_BOOLEAN As String = "Boolean"
Public Const RESULT_HANDLING_REQUIRE_TRUE As String = "RequireTrue"
Public Const EXPECTED_RETURN_TYPE_AUTO As String = "Auto"
Public Const EXPECTED_RETURN_TYPE_BOOLEAN As String = "Boolean"
Public Const EXPECTED_RETURN_TYPE_STRING As String = "String"
Public Const EXPECTED_RETURN_TYPE_LONG As String = "Long"
Public Const EXPECTED_RETURN_TYPE_DOUBLE As String = "Double"
Public Const EXPECTED_RETURN_TYPE_CURRENCY As String = "Currency"
Public Const EXPECTED_RETURN_TYPE_DATE As String = "Date"
Public Const EXPECTED_RETURN_TYPE_VARIANT As String = "Variant"
Public Const EXPECTED_RETURN_TYPE_NONE As String = "None"
' Trace levels and verbosity values.
Public Const TRACE_LEVEL_DETAIL As String = "Detail"
Public Const TRACE_LEVEL_INFO As String = "Info"
Public Const TRACE_LEVEL_WARNING As String = "Warning"
Public Const TRACE_LEVEL_ERROR As String = "Error"
Public Const TRACE_LEVEL_STATUS As String = "Status"
Public Const TRACE_VERBOSITY_NORMAL As String = "Normal"
Public Const TRACE_VERBOSITY_VERBOSE As String = "Verbose"
Public Const TRACE_VERBOSITY_QUIET As String = "Quiet"
' Navigation Pane group-purpose values.
Public Const NAVIGATION_PANE_GROUP_PURPOSE_BACKUP As String = "Backup"
Public Const NAVIGATION_PANE_GROUP_PURPOSE_GENERATED_MACROS As String = "GeneratedMacros"
Public Const NAVIGATION_PANE_GROUP_PURPOSE_SYSTEM_MACROS As String = "SystemMacroCatalog"
Public Const NAVIGATION_PANE_GROUP_PURPOSE_FORMS As String = "Forms"
Public Const NAVIGATION_PANE_GROUP_PURPOSE_CONFIG As String = "Config"
Public Const NAVIGATION_PANE_GROUP_PURPOSE_SETTINGS As String = "Settings"
Public Const NAVIGATION_PANE_GROUP_PURPOSE_OBJECT_DEFINITIONS As String = "ObjectDefinitions"
Public Const NAVIGATION_PANE_GROUP_PURPOSE_FORM_DESIGN As String = "FormDesign"
Public Const NAVIGATION_PANE_GROUP_PURPOSE_NAVIGATION_RIBBON As String = "NavigationRibbon"
Public Const NAVIGATION_PANE_GROUP_PURPOSE_SETUP_MAINTENANCE As String = "SetupMaintenance"
Public Const NAVIGATION_PANE_GROUP_PURPOSE_DIAGNOSTICS As String = "Diagnostics"
Public Const NAVIGATION_PANE_GROUP_PURPOSE_ADD_INS As String = "AddIns"
' Configuration, review, validation, generation, candidate, and health statuses.
Public Const ROW_KIND_CONFIGURATION As String = "Configuration"
Public Const ROW_KIND_CANDIDATE As String = "Candidate"
Public Const CONFIGURATION_STATUS_ACTIVE As String = "Active"
Public Const CONFIGURATION_STATUS_NEEDS_REVIEW As String = "NeedsReview"
Public Const CONFIGURATION_STATUS_INVALID As String = "Invalid"
Public Const CONFIGURATION_STATUS_DISABLED As String = "Disabled"
Public Const CATALOG_REVIEW_STATUS_READY As String = "Ready"
Public Const CATALOG_REVIEW_STATUS_BLOCKING As String = "Blocking"
Public Const CATALOG_REVIEW_STATUS_INACTIVE_OR_REVIEW As String = "InactiveOrReview"
Public Const VALIDATION_RESULT_VALID As String = "Valid"
Public Const VALIDATION_RESULT_WARNING As String = "Warning"
Public Const VALIDATION_RESULT_INVALID As String = "Invalid"
Public Const VALIDATION_RESULT_NEEDS_REVIEW As String = "NeedsReview"
Public Const VALIDATION_RESULT_SKIPPED As String = "Skipped"
Public Const VALIDATION_RESULT_ERROR As String = "Error"
Public Const GENERATED_RESULT_GENERATED As String = "Generated"
Public Const GENERATED_RESULT_SKIPPED As String = "Skipped"
Public Const GENERATED_RESULT_FAILED As String = "Failed"
Public Const GENERATED_RESULT_RESTORED As String = "Restored"
Public Const GENERATED_RESULT_NEEDS_REVIEW As String = "NeedsReview"
Public Const GENERATED_RESULT_INVALID As String = "Invalid"
Public Const MANUAL_REVIEW_STATUS_NEEDS_REVIEW As String = "NeedsReview"
Public Const MANUAL_REVIEW_STATUS_REVIEWED As String = "Reviewed"
Public Const MANUAL_REVIEW_STATUS_DEFERRED As String = "Deferred"
Public Const MANUAL_REVIEW_STATUS_IGNORE As String = "Ignore"
Public Const MANUAL_REVIEW_STATUS_BLOCKED As String = "Blocked"
Public Const HEALTH_SEVERITY_INFO As String = "Info"
Public Const HEALTH_SEVERITY_WARNING As String = "Warning"
Public Const HEALTH_SEVERITY_ERROR As String = "Error"
Public Const HEALTH_SEVERITY_CRITICAL As String = "Critical"
Public Const REVIEW_AREA_SCHEMA As String = "Schema"
Public Const REVIEW_AREA_SETTINGS As String = "Settings"
Public Const REVIEW_AREA_CONFIG As String = "Config"
Public Const REVIEW_AREA_CANDIDATE As String = "Candidate"
Public Const REVIEW_AREA_DUPLICATE As String = "Duplicate"
Public Const REVIEW_AREA_OBJECT_OWNERSHIP As String = "ObjectOwnership"
Public Const REVIEW_AREA_CLEANUP_METADATA As String = "CleanupMetadata"
Public Const REVIEW_AREA_TARGET As String = "Target"
Public Const REVIEW_AREA_VALIDATION As String = "Validation"
Public Const REVIEW_AREA_GENERATION As String = "Generation"
Public Const REVIEW_AREA_GENERATED_MACRO As String = "GeneratedMacro"
Public Const REVIEW_AREA_SYSTEM_MACRO As String = "SystemMacro"
Public Const REVIEW_AREA_PUBLIC_ENTRY_POINT As String = "PublicEntryPoint"
Public Const CANDIDATE_STATUS_NEW As String = "New"
Public Const CANDIDATE_STATUS_STALE As String = "Stale"
Public Const CANDIDATE_STATUS_ACCEPTED As String = "Accepted"
Public Const CANDIDATE_STATUS_REJECTED As String = "Rejected"
Public Const CANDIDATE_STATUS_DUPLICATE As String = "Duplicate"
Public Const CANDIDATE_STATUS_INVALID As String = "Invalid"
Public Const CANDIDATE_STATUS_NEEDS_ARGUMENTS As String = "NeedsArguments"
Public Const CANDIDATE_STATUS_CONFIGURED As String = "Configured"
Public Const CANDIDATE_STATUS_SCAN_ERROR As String = "ScanError"
Public Const DUPLICATE_ISSUE_CONFIG_MACRO_KEY As String = "ConfigMacroKeyDuplicate"
Public Const DUPLICATE_ISSUE_ACTIVE_MACRO_TITLE As String = "ActiveMacroTitleDuplicate"
Public Const ACCEPTANCE_STATUS_ACCEPTED As String = "Accepted"
Public Const ACCEPTANCE_STATUS_ACCEPTED_NEEDS_ARGUMENTS As String = "AcceptedNeedsArguments"
Public Const ACCEPTANCE_STATUS_SKIPPED_ALREADY_ACCEPTED As String = "SkippedAlreadyAccepted"
Public Const ACCEPTANCE_STATUS_DUPLICATE As String = "Duplicate"
Public Const ACCEPTANCE_STATUS_INVALID As String = "Invalid"
Public Const ACCEPTANCE_STATUS_ERROR As String = "Error"
' Setup failure text, rollback action identifiers, and progress completion text.
Private Const SETUP_RESULT_FAILURE_HEADLINE As String = "Access Development Console: Setup/Update failed."
Private Const SETUP_RESULT_FAILURE_TRACE_NOTE As String = "Review the ADCTrace lines and full setup diagnostics in the VBA Immediate Window."
Public Const SETUP_ROLLBACK_RESTORE_MACRO_ACTION As String = "Macro"
Private Const SETUP_ROLLBACK_RESTORE_FORM_ACTION As String = "Form"
Public Const SETUP_ROLLBACK_RESTORE_REPORT_ACTION As String = "Report"
Private Const SETUP_ROLLBACK_RESTORE_QUERY_ACTION As String = "Query"
Private Const SETUP_ROLLBACK_DELETE_TABLE_ACTION As String = "Table"
Private Const SETUP_ROLLBACK_DELETE_NAV_SHORTCUT_ACTION As String = "NavShortcut"
Private Const SETUP_ROLLBACK_RESTORE_NAV_SHORTCUT_ACTION As String = "NavShortcutRestore"
Private Const SETUP_ROLLBACK_DELETE_NAV_OBJECT_ID_ACTION As String = "NavObjectID"
Private Const SETUP_ROLLBACK_RESTORE_NAV_OBJECT_ID_ACTION As String = "NavObjectIDRestore"
Private Const SETUP_ROLLBACK_RESTORE_DELETED_NAV_OBJECT_ID_ACTION As String = "NavObjectIDDeletedRestore"
Private Const SETUP_ROLLBACK_DELETE_NAV_GROUP_ACTION As String = "NavGroup"
Private Const SETUP_ROLLBACK_RESTORE_MANAGED_OBJECT_ROW_ACTION As String = "ManagedObjectRow"
Private Const SETUP_ROLLBACK_RESTORE_MANAGED_DEPENDENCIES_ACTION As String = "ManagedDependencies"
Private Const OPERATION_PROGRESS_COMPLETE_STATUS As String = "ADC process complete."
Private Const OPERATION_PROGRESS_FAILED_STATUS As String = "ADC process failed. Review the Immediate Window trace."
' Trace-file and managed-catalog metadata.
Public Const MODULE_TAG_DELIMITER As String = "|"
Public Const TRACE_LOG_FILE_NAME As String = "Access Development Console Trace.log"
Public Const TRACE_TABLE_WRITE_FAILURE_SUMMARY_INTERVAL As Long = 25
Public Const MANAGED_OBJECT_CATALOG_SCHEMA_VERSION As String = "4.6"
' Managed Object Catalog policies, classifications, and ownership values.
Public Const ADCAT_SETUP_VERIFY_CODEDB_OBJECT As String = "VerifyCodeDbObject"
Public Const ADCAT_SETUP_CREATE_OR_REPAIR_CODEDB As String = "CreateOrRepairCodeDb"
Public Const ADCAT_SETUP_GENERATE_FROM_CONFIG As String = "GenerateFromConfig"
Public Const ADCAT_SETUP_GROUP_EXISTING As String = "GroupExisting"
Public Const ADCAT_SETUP_REFERENCE_ONLY As String = "ReferenceOnly"
Public Const ADCAT_SETUP_DATA_TABLE As String = "DataTable"
Public Const ADCAT_NAV_POLICY_NONE As String = "None"
Public Const ADCAT_NAV_POLICY_ADD_TO_CONFIGURED_GROUP As String = "AddToConfiguredGroup"
Public Const ADCAT_NAV_POLICY_COPY_EXISTING_SHORTCUT_TO_CONFIGURED_GROUP As String = "CopyExistingShortcutToConfiguredGroup"
Public Const ADCAT_ENV_RUNTIME As String = "Runtime"
Public Const ADCAT_ENV_DEVELOPMENT As String = "Development"
Public Const ADCAT_ENV_REVIEW_ONLY As String = "ReviewOnly"
Public Const ADCAT_SETUP_STAGE_BOOTSTRAP As String = "Bootstrap"
Public Const ADCAT_SETUP_STAGE_TABLES As String = "Tables"
Public Const ADCAT_SETUP_STAGE_SYSTEM_MACROS As String = "ADC_SystemMacros"
Public Const ADCAT_SETUP_STAGE_MANAGED_UTILITY_OBJECTS As String = "ManagedUtilityObjects"
Public Const ADCAT_SETUP_STAGE_FORMS As String = "Forms"
Public Const ADCAT_SETUP_STAGE_QUERIES As String = "Queries"
Public Const ADCAT_SETUP_STAGE_REVIEW As String = "Review"
Public Const ADCAT_SETUP_STAGE_LAUNCHER As String = "Launcher"
Public Const ADCAT_SETUP_STAGE_NAVIGATION_PANE_GROUPING As String = "NavigationPaneGrouping"
Public Const ADCAT_SETUP_STAGE_DEFERRED As String = "Deferred"
Public Const ADCAT_PURPOSE_SYSTEM_MACRO As String = "SystemMacro"
Public Const ADCAT_PURPOSE_REVIEW_QUERY As String = "ReviewQuery"
Public Const ADCAT_PURPOSE_FORM_DESIGN_QUERY As String = "FormDesignQuery"
Public Const ADCAT_PURPOSE_UTILITY_SUPPORT As String = "UtilitySupport"
Public Const ADCAT_PURPOSE_RUNTIME_DATA As String = "RuntimeData"
Public Const ADCAT_PURPOSE_CATALOG_DATA As String = "CatalogData"
Public Const ADCAT_PURPOSE_LAUNCHER_SUPPORT As String = "LauncherSupport"
Public Const ADCAT_PURPOSE_REPORT_SUPPORT As String = "ReportSupport"
Public Const ADCAT_EXPECT_SHOULD_EXIST As String = "ShouldExist"
Public Const ADCAT_EXPECT_MAY_EXIST As String = "MayExist"
Public Const ADCAT_DEPENDENCY_TYPE_CATALOG_OBJECT As String = "CatalogObject"
Public Const ADCAT_DEPENDENCY_TYPE_ACCESS_OBJECT As String = "AccessObject"
Public Const ADCAT_DEPENDENCY_TYPE_VBA_PROCEDURE As String = "VBAProcedure"
Public Const ADCAT_DEPENDENCY_TYPE_SETTING_KEY As String = "SettingKey"
Public Const ADCAT_PAYLOAD_ACCESS_MACRO_RUNCODE As String = "AccessMacroRunCode"
Public Const ADCAT_PAYLOAD_QUERY_SQL_BUILDER As String = "AccessQuerySqlBuilder"
Public Const ADCAT_CATEGORY_SYSTEM_MACRO As String = "BuiltInSystemMacro"
Public Const ADCAT_CATEGORY_MANAGED_QUERY As String = "ManagedUtilityQuery"
Public Const ADCAT_CATEGORY_MANAGED_TABLE As String = "ManagedUtilityTable"
Public Const ADCAT_CATEGORY_MANAGED_FORM As String = "ManagedUtilityForm"
Public Const ADCAT_CATEGORY_MANAGED_REPORT As String = "ManagedUtilityReport"
Public Const ADCAT_CATEGORY_RUNTIME_TABLE As String = "RuntimeTable"
Public Const ADCAT_CATEGORY_CATALOG_TABLE As String = "CatalogTable"
Public Const ADCAT_CATEGORY_FORM_QUERY As String = "FormDesignQuery"
Public Const ADCAT_CATEGORY_REVIEW_QUERY As String = "ReviewQuery"
Public Const ADCAT_CATEGORY_REVIEW_TABLE As String = "ReviewTable"
' Generated-form, macro-form, help-document, and catalog lookup metadata.
Public Const GENERATED_FORM_SOURCE_TYPE_UNBOUND As String = "Unbound"
Public Const GENERATED_FORM_SOURCE_TYPE_TABLE_OR_QUERY As String = "TableOrQuery"
Public Const GENERATED_FORM_LOOKUP_KIND_TWO_COLUMN As String = "TwoColumn"
Public Const GENERATED_FORM_LOOKUP_KIND_LIVE_TARGET As String = "LiveTarget"
Public Const GENERATED_FORM_LOOKUP_KIND_LIVE_TABLE_QUERY As String = "LiveTableQuery"
Public Const GENERATED_FORM_LOOKUP_KIND_DISTINCT_FIELD As String = "DistinctField"
Public Const GENERATED_FORM_LOOKUP_KIND_LIVE_ONE_COLUMN As String = "LiveOneColumn"
Public Const GENERATED_FORM_LOOKUP_KIND_VALUE_LIST As String = "ValueList"
Public Const CATALOG_FIELD_LOOKUP_DEFINITION_KIND_VALUE_LIST As String = "ValueList"
Public Const CATALOG_FIELD_LOOKUP_DEFINITION_KIND_TABLE As String = "Table"
Public Const CATALOG_FIELD_LOOKUP_DEFINITION_KIND_LIVE As String = "Live"
Public Const MACRO_FORM_TYPE_DATASHEET As String = "Datasheet"
Public Const MACRO_FORM_TYPE_SPLIT As String = "Split"
Public Const HELP_DOCUMENT_ERROR_TITLE As String = "Access Development Console Help"
Public Const HELP_DOCUMENT_PATH_UNDEFINED_MESSAGE As String = "Error. Help document path undefined."
Public Const HELP_DOCUMENT_NOT_FOUND_MESSAGE As String = "Error. Help document not found."
Public Const HELP_DOCUMENT_OPEN_FAILED_MESSAGE As String = "Error. Unable to open help document."
Public Const HELP_DOCUMENT_PATH_LABEL As String = "Help document path: "
Public Const HELP_DOCUMENT_RELATIVE_PATH_LABEL As String = "Configured relative path: "
Public Const HELP_DOCUMENT_BACKSLASH As String = "\"
' Setup recovery, final-state, evidence, and error identifiers.
Private Const ERR_SETUP_ROLLBACK_RECOVERY_FAILED As Long = vbObjectError + 6361
Private Const ERR_DURABLE_RECOVERY_FAILED As Long = vbObjectError + 6362
Private Const ERR_UI_PREFLIGHT_RECOVERY_FAILED As Long = vbObjectError + 6364
Private Const ERR_SETUP_CANCELLED_BY_USER As Long = vbObjectError + 6498
Private Const SETUP_RECOVERY_KEY As String = "ADC.Core.SetupRollbackRecoveryV1"
Private Const SETUP_ROLLBACK_ACTIVE_DETAIL As String = "Setup rollback tracking is active."
Private Const SETUP_POST_COMMIT_CHECKPOINT_KEY As String = "ADC.Core.SetupPostCommitCheckpointV1"
Private Const SETUP_RUN_CONTEXT_SCHEMA_VERSION As String = "2026-07-21 setup-run context v2 structured evidence"
Private Const SETUP_EVIDENCE_SCHEMA_VERSION As String = "2026-07-21 setup evidence v1"
Private Const FINAL_SETUP_STATE_VERSION As String = "ADC-J1-Final-State-v1"
Private Const FINAL_SETUP_STATE_VERSION_PROPERTY As String = "ADCFinalSetupStateVersion"
Private Const FINAL_SETUP_STATE_TOKEN_PROPERTY As String = "ADCFinalSetupStateToken"
Private Const FINAL_SETUP_CHECKPOINT_PROPERTY As String = "ADCFinalSetupCheckpoint"
Private Const FINAL_SETUP_COMPLETED_PART_PROPERTY As String = "ADCFinalSetupCompletedPart"
Private Const FINAL_SETUP_COMPLETED_ON_PROPERTY As String = "ADCFinalSetupCompletedOn"
Private Const FINAL_SETUP_MODULE_GUARD_PROPERTY As String = "ADCFinalSetupModuleGuardState"
Private Const SETUP_EVIDENCE_STATUS_MATCHED As String = "Matched"
Private Const SETUP_EVIDENCE_STATUS_CHANGED As String = "Changed"
Private Const SETUP_EVIDENCE_STATUS_MISSING As String = "Missing"
Private Const SETUP_EVIDENCE_STATUS_UNSEEDED As String = "Unseeded"
Private Const SETUP_EVIDENCE_STATUS_PROOF_UNAVAILABLE As String = "ProofUnavailable"
Private Const SETUP_EVIDENCE_STATUS_INVALID_DEFINITION As String = "InvalidDefinition"
Private Const SETUP_EVIDENCE_STATUS_MUTATION_REQUIRED As String = "MutationRequired"
Private Const SETUP_EVIDENCE_STATUS_MUTATION_COMPLETED As String = "MutationCompleted"
Private Const SETUP_EVIDENCE_STATUS_POSTCONDITION_VERIFIED As String = "PostconditionVerified"
Private Const SETUP_EVIDENCE_STATUS_DEFERRED_TO_OWNER As String = "DeferredToOwner"
Private Const SETUP_EVIDENCE_STATUS_FAILED As String = "Failed"
Public Const RECOVERY_STATUS_BLOCKED As String = "Blocked"
' Canonical Setup mutation-domain identities. These are semantic ownership labels; callers name the state being changed and Core resolves the only numbered-Setup Part allowed to mutate it.
Public Const SETUP_DOMAIN_A2_PHYSICAL_SCHEMA As String = "A2PhysicalSchema"
Public Const SETUP_DOMAIN_A3_SAVED_QUERY_FOUNDATION As String = "A3SavedQueryFoundation"
Public Const SETUP_DOMAIN_A4_REFERENCE_INVENTORY As String = "A4ReferenceInventory"
Public Const SETUP_DOMAIN_B1_CATALOG_STRUCTURE As String = "B1CatalogStructure"
Public Const SETUP_DOMAIN_B2_CATALOG_REFERENCE_DEFAULTS As String = "B2CatalogReferenceDefaults"
Public Const SETUP_DOMAIN_B3_NONFORM_MANAGED_DEFINITIONS As String = "B3NonFormManagedDefinitions"
Public Const SETUP_DOMAIN_B4_MACRO_MANAGED_DEFINITIONS As String = "B4MacroManagedDefinitions"
Public Const SETUP_DOMAIN_B5_CATALOG_SUPPORT As String = "B5CatalogSupport"
Public Const SETUP_DOMAIN_C2_PHYSICAL_MACRO_BUILD As String = "C2PhysicalMacroBuild"
Public Const SETUP_DOMAIN_C3_PHYSICAL_MACRO_FINALIZATION As String = "C3PhysicalMacroFinalization"
Public Const SETUP_DOMAIN_D1_PERSISTENT_LAUNCHER_FORMS As String = "D1PersistentLauncherForms"
Public Const SETUP_DOMAIN_D2_CHATGPT_UI_FORMS As String = "D2ChatGPTUIForms"
Public Const SETUP_DOMAIN_D3_SUPPORT_FORMS As String = "D3SupportForms"
Public Const SETUP_DOMAIN_D4_MANAGED_FORM_LOGICAL_STATE As String = "D4ManagedFormLogicalState"
Public Const SETUP_DOMAIN_D4_LOOKUP_EDITOR_BINDING As String = "D4LookupEditorBinding"
Public Const SETUP_DOMAIN_D4_MANAGED_FORM_PROOF As String = "D4ManagedFormProof"
Public Const SETUP_DOMAIN_D4_COMPLETE_CATALOG_FINALIZATION As String = "D4CompleteCatalogFinalization"
Public Const SETUP_DOMAIN_E2_PHYSICAL_REPORTS As String = "E2PhysicalReports"
Public Const SETUP_DOMAIN_E3_REPORT_FINALIZATION As String = "E3ReportFinalization"
Public Const SETUP_DOMAIN_F1_MANAGED_DEPENDENCIES As String = "F1ManagedDependencies"
Public Const SETUP_DOMAIN_F1_DERIVED_CATALOG_REFRESH As String = "F1DependencyDerivedCatalogRefresh"
Public Const SETUP_DOMAIN_G1_DESCRIPTIONS As String = "G1Descriptions"
Public Const SETUP_DOMAIN_G1_DERIVED_CATALOG_REFRESH As String = "G1DerivedCatalogRefresh"
Public Const SETUP_DOMAIN_H2_RIBBON_BUILD As String = "H2RibbonBuild"
Public Const SETUP_DOMAIN_H3_RIBBON_COMMIT As String = "H3RibbonCommit"
Public Const SETUP_DOMAIN_I1_NAVIGATION_GROUPS As String = "I1NavigationGroups"
Public Const SETUP_DOMAIN_I3_NAVIGATION_PLACEMENT As String = "I3NavigationPlacement"
Public Const SETUP_DOMAIN_J1_FINAL_STATE As String = "J1FinalState"
' INTENTIONAL ORDERED OPERATION NOTES:
' Appends the current top-level operation's presentation notes in occurrence order. Reset at operation start; Core owns the mutable Collection and callers add/read through Core helpers rather than receiving it
' directly.
Private mOperationNotes As Collection
' Top-level operation warning and detail-trace suppression state.
Private mOperationWarningCount As Long
Private mOperationDetailTraceSuppressionDepth As Long
' Live Setup phase and progress-identity state.
Private mSetupCancelled As Boolean
Private mSetupCurrentPartText As String
Private mSetupPreflightContextActive As Boolean
Private mSetupPhaseIdentity As String
Private mSetupPhaseLabel As String
Private mSetupPhasePartNumber As Long
Private mSetupPhasePartCount As Long
' Retained snapshot of the authoritative live Setup context at the instant of failure.
Private mFailedSetupContextCaptured As Boolean
Private mFailedSetupWasPreflight As Boolean
Private mFailedSetupPhaseIdentity As String
Private mFailedSetupPhaseLabel As String
Private mFailedSetupPartIdentity As String
Private mFailedSetupPhasePartNumber As Long
Private mFailedSetupPhasePartCount As Long
' Setup-run lifecycle state for the active top-level or standalone Setup operation; these flags span Setup phases and are Setup-run-owned state.
Private mSetupRunContextStarted As Boolean
Private mSetupNamingSnapshotStarted As Boolean
Private mSetupNamingSnapshotDirty As Boolean
Private mSetupNamingPolicyDirty As Boolean
' Read-only Module Guard Preflight naming evidence may be transferred into A1 only within the same trace run and HostDb.
Private mSetupPreflightNamingSnapshotReady As Boolean
Private mSetupPreflightNamingSnapshotTraceRunID As String
Private mSetupPreflightNamingSnapshotHostDbName As String
' Setup-run commit/progress state; lifetime and reset semantics match the active Setup operation.
Private mSetupMaterialCommitted As Boolean
Private mSetupCurrentStep As String
Private mFinalSetupJ1ProofToken As String
Private mFinalSetupJ1ModuleGuardStateToken As String
Private mFinalSetupJ1Detail As String
Private mSetupPartRollbackStartIndex As Long
' CANONICAL AUTHORITY: Owns the current Setup run's in-memory context/evidence buckets. Its lifetime is the active Setup run; Core initializes/tears it down, and named-bucket helpers are the supported mutation path.
Private mSetupRunContext As Object
' Current top-level system-command Ribbon-restart state. Persisted Ribbon reload state remains UI-owned.
Private mRibbonRestartRequiredCurrentRun As Boolean
' True only while System Macro 65's generated Access-macro RunCode call is still inside RunSystemCommand.
' Core owns this execution-context lifecycle because Access cannot remove the currently running macro object until its RunCode call has returned.
Private mMacro65RunCodeActive As Boolean
' Deferred Macro 65 deletion timer and completion state.
Private Const MACRO65_DEFERRED_DELETE_TIMER_INTERVAL_MS As Long = 250
Private Const MACRO65_DEFERRED_DELETE_MAX_ATTEMPTS As Long = 20
#If VBA7 Then
Private Declare PtrSafe Function CoreSetTimer Lib "user32" Alias "SetTimer" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Private Declare PtrSafe Function CoreKillTimer Lib "user32" Alias "KillTimer" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
Private mMacro65DeferredDeleteTimerID As LongPtr
#Else
Private Declare Function CoreSetTimer Lib "user32" Alias "SetTimer" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function CoreKillTimer Lib "user32" Alias "KillTimer" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private mMacro65DeferredDeleteTimerID As Long
#End If
Private mMacro65DeferredDeleteRequired As Boolean
Private mMacro65DeferredDeleteName As String
Private mMacro65DeferredDeleteCompletionKind As String
Private mMacro65DeferredDeleteShowMessage As Boolean
Private mMacro65DeferredDeleteCompletionMessage As String
Private mMacro65DeferredDeleteCompletionTitle As String
Private mMacro65DeferredDeleteInactiveTickObserved As Boolean
Private mMacro65DeferredDeleteAttemptCount As Long
Private mMacro65DeferredDeleteCompleteOwnedTrace As Boolean
Private mMacro65DeferredDeleteTraceSuccessMessage As String
Private mMacro65DeferredDeleteTraceFailureMessage As String
' Settings-readiness state. INTENTIONAL KEYED LOOKUP:
' Stores readiness keys already proved for the current database/settings context so repeated consumers do not repeat readiness work. ClearSettingsRuntimeStateForCurrentContext invalidates the set; callers do not
' receive the mutable Collection.
Private mSettingsReadyKeys As Collection
' Setup progress trace ownership and setup rollback state; trace state and notes are owned by
Private mProgressOwnsTraceRun As Boolean
Private mSetupRollbackActive As Boolean
' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION:
' This operation-owned rollback journal records pre-mutation recovery payloads in execution order. It is created/reset for one Setup operation, retained only while rollback/recovery remains possible, and is mutated
' only through rollback-tracking helpers.
Private mSetupRollbackEntries As Collection
Private mSetupRollbackOwnerTraceRunID As String
Private mSetupRollbackOperationName As String
Private mSetupRollbackResumeFirstPartPending As Boolean
' Shared semantic progress dispatch state.
Private Const SETUP_ACTIVITY_DISPATCH_INTERVAL_SECONDS As Double = 0.2
Private mSharedLatestStatusText As String
Private mSharedLastDispatchedStatusText As String
Private mSharedLastActivityDispatchTimer As Single
Private mSharedProgressDispatchActive As Boolean
Private mDisplayOnlyStatusRegistrationCount As Long
Private mDisplayStatusDispatchCount As Long
Private mDisplayStatusDuplicateSkipCount As Long
' Canonical runtime database-context state.
Private mHostDb As DAO.Database
' Setup progress and process timing state.
Private mSetupProgressActive As Boolean
Private mSetupProgressCurrentStep As Long
Private mSetupProgressMaxSteps As Long
Private mProgressOwnerTraceRunID As String
Private mProgressStatusPrefix As String
Private mProgressCompleteStatus As String
Private mProgressFailedStatus As String
Private mProgressScopeLabel As String
Private mLastFailedProgressPartPrefix As String
Private mLastFailedProgressText As String
Private mProgressChildDetailDepth As Long
Private mProgressChildScopeLabel As String
' INTENTIONAL ORDERED ITERATION:
' The child-scope and tracked-process Collections are LIFO stacks whose order is runtime state, not a projection of canonical data. They live only for the active progress/process nesting and are cleared by the
' existing progress/process teardown paths.
Private mProgressChildScopeStack As Collection
Private mAccessProcessPriorityLoggedTraceRunID As String
Private mCurrentProcessName As String
Private mCurrentProcessStartedOn As Date
Private mCurrentProcessStartTimer As Single
Private mCurrentProcessActive As Boolean
Private mCurrentProcessStack As Collection
Private mLastEndedProcessName As String
Private mLastEndedProcessElapsedText As String
Private mLastEndedProcessSummaryAvailable As Boolean
' Standalone Setup-parts operation state.
Private mStandaloneSetupPartsOperationActive As Boolean
' Phase C system-macro operation-scoped planning and physical-state proof.
' INTENTIONAL KEYED LOOKUP + PLAN/PROOF SEPARATION:
' mPartCExpectedMacros is the keyed view of the authoritative MC_Sys/Help_Docs source captured for C1. mPartCPhysicalEvidence is the one authoritative operation-scoped physical payload evidence snapshot: C1
' constructs it, C2/C3 consume it, and only a successfully proven Phase C mutation changes an existing observation afterward.
' Physical evidence item contract: Array(physical name, C1/current proven existence Boolean, proven RunCode, expected-name Boolean, C1 stale-owned Boolean). C1 records expected macros plus generated-looking
' unexpected candidates it physically inspects; unrelated macro payloads are never copied into another equivalent snapshot.
' C2 build-plan item contract: item(0)=MacroTitle; item(1)=authoritative expected RunCode; item(2)=reason ("Missing" or "Stale"). C1-observed existence/RunCode remain only in mPartCPhysicalEvidence and C2 validates
' planned mutation targets against that authoritative evidence.
' C3 validates the C1 delete-name set without a whole-domain payload reread, rereads each destructive target immediately before deletion, and derives final exact state from C1 evidence plus C2/C3 physically proven
' mutations and a cheap current name inventory.
Private mPartCExpectedMacros As Object
Private mPartCPhysicalEvidence As Object
Private mPartCBuildPlan As Collection
Private mPartCDeletePlan As Collection
Private mPartCExpectedStateToken As String
Private mPartCExpectedCount As Long
Private mPartCMatchingCount As Long
Private mPartCMissingCount As Long
Private mPartCReplacementCount As Long
Private mPartCStaleOwnedCount As Long
Private mPartCPlanReady As Boolean
' Operation-scoped authority boundary captured by C1; C2/C3 consume the same Help Document scope and the C1 clear path invalidates it with the plan/evidence.
Private mPartCHelpDocumentSynchronizationScope As HelpDocumentSynchronizationScope
' Build/Repair Help Document Macros may prove the complete System Macro and Help Document Macro state while authorizing mutations only for Help-document macros. This keeps C3's global proof exact without allowing
' operational Build/Repair Help Document Macros to rewrite unrelated System Macros.
Private mPartCHelpDocumentMutationsOnly As Boolean
Private mPartCLastPhysicalStateToken As String
Private mPartCPlanSourceToken As String
' Lightweight aggregate Phase C performance evidence; reset with the operation-scoped plan/evidence.
Private mPartCPayloadReadCount As Long
Private mPartCRollbackBackupCount As Long
Private mPartCTargetedStaleRereadCount As Long
Private mPartCCreatedCount As Long
Private mPartCReplacedCount As Long
Private mPartCDeletedCount As Long
Private mPartCWholeSetSecondPayloadPassOccurred As Boolean

' Operational System Command dispatcher re-entry and deferred result-detail state.
Private mOperationalSystemCommandActive As Boolean
Private mBuildRepairHelpDocumentMacrosMacroResultDetail As String

' ===============================================================================================================================================================================
' PROCEDURES
' ===============================================================================================================================================================================

' ===============================================================================================================================================================================
' PUBLIC SETUP ENTRY / DATABASE / PROCESS CONTEXT PROCEDURES
' ===============================================================================================================================================================================

' Purpose: Provides the Access macro/RunCode Function wrapper for RunFullSetup, recording the physical pre-Setup entry marker before any Setup processing.
'   Invokes the public full-Setup entry point with messages enabled and returns True after that call completes.
' Public role: Macro-callable public entry point. The Function form and Public visibility allow Access macros/RunCode expressions to invoke the underlying ADC operation.
' Parameters: None.
' Returns: Boolean: Returns True after RunFullSetup completes normally. RunFullSetup failures are not converted to False; they retain the underlying public Setup entry point's error behavior.
Public Function Run_RunFullSetup() As Boolean
    'Intended use: Macro-callable RunCode entry point for RunFullSetup.
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim savedHelpFile As String
    Dim savedHelpContext As Long
    LogRunFullSetupEntryPhysicalMarker
    On Error GoTo ErrHandler
    RunFullSetup True
    ClearRunFullSetupEntryPhysicalMarker
    Run_RunFullSetup = True
    Exit Function
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    savedHelpFile = Err.HelpFile
    savedHelpContext = Err.HelpContext
    ClearRunFullSetupEntryPhysicalMarker
    Err.Raise savedErrNumber, savedErrSource, savedErrDescription, savedHelpFile, savedHelpContext
End Function

' Purpose: Resolves ADC DB from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters: None.
' Returns: DAO.Database: Returns the resolved object reference. `Nothing` remains the no-object result where the existing routine permits an unresolved or unavailable object.
Public Function ADCDb() As DAO.Database
    ' Wrapper purpose: Core-owned database-context facade that returns the Access Development Console code-owning DAO.Database for runtime callers.
    ' Why retained: Retained to centralize CodeDb selection in Core so other modules do not independently choose or retain the code-owning database context.
    'Returns the Access Development Console code-owning database for global runtime callers.
    Set ADCDb = codeDb
End Function

' Purpose: Returns the retained active HostDb so child DAO objects do not outlive a temporary CurrentDb object.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: DAO.Database: Returns the DAO.Database object representing the resolved HostDb role for the current ADC configuration.
Public Function GetHostDb() As DAO.Database
    'Returns the retained active HostDb so child DAO objects do not outlive a temporary CurrentDb object.
    On Error GoTo RefreshDatabase
    Dim databaseName As String
    If mHostDb Is Nothing Then GoTo RefreshDatabase
    databaseName = mHostDb.Name
    Set GetHostDb = mHostDb
    Exit Function
RefreshDatabase:
    Err.Clear
    Set mHostDb = currentDb
    Set GetHostDb = mHostDb
End Function

' Purpose: Returns the canonical retained database for a Config owner role. Existing invalid-role behavior remains HostDb unless the caller validates first.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   ConfigOwnerDatabaseRole (Input; String): Identifies the database role whose configuration is authoritative for this operation, such as the CodeDb or HostDb role defined by ADC configuration.
' Returns: DAO.Database: Returns the DAO.Database object resolved for the requested ADC database role (such as CodeDb or HostDb); invalid/unresolvable roles follow the function’s existing error path.
Public Function DatabaseForRole(ByVal ConfigOwnerDatabaseRole As String) As DAO.Database
    'Returns the canonical retained database for a Config owner role. Existing invalid-role behavior remains HostDb unless the caller validates first.
    Select Case LCase$(NormalizedConfigOwnerDatabaseRole(ConfigOwnerDatabaseRole))
        Case LCase$(CONFIG_OWNER_DATABASE_ROLE_CODE_DB)
            Set DatabaseForRole = codeDb
        Case Else
            Set DatabaseForRole = GetHostDb()
    End Select
End Function

' Purpose: Resolves and returns database path for role from the supplied inputs and the authoritative ADC definitions/state used by this module.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   ConfigOwnerDatabaseRole (Input; String): Identifies the database role whose configuration is authoritative for this operation, such as the CodeDb or HostDb role defined by ADC configuration.
' Returns: String: Returns the normalized/resolved database or filesystem path for database path for role; vbNullString indicates that the requested role/source has no configured path where that is an allowed
'   state.
Public Function DatabasePathForRole(ByVal ConfigOwnerDatabaseRole As String) As String
    On Error GoTo ErrHandler
    DatabasePathForRole = DatabaseForRole(ConfigOwnerDatabaseRole).Name
    Exit Function
ErrHandler:
    DatabasePathForRole = vbNullString
End Function

' Purpose: Resolves the loaded VBProject from the canonical database role/path authority.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   ConfigOwnerDatabaseRole (Input; String): Identifies the database role whose configuration is authoritative for this operation, such as the CodeDb or HostDb role defined by ADC configuration.
' Returns: Object: Returns the VBE project associated with the requested configuration-owner database role so source/procedure inspection occurs in the authoritative project.
Public Function VBProjectForConfigOwnerDatabaseRole(ByVal ConfigOwnerDatabaseRole As String) As Object
    'Resolves the loaded VBProject from the canonical database role/path authority.
    Dim normalizedRole As String
    Dim databasePath As String
    normalizedRole = NormalizedConfigOwnerDatabaseRole(ConfigOwnerDatabaseRole)
    If Len(normalizedRole) = 0 Then Exit Function
    databasePath = DatabasePathForRole(normalizedRole)
    If Len(Trim$(databasePath)) = 0 Then Exit Function
    Set VBProjectForConfigOwnerDatabaseRole = VBProjectForDatabasePath(databasePath)
End Function

Private Function VBProjectForDatabasePath(ByVal databasePath As String) As Object
    On Error GoTo ErrHandler
    Dim vbProject As Object
    For Each vbProject In Application.vbe.VBProjects
        PumpSetupResponsiveness
        If VbProjectLooksLikeDatabase(vbProject, databasePath) Then
            Set VBProjectForDatabasePath = vbProject
            Exit Function
        End If
    Next vbProject
    If VbProjectLooksLikeDatabase(Application.vbe.ActiveVBProject, databasePath) Then
        Set VBProjectForDatabasePath = Application.vbe.ActiveVBProject
    End If
    Exit Function
ErrHandler:
    Set VBProjectForDatabasePath = Nothing
End Function

Private Function VbProjectLooksLikeDatabase(ByVal vbProject As Object, ByVal databasePath As String) As Boolean
    On Error GoTo ErrHandler
    Dim fileNameText As String
    If vbProject Is Nothing Then Exit Function
    fileNameText = vbProject.fileName
    If Len(Trim$(fileNameText)) = 0 Then Exit Function
    VbProjectLooksLikeDatabase = (StrComp(NormalizedDatabasePath(fileNameText), NormalizedDatabasePath(databasePath), vbTextCompare) = 0)
    Exit Function
ErrHandler:
    VbProjectLooksLikeDatabase = False
End Function

' Purpose: Compares the configured CodeDb and HostDb identities and reports whether both roles resolve to the same Access database.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: Boolean: True when begin operation detail trace suppression completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or not-applicable
'   path.
Public Function CodeDbIsHostDb() As Boolean
    'Intended use: Returns whether the canonical CodeDb and retained active HostDb resolve to the same database.
    On Error GoTo ErrHandler
    CodeDbIsHostDb = (StrComp(NormalizedDatabasePath(ADCDb().Name), NormalizedDatabasePath(GetHostDb().Name), vbTextCompare) = 0)
    Exit Function
ErrHandler:
    CodeDbIsHostDb = False
End Function

' Purpose: Returns whether the supplied database is the active Access CurrentProject database.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   db (Input; DAO.Database): DAO.Database instance whose schema, data, queries, or properties the routine reads or modifies.
' Returns: Boolean: True when the supplied database is the active Access CurrentProject database; False otherwise.
Public Function DatabaseIsCurrentProject(ByVal db As DAO.Database) As Boolean
    'Returns whether the supplied database is the active Access CurrentProject database.
    On Error GoTo ErrHandler
    If db Is Nothing Then Exit Function
    DatabaseIsCurrentProject = (StrComp(NormalizedDatabasePath(db.Name), NormalizedDatabasePath(Application.CurrentProject.FullName), vbTextCompare) = 0)
    Exit Function
ErrHandler:
    DatabaseIsCurrentProject = False
End Function


' Purpose: Performs the public core reset operation notes operation using the routine's existing ADC workflow and side effects.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub CoreResetOperationNotes()
    Set mOperationNotes = New Collection
    mOperationWarningCount = 0
End Sub

Private Sub PrepareOperationNotes()
    If mOperationNotes Is Nothing Then Set mOperationNotes = New Collection
End Sub

' Purpose: Performs the public core add operation note operation using the routine's existing ADC workflow and side effects.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   noteText (Input; String): Note text text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub CoreAddOperationNote(ByVal noteText As String)
    'Single operation-note/warning authority shared by Core, TABLES, CATALOG, and runtime workflows.
    Dim traceLevel As String
    noteText = Trim$(noteText)
    PrepareOperationNotes
    If Len(noteText) > 0 Then
        mOperationNotes.Add " - " & noteText
        traceLevel = IIf(CommonTextIndicatesWarning(noteText), TRACE_LEVEL_WARNING, TRACE_LEVEL_DETAIL)
        If StrComp(traceLevel, TRACE_LEVEL_WARNING, vbTextCompare) = 0 Then
            mOperationWarningCount = mOperationWarningCount + 1
            LogTraceMessage traceLevel, "Operation", "CoreAddOperationNote", noteText, vbNullString, vbNullString, 0, vbNullString, vbNullString, vbNullString, vbNullString, "CodeDb", "OperationNote"
        ElseIf mOperationDetailTraceSuppressionDepth = 0 Then
            LogTraceMessage traceLevel, "Operation", "CoreAddOperationNote", noteText, vbNullString, vbNullString, 0, vbNullString, vbNullString, vbNullString, vbNullString, "CodeDb", "OperationNote"
        End If
    End If
End Sub

' Purpose: Reports whether the current tracked core operation has accumulated one or more warning conditions.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: Boolean: True when begin tracked process completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or not-applicable path.
Public Function CoreOperationHasWarnings() As Boolean
    CoreOperationHasWarnings = (mOperationWarningCount > 0)
End Function

' Purpose: Initializes operation detail trace suppression, coordinating `End` as the implementation requires.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub BeginOperationDetailTraceSuppression()
    mOperationDetailTraceSuppressionDepth = mOperationDetailTraceSuppressionDepth + 1
End Sub

' Purpose: Closes operation detail trace suppression, coordinating `End` as the implementation requires.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub EndOperationDetailTraceSuppression()
    If mOperationDetailTraceSuppressionDepth > 0 Then mOperationDetailTraceSuppressionDepth = mOperationDetailTraceSuppressionDepth - 1
End Sub

' Purpose: Resolves core operation summary for message from the supplied/current ADC state, using `SharedCollection_Join` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function CoreOperationSummaryForMessage() As String
    PrepareOperationNotes
    If mOperationNotes.count = 0 Then Exit Function
    CoreOperationSummaryForMessage = vbCrLf & vbCrLf & "Table creation/maintenance operations:" & vbCrLf & SharedCollection_Join(mOperationNotes, vbCrLf)
End Function

' Purpose: Processes run summary for message for source procedure.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   sourceProcedure (Input; String): Procedure name recorded as the source of trace, error, status, or diagnostic information emitted by this call.
' Returns: String: Returns formatted process run summary for message text intended for tracing, diagnostics, or user-facing reporting; vbNullString is returned when there is nothing applicable to report.
Public Function ProcessRunSummaryForMessage(ByVal sourceProcedure As String) As String
    Dim processText As String
    Dim elapsedText As String
    Dim elapsedLabel As String
    elapsedLabel = "Elapsed time"
    If mLastEndedProcessSummaryAvailable Then
        processText = mLastEndedProcessName
        elapsedText = mLastEndedProcessElapsedText
    ElseIf mCurrentProcessActive Then
        processText = TrackedProcessName(sourceProcedure)
        elapsedText = FormattedElapsedSeconds(ElapsedSecondsFromTimer(mCurrentProcessStartTimer))
        elapsedLabel = "Top-level process elapsed"
    Else
        processText = TrackedProcessName(sourceProcedure)
        elapsedText = "not tracked"
    End If
    If Len(Trim$(processText)) = 0 Then processText = TrackedProcessName(sourceProcedure)
    ProcessRunSummaryForMessage = vbCrLf & vbCrLf & "Process run:" & vbCrLf & "Process: " & processText & vbCrLf & elapsedLabel & ": " & elapsedText & FailedProgressSummaryForMessage()
End Function

Private Sub ClearLastEndedProcessSummary()
    mLastEndedProcessName = vbNullString
    mLastEndedProcessElapsedText = vbNullString
    mLastEndedProcessSummaryAvailable = False
End Sub

' Purpose: Performs the public capture tracked process summary operation using the routine's existing ADC workflow and side effects.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub CaptureTrackedProcessSummary()
    If mCurrentProcessActive Then
        mLastEndedProcessName = TrackedProcessName(vbNullString)
        mLastEndedProcessElapsedText = FormattedElapsedSeconds(ElapsedSecondsFromTimer(mCurrentProcessStartTimer))
        mLastEndedProcessSummaryAvailable = True
    End If
End Sub

' Purpose: Initializes tracked process, coordinating `End` as the implementation requires.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   processName (Input; String): Logical process name used to identify the tracked progress/trace run.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub BeginTrackedProcess(ByVal processName As String)
    ClearLastEndedProcessSummary
    processName = Trim$(processName)
    If Len(processName) = 0 Then processName = DEFAULT_UTILITY_LONG_NAME & " process"
    If mCurrentProcessStack Is Nothing Then Set mCurrentProcessStack = New Collection
    mCurrentProcessStack.Add Array(mCurrentProcessActive, mCurrentProcessName, mCurrentProcessStartedOn, mCurrentProcessStartTimer)
    mCurrentProcessName = processName
    mCurrentProcessStartedOn = Now
    mCurrentProcessStartTimer = Timer
    mCurrentProcessActive = True
End Sub

' Purpose: Closes tracked process, coordinating `mCurrentProcessStack` and `End` as the implementation requires.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub EndTrackedProcess()
    Dim priorState As Variant
    If Not mCurrentProcessStack Is Nothing Then
        If mCurrentProcessStack.count > 0 Then
            priorState = mCurrentProcessStack(mCurrentProcessStack.count)
            mCurrentProcessStack.Remove mCurrentProcessStack.count
            mCurrentProcessActive = CBool(priorState(0))
            mCurrentProcessName = CStr(priorState(1))
            mCurrentProcessStartedOn = CDate(priorState(2))
            mCurrentProcessStartTimer = CSng(priorState(3))
            Exit Sub
        End If
    End If
    mCurrentProcessActive = False
    mCurrentProcessName = vbNullString
    mCurrentProcessStartedOn = 0
    mCurrentProcessStartTimer = 0
End Sub

Private Function TrackedProcessName(ByVal fallbackProcedure As String) As String
    If mCurrentProcessActive And Len(mCurrentProcessName) > 0 Then
        TrackedProcessName = mCurrentProcessName
    Else
        TrackedProcessName = Trim$(fallbackProcedure)
    End If
    If Len(TrackedProcessName) = 0 Then TrackedProcessName = DEFAULT_UTILITY_LONG_NAME & " process"
End Function

' ===============================================================================================================================================================================
' SHARED PROGRESS / SETUP UI INTEGRATION PROCEDURES
' ===============================================================================================================================================================================

Private Sub ClearFailedProgressContext()
    mLastFailedProgressPartPrefix = vbNullString
    mLastFailedProgressText = vbNullString
    mFailedSetupContextCaptured = False
    mFailedSetupWasPreflight = False
    mFailedSetupPhaseIdentity = vbNullString
    mFailedSetupPhaseLabel = vbNullString
    mFailedSetupPartIdentity = vbNullString
    mFailedSetupPhasePartNumber = 0
    mFailedSetupPhasePartCount = 0
End Sub

Private Sub CaptureFailedSetupPhaseContext()
    Dim partIdentity As String
    If mSetupPreflightContextActive Then
        mFailedSetupContextCaptured = True
        mFailedSetupWasPreflight = True
        mFailedSetupPhaseIdentity = vbNullString
        mFailedSetupPhaseLabel = vbNullString
        mFailedSetupPartIdentity = vbNullString
        mFailedSetupPhasePartNumber = 0
        mFailedSetupPhasePartCount = 0
        Exit Sub
    End If
    partIdentity = mSetupPhaseIdentity
    If Len(partIdentity) = 0 Then Exit Sub
    mFailedSetupContextCaptured = True
    mFailedSetupWasPreflight = False
    mFailedSetupPartIdentity = partIdentity
    mFailedSetupPhaseIdentity = Left$(partIdentity, 1)
    mFailedSetupPhaseLabel = mSetupPhaseLabel
    mFailedSetupPhasePartNumber = mSetupPhasePartNumber
    mFailedSetupPhasePartCount = mSetupPhasePartCount
End Sub

' Purpose: Resolves failed setup phase context for message from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function FailedSetupPhaseContextForMessage() As String
    Dim resultText As String
    If mFailedSetupContextCaptured = False Then
        FailedSetupPhaseContextForMessage = "Setup phase: Not yet established"
        Exit Function
    End If
    If mFailedSetupWasPreflight Then
        FailedSetupPhaseContextForMessage = "Setup phase: Preflight"
        Exit Function
    End If
    resultText = "Setup phase: " & mFailedSetupPhaseIdentity
    If Len(mFailedSetupPhaseLabel) > 0 Then resultText = resultText & " - " & mFailedSetupPhaseLabel
    If Len(mFailedSetupPartIdentity) > 0 Then
        resultText = resultText & vbCrLf & "Setup part: " & mFailedSetupPartIdentity
        If mFailedSetupPhasePartNumber > 0 And mFailedSetupPhasePartCount > 0 Then
            resultText = resultText & " (" & CStr(mFailedSetupPhasePartNumber) & " of " & CStr(mFailedSetupPhasePartCount) & ")"
        End If
    End If
    FailedSetupPhaseContextForMessage = resultText
End Function

Private Function FailureProgressContextForMessage() As String
    Dim partText As String
    partText = Trim$(mLastFailedProgressPartPrefix)
    If Right$(partText, 1) = ">" Then partText = Trim$(Left$(partText, Len(partText) - 1))
    If Len(partText) = 0 And InStr(1, mLastFailedProgressText, "Part ", vbTextCompare) > 0 Then
        partText = Trim$(mLastFailedProgressText)
        If InStr(1, partText, " > ", vbBinaryCompare) > 0 Then partText = Left$(partText, InStr(1, partText, " > ", vbBinaryCompare) - 1)
    End If
    If Len(partText) > 0 Then FailureProgressContextForMessage = vbCrLf & "Failure point: " & partText
End Function

Private Function FailedProgressSummaryForMessage() As String
    If Len(Trim$(mLastFailedProgressText)) > 0 Then FailedProgressSummaryForMessage = vbCrLf & "Failed progress: " & mLastFailedProgressText
End Function

Private Function ActiveProgressBelongsToCurrentTraceRun() As Boolean
    Dim currentTraceRunID As String
    If mSetupProgressActive = False Then Exit Function
    If TraceRunActive() = False Then Exit Function
    currentTraceRunID = GetActiveTraceRunID()
    If Len(currentTraceRunID) = 0 Or Len(mProgressOwnerTraceRunID) = 0 Then Exit Function
    ActiveProgressBelongsToCurrentTraceRun = (StrComp(currentTraceRunID, mProgressOwnerTraceRunID, vbBinaryCompare) = 0)
End Function

' Purpose: Returns live Core progress ownership so child modules do not trust stale TempVar state after a completed or failed run.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: Boolean: True when the condition represented by CoreQualifyProgressStatusText is present under the routine's documented inputs/current state; False when that condition is absent.
Public Function CoreParentProgressActive() As Boolean
    'Returns live Core progress ownership so child modules do not trust stale TempVar state after a completed or failed run.
    On Error GoTo NotActive
    CoreParentProgressActive = ActiveProgressBelongsToCurrentTraceRun()
    Exit Function
NotActive:
    Err.Clear
End Function

Private Function CoreParentProgressScope() As String
    'Returns the live phase-aware top-level progress scope while the current trace run owns that progress state.
    If CoreParentProgressActive() = False Then Exit Function
    If Len(SetupPhaseProgressScopeText()) > 0 Then
        CoreParentProgressScope = SetupPhaseProgressScopeText()
    Else
        CoreParentProgressScope = mProgressScopeLabel
    End If
End Function

Private Function CoreParentProgressStep() As Long
    'Returns zero for phase-aware Setup so child modules use the already-complete phase scope instead of reconstructing flat Part X of X text.
    If CoreParentProgressActive() = False Then Exit Function
    If Len(SetupPhaseProgressScopeText()) > 0 Then Exit Function
    CoreParentProgressStep = mSetupProgressCurrentStep
End Function

Private Function CoreParentProgressMaxSteps() As Long
    'Returns zero for phase-aware Setup so child modules do not append noncurrent flat Part X of X text.
    If CoreParentProgressActive() = False Then Exit Function
    If Len(SetupPhaseProgressScopeText()) > 0 Then Exit Function
    CoreParentProgressMaxSteps = mSetupProgressMaxSteps
End Function

Private Function CoreParentProgressPrefixText() As String
    'Returns the canonical visible prefix for child progress nested beneath the live Core-owned parent progress.
    Dim parentScopeText As String
    Dim parentStepValue As Long
    Dim parentMaxValue As Long
    parentScopeText = CoreParentProgressScope()
    parentStepValue = CoreParentProgressStep()
    parentMaxValue = CoreParentProgressMaxSteps()
    If Len(parentScopeText) = 0 Then parentScopeText = "Setup"
    If parentStepValue > 0 And parentMaxValue > 0 Then
        CoreParentProgressPrefixText = parentScopeText & " Part " & CStr(parentStepValue) & " of " & CStr(parentMaxValue) & " > "
    Else
        CoreParentProgressPrefixText = parentScopeText & " > "
    End If
End Function

' Purpose: Resolves core parent nested progress text from the supplied/current ADC state, using `CoreParentProgressPrefixText` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   statusText (Input; String): Status/progress message supplied for the current operation; the routine forwards it to the applicable UI/progress and trace/reporting path.
'   childScopeText (Input Optional; String): Input child scope text for core parent nested progress text; is tested by If Len(childScopeText) > 0; is passed to `CoreParentProgressPrefixText`. Default: vbNullString.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function CoreParentNestedProgressText(ByVal statusText As String, Optional ByVal childScopeText As String = vbNullString) As String
    'Returns the canonical visible child-progress text nested beneath the live Core-owned parent progress.
    statusText = Trim$(statusText)
    childScopeText = Trim$(childScopeText)
    If Len(statusText) = 0 Then Exit Function
    If Len(childScopeText) > 0 Then
        CoreParentNestedProgressText = CoreParentProgressPrefixText() & childScopeText & " > " & statusText
    Else
        CoreParentNestedProgressText = CoreParentProgressPrefixText() & statusText
    End If
End Function

' Purpose: Resolves core qualify progress status text from the supplied/current ADC state, using `StatusTextHasPartOfContext` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   statusText (Input; String): Status/progress message supplied for the current operation; the routine forwards it to the applicable UI/progress and trace/reporting path.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function CoreQualifyProgressStatusText(ByVal statusText As String) As String
    Dim partNumber As Long
    Dim phasePrefix As String
    statusText = Trim$(statusText)
    If Len(statusText) = 0 Then Exit Function
    If StatusTextHasPartOfContext(statusText) Or StatusTextHasSetupPhaseContext(statusText) Then
        CoreQualifyProgressStatusText = statusText
        Exit Function
    End If
    phasePrefix = SetupPhaseProgressPrefixText()
    If Len(phasePrefix) > 0 Then
        CoreQualifyProgressStatusText = phasePrefix & statusText
    ElseIf mSetupProgressMaxSteps > 0 And Len(mProgressScopeLabel) > 0 Then
        partNumber = mSetupProgressCurrentStep
        If partNumber < 1 Then partNumber = 1
        If partNumber > mSetupProgressMaxSteps Then partNumber = mSetupProgressMaxSteps
        CoreQualifyProgressStatusText = ProgressPartPrefixText() & "Part " & CStr(partNumber) & " of " & CStr(mSetupProgressMaxSteps) & " > " & statusText
    ElseIf Len(mLastFailedProgressPartPrefix) > 0 Then
        CoreQualifyProgressStatusText = mLastFailedProgressPartPrefix & statusText
    Else
        CoreQualifyProgressStatusText = "Operation Part 1 of 1 > " & statusText
    End If
End Function

' Purpose: Resolves core visible progress text from the supplied/current ADC state, using `CoreParentProgressActive` and `CoreParentNestedProgressText` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   statusText (Input; String): Status/progress message supplied for the current operation; the routine forwards it to the applicable UI/progress and trace/reporting path.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function CoreVisibleProgressText(ByVal statusText As String) As String
    'Returns worker status nested beneath an active parent progress scope without introducing presentation work.
    statusText = Trim$(statusText)
    If Len(statusText) = 0 Then Exit Function
    If CoreParentProgressActive() Then
        CoreVisibleProgressText = CoreParentNestedProgressText(statusText)
    Else
        CoreVisibleProgressText = statusText
    End If
End Function

' Purpose: Performs the public core reset shared progress diagnostics operation using the routine's existing ADC workflow and side effects.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub CoreResetSharedProgressDiagnostics()
    'Resets Core-owned semantic progress dispatch state at a new trace/run boundary.
    mSharedLatestStatusText = vbNullString
    mSharedLastDispatchedStatusText = vbNullString
    mSharedLastActivityDispatchTimer = 0
    mSharedProgressDispatchActive = False
    mDisplayOnlyStatusRegistrationCount = 0
    mDisplayStatusDispatchCount = 0
    mDisplayStatusDuplicateSkipCount = 0
End Sub

' Purpose: Resolves core progress performance diagnostics summary from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function CoreProgressPerformanceDiagnosticsSummary() As String
    CoreProgressPerformanceDiagnosticsSummary = _
        "Display-only activity registrations: " & CStr(mDisplayOnlyStatusRegistrationCount) & vbCrLf & _
        "Actual shared display dispatches: " & CStr(mDisplayStatusDispatchCount) & vbCrLf & _
        "Duplicate display dispatches skipped: " & CStr(mDisplayStatusDuplicateSkipCount)
End Function

Private Sub LogSharedProgressFailure(ByVal componentName As String, ByVal sourceProcedure As String, ByVal contextName As String, ByVal progressAction As String, ByVal statusText As String, ByVal errNumber As Long, ByVal errSource As String, ByVal errDescription As String)
    On Error Resume Next
    Dim componentPrefix As String
    componentName = Trim$(componentName)
    If Len(componentName) > 0 Then componentPrefix = componentName & " "
    Debug.Print DEFAULT_UTILITY_LONG_NAME & " " & componentPrefix & "progress/status display failed in " & sourceProcedure & " during " & progressAction & "; operation continues. Error " & CStr(errNumber) & ": " & errDescription
    LogTraceMessage TRACE_LEVEL_WARNING, "Progress", sourceProcedure, componentPrefix & "progress/status display failed; operation continues.", "Action: " & progressAction & vbCrLf & "Status text: " & statusText & vbCrLf & "Err.Number: " & CStr(errNumber) & vbCrLf & "Err.Source: " & errSource & vbCrLf & "Err.Description: " & errDescription, "Progress " & progressAction, errNumber, errSource, errDescription, vbNullString, "StatusBar", "CodeDb", contextName
End Sub

' Purpose: Resolves shared latest progress status text from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function SharedLatestProgressStatusText() As String
    'Allows the independent setup heartbeat to recover the latest registered activity even when UI dispatch was coalesced.
    SharedLatestProgressStatusText = mSharedLatestStatusText
End Function

Private Sub DispatchSharedProgressActivity(ByVal progressActive As Boolean, ByVal statusText As String, ByVal statusPrefix As String, ByVal sourceProcedure As String, ByVal contextName As String, ByVal componentName As String, Optional ByVal forceDisplay As Boolean = False)
    On Error GoTo DisplayFailure
    Dim currentTimer As Single
    Dim elapsedSeconds As Double
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim statusBarText As String
    mDisplayOnlyStatusRegistrationCount = mDisplayOnlyStatusRegistrationCount + 1
    mSharedLatestStatusText = statusText
    currentTimer = Timer
    If forceDisplay = False And mSharedLastActivityDispatchTimer > 0 Then
        elapsedSeconds = ElapsedSecondsFromTimer(mSharedLastActivityDispatchTimer)
        If elapsedSeconds < SETUP_ACTIVITY_DISPATCH_INTERVAL_SECONDS Then Exit Sub
    End If
    mSharedLastActivityDispatchTimer = currentTimer
    If forceDisplay = False Then
        If StrComp(mSharedLastDispatchedStatusText, statusText, vbBinaryCompare) = 0 Then
            mDisplayStatusDuplicateSkipCount = mDisplayStatusDuplicateSkipCount + 1
            SetupControlPump False
            Exit Sub
        End If
    End If
    If mSharedProgressDispatchActive Then Exit Sub
    mSharedProgressDispatchActive = True
    'Keep shared progress semantic through dispatch; apply hierarchy expansion only to the status-bar copy.
    statusBarText = StatusBarProgressText(statusText)
    statusPrefix = StatusBarHierarchyText(statusPrefix)
    If mSetupControlActive Then statusBarText = StatusBarTextWithCurrentSetupElapsed(statusBarText)
    If UIRenderSharedProgressActivity(progressActive, statusText, statusBarText, statusPrefix, forceDisplay, errNumber, errSource, errDescription) = False Then
        LogSharedProgressFailure componentName, sourceProcedure, contextName, "display activity", statusText, errNumber, errSource, errDescription
    End If
    mSharedLastDispatchedStatusText = statusText
    mDisplayStatusDispatchCount = mDisplayStatusDispatchCount + 1
SafeExit:
    mSharedProgressDispatchActive = False
    Exit Sub
DisplayFailure:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    LogSharedProgressFailure componentName, sourceProcedure, contextName, "display-only activity dispatch", statusText, savedErrNumber, savedErrSource, savedErrDescription
    On Error GoTo 0
    GoTo SafeExit
End Sub

' Purpose: Attempts to perform shared progress begin for progress max, caption text.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   progressActive (Output; Boolean): Set True when the procedure determines that progress active applies; set False when that condition is not met.
'   progressStep (Output; Long): Receives the resolved progress step produced from this procedure's validated inputs and current ADC state.
'   progressMax (Input/output; Long): Supplies the requested maximum progress steps; values below 1 are normalized to 1 and the normalized maximum is returned to the caller.
'   captionText (Input; String): Caption text text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   statusPrefix (Input; String): Prefix text prepended to or matched against the status content handled by this routine.
'   sourceProcedure (Input; String): Procedure name recorded as the source of trace, error, status, or diagnostic information emitted by this call.
'   contextName (Input; String): Context label recorded with tracing/progress/errors so the caller can distinguish the logical operation producing the message.
'   componentName (Input Optional; String): VBE component/module name used to limit or label source inspection and diagnostics. Default: vbNullString.
' Returns: None (Sub). Results are returned through the documented ByRef parameters and/or the routine's database, UI, or module-state side effects.
Public Sub TrySharedProgressBegin(ByRef progressActive As Boolean, ByRef progressStep As Long, ByRef progressMax As Long, ByVal captionText As String, ByVal statusPrefix As String, ByVal sourceProcedure As String, ByVal contextName As String, Optional ByVal componentName As String = vbNullString)
    'Begins status-only progress and forces one initial display without creating a persistent trace event.
    If progressMax < 1 Then progressMax = 1
    progressStep = 0
    progressActive = True
    DispatchSharedProgressActivity progressActive, captionText, statusPrefix, sourceProcedure, contextName, componentName, True
End Sub

' Purpose: Attempts to perform shared progress update for progress active, progress max.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   progressActive (Input; Boolean): Indicates whether the caller already owns an active progress scope; True lets this routine update/reuse it, while False lets the routine avoid parent-progress assumptions.
'   progressStep (Input/output; Long): Supplies the caller's current progress step and, when progress is active, receives the next step capped at progressMax.
'   progressMax (Input; Long): Upper bound limiting the progress max quantity scanned, displayed, or processed by this call.
'   statusText (Input; String): Status/progress message supplied for the current operation; the routine forwards it to the applicable UI/progress and trace/reporting path.
'   statusPrefix (Input; String): Prefix text prepended to or matched against the status content handled by this routine.
'   sourceProcedure (Input; String): Procedure name recorded as the source of trace, error, status, or diagnostic information emitted by this call.
'   contextName (Input; String): Context label recorded with tracing/progress/errors so the caller can distinguish the logical operation producing the message.
'   componentName (Input Optional; String): VBE component/module name used to limit or label source inspection and diagnostics. Default: vbNullString.
' Returns: None (Sub). Results are returned through the documented ByRef parameters and/or the routine's database, UI, or module-state side effects.
Public Sub TrySharedProgressUpdate(ByRef progressActive As Boolean, ByRef progressStep As Long, ByRef progressMax As Long, ByVal statusText As String, ByVal statusPrefix As String, ByVal sourceProcedure As String, ByVal contextName As String, Optional ByVal componentName As String = vbNullString)
    'Advances logical progress and forces the new Part X of X boundary to the live display.
    If progressActive Then
        progressStep = progressStep + 1
        If progressStep > progressMax Then progressStep = progressMax
    End If
    DispatchSharedProgressActivity progressActive, statusText, statusPrefix, sourceProcedure, contextName, componentName, True
End Sub

Private Sub TrySharedProgressPulse(ByVal progressActive As Boolean, ByVal statusText As String, ByVal statusPrefix As String, ByVal sourceProcedure As String, ByVal contextName As String, Optional ByVal componentName As String = vbNullString)
    ' Wrapper purpose: Non-forced shared-progress facade that registers detailed activity and delegates UI dispatch with forceDisplay fixed to False.
    ' Why retained: Retained to keep the general dispatcher private and give Core callers the specific throttled pulse contract without repeating the fixed argument.
    'Registers the latest detailed activity on every call and throttles actual UI work.
    DispatchSharedProgressActivity progressActive, statusText, statusPrefix, sourceProcedure, contextName, componentName, False
End Sub

Private Sub TrySharedProgressComplete(ByRef progressActive As Boolean, ByRef progressStep As Long, ByRef progressMax As Long, ByVal statusText As String, ByVal clearStatus As Boolean, ByVal sourceProcedure As String, ByVal contextName As String, Optional ByVal componentName As String = vbNullString)
    'Forces the completion text to the live display, then closes the owned shared-progress state.
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    DispatchSharedProgressActivity progressActive, statusText, vbNullString, sourceProcedure, contextName, componentName, True
    If UIEndSharedProgress(progressActive, progressStep, progressMax, clearStatus, errNumber, errSource, errDescription) = False Then LogSharedProgressFailure componentName, sourceProcedure, contextName, "complete end", statusText, errNumber, errSource, errDescription
End Sub

' Purpose: Attempts to perform shared progress fail for progress active, status text.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   progressActive (Input/output; Boolean): Indicates whether the caller already owns an active progress scope; True lets this routine update/reuse it, while False lets the routine avoid parent-progress
'   assumptions.
'   progressStep (Output; Long): Receives the resolved progress step produced from this procedure's validated inputs and current ADC state.
'   progressMax (Output; Long): Receives the resolved progress max produced from this procedure's validated inputs and current ADC state.
'   statusText (Input; String): Status/progress message supplied for the current operation; the routine forwards it to the applicable UI/progress and trace/reporting path.
'   sourceProcedure (Input; String): Procedure name recorded as the source of trace, error, status, or diagnostic information emitted by this call.
'   contextName (Input; String): Context label recorded with tracing/progress/errors so the caller can distinguish the logical operation producing the message.
'   componentName (Input Optional; String): VBE component/module name used to limit or label source inspection and diagnostics. Default: vbNullString.
' Returns: None (Sub). Results are returned through the documented ByRef parameters and/or the routine's database, UI, or module-state side effects.
Public Sub TrySharedProgressFail(ByRef progressActive As Boolean, ByRef progressStep As Long, ByRef progressMax As Long, ByVal statusText As String, ByVal sourceProcedure As String, ByVal contextName As String, Optional ByVal componentName As String = vbNullString)
    'Forces failure text to the display while preserving the failure text in the status bar.
    DispatchSharedProgressActivity progressActive, statusText, vbNullString, sourceProcedure, contextName, componentName, True
    progressActive = False
    progressStep = 0
    progressMax = 0
End Sub

' Purpose: Ends status-only progress and clears the shared activity cache when requested.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   progressActive (Input/output; Boolean): Supplies whether shared progress is active and is reset False when the UI progress state is ended or normalized after a display failure.
'   progressStep (Input/output; Long): Supplies the caller's current progress step and is reset to 0 when the shared progress state is ended.
'   progressMax (Input/output; Long): Supplies the caller's current progress maximum and is reset to 0 when the shared progress state is ended.
'   clearStatus (Input; Boolean): When True, enables or requests status for this call; when False, the routine follows its normal non-forced/default path.
'   sourceProcedure (Input; String): Procedure name recorded as the source of trace, error, status, or diagnostic information emitted by this call.
'   contextName (Input; String): Context label recorded with tracing/progress/errors so the caller can distinguish the logical operation producing the message.
'   componentName (Input Optional; String): VBE component/module name used to limit or label source inspection and diagnostics. Default: vbNullString.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub TrySharedProgressEnd(ByRef progressActive As Boolean, ByRef progressStep As Long, ByRef progressMax As Long, ByVal clearStatus As Boolean, ByVal sourceProcedure As String, ByVal contextName As String, Optional ByVal componentName As String = vbNullString)
    'Ends status-only progress and clears the shared activity cache when requested.
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    If UIEndSharedProgress(progressActive, progressStep, progressMax, clearStatus, errNumber, errSource, errDescription) = False Then LogSharedProgressFailure componentName, sourceProcedure, contextName, "end", "ClearStatus=" & CStr(clearStatus), errNumber, errSource, errDescription
    If clearStatus Then
        mSharedLatestStatusText = vbNullString
        mSharedLastDispatchedStatusText = vbNullString
        mSharedLastActivityDispatchTimer = 0
    End If
    SetupControlPump True
End Sub

Private Sub ResetOrphanedProgressState(ByVal requestedProgressTitle As String)
    On Error Resume Next
    Dim priorScope As String
    Dim priorPart As String
    Dim priorOwnerTraceRunID As String
    Dim currentTraceRunID As String
    priorScope = mProgressScopeLabel
    priorPart = ProgressPartText("orphaned progress state", False)
    priorOwnerTraceRunID = mProgressOwnerTraceRunID
    currentTraceRunID = GetActiveOrLastTraceRunID()
    TrySharedProgressEnd mSetupProgressActive, mSetupProgressCurrentStep, mSetupProgressMaxSteps, True, "ResetOrphanedProgressState", "Core", "Core"
    mSetupProgressActive = False
    mSetupProgressCurrentStep = 0
    mSetupProgressMaxSteps = 0
    mProgressOwnsTraceRun = False
    mProgressOwnerTraceRunID = vbNullString
    mProgressStatusPrefix = vbNullString
    mProgressCompleteStatus = vbNullString
    mProgressFailedStatus = vbNullString
    mProgressScopeLabel = vbNullString
    ClearProgressChildDetailState
    ClearParentProgressState
    Set mCurrentProcessStack = Nothing
    mCurrentProcessActive = False
    mCurrentProcessName = vbNullString
    mCurrentProcessStartedOn = 0
    mCurrentProcessStartTimer = 0
    ClearLastEndedProcessSummary
    LogTraceMessage TRACE_LEVEL_WARNING, "Progress", "ResetOrphanedProgressState", "Cleared orphaned " & DEFAULT_UTILITY_LONG_NAME & " progress state before starting a new operation.", "Requested operation: " & requestedProgressTitle & vbCrLf & "Prior scope: " & priorScope & vbCrLf & "Prior progress: " & priorPart & vbCrLf & "Prior owner trace run ID: " & priorOwnerTraceRunID & vbCrLf & "Current trace run ID: " & currentTraceRunID, "Reset orphaned progress state", 0, vbNullString, vbNullString, vbNullString, "Progress", "CodeDb", "Core"
    If Err.Number <> 0 Then Debug.Print "ADCProgress: orphaned progress state was cleared before [" & requestedProgressTitle & "], but warning trace output failed: " & Err.Description
    On Error GoTo 0
End Sub

Private Sub BeginProgressInternalForSetup(ByVal progressTitle As String, ByVal maxSteps As Long, ByVal statusPrefix As String, ByVal completeStatus As String, ByVal failedStatus As String, ByVal progressScopeLabel As String)
    On Error GoTo ErrHandler
    Dim trackedProcessStarted As Boolean
    Dim initializationStep As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    ClearFailedProgressContext
    ClearProgressChildDetailState
    initializationStep = "Clear prior progress state"
    EndSetupProgress True
    initializationStep = "Start tracked process"
    BeginTrackedProcess progressTitle
    trackedProcessStarted = True
    mProgressOwnsTraceRun = False
    initializationStep = "Start or reuse Setup trace run"
    If TraceRunActive() = False Then
        StartTraceRunForSetup "CoreProgress"
        mProgressOwnsTraceRun = True
    End If
    initializationStep = "Record progress owner trace run"
    mProgressOwnerTraceRunID = GetActiveOrLastTraceRunID()
    If Len(mProgressOwnerTraceRunID) = 0 Then Err.Raise vbObjectError + 6328, "BeginProgressInternalForSetup", "An " & DEFAULT_UTILITY_LONG_NAME & " Setup progress owner trace run ID could not be established."
    initializationStep = "Initialize progress state"
    mSetupProgressMaxSteps = maxSteps
    mProgressStatusPrefix = statusPrefix
    mProgressCompleteStatus = completeStatus
    mProgressFailedStatus = failedStatus
    mProgressScopeLabel = CleanProgressScopeLabel(progressScopeLabel)
    initializationStep = "Open shared progress display"
    TrySharedProgressBegin mSetupProgressActive, mSetupProgressCurrentStep, mSetupProgressMaxSteps, progressTitle, ProgressStatusPrefixText(), "BeginProgressInternalForSetup", "Core", "Setup"
    initializationStep = "Publish parent progress state"
    PublishParentProgressState
    initializationStep = "Write progress-start trace"
    LogTraceMessage TRACE_LEVEL_INFO, "Progress", "BeginProgressInternalForSetup", progressTitle, "Scope: " & mProgressScopeLabel & vbCrLf & "Max steps: " & CStr(maxSteps), progressTitle, 0, vbNullString, vbNullString, vbNullString, vbNullString, "CodeDb", "Setup"
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    mLastFailedProgressText = "Progress initialization failed > " & progressTitle & " > " & initializationStep
    On Error Resume Next
    LogTraceMessage TRACE_LEVEL_ERROR, "Progress", "BeginProgressInternalForSetup", mLastFailedProgressText, "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, initializationStep, savedErrNumber, savedErrSource, savedErrDescription, vbNullString, vbNullString, "CodeDb", "Setup"
    Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss"); " | ERROR | BeginProgressInternalForSetup | "; mLastFailedProgressText; " | "; CStr(savedErrNumber); " | "; savedErrDescription
    ClearParentProgressState
    TrySharedProgressEnd mSetupProgressActive, mSetupProgressCurrentStep, mSetupProgressMaxSteps, True, "BeginProgressInternalForSetup", "Core", "Setup"
    If mProgressOwnsTraceRun Then CompleteTraceRun mLastFailedProgressText, False
    If trackedProcessStarted Then
        CaptureTrackedProcessSummary
        EndTrackedProcess
    End If
    mProgressOwnsTraceRun = False
    mProgressOwnerTraceRunID = vbNullString
    mProgressStatusPrefix = vbNullString
    mProgressCompleteStatus = vbNullString
    mProgressFailedStatus = vbNullString
    mProgressScopeLabel = vbNullString
    On Error GoTo 0
    Err.Raise savedErrNumber, "BeginProgressInternalForSetup", "Could not initialize " & DEFAULT_UTILITY_LONG_NAME & " Setup progress at step '" & initializationStep & "'." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

' Purpose: Performs the public core pulse progress operation using the routine's existing ADC workflow and side effects.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   statusText (Input; String): Status/progress message supplied for the current operation; the routine forwards it to the applicable UI/progress and trace/reporting path.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub CorePulseProgress(ByVal statusText As String)
    'Registers detailed display-only setup activity; the shared dispatcher throttles actual UI work.
    Dim partStatusText As String
    partStatusText = ProgressPartText(statusText, False)
    TrySharedProgressPulse mSetupProgressActive, partStatusText, ProgressStatusPrefixText(), "CorePulseProgress", "Core", "Core"
End Sub

' Purpose: Performs pulse progress every as the module's public orchestration step using the established ADC state and validation contract.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   currentCount (Input; Long): Supplies the current count used to bound, report, or validate the number of items processed by this operation.
'   intervalCount (Input; Long): Supplies the interval count used to bound, report, or validate the number of items processed by this operation.
'   statusText (Input; String): Status/progress message supplied for the current operation; the routine forwards it to the applicable UI/progress and trace/reporting path.
'   totalCount (Input Optional; Long): Supplies the total count used to bound, report, or validate the number of items processed by this operation. Default: 0.
'   sourceProcedure (Input Optional; String): Procedure name recorded as the source of trace, error, status, or diagnostic information emitted by this call. Default: "PulseProgressEvery".
'   contextName (Input Optional; String): Context label recorded with tracing/progress/errors so the caller can distinguish the logical operation producing the message. Default: "Core".
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub PulseProgressEvery(ByVal currentCount As Long, ByVal intervalCount As Long, ByVal statusText As String, Optional ByVal totalCount As Long = 0, Optional ByVal sourceProcedure As String = "PulseProgressEvery", Optional ByVal contextName As String = "Core")
    'Canonical Core pulse authority; optional total/source/context preserve existing TABLES call-site diagnostics.
    Dim shouldLogPulse As Boolean
    Dim effectiveIntervalCount As Long
    Dim partStatusText As String
    effectiveIntervalCount = intervalCount
    If effectiveIntervalCount <= 0 Then effectiveIntervalCount = DEFAULT_STATUS_PULSE_EVERY_N_ROWS
    partStatusText = ProgressPartText(statusText, False)
    TrySharedProgressPulse mSetupProgressActive, partStatusText, ProgressStatusPrefixText(), sourceProcedure, "Core", contextName
    shouldLogPulse = (currentCount = 1)
    If shouldLogPulse = False And effectiveIntervalCount > 0 Then shouldLogPulse = (currentCount Mod effectiveIntervalCount = 0)
    If shouldLogPulse = False And totalCount > 0 Then shouldLogPulse = (currentCount = totalCount)
    If shouldLogPulse Then LogTraceMessage TRACE_LEVEL_STATUS, "Progress", sourceProcedure, partStatusText, "Current count: " & CStr(currentCount) & "; interval: " & CStr(effectiveIntervalCount) & IIf(totalCount > 0, "; total: " & CStr(totalCount), vbNullString), partStatusText, 0, vbNullString, vbNullString, vbNullString, vbNullString, "CodeDb", contextName
End Sub

Private Function ProgressStatusPrefixText() As String
    If Len(mProgressStatusPrefix) = 0 Then
        ProgressStatusPrefixText = DEFAULT_OPERATION_PROGRESS_STATUS_PREFIX
    Else
        ProgressStatusPrefixText = mProgressStatusPrefix
    End If
End Function

Private Function ProgressScopeFromTitle(ByVal progressTitle As String, Optional ByVal fallbackScopeLabel As String = vbNullString) As String
    Dim TitleText As String
    Dim fallbackText As String
    TitleText = LCase$(Trim$(progressTitle))
    fallbackText = CleanProgressScopeLabel(fallbackScopeLabel)
    If Len(fallbackText) > 0 Then
        ProgressScopeFromTitle = fallbackText
    ElseIf TitleText = LCase$(DEFAULT_UTILITY_SHORT_NAME & ": setup") Then
        ProgressScopeFromTitle = "Setup"
    ElseIf TitleText = "access development console codedb repair/update" Then
        ProgressScopeFromTitle = "Repair"
    ElseIf TitleText = LCase$(DEFAULT_UTILITY_SHORT_NAME & ": setup - hostdb") Then
        ProgressScopeFromTitle = "HostDb Setup"
    ElseIf Len(Trim$(progressTitle)) > 0 Then
        ProgressScopeFromTitle = Trim$(progressTitle)
    Else
        ProgressScopeFromTitle = "Operation"
    End If
End Function

Private Sub PublishParentProgressState()
    On Error Resume Next
    SetProgressTempVar PROGRESS_TEMPVAR_ACTIVE, True
    SetProgressTempVar PROGRESS_TEMPVAR_SCOPE, mProgressScopeLabel
    SetProgressTempVar PROGRESS_TEMPVAR_STEP, mSetupProgressCurrentStep
    SetProgressTempVar PROGRESS_TEMPVAR_MAX, mSetupProgressMaxSteps
End Sub

Private Sub ClearParentProgressState()
    On Error Resume Next
    TempVars.Remove PROGRESS_TEMPVAR_ACTIVE
    TempVars.Remove PROGRESS_TEMPVAR_SCOPE
    TempVars.Remove PROGRESS_TEMPVAR_STEP
    TempVars.Remove PROGRESS_TEMPVAR_MAX
End Sub

Private Sub SetProgressTempVar(ByVal tempVarName As String, ByVal tempVarValue As Variant)
    On Error Resume Next
    TempVars.Remove tempVarName
    TempVars.Add tempVarName, tempVarValue
End Sub

Private Function ProgressPartPrefixText() As String
    Dim scopeText As String
    scopeText = mProgressScopeLabel
    If Len(scopeText) > 0 Then ProgressPartPrefixText = scopeText & " "
End Function

Private Function ProgressPartText(ByVal stepText As String, Optional ByVal advanceStep As Boolean = False) As String
    Dim partNumber As Long
    Dim phasePrefix As String
    stepText = ProgressChildDetailText(stepText)
    If Len(stepText) = 0 Then Exit Function
    phasePrefix = SetupPhaseProgressPrefixText()
    If Len(phasePrefix) > 0 Then
        ProgressPartText = phasePrefix & stepText
        Exit Function
    End If
    If mSetupProgressActive = False Or mSetupProgressMaxSteps <= 0 Then
        If Len(mLastFailedProgressPartPrefix) > 0 Then
            ProgressPartText = mLastFailedProgressPartPrefix & stepText
        Else
            ProgressPartText = stepText
        End If
        Exit Function
    End If
    partNumber = mSetupProgressCurrentStep
    If advanceStep Then partNumber = partNumber + 1
    If partNumber < 1 Then partNumber = 1
    If partNumber > mSetupProgressMaxSteps Then partNumber = mSetupProgressMaxSteps
    ProgressPartText = ProgressPartPrefixText() & "Part " & CStr(partNumber) & " of " & CStr(mSetupProgressMaxSteps) & " > " & stepText
End Function

Private Function ProgressCompleteStatusText() As String
    If Len(mProgressCompleteStatus) = 0 Then
        ProgressCompleteStatusText = OPERATION_PROGRESS_COMPLETE_STATUS
    Else
        ProgressCompleteStatusText = mProgressCompleteStatus
    End If
End Function

Private Function ProgressFailedStatusText() As String
    If Len(mProgressFailedStatus) = 0 Then
        ProgressFailedStatusText = OPERATION_PROGRESS_FAILED_STATUS
    Else
        ProgressFailedStatusText = mProgressFailedStatus
    End If
End Function

' Purpose: Performs core append process run summary by coordinating `ProcessRunSummaryForMessage` and `End` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   messageText (Input; String): Message text text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   sourceProcedure (Input; String): Procedure name recorded as the source of trace, error, status, or diagnostic information emitted by this call.
' Returns: String: Returns formatted core append process run summary text intended for tracing, diagnostics, or user-facing reporting; vbNullString is returned when there is nothing applicable to report.
Public Function CoreAppendProcessRunSummary(ByVal messageText As String, ByVal sourceProcedure As String) As String
    'Intended use: Appends the authoritative Core tracked-process name and elapsed-time summary to a diagnostic or completion message without duplicating an existing process-run section.
    If InStr(1, messageText, "Total elapsed time:", vbTextCompare) > 0 Or (InStr(1, messageText, "Process run:", vbTextCompare) > 0 And (InStr(1, messageText, "Elapsed time:", vbTextCompare) > 0 Or InStr(1, messageText, "Top-level process elapsed:", vbTextCompare) > 0)) Then
        CoreAppendProcessRunSummary = messageText
    Else
        CoreAppendProcessRunSummary = messageText & ProcessRunSummaryForMessage(sourceProcedure)
    End If
End Function

Private Sub BeginOperationProgress(ByVal progressTitle As String, Optional ByVal maxSteps As Long = 0, Optional ByVal progressScopeLabel As String = vbNullString)
    If maxSteps <= 0 Then maxSteps = OPERATION_PROGRESS_DEFAULT_MAX_STEPS
    BeginProgressInternal progressTitle, maxSteps, DEFAULT_OPERATION_PROGRESS_STATUS_PREFIX, OPERATION_PROGRESS_COMPLETE_STATUS, OPERATION_PROGRESS_FAILED_STATUS, ProgressScopeFromTitle(progressTitle, progressScopeLabel)
End Sub

' Purpose: Starts a top-level operation progress scope only when the caller does not already have an active parent progress scope.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   progressTitle (Input; String): Title text used as the progress title shown, stored, compared, or resolved by this operation.
'   maxSteps (Input Optional; Long): Supplies the max steps used to bound, report, or validate the number of items processed by this operation. Default: 0.
'   progressScopeLabel (Input Optional; String): Scope value limiting which progress label items or operations are included in this call. Default: vbNullString.
'   forceNewTopLevelOperation (Input Optional; Boolean): When True, enables or requests new top level operation for this call; when False, the routine follows its normal non-forced/default path. Default: False.
' Returns: Boolean: True when this call starts and owns a new top-level operation-progress scope; False when an existing parent progress scope is reused and no new scope is started.
Public Function CoreBeginOperationProgressIfNeeded(ByVal progressTitle As String, Optional ByVal maxSteps As Long = 0, Optional ByVal progressScopeLabel As String = vbNullString, Optional ByVal forceNewTopLevelOperation As Boolean = False) As Boolean
    'Canonical Core authority for starting a top-level operation without competing progress state.
    If mSetupProgressActive Then
        If forceNewTopLevelOperation = False And ActiveProgressBelongsToCurrentTraceRun() Then
            CorePulseProgress progressTitle
            CoreBeginOperationProgressIfNeeded = False
            Exit Function
        End If
        ResetOrphanedProgressState progressTitle
    End If
    BeginOperationProgress progressTitle, maxSteps, progressScopeLabel
    CoreBeginOperationProgressIfNeeded = True
End Function

Private Sub BeginProgressInternal(ByVal progressTitle As String, ByVal maxSteps As Long, ByVal statusPrefix As String, ByVal completeStatus As String, ByVal failedStatus As String, ByVal progressScopeLabel As String)
    On Error GoTo ErrHandler
    Dim trackedProcessStarted As Boolean
    Dim initializationStep As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    ClearFailedProgressContext
    ClearProgressChildDetailState
    initializationStep = "Clear prior progress state"
    EndSetupProgress True
    initializationStep = "Start tracked process"
    BeginTrackedProcess progressTitle
    trackedProcessStarted = True
    mProgressOwnsTraceRun = False
    initializationStep = "Start or reuse trace run"
    If TraceRunActive() = False Then
        StartTraceRun "CoreProgress"
        mProgressOwnsTraceRun = True
    End If
    initializationStep = "Record progress owner trace run"
    mProgressOwnerTraceRunID = GetActiveOrLastTraceRunID()
    If Len(mProgressOwnerTraceRunID) = 0 Then Err.Raise vbObjectError + 6328, "BeginProgressInternal", "An Access Development Console progress owner trace run ID could not be established."
    initializationStep = "Initialize progress state"
    mSetupProgressMaxSteps = maxSteps
    mProgressStatusPrefix = statusPrefix
    mProgressCompleteStatus = completeStatus
    mProgressFailedStatus = failedStatus
    mProgressScopeLabel = CleanProgressScopeLabel(progressScopeLabel)
    initializationStep = "Open shared progress display"
    TrySharedProgressBegin mSetupProgressActive, mSetupProgressCurrentStep, mSetupProgressMaxSteps, progressTitle, ProgressStatusPrefixText(), "BeginProgressInternal", "Core", "Core"
    initializationStep = "Publish parent progress state"
    PublishParentProgressState
    initializationStep = "Write progress-start trace"
    LogTraceMessage TRACE_LEVEL_INFO, "Progress", "BeginProgressInternal", progressTitle, "Scope: " & mProgressScopeLabel & vbCrLf & "Max steps: " & CStr(maxSteps), progressTitle, 0, vbNullString, vbNullString, vbNullString, vbNullString, "CodeDb", "Core"
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    mLastFailedProgressText = "Progress initialization failed > " & progressTitle & " > " & initializationStep
    On Error Resume Next
    LogTraceMessage TRACE_LEVEL_ERROR, "Progress", "BeginProgressInternal", mLastFailedProgressText, "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, initializationStep, savedErrNumber, savedErrSource, savedErrDescription, vbNullString, vbNullString, "CodeDb", "Core"
    Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss"); " | ERROR | BeginProgressInternal | "; mLastFailedProgressText; " | "; CStr(savedErrNumber); " | "; savedErrDescription
    ClearProgressChildDetailState
    ClearParentProgressState
    TrySharedProgressEnd mSetupProgressActive, mSetupProgressCurrentStep, mSetupProgressMaxSteps, True, "BeginProgressInternal", "Core", "Core"
    If mProgressOwnsTraceRun Then CompleteTraceRun mLastFailedProgressText, False
    If trackedProcessStarted Then
        CaptureTrackedProcessSummary
        EndTrackedProcess
    End If
    mProgressOwnsTraceRun = False
    mProgressOwnerTraceRunID = vbNullString
    mProgressStatusPrefix = vbNullString
    mProgressCompleteStatus = vbNullString
    mProgressFailedStatus = vbNullString
    mProgressScopeLabel = vbNullString
    On Error GoTo 0
    Err.Raise savedErrNumber, "BeginProgressInternal", "Could not initialize Access Development Console progress at step '" & initializationStep & "'." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

' Purpose: Performs the public core begin progress child detail operation using the routine's existing ADC workflow and side effects.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   childScopeLabel (Input; String): Input child scope label used by this procedure to perform core begin progress child detail; validation and accepted values are enforced by this routine before dependent work
'     proceeds.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub CoreBeginProgressChildDetail(ByVal childScopeLabel As String)
    'Keeps nested child-progress labeling in the same Core-owned progress state machine.
    If mProgressChildScopeStack Is Nothing Then Set mProgressChildScopeStack = New Collection
    mProgressChildScopeStack.Add mProgressChildScopeLabel
    mProgressChildDetailDepth = mProgressChildDetailDepth + 1
    mProgressChildScopeLabel = CleanProgressScopeLabel(childScopeLabel)
End Sub

' Purpose: Performs the public core end progress child detail operation using the routine's existing ADC workflow and side effects.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub CoreEndProgressChildDetail()
    If mProgressChildDetailDepth > 0 Then mProgressChildDetailDepth = mProgressChildDetailDepth - 1
    If Not mProgressChildScopeStack Is Nothing Then
        If mProgressChildScopeStack.count > 0 Then
            mProgressChildScopeLabel = CStr(mProgressChildScopeStack(mProgressChildScopeStack.count))
            mProgressChildScopeStack.Remove mProgressChildScopeStack.count
        Else
            mProgressChildScopeLabel = vbNullString
        End If
    Else
        mProgressChildScopeLabel = vbNullString
    End If
End Sub

Private Sub ClearProgressChildDetailState()
    mProgressChildDetailDepth = 0
    mProgressChildScopeLabel = vbNullString
    Set mProgressChildScopeStack = Nothing
End Sub

Private Function ProgressChildDetailText(ByVal statusText As String) As String
    statusText = Trim$(statusText)
    If mProgressChildDetailDepth > 0 And Len(mProgressChildScopeLabel) > 0 Then
        ProgressChildDetailText = mProgressChildScopeLabel & " > " & statusText
    Else
        ProgressChildDetailText = statusText
    End If
End Function

' ===============================================================================================================================================================================
' SETUP EXECUTION / PHASE CONTROL / WRITE-GATE PROCEDURES
' ===============================================================================================================================================================================

Private Sub BeginSetupProgress(ByVal progressTitle As String, Optional ByVal maxSteps As Long = 0, Optional ByVal progressScopeLabel As String = "Setup")
    If maxSteps <= 0 Then maxSteps = OPERATION_PROGRESS_DEFAULT_MAX_STEPS
    BeginProgressInternalForSetup progressTitle, maxSteps, DEFAULT_UTILITY_SHORT_NAME & ": Setup > ", DEFAULT_UTILITY_SHORT_NAME & ": Setup complete.", DEFAULT_UTILITY_SHORT_NAME & ": Setup failed. Review the Immediate Window trace.", ProgressScopeFromTitle(progressTitle, progressScopeLabel)
End Sub

Private Function SetupPhaseProgressPrefixText() As String
    If mSetupPreflightContextActive Then
        SetupPhaseProgressPrefixText = "Setup Preflight > "
    ElseIf Len(mSetupPhaseIdentity) > 0 Then
        SetupPhaseProgressPrefixText = "Setup " & mSetupPhaseIdentity & "  " & mSetupPhaseLabel & " " & CStr(mSetupPhasePartNumber) & " of " & CStr(mSetupPhasePartCount) & " > "
    End If
End Function

Private Function SetupPhaseProgressScopeText() As String
    Dim prefixText As String
    prefixText = SetupPhaseProgressPrefixText()
    If Right$(prefixText, 3) = " > " Then prefixText = Left$(prefixText, Len(prefixText) - 3)
    SetupPhaseProgressScopeText = prefixText
End Function

' Purpose: Initializes the shared preflight, trace, progress, Setup Control, and Module Guard state required before a standalone Setup Part or phase can execute.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   partIdentity (Input; String): Canonical Setup Part identity (for example A2 or G1) used to select, validate, and report the exact setup part being operated on.
'   moduleGuardVerificationDetails (Output; String): Receives module guard verification details used for caller diagnostics, verification output, or user-facing reporting.
'   traceRunOwnedByProcedure (Output; Boolean): Set True only when this procedure starts the trace run and therefore owns closing it; False when an existing trace run remains caller-owned.
'   operationIdentity (Input Optional; String): Optional display/trace identity for the overall setup operation; blank means the routine uses the Setup Part identity as the operation label. Default: vbNullString.
'   progressPartCount (Input Optional; Long): Supplies the progress part count used to bound, report, or validate the number of items processed by this operation. Default: 1.
' Returns: None (Sub). Results are returned through the documented ByRef parameters and/or the routine's database, UI, or module-state side effects.
Public Sub CoreBeginStandaloneSetupPartRun(ByVal partIdentity As String, ByRef moduleGuardVerificationDetails As String, ByRef traceRunOwnedByProcedure As Boolean, Optional ByVal operationIdentity As String = vbNullString, Optional ByVal progressPartCount As Long = 1)
    On Error GoTo ErrHandler
    Dim failedErrNumber As Long
    Dim failedErrSource As String
    Dim failedErrDescription As String
    Dim operationName As String
    Dim displayIdentity As String
    Dim infrastructurePreflightDetails As String
    partIdentity = UCase$(Trim$(partIdentity))
    If Len(partIdentity) = 0 Then Err.Raise vbObjectError + 6494, "CoreBeginStandaloneSetupPartRun", "A Setup Part identity is required."
    If mStandaloneSetupPartsOperationActive Then Err.Raise vbObjectError + 6495, "CoreBeginStandaloneSetupPartRun", "Another standalone Setup Part operation is already active."
    displayIdentity = Trim$(operationIdentity)
    If Len(displayIdentity) = 0 Then displayIdentity = partIdentity
    operationName = DEFAULT_UTILITY_LONG_NAME & ": Setup " & displayIdentity
    mSetupCancelled = False
    traceRunOwnedByProcedure = False
    CoreSetSetupPreflightContext True
    PrepareSetupTopLevelOperation operationName, "Setup " & displayIdentity, traceRunOwnedByProcedure
    If progressPartCount < 1 Then progressPartCount = 1
    BeginSetupProgress DEFAULT_UTILITY_SHORT_NAME & ": Setup " & displayIdentity, progressPartCount, "Setup " & displayIdentity
    CorePulseProgress "Resolving Module Guard state"
    If TryResolveModuleUpdateGuardForSetup(failedErrNumber, failedErrSource, failedErrDescription, moduleGuardVerificationDetails) = False Then
        If failedErrNumber = 0 Then failedErrNumber = vbObjectError + 6496
        If Len(Trim$(failedErrSource)) = 0 Then failedErrSource = "CoreBeginStandaloneSetupPartRun"
        If Len(Trim$(failedErrDescription)) = 0 Then failedErrDescription = "Module Guard preflight failed for standalone Setup " & displayIdentity & "."
        Err.Raise failedErrNumber, failedErrSource, failedErrDescription
    End If
    RequireNoUnfinishedSetupRecoveryState
    If partIdentity = "A2" Or StrComp(displayIdentity, "Phase A", vbTextCompare) = 0 Then
        PrepareSetupDataTableBackupForPreflight infrastructurePreflightDetails
        LogTraceMessage TRACE_LEVEL_INFO, "SetupPreflight", "CoreBeginStandaloneSetupPartRun", "Standalone Setup infrastructure ownership preflight completed.", infrastructurePreflightDetails, "Setup Preflight infrastructure", 0, vbNullString, vbNullString, vbNullString, "SetupInfrastructure", "CodeDb", "Setup"
    End If
    LogSetupStartupStage "CoreBeginStandaloneSetupPartRun", "Setup Control startup call starting.", "Setup Part: " & displayIdentity
    BeginSetupControl "Setup Preflight is complete. " & displayIdentity & " is preparing to run."
    LogSetupStartupStage "CoreBeginStandaloneSetupPartRun", "Setup Control startup call completed.", "Setup Part: " & displayIdentity
    mSetupCurrentPartText = vbNullString
    mSetupPartRollbackStartIndex = 0
    mSetupRunContextStarted = False
    mSetupNamingSnapshotStarted = False
    mSetupMaterialCommitted = False
    mSetupCurrentStep = vbNullString
    ResetFinalSetupJ1Staging
    ResetTraceNotes
    mProgressOwnsTraceRun = False
    CoreResetOperationNotes
    CoreBeginSetupNamingSnapshot
    BeginSetupRunContext moduleGuardVerificationDetails
    mSetupRunContextStarted = True
    BeginSetupOperation operationName
    mStandaloneSetupPartsOperationActive = True
    CoreAddOperationNote "Standalone Setup " & displayIdentity & " started after read-only Setup Preflight. Earlier Parts will not be executed automatically."
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim rollbackSummary As String
    Dim rollbackFailed As Boolean
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    CaptureFailedSetupPhaseContext
    On Error Resume Next
    CoreSetSetupPreflightContext False
    If mSetupRollbackActive Then RollbackSetupOperation rollbackSummary, rollbackFailed
    If mSetupRunContextStarted Then FinalizeSetupRunContext "Failed", False, savedErrNumber, savedErrSource, savedErrDescription & vbCrLf & rollbackSummary
    mSetupRunContextStarted = False
    If mSetupNamingSnapshotStarted Then InvalidateNamingSnapshot
    mSetupNamingSnapshotStarted = False
    ReleaseSetupPreflightNamingSnapshot True
    If mSetupControlActive Then
        EndSetupControl
    End If
    If mSetupProgressActive Then CoreFailSetupProgress
    CoreResetSetupPhaseContext
    mStandaloneSetupPartsOperationActive = False
    If traceRunOwnedByProcedure Then CompleteTraceRun operationName & " failed during preflight/lifecycle initialization.", False
    On Error GoTo 0
    If rollbackFailed Then Err.Raise ERR_SETUP_ROLLBACK_RECOVERY_FAILED, "CoreBeginStandaloneSetupPartRun", "Could not begin standalone Setup " & displayIdentity & ", and rollback of the partially established lifecycle did not complete." & vbCrLf & rollbackSummary & vbCrLf & "Original error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription
    Err.Raise savedErrNumber, "CoreBeginStandaloneSetupPartRun", "Could not begin standalone Setup " & displayIdentity & "." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription & IIf(Len(Trim$(rollbackSummary)) = 0, vbNullString, vbCrLf & rollbackSummary)
End Sub

' Purpose: Performs core complete standalone setup part run by coordinating `ReleaseSetupPreflightNamingSnapshot` and `CoreAddOperationNote` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   partIdentity (Input; String): Canonical Setup Part identity (for example A2 or G1) used to select, validate, and report the exact setup part being operated on.
'   timingText (Input Optional; String): Input timing text for core complete standalone setup part run; is tested by If Len(Trim$(timingText)) > 0. Default: vbNullString.
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: False.
'   traceRunOwnedByProcedure (Input Optional; Boolean): Indicates whether this procedure started the active trace run; True makes it responsible for completing that run, while False leaves the caller-owned run
'   open. Default: False.
'   operationIdentity (Input Optional; String): Optional display/trace identity for the overall setup operation; blank means the routine uses the Setup Part identity as the operation label. Default: vbNullString.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub CoreCompleteStandaloneSetupPartRun(ByVal partIdentity As String, Optional ByVal timingText As String = vbNullString, Optional ByVal showMessage As Boolean = False, Optional ByVal traceRunOwnedByProcedure As Boolean = False, Optional ByVal operationIdentity As String = vbNullString)
    On Error GoTo ErrHandler
    Dim operationName As String
    Dim displayIdentity As String
    partIdentity = UCase$(Trim$(partIdentity))
    displayIdentity = Trim$(operationIdentity)
    If Len(displayIdentity) = 0 Then displayIdentity = partIdentity
    operationName = DEFAULT_UTILITY_LONG_NAME & ": Setup " & displayIdentity
    RequireSetupControlContinue IIf(Len(Trim$(mSetupCurrentStep)) = 0, partIdentity & " completion", mSetupCurrentStep)
    CommitSetupOperation
    mSetupMaterialCommitted = True
    If partIdentity = "J1" Then PersistStagedFinalSetupStateForJ1
    If Len(Trim$(timingText)) > 0 Then LogTraceMessage TRACE_LEVEL_INFO, "SetupTiming", "CoreCompleteStandaloneSetupPartRun", "Standalone Setup " & displayIdentity & " timing captured.", timingText, "Complete " & displayIdentity, 0, vbNullString, vbNullString, vbNullString, "SetupPart", "CodeDb", "SetupTiming"
    If mSetupRunContextStarted Then FinalizeSetupRunContext "Completed", CBool(partIdentity = "J1")
    mSetupRunContextStarted = False
    If mSetupNamingSnapshotStarted Then InvalidateNamingSnapshot
    mSetupNamingSnapshotStarted = False
    ReleaseSetupPreflightNamingSnapshot True
    CoreAddOperationNote "Standalone Setup " & displayIdentity & " completed successfully."
    CoreCloseSetupCompletionUI
    If showMessage Then ReportCoreActionSuccess "CoreCompleteStandaloneSetupPartRun", "Setup", "Standalone Setup " & displayIdentity & " completed successfully.", "Only Setup " & displayIdentity & " was executed. Earlier Parts were verified read-only and were not run automatically.", "Complete " & displayIdentity, vbNullString, "SetupPart", "CodeDb", "Setup", True, operationName
    CoreCompleteSetupProgress
    CoreResetSetupPhaseContext
    ResetFinalSetupJ1Staging
    mStandaloneSetupPartsOperationActive = False
    If traceRunOwnedByProcedure Then CompleteTraceRun operationName & " completed."
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "CoreCompleteStandaloneSetupPartRun", Err.Description
End Sub

' Purpose: Performs core handle standalone setup part failure by coordinating `End` and `ReleaseSetupPreflightNamingSnapshot` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   partIdentity (Input; String): Canonical Setup Part identity (for example A2 or G1) used to select, validate, and report the exact setup part being operated on.
'   errNumber (Input; Long): Numeric identifier or ordinal for the err item processed or reported by this operation.
'   errSource (Input; String): Error source text carried into the returned/raised diagnostic; blank allows the procedure to supply its own source when coded to do so.
'   errDescription (Input; String): Error description text carried into the returned/raised diagnostic; blank allows the procedure to supply its own description when coded to do so.
'   traceRunOwnedByProcedure (Input Optional; Boolean): Indicates whether this procedure started the active trace run; True makes it responsible for completing that run, while False leaves the caller-owned run
'   open. Default: False.
'   operationIdentity (Input Optional; String): Optional display/trace identity for the overall setup operation; blank means the routine uses the Setup Part identity as the operation label. Default: vbNullString.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub CoreHandleStandaloneSetupPartFailure(ByVal partIdentity As String, ByVal errNumber As Long, ByVal errSource As String, ByVal errDescription As String, Optional ByVal traceRunOwnedByProcedure As Boolean = False, Optional ByVal operationIdentity As String = vbNullString)
    Dim rollbackSummary As String
    Dim rollbackFailed As Boolean
    Dim cancelRollbackSummary As String
    Dim cancelRollbackFailed As Boolean
    Dim uiRollbackSummary As String
    Dim uiRollbackSucceeded As Boolean
    Dim cancellationMessage As String
    Dim operationName As String
    Dim failedStepText As String
    Dim displayIdentity As String
    failedStepText = mSetupCurrentStep
    partIdentity = UCase$(Trim$(partIdentity))
    displayIdentity = Trim$(operationIdentity)
    If Len(displayIdentity) = 0 Then displayIdentity = partIdentity
    operationName = DEFAULT_UTILITY_LONG_NAME & ": Setup " & displayIdentity
    If mSetupNamingSnapshotStarted Then
        On Error Resume Next
        InvalidateNamingSnapshot
        On Error GoTo 0
        mSetupNamingSnapshotStarted = False
    End If
    ReleaseSetupPreflightNamingSnapshot True
    If errNumber = ERR_SETUP_CANCELLED_BY_USER Then
        uiRollbackSucceeded = CancelSetupPreparedWork(uiRollbackSummary)
        If mSetupRollbackActive Then
            RollbackSetupCurrentPart cancelRollbackSummary, cancelRollbackFailed
        Else
            cancelRollbackSummary = "No rollback-tracked Setup changes were active at the cancellation checkpoint."
        End If
        If cancelRollbackFailed Or uiRollbackSucceeded = False Then
            If mSetupRunContextStarted Then FinalizeSetupRunContext "Failed", False, ERR_SETUP_ROLLBACK_RECOVERY_FAILED, "CoreHandleStandaloneSetupPartFailure", cancelRollbackSummary & vbCrLf & uiRollbackSummary
            mSetupRunContextStarted = False
            mStandaloneSetupPartsOperationActive = False
            CoreFailSetupProgress
            If traceRunOwnedByProcedure Then CompleteTraceRun operationName & " cancellation rollback failed.", False
            Err.Raise ERR_SETUP_ROLLBACK_RECOVERY_FAILED, "CoreHandleStandaloneSetupPartFailure", "Standalone Setup " & displayIdentity & " cancellation reached a safe checkpoint, but rollback did not complete." & vbCrLf & cancelRollbackSummary & vbCrLf & uiRollbackSummary
        End If
        cancellationMessage = "Standalone Setup " & displayIdentity & " was cancelled at a safe Setup checkpoint." & vbCrLf & vbCrLf & cancelRollbackSummary & vbCrLf & uiRollbackSummary
        If mSetupRunContextStarted Then FinalizeSetupRunContext "Cancelled", False, ERR_SETUP_CANCELLED_BY_USER, errSource, cancellationMessage
        mSetupRunContextStarted = False
        mStandaloneSetupPartsOperationActive = False
        RecordSetupCancellation cancellationMessage, False
        EndSetupProgress False
        CoreResetSetupPhaseContext
        If traceRunOwnedByProcedure Then CompleteTraceRun operationName & " cancelled.", False
        Exit Sub
    End If
    If mSetupRollbackActive Then RollbackSetupOperation rollbackSummary, rollbackFailed
    If mSetupRunContextStarted Then FinalizeSetupRunContext "Failed", mSetupMaterialCommitted, errNumber, errSource, errDescription & vbCrLf & rollbackSummary
    mSetupRunContextStarted = False
    If mSetupControlActive Then
        On Error Resume Next
        EndSetupControl
        On Error GoTo 0
    End If
    CoreFailSetupProgress
    CoreResetSetupPhaseContext
    mStandaloneSetupPartsOperationActive = False
    If traceRunOwnedByProcedure Then CompleteTraceRun operationName & " failed.", False
    If rollbackFailed Then Err.Raise ERR_SETUP_ROLLBACK_RECOVERY_FAILED, "CoreHandleStandaloneSetupPartFailure", "Standalone Setup " & displayIdentity & " failed and rollback did not complete." & vbCrLf & rollbackSummary & vbCrLf & "Original operation error " & CStr(errNumber) & " from " & errSource & ": " & errDescription
    Err.Raise errNumber, "CoreHandleStandaloneSetupPartFailure", "Could not run standalone Setup " & displayIdentity & "." & vbCrLf & "Setup step: " & IIf(Len(Trim$(failedStepText)) = 0, "(not recorded)", failedStepText) & vbCrLf & "Err.Source: " & errSource & vbCrLf & "Err.Description: " & errDescription & vbCrLf & rollbackSummary
End Sub

' Purpose: Performs the public core update setup progress operation using the routine's existing ADC workflow and side effects.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   stepText (Input; String): Input step text for core update setup progress; is passed to `ProgressPartText`.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub CoreUpdateSetupProgress(ByVal stepText As String)
    'Updates the live setup display and retains the existing persistent phase-boundary trace row.
    Dim partStepText As String
    partStepText = ProgressPartText(stepText, True)
    TrySharedProgressUpdate mSetupProgressActive, mSetupProgressCurrentStep, mSetupProgressMaxSteps, partStepText, ProgressStatusPrefixText(), "CoreUpdateSetupProgress", "Core", "Core"
    PublishParentProgressState
End Sub

' Purpose: Performs the public core complete setup progress operation using the routine's existing ADC workflow and side effects.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub CoreCompleteSetupProgress()
    ClearFailedProgressContext
    If mSetupProgressActive Then
        TrySharedProgressComplete mSetupProgressActive, mSetupProgressCurrentStep, mSetupProgressMaxSteps, ProgressCompleteStatusText(), True, "CoreCompleteSetupProgress", "Core", "Setup"
    End If
    If mProgressOwnsTraceRun Then CompleteTraceRun ProgressCompleteStatusText()
    mProgressOwnsTraceRun = False
    mProgressOwnerTraceRunID = vbNullString
    mProgressStatusPrefix = vbNullString
    mProgressCompleteStatus = vbNullString
    mProgressFailedStatus = vbNullString
    mProgressScopeLabel = vbNullString
    ClearProgressChildDetailState
    ClearParentProgressState
    EndTrackedProcess
End Sub

' Purpose: Performs the public core fail setup progress operation using the routine's existing ADC workflow and side effects.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub CoreFailSetupProgress()
    Dim failedStatusText As String
    If mSetupPreflightContextActive Or Len(mSetupPhaseIdentity) > 0 Then CaptureFailedSetupPhaseContext
    failedStatusText = ProgressChildDetailText(ProgressFailedStatusText())
    If Len(mLastFailedProgressPartPrefix) = 0 Then
        If Len(SetupPhaseProgressPrefixText()) > 0 Then
            mLastFailedProgressPartPrefix = SetupPhaseProgressPrefixText()
        ElseIf mSetupProgressActive And mSetupProgressMaxSteps > 0 Then
            mLastFailedProgressPartPrefix = ProgressPartPrefixText() & "Part " & CStr(IIf(mSetupProgressCurrentStep < 1, 1, IIf(mSetupProgressCurrentStep > mSetupProgressMaxSteps, mSetupProgressMaxSteps, mSetupProgressCurrentStep))) & " of " & CStr(mSetupProgressMaxSteps) & " > "
        Else
            mLastFailedProgressPartPrefix = vbNullString
        End If
    End If
    If Len(Trim$(mLastFailedProgressText)) = 0 Then mLastFailedProgressText = mLastFailedProgressPartPrefix & failedStatusText
    TrySharedProgressFail mSetupProgressActive, mSetupProgressCurrentStep, mSetupProgressMaxSteps, mLastFailedProgressText, "CoreFailSetupProgress", "Core", "Setup"
    LogTraceMessage TRACE_LEVEL_ERROR, "Core", "CoreFailSetupProgress", mLastFailedProgressText, ProcessRunSummaryForMessage("CoreFailSetupProgress"), mLastFailedProgressText
    If mProgressOwnsTraceRun Then CompleteTraceRun mLastFailedProgressText, False
    mProgressOwnsTraceRun = False
    mProgressOwnerTraceRunID = vbNullString
    CaptureTrackedProcessSummary
    mProgressStatusPrefix = vbNullString
    mProgressCompleteStatus = vbNullString
    mProgressFailedStatus = vbNullString
    mProgressScopeLabel = vbNullString
    ClearProgressChildDetailState
    ClearParentProgressState
    EndTrackedProcess
End Sub

' Purpose: Closes setup progress, coordinating `TrySharedProgressEnd` and `End` as the implementation requires.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   clearStatus (Input Optional; Boolean): When True, enables or requests status for this call; when False, the routine follows its normal non-forced/default path. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub EndSetupProgress(Optional ByVal clearStatus As Boolean = True)
    ClearFailedProgressContext
    TrySharedProgressEnd mSetupProgressActive, mSetupProgressCurrentStep, mSetupProgressMaxSteps, clearStatus, "EndSetupProgress", "Core", "Core"
    mProgressOwnsTraceRun = False
    mProgressOwnerTraceRunID = vbNullString
    mProgressStatusPrefix = vbNullString
    mProgressCompleteStatus = vbNullString
    mProgressFailedStatus = vbNullString
    mProgressScopeLabel = vbNullString
    ClearProgressChildDetailState
    ClearParentProgressState
End Sub

Private Sub CoreSetSetupPreflightContext(ByVal Active As Boolean)
    mSetupPreflightContextActive = Active
    If Active Then
        mSetupPhaseIdentity = vbNullString
        mSetupPhaseLabel = vbNullString
        mSetupPhasePartNumber = 0
        mSetupPhasePartCount = 0
    End If
End Sub

' Purpose: Performs core begin setup phase part by coordinating `SetStep` and `End` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   phasePartIdentity (Input; String): Stable phase part identity used to select or correlate the exact ADC definition/state handled by this procedure.
'   phaseLabel (Input; String): Display or grouping label used as the phase label for the UI, Ribbon, progress, or definition item being processed.
'   phasePartNumber (Input; Long): Numeric identifier or ordinal for the phase part item processed or reported by this operation.
'   phasePartCount (Input; Long): Supplies the phase part count used to bound, report, or validate the number of items processed by this operation.
'   stepText (Input; String): Input step text for core begin setup phase part; is passed to `SetStep`.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub CoreBeginSetupPhasePart(ByVal phasePartIdentity As String, ByVal phaseLabel As String, ByVal phasePartNumber As Long, ByVal phasePartCount As Long, ByVal stepText As String)
    mSetupPreflightContextActive = False
    mSetupPhaseIdentity = UCase$(Trim$(phasePartIdentity))
    mSetupPhaseLabel = Trim$(phaseLabel)
    mSetupPhasePartNumber = phasePartNumber
    mSetupPhasePartCount = phasePartCount
    SetStep mSetupCurrentStep, stepText
End Sub

Private Sub CoreHandleSetupRunFailure(ByVal scopeLabel As String, ByVal completionLabel As String, ByVal errNumber As Long, ByVal errSource As String, ByVal errDescription As String)
    Dim rollbackSummary As String
    Dim rollbackFailed As Boolean
    Dim cancelRollbackSummary As String
    Dim cancelRollbackFailed As Boolean
    Dim uiRollbackSummary As String
    Dim uiRollbackSucceeded As Boolean
    Dim cancellationMessage As String
    Dim failedStepText As String
    failedStepText = mSetupCurrentStep
    If mSetupNamingSnapshotStarted Then
        On Error Resume Next
        InvalidateNamingSnapshot
        On Error GoTo 0
        mSetupNamingSnapshotStarted = False
    End If
    ReleaseSetupPreflightNamingSnapshot True
    If errNumber = ERR_SETUP_CANCELLED_BY_USER Then
        uiRollbackSucceeded = CancelSetupPreparedWork(uiRollbackSummary)
        If mSetupRollbackActive Then
            RollbackSetupCurrentPart cancelRollbackSummary, cancelRollbackFailed
        Else
            cancelRollbackSummary = "No rollback-tracked Setup changes were active at the cancellation checkpoint."
        End If
        If cancelRollbackFailed Or uiRollbackSucceeded = False Then
            If mSetupRunContextStarted Then FinalizeSetupRunContext "Failed", False, ERR_SETUP_ROLLBACK_RECOVERY_FAILED, "CoreHandleSetupRunFailure", cancelRollbackSummary & vbCrLf & uiRollbackSummary
            mSetupRunContextStarted = False
            CoreFailSetupProgress
            Err.Raise ERR_SETUP_ROLLBACK_RECOVERY_FAILED, "CoreHandleSetupRunFailure", "Setup cancellation reached a safe checkpoint, but rollback did not complete." & vbCrLf & cancelRollbackSummary & vbCrLf & uiRollbackSummary
        End If
        cancellationMessage = DEFAULT_UTILITY_LONG_NAME & ": Setup - CodeDb cancelled by user at a safe Setup checkpoint." & vbCrLf & vbCrLf & cancelRollbackSummary & vbCrLf & uiRollbackSummary
        If mSetupRunContextStarted Then FinalizeSetupRunContext "Cancelled", False, ERR_SETUP_CANCELLED_BY_USER, errSource, cancellationMessage
        mSetupRunContextStarted = False
        RecordSetupCancellation cancellationMessage, False
        EndSetupProgress False
        CoreResetSetupPhaseContext
        Exit Sub
    End If
    If mSetupRollbackActive Then RollbackSetupOperation rollbackSummary, rollbackFailed
    If mSetupRunContextStarted Then FinalizeSetupRunContext "Failed", mSetupMaterialCommitted, errNumber, errSource, errDescription & vbCrLf & rollbackSummary
    mSetupRunContextStarted = False
    CoreFailSetupProgress
    CoreResetSetupPhaseContext
    If rollbackFailed Then Err.Raise ERR_SETUP_ROLLBACK_RECOVERY_FAILED, "CoreHandleSetupRunFailure", "Setup through " & scopeLabel & " failed and rollback did not complete." & vbCrLf & rollbackSummary & vbCrLf & "Original operation error " & CStr(errNumber) & " from " & errSource & ": " & errDescription
    Err.Raise errNumber, "CoreHandleSetupRunFailure", "Could not run Setup Preflight + " & completionLabel & "." & vbCrLf & "Setup step: " & IIf(Len(Trim$(failedStepText)) = 0, "(not recorded)", failedStepText) & vbCrLf & "Err.Source: " & errSource & vbCrLf & "Err.Description: " & errDescription & vbCrLf & rollbackSummary
End Sub

Private Sub CoreCloseSetupCompletionUI()
    If mSetupControlActive Then
        EndSetupControl
    End If
    If mSetupProgressActive Then
        TrySharedProgressComplete mSetupProgressActive, mSetupProgressCurrentStep, mSetupProgressMaxSteps, ProgressCompleteStatusText(), True, "CoreCloseSetupCompletionUI", "Core", "Setup"
    End If
    DoEvents
End Sub

Private Sub RecordSetupCancellation(ByVal cancelMessage As String, Optional ByVal callerOwnsLifecycle As Boolean = False)
    'Records Setup/Repair cancellation without closing progress, trace, or tracked-process state owned by the direct entry point or system-command dispatcher.
    Dim cancellationContext As String
    ClearFailedProgressContext
    cancellationContext = IIf(StrComp(mProgressScopeLabel, "Repair", vbTextCompare) = 0, "Repair", "Setup")
    If Len(Trim$(cancelMessage)) = 0 Then cancelMessage = DEFAULT_OPERATION_PROGRESS_STATUS_PREFIX & cancellationContext & " cancelled by user."
    mSetupCancelled = True
    If callerOwnsLifecycle Then MarkSystemCommandCancelled
    On Error Resume Next
    CoreSetVisibleStatusText ProgressPartText(cancelMessage, False)
    If TraceRunActive() Then LogTraceMessage TRACE_LEVEL_WARNING, cancellationContext, "RecordSetupCancellation", cancelMessage, cancelMessage, cancellationContext & " cancelled", 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config"), "Table", "CodeDb", cancellationContext
    On Error GoTo 0
End Sub

Private Sub LogSetupStartupStage(ByVal sourceProcedure As String, ByVal messageText As String, ByVal detailText As String, Optional ByVal traceLevel As String = TRACE_LEVEL_INFO)
    If Len(Trim$(detailText)) = 0 Then detailText = messageText
    LogTraceMessage traceLevel, "SetupStartup", sourceProcedure, messageText, detailText, "Setup Preflight", 0, vbNullString, vbNullString, vbNullString, "SetupControl", "CodeDb", "Setup"
    FlushTraceLogFile
End Sub

Private Function SetupErrorMessage(ByVal ProcedureName As String, ByVal errNumber As Long, ByVal errSource As String, ByVal errDescription As String) As String
    SetupErrorMessage = DEFAULT_UTILITY_LONG_NAME & ": Setup failed." & vbCrLf & vbCrLf & FailedSetupPhaseContextForMessage() & FailureProgressContextForMessage() & vbCrLf & vbCrLf & "Procedure: " & ProcedureName & vbCrLf & "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB & vbCrLf & "Database: " & SafeDatabaseName() & vbCrLf & "MC_Records table: " & SafeTableName("ADC_Config") & vbCrLf & "CodeDb settings table: " & SafeTableName("settings") & vbCrLf & "Err.Number: " & CStr(errNumber) & vbCrLf & "Err.Source: " & errSource & vbCrLf & "Err.Description:" & vbCrLf & errDescription & TableDiagnosticsLocationForMessage() & TableReadinessDiagnosticsForMessage() & TraceSummaryForMessage() & CoreOperationSummaryForMessage() & ProcessRunSummaryForMessage(ProcedureName)
End Function

Private Sub ReportSetupError(ByVal ProcedureName As String, ByVal errNumber As Long, ByVal errSource As String, ByVal errDescription As String)
    On Error Resume Next
    Dim messageText As String
    messageText = SetupErrorMessage(ProcedureName, errNumber, errSource, errDescription)
    LogTraceMessage TRACE_LEVEL_ERROR, "Core", ProcedureName, DEFAULT_UTILITY_LONG_NAME & ": Setup failed.", messageText, vbNullString, errNumber, errSource, errDescription, vbNullString, vbNullString, "CodeDb", "Setup"
    CoreShowErrorMessage messageText, vbCritical, ProcedureName & " - Error Detail", ProcedureName, "Setup"
End Sub

' Purpose: Runs the complete ADC CodeDb Setup lifecycle through the defined setup phases, coordinating tracked progress, recovery, phase orchestration, final verification, and caller-facing failure handling.
' Public role: Public module entry point. This procedure is listed in the module header as an intentional externally callable ADC entry point and may be invoked directly by Access/VBA callers.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub RunFullSetup(Optional ByVal showMessage As Boolean = True)
    'Intended use: Runs Setup Access Development Console Code Db for direct, generated-macro, form-button, or required cross-module callers.
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim messageText As String
    Dim traceRunOwnedByProcedure As Boolean
    If TryRunFullSetup(showMessage, errNumber, errSource, errDescription, False, traceRunOwnedByProcedure) Then Exit Sub
    If errNumber = 0 Then errNumber = vbObjectError + 5236
    If Len(Trim$(errSource)) = 0 Then errSource = "RunFullSetup"
    If Len(Trim$(errDescription)) = 0 Then errDescription = SETUP_RESULT_FAILURE_HEADLINE
    If showMessage Then
        ReportSetupError "RunFullSetup", errNumber, errSource, errDescription
        If traceRunOwnedByProcedure Then CompleteTraceRun DEFAULT_UTILITY_LONG_NAME & ": Setup - CodeDb failed after final diagnostics were recorded.", False
        Exit Sub
    End If
    PrintSetupFailureDiagnostics "RunFullSetup", errNumber, errSource, errDescription
    messageText = SetupErrorMessage("RunFullSetup", errNumber, errSource, errDescription)
    If traceRunOwnedByProcedure Then CompleteTraceRun DEFAULT_UTILITY_LONG_NAME & ": Setup - CodeDb failed after final diagnostics were recorded.", False
    Err.Raise errNumber, "RunFullSetup", messageText
End Sub

Private Sub PrepareSetupTopLevelOperation(ByVal operationName As String, ByVal traceContext As String, ByRef traceRunOwnedByProcedure As Boolean)
    'Establishes the Setup trace/progress boundary without repairing prerequisite recovery owned by later domains. Preflight performs the read-only recovery gate before A1.
    Dim previousTraceRunID As String
    Dim newTraceRunID As String
    operationName = Trim$(operationName)
    traceContext = Trim$(traceContext)
    If Len(operationName) = 0 Then operationName = DEFAULT_UTILITY_LONG_NAME & " Setup"
    If Len(traceContext) = 0 Then traceContext = "Setup"
    traceRunOwnedByProcedure = False
    If mSetupProgressActive Then
        ResetOrphanedProgressState operationName
    Else
        ClearParentProgressState
        Set mCurrentProcessStack = Nothing
        mCurrentProcessActive = False
        mCurrentProcessName = vbNullString
        mCurrentProcessStartedOn = 0
        mCurrentProcessStartTimer = 0
        ClearLastEndedProcessSummary
    End If
    If TraceRunActive() Then
        previousTraceRunID = GetActiveTraceRunID()
        CompleteTraceRun "Previous active trace run closed before Setup started."
        If TraceRunActive() Then Err.Raise vbObjectError + 6330, "PrepareSetupTopLevelOperation", "The previous active " & DEFAULT_UTILITY_LONG_NAME & " trace run could not be closed before Setup."
    End If
    mProgressOwnsTraceRun = False
    mProgressOwnerTraceRunID = vbNullString
    StartTraceRunForSetup traceContext
    LogAccessProcessPriorityForCurrentTrace "PrepareSetupTopLevelOperation"
    newTraceRunID = GetActiveTraceRunID()
    If Len(newTraceRunID) = 0 Then Err.Raise vbObjectError + 6331, "PrepareSetupTopLevelOperation", "A fresh " & DEFAULT_UTILITY_LONG_NAME & " trace run could not be started for Setup."
    If Len(previousTraceRunID) > 0 And StrComp(newTraceRunID, previousTraceRunID, vbBinaryCompare) = 0 Then Err.Raise vbObjectError + 6332, "PrepareSetupTopLevelOperation", "Setup reused the previous trace run ID instead of starting a fresh trace."
    traceRunOwnedByProcedure = True
    LogTraceMessage TRACE_LEVEL_INFO, "Progress", "PrepareSetupTopLevelOperation", "Fresh Setup operation boundary started.", "Operation: " & operationName & vbCrLf & "Trace context: " & traceContext & vbCrLf & "Previous active trace run ID: " & IIf(Len(previousTraceRunID) = 0, "(none)", previousTraceRunID) & vbCrLf & "New active trace run ID: " & newTraceRunID, operationName, 0, vbNullString, vbNullString, vbNullString, "OperationBoundary", "CodeDb", traceContext
End Sub

Private Function TryRunFullSetup(ByVal showMessage As Boolean, ByRef failedErrNumber As Long, ByRef failedErrSource As String, ByRef failedErrDescription As String, Optional ByVal callerOwnsLifecycle As Boolean = False, Optional ByRef traceRunOwnedByProcedure As Boolean = False) As Boolean
    On Error GoTo ErrHandler
    Dim setupPopupMessagesEnabled As Boolean
    Dim setupProcessStartTimer As Single
    Dim moduleGuardResolved As Boolean
    Dim response As VbMsgBoxResult
    Dim guardVerificationDetails As String
    Dim infrastructurePreflightDetails As String
    failedErrNumber = 0
    failedErrSource = vbNullString
    failedErrDescription = vbNullString
    traceRunOwnedByProcedure = False
    mSetupCancelled = False
    response = ConfirmRunFullSetupImmediately()
    If response <> vbYes Then
        RecordSetupCancellation DEFAULT_UTILITY_LONG_NAME & ": Setup - CodeDb cancelled by user before maintenance started.", callerOwnsLifecycle
        TryRunFullSetup = True
        Exit Function
    End If
    CoreSetSetupPreflightContext True
    If callerOwnsLifecycle = False Then PrepareSetupTopLevelOperation DEFAULT_UTILITY_SHORT_NAME & ": Run Full ADC Setup", "Setup", traceRunOwnedByProcedure
    BeginSetupProgress DEFAULT_UTILITY_SHORT_NAME & ": Setup", ImplementedSetupPartCount(), "Setup"
    setupProcessStartTimer = mCurrentProcessStartTimer
    CorePulseProgress "Resolving Module Guard state"
    LogTraceMessage TRACE_LEVEL_INFO, "ModuleUpdateGuard", "TryRunFullSetup", "Module Guard preflight state resolution starting.", "Preflight distinguishes not installed, installed/enabled, and installed/disabled before A1 begins.", "Setup Preflight", 0, vbNullString, vbNullString, vbNullString, "ModuleGuard", "CodeDb", "Setup"
    moduleGuardResolved = TryResolveModuleUpdateGuardForSetup(failedErrNumber, failedErrSource, failedErrDescription, guardVerificationDetails)
    If moduleGuardResolved Then
        LogModuleGuardSetupTransition "TryRunFullSetup", "Module Guard state resolution returned to Setup.", "Result: success. Next stage: unfinished Setup recovery-state validation."
    Else
        LogModuleGuardSetupTransition "TryRunFullSetup", "Module Guard state resolution returned to Setup without success.", "Err.Number: " & CStr(failedErrNumber) & vbCrLf & "Err.Source: " & failedErrSource & vbCrLf & "Err.Description: " & failedErrDescription, TRACE_LEVEL_WARNING
    End If
    If moduleGuardResolved = False Then
        If IsCancellationError(failedErrNumber) Then
            RecordSetupCancellation failedErrDescription, callerOwnsLifecycle
            EndSetupProgress False
            CoreSetSetupPreflightContext False
            If traceRunOwnedByProcedure Then CompleteTraceRun DEFAULT_UTILITY_LONG_NAME & ": Setup - CodeDb cancelled during Setup Preflight.", False
            TryRunFullSetup = True
            Exit Function
        End If
        mLastFailedProgressPartPrefix = "Setup Preflight > "
        mLastFailedProgressText = mLastFailedProgressPartPrefix & "Module Guard state resolution failed"
        CoreFailSetupProgress
        CoreSetSetupPreflightContext False
        TryRunFullSetup = False
        Exit Function
    End If
    CorePulseProgress "Module Guard state resolved"
    LogSetupRecoveryPreflightStage "Setup recovery-state validation call starting.", "Caller: TryRunFullSetup"
    RequireNoUnfinishedSetupRecoveryState DEFAULT_SETUP_FULL_OPERATION_NAME
    LogSetupRecoveryPreflightStage "Setup recovery-state validation returned to Setup.", "Result: success. Next stage: establish the A2 backup decision in unnumbered Preflight."
    PrepareSetupDataTableBackupForPreflight infrastructurePreflightDetails
    LogTraceMessage TRACE_LEVEL_INFO, "SetupPreflight", "TryRunFullSetup", "Setup infrastructure ownership preflight completed.", infrastructurePreflightDetails, "Setup Preflight infrastructure", 0, vbNullString, vbNullString, vbNullString, "SetupInfrastructure", "CodeDb", "Setup"
    LogSetupStartupStage "TryRunFullSetup", "Setup Control startup call starting.", "Initial status: Setup Preflight is complete. A1 is preparing to run."
    BeginSetupControl "Setup Preflight is complete. A1 is preparing to run."
    LogSetupStartupStage "TryRunFullSetup", "Setup Control startup call completed.", "BeginSetupControl returned successfully."
    setupPopupMessagesEnabled = ShouldShowSetupPopupMessages(showMessage)
    RunSetupThroughPhaseJ setupPopupMessagesEnabled, guardVerificationDetails
    If mSetupControlActive Then
        EndSetupControl
        DoEvents
    End If
    TryRunFullSetup = True
    If mSetupCancelled Then
        If callerOwnsLifecycle Then MarkSystemCommandCancelled
        If traceRunOwnedByProcedure Then CompleteTraceRun DEFAULT_UTILITY_LONG_NAME & ": Setup - CodeDb cancelled.", False
        Exit Function
    End If
    If mSetupCancelled = False And showMessage And setupPopupMessagesEnabled = False Then ShowSetupSuccessStatusMessage FormattedElapsedSeconds(ElapsedSecondsFromTimer(setupProcessStartTimer))
    If traceRunOwnedByProcedure Then CompleteTraceRun DEFAULT_UTILITY_LONG_NAME & ": Setup Preflight + Phases A-J completed."
    Exit Function
ErrHandler:
    failedErrNumber = Err.Number
    failedErrSource = Err.Source
    failedErrDescription = Err.Description
    CaptureFailedSetupPhaseContext
    On Error Resume Next
    CoreSetSetupPreflightContext False
    If mSetupControlActive Then
        EndSetupControl
    End If
    On Error GoTo 0
    If mSetupProgressActive Then CoreFailSetupProgress
    TryRunFullSetup = False
End Function

Private Function ConfirmRunFullSetupImmediately() As VbMsgBoxResult
    'Shows the full-setup confirmation before settings reads, object scans, trace-detail writes, or other preparatory work can make Access appear hung.
    On Error GoTo ErrHandler
    ConfirmRunFullSetupImmediately = UtilityMsgBox(RunFullSetupWarningText(), vbExclamation Or vbYesNo Or vbDefaultButton2, DEFAULT_UTILITY_SHORT_NAME & ": Run Full ADC Setup")
    LogPopupResult ConfirmRunFullSetupImmediately, DEFAULT_UTILITY_SHORT_NAME & ": Run Full ADC Setup", "ConfirmRunFullSetupImmediately"
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ConfirmRunFullSetupImmediately", "Could not show the immediate " & DEFAULT_UTILITY_LONG_NAME & " full-setup confirmation prompt." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Sub ShowSetupSuccessStatusMessage(ByVal totalElapsedText As String)
    Dim messageText As String
    messageText = DEFAULT_UTILITY_LONG_NAME & " Setup Preflight and Phases A-J completed successfully." & vbCrLf & vbCrLf & _
                  "Completed: A1-A4, B1-B5, C1-C3, D1-D4, E1-E3, F1, G1, H1-H3, I1-I3, and J1." & vbCrLf & _
                  "J1 verified the committed final state and persisted the final setup checkpoint." & vbCrLf & vbCrLf & _
                  "Total elapsed time: " & totalElapsedText
    ReportCoreActionSuccess "ShowSetupSuccessStatusMessage", "Setup", "Setup Preflight + Phases A-J completed successfully.", messageText, "Show Phase J setup success status", ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config"), "Table", "CodeDb", "Setup", True, DEFAULT_UTILITY_LONG_NAME & ": Setup"
End Sub

Private Sub PrintSetupFailureDiagnostics(ByVal ProcedureName As String, ByVal errNumber As Long, ByVal errSource As String, ByVal errDescription As String)
    On Error Resume Next
    Dim messageText As String
    messageText = SetupErrorMessage(ProcedureName, errNumber, errSource, errDescription)
    LogTraceMessage TRACE_LEVEL_ERROR, "Core", ProcedureName, SETUP_RESULT_FAILURE_TRACE_NOTE, messageText, vbNullString, errNumber, errSource, errDescription, vbNullString, vbNullString, "CodeDb", "Setup"
End Sub

Private Sub BeginSetupOperation(ByVal operationName As String)
    operationName = Trim$(operationName)
    If mSetupRollbackActive Then
        If mSetupRollbackResumeFirstPartPending And StrComp(mSetupRollbackOperationName, operationName, vbBinaryCompare) = 0 Then
            mSetupRollbackOwnerTraceRunID = GetActiveTraceRunID()
            mSetupPartRollbackStartIndex = 1
            PersistDurableRecoveryMarker SETUP_RECOVERY_KEY, operationName, SETUP_ROLLBACK_ACTIVE_DETAIL, SerializeSetupRollbackEntries()
            CoreAddOperationNote "Interrupted setup/update rollback tracking was adopted by the matching full Setup rerun. Operation: " & operationName & "; preserved rollback entries: " & CStr(mSetupRollbackEntries.count) & "; new owner trace run ID: " & IIf(Len(mSetupRollbackOwnerTraceRunID) = 0, "(pending trace start)", mSetupRollbackOwnerTraceRunID)
            Exit Sub
        End If
        RequireNoUnfinishedSetupRollback operationName
    End If
    mSetupRollbackActive = True
    Set mSetupRollbackEntries = New Collection
    mSetupPartRollbackStartIndex = 1
    mSetupRollbackOperationName = operationName
    mSetupRollbackOwnerTraceRunID = GetActiveTraceRunID()
    mSetupRollbackResumeFirstPartPending = False
    PersistDurableRecoveryMarker SETUP_RECOVERY_KEY, operationName, SETUP_ROLLBACK_ACTIVE_DETAIL, SerializeSetupRollbackEntries()
    CoreAddOperationNote "Setup/update rollback was started. Operation: " & operationName & "; owner trace run ID: " & IIf(Len(mSetupRollbackOwnerTraceRunID) = 0, "(pending trace start)", mSetupRollbackOwnerTraceRunID)
End Sub

Private Sub CommitSetupOperation()
    Dim cleanupWarnings As String
    ValidateSetupRollbackOwnership "CommitSetupOperation"
    DeleteDurableRecoveryMarker SETUP_RECOVERY_KEY
    cleanupWarnings = CleanupSetupRollbackBackupFilesForEntries(mSetupRollbackEntries)
    If Len(cleanupWarnings) > 0 Then CoreAddOperationNote cleanupWarnings
    mSetupRollbackActive = False
    Set mSetupRollbackEntries = Nothing
    mSetupPartRollbackStartIndex = 0
    mSetupRollbackOwnerTraceRunID = vbNullString
    mSetupRollbackOperationName = vbNullString
    mSetupRollbackResumeFirstPartPending = False
    CoreAddOperationNote "Setup/update rollback tracking was committed."
End Sub

' Purpose: Records setup query rollback, coordinating `ValidateSetupRollbackOwnership` and `SerializeSetupRollbackEntries` as the implementation requires.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   queryName (Input; String): Physical Access query name identifying the saved query the routine creates, validates, reads, or updates.
'   targetExisted (Input; Boolean): Records whether the target/row existed before the change so rollback can restore the prior object/row rather than treating it as newly created.
'   previousSql (Input; String): SQL text or SQL-derived value used to define, validate, compare, or execute the previous SQL operation.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub TrackSetupQueryRollback(ByVal queryName As String, ByVal targetExisted As Boolean, ByVal previousSql As String)
    'Intended use: record complete prior query SQL without delimiter serialization so setup rollback can restore it exactly.
    Dim entryData As Variant
    If mSetupRollbackActive = False Then Exit Sub
    ValidateSetupRollbackOwnership "TrackSetupQueryRollback"
    If mSetupRollbackEntries Is Nothing Then Set mSetupRollbackEntries = New Collection
    entryData = Array(SETUP_ROLLBACK_RESTORE_QUERY_ACTION, queryName, targetExisted, previousSql)
    mSetupRollbackEntries.Add entryData
    PersistDurableRecoveryMarker SETUP_RECOVERY_KEY, mSetupRollbackOperationName, SETUP_ROLLBACK_ACTIVE_DETAIL, SerializeSetupRollbackEntries()
End Sub

' Purpose: Records setup created table rollback, coordinating `ValidateSetupRollbackOwnership` and `SerializeSetupRollbackEntries` as the implementation requires.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   tableName (Input; String): Physical Access table name identifying the table whose schema, data, or metadata the routine processes.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub TrackSetupCreatedTableRollback(ByVal tableName As String)
    'Intended use: record a setup-created table so a failed setup can remove or report it during rollback.
    Dim entryData As Variant
    If mSetupRollbackActive = False Then Exit Sub
    ValidateSetupRollbackOwnership "TrackSetupCreatedTableRollback"
    If mSetupRollbackEntries Is Nothing Then Set mSetupRollbackEntries = New Collection
    entryData = Array(SETUP_ROLLBACK_DELETE_TABLE_ACTION, tableName)
    mSetupRollbackEntries.Add entryData
    PersistDurableRecoveryMarker SETUP_RECOVERY_KEY, mSetupRollbackOperationName, SETUP_ROLLBACK_ACTIVE_DETAIL, SerializeSetupRollbackEntries()
End Sub

Private Function ShouldShowSetupPopupMessages(Optional ByVal requestedShowMessage As Boolean = True) As Boolean
    On Error GoTo ErrHandler
    If requestedShowMessage = False Then Exit Function
    ShouldShowSetupPopupMessages = ExistingSettingBoolWithoutRepair(SETTING_GENERATE_SETUP_POPUP_MESSAGES, DEFAULT_GENERATE_SETUP_POPUP_MESSAGES)
    Exit Function
ErrHandler:
    ShouldShowSetupPopupMessages = DEFAULT_GENERATE_SETUP_POPUP_MESSAGES
End Function

Private Function RunFullSetupWarningText() As String
    'Builds the canonical static Setup confirmation without scanning macros or reading settings so the prompt can appear immediately.
    RunFullSetupWarningText = "Setup will verify and reconcile all ADC-managed components in this database, creating or repairing required tables, queries, macros, generated forms, reports, descriptions, Ribbon configuration, Navigation Pane organization, and reference/Add-In inventory as needed." & vbCrLf & vbCrLf & "Continue?"
End Function

' Purpose: Records setup access object rollback, coordinating `ValidateSetupRollbackOwnership` and `StrComp` as the implementation requires.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   actionText (Input; String): Input action text for track setup access object rollback; is tested by If StrComp(actionText, SETUP_ROLLBACK_RESTORE_MACRO_ACTION, vbBinaryCompare) <> 0 And StrCom...; is passed to
'     `Array`.
'   objectName (Input; String): Physical Access object name identifying the database object targeted by the operation.
'   targetExisted (Input; Boolean): Records whether the target/row existed before the change so rollback can restore the prior object/row rather than treating it as newly created.
'   backupPath (Input; String): Filesystem or database path identifying the backup path location used by the operation.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub TrackSetupAccessObjectRollback(ByVal actionText As String, ByVal objectName As String, ByVal targetExisted As Boolean, ByVal backupPath As String)
    Dim entryData As Variant
    If mSetupRollbackActive = False Then Exit Sub
    ValidateSetupRollbackOwnership "TrackSetupAccessObjectRollback"
    If StrComp(actionText, SETUP_ROLLBACK_RESTORE_MACRO_ACTION, vbBinaryCompare) <> 0 And StrComp(actionText, SETUP_ROLLBACK_RESTORE_REPORT_ACTION, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 6438, "TrackSetupAccessObjectRollback", "Only Setup-owned physical Macro and Report replacement rollback entries are supported by this tracker."
    If Len(Trim$(objectName)) = 0 Then Err.Raise vbObjectError + 6439, "TrackSetupAccessObjectRollback", "A tracked Access-object rollback entry requires an object name."
    If targetExisted And Len(Trim$(backupPath)) = 0 Then Err.Raise vbObjectError + 6440, "TrackSetupAccessObjectRollback", "An existing tracked Access object requires a rollback backup path."
    If mSetupRollbackEntries Is Nothing Then Set mSetupRollbackEntries = New Collection
    entryData = Array(actionText, objectName, targetExisted, backupPath)
    mSetupRollbackEntries.Add entryData
    PersistDurableRecoveryMarker SETUP_RECOVERY_KEY, mSetupRollbackOperationName, SETUP_ROLLBACK_ACTIVE_DETAIL, SerializeSetupRollbackEntries()
End Sub
' Purpose: Checks whether the currently active Setup phase owns and permits a2 physical schema writes under the phase-boundary enforcement rules.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: Boolean: True when phase c3 matches the exact state required by the current Setup contract; False when the owning phase still needs repair, refresh, or reconciliation.
Public Function A2PhysicalSchemaWritesAllowedForCurrentSetupPhase() As Boolean
    'A2 is the sole numbered-Setup owner for core ADC table schema mutation. Runtime/manual maintenance remains enabled outside an active Setup run.
    If mSetupRunContextStarted = False Then
        A2PhysicalSchemaWritesAllowedForCurrentSetupPhase = True
        Exit Function
    End If
    A2PhysicalSchemaWritesAllowedForCurrentSetupPhase = (mSetupPhaseIdentity = "A2")
End Function

' Purpose: Checks whether the currently active Setup phase owns and permits B1 catalog structure writes under the phase-boundary enforcement rules.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: Boolean: True when the active Setup phase owns and permits B1 catalog structure writes; False when another phase owns that mutation or the write is otherwise prohibited.
Public Function B1CatalogStructureWritesAllowedForCurrentSetupPhase() As Boolean
    'Compatibility API: delegates B1 ownership to the central semantic-domain map so callers do not maintain a second phase-ownership rule.
    B1CatalogStructureWritesAllowedForCurrentSetupPhase = SetupDomainMutationWritesAllowedForCurrentPart(SETUP_DOMAIN_B1_CATALOG_STRUCTURE)
End Function
' Purpose: Performs setup mutation writes allowed for current part by coordinating `End` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   expectedPartIdentity (Input; String): Stable expected part identity used to select or correlate the exact ADC definition/state handled by this procedure.
' Returns: Boolean result of setup mutation writes allowed for current part; True and False correspond to the affirmative and negative branches explicitly implemented by this function.
'   not-applicable path.
Public Function SetupMutationWritesAllowedForCurrentPart(ByVal expectedPartIdentity As String) As Boolean
    'Fail-closed numbered-Setup mutation guard retained for existing exact-Part callers. New mutation choke points should identify their semantic domain and use RequireSetupDomainMutationAllowed instead.
    expectedPartIdentity = UCase$(Trim$(expectedPartIdentity))
    If Len(expectedPartIdentity) = 0 Then Exit Function
    If mSetupRunContextStarted = False Then
        SetupMutationWritesAllowedForCurrentPart = True
        Exit Function
    End If
    SetupMutationWritesAllowedForCurrentPart = (mSetupPhaseIdentity = expectedPartIdentity)
End Function

Public Function SetupDomainOwnerPart(ByVal domainIdentity As String) As String
    'Single authority for numbered-Setup mutation ownership. Shared helpers may inspect broadly, but their write path names a domain here rather than independently deciding which Part owns it.
    Select Case UCase$(Trim$(domainIdentity))
        Case UCase$(SETUP_DOMAIN_A2_PHYSICAL_SCHEMA): SetupDomainOwnerPart = "A2"
        Case UCase$(SETUP_DOMAIN_A3_SAVED_QUERY_FOUNDATION): SetupDomainOwnerPart = "A3"
        Case UCase$(SETUP_DOMAIN_A4_REFERENCE_INVENTORY): SetupDomainOwnerPart = "A4"
        Case UCase$(SETUP_DOMAIN_B1_CATALOG_STRUCTURE): SetupDomainOwnerPart = "B1"
        Case UCase$(SETUP_DOMAIN_B2_CATALOG_REFERENCE_DEFAULTS): SetupDomainOwnerPart = "B2"
        Case UCase$(SETUP_DOMAIN_B3_NONFORM_MANAGED_DEFINITIONS): SetupDomainOwnerPart = "B3"
        Case UCase$(SETUP_DOMAIN_B4_MACRO_MANAGED_DEFINITIONS): SetupDomainOwnerPart = "B4"
        Case UCase$(SETUP_DOMAIN_B5_CATALOG_SUPPORT): SetupDomainOwnerPart = "B5"
        Case UCase$(SETUP_DOMAIN_C2_PHYSICAL_MACRO_BUILD): SetupDomainOwnerPart = "C2"
        Case UCase$(SETUP_DOMAIN_C3_PHYSICAL_MACRO_FINALIZATION): SetupDomainOwnerPart = "C3"
        Case UCase$(SETUP_DOMAIN_D1_PERSISTENT_LAUNCHER_FORMS): SetupDomainOwnerPart = "D1"
        Case UCase$(SETUP_DOMAIN_D2_CHATGPT_UI_FORMS): SetupDomainOwnerPart = "D2"
        Case UCase$(SETUP_DOMAIN_D3_SUPPORT_FORMS): SetupDomainOwnerPart = "D3"
        Case UCase$(SETUP_DOMAIN_D4_MANAGED_FORM_LOGICAL_STATE), UCase$(SETUP_DOMAIN_D4_LOOKUP_EDITOR_BINDING), UCase$(SETUP_DOMAIN_D4_MANAGED_FORM_PROOF), UCase$(SETUP_DOMAIN_D4_COMPLETE_CATALOG_FINALIZATION): SetupDomainOwnerPart = "D4"
        Case UCase$(SETUP_DOMAIN_E2_PHYSICAL_REPORTS): SetupDomainOwnerPart = "E2"
        Case UCase$(SETUP_DOMAIN_E3_REPORT_FINALIZATION): SetupDomainOwnerPart = "E3"
        Case UCase$(SETUP_DOMAIN_F1_MANAGED_DEPENDENCIES), UCase$(SETUP_DOMAIN_F1_DERIVED_CATALOG_REFRESH): SetupDomainOwnerPart = "F1"
        Case UCase$(SETUP_DOMAIN_G1_DESCRIPTIONS), UCase$(SETUP_DOMAIN_G1_DERIVED_CATALOG_REFRESH): SetupDomainOwnerPart = "G1"
        Case UCase$(SETUP_DOMAIN_H2_RIBBON_BUILD): SetupDomainOwnerPart = "H2"
        Case UCase$(SETUP_DOMAIN_H3_RIBBON_COMMIT): SetupDomainOwnerPart = "H3"
        Case UCase$(SETUP_DOMAIN_I1_NAVIGATION_GROUPS): SetupDomainOwnerPart = "I1"
        Case UCase$(SETUP_DOMAIN_I3_NAVIGATION_PLACEMENT): SetupDomainOwnerPart = "I3"
        Case UCase$(SETUP_DOMAIN_J1_FINAL_STATE): SetupDomainOwnerPart = "J1"
    End Select
End Function

Public Function SetupDomainMutationWritesAllowedForCurrentPart(ByVal domainIdentity As String) As Boolean
    Dim ownerPartIdentity As String
    ownerPartIdentity = SetupDomainOwnerPart(domainIdentity)
    If Len(ownerPartIdentity) = 0 Then Exit Function
    If mSetupRunContextStarted = False Then
        SetupDomainMutationWritesAllowedForCurrentPart = True
        Exit Function
    End If
    SetupDomainMutationWritesAllowedForCurrentPart = (mSetupPhaseIdentity = ownerPartIdentity)
End Function

Public Sub RequireSetupDomainMutationAllowed(ByVal domainIdentity As String, ByVal sourceProcedure As String)
    Dim ownerPartIdentity As String
    Dim currentPartIdentity As String
    ownerPartIdentity = SetupDomainOwnerPart(domainIdentity)
    If Len(ownerPartIdentity) = 0 Then Err.Raise vbObjectError + 10150, "RequireSetupDomainMutationAllowed", "Unknown Setup mutation domain: " & IIf(Len(Trim$(domainIdentity)) = 0, "(blank)", domainIdentity) & "."
    If mSetupRunContextStarted = False Then Exit Sub
    currentPartIdentity = mSetupPhaseIdentity
    If StrComp(currentPartIdentity, ownerPartIdentity, vbBinaryCompare) = 0 Then Exit Sub
    Err.Raise vbObjectError + 10151, sourceProcedure, "Setup phase ownership violation." & vbCrLf & "Current Part: " & IIf(Len(currentPartIdentity) = 0, "(none)", currentPartIdentity) & vbCrLf & "Attempted mutation domain: " & domainIdentity & vbCrLf & "Domain owner: " & ownerPartIdentity & vbCrLf & "Mutation refused."
End Sub

' Purpose: Checks whether phase c3 already satisfies the exact state required by the current Setup run so the owning phase can determine whether repair is necessary.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   detailText (Input/output; String): Receives or supplies diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
' Returns: Boolean: True when the condition represented by PhaseCVerifyPhysicalSystemHelpMacroState is present under the routine's documented inputs/current state; False when that condition is absent.
Public Function PhaseC3CurrentForSetup(ByRef detailText As String) As Boolean
    'Independent read-only successor proof: authoritative B4 source, current physical macros, and persisted C3 proof must still match exactly.
    'J1 and standalone successor checks use this fresh proof; a fresh C1 snapshot is constructed only when persisted proof exists.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim comparisonDetail As String
    Dim expectedToken As String
    Dim actualToken As String
    Dim storedExpectedToken As String
    Dim storedActualToken As String
    Dim storedCountText As String
    Dim propertyExists As Boolean
    Dim failureText As String
    PartCClearComparisonPlan
    Set db = currentDb
    storedExpectedToken = DatabaseTextPropertyValue(db, SYSTEM_MACRO_SET_FINGERPRINT_PROPERTY, propertyExists)
    If propertyExists = False Or Len(storedExpectedToken) = 0 Then detailText = "Persisted C3 expected-state fingerprint is missing.": GoTo NotCurrent
    storedCountText = DatabaseTextPropertyValue(db, SYSTEM_MACRO_EXPECTED_COUNT_PROPERTY, propertyExists)
    If propertyExists = False Or Len(storedCountText) = 0 Then detailText = "Persisted C3 expected macro count is missing.": GoTo NotCurrent
    storedActualToken = DatabaseTextPropertyValue(db, SYSTEM_MACRO_INSTALLED_INVENTORY_TOKEN_PROPERTY, propertyExists)
    If propertyExists = False Or Len(storedActualToken) = 0 Then detailText = "Persisted C3 installed-inventory token is missing.": GoTo NotCurrent
    PartCBuildComparisonPlan comparisonDetail, CoreSetupHelpDocumentSynchronizationScope()
    expectedToken = mPartCExpectedStateToken
    actualToken = mPartCLastPhysicalStateToken
    If mPartCMissingCount <> 0 Or mPartCReplacementCount <> 0 Or mPartCStaleOwnedCount <> 0 Then failureText = SharedText_AppendNonBlankLine(failureText, "Current physical macro state is not exact. Missing=" & CStr(mPartCMissingCount) & "; replacement=" & CStr(mPartCReplacementCount) & "; stale-owned=" & CStr(mPartCStaleOwnedCount) & ".")
    If Len(expectedToken) = 0 Or Len(actualToken) = 0 Or StrComp(expectedToken, actualToken, vbBinaryCompare) <> 0 Then failureText = SharedText_AppendNonBlankLine(failureText, "Current physical macro token does not match the authoritative expected token.")
    If StrComp(storedExpectedToken, expectedToken, vbBinaryCompare) <> 0 Then failureText = SharedText_AppendNonBlankLine(failureText, "Persisted C3 expected-state fingerprint is stale.")
    If Val(storedCountText) <> mPartCExpectedCount Then failureText = SharedText_AppendNonBlankLine(failureText, "Persisted C3 expected macro count is stale.")
    If StrComp(storedActualToken, actualToken, vbBinaryCompare) <> 0 Then failureText = SharedText_AppendNonBlankLine(failureText, "Persisted C3 installed-inventory token is stale.")
    detailText = "C3 current-state proof: expected=" & CStr(mPartCExpectedCount) & "; matching=" & CStr(mPartCMatchingCount) & "; missing=" & CStr(mPartCMissingCount) & "; replacement=" & CStr(mPartCReplacementCount) & "; stale-owned=" & CStr(mPartCStaleOwnedCount) & "."
    If Len(failureText) > 0 Then detailText = detailText & vbCrLf & failureText
    PhaseC3CurrentForSetup = (Len(failureText) = 0)
    PartCClearComparisonPlan
    Exit Function
NotCurrent:
    PartCClearComparisonPlan
    PhaseC3CurrentForSetup = False
    Exit Function
ErrHandler:
    detailText = "C3 current-state proof failed. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    PartCClearComparisonPlan
    PhaseC3CurrentForSetup = False
End Function

Private Sub CoreResetSetupPhaseContext()
    On Error Resume Next
    EndSystemMacroDefinitionSnapshot
    EndManagedObjectDefinitionSnapshot
    On Error GoTo 0
    mSetupPreflightContextActive = False
    mSetupPhaseIdentity = vbNullString
    mSetupPhaseLabel = vbNullString
    mSetupPhasePartNumber = 0
    mSetupPhasePartCount = 0
    mSetupCurrentStep = vbNullString
End Sub

Private Function StatusTextHasSetupPhaseContext(ByVal statusText As String) As Boolean
    Dim normalizedText As String
    normalizedText = LCase$(statusText)
    If Left$(normalizedText, Len("setup preflight >")) = "setup preflight >" Then
        StatusTextHasSetupPhaseContext = True
        Exit Function
    End If
    If Left$(normalizedText, Len("setup ")) = "setup " Then
        If InStr(1, normalizedText, " of ", vbBinaryCompare) > 0 And InStr(1, normalizedText, " > ", vbBinaryCompare) > 0 Then StatusTextHasSetupPhaseContext = True
    End If
End Function

' ===============================================================================================================================================================================
' PHASE C PHYSICAL MACRO LIFECYCLE PROCEDURES
' ===============================================================================================================================================================================

' Purpose: Resolves part C extract run code argument from the supplied/current ADC state, using `End` and `ElseIf` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   macroText (Input; String): Input macro text for part C extract run code argument; is passed to `InStr`; is passed to `Len`.
' Returns: String: Returns the part c extract run code argument text value. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function PartCExtractRunCodeArgument(ByVal macroText As String) As String
    'Extracts the RunCode argument text from serialized Access macro text.
    Dim actionPosition As Long
    Dim argumentPosition As Long
    Dim valueStart As Long
    Dim valuePosition As Long
    Dim currentCharacter As String
    Dim resultText As String
    actionPosition = InStr(1, macroText, "Action =" & Chr$(34) & "RunCode" & Chr$(34), vbTextCompare)
    If actionPosition = 0 Then Exit Function
    argumentPosition = InStr(actionPosition, macroText, "Argument =" & Chr$(34), vbTextCompare)
    If argumentPosition = 0 Then Exit Function
    valueStart = argumentPosition + Len("Argument =" & Chr$(34))
    valuePosition = valueStart
    Do While valuePosition <= Len(macroText)
        currentCharacter = Mid$(macroText, valuePosition, 1)
        If currentCharacter = Chr$(92) Then
            If valuePosition < Len(macroText) And Mid$(macroText, valuePosition + 1, 1) = Chr$(34) Then
                resultText = resultText & Chr$(34)
                valuePosition = valuePosition + 2
            Else
                resultText = resultText & currentCharacter
                valuePosition = valuePosition + 1
            End If
        ElseIf currentCharacter = Chr$(34) Then
            If valuePosition < Len(macroText) And Mid$(macroText, valuePosition + 1, 1) = Chr$(34) Then
                resultText = resultText & Chr$(34)
                valuePosition = valuePosition + 2
            Else
                PartCExtractRunCodeArgument = resultText
                Exit Function
            End If
        Else
            resultText = resultText & currentCharacter
            valuePosition = valuePosition + 1
        End If
    Loop
End Function

' Purpose: Performs phase C verify physical system help macro state by coordinating `PartCBuildComparisonPlan` and `End` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   detailText (Output; String): Receives the full C1 physical System/Help Macro comparison-plan summary produced by PartCBuildComparisonPlan before the phase completes.
'   synchronizationScope (Input Optional; HelpDocumentSynchronizationScope): Scope value limiting which synchronization items or operations are included in this call. Default:
'   HelpDocumentScopeSetupSystemAndAllActive.
'   helpDocumentMutationsOnly (Input Optional; Boolean): Boolean state describing whether help document mutations only applies to this invocation; True selects the named condition/path and False selects its
'   complementary/default path. Default: False.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub PhaseCVerifyPhysicalSystemHelpMacroState(ByRef detailText As String, Optional ByVal synchronizationScope As HelpDocumentSynchronizationScope = HelpDocumentScopeSetupSystemAndAllActive, Optional ByVal helpDocumentMutationsOnly As Boolean = False)
    On Error GoTo ErrHandler
    PartCBuildComparisonPlan detailText, synchronizationScope, helpDocumentMutationsOnly
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    PartCClearComparisonPlan
    Err.Raise savedErrNumber, "PhaseCVerifyPhysicalSystemHelpMacroState", "C1 could not verify the physical System Macro and Help Document Macro state without mutation." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

' Purpose: Performs the public phase c build required system help macros operation using the routine's existing ADC workflow and side effects.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   detailText (Output/in-out; String): Carries human-readable diagnostic detail for the operation; when passed ByRef, the routine can return or augment the text for the caller.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects and the documented ByRef parameter(s): detailText.
Public Sub PhaseCBuildRequiredSystemHelpMacros(ByRef detailText As String)
    RequireSetupDomainMutationAllowed SETUP_DOMAIN_C2_PHYSICAL_MACRO_BUILD, "PhaseCBuildRequiredSystemHelpMacros"
    On Error GoTo ErrHandler
    Dim item As Variant
    Dim processedCount As Long
    Dim totalCount As Long
    Dim phaseTimer As Single
    Dim payloadReadsAtStart As Long
    Dim backupsAtStart As Long
    Dim staleRereadsAtStart As Long
    phaseTimer = Timer
    payloadReadsAtStart = mPartCPayloadReadCount
    backupsAtStart = mPartCRollbackBackupCount
    staleRereadsAtStart = mPartCTargetedStaleRereadCount
    If mPartCPlanReady = False Then Err.Raise vbObjectError + 6520, "PhaseCBuildRequiredSystemHelpMacros", "C1 must complete successfully before C2 can mutate physical macros."
    PartCValidatePlanSourceCurrent "C2"
    PartCValidateBuildPlanTargetsCurrentForC2
    totalCount = mPartCBuildPlan.count
    For Each item In mPartCBuildPlan
        processedCount = processedCount + 1
        CoreUpdateSetupProgress "Build required System Macro and Help Document Macro " & CStr(processedCount) & " of " & CStr(totalCount) & ": " & CStr(item(0))
        PartCCreateOrReplaceMacro CStr(item(0)), CStr(item(1))
        Select Case CStr(item(2))
            Case "Missing"
                If mPartCMissingCount > 0 Then mPartCMissingCount = mPartCMissingCount - 1
                mPartCCreatedCount = mPartCCreatedCount + 1
            Case "Stale"
                If mPartCReplacementCount > 0 Then mPartCReplacementCount = mPartCReplacementCount - 1
                mPartCReplacedCount = mPartCReplacedCount + 1
        End Select
        mPartCMatchingCount = mPartCMatchingCount + 1
        detailText = SharedText_AppendNonBlankLine(detailText, " - Built: " & CStr(item(0)) & " (" & CStr(item(2)) & ")")
        RequireSafeWorkContinue "PhaseCBuildRequiredSystemHelpMacros", "Macro [" & CStr(item(0)) & "]"
    Next item
    If totalCount = 0 Then detailText = SharedText_AppendNonBlankLine(detailText, " - C1 produced an empty C2 build plan; no physical macro objects were rewritten.")
    ReconcileStandalonePhaseCSystemMacroDescriptions detailText
    LogTraceMessage TRACE_LEVEL_INFO, "PhysicalMacros", "PhaseCBuildRequiredSystemHelpMacros", "C2 physical macro mutation plan completed.", _
        "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseTimer)) & vbCrLf & _
        "Payload SaveAsText reads: " & CStr(mPartCPayloadReadCount - payloadReadsAtStart) & vbCrLf & _
        "Rollback SaveAsText backups: " & CStr(mPartCRollbackBackupCount - backupsAtStart) & vbCrLf & _
        "Expected macros: " & CStr(mPartCExpectedCount) & vbCrLf & _
        "Matching after C2: " & CStr(mPartCMatchingCount) & vbCrLf & _
        "Missing after C2: " & CStr(mPartCMissingCount) & vbCrLf & _
        "Created: " & CStr(mPartCCreatedCount) & vbCrLf & _
        "Replaced: " & CStr(mPartCReplacedCount) & vbCrLf & _
        "Deleted: " & CStr(mPartCDeletedCount) & vbCrLf & _
        "Targeted stale-plan payload rereads: " & CStr(mPartCTargetedStaleRereadCount - staleRereadsAtStart) & vbCrLf & _
        "Whole-set second payload pass occurred: " & CStr(mPartCWholeSetSecondPayloadPassOccurred), _
        "C2 physical macros", 0, vbNullString, vbNullString, vbNullString, "SystemMacro", "CodeDb", "Performance"
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    LogTraceMessage TRACE_LEVEL_ERROR, "PhysicalMacros", "PhaseCBuildRequiredSystemHelpMacros", "C2 physical macro mutation plan failed.", _
        "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseTimer)) & vbCrLf & _
        "Payload SaveAsText reads: " & CStr(mPartCPayloadReadCount - payloadReadsAtStart) & vbCrLf & _
        "Rollback SaveAsText backups: " & CStr(mPartCRollbackBackupCount - backupsAtStart) & vbCrLf & _
        "Expected macros: " & CStr(mPartCExpectedCount) & vbCrLf & _
        "Matching: " & CStr(mPartCMatchingCount) & vbCrLf & _
        "Missing: " & CStr(mPartCMissingCount) & vbCrLf & _
        "Created: " & CStr(mPartCCreatedCount) & vbCrLf & _
        "Replaced: " & CStr(mPartCReplacedCount) & vbCrLf & _
        "Deleted: " & CStr(mPartCDeletedCount) & vbCrLf & _
        "Targeted stale-plan payload rereads: " & CStr(mPartCTargetedStaleRereadCount - staleRereadsAtStart) & vbCrLf & _
        "Whole-set second payload pass occurred: " & CStr(mPartCWholeSetSecondPayloadPassOccurred), _
        "C2 physical macros", savedErrNumber, savedErrSource, savedErrDescription, vbNullString, "SystemMacro", "CodeDb", "Performance"
    On Error GoTo 0
    PartCInvalidatePersistedPhysicalMacroState
    PartCClearComparisonPlan
    Err.Raise savedErrNumber, "PhaseCBuildRequiredSystemHelpMacros", "C2 could not complete the exact C1 physical macro change plan." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub ReconcileStandalonePhaseCSystemMacroDescriptions(ByRef detailText As String)
    'C2 stays physical-macro-only during numbered Setup. A direct standalone/general C2 call completes required System Macro Descriptions; the scoped Help-document repair path remains limited to Help objects.
    Dim expectedKeys As Variant
    Dim expectedItem As Variant
    Dim keyText As String
    Dim descriptionChanged As Variant
    Dim changedCount As Long
    Dim checkedCount As Long
    Dim i As Long
    If SetupDomainMutationWritesAllowedForCurrentPart(SETUP_DOMAIN_G1_DESCRIPTIONS) = False Then Exit Sub
    If mPartCHelpDocumentMutationsOnly Then Exit Sub
    If mPartCExpectedMacros Is Nothing Then Err.Raise vbObjectError + 10134, "ReconcileStandalonePhaseCSystemMacroDescriptions", "The authoritative C1 macro snapshot is unavailable for standalone System Macro Description completion."
    expectedKeys = SortedDictionaryKeys(mPartCExpectedMacros)
    For i = LBound(expectedKeys) To UBound(expectedKeys)
        keyText = CStr(expectedKeys(i))
        expectedItem = mPartCExpectedMacros(keyText)
        If InStr(1, CStr(expectedItem(2)), "MC_Sys:", vbTextCompare) = 1 Then
            checkedCount = checkedCount + 1
            descriptionChanged = False
            If ApplyAndVerifyManagedObjectDescriptionForStandaloneRepair(CStr(expectedItem(0)), ADCAT_OBJECT_TYPE_MACRO, descriptionChanged) = False Then Err.Raise vbObjectError + 10135, "ReconcileStandalonePhaseCSystemMacroDescriptions", "System Macro Description reconciliation returned False." & vbCrLf & "Macro: " & CStr(expectedItem(0))
            If CBool(descriptionChanged) Then changedCount = changedCount + 1
        End If
    Next i
    detailText = SharedText_AppendNonBlankLine(detailText, " - Standalone System Macro descriptions reconciled: checked=" & CStr(checkedCount) & "; changed=" & CStr(changedCount) & ".")
End Sub

' Purpose: Performs the public phase c finalize physical system help macro state operation using the routine's existing ADC workflow and side effects.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   detailText (Output/in-out; String): Carries human-readable diagnostic detail for the operation; when passed ByRef, the routine can return or augment the text for the caller.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects and the documented ByRef parameter(s): detailText.
Public Sub PhaseCFinalizePhysicalSystemHelpMacroState(ByRef detailText As String)
    RequireSetupDomainMutationAllowed SETUP_DOMAIN_C3_PHYSICAL_MACRO_FINALIZATION, "PhaseCFinalizePhysicalSystemHelpMacroState"
    On Error GoTo ErrHandler
    Dim item As Variant
    Dim deletedCount As Long
    Dim verifyDetail As String
    Dim actualStateToken As String
    Dim phaseTimer As Single
    Dim payloadReadsAtStart As Long
    Dim backupsAtStart As Long
    Dim staleRereadsAtStart As Long
    phaseTimer = Timer
    payloadReadsAtStart = mPartCPayloadReadCount
    backupsAtStart = mPartCRollbackBackupCount
    staleRereadsAtStart = mPartCTargetedStaleRereadCount
    If mPartCPlanReady = False Then Err.Raise vbObjectError + 6522, "PhaseCFinalizePhysicalSystemHelpMacroState", "C1 must complete successfully before C3 finalization."
    PartCValidatePlanSourceCurrent "C3"
    If mPartCDeletePlan.count > 0 Then PartCValidateDeletePlanTargetsCurrentForC3
    For Each item In mPartCDeletePlan
        PartCDeleteStaleOwnedMacro CStr(item)
        deletedCount = deletedCount + 1
        If mPartCStaleOwnedCount > 0 Then mPartCStaleOwnedCount = mPartCStaleOwnedCount - 1
        detailText = SharedText_AppendNonBlankLine(detailText, " - Deleted stale " & DEFAULT_UTILITY_SHORT_NAME & "-owned macro: " & CStr(item))
        RequireSafeWorkContinue "PhaseCFinalizePhysicalSystemHelpMacroState", "Stale macro [" & CStr(item) & "]"
    Next item
    PartCValidateFinalPhysicalStateFromEvidence actualStateToken, verifyDetail
    If mPartCMissingCount <> 0 Or mPartCReplacementCount <> 0 Or mPartCStaleOwnedCount <> 0 Then Err.Raise vbObjectError + 6524, "PhaseCFinalizePhysicalSystemHelpMacroState", "C3 final verification did not reach an exact physical macro state." & vbCrLf & verifyDetail
    If Len(actualStateToken) = 0 Or StrComp(actualStateToken, mPartCExpectedStateToken, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 6525, "PhaseCFinalizePhysicalSystemHelpMacroState", "C3 physical macro state token did not match the authoritative expected token."
    Call EnsureDaoPropertyValue(currentDb, SYSTEM_MACRO_SET_FINGERPRINT_PROPERTY, dbText, mPartCExpectedStateToken, "PhaseCFinalizePhysicalSystemHelpMacroState")
    Call EnsureDaoPropertyValue(currentDb, SYSTEM_MACRO_EXPECTED_COUNT_PROPERTY, dbText, CStr(mPartCExpectedCount), "PhaseCFinalizePhysicalSystemHelpMacroState")
    Call EnsureDaoPropertyValue(currentDb, SYSTEM_MACRO_INSTALLED_INVENTORY_TOKEN_PROPERTY, dbText, actualStateToken, "PhaseCFinalizePhysicalSystemHelpMacroState")
    detailText = SharedText_AppendNonBlankLine(detailText, " - Final verification matched " & CStr(mPartCExpectedCount) & " authoritative System Macros and Help Document Macros; stale owned macros deleted: " & CStr(deletedCount) & ".")
    LogTraceMessage TRACE_LEVEL_INFO, "PhysicalMacros", "PhaseCFinalizePhysicalSystemHelpMacroState", "C3 physical macro exact-state finalization completed.", _
        "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseTimer)) & vbCrLf & _
        "Payload SaveAsText reads: " & CStr(mPartCPayloadReadCount - payloadReadsAtStart) & vbCrLf & _
        "Rollback SaveAsText backups: " & CStr(mPartCRollbackBackupCount - backupsAtStart) & vbCrLf & _
        "Expected macros: " & CStr(mPartCExpectedCount) & vbCrLf & _
        "Matching: " & CStr(mPartCMatchingCount) & vbCrLf & _
        "Missing: " & CStr(mPartCMissingCount) & vbCrLf & _
        "Created: " & CStr(mPartCCreatedCount) & vbCrLf & _
        "Replaced: " & CStr(mPartCReplacedCount) & vbCrLf & _
        "Deleted: " & CStr(mPartCDeletedCount) & vbCrLf & _
        "Targeted stale-plan payload rereads: " & CStr(mPartCTargetedStaleRereadCount - staleRereadsAtStart) & vbCrLf & _
        "Whole-set second payload pass occurred: " & CStr(mPartCWholeSetSecondPayloadPassOccurred), _
        "C3 physical macros", 0, vbNullString, vbNullString, vbNullString, "SystemMacro", "CodeDb", "Performance"
    PartCClearComparisonPlan
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    LogTraceMessage TRACE_LEVEL_ERROR, "PhysicalMacros", "PhaseCFinalizePhysicalSystemHelpMacroState", "C3 physical macro exact-state finalization failed.", _
        "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseTimer)) & vbCrLf & _
        "Payload SaveAsText reads: " & CStr(mPartCPayloadReadCount - payloadReadsAtStart) & vbCrLf & _
        "Rollback SaveAsText backups: " & CStr(mPartCRollbackBackupCount - backupsAtStart) & vbCrLf & _
        "Expected macros: " & CStr(mPartCExpectedCount) & vbCrLf & _
        "Matching: " & CStr(mPartCMatchingCount) & vbCrLf & _
        "Missing: " & CStr(mPartCMissingCount) & vbCrLf & _
        "Created: " & CStr(mPartCCreatedCount) & vbCrLf & _
        "Replaced: " & CStr(mPartCReplacedCount) & vbCrLf & _
        "Deleted: " & CStr(mPartCDeletedCount) & vbCrLf & _
        "Targeted stale-plan payload rereads: " & CStr(mPartCTargetedStaleRereadCount - staleRereadsAtStart) & vbCrLf & _
        "Whole-set second payload pass occurred: " & CStr(mPartCWholeSetSecondPayloadPassOccurred), _
        "C3 physical macros", savedErrNumber, savedErrSource, savedErrDescription, vbNullString, "SystemMacro", "CodeDb", "Performance"
    On Error GoTo 0
    PartCInvalidatePersistedPhysicalMacroState
    PartCClearComparisonPlan
    Err.Raise savedErrNumber, "PhaseCFinalizePhysicalSystemHelpMacroState", "C3 could not prove and persist the exact physical System Macro and Help Document Macro state." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub PartCAddActiveHelpDocumentDeletionMacros(ByRef expectedMacros As Object, ByRef rowsRead As Long)
    On Error GoTo ErrHandler
    Dim sourceDb As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim sourceKey As String
    Dim rowIdentity As String
    Dim helpDocumentNumber As Long
    Dim MacroTitle As String
    Dim runCodeExpression As String
    rowsRead = 0
    Set sourceDb = ADCDb()
    If SharedDao_TableExists(sourceDb, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HelpDocuments")) = False Then Err.Raise vbObjectError + 6527, "PartCAddActiveHelpDocumentDeletionMacros", "Help_Docs is missing. Phase C deletion does not repair B4."
    sqlText = "SELECT " & SQL_FIELD_ID & ",[" & FIELD_HELP_DOCUMENT_KEY & "],[" & FIELD_HELP_DOCUMENT_NUMBER & "],[" & FIELD_DEFAULT_TITLE_TEXT & "],[" & FIELD_MACRO_TITLE_OVERRIDE & "] FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HelpDocuments") & "] WHERE Nz(" & SQL_FIELD_HELP_DOCUMENT_STATUS & "," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ";"
    Set rs = sourceDb.OpenRecordset(sqlText, dbOpenSnapshot)
    Do While rs.EOF = False
        PumpSetupResponsiveness
        rowsRead = rowsRead + 1
        sourceKey = CStr(Nz(rs.fields(FIELD_HELP_DOCUMENT_KEY).value, vbNullString))
        rowIdentity = CStr(Nz(rs.fields(FIELD_ID).value, vbNullString))
        helpDocumentNumber = CLng(Nz(rs.fields(FIELD_HELP_DOCUMENT_NUMBER).value, 0))
        ResolveHelpDocumentMaterialization sourceKey, helpDocumentNumber, CStr(Nz(rs.fields(FIELD_DEFAULT_TITLE_TEXT).value, vbNullString)), CStr(Nz(rs.fields(FIELD_MACRO_TITLE_OVERRIDE).value, vbNullString)), rowIdentity, HelpDocumentScopeAllActive, MacroTitle, runCodeExpression
        PartCAddExpectedMacroToDictionary expectedMacros, MacroTitle, runCodeExpression, "Help_Docs:" & IIf(Len(Trim$(sourceKey)) = 0, "(blank)", sourceKey)
        rs.MoveNext
    Loop
    rs.Close
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
    Err.Raise savedErrNumber, "PartCAddActiveHelpDocumentDeletionMacros", "Could not read active Help_Docs for physical macro deletion." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

' Purpose: Prepares system help macro deletion plan.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   plannedCountResult (Output Optional; Long): Receives the planned count result produced by the operation so the caller can inspect the outcome without re-running the work. Default: 0.
' Returns: Collection: Returns the prepared deletion-plan object/collection describing the exact managed objects eligible for the requested deletion scope; absence follows the routine's existing no-work
'   representation.
Public Function PrepareSystemHelpMacroDeletionPlan(Optional ByRef plannedCountResult As Long = 0) As Collection
    'Phase C authoritative preparation for one operation-scoped physical ADC System Macro and Help Document Macro deletion plan. The returned Collection is the exact plan later consumed by DeleteSystemHelpMacros.
    On Error GoTo ErrHandler
    Dim expectedMacros As Object
    Dim systemRows As Collection
    Dim deletionNames As Collection
    Dim allMacro As accessObject
    Dim rowData As Variant
    Dim macroName As String
    Dim runCodeExpression As String
    Dim stageTimer As Single
    Dim totalTimer As Single
    Dim systemRowsRead As Long
    Dim helpRowsRead As Long
    Dim physicalMacrosEnumerated As Long
    Dim generatedNameCandidates As Long
    Dim runCodePayloadReads As Long
    Dim namingSnapshotOwned As Boolean
    Dim expectedBuildSeconds As Double
    Dim enumerateSeconds As Double
    totalTimer = Timer
    plannedCountResult = 0
    If CodeDbIsHostDb() = False Then Err.Raise vbObjectError + 6560, "PrepareSystemHelpMacroDeletionPlan", "ADC macro deletion planning requires CodeDb and HostDb to be the same database." & vbCrLf & "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB
    ' INTENTIONAL OPERATION SNAPSHOT: Phase C macro-plan construction resolves the same managed names repeatedly; reuse any caller-owned snapshot, otherwise own one only for this read-only plan build.
    If NamingSnapshotIsReady() = False Then
        BeginNamingSnapshot
        namingSnapshotOwned = True
    End If
    ' INTENTIONAL KEYED LOOKUP: expectedMacros is the operation-owned name membership index built from one narrow MC_Sys read plus one active Help_Docs read.
    Set expectedMacros = CreateObject("Scripting.Dictionary")
    expectedMacros.CompareMode = vbTextCompare
    stageTimer = Timer
    Set systemRows = SystemMacroDeletionMembershipRows(systemRowsRead)
    For Each rowData In systemRows
        PartCAddExpectedMacroToDictionary expectedMacros, CStr(rowData(1)), CStr(rowData(2)), "MC_Sys:" & CStr(rowData(0))
    Next rowData
    LogTraceMessage TRACE_LEVEL_INFO, "Performance", "PrepareSystemHelpMacroDeletionPlan", "Macro 65 active MC_Sys deletion membership loaded.", "Rows read=" & CStr(systemRowsRead) & "; expected names after MC_Sys=" & CStr(expectedMacros.count) & "; elapsed=" & FormattedElapsedSeconds(ElapsedSecondsFromTimer(stageTimer)) & "; System Macro definitions constructed=False; Setup Part roster initialized=False", "Macro 65 macro membership", 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SystemMacros"), "Table", "CodeDb", "SystemMacro65"
    stageTimer = Timer
    PartCAddActiveHelpDocumentDeletionMacros expectedMacros, helpRowsRead
    expectedBuildSeconds = ElapsedSecondsFromTimer(stageTimer)
    LogTraceMessage TRACE_LEVEL_INFO, "Performance", "PrepareSystemHelpMacroDeletionPlan", "Macro 65 active Help_Docs deletion membership loaded and expected-name index completed.", "Help_Docs rows read=" & CStr(helpRowsRead) & "; current expected macros=" & CStr(expectedMacros.count) & "; elapsed=" & FormattedElapsedSeconds(expectedBuildSeconds) & "; HelpDocumentDefinitions constructed=False", "Macro 65 help membership", 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HelpDocuments"), "Table", "CodeDb", "SystemMacro65"
    ' INTENTIONAL OPERATION SNAPSHOT: deletionNames is constructed once by the Phase C owner and is consumed unchanged for the lifetime of the deletion operation.
    Set deletionNames = New Collection
    stageTimer = Timer
    For Each allMacro In CurrentProject.AllMacros
        PumpSetupResponsiveness
        physicalMacrosEnumerated = physicalMacrosEnumerated + 1
        macroName = allMacro.Name
        If expectedMacros.Exists(macroName) Then
            deletionNames.Add macroName
        ElseIf PartCPhysicalMacroNameLooksGeneratedOwned(macroName) Then
            generatedNameCandidates = generatedNameCandidates + 1
            runCodeExpression = vbNullString
            runCodePayloadReads = runCodePayloadReads + 1
            If PartCPhysicalMacroRunCode(macroName, runCodeExpression) Then
                If PartCPhysicalMacroOwnedInSynchronizationScope(runCodeExpression, HelpDocumentScopeAllActive) Then deletionNames.Add macroName
            End If
        End If
    Next allMacro
    enumerateSeconds = ElapsedSecondsFromTimer(stageTimer)
    plannedCountResult = deletionNames.count
    LogTraceMessage TRACE_LEVEL_INFO, "Performance", "PrepareSystemHelpMacroDeletionPlan", "Macro 65 physical macro deletion plan prepared.", "MC_Sys rows read=" & CStr(systemRowsRead) & "; Help_Docs rows read=" & CStr(helpRowsRead) & "; physical macros enumerated=" & CStr(physicalMacrosEnumerated) & "; generated-name candidates=" & CStr(generatedNameCandidates) & "; stale-owned RunCode payload reads=" & CStr(runCodePayloadReads) & "; deletion names retained=" & CStr(deletionNames.count) & "; enumeration elapsed=" & FormattedElapsedSeconds(enumerateSeconds) & "; total plan elapsed=" & FormattedElapsedSeconds(ElapsedSecondsFromTimer(totalTimer)), "Macro 65 macro deletion plan", 0, vbNullString, vbNullString, vbNullString, "MacroSet", "CodeDb", "SystemMacro65"
    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "PrepareSystemHelpMacroDeletionPlan", "DELETE PLAN", "Scope=Macros; Planned=" & CStr(deletionNames.count), "Macro 65 deletion plan", 0, vbNullString, vbNullString, vbNullString, "MacroSet", "CodeDb", "SystemMacro65"
    If deletionNames.count = 0 Then LogTraceMessage TRACE_LEVEL_WARNING, "DeleteObjects", "PrepareSystemHelpMacroDeletionPlan", "DELETE PLAN EMPTY", "Scope=Macros", "Macro 65 deletion plan", 0, vbNullString, vbNullString, vbNullString, "MacroSet", "CodeDb", "SystemMacro65"
    If namingSnapshotOwned Then InvalidateNamingSnapshot
    namingSnapshotOwned = False
    Set PrepareSystemHelpMacroDeletionPlan = deletionNames
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If namingSnapshotOwned Then InvalidateNamingSnapshot
    namingSnapshotOwned = False
    On Error GoTo 0
    Err.Raise savedErrNumber, "PrepareSystemHelpMacroDeletionPlan", "Could not prepare the authoritative ADC System Macro and Help Document Macro deletion plan." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

' Purpose: Deletes system help macros for confirm destructive action, prepared deletion plan.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
'   confirmDestructiveAction (Input Optional; Boolean): When True, enables or requests destructive action for this call; when False, the routine follows its normal non-forced/default path. Default: True.
'   plannedCountResult (Output Optional; Long): Receives the planned count result produced by the operation so the caller can inspect the outcome without re-running the work. Default: 0.
'   attemptedCountResult (Output Optional; Long): Receives the attempted count result produced by the operation so the caller can inspect the outcome without re-running the work. Default: 0.
'   deletedCountResult (Output Optional; Long): Receives the deleted count result produced by the operation so the caller can inspect the outcome without re-running the work. Default: 0.
'   skippedCountResult (Output Optional; Long): Receives the skipped count result produced by the operation so the caller can inspect the outcome without re-running the work. Default: 0.
'   failedCountResult (Output Optional; Long): Receives the failed count result produced by the operation so the caller can inspect the outcome without re-running the work. Default: 0.
'   preparedDeletionPlan (Input Optional; Variant): Previously prepared deletion-plan payload describing the exact objects/actions approved for deletion; omitted means the routine must follow its normal
'   prepare/resolve path.
' Returns: Boolean: True on the function's affirmative/success branch for delete system help macros; False on its explicit negative, unavailable, rejected, or handled-failure branch.
Public Function DeleteSystemHelpMacros(Optional ByVal showMessage As Boolean = True, Optional ByVal confirmDestructiveAction As Boolean = True, Optional ByRef plannedCountResult As Long = 0, Optional ByRef attemptedCountResult As Long = 0, Optional ByRef deletedCountResult As Long = 0, Optional ByRef skippedCountResult As Long = 0, Optional ByRef failedCountResult As Long = 0, Optional ByVal preparedDeletionPlan As Variant) As Boolean
    'Phase C ownership authority for destructive removal of the physical ADC System Macro and Help Document Macro set. It executes exactly one operation-scoped plan prepared by the same owner; no unrelated Access
    'macros are deleted.
    On Error GoTo ErrHandler
    Dim deletionNames As Collection
    Dim macroName As String
    Dim item As Variant
    Dim deletedCount As Long
    Dim attemptedCount As Long
    Dim skippedCount As Long
    Dim failedCount As Long
    Dim response As VbMsgBoxResult
    Dim typedText As String
    Dim messageText As String
    Dim totalTimer As Single
    Dim stageTimer As Single
    Dim deleteSeconds As Double
    Dim invalidateSeconds As Double
    Dim invokingMacroName As String
    Dim deferredInvokingMacro As Boolean
    totalTimer = Timer
    plannedCountResult = 0
    attemptedCountResult = 0
    deletedCountResult = 0
    skippedCountResult = 0
    failedCountResult = 0
    If CodeDbIsHostDb() = False Then Err.Raise vbObjectError + 6560, "DeleteSystemHelpMacros", "ADC macro deletion requires CodeDb and HostDb to be the same database." & vbCrLf & "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB
    If isMissing(preparedDeletionPlan) Then
        Set deletionNames = PrepareSystemHelpMacroDeletionPlan(plannedCountResult)
    ElseIf IsObject(preparedDeletionPlan) Then
        Set deletionNames = preparedDeletionPlan
        If deletionNames Is Nothing Then Err.Raise vbObjectError + 6562, "DeleteSystemHelpMacros", "The supplied Phase C macro deletion plan is Nothing."
        plannedCountResult = deletionNames.count
        LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "DeleteSystemHelpMacros", "Prepared deletion plan reused", "Scope=Macros; Planned=" & CStr(plannedCountResult), "Macro 65 deletion plan reuse", 0, vbNullString, vbNullString, vbNullString, "MacroSet", "CodeDb", "SystemMacro65"
    Else
        Err.Raise vbObjectError + 6563, "DeleteSystemHelpMacros", "The supplied Phase C macro deletion plan is not an object."
    End If
    If confirmDestructiveAction Then
        messageText = "This will delete the physical ADC system macros and active Help Document macros recognized by the Phase C ownership rules." & vbCrLf & vbCrLf & "Unrelated user macros will remain. No backup will be created. This action is destructive. Continue?"
        response = CoreShowMessageResult(messageText, vbCritical Or vbYesNo Or vbDefaultButton2, "Delete All ADC Objects", "ShowADCMessageResult", "Macros")
        If response <> vbYes Then MarkSystemCommandCancelled: Exit Function
        typedText = InputBox("Type DELETE ALL to delete the ADC System Macros and Help Document Macros." & vbCrLf & vbCrLf & "Unrelated user macros will remain. No backup will be created.", "Confirm Delete All ADC Objects")
        If StrComp(Trim$(typedText), "DELETE ALL", vbBinaryCompare) <> 0 Then
            If showMessage Then CoreShowMessage "Delete All ADC Objects was cancelled because the confirmation text did not match DELETE ALL.", vbInformation, "Delete All ADC Objects", "ShowADCMessage", "Macros"
            MarkSystemCommandCancelled
            Exit Function
        End If
    End If
    If mMacro65RunCodeActive Then
        invokingMacroName = Trim$(CoreInvokingMacro65ObjectName())
        If Len(invokingMacroName) = 0 Then Err.Raise vbObjectError + 6580, "DeleteSystemHelpMacros", "The invoking Macro 65 physical name could not be resolved before destructive macro deletion."
    End If
    stageTimer = Timer
    For Each item In deletionNames
        macroName = CStr(item)
        If Len(invokingMacroName) > 0 And StrComp(Trim$(macroName), invokingMacroName, vbTextCompare) = 0 Then
            If RollbackObjectExistsStrict(currentDb, acMacro, macroName, "before deferred Macro 65 deletion") Then
                CoreMarkInvokingMacro65DeletionDeferred macroName, "Normal"
                deferredInvokingMacro = True
                LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "DeleteSystemHelpMacros", "DELETE DEFERRED", "Type=Macro; Name=" & macroName & "; Reason=Invoking Macro 65 RunCode is still active", "Macro 65 terminal delete deferral", 0, vbNullString, vbNullString, macroName, "Macro", "CodeDb", "SystemMacro65"
            Else
                skippedCount = skippedCount + 1
                skippedCountResult = skippedCount
                LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "DeleteSystemHelpMacros", "DELETE SKIP", "Type=Macro; Name=" & macroName & "; Reason=Not found at deletion time", "Macro 65 delete skip", 0, vbNullString, vbNullString, macroName, "Macro", "CodeDb", "SystemMacro65"
            End If
        ElseIf RollbackObjectExistsStrict(currentDb, acMacro, macroName, "before Delete All ADC Objects macro deletion") Then
            attemptedCount = attemptedCount + 1
            attemptedCountResult = attemptedCount
            LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "DeleteSystemHelpMacros", "DELETE ATTEMPT", "Type=Macro; Name=" & macroName, "Macro 65 delete attempt", 0, vbNullString, vbNullString, macroName, "Macro", "CodeDb", "SystemMacro65"
            On Error Resume Next
            DoCmd.Close acMacro, macroName, acSaveNo
            Err.Clear
            On Error GoTo ErrHandler
            DoCmd.DeleteObject acMacro, macroName
            If RollbackObjectExistsStrict(currentDb, acMacro, macroName, "after Delete All ADC Objects macro deletion") Then Err.Raise vbObjectError + 6561, "DeleteSystemHelpMacros", "An ADC-owned physical macro remained after deletion." & vbCrLf & "Macro: " & macroName
            deletedCount = deletedCount + 1
            deletedCountResult = deletedCount
            LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "DeleteSystemHelpMacros", "DELETE SUCCESS", "Type=Macro; Name=" & macroName, "Macro 65 delete success", 0, vbNullString, vbNullString, macroName, "Macro", "CodeDb", "SystemMacro65"
        Else
            skippedCount = skippedCount + 1
            skippedCountResult = skippedCount
            LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "DeleteSystemHelpMacros", "DELETE SKIP", "Type=Macro; Name=" & macroName & "; Reason=Not found at deletion time", "Macro 65 delete skip", 0, vbNullString, vbNullString, macroName, "Macro", "CodeDb", "SystemMacro65"
        End If
        PumpSetupResponsiveness
    Next item
    deleteSeconds = ElapsedSecondsFromTimer(stageTimer)
    stageTimer = Timer
    PartCClearComparisonPlan
    PartCInvalidatePersistedPhysicalMacroState
    invalidateSeconds = ElapsedSecondsFromTimer(stageTimer)
    LogTraceMessage TRACE_LEVEL_INFO, "Performance", "DeleteSystemHelpMacros", IIf(deferredInvokingMacro, "Macro 65 physical macro ordinary deletion completed; invoking Macro 65 remains deferred.", "Macro 65 physical macro deletion completed."), "Deletion names consumed=" & CStr(deletionNames.count) & "; Attempted=" & CStr(attemptedCount) & "; Deleted=" & CStr(deletedCount) & "; Skipped=" & CStr(skippedCount) & "; Failed=" & CStr(failedCount) & "; DeferredInvokingMacro=" & CStr(deferredInvokingMacro) & "; delete elapsed=" & FormattedElapsedSeconds(deleteSeconds) & "; Phase C proof invalidation elapsed=" & FormattedElapsedSeconds(invalidateSeconds) & "; execution elapsed=" & FormattedElapsedSeconds(ElapsedSecondsFromTimer(totalTimer)), "Macro 65 macro deletion", 0, vbNullString, vbNullString, vbNullString, "MacroSet", "CodeDb", "SystemMacro65"
    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "DeleteSystemHelpMacros", IIf(deferredInvokingMacro, "DELETE ALL ADC OBJECTS ORDINARY MACRO WORK COMPLETE", "DELETE ALL ADC OBJECTS COMPLETE"), "Planned=" & CStr(deletionNames.count) & "; Attempted=" & CStr(attemptedCount) & "; Deleted=" & CStr(deletedCount) & "; Skipped=" & CStr(skippedCount) & "; Failed=" & CStr(failedCount) & "; DeferredInvokingMacro=" & CStr(deferredInvokingMacro), "Macro 65 deletion summary", 0, vbNullString, vbNullString, vbNullString, "MacroSet", "CodeDb", "SystemMacro65"
    messageText = "Delete All ADC Objects completed." & vbCrLf & vbCrLf & "Deleted ADC System Macros and Help Document Macros: " & CStr(deletedCount + IIf(deferredInvokingMacro, 1, 0)) & vbCrLf & "Unrelated user macros were retained."
    If deferredInvokingMacro Then
        If showMessage Then CoreConfigureDeferredMacro65NormalCompletion True, messageText, "Delete All ADC Objects"
    ElseIf showMessage Then
        CoreShowMessage messageText, vbInformation, "Delete All ADC Objects", "ShowADCMessage", "Macros"
    End If
    DeleteSystemHelpMacros = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    PartCClearComparisonPlan
    PartCInvalidatePersistedPhysicalMacroState
    failedCount = failedCount + 1
    failedCountResult = failedCount
    LogTraceMessage TRACE_LEVEL_ERROR, "DeleteObjects", "DeleteSystemHelpMacros", "DELETE FAILED", "Type=Macro; Name=" & IIf(Len(macroName) = 0, "(unknown)", macroName) & "; Error=" & CStr(savedErrNumber) & "; Description=" & savedErrDescription, "Macro 65 delete failure", savedErrNumber, savedErrSource, savedErrDescription, macroName, "Macro", "CodeDb", "SystemMacro65"
    messageText = "Could not delete the ADC System Macros and Help Document Macros." & vbCrLf & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description:" & vbCrLf & savedErrDescription
    If showMessage Then
        CoreShowErrorMessage messageText, vbCritical, "Delete All ADC Objects", "DeleteSystemHelpMacros", "Macros"
        Exit Function
    End If
    Err.Raise savedErrNumber, "DeleteSystemHelpMacros", messageText
End Function

Private Sub PartCClearComparisonPlan()
    Set mPartCExpectedMacros = Nothing
    Set mPartCPhysicalEvidence = Nothing
    Set mPartCBuildPlan = Nothing
    Set mPartCDeletePlan = Nothing
    mPartCExpectedStateToken = vbNullString
    mPartCExpectedCount = 0
    mPartCMatchingCount = 0
    mPartCMissingCount = 0
    mPartCReplacementCount = 0
    mPartCStaleOwnedCount = 0
    mPartCPlanSourceToken = vbNullString
    mPartCLastPhysicalStateToken = vbNullString
    mPartCPlanReady = False
    mPartCHelpDocumentSynchronizationScope = 0
    mPartCHelpDocumentMutationsOnly = False
    mPartCPayloadReadCount = 0
    mPartCRollbackBackupCount = 0
    mPartCTargetedStaleRereadCount = 0
    mPartCCreatedCount = 0
    mPartCReplacedCount = 0
    mPartCDeletedCount = 0
    mPartCWholeSetSecondPayloadPassOccurred = False
End Sub

Private Sub PartCLoadAuthoritativeExpectedMacros(ByRef expectedMacros As Object, ByRef expectedStateToken As String, ByVal synchronizationScope As HelpDocumentSynchronizationScope)
    On Error GoTo ErrHandler
    Dim sourceDb As DAO.Database
    Dim rowData As Variant
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim MacroTitle As String
    Dim runCodeExpression As String
    Dim sourceKey As String
    Dim rowIdentity As String
    Dim helpDocumentNumber As Long
    Dim expectedPayload As String
    Dim expectedKeys As Variant
    Dim item As Variant
    Dim i As Long
    Set sourceDb = ADCDb()
    If SharedDao_TableExists(sourceDb, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SystemMacros")) = False Then Err.Raise vbObjectError + 6526, "PartCLoadAuthoritativeExpectedMacros", "B4 source is invalid: [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SystemMacros") & "] is missing. Phase C does not repair B4."
    If synchronizationScope <> HelpDocumentScopeSetupSystemOnly Then
        If SharedDao_TableExists(sourceDb, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HelpDocuments")) = False Then Err.Raise vbObjectError + 6527, "PartCLoadAuthoritativeExpectedMacros", "B4 source is invalid: [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HelpDocuments") & "] is missing. Phase C does not repair B4."
    End If
    If TableCanOpen(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SystemMacros"), True) = False Then Err.Raise vbObjectError + 6528, "PartCLoadAuthoritativeExpectedMacros", "B4 source is invalid: [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SystemMacros") & "] cannot be opened. Phase C does not repair B4."
    If synchronizationScope <> HelpDocumentScopeSetupSystemOnly Then
        If TableCanOpen(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HelpDocuments"), True) = False Then Err.Raise vbObjectError + 6529, "PartCLoadAuthoritativeExpectedMacros", "B4 source is invalid: [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HelpDocuments") & "] cannot be opened. Phase C does not repair B4."
    End If
    ' INTENTIONAL KEYED LOOKUP: expectedMacros provides keyed membership/lookup over source rows or state already read by PartCLoadAuthoritativeExpectedMacros; it is a derived access path, not a second authority.
    Set expectedMacros = CreateObject("Scripting.Dictionary")
    expectedMacros.CompareMode = vbTextCompare
    For Each rowData In SystemMacroCatalog()
        PumpSetupResponsiveness
        sourceKey = Trim$(CStr(rowData(0)))
        MacroTitle = CStr(rowData(2))
        runCodeExpression = Trim$(CStr(rowData(3)))
        PartCAddExpectedMacroToDictionary expectedMacros, MacroTitle, runCodeExpression, "MC_Sys:" & sourceKey
    Next rowData
    If synchronizationScope <> HelpDocumentScopeSetupSystemOnly Then
        sqlText = "SELECT " & SQL_FIELD_ID & ",[" & FIELD_HELP_DOCUMENT_KEY & "],[" & FIELD_HELP_DOCUMENT_NUMBER & "],[" & FIELD_DEFAULT_TITLE_TEXT & "],[" & FIELD_MACRO_TITLE_OVERRIDE & "] FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HelpDocuments") & "] WHERE Nz(" & SQL_FIELD_HELP_DOCUMENT_STATUS & "," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & " ORDER BY " & OrderByPolicyText("CATALOG_HELP_DOCUMENT") & ";"
        Set rs = sourceDb.OpenRecordset(sqlText, dbOpenSnapshot)
        Do While rs.EOF = False
            PumpSetupResponsiveness
            sourceKey = CStr(Nz(rs.fields(FIELD_HELP_DOCUMENT_KEY).value, vbNullString))
            If HelpDocumentKeyInSynchronizationScope(synchronizationScope) Then
                rowIdentity = CStr(Nz(rs.fields(FIELD_ID).value, vbNullString))
                helpDocumentNumber = CLng(Nz(rs.fields(FIELD_HELP_DOCUMENT_NUMBER).value, 0))
                ResolveHelpDocumentMaterialization sourceKey, helpDocumentNumber, CStr(Nz(rs.fields(FIELD_DEFAULT_TITLE_TEXT).value, vbNullString)), CStr(Nz(rs.fields(FIELD_MACRO_TITLE_OVERRIDE).value, vbNullString)), rowIdentity, synchronizationScope, MacroTitle, runCodeExpression
                PartCAddExpectedMacroToDictionary expectedMacros, MacroTitle, runCodeExpression, "Help_Docs:" & IIf(Len(Trim$(sourceKey)) = 0, "(blank)", sourceKey)
            End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
    If expectedMacros.count = 0 Then Err.Raise vbObjectError + 6530, "PartCLoadAuthoritativeExpectedMacros", "B4 source is invalid: the authoritative System Macro and Help Document Macro set is empty."
    expectedKeys = SortedDictionaryKeys(expectedMacros)
    For i = LBound(expectedKeys) To UBound(expectedKeys)
        item = expectedMacros(CStr(expectedKeys(i)))
        expectedPayload = expectedPayload & FingerprintTokenPart(CStr(item(0))) & FingerprintTokenPart(CStr(item(1)))
    Next i
    expectedStateToken = StableTextHashHex("ADC-PhaseC-v1|" & expectedPayload) & ":" & CStr(Len(expectedPayload))
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
    Set expectedMacros = Nothing
    On Error GoTo 0
    Err.Raise savedErrNumber, "PartCLoadAuthoritativeExpectedMacros", "Could not read the authoritative B4 System Macro and Help Document Macro source without mutation." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Function PartCPhysicalStateTokenFromEvidence() As String
    'Canonical Phase C installed-state token calculation over the one operation-scoped physical evidence snapshot. Missing expected names deliberately return no token.
    Dim expectedKeys As Variant
    Dim expectedItem As Variant
    Dim evidenceItem As Variant
    Dim keyText As String
    Dim physicalPayload As String
    Dim physicalRunCode As String
    Dim i As Long
    If mPartCExpectedMacros Is Nothing Then Exit Function
    If mPartCPhysicalEvidence Is Nothing Then Exit Function
    If mPartCExpectedMacros.count = 0 Then Exit Function
    expectedKeys = SortedDictionaryKeys(mPartCExpectedMacros)
    For i = LBound(expectedKeys) To UBound(expectedKeys)
        keyText = CStr(expectedKeys(i))
        expectedItem = mPartCExpectedMacros(keyText)
        If mPartCPhysicalEvidence.Exists(keyText) = False Then Exit Function
        evidenceItem = mPartCPhysicalEvidence(keyText)
        If CBool(evidenceItem(PART_C_EVIDENCE_EXISTED_INDEX)) = False Then Exit Function
        physicalRunCode = CStr(evidenceItem(PART_C_EVIDENCE_RUNCODE_INDEX))
        physicalPayload = physicalPayload & FingerprintTokenPart(CStr(expectedItem(0))) & FingerprintTokenPart(physicalRunCode)
    Next i
    PartCPhysicalStateTokenFromEvidence = StableTextHashHex("ADC-PhaseC-v1|" & physicalPayload) & ":" & CStr(Len(physicalPayload))
End Function

Private Sub PartCBuildComparisonPlan(ByRef detailText As String, ByVal synchronizationScope As HelpDocumentSynchronizationScope, Optional ByVal helpDocumentMutationsOnly As Boolean = False)
    On Error GoTo ErrHandler
    Dim expectedItem As Variant
    Dim physicalRunCode As String
    Dim allMacro As accessObject
    Dim physicalName As String
    Dim keyText As String
    Dim expectedKeys As Variant
    Dim i As Long
    Dim expectedTotal As Long
    Dim physicalIndex As Long
    Dim physicalTotal As Long
    Dim isOwned As Boolean
    Dim expectedIsHelpDocument As Boolean
    Dim unrelatedSystemIssueCount As Long
    Dim unrelatedSystemIssues As String
    Dim phaseTimer As Single
    Dim secondWholeSetPass As Boolean
    phaseTimer = Timer
    secondWholeSetPass = mPartCPlanReady
    PartCClearComparisonPlan
    mPartCWholeSetSecondPayloadPassOccurred = secondWholeSetPass
    mPartCHelpDocumentSynchronizationScope = synchronizationScope
    mPartCHelpDocumentMutationsOnly = helpDocumentMutationsOnly
    Set mPartCPhysicalEvidence = CreateObject("Scripting.Dictionary")
    mPartCPhysicalEvidence.CompareMode = vbTextCompare
    PartCLoadAuthoritativeExpectedMacros mPartCExpectedMacros, mPartCExpectedStateToken, synchronizationScope
    Set mPartCBuildPlan = New Collection
    Set mPartCDeletePlan = New Collection
    expectedKeys = SortedDictionaryKeys(mPartCExpectedMacros)
    expectedTotal = UBound(expectedKeys) - LBound(expectedKeys) + 1
    For i = LBound(expectedKeys) To UBound(expectedKeys)
        keyText = CStr(expectedKeys(i))
        expectedItem = mPartCExpectedMacros(keyText)
        CorePulseProgress "C1 verifying macro " & CStr(i - LBound(expectedKeys) + 1) & " of " & CStr(expectedTotal) & ": " & CStr(expectedItem(0))
        PumpSetupResponsiveness
        expectedIsHelpDocument = (InStr(1, CStr(expectedItem(2)), "Help_Docs:", vbTextCompare) = 1)
        physicalRunCode = vbNullString
        If PartCPhysicalMacroRunCode(CStr(expectedItem(0)), physicalRunCode) = False Then
            mPartCPhysicalEvidence(keyText) = Array(CStr(expectedItem(0)), False, vbNullString, True, False)
            mPartCMissingCount = mPartCMissingCount + 1
            If helpDocumentMutationsOnly And expectedIsHelpDocument = False Then
                unrelatedSystemIssueCount = unrelatedSystemIssueCount + 1
                unrelatedSystemIssues = SharedText_AppendNonBlankLine(unrelatedSystemIssues, "Missing unrelated System Macro: [" & CStr(expectedItem(0)) & "]")
            Else
                mPartCBuildPlan.Add Array(CStr(expectedItem(0)), CStr(expectedItem(1)), "Missing")
            End If
        Else
            mPartCPhysicalEvidence(keyText) = Array(CStr(expectedItem(0)), True, physicalRunCode, True, False)
            If StrComp(Trim$(physicalRunCode), Trim$(CStr(expectedItem(1))), vbBinaryCompare) <> 0 Then
                mPartCReplacementCount = mPartCReplacementCount + 1
                If helpDocumentMutationsOnly And expectedIsHelpDocument = False Then
                    unrelatedSystemIssueCount = unrelatedSystemIssueCount + 1
                    unrelatedSystemIssues = SharedText_AppendNonBlankLine(unrelatedSystemIssues, "Stale unrelated System Macro: [" & CStr(expectedItem(0)) & "]")
                Else
                    mPartCBuildPlan.Add Array(CStr(expectedItem(0)), CStr(expectedItem(1)), "Stale")
                End If
            Else
                mPartCMatchingCount = mPartCMatchingCount + 1
            End If
        End If
        RequireSafeWorkContinue "PartCBuildComparisonPlan", "C1 macro [" & CStr(expectedItem(0)) & "]"
    Next i
    physicalTotal = CurrentProject.AllMacros.count
    For Each allMacro In CurrentProject.AllMacros
        physicalIndex = physicalIndex + 1
        CorePulseProgress "C1 checking physical macro ownership " & CStr(physicalIndex) & " of " & CStr(physicalTotal) & ": " & allMacro.Name
        PumpSetupResponsiveness
        physicalName = allMacro.Name
        keyText = UCase$(physicalName)
        If mPartCExpectedMacros.Exists(keyText) = False Then
            If PartCPhysicalMacroNameLooksGeneratedOwned(physicalName) Then
                physicalRunCode = vbNullString
                If PartCPhysicalMacroRunCode(physicalName, physicalRunCode) Then
                    isOwned = PartCPhysicalMacroOwnedInSynchronizationScope(physicalRunCode, synchronizationScope)
                    mPartCPhysicalEvidence(keyText) = Array(physicalName, True, physicalRunCode, False, isOwned)
                    If isOwned Then
                        mPartCStaleOwnedCount = mPartCStaleOwnedCount + 1
                        If helpDocumentMutationsOnly And InStr(1, Trim$(physicalRunCode), "RunSystemCommand(", vbTextCompare) = 1 Then
                            unrelatedSystemIssueCount = unrelatedSystemIssueCount + 1
                            unrelatedSystemIssues = SharedText_AppendNonBlankLine(unrelatedSystemIssues, "Unexpected unrelated System Macro owned by ADC: [" & physicalName & "]")
                        Else
                            mPartCDeletePlan.Add physicalName
                        End If
                    End If
                Else
                    mPartCPhysicalEvidence(keyText) = Array(physicalName, False, vbNullString, False, False)
                End If
            End If
        End If
        RequireSafeWorkContinue "PartCBuildComparisonPlan", "C1 physical macro [" & physicalName & "]"
    Next allMacro
    If helpDocumentMutationsOnly And unrelatedSystemIssueCount > 0 Then
        Err.Raise vbObjectError + 6577, "PartCBuildComparisonPlan", "Build/Repair Help Document Macros found unrelated System Macro physical-state defects. The Help-only operational path will not rewrite or delete unrelated System Macros." & vbCrLf & _
            "Unrelated System Macro issues: " & CStr(unrelatedSystemIssueCount) & vbCrLf & unrelatedSystemIssues
    End If
    mPartCExpectedCount = mPartCExpectedMacros.count
    mPartCPlanSourceToken = mPartCExpectedStateToken
    mPartCLastPhysicalStateToken = PartCPhysicalStateTokenFromEvidence()
    mPartCPlanReady = True
    detailText = "C1 physical System Macro and Help Document Macro comparison:" & vbCrLf & _
                 " - Expected: " & CStr(mPartCExpectedCount) & vbCrLf & _
                 " - Matching: " & CStr(mPartCMatchingCount) & vbCrLf & _
                 " - Missing: " & CStr(mPartCMissingCount) & vbCrLf & _
                 " - Needing replacement: " & CStr(mPartCReplacementCount) & vbCrLf & _
                 " - Stale " & DEFAULT_UTILITY_SHORT_NAME & "-owned names: " & CStr(mPartCStaleOwnedCount) & vbCrLf & _
                 " - C2 build operations: " & CStr(mPartCBuildPlan.count)
    LogTraceMessage TRACE_LEVEL_INFO, "PhysicalMacros", "PartCBuildComparisonPlan", "C1 physical macro evidence snapshot completed.", _
        "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseTimer)) & vbCrLf & _
        "Payload SaveAsText reads: " & CStr(mPartCPayloadReadCount) & vbCrLf & _
        "Rollback SaveAsText backups: " & CStr(mPartCRollbackBackupCount) & vbCrLf & _
        "Expected macros: " & CStr(mPartCExpectedCount) & vbCrLf & _
        "Matching: " & CStr(mPartCMatchingCount) & vbCrLf & _
        "Missing: " & CStr(mPartCMissingCount) & vbCrLf & _
        "Replacement targets: " & CStr(mPartCReplacementCount) & vbCrLf & _
        "Created: " & CStr(mPartCCreatedCount) & vbCrLf & _
        "Replaced: " & CStr(mPartCReplacedCount) & vbCrLf & _
        "Deleted: " & CStr(mPartCDeletedCount) & vbCrLf & _
        "Targeted stale-plan payload rereads: " & CStr(mPartCTargetedStaleRereadCount) & vbCrLf & _
        "Whole-set second payload pass occurred: " & CStr(mPartCWholeSetSecondPayloadPassOccurred), _
        "C1 physical macros", 0, vbNullString, vbNullString, vbNullString, "SystemMacro", "CodeDb", "Performance"
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    LogTraceMessage TRACE_LEVEL_ERROR, "PhysicalMacros", "PartCBuildComparisonPlan", "C1 physical macro evidence snapshot failed.", _
        "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseTimer)) & vbCrLf & _
        "Payload SaveAsText reads: " & CStr(mPartCPayloadReadCount) & vbCrLf & _
        "Rollback SaveAsText backups: " & CStr(mPartCRollbackBackupCount) & vbCrLf & _
        "Expected macros: " & CStr(mPartCExpectedCount) & vbCrLf & _
        "Matching: " & CStr(mPartCMatchingCount) & vbCrLf & _
        "Missing: " & CStr(mPartCMissingCount) & vbCrLf & _
        "Replacement targets: " & CStr(mPartCReplacementCount) & vbCrLf & _
        "Created: " & CStr(mPartCCreatedCount) & vbCrLf & _
        "Replaced: " & CStr(mPartCReplacedCount) & vbCrLf & _
        "Deleted: " & CStr(mPartCDeletedCount) & vbCrLf & _
        "Targeted stale-plan payload rereads: " & CStr(mPartCTargetedStaleRereadCount) & vbCrLf & _
        "Whole-set second payload pass occurred: " & CStr(mPartCWholeSetSecondPayloadPassOccurred), _
        "C1 physical macros", savedErrNumber, savedErrSource, savedErrDescription, vbNullString, "SystemMacro", "CodeDb", "Performance"
    On Error GoTo 0
    PartCClearComparisonPlan
    Err.Raise savedErrNumber, "PartCBuildComparisonPlan", "Could not build the read-only C1 physical macro comparison plan." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub PartCAddExpectedMacroToDictionary(ByVal expectedMacros As Object, ByVal MacroTitle As String, ByVal runCodeExpression As String, ByVal sourceLabel As String)
    runCodeExpression = Trim$(runCodeExpression)
    If expectedMacros Is Nothing Then Err.Raise vbObjectError + 6533, "PartCAddExpectedMacroToDictionary", "The expected-macro dictionary was not initialized."
    If Len(MacroTitle) = 0 Then Err.Raise vbObjectError + 6534, "PartCAddExpectedMacroToDictionary", "B4 source contains a blank physical macro title. Source: " & sourceLabel
    If Len(runCodeExpression) = 0 Then Err.Raise vbObjectError + 6535, "PartCAddExpectedMacroToDictionary", "B4 source contains a blank RunCode expression. Source: " & sourceLabel
    If Len(MacroTitle) > MAX_ACCESS_OBJECT_NAME_LENGTH Then Err.Raise vbObjectError + 6536, "PartCAddExpectedMacroToDictionary", "B4 source contains a macro title longer than the Access object-name limit." & vbCrLf & MacroTitle
    If expectedMacros.Exists(UCase$(MacroTitle)) Then Err.Raise vbObjectError + 6537, "PartCAddExpectedMacroToDictionary", "B4 source resolves more than one authoritative row to the same physical macro name." & vbCrLf & MacroTitle
    expectedMacros.Add UCase$(MacroTitle), Array(MacroTitle, runCodeExpression, sourceLabel)
End Sub

Private Sub PartCValidatePlanSourceCurrent(ByVal phaseIdentity As String)
    Dim currentExpected As Object
    Dim currentToken As String
    PartCLoadAuthoritativeExpectedMacros currentExpected, currentToken, mPartCHelpDocumentSynchronizationScope
    If StrComp(currentToken, mPartCPlanSourceToken, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 6538, "PartCValidatePlanSourceCurrent", "B4 authoritative System Macro and Help Document Macro source changed after C1. " & phaseIdentity & " refused to use the stale C1 plan. Run C1 again."
End Sub

Private Sub PartCValidateBuildPlanTargetsCurrentForC2()
    Dim item As Variant
    Dim evidenceItem As Variant
    Dim expectedItem As Variant
    Dim macroName As String
    Dim keyText As String
    Dim currentRunCode As String
    Dim c1ObservedExisting As Boolean
    Dim c1ObservedRunCode As String
    If mPartCPhysicalEvidence Is Nothing Then Err.Raise vbObjectError + 6567, "PartCValidateBuildPlanTargetsCurrentForC2", "The authoritative C1 physical-evidence snapshot is unavailable before C2 target validation."
    For Each item In mPartCBuildPlan
        macroName = CStr(item(0))
        keyText = UCase$(macroName)
        If mPartCPhysicalEvidence.Exists(keyText) = False Then Err.Raise vbObjectError + 6568, "PartCValidateBuildPlanTargetsCurrentForC2", "C1 physical evidence is missing for a planned C2 mutation target." & vbCrLf & "Macro: " & macroName
        evidenceItem = mPartCPhysicalEvidence(keyText)
        If CBool(evidenceItem(PART_C_EVIDENCE_EXPECTED_INDEX)) = False Then Err.Raise vbObjectError + 6569, "PartCValidateBuildPlanTargetsCurrentForC2", "C1 physical evidence does not classify a planned C2 mutation target as an authoritative expected macro." & vbCrLf & "Macro: " & macroName
        If mPartCHelpDocumentMutationsOnly Then
            expectedItem = mPartCExpectedMacros(keyText)
            If InStr(1, CStr(expectedItem(2)), "Help_Docs:", vbTextCompare) <> 1 Then Err.Raise vbObjectError + 6578, "PartCValidateBuildPlanTargetsCurrentForC2", "Build/Repair Help Document Macros C2 refused a planned mutation outside the Help_Docs physical-macro scope." & vbCrLf & "Macro: " & macroName
        End If
        If StrComp(CStr(evidenceItem(PART_C_EVIDENCE_NAME_INDEX)), macroName, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 6570, "PartCValidateBuildPlanTargetsCurrentForC2", "C1 physical evidence resolved a different physical name for a planned C2 mutation target." & vbCrLf & "Macro: " & macroName
        c1ObservedExisting = CBool(evidenceItem(PART_C_EVIDENCE_EXISTED_INDEX))
        If c1ObservedExisting Then
            currentRunCode = vbNullString
            mPartCTargetedStaleRereadCount = mPartCTargetedStaleRereadCount + 1
            If PartCPhysicalMacroRunCode(macroName, currentRunCode) = False Then Err.Raise vbObjectError + 6539, "PartCValidateBuildPlanTargetsCurrentForC2", "The C1-observed existing C2 mutation target disappeared before C2. C2 refused to execute a stale change plan. Run C1 again." & vbCrLf & "Macro: " & macroName
            c1ObservedRunCode = CStr(evidenceItem(PART_C_EVIDENCE_RUNCODE_INDEX))
            If StrComp(currentRunCode, c1ObservedRunCode, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 6540, "PartCValidateBuildPlanTargetsCurrentForC2", "The C1-observed physical RunCode for a C2 mutation target changed before C2. C2 refused to execute a stale change plan. Run C1 again." & vbCrLf & "Macro: " & macroName
        Else
            If RollbackObjectExistsStrict(currentDb, acMacro, macroName, "during C2 planned-target validation") Then Err.Raise vbObjectError + 6539, "PartCValidateBuildPlanTargetsCurrentForC2", "A C2 mutation target that C1 observed as missing now exists. C2 refused to execute a stale change plan. Run C1 again." & vbCrLf & "Macro: " & macroName
        End If
    Next item
End Sub

Private Sub PartCValidateDeletePlanTargetsCurrentForC3()
    'Targeted C3 stale-delete protection: compare current names to C1 evidence, inspect only newly appearing generated-looking candidates, and leave each planned target's exact payload reread to the delete routine
    'immediately before mutation.
    On Error GoTo ErrHandler
    Dim currentNames As Object
    Dim allMacro As accessObject
    Dim currentKeys As Variant
    Dim item As Variant
    Dim evidenceItem As Variant
    Dim physicalName As String
    Dim macroName As String
    Dim keyText As String
    Dim runCodeExpression As String
    Dim i As Long
    If mPartCExpectedMacros Is Nothing Then Err.Raise vbObjectError + 6571, "PartCValidateDeletePlanTargetsCurrentForC3", "The authoritative expected-macro snapshot is unavailable before C3 delete-plan validation."
    If mPartCPhysicalEvidence Is Nothing Then Err.Raise vbObjectError + 6572, "PartCValidateDeletePlanTargetsCurrentForC3", "The authoritative C1 physical-evidence snapshot is unavailable before C3 delete-plan validation."
    Set currentNames = CreateObject("Scripting.Dictionary")
    currentNames.CompareMode = vbTextCompare
    For Each allMacro In CurrentProject.AllMacros
        physicalName = allMacro.Name
        currentNames(UCase$(physicalName)) = physicalName
    Next allMacro
    For Each item In mPartCDeletePlan
        macroName = CStr(item)
        keyText = UCase$(macroName)
        If currentNames.Exists(keyText) = False Then Err.Raise vbObjectError + 6523, "PartCValidateDeletePlanTargetsCurrentForC3", "A C1 stale-macro delete target disappeared before C3. C3 refused to execute a stale delete plan. Run C1 again." & vbCrLf & "Macro: " & macroName
        If mPartCPhysicalEvidence.Exists(keyText) = False Then Err.Raise vbObjectError + 6573, "PartCValidateDeletePlanTargetsCurrentForC3", "C1 physical evidence is missing for a planned C3 delete target." & vbCrLf & "Macro: " & macroName
        evidenceItem = mPartCPhysicalEvidence(keyText)
        If CBool(evidenceItem(PART_C_EVIDENCE_EXISTED_INDEX)) = False Or CBool(evidenceItem(PART_C_EVIDENCE_STALE_OWNED_INDEX)) = False Then Err.Raise vbObjectError + 6574, "PartCValidateDeletePlanTargetsCurrentForC3", "C1 physical evidence no longer authorizes a planned C3 delete target." & vbCrLf & "Macro: " & macroName
        If StrComp(CStr(evidenceItem(PART_C_EVIDENCE_NAME_INDEX)), macroName, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 6575, "PartCValidateDeletePlanTargetsCurrentForC3", "C1 physical evidence resolved a different physical name for a planned C3 delete target." & vbCrLf & "Macro: " & macroName
    Next item
    If currentNames.count > 0 Then
        currentKeys = currentNames.keys
        For i = LBound(currentKeys) To UBound(currentKeys)
            keyText = CStr(currentKeys(i))
            physicalName = CStr(currentNames(keyText))
            If mPartCExpectedMacros.Exists(keyText) = False Then
                If PartCPhysicalMacroNameLooksGeneratedOwned(physicalName) Then
                    If mPartCPhysicalEvidence.Exists(keyText) = False Then
                        runCodeExpression = vbNullString
                        If PartCPhysicalMacroRunCode(physicalName, runCodeExpression) Then
                            If PartCPhysicalMacroOwnedInSynchronizationScope(runCodeExpression, mPartCHelpDocumentSynchronizationScope) Then Err.Raise vbObjectError + 6523, "PartCValidateDeletePlanTargetsCurrentForC3", "A new relevant ADC-owned physical macro appeared after C1. C3 refused to execute a stale delete plan. Run C1 again." & vbCrLf & "Macro: " & physicalName
                        End If
                    Else
                        evidenceItem = mPartCPhysicalEvidence(keyText)
                        If CBool(evidenceItem(PART_C_EVIDENCE_EXISTED_INDEX)) = False Then
                            runCodeExpression = vbNullString
                            If PartCPhysicalMacroRunCode(physicalName, runCodeExpression) Then
                                If PartCPhysicalMacroOwnedInSynchronizationScope(runCodeExpression, mPartCHelpDocumentSynchronizationScope) Then Err.Raise vbObjectError + 6523, "PartCValidateDeletePlanTargetsCurrentForC3", "A relevant ADC-owned physical macro reappeared after C1. C3 refused to execute a stale delete plan. Run C1 again." & vbCrLf & "Macro: " & physicalName
                            End If
                        End If
                    End If
                End If
            End If
        Next i
    End If
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "PartCValidateDeletePlanTargetsCurrentForC3", "Could not validate the targeted C3 stale-delete plan against current physical macro names." & vbCrLf & Err.Description
End Sub

Private Sub PartCValidateFinalPhysicalStateFromEvidence(ByRef actualStateToken As String, ByRef verifyDetail As String)
    'Final C3 proof consumes C1/C2/C3 evidence and a cheap current name inventory. Only a newly appearing generated-looking candidate requires a new payload read.
    On Error GoTo ErrHandler
    Dim currentNames As Object
    Dim allMacro As accessObject
    Dim expectedKeys As Variant
    Dim currentKeys As Variant
    Dim expectedItem As Variant
    Dim evidenceItem As Variant
    Dim physicalName As String
    Dim keyText As String
    Dim runCodeExpression As String
    Dim finalMatching As Long
    Dim finalMissing As Long
    Dim finalReplacement As Long
    Dim finalStaleOwned As Long
    Dim i As Long
    actualStateToken = vbNullString
    verifyDetail = vbNullString
    If mPartCExpectedMacros Is Nothing Then Err.Raise vbObjectError + 6548, "PartCValidateFinalPhysicalStateFromEvidence", "The authoritative expected-macro snapshot is unavailable during C3 final verification."
    If mPartCPhysicalEvidence Is Nothing Then Err.Raise vbObjectError + 6549, "PartCValidateFinalPhysicalStateFromEvidence", "The authoritative C1 physical-evidence snapshot is unavailable during C3 final verification."
    Set currentNames = CreateObject("Scripting.Dictionary")
    currentNames.CompareMode = vbTextCompare
    For Each allMacro In CurrentProject.AllMacros
        physicalName = allMacro.Name
        currentNames(UCase$(physicalName)) = physicalName
    Next allMacro
    expectedKeys = SortedDictionaryKeys(mPartCExpectedMacros)
    For i = LBound(expectedKeys) To UBound(expectedKeys)
        keyText = CStr(expectedKeys(i))
        expectedItem = mPartCExpectedMacros(keyText)
        If currentNames.Exists(keyText) = False Then
            finalMissing = finalMissing + 1
        ElseIf mPartCPhysicalEvidence.Exists(keyText) = False Then
            finalMissing = finalMissing + 1
        Else
            evidenceItem = mPartCPhysicalEvidence(keyText)
            If CBool(evidenceItem(PART_C_EVIDENCE_EXPECTED_INDEX)) = False Or StrComp(CStr(evidenceItem(PART_C_EVIDENCE_NAME_INDEX)), CStr(expectedItem(0)), vbTextCompare) <> 0 Then
                finalMissing = finalMissing + 1
            ElseIf CBool(evidenceItem(PART_C_EVIDENCE_EXISTED_INDEX)) = False Then
                finalMissing = finalMissing + 1
            ElseIf StrComp(Trim$(CStr(evidenceItem(PART_C_EVIDENCE_RUNCODE_INDEX))), Trim$(CStr(expectedItem(1))), vbBinaryCompare) <> 0 Then
                finalReplacement = finalReplacement + 1
            Else
                finalMatching = finalMatching + 1
            End If
        End If
    Next i
    If currentNames.count > 0 Then
        currentKeys = currentNames.keys
        For i = LBound(currentKeys) To UBound(currentKeys)
            keyText = CStr(currentKeys(i))
            physicalName = CStr(currentNames(keyText))
            If mPartCExpectedMacros.Exists(keyText) = False Then
                If PartCPhysicalMacroNameLooksGeneratedOwned(physicalName) Then
                    If mPartCPhysicalEvidence.Exists(keyText) Then
                        evidenceItem = mPartCPhysicalEvidence(keyText)
                        If CBool(evidenceItem(PART_C_EVIDENCE_EXISTED_INDEX)) Then
                            If CBool(evidenceItem(PART_C_EVIDENCE_STALE_OWNED_INDEX)) Then finalStaleOwned = finalStaleOwned + 1
                        Else
                            runCodeExpression = vbNullString
                            If PartCPhysicalMacroRunCode(physicalName, runCodeExpression) Then
                                If PartCPhysicalMacroOwnedInSynchronizationScope(runCodeExpression, mPartCHelpDocumentSynchronizationScope) Then finalStaleOwned = finalStaleOwned + 1
                            End If
                        End If
                    Else
                        runCodeExpression = vbNullString
                        If PartCPhysicalMacroRunCode(physicalName, runCodeExpression) Then
                            If PartCPhysicalMacroOwnedInSynchronizationScope(runCodeExpression, mPartCHelpDocumentSynchronizationScope) Then finalStaleOwned = finalStaleOwned + 1
                        End If
                    End If
                End If
            End If
        Next i
    End If
    mPartCMatchingCount = finalMatching
    mPartCMissingCount = finalMissing
    mPartCReplacementCount = finalReplacement
    mPartCStaleOwnedCount = finalStaleOwned
    actualStateToken = PartCPhysicalStateTokenFromEvidence()
    mPartCLastPhysicalStateToken = actualStateToken
    verifyDetail = "C3 evidence-based physical System Macro and Help Document Macro proof: expected=" & CStr(mPartCExpectedCount) & "; matching=" & CStr(finalMatching) & "; missing=" & CStr(finalMissing) & "; replacement=" & CStr(finalReplacement) & "; stale-owned=" & CStr(finalStaleOwned) & "."
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "PartCValidateFinalPhysicalStateFromEvidence", "Could not derive the final Phase C physical macro state from authoritative operation evidence." & vbCrLf & Err.Description
End Sub

Private Function PartCPhysicalMacroOwnedInSynchronizationScope(ByVal runCodeExpression As String, ByVal synchronizationScope As HelpDocumentSynchronizationScope) As Boolean
    runCodeExpression = Trim$(runCodeExpression)
    If InStr(1, runCodeExpression, "RunSystemCommand(", vbTextCompare) = 1 Then PartCPhysicalMacroOwnedInSynchronizationScope = True: Exit Function
    If InStr(1, runCodeExpression, "Run_OpenHelpDocument(", vbTextCompare) <> 1 Then Exit Function
    PartCPhysicalMacroOwnedInSynchronizationScope = HelpDocumentRunCodeInSynchronizationScope(runCodeExpression, synchronizationScope)
End Function
Private Function PartCPhysicalMacroRunCode(ByVal macroName As String, ByRef runCodeExpression As String) As Boolean
    On Error GoTo ErrHandler
    Dim tempPath As String
    Dim macroText As String
    runCodeExpression = vbNullString
    If Len(macroName) = 0 Then Exit Function
    If RollbackObjectExistsStrict(currentDb, acMacro, macroName, "during Phase C physical macro comparison") = False Then Exit Function
    tempPath = Environ$("TEMP") & "\ADC_C1_" & StableTextHashHex(macroName & "|" & Format$(Now, "yyyymmddhhnnss") & "|" & CStr(Timer)) & ".txt"
    Application.SaveAsText acMacro, macroName, tempPath
    mPartCPayloadReadCount = mPartCPayloadReadCount + 1
    macroText = PartCReadSaveAsTextFile(tempPath)
    runCodeExpression = PartCExtractRunCodeArgument(macroText)
    PartCPhysicalMacroRunCode = True
    On Error Resume Next
    If Len(tempPath) > 0 Then If Len(Dir$(tempPath)) > 0 Then Kill tempPath
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Len(tempPath) > 0 Then If Len(Dir$(tempPath)) > 0 Then Kill tempPath
    On Error GoTo 0
    Err.Raise savedErrNumber, "PartCPhysicalMacroRunCode", "Could not read the physical Access macro during Phase C comparison." & vbCrLf & "Macro: " & macroName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function PartCReadSaveAsTextFile(ByVal filePath As String) As String
    On Error GoTo ErrHandler
    Dim fileNum As Integer
    Dim fileLength As Long
    Dim fileBytes() As Byte
    Dim resultText As String
    Dim startAt As Long
    Dim charCount As Long
    Dim i As Long
    Dim codeUnit As Long
    fileNum = FreeFile
    Open filePath For Binary Access Read As #fileNum
    fileLength = LOF(fileNum)
    If fileLength = 0 Then
        Close #fileNum
        fileNum = 0
        Exit Function
    End If
    ReDim fileBytes(0 To fileLength - 1)
    Get #fileNum, , fileBytes
    Close #fileNum
    fileNum = 0
    ' Access SaveAsText writes UTF-16 little-endian text with an FF FE BOM.
    If fileLength >= 2 Then
        If fileBytes(0) = &HFF And fileBytes(1) = &HFE Then
            startAt = 2
            charCount = (fileLength - startAt) \ 2
            If charCount > 0 Then
                resultText = String$(charCount, vbNullChar)
                For i = 0 To charCount - 1
                    codeUnit = CLng(fileBytes(startAt + (i * 2))) + _
                               (CLng(fileBytes(startAt + (i * 2) + 1)) * &H100&)
                    Mid$(resultText, i + 1, 1) = ChrW(codeUnit)
                Next i
            End If
            PartCReadSaveAsTextFile = resultText
            Exit Function
        End If
    End If
    ' Fail closed if Access ever returns an unexpected encoding.
    Err.Raise vbObjectError + 6545, _
              "PartCReadSaveAsTextFile", _
              "Access SaveAsText returned an unexpected text encoding." & vbCrLf & _
              "File: " & filePath
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
    Err.Raise savedErrNumber, _
              "PartCReadSaveAsTextFile", _
              "Could not read the Access SaveAsText output correctly." & vbCrLf & _
              "File: " & filePath & vbCrLf & _
              "Err.Source: " & savedErrSource & vbCrLf & _
              "Err.Description: " & savedErrDescription
End Function


Private Function PartCPhysicalMacroNameLooksGeneratedOwned(ByVal macroName As String) As Boolean
    Dim logicalName As String
    Dim rebuiltName As String
    Dim separatorPosition As Long
    Dim designatorText As String
    Dim usesConfiguredSystemMacroAffix As Boolean
    If Len(macroName) = 0 Then Exit Function
    logicalName = LogicalObjectNameFromPhysicalName(NAMING_CATEGORY_SYSTEM_MACRO, macroName)
    usesConfiguredSystemMacroAffix = (StrComp(logicalName, macroName, vbBinaryCompare) <> 0)
    separatorPosition = InStr(1, logicalName, " ", vbBinaryCompare)
    If separatorPosition > 1 Then
        designatorText = Left$(logicalName, separatorPosition - 1)
        ' Text designators cannot safely make every arbitrary first word an ADC-ownership signal. A configured System Macro affix establishes the ADC namespace; the decimal-only branch preserves the preexisting
        ' no-affix legacy heuristic. Payload ownership is still proved separately before mutation.
        If usesConfiguredSystemMacroAffix Or TextContainsDecimalDigitsOnly(designatorText) Then
            rebuiltName = TargetManagedObjectPhysicalName(NAMING_CATEGORY_SYSTEM_MACRO, designatorText & "|" & Mid$(logicalName, separatorPosition + 1))
            If StrComp(rebuiltName, macroName, vbBinaryCompare) = 0 Then PartCPhysicalMacroNameLooksGeneratedOwned = True: Exit Function
        End If
    End If
    If PartCSetupPartPhysicalPrefix(macroName) Then
        PartCPhysicalMacroNameLooksGeneratedOwned = True
        Exit Function
    End If
End Function

Private Function PartCSetupPartPhysicalPrefix(ByVal macroName As String) As Boolean
    If Len(macroName) < 5 Then Exit Function
    If Left$(macroName, 1) <> "_" Then Exit Function
    If InStr(1, "ABCDEFGHIJ", UCase$(Mid$(macroName, 2, 1)), vbBinaryCompare) = 0 Then Exit Function
    If (Mid$(macroName, 3, 1) Like "#") = False Then Exit Function
    PartCSetupPartPhysicalPrefix = (Mid$(macroName, 4, 1) = " ")
End Function

Private Sub PartCCreateOrReplaceMacro(ByVal MacroTitle As String, ByVal runCodeExpression As String)
    On Error GoTo ErrHandler
    Dim targetExisted As Boolean
    Dim backupPath As String
    Dim sourcePath As String
    Dim rollbackTracking As Boolean
    Dim replacementLoaded As Boolean
    Dim cleanupWarnings As String
    Dim backupFailureText As String
    Dim loadedRunCode As String
    Dim generatedMacroText As String
    targetExisted = RollbackObjectExistsStrict(currentDb, acMacro, MacroTitle, "before C2 macro replacement")
    rollbackTracking = SetupRollbackTrackingActive()
    If targetExisted Then
        backupPath = PartCMacroTempPath("rollback", MacroTitle)
        Application.SaveAsText acMacro, MacroTitle, backupPath
        mPartCRollbackBackupCount = mPartCRollbackBackupCount + 1
        If SetupRollbackBackupFileReadable(backupPath, backupFailureText) = False Then Err.Raise vbObjectError + 6541, "PartCCreateOrReplaceMacro", "The existing macro could not be safely backed up before C2 replacement." & vbCrLf & "Macro: " & MacroTitle & vbCrLf & backupFailureText
    End If
    If rollbackTracking Then TrackSetupAccessObjectRollback SETUP_ROLLBACK_RESTORE_MACRO_ACTION, MacroTitle, targetExisted, backupPath
    If targetExisted Then
        On Error Resume Next
        DoCmd.Close acMacro, MacroTitle, acSaveYes
        Err.Clear
        On Error GoTo ErrHandler
        DoCmd.DeleteObject acMacro, MacroTitle
    End If
    If RollbackObjectExistsStrict(currentDb, acMacro, MacroTitle, "after C2 deleting the prior macro") Then Err.Raise vbObjectError + 6542, "PartCCreateOrReplaceMacro", "The prior physical macro remained present after deletion." & vbCrLf & MacroTitle
    sourcePath = PartCMacroTempPath("build", MacroTitle)
    generatedMacroText = DirectRunCodeMacroText(runCodeExpression)
    PartCWriteTextFile sourcePath, generatedMacroText
    Application.LoadFromText acMacro, MacroTitle, sourcePath
    replacementLoaded = True
    If PartCPhysicalMacroRunCode(MacroTitle, loadedRunCode) = False Then Err.Raise vbObjectError + 6543, "PartCCreateOrReplaceMacro", "The replacement physical macro was not present after LoadFromText." & vbCrLf & MacroTitle
    If StrComp(Trim$(loadedRunCode), Trim$(runCodeExpression), vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 6544, "PartCCreateOrReplaceMacro", "The replacement macro loaded, but its physical RunCode payload did not match the C1 plan." & vbCrLf & "Macro: " & MacroTitle & vbCrLf & "Expected RunCode: " & runCodeExpression & vbCrLf & "Physical RunCode: " & loadedRunCode & vbCrLf & "Generated macro source: " & Replace(generatedMacroText, vbCrLf, " | ")
    DeleteRollbackBackupFile sourcePath, cleanupWarnings
    If rollbackTracking = False And targetExisted Then DeleteRollbackBackupFile backupPath, cleanupWarnings
    If Len(cleanupWarnings) > 0 Then CoreAddOperationNote cleanupWarnings
    If mPartCPhysicalEvidence Is Nothing Then Err.Raise vbObjectError + 6562, "PartCCreateOrReplaceMacro", "The authoritative Phase C physical-evidence snapshot was unavailable after the verified C2 write."
    mPartCPhysicalEvidence(UCase$(MacroTitle)) = Array(MacroTitle, True, loadedRunCode, True, False)
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If replacementLoaded Or RollbackObjectExistsStrict(currentDb, acMacro, MacroTitle, "during failed C2 replacement cleanup") Then DoCmd.DeleteObject acMacro, MacroTitle
    If targetExisted And Len(backupPath) > 0 And Len(Dir$(backupPath)) > 0 Then Application.LoadFromText acMacro, MacroTitle, backupPath
    If Len(sourcePath) > 0 Then If Len(Dir$(sourcePath)) > 0 Then Kill sourcePath
    On Error GoTo 0
    Err.Raise savedErrNumber, "PartCCreateOrReplaceMacro", "Could not safely create or replace System Macro and Help Document Macro [" & MacroTitle & "]." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub PartCDeleteStaleOwnedMacro(ByVal macroName As String)
    On Error GoTo ErrHandler
    Dim backupPath As String
    Dim rollbackTracking As Boolean
    Dim cleanupWarnings As String
    Dim backupFailureText As String
    Dim evidenceItem As Variant
    Dim c1RunCode As String
    Dim currentRunCode As String
    Dim keyText As String
    keyText = UCase$(macroName)
    If mPartCPhysicalEvidence Is Nothing Then Err.Raise vbObjectError + 6563, "PartCDeleteStaleOwnedMacro", "The authoritative C1 physical-evidence snapshot is unavailable before C3 deletion." & vbCrLf & macroName
    If mPartCPhysicalEvidence.Exists(keyText) = False Then Err.Raise vbObjectError + 6564, "PartCDeleteStaleOwnedMacro", "The C1 stale-macro evidence is unavailable before C3 deletion." & vbCrLf & macroName
    evidenceItem = mPartCPhysicalEvidence(keyText)
    If CBool(evidenceItem(PART_C_EVIDENCE_EXISTED_INDEX)) = False Or CBool(evidenceItem(PART_C_EVIDENCE_STALE_OWNED_INDEX)) = False Or CBool(evidenceItem(PART_C_EVIDENCE_EXPECTED_INDEX)) Then Err.Raise vbObjectError + 6565, "PartCDeleteStaleOwnedMacro", "The C1 physical evidence does not authorize this stale-macro deletion." & vbCrLf & macroName
    If StrComp(CStr(evidenceItem(PART_C_EVIDENCE_NAME_INDEX)), macroName, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 6566, "PartCDeleteStaleOwnedMacro", "The C1 physical evidence resolved a different physical name for this stale-macro deletion." & vbCrLf & macroName
    c1RunCode = CStr(evidenceItem(PART_C_EVIDENCE_RUNCODE_INDEX))
    If mPartCHelpDocumentMutationsOnly And InStr(1, Trim$(c1RunCode), "RunSystemCommand(", vbTextCompare) = 1 Then Err.Raise vbObjectError + 6579, "PartCDeleteStaleOwnedMacro", "Build/Repair Help Document Macros C3 refused to delete an unrelated System Macro." & vbCrLf & macroName
    currentRunCode = vbNullString
    mPartCTargetedStaleRereadCount = mPartCTargetedStaleRereadCount + 1
    If PartCPhysicalMacroRunCode(macroName, currentRunCode) = False Then Err.Raise vbObjectError + 6545, "PartCDeleteStaleOwnedMacro", "The C1 stale-macro target disappeared before C3 deletion. C3 refused to execute a stale delete plan." & vbCrLf & macroName
    If StrComp(currentRunCode, c1RunCode, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 6576, "PartCDeleteStaleOwnedMacro", "The C1-observed physical RunCode for a C3 delete target changed before deletion. C3 refused to execute a stale delete plan. Run C1 again." & vbCrLf & "Macro: " & macroName
    rollbackTracking = SetupRollbackTrackingActive()
    backupPath = PartCMacroTempPath("stale", macroName)
    Application.SaveAsText acMacro, macroName, backupPath
    mPartCRollbackBackupCount = mPartCRollbackBackupCount + 1
    If SetupRollbackBackupFileReadable(backupPath, backupFailureText) = False Then Err.Raise vbObjectError + 6546, "PartCDeleteStaleOwnedMacro", "The stale macro could not be safely backed up before C3 deletion." & vbCrLf & "Macro: " & macroName & vbCrLf & backupFailureText
    If rollbackTracking Then TrackSetupAccessObjectRollback SETUP_ROLLBACK_RESTORE_MACRO_ACTION, macroName, True, backupPath
    On Error Resume Next
    DoCmd.Close acMacro, macroName, acSaveYes
    Err.Clear
    On Error GoTo ErrHandler
    DoCmd.DeleteObject acMacro, macroName
    If RollbackObjectExistsStrict(currentDb, acMacro, macroName, "after C3 stale macro deletion") Then Err.Raise vbObjectError + 6547, "PartCDeleteStaleOwnedMacro", "A stale " & DEFAULT_UTILITY_SHORT_NAME & "-owned macro remained present after deletion." & vbCrLf & macroName
    If rollbackTracking = False Then DeleteRollbackBackupFile backupPath, cleanupWarnings
    If Len(cleanupWarnings) > 0 Then CoreAddOperationNote cleanupWarnings
    mPartCPhysicalEvidence(keyText) = Array(macroName, False, currentRunCode, False, True)
    mPartCDeletedCount = mPartCDeletedCount + 1
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If rollbackTracking = False Then
        On Error Resume Next
        If RollbackObjectExistsStrict(currentDb, acMacro, macroName, "during failed C3 stale-macro cleanup") = False Then
            If Len(backupPath) > 0 And Len(Dir$(backupPath)) > 0 Then Application.LoadFromText acMacro, macroName, backupPath
        End If
        On Error GoTo 0
    End If
    Err.Raise savedErrNumber, "PartCDeleteStaleOwnedMacro", "Could not safely delete stale " & DEFAULT_UTILITY_SHORT_NAME & "-owned macro [" & macroName & "]." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Function PartCMacroTempPath(ByVal purposeText As String, ByVal macroName As String) As String
    PartCMacroTempPath = Environ$("TEMP") & "\ADC_C_" & purposeText & "_" & StableTextHashHex(macroName & "|" & Format$(Now, "yyyymmddhhnnss") & "|" & CStr(Timer)) & ".txt"
End Function

Private Sub PartCWriteTextFile(ByVal filePath As String, ByVal textValue As String)
    On Error GoTo ErrHandler
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, textValue;
    Close #fileNum
    fileNum = 0
    Exit Sub
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
    Err.Raise savedErrNumber, "PartCWriteTextFile", "Could not write temporary Access macro source file." & vbCrLf & filePath & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub PartCInvalidatePersistedPhysicalMacroState()
    On Error Resume Next
    Call EnsureDaoPropertyValue(currentDb, SYSTEM_MACRO_SET_FINGERPRINT_PROPERTY, dbText, vbNullString, "PartCInvalidatePersistedPhysicalMacroState")
    Call EnsureDaoPropertyValue(currentDb, SYSTEM_MACRO_EXPECTED_COUNT_PROPERTY, dbText, vbNullString, "PartCInvalidatePersistedPhysicalMacroState")
    Call EnsureDaoPropertyValue(currentDb, SYSTEM_MACRO_INSTALLED_INVENTORY_TOKEN_PROPERTY, dbText, vbNullString, "PartCInvalidatePersistedPhysicalMacroState")
    On Error GoTo 0
End Sub

' ===============================================================================================================================================================================
' SETUP PREFLIGHT / EVIDENCE / SETTINGS / RECOVERY-GATE PROCEDURES
' ===============================================================================================================================================================================

Private Sub ReleaseSetupPreflightNamingSnapshot(Optional ByVal invalidateCommonSnapshot As Boolean = True)
    If mSetupPreflightNamingSnapshotReady And invalidateCommonSnapshot And NamingSnapshotIsReady() Then InvalidateNamingSnapshot
    mSetupPreflightNamingSnapshotReady = False
    mSetupPreflightNamingSnapshotTraceRunID = vbNullString
    mSetupPreflightNamingSnapshotHostDbName = vbNullString
End Sub

Private Function SetupPreflightNamingSnapshotCurrent() As Boolean
    Dim activeTraceRunID As String
    If mSetupPreflightNamingSnapshotReady = False Or NamingSnapshotIsReady() = False Then Exit Function
    activeTraceRunID = GetActiveOrLastTraceRunID()
    If Len(activeTraceRunID) = 0 Or StrComp(activeTraceRunID, mSetupPreflightNamingSnapshotTraceRunID, vbBinaryCompare) <> 0 Then Exit Function
    If StrComp(currentDb.Name, mSetupPreflightNamingSnapshotHostDbName, vbTextCompare) <> 0 Then Exit Function
    SetupPreflightNamingSnapshotCurrent = True
End Function

' Purpose: Performs the public core begin setup naming snapshot operation using the routine's existing ADC workflow and side effects.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub CoreBeginSetupNamingSnapshot()
    If mSetupNamingSnapshotStarted Then Exit Sub
    On Error GoTo ErrHandler
    Dim reusedPreflightNamingSnapshot As Boolean
    BeginManagedObjectDefinitionSnapshot
    BeginSystemMacroDefinitionSnapshot
    If SetupPreflightNamingSnapshotCurrent() Then
        reusedPreflightNamingSnapshot = True
        ReleaseSetupPreflightNamingSnapshot False
        CoreAddOperationNote "A1 adopted the exact trace/HostDb-bound naming snapshot established by read-only Module Guard Preflight; duplicate naming-policy and managed-name reconstruction was skipped."
    Else
        If mSetupPreflightNamingSnapshotReady Then ReleaseSetupPreflightNamingSnapshot True
        BeginNamingSnapshot
    End If
    mSetupNamingSnapshotStarted = True
    mSetupNamingSnapshotDirty = False
    mSetupNamingPolicyDirty = False
    CoreAddOperationNote "Started the setup-scoped pre-resolved naming policy, managed-name index, and operation-scoped managed-object and System Macro definition indexes."
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    EndSystemMacroDefinitionSnapshot
    EndManagedObjectDefinitionSnapshot
    If reusedPreflightNamingSnapshot Then InvalidateNamingSnapshot
    If mSetupPreflightNamingSnapshotReady Then ReleaseSetupPreflightNamingSnapshot True
    On Error GoTo 0
    Err.Raise savedErrNumber, "CoreBeginSetupNamingSnapshot", "Could not initialize the setup-scoped naming and managed-object definition snapshots." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

' Purpose: Performs core mark setup naming snapshot dirty by coordinating `End` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   sourceDetail (Input Optional; String): Receives or supplies source detail used for caller diagnostics, verification output, or user-facing reporting. Default: vbNullString.
'   namingPolicyChanged (Input Optional; Boolean): Boolean state describing whether naming policy changed applies to this invocation; True selects the named condition/path and False selects its
'   complementary/default path. Default: False.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub CoreMarkSetupNamingSnapshotDirty(Optional ByVal sourceDetail As String = vbNullString, Optional ByVal namingPolicyChanged As Boolean = False)
    If mSetupNamingSnapshotStarted = False Then Exit Sub
    If mSetupNamingSnapshotDirty = False Then
        mSetupNamingSnapshotDirty = True
        If Len(Trim$(sourceDetail)) > 0 Then CoreAddOperationNote "Managed-name snapshot marked dirty by " & sourceDetail & "."
    End If
    If namingPolicyChanged Then mSetupNamingPolicyDirty = True
End Sub

' Purpose: Performs the public core ensure setup naming snapshot current operation using the routine's existing ADC workflow and side effects.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub CoreEnsureSetupNamingSnapshotCurrent()
    If mSetupNamingSnapshotStarted = False Then
        CoreBeginSetupNamingSnapshot
        Exit Sub
    End If
    If mSetupNamingSnapshotDirty Then
        If mSetupNamingPolicyDirty Then
            InvalidateNamingSnapshot
            BeginNamingSnapshot
        Else
            RefreshManagedNameSnapshot
        End If
        mSetupNamingSnapshotDirty = False
        mSetupNamingPolicyDirty = False
    End If
End Sub

Private Function StatusTextHasPartOfContext(ByVal statusText As String) As Boolean
    Dim normalizedText As String
    Dim partPosition As Long
    Dim scanPosition As Long
    Dim firstDigitStart As Long
    Dim firstDigitEnd As Long
    Dim totalDigitStart As Long
    Dim totalDigitEnd As Long
    normalizedText = " " & LCase$(statusText)
    partPosition = InStr(1, normalizedText, " part ", vbBinaryCompare)
    Do While partPosition > 0
        PumpSetupResponsiveness
        scanPosition = partPosition + 6
        Do While scanPosition <= Len(normalizedText) And Mid$(normalizedText, scanPosition, 1) = " "
            PumpSetupResponsiveness
            scanPosition = scanPosition + 1
        Loop
        firstDigitStart = scanPosition
        Do While scanPosition <= Len(normalizedText) And Mid$(normalizedText, scanPosition, 1) Like "#"
            PumpSetupResponsiveness
            scanPosition = scanPosition + 1
        Loop
        firstDigitEnd = scanPosition
        Do While scanPosition <= Len(normalizedText) And Mid$(normalizedText, scanPosition, 1) = " "
            PumpSetupResponsiveness
            scanPosition = scanPosition + 1
        Loop
        If firstDigitEnd > firstDigitStart And Mid$(normalizedText, scanPosition, 2) = "of" Then
            scanPosition = scanPosition + 2
            Do While scanPosition <= Len(normalizedText) And Mid$(normalizedText, scanPosition, 1) = " "
                PumpSetupResponsiveness
                scanPosition = scanPosition + 1
            Loop
            totalDigitStart = scanPosition
            Do While scanPosition <= Len(normalizedText) And Mid$(normalizedText, scanPosition, 1) Like "#"
                PumpSetupResponsiveness
                scanPosition = scanPosition + 1
            Loop
            totalDigitEnd = scanPosition
            If totalDigitEnd > totalDigitStart Then
                StatusTextHasPartOfContext = True
                Exit Function
            End If
        End If
        partPosition = InStr(partPosition + 6, normalizedText, " part ", vbBinaryCompare)
    Loop
End Function

Private Sub CoreSetVisibleStatusText(ByVal statusText As String)
    On Error Resume Next
    statusText = Trim$(statusText)
    If Len(statusText) = 0 Then Exit Sub
    SetVisibleStatusText statusText, True, False
    On Error GoTo 0
End Sub

' Purpose: Performs the public core pulse module guard preflight operation using the routine's existing ADC workflow and side effects.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   statusText (Input; String): Status/progress message supplied for the current operation; the routine forwards it to the applicable UI/progress and trace/reporting path.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub CorePulseModuleGuardPreflight(ByVal statusText As String)
    'Publishes Module Guard detail only while the owned setup progress/controller context is active.
    If mSetupProgressActive = False And mSetupControlActive = False Then Exit Sub
    PulseSetupModuleGuardPreflight statusText
End Sub

Private Sub PulseSetupModuleGuardPreflight(ByVal statusText As String)
    'Publishes the setup-only Module Guard preflight outside the numbered Setup Part sequence.
    Dim visibleText As String
    statusText = Trim$(statusText)
    If Len(statusText) = 0 Then Exit Sub
    If mSetupProgressActive Then
        visibleText = ProgressPartPrefixText() & "Setup Preflight > " & statusText
    Else
        visibleText = "Setup Preflight > " & statusText
    End If
    TrySharedProgressPulse mSetupProgressActive, visibleText, ProgressStatusPrefixText(), "PulseSetupModuleGuardPreflight", "Core", "Core"
End Sub

Private Function TryResolveModuleUpdateGuardForSetup(ByRef failedErrNumber As Long, ByRef failedErrSource As String, ByRef failedErrDescription As String, Optional ByRef verificationDetails As String = vbNullString) As Boolean
    On Error GoTo ErrHandler
    Dim missingTableNames As String
    Dim registrationResult As TModuleGuardRegistrationResult
    Dim registrationContext As String
    Dim verifyDetails As String
    Dim infrastructureExists As Boolean
    Dim manifestPopulated As Boolean
    Dim guardEnabled As Boolean
    Dim approvalPromptRequired As Boolean
    Dim approvalPromptText As String
    Dim bootstrapRegistration As Boolean
    Dim registrationEnableAfterRegistration As Boolean
    Dim registrationOptions As ModuleGuardRegistrationOptions
    Dim response As VbMsgBoxResult
    Dim namingSnapshotTimer As Single
    Dim readinessTimer As Single
    Dim cleanupTimer As Single
    Dim cleanupErrNumber As Long
    Dim cleanupErrDescription As String
    Dim returnDetails As String
    Dim namingSnapshotStarted As Boolean
    Dim namingSnapshotReusable As Boolean
    failedErrNumber = 0
    failedErrSource = vbNullString
    failedErrDescription = vbNullString
    verificationDetails = vbNullString
    ReleaseSetupPreflightNamingSnapshot False
    namingSnapshotTimer = Timer
    BeginNamingSnapshot
    namingSnapshotStarted = True
    namingSnapshotReusable = True
    LogTraceMessage TRACE_LEVEL_INFO, "ModuleUpdateGuard", "TryResolveModuleUpdateGuardForSetup", "Module Guard preflight naming snapshot ready.", "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(namingSnapshotTimer)), "Setup Preflight", 0, vbNullString, vbNullString, vbNullString, "NamingSnapshot", "CodeDb", "ModuleUpdateGuard"
    readinessTimer = Timer
    ResolveModuleUpdateGuardReadiness infrastructureExists, manifestPopulated, missingTableNames
    LogTraceMessage _
        TRACE_LEVEL_INFO, _
        "ModuleUpdateGuard", _
        "TryResolveModuleUpdateGuardForSetup", _
        "Module Guard readiness resolved once for Setup.", _
        "Infrastructure ready: " & CStr(infrastructureExists) & vbCrLf & "Active manifest populated: " & CStr(manifestPopulated) & vbCrLf & "Missing tables: " & IIf(Len(missingTableNames) = 0, "(none)", missingTableNames) & vbCrLf & "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(readinessTimer)), _
        "Setup Preflight", _
        0, _
        vbNullString, _
        vbNullString, _
        vbNullString, _
        "Readiness", _
        "CodeDb", _
        "ModuleUpdateGuard"
    If infrastructureExists = False Or manifestPopulated = False Then
        CorePulseProgress "Bootstrapping Module Guard infrastructure and approved coordinated-build registration"
        namingSnapshotReusable = False
        If TryRegisterApprovedBuild(registrationResult, MODULE_GUARD_REGISTRATION_OPTION_ENABLE_AFTER_REGISTRATION) Then
            verificationDetails = "Module Guard state: first-run/bootstrap registration completed and its registration-owned exact-build verification succeeded without forcing enablement." & vbCrLf & registrationResult.RegistrationDetails
            TryResolveModuleUpdateGuardForSetup = True
            GoTo CleanExit
        End If
        failedErrNumber = registrationResult.failedErrNumber
        failedErrSource = registrationResult.failedErrSource
        failedErrDescription = registrationResult.failedErrDescription
        If registrationResult.ExplicitApprovalCandidate = False Then GoTo CleanExit
        bootstrapRegistration = True
        approvalPromptRequired = True
        LogModuleGuardSetupTransition "TryResolveModuleUpdateGuardForSetup", "Strict bootstrap registration found an installed-build identity mismatch eligible for explicit approval.", "Strict embedded-build registration failed before any manifest write." & vbCrLf & "Err.Number: " & CStr(failedErrNumber) & vbCrLf & "Err.Source: " & failedErrSource & vbCrLf & "Err.Description: " & failedErrDescription, TRACE_LEVEL_WARNING
    Else
        guardEnabled = ModuleUpdateGuardEnabled(True)
        If guardEnabled Then
            CorePulseProgress "Verifying enabled Module Guard coordinated build"
        Else
            CorePulseProgress "Verifying coordinated build before disabled Module Guard choice"
        End If
        If TryPrepareAndVerifyModuleUpdateGuard(failedErrNumber, failedErrSource, failedErrDescription, verifyDetails, True, infrastructureExists, manifestPopulated, missingTableNames) = False Then approvalPromptRequired = True
    End If
    If approvalPromptRequired Then
        If bootstrapRegistration Then
            approvalPromptText = "Module Guard does not yet have an active approved manifest, and the installed ADC module set does not match the embedded approved coordinated build." & vbCrLf & vbCrLf
        Else
            approvalPromptText = "The installed ADC module set does not match the currently approved Module Guard build." & vbCrLf & vbCrLf
        End If
        ReportBlockingPromptWait "TryResolveModuleUpdateGuardForSetup", "choosing whether to validate and register the complete installed coordinated build after Module Guard verification or strict bootstrap registration failed"
        response = UtilityMsgBox(approvalPromptText & _
            "Yes = Validate and register the complete installed coordinated build as approved, then continue Setup if validation succeeds." & vbCrLf & _
            "No = Stop Setup without changing the approved build." & vbCrLf & _
            "Cancel = Cancel Setup without changing the approved build." & vbCrLf & vbCrLf & _
            "Choose Yes only when all installed ADC modules came from one complete reviewed package." & vbCrLf & vbCrLf & _
            "Detailed validation information was written to TraceLog and the Access Development Console trace log file.", vbQuestion Or vbYesNoCancel Or vbDefaultButton3, DEFAULT_UTILITY_LONG_NAME & ": Register Installed ADC Build?")
        LogModuleGuardSetupTransition "TryResolveModuleUpdateGuardForSetup", "Module Guard installed-build approval choice returned.", "Registration path: " & IIf(bootstrapRegistration, "bootstrap/unpopulated manifest", "populated manifest mismatch") & vbCrLf & "Response: " & MsgBoxResultText(response) & " (" & CStr(response) & ")"
        ReportBlockingPromptResult "TryResolveModuleUpdateGuardForSetup", "choosing whether to validate and register the complete installed coordinated build after Module Guard verification or strict bootstrap registration failed", (response <> vbCancel)
        If response <> vbYes Then
            failedErrNumber = ERR_SETUP_CANCELLED_BY_USER
            failedErrSource = "TryResolveModuleUpdateGuardForSetup"
            If response = vbCancel Then
                failedErrDescription = DEFAULT_UTILITY_LONG_NAME & ": Setup - CodeDb cancelled by user at the Module Guard registration prompt. The approved build and Module Guard enabled/disabled state were not changed. No material Setup part had started."
            Else
                failedErrDescription = DEFAULT_UTILITY_LONG_NAME & ": Setup - CodeDb stopped after the Module Guard registration prompt returned No. The approved build and Module Guard enabled/disabled state were not changed. No material Setup part had started."
            End If
            GoTo CleanExit
        End If
        If bootstrapRegistration Then
            registrationEnableAfterRegistration = True
        Else
            registrationEnableAfterRegistration = guardEnabled
        End If
        LogModuleGuardSetupTransition "TryResolveModuleUpdateGuardForSetup", "User chose explicit installed-build validation.", "Registration path: " & IIf(bootstrapRegistration, "bootstrap/unpopulated manifest", "populated manifest mismatch")
        CorePulseProgress "Validating and registering the explicitly approved complete coordinated build"
        namingSnapshotReusable = False
        registrationOptions = MODULE_GUARD_REGISTRATION_OPTION_EXPLICITLY_APPROVE_INSTALLED_BUILD
        If registrationEnableAfterRegistration Then registrationOptions = registrationOptions Or MODULE_GUARD_REGISTRATION_OPTION_ENABLE_AFTER_REGISTRATION
        If TryRegisterApprovedBuild(registrationResult, registrationOptions) = False Then
            failedErrNumber = registrationResult.failedErrNumber
            failedErrSource = registrationResult.failedErrSource
            failedErrDescription = registrationResult.failedErrDescription
            LogModuleGuardSetupTransition "TryResolveModuleUpdateGuardForSetup", "Explicit installed-build validation/registration failed.", "Registration path: " & IIf(bootstrapRegistration, "bootstrap/unpopulated manifest", "populated manifest mismatch") & vbCrLf & "Err.Number: " & CStr(failedErrNumber) & vbCrLf & "Err.Source: " & failedErrSource & vbCrLf & "Err.Description: " & failedErrDescription, TRACE_LEVEL_ERROR
            GoTo CleanExit
        End If
        LogModuleGuardSetupTransition "TryResolveModuleUpdateGuardForSetup", "Explicit installed-build validation/registration succeeded.", "Registration path: " & IIf(bootstrapRegistration, "bootstrap/unpopulated manifest", "populated manifest mismatch")
        If bootstrapRegistration Then
            verificationDetails = "Module Guard state: first-run/bootstrap installed build was explicitly approved, validated, registered, and transactionally verified." & vbCrLf & registrationResult.RegistrationDetails
            TryResolveModuleUpdateGuardForSetup = True
            GoTo CleanExit
        End If
        verifyDetails = "The registration authority completed its transaction-owned exact-build verification successfully."
        registrationContext = "Registration was explicitly approved through the Module Guard installed-build approval prompt." & vbCrLf & registrationResult.RegistrationDetails & vbCrLf & vbCrLf
    End If
    If guardEnabled Then
        If Len(registrationContext) > 0 Then
            verificationDetails = "Module Guard state: the complete installed build was explicitly registered, verified, and left enabled." & vbCrLf & registrationContext & verifyDetails
        Else
            verificationDetails = "Module Guard state: installed and enabled." & vbCrLf & verifyDetails
        End If
        TryResolveModuleUpdateGuardForSetup = True
        GoTo CleanExit
    End If
    If SettingRowExists(SETTING_ENABLE_MODULE_UPDATE_GUARD) = False Then
        LogModuleGuardSetupTransition "TryResolveModuleUpdateGuardForSetup", "Disabled Module Guard re-enable choice deferred.", "The canonical EnableModuleUpdateGuard setting row is not available yet; A2 owns that Settings bootstrap contract, so Setup will continue with the verified disabled state."
        verificationDetails = "Module Guard state: installed and verified; the re-enable choice was not offered because the A2-owned EnableModuleUpdateGuard setting row is not yet available. The verified disabled state was retained so Setup can continue to the Settings bootstrap owner." & vbCrLf & registrationContext & verifyDetails
        TryResolveModuleUpdateGuardForSetup = True
        GoTo CleanExit
    End If
    ReportBlockingPromptWait "TryResolveModuleUpdateGuardForSetup", "choosing whether to re-enable the verified disabled Module Guard"
    response = UtilityMsgBox("Module Guard is installed and disabled. The installed coordinated build has already been verified." & vbCrLf & vbCrLf & _
        "Yes = Re-enable Module Guard and continue." & vbCrLf & _
        "No = Leave Module Guard disabled and continue." & vbCrLf & _
        "Cancel = Cancel Setup before material work begins.", vbQuestion Or vbYesNoCancel Or vbDefaultButton2, DEFAULT_UTILITY_LONG_NAME & ": Module Guard")
    LogModuleGuardSetupTransition "TryResolveModuleUpdateGuardForSetup", "Disabled Module Guard choice returned.", "Response: " & MsgBoxResultText(response) & " (" & CStr(response) & ")"
    ReportBlockingPromptResult "TryResolveModuleUpdateGuardForSetup", "choosing whether to re-enable the verified disabled Module Guard", (response <> vbCancel)
    If response = vbCancel Then
        failedErrNumber = ERR_SETUP_CANCELLED_BY_USER
        failedErrSource = "TryResolveModuleUpdateGuardForSetup"
        failedErrDescription = DEFAULT_UTILITY_LONG_NAME & ": Setup - CodeDb cancelled by user while resolving the disabled Module Guard state. No material Setup part had started."
        GoTo CleanExit
    End If
    If response = vbYes Then
        LogModuleGuardSetupTransition "TryResolveModuleUpdateGuardForSetup", "Module Guard re-enable starting.", "Requested state: enabled"
        namingSnapshotReusable = False
        SetModuleUpdateGuardEnabledForSetup True
        LogModuleGuardSetupTransition "TryResolveModuleUpdateGuardForSetup", "Module Guard re-enable completed.", "Resulting state: enabled"
        verificationDetails = "Module Guard state: installed, verified, and explicitly re-enabled by the user." & vbCrLf & registrationContext & verifyDetails
    Else
        LogModuleGuardSetupTransition "TryResolveModuleUpdateGuardForSetup", "Module Guard re-enable skipped.", "The user selected " & MsgBoxResultText(response) & "; the verified disabled state was retained."
        verificationDetails = "Module Guard state: installed, verified, and deliberately left disabled by the user." & vbCrLf & registrationContext & verifyDetails
    End If
    TryResolveModuleUpdateGuardForSetup = True
    LogModuleGuardSetupTransition "TryResolveModuleUpdateGuardForSetup", "Module Guard state resolution result prepared.", "Result: success"
CleanExit:
    On Error Resume Next
    If namingSnapshotStarted Then
        If TryResolveModuleUpdateGuardForSetup And namingSnapshotReusable And NamingSnapshotIsReady() Then
            mSetupPreflightNamingSnapshotReady = True
            mSetupPreflightNamingSnapshotTraceRunID = GetActiveOrLastTraceRunID()
            mSetupPreflightNamingSnapshotHostDbName = currentDb.Name
            namingSnapshotStarted = False
            LogModuleGuardSetupTransition "TryResolveModuleUpdateGuardForSetup", "Module Guard preflight naming snapshot retained for A1 transfer.", "The exact read-only naming evidence remains COMMON-owned and trace/HostDb-bound until A1 adopts it or a consumer invalidates it."
        Else
            cleanupTimer = Timer
            LogModuleGuardSetupTransition "TryResolveModuleUpdateGuardForSetup", "Module Guard preflight naming snapshot cleanup starting.", "Snapshot ownership: setup preflight"
            Err.Clear
            InvalidateNamingSnapshot
            ReleaseSetupPreflightNamingSnapshot False
            cleanupErrNumber = Err.Number
            cleanupErrDescription = Err.Description
            Err.Clear
            If cleanupErrNumber = 0 Then
                namingSnapshotStarted = False
                LogModuleGuardSetupTransition "TryResolveModuleUpdateGuardForSetup", "Module Guard preflight naming snapshot cleanup completed.", "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(cleanupTimer))
            Else
                LogModuleGuardSetupTransition "TryResolveModuleUpdateGuardForSetup", "Module Guard preflight naming snapshot cleanup raised an error that remains non-blocking under the existing cleanup policy.", "Err.Number: " & CStr(cleanupErrNumber) & vbCrLf & "Err.Description: " & cleanupErrDescription, TRACE_LEVEL_WARNING
            End If
        End If
    Else
        LogModuleGuardSetupTransition "TryResolveModuleUpdateGuardForSetup", "Module Guard preflight naming snapshot cleanup skipped.", "No setup-preflight naming snapshot remained owned."
    End If
    If TryResolveModuleUpdateGuardForSetup Then
        returnDetails = "Result: success"
    Else
        returnDetails = "Result: failure or cancellation" & vbCrLf & "Err.Number: " & CStr(failedErrNumber) & vbCrLf & "Err.Source: " & failedErrSource & vbCrLf & "Err.Description: " & failedErrDescription
    End If
    LogModuleGuardSetupTransition "TryResolveModuleUpdateGuardForSetup", "Module Guard state resolution returning to caller.", returnDetails
    FlushTraceLogFile
    On Error GoTo 0
    Exit Function
ErrHandler:
    failedErrNumber = Err.Number
    failedErrSource = Err.Source
    failedErrDescription = Err.Description
    If failedErrNumber = 0 Then failedErrNumber = vbObjectError + 6293
    If Len(Trim$(failedErrSource)) = 0 Then failedErrSource = "TryResolveModuleUpdateGuardForSetup"
    If Len(Trim$(failedErrDescription)) = 0 Then failedErrDescription = "Module Guard state resolution did not complete."
    Resume CleanExit
End Function

Private Sub LogModuleGuardSetupTransition(ByVal sourceProcedure As String, ByVal messageText As String, ByVal detailText As String, Optional ByVal traceLevel As String = TRACE_LEVEL_INFO)
    If Len(Trim$(detailText)) = 0 Then detailText = messageText
    LogTraceMessage traceLevel, "ModuleUpdateGuard", sourceProcedure, messageText, detailText, "Setup Preflight", 0, vbNullString, vbNullString, vbNullString, "Transition", "CodeDb", "ModuleUpdateGuard"
End Sub

' Purpose: Performs report blocking prompt wait by coordinating `CorePulseProgress` and `LogTraceMessage` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   sourceProcedure (Input; String): Procedure name recorded as the source of trace, error, status, or diagnostic information emitted by this call.
'   promptReason (Input; String): Prompt reason text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub ReportBlockingPromptWait(ByVal sourceProcedure As String, ByVal promptReason As String)
    'Leaves a visible status/trace marker while an Access confirmation dialog may be waiting behind another window.
    Dim statusText As String
    statusText = "Checking open objects; confirmation may be required: " & promptReason
    On Error Resume Next
    CorePulseProgress statusText
    LogTraceMessage TRACE_LEVEL_STATUS, "PromptWait", sourceProcedure, statusText, "A modal confirmation may remain open until the user responds.", statusText, 0, vbNullString, vbNullString, vbNullString, "Prompt", "CodeDb", "Setup"
    On Error GoTo 0
End Sub

' Purpose: Performs report blocking prompt result by coordinating `End` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   sourceProcedure (Input; String): Procedure name recorded as the source of trace, error, status, or diagnostic information emitted by this call.
'   promptReason (Input; String): Prompt reason text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   approved (Input; Boolean): Approval state to persist/apply for the reviewed item: True marks it approved; False marks or leaves it unapproved.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub ReportBlockingPromptResult(ByVal sourceProcedure As String, ByVal promptReason As String, ByVal approved As Boolean)
    Dim resultText As String
    resultText = IIf(approved, "Open-object confirmation completed or was not required.", "Open-object confirmation was declined or could not complete.")
    On Error Resume Next
    LogTraceMessage IIf(approved, TRACE_LEVEL_INFO, TRACE_LEVEL_WARNING), "PromptWait", sourceProcedure, resultText, "Reason: " & promptReason, resultText, 0, vbNullString, vbNullString, vbNullString, "Prompt", "CodeDb", "Setup"
    On Error GoTo 0
End Sub

Private Sub LogSetupRecoveryPreflightStage(ByVal messageText As String, ByVal detailText As String, Optional ByVal traceLevel As String = TRACE_LEVEL_INFO)
    If Len(Trim$(detailText)) = 0 Then detailText = messageText
    LogTraceMessage traceLevel, "SetupRecovery", "RequireNoUnfinishedSetupRecoveryState", messageText, detailText, "Setup Preflight", 0, vbNullString, vbNullString, vbNullString, "RecoveryState", "CodeDb", "Setup"
    FlushTraceLogFile
End Sub

Private Sub RequireNoUnfinishedSetupRecoveryState(Optional ByVal resumableSetupOperationName As String = vbNullString)
    On Error GoTo ErrHandler
    Dim recoverySnapshot As Object
    Dim recoveryOperation As String
    Dim recoveryDetail As String
    Dim recoveryPayload As String
    Dim uiDetail As String
    Dim stageName As String
    Dim stageTimer As Single
    Dim coreRollbackJournalAdopted As Boolean
    stageName = "Build durable recovery snapshot"
    stageTimer = Timer
    LogSetupRecoveryPreflightStage "Setup recovery-state snapshot starting.", "Physical authority: the existing Setup_RecoveryState managed table. One read-only snapshot will be shared by the existing Core, Catalog, Review, and UI recovery gates."
    Set recoverySnapshot = CoreDurableRecoverySnapshot()
    LogSetupRecoveryPreflightStage "Setup recovery-state snapshot completed.", "Blocked markers loaded: " & CStr(recoverySnapshot.count) & vbCrLf & "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(stageTimer))
    stageName = "Core rollback marker"
    stageTimer = Timer
    LogSetupRecoveryPreflightStage "Core rollback recovery check starting.", "Recovery key remains owned by Core."
    If CoreDurableRecoveryMarkerExistsInSnapshot(recoverySnapshot, SETUP_RECOVERY_KEY, recoveryOperation, recoveryDetail, recoveryPayload) Then
        coreRollbackJournalAdopted = TryAdoptSetupRollbackRecoveryForResume(resumableSetupOperationName, recoveryOperation, recoveryDetail, recoveryPayload)
        If coreRollbackJournalAdopted = False Then Err.Raise ERR_DURABLE_RECOVERY_FAILED, "RequireNoUnfinishedSetupRecoveryState", "Setup cannot begin while a prior Core rollback recovery marker remains unresolved." & vbCrLf & recoveryOperation & vbCrLf & recoveryDetail
    End If
    LogSetupRecoveryPreflightStage "Core rollback recovery check completed.", "Blocked: False or matching full-Setup rollback journal adopted." & vbCrLf & "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(stageTimer))
    stageName = "Core post-commit marker"
    stageTimer = Timer
    LogSetupRecoveryPreflightStage "Core post-commit recovery check starting.", "Recovery key remains owned by Core."
    If CoreDurableRecoveryMarkerExistsInSnapshot(recoverySnapshot, SETUP_POST_COMMIT_CHECKPOINT_KEY, recoveryOperation, recoveryDetail, recoveryPayload) Then Err.Raise ERR_DURABLE_RECOVERY_FAILED, "RequireNoUnfinishedSetupRecoveryState", "Setup cannot begin while a prior post-commit checkpoint remains unresolved." & vbCrLf & recoveryOperation & vbCrLf & recoveryDetail
    LogSetupRecoveryPreflightStage "Core post-commit recovery check completed.", "Blocked: False" & vbCrLf & "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(stageTimer))
    stageName = "Catalog recovery"
    stageTimer = Timer
    LogSetupRecoveryPreflightStage "Catalog recovery check starting.", "The existing Catalog recovery authority will inspect the shared snapshot."
    If CatalogRecoveryBlocked(recoverySnapshot) Then Err.Raise ERR_DURABLE_RECOVERY_FAILED, "RequireNoUnfinishedSetupRecoveryState", "Setup cannot begin while Catalog recovery remains unresolved. Run the dedicated recovery command first."
    LogSetupRecoveryPreflightStage "Catalog recovery check completed.", "Blocked: False" & vbCrLf & "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(stageTimer))
    stageName = "candidate-marking recovery"
    stageTimer = Timer
    LogSetupRecoveryPreflightStage "candidate-marking recovery check starting.", "The existing candidate-marking recovery authority will inspect the shared snapshot."
    If ConfigCandidateMarkingRecoveryBlocked(recoverySnapshot) Then Err.Raise ERR_DURABLE_RECOVERY_FAILED, "RequireNoUnfinishedSetupRecoveryState", "Setup cannot begin while candidate-marking recovery remains unresolved. Run the dedicated recovery command first."
    LogSetupRecoveryPreflightStage "candidate-marking recovery check completed.", "Blocked: False" & vbCrLf & "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(stageTimer))
    stageName = "UI and ribbon recovery"
    stageTimer = Timer
    LogSetupRecoveryPreflightStage "UI and ribbon recovery check starting.", "The existing UI recovery authority will inspect its in-memory state and the shared snapshot."
    If UIRecoveryBlockedForSetupPreflight(uiDetail, recoverySnapshot) Then
        If coreRollbackJournalAdopted Then
            LogSetupRecoveryPreflightStage "UI recovery retry starting for adopted full Setup rollback journal.", "A matching interrupted full-Setup Core journal was adopted. Resolve its dependency-owned UI recovery boundary before any new Setup mutation begins."
            RetryUIRecovery False
            uiDetail = vbNullString
            If UIRecoveryBlockedForSetupPreflight(uiDetail) Then Err.Raise ERR_UI_PREFLIGHT_RECOVERY_FAILED, "RequireNoUnfinishedSetupRecoveryState", "Setup cannot resume because UI/ribbon recovery remained unresolved after the dependency-owned retry." & IIf(Len(Trim$(uiDetail)) = 0, vbNullString, vbCrLf & uiDetail)
            LogSetupRecoveryPreflightStage "UI recovery retry completed for adopted full Setup rollback journal.", "Blocked: False"
        Else
            Err.Raise ERR_UI_PREFLIGHT_RECOVERY_FAILED, "RequireNoUnfinishedSetupRecoveryState", "Setup cannot begin while UI/ribbon recovery remains unresolved. Setup Preflight does not repair an unrelated later-domain recovery state." & IIf(Len(Trim$(uiDetail)) = 0, vbNullString, vbCrLf & uiDetail)
        End If
    End If
    LogSetupRecoveryPreflightStage "UI and ribbon recovery check completed.", "Blocked: False" & vbCrLf & "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(stageTimer))
    LogSetupRecoveryPreflightStage "Setup recovery-state validation completed.", "All existing recovery authorities returned ready."
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    LogSetupRecoveryPreflightStage "Setup recovery-state validation failed.", "Stage: " & stageName & vbCrLf & "Elapsed in stage: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(stageTimer)) & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, TRACE_LEVEL_ERROR
    Err.Raise savedErrNumber, savedErrSource, savedErrDescription
End Sub

' Purpose: Persists a fail-closed recovery marker through an independent DAO workspace so the marker is not part of the failing transaction.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   recoveryKey (Input; String): Canonical key identifying the recovery definition, setting, or managed item to use.
'   operationName (Input; String): Logical operation name written to progress, trace, warning, or error context for this call.
'   detailText (Input; String): Receives or supplies diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
'   payloadText (Input Optional; String): Input payload text for persist durable recovery marker; is passed to `CoreRecoveryEncode`. Default: vbNullString.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub PersistDurableRecoveryMarker(ByVal recoveryKey As String, ByVal operationName As String, ByVal detailText As String, Optional ByVal payloadText As String = vbNullString)
    'Persists a fail-closed recovery marker through an independent DAO workspace so the marker is not part of the failing transaction.
    On Error GoTo ErrHandler
    Dim recoveryWs As DAO.workspace
    Dim recoveryDb As DAO.Database
    Dim rs As DAO.Recordset
    Dim notesText As String
    recoveryKey = Trim$(recoveryKey)
    If Len(recoveryKey) = 0 Then Err.Raise ERR_DURABLE_RECOVERY_FAILED, "PersistDurableRecoveryMarker", "The recovery key is blank."
    OpenCoreRecoveryDatabase recoveryWs, recoveryDb
    EnsureCoreRecoveryStorage recoveryDb
    notesText = "Version=1" & vbCrLf & "Operation=" & CoreRecoveryEncode(operationName) & vbCrLf & "Detail=" & CoreRecoveryEncode(detailText) & vbCrLf & "Payload=" & CoreRecoveryEncode(payloadText)
    Set rs = recoveryDb.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SetupRecoveryState") & "] WHERE " & SQL_FIELD_RECOVERY_KEY & "=" & SharedDao_SqlQuote(recoveryKey) & ";", dbOpenDynaset)
    If rs.EOF Then rs.AddNew Else rs.Edit
    rs.fields(FIELD_RECOVERY_KEY).value = recoveryKey
    rs.fields(FIELD_RECOVERY_STATUS).value = RECOVERY_STATUS_BLOCKED
    rs.fields(FIELD_UPDATED_ON).value = Now
    rs.fields(FIELD_MANAGED_OBJECT_DEPENDENCY_NOTES).value = notesText
    rs.Update
    rs.Close
    recoveryDb.Close
    recoveryWs.Close
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not recoveryDb Is Nothing Then recoveryDb.Close
    If Not recoveryWs Is Nothing Then recoveryWs.Close
    On Error GoTo 0
    Err.Raise ERR_DURABLE_RECOVERY_FAILED, "PersistDurableRecoveryMarker", "Could not persist the durable recovery marker. Key: " & recoveryKey & vbCrLf & "Original error " & CStr(savedErrNumber) & ": " & savedErrDescription
End Sub

Private Function CoreDurableRecoverySnapshot() As Object
    'Reads all blocked Setup_RecoveryState rows once so Setup Preflight can share one physical database snapshot across the existing domain recovery authorities.
    On Error GoTo ErrHandler
    Dim recoverySnapshot As Object
    Dim recoveryWs As DAO.workspace
    Dim recoveryDb As DAO.Database
    Dim rs As DAO.Recordset
    Dim recoveryKey As String
    Dim recoveryTableName As String
    ' INTENTIONAL OPERATION SNAPSHOT: recoverySnapshot captures one coherent source state for this invocation of CoreDurableRecoverySnapshot; validity ends with the operation rather than becoming a live data
    ' authority.
    Set recoverySnapshot = CreateObject("Scripting.Dictionary")
    recoverySnapshot.CompareMode = vbTextCompare
    recoveryTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SetupRecoveryState")
    OpenCoreRecoveryDatabase recoveryWs, recoveryDb
    If SharedDao_TableExists(recoveryDb, recoveryTableName) = False Then GoTo CleanExit
    Set rs = recoveryDb.OpenRecordset("SELECT " & SQL_FIELD_RECOVERY_KEY & "," & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_NOTES & " FROM [" & recoveryTableName & "] WHERE " & SQL_FIELD_RECOVERY_STATUS & "=" & SharedDao_SqlQuote(RECOVERY_STATUS_BLOCKED) & ";", dbOpenSnapshot)
    Do While rs.EOF = False
        recoveryKey = Trim$(CStr(Nz(rs.fields(FIELD_RECOVERY_KEY).value, vbNullString)))
        If Len(recoveryKey) > 0 Then recoverySnapshot.item(recoveryKey) = CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_DEPENDENCY_NOTES).value, vbNullString))
        rs.MoveNext
    Loop
CleanExit:
    If Not rs Is Nothing Then rs.Close
    If Not recoveryDb Is Nothing Then recoveryDb.Close
    If Not recoveryWs Is Nothing Then recoveryWs.Close
    Set CoreDurableRecoverySnapshot = recoverySnapshot
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
    If Not recoveryDb Is Nothing Then recoveryDb.Close
    If Not recoveryWs Is Nothing Then recoveryWs.Close
    On Error GoTo 0
    Err.Raise ERR_DURABLE_RECOVERY_FAILED, "CoreDurableRecoverySnapshot", "Could not build the read-only durable recovery snapshot." & vbCrLf & "Original error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription
End Function

' Purpose: Resolves core recovery encode from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   sourceText (Input; String): Input source text examined or transformed by this procedure according to its documented core recovery encode contract.
' Returns: String: Returns the core recovery encode text value. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function CoreRecoveryEncode(ByVal sourceText As String) As String
    'Percent-encodes recovery text so control characters and percent signs can be stored safely in line-oriented payloads.
    sourceText = Replace(sourceText, "%", "%25")
    sourceText = Replace(sourceText, Chr$(30), "%1E")
    sourceText = Replace(sourceText, vbCrLf, "%0D%0A")
    sourceText = Replace(sourceText, vbCr, "%0D")
    sourceText = Replace(sourceText, vbLf, "%0A")
    sourceText = Replace(sourceText, vbTab, "%09")
    CoreRecoveryEncode = sourceText
End Function

' Purpose: Resolves core recovery decode from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   sourceText (Input; String): Input source text examined or transformed by this procedure according to its documented core recovery decode contract.
' Returns: String: Returns the core recovery decode text value. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function CoreRecoveryDecode(ByVal sourceText As String) As String
    'Decodes percent-escaped recovery text back to its original control characters and percent signs.
    sourceText = Replace(sourceText, "%09", vbTab)
    sourceText = Replace(sourceText, "%0D%0A", vbCrLf)
    sourceText = Replace(sourceText, "%0D", vbCr)
    sourceText = Replace(sourceText, "%0A", vbLf)
    sourceText = Replace(sourceText, "%1E", Chr$(30))
    sourceText = Replace(sourceText, "%25", "%")
    CoreRecoveryDecode = sourceText
End Function

' Purpose: Resolves core recovery note value from the supplied/current ADC state, using `StrComp` and `Next` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   notesText (Input; String): Notes text text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   keyName (Input; String): Canonical or physical name identifying the key targeted by this operation.
' Returns: String: Returns the core recovery note value text value. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function CoreRecoveryNoteValue(ByVal notesText As String, ByVal keyName As String) As String
    'Finds a named key in line-oriented recovery notes and returns its decoded value.
    Dim lineItem As Variant
    Dim prefixText As String
    prefixText = keyName & "="
    For Each lineItem In Split(notesText, vbCrLf)
        PumpSetupResponsiveness
        If StrComp(Left$(CStr(lineItem), Len(prefixText)), prefixText, vbTextCompare) = 0 Then CoreRecoveryNoteValue = CoreRecoveryDecode(Mid$(CStr(lineItem), Len(prefixText) + 1)): Exit Function
    Next lineItem
End Function

' Purpose: Checks for core durable recovery marker exists in snapshot without requiring the caller to handle the routine's normal not-found path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   recoverySnapshot (Input; Object): Snapshot of recovery state captured before or during the operation and used for validation, rollback, or recovery decisions.
'   recoveryKey (Input; String): Canonical key identifying the recovery definition, setting, or managed item to use.
'   operationName (Output; String): Receives the operation name returned by `CoreRecoveryNoteValue` after this procedure completes its validation/resolution work.
'   detailText (Output; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
'   payloadText (Output; String): Receives the payload text returned by `CoreRecoveryNoteValue` after this procedure completes its validation/resolution work.
' Returns: Boolean: True when the condition represented by DeleteDurableRecoveryMarker is present under the routine's documented inputs/current state; False when that condition is absent.
Public Function CoreDurableRecoveryMarkerExistsInSnapshot(ByVal recoverySnapshot As Object, ByVal recoveryKey As String, ByRef operationName As String, ByRef detailText As String, ByRef payloadText As String) As Boolean
    'Interprets one owner-supplied recovery key from the Core-built Setup_RecoveryState snapshot; ownership of the key and blocking rule remains with the calling domain.
    Dim notesText As String
    operationName = vbNullString
    detailText = vbNullString
    payloadText = vbNullString
    If recoverySnapshot Is Nothing Then Err.Raise ERR_DURABLE_RECOVERY_FAILED, "CoreDurableRecoveryMarkerExistsInSnapshot", "The durable recovery snapshot is not available."
    recoveryKey = Trim$(recoveryKey)
    If Len(recoveryKey) = 0 Then Err.Raise ERR_DURABLE_RECOVERY_FAILED, "CoreDurableRecoveryMarkerExistsInSnapshot", "The recovery key is blank."
    If recoverySnapshot.Exists(recoveryKey) = False Then Exit Function
    notesText = CStr(recoverySnapshot.item(recoveryKey))
    operationName = CoreRecoveryNoteValue(notesText, "Operation")
    detailText = CoreRecoveryNoteValue(notesText, "Detail")
    payloadText = CoreRecoveryNoteValue(notesText, "Payload")
    CoreDurableRecoveryMarkerExistsInSnapshot = True
End Function

' Purpose: Reads one durable recovery marker without invoking Catalog or UI repair paths.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   recoveryKey (Input; String): Canonical key identifying the recovery definition, setting, or managed item to use.
'   operationName (Output; String): Receives the operation name returned by `CoreRecoveryNoteValue` after this procedure completes its validation/resolution work.
'   detailText (Output; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
'   payloadText (Output; String): Receives the payload text returned by `CoreRecoveryNoteValue` after this procedure completes its validation/resolution work.
' Returns: Boolean: True when core durable recovery marker exists in the inspected context; False when it does not.
Public Function CoreDurableRecoveryMarkerExists(ByVal recoveryKey As String, ByRef operationName As String, ByRef detailText As String, ByRef payloadText As String) As Boolean
    'Reads one durable recovery marker without invoking Catalog or UI repair paths.
    On Error GoTo ErrHandler
    Dim recoveryWs As DAO.workspace
    Dim recoveryDb As DAO.Database
    Dim rs As DAO.Recordset
    Dim notesText As String
    operationName = vbNullString
    detailText = vbNullString
    payloadText = vbNullString
    OpenCoreRecoveryDatabase recoveryWs, recoveryDb
    If SharedDao_TableExists(recoveryDb, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SetupRecoveryState")) = False Then GoTo CleanExit
    Set rs = recoveryDb.OpenRecordset("SELECT TOP 1 " & SQL_FIELD_RECOVERY_STATUS & "," & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_NOTES & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SetupRecoveryState") & "] WHERE " & SQL_FIELD_RECOVERY_KEY & "=" & SharedDao_SqlQuote(Trim$(recoveryKey)) & ";", dbOpenSnapshot)
    If rs.EOF Then GoTo CleanExit
    If StrComp(Trim$(CStr(Nz(rs.fields(FIELD_RECOVERY_STATUS).value, vbNullString))), RECOVERY_STATUS_BLOCKED, vbTextCompare) <> 0 Then GoTo CleanExit
    notesText = CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_DEPENDENCY_NOTES).value, vbNullString))
    operationName = CoreRecoveryNoteValue(notesText, "Operation")
    detailText = CoreRecoveryNoteValue(notesText, "Detail")
    payloadText = CoreRecoveryNoteValue(notesText, "Payload")
    CoreDurableRecoveryMarkerExists = True
CleanExit:
    If Not rs Is Nothing Then rs.Close
    recoveryDb.Close
    recoveryWs.Close
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not recoveryDb Is Nothing Then recoveryDb.Close
    If Not recoveryWs Is Nothing Then recoveryWs.Close
    On Error GoTo 0
    Err.Raise ERR_DURABLE_RECOVERY_FAILED, "CoreDurableRecoveryMarkerExists", "Could not inspect the durable recovery marker. Key: " & recoveryKey & vbCrLf & "Original error " & CStr(savedErrNumber) & ": " & savedErrDescription
End Function

' Purpose: Deletes one durable recovery marker only after its caller has verified recovery completion.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   recoveryKey (Input; String): Canonical key identifying the recovery definition, setting, or managed item to use.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub DeleteDurableRecoveryMarker(ByVal recoveryKey As String)
    'Deletes one durable recovery marker only after its caller has verified recovery completion.
    On Error GoTo ErrHandler
    Dim recoveryWs As DAO.workspace
    Dim recoveryDb As DAO.Database
    OpenCoreRecoveryDatabase recoveryWs, recoveryDb
    If SharedDao_TableExists(recoveryDb, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SetupRecoveryState")) Then recoveryDb.Execute "DELETE FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SetupRecoveryState") & "] WHERE " & SQL_FIELD_RECOVERY_KEY & "=" & SharedDao_SqlQuote(Trim$(recoveryKey)) & ";", dbFailOnError
    recoveryDb.Close
    recoveryWs.Close
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not recoveryDb Is Nothing Then recoveryDb.Close
    If Not recoveryWs Is Nothing Then recoveryWs.Close
    On Error GoTo 0
    Err.Raise ERR_DURABLE_RECOVERY_FAILED, "DeleteDurableRecoveryMarker", "Could not delete the durable recovery marker. Key: " & recoveryKey & vbCrLf & "Original error " & CStr(savedErrNumber) & ": " & savedErrDescription
End Sub

Private Function ExistingSettingBoolWithoutRepair(ByVal SettingKey As String, ByVal DefaultValue As Boolean) As Boolean
    On Error GoTo UseDefault
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim valueText As String
    Set db = ADCDb()
    If SharedDao_DatabaseTableExistsSafe(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID)) = False Then GoTo UseDefault
    If SharedDao_DatabaseTableFieldExistsSafe(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID), FIELD_SETTING_KEY) = False Then GoTo UseDefault
    If SharedDao_DatabaseTableFieldExistsSafe(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID), FIELD_SETTING_VALUE) = False Then GoTo UseDefault
    sqlText = "SELECT [" & FIELD_SETTING_VALUE & "] FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID) & "] WHERE [" & FIELD_SETTING_KEY & "]=" & SharedDao_SqlQuote(SettingKey) & ";"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    If rs.EOF Then GoTo UseDefault
    valueText = Nz(rs.fields(FIELD_SETTING_VALUE).value, vbNullString)
    rs.Close
    Set rs = Nothing
    On Error GoTo 0
    ExistingSettingBoolWithoutRepair = BooleanTextValue(valueText, DefaultValue)
    Exit Function
UseDefault:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    On Error GoTo 0
    ExistingSettingBoolWithoutRepair = DefaultValue
End Function

' Purpose: Performs the public clear settings runtime state for current context operation using the routine's existing ADC workflow and side effects.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub ClearSettingsRuntimeStateForCurrentContext()
    'Collections do not support removing a known keyed item safely without raising an error.
    'The state is cheap to rebuild, so clear the readiness cache when a schema repair/rebuild could have changed the underlying settings table object.
    Set mSettingsReadyKeys = New Collection
End Sub

' Purpose: Ensures the CodeDb Settings domain is structurally readable and default-complete for the current context, repairing the owned foundation as allowed and caching successful readiness.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub EnsureSettingsReadyForRead()
    'Owns Settings read-readiness state for all direct and cross-module callers. Generic readiness may perform B2-owned lookup mutation only outside numbered Setup or while B2 is active.
    On Error GoTo ErrHandler
    Dim settingsTableName As String
    Dim applyLookupMetadata As Boolean
    settingsTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID)
    applyLookupMetadata = SetupDomainMutationWritesAllowedForCurrentPart(SETUP_DOMAIN_B2_CATALOG_REFERENCE_DEFAULTS)
    If SharedCollection_KeyExists(mSettingsReadyKeys, StateKey()) Then
        If TableCanOpen(settingsTableName) Then Exit Sub
        'The marker was stale. Clear and repair the state for the CodeDb workflow.
        ClearSettingsRuntimeStateForCurrentContext
    End If
    RepairCodeDbSettingsTableForSetup applyLookupMetadata
    MarkSettingsReadyForCurrentContext
    RepairSettingDefaultsForSetup False
    RepairLauncherTableFoundationForSetup False, True, False, applyLookupMetadata
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If Len(settingsTableName) = 0 Then settingsTableName = SafeTableName("settings")
    Err.Raise savedErrNumber, "EnsureSettingsReadyForRead", _
        "Could not prepare macro-creator settings for read." & vbCrLf & _
        "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB & vbCrLf & _
        "CodeDb settings table: " & settingsTableName & vbCrLf & _
        "Err.Number: " & CStr(savedErrNumber) & vbCrLf & _
        "Err.Source: " & savedErrSource & vbCrLf & _
        "Err.Description: " & savedErrDescription
End Sub

' Purpose: Marks settings ready for current context.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub MarkSettingsReadyForCurrentContext()
    ' Wrapper purpose: Semantic public facade that marks the current Core context as settings-ready.
    ' Why retained: Retained so callers do not manipulate Core's private readiness collection or reproduce the StateKey and keyed-state update mechanics.
    'Intended use: Runs Mark Access Development Console Settings Ready For Current Context for direct, generated-macro, form-button, or required cross-module callers.
    MarkCollectionTextKey mSettingsReadyKeys, StateKey()
End Sub

Private Sub AddTraceStep(ByVal stepText As String)
    On Error Resume Next
    Dim partStepText As String
    partStepText = ProgressPartText(stepText, True)
    If Len(Trim$(partStepText)) > 0 Then
        AddTraceNote partStepText, True
        LogTraceMessage TRACE_LEVEL_INFO, "Core", "SetStep", partStepText, vbNullString, partStepText
    End If
End Sub

Private Sub LogAccessProcessPriorityForCurrentTrace(ByVal sourceProcedure As String)
    On Error Resume Next
    Dim activeTraceRunID As String
    Dim priorityDetails As String
    activeTraceRunID = GetActiveOrLastTraceRunID()
    If Len(activeTraceRunID) = 0 Then Exit Sub
    If StrComp(activeTraceRunID, mAccessProcessPriorityLoggedTraceRunID, vbBinaryCompare) = 0 Then Exit Sub
    priorityDetails = AccessProcessPriorityDetails()
    LogTraceMessage TRACE_LEVEL_INFO, "Process", sourceProcedure, "Access process priority captured at run start.", priorityDetails, "Capture Access process priority", 0, vbNullString, vbNullString, "MSACCESS.EXE", "Process", "CodeDb", "ProcessPriority"
    mAccessProcessPriorityLoggedTraceRunID = activeTraceRunID
End Sub

Private Function NewSetupRunContextDictionary() As Object
    Dim dictionaryValue As Object
    ' CANONICAL AUTHORITY: dictionaryValue is a newly owned mutable bucket in the current Setup run context; the run-context dictionary is the sole in-memory authority for that bucket and is discarded with the run.
    Set dictionaryValue = CreateObject("Scripting.Dictionary")
    dictionaryValue.CompareMode = vbTextCompare
    Set NewSetupRunContextDictionary = dictionaryValue
End Function

Private Function SetupRunContextIsActive() As Boolean
    On Error GoTo ExitPoint
    If mSetupRunContext Is Nothing Then Exit Function
    If mSetupRunContext.Exists("Active") = False Then Exit Function
    SetupRunContextIsActive = CBool(mSetupRunContext.item("Active"))
ExitPoint:
End Function

Private Function SetupRunContextBucket(ByVal bucketName As String) As Object
    If mSetupRunContext Is Nothing Then Exit Function
    If mSetupRunContext.Exists(bucketName) = False Then Err.Raise vbObjectError + 6501, "SetupRunContextBucket", "The active setup-run context does not contain the requested bucket." & vbCrLf & "Bucket: " & bucketName
    If IsObject(mSetupRunContext.item(bucketName)) = False Then Err.Raise vbObjectError + 6502, "SetupRunContextBucket", "The requested setup-run context bucket is not an object dictionary." & vbCrLf & "Bucket: " & bucketName
    Set SetupRunContextBucket = mSetupRunContext.item(bucketName)
End Function

Private Sub SetupRunContextSetBucketObject(ByVal bucketName As String, ByVal itemKey As String, ByVal itemValue As Object)
    Dim bucketValue As Object
    If SetupRunContextIsActive() = False Then Exit Sub
    If itemValue Is Nothing Then Err.Raise vbObjectError + 6505, "SetupRunContextSetBucketObject", "A setup-run context object value cannot be Nothing." & vbCrLf & "Bucket: " & bucketName & vbCrLf & "Key: " & itemKey
    Set bucketValue = SetupRunContextBucket(bucketName)
    If bucketValue.Exists(itemKey) Then
        Set bucketValue.item(itemKey) = itemValue
    Else
        bucketValue.Add itemKey, itemValue
    End If
    mSetupRunContext.item("LastUpdatedOn") = Now
End Sub

Private Function SetupEvidenceStatusIsValid(ByVal evidenceStatus As String) As Boolean
    Select Case evidenceStatus
        Case SETUP_EVIDENCE_STATUS_MATCHED, SETUP_EVIDENCE_STATUS_CHANGED, SETUP_EVIDENCE_STATUS_MISSING, SETUP_EVIDENCE_STATUS_UNSEEDED, SETUP_EVIDENCE_STATUS_PROOF_UNAVAILABLE, SETUP_EVIDENCE_STATUS_INVALID_DEFINITION, SETUP_EVIDENCE_STATUS_MUTATION_REQUIRED, SETUP_EVIDENCE_STATUS_MUTATION_COMPLETED, SETUP_EVIDENCE_STATUS_POSTCONDITION_VERIFIED, SETUP_EVIDENCE_STATUS_DEFERRED_TO_OWNER, SETUP_EVIDENCE_STATUS_FAILED
            SetupEvidenceStatusIsValid = True
    End Select
End Function

Private Function SetupEvidenceIdentityPart(ByVal identityValue As String, ByVal identityLabel As String, ByVal allowBlank As Boolean) As String
    If Len(identityValue) = 0 And allowBlank = False Then Err.Raise vbObjectError + 6506, "SetupEvidenceIdentityPart", "A structured setup-evidence identity field is blank." & vbCrLf & "Field: " & identityLabel
    SetupEvidenceIdentityPart = identityValue
End Function

Private Function SetupEvidenceKey(ByVal domainName As String, ByVal subdomainName As String, ByVal objectIdentity As String) As String
    domainName = SetupEvidenceIdentityPart(domainName, "Domain", False)
    subdomainName = SetupEvidenceIdentityPart(subdomainName, "Subdomain", False)
    objectIdentity = SetupEvidenceIdentityPart(objectIdentity, "ObjectIdentity", True)
    SetupEvidenceKey = CStr(Len(domainName)) & ":" & domainName & CStr(Len(subdomainName)) & ":" & subdomainName & CStr(Len(objectIdentity)) & ":" & objectIdentity
End Function

Private Function NewSetupEvidenceRecord(ByVal domainName As String, ByVal subdomainName As String, ByVal objectIdentity As String) As Object
    Dim recordValue As Object
    Set recordValue = NewSetupRunContextDictionary()
    recordValue.Add "SchemaVersion", SETUP_EVIDENCE_SCHEMA_VERSION
    recordValue.Add "EvidenceKey", SetupEvidenceKey(domainName, subdomainName, objectIdentity)
    recordValue.Add "Domain", domainName
    recordValue.Add "Subdomain", subdomainName
    recordValue.Add "ObjectIdentity", objectIdentity
    recordValue.Add "AuthoritativeSourceFingerprint", vbNullString
    recordValue.Add "CurrentStoredStateFingerprint", vbNullString
    recordValue.Add "CurrentPhysicalStateFingerprint", vbNullString
    recordValue.Add "Status", SETUP_EVIDENCE_STATUS_UNSEEDED
    recordValue.Add "ExactFailedProof", vbNullString
    recordValue.Add "DirtyObjects", vbNullString
    recordValue.Add "RequiredAction", vbNullString
    recordValue.Add "MutationOwner", vbNullString
    recordValue.Add "VerificationOwner", vbNullString
    recordValue.Add "Dependencies", vbNullString
    recordValue.Add "InvalidationReason", vbNullString
    recordValue.Add "SourceProcedure", vbNullString
    recordValue.Add "Detail", vbNullString
    recordValue.Add "IsDirty", False
    recordValue.Add "IsInvalidated", False
    recordValue.Add "ObservationalOnly", True
    recordValue.Add "RecordedOn", Now
    recordValue.Add "UpdatedOn", Now
    Set NewSetupEvidenceRecord = recordValue
End Function

Private Sub RecordSetupEvidenceFrameworkFailure(ByVal ProcedureName As String, ByVal failureNumber As Long, ByVal failureSource As String, ByVal failureDescription As String)
    On Error Resume Next
    If mSetupRunContext Is Nothing Then Exit Sub
    mSetupRunContext.item("EvidenceFrameworkFailureCount") = CLng(mSetupRunContext.item("EvidenceFrameworkFailureCount")) + 1
    mSetupRunContext.item("EvidenceFrameworkLastFailure") = "Procedure: " & ProcedureName & vbCrLf & "Err.Number: " & CStr(failureNumber) & vbCrLf & "Err.Source: " & failureSource & vbCrLf & "Err.Description: " & failureDescription
    mSetupRunContext.item("LastUpdatedOn") = Now
End Sub

Private Sub RegisterSetupEvidence(ByVal domainName As String, ByVal subdomainName As String, ByVal objectIdentity As String, ByVal evidenceStatus As String, Optional ByVal authoritativeSourceFingerprint As String = vbNullString, Optional ByVal storedStateFingerprint As String = vbNullString, Optional ByVal physicalStateFingerprint As String = vbNullString, Optional ByVal exactFailedProof As String = vbNullString, Optional ByVal dirtyObjects As String = vbNullString, Optional ByVal requiredAction As String = vbNullString, Optional ByVal mutationOwner As String = vbNullString, Optional ByVal verificationOwner As String = vbNullString, Optional ByVal dependencies As String = vbNullString, Optional ByVal invalidationReason As String = vbNullString, Optional ByVal sourceProcedure As String = vbNullString, Optional ByVal detailText As String = vbNullString)
    On Error GoTo FrameworkFailure
    Dim evidenceBucket As Object
    Dim recordValue As Object
    Dim evidenceKey As String
    If SetupRunContextIsActive() = False Then Exit Sub
    If SetupEvidenceStatusIsValid(evidenceStatus) = False Then Err.Raise vbObjectError + 6507, "RegisterSetupEvidence", "An unsupported structured setup-evidence status was supplied." & vbCrLf & "Status: " & evidenceStatus
    evidenceKey = SetupEvidenceKey(domainName, subdomainName, objectIdentity)
    Set evidenceBucket = SetupRunContextBucket("Evidence")
    If evidenceBucket.Exists(evidenceKey) Then
        Set recordValue = evidenceBucket.item(evidenceKey)
    Else
        Set recordValue = NewSetupEvidenceRecord(domainName, subdomainName, objectIdentity)
    End If
    recordValue.item("AuthoritativeSourceFingerprint") = authoritativeSourceFingerprint
    recordValue.item("CurrentStoredStateFingerprint") = storedStateFingerprint
    recordValue.item("CurrentPhysicalStateFingerprint") = physicalStateFingerprint
    recordValue.item("Status") = evidenceStatus
    recordValue.item("ExactFailedProof") = exactFailedProof
    recordValue.item("DirtyObjects") = dirtyObjects
    recordValue.item("RequiredAction") = requiredAction
    recordValue.item("MutationOwner") = mutationOwner
    recordValue.item("VerificationOwner") = verificationOwner
    recordValue.item("Dependencies") = dependencies
    recordValue.item("InvalidationReason") = invalidationReason
    recordValue.item("SourceProcedure") = sourceProcedure
    recordValue.item("Detail") = detailText
    recordValue.item("IsDirty") = (Len(dirtyObjects) > 0)
    recordValue.item("IsInvalidated") = (Len(invalidationReason) > 0)
    recordValue.item("UpdatedOn") = Now
    SetupRunContextSetBucketObject "Evidence", evidenceKey, recordValue
    Exit Sub
FrameworkFailure:
    RecordSetupEvidenceFrameworkFailure "RegisterSetupEvidence", Err.Number, Err.Source, Err.Description
End Sub

Private Function SetupEvidenceStatusSummary() As String
    On Error GoTo Unavailable
    Dim evidenceBucket As Object
    Dim statusCounts As Object
    Dim itemKey As Variant
    Dim recordValue As Object
    Dim statusValue As String
    Dim summaryText As String
    If mSetupRunContext Is Nothing Then Exit Function
    Set evidenceBucket = SetupRunContextBucket("Evidence")
    Set statusCounts = NewSetupRunContextDictionary()
    For Each itemKey In evidenceBucket.keys
        Set recordValue = evidenceBucket.item(itemKey)
        statusValue = CStr(recordValue.item("Status"))
        If statusCounts.Exists(statusValue) Then
            statusCounts.item(statusValue) = CLng(statusCounts.item(statusValue)) + 1
        Else
            statusCounts.Add statusValue, 1&
        End If
    Next itemKey
    For Each itemKey In statusCounts.keys
        If Len(summaryText) > 0 Then summaryText = summaryText & "; "
        summaryText = summaryText & CStr(itemKey) & "=" & CStr(statusCounts.item(itemKey))
    Next itemKey
    SetupEvidenceStatusSummary = summaryText
    Exit Function
Unavailable:
    SetupEvidenceStatusSummary = "Unavailable"
End Function

Private Sub BeginSetupRunContext(ByVal moduleGuardVerificationDetails As String)
    On Error GoTo ErrHandler
    Dim contextValue As Object
    Dim bucketValue As Object
    Dim bucketNames As Variant
    Dim bucketName As Variant
    Dim traceRunID As String
    Dim coreBuildID As String
    Dim maintainHelpDocumentsDuringSetup As Boolean
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    traceRunID = GetActiveOrLastTraceRunID()
    If Len(traceRunID) = 0 Then Err.Raise vbObjectError + 6503, "BeginSetupRunContext", "A setup-run context cannot be created without an active trace-run identity."
    coreBuildID = x__UTIL_ADC_MODULE_GUARD.ModuleGuardChatGPTVersionIDByModuleName("x__UTIL_ADC_Core")
    If Len(coreBuildID) = 0 Then Err.Raise vbObjectError + 6506, "BeginSetupRunContext", "The Core build identity could not be read from the managed rewrite header after Module Guard preflight."
    maintainHelpDocumentsDuringSetup = SettingBool(SETTING_MAINTAIN_HELP_DOCUMENTS_DURING_SETUP, DEFAULT_MAINTAIN_HELP_DOCUMENTS_DURING_SETUP)
    If Not (mSetupRunContext Is Nothing) Then
        If SetupRunContextIsActive() Then Err.Raise vbObjectError + 6504, "BeginSetupRunContext", "A setup-run context is already active." & vbCrLf & "Trace run: " & CStr(mSetupRunContext.item("TraceRunID"))
    End If
    Set contextValue = NewSetupRunContextDictionary()
    contextValue.Add "SchemaVersion", SETUP_RUN_CONTEXT_SCHEMA_VERSION
    contextValue.Add "EvidenceSchemaVersion", SETUP_EVIDENCE_SCHEMA_VERSION
    contextValue.Add "ContextID", traceRunID & "|RunFullSetup"
    contextValue.Add "TraceRunID", traceRunID
    contextValue.Add "CoreBuildIdentity", coreBuildID
    contextValue.Add "ModuleGuardResult", "Resolved"
    contextValue.Add "ModuleGuardDetails", IIf(Len(Trim$(moduleGuardVerificationDetails)) = 0, "Module Guard preflight completed successfully.", moduleGuardVerificationDetails)
    contextValue.Add "MaintainHelpDocumentsDuringSetup", maintainHelpDocumentsDuringSetup
    contextValue.Add "StartedOn", Now
    contextValue.Add "StartTimer", Timer
    contextValue.Add "EndedOn", 0#
    contextValue.Add "ElapsedSeconds", 0#
    contextValue.Add "LastUpdatedOn", Now
    contextValue.Add "CurrentPart", 0&
    contextValue.Add "CurrentStep", "Setup-run context initialized."
    contextValue.Add "CheckpointEligible", False
    contextValue.Add "CompletionStatus", "Active"
    contextValue.Add "FailureNumber", 0&
    contextValue.Add "FailureSource", vbNullString
    contextValue.Add "FailureDescription", vbNullString
    contextValue.Add "EvidenceSequence", 0&
    contextValue.Add "MutationSequence", 0&
    contextValue.Add "PostconditionSequence", 0&
    contextValue.Add "InvalidationSequence", 0&
    contextValue.Add "EvidenceFrameworkFailureCount", 0&
    contextValue.Add "EvidenceFrameworkLastFailure", vbNullString
    contextValue.Add "Active", True
    bucketNames = Array("LoadedCatalogRows", "LoadedCategoryDefaults", "LoadedSettings", "PhysicalObjectInventories", "TableSchemaFingerprints", "FieldDescriptionProofs", "QueryDefinitionProofs", "ManagedDefinitionProofs", "PlacementResolutionSnapshot", "MacroVerificationBuildResult", "GeneratedFormPreflightPlan", "ObjectFingerprints", "SubdomainFingerprints", "DirtyObjects", "DirtySubdomains", "Mutations", "Postconditions", "Evidence", "Invalidations")
    For Each bucketName In bucketNames
        Set bucketValue = NewSetupRunContextDictionary()
        contextValue.Add CStr(bucketName), bucketValue
    Next bucketName
    Set mSetupRunContext = contextValue
    RegisterSetupEvidence "ModuleGuard", "Preflight", DEFAULT_UTILITY_SHORT_NAME & " Module Guard", SETUP_EVIDENCE_STATUS_POSTCONDITION_VERIFIED, verificationOwner:="x__UTIL_ADC_MODULE_GUARD", dependencies:="Active trace-run identity; managed-header coordinated-build definition; installed " & DEFAULT_UTILITY_SHORT_NAME & "-owned module inventory", sourceProcedure:="BeginSetupRunContext", detailText:=CStr(contextValue.item("ModuleGuardDetails"))
    LogTraceMessage TRACE_LEVEL_INFO, "SetupContext", "BeginSetupRunContext", "Setup-run context initialized.", "Context ID: " & CStr(contextValue.item("ContextID")) & vbCrLf & "Schema: " & SETUP_RUN_CONTEXT_SCHEMA_VERSION & vbCrLf & "Core build: " & coreBuildID & vbCrLf & "Module Guard: Resolved", "Initialize setup-run context", 0, vbNullString, vbNullString, vbNullString, "SetupContext", "CodeDb", "Setup"
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Set mSetupRunContext = Nothing
    Err.Raise savedErrNumber, "BeginSetupRunContext", "Could not initialize the setup-run context before material setup work." & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Original description: " & savedErrDescription
End Sub

' Purpose: Resolves core setup help document synchronization scope from the supplied/current ADC state, using `SetupRunContextIsActive` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: HelpDocumentSynchronizationScope: Returns the core setup help document synchronization scope result using the type and sentinel semantics already implemented by the routine.
Public Function CoreSetupHelpDocumentSynchronizationScope() As HelpDocumentSynchronizationScope
    'Returns the operation-scoped Setup ownership decision captured once when the Setup run context was created.
    Dim maintainHelpDocumentsDuringSetup As Boolean
    If SetupRunContextIsActive() Then
        If mSetupRunContext.Exists("MaintainHelpDocumentsDuringSetup") Then
            maintainHelpDocumentsDuringSetup = CBool(mSetupRunContext.item("MaintainHelpDocumentsDuringSetup"))
        Else
            maintainHelpDocumentsDuringSetup = SettingBool(SETTING_MAINTAIN_HELP_DOCUMENTS_DURING_SETUP, DEFAULT_MAINTAIN_HELP_DOCUMENTS_DURING_SETUP)
        End If
    Else
        maintainHelpDocumentsDuringSetup = SettingBool(SETTING_MAINTAIN_HELP_DOCUMENTS_DURING_SETUP, DEFAULT_MAINTAIN_HELP_DOCUMENTS_DURING_SETUP)
    End If
    If maintainHelpDocumentsDuringSetup Then
        CoreSetupHelpDocumentSynchronizationScope = HelpDocumentScopeSetupSystemAndAllActive
    Else
        CoreSetupHelpDocumentSynchronizationScope = HelpDocumentScopeSetupSystemOnly
    End If
End Function

Private Sub UpdateSetupRunContextStep(ByVal partNumber As Long, ByVal stepText As String)
    If SetupRunContextIsActive() = False Then Exit Sub
    mSetupRunContext.item("CurrentPart") = partNumber
    mSetupRunContext.item("CurrentStep") = stepText
    mSetupRunContext.item("LastUpdatedOn") = Now
    RegisterSetupEvidence "Setup", "PartTransition", "Part" & Format$(partNumber, "00"), SETUP_EVIDENCE_STATUS_POSTCONDITION_VERIFIED, verificationOwner:="SetStep", sourceProcedure:="UpdateSetupRunContextStep", detailText:=stepText
End Sub

Private Function SetupRunContextBucketCount(ByVal bucketName As String) As Long
    Dim bucketValue As Object
    On Error GoTo ExitPoint
    Set bucketValue = SetupRunContextBucket(bucketName)
    If bucketValue Is Nothing Then Exit Function
    SetupRunContextBucketCount = CLng(bucketValue.count)
ExitPoint:
End Function

Private Function SetupRunContextSummary() As String
    Dim evidenceStatusText As String
    If mSetupRunContext Is Nothing Then Exit Function
    evidenceStatusText = SetupEvidenceStatusSummary()
    SetupRunContextSummary = "Context ID: " & CStr(mSetupRunContext.item("ContextID"))
    SetupRunContextSummary = SetupRunContextSummary & vbCrLf & "Trace run: " & CStr(mSetupRunContext.item("TraceRunID"))
    SetupRunContextSummary = SetupRunContextSummary & vbCrLf & "Core build: " & CStr(mSetupRunContext.item("CoreBuildIdentity"))
    SetupRunContextSummary = SetupRunContextSummary & vbCrLf & "Evidence schema: " & CStr(mSetupRunContext.item("EvidenceSchemaVersion"))
    SetupRunContextSummary = SetupRunContextSummary & vbCrLf & "Status: " & CStr(mSetupRunContext.item("CompletionStatus"))
    SetupRunContextSummary = SetupRunContextSummary & vbCrLf & "Elapsed: " & FormattedElapsedSeconds(CDbl(mSetupRunContext.item("ElapsedSeconds")))
    SetupRunContextSummary = SetupRunContextSummary & vbCrLf & "Current part: " & CStr(mSetupRunContext.item("CurrentPart"))
    SetupRunContextSummary = SetupRunContextSummary & vbCrLf & "Current step: " & CStr(mSetupRunContext.item("CurrentStep"))
    SetupRunContextSummary = SetupRunContextSummary & vbCrLf & "Checkpoint eligible: " & CStr(CBool(mSetupRunContext.item("CheckpointEligible")))
    SetupRunContextSummary = SetupRunContextSummary & vbCrLf & "Evidence records: " & CStr(SetupRunContextBucketCount("Evidence"))
    SetupRunContextSummary = SetupRunContextSummary & vbCrLf & "Evidence statuses: " & IIf(Len(evidenceStatusText) = 0, "(none)", evidenceStatusText)
    SetupRunContextSummary = SetupRunContextSummary & vbCrLf & "Mutations recorded: " & CStr(SetupRunContextBucketCount("Mutations"))
    SetupRunContextSummary = SetupRunContextSummary & vbCrLf & "Postconditions recorded: " & CStr(SetupRunContextBucketCount("Postconditions"))
    SetupRunContextSummary = SetupRunContextSummary & vbCrLf & "Dirty objects: " & CStr(SetupRunContextBucketCount("DirtyObjects"))
    SetupRunContextSummary = SetupRunContextSummary & vbCrLf & "Dirty subdomains: " & CStr(SetupRunContextBucketCount("DirtySubdomains"))
    SetupRunContextSummary = SetupRunContextSummary & vbCrLf & "Invalidations: " & CStr(SetupRunContextBucketCount("Invalidations"))
    SetupRunContextSummary = SetupRunContextSummary & vbCrLf & "Evidence-framework recording failures: " & CStr(mSetupRunContext.item("EvidenceFrameworkFailureCount"))
    If CLng(mSetupRunContext.item("EvidenceFrameworkFailureCount")) > 0 Then SetupRunContextSummary = SetupRunContextSummary & vbCrLf & "Last evidence-framework failure: " & CStr(mSetupRunContext.item("EvidenceFrameworkLastFailure"))
End Function

Private Sub FinalizeSetupRunContext(ByVal completionStatus As String, ByVal checkpointEligible As Boolean, Optional ByVal failureNumber As Long = 0, Optional ByVal failureSource As String = vbNullString, Optional ByVal failureDescription As String = vbNullString)
    On Error GoTo ExitPoint
    Dim severity As String
    Dim detailText As String
    If mSetupRunContext Is Nothing Then Exit Sub
    mSetupRunContext.item("CompletionStatus") = completionStatus
    mSetupRunContext.item("CheckpointEligible") = checkpointEligible
    mSetupRunContext.item("FailureNumber") = failureNumber
    mSetupRunContext.item("FailureSource") = failureSource
    mSetupRunContext.item("FailureDescription") = failureDescription
    mSetupRunContext.item("EndedOn") = Now
    mSetupRunContext.item("ElapsedSeconds") = ElapsedSecondsFromTimer(CSng(mSetupRunContext.item("StartTimer")))
    mSetupRunContext.item("LastUpdatedOn") = Now
    mSetupRunContext.item("Active") = False
    detailText = SetupRunContextSummary()
    If failureNumber <> 0 Then detailText = detailText & vbCrLf & "Failure number: " & CStr(failureNumber) & vbCrLf & "Failure source: " & IIf(Len(Trim$(failureSource)) = 0, "(blank)", failureSource) & vbCrLf & "Failure description: " & IIf(Len(Trim$(failureDescription)) = 0, "(blank)", failureDescription)
    severity = IIf(StrComp(completionStatus, "Completed", vbTextCompare) = 0, TRACE_LEVEL_INFO, IIf(StrComp(completionStatus, "Cancelled", vbTextCompare) = 0, TRACE_LEVEL_WARNING, TRACE_LEVEL_ERROR))
    LogTraceMessage severity, "SetupContext", "FinalizeSetupRunContext", "Setup-run context finalized.", detailText, "Finalize setup-run context", failureNumber, failureSource, failureDescription, vbNullString, "SetupContext", "CodeDb", "Setup"
ExitPoint:
End Sub

Private Sub SetStep(ByRef currentStep As String, ByVal stepText As String)
    Dim checkpointText As String
    If mSetupControlActive Then
        checkpointText = IIf(Len(Trim$(currentStep)) = 0, stepText, currentStep)
        RequireSetupControlContinue checkpointText
        BeginSetupPartRollbackBoundary
    End If
    currentStep = stepText
    mSetupCurrentPartText = stepText
    AddTraceStep stepText
    CoreUpdateSetupProgress stepText
    UpdateSetupRunContextStep mSetupProgressCurrentStep, stepText
End Sub

Private Sub RequireSetupControlContinue(ByVal checkpointText As String)
    If mSetupControlActive = False Then Exit Sub
    If SetupControlCheckpoint(checkpointText) Then Exit Sub
    Err.Raise ERR_SETUP_CANCELLED_BY_USER, "RequireSetupControlContinue", _
        "Setup cancellation was requested and reached a rollback-safe checkpoint." & vbCrLf & _
        "Active part: " & IIf(Len(Trim$(checkpointText)) = 0, "(not recorded)", checkpointText)
End Sub

' Purpose: Requires safe work continue for worker routine, logical item.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   workerRoutine (Input; String): Name of the worker procedure being invoked/tracked; used in dispatch, tracing, status, and error context.
'   logicalItem (Input Optional; String): Stable logical item label associated with the worker operation for progress, tracing, or user-facing diagnostics. Default: vbNullString.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub RequireSafeWorkContinue(ByVal workerRoutine As String, Optional ByVal logicalItem As String = vbNullString)
    'Called only after one rollback-safe or otherwise logically atomic work unit has completed. No-op outside an active Setup/Repair or scan progress controller.
    Dim checkpointText As String
    If mSetupControlActive = False Then Exit Sub
    checkpointText = Trim$(mSetupCurrentPartText)
    If Len(checkpointText) = 0 Then checkpointText = DEFAULT_UTILITY_SHORT_NAME & " work in progress"
    If SetupControlWorkCheckpoint(checkpointText, workerRoutine, logicalItem) Then Exit Sub
    Err.Raise ERR_SETUP_CANCELLED_BY_USER, "RequireSafeWorkContinue", _
        "Cancellation was requested and reached a safe internal work boundary." & vbCrLf & _
        "Active part: " & checkpointText & vbCrLf & _
        "Worker: " & IIf(Len(Trim$(workerRoutine)) = 0, "(not recorded)", workerRoutine) & vbCrLf & _
        "Logical item: " & IIf(Len(Trim$(logicalItem)) = 0, "(not recorded)", logicalItem)
End Sub

' Purpose: Evaluates the cancellation error predicate used by the caller to choose the correct ADC processing path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   errorNumber (Input; Long): Numeric identifier or ordinal for the error item processed or reported by this operation.
' Returns: Boolean: True when cancellation error is true for the supplied/current context; False otherwise.
Public Function IsCancellationError(ByVal errorNumber As Long) As Boolean
    IsCancellationError = (errorNumber = ERR_SETUP_CANCELLED_BY_USER)
End Function

' ===============================================================================================================================================================================
' ROLLBACK / DURABLE RECOVERY / DEFERRED CLEANUP PROCEDURES
' ===============================================================================================================================================================================

Private Sub BeginSetupPartRollbackBoundary()
    If mSetupRollbackActive = False Then
        mSetupPartRollbackStartIndex = 0
    ElseIf mSetupRollbackResumeFirstPartPending Then
        mSetupPartRollbackStartIndex = 1
        mSetupRollbackResumeFirstPartPending = False
    ElseIf mSetupRollbackEntries Is Nothing Then
        mSetupPartRollbackStartIndex = 1
    Else
        mSetupPartRollbackStartIndex = mSetupRollbackEntries.count + 1
    End If
End Sub

Private Function StateKey() As String
    'Intended use: Returns the Access Development Console State Key result for direct, query, form, or required cross-module callers.
    On Error GoTo ErrHandler
    StateKey = LCase$(CONFIG_OWNER_DATABASE_ROLE_CODE_DB) & MODULE_TAG_DELIMITER & LCase$(SafeDatabaseName()) & MODULE_TAG_DELIMITER & LCase$(SafeTableName("settings"))
    Exit Function
ErrHandler:
    StateKey = LCase$(CONFIG_OWNER_DATABASE_ROLE_CODE_DB) & "|unknown"
End Function

' Purpose: Resolves safe database name from the supplied/current ADC state, using `SharedDao_SafeDatabaseName` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: String: Returns the resolved name produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function SafeDatabaseName() As String
    On Error GoTo ErrHandler
    SafeDatabaseName = SharedDao_SafeDatabaseName(ADCDb())
    Exit Function
ErrHandler:
    SafeDatabaseName = "(could not resolve ADCDb: " & Err.Description & ")"
End Function

' Purpose: Resolves and returns safe table name from the supplied inputs and the authoritative ADC definitions/state used by this module.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   tableRole (Input; String): Role value selecting which table ownership or database context applies to this operation.
' Returns: String: Returns the supplied name converted to the safe table-name form accepted by this ADC operation; blank/invalid input follows the routine's existing fallback/error rule.
Public Function SafeTableName(ByVal tableRole As String) As String
    'Intended use: Returns the Safe Access Development Console Table Name result for direct, query, form, or required cross-module callers.
    On Error GoTo ErrHandler
    Select Case LCase$(Trim$(tableRole))
        Case "settings"
            SafeTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID)
        Case "adc_config"
            SafeTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")
        Case "health", "healthreview"
            SafeTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReview")
        Case "candidates"
            SafeTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")
        Case Else
            SafeTableName = "(unknown table role: " & tableRole & ")"
    End Select
    Exit Function
ErrHandler:
    SafeTableName = "(could not resolve " & tableRole & " table name: " & Err.Description & ")"
End Function


Private Sub LogPopupDisplayAttempted(ByVal promptText As String, ByVal TitleText As String, ByVal sourceProcedure As String, ByVal sourceArea As String)
    On Error Resume Next
    Dim titleForTrace As String
    titleForTrace = Trim$(TitleText)
    If Len(titleForTrace) = 0 Then titleForTrace = DEFAULT_UTILITY_LONG_NAME & " popup"
    LogTraceMessage TRACE_LEVEL_INFO, "Popup", sourceProcedure, "Popup display attempted: " & titleForTrace, "Prompt length: " & CStr(Len(promptText)) & vbCrLf & promptText, vbNullString, 0, vbNullString, vbNullString, titleForTrace, "Popup", "CodeDb", sourceArea
End Sub

Private Sub LogPopupShown(ByVal promptText As String, ByVal TitleText As String, ByVal sourceProcedure As String, ByVal sourceArea As String)
    On Error Resume Next
    Dim titleForTrace As String
    titleForTrace = Trim$(TitleText)
    If Len(titleForTrace) = 0 Then titleForTrace = DEFAULT_UTILITY_LONG_NAME & " popup"
    LogTraceMessage TRACE_LEVEL_INFO, "Popup", sourceProcedure, "Popup shown: " & titleForTrace, "User-facing prompt length: " & CStr(Len(promptText)) & vbCrLf & "User-facing text:" & vbCrLf & promptText, vbNullString, 0, vbNullString, vbNullString, titleForTrace, "Popup", "CodeDb", sourceArea
End Sub

Private Sub SuppressPopupMessagesForCurrentRun(ByVal sourceProcedure As String, ByVal TitleText As String, ByVal sourceArea As String, Optional ByVal statusArea As String = "Core")
    On Error Resume Next
    Dim statusText As String
    TempVars.Remove TEMPVAR_SUPPRESS_POPUPS
    TempVars.Add TEMPVAR_SUPPRESS_POPUPS, True
    TempVars.Remove TEMPVAR_SUPPRESS_POPUPS_REASON
    TempVars.Add TEMPVAR_SUPPRESS_POPUPS_REASON, "Suppressed from " & sourceArea & "." & sourceProcedure & " at " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & ". Title: " & TitleText
    LogTraceMessage TRACE_LEVEL_WARNING, "Popup", sourceProcedure, "Further " & DEFAULT_UTILITY_LONG_NAME & " popup messages suppressed for this " & DEFAULT_UTILITY_LONG_NAME & " run.", "Title: " & TitleText & vbCrLf & "Status bar, Immediate Window, TraceLog, and physical trace log reporting remain active.", vbNullString, 0, vbNullString, vbNullString, TitleText, "Popup", "CodeDb", sourceArea
    statusText = "Further " & DEFAULT_UTILITY_SHORT_NAME & " popups suppressed for this " & DEFAULT_UTILITY_SHORT_NAME & " run; trace/status reporting continues."
    Select Case UCase$(Trim$(statusArea))
        Case "TABLES": CorePulseProgress statusText
        Case "UI": SetUIStatusBarText statusText
        Case Else: CorePulseProgress statusText
    End Select
End Sub

Private Sub LogPopupSuppressed(ByVal promptText As String, ByVal TitleText As String, ByVal sourceProcedure As String, ByVal sourceArea As String)
    On Error Resume Next
    Dim titleForTrace As String
    titleForTrace = Trim$(TitleText)
    If Len(titleForTrace) = 0 Then
        If StrComp(Trim$(sourceArea), "UI", vbTextCompare) = 0 Then titleForTrace = DEFAULT_UTILITY_LONG_NAME & " UI popup" Else titleForTrace = DEFAULT_UTILITY_LONG_NAME & " popup"
    End If
    LogTraceMessage TRACE_LEVEL_INFO, "Popup", sourceProcedure, "Popup suppressed for current " & DEFAULT_UTILITY_LONG_NAME & " run: " & titleForTrace, "Prompt length: " & CStr(Len(promptText)) & vbCrLf & promptText, vbNullString, 0, vbNullString, vbNullString, titleForTrace, "Popup", "CodeDb", sourceArea
End Sub

' Purpose: Performs core show message by coordinating `PopupMessagesSuppressedForCurrentRun` and `LogPopupSuppressed` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   promptText (Input; String): Prompt text text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   iconStyle (Input; VbMsgBoxStyle): Message-box style/icon flags controlling the visual treatment and available choices for this UI operation.
'   TitleText (Input; String): Title text text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   sourceProcedure (Input Optional; String): Procedure name recorded as the source of trace, error, status, or diagnostic information emitted by this call. Default: "ShowADCMessage".
'   sourceArea (Input Optional; String): Logical application area used to categorize the source message, trace, or UI operation. Default: "Core".
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub CoreShowMessage(ByVal promptText As String, ByVal iconStyle As VbMsgBoxStyle, ByVal TitleText As String, Optional ByVal sourceProcedure As String = "ShowADCMessage", Optional ByVal sourceArea As String = "Core")
    'Intended use: Central popup-policy authority for Core, Tables, UI, and required cross-module callers.
    'Full detail text is written to TraceLog and the trace log file.
    'The Access popup is kept to one diagnostic overview so long reports do not become multi-page UtilityMsgBox runs.
    'Non-critical popups retain Cancel-based suppression. Critical error popups always display and cannot be suppressed.
    On Error GoTo ErrHandler
    Dim popupPrompt As String
    Dim displayPrompt As String
    Dim popupStyle As VbMsgBoxStyle
    Dim resultValue As VbMsgBoxResult
    Dim terminalError As Boolean
    Dim fallbackProcedure As String
    sourceProcedure = Trim$(sourceProcedure)
    sourceArea = Trim$(sourceArea)
    If Len(sourceProcedure) = 0 Then sourceProcedure = "ShowADCMessage"
    If Len(sourceArea) = 0 Then sourceArea = "Core"
    fallbackProcedure = sourceProcedure & ".Fallback"
    terminalError = ((CLng(iconStyle) And &H70) = vbCritical)
    If terminalError = False And PopupMessagesSuppressedForCurrentRun() Then
        LogPopupSuppressed promptText, TitleText, sourceProcedure, sourceArea
        Exit Sub
    End If
    If terminalError Then
        popupPrompt = promptText
        popupStyle = iconStyle
    Else
        popupPrompt = SuppressiblePopupPrompt(promptText, False)
        popupStyle = MsgBoxStyleWithSuppressionCancel(iconStyle, False)
    End If
    LogPopupDisplayAttempted popupPrompt, TitleText, sourceProcedure, sourceArea
    displayPrompt = HumanReadablePopupText(popupPrompt, TitleText, popupStyle)
    resultValue = UtilityMsgBox(displayPrompt, popupStyle, TitleText)
    LogPopupShown displayPrompt, TitleText, sourceProcedure, sourceArea
    LogPopupResult resultValue, TitleText, sourceProcedure, sourceArea
    If terminalError = False And resultValue = vbCancel Then SuppressPopupMessagesForCurrentRun sourceProcedure, TitleText, sourceArea
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim fallbackErrNumber As Long
    Dim fallbackErrSource As String
    Dim fallbackErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If terminalError Then
        popupPrompt = promptText
        popupStyle = iconStyle
    Else
        popupPrompt = SuppressiblePopupPrompt(promptText, False)
        popupStyle = MsgBoxStyleWithSuppressionCancel(iconStyle, False)
    End If
    LogTraceMessage TRACE_LEVEL_WARNING, "Popup", sourceProcedure, "Popup display helper failed; fallback popup attempted.", "Title: " & TitleText & vbCrLf & "Prompt length: " & CStr(Len(promptText)) & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, vbNullString, savedErrNumber, savedErrSource, savedErrDescription, TitleText, "Popup", "CodeDb", sourceArea
    Err.Clear
    displayPrompt = HumanReadablePopupText(popupPrompt, TitleText, popupStyle)
    resultValue = UtilityMsgBox(displayPrompt, popupStyle, TitleText)
    fallbackErrNumber = Err.Number
    fallbackErrSource = Err.Source
    fallbackErrDescription = Err.Description
    If fallbackErrNumber = 0 Then
        LogPopupShown displayPrompt, TitleText, fallbackProcedure, sourceArea
        LogPopupResult resultValue, TitleText, fallbackProcedure, sourceArea
        If terminalError = False And resultValue = vbCancel Then SuppressPopupMessagesForCurrentRun fallbackProcedure, TitleText, sourceArea
    Else
        LogTraceMessage TRACE_LEVEL_ERROR, "Popup", sourceProcedure, "Popup display fallback failed; no popup could be displayed.", "Title: " & TitleText & vbCrLf & "Prompt length: " & CStr(Len(promptText)) & vbCrLf & "Original Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Original Err.Source: " & savedErrSource & vbCrLf & "Original Err.Description: " & savedErrDescription & vbCrLf & "Fallback Err.Number: " & CStr(fallbackErrNumber) & vbCrLf & "Fallback Err.Source: " & fallbackErrSource & vbCrLf & "Fallback Err.Description: " & fallbackErrDescription, vbNullString, fallbackErrNumber, fallbackErrSource, fallbackErrDescription, TitleText, "Popup", "CodeDb", sourceArea
        Debug.Print "ADCPopup: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " - Popup display failed for [" & TitleText & "]. Review TraceLog and " & DEFAULT_UTILITY_LONG_NAME & " Trace.log."
        CorePulseProgress "Popup display failed. Review TraceLog and " & TRACE_LOG_FILE_NAME & "."
    End If
End Sub

' Purpose: Performs core show error message by coordinating `CoreShowMessage` and `ManagedObjectPhysicalName` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   promptText (Input; String): Prompt text text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   iconStyle (Input; VbMsgBoxStyle): Message-box style/icon flags controlling the visual treatment and available choices for this UI operation.
'   TitleText (Input; String): Title text text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   sourceProcedure (Input Optional; String): Procedure name recorded as the source of trace, error, status, or diagnostic information emitted by this call. Default: "ShowADCErrorMessage".
'   sourceArea (Input Optional; String): Logical application area used to categorize the source message, trace, or UI operation. Default: "Core".
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub CoreShowErrorMessage(ByVal promptText As String, ByVal iconStyle As VbMsgBoxStyle, ByVal TitleText As String, Optional ByVal sourceProcedure As String = "ShowADCErrorMessage", Optional ByVal sourceArea As String = "Core")
    'Intended use: User-facing terminal error authority. Displays the ADC error popup first, then opens the Setup-managed Report Error form after the popup is dismissed.
    'Failures while opening Report Error are logged and consumed so Access does not replace the ADC error workflow with a native runtime error dialog.
    On Error GoTo ErrHandler
    Dim formName As String
    CoreShowMessage promptText, iconStyle, TitleText, sourceProcedure, sourceArea
    formName = ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Report_Error_Dialog")
    If EnsureReportErrorDialogAvailable() = False Then Err.Raise vbObjectError + 9148, "CoreShowErrorMessage", "The managed Report Error form is unavailable after the repair attempt." & vbCrLf & "Form: " & formName
    If CurrentProject.AllForms(formName).IsLoaded Then
        forms(formName).Visible = True
        DoCmd.SelectObject acForm, formName, False
    Else
        DoCmd.OpenForm formName, acNormal
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
    LogTraceMessage TRACE_LEVEL_ERROR, "Popup", sourceProcedure, "The ADC error popup was displayed, but the Report Error form could not be opened.", "Title: " & TitleText & vbCrLf & "Report Error form: " & formName & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, "Open Report Error after terminal error", savedErrNumber, savedErrSource, savedErrDescription, formName, "Form", "CodeDb", sourceArea
    Err.Clear
End Sub


' Purpose: Performs core show message result by coordinating `PopupMessagesSuppressedForCurrentRun` and `SuppressedPopupDefaultResult` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   promptText (Input; String): Prompt text text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   iconStyle (Input; VbMsgBoxStyle): Message-box style/icon flags controlling the visual treatment and available choices for this UI operation.
'   TitleText (Input; String): Title text text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   sourceProcedure (Input Optional; String): Procedure name recorded as the source of trace, error, status, or diagnostic information emitted by this call. Default: "ShowADCMessageResult".
'   sourceArea (Input Optional; String): Logical application area used to categorize the source message, trace, or UI operation. Default: "Core".
' Returns: VbMsgBoxResult: Returns the effective user-choice result from Core message display, including the configured noninteractive/suppressed-popup default when no dialog is shown.
Public Function CoreShowMessageResult(ByVal promptText As String, ByVal iconStyle As VbMsgBoxStyle, ByVal TitleText As String, Optional ByVal sourceProcedure As String = "ShowADCMessageResult", Optional ByVal sourceArea As String = "Core") As VbMsgBoxResult
    'Intended use: Central result-returning popup-policy authority for Core, Tables, UI, and required cross-module callers.
    'Critical error popups always display and cannot be suppressed; non-critical result prompts retain Cancel-based suppression.
    On Error GoTo ErrHandler
    Dim popupPrompt As String
    Dim displayPrompt As String
    Dim popupStyle As VbMsgBoxStyle
    Dim resultValue As VbMsgBoxResult
    Dim terminalError As Boolean
    Dim fallbackProcedure As String
    sourceProcedure = Trim$(sourceProcedure)
    sourceArea = Trim$(sourceArea)
    If Len(sourceProcedure) = 0 Then sourceProcedure = "ShowADCMessageResult"
    If Len(sourceArea) = 0 Then sourceArea = "Core"
    fallbackProcedure = sourceProcedure & ".Fallback"
    terminalError = ((CLng(iconStyle) And &H70) = vbCritical)
    If terminalError = False And PopupMessagesSuppressedForCurrentRun() Then
        resultValue = SuppressedPopupDefaultResult(iconStyle)
        LogPopupSuppressed promptText, TitleText, sourceProcedure, sourceArea
        LogPopupResult resultValue, TitleText, sourceProcedure & ".Suppressed", sourceArea
        CoreShowMessageResult = resultValue
        Exit Function
    End If
    If terminalError Then
        popupPrompt = promptText
        popupStyle = iconStyle
    Else
        popupPrompt = SuppressiblePopupPrompt(promptText, True)
        popupStyle = MsgBoxStyleWithSuppressionCancel(iconStyle, True)
    End If
    LogPopupDisplayAttempted popupPrompt, TitleText, sourceProcedure, sourceArea
    displayPrompt = HumanReadablePopupText(popupPrompt, TitleText, popupStyle)
    resultValue = UtilityMsgBox(displayPrompt, popupStyle, TitleText)
    LogPopupShown displayPrompt, TitleText, sourceProcedure, sourceArea
    LogPopupResult resultValue, TitleText, sourceProcedure, sourceArea
    If terminalError = False And resultValue = vbCancel Then SuppressPopupMessagesForCurrentRun sourceProcedure, TitleText, sourceArea
    CoreShowMessageResult = resultValue
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim fallbackErrNumber As Long
    Dim fallbackErrSource As String
    Dim fallbackErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If terminalError Then
        popupPrompt = promptText
        popupStyle = iconStyle
    Else
        popupPrompt = SuppressiblePopupPrompt(promptText, True)
        popupStyle = MsgBoxStyleWithSuppressionCancel(iconStyle, True)
    End If
    LogTraceMessage TRACE_LEVEL_WARNING, "Popup", sourceProcedure, "Popup result prompt failed; fallback prompt attempted.", "Title: " & TitleText & vbCrLf & "Prompt length: " & CStr(Len(promptText)) & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, vbNullString, savedErrNumber, savedErrSource, savedErrDescription, TitleText, "Popup", "CodeDb", sourceArea
    Err.Clear
    displayPrompt = HumanReadablePopupText(popupPrompt, TitleText, popupStyle)
    resultValue = UtilityMsgBox(displayPrompt, popupStyle, TitleText)
    fallbackErrNumber = Err.Number
    fallbackErrSource = Err.Source
    fallbackErrDescription = Err.Description
    If fallbackErrNumber = 0 Then
        LogPopupShown displayPrompt, TitleText, fallbackProcedure, sourceArea
        LogPopupResult resultValue, TitleText, fallbackProcedure, sourceArea
        If terminalError = False And resultValue = vbCancel Then SuppressPopupMessagesForCurrentRun fallbackProcedure, TitleText, sourceArea
        CoreShowMessageResult = resultValue
    Else
        LogTraceMessage TRACE_LEVEL_ERROR, "Popup", sourceProcedure, "Popup result fallback failed; returning safe no/cancel result.", "Title: " & TitleText & vbCrLf & "Prompt length: " & CStr(Len(promptText)) & vbCrLf & "Original Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Original Err.Source: " & savedErrSource & vbCrLf & "Original Err.Description: " & savedErrDescription & vbCrLf & "Fallback Err.Number: " & CStr(fallbackErrNumber) & vbCrLf & "Fallback Err.Source: " & fallbackErrSource & vbCrLf & "Fallback Err.Description: " & fallbackErrDescription, vbNullString, fallbackErrNumber, fallbackErrSource, fallbackErrDescription, TitleText, "Popup", "CodeDb", sourceArea
        CoreShowMessageResult = SuppressedPopupDefaultResult(iconStyle)
    End If
End Function

Private Sub LogPopupResult(ByVal resultValue As VbMsgBoxResult, ByVal TitleText As String, ByVal sourceProcedure As String, Optional ByVal sourceArea As String = "Core")
    On Error Resume Next
    Dim titleForTrace As String
    titleForTrace = Trim$(TitleText)
    If Len(titleForTrace) = 0 Then
        If StrComp(Trim$(sourceArea), "UI", vbTextCompare) = 0 Then titleForTrace = DEFAULT_UTILITY_LONG_NAME & " UI popup" Else titleForTrace = DEFAULT_UTILITY_LONG_NAME & " popup"
    End If
    LogTraceMessage TRACE_LEVEL_INFO, "Popup", sourceProcedure, "Popup result: " & MsgBoxResultText(resultValue), "Title: " & titleForTrace & vbCrLf & "Result value: " & CStr(resultValue), vbNullString, 0, vbNullString, vbNullString, titleForTrace, "Popup", "CodeDb", sourceArea
End Sub

' Purpose: Performs report core action success by coordinating `End` and `TableDiagnosticsLocationForMessage` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   sourceProcedure (Input; String): Procedure name recorded as the source of trace, error, status, or diagnostic information emitted by this call.
'   traceCategory (Input; String): Category value classifying the trace item so the routine can select the correct naming, logging, or processing rules.
'   successMessage (Input; String): Success message text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   popupMessage (Input Optional; String): Popup message text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine. Default: vbNullString.
'   stepName (Input Optional; String): Setup/progress step label recorded or displayed for the current unit of work. Default: vbNullString.
'   objectName (Input Optional; String): Physical Access object name identifying the database object targeted by the operation. Default: vbNullString.
'   objectType (Input Optional; String): Object-type value identifying the Access/VBA object class targeted by the operation. Default: vbNullString.
'   databaseRole (Input Optional; String): Role value selecting which database ownership or database context applies to this operation. Default: "CodeDb".
'   contextName (Input Optional; String): Context label recorded with tracing/progress/errors so the caller can distinguish the logical operation producing the message. Default: vbNullString.
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: False.
'   popupTitle (Input Optional; String): Title text used as the popup title shown, stored, compared, or resolved by this operation. Default: "ADC".
'   popupIcon (Input Optional; VbMsgBoxStyle): Message-box style/icon flags controlling the visual treatment and available choices for this UI operation. Default: vbInformation.
'   applyRunWarnings (Input Optional; Boolean): When True, enables or requests run warnings for this call; when False, the routine follows its normal non-forced/default path. Default: True.
'   reportingArea (Input Optional; String): Logical application area used to categorize the reporting message, trace, or UI operation. Default: "Core".
'   detectMessageWarnings (Input Optional; Boolean): Boolean state describing whether detect message warnings applies to this invocation; True selects the named condition/path and False selects its
'   complementary/default path. Default: False.
'   visibleStatusMessage (Input Optional; String): Visible status message text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine. Default:
'     vbNullString.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub ReportCoreActionSuccess( _
    ByVal sourceProcedure As String, _
    ByVal traceCategory As String, _
    ByVal successMessage As String, _
    Optional ByVal popupMessage As String = vbNullString, _
    Optional ByVal stepName As String = vbNullString, _
    Optional ByVal objectName As String = vbNullString, _
    Optional ByVal objectType As String = vbNullString, _
    Optional ByVal databaseRole As String = "CodeDb", _
    Optional ByVal contextName As String = vbNullString, _
    Optional ByVal showMessage As Boolean = False, _
    Optional ByVal popupTitle As String = "ADC", _
    Optional ByVal popupIcon As VbMsgBoxStyle = vbInformation, _
    Optional ByVal applyRunWarnings As Boolean = True, _
    Optional ByVal reportingArea As String = "Core", _
    Optional ByVal detectMessageWarnings As Boolean = False, _
    Optional ByVal visibleStatusMessage As String = vbNullString)
    'Sole success-reporting policy authority for Core, Tables, UI, and cross-module callers.
    Dim messageText As String
    Dim reportingFailureText As String
    Dim visibleReportingFailureText As String
    Dim visibleStatusSuccessMessage As String
    Dim visibleSuccessMessage As String
    Dim visibleStepName As String
    Dim warningDetected As Boolean
    Dim normalizedReportingArea As String
    On Error GoTo ErrHandler
    normalizedReportingArea = UCase$(Trim$(reportingArea))
    If Len(normalizedReportingArea) = 0 Then normalizedReportingArea = "CORE"
    If Len(Trim$(popupTitle)) = 0 Then
        Select Case normalizedReportingArea
            Case "TABLES": popupTitle = DEFAULT_TABLE_MAINTENANCE_MESSAGE_TITLE
            Case "UI": popupTitle = DEFAULT_UI_MESSAGE_TITLE
            Case Else: popupTitle = "ADC"
        End Select
    End If
    successMessage = Trim$(successMessage)
    If Len(successMessage) = 0 Then successMessage = sourceProcedure & " completed successfully."
    visibleStatusMessage = Trim$(visibleStatusMessage)
    If Len(Trim$(popupMessage)) = 0 Then
        Select Case normalizedReportingArea
            Case "TABLES"
                messageText = successMessage & TableDiagnosticsLocationForMessage()
            Case "UI"
                messageText = successMessage & TableDiagnosticsLocationForMessage() & UIScreenContextForMessage()
            Case Else
                messageText = successMessage & TableDiagnosticsLocationForMessage() & TraceSummaryForMessage("success")
        End Select
    Else
        messageText = popupMessage
    End If
    warningDetected = (applyRunWarnings And CoreOperationHasWarnings())
    If detectMessageWarnings Then warningDetected = warningDetected Or popupIcon = vbExclamation Or CommonTextIndicatesWarning(successMessage & vbCrLf & messageText)
    If warningDetected Then
        successMessage = CompletionSuccessText(successMessage, True)
        messageText = CompletionPopupMessage(messageText, True)
        If normalizedReportingArea = "TABLES" Or normalizedReportingArea = "UI" Then
            popupIcon = CompletionPopupIcon(popupIcon, True)
        ElseIf popupIcon = vbInformation Then
            popupIcon = vbExclamation
        End If
    End If
    messageText = CoreAppendProcessRunSummary(messageText, sourceProcedure)
    visibleStatusSuccessMessage = visibleStatusMessage
    If Len(visibleStatusSuccessMessage) = 0 Then
        visibleStatusSuccessMessage = successMessage
    ElseIf warningDetected Then
        visibleStatusSuccessMessage = CompletionSuccessText(visibleStatusSuccessMessage, True)
    End If
    Select Case normalizedReportingArea
        Case "TABLES"
            visibleSuccessMessage = successMessage
            If CoreParentProgressActive() Then visibleSuccessMessage = CoreParentNestedProgressText(visibleSuccessMessage)
        Case "UI"
            visibleSuccessMessage = UIVisibleSuccessText(successMessage)
        Case Else
            visibleSuccessMessage = ProgressPartText(successMessage, False)
    End Select
    If Len(Trim$(visibleSuccessMessage)) = 0 Then visibleSuccessMessage = successMessage
    If Len(Trim$(stepName)) = 0 Then
        visibleStepName = visibleSuccessMessage
    Else
        Select Case normalizedReportingArea
            Case "TABLES"
                visibleStepName = stepName
                If CoreParentProgressActive() Then visibleStepName = CoreParentNestedProgressText(visibleStepName)
            Case "UI"
                visibleStepName = UIVisibleSuccessText(stepName)
            Case Else
                visibleStepName = ProgressPartText(stepName, False)
        End Select
        If Len(Trim$(visibleStepName)) = 0 Then visibleStepName = stepName
    End If
    On Error Resume Next
    Select Case normalizedReportingArea
        Case "TABLES": CorePulseProgress visibleStatusSuccessMessage
        Case "UI": SetUIStatusBarText UIVisibleSuccessText(visibleStatusSuccessMessage)
        Case Else: CorePulseProgress visibleStatusSuccessMessage
    End Select
    If Err.Number <> 0 Then
        reportingFailureText = DEFAULT_UTILITY_SHORT_NAME & " success reporting fallback: status update failed in " & sourceProcedure & ". Err.Number: " & CStr(Err.Number) & "; Err.Description: " & Err.Description
        visibleReportingFailureText = "Success reporting fallback: status update failed in " & sourceProcedure & ". Err.Number: " & CStr(Err.Number) & "; Err.Description: " & Err.Description
        Debug.Print reportingFailureText
        Err.Clear
        Select Case normalizedReportingArea
            Case "TABLES": CorePulseProgress visibleReportingFailureText
            Case "UI": SetUIStatusBarText visibleReportingFailureText
            Case Else: CoreSetVisibleStatusText visibleReportingFailureText
        End Select
        Err.Clear
    End If
    LogTraceMessage TRACE_LEVEL_INFO, traceCategory, sourceProcedure, visibleSuccessMessage, messageText, visibleStepName, 0, vbNullString, vbNullString, objectName, objectType, databaseRole, contextName
    If Err.Number <> 0 Then
        reportingFailureText = DEFAULT_UTILITY_SHORT_NAME & " success reporting fallback: trace/log write failed in " & sourceProcedure & ". Err.Number: " & CStr(Err.Number) & "; Err.Description: " & Err.Description
        visibleReportingFailureText = "Success reporting fallback: trace/log write failed in " & sourceProcedure & ". Err.Number: " & CStr(Err.Number) & "; Err.Description: " & Err.Description
        Debug.Print reportingFailureText
        Err.Clear
        Select Case normalizedReportingArea
            Case "TABLES": CorePulseProgress visibleReportingFailureText
            Case "UI": SetUIStatusBarText visibleReportingFailureText
            Case Else: CoreSetVisibleStatusText visibleReportingFailureText
        End Select
        Err.Clear
    End If
    If showMessage Then
        CoreShowMessage messageText, popupIcon, popupTitle, "ShowADCMessage", reportingArea
        If Err.Number <> 0 Then
            reportingFailureText = DEFAULT_UTILITY_SHORT_NAME & " success reporting fallback: success popup failed in " & sourceProcedure & ". Err.Number: " & CStr(Err.Number) & "; Err.Description: " & Err.Description
            visibleReportingFailureText = "Success reporting fallback: success popup failed in " & sourceProcedure & ". Err.Number: " & CStr(Err.Number) & "; Err.Description: " & Err.Description
            Debug.Print reportingFailureText
            Err.Clear
            Select Case normalizedReportingArea
                Case "TABLES": CorePulseProgress visibleReportingFailureText
                Case "UI": SetUIStatusBarText visibleReportingFailureText
                Case Else: CoreSetVisibleStatusText visibleReportingFailureText
            End Select
            Err.Clear
        End If
    End If
    On Error GoTo 0
    Exit Sub
ErrHandler:
    Debug.Print DEFAULT_UTILITY_LONG_NAME & " success reporting failed before reporting sinks were reached in " & sourceProcedure & ". Err.Number: " & CStr(Err.Number) & "; Err.Description: " & Err.Description
    On Error Resume Next
    reportingFailureText = DEFAULT_UTILITY_SHORT_NAME & " success reporting failed in " & sourceProcedure & ". Review Immediate Window."
    visibleReportingFailureText = "Success reporting failed in " & sourceProcedure & ". Review Immediate Window."
    Select Case normalizedReportingArea
        Case "TABLES": CorePulseProgress visibleReportingFailureText
        Case "UI": SetUIStatusBarText visibleReportingFailureText
        Case Else: CoreSetVisibleStatusText visibleReportingFailureText
    End Select
End Sub

' Purpose: Performs core set table repair warning result by coordinating `End` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   hadWarnings (Input; Boolean): Indicates whether warnings have already been detected for the operation; the routine preserves/updates warning handling based on this state.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub CoreSetTableRepairWarningResult(ByVal hadWarnings As Boolean)
    'Keeps ownership of the shared table-repair warning TempVar inside Core while exposing a stable cross-module write contract.
    On Error Resume Next
    TempVars.Remove TEMPVAR_TABLE_REPAIR_WARNINGS
    TempVars.Add TEMPVAR_TABLE_REPAIR_WARNINGS, hadWarnings
    On Error GoTo 0
End Sub

' Purpose: Performs the public clear table repair warning result operation using the routine's existing ADC workflow and side effects.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub ClearTableRepairWarningResult()
    On Error Resume Next
    TempVars.Remove TEMPVAR_TABLE_REPAIR_WARNINGS
End Sub

Private Function TableRepairHadWarnings() As Boolean
    On Error GoTo ExitHere
    TableRepairHadWarnings = CBool(TempVars(TEMPVAR_TABLE_REPAIR_WARNINGS).value)
ExitHere:
End Function

' Purpose: Adds table repair warning to parent.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub AddTableRepairWarningToParent()
    If TableRepairHadWarnings() Then CoreAddOperationNote "Table repair reported Warnings=1 or more. Review the child table-maintenance diagnostics before treating the parent workflow as warning-free."
    ClearTableRepairWarningResult
End Sub

Private Function TryAdoptSetupRollbackRecoveryForResume(ByVal requestedOperationName As String, ByVal recoveryOperation As String, ByVal recoveryDetail As String, ByVal recoveryPayload As String) As Boolean
    Dim activeTraceRunID As String
    Dim recoveredEntries As Collection
    requestedOperationName = Trim$(requestedOperationName)
    recoveryOperation = Trim$(recoveryOperation)
    recoveryDetail = Trim$(recoveryDetail)
    If Len(requestedOperationName) = 0 Then Exit Function
    If StrComp(requestedOperationName, DEFAULT_SETUP_FULL_OPERATION_NAME, vbBinaryCompare) <> 0 Then Exit Function
    If StrComp(recoveryOperation, requestedOperationName, vbBinaryCompare) <> 0 Then Exit Function
    Set recoveredEntries = DeserializeSetupRollbackEntries(recoveryPayload)
    ValidateSetupRollbackRecoveryEntriesForResume recoveredEntries
    activeTraceRunID = GetActiveTraceRunID()
    mSetupRollbackActive = True
    Set mSetupRollbackEntries = recoveredEntries
    mSetupRollbackOperationName = recoveryOperation
    mSetupRollbackOwnerTraceRunID = activeTraceRunID
    mSetupPartRollbackStartIndex = 1
    mSetupRollbackResumeFirstPartPending = True
    PersistDurableRecoveryMarker SETUP_RECOVERY_KEY, mSetupRollbackOperationName, SETUP_ROLLBACK_ACTIVE_DETAIL, SerializeSetupRollbackEntries()
    LogTraceMessage TRACE_LEVEL_INFO, "SetupRecovery", "TryAdoptSetupRollbackRecoveryForResume", "Interrupted full Setup rollback journal adopted for the rerun.", "Operation: " & mSetupRollbackOperationName & " | Recovered rollback entries: " & CStr(recoveredEntries.count) & " | Prior marker detail: " & IIf(Len(recoveryDetail) = 0, "(blank)", SharedMessage_OneLine(recoveryDetail)) & " | New owner trace run ID: " & IIf(Len(activeTraceRunID) = 0, "(none)", activeTraceRunID) & " | The valid same-operation rollback journal was preserved; A1-J1 will be rerun and normal completion will commit it.", "Adopt interrupted Setup rollback journal", 0, vbNullString, vbNullString, vbNullString, "SetupRecovery", "CodeDb", "Setup"
    TryAdoptSetupRollbackRecoveryForResume = True
End Function

Private Sub ValidateSetupRollbackRecoveryEntriesForResume(ByVal recoveryEntries As Collection)
    Dim entryData As Variant
    Dim actionText As String
    Dim minimumUpperBound As Long
    If recoveryEntries Is Nothing Then Err.Raise ERR_DURABLE_RECOVERY_FAILED, "ValidateSetupRollbackRecoveryEntriesForResume", "The interrupted full Setup recovery marker did not deserialize to a rollback-entry collection."
    For Each entryData In recoveryEntries
        PumpSetupResponsiveness
        If IsArray(entryData) = False Then Err.Raise ERR_DURABLE_RECOVERY_FAILED, "ValidateSetupRollbackRecoveryEntriesForResume", "The interrupted full Setup recovery marker contains a non-array rollback entry."
        If UBound(entryData) < 0 Then Err.Raise ERR_DURABLE_RECOVERY_FAILED, "ValidateSetupRollbackRecoveryEntriesForResume", "The interrupted full Setup recovery marker contains an empty rollback entry."
        actionText = CStr(entryData(0))
        Select Case actionText
            Case SETUP_ROLLBACK_RESTORE_QUERY_ACTION, SETUP_ROLLBACK_RESTORE_MACRO_ACTION, SETUP_ROLLBACK_RESTORE_REPORT_ACTION, SETUP_ROLLBACK_RESTORE_MANAGED_OBJECT_ROW_ACTION, SETUP_ROLLBACK_DELETE_NAV_OBJECT_ID_ACTION
                minimumUpperBound = 3
            Case SETUP_ROLLBACK_RESTORE_NAV_OBJECT_ID_ACTION
                minimumUpperBound = 5
            Case SETUP_ROLLBACK_RESTORE_DELETED_NAV_OBJECT_ID_ACTION
                minimumUpperBound = 3
            Case SETUP_ROLLBACK_RESTORE_NAV_SHORTCUT_ACTION
                minimumUpperBound = 6
            Case SETUP_ROLLBACK_DELETE_NAV_SHORTCUT_ACTION, SETUP_ROLLBACK_DELETE_NAV_GROUP_ACTION
                minimumUpperBound = 2
            Case SETUP_ROLLBACK_DELETE_TABLE_ACTION, SETUP_ROLLBACK_RESTORE_MANAGED_DEPENDENCIES_ACTION
                minimumUpperBound = 1
            Case Else
                Err.Raise ERR_DURABLE_RECOVERY_FAILED, "ValidateSetupRollbackRecoveryEntriesForResume", "The interrupted full Setup recovery marker contains an unsupported rollback action: " & actionText
        End Select
        If UBound(entryData) < minimumUpperBound Then Err.Raise ERR_DURABLE_RECOVERY_FAILED, "ValidateSetupRollbackRecoveryEntriesForResume", "The interrupted full Setup recovery marker contains an incomplete rollback entry for action: " & actionText
    Next entryData
End Sub

Private Function SerializeSetupRollbackEntries() As String
    Dim entryData As Variant
    If mSetupRollbackEntries Is Nothing Then Exit Function
    For Each entryData In mSetupRollbackEntries
        PumpSetupResponsiveness
        If Len(SerializeSetupRollbackEntries) > 0 Then SerializeSetupRollbackEntries = SerializeSetupRollbackEntries & vbCrLf
        SerializeSetupRollbackEntries = SerializeSetupRollbackEntries & SerializeRecoveryArray("SETUP", entryData)
    Next entryData
End Function

Private Function DeserializeSetupRollbackEntries(ByVal payloadText As String) As Collection
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: result holds planned/recovery work or independently observed validation evidence for DeserializeSetupRollbackEntries; keep it distinct from the source authority.
    Dim result As New Collection
    Dim lineItem As Variant
    For Each lineItem In Split(payloadText, vbCrLf)
        PumpSetupResponsiveness
        If StrComp(Left$(CStr(lineItem), 6), "SETUP" & Chr$(30), vbBinaryCompare) = 0 Then result.Add DeserializeRecoveryArray(CStr(lineItem))
    Next lineItem
    Set DeserializeSetupRollbackEntries = result
End Function

Private Sub RequireNoUnfinishedSetupRollback(ByVal requestedOperationName As String)
    'Stops a new top-level operation without discarding rollback resources owned by an interrupted setup or repair.
    Dim trackedEntryCount As Long
    Dim durableOperation As String
    Dim durableDetail As String
    Dim durablePayload As String
    If mSetupRollbackActive = False Then
        If CoreDurableRecoveryMarkerExists(SETUP_RECOVERY_KEY, durableOperation, durableDetail, durablePayload) Then
            mSetupRollbackActive = True
            mSetupRollbackOperationName = durableOperation
            Set mSetupRollbackEntries = DeserializeSetupRollbackEntries(durablePayload)
        End If
    End If
    If mSetupRollbackActive = False Then Exit Sub
    If Not mSetupRollbackEntries Is Nothing Then trackedEntryCount = mSetupRollbackEntries.count
    Err.Raise vbObjectError + 6333, "RequireNoUnfinishedSetupRollback", "A new top-level " & DEFAULT_UTILITY_LONG_NAME & " operation cannot start because an earlier setup rollback scope is still active." & vbCrLf & "Requested operation: " & requestedOperationName & vbCrLf & "Unfinished operation: " & IIf(Len(mSetupRollbackOperationName) = 0, "(not recorded)", mSetupRollbackOperationName) & vbCrLf & "Rollback owner trace run ID: " & IIf(Len(mSetupRollbackOwnerTraceRunID) = 0, "(reloaded from durable state)", mSetupRollbackOwnerTraceRunID) & vbCrLf & "Tracked rollback entries: " & CStr(trackedEntryCount) & vbCrLf & "Run RetryADCSetupRollback before starting another operation."
End Sub

Private Sub ValidateSetupRollbackOwnership(ByVal sourceProcedure As String)
    Dim activeTraceRunID As String
    If mSetupRollbackActive = False Then Exit Sub
    activeTraceRunID = GetActiveTraceRunID()
    If Len(mSetupRollbackOwnerTraceRunID) = 0 And Len(activeTraceRunID) > 0 Then mSetupRollbackOwnerTraceRunID = activeTraceRunID
    If Len(mSetupRollbackOwnerTraceRunID) > 0 And StrComp(activeTraceRunID, mSetupRollbackOwnerTraceRunID, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 6334, sourceProcedure, "The active setup rollback scope belongs to a different or completed trace run." & vbCrLf & "Rollback operation: " & mSetupRollbackOperationName & vbCrLf & "Rollback owner trace run ID: " & mSetupRollbackOwnerTraceRunID & vbCrLf & "Current active trace run ID: " & IIf(Len(activeTraceRunID) = 0, "(none)", activeTraceRunID)
End Sub

' Purpose: Performs core commit setup rollback group boundary by coordinating `BeginSetupOperation` and `CoreAddOperationNote` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   completedGroup (Input; String): Name/key of the group that has just completed; used to advance or report grouped progress state.
'   nextGroup (Input; String): Name/key of the next group to be processed; used to advance or report grouped progress state.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub CoreCommitSetupRollbackGroupBoundary(ByVal completedGroup As String, ByVal nextGroup As String)
    On Error GoTo ErrHandler
    Dim operationName As String
    completedGroup = Trim$(completedGroup)
    nextGroup = Trim$(nextGroup)
    If Len(completedGroup) = 0 Then Err.Raise vbObjectError + 6488, "CoreCommitSetupRollbackGroupBoundary", "The completed Setup rollback group is blank."
    If Len(nextGroup) = 0 Then Err.Raise vbObjectError + 6489, "CoreCommitSetupRollbackGroupBoundary", "The next Setup rollback group is blank."
    If mSetupRollbackActive = False Then Err.Raise vbObjectError + 6490, "CoreCommitSetupRollbackGroupBoundary", "No active Setup rollback journal exists at the requested durable group boundary."
    operationName = Trim$(mSetupRollbackOperationName)
    If Len(operationName) = 0 Then Err.Raise vbObjectError + 6491, "CoreCommitSetupRollbackGroupBoundary", "The active Setup rollback journal has no operation name."
    CommitSetupOperation
    mSetupMaterialCommitted = True
    BeginSetupOperation operationName
    CoreAddOperationNote "Durable Setup rollback boundary committed. Completed group: " & completedGroup & "; next rollback group: " & nextGroup & ". Earlier verified work is no longer part of the active rollback journal."
    LogTraceMessage TRACE_LEVEL_INFO, "SetupCheckpoint", "CoreCommitSetupRollbackGroupBoundary", "Setup rollback group committed.", "Completed group: " & completedGroup & vbCrLf & "Next rollback group: " & nextGroup & vbCrLf & "The new active journal contains only later work.", "Commit Setup rollback group", 0, vbNullString, vbNullString, vbNullString, "SetupCheckpoint", "CodeDb", "Setup"
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "CoreCommitSetupRollbackGroupBoundary", Err.Description
End Sub

Private Sub RollbackSetupCreatedTable(ByVal tableName As String, ByRef deletedCount As Long)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim relationshipIndex As Long
    Dim traceTableName As String
    Dim traceWriteRecordsetReleased As Boolean
    If Len(tableName) = 0 Then Err.Raise vbObjectError + 6425, "RollbackSetupCreatedTable", "A setup-created table rollback entry contained a blank table name."
    traceTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_TraceLog")
    Set db = currentDb
    If RollbackObjectExistsStrict(db, acTable, tableName, "before deleting a setup-created table") = False Then Exit Sub
    If StrComp(tableName, traceTableName, vbTextCompare) = 0 Then
        ReleaseTraceLogWriteRecordsetForSchemaMutation
        traceWriteRecordsetReleased = True
    End If
    'A table tracked by this rollback action was created during the active Setup scope, so any relationship involving it was necessarily created after the table and must be removed before the table can be deleted.
    'The retained TraceLog writer must be released before deleting any relationship that touches TraceLog.
    For relationshipIndex = db.Relations.count - 1 To 0 Step -1
        If StrComp(db.Relations(relationshipIndex).Table, tableName, vbTextCompare) = 0 Or StrComp(db.Relations(relationshipIndex).foreignTable, tableName, vbTextCompare) = 0 Then
            If traceWriteRecordsetReleased = False Then
                If StrComp(db.Relations(relationshipIndex).Table, traceTableName, vbTextCompare) = 0 Or StrComp(db.Relations(relationshipIndex).foreignTable, traceTableName, vbTextCompare) = 0 Then
                    ReleaseTraceLogWriteRecordsetForSchemaMutation
                    traceWriteRecordsetReleased = True
                End If
            End If
            db.Relations.Delete db.Relations(relationshipIndex).Name
        End If
    Next relationshipIndex
    db.Relations.Refresh
    db.TableDefs.Delete tableName
    db.TableDefs.Refresh
    If RollbackObjectExistsStrict(db, acTable, tableName, "after deleting a setup-created table") Then Err.Raise vbObjectError + 6401, "RollbackSetupCreatedTable", "The rollback table remained present after deletion. Table: " & tableName
    deletedCount = deletedCount + 1
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "RollbackSetupCreatedTable", "Could not delete a setup-created table during rollback." & vbCrLf & "Table: " & tableName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Sub

Private Sub RollbackSetupCurrentPart(ByRef rollbackSummary As String, ByRef rollbackFailed As Boolean)
    On Error GoTo ErrHandler
    Dim firstIndex As Long
    Dim itemIndex As Long
    Dim entryData As Variant
    Dim actionText As String
    Dim restoredCount As Long
    Dim deletedCount As Long
    Dim failedCount As Long
    Dim detailText As String
    Dim entryErrNumber As Long
    Dim entryErrSource As String
    Dim entryErrDescription As String
    Dim tableName As String
    Dim cleanupWarnings As String
    Dim priorCleanupWarnings As String
    Dim currentEntries As Collection
    Dim priorEntries As Collection
    Dim rollbackPass As Long
    Dim retainedSharedIdentityCount As Long
    rollbackSummary = vbNullString
    rollbackFailed = False
    ValidateSetupRollbackOwnership "RollbackSetupCurrentPart"
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: currentEntries, priorEntries holds planned/recovery work or independently observed validation evidence for RollbackSetupCurrentPart; keep it distinct from the
    ' source authority.
    Set currentEntries = New Collection
    Set priorEntries = New Collection
    If mSetupRollbackEntries Is Nothing Then GoTo CommitCompletedParts
    firstIndex = mSetupPartRollbackStartIndex
    If firstIndex < 1 Then firstIndex = 1
    If firstIndex > mSetupRollbackEntries.count + 1 Then
        firstIndex = mSetupRollbackEntries.count + 1
    End If
    For itemIndex = 1 To mSetupRollbackEntries.count
        PumpSetupResponsiveness
        If itemIndex < firstIndex Then
            priorEntries.Add mSetupRollbackEntries(itemIndex)
        Else
            currentEntries.Add mSetupRollbackEntries(itemIndex)
        End If
    Next itemIndex
    'Dependency pass 1 removes every setup-created Navigation Pane membership from the active part before identity or group rollback.
    For rollbackPass = 1 To 2
        For itemIndex = currentEntries.count To 1 Step -1
            PumpSetupResponsiveness
        entryData = currentEntries(itemIndex)
        actionText = vbNullString
        If IsArray(entryData) Then
            If UBound(entryData) >= 0 Then actionText = CStr(entryData(0))
        End If
        If rollbackPass = 1 Then
            If StrComp(actionText, SETUP_ROLLBACK_DELETE_NAV_SHORTCUT_ACTION, vbBinaryCompare) <> 0 Then GoTo NextCurrentPartRollbackEntry
        ElseIf StrComp(actionText, SETUP_ROLLBACK_DELETE_NAV_SHORTCUT_ACTION, vbBinaryCompare) = 0 Then
            GoTo NextCurrentPartRollbackEntry
        End If
        entryErrNumber = 0
        entryErrSource = vbNullString
        entryErrDescription = vbNullString
        Err.Clear
        On Error Resume Next
        If IsArray(entryData) = False Then
            Err.Raise vbObjectError + 6402, "RollbackSetupCurrentPart", _
                "A setup rollback entry was not a structured array."
        Else
            Select Case actionText
                Case SETUP_ROLLBACK_RESTORE_QUERY_ACTION
                    RollbackSetupQuery CStr(entryData(1)), CBool(entryData(2)), CStr(entryData(3)), restoredCount, deletedCount, failedCount, detailText
                Case SETUP_ROLLBACK_RESTORE_MACRO_ACTION
                    RollbackSetupAccessObject acMacro, CStr(entryData(1)), CBool(entryData(2)), CStr(entryData(3)), restoredCount, deletedCount, failedCount, detailText
                Case SETUP_ROLLBACK_RESTORE_REPORT_ACTION
                    RollbackSetupAccessObject acReport, CStr(entryData(1)), CBool(entryData(2)), CStr(entryData(3)), restoredCount, deletedCount, failedCount, detailText
                Case SETUP_ROLLBACK_RESTORE_MANAGED_OBJECT_ROW_ACTION
                    RollbackSetupManagedObjectRow CLng(entryData(1)), CBool(entryData(2)), CStr(entryData(3)), restoredCount, deletedCount, detailText
                Case SETUP_ROLLBACK_RESTORE_MANAGED_DEPENDENCIES_ACTION
                    RestoreManagedObjectDependenciesFromSetupRollback CStr(entryData(1))
                    restoredCount = restoredCount + 1
                    detailText = SharedText_AppendNonBlankLine(detailText, "Restored the prior Setup F1-owned MAC_Deps generated-row set.")
                Case SETUP_ROLLBACK_DELETE_TABLE_ACTION
                    tableName = Trim$(CStr(entryData(1)))
                    RollbackSetupCreatedTable tableName, deletedCount
                Case SETUP_ROLLBACK_DELETE_NAV_SHORTCUT_ACTION
                    RollbackSetupNavPaneShortcut CLng(entryData(1)), CLng(entryData(2)), deletedCount, detailText
                Case SETUP_ROLLBACK_RESTORE_NAV_SHORTCUT_ACTION
                    RollbackSetupNavPaneShortcutRestore CLng(entryData(1)), CLng(entryData(2)), entryData(3), entryData(4), entryData(5), entryData(6), restoredCount, detailText
                Case SETUP_ROLLBACK_DELETE_NAV_OBJECT_ID_ACTION
                    RollbackSetupNavPaneObjectID CLng(entryData(1)), CStr(entryData(2)), CLng(entryData(3)), deletedCount, retainedSharedIdentityCount, detailText
                Case SETUP_ROLLBACK_RESTORE_NAV_OBJECT_ID_ACTION
                    RollbackSetupNavPaneObjectIDRestore CLng(entryData(1)), CStr(entryData(2)), CLng(entryData(3)), CStr(entryData(4)), CLng(entryData(5)), restoredCount, detailText
                Case SETUP_ROLLBACK_RESTORE_DELETED_NAV_OBJECT_ID_ACTION
                    RollbackSetupDeletedNavPaneObjectIDRestore CLng(entryData(1)), CStr(entryData(2)), CLng(entryData(3)), restoredCount, detailText
                Case SETUP_ROLLBACK_DELETE_NAV_GROUP_ACTION
                    RollbackSetupNavPaneGroup CLng(entryData(1)), CStr(entryData(2)), deletedCount, detailText
                Case Else
                    Err.Raise vbObjectError + 6402, "RollbackSetupCurrentPart", "Setup rollback encountered an unsupported rollback action: " & actionText
            End Select
        End If
        entryErrNumber = Err.Number
        entryErrSource = Err.Source
        entryErrDescription = Err.Description
        Err.Clear
        On Error GoTo ErrHandler
        If entryErrNumber <> 0 Then
            failedCount = failedCount + 1
            detailText = SharedText_AppendNonBlankLine(detailText, _
                "Current-part rollback entry failed. " & _
                SetupRollbackEntryForMessage(entryData) & "; Error " & _
                CStr(entryErrNumber) & " from " & entryErrSource & ": " & _
                entryErrDescription)
        End If
NextCurrentPartRollbackEntry:
        Next itemIndex
    Next rollbackPass
CommitCompletedParts:
    priorCleanupWarnings = _
        CleanupSetupRollbackBackupFilesForEntries(priorEntries)
    If Len(priorCleanupWarnings) > 0 Then
        detailText = SharedText_AppendNonBlankLine( _
            detailText, priorCleanupWarnings)
    End If
    If failedCount = 0 Then
        cleanupWarnings = _
            CleanupSetupRollbackBackupFilesForEntries(currentEntries)
        If Len(cleanupWarnings) > 0 Then
            detailText = SharedText_AppendNonBlankLine( _
                detailText, cleanupWarnings)
        End If
        DeleteDurableRecoveryMarker SETUP_RECOVERY_KEY
        mSetupRollbackActive = False
        Set mSetupRollbackEntries = Nothing
        mSetupPartRollbackStartIndex = 0
        mSetupRollbackOwnerTraceRunID = vbNullString
        mSetupRollbackOperationName = vbNullString
        mSetupRollbackResumeFirstPartPending = False
    Else
        rollbackFailed = True
        Set mSetupRollbackEntries = currentEntries
        mSetupPartRollbackStartIndex = 1
        detailText = SharedText_AppendNonBlankLine(detailText, _
            "Only the active part rollback journal was retained for " & _
            "verified recovery. Completed prior parts remain committed.")
        PersistDurableRecoveryMarker SETUP_RECOVERY_KEY, _
            mSetupRollbackOperationName, detailText, _
            SerializeSetupRollbackEntries()
    End If
    rollbackSummary = "Current setup part rollback completed." & vbCrLf & _
        "Part: " & IIf(Len(Trim$(mSetupCurrentPartText)) = 0, _
        "(not recorded)", mSetupCurrentPartText) & vbCrLf & _
        "Deleted: " & CStr(deletedCount) & vbCrLf & _
        "Restored: " & CStr(restoredCount) & vbCrLf & _
        "Retained shared Navigation Pane identities: " & CStr(retainedSharedIdentityCount) & vbCrLf & _
        "Rollback operations failed: " & CStr(failedCount)
    If currentEntries.count = 0 Then
        rollbackSummary = rollbackSummary & vbCrLf & _
            "No rollback-journal entries were created by the active part."
    End If
    If Len(Trim$(detailText)) > 0 Then
        rollbackSummary = rollbackSummary & vbCrLf & detailText
    End If
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    rollbackFailed = True
    rollbackSummary = "Current setup part rollback failed." & vbCrLf & _
        "Part: " & IIf(Len(Trim$(mSetupCurrentPartText)) = 0, _
        "(not recorded)", mSetupCurrentPartText) & vbCrLf & _
        "Error: " & CStr(savedErrNumber) & vbCrLf & _
        "Source: " & savedErrSource & vbCrLf & _
        "Description: " & savedErrDescription
    On Error Resume Next
    PersistDurableRecoveryMarker SETUP_RECOVERY_KEY, _
        mSetupRollbackOperationName, rollbackSummary, _
        SerializeSetupRollbackEntries()
    On Error GoTo 0
End Sub

Private Sub RollbackSetupOperation(ByRef rollbackSummary As String, ByRef rollbackFailed As Boolean)
    On Error GoTo ErrHandler
    Dim itemIndex As Long
    Dim entryData As Variant
    Dim actionText As String
    Dim restoredCount As Long
    Dim deletedCount As Long
    Dim failedCount As Long
    Dim detailText As String
    Dim entryErrNumber As Long
    Dim entryErrSource As String
    Dim entryErrDescription As String
    Dim tableName As String
    Dim cleanupWarnings As String
    Dim rollbackPass As Long
    Dim retainedSharedIdentityCount As Long
    Dim rollbackEntryCount As Long
    Dim rollbackWorkCount As Long
    Dim rollbackProgressText As String
    Dim rollbackTimer As Single
    Dim managedObjectTableWillBeDeleted As Boolean
    Dim supersededManagedRowRollbackCount As Long
    Dim plannedEntryData As Variant
    Dim managedObjectTableName As String
    rollbackSummary = vbNullString
    rollbackFailed = False
    ValidateSetupRollbackOwnership "RollbackSetupOperation"
    If mSetupRollbackEntries Is Nothing Then
        DeleteDurableRecoveryMarker SETUP_RECOVERY_KEY
        rollbackSummary = "Setup/update rollback had no tracked entries."
        mSetupRollbackActive = False
        mSetupPartRollbackStartIndex = 0
        mSetupRollbackOwnerTraceRunID = vbNullString
        mSetupRollbackOperationName = vbNullString
        mSetupRollbackResumeFirstPartPending = False
        Exit Sub
    End If
    rollbackEntryCount = mSetupRollbackEntries.count
    On Error Resume Next
    managedObjectTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID)
    For itemIndex = rollbackEntryCount To 1 Step -1
        plannedEntryData = mSetupRollbackEntries(itemIndex)
        If IsArray(plannedEntryData) Then
            If UBound(plannedEntryData) >= 1 Then
                If StrComp(CStr(plannedEntryData(0)), SETUP_ROLLBACK_DELETE_TABLE_ACTION, vbBinaryCompare) = 0 Then
                    If Len(managedObjectTableName) > 0 And StrComp(Trim$(CStr(plannedEntryData(1))), managedObjectTableName, vbTextCompare) = 0 Then
                        managedObjectTableWillBeDeleted = True
                        Exit For
                    End If
                End If
            End If
        End If
    Next itemIndex
    Err.Clear
    On Error GoTo ErrHandler
    rollbackTimer = Timer
    rollbackProgressText = "Setup failed - rolling back " & CStr(rollbackEntryCount) & " tracked changes."
    On Error Resume Next
    CoreUpdateSetupProgress rollbackProgressText
    LogTraceMessage TRACE_LEVEL_WARNING, "SetupRecovery", "RollbackSetupOperation", "Setup rollback starting.", "Tracked entries: " & CStr(rollbackEntryCount) & " | Dependency passes: 2 | Managed-object table scheduled for deletion: " & CStr(managedObjectTableWillBeDeleted) & " | The Setup Control remains open and reports rollback progress until cleanup completes.", "Rollback Setup", 0, vbNullString, vbNullString, vbNullString, "SetupRollback", "CodeDb", "Setup"
    On Error GoTo ErrHandler
    'Dependency pass 1 removes every setup-created Navigation Pane membership before any object identity or group rollback is attempted.
    For rollbackPass = 1 To 2
        For itemIndex = mSetupRollbackEntries.count To 1 Step -1
            PumpSetupResponsiveness
        entryData = mSetupRollbackEntries(itemIndex)
        actionText = vbNullString
        If IsArray(entryData) Then
            If UBound(entryData) >= 0 Then actionText = CStr(entryData(0))
        End If
        If rollbackPass = 1 Then
            If StrComp(actionText, SETUP_ROLLBACK_DELETE_NAV_SHORTCUT_ACTION, vbBinaryCompare) <> 0 Then GoTo NextRollbackEntry
        ElseIf StrComp(actionText, SETUP_ROLLBACK_DELETE_NAV_SHORTCUT_ACTION, vbBinaryCompare) = 0 Then
            GoTo NextRollbackEntry
        End If
        rollbackWorkCount = rollbackWorkCount + 1
        If rollbackWorkCount = 1 Or rollbackWorkCount Mod 10 = 0 Then
            rollbackProgressText = "Rolling back Setup changes: " & CStr(rollbackWorkCount) & " of " & CStr(rollbackEntryCount)
            On Error Resume Next
            CoreUpdateSetupProgress rollbackProgressText
            LogTraceMessage TRACE_LEVEL_STATUS, "SetupRecovery", "RollbackSetupOperation", rollbackProgressText, "Rollback pass: " & CStr(rollbackPass) & " of 2 | Current action: " & IIf(Len(actionText) = 0, "(invalid)", actionText) & " | Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(rollbackTimer)), "Rollback Setup", 0, vbNullString, vbNullString, vbNullString, "SetupRollback", "CodeDb", "Setup"
            On Error GoTo ErrHandler
        End If
        entryErrNumber = 0
        entryErrSource = vbNullString
        entryErrDescription = vbNullString
        Err.Clear
        On Error Resume Next
        If IsArray(entryData) = False Then
            Err.Raise vbObjectError + 6402, "RollbackSetupOperation", "A setup rollback entry was not a structured array."
        Else
            Select Case actionText
                Case SETUP_ROLLBACK_RESTORE_QUERY_ACTION
                    RollbackSetupQuery CStr(entryData(1)), CBool(entryData(2)), CStr(entryData(3)), restoredCount, deletedCount, failedCount, detailText
                Case SETUP_ROLLBACK_RESTORE_MACRO_ACTION
                    RollbackSetupAccessObject acMacro, CStr(entryData(1)), CBool(entryData(2)), CStr(entryData(3)), restoredCount, deletedCount, failedCount, detailText
                Case SETUP_ROLLBACK_RESTORE_REPORT_ACTION
                    RollbackSetupAccessObject acReport, CStr(entryData(1)), CBool(entryData(2)), CStr(entryData(3)), restoredCount, deletedCount, failedCount, detailText
                Case SETUP_ROLLBACK_RESTORE_MANAGED_OBJECT_ROW_ACTION
                    If managedObjectTableWillBeDeleted Then
                        supersededManagedRowRollbackCount = supersededManagedRowRollbackCount + 1
                    Else
                        RollbackSetupManagedObjectRow CLng(entryData(1)), CBool(entryData(2)), CStr(entryData(3)), restoredCount, deletedCount, detailText
                    End If
                Case SETUP_ROLLBACK_RESTORE_MANAGED_DEPENDENCIES_ACTION
                    RestoreManagedObjectDependenciesFromSetupRollback CStr(entryData(1))
                    restoredCount = restoredCount + 1
                    detailText = SharedText_AppendNonBlankLine(detailText, "Restored the prior Setup F1-owned MAC_Deps generated-row set.")
                Case SETUP_ROLLBACK_DELETE_TABLE_ACTION
                    tableName = Trim$(CStr(entryData(1)))
                    RollbackSetupCreatedTable tableName, deletedCount
                Case SETUP_ROLLBACK_DELETE_NAV_SHORTCUT_ACTION
                    RollbackSetupNavPaneShortcut CLng(entryData(1)), CLng(entryData(2)), deletedCount, detailText
                Case SETUP_ROLLBACK_RESTORE_NAV_SHORTCUT_ACTION
                    RollbackSetupNavPaneShortcutRestore CLng(entryData(1)), CLng(entryData(2)), entryData(3), entryData(4), entryData(5), entryData(6), restoredCount, detailText
                Case SETUP_ROLLBACK_DELETE_NAV_OBJECT_ID_ACTION
                    RollbackSetupNavPaneObjectID CLng(entryData(1)), CStr(entryData(2)), CLng(entryData(3)), deletedCount, retainedSharedIdentityCount, detailText
                Case SETUP_ROLLBACK_RESTORE_NAV_OBJECT_ID_ACTION
                    RollbackSetupNavPaneObjectIDRestore CLng(entryData(1)), CStr(entryData(2)), CLng(entryData(3)), CStr(entryData(4)), CLng(entryData(5)), restoredCount, detailText
                Case SETUP_ROLLBACK_RESTORE_DELETED_NAV_OBJECT_ID_ACTION
                    RollbackSetupDeletedNavPaneObjectIDRestore CLng(entryData(1)), CStr(entryData(2)), CLng(entryData(3)), restoredCount, detailText
                Case SETUP_ROLLBACK_DELETE_NAV_GROUP_ACTION
                    RollbackSetupNavPaneGroup CLng(entryData(1)), CStr(entryData(2)), deletedCount, detailText
                Case Else
                    Err.Raise vbObjectError + 6402, "RollbackSetupOperation", "Setup rollback encountered an unsupported rollback action: " & actionText
            End Select
        End If
        entryErrNumber = Err.Number
        entryErrSource = Err.Source
        entryErrDescription = Err.Description
        Err.Clear
        On Error GoTo ErrHandler
        If entryErrNumber <> 0 Then
            failedCount = failedCount + 1
            detailText = SharedText_AppendNonBlankLine(detailText, "Rollback entry failed. " & SetupRollbackEntryForMessage(entryData) & "; Error " & CStr(entryErrNumber) & " from " & entryErrSource & ": " & entryErrDescription)
        End If
NextRollbackEntry:
        Next itemIndex
    Next rollbackPass
    If failedCount = 0 Then
        DeleteDurableRecoveryMarker SETUP_RECOVERY_KEY
        cleanupWarnings = CleanupSetupRollbackBackupFilesForEntries(mSetupRollbackEntries)
        If Len(cleanupWarnings) > 0 Then detailText = SharedText_AppendNonBlankLine(detailText, cleanupWarnings)
    Else
        rollbackFailed = True
        detailText = SharedText_AppendNonBlankLine(detailText, "Rollback backup files were preserved for verified retry because one or more rollback steps failed.")
        PersistDurableRecoveryMarker SETUP_RECOVERY_KEY, mSetupRollbackOperationName, detailText, SerializeSetupRollbackEntries()
    End If
    If failedCount = 0 Then
        mSetupRollbackActive = False
        Set mSetupRollbackEntries = Nothing
        mSetupPartRollbackStartIndex = 0
        mSetupRollbackOwnerTraceRunID = vbNullString
        mSetupRollbackOperationName = vbNullString
        mSetupRollbackResumeFirstPartPending = False
    Else
        mSetupRollbackActive = True
    End If
    rollbackSummary = "Setup/update rollback completed." & vbCrLf & "Deleted: " & CStr(deletedCount) & vbCrLf & "Restored: " & CStr(restoredCount) & vbCrLf & "Managed-row reversals superseded by tracked MAC table deletion: " & CStr(supersededManagedRowRollbackCount) & vbCrLf & "Retained shared Navigation Pane identities: " & CStr(retainedSharedIdentityCount) & vbCrLf & "Rollback operations failed: " & CStr(failedCount)
    If Len(Trim$(detailText)) > 0 Then rollbackSummary = rollbackSummary & vbCrLf & detailText
    rollbackProgressText = IIf(rollbackFailed, "Setup rollback finished with failures.", "Setup rollback completed.")
    On Error Resume Next
    CoreUpdateSetupProgress rollbackProgressText
    LogTraceMessage IIf(rollbackFailed, TRACE_LEVEL_ERROR, TRACE_LEVEL_INFO), "SetupRecovery", "RollbackSetupOperation", rollbackProgressText, rollbackSummary & " | Tracked entries: " & CStr(rollbackEntryCount) & " | Work entries processed: " & CStr(rollbackWorkCount) & " | Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(rollbackTimer)), "Rollback Setup", 0, vbNullString, vbNullString, vbNullString, "SetupRollback", "CodeDb", "Setup"
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    rollbackFailed = True
    rollbackSummary = "Setup/update rollback failed." & vbCrLf & "Error: " & CStr(savedErrNumber) & vbCrLf & "Source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription & vbCrLf & "Rollback state and backup resources remain active for recovery."
    mSetupRollbackActive = True
    On Error Resume Next
    PersistDurableRecoveryMarker SETUP_RECOVERY_KEY, mSetupRollbackOperationName, rollbackSummary, SerializeSetupRollbackEntries()
    On Error GoTo 0
End Sub

Private Function SetupRollbackEntryForMessage(ByVal entryData As Variant) As String
    On Error GoTo InvalidEntry
    Dim actionText As String
    Dim targetText As String
    If IsArray(entryData) = False Then GoTo InvalidEntry
    actionText = CStr(entryData(0))
    Select Case actionText
        Case SETUP_ROLLBACK_RESTORE_QUERY_ACTION, SETUP_ROLLBACK_DELETE_TABLE_ACTION, SETUP_ROLLBACK_RESTORE_MACRO_ACTION, SETUP_ROLLBACK_RESTORE_REPORT_ACTION
            targetText = CStr(entryData(1))
        Case SETUP_ROLLBACK_RESTORE_MANAGED_OBJECT_ROW_ACTION
            targetText = "ManagedObjectID=" & CStr(entryData(1))
        Case SETUP_ROLLBACK_RESTORE_MANAGED_DEPENDENCIES_ACTION
            targetText = "prior Setup F1-owned MAC_Deps generated-row set"
        Case SETUP_ROLLBACK_DELETE_NAV_SHORTCUT_ACTION
            targetText = "GroupID=" & CStr(entryData(1)) & "; ObjectID=" & CStr(entryData(2))
        Case SETUP_ROLLBACK_RESTORE_NAV_SHORTCUT_ACTION
            targetText = "GroupID=" & CStr(entryData(1)) & "; ObjectID=" & CStr(entryData(2)) & "; restore exact prior membership metadata"
        Case SETUP_ROLLBACK_DELETE_NAV_OBJECT_ID_ACTION
            targetText = "ObjectID=" & CStr(entryData(1)) & "; ObjectName=" & CStr(entryData(2)) & "; ObjectType=" & CStr(entryData(3))
        Case SETUP_ROLLBACK_RESTORE_NAV_OBJECT_ID_ACTION
            targetText = "ObjectID=" & CStr(entryData(1)) & "; RepairedName=" & CStr(entryData(2)) & "; RepairedType=" & CStr(entryData(3)) & "; PriorName=" & CStr(entryData(4)) & "; PriorType=" & CStr(entryData(5))
        Case SETUP_ROLLBACK_RESTORE_DELETED_NAV_OBJECT_ID_ACTION
            targetText = "ObjectID=" & CStr(entryData(1)) & "; RestoreDeletedName=" & CStr(entryData(2)) & "; RestoreDeletedType=" & CStr(entryData(3))
        Case SETUP_ROLLBACK_DELETE_NAV_GROUP_ACTION
            targetText = "GroupID=" & CStr(entryData(1)) & "; GroupName=" & CStr(entryData(2))
        Case Else
            targetText = "(unsupported rollback action)"
    End Select
    SetupRollbackEntryForMessage = "Action: " & IIf(Len(actionText) = 0, "(blank)", actionText) & "; Target: " & targetText
    Exit Function
InvalidEntry:
    SetupRollbackEntryForMessage = "Action: (invalid structured rollback entry)"
End Function

' Purpose: Returns whether Core currently owns an active setup rollback scope for required cross-module callers.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: Boolean: True while setup rollback tracking state is active; False when it is not active.
Public Function SetupRollbackTrackingActive() As Boolean
    'Returns whether Core currently owns an active setup rollback scope for required cross-module callers.
    SetupRollbackTrackingActive = mSetupRollbackActive
End Function

Private Sub RollbackSetupManagedObjectRow(ByVal managedObjectID As Long, ByVal priorRowExisted As Boolean, ByVal rollbackSql As String, ByRef restoredCount As Long, ByRef deletedCount As Long, ByRef detailText As String)
    On Error GoTo ErrHandler
    Dim resultValue As Variant
    Dim resultText As String
    If managedObjectID <= 0 Then Err.Raise vbObjectError + 6435, "RollbackSetupManagedObjectRow", "A managed-object rollback entry contained an invalid ManagedObjectID."
    If priorRowExisted And Len(Trim$(rollbackSql)) = 0 Then Err.Raise vbObjectError + 6436, "RollbackSetupManagedObjectRow", "An existing managed-object rollback entry contained no restore SQL."
    resultValue = RestoreManagedObjectRowFromSetupRollback(managedObjectID, priorRowExisted, rollbackSql)
    resultText = Trim$(CStr(Nz(resultValue, vbNullString)))
    Select Case UCase$(resultText)
        Case "RESTORED"
            restoredCount = restoredCount + 1
        Case "DELETED"
            deletedCount = deletedCount + 1
        Case Else
            Err.Raise vbObjectError + 6437, "RollbackSetupManagedObjectRow", "Catalog rollback returned an unrecognized result. ManagedObjectID=" & CStr(managedObjectID) & "; Result=" & IIf(Len(resultText) = 0, "(blank)", resultText)
    End Select
    detailText = SharedText_AppendNonBlankLine(detailText, "Managed-object row rollback " & LCase$(resultText) & ": ID=" & CStr(managedObjectID) & ".")
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "RollbackSetupManagedObjectRow", "Could not roll back a tracked MAC row." & vbCrLf & "ManagedObjectID: " & CStr(managedObjectID) & vbCrLf & "Prior row existed: " & CStr(priorRowExisted) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub RollbackSetupQuery(ByVal queryName As String, ByVal targetExisted As Boolean, ByVal previousSql As String, ByRef restoredCount As Long, ByRef deletedCount As Long, ByRef failedCount As Long, ByRef detailText As String)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim qdf As DAO.queryDef
    Dim queryWasPresent As Boolean
    Dim actualSql As String
    queryName = Trim$(queryName)
    If Len(queryName) = 0 Then Err.Raise vbObjectError + 6413, "RollbackSetupQuery", "A rollback query name was blank."
    Set db = ADCDb()
    queryWasPresent = RollbackObjectExistsStrict(db, acQuery, queryName, "before rolling back a saved query")
    If targetExisted Then
        If queryWasPresent Then
            Set qdf = db.QueryDefs(queryName)
            qdf.sql = previousSql
        Else
            Set qdf = db.CreateQueryDef(queryName, previousSql)
        End If
        db.QueryDefs.Refresh
        If RollbackObjectExistsStrict(db, acQuery, queryName, "after restoring a saved query") = False Then Err.Raise vbObjectError + 6414, "RollbackSetupQuery", "The restored query was not present after rollback. Query: " & queryName
        Set qdf = db.QueryDefs(queryName)
        RestoreManagedQueryDescriptionAfterSetupRollback qdf, queryName
        actualSql = qdf.sql
        If StrComp(FormattingInsensitiveQuerySqlKey(actualSql), FormattingInsensitiveQuerySqlKey(previousSql), vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 6415, "RollbackSetupQuery", "The restored query SQL did not match the pre-setup SQL." & vbCrLf & "Query: " & queryName & vbCrLf & "Expected SQL: " & Left$(previousSql, 3000) & vbCrLf & "Actual SQL: " & Left$(actualSql, 3000)
        VerifyRollbackQueryOpenable qdf, queryName
        restoredCount = restoredCount + 1
    Else
        If queryWasPresent Then db.QueryDefs.Delete queryName
        db.QueryDefs.Refresh
        If RollbackObjectExistsStrict(db, acQuery, queryName, "after deleting a newly created saved query") Then Err.Raise vbObjectError + 6416, "RollbackSetupQuery", "A newly created query remained present after rollback deletion. Query: " & queryName
        If queryWasPresent Then deletedCount = deletedCount + 1
    End If
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    failedCount = failedCount + 1
    detailText = SharedText_AppendNonBlankLine(detailText, "Could not roll back query [" & queryName & "]. Error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription)
End Sub

Private Sub RestoreManagedQueryDescriptionAfterSetupRollback(ByVal qdf As DAO.queryDef, ByVal queryName As String)
    'Setup rollback is a recovery boundary rather than ordinary phase work. If the restored query is managed, rollback reapplies the canonical Description so restoration does not leave managed metadata incomplete.
    Dim definitionData As Variant
    Dim canonicalDescription As String
    If TryGetManagedObjectDefinition(queryName, ADCAT_OBJECT_TYPE_QUERY, definitionData) = False Then Exit Sub
    canonicalDescription = ManagedObjectCanonicalDescription(queryName, ADCAT_OBJECT_TYPE_QUERY)
    If Len(Trim$(canonicalDescription)) = 0 Then Err.Raise vbObjectError + 10113, "RestoreManagedQueryDescriptionAfterSetupRollback", "A managed query restored by Setup rollback has no canonical Description." & vbCrLf & "Query: " & queryName
    Call EnsureDaoPropertyValue(qdf, DAO_DESCRIPTION_PROPERTY_NAME, dbText, canonicalDescription, "RestoreManagedQueryDescriptionAfterSetupRollback", queryName)
End Sub

Private Function CleanupSetupRollbackBackupFilesForEntries(ByVal rollbackEntries As Collection) As String
    Dim item As Variant
    Dim entryData As Variant
    Dim actionText As String
    Dim backupPath As String
    Dim backupCleanupCount As Long
    Dim failedPaths As String
    If rollbackEntries Is Nothing Then Exit Function
    For Each item In rollbackEntries
        PumpSetupResponsiveness
        entryData = item
        actionText = vbNullString
        backupPath = vbNullString
        If IsArray(entryData) Then
            If UBound(entryData) >= 0 Then actionText = CStr(entryData(0))
            If actionText = SETUP_ROLLBACK_RESTORE_MACRO_ACTION Or actionText = SETUP_ROLLBACK_RESTORE_FORM_ACTION Or actionText = SETUP_ROLLBACK_RESTORE_REPORT_ACTION Then
                If UBound(entryData) >= 3 Then backupPath = Trim$(CStr(entryData(3)))
            End If
        End If
        If Len(backupPath) > 0 Then
            backupCleanupCount = backupCleanupCount + 1
            PulseProgressEvery backupCleanupCount, 25, _
                "Setup rollback backup files cleaned"
            DeleteRollbackBackupFile backupPath, failedPaths
        End If
    Next item
    If Len(failedPaths) > 0 Then
        CleanupSetupRollbackBackupFilesForEntries = _
            "Setup rollback backup cleanup completed with warnings. " & _
            "The following backup files could not be removed:" & _
            vbCrLf & failedPaths
    End If
End Function

' Purpose: Deletes rollback backup file for backup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   backupPath (Input; String): Filesystem or database path identifying the backup path location used by the operation.
'   failedPaths (Output; String): Receives the failed paths returned by `SharedText_AppendNonBlankLine` after this procedure completes its validation/resolution work.
' Returns: None (Sub). Results are returned through the documented ByRef parameters and/or the routine's database, UI, or module-state side effects.
Public Sub DeleteRollbackBackupFile(ByVal backupPath As String, ByRef failedPaths As String)
    Dim fileExists As Boolean
    backupPath = Trim$(backupPath)
    If Len(backupPath) = 0 Then Exit Sub
    On Error Resume Next
    Err.Clear
    fileExists = (Len(Dir$(backupPath)) > 0)
    If Err.Number <> 0 Then
        failedPaths = SharedText_AppendNonBlankLine(failedPaths, backupPath & " - existence check failed. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    If fileExists Then Kill backupPath
    If Err.Number <> 0 Then
        failedPaths = SharedText_AppendNonBlankLine(failedPaths, backupPath & " - deletion failed. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    fileExists = (Len(Dir$(backupPath)) > 0)
    If Err.Number <> 0 Then
        failedPaths = SharedText_AppendNonBlankLine(failedPaths, backupPath & " - post-delete verification failed. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
        Err.Clear
    ElseIf fileExists Then
        failedPaths = SharedText_AppendNonBlankLine(failedPaths, backupPath & " - file remained after deletion request.")
    End If
    On Error GoTo 0
End Sub

Private Sub RollbackSetupAccessObject(ByVal objectType As AcObjectType, ByVal objectName As String, ByVal targetExisted As Boolean, ByVal backupPath As String, ByRef restoredCount As Long, ByRef deletedCount As Long, ByRef failedCount As Long, ByRef detailText As String)
    On Error GoTo ErrHandler
    Dim objectWasPresent As Boolean
    Dim backupFailureText As String
    objectName = Trim$(objectName)
    backupPath = Trim$(backupPath)
    If Len(objectName) = 0 Then Err.Raise vbObjectError + 6403, "RollbackSetupAccessObject", "A rollback Access object name was blank."
    objectWasPresent = RollbackObjectExistsStrict(currentDb, objectType, objectName, "before rolling back an Access object")
    If targetExisted Then
        If SetupRollbackBackupFileReadable(backupPath, backupFailureText) = False Then
            failedCount = failedCount + 1
            detailText = SharedText_AppendNonBlankLine(detailText, "Preserved current object [" & objectName & "] because its required rollback backup was unavailable. " & backupFailureText)
            Exit Sub
        End If
        If objectWasPresent Then DoCmd.DeleteObject objectType, objectName
        If RollbackObjectExistsStrict(currentDb, objectType, objectName, "after deleting the replacement Access object") Then Err.Raise vbObjectError + 6404, "RollbackSetupAccessObject", "The current object remained present after rollback deletion. Object: " & objectName
        Application.LoadFromText objectType, objectName, backupPath
        If RollbackObjectExistsStrict(currentDb, objectType, objectName, "after restoring the original Access object") = False Then Err.Raise vbObjectError + 6405, "RollbackSetupAccessObject", "The original object was not present after rollback restoration. Object: " & objectName
        restoredCount = restoredCount + 1
    Else
        If objectWasPresent Then DoCmd.DeleteObject objectType, objectName
        If RollbackObjectExistsStrict(currentDb, objectType, objectName, "after deleting a newly created Access object") Then Err.Raise vbObjectError + 6406, "RollbackSetupAccessObject", "A newly created object remained present after rollback deletion. Object: " & objectName
        If objectWasPresent Then deletedCount = deletedCount + 1
    End If
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    failedCount = failedCount + 1
    detailText = SharedText_AppendNonBlankLine(detailText, "Could not roll back Access object [" & objectName & "]. Error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription & IIf(Len(backupPath) = 0, vbNullString, "; Backup: " & backupPath))
End Sub

Private Function SetupRollbackBackupFileReadable(ByVal backupPath As String, ByRef failureText As String) As Boolean
    On Error GoTo ErrHandler
    Dim fileNum As Integer
    backupPath = Trim$(backupPath)
    failureText = vbNullString
    If Len(backupPath) = 0 Then failureText = "Backup path was blank.": Exit Function
    If Len(Dir$(backupPath)) = 0 Then failureText = "Backup file was not found. Path: " & backupPath: Exit Function
    fileNum = FreeFile
    Open backupPath For Binary Access Read As #fileNum
    If LOF(fileNum) <= 0 Then Close #fileNum: fileNum = 0: failureText = "Backup file was empty. Path: " & backupPath: Exit Function
    Close #fileNum
    SetupRollbackBackupFileReadable = True
    Exit Function
ErrHandler:
    On Error Resume Next
    If fileNum <> 0 Then Close #fileNum
    failureText = "Backup file could not be read. Path: " & backupPath & "; Error " & CStr(Err.Number) & ": " & Err.Description
End Function

' Purpose: Runs system command for system macro key.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   systemMacroKey (Input; String): Canonical system-macro key that identifies the system macro definition or Ribbon command being resolved.
' Returns: Boolean result of run system command; True and False correspond to the affirmative and negative branches explicitly implemented by this function.
Public Function RunSystemCommand(ByVal systemMacroKey As String) As Boolean
    On Error GoTo ErrHandler
    Dim keyText As String
    Dim executionKind As String
    Dim partIdentity As String
    Dim phaseIdentity As String
    Dim definitionData As Variant
    Dim resultValue As Variant
    keyText = Trim$(systemMacroKey)
    If Len(keyText) = 0 Then Err.Raise vbObjectError + 6550, "RunSystemCommand", "The " & DEFAULT_UTILITY_LONG_NAME & " system-command key is blank."
    'REGRESSION GUARD: System Macro 65 is a terminal/destructive command and must reach its dedicated operational path before any full SystemMacro definition lookup.
    'SystemMacroDefinitionByKey can construct the complete immutable SystemMacro definition set; those rows resolve many managed physical names and can fall back to MAC database lookups when no naming snapshot is
    'active.
    'Moving this branch below SystemMacroDefinitionByKey previously made Delete All ADC Objects appear frozen before its scope dialog or deletion progress could start. Keep this early dispatch ahead of definition
    'construction; normal commands still use the single-definition reuse path below.
    If StrComp(keyText, SYS_MACRO_KEY_DELETE_ALL_OBJECTS, vbTextCompare) = 0 Then
        If CoreDeferredMacro65DeletionRequired() Or CoreDeferredMacro65DeletionPending() Then Err.Raise vbObjectError + 9669, "RunSystemCommand", "A prior Macro 65 terminal deletion is still pending, so another Delete All ADC Objects command cannot start yet."
        'The dedicated operational owner starts the single Macro 65 trace before emitting the first performance record.
        'Keep the RunCode-active flag true until the final statement before returning to the generated Access macro.
        mRibbonRestartRequiredCurrentRun = False
        mMacro65RunCodeActive = True
        RunSystemCommand = RunOperationalSystemCommand(keyText)
        LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "RunSystemCommand", "Macro 65 operational coordinator returned to RunCode wrapper.", "Result=" & CStr(RunSystemCommand) & "; DeferredRequired=" & CStr(CoreDeferredMacro65DeletionRequired()) & "; DeferredPending=" & CStr(CoreDeferredMacro65DeletionPending()) & "; RunCodeActive=" & CStr(mMacro65RunCodeActive), "Macro 65 terminal handoff", 0, vbNullString, vbNullString, keyText, "SystemMacro", "CodeDb", "SystemMacro65"
        FlushTraceLogFile
        LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "RunSystemCommand", "Macro 65 RunCode wrapper is about to mark the invoking RunCode stack inactive.", "RunCodeActive=" & CStr(mMacro65RunCodeActive), "Macro 65 terminal handoff", 0, vbNullString, vbNullString, keyText, "SystemMacro", "CodeDb", "SystemMacro65"
        FlushTraceLogFile
        mMacro65RunCodeActive = False
        LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "RunSystemCommand", "Macro 65 RunCode stack marked inactive; control is returning to the generated Access macro.", "DeferredRequired=" & CStr(CoreDeferredMacro65DeletionRequired()) & "; DeferredPending=" & CStr(CoreDeferredMacro65DeletionPending()) & "; RunCodeActive=False", "Macro 65 terminal handoff", 0, vbNullString, vbNullString, keyText, "SystemMacro", "CodeDb", "SystemMacro65"
        FlushTraceLogFile
        Exit Function
    End If
    definitionData = SystemMacroDefinitionByKey(keyText)
    executionKind = SystemMacroExecutionKind(keyText, definitionData)
    If Len(executionKind) = 0 Then Err.Raise vbObjectError + 6551, "RunSystemCommand", "The " & DEFAULT_UTILITY_LONG_NAME & " system-command key is not defined in the authoritative SystemMacro catalog." & vbCrLf & "Command key: " & systemMacroKey
    mRibbonRestartRequiredCurrentRun = False
    Select Case UCase$(executionKind)
        Case UCase$(SYSTEM_MACRO_EXECUTION_FULL_SETUP)
            RunFullSetup True
            RunSystemCommand = True
        Case UCase$(SYSTEM_MACRO_EXECUTION_VERIFY_MODULE_GUARD)
            Call VerifyModuleUpdateGuard(False)
            RunSystemCommand = True
        Case UCase$(SYSTEM_MACRO_EXECUTION_SETUP_PART)
            partIdentity = SystemMacroSetupPartIdentity(keyText, definitionData)
            If Len(partIdentity) = 0 Then Err.Raise vbObjectError + 6552, "RunSystemCommand", "The authoritative SystemMacro definition identifies this command as a Setup Part, but its key does not resolve to an implemented Setup Part." & vbCrLf & "Command key: " & systemMacroKey
            resultValue = RunSetupPartByIdentity(partIdentity, False)
            If VarType(resultValue) = vbBoolean Then
                RunSystemCommand = CBool(resultValue)
            Else
                RunSystemCommand = True
            End If
        Case UCase$(SYSTEM_MACRO_EXECUTION_SETUP_PHASE)
            phaseIdentity = SystemMacroSetupPhaseIdentity(keyText, definitionData)
            If Len(phaseIdentity) = 0 Then Err.Raise vbObjectError + 6553, "RunSystemCommand", "The authoritative SystemMacro definition identifies this command as a Setup Phase, but its key does not resolve to a supported Setup Phase." & vbCrLf & "Command key: " & systemMacroKey
            resultValue = RunSetupPhaseByIdentity(phaseIdentity, False)
            If VarType(resultValue) = vbBoolean Then
                RunSystemCommand = CBool(resultValue)
            Else
                RunSystemCommand = True
            End If
        Case UCase$(SYSTEM_MACRO_EXECUTION_OPERATIONAL)
            RunSystemCommand = RunOperationalSystemCommand(keyText)
        Case Else
            Err.Raise vbObjectError + 6554, "RunSystemCommand", "The authoritative SystemMacro definition resolved to an unsupported execution kind." & vbCrLf & "Command key: " & systemMacroKey & vbCrLf & "Execution kind: " & executionKind
    End Select
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim messageText As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    mMacro65RunCodeActive = False
    If savedErrNumber = 0 Then savedErrNumber = vbObjectError + 6555
    If Len(Trim$(savedErrSource)) = 0 Then savedErrSource = "RunSystemCommand"
    messageText = DEFAULT_UTILITY_LONG_NAME & " system command failed." & vbCrLf & vbCrLf & _
                  "Requested key: " & systemMacroKey & vbCrLf & _
                  "Resolved key: " & IIf(Len(keyText) = 0, "(blank)", keyText) & vbCrLf & _
                  "Execution kind: " & IIf(Len(executionKind) = 0, "(unresolved)", executionKind) & vbCrLf & _
                  "Err.Number: " & CStr(savedErrNumber) & vbCrLf & _
                  "Err.Source: " & savedErrSource & vbCrLf & _
                  "Err.Description:" & vbCrLf & savedErrDescription
    On Error Resume Next
    LogTraceMessage TRACE_LEVEL_ERROR, "SystemCommand", "RunSystemCommand", "System command failed.", messageText, keyText, savedErrNumber, savedErrSource, savedErrDescription, systemMacroKey, "SystemMacro", "CodeDb", "SystemCommand"
    CoreShowErrorMessage messageText, vbCritical, DEFAULT_UTILITY_LONG_NAME & " - System Command", "RunSystemCommand", "Core"
    On Error GoTo 0
    RunSystemCommand = False
End Function

' Purpose: Resolves core invoking macro65 object name from the supplied/current ADC state, using `SystemMacroPersistedTitleByKey` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: String: Returns the resolved name produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function CoreInvokingMacro65ObjectName() As String
    'Resolves the invoking Macro 65 physical Access object name from the current persisted System Macro row first; built-in defaults are fallback only when the catalog row is unavailable.
    On Error GoTo CleanExit
    Dim definitionData As Variant
    Dim persistedMacroTitle As String
    persistedMacroTitle = Trim$(SystemMacroPersistedTitleByKey(SYS_MACRO_KEY_DELETE_ALL_OBJECTS))
    If Len(persistedMacroTitle) > 0 Then
        CoreInvokingMacro65ObjectName = persistedMacroTitle
        Exit Function
    End If
    definitionData = SystemMacroDefinitionByKey(SYS_MACRO_KEY_DELETE_ALL_OBJECTS)
    If IsArray(definitionData) = False Then Exit Function
    CoreInvokingMacro65ObjectName = SystemMacroCalculatedTitle(CStr(definitionData(1)), CStr(definitionData(2)), CStr(definitionData(0)), definitionData)
CleanExit:
End Function

' Purpose: Performs core should defer invoking macro65 deletion by coordinating `CoreInvokingMacro65ObjectName` and `StrComp` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   macroName (Input; String): Physical Access macro name identifying the macro object to inspect, run, delete, or verify.
' Returns: Boolean: True when core deferred macro65 deletion is still pending in module/setup state; False when no pending work remains.
Public Function CoreShouldDeferInvokingMacro65Deletion(ByVal macroName As String) As Boolean
    'True only for the currently executing Macro 65 object while its generated Access-macro RunCode call is active.
    Dim invokingMacroName As String
    If mMacro65RunCodeActive = False Then Exit Function
    macroName = Trim$(macroName)
    If Len(macroName) = 0 Then Exit Function
    invokingMacroName = Trim$(CoreInvokingMacro65ObjectName())
    If Len(invokingMacroName) = 0 Then Exit Function
    CoreShouldDeferInvokingMacro65Deletion = (StrComp(macroName, invokingMacroName, vbTextCompare) = 0)
End Function

' Purpose: Performs core mark invoking macro65 deletion deferred by coordinating `CoreShouldDeferInvokingMacro65Deletion` and `StrComp` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   macroName (Input; String): Physical Access macro name identifying the macro object to inspect, run, delete, or verify.
'   completionKind (Input Optional; String): Kind discriminator identifying the completion variant handled by this operation. Default: "Normal".
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub CoreMarkInvokingMacro65DeletionDeferred(ByVal macroName As String, Optional ByVal completionKind As String = "Normal")
    'Records the one Access macro whose deletion must wait until the current RunCode call returns. Deletion owners retain their ordinary deletion plans.
    macroName = Trim$(macroName)
    completionKind = Trim$(completionKind)
    If Len(macroName) = 0 Then Err.Raise vbObjectError + 9660, "CoreMarkInvokingMacro65DeletionDeferred", "The deferred invoking Macro 65 name is blank."
    If CoreShouldDeferInvokingMacro65Deletion(macroName) = False Then Err.Raise vbObjectError + 9661, "CoreMarkInvokingMacro65DeletionDeferred", "The requested deferred macro is not the currently executing Macro 65 object." & vbCrLf & "Macro: " & macroName
    If mMacro65DeferredDeleteRequired Then
        If StrComp(mMacro65DeferredDeleteName, macroName, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 9662, "CoreMarkInvokingMacro65DeletionDeferred", "A different Macro 65 terminal deletion is already registered." & vbCrLf & "Existing: " & mMacro65DeferredDeleteName & vbCrLf & "Requested: " & macroName
        If Len(completionKind) > 0 And Len(mMacro65DeferredDeleteCompletionKind) > 0 Then
            If StrComp(mMacro65DeferredDeleteCompletionKind, completionKind, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 9663, "CoreMarkInvokingMacro65DeletionDeferred", "The existing deferred Macro 65 terminal deletion has a different completion owner."
        End If
        Exit Sub
    End If
    mMacro65DeferredDeleteRequired = True
    mMacro65DeferredDeleteName = macroName
    If Len(completionKind) = 0 Then completionKind = "Normal"
    mMacro65DeferredDeleteCompletionKind = completionKind
    mMacro65DeferredDeleteInactiveTickObserved = False
    mMacro65DeferredDeleteAttemptCount = 0
    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "CoreMarkInvokingMacro65DeletionDeferred", "Invoking Macro 65 deletion deferred until its RunCode call returns.", "CompletionOwner=" & completionKind, macroName, 0, vbNullString, vbNullString, macroName, "Macro", "CodeDb", "SystemMacro65"
End Sub

' Purpose: Performs core configure deferred macro65 normal completion by coordinating `StrComp` and `End` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   showMessage (Input; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the underlying
'   operation.
'   completionMessage (Input; String): Completion message text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   completionTitle (Input Optional; String): Title text used as the completion title shown, stored, compared, or resolved by this operation. Default: "Delete All ADC Objects".
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub CoreConfigureDeferredMacro65NormalCompletion(ByVal showMessage As Boolean, ByVal completionMessage As String, Optional ByVal completionTitle As String = "Delete All ADC Objects")
    'Stores only the normal managed-deletion completion presentation. The terminal object deletion remains Core-owned execution-context work.
    If mMacro65DeferredDeleteRequired = False Then Err.Raise vbObjectError + 9664, "CoreConfigureDeferredMacro65NormalCompletion", "No invoking Macro 65 deletion is currently deferred."
    If StrComp(mMacro65DeferredDeleteCompletionKind, "Standalone", vbTextCompare) = 0 Then Exit Sub
    mMacro65DeferredDeleteCompletionKind = "Normal"
    mMacro65DeferredDeleteShowMessage = showMessage
    mMacro65DeferredDeleteCompletionMessage = completionMessage
    mMacro65DeferredDeleteCompletionTitle = completionTitle
End Sub

' Purpose: Determines whether core deferred macro65 deletion is required by the effective configuration, metadata, or current operation state.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: Boolean: True when core deferred macro65 deletion is required under the current state/policy; False when it is unnecessary.
Public Function CoreDeferredMacro65DeletionRequired() As Boolean
    ' Wrapper purpose: Read-only facade exposing whether Core's private deferred Macro 65 state requires eventual deletion.
    ' Why retained: Retained so CATALOG can inspect the lifecycle state without making the underlying Core flag public or writable.
    CoreDeferredMacro65DeletionRequired = mMacro65DeferredDeleteRequired
End Function

' Purpose: Performs core deferred macro65 deletion pending by coordinating `End` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: Boolean: True when core deferred macro65 deletion is still pending in module/setup state; False when no pending work remains.
Public Function CoreDeferredMacro65DeletionPending() As Boolean
    ' Wrapper purpose: Read-only facade translating Core's private deferred-delete timer state into the semantic question of whether Macro 65 deletion is pending.
    ' Why retained: Retained so cross-module callers do not depend on the timer ID representation or manipulate Core's private timer state.
    CoreDeferredMacro65DeletionPending = (mMacro65DeferredDeleteTimerID <> 0)
End Function

' Purpose: Resolves core deferred macro65 deletion completion kind from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: String: Returns the core deferred macro 65 deletion completion kind text value. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not
'   applicable, or not resolved.
Public Function CoreDeferredMacro65DeletionCompletionKind() As String
    ' Wrapper purpose: Read-only facade exposing Core's private completion-kind value for deferred Macro 65 deletion.
    ' Why retained: Retained so CATALOG can consume the lifecycle result without exposing the underlying Core state variable for direct access or mutation.
    CoreDeferredMacro65DeletionCompletionKind = mMacro65DeferredDeleteCompletionKind
End Function

' Purpose: Performs core is deferred invoking macro65 name by coordinating `StrComp` and `End` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   macroName (Input; String): Physical Access macro name identifying the macro object to inspect, run, delete, or verify.
' Returns: Boolean: True when the condition represented by CoreIsDeferredInvokingMacro65Name is present under the routine's documented inputs/current state; False when that condition is absent.
Public Function CoreIsDeferredInvokingMacro65Name(ByVal macroName As String) As Boolean
    If mMacro65DeferredDeleteRequired = False Then Exit Function
    CoreIsDeferredInvokingMacro65Name = (StrComp(Trim$(macroName), mMacro65DeferredDeleteName, vbTextCompare) = 0)
End Function

' Purpose: Performs the public core cancel deferred macro 65 deletion request operation using the routine's existing ADC workflow and side effects.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub CoreCancelDeferredMacro65DeletionRequest()
    'Cancels an unscheduled/failed terminal request. It is used when the owning destructive operation did not reach a state that may safely delete the invoking macro.
    CoreClearDeferredMacro65DeletionState
End Sub

Private Sub CoreClearDeferredMacro65DeletionState(Optional ByVal killActiveTimer As Boolean = True)
    On Error Resume Next
    If killActiveTimer Then
        If mMacro65DeferredDeleteTimerID <> 0 Then CoreKillTimer 0, mMacro65DeferredDeleteTimerID
    End If
    mMacro65DeferredDeleteTimerID = 0
    mMacro65DeferredDeleteRequired = False
    mMacro65DeferredDeleteName = vbNullString
    mMacro65DeferredDeleteCompletionKind = vbNullString
    mMacro65DeferredDeleteShowMessage = False
    mMacro65DeferredDeleteCompletionMessage = vbNullString
    mMacro65DeferredDeleteCompletionTitle = vbNullString
    mMacro65DeferredDeleteInactiveTickObserved = False
    mMacro65DeferredDeleteAttemptCount = 0
    mMacro65DeferredDeleteCompleteOwnedTrace = False
    mMacro65DeferredDeleteTraceSuccessMessage = vbNullString
    mMacro65DeferredDeleteTraceFailureMessage = vbNullString
    On Error GoTo 0
End Sub

' Purpose: Performs core schedule deferred macro65 deletion by coordinating `CoreSetTimer` and `End` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   completeOwnedTrace (Input; Boolean): Boolean state describing whether complete owned trace applies to this invocation; True selects the named condition/path and False selects its complementary/default path.
'   traceSuccessMessage (Input; String): Trace success message text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   traceFailureMessage (Input; String): Trace failure message text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub CoreScheduleDeferredMacro65Deletion(ByVal completeOwnedTrace As Boolean, ByVal traceSuccessMessage As String, ByVal traceFailureMessage As String)
    On Error GoTo ErrHandler
    If mMacro65DeferredDeleteRequired = False Then Err.Raise vbObjectError + 9665, "CoreScheduleDeferredMacro65Deletion", "No invoking Macro 65 deletion has been deferred."
    If mMacro65DeferredDeleteTimerID <> 0 Then Err.Raise vbObjectError + 9666, "CoreScheduleDeferredMacro65Deletion", "A deferred Macro 65 deletion timer is already active."
    If Len(mMacro65DeferredDeleteName) = 0 Then Err.Raise vbObjectError + 9667, "CoreScheduleDeferredMacro65Deletion", "The deferred invoking Macro 65 name is unavailable."
    mMacro65DeferredDeleteCompleteOwnedTrace = completeOwnedTrace
    mMacro65DeferredDeleteTraceSuccessMessage = traceSuccessMessage
    mMacro65DeferredDeleteTraceFailureMessage = traceFailureMessage
    mMacro65DeferredDeleteInactiveTickObserved = False
    mMacro65DeferredDeleteAttemptCount = 0
    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "CoreScheduleDeferredMacro65Deletion", "Deferred invoking Macro 65 timer creation starting.", "TimerIntervalMs=" & CStr(MACRO65_DEFERRED_DELETE_TIMER_INTERVAL_MS) & "; CompletionOwner=" & mMacro65DeferredDeleteCompletionKind & "; RunCodeActive=" & CStr(mMacro65RunCodeActive) & "; CompleteOwnedTrace=" & CStr(completeOwnedTrace), mMacro65DeferredDeleteName, 0, vbNullString, vbNullString, mMacro65DeferredDeleteName, "Macro", "CodeDb", "SystemMacro65"
    FlushTraceLogFile
    mMacro65DeferredDeleteTimerID = CoreSetTimer(0, 0, MACRO65_DEFERRED_DELETE_TIMER_INTERVAL_MS, AddressOf CoreDeferredMacro65DeleteTimerProc)
    If mMacro65DeferredDeleteTimerID = 0 Then Err.Raise vbObjectError + 9668, "CoreScheduleDeferredMacro65Deletion", "Windows could not create the timer required to delete the invoking Macro 65 after its RunCode call returns."
    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "CoreScheduleDeferredMacro65Deletion", "Deferred invoking Macro 65 terminal deletion scheduled.", "TimerID=" & CStr(mMacro65DeferredDeleteTimerID) & "; TimerIntervalMs=" & CStr(MACRO65_DEFERRED_DELETE_TIMER_INTERVAL_MS) & "; CompletionOwner=" & mMacro65DeferredDeleteCompletionKind & "; RunCodeActive=" & CStr(mMacro65RunCodeActive), mMacro65DeferredDeleteName, 0, vbNullString, vbNullString, mMacro65DeferredDeleteName, "Macro", "CodeDb", "SystemMacro65"
    FlushTraceLogFile
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    CoreClearDeferredMacro65DeletionState
    Err.Raise savedErrNumber, "CoreScheduleDeferredMacro65Deletion", savedErrDescription
End Sub

Private Function PhaseCTryDeleteDeferredInvokingMacro65(ByVal macroName As String, ByRef failureText As String, ByRef wasMissing As Boolean) As Boolean
    'Phase C retains physical system-macro deletion ownership. Core calls this only after the invoking Macro 65 RunCode stack has unwound.
    On Error GoTo ErrHandler
    Dim objectExists As Boolean
    Dim closeErrNumber As Long
    Dim closeErrSource As String
    Dim closeErrDescription As String
    failureText = vbNullString
    wasMissing = False
    macroName = Trim$(macroName)
    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "PhaseCTryDeleteDeferredInvokingMacro65", "Deferred Macro 65 pre-delete existence inspection starting.", "Owner=PhaseC", macroName, 0, vbNullString, vbNullString, macroName, "Macro", "CodeDb", "SystemMacro65"
    FlushTraceLogFile
    objectExists = RollbackObjectExistsStrict(currentDb, acMacro, macroName, "before deferred Phase C Macro 65 deletion")
    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "PhaseCTryDeleteDeferredInvokingMacro65", "Deferred Macro 65 pre-delete existence inspection returned.", "Owner=PhaseC; Exists=" & CStr(objectExists), macroName, 0, vbNullString, vbNullString, macroName, "Macro", "CodeDb", "SystemMacro65"
    FlushTraceLogFile
    If objectExists = False Then
        wasMissing = True
        PhaseCTryDeleteDeferredInvokingMacro65 = True
        Exit Function
    End If
    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "PhaseCTryDeleteDeferredInvokingMacro65", "Deferred Macro 65 close starting.", "Owner=PhaseC", macroName, 0, vbNullString, vbNullString, macroName, "Macro", "CodeDb", "SystemMacro65"
    FlushTraceLogFile
    On Error Resume Next
    DoCmd.Close acMacro, macroName, acSaveNo
    closeErrNumber = Err.Number
    closeErrSource = Err.Source
    closeErrDescription = Err.Description
    Err.Clear
    On Error GoTo ErrHandler
    LogTraceMessage IIf(closeErrNumber = 0, TRACE_LEVEL_INFO, TRACE_LEVEL_WARNING), "DeleteObjects", "PhaseCTryDeleteDeferredInvokingMacro65", "Deferred Macro 65 close returned.", "Owner=PhaseC; CloseErr.Number=" & CStr(closeErrNumber) & "; CloseErr.Source=" & closeErrSource & "; CloseErr.Description=" & closeErrDescription, macroName, closeErrNumber, closeErrSource, closeErrDescription, macroName, "Macro", "CodeDb", "SystemMacro65"
    FlushTraceLogFile
    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "PhaseCTryDeleteDeferredInvokingMacro65", "PHYSICAL MACRO DELETE STARTING.", "Owner=PhaseC", macroName, 0, vbNullString, vbNullString, macroName, "Macro", "CodeDb", "SystemMacro65"
    FlushTraceLogFile
    DoCmd.DeleteObject acMacro, macroName
    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "PhaseCTryDeleteDeferredInvokingMacro65", "PHYSICAL MACRO DELETE RETURNED.", "Owner=PhaseC", macroName, 0, vbNullString, vbNullString, macroName, "Macro", "CodeDb", "SystemMacro65"
    FlushTraceLogFile
    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "PhaseCTryDeleteDeferredInvokingMacro65", "Deferred Macro 65 post-delete existence inspection starting.", "Owner=PhaseC", macroName, 0, vbNullString, vbNullString, macroName, "Macro", "CodeDb", "SystemMacro65"
    FlushTraceLogFile
    objectExists = RollbackObjectExistsStrict(currentDb, acMacro, macroName, "after deferred Phase C Macro 65 deletion")
    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "PhaseCTryDeleteDeferredInvokingMacro65", "Deferred Macro 65 post-delete existence inspection returned.", "Owner=PhaseC; Exists=" & CStr(objectExists), macroName, 0, vbNullString, vbNullString, macroName, "Macro", "CodeDb", "SystemMacro65"
    FlushTraceLogFile
    If objectExists Then
        failureText = "Access still reports the invoking Macro 65 object present after the Phase C deletion attempt."
        Exit Function
    End If
    PhaseCTryDeleteDeferredInvokingMacro65 = True
    Exit Function
ErrHandler:
    failureText = "Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
End Function

Private Function CoreTryDeleteDeferredMacro65(ByVal macroName As String, ByRef failureText As String, ByRef wasMissing As Boolean) As Boolean
    'Core owns only the post-RunCode timing/handoff. The selected deletion implementation performs the final physical macro deletion.
    If StrComp(mMacro65DeferredDeleteCompletionKind, "Standalone", vbTextCompare) = 0 Then
        CoreTryDeleteDeferredMacro65 = StandaloneTryDeleteDeferredInvokingMacro65(macroName, failureText, wasMissing)
    Else
        CoreTryDeleteDeferredMacro65 = PhaseCTryDeleteDeferredInvokingMacro65(macroName, failureText, wasMissing)
    End If
End Function

Private Sub CoreCompleteDeferredMacro65Deletion(ByVal deleteSucceeded As Boolean, ByVal wasMissing As Boolean, ByVal failureText As String)
    On Error Resume Next
    Dim completionSucceeded As Boolean
    Dim terminalOperationSucceeded As Boolean
    Dim messageText As String
    Dim macroName As String
    Dim completionKind As String
    macroName = mMacro65DeferredDeleteName
    completionKind = mMacro65DeferredDeleteCompletionKind
    completionSucceeded = deleteSucceeded
    terminalOperationSucceeded = completionSucceeded
    LogTraceMessage IIf(completionSucceeded, TRACE_LEVEL_INFO, TRACE_LEVEL_ERROR), "DeleteObjects", "CoreCompleteDeferredMacro65Deletion", "Deferred Macro 65 completion processing entered.", "DeleteSucceeded=" & CStr(deleteSucceeded) & "; WasMissing=" & CStr(wasMissing) & "; Attempts=" & CStr(mMacro65DeferredDeleteAttemptCount) & "; CompletionOwner=" & completionKind & "; FailureText=" & failureText, macroName, 0, vbNullString, vbNullString, macroName, "Macro", "CodeDb", "SystemMacro65"
    FlushTraceLogFile
    If mMacro65DeferredDeleteTimerID <> 0 Then CoreKillTimer 0, mMacro65DeferredDeleteTimerID
    mMacro65DeferredDeleteTimerID = 0
    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "CoreCompleteDeferredMacro65Deletion", "Deferred Macro 65 timer disabled; terminal completion processing continues.", "CompletionOwner=" & completionKind, macroName, 0, vbNullString, vbNullString, macroName, "Macro", "CodeDb", "SystemMacro65"
    FlushTraceLogFile
    If completionSucceeded Then
        LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "CoreDeferredMacro65DeleteTimerProc", IIf(wasMissing, "Deferred invoking Macro 65 was already absent after RunCode returned.", "Deferred invoking Macro 65 deletion completed after RunCode returned."), "Attempts=" & CStr(mMacro65DeferredDeleteAttemptCount), macroName, 0, vbNullString, vbNullString, macroName, "Macro", "CodeDb", "SystemMacro65"
    Else
        LogTraceMessage TRACE_LEVEL_ERROR, "DeleteObjects", "CoreDeferredMacro65DeleteTimerProc", "Deferred invoking Macro 65 deletion failed after all post-RunCode retries.", failureText, macroName, 0, vbNullString, vbNullString, macroName, "Macro", "CodeDb", "SystemMacro65"
    End If
    FlushTraceLogFile
    If StrComp(completionKind, "Standalone", vbTextCompare) = 0 Then
        LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "CoreCompleteDeferredMacro65Deletion", "Standalone deferred-completion owner callback starting.", "CompletionSucceeded=" & CStr(completionSucceeded) & "; WasMissing=" & CStr(wasMissing), macroName, 0, vbNullString, vbNullString, macroName, "Macro", "CodeDb", "SystemMacro65"
        FlushTraceLogFile
        terminalOperationSucceeded = StandaloneDeferredMacroDeletionCompleted(completionSucceeded And (wasMissing = False), completionSucceeded And wasMissing, IIf(completionSucceeded, vbNullString, failureText))
        LogTraceMessage IIf(terminalOperationSucceeded, TRACE_LEVEL_INFO, TRACE_LEVEL_ERROR), "DeleteObjects", "CoreCompleteDeferredMacro65Deletion", "Standalone deferred-completion owner callback returned.", "TerminalOperationSucceeded=" & CStr(terminalOperationSucceeded), macroName, 0, vbNullString, vbNullString, macroName, "Macro", "CodeDb", "SystemMacro65"
        FlushTraceLogFile
    Else
        If mMacro65DeferredDeleteShowMessage Then
            LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "CoreCompleteDeferredMacro65Deletion", "Deferred Macro 65 completion message starting.", "CompletionSucceeded=" & CStr(completionSucceeded), macroName, 0, vbNullString, vbNullString, macroName, "Macro", "CodeDb", "SystemMacro65"
            FlushTraceLogFile
            If completionSucceeded Then
                CoreShowMessage mMacro65DeferredDeleteCompletionMessage, vbInformation, IIf(Len(Trim$(mMacro65DeferredDeleteCompletionTitle)) = 0, "Delete All ADC Objects", mMacro65DeferredDeleteCompletionTitle), "CoreDeferredMacro65DeleteTimerProc", "Macros"
            Else
                messageText = "Delete All ADC Objects did not fully complete because Access could not delete the invoking Macro 65 after its RunCode call returned." & vbCrLf & vbCrLf & "Macro: " & macroName & vbCrLf & failureText
                CoreShowErrorMessage messageText, vbCritical, "Delete All ADC Objects", "CoreDeferredMacro65DeleteTimerProc", "Macros"
            End If
            LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "CoreCompleteDeferredMacro65Deletion", "Deferred Macro 65 completion message returned.", "CompletionSucceeded=" & CStr(completionSucceeded), macroName, 0, vbNullString, vbNullString, macroName, "Macro", "CodeDb", "SystemMacro65"
            FlushTraceLogFile
        End If
    End If
    If mMacro65DeferredDeleteCompleteOwnedTrace Then
        If TraceRunActive() Then
            LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "CoreCompleteDeferredMacro65Deletion", "Deferred Macro 65 trace completion starting.", "TerminalOperationSucceeded=" & CStr(terminalOperationSucceeded) & "; CompletionOwner=" & completionKind, macroName, 0, vbNullString, vbNullString, macroName, "Macro", "CodeDb", "SystemMacro65"
            FlushTraceLogFile
            If terminalOperationSucceeded Then
                CompleteTraceRun mMacro65DeferredDeleteTraceSuccessMessage, False
            Else
                CompleteTraceRun mMacro65DeferredDeleteTraceFailureMessage, False
            End If
        End If
    End If
    CoreClearDeferredMacro65DeletionState False
    On Error GoTo 0
End Sub

' Purpose: Handles the Windows timer callback used to complete deferred deletion of the invoking Macro 65 object after its current invocation has unwound.
' Public role: Windows callback public entry point. The Windows timer/message API invokes this routine through its callback address, so the Public signature and parameter order must remain stable.
' Parameters:
'   hWnd (Input; LongPtr): Window handle supplied by Windows to the callback for the window associated with the timer/event message.
'   uMsg (Input; Long): Windows message identifier supplied to the callback.
'   idEvent (Input; LongPtr): Windows timer/event identifier supplied to the callback and used to distinguish the registered event.
'   dwTime (Input; Long): Windows system-time value supplied with the callback notification.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
#If VBA7 Then
Public Sub CoreDeferredMacro65DeleteTimerProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal idEvent As LongPtr, ByVal dwTime As Long)
#Else
Public Sub CoreDeferredMacro65DeleteTimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
#End If
    On Error GoTo ErrHandler
    Dim failureText As String
    Dim wasMissing As Boolean
    Dim deleteSucceeded As Boolean
    If mMacro65DeferredDeleteTimerID = 0 Then Exit Sub
    If idEvent <> mMacro65DeferredDeleteTimerID Then Exit Sub
    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "CoreDeferredMacro65DeleteTimerProc", "Deferred Macro 65 timer callback entered.", "TimerID=" & CStr(idEvent) & "; ExpectedTimerID=" & CStr(mMacro65DeferredDeleteTimerID) & "; RunCodeActive=" & CStr(mMacro65RunCodeActive) & "; InactiveTickObserved=" & CStr(mMacro65DeferredDeleteInactiveTickObserved) & "; Attempts=" & CStr(mMacro65DeferredDeleteAttemptCount), mMacro65DeferredDeleteName, 0, vbNullString, vbNullString, mMacro65DeferredDeleteName, "Macro", "CodeDb", "SystemMacro65"
    FlushTraceLogFile
    If mMacro65RunCodeActive Then
        mMacro65DeferredDeleteInactiveTickObserved = False
        LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "CoreDeferredMacro65DeleteTimerProc", "Deferred Macro 65 timer observed active RunCode; deletion remains deferred.", "RunCodeActive=True; InactiveTickObserved reset=False", mMacro65DeferredDeleteName, 0, vbNullString, vbNullString, mMacro65DeferredDeleteName, "Macro", "CodeDb", "SystemMacro65"
        FlushTraceLogFile
        Exit Sub
    End If
    'A Windows timer callback can be dispatched by DoEvents. Require a full inactive timer interval after RunSystemCommand first reports the RunCode stack inactive.
    If mMacro65DeferredDeleteInactiveTickObserved = False Then
        mMacro65DeferredDeleteInactiveTickObserved = True
        LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "CoreDeferredMacro65DeleteTimerProc", "First fully inactive deferred Macro 65 timer interval observed; one stabilization interval is required before deletion.", "RunCodeActive=False; InactiveTickObserved=True", mMacro65DeferredDeleteName, 0, vbNullString, vbNullString, mMacro65DeferredDeleteName, "Macro", "CodeDb", "SystemMacro65"
        FlushTraceLogFile
        Exit Sub
    End If
    mMacro65DeferredDeleteAttemptCount = mMacro65DeferredDeleteAttemptCount + 1
    failureText = vbNullString
    wasMissing = False
    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "CoreDeferredMacro65DeleteTimerProc", "Deferred Macro 65 deletion attempt starting.", "Attempt=" & CStr(mMacro65DeferredDeleteAttemptCount) & "; CompletionOwner=" & mMacro65DeferredDeleteCompletionKind, mMacro65DeferredDeleteName, 0, vbNullString, vbNullString, mMacro65DeferredDeleteName, "Macro", "CodeDb", "SystemMacro65"
    FlushTraceLogFile
    deleteSucceeded = CoreTryDeleteDeferredMacro65(mMacro65DeferredDeleteName, failureText, wasMissing)
    LogTraceMessage IIf(deleteSucceeded, TRACE_LEVEL_INFO, TRACE_LEVEL_WARNING), "DeleteObjects", "CoreDeferredMacro65DeleteTimerProc", "Deferred Macro 65 deletion attempt returned.", "Attempt=" & CStr(mMacro65DeferredDeleteAttemptCount) & "; DeleteSucceeded=" & CStr(deleteSucceeded) & "; WasMissing=" & CStr(wasMissing) & "; FailureText=" & failureText, mMacro65DeferredDeleteName, 0, vbNullString, vbNullString, mMacro65DeferredDeleteName, "Macro", "CodeDb", "SystemMacro65"
    FlushTraceLogFile
    If deleteSucceeded Then
        CoreCompleteDeferredMacro65Deletion True, wasMissing, vbNullString
        Exit Sub
    End If
    If mMacro65DeferredDeleteAttemptCount < MACRO65_DEFERRED_DELETE_MAX_ATTEMPTS Then
        LogTraceMessage TRACE_LEVEL_WARNING, "DeleteObjects", "CoreDeferredMacro65DeleteTimerProc", "Deferred invoking Macro 65 deletion was not yet accepted by Access; the terminal timer will retry.", "Attempt=" & CStr(mMacro65DeferredDeleteAttemptCount) & "; " & failureText, mMacro65DeferredDeleteName, 0, vbNullString, vbNullString, mMacro65DeferredDeleteName, "Macro", "CodeDb", "SystemMacro65"
        FlushTraceLogFile
        Exit Sub
    End If
    CoreCompleteDeferredMacro65Deletion False, False, failureText
    Exit Sub
ErrHandler:
    failureText = "Fatal timer error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    CoreCompleteDeferredMacro65Deletion False, False, failureText
End Sub

' Purpose: Creates an independent DAO recovery workspace and opens the current CodeDb through it so durable recovery operations can run without reusing the caller's normal database context.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   recoveryWs (Output; DAO.workspace): Receives the DAO.workspace reference resolved or created for recovery ws; Nothing is left only on the procedure's documented unavailable/failure path.
'   recoveryDb (Output; DAO.Database): Receives the DAO.Database reference resolved or created for recovery DB; Nothing is left only on the procedure's documented unavailable/failure path.
' Returns: None (Sub). Results are returned through the documented ByRef parameters and/or the routine's database, UI, or module-state side effects.
Public Sub OpenCoreRecoveryDatabase(ByRef recoveryWs As DAO.workspace, ByRef recoveryDb As DAO.Database)
    Dim workspaceName As String
    workspaceName = "ADCCoreRecovery_" & Format$(Now, "hhnnss") & "_" & Format$(CLng(Timer * 1000), "000000000")
    Set recoveryWs = DBEngine.CreateWorkspace(workspaceName, "Admin", vbNullString, dbUseJet)
    Set recoveryDb = recoveryWs.OpenDatabase(currentDb.Name)
End Sub

' Purpose: Performs ensure core recovery storage ready by coordinating `OpenCoreRecoveryDatabase` and `EnsureCoreRecoveryStorage` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub EnsureCoreRecoveryStorageReady()
    'Owns the independent recovery workspace/database open, Setup_RecoveryState readiness check, and deterministic cleanup for callers that only need durable storage prepared.
    On Error GoTo ErrHandler
    Dim recoveryWs As DAO.workspace
    Dim recoveryDb As DAO.Database
    OpenCoreRecoveryDatabase recoveryWs, recoveryDb
    EnsureCoreRecoveryStorage recoveryDb
    recoveryDb.Close
    Set recoveryDb = Nothing
    recoveryWs.Close
    Set recoveryWs = Nothing
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not recoveryDb Is Nothing Then recoveryDb.Close
    If Not recoveryWs Is Nothing Then recoveryWs.Close
    On Error GoTo 0
    Err.Raise ERR_DURABLE_RECOVERY_FAILED, "EnsureCoreRecoveryStorageReady", "Durable CodeDb recovery storage could not be prepared through the Core recovery infrastructure." & vbCrLf & "Original error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription
End Sub

Private Sub EnsureCoreRecoveryStorage(ByVal recoveryDb As DAO.Database)
    'Intended use: Provides the single physical Setup_RecoveryState schema authority for Core and UI durable-recovery workflows.
    On Error GoTo ErrHandler
    Dim tableName As String
    Dim tdf As DAO.tableDef
    tableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SetupRecoveryState")
    If SharedDao_TableExists(recoveryDb, tableName) = False Then
        Set tdf = recoveryDb.CreateTableDef(tableName)
        AppendCoreRecoveryField tdf, FIELD_RECOVERY_KEY, dbText, 255, True, False
        AppendCoreRecoveryField tdf, FIELD_RECOVERY_STATUS, dbText, 50, True, False, "In (" & Chr$(34) & RECOVERY_STATUS_BLOCKED & Chr$(34) & ")", "RecoveryStatus must be " & RECOVERY_STATUS_BLOCKED & "."
        AppendCoreRecoveryField tdf, FIELD_UPDATED_ON, dbDate, 0, True
        AppendCoreRecoveryField tdf, FIELD_MANAGED_OBJECT_DEPENDENCY_NOTES, dbMemo, 0, False
        EnsureCoreRecoveryPrimaryIndex tdf
        recoveryDb.TableDefs.Append tdf
        recoveryDb.TableDefs.Refresh
    Else
        recoveryDb.TableDefs.Refresh
        Set tdf = recoveryDb.TableDefs(tableName)
        EnsureCoreRecoveryField tdf, FIELD_RECOVERY_KEY, dbText, 255, True, False
        EnsureCoreRecoveryField tdf, FIELD_RECOVERY_STATUS, dbText, 50, True, False, "In (" & Chr$(34) & RECOVERY_STATUS_BLOCKED & Chr$(34) & ")", "RecoveryStatus must be " & RECOVERY_STATUS_BLOCKED & "."
        EnsureCoreRecoveryField tdf, FIELD_UPDATED_ON, dbDate, 0, True
        EnsureCoreRecoveryField tdf, FIELD_MANAGED_OBJECT_DEPENDENCY_NOTES, dbMemo, 0, False
        EnsureCoreRecoveryPrimaryIndex tdf
    End If
    ' Normal numbered Setup leaves managed Description ownership to G1. Independent recovery callers must still leave this managed table and its fields fully described.
    If mSetupRunContextStarted = False Then EnsureCoreRecoveryStorageDescriptions tdf
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "EnsureCoreRecoveryStorage", "Could not create or repair the canonical Setup_RecoveryState physical schema." & vbCrLf & "Err.Description: " & Err.Description
End Sub

Private Sub EnsureCoreRecoveryStorageDescriptions(ByVal tdf As DAO.tableDef)
    'Standalone recovery-storage creation cannot rely on a later G1 pass. It uses DEFAULTS' canonical managed-object and field-description authorities and verifies each DAO Description immediately.
    Dim tableName As String
    Dim tableDescription As String
    Dim fieldName As Variant
    Dim fieldDescription As String
    Dim fieldDescriptionSource As Variant
    tableName = tdf.Name
    tableDescription = ManagedObjectCanonicalDescription(tableName, ADCAT_OBJECT_TYPE_TABLE)
    If Len(Trim$(tableDescription)) = 0 Then Err.Raise vbObjectError + 10109, "EnsureCoreRecoveryStorageDescriptions", "No canonical managed-object Description exists for the standalone recovery table." & vbCrLf & "Table: " & tableName
    Call EnsureDaoPropertyValue(tdf, DAO_DESCRIPTION_PROPERTY_NAME, dbText, tableDescription, "EnsureCoreRecoveryStorageDescriptions", tableName)
    For Each fieldName In Array(FIELD_RECOVERY_KEY, FIELD_RECOVERY_STATUS, FIELD_UPDATED_ON, FIELD_MANAGED_OBJECT_DEPENDENCY_NOTES)
        fieldDescriptionSource = vbNullString
        fieldDescription = ManagedObjectCatalogFieldDescriptionShort(tableName, CStr(fieldName), fieldDescriptionSource)
        If Len(Trim$(fieldDescription)) = 0 Then Err.Raise vbObjectError + 10111, "EnsureCoreRecoveryStorageDescriptions", "No authoritative short field Description exists for the standalone recovery table." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & CStr(fieldName)
        If Len(fieldDescription) > 255 Then Err.Raise vbObjectError + 10112, "EnsureCoreRecoveryStorageDescriptions", "The authoritative short field Description exceeds the Access DAO Field.Description 255-character limit." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & CStr(fieldName) & vbCrLf & "Characters: " & CStr(Len(fieldDescription)) & vbCrLf & "Default/value source: " & IIf(Len(Trim$(CStr(fieldDescriptionSource))) > 0, CStr(fieldDescriptionSource), "ManagedObjectCatalogFieldDescriptionShort")
        Call EnsureDaoPropertyValue(tdf.fields(CStr(fieldName)), DAO_DESCRIPTION_PROPERTY_NAME, dbText, fieldDescription, "EnsureCoreRecoveryStorageDescriptions", tableName & "." & CStr(fieldName))
    Next fieldName
End Sub

Private Sub AppendCoreRecoveryField(ByVal tdf As DAO.tableDef, ByVal fieldName As String, ByVal fieldType As DAO.DataTypeEnum, ByVal fieldSize As Long, ByVal requiredField As Boolean, Optional ByVal allowZeroLength As Variant, Optional ByVal validationRule As Variant, Optional ByVal validationText As Variant)
    Dim fld As DAO.Field
    If fieldType = dbText And fieldSize > 0 Then Set fld = tdf.CreateField(fieldName, fieldType, fieldSize) Else Set fld = tdf.CreateField(fieldName, fieldType)
    fld.Required = requiredField
    If isMissing(allowZeroLength) = False Then If fld.Type = dbText Or fld.Type = dbMemo Then fld.allowZeroLength = CBool(allowZeroLength)
    If isMissing(validationRule) = False Then fld.validationRule = CStr(validationRule)
    If isMissing(validationText) = False Then fld.validationText = CStr(validationText)
    tdf.fields.Append fld
End Sub

Private Sub EnsureCoreRecoveryField(ByVal tdf As DAO.tableDef, ByVal fieldName As String, ByVal fieldType As DAO.DataTypeEnum, ByVal fieldSize As Long, ByVal requiredField As Boolean, Optional ByVal allowZeroLength As Variant, Optional ByVal validationRule As Variant, Optional ByVal validationText As Variant)
    Dim fld As DAO.Field
    tdf.fields.Refresh
    If SharedDao_TableDefFieldExistsSafe(tdf, fieldName) = False Then
        If isMissing(allowZeroLength) Then
            AppendCoreRecoveryField tdf, fieldName, fieldType, fieldSize, requiredField
        ElseIf isMissing(validationRule) Then
            AppendCoreRecoveryField tdf, fieldName, fieldType, fieldSize, requiredField, allowZeroLength
        ElseIf isMissing(validationText) Then
            AppendCoreRecoveryField tdf, fieldName, fieldType, fieldSize, requiredField, allowZeroLength, validationRule
        Else
            AppendCoreRecoveryField tdf, fieldName, fieldType, fieldSize, requiredField, allowZeroLength, validationRule, validationText
        End If
        tdf.fields.Refresh
    End If
    Set fld = tdf.fields(fieldName)
    If SharedDao_FieldDefinitionMatches(fld, fieldType, fieldSize, False) = False Then Err.Raise vbObjectError + 9881, "EnsureCoreRecoveryField", "Setup_RecoveryState field has an incompatible type or size." & vbCrLf & "Field: " & fieldName
    If fld.Required <> requiredField Then fld.Required = requiredField
    If isMissing(allowZeroLength) = False Then If fld.Type = dbText Or fld.Type = dbMemo Then If fld.allowZeroLength <> CBool(allowZeroLength) Then fld.allowZeroLength = CBool(allowZeroLength)
    If isMissing(validationRule) = False Then If StrComp(CStr(Nz(fld.validationRule, vbNullString)), CStr(validationRule), vbBinaryCompare) <> 0 Then fld.validationRule = CStr(validationRule)
    If isMissing(validationText) = False Then If StrComp(CStr(Nz(fld.validationText, vbNullString)), CStr(validationText), vbBinaryCompare) <> 0 Then fld.validationText = CStr(validationText)
End Sub

Private Sub EnsureCoreRecoveryPrimaryIndex(ByVal tdf As DAO.tableDef)
    Dim idx As DAO.index
    Dim indexName As Variant
    ' INTENTIONAL FILTERED PARTITION: indexesToDelete is a consumer-required subset/projection derived from the source enumerated by EnsureCoreRecoveryPrimaryIndex; its membership is not independently authoritative.
    Dim indexesToDelete As New Collection
    tdf.Indexes.Refresh
    For Each idx In tdf.Indexes
        If CBool(idx.Primary) Then
            If StrComp(idx.Name, INDEX_PRIMARY_KEY, vbTextCompare) <> 0 Or idx.fields.count <> 1 Or StrComp(idx.fields(0).Name, FIELD_RECOVERY_KEY, vbTextCompare) <> 0 Or CBool(idx.Unique) = False Then indexesToDelete.Add idx.Name
        ElseIf StrComp(idx.Name, INDEX_PRIMARY_KEY, vbTextCompare) = 0 Then
            indexesToDelete.Add idx.Name
        End If
    Next idx
    For Each indexName In indexesToDelete
        tdf.Indexes.Delete CStr(indexName)
    Next indexName
    tdf.Indexes.Refresh
    Set idx = Nothing
    On Error Resume Next
    Set idx = tdf.Indexes(INDEX_PRIMARY_KEY)
    On Error GoTo 0
    If idx Is Nothing Then
        Set idx = tdf.CreateIndex(INDEX_PRIMARY_KEY)
        idx.Primary = True
        idx.Unique = True
        idx.fields.Append idx.CreateField(FIELD_RECOVERY_KEY)
        tdf.Indexes.Append idx
        tdf.Indexes.Refresh
    End If
End Sub


Private Function RecoveryTypedValue(ByVal sourceValue As Variant) As String
    Select Case VarType(sourceValue)
        Case vbEmpty: RecoveryTypedValue = "E:"
        Case vbNull: RecoveryTypedValue = "N:"
        Case vbBoolean: RecoveryTypedValue = "B:" & IIf(CBool(sourceValue), "1", "0")
        Case vbDate: RecoveryTypedValue = "D:" & CStr(CDbl(CDate(sourceValue)))
        Case vbByte, vbInteger, vbLong: RecoveryTypedValue = "L:" & CStr(CLng(sourceValue))
        Case vbSingle, vbDouble, vbCurrency: RecoveryTypedValue = "R:" & CStr(sourceValue)
        Case Else: RecoveryTypedValue = "S:" & CoreRecoveryEncode(CStr(sourceValue))
    End Select
End Function

Private Function RecoveryValueFromTypedText(ByVal typedText As String) As Variant
    Dim typeCode As String
    Dim valueText As String
    typeCode = Left$(typedText, 2)
    valueText = Mid$(typedText, 3)
    Select Case typeCode
        Case "E:": RecoveryValueFromTypedText = Empty
        Case "N:": RecoveryValueFromTypedText = Null
        Case "B:": RecoveryValueFromTypedText = (valueText = "1")
        Case "D:": RecoveryValueFromTypedText = CDate(CDbl(valueText))
        Case "L:": RecoveryValueFromTypedText = CLng(valueText)
        Case "R:": RecoveryValueFromTypedText = CDbl(valueText)
        Case "S:": RecoveryValueFromTypedText = CoreRecoveryDecode(valueText)
        Case Else: Err.Raise ERR_DURABLE_RECOVERY_FAILED, "RecoveryValueFromTypedText", "Unknown durable recovery value type: " & typeCode
    End Select
End Function

Private Function SerializeRecoveryArray(ByVal linePrefix As String, ByVal entryData As Variant) As String
    Dim itemIndex As Long
    SerializeRecoveryArray = linePrefix
    For itemIndex = LBound(entryData) To UBound(entryData)
        PumpSetupResponsiveness
        SerializeRecoveryArray = SerializeRecoveryArray & Chr$(30) & RecoveryTypedValue(entryData(itemIndex))
    Next itemIndex
End Function

Private Function DeserializeRecoveryArray(ByVal lineText As String) As Variant
    Dim parts As Variant
    Dim result As Variant
    Dim itemIndex As Long
    parts = Split(lineText, Chr$(30))
    ReDim result(0 To UBound(parts) - 1)
    For itemIndex = 1 To UBound(parts)
        PumpSetupResponsiveness
        result(itemIndex - 1) = RecoveryValueFromTypedText(CStr(parts(itemIndex)))
    Next itemIndex
    DeserializeRecoveryArray = result
End Function

Private Function AccessProcessPriorityName(ByVal priorityValue As Long) As String
    Select Case priorityValue
        Case 4: AccessProcessPriorityName = "Idle"
        Case 6: AccessProcessPriorityName = "Below Normal"
        Case 8: AccessProcessPriorityName = "Normal"
        Case 10: AccessProcessPriorityName = "Above Normal"
        Case 13: AccessProcessPriorityName = "High"
        Case 24: AccessProcessPriorityName = "Real Time"
        Case Else: AccessProcessPriorityName = "Other"
    End Select
End Function

Private Function AccessProcessPriorityDetails() As String
    On Error GoTo Unavailable
    Dim service As Object
    Dim processes As Object
    Dim processItem As Object
    Dim selectedProcess As Object
    Dim currentProjectPath As String
    Dim commandLineText As String
    Dim accessProcessCount As Long
    Dim currentProjectMatchCount As Long
    Dim priorityValue As Long
    Dim unavailableReason As String
    currentProjectPath = UCase$(Trim$(GetHostDb().Name))
    Set service = GetObject("winmgmts:\\.\root\cimv2")
    Set processes = service.ExecQuery("SELECT ProcessId, Priority, CommandLine FROM Win32_Process WHERE Name='MSACCESS.EXE'")
    For Each processItem In processes
        PumpSetupResponsiveness
        accessProcessCount = accessProcessCount + 1
        commandLineText = vbNullString
        If IsNull(processItem.commandLine) = False Then commandLineText = CStr(processItem.commandLine)
        If Len(currentProjectPath) > 0 And InStr(1, UCase$(commandLineText), currentProjectPath, vbBinaryCompare) > 0 Then
            currentProjectMatchCount = currentProjectMatchCount + 1
            Set selectedProcess = processItem
        ElseIf accessProcessCount = 1 Then
            Set selectedProcess = processItem
        End If
    Next processItem
    If currentProjectMatchCount = 0 And accessProcessCount <> 1 Then GoTo Ambiguous
    If currentProjectMatchCount > 1 Then GoTo Ambiguous
    If selectedProcess Is Nothing Then
        unavailableReason = "No matching MSACCESS.EXE process was returned."
        GoTo Unavailable
    End If
    If IsNull(selectedProcess.Priority) Then
        unavailableReason = "WMI returned a Null process priority value."
        GoTo Unavailable
    End If
    priorityValue = CLng(selectedProcess.Priority)
    AccessProcessPriorityDetails = "Access process priority: " & AccessProcessPriorityName(priorityValue) & " (" & CStr(priorityValue) & ")" & vbCrLf & "Process ID: " & CStr(selectedProcess.processID) & vbCrLf & "MSACCESS.EXE processes found: " & CStr(accessProcessCount) & vbCrLf & "Current database command-line matches: " & CStr(currentProjectMatchCount)
    Exit Function
Ambiguous:
    AccessProcessPriorityDetails = "Access process priority: unavailable because the current MSACCESS.EXE process could not be identified uniquely." & vbCrLf & "MSACCESS.EXE processes found: " & CStr(accessProcessCount) & vbCrLf & "Current database command-line matches: " & CStr(currentProjectMatchCount)
    Exit Function
Unavailable:
    If Len(unavailableReason) = 0 Then unavailableReason = Err.Description
    If Len(unavailableReason) = 0 Then unavailableReason = "WMI did not return a usable process-priority result."
    AccessProcessPriorityDetails = "Access process priority: unavailable." & vbCrLf & "Reason: " & unavailableReason
End Function

Private Function RollbackObjectExistsStrict(ByVal db As DAO.Database, ByVal objectType As AcObjectType, ByVal objectName As String, ByVal stageText As String) As Boolean
    On Error GoTo ErrHandler
    Dim qdf As DAO.queryDef
    Dim doc As DAO.Document
    Dim containerObject As DAO.Container
    Dim containerName As String
    Dim objectLookupStarted As Boolean
    objectName = Trim$(objectName)
    If Len(objectName) = 0 Then Err.Raise vbObjectError + 6410, "RollbackObjectExistsStrict", "A strict rollback object inspection received a blank object name." & vbCrLf & "Stage: " & stageText
    If db Is Nothing Then Err.Raise vbObjectError + 6424, "RollbackObjectExistsStrict", "A strict rollback object inspection received no database object." & vbCrLf & "Object: " & objectName & vbCrLf & "Stage: " & stageText
    Select Case objectType
        Case acTable
            db.TableDefs.Refresh
            objectLookupStarted = True
            RollbackObjectExistsStrict = DaoTableExists(db, objectName)
            Exit Function
        Case acQuery
            db.QueryDefs.Refresh
            objectLookupStarted = True
            Set qdf = db.QueryDefs(objectName)
        Case acForm
            containerName = FORM_DAO_CONTAINER_NAME
            Set containerObject = db.Containers(containerName)
            containerObject.Documents.Refresh
            objectLookupStarted = True
            Set doc = containerObject.Documents(objectName)
        Case acReport
            containerName = REPORT_DAO_CONTAINER_NAME
            Set containerObject = db.Containers(containerName)
            containerObject.Documents.Refresh
            objectLookupStarted = True
            Set doc = containerObject.Documents(objectName)
        Case acMacro
            containerName = MACRO_DAO_CONTAINER_NAME
            Set containerObject = db.Containers(containerName)
            containerObject.Documents.Refresh
            objectLookupStarted = True
            Set doc = containerObject.Documents(objectName)
        Case Else
            Err.Raise vbObjectError + 6411, "RollbackObjectExistsStrict", "Unsupported strict rollback Access object type: " & CStr(objectType) & vbCrLf & "Object: " & objectName & vbCrLf & "Stage: " & stageText
    End Select
    RollbackObjectExistsStrict = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If savedErrNumber = 3265 And objectLookupStarted Then Exit Function
    Err.Raise savedErrNumber, "RollbackObjectExistsStrict", "Could not authoritatively inspect an Access object during rollback. The error was not treated as object absence." & vbCrLf & "Database: " & SharedDao_SafeDatabaseName(db) & vbCrLf & "Object type: " & CStr(objectType) & vbCrLf & "Object: " & objectName & vbCrLf & "Stage: " & stageText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Function

Private Sub VerifyRollbackQueryOpenable(ByVal qdf As DAO.queryDef, ByVal queryName As String)
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
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
    Err.Raise savedErrNumber, "VerifyRollbackQueryOpenable", "The restored query could not be opened after rollback." & vbCrLf & "Query: " & queryName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Sub

 ' Purpose: Records the exact prior Setup F1-owned MAC_Deps generated-row set so a later failure in the same rollback group can restore Phase F without touching manual dependency rows.
' Public role: Inter-module public API used only by Setup F1 dependency reconciliation.
Public Sub TrackSetupManagedObjectDependenciesRollback(ByVal rollbackSql As String)
    Dim entryData As Variant
    Dim existingEntry As Variant
    If mSetupRollbackActive = False Then Exit Sub
    ValidateSetupRollbackOwnership "TrackSetupManagedObjectDependenciesRollback"
    If Len(rollbackSql) = 0 Then Err.Raise vbObjectError + 6492, "TrackSetupManagedObjectDependenciesRollback", "The managed-dependency rollback SQL is blank."
    If mSetupRollbackEntries Is Nothing Then Set mSetupRollbackEntries = New Collection
    For Each existingEntry In mSetupRollbackEntries
        If IsArray(existingEntry) Then
            If UBound(existingEntry) >= 0 Then
                If StrComp(CStr(existingEntry(0)), SETUP_ROLLBACK_RESTORE_MANAGED_DEPENDENCIES_ACTION, vbBinaryCompare) = 0 Then Exit Sub
            End If
        End If
    Next existingEntry
    entryData = Array(SETUP_ROLLBACK_RESTORE_MANAGED_DEPENDENCIES_ACTION, rollbackSql)
    mSetupRollbackEntries.Add entryData
    PersistDurableRecoveryMarker SETUP_RECOVERY_KEY, mSetupRollbackOperationName, SETUP_ROLLBACK_ACTIVE_DETAIL, SerializeSetupRollbackEntries()
End Sub

' Purpose: Records setup managed object row rollback, coordinating `ValidateSetupRollbackOwnership` and `IsArray` as the implementation requires.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   managedObjectID (Input; Long): Primary-key ID of the managed-object catalog row targeted by the operation/rollback record.
'   priorRowExisted (Input; Boolean): Records whether the target/row existed before the change so rollback can restore the prior object/row rather than treating it as newly created.
'   rollbackSql (Input; String): SQL text or SQL-derived value used to define, validate, compare, or execute the rollback SQL operation.
' Returns: Boolean: True on the function's affirmative/success branch for track setup managed object row rollback; False on its explicit negative, unavailable, rejected, or handled-failure branch.
'   path.
Public Function TrackSetupManagedObjectRowRollback(ByVal managedObjectID As Long, ByVal priorRowExisted As Boolean, ByVal rollbackSql As String) As Boolean
    'Records the earliest exact prior utility-controlled MAC row state for durable setup rollback and returns True only when a new registration is created.
    Dim entryData As Variant
    Dim existingEntry As Variant
    Dim duplicateEntry As Boolean
    If mSetupRollbackActive = False Then Exit Function
    ValidateSetupRollbackOwnership "TrackSetupManagedObjectRowRollback"
    If managedObjectID <= 0 Then Err.Raise vbObjectError + 6433, "TrackSetupManagedObjectRowRollback", "A managed-object rollback entry requires a positive ManagedObjectID."
    If priorRowExisted And Len(Trim$(rollbackSql)) = 0 Then Err.Raise vbObjectError + 6434, "TrackSetupManagedObjectRowRollback", "An existing managed-object rollback entry requires exact restore SQL."
    If mSetupRollbackEntries Is Nothing Then Set mSetupRollbackEntries = New Collection
    For Each existingEntry In mSetupRollbackEntries
        duplicateEntry = False
        If IsArray(existingEntry) Then
            On Error Resume Next
            duplicateEntry = (StrComp(CStr(existingEntry(0)), SETUP_ROLLBACK_RESTORE_MANAGED_OBJECT_ROW_ACTION, vbBinaryCompare) = 0 And CLng(existingEntry(1)) = managedObjectID)
            On Error GoTo 0
        End If
        If duplicateEntry Then Exit Function
    Next existingEntry
    entryData = Array(SETUP_ROLLBACK_RESTORE_MANAGED_OBJECT_ROW_ACTION, managedObjectID, priorRowExisted, rollbackSql)
    mSetupRollbackEntries.Add entryData
    PersistDurableRecoveryMarker SETUP_RECOVERY_KEY, mSetupRollbackOperationName, SETUP_ROLLBACK_ACTIVE_DETAIL, SerializeSetupRollbackEntries()
    TrackSetupManagedObjectRowRollback = True
End Function

' Purpose: Records setup nav pane shortcut rollback, coordinating `ValidateSetupRollbackOwnership` and `IsArray` as the implementation requires.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   groupID (Input; Long): Identifier selecting the logical Ribbon, form, catalog, or definition group affected by the operation.
'   objectID (Input; Long): Access/Navigation Pane object ID captured for the target object so placement and rollback can restore the exact prior identity/state.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub TrackSetupNavPaneShortcutRollback(ByVal groupID As Long, ByVal objectID As Long)
    'Records a Navigation Pane membership created by setup so rollback can delete only that exact GroupID/ObjectID pair.
    Dim entryData As Variant
    Dim existingEntry As Variant
    Dim duplicateEntry As Boolean
    If mSetupRollbackActive = False Then Exit Sub
    ValidateSetupRollbackOwnership "TrackSetupNavPaneShortcutRollback"
    If groupID <= 0 Or objectID = 0 Then Err.Raise vbObjectError + 6441, "TrackSetupNavPaneShortcutRollback", "A Navigation Pane shortcut rollback entry requires a positive GroupID and a nonzero ObjectID."
    If mSetupRollbackEntries Is Nothing Then Set mSetupRollbackEntries = New Collection
    For Each existingEntry In mSetupRollbackEntries
        duplicateEntry = False
        If IsArray(existingEntry) Then
            On Error Resume Next
            duplicateEntry = (StrComp(CStr(existingEntry(0)), SETUP_ROLLBACK_DELETE_NAV_SHORTCUT_ACTION, vbBinaryCompare) = 0 And CLng(existingEntry(1)) = groupID And CLng(existingEntry(2)) = objectID)
            On Error GoTo 0
        End If
        If duplicateEntry Then Exit Sub
    Next existingEntry
    entryData = Array(SETUP_ROLLBACK_DELETE_NAV_SHORTCUT_ACTION, groupID, objectID)
    mSetupRollbackEntries.Add entryData
    PersistDurableRecoveryMarker SETUP_RECOVERY_KEY, mSetupRollbackOperationName, SETUP_ROLLBACK_ACTIVE_DETAIL, SerializeSetupRollbackEntries()
End Sub

' Purpose: Records setup nav pane shortcut restore rollback, coordinating `ValidateSetupRollbackOwnership` and `IsArray` as the implementation requires.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   groupID (Input; Long): Identifier selecting the logical Ribbon, form, catalog, or definition group affected by the operation.
'   objectID (Input; Long): Access/Navigation Pane object ID captured for the target object so placement and rollback can restore the exact prior identity/state.
'   priorName (Input; Variant): Pre-change object/name value captured for rollback; may be Empty/Null when no prior value existed.
'   priorType (Input; Variant): Pre-change object/property type captured for rollback; Variant preserves Empty/Null when no prior typed value existed.
'   priorFlags (Input; Variant): Bit flags preserving or selecting the prior attributes used by the underlying Access/DAO operation.
'   priorPosition (Input; Variant): Numeric position locating the prior item within its source, UI layout, or Navigation Pane ordering.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub TrackSetupNavPaneShortcutRestoreRollback(ByVal groupID As Long, ByVal objectID As Long, ByVal priorName As Variant, ByVal priorType As Variant, ByVal priorFlags As Variant, ByVal priorPosition As Variant)
    'Records one preexisting Navigation Pane membership removed for a verified stale identity repair so rollback can restore that exact row through the existing durable journal.
    Dim entryData As Variant
    Dim existingEntry As Variant
    Dim duplicateEntry As Boolean
    If mSetupRollbackActive = False Then Exit Sub
    ValidateSetupRollbackOwnership "TrackSetupNavPaneShortcutRestoreRollback"
    If groupID <= 0 Or objectID = 0 Then Err.Raise vbObjectError + 6458, "TrackSetupNavPaneShortcutRestoreRollback", "A Navigation Pane shortcut restore entry requires a positive GroupID and a nonzero ObjectID."
    If mSetupRollbackEntries Is Nothing Then Set mSetupRollbackEntries = New Collection
    For Each existingEntry In mSetupRollbackEntries
        duplicateEntry = False
        If IsArray(existingEntry) Then
            On Error Resume Next
            duplicateEntry = (StrComp(CStr(existingEntry(0)), SETUP_ROLLBACK_RESTORE_NAV_SHORTCUT_ACTION, vbBinaryCompare) = 0 And CLng(existingEntry(1)) = groupID And CLng(existingEntry(2)) = objectID)
            On Error GoTo 0
        End If
        If duplicateEntry Then Exit Sub
    Next existingEntry
    entryData = Array(SETUP_ROLLBACK_RESTORE_NAV_SHORTCUT_ACTION, groupID, objectID, priorName, priorType, priorFlags, priorPosition)
    mSetupRollbackEntries.Add entryData
    PersistDurableRecoveryMarker SETUP_RECOVERY_KEY, mSetupRollbackOperationName, SETUP_ROLLBACK_ACTIVE_DETAIL, SerializeSetupRollbackEntries()
End Sub

' Purpose: Records setup nav pane object ID rollback, coordinating `ValidateSetupRollbackOwnership` and `IsArray` as the implementation requires.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   objectID (Input; Long): Access/Navigation Pane object ID captured for the target object so placement and rollback can restore the exact prior identity/state.
'   objectName (Input; String): Physical Access object name identifying the database object targeted by the operation.
'   objectType (Input; Long): Object-type value identifying the Access/VBA object class targeted by the operation.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub TrackSetupNavPaneObjectIDRollback(ByVal objectID As Long, ByVal objectName As String, ByVal objectType As Long)
    'Records an MSysNavPaneObjectIDs row materialized by setup so rollback can delete only that exact identity after memberships are removed.
    Dim entryData As Variant
    Dim existingEntry As Variant
    Dim duplicateEntry As Boolean
    objectName = Trim$(objectName)
    If mSetupRollbackActive = False Then Exit Sub
    ValidateSetupRollbackOwnership "TrackSetupNavPaneObjectIDRollback"
    If objectID = 0 Or Len(objectName) = 0 Or objectType = 0 Then Err.Raise vbObjectError + 6447, "TrackSetupNavPaneObjectIDRollback", "A Navigation Pane object-identity rollback entry requires a nonzero ObjectID, nonblank object name, and nonzero object type."
    If mSetupRollbackEntries Is Nothing Then Set mSetupRollbackEntries = New Collection
    For Each existingEntry In mSetupRollbackEntries
        duplicateEntry = False
        If IsArray(existingEntry) Then
            On Error Resume Next
            duplicateEntry = (StrComp(CStr(existingEntry(0)), SETUP_ROLLBACK_DELETE_NAV_OBJECT_ID_ACTION, vbBinaryCompare) = 0 And CLng(existingEntry(1)) = objectID And StrComp(CStr(existingEntry(2)), objectName, vbTextCompare) = 0 And CLng(existingEntry(3)) = objectType)
            On Error GoTo 0
        End If
        If duplicateEntry Then Exit Sub
    Next existingEntry
    entryData = Array(SETUP_ROLLBACK_DELETE_NAV_OBJECT_ID_ACTION, objectID, objectName, objectType)
    mSetupRollbackEntries.Add entryData
    PersistDurableRecoveryMarker SETUP_RECOVERY_KEY, mSetupRollbackOperationName, SETUP_ROLLBACK_ACTIVE_DETAIL, SerializeSetupRollbackEntries()
End Sub

' Purpose: Records setup nav pane object ID restore rollback, coordinating `ValidateSetupRollbackOwnership` and `IsArray` as the implementation requires.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   objectID (Input; Long): Access/Navigation Pane object ID captured for the target object so placement and rollback can restore the exact prior identity/state.
'   repairedObjectName (Input; String): Physical object name after repair/rename; rollback bookkeeping uses it to locate the repaired object.
'   repairedObjectType (Input; Long): Access object-type code for the repaired/current object, used by rollback to locate/delete/restore the correct object class.
'   priorObjectName (Input; String): Physical object name that existed before repair/rename and must be restored if rollback occurs.
'   priorObjectType (Input; Long): Access object-type code for the pre-change object, used by rollback when restoring its prior identity/state.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub TrackSetupNavPaneObjectIDRestoreRollback(ByVal objectID As Long, ByVal repairedObjectName As String, ByVal repairedObjectType As Long, ByVal priorObjectName As String, ByVal priorObjectType As Long)
    'Records the exact prior MSysNavPaneObjectIDs Name/Type before setup repairs a stale row in place. The earliest prior state wins.
    Dim entryData As Variant
    Dim existingEntry As Variant
    Dim duplicateEntry As Boolean
    repairedObjectName = Trim$(repairedObjectName)
    If mSetupRollbackActive = False Then Exit Sub
    ValidateSetupRollbackOwnership "TrackSetupNavPaneObjectIDRestoreRollback"
    If objectID = 0 Or Len(repairedObjectName) = 0 Or repairedObjectType = 0 Then Err.Raise vbObjectError + 6452, "TrackSetupNavPaneObjectIDRestoreRollback", "A Navigation Pane object-identity restore entry requires a nonzero ObjectID, nonblank repaired name, and nonzero repaired type. Prior stale metadata is preserved even when blank or zero."
    If mSetupRollbackEntries Is Nothing Then Set mSetupRollbackEntries = New Collection
    For Each existingEntry In mSetupRollbackEntries
        duplicateEntry = False
        If IsArray(existingEntry) Then
            On Error Resume Next
            duplicateEntry = (StrComp(CStr(existingEntry(0)), SETUP_ROLLBACK_RESTORE_NAV_OBJECT_ID_ACTION, vbBinaryCompare) = 0 And CLng(existingEntry(1)) = objectID)
            On Error GoTo 0
        End If
        If duplicateEntry Then Exit Sub
    Next existingEntry
    entryData = Array(SETUP_ROLLBACK_RESTORE_NAV_OBJECT_ID_ACTION, objectID, repairedObjectName, repairedObjectType, priorObjectName, priorObjectType)
    mSetupRollbackEntries.Add entryData
    PersistDurableRecoveryMarker SETUP_RECOVERY_KEY, mSetupRollbackOperationName, SETUP_ROLLBACK_ACTIVE_DETAIL, SerializeSetupRollbackEntries()
End Sub

' Purpose: Records setup deleted nav pane object ID restore rollback, coordinating `ValidateSetupRollbackOwnership` and `IsArray` as the implementation requires.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   objectID (Input; Long): Access/Navigation Pane object ID captured for the target object so placement and rollback can restore the exact prior identity/state.
'   objectName (Input; String): Physical Access object name identifying the database object targeted by the operation.
'   objectType (Input; Long): Object-type value identifying the Access/VBA object class targeted by the operation.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub TrackSetupDeletedNavPaneObjectIDRestoreRollback(ByVal objectID As Long, ByVal objectName As String, ByVal objectType As Long)
    'Records an exact preexisting MSysNavPaneObjectIDs row removed during stale logical-identity remapping so rollback can reinsert it before restoring its memberships.
    Dim entryData As Variant
    Dim existingEntry As Variant
    Dim duplicateEntry As Boolean
    objectName = Trim$(objectName)
    If mSetupRollbackActive = False Then Exit Sub
    ValidateSetupRollbackOwnership "TrackSetupDeletedNavPaneObjectIDRestoreRollback"
    If objectID = 0 Or Len(objectName) = 0 Or objectType = 0 Then Err.Raise vbObjectError + 6470, "TrackSetupDeletedNavPaneObjectIDRestoreRollback", "A deleted Navigation Pane object-identity restore entry requires a nonzero ObjectID, nonblank object name, and nonzero object type."
    If mSetupRollbackEntries Is Nothing Then Set mSetupRollbackEntries = New Collection
    For Each existingEntry In mSetupRollbackEntries
        duplicateEntry = False
        If IsArray(existingEntry) Then
            On Error Resume Next
            duplicateEntry = (StrComp(CStr(existingEntry(0)), SETUP_ROLLBACK_RESTORE_DELETED_NAV_OBJECT_ID_ACTION, vbBinaryCompare) = 0 And CLng(existingEntry(1)) = objectID And StrComp(CStr(existingEntry(2)), objectName, vbTextCompare) = 0 And CLng(existingEntry(3)) = objectType)
            On Error GoTo 0
        End If
        If duplicateEntry Then Exit Sub
    Next existingEntry
    entryData = Array(SETUP_ROLLBACK_RESTORE_DELETED_NAV_OBJECT_ID_ACTION, objectID, objectName, objectType)
    mSetupRollbackEntries.Add entryData
    PersistDurableRecoveryMarker SETUP_RECOVERY_KEY, mSetupRollbackOperationName, SETUP_ROLLBACK_ACTIVE_DETAIL, SerializeSetupRollbackEntries()
End Sub

' Purpose: Records setup nav pane group rollback, coordinating `ValidateSetupRollbackOwnership` and `IsArray` as the implementation requires.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   groupID (Input; Long): Identifier selecting the logical Ribbon, form, catalog, or definition group affected by the operation.
'   groupName (Input; String): Name selecting the logical Ribbon, form, catalog, or definition group affected by the operation.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub TrackSetupNavPaneGroupRollback(ByVal groupID As Long, ByVal groupName As String)
    'Records a custom Navigation Pane group created by setup. Rollback removes it only after all tracked memberships have been removed.
    Dim entryData As Variant
    Dim existingEntry As Variant
    Dim duplicateEntry As Boolean
    groupName = Trim$(groupName)
    If mSetupRollbackActive = False Then Exit Sub
    ValidateSetupRollbackOwnership "TrackSetupNavPaneGroupRollback"
    If groupID <= 0 Or Len(groupName) = 0 Then Err.Raise vbObjectError + 6442, "TrackSetupNavPaneGroupRollback", "A Navigation Pane group rollback entry requires a positive GroupID and nonblank group name."
    If mSetupRollbackEntries Is Nothing Then Set mSetupRollbackEntries = New Collection
    For Each existingEntry In mSetupRollbackEntries
        duplicateEntry = False
        If IsArray(existingEntry) Then
            On Error Resume Next
            duplicateEntry = (StrComp(CStr(existingEntry(0)), SETUP_ROLLBACK_DELETE_NAV_GROUP_ACTION, vbBinaryCompare) = 0 And CLng(existingEntry(1)) = groupID)
            On Error GoTo 0
        End If
        If duplicateEntry Then Exit Sub
    Next existingEntry
    entryData = Array(SETUP_ROLLBACK_DELETE_NAV_GROUP_ACTION, groupID, groupName)
    mSetupRollbackEntries.Add entryData
    PersistDurableRecoveryMarker SETUP_RECOVERY_KEY, mSetupRollbackOperationName, SETUP_ROLLBACK_ACTIVE_DETAIL, SerializeSetupRollbackEntries()
End Sub

Private Function DirectRunCodeMacroText(ByVal runCodeExpression As String) As String
    DirectRunCodeMacroText = "Version =196611" & vbCrLf & "PublishOption =1" & vbCrLf & "ColumnsShown =0" & vbCrLf & "Begin" & vbCrLf & "    Action =" & Chr$(34) & "RunCode" & Chr$(34) & vbCrLf & "    Argument =" & Chr$(34) & Replace(runCodeExpression, Chr$(34), Chr$(92) & Chr$(34)) & Chr$(34) & vbCrLf & "End" & vbCrLf
End Function

' ===============================================================================================================================================================================
' FINAL J1 STAGING / RUN-THROUGH-I CONTROL PROCEDURES
' ===============================================================================================================================================================================

' Purpose: Performs the public core stage final setup state for i 1 operation using the routine's existing ADC workflow and side effects.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   finalProofToken (Input; String): Proof/state token representing the final proof token prerequisite or resulting state; callers use it to prove that coordinated setup work still refers to the same
'   validated state.
'   moduleGuardStateToken (Input; String): Proof/state token representing the module guard state token prerequisite or resulting state; callers use it to prove that coordinated setup work still refers to
'   the same validated state.
'   detailText (Input; String): Carries human-readable diagnostic detail for the operation; when passed ByRef, the routine can return or augment the text for the caller.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub CoreStageFinalSetupStateForJ1(ByVal finalProofToken As String, ByVal moduleGuardStateToken As String, ByVal detailText As String)
    RequireSetupDomainMutationAllowed SETUP_DOMAIN_J1_FINAL_STATE, "CoreStageFinalSetupStateForJ1"
    finalProofToken = Trim$(finalProofToken)
    moduleGuardStateToken = Trim$(moduleGuardStateToken)
    If Len(finalProofToken) = 0 Then Err.Raise vbObjectError + 6554, "CoreStageFinalSetupStateForJ1", "J1 final verification token is blank."
    If Len(moduleGuardStateToken) = 0 Then Err.Raise vbObjectError + 6555, "CoreStageFinalSetupStateForJ1", "J1 Module Guard state token is blank."
    mFinalSetupJ1ProofToken = finalProofToken
    mFinalSetupJ1ModuleGuardStateToken = moduleGuardStateToken
    mFinalSetupJ1Detail = detailText
    RegisterSetupEvidence "Setup", "J1FinalProof", "A1-J1 final committed state", SETUP_EVIDENCE_STATUS_POSTCONDITION_VERIFIED, authoritativeSourceFingerprint:=finalProofToken, physicalStateFingerprint:=finalProofToken, mutationOwner:="J1", verificationOwner:="x__UTIL_ADC_SETUP.VerifySetupJ1FinalState", dependencies:="Committed A1-I3 outcomes including E3 Reports; Module Guard preflight state", sourceProcedure:="CoreStageFinalSetupStateForJ1", detailText:=detailText
End Sub

Private Sub PersistStagedFinalSetupStateForJ1()
    RequireSetupDomainMutationAllowed SETUP_DOMAIN_J1_FINAL_STATE, "PersistStagedFinalSetupStateForJ1"
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim completedOnText As String
    Dim verifyToken As String
    Dim verifyCheckpoint As String
    Dim propertyExists As Boolean
    If Len(mFinalSetupJ1ProofToken) = 0 Or Len(mFinalSetupJ1ModuleGuardStateToken) = 0 Then Err.Raise vbObjectError + 6556, "PersistStagedFinalSetupStateForJ1", "J1 final setup proof was not staged before completion."
    Set db = currentDb
    completedOnText = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Call EnsureDatabasePropertyRemoved(db, FINAL_SETUP_STATE_TOKEN_PROPERTY, "PersistStagedFinalSetupStateForJ1")
    Call EnsureDaoPropertyValue(db, FINAL_SETUP_STATE_VERSION_PROPERTY, dbText, FINAL_SETUP_STATE_VERSION, "PersistStagedFinalSetupStateForJ1")
    Call EnsureDaoPropertyValue(db, FINAL_SETUP_CHECKPOINT_PROPERTY, dbText, mFinalSetupJ1ProofToken, "PersistStagedFinalSetupStateForJ1")
    Call EnsureDaoPropertyValue(db, FINAL_SETUP_COMPLETED_PART_PROPERTY, dbText, "J1", "PersistStagedFinalSetupStateForJ1")
    Call EnsureDaoPropertyValue(db, FINAL_SETUP_COMPLETED_ON_PROPERTY, dbText, completedOnText, "PersistStagedFinalSetupStateForJ1")
    Call EnsureDaoPropertyValue(db, FINAL_SETUP_MODULE_GUARD_PROPERTY, dbText, mFinalSetupJ1ModuleGuardStateToken, "PersistStagedFinalSetupStateForJ1")
    Call EnsureDaoPropertyValue(db, FINAL_SETUP_STATE_TOKEN_PROPERTY, dbText, mFinalSetupJ1ProofToken, "PersistStagedFinalSetupStateForJ1")
    verifyToken = DatabaseTextPropertyValue(db, FINAL_SETUP_STATE_TOKEN_PROPERTY, propertyExists)
    If propertyExists = False Or StrComp(verifyToken, mFinalSetupJ1ProofToken, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 6557, "PersistStagedFinalSetupStateForJ1", "The final J1 setup-state token did not persist exactly."
    verifyCheckpoint = DatabaseTextPropertyValue(db, FINAL_SETUP_CHECKPOINT_PROPERTY, propertyExists)
    If propertyExists = False Or StrComp(verifyCheckpoint, mFinalSetupJ1ProofToken, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 6558, "PersistStagedFinalSetupStateForJ1", "The final J1 setup checkpoint did not persist exactly."
    RegisterSetupEvidence "Setup", "J1FinalProof", "Persisted final setup checkpoint", SETUP_EVIDENCE_STATUS_POSTCONDITION_VERIFIED, authoritativeSourceFingerprint:=mFinalSetupJ1ProofToken, storedStateFingerprint:=verifyToken, physicalStateFingerprint:=verifyCheckpoint, mutationOwner:="J1", verificationOwner:="x__UTIL_ADC_Core", dependencies:="Successful J1 final verification", sourceProcedure:="PersistStagedFinalSetupStateForJ1", detailText:="CompletedOn=" & completedOnText & vbCrLf & mFinalSetupJ1Detail
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not db Is Nothing Then
        Call EnsureDatabasePropertyRemoved(db, FINAL_SETUP_STATE_TOKEN_PROPERTY, "PersistStagedFinalSetupStateForJ1")
    End If
    On Error GoTo 0
    Err.Raise savedErrNumber, "PersistStagedFinalSetupStateForJ1", "Could not persist the final J1 setup state. The final completion-token property was removed so this run cannot be mistaken for a completed J1 checkpoint." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub ResetFinalSetupJ1Staging()
    mFinalSetupJ1ProofToken = vbNullString
    mFinalSetupJ1ModuleGuardStateToken = vbNullString
    mFinalSetupJ1Detail = vbNullString
End Sub

' Purpose: Performs the public core begin setup run through phase i operation using the routine's existing ADC workflow and side effects.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   moduleGuardVerificationDetails (Input; String): Carries human-readable diagnostic detail for the operation; when passed ByRef, the routine can return or augment the text for the caller.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub CoreBeginSetupRunThroughPhaseJ(ByVal moduleGuardVerificationDetails As String)
    mSetupCancelled = False
    mSetupCurrentPartText = vbNullString
    mSetupPartRollbackStartIndex = 0
    mSetupRunContextStarted = False
    mSetupNamingSnapshotStarted = False
    mSetupMaterialCommitted = False
    mSetupCurrentStep = vbNullString
    ResetFinalSetupJ1Staging
    ResetTraceNotes
    mProgressOwnsTraceRun = False
    CoreResetOperationNotes
    CoreBeginSetupNamingSnapshot
    BeginSetupRunContext moduleGuardVerificationDetails
    mSetupRunContextStarted = True
    BeginSetupOperation DEFAULT_SETUP_FULL_OPERATION_NAME
    CoreAddOperationNote "Preflight and A1-I3 establish the committed managed database, Forms, Reports, descriptions, ADC Ribbon, and Navigation Pane state. J1 verifies committed outcomes read-only, records final setup evidence, and closes the setup lifecycle without predecessor repair."
End Sub

' Purpose: Performs the public core complete setup run through phase i operation using the routine's existing ADC workflow and side effects.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   timingText (Input Optional; String): Input timing text for core complete setup run through phase I; is tested by If Len(Trim$(timingText)) > 0. Default: vbNullString.
'   showMessage (Input Optional; Boolean): Controls whether the routine shows its user-facing completion, status, warning, or error message in addition to performing the underlying work. Default: False.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub CoreCompleteSetupRunThroughPhaseJ(Optional ByVal timingText As String = vbNullString, Optional ByVal showMessage As Boolean = False)
    Dim totalElapsedText As String
    On Error GoTo ErrHandler
    RequireSetupControlContinue IIf(Len(Trim$(mSetupCurrentStep)) = 0, "J1 - Phase J completion", mSetupCurrentStep)
    CommitSetupOperation
    mSetupMaterialCommitted = True
    PersistStagedFinalSetupStateForJ1
    If Len(Trim$(timingText)) > 0 Then LogTraceMessage TRACE_LEVEL_INFO, "SetupTiming", "CoreCompleteSetupRunThroughPhaseJ", "Preflight and Phases A-J timing captured.", timingText, "Complete Phase J", 0, vbNullString, vbNullString, vbNullString, "Setup", "CodeDb", "SetupTiming"
    If mSetupRunContextStarted Then FinalizeSetupRunContext "Completed", True
    mSetupRunContextStarted = False
    If mSetupNamingSnapshotStarted Then InvalidateNamingSnapshot
    mSetupNamingSnapshotStarted = False
    ReleaseSetupPreflightNamingSnapshot True
    CoreAddOperationNote "Preflight and Phases A-J completed. J1 persisted the final verified setup checkpoint and performed no predecessor repair."
    totalElapsedText = FormattedElapsedSeconds(ElapsedSecondsFromTimer(mCurrentProcessStartTimer))
    CoreCloseSetupCompletionUI
    If showMessage Then ReportCoreActionSuccess "CoreCompleteSetupRunThroughPhaseJ", "Setup", DEFAULT_UTILITY_SHORT_NAME & " Setup Preflight and Phases A-J completed successfully.", "A1-A4, B1-B5, C1-C3, D1-D4, E1-E3, F1, G1, H1-H3, I1-I3, and J1 completed. J1 verified and persisted the final committed setup state." & vbCrLf & vbCrLf & "Total elapsed time: " & totalElapsedText, "Complete Phase J", ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config"), "Table", "CodeDb", "Setup", True, DEFAULT_UTILITY_LONG_NAME & ": Setup - Through Phase J", visibleStatusMessage:="Setup Preflight and Phases A-J completed successfully."
    CoreCompleteSetupProgress
    CoreResetSetupPhaseContext
    ResetFinalSetupJ1Staging
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "CoreCompleteSetupRunThroughPhaseJ", Err.Description
End Sub

' Purpose: Performs the public core handle setup run through phase i failure operation using the routine's existing ADC workflow and side effects.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   errNumber (Input; Long): Receives the VBA error number associated with a failed operation so the caller can preserve and report the original failure.
'   errSource (Input; String): Receives the VBA error source associated with a failed operation so the caller can identify the originating procedure or component.
'   errDescription (Input; String): Receives the VBA error description associated with a failed operation for caller-level diagnostics or rethrow/reporting.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub CoreHandleSetupRunThroughPhaseJFailure(ByVal errNumber As Long, ByVal errSource As String, ByVal errDescription As String)
    ResetFinalSetupJ1Staging
    CoreHandleSetupRunFailure "Phase J", "Phases A-J", errNumber, errSource, errDescription
End Sub

' Purpose: Resolves validation result from target status from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   statusText (Input; String): Status/progress message supplied for the current operation; the routine forwards it to the applicable UI/progress and trace/reporting path.
' Returns: String: Returns the validation result from target status text value. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function ValidationResultFromTargetStatus(ByVal statusText As String) As String
    'Canonical target-resolution status -> persisted LastValidationResult policy.
    'VALID -> Valid; MISSING/AMBIGUOUS/ARGUMENT MISMATCH -> Invalid; VBE ACCESS ERROR -> Error; UNRESOLVED -> NeedsReview; NON-RUNCODE -> Skipped.
    'Already-normalized validation results are preserved; blank or unrecognized statuses require review rather than being treated as definitively invalid.
    Select Case UCase$(Trim$(statusText))
        Case "VALID"
            ValidationResultFromTargetStatus = VALIDATION_RESULT_VALID
        Case "MISSING", "AMBIGUOUS", "ARGUMENT MISMATCH", "INVALID"
            ValidationResultFromTargetStatus = VALIDATION_RESULT_INVALID
        Case "VBE ACCESS ERROR", "ERROR"
            ValidationResultFromTargetStatus = VALIDATION_RESULT_ERROR
        Case "NON-RUNCODE", "SKIPPED"
            ValidationResultFromTargetStatus = VALIDATION_RESULT_SKIPPED
        Case "WARNING"
            ValidationResultFromTargetStatus = VALIDATION_RESULT_WARNING
        Case "UNRESOLVED", "NEEDSREVIEW", "NEEDS REVIEW"
            ValidationResultFromTargetStatus = VALIDATION_RESULT_NEEDS_REVIEW
        Case Else
            ValidationResultFromTargetStatus = VALIDATION_RESULT_NEEDS_REVIEW
    End Select
End Function

' Purpose: Marks ribbon restart required for current run.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub MarkRibbonRestartRequiredForCurrentRun()
    mRibbonRestartRequiredCurrentRun = True
End Sub

' ===============================================================================================================================================================================
' CONSOLIDATED CORE SYSTEM-COMMAND / RECOVERY PROCEDURES
' ===============================================================================================================================================================================

' Purpose: Resolves entry point error message from the supplied/current ADC state, using `FailureProgressContextForMessage` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   headlineText (Input; String): Input headline text for entry point error message; is passed to `EntryPointErrorMessage`.
'   ProcedureName (Input; String): Canonical or physical name identifying the procedure targeted by this operation.
'   errNumber (Input; Long): Receives the VBA error number associated with a failed operation so the caller can preserve and report the original failure.
'   errSource (Input; String): Receives the VBA error source associated with a failed operation so the caller can identify the originating procedure or component.
'   errDescription (Input; String): Receives the VBA error description associated with a failed operation for caller-level diagnostics or rethrow/reporting.
'   detailText (Input Optional; String): Carries human-readable diagnostic detail for the operation; when passed ByRef, the routine can return or augment the text for the caller. Default: vbNullString.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function EntryPointErrorMessage(ByVal headlineText As String, ByVal ProcedureName As String, ByVal errNumber As Long, ByVal errSource As String, ByVal errDescription As String, Optional ByVal detailText As String = vbNullString) As String
    EntryPointErrorMessage = headlineText & FailureProgressContextForMessage() & vbCrLf & vbCrLf & "Procedure: " & ProcedureName & vbCrLf & "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB & vbCrLf & "Database: " & SafeDatabaseName()
    If Len(Trim$(detailText)) > 0 Then EntryPointErrorMessage = EntryPointErrorMessage & vbCrLf & detailText
    EntryPointErrorMessage = EntryPointErrorMessage & vbCrLf & "Err.Number: " & CStr(errNumber) & vbCrLf & "Err.Source: " & errSource & vbCrLf & "Err.Description:" & vbCrLf & errDescription & TableDiagnosticsLocationForMessage() & TraceSummaryForMessage() & CoreOperationSummaryForMessage()
End Function

Private Sub PulseSystemCommandStatus(ByVal statusText As String)
    On Error Resume Next
    CorePulseProgress statusText
    If mSetupControlActive = False Then DoEvents
End Sub

Private Function SetupSystemCommandStatusText(ByVal keyText As String, ByVal statusText As String, ByVal finalStatus As Boolean) As String
    statusText = Trim$(statusText)
    If Len(statusText) = 0 Then Exit Function
    If StrComp(keyText, UCase$(SYS_MACRO_KEY_RUN_FULL_SETUP), vbBinaryCompare) <> 0 Then
        SetupSystemCommandStatusText = statusText
    ElseIf InStr(1, statusText, "Setup Part ", vbTextCompare) > 0 Or InStr(1, statusText, "Setup Preflight", vbTextCompare) > 0 Then
        SetupSystemCommandStatusText = statusText
    ElseIf finalStatus Then
        SetupSystemCommandStatusText = "Setup Part " & CStr(ImplementedSetupPartCount()) & " of " & CStr(ImplementedSetupPartCount()) & " > " & statusText
    Else
        SetupSystemCommandStatusText = "Setup Preflight > " & statusText
    End If
End Function

Private Sub ClearSystemCommandResult()
    On Error Resume Next
    TempVars.Remove TEMPVAR_SYSTEM_COMMAND_RESULT
End Sub

Private Function SystemCommandIsScanCommand(ByVal keyText As String) As Boolean
    Select Case keyText
        Case UCase$(SYS_MACRO_KEY_SCAN_ALL_DB_INVENTORY), _
             UCase$(SYS_MACRO_KEY_SCAN_RUNNER_MACRO_CANDIDATES), _
             UCase$(SYS_MACRO_KEY_SCAN_DB_INVENTORY_ACTIVE_REFERENCES), _
             UCase$(SYS_MACRO_KEY_SCAN_DB_INVENTORY_AVAILABLE_REFERENCES), _
             UCase$(SYS_MACRO_KEY_SCAN_DB_INVENTORY_LOADED_VBA_PROJECTS), _
             UCase$(SYS_MACRO_KEY_SCAN_DB_INVENTORY_ADD_INS_AND_LOADED_VBA_PROJECTS), _
             UCase$(SYS_MACRO_KEY_SCAN_DB_INVENTORY_LOADED_VBA_SOURCE)
            SystemCommandIsScanCommand = True
    End Select
End Function

' Purpose: Publishes the shared System Command cancellation result in Core-owned TempVars state so later command/result consumers can distinguish cancellation from success or failure.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub MarkSystemCommandCancelled()
    'Publishes the cancelled result through Core's shared System Command result-state authority.
    On Error Resume Next
    TempVars.Remove TEMPVAR_SYSTEM_COMMAND_RESULT
    TempVars.Add TEMPVAR_SYSTEM_COMMAND_RESULT, SYSTEM_COMMAND_RESULT_CANCELLED
End Sub

' Purpose: Reads the shared System Command cancellation result through Core's result-state authority.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: Boolean: True when the condition represented by SystemCommandWasCancelled is present under the routine's documented inputs/current state; False when that condition is absent.
Public Function SystemCommandWasCancelled() As Boolean
    'Reads the shared System Command cancellation result through Core's result-state authority.
    On Error GoTo ExitHere
    SystemCommandWasCancelled = (StrComp(Trim$(CStr(TempVars(TEMPVAR_SYSTEM_COMMAND_RESULT).value)), SYSTEM_COMMAND_RESULT_CANCELLED, vbTextCompare) = 0)
ExitHere:
End Function

Private Sub ScanAllDBInventory()
    'Coordinates the complete DB Inventory scan families under the caller-owned progress controller and operation snapshots. Runner Macro Candidate discovery is owned exclusively by System Macro 20.
    Dim foundationDetail As String
    PulseSystemCommandStatus "Scan All DB Inventory | verifying shared reference/add-in foundation"
    If ReferenceInventoryFoundationCurrentForSetup(foundationDetail) = False Then
        Err.Raise vbObjectError + 9799, "ScanAllDBInventory", "The A4 reference/Add-In inventory foundation is not current. Scan All DB Inventory does not repair setup-owned physical objects." & vbCrLf & foundationDetail
    End If
    RequireSafeWorkContinue "ScanAllDBInventory", "Reference/add-in inventory foundation current"
    PulseSystemCommandStatus "Scan All DB Inventory | step 1/5 | Active References | starting"
    ScanDBInventoryActiveReferences False, False, False
    RequireSafeWorkContinue "ScanAllDBInventory", "Active References scan complete"
    PulseSystemCommandStatus "Scan All DB Inventory | step 2/5 | Available References | starting"
    ScanDBInventoryAvailableReferences False, False
    RequireSafeWorkContinue "ScanAllDBInventory", "Available References scan complete"
    PulseSystemCommandStatus "Scan All DB Inventory | step 3/5 | Loaded VBA Projects | starting"
    ScanDBInventoryLoadedVBAProjects False, False, False
    RequireSafeWorkContinue "ScanAllDBInventory", "Loaded VBA Projects scan complete"
    PulseSystemCommandStatus "Scan All DB Inventory | step 4/5 | Add-Ins | starting"
    ScanDBInventoryAddInsAndLoadedVBAProjects False, False, False, False
    RequireSafeWorkContinue "ScanAllDBInventory", "Add-Ins scan complete"
    PulseSystemCommandStatus "Scan All DB Inventory | step 5/5 | Loaded VBA Source | starting"
    ' Reuse the loaded-project inventory produced by step 3. Finalize Add-In-to-loaded-project correlation only after step 5 has made the procedure inventory current.
    ScanDBInventoryLoadedVBASource False, False, False, True
    RequireSafeWorkContinue "ScanAllDBInventory", "Loaded VBA Source scan complete"
    PulseSystemCommandStatus "Scan All DB Inventory | step 5/5 | complete"
End Sub

Private Function BuildRepairHelpDocumentMacros(ByRef detailText As String) As Boolean
    'Operational orchestrator only: delegates B4/C/D4/F/G work to the existing Catalog, physical-macro, logical-catalog, managed-description, Ribbon, and Navigation Pane authorities in phase order.
    On Error GoTo ErrHandler
    Dim catalogDetail As String
    Dim physicalDetail As String
    Dim logicalCatalogDetail As String
    Dim f1Detail As String
    Dim preBuildRepairPredecessorDetail As String
    Dim helpG1ProofDetail As String
    Dim helpG1ProofToken As String
    Dim helpG1ProofTraceRunID As String
    Dim fullPredecessorProofDetail As String
    Dim fullPredecessorProofToken As String
    Dim fullPredecessorProofTraceRunID As String
    Dim ribbonCatalogProofDetail As String
    Dim ribbonCatalogProofToken As String
    Dim ribbonCatalogProofTraceRunID As String
    Dim currentCatalogTokenForPredecessor As String
    Dim preBuildRepairCatalogToken As String
    Dim ribbonDetail As String
    Dim navDetail As String
    Dim managedDefinitionMutationCount As Long
    Dim managedDefinitionDirtyMask As Long
    Dim d4RanThisOperation As Boolean
    Dim f1CatalogDescriptionMutated As Boolean
    Dim f1RefreshedCatalogToken As String
    Dim activeTraceRunID As String
    Dim namingSnapshotOwned As Boolean
    Dim stageTimer As Single
    RequireSameDatabaseUICommand "BuildRepairHelpDocumentMacros", "Building/repairing Help Document Macros"
    activeTraceRunID = GetActiveOrLastTraceRunID()
    If Len(activeTraceRunID) = 0 Then Err.Raise vbObjectError + 6298, "BuildRepairHelpDocumentMacros", "Build/Repair Help Document Macros requires an active trace-run identity before same-run proof orchestration."
    ' INTENTIONAL OPERATION SNAPSHOT: the read-only predecessor proof resolves canonical names repeatedly; COMMON remains the naming authority and this snapshot is released before B4 can mutate managed-definition
    ' naming inputs.
    PulseSystemCommandStatus "Build/Repair Help Document Macros | preparing predecessor naming snapshot"
    stageTimer = Timer
    BeginNamingSnapshot
    namingSnapshotOwned = True
    LogTraceMessage TRACE_LEVEL_INFO, "Performance", "BuildRepairHelpDocumentMacros", "Build/Repair Help Document Macros predecessor naming snapshot prepared.", "COMMON naming authority is retained only for the read-only predecessor proof scope." & vbCrLf & "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(stageTimer)), "Build/Repair Help Document Macros predecessor naming snapshot", 0, vbNullString, vbNullString, vbNullString, "NamingSnapshot", "CodeDb", "Performance"
    PulseSystemCommandStatus "Build/Repair Help Document Macros | proving unchanged predecessor state"
    stageTimer = Timer
    If BeginHelpDocumentMacroBuildRepairFullRibbonPredecessorProof(preBuildRepairPredecessorDetail, preBuildRepairCatalogToken) = False Then Err.Raise vbObjectError + 6289, "BuildRepairHelpDocumentMacros", preBuildRepairPredecessorDetail
    LogTraceMessage TRACE_LEVEL_INFO, "Performance", "BuildRepairHelpDocumentMacros", "Build/Repair Help Document Macros predecessor proof completed.", "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(stageTimer)) & vbCrLf & NamingPerformanceSummary(), "Build/Repair Help Document Macros predecessor proof", 0, vbNullString, vbNullString, vbNullString, "CatalogValidation", "CodeDb", "Performance"
    InvalidateNamingSnapshot
    namingSnapshotOwned = False
    PulseSystemCommandStatus "Build/Repair Help Document Macros | reconciling Help managed definitions"
    If ReconcileHelpDocumentMacroDefinitions(catalogDetail, managedDefinitionMutationCount, managedDefinitionDirtyMask) = False Then Err.Raise vbObjectError + 6282, "BuildRepairHelpDocumentMacros", "Catalog Help macro reconciliation returned False."
    LogTraceMessage TRACE_LEVEL_INFO, "ManagedObjectCatalog", "BuildRepairHelpDocumentMacros", "Build/Repair Help Document Macros managed-definition mutation result captured.", "Mutations=" & CStr(managedDefinitionMutationCount) & "; DirtyDomainMask=" & CStr(managedDefinitionDirtyMask), "Build/Repair Help Document Macros B4 mutation result", 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID), "ManagedObjectDefinitionsMacros", "CodeDb", "Performance"
    ' INTENTIONAL OPERATION SNAPSHOT: B4 has completed the only Build/Repair Help Document Macros managed-name mutation; rebuild COMMON's authoritative naming snapshot once and retain it through the remaining
    ' non-naming-mutating stages.
    PulseSystemCommandStatus "Build/Repair Help Document Macros | preparing post-B4 naming snapshot"
    stageTimer = Timer
    BeginNamingSnapshot
    namingSnapshotOwned = True
    LogTraceMessage TRACE_LEVEL_INFO, "Performance", "BuildRepairHelpDocumentMacros", "Build/Repair Help Document Macros post-B4 naming snapshot prepared.", "Snapshot reflects the completed Help managed-definition reconciliation and remains valid through physical macros, D4, scoped G1, Ribbon, and Navigation Pane work because those stages do not own managed-name or naming-policy mutation." & vbCrLf & "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(stageTimer)), "Build/Repair Help Document Macros post-B4 naming snapshot", 0, vbNullString, vbNullString, vbNullString, "NamingSnapshot", "CodeDb", "Performance"
    PulseSystemCommandStatus "Build/Repair Help Document Macros | verifying physical Help macros"
    PhaseCVerifyPhysicalSystemHelpMacroState physicalDetail, HelpDocumentScopeAllActive, True
    PulseSystemCommandStatus "Build/Repair Help Document Macros | building changed Help macros"
    PhaseCBuildRequiredSystemHelpMacros physicalDetail
    PulseSystemCommandStatus "Build/Repair Help Document Macros | finalizing physical Help macros"
    PhaseCFinalizePhysicalSystemHelpMacroState physicalDetail
    ribbonCatalogProofToken = vbNullString
    ribbonCatalogProofTraceRunID = vbNullString
    ribbonCatalogProofDetail = vbNullString
    currentCatalogTokenForPredecessor = vbNullString
    If managedDefinitionMutationCount = 0 And managedDefinitionDirtyMask = 0 Then
        If Len(preBuildRepairCatalogToken) > 0 Then
            currentCatalogTokenForPredecessor = preBuildRepairCatalogToken
            ribbonCatalogProofDetail = "Complete catalog token reused from the exact successful pre-mutation predecessor proof because B4 made no logical catalog mutation and marked no catalog domain dirty."
        Else
            currentCatalogTokenForPredecessor = CompleteCatalogStateTokenForSetup(ribbonCatalogProofDetail)
        End If
    End If
    If Len(currentCatalogTokenForPredecessor) > 0 Then
        logicalCatalogDetail = "D4 exact-state fast path used. System Macro and Help Document Macro managed-definition reconciliation made zero logical catalog mutations and marked no catalog domains dirty, and the persisted complete-catalog proof remains exact/current. The current catalog token is retained for the same-run G1/H1 handoff; it is not mislabeled as a D4 same-run proof." & vbCrLf & ribbonCatalogProofDetail
        LogTraceMessage TRACE_LEVEL_INFO, "ManagedObjectCatalog", "BuildRepairHelpDocumentMacros", "Build/Repair Help Document Macros skipped full D4 because the authoritative complete-catalog proof is already current.", logicalCatalogDetail, "Build/Repair Help Document Macros D4 fast path", 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID), "CatalogValidation", "CodeDb", "Performance"
    Else
        PulseSystemCommandStatus "Build/Repair Help Document Macros | validating changed catalog state"
        If FinalizeCompleteCatalogStateForSetup(False, logicalCatalogDetail, currentCatalogTokenForPredecessor) = False Then Err.Raise vbObjectError + 6285, "BuildRepairHelpDocumentMacros", "D4 complete-catalog finalization returned False."
        If Len(currentCatalogTokenForPredecessor) = 0 Then Err.Raise vbObjectError + 6302, "BuildRepairHelpDocumentMacros", "D4 completed without returning the finalized complete-catalog token."
        d4RanThisOperation = True
        If CompleteCatalogSameRunProofForRibbon(ribbonCatalogProofToken, ribbonCatalogProofTraceRunID, ribbonCatalogProofDetail, currentCatalogTokenForPredecessor, activeTraceRunID) = False Then Err.Raise vbObjectError + 6286, "BuildRepairHelpDocumentMacros", "D4 completed but Catalog did not expose an authoritative same-run complete-catalog Ribbon handoff." & vbCrLf & ribbonCatalogProofDetail
        currentCatalogTokenForPredecessor = ribbonCatalogProofToken
    End If
    PulseSystemCommandStatus "Build/Repair Help Document Macros | applying and verifying Help macro Descriptions"
    If ApplyAndVerifyHelpDocumentManagedDescriptionsForBuildRepair(f1Detail, f1CatalogDescriptionMutated, f1RefreshedCatalogToken) = False Then Err.Raise vbObjectError + 6287, "BuildRepairHelpDocumentMacros", "Scoped G1 Help managed-description application returned False." & vbCrLf & f1Detail
    If f1CatalogDescriptionMutated Then
        If Len(f1RefreshedCatalogToken) = 0 Then Err.Raise vbObjectError + 6301, "BuildRepairHelpDocumentMacros", "Scoped Help G1 changed catalog descriptions without returning the refreshed complete-catalog token."
        currentCatalogTokenForPredecessor = f1RefreshedCatalogToken
    End If
    If HelpDocumentDescriptionSameRunG1ProofForRibbon(helpG1ProofToken, helpG1ProofTraceRunID, helpG1ProofDetail) = False Then Err.Raise vbObjectError + 6288, "BuildRepairHelpDocumentMacros", "Scoped G1 Help managed-description proof was not available for the complete predecessor proof." & vbCrLf & helpG1ProofDetail
    If Len(currentCatalogTokenForPredecessor) = 0 Then Err.Raise vbObjectError + 6290, "BuildRepairHelpDocumentMacros", "The current complete-catalog token was unavailable after scoped Help G1." & vbCrLf & ribbonCatalogProofDetail
    If d4RanThisOperation Then
        ribbonCatalogProofToken = vbNullString
        ribbonCatalogProofTraceRunID = vbNullString
        ribbonCatalogProofDetail = vbNullString
        If CompleteCatalogSameRunProofForRibbon(ribbonCatalogProofToken, ribbonCatalogProofTraceRunID, ribbonCatalogProofDetail, currentCatalogTokenForPredecessor, activeTraceRunID) = False Then Err.Raise vbObjectError + 6286, "BuildRepairHelpDocumentMacros", "D4/G1 completed but Catalog did not expose the authoritative same-run complete-catalog Ribbon handoff required by H1." & vbCrLf & ribbonCatalogProofDetail
    Else
        ribbonCatalogProofToken = vbNullString
        ribbonCatalogProofTraceRunID = vbNullString
    End If
    If FinalizeHelpDocumentMacroBuildRepairFullRibbonPredecessorProof(currentCatalogTokenForPredecessor, helpG1ProofToken, fullPredecessorProofToken, fullPredecessorProofTraceRunID, fullPredecessorProofDetail) = False Then Err.Raise vbObjectError + 6291, "BuildRepairHelpDocumentMacros", fullPredecessorProofDetail
    PulseSystemCommandStatus "Build/Repair Help Document Macros | updating Ribbon"
    BuildRibbon False, ribbonCatalogProofToken, ribbonCatalogProofTraceRunID, fullPredecessorProofToken, fullPredecessorProofTraceRunID
    ribbonDetail = "Ribbon reconciled through H using the same-run complete Build/Repair Help Document Macros predecessor proof" & IIf(Len(ribbonCatalogProofToken) > 0, " and same-run D4 proof.", ".") & " " & IIf(mRibbonRestartRequiredCurrentRun, "Database reopen required.", "No reopen required.")
    PulseSystemCommandStatus "Build/Repair Help Document Macros | updating Navigation Pane"
    If ReconcileHelpDocumentMacroNavigationPane(navDetail) = False Then Err.Raise vbObjectError + 6283, "BuildRepairHelpDocumentMacros", "Navigation Pane Help macro reconciliation returned False."
    PulseSystemCommandStatus "Build/Repair Help Document Macros | finalizing build/repair"
    detailText = preBuildRepairPredecessorDetail & vbCrLf & vbCrLf & catalogDetail & vbCrLf & vbCrLf & physicalDetail & vbCrLf & vbCrLf & logicalCatalogDetail & vbCrLf & vbCrLf & f1Detail & vbCrLf & vbCrLf & fullPredecessorProofDetail & vbCrLf & vbCrLf & ribbonDetail & vbCrLf & vbCrLf & navDetail
    If namingSnapshotOwned Then InvalidateNamingSnapshot
    namingSnapshotOwned = False
    BuildRepairHelpDocumentMacros = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If namingSnapshotOwned Then InvalidateNamingSnapshot
    namingSnapshotOwned = False
    On Error GoTo 0
    detailText = "Build/Repair Help Document Macros failed." & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
    Err.Raise savedErrNumber, "BuildRepairHelpDocumentMacros", detailText
End Function

Private Function DeleteAllObjectsWithStandaloneFallback(ByVal systemMacroKey As String, ByVal completeOwnedTraceAfterStandalone As Boolean) As Boolean
    'Macro 65 All-scope routing coordinator. Catalog owns the read-only eligibility decision before either destructive engine starts.
    On Error GoTo ErrHandler
    Dim catalogEligible As Boolean
    Dim approvedStandaloneFallback As Boolean
    Dim eligibilityReason As String
    Dim eligibilityDetail As String
    Dim deleteSucceeded As Boolean
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", _
                       "DeleteAllObjectsWithStandaloneFallback", _
                       "Macro 65 All-scope fallback eligibility evaluation started.", _
                       "CommandKey=" & systemMacroKey & "; Authority=SystemOwnedDeletionPathEligible", _
                       "Macro 65 deletion routing", 0, vbNullString, vbNullString, _
                       systemMacroKey, "SystemMacro", "CodeDb", "SystemMacro65"
    catalogEligible = SystemOwnedDeletionPathEligible(DELETE_SCOPE_ALL, approvedStandaloneFallback, eligibilityReason, eligibilityDetail)
    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", _
                       "DeleteAllObjectsWithStandaloneFallback", _
                       "Macro 65 Catalog eligibility evaluation completed.", _
                       "Eligible=" & CStr(catalogEligible) & "; ApprovedStandaloneFallback=" & CStr(approvedStandaloneFallback) & "; Reason=" & eligibilityReason & vbCrLf & eligibilityDetail, _
                       "Macro 65 deletion routing", 0, vbNullString, vbNullString, _
                       systemMacroKey, "SystemMacro", "CodeDb", "SystemMacro65"
    If catalogEligible Then
        LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", _
                           "DeleteAllObjectsWithStandaloneFallback", _
                           "Macro 65 selected the authoritative Catalog deletion route.", _
                           "Reason=" & eligibilityReason, "Macro 65 deletion routing", 0, vbNullString, vbNullString, _
                           systemMacroKey, "SystemMacro", "CodeDb", "SystemMacro65"
        deleteSucceeded = DeleteSystemOwnedObjects(DELETE_SCOPE_ALL, True, True)
        LogTraceMessage IIf(deleteSucceeded, TRACE_LEVEL_INFO, TRACE_LEVEL_ERROR), "DeleteObjects", "DeleteAllObjectsWithStandaloneFallback", "Catalog Delete All engine returned to the routing coordinator.", "Result=" & CStr(deleteSucceeded) & "; Cancelled=" & CStr(SystemCommandWasCancelled()) & "; DeferredRequired=" & CStr(CoreDeferredMacro65DeletionRequired()) & "; DeferredPending=" & CStr(CoreDeferredMacro65DeletionPending()), "Macro 65 deletion routing", 0, vbNullString, vbNullString, systemMacroKey, "SystemMacro", "CodeDb", "SystemMacro65"
        FlushTraceLogFile
        If SystemCommandWasCancelled() Then Exit Function
        DeleteAllObjectsWithStandaloneFallback = deleteSucceeded
        Exit Function
    End If
    If approvedStandaloneFallback = False Then
        Err.Raise vbObjectError + 6559, "DeleteAllObjectsWithStandaloneFallback", "Catalog deletion was not eligible and the Catalog authority did not approve standalone fallback." & vbCrLf & "Reason: " & eligibilityReason & vbCrLf & eligibilityDetail
    End If
    LogTraceMessage TRACE_LEVEL_WARNING, "DeleteObjects", _
                       "DeleteAllObjectsWithStandaloneFallback", _
                       "Macro 65 selected standalone name-based deletion before normal destructive deletion began.", _
                       "Reason=" & eligibilityReason & vbCrLf & eligibilityDetail & vbCrLf & "Normal catalog deletion invoked=False", _
                       "Macro 65 standalone preflight fallback", 0, vbNullString, vbNullString, _
                       systemMacroKey, "SystemMacro", "CodeDb", "SystemMacro65"
    If ConfirmDBObjectsByNameDeletion(True) = False Then
        MarkSystemCommandCancelled
        LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "DeleteAllObjectsWithStandaloneFallback", _
                           "Macro 65 standalone fallback was cancelled before destructive work.", _
                           "Reason=" & eligibilityReason & "; Normal catalog deletion invoked=False", _
                           "Macro 65 standalone preflight fallback", 0, vbNullString, vbNullString, _
                           systemMacroKey, "SystemMacro", "CodeDb", "SystemMacro65"
        Exit Function
    End If
    deleteSucceeded = DeleteDBObjectsByName(False, True, completeOwnedTraceAfterStandalone, "System command complete: " & systemMacroKey, "System command failed: " & systemMacroKey)
    LogTraceMessage IIf(deleteSucceeded, TRACE_LEVEL_INFO, TRACE_LEVEL_WARNING), "DeleteObjects", "DeleteAllObjectsWithStandaloneFallback", "Standalone Delete All engine returned to the routing coordinator.", "Result=" & CStr(deleteSucceeded) & "; Cancelled=" & CStr(SystemCommandWasCancelled()) & "; DeferredRequired=" & CStr(CoreDeferredMacro65DeletionRequired()) & "; DeferredPending=" & CStr(CoreDeferredMacro65DeletionPending()), "Macro 65 deletion routing", 0, vbNullString, vbNullString, systemMacroKey, "SystemMacro", "CodeDb", "SystemMacro65"
    FlushTraceLogFile
    DeleteAllObjectsWithStandaloneFallback = deleteSucceeded
    Exit Function
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If SystemCommandWasCancelled() Or IsCancellationError(savedErrNumber) Then
        If SystemCommandWasCancelled() = False Then MarkSystemCommandCancelled
        Exit Function
    End If
    LogTraceMessage TRACE_LEVEL_ERROR, "DeleteObjects", _
                       "DeleteAllObjectsWithStandaloneFallback", _
                       "Macro 65 All-scope deletion routing or selected deletion engine failed.", _
                       "CatalogEligible=" & CStr(catalogEligible) & "; ApprovedStandaloneFallback=" & CStr(approvedStandaloneFallback) & "; Reason=" & eligibilityReason, _
                       "Macro 65 deletion failure", savedErrNumber, savedErrSource, savedErrDescription, _
                       systemMacroKey, "SystemMacro", "CodeDb", "SystemMacro65"
    Err.Raise savedErrNumber, "DeleteAllObjectsWithStandaloneFallback", savedErrDescription
End Function

Private Sub BuildRepairAllTablesAndQueries(Optional ByVal showMessage As Boolean = True)
    ' Operational System Macro #71. Delegates table/query families to their established owners; this coordinator contains no duplicate schema builders.
    ' Standalone completion also reconciles authoritative Descriptions only for this command's managed table/query inventory. Numbered Setup continues to defer that metadata work to G1.
    On Error GoTo ErrHandler
    Dim details As String
    Dim descriptionDetail As String
    Dim definitionData As Variant
    Dim objectName As String
    Dim objectTypeCode As String
    Dim objectDescriptionChanged As Variant
    Dim fieldDescriptionChangeCount As Variant
    Dim objectDescriptionChanges As Long
    Dim fieldDescriptionChanges As Long
    Dim managedRepairInventory As Collection
    Dim Missing As New Collection
    Set managedRepairInventory = ManagedTableAndQueryRepairInventory()
    'This coordinator owns the final complete managed table/query Description pass, so suppress the narrower nested TABLES standalone reconciliation to avoid duplicate physical scans.
    If RepairTables(False, details, False) = False Then Err.Raise vbObjectError + 6281, "BuildRepairAllTablesAndQueries", "TABLES-owned table/query repair returned False." & vbCrLf & details
    RepairReferenceInventoryFoundationForSetup False
    If RepairCatalogTableAndQueryFamilies(False, details) = False Then Err.Raise vbObjectError + 6282, "BuildRepairAllTablesAndQueries", details
    'This coordinator performs the final scoped Description reconciliation once, so the nested Module Guard structural repair must not run its standalone Description pass here.
    RepairModuleUpdateGuard False, False
    For Each definitionData In managedRepairInventory
        objectName = CStr(definitionData(0))
        objectTypeCode = CStr(definitionData(1))
        If StrComp(objectTypeCode, "Table", vbTextCompare) = 0 Then
            If SharedDao_AccessObjectExists(currentDb, acTable, objectName) = False Then Missing.Add objectTypeCode & ": " & objectName
        ElseIf StrComp(objectTypeCode, "Query", vbTextCompare) = 0 Then
            If SharedDao_AccessObjectExists(currentDb, acQuery, objectName) = False Then Missing.Add objectTypeCode & ": " & objectName
        End If
    Next definitionData
    If Missing.count > 0 Then Err.Raise vbObjectError + 6288, "BuildRepairAllTablesAndQueries", "The complete authoritative ADC table/query inventory was not present after repair." & vbCrLf & SharedCollection_Join(Missing, vbCrLf)
    For Each definitionData In managedRepairInventory
        objectName = CStr(definitionData(0))
        objectTypeCode = CStr(definitionData(1))
        objectDescriptionChanged = False
        If ApplyAndVerifyManagedObjectDescriptionForStandaloneRepair(objectName, objectTypeCode, objectDescriptionChanged) = False Then Err.Raise vbObjectError + 10115, "BuildRepairAllTablesAndQueries", "Managed-object Description reconciliation returned False." & vbCrLf & objectTypeCode & ": " & objectName
        If CBool(objectDescriptionChanged) Then objectDescriptionChanges = objectDescriptionChanges + 1
        If StrComp(objectTypeCode, "Table", vbTextCompare) = 0 Then
            fieldDescriptionChangeCount = CLng(0)
            descriptionDetail = vbNullString
            If ApplyAndVerifyManagedTableFieldDescriptionsForStandaloneRepair(objectName, fieldDescriptionChangeCount, descriptionDetail) = False Then Err.Raise vbObjectError + 10116, "BuildRepairAllTablesAndQueries", descriptionDetail
            fieldDescriptionChanges = fieldDescriptionChanges + CLng(Nz(fieldDescriptionChangeCount, 0))
        End If
    Next definitionData
    If showMessage Then ReportCoreActionSuccess "BuildRepairAllTablesAndQueries", "Tables", "All ADC-managed tables and queries are current.", "All authoritative table/query repair owners completed, the complete managed inventory was verified, and standalone managed descriptions were reconciled." & vbCrLf & "Object descriptions changed: " & CStr(objectDescriptionChanges) & vbCrLf & "Field descriptions changed: " & CStr(fieldDescriptionChanges), "Build/repair all ADC tables and queries", vbNullString, "TableQuerySet", "CodeDb", "Repair", True, "Build/Repair All ADC Tables and Queries", vbInformation, True, "Tables", False
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "BuildRepairAllTablesAndQueries", "Could not build/repair the complete ADC-managed table/query inventory." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Function RunOperationalSystemCommand(ByVal systemMacroKey As String) As Boolean
    'Intended use: Returns the Run Access Development Console System Command result for direct, query, form, or required cross-module callers.
    On Error GoTo ErrHandler
    Dim keyText As String
    Dim traceRunOwnedByProcedure As Boolean
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim terminalDestructiveCommand As Boolean
    Dim deletionScope As ObjectDeletionScope
    Dim macro65SelectorExistedInitially As Boolean
    Dim macro65PreparationShowSeconds As Double
    Dim macro65UserInteractionSeconds As Double
    Dim macro65SelectedScopeText As String
    Dim scanManagedDefinitionSnapshotOwned As Boolean
    Dim scanSystemMacroDefinitionSnapshotOwned As Boolean
    Dim scanNamingSnapshotOwned As Boolean
    Dim scanSnapshotTimer As Single
    Dim modalProgressControlOwned As Boolean
    Dim deletionOperationSucceeded As Boolean
    Dim deferredMacro65CompletionKind As String
    keyText = UCase$(SystemMacroCommandKeyFromRunArgument(systemMacroKey))
    mBuildRepairHelpDocumentMacrosMacroResultDetail = vbNullString
    If mOperationalSystemCommandActive Then
        RunOperationalSystemCommand = False
        Exit Function
    End If
    mOperationalSystemCommandActive = True
    terminalDestructiveCommand = (StrComp(keyText, UCase$(SYS_MACRO_KEY_DELETE_ALL_OBJECTS), vbBinaryCompare) = 0)
    'This is the paired terminal implementation for the early RunSystemCommand branch. Do not require a SystemMacroDefinitionByKey lookup before reaching this block; Macro 65 must remain able to start its
    'scope/progress UI without rebuilding the full definition set.
    If terminalDestructiveCommand Then
        ClearSystemCommandResult
        If TraceRunActive() = False Then
            StartTraceRun "SystemCommand"
            traceRunOwnedByProcedure = True
        End If
        LogAccessProcessPriorityForCurrentTrace "RunOperationalSystemCommand.Macro65"
        LogTraceMessage TRACE_LEVEL_INFO, "Performance", "RunOperationalSystemCommand", "Macro 65 dedicated early-dispatch path entered after trace initialization.", "Dedicated early-dispatch path used=True; full SystemMacro definition lookup bypassed=True; one operation trace active=True", "Macro 65 operational entry", 0, vbNullString, vbNullString, keyText, "SystemMacro", "CodeDb", "SystemMacro65"
        deletionScope = PromptObjectDeletionScope(macro65SelectorExistedInitially, macro65PreparationShowSeconds, macro65UserInteractionSeconds, macro65SelectedScopeText)
        LogTraceMessage TRACE_LEVEL_INFO, "Performance", "RunOperationalSystemCommand", "Macro 65 launch summary.", "Dedicated early-dispatch path used=Yes; selector existed initially=" & IIf(macro65SelectorExistedInitially, "Yes", "No") & "; currentness check used=No; selector recreation used=No; preparation/show elapsed=" & FormattedElapsedSeconds(macro65PreparationShowSeconds) & "; user-interaction elapsed=" & FormattedElapsedSeconds(macro65UserInteractionSeconds) & "; selected scope=" & macro65SelectedScopeText & "; canceled=" & IIf(deletionScope = DELETE_SCOPE_CANCEL, "Yes", "No"), "Macro 65 launch summary", 0, vbNullString, vbNullString, keyText, "SystemMacro", "CodeDb", "SystemMacro65"
        If deletionScope = DELETE_SCOPE_CANCEL Then
            LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "RunOperationalSystemCommand", "DELETE ALL ADC OBJECTS CANCELLED", "Scope=Cancel; Explicit selector callback=True; Dispatch skipped=True", "Macro 65 deletion cancellation", 0, vbNullString, vbNullString, keyText, "SystemMacro", "CodeDb", "SystemMacro65"
        Else
            LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "RunOperationalSystemCommand", "DELETE ALL ADC OBJECTS START", "Scope=" & macro65SelectedScopeText & "; CallingCommand=System Macro 65; CommandKey=" & keyText, "Macro 65 deletion operation", 0, vbNullString, vbNullString, keyText, "SystemMacro", "CodeDb", "SystemMacro65"
            LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "RunOperationalSystemCommand", "DELETE DISPATCH", "Scope=" & macro65SelectedScopeText, "Macro 65 deletion dispatcher", 0, vbNullString, vbNullString, keyText, "SystemMacro", "CodeDb", "SystemMacro65"
        End If
        Select Case deletionScope
            Case DELETE_SCOPE_CANCEL
                MarkSystemCommandCancelled
                RunOperationalSystemCommand = False
            Case DELETE_SCOPE_GENERATED_FORMS
                deletionOperationSucceeded = False
                DeleteGeneratedForms True, True, operationSucceededResult:=deletionOperationSucceeded
                RunOperationalSystemCommand = deletionOperationSucceeded And (SystemCommandWasCancelled() = False)
            Case DELETE_SCOPE_MACROS
                RunOperationalSystemCommand = DeleteSystemHelpMacros(True, True)
            Case DELETE_SCOPE_TABLES, DELETE_SCOPE_QUERIES, DELETE_SCOPE_REPORTS
                RunOperationalSystemCommand = DeleteSystemOwnedObjects(deletionScope, True, True)
            Case DELETE_SCOPE_ALL
                RunOperationalSystemCommand = DeleteAllObjectsWithStandaloneFallback(keyText, traceRunOwnedByProcedure)
            Case Else
                Err.Raise vbObjectError + 6294, "RunOperationalSystemCommand", "System Macro 65 returned an unsupported deletion scope: " & CStr(CLng(deletionScope))
        End Select
        LogTraceMessage IIf(RunOperationalSystemCommand, TRACE_LEVEL_INFO, TRACE_LEVEL_WARNING), "DeleteObjects", "RunOperationalSystemCommand", "Delete engine returned to the Macro 65 operational coordinator.", "Scope=" & macro65SelectedScopeText & "; Result=" & CStr(RunOperationalSystemCommand) & "; Cancelled=" & CStr(SystemCommandWasCancelled()) & "; DeferredRequired=" & CStr(CoreDeferredMacro65DeletionRequired()) & "; DeferredPending=" & CStr(CoreDeferredMacro65DeletionPending()), "Macro 65 terminal handoff", 0, vbNullString, vbNullString, keyText, "SystemMacro", "CodeDb", "SystemMacro65"
        FlushTraceLogFile
        If CoreDeferredMacro65DeletionRequired() And CoreDeferredMacro65DeletionPending() = False Then
            deferredMacro65CompletionKind = CoreDeferredMacro65DeletionCompletionKind()
            If SystemCommandWasCancelled() Then
                CoreCancelDeferredMacro65DeletionRequest
            ElseIf RunOperationalSystemCommand Or StrComp(deferredMacro65CompletionKind, "Standalone", vbTextCompare) = 0 Then
                LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "RunOperationalSystemCommand", "Deferred Macro 65 scheduling call starting.", "CompletionOwner=" & deferredMacro65CompletionKind & "; TraceOwnedByMacro65=" & CStr(traceRunOwnedByProcedure), "Macro 65 terminal handoff", 0, vbNullString, vbNullString, keyText, "SystemMacro", "CodeDb", "SystemMacro65"
                FlushTraceLogFile
                CoreScheduleDeferredMacro65Deletion traceRunOwnedByProcedure, "System command complete: " & keyText, "System command failed: " & keyText
                LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "RunOperationalSystemCommand", "Deferred Macro 65 scheduling call returned.", "CompletionOwner=" & deferredMacro65CompletionKind & "; DeferredPending=" & CStr(CoreDeferredMacro65DeletionPending()), "Macro 65 terminal handoff", 0, vbNullString, vbNullString, keyText, "SystemMacro", "CodeDb", "SystemMacro65"
                FlushTraceLogFile
            Else
                'A normal managed deletion that did not complete must not continue by deleting its invoking macro afterward.
                CoreCancelDeferredMacro65DeletionRequest
            End If
        End If
        If CoreDeferredMacro65DeletionPending() Then
            'Core owns the post-RunCode timing for the one invoking Access macro. The deletion owner has completed all ordinary work; final result/trace completion waits for the timer callback.
            'This handoff trace is best-effort because the deferred terminal deletion is already scheduled; no diagnostic write may convert that committed handoff into a second synchronous failure path.
            On Error Resume Next
            LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "RunOperationalSystemCommand", "Macro 65 synchronous RunCode work completed; terminal invoking-macro deletion and final trace completion remain pending.", "Scope=" & macro65SelectedScopeText & "; CompletionOwner=" & CoreDeferredMacro65DeletionCompletionKind() & "; TraceOwnedByMacro65=" & CStr(traceRunOwnedByProcedure), "Macro 65 terminal handoff", 0, vbNullString, vbNullString, keyText, "SystemMacro", "CodeDb", "SystemMacro65"
            FlushTraceLogFile
            Err.Clear
            On Error GoTo ErrHandler
        ElseIf SystemCommandWasCancelled() Then
            If traceRunOwnedByProcedure Then CompleteTraceRun "System command cancelled: " & keyText, False
        ElseIf RunOperationalSystemCommand = False Then
            LogTraceMessage TRACE_LEVEL_ERROR, "DeleteObjects", "RunOperationalSystemCommand", "DELETE ALL ADC OBJECTS FAILED", "Scope=" & macro65SelectedScopeText & "; operation result=False; cancelled=False", "Macro 65 deletion terminal result", 0, vbNullString, vbNullString, keyText, "SystemMacro", "CodeDb", "SystemMacro65"
            If traceRunOwnedByProcedure Then CompleteTraceRun "System command failed: " & keyText, False
        ElseIf traceRunOwnedByProcedure Then
            CompleteTraceRun "System command complete: " & keyText
        End If
        LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "RunOperationalSystemCommand", "Macro 65 operational coordinator returning to RunSystemCommand.", "Scope=" & macro65SelectedScopeText & "; Result=" & CStr(RunOperationalSystemCommand) & "; DeferredPending=" & CStr(CoreDeferredMacro65DeletionPending()), "Macro 65 terminal handoff", 0, vbNullString, vbNullString, keyText, "SystemMacro", "CodeDb", "SystemMacro65"
        FlushTraceLogFile
        mOperationalSystemCommandActive = False
        Exit Function
    End If
    RecoverUIBeforeSetup keyText
    ClearSystemCommandResult
    BeginTrackedProcess "System command: " & keyText
    If TraceRunActive() = False Then
        StartTraceRun "SystemCommand"
        traceRunOwnedByProcedure = True
    End If
    LogAccessProcessPriorityForCurrentTrace "RunOperationalSystemCommand"
    If SystemCommandIsScanCommand(keyText) And mSetupControlActive = False Then
        BeginScanProgressControl "Scan is starting: " & keyText
        modalProgressControlOwned = True
    ElseIf StrComp(keyText, UCase$(SYS_MACRO_KEY_BUILD_REPAIR_HELP_DOCUMENT_MACROS), vbBinaryCompare) = 0 And mSetupControlActive = False Then
        BeginScanProgressControl "Preparing operation snapshots...", "Rebuilding Help Document Macros", DEFAULT_UTILITY_SHORT_NAME & " Help Document macros are being rebuilt.", False
        modalProgressControlOwned = True
    ElseIf StrComp(keyText, UCase$(SYS_MACRO_KEY_RUN_HEALTH_REVIEW), vbBinaryCompare) = 0 And mSetupControlActive = False Then
        BeginScanProgressControl "Preparing Health Review...", "Health Review", DEFAULT_UTILITY_SHORT_NAME & " health review is running.", False
        modalProgressControlOwned = True
    ElseIf StrComp(keyText, UCase$(SYS_MACRO_KEY_GENERATE_SELECTED_RUNNER_MACROS), vbBinaryCompare) = 0 And mSetupControlActive = False Then
        BeginScanProgressControl "Preparing macro generation...", "Generate Selected Runner Macros", DEFAULT_UTILITY_SHORT_NAME & " selected Runner Macros are being generated.", False
        modalProgressControlOwned = True
        SetMacroGenerationCompletionDeferral True
    ElseIf StrComp(keyText, UCase$(SYS_MACRO_KEY_GENERATE_ALL_CONFIGURED_RUNNER_MACROS), vbBinaryCompare) = 0 And mSetupControlActive = False Then
        BeginScanProgressControl "Preparing macro generation...", "Generate All Configured Runner Macros", DEFAULT_UTILITY_SHORT_NAME & " configured Runner Macros are being generated.", False
        modalProgressControlOwned = True
        SetMacroGenerationCompletionDeferral True
    End If
    LogTraceMessage "Info", "SystemCommand", "RunOperationalSystemCommand", "System command started.", "Requested key: " & systemMacroKey & vbCrLf & "Resolved key: " & keyText, keyText, 0, vbNullString, vbNullString, keyText, "SystemMacro", "CodeDb", "SystemCommand"
    PulseSystemCommandStatus SetupSystemCommandStatusText(keyText, "System command started: " & keyText, False)
    PulseSystemCommandStatus SetupSystemCommandStatusText(keyText, "Running system command: " & keyText, False)
    'The keyed System Macro command owns scan-operation snapshots through prerequisite work, nested scan phases, and finalization. Child scans must not end them.
    Select Case keyText
        Case UCase$(SYS_MACRO_KEY_SCAN_ALL_DB_INVENTORY)
            scanSnapshotTimer = Timer
            PulseSystemCommandStatus "Scan All DB Inventory | preparing operation snapshots"
            BeginSystemMacroDefinitionSnapshot
            scanSystemMacroDefinitionSnapshotOwned = True
            BeginManagedObjectDefinitionSnapshot
            scanManagedDefinitionSnapshotOwned = True
            BeginNamingSnapshot
            scanNamingSnapshotOwned = True
            PumpSetupResponsiveness True
            LogTraceMessage "Info", "Scan", "RunOperationalSystemCommand", "Scan All DB Inventory operation snapshots prepared.", "System Macro definition snapshot: armed" & vbCrLf & "Managed-object definition snapshot: armed" & vbCrLf & "Naming snapshot: built once for the scan operation" & vbCrLf & "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(scanSnapshotTimer)), keyText, 0, vbNullString, vbNullString, keyText, "SystemMacro", "CodeDb", "SystemCommand"
        Case UCase$(SYS_MACRO_KEY_SCAN_RUNNER_MACRO_CANDIDATES)
            scanSnapshotTimer = Timer
            PulseSystemCommandStatus "Candidate scan | preparing operation snapshots"
            BeginSystemMacroDefinitionSnapshot
            scanSystemMacroDefinitionSnapshotOwned = True
            BeginNamingSnapshot
            scanNamingSnapshotOwned = True
            PumpSetupResponsiveness True
            LogTraceMessage "Info", "Scan", "RunOperationalSystemCommand", "Candidate scan operation snapshots prepared.", "System Macro definition snapshot: armed" & vbCrLf & "Naming snapshot: built once for the scan operation" & vbCrLf & "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(scanSnapshotTimer)), keyText, 0, vbNullString, vbNullString, keyText, "SystemMacro", "CodeDb", "SystemCommand"
        Case UCase$(SYS_MACRO_KEY_SCAN_DB_INVENTORY_ACTIVE_REFERENCES), UCase$(SYS_MACRO_KEY_SCAN_DB_INVENTORY_AVAILABLE_REFERENCES), UCase$(SYS_MACRO_KEY_SCAN_DB_INVENTORY_LOADED_VBA_PROJECTS), UCase$(SYS_MACRO_KEY_SCAN_DB_INVENTORY_ADD_INS_AND_LOADED_VBA_PROJECTS), UCase$(SYS_MACRO_KEY_SCAN_DB_INVENTORY_LOADED_VBA_SOURCE)
            scanSnapshotTimer = Timer
            PulseSystemCommandStatus "Reference/Add-in scan | preparing operation snapshots | command=" & keyText
            BeginManagedObjectDefinitionSnapshot
            scanManagedDefinitionSnapshotOwned = True
            BeginNamingSnapshot
            scanNamingSnapshotOwned = True
            PumpSetupResponsiveness True
            LogTraceMessage "Info", "Scan", "RunOperationalSystemCommand", "Reference/Add-in scan operation snapshots prepared.", "Command: " & keyText & vbCrLf & "Managed-object definition snapshot: armed" & vbCrLf & "Naming snapshot: built once for the scan operation" & vbCrLf & "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(scanSnapshotTimer)), keyText, 0, vbNullString, vbNullString, keyText, "SystemMacro", "CodeDb", "SystemCommand"
    End Select
    Select Case keyText
        Case UCase$(SYS_MACRO_KEY_APPLY_OBJECT_NAMING_RULES)
            ApplyObjectNamingRules False
        Case UCase$(SYS_MACRO_KEY_REGISTER_APPROVED_BUILD)
            RegisterApprovedBuild False
        Case UCase$(SYS_MACRO_KEY_RESOLVE_PENDING_SETUP_RECOVERY)
            ResolvePendingSetupRecovery False
        Case UCase$(SYS_MACRO_KEY_GENERATE_SELECTED_RUNNER_MACROS)
            If CodeDbIsHostDb() = False Then Err.Raise vbObjectError + 5453, "RunOperationalSystemCommand.GenerateSelectedRunnerMacros", "CodeDb macro generation requires ConfigDb=CodeDb and GenerationTargetDb=CodeDb."
            GenerateSelectedRunnerMacrosForActiveRoute "RunOperationalSystemCommand.GenerateSelectedRunnerMacros", "Generate Selected Runner Macros"
        Case UCase$(SYS_MACRO_KEY_GENERATE_ALL_CONFIGURED_RUNNER_MACROS)
            If CodeDbIsHostDb() = False Then Err.Raise vbObjectError + 5454, "RunOperationalSystemCommand.GenerateAllConfiguredRunnerMacros", "CodeDb macro generation requires ConfigDb=CodeDb and GenerationTargetDb=CodeDb."
            GenerateAllConfiguredRunnerMacrosForActiveRoute "RunOperationalSystemCommand.GenerateAllConfiguredRunnerMacros", "Generate All Configured Runner Macros"
        Case UCase$(SYS_MACRO_KEY_GENERATE_PROMPTED_RUNNER_MACROS)
            If CodeDbIsHostDb() = False Then Err.Raise vbObjectError + 5457, "RunOperationalSystemCommand.GeneratePromptedRunnerMacros", "Prompted Runner Macro generation requires ConfigDb=CodeDb and GenerationTargetDb=CodeDb."
            GeneratePromptedRunnerMacrosForActiveRoute "RunOperationalSystemCommand.GeneratePromptedRunnerMacros", "Generate Prompted Runner Macros", "Generate Prompted Runner Macros"
        Case UCase$(SYS_MACRO_KEY_SCAN_ALL_DB_INVENTORY)
            ScanAllDBInventory
        Case UCase$(SYS_MACRO_KEY_SCAN_RUNNER_MACRO_CANDIDATES)
            ScanCandidatesForRouting True, False, codeDb, CONFIG_OWNER_DATABASE_ROLE_CODE_DB, "Find CodeDb " & DEFAULT_UTILITY_SHORT_NAME & " candidates", False
        Case UCase$(SYS_MACRO_KEY_ACCEPT_SELECTED_RUNNER_MACRO_CANDIDATES)
            AcceptSelectedRunnerMacroCandidates False
        Case UCase$(SYS_MACRO_KEY_RESOLVE_RUNNER_MACRO_TARGETS)
            ResolveRunnerMacroTargets False
        Case UCase$(SYS_MACRO_KEY_VALIDATE_RUNNER_MACRO_TARGETS)
            ValidateRunnerMacroTargets
        Case UCase$(SYS_MACRO_KEY_RUN_HEALTH_REVIEW)
            RunHealthReview False
        Case UCase$(SYS_MACRO_KEY_FIND_BROKEN_FORM_AND_REPORT_REFERENCES)
            RequireSameDatabaseUICommand "RunOperationalSystemCommand.FindBrokenFormAndReportReferences", "Finding broken non-generated Form and Report references"
            FindBrokenFormAndReportReferences False
        Case UCase$(SYS_MACRO_KEY_REVIEW_STALE_CANDIDATES)
            ReviewStaleCandidates False, True
        Case UCase$(SYS_MACRO_KEY_MARK_STALE_CANDIDATES)
            MarkStaleCandidates False, True
        Case UCase$(SYS_MACRO_KEY_DELETE_SELECTED_DUPLICATE_ROWS)
            DeleteSelectedDuplicateRows False
        Case UCase$(SYS_MACRO_KEY_SCAN_DB_INVENTORY_ACTIVE_REFERENCES)
            ScanDBInventoryActiveReferences False
        Case UCase$(SYS_MACRO_KEY_SCAN_DB_INVENTORY_AVAILABLE_REFERENCES)
            ScanDBInventoryAvailableReferences False
        Case UCase$(SYS_MACRO_KEY_SCAN_DB_INVENTORY_LOADED_VBA_PROJECTS)
            ScanDBInventoryLoadedVBAProjects False
        Case UCase$(SYS_MACRO_KEY_SCAN_DB_INVENTORY_ADD_INS_AND_LOADED_VBA_PROJECTS)
            ScanDBInventoryAddInsAndLoadedVBAProjects False
        Case UCase$(SYS_MACRO_KEY_SCAN_DB_INVENTORY_LOADED_VBA_SOURCE)
            ScanDBInventoryLoadedVBASource False
        Case UCase$(SYS_MACRO_KEY_BUILD_REPAIR_CUSTOM_MACRO_LAUNCHERS)
            BuildRepairCustomMacroLaunchers False
        Case UCase$(SYS_MACRO_KEY_BUILD_UPDATE_CUSTOM_MACRO_LAUNCHER_BUTTONS): BuildUpdateCustomMacroLauncherButtons False
        Case UCase$(SYS_MACRO_KEY_REVIEW_CUSTOM_MACRO_LAUNCHER_BUTTONS)
            ReviewCustomMacroLauncherButtons False
        Case UCase$(SYS_MACRO_KEY_CLEAN_STALE_CUSTOM_MACRO_LAUNCHER_BUTTONS)
            CleanStaleCustomMacroLauncherButtons False
        Case UCase$(SYS_MACRO_KEY_ENABLE_CUSTOM_MACRO_LAUNCHER_AUTOMATION)
            EnableCustomMacroLauncherAutomation False
        Case UCase$(SYS_MACRO_KEY_DISABLE_CUSTOM_MACRO_LAUNCHER_AUTOMATION)
            DisableCustomMacroLauncherAutomation False
        Case UCase$(SYS_MACRO_KEY_BUILD_CHATGPT_UI_FORMS)
            RequireSameDatabaseUICommand "RunOperationalSystemCommand.BuildChatGPTUIForms", "Building ChatGPT UI Forms"
            BuildGeneratedForms True, False, buildScope:="ChatGPTUIForms"
        Case UCase$(SYS_MACRO_KEY_DELETE_CHATGPT_UI_FORMS)
            RequireSameDatabaseUICommand "RunOperationalSystemCommand.DeleteChatGPTUIForms", "Deleting ChatGPT UI Forms"
            deletionOperationSucceeded = False
            DeleteGeneratedForms False, True, operationSucceededResult:=deletionOperationSucceeded, buildScope:="ChatGPTUIForms"
            If deletionOperationSucceeded = False And SystemCommandWasCancelled() = False Then Err.Raise vbObjectError + 6285, "RunOperationalSystemCommand.DeleteChatGPTUIForms", "ChatGPT UI Form deletion did not complete successfully."
        Case UCase$(SYS_MACRO_KEY_BUILD_RIBBON)
            RequireSameDatabaseUICommand "RunOperationalSystemCommand.BuildRibbon", "Building the ADC Ribbon"
            BuildRibbon False
        Case UCase$(SYS_MACRO_KEY_DISABLE_RIBBON)
            RequireSameDatabaseUICommand "RunOperationalSystemCommand.DisableRibbon", "Disabling the ADC Ribbon"
            DisableRibbon False
        Case UCase$(SYS_MACRO_KEY_ADD_TABS_TO_ACCESS_RIBBON)
            RequireSameDatabaseUICommand "RunOperationalSystemCommand.AddTabsToAccessRibbon", "Adding the long-tooltip Access Development Console tabs to the Access Ribbon"
            AddTabsToAccessRibbon False
        Case UCase$(SYS_MACRO_KEY_REPAIR_MANAGED_OBJECT_CATALOG)
            RepairManagedObjectCatalog False
        Case UCase$(SYS_MACRO_KEY_VALIDATE_MANAGED_OBJECT_CATALOG)
            If ValidateManagedObjectCatalog(False, True) = False Then Err.Raise vbObjectError + 6272, "RunOperationalSystemCommand.ValidateManagedObjectCatalog", "Managed Object Catalog validation failed."
        Case UCase$(SYS_MACRO_KEY_BACKUP_ALL_USER_DATA)
            BackupAllUserData False
        Case UCase$(SYS_MACRO_KEY_BUILD_REPAIR_ALL_TABLES_AND_QUERIES)
            ClearTableRepairWarningResult
            BuildRepairAllTablesAndQueries False
            AddTableRepairWarningToParent
        Case UCase$(SYS_MACRO_KEY_RESET_ALL_USER_DATA)
            ResetAllUserData False, True
        Case UCase$(SYS_MACRO_KEY_BUILD_REPAIR_HELP_DOCUMENT_MACROS)
            If BuildRepairHelpDocumentMacros(mBuildRepairHelpDocumentMacrosMacroResultDetail) = False Then Err.Raise vbObjectError + 6284, "RunOperationalSystemCommand.BuildRepairHelpDocumentMacros", "Help Document macro synchronization returned False."
        Case UCase$(SYS_MACRO_KEY_CLEAR_TRACE_LOG_FILE)
            ClearTraceLogFile False, True
        Case Else
            Err.Raise vbObjectError + 6270, "RunOperationalSystemCommand", "Unknown Access Development Console system command key: " & systemMacroKey
    End Select
    If SystemCommandWasCancelled() Then
        If modalProgressControlOwned Then EndSetupControl
        modalProgressControlOwned = False
        If StrComp(keyText, UCase$(SYS_MACRO_KEY_GENERATE_SELECTED_RUNNER_MACROS), vbBinaryCompare) = 0 Or StrComp(keyText, UCase$(SYS_MACRO_KEY_GENERATE_ALL_CONFIGURED_RUNNER_MACROS), vbBinaryCompare) = 0 Then SetMacroGenerationCompletionDeferral False
        ReportSystemCommandCancelled systemMacroKey, keyText
        ClearSystemCommandResult
        On Error Resume Next
        If scanNamingSnapshotOwned Then InvalidateNamingSnapshot
        If scanSystemMacroDefinitionSnapshotOwned Then EndSystemMacroDefinitionSnapshot
        If scanManagedDefinitionSnapshotOwned Then EndManagedObjectDefinitionSnapshot
        On Error GoTo 0
        EndTrackedProcess
        If traceRunOwnedByProcedure Then CompleteTraceRun "System command cancelled: " & keyText, False
        RunOperationalSystemCommand = False
        mOperationalSystemCommandActive = False
        Exit Function
    End If
    If (StrComp(keyText, UCase$(SYS_MACRO_KEY_GENERATE_SELECTED_RUNNER_MACROS), vbBinaryCompare) = 0 Or StrComp(keyText, UCase$(SYS_MACRO_KEY_GENERATE_ALL_CONFIGURED_RUNNER_MACROS), vbBinaryCompare) = 0) And modalProgressControlOwned Then
        EndSetupControl
        modalProgressControlOwned = False
        SetMacroGenerationCompletionDeferral False, True
    End If
    PulseSystemCommandStatus SetupSystemCommandStatusText(keyText, "System command complete: " & keyText, True)
    If modalProgressControlOwned Then EndSetupControl
    modalProgressControlOwned = False
    LogTraceMessage "Info", "SystemCommand", "RunOperationalSystemCommand", "System command completed.", "Requested key: " & systemMacroKey & vbCrLf & "Resolved key: " & keyText, keyText, 0, vbNullString, vbNullString, keyText, "SystemMacro", "CodeDb", "SystemCommand"
    If SystemCommandShowsSuccessPopup(keyText) Then ReportSystemCommandSuccess systemMacroKey, keyText
    ClearSystemCommandResult
    On Error Resume Next
    If scanNamingSnapshotOwned Then InvalidateNamingSnapshot
    If scanSystemMacroDefinitionSnapshotOwned Then EndSystemMacroDefinitionSnapshot
    If scanManagedDefinitionSnapshotOwned Then EndManagedObjectDefinitionSnapshot
    On Error GoTo 0
    EndTrackedProcess
    If traceRunOwnedByProcedure Then CompleteTraceRun "System command complete: " & keyText
    RunOperationalSystemCommand = True
    mOperationalSystemCommandActive = False
    Exit Function
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If SystemCommandWasCancelled() Or IsCancellationError(savedErrNumber) Then
        If SystemCommandWasCancelled() = False Then MarkSystemCommandCancelled
        If modalProgressControlOwned Then EndSetupControl
        modalProgressControlOwned = False
        If StrComp(keyText, UCase$(SYS_MACRO_KEY_GENERATE_SELECTED_RUNNER_MACROS), vbBinaryCompare) = 0 Or StrComp(keyText, UCase$(SYS_MACRO_KEY_GENERATE_ALL_CONFIGURED_RUNNER_MACROS), vbBinaryCompare) = 0 Then SetMacroGenerationCompletionDeferral False
        ReportSystemCommandCancelled systemMacroKey, keyText
        ClearSystemCommandResult
        On Error Resume Next
        If scanNamingSnapshotOwned Then InvalidateNamingSnapshot
        If scanSystemMacroDefinitionSnapshotOwned Then EndSystemMacroDefinitionSnapshot
        If scanManagedDefinitionSnapshotOwned Then EndManagedObjectDefinitionSnapshot
        On Error GoTo 0
        EndTrackedProcess
        If traceRunOwnedByProcedure Then CompleteTraceRun "System command cancelled: " & keyText, False
        RunOperationalSystemCommand = False
        mOperationalSystemCommandActive = False
        Exit Function
    End If
    If (StrComp(keyText, UCase$(SYS_MACRO_KEY_BUILD_REPAIR_HELP_DOCUMENT_MACROS), vbBinaryCompare) = 0 Or StrComp(keyText, UCase$(SYS_MACRO_KEY_RUN_HEALTH_REVIEW), vbBinaryCompare) = 0 Or StrComp(keyText, UCase$(SYS_MACRO_KEY_GENERATE_SELECTED_RUNNER_MACROS), vbBinaryCompare) = 0 Or StrComp(keyText, UCase$(SYS_MACRO_KEY_GENERATE_ALL_CONFIGURED_RUNNER_MACROS), vbBinaryCompare) = 0) And modalProgressControlOwned Then
        If Len(keyText) > 0 Then
            PulseSystemCommandStatus "System command failed: " & keyText
        Else
            PulseSystemCommandStatus "System command failed."
        End If
        EndSetupControl
        modalProgressControlOwned = False
    Else
        If modalProgressControlOwned Then EndSetupControl
        modalProgressControlOwned = False
        If Len(keyText) > 0 Then
                PulseSystemCommandStatus "System command failed: " & keyText
            Else: PulseSystemCommandStatus "System command failed."
        End If
    End If
    If StrComp(keyText, UCase$(SYS_MACRO_KEY_GENERATE_SELECTED_RUNNER_MACROS), vbBinaryCompare) = 0 Or StrComp(keyText, UCase$(SYS_MACRO_KEY_GENERATE_ALL_CONFIGURED_RUNNER_MACROS), vbBinaryCompare) = 0 Then SetMacroGenerationCompletionDeferral False
    On Error Resume Next
    If scanNamingSnapshotOwned Then InvalidateNamingSnapshot
    If scanSystemMacroDefinitionSnapshotOwned Then EndSystemMacroDefinitionSnapshot
    If scanManagedDefinitionSnapshotOwned Then EndManagedObjectDefinitionSnapshot
    On Error GoTo 0
    ReportSystemCommandError systemMacroKey, keyText, savedErrNumber, savedErrSource, savedErrDescription
    ClearSystemCommandResult
    EndTrackedProcess
    If traceRunOwnedByProcedure Then CompleteTraceRun "System command failed: " & keyText, False
    RunOperationalSystemCommand = False
    mOperationalSystemCommandActive = False
End Function

Private Sub ReportSystemCommandCancelled(ByVal rawSystemMacroKey As String, ByVal resolvedKeyText As String)
    On Error Resume Next
    Dim detailText As String
    detailText = "Requested system macro key: " & rawSystemMacroKey
    If Len(resolvedKeyText) > 0 Then detailText = detailText & vbCrLf & "Resolved command key: " & resolvedKeyText
    PulseSystemCommandStatus "System command cancelled: " & resolvedKeyText
    LogTraceMessage "Info", "SystemCommand", "RunOperationalSystemCommand", "System command reported cancellation. Review the child diagnostics because no changes may have occurred or the command may have stopped after partial work.", detailText, resolvedKeyText, 0, vbNullString, vbNullString, resolvedKeyText, "SystemMacro", "CodeDb", "SystemCommand"
End Sub

Private Sub ReportSystemCommandError(ByVal rawSystemMacroKey As String, ByVal resolvedKeyText As String, ByVal errNumber As Long, ByVal errSource As String, ByVal errDescription As String)
    On Error Resume Next
    Dim detailText As String
    Dim messageText As String
    Dim popupText As String
    Dim popupTitle As String
    detailText = "Requested system macro key: " & rawSystemMacroKey
    If Len(resolvedKeyText) > 0 Then detailText = detailText & vbCrLf & "Resolved command key: " & resolvedKeyText
    messageText = EntryPointErrorMessage("Access Development Console system command failed.", "RunOperationalSystemCommand", errNumber, errSource, errDescription, detailText)
    LogTraceMessage "Error", "SystemCommand", "RunOperationalSystemCommand", "System command failed.", messageText, resolvedKeyText, errNumber, errSource, errDescription, resolvedKeyText, "SystemMacro", "CodeDb", "SystemCommand"
    If IsModuleGuardVerificationFailure(errSource, errDescription) Then
        popupText = Trim$(FailureProgressContextForMessage())
        If Len(popupText) > 0 Then popupText = popupText & vbCrLf & vbCrLf
        popupText = popupText & Trim$(errDescription)
        popupTitle = "ADC Module Guard"
    Else
        popupText = messageText
        popupTitle = "Access Development Console System Command Failed"
    End If
    CoreShowErrorMessage popupText, vbCritical, popupTitle, "RunOperationalSystemCommand", "Core"
End Sub

Private Sub ReportSystemCommandSuccess(ByVal rawSystemMacroKey As String, ByVal resolvedKeyText As String)
    On Error Resume Next
    Dim messageText As String
    messageText = "Access Development Console system command completed successfully." & vbCrLf & vbCrLf & "Requested system macro key: " & rawSystemMacroKey
    If Len(resolvedKeyText) > 0 Then messageText = messageText & vbCrLf & "Resolved command key: " & resolvedKeyText
    If StrComp(resolvedKeyText, UCase$(SYS_MACRO_KEY_BUILD_REPAIR_HELP_DOCUMENT_MACROS), vbBinaryCompare) = 0 And Len(Trim$(mBuildRepairHelpDocumentMacrosMacroResultDetail)) > 0 Then messageText = messageText & vbCrLf & vbCrLf & mBuildRepairHelpDocumentMacrosMacroResultDetail
    If SystemCommandRequiresRibbonRestartNotice(resolvedKeyText) Then messageText = messageText & vbCrLf & vbCrLf & "CLOSE AND REOPEN THE DATABASE TO LOAD THE UPDATED RIBBON."
    If CoreOperationHasWarnings() Then messageText = messageText & CoreOperationSummaryForMessage()
    messageText = messageText & TableDiagnosticsLocationForMessage() & TraceSummaryForMessage("success")
    ReportCoreActionSuccess "RunOperationalSystemCommand", "SystemCommand", SetupSystemCommandStatusText(resolvedKeyText, "System command completed successfully: " & resolvedKeyText, True), messageText, resolvedKeyText, resolvedKeyText, "SystemMacro", "CodeDb", "SystemCommand", True, "Access Development Console System Command"
End Sub

Private Function SystemCommandRequiresRibbonRestartNotice(ByVal resolvedKeyText As String) As Boolean
    If mRibbonRestartRequiredCurrentRun = False Then Exit Function
    Select Case resolvedKeyText
        Case UCase$(SYS_MACRO_KEY_RUN_FULL_SETUP), UCase$(SYS_MACRO_KEY_BUILD_RIBBON), "APPLYCUSTOMRIBBON", UCase$(SYS_MACRO_KEY_DISABLE_RIBBON), UCase$(SYS_MACRO_KEY_ADD_TABS_TO_ACCESS_RIBBON), UCase$(SYS_MACRO_KEY_BUILD_REPAIR_HELP_DOCUMENT_MACROS)
            SystemCommandRequiresRibbonRestartNotice = True
    End Select
End Function

Private Function SystemCommandShowsSuccessPopup(ByVal resolvedKeyText As String) As Boolean
    Select Case resolvedKeyText
        Case UCase$(SYS_MACRO_KEY_RUN_FULL_SETUP), _
             UCase$(SYS_MACRO_KEY_BUILD_REPAIR_ALL_TABLES_AND_QUERIES), "UPDATESTRUCTURE", UCase$(SYS_MACRO_KEY_RESOLVE_PENDING_SETUP_RECOVERY), _
             UCase$(SYS_MACRO_KEY_REGISTER_APPROVED_BUILD)
            SystemCommandShowsSuccessPopup = True
        Case UCase$(SYS_MACRO_KEY_BUILD_CHATGPT_UI_FORMS), UCase$(SYS_MACRO_KEY_DELETE_CHATGPT_UI_FORMS), UCase$(SYS_MACRO_KEY_BUILD_RIBBON), "APPLYCUSTOMRIBBON", _
             UCase$(SYS_MACRO_KEY_REPAIR_MANAGED_OBJECT_CATALOG), UCase$(SYS_MACRO_KEY_VALIDATE_MANAGED_OBJECT_CATALOG), _
             UCase$(SYS_MACRO_KEY_BACKUP_ALL_USER_DATA), UCase$(SYS_MACRO_KEY_RESET_ALL_USER_DATA), UCase$(SYS_MACRO_KEY_BUILD_REPAIR_HELP_DOCUMENT_MACROS), UCase$(SYS_MACRO_KEY_CLEAR_TRACE_LOG_FILE), _
             UCase$(SYS_MACRO_KEY_BUILD_REPAIR_CUSTOM_MACRO_LAUNCHERS), UCase$(SYS_MACRO_KEY_BUILD_UPDATE_CUSTOM_MACRO_LAUNCHER_BUTTONS), UCase$(SYS_MACRO_KEY_CLEAN_STALE_CUSTOM_MACRO_LAUNCHER_BUTTONS), _
             UCase$(SYS_MACRO_KEY_ENABLE_CUSTOM_MACRO_LAUNCHER_AUTOMATION), UCase$(SYS_MACRO_KEY_DISABLE_CUSTOM_MACRO_LAUNCHER_AUTOMATION)
            SystemCommandShowsSuccessPopup = True
        Case UCase$(SYS_MACRO_KEY_SCAN_ALL_DB_INVENTORY), UCase$(SYS_MACRO_KEY_SCAN_RUNNER_MACRO_CANDIDATES), UCase$(SYS_MACRO_KEY_ACCEPT_SELECTED_RUNNER_MACRO_CANDIDATES), UCase$(SYS_MACRO_KEY_MARK_STALE_CANDIDATES), UCase$(SYS_MACRO_KEY_DELETE_SELECTED_DUPLICATE_ROWS), _
             UCase$(SYS_MACRO_KEY_RESOLVE_RUNNER_MACRO_TARGETS), UCase$(SYS_MACRO_KEY_VALIDATE_RUNNER_MACRO_TARGETS), UCase$(SYS_MACRO_KEY_RUN_HEALTH_REVIEW), UCase$(SYS_MACRO_KEY_FIND_BROKEN_FORM_AND_REPORT_REFERENCES), UCase$(SYS_MACRO_KEY_REVIEW_STALE_CANDIDATES), _
             UCase$(SYS_MACRO_KEY_REVIEW_CUSTOM_MACRO_LAUNCHER_BUTTONS), UCase$(SYS_MACRO_KEY_SCAN_DB_INVENTORY_ACTIVE_REFERENCES), UCase$(SYS_MACRO_KEY_SCAN_DB_INVENTORY_AVAILABLE_REFERENCES), _
             UCase$(SYS_MACRO_KEY_SCAN_DB_INVENTORY_LOADED_VBA_PROJECTS), UCase$(SYS_MACRO_KEY_SCAN_DB_INVENTORY_ADD_INS_AND_LOADED_VBA_PROJECTS), UCase$(SYS_MACRO_KEY_GENERATE_SELECTED_RUNNER_MACROS), UCase$(SYS_MACRO_KEY_GENERATE_ALL_CONFIGURED_RUNNER_MACROS), _
             UCase$(SYS_MACRO_KEY_GENERATE_PROMPTED_RUNNER_MACROS)
            SystemCommandShowsSuccessPopup = True
        Case Else
            SystemCommandShowsSuccessPopup = False
    End Select
End Function

Private Sub RecoverUIBeforeSetup(ByVal systemCommandKey As String)
    On Error GoTo ErrHandler
    Dim normalizedKey As String
    normalizedKey = UCase$(Trim$(systemCommandKey))
    If normalizedKey <> UCase$(SYS_MACRO_KEY_RUN_FULL_SETUP) Then Exit Sub
    If UIRecoveryBlocked() = False Then Exit Sub
    Debug.Print "ADC UI recovery preflight started."
    Debug.Print "System command: " & normalizedKey
    Debug.Print "Stage: Before trace initialization and top-level ADC Setup recovery gates."
    RetryUIRecovery False
    If UIRecoveryBlocked() Then Err.Raise ERR_UI_PREFLIGHT_RECOVERY_FAILED, "RecoverUIBeforeSetup", "Access Development Console UI recovery remained blocked after the retry returned."
    Debug.Print "ADC UI recovery preflight completed and was verified."
    Debug.Print "System command: " & normalizedKey
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim recoveryStillBlocked As Boolean
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    recoveryStillBlocked = UIRecoveryBlocked()
    On Error GoTo 0
    Debug.Print "ADC UI recovery preflight failed."
    Debug.Print "System command: " & normalizedKey
    Debug.Print "Stage: Before trace initialization and top-level ADC Setup recovery gates."
    Debug.Print "Recovery still blocked: " & CStr(recoveryStillBlocked)
    Debug.Print "Original Err.Number: " & CStr(savedErrNumber)
    Debug.Print "Original Err.Source: " & savedErrSource
    Debug.Print "Original Err.Description: " & savedErrDescription
    Err.Raise ERR_UI_PREFLIGHT_RECOVERY_FAILED, "RecoverUIBeforeSetup", "Could not recover retained Access Development Console UI state before starting ADC Setup." & vbCrLf & "System command: " & normalizedKey & vbCrLf & "Recovery still blocked: " & CStr(recoveryStillBlocked) & vbCrLf & "Original Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Original Err.Source: " & savedErrSource & vbCrLf & "Original Err.Description: " & savedErrDescription
End Sub

Private Function IsModuleGuardVerificationFailure(ByVal errSource As String, ByVal errDescription As String) As Boolean
    IsModuleGuardVerificationFailure = (InStr(1, errSource, "ModuleGuard", vbTextCompare) > 0 Or InStr(1, errSource, "ADCModuleUpdate", vbTextCompare) > 0 Or InStr(1, errDescription, "module update guard", vbTextCompare) > 0 Or InStr(1, errDescription, "Module Guard expects", vbTextCompare) > 0 Or InStr(1, errDescription, "VBA component name", vbTextCompare) > 0)
End Function

Private Function SystemMacroCommandKeyFromRunArgument(ByVal systemMacroArgument As String) As String
    Dim argumentText As String
    Dim keyArgumentText As String
    Dim rowData As Variant
    Dim designatorMatchKey As String
    Dim rowKey As String
    argumentText = systemMacroArgument
    keyArgumentText = Trim$(systemMacroArgument)
    If Len(keyArgumentText) = 0 Then Exit Function
    For Each rowData In SystemMacroCatalog()
        PumpSetupResponsiveness
        rowKey = Trim$(CStr(rowData(0)))
        If StrComp(rowKey, keyArgumentText, vbTextCompare) = 0 Then
            SystemMacroCommandKeyFromRunArgument = rowKey
            Exit Function
        End If
        If Len(designatorMatchKey) = 0 Then
            If StrComp(CStr(rowData(1)), argumentText, vbTextCompare) = 0 Then designatorMatchKey = rowKey
        End If
    Next rowData
    If Len(designatorMatchKey) > 0 Then SystemMacroCommandKeyFromRunArgument = designatorMatchKey Else SystemMacroCommandKeyFromRunArgument = keyArgumentText
End Function

Private Sub ResolvePendingSetupRecovery(Optional ByVal showMessage As Boolean = True)
    'Confirms and retries only the durable recovery boundaries that are currently supported, then verifies every boundary is clear.
    On Error GoTo ErrHandler
    Dim catalogBlocked As Boolean
    Dim reviewBlocked As Boolean
    Dim uiBlocked As Boolean
    Dim pendingSummary As String
    Dim recoveredSummary As String
    Dim promptText As String
    Dim response As VbMsgBoxResult
    CoreResetOperationNotes
    catalogBlocked = CatalogRecoveryBlocked()
    reviewBlocked = ConfigCandidateMarkingRecoveryBlocked()
    uiBlocked = UIRecoveryBlocked()
    If catalogBlocked Then pendingSummary = SharedText_AppendNonBlankLine(pendingSummary, " - Managed Object Catalog transaction recovery")
    If reviewBlocked Then pendingSummary = SharedText_AppendNonBlankLine(pendingSummary, " - Config/Candidate marking transaction recovery")
    If uiBlocked Then pendingSummary = SharedText_AppendNonBlankLine(pendingSummary, " - UI form or ribbon transaction recovery")
    If Len(pendingSummary) = 0 Then
        CoreAddOperationNote "No pending Access Development Console recovery work was found; no recovery routine ran."
        If showMessage Then CoreShowMessage "No pending Access Development Console recovery work was found." & TableDiagnosticsLocationForMessage(), vbInformation, "Resolve Pending Setup Recovery"
        Exit Sub
    End If
    promptText = "Access Development Console found the following unfinished recovery work:" & vbCrLf & vbCrLf & pendingSummary
    promptText = promptText & vbCrLf & vbCrLf & "Recovery may roll back incomplete changes or finish verified cleanup retained by the failed operation."
    promptText = promptText & vbCrLf & "Only the recovery boundaries listed above will be retried. No setup, repair, generation, or ribbon rebuild will run."
    promptText = promptText & vbCrLf & vbCrLf & "Continue?"
    response = CoreShowMessageResult(promptText, vbExclamation Or vbYesNo Or vbDefaultButton2, "Resolve Pending Setup Recovery")
    If response <> vbYes Then
        CoreAddOperationNote "Pending Access Development Console recovery was not started because confirmation was declined."
        MarkSystemCommandCancelled
        Exit Sub
    End If
    If CatalogRecoveryBlocked() Then
        CorePulseProgress "Resolve pending Setup recovery: Catalog recovery"
        RetryCatalogRecovery False
        If CatalogRecoveryBlocked() Then Err.Raise vbObjectError + 6339, "ResolvePendingSetupRecovery", "Catalog recovery returned without clearing its durable blocker."
        recoveredSummary = SharedText_AppendNonBlankLine(recoveredSummary, " - Catalog recovery completed and verified")
    End If
    If ConfigCandidateMarkingRecoveryBlocked() Then
        CorePulseProgress "Resolve pending Setup recovery: candidate-marking recovery"
        RetryConfigCandidateMarkingRecovery False
        If ConfigCandidateMarkingRecoveryBlocked() Then Err.Raise vbObjectError + 6340, "ResolvePendingSetupRecovery", "candidate-marking recovery returned without clearing its durable blocker."
        recoveredSummary = SharedText_AppendNonBlankLine(recoveredSummary, " - candidate-marking recovery completed and verified")
    End If
    If UIRecoveryBlocked() Then
        CorePulseProgress "Resolve pending Setup recovery: UI recovery"
        RetryUIRecovery False
        If UIRecoveryBlocked() Then Err.Raise vbObjectError + 6341, "ResolvePendingSetupRecovery", "UI recovery returned without clearing its durable blocker."
        recoveredSummary = SharedText_AppendNonBlankLine(recoveredSummary, " - UI recovery completed and verified")
    End If
    catalogBlocked = CatalogRecoveryBlocked()
    reviewBlocked = ConfigCandidateMarkingRecoveryBlocked()
    uiBlocked = UIRecoveryBlocked()
    If catalogBlocked Or reviewBlocked Or uiBlocked Then
        pendingSummary = vbNullString
        If catalogBlocked Then pendingSummary = SharedText_AppendNonBlankLine(pendingSummary, " - Catalog recovery remains blocked")
        If reviewBlocked Then pendingSummary = SharedText_AppendNonBlankLine(pendingSummary, " - candidate-marking recovery remains blocked")
        If uiBlocked Then pendingSummary = SharedText_AppendNonBlankLine(pendingSummary, " - UI recovery remains blocked")
        Err.Raise vbObjectError + 6342, "ResolvePendingSetupRecovery", "Recovery stopped because one or more durable blockers remain:" & vbCrLf & pendingSummary
    End If
    If Len(recoveredSummary) = 0 Then Err.Raise vbObjectError + 6343, "ResolvePendingSetupRecovery", "Recovery was confirmed, but no detected recovery boundary completed. No success was reported."
    CoreAddOperationNote "Pending Access Development Console recovery completed and every supported recovery boundary was verified clear." & vbCrLf & recoveredSummary
    If showMessage Then CoreShowMessage "Pending Access Development Console recovery completed successfully." & vbCrLf & vbCrLf & recoveredSummary & TableDiagnosticsLocationForMessage(), vbInformation, "Resolve Pending Setup Recovery"
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    CoreAddOperationNote "Pending Access Development Console recovery remains unresolved. Error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & SharedMessage_OneLine(savedErrDescription)
    If showMessage Then
        CoreShowErrorMessage "Could not complete and verify pending Access Development Console recovery." & vbCrLf & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, vbCritical, "Recover Pending Access Development Console Work", "ResolvePendingSetupRecovery", "Core"
        Exit Sub
    End If
    Err.Raise savedErrNumber, "ResolvePendingSetupRecovery", "Could not complete and verify pending Access Development Console recovery." & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Sub

' Purpose: Requires same database UI command for source procedure, action text.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   sourceProcedure (Input; String): Procedure name recorded as the source of trace, error, status, or diagnostic information emitted by this call.
'   actionText (Input; String): Input action text for require same database UI command; is tested by If CodeDbIsHostDb() = False.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub RequireSameDatabaseUICommand(ByVal sourceProcedure As String, ByVal actionText As String)
    If CodeDbIsHostDb() = False Then Err.Raise vbObjectError + 6292, sourceProcedure, actionText & " requires CodeDb and HostDb to be the same database. This system command was stopped before invoking the UI module so the parent command cannot report a false success." & vbCrLf & "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB
End Sub
