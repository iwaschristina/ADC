Attribute VB_Name = "x__UTIL_ADC_REVIEW"
' ================================================================================================================================================================================
' ChatGPT Rewrite Tracking Header
' ================================================================================================================================================================================
' ChatGPT Previous Version ID: CGPT-GCT8-TPE5-KW8F-3TGD-AQGF
' ChatGPT Baseline Version ID: CGPT-GCT8-TPE5-KW8F-3TGD-AQGF
' ChatGPT Version ID: CGPT-VTN2-HDJA-KYDY-4LR5-U9PU
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
' STANDARD MODULE: Access Development Console Review and Navigation
' VB_Name: x__UTIL_ADC_REVIEW
' Entry Points:
'     EnsureConfiguredNavigationPaneGroupsExist
'     PreflightManagedObjectsNavigationPaneForSetup
'     ApplyPreflightedManagedObjectsNavigationPaneForSetup
' SHARED CODE:
'     SharedCollection_Join(ByVal values As Collection, Optional ByVal delimiter As String = ", ") -> String
'     SharedCollection_KeyExists(ByVal values As Collection, ByVal key As String) -> Boolean
'     SharedDao_QueryExists(ByVal db As DAO.Database, ByVal queryName As String) -> Boolean
'     SharedDao_SafeDatabaseName(ByVal db As DAO.Database) -> String
'     SharedDao_SqlQuote(ByVal value As Variant) -> String
'     SharedDao_TableDefFieldExistsSafe(ByVal tdf As DAO.tableDef, ByVal fieldName As String) -> Boolean
'     SharedDao_TableExists(ByVal db As DAO.Database, ByVal tableName As String) -> Boolean
'     SharedDao_TableFieldExists(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String) -> Boolean
'     SharedNavPane_AddAccessObjectToExistingCustomGroup(ByVal db As DAO.Database, ByVal objectName As String, ByVal objectType As AcObjectType, ByVal groupID As Long, ByRef detailText As String, Optional ByVal
'     forceRefreshShortcut As Boolean = False, Optional ByVal ensureObjectIDRow As Boolean = True) -> Boolean
'     SharedNavPane_EnsureObjectIDRow(ByVal objectName As String, Optional ByVal objectType As AcObjectType = acForm, Optional ByVal tempPrefix As String = "Shared_NavPane_") -> None
'     SharedNavPane_ExistingCustomGroupID(ByVal db As DAO.Database, ByVal groupName As String) -> Long
'     SharedNavPane_MSysObjectID(ByVal db As DAO.Database, ByVal objectName As String, ByVal objectType As AcObjectType) -> Long
'     SharedText_AppendNonBlankLine(ByVal existingText As String, ByVal lineText As String) -> String
'     SharedText_Nz(ByVal v As Variant) -> String
'     SharedValue_NzLong(ByVal value As Variant, Optional ByVal defaultValue As Long = 0) -> Long
' CLASSES USED: None
' DYNAMIC STRING REFERENCES:
'     Dynamic-reference scan stopped after 2000 logical line(s) to prevent a long-running scan; review this module manually if needed.
' LastScanTimestamp: Saturday 8/15/26 8:46 AM
' Purpose:
'     Owns genuine review/health behavior: Health Review, duplicate-review expressions/findings, managed-description review, G1 managed-description reconciliation/whole-catalog exact-state verification, and the
'     I1-I3 complete-domain Navigation Pane engine/placement authority. The read-only D4/E3/G1 predecessor state is consumed by H1. Candidate/configuration operational behavior
'     and configured macro generation/execution are
'     owned by x__UTIL_ADC_MACRO_RUNTIME.
' Author: Christina Shonts
' Support: Access Development Console CodeDb utility workflows.
' Code Category: Access Development Console review/health and Navigation Pane authority.
' Requirements: Microsoft Access, DAO, the companion Access Development Console modules, and x__SHARED_CODE.
' Reference Dependencies:
'     Access object model, DAO, and current Access Development Console shared/common modules.
' DB Object Dependencies:
'     Uses the current CodeDb-owned Access Development Console tables, queries, forms, macros, and Navigation Pane metadata appropriate to this module's scope.
' Usage:
'     Import with the companion Access Development Console modules. Existing public procedure names remain stable for macros, forms, queries, and Application.Run callers; A3 saved queries call the restored
'     duplicate-review Public Functions directly from Access SQL.
' Version:
'     Shifts managed-description and Navigation Pane Setup ownership to G1 and I1-I3 and their Ribbon predecessor handoff to H1 after insertion of dependency Phase F.
'     Removes safe remaining Trim$ normalization from normal H2/H3 Navigation Pane physical-object identity paths while retaining backup/recovery tolerance.
'     Adds explicit managed-description, duplicate-review, Navigation Pane H1-H3, and Health Review section boundaries without changing Review ownership or behavior.
'     Routes macro-runtime calls to x__UTIL_ADC_MACRO_RUNTIME and adopts TMacroRecord/MacroRecordFromRecordset while preserving review behavior and Setup-phase ownership.
'     Removes the two remaining direct Trim$ mutations from the F1 managed-object physical Description path; ownership, validation, and the existing Access property-size policy are unchanged.
'     Routes Managed Object Catalog references to x__UTIL_ADC_MOC and physical lifecycle references to x__UTIL_ADC_MOC_LIFECYCLE after the responsibility split; runtime and Setup-phase behavior are unchanged.
'     Adds reusable standalone managed-object Description reconciliation and documents blank Help Document macro descriptions as an intentional managed-object exception.
'     Replaces overlapping managed-description scope Booleans with one explicit internal scope enum and narrows module-internal procedure visibility.
'     Converged Setup physical Navigation Pane authority into H1 group creation plus one complete-domain H2 plan/H3 apply-proof, while retaining explicit direct/runtime grouping paths outside Setup.
'     Managed-header dependency metadata updated because x__UTIL_ADC_DEFAULTS is now a standard module; executable behavior unchanged.
'     Inlined the duplicate-group length-prefixed signature expression into DuplicateGroupKeyFromValues and removed the unnecessary private micro-helper.
'     Centralized every HealthReview current-run lifecycle, deduplication context, and finding write under Review; stale-candidate preview refreshes only its own findings inside the single current run (or starts one
'     only when none exists), fails closed on legacy competing current runs, and AddHealthFinding is private.
'     Consolidated genuine review/health behavior here while moving Config/generation runtime behavior to x__UTIL_ADC_CONFIG and removing duplicate review helpers.
'     Consolidated managed-object identity normalization under DEFAULTS and DBEngine review diagnostics under REVIEW.
'     Exposed the existing managed-object Description placeholder policy for direct Health Review reuse.
'     Removed the duplicate Review recovery blocker/key and redirected Navigation Pane setup rollback tracking to the core setup journal authority.
'     Clarified that duplicate-review cache invalidation is source-mutation-owned rather than a pre-open responsibility of the duplicate-review query shortcut.
' Notes:
'     Authority consolidation moves procedure bodies directly; it does not add forwarding wrappers or change the table-driven command contract. The A3 duplicate-review query-expression dependency closure is retained
'     because Access resolves those Public Functions at saved-query preflight/open time. Phase I adds no Phase dependency: I1 owns group creation, I2 is read-only, and I3 applies only the validated I2 membership
'     plan, including transient physical ADC backup tables discovered outside MAC_Resolved.
' Core Rules Enforced Here:
'     Preserve diagnostics, cancellation, destructive safety, rollback, scope-first naming, and explicit CodeDb/HostDb ownership; Review is the only physical Navigation Pane group/membership/object-identity mutation
'     and verification owner; I2 performs no physical mutation; I3 requires the successful same-session I2 plan, applies missing memberships, and enforces rollback-safe exclusive Backup-group placement for transient
'     ADC backup tables.
' Design Map:
'     Procedures are grouped by the module scope named above; Phase I uses I1 configured-group creation/proof, I2 cached read-only physical/group/object/membership preflight plus dynamic ADC backup discovery from
'     the existing MSysObjects snapshot, and I3 rollback-safe exclusive backup cleanup, batch membership insertion, one refresh, final verification, and rollback-safe persistence. Generated-form placement is owned
'     here for direct UI build requests; Setup form generation defers physical placement to the complete-domain I3 pass.
' Configuration: Existing settings and catalog rows remain authoritative.
' Behavior:
'     Existing entry-point behavior and names are preserved. MacroCatalog_Records_DuplicateRows and its duplicate-review expressions remain openable during A3 on a fresh Phase-A setup. I2 remains read-only and
'     proves MSysObjects identity plus any already-materialized Navigation Pane mirror while adding canonical ADCBak_* tables to the complete-domain plan without MAC rows; I3 materializes only missing exact mirror
'     identities under rollback before membership writes and verify the settled state.
'     G1 retains one same-trace authoritative target plan, physical Description snapshot, and DAO target-object cache across exact-state detection, mutation, and compact-state persistence, then releases the cache
'     after completion.
' Code Status:
'     Active canonical Review authority; Phase I I1-I3 Navigation Pane engine and Phase G G1 managed-description finalization implemented.
' Todo:
'     Regenerate managed header public procedure metadata after importing this rewrite because Review now owns its mutable state and exposes the grouping-warning query.
' Reporting Bugs: None currently recorded in this header.
' Module Level Bugs: None currently recorded in this header.
' Public Subs:
'     EnsureConfiguredNavigationPaneGroupsExist(ByVal sourceName As String, Optional ByVal failOnCreationError As Boolean = True) -> None
'     RollbackSetupNavPaneShortcut(ByVal groupID As Long, ByVal objectID As Long, ByRef deletedCount As Long, ByRef detailText As String) -> None
'     RollbackSetupNavPaneShortcutRestore(ByVal groupID As Long, ByVal objectID As Long, ByVal priorName As Variant, ByVal priorType As Variant, ByVal priorFlags As Variant, ByVal priorPosition As Variant, ByRef
'     restoredCount As Long, ByRef detailText As String) -> None
'     RollbackSetupNavPaneObjectIDRestore(ByVal objectID As Long, ByVal repairedObjectName As String, ByVal repairedObjectType As Long, ByVal priorObjectName As String, ByVal priorObjectType As Long, ByRef
'     restoredCount As Long, ByRef detailText As String) -> None
'     RollbackSetupDeletedNavPaneObjectIDRestore(ByVal objectID As Long, ByVal objectName As String, ByVal objectType As Long, ByRef restoredCount As Long, ByRef detailText As String) -> None
'     RollbackSetupNavPaneObjectID(ByVal objectID As Long, ByVal objectName As String, ByVal objectType As Long, ByRef deletedCount As Long, ByRef retainedSharedIdentityCount As Long, ByRef detailText As String)
'     -> None
'     RollbackSetupNavPaneGroup(ByVal groupID As Long, ByVal groupName As String, ByRef deletedCount As Long, ByRef detailText As String) -> None
'     ApplyGeneratedFormsNavigationPane(ByVal formNames As Collection, ByRef navNotes As Collection) -> None
'     PreflightManagedObjectsNavigationPaneForSetup() -> None
'     ApplyPreflightedManagedObjectsNavigationPaneForSetup(Optional ByVal showMessage As Boolean = False) -> None
'     AddMacroToConfiguredNavigationPaneGroup(ByVal MacroTitle As String, ByVal configuredGroupName As String, Optional ByVal forceRefreshShortcut As Boolean = False) -> None
'     InitializeNavigationPaneGroupingNotes() -> None
'     RunHealthReview(Optional ByVal showMessage As Boolean = True) -> None
'     PersistStaleCandidateReviewHealthFindings(ByVal findings As Collection) -> None
'     OpenReviewQuery(ByVal queryName As String, ByVal sourceName As String) -> None
'     RequireDefaultCodeDbGeneratedMacrosNavigationPaneGroupForRoleContext(ByVal sourceName As String, Optional ByVal showMessage As Boolean = True) -> None
' Public Functions:
'     InvalidateDuplicateReviewCache() -> Boolean
'     ManagedDescriptionExpectedPhysicalText(ByVal authoritativeDescription As String) -> String
'     ApplyAndVerifyManagedObjectDescriptionForStandaloneRepair(ByVal objectName As String, ByVal objectTypeCode As String, Optional ByRef descriptionChanged As Variant) -> Boolean
'     ApplyAndVerifyManagedDescriptionsForSetup(Optional ByVal showMessage As Boolean = False, Optional ByRef detailsForMessage As String = vbNullString) -> Boolean
'     BeginHelpDocumentMacroBuildRepairFullRibbonPredecessorProof(ByRef detailText As String, Optional ByRef provenCatalogTokenForBuildRepair As String = vbNullString) -> Boolean
'     ApplyAndVerifyHelpDocumentManagedDescriptionsForBuildRepair(Optional ByRef detailsForMessage As String = vbNullString, Optional ByRef catalogDescriptionMutated As Boolean = False, Optional ByRef
'     refreshedCatalogToken As String = vbNullString) -> Boolean
'     HelpDocumentDescriptionSameRunG1ProofForRibbon(ByRef proofToken As String, ByRef proofTraceRunID As String, ByRef detailText As String) -> Boolean
'     FinalizeHelpDocumentMacroBuildRepairFullRibbonPredecessorProof(ByVal currentCatalogToken As String, ByVal scopedHelpProofToken As String, ByRef proofToken As String, ByRef proofTraceRunID As String, ByRef
'     detailText As String) -> Boolean
'     HelpDocumentMacroBuildRepairSameRunFullRibbonPredecessorProof(ByRef proofToken As String, ByRef proofTraceRunID As String, ByRef detailText As String, Optional ByRef proofCatalogToken As String =
'     vbNullString) -> Boolean
'     ConfigRowHasDuplicateForQuery(ByVal rowID As Variant) -> Boolean
'     DuplicateIssueTypeForQuery(ByVal rowID As Variant) -> String
'     DuplicateIssueDetailForQuery(ByVal rowID As Variant) -> String
'     DuplicateGroupKeyForQuery(ByVal rowID As Variant) -> String
'     DuplicateGroupCountForQuery(ByVal rowID As Variant) -> Long
'     DuplicateSelectedCountForQuery(ByVal rowID As Variant) -> Long
'     DuplicateAllRowsSelectedForQuery(ByVal rowID As Variant) -> Boolean
'     ConfiguredNavigationPaneGroupsCurrentForSetup(ByRef detailText As String, Optional ByVal allowSameRunProofReuse As Boolean = False) -> Boolean
'     ReconcileHelpDocumentMacroNavigationPane(Optional ByRef detailText As String = vbNullString) -> Boolean
'     SetupNavigationPanePlanReady(ByRef planToken As String, ByRef planRowCount As Long) -> Boolean
'     CompleteNavigationPanePlacementTokenForSetup(ByRef detailText As String, Optional ByVal allowSameRunProofReuse As Boolean = False) -> String
'     ReviewDBEngineErrorsText() -> String
'     RibbonDependencyTokenForSetup(ByRef detailText As String, Optional ByRef provenObjectDescriptionToken As String = vbNullString, Optional ByRef provenFieldDescriptionToken As String = vbNullString, Optional
'     ByRef provenCatalogToken As String = vbNullString, Optional ByVal persistedProofOnly As Boolean = False, Optional ByVal sameRunProvenCatalogToken As String = vbNullString, Optional ByVal
'     sameRunProvenCatalogTraceRunID As String = vbNullString) -> String
'     NavigationPaneGroupingSummaryForMessage() -> String
'     TextPropertyMatches(ByVal targetObject As Object, ByVal propertyName As String, ByVal expectedValue As String) -> Boolean
'     CountDuplicateRowsBySelectedState(ByVal selectedOnly As Boolean) -> Long
'     CountTargetSignatureRows(ByVal targetSignature As String) -> Long
' Public Properties: None
' ================================================================================================================================================================================

Option Compare Database
Option Explicit

' ===============================================================================================================================================================================
' MODULE-LEVEL DECLARATIONS
' ===============================================================================================================================================================================

' Managed-object description application scope contract.
Private Enum ManagedObjectDescriptionScope
    MANAGED_OBJECT_DESCRIPTION_SCOPE_ALL = 0
    MANAGED_OBJECT_DESCRIPTION_SCOPE_HELP_DOCUMENT_MACROS_ONLY = 3
End Enum

' Managed-object description token state.
Private mManagedObjectDescriptionToken As String
Private mManagedObjectDescriptionTokenTraceRunID As String
Private mHelpDocumentDescriptionProofToken As String
Private mHelpDocumentDescriptionProofTraceRunID As String
' Build/Repair Help Document Macros proves all unchanged predecessor domains before mutation (using the complete fast path when possible, or a Help-tolerant non-Help fallback), then advances them through same-run
' Help-only C/F changes. H1 receives only the resulting complete canonical proof.
Private mHelpDocumentMacroBuildRepairBaselineScopeProofToken As String
Private mHelpDocumentMacroBuildRepairBaselineScopeProofTraceRunID As String
Private mHelpDocumentMacroBuildRepairBaselineFieldDescriptionToken As String
' Operation-scoped compact-state cache: the pre-mutation Build/Repair Help Document Macros proof captures every non-Help object's compact physical metadata once; post-Help finalization reuses those proven states and
' rereads only Help macros.
Private mHelpDocumentMacroBuildRepairBaselineNonHelpCompactStateByObject As Object
Private mHelpDocumentMacroBuildRepairBaselineCompactStateTraceRunID As String
Private mHelpDocumentMacroBuildRepairCompactStateCaptureActive As Boolean
Private mHelpDocumentMacroBuildRepairFullPredecessorToken As String
Private mHelpDocumentMacroBuildRepairFullPredecessorTraceRunID As String
Private mHelpDocumentMacroBuildRepairFullPredecessorCatalogToken As String
' Operation-scoped keyed projection of active Help-document macros whose authoritative ObjectDescription may legitimately be blank.
Private mManagedDescriptionOptionalHelpMacroIndex As Object
Private mManagedDescriptionOptionalHelpPolicyDepth As Long

' Complete managed-object Setup Navigation Pane plan and membership token.
Private mSetupNavigationPanePlanReady As Boolean
' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION:
' This Collection is the complete managed-object Navigation Pane placement plan produced by preflight and later consumed by setup. It is derived from canonical managed-object/group definitions plus one observed pane
' state, remains valid only for its recorded plan token/current run, and is not a second placement authority.
Private mSetupNavigationPanePlan As Collection
Private mSetupNavigationPanePlanCheckedCount As Long
Private mSetupNavigationPanePlanAlreadyPresentCount As Long
Private mSetupNavigationPanePlanOptionalMissingCount As Long
Private mSetupNavigationPanePlanToken As String
Private Const NAVIGATION_PANE_MEMBERSHIP_TOKEN_PROPERTY As String = "ADCManagedNavigationPaneMembershipToken"

' Managed-description fingerprint policy.
Private Const MANAGED_DESCRIPTION_FINGERPRINT_PROPERTY As String = "ADCManagedDescriptionFingerprint"
Private Const MANAGED_DESCRIPTION_FINGERPRINT_VERSION As String = "2026-08-18 optional Help-document Description policy v4"
Private Const MANAGED_DESCRIPTION_PHYSICAL_STATE_PROPERTY As String = "ADCManagedDescriptionPhysicalStateV1"
Private Const MANAGED_DESCRIPTION_EXACT_TOKEN_PROPERTY As String = "ADCManagedDescriptionExactTokenV1"
Private Const MANAGED_DESCRIPTION_WHOLE_CATALOG_EXACT_TOKEN_PROPERTY As String = "ADCManagedDescriptionWholeCatalogExactTokenV1"
Private Const MANAGED_DESCRIPTION_WHOLE_CATALOG_EXACT_TOKEN_VERSION As String = "2026-08-18 G1 whole-catalog managed description exact state v3"
Private Const HELP_DOCUMENT_DESCRIPTION_PROOF_VERSION As String = "2026-08-18 operational Help-document managed description proof v1"
Private Const MANAGED_DESCRIPTION_PHYSICAL_STATE_VERSION As String = "2026-07-20 keyed managed-description compact physical state v3"
Private Const ACCESS_DESCRIPTION_PROPERTY_SAFE_BYTES As Long = 2000
' Duplicate-review cache.
Private mDuplicateReviewCacheReady As Boolean
' INTENTIONAL KEYED LOOKUP CACHE:
' These four dictionaries are one duplicate-review calculation cache: row -> issue/group plus group -> total/selected counts, all derived from current MC_Records
' configuration rows. InvalidateDuplicateReviewCache is the
' lifetime boundary; consumers read calculated results and do not own or mutate the dictionaries.
Private mDuplicateReviewIssueTypeByRowID As Object
Private mDuplicateReviewGroupKeyByRowID As Object
Private mDuplicateReviewGroupCountByKey As Object
Private mDuplicateReviewSelectedCountByKey As Object


' Navigation Pane instrumentation and metadata cache.
Private mNavPaneInstrumentationTraceRunID As String
Private mNavPanePhysicalGroupSnapshotBuildCount As Long
Private mNavPaneDefinitionSnapshotLoadCount As Long
Private mNavPaneCapabilityDetectionCount As Long
Private mNavPaneCustomCategoryLookupCount As Long
Private mNavPaneSameRunProofReuseCount As Long
Private mNavPaneMetadataCacheValid As Boolean
Private mNavPaneMetadataCacheTraceRunID As String
Private mNavPaneMetadataCacheHostDbName As String
Private mNavPaneMetadataHasName As Boolean
Private mNavPaneMetadataHasGroupCategoryID As Boolean
Private mNavPaneMetadataHasCategoryID As Boolean
Private mNavPaneMetadataHasType As Boolean
Private mNavPaneMetadataHasPosition As Boolean
Private mNavPaneMetadataCustomCategoryID As Long

' Complete-domain same-run Navigation Pane proof.
' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION:
' I3 retains the exact verified complete-domain rows for same-run proof reuse. The trace-run/HostDb/token checks bound its lifetime; it is distinct from the I2 preflight plan and is not a second placement authority.
Private mCompleteNavigationPaneProofRows As Collection
Private mCompleteNavigationPaneProofToken As String
Private mCompleteNavigationPaneProofTraceRunID As String
Private mCompleteNavigationPaneProofHostDbName As String

' I1 Navigation Pane proof reuse state.
Private mI1NavigationPaneProofRows As Collection
Private mI1NavigationPaneProofTraceRunID As String
Private mI1NavigationPaneProofHostDbName As String
Private mI1NavigationPaneProofReuseCount As Long

' Final I1 physical-group snapshot is retained only for the same trace-run/HostDb successor I2.
Private mI1NavigationPanePhysicalGroupSnapshot As Object


' Navigation Pane review/grouping run state.
' INTENTIONAL ORDERED RUN STATE / KEYED REPORT SUPPRESSION:
' Navigation-pane notes preserve diagnostic occurrence order for the active run, while mAccessObjectCatalogReportedTypes is a keyed membership set preventing duplicate success reporting within that run. Both are
' reset by their existing run-identity paths and are not external data authorities.
Private mNavigationPaneGroupingNotes As Collection
Private mNavigationPaneGroupingOwnerTraceRunID As String
Private mAccessObjectCatalogReportedRunID As String
Private mAccessObjectCatalogReportedTypes As Collection

' ===============================================================================================================================================================================
' PROCEDURES
' ===============================================================================================================================================================================

' =============================================================================================================================================================================
' MANAGED DESCRIPTION / G1 PROCEDURES
' =============================================================================================================================================================================

' Purpose: Clears the duplicate-review calculation cache when Config-changing workflows alter source rows; query opens no longer own cache invalidation.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters: None.
' Returns: Boolean: True on the function's affirmative/success branch for invalidate duplicate review cache; False on its explicit negative, unavailable, rejected, or handled-failure branch.
Public Function InvalidateDuplicateReviewCache() As Boolean
    'Clears the duplicate-review calculation cache when Config-changing workflows alter source rows; query opens no longer own cache invalidation.
    Set mDuplicateReviewIssueTypeByRowID = Nothing
    Set mDuplicateReviewGroupKeyByRowID = Nothing
    Set mDuplicateReviewGroupCountByKey = Nothing
    Set mDuplicateReviewSelectedCountByKey = Nothing
    mDuplicateReviewCacheReady = False
    InvalidateDuplicateReviewCache = True
End Function

Private Function ManagedDescriptionPropertySafeText(ByVal authoritativeDescription As String, ByRef wasBounded As Boolean, ByRef authoritativeCharacterLength As Long, ByRef authoritativeByteLength As Long, ByRef writtenCharacterLength As Long, ByRef writtenByteLength As Long) As String
    'Access error 3309 rejects property settings larger than 2 KB. The complete authoritative text remains in MAC.ObjectDescription.
    Dim markerText As String
    Dim prefixText As String
    Dim prefixByteBudget As Long
    Dim physicalDescription As String
    wasBounded = False
    authoritativeCharacterLength = 0
    authoritativeByteLength = 0
    writtenCharacterLength = 0
    writtenByteLength = 0
    authoritativeCharacterLength = Len(authoritativeDescription)
    authoritativeByteLength = LenB(authoritativeDescription)
    If authoritativeByteLength <= ACCESS_DESCRIPTION_PROPERTY_SAFE_BYTES Then
        physicalDescription = authoritativeDescription
    Else
        wasBounded = True
        markerText = vbCrLf & "[Full description retained in " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID) & "." & FIELD_OBJECT_DESCRIPTION & "; characters=" & CStr(authoritativeCharacterLength) & "; bytes=" & CStr(authoritativeByteLength) & "; fingerprint=" & StableTextHashHex(authoritativeDescription) & ".]"
        prefixByteBudget = ACCESS_DESCRIPTION_PROPERTY_SAFE_BYTES - LenB(markerText)
        If prefixByteBudget < 2 Then Err.Raise vbObjectError + 6450, "ManagedDescriptionPropertySafeText", "The deterministic managed-description marker exceeds the Access Description property safety budget."
        prefixText = ManagedDescriptionPrefixWithinByteBudget(authoritativeDescription, prefixByteBudget)
        physicalDescription = prefixText & markerText
        Do While LenB(physicalDescription) > ACCESS_DESCRIPTION_PROPERTY_SAFE_BYTES And Len(prefixText) > 0
            prefixText = Left$(prefixText, Len(prefixText) - 1)
            physicalDescription = RTrim$(prefixText) & markerText
        Loop
        If LenB(physicalDescription) > ACCESS_DESCRIPTION_PROPERTY_SAFE_BYTES Then Err.Raise vbObjectError + 6451, "ManagedDescriptionPropertySafeText", "The deterministic managed-description representation still exceeds the Access Description property safety budget."
    End If
    writtenCharacterLength = Len(physicalDescription)
    writtenByteLength = LenB(physicalDescription)
    If writtenByteLength > DAO_PROPERTY_TEXT_LIMIT_BYTES Then Err.Raise vbObjectError + 6452, "ManagedDescriptionPropertySafeText", "The managed-description representation exceeds the Access 2 KB property limit."
    ManagedDescriptionPropertySafeText = physicalDescription
End Function


' Purpose: Resolves managed description expected physical text from the supplied/current ADC state, using `ManagedDescriptionPropertySafeText` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   authoritativeDescription (Input; String): Authoritative description text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
' Returns: String: Returns the description text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function ManagedDescriptionExpectedPhysicalText(ByVal authoritativeDescription As String) As String
    ' Wrapper purpose: REVIEW facade that returns only the expected physical description text from the generalized managed-description resolver.
    ' Why retained: Retained to hide diagnostic ByRef outputs from callers that need only the text and avoid forcing dummy arguments into the simpler contract.
    Dim wasBounded As Boolean
    Dim authoritativeCharacterLength As Long
    Dim authoritativeByteLength As Long
    Dim writtenCharacterLength As Long
    Dim writtenByteLength As Long
    ManagedDescriptionExpectedPhysicalText = ManagedDescriptionPropertySafeText(authoritativeDescription, wasBounded, authoritativeCharacterLength, authoritativeByteLength, writtenCharacterLength, writtenByteLength)
End Function

Private Function ManagedDescriptionCachedTarget(ByVal targetObjects As Object, ByVal objectTypeCode As String, ByVal objectName As String) As Object
    Dim snapshotKey As String
    If targetObjects Is Nothing Then Exit Function
    snapshotKey = ManagedObjectDefinitionIdentityKey(objectName, objectTypeCode)
    If targetObjects.Exists(snapshotKey) = False Then Exit Function
    Set ManagedDescriptionCachedTarget = targetObjects(snapshotKey)
End Function

Private Function ManagedDescriptionCompactObjectStateFromTarget(ByVal objectTypeCode As String, ByVal objectName As String, ByVal targetObject As Object) As String
    Dim stateText As String
    objectTypeCode = UCase$(Trim$(objectTypeCode))
    If Len(objectName) = 0 Then Err.Raise vbObjectError + 6470, "ManagedDescriptionCompactObjectStateFromTarget", "An active managed object has a blank object name."
    Select Case objectTypeCode
        Case UCase$(ADCAT_OBJECT_TYPE_TABLE), UCase$(ADCAT_OBJECT_TYPE_QUERY), UCase$(ADCAT_OBJECT_TYPE_FORM), UCase$(ADCAT_OBJECT_TYPE_REPORT), UCase$(ADCAT_OBJECT_TYPE_MACRO)
        Case Else
            ManagedDescriptionCompactObjectStateFromTarget = FingerprintTokenPart("Unsupported")
            Exit Function
    End Select
    If targetObject Is Nothing Then
        ManagedDescriptionCompactObjectStateFromTarget = FingerprintTokenPart("Missing")
        Exit Function
    End If
    stateText = FingerprintTokenPart("Present")
    Select Case objectTypeCode
        Case UCase$(ADCAT_OBJECT_TYPE_TABLE)
            stateText = stateText & FingerprintTokenPart(CStr(targetObject.Name)) & FingerprintTokenPart(Format$(CDate(targetObject.DateCreated), "yyyy-mm-dd hh:nn:ss")) & FingerprintTokenPart(Format$(CDate(targetObject.LastUpdated), "yyyy-mm-dd hh:nn:ss")) & FingerprintTokenPart(CStr(targetObject.Attributes))
        Case UCase$(ADCAT_OBJECT_TYPE_QUERY)
            stateText = stateText & FingerprintTokenPart(CStr(targetObject.Name)) & FingerprintTokenPart(Format$(CDate(targetObject.DateCreated), "yyyy-mm-dd hh:nn:ss")) & FingerprintTokenPart(Format$(CDate(targetObject.LastUpdated), "yyyy-mm-dd hh:nn:ss")) & FingerprintTokenPart(CStr(targetObject.Type))
        Case Else
            stateText = stateText & FingerprintTokenPart(CStr(targetObject.Name)) & FingerprintTokenPart(Format$(CDate(targetObject.DateCreated), "yyyy-mm-dd hh:nn:ss")) & FingerprintTokenPart(Format$(CDate(targetObject.LastUpdated), "yyyy-mm-dd hh:nn:ss"))
    End Select
    ManagedDescriptionCompactObjectStateFromTarget = stateText
End Function

Private Function ManagedDescriptionCompactPhysicalStateFromCache(ByVal targetObjects As Object) As String
    On Error GoTo ErrHandler
    Dim catalogDb As DAO.Database
    Dim rs As DAO.Recordset
    Dim payloadText As String
    Dim objectName As String
    Dim objectTypeCode As String
    Dim targetObject As Object
    Dim checkedCount As Long
    Set catalogDb = ADCDb()
    payloadText = MANAGED_DESCRIPTION_PHYSICAL_STATE_VERSION & "|"
    Set rs = catalogDb.OpenRecordset("SELECT DISTINCT " & SQL_FIELD_OBJECT_TYPE & "," & SQL_FIELD_OBJECT_NAME & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID) & "] WHERE Nz([" & FIELD_MANAGED_OBJECT_STATUS & "]," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & " ORDER BY " & SQL_FIELD_OBJECT_TYPE & "," & SQL_FIELD_OBJECT_NAME & ";", dbOpenSnapshot)
    Do While rs.EOF = False
        PumpSetupResponsiveness
        checkedCount = checkedCount + 1
        objectTypeCode = CStr(Nz(rs.fields(FIELD_OBJECT_TYPE).value, vbNullString))
        objectName = CStr(Nz(rs.fields(FIELD_OBJECT_NAME).value, vbNullString))
        Set targetObject = ManagedDescriptionCachedTarget(targetObjects, objectTypeCode, objectName)
        payloadText = payloadText & FingerprintTokenPart(objectTypeCode) & FingerprintTokenPart(objectName) & ManagedDescriptionCompactObjectStateFromTarget(objectTypeCode, objectName, targetObject) & vbLf
        rs.MoveNext
    Loop
    rs.Close
    ManagedDescriptionCompactPhysicalStateFromCache = FingerprintFromPayload(payloadText)
    CorePulseProgress "Cached compact managed-object description state completed: objects=" & CStr(checkedCount) & "."
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
    Err.Raise savedErrNumber, "ManagedDescriptionCompactPhysicalStateFromCache", "Could not calculate the compact managed-object physical state from the retained G1 target cache." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function ManagedDescriptionCompactObjectState(ByVal db As DAO.Database, ByVal objectTypeCode As String, ByVal objectName As String) As String
    On Error GoTo ErrHandler
    Dim tdf As DAO.tableDef
    Dim qdf As DAO.queryDef
    Dim doc As DAO.Document
    Dim stateText As String
    Dim lookupErrNumber As Long
    Dim lookupErrSource As String
    Dim lookupErrDescription As String
    objectTypeCode = UCase$(Trim$(objectTypeCode))
    If Len(objectName) = 0 Then Err.Raise vbObjectError + 6445, "ManagedDescriptionCompactObjectState", "An active managed object has a blank object name."
    On Error Resume Next
    Select Case objectTypeCode
        Case UCase$(ADCAT_OBJECT_TYPE_TABLE)
            Set tdf = db.TableDefs(objectName)
        Case UCase$(ADCAT_OBJECT_TYPE_QUERY)
            Set qdf = db.QueryDefs(objectName)
        Case UCase$(ADCAT_OBJECT_TYPE_FORM)
            Set doc = db.Containers(FORM_DAO_CONTAINER_NAME).Documents(objectName)
        Case UCase$(ADCAT_OBJECT_TYPE_REPORT)
            Set doc = db.Containers(REPORT_DAO_CONTAINER_NAME).Documents(objectName)
        Case UCase$(ADCAT_OBJECT_TYPE_MACRO)
            Set doc = db.Containers(MACRO_DAO_CONTAINER_NAME).Documents(objectName)
        Case Else
            On Error GoTo ErrHandler
            ManagedDescriptionCompactObjectState = FingerprintTokenPart("Unsupported")
            Exit Function
    End Select
    lookupErrNumber = Err.Number
    lookupErrSource = Err.Source
    lookupErrDescription = Err.Description
    Err.Clear
    On Error GoTo ErrHandler
    If lookupErrNumber <> 0 Then
        If lookupErrNumber = 3265 Then ManagedDescriptionCompactObjectState = FingerprintTokenPart("Missing"): Exit Function
        Err.Raise lookupErrNumber, lookupErrSource, lookupErrDescription
    End If
    stateText = FingerprintTokenPart("Present")
    Select Case objectTypeCode
        Case UCase$(ADCAT_OBJECT_TYPE_TABLE)
            If tdf Is Nothing Then Err.Raise vbObjectError + 6446, "ManagedDescriptionCompactObjectState", "A keyed TableDef lookup returned Nothing without an error." & vbCrLf & "Table: " & objectName
            stateText = stateText & FingerprintTokenPart(tdf.Name) & FingerprintTokenPart(Format$(CDate(tdf.DateCreated), "yyyy-mm-dd hh:nn:ss")) & FingerprintTokenPart(Format$(CDate(tdf.LastUpdated), "yyyy-mm-dd hh:nn:ss")) & FingerprintTokenPart(CStr(tdf.Attributes))
        Case UCase$(ADCAT_OBJECT_TYPE_QUERY)
            If qdf Is Nothing Then Err.Raise vbObjectError + 6447, "ManagedDescriptionCompactObjectState", "A keyed QueryDef lookup returned Nothing without an error." & vbCrLf & "Query: " & objectName
            stateText = stateText & FingerprintTokenPart(qdf.Name) & FingerprintTokenPart(Format$(CDate(qdf.DateCreated), "yyyy-mm-dd hh:nn:ss")) & FingerprintTokenPart(Format$(CDate(qdf.LastUpdated), "yyyy-mm-dd hh:nn:ss")) & FingerprintTokenPart(CStr(qdf.Type))
        Case Else
            If doc Is Nothing Then Err.Raise vbObjectError + 6448, "ManagedDescriptionCompactObjectState", "A keyed DAO Document lookup returned Nothing without an error." & vbCrLf & "Object: " & objectTypeCode & " [" & objectName & "]"
            stateText = stateText & FingerprintTokenPart(doc.Name) & FingerprintTokenPart(Format$(CDate(doc.DateCreated), "yyyy-mm-dd hh:nn:ss")) & FingerprintTokenPart(Format$(CDate(doc.LastUpdated), "yyyy-mm-dd hh:nn:ss"))
    End Select
    ManagedDescriptionCompactObjectState = stateText
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ManagedDescriptionCompactObjectState", "Could not calculate keyed compact metadata for a managed object." & vbCrLf & "Object: " & objectTypeCode & " [" & objectName & "]" & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function ManagedDescriptionCompactPhysicalState(ByVal db As DAO.Database) As String
    On Error GoTo ErrHandler
    Dim catalogDb As DAO.Database
    Dim rs As DAO.Recordset
    Dim payloadText As String
    Dim objectName As String
    Dim objectTypeCode As String
    Dim objectState As String
    Dim identityKey As String
    Dim checkedCount As Long
    Dim captureBuildRepairHelpDocumentMacrosBaseline As Boolean
    Dim capturePolicyScopeStarted As Boolean
    Dim capturedNonHelpStates As Object
    Dim helpMacroIndex As Object
    Set catalogDb = ADCDb()
    captureBuildRepairHelpDocumentMacrosBaseline = mHelpDocumentMacroBuildRepairCompactStateCaptureActive
    If captureBuildRepairHelpDocumentMacrosBaseline Then
        BeginManagedDescriptionOptionalHelpPolicyScope
        capturePolicyScopeStarted = True
        Set helpMacroIndex = ManagedDescriptionOptionalHelpMacroIndex()
        Set capturedNonHelpStates = CreateObject("Scripting.Dictionary")
        capturedNonHelpStates.CompareMode = vbBinaryCompare
    End If
    payloadText = MANAGED_DESCRIPTION_PHYSICAL_STATE_VERSION & "|"
    Set rs = catalogDb.OpenRecordset("SELECT DISTINCT " & SQL_FIELD_OBJECT_TYPE & "," & SQL_FIELD_OBJECT_NAME & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID) & "] WHERE Nz([" & FIELD_MANAGED_OBJECT_STATUS & "]," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & " ORDER BY " & SQL_FIELD_OBJECT_TYPE & "," & SQL_FIELD_OBJECT_NAME & ";", dbOpenSnapshot)
    Do While rs.EOF = False
        checkedCount = checkedCount + 1
        objectTypeCode = CStr(Nz(rs.fields(FIELD_OBJECT_TYPE).value, vbNullString))
        objectName = CStr(Nz(rs.fields(FIELD_OBJECT_NAME).value, vbNullString))
        PulseProgressEvery checkedCount, 25, "Checking compact managed-object description state " & CStr(checkedCount) & ": " & objectTypeCode & " [" & objectName & "].", 0, "ReviewDescriptionCompactState", "Review"
        objectState = ManagedDescriptionCompactObjectState(db, objectTypeCode, objectName)
        payloadText = payloadText & FingerprintTokenPart(objectTypeCode) & FingerprintTokenPart(objectName) & objectState & vbLf
        If captureBuildRepairHelpDocumentMacrosBaseline Then
            identityKey = ManagedObjectDefinitionIdentityKey(objectName, objectTypeCode)
            If Len(identityKey) = 0 Then Err.Raise vbObjectError + 6806, "ManagedDescriptionCompactPhysicalState", "Could not derive a canonical identity while capturing the Build/Repair Help Document Macros compact-state baseline." & vbCrLf & "Object: " & objectTypeCode & " [" & objectName & "]"
            If helpMacroIndex.Exists(identityKey) = False Then capturedNonHelpStates.Add identityKey, objectState
        End If
        If checkedCount = 1 Or checkedCount Mod 25 = 0 Then RequireSafeWorkContinue "ManagedDescriptionCompactPhysicalState", objectTypeCode & " [" & objectName & "]"
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    If captureBuildRepairHelpDocumentMacrosBaseline Then
        If Len(GetActiveTraceRunID()) = 0 Then Err.Raise vbObjectError + 6807, "ManagedDescriptionCompactPhysicalState", "A trace-run identity is required while capturing the Build/Repair Help Document Macros compact-state baseline."
        Set mHelpDocumentMacroBuildRepairBaselineNonHelpCompactStateByObject = capturedNonHelpStates
        mHelpDocumentMacroBuildRepairBaselineCompactStateTraceRunID = GetActiveTraceRunID()
    End If
    ManagedDescriptionCompactPhysicalState = FingerprintFromPayload(payloadText)
    If capturePolicyScopeStarted Then
        EndManagedDescriptionOptionalHelpPolicyScope
        capturePolicyScopeStarted = False
    End If
    CorePulseProgress "Compact managed-object description state completed: objects=" & CStr(checkedCount) & "."
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
    If capturePolicyScopeStarted Then EndManagedDescriptionOptionalHelpPolicyScope
    If captureBuildRepairHelpDocumentMacrosBaseline Then
        Set mHelpDocumentMacroBuildRepairBaselineNonHelpCompactStateByObject = Nothing
        mHelpDocumentMacroBuildRepairBaselineCompactStateTraceRunID = vbNullString
    End If
    On Error GoTo 0
    Err.Raise savedErrNumber, "ManagedDescriptionCompactPhysicalState", "Could not calculate the compact managed-object physical state." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function ManagedDescriptionCompactPhysicalStateAfterScopedHelpDocumentMacroBuildRepair(ByVal db As DAO.Database, ByRef detailText As String) As String
    ' Rebuild the canonical compact-state payload from the pre-mutation non-Help cache plus fresh Help-macro metadata. Build/Repair Help Document Macros owns only the Help macro mutations, so rereading unchanged
    ' non-Help objects would duplicate the proof already established at the operation boundary.
    On Error GoTo ErrHandler
    Dim catalogDb As DAO.Database
    Dim rs As DAO.Recordset
    Dim payloadText As String
    Dim objectName As String
    Dim objectTypeCode As String
    Dim objectState As String
    Dim identityKey As String
    Dim baselineKey As Variant
    Dim activeTraceRunID As String
    Dim checkedCount As Long
    Dim reusedNonHelpCount As Long
    Dim refreshedHelpCount As Long
    Dim policyScopeStarted As Boolean
    Dim helpMacroIndex As Object
    Dim currentNonHelpKeys As Object
    detailText = vbNullString
    activeTraceRunID = GetActiveTraceRunID()
    If Len(activeTraceRunID) = 0 Then Err.Raise vbObjectError + 6808, "ManagedDescriptionCompactPhysicalStateAfterScopedHelpDocumentMacroBuildRepair", "No active trace run exists."
    If mHelpDocumentMacroBuildRepairBaselineNonHelpCompactStateByObject Is Nothing Then Err.Raise vbObjectError + 6809, "ManagedDescriptionCompactPhysicalStateAfterScopedHelpDocumentMacroBuildRepair", "The pre-mutation non-Help compact-state cache is unavailable."
    If StrComp(mHelpDocumentMacroBuildRepairBaselineCompactStateTraceRunID, activeTraceRunID, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 6810, "ManagedDescriptionCompactPhysicalStateAfterScopedHelpDocumentMacroBuildRepair", "The pre-mutation non-Help compact-state cache belongs to a different trace run."
    BeginManagedDescriptionOptionalHelpPolicyScope
    policyScopeStarted = True
    Set helpMacroIndex = ManagedDescriptionOptionalHelpMacroIndex()
    Set currentNonHelpKeys = CreateObject("Scripting.Dictionary")
    currentNonHelpKeys.CompareMode = vbBinaryCompare
    Set catalogDb = ADCDb()
    payloadText = MANAGED_DESCRIPTION_PHYSICAL_STATE_VERSION & "|"
    Set rs = catalogDb.OpenRecordset("SELECT DISTINCT " & SQL_FIELD_OBJECT_TYPE & "," & SQL_FIELD_OBJECT_NAME & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID) & "] WHERE Nz([" & FIELD_MANAGED_OBJECT_STATUS & "]," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & " ORDER BY " & SQL_FIELD_OBJECT_TYPE & "," & SQL_FIELD_OBJECT_NAME & ";", dbOpenSnapshot)
    Do While rs.EOF = False
        checkedCount = checkedCount + 1
        objectTypeCode = CStr(Nz(rs.fields(FIELD_OBJECT_TYPE).value, vbNullString))
        objectName = CStr(Nz(rs.fields(FIELD_OBJECT_NAME).value, vbNullString))
        identityKey = ManagedObjectDefinitionIdentityKey(objectName, objectTypeCode)
        If Len(identityKey) = 0 Then Err.Raise vbObjectError + 6811, "ManagedDescriptionCompactPhysicalStateAfterScopedHelpDocumentMacroBuildRepair", "Could not derive a canonical identity while advancing the Build/Repair Help Document Macros compact-state proof." & vbCrLf & "Object: " & objectTypeCode & " [" & objectName & "]"
        If helpMacroIndex.Exists(identityKey) Then
            objectState = ManagedDescriptionCompactObjectState(db, objectTypeCode, objectName)
            refreshedHelpCount = refreshedHelpCount + 1
        Else
            If mHelpDocumentMacroBuildRepairBaselineNonHelpCompactStateByObject.Exists(identityKey) = False Then Err.Raise vbObjectError + 6812, "ManagedDescriptionCompactPhysicalStateAfterScopedHelpDocumentMacroBuildRepair", "A non-Help managed-object identity appeared after the pre-mutation proof; Build/Repair Help Document Macros will not bless unrelated state changes." & vbCrLf & "Object: " & objectTypeCode & " [" & objectName & "]"
            objectState = CStr(mHelpDocumentMacroBuildRepairBaselineNonHelpCompactStateByObject(identityKey))
            currentNonHelpKeys.Add identityKey, True
            reusedNonHelpCount = reusedNonHelpCount + 1
        End If
        payloadText = payloadText & FingerprintTokenPart(objectTypeCode) & FingerprintTokenPart(objectName) & objectState & vbLf
        If checkedCount = 1 Or checkedCount Mod 25 = 0 Then RequireSafeWorkContinue "ManagedDescriptionCompactPhysicalStateAfterScopedHelpDocumentMacroBuildRepair", objectTypeCode & " [" & objectName & "]"
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    If currentNonHelpKeys.count <> mHelpDocumentMacroBuildRepairBaselineNonHelpCompactStateByObject.count Then
        For Each baselineKey In mHelpDocumentMacroBuildRepairBaselineNonHelpCompactStateByObject.keys
            If currentNonHelpKeys.Exists(CStr(baselineKey)) = False Then Err.Raise vbObjectError + 6813, "ManagedDescriptionCompactPhysicalStateAfterScopedHelpDocumentMacroBuildRepair", "A pre-mutation non-Help managed-object identity is no longer active after Help reconciliation; Build/Repair Help Document Macros will not bless unrelated state changes." & vbCrLf & "Identity: " & CStr(baselineKey)
        Next baselineKey
        Err.Raise vbObjectError + 6814, "ManagedDescriptionCompactPhysicalStateAfterScopedHelpDocumentMacroBuildRepair", "The non-Help managed-object identity count changed during Build/Repair Help Document Macros."
    End If
    ManagedDescriptionCompactPhysicalStateAfterScopedHelpDocumentMacroBuildRepair = FingerprintFromPayload(payloadText)
    detailText = "Compact managed-object physical state advanced from the operation-scoped pre-mutation cache plus current Help-macro metadata; non-Help states reused=" & CStr(reusedNonHelpCount) & "; Help states reread=" & CStr(refreshedHelpCount) & "; full post-Help compact scan=False."
    EndManagedDescriptionOptionalHelpPolicyScope
    policyScopeStarted = False
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
    If policyScopeStarted Then EndManagedDescriptionOptionalHelpPolicyScope
    On Error GoTo 0
    detailText = "Could not advance the compact managed-object physical state after scoped Help reconciliation. Error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription
    ManagedDescriptionCompactPhysicalStateAfterScopedHelpDocumentMacroBuildRepair = vbNullString
End Function

Private Sub InvalidateHelpDocumentMacroBuildRepairBaselineCompactStateCache()
    ' Explicit lifetime boundary for the Build/Repair Help Document Macros pre-mutation compact-state cache.
    mHelpDocumentMacroBuildRepairCompactStateCaptureActive = False
    Set mHelpDocumentMacroBuildRepairBaselineNonHelpCompactStateByObject = Nothing
    mHelpDocumentMacroBuildRepairBaselineCompactStateTraceRunID = vbNullString
End Sub

Private Sub BeginManagedDescriptionOptionalHelpPolicyScope()
    'The first owner builds a fresh operation-scoped keyed projection; nested description work reuses it until the outer operation ends.
    If mManagedDescriptionOptionalHelpPolicyDepth = 0 Then Set mManagedDescriptionOptionalHelpMacroIndex = Nothing
    mManagedDescriptionOptionalHelpPolicyDepth = mManagedDescriptionOptionalHelpPolicyDepth + 1
End Sub

Private Sub EndManagedDescriptionOptionalHelpPolicyScope()
    If mManagedDescriptionOptionalHelpPolicyDepth > 0 Then mManagedDescriptionOptionalHelpPolicyDepth = mManagedDescriptionOptionalHelpPolicyDepth - 1
    If mManagedDescriptionOptionalHelpPolicyDepth = 0 Then Set mManagedDescriptionOptionalHelpMacroIndex = Nothing
End Sub

Private Function ManagedDescriptionOptionalHelpMacroIndex() As Object
    'One operation-scoped keyed projection over Catalog-owned action payloads. Run_OpenHelpDocument(...) identifies Help-document macros without relying on macro numbers or titles; the stored value is the B4-derived
    'Help_Docs.ObjectDescription.
    ' INTENTIONAL MANAGED-OBJECT EXCEPTION: a Help Document macro may have a blank authoritative ObjectDescription. G1 and scoped Help repair treat that blank as valid optional metadata, not as a missing-
    ' description defect.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim objectName As String
    Dim objectTypeCode As String
    Dim runCodeWrapperName As String
    Dim macroDescription As String
    Dim identityKey As String
    If Not mManagedDescriptionOptionalHelpMacroIndex Is Nothing Then
        Set ManagedDescriptionOptionalHelpMacroIndex = mManagedDescriptionOptionalHelpMacroIndex
        Exit Function
    End If
    Set mManagedDescriptionOptionalHelpMacroIndex = CreateObject("Scripting.Dictionary")
    mManagedDescriptionOptionalHelpMacroIndex.CompareMode = vbTextCompare
    Set db = ADCDb()
    sqlText = "SELECT m." & SQL_FIELD_ID & ",m." & SQL_FIELD_OBJECT_NAME & ",m." & SQL_FIELD_OBJECT_TYPE & ",p.[" & FIELD_PAYLOAD_KIND & "],p.[" & FIELD_RUN_CODE_WRAPPER_NAME & "],p.[" & FIELD_MACRO_DESCRIPTION & "] " & _
              "FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID) & "] AS m INNER JOIN [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectActionPayloads") & "] AS p ON m." & SQL_FIELD_ID & "=p.[" & FIELD_MANAGED_OBJECT_ID & "] " & _
              "WHERE Nz(m.[" & FIELD_MANAGED_OBJECT_STATUS & "]," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & " AND Nz(p.[" & FIELD_PAYLOAD_STATUS & "]," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & " AND UCase(Trim(Nz(m." & SQL_FIELD_OBJECT_TYPE & ",'')))=" & SharedDao_SqlQuote(UCase$(ADCAT_OBJECT_TYPE_MACRO)) & " AND UCase(Trim(Nz(m.[" & FIELD_CATEGORY_CODE & "],'')))=" & SharedDao_SqlQuote(UCase$(ADCAT_CATEGORY_SYSTEM_MACRO)) & " AND UCase(Trim(Nz(m.[" & FIELD_CATALOG_OWNER & "],'')))=" & SharedDao_SqlQuote(UCase$(ADCAT_OWNER_SYSTEM_DEFAULT)) & " ORDER BY m." & SQL_FIELD_ID & ",p." & SQL_FIELD_ID & ";"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    Do While rs.EOF = False
        objectName = Trim$(CStr(Nz(rs.fields(FIELD_OBJECT_NAME).value, vbNullString)))
        objectTypeCode = Trim$(CStr(rs.fields(FIELD_OBJECT_TYPE).value))
        runCodeWrapperName = Trim$(CStr(Nz(rs.fields(FIELD_RUN_CODE_WRAPPER_NAME).value, vbNullString)))
        macroDescription = CStr(Nz(rs.fields(FIELD_MACRO_DESCRIPTION).value, vbNullString))
        If StrComp(Trim$(CStr(Nz(rs.fields(FIELD_PAYLOAD_KIND).value, vbNullString))), ADCAT_PAYLOAD_ACCESS_MACRO_RUNCODE, vbTextCompare) = 0 Then
            If StrComp(Left$(runCodeWrapperName, Len("Run_OpenHelpDocument(")), "Run_OpenHelpDocument(", vbTextCompare) = 0 Then
                identityKey = ManagedObjectDefinitionIdentityKey(objectName, objectTypeCode)
                If Len(identityKey) > 0 Then
                    If mManagedDescriptionOptionalHelpMacroIndex.Exists(identityKey) Then
                        If StrComp(CStr(mManagedDescriptionOptionalHelpMacroIndex(identityKey)), macroDescription, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 6480, "ManagedDescriptionOptionalHelpMacroIndex", "Two active Help-document payload rows disagree on the authoritative optional Description." & vbCrLf & "Object: " & objectName
                    Else
                        mManagedDescriptionOptionalHelpMacroIndex.Add identityKey, macroDescription
                    End If
                End If
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Set ManagedDescriptionOptionalHelpMacroIndex = mManagedDescriptionOptionalHelpMacroIndex
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
    Set mManagedDescriptionOptionalHelpMacroIndex = Nothing
    On Error GoTo 0
    Err.Raise savedErrNumber, "ManagedDescriptionOptionalHelpMacroIndex", "Could not resolve the active Help-document macro Description policy from Catalog-owned action payloads." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function ManagedDescriptionOptionalHelpMacroDescription(ByVal objectName As String, ByVal objectTypeCode As String, ByRef authoritativeDescription As String) As Boolean
    Dim identityKey As String
    Dim optionalHelpIndex As Object
    authoritativeDescription = vbNullString
    identityKey = ManagedObjectDefinitionIdentityKey(objectName, objectTypeCode)
    If Len(identityKey) = 0 Then Exit Function
    Set optionalHelpIndex = ManagedDescriptionOptionalHelpMacroIndex()
    If optionalHelpIndex.Exists(identityKey) = False Then Exit Function
    authoritativeDescription = CStr(optionalHelpIndex(identityKey))
    ManagedDescriptionOptionalHelpMacroDescription = True
End Function

Private Function ManagedDescriptionIsOptionalHelpMacro(ByVal objectName As String, ByVal objectTypeCode As String) As Boolean
    ' Wrapper purpose: Private Boolean facade over the generalized optional Help Macro description resolver when only the classification result is needed.
    ' Why retained: Retained so callers do not create and discard an otherwise-unused description output argument at every predicate call site.
    Dim authoritativeDescription As String
    ManagedDescriptionIsOptionalHelpMacro = ManagedDescriptionOptionalHelpMacroDescription(objectName, objectTypeCode, authoritativeDescription)
End Function

Private Function ManagedDescriptionInputFingerprint(ByVal db As DAO.Database, Optional ByVal formsOnly As Boolean = False) As String
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim payloadText As String
    Dim objectName As String
    Dim objectTypeCode As String
    Dim authoritativeHelpDescription As String
    Dim optionalHelpDescription As Boolean
    Dim checkedCount As Long
    BeginManagedDescriptionOptionalHelpPolicyScope
    sqlText = "SELECT " & SQL_FIELD_OBJECT_NAME & "," & SQL_FIELD_OBJECT_TYPE & "," & SQL_FIELD_OBJECT_DESCRIPTION & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID) & "] WHERE Nz([" & FIELD_MANAGED_OBJECT_STATUS & "]," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE)
    If formsOnly Then sqlText = sqlText & " AND UCase(Trim(Nz(" & SQL_FIELD_OBJECT_TYPE & ",'')))=" & SharedDao_SqlQuote(UCase$(ADCAT_OBJECT_TYPE_FORM))
    sqlText = sqlText & " ORDER BY " & OrderByPolicyText("REVIEW_MANAGED_DESCRIPTION_UNALIASED") & ";"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    payloadText = MANAGED_DESCRIPTION_FINGERPRINT_VERSION & "|"
    Do While rs.EOF = False
        checkedCount = checkedCount + 1
        If checkedCount = 1 Or checkedCount Mod 50 = 0 Then RequireSafeWorkContinue "ManagedDescriptionInputFingerprint", "fingerprinted " & CStr(checkedCount) & " managed description rows"
        objectName = CStr(Nz(rs.fields(FIELD_OBJECT_NAME).value, vbNullString))
        objectTypeCode = CStr(Nz(rs.fields(FIELD_OBJECT_TYPE).value, vbNullString))
        payloadText = payloadText & FingerprintTokenPart(objectName)
        payloadText = payloadText & FingerprintTokenPart(objectTypeCode)
        authoritativeHelpDescription = vbNullString
        optionalHelpDescription = ManagedDescriptionOptionalHelpMacroDescription(objectName, objectTypeCode, authoritativeHelpDescription)
        payloadText = payloadText & FingerprintTokenPart(CStr(Nz(rs.fields(FIELD_OBJECT_DESCRIPTION).value, vbNullString)))
        payloadText = payloadText & FingerprintTokenPart(CStr(optionalHelpDescription))
        payloadText = payloadText & FingerprintTokenPart(IIf(optionalHelpDescription, authoritativeHelpDescription, vbNullString)) & vbLf
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    ManagedDescriptionInputFingerprint = FingerprintFromPayload(payloadText)
    EndManagedDescriptionOptionalHelpPolicyScope
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
    EndManagedDescriptionOptionalHelpPolicyScope
    On Error GoTo 0
    Err.Raise savedErrNumber, "ManagedDescriptionInputFingerprint", "Could not fingerprint managed-object description inputs." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function ManagedDescriptionPhysicalSnapshot(ByVal db As DAO.Database, ByVal targetRows As Collection, ByRef targetObjectCache As Object) As Object
    'Builds an exact Description snapshot only for authoritative managed targets; unmanaged and Access-generated objects are outside this proof.
    On Error GoTo ErrHandler
    Dim snapshot As Object
    Dim tdf As DAO.tableDef
    Dim qdf As DAO.queryDef
    Dim doc As DAO.Document
    Dim rowData As Variant
    Dim objectName As String
    Dim objectTypeCode As String
    Dim lookupErrNumber As Long
    Dim lookupErrSource As String
    Dim lookupErrDescription As String
    Dim checkedCount As Long
    ' INTENTIONAL OPERATION SNAPSHOT: snapshot captures one coherent source state for this invocation of ManagedDescriptionPhysicalSnapshot; validity ends with the operation rather than becoming a live data
    ' authority.
    ' INTENTIONAL KEYED LOOKUP: targetObjectCache provides keyed membership/lookup over source rows or state already read by ManagedDescriptionPhysicalSnapshot; it is a derived access path, not a second authority.
    Set snapshot = CreateObject("Scripting.Dictionary")
    snapshot.CompareMode = vbTextCompare
    Set targetObjectCache = CreateObject("Scripting.Dictionary")
    targetObjectCache.CompareMode = vbTextCompare
    If targetRows Is Nothing Then Err.Raise vbObjectError + 6467, "ManagedDescriptionPhysicalSnapshot", "The authoritative managed-description target collection is missing."
    For Each rowData In targetRows
        checkedCount = checkedCount + 1
        objectName = CStr(rowData(0))
        objectTypeCode = UCase$(Trim$(CStr(rowData(1))))
        PulseProgressEvery checkedCount, 50, "Snapshotting managed " & objectTypeCode & " description " & CStr(checkedCount) & " of " & CStr(targetRows.count) & ": [" & objectName & "].", targetRows.count, "ReviewDescriptionPhysicalSnapshot", "Review"
        Set tdf = Nothing
        Set qdf = Nothing
        Set doc = Nothing
        lookupErrNumber = 0
        lookupErrSource = vbNullString
        lookupErrDescription = vbNullString
        On Error Resume Next
        Select Case objectTypeCode
            Case UCase$(ADCAT_OBJECT_TYPE_TABLE)
                Set tdf = db.TableDefs(objectName)
            Case UCase$(ADCAT_OBJECT_TYPE_QUERY)
                Set qdf = db.QueryDefs(objectName)
            Case UCase$(ADCAT_OBJECT_TYPE_FORM)
                Set doc = db.Containers(FORM_DAO_CONTAINER_NAME).Documents(objectName)
            Case UCase$(ADCAT_OBJECT_TYPE_REPORT)
                Set doc = db.Containers(REPORT_DAO_CONTAINER_NAME).Documents(objectName)
            Case UCase$(ADCAT_OBJECT_TYPE_MACRO)
                Set doc = db.Containers(MACRO_DAO_CONTAINER_NAME).Documents(objectName)
        End Select
        lookupErrNumber = Err.Number
        lookupErrSource = Err.Source
        lookupErrDescription = Err.Description
        Err.Clear
        On Error GoTo ErrHandler
        If lookupErrNumber <> 0 And lookupErrNumber <> 3265 Then Err.Raise lookupErrNumber, lookupErrSource, lookupErrDescription
        If lookupErrNumber = 0 Then
            Select Case objectTypeCode
                Case UCase$(ADCAT_OBJECT_TYPE_TABLE)
                    If Not tdf Is Nothing Then
                        ManagedDescriptionSnapshotAdd snapshot, targetObjectCache, ADCAT_OBJECT_TYPE_TABLE, objectName, tdf
                    End If
                Case UCase$(ADCAT_OBJECT_TYPE_QUERY)
                    If Not qdf Is Nothing Then ManagedDescriptionSnapshotAdd snapshot, targetObjectCache, ADCAT_OBJECT_TYPE_QUERY, objectName, qdf
                Case UCase$(ADCAT_OBJECT_TYPE_FORM)
                    If Not doc Is Nothing Then ManagedDescriptionSnapshotAdd snapshot, targetObjectCache, ADCAT_OBJECT_TYPE_FORM, objectName, doc
                Case UCase$(ADCAT_OBJECT_TYPE_REPORT)
                    If Not doc Is Nothing Then ManagedDescriptionSnapshotAdd snapshot, targetObjectCache, ADCAT_OBJECT_TYPE_REPORT, objectName, doc
                Case UCase$(ADCAT_OBJECT_TYPE_MACRO)
                    If Not doc Is Nothing Then ManagedDescriptionSnapshotAdd snapshot, targetObjectCache, ADCAT_OBJECT_TYPE_MACRO, objectName, doc
            End Select
        End If
        If checkedCount = 1 Or checkedCount Mod 25 = 0 Or checkedCount = targetRows.count Then RequireSafeWorkContinue "ManagedDescriptionPhysicalSnapshot", objectTypeCode & " [" & objectName & "]"
    Next rowData
    CorePulseProgress "Managed-description physical snapshot completed: authoritative targets=" & CStr(checkedCount) & "."
    Set ManagedDescriptionPhysicalSnapshot = snapshot
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ManagedDescriptionPhysicalSnapshot", "Could not collect the exact physical managed-object Description snapshot." & vbCrLf & "Object: " & objectTypeCode & " [" & objectName & "]" & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function ManagedDescriptionPropertyText(ByVal targetObject As Object) As String
    On Error GoTo MissingProperty
    ManagedDescriptionPropertyText = CStr(Nz(targetObject.properties(DAO_DESCRIPTION_PROPERTY_NAME).value, vbNullString))
    Exit Function
MissingProperty:
    Err.Clear
End Function

Private Sub ManagedDescriptionSnapshotAdd(ByRef snapshot As Object, ByVal targetObjectCache As Object, ByVal objectTypeCode As String, ByVal objectName As String, ByVal targetObject As Object)
    Dim snapshotKey As String
    snapshotKey = ManagedObjectDefinitionIdentityKey(objectName, objectTypeCode)
    If Len(snapshotKey) = 0 Then Exit Sub
    If snapshot.Exists(snapshotKey) = False Then snapshot.Add snapshotKey, ManagedDescriptionPropertyText(targetObject)
    If Not targetObjectCache Is Nothing Then
        If targetObjectCache.Exists(snapshotKey) = False Then targetObjectCache.Add snapshotKey, targetObject
    End If
End Sub

Private Function ManagedDescriptionSnapshotContains(ByVal snapshot As Object, ByVal objectTypeCode As String, ByVal objectName As String) As Boolean
    If snapshot Is Nothing Then Exit Function
    ManagedDescriptionSnapshotContains = snapshot.Exists(ManagedObjectDefinitionIdentityKey(objectName, objectTypeCode))
End Function

Private Function ManagedDescriptionSnapshotMatches(ByVal snapshot As Object, ByVal objectTypeCode As String, ByVal objectName As String, ByVal expectedDescription As String) As Boolean
    Dim snapshotKey As String
    Dim expectedPhysicalDescription As String
    If snapshot Is Nothing Then Exit Function
    snapshotKey = ManagedObjectDefinitionIdentityKey(objectName, objectTypeCode)
    If snapshot.Exists(snapshotKey) = False Then Exit Function
    expectedPhysicalDescription = ManagedDescriptionExpectedPhysicalText(expectedDescription)
    ManagedDescriptionSnapshotMatches = (StrComp(CStr(snapshot(snapshotKey)), expectedPhysicalDescription, vbBinaryCompare) = 0)
End Function

Private Function ManagedObjectDescriptionIsPlaceholder(ByVal descriptionText As String) As Boolean
    'Intended use: Single managed-object Description placeholder policy for repair, validation, and Health Review.
    Dim normalizedText As String
    normalizedText = UCase$(Trim$(descriptionText))
    If Len(normalizedText) = 0 Then ManagedObjectDescriptionIsPlaceholder = True: Exit Function
    If Left$(normalizedText, 36) = "UNMAPPED MANAGED OBJECT DESCRIPTION:" Then ManagedObjectDescriptionIsPlaceholder = True: Exit Function
    If Left$(normalizedText, 29) = "DEVELOPMENT METADATA MISSING" Then ManagedObjectDescriptionIsPlaceholder = True: Exit Function
    If Left$(normalizedText, 12) = "PLACEHOLDER:" Then ManagedObjectDescriptionIsPlaceholder = True: Exit Function
    If Left$(normalizedText, 5) = "TODO:" Or normalizedText = "TODO" Then ManagedObjectDescriptionIsPlaceholder = True: Exit Function
    If Left$(normalizedText, 4) = "TBD:" Or normalizedText = "TBD" Then ManagedObjectDescriptionIsPlaceholder = True
End Function

Private Sub ApplyManagedObjectDescriptionsFromCatalog(ByVal descriptionScope As ManagedObjectDescriptionScope, Optional ByRef physicalChangeCount As Variant, Optional ByRef formPhysicalChangeCount As Variant, Optional ByRef scopedExactToken As Variant)
    'Loads the authoritative managed-description rows once, then conditionally writes and verifies each physical Description in one pass. Active Help-document macro descriptions are optional; supplied values remain
    'authoritative.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim descriptionRows As Collection
    Dim rowData As Variant
    Dim sqlText As String
    Dim objectName As String
    Dim objectTypeCode As String
    Dim descriptionText As String
    Dim physicalDescriptionText As String
    Dim descriptionWasBounded As Boolean
    Dim authoritativeCharacterLength As Long
    Dim authoritativeByteLength As Long
    Dim writtenCharacterLength As Long
    Dim writtenByteLength As Long
    Dim boundedDescriptionCount As Long
    Dim boundedDescriptionDetails As String
    Dim expectedExistencePolicy As String
    Dim setupAction As String
    Dim optionalReason As String
    Dim changedCount As Long
    Dim unchangedCount As Long
    Dim skippedCount As Long
    Dim repairedBlankDescriptionCount As Long
    Dim optionalAbsentCount As Long
    Dim optionalAbsentDetails As String
    Dim missingOrUnsupportedCount As Long
    Dim missingOrUnsupportedDetails As String
    Dim descriptionResult As Long
    Dim checkedCount As Long
    Dim snapshotLoadCount As Long
    Dim inputFingerprint As String
    Dim postCommitPayload As String
    Dim activeTraceRunID As String
    Dim fullCatalogScope As Boolean
    Dim resultDetail As String
    Dim descriptionOptionalForHelpMacro As Boolean
    Dim helpDocumentMacrosOnly As Boolean
    Select Case descriptionScope
        Case MANAGED_OBJECT_DESCRIPTION_SCOPE_ALL
            helpDocumentMacrosOnly = False
        Case MANAGED_OBJECT_DESCRIPTION_SCOPE_HELP_DOCUMENT_MACROS_ONLY
            helpDocumentMacrosOnly = True
        Case Else
            Err.Raise vbObjectError + 6448, "ApplyManagedObjectDescriptionsFromCatalog", "Unsupported managed-object description scope: " & CStr(descriptionScope)
    End Select
    BeginManagedDescriptionOptionalHelpPolicyScope
    physicalChangeCount = CLng(0)
    formPhysicalChangeCount = CLng(0)
    fullCatalogScope = Not helpDocumentMacrosOnly
    RequireManagedObjectCatalogReadyForDescriptionWork "ApplyManagedObjectDescriptionsFromCatalog"
    Set db = ADCDb()
    activeTraceRunID = GetActiveTraceRunID()
    sqlText = "SELECT m." & SQL_FIELD_ID & ",m." & SQL_FIELD_OBJECT_NAME & ",m." & SQL_FIELD_OBJECT_TYPE & ",m.[" & FIELD_CATEGORY_CODE & "],m." & SQL_FIELD_OBJECT_DESCRIPTION & ","
    sqlText = sqlText & ResolvedExpectedExistencePolicySql("o", "c") & " " & "AS " & SQL_FIELD_RESOLVED_EXPECTED_EXISTENCE_POLICY & ","
    sqlText = sqlText & ResolvedSetupActionSql("o", "c") & " " & "AS " & SQL_FIELD_RESOLVED_SETUP_ACTION & " "
    sqlText = sqlText & "FROM ([" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID) & "] AS m INNER JOIN [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ObjectCategories") & "] AS c ON m.[" & FIELD_CATEGORY_CODE & "]=c.[" & FIELD_CATEGORY_CODE & "]) LEFT JOIN [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECT_OVERRIDES_TABLE_LOGICAL_ID) & "] AS o ON m." & SQL_FIELD_ID & "=o.[" & FIELD_MANAGED_OBJECT_ID & "] "
    sqlText = sqlText & "WHERE Nz(m.[" & FIELD_MANAGED_OBJECT_STATUS & "]," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & " AND Nz(c." & SQL_FIELD_OBJECT_CATEGORY_STATUS & "," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE)
    sqlText = sqlText & " ORDER BY " & OrderByPolicyText("REVIEW_MANAGED_DESCRIPTION") & ";"
    ' INTENTIONAL OPERATION SNAPSHOT: descriptionRows captures one coherent source state for this invocation of ApplyManagedObjectDescriptionsFromCatalog; validity ends with the operation rather than becoming a live
    ' data authority.
    Set descriptionRows = New Collection
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    Do While rs.EOF = False
        objectName = CStr(Nz(rs.fields(FIELD_OBJECT_NAME).value, vbNullString))
        objectTypeCode = CStr(Nz(rs.fields(FIELD_OBJECT_TYPE).value, vbNullString))
        If helpDocumentMacrosOnly = False Or ManagedDescriptionIsOptionalHelpMacro(objectName, objectTypeCode) Then
            snapshotLoadCount = snapshotLoadCount + 1
            If snapshotLoadCount = 1 Or snapshotLoadCount Mod 50 = 0 Then RequireSafeWorkContinue "ApplyManagedObjectDescriptionsFromCatalog", "loaded " & CStr(snapshotLoadCount) & " description source rows"
            descriptionRows.Add Array(objectName, objectTypeCode, CStr(Nz(rs.fields(FIELD_RESOLVED_EXPECTED_EXISTENCE_POLICY).value, vbNullString)), CStr(Nz(rs.fields(FIELD_RESOLVED_SETUP_ACTION).value, vbNullString)), CStr(Nz(rs.fields(FIELD_OBJECT_DESCRIPTION).value, vbNullString)))
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    If helpDocumentMacrosOnly Then
        inputFingerprint = HELP_DOCUMENT_DESCRIPTION_PROOF_VERSION
        postCommitPayload = HELP_DOCUMENT_DESCRIPTION_PROOF_VERSION & "|"
    Else
        inputFingerprint = ManagedDescriptionInputFingerprint(db, False)
        postCommitPayload = "ADCManagedObjectDescriptionPostCommitV1|" & inputFingerprint & "|"
    End If
    CorePulseProgress "Applying and verifying physical managed-object descriptions."
    For Each rowData In descriptionRows
        checkedCount = checkedCount + 1
        If checkedCount = 1 Or checkedCount Mod 25 = 0 Or checkedCount = descriptionRows.count Then RequireSafeWorkContinue "ApplyManagedObjectDescriptionsFromCatalog", "applying description " & CStr(checkedCount) & " of " & CStr(descriptionRows.count)
        objectName = CStr(rowData(0))
        objectTypeCode = CStr(rowData(1))
        expectedExistencePolicy = CStr(rowData(2))
        setupAction = CStr(rowData(3))
        descriptionText = CStr(rowData(4))
        descriptionOptionalForHelpMacro = ManagedDescriptionIsOptionalHelpMacro(objectName, objectTypeCode)
        If Len(Trim$(descriptionText)) = 0 And descriptionOptionalForHelpMacro = False Then Err.Raise vbObjectError + 6442, "ApplyManagedObjectDescriptionsFromCatalog", "An active Managed Object Catalog description is blank outside the optional Help-document macro policy." & vbCrLf & "Object: " & objectTypeCode & " [" & objectName & "]"
        physicalDescriptionText = ManagedDescriptionPropertySafeText(descriptionText, descriptionWasBounded, authoritativeCharacterLength, authoritativeByteLength, writtenCharacterLength, writtenByteLength)
        If descriptionWasBounded Then
            boundedDescriptionCount = boundedDescriptionCount + 1
            boundedDescriptionDetails = SharedText_AppendNonBlankLine(boundedDescriptionDetails, objectTypeCode & " [" & objectName & "] | Property=Description | Authoritative characters=" & CStr(authoritativeCharacterLength) & " | Authoritative bytes=" & CStr(authoritativeByteLength) & " | Written characters=" & CStr(writtenCharacterLength) & " | Written bytes=" & CStr(writtenByteLength) & " | Fingerprint=" & StableTextHashHex(descriptionText))
        End If
        postCommitPayload = postCommitPayload & FingerprintTokenPart(objectTypeCode) & FingerprintTokenPart(objectName) & FingerprintTokenPart(descriptionText) & FingerprintTokenPart(physicalDescriptionText)
        PulseProgressEvery checkedCount, 25, "Applying object description: " & objectTypeCode & " [" & objectName & "]"
        descriptionResult = SetAccessObjectDescriptionByCatalogType(db, objectName, objectTypeCode, descriptionText, descriptionOptionalForHelpMacro)
        Select Case descriptionResult
            Case 1
                postCommitPayload = postCommitPayload & "Present|"
                changedCount = changedCount + 1
                If StrComp(Trim$(objectTypeCode), ADCAT_OBJECT_TYPE_FORM, vbTextCompare) = 0 Then formPhysicalChangeCount = CLng(Nz(formPhysicalChangeCount, 0)) + 1
            Case 0
                postCommitPayload = postCommitPayload & "Present|"
                unchangedCount = unchangedCount + 1
            Case Else
                skippedCount = skippedCount + 1
                optionalReason = ManagedObjectOptionalMissingReason(expectedExistencePolicy, setupAction)
                If Len(optionalReason) > 0 Then
                    optionalAbsentCount = optionalAbsentCount + 1
                    optionalAbsentDetails = SharedText_AppendNonBlankLine(optionalAbsentDetails, objectTypeCode & " [" & objectName & "] | " & optionalReason)
                    postCommitPayload = postCommitPayload & FingerprintTokenPart("OptionalAbsent:" & optionalReason)
                Else
                    missingOrUnsupportedCount = missingOrUnsupportedCount + 1
                    missingOrUnsupportedDetails = SharedText_AppendNonBlankLine(missingOrUnsupportedDetails, objectTypeCode & " [" & objectName & "]")
                    postCommitPayload = postCommitPayload & FingerprintTokenPart("RequiredMissingOrUnsupported")
                End If
        End Select
        If checkedCount = 1 Or checkedCount Mod 25 = 0 Or checkedCount = descriptionRows.count Then RequireSafeWorkContinue "ApplyManagedObjectDescriptionsFromCatalog", objectTypeCode & " [" & objectName & "]"
    Next rowData
    If missingOrUnsupportedCount > 0 Then Err.Raise vbObjectError + 6437, "ApplyManagedObjectDescriptionsFromCatalog", "One or more required active managed objects were missing or unsupported during strict description application." & vbCrLf & "Required count: " & CStr(missingOrUnsupportedCount) & vbCrLf & missingOrUnsupportedDetails
    If boundedDescriptionCount > 0 Then LogTraceMessage TRACE_LEVEL_DETAIL, "ADC_ManagedObjects", "ApplyManagedObjectDescriptionsFromCatalog", "Deterministic property-safe Access Description representations were applied or verified.", boundedDescriptionDetails, "Apply managed-object descriptions", 0, vbNullString, vbNullString, vbNullString, "ManagedObject", "CodeDb", "ManagedObjectDescriptions"
    If optionalAbsentCount > 0 Then LogTraceMessage TRACE_LEVEL_DETAIL, "ADC_ManagedObjects", "ApplyManagedObjectDescriptionsFromCatalog", "Optional managed objects were absent during managed-description application.", optionalAbsentDetails, "Apply managed-object descriptions", 0, vbNullString, vbNullString, vbNullString, "ManagedObject", "CodeDb", "ManagedObjectDescriptions"
    If fullCatalogScope And Len(activeTraceRunID) > 0 Then
        mManagedObjectDescriptionToken = FingerprintFromPayload(postCommitPayload)
        mManagedObjectDescriptionTokenTraceRunID = activeTraceRunID
    End If
    If fullCatalogScope Then TryPersistManagedDescriptionExactState currentDb, inputFingerprint, FingerprintFromPayload(postCommitPayload), Nothing
    If helpDocumentMacrosOnly Then scopedExactToken = FingerprintFromPayload(postCommitPayload)
    physicalChangeCount = changedCount
    CorePulseProgress "Managed-object description application complete: changed=" & CStr(changedCount) & ", unchanged=" & CStr(unchangedCount) & ", skipped=" & CStr(skippedCount) & "."
    resultDetail = "Blank catalog descriptions repaired: " & CStr(repairedBlankDescriptionCount) & vbCrLf & _
                   "Managed-object Access descriptions changed: " & CStr(changedCount) & vbCrLf & _
                   "Property-safe descriptions applied or verified: " & CStr(boundedDescriptionCount) & vbCrLf & _
                   "Already current: " & CStr(unchangedCount) & vbCrLf & _
                   "Optional absent objects: " & CStr(optionalAbsentCount) & vbCrLf & _
                   "Required missing/unsupported objects: " & CStr(missingOrUnsupportedCount) & vbCrLf & _
                   "Help-document descriptions: optional; supplied values remain authoritative" & vbCrLf & _
                   "Authoritative physical write/verify passes: 1"
    EndManagedDescriptionOptionalHelpPolicyScope
    ReportCoreActionSuccess "ApplyManagedObjectDescriptionsFromCatalog", "ADC_ManagedObjects", "Managed-object Access descriptions applied successfully.", resultDetail, "Apply managed-object descriptions", ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID), "Table", "CodeDb", "ADC_ManagedObjects", False, DEFAULT_UTILITY_LONG_NAME & " Managed Objects"
    Exit Sub
ErrHandler:
    Dim handlerErrNumber As Long
    Dim handlerErrSource As String
    Dim handlerErrDescription As String
    Dim messageText As String
    handlerErrNumber = Err.Number
    handlerErrSource = Err.Source
    handlerErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    EndManagedDescriptionOptionalHelpPolicyScope
    On Error GoTo 0
    If handlerErrNumber = 0 Then handlerErrNumber = vbObjectError + 6436
    messageText = "Managed-object Access description application failed." & vbCrLf & _
        "Procedure: ApplyManagedObjectDescriptionsFromCatalog" & vbCrLf & _
        "Last object type: " & objectTypeCode & vbCrLf & _
        "Last object name: " & objectName & vbCrLf & _
        "Err.Source: " & handlerErrSource & vbCrLf & _
        "Err.Description: " & handlerErrDescription
    LogTraceMessage TRACE_LEVEL_ERROR, "ADC_ManagedObjects", "ApplyManagedObjectDescriptionsFromCatalog", "Managed-object Access description application failed.", messageText, vbNullString, handlerErrNumber, handlerErrSource, handlerErrDescription, objectName, objectTypeCode, "CodeDb", "ManagedObjectDescriptions"
    Err.Raise handlerErrNumber, "ApplyManagedObjectDescriptionsFromCatalog", messageText
End Sub

Private Function ManagedObjectDescriptionPostCommitToken(Optional ByVal persistExactState As Boolean = True) As String
    'Returns a fail-closed token only when every required active managed object exists and every physical Description exactly matches the catalog; active Help-document macro descriptions may be authoritatively
    'blank.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim descriptionRows As Collection
    Dim rowData As Variant
    Dim sqlText As String
    Dim objectName As String
    Dim objectTypeCode As String
    Dim expectedExistencePolicy As String
    Dim setupAction As String
    Dim descriptionText As String
    Dim physicalDescriptionText As String
    Dim optionalReason As String
    Dim payloadText As String
    Dim activeTraceRunID As String
    Dim physicalSnapshot As Object
    Dim targetObjectCache As Object
    Dim inputFingerprint As String
    Dim currentPhysicalState As String
    Dim storedInputFingerprint As String
    Dim storedPhysicalState As String
    Dim storedExactToken As String
    Dim inputPropertyExists As Boolean
    Dim physicalPropertyExists As Boolean
    Dim exactPropertyExists As Boolean
    Dim checkedCount As Long
    Dim snapshotLoadCount As Long
    Dim descriptionOptionalForHelpMacro As Boolean
    activeTraceRunID = GetActiveTraceRunID()
    If persistExactState Then
        If Len(activeTraceRunID) > 0 And StrComp(mManagedObjectDescriptionTokenTraceRunID, activeTraceRunID, vbBinaryCompare) = 0 And Len(mManagedObjectDescriptionToken) > 0 Then
            ManagedObjectDescriptionPostCommitToken = mManagedObjectDescriptionToken
            Exit Function
        End If
    End If
    BeginManagedDescriptionOptionalHelpPolicyScope
    CorePulseProgress "Checking managed-object description input and compact physical-state fingerprints."
    Set db = ADCDb()
    inputFingerprint = ManagedDescriptionInputFingerprint(db, False)
    If persistExactState Then
        storedInputFingerprint = DatabaseTextPropertyValue(currentDb, MANAGED_DESCRIPTION_FINGERPRINT_PROPERTY, inputPropertyExists)
        storedPhysicalState = DatabaseTextPropertyValue(currentDb, MANAGED_DESCRIPTION_PHYSICAL_STATE_PROPERTY, physicalPropertyExists)
        storedExactToken = DatabaseTextPropertyValue(currentDb, MANAGED_DESCRIPTION_EXACT_TOKEN_PROPERTY, exactPropertyExists)
        If inputPropertyExists And physicalPropertyExists And exactPropertyExists And StrComp(storedInputFingerprint, inputFingerprint, vbBinaryCompare) = 0 And Len(storedPhysicalState) > 0 And Len(storedExactToken) > 0 Then
            currentPhysicalState = ManagedDescriptionCompactPhysicalState(currentDb)
            If StrComp(currentPhysicalState, storedPhysicalState, vbBinaryCompare) = 0 Then
                ManagedObjectDescriptionPostCommitToken = storedExactToken
                If Len(activeTraceRunID) > 0 Then
                    mManagedObjectDescriptionToken = storedExactToken
                    mManagedObjectDescriptionTokenTraceRunID = activeTraceRunID
                End If
                CorePulseProgress "Managed-object description compact state matched."
                EndManagedDescriptionOptionalHelpPolicyScope
                Exit Function
            End If
        End If
    End If
    CorePulseProgress "Compact managed-object description state did not match; collecting authoritative object descriptions."
    sqlText = "SELECT m." & SQL_FIELD_OBJECT_NAME & ",m." & SQL_FIELD_OBJECT_TYPE & ",m." & SQL_FIELD_OBJECT_DESCRIPTION & "," & ResolvedExpectedExistencePolicySql("o", "c") & " " & "AS " & SQL_FIELD_RESOLVED_EXPECTED_EXISTENCE_POLICY & "," & _
        ResolvedSetupActionSql("o", "c") & " " & "AS " & SQL_FIELD_RESOLVED_SETUP_ACTION & " FROM ([" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID) & "] AS m INNER JOIN [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ObjectCategories") & "] AS c ON m.[" & FIELD_CATEGORY_CODE & "]=c.[" & FIELD_CATEGORY_CODE & "]) LEFT JOIN [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECT_OVERRIDES_TABLE_LOGICAL_ID) & "] AS o ON m." & SQL_FIELD_ID & "=o.[" & FIELD_MANAGED_OBJECT_ID & "] WHERE Nz(m.[" & FIELD_MANAGED_OBJECT_STATUS & "]," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & _
        " AND Nz(c." & SQL_FIELD_OBJECT_CATEGORY_STATUS & "," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & " ORDER BY " & OrderByPolicyText("REVIEW_MANAGED_DESCRIPTION") & ";"
    ' INTENTIONAL OPERATION SNAPSHOT: descriptionRows captures one coherent source state for this invocation of ManagedObjectDescriptionPostCommitToken; validity ends with the operation rather than becoming a live
    ' data authority.
    Set descriptionRows = New Collection
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    Do While rs.EOF = False
        snapshotLoadCount = snapshotLoadCount + 1
        If snapshotLoadCount = 1 Or snapshotLoadCount Mod 50 = 0 Then RequireSafeWorkContinue "ManagedObjectDescriptionPostCommitToken", "loaded " & CStr(snapshotLoadCount) & " authoritative description rows"
        descriptionRows.Add Array(CStr(Nz(rs.fields(FIELD_OBJECT_NAME).value, vbNullString)), CStr(Nz(rs.fields(FIELD_OBJECT_TYPE).value, vbNullString)), CStr(Nz(rs.fields(FIELD_RESOLVED_EXPECTED_EXISTENCE_POLICY).value, vbNullString)), CStr(Nz(rs.fields(FIELD_RESOLVED_SETUP_ACTION).value, vbNullString)), CStr(Nz(rs.fields(FIELD_OBJECT_DESCRIPTION).value, vbNullString)))
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    CorePulseProgress "Authoritative managed-object descriptions collected: " & CStr(descriptionRows.count) & ". Building the physical object snapshot."
    payloadText = "ADCManagedObjectDescriptionPostCommitV1|" & inputFingerprint & "|"
    Set physicalSnapshot = ManagedDescriptionPhysicalSnapshot(db, descriptionRows, targetObjectCache)
    CorePulseProgress "Physical object snapshot ready; verifying managed-object descriptions."
    For Each rowData In descriptionRows
        checkedCount = checkedCount + 1
        objectName = CStr(rowData(0))
        objectTypeCode = CStr(rowData(1))
        PulseProgressEvery checkedCount, 25, "Verifying managed-object description " & CStr(checkedCount) & " of " & CStr(descriptionRows.count) & ": " & objectTypeCode & " [" & objectName & "].", descriptionRows.count, "ReviewDescriptionExactState", "Review"
        expectedExistencePolicy = CStr(rowData(2))
        setupAction = CStr(rowData(3))
        descriptionText = CStr(rowData(4))
        descriptionOptionalForHelpMacro = ManagedDescriptionIsOptionalHelpMacro(objectName, objectTypeCode)
        If Len(Trim$(descriptionText)) = 0 And descriptionOptionalForHelpMacro = False Then
            If StrComp(Trim$(objectTypeCode), ADCAT_OBJECT_TYPE_FORM, vbTextCompare) = 0 Then
                Err.Raise vbObjectError + 6464, "ManagedObjectDescriptionPostCommitToken", "An active managed-form catalog description is blank. G1 owns authoritative managed-form description repair before exact verification." & vbCrLf & "Object: " & objectTypeCode & " [" & objectName & "]"
            Else
                Err.Raise vbObjectError + 6442, "ManagedObjectDescriptionPostCommitToken", "An active Managed Object Catalog description is blank outside the optional Help-document macro policy." & vbCrLf & "Object: " & objectTypeCode & " [" & objectName & "]"
            End If
        End If
        physicalDescriptionText = ManagedDescriptionExpectedPhysicalText(descriptionText)
        payloadText = payloadText & FingerprintTokenPart(objectTypeCode) & FingerprintTokenPart(objectName) & FingerprintTokenPart(descriptionText) & FingerprintTokenPart(physicalDescriptionText)
        If ManagedDescriptionSnapshotContains(physicalSnapshot, objectTypeCode, objectName) Then
            If ManagedDescriptionSnapshotMatches(physicalSnapshot, objectTypeCode, objectName, descriptionText) = False Then
                If StrComp(Trim$(objectTypeCode), ADCAT_OBJECT_TYPE_FORM, vbTextCompare) = 0 Then
                    Err.Raise vbObjectError + 6465, "ManagedObjectDescriptionPostCommitToken", "A physical managed-form Description does not match its authoritative catalog value. G1 owns physical managed-form Description repair before exact verification." & vbCrLf & "Object: " & objectTypeCode & " [" & objectName & "]"
                Else
                    Err.Raise vbObjectError + 6443, "ManagedObjectDescriptionPostCommitToken", "A physical managed-object Description does not match the deterministic property-safe representation of its authoritative catalog value." & vbCrLf & "Object: " & objectTypeCode & " [" & objectName & "]"
                End If
            End If
            payloadText = payloadText & "Present|"
        Else
            optionalReason = ManagedObjectOptionalMissingReason(expectedExistencePolicy, setupAction)
            If Len(optionalReason) = 0 Then Err.Raise vbObjectError + 6444, "ManagedObjectDescriptionPostCommitToken", "A required active managed-object Description target is missing or unsupported during exact verification." & vbCrLf & "Object: " & objectTypeCode & " [" & objectName & "]"
            payloadText = payloadText & FingerprintTokenPart("OptionalAbsent:" & optionalReason)
        End If
        If checkedCount = 1 Or checkedCount Mod 25 = 0 Or checkedCount = descriptionRows.count Then RequireSafeWorkContinue "ManagedObjectDescriptionPostCommitToken", objectTypeCode & " [" & objectName & "]"
    Next rowData
    ManagedObjectDescriptionPostCommitToken = FingerprintFromPayload(payloadText)
    If persistExactState Then
        TryPersistManagedDescriptionExactState currentDb, inputFingerprint, ManagedObjectDescriptionPostCommitToken, targetObjectCache
        If Len(activeTraceRunID) > 0 Then
            mManagedObjectDescriptionToken = ManagedObjectDescriptionPostCommitToken
            mManagedObjectDescriptionTokenTraceRunID = activeTraceRunID
        End If
    End If
    CorePulseProgress "Managed-object description exact-state verification completed: objects=" & CStr(checkedCount) & "."
    EndManagedDescriptionOptionalHelpPolicyScope
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
    EndManagedDescriptionOptionalHelpPolicyScope
    On Error GoTo 0
    Err.Raise savedErrNumber, "ManagedObjectDescriptionPostCommitToken", "Could not calculate the exact managed-object description post-commit token." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function ManagedObjectDescriptionCurrentTokenForSetup(ByRef detailText As String, Optional ByVal allowFullExactFallback As Boolean = True) As String
    ' Read-only H1/J1 consumer: reuse G1's persisted exact token only when the current input fingerprint and compact physical state still match; otherwise run the existing exact verifier without persisting.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim hostDb As DAO.Database
    Dim inputFingerprint As String
    Dim currentPhysicalState As String
    Dim storedInputFingerprint As String
    Dim storedPhysicalState As String
    Dim storedExactToken As String
    Dim inputPropertyExists As Boolean
    Dim physicalPropertyExists As Boolean
    Dim exactPropertyExists As Boolean
    Dim activeTraceRunID As String
    detailText = vbNullString
    activeTraceRunID = GetActiveTraceRunID()
    If Len(activeTraceRunID) > 0 And StrComp(mManagedObjectDescriptionTokenTraceRunID, activeTraceRunID, vbBinaryCompare) = 0 And Len(mManagedObjectDescriptionToken) > 0 Then
        If mHelpDocumentMacroBuildRepairCompactStateCaptureActive Then
            currentPhysicalState = ManagedDescriptionCompactPhysicalState(currentDb)
            If Len(currentPhysicalState) = 0 Then Err.Raise vbObjectError + 6815, "ManagedObjectDescriptionCurrentTokenForSetup", "Build/Repair Help Document Macros could not capture its required pre-mutation compact physical-state cache from the same-run G1 proof."
        End If
        ManagedObjectDescriptionCurrentTokenForSetup = mManagedObjectDescriptionToken
        detailText = "Managed-object descriptions=current through same-run G1 exact proof; full physical description scan=False" & IIf(mHelpDocumentMacroBuildRepairCompactStateCaptureActive, "; Build/Repair Help Document Macros compact-state baseline captured=True.", ".")
        Exit Function
    End If
    Set db = ADCDb()
    Set hostDb = currentDb
    inputFingerprint = ManagedDescriptionInputFingerprint(db, False)
    storedInputFingerprint = DatabaseTextPropertyValue(hostDb, MANAGED_DESCRIPTION_FINGERPRINT_PROPERTY, inputPropertyExists)
    storedPhysicalState = DatabaseTextPropertyValue(hostDb, MANAGED_DESCRIPTION_PHYSICAL_STATE_PROPERTY, physicalPropertyExists)
    storedExactToken = DatabaseTextPropertyValue(hostDb, MANAGED_DESCRIPTION_EXACT_TOKEN_PROPERTY, exactPropertyExists)
    If inputPropertyExists And physicalPropertyExists And exactPropertyExists Then
        If StrComp(storedInputFingerprint, inputFingerprint, vbBinaryCompare) = 0 And Len(storedPhysicalState) > 0 And Len(storedExactToken) > 0 Then
            currentPhysicalState = ManagedDescriptionCompactPhysicalState(hostDb)
            If StrComp(currentPhysicalState, storedPhysicalState, vbBinaryCompare) = 0 Then
                ManagedObjectDescriptionCurrentTokenForSetup = storedExactToken
                detailText = "Managed-object descriptions=current through G1 persisted exact proof plus compact physical-state gate; full physical description scan=False."
                Exit Function
            End If
        End If
    End If
    If allowFullExactFallback Then
        ManagedObjectDescriptionCurrentTokenForSetup = ManagedObjectDescriptionPostCommitToken(False)
        detailText = "Managed-object description compact gate was unavailable or stale; existing full exact read-only verification fallback executed."
    Else
        detailText = "Managed-object description compact gate was unavailable or stale; full physical description verification was intentionally not executed by this lightweight currentness check."
    End If
    Exit Function
ErrHandler:
    detailText = "Managed-object description current-token proof failed. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    ManagedObjectDescriptionCurrentTokenForSetup = vbNullString
End Function

Private Function ManagedDescriptionWholeCatalogExactToken(ByVal objectDescriptionToken As String, ByVal fieldDescriptionToken As String) As String
    Dim payloadText As String
    payloadText = MANAGED_DESCRIPTION_WHOLE_CATALOG_EXACT_TOKEN_VERSION & "|" & FingerprintTokenPart(objectDescriptionToken) & FingerprintTokenPart(fieldDescriptionToken)
    ManagedDescriptionWholeCatalogExactToken = FingerprintFromPayload(payloadText)
End Function

' Purpose: Applies and verify managed object description for standalone repair for object name, object type code.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   objectName (Input; String): Physical Access object name identifying the database object targeted by the operation.
'   objectTypeCode (Input; String): Access object-type code identifying the class of database object targeted by the operation.
'   descriptionChanged (Input/output Optional; Variant): Description changed text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
' Returns: Boolean: True on the function's affirmative/success branch for apply and verify managed object description for standalone repair; False on its explicit negative, unavailable, rejected, or handled-failure
'     branch.
'   rejection, or not-applicable path.
Public Function ApplyAndVerifyManagedObjectDescriptionForStandaloneRepair(ByVal objectName As String, ByVal objectTypeCode As String, Optional ByRef descriptionChanged As Variant) As Boolean
    'Standalone maintenance reuses REVIEW's canonical physical Description writer. During numbered Setup this API fails closed unless the current phase is G1, preserving G1 ownership.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim canonicalDescription As String
    Dim descriptionResult As Long
    objectName = Trim$(objectName)
    objectTypeCode = Trim$(objectTypeCode)
    If Not isMissing(descriptionChanged) Then descriptionChanged = False
    If Len(objectName) = 0 Or Len(objectTypeCode) = 0 Then Err.Raise vbObjectError + 10101, "ApplyAndVerifyManagedObjectDescriptionForStandaloneRepair", "A managed standalone-description repair requires both object name and object type."
    If SetupDomainMutationWritesAllowedForCurrentPart(SETUP_DOMAIN_G1_DESCRIPTIONS) = False Then Err.Raise vbObjectError + 10102, "ApplyAndVerifyManagedObjectDescriptionForStandaloneRepair", "Managed physical Description mutation is deferred to G1 during numbered Setup."
    canonicalDescription = ManagedObjectCanonicalDescription(objectName, objectTypeCode)
    'Static DEFAULTS definitions remain the first authority. Dynamic managed objects such as generated System Macros fall back to the active Managed Object Catalog row that already owns their canonical Description.
    If Len(Trim$(canonicalDescription)) = 0 Then canonicalDescription = StandaloneManagedObjectCatalogDescription(objectName, objectTypeCode)
    If Len(Trim$(canonicalDescription)) = 0 Then Err.Raise vbObjectError + 10103, "ApplyAndVerifyManagedObjectDescriptionForStandaloneRepair", "No nonblank canonical managed-object Description exists for the requested standalone repair." & vbCrLf & "Object type: " & objectTypeCode & vbCrLf & "Object name: " & objectName
    Set db = ADCDb()
    descriptionResult = SetAccessObjectDescriptionByCatalogType(db, objectName, objectTypeCode, canonicalDescription, False)
    If descriptionResult < 0 Then Err.Raise vbObjectError + 10104, "ApplyAndVerifyManagedObjectDescriptionForStandaloneRepair", "The managed object was missing or unsupported while applying its standalone Description." & vbCrLf & "Object type: " & objectTypeCode & vbCrLf & "Object name: " & objectName
    If Not isMissing(descriptionChanged) Then descriptionChanged = (descriptionResult = 1)
    ApplyAndVerifyManagedObjectDescriptionForStandaloneRepair = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If IsCancellationError(savedErrNumber) Then Err.Raise savedErrNumber, savedErrSource, savedErrDescription
    Err.Raise savedErrNumber, "ApplyAndVerifyManagedObjectDescriptionForStandaloneRepair", "Could not apply and verify the canonical managed-object Description during standalone repair." & vbCrLf & "Object type: " & objectTypeCode & vbCrLf & "Object name: " & objectName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function StandaloneManagedObjectCatalogDescription(ByVal objectName As String, ByVal objectTypeCode As String) As String
    'Fallback authority for dynamic managed objects that do not have a static DEFAULTS row. It is read-only and is used only by standalone Description completion outside numbered Setup.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim managedObjectsTableName As String
    Dim sqlText As String
    managedObjectsTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID)
    Set db = ADCDb()
    If SharedDao_TableExists(db, managedObjectsTableName) = False Then Exit Function
    sqlText = "SELECT " & SQL_FIELD_OBJECT_DESCRIPTION & " FROM [" & managedObjectsTableName & "] WHERE UCase(Trim(Nz(" & SQL_FIELD_OBJECT_NAME & ",'')))=" & SharedDao_SqlQuote(UCase$(Trim$(objectName))) & _
              " AND UCase(Trim(Nz(" & SQL_FIELD_OBJECT_TYPE & ",'')))=" & SharedDao_SqlQuote(UCase$(Trim$(objectTypeCode))) & _
              " AND Nz([" & FIELD_MANAGED_OBJECT_STATUS & "]," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ";"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    If rs.EOF Then GoTo SafeExit
    StandaloneManagedObjectCatalogDescription = CStr(Nz(rs.fields(FIELD_OBJECT_DESCRIPTION).value, vbNullString))
    rs.MoveNext
    If rs.EOF = False Then Err.Raise vbObjectError + 10124, "StandaloneManagedObjectCatalogDescription", "More than one active Managed Object Catalog row resolves to the requested standalone Description identity." & vbCrLf & "Object type: " & objectTypeCode & vbCrLf & "Object name: " & objectName
SafeExit:
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
    Err.Raise savedErrNumber, "StandaloneManagedObjectCatalogDescription", "Could not resolve the active Managed Object Catalog Description for standalone completion." & vbCrLf & "Object type: " & objectTypeCode & vbCrLf & "Object name: " & objectName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

' Purpose: Applies and verifies the G1-owned catalog, physical-object, form/report, and managed-table field descriptions, then establishes the exact same-run description proof required by later Setup work.
' Public role: Inter-module Setup API. The Setup orchestrator calls this routine only for G1; the routine enforces that G1 owns numbered-Setup description mutation before performing any writes.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether G1 displays its normal completion/error message after description work; False keeps the same G1 work and tracing but suppresses that popup. Default: False.
'   detailsForMessage (Output Optional; String): Receives the G1 verification summary, including repaired-description counts, exact proof tokens, catalog refresh state, and elapsed time; on handled failure it
'   receives the failure details. Default: vbNullString.
' Returns: Boolean: True only after G1 finishes its owned description writes and exact-state proof successfully; False only on the handled showMessage error path. Cancellation and unhandled errors are re-raised.
Public Function ApplyAndVerifyManagedDescriptionsForSetup(Optional ByVal showMessage As Boolean = False, Optional ByRef detailsForMessage As String = vbNullString) As Boolean
    RequireSetupDomainMutationAllowed SETUP_DOMAIN_G1_DESCRIPTIONS, "ApplyAndVerifyManagedDescriptionsForSetup"
    ' G1 is the single authoritative Setup path for catalog, field, and physical Access descriptions for every managed object, including forms and reports. Help-document macro descriptions are optional but exact
    ' when supplied.
    On Error GoTo ErrHandler
    Dim fieldDescriptionToken As String
    Dim objectDescriptionToken As String
    Dim wholeCatalogToken As String
    Dim repairedCatalogDescriptionCount As Long
    Dim fieldDescriptionChangeCount As Variant
    Dim physicalDescriptionChangeCount As Variant
    Dim formPhysicalDescriptionChangeCount As Variant
    Dim catalogRefreshDetail As String
    Dim fieldProofFinalizeDetail As String
    Dim descriptionsChanged As Boolean
    Dim phaseTimer As Single
    phaseTimer = Timer
    BeginManagedDescriptionOptionalHelpPolicyScope
    CorePulseProgress "Setup G1: normalizing authoritative managed-object descriptions."
    repairedCatalogDescriptionCount = RepairManagedObjectCatalogDescriptions(False, True, False)
    If repairedCatalogDescriptionCount > 0 Then MarkCatalogManagedObjectDescriptionMutationForSetup
    CorePulseProgress "Setup G1: applying and verifying managed-table field descriptions."
    fieldDescriptionChangeCount = CLng(0)
    fieldDescriptionToken = ManagedTableFieldDescriptionPostCommitToken(fieldDescriptionChangeCount)
    If Len(fieldDescriptionToken) = 0 Then Err.Raise vbObjectError + 6461, "ApplyAndVerifyManagedDescriptionsForSetup", "Managed-table field-description work did not produce an exact token."
    CorePulseProgress "Setup G1: applying and verifying managed-object descriptions, including forms and reports."
    physicalDescriptionChangeCount = CLng(0)
    formPhysicalDescriptionChangeCount = CLng(0)
    ApplyManagedObjectDescriptionsFromCatalog MANAGED_OBJECT_DESCRIPTION_SCOPE_ALL, physicalChangeCount:=physicalDescriptionChangeCount, formPhysicalChangeCount:=formPhysicalDescriptionChangeCount
    descriptionsChanged = repairedCatalogDescriptionCount > 0 Or _
                          CLng(Nz(fieldDescriptionChangeCount, 0)) > 0 Or _
                          CLng(Nz(physicalDescriptionChangeCount, 0)) > 0
    If descriptionsChanged Then
        CorePulseProgress "Setup G1: refreshing the complete catalog exact state once after description writes."
        If RefreshCompleteCatalogStateAfterDescriptionWorkForSetup(catalogRefreshDetail) = False Then Err.Raise vbObjectError + 6463, "ApplyAndVerifyManagedDescriptionsForSetup", catalogRefreshDetail
    End If
    CorePulseProgress "Setup G1: binding field-description proof to final post-description table metadata."
    If FinalizeManagedTableFieldDescriptionProofAfterG1DescriptionWork(fieldDescriptionToken, fieldProofFinalizeDetail) = False Then Err.Raise vbObjectError + 6487, "ApplyAndVerifyManagedDescriptionsForSetup", fieldProofFinalizeDetail
    objectDescriptionToken = mManagedObjectDescriptionToken
    If Len(objectDescriptionToken) = 0 Then objectDescriptionToken = ManagedObjectDescriptionPostCommitToken()
    If Len(objectDescriptionToken) = 0 Then Err.Raise vbObjectError + 6462, "ApplyAndVerifyManagedDescriptionsForSetup", "Managed-object description work did not produce an exact token."
    wholeCatalogToken = ManagedDescriptionWholeCatalogExactToken(objectDescriptionToken, fieldDescriptionToken)
    Call EnsureDaoPropertyValue(currentDb, MANAGED_DESCRIPTION_WHOLE_CATALOG_EXACT_TOKEN_PROPERTY, dbText, wholeCatalogToken, "ApplyAndVerifyManagedDescriptionsForSetup")
    detailsForMessage = "Setup G1 managed descriptions are exact." & vbCrLf & _
                        "Authoritative catalog descriptions repaired: " & CStr(repairedCatalogDescriptionCount) & vbCrLf & _
                        "Physical managed-object descriptions changed: " & CStr(CLng(Nz(physicalDescriptionChangeCount, 0))) & vbCrLf & _
                        "Physical managed-form descriptions changed: " & CStr(CLng(Nz(formPhysicalDescriptionChangeCount, 0))) & vbCrLf & _
                        "Managed-table field descriptions changed: " & CStr(CLng(Nz(fieldDescriptionChangeCount, 0))) & vbCrLf & _
                        "Help-document descriptions: optional; supplied values exact" & vbCrLf & _
                        "Complete catalog state refreshed after mutation: " & CStr(descriptionsChanged) & vbCrLf & _
                        fieldProofFinalizeDetail & vbCrLf & _
                        "Managed-object exact token: " & objectDescriptionToken & vbCrLf & _
                        "Managed-table field-description exact token: " & fieldDescriptionToken & vbCrLf & _
                        "Whole-catalog description token: " & wholeCatalogToken & vbCrLf & _
                        "Form mutation in G1: " & CStr(CLng(Nz(formPhysicalDescriptionChangeCount, 0)) > 0) & vbCrLf & _
                        "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseTimer))
    EndManagedDescriptionOptionalHelpPolicyScope
    LogTraceMessage TRACE_LEVEL_INFO, "ManagedObjectDescriptions", "ApplyAndVerifyManagedDescriptionsForSetup", "Setup G1 managed-description exact state established through the single write path.", detailsForMessage, "G1 managed descriptions", 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID), "ManagedObjectDescriptions", "CodeDb", "ManagedObjectDescriptions"
    If showMessage Then CoreShowMessage detailsForMessage, vbInformation, DEFAULT_UTILITY_LONG_NAME & " Managed Descriptions", "ShowADCMessage", "Tables"
    ApplyAndVerifyManagedDescriptionsForSetup = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    EndManagedDescriptionOptionalHelpPolicyScope
    On Error GoTo 0
    If IsCancellationError(savedErrNumber) Then Err.Raise savedErrNumber, savedErrSource, savedErrDescription
    detailsForMessage = "Setup G1 managed-description work failed." & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
    LogTraceMessage TRACE_LEVEL_ERROR, "ManagedObjectDescriptions", "ApplyAndVerifyManagedDescriptionsForSetup", "Setup G1 managed-description work failed.", detailsForMessage, "G1 managed descriptions", savedErrNumber, savedErrSource, savedErrDescription, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID), "ManagedObjectDescriptions", "CodeDb", "ManagedObjectDescriptions"
    If showMessage Then CoreShowErrorMessage detailsForMessage, vbCritical, DEFAULT_UTILITY_LONG_NAME & " Managed Descriptions", "ApplyAndVerifyManagedDescriptionsForSetup", "Tables": Exit Function
    Err.Raise savedErrNumber, "ApplyAndVerifyManagedDescriptionsForSetup", detailsForMessage
End Function

Private Function NonHelpManagedObjectDescriptionProofForHelpDocumentMacroBuildRepair(ByRef detailText As String) As String
    ' REVIEW-owned fallback proof for Build/Repair Help Document Macros recovery: verifies every non-Help managed-object physical Description exactly while intentionally excluding Help macros that this operation is
    ' authorized to repair.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim descriptionRows As Collection
    Dim rowData As Variant
    Dim physicalSnapshot As Object
    Dim targetObjectCache As Object
    Dim sqlText As String
    Dim payloadText As String
    Dim objectName As String
    Dim objectTypeCode As String
    Dim expectedExistencePolicy As String
    Dim setupAction As String
    Dim descriptionText As String
    Dim optionalReason As String
    Dim checkedCount As Long
    Dim skippedHelpCount As Long
    BeginManagedDescriptionOptionalHelpPolicyScope
    Set db = ADCDb()
    sqlText = "SELECT m." & SQL_FIELD_OBJECT_NAME & ",m." & SQL_FIELD_OBJECT_TYPE & ",m." & SQL_FIELD_OBJECT_DESCRIPTION & "," & ResolvedExpectedExistencePolicySql("o", "c") & " " & "AS " & SQL_FIELD_RESOLVED_EXPECTED_EXISTENCE_POLICY & "," & _
        ResolvedSetupActionSql("o", "c") & " " & "AS " & SQL_FIELD_RESOLVED_SETUP_ACTION & " FROM ([" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID) & "] AS m INNER JOIN [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ObjectCategories") & "] AS c ON m.[" & FIELD_CATEGORY_CODE & "]=c.[" & FIELD_CATEGORY_CODE & "]) LEFT JOIN [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECT_OVERRIDES_TABLE_LOGICAL_ID) & "] AS o ON m." & SQL_FIELD_ID & "=o.[" & FIELD_MANAGED_OBJECT_ID & "] WHERE Nz(m.[" & FIELD_MANAGED_OBJECT_STATUS & "]," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & _
        " AND Nz(c." & SQL_FIELD_OBJECT_CATEGORY_STATUS & "," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & " ORDER BY " & OrderByPolicyText("REVIEW_MANAGED_DESCRIPTION") & ";"
    Set descriptionRows = New Collection
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    Do While rs.EOF = False
        objectName = CStr(Nz(rs.fields(FIELD_OBJECT_NAME).value, vbNullString))
        objectTypeCode = CStr(Nz(rs.fields(FIELD_OBJECT_TYPE).value, vbNullString))
        If ManagedDescriptionIsOptionalHelpMacro(objectName, objectTypeCode) Then
            skippedHelpCount = skippedHelpCount + 1
        Else
            descriptionText = CStr(Nz(rs.fields(FIELD_OBJECT_DESCRIPTION).value, vbNullString))
            If Len(Trim$(descriptionText)) = 0 Then Err.Raise vbObjectError + 6513, "NonHelpManagedObjectDescriptionProofForHelpDocumentMacroBuildRepair", "An active non-Help Managed Object Catalog Description is blank." & vbCrLf & "Object: " & objectTypeCode & " [" & objectName & "]"
            descriptionRows.Add Array(objectName, objectTypeCode, CStr(Nz(rs.fields(FIELD_RESOLVED_EXPECTED_EXISTENCE_POLICY).value, vbNullString)), CStr(Nz(rs.fields(FIELD_RESOLVED_SETUP_ACTION).value, vbNullString)), descriptionText)
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Set physicalSnapshot = ManagedDescriptionPhysicalSnapshot(db, descriptionRows, targetObjectCache)
    payloadText = "ADCBuildRepairHelpDocumentMacrosNonHelpDescriptionProofV1|"
    For Each rowData In descriptionRows
        checkedCount = checkedCount + 1
        objectName = CStr(rowData(0))
        objectTypeCode = CStr(rowData(1))
        expectedExistencePolicy = CStr(rowData(2))
        setupAction = CStr(rowData(3))
        descriptionText = CStr(rowData(4))
        payloadText = payloadText & FingerprintTokenPart(objectTypeCode) & FingerprintTokenPart(objectName) & FingerprintTokenPart(descriptionText) & FingerprintTokenPart(ManagedDescriptionExpectedPhysicalText(descriptionText))
        If ManagedDescriptionSnapshotContains(physicalSnapshot, objectTypeCode, objectName) Then
            If ManagedDescriptionSnapshotMatches(physicalSnapshot, objectTypeCode, objectName, descriptionText) = False Then Err.Raise vbObjectError + 6514, "NonHelpManagedObjectDescriptionProofForHelpDocumentMacroBuildRepair", "A non-Help managed-object physical Description does not match its authoritative catalog value." & vbCrLf & "Object: " & objectTypeCode & " [" & objectName & "]"
            payloadText = payloadText & "Present|"
        Else
            optionalReason = ManagedObjectOptionalMissingReason(expectedExistencePolicy, setupAction)
            If Len(optionalReason) = 0 Then Err.Raise vbObjectError + 6515, "NonHelpManagedObjectDescriptionProofForHelpDocumentMacroBuildRepair", "A required non-Help managed-object Description target is missing or unsupported." & vbCrLf & "Object: " & objectTypeCode & " [" & objectName & "]"
            payloadText = payloadText & FingerprintTokenPart("OptionalAbsent:" & optionalReason)
        End If
        If checkedCount = 1 Or checkedCount Mod 25 = 0 Or checkedCount = descriptionRows.count Then RequireSafeWorkContinue "NonHelpManagedObjectDescriptionProofForHelpDocumentMacroBuildRepair", objectTypeCode & " [" & objectName & "]"
    Next rowData
    NonHelpManagedObjectDescriptionProofForHelpDocumentMacroBuildRepair = FingerprintFromPayload(payloadText)
    detailText = "Non-Help managed-object descriptions=current through exact read-only verification; objects verified=" & CStr(checkedCount) & "; Help macros intentionally excluded=" & CStr(skippedHelpCount) & "."
    EndManagedDescriptionOptionalHelpPolicyScope
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
    EndManagedDescriptionOptionalHelpPolicyScope
    On Error GoTo 0
    If IsCancellationError(savedErrNumber) Then Err.Raise savedErrNumber, savedErrSource, savedErrDescription
    detailText = "Non-Help managed-object Description proof failed. Error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription
    NonHelpManagedObjectDescriptionProofForHelpDocumentMacroBuildRepair = vbNullString
End Function

' Purpose: Initializes help document macro build repair full ribbon predecessor proof, coordinating `GetActiveTraceRunID` and `RibbonDependencyTokenForSetup` as the implementation requires.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   detailText (Input/output; String): Receives or supplies diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
'   provenCatalogTokenForBuildRepair (Output Optional; String): Receives the resolved proven catalog token for build repair produced from this procedure's validated inputs and current ADC state. Default:
'     vbNullString.
' Returns: Boolean: True on the function's affirmative/success branch for begin help document macro build repair full ribbon predecessor proof; False on its explicit negative, unavailable, rejected, or
'     handled-failure branch.
'   rejection, or not-applicable path.
Public Function BeginHelpDocumentMacroBuildRepairFullRibbonPredecessorProof(ByRef detailText As String, Optional ByRef provenCatalogTokenForBuildRepair As String = vbNullString) As Boolean
    ' Prove the unchanged predecessor domains before Build/Repair Help Document Macros mutates Help-owned state. Prefer the complete persisted/compact predecessor fast path; otherwise allow only Help_Docs row input
    ' plus Help physical Description drift while proving all unrelated Catalog/G1 state exact.
    On Error GoTo ErrHandler
    Dim activeTraceRunID As String
    Dim fullPredecessorToken As String
    Dim fullPredecessorFailureDetail As String
    Dim provenObjectDescriptionToken As String
    Dim provenFieldDescriptionToken As String
    Dim provenCatalogToken As String
    Dim nonHelpDescriptionToken As String
    Dim nonHelpDescriptionDetail As String
    Dim catalogToken As String
    Dim exactFallbackCatalogToken As String
    Dim catalogDetail As String
    Dim fieldDetail As String
    Dim catalogProofScopeStarted As Boolean
    Dim capturedBaselinePhysicalState As String
    provenCatalogTokenForBuildRepair = vbNullString
    mHelpDocumentMacroBuildRepairBaselineScopeProofToken = vbNullString
    mHelpDocumentMacroBuildRepairBaselineScopeProofTraceRunID = vbNullString
    mHelpDocumentMacroBuildRepairBaselineFieldDescriptionToken = vbNullString
    mHelpDocumentMacroBuildRepairFullPredecessorToken = vbNullString
    mHelpDocumentMacroBuildRepairFullPredecessorTraceRunID = vbNullString
    mHelpDocumentMacroBuildRepairFullPredecessorCatalogToken = vbNullString
    InvalidateHelpDocumentMacroBuildRepairBaselineCompactStateCache
    mHelpDocumentMacroBuildRepairCompactStateCaptureActive = True
    activeTraceRunID = GetActiveTraceRunID()
    If Len(activeTraceRunID) = 0 Then Err.Raise vbObjectError + 6494, "BeginHelpDocumentMacroBuildRepairFullRibbonPredecessorProof", "A trace-run identity is required before capturing the Build/Repair Help Document Macros predecessor proof."
    BeginBuildRepairHelpDocumentMacrosCatalogPredecessorProofScope
    catalogProofScopeStarted = True
    fullPredecessorToken = RibbonDependencyTokenForSetup(detailText, provenObjectDescriptionToken, provenFieldDescriptionToken, provenCatalogToken, True)
    If Len(fullPredecessorToken) > 0 Then
        If Len(provenFieldDescriptionToken) = 0 Then Err.Raise vbObjectError + 6502, "BeginHelpDocumentMacroBuildRepairFullRibbonPredecessorProof", "The complete pre-mutation proof did not expose its managed-table field-description component token."
        mHelpDocumentMacroBuildRepairBaselineScopeProofToken = fullPredecessorToken
        mHelpDocumentMacroBuildRepairBaselineScopeProofTraceRunID = activeTraceRunID
        mHelpDocumentMacroBuildRepairBaselineFieldDescriptionToken = provenFieldDescriptionToken
        provenCatalogTokenForBuildRepair = provenCatalogToken
        If Len(provenCatalogTokenForBuildRepair) = 0 Then Err.Raise vbObjectError + 6826, "BeginHelpDocumentMacroBuildRepairFullRibbonPredecessorProof", "The complete predecessor proof did not expose its current complete-catalog component token."
        If mHelpDocumentMacroBuildRepairBaselineNonHelpCompactStateByObject Is Nothing Or StrComp(mHelpDocumentMacroBuildRepairBaselineCompactStateTraceRunID, activeTraceRunID, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 6816, "BeginHelpDocumentMacroBuildRepairFullRibbonPredecessorProof", "The complete predecessor fast path did not retain the required operation-scoped non-Help compact-state cache."
        mHelpDocumentMacroBuildRepairCompactStateCaptureActive = False
        detailText = "Build/Repair Help Document Macros unchanged-domain proof=current through the complete persisted/compact D4/G1 predecessor fast path for trace run " & activeTraceRunID & "; full physical Description scans=False; non-Help compact-state cache retained for post-Help proof advancement=True." & vbCrLf & detailText
        EndBuildRepairHelpDocumentMacrosCatalogPredecessorProofScope
        catalogProofScopeStarted = False
        BeginHelpDocumentMacroBuildRepairFullRibbonPredecessorProof = True
        Exit Function
    End If
    fullPredecessorFailureDetail = detailText
    catalogToken = CompleteCatalogUnchangedExceptHelpDocumentsForBuildRepair(catalogDetail, exactFallbackCatalogToken)
    If Len(catalogToken) = 0 Then Err.Raise vbObjectError + 6516, "BeginHelpDocumentMacroBuildRepairFullRibbonPredecessorProof", "The scoped pre-mutation catalog proof is not current; Build/Repair Help Document Macros will not repair an unrelated catalog defect." & vbCrLf & catalogDetail
    If ManagedTableFieldDescriptionsPersistedProofCurrentForBuildRepairHelpDocumentMacros(provenFieldDescriptionToken, fieldDetail) = False Then Err.Raise vbObjectError + 6517, "BeginHelpDocumentMacroBuildRepairFullRibbonPredecessorProof", "The pre-mutation managed-table field-description proof is not current; Build/Repair Help Document Macros will not repair an unrelated field-description defect." & vbCrLf & fieldDetail
    nonHelpDescriptionToken = NonHelpManagedObjectDescriptionProofForHelpDocumentMacroBuildRepair(nonHelpDescriptionDetail)
    If Len(nonHelpDescriptionToken) = 0 Then Err.Raise vbObjectError + 6518, "BeginHelpDocumentMacroBuildRepairFullRibbonPredecessorProof", nonHelpDescriptionDetail
    If mHelpDocumentMacroBuildRepairBaselineNonHelpCompactStateByObject Is Nothing Or StrComp(mHelpDocumentMacroBuildRepairBaselineCompactStateTraceRunID, activeTraceRunID, vbBinaryCompare) <> 0 Then
        capturedBaselinePhysicalState = ManagedDescriptionCompactPhysicalState(currentDb)
        If Len(capturedBaselinePhysicalState) = 0 Or mHelpDocumentMacroBuildRepairBaselineNonHelpCompactStateByObject Is Nothing Or StrComp(mHelpDocumentMacroBuildRepairBaselineCompactStateTraceRunID, activeTraceRunID, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 6817, "BeginHelpDocumentMacroBuildRepairFullRibbonPredecessorProof", "The scoped predecessor fallback could not retain the required operation-scoped non-Help compact-state cache."
    End If
    mHelpDocumentMacroBuildRepairCompactStateCaptureActive = False
    provenCatalogTokenForBuildRepair = exactFallbackCatalogToken
    EndBuildRepairHelpDocumentMacrosCatalogPredecessorProofScope
    catalogProofScopeStarted = False
    mHelpDocumentMacroBuildRepairBaselineScopeProofToken = FingerprintFromPayload("BuildRepairHelpDocumentMacrosUnchangedDomainsV2|" & FingerprintTokenPart(catalogToken) & FingerprintTokenPart(provenFieldDescriptionToken) & FingerprintTokenPart(nonHelpDescriptionToken))
    mHelpDocumentMacroBuildRepairBaselineScopeProofTraceRunID = activeTraceRunID
    mHelpDocumentMacroBuildRepairBaselineFieldDescriptionToken = provenFieldDescriptionToken
    detailText = "Build/Repair Help Document Macros unchanged-domain proof=current through the scoped Help-input recovery fallback. Help_Docs row-content drift and Help physical Description state were allowed because this command owns their reconciliation; unrelated catalog, field-description, and non-Help managed-object Description state all proved current; non-Help compact-state cache retained for post-Help proof advancement=True." & vbCrLf & _
                 "Exact complete-catalog token retained for physical-Help-only recovery=" & CStr(Len(provenCatalogTokenForBuildRepair) > 0) & vbCrLf & _
                 catalogDetail & vbCrLf & fieldDetail & vbCrLf & nonHelpDescriptionDetail & vbCrLf & _
                 "Complete predecessor fast-path rejection retained for diagnostics:" & vbCrLf & fullPredecessorFailureDetail
    BeginHelpDocumentMacroBuildRepairFullRibbonPredecessorProof = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    provenCatalogTokenForBuildRepair = vbNullString
    On Error Resume Next
    If catalogProofScopeStarted Then EndBuildRepairHelpDocumentMacrosCatalogPredecessorProofScope
    On Error GoTo 0
    mHelpDocumentMacroBuildRepairBaselineScopeProofToken = vbNullString
    mHelpDocumentMacroBuildRepairBaselineScopeProofTraceRunID = vbNullString
    mHelpDocumentMacroBuildRepairBaselineFieldDescriptionToken = vbNullString
    mHelpDocumentMacroBuildRepairFullPredecessorToken = vbNullString
    mHelpDocumentMacroBuildRepairFullPredecessorTraceRunID = vbNullString
    mHelpDocumentMacroBuildRepairFullPredecessorCatalogToken = vbNullString
    InvalidateHelpDocumentMacroBuildRepairBaselineCompactStateCache
    If IsCancellationError(savedErrNumber) Then Err.Raise savedErrNumber, savedErrSource, savedErrDescription
    detailText = "Build/Repair Help Document Macros could not establish the unchanged pre-mutation predecessor domains required to safely repair Help state. Error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription
    If Len(fullPredecessorFailureDetail) > 0 Then detailText = detailText & vbCrLf & "Complete predecessor fast-path rejection retained for diagnostics:" & vbCrLf & fullPredecessorFailureDetail
    BeginHelpDocumentMacroBuildRepairFullRibbonPredecessorProof = False
End Function

' Purpose: Applies and verify help document managed descriptions for build repair.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   detailsForMessage (Output Optional; String): Receives detail text formatted for inclusion in the caller's user-facing message. Default: vbNullString.
'   catalogDescriptionMutated (Output Optional; Boolean): Set True when the procedure determines that catalog description mutated applies; set False when that condition is not met. Default: False.
'   refreshedCatalogToken (Output Optional; String): Receives the resolved refreshed catalog token produced from this procedure's validated inputs and current ADC state. Default: vbNullString.
' Returns: Boolean: True on the function's affirmative/success branch for apply and verify help document managed descriptions for build repair; False on its explicit negative, unavailable, rejected, or
'     handled-failure branch.
'   rejection, or not-applicable path.
Public Function ApplyAndVerifyHelpDocumentManagedDescriptionsForBuildRepair(Optional ByRef detailsForMessage As String = vbNullString, Optional ByRef catalogDescriptionMutated As Boolean = False, Optional ByRef refreshedCatalogToken As String = vbNullString) As Boolean
    ' G1-owned operational path: synchronize and verify only active Help-document macro catalog/physical Descriptions. If catalog Description rows change, CATALOG returns the refreshed complete token for the
    ' same-run handoff. Managed-table field descriptions and unrelated managed objects remain outside this operation scope.
    On Error GoTo ErrHandler
    Dim repairedCatalogDescriptionCount As Long
    Dim physicalDescriptionChangeCount As Variant
    Dim formPhysicalDescriptionChangeCount As Variant
    Dim scopedExactToken As Variant
    Dim catalogRefreshDetail As String
    Dim activeTraceRunID As String
    Dim phaseTimer As Single
    phaseTimer = Timer
    catalogDescriptionMutated = False
    refreshedCatalogToken = vbNullString
    activeTraceRunID = GetActiveTraceRunID()
    mHelpDocumentDescriptionProofToken = vbNullString
    mHelpDocumentDescriptionProofTraceRunID = vbNullString
    BeginManagedDescriptionOptionalHelpPolicyScope
    CorePulseProgress "Build/Repair Help Document Macros G1: synchronizing authoritative Help macro descriptions."
    repairedCatalogDescriptionCount = RepairManagedObjectCatalogDescriptions(False, True, False, True)
    If repairedCatalogDescriptionCount > 0 Then MarkCatalogManagedObjectDescriptionMutationForSetup
    CorePulseProgress "Build/Repair Help Document Macros G1: applying and verifying physical Help macro descriptions."
    physicalDescriptionChangeCount = CLng(0)
    formPhysicalDescriptionChangeCount = CLng(0)
    scopedExactToken = vbNullString
    ApplyManagedObjectDescriptionsFromCatalog MANAGED_OBJECT_DESCRIPTION_SCOPE_HELP_DOCUMENT_MACROS_ONLY, physicalChangeCount:=physicalDescriptionChangeCount, formPhysicalChangeCount:=formPhysicalDescriptionChangeCount, scopedExactToken:=scopedExactToken
    If Len(CStr(Nz(scopedExactToken, vbNullString))) = 0 Then Err.Raise vbObjectError + 6491, "ApplyAndVerifyHelpDocumentManagedDescriptionsForBuildRepair", "Help-document managed-description work did not produce a scoped exact token."
    If repairedCatalogDescriptionCount > 0 Then
        CorePulseProgress "Build/Repair Help Document Macros G1: refreshing complete catalog proof after Help description catalog synchronization."
        If RefreshCompleteCatalogStateAfterDescriptionWorkForSetup(catalogRefreshDetail, refreshedCatalogToken) = False Then Err.Raise vbObjectError + 6492, "ApplyAndVerifyHelpDocumentManagedDescriptionsForBuildRepair", catalogRefreshDetail
        If Len(refreshedCatalogToken) = 0 Then Err.Raise vbObjectError + 6821, "ApplyAndVerifyHelpDocumentManagedDescriptionsForBuildRepair", "Catalog description mutation completed without returning the refreshed authoritative complete-catalog token."
        catalogDescriptionMutated = True
    End If
    If Len(activeTraceRunID) = 0 Then Err.Raise vbObjectError + 6493, "ApplyAndVerifyHelpDocumentManagedDescriptionsForBuildRepair", "A trace-run identity is required for the same-run scoped G1 Ribbon handoff."
    mHelpDocumentDescriptionProofToken = CStr(scopedExactToken)
    mHelpDocumentDescriptionProofTraceRunID = activeTraceRunID
    detailsForMessage = "Operational Help-document G1 description state is exact." & vbCrLf & _
                        "Authoritative Help catalog descriptions synchronized: " & CStr(repairedCatalogDescriptionCount) & vbCrLf & _
                        "Complete catalog token advanced by scoped G1: " & CStr(catalogDescriptionMutated) & vbCrLf & _
                        "Physical Help macro descriptions changed: " & CStr(CLng(Nz(physicalDescriptionChangeCount, 0))) & vbCrLf & _
                        "Managed-table field descriptions inspected: 0" & vbCrLf & _
                        "Unrelated managed-object physical descriptions inspected: 0" & vbCrLf & _
                        "Help-document descriptions: optional; supplied values exact" & vbCrLf & _
                        "Scoped exact token: " & mHelpDocumentDescriptionProofToken & vbCrLf & _
                        "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseTimer))
    EndManagedDescriptionOptionalHelpPolicyScope
    LogTraceMessage TRACE_LEVEL_INFO, "ManagedObjectDescriptions", "ApplyAndVerifyHelpDocumentManagedDescriptionsForBuildRepair", "Operational Help-document G1 scoped description state established.", detailsForMessage, "Build/Repair Help Document Macros G1 descriptions", 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID), "ManagedObjectDescriptions", "CodeDb", "Performance"
    ApplyAndVerifyHelpDocumentManagedDescriptionsForBuildRepair = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    mHelpDocumentDescriptionProofToken = vbNullString
    mHelpDocumentDescriptionProofTraceRunID = vbNullString
    EndManagedDescriptionOptionalHelpPolicyScope
    On Error GoTo 0
    If IsCancellationError(savedErrNumber) Then Err.Raise savedErrNumber, savedErrSource, savedErrDescription
    catalogDescriptionMutated = False
    refreshedCatalogToken = vbNullString
    detailsForMessage = "Operational Help-document G1 description work failed." & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
    LogTraceMessage TRACE_LEVEL_ERROR, "ManagedObjectDescriptions", "ApplyAndVerifyHelpDocumentManagedDescriptionsForBuildRepair", "Operational Help-document G1 description work failed.", detailsForMessage, "Build/Repair Help Document Macros G1 descriptions", savedErrNumber, savedErrSource, savedErrDescription, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID), "ManagedObjectDescriptions", "CodeDb", "ManagedObjectDescriptions"
    Err.Raise savedErrNumber, "ApplyAndVerifyHelpDocumentManagedDescriptionsForBuildRepair", detailsForMessage
End Function

' Purpose: Performs help document description same run G1 proof for ribbon by coordinating `GetActiveTraceRunID` and `StrComp` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   proofToken (Output; String): Receives the resolved proof token produced from this procedure's validated inputs and current ADC state.
'   proofTraceRunID (Output; String): Receives the resolved proof trace run ID produced from this procedure's validated inputs and current ADC state.
'   detailText (Output; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
' Returns: Boolean: True when the condition represented by HelpDocumentDescriptionSameRunG1ProofForRibbon is present under the routine's documented inputs/current state; False when that condition is absent.
Public Function HelpDocumentDescriptionSameRunG1ProofForRibbon(ByRef proofToken As String, ByRef proofTraceRunID As String, ByRef detailText As String) As Boolean
    'Fail-closed scoped-G1 handoff: only the Help-description proof established by the G1 owner in the current trace run is accepted by the Build/Repair Help Document Macros full-predecessor finalizer.
    Dim activeTraceRunID As String
    proofToken = vbNullString
    proofTraceRunID = vbNullString
    detailText = vbNullString
    activeTraceRunID = GetActiveTraceRunID()
    If Len(activeTraceRunID) = 0 Then detailText = "No active trace run exists for the scoped Help-description G1 proof.": Exit Function
    If Len(mHelpDocumentDescriptionProofToken) = 0 Or Len(mHelpDocumentDescriptionProofTraceRunID) = 0 Then detailText = "No scoped Help-description G1 proof has been established in this operation.": Exit Function
    If StrComp(mHelpDocumentDescriptionProofTraceRunID, activeTraceRunID, vbBinaryCompare) <> 0 Then detailText = "The scoped Help-description G1 proof belongs to a different trace run.": Exit Function
    proofToken = mHelpDocumentDescriptionProofToken
    proofTraceRunID = mHelpDocumentDescriptionProofTraceRunID
    detailText = "Scoped Help-document G1 description proof=current for trace run " & activeTraceRunID & "."
    HelpDocumentDescriptionSameRunG1ProofForRibbon = True
End Function

Private Function ManagedObjectDescriptionCanonicalTokenAfterScopedHelpDocumentMacroBuildRepair(ByRef inputFingerprint As String) As String
    ' Reconstruct the canonical full managed-object exact token without rereading unrelated physical targets. The pre-refresh compact cache proves non-Help presence/absence; the same-run scoped G1 proof establishes
    ' Help descriptions and only current Help targets are reread.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim payloadText As String
    Dim objectName As String
    Dim objectTypeCode As String
    Dim expectedExistencePolicy As String
    Dim setupAction As String
    Dim descriptionText As String
    Dim physicalDescriptionText As String
    Dim optionalReason As String
    Dim identityKey As String
    Dim baselineObjectState As String
    Dim presentStatePrefix As String
    Dim activeTraceRunID As String
    Dim targetExists As Boolean
    Dim descriptionOptionalForHelpMacro As Boolean
    Dim checkedCount As Long
    BeginManagedDescriptionOptionalHelpPolicyScope
    Set db = ADCDb()
    activeTraceRunID = GetActiveTraceRunID()
    If mHelpDocumentMacroBuildRepairBaselineNonHelpCompactStateByObject Is Nothing Then Err.Raise vbObjectError + 6818, "ManagedObjectDescriptionCanonicalTokenAfterScopedHelpDocumentMacroBuildRepair", "The pre-mutation non-Help compact-state cache is unavailable."
    If StrComp(mHelpDocumentMacroBuildRepairBaselineCompactStateTraceRunID, activeTraceRunID, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 6819, "ManagedObjectDescriptionCanonicalTokenAfterScopedHelpDocumentMacroBuildRepair", "The pre-mutation non-Help compact-state cache belongs to a different trace run."
    inputFingerprint = ManagedDescriptionInputFingerprint(db, False)
    payloadText = "ADCManagedObjectDescriptionPostCommitV1|" & inputFingerprint & "|"
    presentStatePrefix = FingerprintTokenPart("Present")
    sqlText = "SELECT m." & SQL_FIELD_OBJECT_NAME & ",m." & SQL_FIELD_OBJECT_TYPE & ",m." & SQL_FIELD_OBJECT_DESCRIPTION & "," & ResolvedExpectedExistencePolicySql("o", "c") & " " & "AS " & SQL_FIELD_RESOLVED_EXPECTED_EXISTENCE_POLICY & "," & _
        ResolvedSetupActionSql("o", "c") & " " & "AS " & SQL_FIELD_RESOLVED_SETUP_ACTION & " FROM ([" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID) & "] AS m INNER JOIN [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ObjectCategories") & "] AS c ON m.[" & FIELD_CATEGORY_CODE & "]=c.[" & FIELD_CATEGORY_CODE & "]) LEFT JOIN [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECT_OVERRIDES_TABLE_LOGICAL_ID) & "] AS o ON m." & SQL_FIELD_ID & "=o.[" & FIELD_MANAGED_OBJECT_ID & "] WHERE Nz(m.[" & FIELD_MANAGED_OBJECT_STATUS & "]," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & _
        " AND Nz(c." & SQL_FIELD_OBJECT_CATEGORY_STATUS & "," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & " ORDER BY " & OrderByPolicyText("REVIEW_MANAGED_DESCRIPTION") & ";"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    Do While rs.EOF = False
        checkedCount = checkedCount + 1
        objectName = CStr(Nz(rs.fields(FIELD_OBJECT_NAME).value, vbNullString))
        objectTypeCode = CStr(Nz(rs.fields(FIELD_OBJECT_TYPE).value, vbNullString))
        expectedExistencePolicy = CStr(Nz(rs.fields(FIELD_RESOLVED_EXPECTED_EXISTENCE_POLICY).value, vbNullString))
        setupAction = CStr(Nz(rs.fields(FIELD_RESOLVED_SETUP_ACTION).value, vbNullString))
        descriptionText = CStr(Nz(rs.fields(FIELD_OBJECT_DESCRIPTION).value, vbNullString))
        descriptionOptionalForHelpMacro = ManagedDescriptionIsOptionalHelpMacro(objectName, objectTypeCode)
        If Len(Trim$(descriptionText)) = 0 And descriptionOptionalForHelpMacro = False Then Err.Raise vbObjectError + 6503, "ManagedObjectDescriptionCanonicalTokenAfterScopedHelpDocumentMacroBuildRepair", "An active Managed Object Catalog description is blank outside the optional Help-document macro policy." & vbCrLf & "Object: " & objectTypeCode & " [" & objectName & "]"
        physicalDescriptionText = ManagedDescriptionExpectedPhysicalText(descriptionText)
        payloadText = payloadText & FingerprintTokenPart(objectTypeCode) & FingerprintTokenPart(objectName) & FingerprintTokenPart(descriptionText) & FingerprintTokenPart(physicalDescriptionText)
        identityKey = ManagedObjectDefinitionIdentityKey(objectName, objectTypeCode)
        If descriptionOptionalForHelpMacro Then
            targetExists = DatabaseAccessObjectExistsByName(db, acMacro, objectName)
        Else
            If mHelpDocumentMacroBuildRepairBaselineNonHelpCompactStateByObject.Exists(identityKey) = False Then Err.Raise vbObjectError + 6820, "ManagedObjectDescriptionCanonicalTokenAfterScopedHelpDocumentMacroBuildRepair", "A non-Help managed-object identity is not present in the pre-mutation compact-state proof." & vbCrLf & "Object: " & objectTypeCode & " [" & objectName & "]"
            baselineObjectState = CStr(mHelpDocumentMacroBuildRepairBaselineNonHelpCompactStateByObject(identityKey))
            targetExists = (StrComp(Left$(baselineObjectState, Len(presentStatePrefix)), presentStatePrefix, vbBinaryCompare) = 0)
        End If
        If targetExists Then
            payloadText = payloadText & "Present|"
        Else
            optionalReason = ManagedObjectOptionalMissingReason(expectedExistencePolicy, setupAction)
            If Len(optionalReason) = 0 Then Err.Raise vbObjectError + 6504, "ManagedObjectDescriptionCanonicalTokenAfterScopedHelpDocumentMacroBuildRepair", "A required active managed-object description target is missing while advancing the scoped Help proof." & vbCrLf & "Object: " & objectTypeCode & " [" & objectName & "]"
            payloadText = payloadText & FingerprintTokenPart("OptionalAbsent:" & optionalReason)
        End If
        If checkedCount = 1 Or checkedCount Mod 50 = 0 Then RequireSafeWorkContinue "ManagedObjectDescriptionCanonicalTokenAfterScopedHelpDocumentMacroBuildRepair", objectTypeCode & " [" & objectName & "]"
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    ManagedObjectDescriptionCanonicalTokenAfterScopedHelpDocumentMacroBuildRepair = FingerprintFromPayload(payloadText)
    EndManagedDescriptionOptionalHelpPolicyScope
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
    EndManagedDescriptionOptionalHelpPolicyScope
    On Error GoTo 0
    Err.Raise savedErrNumber, "ManagedObjectDescriptionCanonicalTokenAfterScopedHelpDocumentMacroBuildRepair", "Could not reconstruct the canonical full managed-object description token after scoped Help G1." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

' Purpose: Finalizes help document macro build repair full ribbon predecessor proof for current catalog token, scoped help proof token.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   currentCatalogToken (Input; String): Proof/version token identifying the exact current catalog state that was previously validated in this setup run.
'   scopedHelpProofToken (Input; String): Scope value limiting which scoped help proof token items or operations are included in this call.
'   proofToken (Output; String): Receives the resolved proof token produced from this procedure's validated inputs and current ADC state.
'   proofTraceRunID (Output; String): Receives the resolved proof trace run ID produced from this procedure's validated inputs and current ADC state.
'   detailText (Output; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
' Returns: Boolean: True on the function's affirmative/success branch for finalize help document macro build repair full ribbon predecessor proof; False on its explicit negative, unavailable, rejected, or
'     handled-failure branch.
'   cancellation, rejection, or not-applicable path.
Public Function FinalizeHelpDocumentMacroBuildRepairFullRibbonPredecessorProof(ByVal currentCatalogToken As String, ByVal scopedHelpProofToken As String, ByRef proofToken As String, ByRef proofTraceRunID As String, ByRef detailText As String) As Boolean
    ' Advance the pre-mutation complete predecessor proof through the only domains Build/Repair Help Document Macros is authorized to mutate. The returned token uses the normal ADCRibbonDependencyV3 formula and
    ' therefore has the same semantics as standalone/coordinated H1.
    On Error GoTo ErrHandler
    Dim activeTraceRunID As String
    Dim authoritativeHelpToken As String
    Dim authoritativeHelpTraceRunID As String
    Dim helpDetail As String
    Dim authoritativeCatalogToken As String
    Dim objectDescriptionToken As String
    Dim wholeCatalogToken As String
    Dim inputFingerprint As String
    Dim currentPhysicalState As String
    Dim compactPhysicalStateDetail As String
    Dim fieldProofDetail As String
    Dim canonicalPredecessorDetail As String
    Dim canonicalObjectDescriptionToken As String
    Dim canonicalFieldDescriptionToken As String
    Dim canonicalCatalogToken As String
    proofToken = vbNullString
    proofTraceRunID = vbNullString
    detailText = vbNullString
    activeTraceRunID = GetActiveTraceRunID()
    If Len(activeTraceRunID) = 0 Then Err.Raise vbObjectError + 6496, "FinalizeHelpDocumentMacroBuildRepairFullRibbonPredecessorProof", "No active trace run exists."
    If Len(mHelpDocumentMacroBuildRepairBaselineScopeProofToken) = 0 Or StrComp(mHelpDocumentMacroBuildRepairBaselineScopeProofTraceRunID, activeTraceRunID, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 6497, "FinalizeHelpDocumentMacroBuildRepairFullRibbonPredecessorProof", "The pre-mutation Build/Repair Help Document Macros unchanged-domain proof is unavailable for the current trace run."
    If Len(mHelpDocumentMacroBuildRepairBaselineFieldDescriptionToken) = 0 Then Err.Raise vbObjectError + 6505, "FinalizeHelpDocumentMacroBuildRepairFullRibbonPredecessorProof", "The pre-mutation complete proof did not retain the unchanged managed-table field-description token."
    If HelpDocumentDescriptionSameRunG1ProofForRibbon(authoritativeHelpToken, authoritativeHelpTraceRunID, helpDetail) = False Then Err.Raise vbObjectError + 6498, "FinalizeHelpDocumentMacroBuildRepairFullRibbonPredecessorProof", helpDetail
    If StrComp(authoritativeHelpTraceRunID, activeTraceRunID, vbBinaryCompare) <> 0 Or StrComp(Trim$(scopedHelpProofToken), authoritativeHelpToken, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 6499, "FinalizeHelpDocumentMacroBuildRepairFullRibbonPredecessorProof", "The supplied scoped Help proof does not match G1's authoritative same-run proof."
    authoritativeCatalogToken = Trim$(currentCatalogToken)
    If Len(authoritativeCatalogToken) = 0 Then Err.Raise vbObjectError + 6500, "FinalizeHelpDocumentMacroBuildRepairFullRibbonPredecessorProof", "The same-run current complete-catalog token was not supplied to the post-Help predecessor finalizer."
    CorePulseProgress "Build/Repair Help Document Macros predecessor: validating final compact physical state from cached non-Help state plus current Help metadata."
    currentPhysicalState = ManagedDescriptionCompactPhysicalStateAfterScopedHelpDocumentMacroBuildRepair(currentDb, compactPhysicalStateDetail)
    If Len(currentPhysicalState) = 0 Then Err.Raise vbObjectError + 6508, "FinalizeHelpDocumentMacroBuildRepairFullRibbonPredecessorProof", compactPhysicalStateDetail
    CorePulseProgress "Build/Repair Help Document Macros predecessor: reconstructing canonical managed-object description token from proven non-Help state plus scoped Help G1."
    objectDescriptionToken = ManagedObjectDescriptionCanonicalTokenAfterScopedHelpDocumentMacroBuildRepair(inputFingerprint)
    If Len(objectDescriptionToken) = 0 Then Err.Raise vbObjectError + 6506, "FinalizeHelpDocumentMacroBuildRepairFullRibbonPredecessorProof", "The canonical managed-object description token could not be reconstructed after scoped Help G1."
    CorePulseProgress "Build/Repair Help Document Macros predecessor: rebinding unchanged managed-table field-description proof."
    If RebindManagedTableFieldDescriptionProofAfterHelpDocumentMacroBuildRepair(mHelpDocumentMacroBuildRepairBaselineFieldDescriptionToken, authoritativeCatalogToken, activeTraceRunID, fieldProofDetail) = False Then Err.Raise vbObjectError + 6507, "FinalizeHelpDocumentMacroBuildRepairFullRibbonPredecessorProof", fieldProofDetail
    wholeCatalogToken = ManagedDescriptionWholeCatalogExactToken(objectDescriptionToken, mHelpDocumentMacroBuildRepairBaselineFieldDescriptionToken)
    Call EnsureDaoPropertyValue(currentDb, MANAGED_DESCRIPTION_FINGERPRINT_PROPERTY, dbText, inputFingerprint, "FinalizeHelpDocumentMacroBuildRepairFullRibbonPredecessorProof")
    Call EnsureDaoPropertyValue(currentDb, MANAGED_DESCRIPTION_PHYSICAL_STATE_PROPERTY, dbText, currentPhysicalState, "FinalizeHelpDocumentMacroBuildRepairFullRibbonPredecessorProof")
    Call EnsureDaoPropertyValue(currentDb, MANAGED_DESCRIPTION_EXACT_TOKEN_PROPERTY, dbText, objectDescriptionToken, "FinalizeHelpDocumentMacroBuildRepairFullRibbonPredecessorProof")
    Call EnsureDaoPropertyValue(currentDb, MANAGED_DESCRIPTION_WHOLE_CATALOG_EXACT_TOKEN_PROPERTY, dbText, wholeCatalogToken, "FinalizeHelpDocumentMacroBuildRepairFullRibbonPredecessorProof")
    mManagedObjectDescriptionToken = objectDescriptionToken
    mManagedObjectDescriptionTokenTraceRunID = activeTraceRunID
    CorePulseProgress "Build/Repair Help Document Macros predecessor: deriving canonical H1 predecessor token."
    proofToken = RibbonDependencyTokenForSetup(canonicalPredecessorDetail, canonicalObjectDescriptionToken, canonicalFieldDescriptionToken, canonicalCatalogToken, False, authoritativeCatalogToken, activeTraceRunID)
    If Len(proofToken) = 0 Then Err.Raise vbObjectError + 6509, "FinalizeHelpDocumentMacroBuildRepairFullRibbonPredecessorProof", canonicalPredecessorDetail
    If StrComp(canonicalObjectDescriptionToken, objectDescriptionToken, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 6510, "FinalizeHelpDocumentMacroBuildRepairFullRibbonPredecessorProof", "The canonical H1 predecessor authority resolved a different managed-object description token after scoped Help G1."
    If StrComp(canonicalFieldDescriptionToken, mHelpDocumentMacroBuildRepairBaselineFieldDescriptionToken, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 6511, "FinalizeHelpDocumentMacroBuildRepairFullRibbonPredecessorProof", "The canonical H1 predecessor authority resolved a different managed-table field-description token after the Help-only proof rebind."
    If StrComp(canonicalCatalogToken, authoritativeCatalogToken, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 6512, "FinalizeHelpDocumentMacroBuildRepairFullRibbonPredecessorProof", "The canonical H1 predecessor authority resolved a different complete-catalog token after scoped Help G1."
    mHelpDocumentMacroBuildRepairFullPredecessorToken = proofToken
    mHelpDocumentMacroBuildRepairFullPredecessorTraceRunID = activeTraceRunID
    mHelpDocumentMacroBuildRepairFullPredecessorCatalogToken = authoritativeCatalogToken
    proofTraceRunID = activeTraceRunID
    detailText = "Build/Repair Help Document Macros complete predecessor proof advanced without unrelated Description scans and was re-derived by the ordinary RibbonDependencyTokenForSetup authority." & vbCrLf & _
                 "Canonical managed-object token reconstructed from proven unchanged non-Help state plus scoped Help G1: True" & vbCrLf & _
                 "Managed-table field-description physical inspection after Help mutation: 0" & vbCrLf & _
                 "Canonical predecessor authority match: True" & vbCrLf & compactPhysicalStateDetail & vbCrLf & fieldProofDetail & vbCrLf & canonicalPredecessorDetail
    InvalidateHelpDocumentMacroBuildRepairBaselineCompactStateCache
    FinalizeHelpDocumentMacroBuildRepairFullRibbonPredecessorProof = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    mHelpDocumentMacroBuildRepairFullPredecessorToken = vbNullString
    mHelpDocumentMacroBuildRepairFullPredecessorTraceRunID = vbNullString
    mHelpDocumentMacroBuildRepairFullPredecessorCatalogToken = vbNullString
    InvalidateHelpDocumentMacroBuildRepairBaselineCompactStateCache
    If IsCancellationError(savedErrNumber) Then Err.Raise savedErrNumber, savedErrSource, savedErrDescription
    detailText = "Build/Repair Help Document Macros could not establish its complete post-mutation Ribbon predecessor proof. Error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription
    FinalizeHelpDocumentMacroBuildRepairFullRibbonPredecessorProof = False
End Function

' Purpose: Performs help document macro build repair same run full ribbon predecessor proof by coordinating `GetActiveTraceRunID` and `StrComp` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   proofToken (Output; String): Receives the resolved proof token produced from this procedure's validated inputs and current ADC state.
'   proofTraceRunID (Output; String): Receives the resolved proof trace run ID produced from this procedure's validated inputs and current ADC state.
'   detailText (Output; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
'   proofCatalogToken (Output Optional; String): Receives the resolved proof catalog token produced from this procedure's validated inputs and current ADC state. Default: vbNullString.
' Returns: Boolean: True when the condition represented by HelpDocumentMacroBuildRepairSameRunFullRibbonPredecessorProof is present under the routine's documented inputs/current state; False when that condition
'   is absent.
Public Function HelpDocumentMacroBuildRepairSameRunFullRibbonPredecessorProof(ByRef proofToken As String, ByRef proofTraceRunID As String, ByRef detailText As String, Optional ByRef proofCatalogToken As String = vbNullString) As Boolean
    ' H1 handoff for Build/Repair Help Document Macros. Only the complete same-run proof advanced by FinalizeHelpDocumentMacroBuildRepairFullRibbonPredecessorProof is exposed, including its exact CATALOG component
    ' token for reuse by Ribbon planning.
    Dim activeTraceRunID As String
    proofToken = vbNullString
    proofTraceRunID = vbNullString
    proofCatalogToken = vbNullString
    detailText = vbNullString
    activeTraceRunID = GetActiveTraceRunID()
    If Len(activeTraceRunID) = 0 Then detailText = "No active trace run exists for the Build/Repair Help Document Macros complete predecessor proof.": Exit Function
    If Len(mHelpDocumentMacroBuildRepairFullPredecessorToken) = 0 Or Len(mHelpDocumentMacroBuildRepairFullPredecessorTraceRunID) = 0 Or Len(mHelpDocumentMacroBuildRepairFullPredecessorCatalogToken) = 0 Then detailText = "No complete Build/Repair Help Document Macros predecessor proof has been established in this operation.": Exit Function
    If StrComp(mHelpDocumentMacroBuildRepairFullPredecessorTraceRunID, activeTraceRunID, vbBinaryCompare) <> 0 Then detailText = "The complete Build/Repair Help Document Macros predecessor proof belongs to a different trace run.": Exit Function
    proofToken = mHelpDocumentMacroBuildRepairFullPredecessorToken
    proofTraceRunID = mHelpDocumentMacroBuildRepairFullPredecessorTraceRunID
    proofCatalogToken = mHelpDocumentMacroBuildRepairFullPredecessorCatalogToken
    detailText = "Build/Repair Help Document Macros complete Ribbon predecessor proof=current for trace run " & activeTraceRunID & "."
    HelpDocumentMacroBuildRepairSameRunFullRibbonPredecessorProof = True
End Function

Private Sub TryPersistManagedDescriptionExactState(ByVal db As DAO.Database, ByVal inputFingerprint As String, ByVal exactToken As String, ByVal targetObjectCache As Object)
    On Error GoTo OptionalStateWarning
    Dim compactPhysicalState As String
    If targetObjectCache Is Nothing Then
        compactPhysicalState = ManagedDescriptionCompactPhysicalState(db)
    Else
        compactPhysicalState = ManagedDescriptionCompactPhysicalStateFromCache(targetObjectCache)
    End If
    Call EnsureDaoPropertyValue(db, MANAGED_DESCRIPTION_FINGERPRINT_PROPERTY, dbText, inputFingerprint, "TryPersistManagedDescriptionExactState")
    Call EnsureDaoPropertyValue(db, MANAGED_DESCRIPTION_PHYSICAL_STATE_PROPERTY, dbText, compactPhysicalState, "TryPersistManagedDescriptionExactState")
    Call EnsureDaoPropertyValue(db, MANAGED_DESCRIPTION_EXACT_TOKEN_PROPERTY, dbText, exactToken, "TryPersistManagedDescriptionExactState")
    Exit Sub
OptionalStateWarning:
    LogTraceMessage TRACE_LEVEL_DETAIL, "ADC_ManagedObjects", "TryPersistManagedDescriptionExactState", "Optional managed-description compact state could not be persisted; later setup will use authoritative verification.", "Err.Number: " & CStr(Err.Number) & vbCrLf & "Err.Source: " & Err.Source & vbCrLf & "Err.Description: " & Err.Description, "Persist managed-description compact state", Err.Number, Err.Source, Err.Description, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID), "Table", "CodeDb", "ManagedObjectDescriptions"
    Err.Clear
End Sub

Private Function SetAccessObjectDescription(ByVal targetObject As Object, ByVal objectName As String, ByVal objectTypeCode As String, ByVal authoritativeDescription As String, Optional ByVal allowBlankDescription As Boolean = False) As Boolean
    On Error GoTo ErrHandler
    Dim physicalDescription As String
    Dim existingPropertyValue As String
    Dim existingPropertyType As Long
    Dim propertyPreviouslyExisted As Boolean
    Dim operationType As String
    Dim wasBounded As Boolean
    Dim authoritativeCharacterLength As Long
    Dim authoritativeByteLength As Long
    Dim writtenCharacterLength As Long
    Dim writtenByteLength As Long
    Dim verifiedPropertyValue As String
    Dim verifiedPropertyType As Long
    If Len(authoritativeDescription) = 0 And allowBlankDescription = False Then Exit Function
    physicalDescription = ManagedDescriptionPropertySafeText(authoritativeDescription, wasBounded, authoritativeCharacterLength, authoritativeByteLength, writtenCharacterLength, writtenByteLength)
    operationType = "Inspect existing property"
    propertyPreviouslyExisted = ReadDaoTextPropertyState(targetObject, DAO_DESCRIPTION_PROPERTY_NAME, existingPropertyValue, existingPropertyType)
    If propertyPreviouslyExisted Then
        If existingPropertyType = CLng(dbText) And StrComp(existingPropertyValue, physicalDescription, vbBinaryCompare) = 0 Then Exit Function
        If existingPropertyType <> CLng(dbText) Then
            operationType = "Replace wrong-type property"
        Else
            operationType = "Update existing property"
        End If
    Else
        If Len(physicalDescription) = 0 And allowBlankDescription Then Exit Function
        operationType = "Create missing property"
    End If
    If SetupDomainMutationWritesAllowedForCurrentPart(SETUP_DOMAIN_G1_DESCRIPTIONS) = False Then Err.Raise vbObjectError + 9624, "SetAccessObjectDescription", "Managed DAO object Description mutation is owned exclusively by Setup G1." & vbCrLf & "Object type: " & objectTypeCode & vbCrLf & "Object name: " & objectName
    If EnsureDaoPropertyValue(targetObject, DAO_DESCRIPTION_PROPERTY_NAME, dbText, physicalDescription, "SetAccessObjectDescription") = False Then Err.Raise vbObjectError + 6455, "SetAccessObjectDescription", "The verified DAO property writer returned False after a required Description change."
    operationType = "Verify written property"
    If ReadDaoTextPropertyState(targetObject, DAO_DESCRIPTION_PROPERTY_NAME, verifiedPropertyValue, verifiedPropertyType) = False Then Err.Raise vbObjectError + 6456, "SetAccessObjectDescription", "The Description property was absent after the write completed."
    If verifiedPropertyType <> CLng(dbText) Then Err.Raise vbObjectError + 6457, "SetAccessObjectDescription", "The Description property type did not verify after the write." & vbCrLf & "Expected DAO type: " & CStr(dbText) & vbCrLf & "Actual DAO type: " & CStr(verifiedPropertyType)
    If StrComp(verifiedPropertyValue, physicalDescription, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 6458, "SetAccessObjectDescription", "The Description property value did not verify after the write."
    SetAccessObjectDescription = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim daoErrorsText As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    daoErrorsText = ReviewDBEngineErrorsText()
    Err.Raise savedErrNumber, "SetAccessObjectDescription", _
        "Could not ensure and verify the managed-object Access Description property." & vbCrLf & _
        "Object type: " & objectTypeCode & vbCrLf & _
        "Object name: " & objectName & vbCrLf & _
        "Property: " & DAO_DESCRIPTION_PROPERTY_NAME & vbCrLf & _
        "Operation: " & operationType & vbCrLf & _
        "Blank authoritative Description allowed: " & CStr(allowBlankDescription) & vbCrLf & _
        "Property previously existed: " & CStr(propertyPreviouslyExisted) & vbCrLf & _
        "Existing DAO type: " & CStr(existingPropertyType) & vbCrLf & _
        "Existing characters: " & CStr(Len(existingPropertyValue)) & vbCrLf & _
        "Existing bytes: " & CStr(LenB(existingPropertyValue)) & vbCrLf & _
        "Authoritative characters: " & CStr(authoritativeCharacterLength) & vbCrLf & _
        "Authoritative bytes: " & CStr(authoritativeByteLength) & vbCrLf & _
        "Written characters: " & CStr(writtenCharacterLength) & vbCrLf & _
        "Written bytes: " & CStr(writtenByteLength) & vbCrLf & _
        "Property-safe representation used: " & CStr(wasBounded) & vbCrLf & _
        "Err.Number: " & CStr(savedErrNumber) & vbCrLf & _
        "Err.Source: " & savedErrSource & vbCrLf & _
        "Err.Description: " & savedErrDescription & _
        IIf(Len(daoErrorsText) > 0, vbCrLf & "DBEngine.Errors:" & vbCrLf & daoErrorsText, vbNullString)
End Function

' =============================================================================================================================================================================
' DUPLICATE REVIEW PROCEDURES
' =============================================================================================================================================================================

' Purpose: Performs config row has duplicate for query by coordinating `SharedValue_NzLong` and `End` while preserving this module's existing validation and error semantics.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   rowID (Input; Variant): Persistent row identifier selecting the catalog, review, configuration, or history record affected by the operation.
' Returns: Boolean: True when the condition represented by DuplicateGroupKeyForQuery is present under the routine's documented inputs/current state; False when that condition is absent.
Public Function ConfigRowHasDuplicateForQuery(ByVal rowID As Variant) As Boolean
    'Intended use: Returns the Access Development Console Config Row Has Duplicate For Query result for direct, query, form, or required cross-module callers.
    On Error GoTo ErrHandler
    Dim resolvedRowID As Long
    resolvedRowID = SharedValue_NzLong(rowID)
    ConfigRowHasDuplicateForQuery = (Len(DuplicateIssueTypeForRowID(resolvedRowID)) > 0)
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ConfigRowHasDuplicateForQuery", "Could not evaluate duplicate status for the saved duplicate-review query." & vbCrLf & "Configuration row ID: " & CStr(resolvedRowID) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

' Purpose: Resolves and returns duplicate issue type for query from the supplied inputs and the authoritative ADC definitions/state used by this module.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   rowID (Input; Variant): Persistent row identifier selecting the catalog, review, configuration, or history record affected by the operation.
' Returns: String: Returns the canonical duplicate issue type for query value used by downstream policy/dispatch logic; vbNullString indicates that no managed value applies where the API permits absence.
Public Function DuplicateIssueTypeForQuery(ByVal rowID As Variant) As String
    'Intended use: Returns the Access Development Console Duplicate Issue Type For Query result for direct, query, form, or required cross-module callers.
    On Error GoTo ErrHandler
    Dim resolvedRowID As Long
    resolvedRowID = SharedValue_NzLong(rowID)
    DuplicateIssueTypeForQuery = DuplicateIssueTypeForRowID(resolvedRowID)
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "DuplicateIssueTypeForQuery", "Could not evaluate duplicate issue type for the saved duplicate-review query." & vbCrLf & "Configuration row ID: " & CStr(resolvedRowID) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

' Purpose: Resolves and returns duplicate issue detail for query from the supplied inputs and the authoritative ADC definitions/state used by this module.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   rowID (Input; Variant): Persistent row identifier selecting the catalog, review, configuration, or history record affected by the operation.
' Returns: String: Returns the resolved duplicate issue detail for query text in the canonical representation consumed by this module's caller; vbNullString represents no applicable value where the existing
'   contract permits it.
Public Function DuplicateIssueDetailForQuery(ByVal rowID As Variant) As String
    'Intended use: Returns the Access Development Console Duplicate Issue Detail For Query result for direct, query, form, or required cross-module callers.
    On Error GoTo ErrHandler
    Dim resolvedRowID As Long
    resolvedRowID = SharedValue_NzLong(rowID)
    DuplicateIssueDetailForQuery = DuplicateGroupKeyForRowID(resolvedRowID)
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "DuplicateIssueDetailForQuery", "Could not evaluate duplicate issue detail for the saved duplicate-review query." & vbCrLf & "Configuration row ID: " & CStr(resolvedRowID) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

' Purpose: Resolves and returns duplicate group key for query from the supplied inputs and the authoritative ADC definitions/state used by this module.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   rowID (Input; Variant): Persistent row identifier selecting the catalog, review, configuration, or history record affected by the operation.
' Returns: String: Returns the resolved duplicate group key for query text in the canonical representation consumed by this module's caller; vbNullString represents no applicable value where the existing
'   contract permits it.
Public Function DuplicateGroupKeyForQuery(ByVal rowID As Variant) As String
    'Intended use: Returns the Access Development Console Duplicate Group Key For Query result for direct, query, form, or required cross-module callers.
    On Error GoTo ErrHandler
    Dim resolvedRowID As Long
    resolvedRowID = SharedValue_NzLong(rowID)
    DuplicateGroupKeyForQuery = DuplicateGroupKeyForRowID(resolvedRowID)
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "DuplicateGroupKeyForQuery", "Could not evaluate duplicate group key for the saved duplicate-review query." & vbCrLf & "Configuration row ID: " & CStr(resolvedRowID) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

' Purpose: Resolves and returns duplicate group count for query from the supplied inputs and the authoritative ADC definitions/state used by this module.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   rowID (Input; Variant): Persistent row identifier selecting the catalog, review, configuration, or history record affected by the operation.
' Returns: Long: Returns the number of duplicate group count for query items matched, processed, selected, deleted, or reviewed by this operation; zero represents no qualifying items.
Public Function DuplicateGroupCountForQuery(ByVal rowID As Variant) As Long
    'Intended use: Returns the Access Development Console Duplicate Group Count For Query result for direct, query, form, or required cross-module callers.
    On Error GoTo ErrHandler
    Dim resolvedRowID As Long
    Dim groupKey As String
    resolvedRowID = SharedValue_NzLong(rowID)
    groupKey = DuplicateGroupKeyForRowID(resolvedRowID)
    DuplicateGroupCountForQuery = DuplicateGroupCountByKeyCore(groupKey, False)
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "DuplicateGroupCountForQuery", "Could not evaluate duplicate group count for the saved duplicate-review query." & vbCrLf & "Configuration row ID: " & CStr(resolvedRowID) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

' Purpose: Resolves and returns duplicate selected count for query from the supplied inputs and the authoritative ADC definitions/state used by this module.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   rowID (Input; Variant): Persistent row identifier selecting the catalog, review, configuration, or history record affected by the operation.
' Returns: Long: Returns the number of duplicate selected count for query items matched, processed, selected, deleted, or reviewed by this operation; zero represents no qualifying items.
Public Function DuplicateSelectedCountForQuery(ByVal rowID As Variant) As Long
    'Intended use: Returns the Access Development Console Duplicate Selected Count For Query result for direct, query, form, or required cross-module callers.
    On Error GoTo ErrHandler
    Dim resolvedRowID As Long
    Dim groupKey As String
    resolvedRowID = SharedValue_NzLong(rowID)
    groupKey = DuplicateGroupKeyForRowID(resolvedRowID)
    DuplicateSelectedCountForQuery = DuplicateGroupCountByKeyCore(groupKey, True)
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "DuplicateSelectedCountForQuery", "Could not evaluate selected duplicate count for the saved duplicate-review query." & vbCrLf & "Configuration row ID: " & CStr(resolvedRowID) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

' Purpose: Performs duplicate all rows selected for query by coordinating `DuplicateGroupCountForQuery` and `DuplicateSelectedCountForQuery` while preserving this module's existing validation and error semantics.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   rowID (Input; Variant): Persistent row identifier selecting the catalog, review, configuration, or history record affected by the operation.
' Returns: Boolean: True when the condition represented by DuplicateAllRowsSelectedForQuery is present under the routine's documented inputs/current state; False when that condition is absent.
Public Function DuplicateAllRowsSelectedForQuery(ByVal rowID As Variant) As Boolean
    'Intended use: Returns the Access Development Console Duplicate All Rows Selected For Query result for direct, query, form, or required cross-module callers.
    Dim groupCount As Long
    Dim selectedCount As Long
    groupCount = DuplicateGroupCountForQuery(rowID)
    selectedCount = DuplicateSelectedCountForQuery(rowID)
    DuplicateAllRowsSelectedForQuery = (groupCount > 0 And selectedCount >= groupCount)
End Function

' A3 SAVED-QUERY EXPRESSION DEPENDENCY CLOSURE
' MC_Records_DuplicateRows resolves these functions during query preflight/open, so this A3 dependency-closure block must remain available.

Private Sub EnsureDuplicateReviewCache()
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim cfg As TMacroRecord
    Dim issueType As String
    Dim groupKey As String
    Dim rowKey As String
    Dim keyText As String
    Dim configMacroKeys As Object
    Dim activeMacroTitles As Object
    If mDuplicateReviewCacheReady Then Exit Sub
    InvalidateDuplicateReviewCache
    Set mDuplicateReviewIssueTypeByRowID = CreateObject("Scripting.Dictionary")
    Set mDuplicateReviewGroupKeyByRowID = CreateObject("Scripting.Dictionary")
    Set mDuplicateReviewGroupCountByKey = CreateObject("Scripting.Dictionary")
    Set mDuplicateReviewSelectedCountByKey = CreateObject("Scripting.Dictionary")
    ' INTENTIONAL KEYED LOOKUP: these dictionaries are derived access paths over MC_Records configuration rows, not competing authorities.
    Set configMacroKeys = CreateObject("Scripting.Dictionary")
    Set activeMacroTitles = CreateObject("Scripting.Dictionary")
    mDuplicateReviewIssueTypeByRowID.CompareMode = vbBinaryCompare
    mDuplicateReviewGroupKeyByRowID.CompareMode = vbBinaryCompare
    mDuplicateReviewGroupCountByKey.CompareMode = vbBinaryCompare
    mDuplicateReviewSelectedCountByKey.CompareMode = vbBinaryCompare
    configMacroKeys.CompareMode = vbBinaryCompare
    activeMacroTitles.CompareMode = vbBinaryCompare
    Set db = ADCDb()
    Set rs = db.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "] ORDER BY " & SQL_FIELD_ID, dbOpenSnapshot)
    Do While Not rs.EOF
        PumpSetupResponsiveness
        MacroRecordFromRecordset rs, cfg
        keyText = NormalizedDuplicateTextKey(cfg.MacroKey, True)
        IncrementDuplicateCount configMacroKeys, keyText
        If cfg.Active Then IncrementDuplicateCount activeMacroTitles, NormalizedDuplicateTextKey(cfg.MacroTitle, True)
        rs.MoveNext
    Loop
    If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
    Do While Not rs.EOF
        PumpSetupResponsiveness
        MacroRecordFromRecordset rs, cfg
        issueType = DuplicateIssueTypeFromCountIndexes(cfg, configMacroKeys, activeMacroTitles)
        groupKey = DuplicateGroupKeyFromValues(cfg, issueType)
        rowKey = CStr(cfg.ID)
        mDuplicateReviewIssueTypeByRowID(rowKey) = issueType
        mDuplicateReviewGroupKeyByRowID(rowKey) = groupKey
        If Len(groupKey) > 0 Then
            IncrementDuplicateCount mDuplicateReviewGroupCountByKey, groupKey
            If cfg.SelectedForDuplicateDeletion Then IncrementDuplicateCount mDuplicateReviewSelectedCountByKey, groupKey
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    mDuplicateReviewCacheReady = True
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
    InvalidateDuplicateReviewCache
    Err.Raise savedErrNumber, "EnsureDuplicateReviewCache", "Could not build the configuration duplicate-review calculation cache." & vbCrLf & "Table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Function LookupDuplicateCount(ByRef counts As Object, ByVal keyText As String) As Long
    If counts Is Nothing Then Exit Function
    If Len(keyText) = 0 Then Exit Function
    If counts.Exists(keyText) Then LookupDuplicateCount = CLng(counts(keyText))
End Function

Private Function DuplicateIssueTypeFromCountIndexes(ByRef cfg As TMacroRecord, ByRef configMacroKeys As Object, ByRef activeMacroTitles As Object) As String
    Dim keyText As String
    keyText = NormalizedDuplicateTextKey(cfg.MacroKey, True)
    If LookupDuplicateCount(configMacroKeys, keyText) > 1 Then
        DuplicateIssueTypeFromCountIndexes = DUPLICATE_ISSUE_CONFIG_MACRO_KEY
        Exit Function
    End If
    If cfg.Active Then
        If LookupDuplicateCount(activeMacroTitles, NormalizedDuplicateTextKey(cfg.MacroTitle, True)) > 1 Then DuplicateIssueTypeFromCountIndexes = DUPLICATE_ISSUE_ACTIVE_MACRO_TITLE
    End If
End Function

Private Function DuplicateIssueTypeForRowID(ByVal rowID As Long) As String
    On Error GoTo ErrHandler
    Dim rowKey As String
    If rowID <= 0 Then Exit Function
    EnsureDuplicateReviewCache
    rowKey = CStr(rowID)
    If mDuplicateReviewIssueTypeByRowID.Exists(rowKey) = False Then Err.Raise vbObjectError + 6482, "DuplicateIssueTypeForRowID", "Config row could not be found in the duplicate-review cache. Row ID: " & CStr(rowID)
    DuplicateIssueTypeForRowID = CStr(mDuplicateReviewIssueTypeByRowID(rowKey))
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "DuplicateIssueTypeForRowID", "Could not evaluate duplicate issue type for " & DEFAULT_UTILITY_LONG_NAME & " configuration row ID " & CStr(rowID) & "." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function DuplicateGroupKeyForRowID(ByVal rowID As Long) As String
    On Error GoTo ErrHandler
    Dim rowKey As String
    If rowID <= 0 Then Exit Function
    EnsureDuplicateReviewCache
    rowKey = CStr(rowID)
    If mDuplicateReviewGroupKeyByRowID.Exists(rowKey) = False Then Err.Raise vbObjectError + 6483, "DuplicateGroupKeyForRowID", "Config row could not be found in the duplicate-review cache. Row ID: " & CStr(rowID)
    DuplicateGroupKeyForRowID = CStr(mDuplicateReviewGroupKeyByRowID(rowKey))
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "DuplicateGroupKeyForRowID", "Could not evaluate duplicate group key for " & DEFAULT_UTILITY_LONG_NAME & " configuration row ID " & CStr(rowID) & "." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function DuplicateGroupKeyFromValues(ByRef cfg As TMacroRecord, ByVal issueType As String) As String
    Dim valueKey As String
    Select Case issueType
        Case DUPLICATE_ISSUE_CONFIG_MACRO_KEY
            valueKey = NormalizedDuplicateTextKey(cfg.MacroKey, True)
        Case DUPLICATE_ISSUE_ACTIVE_MACRO_TITLE
            valueKey = NormalizedDuplicateTextKey(cfg.MacroTitle, True)
    End Select
    If Len(issueType) > 0 And Len(valueKey) > 0 Then DuplicateGroupKeyFromValues = CStr(Len(issueType)) & ":" & issueType & ";" & valueKey
End Function

Private Function DuplicateGroupCountByKeyCore(ByVal groupKey As String, ByVal selectedOnly As Boolean) As Long
    On Error GoTo ErrHandler
    If Len(Trim$(groupKey)) = 0 Then Exit Function
    EnsureDuplicateReviewCache
    If selectedOnly Then
        If mDuplicateReviewSelectedCountByKey.Exists(groupKey) Then DuplicateGroupCountByKeyCore = CLng(mDuplicateReviewSelectedCountByKey(groupKey))
    Else
        If mDuplicateReviewGroupCountByKey.Exists(groupKey) Then DuplicateGroupCountByKeyCore = CLng(mDuplicateReviewGroupCountByKey(groupKey))
    End If
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "DuplicateGroupCountByKeyCore", "Could not count rows in a duplicate-review group." & vbCrLf & "Group key: " & groupKey & vbCrLf & "Selected only: " & CStr(selectedOnly) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Sub IncrementDuplicateCount(ByRef counts As Object, ByVal keyText As String)
    If counts Is Nothing Then Err.Raise vbObjectError + 6484, "IncrementDuplicateCount", "Duplicate-count dictionary is Nothing."
    If Len(keyText) = 0 Then Exit Sub
    If counts.Exists(keyText) Then
        counts(keyText) = CLng(counts(keyText)) + 1
    Else
        counts.Add keyText, 1&
    End If
End Sub


Private Function ReadDaoTextPropertyState(ByVal targetObject As Object, ByVal propertyName As String, ByRef propertyValue As String, ByRef propertyType As Long) As Boolean
    On Error GoTo ErrHandler
    Dim prp As DAO.Property
    Dim lookupErrNumber As Long
    Dim lookupErrSource As String
    Dim lookupErrDescription As String
    On Error Resume Next
    Err.Clear
    Set prp = targetObject.properties(propertyName)
    lookupErrNumber = Err.Number
    lookupErrSource = Err.Source
    lookupErrDescription = Err.Description
    Err.Clear
    On Error GoTo ErrHandler
    If lookupErrNumber = 3270 Then Exit Function
    If lookupErrNumber <> 0 Then Err.Raise lookupErrNumber, lookupErrSource, lookupErrDescription
    If prp Is Nothing Then Err.Raise vbObjectError + 6453, "ReadDaoTextPropertyState", "The DAO property lookup returned Nothing without a missing-property error."
    propertyType = CLng(prp.Type)
    propertyValue = CStr(Nz(prp.value, vbNullString))
    ReadDaoTextPropertyState = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ReadDaoTextPropertyState", "Could not inspect an existing DAO text property." & vbCrLf & "Property: " & propertyName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function


Private Function ReviewRecordsetFieldValueIfExists(ByVal rs As DAO.Recordset, ByVal fieldName As String) As Variant
    On Error GoTo MissingField
    ReviewRecordsetFieldValueIfExists = rs.fields(fieldName).value
    Exit Function
MissingField:
    Err.Clear
    ReviewRecordsetFieldValueIfExists = Empty
End Function

Private Function ReviewOptionalRecordsetFieldMatches(ByVal rs As DAO.Recordset, ByVal fieldName As String, ByVal expectedValue As Variant) As Boolean
    Dim actualValue As Variant
    If IsEmpty(expectedValue) Then
        ReviewOptionalRecordsetFieldMatches = True
        Exit Function
    End If
    actualValue = ReviewRecordsetFieldValueIfExists(rs, fieldName)
    If IsEmpty(actualValue) Then Exit Function
    If IsNull(expectedValue) Then
        ReviewOptionalRecordsetFieldMatches = IsNull(actualValue)
    ElseIf IsNull(actualValue) = False Then
        ReviewOptionalRecordsetFieldMatches = (StrComp(CStr(actualValue), CStr(expectedValue), vbBinaryCompare) = 0)
    End If
End Function

Private Sub ReviewSetRecordsetFieldIfExists(ByVal rs As DAO.Recordset, ByVal fieldName As String, ByVal fieldValue As Variant)
    On Error GoTo MissingField
    rs.fields(fieldName).value = fieldValue
MissingField:
    Err.Clear
End Sub

' =============================================================================================================================================================================
' NAVIGATION PANE I1-I3 PROCEDURES
' =============================================================================================================================================================================

' Purpose: Ensures configured navigation pane groups exist for source name, fail on creation error.
' Public role: Public module entry point. This procedure is listed in the module header as an intentional externally callable ADC entry point and may be invoked directly by Access/VBA callers.
' Parameters:
'   sourceName (Input; String): Name of the source object/query/table/control from which this operation reads or to which the generated definition is bound.
'   failOnCreationError (Input Optional; Boolean): True treats creation error as a terminating error; False allows the routine's handled/non-fatal path for that condition. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub EnsureConfiguredNavigationPaneGroupsExist(ByVal sourceName As String, Optional ByVal failOnCreationError As Boolean = True)
    RequireSetupDomainMutationAllowed SETUP_DOMAIN_I1_NAVIGATION_GROUPS, "EnsureConfiguredNavigationPaneGroupsExist"
    'I1 owner: creates only missing configured Navigation Pane group containers and verifies each final physical identity without a duplicate whole-set rescan.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim physicalSnapshot As Object
    Dim proofRows As Collection
    Dim groupName As String
    Dim groupKey As String
    Dim groupID As Long
    Dim groupCount As Long
    Dim checkedCount As Long
    Dim existingCount As Long
    Dim createdCount As Long
    Dim failureCount As Long
    Dim failureDetails As String
    Set mI1NavigationPaneProofRows = Nothing
    mI1NavigationPaneProofTraceRunID = vbNullString
    mI1NavigationPaneProofHostDbName = vbNullString
    Set mI1NavigationPanePhysicalGroupSnapshot = Nothing
    sourceName = Trim$(sourceName)
    If Len(sourceName) = 0 Then sourceName = "I1"
    If SettingBool(SETTING_ADD_TO_NAV_GROUP, True) = False Then
        CoreAddOperationNote "I1 Navigation Pane group creation is disabled by " & SETTING_ADD_TO_NAV_GROUP & "."
        Exit Sub
    End If
    If TableCanOpen(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups")) = False Then Err.Raise vbObjectError + 8810, "EnsureConfiguredNavigationPaneGroupsExist", "I1 requires the authoritative UI_Groups table established by Phase B."
    If TableFieldExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups"), FIELD_NAVIGATION_PANE_GROUP_NAME) = False Or TableFieldExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups"), FIELD_CREATE_IF_MISSING) = False Then Err.Raise vbObjectError + 8811, "EnsureConfiguredNavigationPaneGroupsExist", "I1 requires NavigationPaneGroupName and CreateIfMissing fields in UI_Groups."
    Set physicalSnapshot = ReviewPhysicalNavigationPaneGroupSnapshot()
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: proofRows holds planned/recovery work or independently observed validation evidence for EnsureConfiguredNavigationPaneGroupsExist; keep it distinct from the source
    ' authority.
    Set proofRows = New Collection
    Set db = ADCDb()
    Set rs = db.OpenRecordset("SELECT [" & FIELD_NAVIGATION_PANE_GROUP_KEY & "],[" & FIELD_NAVIGATION_PANE_GROUP_NAME & "] FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups") & "] WHERE Nz(" & SQL_FIELD_LOOKUP_ACTIVE & ",True)<>False AND Nz([" & FIELD_CREATE_IF_MISSING & "],False)<>False AND Len(Trim(Nz([" & FIELD_NAVIGATION_PANE_GROUP_NAME & "],'')))>0 ORDER BY " & OrderByPolicyText("REVIEW_CONFIGURED_NAVIGATION_GROUP") & ";", dbOpenSnapshot)
    Do While rs.EOF = False
        PumpSetupResponsiveness
        checkedCount = checkedCount + 1
        groupKey = Trim$(CStr(Nz(rs.fields(FIELD_NAVIGATION_PANE_GROUP_KEY).value, vbNullString)))
        groupName = Trim$(CStr(rs.fields(FIELD_NAVIGATION_PANE_GROUP_NAME).value))
        groupID = 0
        groupCount = ReviewPhysicalNavigationPaneGroupFromSnapshot(physicalSnapshot, groupName, groupID)
        Select Case groupCount
            Case 0
                groupID = ReviewCreateCustomNavigationPaneGroup(groupName, sourceName & IIf(Len(groupKey) = 0, vbNullString, ":" & groupKey), physicalSnapshot)
                If groupID > 0 Then
                    createdCount = createdCount + 1
                Else
                    failureCount = failureCount + 1
                    failureDetails = SharedText_AppendNonBlankLine(failureDetails, "Could not create required group [" & groupName & "].")
                End If
            Case 1
                existingCount = existingCount + 1
            Case Else
                failureCount = failureCount + 1
                failureDetails = SharedText_AppendNonBlankLine(failureDetails, "Required group name is ambiguous: [" & groupName & "] resolves to " & CStr(groupCount) & " physical groups.")
        End Select
        If groupID > 0 Then proofRows.Add Array(groupName, groupID)
        RequireSafeWorkContinue "EnsureConfiguredNavigationPaneGroupsExist", "Navigation Pane group [" & groupName & "]"
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    CoreAddOperationNote "I1 Navigation Pane groups: checked=" & CStr(checkedCount) & ", existing=" & CStr(existingCount) & ", created=" & CStr(createdCount) & ", failures=" & CStr(failureCount) & "."
    If failureCount = 0 Then
        Set mI1NavigationPaneProofRows = proofRows
        mI1NavigationPaneProofTraceRunID = GetActiveTraceRunID()
        mI1NavigationPaneProofHostDbName = currentDb.Name
        Set mI1NavigationPanePhysicalGroupSnapshot = physicalSnapshot
    End If
    LogTraceMessage IIf(failureCount = 0, TRACE_LEVEL_INFO, TRACE_LEVEL_ERROR), "NavigationPane", "EnsureConfiguredNavigationPaneGroupsExist", "I1 configured Navigation Pane group creation completed.", "Checked=" & CStr(checkedCount) & "; Existing=" & CStr(existingCount) & "; Created=" & CStr(createdCount) & "; Failures=" & CStr(failureCount) & "; duplicate final whole-set group proof skipped=True; physical-group snapshot builds=" & CStr(mNavPanePhysicalGroupSnapshotBuildCount) & IIf(Len(failureDetails) = 0, vbNullString, vbCrLf & failureDetails), "I1 create configured groups", 0, vbNullString, vbNullString, vbNullString, "NavigationPaneGroup", "CodeDb", sourceName
    If failOnCreationError And failureCount > 0 Then Err.Raise vbObjectError + 8812, "EnsureConfiguredNavigationPaneGroupsExist", "I1 could not establish the complete configured Navigation Pane group set." & vbCrLf & failureDetails
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
    Set mI1NavigationPaneProofRows = Nothing
    Set mI1NavigationPanePhysicalGroupSnapshot = Nothing
    mI1NavigationPaneProofTraceRunID = vbNullString
    mI1NavigationPaneProofHostDbName = vbNullString
    On Error GoTo 0
    Err.Raise savedErrNumber, "EnsureConfiguredNavigationPaneGroupsExist", "I1 configured Navigation Pane group creation failed." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Function ReviewTryGetSameRunI1PhysicalGroupSnapshot(ByRef physicalSnapshot As Object, ByRef detailText As String) As Boolean
    Dim activeTraceRunID As String
    detailText = vbNullString
    Set physicalSnapshot = Nothing
    If mI1NavigationPaneProofRows Is Nothing Or mI1NavigationPanePhysicalGroupSnapshot Is Nothing Then detailText = "I1 same-run physical-group evidence is unavailable.": Exit Function
    activeTraceRunID = GetActiveTraceRunID()
    If Len(activeTraceRunID) = 0 Or StrComp(activeTraceRunID, mI1NavigationPaneProofTraceRunID, vbBinaryCompare) <> 0 Then detailText = "I1 physical-group evidence belongs to a different or inactive setup trace run.": Exit Function
    If StrComp(currentDb.Name, mI1NavigationPaneProofHostDbName, vbTextCompare) <> 0 Then detailText = "I1 physical-group evidence belongs to a different HostDb.": Exit Function
    Set physicalSnapshot = mI1NavigationPanePhysicalGroupSnapshot
    ReviewTryGetSameRunI1PhysicalGroupSnapshot = True
    detailText = "I1 final physical Navigation Pane group evidence reused by I2 without a duplicate MSysNavPaneGroups snapshot."
End Function

' Purpose: Checks whether configured navigation pane groups already satisfies the exact state required by the current Setup run so the owning phase can determine whether repair is necessary.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   detailText (Output; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
'   allowSameRunProofReuse (Input Optional; Boolean): When True, enables or requests same run proof reuse for this call; when False, the routine follows its normal non-forced/default path. Default: False.
' Returns: Boolean result of configured navigation pane groups current for setup; True and False correspond to the affirmative and negative branches explicitly implemented by this function.
'   not-applicable path.
Public Function ConfiguredNavigationPaneGroupsCurrentForSetup(ByRef detailText As String, Optional ByVal allowSameRunProofReuse As Boolean = False) As Boolean
    'Read-only I1 proof. Every active CreateIfMissing configured group must resolve to exactly one physical custom group from one bulk physical snapshot.
    On Error GoTo ErrHandler
    Dim physicalSnapshot As Object
    Dim settingCache As Object
    Dim groupKeyCache As Object
    Dim configuredCurrent As Boolean
    detailText = vbNullString
    If SettingBool(SETTING_ADD_TO_NAV_GROUP, True) = False Then
        detailText = "Navigation Pane grouping is disabled; I1 group creation is a verified no-op."
        ConfiguredNavigationPaneGroupsCurrentForSetup = True
        Exit Function
    End If
    If allowSameRunProofReuse Then
        If ReviewSameRunI1GroupProofCurrent(detailText) Then
            ConfiguredNavigationPaneGroupsCurrentForSetup = True
            Exit Function
        End If
    End If
    If TableCanOpen(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups")) = False Then
        detailText = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups") & " is missing or cannot be opened."
        Exit Function
    End If
    Set physicalSnapshot = ReviewPhysicalNavigationPaneGroupSnapshot()
    ReviewNavigationPaneGroupDefinitionCaches physicalSnapshot, settingCache, groupKeyCache, configuredCurrent, detailText
    ConfiguredNavigationPaneGroupsCurrentForSetup = configuredCurrent
    Exit Function
ErrHandler:
    detailText = "I1 configured-group proof failed. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    ConfiguredNavigationPaneGroupsCurrentForSetup = False
End Function

Private Function ResolveOrCreateNavigationPaneGroupID(ByVal groupName As String, ByVal SettingKey As String, ByVal sourceName As String, Optional ByRef failureDetail As String) As Long
    'Canonical physical group resolver/creator used by runtime placement policy without forcing the full Setup I1 workflow.
    On Error GoTo ErrHandler
    Dim physicalSnapshot As Object
    Dim groupID As Long
    Dim groupCount As Long
    groupName = Trim$(groupName)
    SettingKey = Trim$(SettingKey)
    sourceName = Trim$(sourceName)
    failureDetail = vbNullString
    If Len(groupName) = 0 Then Exit Function
    Set physicalSnapshot = ReviewPhysicalNavigationPaneGroupSnapshot()
    groupCount = ReviewPhysicalNavigationPaneGroupFromSnapshot(physicalSnapshot, groupName, groupID)
    Select Case groupCount
        Case 0
            groupID = ReviewCreateCustomNavigationPaneGroup(groupName, sourceName, physicalSnapshot, False)
        Case 1
            'Use the uniquely resolved physical group.
        Case Else
            failureDetail = "Group name is ambiguous and resolves to " & CStr(groupCount) & " physical Navigation Pane groups. Setting=" & SettingKey
            Exit Function
    End Select
    If groupID <= 0 Then
        failureDetail = "Group was missing and remained unresolved after the create attempt. Setting=" & SettingKey
        Exit Function
    End If
    ResolveOrCreateNavigationPaneGroupID = groupID
    Exit Function
ErrHandler:
    failureDetail = "Setting=" & SettingKey & "; Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    ResolveOrCreateNavigationPaneGroupID = 0
End Function

Private Function ApplyDeferredNavigationPaneObjectPlacement(ByVal objectName As String, ByVal objectType As AcObjectType, ByVal groupName As String, ByVal sourceName As String, ByVal forceRefreshShortcut As Boolean, ByVal removeOtherGroupShortcuts As Boolean, ByRef groupID As Long, ByRef objectID As Long, ByRef detailText As String, Optional ByRef membershipInsertedResult As Long = 0, Optional ByRef otherGroupRemovedResult As Long = 0, Optional ByRef forcedGroupRemovedResult As Long = 0, Optional ByRef refreshCallCountResult As Long = 0, Optional ByVal deferNavigationPaneRefresh As Boolean = False) As Boolean
    'Physical placement writer; this procedure owns group resolution/creation, mirror identity materialization, shortcut cleanup, membership insertion, and physical verification. By default it refreshes once when
    'this call changes state; a batch owner may explicitly defer that UI refresh and perform one refresh after the batch.
    On Error GoTo ErrHandler
    Dim createdByThisCall As Boolean
    Dim operationDetail As String
    groupName = Trim$(groupName)
    sourceName = Trim$(sourceName)
    groupID = 0
    objectID = 0
    detailText = vbNullString
    membershipInsertedResult = 0
    otherGroupRemovedResult = 0
    forcedGroupRemovedResult = 0
    refreshCallCountResult = 0
    If Len(objectName) = 0 Then
        detailText = "No object name was supplied for physical Navigation Pane placement."
        Exit Function
    End If
    If Len(groupName) = 0 Then
        detailText = "No Navigation Pane group name was supplied for physical placement of [" & objectName & "]."
        Exit Function
    End If
    groupID = ResolveOrCreateNavigationPaneGroupID(groupName, vbNullString, sourceName, operationDetail)
    If groupID <= 0 Then
        detailText = operationDetail
        Exit Function
    End If
    objectID = SharedNavPane_MSysObjectID(currentDb, objectName, objectType)
    If objectID = 0 Then
        SharedNavPane_EnsureObjectIDRow objectName, objectType, "ADC_NavPane_"
        objectID = SharedNavPane_MSysObjectID(currentDb, objectName, objectType)
    End If
    If objectID = 0 Then
        detailText = "Could not resolve the Access object ID required for Navigation Pane placement. Object=[" & objectName & "]."
        Exit Function
    End If
    If removeOtherGroupShortcuts Then
        operationDetail = vbNullString
        otherGroupRemovedResult = ReviewDeleteNavigationPaneGroupObjectShortcutsExcept(groupID, objectID, objectName, operationDetail)
        detailText = SharedText_AppendNonBlankLine(detailText, operationDetail)
        If Len(Trim$(operationDetail)) > 0 Then Exit Function
    End If
    If forceRefreshShortcut Then
        operationDetail = vbNullString
        forcedGroupRemovedResult = ReviewDeleteNavigationPaneGroupObjectShortcuts(groupID, objectID, objectName, operationDetail)
        detailText = SharedText_AppendNonBlankLine(detailText, operationDetail)
        If Len(Trim$(operationDetail)) > 0 Then Exit Function
    End If
    operationDetail = vbNullString
    createdByThisCall = False
    If ReviewEnsureNavigationPaneMembership(currentDb, groupID, objectID, objectName, objectType, createdByThisCall, operationDetail) = False Then
        detailText = SharedText_AppendNonBlankLine(detailText, operationDetail)
        Exit Function
    End If
    If ReviewNavigationPaneMembershipExists(currentDb, groupID, objectID) = False Then
        detailText = SharedText_AppendNonBlankLine(detailText, "The required GroupID/ObjectID membership was not present after physical Navigation Pane placement.")
        Exit Function
    End If
    If createdByThisCall Then membershipInsertedResult = 1
    If createdByThisCall Or forceRefreshShortcut Then TrackSetupNavPaneShortcutRollback groupID, objectID
    If membershipInsertedResult + otherGroupRemovedResult + forcedGroupRemovedResult > 0 And deferNavigationPaneRefresh = False Then
        ReviewRefreshNavigationPaneOnce
        refreshCallCountResult = 1
    End If
    ApplyDeferredNavigationPaneObjectPlacement = True
    Exit Function
ErrHandler:
    detailText = SharedText_AppendNonBlankLine(detailText, "Physical Navigation Pane placement failed. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description)
    ApplyDeferredNavigationPaneObjectPlacement = False
End Function

Private Function ReviewApplySystemMacroNavigationPanePlacement(ByVal objectName As String, ByVal systemMacroGroupCache As Object, ByVal sourceName As String, ByRef detailText As String, Optional ByRef membershipInsertedResult As Long = 0, Optional ByRef otherGroupRemovedResult As Long = 0, Optional ByRef forcedGroupRemovedResult As Long = 0, Optional ByRef refreshCallCountResult As Long = 0) As Boolean
    Dim groupData As Variant
    Dim groupName As String
    Dim groupID As Long
    Dim objectID As Long
    Dim placementDetail As String
    membershipInsertedResult = 0
    otherGroupRemovedResult = 0
    forcedGroupRemovedResult = 0
    refreshCallCountResult = 0
    If Len(objectName) = 0 Then detailText = "A blank System Macro and Help Document Macro name cannot be placed in the Navigation Pane.": Exit Function
    If systemMacroGroupCache Is Nothing Then detailText = "The authoritative System Macro and Help Document Macro Navigation Pane group cache is unavailable.": Exit Function
    If systemMacroGroupCache.Exists(objectName) = False Then detailText = "System Macro and Help Document Macro [" & objectName & "] has no authoritative Navigation Pane group mapping.": Exit Function
    groupData = systemMacroGroupCache(objectName)
    If CLng(groupData(0)) <> 1 Or Len(Trim$(CStr(groupData(1)))) = 0 Or CLng(groupData(2)) <> 1 Then
        detailText = "System Macro and Help Document Macro [" & objectName & "] does not resolve to exactly one authoritative Navigation Pane group definition."
        Exit Function
    End If
    groupName = Trim$(CStr(groupData(3)))
    If Len(groupName) = 0 Or CLng(groupData(4)) <> 1 Or CLng(groupData(5)) <= 0 Then
        detailText = "System Macro and Help Document Macro [" & objectName & "] destination Navigation Pane group is not uniquely present physically."
        Exit Function
    End If
    If ApplyDeferredNavigationPaneObjectPlacement(objectName, acMacro, groupName, sourceName, False, True, groupID, objectID, placementDetail, membershipInsertedResult, otherGroupRemovedResult, forcedGroupRemovedResult, refreshCallCountResult, True) = False Then
        detailText = placementDetail
        Exit Function
    End If
    ReviewApplySystemMacroNavigationPanePlacement = True
End Function

' Purpose: Reconciles help document macro navigation pane.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   detailText (Output Optional; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller. Default: vbNullString.
' Returns: Boolean: True on the function's affirmative/success branch for reconcile help document macro navigation pane; False on its explicit negative, unavailable, rejected, or handled-failure branch.
'   not-applicable path.
Public Function ReconcileHelpDocumentMacroNavigationPane(Optional ByRef detailText As String = vbNullString) As Boolean
    'Operational entry point: reconciles active Help Document shortcuts plus the Build/Repair Help Document Macros command, and removes retired ADC-owned Help Document shortcuts.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim physicalGroupSnapshot As Object
    Dim settingCache As Object
    Dim groupKeyCache As Object
    Dim systemMacroGroupCache As Object
    Dim activeHelpNames As Object
    Dim staleHelpNames As Object
    Dim groupDetail As String
    Dim sqlText As String
    Dim objectName As String
    Dim objectStatus As String
    Dim runCodeWrapperName As String
    Dim buildRepairCommandTitle As String
    Dim keyValue As Variant
    Dim configuredCurrent As Boolean
    Dim activeReconciledCount As Long
    Dim staleShortcutRemovedCount As Long
    Dim staleShortcutPreservedCount As Long
    Dim objectID As Long
    Dim cleanupDetail As String
    Dim activeExactCount As Long
    Dim membershipInsertedCount As Long
    Dim wrongGroupShortcutRemovedCount As Long
    Dim forcedShortcutRemovedCount As Long
    Dim physicalMutationCount As Long
    Dim navigationPaneRefreshCount As Long
    Dim placementInsertedCount As Long
    Dim placementOtherRemovedCount As Long
    Dim placementForcedRemovedCount As Long
    Dim placementRefreshCount As Long
    ' INTENTIONAL KEYED LOOKUP: activeHelpNames, staleHelpNames provides keyed membership/lookup over source rows or state already read by ReconcileHelpDocumentMacroNavigationPane; it is a derived access path, not a
    ' second authority.
    Set activeHelpNames = CreateObject("Scripting.Dictionary")
    activeHelpNames.CompareMode = vbTextCompare
    Set staleHelpNames = CreateObject("Scripting.Dictionary")
    staleHelpNames.CompareMode = vbTextCompare
    Set db = ADCDb()
    sqlText = "SELECT m." & SQL_FIELD_ID & ",m." & SQL_FIELD_OBJECT_NAME & ",m.[" & FIELD_MANAGED_OBJECT_STATUS & "],p.[" & FIELD_PAYLOAD_KIND & "],p.[" & FIELD_RUN_CODE_WRAPPER_NAME & "] " & _
              "FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID) & "] AS m INNER JOIN [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectActionPayloads") & "] AS p ON m." & SQL_FIELD_ID & "=p.[" & FIELD_MANAGED_OBJECT_ID & "] " & _
              "WHERE UCase(Trim(Nz(m." & SQL_FIELD_OBJECT_TYPE & ",'')))=" & SharedDao_SqlQuote(UCase$(ADCAT_OBJECT_TYPE_MACRO)) & " AND UCase(Trim(Nz(m.[" & FIELD_CATEGORY_CODE & "],'')))=" & SharedDao_SqlQuote(UCase$(ADCAT_CATEGORY_SYSTEM_MACRO)) & " AND UCase(Trim(Nz(m.[" & FIELD_CATALOG_OWNER & "],'')))=" & SharedDao_SqlQuote(UCase$(ADCAT_OWNER_SYSTEM_DEFAULT)) & " " & _
              "ORDER BY m." & SQL_FIELD_ID & ",p." & SQL_FIELD_ID & ";"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    Do While rs.EOF = False
        objectName = Trim$(CStr(Nz(rs.fields(FIELD_OBJECT_NAME).value, vbNullString)))
        objectStatus = Trim$(CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_STATUS).value, ADCAT_STATUS_ACTIVE)))
        runCodeWrapperName = Trim$(CStr(Nz(rs.fields(FIELD_RUN_CODE_WRAPPER_NAME).value, vbNullString)))
        If StrComp(Trim$(CStr(Nz(rs.fields(FIELD_PAYLOAD_KIND).value, vbNullString))), ADCAT_PAYLOAD_ACCESS_MACRO_RUNCODE, vbTextCompare) = 0 Then
            If StrComp(Left$(runCodeWrapperName, Len("Run_OpenHelpDocument(")), "Run_OpenHelpDocument(", vbTextCompare) = 0 Then
                If StrComp(objectStatus, ADCAT_STATUS_ACTIVE, vbTextCompare) = 0 Then
                    If activeHelpNames.Exists(objectName) = False Then activeHelpNames.Add objectName, True
                Else
                    If staleHelpNames.Exists(objectName) = False Then staleHelpNames.Add objectName, True
                End If
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    If SettingBool(SETTING_ADD_TO_NAV_GROUP, True) Then
        EnsureConfiguredNavigationPaneGroupsExist "Build/Repair Help Document Macros", True
        Set physicalGroupSnapshot = ReviewPhysicalNavigationPaneGroupSnapshot()
        ReviewNavigationPaneGroupDefinitionCaches physicalGroupSnapshot, settingCache, groupKeyCache, configuredCurrent, groupDetail
        If configuredCurrent = False Then Err.Raise vbObjectError + 8890, "ReconcileHelpDocumentMacroNavigationPane", "Configured Navigation Pane groups are not current." & vbCrLf & groupDetail
        Set systemMacroGroupCache = ReviewSystemMacroNavigationPaneGroupCache(groupKeyCache)
        For Each keyValue In activeHelpNames.keys
            objectName = CStr(keyValue)
            cleanupDetail = vbNullString
            placementInsertedCount = 0
            placementOtherRemovedCount = 0
            placementForcedRemovedCount = 0
            placementRefreshCount = 0
            If ReviewApplySystemMacroNavigationPanePlacement(objectName, systemMacroGroupCache, "Build/Repair Help Document Macros", cleanupDetail, placementInsertedCount, placementOtherRemovedCount, placementForcedRemovedCount, placementRefreshCount) = False Then Err.Raise vbObjectError + 8891, "ReconcileHelpDocumentMacroNavigationPane", cleanupDetail
            activeReconciledCount = activeReconciledCount + 1
            membershipInsertedCount = membershipInsertedCount + placementInsertedCount
            wrongGroupShortcutRemovedCount = wrongGroupShortcutRemovedCount + placementOtherRemovedCount
            forcedShortcutRemovedCount = forcedShortcutRemovedCount + placementForcedRemovedCount
            physicalMutationCount = physicalMutationCount + placementInsertedCount + placementOtherRemovedCount + placementForcedRemovedCount
            navigationPaneRefreshCount = navigationPaneRefreshCount + placementRefreshCount
            If placementInsertedCount + placementOtherRemovedCount + placementForcedRemovedCount = 0 Then activeExactCount = activeExactCount + 1
        Next keyValue
        buildRepairCommandTitle = SystemMacroCatalogTitleByKey(SYS_MACRO_KEY_BUILD_REPAIR_HELP_DOCUMENT_MACROS)
        If Len(buildRepairCommandTitle) = 0 Then Err.Raise vbObjectError + 8892, "ReconcileHelpDocumentMacroNavigationPane", "The Build/Repair Help Document Macros system command title could not be resolved from the authoritative [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SystemMacros") & "] catalog."
        cleanupDetail = vbNullString
        placementInsertedCount = 0
        placementOtherRemovedCount = 0
        placementForcedRemovedCount = 0
        placementRefreshCount = 0
        If ReviewApplySystemMacroNavigationPanePlacement(buildRepairCommandTitle, systemMacroGroupCache, "Build/Repair Help Document Macros", cleanupDetail, placementInsertedCount, placementOtherRemovedCount, placementForcedRemovedCount, placementRefreshCount) = False Then Err.Raise vbObjectError + 8893, "ReconcileHelpDocumentMacroNavigationPane", cleanupDetail
        membershipInsertedCount = membershipInsertedCount + placementInsertedCount
        wrongGroupShortcutRemovedCount = wrongGroupShortcutRemovedCount + placementOtherRemovedCount
        forcedShortcutRemovedCount = forcedShortcutRemovedCount + placementForcedRemovedCount
        physicalMutationCount = physicalMutationCount + placementInsertedCount + placementOtherRemovedCount + placementForcedRemovedCount
        navigationPaneRefreshCount = navigationPaneRefreshCount + placementRefreshCount
    Else
        groupDetail = "Navigation Pane placement is disabled by the existing AddToNavigationPaneGroup setting; active Help shortcuts were not added."
    End If
    For Each keyValue In staleHelpNames.keys
        objectName = CStr(keyValue)
        If AccessObjectExistsInCurrentDatabase(acMacro, objectName) Then
            'C3 preserved a retained-name physical macro because it was not proven obsolete ADC-owned state; preserve its shortcuts too.
            staleShortcutPreservedCount = staleShortcutPreservedCount + 1
        Else
            objectID = SharedNavPane_MSysObjectID(currentDb, objectName, acMacro)
            cleanupDetail = vbNullString
            If objectID <> 0 Then
                staleShortcutRemovedCount = staleShortcutRemovedCount + ReviewDeleteNavigationPaneObjectShortcutsEverywhere(objectID, vbNullString, cleanupDetail)
            Else
                staleShortcutRemovedCount = staleShortcutRemovedCount + ReviewDeleteNavigationPaneObjectShortcutsEverywhere(0, objectName, cleanupDetail)
            End If
            If Len(cleanupDetail) > 0 Then Err.Raise vbObjectError + 8894, "ReconcileHelpDocumentMacroNavigationPane", cleanupDetail
        End If
    Next keyValue
    If staleShortcutRemovedCount > 0 Then physicalMutationCount = physicalMutationCount + staleShortcutRemovedCount
    If physicalMutationCount > 0 Then
        ReviewRefreshNavigationPaneOnce
        navigationPaneRefreshCount = navigationPaneRefreshCount + 1
    End If
    detailText = "Help Document Navigation Pane synchronization completed." & vbCrLf & _
                 "Active Help identities checked: " & CStr(activeHelpNames.count) & vbCrLf & _
                 "Active Help shortcuts reconciled: " & CStr(activeReconciledCount) & vbCrLf & _
                 "Active Help memberships already exact: " & CStr(activeExactCount) & vbCrLf & _
                 "Memberships inserted: " & CStr(membershipInsertedCount) & vbCrLf & _
                 "Wrong-group shortcut memberships removed: " & CStr(wrongGroupShortcutRemovedCount) & vbCrLf & _
                 "Forced same-group shortcut memberships removed: " & CStr(forcedShortcutRemovedCount) & vbCrLf & _
                 "Retired Help identities checked: " & CStr(staleHelpNames.count) & vbCrLf & _
                 "Obsolete shortcut memberships removed: " & CStr(staleShortcutRemovedCount) & vbCrLf & _
                 "Retained-name physical macros preserved: " & CStr(staleShortcutPreservedCount) & vbCrLf & _
                 "Physical Navigation Pane mutations: " & CStr(physicalMutationCount) & vbCrLf & _
                 "Navigation Pane refresh calls: " & CStr(navigationPaneRefreshCount)
    If Len(groupDetail) > 0 Then detailText = detailText & vbCrLf & groupDetail
    LogTraceMessage TRACE_LEVEL_INFO, "NavigationPane", "ReconcileHelpDocumentMacroNavigationPane", "Operational Help Document Navigation Pane synchronization completed.", detailText, "Build/Repair Help Document Macros Navigation Pane", 0, vbNullString, vbNullString, vbNullString, "NavigationPaneHelpDocuments", "CodeDb", "D"
    ReconcileHelpDocumentMacroNavigationPane = True
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
    detailText = "Help Document Navigation Pane synchronization failed." & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
    LogTraceMessage TRACE_LEVEL_ERROR, "NavigationPane", "ReconcileHelpDocumentMacroNavigationPane", "Operational Help Document Navigation Pane synchronization failed.", detailText, "Build/Repair Help Document Macros Navigation Pane", savedErrNumber, savedErrSource, savedErrDescription, vbNullString, "NavigationPaneHelpDocuments", "CodeDb", "D"
    Err.Raise savedErrNumber, "ReconcileHelpDocumentMacroNavigationPane", detailText
End Function

Private Function ReviewNavigationPaneObjectShortcutMatchCriteria(ByVal objectID As Long, ByVal objectName As String) As String
    'Single D-owned matching rule for physical Navigation Pane shortcut cleanup by ObjectID and retained name identity.
    Dim matchCriteria As String
    If objectID <> 0 Then matchCriteria = "[ObjectID]=" & CStr(objectID)
    If Len(objectName) > 0 Then
        If SharedDao_TableFieldExists(currentDb, NAV_PANE_GROUP_OBJECTS_TABLE, "Name") Then
            If Len(matchCriteria) > 0 Then matchCriteria = matchCriteria & " OR "
            matchCriteria = matchCriteria & "UCase(Nz([Name],''))=" & SharedDao_SqlQuote(UCase$(objectName))
        End If
    End If
    ReviewNavigationPaneObjectShortcutMatchCriteria = matchCriteria
End Function

Private Function ReviewDeleteNavigationPaneGroupObjectShortcutsExcept(ByVal keepGroupID As Long, ByVal objectID As Long, ByVal objectName As String, ByRef warningDetail As String) As Long
    'Physical backup-placement cleanup: removes matching shortcuts from every group except the selected backup group and returns the actual row count without refreshing the Navigation Pane itself.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim sqlText As String
    Dim matchCriteria As String
    Dim cleanedObjectName As String
    warningDetail = vbNullString
    cleanedObjectName = objectName
    If keepGroupID <= 0 Then Exit Function
    matchCriteria = ReviewNavigationPaneObjectShortcutMatchCriteria(objectID, cleanedObjectName)
    If Len(matchCriteria) = 0 Then Exit Function
    Set db = currentDb
    sqlText = "DELETE FROM [" & NAV_PANE_GROUP_OBJECTS_TABLE & "] WHERE [GroupID]<>" & CStr(keepGroupID) & " AND (" & matchCriteria & ");"
    db.Execute sqlText, dbFailOnError
    ReviewDeleteNavigationPaneGroupObjectShortcutsExcept = db.RecordsAffected
    Exit Function
ErrHandler:
    warningDetail = "Could not remove backup Navigation Pane shortcut rows from non-backup groups for [" & cleanedObjectName & "]. BackupGroupID=" & CStr(keepGroupID) & "; ObjectID=" & CStr(objectID) & "; Err.Number=" & CStr(Err.Number) & "; Err.Source=" & Err.Source & "; Err.Description=" & Err.Description
    Err.Clear
End Function

Private Function ReviewDeleteNavigationPaneGroupObjectShortcuts(ByVal groupID As Long, ByVal objectID As Long, ByVal objectName As String, ByRef warningDetail As String) As Long
    'Physical stale-shortcut cleanup before a forced runtime regrouping operation; returns the actual row count so the placement owner can coalesce refresh work.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim sqlText As String
    Dim matchCriteria As String
    Dim cleanedObjectName As String
    warningDetail = vbNullString
    cleanedObjectName = objectName
    If groupID <= 0 Then Exit Function
    matchCriteria = ReviewNavigationPaneObjectShortcutMatchCriteria(objectID, cleanedObjectName)
    If Len(matchCriteria) = 0 Then Exit Function
    Set db = currentDb
    sqlText = "DELETE FROM [" & NAV_PANE_GROUP_OBJECTS_TABLE & "] WHERE [GroupID]=" & CStr(groupID) & " AND (" & matchCriteria & ");"
    db.Execute sqlText, dbFailOnError
    ReviewDeleteNavigationPaneGroupObjectShortcuts = db.RecordsAffected
    Exit Function
ErrHandler:
    warningDetail = "Could not remove stale Navigation Pane shortcut row before regrouping [" & cleanedObjectName & "]. GroupID=" & CStr(groupID) & "; ObjectID=" & CStr(objectID) & "; Err.Number=" & CStr(Err.Number) & "; Err.Source=" & Err.Source & "; Err.Description=" & Err.Description
    Err.Clear
End Function

Private Function ReviewDeleteNavigationPaneObjectShortcutsEverywhere(ByVal objectID As Long, ByVal objectName As String, ByRef warningDetail As String) As Long
    'D-owned exact stale cleanup for an obsolete ADC-managed object identity; removes only matching shortcut memberships, never unrelated objects.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim sqlText As String
    Dim matchCriteria As String
    Dim cleanedObjectName As String
    warningDetail = vbNullString
    cleanedObjectName = Trim$(objectName)
    matchCriteria = ReviewNavigationPaneObjectShortcutMatchCriteria(objectID, cleanedObjectName)
    If Len(matchCriteria) = 0 Then Exit Function
    Set db = currentDb
    sqlText = "DELETE FROM [" & NAV_PANE_GROUP_OBJECTS_TABLE & "] WHERE (" & matchCriteria & ");"
    db.Execute sqlText, dbFailOnError
    ReviewDeleteNavigationPaneObjectShortcutsEverywhere = db.RecordsAffected
    Exit Function
ErrHandler:
    warningDetail = "Could not remove obsolete Navigation Pane shortcut rows for [" & cleanedObjectName & "]. ObjectID=" & CStr(objectID) & "; Err.Number=" & CStr(Err.Number) & "; Err.Source=" & Err.Source & "; Err.Description=" & Err.Description
    Err.Clear
End Function

Private Function ReviewCreateCustomNavigationPaneGroup(ByVal groupName As String, ByVal sourceName As String, ByRef physicalSnapshot As Object, Optional ByVal trackSetupRollback As Boolean = True) As Long
    'Single physical group-creation path for Setup and runtime callers. Caller supplies a current bulk snapshot; mandatory post-insert physical verification remains.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim categoryID As Long
    Dim groupCount As Long
    Dim createdGroupID As Long
    Dim rollbackTracked As Boolean
    Dim freshSnapshot As Object
    If Len(groupName) = 0 Then Err.Raise vbObjectError + 8813, "ReviewCreateCustomNavigationPaneGroup", "A configured Navigation Pane group name cannot be blank."
    If physicalSnapshot Is Nothing Then Err.Raise vbObjectError + 8814, "ReviewCreateCustomNavigationPaneGroup", "I1 group creation requires the current physical-group snapshot."
    If CBool(physicalSnapshot("__HasName")) = False Then Err.Raise vbObjectError + 8815, "ReviewCreateCustomNavigationPaneGroup", "MSysNavPaneGroups.Name is unavailable."
    categoryID = CLng(physicalSnapshot("__CustomCategoryID"))
    Set db = currentDb
    Set rs = db.OpenRecordset(NAV_PANE_GROUPS_TABLE, dbOpenDynaset)
    rs.AddNew
    ReviewSetRecordsetFieldIfExists rs, "Name", groupName
    ReviewSetRecordsetFieldIfExists rs, "Type", 0
    ReviewSetRecordsetFieldIfExists rs, "Flags", 0
    If CBool(physicalSnapshot("__HasPosition")) Then ReviewSetRecordsetFieldIfExists rs, "Position", CLng(physicalSnapshot("__MaxPosition")) + 1
    If categoryID > 0 Then
        ReviewSetRecordsetFieldIfExists rs, "GroupCategoryID", categoryID
        ReviewSetRecordsetFieldIfExists rs, "CategoryID", categoryID
    End If
    rs.Update
    rs.Close
    Set rs = Nothing
    Set freshSnapshot = ReviewPhysicalNavigationPaneGroupSnapshot()
    createdGroupID = 0
    groupCount = ReviewPhysicalNavigationPaneGroupFromSnapshot(freshSnapshot, groupName, createdGroupID)
    If groupCount <> 1 Or createdGroupID <= 0 Then Err.Raise vbObjectError + 8816, "ReviewCreateCustomNavigationPaneGroup", "The new Navigation Pane group could not be uniquely resolved after insertion." & vbCrLf & "Group: " & groupName & vbCrLf & "Physical count: " & CStr(groupCount)
    Set physicalSnapshot = freshSnapshot
    If trackSetupRollback And SetupRollbackTrackingActive() Then
        TrackSetupNavPaneGroupRollback createdGroupID, groupName
        rollbackTracked = True
    End If
    ReviewCreateCustomNavigationPaneGroup = createdGroupID
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
    If createdGroupID > 0 And rollbackTracked = False Then
        Set db = currentDb
        db.Execute "DELETE FROM [" & NAV_PANE_GROUPS_TABLE & "] WHERE [ID]=" & CStr(createdGroupID) & " AND UCase(Nz([Name],''))=" & SharedDao_SqlQuote(UCase$(groupName)) & " AND [ID] Not In (SELECT [GroupID] FROM [" & NAV_PANE_GROUP_OBJECTS_TABLE & "]);", dbFailOnError
    End If
    On Error GoTo 0
    Err.Raise savedErrNumber, "ReviewCreateCustomNavigationPaneGroup", "Could not create configured Navigation Pane group [" & groupName & "]." & vbCrLf & "Source: " & sourceName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function ReviewCustomNavigationPaneCategoryID() As Long
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = currentDb
    Set rs = db.OpenRecordset("SELECT TOP 1 [ID] FROM [MSysNavPaneGroupCategories] WHERE UCase(Nz([Name],'')) IN ('CUSTOM','CUSTOM GROUPS','CUSTOM CATEGORY') ORDER BY " & OrderByPolicyText("REVIEW_CUSTOM_NAV_CATEGORY") & ";", dbOpenSnapshot)
    If rs.EOF = False Then ReviewCustomNavigationPaneCategoryID = CLng(Nz(rs.fields("ID").value, 0))
    rs.Close
    Exit Function
ErrHandler:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    ReviewCustomNavigationPaneCategoryID = 0
End Function

Private Sub ResetSetupNavigationPanePlan()
    mSetupNavigationPanePlanReady = False
    Set mSetupNavigationPanePlan = Nothing
    mSetupNavigationPanePlanCheckedCount = 0
    mSetupNavigationPanePlanAlreadyPresentCount = 0
    mSetupNavigationPanePlanOptionalMissingCount = 0
    mSetupNavigationPanePlanToken = vbNullString
End Sub

Private Function ReviewNavigationPaneObjectIDCache(ByVal hostDb As DAO.Database) As Object
    ' I2 read-only cache: Access object type + exact name -> current MSysObjects.ID.
    On Error GoTo ErrHandler
    Dim cache As Object
    Dim rs As DAO.Recordset
    Dim msysType As Long
    Dim accessType As Long
    Dim objectName As String
    Dim cacheKey As String
    ' INTENTIONAL KEYED LOOKUP: cache provides keyed membership/lookup over source rows or state already read by ReviewNavigationPaneObjectIDCache; it is a derived access path, not a second authority.
    Set cache = CreateObject("Scripting.Dictionary")
    cache.CompareMode = vbTextCompare
    Set rs = hostDb.OpenRecordset("SELECT [ID],[Name],[Type] FROM [" & MSYS_OBJECTS_TABLE & "] WHERE Nz([ID],0)<>0 AND Len(Trim(Nz([Name],'')))>0 AND [Type] IN (" & SupportedMSysObjectTypesSqlList() & ") ORDER BY [ID];", dbOpenSnapshot)
    Do While rs.EOF = False
        msysType = CLng(rs.fields("Type").value)
        accessType = CLng(AccessObjectTypeFromMSysType(msysType))
        objectName = CStr(rs.fields("Name").value)
        cacheKey = CStr(accessType) & "|" & UCase$(objectName)
        If cache.Exists(cacheKey) = False Then cache.Add cacheKey, CLng(rs.fields("ID").value)
        rs.MoveNext
    Loop
    rs.Close
    Set ReviewNavigationPaneObjectIDCache = cache
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
    Err.Raise savedErrNumber, "ReviewNavigationPaneObjectIDCache", "I2 could not preload the current MSysObjects identities." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function ReviewNavigationPaneObjectMappingCache(ByVal hostDb As DAO.Database) As Object
    ' I2 read-only cache: Access object type + exact name -> exact MSysNavPaneObjectIDs.ID mirror identity.
    On Error GoTo ErrHandler
    Dim cache As Object
    Dim rs As DAO.Recordset
    Dim navType As Long
    Dim accessType As Long
    Dim objectTypeCode As String
    Dim objectName As String
    Dim cacheKey As String
    ' INTENTIONAL KEYED LOOKUP: cache provides keyed membership/lookup over source rows or state already read by ReviewNavigationPaneObjectMappingCache; it is a derived access path, not a second authority.
    Set cache = CreateObject("Scripting.Dictionary")
    cache.CompareMode = vbTextCompare
    If SharedDao_TableFieldExists(currentDb, NAV_PANE_OBJECT_IDS_TABLE, "ID") = False Or SharedDao_TableFieldExists(currentDb, NAV_PANE_OBJECT_IDS_TABLE, "Name") = False Or SharedDao_TableFieldExists(currentDb, NAV_PANE_OBJECT_IDS_TABLE, "Type") = False Then Err.Raise vbObjectError + 8823, "ReviewNavigationPaneObjectMappingCache", "MSysNavPaneObjectIDs must expose ID, Name, and Type for I2 read-only identity proof."
    Set rs = hostDb.OpenRecordset("SELECT [ID],[Name],[Type] FROM [" & NAV_PANE_OBJECT_IDS_TABLE & "] WHERE Nz([ID],0)<>0 AND Len(Trim(Nz([Name],'')))>0 ORDER BY [ID];", dbOpenSnapshot)
    Do While rs.EOF = False
        navType = CLng(Nz(rs.fields("Type").value, 0))
        objectTypeCode = ObjectTypeCodeFromMSysType(navType)
        accessType = -1
        If Len(objectTypeCode) > 0 Then accessType = CLng(AccessObjectTypeFromCode(objectTypeCode))
        If accessType >= 0 Then
            objectName = CStr(rs.fields("Name").value)
            cacheKey = CStr(accessType) & "|" & UCase$(objectName)
            If cache.Exists(cacheKey) Then Err.Raise vbObjectError + 8824, "ReviewNavigationPaneObjectMappingCache", "MSysNavPaneObjectIDs contains duplicate type/name mappings for [" & objectName & "]."
            cache.Add cacheKey, CLng(rs.fields("ID").value)
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set ReviewNavigationPaneObjectMappingCache = cache
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
    Err.Raise savedErrNumber, "ReviewNavigationPaneObjectMappingCache", "I2 could not preload the read-only Navigation Pane object-ID mirror." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function ReviewNavigationPaneMembershipCache(ByVal hostDb As DAO.Database) As Object
    ' I2/I3 read-only cache: GroupID|ObjectID -> exact membership identity used for presence and backup exclusivity proof.
    On Error GoTo ErrHandler
    Dim cache As Object
    Dim rs As DAO.Recordset
    Dim membershipKey As String
    Dim selectFields As String
    Dim groupID As Long
    Dim objectID As Long
    Dim objectName As Variant
    ' INTENTIONAL KEYED LOOKUP: cache provides keyed membership/lookup over source rows or state already read by ReviewNavigationPaneMembershipCache; it is a derived access path, not a second authority.
    Set cache = CreateObject("Scripting.Dictionary")
    cache.CompareMode = vbTextCompare
    selectFields = "[GroupID],[ObjectID]"
    If SharedDao_TableFieldExists(hostDb, NAV_PANE_GROUP_OBJECTS_TABLE, "Name") Then selectFields = selectFields & ",[Name]"
    Set rs = hostDb.OpenRecordset("SELECT " & selectFields & " FROM [" & NAV_PANE_GROUP_OBJECTS_TABLE & "] WHERE Nz([GroupID],0)<>0 AND Nz([ObjectID],0)<>0;", dbOpenSnapshot)
    Do While rs.EOF = False
        groupID = CLng(rs.fields("GroupID").value)
        objectID = CLng(rs.fields("ObjectID").value)
        objectName = ReviewRecordsetFieldValueIfExists(rs, "Name")
        membershipKey = CStr(groupID) & "|" & CStr(objectID)
        If cache.Exists(membershipKey) = False Then cache.Add membershipKey, Array(groupID, objectID, objectName)
        rs.MoveNext
    Loop
    rs.Close
    Set ReviewNavigationPaneMembershipCache = cache
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
    Err.Raise savedErrNumber, "ReviewNavigationPaneMembershipCache", "I2 could not preload current Navigation Pane memberships." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Sub ReviewEnsureNavigationPaneInstrumentationRun()
    Dim activeTraceRunID As String
    activeTraceRunID = GetActiveTraceRunID()
    If StrComp(mNavPaneInstrumentationTraceRunID, activeTraceRunID, vbBinaryCompare) = 0 Then Exit Sub
    mNavPaneInstrumentationTraceRunID = activeTraceRunID
    mNavPanePhysicalGroupSnapshotBuildCount = 0
    mNavPaneDefinitionSnapshotLoadCount = 0
    mNavPaneCapabilityDetectionCount = 0
    mNavPaneCustomCategoryLookupCount = 0
    mNavPaneSameRunProofReuseCount = 0
    mNavPaneMetadataCacheValid = False
    mNavPaneMetadataCacheTraceRunID = vbNullString
    mNavPaneMetadataCacheHostDbName = vbNullString
End Sub

Private Sub ReviewEnsureNavigationPaneMetadataCache()
    'System-field capabilities and custom category identity are immutable inputs for one active setup/HostDb generation.
    Dim activeTraceRunID As String
    Dim hostDbName As String
    activeTraceRunID = GetActiveTraceRunID()
    hostDbName = currentDb.Name
    If mNavPaneMetadataCacheValid Then
        If Len(activeTraceRunID) > 0 And StrComp(activeTraceRunID, mNavPaneMetadataCacheTraceRunID, vbBinaryCompare) = 0 And StrComp(hostDbName, mNavPaneMetadataCacheHostDbName, vbTextCompare) = 0 Then Exit Sub
    End If
    mNavPaneCapabilityDetectionCount = mNavPaneCapabilityDetectionCount + 1
    mNavPaneMetadataHasName = SharedDao_TableFieldExists(currentDb, NAV_PANE_GROUPS_TABLE, "Name")
    mNavPaneMetadataHasGroupCategoryID = SharedDao_TableFieldExists(currentDb, NAV_PANE_GROUPS_TABLE, "GroupCategoryID")
    mNavPaneMetadataHasCategoryID = SharedDao_TableFieldExists(currentDb, NAV_PANE_GROUPS_TABLE, "CategoryID")
    mNavPaneMetadataHasType = SharedDao_TableFieldExists(currentDb, NAV_PANE_GROUPS_TABLE, "Type")
    mNavPaneMetadataHasPosition = SharedDao_TableFieldExists(currentDb, NAV_PANE_GROUPS_TABLE, "Position")
    mNavPaneCustomCategoryLookupCount = mNavPaneCustomCategoryLookupCount + 1
    mNavPaneMetadataCustomCategoryID = ReviewCustomNavigationPaneCategoryID()
    If mNavPaneMetadataHasName = False Then Err.Raise vbObjectError + 8869, "ReviewEnsureNavigationPaneMetadataCache", "MSysNavPaneGroups.Name is unavailable."
    If mNavPaneMetadataCustomCategoryID <= 0 And mNavPaneMetadataHasType = False Then Err.Raise vbObjectError + 8870, "ReviewEnsureNavigationPaneMetadataCache", "Access does not expose enough Navigation Pane metadata to distinguish custom groups safely."
    mNavPaneMetadataCacheTraceRunID = activeTraceRunID
    mNavPaneMetadataCacheHostDbName = hostDbName
    mNavPaneMetadataCacheValid = (Len(activeTraceRunID) > 0)
End Sub

Private Function ReviewPhysicalNavigationPaneGroupSnapshot() As Object
    'One fresh HostDb generation: detect capabilities/category once, then bulk-load every physical custom Navigation Pane group.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim snapshot As Object
    Dim customCategoryID As Long
    Dim hasName As Boolean
    Dim hasGroupCategoryID As Boolean
    Dim hasCategoryID As Boolean
    Dim hasType As Boolean
    Dim hasPosition As Boolean
    Dim rowIsCustom As Boolean
    Dim rawName As String
    Dim snapshotKey As String
    Dim rowData As Variant
    Dim maxPosition As Long
    ReviewEnsureNavigationPaneInstrumentationRun
    ReviewEnsureNavigationPaneMetadataCache
    hasName = mNavPaneMetadataHasName
    hasGroupCategoryID = mNavPaneMetadataHasGroupCategoryID
    hasCategoryID = mNavPaneMetadataHasCategoryID
    hasType = mNavPaneMetadataHasType
    hasPosition = mNavPaneMetadataHasPosition
    customCategoryID = mNavPaneMetadataCustomCategoryID
    ' INTENTIONAL OPERATION SNAPSHOT: snapshot captures one coherent source state for this invocation of ReviewPhysicalNavigationPaneGroupSnapshot; validity ends with the operation rather than becoming a live data
    ' authority.
    Set snapshot = CreateObject("Scripting.Dictionary")
    snapshot.CompareMode = vbTextCompare
    snapshot.Add "__CustomCategoryID", customCategoryID
    snapshot.Add "__HasName", hasName
    snapshot.Add "__HasGroupCategoryID", hasGroupCategoryID
    snapshot.Add "__HasCategoryID", hasCategoryID
    snapshot.Add "__HasType", hasType
    snapshot.Add "__HasPosition", hasPosition
    snapshot.Add "__MaxPosition", 0&
    Set db = currentDb
    Set rs = db.OpenRecordset("SELECT * FROM [" & NAV_PANE_GROUPS_TABLE & "] ORDER BY [ID];", dbOpenSnapshot)
    Do While rs.EOF = False
        rowIsCustom = False
        If customCategoryID > 0 And (hasGroupCategoryID Or hasCategoryID) Then
            If hasGroupCategoryID Then rowIsCustom = (CLng(Nz(rs.fields("GroupCategoryID").value, 0)) = customCategoryID)
            If rowIsCustom = False And hasCategoryID Then rowIsCustom = (CLng(Nz(rs.fields("CategoryID").value, 0)) = customCategoryID)
        ElseIf hasType Then
            rowIsCustom = (CLng(Nz(rs.fields("Type").value, -1)) = 0)
        End If
        If hasPosition Then
            If CLng(Nz(rs.fields("Position").value, 0)) > maxPosition Then maxPosition = CLng(Nz(rs.fields("Position").value, 0))
        End If
        If rowIsCustom Then
            rawName = CStr(Nz(rs.fields("Name").value, vbNullString))
            snapshotKey = "GROUP|" & UCase$(rawName)
            If snapshot.Exists(snapshotKey) Then
                rowData = snapshot(snapshotKey)
                snapshot(snapshotKey) = Array(CLng(rowData(0)) + 1, CLng(rowData(1)))
            Else
                snapshot.Add snapshotKey, Array(1&, CLng(Nz(rs.fields("ID").value, 0)))
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    snapshot("__MaxPosition") = maxPosition
    mNavPanePhysicalGroupSnapshotBuildCount = mNavPanePhysicalGroupSnapshotBuildCount + 1
    Set ReviewPhysicalNavigationPaneGroupSnapshot = snapshot
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
    Err.Raise savedErrNumber, "ReviewPhysicalNavigationPaneGroupSnapshot", "Could not bulk-load the physical custom Navigation Pane groups." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function ReviewPhysicalNavigationPaneGroupFromSnapshot(ByVal physicalSnapshot As Object, ByVal groupName As String, ByRef resolvedGroupID As Long) As Long
    Dim snapshotKey As String
    Dim rowData As Variant
    resolvedGroupID = 0
    groupName = Trim$(groupName)
    If Len(groupName) = 0 Or physicalSnapshot Is Nothing Then Exit Function
    snapshotKey = "GROUP|" & UCase$(groupName)
    If physicalSnapshot.Exists(snapshotKey) = False Then Exit Function
    rowData = physicalSnapshot(snapshotKey)
    ReviewPhysicalNavigationPaneGroupFromSnapshot = CLng(rowData(0))
    resolvedGroupID = CLng(rowData(1))
End Function

Private Sub ReviewAddNavigationPaneGroupDefinitionCacheEntry(ByVal cache As Object, ByVal cacheKey As String, ByVal groupName As String, ByVal physicalCount As Long, ByVal physicalID As Long)
    Dim rowData As Variant
    Dim definitionCount As Long
    If Len(cacheKey) = 0 Then Exit Sub
    If cache.Exists(cacheKey) Then
        rowData = cache(cacheKey)
        definitionCount = CLng(rowData(0)) + 1
        If StrComp(CStr(rowData(1)), groupName, vbTextCompare) <> 0 Then groupName = CStr(rowData(1)) & " <> " & groupName
        cache(cacheKey) = Array(definitionCount, groupName, CLng(rowData(2)), CLng(rowData(3)))
    Else
        cache.Add cacheKey, Array(1&, groupName, physicalCount, physicalID)
    End If
End Sub

Private Sub ReviewNavigationPaneGroupDefinitionCaches(ByVal physicalSnapshot As Object, ByRef settingCache As Object, ByRef groupKeyCache As Object, ByRef configuredCurrent As Boolean, ByRef configuredDetail As String)
    'One authoritative UI_Groups load builds both indexes and proves I1's configured-group prerequisite from the same row set.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim SettingKey As String
    Dim groupKey As String
    Dim groupName As String
    Dim physicalCount As Long
    Dim physicalID As Long
    Dim checkedCount As Long
    Dim failureCount As Long
    Dim failureDetails As String
    ' INTENTIONAL KEYED LOOKUP: settingCache, groupKeyCache provides keyed membership/lookup over source rows or state already read by ReviewNavigationPaneGroupDefinitionCaches; it is a derived access path, not a
    ' second authority.
    Set settingCache = CreateObject("Scripting.Dictionary")
    settingCache.CompareMode = vbTextCompare
    Set groupKeyCache = CreateObject("Scripting.Dictionary")
    groupKeyCache.CompareMode = vbTextCompare
    Set db = ADCDb()
    Set rs = db.OpenRecordset("SELECT [" & FIELD_SETTING_KEY & "],[" & FIELD_NAVIGATION_PANE_GROUP_KEY & "],[" & FIELD_NAVIGATION_PANE_GROUP_NAME & "],[" & FIELD_CREATE_IF_MISSING & "] FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups") & "] WHERE Nz(" & SQL_FIELD_LOOKUP_ACTIVE & ",True)<>False ORDER BY " & OrderByPolicyText("NAVIGATION_GROUP_SETTING_MATCH") & ";", dbOpenSnapshot)
    Do While rs.EOF = False
        SettingKey = Trim$(CStr(Nz(rs.fields(FIELD_SETTING_KEY).value, vbNullString)))
        groupKey = Trim$(CStr(Nz(rs.fields(FIELD_NAVIGATION_PANE_GROUP_KEY).value, vbNullString)))
        groupName = Trim$(CStr(Nz(rs.fields(FIELD_NAVIGATION_PANE_GROUP_NAME).value, vbNullString)))
        physicalID = 0
        physicalCount = ReviewPhysicalNavigationPaneGroupFromSnapshot(physicalSnapshot, groupName, physicalID)
        ReviewAddNavigationPaneGroupDefinitionCacheEntry settingCache, SettingKey, groupName, physicalCount, physicalID
        ReviewAddNavigationPaneGroupDefinitionCacheEntry groupKeyCache, groupKey, groupName, physicalCount, physicalID
        If CBool(Nz(rs.fields(FIELD_CREATE_IF_MISSING).value, False)) And Len(groupName) > 0 Then
            checkedCount = checkedCount + 1
            If physicalCount <> 1 Or physicalID <= 0 Then
                failureCount = failureCount + 1
                failureDetails = SharedText_AppendNonBlankLine(failureDetails, "Group [" & groupName & "] physical match count=" & CStr(physicalCount) & ".")
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    mNavPaneDefinitionSnapshotLoadCount = mNavPaneDefinitionSnapshotLoadCount + 1
    configuredCurrent = (failureCount = 0)
    configuredDetail = "I1 configured-group proof: checked=" & CStr(checkedCount) & ", failures=" & CStr(failureCount) & "; UI_Groups authoritative loads=" & CStr(mNavPaneDefinitionSnapshotLoadCount) & "; MSysNavPaneGroups bulk loads=" & CStr(mNavPanePhysicalGroupSnapshotBuildCount) & "; capability detections=" & CStr(mNavPaneCapabilityDetectionCount) & "; custom category lookups=" & CStr(mNavPaneCustomCategoryLookupCount) & "."
    If Len(failureDetails) > 0 Then configuredDetail = configuredDetail & vbCrLf & failureDetails
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
    Err.Raise savedErrNumber, "ReviewNavigationPaneGroupDefinitionCaches", "Could not preload authoritative Navigation Pane group definitions." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub InvalidateReviewCompleteNavigationPaneProof()
    Set mCompleteNavigationPaneProofRows = Nothing
    mCompleteNavigationPaneProofToken = vbNullString
    mCompleteNavigationPaneProofTraceRunID = vbNullString
    mCompleteNavigationPaneProofHostDbName = vbNullString
End Sub

Private Sub PublishReviewCompleteNavigationPaneProof(ByVal planRows As Collection, ByVal tokenText As String)
    If planRows Is Nothing Or Len(tokenText) = 0 Then
        InvalidateReviewCompleteNavigationPaneProof
        Exit Sub
    End If
    Set mCompleteNavigationPaneProofRows = planRows
    mCompleteNavigationPaneProofToken = tokenText
    mCompleteNavigationPaneProofTraceRunID = GetActiveTraceRunID()
    mCompleteNavigationPaneProofHostDbName = currentDb.Name
End Sub

Private Function ReviewSameRunNavigationPaneProofCurrent(ByVal proofRows As Collection, ByVal proofToken As String, ByVal proofTraceRunID As String, ByVal proofHostDbName As String, ByVal proofLabel As String, ByRef sharedObjectCache As Object, ByRef sharedNavObjectCache As Object, ByRef sharedMembershipCache As Object, ByRef sharedPhysicalGroupSnapshot As Object, ByRef detailText As String) As Boolean
    'Compact same-run physical drift gate over the exact complete-domain rows already post-verified by I3. No catalog/resolver work is repeated here.
    On Error GoTo NotCurrent
    Dim activeTraceRunID As String
    Dim hostDb As DAO.Database
    Dim objectCache As Object
    Dim navObjectCache As Object
    Dim membershipCache As Object
    Dim physicalGroupSnapshot As Object
    Dim checkedGroups As Object
    Dim planItem As Variant
    Dim objectName As String
    Dim objectType As AcObjectType
    Dim groupName As String
    Dim expectedGroupID As Long
    Dim expectedObjectID As Long
    Dim actualGroupID As Long
    Dim groupCount As Long
    Dim objectKey As String
    Dim membershipKey As String
    Dim groupKey As String
    Dim currentToken As String
    detailText = vbNullString
    If proofRows Is Nothing Or Len(proofToken) = 0 Then detailText = proofLabel & " same-run exact proof is unavailable.": Exit Function
    activeTraceRunID = GetActiveTraceRunID()
    If Len(activeTraceRunID) = 0 Or StrComp(activeTraceRunID, proofTraceRunID, vbBinaryCompare) <> 0 Then detailText = proofLabel & " exact proof belongs to a different or inactive setup trace run.": Exit Function
    Set hostDb = currentDb
    If StrComp(hostDb.Name, proofHostDbName, vbTextCompare) <> 0 Then detailText = proofLabel & " exact proof belongs to a different HostDb.": Exit Function
    currentToken = ReviewNavigationPanePlanToken(proofRows)
    If StrComp(currentToken, proofToken, vbBinaryCompare) <> 0 Then detailText = proofLabel & " retained exact plan rows no longer match their proof token.": Exit Function
    If sharedObjectCache Is Nothing Then Set sharedObjectCache = ReviewNavigationPaneObjectIDCache(hostDb)
    If sharedNavObjectCache Is Nothing Then Set sharedNavObjectCache = ReviewNavigationPaneObjectMappingCache(hostDb)
    If sharedMembershipCache Is Nothing Then Set sharedMembershipCache = ReviewNavigationPaneMembershipCache(hostDb)
    If sharedPhysicalGroupSnapshot Is Nothing Then Set sharedPhysicalGroupSnapshot = ReviewPhysicalNavigationPaneGroupSnapshot()
    Set objectCache = sharedObjectCache
    Set navObjectCache = sharedNavObjectCache
    Set membershipCache = sharedMembershipCache
    Set physicalGroupSnapshot = sharedPhysicalGroupSnapshot
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: checkedGroups holds planned/recovery work or independently observed validation evidence for ReviewSameRunNavigationPaneProofCurrent; keep it distinct from the
    ' source authority.
    Set checkedGroups = CreateObject("Scripting.Dictionary")
    checkedGroups.CompareMode = vbTextCompare
    For Each planItem In proofRows
        objectName = CStr(planItem(0))
        objectType = CLng(planItem(1))
        groupName = CStr(planItem(3))
        expectedGroupID = CLng(planItem(4))
        expectedObjectID = CLng(planItem(5))
        groupKey = UCase$(Trim$(groupName))
        If checkedGroups.Exists(groupKey) = False Then
            actualGroupID = 0
            groupCount = ReviewPhysicalNavigationPaneGroupFromSnapshot(physicalGroupSnapshot, groupName, actualGroupID)
            If groupCount <> 1 Or actualGroupID <> expectedGroupID Then detailText = proofLabel & " compact drift gate detected a changed physical group: [" & groupName & "].": Exit Function
            checkedGroups.Add groupKey, True
        End If
        objectKey = CStr(CLng(objectType)) & "|" & UCase$(objectName)
        If objectCache.Exists(objectKey) = False Then detailText = proofLabel & " compact drift gate detected a missing physical object: [" & objectName & "].": Exit Function
        If CLng(objectCache(objectKey)) <> expectedObjectID Then detailText = proofLabel & " compact drift gate detected a changed physical ObjectID: [" & objectName & "].": Exit Function
        If navObjectCache.Exists(objectKey) = False Then detailText = proofLabel & " compact drift gate detected a missing Navigation Pane mirror identity: [" & objectName & "].": Exit Function
        If CLng(navObjectCache(objectKey)) <> expectedObjectID Then detailText = proofLabel & " compact drift gate detected a changed Navigation Pane mirror identity: [" & objectName & "].": Exit Function
        membershipKey = CStr(expectedGroupID) & "|" & CStr(expectedObjectID)
        If membershipCache.Exists(membershipKey) = False Then detailText = proofLabel & " compact drift gate detected a missing Navigation Pane membership: [" & objectName & "] -> [" & groupName & "].": Exit Function
    Next planItem
    mNavPaneSameRunProofReuseCount = mNavPaneSameRunProofReuseCount + 1
    detailText = proofLabel & " same-run exact proof reused after compact physical drift gate. Rows=" & CStr(proofRows.count) & "; unique groups=" & CStr(checkedGroups.count) & "; same-run reuses=" & CStr(mNavPaneSameRunProofReuseCount) & "."
    ReviewSameRunNavigationPaneProofCurrent = True
    Exit Function
NotCurrent:
    If IsCancellationError(Err.Number) Then Err.Raise Err.Number, Err.Source, Err.Description
    If Len(detailText) = 0 Then detailText = proofLabel & " compact drift gate could not prove current state. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    Err.Clear
End Function

Private Sub ReviewAppendDynamicBackupNavigationPanePlanRows(ByVal hostDb As DAO.Database, ByVal objectIdCache As Object, ByVal navObjectIdCache As Object, ByVal membershipCache As Object, ByVal physicalGroupSnapshot As Object, ByVal requireCurrentPlacement As Boolean, ByRef planRows As Collection, ByRef checkedCount As Long, ByRef alreadyPresentCount As Long, ByRef identityMaterializationCount As Long, ByRef identityRemapCount As Long, ByRef failureCount As Long, ByRef failureText As String)
    ' Shared I2/durable-I3 dynamic placement authority: discovers canonical physical backup tables from the existing MSysObjects cache and appends them without MAC rows.
    Dim cacheKeyValue As Variant
    Dim membershipCacheKey As Variant
    Dim membershipData As Variant
    Dim tableKeyPrefix As String
    Dim candidateName As String
    Dim objectName As String
    Dim objectKey As String
    Dim membershipKey As String
    Dim groupName As String
    Dim objectID As Long
    Dim groupID As Long
    Dim groupCount As Long
    Dim navIdentityMappedID As Long
    Dim membershipGroupID As Long
    Dim membershipObjectID As Long
    Dim membershipObjectName As String
    Dim alreadyPresent As Boolean
    Dim navIdentityPresent As Boolean
    If (hostDb Is Nothing) Or (objectIdCache Is Nothing) Or (navObjectIdCache Is Nothing) Or (membershipCache Is Nothing) Or (physicalGroupSnapshot Is Nothing) Or (planRows Is Nothing) Then Err.Raise vbObjectError + 8897, "ReviewAppendDynamicBackupNavigationPanePlanRows", "Dynamic backup Navigation Pane planning requires the existing I2 physical caches and destination plan."
    tableKeyPrefix = CStr(CLng(acTable)) & "|"
    For Each cacheKeyValue In objectIdCache.keys
        objectKey = CStr(cacheKeyValue)
        If Left$(objectKey, Len(tableKeyPrefix)) = tableKeyPrefix Then
            candidateName = Mid$(objectKey, Len(tableKeyPrefix) + 1)
            If IsBackupAccessObjectName(candidateName) Then
                PumpSetupResponsiveness
                objectName = hostDb.TableDefs(candidateName).Name
                objectID = CLng(objectIdCache(objectKey))
                If groupID <= 0 Then
                    groupName = BuildNavigationPaneGroupName(DEFAULT_BACKUP_NAVIGATION_PANE_GROUP_LOGICAL_NAME, False)
                    groupCount = ReviewPhysicalNavigationPaneGroupFromSnapshot(physicalGroupSnapshot, groupName, groupID)
                    If Len(groupName) = 0 Or groupCount <> 1 Or groupID <= 0 Then
                        failureCount = failureCount + 1
                        failureText = SharedText_AppendNonBlankLine(failureText, "Dynamic ADC backup tables require one physical Backup Navigation Pane group. Group=[" & groupName & "]; count=" & CStr(groupCount) & ".")
                        Exit Sub
                    End If
                End If
                objectKey = CStr(CLng(acTable)) & "|" & UCase$(objectName)
                navIdentityMappedID = 0
                navIdentityPresent = False
                If navObjectIdCache.Exists(objectKey) Then
                    navIdentityMappedID = CLng(navObjectIdCache(objectKey))
                    navIdentityPresent = (navIdentityMappedID = objectID)
                End If
                membershipKey = CStr(groupID) & "|" & CStr(objectID)
                alreadyPresent = membershipCache.Exists(membershipKey)
                If requireCurrentPlacement Then
                    If navIdentityPresent = False Then
                        failureCount = failureCount + 1
                        failureText = SharedText_AppendNonBlankLine(failureText, "ADC backup Table [" & objectName & "] does not have matching physical and Navigation Pane object identity.")
                    End If
                    If alreadyPresent = False Then
                        failureCount = failureCount + 1
                        failureText = SharedText_AppendNonBlankLine(failureText, "Required I3 backup membership is missing: Table [" & objectName & "] -> [" & groupName & "].")
                    End If
                    For Each membershipCacheKey In membershipCache.keys
                        membershipData = membershipCache(membershipCacheKey)
                        membershipGroupID = CLng(membershipData(0))
                        membershipObjectID = CLng(membershipData(1))
                        membershipObjectName = Trim$(CStr(Nz(membershipData(2), vbNullString)))
                        If membershipObjectID = objectID Or (Len(membershipObjectName) > 0 And StrComp(membershipObjectName, objectName, vbTextCompare) = 0) Then
                            If membershipGroupID <> groupID Or membershipObjectID <> objectID Then
                                failureCount = failureCount + 1
                                failureText = SharedText_AppendNonBlankLine(failureText, "ADC backup Table [" & objectName & "] retains a Navigation Pane shortcut outside the canonical Backup group. GroupID=" & CStr(membershipGroupID) & ".")
                            End If
                        End If
                    Next membershipCacheKey
                Else
                    If navIdentityPresent = False Then
                        If navIdentityMappedID = 0 Then
                            identityMaterializationCount = identityMaterializationCount + 1
                        Else
                            identityRemapCount = identityRemapCount + 1
                        End If
                    End If
                End If
                planRows.Add Array(objectName, CLng(acTable), ADCAT_OBJECT_TYPE_TABLE, groupName, groupID, objectID, alreadyPresent, vbNullString, navIdentityPresent, navIdentityMappedID)
                checkedCount = checkedCount + 1
                If alreadyPresent Then alreadyPresentCount = alreadyPresentCount + 1
            End If
        End If
    Next cacheKeyValue
End Sub

Private Function ReviewNavigationPanePlanToken(ByVal planRows As Collection) As String
    Dim item As Variant
    Dim payloadText As String
    payloadText = "ADCPartGCompleteNavigationPanePlanV1|"
    For Each item In planRows
        payloadText = payloadText & CStr(item(0)) & "|" & CStr(item(1)) & "|" & CStr(item(2)) & "|" & CStr(item(3)) & "|" & CStr(item(4)) & "|" & CStr(item(5)) & vbLf
    Next item
    ReviewNavigationPanePlanToken = FingerprintFromPayload(payloadText)
End Function

' Purpose: Evaluates the readiness conditions for setup navigation pane plan and reports whether all required conditions are currently satisfied.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   planToken (Output; String): Receives the resolved plan token produced from this procedure's validated inputs and current ADC state.
'   planRowCount (Output; Long): Receives the plan row count produced by this operation.
' Returns: Boolean: True when setup navigation pane plan is ready for the next dependent operation; False when a prerequisite is not yet satisfied.
Public Function SetupNavigationPanePlanReady(ByRef planToken As String, ByRef planRowCount As Long) As Boolean
    'Read-only same-session I2 plan proof consumed by I3 and Setup orchestration.
    planToken = mSetupNavigationPanePlanToken
    planRowCount = mSetupNavigationPanePlanCheckedCount
    SetupNavigationPanePlanReady = (mSetupNavigationPanePlanReady And Not mSetupNavigationPanePlan Is Nothing And Len(planToken) > 0)
End Function


' Purpose: Performs rollback setup nav pane shortcut as the module's public orchestration step using the established ADC state and validation contract.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   groupID (Input; Long): Identifier selecting the logical Ribbon, form, catalog, or definition group affected by the operation.
'   objectID (Input; Long): Access/Navigation Pane object ID captured for the target object so placement and rollback can restore the exact prior identity/state.
'   deletedCount (Output; Long): Receives the deleted count produced by this operation.
'   detailText (Output; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
' Returns: None (Sub). Results are returned through the documented ByRef parameters and/or the routine's database, UI, or module-state side effects.
Public Sub RollbackSetupNavPaneShortcut(ByVal groupID As Long, ByVal objectID As Long, ByRef deletedCount As Long, ByRef detailText As String)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim existingCount As Long
    If groupID <= 0 Or objectID = 0 Then Err.Raise vbObjectError + 6443, "RollbackSetupNavPaneShortcut", "A Navigation Pane shortcut rollback entry contained an invalid GroupID or zero ObjectID."
    Set db = currentDb
    Set rs = db.OpenRecordset("SELECT Count(*) " & "AS " & SQL_FIELD_CNT & " FROM [" & NAV_PANE_GROUP_OBJECTS_TABLE & "] WHERE [GroupID]=" & CStr(groupID) & " AND [ObjectID]=" & CStr(objectID) & ";", dbOpenSnapshot)
    existingCount = CLng(rs.fields(FIELD_CNT).value)
    rs.Close
    Set rs = Nothing
    If existingCount > 0 Then
        db.Execute "DELETE FROM [" & NAV_PANE_GROUP_OBJECTS_TABLE & "] WHERE [GroupID]=" & CStr(groupID) & " AND [ObjectID]=" & CStr(objectID) & ";", dbFailOnError
        Set rs = db.OpenRecordset("SELECT Count(*) " & "AS " & SQL_FIELD_CNT & " FROM [" & NAV_PANE_GROUP_OBJECTS_TABLE & "] WHERE [GroupID]=" & CStr(groupID) & " AND [ObjectID]=" & CStr(objectID) & ";", dbOpenSnapshot)
        If CLng(rs.fields(FIELD_CNT).value) <> 0 Then Err.Raise vbObjectError + 6444, "RollbackSetupNavPaneShortcut", "The Navigation Pane shortcut remained present after rollback deletion."
        rs.Close
        Set rs = Nothing
        deletedCount = deletedCount + existingCount
        detailText = SharedText_AppendNonBlankLine(detailText, "Navigation Pane shortcut rollback deleted GroupID=" & CStr(groupID) & ", ObjectID=" & CStr(objectID) & ".")
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
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
    Err.Raise savedErrNumber, "RollbackSetupNavPaneShortcut", "Could not roll back a tracked Navigation Pane shortcut." & vbCrLf & "GroupID: " & CStr(groupID) & vbCrLf & "ObjectID: " & CStr(objectID) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub


Private Function SetupRollbackOptionalNavigationPaneFieldMatches(ByVal rs As DAO.Recordset, ByVal fieldName As String, ByVal expectedValue As Variant) As Boolean
    On Error GoTo MissingField
    Dim actualValue As Variant
    If IsEmpty(expectedValue) Then
        SetupRollbackOptionalNavigationPaneFieldMatches = True
        Exit Function
    End If
    actualValue = rs.fields(fieldName).value
    If IsNull(expectedValue) Then
        SetupRollbackOptionalNavigationPaneFieldMatches = IsNull(actualValue)
    ElseIf IsNull(actualValue) = False Then
        SetupRollbackOptionalNavigationPaneFieldMatches = (StrComp(CStr(actualValue), CStr(expectedValue), vbBinaryCompare) = 0)
    End If
    Exit Function
MissingField:
    Err.Clear
End Function


Private Sub SetupRollbackSetOptionalNavigationPaneField(ByVal rs As DAO.Recordset, ByVal fieldName As String, ByVal expectedValue As Variant)
    If IsEmpty(expectedValue) Then Exit Sub
    rs.fields(fieldName).value = expectedValue
End Sub


' Purpose: Performs rollback setup nav pane shortcut restore as the module's public orchestration step using the established ADC state and validation contract.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   groupID (Input; Long): Identifier selecting the logical Ribbon, form, catalog, or definition group affected by the operation.
'   objectID (Input; Long): Access/Navigation Pane object ID captured for the target object so placement and rollback can restore the exact prior identity/state.
'   priorName (Input; Variant): Pre-change object/name value captured for rollback; may be Empty/Null when no prior value existed.
'   priorType (Input; Variant): Pre-change object/property type captured for rollback; Variant preserves Empty/Null when no prior typed value existed.
'   priorFlags (Input; Variant): Bit flags preserving or selecting the prior attributes used by the underlying Access/DAO operation.
'   priorPosition (Input; Variant): Numeric position locating the prior item within its source, UI layout, or Navigation Pane ordering.
'   restoredCount (Output; Long): Receives the restored count produced by this operation.
'   detailText (Output; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
' Returns: None (Sub). Results are returned through the documented ByRef parameters and/or the routine's database, UI, or module-state side effects.
Public Sub RollbackSetupNavPaneShortcutRestore(ByVal groupID As Long, ByVal objectID As Long, ByVal priorName As Variant, ByVal priorType As Variant, ByVal priorFlags As Variant, ByVal priorPosition As Variant, ByRef restoredCount As Long, ByRef detailText As String)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim rowCount As Long
    Dim exactPriorRow As Boolean
    If groupID <= 0 Or objectID = 0 Then Err.Raise vbObjectError + 6459, "RollbackSetupNavPaneShortcutRestore", "A Navigation Pane shortcut restore entry contained an invalid GroupID or ObjectID."
    Set db = currentDb
    Set rs = db.OpenRecordset("SELECT * FROM [" & NAV_PANE_GROUP_OBJECTS_TABLE & "] WHERE [GroupID]=" & CStr(groupID) & " AND [ObjectID]=" & CStr(objectID) & ";", dbOpenSnapshot)
    Do While rs.EOF = False
        rowCount = rowCount + 1
        If rowCount = 1 Then exactPriorRow = SetupRollbackOptionalNavigationPaneFieldMatches(rs, "Name", priorName) And SetupRollbackOptionalNavigationPaneFieldMatches(rs, "Type", priorType) And SetupRollbackOptionalNavigationPaneFieldMatches(rs, "Flags", priorFlags) And SetupRollbackOptionalNavigationPaneFieldMatches(rs, "Position", priorPosition)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    If rowCount > 1 Then Err.Raise vbObjectError + 6460, "RollbackSetupNavPaneShortcutRestore", "The prior Navigation Pane shortcut is ambiguous during rollback."
    If rowCount = 1 Then
        If exactPriorRow = False Then Err.Raise vbObjectError + 6461, "RollbackSetupNavPaneShortcutRestore", "Rollback refused to overwrite a current Navigation Pane shortcut whose metadata no longer matches the exact prior row."
        Exit Sub
    End If
    Set rs = db.OpenRecordset(NAV_PANE_GROUP_OBJECTS_TABLE, dbOpenDynaset)
    rs.AddNew
    rs.fields("GroupID").value = groupID
    rs.fields("ObjectID").value = objectID
    SetupRollbackSetOptionalNavigationPaneField rs, "Name", priorName
    SetupRollbackSetOptionalNavigationPaneField rs, "Type", priorType
    SetupRollbackSetOptionalNavigationPaneField rs, "Flags", priorFlags
    SetupRollbackSetOptionalNavigationPaneField rs, "Position", priorPosition
    rs.Update
    rs.Close
    Set rs = Nothing
    Set rs = db.OpenRecordset("SELECT * FROM [" & NAV_PANE_GROUP_OBJECTS_TABLE & "] WHERE [GroupID]=" & CStr(groupID) & " AND [ObjectID]=" & CStr(objectID) & ";", dbOpenSnapshot)
    If rs.EOF Then Err.Raise vbObjectError + 6462, "RollbackSetupNavPaneShortcutRestore", "The prior Navigation Pane shortcut was not present after rollback restoration."
    exactPriorRow = SetupRollbackOptionalNavigationPaneFieldMatches(rs, "Name", priorName) And SetupRollbackOptionalNavigationPaneFieldMatches(rs, "Type", priorType) And SetupRollbackOptionalNavigationPaneFieldMatches(rs, "Flags", priorFlags) And SetupRollbackOptionalNavigationPaneFieldMatches(rs, "Position", priorPosition)
    rs.MoveNext
    If rs.EOF = False Or exactPriorRow = False Then Err.Raise vbObjectError + 6463, "RollbackSetupNavPaneShortcutRestore", "The prior Navigation Pane shortcut was not exact after rollback restoration."
    rs.Close
    Set rs = Nothing
    restoredCount = restoredCount + 1
    detailText = SharedText_AppendNonBlankLine(detailText, "Navigation Pane shortcut rollback restored GroupID=" & CStr(groupID) & ", ObjectID=" & CStr(objectID) & ".")
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
    Err.Raise savedErrNumber, "RollbackSetupNavPaneShortcutRestore", "Could not restore a preexisting Navigation Pane shortcut removed for stale identity repair." & vbCrLf & "GroupID: " & CStr(groupID) & vbCrLf & "ObjectID: " & CStr(objectID) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub


' Purpose: Performs rollback setup nav pane object ID restore as the module's public orchestration step using the established ADC state and validation contract.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   objectID (Input; Long): Access/Navigation Pane object ID captured for the target object so placement and rollback can restore the exact prior identity/state.
'   repairedObjectName (Input; String): Physical object name after repair/rename; rollback bookkeeping uses it to locate the repaired object.
'   repairedObjectType (Input; Long): Access object-type code for the repaired/current object, used by rollback to locate/delete/restore the correct object class.
'   priorObjectName (Input; String): Physical object name that existed before repair/rename and must be restored if rollback occurs.
'   priorObjectType (Input; Long): Access object-type code for the pre-change object, used by rollback when restoring its prior identity/state.
'   restoredCount (Output; Long): Receives the restored count produced by this operation.
'   detailText (Output; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
' Returns: None (Sub). Results are returned through the documented ByRef parameters and/or the routine's database, UI, or module-state side effects.
Public Sub RollbackSetupNavPaneObjectIDRestore(ByVal objectID As Long, ByVal repairedObjectName As String, ByVal repairedObjectType As Long, ByVal priorObjectName As String, ByVal priorObjectType As Long, ByRef restoredCount As Long, ByRef detailText As String)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim priorExactCount As Long
    Dim repairedExactCount As Long
    repairedObjectName = Trim$(repairedObjectName)
    If objectID = 0 Or Len(repairedObjectName) = 0 Or repairedObjectType = 0 Then Err.Raise vbObjectError + 6453, "RollbackSetupNavPaneObjectIDRestore", "A Navigation Pane object-identity restore entry contained an invalid repaired identity."
    Set db = currentDb
    Set rs = db.OpenRecordset("SELECT Count(*) " & "AS " & SQL_FIELD_CNT & " FROM [" & NAV_PANE_OBJECT_IDS_TABLE & "] WHERE [ID]=" & CStr(objectID) & " AND UCase(Nz([Name],''))=" & SharedDao_SqlQuote(UCase$(priorObjectName)) & " AND Nz([Type],0)=" & CStr(priorObjectType) & ";", dbOpenSnapshot)
    priorExactCount = CLng(rs.fields(FIELD_CNT).value)
    rs.Close
    Set rs = Nothing
    If priorExactCount = 1 Then Exit Sub
    If priorExactCount > 1 Then Err.Raise vbObjectError + 6454, "RollbackSetupNavPaneObjectIDRestore", "The prior Navigation Pane object identity is ambiguous during rollback."
    Set rs = db.OpenRecordset("SELECT Count(*) " & "AS " & SQL_FIELD_CNT & " FROM [" & NAV_PANE_OBJECT_IDS_TABLE & "] WHERE [ID]=" & CStr(objectID) & " AND UCase(Nz([Name],''))=" & SharedDao_SqlQuote(UCase$(repairedObjectName)) & " AND [Type]=" & CStr(repairedObjectType) & ";", dbOpenSnapshot)
    repairedExactCount = CLng(rs.fields(FIELD_CNT).value)
    rs.Close
    Set rs = Nothing
    If repairedExactCount <> 1 Then Err.Raise vbObjectError + 6455, "RollbackSetupNavPaneObjectIDRestore", "Rollback refused to restore a stale Navigation Pane identity because the expected repaired identity is no longer exact."
    Set rs = db.OpenRecordset("SELECT [ID],[Name],[Type] FROM [" & NAV_PANE_OBJECT_IDS_TABLE & "] WHERE [ID]=" & CStr(objectID) & ";", dbOpenDynaset)
    If rs.EOF Then Err.Raise vbObjectError + 6456, "RollbackSetupNavPaneObjectIDRestore", "The expected repaired Navigation Pane object identity disappeared before rollback restore."
    rs.Edit
    rs.fields("Name").value = priorObjectName
    rs.fields("Type").value = priorObjectType
    rs.Update
    rs.Close
    Set rs = Nothing
    Set rs = db.OpenRecordset("SELECT Count(*) " & "AS " & SQL_FIELD_CNT & " FROM [" & NAV_PANE_OBJECT_IDS_TABLE & "] WHERE [ID]=" & CStr(objectID) & " AND UCase(Nz([Name],''))=" & SharedDao_SqlQuote(UCase$(priorObjectName)) & " AND Nz([Type],0)=" & CStr(priorObjectType) & ";", dbOpenSnapshot)
    If CLng(rs.fields(FIELD_CNT).value) <> 1 Then Err.Raise vbObjectError + 6457, "RollbackSetupNavPaneObjectIDRestore", "The prior Navigation Pane object identity was not exact after rollback restore."
    rs.Close
    Set rs = Nothing
    restoredCount = restoredCount + 1
    detailText = SharedText_AppendNonBlankLine(detailText, "Navigation Pane object-identity rollback restored ObjectID=" & CStr(objectID) & " from [" & repairedObjectName & "] to prior identity [" & priorObjectName & "].")
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
    Err.Raise savedErrNumber, "RollbackSetupNavPaneObjectIDRestore", "Could not restore a replaced Navigation Pane object identity." & vbCrLf & "ObjectID: " & CStr(objectID) & vbCrLf & "Repaired identity: [" & repairedObjectName & "], Type=" & CStr(repairedObjectType) & vbCrLf & "Prior identity: [" & priorObjectName & "], Type=" & CStr(priorObjectType) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub


' Purpose: Performs rollback setup deleted nav pane object ID restore as the module's public orchestration step using the established ADC state and validation contract.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   objectID (Input; Long): Access/Navigation Pane object ID captured for the target object so placement and rollback can restore the exact prior identity/state.
'   objectName (Input; String): Physical Access object name identifying the database object targeted by the operation.
'   objectType (Input; Long): Object-type value identifying the Access/VBA object class targeted by the operation.
'   restoredCount (Output; Long): Receives the restored count produced by this operation.
'   detailText (Output; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
' Returns: None (Sub). Results are returned through the documented ByRef parameters and/or the routine's database, UI, or module-state side effects.
Public Sub RollbackSetupDeletedNavPaneObjectIDRestore(ByVal objectID As Long, ByVal objectName As String, ByVal objectType As Long, ByRef restoredCount As Long, ByRef detailText As String)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim exactCount As Long
    Dim idConflictCount As Long
    objectName = Trim$(objectName)
    If objectID = 0 Or Len(objectName) = 0 Or objectType = 0 Then Err.Raise vbObjectError + 6471, "RollbackSetupDeletedNavPaneObjectIDRestore", "A deleted Navigation Pane object-identity restore entry contained an invalid identity."
    Set db = currentDb
    Set rs = db.OpenRecordset("SELECT Count(*) " & "AS " & SQL_FIELD_CNT & " FROM [" & NAV_PANE_OBJECT_IDS_TABLE & "] WHERE [ID]=" & CStr(objectID) & " AND UCase(Nz([Name],''))=" & SharedDao_SqlQuote(UCase$(objectName)) & " AND [Type]=" & CStr(objectType) & ";", dbOpenSnapshot)
    exactCount = CLng(rs.fields(FIELD_CNT).value)
    rs.Close
    Set rs = Nothing
    If exactCount = 1 Then Exit Sub
    If exactCount > 1 Then Err.Raise vbObjectError + 6472, "RollbackSetupDeletedNavPaneObjectIDRestore", "The deleted Navigation Pane object identity is ambiguous during rollback."
    Set rs = db.OpenRecordset("SELECT Count(*) " & "AS " & SQL_FIELD_CNT & " FROM [" & NAV_PANE_OBJECT_IDS_TABLE & "] WHERE [ID]=" & CStr(objectID) & ";", dbOpenSnapshot)
    idConflictCount = CLng(rs.fields(FIELD_CNT).value)
    rs.Close
    Set rs = Nothing
    If idConflictCount <> 0 Then Err.Raise vbObjectError + 6473, "RollbackSetupDeletedNavPaneObjectIDRestore", "Rollback refused to restore a deleted Navigation Pane object identity because its prior ObjectID is now occupied."
    Set rs = db.OpenRecordset(NAV_PANE_OBJECT_IDS_TABLE, dbOpenDynaset)
    rs.AddNew
    rs.fields("ID").value = objectID
    rs.fields("Name").value = objectName
    rs.fields("Type").value = objectType
    rs.Update
    rs.Close
    Set rs = Nothing
    Set rs = db.OpenRecordset("SELECT Count(*) " & "AS " & SQL_FIELD_CNT & " FROM [" & NAV_PANE_OBJECT_IDS_TABLE & "] WHERE [ID]=" & CStr(objectID) & " AND UCase(Nz([Name],''))=" & SharedDao_SqlQuote(UCase$(objectName)) & " AND [Type]=" & CStr(objectType) & ";", dbOpenSnapshot)
    If CLng(rs.fields(FIELD_CNT).value) <> 1 Then Err.Raise vbObjectError + 6475, "RollbackSetupDeletedNavPaneObjectIDRestore", "The deleted Navigation Pane object identity was not exact after rollback restoration."
    rs.Close
    Set rs = Nothing
    restoredCount = restoredCount + 1
    detailText = SharedText_AppendNonBlankLine(detailText, "Navigation Pane object-identity rollback reinserted the exact prior stale identity ObjectID=" & CStr(objectID) & ", ObjectName=[" & objectName & "], ObjectType=" & CStr(objectType) & ". A same-name/type identity at another ObjectID is permitted because that duplicate logical identity was part of the captured pre-Setup state.")
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
    Err.Raise savedErrNumber, "RollbackSetupDeletedNavPaneObjectIDRestore", "Could not restore a preexisting Navigation Pane object identity deleted during stale-ID remapping." & vbCrLf & "ObjectID: " & CStr(objectID) & vbCrLf & "ObjectName: [" & objectName & "]" & vbCrLf & "ObjectType: " & CStr(objectType) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub


' Purpose: Performs rollback setup nav pane object ID as the module's public orchestration step using the established ADC state and validation contract.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   objectID (Input; Long): Access/Navigation Pane object ID captured for the target object so placement and rollback can restore the exact prior identity/state.
'   objectName (Input; String): Physical Access object name identifying the database object targeted by the operation.
'   objectType (Input; Long): Object-type value identifying the Access/VBA object class targeted by the operation.
'   deletedCount (Output; Long): Receives the deleted count produced by this operation.
'   retainedSharedIdentityCount (Output; Long): Receives the retained shared identity count produced by this operation.
'   detailText (Output; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
' Returns: None (Sub). Results are returned through the documented ByRef parameters and/or the routine's database, UI, or module-state side effects.
Public Sub RollbackSetupNavPaneObjectID(ByVal objectID As Long, ByVal objectName As String, ByVal objectType As Long, ByRef deletedCount As Long, ByRef retainedSharedIdentityCount As Long, ByRef detailText As String)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim exactCount As Long
    Dim membershipCount As Long
    objectName = Trim$(objectName)
    If objectID = 0 Or Len(objectName) = 0 Or objectType = 0 Then Err.Raise vbObjectError + 6448, "RollbackSetupNavPaneObjectID", "A Navigation Pane object-identity rollback entry contained an invalid identity."
    Set db = currentDb
    Set rs = db.OpenRecordset("SELECT Count(*) " & "AS " & SQL_FIELD_CNT & " FROM [" & NAV_PANE_OBJECT_IDS_TABLE & "] WHERE [ID]=" & CStr(objectID) & " AND UCase(Nz([Name],''))=" & SharedDao_SqlQuote(UCase$(objectName)) & " AND [Type]=" & CStr(objectType) & ";", dbOpenSnapshot)
    exactCount = CLng(rs.fields(FIELD_CNT).value)
    rs.Close
    Set rs = Nothing
    If exactCount = 0 Then Exit Sub
    If exactCount <> 1 Then Err.Raise vbObjectError + 6450, "RollbackSetupNavPaneObjectID", "The tracked Navigation Pane object identity is ambiguous during rollback."
    Set rs = db.OpenRecordset("SELECT Count(*) " & "AS " & SQL_FIELD_CNT & " FROM [" & NAV_PANE_GROUP_OBJECTS_TABLE & "] WHERE [ObjectID]=" & CStr(objectID) & ";", dbOpenSnapshot)
    membershipCount = CLng(rs.fields(FIELD_CNT).value)
    rs.Close
    Set rs = Nothing
    If membershipCount <> 0 Then
        retainedSharedIdentityCount = retainedSharedIdentityCount + 1
        detailText = SharedText_AppendNonBlankLine(detailText, "Navigation Pane object-identity rollback retained shared ObjectID=" & CStr(objectID) & ", ObjectName=[" & objectName & "], ObjectType=" & CStr(objectType) & " because " & CStr(membershipCount) & " preexisting or Access-managed membership(s) still reference the exact identity after all tracked setup-created memberships were removed.")
        Exit Sub
    End If
    db.Execute "DELETE FROM [" & NAV_PANE_OBJECT_IDS_TABLE & "] WHERE [ID]=" & CStr(objectID) & " AND UCase(Nz([Name],''))=" & SharedDao_SqlQuote(UCase$(objectName)) & " AND [Type]=" & CStr(objectType) & ";", dbFailOnError
    Set rs = db.OpenRecordset("SELECT Count(*) " & "AS " & SQL_FIELD_CNT & " FROM [" & NAV_PANE_OBJECT_IDS_TABLE & "] WHERE [ID]=" & CStr(objectID) & " AND UCase(Nz([Name],''))=" & SharedDao_SqlQuote(UCase$(objectName)) & " AND [Type]=" & CStr(objectType) & ";", dbOpenSnapshot)
    If CLng(rs.fields(FIELD_CNT).value) <> 0 Then Err.Raise vbObjectError + 6451, "RollbackSetupNavPaneObjectID", "The Navigation Pane object identity remained present after rollback deletion."
    rs.Close
    Set rs = Nothing
    deletedCount = deletedCount + exactCount
    detailText = SharedText_AppendNonBlankLine(detailText, "Navigation Pane object-identity rollback deleted ObjectID=" & CStr(objectID) & ", ObjectName=[" & objectName & "], ObjectType=" & CStr(objectType) & ".")
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
    Err.Raise savedErrNumber, "RollbackSetupNavPaneObjectID", "Could not roll back a tracked Navigation Pane object identity." & vbCrLf & "ObjectID: " & CStr(objectID) & vbCrLf & "ObjectName: " & objectName & vbCrLf & "ObjectType: " & CStr(objectType) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub


' Purpose: Performs rollback setup nav pane group as the module's public orchestration step using the established ADC state and validation contract.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   groupID (Input; Long): Identifier selecting the logical Ribbon, form, catalog, or definition group affected by the operation.
'   groupName (Input; String): Name selecting the logical Ribbon, form, catalog, or definition group affected by the operation.
'   deletedCount (Output; Long): Receives the deleted count produced by this operation.
'   detailText (Output; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
' Returns: None (Sub). Results are returned through the documented ByRef parameters and/or the routine's database, UI, or module-state side effects.
Public Sub RollbackSetupNavPaneGroup(ByVal groupID As Long, ByVal groupName As String, ByRef deletedCount As Long, ByRef detailText As String)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim groupCount As Long
    Dim membershipCount As Long
    groupName = Trim$(groupName)
    If groupID <= 0 Or Len(groupName) = 0 Then Err.Raise vbObjectError + 6445, "RollbackSetupNavPaneGroup", "A Navigation Pane group rollback entry contained an invalid GroupID or group name."
    Set db = currentDb
    Set rs = db.OpenRecordset("SELECT Count(*) " & "AS " & SQL_FIELD_CNT & " FROM [" & NAV_PANE_GROUPS_TABLE & "] WHERE [ID]=" & CStr(groupID) & " AND UCase(Nz([Name],''))=" & SharedDao_SqlQuote(UCase$(groupName)) & ";", dbOpenSnapshot)
    groupCount = CLng(rs.fields(FIELD_CNT).value)
    rs.Close
    Set rs = Nothing
    If groupCount = 0 Then Exit Sub
    If groupCount <> 1 Then Err.Raise vbObjectError + 6446, "RollbackSetupNavPaneGroup", "The tracked Navigation Pane group identity is ambiguous during rollback."
    Set rs = db.OpenRecordset("SELECT Count(*) " & "AS " & SQL_FIELD_CNT & " FROM [" & NAV_PANE_GROUP_OBJECTS_TABLE & "] WHERE [GroupID]=" & CStr(groupID) & ";", dbOpenSnapshot)
    membershipCount = CLng(rs.fields(FIELD_CNT).value)
    rs.Close
    Set rs = Nothing
    If membershipCount <> 0 Then Err.Raise vbObjectError + 6447, "RollbackSetupNavPaneGroup", "The setup-created Navigation Pane group still contains memberships, so rollback refused to delete the group." & vbCrLf & "Group: " & groupName & vbCrLf & "GroupID: " & CStr(groupID) & vbCrLf & "Membership count: " & CStr(membershipCount)
    db.Execute "DELETE FROM [" & NAV_PANE_GROUPS_TABLE & "] WHERE [ID]=" & CStr(groupID) & " AND UCase(Nz([Name],''))=" & SharedDao_SqlQuote(UCase$(groupName)) & ";", dbFailOnError
    Set rs = db.OpenRecordset("SELECT Count(*) " & "AS " & SQL_FIELD_CNT & " FROM [" & NAV_PANE_GROUPS_TABLE & "] WHERE [ID]=" & CStr(groupID) & ";", dbOpenSnapshot)
    If CLng(rs.fields(FIELD_CNT).value) <> 0 Then Err.Raise vbObjectError + 6448, "RollbackSetupNavPaneGroup", "The setup-created Navigation Pane group remained present after rollback deletion."
    rs.Close
    Set rs = Nothing
    deletedCount = deletedCount + 1
    detailText = SharedText_AppendNonBlankLine(detailText, "Navigation Pane group rollback deleted [" & groupName & "] (GroupID=" & CStr(groupID) & ").")
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
    Err.Raise savedErrNumber, "RollbackSetupNavPaneGroup", "Could not roll back a setup-created Navigation Pane group." & vbCrLf & "Group: " & groupName & vbCrLf & "GroupID: " & CStr(groupID) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Function ReviewNavigationPaneMembershipExists(ByVal hostDb As DAO.Database, ByVal groupID As Long, ByVal objectID As Long) As Boolean
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Set rs = hostDb.OpenRecordset("SELECT Count(*) " & "AS " & SQL_FIELD_CNT & " FROM [" & NAV_PANE_GROUP_OBJECTS_TABLE & "] WHERE [GroupID]=" & CStr(groupID) & " AND [ObjectID]=" & CStr(objectID) & ";", dbOpenSnapshot)
    ReviewNavigationPaneMembershipExists = (CLng(rs.fields(FIELD_CNT).value) > 0)
    rs.Close
    Exit Function
ErrHandler:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    ReviewNavigationPaneMembershipExists = False
End Function

Private Function ReviewNextNavigationPaneMembershipPosition(ByVal hostDb As DAO.Database, ByVal groupID As Long) As Long
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    If SharedDao_TableFieldExists(currentDb, NAV_PANE_GROUP_OBJECTS_TABLE, "Position") = False Then Exit Function
    Set rs = hostDb.OpenRecordset("SELECT Max([Position]) AS MaxPosition FROM [" & NAV_PANE_GROUP_OBJECTS_TABLE & "] WHERE [GroupID]=" & CStr(groupID) & ";", dbOpenSnapshot)
    If rs.EOF = False Then ReviewNextNavigationPaneMembershipPosition = CLng(Nz(rs.fields("MaxPosition").value, 0)) + 1
    rs.Close
    Exit Function
ErrHandler:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    ReviewNextNavigationPaneMembershipPosition = 0
End Function

Private Function ReviewTryInsertNavigationPaneMembership(ByVal hostDb As DAO.Database, ByVal groupID As Long, ByVal objectID As Long, ByVal objectName As String, ByVal objectType As AcObjectType, ByVal insertMode As Long, ByRef createdByThisCall As Boolean, ByRef detailText As String) As Boolean
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim msysType As Long
    Dim nextPosition As Long
    createdByThisCall = False
    detailText = vbNullString
    If ReviewNavigationPaneMembershipExists(hostDb, groupID, objectID) Then
        ReviewTryInsertNavigationPaneMembership = True
        Exit Function
    End If
    msysType = CanonicalNavigationPaneWriteMSysType(objectType)
    nextPosition = ReviewNextNavigationPaneMembershipPosition(hostDb, groupID)
    Set rs = hostDb.OpenRecordset(NAV_PANE_GROUP_OBJECTS_TABLE, dbOpenDynaset)
    rs.AddNew
    rs.fields("GroupID").value = groupID
    rs.fields("ObjectID").value = objectID
    If insertMode <= 2 Then ReviewSetRecordsetFieldIfExists rs, "Name", objectName
    If insertMode = 1 Then
        If msysType <> 0 Then ReviewSetRecordsetFieldIfExists rs, "Type", msysType
        ReviewSetRecordsetFieldIfExists rs, "Flags", 0
        If nextPosition > 0 Then ReviewSetRecordsetFieldIfExists rs, "Position", nextPosition
    End If
    rs.Update
    rs.Close
    Set rs = Nothing
    If ReviewNavigationPaneMembershipExists(hostDb, groupID, objectID) Then
        createdByThisCall = True
        ReviewTryInsertNavigationPaneMembership = True
        Exit Function
    End If
    detailText = "Insert completed without leaving the expected GroupID/ObjectID membership."
    Exit Function
ErrHandler:
    detailText = "Mode " & CStr(insertMode) & " failed. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
    If ReviewNavigationPaneMembershipExists(hostDb, groupID, objectID) Then
        createdByThisCall = False
        ReviewTryInsertNavigationPaneMembership = True
    End If
End Function

Private Function ReviewNavigationPaneObjectIdentityExistsExact(ByVal hostDb As DAO.Database, ByVal objectID As Long, ByVal objectName As String, ByVal objectType As AcObjectType, ByRef detailText As String) As Boolean
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim exactCount As Long
    Dim msysType As Long
    If objectID = 0 Or Len(objectName) = 0 Then Exit Function
    msysType = CanonicalNavigationPaneWriteMSysType(objectType)
    Set rs = hostDb.OpenRecordset("SELECT Count(*) " & "AS " & SQL_FIELD_CNT & " FROM [" & NAV_PANE_OBJECT_IDS_TABLE & "] WHERE [ID]=" & CStr(objectID) & " AND UCase(Nz([Name],''))=" & SharedDao_SqlQuote(UCase$(objectName)) & " AND [Type]=" & CStr(msysType) & ";", dbOpenSnapshot)
    exactCount = CLng(rs.fields(FIELD_CNT).value)
    rs.Close
    Set rs = Nothing
    If exactCount > 1 Then Err.Raise vbObjectError + 8850, "ReviewNavigationPaneObjectIdentityExistsExact", "MSysNavPaneObjectIDs contains duplicate exact identities for [" & objectName & "]."
    ReviewNavigationPaneObjectIdentityExistsExact = (exactCount = 1)
    Exit Function
ErrHandler:
    detailText = "Could not verify exact MSysNavPaneObjectIDs identity for [" & objectName & "]. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
End Function

Private Sub ReviewTrackAndRemoveStaleNavigationPaneMemberships(ByVal hostDb As DAO.Database, ByVal objectID As Long, ByVal priorObjectName As String, ByVal priorObjectType As Long, ByVal replacementObjectName As String, ByRef removedMembershipCount As Long)
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    ' INTENTIONAL FILTERED PARTITION: membershipRows is a consumer-required subset/projection derived from the source enumerated by ReviewTrackAndRemoveStaleNavigationPaneMemberships; its membership is not
    ' independently authoritative.
    Dim membershipRows As New Collection
    Dim rowData As Variant
    Dim groupID As Long
    Dim previousGroupID As Long
    Dim capturedName As Variant
    Dim capturedType As Variant
    Dim capturedFlags As Variant
    Dim capturedPosition As Variant
    removedMembershipCount = 0
    Set rs = hostDb.OpenRecordset("SELECT * FROM [" & NAV_PANE_GROUP_OBJECTS_TABLE & "] WHERE [ObjectID]=" & CStr(objectID) & " ORDER BY [GroupID];", dbOpenSnapshot)
    Do While rs.EOF = False
        groupID = CLng(Nz(rs.fields("GroupID").value, 0))
        If groupID <= 0 Then Err.Raise vbObjectError + 8870, "ReviewTrackAndRemoveStaleNavigationPaneMemberships", "A stale Navigation Pane membership has an invalid GroupID for ObjectID=" & CStr(objectID) & "."
        If membershipRows.count > 0 And groupID = previousGroupID Then Err.Raise vbObjectError + 8871, "ReviewTrackAndRemoveStaleNavigationPaneMemberships", "Duplicate stale Navigation Pane memberships exist for GroupID=" & CStr(groupID) & ", ObjectID=" & CStr(objectID) & "."
        capturedName = ReviewRecordsetFieldValueIfExists(rs, "Name")
        capturedType = ReviewRecordsetFieldValueIfExists(rs, "Type")
        capturedFlags = ReviewRecordsetFieldValueIfExists(rs, "Flags")
        capturedPosition = ReviewRecordsetFieldValueIfExists(rs, "Position")
        membershipRows.Add Array(groupID, objectID, capturedName, capturedType, capturedFlags, capturedPosition)
        previousGroupID = groupID
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    If membershipRows.count = 0 Then Exit Sub
    For Each rowData In membershipRows
        TrackSetupNavPaneShortcutRestoreRollback CLng(rowData(0)), CLng(rowData(1)), rowData(2), rowData(3), rowData(4), rowData(5)
    Next rowData
    For Each rowData In membershipRows
        groupID = CLng(rowData(0))
        Set rs = hostDb.OpenRecordset("SELECT * FROM [" & NAV_PANE_GROUP_OBJECTS_TABLE & "] WHERE [GroupID]=" & CStr(groupID) & " AND [ObjectID]=" & CStr(objectID) & ";", dbOpenSnapshot)
        If rs.EOF Then Err.Raise vbObjectError + 8872, "ReviewTrackAndRemoveStaleNavigationPaneMemberships", "A tracked stale Navigation Pane membership disappeared before deletion. GroupID=" & CStr(groupID) & ", ObjectID=" & CStr(objectID) & "."
        If ReviewOptionalRecordsetFieldMatches(rs, "Name", rowData(2)) = False Or ReviewOptionalRecordsetFieldMatches(rs, "Type", rowData(3)) = False Or ReviewOptionalRecordsetFieldMatches(rs, "Flags", rowData(4)) = False Or ReviewOptionalRecordsetFieldMatches(rs, "Position", rowData(5)) = False Then Err.Raise vbObjectError + 8873, "ReviewTrackAndRemoveStaleNavigationPaneMemberships", "A stale Navigation Pane membership changed after rollback capture and before deletion. GroupID=" & CStr(groupID) & ", ObjectID=" & CStr(objectID) & "."
        rs.MoveNext
        If rs.EOF = False Then Err.Raise vbObjectError + 8874, "ReviewTrackAndRemoveStaleNavigationPaneMemberships", "A stale Navigation Pane membership became duplicated after rollback capture. GroupID=" & CStr(groupID) & ", ObjectID=" & CStr(objectID) & "."
        rs.Close
        Set rs = Nothing
        hostDb.Execute "DELETE FROM [" & NAV_PANE_GROUP_OBJECTS_TABLE & "] WHERE [GroupID]=" & CStr(groupID) & " AND [ObjectID]=" & CStr(objectID) & ";", dbFailOnError
        If ReviewNavigationPaneMembershipExists(hostDb, groupID, objectID) Then Err.Raise vbObjectError + 8875, "ReviewTrackAndRemoveStaleNavigationPaneMemberships", "A tracked stale Navigation Pane membership remained after deletion. GroupID=" & CStr(groupID) & ", ObjectID=" & CStr(objectID) & "."
        removedMembershipCount = removedMembershipCount + 1
    Next rowData
    LogTraceMessage TRACE_LEVEL_INFO, "NavigationPane", "ReviewTrackAndRemoveStaleNavigationPaneMemberships", "Removed stale Navigation Pane memberships only after durable exact rollback capture.", "ObjectID=" & CStr(objectID) & " | Prior identity=[" & priorObjectName & "], Type=" & CStr(priorObjectType) & " | Replacement identity=[" & replacementObjectName & "] | Memberships removed=" & CStr(removedMembershipCount), "Repair stale Navigation Pane memberships", 0, vbNullString, vbNullString, vbNullString, "NavigationPaneMembership", "CodeDb", "NavigationPaneGrouping"
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
    Err.Raise savedErrNumber, "ReviewTrackAndRemoveStaleNavigationPaneMemberships", "Could not safely remove stale Navigation Pane memberships before identity repair." & vbCrLf & "ObjectID: " & CStr(objectID) & vbCrLf & "Prior identity: [" & priorObjectName & "], Type=" & CStr(priorObjectType) & vbCrLf & "Replacement identity: [" & replacementObjectName & "]" & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub ReviewTrackAndRemovePreflightedBackupNavigationPaneMemberships(ByVal hostDb As DAO.Database, ByVal planRows As Collection, ByRef removedMembershipCount As Long)
    ' I3 batched Setup form of REVIEW's existing exclusive-placement authority: uses the canonical shortcut-match rule, captures exact non-Backup rows into durable rollback, then deletes them through one updatable
    ' recordset.
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: membershipRows holds planned/recovery work or independently observed validation evidence for ReviewTrackAndRemovePreflightedBackupNavigationPaneMemberships; keep it
    ' distinct from the source authority.
    Dim membershipRows As New Collection
    Dim seenRows As Object
    Dim planItem As Variant
    Dim rowData As Variant
    Dim groupID As Long
    Dim objectID As Long
    Dim keepGroupID As Long
    Dim capturedGroupID As Long
    Dim capturedObjectID As Long
    Dim capturedName As Variant
    Dim capturedType As Variant
    Dim capturedFlags As Variant
    Dim capturedPosition As Variant
    Dim objectName As String
    Dim criteriaText As String
    Dim criterion As String
    Dim rowKey As String
    removedMembershipCount = 0
    If planRows Is Nothing Then Exit Sub
    For Each planItem In planRows
        objectName = Trim$(CStr(planItem(0)))
        objectID = CLng(planItem(5))
        keepGroupID = CLng(planItem(4))
        If CLng(planItem(1)) = CLng(acTable) And IsBackupAccessObjectName(objectName) Then
            criterion = ReviewNavigationPaneObjectShortcutMatchCriteria(objectID, objectName)
            If Len(criterion) > 0 Then
                criterion = "((" & criterion & ") AND NOT ([GroupID]=" & CStr(keepGroupID) & " AND [ObjectID]=" & CStr(objectID) & "))"
                If Len(criteriaText) > 0 Then criteriaText = criteriaText & " OR "
                criteriaText = criteriaText & criterion
            End If
        End If
    Next planItem
    If Len(criteriaText) = 0 Then Exit Sub
    Set rs = hostDb.OpenRecordset("SELECT * FROM [" & NAV_PANE_GROUP_OBJECTS_TABLE & "] WHERE " & criteriaText & " ORDER BY [GroupID],[ObjectID];", dbOpenDynaset)
    If rs.EOF Then
        rs.Close
        Exit Sub
    End If
    ' INTENTIONAL KEYED LOOKUP: seenRows indexes captured Navigation Pane GroupID/ObjectID pairs solely to reject duplicate physical shortcut rows while the ordered membershipRows collection retains rollback
    ' evidence.
    Set seenRows = CreateObject("Scripting.Dictionary")
    seenRows.CompareMode = vbTextCompare
    Do While rs.EOF = False
        capturedGroupID = CLng(Nz(rs.fields("GroupID").value, 0))
        capturedObjectID = CLng(Nz(rs.fields("ObjectID").value, 0))
        If capturedGroupID <= 0 Or capturedObjectID = 0 Then Err.Raise vbObjectError + 8898, "ReviewTrackAndRemovePreflightedBackupNavigationPaneMemberships", "A non-Backup shortcut selected for backup exclusivity has an invalid GroupID/ObjectID."
        rowKey = CStr(capturedGroupID) & "|" & CStr(capturedObjectID)
        If seenRows.Exists(rowKey) Then Err.Raise vbObjectError + 8899, "ReviewTrackAndRemovePreflightedBackupNavigationPaneMemberships", "Duplicate Navigation Pane shortcut rows exist for GroupID=" & CStr(capturedGroupID) & ", ObjectID=" & CStr(capturedObjectID) & "."
        seenRows.Add rowKey, True
        capturedName = ReviewRecordsetFieldValueIfExists(rs, "Name")
        capturedType = ReviewRecordsetFieldValueIfExists(rs, "Type")
        capturedFlags = ReviewRecordsetFieldValueIfExists(rs, "Flags")
        capturedPosition = ReviewRecordsetFieldValueIfExists(rs, "Position")
        membershipRows.Add Array(capturedGroupID, capturedObjectID, capturedName, capturedType, capturedFlags, capturedPosition)
        rs.MoveNext
    Loop
    If membershipRows.count = 0 Then
        rs.Close
        Exit Sub
    End If
    If SetupRollbackTrackingActive() = False Then Err.Raise vbObjectError + 8900, "ReviewTrackAndRemovePreflightedBackupNavigationPaneMemberships", "Exclusive Setup backup Navigation Pane cleanup requires active durable rollback tracking."
    For Each rowData In membershipRows
        TrackSetupNavPaneShortcutRestoreRollback CLng(rowData(0)), CLng(rowData(1)), rowData(2), rowData(3), rowData(4), rowData(5)
    Next rowData
    rs.MoveFirst
    For Each rowData In membershipRows
        If rs.EOF Then Err.Raise vbObjectError + 8901, "ReviewTrackAndRemovePreflightedBackupNavigationPaneMemberships", "A captured non-Backup shortcut disappeared before rollback-safe deletion."
        groupID = CLng(Nz(rs.fields("GroupID").value, 0))
        objectID = CLng(Nz(rs.fields("ObjectID").value, 0))
        If groupID <> CLng(rowData(0)) Or objectID <> CLng(rowData(1)) Or ReviewOptionalRecordsetFieldMatches(rs, "Name", rowData(2)) = False Or ReviewOptionalRecordsetFieldMatches(rs, "Type", rowData(3)) = False Or ReviewOptionalRecordsetFieldMatches(rs, "Flags", rowData(4)) = False Or ReviewOptionalRecordsetFieldMatches(rs, "Position", rowData(5)) = False Then Err.Raise vbObjectError + 8902, "ReviewTrackAndRemovePreflightedBackupNavigationPaneMemberships", "A captured non-Backup shortcut changed before rollback-safe deletion. GroupID=" & CStr(CLng(rowData(0))) & ", ObjectID=" & CStr(CLng(rowData(1))) & "."
        rs.Delete
        removedMembershipCount = removedMembershipCount + 1
        rs.MoveNext
    Next rowData
    rs.Close
    Set rs = Nothing
    LogTraceMessage TRACE_LEVEL_INFO, "NavigationPane", "ReviewTrackAndRemovePreflightedBackupNavigationPaneMemberships", "Removed non-Backup shortcut memberships for transient ADC backup tables after exact durable rollback capture.", "Memberships removed=" & CStr(removedMembershipCount), "Enforce exclusive Backup-group placement", 0, vbNullString, vbNullString, vbNullString, "NavigationPaneMembership", "CodeDb", "I3"
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
    Err.Raise savedErrNumber, "ReviewTrackAndRemovePreflightedBackupNavigationPaneMemberships", "Could not safely enforce exclusive Backup-group placement for the preflighted ADC backup tables." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Function ReviewRemovePreflightedStaleNavigationPaneObjectIdentities(ByVal hostDb As DAO.Database, ByVal planRows As Collection, ByVal preMutationNavObjectCache As Object, ByRef removedIdentityCount As Long, ByRef removedMembershipCount As Long, ByRef detailText As String) As Boolean
    'Shared I3 mutation step: consumes the exact bulk mirror snapshot established by the immediately preceding staleness gate; tracked deletes are still verified after mutation.
    On Error GoTo ErrHandler
    Dim verifyRs As DAO.Recordset
    Dim planItem As Variant
    Dim objectName As String
    Dim objectType As AcObjectType
    Dim currentObjectID As Long
    Dim staleObjectID As Long
    Dim currentMappedObjectID As Long
    Dim navIdentityMappedIDAtI2 As Long
    Dim msysType As Long
    Dim staleMembershipCount As Long
    Dim verifyDetail As String
    Dim objectKey As String
    removedIdentityCount = 0
    removedMembershipCount = 0
    detailText = vbNullString
    If planRows Is Nothing Then Err.Raise vbObjectError + 8880, "ReviewRemovePreflightedStaleNavigationPaneObjectIdentities", "A preflighted Navigation Pane plan is required."
    If preMutationNavObjectCache Is Nothing Then Err.Raise vbObjectError + 8880, "ReviewRemovePreflightedStaleNavigationPaneObjectIdentities", "The fresh I3 Navigation Pane mirror snapshot is required."
    If SetupRollbackTrackingActive() = False Then Err.Raise vbObjectError + 8881, "ReviewRemovePreflightedStaleNavigationPaneObjectIdentities", "Stale Navigation Pane identity remapping requires active Setup rollback tracking."
    For Each planItem In planRows
        objectName = Trim$(CStr(planItem(0)))
        objectType = CLng(planItem(1))
        currentObjectID = CLng(planItem(5))
        navIdentityMappedIDAtI2 = CLng(planItem(9))
        msysType = CanonicalNavigationPaneWriteMSysType(objectType)
        objectKey = CStr(CLng(objectType)) & "|" & UCase$(objectName)
        currentMappedObjectID = 0
        If preMutationNavObjectCache.Exists(objectKey) Then currentMappedObjectID = CLng(preMutationNavObjectCache(objectKey))
        staleObjectID = 0
        If currentMappedObjectID <> 0 And currentMappedObjectID <> currentObjectID Then staleObjectID = currentMappedObjectID
        If staleObjectID <> 0 Then
            If navIdentityMappedIDAtI2 = 0 Or staleObjectID <> navIdentityMappedIDAtI2 Then Err.Raise vbObjectError + 8882, "ReviewRemovePreflightedStaleNavigationPaneObjectIdentities", "The fresh I3 mirror snapshot no longer matches the I2 stale identity evidence for [" & objectName & "]."
            verifyDetail = vbNullString
            If ReviewPhysicalObjectIdentityExistsExact(hostDb, currentObjectID, objectName, objectType, verifyDetail) = False Then
                If Len(verifyDetail) = 0 Then verifyDetail = "The current physical object does not own its planned ObjectID."
                Err.Raise vbObjectError + 8884, "ReviewRemovePreflightedStaleNavigationPaneObjectIdentities", "The stale Navigation Pane logical identity could not be remapped because current physical ownership was not exact for [" & objectName & "]." & vbCrLf & verifyDetail
            End If
            verifyDetail = vbNullString
            If ReviewPhysicalObjectIdentityExistsExact(hostDb, staleObjectID, objectName, objectType, verifyDetail) Then Err.Raise vbObjectError + 8885, "ReviewRemovePreflightedStaleNavigationPaneObjectIdentities", "The prior Navigation Pane ObjectID still belongs to the same current physical object and is not stale." & vbCrLf & "Object: [" & objectName & "]" & vbCrLf & "ObjectID: " & CStr(staleObjectID)
            If Len(verifyDetail) > 0 Then Err.Raise vbObjectError + 8885, "ReviewRemovePreflightedStaleNavigationPaneObjectIdentities", verifyDetail
            staleMembershipCount = 0
            ReviewTrackAndRemoveStaleNavigationPaneMemberships hostDb, staleObjectID, objectName, msysType, objectName, staleMembershipCount
            removedMembershipCount = removedMembershipCount + staleMembershipCount
            TrackSetupDeletedNavPaneObjectIDRestoreRollback staleObjectID, objectName, msysType
            hostDb.Execute "DELETE FROM [" & NAV_PANE_OBJECT_IDS_TABLE & "] WHERE [ID]=" & CStr(staleObjectID) & " AND [Type]=" & CStr(msysType) & " AND UCase(Nz([Name],''))=" & SharedDao_SqlQuote(UCase$(objectName)) & ";", dbFailOnError
            Set verifyRs = hostDb.OpenRecordset("SELECT Count(*) " & "AS " & SQL_FIELD_CNT & " FROM [" & NAV_PANE_OBJECT_IDS_TABLE & "] WHERE [ID]=" & CStr(staleObjectID) & " AND [Type]=" & CStr(msysType) & " AND UCase(Nz([Name],''))=" & SharedDao_SqlQuote(UCase$(objectName)) & ";", dbOpenSnapshot)
            If CLng(verifyRs.fields(FIELD_CNT).value) <> 0 Then Err.Raise vbObjectError + 8886, "ReviewRemovePreflightedStaleNavigationPaneObjectIdentities", "The stale Navigation Pane logical identity remained after tracked deletion for [" & objectName & "]."
            verifyRs.Close
            Set verifyRs = Nothing
            removedIdentityCount = removedIdentityCount + 1
            LogTraceMessage TRACE_LEVEL_INFO, "NavigationPane", "ReviewRemovePreflightedStaleNavigationPaneObjectIdentities", "Removed a stale same-name/type Navigation Pane identity before exact current-ID materialization.", "Object=[" & objectName & "] | Prior ObjectID=" & CStr(staleObjectID) & " | Current ObjectID=" & CStr(currentObjectID) & " | Stale memberships removed=" & CStr(staleMembershipCount), "Remap stale Navigation Pane logical identity", 0, vbNullString, vbNullString, vbNullString, "NavigationPaneObjectIdentity", "CodeDb", "NavigationPaneGrouping"
        End If
    Next planItem
    ReviewRemovePreflightedStaleNavigationPaneObjectIdentities = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not verifyRs Is Nothing Then verifyRs.Close
    On Error GoTo 0
    detailText = "Could not remove a preflighted stale Navigation Pane logical identity. Error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription
End Function

Private Function ReviewInsertPreflightedNavigationPaneObjectIdentity(ByVal hostDb As DAO.Database, ByVal objectID As Long, ByVal objectName As String, ByVal objectType As AcObjectType, ByRef createdByThisCall As Boolean, ByRef repairedByThisCall As Boolean, ByRef detailText As String) As Boolean
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim msysType As Long
    Dim verifyDetail As String
    Dim idConflictCount As Long
    Dim nameTypeConflictCount As Long
    Dim conflictID As Long
    Dim conflictName As String
    Dim conflictType As Long
    Dim priorName As String
    Dim priorType As Long
    Dim staleMembershipCount As Long
    Dim interruptedWriteRetries As Long
    Dim exactAfterInterruptedWrite As Boolean
    Dim probeErrNumber As Long
    Dim probeErrSource As String
    Dim probeErrDescription As String
    createdByThisCall = False
    repairedByThisCall = False
    detailText = vbNullString
    If objectID = 0 Or Len(objectName) = 0 Then Err.Raise vbObjectError + 8851, "ReviewInsertPreflightedNavigationPaneObjectIdentity", "A nonzero ObjectID and nonblank object name are required."
    msysType = CanonicalNavigationPaneWriteMSysType(objectType)
    If ReviewNavigationPaneObjectIdentityExistsExact(hostDb, objectID, objectName, objectType, verifyDetail) Then
        ReviewInsertPreflightedNavigationPaneObjectIdentity = True
        Exit Function
    End If
    If Len(verifyDetail) > 0 Then Err.Raise vbObjectError + 8852, "ReviewInsertPreflightedNavigationPaneObjectIdentity", verifyDetail
    'Capture every current collision; exact physical ID ownership below is authoritative for an occupied planned ObjectID.
    Set rs = hostDb.OpenRecordset("SELECT [ID],Nz([Name],'') " & "AS " & SQL_FIELD_OBJECT_NAME & ",Nz([Type],0) " & "AS " & SQL_FIELD_OBJECT_TYPE & " FROM [" & NAV_PANE_OBJECT_IDS_TABLE & "] WHERE [ID]=" & CStr(objectID) & " OR ([Type]=" & CStr(msysType) & " AND UCase(Nz([Name],''))=" & SharedDao_SqlQuote(UCase$(objectName)) & ") ORDER BY [ID];", dbOpenSnapshot)
    Do While rs.EOF = False
        conflictID = CLng(Nz(rs.fields("ID").value, 0))
        conflictName = CStr(rs.fields(FIELD_OBJECT_NAME).value)
        conflictType = CLng(rs.fields(FIELD_OBJECT_TYPE).value)
        If conflictID = objectID Then
            idConflictCount = idConflictCount + 1
            priorName = conflictName
            priorType = conflictType
        End If
        If conflictType = msysType And StrComp(conflictName, objectName, vbTextCompare) = 0 Then nameTypeConflictCount = nameTypeConflictCount + 1
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    'The same logical Navigation Pane identity mapped to a different ID remains a true ambiguity and must fail closed.
    If nameTypeConflictCount <> 0 Then Err.Raise vbObjectError + 8853, "ReviewInsertPreflightedNavigationPaneObjectIdentity", "A conflicting MSysNavPaneObjectIDs type/name identity appeared at a different ObjectID before materialization for [" & objectName & "]."
    If idConflictCount > 1 Then Err.Raise vbObjectError + 8862, "ReviewInsertPreflightedNavigationPaneObjectIdentity", "More than one MSysNavPaneObjectIDs row occupies the planned ObjectID=" & CStr(objectID) & " for [" & objectName & "]."
    If idConflictCount = 1 Then
        'Access can reuse an MSysObjects.ID after System Macros and Help Document Macros are rebuilt while leaving a stale Navigation Pane mirror row on that ID. Repair only after re-proving physical ownership at
        'the mutation boundary.
        verifyDetail = vbNullString
        If ReviewPhysicalObjectIdentityExistsExact(hostDb, objectID, objectName, objectType, verifyDetail) = False Then
            If Len(verifyDetail) = 0 Then verifyDetail = "The current MSysObjects row does not prove that the planned object owns ObjectID=" & CStr(objectID) & "."
            Err.Raise vbObjectError + 8863, "ReviewInsertPreflightedNavigationPaneObjectIdentity", "A stale-looking Navigation Pane ObjectID collision could not be repaired safely." & vbCrLf & verifyDetail
        End If
        If SetupRollbackTrackingActive() = False Then Err.Raise vbObjectError + 8864, "ReviewInsertPreflightedNavigationPaneObjectIdentity", "A stale Navigation Pane object identity requires in-place repair, but Setup rollback tracking is not active."
        ReviewTrackAndRemoveStaleNavigationPaneMemberships hostDb, objectID, priorName, priorType, objectName, staleMembershipCount
        'Persist the exact prior mirror identity before mutation. The restore action is durable across interrupted setup recovery; stale memberships were captured earlier so reverse rollback restores the identity
        'before its memberships.
        TrackSetupNavPaneObjectIDRestoreRollback objectID, objectName, msysType, priorName, priorType
        Set rs = hostDb.OpenRecordset("SELECT [ID],[Name],[Type] FROM [" & NAV_PANE_OBJECT_IDS_TABLE & "] WHERE [ID]=" & CStr(objectID) & ";", dbOpenDynaset)
        If rs.EOF Then Err.Raise vbObjectError + 8865, "ReviewInsertPreflightedNavigationPaneObjectIdentity", "The stale Navigation Pane ObjectID row disappeared before the verified in-place repair."
        If StrComp(CStr(Nz(rs.fields("Name").value, vbNullString)), priorName, vbBinaryCompare) <> 0 Or CLng(Nz(rs.fields("Type").value, 0)) <> priorType Then Err.Raise vbObjectError + 8866, "ReviewInsertPreflightedNavigationPaneObjectIdentity", "The stale Navigation Pane ObjectID row changed after classification and before repair."
        rs.Edit
        rs.fields("Name").value = objectName
        rs.fields("Type").value = msysType
        rs.Update
        rs.Close
        Set rs = Nothing
        verifyDetail = vbNullString
        If ReviewNavigationPaneObjectIdentityExistsExact(hostDb, objectID, objectName, objectType, verifyDetail) = False Then Err.Raise vbObjectError + 8867, "ReviewInsertPreflightedNavigationPaneObjectIdentity", "The repaired MSysNavPaneObjectIDs identity was not exact after in-place stale-row repair." & IIf(Len(verifyDetail) = 0, vbNullString, vbCrLf & verifyDetail)
        repairedByThisCall = True
        LogTraceMessage TRACE_LEVEL_INFO, "NavigationPane", "ReviewInsertPreflightedNavigationPaneObjectIdentity", "Repaired a stale Navigation Pane object identity in place after exact physical-ID proof and tracked stale-membership removal.", "ObjectID=" & CStr(objectID) & " | Prior Name=[" & priorName & "] | Prior Type=" & CStr(priorType) & " | Current Name=[" & objectName & "] | Current Type=" & CStr(msysType) & " | Stale memberships removed=" & CStr(staleMembershipCount), "Repair stale Navigation Pane mirror identity", 0, vbNullString, vbNullString, vbNullString, "NavigationPaneObjectIdentity", "CodeDb", "NavigationPaneGrouping"
        ReviewInsertPreflightedNavigationPaneObjectIdentity = True
        Exit Function
    End If
RetryCreateIdentity:
    Set rs = hostDb.OpenRecordset(NAV_PANE_OBJECT_IDS_TABLE, dbOpenDynaset)
    rs.AddNew
    rs.fields("ID").value = objectID
    rs.fields("Name").value = objectName
    rs.fields("Type").value = msysType
    rs.Update
    rs.Close
    Set rs = Nothing
    verifyDetail = vbNullString
    If ReviewNavigationPaneObjectIdentityExistsExact(hostDb, objectID, objectName, objectType, verifyDetail) = False Then Err.Raise vbObjectError + 8854, "ReviewInsertPreflightedNavigationPaneObjectIdentity", "The MSysNavPaneObjectIDs identity was not exact after materialization." & IIf(Len(verifyDetail) = 0, vbNullString, vbCrLf & verifyDetail)
    createdByThisCall = True
    ReviewInsertPreflightedNavigationPaneObjectIdentity = True
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
    Set rs = Nothing
    On Error GoTo 0
    If savedErrNumber = 3059 And interruptedWriteRetries = 0 And idConflictCount = 0 Then
        interruptedWriteRetries = 1
        verifyDetail = vbNullString
        exactAfterInterruptedWrite = False
        probeErrNumber = 0
        probeErrSource = vbNullString
        probeErrDescription = vbNullString
        Err.Clear
        On Error Resume Next
        exactAfterInterruptedWrite = ReviewNavigationPaneObjectIdentityExistsExact(hostDb, objectID, objectName, objectType, verifyDetail)
        probeErrNumber = Err.Number
        probeErrSource = Err.Source
        probeErrDescription = Err.Description
        Err.Clear
        On Error GoTo ErrHandler
        If probeErrNumber = 0 Then
            If exactAfterInterruptedWrite Then
                createdByThisCall = True
                On Error Resume Next
                LogTraceMessage TRACE_LEVEL_INFO, "NavigationPane", "ReviewInsertPreflightedNavigationPaneObjectIdentity", "Accepted an exact Navigation Pane identity after an interrupted DAO materialization returned error 3059.", "ObjectID=" & CStr(objectID) & " | Object=[" & objectName & "] | Type=" & CStr(msysType) & " | No duplicate retry was issued because the exact row was already present.", "Recover interrupted Navigation Pane identity materialization", 0, vbNullString, vbNullString, vbNullString, "NavigationPaneObjectIdentity", "CodeDb", "NavigationPaneGrouping"
                On Error GoTo ErrHandler
                ReviewInsertPreflightedNavigationPaneObjectIdentity = True
                Exit Function
            End If
            If Len(verifyDetail) > 0 Then
                detailText = "Could not safely retry interrupted MSysNavPaneObjectIDs materialization for [" & objectName & "] with ObjectID=" & CStr(objectID) & ". " & verifyDetail
                Exit Function
            End If
        End If
        DBEngine.Idle dbRefreshCache
        On Error Resume Next
        LogTraceMessage TRACE_LEVEL_INFO, "NavigationPane", "ReviewInsertPreflightedNavigationPaneObjectIdentity", "Navigation Pane identity materialization was interrupted once and will be retried without an event pump.", "ObjectID=" & CStr(objectID) & " | Object=[" & objectName & "] | Type=" & CStr(msysType) & " | Initial DAO error 3059 from " & savedErrSource & "." & IIf(probeErrNumber = 0, vbNullString, " Post-error exact-state probe error " & CStr(probeErrNumber) & " from " & probeErrSource & ": " & probeErrDescription), "Retry Navigation Pane identity materialization", 0, vbNullString, vbNullString, vbNullString, "NavigationPaneObjectIdentity", "CodeDb", "NavigationPaneGrouping"
        On Error GoTo ErrHandler
        GoTo RetryCreateIdentity
    End If
    If savedErrNumber = 3059 Then savedErrDescription = savedErrDescription & vbCrLf & "The single non-pumping materialization retry also failed."
    detailText = "Could not materialize or repair MSysNavPaneObjectIDs identity for [" & objectName & "] with ObjectID=" & CStr(objectID) & ". Error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription
End Function

Private Sub ReviewTrackNavigationPaneObjectIdentityRollbackOrUndo(ByVal hostDb As DAO.Database, ByVal objectID As Long, ByVal objectName As String, ByVal objectType As AcObjectType)
    'A materialized mirror identity must never remain untracked if durable rollback registration fails.
    On Error GoTo TrackErr
    TrackSetupNavPaneObjectIDRollback objectID, objectName, CanonicalNavigationPaneWriteMSysType(objectType)
    Exit Sub
TrackErr:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim cleanupErrNumber As Long
    Dim cleanupErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    hostDb.Execute "DELETE FROM [" & NAV_PANE_OBJECT_IDS_TABLE & "] WHERE [ID]=" & CStr(objectID) & " AND UCase(Nz([Name],''))=" & SharedDao_SqlQuote(UCase$(Trim$(objectName))) & " AND [Type]=" & CStr(CanonicalNavigationPaneWriteMSysType(objectType)) & ";", dbFailOnError
    cleanupErrNumber = Err.Number
    cleanupErrDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If cleanupErrNumber <> 0 Then Err.Raise vbObjectError + 8855, "ReviewTrackNavigationPaneObjectIdentityRollbackOrUndo", "Rollback registration failed after a Navigation Pane identity insert, and immediate cleanup also failed." & vbCrLf & "ObjectID: " & CStr(objectID) & vbCrLf & "Object: [" & objectName & "]" & vbCrLf & "Registration error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription & vbCrLf & "Cleanup error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription
    Err.Raise savedErrNumber, "ReviewTrackNavigationPaneObjectIdentityRollbackOrUndo", "Rollback registration failed after a Navigation Pane identity insert. The inserted identity was removed immediately." & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Original description: " & savedErrDescription
End Sub

Private Function ReviewEnsureNavigationPaneMembership(ByVal hostDb As DAO.Database, ByVal groupID As Long, ByVal objectID As Long, ByVal objectName As String, ByVal objectType As AcObjectType, ByRef createdByThisCall As Boolean, ByRef detailText As String) As Boolean
    Dim modeCreated As Boolean
    Dim modeDetail As String
    Dim allDetails As String
    Dim insertMode As Long
    createdByThisCall = False
    For insertMode = 1 To 3
        modeCreated = False
        modeDetail = vbNullString
        If ReviewTryInsertNavigationPaneMembership(hostDb, groupID, objectID, objectName, objectType, insertMode, modeCreated, modeDetail) Then
            createdByThisCall = modeCreated
            ReviewEnsureNavigationPaneMembership = True
            Exit Function
        End If
        allDetails = SharedText_AppendNonBlankLine(allDetails, modeDetail)
    Next insertMode
    detailText = allDetails
End Function

Private Sub ReviewTrackNavigationPaneShortcutRollbackOrUndo(ByVal hostDb As DAO.Database, ByVal groupID As Long, ByVal objectID As Long)
    'A successful insert must never remain untracked if durable rollback registration fails.
    On Error GoTo TrackErr
    TrackSetupNavPaneShortcutRollback groupID, objectID
    Exit Sub
TrackErr:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim cleanupErrNumber As Long
    Dim cleanupErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    hostDb.Execute "DELETE FROM [" & NAV_PANE_GROUP_OBJECTS_TABLE & "] WHERE [GroupID]=" & CStr(groupID) & " AND [ObjectID]=" & CStr(objectID) & ";", dbFailOnError
    cleanupErrNumber = Err.Number
    cleanupErrDescription = Err.Description
    Err.Clear
    If cleanupErrNumber = 0 Then
        If ReviewNavigationPaneMembershipExists(hostDb, groupID, objectID) Then
            cleanupErrNumber = vbObjectError + 8834
            cleanupErrDescription = "The newly inserted Navigation Pane membership remained present after emergency cleanup."
        End If
    End If
    On Error GoTo 0
    If cleanupErrNumber <> 0 Then Err.Raise vbObjectError + 8835, "ReviewTrackNavigationPaneShortcutRollbackOrUndo", "Rollback registration failed after a Navigation Pane insert, and immediate cleanup also failed." & vbCrLf & "GroupID: " & CStr(groupID) & vbCrLf & "ObjectID: " & CStr(objectID) & vbCrLf & "Registration error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription & vbCrLf & "Cleanup error " & CStr(cleanupErrNumber) & ": " & cleanupErrDescription
    Err.Raise savedErrNumber, "ReviewTrackNavigationPaneShortcutRollbackOrUndo", "Rollback registration failed after a Navigation Pane insert. The inserted membership was removed immediately." & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Original description: " & savedErrDescription
End Sub

Private Function ReviewNavigationPanePlanStillCurrent(ByRef detailText As String, ByRef currentNavObjectCache As Object) As Boolean
    'Read-only I3 gate: rebuilds fresh physical group/object state immediately before I3 mutation; each unique destination group is compared once.
    On Error GoTo ErrHandler
    Dim hostDb As DAO.Database
    Dim objectCache As Object
    Dim navObjectCache As Object
    Dim physicalGroupSnapshot As Object
    Dim checkedGroups As Object
    Dim planItem As Variant
    Dim objectName As String
    Dim objectType As AcObjectType
    Dim groupName As String
    Dim expectedGroupID As Long
    Dim expectedObjectID As Long
    Dim actualGroupID As Long
    Dim groupCount As Long
    Dim objectKey As String
    Dim groupKey As String
    Dim navIdentityMappedIDAtI2 As Long
    Dim currentNavIdentityID As Long
    Dim failureText As String
    Set currentNavObjectCache = Nothing
    If mSetupNavigationPanePlanReady = False Or mSetupNavigationPanePlan Is Nothing Then
        detailText = "No same-session I2 plan is available."
        Exit Function
    End If
    Set hostDb = currentDb
    Set objectCache = ReviewNavigationPaneObjectIDCache(hostDb)
    Set navObjectCache = ReviewNavigationPaneObjectMappingCache(hostDb)
    Set currentNavObjectCache = navObjectCache
    Set physicalGroupSnapshot = ReviewPhysicalNavigationPaneGroupSnapshot()
    ' INTENTIONAL KEYED LOOKUP: checkedGroups provides keyed membership/lookup over source rows or state already read by ReviewNavigationPanePlanStillCurrent; it is a derived access path, not a second authority.
    Set checkedGroups = CreateObject("Scripting.Dictionary")
    checkedGroups.CompareMode = vbTextCompare
    For Each planItem In mSetupNavigationPanePlan
        objectName = CStr(planItem(0))
        objectType = CLng(planItem(1))
        groupName = CStr(planItem(3))
        expectedGroupID = CLng(planItem(4))
        expectedObjectID = CLng(planItem(5))
        navIdentityMappedIDAtI2 = CLng(planItem(9))
        groupKey = UCase$(Trim$(groupName))
        If checkedGroups.Exists(groupKey) = False Then
            actualGroupID = 0
            groupCount = ReviewPhysicalNavigationPaneGroupFromSnapshot(physicalGroupSnapshot, groupName, actualGroupID)
            If groupCount <> 1 Or actualGroupID <> expectedGroupID Then failureText = SharedText_AppendNonBlankLine(failureText, "Group identity changed: [" & groupName & "]; I2 GroupID=" & CStr(expectedGroupID) & "; current GroupID=" & CStr(actualGroupID) & "; count=" & CStr(groupCount) & ".")
            checkedGroups.Add groupKey, True
        End If
        objectKey = CStr(CLng(objectType)) & "|" & UCase$(objectName)
        If objectCache.Exists(objectKey) = False Then
            failureText = SharedText_AppendNonBlankLine(failureText, "Object identity disappeared: " & CStr(planItem(2)) & " [" & objectName & "].")
        ElseIf CLng(objectCache(objectKey)) <> expectedObjectID Then
            failureText = SharedText_AppendNonBlankLine(failureText, "Object identity changed: " & CStr(planItem(2)) & " [" & objectName & "]; I2 ObjectID=" & CStr(expectedObjectID) & "; current ObjectID=" & CStr(objectCache(objectKey)) & ".")
        End If
        currentNavIdentityID = 0
        If navObjectCache.Exists(objectKey) Then currentNavIdentityID = CLng(navObjectCache(objectKey))
        If currentNavIdentityID = expectedObjectID Then
            'Exact identity may have appeared concurrently after a missing or stale I2 observation; I3 can verify and retain it.
        ElseIf currentNavIdentityID = 0 Then
            If navIdentityMappedIDAtI2 <> 0 Then failureText = SharedText_AppendNonBlankLine(failureText, "Navigation Pane object identity disappeared after I2: " & CStr(planItem(2)) & " [" & objectName & "]; I2 mapped ID=" & CStr(navIdentityMappedIDAtI2) & ".")
        ElseIf navIdentityMappedIDAtI2 = 0 Then
            failureText = SharedText_AppendNonBlankLine(failureText, "A mismatched Navigation Pane object identity appeared after I2: " & CStr(planItem(2)) & " [" & objectName & "]; current MSysNavPaneObjectIDs.ID=" & CStr(currentNavIdentityID) & ".")
        ElseIf currentNavIdentityID <> navIdentityMappedIDAtI2 Then
            failureText = SharedText_AppendNonBlankLine(failureText, "Navigation Pane object identity changed after I2: " & CStr(planItem(2)) & " [" & objectName & "]; I2 mapped ID=" & CStr(navIdentityMappedIDAtI2) & "; current mapped ID=" & CStr(currentNavIdentityID) & ".")
        End If
    Next planItem
    detailText = failureText
    If Len(detailText) = 0 Then detailText = "I3 staleness proof used one fresh physical-group bulk snapshot and compared unique destination groups=" & CStr(checkedGroups.count) & "."
    ReviewNavigationPanePlanStillCurrent = (Len(failureText) = 0)
    Exit Function
ErrHandler:
    detailText = "I3 staleness validation failed. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    ReviewNavigationPanePlanStillCurrent = False
End Function

Private Sub ReviewRefreshNavigationPaneOnce()
    On Error Resume Next
    Application.RefreshDatabaseWindow
    DoEvents
    On Error GoTo 0
End Sub

Private Function TryResolveNavigationPanePlacementRow( _
    ByVal rs As DAO.Recordset, ByVal groupCache As Object, ByVal systemMacroGroupCache As Object, ByVal objectIdCache As Object, ByVal h2Preflight As Boolean, _
    ByRef objectName As String, ByRef objectTypeText As String, ByRef SettingKey As String, ByRef groupName As String, ByRef objectType As AcObjectType, _
    ByRef groupID As Long, ByRef objectID As Long, ByRef objectKey As String, ByRef optionalMissing As Boolean, ByRef failureDetail As String) As Boolean
    'Resolves the phase-neutral row identity and authoritative destination used by both I2 planning and I3 proof; it does not inspect or mutate Navigation Pane mirror/membership state.
    Dim categoryCode As String
    Dim expectedExistencePolicy As String
    Dim setupAction As String
    Dim groupData As Variant
    Dim systemMacroGroupData As Variant
    objectName = CStr(Nz(rs.fields(FIELD_OBJECT_NAME).value, vbNullString))
    objectTypeText = Trim$(CStr(Nz(rs.fields(FIELD_RESOLVED_OBJECT_TYPE).value, vbNullString)))
    categoryCode = Trim$(CStr(Nz(rs.fields(FIELD_CATEGORY_CODE).value, vbNullString)))
    expectedExistencePolicy = Trim$(CStr(Nz(rs.fields(FIELD_RESOLVED_EXPECTED_EXISTENCE_POLICY).value, vbNullString)))
    setupAction = Trim$(CStr(Nz(rs.fields(FIELD_RESOLVED_SETUP_ACTION).value, vbNullString)))
    SettingKey = Trim$(CStr(Nz(rs.fields(FIELD_RESOLVED_NAVIGATION_PANE_GROUP_SETTING_KEY).value, vbNullString)))
    groupName = vbNullString
    groupID = 0
    objectID = 0
    objectKey = vbNullString
    optionalMissing = False
    failureDetail = vbNullString
    If Len(objectName) = 0 Or Len(objectTypeText) = 0 Then
        If h2Preflight Then failureDetail = "I2 encountered a grouped catalog row with a blank object name or resolved object type." Else failureDetail = "Grouped catalog row has a blank object name or resolved object type."
        Exit Function
    End If
    If StrComp(categoryCode, ADCAT_CATEGORY_SYSTEM_MACRO, vbTextCompare) = 0 Then
        If systemMacroGroupCache.Exists(objectName) = False Then
            If h2Preflight Then failureDetail = "System Macro and Help Document Macro [" & objectName & "] has no authoritative Navigation Pane group mapping." Else failureDetail = "System macro [" & objectName & "] has no authoritative SystemMacroGroupKey mapping."
            Exit Function
        End If
        systemMacroGroupData = systemMacroGroupCache(objectName)
        SettingKey = Trim$(CStr(systemMacroGroupData(1)))
        If h2Preflight Then
            If CLng(systemMacroGroupData(0)) <> 1 Or Len(SettingKey) = 0 Then
                failureDetail = "System Macro and Help Document Macro [" & objectName & "] has an ambiguous or blank authoritative Navigation Pane group key."
                Exit Function
            End If
            If CLng(systemMacroGroupData(2)) <> 1 Then
                failureDetail = "System Macro and Help Document Macro [" & objectName & "] has " & CStr(systemMacroGroupData(2)) & " active UI_Groups definitions for group key [" & SettingKey & "]."
                Exit Function
            End If
        ElseIf CLng(systemMacroGroupData(0)) <> 1 Or Len(SettingKey) = 0 Or CLng(systemMacroGroupData(2)) <> 1 Then
            failureDetail = "System macro [" & objectName & "] does not have one authoritative Navigation Pane group definition."
            Exit Function
        End If
        groupName = Trim$(CStr(systemMacroGroupData(3)))
        groupID = CLng(systemMacroGroupData(5))
        If Len(groupName) = 0 Or CLng(systemMacroGroupData(4)) <> 1 Or groupID <= 0 Then
            If h2Preflight Then failureDetail = "System Macro and Help Document Macro [" & objectName & "] destination group key [" & SettingKey & "] is not uniquely present physically." Else failureDetail = "System macro [" & objectName & "] destination group key [" & SettingKey & "] is not uniquely present physically."
            Exit Function
        End If
    Else
        If h2Preflight Then
            If Len(SettingKey) = 0 Then
                failureDetail = objectTypeText & " [" & objectName & "] has no resolved Navigation Pane group setting key."
                Exit Function
            End If
            If groupCache.Exists(SettingKey) = False Then
                failureDetail = objectTypeText & " [" & objectName & "] references missing active group setting key [" & SettingKey & "]."
                Exit Function
            End If
        ElseIf Len(SettingKey) = 0 Or groupCache.Exists(SettingKey) = False Then
            failureDetail = objectTypeText & " [" & objectName & "] has no unique active Navigation Pane group setting."
            Exit Function
        End If
        groupData = groupCache(SettingKey)
        If CLng(groupData(0)) <> 1 Then
            If h2Preflight Then failureDetail = objectTypeText & " [" & objectName & "] has ambiguous group setting key [" & SettingKey & "] with " & CStr(groupData(0)) & " active definitions." Else failureDetail = objectTypeText & " [" & objectName & "] has ambiguous Navigation Pane group definitions for [" & SettingKey & "]."
            Exit Function
        End If
        groupName = Trim$(CStr(groupData(1)))
        groupID = CLng(groupData(3))
        If Len(groupName) = 0 Or CLng(groupData(2)) <> 1 Or groupID <= 0 Then
            failureDetail = objectTypeText & " [" & objectName & "] destination group [" & groupName & "] is not uniquely present physically."
            Exit Function
        End If
    End If
    objectType = AccessObjectTypeFromCode(objectTypeText)
    objectKey = CStr(CLng(objectType)) & "|" & UCase$(objectName)
    If objectIdCache.Exists(objectKey) = False Then
        If ManagedObjectMayBeMissing(expectedExistencePolicy, setupAction) Then
            optionalMissing = True
            Exit Function
        End If
        failureDetail = "Required " & objectTypeText & " [" & objectName & "] has no current MSysObjects identity."
        Exit Function
    End If
    objectID = CLng(objectIdCache(objectKey))
    If objectID = 0 Then
        If h2Preflight Then failureDetail = "Required " & objectTypeText & " [" & objectName & "] resolved an invalid MSysObjects ID." Else failureDetail = "Required " & objectTypeText & " [" & objectName & "] does not have matching physical and Navigation Pane object identity."
        Exit Function
    End If
    TryResolveNavigationPanePlacementRow = True
End Function

Private Function CompleteNavigationPanePlacementCurrentForSetup(ByRef detailText As String) As Boolean
    'Read-only durable I3 proof for J1/standalone successors; does not populate or mutate the I2 same-session plan.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim hostDb As DAO.Database
    Dim rs As DAO.Recordset
    Dim objectIdCache As Object
    Dim navObjectIdCache As Object
    Dim membershipCache As Object
    Dim physicalGroupSnapshot As Object
    Dim groupCache As Object
    Dim groupKeyCache As Object
    Dim systemMacroGroupCache As Object
    Dim planRows As Collection
    Dim objectName As String
    Dim objectTypeText As String
    Dim SettingKey As String
    Dim groupName As String
    Dim objectType As AcObjectType
    Dim groupID As Long
    Dim objectID As Long
    Dim objectKey As String
    Dim rowOptionalMissing As Boolean
    Dim rowFailureDetail As String
    Dim membershipKey As String
    Dim currentToken As String
    Dim storedToken As String
    Dim propertyExists As Boolean
    Dim failureCount As Long
    Dim optionalMissingCount As Long
    Dim checkedCount As Long
    Dim backupAlreadyPresentCount As Long
    Dim backupIdentityMaterializationCount As Long
    Dim backupIdentityRemapCount As Long
    Dim failureText As String
    Dim configuredCurrent As Boolean
    Dim sqlText As String
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: planRows holds planned/recovery work or independently observed validation evidence for CompleteNavigationPanePlacementCurrentForSetup; keep it distinct from the
    ' source authority.
    Set planRows = New Collection
    Set hostDb = currentDb
    On Error Resume Next
    Err.Clear
    storedToken = CStr(Nz(hostDb.properties(NAVIGATION_PANE_MEMBERSHIP_TOKEN_PROPERTY).value, vbNullString))
    propertyExists = (Err.Number = 0)
    Err.Clear
    On Error GoTo ErrHandler
    If propertyExists = False Or Len(storedToken) = 0 Then
        detailText = "I3 complete managed-object Navigation Pane state token has not been persisted."
        Exit Function
    End If
    If SettingBool(SETTING_ADD_TO_NAV_GROUP, True) = False Then
        currentToken = ReviewNavigationPanePlanToken(planRows)
        GoTo CompareStoredToken
    End If
    Set physicalGroupSnapshot = ReviewPhysicalNavigationPaneGroupSnapshot()
    ReviewNavigationPaneGroupDefinitionCaches physicalGroupSnapshot, groupCache, groupKeyCache, configuredCurrent, failureText
    If configuredCurrent = False Then
        detailText = "I3 configured-group prerequisite is not current." & vbCrLf & failureText
        Exit Function
    End If
    Set db = ADCDb()
    Set objectIdCache = ReviewNavigationPaneObjectIDCache(hostDb)
    Set navObjectIdCache = ReviewNavigationPaneObjectMappingCache(hostDb)
    Set membershipCache = ReviewNavigationPaneMembershipCache(hostDb)
    Set systemMacroGroupCache = ReviewSystemMacroNavigationPaneGroupCache(groupKeyCache)
    sqlText = "SELECT " & SQL_FIELD_OBJECT_NAME & "," & SQL_FIELD_RESOLVED_OBJECT_TYPE & ",[" & FIELD_CATEGORY_CODE & "]," & SQL_FIELD_RESOLVED_EXPECTED_EXISTENCE_POLICY & "," & SQL_FIELD_RESOLVED_SETUP_ACTION & "," & SQL_FIELD_RESOLVED_NAVIGATION_PANE_GROUP_SETTING_KEY & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_ManagedObjectsResolved") & "] " & _
              "WHERE Nz(" & SQL_FIELD_OBJECT_CATEGORY_STATUS & "," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & " AND Nz([" & FIELD_MANAGED_OBJECT_STATUS & "]," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & " " & _
              "AND UCase(Trim(Nz(" & SQL_FIELD_RESOLVED_NAV_PANE_VISIBILITY_POLICY & ",'')))=" & SharedDao_SqlQuote(UCase$(ADCAT_NAV_POLICY_ADD_TO_CONFIGURED_GROUP)) & " " & _
              "ORDER BY " & OrderByPolicyText("REVIEW_PLACEMENT") & ";"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    Do While rs.EOF = False
        PumpSetupResponsiveness
        If TryResolveNavigationPanePlacementRow(rs, groupCache, systemMacroGroupCache, objectIdCache, False, objectName, objectTypeText, SettingKey, groupName, objectType, groupID, objectID, objectKey, rowOptionalMissing, rowFailureDetail) Then
            If navObjectIdCache.Exists(objectKey) = False Or CLng(navObjectIdCache(objectKey)) <> objectID Then
                failureCount = failureCount + 1
                failureText = SharedText_AppendNonBlankLine(failureText, "Required " & objectTypeText & " [" & objectName & "] does not have matching physical and Navigation Pane object identity.")
            Else
                membershipKey = CStr(groupID) & "|" & CStr(objectID)
                If membershipCache.Exists(membershipKey) = False Then
                    failureCount = failureCount + 1
                    failureText = SharedText_AppendNonBlankLine(failureText, "Required I3 membership is missing: " & objectTypeText & " [" & objectName & "] -> [" & groupName & "].")
                End If
                planRows.Add Array(objectName, CLng(objectType), objectTypeText, groupName, groupID, objectID, True, SettingKey)
                checkedCount = checkedCount + 1
            End If
        ElseIf rowOptionalMissing Then
            optionalMissingCount = optionalMissingCount + 1
        Else
            failureCount = failureCount + 1
            failureText = SharedText_AppendNonBlankLine(failureText, rowFailureDetail)
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    ReviewAppendDynamicBackupNavigationPanePlanRows hostDb, objectIdCache, navObjectIdCache, membershipCache, physicalGroupSnapshot, True, planRows, checkedCount, backupAlreadyPresentCount, backupIdentityMaterializationCount, backupIdentityRemapCount, failureCount, failureText
    If failureCount > 0 Then
        detailText = "I3 complete managed-object Navigation Pane state is not current. Blocking findings=" & CStr(failureCount) & "." & vbCrLf & failureText
        Exit Function
    End If
    currentToken = ReviewNavigationPanePlanToken(planRows)
CompareStoredToken:
    If StrComp(storedToken, currentToken, vbBinaryCompare) <> 0 Then
        detailText = "I3 complete managed-object Navigation Pane state token is stale." & vbCrLf & "Stored: " & storedToken & vbCrLf & "Current: " & currentToken
        Exit Function
    End If
    detailText = "I3 complete managed-object Navigation Pane state is current. Checked=" & CStr(checkedCount) & "; optional missing=" & CStr(optionalMissingCount) & "; token=" & currentToken
    CompleteNavigationPanePlacementCurrentForSetup = True
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
    detailText = "Could not verify the durable I3 complete managed-object Navigation Pane state. Error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription
    CompleteNavigationPanePlacementCurrentForSetup = False
End Function

' Purpose: Returns the authoritative complete-domain I3 physical placement token. Same-run reuse verifies the retained I3 rows against current physical state before accepting the persisted token.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   detailText (Input/output; String): Receives or supplies diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
'   allowSameRunProofReuse (Input Optional; Boolean): When True, enables or requests same run proof reuse for this call; when False, the routine follows its normal non-forced/default path. Default: False.
' Returns: String: Returns the deterministic complete navigation pane placement token for setup token/signature produced from the supplied state, allowing callers to compare exact state without repeating the
'   underlying scan/build work.
Public Function CompleteNavigationPanePlacementTokenForSetup(ByRef detailText As String, Optional ByVal allowSameRunProofReuse As Boolean = False) As String
    'Returns the authoritative complete-domain I3 physical placement token. Same-run reuse verifies the retained I3 rows against current physical state before accepting the persisted token.
    On Error GoTo ErrHandler
    Dim hostDb As DAO.Database
    Dim propertyExists As Boolean
    Dim storedToken As String
    Dim proofDetail As String
    Dim proofReused As Boolean
    Dim sharedObjectCache As Object
    Dim sharedNavObjectCache As Object
    Dim sharedMembershipCache As Object
    Dim sharedPhysicalGroupSnapshot As Object
    If allowSameRunProofReuse Then
        proofReused = ReviewSameRunNavigationPaneProofCurrent(mCompleteNavigationPaneProofRows, mCompleteNavigationPaneProofToken, mCompleteNavigationPaneProofTraceRunID, mCompleteNavigationPaneProofHostDbName, "I3 complete Navigation Pane", sharedObjectCache, sharedNavObjectCache, sharedMembershipCache, sharedPhysicalGroupSnapshot, proofDetail)
    End If
    If proofReused = False Then
        If CompleteNavigationPanePlacementCurrentForSetup(proofDetail) = False Then detailText = proofDetail: Exit Function
    Else
        mNavPaneSameRunProofReuseCount = mNavPaneSameRunProofReuseCount + 1
    End If
    Set hostDb = currentDb
    storedToken = DatabaseTextPropertyValue(hostDb, NAVIGATION_PANE_MEMBERSHIP_TOKEN_PROPERTY, propertyExists)
    If propertyExists = False Or Len(storedToken) = 0 Then detailText = "The persisted I3 complete Navigation Pane exact token is missing.": Exit Function
    If proofReused Then
        If StrComp(storedToken, mCompleteNavigationPaneProofToken, vbBinaryCompare) <> 0 Then detailText = "The retained same-run I3 exact proof does not match the persisted I3 token.": Exit Function
    End If
    CompleteNavigationPanePlacementTokenForSetup = storedToken
    detailText = "I3 complete managed-object Navigation Pane placement is current; token=" & storedToken & "; same-run proof reused=" & CStr(proofReused) & "." & vbCrLf & proofDetail
    Exit Function
ErrHandler:
    If IsCancellationError(Err.Number) Then Err.Raise Err.Number, Err.Source, Err.Description
    detailText = "Could not verify the I3 complete Navigation Pane exact token. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    CompleteNavigationPanePlacementTokenForSetup = vbNullString
End Function


Private Function ReviewGeneratedFormNavigationPaneGroupNameByKey(ByVal groupKey As String) As String
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim resolvedName As String
    If Len(groupKey) = 0 Then Err.Raise vbObjectError + 6696, "ReviewGeneratedFormNavigationPaneGroupNameByKey", "The authoritative Navigation Pane group key is blank."
    Set db = currentDb
    If SharedDao_TableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups")) = False Then
        If StrComp(groupKey, DEFAULT_NAVIGATION_PANE_GROUP_KEY_GENERATED_FORMS, vbTextCompare) = 0 Then
            resolvedName = Trim$(SettingText(SETTING_DEFAULT_FORMS_NAV_GROUP, vbNullString))
            If Len(resolvedName) = 0 Then resolvedName = BuildNavigationPaneGroupName(DEFAULT_GENERATED_FORMS_NAVIGATION_PANE_GROUP_LOGICAL_NAME, False)
            ReviewGeneratedFormNavigationPaneGroupNameByKey = resolvedName
            Exit Function
        End If
        Err.Raise vbObjectError + 6697, "ReviewGeneratedFormNavigationPaneGroupNameByKey", "The authoritative Navigation Pane group table does not exist." & vbCrLf & "Group key: " & groupKey
    End If
    If SharedDao_TableFieldExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups"), FIELD_NAVIGATION_PANE_GROUP_KEY) = False Then Err.Raise vbObjectError + 6698, "ReviewGeneratedFormNavigationPaneGroupNameByKey", "Required field is missing: NavigationPaneGroupKey"
    If SharedDao_TableFieldExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups"), FIELD_NAVIGATION_PANE_GROUP_NAME) = False Then Err.Raise vbObjectError + 6699, "ReviewGeneratedFormNavigationPaneGroupNameByKey", "Required field is missing: NavigationPaneGroupName"
    sqlText = "SELECT [" & FIELD_NAVIGATION_PANE_GROUP_NAME & "] FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups") & "] WHERE [" & FIELD_NAVIGATION_PANE_GROUP_KEY & "]=" & SharedDao_SqlQuote(groupKey)
    If SharedDao_TableFieldExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups"), FIELD_LOOKUP_ACTIVE) Then sqlText = sqlText & " AND Nz(" & SQL_FIELD_LOOKUP_ACTIVE & ",True)<>False"
    If SharedDao_TableFieldExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups"), FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER) Then sqlText = sqlText & " ORDER BY " & OrderByPolicyText("REVIEW_NAVIGATION_GROUP") Else sqlText = sqlText & " ORDER BY [" & FIELD_NAVIGATION_PANE_GROUP_NAME & "]"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    If rs.EOF Then Err.Raise vbObjectError + 6700, "ReviewGeneratedFormNavigationPaneGroupNameByKey", "No active authoritative Navigation Pane group row exists." & vbCrLf & "Group key: " & groupKey
    resolvedName = Trim$(CStr(Nz(rs.fields(FIELD_NAVIGATION_PANE_GROUP_NAME).value, vbNullString)))
    If Len(resolvedName) = 0 Then Err.Raise vbObjectError + 6701, "ReviewGeneratedFormNavigationPaneGroupNameByKey", "The authoritative NavigationPaneGroupName is blank." & vbCrLf & "Group key: " & groupKey
    rs.MoveNext
    If Not rs.EOF Then Err.Raise vbObjectError + 6702, "ReviewGeneratedFormNavigationPaneGroupNameByKey", "More than one active authoritative Navigation Pane group row exists." & vbCrLf & "Group key: " & groupKey
    ReviewGeneratedFormNavigationPaneGroupNameByKey = resolvedName
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
    Err.Raise savedErrNumber, "ReviewGeneratedFormNavigationPaneGroupNameByKey", "Could not resolve an authoritative generated-form Navigation Pane group." & vbCrLf & "Table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups") & vbCrLf & "Group key: " & groupKey & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

' Purpose: Applies generated forms navigation pane for form names, nav notes.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   formNames (Input; Collection): Collection containing the form names items that the routine consumes or updates as part of this operation.
'   navNotes (Input/output; Collection): Collection supplying the nav notes consumed by this procedure; items are read in the collection's existing order unless the body explicitly keys or sorts them.
' Returns: None (Sub). Results are returned through the documented ByRef parameters and/or the routine's database, UI, or module-state side effects.
Public Sub ApplyGeneratedFormsNavigationPane(ByVal formNames As Collection, ByRef navNotes As Collection)
    'Review-owned direct/runtime generated-form placement path. Setup D2 never uses this subset path; Setup defers all physical placement to I3.
    On Error GoTo ErrHandler
    Dim formName As Variant
    Dim groupKey As String
    Dim groupName As String
    Dim groupID As Long
    Dim detailText As String
    Dim copiedCount As Long
    If formNames Is Nothing Then Exit Sub
    If formNames.count = 0 Then Exit Sub
    If SettingBool(SETTING_ADD_GENERATED_FORMS_TO_NAVIGATION_PANE_GROUP, DEFAULT_ADD_GENERATED_FORMS_TO_NAVIGATION_PANE_GROUP) = False Then navNotes.Add "Setting [" & SETTING_ADD_GENERATED_FORMS_TO_NAVIGATION_PANE_GROUP & "] is False, but this explicit generated-form build requested grouping and Review will still verify every authoritative form-specific group."
    For Each formName In formNames
        PumpSetupResponsiveness
        groupKey = Trim$(ManagedObjectNavigationPaneGroupKey(CStr(formName), ADCAT_OBJECT_TYPE_FORM))
        If Len(groupKey) = 0 Then Err.Raise vbObjectError + 6680, "ApplyGeneratedFormsNavigationPane", "No authoritative generated-form definition exists." & vbCrLf & "Form: " & CStr(formName)
        groupName = ReviewGeneratedFormNavigationPaneGroupNameByKey(groupKey)
        If Len(groupName) = 0 Then Err.Raise vbObjectError + 6681, "ApplyGeneratedFormsNavigationPane", "The configured Navigation Pane group name is blank." & vbCrLf & "Form: " & CStr(formName) & vbCrLf & "Group key: " & groupKey
        groupID = SharedNavPane_ExistingCustomGroupID(currentDb, groupName)
        If groupID <= 0 Then Err.Raise vbObjectError + 6682, "ApplyGeneratedFormsNavigationPane", "The configured Navigation Pane group was not found." & vbCrLf & "Form: " & CStr(formName) & vbCrLf & "Group key: " & groupKey & vbCrLf & "Group name: " & groupName
        detailText = vbNullString
        If SharedNavPane_AddAccessObjectToExistingCustomGroup(currentDb, CStr(formName), acForm, groupID, detailText, False, True) = False Then Err.Raise vbObjectError + 6683, "ApplyGeneratedFormsNavigationPane", "Could not copy or verify a generated form in its configured Navigation Pane group." & vbCrLf & "Form: " & CStr(formName) & vbCrLf & "Group key: " & groupKey & vbCrLf & "Group: " & groupName & vbCrLf & "Detail: " & detailText
        copiedCount = copiedCount + 1
    Next formName
    If copiedCount <> formNames.count Then Err.Raise vbObjectError + 6684, "ApplyGeneratedFormsNavigationPane", "Not every generated form was verified in its authoritative Navigation Pane group." & vbCrLf & "Expected: " & CStr(formNames.count) & vbCrLf & "Verified: " & CStr(copiedCount)
    navNotes.Add "Review copied or verified " & CStr(copiedCount) & " generated form shortcuts in their authoritative form-specific Navigation Pane groups."
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    navNotes.Add "Review-owned direct/runtime Navigation Pane grouping failed. Error " & CStr(savedErrNumber) & ": " & savedErrDescription
    Err.Raise savedErrNumber, "ApplyGeneratedFormsNavigationPane", "Could not add generated forms to their configured Navigation Pane groups." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Function ReviewSameRunI1GroupProofCurrent(ByRef detailText As String) As Boolean
    On Error GoTo NotCurrent
    Dim activeTraceRunID As String
    Dim physicalSnapshot As Object
    Dim proofItem As Variant
    Dim groupName As String
    Dim expectedGroupID As Long
    Dim actualGroupID As Long
    Dim groupCount As Long
    detailText = vbNullString
    If mI1NavigationPaneProofRows Is Nothing Then detailText = "I1 same-run configured-group proof is unavailable.": Exit Function
    activeTraceRunID = GetActiveTraceRunID()
    If Len(activeTraceRunID) = 0 Or StrComp(activeTraceRunID, mI1NavigationPaneProofTraceRunID, vbBinaryCompare) <> 0 Then detailText = "I1 configured-group proof belongs to a different or inactive setup trace run.": Exit Function
    If StrComp(currentDb.Name, mI1NavigationPaneProofHostDbName, vbTextCompare) <> 0 Then detailText = "I1 configured-group proof belongs to a different HostDb.": Exit Function
    Set physicalSnapshot = ReviewPhysicalNavigationPaneGroupSnapshot()
    For Each proofItem In mI1NavigationPaneProofRows
        groupName = CStr(proofItem(0))
        expectedGroupID = CLng(proofItem(1))
        actualGroupID = 0
        groupCount = ReviewPhysicalNavigationPaneGroupFromSnapshot(physicalSnapshot, groupName, actualGroupID)
        If groupCount <> 1 Or actualGroupID <> expectedGroupID Then detailText = "I1 compact drift gate detected a changed configured group: [" & groupName & "].": Exit Function
    Next proofItem
    mI1NavigationPaneProofReuseCount = mI1NavigationPaneProofReuseCount + 1
    detailText = "I1 same-run exact configured-group proof reused after one fresh physical-group snapshot. Groups=" & CStr(mI1NavigationPaneProofRows.count) & "; reuses=" & CStr(mI1NavigationPaneProofReuseCount) & "."
    ReviewSameRunI1GroupProofCurrent = True
    Exit Function
NotCurrent:
    If IsCancellationError(Err.Number) Then Err.Raise Err.Number, Err.Source, Err.Description
    If Len(detailText) = 0 Then detailText = "I1 compact configured-group proof could not establish current state. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    Err.Clear
End Function

Private Function ReviewPhysicalObjectIdentityExistsExact(ByVal hostDb As DAO.Database, ByVal objectID As Long, ByVal objectName As String, ByVal objectType As AcObjectType, ByRef detailText As String) As Boolean
    'Mutation-boundary proof: the planned Access object must still own the exact MSysObjects.ID before a stale Navigation Pane mirror row can be repaired or remapped.
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim exactCount As Long
    Dim typePredicate As String
    detailText = vbNullString
    If objectID = 0 Or Len(objectName) = 0 Then Exit Function
    typePredicate = MSysObjectTypePredicate(objectType, "[Type]")
    Set rs = hostDb.OpenRecordset("SELECT Count(*) " & "AS " & SQL_FIELD_CNT & " FROM [" & MSYS_OBJECTS_TABLE & "] WHERE [ID]=" & CStr(objectID) & " AND UCase(Nz([Name],''))=" & SharedDao_SqlQuote(UCase$(objectName)) & " AND " & typePredicate & ";", dbOpenSnapshot)
    exactCount = CLng(rs.fields(FIELD_CNT).value)
    rs.Close
    Set rs = Nothing
    If exactCount > 1 Then Err.Raise vbObjectError + 8861, "ReviewPhysicalObjectIdentityExistsExact", "MSysObjects contains duplicate exact physical identities for [" & objectName & "] with ObjectID=" & CStr(objectID) & "."
    ReviewPhysicalObjectIdentityExistsExact = (exactCount = 1)
    Exit Function
ErrHandler:
    detailText = "Could not verify exact MSysObjects identity for [" & objectName & "] with ObjectID=" & CStr(objectID) & ". Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
End Function

Private Sub ReviewAddSystemMacroNavigationPaneGroupCacheEntry(ByVal cache As Object, ByVal groupDefinitions As Object, ByVal MacroTitle As String, ByVal groupKey As String)
    Dim groupData As Variant
    Dim existingData As Variant
    groupKey = Trim$(groupKey)
    If Len(MacroTitle) = 0 Then Err.Raise vbObjectError + 8848, "ReviewAddSystemMacroNavigationPaneGroupCacheEntry", "An authoritative System Macro and Help Document Macro row has a blank physical macro title."
    If Len(groupKey) = 0 Then Err.Raise vbObjectError + 8849, "ReviewAddSystemMacroNavigationPaneGroupCacheEntry", "System Macro and Help Document Macro [" & MacroTitle & "] has no authoritative Navigation Pane group key."
    If groupDefinitions.Exists(groupKey) Then
        groupData = groupDefinitions(groupKey)
    Else
        groupData = Array(0&, vbNullString, 0&, 0&)
    End If
    If cache.Exists(MacroTitle) Then
        existingData = cache(MacroTitle)
        cache(MacroTitle) = Array(CLng(existingData(0)) + 1, groupKey, CLng(groupData(0)), CStr(groupData(1)), CLng(groupData(2)), CLng(groupData(3)))
    Else
        cache.Add MacroTitle, Array(1&, groupKey, CLng(groupData(0)), CStr(groupData(1)), CLng(groupData(2)), CLng(groupData(3)))
    End If
End Sub

Private Function ReviewSystemMacroNavigationPaneGroupCache(ByVal groupDefinitions As Object) As Object
    'I2/I3 read-only cache: physical System Macro and Help Document Macro title -> Array(mappingCount, authoritative group key, groupDefinitionCount, groupName, physicalCount, physicalGroupID).
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim cache As Object
    Dim catalogRows As Collection
    Dim rowData As Variant
    Dim groupKey As String
    Dim MacroTitle As String
    Dim helpTableName As String
    If groupDefinitions Is Nothing Then Err.Raise vbObjectError + 8871, "ReviewSystemMacroNavigationPaneGroupCache", "System Macro and Help Document Macro Navigation Pane mapping requires the authoritative group-key snapshot."
    ' INTENTIONAL KEYED LOOKUP: cache provides keyed membership/lookup over source rows or state already read by ReviewSystemMacroNavigationPaneGroupCache; it is a derived access path, not a second authority.
    Set cache = CreateObject("Scripting.Dictionary")
    cache.CompareMode = vbTextCompare
    Set catalogRows = SystemMacroCatalog()
    For Each rowData In catalogRows
        MacroTitle = SystemMacroCatalogTitle(rowData)
        groupKey = SystemMacroCatalogGroupKey(rowData)
        ReviewAddSystemMacroNavigationPaneGroupCacheEntry cache, groupDefinitions, MacroTitle, groupKey
    Next rowData
    helpTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HelpDocuments")
    If CodeDbTableExists(helpTableName) Then
        If TableCanOpen(helpTableName, True) Then
            Set db = ADCDb()
            Set rs = db.OpenRecordset("SELECT [" & FIELD_HELP_DOCUMENT_NUMBER & "],[" & FIELD_DEFAULT_TITLE_TEXT & "],[" & FIELD_MACRO_TITLE_OVERRIDE & "] FROM [" & helpTableName & "] WHERE Nz(" & SQL_FIELD_HELP_DOCUMENT_STATUS & "," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & " ORDER BY " & OrderByPolicyText("CATALOG_HELP_DOCUMENT") & ";", dbOpenSnapshot)
            Do While rs.EOF = False
                MacroTitle = HelpDocumentActualTitleFromValues(CLng(Nz(rs.fields(FIELD_HELP_DOCUMENT_NUMBER).value, 0)), CStr(Nz(rs.fields(FIELD_DEFAULT_TITLE_TEXT).value, vbNullString)), CStr(Nz(rs.fields(FIELD_MACRO_TITLE_OVERRIDE).value, vbNullString)))
                ReviewAddSystemMacroNavigationPaneGroupCacheEntry cache, groupDefinitions, MacroTitle, SYSTEM_MACRO_GROUP_SETUP_MAINTENANCE_CLEANUP_HELP
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing
        End If
    End If
    Set ReviewSystemMacroNavigationPaneGroupCache = cache
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
    Err.Raise savedErrNumber, "ReviewSystemMacroNavigationPaneGroupCache", "I2/I3 could not preload authoritative System Macro and Help Document Macro Navigation Pane grouping." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

' Purpose: Performs the public preflight managed objects navigation pane for setup operation using the routine's existing ADC workflow and side effects.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub PreflightManagedObjectsNavigationPaneForSetup()
    'I2 owner: builds the complete managed-object placement plan read-only for the final A-G managed-object set, including managed reports and Ribbon-owned objects. Missing or stale MSysNavPaneObjectIDs mirror
    'identities are I3 reconciliation work, not I2 failures.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim hostDb As DAO.Database
    Dim rs As DAO.Recordset
    Dim objectIdCache As Object
    Dim navObjectIdCache As Object
    Dim membershipCache As Object
    Dim physicalGroupSnapshot As Object
    Dim groupCache As Object
    Dim groupKeyCache As Object
    Dim systemMacroGroupCache As Object
    Dim objectName As String
    Dim objectTypeText As String
    Dim SettingKey As String
    Dim groupName As String
    Dim objectType As AcObjectType
    Dim groupID As Long
    Dim objectID As Long
    Dim objectKey As String
    Dim rowOptionalMissing As Boolean
    Dim rowFailureDetail As String
    Dim membershipKey As String
    Dim alreadyPresent As Boolean
    Dim navIdentityPresent As Boolean
    Dim navIdentityMappedID As Long
    Dim failureCount As Long
    Dim failureText As String
    Dim configuredCurrent As Boolean
    Dim h1PhysicalReuseDetail As String
    Dim skippedCount As Long
    Dim identityMaterializationCount As Long
    Dim identityRemapCount As Long
    Dim sqlText As String
    InvalidateReviewCompleteNavigationPaneProof
    ResetSetupNavigationPanePlan
    Set mSetupNavigationPanePlan = New Collection
    If SettingBool(SETTING_ADD_TO_NAV_GROUP, True) = False Then
        mSetupNavigationPanePlanToken = ReviewNavigationPanePlanToken(mSetupNavigationPanePlan)
        mSetupNavigationPanePlanReady = True
        CoreAddOperationNote "I2 Navigation Pane placement preflight is a verified no-op because grouping is disabled."
        Exit Sub
    End If
    If ReviewTryGetSameRunI1PhysicalGroupSnapshot(physicalGroupSnapshot, h1PhysicalReuseDetail) Then
        AddTraceNote h1PhysicalReuseDetail
    Else
        Set physicalGroupSnapshot = ReviewPhysicalNavigationPaneGroupSnapshot()
    End If
    ReviewNavigationPaneGroupDefinitionCaches physicalGroupSnapshot, groupCache, groupKeyCache, configuredCurrent, failureText
    If configuredCurrent = False Then Err.Raise vbObjectError + 8821, "PreflightManagedObjectsNavigationPaneForSetup", "I2 requires I1's configured-group postcondition." & vbCrLf & failureText
    Set db = ADCDb()
    Set hostDb = currentDb
    Set objectIdCache = ReviewNavigationPaneObjectIDCache(hostDb)
    Set navObjectIdCache = ReviewNavigationPaneObjectMappingCache(hostDb)
    Set membershipCache = ReviewNavigationPaneMembershipCache(hostDb)
    Set systemMacroGroupCache = ReviewSystemMacroNavigationPaneGroupCache(groupKeyCache)
    sqlText = "SELECT " & SQL_FIELD_OBJECT_NAME & "," & SQL_FIELD_RESOLVED_OBJECT_TYPE & "," & SQL_FIELD_CATEGORY_CODE & "," & SQL_FIELD_RESOLVED_EXPECTED_EXISTENCE_POLICY & "," & SQL_FIELD_RESOLVED_SETUP_ACTION & "," & SQL_FIELD_RESOLVED_NAVIGATION_PANE_GROUP_SETTING_KEY & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_ManagedObjectsResolved") & "] " & _
              "WHERE Nz(" & SQL_FIELD_OBJECT_CATEGORY_STATUS & "," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & " AND Nz(" & SQL_FIELD_MANAGED_OBJECT_STATUS & "," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & " " & _
              "AND UCase(Trim(Nz(" & SQL_FIELD_RESOLVED_NAV_PANE_VISIBILITY_POLICY & ",'')))=" & SharedDao_SqlQuote(UCase$(ADCAT_NAV_POLICY_ADD_TO_CONFIGURED_GROUP)) & " " & _
              "ORDER BY " & OrderByPolicyText("REVIEW_PLACEMENT") & ";"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    Do While rs.EOF = False
        PumpSetupResponsiveness
        If TryResolveNavigationPanePlacementRow(rs, groupCache, systemMacroGroupCache, objectIdCache, True, objectName, objectTypeText, SettingKey, groupName, objectType, groupID, objectID, objectKey, rowOptionalMissing, rowFailureDetail) Then
            navIdentityMappedID = 0
            navIdentityPresent = False
            If navObjectIdCache.Exists(objectKey) Then
                navIdentityMappedID = CLng(navObjectIdCache(objectKey))
                navIdentityPresent = (navIdentityMappedID = objectID)
                If navIdentityPresent = False Then identityRemapCount = identityRemapCount + 1
            Else
                identityMaterializationCount = identityMaterializationCount + 1
            End If
            membershipKey = CStr(groupID) & "|" & CStr(objectID)
            alreadyPresent = membershipCache.Exists(membershipKey)
            mSetupNavigationPanePlan.Add Array(objectName, CLng(objectType), objectTypeText, groupName, groupID, objectID, alreadyPresent, SettingKey, navIdentityPresent, navIdentityMappedID)
            mSetupNavigationPanePlanCheckedCount = mSetupNavigationPanePlanCheckedCount + 1
            If alreadyPresent Then mSetupNavigationPanePlanAlreadyPresentCount = mSetupNavigationPanePlanAlreadyPresentCount + 1
        ElseIf rowOptionalMissing Then
            skippedCount = skippedCount + 1
            mSetupNavigationPanePlanOptionalMissingCount = mSetupNavigationPanePlanOptionalMissingCount + 1
        Else
            failureCount = failureCount + 1
            failureText = SharedText_AppendNonBlankLine(failureText, rowFailureDetail)
        End If
        RequireSafeWorkContinue "PreflightManagedObjectsNavigationPaneForSetup", objectTypeText & " [" & objectName & "]"
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    ReviewAppendDynamicBackupNavigationPanePlanRows hostDb, objectIdCache, navObjectIdCache, membershipCache, physicalGroupSnapshot, False, mSetupNavigationPanePlan, mSetupNavigationPanePlanCheckedCount, mSetupNavigationPanePlanAlreadyPresentCount, identityMaterializationCount, identityRemapCount, failureCount, failureText
    If failureCount > 0 Then Err.Raise vbObjectError + 8822, "PreflightManagedObjectsNavigationPaneForSetup", "I2 found " & CStr(failureCount) & " blocking placement problem(s)." & vbCrLf & failureText
    mSetupNavigationPanePlanToken = ReviewNavigationPanePlanToken(mSetupNavigationPanePlan)
    mSetupNavigationPanePlanReady = True
    LogTraceMessage TRACE_LEVEL_INFO, "NavigationPane", "PreflightManagedObjectsNavigationPaneForSetup", "I2 complete managed-object Navigation Pane placement preflight succeeded.", "Plan rows=" & CStr(mSetupNavigationPanePlanCheckedCount) & "; already present=" & CStr(mSetupNavigationPanePlanAlreadyPresentCount) & "; Navigation Pane identities needing I3 materialization=" & CStr(identityMaterializationCount) & "; stale logical identities needing I3 remap=" & CStr(identityRemapCount) & "; optional missing=" & CStr(mSetupNavigationPanePlanOptionalMissingCount) & "; skipped=" & CStr(skippedCount) & "; token=" & mSetupNavigationPanePlanToken, "I2 read-only placement preflight", 0, vbNullString, vbNullString, vbNullString, "NavigationPanePlan", "CodeDb", "I2"
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
    ResetSetupNavigationPanePlan
    Err.Raise savedErrNumber, "PreflightManagedObjectsNavigationPaneForSetup", "I2 complete managed-object Navigation Pane placement preflight failed." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

' Purpose: Applies preflighted managed objects navigation pane for setup.
' Public role: Public module entry point. This procedure is listed in the module header as an intentional externally callable ADC entry point and may be invoked directly by Access/VBA callers.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: False.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub ApplyPreflightedManagedObjectsNavigationPaneForSetup(Optional ByVal showMessage As Boolean = False)
    RequireSetupDomainMutationAllowed SETUP_DOMAIN_I3_NAVIGATION_PLACEMENT, "ApplyPreflightedManagedObjectsNavigationPaneForSetup"
    'I3 owner: remaps stale mirror identities or materializes missing preflighted identities under rollback, applies missing memberships, refreshes once, then verifies the settled plan.
    On Error GoTo ErrHandler
    Dim hostDb As DAO.Database
    Dim membershipVerificationCache As Object
    Dim identityVerificationCache As Object
    Dim planItem As Variant
    Dim objectName As String
    Dim objectType As AcObjectType
    Dim objectKind As String
    Dim groupName As String
    Dim groupID As Long
    Dim objectID As Long
    Dim objectKey As String
    Dim membershipKey As String
    Dim alreadyPresentAtI2 As Boolean
    Dim navIdentityPresentAtI2 As Boolean
    Dim createdByThisCall As Boolean
    Dim repairedByThisCall As Boolean
    Dim detailText As String
    Dim staleDetail As String
    Dim preMutationNavObjectCache As Object
    Dim addedCount As Long
    Dim identityAddedCount As Long
    Dim identityRepairedCount As Long
    Dim identityRemappedCount As Long
    Dim staleIdentityMembershipRemovedCount As Long
    Dim backupOtherMembershipRemovedCount As Long
    Dim identityConcurrentCount As Long
    Dim unchangedCount As Long
    Dim concurrentCount As Long
    Dim verifiedCount As Long
    Dim identityVerifiedCount As Long
    Dim backupExclusiveVerifiedCount As Long
    Dim membershipCacheKey As Variant
    Dim membershipData As Variant
    Dim membershipObjectName As String
    Dim failureText As String
    If mSetupNavigationPanePlanReady = False Or mSetupNavigationPanePlan Is Nothing Or Len(mSetupNavigationPanePlanToken) = 0 Then Err.Raise vbObjectError + 8830, "ApplyPreflightedManagedObjectsNavigationPaneForSetup", "I3 requires a successful same-session I2 plan."
    If ReviewNavigationPanePlanStillCurrent(staleDetail, preMutationNavObjectCache) = False Then Err.Raise vbObjectError + 8831, "ApplyPreflightedManagedObjectsNavigationPaneForSetup", "I3 refused to mutate because the I2 physical identity plan is stale." & vbCrLf & staleDetail
    Set hostDb = currentDb
    detailText = vbNullString
    If ReviewRemovePreflightedStaleNavigationPaneObjectIdentities(hostDb, mSetupNavigationPanePlan, preMutationNavObjectCache, identityRemappedCount, staleIdentityMembershipRemovedCount, detailText) = False Then Err.Raise vbObjectError + 8887, "ApplyPreflightedManagedObjectsNavigationPaneForSetup", detailText
    ReviewTrackAndRemovePreflightedBackupNavigationPaneMemberships hostDb, mSetupNavigationPanePlan, backupOtherMembershipRemovedCount
    For Each planItem In mSetupNavigationPanePlan
        PumpSetupResponsiveness
        objectName = CStr(planItem(0))
        objectType = CLng(planItem(1))
        objectKind = CStr(planItem(2))
        groupName = CStr(planItem(3))
        groupID = CLng(planItem(4))
        objectID = CLng(planItem(5))
        alreadyPresentAtI2 = CBool(planItem(6))
        navIdentityPresentAtI2 = CBool(planItem(8))
        repairedByThisCall = False
        If navIdentityPresentAtI2 = False Then
            detailText = vbNullString
            If ReviewNavigationPaneObjectIdentityExistsExact(hostDb, objectID, objectName, objectType, detailText) Then
                identityConcurrentCount = identityConcurrentCount + 1
            Else
                If Len(detailText) > 0 Then Err.Raise vbObjectError + 8856, "ApplyPreflightedManagedObjectsNavigationPaneForSetup", detailText
                createdByThisCall = False
                repairedByThisCall = False
                If ReviewInsertPreflightedNavigationPaneObjectIdentity(hostDb, objectID, objectName, objectType, createdByThisCall, repairedByThisCall, detailText) = False Then Err.Raise vbObjectError + 8857, "ApplyPreflightedManagedObjectsNavigationPaneForSetup", "I3 could not materialize a preflighted Navigation Pane object identity." & vbCrLf & "Object: " & objectKind & " [" & objectName & "]" & vbCrLf & detailText
                If createdByThisCall Then
                    ReviewTrackNavigationPaneObjectIdentityRollbackOrUndo hostDb, objectID, objectName, objectType
                    identityAddedCount = identityAddedCount + 1
                ElseIf repairedByThisCall Then
                    identityRepairedCount = identityRepairedCount + 1
                Else
                    identityConcurrentCount = identityConcurrentCount + 1
                End If
            End If
        End If
        If alreadyPresentAtI2 And repairedByThisCall = False Then
            unchangedCount = unchangedCount + 1
        ElseIf ReviewNavigationPaneMembershipExists(hostDb, groupID, objectID) Then
            concurrentCount = concurrentCount + 1
        Else
            createdByThisCall = False
            detailText = vbNullString
            If ReviewEnsureNavigationPaneMembership(hostDb, groupID, objectID, objectName, objectType, createdByThisCall, detailText) = False Then Err.Raise vbObjectError + 8832, "ApplyPreflightedManagedObjectsNavigationPaneForSetup", "I3 could not insert a preflighted Navigation Pane membership." & vbCrLf & "Object: " & objectKind & " [" & objectName & "]" & vbCrLf & "Group: " & groupName & vbCrLf & detailText
            If createdByThisCall Then
                ReviewTrackNavigationPaneShortcutRollbackOrUndo hostDb, groupID, objectID
                addedCount = addedCount + 1
            Else
                concurrentCount = concurrentCount + 1
            End If
        End If
        RequireSafeWorkContinue "ApplyPreflightedManagedObjectsNavigationPaneForSetup", objectKind & " [" & objectName & "]"
    Next planItem
    If identityAddedCount > 0 Or identityRepairedCount > 0 Or identityRemappedCount > 0 Or backupOtherMembershipRemovedCount > 0 Or addedCount > 0 Then ReviewRefreshNavigationPaneOnce
    Set membershipVerificationCache = ReviewNavigationPaneMembershipCache(hostDb)
    Set identityVerificationCache = ReviewNavigationPaneObjectMappingCache(hostDb)
    For Each planItem In mSetupNavigationPanePlan
        objectName = CStr(planItem(0))
        objectType = CLng(planItem(1))
        groupName = CStr(planItem(3))
        groupID = CLng(planItem(4))
        objectID = CLng(planItem(5))
        objectKey = CStr(CLng(objectType)) & "|" & UCase$(objectName)
        membershipKey = CStr(groupID) & "|" & CStr(objectID)
        If identityVerificationCache.Exists(objectKey) Then
            If CLng(identityVerificationCache(objectKey)) = objectID Then
                identityVerifiedCount = identityVerifiedCount + 1
            Else
                failureText = SharedText_AppendNonBlankLine(failureText, "Mismatched Navigation Pane identity after I3: Object [" & objectName & "]; ObjectID=" & CStr(objectID) & ".")
            End If
        Else
            failureText = SharedText_AppendNonBlankLine(failureText, "Missing Navigation Pane identity after I3: Object [" & objectName & "]; ObjectID=" & CStr(objectID) & ".")
        End If
        If membershipVerificationCache.Exists(membershipKey) Then
            verifiedCount = verifiedCount + 1
        Else
            failureText = SharedText_AppendNonBlankLine(failureText, "Missing after I3: Object [" & objectName & "]; Group [" & groupName & "]; GroupID=" & CStr(groupID) & "; ObjectID=" & CStr(objectID) & ".")
        End If
        If CLng(objectType) = CLng(acTable) And IsBackupAccessObjectName(objectName) Then
            For Each membershipCacheKey In membershipVerificationCache.keys
                membershipData = membershipVerificationCache(membershipCacheKey)
                membershipObjectName = Trim$(CStr(Nz(membershipData(2), vbNullString)))
                If CLng(membershipData(1)) = objectID Or (Len(membershipObjectName) > 0 And StrComp(membershipObjectName, objectName, vbTextCompare) = 0) Then
                    If CLng(membershipData(0)) <> groupID Or CLng(membershipData(1)) <> objectID Then failureText = SharedText_AppendNonBlankLine(failureText, "Backup placement is not exclusive after I3: Table [" & objectName & "] retains an extra shortcut at GroupID=" & CStr(CLng(membershipData(0))) & ", ObjectID=" & CStr(CLng(membershipData(1))) & ".")
                End If
            Next membershipCacheKey
            backupExclusiveVerifiedCount = backupExclusiveVerifiedCount + 1
        End If
    Next planItem
    If Len(failureText) > 0 Then Err.Raise vbObjectError + 8833, "ApplyPreflightedManagedObjectsNavigationPaneForSetup", "I3 final batch verification found missing or mismatched required Navigation Pane state." & vbCrLf & failureText
    Call EnsureDaoPropertyValue(currentDb, NAVIGATION_PANE_MEMBERSHIP_TOKEN_PROPERTY, dbText, mSetupNavigationPanePlanToken, "ApplyPreflightedManagedObjectsNavigationPaneForSetup")
    PublishReviewCompleteNavigationPaneProof mSetupNavigationPanePlan, mSetupNavigationPanePlanToken
    LogTraceMessage TRACE_LEVEL_INFO, "NavigationPane", "ApplyPreflightedManagedObjectsNavigationPaneForSetup", "I3 complete managed-object Navigation Pane placement applied and verified.", _
        "Plan rows=" & CStr(mSetupNavigationPanePlanCheckedCount) & "; identities materialized=" & CStr(identityAddedCount) & "; identities repaired in place=" & CStr(identityRepairedCount) & "; stale logical identities remapped=" & CStr(identityRemappedCount) & "; stale memberships removed=" & CStr(staleIdentityMembershipRemovedCount) & "; non-Backup backup shortcuts removed=" & CStr(backupOtherMembershipRemovedCount) & _
        "; identities concurrently present=" & CStr(identityConcurrentCount) & "; identities verified=" & CStr(identityVerifiedCount) & "; memberships unchanged=" & CStr(unchangedCount) & "; memberships concurrently present=" & CStr(concurrentCount) & "; memberships added=" & CStr(addedCount) & "; memberships verified=" & CStr(verifiedCount) & "; backup exclusivity verified=" & CStr(backupExclusiveVerifiedCount) & "; optional missing=" & CStr(mSetupNavigationPanePlanOptionalMissingCount) & "; token=" & mSetupNavigationPanePlanToken, _
        "I3 apply and verify", 0, vbNullString, vbNullString, vbNullString, "NavigationPanePlan", "CodeDb", "I3"
    If showMessage Then CoreShowMessage "I3 Navigation Pane placement applied and verified." & vbCrLf & "Identities materialized: " & CStr(identityAddedCount) & vbCrLf & "Identities repaired in place: " & CStr(identityRepairedCount) & vbCrLf & "Stale logical identities remapped: " & CStr(identityRemappedCount) & vbCrLf & "Memberships added: " & CStr(addedCount), vbInformation, DEFAULT_UTILITY_LONG_NAME
    ResetSetupNavigationPanePlan
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    ResetSetupNavigationPanePlan
    If showMessage Then
        CoreShowErrorMessage "I3 complete managed-object Navigation Pane placement failed." & vbCrLf & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, vbCritical, "Managed Object Navigation Pane Placement", "ApplyPreflightedManagedObjectsNavigationPaneForSetup", "Review"
        Exit Sub
    End If
    Err.Raise savedErrNumber, "ApplyPreflightedManagedObjectsNavigationPaneForSetup", "I3 complete managed-object Navigation Pane placement failed." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Function NzBool(ByVal value As Variant, ByVal DefaultValue As Boolean) As Boolean
    If IsNull(value) Then
        NzBool = DefaultValue
    Else
        NzBool = CBool(value)
    End If
End Function

' Purpose: Resolves review DB engine errors text from the supplied/current ADC state, using `SharedText_AppendNonBlankLine` and `Next` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function ReviewDBEngineErrorsText() As String
    On Error Resume Next
    Dim daoErrorItem As DAO.Error
    Dim detailsText As String
    For Each daoErrorItem In DBEngine.Errors
        detailsText = SharedText_AppendNonBlankLine(detailsText, "DAO error " & CStr(daoErrorItem.Number) & " | Source: " & daoErrorItem.Source & " | Description: " & daoErrorItem.Description)
    Next daoErrorItem
    On Error GoTo 0
    ReviewDBEngineErrorsText = detailsText
End Function

Private Sub RequireManagedObjectCatalogReadyForDescriptionWork(ByVal sourceProcedure As String)
    'A-J Setup description work consumes the complete logical catalog established by D4; it must not invoke an extra catalog bootstrap/update path.
    If ManagedObjectCatalogReadyForDescriptionWork() Then Exit Sub
    Err.Raise vbObjectError + 6464, sourceProcedure, _
        "The active Managed Object Catalog is not ready for description work." & vbCrLf & _
        "Active A-J Setup requires the catalog structure, required rows, support query, and prior catalog phases to be established before descriptions are applied." & vbCrLf & _
        "No extra catalog bootstrap/update path was invoked."
End Sub

Private Function RepairManagedObjectCatalogDescriptions(Optional ByVal showMessage As Boolean = False, Optional ByVal includeManagedForms As Boolean = True, Optional ByVal formsOnly As Boolean = False, Optional ByVal helpDocumentMacrosOnly As Boolean = False) As Long
    If SetupDomainMutationWritesAllowedForCurrentPart(SETUP_DOMAIN_G1_DESCRIPTIONS) = False Then Err.Raise vbObjectError + 10021, "RepairManagedObjectCatalogDescriptions", "Managed Object Catalog ObjectDescription mutation is owned by G1 during numbered Setup."
    'Repairs active managed-object descriptions through G1. Help-document macro descriptions are synchronized exactly from their B4 payload derived from Help_Docs.ObjectDescription; blank remains valid and is never
    'synthesized.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim objectName As String
    Dim objectTypeCode As String
    Dim categoryCode As String
    Dim descriptionText As String
    Dim repairedCount As Long
    Dim descriptionOptionalForHelpMacro As Boolean
    Dim authoritativeHelpDescription As String
    Dim checkedCount As Long
    BeginManagedDescriptionOptionalHelpPolicyScope
    RequireManagedObjectCatalogReadyForDescriptionWork "RepairManagedObjectCatalogDescriptions"
    Set db = ADCDb()
    sqlText = "SELECT m." & SQL_FIELD_ID & ",m." & SQL_FIELD_OBJECT_NAME & ",m." & SQL_FIELD_OBJECT_TYPE & ",m.[" & FIELD_CATEGORY_CODE & "],m." & SQL_FIELD_OBJECT_DESCRIPTION & " FROM ([" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID) & "] AS m INNER JOIN [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ObjectCategories") & "] AS c ON m.[" & FIELD_CATEGORY_CODE & "]=c.[" & FIELD_CATEGORY_CODE & "]) LEFT JOIN [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECT_OVERRIDES_TABLE_LOGICAL_ID) & "] AS o ON m." & SQL_FIELD_ID & "=o.[" & FIELD_MANAGED_OBJECT_ID & "] WHERE Nz(m.[" & FIELD_MANAGED_OBJECT_STATUS & "]," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & " AND Nz(c." & SQL_FIELD_OBJECT_CATEGORY_STATUS & "," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE)
    If formsOnly Then
        sqlText = sqlText & " AND UCase(Trim(Nz(m." & SQL_FIELD_OBJECT_TYPE & ",'')))=" & SharedDao_SqlQuote(UCase$(ADCAT_OBJECT_TYPE_FORM))
    ElseIf includeManagedForms = False Then
        sqlText = sqlText & " AND UCase(Trim(Nz(m." & SQL_FIELD_OBJECT_TYPE & ",'')))<>" & SharedDao_SqlQuote(UCase$(ADCAT_OBJECT_TYPE_FORM))
    End If
    sqlText = sqlText & " ORDER BY " & OrderByPolicyText("REVIEW_MANAGED_DESCRIPTION") & ";"
    Set rs = db.OpenRecordset(sqlText, dbOpenDynaset)
    Do While rs.EOF = False
        objectName = CStr(Nz(rs.fields(FIELD_OBJECT_NAME).value, vbNullString))
        objectTypeCode = CStr(Nz(rs.fields(FIELD_OBJECT_TYPE).value, vbNullString))
        categoryCode = CStr(rs.fields(FIELD_CATEGORY_CODE).value)
        descriptionText = CStr(Nz(rs.fields(FIELD_OBJECT_DESCRIPTION).value, vbNullString))
        authoritativeHelpDescription = vbNullString
        descriptionOptionalForHelpMacro = ManagedDescriptionOptionalHelpMacroDescription(objectName, objectTypeCode, authoritativeHelpDescription)
        If helpDocumentMacrosOnly And descriptionOptionalForHelpMacro = False Then GoTo NextManagedCatalogDescriptionRow
        checkedCount = checkedCount + 1
        If checkedCount = 1 Or checkedCount Mod 25 = 0 Then RequireSafeWorkContinue "RepairManagedObjectCatalogDescriptions", "repairing description " & CStr(checkedCount)
        If descriptionOptionalForHelpMacro Then
            If StrComp(descriptionText, authoritativeHelpDescription, vbBinaryCompare) <> 0 Then
                rs.Edit
                If Len(authoritativeHelpDescription) = 0 Then
                    rs.fields(FIELD_OBJECT_DESCRIPTION).value = Null
                Else
                    rs.fields(FIELD_OBJECT_DESCRIPTION).value = authoritativeHelpDescription
                End If
                rs.Update
                repairedCount = repairedCount + 1
            End If
        ElseIf Len(Trim$(descriptionText)) = 0 Or ManagedObjectDescriptionIsPlaceholder(descriptionText) Then
            descriptionText = ManagedObjectDefaultDescriptionForSetup(objectName, categoryCode, objectTypeCode)
            If Len(Trim$(descriptionText)) = 0 Or ManagedObjectDescriptionIsPlaceholder(descriptionText) Then Err.Raise vbObjectError + 6460, "RepairManagedObjectCatalogDescriptions", "No acceptable authoritative managed-object description could be resolved." & vbCrLf & "Object type: " & objectTypeCode & vbCrLf & "Object name: " & objectName & vbCrLf & "Category: " & categoryCode
            rs.Edit
            rs.fields(FIELD_OBJECT_DESCRIPTION).value = descriptionText
            rs.Update
            repairedCount = repairedCount + 1
        End If
NextManagedCatalogDescriptionRow:
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    RepairManagedObjectCatalogDescriptions = repairedCount
    EndManagedDescriptionOptionalHelpPolicyScope
    LogTraceMessage TRACE_LEVEL_INFO, "ADC_ManagedObjects", "RepairManagedObjectCatalogDescriptions", "Active managed-object authoritative descriptions are complete; Help-document macro descriptions were synchronized exactly from B4-derived [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HelpDocuments") & "] values and remain optional.", "Descriptions repaired: " & CStr(repairedCount), "Repair Managed Object Catalog descriptions", 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID), "Table", "CodeDb", "ManagedObjectDescriptions"
    If showMessage Then ReportCoreActionSuccess "RepairManagedObjectCatalogDescriptions", ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID), "Managed Object Catalog descriptions verified.", "Descriptions repaired: " & CStr(repairedCount), "Repair Managed Object Catalog descriptions", ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID), "Table", "CodeDb", "ManagedObjectDescriptions", True, DEFAULT_UTILITY_LONG_NAME & " Managed Objects"
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
    EndManagedDescriptionOptionalHelpPolicyScope
    On Error GoTo 0
    If showMessage Then
        CoreShowErrorMessage "Could not repair and verify active Managed Object Catalog descriptions." & vbCrLf & vbCrLf & "Last object type: " & objectTypeCode & vbCrLf & "Last object name: " & objectName & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, vbCritical, "Managed Object Catalog Descriptions", "RepairManagedObjectCatalogDescriptions", "Review"
        Exit Function
    End If
    Err.Raise savedErrNumber, "RepairManagedObjectCatalogDescriptions", "Could not repair and verify active Managed Object Catalog descriptions." & vbCrLf & "Last object type: " & objectTypeCode & vbCrLf & "Last object name: " & objectName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function SetAccessObjectDescriptionByCatalogType(ByVal db As DAO.Database, ByVal objectName As String, ByVal objectTypeCode As String, ByVal descriptionText As String, Optional ByVal allowBlankDescription As Boolean = False) As Long
    'Returns 1 when changed, 0 when already current, and -1 only when the target is genuinely missing or unsupported. Optional blank Help descriptions clear stale physical text but do not require a Description
    'property to exist.
    On Error GoTo ErrHandler
    Dim doc As DAO.Document
    objectTypeCode = Trim$(objectTypeCode)
    SetAccessObjectDescriptionByCatalogType = -1
    If Len(objectName) = 0 Then Exit Function
    If Len(descriptionText) = 0 And allowBlankDescription = False Then Exit Function
    Select Case UCase$(objectTypeCode)
        Case UCase$(ADCAT_OBJECT_TYPE_MACRO)
            If DatabaseAccessObjectExistsByName(db, acMacro, objectName) = False Then Exit Function
            Set doc = db.Containers(MACRO_DAO_CONTAINER_NAME).Documents(objectName)
            SetAccessObjectDescriptionByCatalogType = IIf(SetAccessObjectDescription(doc, objectName, objectTypeCode, descriptionText, allowBlankDescription), 1, 0)
        Case UCase$(ADCAT_OBJECT_TYPE_TABLE)
            If SharedDao_TableExists(db, objectName) = False Then Exit Function
            SetAccessObjectDescriptionByCatalogType = IIf(SetAccessObjectDescription(db.TableDefs(objectName), objectName, objectTypeCode, descriptionText, allowBlankDescription), 1, 0)
        Case UCase$(ADCAT_OBJECT_TYPE_QUERY)
            If SharedDao_QueryExists(db, objectName) = False Then Exit Function
            SetAccessObjectDescriptionByCatalogType = IIf(SetAccessObjectDescription(db.QueryDefs(objectName), objectName, objectTypeCode, descriptionText, allowBlankDescription), 1, 0)
        Case UCase$(ADCAT_OBJECT_TYPE_FORM)
            If DatabaseAccessObjectExistsByName(db, acForm, objectName) = False Then Exit Function
            Set doc = db.Containers(FORM_DAO_CONTAINER_NAME).Documents(objectName)
            SetAccessObjectDescriptionByCatalogType = IIf(SetAccessObjectDescription(doc, objectName, objectTypeCode, descriptionText, allowBlankDescription), 1, 0)
        Case UCase$(ADCAT_OBJECT_TYPE_REPORT)
            If DatabaseAccessObjectExistsByName(db, acReport, objectName) = False Then Exit Function
            Set doc = db.Containers(REPORT_DAO_CONTAINER_NAME).Documents(objectName)
            SetAccessObjectDescriptionByCatalogType = IIf(SetAccessObjectDescription(doc, objectName, objectTypeCode, descriptionText, allowBlankDescription), 1, 0)
    End Select
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim daoErrorsText As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    daoErrorsText = ReviewDBEngineErrorsText()
    Err.Raise savedErrNumber, "SetAccessObjectDescriptionByCatalogType", "Managed-object Description write failed." & vbCrLf & "Object type: " & objectTypeCode & vbCrLf & "Object name: " & objectName & vbCrLf & "Property: " & DAO_DESCRIPTION_PROPERTY_NAME & vbCrLf & "Blank authoritative Description allowed: " & CStr(allowBlankDescription) & vbCrLf & "Authoritative characters: " & CStr(Len(descriptionText)) & vbCrLf & "Authoritative bytes: " & CStr(LenB(descriptionText)) & vbCrLf & "Target Db: " & SharedDao_SafeDatabaseName(db) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription & IIf(Len(daoErrorsText) > 0, vbCrLf & "DBEngine.Errors:" & vbCrLf & daoErrorsText, vbNullString)
End Function

' Purpose: Resolves and returns ribbon dependency token for setup from the supplied inputs and the authoritative ADC definitions/state used by this module.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   detailText (Input/output; String): Receives or supplies diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
'   provenObjectDescriptionToken (Output Optional; String): Receives the resolved proven object description token produced from this procedure's validated inputs and current ADC state. Default: vbNullString.
'   provenFieldDescriptionToken (Output Optional; String): Receives the resolved proven field description token produced from this procedure's validated inputs and current ADC state. Default: vbNullString.
'   provenCatalogToken (Output Optional; String): Receives the resolved proven catalog token produced from this procedure's validated inputs and current ADC state. Default: vbNullString.
'   persistedProofOnly (Input Optional; Boolean): Boolean state describing whether persisted proof only applies to this invocation; True selects the named condition/path and False selects its
'   complementary/default path. Default: False.
'   sameRunProvenCatalogToken (Input Optional; String): Proof/version token identifying the exact same run proven catalog state that was previously validated in this setup run. Default: vbNullString.
'   sameRunProvenCatalogTraceRunID (Input Optional; String): Trace-run ID carrying previously proven same-run catalog state into a dependent operation without repeating the proof. Default: vbNullString.
' Returns: String: Returns the deterministic ribbon dependency token for setup token/signature produced from the supplied state, allowing callers to compare exact state without repeating the underlying
'   scan/build work.
Public Function RibbonDependencyTokenForSetup(ByRef detailText As String, Optional ByRef provenObjectDescriptionToken As String = vbNullString, Optional ByRef provenFieldDescriptionToken As String = vbNullString, Optional ByRef provenCatalogToken As String = vbNullString, Optional ByVal persistedProofOnly As Boolean = False, Optional ByVal sameRunProvenCatalogToken As String = vbNullString, Optional ByVal sameRunProvenCatalogTraceRunID As String = vbNullString) As String
    'Read-only H1 predecessor proof: consumes D4 logical catalog state plus the exact G1 managed-description state; a trace-bound already-proven CATALOG token may be reused without recalculation. Setup orchestration
    'proves E3 managed-report physical state separately before H1. Physical Navigation Pane membership is intentionally deferred to Phase I.
    On Error GoTo ErrHandler
    Dim hostDb As DAO.Database
    Dim catalogDetail As String
    Dim objectDescriptionToken As String
    Dim objectDescriptionDetail As String
    Dim fieldDescriptionToken As String
    Dim wholeCatalogToken As String
    Dim storedWholeCatalogToken As String
    Dim propertyExists As Boolean
    Dim payloadText As String
    Dim catalogToken As String
    provenObjectDescriptionToken = vbNullString
    provenFieldDescriptionToken = vbNullString
    provenCatalogToken = vbNullString
    Set hostDb = currentDb
    storedWholeCatalogToken = DatabaseTextPropertyValue(hostDb, MANAGED_DESCRIPTION_WHOLE_CATALOG_EXACT_TOKEN_PROPERTY, propertyExists)
    If propertyExists = False Or Len(storedWholeCatalogToken) = 0 Then Err.Raise vbObjectError + 6473, "RibbonDependencyTokenForSetup", "The persisted G1 whole-catalog managed-description exact token is missing."
    If Len(Trim$(sameRunProvenCatalogToken)) > 0 Then
        If Len(Trim$(sameRunProvenCatalogTraceRunID)) = 0 Then Err.Raise vbObjectError + 6822, "RibbonDependencyTokenForSetup", "A same-run proven complete-catalog token was supplied without its trace identity."
        If TraceRunActive() = False Then Err.Raise vbObjectError + 6823, "RibbonDependencyTokenForSetup", "The supplied same-run complete-catalog token cannot be consumed because no trace run is active."
        If StrComp(Trim$(sameRunProvenCatalogTraceRunID), GetActiveOrLastTraceRunID(), vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 6824, "RibbonDependencyTokenForSetup", "The supplied same-run complete-catalog token does not belong to the active trace run."
        catalogToken = Trim$(sameRunProvenCatalogToken)
        catalogDetail = "Complete-catalog token reused from authoritative same-run shared CATALOG handoff; full current-token recalculation skipped."
    ElseIf Len(Trim$(sameRunProvenCatalogTraceRunID)) > 0 Then
        Err.Raise vbObjectError + 6825, "RibbonDependencyTokenForSetup", "A same-run complete-catalog trace identity was supplied without its token."
    Else
        catalogToken = CompleteCatalogStateTokenForSetup(catalogDetail)
        If Len(catalogToken) = 0 Then Err.Raise vbObjectError + 6470, "RibbonDependencyTokenForSetup", catalogDetail
    End If
    objectDescriptionToken = ManagedObjectDescriptionCurrentTokenForSetup(objectDescriptionDetail, Not persistedProofOnly)
    If Len(objectDescriptionToken) = 0 Then Err.Raise vbObjectError + 6472, "RibbonDependencyTokenForSetup", objectDescriptionDetail
    If persistedProofOnly Then
        If ManagedTableFieldDescriptionsPersistedProofCurrent(fieldDescriptionToken, detailText, False) = False Then Err.Raise vbObjectError + 6472, "RibbonDependencyTokenForSetup", detailText
    Else
        If ManagedTableFieldDescriptionsCurrentForSetup(fieldDescriptionToken, detailText) = False Then Err.Raise vbObjectError + 6472, "RibbonDependencyTokenForSetup", detailText
    End If
    If Len(objectDescriptionToken) = 0 Or Len(fieldDescriptionToken) = 0 Then Err.Raise vbObjectError + 6472, "RibbonDependencyTokenForSetup", "G1 exact managed-description verification did not produce complete object and field tokens."
    wholeCatalogToken = ManagedDescriptionWholeCatalogExactToken(objectDescriptionToken, fieldDescriptionToken)
    If StrComp(storedWholeCatalogToken, wholeCatalogToken, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 6474, "RibbonDependencyTokenForSetup", "The persisted G1 whole-catalog managed-description exact token is stale."
    payloadText = "ADCRibbonDependencyV3|" & FingerprintTokenPart(catalogToken) & FingerprintTokenPart(wholeCatalogToken)
    provenObjectDescriptionToken = objectDescriptionToken
    provenFieldDescriptionToken = fieldDescriptionToken
    provenCatalogToken = catalogToken
    RibbonDependencyTokenForSetup = FingerprintFromPayload(payloadText)
    detailText = "Complete logical catalog=current; D4 managed-form proof is independently current; E3 managed-report physical state is an orchestration prerequisite; G1 managed descriptions=current. Ribbon physical Navigation Pane membership is intentionally deferred to Phase I." & vbCrLf & catalogDetail & vbCrLf & objectDescriptionDetail & vbCrLf & detailText
    Exit Function
ErrHandler:
    If IsCancellationError(Err.Number) Then Err.Raise Err.Number, Err.Source, Err.Description
    provenObjectDescriptionToken = vbNullString
    provenFieldDescriptionToken = vbNullString
    provenCatalogToken = vbNullString
    detailText = "Could not verify the complete D4/E3/G1 predecessor state required by H1. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    RibbonDependencyTokenForSetup = vbNullString
End Function

Private Function NavigationPaneGroupNameBySettingKey(ByVal SettingKey As String) As String
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim sqlText As String
    If Len(SettingKey) = 0 Then Exit Function
    If TableCanOpen(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups")) = False Then Exit Function
    If TableFieldExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups"), FIELD_SETTING_KEY) = False Then Exit Function
    If TableFieldExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups"), FIELD_NAVIGATION_PANE_GROUP_NAME) = False Then Exit Function
    sqlText = "SELECT TOP 1 " & SQL_FIELD_NAVIGATION_PANE_GROUP_NAME & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups") & "] WHERE " & SQL_FIELD_SETTING_KEY & "=" & SharedDao_SqlQuote(SettingKey) & " AND Nz(" & SQL_FIELD_LOOKUP_ACTIVE & ",True)<>False ORDER BY Nz(" & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER & ",9999), " & SQL_FIELD_NAVIGATION_PANE_GROUP_NAME & ";"
    Set rs = ADCDb().OpenRecordset(sqlText, dbOpenSnapshot)
    If rs.EOF = False Then
        NavigationPaneGroupNameBySettingKey = CStr(Nz(rs.fields(FIELD_NAVIGATION_PANE_GROUP_NAME).value, vbNullString))
    End If
ExitHere:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Exit Function
ErrHandler:
    NavigationPaneGroupNameBySettingKey = vbNullString
    Resume ExitHere
End Function

Private Function DefaultNavigationPaneGroupNameBySetting(ByVal SettingKey As String, ByVal DefaultValue As String) As String
    Dim catalogGroupName As String
    catalogGroupName = NavigationPaneGroupNameBySettingKey(SettingKey)
    If Len(catalogGroupName) > 0 Then
        DefaultNavigationPaneGroupNameBySetting = catalogGroupName
    Else
        DefaultNavigationPaneGroupNameBySetting = SettingText(SettingKey, DefaultValue)
    End If
End Function

' Purpose: Adds macro to configured navigation pane group for macro title, configured group name.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   MacroTitle (Input; String): Title text used as the macro title shown, stored, compared, or resolved by this operation.
'   configuredGroupName (Input; String): Configured Ribbon/Navigation group name used to resolve the canonical managed group identity.
'   forceRefreshShortcut (Input Optional; Boolean): When True, enables or requests refresh shortcut for this call; when False, the routine follows its normal non-forced/default path. Default: False.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub AddMacroToConfiguredNavigationPaneGroup(ByVal MacroTitle As String, ByVal configuredGroupName As String, Optional ByVal forceRefreshShortcut As Boolean = False)
    On Error GoTo ErrHandler
    Dim effectiveGroupName As String
    Dim defaultGroupName As String
    If SettingBool(SETTING_ADD_TO_NAV_GROUP, True) = False Then Exit Sub
    defaultGroupName = DefaultNavigationPaneGroupNameBySetting(SETTING_DEFAULT_GENERATED_MACROS_NAV_GROUP, BuildNavigationPaneGroupName(DEFAULT_GENERATED_MACROS_NAVIGATION_PANE_GROUP_LOGICAL_NAME, False))
    effectiveGroupName = EffectiveMacroNavigationPaneGroup(configuredGroupName, defaultGroupName)
    If Len(effectiveGroupName) = 0 Then
        Err.Raise vbObjectError + 6456, "AddMacroToConfiguredNavigationPaneGroup", "The macro was created, but no Navigation Pane group could be resolved while automatic grouping was enabled." & vbCrLf & "MacroTitle: " & MacroTitle & vbCrLf & "Configured group: " & IIf(Len(configuredGroupName) = 0, "(blank)", configuredGroupName) & vbCrLf & "Default group: " & IIf(Len(defaultGroupName) = 0, "(blank)", defaultGroupName)
    End If
    AddGeneratedAccessObjectToNavigationPaneGroup MacroTitle, acMacro, ADCAT_OBJECT_TYPE_MACRO, effectiveGroupName, "AddMacroToConfiguredNavigationPaneGroup", forceRefreshShortcut, False
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    AddNavigationPaneGroupingNote "Macro [" & MacroTitle & "] was created, but Navigation Pane grouping failed." & vbCrLf & "Configured group: " & IIf(Len(configuredGroupName) = 0, "(blank)", configuredGroupName) & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
    On Error GoTo 0
    Err.Raise savedErrNumber, "AddMacroToConfiguredNavigationPaneGroup", x__UTIL_ADC_MACRO_RUNTIME.FormatErrorWithContext("The macro was created, but Navigation Pane grouping failed.", "MacroTitle: " & MacroTitle & vbCrLf & "Configured group: " & IIf(Len(configuredGroupName) = 0, "(blank)", configuredGroupName) & vbCrLf & "Original Err.Source: " & savedErrSource & vbCrLf & "Original Err.Description: " & savedErrDescription)
End Sub

Private Function EffectiveMacroNavigationPaneGroup(ByVal configuredGroupName As String, ByVal defaultGroupName As String) As String
    Dim groupName As String
    Dim failureDetail As String
    groupName = configuredGroupName
    If Len(groupName) > 0 Then
        failureDetail = vbNullString
        If ResolveOrCreateNavigationPaneGroupID(groupName, vbNullString, "EffectiveMacroNavigationPaneGroup", failureDetail) > 0 Then
            EffectiveMacroNavigationPaneGroup = groupName
            Exit Function
        End If
        AddNavigationPaneGroupCreationFailure groupName, "EffectiveMacroNavigationPaneGroup", failureDetail
    End If
    If Len(defaultGroupName) > 0 Then
        failureDetail = vbNullString
        If ResolveOrCreateNavigationPaneGroupID(defaultGroupName, SETTING_DEFAULT_GENERATED_MACROS_NAV_GROUP, "EffectiveMacroNavigationPaneGroup", failureDetail) > 0 Then
            EffectiveMacroNavigationPaneGroup = defaultGroupName
            Exit Function
        End If
        AddNavigationPaneGroupCreationFailure defaultGroupName, "EffectiveMacroNavigationPaneGroup", failureDetail
    End If
End Function

Private Sub AddBackupAccessObjectToNavigationPaneGroup(ByVal objectName As String, ByVal objectType As AcObjectType, ByVal objectKind As String, ByVal sourceName As String, Optional ByVal forceRefreshShortcut As Boolean = True, Optional ByVal showMessage As Boolean = True)
    'Runtime backup placement policy: chooses the backup group and requests REVIEW-owned physical placement/removal.
    On Error GoTo ErrHandler
    Dim groupName As String
    Dim groupID As Long
    Dim objectID As Long
    Dim detailText As String
    Dim messageText As String
    If Len(objectName) = 0 Then Err.Raise vbObjectError + 9956, "AddBackupAccessObjectToNavigationPaneGroup", "No object name was supplied for backup Navigation Pane grouping."
    If AccessObjectExistsInCurrentDatabase(objectType, objectName) = False Then
        RefreshAccessObjectCatalog objectType, False
        If AccessObjectExistsInCurrentDatabase(objectType, objectName) = False Then Err.Raise vbObjectError + 9957, "AddBackupAccessObjectToNavigationPaneGroup", objectKind & " backup [" & objectName & "] could not be grouped because the object does not exist." & vbCrLf & "Source: " & sourceName
    End If
    groupName = BuildNavigationPaneGroupName(DEFAULT_BACKUP_NAVIGATION_PANE_GROUP_LOGICAL_NAME, False)
    RefreshAccessObjectCatalog objectType, False
    If ApplyDeferredNavigationPaneObjectPlacement(objectName, objectType, groupName, sourceName, forceRefreshShortcut, True, groupID, objectID, detailText) = False Then
        If groupID <= 0 Then
            AddNavigationPaneGroupCreationFailure groupName, sourceName, detailText
            Err.Raise vbObjectError + 9958, "AddBackupAccessObjectToNavigationPaneGroup", objectKind & " backup [" & objectName & "] could not be grouped because Navigation Pane group [" & groupName & "] could not be created or resolved." & vbCrLf & "Source: " & sourceName & IIf(Len(detailText) = 0, vbNullString, vbCrLf & detailText)
        End If
        If objectID = 0 Then Err.Raise vbObjectError + 9959, "AddBackupAccessObjectToNavigationPaneGroup", objectKind & " backup [" & objectName & "] could not be grouped because its MSysObjects ID could not be resolved." & vbCrLf & "Source: " & sourceName & IIf(Len(detailText) = 0, vbNullString, vbCrLf & detailText)
        Err.Raise vbObjectError + 9960, "AddBackupAccessObjectToNavigationPaneGroup", objectKind & " backup [" & objectName & "] could not be added to Navigation Pane group [" & groupName & "]." & vbCrLf & "Source: " & sourceName & IIf(Len(detailText) = 0, vbNullString, vbCrLf & detailText)
    End If
    If Len(detailText) > 0 Then AddNavigationPaneGroupingNote detailText, "AddBackupAccessObjectToNavigationPaneGroup"
    ReportCoreActionSuccess "AddBackupAccessObjectToNavigationPaneGroup", "NavigationPane", "Backup object Navigation Pane grouping completed successfully.", "Backup object Navigation Pane grouping complete." & vbCrLf & vbCrLf & "Object: " & objectName & vbCrLf & "Object kind: " & objectKind & vbCrLf & "Backup group: " & groupName & vbCrLf & "Source: " & sourceName, "Add backup object to Navigation Pane group", objectName, objectKind, "CodeDb", "NavigationPaneGrouping", showMessage, "Add Backup Object To Navigation Pane Group", vbInformation, showMessage
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    messageText = EntryPointErrorMessage("Backup object Navigation Pane grouping failed.", "AddBackupAccessObjectToNavigationPaneGroup", savedErrNumber, savedErrSource, savedErrDescription, "Object: " & objectName & vbCrLf & "Object kind: " & objectKind & vbCrLf & "Source: " & sourceName)
    On Error Resume Next
    AddNavigationPaneGroupingNote objectKind & " backup [" & objectName & "] Navigation Pane grouping failed." & vbCrLf & "Source: " & sourceName & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Description: " & savedErrDescription, "AddBackupAccessObjectToNavigationPaneGroup"
    LogTraceMessage "Error", "NavigationPane", "AddBackupAccessObjectToNavigationPaneGroup", "Backup object Navigation Pane grouping failed.", messageText, "Add backup object to Navigation Pane group", savedErrNumber, savedErrSource, savedErrDescription, objectName, objectKind, "CodeDb", "NavigationPaneGrouping"
    If showMessage Then CoreShowErrorMessage messageText, vbCritical, "Add Backup Object To Navigation Pane Group", "AddBackupAccessObjectToNavigationPaneGroup", "Core"
    On Error GoTo 0
    If showMessage Then Exit Sub
    Err.Raise savedErrNumber, "AddBackupAccessObjectToNavigationPaneGroup", messageText
End Sub

Private Sub AddGeneratedAccessObjectToNavigationPaneGroup(ByVal objectName As String, ByVal objectType As AcObjectType, ByVal objectKind As String, ByVal groupName As String, ByVal sourceName As String, Optional ByVal forceRefreshShortcut As Boolean = False, Optional ByVal showMessage As Boolean = True)
    'Runtime generated-object placement policy: selects the requested group or backup path and delegates all physical mutation to REVIEW.
    On Error GoTo ErrHandler
    Dim groupID As Long
    Dim objectID As Long
    Dim detailText As String
    Dim messageText As String
    If Len(objectName) = 0 Then Err.Raise vbObjectError + 9950, "AddGeneratedAccessObjectToNavigationPaneGroup", "No object name was supplied for Navigation Pane grouping."
    If IsBackupAccessObjectName(objectName) Then
        AddBackupAccessObjectToNavigationPaneGroup objectName, objectType, objectKind, sourceName, forceRefreshShortcut, showMessage
        Exit Sub
    End If
    If Len(groupName) = 0 Then
        messageText = objectKind & " [" & objectName & "] was created, but no specific Navigation Pane group was supplied." & vbCrLf & "Source: " & sourceName
        Err.Raise vbObjectError + 9951, "AddGeneratedAccessObjectToNavigationPaneGroup", messageText
    End If
    RefreshAccessObjectCatalog objectType, False
    If ApplyDeferredNavigationPaneObjectPlacement(objectName, objectType, groupName, sourceName, forceRefreshShortcut, False, groupID, objectID, detailText) = False Then
        If groupID <= 0 Then
            AddNavigationPaneGroupCreationFailure groupName, sourceName, detailText
            messageText = objectKind & " [" & objectName & "] was created, but Navigation Pane group [" & groupName & "] could not be created or resolved to an ID." & vbCrLf & "Source: " & sourceName & IIf(Len(detailText) = 0, vbNullString, vbCrLf & detailText)
            Err.Raise vbObjectError + 9952, "AddGeneratedAccessObjectToNavigationPaneGroup", messageText
        End If
        messageText = objectKind & " [" & objectName & "] was created, but could not be added to Navigation Pane group [" & groupName & "]." & vbCrLf & "Source: " & sourceName & IIf(Len(detailText) = 0, vbNullString, vbCrLf & detailText)
        Err.Raise vbObjectError + 9953, "AddGeneratedAccessObjectToNavigationPaneGroup", messageText
    End If
    If Len(detailText) > 0 Then AddNavigationPaneGroupingNote detailText, "AddGeneratedAccessObjectToNavigationPaneGroup"
    ReportCoreActionSuccess "AddGeneratedAccessObjectToNavigationPaneGroup", "NavigationPane", "Access object Navigation Pane grouping completed successfully.", "Access object Navigation Pane grouping complete." & vbCrLf & vbCrLf & "Object: " & objectName & vbCrLf & "Object kind: " & objectKind & vbCrLf & "Group: " & groupName & vbCrLf & "Source: " & sourceName, "Add object to Navigation Pane group", objectName, objectKind, "CodeDb", "NavigationPaneGrouping", showMessage, "Add Object To Navigation Pane Group", vbInformation, showMessage
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    messageText = EntryPointErrorMessage("Access object Navigation Pane grouping failed.", "AddGeneratedAccessObjectToNavigationPaneGroup", savedErrNumber, savedErrSource, savedErrDescription, "Object: " & objectName & vbCrLf & "Object kind: " & objectKind & vbCrLf & "Group: " & groupName & vbCrLf & "Source: " & sourceName)
    On Error Resume Next
    AddNavigationPaneGroupingNote objectKind & " [" & objectName & "] was created, but Navigation Pane grouping failed." & vbCrLf & "Group: " & groupName & vbCrLf & "Source: " & sourceName & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Description: " & savedErrDescription, "AddGeneratedAccessObjectToNavigationPaneGroup"
    LogTraceMessage "Error", "NavigationPane", "AddGeneratedAccessObjectToNavigationPaneGroup", "Access object Navigation Pane grouping failed.", messageText, "Add object to Navigation Pane group", savedErrNumber, savedErrSource, savedErrDescription, objectName, objectKind, "CodeDb", "NavigationPaneGrouping"
    If showMessage Then CoreShowErrorMessage messageText, vbCritical, "Add Object To Navigation Pane Group", "AddGeneratedAccessObjectToNavigationPaneGroup", "Core"
    On Error GoTo 0
    If showMessage Then Exit Sub
    Err.Raise savedErrNumber, "AddGeneratedAccessObjectToNavigationPaneGroup", messageText
End Sub

Private Function EnsureBackupNavigationPaneGroupID(ByVal sourceName As String) As Long
    On Error GoTo ErrHandler
    Dim groupName As String
    Dim failureDetail As String
    groupName = BuildNavigationPaneGroupName(DEFAULT_BACKUP_NAVIGATION_PANE_GROUP_LOGICAL_NAME, False)
    EnsureBackupNavigationPaneGroupID = ResolveOrCreateNavigationPaneGroupID(groupName, vbNullString, sourceName, failureDetail)
    If EnsureBackupNavigationPaneGroupID <= 0 Then
        AddNavigationPaneGroupCreationFailure groupName, sourceName, failureDetail
        AddNavigationPaneGroupingNote "Navigation Pane group [" & groupName & "] could not be created or resolved for backup-only grouping." & vbCrLf & "Source: " & sourceName
    End If
    Exit Function
ErrHandler:
    AddNavigationPaneGroupCreationFailure BuildNavigationPaneGroupName(DEFAULT_BACKUP_NAVIGATION_PANE_GROUP_LOGICAL_NAME, False), sourceName, "Backup group verification failed. Error=" & Err.Description
    EnsureBackupNavigationPaneGroupID = 0
End Function

' Purpose: Initializes navigation pane grouping notes.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub InitializeNavigationPaneGroupingNotes()
    Set mNavigationPaneGroupingNotes = New Collection
    mNavigationPaneGroupingOwnerTraceRunID = GetActiveTraceRunID()
End Sub

Private Sub EnsureNavigationPaneGroupingStateForActiveRun(Optional ByVal forceReset As Boolean = False)
    Dim activeTraceRunID As String
    Dim ownerChanged As Boolean
    activeTraceRunID = GetActiveTraceRunID()
    ownerChanged = (StrComp(activeTraceRunID, mNavigationPaneGroupingOwnerTraceRunID, vbBinaryCompare) <> 0)
    If forceReset Or mNavigationPaneGroupingNotes Is Nothing Or ownerChanged Then
        Set mNavigationPaneGroupingNotes = New Collection
        mNavigationPaneGroupingOwnerTraceRunID = activeTraceRunID
    End If
End Sub

Private Sub AddNavigationPaneGroupingNote(ByVal noteText As String, Optional ByVal sourceProcedure As String = "AddNavigationPaneGroupingNote")
    Dim normalizedNote As String
    Dim traceLevel As String
    normalizedNote = Trim$(noteText)
    If Len(sourceProcedure) = 0 Then sourceProcedure = "AddNavigationPaneGroupingNote"
    If Len(normalizedNote) = 0 Then Exit Sub
    EnsureNavigationPaneGroupingStateForActiveRun
    If NavigationPaneGroupingNoteAlreadyRecorded(normalizedNote) Then Exit Sub
    mNavigationPaneGroupingNotes.Add normalizedNote
    traceLevel = "Detail"
    If CommonTextIndicatesWarning(normalizedNote) Then traceLevel = "Warning"
    LogTraceMessage traceLevel, "NavigationPane", sourceProcedure, "Navigation Pane grouping note recorded.", normalizedNote, vbNullString, 0, vbNullString, vbNullString, vbNullString, "NavigationPane", "CodeDb", "Core"
End Sub

Private Function NavigationPaneGroupingNoteAlreadyRecorded(ByVal normalizedNote As String) As Boolean
    Dim existingNote As Variant
    If mNavigationPaneGroupingNotes Is Nothing Then Exit Function
    For Each existingNote In mNavigationPaneGroupingNotes
        PumpSetupResponsiveness
        If StrComp(CStr(existingNote), normalizedNote, vbTextCompare) = 0 Then
            NavigationPaneGroupingNoteAlreadyRecorded = True
            Exit Function
        End If
    Next existingNote
End Function

Private Sub AddNavigationPaneGroupCreationFailure(ByVal groupName As String, ByVal sourceName As String, ByVal detailText As String)
    Dim failureText As String
    EnsureNavigationPaneGroupingStateForActiveRun
    sourceName = Trim$(sourceName)
    detailText = Trim$(detailText)
    If Len(groupName) = 0 Then groupName = "(blank)"
    failureText = "Navigation Pane group [" & groupName & "] was missing and could not be created or resolved."
    If Len(sourceName) > 0 Then failureText = failureText & " Source: " & sourceName & "."
    If Len(detailText) > 0 Then failureText = failureText & " Detail: " & Replace(detailText, vbCrLf, " ")
    AddNavigationPaneGroupingNote failureText
End Sub

' Purpose: Resolves navigation pane grouping summary for message from the supplied/current ADC state, using `End` and `Next` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function NavigationPaneGroupingSummaryForMessage() As String
    Dim note As Variant
    Dim lines As Collection
    Dim noteCount As Long
    Dim maxNotesToShow As Long
    EnsureNavigationPaneGroupingStateForActiveRun
    If mNavigationPaneGroupingNotes Is Nothing Then Exit Function
    If mNavigationPaneGroupingNotes.count = 0 Then Exit Function
    maxNotesToShow = 25
    ' INTENTIONAL ORDERED VIEW: lines preserves the deterministic sequence required by NavigationPaneGroupingSummaryForMessage; keyed-only or unordered storage would not preserve the consumer semantics.
    Set lines = New Collection
    lines.Add vbCrLf & vbCrLf & "Navigation Pane grouping notes: " & CStr(mNavigationPaneGroupingNotes.count)
    For Each note In mNavigationPaneGroupingNotes
        PumpSetupResponsiveness
        noteCount = noteCount + 1
        If noteCount <= maxNotesToShow Then
            lines.Add "- " & Replace(CStr(note), vbCrLf, " ")
        End If
    Next note
    If mNavigationPaneGroupingNotes.count > maxNotesToShow Then
        lines.Add "- Additional grouping notes were omitted from this message."
    End If
    NavigationPaneGroupingSummaryForMessage = SharedCollection_Join(lines, vbCrLf)
End Function

Private Sub RefreshAccessObjectCatalog(ByVal objectType As AcObjectType, Optional ByVal showMessage As Boolean = True, Optional ByVal failOnError As Boolean = False)
    'Refreshes Access and DAO object catalogs; repeated internal refreshes report one success per object type in the active trace run.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim objectTypeText As String
    objectTypeText = CStr(objectType)
    Set db = currentDb
    Select Case objectType
        Case acTable
            db.TableDefs.Refresh
        Case acQuery
            db.QueryDefs.Refresh
        Case acForm
            db.Containers("Forms").Documents.Refresh
        Case acReport
            db.Containers("Reports").Documents.Refresh
        Case acMacro
            db.Containers("Scripts").Documents.Refresh
    End Select
    Application.RefreshDatabaseWindow
    DoEvents
    If showMessage Or ShouldReportAccessObjectCatalogSuccess(objectType) Then
        ReportCoreActionSuccess "RefreshAccessObjectCatalog", "AccessObjectCatalog", "Access object catalog refresh completed successfully.", "Access object catalog refresh complete." & vbCrLf & vbCrLf & "Object type: " & objectTypeText & vbCrLf & "Database: " & SafeDatabaseName(), "Refresh access object catalog", objectTypeText, "AccessObjectCatalog", "CodeDb", "CatalogRefresh", showMessage, "Refresh Access Object Catalog", vbInformation, False
    End If
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim messageText As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    messageText = EntryPointErrorMessage("Access object catalog refresh failed.", "RefreshAccessObjectCatalog", savedErrNumber, savedErrSource, savedErrDescription, "Object type: " & CStr(objectType) & vbCrLf & "Continuing with available catalog state unless the caller requested failOnError=True.")
    CorePulseProgress "Access object catalog refresh failed; continuing with available catalog state."
    LogTraceMessage "Warning", "AccessObjectCatalog", "RefreshAccessObjectCatalog", "Access object catalog refresh failed; continuing with available catalog state.", messageText, "Refresh access object catalog", savedErrNumber, savedErrSource, savedErrDescription, CStr(objectType), "AccessObjectCatalog", "CodeDb", "CatalogRefresh"
    If showMessage Then
        CoreShowErrorMessage messageText, vbCritical, "Refresh Access Object Catalog", "RefreshAccessObjectCatalog", "Core"
        Exit Sub
    End If
    If failOnError Then Err.Raise savedErrNumber, "RefreshAccessObjectCatalog", messageText
    Err.Clear
End Sub

Private Function ShouldReportAccessObjectCatalogSuccess(ByVal objectType As AcObjectType) As Boolean
    Dim traceRunID As String
    Dim objectTypeKey As String
    If TraceRunActive() = False Then
        ShouldReportAccessObjectCatalogSuccess = True
        Exit Function
    End If
    traceRunID = GetActiveOrLastTraceRunID()
    If StrComp(traceRunID, mAccessObjectCatalogReportedRunID, vbBinaryCompare) <> 0 Then
        Set mAccessObjectCatalogReportedTypes = New Collection
        mAccessObjectCatalogReportedRunID = traceRunID
    ElseIf mAccessObjectCatalogReportedTypes Is Nothing Then
        Set mAccessObjectCatalogReportedTypes = New Collection
    End If
    objectTypeKey = CStr(CLng(objectType))
    If SharedCollection_KeyExists(mAccessObjectCatalogReportedTypes, objectTypeKey) Then Exit Function
    MarkCollectionTextKey mAccessObjectCatalogReportedTypes, objectTypeKey
    ShouldReportAccessObjectCatalogSuccess = True
End Function

Private Function HealthArgText(ByVal argValue As Variant) As String
    On Error GoTo Fallback
    If IsError(argValue) Then Exit Function
    If IsNull(argValue) Then Exit Function
    If IsEmpty(argValue) Then Exit Function
    HealthArgText = CStr(argValue)
    Exit Function
Fallback:
    HealthArgText = vbNullString
End Function

' Purpose: Compares the supplied/current values required by text property matches and reports whether they match under this function's exact comparison rules.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   targetObject (Input; Object): Access/DAO/VBE object instance whose named property/state is being inspected, validated, or updated by this helper.
'   propertyName (Input; String): Exact Access/DAO property name to read, compare, create, update, or fingerprint.
'   expectedValue (Input; String): Exact expected text used for comparison/verification against the current Access property, metadata, or source value.
' Returns: Boolean: True when the text property matches condition is satisfied for the supplied or current ADC context; False otherwise.
Public Function TextPropertyMatches(ByVal targetObject As Object, ByVal propertyName As String, ByVal expectedValue As String) As Boolean
    On Error GoTo MissingProperty
    TextPropertyMatches = (StrComp(CStr(targetObject.properties(propertyName).value), expectedValue, vbBinaryCompare) = 0)
    Exit Function
MissingProperty:
    Err.Clear
End Function

' =============================================================================================================================================================================
' HEALTH REVIEW PROCEDURES
' =============================================================================================================================================================================

' Purpose: Runs the CodeDb health-review workflow, collects current diagnostic findings into the managed review output, and coordinates progress plus optional user-facing failure reporting.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub RunHealthReview(Optional ByVal showMessage As Boolean = True)
    On Error GoTo ErrHandler
    Dim progressStarted As Boolean
    progressStarted = CoreBeginOperationProgressIfNeeded("Review " & DEFAULT_UTILITY_SHORT_NAME & " health", 1)
    CoreUpdateSetupProgress "Collect health review findings"
    If RunHealthReviewCore(showMessage) = False Then
        Err.Raise vbObjectError + 5600, "RunHealthReview", "ADC Health Review did not complete successfully." & vbCrLf & "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB & vbCrLf & "Health review table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReview") & vbCrLf & "Review the health-review output table for row-level findings."
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
        CoreShowErrorMessage "Could not complete ADC Health Review." & vbCrLf & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, vbCritical, "Health Review", "RunHealthReview", "Review"
        Exit Sub
    End If
    Err.Raise savedErrNumber, "RunHealthReview", "Could not complete ADC Health Review." & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Function RunHealthReviewCore(Optional ByVal showMessage As Boolean = True) As Boolean
    On Error GoTo ErrHandler
    Dim rsHealth As DAO.Recordset
    Dim issueHashes As Object
    Dim reviewRunID As String
    Dim reviewRunOn As Date
    Dim infoCount As Long
    Dim warningCount As Long
    Dim errorCount As Long
    Dim criticalCount As Long
    Dim totalCount As Long
    Dim messageText As String
    Dim operationalErrorText As String
    OpenHealthReviewWriteContext True, "HR", rsHealth, issueHashes, reviewRunID, reviewRunOn
    CorePulseProgress "Review schema health"
    AddHealthSchemaChecks rsHealth, issueHashes, reviewRunID, reviewRunOn
    CorePulseProgress "Review managed-object and managed-table description coverage"
    AddHealthManagedDescriptionChecks rsHealth, issueHashes, reviewRunID, reviewRunOn, operationalErrorText
    CorePulseProgress "Review settings health"
    AddHealthSettingsChecks rsHealth, issueHashes, reviewRunID, reviewRunOn, operationalErrorText
    CorePulseProgress "Review query health"
    AddHealthQueryChecks rsHealth, issueHashes, reviewRunID, reviewRunOn
    CorePulseProgress "Review config and candidate health"
    AddHealthConfigAndCandidateChecks rsHealth, issueHashes, reviewRunID, reviewRunOn, operationalErrorText
    CorePulseProgress "Review duplicate health"
    AddHealthDuplicateChecks rsHealth, issueHashes, reviewRunID, reviewRunOn
    CorePulseProgress "Review system macro health"
    AddHealthSystemMacroChecks rsHealth, issueHashes, reviewRunID, reviewRunOn
    CorePulseProgress "Review generated macro health"
    AddHealthGeneratedMacroChecks rsHealth, issueHashes, reviewRunID, reviewRunOn, operationalErrorText
    CorePulseProgress "Review managed form health"
    CorePulseProgress "Review public entry point health"
    AddHealthPublicEntryPointChecks rsHealth, issueHashes, reviewRunID, reviewRunOn
    rsHealth.Close
    Set rsHealth = Nothing
    If Len(Trim$(operationalErrorText)) > 0 Then Err.Raise vbObjectError + 5601, "RunHealthReviewCore", "One or more health-review sections encountered operational errors." & vbCrLf & operationalErrorText
    infoCount = HealthFindingCountBySeverity(reviewRunID, HEALTH_SEVERITY_INFO)
    warningCount = HealthFindingCountBySeverity(reviewRunID, HEALTH_SEVERITY_WARNING)
    errorCount = HealthFindingCountBySeverity(reviewRunID, HEALTH_SEVERITY_ERROR)
    criticalCount = HealthFindingCountBySeverity(reviewRunID, HEALTH_SEVERITY_CRITICAL)
    totalCount = infoCount + warningCount + errorCount + criticalCount
    If showMessage Then
        OpenReviewQuery ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_System_HealthReview"), "RunHealthReviewCore"
        messageText = "ADC Health Review complete." & vbCrLf & vbCrLf & "ReviewRunID: " & reviewRunID & vbCrLf & "Total findings: " & CStr(totalCount) & vbCrLf & "Critical: " & CStr(criticalCount) & vbCrLf & "Error: " & CStr(errorCount) & vbCrLf & "Warning: " & CStr(warningCount) & vbCrLf & "Info: " & CStr(infoCount) & vbCrLf & vbCrLf & "Findings were written to: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReview") & vbCrLf & "Current-run query: " & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_System_HealthReview")
        CoreShowMessage messageText, vbInformation, "Health Review"
    End If
    RunHealthReviewCore = True
    Exit Function
ErrHandler:
    Dim handlerErrNumber As Long
    Dim handlerErrDescription As String
    handlerErrNumber = Err.Number
    handlerErrDescription = Err.Description
    On Error Resume Next
    If Not rsHealth Is Nothing Then rsHealth.Close
    RunHealthReviewCore = False
    On Error GoTo 0
    Err.Raise handlerErrNumber, "RunHealthReviewCore", "Could not run the target review health review workflow." & vbCrLf & "Health table: " & SafeTableName("health") & vbCrLf & "Err.Description: " & handlerErrDescription
End Function

' Purpose: Persists stale candidate review health findings for findings.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   findings (Input; Collection): Collection containing the findings items that the routine consumes or updates as part of this operation.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub PersistStaleCandidateReviewHealthFindings(ByVal findings As Collection)
    'Review-owned stale-candidate review persistence entry point: candidate analysis supplies findings; Review refreshes the stale-candidate review subset while preserving other findings in the single current run.
    On Error GoTo ErrHandler
    Dim rsHealth As DAO.Recordset
    Dim issueHashes As Object
    Dim reviewRunID As String
    Dim reviewRunOn As Date
    Dim findingData As Variant
    OpenHealthReviewWriteContext False, "STALE_PREVIEW", rsHealth, issueHashes, reviewRunID, reviewRunOn
    If Not findings Is Nothing Then
        For Each findingData In findings
            AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_CANDIDATE, HEALTH_SEVERITY_WARNING, CStr(findingData(0)), findingData(1), findingData(2), findingData(3), findingData(4), findingData(5), findingData(6), findingData(7), findingData(8), findingData(9), findingData(10), findingData(11), findingData(12), findingData(13), findingData(14), findingData(15), findingData(16), findingData(17), findingData(18), findingData(19)
        Next findingData
    End If
    rsHealth.Close
    Set rsHealth = Nothing
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not rsHealth Is Nothing Then rsHealth.Close
    On Error GoTo 0
    Err.Raise savedErrNumber, "PersistStaleCandidateReviewHealthFindings", "Could not persist stale-candidate review findings through the Health Review authority." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub OpenHealthReviewWriteContext(ByVal startNewRun As Boolean, ByVal runPrefix As String, ByRef rsHealth As DAO.Recordset, ByRef issueHashes As Object, ByRef reviewRunID As String, ByRef reviewRunOn As Date)
    'Single Review-owned HealthReview write-context authority: full Health Review starts a fresh current run; stale preview reuses the one current run and never creates a competing current run.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rsCurrent As DAO.Recordset
    Dim issueHash As String
    Dim currentRunFound As Boolean
    If Len(runPrefix) = 0 Then Err.Raise vbObjectError + 5602, "OpenHealthReviewWriteContext", "Health Review run prefix is blank."
    Set db = ADCDb()
    RepairHealthReviewTableForSetup True
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: issueHashes holds planned/recovery work or independently observed validation evidence for OpenHealthReviewWriteContext; keep it distinct from the source authority.
    Set issueHashes = CreateObject("Scripting.Dictionary")
    If startNewRun Then
        reviewRunOn = Now
        reviewRunID = runPrefix & "_" & Format$(reviewRunOn, "yyyymmdd_hhnnss") & "_" & StableTextHashHex(CStr(Timer) & MODULE_TAG_DELIMITER & db.Name)
        db.Execute "UPDATE [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReview") & "] SET " & SQL_FIELD_IS_CURRENT_RUN & "=False WHERE " & SQL_FIELD_IS_CURRENT_RUN & "<>False OR " & SQL_FIELD_IS_CURRENT_RUN & " Is Null", dbFailOnError
    Else
        Set rsCurrent = db.OpenRecordset("SELECT Nz(" & SQL_FIELD_REVIEW_RUN_ID & ",'') AS [CurrentReviewRunID], Max(" & SQL_FIELD_REVIEW_RUN_ON & ") AS [CurrentReviewRunOn] FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReview") & "] WHERE " & SQL_FIELD_IS_CURRENT_RUN & "<>False OR " & SQL_FIELD_IS_CURRENT_RUN & " Is Null GROUP BY Nz(" & SQL_FIELD_REVIEW_RUN_ID & ",'') ORDER BY Max(" & SQL_FIELD_REVIEW_RUN_ON & ") DESC;", dbOpenSnapshot)
        If rsCurrent.EOF = False Then
            currentRunFound = True
            reviewRunID = Trim$(CStr(rsCurrent.fields("CurrentReviewRunID").value))
            If Len(reviewRunID) = 0 Then Err.Raise vbObjectError + 5603, "OpenHealthReviewWriteContext", "The current Health Review run has a blank ReviewRunID. Run Health Review to restore a canonical current run before persisting stale-candidate review findings."
            If IsNull(rsCurrent.fields("CurrentReviewRunOn").value) Then reviewRunOn = Now Else reviewRunOn = CDate(rsCurrent.fields("CurrentReviewRunOn").value)
            rsCurrent.MoveNext
            If rsCurrent.EOF = False Then Err.Raise vbObjectError + 5604, "OpenHealthReviewWriteContext", "More than one Health Review run is marked current. Run Health Review to retire legacy competing current runs before persisting stale-candidate review findings."
        End If
        rsCurrent.Close
        Set rsCurrent = Nothing
        If currentRunFound = False Then
            reviewRunOn = Now
            reviewRunID = runPrefix & "_" & Format$(reviewRunOn, "yyyymmdd_hhnnss") & "_" & StableTextHashHex(CStr(Timer) & MODULE_TAG_DELIMITER & db.Name)
        Else
            'Refresh only stale-candidate-review-owned findings inside the existing current run; preserve all other Health Review findings in that run.
            db.Execute "DELETE FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReview") & "] WHERE (" & SQL_FIELD_IS_CURRENT_RUN & "<>False OR " & SQL_FIELD_IS_CURRENT_RUN & " Is Null) AND " & SQL_FIELD_REVIEW_RUN_ID & "=" & SharedDao_SqlQuote(reviewRunID) & " AND " & SQL_FIELD_ISSUE_CODE & " IN ('CANDIDATE_STALE_REVIEW','CANDIDATE_STALE_REVIEW_NEEDED')", dbFailOnError
            Set rsCurrent = db.OpenRecordset("SELECT " & SQL_FIELD_ISSUE_HASH & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReview") & "] WHERE (" & SQL_FIELD_IS_CURRENT_RUN & "<>False OR " & SQL_FIELD_IS_CURRENT_RUN & " Is Null) AND " & SQL_FIELD_REVIEW_RUN_ID & "=" & SharedDao_SqlQuote(reviewRunID) & " AND Len(Nz(" & SQL_FIELD_ISSUE_HASH & ",''))>0;", dbOpenSnapshot)
            Do While rsCurrent.EOF = False
                issueHash = Trim$(CStr(rsCurrent.fields(FIELD_ISSUE_HASH).value))
                If Len(issueHash) > 0 And issueHashes.Exists(issueHash) = False Then issueHashes.Add issueHash, True
                rsCurrent.MoveNext
            Loop
            rsCurrent.Close
            Set rsCurrent = Nothing
        End If
    End If
    Set rsHealth = db.OpenRecordset(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReview"), dbOpenDynaset)
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not rsCurrent Is Nothing Then rsCurrent.Close
    On Error GoTo 0
    Err.Raise savedErrNumber, "OpenHealthReviewWriteContext", savedErrDescription & IIf(Len(savedErrSource) > 0, vbCrLf & "Err.Source: " & savedErrSource, vbNullString)
End Sub

Private Sub AddHealthQueryChecks(ByVal rsHealth As DAO.Recordset, ByRef issueHashes As Object, ByVal reviewRunID As String, ByVal reviewRunOn As Date)
    Dim queryName As Variant
    Dim db As DAO.Database
    Dim qdf As DAO.queryDef
    Dim canOpen As Boolean
    Dim detailText As String
    Set db = ADCDb()
    For Each queryName In HealthExpectedQueryNames()
        PumpSetupResponsiveness
        If SharedDao_QueryExists(db, CStr(queryName)) = False Then
            AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_SCHEMA, HEALTH_SEVERITY_ERROR, "QUERY_MISSING", ADCAT_OBJECT_TYPE_QUERY, CStr(queryName), , , , , , , , , , , "query exists", "missing", "Expected Access Development Console saved query is missing.", "Run setup/update to recreate form-support queries."
        Else
            Set qdf = db.QueryDefs(CStr(queryName))
            canOpen = QueryCanOpen(CStr(queryName), detailText)
            If canOpen = False Then
                AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_SCHEMA, HEALTH_SEVERITY_ERROR, "QUERY_CANNOT_OPEN", ADCAT_OBJECT_TYPE_QUERY, CStr(queryName), , , , , , , , , , , "opens", "cannot open", detailText, "Run setup/update or inspect the query SQL."
            End If
            If qdf.parameters.count > 0 Then
                AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_SCHEMA, HEALTH_SEVERITY_WARNING, "QUERY_HAS_PARAMETERS", ADCAT_OBJECT_TYPE_QUERY, CStr(queryName), , , , , , , , , , , "0 parameters", CStr(qdf.parameters.count), "Saved query may prompt for parameters.", "Inspect field/table references in the query SQL."
            End If
        End If
    Next queryName
End Sub

Private Function HealthExpectedQueryNames() As Collection
    'Derived solely from canonical managed-object classifications; Review health owns no independent query-membership roster.
    ' INTENTIONAL FILTERED PARTITION: rows is a consumer-required subset/projection derived from the source enumerated by HealthExpectedQueryNames; its membership is not independently authoritative.
    Dim rows As New Collection
    Dim definitionData As Variant
    For Each definitionData In ManagedObjectReviewQueryDefinitions()
        rows.Add CStr(definitionData(0))
    Next definitionData
    Set HealthExpectedQueryNames = rows
End Function

Private Function QueryCanOpen(ByVal queryName As String, ByRef detailText As String) As Boolean
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Set rs = ADCDb().OpenRecordset(queryName, dbOpenSnapshot)
    QueryCanOpen = True
ExitHere:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Exit Function
ErrHandler:
    detailText = "Could not open query [" & queryName & "]. Err.Number: " & CStr(Err.Number) & "; Err.Description: " & Err.Description
    QueryCanOpen = False
    Resume ExitHere
End Function

' Purpose: Opens review query for query name, source name.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   queryName (Input; String): Physical Access query name identifying the saved query the routine creates, validates, reads, or updates.
'   sourceName (Input; String): Name of the source object/query/table/control from which this operation reads or to which the generated definition is bound.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub OpenReviewQuery(ByVal queryName As String, ByVal sourceName As String)
    On Error GoTo ErrHandler
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    DoCmd.OpenQuery queryName
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    CoreAddOperationNote "Could not open review query/table from " & sourceName & ". Query: " & queryName & "; Err.Description: " & savedErrDescription
    Err.Raise savedErrNumber, "OpenReviewQuery", "Unable to open required Access Development Console review/source query." & vbCrLf & "Source: " & sourceName & vbCrLf & "Query: " & queryName & vbCrLf & "Original error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription
End Sub

Private Sub AddHealthGeneratedMacroChecks(ByVal rsHealth As DAO.Recordset, ByRef issueHashes As Object, ByVal reviewRunID As String, ByVal reviewRunOn As Date, ByRef operationalErrorText As String)
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim cfg As TMacroRecord
    Dim macroName As String
    Dim macroText As String
    Dim errorText As String
    Dim wrapperName As String
    Dim expectedWrapperName As String
    Dim parsedMacroKey As String
    Dim parsedOpenFormTarget As String
    Dim parsedOpenFormTargets As Collection
    Dim obj As accessObject
    Dim orphanKey As String
    Dim configReviewedCount As Long
    Dim macroReviewedCount As Long
    If CodeDbTableExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")) Then
        Set rs = ADCDb().OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "] ORDER BY " & SQL_FIELD_ID, dbOpenSnapshot)
        Do While Not rs.EOF
            PumpSetupResponsiveness
            configReviewedCount = configReviewedCount + 1
            PulseProgressEvery configReviewedCount, 25, "Generated macro config rows reviewed"
            MacroRecordFromRecordset rs, cfg
            macroName = x__UTIL_ADC_MACRO_RUNTIME.ExpectedGeneratedMacroObjectNameForConfig(cfg)
            If cfg.LastGeneratedOn <> 0 Or Len(Trim$(cfg.GeneratedMacroObjectName)) > 0 Then
                If Len(macroName) = 0 Then
                    AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_GENERATED_MACRO, HEALTH_SEVERITY_ERROR, "GENERATED_MACRO_NAME_BLANK", "ConfigRow", x__UTIL_ADC_MACRO_RUNTIME.ExpectedGeneratedMacroObjectNameForConfig(cfg), cfg.TargetProcedureName, cfg.MacroKey, cfg.MacroTitle, cfg.ID, vbNullString, ROW_KIND_CONFIGURATION, cfg.ConfigurationStatus, vbNullString, cfg.ConfigOwnerDatabaseRole, vbNullString, cfg.TargetProcedureName, cfg.TargetProcedureModuleName, cfg.GeneratedMacroObjectName, vbNullString, vbNullString, "Generated macro object name is blank.", "Regenerate after fixing MacroTitle."
                ElseIf AccessObjectExistsInCurrentDatabase(acMacro, macroName) = False Then
                    AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_GENERATED_MACRO, HEALTH_SEVERITY_ERROR, "GENERATED_MACRO_MISSING", "ConfigRow", x__UTIL_ADC_MACRO_RUNTIME.ExpectedGeneratedMacroObjectNameForConfig(cfg), cfg.TargetProcedureName, cfg.MacroKey, cfg.MacroTitle, cfg.ID, vbNullString, ROW_KIND_CONFIGURATION, cfg.ConfigurationStatus, vbNullString, cfg.ConfigOwnerDatabaseRole, vbNullString, cfg.TargetProcedureName, cfg.TargetProcedureModuleName, cfg.GeneratedMacroObjectName, vbNullString, vbNullString, "Generated macro object is missing: " & macroName, "Regenerate this macro if the config row is still active."
                ElseIf x__UTIL_ADC_MACRO_RUNTIME.TryExportMacroText(macroName, macroText, errorText) = False Then
                    AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_GENERATED_MACRO, HEALTH_SEVERITY_WARNING, "GENERATED_MACRO_UNREADABLE", "ConfigRow", x__UTIL_ADC_MACRO_RUNTIME.ExpectedGeneratedMacroObjectNameForConfig(cfg), cfg.TargetProcedureName, cfg.MacroKey, cfg.MacroTitle, cfg.ID, vbNullString, ROW_KIND_CONFIGURATION, cfg.ConfigurationStatus, vbNullString, cfg.ConfigOwnerDatabaseRole, vbNullString, cfg.TargetProcedureName, cfg.TargetProcedureModuleName, cfg.GeneratedMacroObjectName, vbNullString, vbNullString, errorText, "Open or regenerate the macro object."
                Else
                    wrapperName = GeneratedMacroWrapperNameFromMacroText(macroText)
                    parsedMacroKey = GeneratedMacroKeyFromMacroText(macroText)
                    Select Case UCase$(Trim$(cfg.MacroAction))
                        Case UCase$(ACTION_OPENFORM)
                            Set parsedOpenFormTargets = x__UTIL_ADC_MACRO_RUNTIME.OpenFormTargetsFromMacroText(macroText)
                            parsedOpenFormTarget = vbNullString
                            If parsedOpenFormTargets.count > 0 Then parsedOpenFormTarget = CStr(parsedOpenFormTargets(1))
                            If StrComp(parsedOpenFormTarget, x__UTIL_ADC_MACRO_RUNTIME.ConfiguredObjectActionTargetName(cfg), vbTextCompare) <> 0 Then
                                AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_GENERATED_MACRO, HEALTH_SEVERITY_ERROR, "GENERATED_MACRO_WRONG_TARGET", "ConfigRow", x__UTIL_ADC_MACRO_RUNTIME.ExpectedGeneratedMacroObjectNameForConfig(cfg), cfg.TargetProcedureName, cfg.MacroKey, cfg.MacroTitle, cfg.ID, vbNullString, ROW_KIND_CONFIGURATION, cfg.ConfigurationStatus, vbNullString, cfg.ConfigOwnerDatabaseRole, vbNullString, cfg.TargetProcedureName, cfg.TargetProcedureModuleName, cfg.GeneratedMacroObjectName, vbNullString, vbNullString, "Generated OpenForm macro target [" & parsedOpenFormTarget & "] does not match current TargetObjectName [" & x__UTIL_ADC_MACRO_RUNTIME.ConfiguredObjectActionTargetName(cfg) & "].", "Regenerate the macro object."
                            End If
                        Case UCase$(ACTION_RUNCODE), UCase$(ACTION_OPENTABLE), UCase$(ACTION_OPENQUERY), UCase$(ACTION_OPENREPORT), UCase$(ACTION_RUNMACRO)
                            If GeneratedRunCodeMacroRequiresConfirmation(cfg) Then
                                expectedWrapperName = "RunConfirmedConfiguredMacroTarget"
                            Else
                                expectedWrapperName = "RunConfiguredMacroTarget"
                            End If
                            If StrComp(wrapperName, expectedWrapperName, vbTextCompare) <> 0 Then
                                AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_GENERATED_MACRO, HEALTH_SEVERITY_ERROR, "GENERATED_MACRO_WRONG_WRAPPER", "ConfigRow", x__UTIL_ADC_MACRO_RUNTIME.ExpectedGeneratedMacroObjectNameForConfig(cfg), cfg.TargetProcedureName, cfg.MacroKey, cfg.MacroTitle, cfg.ID, vbNullString, ROW_KIND_CONFIGURATION, cfg.ConfigurationStatus, vbNullString, cfg.ConfigOwnerDatabaseRole, vbNullString, cfg.TargetProcedureName, cfg.TargetProcedureModuleName, cfg.GeneratedMacroObjectName, vbNullString, vbNullString, "Generated " & cfg.MacroAction & " dispatcher macro does not call expected wrapper [" & expectedWrapperName & "].", "Regenerate the macro object."
                            End If
                            If StrComp(parsedMacroKey, cfg.MacroKey, vbBinaryCompare) <> 0 Then
                                AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_GENERATED_MACRO, HEALTH_SEVERITY_ERROR, "GENERATED_MACRO_WRONG_MACROKEY", "ConfigRow", x__UTIL_ADC_MACRO_RUNTIME.ExpectedGeneratedMacroObjectNameForConfig(cfg), cfg.TargetProcedureName, cfg.MacroKey, cfg.MacroTitle, cfg.ID, vbNullString, ROW_KIND_CONFIGURATION, cfg.ConfigurationStatus, vbNullString, cfg.ConfigOwnerDatabaseRole, vbNullString, cfg.TargetProcedureName, cfg.TargetProcedureModuleName, cfg.GeneratedMacroObjectName, vbNullString, vbNullString, "Generated macro passes MacroKey [" & parsedMacroKey & "] instead of [" & cfg.MacroKey & "].", "Regenerate the macro object."
                            End If
                    End Select
                End If
            End If
            rs.MoveNext
        Loop
    End If
    For Each obj In Application.CurrentProject.AllMacros
        PumpSetupResponsiveness
        macroReviewedCount = macroReviewedCount + 1
        PulseProgressEvery macroReviewedCount, 25, "Generated macro objects reviewed"
        macroText = vbNullString
        errorText = vbNullString
        If x__UTIL_ADC_MACRO_RUNTIME.TryExportMacroText(obj.Name, macroText, errorText) Then
            wrapperName = GeneratedMacroWrapperNameFromMacroText(macroText)
            If Len(wrapperName) > 0 Then
                orphanKey = GeneratedMacroKeyFromMacroText(macroText)
                If Len(orphanKey) > 0 And ConfigMacroKeyExists(orphanKey) = False Then
                    AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_GENERATED_MACRO, HEALTH_SEVERITY_WARNING, "GENERATED_MACRO_ORPHAN", ADCAT_OBJECT_TYPE_MACRO, obj.Name, , orphanKey, , , , , , , , , "MacroKey exists in config", "missing config row", "Access Development Console-owned generated macro calls a MacroKey that is not configured.", "Review and delete manually if obsolete."
                End If
            End If
        End If
    Next obj
ExitHere:
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
    AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_GENERATED_MACRO, HEALTH_SEVERITY_ERROR, "GENERATED_MACRO_REVIEW_FAILED", ADCAT_OBJECT_TYPE_MACRO, vbNullString, , , , , , , , , , , , , "Generated macro review failed: " & savedErrDescription, "Inspect generated macro objects manually."
    operationalErrorText = SharedText_AppendNonBlankLine(operationalErrorText, "Generated macro health review failed. Error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription)
    Resume ExitHere
End Sub

Private Function GeneratedMacroWrapperNameFromMacroText(ByVal macroText As String) As String
    Dim targets As Collection
    Dim item As Variant
    Dim procName As String
    Set targets = x__UTIL_ADC_MACRO_RUNTIME.RunCodeTargetsFromMacroText(macroText)
    For Each item In targets
        PumpSetupResponsiveness
        procName = ProcedureNameFromRunCodeExpression(CStr(item))
        If StrComp(procName, "RunConfiguredMacroTarget", vbTextCompare) = 0 Or StrComp(procName, "RunConfirmedConfiguredMacroTarget", vbTextCompare) = 0 Then
            GeneratedMacroWrapperNameFromMacroText = procName
            Exit Function
        End If
    Next item
End Function

Private Function GeneratedMacroKeyFromMacroText(ByVal macroText As String) As String
    Dim targets As Collection
    Dim item As Variant
    Dim procName As String
    Set targets = x__UTIL_ADC_MACRO_RUNTIME.RunCodeTargetsFromMacroText(macroText)
    For Each item In targets
        PumpSetupResponsiveness
        procName = ProcedureNameFromRunCodeExpression(CStr(item))
        If StrComp(procName, "RunConfiguredMacroTarget", vbTextCompare) = 0 Or StrComp(procName, "RunConfirmedConfiguredMacroTarget", vbTextCompare) = 0 Then
            GeneratedMacroKeyFromMacroText = FirstStringArgumentFromFunctionCall(CStr(item))
            Exit Function
        End If
    Next item
End Function

' Purpose: Requires default code DB generated macros navigation pane group for role context for source name.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   sourceName (Input; String): Name of the source object/query/table/control from which this operation reads or to which the generated definition is bound.
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub RequireDefaultCodeDbGeneratedMacrosNavigationPaneGroupForRoleContext(ByVal sourceName As String, Optional ByVal showMessage As Boolean = True)
    'Intended use: Runs Require Default Code Db Generated Macros Navigation Pane Group For Role Context for direct, generated-macro, form-button, or required cross-module callers.
    On Error GoTo ErrHandler
    Dim navigationDefinition As Variant
    Dim groupName As String
    Dim SettingKey As String
    Dim checkedCount As Long
    Dim warningCount As Long
    Dim messageText As String
    If mNavigationPaneGroupingNotes Is Nothing Then InitializeNavigationPaneGroupingNotes
    If SettingBool(SETTING_ADD_TO_NAV_GROUP, True) = False Then
        ReportCoreActionSuccess "RequireDefaultCodeDbGeneratedMacrosNavigationPaneGroupForRoleContext", "NavigationPane", "Default Navigation Pane group check skipped because grouping is disabled.", "Default Navigation Pane group check skipped." & vbCrLf & vbCrLf & "Reason: " & SETTING_ADD_TO_NAV_GROUP & " is False." & vbCrLf & "Source: " & sourceName, "Verify default Navigation Pane groups", vbNullString, "NavigationPane", "CodeDb", "NavigationPaneGrouping", showMessage, "Default Navigation Pane Groups"
        Exit Sub
    End If
    For Each navigationDefinition In NavigationPaneGroupDefinitions()
        SettingKey = Trim$(CStr(navigationDefinition(3)))
        If Len(SettingKey) > 0 Then
            PumpSetupResponsiveness
            checkedCount = checkedCount + 1
            groupName = DefaultNavigationPaneGroupNameBySetting(SettingKey, CStr(navigationDefinition(1)))
            If Len(groupName) = 0 Then
                warningCount = warningCount + 1
                AddNavigationPaneGroupingNote "Default Navigation Pane group could not be verified because [" & SettingKey & "] is blank." & vbCrLf & "Source: " & sourceName, "RequireDefaultCodeDbGeneratedMacrosNavigationPaneGroupForRoleContext"
            ElseIf VerifyDefaultCodeDbGeneratedMacrosNavigationPaneGroupExists(groupName, SettingKey) = False Then
                warningCount = warningCount + 1
                AddNavigationPaneGroupingNote "Default Navigation Pane group [" & groupName & "] could not be verified or created before the system macro operation." & vbCrLf & "Setting: " & SettingKey & vbCrLf & "Source: " & sourceName, "RequireDefaultCodeDbGeneratedMacrosNavigationPaneGroupForRoleContext"
            End If
        End If
    Next navigationDefinition
    If EnsureBackupNavigationPaneGroupID(sourceName) <= 0 Then warningCount = warningCount + 1
    If warningCount = 0 Then
        ReportCoreActionSuccess "RequireDefaultCodeDbGeneratedMacrosNavigationPaneGroupForRoleContext", "NavigationPane", "Default Navigation Pane groups verified successfully.", "Default Navigation Pane groups verified." & vbCrLf & vbCrLf & "Source: " & sourceName & vbCrLf & "Checked group settings: " & CStr(checkedCount), "Verify default Navigation Pane groups", vbNullString, "NavigationPane", "CodeDb", "NavigationPaneGrouping", showMessage, "Default Navigation Pane Groups"
    Else
        messageText = "Default Navigation Pane group verification completed with warning(s)." & vbCrLf & vbCrLf & "Procedure: RequireDefaultCodeDbGeneratedMacrosNavigationPaneGroupForRoleContext" & vbCrLf & "Source: " & sourceName & vbCrLf & "Checked group settings: " & CStr(checkedCount) & vbCrLf & "Warnings: " & CStr(warningCount) & TableDiagnosticsLocationForMessage() & NavigationPaneGroupingSummaryForMessage()
        CorePulseProgress "Default Navigation Pane group verification completed with warnings."
        LogTraceMessage "Warning", "NavigationPane", "RequireDefaultCodeDbGeneratedMacrosNavigationPaneGroupForRoleContext", "Default Navigation Pane group verification completed with warnings.", messageText, "Verify default Navigation Pane groups", 0, vbNullString, vbNullString, vbNullString, "NavigationPane", "CodeDb", "NavigationPaneGrouping"
        If showMessage Then CoreShowMessage messageText, vbExclamation, "Default Navigation Pane Groups"
    End If
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    messageText = EntryPointErrorMessage("Could not verify/create the default Navigation Pane group before the system macro operation.", "RequireDefaultCodeDbGeneratedMacrosNavigationPaneGroupForRoleContext", savedErrNumber, savedErrSource, savedErrDescription, "Source: " & sourceName & vbCrLf & "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB)
    AddNavigationPaneGroupingNote "Could not verify/create the default Navigation Pane group before the system macro operation." & vbCrLf & "Source: " & sourceName & vbCrLf & "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, "RequireDefaultCodeDbGeneratedMacrosNavigationPaneGroupForRoleContext"
    LogTraceMessage "Error", "NavigationPane", "RequireDefaultCodeDbGeneratedMacrosNavigationPaneGroupForRoleContext", "Default Navigation Pane group verification failed.", messageText, "Verify default Navigation Pane groups", savedErrNumber, savedErrSource, savedErrDescription, vbNullString, "NavigationPane", "CodeDb", "NavigationPaneGrouping"
    If showMessage Then CoreShowErrorMessage messageText, vbCritical, "Default Navigation Pane Groups", "RequireDefaultCodeDbGeneratedMacrosNavigationPaneGroupForRoleContext", "Core"
End Sub

Private Function VerifyDefaultCodeDbGeneratedMacrosNavigationPaneGroupExists(ByVal groupName As String, Optional ByVal SettingKey As String = vbNullString) As Boolean
    'Intended use: Returns the Verify Default Code Db Generated Macros Navigation Pane Group Exists result for direct, query, form, or required cross-module callers.
    On Error GoTo ErrHandler
    Dim settingKeyForMessage As String
    Dim failureDetail As String
    settingKeyForMessage = SettingKey
    If Len(settingKeyForMessage) = 0 Then settingKeyForMessage = SETTING_DEFAULT_GENERATED_MACROS_NAV_GROUP
    If Len(groupName) = 0 Then Exit Function
    VerifyDefaultCodeDbGeneratedMacrosNavigationPaneGroupExists = (ResolveOrCreateNavigationPaneGroupID(groupName, settingKeyForMessage, "VerifyDefaultCodeDbGeneratedMacrosNavigationPaneGroupExists", failureDetail) > 0)
    If VerifyDefaultCodeDbGeneratedMacrosNavigationPaneGroupExists = False Then AddNavigationPaneGroupCreationFailure groupName, "VerifyDefaultCodeDbGeneratedMacrosNavigationPaneGroupExists", failureDetail
    Exit Function
ErrHandler:
    AddNavigationPaneGroupCreationFailure groupName, "VerifyDefaultCodeDbGeneratedMacrosNavigationPaneGroupExists", "Setting=" & settingKeyForMessage & "; Error=" & Err.Description
    VerifyDefaultCodeDbGeneratedMacrosNavigationPaneGroupExists = False
End Function

Private Sub AddHealthSettingsChecks(ByVal rsHealth As DAO.Recordset, ByRef issueHashes As Object, ByVal reviewRunID As String, ByVal reviewRunOn As Date, ByRef operationalErrorText As String)
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim SettingKey As String
    Dim dataTypeText As String
    Dim valueText As String
    Dim defaultText As String
    Dim requiredValue As Boolean
    Dim dataTypeSupported As Boolean
    Dim valueMatchesDataType As Boolean
    Dim expectedDataTypes As String
    Dim reviewedCount As Long
    expectedDataTypes = SETTING_DATA_TYPE_STRING & "/" & SETTING_DATA_TYPE_LONG & "/" & SETTING_DATA_TYPE_BOOLEAN & "/" & SETTING_DATA_TYPE_DATE & "/" & SETTING_DATA_TYPE_DOUBLE & "/" & SETTING_DATA_TYPE_CURRENCY
    If CodeDbTableExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID)) = False Then Exit Sub
    If TableCanOpen(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID)) = False Then
        AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_SETTINGS, HEALTH_SEVERITY_CRITICAL, "SCHEMA_TABLE_CANNOT_OPEN", ADCAT_OBJECT_TYPE_TABLE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID), , , , , , , , , , , , , "Settings table exists but cannot be opened.", "Run repair tables or setup/update."
        Exit Sub
    End If
    Set rs = ADCDb().OpenRecordset(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID), dbOpenSnapshot)
    Do While Not rs.EOF
        PumpSetupResponsiveness
        reviewedCount = reviewedCount + 1
        PulseProgressEvery reviewedCount, 25, "Settings health rows reviewed"
        SettingKey = SharedText_Nz(FieldValueIfExists(rs, FIELD_SETTING_KEY))
        dataTypeText = SharedText_Nz(FieldValueIfExists(rs, FIELD_MANAGED_OBJECT_FIELD_DATA_TYPE))
        valueText = SharedText_Nz(FieldValueIfExists(rs, FIELD_SETTING_VALUE))
        defaultText = SharedText_Nz(FieldValueIfExists(rs, FIELD_MANAGED_OBJECT_FIELD_DEFAULT_VALUE))
        requiredValue = NzBool(FieldValueIfExists(rs, FIELD_REQUIRED), False)
        If requiredValue And Len(Trim$(valueText & defaultText)) = 0 Then
            AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_SETTINGS, HEALTH_SEVERITY_ERROR, "SETTING_REQUIRED_BLANK", "Setting", SettingKey, , , , 0, SettingKey, , , , , , "nonblank effective value", "blank", "Required setting has no SettingValue or DefaultValue.", "Populate the setting or repair default defaults."
        End If
        valueMatchesDataType = SettingValueMatchesDataType(valueText, dataTypeText, dataTypeSupported)
        If dataTypeSupported = False Then
            AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_SETTINGS, HEALTH_SEVERITY_WARNING, "SETTING_DATATYPE_INVALID", "Setting", SettingKey, , , , 0, SettingKey, , , , , , expectedDataTypes, dataTypeText, "Setting DataType is not recognized.", "Correct DataType or default-repair settings."
        ElseIf Len(Trim$(valueText)) > 0 And valueMatchesDataType = False Then
            AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_SETTINGS, HEALTH_SEVERITY_ERROR, "SETTING_INVALID_VALUE", "Setting", SettingKey, , , , 0, SettingKey, , , , , , dataTypeText, valueText, "SettingValue does not match the declared DataType.", "Correct SettingValue or reset it to the default."
        End If
        rs.MoveNext
    Loop
ExitHere:
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
    AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_SETTINGS, HEALTH_SEVERITY_ERROR, "SETTING_REVIEW_FAILED", ADCAT_OBJECT_TYPE_TABLE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID), , , , , , , , , , , , , "Settings review failed: " & savedErrDescription, "Run setup/update or inspect settings manually."
    operationalErrorText = SharedText_AppendNonBlankLine(operationalErrorText, "Settings health review failed. Error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription)
    Resume ExitHere
End Sub

Private Function SettingValueMatchesDataType(ByVal valueText As String, ByVal dataTypeText As String, ByRef dataTypeSupported As Boolean) As Boolean
    On Error GoTo ErrHandler
    Dim testLong As Long
    Dim testDouble As Double
    Dim testCurrency As Currency
    Dim testDate As Date
    dataTypeSupported = True
    dataTypeText = Trim$(dataTypeText)
    If Len(dataTypeText) = 0 Or StrComp(dataTypeText, SETTING_DATA_TYPE_STRING, vbTextCompare) = 0 Then
        SettingValueMatchesDataType = True
    ElseIf StrComp(dataTypeText, SETTING_DATA_TYPE_BOOLEAN, vbTextCompare) = 0 Then
        If Len(Trim$(valueText)) > 0 Then Call BooleanTextValue(valueText, False)
        SettingValueMatchesDataType = True
    ElseIf StrComp(dataTypeText, SETTING_DATA_TYPE_LONG, vbTextCompare) = 0 Then
        If Len(Trim$(valueText)) > 0 Then testLong = CLng(valueText)
        SettingValueMatchesDataType = True
    ElseIf StrComp(dataTypeText, SETTING_DATA_TYPE_DOUBLE, vbTextCompare) = 0 Then
        If Len(Trim$(valueText)) > 0 Then testDouble = CDbl(valueText)
        SettingValueMatchesDataType = True
    ElseIf StrComp(dataTypeText, SETTING_DATA_TYPE_CURRENCY, vbTextCompare) = 0 Then
        If Len(Trim$(valueText)) > 0 Then testCurrency = CCur(valueText)
        SettingValueMatchesDataType = True
    ElseIf StrComp(dataTypeText, SETTING_DATA_TYPE_DATE, vbTextCompare) = 0 Then
        If Len(Trim$(valueText)) > 0 Then testDate = CDate(valueText)
        SettingValueMatchesDataType = True
    Else
        dataTypeSupported = False
    End If
    Exit Function
ErrHandler:
    SettingValueMatchesDataType = False
End Function

' Purpose: Resolves and returns count duplicate rows by selected state from the supplied inputs and the authoritative ADC definitions/state used by this module.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   selectedOnly (Input; Boolean): True limits the operation to rows/items currently marked selected; False includes the full eligible set.
' Returns: Long: Returns the number of count duplicate rows by selected state items matched, processed, selected, deleted, or reviewed by this operation; zero represents no qualifying items.
Public Function CountDuplicateRowsBySelectedState(ByVal selectedOnly As Boolean) As Long
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim cfg As TMacroRecord
    Set rs = ADCDb().OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "]", dbOpenSnapshot)
    Do While Not rs.EOF
        PumpSetupResponsiveness
        MacroRecordFromRecordset rs, cfg
        If Len(DuplicateIssueTypeForQuery(cfg.ID)) > 0 Then
            If selectedOnly = False Or cfg.SelectedForDuplicateDeletion Then
                CountDuplicateRowsBySelectedState = CountDuplicateRowsBySelectedState + 1
            End If
        End If
        rs.MoveNext
    Loop
ExitHere:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Exit Function
ErrHandler:
    CountDuplicateRowsBySelectedState = 0
    Resume ExitHere
End Function

Private Sub AddDuplicateTargetSignatureHealthChecks(ByVal rsHealth As DAO.Recordset, ByRef issueHashes As Object, ByVal reviewRunID As String, ByVal reviewRunOn As Date)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim cfg As TMacroRecord
    Dim signatureCounts As Object
    Dim signatureDisplay As Object
    Dim signatureKey As Variant
    Dim signatureText As String
    Dim displayText As String
    ' INTENTIONAL KEYED LOOKUP: signatureCounts, signatureDisplay provides keyed membership/lookup over source rows or state already read by AddDuplicateTargetSignatureHealthChecks; it is a derived access path, not
    ' a second authority.
    Set signatureCounts = CreateObject("Scripting.Dictionary")
    Set signatureDisplay = CreateObject("Scripting.Dictionary")
    signatureCounts.CompareMode = vbBinaryCompare
    signatureDisplay.CompareMode = vbBinaryCompare
    If CodeDbTableExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")) = False Then Exit Sub
    Set db = ADCDb()
    Set rs = db.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "] WHERE Nz(" & SQL_FIELD_ACTIVE & ",False)<>False ORDER BY " & SQL_FIELD_ID, dbOpenSnapshot)
    Do While Not rs.EOF
        PumpSetupResponsiveness
        MacroRecordFromRecordset rs, cfg
        signatureText = TargetCallSignature(cfg.ConfigOwnerDatabaseRole, cfg.MacroAction, cfg.TargetProcedureName, cfg.TargetArguments, cfg.TargetObjectName, cfg.ConfirmationMode)
        If Len(signatureText) > 0 Then
            displayText = HealthTargetSignatureDisplay(cfg)
            If signatureCounts.Exists(signatureText) Then
                signatureCounts(signatureText) = CLng(signatureCounts(signatureText)) + 1
            Else
                signatureCounts.Add signatureText, 1&
                signatureDisplay.Add signatureText, displayText
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    For Each signatureKey In signatureCounts.keys
        PumpSetupResponsiveness
        If CLng(signatureCounts(signatureKey)) > 1 Then
            AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_DUPLICATE, HEALTH_SEVERITY_WARNING, "DUPLICATE_TARGET_SIGNATURE", "TargetSignature", CStr(signatureDisplay(signatureKey)), , , , , , , , , , , "unique action-specific target signature", CStr(signatureCounts(signatureKey)) & " rows", "The same action-specific target signature appears in the same ConfigOwnerDatabaseRole.", "Review whether multiple macros intentionally use the same action, target, and arguments."
        End If
    Next signatureKey
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
    Err.Raise savedErrNumber, "AddDuplicateTargetSignatureHealthChecks", "Could not complete duplicate target-signature health checks." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

' Purpose: Resolves and returns count target signature rows from the supplied inputs and the authoritative ADC definitions/state used by this module.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   targetSignature (Input; String): Expected public procedure signature used to compare the discovered callable declaration against the required API contract.
' Returns: Long: Returns the number of count target signature rows items matched, processed, selected, deleted, or reviewed by this operation; zero represents no qualifying items.
Public Function CountTargetSignatureRows(ByVal targetSignature As String) As Long
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim rowSignature As String
    If Len(targetSignature) = 0 Then Exit Function
    Set db = ADCDb()
    Set rs = db.OpenRecordset("SELECT [" & FIELD_CONFIG_OWNER_DATABASE_ROLE & "], " & SQL_FIELD_MACRO_ACTION & ", " & SQL_FIELD_TARGET_PROCEDURE_NAME & ", " & SQL_FIELD_TARGET_ARGUMENTS & ", " & SQL_FIELD_TARGET_OBJECT_NAME & ", " & SQL_FIELD_CONFIRMATION_MODE & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "]", dbOpenSnapshot)
    Do While Not rs.EOF
        PumpSetupResponsiveness
        rowSignature = TargetCallSignature(SharedText_Nz(rs.fields(FIELD_CONFIG_OWNER_DATABASE_ROLE).value), SharedText_Nz(rs.fields(FIELD_MACRO_ACTION).value), SharedText_Nz(rs.fields(FIELD_TARGET_PROCEDURE_NAME).value), SharedText_Nz(rs.fields(FIELD_TARGET_ARGUMENTS).value), SharedText_Nz(rs.fields(FIELD_TARGET_OBJECT_NAME).value), SharedText_Nz(rs.fields(FIELD_CONFIRMATION_MODE).value))
        If StrComp(rowSignature, targetSignature, vbBinaryCompare) = 0 Then CountTargetSignatureRows = CountTargetSignatureRows + 1
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
    Err.Raise savedErrNumber, "CountTargetSignatureRows", "Could not count MC_Records configuration rows by the authoritative action-aware binary target signature." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Sub AddHealthConfigAndCandidateChecks(ByVal rsHealth As DAO.Recordset, ByRef issueHashes As Object, ByVal reviewRunID As String, ByVal reviewRunOn As Date, ByRef operationalErrorText As String)
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim cfg As TMacroRecord
    Dim candidate As TCandidateRecord
    Dim staleReason As String
    Dim reviewReason As String
    Dim latestBatchID As String
    Dim reviewedCount As Long
    If CodeDbTableExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")) Then
        If TableCanOpen(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")) Then
            Set rs = ADCDb().OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "] ORDER BY " & SQL_FIELD_ID, dbOpenSnapshot)
            Do While Not rs.EOF
                PumpSetupResponsiveness
                reviewedCount = reviewedCount + 1
                PulseProgressEvery reviewedCount, 25, "Configuration health rows reviewed"
                MacroRecordFromRecordset rs, cfg
                AddHealthOneConfigRowChecks rsHealth, issueHashes, reviewRunID, reviewRunOn, cfg, 0
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing
        End If
    End If
    If CodeDbTableExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates")) Then
        If TableCanOpen(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates")) Then
            Set rs = ADCDb().OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates") & "] ORDER BY " & SQL_FIELD_ID, dbOpenSnapshot)
            Do While Not rs.EOF
                PumpSetupResponsiveness
                reviewedCount = reviewedCount + 1
                PulseProgressEvery reviewedCount, 25, "Candidate health rows reviewed"
                CandidateRecordFromRecordset rs, candidate
                latestBatchID = LatestCandidateScanBatchID(NormalizedConfigOwnerDatabaseRole(candidate.CandidateOwnerDatabaseRole))
                staleReason = CandidateStaleReason(candidate, latestBatchID)
                reviewReason = CandidateReviewReason(candidate, latestBatchID)
                If Len(staleReason) > 0 Then
                    AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_CANDIDATE, HEALTH_SEVERITY_WARNING, "CANDIDATE_STALE_REVIEW", "CandidateRow", candidate.CandidateMacroKey, candidate.TargetProcedureName, candidate.MacroKey, candidate.MacroTitle, candidate.AcceptedConfigurationID, vbNullString, ROW_KIND_CANDIDATE, vbNullString, candidate.CandidateStatus, candidate.ConfigOwnerDatabaseRole, candidate.CandidateOwnerDatabaseRole, candidate.TargetProcedureName, candidate.TargetProcedureModuleName, vbNullString, vbNullString, vbNullString, "Candidate appears stale. Reason: " & staleReason, "Review and mark stale if appropriate."
                ElseIf Len(reviewReason) > 0 Then
                    AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_CANDIDATE, HEALTH_SEVERITY_WARNING, "CANDIDATE_REVIEW_NEEDED", "CandidateRow", candidate.CandidateMacroKey, candidate.TargetProcedureName, candidate.MacroKey, candidate.MacroTitle, candidate.AcceptedConfigurationID, vbNullString, ROW_KIND_CANDIDATE, vbNullString, candidate.CandidateStatus, candidate.ConfigOwnerDatabaseRole, candidate.CandidateOwnerDatabaseRole, candidate.TargetProcedureName, candidate.TargetProcedureModuleName, vbNullString, vbNullString, vbNullString, "Candidate needs review. Reason: " & reviewReason, "Review candidate metadata before acceptance."
                End If
                If Len(NormalizedConfigOwnerDatabaseRole(candidate.ConfigOwnerDatabaseRole)) = 0 Or Len(NormalizedConfigOwnerDatabaseRole(candidate.CandidateOwnerDatabaseRole)) = 0 Then
                    AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_CANDIDATE, HEALTH_SEVERITY_ERROR, "CANDIDATE_OWNER_ROLE_INVALID", "CandidateRow", candidate.CandidateMacroKey, candidate.TargetProcedureName, candidate.MacroKey, candidate.MacroTitle, candidate.AcceptedConfigurationID, vbNullString, ROW_KIND_CANDIDATE, vbNullString, candidate.CandidateStatus, candidate.ConfigOwnerDatabaseRole, candidate.CandidateOwnerDatabaseRole, candidate.TargetProcedureName, candidate.TargetProcedureModuleName, vbNullString, vbNullString, vbNullString, "Candidate has blank or invalid database-role metadata.", "Correct the candidate database-role metadata before acceptance."
                End If
                If candidate.SelectedForAcceptance And Len(staleReason & reviewReason) > 0 Then
                    AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_CANDIDATE, HEALTH_SEVERITY_WARNING, "CANDIDATE_ACCEPT_SELECTED_NEEDS_REVIEW", "CandidateRow", candidate.CandidateMacroKey, candidate.TargetProcedureName, candidate.MacroKey, candidate.MacroTitle, candidate.AcceptedConfigurationID, vbNullString, ROW_KIND_CANDIDATE, vbNullString, candidate.CandidateStatus, candidate.ConfigOwnerDatabaseRole, candidate.CandidateOwnerDatabaseRole, candidate.TargetProcedureName, candidate.TargetProcedureModuleName, vbNullString, vbNullString, vbNullString, "SelectedForAcceptance=True but this candidate is stale or needs review.", "Clear SelectedForAcceptance or resolve the candidate issue first."
                End If
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing
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
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
    AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_CONFIG, HEALTH_SEVERITY_ERROR, "CONFIG_CANDIDATE_REVIEW_FAILED", ADCAT_OBJECT_TYPE_TABLE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates"), , , , , , , , , , , , , "Configuration/candidate review failed: " & savedErrDescription, "Inspect MC_Records and MC_Candidates manually."
    operationalErrorText = SharedText_AppendNonBlankLine(operationalErrorText, "Configuration/candidate health review failed. Error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription)
End Sub

Private Sub AddHealthSchemaChecks(ByVal rsHealth As DAO.Recordset, ByRef issueHashes As Object, ByVal reviewRunID As String, ByVal reviewRunOn As Date)
    Dim fieldDefinition As Variant
    Dim fieldName As Variant
    Dim fieldNames As Variant
    Dim relationshipFinding As Variant
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: relationshipFindings holds planned/recovery work or independently observed validation evidence for AddHealthSchemaChecks; keep it distinct from the source
    ' authority.
    Dim relationshipFindings As New Collection
    AddCanonicalRelationshipFindings relationshipFindings
    For Each relationshipFinding In relationshipFindings
        PumpSetupResponsiveness
        AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_SCHEMA, HEALTH_SEVERITY_ERROR, "SCHEMA_RELATIONSHIP_INVALID", "Relationship", Left$(CStr(relationshipFinding), 240), , , , , , , , , , , , , , "canonical CodeDb RI relationship", "missing, malformed, cascading, or wrong database scope", CStr(relationshipFinding), "Run Access Development Console: Setup/Update to repair the canonical relationship from the shared relationship registry."
    Next relationshipFinding
    If CodeDbTableExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")) = False Then
        AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_SCHEMA, HEALTH_SEVERITY_CRITICAL, "SCHEMA_CONFIG_TABLE_MISSING", ADCAT_OBJECT_TYPE_TABLE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config"), , , , , , , , , , , , , "Configuration table [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "] is missing.", "Run Access Development Console: Setup/Update or repair tables."
    Else
        fieldNames = Split(ConfigFieldOrderText(), "|")
        For Each fieldName In fieldNames
            PumpSetupResponsiveness
            If TableFieldExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config"), CStr(fieldName)) = False Then
                AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_SCHEMA, HEALTH_SEVERITY_CRITICAL, "SCHEMA_MISSING_FIELD", ADCAT_OBJECT_TYPE_TABLE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config"), , , , , , , , , , , CStr(fieldName), , "Required config field is missing.", "Run Access Development Console: Setup/Update or repair tables."
            End If
        Next fieldName
    End If
    If CodeDbTableExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates")) = False Then
        AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_SCHEMA, HEALTH_SEVERITY_CRITICAL, "SCHEMA_CANDIDATE_TABLE_MISSING", ADCAT_OBJECT_TYPE_TABLE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates"), , , , , , , , , , , , , "Candidate/discovery table is missing.", "Run Access Development Console: Setup/Update or repair tables."
    Else
        fieldNames = Split(CandidateFieldOrderText(), "|")
        For Each fieldName In fieldNames
            PumpSetupResponsiveness
            If TableFieldExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates"), CStr(fieldName)) = False Then
                AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_SCHEMA, HEALTH_SEVERITY_CRITICAL, "SCHEMA_MISSING_FIELD", ADCAT_OBJECT_TYPE_TABLE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates"), , , , , , , , , , , CStr(fieldName), , "Required candidate field is missing.", "Run Access Development Console: Setup/Update or repair tables."
            End If
        Next fieldName
    End If
    If CodeDbTableExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID)) = False Then
        AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_SCHEMA, HEALTH_SEVERITY_CRITICAL, "SCHEMA_SETTINGS_TABLE_MISSING", ADCAT_OBJECT_TYPE_TABLE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID), , , , , , , , , , , , , "Settings table is missing.", "Run Access Development Console: Setup/Update or repair tables."
    Else
        For Each fieldDefinition In SettingsFieldSchemaDefinitions()
            PumpSetupResponsiveness
            If TableFieldExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID), CStr(fieldDefinition(0))) = False Then
                AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_SCHEMA, HEALTH_SEVERITY_CRITICAL, "SCHEMA_MISSING_FIELD", ADCAT_OBJECT_TYPE_TABLE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID), , , , , , , , , , , CStr(fieldDefinition(0)), , "Required settings field is missing.", "Run Access Development Console: Setup/Update or repair tables."
            End If
        Next fieldDefinition
    End If
    If CodeDbTableExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReview")) = False Then
        AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_SCHEMA, HEALTH_SEVERITY_CRITICAL, "SCHEMA_HEALTH_TABLE_MISSING", ADCAT_OBJECT_TYPE_TABLE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReview"), , , , , , , , , , , , , "Health-review table is missing.", "Run Access Development Console: Setup/Update."
    Else
        For Each fieldName In HealthReviewFieldNames()
            PumpSetupResponsiveness
            If TableFieldExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReview"), CStr(fieldName)) = False Then
                AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_SCHEMA, HEALTH_SEVERITY_ERROR, "SCHEMA_MISSING_FIELD", ADCAT_OBJECT_TYPE_TABLE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReview"), , , , , , , , , , , CStr(fieldName), , "Required health-review field is missing.", "Run Access Development Console: Setup/Update."
            End If
        Next fieldName
    End If
End Sub

Private Sub AddHealthOneConfigRowChecks(ByVal rsHealth As DAO.Recordset, ByRef issueHashes As Object, ByVal reviewRunID As String, ByVal reviewRunOn As Date, ByRef cfg As TMacroRecord, ByVal schemaUpdatedOn As Date)
    Dim targetRole As String
    Dim generatedName As String
    targetRole = NormalizedConfigOwnerDatabaseRole(cfg.ConfigOwnerDatabaseRole)
    generatedName = x__UTIL_ADC_MACRO_RUNTIME.ExpectedGeneratedMacroObjectNameForConfig(cfg)
    If cfg.Active And Len(targetRole) = 0 Then
        AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_TARGET, HEALTH_SEVERITY_ERROR, "CONFIG_TARGET_ROLE_INVALID", "ConfigRow", x__UTIL_ADC_MACRO_RUNTIME.ExpectedGeneratedMacroObjectNameForConfig(cfg), cfg.TargetProcedureName, cfg.MacroKey, cfg.MacroTitle, cfg.ID, vbNullString, ROW_KIND_CONFIGURATION, cfg.ConfigurationStatus, vbNullString, cfg.ConfigOwnerDatabaseRole, vbNullString, cfg.TargetProcedureName, cfg.TargetProcedureModuleName, cfg.GeneratedMacroObjectName, vbNullString, vbNullString, "Active config row has blank or invalid ConfigOwnerDatabaseRole.", "Run Resolve Access Development Console Targets or set ConfigOwnerDatabaseRole manually."
    End If
    If cfg.Active And Len(Trim$(cfg.MacroKey)) = 0 Then
        AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_CONFIG, HEALTH_SEVERITY_ERROR, "CONFIG_MACROKEY_BLANK", "ConfigRow", x__UTIL_ADC_MACRO_RUNTIME.ExpectedGeneratedMacroObjectNameForConfig(cfg), cfg.TargetProcedureName, cfg.MacroKey, cfg.MacroTitle, cfg.ID, vbNullString, ROW_KIND_CONFIGURATION, cfg.ConfigurationStatus, vbNullString, cfg.ConfigOwnerDatabaseRole, vbNullString, cfg.TargetProcedureName, cfg.TargetProcedureModuleName, cfg.GeneratedMacroObjectName, vbNullString, vbNullString, "Active config row has blank MacroKey.", "Populate MacroKey before validation/generation."
    End If
    If cfg.Active And Len(Trim$(cfg.MacroTitle)) = 0 Then
        AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_CONFIG, HEALTH_SEVERITY_ERROR, "CONFIG_MACROTITLE_BLANK", "ConfigRow", x__UTIL_ADC_MACRO_RUNTIME.ExpectedGeneratedMacroObjectNameForConfig(cfg), cfg.TargetProcedureName, cfg.MacroKey, cfg.MacroTitle, cfg.ID, vbNullString, ROW_KIND_CONFIGURATION, cfg.ConfigurationStatus, vbNullString, cfg.ConfigOwnerDatabaseRole, vbNullString, cfg.TargetProcedureName, cfg.TargetProcedureModuleName, cfg.GeneratedMacroObjectName, vbNullString, vbNullString, "Active config row has blank MacroTitle.", "Populate MacroTitle before generation."
    End If
    Select Case LCase$(Trim$(cfg.MacroAction))
        Case LCase$(ACTION_RUNCODE)
            If Len(Trim$(cfg.TargetProcedureName)) = 0 Then
                AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_TARGET, HEALTH_SEVERITY_ERROR, "TARGET_PROCEDURE_MISSING", "ConfigRow", x__UTIL_ADC_MACRO_RUNTIME.ExpectedGeneratedMacroObjectNameForConfig(cfg), cfg.TargetProcedureName, cfg.MacroKey, cfg.MacroTitle, cfg.ID, vbNullString, ROW_KIND_CONFIGURATION, cfg.ConfigurationStatus, vbNullString, cfg.ConfigOwnerDatabaseRole, vbNullString, cfg.TargetProcedureName, cfg.TargetProcedureModuleName, cfg.GeneratedMacroObjectName, vbNullString, vbNullString, "RunCode row has blank TargetProcedureName.", "Populate TargetProcedureName."
            End If
        Case LCase$(ACTION_OPENFORM), LCase$(ACTION_OPENTABLE), LCase$(ACTION_OPENQUERY), LCase$(ACTION_OPENREPORT), LCase$(ACTION_RUNMACRO)
            If Len(x__UTIL_ADC_MACRO_RUNTIME.ConfiguredObjectActionTargetName(cfg)) = 0 Then
                AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_TARGET, HEALTH_SEVERITY_ERROR, "TARGET_OBJECT_MISSING", "ConfigRow", x__UTIL_ADC_MACRO_RUNTIME.ExpectedGeneratedMacroObjectNameForConfig(cfg), cfg.TargetProcedureName, cfg.MacroKey, cfg.MacroTitle, cfg.ID, vbNullString, ROW_KIND_CONFIGURATION, cfg.ConfigurationStatus, vbNullString, cfg.ConfigOwnerDatabaseRole, vbNullString, cfg.TargetProcedureName, cfg.TargetProcedureModuleName, cfg.GeneratedMacroObjectName, vbNullString, vbNullString, cfg.MacroAction & " row has blank target object name.", "Populate TargetObjectName with the target Access object name."
            End If
            If StrComp(targetRole, CONFIG_OWNER_DATABASE_ROLE_CODE_DB, vbTextCompare) = 0 Then
                AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_TARGET, HEALTH_SEVERITY_ERROR, "TARGET_OBJECTACTION_CODEDB", "ConfigRow", x__UTIL_ADC_MACRO_RUNTIME.ExpectedGeneratedMacroObjectNameForConfig(cfg), cfg.TargetProcedureName, cfg.MacroKey, cfg.MacroTitle, cfg.ID, vbNullString, ROW_KIND_CONFIGURATION, cfg.ConfigurationStatus, vbNullString, cfg.ConfigOwnerDatabaseRole, vbNullString, cfg.TargetProcedureName, cfg.TargetProcedureModuleName, cfg.GeneratedMacroObjectName, vbNullString, vbNullString, cfg.MacroAction & " + ConfigOwnerDatabaseRole=CodeDb is unsupported.", "Use a HostDb-owned target object or set ConfigOwnerDatabaseRole=HostDb."
            End If
        Case Else
            AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_CONFIG, HEALTH_SEVERITY_ERROR, "CONFIG_MACROACTION_INVALID", "ConfigRow", x__UTIL_ADC_MACRO_RUNTIME.ExpectedGeneratedMacroObjectNameForConfig(cfg), cfg.TargetProcedureName, cfg.MacroKey, cfg.MacroTitle, cfg.ID, vbNullString, ROW_KIND_CONFIGURATION, cfg.ConfigurationStatus, vbNullString, cfg.ConfigOwnerDatabaseRole, vbNullString, cfg.TargetProcedureName, cfg.TargetProcedureModuleName, cfg.GeneratedMacroObjectName, vbNullString, vbNullString, "MacroAction must be RunCode, OpenForm, OpenTable, OpenQuery, OpenReport, or RunMacro.", "Correct MacroAction."
    End Select
    If cfg.LastValidatedOn = 0 Then
        AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_VALIDATION, HEALTH_SEVERITY_WARNING, "VALIDATION_STALE", "ConfigRow", x__UTIL_ADC_MACRO_RUNTIME.ExpectedGeneratedMacroObjectNameForConfig(cfg), cfg.TargetProcedureName, cfg.MacroKey, cfg.MacroTitle, cfg.ID, vbNullString, ROW_KIND_CONFIGURATION, cfg.ConfigurationStatus, vbNullString, cfg.ConfigOwnerDatabaseRole, vbNullString, cfg.TargetProcedureName, cfg.TargetProcedureModuleName, cfg.GeneratedMacroObjectName, vbNullString, vbNullString, "LastValidatedOn is blank.", "Run validation or target resolution."
    ElseIf schemaUpdatedOn <> 0 And cfg.LastValidatedOn < schemaUpdatedOn Then
        AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_VALIDATION, HEALTH_SEVERITY_WARNING, "VALIDATION_STALE", "ConfigRow", x__UTIL_ADC_MACRO_RUNTIME.ExpectedGeneratedMacroObjectNameForConfig(cfg), cfg.TargetProcedureName, cfg.MacroKey, cfg.MacroTitle, cfg.ID, vbNullString, ROW_KIND_CONFIGURATION, cfg.ConfigurationStatus, vbNullString, cfg.ConfigOwnerDatabaseRole, vbNullString, cfg.TargetProcedureName, cfg.TargetProcedureModuleName, cfg.GeneratedMacroObjectName, vbNullString, vbNullString, "LastValidatedOn is older than ADCSchemaUpdatedOn.", "Run validation or target resolution."
    End If
    If cfg.LastGeneratedOn = 0 Then
        AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_GENERATION, HEALTH_SEVERITY_INFO, "GENERATION_STALE", "ConfigRow", x__UTIL_ADC_MACRO_RUNTIME.ExpectedGeneratedMacroObjectNameForConfig(cfg), cfg.TargetProcedureName, cfg.MacroKey, cfg.MacroTitle, cfg.ID, vbNullString, ROW_KIND_CONFIGURATION, cfg.ConfigurationStatus, vbNullString, cfg.ConfigOwnerDatabaseRole, vbNullString, cfg.TargetProcedureName, cfg.TargetProcedureModuleName, cfg.GeneratedMacroObjectName, vbNullString, vbNullString, "LastGeneratedOn is blank.", "Generate when ready."
    ElseIf schemaUpdatedOn <> 0 And cfg.LastGeneratedOn < schemaUpdatedOn Then
        AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_GENERATION, HEALTH_SEVERITY_WARNING, "GENERATION_STALE", "ConfigRow", x__UTIL_ADC_MACRO_RUNTIME.ExpectedGeneratedMacroObjectNameForConfig(cfg), cfg.TargetProcedureName, cfg.MacroKey, cfg.MacroTitle, cfg.ID, vbNullString, ROW_KIND_CONFIGURATION, cfg.ConfigurationStatus, vbNullString, cfg.ConfigOwnerDatabaseRole, vbNullString, cfg.TargetProcedureName, cfg.TargetProcedureModuleName, cfg.GeneratedMacroObjectName, vbNullString, vbNullString, "LastGeneratedOn is older than ADCSchemaUpdatedOn.", "Regenerate when validation is clean."
    End If
    If ResultIndicatesProblem(cfg.LastValidationResult) Then
        AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_VALIDATION, HEALTH_SEVERITY_ERROR, "VALIDATION_RESULT_PROBLEM", "ConfigRow", x__UTIL_ADC_MACRO_RUNTIME.ExpectedGeneratedMacroObjectNameForConfig(cfg), cfg.TargetProcedureName, cfg.MacroKey, cfg.MacroTitle, cfg.ID, vbNullString, ROW_KIND_CONFIGURATION, cfg.ConfigurationStatus, vbNullString, cfg.ConfigOwnerDatabaseRole, vbNullString, cfg.TargetProcedureName, cfg.TargetProcedureModuleName, cfg.GeneratedMacroObjectName, vbNullString, vbNullString, "LastValidationResult indicates a problem: " & cfg.LastValidationResult, "Resolve validation issues before generation."
    End If
    If GeneratedResultIndicatesProblem(cfg.LastGeneratedResult) Then
        AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_GENERATION, HEALTH_SEVERITY_ERROR, "GENERATION_RESULT_PROBLEM", "ConfigRow", x__UTIL_ADC_MACRO_RUNTIME.ExpectedGeneratedMacroObjectNameForConfig(cfg), cfg.TargetProcedureName, cfg.MacroKey, cfg.MacroTitle, cfg.ID, vbNullString, ROW_KIND_CONFIGURATION, cfg.ConfigurationStatus, vbNullString, cfg.ConfigOwnerDatabaseRole, vbNullString, cfg.TargetProcedureName, cfg.TargetProcedureModuleName, cfg.GeneratedMacroObjectName, vbNullString, vbNullString, "LastGeneratedResult indicates a problem: " & cfg.LastGeneratedResult, "Review generated macro status and regenerate when safe."
    End If
    If cfg.LastGeneratedOn <> 0 And Len(generatedName) = 0 Then
        AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_GENERATED_MACRO, HEALTH_SEVERITY_ERROR, "GENERATED_MACRO_NAME_BLANK", "ConfigRow", x__UTIL_ADC_MACRO_RUNTIME.ExpectedGeneratedMacroObjectNameForConfig(cfg), cfg.TargetProcedureName, cfg.MacroKey, cfg.MacroTitle, cfg.ID, vbNullString, ROW_KIND_CONFIGURATION, cfg.ConfigurationStatus, vbNullString, cfg.ConfigOwnerDatabaseRole, vbNullString, cfg.TargetProcedureName, cfg.TargetProcedureModuleName, cfg.GeneratedMacroObjectName, vbNullString, vbNullString, "Generated macro object name is blank after generation.", "Regenerate or repair the config row."
    End If
End Sub

Private Sub AddHealthDuplicateChecks(ByVal rsHealth As DAO.Recordset, ByRef issueHashes As Object, ByVal reviewRunID As String, ByVal reviewRunOn As Date)
    AddDuplicateFieldHealthChecks rsHealth, issueHashes, reviewRunID, reviewRunOn, FIELD_MACRO_KEY, "DUPLICATE_MACROKEY", True
    AddDuplicateFieldHealthChecks rsHealth, issueHashes, reviewRunID, reviewRunOn, FIELD_MACRO_TITLE, "DUPLICATE_ACTIVE_MACROTITLE", True
    AddDuplicateFieldHealthChecks rsHealth, issueHashes, reviewRunID, reviewRunOn, FIELD_GENERATED_MACRO_OBJECT_NAME, "DUPLICATE_GENERATED_MACRO_OBJECT", True
    AddDuplicateTargetSignatureHealthChecks rsHealth, issueHashes, reviewRunID, reviewRunOn
End Sub

Private Function HealthTargetSignatureDisplay(ByRef cfg As TMacroRecord) As String
    If StrComp(Trim$(cfg.MacroAction), ACTION_RUNCODE, vbTextCompare) = 0 Then
        HealthTargetSignatureDisplay = NormalizedConfigOwnerDatabaseRole(cfg.ConfigOwnerDatabaseRole) & " | " & cfg.MacroAction & " | " & Trim$(cfg.TargetProcedureName) & " | " & cfg.TargetArguments
    Else
        HealthTargetSignatureDisplay = NormalizedConfigOwnerDatabaseRole(cfg.ConfigOwnerDatabaseRole) & " | " & cfg.MacroAction & " | " & Trim$(cfg.TargetObjectName) & " | " & Trim$(cfg.ConfirmationMode)
    End If
End Function

Private Sub AddHealthPublicEntryPointChecks(ByVal rsHealth As DAO.Recordset, ByRef issueHashes As Object, ByVal reviewRunID As String, ByVal reviewRunOn As Date)
    Dim detailText As String
    Dim procName As Variant
    Dim procedureKind As String
    Dim procedureLocation As String
    Dim ignoredArgumentSignature As String
    Dim ignoredRequiredArgumentCount As Long
    Dim ignoredOptionalArgumentCount As Long
    Dim ignoredHasParamArray As Boolean
    Dim ignoredExpectedReturnType As String
    If x__UTIL_ADC_MACRO_RUNTIME.CanAccessVBEProjectModel(detailText) = False Then
        AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_PUBLIC_ENTRY_POINT, HEALTH_SEVERITY_WARNING, "VBE_ACCESS_UNAVAILABLE", "VBE", "Application.VBE", , , , , , , , , , , "VBE accessible", "unavailable", detailText, "Trust access to the VBA project object model if public entry-point audit is needed."
        Exit Sub
    End If
    For Each procName In ExpectedPublicEntryPointNames()
        PumpSetupResponsiveness
        procedureKind = vbNullString
        procedureLocation = vbNullString
        If x__UTIL_ADC_MACRO_RUNTIME.PublicCallableProcedureInfoInStandardModule(CStr(procName), CONFIG_OWNER_DATABASE_ROLE_CODE_DB, procedureKind, procedureLocation, ignoredArgumentSignature, ignoredRequiredArgumentCount, ignoredOptionalArgumentCount, ignoredHasParamArray, ignoredExpectedReturnType) = False Then
            AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_PUBLIC_ENTRY_POINT, HEALTH_SEVERITY_ERROR, "PUBLIC_ENTRY_MISSING", "Procedure", CStr(procName), CStr(procName), , , , , , , , , , "public callable procedure", "missing or private", "Expected public entry point is not callable from a standard module.", "Verify the wrapper is Public and not hidden by Option Private Module."
        End If
    Next procName
End Sub

Private Function HealthFindingCountBySeverity(ByVal reviewRunID As String, ByVal severityText As String) As Long
    'Counts health findings in the explicit CodeDb context; database failures are not converted to a valid zero count.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Set db = codeDb
    sqlText = "SELECT Count(*) AS [FindingCount] FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReview") & "] WHERE " & SQL_FIELD_REVIEW_RUN_ID & "=" & SharedDao_SqlQuote(reviewRunID) & " AND " & SQL_FIELD_SEVERITY & "=" & SharedDao_SqlQuote(severityText) & ";"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    HealthFindingCountBySeverity = CLng(rs!findingCount)
    rs.Close
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
    Err.Raise savedErrNumber, "HealthFindingCountBySeverity", "Could not count CodeDb health-review findings." & vbCrLf & "Review run ID: " & reviewRunID & vbCrLf & "Severity: " & severityText & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Sub AddHealthFinding(ByVal rsHealth As DAO.Recordset, ByRef issueHashes As Object, ByVal reviewRunID As String, ByVal reviewRunOn As Date, ByVal reviewArea As String, ByVal severityText As String, ByVal issueCode As String, ParamArray findingValues() As Variant)
    Dim issueHash As String
    Dim objectType As String
    Dim objectName As String
    Dim ProcedureName As String
    Dim MacroKey As String
    Dim MacroTitle As String
    Dim configID As Long
    Dim SettingKey As String
    Dim RowKind As String
    Dim ConfigurationStatus As String
    Dim CandidateStatus As String
    Dim ConfigOwnerDatabaseRole As String
    Dim CandidateOwnerDatabaseRole As String
    Dim TargetProcedureName As String
    Dim TargetProcedureModuleName As String
    Dim GeneratedMacroObjectName As String
    Dim expectedValue As String
    Dim actualValue As String
    Dim detailsText As String
    Dim recommendedActionText As String
    Dim findingUpperBound As Long
    findingUpperBound = -1
    On Error Resume Next
    findingUpperBound = UBound(findingValues)
    If Err.Number <> 0 Then
        Err.Clear
        findingUpperBound = -1
    End If
    On Error GoTo 0
    If findingUpperBound >= 0 Then objectType = HealthArgText(findingValues(0))
    If findingUpperBound >= 1 Then objectName = HealthArgText(findingValues(1))
    If findingUpperBound >= 2 Then ProcedureName = HealthArgText(findingValues(2))
    If findingUpperBound >= 3 Then MacroKey = HealthArgText(findingValues(3))
    If findingUpperBound >= 4 Then MacroTitle = HealthArgText(findingValues(4))
    If findingUpperBound >= 5 Then configID = HealthArgLong(findingValues(5), 0)
    If findingUpperBound >= 6 Then SettingKey = HealthArgText(findingValues(6))
    If findingUpperBound >= 7 Then RowKind = HealthArgText(findingValues(7))
    If findingUpperBound >= 8 Then ConfigurationStatus = HealthArgText(findingValues(8))
    If findingUpperBound >= 9 Then CandidateStatus = HealthArgText(findingValues(9))
    If findingUpperBound >= 10 Then ConfigOwnerDatabaseRole = HealthArgText(findingValues(10))
    If findingUpperBound >= 11 Then CandidateOwnerDatabaseRole = HealthArgText(findingValues(11))
    If findingUpperBound >= 12 Then TargetProcedureName = HealthArgText(findingValues(12))
    If findingUpperBound >= 13 Then TargetProcedureModuleName = HealthArgText(findingValues(13))
    If findingUpperBound >= 14 Then GeneratedMacroObjectName = HealthArgText(findingValues(14))
    If findingUpperBound >= 15 Then expectedValue = HealthArgText(findingValues(15))
    If findingUpperBound >= 16 Then actualValue = HealthArgText(findingValues(16))
    If findingUpperBound >= 17 Then detailsText = HealthArgText(findingValues(17))
    If findingUpperBound >= 18 Then recommendedActionText = HealthArgText(findingValues(18))
    issueHash = StableTextHashHex(LCase$(reviewArea) & MODULE_TAG_DELIMITER & LCase$(issueCode) & MODULE_TAG_DELIMITER & LCase$(objectType) & MODULE_TAG_DELIMITER & LCase$(objectName) & MODULE_TAG_DELIMITER & LCase$(ProcedureName) & MODULE_TAG_DELIMITER & LCase$(MacroKey) & MODULE_TAG_DELIMITER & CStr(configID) & MODULE_TAG_DELIMITER & LCase$(SettingKey) & MODULE_TAG_DELIMITER & LCase$(ConfigOwnerDatabaseRole) & MODULE_TAG_DELIMITER & LCase$(CandidateOwnerDatabaseRole))
    If issueHashes.Exists(issueHash) Then Exit Sub
    issueHashes.Add issueHash, True
    rsHealth.AddNew
    SetRecordsetFieldIfExists rsHealth, FIELD_REVIEW_RUN_ID, reviewRunID
    SetRecordsetFieldIfExists rsHealth, FIELD_REVIEW_RUN_ON, reviewRunOn
    SetRecordsetFieldIfExists rsHealth, FIELD_REVIEW_AREA, reviewArea
    SetRecordsetFieldIfExists rsHealth, FIELD_SEVERITY, severityText
    SetRecordsetFieldIfExists rsHealth, FIELD_ISSUE_CODE, issueCode
    SetRecordsetFieldIfExists rsHealth, FIELD_OBJECT_TYPE, objectType
    SetRecordsetFieldIfExists rsHealth, FIELD_OBJECT_NAME, objectName
    SetRecordsetFieldIfExists rsHealth, FIELD_PROCEDURE_NAME, ProcedureName
    SetRecordsetFieldIfExists rsHealth, FIELD_MACRO_KEY, MacroKey
    SetRecordsetFieldIfExists rsHealth, FIELD_MACRO_TITLE, MacroTitle
    If configID > 0 Then SetRecordsetFieldIfExists rsHealth, FIELD_CONFIG_ID, configID
    SetRecordsetFieldIfExists rsHealth, FIELD_SETTING_KEY, SettingKey
    SetRecordsetFieldIfExists rsHealth, FIELD_ROW_KIND, RowKind
    SetRecordsetFieldIfExists rsHealth, FIELD_CONFIGURATION_STATUS, ConfigurationStatus
    SetRecordsetFieldIfExists rsHealth, FIELD_CANDIDATE_STATUS, CandidateStatus
    SetRecordsetFieldIfExists rsHealth, FIELD_CONFIG_OWNER_DATABASE_ROLE, ConfigOwnerDatabaseRole
    SetRecordsetFieldIfExists rsHealth, FIELD_CANDIDATE_OWNER_DATABASE_ROLE, CandidateOwnerDatabaseRole
    SetRecordsetFieldIfExists rsHealth, FIELD_TARGET_PROCEDURE_NAME, TargetProcedureName
    SetRecordsetFieldIfExists rsHealth, FIELD_TARGET_PROCEDURE_MODULE_NAME, TargetProcedureModuleName
    SetRecordsetFieldIfExists rsHealth, FIELD_GENERATED_MACRO_OBJECT_NAME, GeneratedMacroObjectName
    SetRecordsetFieldIfExists rsHealth, FIELD_EXPECTED_VALUE, expectedValue
    SetRecordsetFieldIfExists rsHealth, FIELD_ACTUAL_VALUE, actualValue
    SetRecordsetFieldIfExists rsHealth, FIELD_DETAILS, detailsText
    SetRecordsetFieldIfExists rsHealth, FIELD_RECOMMENDED_ACTION, recommendedActionText
    SetRecordsetFieldIfExists rsHealth, FIELD_ISSUE_HASH, issueHash
    SetRecordsetFieldIfExists rsHealth, FIELD_IS_CURRENT_RUN, True
    SetRecordsetFieldIfExists rsHealth, FIELD_CREATED_ON, Now
    rsHealth.Update
End Sub

Private Function HealthArgLong(ByVal argValue As Variant, Optional ByVal DefaultValue As Long = 0) As Long
    On Error GoTo Fallback
    HealthArgLong = DefaultValue
    If IsError(argValue) Then Exit Function
    If IsNull(argValue) Then Exit Function
    If IsEmpty(argValue) Then Exit Function
    If Len(Trim$(CStr(argValue))) = 0 Then Exit Function
    HealthArgLong = CLng(argValue)
    Exit Function
Fallback:
    HealthArgLong = DefaultValue
End Function

Private Sub AddDuplicateFieldHealthChecks(ByVal rsHealth As DAO.Recordset, ByRef issueHashes As Object, ByVal reviewRunID As String, ByVal reviewRunOn As Date, ByVal fieldName As String, ByVal issueCode As String, ByVal activeOnly As Boolean)
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim valueText As String
    If CodeDbTableExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")) = False Then Exit Sub
    If TableFieldExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config"), fieldName) = False Then Exit Sub
    sqlText = "SELECT [" & fieldName & "] AS DupValue, Count(*) " & "AS " & SQL_FIELD_DUP_COUNT & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "] WHERE Len(Nz([" & fieldName & "],''))>0 "
    If activeOnly Then sqlText = sqlText & "AND Nz(" & SQL_FIELD_ACTIVE & ",False)<>False "
    sqlText = sqlText & "GROUP BY [" & fieldName & "] HAVING Count(*)>1"
    Set rs = ADCDb().OpenRecordset(sqlText, dbOpenSnapshot)
    Do While Not rs.EOF
        PumpSetupResponsiveness
        valueText = CStr(rs!DupValue)
        AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_DUPLICATE, HEALTH_SEVERITY_ERROR, issueCode, "ConfigField", fieldName, , valueText, , , , , , , , , "unique", CStr(rs.fields(FIELD_DUP_COUNT)) & " rows", "Duplicate value found in MC_Records table field [" & fieldName & "].", "Review duplicate rows and keep only the intended active/configured row."
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
    Err.Raise savedErrNumber, "AddDuplicateFieldHealthChecks", "Could not complete duplicate-field health checks." & vbCrLf & "Field: " & fieldName & vbCrLf & "IssueCode: " & issueCode & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub AddHealthManagedDescriptionChecks(ByVal rsHealth As DAO.Recordset, ByRef issueHashes As Object, ByVal reviewRunID As String, ByVal reviewRunOn As Date, ByRef operationalErrorText As String)
    'Independently reports catalog and physical Access Description coverage without repairing managed objects or fields.
    On Error GoTo ReviewFailed
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim tdf As DAO.tableDef
    Dim fld As DAO.Field
    Dim sqlText As String
    Dim objectName As String
    Dim objectTypeCode As String
    Dim descriptionText As String
    Dim expectedExistencePolicy As String
    Dim setupAction As String
    Dim optionalReason As String
    Dim tableName As String
    Dim fieldName As String
    Dim actualDescription As String
    Dim propertyType As Long
    Dim propertyExists As Boolean
    Dim descriptionOptionalForHelpMacro As Boolean
    Dim fieldDescriptionShortByKey As Object
    Dim fieldDescriptionKey As String
    BeginManagedDescriptionOptionalHelpPolicyScope
    Set db = ADCDb()
    Set fieldDescriptionShortByKey = CreateObject("Scripting.Dictionary")
    fieldDescriptionShortByKey.CompareMode = vbTextCompare
    sqlText = "SELECT o." & SQL_FIELD_OBJECT_NAME & ",f.[" & FIELD_MANAGED_OBJECT_FIELD_NAME & "],f.[" & FIELD_MANAGED_OBJECT_FIELD_DESCRIPTION_SHORT & "] FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFields") & "] AS f INNER JOIN [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID) & "] AS o ON f.[" & FIELD_MANAGED_OBJECT_ID & "]=o." & SQL_FIELD_ID & " WHERE Nz(f.[" & FIELD_MANAGED_OBJECT_FIELD_IS_ACTIVE & "],False)<>False AND Nz(o.[" & FIELD_MANAGED_OBJECT_STATUS & "]," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & " AND UCase(Trim(Nz(o." & SQL_FIELD_OBJECT_TYPE & ",'')))='TABLE';"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    Do While rs.EOF = False
        PumpSetupResponsiveness
        tableName = Trim$(CStr(Nz(rs.fields(FIELD_OBJECT_NAME).value, vbNullString)))
        fieldName = Trim$(CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_NAME).value, vbNullString)))
        If Len(tableName) > 0 And Len(fieldName) > 0 Then fieldDescriptionShortByKey(tableName & "|" & fieldName) = CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_DESCRIPTION_SHORT).value, vbNullString))
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    sqlText = "SELECT m." & SQL_FIELD_OBJECT_NAME & ",m." & SQL_FIELD_OBJECT_TYPE & ",m." & SQL_FIELD_OBJECT_DESCRIPTION & _
        ",IIf(Len(Nz(o." & SQL_FIELD_EXPECTED_EXISTENCE_POLICY_OVERRIDE & ",''))>0,o." & SQL_FIELD_EXPECTED_EXISTENCE_POLICY_OVERRIDE & ",c." & SQL_FIELD_DEFAULT_EXPECTED_EXISTENCE_POLICY & ") " & _
        "AS " & SQL_FIELD_RESOLVED_EXPECTED_EXISTENCE_POLICY & ",IIf(Len(Nz(o." & SQL_FIELD_SETUP_ACTION_OVERRIDE & ",''))>0,o." & SQL_FIELD_SETUP_ACTION_OVERRIDE & ",c." & SQL_FIELD_DEFAULT_SETUP_ACTION & ") " & _
        "AS " & SQL_FIELD_RESOLVED_SETUP_ACTION & " FROM ([" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID) & "] AS m " & _
        "INNER JOIN [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ObjectCategories") & "] AS c ON m." & SQL_FIELD_CATEGORY_CODE & "=c." & SQL_FIELD_CATEGORY_CODE & ") LEFT JOIN [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECT_OVERRIDES_TABLE_LOGICAL_ID) & "] AS o ON m." & SQL_FIELD_ID & "=o.[" & FIELD_MANAGED_OBJECT_ID & "]" & _
        " WHERE Nz(m." & SQL_FIELD_MANAGED_OBJECT_STATUS & "," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & _
        " AND Nz(c." & SQL_FIELD_OBJECT_CATEGORY_STATUS & "," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & _
        " ORDER BY m." & SQL_FIELD_OBJECT_TYPE & ",m." & SQL_FIELD_OBJECT_NAME & ";"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    Do While rs.EOF = False
        PumpSetupResponsiveness
        objectName = CStr(Nz(rs.fields(FIELD_OBJECT_NAME).value, vbNullString))
        objectTypeCode = CStr(Nz(rs.fields(FIELD_OBJECT_TYPE).value, vbNullString))
        descriptionText = CStr(Nz(rs.fields(FIELD_OBJECT_DESCRIPTION).value, vbNullString))
        expectedExistencePolicy = CStr(Nz(rs.fields(FIELD_RESOLVED_EXPECTED_EXISTENCE_POLICY).value, vbNullString))
        setupAction = CStr(Nz(rs.fields(FIELD_RESOLVED_SETUP_ACTION).value, vbNullString))
        descriptionOptionalForHelpMacro = ManagedDescriptionIsOptionalHelpMacro(objectName, objectTypeCode)
        If descriptionOptionalForHelpMacro = False And (Len(Trim$(descriptionText)) = 0 Or ManagedObjectDescriptionIsPlaceholder(descriptionText)) Then
            AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_SCHEMA, HEALTH_SEVERITY_ERROR, "MANAGED_OBJECT_DESCRIPTION_UNMAPPED", objectTypeCode, objectName, , , , , , , , , , , , , , "specific nonblank authoritative description", descriptionText, "The active Managed Object Catalog row has no acceptable authoritative description.", "Define the object-specific description in MAC and rerun setup."
        ElseIf ManagedDescriptionTargetSupportedAndPresent(db, objectName, objectTypeCode) = False Then
            optionalReason = ManagedObjectOptionalMissingReason(expectedExistencePolicy, setupAction)
            If Len(optionalReason) = 0 Then AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_SCHEMA, HEALTH_SEVERITY_ERROR, "MANAGED_OBJECT_DESCRIPTION_TARGET_MISSING", objectTypeCode, objectName, , , , , , , , , , , , , , "required managed-object Description target exists", "missing or unsupported", "The active managed object required by the catalog is missing or does not support Access Description metadata.", "Restore or recreate the required object through its authoritative setup owner."
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Set rs = db.OpenRecordset("SELECT DISTINCT " & SQL_FIELD_OBJECT_NAME & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID) & "] WHERE Nz(" & SQL_FIELD_MANAGED_OBJECT_STATUS & ",'Active')='Active' AND UCase(Trim(Nz(" & SQL_FIELD_OBJECT_TYPE & ",'')))='TABLE' ORDER BY " & SQL_FIELD_OBJECT_NAME & ";", dbOpenSnapshot)
    Do While rs.EOF = False
        PumpSetupResponsiveness
        tableName = CStr(Nz(rs.fields(FIELD_OBJECT_NAME).value, vbNullString))
        If Len(Trim$(tableName)) > 0 Then
            If SharedDao_TableExists(db, tableName) = False Then
                AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_SCHEMA, HEALTH_SEVERITY_CRITICAL, "MANAGED_TABLE_DESCRIPTION_TARGET_MISSING", "Table", tableName, , , , , , , , , , , , , , "active managed table exists", "missing", "Field-description coverage cannot be verified because the active managed table is missing.", "Restore or recreate the managed table through setup."
            Else
                Set tdf = db.TableDefs(tableName)
                For Each fld In tdf.fields
                    PumpSetupResponsiveness
                    fieldName = fld.Name
                    fieldDescriptionKey = tdf.Name & "|" & fieldName
                    descriptionText = vbNullString
                    If fieldDescriptionShortByKey.Exists(fieldDescriptionKey) Then descriptionText = CStr(fieldDescriptionShortByKey(fieldDescriptionKey))
                    If Len(Trim$(descriptionText)) = 0 Or FieldDescriptionTextIsPlaceholder(descriptionText) Then
                        AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_SCHEMA, HEALTH_SEVERITY_ERROR, "MANAGED_FIELD_DESCRIPTION_UNMAPPED", "TableField", tableName & "." & fieldName, , , , , , , , , , , , , , "field-specific authoritative description", descriptionText, "No acceptable authoritative FieldDescriptionShort exists for an active managed-table field.", "Populate MAC_Fields.FieldDescriptionShort with the intended <=255-character DAO field description, then rerun setup."
                    Else
                        actualDescription = vbNullString
                        propertyType = 0
                        propertyExists = False
                        ReadDaoFieldDescriptionState fld, actualDescription, propertyType, propertyExists
                        If Len(descriptionText) > 255 Then
                            AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_SCHEMA, HEALTH_SEVERITY_ERROR, "MANAGED_FIELD_DESCRIPTION_SHORT_TOO_LONG", "TableField", tableName & "." & fieldName, , , , , , , , , , , , , , "<=255 characters", CStr(Len(descriptionText)) & " characters", "MAC_Fields.FieldDescriptionShort exceeds the Access DAO Field.Description limit.", "Shorten MAC_Fields.FieldDescriptionShort intentionally, then rerun G1."
                        ElseIf propertyExists = False Then
                            AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_SCHEMA, HEALTH_SEVERITY_ERROR, "MANAGED_FIELD_DESCRIPTION_MISSING", "TableField", tableName & "." & fieldName, , , , , , , , , , , , , , descriptionText, "property missing", "The physical field has no Access Description property.", "Run setup to create and verify the field Description property."
                        ElseIf propertyType <> CLng(dbText) Then
                            AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_SCHEMA, HEALTH_SEVERITY_ERROR, "MANAGED_FIELD_DESCRIPTION_WRONG_TYPE", "TableField", tableName & "." & fieldName, , , , , , , , , , , , , , "DAO dbText (10)", CStr(propertyType), "The physical field Description property has the wrong DAO type.", "Run setup to replace and verify the Description property."
                        ElseIf StrComp(actualDescription, descriptionText, vbBinaryCompare) <> 0 Then
                            AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_SCHEMA, HEALTH_SEVERITY_ERROR, "MANAGED_FIELD_DESCRIPTION_MISMATCH", "TableField", tableName & "." & fieldName, , , , , , , , , , , , , , descriptionText, actualDescription, "The physical field Description differs from the authoritative schema definition.", "Run setup to repair and verify the field Description property."
                        End If
                    End If
                Next fld
                Set tdf = Nothing
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    If SharedDao_TableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFields")) Then
        sqlText = "SELECT o." & SQL_FIELD_OBJECT_NAME & ",f." & SQL_FIELD_MANAGED_OBJECT_FIELD_NAME & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFields") & "] AS f INNER JOIN [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID) & "] AS o ON f." & SQL_FIELD_MANAGED_OBJECT_ID & "=o." & SQL_FIELD_ID & " WHERE Nz(f." & SQL_FIELD_MANAGED_OBJECT_FIELD_IS_ACTIVE & ",False)<>False AND UCase(Trim(Nz(o." & SQL_FIELD_OBJECT_TYPE & ",'')))='TABLE' ORDER BY o." & SQL_FIELD_OBJECT_NAME & ",f." & SQL_FIELD_MANAGED_OBJECT_FIELD_NAME & ";"
        Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
        Do While rs.EOF = False
            PumpSetupResponsiveness
            tableName = CStr(Nz(rs.fields(FIELD_OBJECT_NAME).value, vbNullString))
            fieldName = CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_NAME).value, vbNullString))
            If Len(tableName) > 0 And Len(fieldName) > 0 Then
                If SharedDao_TableExists(db, tableName) Then
                    Set tdf = db.TableDefs(tableName)
                    If SharedDao_TableDefFieldExistsSafe(tdf, fieldName) = False Then AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_SCHEMA, HEALTH_SEVERITY_WARNING, "MANAGED_FIELD_DESCRIPTION_ORPHANED_MAPPING", "TableField", tableName & "." & fieldName, , , , , , , , , , , , , , "active physical field", "mapping exists but field is absent", "An active MAC_Fields mapping no longer resolves to a physical field.", "Mark the stale mapping inactive or restore the field through its authoritative schema/setup owner."
                    Set tdf = Nothing
                End If
            End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
    EndManagedDescriptionOptionalHelpPolicyScope
    Exit Sub
ReviewFailed:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    EndManagedDescriptionOptionalHelpPolicyScope
    On Error GoTo 0
    AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_SCHEMA, HEALTH_SEVERITY_ERROR, "MANAGED_DESCRIPTION_REVIEW_FAILED", "DescriptionCoverage", IIf(Len(tableName) > 0, tableName & IIf(Len(fieldName) > 0, "." & fieldName, vbNullString), objectName), , , , , , , , , , , , , , "description coverage review completes", "error", "Managed-description health review failed. Error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription, "Inspect the trace and rerun setup before relying on description coverage."
    operationalErrorText = SharedText_AppendNonBlankLine(operationalErrorText, "Managed-description health review failed. Error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription)
End Sub

Private Function ManagedDescriptionTargetSupportedAndPresent(ByVal db As DAO.Database, ByVal objectName As String, ByVal objectTypeCode As String) As Boolean
    'Verifies that one active managed object exists and uses a description-supported Access object type without rewriting its Description property.
    objectTypeCode = Trim$(objectTypeCode)
    If Len(objectName) = 0 Then Exit Function
    Select Case UCase$(objectTypeCode)
        Case UCase$(ADCAT_OBJECT_TYPE_MACRO)
            ManagedDescriptionTargetSupportedAndPresent = DatabaseAccessObjectExistsByName(db, acMacro, objectName)
        Case UCase$(ADCAT_OBJECT_TYPE_TABLE)
            ManagedDescriptionTargetSupportedAndPresent = SharedDao_TableExists(db, objectName)
        Case UCase$(ADCAT_OBJECT_TYPE_QUERY)
            ManagedDescriptionTargetSupportedAndPresent = SharedDao_QueryExists(db, objectName)
        Case UCase$(ADCAT_OBJECT_TYPE_FORM)
            ManagedDescriptionTargetSupportedAndPresent = DatabaseAccessObjectExistsByName(db, acForm, objectName)
        Case UCase$(ADCAT_OBJECT_TYPE_REPORT)
            ManagedDescriptionTargetSupportedAndPresent = DatabaseAccessObjectExistsByName(db, acReport, objectName)
    End Select
End Function

Private Function ResultIndicatesProblem(ByVal resultText As String) As Boolean
    Select Case LCase$(Trim$(resultText))
        Case LCase$(VALIDATION_RESULT_INVALID), LCase$(VALIDATION_RESULT_ERROR), LCase$(VALIDATION_RESULT_NEEDS_REVIEW)
            ResultIndicatesProblem = True
    End Select
End Function

Private Sub AddHealthSystemMacroChecks(ByVal rsHealth As DAO.Recordset, ByRef issueHashes As Object, ByVal reviewRunID As String, ByVal reviewRunOn As Date)
    Dim catalogRows As Collection
    Dim rowData As Variant
    Dim TitleText As String
    Set catalogRows = SystemMacroCatalog()
    For Each rowData In catalogRows
        PumpSetupResponsiveness
        TitleText = SystemMacroCatalogTitle(rowData)
        If AccessObjectExistsInCurrentDatabase(acMacro, TitleText) = False Then
            AddHealthFinding rsHealth, issueHashes, reviewRunID, reviewRunOn, REVIEW_AREA_SYSTEM_MACRO, HEALTH_SEVERITY_ERROR, "SYSTEM_MACRO_MISSING", ADCAT_OBJECT_TYPE_MACRO, TitleText, , , , , , , , , , , "macro exists", "missing", "Expected built-in Access Development Console system macro is missing.", "Run setup/update to recreate system macros."
        End If
    Next rowData
End Sub

Private Function ExpectedPublicEntryPointNames() As Variant
    Dim catalogRows As Collection
    Dim values As Collection
    Dim rowData As Variant
    Dim result() As String
    Dim rowIndex As Long
    Dim procName As String
    Dim runCodeName As String
    ' INTENTIONAL FILTERED PARTITION: values is a consumer-required subset/projection derived from the source enumerated by ExpectedPublicEntryPointNames; its membership is not independently authoritative.
    Set values = New Collection
    Set catalogRows = SystemMacroCatalog()
    For Each rowData In catalogRows
        PumpSetupResponsiveness
        runCodeName = SystemMacroCatalogRunCodeFunctionName(rowData)
        procName = ProcedureNameFromRunCodeExpression(runCodeName)
        If Len(procName) = 0 Then procName = runCodeName
        If Len(Trim$(procName)) > 0 Then
            If SharedCollection_KeyExists(values, LCase$(procName)) = False Then values.Add procName, LCase$(procName)
        End If
    Next rowData
    If SharedCollection_KeyExists(values, LCase$("RunConfiguredMacroTarget")) = False Then values.Add "RunConfiguredMacroTarget", LCase$("RunConfiguredMacroTarget")
    If SharedCollection_KeyExists(values, LCase$("RunConfirmedConfiguredMacroTarget")) = False Then values.Add "RunConfirmedConfiguredMacroTarget", LCase$("RunConfirmedConfiguredMacroTarget")
    ReDim result(0 To values.count - 1)
    For rowIndex = 1 To values.count
        PumpSetupResponsiveness
        result(rowIndex - 1) = CStr(values(rowIndex))
    Next rowIndex
    ExpectedPublicEntryPointNames = result
End Function

Private Function IsBackupAccessObjectName(ByVal objectName As String) As Boolean
    Dim normalizedName As String
    normalizedName = UCase$(objectName)
    If Len(normalizedName) = 0 Then Exit Function
    IsBackupAccessObjectName = (Left$(normalizedName, Len("ADCBAK_")) = "ADCBAK_")
End Function
