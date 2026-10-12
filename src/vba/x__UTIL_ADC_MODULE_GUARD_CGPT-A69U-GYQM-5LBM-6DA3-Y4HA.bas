Attribute VB_Name = "x__UTIL_ADC_MODULE_GUARD"
' ================================================================================================================================================================================
' ChatGPT Rewrite Tracking Header
' ================================================================================================================================================================================
' ChatGPT Previous Version ID: CGPT-HY67-3UY9-NQHX-CL3L-MM36
' ChatGPT Baseline Version ID: CGPT-HY67-3UY9-NQHX-CL3L-MM36
' ChatGPT Version ID: CGPT-A69U-GYQM-5LBM-6DA3-Y4HA
' ChatGPT Version Timestamp: Saturday 8/29/26 10:33 AM
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
' STANDARD MODULE: Access Development Console Module Update Guard
' VB_Name: x__UTIL_ADC_MODULE_GUARD
' Entry Points:
'     PrepareAndVerifyADCModuleUpdateGuard
'     SetModuleUpdateGuardEnabledForSetup
' SHARED CODE:
'     SharedDao_DatabaseTableExistsSafe(ByVal db As DAO.Database, ByVal tableName As String) -> Boolean
'     SharedDao_DatabaseTableFieldExistsSafe(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String) -> Boolean
'     SharedDao_SqlQuote(ByVal value As Variant) -> String
'     Shared_String_ConvertToNullIfBlank(ByVal v As Variant) -> Variant
' CLASSES USED: None
' DYNAMIC STRING REFERENCES:
'     Dynamic-reference scan stopped after 2000 logical line(s) to prevent a long-running scan; review this module manually if needed.
' LastScanTimestamp: Saturday 8/15/26 8:45 AM
' Purpose:
'     Owns ADC Module Guard storage, verification, coordinated-build registration, and Setup Preflight policy.
'     Change-prone approved-build identity and roster definitions are owned by the companion definitions module.
' Author: Christina Shonts
' Support: Access Development Console CodeDb module-version assurance and copy/paste overwrite detection.
' Code Category: Access Development Console update assurance and preflight verification.
' Requirements:
'     Microsoft Access, DAO, trusted access to the VBA project object model, and the companion Access Development Console modules.
' Reference Dependencies: Access object model and DAO. Scripting.Dictionary is created late-bound.
' DB Object Dependencies:
'     ModuleVersion_History_Updates, ModuleVersion_Manifest, ModuleVersion_History_Verifications, and ModuleUpdates_* saved queries.
' Usage:
'     Setup and Repair call PrepareAndVerifyADCModuleUpdateGuard before material work. Ordinary preflight first verifies the database-local active manifest. When that manifest differs, preflight
'     invokes the approved-build registration authority with structure repair enabled, which reads the complete coordinated identity from each guarded module's managed rewrite-header comments and challenges it
'     against persistent manifest/history state. A nonmatching installed set returns failure for the explicit setup choice. When the user chooses Yes, the same registrar switches to explicit-approval mode and
'     requires the exact module identities in the fixed guarded roster. It validates their types and managed headers; rejects missing, extra, invalid, duplicate, or conflicting exact-version history identities;
'     explicitly overrides active-manifest lineage/timestamp transition rejection, then transactionally registers the exact installed Version IDs and timestamps.
'     A2 owns insertion of the missing EnableModuleUpdateGuard row from the canonical
'     GetCoreSettingDefinitions contract while preserving any existing value. I1 reads final Guard infrastructure/manifest/enable-setting state without registration or repair. Future guarded rewrites
'     update module tracking headers. Every coordinated guarded build also advances the definitions module and its approved-build batch identity; Module Guard reads module identities from managed headers.
' Version:
'     Splits changing approved-build identity and guarded-roster definitions into x__UTIL_ADC_MODULE_GUARD_DEFINITIONS while keeping validation/registration policy in this module.
'     Requires that definitions module and its batch identity advance whenever another guarded module changes; removes per-module change descriptions from the active build-definition contract.
'     Restores explicit user-approved registration as a deliberate runtime override of active-manifest lineage/timestamp rejection while keeping automatic Module Guard validation fail-closed.
'     Removes executable hard-coded module Version IDs and source-carried transition exceptions; automatic stale baseline/branch mismatches remain fail-closed against persistent Module Guard history.
'     Synchronizes the authoritative supplied-baseline lineage correction and DB Inventory VBA QueryDef compile fix; Module Guard rollback detection and rejection behavior are unchanged.
'     Synchronizes the coordinated build for A2 Preflight HealthReview/launcher-review schema detection, backup coverage, and final-proof correction; the guarded roster and Setup-phase ownership are unchanged.
'     Synchronizes Module Guard for the TABLES-owned MC_Records.MacroKey required/nonblank/unique schema correction; guarded roster and Setup-phase ownership are unchanged.
'     Synchronizes the lossless MAC field/value identity migration and updated repair help contract; guarded roster and Setup-phase ownership remain unchanged.
'     Synchronizes Module Guard for the B1 structural-family proof gate and B2 error-propagation correction; the guarded roster and Setup-phase ownership are unchanged.
'     Synchronizes Module Guard for the B1 MAC_Fields/MAC_Fields_Vals identity-hardening update while preserving the guarded roster and Setup-phase ownership.
'     Synchronizes Module Guard for TABLES-owned canonical relationship timing diagnostics; guarded behavior, Setup-phase ownership, and the guarded-module roster are unchanged.
'     Synchronizes Module Guard for the TABLES physical-file-only startup trace writer and Ribbon diagnostic routing; Setup-phase ownership and the guarded-module roster are unchanged.
'     Synchronizes Module Guard for Delete All Objects terminal diagnostics in Core/MOC Lifecycle/Ribbon startup; Setup-phase ownership and the guarded-module roster are unchanged.
'     Synchronizes Module Guard for the TABLES physical trace-log delete-and-recreate change; Setup-phase ownership and the guarded-module roster are unchanged.
'     Synchronizes Module Guard for the A2 phase-scoped foundation fingerprint and deferred A2 physical-retirement fast-path repair; B5 relationship ownership and the managed-module roster are unchanged.
'     Synchronizes Module Guard for the safe runtime/candidate/I1/MOC/proof/Ribbon/Defaults exact-identity cleanup; guarded roster and Setup-phase ownership are unchanged.
'     Synchronizes Module Guard for the final Audit 8 DEFAULTS navigation boundary and UI Ribbon source-helper ordering correction; executable behavior, guarded roster, and Setup-phase ownership are unchanged.
'     Synchronizes Module Guard for the safe DB-object exact-identity Trim$ cleanup; the guarded roster and Setup-phase ownership are unchanged.
'     Synchronizes Module Guard for the final Audit 8 DEFAULTS and MOC Lifecycle navigation corrections; executable behavior, module ownership, and Setup-phase boundaries are unchanged.
'     Synchronizes the approved build for the MOC USysRibbons compile correction; module roster, Setup-phase ownership, and lifecycle ownership are unchanged.
'     Synchronizes Module Guard for the final Audit 8 UI Ribbon section-order correction; executable Ribbon behavior and Setup G1-G3/I1 ownership are unchanged.
'     Adds internal Module Guard navigation sections and synchronizes the approved coordinated build for the Audit 8 module-organization correction.
'     Synchronizes Module Guard for the x__UTIL_ADC_CONFIG to x__UTIL_ADC_MACRO_RUNTIME rename and the coordinated TMacroRecord/MacroRecordFromRecordset caller updates while preserving historical
'     predecessor evidence.
'     Restores required fresh D4 editor-form and G3/recovery CustomRibbonID verification reads after concurrency review; Setup-phase ownership and the static no-edit reference module are unchanged.
'     Synchronizes the coordinated build for removal of remaining managed-description Trim$ normalization in MOC; Setup-phase ownership is unchanged.
'     Synchronizes the coordinated build for Audit 6 E3 stale-documentation cleanup; executable behavior and Setup-phase ownership are unchanged.
'     Synchronizes the coordinated build for removal of direct managed-object Description Trim$ mutations in REVIEW and CONFIG; Setup-phase ownership is unchanged.
'     Synchronizes the coordinated build for Audit 6 stale-documentation cleanup; executable behavior and Setup-phase ownership are unchanged.
'     Synchronized the coordinated build for exact managed-object description persistence and removal of the dead shared sentence formatter without changing Setup-phase ownership.
'     Corrects stale Ribbon phase documentation to G1-G3; Ribbon behavior, Setup-phase ownership, and the canonical module roster are unchanged.
'     Synchronizes the coordinated guarded build for the Ribbon dead-parameter cleanup; Ribbon G1-G3 ownership, behavior, and the canonical module roster are unchanged.
'     Synchronizes the coordinated guarded build for the Audit 6 proof-lifetime regression cleanup; the canonical module roster and Setup-phase architecture are unchanged.
'     Synchronizes the coordinated guarded build for the DB Inventory VBA dead-local cleanup; the canonical module roster and Setup-phase architecture are unchanged.
'     Updates the canonical owned-module roster for x__UTIL_ADC_MOC plus new x__UTIL_ADC_MOC_LIFECYCLE and synchronizes every module whose qualified reference changed in the split.
'     Synchronizes the coordinated guarded build for the current private cleanup while preserving the existing Module Guard ownership roster and Setup-phase architecture.
'     Advances the exact coordinated build to 17 managed modules for the x__UTIL_ADC_DB_REFERENCE extraction, with updated Catalog and Module Guard identities and the new module added to the canonical roster.
'     Repairs the Module Guard approved-build definition so its transition-history argument stays within VBA's 24-consecutive explicit line-continuation limit.
'     No Setup-phase ownership or functional workflow changes.
'     Synchronizes the coordinated build after scoping saved-query SQL-mismatch diagnostic throttling to each A3 batch or standalone query-repair operation.
'     Synchronizes the exact coordinated build for standalone managed Description completion in Catalog, Tables, and generated forms, including cold-start-safe TABLES field descriptions and Core
'     nested-scan suppression.
'     Adds optional standalone Module Guard object Description reconciliation while internal build registration and all-table/query repair defer to their owning description paths; synchronizes the coordinated
'     build manifest.
'     Synchronizes the exact coordinated build for the phase-safe A4/B5 relationship-proof repair and single-snapshot loaded-module RI scan ordering in x__UTIL_ADC_REFERENCES.
'     Advances the exact coordinated build for retained Reference/Add-In RI coverage and source-scan parent-before-child ordering, with TABLES and REFERENCES lineage synchronized.
'     Advances the coordinated build for authoritative Discovered User RibbonTitle derivation from the physical object name minus only the leading "ADC " prefix.
'     Replaces registration Boolean/output bundles with named option flags and a typed result contract, advances the coordinated build, and preserves transaction/verification ownership.
'     Synchronized the exact coordinated build for centralized Public Const Long/Short field-description defaults and dedicated runtime-derived description builders in
'     Synchronized the exact coordinated build for the generated-form Long Text sort repair that preserves the shared DAO builder contract while appending the canonical sortable ORDER BY expression.
'     Synchronized the exact coordinated 12-module build for native managed Reports, the new E1-E3 Reports phase, shifted F-I lifecycle ownership, and Phase-I final proof.
'     Synchronized the exact coordinated build for the D4 Build/Repair Help Document Macros performance repair: CONFIG-owned operation-scoped callable-procedure lookup, CATALOG dirty-domain compositional final
'     proof, and coarse D4 performance evidence.
'     Synchronized the exact coordinated 11-module build for the Access/ACE JOIN 3296 repair, dependency-safe canonical Reference query order, entry-point predicate hardening, and exact saved-query A4 diagnostics.
'     Synchronized the exact coordinated 11-module build for the Phase C physical-macro verification performance repair; C1 remains the pre-mutation owner, C2 validates only planned mutation targets, and C3 retains
'     stale-delete protection plus the final exact-state proof.
'     Synchronized the exact coordinated 11-module build for DBInventory_AddIns+VBAPrjs ACE UNION type anchoring and A4 executable validation.
'     Synchronized the exact coordinated 11-module build for Build/Repair Help Document Macros D4 proof renewal and B4/C/D4/F/G authority-ordered orchestration.
'     Synchronized the exact coordinated 11-module build for the B3/B4 MAC.ObjectDescription postcondition authority repair; F1 remains the sole Setup description writer.
'     Synchronized the coordinated Setup phase-authority model now expressed as A-I: D Forms/logical catalog, E Reports, F Descriptions, G Ribbon, H Navigation Pane, I Final Setup Verification.
'     Synchronized the exact 11-module coordinated build for standalone scan overlap cleanup, deferred parent finalization, and updated scan System Macro descriptions.
'     Allows an explicitly reviewed component-type migration only when the exact authorized active predecessor type/version/previous/timestamp is proven by the embedded predecessor definition and append-only
'     history; accidental type mismatches still fail closed.
'     Module Guard synchronization for the DEFAULTS standard-module conversion; managed-header dependency metadata now no longer classifies DEFAULTS as a class.
'     Synchronized the exact 11-module approved build and exact latest supplied predecessor after removing four dead DEFAULTS collection projections and the redundant TABLES foundation-table collection copy.
'     Synchronized the exact 11-module approved build and complete supplied predecessor roster after the operation-local System Macro definition-reuse cleanup.
'     Reused each successful Module Guard operation's already-captured baseline order and installed Version-ID inventory for reporting, then synchronized the exact 11-module build and supplied-baseline predecessors.
'     Synchronized the exact 11-module approved build and immediate predecessor for caller-scoped FormDesign family projection in DEFAULTS/TABLES, eliminating eager module-level FormDesign snapshot collections while
'     preserving one-pass canonical ordering.
'     Synchronized the exact 11-module approved build and immediate predecessor after removing the noncanonical rel_NavigationPaneGroups_SettingKey RI definition from TABLES.
'     Synchronized the exact 11-module approved build and exact supplied predecessor for traditional Access add-in USysRegInfo entry-point inventory.
'     Synchronized the exact approved build for table/query-specific System Macro default descriptions in
'     Synchronized the exact 11-module approved build and immediate predecessor after the Delete All relationship-order and macro-classification exact-schema repair.
'     Synchronized the exact 11-module approved build for canonical static Ribbon tags on known non-generated forms and Help-document macros, with the supplied 2:20 PM Help Document Numbering Compile Fix package as
'     the exact immediate predecessor.
'     Synchronized the exact 11-module coordinated build and exact 10:29 AM predecessor package for the generated continuous-form label-layout correction in UI.
'     Routed Module Guard enable-setting mutation through TABLES with its original verification semantics and synchronized the final authority-cleanup build.
'     Synchronized the exact 11-module approved build and exact 9:39 AM predecessor for the four verified private micro-helper removals.
'     Synchronized the exact 11-module approved build and exact immediate predecessor for DEFAULTS family-vocabulary/query-classification and Core/UI durable-recovery authority consolidation.
'     Synchronized the exact 11-module approved build and immediate predecessor definitions for the HealthReview, Settings, Ribbon target-map, Catalog snapshot, progress-presentation, and TABLES-ownership cleanup.
'     Synchronized the current Module Guard model for the permanent 11-module ADC roster after Diff consolidation while preserving the exact supplied 16-module build as historical predecessor lineage.
'     Advanced the exact 11-module coordinated build for permanent Ribbon button/group XML IDs derived from canonical static tags, with this supplied 11-module package as the exact immediate predecessor.
'     Synchronized the exact approved build for Setup progress qualification/presentation separation and the expanded Setup Control status layout, with the supplied 12:38 PM build as the exact immediate predecessor.
'     Synchronized the exact approved build after making the running Setup Control popup non-modal so generated temporary Design-view forms can own active-form focus during D2.
'     Synchronized the exact 16-module coordinated build after canonical-registry-driven retirement of stale ADC-owned managed-table metadata without deleting physical legacy tables.
'     Synchronized the exact coordinated build after B3 definition-family validation was reattached to the canonical definition providers and the post-cleanup numeric-invariant audit completed.
'     Synchronized the exact 16-module coordinated build after repairing compile-breaking cleanup residue in DEFAULTS and UI.
'     Advanced the exact 17-module coordinated build for the nested Setup progress-dialog routing repair and synchronized the exact immediate predecessor package.
'     Advanced the exact 17-module coordinated build for detailed live status output across Scan Runner Macro Candidates, Active/Available References, Loaded VBA Projects, Add-ins, and External VBA Projects, and
'     synchronized the exact immediate predecessor package.
'     Advanced the exact 17-module coordinated build for the B5 RibbonStaticTag phase-scope repair and synchronized the exact immediate predecessor package.
'     Advanced the exact 17-module coordinated build for the UtilityLazyLoadSelectedTab shared UI message-title compile fix and synchronized the exact immediate predecessor package.
'     Advanced the exact coordinated build for B1 MAC RibbonStaticTag schema-contract invalidation so stale schema evidence cannot bypass required field creation.
'     Advanced the exact 17-module coordinated build for deterministic static Ribbon tags on every generated tab/group/button and bounded remaining canonical field descriptions to 255 characters or fewer.
'     Advanced the exact 17-module coordinated build for the validation-timestamp false-staleness repair, generated-macro contract verification, and aligned validation metadata.
'     Synchronized the managed-header coordinated-build inventory to the exact 15:12 authority-cleanup module identities.
'     Advanced the coordinated approved build for duplicate-authority consolidation.
'     Updated the exact approved coordinated build for Core process-run authority and UI visible-status authority consolidation.
'     Updated the coordinated Module Guard build for setup rollback, Review recovery, and Help_Terms structural authority consolidation.
'     Updated the exact approved coordinated build for the authority-consolidation changes in CATALOG, REVIEWDiff, TABLES, and TABLESDiff.
'     Updated the approved coordinated build after consolidating duplicated Ribbon substitution rules and MAC physical-name identity in DEFAULTS.
'     Updated the approved coordinated build after consolidating generated UI_Frm_Registry write authority in x__UTIL_ADC_UI and removing CATALOG-owned generated ownership/classification rewrites.
'     Updated the approved coordinated build after consolidating generated-form name inventory authority in x__UTIL_ADC_DEFAULTS while preserving UI deletion exclusions for Setup-owned forms.
'     Updated the approved coordinated build for duplicate-review cache/calculation authority consolidation in
'     Updated the approved coordinated build after removal of the confirmed dead-code procedures/declarations and required surviving-reference cleanup.
'     Updated the approved coordinated build for restoration of the three public single-record entry points used by the ADC Config form.
'     Updated the approved coordinated build for duplicate Ribbon CustomID rejection and single-row USysRibbons persistence per CustomRibbonID.
'     Updated the approved coordinated build for unique deterministic Ribbon tags, one authoritative Ribbon group per managed object, and confirmed dead-code removal.
'     Updated the approved coordinated build for stable deterministic per-object Ribbon tags.
'     Updated the approved coordinated build for the System Macro 65 terminal-destructive-path repair.
'     Reference Manager OnOpen preparation and the Setup-managed Report Error form are the current authorities for their preparation and workflow behavior.
'     Synchronized the embedded approved coordinated build to the final terminal error-popup / Report Error audit package after rewrite tracking advanced module identities.
'     Synchronized the embedded approved coordinated build after final canonical command-name presentation cleanup.
'     Updated the approved coordinated build after removing five verified unnecessary wrappers/aliases and routing all callers directly to the existing Module Guard, TABLES classification, and DEFAULTS Ribbon-tag
'     authorities.
'     Synchronized the exact coordinated build after documenting the intentional pre-D Setup UI bootstrap exceptions and correcting stale generated-form phase references to D1/D2/D ownership.
'     Synchronized the exact coordinated build for D4 wrapper-validation lookup reuse and dirty-domain compositional complete-catalog proof.
'     Advanced the exact coordinated build after removing six verified unnecessary wrappers/aliases and routing callers directly to the existing TABLES and DEFAULTS authorities.
'     Advanced the coordinated build for Audit 6 setup/metadata redundancy cleanup while retaining standalone/recovery defenses and existing setup-phase ownership.
' Notes:
'     Baseline creation reads every guarded module identity, type, visible current and previous Version ID, timestamp, and managed-header name from that module's commented rewrite header. Later registration
'     may replace an older active build whose
'     visible manifest identity and lineage match append-only history and whose timestamp precedes the incoming build. A source branch whose managed-header predecessor is already behind the active approved lineage
'     is rejected and must be corrected at the baseline; source code cannot authorize the mismatch by carrying a module Version ID exception. A changed module must directly follow the database's active approved
'     Version ID; intermediate module versions are not inferred or skipped during automatic acceptance. RegisterApprovedBuild remains strict to the fixed roster and header-derived coordinated identity. The Setup
'     mismatch prompt may explicitly approve the exact installed build after structural/header validation; that runtime choice intentionally overrides active-manifest lineage and rollback-timestamp rejection while
'     preserving exact-version history immutability, transactional history/manifest writes, and staged verification.
' Core Rules Enforced Here:
'     Before any structure, manifest, version, history, or registration work, fail closed when an ADC-owned actual VBE component name differs from its managed-header VB_Name or another component claims an ADC-owned
'     VB_Name; never silently accept a mixed module batch; never overwrite conflicting exact-version history; never infer approval from installed Version IDs without the explicit Setup Yes choice and complete
'     structural validation; automatic acceptance permits no skipped module-version deployment lineage and requires every changed module to directly identify the active approved Version ID as its commented
'     predecessor; only the runtime explicit-approval choice may override that active-manifest lineage/timestamp gate; automatic component-type migration is permitted only when the incoming build directly
'     follows the active managed-header Version ID and persistent manifest/history evidence proves the active old type/version lineage; preserve verification evidence even when the guard blocks setup.
' Design Map:
'     Structure repair; initial baseline seeding; transactional batch registration; installed-module scan; manifest comparison; append-only verification snapshot; saved review queries.
' Configuration: x__UTIL_ADC_MODULE_GUARD_DEFINITIONS owns the changing approved-build batch identity and guarded roster; active ModuleVersion_Manifest rows remain authoritative after initial baseline seeding.
' Behavior:
'     Every verification and registration entry path first performs a read-only VBE component/header identity pass and stops immediately on a blank or mismatched ADC-owned VB_Name before database-backed Guard work
'     begins. Initial baseline seeding and approved-build registration govern acceptance by visible component identity, type, current/previous Version IDs, timestamps, required presence, Version-ID uniqueness, and
'     append-only lineage. Any coordinated build that changes another guarded module must also advance the definitions module Version ID and approved-build batch identity.
' Code Status: Active.
' Todo: None.
' Reporting Bugs: None currently recorded in this header.
' Module Level Bugs: None currently recorded in this header.
' Public Subs:
'     ResolveModuleUpdateGuardReadiness(ByRef infrastructureExists As Boolean, ByRef manifestPopulated As Boolean, Optional ByRef missingTableNames As String = vbNullString) -> None
'     RepairModuleUpdateGuard(Optional ByVal showMessage As Boolean = True, Optional ByVal applyStandaloneDescriptions As Boolean = True) -> None
'     VerifyModuleUpdateGuard(Optional ByVal showMessage As Boolean = True, Optional ByRef verificationDetails As String = vbNullString, Optional ByVal suppressFailureReporting As Boolean = False, Optional ByRef
'     installedVersionSummary As String = vbNullString) -> None
'     SetModuleUpdateGuardEnabledForSetup(ByVal enabled As Boolean) -> None
'     RegisterApprovedBuild(Optional ByVal showMessage As Boolean = True) -> None
' Public Functions:
'     ModuleUpdateGuardEnabled(Optional ByVal withoutRepair As Boolean = False) -> Boolean
'     TryPrepareAndVerifyModuleUpdateGuard(ByRef failedErrNumber As Long, ByRef failedErrSource As String, ByRef failedErrDescription As String, Optional ByRef verificationDetails As String = vbNullString,
'     Optional ByVal readinessResolved As Boolean = False, Optional ByVal infrastructureExists As Boolean = False, Optional ByVal manifestPopulated As Boolean = False, Optional ByVal missingTableNames As String
'     = vbNullString, Optional ByRef installedVersionSummary As String = vbNullString) -> Boolean
'     ModuleUpdateGuardFinalStateTokenForI1(ByRef detailText As String) -> String
'     ModuleGuardChatGPTVersionIDFromComponent(ByVal component As Object) -> String
'     ModuleGuardChatGPTVersionIDByModuleName(ByVal ModuleName As String) -> String
'     TryRegisterApprovedBuild(ByRef registrationResult As TModuleGuardRegistrationResult, Optional ByVal registrationOptions As ModuleGuardRegistrationOptions =
'     MODULE_GUARD_REGISTRATION_OPTION_ENABLE_AFTER_REGISTRATION) -> Boolean
'     ModuleUpdateHistoryFieldOrderText() -> String
'     ModuleVersionManifestFieldOrderText() -> String
'     ModuleVersionVerificationFieldOrderText() -> String
' Public Properties: None
' ================================================================================================================================================================================

Option Compare Database
Option Explicit

' ===============================================================================================================================================================================
' MODULE-LEVEL DECLARATIONS
' ===============================================================================================================================================================================

' Approved-build registration result and option contracts shared with Core.
Public Type TModuleGuardRegistrationResult
    failedErrNumber As Long
    failedErrSource As String
    failedErrDescription As String
    RegistrationDetails As String
    ExplicitApprovalCandidate As Boolean
    installedVersionSummary As String
End Type

Public Enum ModuleGuardRegistrationOptions
    MODULE_GUARD_REGISTRATION_OPTION_NONE = 0
    MODULE_GUARD_REGISTRATION_OPTION_ENABLE_AFTER_REGISTRATION = 2
    MODULE_GUARD_REGISTRATION_OPTION_EXPLICITLY_APPROVE_INSTALLED_BUILD = 4
End Enum

' Registration provenance values.
Public Const MODULE_GUARD_RECORDED_BY_INITIAL_BASELINE As String = "InitialBaseline"
Public Const MODULE_GUARD_RECORDED_BY_REGISTERED_BATCH As String = "RegisteredBatch"
Public Const MODULE_GUARD_RECORDED_BY_APPROVED_BUILD As String = "ApprovedCurrentBuild"
Public Const MODULE_GUARD_RECORDED_BY_MANUAL As String = "Manual"

' Persisted structure-state property and version.
Private Const MODULE_GUARD_STRUCTURE_STATE_PROPERTY As String = "ADCModuleGuardStructureStateV1"
' Semantic structure-state schema version; not a module Version ID. Changing it forces one full structure verification/repair before the new steady-state token is accepted.
Private Const MODULE_GUARD_STRUCTURE_STATE_VERSION As String = "MG-STRUCTURE-STATE-V1"

' Guard error numbers and supported component types.
Private Const MODULE_GUARD_ERROR_BASE As Long = vbObjectError + 9720
Private Const MODULE_GUARD_HEADER_IDENTITY_ERROR As Long = MODULE_GUARD_ERROR_BASE + 125
Public Const MODULE_GUARD_COMPONENT_STANDARD As String = "Standard Module"
Public Const MODULE_GUARD_COMPONENT_CLASS As String = "Class Module"
Public Const MODULE_GUARD_COMPONENT_FORM As String = "Form Module"
Public Const MODULE_GUARD_COMPONENT_DOCUMENT As String = "Document Module"
Private Const MODULE_GUARD_DEFINITIONS_MODULE_NAME As String = "x__UTIL_ADC_MODULE_GUARD_DEFINITIONS"

' Verification result values and tracking-ID format.
Public Const MODULE_GUARD_RESULT_MATCH As String = "Match"
Public Const MODULE_GUARD_RESULT_MISMATCH As String = "Mismatch"
Public Const MODULE_GUARD_RESULT_MISSING As String = "Missing"
Public Const MODULE_GUARD_RESULT_UNREGISTERED As String = "Unregistered"
Public Const MODULE_GUARD_RESULT_BROKEN_CHAIN As String = "BrokenChain"
Public Const MODULE_GUARD_RESULT_DUPLICATE_ID As String = "DuplicateVersionID"
Public Const MODULE_GUARD_RESULT_HEADER_MISMATCH As String = "HeaderMismatch"
Public Const MODULE_GUARD_RESULT_INVALID_ID As String = "InvalidVersionID"
Private Const MODULE_GUARD_TRACKING_PREFIX As String = "CGPT-"
Private Const MODULE_GUARD_NO_VERSION_SUFFIX As String = "NONE"
Private Const MODULE_GUARD_PLACEHOLDER_VERSION_SUFFIX As String = "XXXX-XXXX-XXXX-XXXX-XXXX"
Private Const MODULE_GUARD_ID_ALPHABET As String = "23456789ABCDEFGHJKLMNPQRSTUVWXYZ"
Private Const MODULE_GUARD_MAX_DETAIL_ROWS As Long = 40

' Registration transaction, verification, pump, and component-cache state.
Private mModuleGuardRegistrationTransactionOwned As Boolean
Private mModuleGuardNestedVerificationTransactionActive As Boolean
Private mModuleGuardPreflightActive As Boolean
' Transient authoritative installed-module evidence for the active Preflight/registration operation only; valid only while the VBProject is unchanged and cleared at every operation boundary.
Private mModuleGuardInstalledInventoryEvidence As Collection
Private mModuleGuardPumpActive As Boolean
Private mModuleGuardLastPumpTimer As Single
' INTENTIONAL KEYED LOOKUP CACHE:
' Derived from the fixed guarded-module roster owned by x__UTIL_ADC_MODULE_GUARD_DEFINITIONS; it contains module names only and no executable module Version IDs. It is module-owned for
' the loaded code version, has no independent mutation path, and is rebuilt only from the guarded roster when Nothing.
Private mModuleGuardOwnedComponentLookup As Object

' ===============================================================================================================================================================================
' PROCEDURES
' ===============================================================================================================================================================================

' =============================================================================================================================================================================
' READINESS / PREFLIGHT / VERIFICATION PROCEDURES
' =============================================================================================================================================================================

Private Sub PumpModuleGuardSetupResponsiveness(Optional ByVal forcePump As Boolean = False)
    On Error GoTo SafeExit
    Dim currentTimer As Single
    Dim elapsedSeconds As Double
    If mModuleGuardPumpActive Then Exit Sub
    currentTimer = Timer
    If forcePump = False And mModuleGuardLastPumpTimer > 0 Then
        If currentTimer >= mModuleGuardLastPumpTimer Then
            elapsedSeconds = CDbl(currentTimer - mModuleGuardLastPumpTimer)
        Else
            elapsedSeconds = CDbl((86400! - mModuleGuardLastPumpTimer) + currentTimer)
        End If
        If elapsedSeconds < DEFAULT_SETUP_EVENT_PUMP_INTERVAL_SECONDS Then Exit Sub
    End If
    mModuleGuardLastPumpTimer = currentTimer
    mModuleGuardPumpActive = True
    DoEvents
SafeExit:
    mModuleGuardPumpActive = False
End Sub

Private Sub LogModuleGuardPreflightBoundary(ByVal sourceProcedure As String, ByVal messageText As String, Optional ByVal detailText As String = vbNullString, Optional ByVal severityText As String = TRACE_LEVEL_STATUS)
    On Error Resume Next
    Call CorePulseModuleGuardPreflight(messageText)
    On Error GoTo 0
    LogModuleGuardMessage messageText, detailText, severityText, sourceProcedure, messageText
End Sub

' Purpose: Performs module update guard enabled by coordinating `ModuleUpdateGuardEnabledWithoutRepair` and `End` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   withoutRepair (Input Optional; Boolean): True requests the repair-free/read-only path; False allows the routine's normal repair-capable behavior. Default: False.
' Returns: Boolean: True when module update guard is enabled by the effective configuration/state; False when disabled.
Public Function ModuleUpdateGuardEnabled(Optional ByVal withoutRepair As Boolean = False) As Boolean
    'Canonical effective enablement policy. Setup preflight may request a raw no-repair read to avoid readiness/repair re-entry.
    If withoutRepair Then
        ModuleUpdateGuardEnabled = ModuleUpdateGuardEnabledWithoutRepair()
    Else
        ModuleUpdateGuardEnabled = SettingBool(SETTING_ENABLE_MODULE_UPDATE_GUARD, DEFAULT_ENABLE_MODULE_UPDATE_GUARD)
    End If
End Function

Private Function ModuleUpdateGuardEnabledWithoutRepair() As Boolean
    On Error GoTo UseDefault
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim settingsTableName As String
    Dim valueText As String
    settingsTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID)
    Set db = ADCDb()
    If SharedDao_DatabaseTableExistsSafe(db, settingsTableName) = False Then GoTo UseDefault
    If SharedDao_DatabaseTableFieldExistsSafe(db, settingsTableName, FIELD_SETTING_KEY) = False Then GoTo UseDefault
    If SharedDao_DatabaseTableFieldExistsSafe(db, settingsTableName, FIELD_SETTING_VALUE) = False Then GoTo UseDefault
    Set rs = db.OpenRecordset("SELECT [" & FIELD_SETTING_VALUE & "] FROM [" & settingsTableName & "] WHERE [" & FIELD_SETTING_KEY & "]=" & SharedDao_SqlQuote(SETTING_ENABLE_MODULE_UPDATE_GUARD) & ";", dbOpenSnapshot)
    If rs.EOF Then GoTo UseDefault
    valueText = Nz(rs.fields(FIELD_SETTING_VALUE).value, vbNullString)
    rs.Close
    Set rs = Nothing
    On Error GoTo 0
    ModuleUpdateGuardEnabledWithoutRepair = BooleanTextValue(valueText, DEFAULT_ENABLE_MODULE_UPDATE_GUARD)
    Exit Function
UseDefault:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    On Error GoTo 0
    ModuleUpdateGuardEnabledWithoutRepair = DEFAULT_ENABLE_MODULE_UPDATE_GUARD
End Function

' Purpose: Attempts to perform prepare and verify module update guard for readiness resolved, infrastructure exists.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   failedErrNumber (Output; Long): Receives the resolved failed err number produced from this procedure's validated inputs and current ADC state.
'   failedErrSource (Output; String): Receives the resolved failed err source produced from this procedure's validated inputs and current ADC state.
'   failedErrDescription (Output; String): Receives the resolved failed err description produced from this procedure's validated inputs and current ADC state.
'   verificationDetails (Output Optional; String): Receives verification details used for caller diagnostics, verification output, or user-facing reporting. Default: vbNullString.
'   readinessResolved (Input Optional; Boolean): Boolean state describing whether readiness resolved applies to this invocation; True selects the named condition/path and False selects its complementary/default
'   path. Default: False.
'   infrastructureExists (Input Optional; Boolean): Boolean state describing whether infrastructure exists applies to this invocation; True selects the named condition/path and False selects its
'   complementary/default path. Default: False.
'   manifestPopulated (Input Optional; Boolean): Boolean state describing whether manifest populated applies to this invocation; True selects the named condition/path and False selects its complementary/default
'   path. Default: False.
'   missingTableNames (Input Optional; String): Delimited/list-form text receiving or supplying the names of required tables that are absent from the inspected database. Default: vbNullString.
'   installedVersionSummary (Output Optional; String): Receives the resolved installed version summary produced from this procedure's validated inputs and current ADC state. Default: vbNullString.
' Returns: Boolean: True when the attempt to prepare and verify module update guard completes successfully; False for the routine's defined normal unsuccessful/not-applicable outcome.
Public Function TryPrepareAndVerifyModuleUpdateGuard(ByRef failedErrNumber As Long, ByRef failedErrSource As String, ByRef failedErrDescription As String, Optional ByRef verificationDetails As String = vbNullString, Optional ByVal readinessResolved As Boolean = False, Optional ByVal infrastructureExists As Boolean = False, Optional ByVal manifestPopulated As Boolean = False, Optional ByVal missingTableNames As String = vbNullString, Optional ByRef installedVersionSummary As String = vbNullString) As Boolean
    On Error GoTo VerificationFailed
    Dim originalErrNumber As Long
    Dim originalErrSource As String
    Dim originalErrDescription As String
    Dim registrationResult As TModuleGuardRegistrationResult
    Dim transcriptText As String
    Dim namingSnapshotTimer As Single
    Dim readinessTimer As Single
    Dim verificationTimer As Single
    Dim registrationTimer As Single
    Dim namingSnapshotOwned As Boolean
    failedErrNumber = 0
    failedErrSource = vbNullString
    failedErrDescription = vbNullString
    verificationDetails = vbNullString
    installedVersionSummary = vbNullString
    If mModuleGuardPreflightActive Then
        failedErrNumber = MODULE_GUARD_ERROR_BASE + 126
        failedErrSource = "TryPrepareAndVerifyModuleUpdateGuard"
        failedErrDescription = "Module Guard preflight reentry was blocked. Another Module Guard preflight is already active; setup cannot continue until the original verification returns."
        verificationDetails = failedErrDescription
        LogModuleGuardPreflightBoundary "TryPrepareAndVerifyModuleUpdateGuard", "Module Guard preflight reentry blocked.", failedErrDescription, TRACE_LEVEL_ERROR
        Exit Function
    End If
    mModuleGuardPreflightActive = True
    Set mModuleGuardInstalledInventoryEvidence = Nothing
    mModuleGuardLastPumpTimer = 0
    LogModuleGuardPreflightBoundary "TryPrepareAndVerifyModuleUpdateGuard", "Module Guard preflight started.", "Stage sequence: use caller-resolved readiness when supplied; otherwise build one temporary naming snapshot and resolve infrastructure plus active-manifest state once. If the registered manifest differs, invoke approved-build registration only after the managed-header coordinated identity passes rollback and stale-branch checks against persistent history. Otherwise return failure for the explicit setup choice."
    If readinessResolved = False Then
        namingSnapshotTimer = Timer
        BeginNamingSnapshot
        namingSnapshotOwned = True
        LogModuleGuardPreflightBoundary "TryPrepareAndVerifyModuleUpdateGuard", "Module Guard preflight naming snapshot ready.", "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(namingSnapshotTimer)), TRACE_LEVEL_INFO
        readinessTimer = Timer
        ResolveModuleUpdateGuardReadiness infrastructureExists, manifestPopulated, missingTableNames
        LogModuleGuardPreflightBoundary "TryPrepareAndVerifyModuleUpdateGuard", "Module Guard readiness resolved once.", "Infrastructure ready: " & CStr(infrastructureExists) & vbCrLf & "Active manifest populated: " & CStr(manifestPopulated) & vbCrLf & "Missing tables: " & IIf(Len(missingTableNames) = 0, "(none)", missingTableNames) & vbCrLf & "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(readinessTimer)), TRACE_LEVEL_INFO
    End If
    If infrastructureExists = False Then
        Err.Raise MODULE_GUARD_ERROR_BASE + 127, "TryPrepareAndVerifyModuleUpdateGuard", _
            "Module Guard infrastructure is not ready for ordinary setup verification." & vbCrLf & _
            "Missing tables: " & IIf(Len(missingTableNames) = 0, "(not identified)", missingTableNames) & vbCrLf & _
            "Setup did not create or repair Module Guard infrastructure." & vbCrLf & _
            "Run RegisterApprovedBuild, then rerun RunFullSetup."
    End If
    If manifestPopulated = False Then
        Err.Raise MODULE_GUARD_ERROR_BASE + 128, "TryPrepareAndVerifyModuleUpdateGuard", _
            "The active Module Guard manifest is not populated for the installed build." & vbCrLf & _
            "Setup did not seed or register the manifest." & vbCrLf & _
            "Run RegisterApprovedBuild, then rerun RunFullSetup."
    End If
    LogModuleGuardPreflightBoundary "TryPrepareAndVerifyModuleUpdateGuard", "Module Guard preflight: one-pass installed inventory and active-manifest comparison starting."
    verificationTimer = Timer
    On Error GoTo RegisteredManifestVerificationFailed
    VerifyModuleUpdateGuard False, verificationDetails, True, installedVersionSummary
    On Error GoTo VerificationFailed
    LogModuleGuardPreflightBoundary "TryPrepareAndVerifyModuleUpdateGuard", "Module Guard installed-module verification timing.", "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(verificationTimer)), TRACE_LEVEL_INFO
    LogModuleGuardMessage "Module update guard verified the installed guarded module set before material work.", verificationDetails
    LogModuleGuardPreflightBoundary "TryPrepareAndVerifyModuleUpdateGuard", "Module Guard preflight completed successfully.", verificationDetails & vbCrLf & "The already-registered manifest matched; no registration or approved manifest-identity write was performed.", TRACE_LEVEL_INFO
    TryPrepareAndVerifyModuleUpdateGuard = True
CleanExit:
    On Error Resume Next
    If namingSnapshotOwned Then InvalidateNamingSnapshot
    On Error GoTo 0
    Set mModuleGuardInstalledInventoryEvidence = Nothing
    mModuleGuardPreflightActive = False
    mModuleGuardLastPumpTimer = 0
    Exit Function
RegisteredManifestVerificationFailed:
    originalErrNumber = Err.Number
    originalErrSource = Err.Source
    originalErrDescription = Err.Description
    Err.Clear
    On Error GoTo VerificationFailed
    If originalErrNumber <> MODULE_GUARD_ERROR_BASE + 1 Then
        LogModuleGuardPreflightBoundary "TryPrepareAndVerifyModuleUpdateGuard", "Module Guard installed-module verification timing.", "Elapsed before failure: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(verificationTimer)), TRACE_LEVEL_INFO
        failedErrNumber = originalErrNumber
        failedErrSource = originalErrSource
        failedErrDescription = originalErrDescription
        GoTo VerificationFailureReady
    End If
    LogModuleGuardPreflightBoundary "TryPrepareAndVerifyModuleUpdateGuard", "Module Guard installed-module verification timing.", "Elapsed before caller-managed stale-manifest result: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(verificationTimer)), TRACE_LEVEL_INFO
    LogModuleGuardPreflightBoundary "TryPrepareAndVerifyModuleUpdateGuard", "The registered manifest is stale relative to the installed coordinated build; automatic exact embedded-build validation is starting.", "Caller-managed verification result: " & CStr(originalErrNumber) & " from " & originalErrSource, TRACE_LEVEL_INFO
    registrationTimer = Timer
    If TryRegisterApprovedBuild(registrationResult, MODULE_GUARD_REGISTRATION_OPTION_NONE) Then
        installedVersionSummary = registrationResult.installedVersionSummary
        LogModuleGuardPreflightBoundary "TryPrepareAndVerifyModuleUpdateGuard", "Module Guard automatic approved-build registration timing.", "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(registrationTimer)), TRACE_LEVEL_INFO
        verificationDetails = "Module Guard automatically accepted and registered the complete managed-header coordinated build after the prior active manifest proved stale." & vbCrLf & registrationResult.RegistrationDetails
        LogModuleGuardMessage "Module update guard automatically registered and verified the exact managed-header coordinated build before material work.", verificationDetails
        LogModuleGuardPreflightBoundary "TryPrepareAndVerifyModuleUpdateGuard", "Module Guard preflight completed successfully after automatic exact-build registration.", verificationDetails, TRACE_LEVEL_INFO
        TryPrepareAndVerifyModuleUpdateGuard = True
        GoTo CleanExit
    End If
    LogModuleGuardPreflightBoundary "TryPrepareAndVerifyModuleUpdateGuard", "Module Guard automatic approved-build registration timing.", "Elapsed before failure: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(registrationTimer)), TRACE_LEVEL_INFO
    failedErrNumber = originalErrNumber
    failedErrSource = originalErrSource
    failedErrDescription = originalErrDescription & vbCrLf & vbCrLf & _
        "Automatic exact embedded-build acceptance did not complete." & vbCrLf & _
        "Registration Err.Number: " & CStr(registrationResult.failedErrNumber) & vbCrLf & _
        "Registration Err.Source: " & registrationResult.failedErrSource & vbCrLf & _
        "Registration Err.Description: " & registrationResult.failedErrDescription
    GoTo VerificationFailureReady
VerificationFailed:
    originalErrNumber = Err.Number
    originalErrSource = Err.Source
    originalErrDescription = Err.Description
    failedErrNumber = originalErrNumber
    failedErrSource = originalErrSource
    failedErrDescription = originalErrDescription
VerificationFailureReady:
    If failedErrNumber = 0 Then failedErrNumber = MODULE_GUARD_ERROR_BASE + 2
    If Len(Trim$(failedErrSource)) = 0 Then failedErrSource = "TryPrepareAndVerifyModuleUpdateGuard"
    If Len(Trim$(failedErrDescription)) = 0 Then failedErrDescription = "The module update guard verification did not complete."
    If failedErrNumber <> MODULE_GUARD_HEADER_IDENTITY_ERROR Then
        If InStr(1, failedErrDescription, "RegisterApprovedBuild", vbBinaryCompare) = 0 Then
            failedErrDescription = failedErrDescription & vbCrLf & vbCrLf & _
                "Setup did not alter the active manifest or registration history after automatic exact-build acceptance failed." & vbCrLf & _
                "Use the Setup mismatch choice to explicitly approve and register the structurally validated installed module set, or stop without changing approval."
        End If
    End If
    transcriptText = BuildModuleGuardVerificationFailureTranscript( _
        DEFAULT_UTILITY_LONG_NAME & " module update verification failed." & vbCrLf & vbCrLf & _
        "Err.Number: " & CStr(failedErrNumber) & vbCrLf & _
        "Err.Source: " & failedErrSource & vbCrLf & _
        "Err.Description:" & vbCrLf & failedErrDescription)
    If Len(verificationDetails) = 0 Or IsModuleGuardVerificationFailureTranscript(verificationDetails) = False Then verificationDetails = transcriptText
    On Error Resume Next
    Debug.Print verificationDetails
    LogModuleGuardMessage "Module update guard verification failed before material setup work.", "Err.Number: " & CStr(failedErrNumber) & vbCrLf & "Err.Source: " & failedErrSource & vbCrLf & "Err.Description: " & failedErrDescription, TRACE_LEVEL_ERROR, "TryPrepareAndVerifyModuleUpdateGuard", "Verify registered or exact managed-header build"
    LogModuleGuardPreflightBoundary "TryPrepareAndVerifyModuleUpdateGuard", "Module Guard preflight completed with failure.", "Err.Number: " & CStr(failedErrNumber) & vbCrLf & "Err.Source: " & failedErrSource & vbCrLf & "Err.Description: " & failedErrDescription, TRACE_LEVEL_ERROR
    On Error GoTo 0
    GoTo CleanExit
End Function

' Purpose: Resolves Guard table readiness and active-manifest population from one canonical managed-object projection.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   infrastructureExists (Output; Boolean): Set True when the procedure determines that infrastructure exists applies; set False when that condition is not met.
'   manifestPopulated (Output; Boolean): Set True when the procedure determines that manifest populated applies; set False when that condition is not met.
'   missingTableNames (Output Optional; String): Receives the resolved missing table names produced from this procedure's validated inputs and current ADC state. Default: vbNullString.
' Returns: None (Sub). Results are returned through the documented ByRef parameters and/or the routine's database, UI, or module-state side effects.
Public Sub ResolveModuleUpdateGuardReadiness(ByRef infrastructureExists As Boolean, ByRef manifestPopulated As Boolean, Optional ByRef missingTableNames As String = vbNullString)
    'Resolves Guard table readiness and active-manifest population from one canonical managed-object projection.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim definitions As Collection
    Dim Definition As Variant
    Dim objectName As String
    Dim manifestTableName As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    infrastructureExists = False
    manifestPopulated = False
    missingTableNames = vbNullString
    Set db = currentDb
    Set definitions = ModuleGuardManagedObjectDefinitions()
    manifestTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionManifest")
    For Each Definition In definitions
        If StrComp(CStr(Definition(2)), ADCAT_OBJECT_TYPE_TABLE, vbTextCompare) = 0 Then
            objectName = CStr(Definition(0))
            If ModuleGuardTableExists(db, objectName) = False Then AppendGuardDetail missingTableNames, objectName
        End If
    Next Definition
    infrastructureExists = (Len(missingTableNames) = 0)
    If infrastructureExists = False Then Exit Sub
    Set rs = db.OpenRecordset("SELECT " & SQL_FIELD_MODULE_NAME & " FROM [" & Replace(manifestTableName, "]", "]]") & "] WHERE [" & FIELD_ACTIVE & "]=True;", dbOpenSnapshot)
    Do While rs.EOF = False
        PumpModuleGuardSetupResponsiveness
        If ModuleGuardOwnsComponent(CStr(Nz(rs.fields(FIELD_MODULE_NAME), vbNullString))) Then
            manifestPopulated = True
            Exit Do
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    On Error GoTo 0
    If savedErrNumber = 0 Then savedErrNumber = vbObjectError + 6175
    If Len(Trim$(savedErrSource)) = 0 Then savedErrSource = "ResolveModuleUpdateGuardReadiness"
    Err.Raise savedErrNumber, "ResolveModuleUpdateGuardReadiness", "Could not resolve Module Guard readiness from ModuleGuardManagedObjectDefinitions." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

' =============================================================================================================================================================================
' PHYSICAL MODULE GUARD STRUCTURE / QUERY REPAIR PROCEDURES
' =============================================================================================================================================================================

' Purpose: Repairs module update guard for apply standalone descriptions.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
'   applyStandaloneDescriptions (Input Optional; Boolean): When True, enables or requests standalone descriptions for this call; when False, the routine follows its normal non-forced/default path. Default:
'   True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub RepairModuleUpdateGuard(Optional ByVal showMessage As Boolean = True, Optional ByVal applyStandaloneDescriptions As Boolean = True)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim definitions As Collection
    Dim Definition As Variant
    Dim stageText As String
    Dim targetObject As String
    Dim descriptionChanged As Variant
    Dim descriptionChangeCount As Long
    stageText = "Initialize module update guard repair."
    targetObject = "Module update guard"
    LogModuleGuardMessage "Module update guard structure repair started.", "Stage: " & stageText, TRACE_LEVEL_INFO, "RepairModuleUpdateGuard"
    Set db = currentDb
    Set definitions = ModuleGuardManagedObjectDefinitions()
    For Each Definition In definitions
        If StrComp(CStr(Definition(2)), ADCAT_OBJECT_TYPE_TABLE, vbTextCompare) = 0 Then
            targetObject = CStr(Definition(0))
            stageText = "Create or repair canonical Module Guard table."
            LogModuleGuardMessage "Module update guard repair stage started.", "Stage: " & stageText & vbCrLf & "Target: " & targetObject, TRACE_LEVEL_STATUS, "RepairModuleUpdateGuard", stageText
            RepairModuleGuardTableByName db, targetObject
            LogModuleGuardMessage "Module update guard repair stage completed.", "Stage: " & stageText & vbCrLf & "Target: " & targetObject, TRACE_LEVEL_STATUS, "RepairModuleUpdateGuard", stageText
        End If
    Next Definition
    stageText = "Create or repair canonical Module Guard review queries."
    targetObject = "Canonical Module Guard query set"
    LogModuleGuardMessage "Module update guard repair stage started.", "Stage: " & stageText & vbCrLf & "Target: " & targetObject, TRACE_LEVEL_STATUS, "RepairModuleUpdateGuard", stageText
    EnsureModuleGuardQueries db, definitions
    LogModuleGuardMessage "Module update guard repair stage completed.", "Stage: " & stageText & vbCrLf & "Target: " & targetObject, TRACE_LEVEL_STATUS, "RepairModuleUpdateGuard", stageText
    stageText = "Verify complete canonical Module Guard structure."
    targetObject = "Module update guard"
    RequireCompleteModuleGuardStructure db, definitions
    If applyStandaloneDescriptions Then
        stageText = "Apply and verify canonical Module Guard object descriptions."
        For Each Definition In definitions
            targetObject = CStr(Definition(0))
            descriptionChanged = False
            If ApplyAndVerifyManagedObjectDescriptionForStandaloneRepair(targetObject, CStr(Definition(2)), descriptionChanged) = False Then Err.Raise MODULE_GUARD_ERROR_BASE + 58, "RepairModuleUpdateGuard", "Module Guard managed-object Description reconciliation returned False." & vbCrLf & "Object: " & targetObject
            If CBool(descriptionChanged) Then descriptionChangeCount = descriptionChangeCount + 1
        Next Definition
    End If
    PersistModuleGuardStructureState db, definitions
    LogModuleGuardMessage "Module update guard tables and queries were repaired successfully.", ModuleGuardRepairInventoryText(db, definitions) & vbCrLf & "Standalone managed-object descriptions applied: " & CStr(applyStandaloneDescriptions) & vbCrLf & "Managed-object descriptions changed: " & CStr(descriptionChangeCount), TRACE_LEVEL_INFO, "RepairModuleUpdateGuard", stageText
    If showMessage Then UtilityMsgBox DEFAULT_UTILITY_LONG_NAME & " module update guard tables and queries were repaired successfully.", vbInformation, DEFAULT_UTILITY_LONG_NAME & " Module Update Guard"
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim failureText As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    failureText = ModuleGuardRepairFailureText(stageText, targetObject, savedErrNumber, savedErrSource, savedErrDescription, definitions)
    LogModuleGuardMessage "Module update guard structure repair failed.", failureText, TRACE_LEVEL_ERROR, "RepairModuleUpdateGuard", stageText
    If showMessage Then
        CoreShowErrorMessage failureText, vbCritical, DEFAULT_UTILITY_LONG_NAME & " Module Update Guard", "RepairModuleUpdateGuard", "Core"
        Exit Sub
    End If
    Err.Raise savedErrNumber, "RepairModuleUpdateGuard", failureText
End Sub

Private Sub RepairModuleGuardTableByName(ByVal db As DAO.Database, ByVal tableName As String)
    'Maps a canonical Module Guard table identity to its specialized repair behavior. This is build behavior, not table-membership authority.
    Select Case tableName
        Case ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleUpdateHistory")
            EnsureModuleUpdateHistoryTable db
        Case ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionManifest")
            EnsureModuleVersionManifestTable db
            DeactivateNonOwnedManifestRows db
        Case ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionVerification")
            EnsureModuleVersionVerificationTable db
        Case Else
            Err.Raise MODULE_GUARD_ERROR_BASE + 57, "RepairModuleGuardTableByName", "No table repair handler mapping exists for the canonical Module Guard table." & vbCrLf & "Table: " & tableName
    End Select
End Sub

Private Sub PersistModuleGuardStructureState(ByVal db As DAO.Database, Optional ByVal managedDefinitions As Variant)
    'The token is an optimization only. A persistence failure leaves the next run on the full-repair path without weakening verification.
    Dim currentToken As String
    On Error Resume Next
    currentToken = CurrentModuleGuardStructureStateToken(db, managedDefinitions)
    If Len(currentToken) > 0 Then EnsureDaoPropertyValue db, MODULE_GUARD_STRUCTURE_STATE_PROPERTY, dbText, currentToken, "PersistModuleGuardStructureState"
    On Error GoTo 0
End Sub

Private Function CurrentModuleGuardStructureStateToken(ByVal db As DAO.Database, Optional ByVal managedDefinitions As Variant) As String
    On Error GoTo NotCurrent
    Dim definitions As Collection
    Dim Definition As Variant
    Dim objectName As String
    Dim objectType As String
    Dim stateText As String
    db.TableDefs.Refresh
    db.QueryDefs.Refresh
    stateText = MODULE_GUARD_STRUCTURE_STATE_VERSION
    If IsObject(managedDefinitions) Then Set definitions = managedDefinitions
    If definitions Is Nothing Then Set definitions = ModuleGuardManagedObjectDefinitions()
    For Each Definition In definitions
        objectName = CStr(Definition(0))
        objectType = CStr(Definition(2))
        Select Case objectType
            Case ADCAT_OBJECT_TYPE_TABLE
                AppendModuleGuardTableStructureState stateText, db.TableDefs(objectName)
            Case ADCAT_OBJECT_TYPE_QUERY
                AppendModuleGuardQueryStructureState stateText, db.QueryDefs(objectName)
            Case Else
                Err.Raise MODULE_GUARD_ERROR_BASE + 54, "CurrentModuleGuardStructureStateToken", "ModuleGuardManagedObjectDefinitions contains an unsupported object type." & vbCrLf & "Object: " & objectName & vbCrLf & "ObjectType: " & objectType
        End Select
    Next Definition
    CurrentModuleGuardStructureStateToken = MODULE_GUARD_STRUCTURE_STATE_VERSION & "-" & ModuleGuardSourceFingerprint(stateText)
    Exit Function
NotCurrent:
    Err.Clear
End Function

Private Sub AppendModuleGuardTableStructureState(ByRef stateText As String, ByVal tdf As DAO.tableDef)
    Dim fld As DAO.Field
    Dim idx As DAO.index
    Dim indexField As DAO.Field
    stateText = stateText & ModuleGuardStructurePart("TABLE") & ModuleGuardStructurePart(tdf.Name) & ModuleGuardStructurePart(CStr(tdf.Attributes))
    For Each fld In tdf.fields
        PumpModuleGuardSetupResponsiveness
        stateText = stateText & ModuleGuardStructurePart("FIELD") & ModuleGuardStructurePart(fld.Name) & ModuleGuardStructurePart(CStr(fld.Type)) & ModuleGuardStructurePart(CStr(fld.Size)) & ModuleGuardStructurePart(CStr(fld.Attributes))
        stateText = stateText & ModuleGuardStructurePart(ModuleGuardDaoPropertyState(fld, "Required")) & ModuleGuardStructurePart(ModuleGuardDaoPropertyState(fld, "AllowZeroLength")) & ModuleGuardStructurePart(CStr(fld.ordinalPosition))
        stateText = stateText & ModuleGuardStructurePart(ModuleGuardDaoPropertyState(fld, "DefaultValue")) & ModuleGuardStructurePart(ModuleGuardDaoPropertyState(fld, "ValidationRule")) & ModuleGuardStructurePart(ModuleGuardDaoPropertyState(fld, DAO_DESCRIPTION_PROPERTY_NAME))
        stateText = stateText & ModuleGuardStructurePart(ModuleGuardDaoPropertyState(fld, "DisplayControl")) & ModuleGuardStructurePart(ModuleGuardDaoPropertyState(fld, "RowSourceType")) & ModuleGuardStructurePart(ModuleGuardDaoPropertyState(fld, "RowSource"))
        stateText = stateText & ModuleGuardStructurePart(ModuleGuardDaoPropertyState(fld, "BoundColumn")) & ModuleGuardStructurePart(ModuleGuardDaoPropertyState(fld, "ColumnCount")) & ModuleGuardStructurePart(ModuleGuardDaoPropertyState(fld, "LimitToList"))
    Next fld
    For Each idx In tdf.Indexes
        PumpModuleGuardSetupResponsiveness
        stateText = stateText & ModuleGuardStructurePart("INDEX") & ModuleGuardStructurePart(idx.Name) & ModuleGuardStructurePart(CStr(idx.Primary)) & ModuleGuardStructurePart(CStr(idx.Unique)) & ModuleGuardStructurePart(CStr(idx.IgnoreNulls)) & ModuleGuardStructurePart(CStr(idx.Required))
        For Each indexField In idx.fields
            PumpModuleGuardSetupResponsiveness
            stateText = stateText & ModuleGuardStructurePart(indexField.Name) & ModuleGuardStructurePart(CStr(indexField.Attributes))
        Next indexField
    Next idx
End Sub

Private Sub AppendModuleGuardQueryStructureState(ByRef stateText As String, ByVal qdf As DAO.queryDef)
    Dim fld As DAO.Field
    stateText = stateText & ModuleGuardStructurePart("QUERY") & ModuleGuardStructurePart(qdf.Name) & ModuleGuardStructurePart(CStr(qdf.Type)) & ModuleGuardStructurePart(qdf.sql)
    qdf.fields.Refresh
    For Each fld In qdf.fields
        PumpModuleGuardSetupResponsiveness
        stateText = stateText & ModuleGuardStructurePart("QUERYFIELD") & ModuleGuardStructurePart(fld.Name) & ModuleGuardStructurePart(CStr(fld.Type)) & ModuleGuardStructurePart(CStr(fld.Size)) & ModuleGuardStructurePart(CStr(fld.Attributes))
    Next fld
End Sub

Private Function ModuleGuardStructurePart(ByVal valueText As String) As String
    ModuleGuardStructurePart = CStr(Len(valueText)) & ":" & valueText
End Function

Private Function ModuleGuardDaoPropertyState(ByVal targetObject As Object, ByVal propertyName As String) As String
    On Error GoTo MissingProperty
    Dim propertyValue As Variant
    propertyValue = targetObject.properties(propertyName).value
    If IsNull(propertyValue) Then ModuleGuardDaoPropertyState = "#NULL#" Else ModuleGuardDaoPropertyState = CStr(propertyValue)
    Exit Function
MissingProperty:
    Err.Clear
    ModuleGuardDaoPropertyState = "#MISSING#"
End Function

' =============================================================================================================================================================================
' APPROVED BUILD / MANIFEST TRANSITION VERIFICATION PROCEDURES
' =============================================================================================================================================================================

Private Sub ValidateApprovedBuildDefinitions(ByVal definitions As Collection, Optional ByVal expectedModuleCount As Long = 0, Optional ByVal requireCurrentOwnership As Boolean = True)
    On Error GoTo ErrHandler
    Dim names As Object
    Dim Definition As Variant
    Dim ModuleName As String
    Dim ComponentType As String
    Dim versionID As String
    Dim previousVersionID As String
    Dim timestampText As String
    Dim parsedTimestamp As Date
    If definitions Is Nothing Then Err.Raise MODULE_GUARD_ERROR_BASE + 95, "ValidateApprovedBuildDefinitions", "The managed-header coordinated-build definition collection is missing."
    If expectedModuleCount > 0 Then
        If definitions.count <> expectedModuleCount Then Err.Raise MODULE_GUARD_ERROR_BASE + 96, "ValidateApprovedBuildDefinitions", "The managed-header coordinated-build inventory has the wrong " & DEFAULT_UTILITY_SHORT_NAME & "-owned module count. Expected: " & CStr(expectedModuleCount) & "; Found: " & CStr(definitions.count)
    End If
    ' INTENTIONAL KEYED LOOKUP: names provides keyed membership/lookup over source rows or state already read by ValidateApprovedBuildDefinitions; it is a derived access path, not a second authority.
    Set names = CreateObject("Scripting.Dictionary")
    names.CompareMode = vbBinaryCompare
    For Each Definition In definitions
        PumpModuleGuardSetupResponsiveness
        If UBound(Definition) < 4 Then Err.Raise MODULE_GUARD_ERROR_BASE + 97, "ValidateApprovedBuildDefinitions", "A managed-header coordinated-build definition does not contain the complete five-field contract."
        ModuleName = Trim$(CStr(Definition(0)))
        ComponentType = Trim$(CStr(Definition(1)))
        versionID = Trim$(CStr(Definition(2)))
        previousVersionID = Trim$(CStr(Definition(3)))
        timestampText = Trim$(CStr(Definition(4)))
        If requireCurrentOwnership Then
            If ModuleGuardOwnsComponent(ModuleName) = False Then Err.Raise MODULE_GUARD_ERROR_BASE + 98, "ValidateApprovedBuildDefinitions", "The managed-header coordinated-build inventory contains a component outside " & DEFAULT_UTILITY_SHORT_NAME & " Module Guard ownership." & vbCrLf & "Module: " & DisplayGuardValue(ModuleName)
        End If
        If names.Exists(ModuleName) Then Err.Raise MODULE_GUARD_ERROR_BASE + 99, "ValidateApprovedBuildDefinitions", "The managed-header coordinated-build inventory contains a duplicate module." & vbCrLf & "Module: " & ModuleName
        names.Add ModuleName, True
        Select Case ComponentType
            Case MODULE_GUARD_COMPONENT_STANDARD, MODULE_GUARD_COMPONENT_CLASS, MODULE_GUARD_COMPONENT_FORM, MODULE_GUARD_COMPONENT_DOCUMENT
            Case Else
                Err.Raise MODULE_GUARD_ERROR_BASE + 100, "ValidateApprovedBuildDefinitions", "The managed-header coordinated-build definition contains an unsupported component type." & vbCrLf & "Module: " & ModuleName & vbCrLf & "Component type: " & DisplayGuardValue(ComponentType)
        End Select
        If IsValidChatGPTVersionID(versionID) = False Then Err.Raise MODULE_GUARD_ERROR_BASE + 101, "ValidateApprovedBuildDefinitions", "The managed-header coordinated-build definition contains an invalid current Version ID." & vbCrLf & "Module: " & ModuleName & vbCrLf & "Version ID: " & DisplayGuardValue(versionID)
        If Len(previousVersionID) > 0 And StrComp(previousVersionID, MODULE_GUARD_TRACKING_PREFIX & MODULE_GUARD_NO_VERSION_SUFFIX, vbBinaryCompare) <> 0 And IsValidChatGPTVersionID(previousVersionID) = False Then Err.Raise MODULE_GUARD_ERROR_BASE + 102, "ValidateApprovedBuildDefinitions", "The managed-header coordinated-build definition contains an invalid Previous Version ID." & vbCrLf & "Module: " & ModuleName & vbCrLf & "Previous Version ID: " & DisplayGuardValue(previousVersionID)
        If Len(timestampText) = 0 Or TryParseChatGPTVersionTimestamp(timestampText, parsedTimestamp) = False Then Err.Raise MODULE_GUARD_ERROR_BASE + 103, "ValidateApprovedBuildDefinitions", "The managed-header coordinated-build definition contains a missing or invalid timestamp." & vbCrLf & "Module: " & ModuleName & vbCrLf & "Timestamp: " & DisplayGuardValue(timestampText)
    Next Definition
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If savedErrNumber = 0 Then savedErrNumber = MODULE_GUARD_ERROR_BASE + 109
    If Len(Trim$(savedErrSource)) = 0 Then savedErrSource = "ValidateApprovedBuildDefinitions"
    Err.Raise savedErrNumber, savedErrSource, savedErrDescription
End Sub

Private Sub ValidateDefinitionsModuleBuildParticipation(ByVal db As DAO.Database, ByVal definitions As Collection)
    'Requires the definitions module and approved-build batch identity to advance whenever another guarded module changes; explicit approval does not bypass this structural build-definition rule.
    On Error GoTo ErrHandler
    Dim incomingVersions As Object
    Dim activeVersions As Object
    Dim rs As DAO.Recordset
    Dim Definition As Variant
    Dim ModuleName As String
    Dim incomingDefinitionsVersion As String
    Dim activeDefinitionsVersion As String
    Dim activeDefinitionsBatchID As String
    Dim otherBuildChange As Boolean
    Dim key As Variant
    Set incomingVersions = CreateObject("Scripting.Dictionary")
    incomingVersions.CompareMode = vbTextCompare
    For Each Definition In definitions
        ModuleName = Trim$(CStr(Definition(0)))
        If Len(ModuleName) > 0 Then incomingVersions(ModuleName) = Trim$(CStr(Definition(2)))
    Next Definition
    If incomingVersions.Exists(MODULE_GUARD_DEFINITIONS_MODULE_NAME) = False Then Err.Raise MODULE_GUARD_ERROR_BASE + 154, "ValidateDefinitionsModuleBuildParticipation", "The guarded build does not include the required Module Guard definitions module." & vbCrLf & "Module: " & MODULE_GUARD_DEFINITIONS_MODULE_NAME
    incomingDefinitionsVersion = CStr(incomingVersions(MODULE_GUARD_DEFINITIONS_MODULE_NAME))
    Set activeVersions = CreateObject("Scripting.Dictionary")
    activeVersions.CompareMode = vbTextCompare
    Set rs = db.OpenRecordset("SELECT " & SQL_FIELD_MODULE_NAME & ", " & SQL_FIELD_EXPECTED_VERSION_ID & ", " & SQL_FIELD_UPDATE_BATCH_ID & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionManifest") & "] WHERE [" & FIELD_ACTIVE & "]=True;", dbOpenSnapshot)
    Do While rs.EOF = False
        ModuleName = Trim$(CStr(Nz(rs.fields(FIELD_MODULE_NAME), vbNullString)))
        If Len(ModuleName) > 0 Then
            activeVersions(ModuleName) = Trim$(CStr(Nz(rs.fields(FIELD_EXPECTED_VERSION_ID), vbNullString)))
            If StrComp(ModuleName, MODULE_GUARD_DEFINITIONS_MODULE_NAME, vbTextCompare) = 0 Then
                activeDefinitionsVersion = CStr(activeVersions(ModuleName))
                activeDefinitionsBatchID = Trim$(CStr(Nz(rs.fields(FIELD_UPDATE_BATCH_ID), vbNullString)))
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    'No active definitions row means this is the first build using the split architecture; the newly added definitions module satisfies participation by construction.
    If Len(activeDefinitionsVersion) = 0 Then Exit Sub
    For Each key In incomingVersions.keys
        If StrComp(CStr(key), MODULE_GUARD_DEFINITIONS_MODULE_NAME, vbTextCompare) <> 0 Then
            If activeVersions.Exists(CStr(key)) = False Then
                otherBuildChange = True
                Exit For
            End If
            If StrComp(CStr(incomingVersions(CStr(key))), CStr(activeVersions(CStr(key))), vbBinaryCompare) <> 0 Then
                otherBuildChange = True
                Exit For
            End If
        End If
    Next key
    If otherBuildChange = False Then
        For Each key In activeVersions.keys
            If StrComp(CStr(key), MODULE_GUARD_DEFINITIONS_MODULE_NAME, vbTextCompare) <> 0 Then
                If incomingVersions.Exists(CStr(key)) = False Then
                    otherBuildChange = True
                    Exit For
                End If
            End If
        Next key
    End If
    If otherBuildChange = False Then Exit Sub
    If StrComp(incomingDefinitionsVersion, activeDefinitionsVersion, vbBinaryCompare) = 0 Then Err.Raise MODULE_GUARD_ERROR_BASE + 155, "ValidateDefinitionsModuleBuildParticipation", "A guarded module changed, but the Module Guard definitions module Version ID did not advance." & vbCrLf & "Module: " & MODULE_GUARD_DEFINITIONS_MODULE_NAME & vbCrLf & "Active Version ID: " & DisplayGuardValue(activeDefinitionsVersion)
    If Len(activeDefinitionsBatchID) > 0 And StrComp(MODULE_GUARD_APPROVED_CURRENT_BUILD_BATCH_ID, activeDefinitionsBatchID, vbBinaryCompare) = 0 Then Err.Raise MODULE_GUARD_ERROR_BASE + 156, "ValidateDefinitionsModuleBuildParticipation", "A guarded module changed, but the approved-build batch identity in the Module Guard definitions module did not advance." & vbCrLf & "Batch ID: " & MODULE_GUARD_APPROVED_CURRENT_BUILD_BATCH_ID
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
    Set rs = Nothing
    On Error GoTo 0
    If savedErrNumber = 0 Then savedErrNumber = MODULE_GUARD_ERROR_BASE + 157
    If Len(savedErrSource) = 0 Then savedErrSource = "ValidateDefinitionsModuleBuildParticipation"
    Err.Raise savedErrNumber, savedErrSource, savedErrDescription
End Sub

Private Sub PreflightApprovedBuildTransition(ByVal db As DAO.Database, ByVal definitions As Collection)
    'Validates the current managed-header coordinated build before mutation; stale source branches fail closed against persistent active-manifest/history evidence.
    On Error GoTo ErrHandler
    Dim activeRows As Object
    Dim incomingNames As Object
    Dim historyByKey As Object
    Dim rs As DAO.Recordset
    Dim Definition As Variant
    Dim ModuleName As String
    Dim activeData As Variant
    Dim activeType As String
    Dim activeCurrent As String
    Dim activePrevious As String
    Dim activeTimestampText As String
    Dim activeBaselineAccepted As Boolean
    Dim activeEvidenceDetail As String
    Dim incomingType As String
    Dim incomingCurrent As String
    Dim incomingSourcePrevious As String
    Dim incomingTimestampText As String
    Dim activeTimestampValue As Date
    Dim incomingTimestampValue As Date
    Dim activeTimestampParsed As Boolean
    Dim incomingTimestampParsed As Boolean
    Dim typeTransitionAuthorized As Boolean
    Dim typeTransitionDetail As String
    Dim mismatchCount As Long
    Dim mismatchText As String
    Dim changedCount As Long
    Dim unchangedCount As Long
    Dim addedCount As Long
    Dim acceptedTransitionText As String
    Dim key As Variant
    ValidateApprovedBuildDefinitions definitions
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: activeRows, incomingNames holds planned/recovery work or independently observed validation evidence for PreflightApprovedBuildTransition; keep it distinct from the
    ' source authority.
    Set activeRows = CreateObject("Scripting.Dictionary")
    activeRows.CompareMode = vbTextCompare
    Set incomingNames = CreateObject("Scripting.Dictionary")
    incomingNames.CompareMode = vbTextCompare
    Set rs = db.OpenRecordset( _
        "SELECT " & SQL_FIELD_MODULE_NAME & ", " & SQL_FIELD_COMPONENT_TYPE & ", " & SQL_FIELD_EXPECTED_VERSION_ID & ", " & SQL_FIELD_EXPECTED_PREVIOUS_VERSION_ID & ", " & _
        SQL_FIELD_EXPECTED_VERSION_TIMESTAMP & ", " & SQL_FIELD_BASELINE_ACCEPTED & " " & _
        "FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionManifest") & "] WHERE [" & FIELD_ACTIVE & "]=True;", dbOpenSnapshot)
    Do While Not rs.EOF
        PumpModuleGuardSetupResponsiveness
        ModuleName = CStr(rs.fields(FIELD_MODULE_NAME))
        If ModuleGuardOwnsComponent(ModuleName) Then
            If activeRows.Exists(ModuleName) Then
                rs.Close
                Set rs = Nothing
                Err.Raise MODULE_GUARD_ERROR_BASE + 74, "PreflightApprovedBuildTransition", "Duplicate active manifest row: " & ModuleName
            End If
            activeRows.Add ModuleName, Array( _
                Trim$(CStr(Nz(rs.fields(FIELD_COMPONENT_TYPE), vbNullString))), _
                Trim$(CStr(Nz(rs!expectedVersionID, vbNullString))), _
                Trim$(CStr(Nz(rs!expectedPreviousVersionID, vbNullString))), _
                ManifestTimestampText(rs.fields(FIELD_EXPECTED_VERSION_TIMESTAMP)), _
                CBool(Nz(rs!baselineAccepted, False)))
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    If activeRows.count = 0 Then Exit Sub
    Set historyByKey = CreateObject("Scripting.Dictionary")
    historyByKey.CompareMode = vbBinaryCompare
    CaptureModuleGuardHistoryIndex db, historyByKey
    LogModuleGuardPreflightBoundary "PreflightApprovedBuildTransition", "Module Guard approved-build preflight: append-only history loaded once for all transition checks.", "Indexed history rows: " & CStr(historyByKey.count)
    For Each Definition In definitions
        PumpModuleGuardSetupResponsiveness
        ModuleName = CStr(Definition(0))
        If incomingNames.Exists(ModuleName) Then
            Err.Raise MODULE_GUARD_ERROR_BASE + 75, "PreflightApprovedBuildTransition", "Duplicate incoming approved-build module: " & ModuleName
        End If
        incomingNames.Add ModuleName, True
        incomingType = Trim$(CStr(Definition(1)))
        incomingCurrent = Trim$(CStr(Definition(2)))
        incomingSourcePrevious = Trim$(CStr(Definition(3)))
        incomingTimestampText = Trim$(CStr(Definition(4)))
        If activeRows.Exists(ModuleName) = False Then
            addedCount = addedCount + 1
            changedCount = changedCount + 1
            AppendGuardDetail acceptedTransitionText, ModuleName & ": new " & DEFAULT_UTILITY_SHORT_NAME & "-owned module -> " & incomingCurrent & " (authorized by the exact managed-header build)"
        Else
            activeData = activeRows(ModuleName)
            activeType = CStr(activeData(0))
            activeCurrent = CStr(activeData(1))
            activePrevious = CStr(activeData(2))
            activeTimestampText = CStr(activeData(3))
            activeBaselineAccepted = CBool(activeData(4))
            activeEvidenceDetail = vbNullString
            activeTimestampParsed = True
            incomingTimestampParsed = True
            typeTransitionAuthorized = False
            typeTransitionDetail = vbNullString
            If Len(activeTimestampText) > 0 Then activeTimestampParsed = TryParseChatGPTVersionTimestamp(activeTimestampText, activeTimestampValue)
            If Len(incomingTimestampText) > 0 Then incomingTimestampParsed = TryParseChatGPTVersionTimestamp(incomingTimestampText, incomingTimestampValue)
            If StrComp(activeType, incomingType, vbTextCompare) <> 0 Then
                typeTransitionAuthorized = ApprovedComponentTypeTransitionIsAuthorized(Definition, activeType, activeCurrent, typeTransitionDetail)
            End If
            If Len(activeCurrent) = 0 Then
                mismatchCount = mismatchCount + 1
                AppendGuardDetail mismatchText, ApprovedDefinitionTransitionDetail(ModuleName, "The active manifest Current ID is blank.", activeCurrent, activePrevious, incomingCurrent, incomingSourcePrevious)
            ElseIf StrComp(activeType, incomingType, vbTextCompare) <> 0 And typeTransitionAuthorized = False Then
                mismatchCount = mismatchCount + 1
                'Fail closed on type drift unless the incoming header directly follows the active approved version and persistent history proves the old component identity.
                AppendGuardDetail mismatchText, ApprovedDefinitionTransitionDetail(ModuleName, "Component type differs from the active approved manifest and the exact approved predecessor does not authorize that migration. Active type=" & DisplayGuardValue(activeType) & "; Incoming type=" & DisplayGuardValue(incomingType), activeCurrent, activePrevious, incomingCurrent, incomingSourcePrevious) & vbCrLf & typeTransitionDetail
            ElseIf activeBaselineAccepted = False Then
                mismatchCount = mismatchCount + 1
                AppendGuardDetail mismatchText, ApprovedDefinitionTransitionDetail(ModuleName, "The active manifest row was not marked as an accepted baseline.", activeCurrent, activePrevious, incomingCurrent, incomingSourcePrevious)
            ElseIf StrComp(incomingCurrent, activeCurrent, vbBinaryCompare) = 0 Then
                If StrComp(incomingSourcePrevious, activePrevious, vbBinaryCompare) <> 0 Then
                    mismatchCount = mismatchCount + 1
                    AppendGuardDetail mismatchText, ApprovedDefinitionTransitionDetail(ModuleName, "The unchanged Current Version ID has a different managed-header Previous Version ID than the active manifest. Reusing a Version ID with different lineage is not allowed.", activeCurrent, activePrevious, incomingCurrent, incomingSourcePrevious)
                ElseIf StrComp(incomingTimestampText, activeTimestampText, vbBinaryCompare) <> 0 Then
                    mismatchCount = mismatchCount + 1
                    AppendGuardDetail mismatchText, ApprovedDefinitionTransitionDetail(ModuleName, "The unchanged Current Version ID has a different Version Timestamp than the active manifest. Reusing a Version ID with different timestamp metadata is not allowed.", activeCurrent, activePrevious, incomingCurrent, incomingSourcePrevious)
                Else
                    unchangedCount = unchangedCount + 1
                End If
            ElseIf ActiveManifestHasRegisteredHistoryEvidence(historyByKey, ModuleName, activeType, activeCurrent, activePrevious, activeTimestampText, activeEvidenceDetail) = False Then
                mismatchCount = mismatchCount + 1
                AppendGuardDetail mismatchText, ApprovedDefinitionTransitionDetail(ModuleName, "The changed module's active manifest row is not backed by matching append-only registration history.", activeCurrent, activePrevious, incomingCurrent, incomingSourcePrevious) & vbCrLf & activeEvidenceDetail
            ElseIf Len(activeTimestampText) = 0 Then
                mismatchCount = mismatchCount + 1
                AppendGuardDetail mismatchText, ApprovedDefinitionTransitionDetail(ModuleName, "The changed active manifest row has no timestamp, so rollback safety cannot be proven.", activeCurrent, activePrevious, incomingCurrent, incomingSourcePrevious)
            ElseIf Len(incomingTimestampText) = 0 Then
                mismatchCount = mismatchCount + 1
                AppendGuardDetail mismatchText, ApprovedDefinitionTransitionDetail(ModuleName, "The incoming build omits a timestamp required to prove that it is not a rollback.", activeCurrent, activePrevious, incomingCurrent, incomingSourcePrevious)
            ElseIf activeTimestampParsed = False Then
                mismatchCount = mismatchCount + 1
                AppendGuardDetail mismatchText, ApprovedDefinitionTransitionDetail(ModuleName, "The active manifest timestamp is not in the required ChatGPT Version Timestamp format.", activeCurrent, activePrevious, incomingCurrent, incomingSourcePrevious)
            ElseIf incomingTimestampParsed = False Then
                mismatchCount = mismatchCount + 1
                AppendGuardDetail mismatchText, ApprovedDefinitionTransitionDetail(ModuleName, "The managed-header coordinated-build timestamp is not in the required ChatGPT Version Timestamp format.", activeCurrent, activePrevious, incomingCurrent, incomingSourcePrevious)
            ElseIf incomingTimestampValue < activeTimestampValue Then
                mismatchCount = mismatchCount + 1
                AppendGuardDetail mismatchText, ApprovedDefinitionTransitionDetail(ModuleName, "Rollback rejected: the incoming version timestamp is older than the active approved version timestamp.", activeCurrent, activePrevious, incomingCurrent, incomingSourcePrevious)
            ElseIf incomingTimestampValue = activeTimestampValue Then
                mismatchCount = mismatchCount + 1
                AppendGuardDetail mismatchText, ApprovedDefinitionTransitionDetail(ModuleName, "Different Version IDs share the same version timestamp, so deployment order cannot be proven safely.", activeCurrent, activePrevious, incomingCurrent, incomingSourcePrevious)
            ElseIf Len(incomingSourcePrevious) = 0 Or StrComp(incomingSourcePrevious, MODULE_GUARD_TRACKING_PREFIX & MODULE_GUARD_NO_VERSION_SUFFIX, vbBinaryCompare) = 0 Then
                mismatchCount = mismatchCount + 1
                AppendGuardDetail mismatchText, ApprovedDefinitionTransitionDetail(ModuleName, "A changed guarded module must identify the database's active approved Version ID as its managed-header Previous Version ID.", activeCurrent, activePrevious, incomingCurrent, incomingSourcePrevious)
            ElseIf StrComp(incomingSourcePrevious, activeCurrent, vbBinaryCompare) <> 0 Then
                mismatchCount = mismatchCount + 1
                AppendGuardDetail mismatchText, ApprovedDefinitionTransitionDetail(ModuleName, "Baseline lineage mismatch rejected: the incoming managed-header Previous Version ID does not equal the database's active approved Current Version ID. The supplied source was not edited from the currently approved baseline, or the database is on a different build. Module Guard does not infer, skip, or authorize intermediate module versions.", activeCurrent, activePrevious, incomingCurrent, incomingSourcePrevious)
            Else
                changedCount = changedCount + 1
                AppendGuardDetail acceptedTransitionText, ModuleName & ": " & activeCurrent & " -> " & incomingCurrent & _
                    IIf(StrComp(activeType, incomingType, vbTextCompare) = 0, vbNullString, " (direct-predecessor component-type migration: " & activeType & " -> " & incomingType & ")")
            End If
        End If
    Next Definition
    For Each key In activeRows.keys
        PumpModuleGuardSetupResponsiveness
        If incomingNames.Exists(CStr(key)) = False Then
            activeData = activeRows(CStr(key))
            mismatchCount = mismatchCount + 1
            AppendGuardDetail mismatchText, ApprovedDefinitionTransitionDetail(CStr(key), "The active guarded module is missing from the proposed build; removal was not explicitly authorized.", CStr(activeData(1)), CStr(activeData(2)), vbNullString, vbNullString)
        End If
    Next key
    If mismatchCount > 0 Then
        LogModuleGuardMessage "Approved-build deployment preflight failed; no history or manifest write started.", _
            "Mismatch count: " & CStr(mismatchCount) & vbCrLf & mismatchText, _
            TRACE_LEVEL_ERROR, "PreflightApprovedBuildTransition", "Validate exact approved deployment"
        Err.Raise MODULE_GUARD_ERROR_BASE + 72, "PreflightApprovedBuildTransition", _
            "Approved-build deployment preflight found " & CStr(mismatchCount) & " invalid transition(s). No history or manifest rows were written." & vbCrLf & vbCrLf & mismatchText
    End If
    LogModuleGuardMessage "Approved-build deployment preflight passed before mutation.", _
        "Active " & DEFAULT_UTILITY_SHORT_NAME & "-owned manifest rows: " & CStr(activeRows.count) & vbCrLf & _
        "Incoming exact approved rows: " & CStr(incomingNames.count) & vbCrLf & _
        "Changed modules: " & CStr(changedCount) & vbCrLf & _
        "Unchanged modules: " & CStr(unchangedCount) & vbCrLf & _
        "New owned modules: " & CStr(addedCount) & vbCrLf & _
        "Rule: unchanged modules must retain the active approved Current Version ID. Every changed guarded module must name that exact active Current Version ID as its commented Previous Version ID, must have a different Current Version ID, and must have a rollback-safe timestamp. Module Guard does not infer or skip intermediate module versions and accepts no source-carried module Version ID exception. Component-type migration is permitted only on that same direct predecessor and remains subject to persistent manifest/history proof." & _
        IIf(Len(acceptedTransitionText) = 0, vbNullString, vbCrLf & vbCrLf & "Accepted deployment transitions:" & vbCrLf & acceptedTransitionText), _
        TRACE_LEVEL_INFO, "PreflightApprovedBuildTransition", "Validate exact approved deployment"
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
    Set rs = Nothing
    On Error GoTo 0
    If savedErrNumber = 0 Then savedErrNumber = MODULE_GUARD_ERROR_BASE + 73
    If Len(savedErrSource) = 0 Then savedErrSource = "PreflightApprovedBuildTransition"
    Err.Raise savedErrNumber, savedErrSource, savedErrDescription
End Sub

Private Sub RunInteractiveModuleUpdateGuardVerification()
    Dim failedErrNumber As Long
    Dim failedErrSource As String
    Dim failedErrDescription As String
    Dim verificationDetails As String
    Dim installedVersionSummary As String
    If TryPrepareAndVerifyModuleUpdateGuard(failedErrNumber, failedErrSource, failedErrDescription, verificationDetails, installedVersionSummary:=installedVersionSummary) Then
        LogModuleGuardPreflightBoundary "RunInteractiveModuleUpdateGuardVerification", "Module Guard interactive verification: success dialog opening."
        UtilityMsgBox DEFAULT_UTILITY_LONG_NAME & " module update guard verification completed successfully." & vbCrLf & vbCrLf & installedVersionSummary, vbInformation, DEFAULT_UTILITY_LONG_NAME & " Module Update Guard"
        LogModuleGuardPreflightBoundary "RunInteractiveModuleUpdateGuardVerification", "Module Guard interactive verification: success dialog returned."
        Exit Sub
    End If
    Err.Raise failedErrNumber, "VerifyModuleUpdateGuard", failedErrDescription
End Sub

' Purpose: Verifies module update guard for suppress failure reporting.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
'   verificationDetails (Output Optional; String): Receives verification details used for caller diagnostics, verification output, or user-facing reporting. Default: vbNullString.
'   suppressFailureReporting (Input Optional; Boolean): True suppresses failure reporting; False leaves the normal behavior enabled. Default: False.
'   installedVersionSummary (Output Optional; String): Receives the installed version summary returned by `InstalledModuleVersionSummary` after this procedure completes its validation/resolution work. Default:
'     vbNullString.
' Returns: None (Sub). Results are returned through the documented ByRef parameters and/or the routine's database, UI, or module-state side effects.
Public Sub VerifyModuleUpdateGuard(Optional ByVal showMessage As Boolean = True, Optional ByRef verificationDetails As String = vbNullString, Optional ByVal suppressFailureReporting As Boolean = False, Optional ByRef installedVersionSummary As String = vbNullString)
    On Error GoTo ErrHandler
    Dim verificationStage As String
    installedVersionSummary = vbNullString
    If showMessage And suppressFailureReporting = False And mModuleGuardRegistrationTransactionOwned = False Then
        verificationStage = "Run interactive module update guard verification"
        RunInteractiveModuleUpdateGuardVerification
        Exit Sub
    End If
    Dim db As DAO.Database
    Dim ws As DAO.workspace
    Dim initialBaselineDefinitions As Collection
    Dim manifest As DAO.Recordset
    Dim verification As DAO.Recordset
    Dim componentsByName As Object
    Dim timestampTextByName As Object
    Dim typeByName As Object
    Dim versionByName As Object
    Dim previousByName As Object
    Dim timestampByName As Object
    Dim headerNameByName As Object
    Dim versionCounts As Object
    Dim manifestNames As Object
    Dim historyByKey As Object
    Dim inventoryEvidence As Collection
    ' INTENTIONAL CONSOLIDATED VERIFICATION INDEXES: These keyed views serve distinct membership, metadata, duplicate, and history checks over one captured verification state; they avoid repeated project enumeration
    ' without creating parallel expected-state authorities.
    Dim key As Variant
    Dim verificationRunID As String
    Dim verifiedOn As Date
    Dim resultCode As String
    Dim resultNotes As String
    Dim mismatchCount As Long
    Dim matchCount As Long
    Dim unregisteredCount As Long
    Dim requiredCount As Long
    Dim installedCount As Long
    Dim manifestCount As Long
    Dim expectedTimestamp As Variant
    Dim installedTimestamp As Variant
    Dim expectedVersion As String
    Dim installedVersion As String
    Dim expectedPrevious As String
    Dim installedPrevious As String
    Dim expectedType As String
    Dim installedType As String
    Dim ModuleName As String
    Dim expectedHeaderName As String
    Dim batchID As String
    Dim isRequired As Boolean
    Dim transactionStarted As Boolean
    Dim activeManifestCount As Long
    verificationStage = "Capture one installed module inventory and validate identities"
    Set db = currentDb
    Set ws = DBEngine.Workspaces(0)
    ' INTENTIONAL KEYED LOOKUP: componentsByName, timestampTextByName, typeByName, versionByName, previousByName, timestampByName, headerNameByName, versionCounts, manifestNames, historyByKey provides keyed
    ' membership/lookup over source rows or state already read by VerifyModuleUpdateGuard; it is a derived access path, not a second authority.
    Set manifestNames = CreateObject("Scripting.Dictionary")
    Set historyByKey = CreateObject("Scripting.Dictionary")
    ConfigureTextDictionary manifestNames
    historyByKey.CompareMode = vbBinaryCompare
    verificationStage = "Capture or consume installed module inventory"
    If Not mModuleGuardInstalledInventoryEvidence Is Nothing Then Set inventoryEvidence = mModuleGuardInstalledInventoryEvidence
    If inventoryEvidence Is Nothing Then
        Set initialBaselineDefinitions = InitialModuleGuardBaselineDefinitions()
        LogModuleGuardPreflightBoundary "VerifyModuleUpdateGuard", "Module Guard verification: one-pass installed-module inventory and header-identity validation starting."
        Set inventoryEvidence = CaptureInstalledModuleInventoryEvidence(initialBaselineDefinitions)
        If mModuleGuardPreflightActive Or mModuleGuardRegistrationTransactionOwned Then Set mModuleGuardInstalledInventoryEvidence = inventoryEvidence
        LogModuleGuardPreflightBoundary "VerifyModuleUpdateGuard", "Module Guard verification: one-pass installed-module inventory and header-identity validation completed."
    Else
        Set initialBaselineDefinitions = inventoryEvidence("ExpectedDefinitions")
        LogModuleGuardPreflightBoundary "VerifyModuleUpdateGuard", "Module Guard verification: authoritative installed-module inventory and canonical definitions reused; duplicate VBProject/header scan skipped."
    End If
    LoadInstalledModuleInventoryEvidence inventoryEvidence, componentsByName, timestampTextByName, typeByName, versionByName, previousByName, timestampByName, headerNameByName, versionCounts, installedCount
    LogModuleGuardPreflightBoundary "VerifyModuleUpdateGuard", "Module Guard verification: installed-module inventory ready.", "Project components enumerated: " & CStr(installedCount) & vbCrLf & DEFAULT_UTILITY_SHORT_NAME & "-owned modules inspected: " & CStr(componentsByName.count)
    verificationStage = "Prepare active-manifest verification"
    verificationRunID = Format$(Now, "yyyymmddhhnnss") & "-" & Format$(CLng(Timer * 1000) Mod 1000000, "000000")
    verifiedOn = Now
    activeManifestCount = DCount("*", ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionManifest"), "[" & FIELD_ACTIVE & "]=True")
    verificationStage = "Begin verification transaction"
    LogModuleGuardPreflightBoundary "VerifyModuleUpdateGuard", "Module Guard verification: transaction starting.", "Active manifest rows: " & CStr(activeManifestCount)
    ws.BeginTrans
    transactionStarted = True
    If mModuleGuardRegistrationTransactionOwned Then mModuleGuardNestedVerificationTransactionActive = True
    verificationStage = "Open and traverse update history"
    LogModuleGuardPreflightBoundary "VerifyModuleUpdateGuard", "Module Guard verification: history index recordset opening."
    CaptureModuleGuardHistoryIndex db, historyByKey
    LogModuleGuardPreflightBoundary "VerifyModuleUpdateGuard", "Module Guard verification: history index traversal completed.", "Indexed history rows: " & CStr(historyByKey.count)
    verificationStage = "Open active manifest and verification snapshot"
    LogModuleGuardPreflightBoundary "VerifyModuleUpdateGuard", "Module Guard verification: active manifest recordset opening.", "Query source: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionManifest")
    Set manifest = db.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionManifest") & "] WHERE [" & FIELD_ACTIVE & "]=True ORDER BY " & OrderByPolicyText("MODULE_GUARD_MANIFEST_TRAVERSAL"), dbOpenDynaset)
    LogModuleGuardPreflightBoundary "VerifyModuleUpdateGuard", "Module Guard verification: active manifest recordset opened."
    LogModuleGuardPreflightBoundary "VerifyModuleUpdateGuard", "Module Guard verification: verification snapshot recordset opening.", "Table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionVerification")
    Set verification = db.OpenRecordset(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionVerification"), dbOpenDynaset)
    LogModuleGuardPreflightBoundary "VerifyModuleUpdateGuard", "Module Guard verification: verification snapshot recordset opened."
    verificationStage = "Compare active manifest with installed modules"
    LogModuleGuardPreflightBoundary "VerifyModuleUpdateGuard", "Module Guard verification: active manifest comparison loop starting.", "Rows expected: " & CStr(activeManifestCount)
    Do While manifest.EOF = False
        PumpModuleGuardSetupResponsiveness
        manifestCount = manifestCount + 1
        ModuleName = Trim$(Nz(manifest.fields(FIELD_MODULE_NAME), vbNullString))
        LogModuleGuardPreflightBoundary "VerifyModuleUpdateGuard", "Module Guard verification: manifest row " & CStr(manifestCount) & " of " & CStr(activeManifestCount) & ".", "Module: " & DisplayGuardValue(ModuleName)
        expectedType = Trim$(Nz(manifest.fields(FIELD_COMPONENT_TYPE), vbNullString))
        expectedVersion = Trim$(Nz(manifest!expectedVersionID, vbNullString))
        expectedPrevious = Trim$(Nz(manifest!expectedPreviousVersionID, vbNullString))
        expectedTimestamp = manifest.fields(FIELD_EXPECTED_VERSION_TIMESTAMP)
        isRequired = CBool(Nz(manifest!isRequired, True))
        batchID = Trim$(Nz(manifest!updateBatchID, vbNullString))
        manifestNames(LCase$(ModuleName)) = True
        If isRequired Then requiredCount = requiredCount + 1
        resultCode = MODULE_GUARD_RESULT_MATCH
        resultNotes = vbNullString
        installedType = vbNullString
        installedVersion = vbNullString
        installedPrevious = vbNullString
        installedTimestamp = Null
        expectedHeaderName = vbNullString
        If componentsByName.Exists(ModuleName) = False Then
            If isRequired Then resultCode = MODULE_GUARD_RESULT_MISSING Else resultCode = MODULE_GUARD_RESULT_MISMATCH
            AppendGuardDetail resultNotes, "The expected component is not installed."
        Else
            installedType = CStr(typeByName(ModuleName))
            installedVersion = CStr(versionByName(ModuleName))
            installedPrevious = CStr(previousByName(ModuleName))
            installedTimestamp = timestampByName(ModuleName)
            expectedHeaderName = CStr(headerNameByName(ModuleName))
            If Len(expectedHeaderName) = 0 Or StrComp(expectedHeaderName, ModuleName, vbBinaryCompare) <> 0 Then
                resultCode = MODULE_GUARD_RESULT_HEADER_MISMATCH
                AppendGuardDetail resultNotes, "Module identity mismatch."
                AppendGuardDetail resultNotes, "Actual component name: " & ModuleName
                AppendGuardDetail resultNotes, "Header VB_Name: " & IIf(Len(expectedHeaderName) = 0, "(blank)", expectedHeaderName)
                AppendGuardDetail resultNotes, "Manifest module name: " & ModuleName
            Else
                If StrComp(installedType, expectedType, vbTextCompare) <> 0 Then resultCode = MODULE_GUARD_RESULT_MISMATCH: AppendGuardDetail resultNotes, "Component type mismatch. Expected=" & expectedType & "; Installed=" & installedType
                If IsValidChatGPTVersionID(installedVersion) = False Then resultCode = MODULE_GUARD_RESULT_INVALID_ID: AppendGuardDetail resultNotes, "The installed ChatGPT Version ID is missing, malformed, or a placeholder."
                If StrComp(installedVersion, expectedVersion, vbBinaryCompare) <> 0 Then resultCode = MODULE_GUARD_RESULT_MISMATCH: AppendGuardDetail resultNotes, "Version mismatch. Expected=" & expectedVersion & "; Installed=" & installedVersion
                If Len(expectedPrevious) > 0 And StrComp(installedPrevious, expectedPrevious, vbBinaryCompare) <> 0 Then resultCode = MODULE_GUARD_RESULT_BROKEN_CHAIN: AppendGuardDetail resultNotes, "Previous-version mismatch. Expected=" & expectedPrevious & "; Installed=" & installedPrevious
            If IsNull(expectedTimestamp) = False Then
                If IsNull(installedTimestamp) Then
                    resultCode = MODULE_GUARD_RESULT_HEADER_MISMATCH
                    AppendGuardDetail resultNotes, "The installed header timestamp is missing or invalid. Timestamp=" & IIf(Len(CStr(timestampTextByName(ModuleName))) = 0, "(blank)", CStr(timestampTextByName(ModuleName))) & "; Required format=dddd m/d/yy h:nn AM/PM"
                ElseIf FormatChatGPTVersionTimestamp(CDate(installedTimestamp)) <> FormatChatGPTVersionTimestamp(CDate(expectedTimestamp)) Then
                    resultCode = MODULE_GUARD_RESULT_HEADER_MISMATCH
                    AppendGuardDetail resultNotes, "Timestamp mismatch. Expected=" & FormatChatGPTVersionTimestamp(CDate(expectedTimestamp)) & "; Installed=" & FormatChatGPTVersionTimestamp(CDate(installedTimestamp))
                End If
            End If
            If IsValidChatGPTVersionID(installedVersion) Then
                If CLng(versionCounts(installedVersion)) > 1 Then resultCode = MODULE_GUARD_RESULT_DUPLICATE_ID: AppendGuardDetail resultNotes, "The installed ChatGPT Version ID is claimed by more than one component."
            End If
            If ModuleGuardHistoryChainValid(historyByKey, ModuleName, expectedVersion, expectedPrevious, CBool(Nz(manifest!baselineAccepted, False)), resultNotes) = False Then resultCode = MODULE_GUARD_RESULT_BROKEN_CHAIN
            End If
        End If
        AddVerificationRow verification, verificationRunID, verifiedOn, ModuleName, expectedType, isRequired, True, expectedVersion, installedVersion, expectedPrevious, installedPrevious, expectedTimestamp, installedTimestamp, resultCode, resultNotes, batchID
        UpdateManifestVerificationRow manifest, verifiedOn, resultCode, resultNotes
        If resultCode = MODULE_GUARD_RESULT_MATCH Then matchCount = matchCount + 1 Else mismatchCount = mismatchCount + 1
        manifest.MoveNext
    Loop
    LogModuleGuardPreflightBoundary "VerifyModuleUpdateGuard", "Module Guard verification: active manifest comparison loop completed.", "Rows checked: " & CStr(manifestCount) & vbCrLf & "Matches: " & CStr(matchCount) & vbCrLf & "Mismatches: " & CStr(mismatchCount)
    verificationStage = "Check unregistered installed guarded modules"
    LogModuleGuardPreflightBoundary "VerifyModuleUpdateGuard", "Module Guard verification: unregistered-module loop starting.", "Installed " & DEFAULT_UTILITY_SHORT_NAME & "-owned modules: " & CStr(componentsByName.count)
    For Each key In componentsByName.keys
        PumpModuleGuardSetupResponsiveness
        ModuleName = CStr(key)
        If ModuleGuardOwnsComponent(ModuleName) Then
            installedVersion = CStr(versionByName(ModuleName))
            If IsValidChatGPTVersionID(installedVersion) Then
                If manifestNames.Exists(LCase$(ModuleName)) = False Then
                    resultCode = MODULE_GUARD_RESULT_UNREGISTERED
                    resultNotes = "The installed tracked component is not registered in the active expected-current manifest."
                    AddVerificationRow verification, verificationRunID, verifiedOn, ModuleName, CStr(typeByName(ModuleName)), False, False, vbNullString, installedVersion, vbNullString, CStr(previousByName(ModuleName)), Null, timestampByName(ModuleName), resultCode, resultNotes, vbNullString
                    mismatchCount = mismatchCount + 1
                    unregisteredCount = unregisteredCount + 1
                End If
            End If
        End If
    Next key
    LogModuleGuardPreflightBoundary "VerifyModuleUpdateGuard", "Module Guard verification: unregistered-module loop completed.", "Unregistered modules: " & CStr(unregisteredCount)
    verificationStage = "Close verification recordsets and commit"
    LogModuleGuardPreflightBoundary "VerifyModuleUpdateGuard", "Module Guard verification: recordset cleanup and transaction commit starting."
    verification.Close
    Set verification = Nothing
    manifest.Close
    Set manifest = Nothing
    ws.CommitTrans
    transactionStarted = False
    If mModuleGuardNestedVerificationTransactionActive Then mModuleGuardNestedVerificationTransactionActive = False
    LogModuleGuardPreflightBoundary "VerifyModuleUpdateGuard", "Module Guard verification: transaction committed.", "Verification run: " & verificationRunID
    verificationStage = "Evaluate verification result"
    verificationDetails = "Verification run: " & verificationRunID & vbCrLf & "Manifest rows checked: " & CStr(manifestCount) & vbCrLf & "Required modules: " & CStr(requiredCount) & vbCrLf & "Project components enumerated: " & CStr(installedCount) & vbCrLf & DEFAULT_UTILITY_SHORT_NAME & "-owned modules inspected: " & CStr(componentsByName.count) & vbCrLf & "Matches: " & CStr(matchCount) & vbCrLf & "Mismatches: " & CStr(mismatchCount) & vbCrLf & "Unregistered tracked modules: " & CStr(unregisteredCount)
    If mismatchCount > 0 Then
        If suppressFailureReporting Then
            LogModuleGuardPreflightBoundary "VerifyModuleUpdateGuard", "Module Guard verification found a caller-managed manifest mismatch.", "Mismatches: " & CStr(mismatchCount) & vbCrLf & "The caller will validate whether the installed modules are the exact managed-header coordinated build.", TRACE_LEVEL_INFO
        Else
            LogModuleGuardPreflightBoundary "VerifyModuleUpdateGuard", "Module Guard verification: mismatch evaluation starting.", "Mismatches: " & CStr(mismatchCount), TRACE_LEVEL_WARNING
        End If
        Dim firstMismatchDetails As String
        firstMismatchDetails = FirstModuleGuardMismatchSummary(db, verificationRunID)
        verificationDetails = verificationDetails & vbCrLf & vbCrLf & LatestModuleGuardMismatchSummary(db, verificationRunID)
        Err.Raise MODULE_GUARD_ERROR_BASE + 1, "VerifyModuleUpdateGuard", firstMismatchDetails
    End If
    installedVersionSummary = InstalledModuleVersionSummary(initialBaselineDefinitions, versionByName)
    LogModuleGuardPreflightBoundary "VerifyModuleUpdateGuard", "Module Guard verification completed successfully.", verificationDetails, TRACE_LEVEL_INFO
    If showMessage Then
        LogModuleGuardPreflightBoundary "VerifyModuleUpdateGuard", "Module Guard verification: success dialog opening."
        UtilityMsgBox DEFAULT_UTILITY_LONG_NAME & " module update guard verification completed successfully." & vbCrLf & vbCrLf & installedVersionSummary, vbInformation, DEFAULT_UTILITY_LONG_NAME & " Module Update Guard"
        LogModuleGuardPreflightBoundary "VerifyModuleUpdateGuard", "Module Guard verification: success dialog returned."
    End If
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If Len(verificationStage) > 0 Then savedErrDescription = "Verification stage: " & verificationStage & vbCrLf & savedErrDescription
    Dim cleanupDetails As String
    CloseModuleGuardRecordset verification, "verification snapshot recordset", cleanupDetails
    Set verification = Nothing
    CloseModuleGuardRecordset manifest, "manifest recordset", cleanupDetails
    Set manifest = Nothing
    If transactionStarted Then
        If mModuleGuardRegistrationTransactionOwned Then
            If RollbackRegistrationNestedVerificationTransaction(ws, cleanupDetails) Then transactionStarted = False
        ElseIf RollbackStandaloneVerificationTransaction(ws, cleanupDetails) Then
            transactionStarted = False
        End If
    End If
    If Len(cleanupDetails) > 0 Then savedErrDescription = savedErrDescription & vbCrLf & vbCrLf & cleanupDetails
    If savedErrNumber = 0 Then savedErrNumber = MODULE_GUARD_ERROR_BASE + 2
    Dim popupText As String
    popupText = DEFAULT_UTILITY_LONG_NAME & " module update verification failed." & vbCrLf & vbCrLf & _
                "Err.Number: " & CStr(savedErrNumber) & vbCrLf & _
                "Err.Source: " & savedErrSource & vbCrLf & _
                "Err.Description:" & vbCrLf & savedErrDescription
    If Len(Trim$(verificationDetails)) > 0 Then popupText = popupText & vbCrLf & vbCrLf & "Verification summary:" & vbCrLf & verificationDetails
    Dim failureTranscript As String
    failureTranscript = BuildModuleGuardVerificationFailureTranscript(popupText)
    verificationDetails = failureTranscript
    If suppressFailureReporting = False Then
        Debug.Print failureTranscript
        If mModuleGuardRegistrationTransactionOwned = False Then PersistModuleGuardVerificationFailureTranscript failureTranscript, savedErrNumber, savedErrSource
    End If
    If suppressFailureReporting Then
        LogModuleGuardPreflightBoundary "VerifyModuleUpdateGuard", "Module Guard verification returned a caller-managed result.", "Stage: " & verificationStage & vbCrLf & "Result number: " & CStr(savedErrNumber) & vbCrLf & "Result source: " & savedErrSource & vbCrLf & "Result description: " & savedErrDescription, TRACE_LEVEL_INFO
    Else
        LogModuleGuardPreflightBoundary "VerifyModuleUpdateGuard", "Module Guard verification terminated with failure.", "Stage: " & verificationStage & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, TRACE_LEVEL_ERROR
    End If
    If showMessage Then
        LogModuleGuardPreflightBoundary "VerifyModuleUpdateGuard", "Module Guard verification: failure dialog opening."
        CoreShowErrorMessage popupText, vbCritical, DEFAULT_UTILITY_LONG_NAME & " Module Update Guard", "VerifyModuleUpdateGuard", "Core"
        LogModuleGuardPreflightBoundary "VerifyModuleUpdateGuard", "Module Guard verification: failure dialog returned."
        Exit Sub
    End If
    Err.Raise savedErrNumber, "VerifyModuleUpdateGuard", savedErrDescription
End Sub

' =============================================================================================================================================================================
' INSTALLED MODULE INVENTORY / FINGERPRINT VERIFICATION PROCEDURES
' =============================================================================================================================================================================

Private Sub EnsureModuleGuardOwnedComponentLookup(Optional ByVal baselineDefinitions As Variant)
    'Builds the fast ownership lookup from the fixed guarded-module roster supplied by x__UTIL_ADC_MODULE_GUARD_DEFINITIONS; no module Version ID is embedded here.
    Dim definitions As Collection
    Dim Definition As Variant
    Dim ModuleName As String
    If Not mModuleGuardOwnedComponentLookup Is Nothing Then Exit Sub
    Set mModuleGuardOwnedComponentLookup = CreateObject("Scripting.Dictionary")
    mModuleGuardOwnedComponentLookup.CompareMode = vbTextCompare
    If IsObject(baselineDefinitions) Then Set definitions = baselineDefinitions
    If definitions Is Nothing Then Set definitions = InitialModuleGuardBaselineDefinitions()
    For Each Definition In definitions
        ModuleName = Trim$(CStr(Definition(0)))
        If Len(ModuleName) > 0 Then
            If mModuleGuardOwnedComponentLookup.Exists(ModuleName) = False Then mModuleGuardOwnedComponentLookup.Add ModuleName, True
        End If
    Next Definition
End Sub

Private Function ModuleGuardOwnsComponent(ByVal ModuleName As String) As Boolean
    'Returns True only for VBA components present in the fixed guarded-module roster supplied by x__UTIL_ADC_MODULE_GUARD_DEFINITIONS.
    Dim normalizedModuleName As String
    normalizedModuleName = Trim$(ModuleName)
    If Len(normalizedModuleName) = 0 Then Exit Function
    EnsureModuleGuardOwnedComponentLookup
    ModuleGuardOwnsComponent = mModuleGuardOwnedComponentLookup.Exists(normalizedModuleName)
End Function

Private Function ModuleGuardOwnedComponentDescription() As String
    'Returns the current owned roster from the same derived lookup, alphabetized only for stable diagnostic presentation.
    Dim keys As Variant
    Dim leftIndex As Long
    Dim rightIndex As Long
    Dim swapText As String
    Dim resultText As String
    EnsureModuleGuardOwnedComponentLookup
    If mModuleGuardOwnedComponentLookup.count = 0 Then Exit Function
    keys = mModuleGuardOwnedComponentLookup.keys
    For leftIndex = LBound(keys) To UBound(keys) - 1
        For rightIndex = leftIndex + 1 To UBound(keys)
            If StrComp(CStr(keys(leftIndex)), CStr(keys(rightIndex)), vbTextCompare) > 0 Then
                swapText = CStr(keys(leftIndex))
                keys(leftIndex) = keys(rightIndex)
                keys(rightIndex) = swapText
            End If
        Next rightIndex
    Next leftIndex
    For leftIndex = LBound(keys) To UBound(keys)
        If Len(resultText) > 0 Then resultText = resultText & ", "
        resultText = resultText & CStr(keys(leftIndex))
    Next leftIndex
    ModuleGuardOwnedComponentDescription = resultText
End Function

Private Function InitialModuleGuardBaselineDefinitions() As Collection
    'CANONICAL AUTHORITY: x__UTIL_ADC_MODULE_GUARD_DEFINITIONS owns the fixed guarded-module roster and approved-build batch identity.
    'Current/previous Version IDs and timestamps are read from installed managed headers.
    Dim rows As New Collection
    Dim rosterDefinitions As Collection
    Dim rosterDefinition As Variant
    Dim projectObject As Object
    Set projectObject = GuardVBProject()
    Set rosterDefinitions = ModuleGuardBaselineRosterDefinitions()
    If rosterDefinitions Is Nothing Then Err.Raise MODULE_GUARD_ERROR_BASE + 147, "InitialModuleGuardBaselineDefinitions", "The Module Guard definitions module did not return the guarded-module roster."
    For Each rosterDefinition In rosterDefinitions
        If UBound(rosterDefinition) < 1 Then Err.Raise MODULE_GUARD_ERROR_BASE + 148, "InitialModuleGuardBaselineDefinitions", "A guarded-module roster definition does not contain module name and component type."
        AddHeaderDerivedBaselineDefinition rows, projectObject, CStr(rosterDefinition(0)), CStr(rosterDefinition(1))
    Next rosterDefinition
    Set InitialModuleGuardBaselineDefinitions = rows
End Function

Private Sub AddHeaderDerivedBaselineDefinition(ByRef rows As Collection, ByVal projectObject As Object, ByVal ModuleName As String, ByVal expectedComponentType As String)
    'Reads the coordinated module identity only from the module's managed rewrite-header comments; the compatibility description field is intentionally blank for new registrations.
    On Error GoTo ErrHandler
    Dim component As Object
    Dim sourceText As String
    Dim actualComponentType As String
    Dim versionID As String
    Dim previousVersionID As String
    Dim timestampValue As Variant
    Dim timestampText As String
    Dim headerName As String
    Set component = projectObject.VBComponents(ModuleName)
    sourceText = ComponentHeaderSourceText(component)
    actualComponentType = GetComponentTypeText(CLng(component.Type))
    If StrComp(actualComponentType, expectedComponentType, vbTextCompare) <> 0 Then Err.Raise MODULE_GUARD_ERROR_BASE + 149, "AddHeaderDerivedBaselineDefinition", "The managed component type does not match the guarded roster." & vbCrLf & "Module: " & ModuleName & vbCrLf & "Expected: " & expectedComponentType & vbCrLf & "Actual: " & actualComponentType
    ReadNormalizedModuleGuardSourceMetadata sourceText, versionID, previousVersionID, timestampValue, headerName, timestampText
    If StrComp(headerName, ModuleName, vbBinaryCompare) <> 0 Then RaiseModuleHeaderIdentityFailure ModuleName, headerName
    If IsValidChatGPTVersionID(versionID) = False Then Err.Raise MODULE_GUARD_ERROR_BASE + 150, "AddHeaderDerivedBaselineDefinition", "The managed rewrite header contains an invalid current Version ID." & vbCrLf & "Module: " & ModuleName & vbCrLf & "Version ID: " & DisplayGuardValue(versionID)
    If Len(previousVersionID) > 0 And StrComp(previousVersionID, MODULE_GUARD_TRACKING_PREFIX & MODULE_GUARD_NO_VERSION_SUFFIX, vbBinaryCompare) <> 0 And IsValidChatGPTVersionID(previousVersionID) = False Then Err.Raise MODULE_GUARD_ERROR_BASE + 151, "AddHeaderDerivedBaselineDefinition", "The managed rewrite header contains an invalid Previous Version ID." & vbCrLf & "Module: " & ModuleName & vbCrLf & "Previous Version ID: " & DisplayGuardValue(previousVersionID)
    If IsNull(timestampValue) Then Err.Raise MODULE_GUARD_ERROR_BASE + 152, "AddHeaderDerivedBaselineDefinition", "The managed rewrite header contains a missing or invalid Version Timestamp." & vbCrLf & "Module: " & ModuleName & vbCrLf & "Timestamp: " & DisplayGuardValue(timestampText)
    AddBaselineDefinition rows, ModuleName, expectedComponentType, versionID, previousVersionID, timestampText
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If savedErrNumber = 0 Then savedErrNumber = MODULE_GUARD_ERROR_BASE + 153
    If Len(savedErrSource) = 0 Then savedErrSource = "AddHeaderDerivedBaselineDefinition"
    Err.Raise savedErrNumber, savedErrSource, savedErrDescription
End Sub

Private Sub EnsureModuleGuardQueries(ByVal db As DAO.Database, Optional ByVal managedDefinitions As Variant)
    'Query membership comes only from ModuleGuardManagedObjectDefinitions; this procedure supplies SQL only for canonical query members.
    Dim definitions As Collection
    Dim Definition As Variant
    Dim queryName As String
    If IsObject(managedDefinitions) Then Set definitions = managedDefinitions
    If definitions Is Nothing Then Set definitions = ModuleGuardManagedObjectDefinitions()
    For Each Definition In definitions
        If StrComp(CStr(Definition(2)), ADCAT_OBJECT_TYPE_QUERY, vbTextCompare) = 0 Then
            queryName = CStr(Definition(0))
            EnsureModuleGuardSavedQuery db, queryName, ModuleGuardQuerySql(queryName)
            Set db = Nothing
            Set db = currentDb
        End If
    Next Definition
End Sub

Private Function ModuleGuardQuerySql(ByVal queryName As String) As String
    'Maps a canonical Module Guard query identity to its SQL. This is build behavior, not query-membership authority.
    Select Case queryName
        Case ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_ModuleUpdates_Current")
            ModuleGuardQuerySql = "SELECT M." & SQL_FIELD_MODULE_NAME & ",M." & SQL_FIELD_COMPONENT_TYPE & ",M." & SQL_FIELD_EXPECTED_VERSION_ID & ",V." & SQL_FIELD_INSTALLED_VERSION_ID & ",M." & SQL_FIELD_EXPECTED_PREVIOUS_VERSION_ID & ",V." & SQL_FIELD_INSTALLED_PREVIOUS_VERSION_ID & ",M." & SQL_FIELD_EXPECTED_VERSION_TIMESTAMP & ",V." & SQL_FIELD_INSTALLED_VERSION_TIMESTAMP & ",M." & SQL_FIELD_UPDATE_BATCH_ID & ",M." & SQL_FIELD_EXPECTED_CHANGE_DESCRIPTION & ",M." & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_IS_REQUIRED & ",M.[" & FIELD_ACTIVE & "],M." & SQL_FIELD_BASELINE_ACCEPTED & ",M." & SQL_FIELD_MANIFEST_UPDATED_ON & ",M." & SQL_FIELD_LAST_VERIFIED_ON & ",M." & SQL_FIELD_LAST_VERIFICATION_RESULT & ",M." & SQL_FIELD_LAST_VERIFICATION_NOTES & ",V." & SQL_FIELD_VERIFICATION_RUN_ID & ",V." & SQL_FIELD_VERIFIED_ON & ",V." & SQL_FIELD_VERIFICATION_RESULT & ",V." & _
                        SQL_FIELD_VERIFICATION_NOTES & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionManifest") & "] AS M LEFT JOIN [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionVerification") & "] AS V ON (M." & SQL_FIELD_LAST_VERIFIED_ON & "=V." & SQL_FIELD_VERIFIED_ON & ") AND (M." & SQL_FIELD_EXPECTED_VERSION_ID & "=V." & SQL_FIELD_EXPECTED_VERSION_ID & ") AND (M." & SQL_FIELD_MODULE_NAME & "=V." & SQL_FIELD_MODULE_NAME & ") WHERE M.[" & FIELD_ACTIVE & "]=True ORDER BY " & OrderByPolicyText("MODULE_GUARD_CURRENT") & ";"
        Case ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_ModuleUpdates_History")
            ModuleGuardQuerySql = "SELECT H." & SQL_FIELD_ID & ", H." & SQL_FIELD_RECORDED_ON & ", H." & SQL_FIELD_UPDATE_BATCH_ID & ", H." & SQL_FIELD_MODULE_NAME & ", H." & SQL_FIELD_COMPONENT_TYPE & ", H." & SQL_FIELD_CHAT_GPT_VERSION_ID & ", H." & SQL_FIELD_PREVIOUS_CHAT_GPT_VERSION_ID & ", H." & SQL_FIELD_VERSION_TIMESTAMP & ", H." & SQL_FIELD_CHANGE_DESCRIPTION & ", H." & SQL_FIELD_RECORDED_BY & ", H." & SQL_FIELD_IS_BASELINE & ", H." & SQL_FIELD_SUPERSEDED_ON & ", H." & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_NOTES & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleUpdateHistory") & "] AS H ORDER BY " & OrderByPolicyText("MODULE_GUARD_HISTORY") & ";"
        Case ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_ModuleUpdates_ByBatch")
            ModuleGuardQuerySql = "SELECT H." & SQL_FIELD_UPDATE_BATCH_ID & ", H." & SQL_FIELD_RECORDED_ON & ", H." & SQL_FIELD_MODULE_NAME & ", H." & SQL_FIELD_COMPONENT_TYPE & ", H." & SQL_FIELD_CHAT_GPT_VERSION_ID & ", H." & SQL_FIELD_PREVIOUS_CHAT_GPT_VERSION_ID & ", H." & SQL_FIELD_VERSION_TIMESTAMP & ", H." & SQL_FIELD_CHANGE_DESCRIPTION & ", H." & SQL_FIELD_RECORDED_BY & ", H." & SQL_FIELD_IS_BASELINE & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleUpdateHistory") & "] AS H ORDER BY " & OrderByPolicyText("MODULE_GUARD_BY_BATCH") & ";"
        Case ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_ModuleUpdates_Mismatches")
            ModuleGuardQuerySql = "SELECT V." & SQL_FIELD_VERIFICATION_RUN_ID & ", V." & SQL_FIELD_VERIFIED_ON & ", V." & SQL_FIELD_MODULE_NAME & ", V." & SQL_FIELD_COMPONENT_TYPE & ", V." & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_IS_REQUIRED & ", V." & SQL_FIELD_IS_MANIFESTED & ", V." & SQL_FIELD_EXPECTED_VERSION_ID & ", V." & SQL_FIELD_INSTALLED_VERSION_ID & ", V." & SQL_FIELD_EXPECTED_PREVIOUS_VERSION_ID & ", V." & SQL_FIELD_INSTALLED_PREVIOUS_VERSION_ID & ", V." & SQL_FIELD_EXPECTED_VERSION_TIMESTAMP & ", V." & SQL_FIELD_INSTALLED_VERSION_TIMESTAMP & ", V." & SQL_FIELD_VERIFICATION_RESULT & ", V." & SQL_FIELD_VERIFICATION_NOTES & ", V." & SQL_FIELD_UPDATE_BATCH_ID & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionVerification") & "] AS V WHERE V." & SQL_FIELD_VERIFICATION_RUN_ID & "=(SELECT Max(V2." & SQL_FIELD_VERIFICATION_RUN_ID & ") FROM [" & _
                        ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionVerification") & "] AS V2) AND V." & SQL_FIELD_VERIFICATION_RESULT & "<>""" & MODULE_GUARD_RESULT_MATCH & """ ORDER BY " & OrderByPolicyText("MODULE_GUARD_MISMATCHES") & ";"
        Case ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_ModuleUpdates_Missing")
            ModuleGuardQuerySql = "SELECT V." & SQL_FIELD_VERIFICATION_RUN_ID & ", V." & SQL_FIELD_VERIFIED_ON & ", V." & SQL_FIELD_MODULE_NAME & ", V." & SQL_FIELD_COMPONENT_TYPE & ", V." & SQL_FIELD_EXPECTED_VERSION_ID & ", V." & SQL_FIELD_INSTALLED_VERSION_ID & ", V." & SQL_FIELD_VERIFICATION_RESULT & ", V." & SQL_FIELD_VERIFICATION_NOTES & ", V." & SQL_FIELD_UPDATE_BATCH_ID & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionVerification") & "] AS V WHERE V." & SQL_FIELD_VERIFICATION_RUN_ID & "=(SELECT Max(V2." & SQL_FIELD_VERIFICATION_RUN_ID & ") FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionVerification") & "] AS V2) AND V." & SQL_FIELD_VERIFICATION_RESULT & "=""" & MODULE_GUARD_RESULT_MISSING & """ ORDER BY " & OrderByPolicyText("MODULE_GUARD_MISSING") & ";"
        Case ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_ModuleUpdates_Unregistered")
            ModuleGuardQuerySql = "SELECT V." & SQL_FIELD_VERIFICATION_RUN_ID & ", V." & SQL_FIELD_VERIFIED_ON & ", V." & SQL_FIELD_MODULE_NAME & ", V." & SQL_FIELD_COMPONENT_TYPE & ", V." & SQL_FIELD_INSTALLED_VERSION_ID & ", V." & SQL_FIELD_INSTALLED_PREVIOUS_VERSION_ID & ", V." & SQL_FIELD_INSTALLED_VERSION_TIMESTAMP & ", V." & SQL_FIELD_VERIFICATION_RESULT & ", V." & SQL_FIELD_VERIFICATION_NOTES & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionVerification") & "] AS V WHERE V." & SQL_FIELD_VERIFICATION_RUN_ID & "=(SELECT Max(V2." & SQL_FIELD_VERIFICATION_RUN_ID & ") FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionVerification") & "] AS V2) AND V." & SQL_FIELD_VERIFICATION_RESULT & "=""" & MODULE_GUARD_RESULT_UNREGISTERED & """ ORDER BY " & OrderByPolicyText("MODULE_GUARD_UNREGISTERED") & ";"
        Case ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_ModuleUpdates_BrokenChains")
            ModuleGuardQuerySql = "SELECT V." & SQL_FIELD_VERIFICATION_RUN_ID & ", V." & SQL_FIELD_VERIFIED_ON & ", V." & SQL_FIELD_MODULE_NAME & ", V." & SQL_FIELD_EXPECTED_VERSION_ID & ", V." & SQL_FIELD_INSTALLED_VERSION_ID & ", V." & SQL_FIELD_EXPECTED_PREVIOUS_VERSION_ID & ", V." & SQL_FIELD_INSTALLED_PREVIOUS_VERSION_ID & ", V." & SQL_FIELD_VERIFICATION_RESULT & ", V." & SQL_FIELD_VERIFICATION_NOTES & ", V." & SQL_FIELD_UPDATE_BATCH_ID & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionVerification") & "] AS V WHERE V." & SQL_FIELD_VERIFICATION_RUN_ID & "=(SELECT Max(V2." & SQL_FIELD_VERIFICATION_RUN_ID & ") FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionVerification") & "] AS V2) AND V." & SQL_FIELD_VERIFICATION_RESULT & " In (""" & MODULE_GUARD_RESULT_BROKEN_CHAIN & """,""" & MODULE_GUARD_RESULT_DUPLICATE_ID & """,""" & _
                        MODULE_GUARD_RESULT_HEADER_MISMATCH & """,""" & MODULE_GUARD_RESULT_INVALID_ID & """) ORDER BY " & OrderByPolicyText("MODULE_GUARD_BROKEN_CHAINS") & ";"
        Case Else
            Err.Raise MODULE_GUARD_ERROR_BASE + 55, "ModuleGuardQuerySql", "No SQL builder mapping exists for the canonical Module Guard query." & vbCrLf & "Query: " & queryName
    End Select
End Function

Private Sub CreateModuleGuardTableFromDefinitions(ByVal db As DAO.Database, ByVal tableName As String, ByVal definitions As Variant)
    'Creates a new module-guard table from the same field definitions used by repair and logical-order enforcement.
    Dim Definition As Variant
    Dim parts As Variant
    Dim sqlText As String
    Dim fieldSql As String
    Dim expectedType As Long
    Dim expectedSize As Long
    Dim expectedAttributes As Long
    Dim requiredField As Boolean
    sqlText = "CREATE TABLE [" & tableName & "] ("
    For Each Definition In definitions
        parts = Split(CStr(Definition), "|")
        expectedType = CLng(parts(1))
        expectedSize = CLng(parts(2))
        expectedAttributes = CLng(parts(3))
        requiredField = (UBound(parts) >= 4 And CLng(parts(4)) <> 0)
        If (expectedAttributes And dbAutoIncrField) <> 0 Then
            fieldSql = "[" & CStr(parts(0)) & "] COUNTER"
        Else
            fieldSql = "[" & CStr(parts(0)) & "] " & DdlTypeText(expectedType, expectedSize)
        End If
        If requiredField Then fieldSql = fieldSql & " NOT NULL"
        If Right$(sqlText, 1) <> "(" Then sqlText = sqlText & ", "
        sqlText = sqlText & fieldSql
    Next Definition
    sqlText = sqlText & ")"
    db.Execute sqlText, dbFailOnError
End Sub

Private Function ModuleGuardFieldOrderText(ByVal definitions As Variant) As String
    Dim Definition As Variant
    Dim parts As Variant
    Dim orderText As String
    For Each Definition In definitions
        parts = Split(CStr(Definition), "|")
        If Len(orderText) > 0 Then orderText = orderText & "|"
        orderText = orderText & CStr(parts(0))
    Next Definition
    ModuleGuardFieldOrderText = orderText
End Function

Private Function FreshModuleGuardTableDef(ByVal db As DAO.Database, ByVal tableName As String) As DAO.tableDef
    Dim tdf As DAO.tableDef
    db.TableDefs.Refresh
    For Each tdf In db.TableDefs
        PumpModuleGuardSetupResponsiveness
        If StrComp(tdf.Name, tableName, vbTextCompare) = 0 Then
            Set FreshModuleGuardTableDef = tdf
            Exit Function
        End If
    Next tdf
    Err.Raise MODULE_GUARD_ERROR_BASE + 50, "FreshModuleGuardTableDef", "Required module-guard table was not found after refreshing DAO.TableDefs." & vbCrLf & "Table: " & tableName
End Function

Private Sub RequireModuleGuardTable(ByVal db As DAO.Database, ByVal tableName As String)
    Call FreshModuleGuardTableDef(db, tableName)
End Sub

Private Function FindModuleGuardField(ByVal tdf As DAO.tableDef, ByVal fieldName As String) As DAO.Field
    Dim fld As DAO.Field
    For Each fld In tdf.fields
        PumpModuleGuardSetupResponsiveness
        If StrComp(fld.Name, fieldName, vbTextCompare) = 0 Then
            Set FindModuleGuardField = fld
            Exit Function
        End If
    Next fld
End Function

Private Function ModuleGuardExpectedSourceTablesForSql(ByVal sqlText As String) As String
    Dim resultText As String
    Dim sqlKey As String
    sqlKey = FormattingInsensitiveQuerySqlKey(sqlText)
    If InStr(1, sqlKey, LCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionManifest")), vbBinaryCompare) > 0 Then AppendGuardDetail resultText, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionManifest")
    If InStr(1, sqlKey, LCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleUpdateHistory")), vbBinaryCompare) > 0 Then AppendGuardDetail resultText, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleUpdateHistory")
    If InStr(1, sqlKey, LCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionVerification")), vbBinaryCompare) > 0 Then AppendGuardDetail resultText, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionVerification")
    ModuleGuardExpectedSourceTablesForSql = Replace(resultText, vbCrLf, ", ")
End Function

Private Function ModuleGuardDeclaredVBName(ByVal sourceText As String) As String
    Dim lineValue As Variant
    Dim trimmedLine As String
    If Len(sourceText) = 0 Then Exit Function
    For Each lineValue In Split(sourceText, vbCrLf)
        PumpModuleGuardSetupResponsiveness
        trimmedLine = Trim$(CStr(lineValue))
        If Left$(trimmedLine, 1) = "'" Then trimmedLine = Trim$(Mid$(trimmedLine, 2))
        If Left$(trimmedLine, Len("VB_Name:")) = "VB_Name:" Then
            ModuleGuardDeclaredVBName = Trim$(Mid$(trimmedLine, Len("VB_Name:") + 1))
            Exit Function
        End If
    Next lineValue
End Function

Private Sub ReadNormalizedModuleGuardSourceMetadata(ByVal sourceText As String, ByRef versionID As String, ByRef previousVersionID As String, ByRef timestampValue As Variant, ByRef headerName As String, ByRef timestampText As String)
    'ComponentHeaderSourceText supplies only the bounded managed-header block on the normal path. This parser reads identity and version metadata only.
    Dim sourceLines As Variant
    Dim lineValue As Variant
    Dim trimmedLine As String
    Dim parsedTimestamp As Date
    versionID = vbNullString
    previousVersionID = vbNullString
    timestampText = vbNullString
    headerName = vbNullString
    timestampValue = Null
    If Len(sourceText) = 0 Then Exit Sub
    sourceLines = Split(sourceText, vbCrLf)
    For Each lineValue In sourceLines
        PumpModuleGuardSetupResponsiveness
        trimmedLine = Trim$(CStr(lineValue))
        If Left$(trimmedLine, 1) = "'" Then trimmedLine = Trim$(Mid$(trimmedLine, 2))
        If Len(versionID) = 0 And Left$(trimmedLine, Len("ChatGPT Version ID:")) = "ChatGPT Version ID:" Then versionID = Trim$(Mid$(trimmedLine, Len("ChatGPT Version ID:") + 1))
        If Len(previousVersionID) = 0 And Left$(trimmedLine, Len("ChatGPT Previous Version ID:")) = "ChatGPT Previous Version ID:" Then previousVersionID = Trim$(Mid$(trimmedLine, Len("ChatGPT Previous Version ID:") + 1))
        If Len(timestampText) = 0 And Left$(trimmedLine, Len("ChatGPT Version Timestamp:")) = "ChatGPT Version Timestamp:" Then timestampText = Trim$(Mid$(trimmedLine, Len("ChatGPT Version Timestamp:") + 1))
        If Len(headerName) = 0 And Left$(trimmedLine, Len("VB_Name:")) = "VB_Name:" Then headerName = Trim$(Mid$(trimmedLine, Len("VB_Name:") + 1))
        If Len(versionID) > 0 And Len(previousVersionID) > 0 And Len(timestampText) > 0 And Len(headerName) > 0 Then Exit For
    Next lineValue
    If TryParseChatGPTVersionTimestamp(timestampText, parsedTimestamp) Then timestampValue = parsedTimestamp
End Sub

Private Function ModuleGuardSourceFingerprint(ByVal sourceText As String) As String
    sourceText = NormalizeSourceText(sourceText)
    ModuleGuardSourceFingerprint = ModuleGuardFingerprintFromNormalizedText(sourceText)
End Function

Private Function ModuleGuardFingerprintFromNormalizedText(ByVal sourceText As String) As String
    Dim indexValue As Long
    Dim sourceLength As Long
    Dim byteOffset As Long
    Dim reverseByteOffset As Long
    Dim sourceBytes() As Byte
    Dim forwardHash As Double
    Dim reverseHash As Double
    Dim forwardCharacter As Long
    Dim reverseCharacter As Long
    sourceLength = Len(sourceText)
    forwardHash = 5381#
    reverseHash = 5381#
    If sourceLength = 0 Then ModuleGuardFingerprintFromNormalizedText = "0000150500001505": Exit Function
    sourceBytes = sourceText
    For indexValue = 0 To sourceLength - 1
        If (indexValue And 255) = 0 Then PumpModuleGuardSetupResponsiveness
        byteOffset = indexValue * 2
        forwardCharacter = CLng(sourceBytes(byteOffset)) + CLng(sourceBytes(byteOffset + 1)) * 256&
        forwardHash = forwardHash * 33# + forwardCharacter
        forwardHash = forwardHash - Int(forwardHash / 2147483647#) * 2147483647#
        reverseByteOffset = (sourceLength - indexValue - 1) * 2
        reverseCharacter = CLng(sourceBytes(reverseByteOffset)) + CLng(sourceBytes(reverseByteOffset + 1)) * 256&
        reverseHash = reverseHash * 33# + reverseCharacter
        reverseHash = reverseHash - Int(reverseHash / 2147483647#) * 2147483647#
    Next indexValue
    ModuleGuardFingerprintFromNormalizedText = Right$("00000000" & Hex$(CLng(forwardHash)), 8) & Right$("00000000" & Hex$(CLng(reverseHash)), 8)
End Function

Private Function RollbackModuleGuardTransaction(ByVal workspaceObject As DAO.workspace, ByRef cleanupDetails As String) As Boolean
    'Rolls back the transaction started and owned by TryRegisterApprovedBuild.
    On Error GoTo ErrHandler
    If workspaceObject Is Nothing Then
        AppendGuardDetail cleanupDetails, "Rollback failed because the DAO Workspace object was not available."
        Exit Function
    End If
    If mModuleGuardRegistrationTransactionOwned = False Then
        AppendGuardDetail cleanupDetails, "Rollback was not attempted because this module did not own an active registration transaction."
        Exit Function
    End If
    If mModuleGuardNestedVerificationTransactionActive Then
        workspaceObject.Rollback
        mModuleGuardNestedVerificationTransactionActive = False
        AppendGuardDetail cleanupDetails, "A still-active nested verification transaction was rolled back before the owning registration transaction."
    End If
    workspaceObject.Rollback
    mModuleGuardRegistrationTransactionOwned = False
    AppendGuardDetail cleanupDetails, "Rollback completed for the registration-owned DAO Workspace transaction."
    RollbackModuleGuardTransaction = True
    Exit Function
ErrHandler:
    AppendGuardDetail cleanupDetails, "Cleanup failure while rolling back the module-guard transaction. Err.Number=" & CStr(Err.Number) & "; Err.Source=" & Err.Source & "; Err.Description=" & Err.Description
End Function

Private Sub CloseModuleGuardRecordset(ByVal recordsetObject As DAO.Recordset, ByVal recordsetDescription As String, ByRef cleanupDetails As String)
    On Error GoTo ErrHandler
    If recordsetObject Is Nothing Then Exit Sub
    recordsetObject.Close
    Exit Sub
ErrHandler:
    AppendGuardDetail cleanupDetails, "Cleanup failure while closing " & recordsetDescription & ". Err.Number=" & CStr(Err.Number) & "; Err.Source=" & Err.Source & "; Err.Description=" & Err.Description
End Sub

Private Sub CaptureModuleGuardHistoryIndex(ByVal db As DAO.Database, ByVal historyByKey As Object)
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim ModuleName As String
    Dim versionID As String
    Dim historyKey As String
    Dim versionTimestampValue As Variant
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Set rs = db.OpenRecordset("SELECT " & SQL_FIELD_MODULE_NAME & ", " & SQL_FIELD_COMPONENT_TYPE & ", " & SQL_FIELD_CHAT_GPT_VERSION_ID & ", " & SQL_FIELD_PREVIOUS_CHAT_GPT_VERSION_ID & ", " & SQL_FIELD_VERSION_TIMESTAMP & ", " & SQL_FIELD_IS_BASELINE & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleUpdateHistory") & "]", dbOpenSnapshot)
    Do While rs.EOF = False
        PumpModuleGuardSetupResponsiveness
        ModuleName = Trim$(CStr(Nz(rs.fields(FIELD_MODULE_NAME), vbNullString)))
        versionID = Trim$(CStr(Nz(rs.fields(FIELD_CHAT_GPT_VERSION_ID), vbNullString)))
        If ModuleGuardOwnsComponent(ModuleName) And Len(versionID) > 0 Then
            If IsNull(rs.fields(FIELD_VERSION_TIMESTAMP).value) Then
                versionTimestampValue = Null
            Else
                versionTimestampValue = CDate(rs.fields(FIELD_VERSION_TIMESTAMP).value)
            End If
            historyKey = ModuleGuardHistoryKey(ModuleName, versionID)
            If historyByKey.Exists(historyKey) Then Err.Raise MODULE_GUARD_ERROR_BASE + 116, "CaptureModuleGuardHistoryIndex", "Append-only Module Guard history contains duplicate module/version rows." & vbCrLf & "Module: " & ModuleName & vbCrLf & "Version: " & versionID
            historyByKey.Add historyKey, Array( _
                Trim$(CStr(Nz(rs.fields(FIELD_PREVIOUS_CHAT_GPT_VERSION_ID), vbNullString))), _
                CBool(Nz(rs!isBaseline, False)), _
                Trim$(CStr(Nz(rs.fields(FIELD_COMPONENT_TYPE), vbNullString))), _
                versionTimestampValue)
        End If
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
    Err.Raise savedErrNumber, savedErrSource, savedErrDescription
End Sub

Private Function ModuleGuardHistoryKey(ByVal ModuleName As String, ByVal versionID As String) As String
    ModuleGuardHistoryKey = LCase$(Trim$(ModuleName)) & Chr$(30) & Trim$(versionID)
End Function

Private Function ModuleGuardHistoryChainValid(ByVal historyByKey As Object, ByVal ModuleName As String, ByVal expectedVersion As String, ByVal expectedPrevious As String, ByVal baselineAccepted As Boolean, ByRef resultNotes As String) As Boolean
    Dim historyData As Variant
    Dim historyPrevious As String
    Dim historyBaseline As Boolean
    Dim historyKey As String
    historyKey = ModuleGuardHistoryKey(ModuleName, expectedVersion)
    If historyByKey.Exists(historyKey) = False Then AppendGuardDetail resultNotes, "No update-history row exists for the expected current version.": Exit Function
    historyData = historyByKey(historyKey)
    historyPrevious = CStr(historyData(0))
    historyBaseline = CBool(historyData(1))
    If Len(expectedPrevious) > 0 And StrComp(historyPrevious, expectedPrevious, vbBinaryCompare) <> 0 Then AppendGuardDetail resultNotes, "The update-history predecessor does not match the manifest predecessor.": Exit Function
    If baselineAccepted Or historyBaseline Or Len(expectedPrevious) = 0 Or StrComp(expectedPrevious, MODULE_GUARD_TRACKING_PREFIX & MODULE_GUARD_NO_VERSION_SUFFIX, vbBinaryCompare) = 0 Then ModuleGuardHistoryChainValid = True: Exit Function
    If historyByKey.Exists(ModuleGuardHistoryKey(ModuleName, expectedPrevious)) = False Then AppendGuardDetail resultNotes, "The predecessor version is not present in append-only update history.": Exit Function
    ModuleGuardHistoryChainValid = True
End Function

Private Function FirstModuleGuardMismatchSummary(ByVal db As DAO.Database, ByVal runID As String) As String
    'Returns the exact first failing module comparison for the popup and Immediate Window.
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim resultText As String
    Dim notesText As String
    Set rs = db.OpenRecordset("SELECT TOP 1 * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionVerification") & "] WHERE " & SQL_FIELD_VERIFICATION_RUN_ID & "=" & SharedDao_SqlQuote(runID) & " AND " & SQL_FIELD_VERIFICATION_RESULT & "<>""" & MODULE_GUARD_RESULT_MATCH & """ ORDER BY " & OrderByPolicyText("MODULE_GUARD_FIRST_MISMATCH"), dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        FirstModuleGuardMismatchSummary = "Module Guard found a mismatch, but no failing module row was available." & vbCrLf & "Verification run: " & runID
        Exit Function
    End If
    notesText = Trim$(CStr(Nz(rs.fields(FIELD_VERIFICATION_NOTES), vbNullString)))
    resultText = "First failing module:" & vbCrLf & _
                 "Module: " & DisplayGuardValue(CStr(Nz(rs.fields(FIELD_MODULE_NAME), vbNullString))) & vbCrLf & _
                 "Result: " & DisplayGuardValue(CStr(rs.fields(FIELD_VERIFICATION_RESULT))) & vbCrLf & _
                 "Component type: " & DisplayGuardValue(CStr(Nz(rs.fields(FIELD_COMPONENT_TYPE), vbNullString))) & vbCrLf & _
                 "Expected Version ID: " & DisplayGuardValue(CStr(Nz(rs!expectedVersionID, vbNullString))) & vbCrLf & _
                 "Installed Version ID: " & DisplayGuardValue(CStr(Nz(rs!installedVersionID, vbNullString))) & vbCrLf & _
                 "Expected previous Version ID: " & DisplayGuardValue(CStr(Nz(rs!expectedPreviousVersionID, vbNullString))) & vbCrLf & _
                 "Installed previous Version ID: " & DisplayGuardValue(CStr(Nz(rs!installedPreviousVersionID, vbNullString))) & vbCrLf & _
                 "Expected timestamp: " & GuardDisplayTimestamp(rs.fields(FIELD_EXPECTED_VERSION_TIMESTAMP)) & vbCrLf & _
                 "Installed timestamp: " & GuardDisplayTimestamp(rs.fields(FIELD_INSTALLED_VERSION_TIMESTAMP))
    If Len(notesText) > 0 Then resultText = resultText & vbCrLf & "Verification notes:" & vbCrLf & notesText
    resultText = resultText & vbCrLf & "Verification run: " & runID
    rs.Close
    FirstModuleGuardMismatchSummary = resultText
    Exit Function
ErrHandler:
    Dim formattingErrNumber As Long
    Dim formattingErrSource As String
    Dim formattingErrDescription As String
    formattingErrNumber = Err.Number
    formattingErrSource = Err.Source
    formattingErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    On Error GoTo 0
    FirstModuleGuardMismatchSummary = "Module Guard found a module mismatch, but could not format the detailed comparison." & vbCrLf & _
        "Verification run: " & runID & vbCrLf & _
        "Formatting Err.Number: " & CStr(formattingErrNumber) & vbCrLf & _
        "Formatting Err.Source: " & formattingErrSource & vbCrLf & _
        "Formatting Err.Description: " & formattingErrDescription
End Function

Private Function LatestModuleGuardMismatchSummary(ByVal db As DAO.Database, ByVal runID As String) As String
    Dim rs As DAO.Recordset
    Dim resultText As String
    Dim rowCount As Long
    Set rs = db.OpenRecordset("SELECT " & SQL_FIELD_MODULE_NAME & ", " & SQL_FIELD_EXPECTED_VERSION_ID & ", " & SQL_FIELD_INSTALLED_VERSION_ID & ", " & SQL_FIELD_VERIFICATION_RESULT & ", " & SQL_FIELD_VERIFICATION_NOTES & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionVerification") & "] WHERE " & SQL_FIELD_VERIFICATION_RUN_ID & "=" & SharedDao_SqlQuote(runID) & " AND " & SQL_FIELD_VERIFICATION_RESULT & "<>""" & MODULE_GUARD_RESULT_MATCH & """ ORDER BY " & OrderByPolicyText("MODULE_GUARD_MISMATCH_SUMMARY"), dbOpenSnapshot)
    Do While rs.EOF = False
        PumpModuleGuardSetupResponsiveness
        rowCount = rowCount + 1
        If rowCount <= MODULE_GUARD_MAX_DETAIL_ROWS Then
            If Len(resultText) > 0 Then resultText = resultText & vbCrLf & vbCrLf
            resultText = resultText & "Module: " & Nz(rs.fields(FIELD_MODULE_NAME), vbNullString) & vbCrLf & "Result: " & rs.fields(FIELD_VERIFICATION_RESULT) & vbCrLf & "Expected: " & Nz(rs!expectedVersionID, "(not manifested)") & vbCrLf & "Installed: " & Nz(rs!installedVersionID, "(missing)") & vbCrLf & "Details: " & Nz(rs.fields(FIELD_VERIFICATION_NOTES), vbNullString)
        End If
        rs.MoveNext
    Loop
    rs.Close
    If rowCount > MODULE_GUARD_MAX_DETAIL_ROWS Then resultText = resultText & vbCrLf & vbCrLf & "Additional mismatch rows omitted from this message: " & CStr(rowCount - MODULE_GUARD_MAX_DETAIL_ROWS) & ". Open " & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_ModuleUpdates_Mismatches") & " for the complete list."
    LatestModuleGuardMismatchSummary = resultText
End Function

Private Sub InsertOrVerifyApprovedHistoryRow(ByVal db As DAO.Database, ByVal updateBatchID As String, ByVal rowData As Variant, ByVal recordedBy As String)
    'Preserves append-only evidence when an approved current version was already recorded by an earlier batch; otherwise inserts the reviewed reactivation-baseline row.
    Dim rs As DAO.Recordset
    Dim ModuleName As String
    Dim versionID As String
    ModuleName = CStr(rowData(0))
    versionID = CStr(rowData(2))
    Set rs = db.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleUpdateHistory") & "] WHERE " & SQL_FIELD_MODULE_NAME & "=" & SharedDao_SqlQuote(ModuleName) & " AND " & SQL_FIELD_CHAT_GPT_VERSION_ID & "=" & SharedDao_SqlQuote(versionID), dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        InsertOrVerifyHistoryRow db, updateBatchID, rowData, True, recordedBy
        Exit Sub
    End If
    If StrComp(Nz(rs.fields(FIELD_COMPONENT_TYPE), vbNullString), CStr(rowData(1)), vbTextCompare) <> 0 Or StrComp(Nz(rs.fields(FIELD_PREVIOUS_CHAT_GPT_VERSION_ID), vbNullString), CStr(rowData(3)), vbBinaryCompare) <> 0 Then rs.Close: Err.Raise MODULE_GUARD_ERROR_BASE + 68, "InsertOrVerifyApprovedHistoryRow", "An existing history row conflicts with the approved current build." & vbCrLf & "Module: " & ModuleName & vbCrLf & "Version: " & versionID
    If IsNull(rs.fields(FIELD_VERSION_TIMESTAMP)) Xor IsNull(rowData(4)) Then rs.Close: Err.Raise MODULE_GUARD_ERROR_BASE + 69, "InsertOrVerifyApprovedHistoryRow", "An existing history timestamp conflicts with the approved current build." & vbCrLf & "Module: " & ModuleName & vbCrLf & "Version: " & versionID
    If IsNull(rs.fields(FIELD_VERSION_TIMESTAMP)) = False Then
        If FormatChatGPTVersionTimestamp(CDate(rs.fields(FIELD_VERSION_TIMESTAMP))) <> FormatChatGPTVersionTimestamp(CDate(rowData(4))) Then rs.Close: Err.Raise MODULE_GUARD_ERROR_BASE + 70, "InsertOrVerifyApprovedHistoryRow", "An existing history timestamp conflicts with the approved current build." & vbCrLf & "Module: " & ModuleName & vbCrLf & "Version: " & versionID
    End If
    rs.Close
End Sub

Private Function ApprovedBuildComparisonDiagnostic(ByVal ModuleName As String, ByVal Definition As Variant, ByVal installedType As String, ByVal installedVersion As String, ByVal installedPrevious As String, ByVal installedTimestamp As Variant, ByVal installedHeaderName As String, ByVal manifestWriteStarted As Boolean) As String
    'Formats one explicit expected-versus-installed comparison for diagnostics and the blocking popup.
    Dim installedTimestampText As String
    Dim resultText As String
    If IsNull(installedTimestamp) Then installedTimestampText = "(blank)" Else installedTimestampText = FormatChatGPTVersionTimestamp(CDate(installedTimestamp))
    resultText = "Actual VBA component name: " & DisplayGuardValue(ModuleName)
    resultText = resultText & vbCrLf & "Declared header VB_Name: " & DisplayGuardValue(installedHeaderName)
    resultText = resultText & vbCrLf & "Expected component type: " & DisplayGuardValue(CStr(Definition(1)))
    resultText = resultText & vbCrLf & "Installed component type: " & DisplayGuardValue(installedType)
    resultText = resultText & vbCrLf & "Expected Version ID: " & DisplayGuardValue(CStr(Definition(2)))
    resultText = resultText & vbCrLf & "Installed Version ID: " & DisplayGuardValue(installedVersion)
    resultText = resultText & vbCrLf & "Expected previous Version ID: " & DisplayGuardValue(CStr(Definition(3)))
    resultText = resultText & vbCrLf & "Installed previous Version ID: " & DisplayGuardValue(installedPrevious)
    resultText = resultText & vbCrLf & "Expected timestamp: " & DisplayGuardValue(CStr(Definition(4)))
    resultText = resultText & vbCrLf & "Installed timestamp: " & installedTimestampText
    resultText = resultText & vbCrLf & "Approval basis: identity, type, Version-ID lineage, timestamp, required presence, and Version-ID uniqueness"
    resultText = resultText & vbCrLf & "Manifest write started: " & CStr(manifestWriteStarted)
    ApprovedBuildComparisonDiagnostic = resultText
End Function

Private Sub RequireCompleteModuleGuardStructure(ByVal db As DAO.Database, Optional ByVal managedDefinitions As Variant)
    Dim definitions As Collection
    Dim Definition As Variant
    Dim objectName As String
    Dim objectType As String
    db.QueryDefs.Refresh
    If IsObject(managedDefinitions) Then Set definitions = managedDefinitions
    If definitions Is Nothing Then Set definitions = ModuleGuardManagedObjectDefinitions()
    For Each Definition In definitions
        objectName = CStr(Definition(0))
        objectType = CStr(Definition(2))
        Select Case objectType
            Case ADCAT_OBJECT_TYPE_TABLE
                RequireModuleGuardTable db, objectName
            Case ADCAT_OBJECT_TYPE_QUERY
                RequireModuleGuardQuery db, objectName
            Case Else
                Err.Raise MODULE_GUARD_ERROR_BASE + 56, "RequireCompleteModuleGuardStructure", "ModuleGuardManagedObjectDefinitions contains an unsupported object type." & vbCrLf & "Object: " & objectName & vbCrLf & "ObjectType: " & objectType
        End Select
    Next Definition
End Sub

Private Sub RequireModuleGuardQuery(ByVal db As DAO.Database, ByVal queryName As String)
    Dim qdf As DAO.queryDef
    Set qdf = FindQueryDef(db, queryName)
    If qdf Is Nothing Then Err.Raise MODULE_GUARD_ERROR_BASE + 53, "RequireModuleGuardQuery", "Required module update guard query is missing after repair." & vbCrLf & "Query: " & queryName
End Sub

Private Function ModuleGuardRepairFailureText(ByVal stageText As String, ByVal targetObject As String, ByVal savedErrNumber As Long, ByVal savedErrSource As String, ByVal savedErrDescription As String, Optional ByVal managedDefinitions As Variant) As String
    Dim db As DAO.Database
    Dim resultText As String
    resultText = "Could not repair the " & DEFAULT_UTILITY_LONG_NAME & " module update guard." & vbCrLf & vbCrLf
    resultText = resultText & "Stage: " & stageText & vbCrLf
    resultText = resultText & "Target: " & targetObject & vbCrLf
    resultText = resultText & "Err.Number: " & CStr(savedErrNumber) & vbCrLf
    resultText = resultText & "Err.Source: " & savedErrSource & vbCrLf
    resultText = resultText & "Err.Description: " & savedErrDescription
    On Error Resume Next
    Set db = currentDb
    If Not db Is Nothing Then resultText = resultText & vbCrLf & vbCrLf & ModuleGuardRepairInventoryText(db, managedDefinitions)
    If Err.Number <> 0 Then
        resultText = resultText & vbCrLf & "Diagnostic inventory error: " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    ModuleGuardRepairFailureText = resultText
End Function

Private Function ModuleGuardRepairInventoryText(ByVal db As DAO.Database, Optional ByVal managedDefinitions As Variant) As String
    Dim definitions As Collection
    Dim Definition As Variant
    Dim objectName As String
    Dim objectType As String
    Dim resultText As String
    resultText = "Database: " & ModuleGuardSafeDatabaseName(db)
    If IsObject(managedDefinitions) Then Set definitions = managedDefinitions
    If definitions Is Nothing Then Set definitions = ModuleGuardManagedObjectDefinitions()
    For Each Definition In definitions
        objectName = CStr(Definition(0))
        objectType = CStr(Definition(2))
        Select Case objectType
            Case ADCAT_OBJECT_TYPE_TABLE
                resultText = resultText & vbCrLf & ModuleGuardTableInventoryLine(db, objectName)
            Case ADCAT_OBJECT_TYPE_QUERY
                resultText = resultText & vbCrLf & ModuleGuardQueryInventoryLine(db, objectName)
            Case Else
                resultText = resultText & vbCrLf & "Unsupported canonical Module Guard object [" & objectName & "] type [" & objectType & "]."
        End Select
    Next Definition
    ModuleGuardRepairInventoryText = resultText
End Function

Private Function ModuleGuardTableInventoryLine(ByVal db As DAO.Database, ByVal tableName As String) As String
    On Error GoTo ErrHandler
    Dim tdf As DAO.tableDef
    Set tdf = FreshModuleGuardTableDef(db, tableName)
    ModuleGuardTableInventoryLine = "Table [" & tableName & "]: present; fields=" & CStr(tdf.fields.count) & "; indexes=" & CStr(tdf.Indexes.count)
    Exit Function
ErrHandler:
    If Err.Number = MODULE_GUARD_ERROR_BASE + 50 Then
        ModuleGuardTableInventoryLine = "Table [" & tableName & "]: missing"
    Else
        ModuleGuardTableInventoryLine = "Table [" & tableName & "]: inspection failed; error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    End If
End Function

Private Function ModuleGuardQueryInventoryLine(ByVal db As DAO.Database, ByVal queryName As String) As String
    On Error GoTo ErrHandler
    Dim qdf As DAO.queryDef
    db.QueryDefs.Refresh
    Set qdf = FindQueryDef(db, queryName)
    If qdf Is Nothing Then
        ModuleGuardQueryInventoryLine = "Query [" & queryName & "]: missing"
    Else
        ModuleGuardQueryInventoryLine = "Query [" & queryName & "]: present; SQL length=" & CStr(Len(qdf.sql))
    End If
    Exit Function
ErrHandler:
    ModuleGuardQueryInventoryLine = "Query [" & queryName & "]: inspection failed; error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
End Function

Private Function ModuleGuardSafeDatabaseName(ByVal db As DAO.Database) As String
    On Error Resume Next
    ModuleGuardSafeDatabaseName = db.Name
    If Len(ModuleGuardSafeDatabaseName) = 0 Then ModuleGuardSafeDatabaseName = CurrentProject.FullName
    On Error GoTo 0
End Function

Private Function ResolveModuleUpdateGuardRegistrationEnablement(ByVal db As DAO.Database, ByVal requestedEnableAfterRegistration As Boolean, ByRef enablementDeferredForBootstrap As Boolean, ByRef detailText As String) As Boolean
    'Defers only the guard-enable write while setup has not yet created the Settings table or its single guard-setting row.
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim settingsTableName As String
    Dim settingRowCount As Long
    enablementDeferredForBootstrap = False
    detailText = vbNullString
    If requestedEnableAfterRegistration = False Then Exit Function
    settingsTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID)
    If ModuleGuardTableExists(db, settingsTableName) = False Then
        enablementDeferredForBootstrap = True
        detailText = "Settings table [" & settingsTableName & "] is not present. Exact-build registration and verification may proceed without forcing enablement; A2 will create the canonical EnableModuleUpdateGuard row with its default False value."
        LogModuleGuardMessage "The exact approved build will be registered and verified with guard enablement deferred until Settings creation.", detailText, TRACE_LEVEL_INFO, "ResolveModuleUpdateGuardRegistrationEnablement", "Resolve registration enablement"
        Exit Function
    End If
    Set rs = db.OpenRecordset("SELECT Count(*) AS " & SQL_FIELD_SETTING_ROW_COUNT & " FROM [" & settingsTableName & "] WHERE [" & FIELD_SETTING_KEY & "]=" & SharedDao_SqlQuote(SETTING_ENABLE_MODULE_UPDATE_GUARD), dbOpenSnapshot)
    settingRowCount = CLng(rs!settingRowCount)
    rs.Close
    Set rs = Nothing
    If settingRowCount = 0 Then
        enablementDeferredForBootstrap = True
        detailText = "Settings table [" & settingsTableName & "] exists, but the single [" & SETTING_ENABLE_MODULE_UPDATE_GUARD & "] row has not yet been created. Exact-build registration and verification may proceed without forcing enablement; A2 will insert the canonical row with its default False value."
        LogModuleGuardMessage "The exact approved build will be registered and verified with guard enablement deferred until Settings-row creation.", detailText, TRACE_LEVEL_INFO, "ResolveModuleUpdateGuardRegistrationEnablement", "Resolve registration enablement"
        Exit Function
    End If
    If settingRowCount <> 1 Then Err.Raise MODULE_GUARD_ERROR_BASE + 115, "ResolveModuleUpdateGuardRegistrationEnablement", "The EnableModuleUpdateGuard setting must have no more than one row before approved-build registration begins." & vbCrLf & "Rows found: " & CStr(settingRowCount)
    detailText = "Settings table [" & settingsTableName & "] and its single [" & SETTING_ENABLE_MODULE_UPDATE_GUARD & "] row are present; enablement will be preflighted and completed inside the registration transaction."
    ResolveModuleUpdateGuardRegistrationEnablement = True
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
    If savedErrNumber = 0 Then savedErrNumber = MODULE_GUARD_ERROR_BASE + 116
    If Len(Trim$(savedErrSource)) = 0 Then savedErrSource = "ResolveModuleUpdateGuardRegistrationEnablement"
    Err.Raise savedErrNumber, savedErrSource, "Could not resolve whether approved-build registration may defer guard enablement." & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Sub PreflightModuleUpdateGuardEnableTarget(ByVal db As DAO.Database, ByVal enableAfterRegistration As Boolean)
    'Validates every guard-enable write prerequisite before the approved-build registration transaction begins.
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim settingRowCount As Long
    If enableAfterRegistration = False Then Exit Sub
    If ModuleGuardTableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID)) = False Then Err.Raise MODULE_GUARD_ERROR_BASE + 77, "PreflightModuleUpdateGuardEnableTarget", "The Settings table is missing; approved-build registration cannot safely begin because guard enablement could not be completed."
    settingRowCount = DCount("*", ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID), "[" & FIELD_SETTING_KEY & "]=" & SharedDao_SqlQuote(SETTING_ENABLE_MODULE_UPDATE_GUARD))
    If settingRowCount <> 1 Then Err.Raise MODULE_GUARD_ERROR_BASE + 78, "PreflightModuleUpdateGuardEnableTarget", "The EnableModuleUpdateGuard setting must have exactly one row before approved-build registration begins." & vbCrLf & "Rows found: " & CStr(settingRowCount)
    Set rs = db.OpenRecordset("SELECT [" & FIELD_SETTING_KEY & "],[" & FIELD_SETTING_VALUE & "] FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID) & "] WHERE [" & FIELD_SETTING_KEY & "]=" & SharedDao_SqlQuote(SETTING_ENABLE_MODULE_UPDATE_GUARD), dbOpenDynaset)
    If rs.EOF Then Err.Raise MODULE_GUARD_ERROR_BASE + 79, "PreflightModuleUpdateGuardEnableTarget", "The EnableModuleUpdateGuard setting row was not available through an updateable recordset."
    If rs.Updatable = False Then Err.Raise MODULE_GUARD_ERROR_BASE + 80, "PreflightModuleUpdateGuardEnableTarget", "The EnableModuleUpdateGuard setting row is not updateable; approved-build registration cannot safely begin."
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
    Set rs = Nothing
    On Error GoTo 0
    If savedErrNumber = 0 Then savedErrNumber = MODULE_GUARD_ERROR_BASE + 81
    If Len(Trim$(savedErrSource)) = 0 Then savedErrSource = "PreflightModuleUpdateGuardEnableTarget"
    Err.Raise savedErrNumber, savedErrSource, "Could not preflight the Module Guard enable-setting target before approved-build registration." & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

' Purpose: Resolves module update guard final state token for I1 from the supplied/current ADC state, using `ResolveModuleUpdateGuardReadiness` and `ModuleGuardTableExists` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   detailText (Output/in-out; String): Carries human-readable diagnostic detail for the operation; when passed ByRef, the routine can return or augment the text for the caller.
' Returns: String: Returns the proof/state token produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function ModuleUpdateGuardFinalStateTokenForI1(ByRef detailText As String) As String
    'I1 reads final Guard infrastructure and the user's persisted enable/disable choice without registration, repair, or enforcement-state mutation.
    On Error GoTo NotCurrent
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim missingTableNames As String
    Dim infrastructureExists As Boolean
    Dim manifestPopulated As Boolean
    Dim SettingValue As Boolean
    Dim payloadText As String
    detailText = vbNullString
    Set db = currentDb
    ResolveModuleUpdateGuardReadiness infrastructureExists, manifestPopulated, missingTableNames
    If infrastructureExists = False Then detailText = "Module Guard infrastructure is incomplete. Missing: " & IIf(Len(missingTableNames) = 0, "(not identified)", missingTableNames): Exit Function
    If manifestPopulated = False Then detailText = "Module Guard infrastructure exists, but the active manifest is not populated.": Exit Function
    If ModuleGuardTableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID)) = False Then detailText = "The Settings table is missing, so I1 cannot prove the final Module Guard enable-setting state.": Exit Function
    Set rs = db.OpenRecordset("SELECT [" & FIELD_SETTING_VALUE & "] FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID) & "] WHERE [" & FIELD_SETTING_KEY & "]=" & SharedDao_SqlQuote(SETTING_ENABLE_MODULE_UPDATE_GUARD), dbOpenSnapshot)
    If rs.EOF Then detailText = "The EnableModuleUpdateGuard setting row is missing; A2 did not complete the canonical core Settings bootstrap contract.": rs.Close: Set rs = Nothing: Exit Function
    SettingValue = BooleanTextValue(Nz(rs.fields(FIELD_SETTING_VALUE), vbNullString), DEFAULT_ENABLE_MODULE_UPDATE_GUARD)
    rs.MoveNext
    If rs.EOF = False Then detailText = "More than one EnableModuleUpdateGuard setting row exists; I1 requires exactly one.": rs.Close: Set rs = Nothing: Exit Function
    rs.Close
    Set rs = Nothing
    payloadText = "ADCModuleGuardFinalStateI1V1|Infrastructure=Ready|ActiveManifest=Populated|EnableModuleUpdateGuard=" & CStr(SettingValue)
    ModuleUpdateGuardFinalStateTokenForI1 = payloadText
    detailText = "Module Guard infrastructure is ready, the active manifest is populated, and EnableModuleUpdateGuard=" & CStr(SettingValue) & ". I1 did not change the setting."
    Exit Function
NotCurrent:
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
    detailText = "Could not prove final Module Guard state for I1. Error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription
End Function

' Purpose: Sets module update guard enabled for setup for enabled.
' Public role: Public module entry point. This procedure is listed in the module header as an intentional externally callable ADC entry point and may be invoked directly by Access/VBA callers.
' Parameters:
'   enabled (Input; Boolean): Requested enabled state: True enables the managed feature/control; False disables it.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub SetModuleUpdateGuardEnabledForSetup(ByVal enabled As Boolean)
    'Setup Preflight calls this only after the user explicitly chooses Re-enable. It is never called for Leave disabled.
    'Module Guard retains the enablement decision and exact pre/post verification; TABLES owns only the physical Settings.SettingValue mutation.
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Set db = currentDb
    If ModuleGuardTableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID)) = False Then Err.Raise MODULE_GUARD_ERROR_BASE + 73, "SetModuleUpdateGuardEnabledForSetup", "The Settings table is missing; module-guard enforcement could not be enabled."
    sqlText = "SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID) & "] WHERE [" & FIELD_SETTING_KEY & "]=" & SharedDao_SqlQuote(SETTING_ENABLE_MODULE_UPDATE_GUARD)
    Set rs = db.OpenRecordset(sqlText, dbOpenDynaset)
    If rs.EOF Then rs.Close: Err.Raise MODULE_GUARD_ERROR_BASE + 74, "SetModuleUpdateGuardEnabledForSetup", "The EnableModuleUpdateGuard setting row is missing; run the current table repair before module-guard registration."
    rs.Close
    Set rs = Nothing
    SetSettingRawText SETTING_ENABLE_MODULE_UPDATE_GUARD, IIf(enabled, "True", "False"), False, False, False
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Err.Raise MODULE_GUARD_ERROR_BASE + 75, "SetModuleUpdateGuardEnabledForSetup", "The EnableModuleUpdateGuard setting row disappeared during verification."
    If BooleanTextValue(Nz(rs.fields(FIELD_SETTING_VALUE), vbNullString), DEFAULT_ENABLE_MODULE_UPDATE_GUARD) <> enabled Then rs.Close: Err.Raise MODULE_GUARD_ERROR_BASE + 76, "SetModuleUpdateGuardEnabledForSetup", "The EnableModuleUpdateGuard setting did not persist the requested value."
    rs.Close
End Sub

Private Function BuildModuleGuardVerificationFailureTranscript(ByVal failureText As String) As String
    BuildModuleGuardVerificationFailureTranscript = String$(80, "-") & vbCrLf & _
        DEFAULT_UTILITY_SHORT_NAME & " Module Guard verification failure" & vbCrLf & _
        failureText & vbCrLf & _
        String$(80, "-")
End Function

Private Function IsModuleGuardVerificationFailureTranscript(ByVal transcriptText As String) As Boolean
    IsModuleGuardVerificationFailureTranscript = (InStr(1, transcriptText, DEFAULT_UTILITY_SHORT_NAME & " Module Guard verification failure" & vbCrLf & DEFAULT_UTILITY_LONG_NAME & " module update verification failed.", vbBinaryCompare) > 0)
End Function

Private Sub PersistModuleGuardVerificationFailureTranscript(ByVal transcriptText As String, ByVal failureNumber As Long, ByVal failureSource As String)
    Dim traceErrNumber As Long
    Dim traceErrSource As String
    Dim traceErrDescription As String
    On Error Resume Next
    LogTraceMessage TRACE_LEVEL_ERROR, "ModuleUpdateGuard", "VerifyModuleUpdateGuard", DEFAULT_UTILITY_LONG_NAME & " module update verification failed.", transcriptText, "Module update guard", failureNumber, failureSource, transcriptText, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionManifest"), "Table", "CodeDb", "ModuleUpdateGuard"
    traceErrNumber = Err.Number
    traceErrSource = Err.Source
    traceErrDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If traceErrNumber <> 0 Then
        Debug.Print DEFAULT_UTILITY_SHORT_NAME & " Module Guard verification transcript logging failed. Err.Number=" & CStr(traceErrNumber) & "; Err.Source=" & traceErrSource & "; Err.Description=" & traceErrDescription
    End If
End Sub

Private Sub LogModuleGuardMessage(ByVal messageText As String, ByVal detailText As String, Optional ByVal severityText As String = TRACE_LEVEL_INFO, Optional ByVal sourceProcedure As String = "x__UTIL_ADC_MODULE_GUARD", Optional ByVal stageText As String = vbNullString)
    Dim traceDetail As String
    traceDetail = detailText
    If Len(Trim$(stageText)) > 0 Then
        If Len(Trim$(traceDetail)) > 0 Then traceDetail = traceDetail & vbCrLf
        traceDetail = traceDetail & "Stage: " & stageText
    End If
    Debug.Print "ADCModuleGuard [" & severityText & "] " & sourceProcedure & ": " & messageText
    If Len(Trim$(traceDetail)) > 0 Then Debug.Print Replace(traceDetail, vbCrLf, " | ")
    On Error Resume Next
    LogTraceMessage severityText, "ModuleUpdateGuard", sourceProcedure, messageText, traceDetail, "Module update guard", 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionManifest"), "Table", "CodeDb", "ModuleUpdateGuard"
    On Error GoTo 0
End Sub

Private Function ManifestTimestampText(ByVal timestampValue As Variant) As String
    If IsNull(timestampValue) Or IsEmpty(timestampValue) Then Exit Function
    ManifestTimestampText = FormatChatGPTVersionTimestamp(CDate(timestampValue))
End Function

Private Function DdlTypeText(ByVal fieldType As Long, ByVal fieldSize As Long) As String
    Select Case fieldType
        Case dbText
            DdlTypeText = "TEXT(" & CStr(fieldSize) & ")"
        Case dbMemo
            DdlTypeText = "LONGTEXT"
        Case dbLong
            DdlTypeText = "LONG"
        Case dbDate
            DdlTypeText = "DATETIME"
        Case dbBoolean
            DdlTypeText = "YESNO"
        Case Else
            Err.Raise MODULE_GUARD_ERROR_BASE + 34, "DdlTypeText", "Unsupported module-guard field type: " & CStr(fieldType)
    End Select
End Function

Private Function ComponentHeaderSourceText(ByVal component As Object) As String
    Const HEADER_READ_LINE_LIMIT As Long = 512
    Dim codeModule As Object
    Dim moduleLineCount As Long
    Dim linesToRead As Long
    Set codeModule = component.codeModule
    moduleLineCount = CLng(codeModule.CountOfLines)
    If moduleLineCount = 0 Then Exit Function
    linesToRead = moduleLineCount
    If linesToRead > HEADER_READ_LINE_LIMIT Then linesToRead = HEADER_READ_LINE_LIMIT
    ComponentHeaderSourceText = NormalizeSourceText(CStr(codeModule.lines(1, linesToRead)))
    If moduleLineCount > linesToRead And Len(ModuleGuardDeclaredVBName(ComponentHeaderSourceText)) = 0 Then
        ComponentHeaderSourceText = NormalizeSourceText(CStr(codeModule.lines(1, moduleLineCount)))
    End If
End Function

' Purpose: Resolves and returns module guard chat gpt version ID from component from the supplied inputs and the authoritative ADC definitions/state used by this module.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   component (Input; Object): VBE VBComponent object whose module source/header is being inspected for Module Guard version metadata.
' Returns: String: Returns the resolved module guard chat gpt version ID from component text in the canonical representation consumed by this module's caller; vbNullString represents no applicable value where
'   the existing contract permits it.
Public Function ModuleGuardChatGPTVersionIDFromComponent(ByVal component As Object) As String
    'Canonical cross-module rewrite-ID accessor. It reads only a bounded leading source region and delegates parsing and validity semantics to the existing Module Guard metadata parser/validator.
    On Error GoTo Unavailable
    Const VERSION_ID_READ_LINE_LIMIT As Long = 160
    Dim codeModule As Object
    Dim moduleLineCount As Long
    Dim linesToRead As Long
    Dim sourceText As String
    Dim versionID As String
    Dim previousVersionID As String
    Dim timestampValue As Variant
    Dim headerName As String
    Dim timestampText As String
    If component Is Nothing Then Exit Function
    Set codeModule = component.codeModule
    moduleLineCount = CLng(codeModule.CountOfLines)
    If moduleLineCount <= 0 Then Exit Function
    linesToRead = moduleLineCount
    If linesToRead > VERSION_ID_READ_LINE_LIMIT Then linesToRead = VERSION_ID_READ_LINE_LIMIT
    sourceText = NormalizeSourceText(CStr(codeModule.lines(1, linesToRead)))
    ReadNormalizedModuleGuardSourceMetadata sourceText, versionID, previousVersionID, timestampValue, headerName, timestampText
    If IsValidChatGPTVersionID(versionID) Then ModuleGuardChatGPTVersionIDFromComponent = versionID
    Exit Function
Unavailable:
    Err.Clear
End Function
Public Function ModuleGuardChatGPTVersionIDByModuleName(ByVal ModuleName As String) As String
    'Returns the current managed rewrite-header Version ID for one module in the current database project without embedding that ID in executable code.
    On Error GoTo Unavailable
    Dim projectObject As Object
    Dim component As Object
    Set projectObject = GuardVBProject()
    Set component = projectObject.VBComponents(ModuleName)
    ModuleGuardChatGPTVersionIDByModuleName = ModuleGuardChatGPTVersionIDFromComponent(component)
    Exit Function
Unavailable:
    Err.Clear
End Function


Private Function NormalizeSourceText(ByVal sourceText As String) As String
    Dim lines As Variant
    Dim lowerIndex As Long
    Dim upperIndex As Long
    sourceText = Replace(sourceText, vbCrLf, vbLf)
    sourceText = Replace(sourceText, vbCr, vbLf)
    lines = Split(sourceText, vbLf)
    lowerIndex = LBound(lines)
    upperIndex = UBound(lines)
    Do While upperIndex >= lowerIndex
        PumpModuleGuardSetupResponsiveness
        If Len(CStr(lines(upperIndex))) > 0 Then Exit Do
        upperIndex = upperIndex - 1
    Loop
    If upperIndex < lowerIndex Then Exit Function
    ReDim Preserve lines(lowerIndex To upperIndex)
    NormalizeSourceText = Join(lines, vbCrLf)
End Function

Private Function GetComponentTypeText(ByVal ComponentType As Long) As String
    Select Case ComponentType
        Case 1
            GetComponentTypeText = MODULE_GUARD_COMPONENT_STANDARD
        Case 2
            GetComponentTypeText = MODULE_GUARD_COMPONENT_CLASS
        Case 3
            GetComponentTypeText = MODULE_GUARD_COMPONENT_FORM
        Case 100
            GetComponentTypeText = MODULE_GUARD_COMPONENT_DOCUMENT
        Case Else
            GetComponentTypeText = "Type " & CStr(ComponentType)
    End Select
End Function

Private Sub ConfigureTextDictionary(ByVal dictionaryObject As Object)
    dictionaryObject.CompareMode = vbTextCompare
End Sub

Private Function ApprovedComponentTypeTransitionIsAuthorized(ByVal incomingDefinition As Variant, ByVal activeType As String, ByVal activeCurrent As String, ByRef detailText As String) As Boolean
    'A component-type migration is eligible only on the direct managed-header predecessor; the caller separately requires matching persistent active-manifest/history evidence before registration.
    Dim sourcePreviousVersionID As String
    sourcePreviousVersionID = Trim$(CStr(incomingDefinition(3)))
    If StrComp(sourcePreviousVersionID, Trim$(activeCurrent), vbBinaryCompare) <> 0 Then
        detailText = "Component-type migration requires the incoming managed-header Previous Version ID to equal the active approved Version ID. Active=" & DisplayGuardValue(activeCurrent) & "; Header previous=" & DisplayGuardValue(sourcePreviousVersionID)
        Exit Function
    End If
    detailText = "Direct managed-header predecessor matched the active approved Version ID for component-type migration: " & activeType & " " & activeCurrent & " -> " & CStr(incomingDefinition(1)) & " " & CStr(incomingDefinition(2))
    ApprovedComponentTypeTransitionIsAuthorized = True
End Function

Private Function ApprovedDefinitionTransitionDetail(ByVal ModuleName As String, ByVal reasonText As String, ByVal activeCurrent As String, ByVal activePrevious As String, ByVal incomingCurrent As String, ByVal incomingPrevious As String) As String
    ApprovedDefinitionTransitionDetail = _
        "Module: " & ModuleName & vbCrLf & _
        "Reason: " & reasonText & vbCrLf & _
        "Active manifest Current ID: " & DisplayGuardValue(activeCurrent) & vbCrLf & _
        "Active manifest Previous ID: " & DisplayGuardValue(activePrevious) & vbCrLf & _
        "Incoming Current ID: " & DisplayGuardValue(incomingCurrent) & vbCrLf & _
        "Incoming managed-header predecessor ID: " & DisplayGuardValue(incomingPrevious)
End Function

Private Sub EnsurePrimaryIndexExact(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String)
    Dim tdf As DAO.tableDef
    Dim idx As DAO.index
    Dim primaryIndex As DAO.index
    Set tdf = FreshModuleGuardTableDef(db, tableName)
    For Each idx In tdf.Indexes
        PumpModuleGuardSetupResponsiveness
        If idx.Primary Then
            Set primaryIndex = idx
            Exit For
        End If
    Next idx
    If primaryIndex Is Nothing Then
        Set tdf = Nothing
        db.Execute "ALTER TABLE [" & tableName & "] ADD CONSTRAINT [" & INDEX_PRIMARY_KEY & "] PRIMARY KEY ([" & fieldName & "])", dbFailOnError
        Set db = Nothing
        Set db = currentDb
        Set tdf = FreshModuleGuardTableDef(db, tableName)
        For Each idx In tdf.Indexes
            PumpModuleGuardSetupResponsiveness
            If idx.Primary Then
                Set primaryIndex = idx
                Exit For
            End If
        Next idx
    End If
    If primaryIndex Is Nothing Then Err.Raise MODULE_GUARD_ERROR_BASE + 39, "EnsurePrimaryIndexExact", "Primary-key creation could not be verified." & vbCrLf & "Table: " & tableName
    If StrComp(IndexFieldNamesText(primaryIndex), fieldName, vbTextCompare) <> 0 Or primaryIndex.Unique = False Then Err.Raise MODULE_GUARD_ERROR_BASE + 43, "EnsurePrimaryIndexExact", "The existing primary key conflicts with the required module-guard definition." & vbCrLf & "Table: " & tableName & vbCrLf & "Expected field: " & fieldName & vbCrLf & "Actual fields: " & IndexFieldNamesText(primaryIndex)
End Sub

Private Function ActiveManifestHasRegisteredHistoryEvidence(ByVal historyByKey As Object, ByVal ModuleName As String, ByVal activeType As String, ByVal activeCurrent As String, ByVal activePrevious As String, ByVal activeTimestampText As String, ByRef detailText As String) As Boolean
    On Error GoTo ErrHandler
    Dim historyData As Variant
    Dim historyKey As String
    Dim historyTimestampText As String
    If IsValidChatGPTVersionID(activeCurrent) = False Then AppendGuardDetail detailText, "The active manifest Version ID is invalid. Module=" & ModuleName & "; Version=" & DisplayGuardValue(activeCurrent): Exit Function
    If historyByKey Is Nothing Then AppendGuardDetail detailText, "Append-only history evidence was not loaded for the approved-build preflight. Module=" & ModuleName: Exit Function
    historyKey = ModuleGuardHistoryKey(ModuleName, activeCurrent)
    If historyByKey.Exists(historyKey) = False Then AppendGuardDetail detailText, "Append-only history has no row for the active manifest version. Module=" & ModuleName & "; Version=" & activeCurrent: Exit Function
    historyData = historyByKey(historyKey)
    If StrComp(CStr(historyData(2)), activeType, vbTextCompare) <> 0 Then AppendGuardDetail detailText, "Active manifest/history component type mismatch. Module=" & ModuleName: Exit Function
    If StrComp(CStr(historyData(0)), activePrevious, vbBinaryCompare) <> 0 Then AppendGuardDetail detailText, "Active manifest/history Previous Version ID mismatch. Module=" & ModuleName: Exit Function
    historyTimestampText = ManifestTimestampText(historyData(3))
    If Len(activeTimestampText) = 0 Or StrComp(historyTimestampText, activeTimestampText, vbBinaryCompare) <> 0 Then AppendGuardDetail detailText, "Active manifest/history timestamp mismatch. Module=" & ModuleName & "; Manifest=" & DisplayGuardValue(activeTimestampText) & "; History=" & DisplayGuardValue(historyTimestampText): Exit Function
    ActiveManifestHasRegisteredHistoryEvidence = True
    Exit Function
ErrHandler:
    AppendGuardDetail detailText, "Could not prove active manifest/history registration evidence. Module=" & ModuleName & "; Err.Number=" & CStr(Err.Number) & "; Err.Source=" & Err.Source & "; Err.Description=" & Err.Description
End Function




Private Sub WriteManifestRowOnly(ByVal db As DAO.Database, ByVal updateBatchID As String, ByVal rowData As Variant, ByVal baselineAccepted As Boolean)
    Dim rs As DAO.Recordset
    Dim ModuleName As String
    ModuleName = CStr(rowData(0))
    Set rs = db.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionManifest") & "] WHERE " & SQL_FIELD_MODULE_NAME & "=" & SharedDao_SqlQuote(ModuleName), dbOpenDynaset)
    If rs.EOF Then rs.AddNew Else rs.Edit
    rs.fields(FIELD_MODULE_NAME) = ModuleName
    rs.fields(FIELD_COMPONENT_TYPE) = CStr(rowData(1))
    rs!expectedVersionID = CStr(rowData(2))
    rs!expectedPreviousVersionID = Shared_String_ConvertToNullIfBlank(Trim$(CStr(rowData(3))))
    rs.fields(FIELD_EXPECTED_VERSION_TIMESTAMP) = rowData(4)
    rs!updateBatchID = updateBatchID
    rs.fields(FIELD_EXPECTED_CHANGE_DESCRIPTION) = CStr(rowData(5))
    rs!isRequired = True
    rs.fields(FIELD_ACTIVE).value = True
    rs!baselineAccepted = baselineAccepted
    rs.fields(FIELD_MANIFEST_UPDATED_ON) = Now
    rs.fields(FIELD_LAST_VERIFIED_ON) = Null
    rs.fields(FIELD_LAST_VERIFICATION_RESULT) = Null
    rs.fields(FIELD_LAST_VERIFICATION_NOTES) = Null
    rs.Update
    rs.Close
End Sub

Private Sub UpdateManifestVerificationRow(ByVal manifest As DAO.Recordset, ByVal verifiedOn As Date, ByVal resultCode As String, ByVal resultNotes As String)
    manifest.Edit
    manifest.fields(FIELD_LAST_VERIFIED_ON) = verifiedOn
    manifest.fields(FIELD_LAST_VERIFICATION_RESULT) = resultCode
    manifest.fields(FIELD_LAST_VERIFICATION_NOTES) = Shared_String_ConvertToNullIfBlank(Trim$(resultNotes))
    manifest.Update
End Sub

Private Sub UpsertManifestRow(ByVal db As DAO.Database, ByVal updateBatchID As String, ByVal rowData As Variant, ByVal baselineAccepted As Boolean)
    Dim rs As DAO.Recordset
    Dim ModuleName As String
    ModuleName = CStr(rowData(0))
    db.Execute "UPDATE [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleUpdateHistory") & "] SET " & SQL_FIELD_SUPERSEDED_ON & "=Now() WHERE " & SQL_FIELD_MODULE_NAME & "=" & SharedDao_SqlQuote(ModuleName) & " AND " & SQL_FIELD_CHAT_GPT_VERSION_ID & "<>" & SharedDao_SqlQuote(CStr(rowData(2))) & " AND " & SQL_FIELD_SUPERSEDED_ON & " Is Null", dbFailOnError
    Set rs = db.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionManifest") & "] WHERE " & SQL_FIELD_MODULE_NAME & "=" & SharedDao_SqlQuote(ModuleName), dbOpenDynaset)
    If rs.EOF Then rs.AddNew Else rs.Edit
    rs.fields(FIELD_MODULE_NAME) = ModuleName
    rs.fields(FIELD_COMPONENT_TYPE) = CStr(rowData(1))
    rs!expectedVersionID = CStr(rowData(2))
    rs!expectedPreviousVersionID = Shared_String_ConvertToNullIfBlank(Trim$(CStr(rowData(3))))
    rs.fields(FIELD_EXPECTED_VERSION_TIMESTAMP) = rowData(4)
    rs!updateBatchID = updateBatchID
    rs.fields(FIELD_EXPECTED_CHANGE_DESCRIPTION) = CStr(rowData(5))
    rs!isRequired = True
    rs.fields(FIELD_ACTIVE).value = True
    rs!baselineAccepted = baselineAccepted
    rs.fields(FIELD_MANIFEST_UPDATED_ON) = Now
    rs.fields(FIELD_LAST_VERIFIED_ON) = Null
    rs.fields(FIELD_LAST_VERIFICATION_RESULT) = Null
    rs.fields(FIELD_LAST_VERIFICATION_NOTES) = Null
    rs.Update
    rs.Close
End Sub


' =============================================================================================================================================================================
' APPROVED BUILD REGISTRATION PROCEDURES
' =============================================================================================================================================================================

' Purpose: Attempts to perform register approved build for registration options.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   registrationResult (Output; TModuleGuardRegistrationResult): Receives the registration result produced by the operation so the caller can inspect the outcome without re-running the work.
'   registrationOptions (Input Optional; ModuleGuardRegistrationOptions): Option flags selecting the optional registration behaviors enabled for this invocation. Default:
'   MODULE_GUARD_REGISTRATION_OPTION_ENABLE_AFTER_REGISTRATION.
' Returns: Boolean: True when the attempt to register approved build completes successfully; False for the routine's defined normal unsuccessful/not-applicable outcome.
Public Function TryRegisterApprovedBuild(ByRef registrationResult As TModuleGuardRegistrationResult, Optional ByVal registrationOptions As ModuleGuardRegistrationOptions = MODULE_GUARD_REGISTRATION_OPTION_ENABLE_AFTER_REGISTRATION) As Boolean
    On Error GoTo ErrHandler
    Dim enableAfterRegistration As Boolean
    Dim explicitlyApproveInstalledBuild As Boolean
    Dim registrationRecordedBy As String
    Dim db As DAO.Database
    Dim ws As DAO.workspace
    Dim definitions As Collection
    Dim initialBaselineDefinitions As Collection
    Dim installedVersionByName As Object
    Dim inventoryEvidence As Collection
    Dim componentsByName As Object
    Dim timestampTextByName As Object
    Dim typeByName As Object
    Dim versionByName As Object
    Dim previousByName As Object
    Dim timestampByName As Object
    Dim headerNameByName As Object
    Dim versionCounts As Object
    Dim installedCount As Long
    Dim Definition As Variant
    Dim componentData As Collection
    Dim rowData As Variant
    Dim ModuleName As String
    Dim actualType As String
    Dim actualVersion As String
    Dim actualPrevious As String
    Dim actualTimestamp As Variant
    Dim actualHeaderName As String
    Dim verificationDetails As String
    Dim cleanupDetails As String
    Dim transactionStarted As Boolean
    Dim rollbackVerified As Boolean
    Dim comparisonDetail As String
    Dim manifestWriteStarted As Boolean
    Dim approvedDefinitionMismatchText As String
    Dim approvedDefinitionMismatchCount As Long
    Dim effectiveEnableAfterRegistration As Boolean
    Dim enablementDeferredForBootstrap As Boolean
    Dim enablementResolutionDetails As String
    Dim registrationStage As String
    Dim registrationOutcomeMessage As String
    Dim definitionNumber As Long
    Dim manifestWriteNumber As Long
    Dim namingSnapshotStarted As Boolean
    Dim namingSnapshotTimer As Single
    enableAfterRegistration = ((registrationOptions And MODULE_GUARD_REGISTRATION_OPTION_ENABLE_AFTER_REGISTRATION) <> 0)
    explicitlyApproveInstalledBuild = ((registrationOptions And MODULE_GUARD_REGISTRATION_OPTION_EXPLICITLY_APPROVE_INSTALLED_BUILD) <> 0)
    If explicitlyApproveInstalledBuild Then registrationRecordedBy = MODULE_GUARD_RECORDED_BY_MANUAL Else registrationRecordedBy = MODULE_GUARD_RECORDED_BY_APPROVED_BUILD
    registrationResult.failedErrNumber = 0
    registrationResult.failedErrSource = vbNullString
    registrationResult.failedErrDescription = vbNullString
    registrationResult.RegistrationDetails = vbNullString
    registrationResult.installedVersionSummary = vbNullString
    registrationResult.ExplicitApprovalCandidate = False
    registrationStage = "Prepare registration naming snapshot"
    namingSnapshotTimer = Timer
    LogModuleGuardPreflightBoundary "TryRegisterApprovedBuild", "Module Guard approved-build registration: naming snapshot starting."
    BeginNamingSnapshot
    namingSnapshotStarted = True
    LogModuleGuardPreflightBoundary "TryRegisterApprovedBuild", "Module Guard approved-build registration: naming snapshot ready.", "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(namingSnapshotTimer))
    registrationStage = "Capture or consume installed module inventory"
    If Not mModuleGuardInstalledInventoryEvidence Is Nothing Then Set inventoryEvidence = mModuleGuardInstalledInventoryEvidence
    If inventoryEvidence Is Nothing Then
        Set initialBaselineDefinitions = InitialModuleGuardBaselineDefinitions()
        EnsureModuleGuardOwnedComponentLookup initialBaselineDefinitions
        LogModuleGuardPreflightBoundary "TryRegisterApprovedBuild", "Module Guard approved-build registration: one-pass installed inventory/header validation starting."
        Set inventoryEvidence = CaptureInstalledModuleInventoryEvidence(initialBaselineDefinitions)
        LogModuleGuardPreflightBoundary "TryRegisterApprovedBuild", "Module Guard approved-build registration: one-pass installed inventory/header validation completed."
    Else
        Set initialBaselineDefinitions = inventoryEvidence("ExpectedDefinitions")
        EnsureModuleGuardOwnedComponentLookup initialBaselineDefinitions
        LogModuleGuardPreflightBoundary "TryRegisterApprovedBuild", "Module Guard approved-build registration: authoritative installed-module inventory and canonical definitions reused; duplicate VBProject/header inspection skipped."
    End If
    Set mModuleGuardInstalledInventoryEvidence = inventoryEvidence
    LoadInstalledModuleInventoryEvidence inventoryEvidence, componentsByName, timestampTextByName, typeByName, versionByName, previousByName, timestampByName, headerNameByName, versionCounts, installedCount
    registrationStage = "Repair Module Guard structure"
    LogModuleGuardPreflightBoundary "TryRegisterApprovedBuild", "Module Guard approved-build registration: structure repair starting."
    RepairModuleUpdateGuard False, False
    LogModuleGuardPreflightBoundary "TryRegisterApprovedBuild", "Module Guard approved-build registration: structure repair completed."
    If explicitlyApproveInstalledBuild Then
        registrationStage = "Validate explicitly approved installed coordinated build"
        LogModuleGuardPreflightBoundary "TryRegisterApprovedBuild", "Module Guard registration: explicit installed-build validation starting."
        Set definitions = ExplicitlyApprovedInstalledBuildDefinitions(componentData, initialBaselineDefinitions, inventoryEvidence)
        LogModuleGuardPreflightBoundary "TryRegisterApprovedBuild", "Module Guard registration: explicit installed-build validation completed.", "Definitions: " & CStr(definitions.count)
    Else
        registrationStage = "Validate managed-header coordinated-build definitions"
        Set definitions = initialBaselineDefinitions
        LogModuleGuardPreflightBoundary "TryRegisterApprovedBuild", "Module Guard approved-build registration: managed-header definition validation starting.", "Definitions: " & CStr(definitions.count)
        ValidateApprovedBuildDefinitions definitions
        LogModuleGuardPreflightBoundary "TryRegisterApprovedBuild", "Module Guard approved-build registration: managed-header definition validation completed."
        ' INTENTIONAL OPERATION SNAPSHOT: componentData captures one coherent source state for this invocation of TryRegisterApprovedBuild; validity ends with the operation rather than becoming a live data
        ' authority.
        Set componentData = New Collection
        registrationStage = "Inspect installed headers against approved definitions"
        LogModuleGuardPreflightBoundary "TryRegisterApprovedBuild", "Module Guard approved-build registration: installed-header comparison loop starting.", "Modules: " & CStr(definitions.count)
        For Each Definition In definitions
            PumpModuleGuardSetupResponsiveness
            definitionNumber = definitionNumber + 1
            ModuleName = CStr(Definition(0))
            LogModuleGuardPreflightBoundary "TryRegisterApprovedBuild", "Module Guard approved-build registration: module " & CStr(definitionNumber) & " of " & CStr(definitions.count) & " header inspection starting.", "Module: " & ModuleName
            actualType = vbNullString
            actualVersion = vbNullString
            actualPrevious = vbNullString
            actualTimestamp = Null
            actualHeaderName = vbNullString
            If componentsByName.Exists(ModuleName) = False Then Err.Raise MODULE_GUARD_ERROR_BASE + 40, "TryRegisterApprovedBuild", "Required VBA component is not installed." & vbCrLf & "Module: " & ModuleName
            actualType = CStr(typeByName(ModuleName))
            actualVersion = CStr(versionByName(ModuleName))
            actualPrevious = CStr(previousByName(ModuleName))
            actualTimestamp = timestampByName(ModuleName)
            actualHeaderName = CStr(headerNameByName(ModuleName))
            comparisonDetail = ApprovedBuildComparisonDiagnostic(ModuleName, Definition, actualType, actualVersion, actualPrevious, actualTimestamp, actualHeaderName, manifestWriteStarted)
            If StrComp(actualType, CStr(Definition(1)), vbTextCompare) <> 0 Then
                approvedDefinitionMismatchCount = approvedDefinitionMismatchCount + 1
                AppendGuardDetail approvedDefinitionMismatchText, "Component type mismatch:" & vbCrLf & comparisonDetail
            End If
            If StrComp(actualVersion, CStr(Definition(2)), vbBinaryCompare) <> 0 Then
                approvedDefinitionMismatchCount = approvedDefinitionMismatchCount + 1
                AppendGuardDetail approvedDefinitionMismatchText, "Current Version ID mismatch:" & vbCrLf & comparisonDetail
            End If
            If Len(CStr(Definition(3))) > 0 And StrComp(actualPrevious, CStr(Definition(3)), vbBinaryCompare) <> 0 Then
                approvedDefinitionMismatchCount = approvedDefinitionMismatchCount + 1
                AppendGuardDetail approvedDefinitionMismatchText, "Previous Version ID mismatch:" & vbCrLf & comparisonDetail
            End If
            If StrComp(actualHeaderName, ModuleName, vbBinaryCompare) <> 0 Then
                approvedDefinitionMismatchCount = approvedDefinitionMismatchCount + 1
                AppendGuardDetail approvedDefinitionMismatchText, "Module identity mismatch:" & vbCrLf & comparisonDetail
            End If
            If Len(CStr(Definition(4))) > 0 Then
                If IsNull(actualTimestamp) Then
                    approvedDefinitionMismatchCount = approvedDefinitionMismatchCount + 1
                    AppendGuardDetail approvedDefinitionMismatchText, "Installed timestamp is missing:" & vbCrLf & comparisonDetail
                ElseIf FormatChatGPTVersionTimestamp(CDate(actualTimestamp)) <> CStr(Definition(4)) Then
                    approvedDefinitionMismatchCount = approvedDefinitionMismatchCount + 1
                    AppendGuardDetail approvedDefinitionMismatchText, "Approved-definition timestamp mismatch:" & vbCrLf & comparisonDetail
                End If
            End If
            componentData.Add Array(ModuleName, actualType, actualVersion, actualPrevious, actualTimestamp, vbNullString)
            LogModuleGuardPreflightBoundary "TryRegisterApprovedBuild", "Module Guard approved-build registration: module " & CStr(definitionNumber) & " of " & CStr(definitions.count) & " header inspection completed.", "Module: " & ModuleName
        Next Definition
        LogModuleGuardPreflightBoundary "TryRegisterApprovedBuild", "Module Guard approved-build registration: installed-header comparison loop completed.", "Modules inspected: " & CStr(definitionNumber) & vbCrLf & "Definition mismatches: " & CStr(approvedDefinitionMismatchCount)
        If approvedDefinitionMismatchCount > 0 Then
            registrationResult.ExplicitApprovalCandidate = True
            LogModuleGuardMessage "Approved current build preflight found all embedded-definition mismatches; no manifest write started.", "Mismatch count: " & CStr(approvedDefinitionMismatchCount) & vbCrLf & approvedDefinitionMismatchText, TRACE_LEVEL_ERROR, "TryRegisterApprovedBuild", "Preflight complete approved build"
            Err.Raise MODULE_GUARD_ERROR_BASE + 71, "TryRegisterApprovedBuild", "Approved-build preflight found " & CStr(approvedDefinitionMismatchCount) & " embedded-definition mismatch(es). No history or manifest rows were written." & vbCrLf & vbCrLf & approvedDefinitionMismatchText
        End If
    End If
    ' INTENTIONAL KEYED LOOKUP: installedVersionByName provides keyed membership/lookup over source rows or state already read by TryRegisterApprovedBuild; it is a derived access path, not a second authority.
    Set installedVersionByName = CreateObject("Scripting.Dictionary")
    ConfigureTextDictionary installedVersionByName
    For Each rowData In componentData
        installedVersionByName(CStr(rowData(0))) = CStr(rowData(2))
    Next rowData
    registrationResult.installedVersionSummary = InstalledModuleVersionSummary(initialBaselineDefinitions, installedVersionByName)
    Set db = currentDb
    registrationStage = "Validate Module Guard definitions build participation"
    ValidateDefinitionsModuleBuildParticipation db, definitions
    If explicitlyApproveInstalledBuild Then
        registrationStage = "Apply explicit active-manifest transition override"
        LogModuleGuardPreflightBoundary "TryRegisterApprovedBuild", "Module Guard explicit approval: active-manifest lineage/timestamp transition checks intentionally overridden by the user's Setup Yes choice.", "The exact installed coordinated build already passed fixed-roster, component-type, managed-header, Version-ID format/uniqueness, and timestamp validation. Existing exact-version history remains immutable and is verified transactionally before commit.", TRACE_LEVEL_WARNING
    Else
        registrationStage = "Preflight active-manifest transition"
        LogModuleGuardPreflightBoundary "TryRegisterApprovedBuild", "Module Guard approved-build registration: manifest transition preflight starting."
        PreflightApprovedBuildTransition db, definitions
        LogModuleGuardPreflightBoundary "TryRegisterApprovedBuild", "Module Guard approved-build registration: manifest transition preflight completed."
    End If
    registrationStage = "Resolve and verify guard enablement target"
    effectiveEnableAfterRegistration = ResolveModuleUpdateGuardRegistrationEnablement(db, enableAfterRegistration, enablementDeferredForBootstrap, enablementResolutionDetails)
    PreflightModuleUpdateGuardEnableTarget db, effectiveEnableAfterRegistration
    LogModuleGuardPreflightBoundary "TryRegisterApprovedBuild", "Module Guard approved-build registration: guard enablement target verified.", enablementResolutionDetails
    Set ws = DBEngine.Workspaces(0)
    If mModuleGuardRegistrationTransactionOwned Then Err.Raise MODULE_GUARD_ERROR_BASE + 69, "TryRegisterApprovedBuild", "Approved-build registration is already active in this module. A second registration transaction cannot start."
    registrationStage = "Begin approved-build registration transaction"
    LogModuleGuardPreflightBoundary "TryRegisterApprovedBuild", "Module Guard approved-build registration: transaction starting."
    ws.BeginTrans
    transactionStarted = True
    mModuleGuardRegistrationTransactionOwned = True
    mModuleGuardNestedVerificationTransactionActive = False
    'History, current manifest, verification snapshot, verification-status updates, and optional enable switch all remain inside one Workspace transaction.
    'VerifyModuleUpdateGuard uses a nested Workspace transaction; its commit closes only the nested level, leaving this outer registration transaction available for complete rollback.
    registrationStage = "Deactivate manifest rows retired from current Module Guard ownership"
    DeactivateNonOwnedManifestRows db
    manifestWriteStarted = True
    registrationStage = "Write approved history and manifest rows"
    LogModuleGuardPreflightBoundary "TryRegisterApprovedBuild", "Module Guard approved-build registration: history and manifest write loop starting.", "Rows: " & CStr(componentData.count)
    For Each rowData In componentData
        PumpModuleGuardSetupResponsiveness
        manifestWriteNumber = manifestWriteNumber + 1
        LogModuleGuardPreflightBoundary "TryRegisterApprovedBuild", "Module Guard approved-build registration: write row " & CStr(manifestWriteNumber) & " of " & CStr(componentData.count) & ".", "Module: " & CStr(rowData(0))
        InsertOrVerifyApprovedHistoryRow db, MODULE_GUARD_APPROVED_CURRENT_BUILD_BATCH_ID, rowData, registrationRecordedBy
        UpsertManifestRow db, MODULE_GUARD_APPROVED_CURRENT_BUILD_BATCH_ID, rowData, True
    Next rowData
    LogModuleGuardPreflightBoundary "TryRegisterApprovedBuild", "Module Guard approved-build registration: history and manifest write loop completed.", "Rows written or verified: " & CStr(manifestWriteNumber)
    registrationStage = "Reverify staged approved build"
    LogModuleGuardPreflightBoundary "TryRegisterApprovedBuild", "Module Guard approved-build registration: staged-build verification starting."
    VerifyModuleUpdateGuard False, verificationDetails, True, registrationResult.installedVersionSummary
    LogModuleGuardPreflightBoundary "TryRegisterApprovedBuild", "Module Guard approved-build registration: staged-build verification completed.", verificationDetails
    If effectiveEnableAfterRegistration Then SetModuleUpdateGuardEnabledForSetup True
    registrationStage = "Commit approved-build registration transaction"
    LogModuleGuardPreflightBoundary "TryRegisterApprovedBuild", "Module Guard approved-build registration: transaction commit starting."
    ws.CommitTrans
    transactionStarted = False
    LogModuleGuardPreflightBoundary "TryRegisterApprovedBuild", "Module Guard approved-build registration: transaction commit completed."
    mModuleGuardNestedVerificationTransactionActive = False
    mModuleGuardRegistrationTransactionOwned = False
    registrationResult.RegistrationDetails = "Batch: " & MODULE_GUARD_APPROVED_CURRENT_BUILD_BATCH_ID & vbCrLf & _
        "Modules: " & CStr(componentData.count) & vbCrLf & _
        "Guard enablement requested: " & CStr(enableAfterRegistration) & vbCrLf & _
        "Guard enabled after registration: " & CStr(effectiveEnableAfterRegistration) & vbCrLf & _
        "Guard enablement deferred for incomplete-setup bootstrap: " & CStr(enablementDeferredForBootstrap) & _
        IIf(Len(enablementResolutionDetails) = 0, vbNullString, vbCrLf & enablementResolutionDetails) & vbCrLf & _
        "Approval basis: " & IIf(explicitlyApproveInstalledBuild, "explicit user approval of the structurally validated exact installed module set, overriding active-manifest lineage/timestamp transition rejection", "exact managed-header coordinated-build identity with strict active-manifest lineage/timestamp validation") & "; module identity, component type, current and previous Version-ID validity, timestamp validity, required presence, Version-ID uniqueness, and exact-version history immutability." & _
        vbCrLf & vbCrLf & verificationDetails
    If effectiveEnableAfterRegistration Then
        registrationOutcomeMessage = "Approved current module build registered, verified, and enabled."
    ElseIf enablementDeferredForBootstrap Then
        registrationOutcomeMessage = "Approved current module build registered and verified; guard enablement was deferred until Settings repair supplies the canonical default-False setting row."
    Else
        registrationOutcomeMessage = "Approved current module build registered and verified; the existing Module Guard enabled/disabled state was preserved because enablement was not requested."
    End If
    LogModuleGuardMessage registrationOutcomeMessage, registrationResult.RegistrationDetails, TRACE_LEVEL_INFO, "TryRegisterApprovedBuild", "Register approved current build"
    TryRegisterApprovedBuild = True
    GoTo CleanExit
ErrHandler:
    registrationResult.failedErrNumber = Err.Number
    registrationResult.failedErrSource = Err.Source
    registrationResult.failedErrDescription = Err.Description
    If Len(registrationStage) > 0 Then registrationResult.failedErrDescription = "Registration stage: " & registrationStage & vbCrLf & registrationResult.failedErrDescription
    If transactionStarted Then
        rollbackVerified = RollbackModuleGuardTransaction(ws, cleanupDetails)
        If rollbackVerified Then
            transactionStarted = False
            mModuleGuardRegistrationTransactionOwned = False
        End If
    End If
    If transactionStarted Then AppendGuardDetail cleanupDetails, "The registration-owned transaction could not be rolled back successfully; the database must not continue through setup."
    If rollbackVerified Then AppendGuardDetail cleanupDetails, "The staged approved-build history, manifest, verification snapshot, verification-status updates, and any requested guard-enable change were rolled back together."
    If Len(cleanupDetails) > 0 Then registrationResult.failedErrDescription = registrationResult.failedErrDescription & vbCrLf & vbCrLf & cleanupDetails
    If registrationResult.failedErrNumber = 0 Then registrationResult.failedErrNumber = MODULE_GUARD_ERROR_BASE + 68
    If Len(Trim$(registrationResult.failedErrSource)) = 0 Then registrationResult.failedErrSource = "TryRegisterApprovedBuild"
    If Len(Trim$(registrationResult.failedErrDescription)) = 0 Then registrationResult.failedErrDescription = "The approved current module build was not registered."
    If IsModuleGuardVerificationFailureTranscript(verificationDetails) Then
        registrationResult.RegistrationDetails = verificationDetails
        Debug.Print verificationDetails
        PersistModuleGuardVerificationFailureTranscript verificationDetails, registrationResult.failedErrNumber, registrationResult.failedErrSource
    End If
    LogModuleGuardMessage "Approved current module build registration failed.", "Err.Number: " & CStr(registrationResult.failedErrNumber) & vbCrLf & "Err.Source: " & registrationResult.failedErrSource & vbCrLf & "Err.Description: " & registrationResult.failedErrDescription, TRACE_LEVEL_ERROR, "TryRegisterApprovedBuild", "Register approved current build"
    LogModuleGuardPreflightBoundary "TryRegisterApprovedBuild", "Module Guard approved-build registration terminated with failure.", "Stage: " & registrationStage & vbCrLf & "Err.Number: " & CStr(registrationResult.failedErrNumber) & vbCrLf & "Err.Source: " & registrationResult.failedErrSource & vbCrLf & "Err.Description: " & registrationResult.failedErrDescription, TRACE_LEVEL_ERROR
CleanExit:
    On Error Resume Next
    Set mModuleGuardInstalledInventoryEvidence = Nothing
    If namingSnapshotStarted Then
        LogModuleGuardPreflightBoundary "TryRegisterApprovedBuild", "Module Guard approved-build registration: naming snapshot cleanup starting."
        InvalidateNamingSnapshot
        namingSnapshotStarted = False
        LogModuleGuardPreflightBoundary "TryRegisterApprovedBuild", "Module Guard approved-build registration: naming snapshot cleanup completed."
    End If
    On Error GoTo 0
End Function

' Purpose: Performs the public register approved build operation using the routine's existing ADC workflow and side effects.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine shows its user-facing completion, status, warning, or error message in addition to performing the underlying work. Default: True.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub RegisterApprovedBuild(Optional ByVal showMessage As Boolean = True)
    Dim registrationResult As TModuleGuardRegistrationResult
    If TryRegisterApprovedBuild(registrationResult) Then
        If showMessage Then UtilityMsgBox "The approved " & DEFAULT_UTILITY_LONG_NAME & " module build was registered and verified successfully." & vbCrLf & vbCrLf & registrationResult.installedVersionSummary, vbInformation, DEFAULT_UTILITY_LONG_NAME & " Module Update Guard"
        Exit Sub
    End If
    If showMessage Then
        CoreShowErrorMessage registrationResult.failedErrDescription, vbCritical, DEFAULT_UTILITY_LONG_NAME & " Module Update Guard", "RegisterApprovedBuild", "Core"
        Exit Sub
    End If
    Err.Raise registrationResult.failedErrNumber, "RegisterApprovedBuild", registrationResult.failedErrDescription
End Sub

Private Sub DeactivateNonOwnedManifestRows(ByVal db As DAO.Database)
    'Deactivates manifest rows outside the current ADC-owned roster while preserving append-only history.
    Dim rs As DAO.Recordset
    Dim deactivatedCount As Long
    Set rs = db.OpenRecordset("SELECT " & SQL_FIELD_MODULE_NAME & ",[" & FIELD_ACTIVE & "]," & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_IS_REQUIRED & "," & SQL_FIELD_LAST_VERIFICATION_RESULT & "," & SQL_FIELD_LAST_VERIFICATION_NOTES & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionManifest") & "] WHERE [" & FIELD_ACTIVE & "]=True;", dbOpenDynaset)
    Do While rs.EOF = False
        PumpModuleGuardSetupResponsiveness
        If ModuleGuardOwnsComponent(CStr(Nz(rs.fields(FIELD_MODULE_NAME), vbNullString))) = False Then
            rs.Edit
            rs.fields(FIELD_ACTIVE).value = False
            rs!isRequired = False
            rs.fields(FIELD_LAST_VERIFICATION_RESULT) = Null
            rs.fields(FIELD_LAST_VERIFICATION_NOTES) = "This component is outside the current " & DEFAULT_UTILITY_SHORT_NAME & " Module Guard roster. Its active manifest expectation was retired without deleting append-only version history."
            rs.Update
            deactivatedCount = deactivatedCount + 1
        End If
        rs.MoveNext
    Loop
    rs.Close
    If deactivatedCount > 0 Then
        LogModuleGuardMessage "Manifest rows outside the current " & DEFAULT_UTILITY_SHORT_NAME & " roster were deactivated without deleting append-only history.", _
            "Rows deactivated: " & CStr(deactivatedCount) & vbCrLf & _
            "Owned set: " & ModuleGuardOwnedComponentDescription(), _
            TRACE_LEVEL_INFO, "DeactivateNonOwnedManifestRows", "Enforce " & DEFAULT_UTILITY_SHORT_NAME & " Module Guard ownership boundary"
    End If
End Sub

Private Function ExplicitlyApprovedInstalledBuildDefinitions(ByRef componentData As Collection, ByVal baselineDefinitions As Collection, ByVal installedInventoryEvidence As Collection) As Collection
    Dim expectedDefinitions As Collection
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: rows records the structurally validated installed build in approved-definition format; componentData is the corresponding registration payload. Both derive from
    ' observed installed headers and do not replace the fixed guarded-module roster or persistent manifest/history authority.
    Dim rows As New Collection
    Dim componentsByName As Object
    Dim timestampTextByName As Object
    Dim typeByName As Object
    Dim versionByName As Object
    Dim previousByName As Object
    Dim timestampByName As Object
    Dim headerNameByName As Object
    Dim versionCounts As Object
    Dim expectedNames As Object
    Dim Definition As Variant
    Dim key As Variant
    Dim ModuleName As String
    Dim expectedType As String
    Dim actualType As String
    Dim actualVersion As String
    Dim actualPrevious As String
    Dim actualTimestamp As Variant
    Dim actualTimestampText As String
    Dim actualHeaderName As String
    Dim extraModules As String
    Dim installedProjectComponentCount As Long
    Set expectedDefinitions = baselineDefinitions
    If expectedDefinitions Is Nothing Then Err.Raise MODULE_GUARD_ERROR_BASE + 144, "ExplicitlyApprovedInstalledBuildDefinitions", "The explicit registration operation did not supply the canonical initial Module Guard baseline definitions."
    ' INTENTIONAL KEYED LOOKUP: componentsByName, timestampTextByName, typeByName, versionByName, previousByName, timestampByName, headerNameByName, versionCounts, expectedNames provides keyed membership/lookup over
    ' source rows or state already read by the authoritative installed-module inventory evidence; it is a derived access path, not a second authority.
    Set expectedNames = CreateObject("Scripting.Dictionary")
    ConfigureTextDictionary expectedNames
    LoadInstalledModuleInventoryEvidence installedInventoryEvidence, componentsByName, timestampTextByName, typeByName, versionByName, previousByName, timestampByName, headerNameByName, versionCounts, installedProjectComponentCount
    For Each Definition In expectedDefinitions
        PumpModuleGuardSetupResponsiveness
        expectedNames(CStr(Definition(0))) = True
    Next Definition
    For Each key In componentsByName.keys
        PumpModuleGuardSetupResponsiveness
        If expectedNames.Exists(CStr(key)) = False Then AppendGuardDetail extraModules, CStr(key)
    Next key
    If Len(extraModules) > 0 Then Err.Raise MODULE_GUARD_ERROR_BASE + 135, "ExplicitlyApprovedInstalledBuildDefinitions", "The explicitly approved installed build contains unexpected ADC-owned modules." & vbCrLf & vbCrLf & extraModules
    If componentsByName.count <> expectedDefinitions.count Then Err.Raise MODULE_GUARD_ERROR_BASE + 136, "ExplicitlyApprovedInstalledBuildDefinitions", "The explicitly approved installed build is incomplete." & vbCrLf & "Expected ADC-owned modules: " & CStr(expectedDefinitions.count) & vbCrLf & "Installed ADC-owned modules: " & CStr(componentsByName.count)
    ' INTENTIONAL OPERATION SNAPSHOT: componentData captures one coherent source state for this invocation of ExplicitlyApprovedInstalledBuildDefinitions; validity ends with the operation rather than becoming a live
    ' data authority.
    Set componentData = New Collection
    For Each Definition In expectedDefinitions
        PumpModuleGuardSetupResponsiveness
        ModuleName = CStr(Definition(0))
        expectedType = CStr(Definition(1))
        If componentsByName.Exists(ModuleName) = False Then Err.Raise MODULE_GUARD_ERROR_BASE + 137, "ExplicitlyApprovedInstalledBuildDefinitions", "The explicitly approved installed build is missing a required ADC module." & vbCrLf & "Module: " & ModuleName
        actualType = CStr(typeByName(ModuleName))
        actualVersion = CStr(versionByName(ModuleName))
        actualPrevious = CStr(previousByName(ModuleName))
        actualTimestamp = timestampByName(ModuleName)
        actualTimestampText = CStr(timestampTextByName(ModuleName))
        actualHeaderName = CStr(headerNameByName(ModuleName))
        If StrComp(actualType, expectedType, vbTextCompare) <> 0 Then Err.Raise MODULE_GUARD_ERROR_BASE + 138, "ExplicitlyApprovedInstalledBuildDefinitions", "The explicitly approved installed build has the wrong VBA component type." & vbCrLf & "Module: " & ModuleName & vbCrLf & "Expected: " & expectedType & vbCrLf & "Installed: " & actualType
        If StrComp(actualHeaderName, ModuleName, vbBinaryCompare) <> 0 Then Err.Raise MODULE_GUARD_ERROR_BASE + 139, "ExplicitlyApprovedInstalledBuildDefinitions", "The explicitly approved installed build has a managed-header identity mismatch." & vbCrLf & "Module: " & ModuleName & vbCrLf & "Header VB_Name: " & DisplayGuardValue(actualHeaderName)
        If IsValidChatGPTVersionID(actualVersion) = False Then Err.Raise MODULE_GUARD_ERROR_BASE + 140, "ExplicitlyApprovedInstalledBuildDefinitions", "The explicitly approved installed build has an invalid current Version ID." & vbCrLf & "Module: " & ModuleName & vbCrLf & "Version ID: " & DisplayGuardValue(actualVersion)
        If StrComp(actualPrevious, MODULE_GUARD_TRACKING_PREFIX & MODULE_GUARD_NO_VERSION_SUFFIX, vbBinaryCompare) <> 0 And IsValidChatGPTVersionID(actualPrevious) = False Then Err.Raise MODULE_GUARD_ERROR_BASE + 141, "ExplicitlyApprovedInstalledBuildDefinitions", "The explicitly approved installed build has an invalid Previous Version ID." & vbCrLf & "Module: " & ModuleName & vbCrLf & "Previous Version ID: " & DisplayGuardValue(actualPrevious)
        If IsNull(actualTimestamp) Or Len(actualTimestampText) = 0 Then Err.Raise MODULE_GUARD_ERROR_BASE + 142, "ExplicitlyApprovedInstalledBuildDefinitions", "The explicitly approved installed build has a missing or invalid Version Timestamp." & vbCrLf & "Module: " & ModuleName & vbCrLf & "Timestamp: " & DisplayGuardValue(actualTimestampText)
        If CLng(versionCounts(actualVersion)) <> 1 Then Err.Raise MODULE_GUARD_ERROR_BASE + 143, "ExplicitlyApprovedInstalledBuildDefinitions", "The explicitly approved installed build repeats a Version ID across ADC modules." & vbCrLf & "Module: " & ModuleName & vbCrLf & "Version ID: " & actualVersion
        AddBaselineDefinition rows, ModuleName, actualType, actualVersion, actualPrevious, actualTimestampText
        componentData.Add Array(ModuleName, actualType, actualVersion, actualPrevious, actualTimestamp, vbNullString)
    Next Definition
    ValidateApprovedBuildDefinitions rows
    Set ExplicitlyApprovedInstalledBuildDefinitions = rows
End Function

Private Sub AddBaselineDefinition(ByRef rows As Collection, ByVal ModuleName As String, ByVal ComponentType As String, ByVal versionID As String, ByVal previousVersionID As String, ByVal timestampText As String)
    rows.Add Array(ModuleName, ComponentType, versionID, previousVersionID, timestampText)
End Sub

' =============================================================================================================================================================================
' SCHEMA / INDEX / LOOKUP REPAIR PROCEDURES
' =============================================================================================================================================================================

Private Sub EnsureModuleUpdateHistoryTable(ByVal db As DAO.Database)
    Dim fieldDefinitions As Variant
    fieldDefinitions = ModuleUpdateHistoryFieldDefinitions()
    If ModuleGuardTableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleUpdateHistory")) = False Then
        CreateModuleGuardTableFromDefinitions db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleUpdateHistory"), fieldDefinitions
    End If
    Set db = Nothing
    Set db = currentDb
    ValidateOrAddTableFields db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleUpdateHistory"), fieldDefinitions
    Set db = Nothing
    Set db = currentDb
    EnsurePrimaryIndexExact db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleUpdateHistory"), FIELD_ID
    Set db = Nothing
    Set db = currentDb
    EnsureIndex db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleUpdateHistory"), INDEX_UX_MODULE_VERSION, FIELD_MODULE_NAME & "|" & FIELD_CHAT_GPT_VERSION_ID, True, False
    Set db = Nothing
    Set db = currentDb
    EnsureIndex db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleUpdateHistory"), INDEX_IX_UPDATE_BATCH_ID, FIELD_UPDATE_BATCH_ID, False, False
    Set db = Nothing
    Set db = currentDb
    EnsureIndex db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleUpdateHistory"), INDEX_IX_RECORDED_ON, FIELD_RECORDED_ON, False, False
    Set db = Nothing
    Set db = currentDb
    applyFieldDescriptions db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleUpdateHistory"), ModuleGuardFieldDescriptions(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleUpdateHistory"))
    Set db = Nothing
    Set db = currentDb
    ConfigureModuleGuardTableLookup db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleUpdateHistory"), FIELD_COMPONENT_TYPE
    Set db = Nothing
    Set db = currentDb
    ConfigureModuleGuardTableLookup db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleUpdateHistory"), FIELD_RECORDED_BY
    Set db = Nothing
    Set db = currentDb
    ClearUnexpectedModuleGuardLookupFields db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleUpdateHistory"), fieldDefinitions, FIELD_COMPONENT_TYPE & "|" & FIELD_RECORDED_BY
    Set db = Nothing
    Set db = currentDb
    RequireModuleGuardTable db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleUpdateHistory")
End Sub

Private Sub EnsureModuleVersionManifestTable(ByVal db As DAO.Database)
    Dim fieldDefinitions As Variant
    fieldDefinitions = ModuleVersionManifestFieldDefinitions()
    If ModuleGuardTableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionManifest")) = False Then
        CreateModuleGuardTableFromDefinitions db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionManifest"), fieldDefinitions
    End If
    Set db = Nothing
    Set db = currentDb
    ValidateOrAddTableFields db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionManifest"), fieldDefinitions
    Set db = Nothing
    Set db = currentDb
    EnsurePrimaryIndexExact db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionManifest"), FIELD_MODULE_NAME
    Set db = Nothing
    Set db = currentDb
    EnsureIndex db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionManifest"), INDEX_IX_MANIFEST_ACTIVE, FIELD_ACTIVE & "|" & FIELD_MANAGED_OBJECT_DEPENDENCY_IS_REQUIRED, False, False
    Set db = Nothing
    Set db = currentDb
    applyFieldDescriptions db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionManifest"), ModuleGuardFieldDescriptions(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionManifest"))
    Set db = Nothing
    Set db = currentDb
    ConfigureModuleGuardTableLookup db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionManifest"), FIELD_COMPONENT_TYPE
    Set db = Nothing
    Set db = currentDb
    ConfigureModuleGuardTableLookup db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionManifest"), FIELD_LAST_VERIFICATION_RESULT
    Set db = Nothing
    Set db = currentDb
    ClearUnexpectedModuleGuardLookupFields db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionManifest"), fieldDefinitions, FIELD_COMPONENT_TYPE & "|" & FIELD_LAST_VERIFICATION_RESULT
    Set db = Nothing
    Set db = currentDb
    RequireModuleGuardTable db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionManifest")
End Sub

Private Sub EnsureModuleVersionVerificationTable(ByVal db As DAO.Database)
    Dim fieldDefinitions As Variant
    fieldDefinitions = ModuleVersionVerificationFieldDefinitions()
    If ModuleGuardTableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionVerification")) = False Then
        CreateModuleGuardTableFromDefinitions db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionVerification"), fieldDefinitions
    End If
    Set db = Nothing
    Set db = currentDb
    ValidateOrAddTableFields db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionVerification"), fieldDefinitions
    Set db = Nothing
    Set db = currentDb
    EnsurePrimaryIndexExact db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionVerification"), FIELD_ID
    Set db = Nothing
    Set db = currentDb
    EnsureIndex db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionVerification"), INDEX_IX_VERIFICATION_RUN_ID, FIELD_VERIFICATION_RUN_ID, False, False
    Set db = Nothing
    Set db = currentDb
    EnsureIndex db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionVerification"), INDEX_IX_VERIFICATION_RESULT, FIELD_VERIFICATION_RESULT, False, False
    Set db = Nothing
    Set db = currentDb
    RequireNoDuplicateModuleVerificationIdentity db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionVerification")
    EnsureIndex db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionVerification"), INDEX_IX_MODULE_VERIFICATION_RUN, FIELD_VERIFICATION_RUN_ID & "|" & FIELD_MODULE_NAME, True, False
    Set db = Nothing
    Set db = currentDb
    applyFieldDescriptions db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionVerification"), ModuleGuardFieldDescriptions(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionVerification"))
    Set db = Nothing
    Set db = currentDb
    ConfigureModuleGuardTableLookup db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionVerification"), FIELD_COMPONENT_TYPE
    Set db = Nothing
    Set db = currentDb
    ConfigureModuleGuardTableLookup db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionVerification"), FIELD_VERIFICATION_RESULT
    Set db = Nothing
    Set db = currentDb
    ClearUnexpectedModuleGuardLookupFields db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionVerification"), fieldDefinitions, FIELD_COMPONENT_TYPE & "|" & FIELD_VERIFICATION_RESULT
    Set db = Nothing
    Set db = currentDb
    RequireModuleGuardTable db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleVersionVerification")
End Sub

Private Sub RequireNoDuplicateModuleVerificationIdentity(ByVal db As DAO.Database, ByVal tableName As String)
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT TOP 1 [" & FIELD_VERIFICATION_RUN_ID & "],[" & FIELD_MODULE_NAME & "],Count(*) AS " & FIELD_DUP_COUNT & " FROM [" & tableName & "] GROUP BY [" & FIELD_VERIFICATION_RUN_ID & "],[" & FIELD_MODULE_NAME & "] HAVING Count(*)>1 ORDER BY [" & FIELD_VERIFICATION_RUN_ID & "],[" & FIELD_MODULE_NAME & "];", dbOpenSnapshot)
    If rs.EOF = False Then
        Err.Raise MODULE_GUARD_ERROR_BASE + 59, "RequireNoDuplicateModuleVerificationIdentity", "Module Guard verification-history uniqueness repair was blocked because duplicate persisted module/run results already exist. Existing verification evidence was not changed." & vbCrLf & "VerificationRunID: " & CStr(rs.fields(FIELD_VERIFICATION_RUN_ID).value) & vbCrLf & "ModuleName: " & CStr(rs.fields(FIELD_MODULE_NAME).value) & vbCrLf & "Rows: " & CStr(CLng(rs.fields(FIELD_DUP_COUNT)))
    End If
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
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
    Err.Raise savedErrNumber, "RequireNoDuplicateModuleVerificationIdentity", "Could not verify Module Guard verification-history row identity before unique-index repair." & vbCrLf & "Table: " & tableName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub ValidateOrAddTableFields(ByVal db As DAO.Database, ByVal tableName As String, ByVal definitions As Variant)
    Dim Definition As Variant
    Dim parts As Variant
    Dim tdf As DAO.tableDef
    Dim fld As DAO.Field
    Dim fieldName As String
    Dim expectedType As Long
    Dim expectedSize As Long
    Dim expectedAttributes As Long
    Dim requiredField As Boolean
    RequireModuleGuardTable db, tableName
    For Each Definition In definitions
        PumpModuleGuardSetupResponsiveness
        parts = Split(CStr(Definition), "|")
        fieldName = CStr(parts(0))
        expectedType = CLng(parts(1))
        expectedSize = CLng(parts(2))
        expectedAttributes = CLng(parts(3))
        requiredField = (UBound(parts) >= 4 And CLng(parts(4)) <> 0)
        Set tdf = FreshModuleGuardTableDef(db, tableName)
        Set fld = FindModuleGuardField(tdf, fieldName)
        If fld Is Nothing Then
            If fieldName = FIELD_ID Then Err.Raise MODULE_GUARD_ERROR_BASE + 30, "ValidateOrAddTableFields", "Required ID field is missing from an existing module-guard table." & vbCrLf & "Table: " & tableName
            Set tdf = Nothing
            db.Execute "ALTER TABLE [" & tableName & "] ADD COLUMN [" & fieldName & "] " & DdlTypeText(expectedType, expectedSize), dbFailOnError
            Set db = Nothing
            Set db = currentDb
            Set tdf = FreshModuleGuardTableDef(db, tableName)
            Set fld = FindModuleGuardField(tdf, fieldName)
            If fld Is Nothing Then Err.Raise MODULE_GUARD_ERROR_BASE + 44, "ValidateOrAddTableFields", "A missing module-guard field was not present after ALTER TABLE." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName
        End If
        If fld.Type <> expectedType Then Err.Raise MODULE_GUARD_ERROR_BASE + 31, "ValidateOrAddTableFields", "Field type conflict in module-guard table." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName & vbCrLf & "Expected type: " & CStr(expectedType) & vbCrLf & "Actual type: " & CStr(fld.Type)
        If expectedType = dbText And expectedSize > 0 And fld.Size <> expectedSize Then Err.Raise MODULE_GUARD_ERROR_BASE + 32, "ValidateOrAddTableFields", "Field size conflict in module-guard table." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName & vbCrLf & "Expected size: " & CStr(expectedSize) & vbCrLf & "Actual size: " & CStr(fld.Size)
        If expectedAttributes <> 0 And (fld.Attributes And expectedAttributes) <> expectedAttributes Then Err.Raise MODULE_GUARD_ERROR_BASE + 33, "ValidateOrAddTableFields", "Field attribute conflict in module-guard table." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName
        If fld.Required <> requiredField Then fld.Required = requiredField
        If requiredField And expectedType = dbText Then If fld.allowZeroLength Then fld.allowZeroLength = False
        If fld.Required <> requiredField Then Err.Raise MODULE_GUARD_ERROR_BASE + 45, "ValidateOrAddTableFields", "Field Required property did not match the canonical module-guard definition after repair." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName
        If requiredField And expectedType = dbText Then If fld.allowZeroLength Then Err.Raise MODULE_GUARD_ERROR_BASE + 46, "ValidateOrAddTableFields", "Required text identity/status field still allows zero-length strings after module-guard repair." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName
    Next Definition
    ApplyLogicalFieldOrder db, tableName, definitions
End Sub

Private Function ModuleGuardTableExists(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
    db.TableDefs.Refresh
    ModuleGuardTableExists = DaoTableExists(db, tableName)
End Function

Private Function ModuleUpdateHistoryFieldDefinitions() As Variant
    ModuleUpdateHistoryFieldDefinitions = Array(FIELD_ID & "|4|0|16|0", FIELD_UPDATE_BATCH_ID & "|10|100|0|1", FIELD_MODULE_NAME & "|10|128|0|1", FIELD_COMPONENT_TYPE & "|10|50|0|1", FIELD_CHAT_GPT_VERSION_ID & "|10|40|0|1", FIELD_PREVIOUS_CHAT_GPT_VERSION_ID & "|10|40|0|0", FIELD_VERSION_TIMESTAMP & "|8|0|0|0", FIELD_CHANGE_DESCRIPTION & "|12|0|0|0", FIELD_RECORDED_ON & "|8|0|0|1", FIELD_RECORDED_BY & "|10|50|0|0", FIELD_IS_BASELINE & "|1|0|0|0", FIELD_SUPERSEDED_ON & "|8|0|0|0", FIELD_MANAGED_OBJECT_DEPENDENCY_NOTES & "|12|0|0|0")
End Function

Private Function ModuleVersionManifestFieldDefinitions() As Variant
    ModuleVersionManifestFieldDefinitions = Array(FIELD_MODULE_NAME & "|10|128|0|1", FIELD_COMPONENT_TYPE & "|10|50|0|1", FIELD_EXPECTED_VERSION_ID & "|10|40|0|1", FIELD_EXPECTED_PREVIOUS_VERSION_ID & "|10|40|0|0", FIELD_EXPECTED_VERSION_TIMESTAMP & "|8|0|0|0", FIELD_UPDATE_BATCH_ID & "|10|100|0|0", FIELD_EXPECTED_CHANGE_DESCRIPTION & "|12|0|0|0", FIELD_MANAGED_OBJECT_DEPENDENCY_IS_REQUIRED & "|1|0|0|0", FIELD_ACTIVE & "|1|0|0|0", FIELD_BASELINE_ACCEPTED & "|1|0|0|0", FIELD_MANIFEST_UPDATED_ON & "|8|0|0|0", FIELD_LAST_VERIFIED_ON & "|8|0|0|0", FIELD_LAST_VERIFICATION_RESULT & "|10|50|0|0", FIELD_LAST_VERIFICATION_NOTES & "|12|0|0|0")
End Function

Private Function ModuleVersionVerificationFieldDefinitions() As Variant
    ModuleVersionVerificationFieldDefinitions = Array(FIELD_ID & "|4|0|16|0", FIELD_VERIFICATION_RUN_ID & "|10|100|0|1", FIELD_VERIFIED_ON & "|8|0|0|1", FIELD_MODULE_NAME & "|10|128|0|1", FIELD_COMPONENT_TYPE & "|10|50|0|0", FIELD_MANAGED_OBJECT_DEPENDENCY_IS_REQUIRED & "|1|0|0|0", FIELD_IS_MANIFESTED & "|1|0|0|0", FIELD_EXPECTED_VERSION_ID & "|10|40|0|0", FIELD_INSTALLED_VERSION_ID & "|10|40|0|0", FIELD_EXPECTED_PREVIOUS_VERSION_ID & "|10|40|0|0", FIELD_INSTALLED_PREVIOUS_VERSION_ID & "|10|40|0|0", FIELD_EXPECTED_VERSION_TIMESTAMP & "|8|0|0|0", FIELD_INSTALLED_VERSION_TIMESTAMP & "|8|0|0|0", FIELD_VERIFICATION_RESULT & "|10|50|0|1", FIELD_VERIFICATION_NOTES & "|12|0|0|0", FIELD_UPDATE_BATCH_ID & "|10|100|0|0")
End Function

Private Function IndexFieldNamesText(ByVal idx As DAO.index) As String
    Dim fld As DAO.Field
    Dim resultText As String
    For Each fld In idx.fields
        PumpModuleGuardSetupResponsiveness
        If Len(resultText) > 0 Then resultText = resultText & "|"
        resultText = resultText & fld.Name
    Next fld
    IndexFieldNamesText = resultText
End Function

' Purpose: Resolves module update history field order text from the supplied/current ADC state, using `ModuleGuardFieldOrderText` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function ModuleUpdateHistoryFieldOrderText() As String
    ' Wrapper purpose: Module Guard-owned facade returning canonical physical field order directly from the authoritative update-history schema.
    ' Why retained: Retained so external schema consumers do not access Module Guard's private schema/formatting implementation or maintain an independent field-order mirror.
    'Intended use: Returns the current canonical physical field order owned by Module Guard for its update-history table.
    ModuleUpdateHistoryFieldOrderText = ModuleGuardFieldOrderText(ModuleUpdateHistoryFieldDefinitions())
End Function

' Purpose: Resolves module version manifest field order text from the supplied/current ADC state, using `ModuleGuardFieldOrderText` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function ModuleVersionManifestFieldOrderText() As String
    ' Wrapper purpose: Module Guard-owned facade returning canonical physical field order directly from the authoritative version-manifest schema.
    ' Why retained: Retained so external schema consumers do not access Module Guard's private schema/formatting implementation or maintain an independent field-order mirror.
    'Intended use: Returns the current canonical physical field order owned by Module Guard for its version-manifest table.
    ModuleVersionManifestFieldOrderText = ModuleGuardFieldOrderText(ModuleVersionManifestFieldDefinitions())
End Function

' Purpose: Resolves module version verification field order text from the supplied/current ADC state, using `ModuleGuardFieldOrderText` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function ModuleVersionVerificationFieldOrderText() As String
    ' Wrapper purpose: Module Guard-owned facade returning canonical physical field order directly from the authoritative verification schema.
    ' Why retained: Retained so external schema consumers do not access Module Guard's private schema/formatting implementation or maintain an independent field-order mirror.
    'Intended use: Returns the current canonical physical field order owned by Module Guard for its verification table.
    ModuleVersionVerificationFieldOrderText = ModuleGuardFieldOrderText(ModuleVersionVerificationFieldDefinitions())
End Function

Private Sub ApplyLogicalFieldOrder(ByVal db As DAO.Database, ByVal tableName As String, ByVal definitions As Variant)
    Dim tdf As DAO.tableDef
    Dim Definition As Variant
    Dim parts As Variant
    Dim fld As DAO.Field
    Dim ordinal As Long
    Set tdf = FreshModuleGuardTableDef(db, tableName)
    ordinal = 0
    For Each Definition In definitions
        PumpModuleGuardSetupResponsiveness
        parts = Split(CStr(Definition), "|")
        Set fld = FindModuleGuardField(tdf, CStr(parts(0)))
        If fld Is Nothing Then Err.Raise MODULE_GUARD_ERROR_BASE + 45, "ApplyLogicalFieldOrder", "A required field was missing while applying module-guard field order." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & CStr(parts(0))
        fld.ordinalPosition = ordinal
        EnsureDaoPropertyValue fld, "ColumnOrder", dbLong, ordinal + 1, "ApplyLogicalFieldOrder"
        ordinal = ordinal + 1
    Next Definition
    tdf.fields.Refresh
End Sub

Private Sub EnsureIndex(ByVal db As DAO.Database, ByVal tableName As String, ByVal indexName As String, ByVal fieldNamesText As String, ByVal uniqueIndex As Boolean, ByVal primaryIndex As Boolean)
    Dim tdf As DAO.tableDef
    Dim idx As DAO.index
    Dim fieldNames As Variant
    Dim expectedText As String
    Dim actualText As String
    Dim sqlText As String
    Dim fieldName As Variant
    Set tdf = FreshModuleGuardTableDef(db, tableName)
    Set idx = FindIndex(tdf, indexName)
    expectedText = LCase$(Replace(fieldNamesText, " ", vbNullString))
    If Not idx Is Nothing Then
        actualText = IndexFieldNamesText(idx)
        If LCase$(actualText) = expectedText And idx.Unique = uniqueIndex And idx.Primary = primaryIndex Then Exit Sub
        Set idx = Nothing
        Set tdf = Nothing
        db.Execute "DROP INDEX [" & indexName & "] ON [" & tableName & "]", dbFailOnError
        Set db = Nothing
        Set db = currentDb
        Set tdf = FreshModuleGuardTableDef(db, tableName)
        If Not FindIndex(tdf, indexName) Is Nothing Then Err.Raise MODULE_GUARD_ERROR_BASE + 46, "EnsureIndex", "A mismatched module-guard index remained after DROP INDEX." & vbCrLf & "Table: " & tableName & vbCrLf & "Index: " & indexName
    End If
    fieldNames = Split(fieldNamesText, "|")
    For Each fieldName In fieldNames
        PumpModuleGuardSetupResponsiveness
        If FindModuleGuardField(tdf, CStr(fieldName)) Is Nothing Then Err.Raise MODULE_GUARD_ERROR_BASE + 47, "EnsureIndex", "A required index field is missing." & vbCrLf & "Table: " & tableName & vbCrLf & "Index: " & indexName & vbCrLf & "Field: " & CStr(fieldName)
    Next fieldName
    sqlText = "CREATE "
    If uniqueIndex Then sqlText = sqlText & "UNIQUE "
    sqlText = sqlText & "INDEX [" & indexName & "] ON [" & tableName & "] ("
    For Each fieldName In fieldNames
        PumpModuleGuardSetupResponsiveness
        If Right$(sqlText, 1) <> "(" Then sqlText = sqlText & ", "
        sqlText = sqlText & "[" & CStr(fieldName) & "]"
    Next fieldName
    sqlText = sqlText & ")"
    If primaryIndex Then sqlText = sqlText & " WITH PRIMARY"
    Set tdf = Nothing
    db.Execute sqlText, dbFailOnError
    Set db = Nothing
    Set db = currentDb
    Set tdf = FreshModuleGuardTableDef(db, tableName)
    Set idx = FindIndex(tdf, indexName)
    If idx Is Nothing Then Err.Raise MODULE_GUARD_ERROR_BASE + 35, "EnsureIndex", "Index creation could not be verified." & vbCrLf & "Table: " & tableName & vbCrLf & "Index: " & indexName
    If LCase$(IndexFieldNamesText(idx)) <> expectedText Or idx.Unique <> uniqueIndex Or idx.Primary <> primaryIndex Then Err.Raise MODULE_GUARD_ERROR_BASE + 36, "EnsureIndex", "Persisted index definition does not match the requested module-guard definition." & vbCrLf & "Table: " & tableName & vbCrLf & "Index: " & indexName
End Sub

Private Function FindIndex(ByVal tdf As DAO.tableDef, ByVal indexName As String) As DAO.index
    Dim idx As DAO.index
    For Each idx In tdf.Indexes
        PumpModuleGuardSetupResponsiveness
        If StrComp(idx.Name, indexName, vbTextCompare) = 0 Then
            Set FindIndex = idx
            Exit Function
        End If
    Next idx
End Function

Private Sub applyFieldDescriptions(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldDescriptions As Variant)
    On Error GoTo ErrHandler
    Dim tdf As DAO.tableDef
    Dim catalogFieldsTdf As DAO.tableDef
    Dim rs As DAO.Recordset
    Dim item As Variant
    Dim separatorPosition As Long
    Dim fieldName As String
    Dim descriptionText As String
    Dim defaultDescriptionText As String
    Dim defaultDescriptionSource As Variant
    Dim failureSourceName As String
    Dim managedObjectsTableName As String
    Dim managedObjectFieldsTableName As String
    Dim sqlText As String
    Dim persistedShortByField As Object
    Dim catalogShortField As DAO.Field
    Dim fld As DAO.Field
    Set tdf = FreshModuleGuardTableDef(db, tableName)
    Set persistedShortByField = CreateObject("Scripting.Dictionary")
    persistedShortByField.CompareMode = vbTextCompare
    managedObjectsTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID)
    managedObjectFieldsTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFields")
    If ModuleGuardTableExists(db, managedObjectsTableName) And ModuleGuardTableExists(db, managedObjectFieldsTableName) Then
        Set catalogFieldsTdf = db.TableDefs(managedObjectFieldsTableName)
        Set catalogShortField = FindModuleGuardField(catalogFieldsTdf, FIELD_MANAGED_OBJECT_FIELD_DESCRIPTION_SHORT)
        If Not catalogShortField Is Nothing Then
            sqlText = "SELECT f.[" & FIELD_MANAGED_OBJECT_FIELD_NAME & "],f.[" & FIELD_MANAGED_OBJECT_FIELD_DESCRIPTION_SHORT & "] FROM [" & managedObjectFieldsTableName & "] AS f INNER JOIN [" & managedObjectsTableName & "] AS o ON f.[" & FIELD_MANAGED_OBJECT_ID & "]=o." & SQL_FIELD_ID & " WHERE Nz(f.[" & FIELD_MANAGED_OBJECT_FIELD_IS_ACTIVE & "],False)<>False AND Nz(o.[" & FIELD_MANAGED_OBJECT_STATUS & "]," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & " AND UCase(Trim(Nz(o." & SQL_FIELD_OBJECT_TYPE & ",''))) = 'TABLE' AND o." & SQL_FIELD_OBJECT_NAME & "=" & SharedDao_SqlQuote(tableName) & ";"
            Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
            Do While rs.EOF = False
                fieldName = Trim$(CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_NAME).value, vbNullString)))
                If Len(fieldName) > 0 And Len(Trim$(CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_DESCRIPTION_SHORT).value, vbNullString)))) > 0 Then
                    If persistedShortByField.Exists(fieldName) Then Err.Raise MODULE_GUARD_ERROR_BASE + 149, "ApplyFieldDescriptions", "Duplicate MAC_Fields.FieldDescriptionShort rows exist for a Module Guard field." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName
                    persistedShortByField.Add fieldName, CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_DESCRIPTION_SHORT).value, vbNullString))
                End If
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing
        End If
    End If
    For Each item In fieldDescriptions
        PumpModuleGuardSetupResponsiveness
        separatorPosition = InStr(1, CStr(item), "|", vbBinaryCompare)
        fieldName = Left$(CStr(item), separatorPosition - 1)
        descriptionText = Mid$(CStr(item), separatorPosition + 1)
        If persistedShortByField.Exists(fieldName) Then descriptionText = CStr(persistedShortByField(fieldName))
        Set fld = FindModuleGuardField(tdf, fieldName)
        If fld Is Nothing Then Err.Raise MODULE_GUARD_ERROR_BASE + 48, "ApplyFieldDescriptions", "A required field was missing while applying module-guard descriptions." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName
        If Len(descriptionText) > 255 Then
            defaultDescriptionSource = vbNullString
            defaultDescriptionText = ManagedObjectCatalogFieldDescriptionShort(tableName, fieldName, defaultDescriptionSource)
            If Len(defaultDescriptionText) > 0 And StrComp(defaultDescriptionText, descriptionText, vbBinaryCompare) = 0 And Len(Trim$(CStr(defaultDescriptionSource))) > 0 Then
                failureSourceName = CStr(defaultDescriptionSource)
            Else
                failureSourceName = "(not applicable; persisted/runtime value does not match the current DEFAULTS default)"
            End If
            Err.Raise MODULE_GUARD_ERROR_BASE + 150, "ApplyFieldDescriptions", "MAC_Fields.FieldDescriptionShort exceeds the Access DAO Field.Description 255-character limit for a Module Guard field." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName & vbCrLf & "Characters: " & CStr(Len(descriptionText)) & vbCrLf & "Default/value source: " & failureSourceName
        End If
        EnsureDaoPropertyValue fld, DAO_DESCRIPTION_PROPERTY_NAME, dbText, descriptionText, "ApplyFieldDescriptions"
    Next item
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
    Err.Raise savedErrNumber, "ApplyFieldDescriptions", "Could not apply authoritative Module Guard field descriptions." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub ConfigureModuleGuardTableLookup(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String)
    Dim fld As DAO.Field
    Dim lookupTableName As String
    Dim rowSource As String
    lookupTableName = ConvertedValueLookupTableForField(tableName, fieldName)
    If Len(lookupTableName) = 0 Then Err.Raise MODULE_GUARD_ERROR_BASE + 49, "ConfigureModuleGuardTableLookup", "No table-backed Module Guard lookup definition exists." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName
    lookupTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, lookupTableName)
    Set fld = db.TableDefs(tableName).fields(fieldName)
    rowSource = "SELECT " & SQL_FIELD_VALUE_TEXT & ", " & SQL_FIELD_DESCRIPTION & " FROM [" & Replace(lookupTableName, "]", "]]") & "] ORDER BY " & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER & ", " & SQL_FIELD_VALUE_TEXT & ";"
    ApplyLookupFieldProperties fld, "Table/Query", rowSource, True, 2, 1, "0;2", "ConfigureModuleGuardTableLookup", MODULE_GUARD_ERROR_BASE + 49, MODULE_GUARD_ERROR_BASE + 50
End Sub

Private Sub ClearUnexpectedModuleGuardLookupFields(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldDefinitions As Variant, ByVal expectedLookupFieldsText As String)
    'Module Guard owns the lookup designation for its canonical fields. Remove obsolete lookup metadata from canonical fields that are no longer designated as lookups.
    Dim expected As Object
    Dim expectedField As Variant
    Dim Definition As Variant
    Dim parts As Variant
    Dim tdf As DAO.tableDef
    Dim fieldName As String
    Set expected = CreateObject("Scripting.Dictionary")
    expected.CompareMode = vbTextCompare
    For Each expectedField In Split(expectedLookupFieldsText, "|")
        fieldName = Trim$(CStr(expectedField))
        If Len(fieldName) > 0 Then expected(fieldName) = True
    Next expectedField
    Set tdf = FreshModuleGuardTableDef(db, tableName)
    For Each Definition In fieldDefinitions
        PumpModuleGuardSetupResponsiveness
        parts = Split(CStr(Definition), "|")
        fieldName = CStr(parts(0))
        If expected.Exists(fieldName) = False Then ClearModuleGuardFieldLookupProperties tdf, fieldName
    Next Definition
End Sub

Private Sub ClearModuleGuardFieldLookupProperties(ByVal tdf As DAO.tableDef, ByVal fieldName As String)
    On Error GoTo ErrHandler
    Dim fld As DAO.Field
    Dim propertyNames As Variant
    Dim propertyName As Variant
    Dim prp As DAO.Property
    Set fld = FindModuleGuardField(tdf, fieldName)
    If fld Is Nothing Then Exit Sub
    propertyNames = Array("DisplayControl", "RowSourceType", "RowSource", "BoundColumn", "ColumnCount", "ColumnWidths", "ColumnHeads", "ListRows", "ListWidth", "LimitToList", "AllowValueListEdits", "ListItemsEditForm")
    For Each propertyName In propertyNames
        Set prp = Nothing
        On Error Resume Next
        Err.Clear
        Set prp = fld.properties(CStr(propertyName))
        If Err.Number = 0 And Not prp Is Nothing Then fld.properties.Delete CStr(propertyName)
        If Err.Number <> 0 And Err.Number <> 3265 And Err.Number <> 3270 Then GoTo PropertyFailed
        Err.Clear
        On Error GoTo ErrHandler
    Next propertyName
    fld.properties.Refresh
    For Each propertyName In propertyNames
        Set prp = Nothing
        On Error Resume Next
        Err.Clear
        Set prp = fld.properties(CStr(propertyName))
        If Err.Number = 0 And Not prp Is Nothing Then
            On Error GoTo ErrHandler
            Err.Raise MODULE_GUARD_ERROR_BASE + 75, "ClearModuleGuardFieldLookupProperties", "A stale Module Guard lookup property remained after cleanup." & vbCrLf & "Table: " & tdf.Name & vbCrLf & "Field: " & fieldName & vbCrLf & "Property: " & CStr(propertyName)
        End If
        If Err.Number <> 0 And Err.Number <> 3265 And Err.Number <> 3270 Then GoTo PropertyFailed
        Err.Clear
        On Error GoTo ErrHandler
    Next propertyName
    Exit Sub
PropertyFailed:
    Dim propertyErrNumber As Long
    Dim propertyErrSource As String
    Dim propertyErrDescription As String
    propertyErrNumber = Err.Number
    propertyErrSource = Err.Source
    propertyErrDescription = Err.Description
    On Error GoTo ErrHandler
    Err.Raise propertyErrNumber, propertyErrSource, propertyErrDescription
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ClearModuleGuardFieldLookupProperties", "Could not clear and verify obsolete Module Guard lookup metadata." & vbCrLf & "Table: " & tdf.Name & vbCrLf & "Field: " & fieldName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub EnsureModuleGuardSavedQuery(ByVal db As DAO.Database, ByVal queryName As String, ByVal sqlText As String)
    ' Module Guard owns its query lifecycle/open verification. COMMON performs best-effort no-op comparison only. Persisted proof verifies the saved query contract that ACE accepted (SELECT type, source-table
    ' membership, output-field schema, and openability); ACE-only SQL serialization changes are diagnostic rather than fatal.
    On Error GoTo ErrHandler
    Dim qdf As DAO.queryDef
    Dim rs As DAO.Recordset
    Dim fld As DAO.Field
    Dim stageText As String
    Dim comparisonDetails As String
    Dim expectedSourceTables As String
    Dim storedSourceTables As String
    Dim acceptedSqlText As String
    Dim storedSqlText As String
    Dim acceptedFieldContract As String
    Dim storedFieldContract As String
    Dim acceptedQueryType As Long
    Dim storedQueryType As Long
    Dim fieldCount As Long
    expectedSourceTables = ModuleGuardExpectedSourceTablesForSql(sqlText)
    stageText = "Inspect existing saved query."
    Set qdf = FindQueryDef(db, queryName)
    If qdf Is Nothing Then
        stageText = "Create missing saved query."
        Set qdf = db.CreateQueryDef(queryName, sqlText)
    ElseIf QuerySqlEquivalent(db, qdf.sql, sqlText) = False Then
        stageText = "Replace saved query SQL."
        qdf.sql = sqlText
    End If
    acceptedSqlText = qdf.sql
    acceptedQueryType = CLng(qdf.Type)
    If acceptedQueryType <> dbQSelect Then Err.Raise MODULE_GUARD_ERROR_BASE + 38, "EnsureModuleGuardSavedQuery", "Canonical Module Guard SQL did not compile as a SELECT query: " & queryName
    acceptedFieldContract = ModuleGuardQueryFieldContract(qdf)
    If Len(acceptedFieldContract) = 0 Then Err.Raise MODULE_GUARD_ERROR_BASE + 39, "EnsureModuleGuardSavedQuery", "Canonical Module Guard query exposed no fields before persistence verification: " & queryName
    Set qdf = Nothing
    Set db = Nothing
    Set db = currentDb
    db.QueryDefs.Refresh
    stageText = "Verify persisted saved query contract."
    Set qdf = FindQueryDef(db, queryName)
    If qdf Is Nothing Then Err.Raise MODULE_GUARD_ERROR_BASE + 37, "EnsureModuleGuardSavedQuery", "Saved query creation could not be verified: " & queryName
    storedSqlText = qdf.sql
    storedQueryType = CLng(qdf.Type)
    If storedQueryType <> acceptedQueryType Or storedQueryType <> dbQSelect Then Err.Raise MODULE_GUARD_ERROR_BASE + 38, "EnsureModuleGuardSavedQuery", "Saved query type did not preserve the SELECT-query contract: " & queryName & vbCrLf & "ACE-accepted type: " & CStr(acceptedQueryType) & vbCrLf & "Persisted type: " & CStr(storedQueryType)
    storedSourceTables = ModuleGuardExpectedSourceTablesForSql(storedSqlText)
    If StrComp(storedSourceTables, expectedSourceTables, vbTextCompare) <> 0 Then Err.Raise MODULE_GUARD_ERROR_BASE + 38, "EnsureModuleGuardSavedQuery", "Saved query source-table membership did not preserve the Module Guard query contract: " & queryName & vbCrLf & "Expected source tables: " & IIf(Len(expectedSourceTables) = 0, "(none detected)", expectedSourceTables) & vbCrLf & "Persisted source tables: " & IIf(Len(storedSourceTables) = 0, "(none detected)", storedSourceTables)
    storedFieldContract = ModuleGuardQueryFieldContract(qdf)
    If StrComp(storedFieldContract, acceptedFieldContract, vbBinaryCompare) <> 0 Then Err.Raise MODULE_GUARD_ERROR_BASE + 38, "EnsureModuleGuardSavedQuery", "Saved query output-field schema did not preserve the SQL contract accepted by ACE: " & queryName
    If StrComp(FormattingInsensitiveQuerySqlKey(storedSqlText), FormattingInsensitiveQuerySqlKey(acceptedSqlText), vbBinaryCompare) <> 0 Then comparisonDetails = SqlPersistenceComparisonDiagnostics(acceptedSqlText, storedSqlText)
    Set qdf = Nothing
    stageText = "Open saved query and enumerate fields."
    Set rs = db.OpenRecordset(queryName, dbOpenSnapshot)
    For Each fld In rs.fields
        PumpModuleGuardSetupResponsiveness
        fieldCount = fieldCount + 1
    Next fld
    If fieldCount = 0 Then Err.Raise MODULE_GUARD_ERROR_BASE + 39, "EnsureModuleGuardSavedQuery", "Saved query opened but exposed no fields: " & queryName
    rs.Close
    Set rs = Nothing
    If Len(comparisonDetails) > 0 Then LogModuleGuardMessage "ACE normalized persisted Module Guard query SQL; query-contract verification succeeded.", "Query: " & queryName & vbCrLf & comparisonDetails, TRACE_LEVEL_INFO, "EnsureModuleGuardSavedQuery", stageText
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
    If Len(storedSqlText) = 0 Then
        If Not qdf Is Nothing Then storedSqlText = qdf.sql
    End If
    On Error GoTo 0
    Err.Raise savedErrNumber, "EnsureModuleGuardSavedQuery", "Could not create, repair, or open a module update guard saved query." & vbCrLf & "Stage: " & stageText & vbCrLf & "Query: " & queryName & vbCrLf & "Expected source tables: " & IIf(Len(expectedSourceTables) = 0, "(none detected)", expectedSourceTables) & vbCrLf & "Requested SQL: " & sqlText & vbCrLf & "ACE-accepted SQL: " & IIf(Len(acceptedSqlText) = 0, "(unavailable)", acceptedSqlText) & vbCrLf & "Stored SQL: " & IIf(Len(storedSqlText) = 0, "(unavailable)", storedSqlText) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Original description: " & savedErrDescription
End Sub

Private Function ModuleGuardQueryFieldContract(ByVal qdf As DAO.queryDef) As String
    'Persistence-proof projection only: preserves the accepted output-field order, names, DAO types, and sizes without duplicating query SQL authority.
    Dim fld As DAO.Field
    Dim resultText As String
    qdf.fields.Refresh
    For Each fld In qdf.fields
        PumpModuleGuardSetupResponsiveness
        resultText = resultText & ModuleGuardStructurePart(fld.Name) & ModuleGuardStructurePart(CStr(fld.Type)) & ModuleGuardStructurePart(CStr(fld.Size))
    Next fld
    ModuleGuardQueryFieldContract = resultText
End Function

Private Function FindQueryDef(ByVal db As DAO.Database, ByVal queryName As String) As DAO.queryDef
    Dim qdf As DAO.queryDef
    For Each qdf In db.QueryDefs
        PumpModuleGuardSetupResponsiveness
        If StrComp(qdf.Name, queryName, vbTextCompare) = 0 Then Set FindQueryDef = qdf: Exit Function
    Next qdf
End Function

Private Function SqlPersistenceComparisonDiagnostics(ByVal acceptedSql As String, ByVal persistedSql As String) As String
    Dim acceptedKey As String
    Dim persistedKey As String
    Dim mismatchPosition As Long
    acceptedKey = FormattingInsensitiveQuerySqlKey(acceptedSql)
    persistedKey = FormattingInsensitiveQuerySqlKey(persistedSql)
    mismatchPosition = FirstTextDifferencePosition(acceptedKey, persistedKey)
    SqlPersistenceComparisonDiagnostics = "ACE-accepted raw SQL length: " & CStr(Len(acceptedSql)) & vbCrLf & "Persisted raw SQL length: " & CStr(Len(persistedSql)) & vbCrLf & "ACE-accepted comparison-key length: " & CStr(Len(acceptedKey)) & vbCrLf & "Persisted comparison-key length: " & CStr(Len(persistedKey)) & vbCrLf & "First comparison-key mismatch position: " & CStr(mismatchPosition) & vbCrLf & "ACE-accepted comparison-key excerpt: " & SqlDiagnosticExcerpt(acceptedKey, mismatchPosition) & vbCrLf & "Persisted comparison-key excerpt: " & SqlDiagnosticExcerpt(persistedKey, mismatchPosition)
End Function

Private Function SqlDiagnosticExcerpt(ByVal sqlText As String, ByVal mismatchPosition As Long) As String
    Dim startPosition As Long
    If mismatchPosition <= 0 Then mismatchPosition = 1
    startPosition = mismatchPosition - 80
    If startPosition < 1 Then startPosition = 1
    SqlDiagnosticExcerpt = Mid$(sqlText, startPosition, 180)
End Function

Private Function CaptureInstalledModuleInventoryEvidence(ByVal expectedDefinitions As Collection) As Collection
    'Authoritative transient evidence bundle for one unchanged VBProject state. The existing CaptureInstalledModuleInventory constructor performs the single physical scan; this bundle only carries its detached keyed
    'results across verification/registration consumers.
    Dim evidence As New Collection
    Dim componentsByName As Object
    Dim timestampTextByName As Object
    Dim typeByName As Object
    Dim versionByName As Object
    Dim previousByName As Object
    Dim timestampByName As Object
    Dim headerNameByName As Object
    Dim versionCounts As Object
    Dim installedCount As Long
    Set componentsByName = CreateObject("Scripting.Dictionary")
    Set timestampTextByName = CreateObject("Scripting.Dictionary")
    Set typeByName = CreateObject("Scripting.Dictionary")
    Set versionByName = CreateObject("Scripting.Dictionary")
    Set previousByName = CreateObject("Scripting.Dictionary")
    Set timestampByName = CreateObject("Scripting.Dictionary")
    Set headerNameByName = CreateObject("Scripting.Dictionary")
    Set versionCounts = CreateObject("Scripting.Dictionary")
    ConfigureTextDictionary componentsByName
    ConfigureTextDictionary timestampTextByName
    ConfigureTextDictionary typeByName
    ConfigureTextDictionary versionByName
    ConfigureTextDictionary previousByName
    ConfigureTextDictionary timestampByName
    ConfigureTextDictionary headerNameByName
    ConfigureTextDictionary versionCounts
    CaptureInstalledModuleInventory componentsByName, timestampTextByName, typeByName, versionByName, previousByName, timestampByName, headerNameByName, versionCounts, installedCount, expectedDefinitions
    evidence.Add componentsByName, "ComponentsByName"
    evidence.Add timestampTextByName, "TimestampTextByName"
    evidence.Add typeByName, "TypeByName"
    evidence.Add versionByName, "VersionByName"
    evidence.Add previousByName, "PreviousByName"
    evidence.Add timestampByName, "TimestampByName"
    evidence.Add headerNameByName, "HeaderNameByName"
    evidence.Add versionCounts, "VersionCounts"
    evidence.Add installedCount, "InstalledCount"
    evidence.Add expectedDefinitions, "ExpectedDefinitions"
    Set CaptureInstalledModuleInventoryEvidence = evidence
End Function

Private Sub LoadInstalledModuleInventoryEvidence(ByVal evidence As Collection, ByRef componentsByName As Object, ByRef timestampTextByName As Object, ByRef typeByName As Object, ByRef versionByName As Object, ByRef previousByName As Object, ByRef timestampByName As Object, ByRef headerNameByName As Object, ByRef versionCounts As Object, ByRef installedCount As Long)
    If evidence Is Nothing Then Err.Raise MODULE_GUARD_ERROR_BASE + 145, "LoadInstalledModuleInventoryEvidence", "The authoritative installed-module inventory evidence was not supplied."
    Set componentsByName = evidence("ComponentsByName")
    Set timestampTextByName = evidence("TimestampTextByName")
    Set typeByName = evidence("TypeByName")
    Set versionByName = evidence("VersionByName")
    Set previousByName = evidence("PreviousByName")
    Set timestampByName = evidence("TimestampByName")
    Set headerNameByName = evidence("HeaderNameByName")
    Set versionCounts = evidence("VersionCounts")
    installedCount = CLng(evidence("InstalledCount"))
End Sub

Private Sub RaiseModuleHeaderIdentityFailure(ByVal actualName As String, ByVal declaredName As String)
    Err.Raise MODULE_GUARD_HEADER_IDENTITY_ERROR, "CaptureInstalledModuleInventory", DEFAULT_UTILITY_SHORT_NAME & " module identity preflight failed before any structure, manifest, version, history, or registration checks." & vbCrLf & "Actual VBA component name: " & actualName & vbCrLf & "Declared header VB_Name: " & IIf(Len(declaredName) = 0, "(blank)", declaredName) & vbCrLf & "Likely cause: a complete module was pasted over the wrong VBA component." & vbCrLf & "Next step: replace the source in the actual component with the source intended for that component. Do not rename the component merely to match the pasted header unless that rename was intentional." & vbCrLf & "No Module Guard history or active-manifest change was attempted."
End Sub

Private Sub CaptureInstalledModuleInventory(ByVal componentsByName As Object, ByVal timestampTextByName As Object, ByVal typeByName As Object, ByVal versionByName As Object, ByVal previousByName As Object, ByVal timestampByName As Object, ByVal headerNameByName As Object, ByVal versionCounts As Object, ByRef installedCount As Long, Optional ByVal baselineDefinitions As Variant)
    Dim projectObject As Object
    Dim component As Object
    Dim sourceText As String
    Dim versionID As String
    Dim previousVersionID As String
    Dim timestampValue As Variant
    Dim headerName As String
    Dim declaredName As String
    Dim actualName As String
    Dim ComponentType As String
    Dim timestampText As String
    Dim ownedModuleNumber As Long
    Dim expectedOwnedModuleCount As Long
    Dim expectedDefinitions As Collection
    Dim actualOwned As Boolean
    Dim declaredOwned As Boolean
    If IsObject(baselineDefinitions) Then Set expectedDefinitions = baselineDefinitions
    If expectedDefinitions Is Nothing Then Set expectedDefinitions = InitialModuleGuardBaselineDefinitions()
    EnsureModuleGuardOwnedComponentLookup expectedDefinitions
    expectedOwnedModuleCount = expectedDefinitions.count
    Set projectObject = GuardVBProject()
    LogModuleGuardPreflightBoundary "CaptureInstalledModuleInventory", "Module Guard verification: single VBProject component enumeration starting.", "Project components: " & CStr(projectObject.VBComponents.count)
    For Each component In projectObject.VBComponents
        PumpModuleGuardSetupResponsiveness
        installedCount = installedCount + 1
        actualName = CStr(component.Name)
        sourceText = ComponentHeaderSourceText(component)
        declaredName = ModuleGuardDeclaredVBName(sourceText)
        actualOwned = ModuleGuardOwnsComponent(actualName)
        declaredOwned = False
        If Len(declaredName) > 0 Then declaredOwned = ModuleGuardOwnsComponent(declaredName)
        If actualOwned Or declaredOwned Then
            If Len(declaredName) = 0 Then RaiseModuleHeaderIdentityFailure actualName, declaredName
            If StrComp(actualName, declaredName, vbBinaryCompare) <> 0 Then RaiseModuleHeaderIdentityFailure actualName, declaredName
        End If
        If actualOwned Then
            ownedModuleNumber = ownedModuleNumber + 1
            LogModuleGuardPreflightBoundary "CaptureInstalledModuleInventory", "Module Guard verification: installed module " & CStr(ownedModuleNumber) & " of " & CStr(expectedOwnedModuleCount) & " starting.", "Module: " & actualName
            ReadNormalizedModuleGuardSourceMetadata sourceText, versionID, previousVersionID, timestampValue, headerName, timestampText
            ComponentType = GetComponentTypeText(CLng(component.Type))
            componentsByName(actualName) = True
            timestampTextByName(actualName) = timestampText
            typeByName(actualName) = ComponentType
            versionByName(actualName) = versionID
            previousByName(actualName) = previousVersionID
            timestampByName(actualName) = timestampValue
            headerNameByName(actualName) = headerName
            If IsValidChatGPTVersionID(versionID) Then
                If versionCounts.Exists(versionID) Then versionCounts(versionID) = CLng(versionCounts(versionID)) + 1 Else versionCounts(versionID) = 1
            End If
            LogModuleGuardPreflightBoundary "CaptureInstalledModuleInventory", "Module Guard verification: installed module " & CStr(ownedModuleNumber) & " of " & CStr(expectedOwnedModuleCount) & " completed.", "Module: " & actualName & vbCrLf & "Version: " & DisplayGuardValue(versionID) & vbCrLf & "Header identity and version metadata captured."
        End If
    Next component
    LogModuleGuardPreflightBoundary "CaptureInstalledModuleInventory", "Module Guard verification: single VBProject component enumeration completed.", "Project components enumerated: " & CStr(installedCount) & vbCrLf & DEFAULT_UTILITY_SHORT_NAME & "-owned modules inspected: " & CStr(ownedModuleNumber)
End Sub

Private Function GuardVBProject() As Object
    Dim projectObject As Object
    Dim activeProject As Object
    Dim projectFileName As String
    Dim currentFileName As String
    Dim projectName As String
    currentFileName = Replace(CurrentProject.FullName, "/", "\")
    For Each projectObject In Application.vbe.VBProjects
        PumpModuleGuardSetupResponsiveness
        projectFileName = vbNullString
        On Error Resume Next
        projectFileName = Replace(CStr(projectObject.fileName), "/", "\")
        On Error GoTo 0
        If Len(projectFileName) > 0 And StrComp(projectFileName, currentFileName, vbTextCompare) = 0 Then Set GuardVBProject = projectObject: Exit Function
    Next projectObject
    Set activeProject = Application.vbe.ActiveVBProject
    If Not activeProject Is Nothing Then
        projectName = vbNullString
        On Error Resume Next
        projectName = CStr(activeProject.Name)
        On Error GoTo 0
        If StrComp(projectName, CurrentProject.Name, vbTextCompare) = 0 Then Set GuardVBProject = activeProject: Exit Function
    End If
    Err.Raise MODULE_GUARD_ERROR_BASE + 42, "GuardVBProject", "Could not resolve the VBA project for the current CodeDb." & vbCrLf & "Database: " & CurrentProject.FullName
End Function

Private Function InstalledModuleVersionSummary(ByVal orderedDefinitions As Collection, ByVal versionByName As Object) As String
    Dim Definition As Variant
    Dim ModuleName As String
    Dim resultText As String
    On Error GoTo ErrHandler
    If orderedDefinitions Is Nothing Then Err.Raise MODULE_GUARD_ERROR_BASE + 145, "InstalledModuleVersionSummary", "The successful Module Guard operation did not supply its canonical module order for reporting."
    If versionByName Is Nothing Then Err.Raise MODULE_GUARD_ERROR_BASE + 146, "InstalledModuleVersionSummary", "The successful Module Guard operation did not supply its captured installed Version IDs for reporting."
    For Each Definition In orderedDefinitions
        PumpModuleGuardSetupResponsiveness
        ModuleName = CStr(Definition(0))
        If versionByName.Exists(ModuleName) Then
            AppendGuardDetail resultText, ModuleName & ": " & DisplayGuardValue(CStr(versionByName(ModuleName)))
        Else
            AppendGuardDetail resultText, ModuleName & ": (not captured)"
        End If
    Next Definition
    InstalledModuleVersionSummary = "Installed module IDs:" & vbCrLf & resultText
    Exit Function
ErrHandler:
    InstalledModuleVersionSummary = "Installed module IDs could not be summarized from the captured verification state." & vbCrLf & "Err.Description: " & Err.Description
End Function



Private Function TryParseChatGPTVersionTimestamp(ByVal timestampText As String, ByRef timestampValue As Date) As Boolean
    Dim firstSpace As Long
    Dim secondSpace As Long
    Dim thirdSpace As Long
    Dim weekdayText As String
    Dim dateText As String
    Dim timeText As String
    Dim meridiemText As String
    Dim remainingText As String
    Dim dateParts As Variant
    Dim timeParts As Variant
    Dim monthNumber As Long
    Dim dayNumber As Long
    Dim yearNumber As Long
    Dim hourNumber As Long
    Dim minuteNumber As Long
    Dim hour24 As Long
    Dim parsedDate As Date
    timestampText = Trim$(timestampText)
    If Len(timestampText) = 0 Then Exit Function
    firstSpace = InStr(1, timestampText, " ", vbBinaryCompare)
    If firstSpace <= 1 Then Exit Function
    weekdayText = Left$(timestampText, firstSpace - 1)
    remainingText = Trim$(Mid$(timestampText, firstSpace + 1))
    secondSpace = InStr(1, remainingText, " ", vbBinaryCompare)
    If secondSpace <= 1 Then Exit Function
    dateText = Left$(remainingText, secondSpace - 1)
    remainingText = Trim$(Mid$(remainingText, secondSpace + 1))
    thirdSpace = InStr(1, remainingText, " ", vbBinaryCompare)
    If thirdSpace <= 1 Then Exit Function
    timeText = Left$(remainingText, thirdSpace - 1)
    meridiemText = Trim$(Mid$(remainingText, thirdSpace + 1))
    If InStr(1, meridiemText, " ", vbBinaryCompare) > 0 Then Exit Function
    dateParts = Split(dateText, "/")
    If UBound(dateParts) <> 2 Then Exit Function
    timeParts = Split(timeText, ":")
    If UBound(timeParts) <> 1 Then Exit Function
    If TextContainsDecimalDigitsOnly(CStr(dateParts(0))) = False Or TextContainsDecimalDigitsOnly(CStr(dateParts(1))) = False Or TextContainsDecimalDigitsOnly(CStr(dateParts(2))) = False Then Exit Function
    If TextContainsDecimalDigitsOnly(CStr(timeParts(0))) = False Or TextContainsDecimalDigitsOnly(CStr(timeParts(1))) = False Then Exit Function
    If Len(CStr(dateParts(0))) < 1 Or Len(CStr(dateParts(0))) > 2 Then Exit Function
    If Len(CStr(dateParts(1))) < 1 Or Len(CStr(dateParts(1))) > 2 Then Exit Function
    If Len(CStr(dateParts(2))) <> 2 Then Exit Function
    If Len(CStr(timeParts(0))) < 1 Or Len(CStr(timeParts(0))) > 2 Then Exit Function
    If Len(CStr(timeParts(1))) <> 2 Then Exit Function
    monthNumber = CLng(dateParts(0))
    dayNumber = CLng(dateParts(1))
    yearNumber = 2000 + CLng(dateParts(2))
    hourNumber = CLng(timeParts(0))
    minuteNumber = CLng(timeParts(1))
    If monthNumber < 1 Or monthNumber > 12 Then Exit Function
    If dayNumber < 1 Or dayNumber > 31 Then Exit Function
    If hourNumber < 1 Or hourNumber > 12 Then Exit Function
    If minuteNumber < 0 Or minuteNumber > 59 Then Exit Function
    If StrComp(meridiemText, "AM", vbBinaryCompare) <> 0 And StrComp(meridiemText, "PM", vbBinaryCompare) <> 0 Then Exit Function
    parsedDate = DateSerial(yearNumber, monthNumber, dayNumber)
    If Year(parsedDate) <> yearNumber Or Month(parsedDate) <> monthNumber Or Day(parsedDate) <> dayNumber Then Exit Function
    If StrComp(weekdayText, EnglishWeekdayName(Weekday(parsedDate, vbSunday)), vbBinaryCompare) <> 0 Then Exit Function
    hour24 = hourNumber Mod 12
    If StrComp(meridiemText, "PM", vbBinaryCompare) = 0 Then hour24 = hour24 + 12
    timestampValue = parsedDate + TimeSerial(hour24, minuteNumber, 0)
    TryParseChatGPTVersionTimestamp = True
End Function

Private Function EnglishWeekdayName(ByVal weekdayNumber As Long) As String
    Select Case weekdayNumber
        Case vbSunday
            EnglishWeekdayName = "Sunday"
        Case vbMonday
            EnglishWeekdayName = "Monday"
        Case vbTuesday
            EnglishWeekdayName = "Tuesday"
        Case vbWednesday
            EnglishWeekdayName = "Wednesday"
        Case vbThursday
            EnglishWeekdayName = "Thursday"
        Case vbFriday
            EnglishWeekdayName = "Friday"
        Case vbSaturday
            EnglishWeekdayName = "Saturday"
    End Select
End Function

Private Function FormatChatGPTVersionTimestamp(ByVal timestampValue As Date) As String
    Dim hour12 As Long
    Dim meridiemText As String
    hour12 = Hour(timestampValue) Mod 12
    If hour12 = 0 Then hour12 = 12
    If Hour(timestampValue) < 12 Then meridiemText = "AM" Else meridiemText = "PM"
    FormatChatGPTVersionTimestamp = EnglishWeekdayName(Weekday(timestampValue, vbSunday)) & " " & CStr(Month(timestampValue)) & "/" & CStr(Day(timestampValue)) & "/" & Right$("0" & CStr(Year(timestampValue) Mod 100), 2) & " " & CStr(hour12) & ":" & Right$("0" & CStr(Minute(timestampValue)), 2) & " " & meridiemText
End Function

Private Function IsValidChatGPTVersionID(ByVal versionID As String) As Boolean
    Dim groups As Variant
    Dim groupValue As Variant
    Dim characterIndex As Long
    Dim characterText As String
    versionID = Trim$(versionID)
    If StrComp(versionID, MODULE_GUARD_TRACKING_PREFIX & MODULE_GUARD_PLACEHOLDER_VERSION_SUFFIX, vbBinaryCompare) = 0 Then Exit Function
    If Left$(versionID, Len(MODULE_GUARD_TRACKING_PREFIX)) <> MODULE_GUARD_TRACKING_PREFIX Then Exit Function
    groups = Split(Mid$(versionID, Len(MODULE_GUARD_TRACKING_PREFIX) + 1), "-")
    If UBound(groups) <> 4 Then Exit Function
    For Each groupValue In groups
        PumpModuleGuardSetupResponsiveness
        If Len(CStr(groupValue)) <> 4 Then Exit Function
        For characterIndex = 1 To 4
            If (characterIndex And 255) = 0 Then PumpModuleGuardSetupResponsiveness
            characterText = Mid$(CStr(groupValue), characterIndex, 1)
            If InStr(1, MODULE_GUARD_ID_ALPHABET, characterText, vbBinaryCompare) = 0 Then Exit Function
        Next characterIndex
    Next groupValue
    IsValidChatGPTVersionID = True
End Function

Private Function RollbackStandaloneVerificationTransaction(ByVal workspaceObject As DAO.workspace, ByRef cleanupDetails As String) As Boolean
    'Rolls back only a verification transaction that was started outside approved-build registration.
    On Error GoTo ErrHandler
    If workspaceObject Is Nothing Then
        AppendGuardDetail cleanupDetails, "Standalone verification rollback failed because the DAO Workspace object was not available."
        Exit Function
    End If
    If mModuleGuardRegistrationTransactionOwned Then
        AppendGuardDetail cleanupDetails, "Standalone verification rollback was not attempted because the outer approved-build registration owns the active transaction."
        Exit Function
    End If
    workspaceObject.Rollback
    AppendGuardDetail cleanupDetails, "Standalone verification transaction was rolled back."
    RollbackStandaloneVerificationTransaction = True
    Exit Function
ErrHandler:
    AppendGuardDetail cleanupDetails, "Cleanup failure while rolling back the standalone verification transaction. Err.Number=" & CStr(Err.Number) & "; Err.Source=" & Err.Source & "; Err.Description=" & Err.Description
End Function

Private Function RollbackRegistrationNestedVerificationTransaction(ByVal workspaceObject As DAO.workspace, ByRef cleanupDetails As String) As Boolean
    'Resolves only the verification transaction nested inside a module-owned registration transaction.
    On Error GoTo ErrHandler
    If workspaceObject Is Nothing Then
        AppendGuardDetail cleanupDetails, "Nested verification rollback failed because the DAO Workspace object was not available."
        Exit Function
    End If
    If mModuleGuardRegistrationTransactionOwned = False Or mModuleGuardNestedVerificationTransactionActive = False Then
        AppendGuardDetail cleanupDetails, "Nested verification rollback was not attempted because no registration-owned nested verification transaction was recorded."
        Exit Function
    End If
    workspaceObject.Rollback
    mModuleGuardNestedVerificationTransactionActive = False
    AppendGuardDetail cleanupDetails, "The failed nested verification transaction was rolled back; the owning registration transaction remains for outer rollback."
    RollbackRegistrationNestedVerificationTransaction = True
    Exit Function
ErrHandler:
    AppendGuardDetail cleanupDetails, "Cleanup failure while rolling back the nested verification transaction. Err.Number=" & CStr(Err.Number) & "; Err.Source=" & Err.Source & "; Err.Description=" & Err.Description
End Function

Private Sub AddVerificationRow(ByVal rs As DAO.Recordset, ByVal runID As String, ByVal verifiedOn As Date, ByVal ModuleName As String, ByVal ComponentType As String, ByVal isRequired As Boolean, ByVal isManifested As Boolean, ByVal expectedVersionID As String, ByVal installedVersionID As String, ByVal expectedPreviousVersionID As String, ByVal installedPreviousVersionID As String, ByVal expectedTimestamp As Variant, ByVal installedTimestamp As Variant, ByVal resultCode As String, ByVal resultNotes As String, ByVal updateBatchID As String)
    rs.AddNew
    rs!verificationRunID = runID
    rs!verifiedOn = verifiedOn
    rs.fields(FIELD_MODULE_NAME) = ModuleName
    rs.fields(FIELD_COMPONENT_TYPE) = Shared_String_ConvertToNullIfBlank(Trim$(ComponentType))
    rs!isRequired = isRequired
    rs!isManifested = isManifested
    rs!expectedVersionID = Shared_String_ConvertToNullIfBlank(expectedVersionID)
    rs!installedVersionID = Shared_String_ConvertToNullIfBlank(Trim$(installedVersionID))
    rs!expectedPreviousVersionID = Shared_String_ConvertToNullIfBlank(expectedPreviousVersionID)
    rs!installedPreviousVersionID = Shared_String_ConvertToNullIfBlank(Trim$(installedPreviousVersionID))
    rs.fields(FIELD_EXPECTED_VERSION_TIMESTAMP) = expectedTimestamp
    rs.fields(FIELD_INSTALLED_VERSION_TIMESTAMP) = installedTimestamp
    rs.fields(FIELD_VERIFICATION_RESULT) = resultCode
    rs.fields(FIELD_VERIFICATION_NOTES) = Shared_String_ConvertToNullIfBlank(Trim$(resultNotes))
    rs!updateBatchID = Shared_String_ConvertToNullIfBlank(updateBatchID)
    rs.Update
End Sub

Private Function DisplayGuardValue(ByVal valueText As String) As String
    If Len(Trim$(valueText)) = 0 Then
        DisplayGuardValue = "missing"
    Else
        DisplayGuardValue = Trim$(valueText)
    End If
End Function

Private Function GuardDisplayTimestamp(ByVal value As Variant) As String
    If IsNull(value) Then
        GuardDisplayTimestamp = "(blank)"
    Else
        GuardDisplayTimestamp = FormatChatGPTVersionTimestamp(CDate(value))
    End If
End Function

Private Sub AppendGuardDetail(ByRef targetText As String, ByVal detailText As String)
    If Len(Trim$(detailText)) = 0 Then Exit Sub
    If Len(targetText) > 0 Then targetText = targetText & " | "
    targetText = targetText & detailText
End Sub

Private Sub InsertOrVerifyHistoryRow(ByVal db As DAO.Database, ByVal updateBatchID As String, ByVal rowData As Variant, ByVal isBaseline As Boolean, ByVal recordedBy As String)
    Dim rs As DAO.Recordset
    Dim ModuleName As String
    Dim versionID As String
    ModuleName = CStr(rowData(0))
    versionID = CStr(rowData(2))
    Set rs = db.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ModuleUpdateHistory") & "] WHERE " & SQL_FIELD_MODULE_NAME & "=" & SharedDao_SqlQuote(ModuleName) & " AND " & SQL_FIELD_CHAT_GPT_VERSION_ID & "=" & SharedDao_SqlQuote(versionID), dbOpenDynaset)
    If rs.EOF Then
        rs.AddNew
        rs!updateBatchID = updateBatchID
        rs.fields(FIELD_MODULE_NAME) = ModuleName
        rs.fields(FIELD_COMPONENT_TYPE) = CStr(rowData(1))
        rs.fields(FIELD_CHAT_GPT_VERSION_ID) = versionID
        rs.fields(FIELD_PREVIOUS_CHAT_GPT_VERSION_ID) = Shared_String_ConvertToNullIfBlank(Trim$(CStr(rowData(3))))
        rs.fields(FIELD_VERSION_TIMESTAMP) = rowData(4)
        rs!changeDescription = CStr(rowData(5))
        rs.fields(FIELD_RECORDED_ON) = Now
        rs!recordedBy = recordedBy
        rs!isBaseline = isBaseline
        rs.fields(FIELD_SUPERSEDED_ON) = Null
        rs.fields(FIELD_MANAGED_OBJECT_DEPENDENCY_NOTES) = Null
        rs.Update
    Else
        If StrComp(Nz(rs!updateBatchID, vbNullString), updateBatchID, vbBinaryCompare) <> 0 Or StrComp(Nz(rs.fields(FIELD_COMPONENT_TYPE), vbNullString), CStr(rowData(1)), vbTextCompare) <> 0 Or StrComp(Nz(rs.fields(FIELD_PREVIOUS_CHAT_GPT_VERSION_ID), vbNullString), CStr(rowData(3)), vbBinaryCompare) <> 0 Or CBool(Nz(rs!isBaseline, False)) <> isBaseline Then rs.Close: Err.Raise MODULE_GUARD_ERROR_BASE + 50, "InsertOrVerifyHistoryRow", "An existing append-only history row conflicts with the requested module version." & vbCrLf & "Module: " & ModuleName & vbCrLf & "Version: " & versionID
        If IsNull(rs.fields(FIELD_VERSION_TIMESTAMP)) Xor IsNull(rowData(4)) Then rs.Close: Err.Raise MODULE_GUARD_ERROR_BASE + 51, "InsertOrVerifyHistoryRow", "An existing append-only history timestamp conflicts with the requested module version." & vbCrLf & "Module: " & ModuleName & vbCrLf & "Version: " & versionID
        If IsNull(rs.fields(FIELD_VERSION_TIMESTAMP)) = False Then
            If FormatChatGPTVersionTimestamp(CDate(rs.fields(FIELD_VERSION_TIMESTAMP))) <> FormatChatGPTVersionTimestamp(CDate(rowData(4))) Then rs.Close: Err.Raise MODULE_GUARD_ERROR_BASE + 52, "InsertOrVerifyHistoryRow", "An existing append-only history timestamp conflicts with the requested module version." & vbCrLf & "Module: " & ModuleName & vbCrLf & "Version: " & versionID
        End If
    End If
    rs.Close
End Sub
