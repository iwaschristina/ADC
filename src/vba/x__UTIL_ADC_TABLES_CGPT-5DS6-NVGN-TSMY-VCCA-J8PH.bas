Attribute VB_Name = "x__UTIL_ADC_TABLES"
' ================================================================================================================================================================================
' ChatGPT Rewrite Tracking Header
' ================================================================================================================================================================================
' ChatGPT Previous Version ID: CGPT-B4AB-5JDU-JWNL-UARL-FPJY
' ChatGPT Baseline Version ID: CGPT-B4AB-5JDU-JWNL-UARL-FPJY
' ChatGPT Version ID: CGPT-5DS6-NVGN-TSMY-VCCA-J8PH
' ChatGPT Version Timestamp: Tuesday 9/1/26 4:39 PM
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
' STANDARD MODULE: Access Development Console Table Maintenance
' VB_Name: x__UTIL_ADC_TABLES
' Entry Points:
'     StartTraceRunForSetup
'     EnsureSetupDiagnosticsReadyForPreflight
'     PrepareSetupDataTableBackupForPreflight
'     RepairCoreDataTableFoundationForSetup
'     RepairHealthReviewTableForSetup
'     RepairHostDbConfigTable
'     RepairNonFormQueryFoundationForSetup
'     RepairOrReplaceQueryForSetup
' SHARED CODE:
'     SharedAccess_PrepareOpenObjectForDesignChange(ByVal objectType As AcObjectType, ByVal objectName As String, ByVal actionText As String, Optional ByVal confirmationAlreadyProvided As Boolean = False) -> Boolean
'     SharedAccess_PrepareOpenObjectsForDesignChange(ByVal objectType As AcObjectType, ByVal objectNames As Collection, ByVal actionText As String, Optional ByVal confirmationAlreadyProvided As Boolean = False) ->
'     Boolean
'     SharedCollection_Join(ByVal values As Collection, Optional ByVal delimiter As String = ", ") -> String
'     SharedDao_BracketName(ByVal objectName As String) -> String
'     SharedDao_DatabaseTableExistsSafe(ByVal db As DAO.Database, ByVal tableName As String) -> Boolean
'     SharedDao_DatabaseTableFieldExistsSafe(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String) -> Boolean
'     SharedDao_FieldDefinitionMatches(ByVal fld As DAO.Field, ByVal expectedType As DAO.DataTypeEnum, Optional ByVal expectedSize As Long = 0, Optional ByVal mustBeAutoNumber As Boolean = False, Optional ByVal
'     allowLargerTextSize As Boolean = False) -> Boolean
'     SharedDao_QueryExists(ByVal db As DAO.Database, ByVal queryName As String) -> Boolean
'     SharedDao_RecordsetHasFieldSafe(ByVal rs As DAO.Recordset, ByVal fieldName As String) -> Boolean
'     SharedDao_RelationExists(ByVal db As DAO.Database, ByVal relName As String) -> Boolean
'     SharedDao_SafeDatabaseName(ByVal db As DAO.Database) -> String
'     SharedDao_SqlQuote(ByVal value As Variant) -> String
'     SharedDao_TableDefFieldExistsSafe(ByVal tdf As DAO.tableDef, ByVal fieldName As String) -> Boolean
'     SharedDao_TableDefIndexExistsSafe(ByVal tdf As DAO.tableDef, ByVal indexName As String) -> Boolean
'     SharedDao_TableDefSingleFieldIndexMatches(ByVal tdf As DAO.tableDef, ByVal indexName As String, ByVal fieldName As String, ByVal uniqueIndex As Boolean, ByVal primaryIndex As Boolean) -> Boolean
'     SharedDao_TableExists(ByVal db As DAO.Database, ByVal tableName As String) -> Boolean
'     SharedDao_TableFieldExists(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String) -> Boolean
'     SharedMessage_OneLine(ByVal valueText As String, Optional ByVal maxLength As Long = 700, Optional ByVal ellipsisText As String = "...") -> String
'     SharedName_SafeAccessObjectName(ByVal value As String, Optional ByVal maxLength As Long = 0, Optional ByVal firstCharacterPrefix As String = "DB_", Optional ByVal fallbackName As String = vbNullString) ->
'     String
'     SharedText_AppendNonBlankLine(ByVal existingText As String, ByVal lineText As String) -> String
'     SharedText_Nz(ByVal v As Variant) -> String
'     Shared_String_ConvertToNullIfBlank(ByVal v As Variant) -> Variant
'     SharedValue_NzLong(ByVal value As Variant, Optional ByVal defaultValue As Long = 0) -> Long
' CLASSES USED: None
' DYNAMIC STRING REFERENCES:
'     Dynamic-reference scan stopped after 2000 logical line(s) to prevent a long-running scan; review this module manually if needed.
' LastScanTimestamp: Saturday 8/15/26 8:47 AM
' Purpose:
'     Owns physical ADC table/schema creation and repair, Settings/storage foundation already assigned here, data-table backup/reset, launcher table foundation, the complete canonical A3 non-form saved-query
'     membership and QueryDef repair path, trace support, naming-backed physical object helpers, and structural verification required by A2/A3. Physical progress presentation remains UI-owned.
' Author: Christina Shonts
' Support:
'     Access Development Console CodeDb table maintenance, required-row/default-value repair, catalog repair, and HostDb MacroCatalog_Records table repair.
' Code Category:
'     Access Development Console current-schema table maintenance; CodeDb-owned support table repair and table-backed catalog/reference maintenance.
' Requirements:
'     DAO is required.
'     The caller must run this from the CodeDb Access Development Console utility environment for CodeDb-owned table maintenance.
'     HostDb work must remain limited to HostDb MacroCatalog_Records through explicit HostDb MacroCatalog_Records repair paths.
'     Shared DAO, SQL, and progress/status helpers must be available in the project when this module is imported without the older core-module copies.
' Reference Dependencies:
'     DAO, Access object model constants, and shared DAO/status helper modules are required.
' DB Object Dependencies:
'     CodeDb-owned tables include MacroCatalog_Records, Settings, MacroCatalog_Cats, MacroCatalog_Cats_SubCats, MacroCatalog_Sys, TraceLog, UI_Ribbon_Generated_BtnTargets, Help_Docs, Help_Terms, UI_Groups,
'     UI_Frm_Lookup_Types, UI_Frm_Registry, MAC_Cats, MAC, MAC_ActionPayloads,  and MAC_Deps.
'     HostDb-owned data for this module is limited to MacroCatalog_Records in the HostDb.
'     Review, lookup, and form-design support queries remain CodeDb-owned support objects unless generated  macro workflows explicitly target HostDb macro objects.
' Usage:
'     Run RepairCoreDataTableFoundationForSetup for canonical CodeDb core-table readiness; RepairTables owns the runtime/manual table-repair orchestration in this module.
'     Run RepairHostDbConfigTable only for the narrow HostDb MacroCatalog_Records table path.
'     Generated system macros route through RunSystemCommand; keep only parameterized RunCode wrappers  that still have a direct purpose.
'     Do not use this module as a broad HostDb setup routine or as a form/UI builder.
' Version:
'     Shifts current Navigation Pane and Ribbon lifecycle references to Phase I and Phase H after insertion of Setup dependency Phase F.
'     Extends A2 Preflight schema detection, affected-table backup planning, backup-coverage enforcement, and final structural proof to HealthReview and launcher-review physical constraints.
'     A2 ownership remains unchanged.
'     Enforces MC_Records.MacroKey as required, nonblank, and unique in the canonical A2/HostDb physical schema, with lossless pre-repair diagnostics for invalid existing rows.
'     Adds physical-file-only timing around canonical DAO relationship operations so Setup pauses can be isolated without reopening the TraceLog table sink or changing phase-owned work.
'     Adds a physical-file-only trace writer for early startup diagnostics that bypasses TraceLog table access and persisted trace settings while preserving TABLES sink ownership and Setup-phase boundaries.
'     Adds a dedicated cumulative Run_RunFullSetup startup log and state-isolated replay so wrapper-entry evidence survives later Setup invocations and trace-log recreation without touching normal trace handles,
'     counters, or run IDs.
'     Replaces physical trace-log truncation with delete-and-recreate behavior for trace-run starts and the public clear-log command; Setup-phase ownership and trace-table retention are unchanged.
'     Removes relationship state from the A2 foundation fingerprint and moves all canonical relationship creation/repair to B5; Preflight and A2 now own only their table/schema/index work, while deferred
'     legacy physical retirement remains compatible with B5-owned relationship cleanup.
'     Scopes the A2 foundation schema fingerprint to A2-owned relationship state so later B5 relationship cleanup cannot invalidate the persisted A2 proof, and forces the A2 repair path while deferred
'     legacy physical retirement remains pending.
'     Removes safe remaining Trim$ normalization from exact table/query proof/cache and field-description table identity paths.
'     Restores fresh A2 Settings/table existence and structural verification after responsiveness/event boundaries; removes unsafe prepared/exact-schema proof reuse while retaining A2 ownership.
'     Routes Managed Object Catalog references to x__UTIL_ADC_MOC and physical lifecycle references to x__UTIL_ADC_MOC_LIFECYCLE after the responsibility split; runtime and Setup-phase behavior are unchanged.
'     Resets saved-query SQL-mismatch diagnostic throttling at A3 batch start and for each standalone repair call so the five-detail cap has an explicit operation lifetime.
'     Completes direct RepairTables by reconciling Descriptions for the exact A2 table/field and A3 non-form query rosters.
'     Adds an explicit suppression option for broader coordinators that own a later complete Description pass.
'     Preserves direct cold-start repair before MAC_Fields exists by allowing only the canonical Defaults field-description seed fallback in that unavailable-catalog state.
'     Adds B5-owned non-cascading RI for retained Reference/Add-In project, reference, module, procedure, dependency, and resolved-identity paths; A4 remains the participating table/index foundation owner.
'     Uses typed A2 schema-repair evidence and named lookup-property masks, uses the canonical CodeDb role constant directly, and narrows module-internal visibility.
'     Updated Setup table/query creation contracts to defer physical Navigation Pane placement to Phase H and corrected Ribbon infrastructure lifecycle wording to Phase G.
'     Updated DEFAULTS dependency metadata and stale lookup-error wording for the standard-module authority; lookup behavior is unchanged.
'     Removed the redundant ADCFoundationRequiredTableNames collection reconstruction and routed its three consumers directly to FoundationRequiredTableNames; canonical managed-object identity
'     validation already prevents duplicate foundation-table identities.
'     Reused one operation-local Settings, MacroCatalog_Records, MacroCatalog_Cats, and MacroCatalog_Cats_SubCats schema set through A2 detection, repair, and final validation and accepted caller-owned Help_Terms
'     schema collections for B1.
'     Made A3 own one local five-family FormDesign query work plan, resolving the canonical managed-object definitions once and passing each ordered collection directly to its repair procedure instead of requesting
'     five independently cached family projections.
'     Removed rel_NavigationPaneGroups_SettingKey from the canonical RI registry because UI_Groups.SettingKey is a code-owned routing/configuration identifier rather than a Settings foreign key; preserved
'     managed-object Navigation Pane group RI and generic fail-closed enforcement.
'     Enforced exact field membership for the ADC-owned MacroCatalog_Cats and MacroCatalog_Cats_SubCats schemas so A2 detects and removes obsolete physical fields instead of allowing F1 to encounter unmapped stale
'     metadata.
'     Added the authoritative raw Settings writer, routed HostDb MacroCatalog_Records database resolution through Core, and routed cancellation publication through Core.
'     Inlined the duplicate-review MacroCatalog_Records source-table expression into DuplicateReviewSql and removed the unnecessary private micro-helper.
'     Removed TABLES' independent required-query family classifier and routed foundation fingerprint membership through the canonical DEFAULTS classification rows without adding an extra registry scan.
'     Corrected TABLES ownership documentation to cover only the assigned core/shared, Settings-storage, and support-query physical families rather than all ADC physical schema.
'     Consolidated runtime table backup/reset/repair and launcher-foundation storage operations into this physical table/storage authority.
'     Kept shared progress text semantic through dispatch and moved status-bar hierarchy formatting to the display-only boundary so Setup Control never receives presentation-formatted separators.
'     Delegated nested progress composition to Core and owns the shared Null-returning optional recordset field-value helper.
'     Removed TABLES as a physical status-bar writer; table progress/status and trace completion now delegate visible set/clear operations to the UI authority.
'     Redirected the remaining Tables confirmation caller to the centralized Core popup-result authority.
'     Removed hard-coded FormDesign query-family classification; foundation snapshot classification now consumes canonical DEFAULTS metadata, with names and families cached from one registry projection.
'     Consolidated Help_Terms physical table and field creation/repair in TABLES; the Help_Terms path no longer delegates those mutations to Catalog.
'     Centralized physical ADC backup-table naming/copy creation in CreateBackupTableCopy so Setup and runtime workflows share one mutation authority.
'     Removed six confirmed unreachable list/trace/review-foundation procedures.
' Notes:
'     Live Access Development Console tables use stable ADC-prefixed names.
'     Required rows and default values are maintained directly in live CodeDb-owned managed tables by setup/repair routines.
'     This module is not a second table layer and must not describe or populate a parallel table family.
'     Public procedures are intentionally preserved only when useful for manual execution, generated system macros, cross-module orchestration, parameterized catalog list access, or future form buttons; internal
'     helpers remain Private.
'     Fully retired terminology rows use TermStatus=Retired and are not also marked as temporary Deprecated rows.
' Core Rules Enforced Here:
'     Reporting ownership rule: public action entry points own one diagnostic popup; table/shared helpers log/raise upward and must not show duplicate popups.
'     Repair the current live managed-table family in CodeDb unless a caller explicitly routes to the narrow HostDb MacroCatalog_Records table repair path.
'     Do not define a parallel duplicate table family through active table names, procedure names, comments, defaults, or user-facing messages.
'     Use explicit CodeDb/HostDb language and keep any CurrentDb API use isolated behind role-specific helper functions.
'     Keep generated/system macro table rows aligned with RunSystemCommand expressions and callable direct entry points.
' Design Map:
'     RepairCoreDataTableFoundationForSetup owns core physical schema/default readiness including HealthReview; RepairNonFormQueryFoundationForSetup owns complete A3 saved-query membership and physical QueryDef
'     repair; RepairTables orchestrates these TABLES authorities for runtime/manual repair.
'     HostDb MacroCatalog_Records repair is isolated in RepairHostDbConfigTable and supporting helpers.
'     TABLES owns the core/shared table, Settings-storage, and support-query families explicitly assigned to this module, including Help_Terms physical repair and its B2 default/reference rows; domain-specific
'     physical schema owned by References, Module Guard, Catalog, and other authorities remains with those owners. A3 review-foundation QueryDef repair remains TABLES-owned as described above.
'     Catalog owns B-phase catalog reconciliation for MAC and its lookup tables, MacroCatalog_Sys, and Help_Docs.
'     Backup/reset helpers operate on the live table family and should stay table-module owned.
' Configuration:
'     ADCCatalog* name helpers describe current live utility objects and catalog values, not a separate table family.
'     Managed table-name constants must continue to reflect the accepted V6 live table family.
'     Backup tables are grouped only under Access Development Console - Backup and are not placed in normal table/catalog/all-object groups.
'     Help_Terms uses TermStatus for Approved, Conditional, Deferred, and Retired states; TermStatus=Retired means not valid current terminology. Deprecated=True is reserved for terms still potentially valid only in
'     a limited, discouraged, or transitional current context.
' Behavior:
'     Public repair entry points return or report success/failure and may show user-facing messages when requested.
'     The module does not create the full Access Development Console utility environment in HostDb.
'     Catalog/default-value repair preserves user-editable fields where the table workflow allows it.
' Code Status:
'     Active current-design table-maintenance module. Table-related routines already present here are treated as owned by  System macro object descriptions and managed-table field descriptions are
'     maintained as development-facing metadata.
' Todo: None.
' Reporting Bugs: None currently recorded in this header.
' Module Level Bugs: None currently recorded in this header.
' Public Subs:
'     EnsureUSysRibbonsTableSchema() -> None
'     EnsureSetupDiagnosticsReadyForPreflight(Optional ByRef detailText As String = vbNullString) -> None
'     PrepareSetupDataTableBackupForPreflight(Optional ByRef detailText As String = vbNullString) -> None
'     RequireSetupPreflightTerminologyBackupCoverage(Optional ByVal suppliedSchemaDefinitions As Variant, Optional ByVal suppliedIndexDefinitions As Variant) -> None
'     RepairTerminologyDefaultRowsForSetup() -> None
'     EnsureTerminologyReviewQueryForSetup() -> None
'     RepairDuplicateReviewQueryForSetup() -> None
'     RepairCanonicalRelationshipsForSetup() -> None
'     AddCanonicalRelationshipFindings(ByRef findings As Collection) -> None
'     SafeDropTableIfExists(ByVal tableName As String, Optional ByVal successNote As String = vbNullString) -> None
'     BeginTraceTableWriteSuppression() -> None
'     BeginTerminalTraceTableTeardown() -> None
'     EndTraceTableWriteSuppression() -> None
'     StartTraceRun(Optional ByVal contextName As String = vbNullString) -> None
'     StartTraceRunForSetup(Optional ByVal contextName As String = vbNullString) -> None
'     LogRunFullSetupEntryPhysicalMarker() -> None
'     ClearRunFullSetupEntryPhysicalMarker() -> None
'     CompleteTraceRun(Optional ByVal completeMessage As String = vbNullString, Optional ByVal clearStatus As Boolean = True) -> None
'     LogTraceMessage(ByVal traceLevel As String, ByVal traceCategory As String, ByVal sourceProcedure As String, ByVal traceMessage As String, Optional ByVal detailText As String = vbNullString, Optional ByVal
'     stepName As String = vbNullString, Optional ByVal errNumber As Long = 0, Optional ByVal errSource As String = vbNullString, Optional ByVal errDescription As String = vbNullString, Optional ByVal objectName
'     As String = vbNullString, Optional ByVal objectType As String = vbNullString, Optional ByVal databaseRole As String = vbNullString, Optional ByVal contextName As String = vbNullString) -> None
'     LogPhysicalTraceMessageOnly(ByVal traceLevel As String, ByVal traceCategory As String, ByVal sourceProcedure As String, ByVal traceMessage As String, Optional ByVal detailText As String = vbNullString,
'     Optional ByVal stepName As String = vbNullString, Optional ByVal errNumber As Long = 0, Optional ByVal errSource As String = vbNullString, Optional ByVal errDescription As String = vbNullString,
'     Optional ByVal objectName As String = vbNullString, Optional ByVal objectType As String = vbNullString, Optional ByVal databaseRole As String = vbNullString,
'     Optional ByVal contextName As String = vbNullString, Optional ByVal traceRunIDOverride As String = vbNullString) -> None
'     ClearTraceLogFile(Optional ByVal showMessage As Boolean = True, Optional ByVal confirmDestructiveAction As Boolean = True) -> None
'     FlushTraceLogFile() -> None
'     ReleaseTraceLogWriteRecordsetForSchemaMutation() -> None
'     ResetTraceNotes() -> None
'     AddTraceNote(ByVal traceText As String, Optional ByVal suppressImmediateWrite As Boolean = False) -> None
'     SetSettingRawText(ByVal SettingKey As String, ByVal SettingValue As String, Optional ByVal prepareSettings As Boolean = True, Optional ByVal updateTimestamps As Boolean = True, Optional ByVal wrapError As
'     Boolean = True) -> None
'     RepairHealthReviewTableForSetup(Optional ByVal applyLookupMetadata As Boolean = True) -> None
'     EnsureTerminologyTable(Optional ByVal applyLookupMetadata As Boolean = True, Optional ByVal suppliedSchemaDefinitions As Variant, Optional ByVal suppliedIndexDefinitions As Variant) -> None
'     RepairHostDbConfigTable(Optional ByVal showMessage As Boolean = True) -> None
'     RepairLauncherTableFoundationForSetup(Optional ByVal showMessage As Boolean = False, Optional ByVal configMetadataReady As Boolean = False, Optional ByVal settingsDefaultsAlreadyReady As Boolean = False,
'     Optional ByVal applyLookupMetadata As Boolean = True) -> None
'     RepairCodeDbSettingsTableForSetup(Optional ByVal applyLookupMetadata As Boolean = True, Optional ByVal suppliedSchemaDefinitions As Variant, Optional ByVal suppliedIndexDefinitions As Variant) -> None
'     RepairSettingDefaultsForSetup(Optional ByVal showMessage As Boolean = True, Optional ByVal suppliedCoreSettingDefinitions As Variant) -> None
'     RepairOrReplaceQueryForSetup(ByVal queryName As String, ByVal querySql As String, Optional ByVal showMessage As Boolean = True) -> None
'     SetRecordsetFieldIfExists(ByVal rs As DAO.Recordset, ByVal fieldName As String, ByVal fieldValue As Variant, Optional ByVal valueSourceName As String = vbNullString) -> None
'     SetRecordsetFieldValueWithSource(ByVal rs As DAO.Recordset, ByVal fieldName As String, ByVal fieldValue As Variant, ByVal valueSourceName As String) -> None
'     BackupAllUserData(Optional ByVal showMessage As Boolean = True) -> None
'     ResetAllUserData(Optional ByVal showMessage As Boolean = True, Optional ByVal confirmDestructiveAction As Boolean = True) -> None
'     PrepareOpenObjectsForManagedNameUpdate(ByVal codeDb As DAO.Database, ByVal hostDb As DAO.Database, ByVal renames As Collection) -> None
'     RepairLauncherFoundation(Optional ByVal showMessage As Boolean = True, Optional ByVal configMetadataReady As Boolean = False, Optional ByVal settingsDefaultsAlreadyReady As Boolean = False) -> None
'     EnsureDaoFieldDescription(ByVal fld As DAO.Field, ByVal descriptionText As String, Optional ByVal tableName As String = vbNullString) -> None
'     ReadDaoFieldDescriptionState(ByVal fld As DAO.Field, ByRef descriptionText As String, ByRef propertyType As Long, ByRef propertyExists As Boolean) -> None
' Public Functions:
'     USysRibbonsTableSchemaCurrentForSetup(Optional ByRef detailText As String = vbNullString) -> Boolean
'     TryGetSetupPreflightA2SchemaEvidence(ByRef repairNeeded As Boolean, ByRef repairSummary As String) -> Boolean
'     TablesConfiguredPhysicalObjectNamesForSetup() -> Collection
'     CoreDataTableFoundationCurrentForSetup(ByRef detailText As String) -> Boolean
'     CanonicalRelationshipDefinitions() -> Collection
'     LookupRecordReferenceUsage(ByVal lookupTableName As String, ByVal lookupRowValues As Object, ByRef usageDetail As String) -> Boolean
'     SafeDaoTableDefName(ByVal tdf As DAO.tableDef) -> String
'     ClearFieldLookupPresentationProperties(ByVal tdf As DAO.tableDef, ByVal fieldName As String) -> Long
'     EnsureFieldTableLookup(ByVal tdf As DAO.tableDef, ByVal fieldName As String, ByVal lookupTableName As String, ByVal codeFieldName As String, ByVal displayFieldName As String, Optional ByVal LimitToList As
'     Boolean = True, Optional ByVal columnWidthsText As String = "0;2") -> Long
'     CodeDbTableExists(ByVal tableName As String) -> Boolean
'     TableCanOpen(ByVal tableName As String, Optional ByVal tableExistsAlreadyKnown As Boolean = False) -> Boolean
'     TableOpenFailureDetail(ByVal tableName As String) -> String
'     TryDeleteTableDef(ByVal db As DAO.Database, ByVal tableName As String, ByRef errorText As String, Optional ByRef errorNumberResult As Long = 0, Optional ByRef errorSourceResult As String = vbNullString,
'     Optional ByRef errorDescriptionResult As String = vbNullString) -> Boolean
'     TableFieldExists(ByVal tableName As String, ByVal fieldName As String) -> Boolean
'     TableDiagnosticsLocationForMessage() -> String
'     TableReadinessDiagnosticsForMessage() -> String
'     GetActiveOrLastTraceRunID() -> String
'     GetActiveTraceRunID() -> String
'     TraceRunActive() -> Boolean
'     TraceSummaryForMessage(Optional ByVal reportContext As String = "failure") -> String
'     SettingText(ByVal SettingKey As String, ByVal DefaultValue As String) -> String
'     SettingBool(ByVal SettingKey As String, ByVal DefaultValue As Boolean) -> Boolean
'     SettingLong(ByVal SettingKey As String, ByVal DefaultValue As Long) -> Long
'     SettingRowExists(ByVal SettingKey As String) -> Boolean
'     LogReportedError(ByVal descriptionText As String, ByVal screenshotPath As String, ByVal captureWidth As Long, ByVal captureHeight As Long) -> Long
'     TraceBackendDiagnosticsForMessage() -> String
'     HealthReviewFieldOrderText() -> String
'     HealthReviewFieldNames() -> Variant
'     RepairCoreDataTableFoundationForSetup(ByRef schemaRepairEvidence As TSchemaRepairPreflightEvidence, Optional ByVal showMessage As Boolean = False, Optional ByRef repairDetailsForMessage As String =
'     vbNullString, Optional ByRef namingInputsChanged As Boolean = False, Optional ByVal requireSetupPreflightBackup As Boolean = True, Optional ByVal applyStandaloneDescriptions As Boolean = True) -> Boolean
'     MacroClassificationStateCurrentForB1(ByRef detailText As String) -> Boolean
'     SettingsFieldSchemaDefinitions() -> Collection
'     RepairNonFormQueryFoundationForSetup(Optional ByVal showMessage As Boolean = False, Optional ByRef repairDetailsForMessage As String = vbNullString,
'     Optional ByVal applyStandaloneDescriptions As Boolean = True) -> Boolean
'     NonFormQueryFoundationCurrentForSetup(ByRef detailText As String) -> Boolean
'     ConfigFieldOrderText() -> String
'     SettingsFieldOrderText() -> String
'     CustomMacroLauncherReviewFieldOrderText() -> String
'     FoundationOwnedTableFieldOrderText(ByVal tableName As String) -> String
'     TerminologyFieldSchemaDefinitions() -> Collection
'     TerminologyIndexSchemaDefinitions() -> Collection
'     ValidationRuleFromValueList(ByVal fieldName As String, Optional ByVal tableName As String = vbNullString) -> String
'     FieldDefinitionDescription(ByVal fieldType As DAO.DataTypeEnum, Optional ByVal fieldSize As Long = 0, Optional ByVal isAutoNumber As Boolean = False) -> String
'     FieldActualDefinitionDescription(ByVal fld As DAO.Field) -> String
'     EnsureFieldValueListLookup(ByVal tdf As DAO.tableDef, ByVal fieldName As String, ByVal rowSourceText As String, Optional ByVal BoundColumn As Long = 0, Optional ByVal LimitToList As Boolean = True) -> Long
'     FieldValueIfExists(ByVal rs As DAO.Recordset, ByVal fieldName As String) -> Variant
'     TryDeleteQueryDef(ByVal db As DAO.Database, ByVal queryName As String, ByRef errorText As String, Optional ByRef errorNumberResult As Long = 0, Optional ByRef errorSourceResult As String = vbNullString,
'     Optional ByRef errorDescriptionResult As String = vbNullString) -> Boolean
'     FieldDescriptionTextIsPlaceholder(ByVal descriptionText As String) -> Boolean
'     NormalizeLookupColumnWidthsForAccess(ByVal columnWidthsText As String, ByVal columnCount As Long, ByVal BoundColumn As Long, ByVal LimitToList As Boolean) -> String
'     TryDeleteAccessObject(ByVal objectType As AcObjectType, ByVal objectName As String, ByRef errorText As String, Optional ByRef errorNumberResult As Long = 0, Optional ByRef errorSourceResult As String =
'     vbNullString, Optional ByRef errorDescriptionResult As String = vbNullString) -> Boolean
'     FormDesignQueryFoundationCurrentForSetup(ByRef detailText As String) -> Boolean
'     BackupAllUserDataCore(Optional ByVal showMessage As Boolean = True) -> Long
'     RepairTables(Optional ByVal showMessage As Boolean = True, Optional ByRef schemaRepairDetailsForMessage As String = vbNullString, Optional ByVal applyStandaloneDescriptions As Boolean = True) -> Boolean
' Public Properties: None
' ================================================================================================================================================================================

Option Compare Database
Option Explicit

' ===============================================================================================================================================================================
' MODULE-LEVEL DECLARATIONS
' ===============================================================================================================================================================================

' Same-run A2 schema preflight evidence transferred atomically across the Setup/TABLES boundary.
Public Type TSchemaRepairPreflightEvidence
    IsKnown As Boolean
    repairNeeded As Boolean
    Summary As String
End Type

' Foundation and trace support structures.
Private Type FoundationDomainSnapshot
    SchemaFingerprint As String
    Part2DataFingerprint As String
    CoreQueryFingerprint As String
    FormDesignSourceFingerprint As String
    LookupQueryFingerprint As String
    tableCount As Long
    coreQueryCount As Long
    FormDesignSourceCount As Long
    LookupQueryCount As Long
    failureDetail As String
End Type
Private Type TTraceSinkFailure
    Number As Long
    Source As String
    Description As String
End Type

' Settings repair counters and active batch objects.
Private mSettingRepairSummaryActive As Boolean
Private mSettingRowsChecked As Long
Private mSettingRowsExistingPreserved As Long
Private mSettingRowsInserted As Long
Private mSettingBatchActive As Boolean
Private mSettingBatchDb As DAO.Database
Private mSettingBatchRs As DAO.Recordset

' Table-design batch preparation state.
Private mTableDesignBatchPrepared As Boolean

' Saved-query repair batch state and mismatch diagnostics.
Private mQueryMismatchDiagnosticCount As Long
Private mQueryRepairBatchActive As Boolean
Private mQueryRepairBatchChecked As Long
Private mQueryRepairBatchCreated As Long
Private mQueryRepairBatchRefreshed As Long
Private mQueryRepairBatchUnchanged As Long
Private mQueryRepairBatchStarted As Single
Private mQueryRepairBatchDb As DAO.Database
' INTENTIONAL KEYED OPERATION CACHE:
' Stores query name -> SQL already verified inside the active saved-query repair batch so post-batch consumers use the exact checked text. Begin/end batch owns the lifetime; the dictionary is derived from QueryDefs
' and is never an independent query-definition authority.
Private mQueryRepairBatchVerifiedEvidence As Object

' Foundation fingerprint, compact-state, and query-family metadata.
Public Const A2_B2_LOOKUP_DEFINITION_PENDING_FIELDS_PROPERTY As String = "ADCA2B2LookupDefinitionPendingFields"
Private Const FOUNDATION_SCHEMA_FINGERPRINT_PROPERTY As String = "ADCFoundationSchemaFingerprint"
Private Const FOUNDATION_PART2_DATA_FINGERPRINT_PROPERTY As String = "ADCFoundationPart2DataFingerprint"
Private Const FOUNDATION_CORE_QUERY_FINGERPRINT_PROPERTY As String = "ADCFoundationCoreQueryFingerprint"
Private Const FOUNDATION_FORMDESIGN_SOURCE_FINGERPRINT_PROPERTY As String = "ADCFoundationFormDesignSourceFingerprint"
Private Const FOUNDATION_LOOKUP_QUERY_FINGERPRINT_PROPERTY As String = "ADCFoundationLookupQueryFingerprint"
Private Const RETIRED_CONFIG_OWNER_DATABASE_ROLE_LOOKUP_QUERY As String = "MacroCatalog_Records_Lookup_RecordOwnerDBRole"
Private Const FOUNDATION_NONFORM_QUERY_FINGERPRINT_PROPERTY As String = "ADCFoundationNonFormQueryFingerprint"
Private Const FOUNDATION_NONFORM_QUERY_FINGERPRINT_VERSION As String = "A3-v1-2026-07-23"
Private Const FOUNDATION_FINGERPRINT_VERSION As String = "2026-07-15 Part 3 query fingerprint v5 - SQL covers parameter declarations without DAO parameter compilation"
Private Const FOUNDATION_SCHEMA_FINGERPRINT_VERSION As String = "2026-09-01 A2 physical schema fingerprint v13 - excludes all B5-owned relationship state while retaining A2 table, field, index, and physical-retirement ownership"

' Trace-version logging and shared activity-dispatch constants.
Private Const LOG_CHATGPT_VERSION_IDS_AT_TRACE_START As Boolean = True
' Dedicated cumulative startup log retained independently across Setup invocations while the normal physical trace log is deleted and recreated.
Private Const RUN_FULL_SETUP_ENTRY_LOG_FILE_NAME As String = "Access Development Console Setup Startup.log"

' Trace-run, sink, file, table, and write-state objects.
' INTENTIONAL ORDERED TRACE NOTES:
' Holds recent trace/status notes in occurrence order for the active trace run and summary display. ResetTraceNotes owns invalidation; TABLES retains the mutable Collection and exposes formatted summaries rather
' than the Collection itself.
Private mTraceNotes As Collection
Private mTraceRunID As String
Private mLastTraceRunID As String
' Pre-Setup wrapper-entry marker state retained only until the physical trace log is recreated or the wrapper exits.
Private mRunFullSetupEntryMarkerPending As Boolean
Private mRunFullSetupEntryMarkerTimestamp As Date
Private mTraceTableWriteFailureCount As Long
Private mTraceTableLastFailureText As String
Private mTraceLogFileWriteFailureCount As Long
Private mTraceLogFileLastFailureText As String
Private mTraceLogFileClearFailureText As String
Private mTraceSinkDepth As Long
Private mTraceLogTableReadyRunID As String
Private mTraceLogFileNumber As Integer
Private mTraceLogFilePathOpen As String
Private mTraceLogFileOpen As Boolean
Private mTraceLogFileOwnerRunID As String
Private mTraceLogWriteDatabase As DAO.Database
Private mTraceLogWriteRecordset As DAO.Recordset
Private mTraceLogWriteRecordsetOwnerRunID As String
Private mTraceTableWriteSuspensionDepth As Long
Private mTraceStorageVerificationCount As Long
Private mTraceWriteRecordsetOpenCount As Long
Private mTraceRowsWrittenCount As Long
Private mTracePhysicalLogOpenCount As Long
Private mTracePhysicalLogLinesWrittenCount As Long
Private mTraceLogFileMaxLineLengthForRun As Long

' TABLES-local setup-control pump instrumentation.
Private mSetupControlPumpRequestsIssuedCount As Long

' Setup preflight backup and schema-repair state.
Private mSetupPreflightBackupDecisionReady As Boolean
Private mSetupPreflightSchemaRepairNeeded As Boolean
Private mSetupPreflightBackupCount As Long
Private mSetupPreflightA2SchemaEvidenceReady As Boolean
Private mSetupPreflightA2SchemaRepairNeeded As Boolean
Private mSetupPreflightA2SchemaRepairSummary As String
Private mSetupPreflightA2TraceRunID As String
Private mSetupPreflightA2HostDbName As String
Private mSetupPreflightA2SettingsSchemaDefinitions As Collection
Private mSetupPreflightA2SettingsIndexDefinitions As Collection
Private mSetupPreflightA2ConfigSchemaDefinitions As Collection
Private mSetupPreflightA2ConfigIndexDefinitions As Collection
Private mSetupPreflightA2MacroCategorySchemaDefinitions As Collection
Private mSetupPreflightA2MacroSubcategorySchemaDefinitions As Collection
Private mSetupPreflightA2EditableLookupValueTextSizes As Object
' INTENTIONAL OPERATION-SCOPED DAO METADATA SNAPSHOT:
' PrepareSetupDataTableBackupForPreflight owns this transient, case-insensitive index while it plans A2/B1 schema work and resolves the exact existing tables eligible for backup. A single TableDefs enumeration
' replaces dozens of independent missing-table probes. The dictionary contains existence evidence only; canonical table membership and schema definitions remain owned by DEFAULTS/TABLES. It is cleared before any
' backup copy is created and on every error path, so later setup phases cannot consume stale metadata.
Private mSetupPreflightTableNameIndexActive As Boolean
Private mSetupPreflightExistingTableNames As Object
' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION:
' Maps preflight-approved affected table names to their backup identities so later schema phases can prove backup coverage before mutation. It is valid only for the current Setup preflight decision and is cleared by
' the existing preflight reset path; it records safety proof, not table authority.
Private mSetupPreflightBackedTableNames As Object

' Runtime table-maintenance declarations.
Private Const TABLE_MAINTENANCE_MESSAGE_TITLE As String = "Access Development Console Table Maintenance"

' ===============================================================================================================================================================================
' PROCEDURES
' ===============================================================================================================================================================================

' ===============================================================================================================================================================================
' PREFLIGHT / BACKUP / RELATIONSHIP FOUNDATION PROCEDURES
' ===============================================================================================================================================================================

' Purpose: Ensures u sys ribbons table schema.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub EnsureUSysRibbonsTableSchema()
    'A2 owns the complete physical table foundation; Ribbon row/XML content remains H2-owned.
    If SetupDomainMutationWritesAllowedForCurrentPart(SETUP_DOMAIN_A2_PHYSICAL_SCHEMA) = False Then Err.Raise vbObjectError + 9630, "EnsureUSysRibbonsTableSchema", "USysRibbons physical schema mutation is owned by A2 during numbered Setup. Successor phases must verify the A2 schema read-only."
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim fld As DAO.Field
    Set db = TablesWorkspaceCodeDb()
    If SharedDao_TableExists(db, DB_TABLE_USYS_RIBBONS) = False Then
        Set tdf = db.CreateTableDef(DB_TABLE_USYS_RIBBONS)
        Set fld = tdf.CreateField(FIELD_RIBBON_NAME, dbText, 255)
        fld.Required = True
        fld.allowZeroLength = False
        tdf.fields.Append fld
        Set fld = tdf.CreateField(FIELD_RIBBON_XML, dbMemo)
        fld.Required = True
        fld.allowZeroLength = False
        tdf.fields.Append fld
        db.TableDefs.Append tdf
    Else
        Set tdf = db.TableDefs(DB_TABLE_USYS_RIBBONS)
        If SharedDao_TableDefFieldExistsSafe(tdf, FIELD_RIBBON_NAME) = False Then tdf.fields.Append tdf.CreateField(FIELD_RIBBON_NAME, dbText, 255)
        If SharedDao_TableDefFieldExistsSafe(tdf, FIELD_RIBBON_XML) = False Then tdf.fields.Append tdf.CreateField(FIELD_RIBBON_XML, dbMemo)
        tdf.fields.Refresh
        Set fld = tdf.fields(FIELD_RIBBON_NAME)
        If SharedDao_FieldDefinitionMatches(fld, dbText, 255, False) = False Then Err.Raise vbObjectError + 5348, "EnsureUSysRibbonsTableSchema", "USysRibbons.RibbonName has an incompatible type or size."
        If fld.Required = False Then fld.Required = True
        If fld.allowZeroLength Then fld.allowZeroLength = False
        Set fld = tdf.fields(FIELD_RIBBON_XML)
        If fld.Type <> dbMemo Then Err.Raise vbObjectError + 5349, "EnsureUSysRibbonsTableSchema", "USysRibbons.RibbonXml has an incompatible type."
        If fld.Required = False Then fld.Required = True
        If fld.allowZeroLength Then fld.allowZeroLength = False
    End If
    db.TableDefs.Refresh
    EnsureSingleFieldIndex db, DB_TABLE_USYS_RIBBONS, INDEX_UX_U_SYS_RIBBONS_RIBBON_NAME, FIELD_RIBBON_NAME, True, False, mTableDesignBatchPrepared
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "EnsureUSysRibbonsTableSchema", "Could not create or repair the complete A2 USysRibbons schema." & vbCrLf & "Err.Description: " & Err.Description
End Sub

' Purpose: Checks whether u sys ribbons table schema already satisfies the exact state required by the current Setup run so the owning phase can determine whether repair is necessary.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   detailText (Output Optional; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller. Default: vbNullString.
' Returns: Boolean: True when u sys ribbons table schema matches the exact state required by the current Setup contract; False when the owning phase still needs repair, refresh, or reconciliation.
Public Function USysRibbonsTableSchemaCurrentForSetup(Optional ByRef detailText As String = vbNullString) As Boolean
    'Read-only exact A2 predecessor verification for H2. This routine never creates, repairs, refreshes-for-mutation, or normalizes USysRibbons.
    On Error GoTo NotCurrent
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim fld As DAO.Field
    detailText = vbNullString
    Set db = TablesWorkspaceCodeDb()
    If SharedDao_TableExists(db, DB_TABLE_USYS_RIBBONS) = False Then detailText = "A2-owned USysRibbons table is missing.": Exit Function
    Set tdf = db.TableDefs(DB_TABLE_USYS_RIBBONS)
    If SharedDao_TableDefFieldExistsSafe(tdf, FIELD_RIBBON_NAME) = False Then detailText = "A2-owned USysRibbons.RibbonName field is missing.": Exit Function
    If SharedDao_TableDefFieldExistsSafe(tdf, FIELD_RIBBON_XML) = False Then detailText = "A2-owned USysRibbons.RibbonXml field is missing.": Exit Function
    Set fld = tdf.fields(FIELD_RIBBON_NAME)
    If SharedDao_FieldDefinitionMatches(fld, dbText, 255, False) = False Then detailText = "A2-owned USysRibbons.RibbonName type or size is stale.": Exit Function
    If fld.Required = False Then detailText = "A2-owned USysRibbons.RibbonName.Required is stale.": Exit Function
    If fld.allowZeroLength Then detailText = "A2-owned USysRibbons.RibbonName.AllowZeroLength is stale.": Exit Function
    Set fld = tdf.fields(FIELD_RIBBON_XML)
    If fld.Type <> dbMemo Then detailText = "A2-owned USysRibbons.RibbonXml type is stale.": Exit Function
    If fld.Required = False Then detailText = "A2-owned USysRibbons.RibbonXml.Required is stale.": Exit Function
    If fld.allowZeroLength Then detailText = "A2-owned USysRibbons.RibbonXml.AllowZeroLength is stale.": Exit Function
    If SharedDao_TableDefSingleFieldIndexMatches(tdf, INDEX_UX_U_SYS_RIBBONS_RIBBON_NAME, FIELD_RIBBON_NAME, True, False) = False Then detailText = "A2-owned USysRibbons unique RibbonName index is missing or stale.": Exit Function
    detailText = "A2-owned USysRibbons physical schema is current."
    USysRibbonsTableSchemaCurrentForSetup = True
    Exit Function
NotCurrent:
    detailText = "Could not verify the A2-owned USysRibbons physical schema read-only." & vbCrLf & "Err.Number: " & CStr(Err.Number) & vbCrLf & "Err.Source: " & Err.Source & vbCrLf & "Err.Description: " & Err.Description
End Function

Private Function SetupPerformanceDiagnosticsSummary() As String
    'Returns run-scoped display and trace-writer evidence for seeded performance comparison.
    Dim uiSummary As String
    On Error Resume Next
    uiSummary = SetupControlPerformanceSummary()
    On Error GoTo 0
    SetupPerformanceDiagnosticsSummary = _
        CoreProgressPerformanceDiagnosticsSummary() & vbCrLf & _
        "Setup-controller pump requests issued by TABLES: " & CStr(mSetupControlPumpRequestsIssuedCount) & vbCrLf & _
        "TraceLog structural readiness checks: " & CStr(mTraceStorageVerificationCount) & vbCrLf & _
        "TraceLog dynaset opens: " & CStr(mTraceWriteRecordsetOpenCount) & vbCrLf & _
        "Trace rows written before this completion event: " & CStr(mTraceRowsWrittenCount) & vbCrLf & _
        "Physical trace-log opens: " & CStr(mTracePhysicalLogOpenCount) & vbCrLf & _
        "Physical trace lines written before this completion event: " & CStr(mTracePhysicalLogLinesWrittenCount)
    If Len(Trim$(uiSummary)) > 0 Then SetupPerformanceDiagnosticsSummary = SetupPerformanceDiagnosticsSummary & vbCrLf & uiSummary
    SetupPerformanceDiagnosticsSummary = SetupPerformanceDiagnosticsSummary & vbCrLf & NamingPerformanceSummary()
End Function

Private Sub ClearSetupPreflightA2SchemaEvidence()
    'TABLES-owned transient A2 pre-mutation evidence is valid only until A2 mutates the physical schema or the Preflight operation fails.
    mSetupPreflightA2SchemaEvidenceReady = False
    mSetupPreflightA2SchemaRepairNeeded = False
    mSetupPreflightA2SchemaRepairSummary = vbNullString
    mSetupPreflightA2TraceRunID = vbNullString
    mSetupPreflightA2HostDbName = vbNullString
    Set mSetupPreflightA2SettingsSchemaDefinitions = Nothing
    Set mSetupPreflightA2SettingsIndexDefinitions = Nothing
    Set mSetupPreflightA2ConfigSchemaDefinitions = Nothing
    Set mSetupPreflightA2ConfigIndexDefinitions = Nothing
    Set mSetupPreflightA2MacroCategorySchemaDefinitions = Nothing
    Set mSetupPreflightA2MacroSubcategorySchemaDefinitions = Nothing
    Set mSetupPreflightA2EditableLookupValueTextSizes = Nothing
End Sub

Private Sub ResetSetupPreflightInfrastructureState()
    'A new top-level or standalone Setup run cannot inherit table-existence evidence from an interrupted Preflight assessment.
    EndSetupPreflightTableNameIndex
    mSetupPreflightBackupDecisionReady = False
    mSetupPreflightSchemaRepairNeeded = False
    mSetupPreflightBackupCount = 0
    ClearSetupPreflightA2SchemaEvidence
    Set mSetupPreflightBackedTableNames = Nothing
End Sub

Private Function SetupPreflightA2EvidenceCurrent() As Boolean
    Dim db As DAO.Database
    If mSetupPreflightA2SchemaEvidenceReady = False Then Exit Function
    If Len(mSetupPreflightA2TraceRunID) = 0 Or StrComp(GetActiveTraceRunID(), mSetupPreflightA2TraceRunID, vbBinaryCompare) <> 0 Then Exit Function
    Set db = TablesWorkspaceCodeDb()
    If StrComp(db.Name, mSetupPreflightA2HostDbName, vbTextCompare) <> 0 Then Exit Function
    SetupPreflightA2EvidenceCurrent = True
End Function

' Purpose: Attempts to perform get setup preflight a2 schema evidence.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   repairNeeded (Output; Boolean): Set True when the procedure determines that repair needed applies; set False when that condition is not met.
'   repairSummary (Output; String): Receives the resolved repair summary produced from this procedure's validated inputs and current ADC state.
' Returns: Boolean: True when the attempt to get setup preflight a2 schema evidence completes successfully; False for the routine's defined normal unsuccessful/not-applicable outcome.
Public Function TryGetSetupPreflightA2SchemaEvidence(ByRef repairNeeded As Boolean, ByRef repairSummary As String) As Boolean
    repairNeeded = False
    repairSummary = vbNullString
    If SetupPreflightA2EvidenceCurrent() = False Then Exit Function
    repairNeeded = mSetupPreflightA2SchemaRepairNeeded
    repairSummary = mSetupPreflightA2SchemaRepairSummary
    TryGetSetupPreflightA2SchemaEvidence = True
End Function

' Purpose: Ensures setup diagnostics ready for preflight.
' Public role: Public module entry point. This procedure is listed in the module header as an intentional externally callable ADC entry point and may be invoked directly by Access/VBA callers.
' Parameters:
'   detailText (Output Optional; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller. Default: vbNullString.
' Returns: None (Sub). Results are returned through the documented ByRef parameters and/or the routine's database, UI, or module-state side effects.
Public Sub EnsureSetupDiagnosticsReadyForPreflight(Optional ByRef detailText As String = vbNullString)
    'Unnumbered Setup Preflight owns diagnostics table/schema/index readiness. B5 owns all canonical relationship creation and repair.
    On Error GoTo ErrHandler
    Const STAGE_SNAPSHOT As String = "Snapshot existing diagnostics state"
    Const STAGE_TRACE_TABLE As String = "Repair TraceLog schema and indexes"
    Const STAGE_SCREENSHOT_TABLE As String = "Repair TraceLog_Screenshots schema and indexes"
    Dim db As DAO.Database
    Dim traceExisted As Boolean
    Dim screenshotsExisted As Boolean
    Dim currentStage As String
    Dim stageStarted As Single
    Dim stageTimerActive As Boolean
    Dim snapshotElapsed As Single
    Dim traceTableElapsed As Single
    Dim screenshotTableElapsed As Single
    currentStage = STAGE_SNAPSHOT
    stageStarted = Timer
    stageTimerActive = True
    CloseRetainedTraceWriteRecordset
    Set db = TablesWorkspaceCodeDb()
    traceExisted = SharedDao_TableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_TraceLog"))
    screenshotsExisted = SharedDao_TableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_TraceLogScreenshots"))
    snapshotElapsed = ElapsedTimerSeconds(stageStarted)
    stageTimerActive = False
    PumpSetupResponsiveness True
    currentStage = STAGE_TRACE_TABLE
    stageStarted = Timer
    stageTimerActive = True
    EnsureTraceLogWriteTable db
    traceTableElapsed = ElapsedTimerSeconds(stageStarted)
    stageTimerActive = False
    PumpSetupResponsiveness True
    currentStage = STAGE_SCREENSHOT_TABLE
    stageStarted = Timer
    stageTimerActive = True
    EnsureTraceScreenshotWriteTable db
    screenshotTableElapsed = ElapsedTimerSeconds(stageStarted)
    stageTimerActive = False
    detailText = "Setup Preflight diagnostics storage ready." & vbCrLf & _
                 "TraceLog: " & IIf(traceExisted, "verified/repaired", "created") & vbCrLf & _
                 "TraceLog_Screenshots: " & IIf(screenshotsExisted, "verified/repaired", "created") & vbCrLf & _
                 "Diagnostics state snapshot elapsed: " & Format$(snapshotElapsed, "0.0") & " seconds" & vbCrLf & _
                 "TraceLog schema/index elapsed: " & Format$(traceTableElapsed, "0.0") & " seconds" & vbCrLf & _
                 "TraceLog_Screenshots schema/index elapsed: " & Format$(screenshotTableElapsed, "0.0") & " seconds" & vbCrLf & _
                 "B5 owns canonical relationship creation and repair."
    Set db = Nothing
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If stageTimerActive Then
        Select Case currentStage
            Case STAGE_SNAPSHOT
                snapshotElapsed = ElapsedTimerSeconds(stageStarted)
            Case STAGE_TRACE_TABLE
                traceTableElapsed = ElapsedTimerSeconds(stageStarted)
            Case STAGE_SCREENSHOT_TABLE
                screenshotTableElapsed = ElapsedTimerSeconds(stageStarted)
        End Select
        stageTimerActive = False
    End If
    Set db = Nothing
    detailText = "Setup Preflight diagnostics readiness failed." & vbCrLf & _
                 "Failure stage: " & IIf(Len(currentStage) = 0, "(not recorded)", currentStage) & vbCrLf & _
                 "Diagnostics state snapshot elapsed: " & Format$(snapshotElapsed, "0.0") & " seconds" & vbCrLf & _
                 "TraceLog schema/index elapsed: " & Format$(traceTableElapsed, "0.0") & " seconds" & vbCrLf & _
                 "TraceLog_Screenshots schema/index elapsed: " & Format$(screenshotTableElapsed, "0.0") & " seconds"
    Err.Raise savedErrNumber, savedErrSource, detailText & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

' Purpose: Resolves tables configured physical object names for setup from the supplied/current ADC state, using `Next` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: Collection: Returns the ordered Collection assembled by the routine; callers receive the object reference and should treat its membership/order according to the routine contract.
Public Function TablesConfiguredPhysicalObjectNamesForSetup() As Collection
    Dim names As Collection
    Dim item As Variant
    ' INTENTIONAL FILTERED PARTITION: names is a consumer-required subset/projection derived from the source enumerated by TablesConfiguredPhysicalObjectNamesForSetup; its membership is not independently
    ' authoritative.
    Set names = New Collection
    ' INTENTIONAL SEPARATE CANONICAL PROJECTIONS: table and non-form query names are independently reusable authorities; keep them separate rather than adding a setup-only multi-output resolver for one small
    ' registry traversal.
    For Each item In FoundationRequiredTableNames()
        names.Add "TABLE|" & CStr(item)
    Next item
    For Each item In FoundationRequiredNonFormQueryNames()
        names.Add "QUERY|" & CStr(item)
    Next item
    Set TablesConfiguredPhysicalObjectNamesForSetup = names
End Function

' Purpose: Prepares setup data table backup for preflight.
' Public role: Public module entry point. This procedure is listed in the module header as an intentional externally callable ADC entry point and may be invoked directly by Access/VBA callers.
' Parameters:
'   detailText (Output Optional; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller. Default: vbNullString.
' Returns: None (Sub). Results are returned through the documented ByRef parameters and/or the routine's database, UI, or module-state side effects.
Public Sub PrepareSetupDataTableBackupForPreflight(Optional ByRef detailText As String = vbNullString)
    'Preflight decides once whether A2 needs backups; A2 never creates them itself. The read-only schema assessment uses one operation-scoped TableDefs name index so a fresh database does not pay for a separate DAO
    'existence probe for every managed lookup table.
    On Error GoTo ErrHandler
    Dim schemaRepairSummary As String
    Dim schemaRepairNeeded As Boolean
    Dim backupStamp As String
    Dim tableName As Variant
    Dim lookupTableName As Variant
    Dim affectedExistingCount As Long
    Dim backupName As String
    Dim tableNames As Collection
    Dim affectedExistingTableNames As Collection
    Dim terminologyRepairSummary As String
    Dim terminologyRepairNeeded As Boolean
    Dim settingsSchemaDefinitions As Collection
    Dim settingsIndexDefinitions As Collection
    Dim configSchemaDefinitions As Collection
    Dim configIndexDefinitions As Collection
    Dim macroCategorySchemaDefinitions As Collection
    Dim macroSubcategorySchemaDefinitions As Collection
    Dim editableLookupValueTextSizes As Object
    Dim metadataSnapshotStarted As Single
    Dim metadataSnapshotElapsed As Single
    Dim metadataSnapshotTableCount As Long
    If mSetupPreflightBackupDecisionReady Then
        detailText = "Setup Preflight A2 backup decision reused." & vbCrLf & _
                     "Schema repair required: " & CStr(mSetupPreflightSchemaRepairNeeded) & vbCrLf & _
                     "Tables backed up: " & CStr(mSetupPreflightBackupCount)
        Exit Sub
    End If
    Set mSetupPreflightBackedTableNames = CreateObject("Scripting.Dictionary")
    mSetupPreflightBackedTableNames.CompareMode = vbTextCompare
    Set settingsSchemaDefinitions = SettingsFieldSchemaDefinitions()
    Set settingsIndexDefinitions = SettingsIndexSchemaDefinitions()
    Set configSchemaDefinitions = ConfigFieldSchemaDefinitions()
    Set configIndexDefinitions = ConfigIndexSchemaDefinitions()
    Set macroCategorySchemaDefinitions = BuildMacroCategorySchemaDefinitions()
    Set macroSubcategorySchemaDefinitions = BuildMacroSubcategorySchemaDefinitions()
    Set editableLookupValueTextSizes = EditableLookupValueTextSizeIndex(configSchemaDefinitions, BuildHealthReviewSchemaDefinitions(), CustomMacroLauncherReviewSchemaDefinitions())
    CorePulseProgress "Setup Preflight: snapshot existing table names for schema and backup planning."
    metadataSnapshotStarted = Timer
    BeginSetupPreflightTableNameIndex
    schemaRepairNeeded = TableSchemaRepairRequired(schemaRepairSummary, settingsSchemaDefinitions, settingsIndexDefinitions, configSchemaDefinitions, configIndexDefinitions, macroCategorySchemaDefinitions, macroSubcategorySchemaDefinitions, editableLookupValueTextSizes)
    mSetupPreflightA2SchemaRepairNeeded = schemaRepairNeeded
    mSetupPreflightA2SchemaRepairSummary = schemaRepairSummary
    mSetupPreflightA2TraceRunID = GetActiveTraceRunID()
    mSetupPreflightA2HostDbName = TablesWorkspaceCodeDb().Name
    Set mSetupPreflightA2SettingsSchemaDefinitions = settingsSchemaDefinitions
    Set mSetupPreflightA2SettingsIndexDefinitions = settingsIndexDefinitions
    Set mSetupPreflightA2ConfigSchemaDefinitions = configSchemaDefinitions
    Set mSetupPreflightA2ConfigIndexDefinitions = configIndexDefinitions
    Set mSetupPreflightA2MacroCategorySchemaDefinitions = macroCategorySchemaDefinitions
    Set mSetupPreflightA2MacroSubcategorySchemaDefinitions = macroSubcategorySchemaDefinitions
    Set mSetupPreflightA2EditableLookupValueTextSizes = editableLookupValueTextSizes
    mSetupPreflightA2SchemaEvidenceReady = True
    terminologyRepairNeeded = TerminologySchemaRepairRequired(terminologyRepairSummary)
    mSetupPreflightSchemaRepairNeeded = (schemaRepairNeeded Or terminologyRepairNeeded)
    If Len(terminologyRepairSummary) > 0 Then schemaRepairSummary = SharedText_AppendNonBlankLine(schemaRepairSummary, terminologyRepairSummary)
    schemaRepairNeeded = (schemaRepairNeeded Or terminologyRepairNeeded)
    Set tableNames = New Collection
    tableNames.Add ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")
    tableNames.Add ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates")
    tableNames.Add ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReview")
    tableNames.Add ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LauncherControlsReview")
    tableNames.Add ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID)
    tableNames.Add ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_MacroCategories")
    tableNames.Add ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_MacroSubcategories")
    tableNames.Add ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LookupDefinitions")
    tableNames.Add ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_TERMINOLOGY_TABLE_LOGICAL_ID)
    For Each lookupTableName In EditableLookupTableNames()
        tableNames.Add CStr(lookupTableName)
    Next lookupTableName
    Set affectedExistingTableNames = New Collection
    If schemaRepairNeeded Then
        For Each tableName In tableNames
            If RepairSummaryContainsTable(schemaRepairSummary, CStr(tableName)) And CodeDbTableExists(CStr(tableName)) Then affectedExistingTableNames.Add CStr(tableName)
        Next tableName
    End If
    affectedExistingCount = affectedExistingTableNames.count
    metadataSnapshotElapsed = ElapsedTimerSeconds(metadataSnapshotStarted)
    metadataSnapshotTableCount = mSetupPreflightExistingTableNames.count
    EndSetupPreflightTableNameIndex
    CorePulseProgress "Setup Preflight: schema and backup assessment completed."
    If affectedExistingCount > 0 Then
        backupStamp = Format$(Now, "yyyymmdd_hhnnss")
        For Each tableName In affectedExistingTableNames
            If TableCanOpen(CStr(tableName)) = False Then Err.Raise vbObjectError + 9203, "PrepareSetupDataTableBackupForPreflight", "An affected existing table cannot be opened, so Preflight cannot create the required backup." & vbCrLf & "Table: " & CStr(tableName) & vbCrLf & TableOpenFailureDetail(CStr(tableName))
            backupName = BackupOneTable(CStr(tableName), backupStamp)
            CoreAddOperationNote "Backup table [" & backupName & "] was created without Navigation Pane placement; physical placement belongs to Phase I."
            mSetupPreflightBackedTableNames.Add CStr(tableName), backupName
            mSetupPreflightBackupCount = mSetupPreflightBackupCount + 1
        Next tableName
    End If
    mSetupPreflightBackupDecisionReady = True
    If schemaRepairNeeded = False Then
        detailText = "Setup Preflight determined that A2 and B1 Help_Terms require no schema repair; no backup was created."
    ElseIf affectedExistingCount = 0 Then
        detailText = "Setup Preflight found only missing affected tables (unseeded/missing-only state); no empty or redundant backup was created." & vbCrLf & schemaRepairSummary
    Else
        detailText = "Setup Preflight backed up only existing A2 or B1 Help_Terms tables affected by planned schema repair, including editable lookup tables when their persisted schema requires repair." & vbCrLf & _
                     "Affected existing tables: " & CStr(affectedExistingCount) & vbCrLf & _
                     "Tables backed up: " & CStr(mSetupPreflightBackupCount) & vbCrLf & _
                     schemaRepairSummary
    End If
    detailText = "TableDefs enumerations for schema assessment: 1" & vbCrLf & _
                 "Existing table names indexed: " & CStr(metadataSnapshotTableCount) & vbCrLf & _
                 "Metadata snapshot, schema comparison, and backup-candidate resolution elapsed: " & Format$(metadataSnapshotElapsed, "0.0") & " seconds" & vbCrLf & _
                 detailText
    LogTraceMessage TRACE_LEVEL_INFO, "Backup", "PrepareSetupDataTableBackupForPreflight", "Setup Preflight completed the A2 backup decision.", detailText, "Setup Preflight backup decision", 0, vbNullString, vbNullString, vbNullString, "TableBackupPlan", "CodeDb", "Setup"
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    EndSetupPreflightTableNameIndex
    mSetupPreflightBackupDecisionReady = False
    ClearSetupPreflightA2SchemaEvidence
    detailText = "Setup Preflight could not establish the required A2 backup decision."
    Err.Raise savedErrNumber, "PrepareSetupDataTableBackupForPreflight", detailText & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub BeginSetupPreflightTableNameIndex()
    'Builds the sole table-existence snapshot used by the current Preflight assessment. TableDefs is enumerated exactly once; dictionary lookup is then O(1) for core, lookup, registry, and Help_Terms checks nested
    'beneath TableSchemaRepairRequired.
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    EndSetupPreflightTableNameIndex
    Set db = TablesWorkspaceCodeDb()
    Set mSetupPreflightExistingTableNames = CreateObject("Scripting.Dictionary")
    mSetupPreflightExistingTableNames.CompareMode = vbTextCompare
    For Each tdf In db.TableDefs
        If Len(Trim$(tdf.Name)) > 0 Then mSetupPreflightExistingTableNames(tdf.Name) = True
    Next tdf
    mSetupPreflightTableNameIndexActive = True
End Sub

Private Sub EndSetupPreflightTableNameIndex()
    'Invalidates operation-scoped existence evidence before any structural mutation or backup copy can change TableDefs.
    mSetupPreflightTableNameIndexActive = False
    Set mSetupPreflightExistingTableNames = Nothing
End Sub

Private Function ElapsedTimerSeconds(ByVal startedAt As Single) As Single
    'Timer resets at midnight. Normalize the one possible rollover so diagnostics remain accurate for a Setup run that crosses midnight.
    ElapsedTimerSeconds = Timer - startedAt
    If ElapsedTimerSeconds < 0 Then ElapsedTimerSeconds = ElapsedTimerSeconds + 86400!
End Function

Private Sub RequireSetupPreflightBackupCoverage(ByVal schemaRepairNeeded As Boolean, ByVal schemaRepairSummary As String)
    Dim tableName As Variant
    Dim lookupTableName As Variant
    Dim tableNames As Collection
    If mSetupPreflightBackupDecisionReady = False Then Err.Raise vbObjectError + 9204, "RequireSetupPreflightBackupCoverage", "A2 cannot begin because unnumbered Setup Preflight did not establish its backup decision."
    If schemaRepairNeeded = False Then Exit Sub
    Set tableNames = New Collection
    tableNames.Add ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")
    tableNames.Add ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReview")
    tableNames.Add ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LauncherControlsReview")
    tableNames.Add ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID)
    tableNames.Add ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_MacroCategories")
    tableNames.Add ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_MacroSubcategories")
    tableNames.Add ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LookupDefinitions")
    For Each lookupTableName In EditableLookupTableNames()
        tableNames.Add CStr(lookupTableName)
    Next lookupTableName
    For Each tableName In tableNames
        If RepairSummaryContainsTable(schemaRepairSummary, CStr(tableName)) And CodeDbTableExists(CStr(tableName)) Then
            If mSetupPreflightBackedTableNames Is Nothing Then Err.Raise vbObjectError + 9205, "RequireSetupPreflightBackupCoverage", "A2 found an affected existing table, but Preflight recorded no backup coverage." & vbCrLf & "Table: " & CStr(tableName)
            If mSetupPreflightBackedTableNames.Exists(CStr(tableName)) = False Then Err.Raise vbObjectError + 9206, "RequireSetupPreflightBackupCoverage", "A2 found an affected existing table that was not included in the Preflight backup decision." & vbCrLf & "Table: " & CStr(tableName) & vbCrLf & "Rerun Setup so Preflight can reassess the current schema."
        End If
    Next tableName
End Sub

' Purpose: Requires setup preflight terminology backup coverage for supplied schema definitions, supplied index definitions.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   suppliedSchemaDefinitions (Input Optional; Variant): Caller-supplied supplied schema definitions data used as the authoritative definition input for this call, avoiding an unnecessary reload when already
'   available.
'   suppliedIndexDefinitions (Input Optional; Variant): Caller-supplied supplied index definitions data used as the authoritative definition input for this call, avoiding an unnecessary reload when already
'   available.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub RequireSetupPreflightTerminologyBackupCoverage(Optional ByVal suppliedSchemaDefinitions As Variant, Optional ByVal suppliedIndexDefinitions As Variant)
    'B1 calls this before changing an existing Help_Terms schema; missing/unseeded Help_Terms requires no backup.
    Dim repairSummary As String
    Dim repairNeeded As Boolean
    Dim schemaDefinitions As Collection
    Dim indexDefinitions As Collection
    If isMissing(suppliedSchemaDefinitions) = False Then
        If IsObject(suppliedSchemaDefinitions) Then Set schemaDefinitions = suppliedSchemaDefinitions
    End If
    If isMissing(suppliedIndexDefinitions) = False Then
        If IsObject(suppliedIndexDefinitions) Then Set indexDefinitions = suppliedIndexDefinitions
    End If
    If schemaDefinitions Is Nothing Then Set schemaDefinitions = TerminologyFieldSchemaDefinitions()
    If indexDefinitions Is Nothing Then Set indexDefinitions = TerminologyIndexSchemaDefinitions()
    If mSetupPreflightBackupDecisionReady = False Then Err.Raise vbObjectError + 9207, "RequireSetupPreflightTerminologyBackupCoverage", "B1 cannot begin Help_Terms structural repair because Setup Preflight did not establish backup coverage."
    repairNeeded = TerminologySchemaRepairRequired(repairSummary, schemaDefinitions, indexDefinitions)
    If repairNeeded = False Then Exit Sub
    If CodeDbTableExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_TERMINOLOGY_TABLE_LOGICAL_ID)) = False Then Exit Sub
    If mSetupPreflightBackedTableNames Is Nothing Then Err.Raise vbObjectError + 9208, "RequireSetupPreflightTerminologyBackupCoverage", "B1 found an affected existing Help_Terms table, but Preflight recorded no backup coverage." & vbCrLf & "Table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_TERMINOLOGY_TABLE_LOGICAL_ID)
    If mSetupPreflightBackedTableNames.Exists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_TERMINOLOGY_TABLE_LOGICAL_ID)) = False Then Err.Raise vbObjectError + 9209, "RequireSetupPreflightTerminologyBackupCoverage", "B1 found an affected existing Help_Terms table that was not included in the Preflight backup decision." & vbCrLf & "Table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_TERMINOLOGY_TABLE_LOGICAL_ID) & vbCrLf & "Rerun Setup so Preflight can reassess the current schema."
End Sub

Private Function CreateBackupTableCopy(ByVal sourceTableName As String, ByVal backupStamp As String) As String
    'Sole physical authority for choosing an ADC backup-table name and creating the table copy. Callers retain workflow-specific notes and Navigation Pane handling.
    On Error GoTo ErrHandler
    Dim backupTableName As String
    backupTableName = UniqueBackupTableName(sourceTableName, backupStamp)
    TablesWorkspaceCodeDb().Execute "SELECT * INTO [" & backupTableName & "] FROM [" & sourceTableName & "]", dbFailOnError
    ApplyBackupTableDescriptions TablesWorkspaceCodeDb(), sourceTableName, backupTableName, backupStamp
    CreateBackupTableCopy = backupTableName
    Exit Function
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrSource As String
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrSource = Err.Source
    originalErrDescription = Err.Description
    Dim cleanupErrNumber As Long
    Dim cleanupErrSource As String
    Dim cleanupErrDescription As String
    On Error Resume Next
    Err.Clear
    DeleteIncompleteBackupTableVerified TablesWorkspaceCodeDb(), backupTableName
    cleanupErrNumber = Err.Number
    cleanupErrSource = Err.Source
    cleanupErrDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    Err.Raise originalErrNumber, "CreateBackupTableCopy", "Could not create and fully describe an Access Development Console backup-table copy." & vbCrLf & "Source table: " & sourceTableName & vbCrLf & "Incomplete backup cleanup: " & IIf(cleanupErrNumber = 0, "deleted and verified", "FAILED") & IIf(cleanupErrNumber = 0, vbNullString, vbCrLf & "Cleanup Err.Number: " & CStr(cleanupErrNumber) & vbCrLf & "Cleanup Err.Source: " & cleanupErrSource & vbCrLf & "Cleanup Err.Description: " & cleanupErrDescription) & vbCrLf & "Original Err.Source: " & originalErrSource & vbCrLf & "Original Err.Description: " & originalErrDescription
End Function

Private Sub DeleteIncompleteBackupTableVerified(ByVal db As DAO.Database, ByVal backupTableName As String)
    'Failure cleanup is successful only when the incomplete backup is proven absent; inspection failures are never treated as absence.
    On Error GoTo ErrHandler
    Dim tdf As DAO.tableDef
    Dim lookupErrNumber As Long
    Dim lookupErrSource As String
    Dim lookupErrDescription As String
    backupTableName = Trim$(backupTableName)
    If Len(backupTableName) = 0 Then Exit Sub
    db.TableDefs.Refresh
    On Error Resume Next
    Err.Clear
    Set tdf = db.TableDefs(backupTableName)
    lookupErrNumber = Err.Number
    lookupErrSource = Err.Source
    lookupErrDescription = Err.Description
    Err.Clear
    On Error GoTo ErrHandler
    If lookupErrNumber = 3265 Then Exit Sub
    If lookupErrNumber <> 0 Then Err.Raise lookupErrNumber, lookupErrSource, lookupErrDescription
    Set tdf = Nothing
    db.TableDefs.Delete backupTableName
    db.TableDefs.Refresh
    On Error Resume Next
    Err.Clear
    Set tdf = db.TableDefs(backupTableName)
    lookupErrNumber = Err.Number
    lookupErrSource = Err.Source
    lookupErrDescription = Err.Description
    Err.Clear
    On Error GoTo ErrHandler
    If lookupErrNumber = 3265 Then Exit Sub
    If lookupErrNumber <> 0 Then Err.Raise lookupErrNumber, lookupErrSource, lookupErrDescription
    Err.Raise vbObjectError + 9399, "DeleteIncompleteBackupTableVerified", "The incomplete backup table still exists after deletion." & vbCrLf & "Table: " & backupTableName
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "DeleteIncompleteBackupTableVerified", "Could not delete and verify removal of the incomplete ADC backup table [" & backupTableName & "]." & vbCrLf & savedErrDescription
End Sub

Private Sub ApplyBackupTableDescriptions(ByVal db As DAO.Database, ByVal sourceTableName As String, ByVal backupTableName As String, ByVal backupStamp As String)
    'Backup tables are intentionally outside the Managed Object Catalog, so the creating authority owns their complete descriptive metadata.
    On Error GoTo ErrHandler
    Dim sourceTdf As DAO.tableDef
    Dim backupTdf As DAO.tableDef
    Dim backupField As DAO.Field
    Dim sourceDescription As String
    Dim fieldDescription As String
    db.TableDefs.Refresh
    Set sourceTdf = db.TableDefs(sourceTableName)
    Set backupTdf = db.TableDefs(backupTableName)
    Call EnsureDaoPropertyValue(backupTdf, DAO_DESCRIPTION_PROPERTY_NAME, dbText, Left$("ADC recovery backup of table [" & sourceTableName & "] created before schema or data repair. Snapshot stamp: " & backupStamp & ". Preserve it until the associated Setup/recovery operation is verified; it is evidence, not runtime authority.", 255), "ApplyBackupTableDescriptions")
    For Each backupField In backupTdf.fields
        sourceDescription = vbNullString
        On Error Resume Next
        sourceDescription = CStr(sourceTdf.fields(backupField.Name).properties(DAO_DESCRIPTION_PROPERTY_NAME).value)
        Err.Clear
        On Error GoTo ErrHandler
        fieldDescription = "Backup snapshot of [" & sourceTableName & "].[" & backupField.Name & "]. Retains the value captured before repair and must not be treated as current runtime data."
        If Len(Trim$(sourceDescription)) > 0 Then fieldDescription = fieldDescription & " Source meaning: " & sourceDescription
        Call EnsureDaoPropertyValue(backupField, DAO_DESCRIPTION_PROPERTY_NAME, dbText, Left$(fieldDescription, 255), "ApplyBackupTableDescriptions")
    Next backupField
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ApplyBackupTableDescriptions", "Could not apply and verify developer descriptions on ADC backup table [" & backupTableName & "] copied from [" & sourceTableName & "]." & vbCrLf & savedErrDescription
End Sub

Private Function BackupOneTable(ByVal sourceTableName As String, ByVal backupStamp As String) As String
    ' Wrapper purpose: Workflow-neutral backup facade around the sole physical ADC table-copy authority, adding consistent operation logging and normalized error handling.
    ' Why retained: Retained so backup workflows share one physical copy authority while receiving the same diagnostics and failure contract without duplicating boilerplate.
    'Shared workflow-neutral wrapper around the sole physical ADC backup-table copy authority.
    On Error GoTo ErrHandler
    Dim backupTableName As String
    backupTableName = CreateBackupTableCopy(sourceTableName, backupStamp)
    CoreAddOperationNote "Backed up table [" & sourceTableName & "] to [" & backupTableName & "]."
    BackupOneTable = backupTableName
    Exit Function
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "BackupOneTable", "Could not back up access development console table." & vbCrLf & "Source table: " & sourceTableName & vbCrLf & "Err.Description: " & originalErrDescription
End Function

Private Function BackupTablesForSchemaRepair(ByVal schemaRepairSummary As String, ByRef editableLookupBackupCount As Long) As Long
    'Runtime table repair backup coordinator. Membership comes from the authoritative user-data scope; only TABLES-owned existing tables identified by the current schema-repair proof are copied.
    On Error GoTo ErrHandler
    Dim scopeDefinition As Variant
    Dim tableName As String
    Dim authoritativeOwner As String
    Dim backupBehavior As String
    Dim resetContract As String
    Dim backupStamp As String
    Dim backupName As String
    Dim backupCount As Long
    editableLookupBackupCount = 0
    backupStamp = Format$(Now, "yyyymmdd_hhnnss")
    For Each scopeDefinition In UserDataTableScopeDefinitions()
        authoritativeOwner = CStr(scopeDefinition(2))
        If StrComp(authoritativeOwner, "TABLES", vbTextCompare) = 0 Then
            tableName = CStr(scopeDefinition(1))
            If RepairSummaryContainsTable(schemaRepairSummary, tableName) And CodeDbTableExists(tableName) Then
                backupBehavior = CStr(scopeDefinition(3))
                If StrComp(backupBehavior, "FullDataCopy", vbTextCompare) <> 0 Then Err.Raise vbObjectError + 6237, "BackupTablesForSchemaRepair", "Unsupported canonical TABLES schema-repair backup behavior." & vbCrLf & "Table: " & tableName & vbCrLf & "Backup behavior: " & backupBehavior
                If TableCanOpen(tableName, True) = False Then Err.Raise vbObjectError + 6238, "BackupTablesForSchemaRepair", "An affected existing TABLES-owned user-data table cannot be opened, so runtime schema repair cannot confirm its required backup." & vbCrLf & "Table: " & tableName & vbCrLf & TableOpenFailureDetail(tableName)
                backupName = BackupOneTable(tableName, backupStamp)
                If Len(Trim$(backupName)) = 0 Then Err.Raise vbObjectError + 6239, "BackupTablesForSchemaRepair", "Runtime schema-repair backup creation returned no backup identity." & vbCrLf & "Table: " & tableName
                backupCount = backupCount + 1
                resetContract = CStr(scopeDefinition(4))
                If StrComp(resetContract, "EditableLookupTable", vbTextCompare) = 0 Then editableLookupBackupCount = editableLookupBackupCount + 1
            End If
        End If
    Next scopeDefinition
    BackupTablesForSchemaRepair = backupCount
    Exit Function
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrSource As String
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrSource = Err.Source
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "BackupTablesForSchemaRepair", "Could not create the required runtime schema-repair backups." & vbCrLf & "Err.Source: " & originalErrSource & vbCrLf & "Err.Description: " & originalErrDescription
End Function

' Purpose: Repairs terminology default rows for setup.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub RepairTerminologyDefaultRowsForSetup()
    'B2 sole setup owner for authoritative Help_Terms reference/default rows after B1 establishes the table structure.
    RequireSetupDomainMutationAllowed SETUP_DOMAIN_B2_CATALOG_REFERENCE_DEFAULTS, "RepairTerminologyDefaultRowsForSetup"
    On Error GoTo ErrHandler
    If CodeDbTableExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_TERMINOLOGY_TABLE_LOGICAL_ID)) = False Then Err.Raise vbObjectError + 9141, "RepairTerminologyDefaultRowsForSetup", "The B1-owned Help_Terms table is missing before B2 default-row repair." & vbCrLf & "Table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_TERMINOLOGY_TABLE_LOGICAL_ID)
    DefaultTerminologyRows
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "RepairTerminologyDefaultRowsForSetup", "Could not apply the B2-owned Help_Terms default-row repair." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

' Purpose: Ensures terminology review query for setup.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub EnsureTerminologyReviewQueryForSetup()
    'B1 owns this catalog-structure query immediately after the Help_Terms table is structurally current.
    RequireSetupDomainMutationAllowed SETUP_DOMAIN_B1_CATALOG_STRUCTURE, "EnsureTerminologyReviewQueryForSetup"
    If CodeDbTableExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_TERMINOLOGY_TABLE_LOGICAL_ID)) = False Then Err.Raise vbObjectError + 9210, "EnsureTerminologyReviewQueryForSetup", "B1 cannot create or verify " & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Terminology_Review") & " because the B1-owned Help_Terms table is missing." & vbCrLf & "Table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_TERMINOLOGY_TABLE_LOGICAL_ID)
    RepairOrReplaceQueryForSetup ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Terminology_Review"), TerminologyReviewSql(), False
End Sub

' Purpose: Repairs duplicate review query for setup.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub RepairDuplicateReviewQueryForSetup()
    RequireSetupDomainMutationAllowed SETUP_DOMAIN_A3_SAVED_QUERY_FOUNDATION, "RepairDuplicateReviewQueryForSetup"
    Dim querySql As String
    If CodeDbTableExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")) = False Then Err.Raise vbObjectError + 6480, "RepairDuplicateReviewQueryForSetup", "Required CodeDb MC_Records table is missing: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")
    If TableFieldExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config"), FIELD_SELECTED_FOR_DUPLICATE_DELETION) = False Then Err.Raise vbObjectError + 6481, "RepairDuplicateReviewQueryForSetup", "Required duplicate-review field is missing." & vbCrLf & "Table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & vbCrLf & "Field: SelectedForDuplicateDeletion"
    InvalidateDuplicateReviewCache
    querySql = DuplicateReviewSql()
    RepairOrReplaceQueryForSetup ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Config_DuplicateRows"), querySql, False
End Sub

Private Sub EnsureCustomMacroLauncherReviewQueryForSetup()
    RequireSetupDomainMutationAllowed SETUP_DOMAIN_A3_SAVED_QUERY_FOUNDATION, "EnsureCustomMacroLauncherReviewQueryForSetup"
    On Error GoTo ErrHandler
    If CodeDbTableExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LauncherControlsReview")) = False Then Exit Sub
    RepairOrReplaceQueryForSetup ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_LauncherControls_Review"), CustomMacroLauncherReviewSql(), False
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrSource As String
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrSource = Err.Source
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "EnsureCustomMacroLauncherReviewQueryForSetup", "Could not ensure the setup-owned Custom Macro Launcher review query." & vbCrLf & "Query: " & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_LauncherControls_Review") & vbCrLf & "Err.Source: " & originalErrSource & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

' Purpose: Checks whether core data table foundation already satisfies the exact state required by the current Setup run so the owning phase can determine whether repair is necessary.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   detailText (Output; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
' Returns: Boolean: True when core data table foundation matches the exact state required by the current Setup contract; False when the owning phase still needs repair, refresh, or reconciliation.
Public Function CoreDataTableFoundationCurrentForSetup(ByRef detailText As String) As Boolean
    On Error GoTo NotCurrent
    Dim snapshot As FoundationDomainSnapshot
    Dim db As DAO.Database
    Dim storedSchemaFingerprint As String
    Dim storedDataFingerprint As String
    Dim schemaExists As Boolean
    Dim dataExists As Boolean
    Dim schemaMatches As Boolean
    Dim dataMatches As Boolean
    detailText = vbNullString
    Set db = TablesWorkspaceCodeDb()
    storedSchemaFingerprint = DatabaseTextPropertyValue(db, FOUNDATION_SCHEMA_FINGERPRINT_PROPERTY, schemaExists)
    storedDataFingerprint = DatabaseTextPropertyValue(db, FOUNDATION_PART2_DATA_FINGERPRINT_PROPERTY, dataExists)
    If schemaExists = False Or Len(storedSchemaFingerprint) = 0 Or dataExists = False Or Len(storedDataFingerprint) = 0 Then
        detailText = "A2 schema=" & IIf(schemaExists And Len(storedSchemaFingerprint) > 0, "persisted", "unseeded") & "; defaults=" & IIf(dataExists And Len(storedDataFingerprint) > 0, "persisted", "unseeded") & "; physical comparison skipped because required persisted proof is unavailable."
        Exit Function
    End If
    If CalculateFoundationSchemaSnapshot(snapshot) = False Then detailText = snapshot.failureDetail: Exit Function
    schemaMatches = (StrComp(storedSchemaFingerprint, snapshot.SchemaFingerprint, vbBinaryCompare) = 0)
    dataMatches = (StrComp(storedDataFingerprint, snapshot.Part2DataFingerprint, vbBinaryCompare) = 0)
    CoreDataTableFoundationCurrentForSetup = schemaMatches And dataMatches
    detailText = "A2 schema=" & IIf(schemaMatches, "matched", "changed") & "; defaults=" & IIf(dataMatches, "matched", "changed") & "."
    Exit Function
NotCurrent:
    detailText = "A2 current-state verification failed. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
End Function

Private Sub ClearPopupMessagesSuppressionForCurrentRun()
    On Error Resume Next
    TempVars.Remove TEMPVAR_SUPPRESS_POPUPS
    TempVars.Remove TEMPVAR_SUPPRESS_POPUPS_REASON
End Sub

Private Function TableSafeDatabaseName() As String
On Error GoTo HandleError
    TableSafeDatabaseName = SharedDao_SafeDatabaseName(TablesWorkspaceCodeDb())
    Exit Function
HandleError:
    TableSafeDatabaseName = "(could not resolve ADCDb: " & Err.Description & ")"
End Function

Private Function FoundationPart2DataTableNames() As Collection
    'Rows whose default/reference metadata is owned by Part 2; user setting values and volatile timestamps are excluded by FoundationIncludeCatalogDataField.
    Dim names As Collection
    ' INTENTIONAL FILTERED PARTITION: names is a consumer-required subset/projection derived from the source enumerated by FoundationPart2DataTableNames; its membership is not independently authoritative.
    Set names = New Collection
    AddFoundationName names, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID)
    AddFoundationName names, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_MacroCategories")
    AddFoundationName names, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_MacroSubcategories")
    Set FoundationPart2DataTableNames = names
End Function

Private Function FoundationTableDefinitionText(ByVal db As DAO.Database, ByVal tableName As String, ByRef failureDetail As String) As String
    On Error GoTo ErrHandler
    Dim tdf As DAO.tableDef
    Dim fld As DAO.Field
    Dim idx As DAO.index
    Dim idxField As DAO.Field
    Dim resultText As String
    If SharedDao_TableExists(db, tableName) = False Then
        failureDetail = "Required foundation table is missing: " & tableName
        Exit Function
    End If
    Set tdf = db.TableDefs(tableName)
    tdf.fields.Refresh
    tdf.Indexes.Refresh
    resultText = FingerprintTokenPart(tableName)
    resultText = resultText & FingerprintTokenPart(CStr(tdf.Attributes))
    resultText = resultText & FingerprintTokenPart(CStr(Nz(tdf.Connect, vbNullString)))
    resultText = resultText & FingerprintTokenPart(CStr(Nz(tdf.sourceTableName, vbNullString)))
    resultText = resultText & FingerprintTokenPart(DaoObjectPropertyFingerprintText(tdf, "ValidationRule"))
    For Each fld In tdf.fields
        PumpSetupControlEvents
        resultText = resultText & FingerprintTokenPart(fld.Name)
        resultText = resultText & FingerprintTokenPart(CStr(fld.Type))
        resultText = resultText & FingerprintTokenPart(CStr(fld.Size))
        resultText = resultText & FingerprintTokenPart(CStr(fld.Attributes))
        resultText = resultText & FingerprintTokenPart(DaoObjectPropertyFingerprintText(fld, "Required"))
        resultText = resultText & FingerprintTokenPart(DaoObjectPropertyFingerprintText(fld, "AllowZeroLength"))
        resultText = resultText & FingerprintTokenPart(DaoObjectPropertyFingerprintText(fld, "DefaultValue"))
        resultText = resultText & FingerprintTokenPart(DaoObjectPropertyFingerprintText(fld, "ValidationRule"))
        resultText = resultText & FingerprintTokenPart(DaoObjectPropertyFingerprintText(fld, "ValidationText"))
        'A2 proves physical structure only; B2 owns non-form-dependent lookup/caption metadata and editor intent, D4 owns physical editor binding, and G1 owns managed DAO Description metadata, so those later-phase
        'properties are excluded from this fingerprint.
    Next fld
    'A2 structural proof owns foundation indexes, including its explicit primary/unique contracts. DAO's non-unique Foreign indexes are relationship support owned by B5.
    'Exclude only those B5-owned indexes so later relationship work cannot invalidate A2's persisted proof.
    For Each idx In tdf.Indexes
        PumpSetupControlEvents
        If idx.Foreign = False Or idx.Unique Or idx.Primary Then
            resultText = resultText & FingerprintTokenPart(idx.Name)
            resultText = resultText & FingerprintTokenPart(CStr(idx.Primary))
            resultText = resultText & FingerprintTokenPart(CStr(idx.Unique))
            resultText = resultText & FingerprintTokenPart(CStr(idx.Required))
            resultText = resultText & FingerprintTokenPart(CStr(idx.IgnoreNulls))
            For Each idxField In idx.fields
                PumpSetupControlEvents
                resultText = resultText & FingerprintTokenPart(idxField.Name)
            Next idxField
        End If
    Next idx
    FoundationTableDefinitionText = resultText
    Exit Function
ErrHandler:
    failureDetail = "Could not inspect foundation table [" & tableName & "]. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
End Function

Private Function TablesWorkspaceCodeDb() As DAO.Database
    'Uses Core.ADCDb as the runtime CodeDb/database-context authority while preserving this module's existing workspace-handle and setup-event-pump behavior.
    Dim workspaceValue As DAO.workspace
    Dim databaseValue As DAO.Database
    Dim codeDbPath As String
    codeDbPath = ADCDb().Name
    Set workspaceValue = DBEngine.Workspaces(0)
    For Each databaseValue In workspaceValue.Databases
        PumpSetupControlEvents
        If StrComp(databaseValue.Name, codeDbPath, vbTextCompare) = 0 Then
            Set TablesWorkspaceCodeDb = databaseValue
            Exit Function
        End If
    Next databaseValue
    Set TablesWorkspaceCodeDb = workspaceValue.OpenDatabase(codeDbPath)
End Function

' Purpose: Resolves canonical relationship definitions from the supplied/current ADC state, using `ManagedObjectPhysicalName` and `AddCanonicalMetadataRelationshipDefinitions` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: Collection: Returns the ordered Collection assembled by the routine; callers receive the object reference and should treat its membership/order according to the routine contract.
Public Function CanonicalRelationshipDefinitions() As Collection
    'Single canonical relationship definition set consumed by Setup repair and Health Review validation.
    ' CANONICAL AUTHORITY: rows owns the in-memory definition set constructed by CanonicalRelationshipDefinitions; downstream consumers derive views from this set rather than create a competing authority.
    Dim rows As New Collection
    Dim mo As String
    Dim moOverrides As String
    mo = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID)
    moOverrides = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECT_OVERRIDES_TABLE_LOGICAL_ID)
    'Seven relationships already present in the incoming baseline; preserve their established names.
    rows.Add Array(RELATION_MANAGED_OBJECTS_ACTION_PAYLOADS, mo, FIELD_ID, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectActionPayloads"), FIELD_MANAGED_OBJECT_ID, "CodeDb", True, False, False)
    rows.Add Array(RELATION_MANAGED_OBJECTS_DEPENDENCIES, mo, FIELD_ID, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectDependencies"), FIELD_MANAGED_OBJECT_ID, "CodeDb", True, False, False)
    rows.Add Array(RELATION_MANAGED_OBJECTS_DEPENDS_ON, mo, FIELD_ID, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectDependencies"), FIELD_DEPENDS_ON_MANAGED_OBJECT_ID, "CodeDb", True, False, False)
    rows.Add Array(RELATION_MANAGED_OBJECTS_FIELDS, mo, FIELD_ID, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFields"), FIELD_MANAGED_OBJECT_ID, "CodeDb", True, False, False)
    rows.Add Array(RELATION_MANAGED_OBJECT_FIELDS_VALUES, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFields"), FIELD_ID, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFieldValues"), FIELD_MANAGED_OBJECT_FIELD_VALUE_MANAGED_OBJECT_FIELD_ID, "CodeDb", True, False, False)
    rows.Add Array(RELATION_MACRO_CATEGORIES_SUBCATEGORIES, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_MacroCategories"), FIELD_MACRO_CATEGORY_KEY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_MacroSubcategories"), FIELD_MACRO_CATEGORY_KEY, "CodeDb", True, False, False)
    rows.Add Array(RELATION_TRACE_LOG_SCREENSHOTS, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_TraceLog"), FIELD_ID, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_TraceLogScreenshots"), FIELD_TRACE_LOG_ID, "CodeDb", True, False, False)
    'Dedicated sparse managed-object overrides are a one-to-zero-or-one child of MAC; cascade delete prevents orphaned override rows.
    rows.Add Array(RELATION_MANAGED_OBJECTS_OVERRIDES, mo, FIELD_ID, moOverrides, FIELD_MANAGED_OBJECT_ID, "CodeDb", True, False, True)
    'MC_Sys.
    rows.Add Array(RELATION_SYSTEM_MACROS_SYSTEM_MACRO_STATUS, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectStatuses"), FIELD_MANAGED_OBJECT_STATUS_CODE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SystemMacros"), FIELD_SYSTEM_MACRO_STATUS, "CodeDb", True, False, False)
    rows.Add Array(RELATION_SYSTEM_MACROS_MACRO_CATEGORY_KEY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_MacroCategories"), FIELD_MACRO_CATEGORY_KEY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SystemMacros"), FIELD_MACRO_CATEGORY_KEY, "CodeDb", True, False, False)
    rows.Add Array(RELATION_SYSTEM_MACROS_MACRO_SUBCATEGORY_KEY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_MacroSubcategories"), FIELD_MACRO_SUBCATEGORY_KEY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SystemMacros"), FIELD_MACRO_SUBCATEGORY_KEY, "CodeDb", True, False, False)
    rows.Add Array(RELATION_SYSTEM_MACROS_SYSTEM_MACRO_GROUP_KEY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SystemMacroGroup_Lookup"), FIELD_VALUE_TEXT, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SystemMacros"), FIELD_SYSTEM_MACRO_GROUP_KEY, "CodeDb", True, False, False)
    'Help_Docs.
    rows.Add Array(RELATION_HELP_DOCUMENTS_HELP_DOCUMENT_STATUS, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectStatuses"), FIELD_MANAGED_OBJECT_STATUS_CODE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HelpDocuments"), FIELD_HELP_DOCUMENT_STATUS, "CodeDb", True, False, False)
    'MAC_Cats.
    rows.Add Array(RELATION_OBJECT_CATEGORIES_OBJECT_CATEGORY_STATUS, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ObjectCategoryStatuses"), FIELD_OBJECT_CATEGORY_STATUS_CODE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ObjectCategories"), FIELD_OBJECT_CATEGORY_STATUS, "CodeDb", True, False, False)
    rows.Add Array(RELATION_OBJECT_CATEGORIES_CATALOG_OWNER, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_CatalogOwners"), FIELD_CATALOG_OWNER_CODE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ObjectCategories"), FIELD_CATALOG_OWNER, "CodeDb", True, False, False)
    rows.Add Array(RELATION_OBJECT_CATEGORIES_DEFAULT_OBJECT_TYPE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AccessObjectTypes"), FIELD_ACCESS_OBJECT_TYPE_CODE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ObjectCategories"), FIELD_DEFAULT_OBJECT_TYPE, "CodeDb", True, False, False)
    rows.Add Array(RELATION_OBJECT_CATEGORIES_DEFAULT_SETUP_ACTION, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SetupActions"), FIELD_SETUP_ACTION_CODE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ObjectCategories"), FIELD_DEFAULT_SETUP_ACTION, "CodeDb", True, False, False)
    rows.Add Array(RELATION_OBJECT_CATEGORIES_DEFAULT_OBJECT_SOURCE_DB_ROLE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_DatabaseRoles"), FIELD_DATABASE_ROLE_CODE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ObjectCategories"), FIELD_DEFAULT_OBJECT_SOURCE_DB_ROLE, "CodeDb", True, False, False)
    rows.Add Array(RELATION_OBJECT_CATEGORIES_DEFAULT_CONFIG_OWNER_DATABASE_ROLE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_DatabaseRoles"), FIELD_DATABASE_ROLE_CODE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ObjectCategories"), FIELD_DEFAULT_CONFIG_OWNER_DATABASE_ROLE, "CodeDb", True, False, False)
    rows.Add Array(RELATION_OBJECT_CATEGORIES_DEFAULT_ENVIRONMENT_SCOPE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_EnvironmentScopes"), FIELD_ENVIRONMENT_SCOPE_CODE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ObjectCategories"), FIELD_DEFAULT_ENVIRONMENT_SCOPE, "CodeDb", True, False, False)
    rows.Add Array(RELATION_OBJECT_CATEGORIES_DEFAULT_SETUP_STAGE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SetupStages"), FIELD_SETUP_STAGE_CODE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ObjectCategories"), FIELD_DEFAULT_SETUP_STAGE, "CodeDb", True, False, False)
    rows.Add Array(RELATION_OBJECT_CATEGORIES_DEFAULT_MANAGED_OBJECT_PURPOSE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectPurposes"), FIELD_MANAGED_OBJECT_PURPOSE_CODE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ObjectCategories"), FIELD_DEFAULT_MANAGED_OBJECT_PURPOSE, "CodeDb", True, False, False)
    rows.Add Array(RELATION_OBJECT_CATEGORIES_DEFAULT_NAV_PANE_VISIBILITY_POLICY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneVisibilityPolicies"), FIELD_NAV_PANE_VISIBILITY_POLICY_CODE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ObjectCategories"), FIELD_DEFAULT_NAV_PANE_VISIBILITY_POLICY, "CodeDb", True, False, False)
    rows.Add Array(RELATION_OBJECT_CATEGORIES_DEFAULT_EXPECTED_EXISTENCE_POLICY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ExpectedExistencePolicies"), FIELD_EXPECTED_EXISTENCE_POLICY_CODE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ObjectCategories"), FIELD_DEFAULT_EXPECTED_EXISTENCE_POLICY, "CodeDb", True, False, False)
    'MAC.
    rows.Add Array(RELATION_MANAGED_OBJECTS_OBJECT_TYPE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AccessObjectTypes"), FIELD_ACCESS_OBJECT_TYPE_CODE, mo, FIELD_OBJECT_TYPE, "CodeDb", True, False, False)
    rows.Add Array(RELATION_MANAGED_OBJECTS_CATEGORY_CODE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ObjectCategories"), FIELD_CATEGORY_CODE, mo, FIELD_CATEGORY_CODE, "CodeDb", True, False, False)
    rows.Add Array(RELATION_MANAGED_OBJECTS_MANAGED_OBJECT_STATUS, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectStatuses"), FIELD_MANAGED_OBJECT_STATUS_CODE, mo, FIELD_MANAGED_OBJECT_STATUS, "CodeDb", True, False, False)
    rows.Add Array(RELATION_MANAGED_OBJECTS_CATALOG_OWNER, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_CatalogOwners"), FIELD_CATALOG_OWNER_CODE, mo, FIELD_CATALOG_OWNER, "CodeDb", True, False, False)
    rows.Add Array(RELATION_MANAGED_OBJECTS_SETUP_ACTION_OVERRIDE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SetupActions"), FIELD_SETUP_ACTION_CODE, moOverrides, FIELD_SETUP_ACTION_OVERRIDE, "CodeDb", True, False, False)
    rows.Add Array(RELATION_MANAGED_OBJECTS_OBJECT_SOURCE_DB_ROLE_OVERRIDE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_DatabaseRoles"), FIELD_DATABASE_ROLE_CODE, moOverrides, FIELD_OBJECT_SOURCE_DB_ROLE_OVERRIDE, "CodeDb", True, False, False)
    rows.Add Array(RELATION_MANAGED_OBJECTS_CONFIG_OWNER_DATABASE_ROLE_OVERRIDE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_DatabaseRoles"), FIELD_DATABASE_ROLE_CODE, moOverrides, FIELD_CONFIG_OWNER_DATABASE_ROLE_OVERRIDE, "CodeDb", True, False, False)
    rows.Add Array(RELATION_MANAGED_OBJECTS_ENVIRONMENT_SCOPE_OVERRIDE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_EnvironmentScopes"), FIELD_ENVIRONMENT_SCOPE_CODE, moOverrides, FIELD_ENVIRONMENT_SCOPE_OVERRIDE, "CodeDb", True, False, False)
    rows.Add Array(RELATION_MANAGED_OBJECTS_SETUP_STAGE_OVERRIDE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SetupStages"), FIELD_SETUP_STAGE_CODE, moOverrides, FIELD_SETUP_STAGE_OVERRIDE, "CodeDb", True, False, False)
    rows.Add Array(RELATION_MANAGED_OBJECTS_MANAGED_OBJECT_PURPOSE_OVERRIDE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectPurposes"), FIELD_MANAGED_OBJECT_PURPOSE_CODE, moOverrides, FIELD_MANAGED_OBJECT_PURPOSE_OVERRIDE, "CodeDb", True, False, False)
    rows.Add Array(RELATION_MANAGED_OBJECTS_NAV_PANE_VISIBILITY_POLICY_OVERRIDE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneVisibilityPolicies"), FIELD_NAV_PANE_VISIBILITY_POLICY_CODE, moOverrides, FIELD_NAV_PANE_VISIBILITY_POLICY_OVERRIDE, "CodeDb", True, False, False)
    rows.Add Array(RELATION_MANAGED_OBJECTS_EXPECTED_EXISTENCE_POLICY_OVERRIDE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ExpectedExistencePolicies"), FIELD_EXPECTED_EXISTENCE_POLICY_CODE, moOverrides, FIELD_EXPECTED_EXISTENCE_POLICY_OVERRIDE, "CodeDb", True, False, False)
    rows.Add Array(RELATION_MANAGED_OBJECTS_NAVIGATION_PANE_GROUP_SETTING_KEY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups"), FIELD_SETTING_KEY, mo, FIELD_NAVIGATION_PANE_GROUP_SETTING_KEY, "CodeDb", True, False, False)
    rows.Add Array(RELATION_MANAGED_OBJECTS_LAST_SETUP_RESULT, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SetupResults"), FIELD_SETUP_RESULT_CODE, mo, FIELD_LAST_SETUP_RESULT, "CodeDb", True, False, False)
    'UI_Ribbon_Generated_BtnTargets and UI_Groups.
    rows.Add Array(RELATION_RIBBON_BUTTON_TARGETS_OBJECT_TYPE_CODE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AccessObjectTypes"), FIELD_ACCESS_OBJECT_TYPE_CODE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_RIBBON_BUTTON_TARGETS_TABLE_LOGICAL_ID), FIELD_OBJECT_TYPE_CODE, "CodeDb", True, False, False)
    rows.Add Array(RELATION_NAVIGATION_PANE_GROUPS_CONFIG_OWNER_DATABASE_ROLE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_DatabaseRoles"), FIELD_DATABASE_ROLE_CODE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups"), FIELD_CONFIG_OWNER_DATABASE_ROLE, "CodeDb", True, False, False)
    rows.Add Array(RELATION_NAVIGATION_PANE_GROUPS_OBJECT_TYPE_SCOPE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AccessObjectTypes"), FIELD_ACCESS_OBJECT_TYPE_CODE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups"), FIELD_OBJECT_TYPE_SCOPE, "CodeDb", True, False, False)
    'UI_Frm_Registry.
    rows.Add Array(RELATION_FORM_DEFINITIONS_FORM_TYPE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_FormTypes"), FIELD_FORM_TYPE_CODE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_FormDefinitions"), FIELD_FORM_TYPE, "CodeDb", True, False, False)
    rows.Add Array(RELATION_FORM_DEFINITIONS_NAVIGATION_PANE_GROUP_KEY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups"), FIELD_NAVIGATION_PANE_GROUP_KEY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_FormDefinitions"), FIELD_NAVIGATION_PANE_GROUP_KEY, "CodeDb", True, False, False)
    rows.Add Array(RELATION_FORM_DEFINITIONS_CONFIG_OWNER_DATABASE_ROLE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_DatabaseRoles"), FIELD_DATABASE_ROLE_CODE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_FormDefinitions"), FIELD_CONFIG_OWNER_DATABASE_ROLE, "CodeDb", True, False, False)
    rows.Add Array(RELATION_FORM_DEFINITIONS_FORM_STATUS, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectStatuses"), FIELD_MANAGED_OBJECT_STATUS_CODE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_FormDefinitions"), FIELD_FORM_STATUS, "CodeDb", True, False, False)
    rows.Add Array(RELATION_FORM_DEFINITIONS_CATALOG_OWNER, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_CatalogOwners"), FIELD_CATALOG_OWNER_CODE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_FormDefinitions"), FIELD_CATALOG_OWNER, "CodeDb", True, False, False)
    'MAC_ActionPayloads.
    rows.Add Array(RELATION_MANAGED_OBJECT_ACTION_PAYLOADS_PAYLOAD_KIND, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActionPayloadKinds"), FIELD_PAYLOAD_KIND_CODE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectActionPayloads"), FIELD_PAYLOAD_KIND, "CodeDb", True, False, False)
    rows.Add Array(RELATION_MANAGED_OBJECT_ACTION_PAYLOADS_PAYLOAD_STATUS, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectStatuses"), FIELD_MANAGED_OBJECT_STATUS_CODE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectActionPayloads"), FIELD_PAYLOAD_STATUS, "CodeDb", True, False, False)
    rows.Add Array(RELATION_MANAGED_OBJECT_ACTION_PAYLOADS_CATALOG_OWNER, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_CatalogOwners"), FIELD_CATALOG_OWNER_CODE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectActionPayloads"), FIELD_CATALOG_OWNER, "CodeDb", True, False, False)
    'MAC_Deps.
    rows.Add Array(RELATION_MANAGED_OBJECT_DEPENDENCIES_DEPENDENCY_TYPE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_DependencyTypes"), FIELD_DEPENDENCY_TYPE_CODE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectDependencies"), FIELD_DEPENDENCY_TYPE, "CodeDb", True, False, False)
    rows.Add Array(RELATION_MANAGED_OBJECT_DEPENDENCIES_DEPENDENCY_OBJECT_TYPE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AccessObjectTypes"), FIELD_ACCESS_OBJECT_TYPE_CODE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectDependencies"), FIELD_DEPENDENCY_OBJECT_TYPE, "CodeDb", True, False, False)
    rows.Add Array(RELATION_MANAGED_OBJECT_DEPENDENCIES_DEPENDENCY_DATABASE_ROLE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_DatabaseRoles"), FIELD_DATABASE_ROLE_CODE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectDependencies"), FIELD_MANAGED_OBJECT_DEPENDENCY_DATABASE_ROLE, "CodeDb", True, False, False)
    rows.Add Array(RELATION_MANAGED_OBJECT_DEPENDENCIES_DEPENDENCY_STATUS, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_DependencyStatuses"), FIELD_DEPENDENCY_STATUS_CODE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectDependencies"), FIELD_MANAGED_OBJECT_DEPENDENCY_STATUS, "CodeDb", True, False, False)
    'CodeDb configuration/candidate relationships. HostDb MC_Records intentionally receives none of these CodeDb-owned relationships.
    rows.Add Array(RELATION_CONFIG_CONFIG_OWNER_DATABASE_ROLE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_DatabaseRoles"), FIELD_DATABASE_ROLE_CODE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config"), FIELD_CONFIG_OWNER_DATABASE_ROLE, "CodeDb", True, False, False)
    rows.Add Array(RELATION_CANDIDATES_CONFIGURATION, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config"), FIELD_ID, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates"), FIELD_ACCEPTED_CONFIGURATION_ID, "CodeDb", True, False, False)
    rows.Add Array(RELATION_CONFIG_MACRO_CATEGORY_KEY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_MacroCategories"), FIELD_MACRO_CATEGORY_KEY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config"), FIELD_MACRO_CATEGORY_KEY, "CodeDb", True, False, False)
    rows.Add Array(RELATION_CONFIG_MACRO_SUBCATEGORY_KEY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_MacroSubcategories"), FIELD_MACRO_SUBCATEGORY_KEY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config"), FIELD_MACRO_SUBCATEGORY_KEY, "CodeDb", True, False, False)
    'Composite category/subcategory integrity.
    rows.Add Array(RELATION_SYSTEM_MACROS_MACRO_CATEGORY_SUBCATEGORY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_MacroSubcategories"), FIELD_MACRO_CATEGORY_KEY & "|" & FIELD_MACRO_SUBCATEGORY_KEY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SystemMacros"), FIELD_MACRO_CATEGORY_KEY & "|" & FIELD_MACRO_SUBCATEGORY_KEY, "CodeDb", True, False, False)
    rows.Add Array(RELATION_CONFIG_MACRO_CATEGORY_SUBCATEGORY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_MacroSubcategories"), FIELD_MACRO_CATEGORY_KEY & "|" & FIELD_MACRO_SUBCATEGORY_KEY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config"), FIELD_MACRO_CATEGORY_KEY & "|" & FIELD_MACRO_SUBCATEGORY_KEY, "CodeDb", True, False, False)
    rows.Add Array(RELATION_CANDIDATES_MACRO_CATEGORY_SUBCATEGORY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_MacroSubcategories"), FIELD_MACRO_CATEGORY_KEY & "|" & FIELD_MACRO_SUBCATEGORY_KEY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates"), FIELD_MACRO_CATEGORY_KEY & "|" & FIELD_MACRO_SUBCATEGORY_KEY, "CodeDb", True, False, False)
    'Retained Reference/Add-In ownership and resolved-identity integrity. A4 owns the participating physical tables; B5 owns these CodeDb relationships and their child-field indexes.
    rows.Add Array(RELATION_LOADED_VBA_PROJECTS_ACCESS_ADD_INS, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects"), FIELD_PROJECT_IDENTITY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AccessAddIns"), FIELD_PROJECT_IDENTITY, "CodeDb", True, False, False)
    rows.Add Array(RELATION_LOADED_VBA_PROJECTS_PROJECT_REFERENCES, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects"), FIELD_PROJECT_IDENTITY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjectReferences"), FIELD_PROJECT_IDENTITY, "CodeDb", True, False, False)
    rows.Add Array(RELATION_LOADED_VBA_PROJECTS_REFERENCED_PROJECTS, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects"), FIELD_PROJECT_IDENTITY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjectReferences"), FIELD_REFERENCED_PROJECT_IDENTITY, "CodeDb", True, False, False)
    rows.Add Array(RELATION_LOADED_VBA_PROJECTS_MODULES, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects"), FIELD_PROJECT_IDENTITY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAModules"), FIELD_PROJECT_IDENTITY, "CodeDb", True, False, False)
    rows.Add Array(RELATION_LOADED_VBA_PROJECTS_PROCEDURES, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects"), FIELD_PROJECT_IDENTITY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProcedures"), FIELD_PROJECT_IDENTITY, "CodeDb", True, False, False)
    rows.Add Array(RELATION_LOADED_VBA_PROJECTS_DEPENDENCIES, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects"), FIELD_PROJECT_IDENTITY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBADependencies"), FIELD_PROJECT_IDENTITY, "CodeDb", True, False, False)
    rows.Add Array(RELATION_LOADED_VBA_PROJECTS_DEPENDENCY_TARGETS, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects"), FIELD_PROJECT_IDENTITY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBADependencies"), FIELD_TARGET_PROJECT_IDENTITY, "CodeDb", True, False, False)
    rows.Add Array(RELATION_ACTIVE_REFERENCES_PROJECT_LINKS, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActiveReferences"), FIELD_REFERENCE_IDENTITY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActiveReferenceProjectLinks"), FIELD_REFERENCE_IDENTITY, "CodeDb", True, False, False)
    rows.Add Array(RELATION_LOADED_VBA_PROJECTS_ACTIVE_REFERENCE_LINKS, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects"), FIELD_PROJECT_IDENTITY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActiveReferenceProjectLinks"), FIELD_LOADED_PROJECT_IDENTITY, "CodeDb", True, False, False)
    rows.Add Array(RELATION_LOADED_VBA_MODULES_PROCEDURES, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAModules"), FIELD_MODULE_IDENTITY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProcedures"), FIELD_MODULE_IDENTITY, "CodeDb", True, False, False)
    rows.Add Array(RELATION_LOADED_VBA_MODULES_DEPENDENCY_SOURCES, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAModules"), FIELD_MODULE_IDENTITY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBADependencies"), FIELD_SOURCE_MODULE_IDENTITY, "CodeDb", True, False, False)
    rows.Add Array(RELATION_LOADED_VBA_MODULES_DEPENDENCY_TARGETS, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAModules"), FIELD_MODULE_IDENTITY, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBADependencies"), FIELD_TARGET_MODULE_IDENTITY, "CodeDb", True, False, False)
    'Complete the canonical set from the same authoritative table-backed lookup metadata used to configure fields.
    'Only closed-list persisted table lookups whose parent field is the source's authoritative code/key qualify: live/query sources, display fields, and LimitToList=False presentation lists intentionally do not
    'create RI.
    AddCanonicalMetadataRelationshipDefinitions rows
    Set CanonicalRelationshipDefinitions = rows
End Function

Private Sub AddCanonicalMetadataRelationshipDefinitions(ByRef rows As Collection)
    Dim logicalTableNames As Variant
    Dim logicalTableName As Variant
    Dim childTableName As String
    Dim metadataDefinitions As Collection
    Dim metadataDefinition As Variant
    Dim relationshipName As String
    Dim parentTableName As String
    Dim parentFieldName As String
    Dim childFieldName As String
    logicalTableNames = Array( _
        "ADC_SystemMacros", "ADC_HelpDocuments", "ADC_TraceLog", "ADC_ObjectCategories", _
        DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID, "ADC_ManagedObjectFields", _
        DEFAULT_RIBBON_BUTTON_TARGETS_TABLE_LOGICAL_ID, "ADC_NavigationPaneGroups", _
        "ADC_FormDefinitions", "ADC_ManagedObjectActionPayloads", "ADC_ManagedObjectDependencies", _
        "ADC_ManagedObjectFieldValues", "ADC_LoadedVBAProjects", "ADC_VBEAddIns", "ADC_COMAddIns", _
        "ADC_AccessAddIns", "ADC_LoadedVBAProjectReferences", "ADC_LoadedVBAModules", _
        "ADC_LoadedVBAProcedures", "ADC_LoadedVBADependencies", "ADC_ActiveReferenceProjectLinks", _
        "ADC_ReferenceManager", "ADC_ActiveReferences", "ADC_AvailableReferences", _
        DEFAULT_SETTINGS_TABLE_LOGICAL_ID, "ADC_Config", "ADC_Candidates", "ADC_HealthReview", _
        DEFAULT_TERMINOLOGY_TABLE_LOGICAL_ID, "ADC_LauncherControlsReview", _
        "ADC_ModuleUpdateHistory", "ADC_ModuleVersionManifest", "ADC_ModuleVersionVerification", _
        "ADC_LookupDefinitions")
    For Each logicalTableName In logicalTableNames
        childTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, CStr(logicalTableName))
        Set metadataDefinitions = B2SharedFieldMetadataDefinitions(childTableName)
        For Each metadataDefinition In metadataDefinitions
            If StrComp(CStr(metadataDefinition(0)), CATALOG_FIELD_LOOKUP_DEFINITION_KIND_TABLE, vbTextCompare) = 0 Then
                If CBool(metadataDefinition(5)) Then
                    childFieldName = Trim$(CStr(metadataDefinition(1)))
                    parentTableName = Trim$(CStr(metadataDefinition(2)))
                    parentFieldName = Trim$(CStr(metadataDefinition(3)))
                    If CanonicalMetadataRelationshipEligible(parentTableName, parentFieldName) Then
                        If CanonicalRelationshipFieldCovered(rows, parentTableName, parentFieldName, childTableName, childFieldName) = False Then
                            relationshipName = RELATION_PREFIX & CStr(logicalTableName) & "_" & childFieldName
                            If Len(relationshipName) > 64 Then Err.Raise vbObjectError + 9390, "AddCanonicalMetadataRelationshipDefinitions", "The canonical metadata relationship name exceeds the Access object-name limit." & vbCrLf & "Relationship: " & relationshipName
                            If CanonicalRelationshipNameExists(rows, relationshipName) Then Err.Raise vbObjectError + 9391, "AddCanonicalMetadataRelationshipDefinitions", "The canonical metadata relationship name is not unique." & vbCrLf & "Relationship: " & relationshipName
                            rows.Add Array(relationshipName, parentTableName, parentFieldName, childTableName, childFieldName, "CodeDb", True, False, False)
                        End If
                    End If
                End If
            End If
        Next metadataDefinition
    Next logicalTableName
End Sub

Private Function CanonicalMetadataRelationshipEligible(ByVal parentTableName As String, ByVal parentFieldName As String) As Boolean
    'Physical RI requires the lookup source's authoritative stored code/key; a closed display field alone is not relational identity.
    Dim policyData As Variant
    policyData = LookupEditorPolicyForPhysicalSource(parentTableName)
    If IsArray(policyData) = False Then Exit Function
    CanonicalMetadataRelationshipEligible = (StrComp(Trim$(CStr(policyData(LOOKUP_EDITOR_DEF_CODE_FIELD))), parentFieldName, vbTextCompare) = 0)
End Function

Private Function CanonicalRelationshipFieldCovered(ByVal rows As Collection, ByVal parentTableName As String, ByVal parentFieldName As String, ByVal childTableName As String, ByVal childFieldName As String) As Boolean
    Dim Definition As Variant
    Dim parentFields As Variant
    Dim childFields As Variant
    Dim fieldIndex As Long
    For Each Definition In rows
        If StrComp(CStr(Definition(1)), parentTableName, vbTextCompare) = 0 And _
           StrComp(CStr(Definition(3)), childTableName, vbTextCompare) = 0 Then
            parentFields = Split(CStr(Definition(2)), "|")
            childFields = Split(CStr(Definition(4)), "|")
            If UBound(parentFields) - LBound(parentFields) = UBound(childFields) - LBound(childFields) Then
                For fieldIndex = LBound(parentFields) To UBound(parentFields)
                    If StrComp(Trim$(CStr(parentFields(fieldIndex))), parentFieldName, vbTextCompare) = 0 And _
                       StrComp(Trim$(CStr(childFields(fieldIndex))), childFieldName, vbTextCompare) = 0 Then
                        CanonicalRelationshipFieldCovered = True
                        Exit Function
                    End If
                Next fieldIndex
            End If
        End If
    Next Definition
End Function

Private Function CanonicalRelationshipNameExists(ByVal rows As Collection, ByVal relationshipName As String) As Boolean
    Dim Definition As Variant
    For Each Definition In rows
        If StrComp(CStr(Definition(0)), relationshipName, vbTextCompare) = 0 Then
            CanonicalRelationshipNameExists = True
            Exit Function
        End If
    Next Definition
End Function

Private Function LookupReferenceRecordExists(ByVal db As DAO.Database, ByVal foreignTableName As String, ByVal primaryFieldNames As Variant, ByVal foreignFieldNames As Variant, ByVal lookupRowValues As Object) As Boolean
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim whereText As String
    Dim primaryFieldName As String
    Dim foreignFieldName As String
    Dim fieldIndex As Long
    Dim lookupValue As Variant
    For fieldIndex = LBound(foreignFieldNames) To UBound(foreignFieldNames)
        primaryFieldName = Trim$(CStr(primaryFieldNames(fieldIndex)))
        foreignFieldName = Trim$(CStr(foreignFieldNames(fieldIndex)))
        If lookupRowValues.Exists(primaryFieldName) = False Then Err.Raise vbObjectError + 9384, "LookupReferenceRecordExists", "The generated lookup editor did not supply a lookup field required by a reference definition." & vbCrLf & "Lookup field: " & primaryFieldName & vbCrLf & "Referenced table: " & foreignTableName
        lookupValue = lookupRowValues(primaryFieldName)
        If Len(whereText) > 0 Then whereText = whereText & " AND "
        If IsNull(lookupValue) Then
            whereText = whereText & "[" & Replace(foreignFieldName, "]", "]]") & "] Is Null"
        Else
            whereText = whereText & "[" & Replace(foreignFieldName, "]", "]]") & "]=" & SharedDao_SqlQuote(CStr(lookupValue))
        End If
    Next fieldIndex
    sqlText = "SELECT TOP 1 1 AS " & FIELD_MATCH_FOUND & " FROM [" & Replace(foreignTableName, "]", "]]") & "] WHERE " & whereText & ";"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot, dbReadOnly)
    LookupReferenceRecordExists = (rs.EOF = False)
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
    Set rs = Nothing
    On Error GoTo 0
    Err.Raise savedErrNumber, "LookupReferenceRecordExists", "Could not inspect a lookup reference table." & vbCrLf & "Table: " & foreignTableName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function LookupReferenceFieldKey(ByVal tableName As String, ByVal fieldName As String) As String
    LookupReferenceFieldKey = Trim$(tableName) & "|" & fieldName
End Function

Private Function LookupMetadataReferenceUsageInDatabase(ByVal db As DAO.Database, ByVal databaseRole As String, ByVal lookupTableName As String, ByVal lookupCodeFieldName As String, ByVal lookupRowValues As Object, ByVal relationshipCoveredFields As Object, ByRef usageDetail As String, Optional ByVal tableNameFilter As String = vbNullString) As Boolean
    On Error GoTo ErrHandler
    Dim tdf As DAO.tableDef
    Dim metadataDefinitions As Collection
    Dim metadataDefinition As Variant
    Dim sourceFieldName As String
    Dim referenceKey As String
    Dim skipMetadataReference As Boolean
    tableNameFilter = Trim$(tableNameFilter)
    LogPhysicalTraceMessageOnly "Info", "LookupEditorDelete", "LookupMetadataReferenceUsageInDatabase", "START", "Database role: " & databaseRole & vbCrLf & "Lookup table: " & lookupTableName & vbCrLf & "Table filter: " & IIf(Len(tableNameFilter) = 0, "(all tables)", tableNameFilter), "Inspect lookup metadata references", 0, vbNullString, vbNullString, lookupTableName, "Table", databaseRole, "LookupEditor"
    For Each tdf In db.TableDefs
        If Len(tableNameFilter) = 0 Or StrComp(tdf.Name, tableNameFilter, vbTextCompare) = 0 Then
            Set metadataDefinitions = B2SharedFieldMetadataDefinitions(tdf.Name)
            For Each metadataDefinition In metadataDefinitions
                If StrComp(CStr(metadataDefinition(0)), CATALOG_FIELD_LOOKUP_DEFINITION_KIND_TABLE, vbTextCompare) = 0 Then
                    If StrComp(Trim$(CStr(metadataDefinition(2))), lookupTableName, vbTextCompare) = 0 And _
                       StrComp(Trim$(CStr(metadataDefinition(3))), lookupCodeFieldName, vbTextCompare) = 0 Then
                        sourceFieldName = Trim$(CStr(metadataDefinition(1)))
                        referenceKey = LookupReferenceFieldKey(tdf.Name, sourceFieldName)
                        skipMetadataReference = False
                        If Not relationshipCoveredFields Is Nothing Then skipMetadataReference = relationshipCoveredFields.Exists(referenceKey)
                        If skipMetadataReference = False Then
                            If SharedDao_DatabaseTableFieldExistsSafe(db, tdf.Name, sourceFieldName) = False Then Err.Raise vbObjectError + 9386, "LookupMetadataReferenceUsageInDatabase", "A table-backed lookup metadata field is missing from its table." & vbCrLf & "Database role: " & databaseRole & vbCrLf & "Table: " & tdf.Name & vbCrLf & "Field: " & sourceFieldName & vbCrLf & "Lookup table: " & lookupTableName
                            If LookupReferenceRecordExists(db, tdf.Name, Array(lookupCodeFieldName), Array(sourceFieldName), lookupRowValues) Then
                                usageDetail = "Referenced database: " & databaseRole & vbCrLf & "Referenced table: " & tdf.Name & vbCrLf & "Referenced field: " & sourceFieldName
                                LookupMetadataReferenceUsageInDatabase = True
                                LogPhysicalTraceMessageOnly "Info", "LookupEditorDelete", "LookupMetadataReferenceUsageInDatabase", "DONE", "Database role: " & databaseRole & vbCrLf & "Lookup table: " & lookupTableName & vbCrLf & "Referenced: True" & vbCrLf & usageDetail, "Inspect lookup metadata references", 0, vbNullString, vbNullString, lookupTableName, "Table", databaseRole, "LookupEditor"
                                Exit Function
                            End If
                        End If
                    End If
                End If
            Next metadataDefinition
            If Len(tableNameFilter) > 0 Then Exit For
        End If
    Next tdf
    LogPhysicalTraceMessageOnly "Info", "LookupEditorDelete", "LookupMetadataReferenceUsageInDatabase", "DONE", "Database role: " & databaseRole & vbCrLf & "Lookup table: " & lookupTableName & vbCrLf & "Referenced: False", "Inspect lookup metadata references", 0, vbNullString, vbNullString, lookupTableName, "Table", databaseRole, "LookupEditor"
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    LogPhysicalTraceMessageOnly "Error", "LookupEditorDelete", "LookupMetadataReferenceUsageInDatabase", "ERROR", "Database role: " & databaseRole & vbCrLf & "Lookup table: " & lookupTableName, "Inspect lookup metadata references", savedErrNumber, savedErrSource, savedErrDescription, lookupTableName, "Table", databaseRole, "LookupEditor"
    Err.Raise savedErrNumber, "LookupMetadataReferenceUsageInDatabase", "Could not inspect table-backed lookup metadata usage." & vbCrLf & "Database role: " & databaseRole & vbCrLf & "Lookup table: " & lookupTableName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

' Purpose: Checks whether the specified lookup record is referenced by managed data and returns reference details used to decide whether deletion is safe.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   lookupTableName (Input; String): Physical Access lookup-table name whose rows/fields supply the lookup definition or validation target.
'   lookupRowValues (Input; Object): Collection/object containing the lookup-row values to validate or persist, in the field order expected by the lookup operation.
'   usageDetail (Output; String): Receives usage detail used for caller diagnostics, verification output, or user-facing reporting.
' Returns: Boolean: True when the target lookup record is found to be referenced and the routine populates the applicable reference detail; False when no managed reference is found.
Public Function LookupRecordReferenceUsage(ByVal lookupTableName As String, ByVal lookupRowValues As Object, ByRef usageDetail As String) As Boolean
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim hostDb As DAO.Database
    Dim Definition As Variant
    Dim policyData As Variant
    Dim primaryFieldNames As Variant
    Dim foreignFieldNames As Variant
    Dim primaryFieldName As String
    Dim foreignFieldName As String
    Dim foreignTableName As String
    Dim lookupCodeFieldName As String
    Dim configTableName As String
    Dim fieldIndex As Long
    Dim relationshipCoveredFields As Object
    Dim noRelationshipCoveredFields As Object
    Dim relationshipDefinitions As Collection
    lookupTableName = Trim$(lookupTableName)
    usageDetail = vbNullString
    LogPhysicalTraceMessageOnly "Info", "LookupEditorDelete", "LookupRecordReferenceUsage", "START", "Lookup table: " & lookupTableName, "Verify lookup record references", 0, vbNullString, vbNullString, lookupTableName, "Table", "CodeDb", "LookupEditor"
    If Len(lookupTableName) = 0 Then Err.Raise vbObjectError + 9380, "LookupRecordReferenceUsage", "Lookup table name is blank."
    If lookupRowValues Is Nothing Then Err.Raise vbObjectError + 9381, "LookupRecordReferenceUsage", "Lookup row values are unavailable." & vbCrLf & "Table: " & lookupTableName
    Set db = ADCDb()
    If SharedDao_DatabaseTableExistsSafe(db, lookupTableName) = False Then Err.Raise vbObjectError + 9382, "LookupRecordReferenceUsage", "The lookup source table does not exist in CodeDb." & vbCrLf & "Table: " & lookupTableName
    policyData = LookupEditorPolicyForPhysicalSource(lookupTableName)
    If IsArray(policyData) = False Then Err.Raise vbObjectError + 9387, "LookupRecordReferenceUsage", "No lookup-editor policy exists for the lookup source table." & vbCrLf & "Table: " & lookupTableName
    lookupCodeFieldName = Trim$(CStr(policyData(LOOKUP_EDITOR_DEF_CODE_FIELD)))
    If Len(lookupCodeFieldName) = 0 Then Err.Raise vbObjectError + 9388, "LookupRecordReferenceUsage", "The lookup-editor policy has no authoritative code/key field." & vbCrLf & "Table: " & lookupTableName
    If lookupRowValues.Exists(lookupCodeFieldName) = False Then Err.Raise vbObjectError + 9389, "LookupRecordReferenceUsage", "The generated lookup editor did not supply its authoritative code/key value." & vbCrLf & "Lookup table: " & lookupTableName & vbCrLf & "Required field: " & lookupCodeFieldName
    Set relationshipCoveredFields = CreateObject("Scripting.Dictionary")
    relationshipCoveredFields.CompareMode = vbTextCompare
    LogPhysicalTraceMessageOnly "Info", "LookupEditorDelete", "LookupRecordReferenceUsage", "CanonicalRelationshipDefinitions START", "Lookup table: " & lookupTableName, "Build canonical relationship definitions", 0, vbNullString, vbNullString, lookupTableName, "Table", "CodeDb", "LookupEditor"
    Set relationshipDefinitions = CanonicalRelationshipDefinitions()
    LogPhysicalTraceMessageOnly "Info", "LookupEditorDelete", "LookupRecordReferenceUsage", "CanonicalRelationshipDefinitions DONE", "Lookup table: " & lookupTableName & vbCrLf & "Definition count: " & CStr(relationshipDefinitions.count), "Build canonical relationship definitions", 0, vbNullString, vbNullString, lookupTableName, "Table", "CodeDb", "LookupEditor"
    For Each Definition In relationshipDefinitions
        If StrComp(CStr(Definition(5)), "CodeDb", vbTextCompare) = 0 And StrComp(CStr(Definition(1)), lookupTableName, vbTextCompare) = 0 Then
            primaryFieldNames = Split(CStr(Definition(2)), "|")
            foreignFieldNames = Split(CStr(Definition(4)), "|")
            If UBound(primaryFieldNames) - LBound(primaryFieldNames) <> UBound(foreignFieldNames) - LBound(foreignFieldNames) Then Err.Raise vbObjectError + 9383, "LookupRecordReferenceUsage", "A canonical relationship has mismatched primary/foreign field counts." & vbCrLf & "Relationship: " & CStr(Definition(0))
            foreignTableName = Trim$(CStr(Definition(3)))
            For fieldIndex = LBound(primaryFieldNames) To UBound(primaryFieldNames)
                primaryFieldName = Trim$(CStr(primaryFieldNames(fieldIndex)))
                foreignFieldName = Trim$(CStr(foreignFieldNames(fieldIndex)))
                relationshipCoveredFields(LookupReferenceFieldKey(foreignTableName, foreignFieldName)) = True
                If lookupRowValues.Exists(primaryFieldName) = False Then Err.Raise vbObjectError + 9384, "LookupRecordReferenceUsage", "The generated lookup editor did not supply a parent field required by a canonical relationship." & vbCrLf & "Lookup table: " & lookupTableName & vbCrLf & "Relationship: " & CStr(Definition(0)) & vbCrLf & "Required field: " & primaryFieldName
            Next fieldIndex
            If SharedDao_DatabaseTableExistsSafe(db, foreignTableName) Then
                For fieldIndex = LBound(foreignFieldNames) To UBound(foreignFieldNames)
                    foreignFieldName = Trim$(CStr(foreignFieldNames(fieldIndex)))
                    If SharedDao_DatabaseTableFieldExistsSafe(db, foreignTableName, foreignFieldName) = False Then Err.Raise vbObjectError + 9385, "LookupRecordReferenceUsage", "A canonical lookup reference field is missing from its table." & vbCrLf & "Relationship: " & CStr(Definition(0)) & vbCrLf & "Table: " & foreignTableName & vbCrLf & "Field: " & foreignFieldName
                Next fieldIndex
                If LookupReferenceRecordExists(db, foreignTableName, primaryFieldNames, foreignFieldNames, lookupRowValues) Then
                    usageDetail = "Referenced database: CodeDb" & vbCrLf & "Referenced table: " & foreignTableName & vbCrLf & "Referenced field" & IIf(UBound(foreignFieldNames) = LBound(foreignFieldNames), ": ", "s: ") & Replace(CStr(Definition(4)), "|", ", ")
                    LookupRecordReferenceUsage = True
                    LogPhysicalTraceMessageOnly "Info", "LookupEditorDelete", "LookupRecordReferenceUsage", "DONE", "Lookup table: " & lookupTableName & vbCrLf & "Referenced: True" & vbCrLf & usageDetail, "Verify lookup record references", 0, vbNullString, vbNullString, lookupTableName, "Table", "CodeDb", "LookupEditor"
                    Exit Function
                End If
            End If
        End If
    Next Definition
    If LookupMetadataReferenceUsageInDatabase(db, "CodeDb", lookupTableName, lookupCodeFieldName, lookupRowValues, relationshipCoveredFields, usageDetail) Then
        LookupRecordReferenceUsage = True
        LogPhysicalTraceMessageOnly "Info", "LookupEditorDelete", "LookupRecordReferenceUsage", "DONE", "Lookup table: " & lookupTableName & vbCrLf & "Referenced: True" & vbCrLf & usageDetail, "Verify lookup record references", 0, vbNullString, vbNullString, lookupTableName, "Table", "CodeDb", "LookupEditor"
        Exit Function
    End If
    If CodeDbIsHostDb() = False Then
        configTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")
        Set hostDb = GetHostDb()
        If SharedDao_DatabaseTableExistsSafe(hostDb, configTableName) Then
            If LookupMetadataReferenceUsageInDatabase(hostDb, "HostDb", lookupTableName, lookupCodeFieldName, lookupRowValues, noRelationshipCoveredFields, usageDetail, configTableName) Then
                LookupRecordReferenceUsage = True
                LogPhysicalTraceMessageOnly "Info", "LookupEditorDelete", "LookupRecordReferenceUsage", "DONE", "Lookup table: " & lookupTableName & vbCrLf & "Referenced: True" & vbCrLf & usageDetail, "Verify lookup record references", 0, vbNullString, vbNullString, lookupTableName, "Table", "HostDb", "LookupEditor"
                Exit Function
            End If
        End If
    End If
    LogPhysicalTraceMessageOnly "Info", "LookupEditorDelete", "LookupRecordReferenceUsage", "DONE", "Lookup table: " & lookupTableName & vbCrLf & "Referenced: False", "Verify lookup record references", 0, vbNullString, vbNullString, lookupTableName, "Table", "CodeDb", "LookupEditor"
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    LogPhysicalTraceMessageOnly "Error", "LookupEditorDelete", "LookupRecordReferenceUsage", "ERROR", "Lookup table: " & lookupTableName, "Verify lookup record references", savedErrNumber, savedErrSource, savedErrDescription, lookupTableName, "Table", "CodeDb", "LookupEditor"
    Err.Raise savedErrNumber, "LookupRecordReferenceUsage", "Could not verify whether the lookup record is referenced." & vbCrLf & "Lookup table: " & lookupTableName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

' Purpose: Repairs canonical relationships for setup.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub RepairCanonicalRelationshipsForSetup()
    'B5 is the sole Setup mutation owner for the complete canonical CodeDb relationship set. Earlier phases own their required table, field, index, and metadata endpoints but do not create or repair relationships.
    'Health Review may validate the complete set read-only.
    RequireSetupDomainMutationAllowed SETUP_DOMAIN_B5_CATALOG_SUPPORT, "RepairCanonicalRelationshipsForSetup"
    Dim db As DAO.Database
    Dim Definition As Variant
    Set db = TablesWorkspaceCodeDb()
    db.Relations.Refresh
    RetireObsoleteConfigCandidateRelationshipsForSetup db
    For Each Definition In CanonicalRelationshipDefinitions()
        If StrComp(CStr(Definition(5)), "CodeDb", vbBinaryCompare) = 0 Then EnsureCanonicalRelationship db, Definition, True
    Next Definition
End Sub


Private Sub RetireObsoleteConfigCandidateRelationshipsForSetup(ByVal db As DAO.Database)
    'B5-only migration cleanup for relationships whose child fields moved from MC_Records to MC_Candidates.
    Dim relationName As Variant
    Dim obsoleteRelations As Variant
    obsoleteRelations = Array( _
        RELATION_CONFIG_CANDIDATE_OWNER_DATABASE_ROLE, _
        RELATION_PREFIX & "ADC_Config_" & FIELD_ROW_KIND, _
        RELATION_PREFIX & "ADC_Config_" & FIELD_CANDIDATE_STATUS, _
        RELATION_PREFIX & "ADC_Config_" & FIELD_ACCEPTANCE_STATUS, _
        RELATION_PREFIX & "ADC_Config_" & FIELD_CANDIDATE_OWNER_DATABASE_ROLE, _
        RELATION_PREFIX & "ADC_Config_" & FIELD_ACCEPTED_CONFIGURATION_ID, _
        RELATION_PREFIX & "ADC_HealthReview_" & FIELD_ROW_KIND)
    For Each relationName In obsoleteRelations
        On Error Resume Next
        db.Relations.Delete CStr(relationName)
        Err.Clear
        On Error GoTo 0
    Next relationName
    db.Relations.Refresh
End Sub
' Purpose: Adds canonical relationship findings for findings.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   findings (Input/output; Collection): Collection containing the findings items that the routine consumes or updates as part of this operation.
' Returns: None (Sub). Results are returned through the documented ByRef parameters and/or the routine's database, UI, or module-state side effects.
Public Sub AddCanonicalRelationshipFindings(ByRef findings As Collection)
    'Health Review consumer of the exact same canonical relationship definitions used by Setup.
    Dim db As DAO.Database
    Dim hostDb As DAO.Database
    Dim definitions As Collection
    Dim Definition As Variant
    Dim existingRelationship As DAO.Relation
    Set db = TablesWorkspaceCodeDb()
    db.Relations.Refresh
    Set definitions = CanonicalRelationshipDefinitions()
    For Each Definition In definitions
        Set existingRelationship = Nothing
        If StrComp(CStr(Definition(5)), "CodeDb", vbBinaryCompare) = 0 Then
            If TryGetRelationByName(db, CStr(Definition(0)), existingRelationship) = False Then
                findings.Add "Missing required canonical relationship: " & RelationshipDefinitionDisplay(Definition)
            ElseIf RelationshipMatchesDefinition(existingRelationship, Definition) = False Then
                findings.Add "Canonical relationship has the wrong table, field order, RI, or cascade definition: " & RelationshipDefinitionDisplay(Definition)
            End If
        End If
    Next Definition
    If CodeDbIsHostDb() = False Then
        Set hostDb = GetHostDb()
        hostDb.Relations.Refresh
        For Each Definition In definitions
            Set existingRelationship = Nothing
            If StrComp(CStr(Definition(5)), "CodeDb", vbBinaryCompare) = 0 Then
                If TryGetRelationByName(hostDb, CStr(Definition(0)), existingRelationship) Then findings.Add "Canonical relationship is installed in the wrong database scope (HostDb; expected CodeDb): " & CStr(Definition(0))
            End If
        Next Definition
    End If
    Set existingRelationship = Nothing
End Sub


Private Function TryGetRelationByName(ByVal db As DAO.Database, ByVal relationName As String, ByRef rel As DAO.Relation) As Boolean
    'Canonical relationship callers synchronize the DAO.Relations collection at their repair/validation boundary; this lookup then treats absence as an ordinary False result while allowing genuine collection
    'failures to propagate.
    Dim candidate As DAO.Relation
    Set rel = Nothing
    For Each candidate In db.Relations
        If StrComp(candidate.Name, relationName, vbTextCompare) = 0 Then
            Set rel = candidate
            TryGetRelationByName = True
            Exit Function
        End If
    Next candidate
End Function

Private Sub EnsureCanonicalRelationship(ByVal db As DAO.Database, ByVal Definition As Variant, Optional ByVal relationshipsAlreadyRefreshed As Boolean = False)
    On Error GoTo ErrHandler
    Dim rel As DAO.Relation
    Dim relField As DAO.Field
    Dim parentFields As Variant
    Dim childFields As Variant
    Dim i As Long
    Dim orphanCount As Long
    Dim relationName As String
    Dim deleteErrNumber As Long
    Dim deleteErrSource As String
    Dim deleteErrDescription As String
    Dim relationshipFound As Boolean
    Dim relationshipMatches As Boolean
    Dim timingStarted As Single
    Dim timingElapsed As Single
    relationName = CStr(Definition(0))
    If relationshipsAlreadyRefreshed = False Then
        LogPhysicalTraceMessageOnly "Info", "RelationshipTiming", "EnsureCanonicalRelationship", "DAO.Relations.Refresh before inspection starting.", "Relationship: " & relationName, "Canonical relationship repair", 0, vbNullString, vbNullString, relationName, "Relationship", "CodeDb", "Performance"
        timingStarted = Timer
        db.Relations.Refresh
        timingElapsed = ElapsedTimerSeconds(timingStarted)
        LogPhysicalTraceMessageOnly "Info", "RelationshipTiming", "EnsureCanonicalRelationship", "DAO.Relations.Refresh before inspection completed.", "Relationship: " & relationName & vbCrLf & "Elapsed: " & Format$(timingElapsed, "0.000") & " seconds", "Canonical relationship repair", 0, vbNullString, vbNullString, relationName, "Relationship", "CodeDb", "Performance"
    End If
    If StrComp(CStr(Definition(5)), "CodeDb", vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 9211, "EnsureCanonicalRelationship", "Unsupported canonical relationship database scope: " & CStr(Definition(5))
    If SharedDao_TableExists(db, CStr(Definition(1))) = False Then Err.Raise vbObjectError + 9212, "EnsureCanonicalRelationship", "Canonical relationship parent table is missing." & vbCrLf & RelationshipDefinitionDisplay(Definition)
    If SharedDao_TableExists(db, CStr(Definition(3))) = False Then Err.Raise vbObjectError + 9213, "EnsureCanonicalRelationship", "Canonical relationship child table is missing." & vbCrLf & RelationshipDefinitionDisplay(Definition)
    LogPhysicalTraceMessageOnly "Info", "RelationshipTiming", "EnsureCanonicalRelationship", "Existing relationship lookup starting.", "Relationship: " & relationName, "Canonical relationship repair", 0, vbNullString, vbNullString, relationName, "Relationship", "CodeDb", "Performance"
    timingStarted = Timer
    relationshipFound = TryGetRelationByName(db, relationName, rel)
    timingElapsed = ElapsedTimerSeconds(timingStarted)
    LogPhysicalTraceMessageOnly "Info", "RelationshipTiming", "EnsureCanonicalRelationship", "Existing relationship lookup completed.", "Relationship: " & relationName & vbCrLf & "Elapsed: " & Format$(timingElapsed, "0.000") & " seconds" & vbCrLf & "Found: " & CStr(relationshipFound), "Canonical relationship repair", 0, vbNullString, vbNullString, relationName, "Relationship", "CodeDb", "Performance"
    If relationshipFound Then
        LogPhysicalTraceMessageOnly "Info", "RelationshipTiming", "EnsureCanonicalRelationship", "Existing relationship definition comparison starting.", "Relationship: " & relationName, "Canonical relationship repair", 0, vbNullString, vbNullString, relationName, "Relationship", "CodeDb", "Performance"
        timingStarted = Timer
        relationshipMatches = RelationshipMatchesDefinition(rel, Definition)
        timingElapsed = ElapsedTimerSeconds(timingStarted)
        LogPhysicalTraceMessageOnly "Info", "RelationshipTiming", "EnsureCanonicalRelationship", "Existing relationship definition comparison completed.", "Relationship: " & relationName & vbCrLf & "Elapsed: " & Format$(timingElapsed, "0.000") & " seconds" & vbCrLf & "Matches: " & CStr(relationshipMatches), "Canonical relationship repair", 0, vbNullString, vbNullString, relationName, "Relationship", "CodeDb", "Performance"
        If relationshipMatches Then Exit Sub
    End If
    LogPhysicalTraceMessageOnly "Info", "RelationshipTiming", "EnsureCanonicalRelationship", "Relationship orphan-row count starting.", "Relationship: " & relationName, "Canonical relationship repair", 0, vbNullString, vbNullString, relationName, "Relationship", "CodeDb", "Performance"
    timingStarted = Timer
    orphanCount = RelationshipOrphanCountFromDefinition(db, Definition)
    timingElapsed = ElapsedTimerSeconds(timingStarted)
    LogPhysicalTraceMessageOnly "Info", "RelationshipTiming", "EnsureCanonicalRelationship", "Relationship orphan-row count completed.", "Relationship: " & relationName & vbCrLf & "Elapsed: " & Format$(timingElapsed, "0.000") & " seconds" & vbCrLf & "Orphan rows: " & CStr(orphanCount), "Canonical relationship repair", 0, vbNullString, vbNullString, relationName, "Relationship", "CodeDb", "Performance"
    If orphanCount > 0 Then Err.Raise vbObjectError + 9214, "EnsureCanonicalRelationship", "Referential integrity cannot be enabled because orphan child rows exist." & vbCrLf & RelationshipDefinitionDisplay(Definition) & vbCrLf & "Orphan rows: " & CStr(orphanCount)
    If StrComp(CStr(Definition(1)), ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_TraceLog"), vbTextCompare) = 0 Or _
       StrComp(CStr(Definition(3)), ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_TraceLog"), vbTextCompare) = 0 Then
        LogPhysicalTraceMessageOnly "Info", "RelationshipTiming", "EnsureCanonicalRelationship", "Retained TraceLog write-recordset release starting.", "Relationship: " & relationName, "Canonical relationship repair", 0, vbNullString, vbNullString, relationName, "Relationship", "CodeDb", "Performance"
        timingStarted = Timer
        CloseRetainedTraceWriteRecordset
        timingElapsed = ElapsedTimerSeconds(timingStarted)
        LogPhysicalTraceMessageOnly "Info", "RelationshipTiming", "EnsureCanonicalRelationship", "Retained TraceLog write-recordset release completed.", "Relationship: " & relationName & vbCrLf & "Elapsed: " & Format$(timingElapsed, "0.000") & " seconds", "Canonical relationship repair", 0, vbNullString, vbNullString, relationName, "Relationship", "CodeDb", "Performance"
    End If
    If Not rel Is Nothing Then
        Set rel = Nothing
        LogPhysicalTraceMessageOnly "Info", "RelationshipTiming", "EnsureCanonicalRelationship", "DAO.Relations.Delete starting.", "Relationship: " & relationName, "Canonical relationship repair", 0, vbNullString, vbNullString, relationName, "Relationship", "CodeDb", "Performance"
        timingStarted = Timer
        'A relationship removed through another DAO.Database handle can remain represented by stale collection metadata even after it was observed above. Treat DAO's two not-found results as already absent, then
        'synchronize before recreation.
        On Error Resume Next
        Err.Clear
        db.Relations.Delete relationName
        deleteErrNumber = Err.Number
        deleteErrSource = Err.Source
        deleteErrDescription = Err.Description
        Err.Clear
        On Error GoTo ErrHandler
        timingElapsed = ElapsedTimerSeconds(timingStarted)
        LogPhysicalTraceMessageOnly "Info", "RelationshipTiming", "EnsureCanonicalRelationship", "DAO.Relations.Delete completed.", "Relationship: " & relationName & vbCrLf & "Elapsed: " & Format$(timingElapsed, "0.000") & " seconds" & vbCrLf & "DAO error: " & CStr(deleteErrNumber), "Canonical relationship repair", 0, vbNullString, vbNullString, relationName, "Relationship", "CodeDb", "Performance"
        If deleteErrNumber <> 0 And deleteErrNumber <> 3011 And deleteErrNumber <> 3265 Then Err.Raise deleteErrNumber, deleteErrSource, deleteErrDescription
        LogPhysicalTraceMessageOnly "Info", "RelationshipTiming", "EnsureCanonicalRelationship", "DAO.Relations.Refresh after delete starting.", "Relationship: " & relationName, "Canonical relationship repair", 0, vbNullString, vbNullString, relationName, "Relationship", "CodeDb", "Performance"
        timingStarted = Timer
        db.Relations.Refresh
        timingElapsed = ElapsedTimerSeconds(timingStarted)
        LogPhysicalTraceMessageOnly "Info", "RelationshipTiming", "EnsureCanonicalRelationship", "DAO.Relations.Refresh after delete completed.", "Relationship: " & relationName & vbCrLf & "Elapsed: " & Format$(timingElapsed, "0.000") & " seconds", "Canonical relationship repair", 0, vbNullString, vbNullString, relationName, "Relationship", "CodeDb", "Performance"
    End If
    LogPhysicalTraceMessageOnly "Info", "RelationshipTiming", "EnsureCanonicalRelationship", "DAO.CreateRelation and field mapping starting.", "Relationship: " & relationName, "Canonical relationship repair", 0, vbNullString, vbNullString, relationName, "Relationship", "CodeDb", "Performance"
    timingStarted = Timer
    Set rel = db.CreateRelation(relationName, CStr(Definition(1)), CStr(Definition(3)), IIf(CBool(Definition(6)), 0, dbRelationDontEnforce))
    parentFields = Split(CStr(Definition(2)), "|")
    childFields = Split(CStr(Definition(4)), "|")
    If UBound(parentFields) - LBound(parentFields) <> UBound(childFields) - LBound(childFields) Then Err.Raise vbObjectError + 9215, "EnsureCanonicalRelationship", "Canonical relationship field counts do not match." & vbCrLf & RelationshipDefinitionDisplay(Definition)
    For i = LBound(parentFields) To UBound(parentFields)
        Set relField = rel.CreateField(CStr(parentFields(i)))
        relField.ForeignName = CStr(childFields(i))
        rel.fields.Append relField
    Next i
    If CBool(Definition(7)) Then rel.Attributes = rel.Attributes Or dbRelationUpdateCascade
    If CBool(Definition(8)) Then rel.Attributes = rel.Attributes Or dbRelationDeleteCascade
    timingElapsed = ElapsedTimerSeconds(timingStarted)
    LogPhysicalTraceMessageOnly "Info", "RelationshipTiming", "EnsureCanonicalRelationship", "DAO.CreateRelation and field mapping completed.", "Relationship: " & relationName & vbCrLf & "Elapsed: " & Format$(timingElapsed, "0.000") & " seconds", "Canonical relationship repair", 0, vbNullString, vbNullString, relationName, "Relationship", "CodeDb", "Performance"
    LogPhysicalTraceMessageOnly "Info", "RelationshipTiming", "EnsureCanonicalRelationship", "DAO.Relations.Append starting.", "Relationship: " & relationName, "Canonical relationship repair", 0, vbNullString, vbNullString, relationName, "Relationship", "CodeDb", "Performance"
    timingStarted = Timer
    db.Relations.Append rel
    timingElapsed = ElapsedTimerSeconds(timingStarted)
    LogPhysicalTraceMessageOnly "Info", "RelationshipTiming", "EnsureCanonicalRelationship", "DAO.Relations.Append completed.", "Relationship: " & relationName & vbCrLf & "Elapsed: " & Format$(timingElapsed, "0.000") & " seconds", "Canonical relationship repair", 0, vbNullString, vbNullString, relationName, "Relationship", "CodeDb", "Performance"
    LogPhysicalTraceMessageOnly "Info", "RelationshipTiming", "EnsureCanonicalRelationship", "DAO.Relations.Refresh after append starting.", "Relationship: " & relationName, "Canonical relationship repair", 0, vbNullString, vbNullString, relationName, "Relationship", "CodeDb", "Performance"
    timingStarted = Timer
    db.Relations.Refresh
    timingElapsed = ElapsedTimerSeconds(timingStarted)
    LogPhysicalTraceMessageOnly "Info", "RelationshipTiming", "EnsureCanonicalRelationship", "DAO.Relations.Refresh after append completed.", "Relationship: " & relationName & vbCrLf & "Elapsed: " & Format$(timingElapsed, "0.000") & " seconds", "Canonical relationship repair", 0, vbNullString, vbNullString, relationName, "Relationship", "CodeDb", "Performance"
    Set rel = Nothing
    LogPhysicalTraceMessageOnly "Info", "RelationshipTiming", "EnsureCanonicalRelationship", "Post-append relationship lookup starting.", "Relationship: " & relationName, "Canonical relationship repair", 0, vbNullString, vbNullString, relationName, "Relationship", "CodeDb", "Performance"
    timingStarted = Timer
    relationshipFound = TryGetRelationByName(db, relationName, rel)
    timingElapsed = ElapsedTimerSeconds(timingStarted)
    LogPhysicalTraceMessageOnly "Info", "RelationshipTiming", "EnsureCanonicalRelationship", "Post-append relationship lookup completed.", "Relationship: " & relationName & vbCrLf & "Elapsed: " & Format$(timingElapsed, "0.000") & " seconds" & vbCrLf & "Found: " & CStr(relationshipFound), "Canonical relationship repair", 0, vbNullString, vbNullString, relationName, "Relationship", "CodeDb", "Performance"
    If relationshipFound = False Then Err.Raise vbObjectError + 9216, "EnsureCanonicalRelationship", "Canonical relationship was not present after creation."
    LogPhysicalTraceMessageOnly "Info", "RelationshipTiming", "EnsureCanonicalRelationship", "Post-append relationship definition verification starting.", "Relationship: " & relationName, "Canonical relationship repair", 0, vbNullString, vbNullString, relationName, "Relationship", "CodeDb", "Performance"
    timingStarted = Timer
    relationshipMatches = RelationshipMatchesDefinition(rel, Definition)
    timingElapsed = ElapsedTimerSeconds(timingStarted)
    LogPhysicalTraceMessageOnly "Info", "RelationshipTiming", "EnsureCanonicalRelationship", "Post-append relationship definition verification completed.", "Relationship: " & relationName & vbCrLf & "Elapsed: " & Format$(timingElapsed, "0.000") & " seconds" & vbCrLf & "Matches: " & CStr(relationshipMatches), "Canonical relationship repair", 0, vbNullString, vbNullString, relationName, "Relationship", "CodeDb", "Performance"
    If relationshipMatches = False Then Err.Raise vbObjectError + 9217, "EnsureCanonicalRelationship", "Canonical relationship did not match its definition after creation."
    CoreAddOperationNote "Created or repaired non-cascading RI relationship [" & relationName & "]."
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "EnsureCanonicalRelationship", "Could not create or verify canonical ADC relationship." & vbCrLf & RelationshipDefinitionDisplay(Definition) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Function CanonicalRelationshipDescription(ByVal Definition As Variant) As String
    Dim enforcementText As String
    enforcementText = IIf(CBool(Definition(6)), "enforces referential integrity", "documents a non-enforced association")
    CanonicalRelationshipDescription = Left$("ADC-owned relationship from [" & CStr(Definition(1)) & "].[" & Replace(CStr(Definition(2)), "|", ", ") & "] to [" & CStr(Definition(3)) & "].[" & Replace(CStr(Definition(4)), "|", ", ") & "]; " & enforcementText & ". Cascade update=" & CStr(CBool(Definition(7))) & "; cascade delete=" & CStr(CBool(Definition(8))) & ".", 255)
End Function

Private Function RelationshipMatchesDefinition(ByVal rel As DAO.Relation, ByVal Definition As Variant) As Boolean
    On Error GoTo ExitHere
    Dim parentFields As Variant
    Dim childFields As Variant
    Dim i As Long
    If StrComp(rel.Table, CStr(Definition(1)), vbTextCompare) <> 0 Then Exit Function
    If StrComp(rel.foreignTable, CStr(Definition(3)), vbTextCompare) <> 0 Then Exit Function
    parentFields = Split(CStr(Definition(2)), "|")
    childFields = Split(CStr(Definition(4)), "|")
    If rel.fields.count <> UBound(parentFields) - LBound(parentFields) + 1 Then Exit Function
    If UBound(parentFields) - LBound(parentFields) <> UBound(childFields) - LBound(childFields) Then Exit Function
    For i = LBound(parentFields) To UBound(parentFields)
        If StrComp(rel.fields(i - LBound(parentFields)).Name, CStr(parentFields(i)), vbTextCompare) <> 0 Then Exit Function
        If StrComp(rel.fields(i - LBound(parentFields)).ForeignName, CStr(childFields(i)), vbTextCompare) <> 0 Then Exit Function
    Next i
    If CBool(Definition(6)) Then
        If (rel.Attributes And dbRelationDontEnforce) <> 0 Then Exit Function
    ElseIf (rel.Attributes And dbRelationDontEnforce) = 0 Then
        Exit Function
    End If
    If ((rel.Attributes And dbRelationUpdateCascade) <> 0) <> CBool(Definition(7)) Then Exit Function
    If ((rel.Attributes And dbRelationDeleteCascade) <> 0) <> CBool(Definition(8)) Then Exit Function
    RelationshipMatchesDefinition = True
ExitHere:
End Function

Private Function RelationshipOrphanCountFromDefinition(ByVal db As DAO.Database, ByVal Definition As Variant) As Long
    Dim rs As DAO.Recordset
    Dim parentFields As Variant
    Dim childFields As Variant
    Dim joinText As String
    Dim childRequiredText As String
    Dim sqlText As String
    Dim i As Long
    parentFields = Split(CStr(Definition(2)), "|")
    childFields = Split(CStr(Definition(4)), "|")
    For i = LBound(parentFields) To UBound(parentFields)
        If Len(joinText) > 0 Then joinText = joinText & " AND "
        joinText = joinText & "c.[" & CStr(childFields(i)) & "]=p.[" & CStr(parentFields(i)) & "]"
        If Len(childRequiredText) > 0 Then childRequiredText = childRequiredText & " AND "
        childRequiredText = childRequiredText & "c.[" & CStr(childFields(i)) & "] Is Not Null"
    Next i
    sqlText = "SELECT Count(*) AS " & FIELD_CNT & " FROM [" & CStr(Definition(3)) & "] AS c LEFT JOIN [" & CStr(Definition(1)) & "] AS p ON " & joinText & " WHERE " & childRequiredText & " AND p.[" & CStr(parentFields(LBound(parentFields))) & "] Is Null;"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    RelationshipOrphanCountFromDefinition = CLng(rs.fields(FIELD_CNT).value)
    rs.Close
End Function

Private Function RelationshipDefinitionDisplay(ByVal Definition As Variant) As String
    RelationshipDefinitionDisplay = CStr(Definition(0)) & " | " & CStr(Definition(5)) & " | " & CStr(Definition(1)) & ".(" & Replace(CStr(Definition(2)), "|", ",") & ") -> " & CStr(Definition(3)) & ".(" & Replace(CStr(Definition(4)), "|", ",") & ")" & vbCrLf & "Developer description: " & CanonicalRelationshipDescription(Definition)
End Function

Private Function RepairSummaryContainsTable(ByVal repairSummary As String, ByVal tableName As String) As Boolean
    Dim normalizedSummary As String
    normalizedSummary = vbCrLf & Replace(repairSummary, vbCr, vbNullString)
    normalizedSummary = Replace(normalizedSummary, vbLf, vbCrLf)
    RepairSummaryContainsTable = (InStr(1, normalizedSummary, vbCrLf & tableName & ":", vbTextCompare) > 0)
End Function

Private Function UniqueBackupTableName(ByVal sourceTableName As String, ByVal backupStamp As String) As String
    Dim prefixText As String
    Dim sourcePart As String
    Dim candidateName As String
    Dim suffixText As String
    Dim counter As Long
    Dim maxSourceLength As Long
    prefixText = "ADCBak_" & backupStamp & "_"
    maxSourceLength = MAX_ACCESS_OBJECT_NAME_LENGTH - Len(prefixText)
    If maxSourceLength < 8 Then maxSourceLength = 8
    sourcePart = SharedName_SafeAccessObjectName(sourceTableName, 0, "DB_", vbNullString)
    If Len(sourcePart) = 0 Then sourcePart = "ADC"
    sourcePart = Left$(sourcePart, maxSourceLength)
    candidateName = Left$(prefixText & sourcePart, MAX_ACCESS_OBJECT_NAME_LENGTH)
    Do While CodeDbTableExists(candidateName)
        PumpSetupControlEvents
        counter = counter + 1
        suffixText = "_" & Format$(counter, "00")
        candidateName = Left$(prefixText & sourcePart, MAX_ACCESS_OBJECT_NAME_LENGTH - Len(suffixText)) & suffixText
    Loop
    UniqueBackupTableName = candidateName
End Function

Private Sub AddTerminologyTableSchemaIssues(ByRef issues As Collection, Optional ByVal suppliedSchemaDefinitions As Variant, Optional ByVal suppliedIndexDefinitions As Variant)
    Dim indexDefinition As Variant
    Dim schemaDefinitions As Collection
    Dim indexDefinitions As Collection
    If isMissing(suppliedSchemaDefinitions) = False Then
        If IsObject(suppliedSchemaDefinitions) Then Set schemaDefinitions = suppliedSchemaDefinitions
    End If
    If isMissing(suppliedIndexDefinitions) = False Then
        If IsObject(suppliedIndexDefinitions) Then Set indexDefinitions = suppliedIndexDefinitions
    End If
    If schemaDefinitions Is Nothing Then Set schemaDefinitions = TerminologyFieldSchemaDefinitions()
    If indexDefinitions Is Nothing Then Set indexDefinitions = TerminologyIndexSchemaDefinitions()
    If CodeDbTableExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_TERMINOLOGY_TABLE_LOGICAL_ID)) = False Then
        issues.Add ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_TERMINOLOGY_TABLE_LOGICAL_ID) & ": table is missing"
        Exit Sub
    End If
    If TableCanOpen(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_TERMINOLOGY_TABLE_LOGICAL_ID), True) = False Then
        issues.Add ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_TERMINOLOGY_TABLE_LOGICAL_ID) & ": table exists but cannot be opened"
        Exit Sub
    End If
    AddFieldDefinitionIssuesFromSchema issues, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_TERMINOLOGY_TABLE_LOGICAL_ID), schemaDefinitions
    For Each indexDefinition In indexDefinitions
        AddIndexDefinitionIssue issues, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_TERMINOLOGY_TABLE_LOGICAL_ID), CStr(indexDefinition(0)), CStr(indexDefinition(1)), CBool(indexDefinition(2)), CBool(indexDefinition(3))
    Next indexDefinition
End Sub

Private Sub AddEditableLookupTableSchemaIssues(ByRef issues As Collection, ByVal valueTextSizes As Object)
    'A2 Preflight and post-repair proof must cover every persisted editable lookup-table structure that A2 can modify.
    Dim definitions As Collection
    Dim rowData As Variant
    Dim tableName As String
    Dim valueTextSize As Long
    Dim schemaDefinitions As Collection
    Set definitions = EditableLookupTableDefinitionsForSetup()
    For Each rowData In definitions
        tableName = Trim$(CStr(rowData(0)))
        If Len(tableName) = 0 Then
            issues.Add "Editable lookup table: canonical table name is blank"
        ElseIf CodeDbTableExists(tableName) = False Then
            issues.Add tableName & ": table is missing"
        ElseIf TableCanOpen(tableName, True) = False Then
            issues.Add tableName & ": table exists but cannot be opened"
        Else
            valueTextSize = CLng(valueTextSizes(tableName))
            Set schemaDefinitions = New Collection
            schemaDefinitions.Add Array(FIELD_VALUE_TEXT, dbText, valueTextSize, False, True, False, Null, Null, Null)
            schemaDefinitions.Add Array(FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER, dbLong, 0, False, True)
            schemaDefinitions.Add Array(FIELD_DESCRIPTION, dbMemo, 0, False, Null, True)
            AddFieldDefinitionIssuesFromSchema issues, tableName, schemaDefinitions
            AddIndexDefinitionIssue issues, tableName, INDEX_PRIMARY_KEY, FIELD_VALUE_TEXT, True, True
            AddIndexDefinitionIssue issues, tableName, INDEX_IX_SORT_ORDER, FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER, False, False
        End If
    Next rowData
End Sub

Private Sub AddConfigTableSchemaIssues(ByRef issues As Collection, Optional ByVal suppliedSchemaDefinitions As Variant, Optional ByVal suppliedIndexDefinitions As Variant)
    Dim indexDefinition As Variant
    Dim tableName As String
    Dim schemaDefinitions As Collection
    Dim indexDefinitions As Collection
    Dim tdf As DAO.tableDef
    If isMissing(suppliedSchemaDefinitions) = False Then
        If IsObject(suppliedSchemaDefinitions) Then Set schemaDefinitions = suppliedSchemaDefinitions
    End If
    If isMissing(suppliedIndexDefinitions) = False Then
        If IsObject(suppliedIndexDefinitions) Then Set indexDefinitions = suppliedIndexDefinitions
    End If
    If schemaDefinitions Is Nothing Then Set schemaDefinitions = ConfigFieldSchemaDefinitions()
    If indexDefinitions Is Nothing Then Set indexDefinitions = ConfigIndexSchemaDefinitions()
    tableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")
    If CodeDbTableExists(tableName) = False Then Exit Sub
    If TableCanOpen(tableName, True) = False Then
        issues.Add tableName & ": table definition exists but Access cannot open it; it will be treated as stale/broken and recreated"
        Exit Sub
    End If
    Set tdf = TablesWorkspaceCodeDb().TableDefs(tableName)
    AddFieldDefinitionIssuesFromSchema issues, tableName, schemaDefinitions
    For Each indexDefinition In indexDefinitions
        If ConfigIndexSupersededByCodeDbCanonicalRelationship(indexDefinition) Then
            If TableDefIndexExists(tdf, CStr(indexDefinition(0))) Then issues.Add tableName & ": redundant relationship-backed explicit index [" & CStr(indexDefinition(0)) & "] is present and must be removed before B5 canonical RI owns the equivalent CodeDb child-field index"
        Else
            AddIndexDefinitionIssue issues, tableName, CStr(indexDefinition(0)), CStr(indexDefinition(1)), CBool(indexDefinition(2)), CBool(indexDefinition(3))
        End If
    Next indexDefinition
End Sub


Private Sub AddHealthReviewTableSchemaIssues(ByRef issues As Collection, Optional ByVal suppliedSchemaDefinitions As Variant)
    'A2 Preflight and final proof cover every HealthReview field/index that A2 can repair, including the per-run IssueHash unique identity.
    Dim tableName As String
    Dim schemaDefinitions As Collection
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    If isMissing(suppliedSchemaDefinitions) = False Then
        If IsObject(suppliedSchemaDefinitions) Then Set schemaDefinitions = suppliedSchemaDefinitions
    End If
    If schemaDefinitions Is Nothing Then Set schemaDefinitions = BuildHealthReviewSchemaDefinitions()
    tableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReview")
    If CodeDbTableExists(tableName) = False Then issues.Add tableName & ": table is missing": Exit Sub
    If TableCanOpen(tableName, True) = False Then issues.Add tableName & ": table exists but cannot be opened": Exit Sub
    Set db = TablesWorkspaceCodeDb()
    If mTableDesignBatchPrepared = False Then db.TableDefs.Refresh
    Set tdf = db.TableDefs(tableName)
    If mTableDesignBatchPrepared = False Then tdf.fields.Refresh
    If mTableDesignBatchPrepared = False Then tdf.Indexes.Refresh
    AddFieldDefinitionIssuesFromSchema issues, tableName, schemaDefinitions, tdf
    AddIndexDefinitionIssue issues, tableName, INDEX_PRIMARY_KEY, FIELD_ID, True, True, tdf
    AddIndexDefinitionIssue issues, tableName, INDEX_IX_REVIEW_RUN_ID, FIELD_REVIEW_RUN_ID, False, False, tdf
    AddIndexDefinitionIssue issues, tableName, INDEX_IX_IS_CURRENT_RUN, FIELD_IS_CURRENT_RUN, False, False, tdf
    AddIndexDefinitionIssue issues, tableName, INDEX_IX_SEVERITY, FIELD_SEVERITY, False, False, tdf
    AddIndexDefinitionIssue issues, tableName, INDEX_IX_REVIEW_AREA, FIELD_REVIEW_AREA, False, False, tdf
    AddIndexDefinitionIssue issues, tableName, INDEX_IX_ISSUE_CODE, FIELD_ISSUE_CODE, False, False, tdf
    AddIndexDefinitionIssue issues, tableName, INDEX_IX_CONFIG_ID, FIELD_CONFIG_ID, False, False, tdf
    AddIndexDefinitionIssue issues, tableName, INDEX_IX_MACRO_KEY, FIELD_MACRO_KEY, False, False, tdf
    AddIndexDefinitionIssue issues, tableName, INDEX_IX_SETTING_KEY, FIELD_SETTING_KEY, False, False, tdf
    AddIndexDefinitionIssue issues, tableName, INDEX_IX_OBJECT_NAME, FIELD_OBJECT_NAME, False, False, tdf
    AddIndexDefinitionIssue issues, tableName, INDEX_IX_ISSUE_HASH, FIELD_ISSUE_HASH, False, False, tdf
    If TableDefIndexExists(tdf, INDEX_UX_HEALTH_REVIEW_RUN_ISSUE_HASH) = False Then
        issues.Add tableName & ": missing composite index [" & INDEX_UX_HEALTH_REVIEW_RUN_ISSUE_HASH & "] on fields [" & FIELD_REVIEW_RUN_ID & ", " & FIELD_ISSUE_HASH & "]"
    ElseIf IndexMatchesFieldList(tdf.Indexes(INDEX_UX_HEALTH_REVIEW_RUN_ISSUE_HASH), FIELD_REVIEW_RUN_ID & "|" & FIELD_ISSUE_HASH, True, False) = False Then
        issues.Add tableName & ": invalid composite index [" & INDEX_UX_HEALTH_REVIEW_RUN_ISSUE_HASH & "] on fields [" & FIELD_REVIEW_RUN_ID & ", " & FIELD_ISSUE_HASH & "]"
    End If
End Sub

Private Sub AddLauncherReviewTableSchemaIssues(ByRef issues As Collection, Optional ByVal suppliedSchemaDefinitions As Variant)
    'A2 Preflight and final proof cover the launcher-review physical schema and its AutoNumber row primary key before A2 can repair either contract.
    Dim tableName As String
    Dim schemaDefinitions As Collection
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    If isMissing(suppliedSchemaDefinitions) = False Then
        If IsObject(suppliedSchemaDefinitions) Then Set schemaDefinitions = suppliedSchemaDefinitions
    End If
    If schemaDefinitions Is Nothing Then Set schemaDefinitions = CustomMacroLauncherReviewSchemaDefinitions()
    tableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LauncherControlsReview")
    If CodeDbTableExists(tableName) = False Then issues.Add tableName & ": table is missing": Exit Sub
    If TableCanOpen(tableName, True) = False Then issues.Add tableName & ": table exists but cannot be opened": Exit Sub
    Set db = TablesWorkspaceCodeDb()
    If mTableDesignBatchPrepared = False Then db.TableDefs.Refresh
    Set tdf = db.TableDefs(tableName)
    If mTableDesignBatchPrepared = False Then tdf.fields.Refresh
    If mTableDesignBatchPrepared = False Then tdf.Indexes.Refresh
    AddFieldDefinitionIssuesFromSchema issues, tableName, schemaDefinitions, tdf
    AddIndexDefinitionIssue issues, tableName, INDEX_PRIMARY_KEY, FIELD_ID, True, True, tdf
End Sub

Private Sub AddCandidateTableSchemaIssues(ByRef issues As Collection, Optional ByVal suppliedSchemaDefinitions As Variant, Optional ByVal suppliedIndexDefinitions As Variant)
    Dim indexDefinition As Variant
    Dim tableName As String
    Dim schemaDefinitions As Collection
    Dim indexDefinitions As Collection
    Dim tdf As DAO.tableDef
    If isMissing(suppliedSchemaDefinitions) = False Then
        If IsObject(suppliedSchemaDefinitions) Then Set schemaDefinitions = suppliedSchemaDefinitions
    End If
    If isMissing(suppliedIndexDefinitions) = False Then
        If IsObject(suppliedIndexDefinitions) Then Set indexDefinitions = suppliedIndexDefinitions
    End If
    If schemaDefinitions Is Nothing Then Set schemaDefinitions = CandidateFieldSchemaDefinitions()
    If indexDefinitions Is Nothing Then Set indexDefinitions = CandidateIndexSchemaDefinitions()
    tableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates")
    If CodeDbTableExists(tableName) = False Then
        issues.Add tableName & ": table is missing"
        Exit Sub
    End If
    If TableCanOpen(tableName, True) = False Then
        issues.Add tableName & ": table definition exists but Access cannot open it; it will be treated as stale/broken and recreated"
        Exit Sub
    End If
    Set tdf = TablesWorkspaceCodeDb().TableDefs(tableName)
    AddFieldDefinitionIssuesFromSchema issues, tableName, schemaDefinitions
    For Each indexDefinition In indexDefinitions
        If ConfigIndexSupersededByCodeDbCanonicalRelationship(indexDefinition) Then
            If TableDefIndexExists(tdf, CStr(indexDefinition(0))) Then issues.Add tableName & ": redundant relationship-backed explicit index [" & CStr(indexDefinition(0)) & "] is present and must be removed before B5 canonical RI owns the equivalent CodeDb child-field index"
        Else
            AddIndexDefinitionIssue issues, tableName, CStr(indexDefinition(0)), CStr(indexDefinition(1)), CBool(indexDefinition(2)), CBool(indexDefinition(3))
        End If
    Next indexDefinition
End Sub
Private Function BuildMacroCategorySchemaDefinitions() As Collection
    'Single expected physical schema authority for MC_Cats repair and validation.
    ' CANONICAL AUTHORITY: rows owns the in-memory definition set constructed by BuildMacroCategorySchemaDefinitions; downstream consumers derive views from this set rather than create a competing authority.
    Dim rows As New Collection
    rows.Add Array("FIELD", FIELD_ID, dbLong, 0, True, False, vbNullString)
    rows.Add Array("FIELD", FIELD_MACRO_CATEGORY_KEY, dbText, 100, False, False, vbNullString, True, False, Null, Null)
    rows.Add Array("FIELD", FIELD_MACRO_CATEGORY_NAME, dbText, 255, False, False, vbNullString, True, False, Null, Null)
    rows.Add Array("FIELD", FIELD_CATEGORY_DESCRIPTION, dbMemo, 0, False, False, vbNullString)
    rows.Add Array("FIELD", FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER, dbLong, 0, False, True, DEFAULT_MACRO_CATEGORY_SORT_ORDER_FIELD_VALUE, Null, Null, Null, Null)
    rows.Add Array("FIELD", FIELD_ACTIVE, dbBoolean, 0, False, True, DEFAULT_MACRO_CATEGORY_ACTIVE_FIELD_VALUE, Null, Null, Null, Null)
    rows.Add Array("FIELD", FIELD_MANAGED_OBJECT_DEPENDENCY_NOTES, dbMemo, 0, False, False, vbNullString)
    rows.Add Array("INDEX", INDEX_PRIMARY_KEY, FIELD_ID, True, True)
    rows.Add Array("INDEX", "ux_MacroCategoryKey", FIELD_MACRO_CATEGORY_KEY, True, False)
    rows.Add Array("INDEX", "ix_MacroCategoryName", FIELD_MACRO_CATEGORY_NAME, False, False)
    rows.Add Array("INDEX", "ix_MacroCategorySortOrder", FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER, False, False)
    rows.Add Array("INDEX", "ix_MacroCategoryActive", FIELD_ACTIVE, False, False)
    Set BuildMacroCategorySchemaDefinitions = rows
End Function

Private Function BuildMacroSubcategorySchemaDefinitions() As Collection
    'Single expected physical schema authority for MC_Cats_SubCats repair and validation.
    ' CANONICAL AUTHORITY: rows owns the in-memory definition set constructed by BuildMacroSubcategorySchemaDefinitions; downstream consumers derive views from this set rather than create a competing authority.
    Dim rows As New Collection
    rows.Add Array("FIELD", FIELD_ID, dbLong, 0, True, False, vbNullString)
    rows.Add Array("FIELD", FIELD_MACRO_CATEGORY_KEY, dbText, 100, False, False, vbNullString, True, False, Null, Null)
    rows.Add Array("FIELD", FIELD_MACRO_SUBCATEGORY_KEY, dbText, 100, False, False, vbNullString, True, False, Null, Null)
    rows.Add Array("FIELD", FIELD_MACRO_SUBCATEGORY_NAME, dbText, 255, False, False, vbNullString, True, False, Null, Null)
    rows.Add Array("FIELD", FIELD_SUBCATEGORY_DESCRIPTION, dbMemo, 0, False, False, vbNullString)
    rows.Add Array("FIELD", FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER, dbLong, 0, False, True, DEFAULT_MACRO_SUBCATEGORY_SORT_ORDER_FIELD_VALUE, Null, Null, Null, Null)
    rows.Add Array("FIELD", FIELD_ACTIVE, dbBoolean, 0, False, True, DEFAULT_MACRO_SUBCATEGORY_ACTIVE_FIELD_VALUE, Null, Null, Null, Null)
    rows.Add Array("FIELD", FIELD_MANAGED_OBJECT_DEPENDENCY_NOTES, dbMemo, 0, False, False, vbNullString)
    rows.Add Array("INDEX", INDEX_PRIMARY_KEY, FIELD_ID, True, True)
    rows.Add Array("INDEX", "ux_MacroSubcategoryKey", FIELD_MACRO_SUBCATEGORY_KEY, True, False)
    rows.Add Array("MULTI_INDEX", "ux_MacroSubcategoryCategoryKey", FIELD_MACRO_CATEGORY_KEY & "|" & FIELD_MACRO_SUBCATEGORY_KEY, True, False)
    rows.Add Array("INDEX", "ix_MacroSubcategoryCategory", FIELD_MACRO_CATEGORY_KEY, False, False)
    rows.Add Array("INDEX", "ix_MacroSubcategoryName", FIELD_MACRO_SUBCATEGORY_NAME, False, False)
    rows.Add Array("INDEX", "ix_MacroSubcategorySortOrder", FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER, False, False)
    rows.Add Array("INDEX", "ix_MacroSubcategoryActive", FIELD_ACTIVE, False, False)
    Set BuildMacroSubcategorySchemaDefinitions = rows
End Function

Private Sub AddMacroClassificationSchemaIssues(ByRef issues As Collection, ByVal tableName As String, ByVal schemaDefinitions As Collection)
    Dim schemaDefinition As Variant
    If CodeDbTableExists(tableName) = False Then
        issues.Add tableName & ": table is missing"
        Exit Sub
    End If
    If TableCanOpen(tableName, True) = False Then
        issues.Add tableName & ": table exists but cannot be opened"
        Exit Sub
    End If
    For Each schemaDefinition In schemaDefinitions
        Select Case CStr(schemaDefinition(0))
            Case "FIELD"
                AddFieldDefinitionIssue issues, tableName, CStr(schemaDefinition(1)), CLng(schemaDefinition(2)), CLng(schemaDefinition(3)), CBool(schemaDefinition(4))
                AddMacroClassificationPhysicalFieldIssue issues, tableName, schemaDefinition
            Case "INDEX"
                AddIndexDefinitionIssue issues, tableName, CStr(schemaDefinition(1)), CStr(schemaDefinition(2)), CBool(schemaDefinition(3)), CBool(schemaDefinition(4))
            Case "MULTI_INDEX"
                AddMultiFieldIndexDefinitionIssue issues, tableName, CStr(schemaDefinition(1)), CStr(schemaDefinition(2)), CBool(schemaDefinition(3)), CBool(schemaDefinition(4))
        End Select
    Next schemaDefinition
    AddUnexpectedMacroClassificationFieldIssues issues, tableName, schemaDefinitions
End Sub

Private Sub AddMacroClassificationPhysicalFieldIssue(ByRef issues As Collection, ByVal tableName As String, ByVal schemaDefinition As Variant)
    If UBound(schemaDefinition) < 7 Then Exit Sub
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim fld As DAO.Field
    Dim contractDefinition As Variant
    Set db = TablesWorkspaceCodeDb()
    Set tdf = db.TableDefs(tableName)
    If SharedDao_TableDefFieldExistsSafe(tdf, CStr(schemaDefinition(1))) = False Then Exit Sub
    Set fld = tdf.fields(CStr(schemaDefinition(1)))
    contractDefinition = Array(schemaDefinition(1), schemaDefinition(2), schemaDefinition(3), schemaDefinition(4), schemaDefinition(7), schemaDefinition(8), IIf(CBool(schemaDefinition(5)), schemaDefinition(6), Null), schemaDefinition(9), schemaDefinition(10))
    Dim issueText As String
    issueText = SchemaFieldContractIssueText(fld, contractDefinition)
    If Len(issueText) > 0 Then issues.Add tableName & ": field [" & CStr(schemaDefinition(1)) & "] physical contract mismatch; " & issueText
End Sub

Private Sub ApplyMacroClassificationPhysicalFieldContract(ByVal db As DAO.Database, ByVal tableName As String, ByVal schemaDefinition As Variant)
    If UBound(schemaDefinition) < 7 Then Exit Sub
    Dim tdf As DAO.tableDef
    Dim contractDefinition As Variant
    Set tdf = db.TableDefs(tableName)
    contractDefinition = Array(schemaDefinition(1), schemaDefinition(2), schemaDefinition(3), schemaDefinition(4), schemaDefinition(7), schemaDefinition(8), IIf(CBool(schemaDefinition(5)), schemaDefinition(6), Null), schemaDefinition(9), schemaDefinition(10))
    ApplySchemaFieldContract tdf.fields(CStr(schemaDefinition(1))), contractDefinition, tableName, "macro-classification"
End Sub

Private Sub AddUnexpectedMacroClassificationFieldIssues(ByRef issues As Collection, ByVal tableName As String, ByVal schemaDefinitions As Collection)
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim fld As DAO.Field
    Set db = TablesWorkspaceCodeDb()
    Set tdf = db.TableDefs(tableName)
    For Each fld In tdf.fields
        If MacroClassificationSchemaDefinesField(schemaDefinitions, fld.Name) = False Then issues.Add tableName & "." & fld.Name & ": unexpected field is not part of the canonical ADC-owned schema"
    Next fld
End Sub

Private Function MacroClassificationSchemaDefinesField(ByVal schemaDefinitions As Collection, ByVal fieldName As String) As Boolean
    Dim schemaDefinition As Variant
    fieldName = Trim$(fieldName)
    For Each schemaDefinition In schemaDefinitions
        If StrComp(CStr(schemaDefinition(0)), "FIELD", vbBinaryCompare) = 0 Then
            If StrComp(CStr(schemaDefinition(1)), fieldName, vbTextCompare) = 0 Then
                MacroClassificationSchemaDefinesField = True
                Exit Function
            End If
        End If
    Next schemaDefinition
End Function

Private Sub AddMacroClassificationTableSchemaIssues(ByRef issues As Collection, Optional ByVal suppliedCategorySchemaDefinitions As Variant, Optional ByVal suppliedSubcategorySchemaDefinitions As Variant)
    Dim categorySchemaDefinitions As Collection
    Dim subcategorySchemaDefinitions As Collection
    If isMissing(suppliedCategorySchemaDefinitions) = False Then
        If IsObject(suppliedCategorySchemaDefinitions) Then Set categorySchemaDefinitions = suppliedCategorySchemaDefinitions
    End If
    If isMissing(suppliedSubcategorySchemaDefinitions) = False Then
        If IsObject(suppliedSubcategorySchemaDefinitions) Then Set subcategorySchemaDefinitions = suppliedSubcategorySchemaDefinitions
    End If
    If categorySchemaDefinitions Is Nothing Then Set categorySchemaDefinitions = BuildMacroCategorySchemaDefinitions()
    If subcategorySchemaDefinitions Is Nothing Then Set subcategorySchemaDefinitions = BuildMacroSubcategorySchemaDefinitions()
    AddMacroClassificationSchemaIssues issues, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_MacroCategories"), categorySchemaDefinitions
    AddMacroClassificationSchemaIssues issues, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_MacroSubcategories"), subcategorySchemaDefinitions
End Sub

Private Function ConfigTableSchemaIssueText(ByVal schemaDefinitions As Collection, ByVal indexDefinitions As Collection) As String
    Dim issues As Collection
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: issues holds planned/recovery work or independently observed validation evidence for ConfigTableSchemaIssueText; keep it distinct from the source authority.
    Set issues = New Collection
    If CodeDbTableExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")) = False Then
        ConfigTableSchemaIssueText = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & ": table is missing"
        Exit Function
    End If
    If TableCanOpen(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config"), True) = False Then
        ConfigTableSchemaIssueText = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & ": table exists but cannot be opened"
        Exit Function
    End If
    AddConfigTableSchemaIssues issues, schemaDefinitions, indexDefinitions
    If issues.count > 0 Then ConfigTableSchemaIssueText = SharedCollection_Join(issues, vbCrLf)
End Function


Private Function CandidateTableSchemaIssueText(ByVal schemaDefinitions As Collection, ByVal indexDefinitions As Collection) As String
    Dim issues As Collection
    Set issues = New Collection
    AddCandidateTableSchemaIssues issues, schemaDefinitions, indexDefinitions
    If issues.count > 0 Then CandidateTableSchemaIssueText = SharedCollection_Join(issues, vbCrLf)
End Function
Private Sub PrepareOpenFoundationTablesForDesignChange(ByVal actionText As String)
    Dim names As Collection
    Set names = FoundationRequiredTableNames()
    If SharedAccess_PrepareOpenObjectsForDesignChange(acTable, names, actionText) = False Then
        MarkSystemCommandCancelled
        Err.Raise vbObjectError + 8994, "PrepareOpenFoundationTablesForDesignChange", "A2 table design preparation was cancelled."
    End If
End Sub

Private Function ReacquireTableDefField(ByVal tableName As String, ByVal fieldName As String, ByRef ownerDatabase As DAO.Database) As DAO.Field
    On Error GoTo ErrHandler
    Set ownerDatabase = currentDb
    ownerDatabase.TableDefs.Refresh
    Set ReacquireTableDefField = ownerDatabase.TableDefs(tableName).fields(fieldName)
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ReacquireTableDefField", "Could not reacquire the field for post-write Description verification." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

' Purpose: Resolves safe DAO table def name from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   tdf (Input; DAO.tableDef): DAO TableDef representing the target table whose schema or properties are inspected or changed.
' Returns: String: Returns the resolved name produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function SafeDaoTableDefName(ByVal tdf As DAO.tableDef) As String
    On Error GoTo ErrHandler
    SafeDaoTableDefName = tdf.Name
    Exit Function
ErrHandler:
    SafeDaoTableDefName = "<unknown table>"
End Function

Private Function ClearFieldLookupProperties(ByVal tdf As DAO.tableDef, ByVal fieldName As String) As Long
    On Error GoTo ErrHandler
    If SetupDomainMutationWritesAllowedForCurrentPart(SETUP_DOMAIN_D4_LOOKUP_EDITOR_BINDING) = False Then Err.Raise vbObjectError + 9621, "ClearFieldLookupProperties", "Full lookup-property cleanup includes D4-owned editor binding properties and may run during numbered Setup only in D4." & vbCrLf & "Table: " & SafeDaoTableDefName(tdf) & vbCrLf & "Field: " & fieldName
    Dim fld As DAO.Field
    Dim propertyNames As Variant
    Dim propertyName As Variant
    Dim prp As DAO.Property
    Dim changed As Boolean
    If SharedDao_TableDefFieldExistsSafe(tdf, fieldName) = False Then Exit Function
    Set fld = tdf.fields(fieldName)
    propertyNames = Array("DisplayControl", "RowSourceType", "RowSource", "BoundColumn", "ColumnCount", "ColumnWidths", "ColumnHeads", "ListRows", "ListWidth", "LimitToList", "AllowValueListEdits", "ListItemsEditForm")
    For Each propertyName In propertyNames
        Set prp = Nothing
        On Error Resume Next
        Err.Clear
        Set prp = fld.properties(CStr(propertyName))
        If Err.Number = 0 And Not prp Is Nothing Then
            fld.properties.Delete CStr(propertyName)
            If Err.Number = 0 Then
                changed = True
            ElseIf Err.Number <> 3265 And Err.Number <> 3270 Then
                GoTo DeleteFailed
            End If
        ElseIf Err.Number <> 0 And Err.Number <> 3265 And Err.Number <> 3270 Then
            GoTo DeleteFailed
        End If
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
            Err.Raise vbObjectError + 6496, "ClearFieldLookupProperties", "A stale lookup property remained after cleanup." & vbCrLf & "Table: " & tdf.Name & vbCrLf & "Field: " & fieldName & vbCrLf & "Property: " & CStr(propertyName)
        End If
        If Err.Number <> 0 And Err.Number <> 3265 And Err.Number <> 3270 Then GoTo VerifyFailed
        Err.Clear
        On Error GoTo ErrHandler
    Next propertyName
    If changed Then ClearFieldLookupProperties = 1
    Exit Function
DeleteFailed:
    Dim deleteErrNumber As Long
    Dim deleteErrSource As String
    Dim deleteErrDescription As String
    deleteErrNumber = Err.Number: deleteErrSource = Err.Source: deleteErrDescription = Err.Description
    On Error GoTo ErrHandler
    Err.Raise deleteErrNumber, deleteErrSource, deleteErrDescription
VerifyFailed:
    Dim verifyErrNumber As Long
    Dim verifyErrSource As String
    Dim verifyErrDescription As String
    verifyErrNumber = Err.Number: verifyErrSource = Err.Source: verifyErrDescription = Err.Description
    On Error GoTo ErrHandler
    Err.Raise verifyErrNumber, verifyErrSource, verifyErrDescription
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ClearFieldLookupProperties", "Could not clear and verify stale lookup metadata." & vbCrLf & "Table: " & SafeDaoTableDefName(tdf) & vbCrLf & "Field: " & fieldName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

' Purpose: Resolves clear field lookup presentation properties from the supplied/current ADC state, using `SharedDao_TableDefFieldExistsSafe` and `ElseIf` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   tdf (Input; DAO.tableDef): DAO TableDef representing the target table whose schema or properties are inspected or changed.
'   fieldName (Input; String): Canonical or physical name identifying the field targeted by this operation.
' Returns: Long: Returns the numeric clear field lookup presentation properties result; zero and other sentinel values retain the meanings established by the routine's existing implementation.
Public Function ClearFieldLookupPresentationProperties(ByVal tdf As DAO.tableDef, ByVal fieldName As String) As Long
    'B2 presentation-only cleanup. The form-dependent AllowValueListEdits and ListItemsEditForm properties are deliberately excluded because D4 owns their positive and negative Setup state.
    On Error GoTo ErrHandler
    Dim fld As DAO.Field
    Dim propertyNames As Variant
    Dim propertyName As Variant
    Dim prp As DAO.Property
    Dim changed As Boolean
    If SharedDao_TableDefFieldExistsSafe(tdf, fieldName) = False Then Exit Function
    Set fld = tdf.fields(fieldName)
    propertyNames = Array("DisplayControl", "RowSourceType", "RowSource", "BoundColumn", "ColumnCount", "ColumnWidths", "ColumnHeads", "ListRows", "ListWidth", "LimitToList")
    For Each propertyName In propertyNames
        Set prp = Nothing
        On Error Resume Next
        Err.Clear
        Set prp = fld.properties(CStr(propertyName))
        If Err.Number = 0 And Not prp Is Nothing Then
            fld.properties.Delete CStr(propertyName)
            If Err.Number = 0 Then
                changed = True
            ElseIf Err.Number <> 3265 And Err.Number <> 3270 Then
                GoTo DeleteFailed
            End If
        ElseIf Err.Number <> 0 And Err.Number <> 3265 And Err.Number <> 3270 Then
            GoTo DeleteFailed
        End If
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
            Err.Raise vbObjectError + 6497, "ClearFieldLookupPresentationProperties", "A stale B2-owned lookup presentation property remained after cleanup." & vbCrLf & "Table: " & tdf.Name & vbCrLf & "Field: " & fieldName & vbCrLf & "Property: " & CStr(propertyName)
        End If
        If Err.Number <> 0 And Err.Number <> 3265 And Err.Number <> 3270 Then GoTo VerifyFailed
        Err.Clear
        On Error GoTo ErrHandler
    Next propertyName
    If changed Then ClearFieldLookupPresentationProperties = 1
    Exit Function
DeleteFailed:
    Dim deleteErrNumber As Long
    Dim deleteErrSource As String
    Dim deleteErrDescription As String
    deleteErrNumber = Err.Number: deleteErrSource = Err.Source: deleteErrDescription = Err.Description
    On Error GoTo ErrHandler
    Err.Raise deleteErrNumber, deleteErrSource, deleteErrDescription
VerifyFailed:
    Dim verifyErrNumber As Long
    Dim verifyErrSource As String
    Dim verifyErrDescription As String
    verifyErrNumber = Err.Number: verifyErrSource = Err.Source: verifyErrDescription = Err.Description
    On Error GoTo ErrHandler
    Err.Raise verifyErrNumber, verifyErrSource, verifyErrDescription
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ClearFieldLookupPresentationProperties", "Could not clear and verify stale B2-owned lookup presentation metadata." & vbCrLf & "Table: " & SafeDaoTableDefName(tdf) & vbCrLf & "Field: " & fieldName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

' Purpose: Ensures field table lookup for tdf, field name.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   tdf (Input; DAO.tableDef): DAO.TableDef representing the table whose fields, indexes, properties, or metadata are being inspected or modified.
'   fieldName (Input; String): Field name identifying the table/query field or bound control source the routine evaluates or modifies.
'   lookupTableName (Input; String): Physical Access lookup-table name whose rows/fields supply the lookup definition or validation target.
'   codeFieldName (Input; String): Field name containing the stored lookup code/key value used as the bound value for the generated lookup editor.
'   displayFieldName (Input; String): Field name containing the user-facing lookup text shown by the generated lookup editor/list.
'   LimitToList (Input Optional; Boolean): True restricts the Access lookup control/field to values present in its row source; False permits values outside the listed choices when Access allows them. Default:
'   True.
'   columnWidthsText (Input Optional; String): Input column widths text for ensure field table lookup; is tested by If Len(columnWidthsText) = 0. Default: "0;2".
' Returns: Long: Returns the number/result code produced while applying the requested field/lookup metadata; callers use it to detect whether physical metadata required creation or repair.
Public Function EnsureFieldTableLookup(ByVal tdf As DAO.tableDef, ByVal fieldName As String, ByVal lookupTableName As String, ByVal codeFieldName As String, ByVal displayFieldName As String, Optional ByVal LimitToList As Boolean = True, Optional ByVal columnWidthsText As String = "0;2") As Long
    On Error GoTo ErrHandler
    Dim fld As DAO.Field
    Dim rowSourceText As String
    Dim changed As Boolean
    If SharedDao_TableDefFieldExistsSafe(tdf, fieldName) = False Then Exit Function
    lookupTableName = Trim$(lookupTableName)
    codeFieldName = Trim$(codeFieldName)
    displayFieldName = Trim$(displayFieldName)
    columnWidthsText = Trim$(columnWidthsText)
    If Len(columnWidthsText) = 0 Then columnWidthsText = "0;2"
    columnWidthsText = NormalizeLookupColumnWidthsForAccess(columnWidthsText, 2, 1, LimitToList)
    If Len(lookupTableName) = 0 Then Exit Function
    If Len(codeFieldName) = 0 Then Exit Function
    If Len(displayFieldName) = 0 Then Exit Function
    If CodeDbTableExists(lookupTableName) = False Then Exit Function
    If TableFieldExists(lookupTableName, codeFieldName) = False Then Exit Function
    If TableFieldExists(lookupTableName, displayFieldName) = False Then Exit Function
    Set fld = tdf.fields(fieldName)
    rowSourceText = TableLookupRowSource(lookupTableName, codeFieldName, displayFieldName, TableFieldExists(lookupTableName, FIELD_LOOKUP_ACTIVE), StrComp(codeFieldName, FIELD_SETTING_KEY, vbTextCompare) = 0, TableFieldExists(lookupTableName, FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER))
    changed = ApplyLookupFieldProperties(fld, "Table/Query", rowSourceText, LimitToList, 2, 1, columnWidthsText, "EnsureFieldTableLookup", vbObjectError + 6494, vbObjectError + 6495)
    If changed Then EnsureFieldTableLookup = 1
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If savedErrNumber = 0 Then
        CoreAddOperationNote "Skipped table-field lookup metadata because the property update did not complete, but no active VBA error remained for diagnosis. Table=" & SafeDaoTableDefName(tdf) & "; Field=" & fieldName & "; LookupTable=" & lookupTableName & "."
    Else
        CoreAddOperationNote "Skipped table-field lookup metadata after Access rejected the lookup property update. Table=" & SafeDaoTableDefName(tdf) & "; Field=" & fieldName & "; LookupTable=" & lookupTableName & "; Error=" & CStr(savedErrNumber) & " from " & savedErrSource & " - " & savedErrDescription
    End If
    Err.Clear
End Function

Private Sub RepairExplicitBoundedTextFieldInDatabase(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String, ByVal fieldSize As Long, Optional ByVal prepareCodeDbDesignChange As Boolean = False)
    'Repairs only fields carrying an explicit ReqFieldSize contract. It verifies existing data length before changing the physical field definition.
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim maxStoredLength As Long
    If fieldSize < 1 Or fieldSize > MAX_ACCESS_SHORT_TEXT_LENGTH Then Err.Raise vbObjectError + 5252, "RepairExplicitBoundedTextFieldInDatabase", "Required field size is outside the valid Access Short Text range."
    Set rs = db.OpenRecordset("SELECT Max(Len(Nz([" & fieldName & "],''))) AS MaxStoredLength FROM [" & tableName & "];", dbOpenSnapshot, dbReadOnly)
    If rs.EOF = False Then maxStoredLength = CLng(Nz(rs.fields(0).value, 0))
    rs.Close
    Set rs = Nothing
    If maxStoredLength > fieldSize Then Err.Raise vbObjectError + 5253, "RepairExplicitBoundedTextFieldInDatabase", "Existing data exceeds the required field size; no data was truncated and the field was not changed." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName & vbCrLf & "Longest stored value: " & CStr(maxStoredLength) & vbCrLf & "ReqFieldSize: " & CStr(fieldSize) & vbCrLf & "Default/value source: (not applicable; existing persisted row)"
    If prepareCodeDbDesignChange And mTableDesignBatchPrepared = False Then
        If SharedAccess_PrepareOpenObjectForDesignChange(acTable, tableName, "applying ReqFieldSize to field [" & fieldName & "] in table [" & tableName & "]") = False Then
            MarkSystemCommandCancelled
            Err.Raise vbObjectError + 9003, "RepairExplicitBoundedTextFieldInDatabase", "Field-size repair was cancelled because the affected open table was not approved for save and close."
        End If
    End If
    db.Execute "ALTER TABLE [" & tableName & "] ALTER COLUMN [" & fieldName & "] TEXT(" & CStr(fieldSize) & ");", dbFailOnError
    db.TableDefs.Refresh
    CoreAddOperationNote "Applied ReqFieldSize " & CStr(fieldSize) & " to field [" & fieldName & "] in table [" & tableName & "]."
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
    Err.Raise savedErrNumber, "RepairExplicitBoundedTextFieldInDatabase", "Could not safely apply the required field size." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName & vbCrLf & "ReqFieldSize: " & CStr(fieldSize) & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub RepairTableField(ByVal tableName As String, ByVal fieldName As String, ByVal fieldType As DAO.DataTypeEnum, Optional ByVal fieldSize As Long = 0, Optional ByVal newFieldDefaultValue As Variant)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim fld As DAO.Field
    Dim originalErrNumber As Long
    Dim originalErrSource As String
    Dim originalErrDescription As String
    Set db = TablesWorkspaceCodeDb()
    If mTableDesignBatchPrepared = False Then db.TableDefs.Refresh
    Set tdf = db.TableDefs(tableName)
    If mTableDesignBatchPrepared = False Then tdf.fields.Refresh
    If SharedDao_TableDefFieldExistsSafe(tdf, fieldName) Then
        Set fld = tdf.fields(fieldName)
        If SharedDao_FieldDefinitionMatches(fld, fieldType, fieldSize, False) = False Then
            If fieldType = dbText And IsNull(DefaultManagedFieldReqSize(tableName, fieldName)) = False Then
                RepairExplicitBoundedTextFieldInDatabase db, tableName, fieldName, fieldSize, True
                Set db = TablesWorkspaceCodeDb()
                Set tdf = db.TableDefs(tableName)
                Set fld = tdf.fields(fieldName)
            End If
            If SharedDao_FieldDefinitionMatches(fld, fieldType, fieldSize, False) = False Then Err.Raise vbObjectError + 5250, "RepairTableField", "Existing field has the wrong definition. Use table rebuild repair for schema changes." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName & vbCrLf & "Expected: " & FieldDefinitionDescription(fieldType, fieldSize, False) & vbCrLf & "Found: " & FieldActualDefinitionDescription(fld)
        End If
        If isMissing(newFieldDefaultValue) = False Then
            If StrComp(CStr(Nz(fld.DefaultValue, vbNullString)), CStr(newFieldDefaultValue), vbBinaryCompare) <> 0 Then fld.DefaultValue = CStr(newFieldDefaultValue)
        End If
        Exit Sub
    End If
    If fieldSize > 0 Then
        Set fld = tdf.CreateField(fieldName, fieldType, fieldSize)
    Else
        Set fld = tdf.CreateField(fieldName, fieldType)
    End If
    On Error Resume Next
    tdf.fields.Append fld
    If Err.Number = 3191 Or Err.Number = 3367 Then
        Err.Clear
        tdf.fields.Refresh
        If SharedDao_TableDefFieldExistsSafe(tdf, fieldName) Then Exit Sub
    End If
    If Err.Number <> 0 Then GoTo ErrHandler
    On Error GoTo ErrHandler
    If mTableDesignBatchPrepared = False Then tdf.fields.Refresh
    If SharedDao_TableDefFieldExistsSafe(tdf, fieldName) = False Then
        tdf.fields.Refresh
        If SharedDao_TableDefFieldExistsSafe(tdf, fieldName) = False Then Err.Raise vbObjectError + 5251, "RepairTableField", "The newly added field was not visible after DAO append." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName
    End If
    Set fld = tdf.fields(fieldName)
    If isMissing(newFieldDefaultValue) = False Then
        On Error Resume Next
        fld.DefaultValue = newFieldDefaultValue
        If Err.Number <> 0 Then
            originalErrNumber = Err.Number
            originalErrSource = Err.Source
            originalErrDescription = Err.Description
            On Error GoTo 0
            Err.Raise originalErrNumber, "RepairTableField", "Could not assign the declared default to a newly added field." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName & vbCrLf & "Err.Source: " & originalErrSource & vbCrLf & "Err.Description: " & originalErrDescription
        End If
        On Error GoTo ErrHandler
    End If
    CoreAddOperationNote "Added field [" & fieldName & "] to table [" & tableName & "]."
    Exit Sub
ErrHandler:
    originalErrNumber = Err.Number
    originalErrSource = Err.Source
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "RepairTableField", "Could not add or validate field." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

' Purpose: Returns table existence from the active Preflight metadata snapshot when one is owned by PrepareSetupDataTableBackupForPreflight. All other callers retain the live DAO existence check, so runtime
'   behavior
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   tableName (Input; String): Physical Access table name identifying the table whose schema, data, or metadata the routine processes.
' Returns: Boolean: True when code DB table exists in the inspected context; False when it does not.
Public Function CodeDbTableExists(ByVal tableName As String) As Boolean
    'Returns table existence from the active Preflight metadata snapshot when one is owned by PrepareSetupDataTableBackupForPreflight. All other callers retain the live DAO existence check, so runtime behavior
    'outside that read-only planning boundary is unchanged.
    If mSetupPreflightTableNameIndexActive Then
        If mSetupPreflightExistingTableNames Is Nothing Then Err.Raise vbObjectError + 9250, "CodeDbTableExists", "Setup Preflight marked its table-name snapshot active, but the authoritative dictionary is unavailable."
        CodeDbTableExists = mSetupPreflightExistingTableNames.Exists(tableName)
        Exit Function
    End If
    CodeDbTableExists = SharedDao_DatabaseTableExistsSafe(TablesWorkspaceCodeDb(), tableName)
End Function

' Purpose: Performs table can open by coordinating `CodeDbTableExists` and `TablesWorkspaceCodeDb` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   tableName (Input; String): Physical Access table name identifying the table whose schema, data, or metadata the routine processes.
'   tableExistsAlreadyKnown (Input Optional; Boolean): True only when the immediate caller has already established that the same table exists and no intervening mutation can invalidate that proof. Default: False.
' Returns: Boolean: True when the table can open condition is satisfied for the supplied or current ADC context; False otherwise.
Public Function TableCanOpen(ByVal tableName As String, Optional ByVal tableExistsAlreadyKnown As Boolean = False) As Boolean
    'Intended use: Returns the Access Development Console Table Can Open result for direct, query, form, or required cross-module callers.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    If tableExistsAlreadyKnown = False Then
        If CodeDbTableExists(tableName) = False Then Exit Function
    End If
    Set db = TablesWorkspaceCodeDb()
    Set rs = db.OpenRecordset(tableName, dbOpenSnapshot)
    TableCanOpen = True
ExitHere:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Err.Clear
    Exit Function
ErrHandler:
    TableCanOpen = False
    Resume ExitHere
End Function

' Purpose: Resolves table open failure detail from the supplied/current ADC state, using `CodeDbTableExists` and `SharedDao_QueryExists` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   tableName (Input; String): Canonical or physical name identifying the table targeted by this operation.
' Returns: String: Returns the table open failure detail text value. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function TableOpenFailureDetail(ByVal tableName As String) As String
    'Intended use: Returns the Access Development Console Table Open Failure Detail result for direct, query, form, or required cross-module callers.
    On Error GoTo ErrHandler
    Dim directDetail As String
    Dim sqlDetail As String
    If CodeDbTableExists(tableName) = False Then
        If SharedDao_QueryExists(TablesWorkspaceCodeDb(), tableName) Then
            TableOpenFailureDetail = "No table named [" & tableName & "] exists, but a saved query with that exact name exists. " & "That query name blocks table creation and can cause Access error 3078 while the utility is trying to create the table."
        Else
            TableOpenFailureDetail = "No table named [" & tableName & "] exists in TableDefs."
        End If
        Exit Function
    End If
    directDetail = OpenRecordsetFailureDetail(tableName, False)
    sqlDetail = OpenRecordsetFailureDetail(tableName, True)
    If Len(directDetail) = 0 And Len(sqlDetail) = 0 Then
        TableOpenFailureDetail = "No open failure detected. Direct table-name open and SQL open both succeeded."
    Else
        TableOpenFailureDetail = "Direct table-name open: " & IIf(Len(directDetail) = 0, "Succeeded", directDetail) & "; " & "SQL open: " & IIf(Len(sqlDetail) = 0, "Succeeded", sqlDetail)
    End If
    Exit Function
ErrHandler:
    TableOpenFailureDetail = "Open-failure diagnostic itself failed. Err.Number: " & CStr(Err.Number) & "; Err.Source: " & Err.Source & "; Err.Description: " & Err.Description
End Function

Private Sub VerifyTableReady(ByVal tableName As String, ByVal tableRole As String)
    If TableCanOpen(tableName) Then Exit Sub
    Err.Raise vbObjectError + 5291, "VerifyTableReady", DEFAULT_UTILITY_LONG_NAME & " " & tableRole & " table exists in the table collection but cannot be opened after create/repair." & vbCrLf & "Table: " & tableName & vbCrLf & "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB & vbCrLf & "Database: " & TablesWorkspaceCodeDb().Name & vbCrLf & "Open failure: " & TableOpenFailureDetail(tableName)
End Sub

Private Sub RefreshTableDefsSafely()
    On Error Resume Next
    Dim db As DAO.Database
    Set db = TablesWorkspaceCodeDb()
    db.TableDefs.Refresh
    db.QueryDefs.Refresh
    On Error GoTo 0
End Sub

' Purpose: Checks whether safe drop table if exists in the database, collection, or metadata scope supplied to this function.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   tableName (Input; String): Physical Access table name identifying the table whose schema, data, or metadata the routine processes.
'   successNote (Input Optional; String): Optional note emitted/recorded when the requested operation succeeds; blank leaves the normal success text unchanged. Default: vbNullString.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub SafeDropTableIfExists(ByVal tableName As String, Optional ByVal successNote As String = vbNullString)
    'Safely removes table or saved-query objects that conflict with the requested table name.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim droppedAny As Boolean
    Dim tableDropErrorText As String
    Dim queryDropErrorText As String
    If Len(tableName) = 0 Then Exit Sub
    Set db = TablesWorkspaceCodeDb()
    RefreshTableDefsSafely
    If CodeDbTableExists(tableName) Then
        If TryDeleteTableDef(db, tableName, tableDropErrorText) Then
            droppedAny = True
        End If
    End If
    If SharedDao_QueryExists(TablesWorkspaceCodeDb(), tableName) Then
        If TryDeleteQueryDef(db, tableName, queryDropErrorText) Then
            droppedAny = True
        End If
    End If
    If CodeDbIsHostDb() Then
        If AccessObjectExistsInCurrentDatabase(acTable, tableName) Then
            If TryDeleteAccessObject(acTable, tableName, tableDropErrorText) Then droppedAny = True
        End If
        If AccessObjectExistsInCurrentDatabase(acQuery, tableName) Then
            If TryDeleteAccessObject(acQuery, tableName, queryDropErrorText) Then droppedAny = True
        End If
    End If
    RefreshTableDefsSafely
    If ObjectNameConflictExists(tableName) Then
        Err.Raise vbObjectError + 5298, "SafeDropTableIfExists", "A table/query object-name conflict still exists after deletion attempts." & vbCrLf & "Object name: " & tableName & vbCrLf & "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB & vbCrLf & "Database: " & TablesWorkspaceCodeDb().Name & vbCrLf & "Table delete detail: " & IIf(Len(tableDropErrorText) = 0, "(none)", tableDropErrorText) & vbCrLf & "Query delete detail: " & IIf(Len(queryDropErrorText) = 0, "(none)", queryDropErrorText)
    End If
    If droppedAny Then
        If Len(Trim$(successNote)) > 0 Then
            CoreAddOperationNote successNote
        Else
            CoreAddOperationNote "Dropped macro-creator object-name conflict [" & tableName & "]."
        End If
    End If
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrSource As String
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrSource = Err.Source
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "SafeDropTableIfExists", "Could not safely clear macro-creator table/query object-name conflict." & vbCrLf & "Object name: " & tableName & vbCrLf & "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB & vbCrLf & "Database: " & TablesWorkspaceCodeDb().Name & vbCrLf & "Err.Number: " & CStr(originalErrNumber) & vbCrLf & "Err.Source: " & originalErrSource & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

' Purpose: Attempts to perform delete table def for DB, table name.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   db (Input; DAO.Database): DAO.Database instance whose schema, data, queries, or properties the routine reads or modifies.
'   tableName (Input; String): Physical Access table name identifying the table whose schema, data, or metadata the routine processes.
'   errorText (Output; String): Receives the resolved error text produced from this procedure's validated inputs and current ADC state.
'   errorNumberResult (Output Optional; Long): Receives the error number result produced by the operation so the caller can inspect the outcome without re-running the work. Default: 0.
'   errorSourceResult (Output Optional; String): Receives the error source result produced by the operation so the caller can inspect the outcome without re-running the work. Default: vbNullString.
'   errorDescriptionResult (Output Optional; String): Receives the error description result produced by the operation so the caller can inspect the outcome without re-running the work. Default: vbNullString.
' Returns: Boolean: True when the attempt to delete table def completes successfully; False for the routine's defined normal unsuccessful/not-applicable outcome.
Public Function TryDeleteTableDef(ByVal db As DAO.Database, ByVal tableName As String, ByRef errorText As String, Optional ByRef errorNumberResult As Long = 0, Optional ByRef errorSourceResult As String = vbNullString, Optional ByRef errorDescriptionResult As String = vbNullString) As Boolean
    On Error GoTo ErrHandler
    errorNumberResult = 0
    errorSourceResult = vbNullString
    errorDescriptionResult = vbNullString
    db.TableDefs.Delete tableName
    db.TableDefs.Refresh
    TryDeleteTableDef = True
    Exit Function
ErrHandler:
    errorNumberResult = Err.Number
    errorSourceResult = Err.Source
    errorDescriptionResult = Err.Description
    If Err.Number = 3265 Or Err.Number = 3078 Or Err.Number = 3011 Or Err.Number = 7874 Then
        errorText = "Table definition was already unavailable. Err.Number: " & CStr(Err.Number) & "; " & SharedMessage_OneLine(Err.Description, 250)
    Else
        errorText = "Could not delete TableDef. Err.Number: " & CStr(Err.Number) & "; Err.Source: " & Err.Source & "; Err.Description: " & Err.Description
    End If
    Err.Clear
End Function

' Purpose: Checks whether table field exists in the database, collection, or metadata scope supplied to this function.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   tableName (Input; String): Physical Access table name identifying the table whose schema, data, or metadata the routine processes.
'   fieldName (Input; String): Field name identifying the table/query field or bound control source the routine evaluates or modifies.
' Returns: Boolean: True when table field exists in the inspected context; False when it does not.
Public Function TableFieldExists(ByVal tableName As String, ByVal fieldName As String) As Boolean
    ' Wrapper purpose: TABLES-owned facade binding the shared DAO field-existence check to the Access Development Console table database context.
    ' Why retained: Retained so cross-module callers need not choose the DAO.Database themselves and all ADC table-field checks use the same owned context.
    'Intended use: Returns the Access Development Console Table Field Exists result for direct, query, form, or required cross-module callers.
    TableFieldExists = SharedDao_DatabaseTableFieldExistsSafe(TablesWorkspaceCodeDb(), tableName, fieldName)
End Function

' Purpose: Resolves table diagnostics location for message from the supplied/current ADC state, using `ManagedObjectPhysicalName` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function TableDiagnosticsLocationForMessage() As String
    On Error Resume Next
    TableDiagnosticsLocationForMessage = vbCrLf & vbCrLf & "Diagnostics:" & vbCrLf & "Trace table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_TraceLog") & vbCrLf & "Trace run ID: " & GetActiveOrLastTraceRunID() & vbCrLf & "Log file: " & GetTraceLogFilePath() & TraceBackendDiagnosticsForMessage()
End Function

' Purpose: Resolves table readiness diagnostics for message from the supplied/current ADC state, using `ObjectDiagnosticLine` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function TableReadinessDiagnosticsForMessage() As String
    'Intended use: Returns the Access Development Console Table Readiness Diagnostics For Message result for direct, query, form, or required cross-module callers.
    On Error GoTo ErrHandler
    TableReadinessDiagnosticsForMessage = vbCrLf & vbCrLf & DEFAULT_UTILITY_LONG_NAME & " table/object diagnostics:" & vbCrLf & ObjectDiagnosticLine(SafeTableName("settings")) & vbCrLf & ObjectDiagnosticLine(SafeTableName("ADC_Config"))
    Exit Function
ErrHandler:
    TableReadinessDiagnosticsForMessage = vbCrLf & vbCrLf & DEFAULT_UTILITY_LONG_NAME & " table/object diagnostics: diagnostic failed. " & "Err.Number: " & CStr(Err.Number) & "; Err.Description: " & Err.Description
End Function

Private Function FoundationValueText(ByVal fieldValue As Variant) As String
    If IsNull(fieldValue) Then
        FoundationValueText = "<NULL>"
    ElseIf IsDate(fieldValue) Then
        FoundationValueText = Format$(CDate(fieldValue), "yyyy-mm-dd hh:nn:ss")
    ElseIf VarType(fieldValue) = vbBoolean Then
        FoundationValueText = CStr(CBool(fieldValue))
    Else
        FoundationValueText = CStr(fieldValue)
    End If
End Function


' ===============================================================================================================================================================================
' TRACE / DIAGNOSTIC STORAGE PROCEDURES
' ===============================================================================================================================================================================

' Purpose: Initializes trace table write suppression, coordinating `End` as the implementation requires.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub BeginTraceTableWriteSuppression()
    'Temporarily suppresses TraceLog table writes for a caller-owned read-only phase while leaving status/immediate/file telemetry available.
    mTraceTableWriteSuspensionDepth = mTraceTableWriteSuspensionDepth + 1
End Sub

' Purpose: Initializes terminal trace table teardown, coordinating `End` as the implementation requires.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub BeginTerminalTraceTableTeardown()
    'Terminal destructive boundary: stop all database-backed TraceLog writes, release the retained DAO writer, and leave physical trace-file logging available.
    BeginTraceTableWriteSuppression
    CloseRetainedTraceWriteRecordset
    mTraceLogTableReadyRunID = vbNullString
End Sub

' Purpose: Closes trace table write suppression, coordinating `End` as the implementation requires.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub EndTraceTableWriteSuppression()
    If mTraceTableWriteSuspensionDepth <= 0 Then Err.Raise vbObjectError + 9141, "EndTraceTableWriteSuppression", "No TraceLog table-write suppression boundary is active."
    mTraceTableWriteSuspensionDepth = mTraceTableWriteSuspensionDepth - 1
End Sub

' Purpose: Starts trace run, coordinating `ResetSetupControlPerformanceDiagnostics` and `RecreatePhysicalTraceLogFile` as the implementation requires.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   contextName (Input Optional; String): Context label recorded with tracing/progress/errors so the caller can distinguish the logical operation producing the message. Default: vbNullString.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub StartTraceRun(Optional ByVal contextName As String = vbNullString)
    'Intended use: Runs Start Access Development Console Trace Run for direct, generated-macro, form-button, or required cross-module callers.
    If Len(mTraceRunID) > 0 Then Exit Sub
    CoreResetOperationNotes
    ClearPopupMessagesSuppressionForCurrentRun
    Dim traceLogClearSucceeded As Boolean
    Dim traceLogClearDetail As String
    mTraceTableWriteFailureCount = 0
    mTraceTableLastFailureText = vbNullString
    mTraceLogFileWriteFailureCount = 0
    mTraceLogFileLastFailureText = vbNullString
    mTraceLogFileClearFailureText = vbNullString
    mTraceLogTableReadyRunID = vbNullString
    CoreResetSharedProgressDiagnostics
    mSetupControlPumpRequestsIssuedCount = 0
    mTraceStorageVerificationCount = 0
    mTraceWriteRecordsetOpenCount = 0
    mTraceRowsWrittenCount = 0
    mTracePhysicalLogOpenCount = 0
    mTracePhysicalLogLinesWrittenCount = 0
    mTraceLogFileMaxLineLengthForRun = 0
    mTraceTableWriteSuspensionDepth = 0
    ResetNamingPerformanceDiagnostics
    On Error Resume Next
    Call ResetSetupControlPerformanceDiagnostics
    On Error GoTo 0
    CloseRetainedTraceWriteRecordset
    FlushTraceLogFile
    traceLogClearSucceeded = RecreatePhysicalTraceLogFile(traceLogClearDetail)
    Randomize
    mTraceRunID = Format$(Now, "yyyymmddhhnnss") & "-" & Right$("000000" & CStr(Int(999999 * Rnd) + 1), 6)
    mLastTraceRunID = mTraceRunID
    If traceLogClearSucceeded = False Then
        mTraceLogFileClearFailureText = traceLogClearDetail
        mTraceRunID = vbNullString
        mLastTraceRunID = vbNullString
        Err.Raise vbObjectError + 9002, "StartTraceRun", "A new Access Development Console process cannot start because its physical trace log file could not be deleted and recreated." & vbCrLf & traceLogClearDetail
    End If
    LogTraceMessage "Info", "Setup", "StartTraceRun", "Access Development Console trace run started.", traceLogClearDetail & vbCrLf & "Trace-history table retained: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_TraceLog"), vbNullString, 0, vbNullString, vbNullString, vbNullString, vbNullString, "CodeDb", contextName
    If LOG_CHATGPT_VERSION_IDS_AT_TRACE_START Then LogChatGPTModuleVersionInventory contextName
End Sub

' Purpose: Starts the Setup trace for coordinated Setup without changing setup-phase ownership.
' Public role: Public module entry point. This procedure is listed in the module header as an intentional externally callable ADC entry point and may be invoked directly by Access/VBA callers.
' Parameters:
'   contextName (Input Optional; String): Context label recorded with tracing/progress/errors so the caller can distinguish the logical operation producing the message. Default: vbNullString.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub StartTraceRunForSetup(Optional ByVal contextName As String = vbNullString)
    'Starts the Setup trace for coordinated Setup without changing setup-phase ownership.
    If Len(mTraceRunID) > 0 Then Exit Sub
    CoreResetOperationNotes
    ClearPopupMessagesSuppressionForCurrentRun
    Dim traceLogClearSucceeded As Boolean
    Dim traceLogClearDetail As String
    Dim diagnosticsReadyDetail As String
    mTraceTableWriteFailureCount = 0
    mTraceTableLastFailureText = vbNullString
    mTraceLogFileWriteFailureCount = 0
    mTraceLogFileLastFailureText = vbNullString
    mTraceLogFileClearFailureText = vbNullString
    mTraceLogTableReadyRunID = vbNullString
    CoreResetSharedProgressDiagnostics
    mSetupControlPumpRequestsIssuedCount = 0
    mTraceStorageVerificationCount = 0
    mTraceWriteRecordsetOpenCount = 0
    mTraceRowsWrittenCount = 0
    mTracePhysicalLogOpenCount = 0
    mTracePhysicalLogLinesWrittenCount = 0
    mTraceLogFileMaxLineLengthForRun = 0
    mTraceTableWriteSuspensionDepth = 0
    ResetSetupPreflightInfrastructureState
    ResetNamingPerformanceDiagnostics
    On Error Resume Next
    Call ResetSetupControlPerformanceDiagnostics
    On Error GoTo 0
    CloseRetainedTraceWriteRecordset
    FlushTraceLogFile
    traceLogClearSucceeded = RecreatePhysicalTraceLogFile(traceLogClearDetail)
    If traceLogClearSucceeded = False Then
        mTraceLogFileClearFailureText = traceLogClearDetail
        Err.Raise vbObjectError + 9002, "StartTraceRunForSetup", "A new " & DEFAULT_UTILITY_LONG_NAME & " Setup process cannot start because its physical trace log file could not be deleted and recreated." & vbCrLf & traceLogClearDetail
    End If
    ReplayRunFullSetupEntryPhysicalMarkerAfterLogRecreation
    EnsureSetupDiagnosticsReadyForPreflight diagnosticsReadyDetail
    mTraceStorageVerificationCount = mTraceStorageVerificationCount + 1
    Randomize
    mTraceRunID = Format$(Now, "yyyymmddhhnnss") & "-" & Right$("000000" & CStr(Int(999999 * Rnd) + 1), 6)
    mLastTraceRunID = mTraceRunID
    mTraceLogTableReadyRunID = mTraceRunID
    LogTraceMessage TRACE_LEVEL_INFO, "Setup", "StartTraceRunForSetup", DEFAULT_UTILITY_LONG_NAME & " Setup trace run started.", traceLogClearDetail & vbCrLf & diagnosticsReadyDetail, vbNullString, 0, vbNullString, vbNullString, vbNullString, vbNullString, "CodeDb", contextName
    If LOG_CHATGPT_VERSION_IDS_AT_TRACE_START Then LogChatGPTModuleVersionInventory contextName
End Sub

' Purpose: Finalizes trace run, coordinating `GetTraceLogFilePath` and `End` as the implementation requires.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   completeMessage (Input Optional; String): Complete message text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine. Default: vbNullString.
'   clearStatus (Input Optional; Boolean): When True, enables or requests status for this call; when False, the routine follows its normal non-forced/default path. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub CompleteTraceRun(Optional ByVal completeMessage As String = vbNullString, Optional ByVal clearStatus As Boolean = True)
    'Intended use: Runs Complete Access Development Console Trace Run for direct, generated-macro, form-button, or required cross-module callers.
    If Len(mTraceRunID) = 0 Then Exit Sub
    If Len(Trim$(completeMessage)) = 0 Then completeMessage = DEFAULT_UTILITY_LONG_NAME & " trace run completed."
    LogTraceMessage TRACE_LEVEL_INFO, "Setup", "CompleteTraceRun", completeMessage, "Trace log file: " & GetTraceLogFilePath() & vbCrLf & SetupPerformanceDiagnosticsSummary(), vbNullString, 0, vbNullString, vbNullString, vbNullString, vbNullString, "CodeDb", vbNullString
    mLastTraceRunID = mTraceRunID
    CloseRetainedTraceWriteRecordset
    mTraceTableWriteSuspensionDepth = 0
    FlushTraceLogFile
    mTraceRunID = vbNullString
    mTraceLogTableReadyRunID = vbNullString
    mTraceLogFileMaxLineLengthForRun = 0
    If clearStatus Then
        On Error Resume Next
        ClearVisibleStatusText
        On Error GoTo 0
    End If
End Sub

' Purpose: Returns active or last trace run ID.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: String: Returns the persisted/current identifier for get active or last trace run ID; vbNullString indicates that no qualifying run/item is currently available.
Public Function GetActiveOrLastTraceRunID() As String
    'Intended use: Returns the Access Development Console Trace Run ID result for direct, query, form, or required cross-module callers.
    If Len(mTraceRunID) > 0 Then
        GetActiveOrLastTraceRunID = mTraceRunID
    Else
        GetActiveOrLastTraceRunID = mLastTraceRunID
    End If
End Function

' Purpose: Returns the active trace run ID or blank when no trace run is active.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: String: Returns the persisted/current identifier for get active trace run ID; vbNullString indicates that no qualifying run/item is currently available.
Public Function GetActiveTraceRunID() As String
    ' Wrapper purpose: Read-only TABLES facade exposing the active trace run ID while returning blank when no trace run is active.
    ' Why retained: Retained so cross-module diagnostics can read trace identity without exposing TABLES' private trace-run state for direct access or mutation.
    'Returns the active trace run ID or blank when no trace run is active.
    GetActiveTraceRunID = mTraceRunID
End Function

' Purpose: Reports the current trace run active state maintained by this module.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: Boolean: True when the trace run active condition is satisfied for the supplied or current ADC context; False otherwise.
Public Function TraceRunActive() As Boolean
    'Intended use: Returns the Access Development Console Trace Run Active result for direct, query, form, or required cross-module callers.
    TraceRunActive = (Len(mTraceRunID) > 0)
End Function



' Purpose: Writes one trace event without allowing trace-table bootstrap or sink failures to recurse or disappear.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   traceLevel (Input; String): Level value selecting the trace severity, grouping, or hierarchy used by this operation.
'   traceCategory (Input; String): Category value classifying the trace item so the routine can select the correct naming, logging, or processing rules.
'   sourceProcedure (Input; String): Procedure name recorded as the source of trace, error, status, or diagnostic information emitted by this call.
'   traceMessage (Input; String): Trace message text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   detailText (Input Optional; String): Receives or supplies diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller. Default: vbNullString.
'   stepName (Input Optional; String): Setup/progress step label recorded or displayed for the current unit of work. Default: vbNullString.
'   errNumber (Input Optional; Long): Numeric identifier or ordinal for the err item processed or reported by this operation. Default: 0.
'   errSource (Input Optional; String): Error source text carried into the returned/raised diagnostic; blank allows the procedure to supply its own source when coded to do so. Default: vbNullString.
'   errDescription (Input Optional; String): Error description text carried into the returned/raised diagnostic; blank allows the procedure to supply its own description when coded to do so. Default: vbNullString.
'   objectName (Input Optional; String): Physical Access object name identifying the database object targeted by the operation. Default: vbNullString.
'   objectType (Input Optional; String): Object-type value identifying the Access/VBA object class targeted by the operation. Default: vbNullString.
'   databaseRole (Input Optional; String): Role value selecting which database ownership or database context applies to this operation. Default: vbNullString.
'   contextName (Input Optional; String): Context label recorded with tracing/progress/errors so the caller can distinguish the logical operation producing the message. Default: vbNullString.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub LogTraceMessage(ByVal traceLevel As String, ByVal traceCategory As String, ByVal sourceProcedure As String, ByVal traceMessage As String, Optional ByVal detailText As String = vbNullString, Optional ByVal stepName As String = vbNullString, Optional ByVal errNumber As Long = 0, Optional ByVal errSource As String = vbNullString, Optional ByVal errDescription As String = vbNullString, Optional ByVal objectName As String = vbNullString, Optional ByVal objectType As String = vbNullString, Optional ByVal databaseRole As String = vbNullString, Optional ByVal contextName As String = vbNullString)
    'Writes one trace event without allowing trace-table bootstrap or sink failures to recurse or disappear.
    On Error GoTo ErrHandler
    Dim printedToImmediate As Boolean
    Dim traceRunIDForWrite As String
    Dim immediateText As String
    Dim nestedTrace As Boolean
    Dim tableSinkSuspended As Boolean
    Dim tableFailure As TTraceSinkFailure
    Dim fileFailure As TTraceSinkFailure
    mTraceSinkDepth = mTraceSinkDepth + 1
    nestedTrace = (mTraceSinkDepth > 1)
    tableSinkSuspended = (mTraceTableWriteSuspensionDepth > 0)
    traceRunIDForWrite = ResolveTraceRunIDForWrite()
    immediateText = TraceImmediateText(traceLevel, traceCategory, sourceProcedure, traceMessage, objectName)
    If nestedTrace Then
        printedToImmediate = True
        Debug.Print "ADCTraceNested: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " - " & immediateText & " [trace-table sink skipped at depth " & CStr(mTraceSinkDepth) & " to prevent recursive trace logging]"
    Else
        printedToImmediate = TraceShouldPrint(traceLevel, traceMessage, errNumber, errDescription)
        If printedToImmediate Then Debug.Print "ADCTrace: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " - " & immediateText
        If tableSinkSuspended = False Then
            If WriteTraceLogRow(traceRunIDForWrite, traceLevel, traceCategory, sourceProcedure, stepName, traceMessage, detailText, errNumber, errSource, errDescription, objectName, objectType, databaseRole, contextName, printedToImmediate, tableFailure) = False Then ReportTraceTableWriteFailure tableFailure, sourceProcedure, traceMessage
        End If
    End If
    If AppendTraceLogFileLine(traceRunIDForWrite, traceLevel, traceCategory, sourceProcedure, stepName, traceMessage, detailText, errNumber, errSource, errDescription, objectName, objectType, databaseRole, contextName, printedToImmediate, fileFailure) = False Then ReportTraceLogFileWriteFailure fileFailure, sourceProcedure, traceMessage, nestedTrace
ExitHere:
    If mTraceSinkDepth > 0 Then mTraceSinkDepth = mTraceSinkDepth - 1
    Exit Sub
ErrHandler:
    Dim sinkErrNumber As Long
    Dim sinkErrSource As String
    Dim sinkErrDescription As String
    sinkErrNumber = Err.Number
    sinkErrSource = Err.Source
    sinkErrDescription = Err.Description
    Debug.Print "ADCTraceSinkFailure: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " - Trace orchestration failed without suppressing the error. EventSource=" & sourceProcedure & "; EventMessage=" & traceMessage & "; Err.Number=" & CStr(sinkErrNumber) & "; Err.Source=" & sinkErrSource & "; Err.Description=" & sinkErrDescription
    Resume ExitHere
End Sub

' Purpose: Records Run_RunFullSetup entry in a dedicated startup log before any Setup processing and retains the invocation timestamp until the normal trace log can receive the marker.
' Public role: Inter-module startup diagnostic API used by the Core RunCode wrapper before it invokes RunFullSetup.
' Parameters: None.
' Returns: None (Sub). Startup-log failures are reported only to the Immediate window and are not raised into the caller.
Public Sub LogRunFullSetupEntryPhysicalMarker()
    'Records the wrapper-entry marker without opening, closing, counting, creating, changing, or reusing any normal trace-run sink or identity.
    mRunFullSetupEntryMarkerTimestamp = Now
    mRunFullSetupEntryMarkerPending = True
    WriteRunFullSetupEntryMarkerFile GetRunFullSetupEntryMarkerLogPath(), False
End Sub

' Purpose: Clears the retained in-memory Run_RunFullSetup marker after replay or wrapper exit; the separate cumulative startup log remains as durable evidence.
' Public role: Inter-module cleanup API used by the Core RunCode wrapper to prevent a canceled or failed invocation from leaking into a later direct Setup run.
' Parameters: None.
' Returns: None (Sub).
Public Sub ClearRunFullSetupEntryPhysicalMarker()
    'Clears only the pending in-memory pre-Setup marker; normal trace-run identity, sinks, counters, and the durable cumulative startup log are not changed.
    mRunFullSetupEntryMarkerPending = False
    mRunFullSetupEntryMarkerTimestamp = 0
End Sub

Private Sub ReplayRunFullSetupEntryPhysicalMarkerAfterLogRecreation()
    'Copies the original wrapper-entry evidence into the newly recreated normal trace log before diagnostics preflight begins, without participating in normal trace instrumentation.
    If mRunFullSetupEntryMarkerPending = False Then Exit Sub
    WriteRunFullSetupEntryMarkerFile GetTraceLogFilePath(), True
    ClearRunFullSetupEntryPhysicalMarker
End Sub

Private Function GetRunFullSetupEntryMarkerLogPath() As String
    'Places the durable cumulative startup log beside the normal physical trace log while keeping it outside that log's delete-and-recreate lifecycle.
    Dim traceLogPath As String
    Dim folderDelimiterPosition As Long
    traceLogPath = GetTraceLogFilePath()
    folderDelimiterPosition = InStrRev(traceLogPath, "\")
    If folderDelimiterPosition <= 0 Then Exit Function
    GetRunFullSetupEntryMarkerLogPath = Left$(traceLogPath, folderDelimiterPosition) & RUN_FULL_SETUP_ENTRY_LOG_FILE_NAME
End Function

Private Sub WriteRunFullSetupEntryMarkerFile(ByVal targetPath As String, ByVal markerReplay As Boolean)
    'Appends the startup marker with direct file I/O only; this routine deliberately does not call or mutate the normal trace writer, retained handle, counters, table sink, or run-ID state.
    On Error GoTo ErrHandler
    Dim fileNum As Integer
    Dim fileIsOpen As Boolean
    If Len(targetPath) = 0 Then Err.Raise vbObjectError + 9070, "WriteRunFullSetupEntryMarkerFile", "No startup diagnostic log path could be resolved."
    fileNum = FreeFile
    Open targetPath For Append As #fileNum
    fileIsOpen = True
    Print #fileNum, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | Run: pre-setup | Level: Info | Category: SetupStartup"
    Print #fileNum, "  Procedure: Run_RunFullSetup | Step: RunFullSetup wrapper entry"
    Print #fileNum, "  Message: Run_RunFullSetup() invoked before Setup processing."
    Print #fileNum, "  Detail: Original invocation: " & Format$(mRunFullSetupEntryMarkerTimestamp, "yyyy-mm-dd hh:nn:ss")
    If markerReplay Then
        Print #fileNum, "  Detail: Marker replayed after physical trace-log recreation and before Setup diagnostics preflight."
        Print #fileNum, "  DurableStartupLog: " & RUN_FULL_SETUP_ENTRY_LOG_FILE_NAME
    Else
        Print #fileNum, "  Detail: Durable startup marker remains in this file while the normal physical trace log is recreated."
    End If
    Print #fileNum, vbNullString
    Close #fileNum
    fileIsOpen = False
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If fileIsOpen Then Close #fileNum
    Debug.Print "ADCSetupStartupLogFailure: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " - Path=" & targetPath & "; Err.Number=" & CStr(savedErrNumber) & "; Err.Source=" & savedErrSource & "; Err.Description=" & savedErrDescription
    On Error GoTo 0
End Sub

' Purpose: Writes one trace event directly to the physical trace file without reading, validating, creating, or writing any trace table or persisted trace setting.
' Public role: Inter-module diagnostic API for startup boundaries that must remain observable even when database-backed trace storage is unavailable or implicated in a hang.
' Parameters: Mirrors LogTraceMessage so callers can preserve the same diagnostic metadata while selecting only the physical file sink.
'   traceRunIDOverride supplies a physical-file-only identity without reading or changing normal trace-run state.
' Returns: None (Sub). File-write failures are reported only to the Immediate window and are not raised into the caller.
Public Sub LogPhysicalTraceMessageOnly(ByVal traceLevel As String, ByVal traceCategory As String, ByVal sourceProcedure As String, ByVal traceMessage As String, Optional ByVal detailText As String = vbNullString, Optional ByVal stepName As String = vbNullString, Optional ByVal errNumber As Long = 0, Optional ByVal errSource As String = vbNullString, Optional ByVal errDescription As String = vbNullString, Optional ByVal objectName As String = vbNullString, Optional ByVal objectType As String = vbNullString, Optional ByVal databaseRole As String = vbNullString, Optional ByVal contextName As String = vbNullString, Optional ByVal traceRunIDOverride As String = vbNullString)
    'Writes and closes one physical diagnostic event using only in-memory trace state, the code-project path, and the DEFAULTS line-length constant.
    On Error GoTo ErrHandler
    Dim fileNum As Integer
    Dim fileIsOpen As Boolean
    Dim logPath As String
    Dim traceRunIDForWrite As String
    FlushTraceLogFile
    logPath = GetTraceLogFilePath()
    If Len(Trim$(logPath)) = 0 Then Err.Raise vbObjectError + 9069, "LogPhysicalTraceMessageOnly", "No physical trace-log path could be resolved."
    If Len(traceRunIDOverride) > 0 Then
        traceRunIDForWrite = traceRunIDOverride
    Else
        traceRunIDForWrite = ResolveTraceRunIDForWrite()
    End If
    fileNum = FreeFile
    Open logPath For Append As #fileNum
    mTracePhysicalLogOpenCount = mTracePhysicalLogOpenCount + 1
    fileIsOpen = True
    WriteTraceLogEventToOpenFile fileNum, DEFAULT_TRACE_LOG_FILE_MAX_LINE_LENGTH, traceRunIDForWrite, traceLevel, traceCategory, sourceProcedure, stepName, traceMessage, detailText, errNumber, errSource, errDescription, objectName, objectType, databaseRole, contextName, False
    Close #fileNum
    fileIsOpen = False
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If fileIsOpen Then Close #fileNum
    Debug.Print "ADCPhysicalTraceFileOnlyFailure: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " - EventSource=" & sourceProcedure & "; EventMessage=" & traceMessage & "; Err.Number=" & CStr(savedErrNumber) & "; Err.Source=" & savedErrSource & "; Err.Description=" & savedErrDescription
    On Error GoTo 0
End Sub

Private Sub SetTraceSinkFailure(ByRef failure As TTraceSinkFailure, ByVal failureNumber As Long, ByVal failureSource As String, ByVal failureDescription As String)
    failure.Number = failureNumber
    failure.Source = failureSource
    failure.Description = failureDescription
End Sub

Private Function TryCloseTraceRecordset(ByVal rs As DAO.Recordset, ByRef failure As TTraceSinkFailure) As Boolean
    On Error GoTo ErrHandler
    If Not rs Is Nothing Then rs.Close
    TryCloseTraceRecordset = True
    Exit Function
ErrHandler:
    SetTraceSinkFailure failure, Err.Number, Err.Source, Err.Description
End Function

Private Function TryCloseTraceFile(ByVal fileNum As Integer, ByRef failure As TTraceSinkFailure) As Boolean
    On Error GoTo ErrHandler
    Close #fileNum
    TryCloseTraceFile = True
    Exit Function
ErrHandler:
    SetTraceSinkFailure failure, Err.Number, Err.Source, Err.Description
End Function

Private Function ResolveTraceRunIDForWrite() As String
    If Len(mTraceRunID) > 0 Then
        ResolveTraceRunIDForWrite = mTraceRunID
    ElseIf Len(mLastTraceRunID) > 0 Then
        ResolveTraceRunIDForWrite = mLastTraceRunID
    Else
        Randomize
        mLastTraceRunID = Format$(Now, "yyyymmddhhnnss") & "-adhoc-" & Right$("000000" & CStr(Int(999999 * Rnd) + 1), 6)
        ResolveTraceRunIDForWrite = mLastTraceRunID
    End If
End Function

Private Function GetTraceLogFilePath() As String
    'Intended use: Returns the Access Development Console Trace Log File Path result for direct, query, form, or required cross-module callers.
    On Error GoTo ErrHandler
    Dim dbPath As String
    Dim folderPath As String
    dbPath = Application.CodeProject.FullName
    If InStrRev(dbPath, "\") > 0 Then folderPath = Left$(dbPath, InStrRev(dbPath, "\") - 1)
    If Len(Trim$(folderPath)) = 0 Then folderPath = CurrentProject.Path
    If Len(Trim$(folderPath)) = 0 Then folderPath = Environ$("TEMP")
    If Len(Trim$(folderPath)) = 0 Then Exit Function
    GetTraceLogFilePath = folderPath & "\" & TRACE_LOG_FILE_NAME
    Exit Function
ErrHandler:
    GetTraceLogFilePath = Environ$("TEMP") & "\" & TRACE_LOG_FILE_NAME
End Function

' Purpose: Clears only the physical trace log file. ADC_TraceLog remains the complete retained trace history.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
'   confirmDestructiveAction (Input Optional; Boolean): When True, enables or requests destructive action for this call; when False, the routine follows its normal non-forced/default path. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub ClearTraceLogFile(Optional ByVal showMessage As Boolean = True, Optional ByVal confirmDestructiveAction As Boolean = True)
    'Clears only the physical trace log file. ADC_TraceLog remains the complete retained trace history.
    On Error GoTo ErrHandler
    Dim response As VbMsgBoxResult
    Dim rowCount As Long
    Dim messageText As String
    Dim traceLogFilePath As String
    Dim traceLogFileClearSucceeded As Boolean
    Dim traceLogFileClearDetail As String
    EnsureTraceLogStorageForWrite
    rowCount = CatalogCountSql("SELECT Count(*) AS " & FIELD_CNT & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_TraceLog") & "];")
    traceLogFilePath = GetTraceLogFilePath()
    If confirmDestructiveAction Then
        response = CoreShowMessageResult("Clear only the physical Access Development Console trace log file?" & vbCrLf & vbCrLf & "Retained trace-history table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_TraceLog") & vbCrLf & "Rows retained: " & CStr(rowCount) & vbCrLf & "Physical log file: " & IIf(Len(Trim$(traceLogFilePath)) = 0, "(could not resolve path)", traceLogFilePath) & vbCrLf & vbCrLf & "The trace table is not cleared and remains the complete history.", vbQuestion Or vbYesNo Or vbDefaultButton2, "Clear Access Development Console Log File", "ShowADCMessageResult", "Tables")
        If response <> vbYes Then
            LogTraceMessage "Info", "Trace", "ClearTraceLogFile", "Physical trace log file clear canceled by user.", "Trace-history rows retained: " & CStr(rowCount) & vbCrLf & "Physical log file retained: " & IIf(Len(Trim$(traceLogFilePath)) = 0, "(could not resolve path)", traceLogFilePath), vbNullString, 0, vbNullString, vbNullString, TRACE_LOG_FILE_NAME, "LogFile", "CodeDb", "Trace"
            MarkSystemCommandCancelled
            Exit Sub
        End If
    End If
    traceLogFileClearSucceeded = RecreatePhysicalTraceLogFile(traceLogFileClearDetail)
    If traceLogFileClearSucceeded = False Then Err.Raise vbObjectError + 9010, "ClearTraceLogFile", "The physical Access Development Console trace log file could not be deleted and recreated." & vbCrLf & traceLogFileClearDetail
    messageText = "Access Development Console physical trace log file cleared successfully." & vbCrLf & vbCrLf & "Trace-history table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_TraceLog") & vbCrLf & "Rows retained: " & CStr(rowCount) & vbCrLf & "Physical log file: " & IIf(Len(Trim$(traceLogFilePath)) = 0, "(could not resolve path)", traceLogFilePath) & vbCrLf & traceLogFileClearDetail & CoreOperationSummaryForMessage()
    ReportCoreActionSuccess "ClearTraceLogFile", "Trace", DEFAULT_UTILITY_SHORT_NAME & " physical trace log file cleared successfully.", messageText, "Clear physical trace log file", TRACE_LOG_FILE_NAME, "LogFile", "CodeDb", "Trace", showMessage, "Clear Access Development Console Log File", vbInformation, True, "Tables", False, visibleStatusMessage:="Physical trace log file cleared successfully."
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    messageText = "Unable to clear the physical Access Development Console trace log file." & vbCrLf & vbCrLf & "Procedure: ClearTraceLogFile" & vbCrLf & "Trace-history table retained: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_TraceLog") & vbCrLf & "Rows retained: " & CStr(rowCount) & vbCrLf & "Physical log file: " & IIf(Len(Trim$(traceLogFilePath)) = 0, "(could not resolve path)", traceLogFilePath) & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description:" & vbCrLf & savedErrDescription & TableDiagnosticsLocationForMessage()
    LogTraceMessage "Error", "TraceLogFile", "ClearTraceLogFile", "Unable to clear the physical trace log file; retained trace-history rows were not deleted.", messageText, vbNullString, savedErrNumber, savedErrSource, savedErrDescription, TRACE_LOG_FILE_NAME, "LogFile", "CodeDb", "Trace"
    If showMessage Then
        CoreShowErrorMessage messageText, vbCritical, "Clear Access Development Console Log File", "ClearTraceLogFile", "Tables"
        Exit Sub
    End If
    Err.Raise savedErrNumber, "ClearTraceLogFile", messageText
End Sub

Private Function RecreatePhysicalTraceLogFile(ByRef detailText As String) As Boolean
    On Error GoTo ErrHandler
    Dim fileNum As Integer
    Dim logPath As String
    Dim fileIsOpen As Boolean
    Dim fileExisted As Boolean
    Dim closeFailure As TTraceSinkFailure
    FlushTraceLogFile
    logPath = GetTraceLogFilePath()
    If Len(Trim$(logPath)) = 0 Then
        detailText = "No trace log file path could be resolved. Trace output will still attempt trace-table and Immediate Window reporting."
        Exit Function
    End If
    fileExisted = (Len(Dir$(logPath)) > 0)
    If fileExisted Then Kill logPath
    fileNum = FreeFile
    Open logPath For Output As #fileNum
    mTracePhysicalLogOpenCount = mTracePhysicalLogOpenCount + 1
    fileIsOpen = True
    If TryCloseTraceFile(fileNum, closeFailure) = False Then
        detailText = "The trace log file was opened for recreation, but the file could not be closed. Path: " & logPath & "; Err.Number=" & CStr(closeFailure.Number) & "; Err.Source=" & closeFailure.Source & "; Err.Description=" & closeFailure.Description
        Exit Function
    End If
    fileIsOpen = False
    If fileExisted Then
        detailText = "Trace log file was deleted and recreated for this run. Trace log file: " & logPath
    Else
        detailText = "Trace log file did not exist and was created for this run. Trace log file: " & logPath
    End If
    RecreatePhysicalTraceLogFile = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim cleanupFailure As TTraceSinkFailure
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If fileIsOpen Then
        If TryCloseTraceFile(fileNum, cleanupFailure) = False Then savedErrDescription = savedErrDescription & vbCrLf & "Trace-file cleanup also failed. Err.Number=" & CStr(cleanupFailure.Number) & "; Err.Source=" & cleanupFailure.Source & "; Err.Description=" & cleanupFailure.Description
    End If
    detailText = "Could not delete and recreate the " & DEFAULT_UTILITY_LONG_NAME & " trace log file. Path: " & logPath & "; Err.Number=" & CStr(savedErrNumber) & "; Err.Source=" & savedErrSource & "; Err.Description=" & savedErrDescription
End Function

' Purpose: Performs the public flush trace log file operation using the routine's existing ADC workflow and side effects.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub FlushTraceLogFile()
    'Closes only the retained physical trace-log handle so the most recent complete trace line is visible on disk; the active trace run and trace-table sink remain open.
    On Error Resume Next
    If mTraceLogFileOpen Then Close #mTraceLogFileNumber
    mTraceLogFileNumber = 0
    mTraceLogFilePathOpen = vbNullString
    mTraceLogFileOwnerRunID = vbNullString
    mTraceLogFileOpen = False
    On Error GoTo 0
End Sub

Private Sub CloseRetainedTraceWriteRecordset()
    On Error Resume Next
    If Not mTraceLogWriteRecordset Is Nothing Then mTraceLogWriteRecordset.Close
    Set mTraceLogWriteRecordset = Nothing
    Set mTraceLogWriteDatabase = Nothing
    mTraceLogWriteRecordsetOwnerRunID = vbNullString
    On Error GoTo 0
End Sub

' Purpose: Performs the public release trace log write recordset for schema mutation operation using the routine's existing ADC workflow and side effects.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub ReleaseTraceLogWriteRecordsetForSchemaMutation()
    'Releases the retained TraceLog dynaset without ending the active trace run so cross-module schema/relationship rollback can obtain the required Access table lock; later trace writes reopen the sink on demand.
    CloseRetainedTraceWriteRecordset
End Sub

Private Function AppendTraceLogFileLine(ByVal traceRunID As String, ByVal traceLevel As String, ByVal traceCategory As String, ByVal sourceProcedure As String, ByVal stepName As String, ByVal traceMessage As String, ByVal detailText As String, ByVal errNumber As Long, ByVal errSource As String, ByVal errDescription As String, ByVal objectName As String, ByVal objectType As String, ByVal databaseRole As String, ByVal contextName As String, ByVal printedToImmediate As Boolean, ByRef failure As TTraceSinkFailure) As Boolean
    On Error GoTo ErrHandler
    Dim fileNum As Integer
    Dim logPath As String
    Dim retainedHandle As Boolean
    Dim fileIsOpen As Boolean
    Dim closeFailure As TTraceSinkFailure
    Dim maxLineLength As Long
    logPath = GetTraceLogFilePath()
    If Len(Trim$(logPath)) = 0 Then Err.Raise vbObjectError + 9068, "AppendTraceLogFileLine", "No physical trace-log path could be resolved."
    maxLineLength = TraceLogFileMaxLineLengthForWrite()
    retainedHandle = (Len(mTraceRunID) > 0 And StrComp(traceRunID, mTraceRunID, vbBinaryCompare) = 0)
    If retainedHandle Then
        If mTraceLogFileOpen = False Or StrComp(mTraceLogFilePathOpen, logPath, vbBinaryCompare) <> 0 Or StrComp(mTraceLogFileOwnerRunID, traceRunID, vbBinaryCompare) <> 0 Then
            FlushTraceLogFile
            mTraceLogFileNumber = FreeFile
            Open logPath For Append As #mTraceLogFileNumber
            mTracePhysicalLogOpenCount = mTracePhysicalLogOpenCount + 1
            mTraceLogFileOpen = True
            mTraceLogFilePathOpen = logPath
            mTraceLogFileOwnerRunID = traceRunID
        End If
        WriteTraceLogEventToOpenFile mTraceLogFileNumber, maxLineLength, traceRunID, traceLevel, traceCategory, sourceProcedure, stepName, traceMessage, detailText, errNumber, errSource, errDescription, objectName, objectType, databaseRole, contextName, printedToImmediate
    Else
        fileNum = FreeFile
        Open logPath For Append As #fileNum
        mTracePhysicalLogOpenCount = mTracePhysicalLogOpenCount + 1
        fileIsOpen = True
        WriteTraceLogEventToOpenFile fileNum, maxLineLength, traceRunID, traceLevel, traceCategory, sourceProcedure, stepName, traceMessage, detailText, errNumber, errSource, errDescription, objectName, objectType, databaseRole, contextName, printedToImmediate
        If TryCloseTraceFile(fileNum, closeFailure) = False Then
            SetTraceSinkFailure failure, closeFailure.Number, closeFailure.Source, "The trace event was written, but the physical log file could not be closed. " & closeFailure.Description
            Exit Function
        End If
        fileIsOpen = False
    End If
    AppendTraceLogFileLine = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim cleanupFailure As TTraceSinkFailure
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If retainedHandle Then FlushTraceLogFile
    If fileIsOpen Then
        If TryCloseTraceFile(fileNum, cleanupFailure) = False Then savedErrDescription = savedErrDescription & vbCrLf & "Trace-file cleanup also failed. Err.Number=" & CStr(cleanupFailure.Number) & "; Err.Source=" & cleanupFailure.Source & "; Err.Description=" & cleanupFailure.Description
    End If
    SetTraceSinkFailure failure, savedErrNumber, savedErrSource, savedErrDescription
End Function

Private Function TraceLogFileMaxLineLengthForWrite() As Long
    'Resolves the physical trace-file line-length setting once per active trace run without allowing an invalid trace setting to break trace logging.
    On Error GoTo Fallback
    Dim configuredText As String
    Dim configuredValue As Long
    If Len(mTraceRunID) > 0 And mTraceLogFileMaxLineLengthForRun > 0 Then
        TraceLogFileMaxLineLengthForWrite = mTraceLogFileMaxLineLengthForRun
        Exit Function
    End If
    configuredText = Trim$(TraceSettingText(SETTING_TRACE_LOG_FILE_MAX_LINE_LENGTH, CStr(DEFAULT_TRACE_LOG_FILE_MAX_LINE_LENGTH)))
    If IsNumeric(configuredText) = False Then GoTo Fallback
    configuredValue = CLng(configuredText)
    If configuredValue <= 0 Then GoTo Fallback
    If Len(mTraceRunID) > 0 Then mTraceLogFileMaxLineLengthForRun = configuredValue
    TraceLogFileMaxLineLengthForWrite = configuredValue
    Exit Function
Fallback:
    If Len(mTraceRunID) > 0 Then mTraceLogFileMaxLineLengthForRun = DEFAULT_TRACE_LOG_FILE_MAX_LINE_LENGTH
    TraceLogFileMaxLineLengthForWrite = DEFAULT_TRACE_LOG_FILE_MAX_LINE_LENGTH
End Function

Private Sub WriteTraceLogEventToOpenFile(ByVal fileNum As Integer, ByVal maxLineLength As Long, ByVal traceRunID As String, ByVal traceLevel As String, ByVal traceCategory As String, ByVal sourceProcedure As String, ByVal stepName As String, ByVal traceMessage As String, ByVal detailText As String, ByVal errNumber As Long, ByVal errSource As String, ByVal errDescription As String, ByVal objectName As String, ByVal objectType As String, ByVal databaseRole As String, ByVal contextName As String, ByVal printedToImmediate As Boolean)
    'Writes every trace-event value in a compact labeled block and wraps physical lines without truncating any trace information.
    PrintWrappedTraceLogLine fileNum, maxLineLength, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | Run: " & TraceFileSafeText(traceRunID) & " | Level: " & TraceFileSafeText(traceLevel) & " | Category: " & TraceFileSafeText(traceCategory)
    PrintWrappedTraceLogLine fileNum, maxLineLength, "  Procedure: " & TraceFileSafeText(sourceProcedure) & " | Step: " & TraceFileSafeText(stepName)
    PrintWrappedTraceLogLine fileNum, maxLineLength, "  Message: " & TraceFileSafeText(traceMessage)
    PrintWrappedTraceLogLine fileNum, maxLineLength, "  Detail: " & TraceFileSafeText(detailText)
    PrintWrappedTraceLogLine fileNum, maxLineLength, "  Error: " & CStr(errNumber) & " | Source: " & TraceFileSafeText(errSource) & " | Description: " & TraceFileSafeText(errDescription)
    PrintWrappedTraceLogLine fileNum, maxLineLength, "  Object: " & TraceFileSafeText(objectName) & " | Type: " & TraceFileSafeText(objectType) & " | DatabaseRole: " & TraceFileSafeText(databaseRole) & " | Context: " & TraceFileSafeText(contextName)
    PrintWrappedTraceLogLine fileNum, maxLineLength, "  PrintedToImmediate: " & CStr(printedToImmediate)
    Print #fileNum, vbNullString
    mTracePhysicalLogLinesWrittenCount = mTracePhysicalLogLinesWrittenCount + 1
End Sub

Private Sub PrintWrappedTraceLogLine(ByVal fileNum As Integer, ByVal maxLineLength As Long, ByVal lineText As String)
    'Writes one logical trace line as one or more bounded physical lines while preserving embedded line breaks and every original trace character.
    Dim normalizedText As String
    Dim logicalLines As Variant
    Dim logicalLine As Variant
    Dim remainingText As String
    Dim segmentText As String
    Dim continuationPrefix As String
    normalizedText = Replace(lineText, vbCrLf, vbLf)
    normalizedText = Replace(normalizedText, vbCr, vbLf)
    logicalLines = Split(normalizedText, vbLf)
    For Each logicalLine In logicalLines
        remainingText = CStr(logicalLine)
        If Len(remainingText) = 0 Then
            Print #fileNum, vbNullString
            mTracePhysicalLogLinesWrittenCount = mTracePhysicalLogLinesWrittenCount + 1
        Else
            continuationPrefix = vbNullString
            Do While Len(remainingText) > 0
                If Len(continuationPrefix) >= maxLineLength Then continuationPrefix = vbNullString
                segmentText = Left$(remainingText, maxLineLength - Len(continuationPrefix))
                Print #fileNum, continuationPrefix & segmentText
                mTracePhysicalLogLinesWrittenCount = mTracePhysicalLogLinesWrittenCount + 1
                remainingText = Mid$(remainingText, Len(segmentText) + 1)
                If Len(remainingText) > 0 Then continuationPrefix = IIf(maxLineLength > 4, "    ", vbNullString)
            Loop
        End If
    Next logicalLine
End Sub

Private Sub ReportTraceLogFileWriteFailure(ByRef failure As TTraceSinkFailure, ByVal eventSourceProcedure As String, ByVal eventMessage As String, ByVal nestedTrace As Boolean)
    Dim failureText As String
    Dim shouldReport As Boolean
    mTraceLogFileWriteFailureCount = mTraceLogFileWriteFailureCount + 1
    failureText = "Could not append to the " & DEFAULT_UTILITY_LONG_NAME & " trace log file. Err.Number=" & CStr(failure.Number) & "; Err.Source=" & failure.Source & "; Err.Description=" & failure.Description & "; EventSource=" & eventSourceProcedure & "; EventMessage=" & eventMessage & "; NestedTrace=" & CStr(nestedTrace)
    shouldReport = (mTraceLogFileWriteFailureCount = 1 Or mTraceLogFileWriteFailureCount Mod TRACE_TABLE_WRITE_FAILURE_SUMMARY_INTERVAL = 0 Or StrComp(failureText, mTraceLogFileLastFailureText, vbTextCompare) <> 0)
    mTraceLogFileLastFailureText = failureText
    If shouldReport Then Debug.Print "ADCTraceLogFileFailure: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " - " & failureText & " (failure count: " & CStr(mTraceLogFileWriteFailureCount) & ")"
End Sub


' Purpose: Clears the canonical in-memory trace-note collection for a new logical operation while retaining TraceLog history.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub ResetTraceNotes()
    'Clears the canonical in-memory trace-note collection for a new logical operation while retaining TraceLog history.
    Set mTraceNotes = New Collection
End Sub

Private Sub EnsureTraceNotesReady()
    If mTraceNotes Is Nothing Then Set mTraceNotes = New Collection
End Sub

' Purpose: Resolves trace summary for message from the supplied/current ADC state, using `End` and `JoinLastCollectionItems` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   reportContext (Input Optional; String): Report context text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine. Default: "failure".
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function TraceSummaryForMessage(Optional ByVal reportContext As String = "failure") As String
    'Supports required cross-module Access Development Console table diagnostics.
    On Error GoTo ErrHandler
    Dim headingText As String
    EnsureTraceNotesReady
    If mTraceNotes.count = 0 Then Exit Function
    Select Case LCase$(Trim$(reportContext))
        Case "success"
            headingText = "Recent trace steps:"
        Case "cancellation", "cancelled", "canceled"
            headingText = "Trace steps captured before cancellation:"
        Case Else
            headingText = "Trace steps captured before failure:"
    End Select
    TraceSummaryForMessage = vbCrLf & vbCrLf & headingText & vbCrLf & JoinLastCollectionItems(mTraceNotes, vbCrLf, 40)
    Exit Function
ErrHandler:
    TraceSummaryForMessage = vbCrLf & vbCrLf & "Trace summary failed: " & Err.Description
End Function

' Purpose: Adds trace note for trace text, suppress immediate write.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   traceText (Input; String): Input trace text for add trace note; is tested by If Len(traceText) > 0; is passed to `mTraceNotes`.
'   suppressImmediateWrite (Input Optional; Boolean): True suppresses immediate write; False leaves the normal behavior enabled. Default: False.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub AddTraceNote(ByVal traceText As String, Optional ByVal suppressImmediateWrite As Boolean = False)
    'Retains every operation note in memory. During full Setup, only diagnostic notes are written immediately to TraceLog and the physical file.
    On Error Resume Next
    EnsureTraceNotesReady
    traceText = Trim$(traceText)
    If Len(traceText) > 0 Then
        mTraceNotes.Add Format$(Now, "hh:nn:ss") & " - " & traceText
        If suppressImmediateWrite = False Then
            If mSetupControlActive = False Or TraceNoteRequiresImmediateWrite(traceText) Then LogTraceMessage TRACE_LEVEL_DETAIL, "Tables", "AddTraceNote", traceText
        End If
    End If
End Sub

Private Function TraceNoteRequiresImmediateWrite(ByVal traceText As String) As Boolean
    Dim normalizedText As String
    normalizedText = LCase$(traceText)
    If Len(normalizedText) = 0 Then Exit Function
    TraceNoteRequiresImmediateWrite = _
        InStr(1, normalizedText, "fail", vbBinaryCompare) > 0 Or _
        InStr(1, normalizedText, "error", vbBinaryCompare) > 0 Or _
        InStr(1, normalizedText, "warning", vbBinaryCompare) > 0 Or _
        InStr(1, normalizedText, "conflict", vbBinaryCompare) > 0 Or _
        InStr(1, normalizedText, "rollback", vbBinaryCompare) > 0 Or _
        InStr(1, normalizedText, "cancel", vbBinaryCompare) > 0 Or _
        InStr(1, normalizedText, "could not", vbBinaryCompare) > 0 Or _
        InStr(1, normalizedText, "missing", vbBinaryCompare) > 0 Or _
        InStr(1, normalizedText, "unavailable", vbBinaryCompare) > 0 Or _
        InStr(1, normalizedText, "mismatch", vbBinaryCompare) > 0 Or _
        InStr(1, normalizedText, "stale", vbBinaryCompare) > 0 Or _
        InStr(1, normalizedText, "recovery", vbBinaryCompare) > 0
End Function

Private Sub AddTablesTraceStep(ByVal stepText As String)
    On Error Resume Next
    Dim partStepText As String
    EnsureTraceNotesReady
    partStepText = CoreVisibleProgressText(stepText)
    If Len(Trim$(partStepText)) > 0 Then
        mTraceNotes.Add Format$(Now, "hh:nn:ss") & " - " & partStepText
        LogTraceMessage TRACE_LEVEL_INFO, "Tables", "SetTablesStep", partStepText, vbNullString, partStepText
    End If
End Sub

Private Function TraceShouldPrint(ByVal traceLevel As String, ByVal traceMessage As String, ByVal errNumber As Long, ByVal errDescription As String) As Boolean
    Dim verbosityText As String
    If errNumber <> 0 Then
        TraceShouldPrint = True
        Exit Function
    End If
    If TraceMessageIsPopupLifecycle(traceMessage) Then
        TraceShouldPrint = BooleanTextValue(TraceSettingText(SETTING_LOG_POPUP_LIFECYCLE_TO_IMMEDIATE, CStr(DEFAULT_LOG_POPUP_LIFECYCLE_TO_IMMEDIATE)), DEFAULT_LOG_POPUP_LIFECYCLE_TO_IMMEDIATE)
        Exit Function
    End If
    verbosityText = UCase$(Trim$(TraceSettingText(SETTING_IMMEDIATE_TRACE_VERBOSITY, DEFAULT_IMMEDIATE_TRACE_VERBOSITY)))
    If verbosityText = UCase$(TRACE_VERBOSITY_VERBOSE) Then
        TraceShouldPrint = True
        Exit Function
    End If
    If verbosityText = UCase$(TRACE_VERBOSITY_QUIET) Then
        Select Case UCase$(Trim$(traceLevel))
            Case UCase$(TRACE_LEVEL_ERROR), UCase$(TRACE_LEVEL_WARNING)
                TraceShouldPrint = True
            Case Else
                TraceShouldPrint = TraceMessageIsHighPriority(traceMessage) Or TraceMessageIsHighPriority(errDescription)
        End Select
        Exit Function
    End If
    Select Case UCase$(Trim$(traceLevel))
        Case UCase$(TRACE_LEVEL_ERROR), UCase$(TRACE_LEVEL_WARNING), UCase$(TRACE_LEVEL_INFO)
            TraceShouldPrint = True
        Case UCase$(TRACE_LEVEL_STATUS), UCase$(TRACE_LEVEL_DETAIL)
            TraceShouldPrint = TraceMessageIsHighPriority(traceMessage) Or TraceMessageIsHighPriority(errDescription)
        Case Else
            TraceShouldPrint = TraceMessageIsHighPriority(traceMessage) Or TraceMessageIsHighPriority(errDescription)
    End Select
End Function

Private Function TraceMessageIsPopupLifecycle(ByVal traceMessage As String) As Boolean
    Dim messageKey As String
    messageKey = LCase$(Trim$(traceMessage))
    If Left$(messageKey, Len("popup display attempted")) = "popup display attempted" Then TraceMessageIsPopupLifecycle = True
    If Left$(messageKey, Len("popup shown")) = "popup shown" Then TraceMessageIsPopupLifecycle = True
    If Left$(messageKey, Len("popup result")) = "popup result" Then TraceMessageIsPopupLifecycle = True
    If Left$(messageKey, Len("popup detail text recorded")) = "popup detail text recorded" Then TraceMessageIsPopupLifecycle = True
End Function

' Canonical normal persisted-Settings read authority. These reads do not repair schema or insert rows.
Private Function SettingRawTextNoRepair(ByVal SettingKey As String, Optional ByVal logFallbackTrace As Boolean = True) As String
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim sqlText As String
    If CodeDbTableExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID)) = False Then
        If logFallbackTrace Then CoreAddOperationNote "Direct " & DEFAULT_UTILITY_LONG_NAME & " setting read fell back because the settings table is missing." & " " & SETTING_DIRECT_FALLBACK_KEY_PREFIX & SettingKey
        Exit Function
    End If
    If TableCanOpen(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID), True) = False Then
        If logFallbackTrace Then CoreAddOperationNote "Direct " & DEFAULT_UTILITY_LONG_NAME & " setting read fell back because the settings table exists but cannot be opened." & " " & SETTING_DIRECT_FALLBACK_KEY_PREFIX & SettingKey
        Exit Function
    End If
    If TableFieldExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID), FIELD_SETTING_KEY) = False Then
        If logFallbackTrace Then CoreAddOperationNote "Direct " & DEFAULT_UTILITY_LONG_NAME & " setting read fell back because the settings table is missing the SettingKey field." & " " & SETTING_DIRECT_FALLBACK_KEY_PREFIX & SettingKey
        Exit Function
    End If
    If TableFieldExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID), FIELD_SETTING_VALUE) = False Then
        If logFallbackTrace Then CoreAddOperationNote "Direct " & DEFAULT_UTILITY_LONG_NAME & " setting read fell back because the settings table is missing the SettingValue field." & " " & SETTING_DIRECT_FALLBACK_KEY_PREFIX & SettingKey
        Exit Function
    End If
    sqlText = "SELECT TOP 1 [" & FIELD_SETTING_VALUE & "] FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID) & "] " & "WHERE [" & FIELD_SETTING_KEY & "]=" & SharedDao_SqlQuote(SettingKey)
    Set rs = ADCDb().OpenRecordset(sqlText, dbOpenSnapshot)
    If Not rs.EOF Then
        SettingRawTextNoRepair = SharedText_Nz(rs.fields(FIELD_SETTING_VALUE))
    Else
        If logFallbackTrace Then CoreAddOperationNote "Direct " & DEFAULT_UTILITY_LONG_NAME & " setting read fell back because the settings row is missing." & " " & SETTING_DIRECT_FALLBACK_KEY_PREFIX & SettingKey
    End If
ExitHere:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Exit Function
ErrHandler:
    If logFallbackTrace Then CoreAddOperationNote "Direct " & DEFAULT_UTILITY_LONG_NAME & " setting read fell back after error: " & CStr(Err.Number) & ": " & Err.Description & " " & SETTING_DIRECT_FALLBACK_KEY_PREFIX & SettingKey
    SettingRawTextNoRepair = vbNullString
    Resume ExitHere
End Function

' Purpose: Reads an existing text setting without repairing schema or inserting a missing row; returns the caller fallback when unavailable or blank.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   SettingKey (Input; String): Canonical settings key identifying the configuration value to read, validate, default, or update.
'   DefaultValue (Input; String): Fallback value used when no persisted/supplied value is available; the routine returns/applies it only on the defined default path.
' Returns: String: Returns the effective text value of the requested ADC setting after applying persisted-value and default/fallback rules.
Public Function SettingText(ByVal SettingKey As String, ByVal DefaultValue As String) As String
    'Reads an existing text setting without repairing schema or inserting a missing row; returns the caller fallback when unavailable or blank.
    On Error GoTo ErrHandler
    Dim valueText As String
    valueText = SettingRawTextNoRepair(SettingKey, False)
    If Len(Trim$(valueText)) = 0 Then
        SettingText = DefaultValue
    Else
        SettingText = valueText
    End If
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ADCSettingRead", "Could not read Access Development Console text setting without mutation." & vbCrLf & "SettingKey: " & SettingKey & vbCrLf & "Fallback value: " & DefaultValue & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

' Purpose: Reads the requested setting and returns its effective Boolean value using the persisted/default setting rules.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   SettingKey (Input; String): Canonical settings key identifying the configuration value to read, validate, default, or update.
'   DefaultValue (Input; Boolean): Fallback Boolean used when the setting/source has no usable persisted value; the routine returns/applies this value in that absence case.
' Returns: Boolean: Returns the effective Boolean value of the requested ADC setting after applying the routine's persisted/default-value rules.
Public Function SettingBool(ByVal SettingKey As String, ByVal DefaultValue As Boolean) As Boolean
    'Intended use: Returns the Access Development Console Setting Bool result for direct, query, form, or required cross-module callers.
    On Error GoTo ErrHandler
    Dim valueText As String
    valueText = SettingRawTextNoRepair(SettingKey, False)
    SettingBool = BooleanTextValue(valueText, DefaultValue)
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ADCSettingRead", "Could not read Access Development Console Boolean setting without mutation." & vbCrLf & "SettingKey: " & SettingKey & vbCrLf & "Fallback value: " & CStr(DefaultValue) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

' Purpose: Reads an existing Long setting without repairing or inserting rows; uses the caller fallback when unavailable or blank.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   SettingKey (Input; String): Canonical settings key identifying the configuration value to read, validate, default, or update.
'   DefaultValue (Input; Long): Fallback value used when no persisted/supplied value is available; the routine returns/applies it only on the defined default path.
' Returns: Long: Returns the effective Long value of the requested ADC setting after applying persisted-value and default/fallback rules.
Public Function SettingLong(ByVal SettingKey As String, ByVal DefaultValue As Long) As Long
    'Reads an existing Long setting without repairing or inserting rows; uses the caller fallback when unavailable or blank.
    On Error GoTo ErrHandler
    Dim valueText As String
    valueText = Trim$(SettingRawTextNoRepair(SettingKey, False))
    If Len(valueText) = 0 Then
        SettingLong = DefaultValue
    ElseIf IsNumeric(valueText) Then
        SettingLong = CLng(valueText)
    Else
        Err.Raise vbObjectError + 5405, "ADCSettingRead", "Access Development Console Long setting has an invalid stored value." & vbCrLf & "SettingKey: " & SettingKey & vbCrLf & "Stored value: " & valueText
    End If
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ADCSettingRead", "Could not read Access Development Console Long setting without mutation." & vbCrLf & "SettingKey: " & SettingKey & vbCrLf & "Fallback value: " & CStr(DefaultValue) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

' Purpose: Checks whether setting row exists in the database, collection, or metadata scope supplied to this function.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   SettingKey (Input; String): Canonical settings key identifying the configuration value to read, validate, default, or update.
' Returns: Boolean: True when setting row exists in the inspected context; False when it does not.
Public Function SettingRowExists(ByVal SettingKey As String) As Boolean
    'Intended use: Returns the Access Development Console Setting Row Exists result for direct, query, form, or required cross-module callers.
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim sqlText As String
    If CodeDbTableExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID)) = False Then Exit Function
    If TableCanOpen(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID), True) = False Then Exit Function
    If TableFieldExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID), FIELD_SETTING_KEY) = False Then Exit Function
    sqlText = "SELECT TOP 1 " & SQL_FIELD_SETTING_KEY & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID) & "] " & "WHERE " & SQL_FIELD_SETTING_KEY & "=" & SharedDao_SqlQuote(SettingKey)
    Set rs = ADCDb().OpenRecordset(sqlText, dbOpenSnapshot)
    SettingRowExists = Not rs.EOF
ExitHere:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Exit Function
ErrHandler:
    SettingRowExists = False
    Resume ExitHere
End Function

' Purpose: Sets setting raw text for setting key, setting value.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   SettingKey (Input; String): Canonical settings key identifying the configuration value to read, validate, default, or update.
'   SettingValue (Input; String): Text value to persist or compare for the identified setting; it must already be in the representation accepted by that setting contract.
'   prepareSettings (Input Optional; Boolean): When True, enables or requests settings for this call; when False, the routine follows its normal non-forced/default path. Default: True.
'   updateTimestamps (Input Optional; Boolean): True updates timestamps as part of the operation; False preserves the existing value/state. Default: True.
'   wrapError (Input Optional; Boolean): Boolean state describing whether wrap error applies to this invocation; True selects the named condition/path and False selects its complementary/default path. Default:
'   True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub SetSettingRawText(ByVal SettingKey As String, ByVal SettingValue As String, Optional ByVal prepareSettings As Boolean = True, Optional ByVal updateTimestamps As Boolean = True, Optional ByVal wrapError As Boolean = True)
    'Authoritative physical Settings.SettingValue writer. Callers retain domain policy; flags preserve established readiness/timestamp/error semantics.
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    If prepareSettings Then
        EnsureSettingsReadyForRead
        If SettingRowExists(SettingKey) = False Then
            Err.Raise vbObjectError + 5911, "SetSettingRawText", "Setting row is missing after authoritative Settings readiness/repair." & vbCrLf & "SettingKey: " & SettingKey
        End If
    End If
    Set rs = ADCDb().OpenRecordset(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID), dbOpenDynaset)
    rs.FindFirst "[" & FIELD_SETTING_KEY & "]=" & SharedDao_SqlQuote(SettingKey)
    If rs.NoMatch Then
        Err.Raise vbObjectError + 5911, "SetSettingRawText", "Setting row could not be found after authoritative Settings readiness/repair." & vbCrLf & "SettingKey: " & SettingKey
    End If
    rs.Edit
    SetRecordsetFieldIfExists rs, FIELD_SETTING_VALUE, SettingValue
    If updateTimestamps Then
        SetRecordsetFieldIfExists rs, FIELD_LAST_UPDATED_ON, Now
        SetRecordsetFieldIfExists rs, FIELD_UPDATED_ON, Now
    End If
    rs.Update
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrSource As String
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrSource = Err.Source
    originalErrDescription = Err.Description
    If wrapError Then
        Err.Raise originalErrNumber, "SetSettingRawText", "Could not set Access Development Console setting." & vbCrLf & "SettingKey: " & SettingKey & vbCrLf & "Err.Description: " & originalErrDescription
    End If
    Err.Raise originalErrNumber, originalErrSource, originalErrDescription
End Sub

Private Function TraceSettingText(ByVal SettingKey As String, ByVal DefaultValue As String) As String
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim closeFailure As TTraceSinkFailure
    If CodeDbTableExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID)) = False Then GoTo Fallback
    If TableCanOpen(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID), True) = False Then GoTo Fallback
    If TableFieldExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID), FIELD_SETTING_KEY) = False Then GoTo Fallback
    If TableFieldExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID), FIELD_SETTING_VALUE) = False Then GoTo Fallback
    sqlText = "SELECT TOP 1 [" & FIELD_SETTING_VALUE & "] FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID) & "] WHERE [" & FIELD_SETTING_KEY & "]=" & SharedDao_SqlQuote(SettingKey)
    Set rs = TablesWorkspaceCodeDb().OpenRecordset(sqlText, dbOpenSnapshot)
    If Not rs.EOF Then TraceSettingText = SharedText_Nz(rs.fields(FIELD_SETTING_VALUE))
    If Len(Trim$(TraceSettingText)) = 0 Then TraceSettingText = DefaultValue
CloseAndExit:
    If Not rs Is Nothing Then
        If TryCloseTraceRecordset(rs, closeFailure) = False Then Debug.Print "ADCTraceSettingCloseFailure: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " - Could not close the trace-setting recordset. SettingKey=" & SettingKey & "; Err.Number=" & CStr(closeFailure.Number) & "; Err.Source=" & closeFailure.Source & "; Err.Description=" & closeFailure.Description
    End If
    Exit Function
Fallback:
    TraceSettingText = DefaultValue
    GoTo CloseAndExit
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    TraceSettingText = DefaultValue
    Debug.Print "ADCTraceSettingReadFailure: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " - Falling back to the default trace setting without suppressing the read error. SettingKey=" & SettingKey & "; Err.Number=" & CStr(savedErrNumber) & "; Err.Source=" & savedErrSource & "; Err.Description=" & savedErrDescription
    Resume CloseAndExit
End Function

Private Function TraceMessageIsHighPriority(ByVal valueText As String) As Boolean
    Dim valueKey As String
    valueKey = valueText
    If Len(valueKey) = 0 Then Exit Function
    If InStr(1, valueKey, "failed", vbTextCompare) > 0 Then TraceMessageIsHighPriority = True
    If InStr(1, valueKey, "failure", vbTextCompare) > 0 Then TraceMessageIsHighPriority = True
    If InStr(1, valueKey, "unable", vbTextCompare) > 0 Then TraceMessageIsHighPriority = True
    If InStr(1, valueKey, "could not", vbTextCompare) > 0 Then TraceMessageIsHighPriority = True
    If InStr(1, valueKey, "not found", vbTextCompare) > 0 Then TraceMessageIsHighPriority = True
    If InStr(1, valueKey, "not repaired", vbTextCompare) > 0 Then TraceMessageIsHighPriority = True
    If InStr(1, valueKey, "not created", vbTextCompare) > 0 Then TraceMessageIsHighPriority = True
    If InStr(1, valueKey, "not updated", vbTextCompare) > 0 Then TraceMessageIsHighPriority = True
    If InStr(1, valueKey, "not deleted", vbTextCompare) > 0 Then TraceMessageIsHighPriority = True
    If InStr(1, valueKey, "blocked", vbTextCompare) > 0 Then TraceMessageIsHighPriority = True
    If InStr(1, valueKey, "mismatch", vbTextCompare) > 0 Then TraceMessageIsHighPriority = True
    If InStr(1, valueKey, "conflict", vbTextCompare) > 0 Then TraceMessageIsHighPriority = True
    If InStr(1, valueKey, "rollback", vbTextCompare) > 0 Then TraceMessageIsHighPriority = True
    If InStr(1, valueKey, "exception", vbTextCompare) > 0 Then TraceMessageIsHighPriority = True
    If InStr(1, valueKey, "source group missing", vbTextCompare) > 0 Then TraceMessageIsHighPriority = True
End Function

Private Function TraceImmediateText(ByVal traceLevel As String, ByVal traceCategory As String, ByVal sourceProcedure As String, ByVal traceMessage As String, ByVal objectName As String) As String
    Dim resultText As String
    resultText = Trim$(traceLevel)
    If Len(Trim$(traceCategory)) > 0 Then resultText = resultText & "/" & Trim$(traceCategory)
    If Len(Trim$(sourceProcedure)) > 0 Then resultText = resultText & " - " & Trim$(sourceProcedure)
    If Len(Trim$(objectName)) > 0 Then resultText = resultText & " - " & Trim$(objectName)
    If Len(Trim$(traceMessage)) > 0 Then resultText = resultText & ": " & traceMessage
    TraceImmediateText = resultText
End Function

Private Sub EnsureTraceLogStorageForWrite()
    'Intended use: Provides the single physical TraceLog field/index schema authority for Catalog repair and the low-level trace sink.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    CloseRetainedTraceWriteRecordset
    Set db = TablesWorkspaceCodeDb()
    EnsureTraceLogWriteTable db
    Set db = Nothing
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Set db = Nothing
    Err.Raise savedErrNumber, "EnsureTraceLogStorageForWrite", "Could not create or validate the authoritative TraceLog storage schema." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub EnsureTraceLogScreenshotsStorageForWrite()
    'Intended use: Provides the single physical TraceLog_Screenshots field/index schema authority for Catalog repair and error-report screenshot storage.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Set db = TablesWorkspaceCodeDb()
    EnsureTraceScreenshotWriteTable db
    Set db = Nothing
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Set db = Nothing
    Err.Raise savedErrNumber, "EnsureTraceLogScreenshotsStorageForWrite", "Could not create or validate the authoritative TraceLog_Screenshots storage schema." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub EnsureTraceStorageForWrite()
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    CloseRetainedTraceWriteRecordset
    Set db = TablesWorkspaceCodeDb()
    EnsureTraceLogWriteTable db
    EnsureTraceScreenshotWriteTable db
    Set db = Nothing
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Set db = Nothing
    Err.Raise savedErrNumber, "EnsureTraceStorageForWrite", "Could not create or validate the low-level trace write storage." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub EnsureTraceLogWriteTable(ByVal db As DAO.Database)
    On Error GoTo ErrHandler
    Dim tableName As String
    Dim tdf As DAO.tableDef
    Dim createdTable As Boolean
    Dim fieldDefinition As Variant
    tableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_TraceLog")
    If SharedDao_QueryExists(db, tableName) Then Err.Raise vbObjectError + 9130, "EnsureTraceLogWriteTable", "A saved query uses the required TraceLog table name." & vbCrLf & "Name: " & tableName
    If SharedDao_TableExists(db, tableName) Then
        Set tdf = db.TableDefs(tableName)
    Else
        Set tdf = db.CreateTableDef(tableName)
        createdTable = True
    End If
    For Each fieldDefinition In TraceLogFieldSchemaDefinitions()
        EnsureTraceWriteField tdf, CStr(fieldDefinition(0)), CLng(fieldDefinition(1)), CLng(fieldDefinition(2)), CBool(fieldDefinition(3))
        ApplySchemaFieldContract tdf.fields(CStr(fieldDefinition(0))), fieldDefinition, tableName, "trace-log"
        PumpSetupResponsiveness
    Next fieldDefinition
    If createdTable Then
        db.TableDefs.Append tdf
        db.TableDefs.Refresh
    End If
    Set tdf = db.TableDefs(tableName)
    tdf.fields.Refresh
    tdf.Indexes.Refresh
    EnsureSingleFieldIndex db, tableName, INDEX_PRIMARY_KEY, FIELD_ID, True, True, True, False
    EnsureSingleFieldIndex db, tableName, INDEX_IX_TRACE_RUN_ID, FIELD_TRACE_RUN_ID, False, False, True, False
    EnsureSingleFieldIndex db, tableName, INDEX_IX_TRACE_LEVEL, FIELD_TRACE_LEVEL, False, False, True, False
    EnsureSingleFieldIndex db, tableName, INDEX_IX_TRACE_CATEGORY, FIELD_TRACE_CATEGORY, False, False, True, False
    EnsureSingleFieldIndex db, tableName, INDEX_IX_TRACE_ON, FIELD_TRACE_ON, False, False, True, False
    tdf.Indexes.Refresh
    If SharedDao_TableDefSingleFieldIndexMatches(tdf, INDEX_PRIMARY_KEY, FIELD_ID, True, True) = False Then Err.Raise vbObjectError + 9251, "EnsureTraceLogWriteTable", "The batched TraceLog primary-key verification failed."
    If SharedDao_TableDefSingleFieldIndexMatches(tdf, INDEX_IX_TRACE_RUN_ID, FIELD_TRACE_RUN_ID, False, False) = False Then Err.Raise vbObjectError + 9252, "EnsureTraceLogWriteTable", "The batched TraceRunID index verification failed."
    If SharedDao_TableDefSingleFieldIndexMatches(tdf, INDEX_IX_TRACE_LEVEL, FIELD_TRACE_LEVEL, False, False) = False Then Err.Raise vbObjectError + 9253, "EnsureTraceLogWriteTable", "The batched TraceLevel index verification failed."
    If SharedDao_TableDefSingleFieldIndexMatches(tdf, INDEX_IX_TRACE_CATEGORY, FIELD_TRACE_CATEGORY, False, False) = False Then Err.Raise vbObjectError + 9254, "EnsureTraceLogWriteTable", "The batched TraceCategory index verification failed."
    If SharedDao_TableDefSingleFieldIndexMatches(tdf, INDEX_IX_TRACE_ON, FIELD_TRACE_ON, False, False) = False Then Err.Raise vbObjectError + 9255, "EnsureTraceLogWriteTable", "The batched TraceOn index verification failed."
    If createdTable Then ReconcileStandaloneManagedTableDescriptionIfAllowed tableName, True
    Set tdf = Nothing
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Set tdf = Nothing
    Err.Raise savedErrNumber, "EnsureTraceLogWriteTable", "Could not create or validate the TraceLog write table." & vbCrLf & "Table: " & tableName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub EnsureTraceScreenshotWriteTable(ByVal db As DAO.Database)
    On Error GoTo ErrHandler
    Dim tableName As String
    Dim tdf As DAO.tableDef
    Dim createdTable As Boolean
    tableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_TraceLogScreenshots")
    If SharedDao_QueryExists(db, tableName) Then Err.Raise vbObjectError + 9131, "EnsureTraceScreenshotWriteTable", "A saved query uses the required TraceLog_Screenshots table name." & vbCrLf & "Name: " & tableName
    If SharedDao_TableExists(db, tableName) Then
        Set tdf = db.TableDefs(tableName)
    Else
        Set tdf = db.CreateTableDef(tableName)
        createdTable = True
    End If
    EnsureTraceWriteField tdf, FIELD_ID, dbLong, 0, True
    EnsureTraceWriteField tdf, FIELD_TRACE_LOG_ID, dbLong, 0, False, True
    EnsureTraceWriteField tdf, FIELD_SCREENSHOT, dbAttachment, 0, False
    EnsureTraceWriteField tdf, FIELD_SCREENSHOT_FILE_NAME, dbText, 255, False
    EnsureTraceWriteField tdf, FIELD_CAPTURED_ON, dbDate, 0, False
    EnsureTraceWriteField tdf, FIELD_CAPTURE_WIDTH, dbLong, 0, False
    EnsureTraceWriteField tdf, FIELD_CAPTURE_HEIGHT, dbLong, 0, False
    EnsureTraceWriteField tdf, FIELD_CREATED_ON, dbDate, 0, False
    If createdTable Then
        db.TableDefs.Append tdf
        db.TableDefs.Refresh
    End If
    Set tdf = db.TableDefs(tableName)
    tdf.fields.Refresh
    tdf.Indexes.Refresh
    EnsureSingleFieldIndex db, tableName, INDEX_PRIMARY_KEY, FIELD_ID, True, True, True, False
    EnsureSingleFieldIndex db, tableName, INDEX_UX_TRACE_LOG_SCREENSHOTS_TRACE_LOG_ID, FIELD_TRACE_LOG_ID, True, False, True, False
    tdf.Indexes.Refresh
    If SharedDao_TableDefSingleFieldIndexMatches(tdf, INDEX_PRIMARY_KEY, FIELD_ID, True, True) = False Then Err.Raise vbObjectError + 9256, "EnsureTraceScreenshotWriteTable", "The batched TraceLog_Screenshots primary-key verification failed."
    If SharedDao_TableDefSingleFieldIndexMatches(tdf, INDEX_UX_TRACE_LOG_SCREENSHOTS_TRACE_LOG_ID, FIELD_TRACE_LOG_ID, True, False) = False Then Err.Raise vbObjectError + 9257, "EnsureTraceScreenshotWriteTable", "The batched TraceLogID unique-index verification failed."
    Call EnsureDefinedFieldLookup(tdf, FIELD_TRACE_LOG_ID)
    If createdTable Then ReconcileStandaloneManagedTableDescriptionIfAllowed tableName, True
    Set tdf = Nothing
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Set tdf = Nothing
    Err.Raise savedErrNumber, "EnsureTraceScreenshotWriteTable", "Could not create or validate the TraceLog_Screenshots write table." & vbCrLf & "Table: " & tableName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub EnsureTraceWriteField(ByVal tdf As DAO.tableDef, ByVal fieldName As String, ByVal fieldType As DAO.DataTypeEnum, ByVal fieldSize As Long, ByVal autoNumber As Boolean, Optional ByVal requiredField As Variant)
    On Error GoTo ErrHandler
    Dim fld As DAO.Field
    If SharedDao_TableDefFieldExistsSafe(tdf, fieldName) Then
        Set fld = tdf.fields(fieldName)
        If SharedDao_FieldDefinitionMatches(fld, fieldType, fieldSize, autoNumber) = False Then
            Err.Raise vbObjectError + 9132, "EnsureTraceWriteField", "An existing low-level trace field has an incompatible definition." & vbCrLf & "Table: " & tdf.Name & vbCrLf & "Field: " & fieldName & vbCrLf & "Expected: " & FieldDefinitionDescription(fieldType, fieldSize, autoNumber) & vbCrLf & "Found: " & FieldActualDefinitionDescription(fld)
        End If
        If isMissing(requiredField) = False Then If fld.Required <> CBool(requiredField) Then fld.Required = CBool(requiredField)
        Set fld = Nothing
        Exit Sub
    End If
    If fieldType = dbText And fieldSize > 0 Then
        Set fld = tdf.CreateField(fieldName, fieldType, fieldSize)
    Else
        Set fld = tdf.CreateField(fieldName, fieldType)
    End If
    If autoNumber Then fld.Attributes = fld.Attributes Or dbAutoIncrField
    If isMissing(requiredField) = False Then fld.Required = CBool(requiredField)
    tdf.fields.Append fld
    Set fld = Nothing
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Set fld = Nothing
    Err.Raise savedErrNumber, "EnsureTraceWriteField", "Could not create or validate a low-level trace field." & vbCrLf & "Table: " & tdf.Name & vbCrLf & "Field: " & fieldName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub EnsureTraceLogTableReadyForWrite(ByVal traceRunID As String)
    Dim readinessKey As String
    readinessKey = traceRunID
    If Len(readinessKey) > 0 Then
        If StrComp(mTraceLogTableReadyRunID, readinessKey, vbBinaryCompare) = 0 Then Exit Sub
    End If
    mTraceStorageVerificationCount = mTraceStorageVerificationCount + 1
    EnsureTraceStorageForWrite
    If Len(readinessKey) > 0 Then
        mTraceLogTableReadyRunID = readinessKey
    Else
        mTraceLogTableReadyRunID = vbNullString
    End If
End Sub

Private Function WriteTraceLogRow(ByVal traceRunID As String, ByVal traceLevel As String, ByVal traceCategory As String, ByVal sourceProcedure As String, ByVal stepName As String, ByVal traceMessage As String, ByVal detailText As String, ByVal errNumber As Long, ByVal errSource As String, ByVal errDescription As String, ByVal objectName As String, ByVal objectType As String, ByVal databaseRole As String, ByVal contextName As String, ByVal wasPrintedToImmediate As Boolean, ByRef failure As TTraceSinkFailure, Optional ByRef traceLogID As Long = 0) As Boolean
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim closeFailure As TTraceSinkFailure
    Dim retainedWriter As Boolean
    EnsureTraceLogTableReadyForWrite traceRunID
    retainedWriter = (Len(mTraceRunID) > 0 And StrComp(traceRunID, mTraceRunID, vbBinaryCompare) = 0)
    If retainedWriter Then
        If mTraceLogWriteRecordset Is Nothing Or StrComp(mTraceLogWriteRecordsetOwnerRunID, traceRunID, vbBinaryCompare) <> 0 Then
            CloseRetainedTraceWriteRecordset
            Set mTraceLogWriteDatabase = TablesWorkspaceCodeDb()
            Set mTraceLogWriteRecordset = mTraceLogWriteDatabase.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_TraceLog") & "] WHERE 1=0;", dbOpenDynaset)
            mTraceLogWriteRecordsetOwnerRunID = traceRunID
            mTraceWriteRecordsetOpenCount = mTraceWriteRecordsetOpenCount + 1
        End If
        Set rs = mTraceLogWriteRecordset
    Else
        Set rs = TablesWorkspaceCodeDb().OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_TraceLog") & "] WHERE 1=0;", dbOpenDynaset)
        mTraceWriteRecordsetOpenCount = mTraceWriteRecordsetOpenCount + 1
    End If
    rs.AddNew
    rs.fields(FIELD_TRACE_RUN_ID).value = TraceValueForField(rs, FIELD_TRACE_RUN_ID, traceRunID)
    rs.fields(FIELD_TRACE_ON).value = Now
    rs.fields(FIELD_TRACE_LEVEL).value = TraceValueForField(rs, FIELD_TRACE_LEVEL, traceLevel)
    rs.fields(FIELD_TRACE_CATEGORY).value = TraceValueForField(rs, FIELD_TRACE_CATEGORY, traceCategory)
    rs.fields(FIELD_SOURCE_PROCEDURE).value = TraceValueForField(rs, FIELD_SOURCE_PROCEDURE, sourceProcedure)
    rs.fields(FIELD_STEP_NAME).value = TraceValueForField(rs, FIELD_STEP_NAME, stepName)
    rs.fields(FIELD_TRACE_MESSAGE).value = TraceValueForField(rs, FIELD_TRACE_MESSAGE, traceMessage)
    rs.fields(FIELD_DETAIL_TEXT).value = TraceValueForField(rs, FIELD_DETAIL_TEXT, detailText)
    rs.fields(FIELD_ERR_NUMBER).value = errNumber
    rs.fields(FIELD_ERR_SOURCE).value = TraceValueForField(rs, FIELD_ERR_SOURCE, errSource)
    rs.fields(FIELD_ERR_DESCRIPTION).value = TraceValueForField(rs, FIELD_ERR_DESCRIPTION, errDescription)
    rs.fields(FIELD_OBJECT_NAME).value = TraceValueForField(rs, FIELD_OBJECT_NAME, objectName)
    rs.fields(FIELD_OBJECT_TYPE).value = TraceValueForField(rs, FIELD_OBJECT_TYPE, objectType)
    rs.fields(FIELD_DATABASE_ROLE).value = TraceValueForField(rs, FIELD_DATABASE_ROLE, databaseRole)
    rs.fields(FIELD_CONTEXT_NAME).value = TraceValueForField(rs, FIELD_CONTEXT_NAME, contextName)
    rs.fields(FIELD_WAS_PRINTED_TO_IMMEDIATE).value = wasPrintedToImmediate
    rs.fields(FIELD_CREATED_ON).value = Now
    rs.Update
    rs.Bookmark = rs.LastModified
    traceLogID = CLng(rs.fields(FIELD_ID).value)
    mTraceRowsWrittenCount = mTraceRowsWrittenCount + 1
    If retainedWriter = False Then
        If TryCloseTraceRecordset(rs, closeFailure) = False Then
            SetTraceSinkFailure failure, closeFailure.Number, closeFailure.Source, "The trace row was written, but its DAO recordset could not be closed. " & closeFailure.Description
            Exit Function
        End If
    End If
    Set rs = Nothing
    WriteTraceLogRow = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim cleanupFailure As TTraceSinkFailure
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If retainedWriter Then
        CloseRetainedTraceWriteRecordset
        mTraceLogTableReadyRunID = vbNullString
    ElseIf Not rs Is Nothing Then
        If TryCloseTraceRecordset(rs, cleanupFailure) = False Then savedErrDescription = savedErrDescription & vbCrLf & "Trace-table recordset cleanup also failed. Err.Number=" & CStr(cleanupFailure.Number) & "; Err.Source=" & cleanupFailure.Source & "; Err.Description=" & cleanupFailure.Description
    End If
    SetTraceSinkFailure failure, savedErrNumber, savedErrSource, savedErrDescription
End Function

Private Function TraceValueForField(ByVal rs As DAO.Recordset, ByVal fieldName As String, ByVal valueText As Variant) As Variant
    On Error GoTo ErrHandler
    Dim textValue As String
    If IsNull(valueText) Then
        TraceValueForField = Null
        Exit Function
    End If
    textValue = CStr(valueText)
    If Len(Trim$(textValue)) = 0 Then
        TraceValueForField = Null
        Exit Function
    End If
    If rs.fields(fieldName).Type = dbText Then
        If rs.fields(fieldName).Size > 0 Then
            If Len(textValue) > rs.fields(fieldName).Size Then textValue = Left$(textValue, rs.fields(fieldName).Size)
        End If
    End If
    TraceValueForField = textValue
    Exit Function
ErrHandler:
    TraceValueForField = Shared_String_ConvertToNullIfBlank(valueText)
End Function

Private Function InsertReportedErrorScreenshot(ByVal traceLogID As Long, ByVal screenshotPath As String, ByVal captureWidth As Long, ByVal captureHeight As Long) As Long
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset2
    Dim attachmentRs As DAO.Recordset2
    Dim screenshotID As Long
    EnsureTraceLogScreenshotsStorageForWrite
    Set db = TablesWorkspaceCodeDb()
    Set rs = db.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_TraceLogScreenshots") & "] WHERE 1=0;", dbOpenDynaset)
    rs.AddNew
    rs.fields(FIELD_TRACE_LOG_ID).value = traceLogID
    rs.fields(FIELD_SCREENSHOT_FILE_NAME).value = FileNameFromPath(screenshotPath)
    rs.fields(FIELD_CAPTURED_ON).value = Now
    rs.fields(FIELD_CAPTURE_WIDTH).value = captureWidth
    rs.fields(FIELD_CAPTURE_HEIGHT).value = captureHeight
    rs.fields(FIELD_CREATED_ON).value = Now
    rs.Update
    rs.Bookmark = rs.LastModified
    screenshotID = CLng(rs.fields(FIELD_ID).value)
    rs.Edit
    Set attachmentRs = rs.fields(FIELD_SCREENSHOT).value
    attachmentRs.AddNew
    attachmentRs.fields("FileData").LoadFromFile screenshotPath
    attachmentRs.Update
    attachmentRs.Close
    Set attachmentRs = Nothing
    rs.Update
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    InsertReportedErrorScreenshot = screenshotID
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not attachmentRs Is Nothing Then attachmentRs.Close
    If Not rs Is Nothing Then rs.Close
    If screenshotID > 0 Then db.Execute "DELETE FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_TraceLogScreenshots") & "] WHERE " & SQL_FIELD_ID & "=" & CStr(screenshotID) & ";", dbFailOnError
    Set attachmentRs = Nothing
    Set rs = Nothing
    Set db = Nothing
    On Error GoTo 0
    Err.Raise savedErrNumber, "InsertReportedErrorScreenshot", "Could not attach the captured screenshot to the linked TraceLog row." & vbCrLf & "TraceLog ID: " & CStr(traceLogID) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

' Purpose: Logs reported error for description text, screenshot path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   descriptionText (Input; String): Description text text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   screenshotPath (Input; String): Filesystem or database path identifying the screenshot path location used by the operation.
'   captureWidth (Input; Long): Access layout dimension, in the units expected by the surrounding form/report builder, for the capture width value.
'   captureHeight (Input; Long): Access layout dimension, in the units expected by the surrounding form/report builder, for the capture height value.
' Returns: Long: Returns the persistent identifier/result produced when the reported error is logged, using zero or the routine's defined sentinel when no row is created.
Public Function LogReportedError(ByVal descriptionText As String, ByVal screenshotPath As String, ByVal captureWidth As Long, ByVal captureHeight As Long) As Long
    On Error GoTo ErrHandler
    Dim traceRunID As String
    Dim traceLogID As Long
    Dim screenshotID As Long
    Dim tableFailure As TTraceSinkFailure
    Dim fileFailure As TTraceSinkFailure
    Dim physicalDetail As String
    If Len(Trim$(descriptionText)) = 0 Then Err.Raise vbObjectError + 9110, "LogReportedError", "An error description is required."
    If Len(Trim$(screenshotPath)) = 0 Or Len(Dir$(screenshotPath)) = 0 Then Err.Raise vbObjectError + 9111, "LogReportedError", "The captured screenshot file does not exist." & vbCrLf & "Path: " & screenshotPath
    traceRunID = ResolveTraceRunIDForWrite()
    Debug.Print "ADCTrace: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " - Error/UserReportedError - CompleteReportedError: User-reported error submitted."
    If WriteTraceLogRow(traceRunID, "Error", "UserReportedError", "CompleteReportedError", "Capture and store user-reported error", "User-reported error submitted.", descriptionText, 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_TraceLogScreenshots"), "Table", "CodeDb", "UserReportedError", True, tableFailure, traceLogID) = False Then
        Err.Raise tableFailure.Number, "LogReportedError", "The user-reported error could not be written to TraceLog." & vbCrLf & "Err.Source: " & tableFailure.Source & vbCrLf & "Err.Description: " & tableFailure.Description
    End If
    screenshotID = InsertReportedErrorScreenshot(traceLogID, screenshotPath, captureWidth, captureHeight)
    physicalDetail = descriptionText & vbCrLf & "TraceLog ID: " & CStr(traceLogID) & vbCrLf & "Screenshot table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_TraceLogScreenshots") & vbCrLf & "Screenshot row ID: " & CStr(screenshotID) & vbCrLf & "Screenshot file name: " & FileNameFromPath(screenshotPath) & vbCrLf & "Capture dimensions: " & CStr(captureWidth) & " x " & CStr(captureHeight)
    If AppendTraceLogFileLine(traceRunID, "Error", "UserReportedError", "CompleteReportedError", "Capture and store user-reported error", "User-reported error submitted.", physicalDetail, 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_TraceLogScreenshots"), "Table", "CodeDb", "UserReportedError", True, fileFailure) = False Then
        Err.Raise fileFailure.Number, "LogReportedError", "The report and screenshot were stored in the database, but the physical trace log entry could not be written." & vbCrLf & "Err.Source: " & fileFailure.Source & vbCrLf & "Err.Description: " & fileFailure.Description
    End If
    LogReportedError = traceLogID
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim failureDetail As String
    Dim fallbackFailure As TTraceSinkFailure
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    failureDetail = descriptionText & vbCrLf & "Screenshot path: " & screenshotPath & vbCrLf & "Capture dimensions: " & CStr(captureWidth) & " x " & CStr(captureHeight) & vbCrLf & "Storage failure: " & savedErrDescription
    On Error Resume Next
    If traceLogID > 0 Then MarkReportedErrorStorageFailure traceLogID, savedErrNumber, savedErrSource, savedErrDescription
    If Len(traceRunID) > 0 Then AppendTraceLogFileLine traceRunID, "Error", "UserReportedError", "CompleteReportedError", "Capture and store user-reported error", "User-reported error storage failed.", failureDetail, savedErrNumber, savedErrSource, savedErrDescription, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_TraceLogScreenshots"), "Table", "CodeDb", "UserReportedError", True, fallbackFailure
    On Error GoTo 0
    If savedErrNumber = 0 Then savedErrNumber = vbObjectError + 9112
    If Len(Trim$(savedErrSource)) = 0 Then savedErrSource = "LogReportedError"
    Err.Raise savedErrNumber, "LogReportedError", "The user-reported error could not be stored completely." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Sub MarkReportedErrorStorageFailure(ByVal traceLogID As Long, ByVal failureNumber As Long, ByVal failureSource As String, ByVal failureDescription As String)
    On Error GoTo ExitHere
    Dim rs As DAO.Recordset
    EnsureTraceLogStorageForWrite
    Set rs = TablesWorkspaceCodeDb().OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_TraceLog") & "] WHERE " & SQL_FIELD_ID & "=" & CStr(traceLogID) & ";", dbOpenDynaset)
    If rs.EOF = False Then
        rs.Edit
        rs.fields(FIELD_TRACE_MESSAGE).value = "User-reported error screenshot or physical-log storage failed."
        rs.fields(FIELD_ERR_NUMBER).value = failureNumber
        rs.fields(FIELD_ERR_SOURCE).value = TraceValueForField(rs, FIELD_ERR_SOURCE, failureSource)
        rs.fields(FIELD_ERR_DESCRIPTION).value = TraceValueForField(rs, FIELD_ERR_DESCRIPTION, failureDescription)
        rs.Update
    End If
    rs.Close
ExitHere:
    Set rs = Nothing
End Sub


Private Sub ReportTraceTableWriteFailure(ByRef failure As TTraceSinkFailure, ByVal eventSourceProcedure As String, ByVal eventMessage As String)
    Dim failureText As String
    Dim shouldReport As Boolean
    mTraceTableWriteFailureCount = mTraceTableWriteFailureCount + 1
    failureText = "Trace table write failed. Err.Number=" & CStr(failure.Number) & "; Err.Source=" & failure.Source & "; Err.Description=" & failure.Description & "; EventSource=" & eventSourceProcedure & "; EventMessage=" & eventMessage
    shouldReport = (mTraceTableWriteFailureCount = 1 Or mTraceTableWriteFailureCount Mod TRACE_TABLE_WRITE_FAILURE_SUMMARY_INTERVAL = 0 Or StrComp(failureText, mTraceTableLastFailureText, vbTextCompare) <> 0)
    mTraceTableLastFailureText = failureText
    If shouldReport Then Debug.Print "ADCTraceTableFailure: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " - " & failureText & " (failure count: " & CStr(mTraceTableWriteFailureCount) & ")"
End Sub

' Purpose: Resolves trace backend diagnostics for message from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function TraceBackendDiagnosticsForMessage() As String
    'Intended use: Returns the Access Development Console Trace Backend Diagnostics For Message result for direct, query, form, or required cross-module callers.
    On Error GoTo ErrHandler
    Dim warningText As String
    If mTraceTableWriteFailureCount > 0 Then
        warningText = warningText & "Trace table write failures this run: " & CStr(mTraceTableWriteFailureCount)
        If Len(Trim$(mTraceTableLastFailureText)) > 0 Then warningText = warningText & ". Last failure: " & mTraceTableLastFailureText
    End If
    If Len(Trim$(mTraceLogFileClearFailureText)) > 0 Then
        If Len(warningText) > 0 Then warningText = warningText & vbCrLf
        warningText = warningText & "Trace log file clear warning: " & mTraceLogFileClearFailureText
    End If
    If mTraceLogFileWriteFailureCount > 0 Then
        If Len(warningText) > 0 Then warningText = warningText & vbCrLf
        warningText = warningText & "Trace log file append failures this run: " & CStr(mTraceLogFileWriteFailureCount)
        If Len(Trim$(mTraceLogFileLastFailureText)) > 0 Then warningText = warningText & ". Last failure: " & mTraceLogFileLastFailureText
    End If
    If Len(Trim$(warningText)) > 0 Then TraceBackendDiagnosticsForMessage = vbCrLf & vbCrLf & "Trace storage warnings:" & vbCrLf & warningText
    Exit Function
ErrHandler:
    TraceBackendDiagnosticsForMessage = vbCrLf & vbCrLf & "Trace storage warnings: (could not read trace storage warning state: " & Err.Description & ")"
End Function

Private Function TraceLogFieldSchemaDefinitions() As Collection
    'Single physical field-schema authority for TraceLog creation, current field-definition repair/verification, and canonical order.
    ' CANONICAL AUTHORITY: rows owns the in-memory definition set constructed by TraceLogFieldSchemaDefinitions; downstream consumers derive views from this set rather than create a competing authority.
    Dim rows As New Collection
    rows.Add Array(FIELD_ID, dbLong, 0, True)
    rows.Add Array(FIELD_TRACE_RUN_ID, dbText, 100, False)
    rows.Add Array(FIELD_TRACE_ON, dbDate, 0, False)
    rows.Add Array(FIELD_TRACE_LEVEL, dbText, 40, False, Null, Null, Null, Null, Null)
    rows.Add Array(FIELD_TRACE_CATEGORY, dbText, 80, False)
    rows.Add Array(FIELD_SOURCE_PROCEDURE, dbText, 150, False)
    rows.Add Array(FIELD_STEP_NAME, dbText, 150, False)
    rows.Add Array(FIELD_TRACE_MESSAGE, dbMemo, 0, False)
    rows.Add Array(FIELD_DETAIL_TEXT, dbMemo, 0, False)
    rows.Add Array(FIELD_OBJECT_NAME, dbText, 255, False)
    rows.Add Array(FIELD_OBJECT_TYPE, dbText, 80, False)
    rows.Add Array(FIELD_DATABASE_ROLE, dbText, 40, False)
    rows.Add Array(FIELD_CONTEXT_NAME, dbText, 150, False)
    rows.Add Array(FIELD_ERR_NUMBER, dbLong, 0, False)
    rows.Add Array(FIELD_ERR_SOURCE, dbText, 255, False)
    rows.Add Array(FIELD_ERR_DESCRIPTION, dbMemo, 0, False)
    rows.Add Array(FIELD_WAS_PRINTED_TO_IMMEDIATE, dbBoolean, 0, False)
    rows.Add Array(FIELD_CREATED_ON, dbDate, 0, False)
    Set TraceLogFieldSchemaDefinitions = rows
End Function

Private Sub EnsureKeyIntegrityBeforeSchemaRepair()
    Dim keyIntegritySummary As String
    keyIntegritySummary = DuplicateKeySummaryForTable(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID), FIELD_SETTING_KEY)
    keyIntegritySummary = SharedText_AppendNonBlankLine(keyIntegritySummary, ConfigMacroKeyIntegritySummary(TablesWorkspaceCodeDb(), ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config"), True))
    keyIntegritySummary = SharedText_AppendNonBlankLine(keyIntegritySummary, DuplicateKeySummaryForTable(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LookupDefinitions"), "LookupLogicalID"))
    If Len(Trim$(keyIntegritySummary)) > 0 Then
        Err.Raise vbObjectError + 5260, "EnsureKeyIntegrityBeforeSchemaRepair", "Key-integrity preflight failed. Schema repair cannot safely enforce required/unique key contracts until existing invalid keys are corrected." & vbCrLf & vbCrLf & keyIntegritySummary & vbCrLf & vbCrLf & "Correct the invalid key values, then rerun the table repair."
    End If
End Sub

Private Function ConfigMacroKeyIntegritySummary(ByVal db As DAO.Database, ByVal tableName As String, Optional ByVal legacyConfigurationRowsOnly As Boolean = False) As String
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim rowKindFilter As String
    Dim summaryText As String
    Dim invalidCount As Long
    If db Is Nothing Then Err.Raise vbObjectError + 5261, "ConfigMacroKeyIntegritySummary", "Database reference is Nothing."
    If SharedDao_TableExists(db, tableName) = False Then Exit Function
    If TableCanOpenInDatabase(db, tableName) = False Then Exit Function
    If SharedDao_DatabaseTableFieldExistsSafe(db, tableName, FIELD_MACRO_KEY) = False Then Exit Function
    If legacyConfigurationRowsOnly And SharedDao_DatabaseTableFieldExistsSafe(db, tableName, FIELD_ROW_KIND) Then rowKindFilter = " AND (" & SQL_FIELD_ROW_KIND & " Is Null OR Trim(" & SQL_FIELD_ROW_KIND & ")='' OR " & SQL_FIELD_ROW_KIND & "=" & SharedDao_SqlQuote(ROW_KIND_CONFIGURATION) & ")"
    sqlText = "SELECT Count(*) AS " & FIELD_INVALID_COUNT & " FROM [" & tableName & "] WHERE (" & SQL_FIELD_MACRO_KEY & " Is Null OR Trim(" & SQL_FIELD_MACRO_KEY & ")='')" & rowKindFilter
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    If Not rs.EOF Then invalidCount = CLng(Nz(rs.fields(FIELD_INVALID_COUNT).value, 0))
    rs.Close
    Set rs = Nothing
    If invalidCount > 0 Then summaryText = tableName & "." & SQL_FIELD_MACRO_KEY & " blank/whitespace key rows: " & CStr(invalidCount)
    sqlText = "SELECT " & SQL_FIELD_MACRO_KEY & " AS " & FIELD_KEY_VALUE & ", Count(*) AS " & FIELD_DUP_COUNT & " FROM [" & tableName & "] WHERE " & SQL_FIELD_MACRO_KEY & " Is Not Null AND Trim(" & SQL_FIELD_MACRO_KEY & ")<>''" & rowKindFilter & " GROUP BY " & SQL_FIELD_MACRO_KEY & " HAVING Count(*)>1 ORDER BY " & SQL_FIELD_MACRO_KEY
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    Do While Not rs.EOF
        PumpSetupResponsiveness
        summaryText = SharedText_AppendNonBlankLine(summaryText, tableName & "." & SQL_FIELD_MACRO_KEY & " duplicate config key: " & CStr(rs.fields(FIELD_KEY_VALUE).value) & " (rows: " & CStr(CLng(rs.fields(FIELD_DUP_COUNT).value)) & ")")
        rs.MoveNext
    Loop
    ConfigMacroKeyIntegritySummary = summaryText
ExitHere:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Exit Function
ErrHandler:
    ConfigMacroKeyIntegritySummary = tableName & "." & SQL_FIELD_MACRO_KEY & ": key-integrity preflight could not inspect this field. " & Err.Description
    Resume ExitHere
End Function

Private Function DuplicateKeySummaryForTable(ByVal tableName As String, ByVal fieldName As String) As String
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim lines As Collection
    If CodeDbTableExists(tableName) = False Then Exit Function
    If TableCanOpen(tableName, True) = False Then
        CoreAddOperationNote "Skipped duplicate-key preflight for unusable table definition [" & tableName & "]. " & "The table will be recreated if repair is required. " & SharedMessage_OneLine(TableOpenFailureDetail(tableName), 300)
        Exit Function
    End If
    If TableFieldExists(tableName, fieldName) = False Then Exit Function
    ' INTENTIONAL ORDERED VIEW: lines preserves the deterministic sequence required by DuplicateKeySummaryForTable; keyed-only or unordered storage would not preserve the consumer semantics.
    Set lines = New Collection
    sqlText = "SELECT [" & fieldName & "] AS " & FIELD_KEY_VALUE & ", Count(*) AS " & FIELD_DUP_COUNT & " " & "FROM [" & tableName & "] " & "WHERE [" & fieldName & "] Is Not Null AND Trim([" & fieldName & "])<>'' " & "GROUP BY [" & fieldName & "] " & "HAVING Count(*)>1 " & "ORDER BY [" & fieldName & "]"
    Set rs = TablesWorkspaceCodeDb().OpenRecordset(sqlText, dbOpenSnapshot)
    Do While Not rs.EOF
        PumpSetupResponsiveness
        lines.Add tableName & ".[" & fieldName & "] duplicate: " & CStr(rs!keyValue) & " (rows: " & CStr(CLng(rs.fields(FIELD_DUP_COUNT))) & ")"
        rs.MoveNext
    Loop
    If lines.count > 0 Then DuplicateKeySummaryForTable = SharedCollection_Join(lines, vbCrLf)
ExitHere:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Exit Function
ErrHandler:
    DuplicateKeySummaryForTable = tableName & ".[" & fieldName & "]: duplicate-key preflight could not inspect this field. " & Err.Description
    Resume ExitHere
End Function

' ===============================================================================================================================================================================
' CORE DATA TABLE / LOOKUP / LAUNCHER FOUNDATION PROCEDURES
' ===============================================================================================================================================================================

Private Sub RepairDataTablesCoreForSetup(ByVal settingsSchemaDefinitions As Collection, ByVal settingsIndexDefinitions As Collection, ByVal configSchemaDefinitions As Collection, ByVal configIndexDefinitions As Collection, ByVal macroCategorySchemaDefinitions As Collection, ByVal macroSubcategorySchemaDefinitions As Collection, ByVal coreSettingDefinitions As Collection, ByVal editableLookupValueTextSizes As Object, Optional ByVal schemaRepairNeededKnown As Boolean = False, Optional ByVal schemaRepairNeededValue As Boolean = False, Optional ByVal schemaRepairSummaryValue As String = vbNullString, Optional ByVal designChangeAlreadyPrepared As Boolean = False)
    RequireCatalogMutationRecoveryReady "RepairDataTablesCoreForSetup"
    'A2-only physical table structure repair. Saved queries, FormDesign, Navigation Pane placement, and historical schema/version conversion are excluded.
    On Error GoTo ErrHandler
    Dim schemaRepairNeeded As Boolean
    Dim schemaRepairSummary As String
    Dim currentStep As String
    Dim configTableName As String
    Dim candidateTableName As String
    Dim settingsTableName As String
    Dim databaseName As String
    Dim settingsTablePreparedThisPass As Boolean
    Dim previousDesignBatchPrepared As Boolean
    previousDesignBatchPrepared = mTableDesignBatchPrepared
    If designChangeAlreadyPrepared Then mTableDesignBatchPrepared = True
    configTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")
    candidateTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates")
    settingsTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID)
    databaseName = TableSafeDatabaseName()
    SetTablesStep currentStep, "Synchronize A2 table definitions when the caller has not already prepared the design batch."
    If designChangeAlreadyPrepared Then
        AddTraceNote "A2 table-design batch reused the caller-owned TableDefs synchronization; duplicate batch-entry refresh skipped."
    Else
        RefreshTableDefsSafely
    End If
    SetTablesStep currentStep, "Create or repair the USysRibbons table schema."
    EnsureUSysRibbonsTableSchema
    SetTablesStep currentStep, "Prepare settings table physical name: " & settingsTableName
    PrepareTableNameForCreateOrRepair settingsTableName, "settings"
    If CodeDbTableExists(settingsTableName) = False Then
        SetTablesStep currentStep, "Create settings table before other A2 structures: " & settingsTableName
        RepairCodeDbSettingsTableForSetup False, settingsSchemaDefinitions, settingsIndexDefinitions
        settingsTablePreparedThisPass = True
        MarkSettingsReadyForCurrentContext
        RefreshTableDefsSafely
        RequireSafeWorkContinue "RepairDataTablesCoreForSetup", "Settings table [" & settingsTableName & "]"
    Else
        RequirePreparedTableCanOpen settingsTableName, "settings"
        If SettingsTableSchemaIsValid(settingsSchemaDefinitions, settingsIndexDefinitions) Then MarkSettingsReadyForCurrentContext
    End If
    SetTablesStep currentStep, "Detect A2 schema repair need."
    If schemaRepairNeededKnown Then
        schemaRepairNeeded = schemaRepairNeededValue
        schemaRepairSummary = schemaRepairSummaryValue
        AddTraceNote "Reused the A2 schema-repair preflight already calculated by the outer foundation repair; duplicate schema inspection was skipped."
    Else
        schemaRepairNeeded = TableSchemaRepairRequired(schemaRepairSummary, settingsSchemaDefinitions, settingsIndexDefinitions, configSchemaDefinitions, configIndexDefinitions, macroCategorySchemaDefinitions, macroSubcategorySchemaDefinitions, editableLookupValueTextSizes)
    End If
    If schemaRepairNeeded Then ClearSettingsRuntimeStateForCurrentContext
    If schemaRepairNeeded Then
        SetTablesStep currentStep, "Run key-integrity preflight before A2 schema repair."
        EnsureKeyIntegrityBeforeSchemaRepair
    End If
    If settingsTablePreparedThisPass = False Then
        SetTablesStep currentStep, "Repair and verify settings table structure: " & settingsTableName
        RepairCodeDbSettingsTableForSetup False, settingsSchemaDefinitions, settingsIndexDefinitions
        RequireSafeWorkContinue "RepairDataTablesCoreForSetup", "Settings table [" & settingsTableName & "]"
    End If
    MarkSettingsReadyForCurrentContext
    SetTablesStep currentStep, "Repair macro classification table structures; B5 owns category/subcategory relationship repair."
    RepairMacroClassificationTables False, mTableDesignBatchPrepared, False, macroCategorySchemaDefinitions, macroSubcategorySchemaDefinitions
    RequireSafeWorkContinue "RepairDataTablesCoreForSetup", "Macro classification table family"
    SetTablesStep currentStep, "Retire obsolete RowKind lookup storage when no B5-owned relationship still references it."
    RetireObsoleteRowKindLookupTableForSetup
    SetTablesStep currentStep, "Create or repair persistent editable lookup tables and seed only newly created tables."
    EnsureEditableLookupTablesForSetup editableLookupValueTextSizes
    RequireSafeWorkContinue "RepairDataTablesCoreForSetup", "Editable lookup table family"
    SetTablesStep currentStep, "Create or repair MC_Candidates before migrating legacy mixed MC_Records rows: " & candidateTableName
    EnsureCandidateTable False
    RequireSafeWorkContinue "RepairDataTablesCoreForSetup", "MC_Candidates table [" & candidateTableName & "]"
    SetTablesStep currentStep, "Migrate legacy candidate/discovery evidence out of MC_Records."
    MigrateLegacyConfigCandidateRowsForSetup configTableName, candidateTableName
    SetTablesStep currentStep, "Repair configuration-only MC_Records table structure: " & configTableName
    EnsureConfigTable False, configSchemaDefinitions, configIndexDefinitions
    RetireLegacyCandidateFieldsFromConfigTableForSetup configTableName
    RequireSafeWorkContinue "RepairDataTablesCoreForSetup", "MC_Records table [" & configTableName & "]"
    SetTablesStep currentStep, "Materialize the complete canonical Settings catalog required before later Setup consumers."
    RepairSettingDefaultsForSetup False, coreSettingDefinitions
    RequireSafeWorkContinue "RepairDataTablesCoreForSetup", "Settings default batch"
    SetTablesStep currentStep, "Repair health-review storage."
    RepairHealthReviewTableForSetup False
    SetTablesStep currentStep, "Repair launcher-review structural storage foundation; defer lookup metadata to B2 and its saved-query foundation to A3."
    RepairLauncherTableFoundationForSetup False, True, True, False
    RequireSafeWorkContinue "RepairDataTablesCoreForSetup", "Health and launcher review storage"
    SetTablesStep currentStep, "Verify the complete A2 physical table schema, including HealthReview and launcher-review storage."
    VerifyTableReady settingsTableName, "settings"
    VerifyTableReady configTableName, configTableName
    VerifyTableReady candidateTableName, candidateTableName
    'Final proof owns one strict TableDefs synchronization. Do not use the error-suppressing general refresh here: the performance repair removes per-field/per-index refreshes, so this boundary must remain
    'fail-closed.
    TablesWorkspaceCodeDb().TableDefs.Refresh
    If TableSchemaRepairRequired(schemaRepairSummary, settingsSchemaDefinitions, settingsIndexDefinitions, configSchemaDefinitions, configIndexDefinitions, macroCategorySchemaDefinitions, macroSubcategorySchemaDefinitions, editableLookupValueTextSizes) Then Err.Raise vbObjectError + 5293, "RepairDataTablesCoreForSetup", "A2 physical table schema remains invalid after repair." & vbCrLf & "Schema issues remaining:" & vbCrLf & schemaRepairSummary
    CoreAddOperationNote "A2 physical data-table foundation is current. No A3-owned saved-query foundation, Navigation Pane, or FormDesign work was performed. A2-owned obsolete lookup-query retirement may have been performed as part of physical lookup-table conversion."
    mTableDesignBatchPrepared = previousDesignBatchPrepared
    Exit Sub
ErrHandler:
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    mTableDesignBatchPrepared = previousDesignBatchPrepared
    If IsCancellationError(errNumber) Then Err.Raise errNumber, errSource, errDescription
    AddTraceNote "ERROR in RepairDataTablesCoreForSetup. Step=" & currentStep & "; Err.Number=" & CStr(errNumber) & "; Err.Source=" & errSource & "; Err.Description=" & SharedMessage_OneLine(errDescription, 600)
    Err.Raise errNumber, "RepairDataTablesCoreForSetup", "Could not repair the A2 physical data-table foundation." & vbCrLf & "Table setup step: " & IIf(Len(Trim$(currentStep)) = 0, "(not recorded)", currentStep) & vbCrLf & "Database: " & databaseName & vbCrLf & "Err.Source: " & errSource & vbCrLf & "Err.Description: " & errDescription
End Sub

Private Sub RetireObsoleteRowKindLookupTableForSetup()
    'A2 owns physical table retirement. If an older B5 relationship still references the obsolete RowKind lookup table, A2 leaves the table intact.
    'B5 retires that relationship later in the same Setup run; the next A2 pass removes the table without crossing phase ownership.
    Dim db As DAO.Database
    Dim rel As DAO.Relation
    Dim relationStillReferencesTable As Boolean
    If CodeDbTableExists(LEGACY_TABLE_PHYSICAL_ROWKIND_LOOKUP) = False Then Exit Sub
    Set db = TablesWorkspaceCodeDb()
    db.Relations.Refresh
    For Each rel In db.Relations
        If StrComp(rel.Table, LEGACY_TABLE_PHYSICAL_ROWKIND_LOOKUP, vbTextCompare) = 0 Or StrComp(rel.foreignTable, LEGACY_TABLE_PHYSICAL_ROWKIND_LOOKUP, vbTextCompare) = 0 Then
            relationStillReferencesTable = True
            Exit For
        End If
    Next rel
    If relationStillReferencesTable Then
        CoreAddOperationNote "A2 deferred obsolete RowKind lookup-table retirement because an existing B5-owned relationship still references [" & LEGACY_TABLE_PHYSICAL_ROWKIND_LOOKUP & "]. B5 removes that relationship in this run; the next A2 pass can then remove the obsolete table."
        Exit Sub
    End If
    SafeDropTableIfExists LEGACY_TABLE_PHYSICAL_ROWKIND_LOOKUP, "Retired obsolete RowKind lookup table after RowKind became a calculated/diagnostic value rather than stored entity state."
End Sub

Private Sub EnsureEditableLookupTablesForSetup(ByVal editableLookupValueTextSizes As Object)
    'A2 owns physical creation/repair of the editable lookup-table family. Built-in definitions seed only brand-new tables; persisted rows are thereafter authoritative and are never re-defaulted here.
    Dim definitions As Collection
    Dim setupDefinitions As Collection
    Dim rowData As Variant
    Dim tableName As String
    Dim valueTextSize As Long
    Set definitions = EditableLookupTableDefinitions()
    Set setupDefinitions = EditableLookupTableDefinitionsForSetup(definitions)
    RetireConfigOwnerDatabaseRoleLookupQueryForSetup
    For Each rowData In setupDefinitions
        tableName = CStr(rowData(0))
        valueTextSize = CLng(editableLookupValueTextSizes(tableName))
        CorePulseProgress "Table foundation: " & CStr(rowData(3))
        EnsureEditableLookupTable tableName, EditableLookupTableSeedDefinitions(tableName, CStr(rowData(1))), valueTextSize, "EditableLookupTableSeedDefinitions"
    Next rowData
    EnsureLookupDefinitionRegistryForSetup
End Sub

Private Sub EnsureLookupDefinitionRegistryForSetup()
    'A2 owns only the physical ADC_LookupDefinitions table schema. B2 owns persistent lookup-editor policy/intent rows. Before A2 appends a missing B2-owned field to an existing registry, it durably records that
    'exact field name so B2 can initialize it without inferring provenance from stored values.
    Dim db As DAO.Database, tdf As DAO.tableDef, fld As DAO.Field
    Dim tableName As String
    Dim pendingAddedFieldNames As String
    Dim pendingPropertyExists As Boolean
    tableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LookupDefinitions")
    Set db = TablesWorkspaceCodeDb()
    pendingAddedFieldNames = DatabaseTextPropertyValue(db, A2_B2_LOOKUP_DEFINITION_PENDING_FIELDS_PROPERTY, pendingPropertyExists)
    If SharedDao_TableExists(db, tableName) = False Then
        Set tdf = db.CreateTableDef(tableName)
        Set fld = tdf.CreateField(FIELD_ID, dbLong): fld.Attributes = dbAutoIncrField: tdf.fields.Append fld
        tdf.fields.Append tdf.CreateField(FIELD_LOOKUP_LOGICAL_ID, dbText, 100): tdf.fields.Append tdf.CreateField(FIELD_LOOKUP_DISPLAY_NAME, dbText, 150)
        tdf.fields.Append tdf.CreateField(FIELD_LOOKUP_TABLE_NAME, dbText, 150): tdf.fields.Append tdf.CreateField(FIELD_LOOKUP_EDITOR_FORM_NAME, dbText, 150)
        tdf.fields.Append tdf.CreateField(FIELD_EDITOR_MODE, dbText, 20): tdf.fields.Append tdf.CreateField(FIELD_ALLOW_ADDITIONS, dbBoolean)
        tdf.fields.Append tdf.CreateField(FIELD_ALLOW_DELETIONS, dbBoolean): tdf.fields.Append tdf.CreateField(FIELD_ALLOW_CODE_EDITS, dbBoolean)
        tdf.fields.Append tdf.CreateField(FIELD_ALLOW_DISPLAY_NAME_EDITS, dbBoolean): tdf.fields.Append tdf.CreateField(FIELD_ALLOW_DESCRIPTION_EDITS, dbBoolean)
        tdf.fields.Append tdf.CreateField(FIELD_ALLOW_SORT_ORDER_EDITS, dbBoolean): tdf.fields.Append tdf.CreateField(FIELD_ALLOW_ACTIVE_EDITS, dbBoolean)
        tdf.fields.Append tdf.CreateField(FIELD_FINALIZED_ON, dbDate): tdf.fields.Append tdf.CreateField(FIELD_FINALIZED_BY, dbText, 100)
        tdf.fields.Append tdf.CreateField(FIELD_MANAGED_OBJECT_DEPENDENCY_NOTES, dbMemo): db.TableDefs.Append tdf
        TrackSetupCreatedTableRollback tableName
        pendingAddedFieldNames = vbNullString
        EnsureDatabasePropertyRemoved db, A2_B2_LOOKUP_DEFINITION_PENDING_FIELDS_PROPERTY, "EnsureLookupDefinitionRegistryForSetup"
    Else
        Set tdf = db.TableDefs(tableName)
        Call EnsureLookupDefinitionRegistryField(tdf, FIELD_ID, dbLong, 0, True)
        Call EnsureLookupDefinitionRegistryField(tdf, FIELD_LOOKUP_LOGICAL_ID, dbText, 100, False)
        EnsureLookupDefinitionRegistryB2OwnedField db, tdf, pendingAddedFieldNames, FIELD_LOOKUP_DISPLAY_NAME, dbText, 150
        EnsureLookupDefinitionRegistryB2OwnedField db, tdf, pendingAddedFieldNames, FIELD_LOOKUP_TABLE_NAME, dbText, 150
        EnsureLookupDefinitionRegistryB2OwnedField db, tdf, pendingAddedFieldNames, FIELD_LOOKUP_EDITOR_FORM_NAME, dbText, 150
        EnsureLookupDefinitionRegistryB2OwnedField db, tdf, pendingAddedFieldNames, FIELD_EDITOR_MODE, dbText, 20
        EnsureLookupDefinitionRegistryB2OwnedField db, tdf, pendingAddedFieldNames, FIELD_ALLOW_ADDITIONS, dbBoolean, 0
        EnsureLookupDefinitionRegistryB2OwnedField db, tdf, pendingAddedFieldNames, FIELD_ALLOW_DELETIONS, dbBoolean, 0
        EnsureLookupDefinitionRegistryB2OwnedField db, tdf, pendingAddedFieldNames, FIELD_ALLOW_CODE_EDITS, dbBoolean, 0
        EnsureLookupDefinitionRegistryB2OwnedField db, tdf, pendingAddedFieldNames, FIELD_ALLOW_DISPLAY_NAME_EDITS, dbBoolean, 0
        EnsureLookupDefinitionRegistryB2OwnedField db, tdf, pendingAddedFieldNames, FIELD_ALLOW_DESCRIPTION_EDITS, dbBoolean, 0
        EnsureLookupDefinitionRegistryB2OwnedField db, tdf, pendingAddedFieldNames, FIELD_ALLOW_SORT_ORDER_EDITS, dbBoolean, 0
        EnsureLookupDefinitionRegistryB2OwnedField db, tdf, pendingAddedFieldNames, FIELD_ALLOW_ACTIVE_EDITS, dbBoolean, 0
        Call EnsureLookupDefinitionRegistryField(tdf, FIELD_FINALIZED_ON, dbDate, 0, False)
        Call EnsureLookupDefinitionRegistryField(tdf, FIELD_FINALIZED_BY, dbText, 100, False)
        Call EnsureLookupDefinitionRegistryField(tdf, FIELD_MANAGED_OBJECT_DEPENDENCY_NOTES, dbMemo, 0, False)
        If Len(pendingAddedFieldNames) = 0 And pendingPropertyExists Then
            EnsureDatabasePropertyRemoved db, A2_B2_LOOKUP_DEFINITION_PENDING_FIELDS_PROPERTY, "EnsureLookupDefinitionRegistryForSetup"
        End If
    End If
    EnsureSingleFieldIndex db, tableName, INDEX_PRIMARY_KEY, FIELD_ID, True, True, mTableDesignBatchPrepared
    EnsureSingleFieldIndex db, tableName, INDEX_UX_LOOKUP_LOGICAL_ID, FIELD_LOOKUP_LOGICAL_ID, True, False, mTableDesignBatchPrepared
    Set tdf = db.TableDefs(tableName)
    EnsureTableDefFieldOrderFromText tdf, FIELD_ID & "|" & FIELD_LOOKUP_LOGICAL_ID & "|" & FIELD_LOOKUP_DISPLAY_NAME & "|" & FIELD_LOOKUP_TABLE_NAME & "|" & FIELD_LOOKUP_EDITOR_FORM_NAME & "|" & FIELD_EDITOR_MODE & "|" & FIELD_ALLOW_ADDITIONS & "|" & FIELD_ALLOW_DELETIONS & "|" & FIELD_ALLOW_CODE_EDITS & "|" & FIELD_ALLOW_DISPLAY_NAME_EDITS & "|" & FIELD_ALLOW_DESCRIPTION_EDITS & "|" & FIELD_ALLOW_SORT_ORDER_EDITS & "|" & FIELD_ALLOW_ACTIVE_EDITS & "|" & FIELD_FINALIZED_ON & "|" & FIELD_FINALIZED_BY & "|" & FIELD_MANAGED_OBJECT_DEPENDENCY_NOTES, mTableDesignBatchPrepared
End Sub

Private Sub EnsureLookupDefinitionRegistryB2OwnedField(ByVal db As DAO.Database, ByVal tdf As DAO.tableDef, ByRef pendingFieldNames As String, ByVal fieldName As String, ByVal fieldType As DAO.DataTypeEnum, ByVal fieldSize As Long)
    'Persists the A2-to-B2 handoff before the physical append so a later A2 failure cannot strand a newly introduced B2-owned field without provenance.
    If SharedDao_TableDefFieldExistsSafe(tdf, fieldName) Then
        Call EnsureLookupDefinitionRegistryField(tdf, fieldName, fieldType, fieldSize, False)
        Exit Sub
    End If
    AppendLookupDefinitionPendingField pendingFieldNames, fieldName
    Call EnsureDaoPropertyValue(db, A2_B2_LOOKUP_DEFINITION_PENDING_FIELDS_PROPERTY, dbText, pendingFieldNames, "EnsureLookupDefinitionRegistryForSetup")
    Call EnsureLookupDefinitionRegistryField(tdf, fieldName, fieldType, fieldSize, False)
End Sub

Private Sub EnsureLookupDefinitionRegistryField(ByVal tdf As DAO.tableDef, ByVal fieldName As String, ByVal fieldType As DAO.DataTypeEnum, ByVal fieldSize As Long, ByVal autoNumber As Boolean)
    Dim fld As DAO.Field
    If SharedDao_TableDefFieldExistsSafe(tdf, fieldName) Then
        Set fld = tdf.fields(fieldName)
        If SharedDao_FieldDefinitionMatches(fld, fieldType, fieldSize, autoNumber) = False Then Err.Raise vbObjectError + 9612, "EnsureLookupDefinitionRegistryField", "The lookup-policy registry field has an incompatible definition and cannot be safely altered in place." & vbCrLf & "Table: " & tdf.Name & vbCrLf & "Field: " & fieldName & vbCrLf & "Expected: " & FieldDefinitionDescription(fieldType, fieldSize, autoNumber) & vbCrLf & "Found: " & FieldActualDefinitionDescription(fld)
        Exit Sub
    End If
    If autoNumber Then Err.Raise vbObjectError + 9613, "EnsureLookupDefinitionRegistryField", "The lookup-policy registry is missing its AutoNumber identity field; safe repair requires restoring the Preflight backup." & vbCrLf & "Table: " & tdf.Name & vbCrLf & "Field: " & fieldName
    If fieldSize > 0 Then Set fld = tdf.CreateField(fieldName, fieldType, fieldSize) Else Set fld = tdf.CreateField(fieldName, fieldType)
    tdf.fields.Append fld
End Sub

Private Sub AppendLookupDefinitionPendingField(ByRef pendingFieldNames As String, ByVal fieldName As String)
    'Maintains a compact, de-duplicated A2->B2 handoff list of B2-owned fields physically introduced into an existing lookup-definition registry.
    If Len(fieldName) = 0 Then Exit Sub
    If InStr(1, "|" & pendingFieldNames & "|", "|" & fieldName & "|", vbTextCompare) > 0 Then Exit Sub
    If Len(pendingFieldNames) > 0 Then pendingFieldNames = pendingFieldNames & "|"
    pendingFieldNames = pendingFieldNames & fieldName
End Sub

Private Function EditableLookupValueTextSizeIndex(ByVal configSchemaDefinitions As Collection, ByVal healthReviewSchemaDefinitions As Collection, ByVal launcherReviewSchemaDefinitions As Collection) As Object
    'Derives every editable lookup-table key width once from the canonical destination schemas; the narrowest destination remains each persisted vocabulary contract.
    Dim valueTextSizes As Object
    Dim setupDefinitions As Collection
    Dim rowData As Variant
    Dim lookupTableName As String
    Set valueTextSizes = CreateObject("Scripting.Dictionary")
    valueTextSizes.CompareMode = vbTextCompare
    Set setupDefinitions = EditableLookupTableDefinitionsForSetup()
    For Each rowData In setupDefinitions
        lookupTableName = Trim$(CStr(rowData(0)))
        If Len(lookupTableName) = 0 Then Err.Raise vbObjectError + 9600, "EditableLookupValueTextSizeIndex", "The canonical editable lookup-table definition contains a blank table name."
        If valueTextSizes.Exists(lookupTableName) = False Then
            If UBound(rowData) >= 4 Then
                valueTextSizes.Add lookupTableName, CLng(rowData(4))
            Else
                valueTextSizes.Add lookupTableName, CLng(0)
            End If
        End If
    Next rowData
    ResolveEditableLookupValueTextSizesFromSchema ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config"), configSchemaDefinitions, valueTextSizes
    ResolveEditableLookupValueTextSizesFromSchema ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates"), CandidateFieldSchemaDefinitions(), valueTextSizes
    ResolveEditableLookupValueTextSizesFromSchema ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReview"), healthReviewSchemaDefinitions, valueTextSizes
    ResolveEditableLookupValueTextSizesFromSchema ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LauncherControlsReview"), launcherReviewSchemaDefinitions, valueTextSizes
    For Each rowData In setupDefinitions
        lookupTableName = CStr(rowData(0))
        If CLng(valueTextSizes(lookupTableName)) <= 0 Then Err.Raise vbObjectError + 9600, "EditableLookupValueTextSizeIndex", "No canonical destination Short Text field was found for the editable lookup vocabulary." & vbCrLf & "Lookup table: " & lookupTableName
    Next rowData
    Set EditableLookupValueTextSizeIndex = valueTextSizes
End Function

Private Sub ResolveEditableLookupValueTextSizesFromSchema(ByVal sourceTableName As String, ByVal schemaDefinitions As Collection, ByVal valueTextSizes As Object)
    Dim fieldDefinition As Variant
    Dim fieldName As String
    Dim lookupDefinition As Variant
    Dim lookupTableName As String
    Dim fieldSize As Long
    Dim resolvedSize As Long
    For Each fieldDefinition In schemaDefinitions
        fieldName = CStr(fieldDefinition(0))
        lookupDefinition = Empty
        If LookupFieldRequiresTableOrQuerySource(fieldName) Then lookupDefinition = FieldTableLookupDefinition(sourceTableName, fieldName)
        If IsArray(lookupDefinition) Then
            If StrComp(CStr(lookupDefinition(0)), FIELD_LOOKUP_DEFINITION_KIND_TABLE, vbTextCompare) = 0 Then
                lookupTableName = CStr(lookupDefinition(1))
                If valueTextSizes.Exists(lookupTableName) Then
                    If CLng(fieldDefinition(1)) <> CLng(dbText) Or CLng(fieldDefinition(2)) <= 0 Then Err.Raise vbObjectError + 9601, "ResolveEditableLookupValueTextSizesFromSchema", "An editable lookup vocabulary is assigned to a destination field that is not a bounded Short Text field." & vbCrLf & "Lookup table: " & lookupTableName & vbCrLf & "Destination table: " & sourceTableName & vbCrLf & "Destination field: " & fieldName
                    fieldSize = CLng(fieldDefinition(2))
                    resolvedSize = CLng(valueTextSizes(lookupTableName))
                    If resolvedSize = 0 Or fieldSize < resolvedSize Then valueTextSizes(lookupTableName) = fieldSize
                End If
            End If
        End If
    Next fieldDefinition
End Sub

Private Sub RetireConfigOwnerDatabaseRoleLookupQueryForSetup()
    'The former generated DB-role lookup query is unnecessary because Config owner-role consumers use the existing editable MAC_Lookup_DBRoles table directly.
    Dim db As DAO.Database
    Dim previousQuerySql As String
    Dim deleteErrorText As String
    Set db = TablesWorkspaceCodeDb()
    If SharedDao_QueryExists(db, RETIRED_CONFIG_OWNER_DATABASE_ROLE_LOOKUP_QUERY) = False Then Exit Sub
    If CodeDbIsHostDb() Then
        If AccessObjectExistsInCurrentDatabase(acQuery, RETIRED_CONFIG_OWNER_DATABASE_ROLE_LOOKUP_QUERY) Then
            If SharedAccess_PrepareOpenObjectForDesignChange(acQuery, RETIRED_CONFIG_OWNER_DATABASE_ROLE_LOOKUP_QUERY, "retiring the obsolete database-role lookup query") = False Then
                MarkSystemCommandCancelled
                Err.Raise vbObjectError + 9574, "RetireConfigOwnerDatabaseRoleLookupQueryForSetup", "Lookup-query retirement was cancelled because the affected open query was not approved for save and close."
            End If
        End If
    End If
    previousQuerySql = CStr(db.QueryDefs(RETIRED_CONFIG_OWNER_DATABASE_ROLE_LOOKUP_QUERY).sql)
    TrackSetupQueryRollback RETIRED_CONFIG_OWNER_DATABASE_ROLE_LOOKUP_QUERY, True, previousQuerySql
    If TryDeleteQueryDef(db, RETIRED_CONFIG_OWNER_DATABASE_ROLE_LOOKUP_QUERY, deleteErrorText) = False Then Err.Raise vbObjectError + 9575, "RetireConfigOwnerDatabaseRoleLookupQueryForSetup", "Could not retire the obsolete generated database-role lookup query." & vbCrLf & "Query: " & RETIRED_CONFIG_OWNER_DATABASE_ROLE_LOOKUP_QUERY & vbCrLf & deleteErrorText
    CoreAddOperationNote "Retired obsolete generated database-role lookup query [" & RETIRED_CONFIG_OWNER_DATABASE_ROLE_LOOKUP_QUERY & "]; database-role lookups use MAC_Lookup_DBRoles directly."
End Sub

Private Sub EnsureEditableLookupTable(ByVal tableName As String, ByVal seedDefinitions As Collection, ByVal valueTextSize As Long, ByVal seedDefaultSourceName As String)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim fld As DAO.Field
    Dim seedRow As Variant
    Dim createdThisCall As Boolean
    Dim previousQuerySql As String
    If Len(tableName) = 0 Then Err.Raise vbObjectError + 9571, "EnsureEditableLookupTable", "The editable lookup-table name is blank."
    If seedDefinitions Is Nothing Then Err.Raise vbObjectError + 9572, "EnsureEditableLookupTable", "The initial seed definition collection is missing." & vbCrLf & "Table: " & tableName
    If valueTextSize <= 0 Or valueTextSize > 255 Then Err.Raise vbObjectError + 9604, "EnsureEditableLookupTable", "The editable lookup-table ValueText size is outside the Access Short Text range." & vbCrLf & "Table: " & tableName & vbCrLf & "ValueText size: " & CStr(valueTextSize)
    For Each seedRow In seedDefinitions
        If Len(CStr(seedRow(0))) > valueTextSize Then Err.Raise vbObjectError + 9605, "EnsureEditableLookupTable", "A built-in editable lookup seed value exceeds the authoritative destination-field length." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & FIELD_VALUE_TEXT & vbCrLf & "Maximum length: " & CStr(valueTextSize) & vbCrLf & "Value length: " & CStr(Len(CStr(seedRow(0)))) & vbCrLf & "Default/value source: " & seedDefaultSourceName & vbCrLf & "Seed value: " & CStr(seedRow(0))
    Next seedRow
    Set db = TablesWorkspaceCodeDb()
    If SharedDao_QueryExists(db, tableName) Then
        If CodeDbIsHostDb() Then
            If AccessObjectExistsInCurrentDatabase(acQuery, tableName) Then
                If SharedAccess_PrepareOpenObjectForDesignChange(acQuery, tableName, "replacing a generated lookup query with its editable lookup table") = False Then
                    MarkSystemCommandCancelled
                    Err.Raise vbObjectError + 9576, "EnsureEditableLookupTable", "Lookup-table conversion was cancelled because the affected open query was not approved for save and close." & vbCrLf & "Query: " & tableName
                End If
            End If
        End If
        previousQuerySql = CStr(db.QueryDefs(tableName).sql)
        TrackSetupQueryRollback tableName, True, previousQuerySql
    End If
    PrepareTableNameForCreateOrRepair tableName, "editable lookup"
    If CodeDbTableExists(tableName) = False Then
        Set tdf = db.CreateTableDef(tableName)
        Set fld = tdf.CreateField(FIELD_VALUE_TEXT, dbText, valueTextSize)
        fld.Required = True
        fld.allowZeroLength = False
        tdf.fields.Append fld
        Set fld = tdf.CreateField(FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER, dbLong)
        fld.Required = True
        tdf.fields.Append fld
        Set fld = tdf.CreateField(FIELD_DESCRIPTION, dbMemo)
        fld.allowZeroLength = True
        tdf.fields.Append fld
        db.TableDefs.Append tdf
        TrackSetupCreatedTableRollback tableName
        'Append updates this batch-owned TableDefs collection immediately; defer a full refresh to the A2 batch boundary.
        If mTableDesignBatchPrepared = False Then db.TableDefs.Refresh
        createdThisCall = True
        CoreAddOperationNote "Created persistent editable lookup table [" & tableName & "] with ValueText size " & CStr(valueTextSize) & "."
    Else
        RequirePreparedTableCanOpen tableName, "editable lookup"
        EnsureEditableLookupValueTextField db, tableName, valueTextSize
        RepairTableField tableName, FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER, dbLong
        RepairTableField tableName, FIELD_DESCRIPTION, dbMemo
        Set tdf = db.TableDefs(tableName)
        tdf.fields(FIELD_VALUE_TEXT).Required = True
        tdf.fields(FIELD_VALUE_TEXT).allowZeroLength = False
        tdf.fields(FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER).Required = True
        tdf.fields(FIELD_DESCRIPTION).allowZeroLength = True
    End If
    EnsureSingleFieldIndex db, tableName, INDEX_PRIMARY_KEY, FIELD_VALUE_TEXT, True, True, mTableDesignBatchPrepared
    EnsureSingleFieldIndex db, tableName, INDEX_IX_SORT_ORDER, FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER, False, False, mTableDesignBatchPrepared
    Set tdf = db.TableDefs(tableName)
    EnsureTableDefFieldOrderFromText tdf, FIELD_VALUE_TEXT & "|" & FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER & "|" & FIELD_DESCRIPTION, mTableDesignBatchPrepared
    If createdThisCall Then
        SeedEditableLookupTable db, tableName, seedDefinitions
    Else
        CoreAddOperationNote "Preserved existing editable lookup rows in [" & tableName & "]; Setup did not re-seed or restore values."
    End If
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrSource As String
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrSource = Err.Source
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "EnsureEditableLookupTable", "Could not create or repair an editable lookup table." & vbCrLf & "Table: " & tableName & vbCrLf & "Err.Source: " & originalErrSource & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

Private Sub EnsureEditableLookupValueTextField(ByVal db As DAO.Database, ByVal tableName As String, ByVal valueTextSize As Long)
    On Error GoTo ErrHandler
    Dim tdf As DAO.tableDef
    Dim fld As DAO.Field
    Set tdf = db.TableDefs(tableName)
    If SharedDao_TableDefFieldExistsSafe(tdf, FIELD_VALUE_TEXT) = False Then
        RepairTableField tableName, FIELD_VALUE_TEXT, dbText, valueTextSize
        Exit Sub
    End If
    Set fld = tdf.fields(FIELD_VALUE_TEXT)
    If fld.Type <> dbText Then Err.Raise vbObjectError + 9606, "EnsureEditableLookupValueTextField", "The existing editable lookup ValueText field is not Short Text; ordinary repair will not coerce an incompatible field type." & vbCrLf & "Table: " & tableName & vbCrLf & "Actual field type: " & CStr(fld.Type)
    If CLng(fld.Size) = valueTextSize Then Exit Sub
    ValidateEditableLookupValueTextLength db, tableName, valueTextSize
    db.Execute "ALTER TABLE [" & Replace(tableName, "]", "]]") & "] ALTER COLUMN " & SQL_FIELD_VALUE_TEXT & " TEXT(" & CStr(valueTextSize) & ");", dbFailOnError
    db.TableDefs.Refresh
    Set tdf = db.TableDefs(tableName)
    tdf.fields.Refresh
    Set fld = tdf.fields(FIELD_VALUE_TEXT)
    If SharedDao_FieldDefinitionMatches(fld, dbText, valueTextSize, False) = False Then Err.Raise vbObjectError + 9603, "EnsureEditableLookupValueTextField", "The editable lookup ValueText field did not match its authoritative Short Text definition after ALTER COLUMN." & vbCrLf & "Table: " & tableName & vbCrLf & "Expected size: " & CStr(valueTextSize)
    CoreAddOperationNote "Repaired editable lookup field [" & tableName & "]." & SQL_FIELD_VALUE_TEXT & " to Short Text(" & CStr(valueTextSize) & ") without changing persisted lookup rows."
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "EnsureEditableLookupValueTextField", "Could not enforce the authoritative editable lookup ValueText field definition." & vbCrLf & "Table: " & tableName & vbCrLf & "Expected size: " & CStr(valueTextSize) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub ValidateEditableLookupValueTextLength(ByVal db As DAO.Database, ByVal tableName As String, ByVal valueTextSize As Long)
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim offendingValue As String
    sqlText = "SELECT TOP 1 " & SQL_FIELD_VALUE_TEXT & " FROM " & SharedDao_BracketName(tableName) & " WHERE Len(Nz(" & SQL_FIELD_VALUE_TEXT & ",''))>" & CStr(valueTextSize) & ";"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    If rs.EOF = False Then
        offendingValue = CStr(rs.fields(FIELD_VALUE_TEXT).value)
        rs.Close
        Set rs = Nothing
        Err.Raise vbObjectError + 9602, "ValidateEditableLookupValueTextLength", "An existing editable lookup value is too long for the field that stores this vocabulary. Shorten or delete the value before rerunning Setup." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & FIELD_VALUE_TEXT & vbCrLf & "Maximum length: " & CStr(valueTextSize) & vbCrLf & "Default/value source: (not applicable; existing persisted row)" & vbCrLf & "Existing value length: " & CStr(Len(offendingValue)) & vbCrLf & "Existing value: " & offendingValue
    End If
    rs.Close
    Set rs = Nothing
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
    Err.Raise savedErrNumber, "ValidateEditableLookupValueTextLength", "Could not validate editable lookup values before enforcing the authoritative ValueText length." & vbCrLf & "Table: " & tableName & vbCrLf & "Maximum length: " & CStr(valueTextSize) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub SeedEditableLookupTable(ByVal db As DAO.Database, ByVal tableName As String, ByVal seedDefinitions As Collection)
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim rowData As Variant
    Dim SortOrder As Long
    If seedDefinitions Is Nothing Then Err.Raise vbObjectError + 9573, "SeedEditableLookupTable", "The initial seed definition collection is missing." & vbCrLf & "Table: " & tableName
    Set rs = db.OpenRecordset(tableName, dbOpenDynaset)
    For Each rowData In seedDefinitions
        SortOrder = SortOrder + 1
        rs.AddNew
        rs.fields(FIELD_VALUE_TEXT).value = CStr(rowData(0))
        rs.fields(FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER).value = SortOrder
        rs.fields(FIELD_DESCRIPTION).value = CStr(rowData(1))
        rs.Update
    Next rowData
    rs.Close
    Set rs = Nothing
    CoreAddOperationNote "Seeded newly created editable lookup table [" & tableName & "] with " & CStr(SortOrder) & " built-in value(s)."
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrSource As String
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrSource = Err.Source
    originalErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
    Err.Raise originalErrNumber, "SeedEditableLookupTable", "Could not seed a newly created editable lookup table." & vbCrLf & "Table: " & tableName & vbCrLf & "Err.Source: " & originalErrSource & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

' Purpose: Repairs health review table for setup for apply lookup metadata.
' Public role: Public module entry point. This procedure is listed in the module header as an intentional externally callable ADC entry point and may be invoked directly by Access/VBA callers.
' Parameters:
'   applyLookupMetadata (Input Optional; Boolean): When True, enables or requests lookup metadata for this call; when False, the routine follows its normal non-forced/default path. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub RepairHealthReviewTableForSetup(Optional ByVal applyLookupMetadata As Boolean = True)
    RequireSetupDomainMutationAllowed SETUP_DOMAIN_A2_PHYSICAL_SCHEMA, "RepairHealthReviewTableForSetup"
    RequireCatalogMutationRecoveryReady "RepairHealthReviewTableForSetup"
    On Error GoTo ErrHandler
    Dim schemaDefinitions As Collection
    Set schemaDefinitions = BuildHealthReviewSchemaDefinitions()
    If CodeDbTableExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReview")) = False Then
        CreateHealthReviewTable ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReview"), schemaDefinitions
    Else
        EnsureHealthReviewFields ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReview"), schemaDefinitions
    End If
    If applyLookupMetadata Then EnsureHealthReviewLookupFields ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReview")
    EnsureHealthReviewIndexes ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReview")
    ReconcileStandaloneManagedTableDescriptionIfAllowed ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReview")
    LogTraceMessage TRACE_LEVEL_INFO, "TableRepair", "RepairHealthReviewTableForSetup", "A2 health-review storage repaired and verified.", "Table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReview"), "Repair A2 health-review storage", 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReview"), "Table", "CodeDb", "TableRepair"
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "RepairHealthReviewTableForSetup", "Could not create or repair A2 health-review storage." & vbCrLf & "Table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReview") & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

' Purpose: Resolves health review field order text from the supplied/current ADC state, using `BuildHealthReviewSchemaDefinitions` and `Next` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function HealthReviewFieldOrderText() As String
    ' Wrapper purpose: TABLES-owned facade returning HealthReview canonical physical field order from its authoritative schema definitions.
    ' Why retained: Retained so external schema consumers do not access private TABLES schema machinery or maintain an independent HealthReview field-order mirror.
    'Intended use: Returns HealthReview canonical physical field order directly from the authoritative HealthReview schema definitions.
    Dim fieldDefinition As Variant
    Dim valueText As String
    For Each fieldDefinition In BuildHealthReviewSchemaDefinitions()
        If Len(valueText) > 0 Then valueText = valueText & "|"
        valueText = valueText & CStr(fieldDefinition(0))
    Next fieldDefinition
    HealthReviewFieldOrderText = valueText
End Function

Private Function BuildHealthReviewSchemaDefinitions() As Collection
    'TABLES-owned single physical field-schema authority for HealthReview creation, structural repair, contract enforcement, and field-name enumeration; G1 owns managed DAO field descriptions.
    ' CANONICAL AUTHORITY: rows owns the in-memory definition set constructed by BuildHealthReviewSchemaDefinitions; downstream consumers derive views from this set rather than create a competing authority.
    Dim rows As New Collection
    rows.Add Array(FIELD_ID, dbLong, 0, True)
    rows.Add Array(FIELD_REVIEW_RUN_ID, dbText, 100, False)
    rows.Add Array(FIELD_REVIEW_RUN_ON, dbDate, 0, False)
    rows.Add Array(FIELD_IS_CURRENT_RUN, dbBoolean, 0, False)
    rows.Add Array(FIELD_SEVERITY, dbText, 30, False, Null, Null, Null, Null, Null)
    rows.Add Array(FIELD_REVIEW_AREA, dbText, 100, False, Null, Null, Null, Null, Null)
    rows.Add Array(FIELD_ISSUE_CODE, dbText, 100, False)
    rows.Add Array(FIELD_DETAILS, dbMemo, 0, False)
    rows.Add Array(FIELD_RECOMMENDED_ACTION, dbMemo, 0, False)
    rows.Add Array(FIELD_OBJECT_TYPE, dbText, 100, False)
    rows.Add Array(FIELD_OBJECT_NAME, dbText, 255, False)
    rows.Add Array(FIELD_PROCEDURE_NAME, dbText, 255, False)
    rows.Add Array(FIELD_CONFIG_ID, dbLong, 0, False)
    rows.Add Array(FIELD_MACRO_KEY, dbText, 100, False)
    rows.Add Array(FIELD_MACRO_TITLE, dbText, 255, False)
    rows.Add Array(FIELD_SETTING_KEY, dbText, 255, False)
    rows.Add Array(FIELD_ROW_KIND, dbText, 30, False)
    rows.Add Array(FIELD_CONFIGURATION_STATUS, dbText, 30, False)
    rows.Add Array(FIELD_CANDIDATE_STATUS, dbText, 100, False)
    rows.Add Array(FIELD_CONFIG_OWNER_DATABASE_ROLE, dbText, 30, False)
    rows.Add Array(FIELD_CANDIDATE_OWNER_DATABASE_ROLE, dbText, 30, False)
    rows.Add Array(FIELD_TARGET_PROCEDURE_NAME, dbText, 255, False)
    rows.Add Array(FIELD_TARGET_PROCEDURE_MODULE_NAME, dbText, 255, False)
    rows.Add Array(FIELD_GENERATED_MACRO_OBJECT_NAME, dbText, 255, False)
    rows.Add Array(FIELD_EXPECTED_VALUE, dbMemo, 0, False)
    rows.Add Array(FIELD_ACTUAL_VALUE, dbMemo, 0, False)
    rows.Add Array(FIELD_ISSUE_HASH, dbText, 255, False)
    rows.Add Array(FIELD_CREATED_ON, dbDate, 0, False)
    Set BuildHealthReviewSchemaDefinitions = rows
End Function

Private Sub CreateHealthReviewTable(ByVal tableName As String, ByVal schemaDefinitions As Collection)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim fld As DAO.Field
    Dim fieldDefinition As Variant
    Set db = TablesWorkspaceCodeDb()
    Set tdf = db.CreateTableDef(tableName)
    For Each fieldDefinition In schemaDefinitions
        If CLng(fieldDefinition(1)) = CLng(dbText) Then
            Set fld = tdf.CreateField(CStr(fieldDefinition(0)), CLng(fieldDefinition(1)), CLng(fieldDefinition(2)))
        Else
            Set fld = tdf.CreateField(CStr(fieldDefinition(0)), CLng(fieldDefinition(1)))
        End If
        If CBool(fieldDefinition(3)) Then fld.Attributes = dbAutoIncrField
        ApplySchemaFieldContract fld, fieldDefinition, tableName, "health-review"
        tdf.fields.Append fld
    Next fieldDefinition
    db.TableDefs.Append tdf
    TrackSetupCreatedTableRollback tableName
    db.TableDefs.Refresh
    EnsureTableDefFieldOrderFromText db.TableDefs(tableName), SchemaDefinitionFieldOrderText(schemaDefinitions), mTableDesignBatchPrepared
    CoreAddOperationNote "Created health-review output table [" & tableName & "]."
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "CreateHealthReviewTable", "Could not create health-review output table." & vbCrLf & "Table: " & tableName & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

Private Sub EnsureHealthReviewFields(ByVal tableName As String, ByVal schemaDefinitions As Collection)
    Dim fieldDefinition As Variant
    Dim tdf As DAO.tableDef
    Set tdf = TablesWorkspaceCodeDb().TableDefs(tableName)
    For Each fieldDefinition In schemaDefinitions
        If CBool(fieldDefinition(3)) = False Then
            If SharedDao_TableDefFieldExistsSafe(tdf, CStr(fieldDefinition(0))) Then
                RepairTableField tableName, CStr(fieldDefinition(0)), CLng(fieldDefinition(1)), CLng(fieldDefinition(2))
            Else
                RepairTableField tableName, CStr(fieldDefinition(0)), CLng(fieldDefinition(1)), CLng(fieldDefinition(2))
                tdf.fields.Refresh
            End If
        End If
        ApplySchemaFieldContract tdf.fields(CStr(fieldDefinition(0))), fieldDefinition, tableName, "health-review"
    Next fieldDefinition
    EnsureTableDefFieldOrderFromText TablesWorkspaceCodeDb().TableDefs(tableName), SchemaDefinitionFieldOrderText(schemaDefinitions), mTableDesignBatchPrepared
End Sub

Private Sub EnsureHealthReviewIndexes(ByVal tableName As String)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim duplicateIssueRows As DAO.Recordset
    Set db = TablesWorkspaceCodeDb()
    EnsureSingleFieldIndex db, tableName, INDEX_PRIMARY_KEY, FIELD_ID, True, True, mTableDesignBatchPrepared
    EnsureSingleFieldIndex db, tableName, INDEX_IX_REVIEW_RUN_ID, FIELD_REVIEW_RUN_ID, False, False, mTableDesignBatchPrepared
    EnsureSingleFieldIndex db, tableName, INDEX_IX_IS_CURRENT_RUN, FIELD_IS_CURRENT_RUN, False, False, mTableDesignBatchPrepared
    EnsureSingleFieldIndex db, tableName, INDEX_IX_SEVERITY, FIELD_SEVERITY, False, False, mTableDesignBatchPrepared
    EnsureSingleFieldIndex db, tableName, INDEX_IX_REVIEW_AREA, FIELD_REVIEW_AREA, False, False, mTableDesignBatchPrepared
    EnsureSingleFieldIndex db, tableName, INDEX_IX_ISSUE_CODE, FIELD_ISSUE_CODE, False, False, mTableDesignBatchPrepared
    EnsureSingleFieldIndex db, tableName, INDEX_IX_CONFIG_ID, FIELD_CONFIG_ID, False, False, mTableDesignBatchPrepared
    EnsureSingleFieldIndex db, tableName, INDEX_IX_MACRO_KEY, FIELD_MACRO_KEY, False, False, mTableDesignBatchPrepared
    EnsureSingleFieldIndex db, tableName, INDEX_IX_SETTING_KEY, FIELD_SETTING_KEY, False, False, mTableDesignBatchPrepared
    EnsureSingleFieldIndex db, tableName, INDEX_IX_OBJECT_NAME, FIELD_OBJECT_NAME, False, False, mTableDesignBatchPrepared
    EnsureSingleFieldIndex db, tableName, INDEX_IX_ISSUE_HASH, FIELD_ISSUE_HASH, False, False, mTableDesignBatchPrepared
    Set duplicateIssueRows = db.OpenRecordset("SELECT TOP 1 [" & FIELD_REVIEW_RUN_ID & "],[" & FIELD_ISSUE_HASH & "],Count(*) AS " & FIELD_DUP_COUNT & " FROM [" & tableName & "] WHERE [" & FIELD_REVIEW_RUN_ID & "] Is Not Null AND [" & FIELD_ISSUE_HASH & "] Is Not Null GROUP BY [" & FIELD_REVIEW_RUN_ID & "],[" & FIELD_ISSUE_HASH & "] HAVING Count(*)>1 ORDER BY [" & FIELD_REVIEW_RUN_ID & "],[" & FIELD_ISSUE_HASH & "];", dbOpenSnapshot)
    If duplicateIssueRows.EOF = False Then
        Err.Raise vbObjectError + 5338, "EnsureHealthReviewIndexes", "HealthReview per-run IssueHash uniqueness repair was blocked because duplicate persisted findings already exist. Existing review evidence was not changed." & vbCrLf & "ReviewRunID: " & CStr(Nz(duplicateIssueRows.fields(FIELD_REVIEW_RUN_ID).value, vbNullString)) & vbCrLf & "IssueHash: " & CStr(Nz(duplicateIssueRows.fields(FIELD_ISSUE_HASH).value, vbNullString)) & vbCrLf & "Rows: " & CStr(CLng(duplicateIssueRows.fields(FIELD_DUP_COUNT)))
    End If
    duplicateIssueRows.Close
    Set duplicateIssueRows = Nothing
    EnsureMultiFieldIndex db, tableName, INDEX_UX_HEALTH_REVIEW_RUN_ISSUE_HASH, FIELD_REVIEW_RUN_ID & "|" & FIELD_ISSUE_HASH, True, False
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    On Error Resume Next
    If Not duplicateIssueRows Is Nothing Then duplicateIssueRows.Close
    On Error GoTo 0
    Err.Raise originalErrNumber, "EnsureHealthReviewIndexes", "Could not repair health-review indexes." & vbCrLf & "Table: " & tableName & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

Private Sub EnsureHealthReviewLookupFields(ByVal tableName As String)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim changedCount As Long
    If CodeDbTableExists(tableName) = False Then Exit Sub
    Set db = TablesWorkspaceCodeDb()
    Set tdf = db.TableDefs(tableName)
    tdf.fields.Refresh
    changedCount = changedCount + EnsureDefinedFieldLookup(tdf, FIELD_SEVERITY)
    changedCount = changedCount + EnsureDefinedFieldLookup(tdf, FIELD_REVIEW_AREA)
    changedCount = changedCount + EnsureDefinedFieldLookup(tdf, FIELD_OBJECT_TYPE)
    changedCount = changedCount + EnsureDefinedFieldLookup(tdf, FIELD_CONFIG_ID)
    changedCount = changedCount + EnsureDefinedFieldLookup(tdf, FIELD_SETTING_KEY)
    changedCount = changedCount + EnsureDefinedFieldLookup(tdf, FIELD_ROW_KIND)
    changedCount = changedCount + EnsureDefinedFieldLookup(tdf, FIELD_CONFIGURATION_STATUS)
    changedCount = changedCount + EnsureDefinedFieldLookup(tdf, FIELD_CANDIDATE_STATUS)
    changedCount = changedCount + EnsureDefinedFieldLookup(tdf, FIELD_CONFIG_OWNER_DATABASE_ROLE)
    changedCount = changedCount + EnsureDefinedFieldLookup(tdf, FIELD_CANDIDATE_OWNER_DATABASE_ROLE)
    If changedCount > 0 Then CoreAddOperationNote "Repaired " & CStr(changedCount) & " lookup field(s) on table [" & tableName & "]."
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "EnsureHealthReviewLookupFields", "Could not repair health-review lookup metadata." & vbCrLf & "Table: " & tableName & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

' Purpose: Resolves health review field names from the supplied/current ADC state, using `fieldNames` and `BuildHealthReviewSchemaDefinitions` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: Variant: Returns the health review field names value in the scalar, array, object, Empty, or Null form produced by the routine's existing contract.
Public Function HealthReviewFieldNames() As Variant
    'Intended use: Returns HealthReview field names directly from the single physical schema authority.
    Dim schemaDefinitions As Collection
    Dim fieldDefinition As Variant
    Dim fieldNames() As String
    Dim indexValue As Long
    Set schemaDefinitions = BuildHealthReviewSchemaDefinitions()
    ReDim fieldNames(0 To schemaDefinitions.count - 1)
    For Each fieldDefinition In schemaDefinitions
        fieldNames(indexValue) = CStr(fieldDefinition(0))
        indexValue = indexValue + 1
    Next fieldDefinition
    HealthReviewFieldNames = fieldNames
End Function

' Purpose: Repairs core data table foundation for setup for schema repair evidence, require setup preflight backup.
' Public role: Public module entry point. This procedure is listed in the module header as an intentional externally callable ADC entry point and may be invoked directly by Access/VBA callers.
' Parameters:
'   schemaRepairEvidence (Input; TSchemaRepairPreflightEvidence): Preflight evidence describing the schema-repair state already established for this Setup run; the routine uses it to authorize and constrain
'   foundation repair work.
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: False.
'   repairDetailsForMessage (Output Optional; String): Receives repair details for message used for caller diagnostics, verification output, or user-facing reporting. Default: vbNullString.
'   namingInputsChanged (Output Optional; Boolean): Set True when the procedure determines that naming inputs changed applies; set False when that condition is not met. Default: False.
'   requireSetupPreflightBackup (Input Optional; Boolean): True requires setup preflight backup before the operation may proceed; False allows the operation to continue without making that condition mandatory.
'   Default: True.
'   applyStandaloneDescriptions (Input Optional; Boolean): True completes the A2 managed table/field Descriptions when this public coordinator runs outside numbered Setup; False lets a broader caller own
'   a later complete Description pass. Default: True.
' Returns: Boolean: True on the function's affirmative/success branch for repair core data table foundation for setup; False on its explicit negative, unavailable, rejected, or handled-failure branch.
'   not-applicable path.
Public Function RepairCoreDataTableFoundationForSetup(ByRef schemaRepairEvidence As TSchemaRepairPreflightEvidence, Optional ByVal showMessage As Boolean = False, Optional ByRef repairDetailsForMessage As String = vbNullString, Optional ByRef namingInputsChanged As Boolean = False, Optional ByVal requireSetupPreflightBackup As Boolean = True, Optional ByVal applyStandaloneDescriptions As Boolean = True) As Boolean
    RequireSetupDomainMutationAllowed SETUP_DOMAIN_A2_PHYSICAL_SCHEMA, "RepairCoreDataTableFoundationForSetup"
    On Error GoTo ErrHandler
    Dim snapshot As FoundationDomainSnapshot
    Dim refreshedSnapshot As FoundationDomainSnapshot
    Dim db As DAO.Database
    Dim schemaPropertyExists As Boolean
    Dim dataPropertyExists As Boolean
    Dim storedSchemaFingerprint As String
    Dim storedPart2DataFingerprint As String
    Dim schemaMatches As Boolean
    Dim dataMatches As Boolean
    Dim phaseStarted As Single
    Dim schemaRepairDetailsText As String
    Dim messageText As String
    Dim coreSettingDefinitions As Collection
    Dim previousDesignBatchPrepared As Boolean
    Dim deferredPhysicalRetirementPending As Boolean
    previousDesignBatchPrepared = mTableDesignBatchPrepared
    phaseStarted = Timer
    repairDetailsForMessage = vbNullString
    namingInputsChanged = False
    CoreSetTableRepairWarningResult False
    Set coreSettingDefinitions = x__UTIL_ADC_DEFAULTS.GetCoreSettingDefinitions()
    Set db = TablesWorkspaceCodeDb()
    deferredPhysicalRetirementPending = A2DeferredPhysicalRetirementPendingForSetup()
    storedSchemaFingerprint = DatabaseTextPropertyValue(db, FOUNDATION_SCHEMA_FINGERPRINT_PROPERTY, schemaPropertyExists)
    storedPart2DataFingerprint = DatabaseTextPropertyValue(db, FOUNDATION_PART2_DATA_FINGERPRINT_PROPERTY, dataPropertyExists)
    CoreUpdateSetupProgress "Checking the core CodeDb data-table foundation fingerprint and required table inventory."
    If schemaRepairEvidence.IsKnown And schemaRepairEvidence.repairNeeded Then
        AddTraceNote "A2 reused Preflight schema evidence showing repair is required; duplicate pre-mutation fingerprint construction was skipped."
    ElseIf deferredPhysicalRetirementPending Then
        AddTraceNote "A2 full repair required because A2-owned legacy physical retirement is still pending; the persisted fingerprint fast path was bypassed without inspecting or mutating B5 relationship state."
    ElseIf schemaPropertyExists And dataPropertyExists And Len(storedSchemaFingerprint) > 0 And Len(storedPart2DataFingerprint) > 0 Then
        If CalculateFoundationSchemaSnapshot(snapshot, coreSettingDefinitions) Then
            schemaMatches = DatabaseTextPropertyMatches(db, FOUNDATION_SCHEMA_FINGERPRINT_PROPERTY, snapshot.SchemaFingerprint, schemaPropertyExists)
            dataMatches = DatabaseTextPropertyMatches(db, FOUNDATION_PART2_DATA_FINGERPRINT_PROPERTY, snapshot.Part2DataFingerprint, dataPropertyExists)
            If schemaMatches And dataMatches Then
                repairDetailsForMessage = vbCrLf & vbCrLf & "Core data-table foundation fast path: Used" & vbCrLf & _
                                          "Schema=matched" & vbCrLf & _
                                          "Part2Defaults=matched" & vbCrLf & _
                                          "Foundation tables: " & CStr(snapshot.tableCount) & vbCrLf & _
                                          "Verification elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseStarted))
                MarkSettingsReadyForCurrentContext
                If requireSetupPreflightBackup Then ClearSetupPreflightA2SchemaEvidence
                If applyStandaloneDescriptions Then ReconcileStandaloneFoundationTableDescriptions
                RepairCoreDataTableFoundationForSetup = True
                ReportCoreActionSuccess "RepairCoreDataTableFoundationForSetup", "Tables", "Core CodeDb data-table foundation verified successfully.", "Core CodeDb data-table foundation verified by complete schema/default-data fingerprint." & repairDetailsForMessage, "Verify core CodeDb data-table foundation", vbNullString, "FoundationPart2Fingerprint", "CodeDb", "TableRepair", showMessage, "Repair " & DEFAULT_UTILITY_LONG_NAME & " Tables", vbInformation, True, "Tables", False
                Exit Function
            End If
            AddTraceNote "Part 2 full repair required. Schema=" & IIf(schemaMatches, "matched", "changed") & "; Part2Defaults=" & IIf(dataMatches, "matched", "changed")
        End If
    Else
        AddTraceNote "A2 persisted fingerprint proof is absent; skipped impossible compact comparison and entered the authoritative repair path."
    End If
    CoreUpdateSetupProgress "Repairing the core CodeDb data-table foundation while deferring saved-query families to the next setup part."
    PrepareOpenFoundationTablesForDesignChange "repairing the A2 physical data-table foundation"
    'Closing/saving an open foundation table can change DAO metadata after the fingerprint snapshot.
    'Synchronize once here before the A2 batch begins so the pre-repair schema proof cannot reuse stale TableDefs.
    TablesWorkspaceCodeDb().TableDefs.Refresh
    mTableDesignBatchPrepared = True
    RepairDataTablesPublicCoreForSetup False, schemaRepairDetailsText, requireSetupPreflightBackup, schemaRepairEvidence.IsKnown, schemaRepairEvidence.repairNeeded, schemaRepairEvidence.Summary, coreSettingDefinitions
    If requireSetupPreflightBackup Then ClearSetupPreflightA2SchemaEvidence
    mTableDesignBatchPrepared = previousDesignBatchPrepared
    If CalculateFoundationSchemaSnapshot(refreshedSnapshot, coreSettingDefinitions, True) = False Then Err.Raise vbObjectError + 9042, "RepairCoreDataTableFoundationForSetup", "Core data-table repair completed, but the schema/default-data fingerprint could not be recalculated. " & refreshedSnapshot.failureDetail
    AddTraceNote "A2 final fingerprint construction reused the strict TableDefs synchronization performed by the authoritative final structural proof."
    PersistFoundationSchemaSnapshot refreshedSnapshot
    namingInputsChanged = True
    MarkSettingsReadyForCurrentContext
    If applyStandaloneDescriptions Then ReconcileStandaloneFoundationTableDescriptions
    repairDetailsForMessage = schemaRepairDetailsText & vbCrLf & vbCrLf & "Core data-table foundation path: Full repair" & vbCrLf & "Schema/default-data fingerprints: refreshed" & vbCrLf & "Foundation tables: " & CStr(refreshedSnapshot.tableCount) & vbCrLf & "Verification and repair elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseStarted))
    CoreSetTableRepairWarningResult CoreOperationHasWarnings()
    RepairCoreDataTableFoundationForSetup = True
    messageText = "Core CodeDb data-table foundation repaired successfully." & repairDetailsForMessage & CoreOperationSummaryForMessage() & TableDiagnosticsLocationForMessage()
    ReportCoreActionSuccess "RepairCoreDataTableFoundationForSetup", "Tables", "Core CodeDb data-table foundation repair completed successfully.", messageText, "Repair core CodeDb data-table foundation", vbNullString, "FoundationSchemaSet", "CodeDb", "TableRepair", showMessage, "Repair " & DEFAULT_UTILITY_LONG_NAME & " Tables", vbInformation, True, "Tables", False
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    mTableDesignBatchPrepared = previousDesignBatchPrepared
    If requireSetupPreflightBackup Then ClearSetupPreflightA2SchemaEvidence
    RepairCoreDataTableFoundationForSetup = False
    repairDetailsForMessage = vbCrLf & vbCrLf & "Core data-table foundation repair failed." & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
    messageText = "Could not verify or repair the core CodeDb data-table foundation." & repairDetailsForMessage & TableDiagnosticsLocationForMessage() & TableReadinessDiagnosticsForMessage() & CoreOperationSummaryForMessage()
    LogTraceMessage TRACE_LEVEL_ERROR, "Tables", "RepairCoreDataTableFoundationForSetup", "Core CodeDb data-table foundation repair failed.", messageText, "Repair core CodeDb data-table foundation", savedErrNumber, savedErrSource, savedErrDescription, vbNullString, "FoundationSchemaSet", "CodeDb", "TableRepair"
    If showMessage Then
        CoreShowErrorMessage messageText, vbCritical, DEFAULT_TABLE_MAINTENANCE_MESSAGE_TITLE, "RepairCoreDataTableFoundationForSetup", "Tables"
        Exit Function
    End If
    Err.Raise savedErrNumber, "RepairCoreDataTableFoundationForSetup", messageText
End Function

Private Sub EnsureConfigFieldsOnTableDef(ByVal tdf As DAO.tableDef, ByVal schemaDefinitions As Collection)
    Dim fieldDefinition As Variant
    Dim reqSize As Variant
    Dim currentTableName As String
    currentTableName = tdf.Name
    For Each fieldDefinition In schemaDefinitions
        reqSize = DefaultManagedFieldReqSize(currentTableName, CStr(fieldDefinition(0)))
        If SharedDao_TableDefFieldExistsSafe(tdf, CStr(fieldDefinition(0))) And IsNull(reqSize) = False And CLng(fieldDefinition(1)) = CLng(dbText) Then
            If SharedDao_FieldDefinitionMatches(tdf.fields(CStr(fieldDefinition(0))), CLng(fieldDefinition(1)), CLng(fieldDefinition(2)), CBool(fieldDefinition(3))) = False Then
                RepairExplicitBoundedTextFieldInDatabase TablesWorkspaceCodeDb(), currentTableName, CStr(fieldDefinition(0)), CLng(fieldDefinition(2)), True
                Set tdf = TablesWorkspaceCodeDb().TableDefs(currentTableName)
            End If
        End If
        AppendSchemaFieldIfMissing tdf, CStr(fieldDefinition(0)), CLng(fieldDefinition(1)), CLng(fieldDefinition(2)), CBool(fieldDefinition(3)), tdf.Name, fieldDefinition
        ApplySchemaFieldContract tdf.fields(CStr(fieldDefinition(0))), fieldDefinition, tdf.Name, tdf.Name
    Next fieldDefinition
    EnsureTableDefFieldOrderFromText tdf, SchemaDefinitionFieldOrderText(schemaDefinitions), mTableDesignBatchPrepared
End Sub

Private Sub RepairDataTablesPublicCoreForSetup(Optional ByVal showMessage As Boolean = False, Optional ByRef schemaRepairDetailsForMessage As String = vbNullString, Optional ByVal requireSetupPreflightBackup As Boolean = True, Optional ByVal schemaRepairNeededKnown As Boolean = False, Optional ByVal schemaRepairNeededValue As Boolean = False, Optional ByVal schemaRepairSummaryValue As String = vbNullString, Optional ByVal suppliedCoreSettingDefinitions As Variant)
    On Error GoTo ErrHandler
    CoreResetOperationNotes
    Dim schemaRepairNeeded As Boolean
    Dim schemaRepairSummary As String
    Dim messageText As String
    Dim settingsSchemaDefinitions As Collection
    Dim settingsIndexDefinitions As Collection
    Dim configSchemaDefinitions As Collection
    Dim configIndexDefinitions As Collection
    Dim macroCategorySchemaDefinitions As Collection
    Dim macroSubcategorySchemaDefinitions As Collection
    Dim coreSettingDefinitions As Collection
    Dim editableLookupValueTextSizes As Object
    If schemaRepairNeededKnown And SetupPreflightA2EvidenceCurrent() Then
        Set settingsSchemaDefinitions = mSetupPreflightA2SettingsSchemaDefinitions
        Set settingsIndexDefinitions = mSetupPreflightA2SettingsIndexDefinitions
        Set configSchemaDefinitions = mSetupPreflightA2ConfigSchemaDefinitions
        Set configIndexDefinitions = mSetupPreflightA2ConfigIndexDefinitions
        Set macroCategorySchemaDefinitions = mSetupPreflightA2MacroCategorySchemaDefinitions
        Set macroSubcategorySchemaDefinitions = mSetupPreflightA2MacroSubcategorySchemaDefinitions
        Set editableLookupValueTextSizes = mSetupPreflightA2EditableLookupValueTextSizes
        CoreAddOperationNote "A2 reused the TABLES-owned immutable schema-plan evidence materialized by Setup Preflight."
    Else
        Set settingsSchemaDefinitions = SettingsFieldSchemaDefinitions()
        Set settingsIndexDefinitions = SettingsIndexSchemaDefinitions()
        Set configSchemaDefinitions = ConfigFieldSchemaDefinitions()
        Set configIndexDefinitions = ConfigIndexSchemaDefinitions()
        Set macroCategorySchemaDefinitions = BuildMacroCategorySchemaDefinitions()
        Set macroSubcategorySchemaDefinitions = BuildMacroSubcategorySchemaDefinitions()
        Set editableLookupValueTextSizes = EditableLookupValueTextSizeIndex(configSchemaDefinitions, BuildHealthReviewSchemaDefinitions(), CustomMacroLauncherReviewSchemaDefinitions())
    End If
    If isMissing(suppliedCoreSettingDefinitions) = False Then
        If IsObject(suppliedCoreSettingDefinitions) Then Set coreSettingDefinitions = suppliedCoreSettingDefinitions
    End If
    If coreSettingDefinitions Is Nothing Then Set coreSettingDefinitions = x__UTIL_ADC_DEFAULTS.GetCoreSettingDefinitions()
    schemaRepairDetailsForMessage = vbNullString
    If schemaRepairNeededKnown Then
        schemaRepairNeeded = schemaRepairNeededValue
        schemaRepairSummary = schemaRepairSummaryValue
    Else
        schemaRepairNeeded = TableSchemaRepairRequired(schemaRepairSummary, settingsSchemaDefinitions, settingsIndexDefinitions, configSchemaDefinitions, configIndexDefinitions, macroCategorySchemaDefinitions, macroSubcategorySchemaDefinitions, editableLookupValueTextSizes)
    End If
    If requireSetupPreflightBackup Then
        RequireSetupPreflightBackupCoverage schemaRepairNeeded, schemaRepairSummary
        CoreAddOperationNote "A2 consumed the unnumbered Preflight backup decision and performed no backup work inside the numbered phase."
    Else
        CoreAddOperationNote "Runtime table repair reused the caller-owned backup decision before invoking the canonical TABLES structural authority."
    End If
    CoreAddOperationNote "Foundation table repair is structural-only for existing rows: existing MC_Records, Settings, and catalog values are preserved."
    RepairDataTablesCoreForSetup settingsSchemaDefinitions, settingsIndexDefinitions, configSchemaDefinitions, configIndexDefinitions, macroCategorySchemaDefinitions, macroSubcategorySchemaDefinitions, coreSettingDefinitions, editableLookupValueTextSizes, True, schemaRepairNeeded, schemaRepairSummary, mTableDesignBatchPrepared
    schemaRepairDetailsForMessage = SchemaRepairSummaryForMessage(schemaRepairNeeded, schemaRepairSummary)
    If showMessage Then
        messageText = DEFAULT_UTILITY_LONG_NAME & " foundation table repair complete."
        If schemaRepairNeeded And requireSetupPreflightBackup Then messageText = messageText & vbCrLf & vbCrLf & "Backup ownership: unnumbered Setup Preflight" & vbCrLf & "Preflight tables backed up: " & CStr(mSetupPreflightBackupCount)
        messageText = messageText & schemaRepairDetailsForMessage & CoreOperationSummaryForMessage()
        CoreShowMessage messageText, vbInformation, "Repair " & DEFAULT_UTILITY_LONG_NAME & " Foundation Tables", "ShowADCMessage", "Tables"
    End If
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrSource As String
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrSource = Err.Source
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "RepairDataTablesPublicCoreForSetup", "Could not repair the " & DEFAULT_UTILITY_LONG_NAME & " data-table foundation." & vbCrLf & "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB & vbCrLf & "Err.Source: " & originalErrSource & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

' Purpose: Performs macro classification state current for B1 by coordinating `AddMacroClassificationTableSchemaIssues` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   detailText (Output; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
' Returns: Boolean: True when the condition represented by MacroClassificationStateCurrentForB1 is present under the routine's documented inputs/current state; False when that condition is absent.
Public Function MacroClassificationStateCurrentForB1(ByRef detailText As String) As Boolean
    'Read-only B1 predecessor verification of the A2-owned macro-classification tables, indexes, and physical field contracts. B5 owns their canonical relationship.
    On Error GoTo ErrHandler
    Dim issues As New Collection
    AddMacroClassificationTableSchemaIssues issues
    If issues.count > 0 Then
        detailText = SharedCollection_Join(issues, vbCrLf)
        Exit Function
    End If
    detailText = "A2-owned macro-classification tables, indexes, and physical field contracts are current; B5 owns canonical relationship creation and repair."
    MacroClassificationStateCurrentForB1 = True
    Exit Function
ErrHandler:
    detailText = "Could not verify A2-owned macro-classification state for B1. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
End Function

Private Sub RepairMacroClassificationTables(Optional ByVal showMessage As Boolean = True, Optional ByVal designChangeAlreadyPrepared As Boolean = False, Optional ByVal applyLookupMetadata As Boolean = True, Optional ByVal suppliedCategorySchemaDefinitions As Variant, Optional ByVal suppliedSubcategorySchemaDefinitions As Variant)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim createdCount As Long
    Dim previousDesignBatchPrepared As Boolean
    Dim familyNames As Collection
    Dim tableName As Variant
    Dim categoryTableName As String
    Dim subcategoryTableName As String
    Dim categorySchemaDefinitions As Collection
    Dim subcategorySchemaDefinitions As Collection
    If isMissing(suppliedCategorySchemaDefinitions) = False Then
        If IsObject(suppliedCategorySchemaDefinitions) Then Set categorySchemaDefinitions = suppliedCategorySchemaDefinitions
    End If
    If isMissing(suppliedSubcategorySchemaDefinitions) = False Then
        If IsObject(suppliedSubcategorySchemaDefinitions) Then Set subcategorySchemaDefinitions = suppliedSubcategorySchemaDefinitions
    End If
    If categorySchemaDefinitions Is Nothing Then Set categorySchemaDefinitions = BuildMacroCategorySchemaDefinitions()
    If subcategorySchemaDefinitions Is Nothing Then Set subcategorySchemaDefinitions = BuildMacroSubcategorySchemaDefinitions()
    previousDesignBatchPrepared = mTableDesignBatchPrepared
    mTableDesignBatchPrepared = previousDesignBatchPrepared Or designChangeAlreadyPrepared
    Set db = TablesWorkspaceCodeDb()
    Set familyNames = ManagedObjectNamesByB1Family(ManagedObjectB1MacroClassificationFamilyCode())
    For Each tableName In familyNames
        Select Case CStr(tableName)
            Case ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_MacroCategories")
                categoryTableName = CStr(tableName)
                RepairMacroClassificationTableFromSchema db, categoryTableName, categorySchemaDefinitions, createdCount, "category"
            Case ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_MacroSubcategories")
                subcategoryTableName = CStr(tableName)
                RepairMacroClassificationTableFromSchema db, subcategoryTableName, subcategorySchemaDefinitions, createdCount, "subcategory"
            Case Else
                Err.Raise vbObjectError + 9137, "RepairMacroClassificationTables", "The canonical A2-owned macro-classification family used by the B1 proof contains a table with no repair implementation mapping." & vbCrLf & "Table: " & CStr(tableName)
        End Select
    Next tableName
    If Len(categoryTableName) = 0 Or Len(subcategoryTableName) = 0 Then Err.Raise vbObjectError + 9138, "RepairMacroClassificationTables", "The canonical A2-owned macro-classification family used by the B1 proof did not provide both required implementation roles." & vbCrLf & "Category table: " & categoryTableName & vbCrLf & "Subcategory table: " & subcategoryTableName
    If applyLookupMetadata Then EnsureDefinedFieldLookup db.TableDefs(subcategoryTableName), FIELD_MACRO_CATEGORY_KEY
    If createdCount > 0 Then CoreAddOperationNote "Deferred Navigation Pane grouping for generated-macro category tables until the final catalog-driven placement pass."
    mTableDesignBatchPrepared = previousDesignBatchPrepared
    ReportCoreActionSuccess "RepairMacroClassificationTables", "TableRepair", "Generated-macro category and subcategory tables repaired successfully.", "Generated-macro classification structure is ready." & vbCrLf & vbCrLf & "Category table: " & categoryTableName & vbCrLf & "Subcategory table: " & subcategoryTableName, "Repair generated-macro classification tables", categoryTableName, "TableSet", "CodeDb", "TableRepair", showMessage, "Generated Macro Classification", vbInformation, True, "Tables", False
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    mTableDesignBatchPrepared = previousDesignBatchPrepared
    If showMessage Then
        CoreShowErrorMessage "Could not create or repair generated-macro category/subcategory storage." & vbCrLf & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, vbCritical, "Macro Classification Tables", "RepairMacroClassificationTables", "Tables"
        Exit Sub
    End If
    Err.Raise savedErrNumber, "RepairMacroClassificationTables", "Could not create or repair generated-macro category/subcategory storage." & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Sub

Private Sub RepairMacroClassificationTableFromSchema(ByVal db As DAO.Database, ByVal tableName As String, ByVal schemaDefinitions As Collection, ByRef createdCount As Long, ByVal classificationLabel As String)
    Dim tdf As DAO.tableDef
    Dim fld As DAO.Field
    Dim schemaDefinition As Variant
    If CodeDbTableExists(tableName) = False Then
        Set tdf = db.CreateTableDef(tableName)
        For Each schemaDefinition In schemaDefinitions
            If CStr(schemaDefinition(0)) = "FIELD" And CBool(schemaDefinition(4)) Then
                Set fld = tdf.CreateField(CStr(schemaDefinition(1)), CLng(schemaDefinition(2)))
                fld.Attributes = fld.Attributes Or dbAutoIncrField
                tdf.fields.Append fld
            End If
        Next schemaDefinition
        db.TableDefs.Append tdf
        createdCount = createdCount + 1
        CoreAddOperationNote "Created generated-macro " & classificationLabel & " table [" & tableName & "]."
    End If
    RemoveUnexpectedMacroClassificationFields db, tableName, schemaDefinitions
    For Each schemaDefinition In schemaDefinitions
        Select Case CStr(schemaDefinition(0))
            Case "FIELD"
                If CBool(schemaDefinition(4)) = False Then
                    If CBool(schemaDefinition(5)) Then
                        RepairTableField tableName, CStr(schemaDefinition(1)), CLng(schemaDefinition(2)), CLng(schemaDefinition(3)), schemaDefinition(6)
                    Else
                        RepairTableField tableName, CStr(schemaDefinition(1)), CLng(schemaDefinition(2)), CLng(schemaDefinition(3))
                    End If
                    ApplyMacroClassificationPhysicalFieldContract db, tableName, schemaDefinition
                End If
            Case "INDEX"
                EnsureSingleFieldIndex db, tableName, CStr(schemaDefinition(1)), CStr(schemaDefinition(2)), CBool(schemaDefinition(3)), CBool(schemaDefinition(4)), mTableDesignBatchPrepared
            Case "MULTI_INDEX"
                EnsureMultiFieldIndex db, tableName, CStr(schemaDefinition(1)), CStr(schemaDefinition(2)), CBool(schemaDefinition(3)), CBool(schemaDefinition(4))
        End Select
    Next schemaDefinition
End Sub

Private Sub RemoveUnexpectedMacroClassificationFields(ByVal db As DAO.Database, ByVal tableName As String, ByVal schemaDefinitions As Collection)
    Dim tdf As DAO.tableDef
    Dim fld As DAO.Field
    Dim idx As DAO.index
    Dim idxField As DAO.Field
    ' INTENTIONAL FILTERED PARTITION: fieldNamesToDelete is a consumer-required subset/projection derived from the source enumerated by RemoveUnexpectedMacroClassificationFields; its membership is not independently
    ' authoritative.
    Dim fieldNamesToDelete As New Collection
    Dim indexNamesToDelete As Collection
    Dim fieldName As Variant
    Dim indexName As Variant
    Dim indexUsesField As Boolean
    Set tdf = db.TableDefs(tableName)
    For Each fld In tdf.fields
        If MacroClassificationSchemaDefinesField(schemaDefinitions, fld.Name) = False Then fieldNamesToDelete.Add fld.Name
    Next fld
    For Each fieldName In fieldNamesToDelete
        ' INTENTIONAL FILTERED PARTITION: indexNamesToDelete is a consumer-required subset/projection derived from the source enumerated by RemoveUnexpectedMacroClassificationFields; its membership is not
        ' independently authoritative.
        Set indexNamesToDelete = New Collection
        For Each idx In tdf.Indexes
            indexUsesField = False
            For Each idxField In idx.fields
                If StrComp(idxField.Name, CStr(fieldName), vbTextCompare) = 0 Then
                    indexUsesField = True
                    Exit For
                End If
            Next idxField
            If indexUsesField Then indexNamesToDelete.Add idx.Name
        Next idx
        For Each indexName In indexNamesToDelete
            tdf.Indexes.Delete CStr(indexName)
        Next indexName
        tdf.fields.Delete CStr(fieldName)
        CoreAddOperationNote "Removed unexpected obsolete field [" & CStr(fieldName) & "] from canonical ADC-owned table [" & tableName & "]."
    Next fieldName
    If fieldNamesToDelete.count > 0 Then
        tdf.Indexes.Refresh
        tdf.fields.Refresh
    End If
End Sub

Private Sub AddMultiFieldIndexDefinitionIssue(ByRef issues As Collection, ByVal tableName As String, ByVal indexName As String, ByVal fieldNamesText As String, ByVal isUnique As Boolean, ByVal isPrimary As Boolean)
    On Error GoTo MissingOrWrong
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim idx As DAO.index
    Set db = TablesWorkspaceCodeDb()
    Set tdf = db.TableDefs(tableName)
    Set idx = tdf.Indexes(indexName)
    If IndexMatchesFieldList(idx, fieldNamesText, isUnique, isPrimary) Then Exit Sub
MissingOrWrong:
    issues.Add tableName & "." & indexName & ": composite index is missing or does not match fields/order/uniqueness/primary-key requirements [" & Replace(fieldNamesText, "|", ", ") & "]"
End Sub

Private Function TableDefIndexExists(ByVal tdf As DAO.tableDef, ByVal indexName As String) As Boolean
    On Error GoTo NotFound
    Dim idx As DAO.index
    Set idx = tdf.Indexes(indexName)
    TableDefIndexExists = True
    Exit Function
NotFound:
End Function

Private Function IndexMatchesFieldList(ByVal idx As DAO.index, ByVal fieldNamesText As String, ByVal isUnique As Boolean, ByVal isPrimary As Boolean) As Boolean
    On Error GoTo ExitHere
    Dim fieldNames As Variant
    Dim i As Long
    fieldNames = Split(fieldNamesText, "|")
    If idx.fields.count <> UBound(fieldNames) - LBound(fieldNames) + 1 Then Exit Function
    If CBool(idx.Unique) <> isUnique Then Exit Function
    If CBool(idx.Primary) <> isPrimary Then Exit Function
    For i = LBound(fieldNames) To UBound(fieldNames)
        If StrComp(idx.fields(i - LBound(fieldNames)).Name, CStr(fieldNames(i)), vbTextCompare) <> 0 Then Exit Function
    Next i
    IndexMatchesFieldList = True
ExitHere:
End Function

Private Sub EnsureMultiFieldIndex(ByVal db As DAO.Database, ByVal tableName As String, ByVal indexName As String, ByVal fieldNamesText As String, ByVal isUnique As Boolean, ByVal isPrimary As Boolean)
    On Error GoTo ErrHandler
    Dim tdf As DAO.tableDef
    Dim idx As DAO.index
    Dim idxField As DAO.Field
    Dim fieldNames As Variant
    Dim fieldName As Variant
    If db Is Nothing Then Err.Raise vbObjectError + 9142, "EnsureMultiFieldIndex", "Database reference is Nothing."
    If mTableDesignBatchPrepared = False Then db.TableDefs.Refresh
    Set tdf = db.TableDefs(tableName)
    If mTableDesignBatchPrepared = False Then tdf.Indexes.Refresh
    If TableDefIndexExists(tdf, indexName) Then
        If IndexMatchesFieldList(tdf.Indexes(indexName), fieldNamesText, isUnique, isPrimary) Then Exit Sub
    End If
    If mTableDesignBatchPrepared = False Then
        If SharedAccess_PrepareOpenObjectForDesignChange(acTable, tableName, "creating or repairing composite index [" & indexName & "] on table [" & tableName & "]") = False Then
            MarkSystemCommandCancelled
            Err.Raise vbObjectError + 9143, "EnsureMultiFieldIndex", "Composite-index repair was cancelled because the affected open table was not approved for save and close."
        End If
        Set tdf = Nothing
        Set db = TablesWorkspaceCodeDb()
        db.TableDefs.Refresh
        Set tdf = db.TableDefs(tableName)
        tdf.Indexes.Refresh
    End If
    If TableDefIndexExists(tdf, indexName) Then
        tdf.Indexes.Delete indexName
        If mTableDesignBatchPrepared = False Then tdf.Indexes.Refresh
    End If
    Set idx = tdf.CreateIndex(indexName)
    idx.Unique = isUnique
    idx.Primary = isPrimary
    fieldNames = Split(fieldNamesText, "|")
    For Each fieldName In fieldNames
        If SharedDao_TableDefFieldExistsSafe(tdf, CStr(fieldName)) = False Then Err.Raise vbObjectError + 9144, "EnsureMultiFieldIndex", "Composite-index field is missing." & vbCrLf & "Field: " & CStr(fieldName)
        Set idxField = idx.CreateField(CStr(fieldName))
        idx.fields.Append idxField
    Next fieldName
    tdf.Indexes.Append idx
    If mTableDesignBatchPrepared = False Then tdf.Indexes.Refresh
    If TableDefIndexExists(tdf, indexName) = False Then
        tdf.Indexes.Refresh
        If TableDefIndexExists(tdf, indexName) = False Then Err.Raise vbObjectError + 9140, "EnsureMultiFieldIndex", "Composite index was not present after creation."
    End If
    If IndexMatchesFieldList(tdf.Indexes(indexName), fieldNamesText, isUnique, isPrimary) = False Then Err.Raise vbObjectError + 9141, "EnsureMultiFieldIndex", "Composite index did not match its canonical definition after creation."
    CoreAddOperationNote "Created/repaired composite index [" & indexName & "] on table [" & tableName & "]."
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "EnsureMultiFieldIndex", "Could not create or verify canonical composite index." & vbCrLf & "Table: " & tableName & vbCrLf & "Index: " & indexName & vbCrLf & "Fields: " & Replace(fieldNamesText, "|", ", ") & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Function TableSchemaRepairRequired(ByRef repairSummary As String, Optional ByVal suppliedSettingsSchemaDefinitions As Variant, Optional ByVal suppliedSettingsIndexDefinitions As Variant, Optional ByVal suppliedConfigSchemaDefinitions As Variant, Optional ByVal suppliedConfigIndexDefinitions As Variant, Optional ByVal suppliedMacroCategorySchemaDefinitions As Variant, Optional ByVal suppliedMacroSubcategorySchemaDefinitions As Variant, Optional ByVal suppliedEditableLookupValueTextSizes As Variant) As Boolean
    On Error GoTo ErrHandler
    Dim issues As Collection
    Dim settingsSchemaDefinitions As Collection
    Dim settingsIndexDefinitions As Collection
    Dim configSchemaDefinitions As Collection
    Dim configIndexDefinitions As Collection
    Dim macroCategorySchemaDefinitions As Collection
    Dim macroSubcategorySchemaDefinitions As Collection
    Dim editableLookupValueTextSizes As Object
    If isMissing(suppliedSettingsSchemaDefinitions) = False Then
        If IsObject(suppliedSettingsSchemaDefinitions) Then Set settingsSchemaDefinitions = suppliedSettingsSchemaDefinitions
    End If
    If isMissing(suppliedSettingsIndexDefinitions) = False Then
        If IsObject(suppliedSettingsIndexDefinitions) Then Set settingsIndexDefinitions = suppliedSettingsIndexDefinitions
    End If
    If isMissing(suppliedConfigSchemaDefinitions) = False Then
        If IsObject(suppliedConfigSchemaDefinitions) Then Set configSchemaDefinitions = suppliedConfigSchemaDefinitions
    End If
    If isMissing(suppliedConfigIndexDefinitions) = False Then
        If IsObject(suppliedConfigIndexDefinitions) Then Set configIndexDefinitions = suppliedConfigIndexDefinitions
    End If
    If isMissing(suppliedMacroCategorySchemaDefinitions) = False Then
        If IsObject(suppliedMacroCategorySchemaDefinitions) Then Set macroCategorySchemaDefinitions = suppliedMacroCategorySchemaDefinitions
    End If
    If isMissing(suppliedMacroSubcategorySchemaDefinitions) = False Then
        If IsObject(suppliedMacroSubcategorySchemaDefinitions) Then Set macroSubcategorySchemaDefinitions = suppliedMacroSubcategorySchemaDefinitions
    End If
    If isMissing(suppliedEditableLookupValueTextSizes) = False Then
        If IsObject(suppliedEditableLookupValueTextSizes) Then Set editableLookupValueTextSizes = suppliedEditableLookupValueTextSizes
    End If
    If settingsSchemaDefinitions Is Nothing Then Set settingsSchemaDefinitions = SettingsFieldSchemaDefinitions()
    If settingsIndexDefinitions Is Nothing Then Set settingsIndexDefinitions = SettingsIndexSchemaDefinitions()
    If configSchemaDefinitions Is Nothing Then Set configSchemaDefinitions = ConfigFieldSchemaDefinitions()
    If configIndexDefinitions Is Nothing Then Set configIndexDefinitions = ConfigIndexSchemaDefinitions()
    If macroCategorySchemaDefinitions Is Nothing Then Set macroCategorySchemaDefinitions = BuildMacroCategorySchemaDefinitions()
    If macroSubcategorySchemaDefinitions Is Nothing Then Set macroSubcategorySchemaDefinitions = BuildMacroSubcategorySchemaDefinitions()
    If editableLookupValueTextSizes Is Nothing Then Set editableLookupValueTextSizes = EditableLookupValueTextSizeIndex(configSchemaDefinitions, BuildHealthReviewSchemaDefinitions(), CustomMacroLauncherReviewSchemaDefinitions())
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: issues holds planned/recovery work or independently observed validation evidence for TableSchemaRepairRequired; keep it distinct from the source authority.
    Set issues = New Collection
    AddSettingsTableSchemaIssues issues, settingsSchemaDefinitions, settingsIndexDefinitions
    AddConfigTableSchemaIssues issues, configSchemaDefinitions, configIndexDefinitions
    AddCandidateTableSchemaIssues issues, CandidateFieldSchemaDefinitions(), CandidateIndexSchemaDefinitions()
    AddHealthReviewTableSchemaIssues issues, BuildHealthReviewSchemaDefinitions()
    AddLauncherReviewTableSchemaIssues issues, CustomMacroLauncherReviewSchemaDefinitions()
    AddMacroClassificationTableSchemaIssues issues, macroCategorySchemaDefinitions, macroSubcategorySchemaDefinitions
    AddEditableLookupTableSchemaIssues issues, editableLookupValueTextSizes
    AddLookupDefinitionRegistrySchemaIssues issues
    If issues.count > 0 Then
        repairSummary = SharedCollection_Join(issues, vbCrLf)
        TableSchemaRepairRequired = True
    Else
        repairSummary = vbNullString
        TableSchemaRepairRequired = False
    End If
    Exit Function
ErrHandler:
    repairSummary = "Schema validation failed before repair: " & Err.Description
    TableSchemaRepairRequired = True
End Function

Private Sub AddLookupDefinitionRegistrySchemaIssues(ByRef issues As Collection)
    Dim tableName As String
    Dim schemaDefinitions As New Collection
    tableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LookupDefinitions")
    If CodeDbTableExists(tableName) = False Then issues.Add tableName & ": table is missing": Exit Sub
    If TableCanOpen(tableName, True) = False Then issues.Add tableName & ": table exists but cannot be opened": Exit Sub
    schemaDefinitions.Add Array(FIELD_ID, dbLong, 0, True)
    schemaDefinitions.Add Array(FIELD_LOOKUP_LOGICAL_ID, dbText, 100, False)
    schemaDefinitions.Add Array(FIELD_LOOKUP_DISPLAY_NAME, dbText, 150, False)
    schemaDefinitions.Add Array(FIELD_LOOKUP_TABLE_NAME, dbText, 150, False)
    schemaDefinitions.Add Array(FIELD_LOOKUP_EDITOR_FORM_NAME, dbText, 150, False)
    schemaDefinitions.Add Array(FIELD_EDITOR_MODE, dbText, 20, False)
    schemaDefinitions.Add Array(FIELD_ALLOW_ADDITIONS, dbBoolean, 0, False)
    schemaDefinitions.Add Array(FIELD_ALLOW_DELETIONS, dbBoolean, 0, False)
    schemaDefinitions.Add Array(FIELD_ALLOW_CODE_EDITS, dbBoolean, 0, False)
    schemaDefinitions.Add Array(FIELD_ALLOW_DISPLAY_NAME_EDITS, dbBoolean, 0, False)
    schemaDefinitions.Add Array(FIELD_ALLOW_DESCRIPTION_EDITS, dbBoolean, 0, False)
    schemaDefinitions.Add Array(FIELD_ALLOW_SORT_ORDER_EDITS, dbBoolean, 0, False)
    schemaDefinitions.Add Array(FIELD_ALLOW_ACTIVE_EDITS, dbBoolean, 0, False)
    schemaDefinitions.Add Array(FIELD_FINALIZED_ON, dbDate, 0, False)
    schemaDefinitions.Add Array(FIELD_FINALIZED_BY, dbText, 100, False)
    schemaDefinitions.Add Array(FIELD_MANAGED_OBJECT_DEPENDENCY_NOTES, dbMemo, 0, False)
    AddFieldDefinitionIssuesFromSchema issues, tableName, schemaDefinitions
    AddIndexDefinitionIssue issues, tableName, INDEX_PRIMARY_KEY, FIELD_ID, True, True
    AddIndexDefinitionIssue issues, tableName, INDEX_UX_LOOKUP_LOGICAL_ID, FIELD_LOOKUP_LOGICAL_ID, True, False
End Sub

' Purpose: Ensures terminology table for apply lookup metadata, supplied schema definitions.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   applyLookupMetadata (Input Optional; Boolean): When True, enables or requests lookup metadata for this call; when False, the routine follows its normal non-forced/default path. Default: True.
'   suppliedSchemaDefinitions (Input Optional; Variant): Caller-supplied supplied schema definitions data used as the authoritative definition input for this call, avoiding an unnecessary reload when already
'   available.
'   suppliedIndexDefinitions (Input Optional; Variant): Caller-supplied supplied index definitions data used as the authoritative definition input for this call, avoiding an unnecessary reload when already
'   available.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub EnsureTerminologyTable(Optional ByVal applyLookupMetadata As Boolean = True, Optional ByVal suppliedSchemaDefinitions As Variant, Optional ByVal suppliedIndexDefinitions As Variant)
    'TABLES is the sole physical creation/repair authority for the Help_Terms table; Catalog remains a caller/data consumer only.
    On Error GoTo ErrHandler
    Dim tableName As String
    Dim schemaDefinitions As Collection
    Dim indexDefinitions As Collection
    tableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_TERMINOLOGY_TABLE_LOGICAL_ID)
    If isMissing(suppliedSchemaDefinitions) = False Then
        If IsObject(suppliedSchemaDefinitions) Then Set schemaDefinitions = suppliedSchemaDefinitions
    End If
    If isMissing(suppliedIndexDefinitions) = False Then
        If IsObject(suppliedIndexDefinitions) Then Set indexDefinitions = suppliedIndexDefinitions
    End If
    If schemaDefinitions Is Nothing Then Set schemaDefinitions = TerminologyFieldSchemaDefinitions()
    If indexDefinitions Is Nothing Then Set indexDefinitions = TerminologyIndexSchemaDefinitions()
    If CodeDbTableExists(tableName) Then
        EnsureTerminologyFields tableName, schemaDefinitions
    Else
        CreateTerminologyTable tableName, schemaDefinitions
    End If
    EnsureTerminologyIndexes tableName, indexDefinitions
    If applyLookupMetadata Then EnsureTerminologyLookupFields tableName
    ReconcileStandaloneManagedTableDescriptionIfAllowed tableName
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "EnsureTerminologyTable", "Could not create or repair the " & DEFAULT_UTILITY_LONG_NAME & " Help_Terms table." & vbCrLf & "Table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_TERMINOLOGY_TABLE_LOGICAL_ID) & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

Private Sub CreateSettingsTable(ByVal tableName As String, ByVal applyLookupMetadata As Boolean, ByVal schemaDefinitions As Collection, ByVal indexDefinitions As Collection)
    On Error GoTo FirstAttemptFailed
    If SharedDao_QueryExists(TablesWorkspaceCodeDb(), tableName) Then
        SafeDropTableIfExists tableName, "Dropped saved query name conflict [" & tableName & "] before creating the macro-creator settings table."
    End If
    CreateSettingsTableAttempt tableName, applyLookupMetadata, schemaDefinitions, indexDefinitions
    Exit Sub
FirstAttemptFailed:
    Dim firstErrNumber As Long
    Dim firstErrSource As String
    Dim firstErrDescription As String
    firstErrNumber = Err.Number
    firstErrSource = Err.Source
    firstErrDescription = Err.Description
    If firstErrNumber = 3010 Or firstErrNumber = 3078 Or firstErrNumber = 3265 Or firstErrNumber = 7874 Then
        On Error GoTo RetryFailed
        SafeDropTableIfExists tableName, "Removed conflicting/stale settings table object [" & tableName & "] before retrying table creation. " & "First create attempt failed with Err.Number " & CStr(firstErrNumber) & ": " & SharedMessage_OneLine(firstErrDescription, 250)
        CreateSettingsTableAttempt tableName, applyLookupMetadata, schemaDefinitions, indexDefinitions
        Exit Sub
    End If
    Err.Raise firstErrNumber, "CreateSettingsTable", "Could not create access development console settings table." & vbCrLf & "Table: " & tableName & vbCrLf & "Err.Number: " & CStr(firstErrNumber) & vbCrLf & "Err.Source: " & firstErrSource & vbCrLf & "Err.Description: " & firstErrDescription
RetryFailed:
    Dim retryErrNumber As Long
    Dim retryErrSource As String
    Dim retryErrDescription As String
    retryErrNumber = Err.Number
    retryErrSource = Err.Source
    retryErrDescription = Err.Description
    Err.Raise retryErrNumber, "CreateSettingsTable", "Could not create access development console settings table after dropping a conflicting/stale table object." & vbCrLf & "Table: " & tableName & vbCrLf & "First Err.Number: " & CStr(firstErrNumber) & vbCrLf & "First Err.Source: " & firstErrSource & vbCrLf & "First Err.Description: " & firstErrDescription & vbCrLf & "Retry Err.Number: " & CStr(retryErrNumber) & vbCrLf & "Retry Err.Source: " & retryErrSource & vbCrLf & "Retry Err.Description: " & retryErrDescription
End Sub

Private Sub CreateSettingsTableAttempt(ByVal tableName As String, ByVal applyLookupMetadata As Boolean, ByVal schemaDefinitions As Collection, ByVal indexDefinitions As Collection)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim fieldDefinition As Variant
    Set db = TablesWorkspaceCodeDb()
    Set tdf = db.CreateTableDef(tableName)
    For Each fieldDefinition In schemaDefinitions
        AppendSchemaField tdf, CStr(fieldDefinition(0)), CLng(fieldDefinition(1)), CLng(fieldDefinition(2)), CBool(fieldDefinition(3)), "settings", fieldDefinition
    Next fieldDefinition
    db.TableDefs.Append tdf
    TrackSetupCreatedTableRollback tableName
    db.TableDefs.Refresh
    EnsureTableDefFieldOrderFromText db.TableDefs(tableName), SchemaDefinitionFieldOrderText(schemaDefinitions), mTableDesignBatchPrepared
    If applyLookupMetadata Then EnsureSettingsLookupFields tableName
    CoreAddOperationNote "Created settings table [" & tableName & "]."
    CoreAddOperationNote "Deferred Navigation Pane grouping for newly created settings table [" & tableName & "] until the final catalog-driven setup pass."
    EnsureSettingsIndexes tableName, indexDefinitions
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrSource As String
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrSource = Err.Source
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "CreateSettingsTableAttempt", "Could not create access development console settings table." & vbCrLf & "Table: " & tableName & vbCrLf & "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB & vbCrLf & "Database: " & TablesWorkspaceCodeDb().Name & vbCrLf & "Err.Number: " & CStr(originalErrNumber) & vbCrLf & "Err.Source: " & originalErrSource & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

Private Sub EnsureConfigTable(Optional ByVal applyLookupMetadata As Boolean = True, Optional ByVal suppliedSchemaDefinitions As Variant, Optional ByVal suppliedIndexDefinitions As Variant)
    On Error GoTo ErrHandler
    Dim tableName As String
    Dim remainingSchemaIssues As String
    Dim tableCreatedThisCall As Boolean
    Dim schemaDefinitions As Collection
    Dim indexDefinitions As Collection
    tableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")
    If isMissing(suppliedSchemaDefinitions) = False Then
        If IsObject(suppliedSchemaDefinitions) Then Set schemaDefinitions = suppliedSchemaDefinitions
    End If
    If isMissing(suppliedIndexDefinitions) = False Then
        If IsObject(suppliedIndexDefinitions) Then Set indexDefinitions = suppliedIndexDefinitions
    End If
    If schemaDefinitions Is Nothing Then Set schemaDefinitions = ConfigFieldSchemaDefinitions()
    If indexDefinitions Is Nothing Then Set indexDefinitions = ConfigIndexSchemaDefinitions()
    PrepareTableNameForCreateOrRepair tableName, tableName
    If CodeDbTableExists(tableName) = False Then
        CreateConfigTable tableName, applyLookupMetadata, schemaDefinitions, indexDefinitions
        tableCreatedThisCall = True
    Else
        RequirePreparedTableCanOpen tableName, tableName
        EnsureConfigCurrentDesignFields tableName, schemaDefinitions
    End If
    If tableCreatedThisCall Then
        CoreAddOperationNote "MC_Records table creation completed structural field order and indexes; B2 owns lookup presentation/editor intent, D4 owns editor binding, and G1 owns managed DAO descriptions during Setup."
    Else
        RepairConfigIndexes tableName, False, indexDefinitions
    End If
    remainingSchemaIssues = ConfigTableSchemaIssueText(schemaDefinitions, indexDefinitions)
    If Len(remainingSchemaIssues) > 0 Then
        Err.Raise vbObjectError + 5331, "EnsureConfigTable", DEFAULT_UTILITY_LONG_NAME & " MC_Records table still does not match the current managed-table design schema after additive field and index repair." & vbCrLf & "MC_Records table: " & tableName & vbCrLf & "Remaining schema issues:" & vbCrLf & remainingSchemaIssues
    End If
    AddTraceNote "MC_Records schema repair preserved all existing row values; normalization remains an explicit manual action."
    If tableCreatedThisCall = False Then
        If applyLookupMetadata Then RepairConfigLookupFields tableName, False
    End If
    VerifyTableReady tableName, tableName
    Exit Sub
ErrHandler:
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim openFailureDetail As String
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    openFailureDetail = TableOpenFailureDetail(tableName)
    Err.Raise errNumber, "EnsureConfigTable", "Could not create or repair the macro-creator MC_Records table." & vbCrLf & "MC_Records table: " & tableName & vbCrLf & "Err.Number: " & CStr(errNumber) & vbCrLf & "Err.Source: " & errSource & vbCrLf & "Err.Description: " & errDescription & vbCrLf & "Open failure detail: " & openFailureDetail & TableReadinessDiagnosticsForMessage()
End Sub


Private Sub EnsureCandidateTable(Optional ByVal applyLookupMetadata As Boolean = True, Optional ByVal suppliedSchemaDefinitions As Variant, Optional ByVal suppliedIndexDefinitions As Variant)
    On Error GoTo ErrHandler
    Dim tableName As String
    Dim remainingSchemaIssues As String
    Dim tableCreatedThisCall As Boolean
    Dim schemaDefinitions As Collection
    Dim indexDefinitions As Collection
    Dim duplicateSignatureRows As DAO.Recordset
    tableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates")
    If isMissing(suppliedSchemaDefinitions) = False Then
        If IsObject(suppliedSchemaDefinitions) Then Set schemaDefinitions = suppliedSchemaDefinitions
    End If
    If isMissing(suppliedIndexDefinitions) = False Then
        If IsObject(suppliedIndexDefinitions) Then Set indexDefinitions = suppliedIndexDefinitions
    End If
    If schemaDefinitions Is Nothing Then Set schemaDefinitions = CandidateFieldSchemaDefinitions()
    If indexDefinitions Is Nothing Then Set indexDefinitions = CandidateIndexSchemaDefinitions()
    PrepareTableNameForCreateOrRepair tableName, tableName
    If CodeDbTableExists(tableName) = False Then
        CreateConfigTable tableName, applyLookupMetadata, schemaDefinitions, indexDefinitions
        tableCreatedThisCall = True
    Else
        RequirePreparedTableCanOpen tableName, tableName
        EnsureConfigCurrentDesignFields tableName, schemaDefinitions
    End If
    If tableCreatedThisCall Then
        CoreAddOperationNote "MC_Candidates table creation completed structural field order and indexes; B2 owns lookup presentation/editor intent, D4 owns editor binding, and G1 owns managed DAO descriptions during Setup."
    Else
        Set duplicateSignatureRows = TablesWorkspaceCodeDb().OpenRecordset("SELECT TOP 1 [" & FIELD_SCAN_SIGNATURE & "], Count(*) AS " & FIELD_DUP_COUNT & " FROM [" & tableName & "] WHERE [" & FIELD_SCAN_SIGNATURE & "] Is Not Null GROUP BY [" & FIELD_SCAN_SIGNATURE & "] HAVING Count(*)>1 ORDER BY [" & FIELD_SCAN_SIGNATURE & "];", dbOpenSnapshot)
        If duplicateSignatureRows.EOF = False Then
            Err.Raise vbObjectError + 5337, "EnsureCandidateTable", "MC_Candidates ScanSignature uniqueness repair was blocked because duplicate persisted signatures already exist. Existing candidate evidence was not changed." & vbCrLf & "ScanSignature: " & CStr(Nz(duplicateSignatureRows.fields(FIELD_SCAN_SIGNATURE).value, vbNullString)) & vbCrLf & "Rows: " & CStr(CLng(duplicateSignatureRows.fields(FIELD_DUP_COUNT)))
        End If
        duplicateSignatureRows.Close
        Set duplicateSignatureRows = Nothing
        RepairConfigIndexes tableName, False, indexDefinitions
    End If
    remainingSchemaIssues = CandidateTableSchemaIssueText(schemaDefinitions, indexDefinitions)
    If Len(remainingSchemaIssues) > 0 Then Err.Raise vbObjectError + 5332, "EnsureCandidateTable", DEFAULT_UTILITY_LONG_NAME & " MC_Candidates table still does not match the current managed-table design schema after additive field and index repair." & vbCrLf & "MC_Candidates table: " & tableName & vbCrLf & "Remaining schema issues:" & vbCrLf & remainingSchemaIssues
    If tableCreatedThisCall = False Then
        If applyLookupMetadata Then RepairConfigLookupFields tableName, False
    End If
    VerifyTableReady tableName, tableName
    Exit Sub
ErrHandler:
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    On Error Resume Next
    If Not duplicateSignatureRows Is Nothing Then duplicateSignatureRows.Close
    On Error GoTo 0
    Err.Raise errNumber, "EnsureCandidateTable", "Could not create or repair MC_Candidates." & vbCrLf & "MC_Candidates table: " & tableName & vbCrLf & "Err.Number: " & CStr(errNumber) & vbCrLf & "Err.Source: " & errSource & vbCrLf & "Err.Description: " & errDescription
End Sub

Private Sub MigrateLegacyConfigCandidateRowsForSetup(ByVal configTableName As String, ByVal candidateTableName As String)
    'A2 migration only: move legacy candidate/discovery rows out of the former mixed MC_Records table and preserve accepted-candidate evidence beside the configuration it produced.
    'The copy/delete work is one DAO workspace transaction so an interrupted migration cannot leave duplicated unaccepted candidate rows.
    On Error GoTo ErrHandler
    Dim ws As DAO.workspace
    Dim db As DAO.Database
    Dim configTdf As DAO.tableDef
    Dim rsLegacy As DAO.Recordset
    Dim rsCandidates As DAO.Recordset
    Dim fieldDefinition As Variant
    Dim fieldName As String
    Dim rowKindText As String
    Dim configID As Long
    Dim AcceptedConfigurationID As Long
    Dim migrateRow As Boolean
    Dim deleteLegacyRow As Boolean
    Dim copiedCount As Long
    Dim deletedCount As Long
    Dim acceptedEvidenceCount As Long
    Dim inTransaction As Boolean
    Set db = TablesWorkspaceCodeDb()
    If SharedDao_DatabaseTableExistsSafe(db, configTableName) = False Then Exit Sub
    Set configTdf = db.TableDefs(configTableName)
    If SharedDao_TableDefFieldExistsSafe(configTdf, FIELD_ROW_KIND) = False Then Exit Sub
    If SharedDao_DatabaseTableExistsSafe(db, candidateTableName) = False Then Err.Raise vbObjectError + 5333, "MigrateLegacyConfigCandidateRowsForSetup", "MC_Candidates must exist before legacy candidate migration."
    Set ws = DBEngine.Workspaces(0)
    Set rsLegacy = db.OpenRecordset(configTableName, dbOpenDynaset)
    Set rsCandidates = db.OpenRecordset(candidateTableName, dbOpenDynaset)
    ws.BeginTrans
    inTransaction = True
    Do While rsLegacy.EOF = False
        rowKindText = SharedText_Nz(FieldValueIfExists(rsLegacy, FIELD_ROW_KIND))
        configID = SharedValue_NzLong(FieldValueIfExists(rsLegacy, FIELD_ID))
        migrateRow = LegacyConfigRowHasCandidateEvidence(rsLegacy, rowKindText)
        deleteLegacyRow = (Len(rowKindText) > 0 And StrComp(rowKindText, ROW_KIND_CONFIGURATION, vbTextCompare) <> 0)
        If migrateRow Then
            AcceptedConfigurationID = SharedValue_NzLong(FieldValueIfExists(rsLegacy, FIELD_ACCEPTED_CONFIGURATION_ID))
            If StrComp(rowKindText, ROW_KIND_CONFIGURATION, vbTextCompare) = 0 Then
                'Legacy acceptance promoted the candidate in place, so candidate evidence carried by a configuration row belongs to this exact configuration identity.
                'Do not preserve a stale/corrupt self-link to another configuration during the split migration.
                AcceptedConfigurationID = configID
                If CandidateEvidenceLinkedToConfigurationExists(db, candidateTableName, AcceptedConfigurationID) = False Then
                    rsCandidates.AddNew
                    For Each fieldDefinition In CandidateFieldSchemaDefinitions()
                        fieldName = CStr(fieldDefinition(0))
                        If StrComp(fieldName, FIELD_ID, vbTextCompare) <> 0 And StrComp(fieldName, FIELD_ACCEPTED_CONFIGURATION_ID, vbTextCompare) <> 0 Then SetRecordsetFieldIfExists rsCandidates, fieldName, FieldValueIfExists(rsLegacy, fieldName)
                    Next fieldDefinition
                    SetRecordsetFieldIfExists rsCandidates, FIELD_ACCEPTED_CONFIGURATION_ID, AcceptedConfigurationID
                    rsCandidates.Update
                    copiedCount = copiedCount + 1
                    acceptedEvidenceCount = acceptedEvidenceCount + 1
                End If
            Else
                rsCandidates.AddNew
                For Each fieldDefinition In CandidateFieldSchemaDefinitions()
                    fieldName = CStr(fieldDefinition(0))
                    If StrComp(fieldName, FIELD_ID, vbTextCompare) <> 0 Then SetRecordsetFieldIfExists rsCandidates, fieldName, FieldValueIfExists(rsLegacy, fieldName)
                Next fieldDefinition
                rsCandidates.Update
                copiedCount = copiedCount + 1
            End If
        End If
        If deleteLegacyRow Then
            rsLegacy.Delete
            deletedCount = deletedCount + 1
        End If
        rsLegacy.MoveNext
    Loop
    ws.CommitTrans
    inTransaction = False
    If copiedCount > 0 Or deletedCount > 0 Then CoreAddOperationNote "A2 migrated legacy MC_Records candidate evidence to MC_Candidates. Candidate rows copied: " & CStr(copiedCount) & "; legacy non-configuration rows removed from MC_Records: " & CStr(deletedCount) & "; accepted configuration evidence rows linked: " & CStr(acceptedEvidenceCount) & "."
ExitHere:
    On Error Resume Next
    If Not rsLegacy Is Nothing Then rsLegacy.Close
    If Not rsCandidates Is Nothing Then rsCandidates.Close
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If inTransaction Then ws.Rollback
    If Not rsLegacy Is Nothing Then rsLegacy.Close
    If Not rsCandidates Is Nothing Then rsCandidates.Close
    On Error GoTo 0
    Err.Raise savedErrNumber, "MigrateLegacyConfigCandidateRowsForSetup", "Could not migrate legacy candidate/discovery state out of MC_Records." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Function LegacyConfigRowHasCandidateEvidence(ByVal rsLegacy As DAO.Recordset, ByVal rowKindText As String) As Boolean
    If Len(rowKindText) > 0 And StrComp(rowKindText, ROW_KIND_CONFIGURATION, vbTextCompare) <> 0 Then LegacyConfigRowHasCandidateEvidence = True: Exit Function
    If Len(SharedText_Nz(FieldValueIfExists(rsLegacy, FIELD_CANDIDATE_STATUS))) > 0 Then LegacyConfigRowHasCandidateEvidence = True: Exit Function
    If Len(SharedText_Nz(FieldValueIfExists(rsLegacy, FIELD_CANDIDATE_MACRO_KEY))) > 0 Then LegacyConfigRowHasCandidateEvidence = True: Exit Function
    If Len(SharedText_Nz(FieldValueIfExists(rsLegacy, FIELD_SCAN_SIGNATURE))) > 0 Then LegacyConfigRowHasCandidateEvidence = True: Exit Function
    If Len(SharedText_Nz(FieldValueIfExists(rsLegacy, FIELD_CANDIDATE_OWNER_DATABASE_ROLE))) > 0 Then LegacyConfigRowHasCandidateEvidence = True: Exit Function
    If SharedValue_NzLong(FieldValueIfExists(rsLegacy, FIELD_ACCEPTED_CONFIGURATION_ID)) > 0 Then LegacyConfigRowHasCandidateEvidence = True
End Function

Private Function CandidateEvidenceLinkedToConfigurationExists(ByVal db As DAO.Database, ByVal candidateTableName As String, ByVal configurationID As Long) As Boolean
    Dim rs As DAO.Recordset
    If configurationID <= 0 Then Exit Function
    Set rs = db.OpenRecordset("SELECT TOP 1 [" & FIELD_ID & "] FROM [" & candidateTableName & "] WHERE [" & FIELD_ACCEPTED_CONFIGURATION_ID & "]=" & CStr(configurationID), dbOpenSnapshot)
    CandidateEvidenceLinkedToConfigurationExists = (rs.EOF = False)
    rs.Close
End Function

Private Function LegacyConfigCandidateFieldNames() As Variant
    'Single A2-owned roster shared by legacy-field retirement and the fast-path pending-retirement check so the two paths cannot drift apart.
    LegacyConfigCandidateFieldNames = Array(FIELD_ROW_KIND, FIELD_CANDIDATE_STATUS, FIELD_CANDIDATE_MACRO_KEY, FIELD_CANDIDATE_OWNER_DATABASE_ROLE, FIELD_CANDIDATE_SCAN_TARGET_DB_NAME, FIELD_CANDIDATE_SCAN_TARGET_DB_PATH, FIELD_CANDIDATE_VB_PROJECT_NAME, FIELD_CANDIDATE_MODULE_NAME, FIELD_SCAN_SIGNATURE, FIELD_LAST_SCAN_BATCH_ID, FIELD_LAST_SCANNED_ON, FIELD_DISCOVERED_ON, FIELD_MODULE_HEADER_FOUND, FIELD_MODULE_HEADER_MALFORMED, FIELD_MODULE_HEADER_TITLE, FIELD_MODULE_HEADER_PURPOSE, FIELD_MODULE_HEADER_ENTRY_POINTS, FIELD_PROCEDURE_LISTED_IN_ENTRY_POINTS, FIELD_ENTRY_POINT_MATCH_CONFIDENCE, FIELD_MODULE_HEADER_SCAN_NOTES, FIELD_CANDIDATE_PROCEDURE_DECLARATION, FIELD_CANDIDATE_ARGUMENT_SIGNATURE, FIELD_HAS_REQUIRED_ARGUMENTS, FIELD_CANDIDATE_NOTES, FIELD_SELECTED_FOR_ACCEPTANCE, FIELD_ACCEPTANCE_STATUS, FIELD_ACCEPTED_ON, FIELD_ACCEPTED_CONFIGURATION_ID, FIELD_ACCEPTANCE_NOTES, FIELD_REJECTED_ON)
End Function

Private Function A2DeferredPhysicalRetirementPendingForSetup() As Boolean
    'Read-only A2 fast-path guard. It inspects only A2-owned obsolete physical storage; it does not inspect, delete, or otherwise take ownership of B5 relationships.
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim fieldName As Variant
    Dim configTableName As String
    Set db = TablesWorkspaceCodeDb()
    db.TableDefs.Refresh
    If SharedDao_TableExists(db, LEGACY_TABLE_PHYSICAL_ROWKIND_LOOKUP) Then A2DeferredPhysicalRetirementPendingForSetup = True: Exit Function
    configTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")
    If SharedDao_TableExists(db, configTableName) = False Then Exit Function
    Set tdf = db.TableDefs(configTableName)
    For Each fieldName In LegacyConfigCandidateFieldNames()
        PumpSetupControlEvents
        If SharedDao_TableDefFieldExistsSafe(tdf, CStr(fieldName)) Then A2DeferredPhysicalRetirementPendingForSetup = True: Exit Function
    Next fieldName
End Function

Private Sub RetireLegacyCandidateFieldsFromConfigTableForSetup(ByVal configTableName As String)
    'A2 owns physical field retirement. Existing B5 relationships can legitimately keep a retired legacy field alive until B5 removes the old relationship later in this Setup run.
    'The next A2 pass then removes that field.
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim fieldName As Variant
    Dim retiredFields As Variant
    Dim deferredCount As Long
    Dim removedCount As Long
    If CodeDbTableExists(configTableName) = False Then Exit Sub
    Set db = TablesWorkspaceCodeDb()
    Set tdf = db.TableDefs(configTableName)
    retiredFields = LegacyConfigCandidateFieldNames()
    For Each fieldName In retiredFields
        If SharedDao_TableDefFieldExistsSafe(tdf, CStr(fieldName)) Then
            On Error Resume Next
            Err.Clear
            tdf.fields.Delete CStr(fieldName)
            If Err.Number = 0 Then
                removedCount = removedCount + 1
            Else
                deferredCount = deferredCount + 1
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next fieldName
    If removedCount > 0 Then CoreAddOperationNote "A2 removed " & CStr(removedCount) & " retired candidate/discovery field(s) from legacy MC_Records storage after their data was migrated."
    If deferredCount > 0 Then CoreAddOperationNote "A2 deferred " & CStr(deferredCount) & " retired MC_Records candidate/discovery field(s) because existing B5-owned relationships still reference them. B5 removes those obsolete relationships in this run; the next A2 repair can then remove the remaining physical legacy columns without crossing phase ownership."
End Sub
Private Sub CreateConfigTable(ByVal tableName As String, ByVal applyLookupMetadata As Boolean, ByVal schemaDefinitions As Collection, ByVal indexDefinitions As Collection)
    On Error GoTo FirstAttemptFailed
    If SharedDao_QueryExists(TablesWorkspaceCodeDb(), tableName) Then
        SafeDropTableIfExists tableName, "Dropped saved query name conflict [" & tableName & "] before creating the macro-creator MC_Records table."
    End If
    CreateConfigTableAttempt tableName, applyLookupMetadata, schemaDefinitions, indexDefinitions
    Exit Sub
FirstAttemptFailed:
    Dim firstErrNumber As Long
    Dim firstErrSource As String
    Dim firstErrDescription As String
    firstErrNumber = Err.Number
    firstErrSource = Err.Source
    firstErrDescription = Err.Description
    If firstErrNumber = 3010 Or firstErrNumber = 3078 Or firstErrNumber = 3265 Or firstErrNumber = 7874 Then
        On Error GoTo RetryFailed
        SafeDropTableIfExists tableName, "Removed conflicting/stale MC_Records table object [" & tableName & "] before retrying table creation. " & "First create attempt failed with Err.Number " & CStr(firstErrNumber) & ": " & SharedMessage_OneLine(firstErrDescription, 250)
        CreateConfigTableAttempt tableName, applyLookupMetadata, schemaDefinitions, indexDefinitions
        Exit Sub
    End If
    Err.Raise firstErrNumber, "CreateConfigTable", "Could not create access development console MC_Records table." & vbCrLf & "Table: " & tableName & vbCrLf & "Err.Number: " & CStr(firstErrNumber) & vbCrLf & "Err.Source: " & firstErrSource & vbCrLf & "Err.Description: " & firstErrDescription
RetryFailed:
    Dim retryErrNumber As Long
    Dim retryErrSource As String
    Dim retryErrDescription As String
    retryErrNumber = Err.Number
    retryErrSource = Err.Source
    retryErrDescription = Err.Description
    Err.Raise retryErrNumber, "CreateConfigTable", "Could not create access development console MC_Records table after dropping a conflicting/stale table object." & vbCrLf & "Table: " & tableName & vbCrLf & "First Err.Number: " & CStr(firstErrNumber) & vbCrLf & "First Err.Source: " & firstErrSource & vbCrLf & "First Err.Description: " & firstErrDescription & vbCrLf & "Retry Err.Number: " & CStr(retryErrNumber) & vbCrLf & "Retry Err.Source: " & retryErrSource & vbCrLf & "Retry Err.Description: " & retryErrDescription
End Sub

Private Sub CreateConfigTableAttempt(ByVal tableName As String, ByVal applyLookupMetadata As Boolean, ByVal schemaDefinitions As Collection, ByVal indexDefinitions As Collection)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim fieldDefinition As Variant
    Set db = TablesWorkspaceCodeDb()
    Set tdf = db.CreateTableDef(tableName)
    For Each fieldDefinition In schemaDefinitions
        AppendSchemaField tdf, CStr(fieldDefinition(0)), CLng(fieldDefinition(1)), CLng(fieldDefinition(2)), CBool(fieldDefinition(3)), tdf.Name, fieldDefinition
    Next fieldDefinition
    db.TableDefs.Append tdf
    TrackSetupCreatedTableRollback tableName
    db.TableDefs.Refresh
    EnsureTableDefFieldOrderFromText db.TableDefs(tableName), SchemaDefinitionFieldOrderText(schemaDefinitions), mTableDesignBatchPrepared
    If applyLookupMetadata Then RepairConfigLookupFields tableName, False
    CoreAddOperationNote "Created configuration-only MC_Records table [" & tableName & "]."
    CoreAddOperationNote "Deferred Navigation Pane grouping for newly created MC_Records table [" & tableName & "] until the final catalog-driven setup pass."
    RepairConfigIndexes tableName, False, indexDefinitions
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrSource As String
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrSource = Err.Source
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "CreateConfigTableAttempt", "Could not create access development console MC_Records table." & vbCrLf & "Table: " & tableName & vbCrLf & "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB & vbCrLf & "Database: " & TablesWorkspaceCodeDb().Name & vbCrLf & "Err.Number: " & CStr(originalErrNumber) & vbCrLf & "Err.Source: " & originalErrSource & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

' Purpose: Creates or repairs only the HostDb MC_Records/config table from canonical schema, index, lookup, and field-description definitions, without installing CodeDb-owned utility objects.
' Public role: Public module entry point. This procedure is listed in the module header as an intentional externally callable ADC entry point and may be invoked directly by Access/VBA callers.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub RepairHostDbConfigTable(Optional ByVal showMessage As Boolean = True)
    'Narrow HostDb MC_Records storage authority: creates/repairs only that table while Core resolves the HostDb database context.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim tableName As String
    Dim messageText As String
    Set db = GetHostDb()
    tableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")
    If SharedAccess_PrepareOpenObjectForDesignChange(acTable, tableName, "repairing the HostDb MC_Records table") = False Then
        MarkSystemCommandCancelled
        Err.Raise vbObjectError + 8991, "RepairHostDbConfigTable", "HostDb MC_Records repair was cancelled because the affected open table was not approved for save and close."
    End If
    EnsureConfigTableInDatabase db, tableName
    messageText = "Access Development Console HostDb MC_Records table repair complete." & vbCrLf & vbCrLf & "Table: " & tableName & vbCrLf & "Database: " & SharedDao_SafeDatabaseName(db) & TableDiagnosticsLocationForMessage() & TraceBackendDiagnosticsForMessage()
    ReportCoreActionSuccess "RepairHostDbConfigTable", "TableRepair", DEFAULT_UTILITY_SHORT_NAME & " HostDb MC_Records table repair completed successfully.", messageText, "Repair HostDb MC_Records table", tableName, "Table", "HostDb", "HostDbConfigRepair", showMessage, DEFAULT_TABLE_MAINTENANCE_MESSAGE_TITLE, vbInformation, True, "Tables", False, visibleStatusMessage:="HostDb MC_Records table repair completed successfully."
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim databaseNameText As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not db Is Nothing Then databaseNameText = SharedDao_SafeDatabaseName(db)
    On Error GoTo 0
    If Len(Trim$(databaseNameText)) = 0 Then databaseNameText = "(HostDb could not be resolved)"
    messageText = "Access Development Console HostDb MC_Records table repair failed." & vbCrLf & vbCrLf & "Procedure: RepairHostDbConfigTable" & vbCrLf & "Table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & vbCrLf & "Database: " & databaseNameText & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description:" & vbCrLf & savedErrDescription & TableDiagnosticsLocationForMessage()
    LogTraceMessage TRACE_LEVEL_ERROR, "TableRepair", "RepairHostDbConfigTable", "HostDb MC_Records table repair failed.", messageText, vbNullString, savedErrNumber, savedErrSource, savedErrDescription, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config"), "Table", "HostDb", "HostDbConfigRepair"
    If showMessage Then
        CoreShowErrorMessage messageText, vbCritical, DEFAULT_TABLE_MAINTENANCE_MESSAGE_TITLE, "RepairHostDbConfigTable", "Tables"
        Exit Sub
    End If
    Err.Raise savedErrNumber, "RepairHostDbConfigTable", messageText
End Sub

Private Sub EnsureConfigTableInDatabase(ByVal db As DAO.Database, ByVal tableName As String)
    On Error GoTo ErrHandler
    Dim schemaDefinitions As Collection
    Dim indexDefinitions As Collection
    Dim macroKeyIntegritySummary As String
    If db Is Nothing Then Err.Raise vbObjectError + 6120, "EnsureConfigTableInDatabase", "Database reference is Nothing."
    Set schemaDefinitions = ConfigFieldSchemaDefinitions()
    Set indexDefinitions = ConfigIndexSchemaDefinitions()
    If Len(tableName) = 0 Then Err.Raise vbObjectError + 6121, "EnsureConfigTableInDatabase", "MC_Records table name is blank."
    If SharedDao_QueryExists(db, tableName) Then Err.Raise vbObjectError + 6122, "EnsureConfigTableInDatabase", "A saved query has the same name as the required HostDb MC_Records table." & vbCrLf & "Table/query name: " & tableName & vbCrLf & "Database: " & SharedDao_SafeDatabaseName(db)
    If SharedDao_TableExists(db, tableName) = False Then
        CreateConfigTableInDatabase db, tableName, schemaDefinitions
    ElseIf TableCanOpenInDatabase(db, tableName) = False Then
        Err.Raise vbObjectError + 6123, "EnsureConfigTableInDatabase", "HostDb MC_Records table exists but cannot be opened." & vbCrLf & "Table: " & tableName & vbCrLf & "Database: " & SharedDao_SafeDatabaseName(db)
    Else
        macroKeyIntegritySummary = ConfigMacroKeyIntegritySummary(db, tableName)
        If Len(Trim$(macroKeyIntegritySummary)) > 0 Then Err.Raise vbObjectError + 6129, "EnsureConfigTableInDatabase", "HostDb MC_Records key-integrity preflight failed. Existing MacroKey values must be nonblank and unique before the canonical required/unique schema can be applied." & vbCrLf & vbCrLf & macroKeyIntegritySummary
        EnsureConfigFieldsInDatabase db, tableName, schemaDefinitions
    End If
    RepairConfigIndexesInDatabase db, tableName, indexDefinitions
    RepairConfigFieldDescriptionsInDatabase db, tableName, schemaDefinitions
    EnsureConfigLookupsInDatabase db, tableName, schemaDefinitions
    CoreAddOperationNote "HostDb MC_Records table [" & tableName & "] was created/repaired from the canonical TABLES schema/default metadata without installing CodeDb-owned utility objects."
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrSource As String
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrSource = Err.Source
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "EnsureConfigTableInDatabase", "Could not create or repair HostDb MC_Records table." & vbCrLf & "Table: " & tableName & vbCrLf & "Database: " & SharedDao_SafeDatabaseName(db) & vbCrLf & "Err.Number: " & CStr(originalErrNumber) & vbCrLf & "Err.Source: " & originalErrSource & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

Private Sub CreateConfigTableInDatabase(ByVal db As DAO.Database, ByVal tableName As String, ByVal schemaDefinitions As Collection)
    On Error GoTo ErrHandler
    Dim tdf As DAO.tableDef
    Dim fieldDefinition As Variant
    Set tdf = db.CreateTableDef(tableName)
    For Each fieldDefinition In schemaDefinitions
        AppendSchemaField tdf, CStr(fieldDefinition(0)), CLng(fieldDefinition(1)), CLng(fieldDefinition(2)), CBool(fieldDefinition(3)), tdf.Name, fieldDefinition
    Next fieldDefinition
    db.TableDefs.Append tdf
    db.TableDefs.Refresh
    Set tdf = db.TableDefs(tableName)
    EnsureTableDefFieldOrderFromText tdf, SchemaDefinitionFieldOrderText(schemaDefinitions), mTableDesignBatchPrepared
    CoreAddOperationNote "Created HostDb MC_Records table [" & tableName & "] from the canonical TABLES MC_Records field schema."
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "CreateConfigTableInDatabase", "Could not create HostDb MC_Records table." & vbCrLf & "Table: " & tableName & vbCrLf & "Database: " & SharedDao_SafeDatabaseName(db) & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

Private Sub EnsureConfigFieldsInDatabase(ByVal db As DAO.Database, ByVal tableName As String, ByVal schemaDefinitions As Collection)
    Dim tdf As DAO.tableDef
    Dim fieldDefinition As Variant
    Dim reqSize As Variant
    Set tdf = db.TableDefs(tableName)
    For Each fieldDefinition In schemaDefinitions
        reqSize = DefaultManagedFieldReqSize(tableName, CStr(fieldDefinition(0)))
        If SharedDao_TableDefFieldExistsSafe(tdf, CStr(fieldDefinition(0))) And IsNull(reqSize) = False And CLng(fieldDefinition(1)) = CLng(dbText) Then
            If SharedDao_FieldDefinitionMatches(tdf.fields(CStr(fieldDefinition(0))), CLng(fieldDefinition(1)), CLng(fieldDefinition(2)), CBool(fieldDefinition(3))) = False Then
                RepairExplicitBoundedTextFieldInDatabase db, tableName, CStr(fieldDefinition(0)), CLng(fieldDefinition(2)), False
                Set tdf = db.TableDefs(tableName)
            End If
        End If
        AppendSchemaFieldIfMissing tdf, CStr(fieldDefinition(0)), CLng(fieldDefinition(1)), CLng(fieldDefinition(2)), CBool(fieldDefinition(3)), tdf.Name, fieldDefinition
        ApplySchemaFieldContract tdf.fields(CStr(fieldDefinition(0))), fieldDefinition, tdf.Name, tdf.Name
    Next fieldDefinition
    tdf.fields.Refresh
    EnsureTableDefFieldOrderFromText tdf, SchemaDefinitionFieldOrderText(schemaDefinitions), mTableDesignBatchPrepared
End Sub

Private Function TableCanOpenInDatabase(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    If SharedDao_TableExists(db, tableName) = False Then Exit Function
    Set rs = db.OpenRecordset(tableName, dbOpenSnapshot)
    TableCanOpenInDatabase = True
ExitHere:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Err.Clear
    Exit Function
ErrHandler:
    TableCanOpenInDatabase = False
    Resume ExitHere
End Function

Private Sub RepairConfigIndexesInDatabase(ByVal db As DAO.Database, ByVal tableName As String, ByVal indexDefinitions As Collection)
    Dim indexDefinition As Variant
    For Each indexDefinition In indexDefinitions
        EnsureSingleFieldIndex db, tableName, CStr(indexDefinition(0)), CStr(indexDefinition(1)), CBool(indexDefinition(2)), CBool(indexDefinition(3)), False, False
    Next indexDefinition
End Sub

Private Sub RepairConfigFieldDescriptionsInDatabase(ByVal db As DAO.Database, ByVal tableName As String, ByVal schemaDefinitions As Collection)
    On Error GoTo ErrHandler
    Dim codeDb As DAO.Database
    Dim rs As DAO.Recordset
    Dim tdf As DAO.tableDef
    Dim fieldDefinition As Variant
    Dim fieldName As String
    Dim descriptionText As String
    Dim descriptionShortByField As Object
    Dim managedObjectsTableName As String
    Dim managedObjectFieldsTableName As String
    Dim sqlText As String
    Set tdf = db.TableDefs(tableName)
    Set descriptionShortByField = CreateObject("Scripting.Dictionary")
    descriptionShortByField.CompareMode = vbTextCompare
    Set codeDb = ADCDb()
    managedObjectsTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID)
    managedObjectFieldsTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFields")
    If SharedDao_DatabaseTableExistsSafe(codeDb, managedObjectsTableName) = False Or SharedDao_DatabaseTableExistsSafe(codeDb, managedObjectFieldsTableName) = False Then Err.Raise vbObjectError + 6125, "RepairConfigFieldDescriptionsInDatabase", "The authoritative MAC_Fields.FieldDescriptionShort catalog is unavailable; HostDb DAO field descriptions cannot be repaired from a different authority."
    sqlText = "SELECT f.[" & FIELD_MANAGED_OBJECT_FIELD_NAME & "],f.[" & FIELD_MANAGED_OBJECT_FIELD_DESCRIPTION_SHORT & "] FROM [" & managedObjectFieldsTableName & "] AS f INNER JOIN [" & managedObjectsTableName & "] AS o ON f.[" & FIELD_MANAGED_OBJECT_ID & "]=o." & SQL_FIELD_ID & " WHERE Nz(f.[" & FIELD_MANAGED_OBJECT_FIELD_IS_ACTIVE & "],False)<>False AND Nz(o.[" & FIELD_MANAGED_OBJECT_STATUS & "]," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & " AND UCase(Trim(Nz(o." & SQL_FIELD_OBJECT_TYPE & ",'')))='TABLE' AND o." & SQL_FIELD_OBJECT_NAME & "=" & SharedDao_SqlQuote(tableName) & ";"
    Set rs = codeDb.OpenRecordset(sqlText, dbOpenSnapshot)
    Do While rs.EOF = False
        fieldName = Trim$(CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_NAME).value, vbNullString)))
        If Len(fieldName) > 0 Then
            If descriptionShortByField.Exists(fieldName) Then Err.Raise vbObjectError + 6128, "RepairConfigFieldDescriptionsInDatabase", "Duplicate active MAC_Fields.FieldDescriptionShort mappings exist for the HostDb MC_Records field." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName
            descriptionShortByField.Add fieldName, CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_DESCRIPTION_SHORT).value, vbNullString))
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    For Each fieldDefinition In schemaDefinitions
        fieldName = CStr(fieldDefinition(0))
        If SharedDao_TableDefFieldExistsSafe(tdf, fieldName) Then
            If descriptionShortByField.Exists(fieldName) = False Then Err.Raise vbObjectError + 6126, "RepairConfigFieldDescriptionsInDatabase", "No active MAC_Fields.FieldDescriptionShort mapping exists for the HostDb MC_Records field." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName
            descriptionText = CStr(descriptionShortByField(fieldName))
            If Len(Trim$(descriptionText)) = 0 Then Err.Raise vbObjectError + 6127, "RepairConfigFieldDescriptionsInDatabase", "The authoritative MAC_Fields.FieldDescriptionShort value is blank." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName
            If Len(descriptionText) > 255 Then Err.Raise vbObjectError + 6124, "RepairConfigFieldDescriptionsInDatabase", "The authoritative FieldDescriptionShort exceeds the Access DAO Field.Description 255-character limit." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName & vbCrLf & "Characters: " & CStr(Len(descriptionText)) & vbCrLf & "Default/value source: " & FieldDescriptionFailureSourceName(tableName, fieldName, descriptionText)
            EnsureDaoFieldDescription tdf.fields(fieldName), descriptionText, tdf.Name
        End If
    Next fieldDefinition
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
    Err.Raise originalErrNumber, "RepairConfigFieldDescriptionsInDatabase", "Could not repair HostDb MC_Records field descriptions." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

Private Function EnsureExternalCodeDbLookup(ByVal tdf As DAO.tableDef, ByVal fieldName As String, ByVal lookupDefinition As Variant) As Long
    'HostDb MC_Records stores an external RowSource back to the single CodeDb lookup-table authority instead of duplicating editable lookup data in HostDb.
    On Error GoTo ErrHandler
    Dim codeDb As DAO.Database
    Dim lookupTdf As DAO.tableDef
    Dim fld As DAO.Field
    Dim lookupTableName As String
    Dim codeFieldName As String
    Dim displayFieldName As String
    Dim rowSourceText As String
    Dim columnWidthsText As String
    Dim changed As Boolean
    If IsArray(lookupDefinition) = False Then Exit Function
    lookupTableName = Trim$(CStr(lookupDefinition(1)))
    codeFieldName = Trim$(CStr(lookupDefinition(2)))
    displayFieldName = Trim$(CStr(lookupDefinition(3)))
    columnWidthsText = Trim$(CStr(lookupDefinition(5)))
    If Len(columnWidthsText) = 0 Then columnWidthsText = "0;2"
    If SharedDao_TableDefFieldExistsSafe(tdf, fieldName) = False Then Exit Function
    Set codeDb = TablesWorkspaceCodeDb()
    If SharedDao_TableExists(codeDb, lookupTableName) = False Then Err.Raise vbObjectError + 9577, "EnsureExternalCodeDbLookup", "The required CodeDb editable lookup table is missing." & vbCrLf & "Table: " & lookupTableName
    Set lookupTdf = codeDb.TableDefs(lookupTableName)
    If SharedDao_TableDefFieldExistsSafe(lookupTdf, codeFieldName) = False Or SharedDao_TableDefFieldExistsSafe(lookupTdf, displayFieldName) = False Then Err.Raise vbObjectError + 9578, "EnsureExternalCodeDbLookup", "The required CodeDb editable lookup-table fields are missing." & vbCrLf & "Table: " & lookupTableName & vbCrLf & "Code field: " & codeFieldName & vbCrLf & "Display field: " & displayFieldName
    Set fld = tdf.fields(fieldName)
    rowSourceText = TableLookupRowSource(lookupTableName, codeFieldName, displayFieldName, SharedDao_TableDefFieldExistsSafe(lookupTdf, FIELD_LOOKUP_ACTIVE), False, SharedDao_TableDefFieldExistsSafe(lookupTdf, FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER), codeDb.Name)
    changed = ApplyLookupFieldProperties(fld, "Table/Query", rowSourceText, CBool(lookupDefinition(4)), 2, 1, NormalizeLookupColumnWidthsForAccess(columnWidthsText, 2, 1, CBool(lookupDefinition(4))), "EnsureExternalCodeDbLookup", vbObjectError + 6494, vbObjectError + 6495)
    If changed Then EnsureExternalCodeDbLookup = 1
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "EnsureExternalCodeDbLookup", "Could not configure a HostDb lookup field to use the authoritative CodeDb editable lookup table." & vbCrLf & "Host table: " & SafeDaoTableDefName(tdf) & vbCrLf & "Field: " & fieldName & vbCrLf & "Lookup table: " & lookupTableName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Sub EnsureConfigLookupsInDatabase(ByVal db As DAO.Database, ByVal tableName As String, ByVal schemaDefinitions As Collection)
    On Error GoTo ErrHandler
    Dim tdf As DAO.tableDef
    Dim fieldDefinition As Variant
    Dim fieldName As String
    Dim rowSourceText As String
    Dim lookupDefinition As Variant
    Dim lookupConfigured As Boolean
    Set tdf = db.TableDefs(tableName)
    For Each fieldDefinition In schemaDefinitions
        fieldName = CStr(fieldDefinition(0))
        rowSourceText = vbNullString
        lookupConfigured = False
        lookupDefinition = Empty
        If LookupFieldRequiresTableOrQuerySource(fieldName) Then lookupDefinition = FieldTableLookupDefinition(tableName, fieldName)
        If IsArray(lookupDefinition) Then
            If StrComp(CStr(lookupDefinition(0)), FIELD_LOOKUP_DEFINITION_KIND_TABLE, vbTextCompare) = 0 Then
                If EditableLookupTableName(CStr(lookupDefinition(1))) Then
                    Call EnsureExternalCodeDbLookup(tdf, fieldName, lookupDefinition)
                    lookupConfigured = True
                Else
                    rowSourceText = TableFieldLookupValueList(tableName, fieldName)
                    If Len(rowSourceText) > 0 Then Call EnsureFieldValueListLookup(tdf, fieldName, rowSourceText): lookupConfigured = True
                End If
            Else
                rowSourceText = TableFieldLookupValueList(tableName, fieldName)
                If Len(rowSourceText) > 0 Then Call EnsureFieldValueListLookup(tdf, fieldName, rowSourceText): lookupConfigured = True
            End If
        Else
            rowSourceText = TableFieldLookupValueList(tableName, fieldName)
            If Len(rowSourceText) > 0 Then Call EnsureFieldValueListLookup(tdf, fieldName, rowSourceText): lookupConfigured = True
        End If
        If lookupConfigured = False Then Call ClearFieldLookupProperties(tdf, fieldName)
    Next fieldDefinition
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "EnsureConfigLookupsInDatabase", "Could not repair HostDb MC_Records lookup metadata." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName & vbCrLf & "Err.Description: " & originalErrDescription
End Sub


' Purpose: Repairs launcher table foundation for setup for config metadata ready, settings defaults already ready.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: False.
'   configMetadataReady (Input Optional; Boolean): Boolean state describing whether config metadata ready applies to this invocation; True selects the named condition/path and False selects its
'   complementary/default path. Default: False.
'   settingsDefaultsAlreadyReady (Input Optional; Boolean): Boolean state describing whether settings defaults already ready applies to this invocation; True selects the named condition/path and False selects
'   its complementary/default path. Default: False.
'   applyLookupMetadata (Input Optional; Boolean): When True, enables or requests lookup metadata for this call; when False, the routine follows its normal non-forced/default path. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub RepairLauncherTableFoundationForSetup(Optional ByVal showMessage As Boolean = False, Optional ByVal configMetadataReady As Boolean = False, Optional ByVal settingsDefaultsAlreadyReady As Boolean = False, Optional ByVal applyLookupMetadata As Boolean = True)
    RequireSetupDomainMutationAllowed SETUP_DOMAIN_A2_PHYSICAL_SCHEMA, "RepairLauncherTableFoundationForSetup"
    On Error GoTo ErrHandler
    If settingsDefaultsAlreadyReady = False Then RepairLauncherSettingDefaults False
    If configMetadataReady = False And CodeDbTableExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")) Then
        EnsureLauncherConfigFields ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")
        ReconcileStandaloneManagedTableDescriptionIfAllowed ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")
    End If
    EnsureCustomMacroLauncherReviewTable
    If applyLookupMetadata Then EnsureCustomMacroLauncherReviewLookupFields ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LauncherControlsReview")
    ReconcileStandaloneManagedTableDescriptionIfAllowed ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LauncherControlsReview")
    If showMessage Then ReportCoreActionSuccess "RepairLauncherTableFoundationForSetup", "Launcher", "Launcher review table foundation repaired successfully.", "Saved-query repair is deferred to A3.", "Repair launcher table foundation", ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LauncherControlsReview"), "Table", "CodeDb", "Launcher", True, "Custom Macro Launcher Table Foundation", vbInformation, True, "Tables", False
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If showMessage Then
        CoreShowErrorMessage "Could not repair the launcher table foundation." & vbCrLf & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, vbCritical, "Custom Macro Launcher Table Foundation", "RepairLauncherTableFoundationForSetup", "Tables"
        Exit Sub
    End If
    Err.Raise savedErrNumber, "RepairLauncherTableFoundationForSetup", savedErrDescription
End Sub

Private Sub EnsureCustomMacroLauncherReviewTable()
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim fld As DAO.Field
    Dim fieldDefinition As Variant
    Dim schemaDefinitions As Collection
    Dim tableName As String
    Dim tableCreated As Boolean
    tableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LauncherControlsReview")
    Set schemaDefinitions = CustomMacroLauncherReviewSchemaDefinitions()
    tableCreated = (CodeDbTableExists(tableName) = False)
    If tableCreated Then
        Set db = TablesWorkspaceCodeDb()
        Set tdf = db.CreateTableDef(tableName)
        For Each fieldDefinition In schemaDefinitions
            If CLng(fieldDefinition(1)) = CLng(dbText) Then
                Set fld = tdf.CreateField(CStr(fieldDefinition(0)), CLng(fieldDefinition(1)), CLng(fieldDefinition(2)))
            Else
                Set fld = tdf.CreateField(CStr(fieldDefinition(0)), CLng(fieldDefinition(1)))
            End If
            If CBool(fieldDefinition(3)) Then fld.Attributes = dbAutoIncrField
            ApplySchemaFieldContract fld, fieldDefinition, tableName, "Custom Macro Launcher review"
            tdf.fields.Append fld
        Next fieldDefinition
        db.TableDefs.Append tdf
        TrackSetupCreatedTableRollback tableName
        db.TableDefs.Refresh
        CoreAddOperationNote "Created Custom Macro Launcher review table [" & tableName & "]."
        CoreAddOperationNote "Deferred Navigation Pane grouping for newly created Custom Macro Launcher review table [" & tableName & "] until the final catalog-driven setup pass."
    Else
        Set tdf = TablesWorkspaceCodeDb().TableDefs(tableName)
        For Each fieldDefinition In schemaDefinitions
            If CBool(fieldDefinition(3)) = False Then
                If SharedDao_TableDefFieldExistsSafe(tdf, CStr(fieldDefinition(0))) Then
                    RepairTableField tableName, CStr(fieldDefinition(0)), CLng(fieldDefinition(1)), CLng(fieldDefinition(2))
                Else
                    RepairTableField tableName, CStr(fieldDefinition(0)), CLng(fieldDefinition(1)), CLng(fieldDefinition(2))
                    tdf.fields.Refresh
                End If
            End If
            ApplySchemaFieldContract tdf.fields(CStr(fieldDefinition(0))), fieldDefinition, tableName, "Custom Macro Launcher review"
        Next fieldDefinition
    End If
    EnsureTableDefFieldOrderFromText TablesWorkspaceCodeDb().TableDefs(tableName), SchemaDefinitionFieldOrderText(schemaDefinitions), mTableDesignBatchPrepared
    Set db = TablesWorkspaceCodeDb()
    EnsureSingleFieldIndex db, tableName, INDEX_PRIMARY_KEY, FIELD_ID, True, True, mTableDesignBatchPrepared
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "EnsureCustomMacroLauncherReviewTable", "Could not ensure Custom Macro Launcher review table." & vbCrLf & "Table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LauncherControlsReview") & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

Private Sub EnsureCustomMacroLauncherReviewLookupFields(ByVal tableName As String)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim changedCount As Long
    Set db = TablesWorkspaceCodeDb()
    Set tdf = db.TableDefs(tableName)
    tdf.fields.Refresh
    changedCount = changedCount + EnsureDefinedFieldLookup(tdf, FIELD_CONFIG_ID)
    changedCount = changedCount + EnsureDefinedFieldLookup(tdf, FIELD_SEVERITY)
    changedCount = changedCount + EnsureDefinedFieldLookup(tdf, FIELD_ISSUE_TYPE)
    If changedCount > 0 Then CoreAddOperationNote "Repaired " & CStr(changedCount) & " lookup field(s) on table [" & tableName & "]."
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "EnsureCustomMacroLauncherReviewLookupFields", "Could not repair Custom Macro Launcher lookup metadata." & vbCrLf & "Table: " & tableName & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

Private Sub PrepareTableNameForCreateOrRepair(ByVal tableName As String, ByVal tableRole As String)
    On Error GoTo ErrHandler
    'A prepared A2 design batch already owns one coherent DAO collection synchronization.
    'Do not repeat full TableDefs/QueryDefs refreshes for every table name unless a conflict is actually removed.
    If mTableDesignBatchPrepared = False Then RefreshTableDefsSafely
    If SharedDao_QueryExists(TablesWorkspaceCodeDb(), tableName) Then
        SafeDropTableIfExists tableName, "Dropped saved query name conflict [" & tableName & "] before creating the macro-creator " & tableRole & " table."
        If mTableDesignBatchPrepared = False Then RefreshTableDefsSafely
    End If
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrSource As String
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrSource = Err.Source
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "PrepareTableNameForCreateOrRepair", "Could not prepare macro-creator table name for create/repair." & vbCrLf & "Table role: " & tableRole & vbCrLf & "Table: " & tableName & vbCrLf & "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB & vbCrLf & "Database: " & TablesWorkspaceCodeDb().Name & vbCrLf & "Err.Number: " & CStr(originalErrNumber) & vbCrLf & "Err.Source: " & originalErrSource & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

Private Sub RequirePreparedTableCanOpen(ByVal tableName As String, ByVal tableRole As String)
    On Error GoTo ErrHandler
    'The caller has just established physical existence after name-conflict preparation.
    'Open the known-existing table directly so the openability guard does not repeat the same TableDefs existence lookup.
    If TableCanOpen(tableName, True) Then Exit Sub
    Err.Raise vbObjectError + 6233, "RequirePreparedTableCanOpen", "Ordinary repair stopped because Access lists an existing macro-creator " & tableRole & " table definition that cannot be opened. The table was not deleted. Preserve or recover the table separately before retrying repair." & vbCrLf & "Table: " & tableName & vbCrLf & TableOpenFailureDetail(tableName)
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrSource As String
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrSource = Err.Source
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "RequirePreparedTableCanOpen", "Could not verify that the prepared macro-creator table can be opened for repair." & vbCrLf & "Table role: " & tableRole & vbCrLf & "Table: " & tableName & vbCrLf & "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB & vbCrLf & "Database: " & TablesWorkspaceCodeDb().Name & vbCrLf & "Err.Number: " & CStr(originalErrNumber) & vbCrLf & "Err.Source: " & originalErrSource & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

Private Sub PumpSetupControlEvents(Optional ByVal forcePump As Boolean = False)
    On Error Resume Next
    mSetupControlPumpRequestsIssuedCount = mSetupControlPumpRequestsIssuedCount + 1
    Call SetupControlPump(forcePump)
    On Error GoTo 0
End Sub

Private Function ChatGPTVersionIDFromCodeModule(ByVal codeModule As Object) As String
    On Error GoTo Unavailable
    Dim lineCount As Long
    Dim inspectLineCount As Long
    Dim moduleText As String
    Dim moduleLines As Variant
    Dim lineItem As Variant
    Dim trimmedLine As String
    Const VERSION_PREFIX As String = "' ChatGPT Version ID:"
    lineCount = CLng(codeModule.CountOfLines)
    If lineCount <= 0 Then
        ChatGPTVersionIDFromCodeModule = "<not present>"
        Exit Function
    End If
    inspectLineCount = lineCount
    If inspectLineCount > 160 Then inspectLineCount = 160
    moduleText = CStr(codeModule.lines(1, inspectLineCount))
    moduleText = Replace(moduleText, vbCrLf, vbLf)
    moduleText = Replace(moduleText, vbCr, vbLf)
    moduleLines = Split(moduleText, vbLf)
    For Each lineItem In moduleLines
        PumpSetupControlEvents
        trimmedLine = Trim$(CStr(lineItem))
        If StrComp(Left$(trimmedLine, Len(VERSION_PREFIX)), VERSION_PREFIX, vbTextCompare) = 0 Then
            ChatGPTVersionIDFromCodeModule = Trim$(Mid$(trimmedLine, Len(VERSION_PREFIX) + 1))
            If Len(ChatGPTVersionIDFromCodeModule) = 0 Then ChatGPTVersionIDFromCodeModule = "<blank>"
            Exit Function
        End If
    Next lineItem
    ChatGPTVersionIDFromCodeModule = "<not present>"
    Exit Function
Unavailable:
    ChatGPTVersionIDFromCodeModule = "<unavailable>"
    Err.Clear
End Function

Private Sub LogChatGPTModuleVersionInventory(ByVal contextName As String)
    On Error GoTo InventoryWarning
    Dim vbProject As Object
    Dim vbComponent As Object
    Dim inventoryRows() As String
    Dim componentCount As Long
    Dim rowIndex As Long
    Dim trackedCount As Long
    Dim absentCount As Long
    Dim unavailableCount As Long
    Dim versionText As String
    Dim detailText As String
    Dim inventoryStarted As Single
    inventoryStarted = Timer
    On Error Resume Next
    Set vbProject = VBProjectForConfigOwnerDatabaseRole("CodeDb")
    On Error GoTo InventoryWarning
    If vbProject Is Nothing Then Set vbProject = Application.vbe.ActiveVBProject
    If vbProject Is Nothing Then Err.Raise vbObjectError + 9040, "LogChatGPTModuleVersionInventory", "The CodeDb VBProject could not be resolved."
    componentCount = CLng(vbProject.VBComponents.count)
    If componentCount <= 0 Then
        detailText = "VBProject: " & CStr(vbProject.Name) & vbCrLf & "Modules: 0"
    Else
        ReDim inventoryRows(1 To componentCount)
        For Each vbComponent In vbProject.VBComponents
            PumpSetupControlEvents
            rowIndex = rowIndex + 1
            versionText = ChatGPTVersionIDFromCodeModule(vbComponent.codeModule)
            inventoryRows(rowIndex) = CStr(vbComponent.Name) & " = " & versionText
            If StrComp(versionText, "<not present>", vbTextCompare) = 0 Or StrComp(versionText, "<blank>", vbTextCompare) = 0 Then
                absentCount = absentCount + 1
            ElseIf StrComp(versionText, "<unavailable>", vbTextCompare) = 0 Then
                unavailableCount = unavailableCount + 1
            Else
                trackedCount = trackedCount + 1
            End If
        Next vbComponent
        SortVersionInventoryRows inventoryRows, rowIndex
        detailText = "VBProject: " & CStr(vbProject.Name)
        detailText = detailText & vbCrLf & "Modules inspected: " & CStr(rowIndex)
        detailText = detailText & vbCrLf & "Tracked versions found: " & CStr(trackedCount)
        detailText = detailText & vbCrLf & "Headers absent or blank: " & CStr(absentCount)
        detailText = detailText & vbCrLf & "Read failures: " & CStr(unavailableCount)
        detailText = detailText & vbCrLf & "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(inventoryStarted))
        For rowIndex = 1 To componentCount
            PumpSetupControlEvents
            detailText = detailText & vbCrLf & inventoryRows(rowIndex)
        Next rowIndex
    End If
    If componentCount <= 0 Then detailText = detailText & vbCrLf & "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(inventoryStarted))
    LogTraceMessage TRACE_LEVEL_INFO, "VersionInventory", "LogChatGPTModuleVersionInventory", "ChatGPT module version inventory captured.", detailText, "Read leading module rewrite-tracking headers", 0, vbNullString, vbNullString, vbNullString, "VBProject", "CodeDb", contextName
    Exit Sub
InventoryWarning:
    Dim inventoryErrNumber As Long
    Dim inventoryErrSource As String
    Dim inventoryErrDescription As String
    inventoryErrNumber = Err.Number
    inventoryErrSource = Err.Source
    inventoryErrDescription = Err.Description
    On Error Resume Next
    LogTraceMessage TRACE_LEVEL_WARNING, "VersionInventory", "LogChatGPTModuleVersionInventory", "ChatGPT module version inventory could not be captured.", "The trace run will continue without the optional version inventory." & vbCrLf & "Elapsed before failure: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(inventoryStarted)) & vbCrLf & "Err.Number: " & CStr(inventoryErrNumber) & vbCrLf & "Err.Source: " & inventoryErrSource & vbCrLf & "Err.Description: " & inventoryErrDescription, "Read leading module rewrite-tracking headers", inventoryErrNumber, inventoryErrSource, inventoryErrDescription, vbNullString, "VBProject", "CodeDb", contextName
    Err.Clear
End Sub

Private Sub SortVersionInventoryRows(ByRef inventoryRows() As String, ByVal rowCount As Long)
    Dim outerIndex As Long
    Dim innerIndex As Long
    Dim swapText As String
    For outerIndex = 1 To rowCount - 1
        PumpSetupControlEvents
        For innerIndex = outerIndex + 1 To rowCount
            PumpSetupControlEvents
            If StrComp(inventoryRows(outerIndex), inventoryRows(innerIndex), vbTextCompare) > 0 Then
                swapText = inventoryRows(outerIndex)
                inventoryRows(outerIndex) = inventoryRows(innerIndex)
                inventoryRows(innerIndex) = swapText
            End If
        Next innerIndex
    Next outerIndex
End Sub

Private Sub SetTablesStep(ByRef currentStep As String, ByVal stepText As String)
    currentStep = stepText
    AddTablesTraceStep stepText
    CoreUpdateSetupProgress stepText
End Sub

Private Sub AddFoundationName(ByRef names As Collection, ByVal candidateName As String)
    Dim item As Variant
    candidateName = Trim$(candidateName)
    If Len(candidateName) = 0 Then Exit Sub
    For Each item In names
        PumpSetupControlEvents
        If StrComp(CStr(item), candidateName, vbTextCompare) = 0 Then Exit Sub
    Next item
    names.Add candidateName
End Sub

Private Function CalculateFoundationSchemaSnapshot(ByRef snapshot As FoundationDomainSnapshot, Optional ByVal suppliedCoreSettingDefinitions As Variant, Optional ByVal tableDefsAlreadySynchronized As Boolean = False) As Boolean
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim tableNames As Collection
    Dim coreSettingDefinitions As Collection
    Dim item As Variant
    Dim itemText As String
    Dim schemaPayload As String
    snapshot.failureDetail = vbNullString
    If isMissing(suppliedCoreSettingDefinitions) = False Then
        If IsObject(suppliedCoreSettingDefinitions) Then Set coreSettingDefinitions = suppliedCoreSettingDefinitions
    End If
    If coreSettingDefinitions Is Nothing Then Set coreSettingDefinitions = x__UTIL_ADC_DEFAULTS.GetCoreSettingDefinitions()
    Set db = TablesWorkspaceCodeDb()
    If tableDefsAlreadySynchronized = False Then db.TableDefs.Refresh
    db.Relations.Refresh
    Set tableNames = FoundationRequiredTableNames()
    schemaPayload = FingerprintTokenPart(FOUNDATION_SCHEMA_FINGERPRINT_VERSION & "|Schema")
    For Each item In tableNames
        PumpSetupControlEvents
        itemText = FoundationTableDefinitionText(db, CStr(item), snapshot.failureDetail)
        If Len(snapshot.failureDetail) > 0 Then Exit Function
        schemaPayload = schemaPayload & itemText
        snapshot.tableCount = snapshot.tableCount + 1
    Next item
    Dim part2DataSnapshot As FoundationDomainSnapshot
    If CalculateFoundationPart2DataSnapshot(part2DataSnapshot, coreSettingDefinitions, True) = False Then
        snapshot.failureDetail = part2DataSnapshot.failureDetail
        Exit Function
    End If
    snapshot.Part2DataFingerprint = part2DataSnapshot.Part2DataFingerprint
    snapshot.SchemaFingerprint = FingerprintFromPayload(schemaPayload)
    CalculateFoundationSchemaSnapshot = True
    Exit Function
ErrHandler:
    snapshot.failureDetail = "Could not calculate the " & DEFAULT_UTILITY_LONG_NAME & " core data-table foundation fingerprint. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
End Function



Private Function CalculateFoundationPart2DataSnapshot(ByRef snapshot As FoundationDomainSnapshot, Optional ByVal suppliedCoreSettingDefinitions As Variant, Optional ByVal tableDefsAlreadySynchronized As Boolean = False) As Boolean
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim tableNames As Collection
    Dim coreSettingDefinitions As Collection
    Dim item As Variant
    Dim itemText As String
    Dim payloadText As String
    snapshot.failureDetail = vbNullString
    If isMissing(suppliedCoreSettingDefinitions) = False Then
        If IsObject(suppliedCoreSettingDefinitions) Then Set coreSettingDefinitions = suppliedCoreSettingDefinitions
    End If
    If coreSettingDefinitions Is Nothing Then Set coreSettingDefinitions = x__UTIL_ADC_DEFAULTS.GetCoreSettingDefinitions()
    Set db = TablesWorkspaceCodeDb()
    If tableDefsAlreadySynchronized = False Then db.TableDefs.Refresh
    Set tableNames = FoundationPart2DataTableNames()
    payloadText = FingerprintTokenPart(FOUNDATION_FINGERPRINT_VERSION & "|Part2Data")
    For Each item In tableNames
        PumpSetupControlEvents
        itemText = FoundationCatalogDataText(db, CStr(item), snapshot.failureDetail)
        If Len(snapshot.failureDetail) > 0 Then Exit Function
        payloadText = payloadText & itemText
    Next item
    payloadText = payloadText & FoundationCoreSettingDefinitionText(coreSettingDefinitions)
    snapshot.Part2DataFingerprint = FingerprintFromPayload(payloadText)
    CalculateFoundationPart2DataSnapshot = True
    Exit Function
ErrHandler:
    snapshot.failureDetail = "Could not calculate the " & DEFAULT_UTILITY_LONG_NAME & " Part 2 default/reference-data fingerprint. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
End Function

Private Sub PersistFoundationSchemaSnapshot(ByRef snapshot As FoundationDomainSnapshot)
    Dim db As DAO.Database
    Set db = TablesWorkspaceCodeDb()
    Call EnsureDaoPropertyValue(db, FOUNDATION_SCHEMA_FINGERPRINT_PROPERTY, dbText, snapshot.SchemaFingerprint, "PersistFoundationSchemaSnapshot")
    Call EnsureDaoPropertyValue(db, FOUNDATION_PART2_DATA_FINGERPRINT_PROPERTY, dbText, snapshot.Part2DataFingerprint, "PersistFoundationSchemaSnapshot")
End Sub

Private Function FoundationDomainSummary(ByRef snapshot As FoundationDomainSnapshot) As String
    FoundationDomainSummary = "Foundation tables: " & CStr(snapshot.tableCount) & "; core queries: " & CStr(snapshot.coreQueryCount) & "; retained form-design source queries: " & CStr(snapshot.FormDesignSourceCount) & "."
End Function

Private Function FoundationIncludeCatalogDataField(ByVal fieldName As String) As Boolean
    Select Case UCase$(fieldName)
        Case UCase$(FIELD_SETTING_VALUE), UCase$(FIELD_UPDATED_ON), UCase$(FIELD_LAST_UPDATED_ON), UCase$(FIELD_LAST_DEFAULTED_ON), UCase$(FIELD_CREATED_ON), "INSTALLEDON", UCase$(FIELD_LAST_VERIFIED_ON), "LASTSEENON", "LASTRUNON"
            Exit Function
    End Select
    FoundationIncludeCatalogDataField = True
End Function

Private Function FoundationCatalogDataText(ByVal db As DAO.Database, ByVal tableName As String, ByRef failureDetail As String) As String
    On Error GoTo ErrHandler
    Dim tdf As DAO.tableDef
    Dim fld As DAO.Field
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim orderFieldName As String
    Dim resultText As String
    If SharedDao_TableExists(db, tableName) = False Then
        failureDetail = "Required catalog/default table is missing: " & tableName
        Exit Function
    End If
    Set tdf = db.TableDefs(tableName)
    tdf.fields.Refresh
    If SharedDao_TableDefFieldExistsSafe(tdf, FIELD_ID) Then
        orderFieldName = FIELD_ID
    ElseIf tdf.fields.count > 0 Then
        orderFieldName = tdf.fields(0).Name
    End If
    sqlText = "SELECT * FROM " & SharedDao_BracketName(tableName)
    If Len(orderFieldName) > 0 Then sqlText = sqlText & " ORDER BY " & SharedDao_BracketName(orderFieldName)
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    resultText = FingerprintTokenPart("DATA:" & tableName)
    Do While Not rs.EOF
        PumpSetupControlEvents
        For Each fld In rs.fields
            PumpSetupControlEvents
            If FoundationIncludeCatalogDataField(fld.Name) Then
                resultText = resultText & FingerprintTokenPart(fld.Name)
                resultText = resultText & FingerprintTokenPart(FoundationValueText(fld.value))
            End If
        Next fld
        resultText = resultText & vbLf
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    FoundationCatalogDataText = resultText
    Exit Function
ErrHandler:
    failureDetail = "Could not fingerprint foundation catalog/default rows in [" & tableName & "]. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
End Function

Private Function FoundationCoreSettingDefinitionText(ByVal coreSettingDefinitions As Collection) As String
    'Makes the Part-2 fast path depend on the canonical required-row contract instead of only the rows already present in Settings.
    Dim defaultRow As Variant
    Dim fieldIndex As Long
    Dim resultText As String
    resultText = FingerprintTokenPart("AUTHORITY:CoreSettingDefinitions")
    For Each defaultRow In coreSettingDefinitions
        PumpSetupControlEvents
        For fieldIndex = LBound(defaultRow) To UBound(defaultRow)
            resultText = resultText & FingerprintTokenPart(FoundationValueText(defaultRow(fieldIndex)))
        Next fieldIndex
    Next defaultRow
    FoundationCoreSettingDefinitionText = resultText
End Function

' ===============================================================================================================================================================================
' SETTINGS / DEFAULT STORAGE PROCEDURES
' ===============================================================================================================================================================================

' Purpose: Returns the single authoritative Settings physical field definitions for schema, repair, validation, and metadata consumers.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: Collection: Returns the collection/object containing the canonical settings field schema definitions selected by the supplied filters/context; an empty collection/object is used where the
'   implementation defines no matches.
Public Function SettingsFieldSchemaDefinitions() As Collection
    'Returns the single authoritative Settings physical field definitions for schema, repair, validation, and metadata consumers.
    ' CANONICAL AUTHORITY: rows owns the in-memory definition set constructed by SettingsFieldSchemaDefinitions; downstream consumers derive views from this set rather than create a competing authority.
    Dim rows As New Collection
    rows.Add Array(FIELD_ID, dbLong, 0, True)
    rows.Add Array(FIELD_SETTING_KEY, dbText, 100, False, True, False, Null, Null, Null)
    rows.Add Array(FIELD_SETTING_VALUE, dbMemo, 0, False)
    rows.Add Array(FIELD_MANAGED_OBJECT_FIELD_DEFAULT_VALUE, dbMemo, 0, False)
    rows.Add Array(FIELD_MANAGED_OBJECT_FIELD_DATA_TYPE, dbText, 50, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_DESCRIPTION, dbMemo, 0, False)
    rows.Add Array(FIELD_SETTING_GROUP, dbText, 100, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER, dbLong, 0, False)
    rows.Add Array(FIELD_ACTIVE, dbBoolean, 0, False)
    rows.Add Array(FIELD_ADVANCED, dbBoolean, 0, False)
    rows.Add Array(FIELD_READ_ONLY, dbBoolean, 0, False)
    rows.Add Array(FIELD_REQUIRED, dbBoolean, 0, False)
    rows.Add Array(FIELD_RESTART_REQUIRED, dbBoolean, 0, False)
    rows.Add Array(FIELD_VALIDATION_NOTES, dbMemo, 0, False)
    rows.Add Array(FIELD_LAST_DEFAULTED_ON, dbDate, 0, False)
    rows.Add Array(FIELD_LAST_UPDATED_ON, dbDate, 0, False)
    Set SettingsFieldSchemaDefinitions = rows
End Function

Private Function SettingsIndexSchemaDefinitions() As Collection
    'Single Settings index-definition authority consumed by repair and validation.
    ' CANONICAL AUTHORITY: rows owns the in-memory definition set constructed by SettingsIndexSchemaDefinitions; downstream consumers derive views from this set rather than create a competing authority.
    Dim rows As New Collection
    rows.Add Array(INDEX_PK_SETTINGS_ID, FIELD_ID, True, True)
    rows.Add Array(INDEX_UX_SETTING_KEY, FIELD_SETTING_KEY, True, False)
    rows.Add Array(INDEX_IX_SETTING_GROUP, FIELD_SETTING_GROUP, False, False)
    rows.Add Array(INDEX_IX_SETTING_SORT_ORDER, FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER, False, False)
    rows.Add Array(INDEX_IX_SETTING_ACTIVE, FIELD_ACTIVE, False, False)
    Set SettingsIndexSchemaDefinitions = rows
End Function

Private Sub AddSettingsTableSchemaIssues(ByRef issues As Collection, Optional ByVal suppliedSchemaDefinitions As Variant, Optional ByVal suppliedIndexDefinitions As Variant, Optional ByVal tableReadyAlreadyKnown As Boolean = False)
    Dim indexDefinition As Variant
    Dim tableName As String
    Dim schemaDefinitions As Collection
    Dim indexDefinitions As Collection
    If isMissing(suppliedSchemaDefinitions) = False Then
        If IsObject(suppliedSchemaDefinitions) Then Set schemaDefinitions = suppliedSchemaDefinitions
    End If
    If isMissing(suppliedIndexDefinitions) = False Then
        If IsObject(suppliedIndexDefinitions) Then Set indexDefinitions = suppliedIndexDefinitions
    End If
    If schemaDefinitions Is Nothing Then Set schemaDefinitions = SettingsFieldSchemaDefinitions()
    If indexDefinitions Is Nothing Then Set indexDefinitions = SettingsIndexSchemaDefinitions()
    tableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID)
    If tableReadyAlreadyKnown = False Then
        If CodeDbTableExists(tableName) = False Then Exit Sub
        If TableCanOpen(tableName, True) = False Then
            issues.Add tableName & ": table definition exists but Access cannot open it; it will be treated as stale/broken and recreated"
            Exit Sub
        End If
    End If
    AddFieldDefinitionIssuesFromSchema issues, tableName, schemaDefinitions
    For Each indexDefinition In indexDefinitions
        AddIndexDefinitionIssue issues, tableName, CStr(indexDefinition(0)), CStr(indexDefinition(1)), CBool(indexDefinition(2)), CBool(indexDefinition(3))
    Next indexDefinition
End Sub

Private Function SettingsTableSchemaIsValid(Optional ByVal suppliedSchemaDefinitions As Variant, Optional ByVal suppliedIndexDefinitions As Variant) As Boolean
    Dim issues As Collection
    Dim schemaDefinitions As Collection
    Dim indexDefinitions As Collection
    If isMissing(suppliedSchemaDefinitions) = False Then
        If IsObject(suppliedSchemaDefinitions) Then Set schemaDefinitions = suppliedSchemaDefinitions
    End If
    If isMissing(suppliedIndexDefinitions) = False Then
        If IsObject(suppliedIndexDefinitions) Then Set indexDefinitions = suppliedIndexDefinitions
    End If
    If schemaDefinitions Is Nothing Then Set schemaDefinitions = SettingsFieldSchemaDefinitions()
    If indexDefinitions Is Nothing Then Set indexDefinitions = SettingsIndexSchemaDefinitions()
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: issues holds planned/recovery work or independently observed validation evidence for SettingsTableSchemaIsValid; keep it distinct from the source authority.
    Set issues = New Collection
    If CodeDbTableExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID)) = False Then Exit Function
    If TableCanOpen(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID), True) = False Then Exit Function
    AddSettingsTableSchemaIssues issues, schemaDefinitions, indexDefinitions, True
    SettingsTableSchemaIsValid = (issues.count = 0)
End Function

' Purpose: Repairs code DB settings table for setup for apply lookup metadata, supplied schema definitions.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   applyLookupMetadata (Input Optional; Boolean): When True, enables or requests lookup metadata for this call; when False, the routine follows its normal non-forced/default path. Default: True.
'   suppliedSchemaDefinitions (Input Optional; Variant): Caller-supplied supplied schema definitions data used as the authoritative definition input for this call, avoiding an unnecessary reload when already
'   available.
'   suppliedIndexDefinitions (Input Optional; Variant): Caller-supplied supplied index definitions data used as the authoritative definition input for this call, avoiding an unnecessary reload when already
'   available.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub RepairCodeDbSettingsTableForSetup(Optional ByVal applyLookupMetadata As Boolean = True, Optional ByVal suppliedSchemaDefinitions As Variant, Optional ByVal suppliedIndexDefinitions As Variant)
    RequireSetupDomainMutationAllowed SETUP_DOMAIN_A2_PHYSICAL_SCHEMA, "RepairCodeDbSettingsTableForSetup"
    On Error GoTo ErrHandler
    Dim tableName As String
    Dim tableCreatedThisCall As Boolean
    Dim schemaDefinitions As Collection
    Dim indexDefinitions As Collection
    tableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID)
    If isMissing(suppliedSchemaDefinitions) = False Then
        If IsObject(suppliedSchemaDefinitions) Then Set schemaDefinitions = suppliedSchemaDefinitions
    End If
    If isMissing(suppliedIndexDefinitions) = False Then
        If IsObject(suppliedIndexDefinitions) Then Set indexDefinitions = suppliedIndexDefinitions
    End If
    If schemaDefinitions Is Nothing Then Set schemaDefinitions = SettingsFieldSchemaDefinitions()
    If indexDefinitions Is Nothing Then Set indexDefinitions = SettingsIndexSchemaDefinitions()
    PrepareTableNameForCreateOrRepair tableName, "settings"
    If CodeDbTableExists(tableName) = False Then
        CreateSettingsTable tableName, applyLookupMetadata, schemaDefinitions, indexDefinitions
        tableCreatedThisCall = True
    Else
        RequirePreparedTableCanOpen tableName, "settings"
        If SettingsTableHasIncompatibleExistingFieldDefinitions(tableName, schemaDefinitions) Then
            Err.Raise vbObjectError + 5329, "RepairCodeDbSettingsTableForSetup", DEFAULT_UTILITY_LONG_NAME & " CodeDb settings table exists but does not match the current managed-table design schema." & vbCrLf & "CodeDb settings table: " & tableName
        Else
            EnsureSettingsSchemaFields tableName, schemaDefinitions
            EnsureSettingsIndexes tableName, indexDefinitions
        End If
    End If
    If tableCreatedThisCall = False Then
        If applyLookupMetadata Then EnsureSettingsLookupFields tableName
    End If
    VerifyTableReady tableName, "settings"
    SetNamingBootstrapTableNames tableName, BootstrapManagedObjectsTableName()
    ReconcileStandaloneManagedTableDescriptionIfAllowed tableName
    LogTraceMessage TRACE_LEVEL_INFO, "TableRepair", "RepairCodeDbSettingsTableForSetup", "A2 settings-table structure repaired and verified.", "Table: " & tableName, "Repair A2 settings table", 0, vbNullString, vbNullString, tableName, "Table", "CodeDb", "TableRepair"
    Exit Sub
ErrHandler:
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Err.Raise errNumber, "RepairCodeDbSettingsTableForSetup", "Could not create or repair the A2 settings table." & vbCrLf & "CodeDb settings table: " & tableName & vbCrLf & "Err.Source: " & errSource & vbCrLf & "Err.Description: " & errDescription
End Sub

Private Function SettingsTableHasIncompatibleExistingFieldDefinitions(ByVal tableName As String, ByVal schemaDefinitions As Collection) As Boolean
    Dim fieldDefinition As Variant
    For Each fieldDefinition In schemaDefinitions
        If SettingsExistingFieldWrong(tableName, CStr(fieldDefinition(0)), CLng(fieldDefinition(1)), CLng(fieldDefinition(2)), CBool(fieldDefinition(3))) Then
            SettingsTableHasIncompatibleExistingFieldDefinitions = True
            Exit Function
        End If
    Next fieldDefinition
End Function

Private Function SettingsExistingFieldWrong(ByVal tableName As String, ByVal fieldName As String, ByVal expectedType As DAO.DataTypeEnum, Optional ByVal expectedSize As Long = 0, Optional ByVal mustBeAutoNumber As Boolean = False) As Boolean
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim fld As DAO.Field
    If CodeDbTableExists(tableName) = False Then Exit Function
    Set db = TablesWorkspaceCodeDb()
    db.TableDefs.Refresh
    Set tdf = db.TableDefs(tableName)
    tdf.fields.Refresh
    If SharedDao_TableDefFieldExistsSafe(tdf, fieldName) = False Then Exit Function
    Set fld = tdf.fields(fieldName)
    SettingsExistingFieldWrong = (SharedDao_FieldDefinitionMatches(fld, expectedType, expectedSize, mustBeAutoNumber) = False)
    Exit Function
ErrHandler:
    SettingsExistingFieldWrong = True
End Function

Private Sub EnsureSettingsSchemaFields(ByVal tableName As String, ByVal schemaDefinitions As Collection)
    On Error GoTo ErrHandler
    Dim fieldDefinition As Variant
    For Each fieldDefinition In schemaDefinitions
        EnsureSettingsField tableName, CStr(fieldDefinition(0)), CLng(fieldDefinition(1)), CLng(fieldDefinition(2)), CBool(fieldDefinition(3)), fieldDefinition
    Next fieldDefinition
    EnsureTableDefFieldOrderFromText TablesWorkspaceCodeDb().TableDefs(tableName), SchemaDefinitionFieldOrderText(schemaDefinitions), mTableDesignBatchPrepared
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "EnsureSettingsSchemaFields", "Could not add/verify settings-table fields." & vbCrLf & "Table: " & tableName & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

Private Sub EnsureSettingsField(ByVal tableName As String, ByVal fieldName As String, ByVal fieldType As DAO.DataTypeEnum, Optional ByVal fieldSize As Long = 0, Optional ByVal autoNumber As Boolean = False, Optional ByVal fieldDefinition As Variant)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim fld As DAO.Field
    Set db = TablesWorkspaceCodeDb()
    db.TableDefs.Refresh
    Set tdf = db.TableDefs(tableName)
    tdf.fields.Refresh
    If SharedDao_TableDefFieldExistsSafe(tdf, fieldName) Then
        Set fld = tdf.fields(fieldName)
        If SharedDao_FieldDefinitionMatches(fld, fieldType, fieldSize, autoNumber) = False Then
            Err.Raise vbObjectError + 5327, "EnsureSettingsField", "Existing settings field has an incompatible definition and cannot be safely altered in place." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName & vbCrLf & "Expected: " & FieldDefinitionDescription(fieldType, fieldSize, autoNumber) & vbCrLf & "Found: " & FieldActualDefinitionDescription(fld)
        End If
        If isMissing(fieldDefinition) = False Then ApplySchemaFieldContract fld, fieldDefinition, tableName, "settings"
        Exit Sub
    End If
    If fieldType = dbText And fieldSize > 0 Then
        Set fld = tdf.CreateField(fieldName, fieldType, fieldSize)
    Else
        Set fld = tdf.CreateField(fieldName, fieldType)
    End If
    If autoNumber Then fld.Attributes = fld.Attributes Or dbAutoIncrField
    If isMissing(fieldDefinition) = False Then ApplySchemaFieldContract fld, fieldDefinition, tableName, "settings"
    tdf.fields.Append fld
    tdf.fields.Refresh
    CoreAddOperationNote "Added missing settings field [" & fieldName & "] to [" & tableName & "]."
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "EnsureSettingsField", "Could not add or verify a settings-table field." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

Private Sub EnsureSettingsIndexes(ByVal tableName As String, ByVal indexDefinitions As Collection)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim indexDefinition As Variant
    Set db = TablesWorkspaceCodeDb()
    EnsureNoDuplicateSettingKeysBeforeIndex tableName
    For Each indexDefinition In indexDefinitions
        If TableFieldExists(tableName, CStr(indexDefinition(1))) Then
            EnsureSingleFieldIndex db, tableName, CStr(indexDefinition(0)), CStr(indexDefinition(1)), CBool(indexDefinition(2)), CBool(indexDefinition(3)), mTableDesignBatchPrepared
        End If
    Next indexDefinition
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "EnsureSettingsIndexes", "Could not create or repair settings-table indexes." & vbCrLf & "Table: " & tableName & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

Private Sub EnsureNoDuplicateSettingKeysBeforeIndex(ByVal tableName As String)
    Dim duplicateSummary As String
    duplicateSummary = DuplicateKeySummaryForTable(tableName, FIELD_SETTING_KEY)
    If Len(Trim$(duplicateSummary)) > 0 Then
        Err.Raise vbObjectError + 5336, "EnsureNoDuplicateSettingKeysBeforeIndex", "Duplicate SettingKey values must be corrected before the unique SettingKey index can be created or repaired." & vbCrLf & vbCrLf & duplicateSummary
    End If
End Sub

Private Sub EnsureSettingsLookupFields(ByVal tableName As String)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim changedCount As Long
    Dim ignoredWarningCount As Long
    If CodeDbTableExists(tableName) = False Then Exit Sub
    If StrComp(tableName, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_TraceLog"), vbTextCompare) = 0 Then CloseRetainedTraceWriteRecordset
    Set db = TablesWorkspaceCodeDb()
    db.TableDefs.Refresh
    Set tdf = db.TableDefs(tableName)
    changedCount = ApplyB2SharedFieldMetadataDefinitionsToTableDef(tdf, ignoredWarningCount)
    If changedCount > 0 Then CoreAddOperationNote "Updated " & CStr(changedCount) & " lookup field setting(s) on settings table [" & tableName & "]."
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "EnsureSettingsLookupFields", "Could not create or update macro-creator settings lookup fields." & vbCrLf & "Table: " & tableName & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

Private Sub RepairLauncherSettingDefaults(Optional ByVal forceDefault As Boolean = False)
    On Error GoTo ErrHandler
    Dim defaultRow As Variant
    For Each defaultRow In LauncherSettingDefinitions()
        UpsertSetting CStr(defaultRow(0)), CStr(defaultRow(1)), CStr(defaultRow(2)), CStr(defaultRow(3)), CStr(defaultRow(4)), CLng(defaultRow(5)), CBool(defaultRow(6)), CBool(defaultRow(7)), CBool(defaultRow(8)), CBool(defaultRow(9)), CBool(defaultRow(10)), CStr(defaultRow(11))
    Next defaultRow
    If forceDefault Then CoreAddOperationNote "Custom Macro Launcher setting metadata/defaults were refreshed; existing SettingValue choices were preserved."
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "RepairLauncherSettingDefaults", "Could not repair Custom Macro Launcher setting defaults while preserving existing SettingValue choices." & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

Private Sub PreflightSettingDefaultBatch()
    'Validates the complete Settings-row mutation target before the batched default repair begins.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim requiredFields As Variant
    Dim fieldName As Variant
    If TableCanOpen(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID)) = False Then Err.Raise vbObjectError + 5340, "PreflightSettingDefaultBatch", "The Settings table is not openable before batched default repair." & vbCrLf & "Open failure detail: " & TableOpenFailureDetail(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID))
    Set db = TablesWorkspaceCodeDb()
    Set rs = db.OpenRecordset(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID), dbOpenSnapshot)
    requiredFields = Array(FIELD_SETTING_KEY, FIELD_SETTING_VALUE, FIELD_MANAGED_OBJECT_FIELD_DEFAULT_VALUE, FIELD_MANAGED_OBJECT_FIELD_DATA_TYPE, FIELD_DESCRIPTION, FIELD_SETTING_GROUP, FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER, FIELD_ADVANCED, FIELD_ACTIVE, FIELD_RESTART_REQUIRED, FIELD_READ_ONLY, FIELD_REQUIRED, FIELD_VALIDATION_NOTES, FIELD_LAST_DEFAULTED_ON, FIELD_LAST_UPDATED_ON)
    For Each fieldName In requiredFields
        PumpSetupControlEvents
        If SharedDao_RecordsetHasFieldSafe(rs, CStr(fieldName)) = False Then Err.Raise vbObjectError + 5341, "PreflightSettingDefaultBatch", "The Settings table is missing a field required by the complete default-repair plan." & vbCrLf & "Missing field: " & CStr(fieldName)
    Next fieldName
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
    Err.Raise savedErrNumber, "PreflightSettingDefaultBatch", "Could not preflight the complete batched Settings default-repair target." & vbCrLf & "Settings table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID) & vbCrLf & "Database: " & TableSafeDatabaseName() & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub BeginSettingDefaultBatch()
    'Reuses one database handle and one dynaset for the complete insert-missing-default-row pass.
    On Error GoTo ErrHandler
    If mSettingBatchActive Then Err.Raise vbObjectError + 5342, "BeginSettingDefaultBatch", "A Settings default batch is already active."
    Set mSettingBatchDb = TablesWorkspaceCodeDb()
    Set mSettingBatchRs = mSettingBatchDb.OpenRecordset(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID), dbOpenDynaset)
    mSettingBatchActive = True
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    EndSettingDefaultBatch
    Err.Raise savedErrNumber, "BeginSettingDefaultBatch", "Could not start the batched Settings default repair." & vbCrLf & "Settings table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID) & vbCrLf & "Database: " & TableSafeDatabaseName() & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub EndSettingDefaultBatch()
    On Error Resume Next
    If Not mSettingBatchRs Is Nothing Then
        If mSettingBatchRs.EditMode <> 0 Then mSettingBatchRs.CancelUpdate
        mSettingBatchRs.Close
    End If
    Set mSettingBatchRs = Nothing
    Set mSettingBatchDb = Nothing
    mSettingBatchActive = False
End Sub

' Purpose: Repairs setting defaults for setup for supplied core setting definitions.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
'   suppliedCoreSettingDefinitions (Input Optional; Variant): Caller-supplied supplied core setting definitions data used as the authoritative definition input for this call, avoiding an unnecessary reload when
'   already available.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub RepairSettingDefaultsForSetup(Optional ByVal showMessage As Boolean = True, Optional ByVal suppliedCoreSettingDefinitions As Variant)
    RequireSetupDomainMutationAllowed SETUP_DOMAIN_A2_PHYSICAL_SCHEMA, "RepairSettingDefaultsForSetup"
    On Error GoTo ErrHandler
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim messageText As String
    Dim defaultRow As Variant
    Dim coreSettingDefinitions As Collection
    If isMissing(suppliedCoreSettingDefinitions) = False Then
        If IsObject(suppliedCoreSettingDefinitions) Then Set coreSettingDefinitions = suppliedCoreSettingDefinitions
    End If
    If coreSettingDefinitions Is Nothing Then Set coreSettingDefinitions = x__UTIL_ADC_DEFAULTS.GetCoreSettingDefinitions()
    PreflightSettingDefaultBatch
    BeginSettingDefaultBatch
    mSettingRepairSummaryActive = True
    mSettingRowsChecked = 0
    mSettingRowsExistingPreserved = 0
    mSettingRowsInserted = 0
    For Each defaultRow In coreSettingDefinitions
        UpsertSetting CStr(defaultRow(0)), CStr(defaultRow(1)), CStr(defaultRow(2)), CStr(defaultRow(3)), CStr(defaultRow(4)), CLng(defaultRow(5)), CBool(defaultRow(6)), CBool(defaultRow(7)), CBool(defaultRow(8)), CBool(defaultRow(9)), CBool(defaultRow(10)), CStr(defaultRow(11))
    Next defaultRow
    'LauncherSettingDefinitions remains the single launcher-setting definition authority; A2 materializes it in the same Settings write batch before any later launcher consumer.
    For Each defaultRow In LauncherSettingDefinitions()
        UpsertSetting CStr(defaultRow(0)), CStr(defaultRow(1)), CStr(defaultRow(2)), CStr(defaultRow(3)), CStr(defaultRow(4)), CLng(defaultRow(5)), CBool(defaultRow(6)), CBool(defaultRow(7)), CBool(defaultRow(8)), CBool(defaultRow(9)), CBool(defaultRow(10)), CStr(defaultRow(11))
    Next defaultRow
    AddTraceNote "A2 repaired the complete canonical Settings materialization contract, including foundational/naming/navigation/diagnostic/help/generated-form settings and the launcher-setting authority. Rows checked=" & CStr(mSettingRowsChecked) & "; rows inserted=" & CStr(mSettingRowsInserted) & "; existing rows preserved=" & CStr(mSettingRowsExistingPreserved) & "."
    EndSettingDefaultBatch
    mSettingRepairSummaryActive = False
    If showMessage Then ReportCoreActionSuccess "RepairSettingDefaultsForSetup", "DefaultRepair", "Canonical Settings defaults verified successfully.", "A2 materialized the complete foundational and launcher Settings catalogs required by later Setup consumers while preserving existing values.", "Repair canonical Settings defaults", ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID), "Table", "CodeDb", "DefaultRepair", True, "Repair Canonical Settings Defaults", vbInformation, True, "Tables", False
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    EndSettingDefaultBatch
    mSettingRepairSummaryActive = False
    messageText = "Could not materialize the complete A2 canonical Settings rows." & vbCrLf & "Settings table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
    LogTraceMessage TRACE_LEVEL_ERROR, "DefaultRepair", "RepairSettingDefaultsForSetup", "A2 canonical Settings materialization failed.", messageText, "Repair canonical Settings defaults", savedErrNumber, savedErrSource, savedErrDescription, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID), "Table", "CodeDb", "DefaultRepair"
    If showMessage Then
        CoreShowErrorMessage messageText, vbCritical, "Repair Canonical Settings Defaults - Error Detail", "RepairSettingDefaultsForSetup", "Tables"
        Exit Sub
    End If
    Err.Raise savedErrNumber, "RepairSettingDefaultsForSetup", messageText
End Sub

Private Sub UpsertSetting(ByVal SettingKey As String, ByVal SettingValue As String, ByVal descriptionText As String, Optional ByVal dataTypeText As String = SETTING_DATA_TYPE_STRING, Optional ByVal settingGroupText As String = "General", Optional ByVal sortOrderValue As Long = 0, Optional ByVal advancedValue As Boolean = False, Optional ByVal activeValue As Boolean = True, Optional ByVal restartRequiredValue As Boolean = False, Optional ByVal readOnlyValue As Boolean = False, Optional ByVal requiredValue As Boolean = False, Optional ByVal validationNotesText As String = vbNullString)
    'Intended use: insert a missing setting row with current defaults; preserve every field of an existing row.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim foundExisting As Boolean
    Dim nowValue As Date
    Dim storedSettingValue As Variant
    Dim defaultSourceName As String
    Dim ownsRecordset As Boolean
    If mSettingRepairSummaryActive Then mSettingRowsChecked = mSettingRowsChecked + 1
    nowValue = Now
    defaultSourceName = CanonicalSettingDefaultSourceName(SettingKey)
    If Len(SettingValue) = 0 Then
        storedSettingValue = Null
    Else
        storedSettingValue = SettingValue
    End If
    If mSettingBatchActive Then
        Set db = mSettingBatchDb
        Set rs = mSettingBatchRs
    Else
        Set db = TablesWorkspaceCodeDb()
        If TableCanOpen(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID)) = False Then Err.Raise vbObjectError + 5311, "UpsertSetting", "The settings table is not openable immediately before setting metadata repair." & vbCrLf & "SettingKey: " & SettingKey & vbCrLf & "CodeDb settings table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID) & vbCrLf & "Open failure detail: " & TableOpenFailureDetail(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID))
        Set rs = db.OpenRecordset(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID), dbOpenDynaset)
        ownsRecordset = True
    End If
    rs.FindFirst "[" & FIELD_SETTING_KEY & "]=" & SharedDao_SqlQuote(SettingKey)
    foundExisting = Not rs.NoMatch
    If foundExisting = False Then
        rs.AddNew
        SetRecordsetFieldIfExists rs, FIELD_SETTING_KEY, SettingKey
        SetRecordsetFieldIfExists rs, FIELD_SETTING_VALUE, storedSettingValue, defaultSourceName
        SetRecordsetFieldIfExists rs, FIELD_MANAGED_OBJECT_FIELD_DEFAULT_VALUE, storedSettingValue, defaultSourceName
        SetRecordsetFieldIfExists rs, FIELD_MANAGED_OBJECT_FIELD_DATA_TYPE, dataTypeText
        SetRecordsetFieldIfExists rs, FIELD_DESCRIPTION, descriptionText
        SetRecordsetFieldIfExists rs, FIELD_SETTING_GROUP, settingGroupText
        SetRecordsetFieldIfExists rs, FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER, sortOrderValue
        SetRecordsetFieldIfExists rs, FIELD_ADVANCED, advancedValue
        SetRecordsetFieldIfExists rs, FIELD_ACTIVE, activeValue
        SetRecordsetFieldIfExists rs, FIELD_RESTART_REQUIRED, restartRequiredValue
        SetRecordsetFieldIfExists rs, FIELD_READ_ONLY, readOnlyValue
        SetRecordsetFieldIfExists rs, FIELD_REQUIRED, requiredValue
        SetRecordsetFieldIfExists rs, FIELD_VALIDATION_NOTES, validationNotesText
        SetRecordsetFieldIfExists rs, FIELD_LAST_DEFAULTED_ON, nowValue
        SetRecordsetFieldIfExists rs, FIELD_LAST_UPDATED_ON, nowValue
        rs.Update
        If mSettingRepairSummaryActive Then
            mSettingRowsInserted = mSettingRowsInserted + 1
        Else
            CoreAddOperationNote "Inserted default setting [" & SettingKey & "] into [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID) & "]."
        End If
    Else
        If mSettingRepairSummaryActive Then
            mSettingRowsExistingPreserved = mSettingRowsExistingPreserved + 1
        Else
            AddTraceNote "Preserved existing Settings row exactly: " & SettingKey
        End If
    End If
    On Error Resume Next
    If ownsRecordset Then
        If Not rs Is Nothing Then rs.Close
    End If
    Exit Sub
ErrHandler:
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.EditMode <> 0 Then rs.CancelUpdate
    End If
    If ownsRecordset Then
        If Not rs Is Nothing Then rs.Close
    End If
    On Error GoTo 0
    Err.Raise errNumber, "UpsertSetting", "Could not insert/update macro-creator default setting metadata." & vbCrLf & "SettingKey: " & SettingKey & vbCrLf & "CodeDb settings table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID) & vbCrLf & "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB & vbCrLf & "Database: " & TableSafeDatabaseName() & vbCrLf & "Err.Number: " & CStr(errNumber) & vbCrLf & "Err.Source: " & errSource & vbCrLf & "Err.Description: " & errDescription & TableReadinessDiagnosticsForMessage()
End Sub

' ===============================================================================================================================================================================
' SAVED QUERY / FORM-SOURCE FOUNDATION PROCEDURES
' ===============================================================================================================================================================================

Private Function QueryRepairDb() As DAO.Database
    If mQueryRepairBatchActive Then
        If mQueryRepairBatchDb Is Nothing Then Err.Raise vbObjectError + 9050, "QueryRepairDb", "The active saved-query repair batch has no database context."
        Set QueryRepairDb = mQueryRepairBatchDb
    Else
        Set QueryRepairDb = TablesWorkspaceCodeDb()
    End If
End Function

Private Sub BeginQueryRepairBatch()
    If mQueryRepairBatchActive Then Err.Raise vbObjectError + 9048, "BeginQueryRepairBatch", "A saved-query repair batch is already active."
    Set mQueryRepairBatchDb = TablesWorkspaceCodeDb()
    mQueryMismatchDiagnosticCount = 0
    mQueryRepairBatchChecked = 0
    mQueryRepairBatchCreated = 0
    mQueryRepairBatchRefreshed = 0
    mQueryRepairBatchUnchanged = 0
    Set mQueryRepairBatchVerifiedEvidence = CreateObject("Scripting.Dictionary")
    mQueryRepairBatchVerifiedEvidence.CompareMode = vbTextCompare
    mQueryRepairBatchStarted = Timer
    mQueryRepairBatchActive = True
End Sub

Private Function EndQueryRepairBatch() As String
    Dim summaryText As String
    If mQueryRepairBatchActive = False Then Err.Raise vbObjectError + 9049, "EndQueryRepairBatch", "No saved-query repair batch is active."
    summaryText = "Saved-query batch summary: checked=" & CStr(mQueryRepairBatchChecked) & "; created=" & CStr(mQueryRepairBatchCreated) & "; refreshed=" & CStr(mQueryRepairBatchRefreshed) & "; unchanged=" & CStr(mQueryRepairBatchUnchanged) & "; elapsed=" & FormattedElapsedSeconds(ElapsedSecondsFromTimer(mQueryRepairBatchStarted)) & ". One database context was reused and per-query progress and ordinary success notes were suppressed; errors, mismatch diagnostics, verification, rollback tracking, and Navigation Pane warnings remain fail-closed."
    mQueryRepairBatchActive = False
    Set mQueryRepairBatchDb = Nothing
    Set mQueryRepairBatchVerifiedEvidence = Nothing
    mQueryRepairBatchStarted = 0
    AddTraceNote summaryText
    EndQueryRepairBatch = summaryText
End Function

Private Sub AbortQueryRepairBatch()
    mQueryRepairBatchActive = False
    Set mQueryRepairBatchDb = Nothing
    Set mQueryRepairBatchVerifiedEvidence = Nothing
    mQueryRepairBatchChecked = 0
    mQueryRepairBatchCreated = 0
    mQueryRepairBatchRefreshed = 0
    mQueryRepairBatchUnchanged = 0
    mQueryRepairBatchStarted = 0
End Sub

Private Sub CaptureQueryRepairBatchVerifiedEvidence(ByVal queryName As String, ByVal qdf As DAO.queryDef)
    If mQueryRepairBatchActive = False Then Exit Sub
    If mQueryRepairBatchVerifiedEvidence Is Nothing Then Err.Raise vbObjectError + 9052, "CaptureQueryRepairBatchVerifiedEvidence", "The active saved-query repair batch has no verified QueryDef evidence index."
    If qdf Is Nothing Then Err.Raise vbObjectError + 9053, "CaptureQueryRepairBatchVerifiedEvidence", "The verified saved query has no QueryDef object." & vbCrLf & "Query: " & queryName
    mQueryRepairBatchVerifiedEvidence.item(queryName) = Array(CStr(qdf.sql), CLng(qdf.Type), CStr(qdf.Name))
End Sub

Private Function QueryRepairBatchVerifiedEvidenceSnapshot() As Object
    Dim result As Object
    Dim queryName As Variant
    Dim evidence As Variant
    ' INTENTIONAL OPERATION SNAPSHOT: result captures the one authoritative post-repair QueryDef state already verified by this A3 batch; validity ends with this A3 operation.
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    If Not mQueryRepairBatchVerifiedEvidence Is Nothing Then
        For Each queryName In mQueryRepairBatchVerifiedEvidence.keys
            evidence = mQueryRepairBatchVerifiedEvidence.item(queryName)
            result.Add CStr(queryName), Array(CStr(evidence(0)), CLng(evidence(1)), CStr(evidence(2)))
        Next queryName
    End If
    Set QueryRepairBatchVerifiedEvidenceSnapshot = result
End Function

Private Sub PreflightQuerySql(ByVal queryName As String, ByVal querySql As String)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim qdf As DAO.queryDef
    Dim rs As DAO.Recordset
    Set db = QueryRepairDb()
    Set qdf = db.CreateQueryDef(vbNullString, querySql)
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not qdf Is Nothing Then qdf.Close
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
    If Not qdf Is Nothing Then qdf.Close
    On Error GoTo 0
    Err.Raise savedErrNumber, "PreflightQuerySql", "Replacement SQL could not be opened before changing the saved query." & vbCrLf & "Query: " & queryName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription & vbCrLf & "SQL:" & vbCrLf & Left$(querySql, 3500)
End Sub

' TABLES owns A3 mismatch diagnostics only; COMMON owns ACE-canonical SQL equivalence for every saved-query family.
Private Function QuerySqlMatches(ByVal db As DAO.Database, ByVal existingSql As String, ByVal requestedSql As String, ByRef mismatchDetail As String) As Boolean
    QuerySqlMatches = QuerySqlEquivalent(db, existingSql, requestedSql)
    If QuerySqlMatches = False Then mismatchDetail = TextMismatchDetail(NormalizedQuerySqlForComparison(db, existingSql), NormalizedQuerySqlForComparison(db, requestedSql))
End Function

Private Sub VerifyQueryOpenable(ByVal queryName As String, ByVal querySql As String)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim qdf As DAO.queryDef
    Dim rs As DAO.Recordset
    Set db = QueryRepairDb()
    Set qdf = db.QueryDefs(queryName)
    'Opening the snapshot is the safest compile check.  Do not call qdf.Parameters.Refresh before opening: DAO can raise false 3067 errors for valid constant/lookup SELECT queries.  If opening fails, the error
    'handler tries to collect parameter names only as diagnostic detail.
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Exit Sub
ErrHandler:
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim unresolvedParameters As String
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    unresolvedParameters = QueryParameterListForMessage(queryName)
    If Len(unresolvedParameters) > 0 Then
        errDescription = errDescription & vbCrLf & "Unresolved parameter(s): " & unresolvedParameters
    End If
    Err.Raise errNumber, "VerifyQueryOpenable", "Saved query could not be opened/compiled." & vbCrLf & "Query: " & queryName & vbCrLf & "Err.Number: " & CStr(errNumber) & vbCrLf & "Err.Source: " & errSource & vbCrLf & "Err.Description: " & errDescription & vbCrLf & "SQL:" & vbCrLf & Left$(querySql, 3500)
End Sub

Private Sub ValidateQueryFieldOrderSql(ByVal queryName As String, ByVal querySql As String)
    On Error GoTo ErrHandler
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.ignoreCase = True
    regex.multiLine = True
    regex.pattern = "(SELECT\s+(DISTINCTROW\s+|DISTINCT\s+|TOP\s+\d+(\s+PERCENT)?\s+)*|,\s*)(\*|\[[^\]]+\]\.\*|[A-Z_][A-Z0-9_]*\.\*)"
    If regex.Test(querySql) Then Err.Raise vbObjectError + 5935, "ValidateQueryFieldOrderSql", "Managed saved queries must use an explicit, intentionally ordered output-field list. Wildcard SELECT output is not allowed."
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    If originalErrNumber = vbObjectError + 5935 Then Err.Raise originalErrNumber, "ValidateQueryFieldOrderSql", originalErrDescription & vbCrLf & "Query: " & queryName
    Err.Raise originalErrNumber, "ValidateQueryFieldOrderSql", "Could not validate the managed query output-field order." & vbCrLf & "Query: " & queryName & vbCrLf & "Description: " & originalErrDescription
End Sub

Private Function TextMismatchDetail(ByVal existingText As String, ByVal requestedText As String) As String
    Dim differencePosition As Long
    differencePosition = FirstTextDifferencePosition(existingText, requestedText)
    TextMismatchDetail = "Existing length/hash: " & CStr(Len(existingText)) & "/" & AdlerTextHash(existingText)
    TextMismatchDetail = TextMismatchDetail & vbCrLf & "Requested length/hash: " & CStr(Len(requestedText)) & "/" & AdlerTextHash(requestedText)
    TextMismatchDetail = TextMismatchDetail & vbCrLf & "First differing character position: " & CStr(differencePosition)
    TextMismatchDetail = TextMismatchDetail & vbCrLf & "Existing excerpt: " & TextComparisonExcerpt(existingText, differencePosition)
    TextMismatchDetail = TextMismatchDetail & vbCrLf & "Requested excerpt: " & TextComparisonExcerpt(requestedText, differencePosition)
End Function

' Purpose: Repairs one managed saved query after normalized no-op detection, preflights only changed SQL, and restores the prior query locally if replacement verification fails.
' Public role: Public module entry point. This procedure is listed in the module header as an intentional externally callable ADC entry point and may be invoked directly by Access/VBA callers.
' Parameters:
'   queryName (Input; String): Physical Access query name identifying the saved query the routine creates, validates, reads, or updates.
'   querySql (Input; String): SQL text or SQL-derived value used to define, validate, compare, or execute the query SQL operation.
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub RepairOrReplaceQueryForSetup(ByVal queryName As String, ByVal querySql As String, Optional ByVal showMessage As Boolean = True)
    RequireCatalogMutationRecoveryReady "RepairOrReplaceQueryForSetup"
    'Repairs one managed saved query after normalized no-op detection, preflights only changed SQL, and restores the prior query locally if replacement verification fails.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim qdf As DAO.queryDef
    Dim currentAction As String
    Dim targetExisted As Boolean
    Dim previousSql As String
    Dim replacementApplied As Boolean
    Dim createdNewQuery As Boolean
    Dim queryChanged As Boolean
    Dim successMessage As String
    Dim queryMismatchDetail As String
    If mQueryRepairBatchActive = False Then mQueryMismatchDiagnosticCount = 0
    Set db = QueryRepairDb()
    If mQueryRepairBatchActive = False Then CorePulseProgress "Saved query: " & queryName
    If mQueryRepairBatchActive Then mQueryRepairBatchChecked = mQueryRepairBatchChecked + 1
    currentAction = "validate explicit field order"
    ValidateQueryFieldOrderSql queryName, querySql
    targetExisted = SharedDao_QueryExists(db, queryName)
    If targetExisted Then
        Set qdf = db.QueryDefs(queryName)
        previousSql = qdf.sql
        currentAction = "compare normalized SQL"
        queryChanged = Not QuerySqlMatches(db, previousSql, querySql, queryMismatchDetail)
        If queryChanged Then
            If mQueryMismatchDiagnosticCount < 5 Then
                mQueryMismatchDiagnosticCount = mQueryMismatchDiagnosticCount + 1
                LogTraceMessage TRACE_LEVEL_DETAIL, "QueryComparison", "RepairOrReplaceQueryForSetup", "Managed query SQL mismatch diagnostic.", queryMismatchDetail, "Compare normalized SQL", 0, vbNullString, vbNullString, queryName, "Query", "CodeDb", "QueryComparison"
            End If
            currentAction = "preflight replacement SQL"
            PreflightQuerySql queryName, querySql
            currentAction = "prepare affected open query"
            If SharedAccess_PrepareOpenObjectForDesignChange(acQuery, queryName, "replacing saved-query SQL for [" & queryName & "]") = False Then
                MarkSystemCommandCancelled
                Err.Raise vbObjectError + 8990, "RepairOrReplaceQueryForSetup", "Saved-query repair was cancelled because the affected open query was not approved for save and close."
            End If
            Set qdf = db.QueryDefs(queryName)
            TrackSetupQueryRollback queryName, True, previousSql
            currentAction = "refresh SQL"
            qdf.sql = querySql
            replacementApplied = True
            If mQueryRepairBatchActive Then
                mQueryRepairBatchRefreshed = mQueryRepairBatchRefreshed + 1
            Else
                CoreAddOperationNote "Refreshed query [" & queryName & "]."
            End If
        Else
            currentAction = "verify unchanged query opens/compiles"
            If mQueryRepairBatchActive Then mQueryRepairBatchUnchanged = mQueryRepairBatchUnchanged + 1
        End If
    Else
        currentAction = "preflight replacement SQL"
        PreflightQuerySql queryName, querySql
        TrackSetupQueryRollback queryName, False, vbNullString
        currentAction = "create query"
        Set qdf = db.CreateQueryDef(queryName, querySql)
        replacementApplied = True
        createdNewQuery = True
        queryChanged = True
        If mQueryRepairBatchActive Then
            mQueryRepairBatchCreated = mQueryRepairBatchCreated + 1
        Else
            CoreAddOperationNote "Created query [" & queryName & "]."
        End If
    End If
    currentAction = "verify query opens/compiles"
    VerifyQueryOpenable queryName, IIf(queryChanged, querySql, previousSql)
    CaptureQueryRepairBatchVerifiedEvidence queryName, qdf
    replacementApplied = False
    ReconcileStandaloneManagedQueryDescriptionIfAllowed queryName
    currentAction = "defer Navigation Pane placement to Phase I"
    successMessage = DEFAULT_UTILITY_LONG_NAME & " query repair complete." & vbCrLf & vbCrLf & "Query: " & queryName & vbCrLf & "SQL changed: " & CStr(queryChanged) & vbCrLf & "Navigation Pane placement deferred to Phase I: True"
    If showMessage Then ReportCoreActionSuccess "RepairOrReplaceQueryForSetup", "QueryRepair", DEFAULT_UTILITY_SHORT_NAME & " query repair completed successfully.", successMessage, "Repair or replace query", queryName, "Query", "CodeDb", "QueryRepair", True, "Repair Or Replace " & DEFAULT_UTILITY_LONG_NAME & " Query", vbInformation, True, "Tables", False, visibleStatusMessage:="Query repair completed successfully."
    Exit Sub
ErrHandler:
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim messageText As String
    Dim restoreFailureText As String
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    On Error Resume Next
    If replacementApplied Then
        If createdNewQuery Then
            db.QueryDefs.Delete queryName
            If Err.Number <> 0 Then restoreFailureText = "Could not delete the failed newly created query: " & Err.Description
        Else
            Err.Clear
            Set qdf = db.QueryDefs(queryName)
            qdf.sql = previousSql
            If Err.Number <> 0 Then restoreFailureText = "Could not restore the previous saved-query SQL: " & Err.Description
        End If
        db.QueryDefs.Refresh
    End If
    On Error GoTo 0
    messageText = "Could not create, refresh, or verify saved query." & vbCrLf & vbCrLf & "Procedure: RepairOrReplaceQueryForSetup" & vbCrLf & "Query: " & queryName & vbCrLf & "Action: " & IIf(Len(currentAction) = 0, "(not recorded)", currentAction) & vbCrLf & "Err.Number: " & CStr(errNumber) & vbCrLf & "Err.Source: " & errSource & vbCrLf & "Err.Description:" & vbCrLf & errDescription
    If Len(restoreFailureText) > 0 Then messageText = messageText & vbCrLf & "Local rollback failure: " & restoreFailureText
    messageText = messageText & vbCrLf & "SQL:" & vbCrLf & Left$(querySql, 3500) & TableReadinessDiagnosticsForMessage() & TraceBackendDiagnosticsForMessage()
    LogTraceMessage TRACE_LEVEL_ERROR, "QueryRepair", "RepairOrReplaceQueryForSetup", "Saved query repair failed.", messageText, IIf(Len(currentAction) = 0, "Repair saved query", currentAction), errNumber, errSource, errDescription, queryName, "Query", "CodeDb", "ReviewFoundation"
    If showMessage Then
        CoreShowErrorMessage messageText, vbCritical, "Repair Or Replace Saved Query", "RepairOrReplaceQueryForSetup", "Tables"
        Exit Sub
    End If
    Err.Raise errNumber, "RepairOrReplaceQueryForSetup", messageText
End Sub


Private Sub RepairNonFormSavedQueryDomainsForSetup(ByRef verifiedQueryEvidenceByName As Object)
    RequireCatalogMutationRecoveryReady "RepairNonFormSavedQueryDomainsForSetup"
    'TABLES is the sole owner of the non-form saved-query families assigned to A3 and their physical QueryDef repair; persistent editable lookup vocabularies are A2 tables and are excluded from this QueryDef phase.
    'Specialized Reference/Add-In, Module Guard, and catalog query families remain with their owning domains.
    On Error GoTo ErrHandler
    Dim queryBatchDetail As String
    Dim formDesignSourceNames As Collection
    BeginQueryRepairBatch
    CorePulseProgress "Review foundation: " & "A3.1 repair core and support saved-query families"
    EnsureCoreFormSupportQueriesForSetup
    RepairDuplicateReviewQueryForSetup
    EnsureCustomMacroLauncherReviewQueryForSetup
    CorePulseProgress "Review foundation: " & "A3.2 repair FormDesign saved-query families"
    ResolveFormDesignBuildFamilies formDesignSourceNames
    EnsureFormDesignerCoreSourceQueries formDesignSourceNames
    CorePulseProgress "Review foundation: " & "A3.3 repair Custom Macro Launcher source saved queries"
    EnsureFormDesignerLauncherSourceQueries
    Set verifiedQueryEvidenceByName = QueryRepairBatchVerifiedEvidenceSnapshot()
    queryBatchDetail = EndQueryRepairBatch()
    AddTraceNote "A3 complete non-form saved-query batch completed under TABLES authority. A3.1=core/support; A3.2=FormDesign; A3.3=Custom Macro Launcher sources. " & queryBatchDetail
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    AbortQueryRepairBatch
    Err.Raise savedErrNumber, "RepairNonFormSavedQueryDomainsForSetup", "Could not repair the complete A3 non-form saved-query foundation." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub



Private Sub EnsureCoreFormSupportQueriesForSetup()
    CorePulseProgress "Review foundation: " & "refresh configuration form source query"
    RepairOrReplaceQueryForSetup ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Config_Management"), ConfigFormSql(), False
    CorePulseProgress "Review foundation: " & "refresh macro report source query"
    RepairOrReplaceQueryForSetup ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Macro_Report_Source"), MacroReportSourceSql(), False
    CorePulseProgress "Review foundation: " & "refresh candidate review form source query"
    RepairOrReplaceQueryForSetup ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Config_CandidateReview"), CandidatesFormSql(), False
    CorePulseProgress "Review foundation: " & "refresh settings form source query"
    RepairOrReplaceQueryForSetup ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Settings_Resolved"), SettingsFormSql(), False
    CorePulseProgress "Review foundation: " & "refresh health review query"
    RepairOrReplaceQueryForSetup ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_System_HealthReview"), HealthReviewFormSql(), False
    CorePulseProgress "Review foundation: " & "refresh target-resolution review query"
    RepairOrReplaceQueryForSetup ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Config_TargetResolutionAudit"), TargetResolutionFormSql(), False
    CorePulseProgress "Review foundation: " & "refresh stale-candidate review query"
    RepairOrReplaceQueryForSetup ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Config_StaleCandidates"), StaleCandidatesFormSql(), False
    CorePulseProgress "Review foundation: " & "refresh generation-status review query"
    RepairOrReplaceQueryForSetup ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Config_GenerationStatus"), GenerationStatusFormSql(), False
    CorePulseProgress "Review foundation: " & "refresh manual-review query"
    RepairOrReplaceQueryForSetup ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Config_ManualReview"), ManualReviewFormSql(), False
End Sub



Private Sub EnsureFormDesignerCoreSourceQueries(ByVal queryNames As Collection)
    'A3.2 source rows: all FormDesign source queries except the Custom Macro Launcher sources owned by A3.3.
    Dim queryName As Variant
    For Each queryName In queryNames
        If StrComp(CStr(queryName), ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_LauncherControls_BackfillSource"), vbBinaryCompare) <> 0 And StrComp(CStr(queryName), ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_LauncherControls_ReviewSource"), vbBinaryCompare) <> 0 Then
            RepairOrReplaceQueryForSetup CStr(queryName), FormDesignQuerySql(CStr(queryName)), False
        End If
    Next queryName
End Sub



Private Sub EnsureFormDesignerLauncherSourceQueries()
    'A3.3 Custom Macro Launcher sources remain part of the same A3 repair batch and canonical query registry.
    RepairOrReplaceQueryForSetup ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_LauncherControls_BackfillSource"), FormDesignQuerySql(ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_LauncherControls_BackfillSource")), False
    RepairOrReplaceQueryForSetup ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_LauncherControls_ReviewSource"), FormDesignQuerySql(ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_LauncherControls_ReviewSource")), False
End Sub



' Purpose: Repairs non form query foundation for setup.
' Public role: Public module entry point. This procedure is listed in the module header as an intentional externally callable ADC entry point and may be invoked directly by Access/VBA callers.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: False.
'   repairDetailsForMessage (Output Optional; String): Receives repair details for message used for caller diagnostics, verification output, or user-facing reporting. Default: vbNullString.
'   applyStandaloneDescriptions (Input Optional; Boolean): True completes the A3 managed query Descriptions when this public coordinator runs outside numbered Setup; False lets a broader caller own a later
'   complete Description pass. Default: True.
' Returns: Boolean: True on the function's affirmative/success branch for repair non form query foundation for setup; False on its explicit negative, unavailable, rejected, or handled-failure branch.
'   not-applicable path.
Public Function RepairNonFormQueryFoundationForSetup(Optional ByVal showMessage As Boolean = False, Optional ByRef repairDetailsForMessage As String = vbNullString, Optional ByVal applyStandaloneDescriptions As Boolean = True) As Boolean
    RequireSetupDomainMutationAllowed SETUP_DOMAIN_A3_SAVED_QUERY_FOUNDATION, "RepairNonFormQueryFoundationForSetup"
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim phaseStarted As Single
    Dim messageText As String
    Dim QueryFingerprint As String
    Dim queryFingerprintFailure As String
    Dim verifiedCount As Long
    Dim verifiedQueryEvidenceByName As Object
    Dim queryRows As Collection
    Dim requiredNames As Collection
    Dim refreshedSnapshot As FoundationDomainSnapshot
    phaseStarted = Timer
    repairDetailsForMessage = vbNullString
    Set queryRows = FoundationRequiredNonFormQueryFamilyRows()
    Set requiredNames = FoundationRequiredNonFormQueryNames(queryRows)
    CoreUpdateSetupProgress "Setup A3 - repairing all non-form saved-query foundations."
    RepairNonFormSavedQueryDomainsForSetup verifiedQueryEvidenceByName
    Set db = TablesWorkspaceCodeDb()
    QueryFingerprint = NonFormQueryFoundationStateTokenFromVerifiedEvidence(verifiedQueryEvidenceByName, queryFingerprintFailure, verifiedCount, requiredNames)
    If Len(QueryFingerprint) = 0 Then Err.Raise vbObjectError + 9045, "RepairNonFormQueryFoundationForSetup", "A3 repair completed, but the complete non-form saved-query fingerprint could not be calculated. " & queryFingerprintFailure
    Call EnsureDaoPropertyValue(db, FOUNDATION_NONFORM_QUERY_FINGERPRINT_PROPERTY, dbText, QueryFingerprint, "RepairNonFormQueryFoundationForSetup")
    If CalculateFoundationQuerySnapshot(refreshedSnapshot, queryRows, requiredNames, verifiedQueryEvidenceByName) = False Then Err.Raise vbObjectError + 9045, "RepairNonFormQueryFoundationForSetup", "A3 repair completed, but the detailed FormDesign query-domain fingerprints could not be calculated. " & refreshedSnapshot.failureDetail
    AddTraceNote "A3 reused one authoritative post-repair QueryDef evidence set for the complete non-form token and all query-family fingerprints; no second QueryDefs.Refresh or complete QueryDefs enumeration was performed."
    PersistFoundationQuerySnapshot refreshedSnapshot
    If applyStandaloneDescriptions Then ReconcileStandaloneNonFormQueryDescriptions
    repairDetailsForMessage = vbCrLf & vbCrLf & "Complete non-form saved-query foundation path: Repaired/verified" & vbCrLf & "A3.1 core/support queries: completed" & vbCrLf & "A3.2 FormDesign queries: completed" & vbCrLf & "A3.3 Custom Macro Launcher source queries: completed" & vbCrLf & "Required non-form saved queries verified: " & CStr(verifiedCount) & vbCrLf & FoundationDomainSummary(refreshedSnapshot) & vbCrLf & "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseStarted))
    RepairNonFormQueryFoundationForSetup = True
    ReportCoreActionSuccess "RepairNonFormQueryFoundationForSetup", "ReviewFoundation", "Complete non-form saved-query foundation verified successfully.", "A3 owns the non-form saved-query families assigned to A3; specialized Reference/Add-In, Module Guard, and catalog query families remain with their owning domains." & repairDetailsForMessage, "Verify complete non-form saved-query foundation", vbNullString, "FoundationNonFormQuerySet", "CodeDb", "ReviewFoundation", showMessage, "Repair " & DEFAULT_UTILITY_LONG_NAME & " Queries", vbInformation, True, "Tables", False
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    RepairNonFormQueryFoundationForSetup = False
    repairDetailsForMessage = vbCrLf & vbCrLf & "Complete non-form saved-query foundation repair failed." & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
    messageText = "Could not verify or repair the complete non-form saved-query foundation." & repairDetailsForMessage & TableDiagnosticsLocationForMessage()
    LogTraceMessage TRACE_LEVEL_ERROR, "ReviewFoundation", "RepairNonFormQueryFoundationForSetup", "Complete non-form saved-query foundation repair failed.", messageText, "Repair complete non-form saved-query foundation", savedErrNumber, savedErrSource, savedErrDescription, vbNullString, "FoundationNonFormQuerySet", "CodeDb", "ReviewFoundation"
    If showMessage Then
        CoreShowErrorMessage messageText, vbCritical, DEFAULT_TABLE_MAINTENANCE_MESSAGE_TITLE, "RepairNonFormQueryFoundationForSetup", "Tables"
        Exit Function
    End If
    Err.Raise savedErrNumber, "RepairNonFormQueryFoundationForSetup", messageText
End Function

' Purpose: Checks whether non form query foundation already satisfies the exact state required by the current Setup run so the owning phase can determine whether repair is necessary.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   detailText (Output; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
' Returns: Boolean: True when non form query foundation matches the exact state required by the current Setup contract; False when the owning phase still needs repair, refresh, or reconciliation.
Public Function NonFormQueryFoundationCurrentForSetup(ByRef detailText As String) As Boolean
    On Error GoTo NotCurrent
    Dim db As DAO.Database
    Dim currentToken As String
    Dim storedToken As String
    Dim propertyExists As Boolean
    detailText = vbNullString
    Set db = TablesWorkspaceCodeDb()
    storedToken = DatabaseTextPropertyValue(db, FOUNDATION_NONFORM_QUERY_FINGERPRINT_PROPERTY, propertyExists)
    If propertyExists = False Or Len(storedToken) = 0 Then
        detailText = "A3 non-form saved-query fingerprint=unseeded; physical comparison skipped because required persisted proof is unavailable."
        Exit Function
    End If
    currentToken = NonFormQueryFoundationStateToken(detailText)
    If Len(currentToken) = 0 Then Exit Function
    NonFormQueryFoundationCurrentForSetup = (StrComp(storedToken, currentToken, vbBinaryCompare) = 0)
    detailText = "A3 non-form saved-query fingerprint=" & IIf(NonFormQueryFoundationCurrentForSetup, "matched", "changed") & "."
    Exit Function
NotCurrent:
    detailText = "A3 current-state verification failed. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
End Function

Private Function NonFormQueryFoundationStateTokenFromVerifiedEvidence(ByVal verifiedQueryEvidenceByName As Object, ByRef failureDetail As String, Optional ByRef verifiedCount As Long = 0, Optional ByVal suppliedRequiredNames As Variant) As String
    On Error GoTo ErrHandler
    Dim requiredNames As Collection
    Dim item As Variant
    Dim queryName As String
    Dim payloadText As String
    Dim evidence As Variant
    failureDetail = vbNullString
    verifiedCount = 0
    If verifiedQueryEvidenceByName Is Nothing Then failureDetail = "The A3 saved-query repair batch returned no verified QueryDef evidence index.": Exit Function
    If isMissing(suppliedRequiredNames) = False Then
        If IsObject(suppliedRequiredNames) Then Set requiredNames = suppliedRequiredNames
    End If
    If requiredNames Is Nothing Then Set requiredNames = FoundationRequiredNonFormQueryNames()
    payloadText = FOUNDATION_NONFORM_QUERY_FINGERPRINT_VERSION & "|"
    For Each item In requiredNames
        queryName = CStr(item)
        If verifiedQueryEvidenceByName.Exists(queryName) = False Then failureDetail = "The A3 saved-query repair batch did not retain verified QueryDef evidence for: " & queryName: Exit Function
        evidence = verifiedQueryEvidenceByName.item(queryName)
        verifiedCount = verifiedCount + 1
        payloadText = payloadText & FingerprintTokenPart(queryName) & FingerprintTokenPart(CStr(evidence(0)))
    Next item
    NonFormQueryFoundationStateTokenFromVerifiedEvidence = FingerprintFromPayload(payloadText)
    Exit Function
ErrHandler:
    failureDetail = "Could not calculate the A3 non-form saved-query fingerprint from the repair batch evidence. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
End Function

Private Function NonFormQueryFoundationStateToken(ByRef failureDetail As String, Optional ByRef verifiedCount As Long = 0) As String
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim requiredNames As Collection
    Dim item As Variant
    Dim qdf As DAO.queryDef
    Dim payloadText As String
    failureDetail = vbNullString
    verifiedCount = 0
    Set db = TablesWorkspaceCodeDb()
    db.QueryDefs.Refresh
    Set requiredNames = FoundationRequiredNonFormQueryNames()
    payloadText = FOUNDATION_NONFORM_QUERY_FINGERPRINT_VERSION & "|"
    For Each item In requiredNames
        PumpSetupControlEvents
        If SharedDao_QueryExists(db, CStr(item)) = False Then failureDetail = "Missing non-form saved query: " & CStr(item): Exit Function
        Set qdf = db.QueryDefs(CStr(item))
        verifiedCount = verifiedCount + 1
        payloadText = payloadText & FingerprintTokenPart(CStr(item)) & FingerprintTokenPart(qdf.sql)
    Next item
    NonFormQueryFoundationStateToken = FingerprintFromPayload(payloadText)
    Exit Function
ErrHandler:
    failureDetail = "Could not calculate the A3 non-form saved-query fingerprint. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
End Function

Private Function FoundationQueryDefinitionSnapshotFromVerifiedEvidence(ByVal verifiedQueryEvidenceByName As Object, ByVal queryNames As Collection, ByRef failureDetail As String) As Object
    Dim cache As Object
    Dim queryItem As Variant
    Dim evidence As Variant
    Set cache = CreateObject("Scripting.Dictionary")
    cache.CompareMode = vbTextCompare
    If verifiedQueryEvidenceByName Is Nothing Then failureDetail = "The authoritative A3 QueryDef evidence set is unavailable.": Exit Function
    For Each queryItem In queryNames
        If verifiedQueryEvidenceByName.Exists(CStr(queryItem)) = False Then failureDetail = "Required foundation query is missing from authoritative A3 evidence: " & CStr(queryItem): Exit Function
        evidence = verifiedQueryEvidenceByName.item(CStr(queryItem))
        cache(CStr(queryItem)) = FingerprintTokenPart(CStr(evidence(2))) & FingerprintTokenPart(CStr(CLng(evidence(1)))) & FingerprintTokenPart(CStr(evidence(0)))
    Next queryItem
    Set FoundationQueryDefinitionSnapshotFromVerifiedEvidence = cache
End Function

Private Function FoundationQueryDefinitionSnapshot(ByVal db As DAO.Database, ByVal queryNames As Collection, ByRef failureDetail As String) As Object
    On Error GoTo ErrHandler
    Dim cache As Object
    Dim requiredNames As Object
    Dim queryItem As Variant
    Dim qdf As DAO.queryDef
    Dim resultText As String
    ' INTENTIONAL KEYED LOOKUP: cache provides keyed membership/lookup over source rows or state already read by FoundationQueryDefinitionSnapshot; it is a derived access path, not a second authority.
    ' INTENTIONAL OPERATION SNAPSHOT: requiredNames captures one coherent source state for this invocation of FoundationQueryDefinitionSnapshot; validity ends with the operation rather than becoming a live data
    ' authority.
    Set cache = CreateObject("Scripting.Dictionary")
    cache.CompareMode = vbTextCompare
    Set requiredNames = CreateObject("Scripting.Dictionary")
    requiredNames.CompareMode = vbTextCompare
    For Each queryItem In queryNames
        PumpSetupControlEvents
        If requiredNames.Exists(CStr(queryItem)) = False Then requiredNames.Add CStr(queryItem), True
    Next queryItem
    For Each qdf In db.QueryDefs
        PumpSetupControlEvents
        If requiredNames.Exists(qdf.Name) Then
            resultText = FingerprintTokenPart(qdf.Name)
            resultText = resultText & FingerprintTokenPart(CStr(qdf.Type))
            resultText = resultText & FingerprintTokenPart(qdf.sql)
            cache(qdf.Name) = resultText
        End If
    Next qdf
    For Each queryItem In queryNames
        PumpSetupControlEvents
        If cache.Exists(CStr(queryItem)) = False Then
            failureDetail = "Required foundation query is missing: " & CStr(queryItem)
            Set FoundationQueryDefinitionSnapshot = Nothing
            Exit Function
        End If
    Next queryItem
    Set FoundationQueryDefinitionSnapshot = cache
    Exit Function
ErrHandler:
    failureDetail = "Could not preload the required QueryDefs snapshot. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    Set FoundationQueryDefinitionSnapshot = Nothing
End Function

Private Function FoundationQueryDefinitionTextFromSnapshot(ByVal queryDefinitions As Object, ByVal queryName As String, ByRef failureDetail As String) As String
    If queryDefinitions Is Nothing Then
        failureDetail = "The QueryDefs snapshot is unavailable."
        Exit Function
    End If
    If queryDefinitions.Exists(queryName) = False Then
        failureDetail = "Required foundation query is missing: " & queryName
        Exit Function
    End If
    FoundationQueryDefinitionTextFromSnapshot = CStr(queryDefinitions(queryName))
End Function


Private Function CalculateFoundationQuerySnapshot(ByRef snapshot As FoundationDomainSnapshot, Optional ByVal suppliedQueryRows As Variant, Optional ByVal suppliedQueryNames As Variant, Optional ByVal suppliedVerifiedQueryEvidence As Variant) As Boolean
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim queryNames As Collection
    Dim queryRows As Collection
    Dim queryDefinitions As Object
    Dim queryRow As Variant
    Dim item As String
    Dim itemText As String
    Dim corePayload As String
    Dim formDesignSourcePayload As String
    Dim lookupPayload As String
    snapshot.failureDetail = vbNullString
    Set db = TablesWorkspaceCodeDb()
    If isMissing(suppliedQueryRows) = False Then
        If IsObject(suppliedQueryRows) Then Set queryRows = suppliedQueryRows
    End If
    If queryRows Is Nothing Then Set queryRows = FoundationRequiredNonFormQueryFamilyRows()
    If isMissing(suppliedQueryNames) = False Then
        If IsObject(suppliedQueryNames) Then Set queryNames = suppliedQueryNames
    End If
    If queryNames Is Nothing Then Set queryNames = FoundationRequiredNonFormQueryNames(queryRows)
    If isMissing(suppliedVerifiedQueryEvidence) = False Then
        If IsObject(suppliedVerifiedQueryEvidence) Then
            Set queryDefinitions = FoundationQueryDefinitionSnapshotFromVerifiedEvidence(suppliedVerifiedQueryEvidence, queryNames, snapshot.failureDetail)
            If queryDefinitions Is Nothing Then Exit Function
        End If
    End If
    If queryDefinitions Is Nothing Then
        db.QueryDefs.Refresh
        Set queryDefinitions = FoundationQueryDefinitionSnapshot(db, queryNames, snapshot.failureDetail)
    End If
    If queryDefinitions Is Nothing Then Exit Function
    corePayload = FingerprintTokenPart(FOUNDATION_FINGERPRINT_VERSION & "|CoreQueries")
    formDesignSourcePayload = FingerprintTokenPart(FOUNDATION_FINGERPRINT_VERSION & "|" & ManagedObjectFormDesignSourceFamilyCode())
    lookupPayload = FingerprintTokenPart(FOUNDATION_FINGERPRINT_VERSION & "|LookupQueries")
    For Each queryRow In queryRows
        PumpSetupControlEvents
        item = CStr(queryRow(0))
        itemText = FoundationQueryDefinitionTextFromSnapshot(queryDefinitions, item, snapshot.failureDetail)
        If Len(snapshot.failureDetail) > 0 Then Exit Function
        If CBool(queryRow(2)) Then
            lookupPayload = lookupPayload & itemText
            snapshot.LookupQueryCount = snapshot.LookupQueryCount + 1
        ElseIf CBool(queryRow(3)) Then
            formDesignSourcePayload = formDesignSourcePayload & itemText
            snapshot.FormDesignSourceCount = snapshot.FormDesignSourceCount + 1
        Else
            corePayload = corePayload & itemText
            snapshot.coreQueryCount = snapshot.coreQueryCount + 1
        End If
    Next queryRow
    snapshot.CoreQueryFingerprint = FingerprintFromPayload(corePayload)
    snapshot.FormDesignSourceFingerprint = FingerprintFromPayload(formDesignSourcePayload)
    snapshot.LookupQueryFingerprint = FingerprintFromPayload(lookupPayload)
    CalculateFoundationQuerySnapshot = True
    Exit Function
ErrHandler:
    snapshot.failureDetail = "Could not calculate the required saved-query foundation snapshot. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
End Function

Private Sub PersistFoundationQuerySnapshot(ByRef snapshot As FoundationDomainSnapshot)
    Dim db As DAO.Database
    Set db = TablesWorkspaceCodeDb()
    Call EnsureDaoPropertyValue(db, FOUNDATION_CORE_QUERY_FINGERPRINT_PROPERTY, dbText, snapshot.CoreQueryFingerprint, "PersistFoundationQuerySnapshot")
    Call EnsureDaoPropertyValue(db, FOUNDATION_FORMDESIGN_SOURCE_FINGERPRINT_PROPERTY, dbText, snapshot.FormDesignSourceFingerprint, "PersistFoundationQuerySnapshot")
    Call EnsureDaoPropertyValue(db, FOUNDATION_LOOKUP_QUERY_FINGERPRINT_PROPERTY, dbText, snapshot.LookupQueryFingerprint, "PersistFoundationQuerySnapshot")
End Sub

' ===============================================================================================================================================================================
' TABLE / QUERY SCHEMA DEFINITIONS AND REPAIR HELPERS
' ===============================================================================================================================================================================

' Purpose: Resolves config field order text from the supplied/current ADC state, using `SchemaDefinitionFieldOrderText` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function ConfigFieldOrderText() As String
    ' Wrapper purpose: TABLES-owned facade returning MC_Records/Config canonical physical field order from its authoritative schema definitions.
    ' Why retained: Retained so external schema consumers do not access private TABLES schema machinery or maintain an independent Config field-order mirror.
    'Intended use: Returns MC_Records canonical physical field order directly from the authoritative MC_Records schema definitions.
    ConfigFieldOrderText = SchemaDefinitionFieldOrderText(ConfigFieldSchemaDefinitions())
End Function


Public Function CandidateFieldOrderText() As String
    ' Wrapper purpose: TABLES-owned facade returning MC_Candidates canonical physical field order from its authoritative schema definitions.
    ' Why retained: Candidate discovery/review code can validate its physical storage without maintaining an independent field-order mirror.
    CandidateFieldOrderText = SchemaDefinitionFieldOrderText(CandidateFieldSchemaDefinitions())
End Function
' Purpose: Resolves and returns settings field order text from the supplied inputs and the authoritative ADC definitions/state used by this module.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: String: Returns the resolved settings field order text text in the canonical representation consumed by this module's caller; vbNullString represents no applicable value where the existing contract
'   permits it.
Public Function SettingsFieldOrderText() As String
    ' Wrapper purpose: TABLES-owned facade returning Settings canonical physical field order from its authoritative schema definitions.
    ' Why retained: Retained so external schema consumers do not access private TABLES schema machinery or maintain an independent Settings field-order mirror.
    'Intended use: Returns Settings canonical physical field order directly from the authoritative Settings schema definitions.
    SettingsFieldOrderText = SchemaDefinitionFieldOrderText(SettingsFieldSchemaDefinitions())
End Function

' Purpose: Resolves custom macro launcher review field order text from the supplied/current ADC state, using `SchemaDefinitionFieldOrderText` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function CustomMacroLauncherReviewFieldOrderText() As String
    ' Wrapper purpose: TABLES-owned facade returning the custom macro launcher review table's canonical physical field order from its authoritative schema definitions.
    ' Why retained: Retained so external schema consumers do not access private TABLES schema machinery or maintain an independent field-order mirror.
    'Intended use: Returns MC_Generated_LauncherCtrls_Review canonical physical field order directly from its authoritative schema definitions.
    CustomMacroLauncherReviewFieldOrderText = SchemaDefinitionFieldOrderText(CustomMacroLauncherReviewSchemaDefinitions())
End Function
' Purpose: Resolves foundation owned table field order text from the supplied/current ADC state, using `End` and `ManagedObjectPhysicalName` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   tableName (Input; String): Canonical or physical name identifying the table targeted by this operation.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function FoundationOwnedTableFieldOrderText(ByVal tableName As String) As String
    ' Wrapper purpose: TABLES-owned resolver for canonical field order across foundation tables whose authoritative schemas are owned by TABLES.
    ' Why retained: Retained so callers can request an owned table's canonical order through one contract instead of maintaining independent field-order mirrors.
    'Returns canonical field order for TABLES-owned schemas that previously had independent field-order mirrors.
    If EditableLookupTableName(tableName) Then
        FoundationOwnedTableFieldOrderText = FIELD_VALUE_TEXT & "|" & FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER & "|" & FIELD_DESCRIPTION
        Exit Function
    End If
    Select Case tableName
        Case ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_TraceLog")
            FoundationOwnedTableFieldOrderText = SchemaDefinitionFieldOrderText(TraceLogFieldSchemaDefinitions())
        Case ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_TERMINOLOGY_TABLE_LOGICAL_ID)
            FoundationOwnedTableFieldOrderText = SchemaDefinitionFieldOrderText(TerminologyFieldSchemaDefinitions())
    End Select
End Function

' Purpose: Resolves terminology field schema definitions from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: Collection: Returns the ordered Collection assembled by the routine; callers receive the object reference and should treat its membership/order according to the routine contract.
Public Function TerminologyFieldSchemaDefinitions() As Collection
    'Single physical field-schema authority for Help_Terms creation/repair and canonical order.
    ' CANONICAL AUTHORITY: rows owns the in-memory definition set constructed by TerminologyFieldSchemaDefinitions; downstream consumers derive views from this set rather than create a competing authority.
    Dim rows As New Collection
    rows.Add Array(FIELD_ID, dbLong, 0, True)
    rows.Add Array(FIELD_TERM_KEY, dbText, 120, False, True, False, Null, Null, Null)
    rows.Add Array(FIELD_TERM_TEXT, dbText, 255, False, True, False, Null, Null, Null)
    rows.Add Array(FIELD_DEFINITION, dbMemo, 0, False)
    rows.Add Array(FIELD_TERM_CATEGORY, dbText, 60, False, Null, Null, Null, Null, Null)
    rows.Add Array(FIELD_TERM_STATUS, dbText, 30, False, Null, Null, Null, Null, Null)
    rows.Add Array(FIELD_DEPRECATED, dbBoolean, 0, False)
    rows.Add Array(FIELD_APPROVED_USAGE, dbMemo, 0, False)
    rows.Add Array(FIELD_DISALLOWED_USAGE, dbMemo, 0, False)
    rows.Add Array(FIELD_REPLACEMENT_GUIDANCE, dbMemo, 0, False)
    rows.Add Array(FIELD_EXAMPLE_GOOD, dbMemo, 0, False)
    rows.Add Array(FIELD_EXAMPLE_BAD, dbMemo, 0, False)
    rows.Add Array(FIELD_DECISION_SCOPE, dbText, 60, False, Null, Null, Null, Null, Null)
    rows.Add Array(FIELD_DECISION_STATUS, dbText, 30, False, Null, Null, Null, Null, Null)
    rows.Add Array(FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER, dbLong, 0, False)
    rows.Add Array(FIELD_DEFAULTED_ON, dbDate, 0, False)
    rows.Add Array(FIELD_LAST_DEFAULT_REPAIR_ON, dbDate, 0, False)
    rows.Add Array(FIELD_MANAGED_OBJECT_DEPENDENCY_NOTES, dbMemo, 0, False)
    Set TerminologyFieldSchemaDefinitions = rows
End Function

' Purpose: Resolves terminology index schema definitions from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: Collection: Returns the ordered Collection assembled by the routine; callers receive the object reference and should treat its membership/order according to the routine contract.
Public Function TerminologyIndexSchemaDefinitions() As Collection
    'Single Help_Terms index-definition authority consumed by repair and validation.
    ' CANONICAL AUTHORITY: rows owns the in-memory definition set constructed by TerminologyIndexSchemaDefinitions; downstream consumers derive views from this set rather than create a competing authority.
    Dim rows As New Collection
    rows.Add Array(INDEX_PRIMARY_KEY, FIELD_ID, True, True)
    rows.Add Array(INDEX_UX_TERM_KEY, FIELD_TERM_KEY, True, False)
    rows.Add Array(INDEX_IX_TERM_TEXT, FIELD_TERM_TEXT, False, False)
    rows.Add Array(INDEX_IX_TERM_STATUS, FIELD_TERM_STATUS, False, False)
    rows.Add Array(INDEX_IX_TERM_CATEGORY, FIELD_TERM_CATEGORY, False, False)
    rows.Add Array(INDEX_IX_DEPRECATED, FIELD_DEPRECATED, False, False)
    rows.Add Array(INDEX_IX_DECISION_SCOPE, FIELD_DECISION_SCOPE, False, False)
    rows.Add Array(INDEX_IX_SORT_ORDER, FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER, False, False)
    Set TerminologyIndexSchemaDefinitions = rows
End Function

' Purpose: Resolves validation rule from value list from the supplied/current ADC state, using `Next` and `In` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   fieldName (Input; String): Canonical or physical name identifying the field targeted by this operation.
'   tableName (Input Optional; String): Canonical or physical name identifying the table targeted by this operation. Default: vbNullString.
' Returns: String: Returns the validation rule from value list text value. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function ValidationRuleFromValueList(ByVal fieldName As String, Optional ByVal tableName As String = vbNullString) As String
    Dim item As Variant
    Dim listText As String
    Dim ruleText As String
    listText = Trim$(ValueListForField(fieldName))
    If Len(listText) = 0 And Len(Trim$(tableName)) > 0 Then listText = Trim$(TableFieldLookupValueList(tableName, fieldName))
    If Len(listText) = 0 Then Exit Function
    For Each item In Split(listText, ";")
        If Len(ruleText) > 0 Then ruleText = ruleText & ","
        ruleText = ruleText & Chr$(34) & Replace(CStr(item), Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
    Next item
    ValidationRuleFromValueList = "In (" & ruleText & ")"
End Function

Private Function ConfigFieldSchemaDefinitions() As Collection
    'Single physical field-definition authority for configuration-only MC_Records rows.
    ' CANONICAL AUTHORITY: rows owns the in-memory definition set constructed by ConfigFieldSchemaDefinitions; downstream consumers derive views from this set rather than create a competing authority.
    Dim rows As New Collection
    Dim tableName As String
    tableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")
    rows.Add Array(FIELD_ID, dbLong, 0, True)
    rows.Add Array(FIELD_ACTIVE, dbBoolean, 0, False)
    rows.Add Array(FIELD_CONFIGURATION_STATUS, dbText, 30, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_CONFIGURED_ON, dbDate, 0, False)
    rows.Add Array(FIELD_MACRO_KEY, dbText, 100, False, True, False, Null, "Like ""*[! ]*""", "MacroKey is required and cannot be blank.")
    rows.Add Array(FIELD_MACRO_TITLE, dbText, ManagedFieldReqSizeForSetup(tableName, FIELD_MACRO_TITLE), False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_MACRO_ACTION, dbText, 30, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_MACRO_CATEGORY_KEY, dbText, 100, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_MACRO_SUBCATEGORY_KEY, dbText, 100, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_SELECTED_FOR_RUNNER_MACRO_GENERATION, dbBoolean, 0, False)
    rows.Add Array(FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER, dbLong, 0, False)
    rows.Add Array(FIELD_CONFIG_OWNER_DATABASE_ROLE, dbText, 30, False, False, False, Null, ValidationRuleFromValueList(FIELD_CONFIG_OWNER_DATABASE_ROLE), "Use a supported config owner database role.")
    rows.Add Array(FIELD_TARGET_PROCEDURE_NAME, dbText, 255, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_TARGET_PROCEDURE_MODULE_NAME, dbText, 255, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_TARGET_PROCEDURE_KIND, dbText, 30, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_TARGET_ARGUMENTS, dbMemo, 0, False)
    rows.Add Array(FIELD_RESULT_HANDLING, dbText, 30, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_EXPECTED_RETURN_TYPE, dbText, 30, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_TARGET_OBJECT_NAME, dbText, ManagedFieldReqSizeForSetup(tableName, FIELD_TARGET_OBJECT_NAME), False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_TARGET_REFERENCE_NAME, dbText, 255, False)
    rows.Add Array(FIELD_TARGET_VB_PROJECT_NAME, dbText, 255, False)
    rows.Add Array(FIELD_GENERATION_TARGET_DB_NAME, dbText, 255, False)
    rows.Add Array(FIELD_GENERATION_TARGET_DB_PATH, dbText, 255, False)
    rows.Add Array(FIELD_NAVIGATION_PANE_GROUP, dbText, 255, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_CONFIRMATION_MODE, dbText, 30, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_ACTION_DESCRIPTION, dbMemo, 0, False)
    rows.Add Array(FIELD_LAUNCHER_TOOLTIP, dbText, ManagedFieldReqSizeForSetup(tableName, FIELD_LAUNCHER_TOOLTIP), False)
    rows.Add Array(FIELD_OBJECT_DESCRIPTION, dbMemo, 0, False)
    rows.Add Array(FIELD_SOURCE_NOTES, dbMemo, 0, False)
    rows.Add Array(FIELD_REQUIRED_ARGUMENT_COUNT, dbLong, 0, False, False, Null, Null, ">=0", "Required argument count cannot be negative.")
    rows.Add Array(FIELD_OPTIONAL_ARGUMENT_COUNT, dbLong, 0, False, False, Null, Null, ">=0", "Optional argument count cannot be negative.")
    rows.Add Array(FIELD_REVIEW_PRIORITY, dbLong, 0, False)
    rows.Add Array(FIELD_MANUAL_REVIEW_STATUS, dbText, 30, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_MANUAL_REVIEW_NOTES, dbMemo, 0, False)
    rows.Add Array(FIELD_SELECTED_FOR_DUPLICATE_DELETION, dbBoolean, 0, False)
    rows.Add Array(FIELD_LAST_VALIDATED_ON, dbDate, 0, False)
    rows.Add Array(FIELD_LAST_VALIDATION_RESULT, dbText, 30, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_LAST_VALIDATION_NOTES, dbMemo, 0, False)
    rows.Add Array(FIELD_GENERATED_MACRO_OBJECT_NAME, dbText, ManagedFieldReqSizeForSetup(tableName, FIELD_GENERATED_MACRO_OBJECT_NAME), False)
    rows.Add Array(FIELD_LAST_GENERATED_ON, dbDate, 0, False)
    rows.Add Array(FIELD_LAST_GENERATED_RESULT, dbText, 30, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_LAST_GENERATED_NOTES, dbMemo, 0, False)
    Set ConfigFieldSchemaDefinitions = rows
End Function

Private Function CandidateFieldSchemaDefinitions() As Collection
    'Single physical field-definition authority for MC_Candidates discovery, proposed-configuration, review, and acceptance evidence.
    ' CANONICAL AUTHORITY: rows owns the in-memory definition set constructed by CandidateFieldSchemaDefinitions; downstream consumers derive views from this set rather than create a competing authority.
    Dim rows As New Collection
    rows.Add Array(FIELD_ID, dbLong, 0, True)
    rows.Add Array(FIELD_CANDIDATE_STATUS, dbText, 100, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_CANDIDATE_MACRO_KEY, dbText, 100, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_MACRO_KEY, dbText, 100, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_MACRO_TITLE, dbText, 255, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_MACRO_ACTION, dbText, 30, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_MACRO_CATEGORY_KEY, dbText, 100, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_MACRO_SUBCATEGORY_KEY, dbText, 100, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER, dbLong, 0, False)
    rows.Add Array(FIELD_CONFIG_OWNER_DATABASE_ROLE, dbText, 30, False, False, False, Null, ValidationRuleFromValueList(FIELD_CONFIG_OWNER_DATABASE_ROLE), "Use a supported config owner database role.")
    rows.Add Array(FIELD_CANDIDATE_OWNER_DATABASE_ROLE, dbText, 30, False, False, False, Null, ValidationRuleFromValueList(FIELD_CANDIDATE_OWNER_DATABASE_ROLE), "Use a supported candidate owner database role.")
    rows.Add Array(FIELD_CANDIDATE_SCAN_TARGET_DB_NAME, dbText, 255, False)
    rows.Add Array(FIELD_CANDIDATE_SCAN_TARGET_DB_PATH, dbText, 255, False)
    rows.Add Array(FIELD_CANDIDATE_VB_PROJECT_NAME, dbText, 255, False)
    rows.Add Array(FIELD_CANDIDATE_MODULE_NAME, dbText, 255, False)
    rows.Add Array(FIELD_TARGET_PROCEDURE_NAME, dbText, 255, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_TARGET_PROCEDURE_MODULE_NAME, dbText, 255, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_TARGET_PROCEDURE_KIND, dbText, 30, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_TARGET_ARGUMENTS, dbMemo, 0, False)
    rows.Add Array(FIELD_RESULT_HANDLING, dbText, 30, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_EXPECTED_RETURN_TYPE, dbText, 30, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_TARGET_OBJECT_NAME, dbText, 255, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_TARGET_REFERENCE_NAME, dbText, 255, False)
    rows.Add Array(FIELD_TARGET_VB_PROJECT_NAME, dbText, 255, False)
    rows.Add Array(FIELD_GENERATION_TARGET_DB_NAME, dbText, 255, False)
    rows.Add Array(FIELD_GENERATION_TARGET_DB_PATH, dbText, 255, False)
    rows.Add Array(FIELD_NAVIGATION_PANE_GROUP, dbText, 255, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_CONFIRMATION_MODE, dbText, 30, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_ACTION_DESCRIPTION, dbMemo, 0, False)
    rows.Add Array(FIELD_LAUNCHER_TOOLTIP, dbMemo, 0, False)
    rows.Add Array(FIELD_OBJECT_DESCRIPTION, dbMemo, 0, False)
    rows.Add Array(FIELD_SOURCE_NOTES, dbMemo, 0, False)
    rows.Add Array(FIELD_REVIEW_PRIORITY, dbLong, 0, False)
    rows.Add Array(FIELD_MANUAL_REVIEW_STATUS, dbText, 30, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_MANUAL_REVIEW_NOTES, dbMemo, 0, False)
    rows.Add Array(FIELD_SCAN_SIGNATURE, dbText, 255, False)
    rows.Add Array(FIELD_LAST_SCAN_BATCH_ID, dbText, 100, False)
    rows.Add Array(FIELD_LAST_SCANNED_ON, dbDate, 0, False)
    rows.Add Array(FIELD_DISCOVERED_ON, dbDate, 0, False)
    rows.Add Array(FIELD_MODULE_HEADER_FOUND, dbBoolean, 0, False)
    rows.Add Array(FIELD_MODULE_HEADER_MALFORMED, dbBoolean, 0, False)
    rows.Add Array(FIELD_MODULE_HEADER_TITLE, dbText, 255, False)
    rows.Add Array(FIELD_MODULE_HEADER_PURPOSE, dbMemo, 0, False)
    rows.Add Array(FIELD_MODULE_HEADER_ENTRY_POINTS, dbMemo, 0, False)
    rows.Add Array(FIELD_PROCEDURE_LISTED_IN_ENTRY_POINTS, dbBoolean, 0, False)
    rows.Add Array(FIELD_ENTRY_POINT_MATCH_CONFIDENCE, dbText, 30, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_MODULE_HEADER_SCAN_NOTES, dbMemo, 0, False)
    rows.Add Array(FIELD_CANDIDATE_PROCEDURE_DECLARATION, dbMemo, 0, False)
    rows.Add Array(FIELD_CANDIDATE_ARGUMENT_SIGNATURE, dbMemo, 0, False)
    rows.Add Array(FIELD_HAS_REQUIRED_ARGUMENTS, dbBoolean, 0, False)
    rows.Add Array(FIELD_REQUIRED_ARGUMENT_COUNT, dbLong, 0, False, False, Null, Null, ">=0", "Required argument count cannot be negative.")
    rows.Add Array(FIELD_OPTIONAL_ARGUMENT_COUNT, dbLong, 0, False, False, Null, Null, ">=0", "Optional argument count cannot be negative.")
    rows.Add Array(FIELD_CANDIDATE_NOTES, dbMemo, 0, False)
    rows.Add Array(FIELD_SELECTED_FOR_ACCEPTANCE, dbBoolean, 0, False)
    rows.Add Array(FIELD_ACCEPTANCE_STATUS, dbText, 100, False, False, False, Null, Null, Null)
    rows.Add Array(FIELD_ACCEPTED_ON, dbDate, 0, False)
    rows.Add Array(FIELD_ACCEPTED_CONFIGURATION_ID, dbLong, 0, False)
    rows.Add Array(FIELD_ACCEPTANCE_NOTES, dbMemo, 0, False)
    rows.Add Array(FIELD_REJECTED_ON, dbDate, 0, False)
    Set CandidateFieldSchemaDefinitions = rows
End Function
Private Function ConfigIndexSchemaDefinitions() As Collection
    'Single configuration-only MC_Records index-definition authority shared by CodeDb and narrow HostDb structural repair paths.
    'The fifth value marks an explicit index that CodeDb removes because B5 canonical RI supplies the same child-field index; HostDb retains it because HostDb receives no CodeDb-owned RI.
    ' CANONICAL AUTHORITY: rows owns the in-memory definition set constructed by ConfigIndexSchemaDefinitions; downstream consumers derive views from this set rather than create a competing authority.
    Dim rows As New Collection
    rows.Add Array(INDEX_PRIMARY_KEY, FIELD_ID, True, True, False)
    rows.Add Array(INDEX_IX_MACRO_KEY, FIELD_MACRO_KEY, True, False, False)
    rows.Add Array(INDEX_IX_MACRO_CATEGORY_KEY, FIELD_MACRO_CATEGORY_KEY, False, False, True)
    rows.Add Array(INDEX_IX_MACRO_SUBCATEGORY_KEY, FIELD_MACRO_SUBCATEGORY_KEY, False, False, True)
    rows.Add Array(INDEX_IX_CONFIG_OWNER_DATABASE_ROLE, FIELD_CONFIG_OWNER_DATABASE_ROLE, False, False, True)
    rows.Add Array(INDEX_IX_CONFIGURATION_STATUS, FIELD_CONFIGURATION_STATUS, False, False, True)
    rows.Add Array(INDEX_IX_MANUAL_REVIEW_STATUS, FIELD_MANUAL_REVIEW_STATUS, False, False, True)
    rows.Add Array(INDEX_IX_LAST_VALIDATION_RESULT, FIELD_LAST_VALIDATION_RESULT, False, False, True)
    rows.Add Array(INDEX_IX_LAST_GENERATED_RESULT, FIELD_LAST_GENERATED_RESULT, False, False, True)
    rows.Add Array(INDEX_IX_SELECTED_FOR_DUPLICATE_DELETION, FIELD_SELECTED_FOR_DUPLICATE_DELETION, False, False, False)
    rows.Add Array(INDEX_IX_GENERATED_MACRO_OBJECT_NAME, FIELD_GENERATED_MACRO_OBJECT_NAME, False, False, False)
    Set ConfigIndexSchemaDefinitions = rows
End Function

Private Function CandidateIndexSchemaDefinitions() As Collection
    'Single MC_Candidates index-definition authority. Relationship-backed child-field indexes are removed explicitly in CodeDb so B5 owns those indexes.
    ' CANONICAL AUTHORITY: rows owns the in-memory definition set constructed by CandidateIndexSchemaDefinitions; downstream consumers derive views from this set rather than create a competing authority.
    Dim rows As New Collection
    rows.Add Array(INDEX_PRIMARY_KEY, FIELD_ID, True, True, False)
    rows.Add Array(INDEX_IX_CANDIDATE_MACRO_KEY, FIELD_CANDIDATE_MACRO_KEY, False, False, False)
    rows.Add Array(INDEX_IX_SCAN_SIGNATURE, FIELD_SCAN_SIGNATURE, True, False, False)
    rows.Add Array(INDEX_IX_MACRO_CATEGORY_KEY, FIELD_MACRO_CATEGORY_KEY, False, False, True)
    rows.Add Array(INDEX_IX_MACRO_SUBCATEGORY_KEY, FIELD_MACRO_SUBCATEGORY_KEY, False, False, True)
    rows.Add Array(INDEX_IX_CONFIG_OWNER_DATABASE_ROLE, FIELD_CONFIG_OWNER_DATABASE_ROLE, False, False, True)
    rows.Add Array(INDEX_IX_CANDIDATE_OWNER_DATABASE_ROLE, FIELD_CANDIDATE_OWNER_DATABASE_ROLE, False, False, True)
    rows.Add Array(INDEX_IX_CANDIDATE_STATUS, FIELD_CANDIDATE_STATUS, False, False, True)
    rows.Add Array(INDEX_IX_MANUAL_REVIEW_STATUS, FIELD_MANUAL_REVIEW_STATUS, False, False, True)
    rows.Add Array(INDEX_IX_SELECTED_FOR_ACCEPTANCE, FIELD_SELECTED_FOR_ACCEPTANCE, False, False, False)
    rows.Add Array(INDEX_IX_ACCEPTANCE_STATUS, FIELD_ACCEPTANCE_STATUS, False, False, True)
    rows.Add Array(INDEX_UX_ACCEPTED_CONFIGURATION_ID, FIELD_ACCEPTED_CONFIGURATION_ID, True, False, False)
    Set CandidateIndexSchemaDefinitions = rows
End Function
Private Function ConfigIndexSupersededByCodeDbCanonicalRelationship(ByVal indexDefinition As Variant) As Boolean
    'The optional fifth ConfigIndexSchemaDefinitions value is True only when CodeDb B5 RI owns the equivalent child-field index.
    If UBound(indexDefinition) < 4 Then Exit Function
    ConfigIndexSupersededByCodeDbCanonicalRelationship = CBool(indexDefinition(4))
End Function

Private Sub ApplySchemaFieldContract(ByVal fld As DAO.Field, ByVal fieldDefinition As Variant, ByVal tableName As String, ByVal schemaLabel As String)
    On Error GoTo ErrHandler
    If UBound(fieldDefinition) >= 4 Then If IsNull(fieldDefinition(4)) = False Then If fld.Required <> CBool(fieldDefinition(4)) Then fld.Required = CBool(fieldDefinition(4))
    If UBound(fieldDefinition) >= 5 Then
        If IsNull(fieldDefinition(5)) = False And (fld.Type = dbText Or fld.Type = dbMemo) Then If fld.allowZeroLength <> CBool(fieldDefinition(5)) Then fld.allowZeroLength = CBool(fieldDefinition(5))
    End If
    If UBound(fieldDefinition) >= 6 Then If IsNull(fieldDefinition(6)) = False Then If StrComp(CStr(Nz(fld.DefaultValue, vbNullString)), CStr(fieldDefinition(6)), vbBinaryCompare) <> 0 Then fld.DefaultValue = CStr(fieldDefinition(6))
    If UBound(fieldDefinition) >= 7 Then If IsNull(fieldDefinition(7)) = False Then If StrComp(CStr(Nz(fld.validationRule, vbNullString)), CStr(fieldDefinition(7)), vbBinaryCompare) <> 0 Then fld.validationRule = CStr(fieldDefinition(7))
    If UBound(fieldDefinition) >= 8 Then If IsNull(fieldDefinition(8)) = False Then If StrComp(CStr(Nz(fld.validationText, vbNullString)), CStr(fieldDefinition(8)), vbBinaryCompare) <> 0 Then fld.validationText = CStr(fieldDefinition(8))
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "ApplySchemaFieldContract", "Could not enforce the canonical " & schemaLabel & " field contract." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fld.Name & vbCrLf & "Err.Description: " & Err.Description
End Sub

Private Function SchemaFieldContractIssueText(ByVal fld As DAO.Field, ByVal fieldDefinition As Variant) As String
    Dim issueText As String
    If UBound(fieldDefinition) >= 4 Then If IsNull(fieldDefinition(4)) = False Then If fld.Required <> CBool(fieldDefinition(4)) Then issueText = issueText & "; Required expected " & CStr(CBool(fieldDefinition(4))) & " found " & CStr(fld.Required)
    If UBound(fieldDefinition) >= 5 Then If IsNull(fieldDefinition(5)) = False And (fld.Type = dbText Or fld.Type = dbMemo) Then If fld.allowZeroLength <> CBool(fieldDefinition(5)) Then issueText = issueText & "; AllowZeroLength expected " & CStr(CBool(fieldDefinition(5))) & " found " & CStr(fld.allowZeroLength)
    If UBound(fieldDefinition) >= 6 Then If IsNull(fieldDefinition(6)) = False Then If StrComp(CStr(Nz(fld.DefaultValue, vbNullString)), CStr(fieldDefinition(6)), vbBinaryCompare) <> 0 Then issueText = issueText & "; DefaultValue expected [" & CStr(fieldDefinition(6)) & "] found [" & CStr(Nz(fld.DefaultValue, vbNullString)) & "]"
    If UBound(fieldDefinition) >= 7 Then If IsNull(fieldDefinition(7)) = False Then If StrComp(CStr(Nz(fld.validationRule, vbNullString)), CStr(fieldDefinition(7)), vbBinaryCompare) <> 0 Then issueText = issueText & "; ValidationRule expected [" & CStr(fieldDefinition(7)) & "] found [" & CStr(Nz(fld.validationRule, vbNullString)) & "]"
    If UBound(fieldDefinition) >= 8 Then If IsNull(fieldDefinition(8)) = False Then If StrComp(CStr(Nz(fld.validationText, vbNullString)), CStr(fieldDefinition(8)), vbBinaryCompare) <> 0 Then issueText = issueText & "; ValidationText expected [" & CStr(fieldDefinition(8)) & "] found [" & CStr(Nz(fld.validationText, vbNullString)) & "]"
    If Len(issueText) > 0 Then SchemaFieldContractIssueText = Mid$(issueText, 3)
End Function

Private Sub AppendSchemaFieldIfMissing(ByVal tdf As DAO.tableDef, ByVal fieldName As String, ByVal fieldType As DAO.DataTypeEnum, Optional ByVal fieldSize As Long = 0, Optional ByVal autoNumber As Boolean = False, Optional ByVal schemaLabel As String = "managed", Optional ByVal fieldDefinition As Variant)
    If SharedDao_TableDefFieldExistsSafe(tdf, fieldName) Then Exit Sub
    If isMissing(fieldDefinition) Then
        AppendSchemaField tdf, fieldName, fieldType, fieldSize, autoNumber, schemaLabel
    Else
        AppendSchemaField tdf, fieldName, fieldType, fieldSize, autoNumber, schemaLabel, fieldDefinition
    End If
End Sub

Private Sub AppendSchemaField(ByVal tdf As DAO.tableDef, ByVal fieldName As String, ByVal fieldType As DAO.DataTypeEnum, Optional ByVal fieldSize As Long = 0, Optional ByVal autoNumber As Boolean = False, Optional ByVal schemaLabel As String = "managed", Optional ByVal fieldDefinition As Variant)
    'Appends a field to a fresh or already-inspected TableDef without introducing a second schema definition.
    On Error GoTo ErrHandler
    Dim fld As DAO.Field
    If fieldSize > 0 Then
        Set fld = tdf.CreateField(fieldName, fieldType, fieldSize)
    Else
        Set fld = tdf.CreateField(fieldName, fieldType)
    End If
    If autoNumber Then fld.Attributes = fld.Attributes Or dbAutoIncrField
    If isMissing(fieldDefinition) = False Then ApplySchemaFieldContract fld, fieldDefinition, tdf.Name, schemaLabel
    tdf.fields.Append fld
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "AppendSchemaField", "Could not append " & DEFAULT_UTILITY_LONG_NAME & " " & schemaLabel & " field." & vbCrLf & "Table: " & tdf.Name & vbCrLf & "Field: " & fieldName & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

Private Sub AddFieldDefinitionIssuesFromSchema(ByRef issues As Collection, ByVal tableName As String, ByVal definitions As Collection, Optional ByVal suppliedTableDef As Variant)
    Dim fieldDefinition As Variant
    For Each fieldDefinition In definitions
        If isMissing(suppliedTableDef) Then
            AddFieldDefinitionIssue issues, tableName, CStr(fieldDefinition(0)), CLng(fieldDefinition(1)), CLng(fieldDefinition(2)), CBool(fieldDefinition(3)), fieldDefinition
        Else
            AddFieldDefinitionIssue issues, tableName, CStr(fieldDefinition(0)), CLng(fieldDefinition(1)), CLng(fieldDefinition(2)), CBool(fieldDefinition(3)), fieldDefinition, suppliedTableDef
        End If
    Next fieldDefinition
End Sub

Private Sub AddFieldDefinitionIssue(ByRef issues As Collection, ByVal tableName As String, ByVal fieldName As String, ByVal expectedType As DAO.DataTypeEnum, Optional ByVal expectedSize As Long = 0, Optional ByVal mustBeAutoNumber As Boolean = False, Optional ByVal fieldDefinition As Variant, Optional ByVal suppliedTableDef As Variant)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim fld As DAO.Field
    If CodeDbTableExists(tableName) = False Then
        issues.Add tableName & ": table is missing"
        Exit Sub
    End If
    'Keep the DAO.Database object alive while the TableDef/Field objects are inspected.
    'Using ADCDb().TableDefs(tableName) inline can leave child DAO objects invalid after Access/DAO rebuilds or renames a table. That produced false "Object invalid or no longer set" schema failures after a
    'successful rebuild.
    'A2 performs an explicit batch-boundary refresh before final verification. Reusing that synchronized TableDef avoids an expensive collection refresh for every individual field inspection.
    If isMissing(suppliedTableDef) = False Then
        If IsObject(suppliedTableDef) Then Set tdf = suppliedTableDef
    End If
    If tdf Is Nothing Then
        Set db = TablesWorkspaceCodeDb()
        If mTableDesignBatchPrepared = False Then db.TableDefs.Refresh
        Set tdf = db.TableDefs(tableName)
        If mTableDesignBatchPrepared = False Then tdf.fields.Refresh
    End If
    If SharedDao_TableDefFieldExistsSafe(tdf, fieldName) = False Then
        issues.Add tableName & ": missing field [" & fieldName & "]"
        Exit Sub
    End If
    Set fld = tdf.fields(fieldName)
    If SharedDao_FieldDefinitionMatches(fld, expectedType, expectedSize, mustBeAutoNumber) = False Then
        issues.Add tableName & ": field [" & fieldName & "] has wrong definition; expected " & FieldDefinitionDescription(expectedType, expectedSize, mustBeAutoNumber) & ", found " & FieldActualDefinitionDescription(fld)
    ElseIf isMissing(fieldDefinition) = False Then
        Dim contractIssue As String
        contractIssue = SchemaFieldContractIssueText(fld, fieldDefinition)
        If Len(contractIssue) > 0 Then issues.Add tableName & ": field [" & fieldName & "] physical contract mismatch; " & contractIssue
    End If
    Exit Sub
ErrHandler:
    issues.Add tableName & ": could not inspect field [" & fieldName & "] - " & "Err.Number " & CStr(Err.Number) & "; Err.Source " & Err.Source & "; Err.Description " & Err.Description
End Sub

' Purpose: Resolves and returns field definition description from the supplied inputs and the authoritative ADC definitions/state used by this module.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   fieldType (Input; DAO.DataTypeEnum): DAO DataTypeEnum value specifying the exact physical Access field type to create or validate.
'   fieldSize (Input Optional; Long): DAO field-size value used when creating or validating a variable-length field; zero leaves sizing to the routine’s type-specific/default handling. Default: 0.
'   isAutoNumber (Input Optional; Boolean): Indicates whether the supplied item/state is auto number; the routine uses this distinction to apply the matching validation or persistence rules. Default: False.
' Returns: String: Returns the canonical description text for field definition description, suitable for the managed catalog/field/object metadata destination; vbNullString indicates that no description applies
'   where allowed.
Public Function FieldDefinitionDescription(ByVal fieldType As DAO.DataTypeEnum, Optional ByVal fieldSize As Long = 0, Optional ByVal isAutoNumber As Boolean = False) As String
    FieldDefinitionDescription = DaoFieldTypeName(fieldType)
    If fieldType = dbText And fieldSize > 0 Then
        FieldDefinitionDescription = FieldDefinitionDescription & "(" & CStr(fieldSize) & ")"
    End If
    If isAutoNumber Then
        FieldDefinitionDescription = FieldDefinitionDescription & " AutoNumber"
    End If
End Function

' Purpose: Resolves and returns field actual definition description from the supplied inputs and the authoritative ADC definitions/state used by this module.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   fld (Input; DAO.Field): DAO.Field representing the field whose properties or metadata are being inspected or modified.
' Returns: String: Returns the canonical description text for field actual definition description, suitable for the managed catalog/field/object metadata destination; vbNullString indicates that no description
'   applies where allowed.
Public Function FieldActualDefinitionDescription(ByVal fld As DAO.Field) As String
    FieldActualDefinitionDescription = DaoFieldTypeName(fld.Type)
    If fld.Type = dbText Then
        FieldActualDefinitionDescription = FieldActualDefinitionDescription & "(" & CStr(fld.Size) & ")"
    End If
    If (fld.Attributes And dbAutoIncrField) <> 0 Then
        FieldActualDefinitionDescription = FieldActualDefinitionDescription & " AutoNumber"
    End If
End Function

Private Function DaoFieldTypeName(ByVal fieldType As DAO.DataTypeEnum) As String
    Select Case fieldType
        Case dbBoolean
            DaoFieldTypeName = "dbBoolean"
        Case dbByte
            DaoFieldTypeName = "dbByte"
        Case dbInteger
            DaoFieldTypeName = "dbInteger"
        Case dbLong
            DaoFieldTypeName = "dbLong"
        Case dbCurrency
            DaoFieldTypeName = "dbCurrency"
        Case dbSingle
            DaoFieldTypeName = "dbSingle"
        Case dbDouble
            DaoFieldTypeName = "dbDouble"
        Case dbDate
            DaoFieldTypeName = "dbDate"
        Case dbText
            DaoFieldTypeName = "dbText"
        Case dbLongBinary
            DaoFieldTypeName = "dbLongBinary"
        Case dbMemo
            DaoFieldTypeName = "dbMemo"
        Case Else
            DaoFieldTypeName = "DAO type " & CStr(fieldType)
    End Select
End Function

Private Sub EnsureTerminologyFields(ByVal tableName As String, ByVal schemaDefinitions As Collection)
    'Uses the TABLES-owned schema directly; existing incompatible fields remain a hard failure instead of being silently rebuilt.
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim fld As DAO.Field
    Dim fieldDefinition As Variant
    Set db = TablesWorkspaceCodeDb()
    db.TableDefs.Refresh
    Set tdf = db.TableDefs(tableName)
    tdf.fields.Refresh
    For Each fieldDefinition In schemaDefinitions
        If CBool(fieldDefinition(3)) Then
            If SharedDao_TableDefFieldExistsSafe(tdf, CStr(fieldDefinition(0))) Then
                Set fld = tdf.fields(CStr(fieldDefinition(0)))
                If SharedDao_FieldDefinitionMatches(fld, CLng(fieldDefinition(1)), CLng(fieldDefinition(2)), True) = False Then Err.Raise vbObjectError + 5901, "EnsureTerminologyFields", "Help_Terms field has an incompatible definition." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & CStr(fieldDefinition(0)) & vbCrLf & "Expected: " & FieldDefinitionDescription(CLng(fieldDefinition(1)), CLng(fieldDefinition(2)), True) & vbCrLf & "Actual: " & FieldActualDefinitionDescription(fld)
            Else
                AppendSchemaField tdf, CStr(fieldDefinition(0)), CLng(fieldDefinition(1)), CLng(fieldDefinition(2)), True, "help_terms_schema", fieldDefinition
                tdf.fields.Refresh
            End If
        Else
            If SharedDao_TableDefFieldExistsSafe(tdf, CStr(fieldDefinition(0))) Then
                RepairTableField tableName, CStr(fieldDefinition(0)), CLng(fieldDefinition(1)), CLng(fieldDefinition(2))
            Else
                RepairTableField tableName, CStr(fieldDefinition(0)), CLng(fieldDefinition(1)), CLng(fieldDefinition(2))
                tdf.fields.Refresh
            End If
        End If
        ApplySchemaFieldContract tdf.fields(CStr(fieldDefinition(0))), fieldDefinition, tableName, "help_terms_schema"
    Next fieldDefinition
    EnsureTableDefFieldOrderFromText TablesWorkspaceCodeDb().TableDefs(tableName), SchemaDefinitionFieldOrderText(schemaDefinitions), mTableDesignBatchPrepared
End Sub

Private Sub CreateTerminologyTable(ByVal tableName As String, ByVal schemaDefinitions As Collection)
    'Creates the complete Help_Terms structure in canonical field order from the single TABLES-owned schema definition.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim fieldDefinition As Variant
    Set db = TablesWorkspaceCodeDb()
    Set tdf = db.CreateTableDef(tableName)
    For Each fieldDefinition In schemaDefinitions
        AppendSchemaField tdf, CStr(fieldDefinition(0)), CLng(fieldDefinition(1)), CLng(fieldDefinition(2)), CBool(fieldDefinition(3)), "help_terms_schema", fieldDefinition
    Next fieldDefinition
    db.TableDefs.Append tdf
    TrackSetupCreatedTableRollback tableName
    db.TableDefs.Refresh
    If SharedDao_TableExists(db, tableName) = False Then Err.Raise vbObjectError + 9420, "CreateTerminologyTable", "The Help_Terms table was not visible after DAO appended it."
    CoreAddOperationNote "Created Help_Terms table [" & tableName & "] from the TABLES-owned canonical schema."
    CoreAddOperationNote "Deferred Navigation Pane grouping for newly created table [" & tableName & "] until the final catalog-driven grouping pass after MAC defaults exist."
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrSource As String
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrSource = Err.Source
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "CreateTerminologyTable", "Could not create the complete Help_Terms table." & vbCrLf & "Table: " & tableName & vbCrLf & "Err.Source: " & originalErrSource & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

Private Sub EnsureTerminologyLookupFields(ByVal tableName As String)
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim ignoredWarningCount As Long
    Set db = TablesWorkspaceCodeDb()
    db.TableDefs.Refresh
    Set tdf = db.TableDefs(tableName)
    Call ApplyB2SharedFieldMetadataDefinitionsToTableDef(tdf, ignoredWarningCount)
End Sub

Private Sub EnsureConfigCurrentDesignFields(ByVal tableName As String, ByVal schemaDefinitions As Collection)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    If CodeDbTableExists(tableName) = False Then Exit Sub
    If TableCanOpen(tableName, True) = False Then Exit Sub
    Set db = TablesWorkspaceCodeDb()
    db.TableDefs.Refresh
    Set tdf = db.TableDefs(tableName)
    EnsureConfigFieldsOnTableDef tdf, schemaDefinitions
    tdf.fields.Refresh
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrSource As String
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrSource = Err.Source
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "EnsureConfigCurrentDesignFields", "Unable to add missing current-design MC_Records table fields before strict schema validation." & vbCrLf & "Table: " & tableName & vbCrLf & "Expected additive fields: current MC_Records schema fields" & vbCrLf & "Error number: " & CStr(originalErrNumber) & vbCrLf & "Source: " & originalErrSource & vbCrLf & "Description: " & originalErrDescription
End Sub

Private Sub EnsureLauncherConfigFields(ByVal tableName As String)
    On Error GoTo ErrHandler
    If CodeDbTableExists(tableName) = False Then Exit Sub
    RepairTableField tableName, FIELD_LAUNCHER_TOOLTIP, dbText, ManagedFieldReqSizeForSetup(tableName, FIELD_LAUNCHER_TOOLTIP)
    RepairTableField tableName, FIELD_OBJECT_DESCRIPTION, dbMemo
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "EnsureLauncherConfigFields", "Could not ensure launcher config fields." & vbCrLf & "Table: " & tableName & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

Private Function SafeDAOFieldName(ByVal fld As DAO.Field) As String
    On Error GoTo ErrHandler
    SafeDAOFieldName = fld.Name
    Exit Function
ErrHandler:
    SafeDAOFieldName = "(could not resolve field name)"
End Function

Private Sub RepairConfigLookupFields(ByVal tableName As String, Optional ByVal showMessage As Boolean = True)
    'Intended use: Runs Repair Access Development Console Config Lookup Fields for direct, generated-macro, form-button, or required cross-module callers.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim changedCount As Long
    Dim warningCount As Long
    If CodeDbTableExists(tableName) = False Then Exit Sub
    If StrComp(tableName, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config"), vbTextCompare) = 0 Then
        If CodeDbTableExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_MacroCategories")) = False Or CodeDbTableExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_MacroSubcategories")) = False Then RepairMacroClassificationTables False
    End If
    Set db = TablesWorkspaceCodeDb()
    Set tdf = db.TableDefs(tableName)
    changedCount = ApplyB2SharedFieldMetadataDefinitionsToTableDef(tdf, warningCount)
    If changedCount > 0 Then CoreAddOperationNote "Repaired " & CStr(changedCount) & " lookup field(s) on table [" & tableName & "]."
    If warningCount > 0 Then CoreAddOperationNote "Config lookup metadata repair completed with " & CStr(warningCount) & " warning(s). Review the preceding lookup-property diagnostic note(s)."
    ReportCoreActionSuccess "RepairConfigLookupFields", "TableRepair", IIf(warningCount > 0, DEFAULT_UTILITY_SHORT_NAME & " config lookup metadata repair completed with warning(s).", DEFAULT_UTILITY_SHORT_NAME & " config lookup metadata repair completed successfully."), DEFAULT_UTILITY_LONG_NAME & " config lookup metadata repair complete." & vbCrLf & vbCrLf & "Table: " & tableName & vbCrLf & "Lookup fields changed: " & CStr(changedCount) & vbCrLf & "Lookup warnings: " & CStr(warningCount), "Repair config lookup fields", tableName, "Table", "CodeDb", "TableRepair", showMessage, "Repair Config Lookup Fields", vbInformation, True, "Tables", False, visibleStatusMessage:=IIf(warningCount > 0, "Config lookup metadata repair completed with warning(s).", "Config lookup metadata repair completed successfully.")
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim messageText As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    messageText = "Could not repair macro-creator config lookup metadata." & vbCrLf & vbCrLf & "Procedure: RepairConfigLookupFields" & vbCrLf & "Table: " & tableName & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description:" & vbCrLf & savedErrDescription & TableReadinessDiagnosticsForMessage() & TraceBackendDiagnosticsForMessage()
    LogTraceMessage TRACE_LEVEL_ERROR, "TableRepair", "RepairConfigLookupFields", "Config lookup metadata repair failed.", messageText, "Repair config lookup fields", savedErrNumber, savedErrSource, savedErrDescription, tableName, "Table", "CodeDb", "TableRepair"
    If showMessage Then
        CoreShowErrorMessage messageText, vbCritical, "Repair Config Lookup Fields", "RepairConfigLookupFields", "Tables"
        Exit Sub
    End If
    Err.Raise savedErrNumber, "RepairConfigLookupFields", messageText
End Sub

Private Function EnsureDefinedFieldLookup(ByVal tdf As DAO.tableDef, ByVal fieldName As String) As Long
    Dim Definition As Variant
    Dim lookupKind As String
    Definition = FieldTableLookupDefinition(tdf.Name, fieldName)
    If IsEmpty(Definition) Then Exit Function
    lookupKind = CStr(Definition(0))
    Select Case lookupKind
        Case FIELD_LOOKUP_DEFINITION_KIND_TABLE
            EnsureDefinedFieldLookup = EnsureFieldTableLookup(tdf, fieldName, CStr(Definition(1)), CStr(Definition(2)), CStr(Definition(3)), CBool(Definition(4)), CStr(Definition(5)))
        Case FIELD_LOOKUP_DEFINITION_KIND_LIVE_TARGET
            EnsureDefinedFieldLookup = EnsureLiveObjectLookupField(tdf, fieldName, LiveTargetObjectsLookupRowSource(), CBool(Definition(4)), CLng(Definition(6)), CLng(Definition(7)), CStr(Definition(5)))
        Case Else
            Err.Raise vbObjectError + 6517, "EnsureDefinedFieldLookup", "Unknown DEFAULTS-owned " & DEFAULT_UTILITY_SHORT_NAME & " lookup definition kind." & vbCrLf & "Table: " & tdf.Name & vbCrLf & "Field: " & fieldName & vbCrLf & "Kind: " & lookupKind
    End Select
End Function

' Purpose: Ensures field value list lookup for tdf, field name.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   tdf (Input; DAO.tableDef): DAO.TableDef representing the table whose fields, indexes, properties, or metadata are being inspected or modified.
'   fieldName (Input; String): Field name identifying the table/query field or bound control source the routine evaluates or modifies.
'   rowSourceText (Input; String): SQL/row-source text consumed by this procedure; it is parsed, compared, or assigned as coded without changing the caller's original string.
'   BoundColumn (Input Optional; Long): One-based Access lookup column number whose value is stored as the control/field value. Default: 0.
'   LimitToList (Input Optional; Boolean): True restricts the Access lookup control/field to values present in its row source; False permits values outside the listed choices when Access allows them. Default:
'   True.
' Returns: Long: Returns the number/result code produced while applying the requested field/lookup metadata; callers use it to detect whether physical metadata required creation or repair.
Public Function EnsureFieldValueListLookup(ByVal tdf As DAO.tableDef, ByVal fieldName As String, ByVal rowSourceText As String, Optional ByVal BoundColumn As Long = 0, Optional ByVal LimitToList As Boolean = True) As Long
    On Error GoTo ErrHandler
    Dim fld As DAO.Field
    Dim changed As Boolean
    Dim propertyMask As LookupFieldPropertyMask
    If SharedDao_TableDefFieldExistsSafe(tdf, fieldName) = False Then Exit Function
    Set fld = tdf.fields(fieldName)
    propertyMask = LOOKUP_FIELD_PROPERTY_DISPLAY_CONTROL Or LOOKUP_FIELD_PROPERTY_ROW_SOURCE_TYPE Or LOOKUP_FIELD_PROPERTY_ROW_SOURCE Or LOOKUP_FIELD_PROPERTY_COLUMN_COUNT Or LOOKUP_FIELD_PROPERTY_LIMIT_TO_LIST
    If BoundColumn > 0 Then propertyMask = propertyMask Or LOOKUP_FIELD_PROPERTY_BOUND_COLUMN
    changed = ApplyLookupFieldProperties(fld, "Value List", rowSourceText, LimitToList, 1, IIf(BoundColumn > 0, BoundColumn, 1), vbNullString, "EnsureFieldValueListLookup", vbObjectError + 6494, vbObjectError + 6495, propertyMask)
    If changed Then EnsureFieldValueListLookup = 1
    Exit Function
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "EnsureFieldValueListLookup", "Could not apply hard-coded lookup options to field." & vbCrLf & "Field: " & fieldName & vbCrLf & "Err.Description: " & originalErrDescription
End Function

' Purpose: Checks whether set recordset field if exists in the database, collection, or metadata scope supplied to this function.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   rs (Input; DAO.Recordset): DAO.Recordset supplying or receiving the row data used by the operation.
'   fieldName (Input; String): Field name identifying the table/query field or bound control source the routine evaluates or modifies.
'   fieldValue (Input; Variant): Field value to test/format/store; may be Null/Empty or another Variant subtype when the surrounding field contract permits it.
'   valueSourceName (Input Optional; String): DEFAULTS constant/expression, persisted field, or runtime source that supplied the value; used only for precise overlength diagnostics. Default: vbNullString.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub SetRecordsetFieldIfExists(ByVal rs As DAO.Recordset, ByVal fieldName As String, ByVal fieldValue As Variant, Optional ByVal valueSourceName As String = vbNullString)
    'Intended use: Runs Set Recordset Field If Exists for direct, generated-macro, form-button, or required cross-module callers.
    On Error GoTo ErrHandler
    Dim targetField As DAO.Field
    Dim storageValue As Variant
    If SharedDao_RecordsetHasFieldSafe(rs, fieldName) Then
        Set targetField = rs.fields(fieldName)
        storageValue = RecordsetFieldValueForStorage(targetField, fieldValue)
        ValidateRecordsetTextFieldValueLength targetField, fieldName, storageValue, valueSourceName
        targetField.value = storageValue
    End If
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "SetRecordsetFieldIfExists", "Could not assign recordset field." & vbCrLf & "Field: " & fieldName & vbCrLf & "Default/value source: " & IIf(Len(Trim$(valueSourceName)) > 0, valueSourceName, "(not applicable; runtime/caller value)") & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

' Purpose: Assigns a required recordset field while preserving direct DAO assignment semantics and adding source-aware Short Text length validation.
' Public role: Inter-module public API for existing required-field writes that must retain missing-field and zero-length behavior while producing precise overlength diagnostics.
' Parameters:
'   rs (Input; DAO.Recordset): DAO.Recordset receiving the field value.
'   fieldName (Input; String): Required destination field name.
'   fieldValue (Input; Variant): Value assigned without storage normalization.
'   valueSourceName (Input; String): DEFAULTS constant/function or other named source that supplied the value for overlength diagnostics.
' Returns: None (Sub). The procedure assigns the value or raises a source-aware error.
Public Sub SetRecordsetFieldValueWithSource(ByVal rs As DAO.Recordset, ByVal fieldName As String, ByVal fieldValue As Variant, ByVal valueSourceName As String)
    On Error GoTo ErrHandler
    Dim targetField As DAO.Field
    Set targetField = rs.fields(fieldName)
    ValidateRecordsetTextFieldValueLength targetField, fieldName, fieldValue, valueSourceName
    targetField.value = fieldValue
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "SetRecordsetFieldValueWithSource", "Could not assign required recordset field." & vbCrLf & "Field: " & fieldName & vbCrLf & "Default/value source: " & IIf(Len(Trim$(valueSourceName)) > 0, valueSourceName, "(not applicable; runtime/caller value)") & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

Private Sub ValidateRecordsetTextFieldValueLength(ByVal targetField As DAO.Field, ByVal fieldName As String, ByVal fieldValue As Variant, ByVal valueSourceName As String)
    If targetField.Type <> dbText Or IsNull(fieldValue) Then Exit Sub
    If Len(CStr(fieldValue)) <= targetField.Size Then Exit Sub
    Err.Raise vbObjectError + 10119, "ValidateRecordsetTextFieldValueLength", "A value exceeds the destination Short Text field length." & vbCrLf & "Field: " & fieldName & vbCrLf & "Maximum length: " & CStr(targetField.Size) & vbCrLf & "Value length: " & CStr(Len(CStr(fieldValue))) & vbCrLf & "Default/value source: " & IIf(Len(Trim$(valueSourceName)) > 0, valueSourceName, "(not applicable; runtime/caller value)")
End Sub

' Purpose: Checks for field value if exists without requiring the caller to handle the routine's normal not-found path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   rs (Input; DAO.Recordset): DAO Recordset supplying the current rows used or updated by the routine.
'   fieldName (Input; String): Canonical or physical name identifying the field targeted by this operation.
' Returns: Variant: Returns the field value if exists value in the scalar, array, object, Empty, or Null form produced by the routine's existing contract.
Public Function FieldValueIfExists(ByVal rs As DAO.Recordset, ByVal fieldName As String) As Variant
    'Returns Null when the requested field is not present in the recordset.
    If SharedDao_RecordsetHasFieldSafe(rs, fieldName) Then
        FieldValueIfExists = rs.fields(fieldName).value
    Else
        FieldValueIfExists = Null
    End If
End Function

Private Function RecordsetFieldValueForStorage(ByVal targetField As DAO.Field, ByVal fieldValue As Variant) As Variant
    If IsNull(fieldValue) Then
        RecordsetFieldValueForStorage = Null
        Exit Function
    End If
    If VarType(fieldValue) = vbString Then
        If Len(CStr(fieldValue)) = 0 Then
            If FieldDisallowsZeroLength(targetField) Then
                RecordsetFieldValueForStorage = Null
            Else
                RecordsetFieldValueForStorage = vbNullString
            End If
            Exit Function
        End If
    End If
    RecordsetFieldValueForStorage = fieldValue
End Function

Private Function FieldDisallowsZeroLength(ByVal targetField As DAO.Field) As Boolean
    If targetField.Type <> dbText And targetField.Type <> dbMemo Then Exit Function
    FieldDisallowsZeroLength = True
    On Error Resume Next
    FieldDisallowsZeroLength = Not CBool(targetField.properties("AllowZeroLength").value)
    Err.Clear
    On Error GoTo 0
End Function

Private Function TerminologySchemaRepairRequired(ByRef repairSummary As String, Optional ByVal suppliedSchemaDefinitions As Variant, Optional ByVal suppliedIndexDefinitions As Variant) As Boolean
    On Error GoTo ErrHandler
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: issues holds planned/recovery work or independently observed validation evidence for TerminologySchemaRepairRequired; keep it distinct from the source authority.
    Dim issues As New Collection
    Dim schemaDefinitions As Collection
    Dim indexDefinitions As Collection
    If isMissing(suppliedSchemaDefinitions) = False Then
        If IsObject(suppliedSchemaDefinitions) Then Set schemaDefinitions = suppliedSchemaDefinitions
    End If
    If isMissing(suppliedIndexDefinitions) = False Then
        If IsObject(suppliedIndexDefinitions) Then Set indexDefinitions = suppliedIndexDefinitions
    End If
    If schemaDefinitions Is Nothing Then Set schemaDefinitions = TerminologyFieldSchemaDefinitions()
    If indexDefinitions Is Nothing Then Set indexDefinitions = TerminologyIndexSchemaDefinitions()
    AddTerminologyTableSchemaIssues issues, schemaDefinitions, indexDefinitions
    If issues.count > 0 Then
        repairSummary = SharedCollection_Join(issues, vbCrLf)
        TerminologySchemaRepairRequired = True
    Else
        repairSummary = vbNullString
    End If
    Exit Function
ErrHandler:
    repairSummary = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_TERMINOLOGY_TABLE_LOGICAL_ID) & ": schema validation failed before B1 repair: " & Err.Description
    TerminologySchemaRepairRequired = True
End Function

Private Sub EnsureTerminologyIndexes(ByVal tableName As String, ByVal indexDefinitions As Collection)
    Dim db As DAO.Database
    Dim indexDefinition As Variant
    Set db = TablesWorkspaceCodeDb()
    For Each indexDefinition In indexDefinitions
        EnsureSingleFieldIndex db, tableName, CStr(indexDefinition(0)), CStr(indexDefinition(1)), CBool(indexDefinition(2)), CBool(indexDefinition(3)), mTableDesignBatchPrepared
    Next indexDefinition
End Sub

Private Sub AddIndexDefinitionIssue(ByRef issues As Collection, ByVal tableName As String, ByVal indexName As String, ByVal fieldName As String, ByVal uniqueIndex As Boolean, ByVal primaryIndex As Boolean, Optional ByVal suppliedTableDef As Variant)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    If CodeDbTableExists(tableName) = False Then Exit Sub
    'Keep the parent DAO.Database alive while the TableDef/Index objects are inspected.
    'A2 final verification reuses the batch-boundary DAO metadata snapshot instead of synchronizing TableDefs, Fields, and Indexes once per index. Non-batch callers retain the prior refresh behavior.
    If isMissing(suppliedTableDef) = False Then
        If IsObject(suppliedTableDef) Then Set tdf = suppliedTableDef
    End If
    If tdf Is Nothing Then
        Set db = TablesWorkspaceCodeDb()
        If mTableDesignBatchPrepared = False Then db.TableDefs.Refresh
        Set tdf = db.TableDefs(tableName)
        If mTableDesignBatchPrepared = False Then tdf.fields.Refresh
        If mTableDesignBatchPrepared = False Then tdf.Indexes.Refresh
    End If
    If Len(Trim$(fieldName)) = 0 Then
        If SharedDao_TableDefIndexExistsSafe(tdf, indexName) = False Then
            issues.Add tableName & ": missing index [" & indexName & "]"
        End If
    ElseIf SharedDao_TableDefSingleFieldIndexMatches(tdf, indexName, fieldName, uniqueIndex, primaryIndex) = False Then
        issues.Add tableName & ": missing or invalid index [" & indexName & "] on field [" & fieldName & "]"
    End If
    Exit Sub
ErrHandler:
    issues.Add tableName & ": could not validate index [" & indexName & "] - " & "Err.Number " & CStr(Err.Number) & "; Err.Source " & Err.Source & "; Err.Description " & Err.Description
End Sub

Private Function TerminologyReviewSql() As String
    Dim sqlText As String
    sqlText = "SELECT " & SQL_FIELD_ID & ", " & SQL_FIELD_TERM_KEY & ", " & SQL_FIELD_TERM_TEXT & ", " & SQL_FIELD_DEFINITION & ", " & SQL_FIELD_TERM_CATEGORY & ", " & SQL_FIELD_TERM_STATUS & ", " & SQL_FIELD_DEPRECATED & ", "
    sqlText = sqlText & SQL_FIELD_APPROVED_USAGE & ", " & SQL_FIELD_DISALLOWED_USAGE & ", " & SQL_FIELD_REPLACEMENT_GUIDANCE & ", " & SQL_FIELD_EXAMPLE_GOOD & ", " & SQL_FIELD_EXAMPLE_BAD & ", "
    sqlText = sqlText & SQL_FIELD_DECISION_SCOPE & ", " & SQL_FIELD_DECISION_STATUS & ", " & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER & ", " & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_NOTES & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_TERMINOLOGY_TABLE_LOGICAL_ID) & "] "
    sqlText = sqlText & "ORDER BY " & OrderByPolicyText("TABLES_TERMINOLOGY_REVIEW") & ";"
    TerminologyReviewSql = sqlText
End Function

Private Function DuplicateReviewSql() As String
    Dim sourceTable As String
    Dim sqlText As String
    sourceTable = "[" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "]"
    sqlText = "SELECT cfg." & SQL_FIELD_ID & ", " & SharedDao_SqlQuote(ROW_KIND_CONFIGURATION) & " AS " & SQL_FIELD_ROW_KIND & ", cfg." & SQL_FIELD_MACRO_KEY & ", Null AS " & SQL_FIELD_CANDIDATE_MACRO_KEY & ", cfg." & SQL_FIELD_MACRO_TITLE & ", cfg." & SQL_FIELD_CONFIGURATION_STATUS & ", Null AS " & SQL_FIELD_CANDIDATE_STATUS & ", "
    sqlText = sqlText & "cfg." & SQL_FIELD_MANUAL_REVIEW_STATUS & ", cfg." & SQL_FIELD_REVIEW_PRIORITY & ", cfg.[" & FIELD_CONFIG_OWNER_DATABASE_ROLE & "], Null AS [" & FIELD_CANDIDATE_OWNER_DATABASE_ROLE & "], "
    sqlText = sqlText & "cfg." & SQL_FIELD_TARGET_PROCEDURE_MODULE_NAME & ", cfg." & SQL_FIELD_TARGET_PROCEDURE_NAME & ", cfg." & SQL_FIELD_TARGET_ARGUMENTS & " AS " & FIELD_TARGET_ARGUMENT_SIGNATURE & ", Null AS " & SQL_FIELD_CANDIDATE_ARGUMENT_SIGNATURE & ", cfg." & SQL_FIELD_GENERATED_MACRO_OBJECT_NAME & ", "
    sqlText = sqlText & "DuplicateIssueTypeForQuery(cfg." & SQL_FIELD_ID & ") AS " & FIELD_DUPLICATE_ISSUE & ", "
    sqlText = sqlText & "DuplicateIssueDetailForQuery(cfg." & SQL_FIELD_ID & ") AS " & FIELD_DUPLICATE_ISSUE_DETAIL & ", "
    sqlText = sqlText & "DuplicateGroupKeyForQuery(cfg." & SQL_FIELD_ID & ") AS " & FIELD_DUPLICATE_GROUP_KEY & ", "
    sqlText = sqlText & "DuplicateGroupCountForQuery(cfg." & SQL_FIELD_ID & ") AS " & FIELD_DUPLICATE_GROUP_COUNT & ", "
    sqlText = sqlText & "DuplicateSelectedCountForQuery(cfg." & SQL_FIELD_ID & ") AS " & FIELD_DUPLICATE_SELECTED_COUNT & ", "
    sqlText = sqlText & "DuplicateAllRowsSelectedForQuery(cfg." & SQL_FIELD_ID & ") AS " & FIELD_ALL_ROWS_IN_GROUP_SELECTED_FOR_DELETE & ", "
    sqlText = sqlText & "cfg." & SQL_FIELD_SELECTED_FOR_DUPLICATE_DELETION & ", cfg." & SQL_FIELD_LAST_VALIDATED_ON & ", cfg." & SQL_FIELD_LAST_VALIDATION_RESULT & ", cfg." & SQL_FIELD_LAST_VALIDATION_NOTES & ", cfg." & SQL_FIELD_LAST_GENERATED_ON & ", cfg." & SQL_FIELD_LAST_GENERATED_RESULT & ", cfg." & SQL_FIELD_LAST_GENERATED_NOTES & " "
    sqlText = sqlText & "FROM " & sourceTable & " AS cfg WHERE ConfigRowHasDuplicateForQuery(cfg." & SQL_FIELD_ID & ")<>False "
    sqlText = sqlText & "ORDER BY " & OrderByPolicyText("TABLES_DUPLICATE_REVIEW") & ";"
    DuplicateReviewSql = sqlText
End Function

Private Function CustomMacroLauncherReviewSql() As String
    CustomMacroLauncherReviewSql = "SELECT " & OrderedTableSelectList(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LauncherControlsReview")) & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LauncherControlsReview") & "] ORDER BY " & OrderByPolicyText("TABLES_LAUNCHER_CONTROLS_REVIEW") & ";"
End Function

Private Function TableCanOpenSql(ByVal tableName As String) As Boolean
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    If CodeDbTableExists(tableName) = False Then Exit Function
    Set db = TablesWorkspaceCodeDb()
    Set rs = db.OpenRecordset("SELECT TOP 1 * FROM [" & tableName & "]", dbOpenSnapshot)
    TableCanOpenSql = True
ExitHere:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Err.Clear
    Exit Function
ErrHandler:
    TableCanOpenSql = False
    Resume ExitHere
End Function

' Purpose: Attempts to perform delete query def for DB, query name.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   db (Input; DAO.Database): DAO.Database instance whose schema, data, queries, or properties the routine reads or modifies.
'   queryName (Input; String): Physical Access query name identifying the saved query the routine creates, validates, reads, or updates.
'   errorText (Output; String): Receives the resolved error text produced from this procedure's validated inputs and current ADC state.
'   errorNumberResult (Output Optional; Long): Receives the error number result produced by the operation so the caller can inspect the outcome without re-running the work. Default: 0.
'   errorSourceResult (Output Optional; String): Receives the error source result produced by the operation so the caller can inspect the outcome without re-running the work. Default: vbNullString.
'   errorDescriptionResult (Output Optional; String): Receives the error description result produced by the operation so the caller can inspect the outcome without re-running the work. Default: vbNullString.
' Returns: Boolean: True when the attempt to delete query def completes successfully; False for the routine's defined normal unsuccessful/not-applicable outcome.
Public Function TryDeleteQueryDef(ByVal db As DAO.Database, ByVal queryName As String, ByRef errorText As String, Optional ByRef errorNumberResult As Long = 0, Optional ByRef errorSourceResult As String = vbNullString, Optional ByRef errorDescriptionResult As String = vbNullString) As Boolean
    On Error GoTo ErrHandler
    errorNumberResult = 0
    errorSourceResult = vbNullString
    errorDescriptionResult = vbNullString
    db.QueryDefs.Delete queryName
    db.QueryDefs.Refresh
    TryDeleteQueryDef = True
    Exit Function
ErrHandler:
    errorNumberResult = Err.Number
    errorSourceResult = Err.Source
    errorDescriptionResult = Err.Description
    If Err.Number = 3265 Or Err.Number = 3078 Or Err.Number = 3011 Or Err.Number = 7874 Then
        errorText = "Query definition was already unavailable. Err.Number: " & CStr(Err.Number) & "; " & SharedMessage_OneLine(Err.Description, 250)
    Else
        errorText = "Could not delete QueryDef. Err.Number: " & CStr(Err.Number) & "; Err.Source: " & Err.Source & "; Err.Description: " & Err.Description
    End If
    Err.Clear
End Function

Private Sub DefaultTerminologyRows()
    Dim defaultRow As Variant
    Dim exampleGoodText As String
    Dim exampleBadText As String
    For Each defaultRow In TerminologyDefinitions()
        exampleGoodText = vbNullString
        exampleBadText = vbNullString
        If UBound(defaultRow) >= 13 Then exampleGoodText = CStr(defaultRow(13))
        If UBound(defaultRow) >= 14 Then exampleBadText = CStr(defaultRow(14))
        DefaultTerminologyRow CStr(defaultRow(0)), CStr(defaultRow(1)), CStr(defaultRow(2)), CStr(defaultRow(3)), CBool(defaultRow(4)), CStr(defaultRow(5)), CStr(defaultRow(6)), CStr(defaultRow(7)), CStr(defaultRow(8)), CStr(defaultRow(9)), CStr(defaultRow(10)), CLng(defaultRow(11)), CStr(defaultRow(12)), exampleGoodText, exampleBadText
    Next defaultRow
End Sub

Private Sub DefaultTerminologyRow(ByVal termKey As String, ByVal termText As String, ByVal termStatus As String, ByVal termCategory As String, ByVal deprecatedValue As Boolean, ByVal definitionText As String, ByVal approvedUsageText As String, ByVal disallowedUsageText As String, ByVal replacementGuidanceText As String, ByVal decisionScopeText As String, ByVal decisionStatusText As String, ByVal SortOrder As Long, Optional ByVal notesText As String = vbNullString, Optional ByVal exampleGoodText As String = vbNullString, Optional ByVal exampleBadText As String = vbNullString)
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim sqlText As String
    If Len(termKey) = 0 Then Exit Sub
    sqlText = "SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_TERMINOLOGY_TABLE_LOGICAL_ID) & "] WHERE " & SQL_FIELD_TERM_KEY & "=" & SharedDao_SqlQuote(termKey) & ";"
    Set rs = TablesWorkspaceCodeDb().OpenRecordset(sqlText, dbOpenDynaset)
    If rs.EOF = False Then
        rs.Close
        Exit Sub
    End If
    rs.AddNew
    rs.fields(FIELD_TERM_KEY).value = termKey
    rs.fields(FIELD_DEFAULTED_ON).value = Now
    rs.fields(FIELD_TERM_TEXT).value = termText
    rs.fields(FIELD_TERM_STATUS).value = termStatus
    rs.fields(FIELD_TERM_CATEGORY).value = termCategory
    rs.fields(FIELD_DEPRECATED).value = deprecatedValue
    rs.fields(FIELD_DEFINITION).value = definitionText
    rs.fields(FIELD_APPROVED_USAGE).value = approvedUsageText
    rs.fields(FIELD_DISALLOWED_USAGE).value = disallowedUsageText
    rs.fields(FIELD_REPLACEMENT_GUIDANCE).value = IIf(Len(Trim$(replacementGuidanceText)) = 0, Null, replacementGuidanceText)
    rs.fields(FIELD_DECISION_SCOPE).value = decisionScopeText
    rs.fields(FIELD_DECISION_STATUS).value = decisionStatusText
    rs.fields(FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER).value = SortOrder
    rs.fields(FIELD_EXAMPLE_GOOD).value = IIf(Len(exampleGoodText) = 0, Null, exampleGoodText)
    rs.fields(FIELD_EXAMPLE_BAD).value = IIf(Len(exampleBadText) = 0, Null, exampleBadText)
    If Len(Trim$(notesText)) > 0 Then rs.fields(FIELD_MANAGED_OBJECT_DEPENDENCY_NOTES).value = notesText
    rs.Update
    rs.Close
    Exit Sub
ErrHandler:
    Dim handlerErrNumber As Long
    Dim handlerErrDescription As String
    handlerErrNumber = Err.Number
    handlerErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
    Err.Raise handlerErrNumber, "DefaultTerminologyRow", "Unable to insert a missing " & DEFAULT_UTILITY_LONG_NAME & " terminology row without changing existing rows." & vbCrLf & "TermKey: " & termKey & vbCrLf & "Table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_TERMINOLOGY_TABLE_LOGICAL_ID) & vbCrLf & "Err.Description: " & handlerErrDescription
End Sub

Private Function CustomMacroLauncherReviewSchemaDefinitions() As Collection
    'Single physical field-schema authority for MC_Generated_LauncherCtrls_Review creation and repair.
    ' CANONICAL AUTHORITY: rows owns the in-memory definition set constructed by CustomMacroLauncherReviewSchemaDefinitions; downstream consumers derive views from this set rather than create a competing authority.
    Dim rows As New Collection
    Dim tableName As String
    tableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LauncherControlsReview")
    rows.Add Array(FIELD_ID, dbLong, 0, True)
    rows.Add Array(FIELD_REVIEW_BATCH_ID, dbText, 100, False)
    rows.Add Array(FIELD_REVIEWED_ON, dbDate, 0, False)
    rows.Add Array(FIELD_SEVERITY, dbText, 30, False, Null, Null, Null, Null, Null)
    rows.Add Array(FIELD_ISSUE_TYPE, dbText, 100, False, Null, Null, Null, Null, Null)
    rows.Add Array(FIELD_CONFIG_ID, dbLong, 0, False)
    rows.Add Array(FIELD_MACRO_KEY, dbText, 100, False)
    rows.Add Array(FIELD_MACRO_TITLE, dbText, 255, False)
    rows.Add Array(FIELD_GENERATED_MACRO_OBJECT_NAME, dbText, 255, False)
    rows.Add Array(FIELD_CONTROL_NAME, dbText, 255, False)
    rows.Add Array(FIELD_CONTROL_CAPTION, dbText, 255, False)
    rows.Add Array(FIELD_EXPECTED_CAPTION, dbText, 255, False)
    rows.Add Array(FIELD_EXPECTED_ON_CLICK, dbMemo, 0, False)
    rows.Add Array(FIELD_ACTUAL_ON_CLICK, dbMemo, 0, False)
    rows.Add Array(FIELD_EXPECTED_TOOLTIP, dbText, ManagedFieldReqSizeForSetup(tableName, FIELD_EXPECTED_TOOLTIP), False)
    rows.Add Array(FIELD_ACTUAL_TOOLTIP, dbText, ManagedFieldReqSizeForSetup(tableName, FIELD_ACTUAL_TOOLTIP), False)
    rows.Add Array(FIELD_FINDING_NOTES, dbMemo, 0, False)
    Set CustomMacroLauncherReviewSchemaDefinitions = rows
End Function

' Purpose: Ensures one authoritative managed field Description, rejects blank/placeholder definitions, and independently verifies the persisted DAO type and value. G1 is the only numbered-Setup mutation owner.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   fld (Input; DAO.Field): DAO.Field representing the field whose properties or metadata are being inspected or modified.
'   descriptionText (Input; String): Description text text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   tableName (Input Optional; String): Physical Access table name identifying the table whose schema, data, or metadata the routine processes. Default: vbNullString.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Private Function FieldDescriptionFailureSourceName(ByVal tableName As String, ByVal fieldName As String, ByVal descriptionText As String) As String
    'Identifies the exact current DEFAULTS constant/builder when an overlength description still equals its canonical default; otherwise identifies the persisted/runtime authority.
    Dim defaultDescription As String
    Dim resolutionSource As Variant
    If Len(tableName) > 0 And Len(fieldName) > 0 Then
        resolutionSource = vbNullString
        defaultDescription = ManagedObjectCatalogFieldDescriptionShort(tableName, fieldName, resolutionSource)
        If Len(defaultDescription) > 0 And StrComp(defaultDescription, descriptionText, vbBinaryCompare) = 0 And Len(Trim$(CStr(resolutionSource))) > 0 Then
            FieldDescriptionFailureSourceName = CStr(resolutionSource)
            Exit Function
        End If
    End If
    FieldDescriptionFailureSourceName = "(not applicable; persisted/runtime value does not match the current DEFAULTS default)"
End Function

Public Sub EnsureDaoFieldDescription(ByVal fld As DAO.Field, ByVal descriptionText As String, Optional ByVal tableName As String = vbNullString)
    'Ensures one authoritative managed field Description, rejects blank/placeholder definitions, and independently verifies the persisted DAO type and value. G1 is the only numbered-Setup mutation owner.
    On Error GoTo ErrHandler
    Dim currentDescription As String
    Dim currentPropertyType As Long
    Dim propertyExists As Boolean
    Dim writerChanged As Boolean
    Dim repairOutcome As Variant
    Dim targetContext As String
    Dim fieldName As String
    Dim verificationField As DAO.Field
    Dim verificationDatabase As DAO.Database
    fieldName = SafeDAOFieldName(fld)
    targetContext = "Object type: DAO.Field" & vbCrLf & "Table: " & IIf(Len(tableName) = 0, "(not supplied)", tableName) & vbCrLf & "Field: " & fieldName
    descriptionText = Trim$(descriptionText)
    If Len(descriptionText) = 0 Then Err.Raise vbObjectError + 6890, "EnsureDaoFieldDescription", "No authoritative field description exists." & vbCrLf & "Field: " & SafeDAOFieldName(fld)
    If FieldDescriptionTextIsPlaceholder(descriptionText) Then Err.Raise vbObjectError + 6891, "EnsureDaoFieldDescription", "The field description is an unmapped or placeholder definition and cannot be applied as authoritative metadata." & vbCrLf & "Field: " & SafeDAOFieldName(fld) & vbCrLf & "Description: " & descriptionText
    If Len(descriptionText) > 255 Then Err.Raise vbObjectError + 6899, "EnsureDaoFieldDescription", "The authoritative field description exceeds the Access DAO Field.Description 255-character limit." & vbCrLf & "Table: " & IIf(Len(tableName) = 0, "(not supplied)", tableName) & vbCrLf & "Field: " & fieldName & vbCrLf & "Characters: " & CStr(Len(descriptionText)) & vbCrLf & "Default/value source: " & FieldDescriptionFailureSourceName(tableName, fieldName, descriptionText)
    ReadDaoFieldDescriptionState fld, currentDescription, currentPropertyType, propertyExists
    If propertyExists And currentPropertyType = CLng(dbText) Then
        If StrComp(currentDescription, descriptionText, vbBinaryCompare) = 0 Then Exit Sub
    End If
    If SetupDomainMutationWritesAllowedForCurrentPart(SETUP_DOMAIN_G1_DESCRIPTIONS) = False Then Err.Raise vbObjectError + 9623, "EnsureDaoFieldDescription", "Managed DAO field Description mutation is owned exclusively by Setup G1." & vbCrLf & "Table: " & IIf(Len(tableName) = 0, "(not supplied)", tableName) & vbCrLf & "Field: " & fieldName
    writerChanged = EnsureDaoPropertyValue(fld, DAO_DESCRIPTION_PROPERTY_NAME, dbText, descriptionText, "EnsureDaoFieldDescription", targetContext, repairOutcome)
    If StrComp(CStr(repairOutcome), "CapacityRecreated", vbTextCompare) = 0 Then AddTraceNote "Recreated an undersized DAO Description property after Access error 3309. Table: " & IIf(Len(tableName) = 0, "(not supplied)", tableName) & "; Field: " & SafeDAOFieldName(fld) & "; Requested characters: " & CStr(Len(descriptionText)) & "; Requested bytes: " & CStr(LenB(descriptionText))
    If writerChanged = False Then Err.Raise vbObjectError + 6892, "EnsureDaoFieldDescription", "The verified DAO property writer reported no change even though the existing field Description was not authoritative." & vbCrLf & "Field: " & SafeDAOFieldName(fld)
    currentDescription = vbNullString
    currentPropertyType = 0
    propertyExists = False
    Set verificationField = fld
    If Len(tableName) > 0 Then Set verificationField = ReacquireTableDefField(tableName, fieldName, verificationDatabase)
    ReadDaoFieldDescriptionState verificationField, currentDescription, currentPropertyType, propertyExists
    If propertyExists = False Then Err.Raise vbObjectError + 6893, "EnsureDaoFieldDescription", "The field Description property was absent after the write completed." & vbCrLf & "Field: " & SafeDAOFieldName(fld)
    If currentPropertyType <> CLng(dbText) Then Err.Raise vbObjectError + 6894, "EnsureDaoFieldDescription", "The field Description property type did not verify after the write." & vbCrLf & "Field: " & SafeDAOFieldName(fld) & vbCrLf & "Expected DAO type: " & CStr(dbText) & vbCrLf & "Actual DAO type: " & CStr(currentPropertyType)
    If StrComp(currentDescription, descriptionText, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 6895, "EnsureDaoFieldDescription", "The field Description property value did not verify after the write." & vbCrLf & "Field: " & SafeDAOFieldName(fld) & vbCrLf & "Expected: " & descriptionText & vbCrLf & "Actual: " & currentDescription
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrSource As String
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrSource = Err.Source
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "EnsureDaoFieldDescription", "Could not set and verify the authoritative Access field Description property." & vbCrLf & "Table: " & IIf(Len(tableName) = 0, "(not supplied)", tableName) & vbCrLf & "Field: " & SafeDAOFieldName(fld) & vbCrLf & "Requested characters: " & CStr(Len(descriptionText)) & vbCrLf & "Requested bytes: " & CStr(LenB(descriptionText)) & vbCrLf & "Property previously existed: " & CStr(propertyExists) & vbCrLf & "Existing DAO type: " & IIf(propertyExists, CStr(currentPropertyType), "(absent)") & vbCrLf & "Err.Number: " & CStr(originalErrNumber) & vbCrLf & "Err.Source: " & originalErrSource & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

' Purpose: Reads one field Description without converting a real DAO failure into a missing-property result.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   fld (Input; DAO.Field): DAO.Field representing the field whose properties or metadata are being inspected or modified.
'   descriptionText (Output; String): Receives the resolved description text produced from this procedure's validated inputs and current ADC state.
'   propertyType (Output; Long): Receives the property type returned by `CLng` after this procedure completes its validation/resolution work.
'   propertyExists (Output; Boolean): Set True when the procedure determines that property exists applies; set False when that condition is not met.
' Returns: None (Sub). Results are returned through the documented ByRef parameters and/or the routine's database, UI, or module-state side effects.
Public Sub ReadDaoFieldDescriptionState(ByVal fld As DAO.Field, ByRef descriptionText As String, ByRef propertyType As Long, ByRef propertyExists As Boolean)
    'Reads one field Description without converting a real DAO failure into a missing-property result.
    On Error GoTo ErrHandler
    Dim prp As DAO.Property
    Dim lookupErrNumber As Long
    Dim lookupErrSource As String
    Dim lookupErrDescription As String
    descriptionText = vbNullString
    propertyType = 0
    propertyExists = False
    If fld Is Nothing Then Err.Raise vbObjectError + 6896, "ReadDaoFieldDescriptionState", "The DAO field is Nothing."
    On Error Resume Next
    Err.Clear
    Set prp = fld.properties(DAO_DESCRIPTION_PROPERTY_NAME)
    lookupErrNumber = Err.Number
    lookupErrSource = Err.Source
    lookupErrDescription = Err.Description
    Err.Clear
    On Error GoTo ErrHandler
    If lookupErrNumber = 3270 Then
        Exit Sub
    End If
    If lookupErrNumber <> 0 Then Err.Raise lookupErrNumber, lookupErrSource, lookupErrDescription
    propertyExists = True
    propertyType = CLng(prp.Type)
    If IsNull(prp.value) = False Then descriptionText = CStr(prp.value)
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ReadDaoFieldDescriptionState", "Could not read the Access field Description property state." & vbCrLf & "Field: " & SafeDAOFieldName(fld) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

' Purpose: Performs field description text is placeholder by coordinating `End` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   descriptionText (Input; String): Description text text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
' Returns: Boolean: True when the condition represented by NormalizeLookupColumnWidthsForAccess is present under the routine's documented inputs/current state; False when that condition is absent.
Public Function FieldDescriptionTextIsPlaceholder(ByVal descriptionText As String) As Boolean
    'Intended use: Provides the single physical-field Description placeholder policy used by table repair and catalog validation.
    Dim normalizedText As String
    normalizedText = UCase$(Trim$(descriptionText))
    If Len(normalizedText) = 0 Then FieldDescriptionTextIsPlaceholder = True: Exit Function
    If Left$(normalizedText, 27) = "UNMAPPED FIELD DESCRIPTION:" Then FieldDescriptionTextIsPlaceholder = True: Exit Function
    If Left$(normalizedText, 29) = "DEVELOPMENT METADATA MISSING" Then FieldDescriptionTextIsPlaceholder = True: Exit Function
    If Left$(normalizedText, 12) = "PLACEHOLDER:" Then FieldDescriptionTextIsPlaceholder = True: Exit Function
    If Left$(normalizedText, 5) = "TODO:" Or normalizedText = "TODO" Then FieldDescriptionTextIsPlaceholder = True: Exit Function
    If Left$(normalizedText, 4) = "TBD:" Or normalizedText = "TBD" Then FieldDescriptionTextIsPlaceholder = True: Exit Function
    If normalizedText = "DESCRIPTION UNAVAILABLE" Then FieldDescriptionTextIsPlaceholder = True: Exit Function
    If normalizedText = "UNKNOWN" Then FieldDescriptionTextIsPlaceholder = True
End Function

Private Sub RepairConfigIndexes(ByVal tableName As String, Optional ByVal showMessage As Boolean = True, Optional ByVal suppliedIndexDefinitions As Variant)
    'Intended use: Runs Repair Access Development Console Config Indexes for direct, generated-macro, form-button, or required cross-module callers.
    'CodeDb removes explicit child-field indexes that B5 canonical RI will own, preventing duplicate index-slot consumption while HostDb keeps the complete explicit performance set.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim indexDefinition As Variant
    Dim indexDefinitions As Collection
    Dim indexName As String
    If CodeDbTableExists(tableName) = False Then Exit Sub
    If isMissing(suppliedIndexDefinitions) = False Then
        If IsObject(suppliedIndexDefinitions) Then Set indexDefinitions = suppliedIndexDefinitions
    End If
    If indexDefinitions Is Nothing Then Set indexDefinitions = ConfigIndexSchemaDefinitions()
    Set db = TablesWorkspaceCodeDb()
    For Each indexDefinition In indexDefinitions
        indexName = CStr(indexDefinition(0))
        If ConfigIndexSupersededByCodeDbCanonicalRelationship(indexDefinition) Then
            Set tdf = db.TableDefs(tableName)
            If TableDefIndexExists(tdf, indexName) Then
                If mTableDesignBatchPrepared = False Then
                    If SharedAccess_PrepareOpenObjectForDesignChange(acTable, tableName, "removing redundant relationship-backed index [" & indexName & "] from table [" & tableName & "]") = False Then
                        MarkSystemCommandCancelled
                        Err.Raise vbObjectError + 10035, "RepairConfigIndexes", "Config-index repair was cancelled because the affected open table was not approved for save and close."
                    End If
                    Set db = TablesWorkspaceCodeDb()
                    Set tdf = db.TableDefs(tableName)
                End If
                tdf.Indexes.Delete indexName
                If mTableDesignBatchPrepared = False Then tdf.Indexes.Refresh
                CoreAddOperationNote "Removed redundant explicit index [" & indexName & "] from table [" & tableName & "]; B5 canonical RI owns the equivalent CodeDb child-field index."
            End If
        Else
            EnsureSingleFieldIndex db, tableName, indexName, CStr(indexDefinition(1)), CBool(indexDefinition(2)), CBool(indexDefinition(3)), mTableDesignBatchPrepared
        End If
    Next indexDefinition
    ReportCoreActionSuccess "RepairConfigIndexes", "TableRepair", DEFAULT_UTILITY_SHORT_NAME & " config index repair completed successfully.", DEFAULT_UTILITY_LONG_NAME & " config index repair complete." & vbCrLf & vbCrLf & "Table: " & tableName, "Repair config indexes", tableName, "Table", "CodeDb", "TableRepair", showMessage, "Repair Config Indexes", vbInformation, True, "Tables", False, visibleStatusMessage:="Config index repair completed successfully."
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim messageText As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    messageText = "Could not create or repair macro-creator config indexes." & vbCrLf & vbCrLf & "Procedure: RepairConfigIndexes" & vbCrLf & "Table: " & tableName & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description:" & vbCrLf & savedErrDescription & TableReadinessDiagnosticsForMessage() & TraceBackendDiagnosticsForMessage()
    LogTraceMessage TRACE_LEVEL_ERROR, "TableRepair", "RepairConfigIndexes", "Config index repair failed.", messageText, "Repair config indexes", savedErrNumber, savedErrSource, savedErrDescription, tableName, "Table", "CodeDb", "TableRepair"
    If showMessage Then
        CoreShowErrorMessage messageText, vbCritical, "Repair Config Indexes", "RepairConfigIndexes", "Tables"
        Exit Sub
    End If
    Err.Raise savedErrNumber, "RepairConfigIndexes", messageText
End Sub

' Purpose: Normalizes lookup widths for current Microsoft Access behavior, especially when LimitToList=False makes a hidden bound column invalid or problematic.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   columnWidthsText (Input; String): Input column widths text for normalize lookup column widths for access; is assigned to `NormalizeLookupColumnWidthsForAccess`; is passed to `Split`.
'   columnCount (Input; Long): Supplies the column count used to bound, report, or validate the number of items processed by this operation.
'   BoundColumn (Input; Long): One-based Access lookup column number whose value is stored as the control/field value.
'   LimitToList (Input; Boolean): True restricts the Access lookup control/field to values present in its row source; False permits values outside the listed choices when Access allows them.
' Returns: String: Returns the normalized lookup column widths for access text after applying exactly the canonicalization rules implemented by this helper; vbNullString is preserved/returned when no usable
'   text remains.
Public Function NormalizeLookupColumnWidthsForAccess(ByVal columnWidthsText As String, ByVal columnCount As Long, ByVal BoundColumn As Long, ByVal LimitToList As Boolean) As String
    'Normalizes lookup widths for current Microsoft Access behavior, especially when LimitToList=False makes a hidden bound column invalid or problematic.
    Dim widths As Variant
    Dim item As Variant
    Dim normalizedText As String
    Dim i As Long
    Dim widthValue As Double
    Dim boundWidth As Double
    columnWidthsText = Trim$(columnWidthsText)
    If LimitToList Or columnCount <= 1 Or BoundColumn < 1 Or BoundColumn > columnCount Then
        NormalizeLookupColumnWidthsForAccess = columnWidthsText
        Exit Function
    End If
    widths = Split(columnWidthsText, ";")
    For i = LBound(widths) To UBound(widths)
        PumpSetupControlEvents
        item = Trim$(CStr(widths(i)))
        widthValue = Val(CStr(item))
        If i + 1 = BoundColumn Then boundWidth = widthValue
    Next i
    If boundWidth <= 0 Then
        If UBound(widths) < BoundColumn - 1 Then ReDim Preserve widths(0 To BoundColumn - 1)
        widths(BoundColumn - 1) = "1.25"
    End If
    For i = LBound(widths) To UBound(widths)
        PumpSetupControlEvents
        If Len(normalizedText) > 0 Then normalizedText = normalizedText & ";"
        normalizedText = normalizedText & Trim$(CStr(widths(i)))
    Next i
    NormalizeLookupColumnWidthsForAccess = normalizedText
End Function

Private Function OpenRecordsetFailureDetail(ByVal tableName As String, ByVal useSql As Boolean) As String
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sourceText As String
    If useSql Then
        sourceText = "SELECT TOP 1 * FROM [" & tableName & "]"
    Else
        sourceText = tableName
    End If
    Set db = TablesWorkspaceCodeDb()
    Set rs = db.OpenRecordset(sourceText, dbOpenSnapshot)
ExitHere:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Err.Clear
    Exit Function
ErrHandler:
    OpenRecordsetFailureDetail = "Err.Number " & CStr(Err.Number) & ", Source " & Err.Source & ", Description " & Err.Description
    Resume ExitHere
End Function

Private Function ObjectDiagnosticLine(ByVal objectName As String) As String
    On Error GoTo ErrHandler
    Dim tableDefExistsValue As Boolean
    Dim queryDefExistsValue As Boolean
    Dim allTablesExistsValue As Boolean
    Dim allQueriesExistsValue As Boolean
    Dim canOpenDirect As Boolean
    Dim canOpenSql As Boolean
    Dim detailText As String
    tableDefExistsValue = CodeDbTableExists(objectName)
    queryDefExistsValue = SharedDao_QueryExists(TablesWorkspaceCodeDb(), objectName)
    If CodeDbIsHostDb() Then
        allTablesExistsValue = AccessObjectExistsInCurrentDatabase(acTable, objectName)
        allQueriesExistsValue = AccessObjectExistsInCurrentDatabase(acQuery, objectName)
    End If
    canOpenDirect = TableCanOpen(objectName)
    canOpenSql = TableCanOpenSql(objectName)
    detailText = TableOpenFailureDetail(objectName)
    ObjectDiagnosticLine = " - [" & objectName & "] " & "TableDef=" & CStr(tableDefExistsValue) & "; " & "QueryDef=" & CStr(queryDefExistsValue) & "; " & "HostDb.AllTables=" & CStr(allTablesExistsValue) & "; " & "HostDb.AllQueries=" & CStr(allQueriesExistsValue) & "; " & "OpenByName=" & CStr(canOpenDirect) & "; " & "OpenBySql=" & CStr(canOpenSql) & "; " & "Detail=" & SharedMessage_OneLine(detailText, 900)
    Exit Function
ErrHandler:
    ObjectDiagnosticLine = " - [" & objectName & "] diagnostic failed. Err.Number: " & CStr(Err.Number) & "; Err.Source: " & Err.Source & "; Err.Description: " & Err.Description
End Function

Private Function ObjectNameConflictExists(ByVal objectName As String) As Boolean
    On Error GoTo ErrHandler
    If CodeDbTableExists(objectName) Then
        ObjectNameConflictExists = True
        Exit Function
    End If
    If SharedDao_QueryExists(TablesWorkspaceCodeDb(), objectName) Then
        ObjectNameConflictExists = True
        Exit Function
    End If
    If CodeDbIsHostDb() Then
        If AccessObjectExistsInCurrentDatabase(acTable, objectName) Or AccessObjectExistsInCurrentDatabase(acQuery, objectName) Then
            ObjectNameConflictExists = True
        End If
    End If
    Exit Function
ErrHandler:
    ObjectNameConflictExists = False
End Function

' Purpose: Attempts to perform delete access object for object type, object name.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   objectType (Input; AcObjectType): Object-type value identifying the Access/VBA object class targeted by the operation.
'   objectName (Input; String): Physical Access object name identifying the database object targeted by the operation.
'   errorText (Output; String): Receives the resolved error text produced from this procedure's validated inputs and current ADC state.
'   errorNumberResult (Output Optional; Long): Receives the error number result produced by the operation so the caller can inspect the outcome without re-running the work. Default: 0.
'   errorSourceResult (Output Optional; String): Receives the error source result produced by the operation so the caller can inspect the outcome without re-running the work. Default: vbNullString.
'   errorDescriptionResult (Output Optional; String): Receives the error description result produced by the operation so the caller can inspect the outcome without re-running the work. Default: vbNullString.
' Returns: Boolean: True when the attempt to delete access object completes successfully; False for the routine's defined normal unsuccessful/not-applicable outcome.
Public Function TryDeleteAccessObject(ByVal objectType As AcObjectType, ByVal objectName As String, ByRef errorText As String, Optional ByRef errorNumberResult As Long = 0, Optional ByRef errorSourceResult As String = vbNullString, Optional ByRef errorDescriptionResult As String = vbNullString) As Boolean
    On Error GoTo ErrHandler
    errorNumberResult = 0
    errorSourceResult = vbNullString
    errorDescriptionResult = vbNullString
    If CodeDbIsHostDb() = False Then
        errorNumberResult = vbObjectError + 6710
        errorSourceResult = "TryDeleteAccessObject"
        errorDescriptionResult = "DoCmd.DeleteObject is active-project scoped and CodeDb differs from HostDb."
        errorText = "Could not delete Access object because " & errorDescriptionResult & " Object: " & objectName
        Exit Function
    End If
    DoCmd.DeleteObject objectType, objectName
    TryDeleteAccessObject = True
    Exit Function
ErrHandler:
    errorNumberResult = Err.Number
    errorSourceResult = Err.Source
    errorDescriptionResult = Err.Description
    If Err.Number = 3265 Or Err.Number = 3078 Or Err.Number = 3011 Or Err.Number = 7874 Or Err.Number = 2501 Then
        errorText = "Access object was already unavailable. Err.Number: " & CStr(Err.Number) & "; " & SharedMessage_OneLine(Err.Description, 250)
    Else
        errorText = "Could not delete Access object. Err.Number: " & CStr(Err.Number) & "; Err.Source: " & Err.Source & "; Err.Description: " & Err.Description
    End If
    Err.Clear
End Function

' Purpose: Checks whether form design query foundation already satisfies the exact state required by the current Setup run so the owning phase can determine whether repair is necessary.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   detailText (Output; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
' Returns: Boolean: True when the condition represented by BackupAllUserDataCore is present under the routine's documented inputs/current state; False when that condition is absent.
Public Function FormDesignQueryFoundationCurrentForSetup(ByRef detailText As String) As Boolean
    On Error GoTo NotCurrent
    Dim snapshot As FoundationDomainSnapshot
    Dim db As DAO.Database
    Dim storedSourceFingerprint As String
    Dim sourceExists As Boolean
    Dim sourceMatches As Boolean
    detailText = vbNullString
    Set db = TablesWorkspaceCodeDb()
    storedSourceFingerprint = DatabaseTextPropertyValue(db, FOUNDATION_FORMDESIGN_SOURCE_FINGERPRINT_PROPERTY, sourceExists)
    If sourceExists = False Or Len(storedSourceFingerprint) = 0 Then
        detailText = "A3 retained form-design source-query fingerprint: source=unseeded; physical comparison skipped because required persisted proof is unavailable. Retired FormDesign_* FieldMap/Tab/Rule/Layout saved-query families are not physical Setup requirements."
        Exit Function
    End If
    If CalculateFoundationQuerySnapshot(snapshot) = False Then
        detailText = snapshot.failureDetail
        Exit Function
    End If
    sourceMatches = (StrComp(storedSourceFingerprint, snapshot.FormDesignSourceFingerprint, vbBinaryCompare) = 0)
    FormDesignQueryFoundationCurrentForSetup = sourceMatches
    detailText = "A3 retained form-design source-query fingerprint: source=" & IIf(sourceMatches, "matched", "changed") & ". Retired FormDesign_* FieldMap/Tab/Rule/Layout saved-query families are not physical Setup requirements."
    Exit Function
NotCurrent:
    detailText = "A3 retained form-design source-query current-state verification failed. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
End Function

' ===============================================================================================================================================================================
' CONSOLIDATED TABLE / DATABASE FOUNDATION PROCEDURES
' ===============================================================================================================================================================================

' Purpose: Performs the public backup all user data operation using the routine's existing ADC workflow and side effects.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine shows its user-facing completion, status, warning, or error message in addition to performing the underlying work. Default: True.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub BackupAllUserData(Optional ByVal showMessage As Boolean = True)
    ' Operational System Macro #73. Uses the one Defaults-owned user-data table scope; no independent table list is permitted here.
    On Error GoTo ErrHandler
    Dim progressStarted As Boolean
    Dim messageText As String
    Dim backupCount As Long
    LogTraceMessage "Info", "Backup", "BackupAllUserData", "ADC user-data backup started.", vbNullString, vbNullString, 0, vbNullString, vbNullString, vbNullString, "UserDataTableSet", "CodeDb", "Backup"
    progressStarted = CoreBeginOperationProgressIfNeeded("Backup all " & DEFAULT_UTILITY_SHORT_NAME & " user data", 3)
    CorePulseProgress "Check default Navigation Pane group"
    RequireDefaultCodeDbGeneratedMacrosNavigationPaneGroupForRoleContext "BackupAllUserData", False
    CorePulseProgress "Back up authoritative " & DEFAULT_UTILITY_SHORT_NAME & " user data"
    backupCount = BackupAllUserDataCore(False)
    CorePulseProgress "Report user-data backup completion"
    messageText = "ADC user-data backup complete." & vbCrLf & vbCrLf & "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB & vbCrLf & "Tables backed up: " & CStr(backupCount) & CoreOperationSummaryForMessage() & TableDiagnosticsLocationForMessage() & TraceBackendDiagnosticsForMessage()
    ReportCoreActionSuccess "BackupAllUserData", "Backup", DEFAULT_UTILITY_SHORT_NAME & " user-data backup completed successfully.", messageText, "Back up all ADC user data", vbNullString, "UserDataTableSet", "CodeDb", "Backup", showMessage, TABLE_MAINTENANCE_MESSAGE_TITLE, vbInformation, True, "Tables", False, visibleStatusMessage:="User-data backup completed successfully."
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
    messageText = "Could not back up all ADC user data." & vbCrLf & vbCrLf & "Procedure: BackupAllUserData" & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description:" & vbCrLf & savedErrDescription & TableDiagnosticsLocationForMessage() & CoreOperationSummaryForMessage()
    LogTraceMessage "Error", "Backup", "BackupAllUserData", "ADC user-data backup failed.", messageText, vbNullString, savedErrNumber, savedErrSource, savedErrDescription, vbNullString, "UserDataTableSet", "CodeDb", "Backup"
    If showMessage Then
        CoreShowErrorMessage messageText, vbCritical, TABLE_MAINTENANCE_MESSAGE_TITLE, "BackupAllUserData", "Tables"
        Exit Sub
    End If
    Err.Raise savedErrNumber, "BackupAllUserData", messageText
End Sub

' Purpose: Resolves backup all user data core from the supplied/current ADC state, using `StrComp` and `CodeDbTableExists` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine shows its user-facing completion, status, warning, or error message in addition to performing the underlying work. Default: True.
' Returns: Long: Returns the numeric backup all user data core result; zero and other sentinel values retain the meanings established by the routine's existing implementation.
Public Function BackupAllUserDataCore(Optional ByVal showMessage As Boolean = True) As Long
    ' One authoritative Backup implementation for every currently existing table in UserDataTableScopeDefinitions().
    On Error GoTo ErrHandler
    Dim backupStamp As String
    Dim backupCount As Long
    Dim scopeDefinition As Variant
    Dim tableName As String
    Dim backupName As String
    Dim details As String
    backupStamp = Format$(Now, "yyyymmdd_hhnnss")
    For Each scopeDefinition In UserDataTableScopeDefinitions()
        tableName = CStr(scopeDefinition(1))
        If StrComp(CStr(scopeDefinition(3)), "FullDataCopy", vbTextCompare) <> 0 Then Err.Raise vbObjectError + 6233, "BackupAllUserDataCore", "Unsupported canonical ADC user-data backup behavior." & vbCrLf & "Logical identity: " & CStr(scopeDefinition(0)) & vbCrLf & "Backup behavior: " & CStr(scopeDefinition(3))
        If CodeDbTableExists(tableName) Then
            If TableCanOpen(tableName, True) = False Then Err.Raise vbObjectError + 6230, "BackupAllUserDataCore", "An existing authoritative ADC user-data table cannot be opened, so its required backup could not be confirmed." & vbCrLf & "Table: " & tableName & vbCrLf & TableOpenFailureDetail(tableName)
            backupName = BackupOneTable(tableName, backupStamp)
            If Len(Trim$(backupName)) = 0 Then Err.Raise vbObjectError + 6231, "BackupAllUserDataCore", "Backup creation returned no backup identity." & vbCrLf & "Table: " & tableName
            backupCount = backupCount + 1
            If Len(details) > 0 Then details = details & vbCrLf
            details = details & tableName & " -> " & backupName
        End If
    Next scopeDefinition
    BackupAllUserDataCore = backupCount
    If showMessage Then ReportCoreActionSuccess "BackupAllUserDataCore", "Backup", "ADC user-data backup completed successfully.", "Tables backed up: " & CStr(backupCount) & IIf(Len(details) = 0, vbNullString, vbCrLf & vbCrLf & details), "Back up all ADC user data", vbNullString, "UserDataTableSet", "CodeDb", "Backup", True, "Backup All ADC User Data", vbInformation, True, "Tables", False
    Exit Function
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrSource As String
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrSource = Err.Source
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "BackupAllUserDataCore", "Could not back up all ADC user data." & vbCrLf & "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB & vbCrLf & "Err.Source: " & originalErrSource & vbCrLf & "Err.Description: " & originalErrDescription
End Function

' Purpose: Clears ADC user/operational data through the established reset path while preserving the schema and managed setup definitions required for subsequent use.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
'   confirmDestructiveAction (Input Optional; Boolean): When True, enables or requests destructive action for this call; when False, the routine follows its normal non-forced/default path. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub ResetAllUserData(Optional ByVal showMessage As Boolean = True, Optional ByVal confirmDestructiveAction As Boolean = True)
    ' Operational System Macro #72. Backup and destructive reset consume the exact same Defaults-owned scope.
    On Error GoTo ErrHandler
    Dim progressStarted As Boolean
    Dim messageText As String
    Dim backupCount As Long
    progressStarted = CoreBeginOperationProgressIfNeeded("Reset all " & DEFAULT_UTILITY_SHORT_NAME & " user data", OPERATION_PROGRESS_DEFAULT_MAX_STEPS)
    CorePulseProgress "Confirm and back up authoritative ADC user-data scope"
    backupCount = ResetAllUserDataCore(False, confirmDestructiveAction)
    If backupCount = -1 Then
        If progressStarted Then CoreCompleteSetupProgress
        Exit Sub
    End If
    messageText = "ADC user-data reset complete." & vbCrLf & vbCrLf & "Confirmed backup tables created before reset: " & CStr(backupCount) & vbCrLf & "The same authoritative user-data scope was used for Backup and Reset." & CoreOperationSummaryForMessage() & TableDiagnosticsLocationForMessage() & TraceBackendDiagnosticsForMessage()
    ReportCoreActionSuccess "ResetAllUserData", "Reset", DEFAULT_UTILITY_SHORT_NAME & " user-data reset completed successfully.", messageText, "Reset all ADC user data", vbNullString, "UserDataTableSet", "CodeDb", "Reset", showMessage, TABLE_MAINTENANCE_MESSAGE_TITLE, vbInformation, True, "Tables", False, visibleStatusMessage:="User-data reset completed successfully."
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
    messageText = "Could not reset all ADC user data." & vbCrLf & vbCrLf & "Procedure: ResetAllUserData" & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description:" & vbCrLf & savedErrDescription & TableDiagnosticsLocationForMessage() & CoreOperationSummaryForMessage()
    LogTraceMessage "Error", "Reset", "ResetAllUserData", "ADC user-data reset failed.", messageText, vbNullString, savedErrNumber, savedErrSource, savedErrDescription, vbNullString, "UserDataTableSet", "CodeDb", "Reset"
    If showMessage Then
        CoreShowErrorMessage messageText, vbCritical, TABLE_MAINTENANCE_MESSAGE_TITLE, "ResetAllUserData", "Tables"
        Exit Sub
    End If
    Err.Raise savedErrNumber, "ResetAllUserData", messageText
End Sub

' Purpose: Repairs and verifies the CodeDb table foundation and managed non-form saved-query definitions for runtime/manual use, including schema evidence, query currentness, and optional standalone Description
'   reconciliation.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
'   schemaRepairDetailsForMessage (Output Optional; String): Receives schema repair details for message used for caller diagnostics, verification output, or user-facing reporting. Default:
'   vbNullString.
'   applyStandaloneDescriptions (Input Optional; Boolean): True reconciles authoritative Descriptions for the A2/A3 objects owned by this direct/runtime operation before success; False lets a broader caller
'   own a later complete Description pass. Default: True.
' Returns: Boolean: True on the function's affirmative/success branch for repair tables; False on its explicit negative, unavailable, rejected, or handled-failure branch.
Public Function RepairTables(Optional ByVal showMessage As Boolean = True, Optional ByRef schemaRepairDetailsForMessage As String = vbNullString, Optional ByVal applyStandaloneDescriptions As Boolean = True) As Boolean
    'Runtime/manual orchestrator: TABLES owns physical readiness plus complete non-form saved-query definition, verification, fingerprint, and repair authority.
    On Error GoTo ErrHandler
    Dim coreFoundationDetails As String
    Dim queryRepairDetails As String
    Dim schemaRepairSummary As String
    Dim messageText As String
    Dim traceRunOwnedByProcedure As Boolean
    Dim progressStarted As Boolean
    Dim queryFoundationCurrent As Boolean
    Dim schemaRepairNeeded As Boolean
    Dim backupCount As Long
    Dim lookupBackupCount As Long
    Dim namingInputsChanged As Boolean
    Dim schemaRepairEvidence As TSchemaRepairPreflightEvidence
    Dim queryStateDetail As String
    Dim fastSkipStarted As Single
    Dim objectDescriptionChanges As Long
    Dim fieldDescriptionChanges As Long
    Dim descriptionDetail As String
    schemaRepairDetailsForMessage = vbNullString
    CoreSetTableRepairWarningResult False
    If TraceRunActive() = False Then
        StartTraceRun "TableRepair"
        traceRunOwnedByProcedure = True
    End If
    progressStarted = CoreBeginOperationProgressIfNeeded("Repair " & DEFAULT_UTILITY_SHORT_NAME & " table/query foundation", OPERATION_PROGRESS_DEFAULT_MAX_STEPS)
    LogTraceMessage "Info", "Tables", "RepairTables", "Access Development Console runtime table/query foundation repair started.", vbNullString, vbNullString, 0, vbNullString, vbNullString, vbNullString, "FoundationSet", "CodeDb", "TableRepair"
    CorePulseProgress "Check canonical TABLES structural readiness and backup need."
    schemaRepairNeeded = TableSchemaRepairRequired(schemaRepairSummary)
    schemaRepairEvidence.IsKnown = True
    schemaRepairEvidence.repairNeeded = schemaRepairNeeded
    schemaRepairEvidence.Summary = schemaRepairSummary
    If schemaRepairNeeded Then
        RequireDefaultCodeDbGeneratedMacrosNavigationPaneGroupForRoleContext "RepairTables", False
        backupCount = BackupTablesForSchemaRepair(schemaRepairSummary, lookupBackupCount)
        CoreAddOperationNote "Runtime table repair created " & CStr(backupCount) & " automatic backup table(s), including " & CStr(lookupBackupCount) & " affected editable lookup table(s), before canonical TABLES structural repair."
    End If
    CorePulseProgress "Ask TABLES to verify or repair the core physical table/default foundation."
    CoreBeginProgressChildDetail "Canonical Core Table Foundation"
    If RepairCoreDataTableFoundationForSetup(schemaRepairEvidence, False, coreFoundationDetails, namingInputsChanged, False, False) = False Then Err.Raise vbObjectError + 9048, "RepairTables", "The canonical TABLES core data-table foundation repair returned False."
    CoreEndProgressChildDetail
    schemaRepairDetailsForMessage = coreFoundationDetails
    CorePulseProgress "Ask TABLES to verify the complete non-form saved-query foundation."
    fastSkipStarted = Timer
    queryFoundationCurrent = NonFormQueryFoundationCurrentForSetup(queryStateDetail)
    If queryFoundationCurrent Then
        schemaRepairDetailsForMessage = coreFoundationDetails & vbCrLf & vbCrLf & "Runtime Part 3 query-foundation fast path: Used" & vbCrLf & queryStateDetail & vbCrLf & "Query-foundation verification elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(fastSkipStarted))
        If applyStandaloneDescriptions Then
            CorePulseProgress "Reconcile standalone TABLES-owned descriptions"
            ReconcileRepairTablesStandaloneDescriptions objectDescriptionChanges, fieldDescriptionChanges, descriptionDetail
            schemaRepairDetailsForMessage = SharedText_AppendNonBlankLine(schemaRepairDetailsForMessage, descriptionDetail)
        End If
        messageText = "ADC table/query foundation is ready and TABLES-owned saved-query foundation matched."
        messageText = messageText & schemaRepairDetailsForMessage & CoreOperationSummaryForMessage() & TableDiagnosticsLocationForMessage() & TraceBackendDiagnosticsForMessage()
        RepairTables = True
        CoreSetTableRepairWarningResult CoreOperationHasWarnings()
        ReportCoreActionSuccess "RepairTables", "Tables", DEFAULT_UTILITY_SHORT_NAME & " table/query foundation verified successfully.", messageText, "Verify TABLES-owned non-form query foundation", vbNullString, "FoundationNonFormQuerySet", "CodeDb", "TableRepair", showMessage, "Repair Access Development Console Tables", vbInformation, True, "Tables", False, visibleStatusMessage:="Table/query foundation verified successfully."
        If progressStarted Then CoreCompleteSetupProgress
        If traceRunOwnedByProcedure Then CompleteTraceRun "ADC table/query foundation and TABLES-owned saved-query foundation verified."
        Exit Function
    End If
    If Len(Trim$(queryStateDetail)) > 0 Then CoreAddOperationNote "TABLES-owned non-form query repair required. " & queryStateDetail
    CorePulseProgress "Ask TABLES to repair the complete non-form saved-query foundation."
    CoreBeginProgressChildDetail "TABLES Non-Form Query Foundation"
    If RepairNonFormQueryFoundationForSetup(False, queryRepairDetails, False) = False Then Err.Raise vbObjectError + 9047, "RepairTables", "The TABLES-owned non-form query-foundation repair returned False during runtime foundation repair."
    CoreEndProgressChildDetail
    If Len(Trim$(queryRepairDetails)) > 0 Then CoreAddOperationNote queryRepairDetails
    schemaRepairDetailsForMessage = coreFoundationDetails & vbCrLf & vbCrLf & "Runtime Part 3 query-foundation path: Repair" & vbCrLf & "Initial TABLES-owned query state: " & IIf(Len(Trim$(queryStateDetail)) = 0, "(not recorded)", queryStateDetail) & queryRepairDetails & vbCrLf & "Query-foundation verification and repair elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(fastSkipStarted))
    If applyStandaloneDescriptions Then
        CorePulseProgress "Reconcile standalone TABLES-owned descriptions"
        ReconcileRepairTablesStandaloneDescriptions objectDescriptionChanges, fieldDescriptionChanges, descriptionDetail
        schemaRepairDetailsForMessage = SharedText_AppendNonBlankLine(schemaRepairDetailsForMessage, descriptionDetail)
    End If
    messageText = "ADC table/query foundation is ready and TABLES-owned saved-query foundation repaired."
    messageText = messageText & schemaRepairDetailsForMessage & CoreOperationSummaryForMessage() & TableDiagnosticsLocationForMessage() & TraceBackendDiagnosticsForMessage()
    RepairTables = True
    CoreSetTableRepairWarningResult CoreOperationHasWarnings()
    ReportCoreActionSuccess "RepairTables", "Tables", DEFAULT_UTILITY_SHORT_NAME & " table/query foundation repair completed successfully.", messageText, "Repair TABLES-owned non-form query foundation", vbNullString, "FoundationNonFormQuerySet", "CodeDb", "TableRepair", showMessage, "Repair Access Development Console Tables", vbInformation, True, "Tables", False, visibleStatusMessage:="Table/query foundation repair completed successfully."
    If progressStarted Then CoreCompleteSetupProgress
    If traceRunOwnedByProcedure Then CompleteTraceRun "ADC table/query foundation ready; TABLES-owned saved-query foundation repaired."
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    RepairTables = False
    If progressStarted Then CoreFailSetupProgress
    messageText = "Access Development Console table/query foundation repair failed." & vbCrLf & vbCrLf & _
        "Procedure: RepairTables" & vbCrLf & "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB & vbCrLf & "Database: " & TableSafeDatabaseName() & vbCrLf & _
        "MC_Records table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & vbCrLf & "CodeDb settings table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID) & vbCrLf & _
        "TABLES query-state detail: " & IIf(Len(Trim$(queryStateDetail)) = 0, "(not recorded)", queryStateDetail) & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & _
        "Err.Description:" & vbCrLf & savedErrDescription & TableDiagnosticsLocationForMessage() & TableReadinessDiagnosticsForMessage() & TraceSummaryForMessage() & CoreOperationSummaryForMessage() & CoreAppendProcessRunSummary(vbNullString, "RepairTables")
    LogTraceMessage "Error", "Tables", "RepairTables", "Access Development Console table/query foundation repair failed.", messageText, vbNullString, savedErrNumber, savedErrSource, savedErrDescription, vbNullString, "FoundationSet", "CodeDb", "TableRepair"
    If showMessage Then
        CoreShowErrorMessage messageText, vbCritical, TABLE_MAINTENANCE_MESSAGE_TITLE, "RepairTables", "Tables"
    Else
        If traceRunOwnedByProcedure Then CompleteTraceRun "Access Development Console table/query foundation repair failed.", False
        Err.Raise savedErrNumber, "RepairTables", messageText
    End If
    If traceRunOwnedByProcedure Then CompleteTraceRun "Access Development Console table/query foundation repair failed.", False
End Function

Private Sub ReconcileStandaloneManagedTableDescriptionIfAllowed(ByVal tableName As String, Optional ByVal ignoreDesignBatch As Boolean = False)
    'Direct A2 helpers may complete the managed table metadata they own outside numbered Setup. Nested A2 batch calls skip this per-table pass and use one complete foundation pass at the coordinator boundary.
    Dim definitionData As Variant
    Dim descriptionChanged As Variant
    Dim fieldDescriptionChangeCount As Variant
    Dim fieldDetail As String
    If mTableDesignBatchPrepared And ignoreDesignBatch = False Then Exit Sub
    If SetupDomainMutationWritesAllowedForCurrentPart(SETUP_DOMAIN_G1_DESCRIPTIONS) = False Then Exit Sub
    If TryGetManagedObjectDefinition(tableName, ADCAT_OBJECT_TYPE_TABLE, definitionData) = False Then Exit Sub
    descriptionChanged = False
    If ApplyAndVerifyManagedObjectDescriptionForStandaloneRepair(tableName, ADCAT_OBJECT_TYPE_TABLE, descriptionChanged) = False Then Err.Raise vbObjectError + 10131, "ReconcileStandaloneManagedTableDescriptionIfAllowed", "Managed-table Description reconciliation returned False." & vbCrLf & "Table: " & tableName
    fieldDescriptionChangeCount = CLng(0)
    fieldDetail = vbNullString
    If ApplyAndVerifyManagedTableFieldDescriptionsForStandaloneRepair(tableName, fieldDescriptionChangeCount, fieldDetail, True) = False Then Err.Raise vbObjectError + 10132, "ReconcileStandaloneManagedTableDescriptionIfAllowed", fieldDetail
End Sub

Private Sub ReconcileStandaloneFoundationTableDescriptions()
    'The direct complete A2 API owns the entire TABLES foundation roster; this one pass covers both its fast path and its full structural-repair path without moving Description writes into A2 during numbered Setup.
    Dim tableNames As Collection
    Dim item As Variant
    If SetupDomainMutationWritesAllowedForCurrentPart(SETUP_DOMAIN_G1_DESCRIPTIONS) = False Then Exit Sub
    Set tableNames = FoundationRequiredTableNames()
    For Each item In tableNames
        ReconcileStandaloneManagedTableDescriptionIfAllowed CStr(item), True
    Next item
End Sub

Private Sub ReconcileStandaloneManagedQueryDescriptionIfAllowed(ByVal queryName As String, Optional ByVal ignoreQueryBatch As Boolean = False)
    'Single-query repair completes canonical managed metadata only when it is a true standalone call. A3 batch repair defers to one complete query-roster pass at its outer boundary.
    Dim definitionData As Variant
    Dim descriptionChanged As Variant
    If mQueryRepairBatchActive And ignoreQueryBatch = False Then Exit Sub
    If SetupDomainMutationWritesAllowedForCurrentPart(SETUP_DOMAIN_G1_DESCRIPTIONS) = False Then Exit Sub
    If TryGetManagedObjectDefinition(queryName, ADCAT_OBJECT_TYPE_QUERY, definitionData) = False Then Exit Sub
    descriptionChanged = False
    If ApplyAndVerifyManagedObjectDescriptionForStandaloneRepair(queryName, ADCAT_OBJECT_TYPE_QUERY, descriptionChanged) = False Then Err.Raise vbObjectError + 10133, "ReconcileStandaloneManagedQueryDescriptionIfAllowed", "Managed-query Description reconciliation returned False." & vbCrLf & "Query: " & queryName
End Sub

Private Sub ReconcileStandaloneNonFormQueryDescriptions()
    'The direct complete A3 API owns the full non-form saved-query roster; numbered Setup A3 skips this physical metadata pass and leaves Description mutation to G1.
    Dim queryNames As Collection
    Dim item As Variant
    If SetupDomainMutationWritesAllowedForCurrentPart(SETUP_DOMAIN_G1_DESCRIPTIONS) = False Then Exit Sub
    Set queryNames = FoundationRequiredNonFormQueryNames()
    For Each item In queryNames
        ReconcileStandaloneManagedQueryDescriptionIfAllowed CStr(item), True
    Next item
End Sub

Private Sub ReconcileRepairTablesStandaloneDescriptions(ByRef objectDescriptionChanges As Long, ByRef fieldDescriptionChanges As Long, ByRef detailText As String)
    'Standalone completion uses the same authoritative A2 table and A3 non-form query rosters consumed by TABLES structural repair; it does not broaden into Catalog, Reference, Module Guard, form, or report
    'ownership.
    On Error GoTo ErrHandler
    Dim tableNames As Collection
    Dim queryNames As Collection
    Dim item As Variant
    Dim descriptionChanged As Variant
    Dim fieldDescriptionChangeCount As Variant
    Dim fieldDetail As String
    objectDescriptionChanges = 0
    fieldDescriptionChanges = 0
    detailText = vbNullString
    Set tableNames = FoundationRequiredTableNames()
    For Each item In tableNames
        descriptionChanged = False
        If ApplyAndVerifyManagedObjectDescriptionForStandaloneRepair(CStr(item), ADCAT_OBJECT_TYPE_TABLE, descriptionChanged) = False Then Err.Raise vbObjectError + 10120, "ReconcileRepairTablesStandaloneDescriptions", "TABLES-owned managed-table Description reconciliation returned False." & vbCrLf & "Table: " & CStr(item)
        If CBool(descriptionChanged) Then objectDescriptionChanges = objectDescriptionChanges + 1
        fieldDescriptionChangeCount = CLng(0)
        fieldDetail = vbNullString
        If ApplyAndVerifyManagedTableFieldDescriptionsForStandaloneRepair(CStr(item), fieldDescriptionChangeCount, fieldDetail, True) = False Then Err.Raise vbObjectError + 10121, "ReconcileRepairTablesStandaloneDescriptions", fieldDetail
        fieldDescriptionChanges = fieldDescriptionChanges + CLng(Nz(fieldDescriptionChangeCount, 0))
    Next item
    Set queryNames = FoundationRequiredNonFormQueryNames()
    For Each item In queryNames
        descriptionChanged = False
        If ApplyAndVerifyManagedObjectDescriptionForStandaloneRepair(CStr(item), ADCAT_OBJECT_TYPE_QUERY, descriptionChanged) = False Then Err.Raise vbObjectError + 10122, "ReconcileRepairTablesStandaloneDescriptions", "TABLES-owned managed-query Description reconciliation returned False." & vbCrLf & "Query: " & CStr(item)
        If CBool(descriptionChanged) Then objectDescriptionChanges = objectDescriptionChanges + 1
    Next item
    detailText = "Standalone TABLES-owned managed descriptions reconciled." & vbCrLf & "Object descriptions changed: " & CStr(objectDescriptionChanges) & vbCrLf & "Field descriptions changed: " & CStr(fieldDescriptionChanges)
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    detailText = "Standalone TABLES-owned Description reconciliation failed." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
    Err.Raise savedErrNumber, "ReconcileRepairTablesStandaloneDescriptions", detailText
End Sub

Private Sub ResetUserDataTableByContract(ByVal scopeDefinition As Variant)
    ' Destructive reset is explicit by canonical ownership contract. This prevents mixed/system-user stores from falling through an accidental blanket-delete path.
    Dim logicalIdentity As String
    Dim tableName As String
    Dim authoritativeOwner As String
    Dim resetContract As String
    logicalIdentity = CStr(scopeDefinition(0))
    tableName = CStr(scopeDefinition(1))
    authoritativeOwner = CStr(scopeDefinition(2))
    resetContract = CStr(scopeDefinition(4))
    Select Case UCase$(resetContract)
        Case "CORETABLE", "EDITABLELOOKUPTABLE", "LOOKUPDEFINITIONREGISTRY"
            If StrComp(authoritativeOwner, "TABLES", vbTextCompare) <> 0 Then Err.Raise vbObjectError + 6234, "ResetUserDataTableByContract", "TABLES-owned reset contract has the wrong authoritative owner." & vbCrLf & logicalIdentity
            SafeDropTableIfExists tableName, "Reset TABLES-owned ADC user-data table [" & tableName & "] under its authoritative reset contract."
        Case "CATALOGTABLE"
            If StrComp(authoritativeOwner, "CATALOG", vbTextCompare) <> 0 Then Err.Raise vbObjectError + 6235, "ResetUserDataTableByContract", "Catalog-owned reset contract has the wrong authoritative owner." & vbCrLf & logicalIdentity
            ResetCatalogUserDataTableForOperationalReset logicalIdentity, tableName
        Case Else
            Err.Raise vbObjectError + 6236, "ResetUserDataTableByContract", "Unsupported ADC user-data reset contract." & vbCrLf & "Logical identity: " & logicalIdentity & vbCrLf & "Table: " & tableName & vbCrLf & "Reset contract: " & resetContract
    End Select
End Sub

Private Function ResetAllUserDataCore(Optional ByVal showMessage As Boolean = True, Optional ByVal confirmDestructiveAction As Boolean = True) As Long
    On Error GoTo ErrHandler
    CoreResetOperationNotes
    Dim scopeDefinitions As Collection
    Dim scopeDefinition As Variant
    Dim backupCount As Long
    Dim requiredBackupCount As Long
    Dim promptText As String
    Dim confirmationText As String
    Dim coreFoundationDetails As String
    Dim queryRepairDetails As String
    Dim catalogDetails As String
    Dim namingInputsChanged As Boolean
    Dim schemaRepairEvidence As TSchemaRepairPreflightEvidence
    Set scopeDefinitions = UserDataTableScopeDefinitions()
    If confirmDestructiveAction Then
        promptText = "This will reset all authoritative ADC user-maintained/persistent data and restore each scoped store to its canonical initial/default state." & vbCrLf & vbCrLf & _
                     "Every existing table in the reset scope will be backed up first. Generated/transient diagnostic and inventory data outside the user-data classification is not part of this command." & vbCrLf & vbCrLf & _
                     "To continue, type:" & vbCrLf & "RESET ALL ADC USER DATA"
        confirmationText = InputBox(promptText, "Reset All ADC User Data")
        If StrComp(Trim$(confirmationText), "RESET ALL ADC USER DATA", vbBinaryCompare) <> 0 Then
            CoreAddOperationNote "Reset cancelled. No ADC user data was changed."
            If showMessage Then CoreShowMessage "Reset cancelled. No ADC user data was changed.", vbInformation, "Reset All ADC User Data", "ShowADCMessage", "Tables"
            MarkSystemCommandCancelled
            ResetAllUserDataCore = -1
            Exit Function
        End If
    End If
    For Each scopeDefinition In scopeDefinitions
        If CodeDbTableExists(CStr(scopeDefinition(1))) Then requiredBackupCount = requiredBackupCount + 1
    Next scopeDefinition
    backupCount = BackupAllUserDataCore(False)
    If backupCount <> requiredBackupCount Then Err.Raise vbObjectError + 6232, "ResetAllUserDataCore", "Reset was stopped because every existing table in the authoritative reset scope must have a confirmed backup before destructive work begins." & vbCrLf & "Existing scoped tables: " & CStr(requiredBackupCount) & vbCrLf & "Confirmed backups: " & CStr(backupCount)
    For Each scopeDefinition In scopeDefinitions
        ResetUserDataTableByContract scopeDefinition
    Next scopeDefinition
    If RepairCoreDataTableFoundationForSetup(schemaRepairEvidence, False, coreFoundationDetails, namingInputsChanged, False) = False Then Err.Raise vbObjectError + 9049, "ResetAllUserDataCore", "The canonical TABLES core foundation owner returned False after reset."
    If RepairNonFormQueryFoundationForSetup(False, queryRepairDetails) = False Then Err.Raise vbObjectError + 9050, "ResetAllUserDataCore", "The canonical TABLES query owner returned False after reset."
    If RepairCatalogTableAndQueryFamilies(False, catalogDetails) = False Then Err.Raise vbObjectError + 9051, "ResetAllUserDataCore", catalogDetails
    If Len(Trim$(coreFoundationDetails)) > 0 Then CoreAddOperationNote coreFoundationDetails
    If Len(Trim$(queryRepairDetails)) > 0 Then CoreAddOperationNote queryRepairDetails
    If Len(Trim$(catalogDetails)) > 0 Then CoreAddOperationNote catalogDetails
    ResetAllUserDataCore = backupCount
    Exit Function
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrSource As String
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrSource = Err.Source
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "ResetAllUserDataCore", "Could not reset all ADC user data." & vbCrLf & "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB & vbCrLf & "Err.Source: " & originalErrSource & vbCrLf & "Err.Description: " & originalErrDescription
End Function

' Purpose: Prepares open objects for managed name update for code DB, host DB.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   codeDb (Input; DAO.Database): DAO.Database reference for the database context that this operation reads or updates.
'   hostDb (Input; DAO.Database): DAO.Database reference for the database context that this operation reads or updates.
'   renames (Input; Collection): Collection containing the renames items that the routine consumes or updates as part of this operation.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub PrepareOpenObjectsForManagedNameUpdate(ByVal codeDb As DAO.Database, ByVal hostDb As DAO.Database, ByVal renames As Collection)
    'Managed-object physical names come from current configured naming rules. This workflow prepares open objects so current physical names can be reconciled safely with that authority; stored references are
    'repaired, results are verified by the owning rename workflow, and failures are rolled back/recovered there. This is not historical database-version migration. Do not add hardcoded historical rename maps or
    'compatibility aliases.
    ' INTENTIONAL FILTERED PARTITION: tableNames, queryNames is a consumer-required subset/projection derived from the source enumerated by PrepareOpenObjectsForManagedNameUpdate; its membership is not independently
    ' authoritative.
    Dim tableNames As New Collection
    Dim queryNames As New Collection
    Dim item As Variant
    Dim qdf As DAO.queryDef
    Dim repairedSql As String
    If DatabaseIsCurrentProject(codeDb) Then
        For Each item In renames
            PumpSetupResponsiveness
            If StrComp(CStr(item(2)), "Table", vbTextCompare) = 0 Then
                AddUniqueObjectName tableNames, CStr(item(0))
                AddUniqueObjectName tableNames, CStr(item(1))
            End If
            If StrComp(CStr(item(2)), "Query", vbTextCompare) = 0 Then
                AddUniqueObjectName queryNames, CStr(item(0))
                AddUniqueObjectName queryNames, CStr(item(1))
            End If
        Next item
        For Each qdf In codeDb.QueryDefs
            PumpSetupResponsiveness
            If Left$(qdf.Name, 1) <> "~" Then
                repairedSql = x__UTIL_ADC_MOC_LIFECYCLE.ApplyManagedObjectRenameMap(qdf.sql, renames)
                If StrComp(qdf.sql, repairedSql, vbBinaryCompare) <> 0 Then AddUniqueObjectName queryNames, qdf.Name
            End If
        Next qdf
        If SharedAccess_PrepareOpenObjectsForDesignChange(acQuery, queryNames, "renaming managed queries and repairing saved-query references") = False Then
            MarkSystemCommandCancelled
            Err.Raise vbObjectError + 8995, "PrepareOpenObjectsForManagedNameUpdate", "Managed query rename/update was cancelled."
        End If
        If SharedAccess_PrepareOpenObjectsForDesignChange(acTable, tableNames, "renaming managed tables and repairing stored object references") = False Then
            MarkSystemCommandCancelled
            Err.Raise vbObjectError + 8996, "PrepareOpenObjectsForManagedNameUpdate", "Managed table rename/update was cancelled."
        End If
    End If
    If DatabaseIsCurrentProject(hostDb) Then
        ' INTENTIONAL FILTERED PARTITION: tableNames is a consumer-required subset/projection derived from the source enumerated by PrepareOpenObjectsForManagedNameUpdate; its membership is not independently
        ' authoritative.
        Set tableNames = New Collection
        AddUniqueObjectName tableNames, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config")
        AddUniqueObjectName tableNames, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Candidates")
        If SharedAccess_PrepareOpenObjectsForDesignChange(acTable, tableNames, "updating the HostDb MC_Records/MC_Candidates table names and stored references") = False Then
            MarkSystemCommandCancelled
            Err.Raise vbObjectError + 8997, "PrepareOpenObjectsForManagedNameUpdate", "HostDb MC_Records/MC_Candidates rename/update was cancelled."
        End If
    End If
End Sub

Private Sub AddUniqueObjectName(ByRef names As Collection, ByVal objectName As String)
    If Len(objectName) = 0 Then Exit Sub
    On Error Resume Next
    names.Add objectName, UCase$(objectName)
    On Error GoTo 0
End Sub

' Purpose: Repairs launcher foundation for config metadata ready, settings defaults already ready.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
'   configMetadataReady (Input Optional; Boolean): Boolean state describing whether config metadata ready applies to this invocation; True selects the named condition/path and False selects its
'   complementary/default path. Default: False.
'   settingsDefaultsAlreadyReady (Input Optional; Boolean): Boolean state describing whether settings defaults already ready applies to this invocation; True selects the named condition/path and False selects
'   its complementary/default path. Default: False.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub RepairLauncherFoundation(Optional ByVal showMessage As Boolean = True, Optional ByVal configMetadataReady As Boolean = False, Optional ByVal settingsDefaultsAlreadyReady As Boolean = False)
    'Launcher foundation orchestration: TABLES owns physical MC_Records/Settings/launcher readiness and the launcher review saved-query definition/repair; this procedure retains orchestration/reporting only.
    On Error GoTo ErrHandler
    RepairLauncherTableFoundationForSetup False, configMetadataReady, settingsDefaultsAlreadyReady
    EnsureCustomMacroLauncherReviewQueryForSetup
    ReportCoreActionSuccess "RepairLauncherFoundation", "Launcher", "Custom Macro Launcher automation foundation repair completed successfully.", "Custom Macro Launcher automation foundation is ready." & vbCrLf & vbCrLf & "MC_Records table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & vbCrLf & "Launcher review table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LauncherControlsReview") & vbCrLf & "Launcher review query: " & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_LauncherControls_Review"), "Repair Custom Macro Launcher foundation", ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LauncherControlsReview"), "Table", "CodeDb", "Launcher", showMessage, "Custom Macro Launcher Foundation", vbInformation, True, "Tables", False
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim messageText As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    messageText = "Could not ensure Custom Macro Launcher automation foundation." & vbCrLf & vbCrLf & "Procedure: RepairLauncherFoundation" & vbCrLf & "MC_Records table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & vbCrLf & "Launcher review table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LauncherControlsReview") & vbCrLf & "Launcher review query: " & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_LauncherControls_Review") & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description:" & vbCrLf & savedErrDescription & TableReadinessDiagnosticsForMessage() & TraceBackendDiagnosticsForMessage()
    LogTraceMessage "Error", "Launcher", "RepairLauncherFoundation", "Custom Macro Launcher automation foundation repair failed.", messageText, "Repair Custom Macro Launcher foundation", savedErrNumber, savedErrSource, savedErrDescription, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LauncherControlsReview"), "Table", "CodeDb", "LauncherFoundation"
    If showMessage Then
        CoreShowErrorMessage messageText, vbCritical, "Custom Macro Launcher Foundation - Error Detail", "RepairLauncherFoundation", "Tables"
        Exit Sub
    End If
    Err.Raise savedErrNumber, "RepairLauncherFoundation", messageText
End Sub
