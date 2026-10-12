Attribute VB_Name = "x__UTIL_ADC_COMMON"
' ================================================================================================================================================================================
' ChatGPT Rewrite Tracking Header
' ================================================================================================================================================================================
' ChatGPT Previous Version ID: CGPT-V2WE-CLTY-4NJD-G523-2YNS
' ChatGPT Baseline Version ID: CGPT-V2WE-CLTY-4NJD-G523-2YNS
' ChatGPT Version ID: CGPT-UWHT-Y8H3-273C-WNGX-LD79
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
' STANDARD MODULE: Access Development Console Common
' VB_Name: x__UTIL_ADC_COMMON
' Entry Points: None; this module is support-only in the Phase-A build.
' SHARED CODE:
'     SharedAccess_PrepareOpenObjectForDesignChange(ByVal objectType As AcObjectType, ByVal objectName As String, ByVal actionText As String, Optional ByVal confirmationAlreadyProvided As Boolean = False) -> Boolean
'     SharedCollection_KeyExists(ByVal values As Collection, ByVal key As String) -> Boolean
'     SharedDao_AccessObjectExists(ByVal db As DAO.Database, ByVal objectType As AcObjectType, ByVal objectName As String) -> Boolean
'     SharedDao_QueryExists(ByVal db As DAO.Database, ByVal queryName As String) -> Boolean
'     SharedDao_TableDefFieldExistsSafe(ByVal tdf As DAO.tableDef, ByVal fieldName As String) -> Boolean
'     SharedDao_TableDefIndexExistsSafe(ByVal tdf As DAO.tableDef, ByVal indexName As String) -> Boolean
'     SharedDao_TableDefSingleFieldIndexMatches(ByVal tdf As DAO.tableDef, ByVal indexName As String, ByVal fieldName As String, ByVal uniqueIndex As Boolean, ByVal primaryIndex As Boolean) -> Boolean
'     SharedText_Nz(ByVal v As Variant) -> String
'     SharedValue_NzLong(ByVal value As Variant, Optional ByVal defaultValue As Long = 0) -> Long
'     SharedVba_IsIdentifierChar(ByVal value As String) -> Boolean
' CLASSES USED: None
' DYNAMIC STRING REFERENCES:
'     Dynamic-reference scan stopped after 2000 logical line(s) to prevent a long-running scan; review this module manually if needed.
' LastScanTimestamp: Saturday 8/15/26 8:43 AM
' Purpose:
'     Provides the shared TMacroRecord record contract plus naming, diagnostics, DAO-property, popup, timing, stateless Settings Boolean interpretation, and generic physical DAO schema-mutation mechanics
'     required by the active Setup build, including the standard-Ribbon title setting reader consumed by Part G.
' Author: Christina Shonts
' Support:
'     Access Development Console project-specific common helpers only; this is not a general shared-code module.
' Code Category: Access Development Console common infrastructure.
' Requirements:
'     Microsoft Access TempVars, DAO, and VBA UtilityMsgBox constants. SharedCollection_KeyExists from the shared-code module must be available.
' Reference Dependencies: Access object model, DAO, and the shared-code module.
' DB Object Dependencies: None.
' Usage:
'     Import with the Access Development Console core, table-maintenance, and UI modules. Keep workflow state, tracing adapters, semantic progress ownership, and runtime database/session state in their owning
'     modules.
' Version:
'     Updates current Setup lifecycle references so standard-Ribbon settings are consumed by Phase H after insertion of the new dependency Phase F.
'     Removes remaining TargetObjectName Trim$ normalization from action signatures so configured Access object identity remains exact.
'     Moves domain-owned macro, Phase C/recovery, Reference/Add-In, UI, Ribbon, managed-object rename, and Defaults helpers to their established owner modules; shared/common behavior is retained.
'     Splits the shared row contracts so TMacroRecord/MacroRecordFromRecordset hydrate configuration-only MC_Records rows while TCandidateRecord/CandidateRecordFromRecordset hydrate MC_Candidates rows.
'     This prevents candidate-only state from leaking back into configured-macro runtime logic.
'     Removed the now-dead shared sentence-formatting helper after managed-object descriptions were changed to persist exactly as resolved without trimming or punctuation repair.
'     Replaces seven positional lookup-property Boolean switches with one named property mask and narrows module-internal procedure/constant visibility.
'     Updated shared lifecycle documentation so standard-Ribbon settings are consumed by Phase G under the reordered Setup architecture.
'     Managed-header dependency metadata updated because x__UTIL_ADC_DEFAULTS is now a standard module; executable behavior unchanged.
'     Removed Common-owned System Command result state and retained only the low-level popup/suppression primitives assigned to Common.
'     Centralized shared system-command cancellation-result interpretation beside the existing cancellation publisher.
'     Canonical terminology name resolution is object-owned and does not depend on a System Macro identity.
' Notes:
'     This module deliberately does not own Access Development Console logging, popup policy, setup orchestration, runtime database/session policy, or progress state. It owns only stateless database-role/path
'     normalization. SetupParts is a current normalized system-macro group key used by the dedicated setup-macro Navigation Pane metadata.
' Core Rules Enforced Here:
'     Common helpers must remain project-specific, side-effect limited, and independent of module-private constants.
' Design Map:
'     TMacroRecord and TCandidateRecord value-type contracts and recordset hydration; Config target-signature and safe object-name helpers; stateless database-role/path normalization; DAO lookup-property mutation;
'     dictionary-key sorting; popup formatting/result helpers and low-level UtilityMsgBox display; warning-metric and elapsed-time text helpers; keyed collection marker helper; TempVar suppression-state reader.
' Configuration: TempVar names must remain compatible with the existing Access Development Console modules.
' Behavior:
'     Missing suppression TempVars return False; elapsed-time formatting handles midnight rollover; keyed collection markers remain idempotent; DAO property writes replace only a mismatched property entry and never
'     delete or rebuild its owning field, table, or stored data.
' Code Status: Active.
' Todo: None.
' Reporting Bugs:
' Module Level Bugs:
' Public Subs:
'     MacroRecordFromRecordset(ByVal rs As DAO.Recordset, ByRef cfg As TMacroRecord) -> None
'     CandidateRecordFromRecordset(ByVal rs As DAO.Recordset, ByRef candidate As TCandidateRecord) -> None
'     ResetNamingPerformanceDiagnostics() -> None
'     BeginNamingSnapshot() -> None
'     RefreshManagedNameSnapshot() -> None
'     MarkCollectionTextKey(ByRef values As Collection, ByVal keyText As String) -> None
'     InvalidateNamingSnapshot() -> None
'     PumpSetupResponsiveness(Optional ByVal forceDispatch As Boolean = False) -> None
'     SetNamingBootstrapTableNames(ByVal settingsTableName As String, ByVal managedObjectsTableName As String) -> None
'     EnsureTableDefFieldOrderFromText(ByVal tdf As DAO.tableDef, ByVal fieldOrderText As String, Optional ByVal designChangeAlreadyPrepared As Boolean = False) -> None
'     EnsureSingleFieldIndex(ByVal db As DAO.Database, ByVal tableName As String, ByVal indexName As String, ByVal fieldName As String, ByVal uniqueIndex As Boolean, ByVal primaryIndex As Boolean, Optional ByVal
'     designChangeAlreadyPrepared As Boolean = False, Optional ByVal prepareOpenObjectForDesignChange As Boolean = True) -> None
'     EnsureDatabasePropertyRemoved(ByVal db As DAO.Database, ByVal propertyName As String, Optional ByVal sourceProcedure As String = "EnsureDatabasePropertyRemoved") -> None
'     ValidateGeneratedFormPreparationIdentity(ByVal physicalFormName As String) -> None
' Public Functions:
'     AccessObjectTypeFromCode(ByVal objectTypeCode As String) -> AcObjectType
'     AccessObjectTypeFromMSysType(ByVal msysType As Long) -> AcObjectType
'     CanonicalNavigationPaneWriteMSysType(ByVal objectType As AcObjectType) -> Long
'     NamingSnapshotIsReady() -> Boolean
'     ManagedReportChatGPTUIClassification(ByVal physicalReportName As String, ByRef isCanonicalReport As Boolean) -> Boolean
'     IncludeSystemMacroDesignatorInRibbonTitle() -> Boolean
'     SchemaDefinitionFieldOrderText(ByVal schemaDefinitions As Collection) -> String
'     DaoTableExists(ByVal db As DAO.Database, ByVal tableName As String) -> Boolean
'     UtilityMsgBox(ByVal prompt As Variant, Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As Variant) -> VbMsgBoxResult
'     MsgBoxStyleWithSuppressionCancel(ByVal requestedStyle As VbMsgBoxStyle, ByVal resultPrompt As Boolean) -> VbMsgBoxStyle
'     PopupMessagesSuppressedForCurrentRun() -> Boolean
'     SuppressedPopupDefaultResult(ByVal requestedStyle As VbMsgBoxStyle) -> VbMsgBoxResult
'     CompletionPopupIcon(ByVal requestedIcon As VbMsgBoxStyle, ByVal warningDetected As Boolean) -> VbMsgBoxStyle
'     ApplyLookupFieldProperties(ByVal fld As DAO.Field, ByVal rowSourceType As String, ByVal rowSourceText As String, ByVal LimitToList As Boolean, ByVal columnCount As Long, ByVal BoundColumn As Long, ByVal
'     columnWidthsText As String, Optional ByVal sourceProcedure As String = "ApplyLookupFieldProperties", Optional ByVal nothingErrorNumber As Long = vbObjectError + 6804, Optional ByVal blankNameErrorNumber As
'     Long = vbObjectError + 6805, Optional ByVal propertyMask As LookupFieldPropertyMask = LOOKUP_FIELD_PROPERTY_ALL) -> Boolean
'     ElapsedSecondsFromTimer(ByVal startTimer As Single) -> Double
'     QuotedTextCollectionList(ByVal values As Collection) -> String
'     Part8AppendListItem(ByVal listText As String, ByVal itemText As String) -> String
'     LengthPrefixedVariant(ByVal fieldValue As Variant) -> String
'     AppendLine(ByVal existingText As String, ByVal newLine As String) -> String
'     TextInsideFirstParentheses(ByVal value As String) -> String
'     SplitTopLevelCommaText(ByVal value As String) -> Collection
'     EmptyToPlaceholder(ByVal valueText As String, ByVal placeholderText As String) -> String
'     FunctionReturnTypeFromDeclaration(ByVal declarationText As String) -> String
'     ProcedureArgumentIsOptional(ByVal argText As String) -> Boolean
'     ParseProcedureArguments(ByVal argumentsText As String) -> Collection
'     ProcedureSignatureHasParamArray(ByVal argumentSignature As String) -> Boolean
'     ProcedureNameFromRunCodeExpression(ByVal runCodeExpression As String) -> String
'     FirstStringArgumentFromFunctionCall(ByVal callText As String) -> String
'     PublicCallableProcedureNameFromDeclaration(ByVal declarationText As String) -> String
'     CodeLineMayDeclareProcedure(ByVal lineText As String, ByVal ProcedureName As String) -> Boolean
'     ProcedureNameFromDeclarationRemainder(ByVal remainderText As String) -> String
'     FirstNonBlankLineIndex(ByVal codeLines As Variant, ByVal maxScanIndex As Long) -> Long
'     ManagedHeaderEndIndex(ByVal codeLines As Variant, ByVal headerStartIndex As Long, ByVal maxScanIndex As Long) -> Long
'     HeaderTitleFromLines(ByVal codeLines As Variant, ByVal headerStartIndex As Long, ByVal maxScanIndex As Long) -> String
'     HeaderValueByPrefix(ByVal codeLines As Variant, ByVal headerStartIndex As Long, ByVal headerEndIndex As Long, ByVal prefixText As String) -> String
'     IsHeaderSeparatorLine(ByVal lineText As String) -> Boolean
'     IsComponentHeaderLine(ByVal lineText As String) -> Boolean
'     FirstIdentifierToken(ByVal valueText As String) -> String
'     IsIdentifierStartCharacter(ByVal ch As String) -> Boolean
'     VbaLineContinues(ByVal lineText As String) -> Boolean
'     RemoveVbaLineContinuation(ByVal lineText As String) -> String
'     CodeLineWithoutTrailingComment(ByVal lineText As String) -> String
'     AppendedNote(ByVal existingNotes As String, ByVal newNoteText As String) -> String
'     StableTextHashHex(ByVal value As String) -> String
'     AppendUniquePipeDelimitedNames(ByVal existingNames As String, ByVal additionalNames As String) -> String
'     PipeDelimitedCollectionNames(ByVal names As Collection) -> String
'     VbaSourceIdentifierCharacter(ByVal characterText As String) -> Boolean
'     LeadingIdentifier(ByVal valueText As String) -> String
'     VbaRemCommentStartsAt(ByVal sourceLine As String, ByVal characterIndex As Long) -> Boolean
'     StripVbaComment(ByVal sourceLine As String) -> String
'     AppendDelimitedValue(ByVal existingText As String, ByVal newValue As String) -> String
'     ManagedDescriptionPrefixWithinByteBudget(ByVal authoritativeDescription As String, ByVal byteBudget As Long) -> String
'     NormalizedDuplicateTextKey(ByVal valueText As String, ByVal trimValue As Boolean) -> String
'     TraceFileSafeText(ByVal valueText As Variant) -> String
'     FileNameFromPath(ByVal filePath As String) -> String
'     AppendRecoveryPayloadLine(ByVal payloadText As String, ByVal lineText As String) -> String
'     UIDurableEncode(ByVal sourceText As String) -> String
'     UIDurableNoteValue(ByVal notesText As String, ByVal keyName As String) -> String
'     NullDisplayText(ByVal valueText As String) -> String
'     NormalizeFieldInput(ByVal value As String) -> String
'     NormalizedConfigOwnerDatabaseRole(ByVal ConfigOwnerDatabaseRole As String) -> String
'     NormalizedDatabasePath(ByVal value As String) -> String
'     DatabaseBaseNameFromPath(ByVal databasePath As String) -> String
'     FormattingInsensitiveQuerySqlKey(ByVal sqlText As String) -> String
'     NavigationPaneLogicalGroupName(ByVal groupName As String, Optional ByVal systemMacroGroup As Boolean = False) -> String
'     TextContainsDecimalDigitsOnly(ByVal textValue As String) -> Boolean
'     CommonTextIndicatesWarning(ByVal sourceText As String) -> Boolean
'     FingerprintTokenPart(ByVal valueText As String) -> String
'     AdlerTextHash(ByVal textValue As String) -> String
'     FirstTextDifferencePosition(ByVal leftText As String, ByVal rightText As String) -> Long
'     TextComparisonExcerpt(ByVal textValue As String, ByVal differencePosition As Long) -> String
'     CleanProgressScopeLabel(ByVal scopeText As String) -> String
'     LogicalObjectNameFromPhysicalName(ByVal namingCategory As String, ByVal physicalObjectName As String) -> String
'     BooleanTextValue(ByVal valueText As String, ByVal DefaultValue As Boolean) -> Boolean
'     ApplyDisplayBranding(ByVal sourceText As String) -> String
'     HumanReadablePopupText(ByVal promptText As String, ByVal TitleText As String, ByVal buttons As VbMsgBoxStyle) -> String
'     JoinLastCollectionItems(ByVal values As Collection, ByVal delimiter As String, Optional ByVal maxItems As Long = 40) -> String
'     ProcedureCallText(ByVal ProcedureName As String, ByVal argumentsText As String) -> String
'     CleanDescriptionForCaption(ByVal descriptionText As String) -> String
'     TargetCallSignature(ByVal ConfigOwnerDatabaseRole As String, ByVal MacroAction As String, ByVal TargetProcedureName As String, ByVal TargetArguments As String, ByVal TargetObjectName As String, ByVal
'     ConfirmationMode As String) -> String
'     NamingPerformanceSummary() -> String
'     NormalizedQuerySqlForComparison(ByVal db As DAO.Database, ByVal querySql As String) -> String
'     FormattedElapsedSeconds(ByVal elapsedSeconds As Double) -> String
'     NormalizedObjectTypeCode(ByVal objectTypeCode As String) -> String
'     ObjectTypeCodeFromMSysType(ByVal msysType As Long) -> String
'     SupportedMSysObjectTypesSqlList() -> String
'     MSysObjectTypePredicate(ByVal objectType As AcObjectType, ByVal fieldExpression As String) -> String
'     BuildNavigationPaneGroupName(ByVal logicalGroupName As String, Optional ByVal systemMacroGroup As Boolean = False) -> String
'     BootstrapSettingsTableName() -> String
'     DiagnosticDateText(ByVal valueData As Variant) -> String
'     MsgBoxResultText(ByVal resultValue As VbMsgBoxResult) -> String
'     CompletionSuccessText(ByVal cleanSuccessText As String, ByVal warningDetected As Boolean) -> String
'     FingerprintFromPayload(ByVal payloadText As String) -> String
'     DaoObjectPropertyFingerprintText(ByVal target As Object, ByVal propertyName As String) -> String
'     GeneratedFormPreparationRequiredPrefix(ByVal physicalFormName As String) -> String
'     ManagedObjectPhysicalName(ByVal namingCategory As String, ByVal logicalObjectName As String) -> String
'     TargetManagedObjectPhysicalName(ByVal namingCategory As String, ByVal logicalObjectName As String) -> String
'     ManagedObjectNamingCategory(ByVal objectTypeCode As String, ByVal categoryCode As String, ByVal physicalObjectName As String) -> String
'     BootstrapManagedObjectsTableName() -> String
'     StatusBarProgressText(ByVal statusText As String) -> String
'     StatusBarHierarchyText(ByVal statusText As String) -> String
'     SuppressiblePopupPrompt(ByVal promptText As String, ByVal resultPrompt As Boolean) -> String
'     CompletionPopupMessage(ByVal messageText As String, ByVal warningDetected As Boolean) -> String
'     DatabaseTextPropertyValue(ByVal db As DAO.Database, ByVal propertyName As String, ByRef propertyExists As Boolean) -> String
'     IsValidConfigOwnerDatabaseRole(ByVal ConfigOwnerDatabaseRole As String) -> Boolean
'     QuerySqlEquivalent(ByVal db As DAO.Database, ByVal existingSql As String, ByVal requestedSql As String) -> Boolean
'     NavigationPaneGroupUsesConfiguredAffixes(ByVal groupName As String, Optional ByVal systemMacroGroup As Boolean = False) -> Boolean
'     SortedDictionaryKeys(ByVal values As Object) -> Variant
'     DatabaseAccessObjectExistsByName(ByVal db As DAO.Database, ByVal objectType As AcObjectType, ByVal objectName As String) -> Boolean
'     AccessObjectExistsInCurrentDatabase(ByVal objectType As AcObjectType, ByVal objectName As String) -> Boolean
'     EnsureFieldColumnOrder(ByVal fld As DAO.Field, ByVal tableNameText As String, ByVal expectedColumnOrder As Long) -> Boolean
'     DaoFieldExists(ByVal tdf As DAO.tableDef, ByVal fieldName As String) -> Boolean
'     EnsureDaoPropertyValue(ByVal targetObject As Object, ByVal propertyName As String, ByVal requestedType As DAO.DataTypeEnum, ByVal requestedValue As Variant, Optional ByVal sourceProcedure As String =
'     "EnsureDaoPropertyValue", Optional ByVal targetDiagnosticContext As String = vbNullString, Optional ByRef repairOutcome As Variant) -> Boolean
'     DatabaseTextPropertyMatches(ByVal db As DAO.Database, ByVal propertyName As String, ByVal expectedValue As String, ByRef propertyExists As Boolean) -> Boolean
' Public Properties: None
' ================================================================================================================================================================================

Option Compare Database
Option Explicit

' ===============================================================================================================================================================================
' MODULE-LEVEL DECLARATIONS
' ===============================================================================================================================================================================

' MC_Records row-hydration cache.
' INTENTIONAL KEYED LOOKUP CACHE:
' Maps MC_Records field names to ordinals for one observed recordset field signature so hydration can use stable keyed access without repeated field scans. The signature change/reset path invalidates it; callers
' cannot mutate the module-owned dictionary.
Private mConfigFieldOrdinals As Object
Private mConfigFieldSignature As String

' Shared database-context role identities.
Public Const CONFIG_OWNER_DATABASE_ROLE_HOST_DB As String = "HostDb"
Public Const CONFIG_OWNER_DATABASE_ROLE_CODE_DB As String = "CodeDb"

' Access object-name limit shared by generated/configured naming code.
Public Const MAX_ACCESS_OBJECT_NAME_LENGTH As Long = 64
' Access field-name limit shared by schema-definition and validation code.
Public Const MAX_ACCESS_FIELD_NAME_LENGTH As Long = 64
' Maximum physical size of an Access Short Text field.
Public Const MAX_ACCESS_SHORT_TEXT_LENGTH As Long = 255
' Maximum stored length permitted for an Access field ValidationRule expression in the managed catalog.
Public Const MAX_ACCESS_FIELD_VALIDATION_RULE_LENGTH As Long = 2048

' Named lookup-property selection mask used by cross-module DAO lookup-field mutation callers.
Public Enum LookupFieldPropertyMask
    LOOKUP_FIELD_PROPERTY_BOUND_COLUMN = 1
    LOOKUP_FIELD_PROPERTY_COLUMN_WIDTHS = 2
    LOOKUP_FIELD_PROPERTY_DISPLAY_CONTROL = 4
    LOOKUP_FIELD_PROPERTY_ROW_SOURCE_TYPE = 8
    LOOKUP_FIELD_PROPERTY_ROW_SOURCE = 16
    LOOKUP_FIELD_PROPERTY_COLUMN_COUNT = 32
    LOOKUP_FIELD_PROPERTY_LIMIT_TO_LIST = 64
    LOOKUP_FIELD_PROPERTY_ALL = 127
End Enum

' Shared Help Document synchronization-scope type contract. Catalog remains the behavioral authority; the shared type lives here so Catalog and Phase C do not create a public-type dependency cycle.
Public Enum HelpDocumentSynchronizationScope
    HelpDocumentScopeSetupSystemAndAllActive = 1
    HelpDocumentScopeAllActive = 2
    HelpDocumentScopeSetupSystemOnly = 3
End Enum

' One loaded accepted/current macro configuration record shared across ADC modules.
Public Type TMacroRecord
    ' Configuration identity, activation, and status.
    ID As Long
    Active As Boolean
    ConfigurationStatus As String
    ' Macro identity and classification.
    MacroKey As String
    MacroTitle As String
    MacroAction As String
    MacroCategoryKey As String
    MacroSubcategoryKey As String
    ' Configuration routing ownership.
    ConfigOwnerDatabaseRole As String
    ' Procedure execution target and result handling.
    TargetProcedureName As String
    TargetProcedureModuleName As String
    TargetProcedureKind As String
    TargetArguments As String
    ResultHandling As String
    ExpectedReturnType As String
    ' Object, database, and Navigation Pane targets.
    TargetObjectName As String
    GenerationTargetDbPath As String
    navigationPaneGroupName As String
    ' Confirmation behavior and user-facing descriptions.
    ConfirmationMode As String
    ActionDescription As String
    LauncherTooltip As String
    ObjectDescription As String
    SourceNotes As String
    ' Resolved target argument counts retained with the accepted configuration because runtime target validation and generation consume them after candidate acceptance.
    RequiredArgumentCount As Long
    OptionalArgumentCount As Long
    ' Duplicate handling.
    SelectedForDuplicateDeletion As Boolean
    ' Validation and generated-macro results.
    LastValidatedOn As Date
    LastValidationResult As String
    GeneratedMacroObjectName As String
    LastGeneratedOn As Date
    LastGeneratedResult As String
End Type

' One loaded candidate/discovery record shared by candidate scan, acceptance, stale-review, and health-review workflows.
Public Type TCandidateRecord
    ID As Long
    CandidateStatus As String
    CandidateMacroKey As String
    MacroKey As String
    MacroTitle As String
    MacroAction As String
    MacroCategoryKey As String
    MacroSubcategoryKey As String
    ConfigOwnerDatabaseRole As String
    CandidateOwnerDatabaseRole As String
    TargetProcedureName As String
    TargetProcedureModuleName As String
    TargetProcedureKind As String
    TargetArguments As String
    ResultHandling As String
    ExpectedReturnType As String
    TargetObjectName As String
    GenerationTargetDbPath As String
    navigationPaneGroupName As String
    ConfirmationMode As String
    ActionDescription As String
    LauncherTooltip As String
    ObjectDescription As String
    SourceNotes As String
    ScanSignature As String
    LastScanBatchID As String
    LastScannedOn As Date
    HasRequiredArguments As Boolean
    RequiredArgumentCount As Long
    OptionalArgumentCount As Long
    SelectedForAcceptance As Boolean
    AcceptanceStatus As String
    AcceptedConfigurationID As Long
    AcceptanceNotes As String
End Type
' Run-scoped popup-suppression TempVars.
Public Const TEMPVAR_SUPPRESS_POPUPS As String = "ADCSuppressPopupsForCurrentRun"
Public Const TEMPVAR_SUPPRESS_POPUPS_REASON As String = "ADCSuppressPopupsReasonForCurrentRun"

' Object naming categories.
Public Const NAMING_CATEGORY_TABLE As String = "Table"
Public Const NAMING_CATEGORY_QUERY As String = "Query"
Public Const NAMING_CATEGORY_REPORT As String = "Report"
Public Const NAMING_CATEGORY_GENERATED_FORM As String = "GeneratedForm"
Public Const NAMING_CATEGORY_GENERATED_MACRO As String = "GeneratedMacro"
Public Const NAMING_CATEGORY_SYSTEM_MACRO As String = "SystemMacro"
Public Const NAMING_CATEGORY_CUSTOM_FORM As String = "CustomForm"

' Naming-policy setting keys.
Public Const SETTING_TABLE_OBJECT_PREFIX As String = "TableObjectPrefix"
Public Const SETTING_TABLE_OBJECT_POSTFIX As String = "TableObjectPostfix"
Public Const SETTING_QUERY_OBJECT_PREFIX As String = "QueryObjectPrefix"
Public Const SETTING_QUERY_OBJECT_POSTFIX As String = "QueryObjectPostfix"
Public Const SETTING_GENERATED_MACRO_OBJECT_PREFIX As String = "GeneratedMacroObjectPrefix"
Public Const SETTING_GENERATED_MACRO_OBJECT_POSTFIX As String = "GeneratedMacroObjectPostfix"
Public Const SETTING_SYSTEM_MACRO_OBJECT_PREFIX As String = "SystemMacroObjectPrefix"
Public Const SETTING_SYSTEM_MACRO_OBJECT_POSTFIX As String = "SystemMacroObjectPostfix"
Public Const SETTING_CUSTOM_FORM_OBJECT_PREFIX As String = "CustomFormObjectPrefix"
Public Const SETTING_CUSTOM_FORM_OBJECT_POSTFIX As String = "CustomFormObjectPostfix"
Public Const SETTING_INCLUDE_SYSTEM_MACRO_DESIGNATOR_IN_RIBBON_TITLE As String = "IncludeSystemMacroDesignatorInRibbonTitle"

' Bootstrap property names and DAO property limits.
Private Const BOOTSTRAP_SETTINGS_TABLE_PROPERTY As String = "ADC_CurrentSettingsTableName"
Private Const BOOTSTRAP_MANAGED_OBJECTS_TABLE_PROPERTY As String = "ADC_CurrentManagedObjectsTableName"
Public Const DAO_PROPERTY_TEXT_LIMIT_BYTES As Long = 2048
Public Const DAO_DESCRIPTION_PROPERTY_NAME As String = "Description"

' Access system-table, object-type, and DAO container identifiers.
Public Const NAV_PANE_GROUPS_TABLE As String = "MSysNavPaneGroups"
Public Const NAV_PANE_GROUP_OBJECTS_TABLE As String = "MSysNavPaneGroupToObjects"
Public Const NAV_PANE_OBJECT_IDS_TABLE As String = "MSysNavPaneObjectIDs"
Public Const MSYS_OBJECTS_TABLE As String = "MSysObjects"
Public Const MSYS_OBJECT_TYPE_LOCAL_TABLE As Long = 1
Public Const MSYS_OBJECT_TYPE_ODBC_TABLE As Long = 4
Public Const MSYS_OBJECT_TYPE_QUERY As Long = 5
Public Const MSYS_OBJECT_TYPE_LINKED_TABLE As Long = 6
Public Const MSYS_OBJECT_TYPE_FORM As Long = -32768
Private Const NAV_PANE_OBJECT_TYPE_FORM As Long = 32768
Public Const MSYS_OBJECT_TYPE_MACRO As Long = -32766
Private Const NAV_PANE_OBJECT_TYPE_MACRO As Long = 32770
Public Const MSYS_OBJECT_TYPE_REPORT As Long = -32764
Private Const NAV_PANE_OBJECT_TYPE_REPORT As Long = 32772
Public Const FORM_DAO_CONTAINER_NAME As String = "Forms"
Public Const REPORT_DAO_CONTAINER_NAME As String = "Reports"
Public Const MACRO_DAO_CONTAINER_NAME As String = "Scripts"

' Cached naming-policy values and instrumentation counters.
Private mNamingSnapshotActive As Boolean
Private mNamingPolicyReady As Boolean
' INTENTIONAL OPERATION SNAPSHOT / KEYED LOOKUP:
' Captures managed logical-name -> physical-name mappings for the active naming snapshot so all naming consumers observe one coherent MAC state. Begin/end/invalidation of the naming snapshot owns its lifetime; this
' dictionary is derived from MAC and is not a second naming authority.
Private mNamingSnapshotManagedNames As Object
Private mNamingTableObjectPrefix As String
Private mNamingTableObjectPostfix As String
Private mNamingQueryObjectPrefix As String
Private mNamingQueryObjectPostfix As String
Private mNamingGeneratedMacroObjectPrefix As String
Private mNamingGeneratedMacroObjectPostfix As String
Private mNamingSystemMacroObjectPrefix As String
Private mNamingSystemMacroObjectPostfix As String
Private mNamingCustomFormObjectPrefix As String
Private mNamingCustomFormObjectPostfix As String
Private mNamingNavigationPaneGroupPrefix As String
Private mNamingNavigationPaneGroupPostfix As String
Private mNamingSystemMacroNavigationPaneGroupPrefix As String
Private mNamingSystemMacroNavigationPaneGroupPostfix As String
Private mNamingIncludeSystemMacroDesignatorInRibbonTitle As Boolean
Private mNamingIncludeSystemMacroDesignatorInNavigationPaneTitle As Boolean
Private mNamingSnapshotBuildCount As Long
Private mNamingManagedNameRefreshCount As Long
Private mNamingSnapshotInvalidationCount As Long
Private mNamingDatabaseLookupCount As Long

' ===============================================================================================================================================================================
' PROCEDURES
' ===============================================================================================================================================================================




Private Sub EnsureConfigFieldOrdinalCache(ByVal rs As DAO.Recordset)
    'Builds or refreshes the cached macro-record recordset field-ordinal lookup for the supplied recordset shape.
    On Error GoTo ErrHandler
    Dim fieldIndex As Long
    Dim fieldName As String
    Dim schemaSignature As String
    Dim fieldOrdinals As Object
    If rs Is Nothing Then Err.Raise vbObjectError + 5960, "EnsureConfigFieldOrdinalCache", "The configuration recordset is Nothing."
    schemaSignature = CStr(rs.fields.count) & ":"
    For fieldIndex = 0 To rs.fields.count - 1
        PumpSetupResponsiveness
        fieldName = CStr(rs.fields(fieldIndex).Name)
        schemaSignature = schemaSignature & CStr(Len(fieldName)) & ":" & fieldName & ":" & CStr(rs.fields(fieldIndex).Type) & ";"
    Next fieldIndex
    If Not mConfigFieldOrdinals Is Nothing Then
        If StrComp(schemaSignature, mConfigFieldSignature, vbBinaryCompare) = 0 Then Exit Sub
    End If
    ' INTENTIONAL KEYED LOOKUP: fieldOrdinals provides keyed membership/lookup over source rows or state already read by EnsureConfigFieldOrdinalCache; it is a derived access path, not a second authority.
    Set fieldOrdinals = CreateObject("Scripting.Dictionary")
    fieldOrdinals.CompareMode = vbTextCompare
    For fieldIndex = 0 To rs.fields.count - 1
        PumpSetupResponsiveness
        fieldName = CStr(rs.fields(fieldIndex).Name)
        If fieldOrdinals.Exists(fieldName) = False Then fieldOrdinals.Add fieldName, fieldIndex
    Next fieldIndex
    Set mConfigFieldOrdinals = fieldOrdinals
    mConfigFieldSignature = schemaSignature
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Set mConfigFieldOrdinals = Nothing
    mConfigFieldSignature = vbNullString
    Err.Raise savedErrNumber, "EnsureConfigFieldOrdinalCache", "Could not prepare the cached configuration-record field map." & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Function ConfigFieldValue(ByVal rs As DAO.Recordset, ByVal fieldName As String) As Variant
    'Returns a macro-record recordset field value by name through the cached field-ordinal lookup.
    On Error GoTo ErrHandler
    If mConfigFieldOrdinals Is Nothing Then EnsureConfigFieldOrdinalCache rs
    If mConfigFieldOrdinals.Exists(fieldName) Then
        ConfigFieldValue = rs.fields(CLng(mConfigFieldOrdinals(fieldName))).value
    Else
        ConfigFieldValue = Null
    End If
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ConfigFieldValue", "Could not read a cached configuration-record field." & vbCrLf & "Field: " & fieldName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function ConfigDateOrZero(ByVal value As Variant) As Date
    'Converts a macro-record value to Date, returning zero for Null or Empty input.
    If IsNull(value) Or IsEmpty(value) Then
        ConfigDateOrZero = 0
    Else
        ConfigDateOrZero = CDate(value)
    End If
End Function

Private Function ConfigBool(ByVal value As Variant, ByVal DefaultValue As Boolean) As Boolean
    'Converts a macro-record value to Boolean, applying the supplied default for Null or Empty input.
    If IsNull(value) Or IsEmpty(value) Then
        ConfigBool = DefaultValue
    Else
        ConfigBool = CBool(value)
    End If
End Function



' Purpose: Populates a TMacroRecord value from the current macro-record recordset row.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   rs (Input; DAO.Recordset): DAO.Recordset supplying or receiving the row data used by the operation.
'   cfg (Input/output; TMacroRecord): ADC macro record containing persisted accepted/current configuration values; the routine reads it and may populate or update its fields when passed ByRef.
' Returns: None (Sub). Results are returned through the documented ByRef parameters and/or the routine's database, UI, or module-state side effects.
Public Sub MacroRecordFromRecordset(ByVal rs As DAO.Recordset, ByRef cfg As TMacroRecord)
    'Populates the configuration-only TMacroRecord value from the current MC_Records row.
    EnsureConfigFieldOrdinalCache rs
    With cfg
        .ID = SharedValue_NzLong(ConfigFieldValue(rs, FIELD_ID))
        .Active = ConfigBool(ConfigFieldValue(rs, FIELD_ACTIVE), True)
        .ConfigurationStatus = SharedText_Nz(ConfigFieldValue(rs, FIELD_CONFIGURATION_STATUS))
        .MacroKey = SharedText_Nz(ConfigFieldValue(rs, FIELD_MACRO_KEY))
        .MacroTitle = SharedText_Nz(ConfigFieldValue(rs, FIELD_MACRO_TITLE))
        .MacroAction = SharedText_Nz(ConfigFieldValue(rs, FIELD_MACRO_ACTION))
        .MacroCategoryKey = SharedText_Nz(ConfigFieldValue(rs, FIELD_MACRO_CATEGORY_KEY))
        .MacroSubcategoryKey = SharedText_Nz(ConfigFieldValue(rs, FIELD_MACRO_SUBCATEGORY_KEY))
        .ConfigOwnerDatabaseRole = SharedText_Nz(ConfigFieldValue(rs, FIELD_CONFIG_OWNER_DATABASE_ROLE))
        .TargetProcedureName = SharedText_Nz(ConfigFieldValue(rs, FIELD_TARGET_PROCEDURE_NAME))
        .TargetProcedureModuleName = SharedText_Nz(ConfigFieldValue(rs, FIELD_TARGET_PROCEDURE_MODULE_NAME))
        .TargetProcedureKind = SharedText_Nz(ConfigFieldValue(rs, FIELD_TARGET_PROCEDURE_KIND))
        .TargetArguments = SharedText_Nz(ConfigFieldValue(rs, FIELD_TARGET_ARGUMENTS))
        .ResultHandling = SharedText_Nz(ConfigFieldValue(rs, FIELD_RESULT_HANDLING))
        .ExpectedReturnType = SharedText_Nz(ConfigFieldValue(rs, FIELD_EXPECTED_RETURN_TYPE))
        .TargetObjectName = SharedText_Nz(ConfigFieldValue(rs, FIELD_TARGET_OBJECT_NAME))
        .GenerationTargetDbPath = SharedText_Nz(ConfigFieldValue(rs, FIELD_GENERATION_TARGET_DB_PATH))
        .navigationPaneGroupName = SharedText_Nz(ConfigFieldValue(rs, FIELD_NAVIGATION_PANE_GROUP))
        .ConfirmationMode = SharedText_Nz(ConfigFieldValue(rs, FIELD_CONFIRMATION_MODE))
        .ActionDescription = SharedText_Nz(ConfigFieldValue(rs, FIELD_ACTION_DESCRIPTION))
        .LauncherTooltip = SharedText_Nz(ConfigFieldValue(rs, FIELD_LAUNCHER_TOOLTIP))
        .ObjectDescription = SharedText_Nz(ConfigFieldValue(rs, FIELD_OBJECT_DESCRIPTION))
        .SourceNotes = SharedText_Nz(ConfigFieldValue(rs, FIELD_SOURCE_NOTES))
        .RequiredArgumentCount = SharedValue_NzLong(ConfigFieldValue(rs, FIELD_REQUIRED_ARGUMENT_COUNT))
        .OptionalArgumentCount = SharedValue_NzLong(ConfigFieldValue(rs, FIELD_OPTIONAL_ARGUMENT_COUNT))
        .SelectedForDuplicateDeletion = ConfigBool(ConfigFieldValue(rs, FIELD_SELECTED_FOR_DUPLICATE_DELETION), False)
        .LastValidatedOn = ConfigDateOrZero(ConfigFieldValue(rs, FIELD_LAST_VALIDATED_ON))
        .LastValidationResult = SharedText_Nz(ConfigFieldValue(rs, FIELD_LAST_VALIDATION_RESULT))
        .GeneratedMacroObjectName = SharedText_Nz(ConfigFieldValue(rs, FIELD_GENERATED_MACRO_OBJECT_NAME))
        .LastGeneratedOn = ConfigDateOrZero(ConfigFieldValue(rs, FIELD_LAST_GENERATED_ON))
        .LastGeneratedResult = SharedText_Nz(ConfigFieldValue(rs, FIELD_LAST_GENERATED_RESULT))
    End With
    If Len(Trim$(cfg.MacroAction)) = 0 Then cfg.MacroAction = ACTION_RUNCODE
    If Len(Trim$(cfg.TargetProcedureKind)) = 0 Then cfg.TargetProcedureKind = TARGET_PROCEDURE_KIND_AUTO
    If Len(Trim$(cfg.ResultHandling)) = 0 Then cfg.ResultHandling = RESULT_HANDLING_AUTO
    If Len(Trim$(cfg.ConfirmationMode)) = 0 Then cfg.ConfirmationMode = CONFIRM_AUTO
    If Len(Trim$(cfg.ConfigurationStatus)) = 0 Then cfg.ConfigurationStatus = CONFIGURATION_STATUS_ACTIVE
    cfg.ConfigOwnerDatabaseRole = NormalizedConfigOwnerDatabaseRole(cfg.ConfigOwnerDatabaseRole)
    If Len(Trim$(cfg.ConfigOwnerDatabaseRole)) = 0 Then cfg.ConfigOwnerDatabaseRole = CONFIG_OWNER_DATABASE_ROLE_HOST_DB
End Sub

' Purpose: Populates a TCandidateRecord from the current MC_Candidates row.
Public Sub CandidateRecordFromRecordset(ByVal rs As DAO.Recordset, ByRef candidate As TCandidateRecord)
    EnsureConfigFieldOrdinalCache rs
    With candidate
        .ID = SharedValue_NzLong(ConfigFieldValue(rs, FIELD_ID))
        .CandidateStatus = SharedText_Nz(ConfigFieldValue(rs, FIELD_CANDIDATE_STATUS))
        .CandidateMacroKey = SharedText_Nz(ConfigFieldValue(rs, FIELD_CANDIDATE_MACRO_KEY))
        .MacroKey = SharedText_Nz(ConfigFieldValue(rs, FIELD_MACRO_KEY))
        .MacroTitle = SharedText_Nz(ConfigFieldValue(rs, FIELD_MACRO_TITLE))
        .MacroAction = SharedText_Nz(ConfigFieldValue(rs, FIELD_MACRO_ACTION))
        .MacroCategoryKey = SharedText_Nz(ConfigFieldValue(rs, FIELD_MACRO_CATEGORY_KEY))
        .MacroSubcategoryKey = SharedText_Nz(ConfigFieldValue(rs, FIELD_MACRO_SUBCATEGORY_KEY))
        .ConfigOwnerDatabaseRole = SharedText_Nz(ConfigFieldValue(rs, FIELD_CONFIG_OWNER_DATABASE_ROLE))
        .CandidateOwnerDatabaseRole = SharedText_Nz(ConfigFieldValue(rs, FIELD_CANDIDATE_OWNER_DATABASE_ROLE))
        .TargetProcedureName = SharedText_Nz(ConfigFieldValue(rs, FIELD_TARGET_PROCEDURE_NAME))
        .TargetProcedureModuleName = SharedText_Nz(ConfigFieldValue(rs, FIELD_TARGET_PROCEDURE_MODULE_NAME))
        .TargetProcedureKind = SharedText_Nz(ConfigFieldValue(rs, FIELD_TARGET_PROCEDURE_KIND))
        .TargetArguments = SharedText_Nz(ConfigFieldValue(rs, FIELD_TARGET_ARGUMENTS))
        .ResultHandling = SharedText_Nz(ConfigFieldValue(rs, FIELD_RESULT_HANDLING))
        .ExpectedReturnType = SharedText_Nz(ConfigFieldValue(rs, FIELD_EXPECTED_RETURN_TYPE))
        .TargetObjectName = SharedText_Nz(ConfigFieldValue(rs, FIELD_TARGET_OBJECT_NAME))
        .GenerationTargetDbPath = SharedText_Nz(ConfigFieldValue(rs, FIELD_GENERATION_TARGET_DB_PATH))
        .navigationPaneGroupName = SharedText_Nz(ConfigFieldValue(rs, FIELD_NAVIGATION_PANE_GROUP))
        .ConfirmationMode = SharedText_Nz(ConfigFieldValue(rs, FIELD_CONFIRMATION_MODE))
        .ActionDescription = SharedText_Nz(ConfigFieldValue(rs, FIELD_ACTION_DESCRIPTION))
        .LauncherTooltip = SharedText_Nz(ConfigFieldValue(rs, FIELD_LAUNCHER_TOOLTIP))
        .ObjectDescription = SharedText_Nz(ConfigFieldValue(rs, FIELD_OBJECT_DESCRIPTION))
        .SourceNotes = SharedText_Nz(ConfigFieldValue(rs, FIELD_SOURCE_NOTES))
        .ScanSignature = SharedText_Nz(ConfigFieldValue(rs, FIELD_SCAN_SIGNATURE))
        .LastScanBatchID = SharedText_Nz(ConfigFieldValue(rs, FIELD_LAST_SCAN_BATCH_ID))
        .LastScannedOn = ConfigDateOrZero(ConfigFieldValue(rs, FIELD_LAST_SCANNED_ON))
        .HasRequiredArguments = ConfigBool(ConfigFieldValue(rs, FIELD_HAS_REQUIRED_ARGUMENTS), False)
        .RequiredArgumentCount = SharedValue_NzLong(ConfigFieldValue(rs, FIELD_REQUIRED_ARGUMENT_COUNT))
        .OptionalArgumentCount = SharedValue_NzLong(ConfigFieldValue(rs, FIELD_OPTIONAL_ARGUMENT_COUNT))
        .SelectedForAcceptance = ConfigBool(ConfigFieldValue(rs, FIELD_SELECTED_FOR_ACCEPTANCE), False)
        .AcceptanceStatus = SharedText_Nz(ConfigFieldValue(rs, FIELD_ACCEPTANCE_STATUS))
        .AcceptedConfigurationID = SharedValue_NzLong(ConfigFieldValue(rs, FIELD_ACCEPTED_CONFIGURATION_ID))
        .AcceptanceNotes = SharedText_Nz(ConfigFieldValue(rs, FIELD_ACCEPTANCE_NOTES))
    End With
    If Len(Trim$(candidate.MacroKey)) = 0 Then candidate.MacroKey = candidate.CandidateMacroKey
    If Len(Trim$(candidate.MacroAction)) = 0 Then candidate.MacroAction = ACTION_RUNCODE
    If Len(Trim$(candidate.TargetProcedureKind)) = 0 Then candidate.TargetProcedureKind = TARGET_PROCEDURE_KIND_AUTO
    If Len(Trim$(candidate.ResultHandling)) = 0 Then candidate.ResultHandling = RESULT_HANDLING_AUTO
    If Len(Trim$(candidate.ConfirmationMode)) = 0 Then candidate.ConfirmationMode = CONFIRM_AUTO
    If Len(Trim$(candidate.ConfigOwnerDatabaseRole)) = 0 Then candidate.ConfigOwnerDatabaseRole = candidate.CandidateOwnerDatabaseRole
    candidate.ConfigOwnerDatabaseRole = NormalizedConfigOwnerDatabaseRole(candidate.ConfigOwnerDatabaseRole)
End Sub
' Purpose: Clears and reinitializes naming performance diagnostics, coordinating `End` as the implementation requires.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub ResetNamingPerformanceDiagnostics()
    'Resets the counters used to diagnose naming-snapshot and naming-policy performance.
    mNamingSnapshotBuildCount = 0
    mNamingManagedNameRefreshCount = 0
    mNamingSnapshotInvalidationCount = 0
    mNamingDatabaseLookupCount = 0
End Sub


Private Function ObjectTypeDefinitionForMSysType(ByVal msysType As Long, ByRef Definition As Variant) As Boolean
    'Finds the canonical object-type definition associated with an MSysObjects type value.
    Dim rowData As Variant
    Dim acceptedTypes As Variant
    Dim acceptedType As Variant
    Dim matched As Boolean
    For Each rowData In ObjectTypeDefinitionRows()
        matched = (CLng(rowData(3)) = msysType)
        If matched = False Then
            acceptedTypes = rowData(2)
            For Each acceptedType In acceptedTypes
                If CLng(acceptedType) = msysType Then matched = True: Exit For
            Next acceptedType
        End If
        If matched Then
            Definition = rowData
            ObjectTypeDefinitionForMSysType = True
            Exit Function
        End If
    Next rowData
End Function





Private Function ObjectTypeDefinitionRows() As Variant
    'Single authority for ADC Access object-type identity, Access enum value, accepted MSys types, and canonical local Navigation Pane write type.
    Static definitions As Variant
    If IsEmpty(definitions) Then definitions = Array( _
        Array(ADCAT_OBJECT_TYPE_TABLE, CLng(acTable), Array(MSYS_OBJECT_TYPE_LOCAL_TABLE, MSYS_OBJECT_TYPE_ODBC_TABLE, MSYS_OBJECT_TYPE_LINKED_TABLE), MSYS_OBJECT_TYPE_LOCAL_TABLE), _
        Array(ADCAT_OBJECT_TYPE_QUERY, CLng(acQuery), Array(MSYS_OBJECT_TYPE_QUERY), MSYS_OBJECT_TYPE_QUERY), _
        Array(ADCAT_OBJECT_TYPE_FORM, CLng(acForm), Array(MSYS_OBJECT_TYPE_FORM), NAV_PANE_OBJECT_TYPE_FORM), _
        Array(ADCAT_OBJECT_TYPE_REPORT, CLng(acReport), Array(MSYS_OBJECT_TYPE_REPORT), NAV_PANE_OBJECT_TYPE_REPORT), _
        Array(ADCAT_OBJECT_TYPE_MACRO, CLng(acMacro), Array(MSYS_OBJECT_TYPE_MACRO), NAV_PANE_OBJECT_TYPE_MACRO))
    ObjectTypeDefinitionRows = definitions
End Function


Private Function ObjectTypeDefinitionForAccessType(ByVal objectType As AcObjectType, ByRef Definition As Variant) As Boolean
    'Finds the canonical object-type definition associated with an Access object type.
    Dim rowData As Variant
    For Each rowData In ObjectTypeDefinitionRows()
        If CLng(rowData(1)) = CLng(objectType) Then
            Definition = rowData
            ObjectTypeDefinitionForAccessType = True
            Exit Function
        End If
    Next rowData
End Function



' Purpose: Resolves access object type from code from the supplied/current ADC state, using `ObjectTypeDefinitionForCode` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   objectTypeCode (Input; String): Canonical code used to classify or select the object type code behavior.
' Returns: AcObjectType: Returns the resolved object reference. `Nothing` remains the no-object result where the existing routine permits an unresolved or unavailable object.
Public Function AccessObjectTypeFromCode(ByVal objectTypeCode As String) As AcObjectType
    'Converts a supported ADC object-type code to its Access AcObjectType value.
    Dim Definition As Variant
    If ObjectTypeDefinitionForCode(objectTypeCode, Definition) = False Then Err.Raise vbObjectError + 6897, "AccessObjectTypeFromCode", "Unsupported " & DEFAULT_UTILITY_SHORT_NAME & " Access object type: " & objectTypeCode
    AccessObjectTypeFromCode = CLng(Definition(1))
End Function

' Purpose: Converts a supported MSys object type to the corresponding Access AcObjectType value.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   msysType (Input; Long): MSysObjects Type code supplied/inspected when mapping between system-table rows and Access object types.
' Returns: AcObjectType: Returns the Access object-type constant corresponding to the supplied MSysObjects Type value; returns the routine’s defined fallback/unknown type when no supported mapping exists.
Public Function AccessObjectTypeFromMSysType(ByVal msysType As Long) As AcObjectType
    'Converts a supported MSys object type to the corresponding Access AcObjectType value.
    Dim Definition As Variant
    If ObjectTypeDefinitionForMSysType(msysType, Definition) = False Then Err.Raise vbObjectError + 6898, "AccessObjectTypeFromMSysType", "Unsupported MSys Access object type: " & CStr(msysType)
    AccessObjectTypeFromMSysType = CLng(Definition(1))
End Function




' Purpose: Returns the explicit Navigation Pane write type; acTable intentionally writes the known-local table type rather than pretending the MSys mapping is reversible.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   objectType (Input; AcObjectType): Object-type value identifying the Access/VBA object class targeted by the operation.
' Returns: Long: Returns the MSysObjects Type code used by the Navigation Pane write/restore path for the supplied managed object category.
Public Function CanonicalNavigationPaneWriteMSysType(ByVal objectType As AcObjectType) As Long
    'Returns the explicit Navigation Pane write type; acTable intentionally writes the known-local table type rather than pretending the MSys mapping is reversible.
    Dim Definition As Variant
    If ObjectTypeDefinitionForAccessType(objectType, Definition) = False Then Exit Function
    CanonicalNavigationPaneWriteMSysType = CLng(Definition(3))
End Function




Private Sub NavigationPaneGroupAffixes(ByVal systemMacroGroup As Boolean, ByRef prefixText As String, ByRef postfixText As String)
    'Returns the configured Navigation Pane prefix and postfix for regular or System Macro groups.
    If mNamingSnapshotActive And mNamingPolicyReady Then
        If systemMacroGroup Then
            prefixText = mNamingSystemMacroNavigationPaneGroupPrefix
            postfixText = mNamingSystemMacroNavigationPaneGroupPostfix
        Else
            prefixText = mNamingNavigationPaneGroupPrefix
            postfixText = mNamingNavigationPaneGroupPostfix
        End If
    ElseIf systemMacroGroup Then
        prefixText = NamingSettingText(SETTING_SYSTEM_MACRO_NAVIGATION_PANE_GROUP_PREFIX, DEFAULT_SYSTEM_MACRO_NAVIGATION_PANE_GROUP_PREFIX)
        postfixText = NamingSettingText(SETTING_SYSTEM_MACRO_NAVIGATION_PANE_GROUP_POSTFIX, DEFAULT_SYSTEM_MACRO_NAVIGATION_PANE_GROUP_POSTFIX)
    Else
        prefixText = NamingSettingText(SETTING_NAVIGATION_PANE_GROUP_PREFIX, DEFAULT_NAVIGATION_PANE_GROUP_PREFIX)
        postfixText = NamingSettingText(SETTING_NAVIGATION_PANE_GROUP_POSTFIX, DEFAULT_NAVIGATION_PANE_GROUP_POSTFIX)
    End If
End Sub

' Purpose: Starts a naming snapshot and loads the naming policy and managed-name cache for consistent repeated lookups.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub BeginNamingSnapshot()
    'Starts a naming snapshot and loads the naming policy and managed-name cache for consistent repeated lookups.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    ClearNamingSnapshotState
    mNamingSnapshotBuildCount = mNamingSnapshotBuildCount + 1
    Set db = currentDb
    LoadNamingPolicy db
    LoadManagedNameSnapshot db
    mNamingPolicyReady = True
    mNamingSnapshotActive = True
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    InvalidateNamingSnapshot
    On Error GoTo 0
    Err.Raise savedErrNumber, IIf(Len(savedErrSource) > 0, savedErrSource, "BeginNamingSnapshot"), savedErrDescription
End Sub

' Purpose: Returns whether COMMON currently holds a complete active naming-policy and managed-name snapshot that callers may reuse without rebuilding it.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: Boolean: True when naming snapshot is is ready for the next dependent operation; False when a prerequisite is not yet satisfied.
Public Function NamingSnapshotIsReady() As Boolean
    'Returns whether COMMON currently holds a complete active naming-policy and managed-name snapshot that callers may reuse without rebuilding it.
    NamingSnapshotIsReady = (mNamingSnapshotActive And mNamingPolicyReady)
End Function

Private Sub LoadNamingPolicy(ByVal db As DAO.Database)
    'Loads naming-policy settings from the supplied database into the current naming snapshot.
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim tableName As String
    Dim keyText As String
    InitializeNamingPolicyDefaults
    tableName = BootstrapSettingsTableName()
    If DaoTableExists(db, tableName) = False Then Exit Sub
    Set rs = db.OpenRecordset("SELECT [" & FIELD_SETTING_KEY & "],[" & FIELD_SETTING_VALUE & "] FROM [" & Replace(tableName, "]", "]]") & "] WHERE Len(Nz([" & FIELD_SETTING_KEY & "],''))>0;", dbOpenSnapshot)
    Do While rs.EOF = False
        keyText = CStr(rs.fields(FIELD_SETTING_KEY).value)
        ApplyNamingPolicySetting keyText, rs.fields(FIELD_SETTING_VALUE).value
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
    Err.Raise savedErrNumber, IIf(Len(savedErrSource) > 0, savedErrSource, "LoadNamingPolicy"), savedErrDescription
End Sub

Private Sub InitializeNamingPolicyDefaults()
    'Initializes in-memory naming-policy values from the canonical defaults.
    mNamingTableObjectPrefix = DEFAULT_TABLE_OBJECT_PREFIX
    mNamingTableObjectPostfix = DEFAULT_TABLE_OBJECT_POSTFIX
    mNamingQueryObjectPrefix = DEFAULT_QUERY_OBJECT_PREFIX
    mNamingQueryObjectPostfix = DEFAULT_QUERY_OBJECT_POSTFIX
    mNamingGeneratedMacroObjectPrefix = DEFAULT_GENERATED_MACRO_OBJECT_PREFIX
    mNamingGeneratedMacroObjectPostfix = DEFAULT_GENERATED_MACRO_OBJECT_POSTFIX
    mNamingSystemMacroObjectPrefix = DEFAULT_SYSTEM_MACRO_OBJECT_PREFIX
    mNamingSystemMacroObjectPostfix = DEFAULT_SYSTEM_MACRO_OBJECT_POSTFIX
    mNamingCustomFormObjectPrefix = DEFAULT_CUSTOM_FORM_OBJECT_PREFIX
    mNamingCustomFormObjectPostfix = DEFAULT_CUSTOM_FORM_OBJECT_POSTFIX
    mNamingNavigationPaneGroupPrefix = DEFAULT_NAVIGATION_PANE_GROUP_PREFIX
    mNamingNavigationPaneGroupPostfix = DEFAULT_NAVIGATION_PANE_GROUP_POSTFIX
    mNamingSystemMacroNavigationPaneGroupPrefix = DEFAULT_SYSTEM_MACRO_NAVIGATION_PANE_GROUP_PREFIX
    mNamingSystemMacroNavigationPaneGroupPostfix = DEFAULT_SYSTEM_MACRO_NAVIGATION_PANE_GROUP_POSTFIX
    mNamingIncludeSystemMacroDesignatorInRibbonTitle = DEFAULT_INCLUDE_SYSTEM_MACRO_DESIGNATOR_IN_RIBBON_TITLE
    mNamingIncludeSystemMacroDesignatorInNavigationPaneTitle = DEFAULT_INCLUDE_SYSTEM_MACRO_DESIGNATOR_IN_NAVIGATION_PANE_TITLE
End Sub

Private Sub ClearNamingSnapshotState()
    'Clears the active naming snapshot, cached names, and loaded naming-policy state.
    mNamingSnapshotActive = False
    mNamingPolicyReady = False
    Set mNamingSnapshotManagedNames = Nothing
    mNamingTableObjectPrefix = vbNullString
    mNamingTableObjectPostfix = vbNullString
    mNamingQueryObjectPrefix = vbNullString
    mNamingQueryObjectPostfix = vbNullString
    mNamingGeneratedMacroObjectPrefix = vbNullString
    mNamingGeneratedMacroObjectPostfix = vbNullString
    mNamingSystemMacroObjectPrefix = vbNullString
    mNamingSystemMacroObjectPostfix = vbNullString
    mNamingCustomFormObjectPrefix = vbNullString
    mNamingCustomFormObjectPostfix = vbNullString
    mNamingNavigationPaneGroupPrefix = vbNullString
    mNamingNavigationPaneGroupPostfix = vbNullString
    mNamingSystemMacroNavigationPaneGroupPrefix = vbNullString
    mNamingSystemMacroNavigationPaneGroupPostfix = vbNullString
End Sub


Private Sub NamingAffixes(ByVal namingCategory As String, ByRef prefixText As String, ByRef postfixText As String)
    'Returns the configured prefix and postfix for the requested managed naming category.
    If mNamingSnapshotActive And mNamingPolicyReady Then
        Select Case UCase$(Trim$(namingCategory))
            Case UCase$(NAMING_CATEGORY_TABLE): prefixText = mNamingTableObjectPrefix: postfixText = mNamingTableObjectPostfix
            Case UCase$(NAMING_CATEGORY_QUERY): prefixText = mNamingQueryObjectPrefix: postfixText = mNamingQueryObjectPostfix
            Case UCase$(NAMING_CATEGORY_GENERATED_MACRO): prefixText = mNamingGeneratedMacroObjectPrefix: postfixText = mNamingGeneratedMacroObjectPostfix
            Case UCase$(NAMING_CATEGORY_SYSTEM_MACRO): prefixText = mNamingSystemMacroObjectPrefix: postfixText = mNamingSystemMacroObjectPostfix
            Case UCase$(NAMING_CATEGORY_CUSTOM_FORM): prefixText = mNamingCustomFormObjectPrefix: postfixText = mNamingCustomFormObjectPostfix
            Case Else: Err.Raise vbObjectError + 6801, "NamingAffixes", "Unsupported managed-object naming category: " & namingCategory
        End Select
        Exit Sub
    End If
    Select Case UCase$(Trim$(namingCategory))
        Case UCase$(NAMING_CATEGORY_TABLE): prefixText = NamingSettingText(SETTING_TABLE_OBJECT_PREFIX, DEFAULT_TABLE_OBJECT_PREFIX): postfixText = NamingSettingText(SETTING_TABLE_OBJECT_POSTFIX, DEFAULT_TABLE_OBJECT_POSTFIX)
        Case UCase$(NAMING_CATEGORY_QUERY): prefixText = NamingSettingText(SETTING_QUERY_OBJECT_PREFIX, DEFAULT_QUERY_OBJECT_PREFIX): postfixText = NamingSettingText(SETTING_QUERY_OBJECT_POSTFIX, DEFAULT_QUERY_OBJECT_POSTFIX)
        Case UCase$(NAMING_CATEGORY_GENERATED_MACRO): prefixText = NamingSettingText(SETTING_GENERATED_MACRO_OBJECT_PREFIX, DEFAULT_GENERATED_MACRO_OBJECT_PREFIX): postfixText = NamingSettingText(SETTING_GENERATED_MACRO_OBJECT_POSTFIX, DEFAULT_GENERATED_MACRO_OBJECT_POSTFIX)
        Case UCase$(NAMING_CATEGORY_SYSTEM_MACRO): prefixText = NamingSettingText(SETTING_SYSTEM_MACRO_OBJECT_PREFIX, DEFAULT_SYSTEM_MACRO_OBJECT_PREFIX): postfixText = NamingSettingText(SETTING_SYSTEM_MACRO_OBJECT_POSTFIX, DEFAULT_SYSTEM_MACRO_OBJECT_POSTFIX)
        Case UCase$(NAMING_CATEGORY_CUSTOM_FORM): prefixText = NamingSettingText(SETTING_CUSTOM_FORM_OBJECT_PREFIX, DEFAULT_CUSTOM_FORM_OBJECT_PREFIX): postfixText = NamingSettingText(SETTING_CUSTOM_FORM_OBJECT_POSTFIX, DEFAULT_CUSTOM_FORM_OBJECT_POSTFIX)
        Case Else: Err.Raise vbObjectError + 6801, "NamingAffixes", "Unsupported managed-object naming category: " & namingCategory
    End Select
End Sub

' Purpose: Rebuilds the active managed-name snapshot from the current database state.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub RefreshManagedNameSnapshot()
    'Rebuilds the active managed-name snapshot from the current database state.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    If mNamingSnapshotActive = False Or mNamingPolicyReady = False Then
        BeginNamingSnapshot
        Exit Sub
    End If
    Set db = currentDb
    LoadManagedNameSnapshot db
    mNamingManagedNameRefreshCount = mNamingManagedNameRefreshCount + 1
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    InvalidateNamingSnapshot
    On Error GoTo 0
    Err.Raise savedErrNumber, IIf(Len(savedErrSource) > 0, savedErrSource, "RefreshManagedNameSnapshot"), savedErrDescription
End Sub

Private Sub LoadManagedNameSnapshot(ByVal db As DAO.Database)
    'Loads managed logical-to-physical names from the supplied database into the naming snapshot.
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim tdf As DAO.tableDef
    Dim tableName As String
    Dim keyText As String
    Set mNamingSnapshotManagedNames = CreateObject("Scripting.Dictionary")
    mNamingSnapshotManagedNames.CompareMode = vbTextCompare
    tableName = BootstrapManagedObjectsTableName()
    If DaoTableExists(db, tableName) = False Then Exit Sub
    Set tdf = db.TableDefs(tableName)
    If DAOFieldExists(tdf, FIELD_LOGICAL_OBJECT_NAME) = False Then Exit Sub
    If DAOFieldExists(tdf, FIELD_OBJECT_NAMING_CATEGORY) = False Then Exit Sub
    Set rs = db.OpenRecordset("SELECT [" & FIELD_OBJECT_NAMING_CATEGORY & "],[" & FIELD_LOGICAL_OBJECT_NAME & "]," & SQL_FIELD_OBJECT_NAME & " FROM [" & Replace(tableName, "]", "]]") & "] WHERE Nz([" & FIELD_MANAGED_OBJECT_STATUS & "],'" & ADCAT_STATUS_ACTIVE & "')='" & ADCAT_STATUS_ACTIVE & "' AND Len(Nz([" & FIELD_LOGICAL_OBJECT_NAME & "],''))>0 AND Len(Nz([" & FIELD_OBJECT_NAMING_CATEGORY & "],''))>0 ORDER BY [" & FIELD_OBJECT_NAMING_CATEGORY & "],[" & FIELD_LOGICAL_OBJECT_NAME & "]," & SQL_FIELD_ID & ";", dbOpenSnapshot)
    Do While rs.EOF = False
        keyText = NamingSnapshotManagedNameKey(CStr(rs.fields(FIELD_OBJECT_NAMING_CATEGORY).value), CStr(rs.fields(FIELD_LOGICAL_OBJECT_NAME).value))
        If mNamingSnapshotManagedNames.Exists(keyText) Then Err.Raise vbObjectError + 6803, "LoadManagedNameSnapshot", "More than one active managed-object row has the same naming category and logical name." & vbCrLf & "Key: " & keyText
        mNamingSnapshotManagedNames.Add keyText, CStr(Nz(rs.fields(FIELD_OBJECT_NAME).value, vbNullString))
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
    Err.Raise savedErrNumber, IIf(Len(savedErrSource) > 0, savedErrSource, "LoadManagedNameSnapshot"), savedErrDescription
End Sub

Private Sub ApplyNamingPolicySetting(ByVal SettingKey As String, ByVal SettingValue As Variant)
    'Applies one persisted naming-policy setting to the in-memory naming policy.
    Dim textValue As String
    If IsNull(SettingValue) Then textValue = vbNullString Else textValue = CStr(SettingValue)
    Select Case UCase$(Trim$(SettingKey))
        Case UCase$(SETTING_TABLE_OBJECT_PREFIX): mNamingTableObjectPrefix = textValue
        Case UCase$(SETTING_TABLE_OBJECT_POSTFIX): mNamingTableObjectPostfix = textValue
        Case UCase$(SETTING_QUERY_OBJECT_PREFIX): mNamingQueryObjectPrefix = textValue
        Case UCase$(SETTING_QUERY_OBJECT_POSTFIX): mNamingQueryObjectPostfix = textValue
        Case UCase$(SETTING_GENERATED_MACRO_OBJECT_PREFIX): mNamingGeneratedMacroObjectPrefix = textValue
        Case UCase$(SETTING_GENERATED_MACRO_OBJECT_POSTFIX): mNamingGeneratedMacroObjectPostfix = textValue
        Case UCase$(SETTING_SYSTEM_MACRO_OBJECT_PREFIX): mNamingSystemMacroObjectPrefix = textValue
        Case UCase$(SETTING_SYSTEM_MACRO_OBJECT_POSTFIX): mNamingSystemMacroObjectPostfix = textValue
        Case UCase$(SETTING_CUSTOM_FORM_OBJECT_PREFIX): mNamingCustomFormObjectPrefix = textValue
        Case UCase$(SETTING_CUSTOM_FORM_OBJECT_POSTFIX): mNamingCustomFormObjectPostfix = textValue
        Case UCase$(SETTING_NAVIGATION_PANE_GROUP_PREFIX): mNamingNavigationPaneGroupPrefix = textValue
        Case UCase$(SETTING_NAVIGATION_PANE_GROUP_POSTFIX): mNamingNavigationPaneGroupPostfix = textValue
        Case UCase$(SETTING_SYSTEM_MACRO_NAVIGATION_PANE_GROUP_PREFIX): mNamingSystemMacroNavigationPaneGroupPrefix = textValue
        Case UCase$(SETTING_SYSTEM_MACRO_NAVIGATION_PANE_GROUP_POSTFIX): mNamingSystemMacroNavigationPaneGroupPostfix = textValue
        Case UCase$(SETTING_INCLUDE_SYSTEM_MACRO_DESIGNATOR_IN_RIBBON_TITLE): mNamingIncludeSystemMacroDesignatorInRibbonTitle = BooleanTextValue(textValue, DEFAULT_INCLUDE_SYSTEM_MACRO_DESIGNATOR_IN_RIBBON_TITLE)
        Case UCase$(SETTING_INCLUDE_SYSTEM_MACRO_DESIGNATOR_IN_NAVIGATION_PANE_TITLE): mNamingIncludeSystemMacroDesignatorInNavigationPaneTitle = BooleanTextValue(textValue, DEFAULT_INCLUDE_SYSTEM_MACRO_DESIGNATOR_IN_NAVIGATION_PANE_TITLE)
    End Select
End Sub


















' Purpose: Adds a keyed text marker once, creating the Collection when needed.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   values (Input/output; Collection): Caller-supplied value collection used as the source set for this operation; when passed ByRef the routine may append/update items according to its documented behavior.
'   keyText (Input; String): Input key text for mark collection text key; is tested by If SharedCollection_KeyExists(values, keyText) = False.
' Returns: None (Sub). Results are returned through the documented ByRef parameters and/or the routine's database, UI, or module-state side effects.
Public Sub MarkCollectionTextKey(ByRef values As Collection, ByVal keyText As String)
    'Adds a keyed text marker once, creating the Collection when needed.
    On Error Resume Next
    ' INTENTIONAL FILTERED PARTITION: values is a consumer-required subset/projection derived from the source enumerated by MarkCollectionTextKey; its membership is not independently authoritative.
    If values Is Nothing Then Set values = New Collection
    If SharedCollection_KeyExists(values, keyText) = False Then values.Add keyText, keyText
End Sub

' Purpose: Marks the naming snapshot stale so the next naming lookup rebuilds current state.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub InvalidateNamingSnapshot()
    'Marks the naming snapshot stale so the next naming lookup rebuilds current state.
    mNamingSnapshotInvalidationCount = mNamingSnapshotInvalidationCount + 1
    ClearNamingSnapshotState
End Sub







' Purpose: Processes setup responsiveness, coordinating `Static` and `ElapsedSecondsFromTimer` as the implementation requires.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   forceDispatch (Input Optional; Boolean): When True, enables or requests dispatch for this call; when False, the routine follows its normal non-forced/default path. Default: False.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub PumpSetupResponsiveness(Optional ByVal forceDispatch As Boolean = False)
    'Intended use: Uses the existing setup-control pump when Setup Control is active and otherwise yields directly; ordinary calls remain interval-throttled, while forceDispatch=True yields immediately at explicit
    'scan phase boundaries.
    Static lastDispatchTimer As Single
    Static dispatchActive As Boolean
    On Error Resume Next
    If dispatchActive Then Exit Sub
    If forceDispatch = False Then
        If lastDispatchTimer > 0 Then
            If ElapsedSecondsFromTimer(lastDispatchTimer) < DEFAULT_SETUP_EVENT_PUMP_INTERVAL_SECONDS Then Exit Sub
        End If
    End If
    dispatchActive = True
    lastDispatchTimer = Timer
    If mSetupControlActive Then
        Call SetupControlPump(True)
    Else
        DoEvents
    End If
    dispatchActive = False
    On Error GoTo 0
End Sub




' Purpose: Returns the Defaults-owned ChatGPTUIReport classification for an exact canonical physical report name. False/noncanonical rows remain outside this generated-report namespace.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   physicalReportName (Input; String): Exact physical Access report name used for managed-report classification against Defaults.
'   isCanonicalReport (Output; Boolean): Set True when the procedure determines that is canonical report applies; set False when that condition is not met.
' Returns: Boolean: True when DAO table exists in the inspected database/metadata scope; False when it does not.
Public Function ManagedReportChatGPTUIClassification(ByVal physicalReportName As String, ByRef isCanonicalReport As Boolean) As Boolean
    'Returns the Defaults-owned ChatGPTUIReport classification for an exact canonical physical report name. False/noncanonical rows remain outside this generated-report namespace.
    Dim logicalName As String
    Dim chatGPTUIReport As Boolean
    isCanonicalReport = False
    If CanonicalReportObjectName(physicalReportName, True, logicalName, chatGPTUIReport) = False Then Exit Function
    isCanonicalReport = True
    ManagedReportChatGPTUIClassification = chatGPTUIReport
End Function

' Purpose: Returns the effective setting controlling System Macro designators in Ribbon titles.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: Boolean: True when the condition represented by EnsureTableDefFieldOrderFromText is present under the routine's documented inputs/current state; False when that condition is absent.
Public Function IncludeSystemMacroDesignatorInRibbonTitle() As Boolean
    'Returns the effective setting controlling System Macro designators in Ribbon titles.
    If mNamingSnapshotActive And mNamingPolicyReady Then
        IncludeSystemMacroDesignatorInRibbonTitle = mNamingIncludeSystemMacroDesignatorInRibbonTitle
    Else
        IncludeSystemMacroDesignatorInRibbonTitle = NamingSettingBoolean(SETTING_INCLUDE_SYSTEM_MACRO_DESIGNATOR_IN_RIBBON_TITLE, DEFAULT_INCLUDE_SYSTEM_MACRO_DESIGNATOR_IN_RIBBON_TITLE)
    End If
End Function

Private Function IncludeSystemMacroDesignatorInNavigationPaneTitle() As Boolean
    'Returns the effective setting controlling System Macro designators in Navigation Pane titles.
    If mNamingSnapshotActive And mNamingPolicyReady Then
        IncludeSystemMacroDesignatorInNavigationPaneTitle = mNamingIncludeSystemMacroDesignatorInNavigationPaneTitle
    Else
        IncludeSystemMacroDesignatorInNavigationPaneTitle = NamingSettingBoolean(SETTING_INCLUDE_SYSTEM_MACRO_DESIGNATOR_IN_NAVIGATION_PANE_TITLE, DEFAULT_INCLUDE_SYSTEM_MACRO_DESIGNATOR_IN_NAVIGATION_PANE_TITLE)
    End If
End Function

' Purpose: Sets naming bootstrap table names for settings table name, managed objects table name.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   settingsTableName (Input; String): Physical settings table name used as the schema/data target for the operation.
'   managedObjectsTableName (Input; String): Physical managed-objects catalog table name used by the catalog/query operation.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub SetNamingBootstrapTableNames(ByVal settingsTableName As String, ByVal managedObjectsTableName As String)
    'Caches the bootstrap Settings and Managed Objects table names used before the naming snapshot is ready.
    Dim db As DAO.Database
    InvalidateNamingSnapshot
    Set db = currentDb
    If Len(settingsTableName) > 0 Then Call EnsureDaoPropertyValue(db, BOOTSTRAP_SETTINGS_TABLE_PROPERTY, dbText, settingsTableName, "ADCSetDatabaseTextProperty")
    If Len(managedObjectsTableName) > 0 Then Call EnsureDaoPropertyValue(db, BOOTSTRAP_MANAGED_OBJECTS_TABLE_PROPERTY, dbText, managedObjectsTableName, "ADCSetDatabaseTextProperty")
End Sub





' Purpose: Checks whether DAO table exists in the database, collection, or metadata scope supplied to this function.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   db (Input; DAO.Database): DAO.Database instance whose schema, data, queries, or properties the routine reads or modifies.
'   tableName (Input; String): Physical Access table name identifying the table whose schema, data, or metadata the routine processes.
' Returns: Boolean: True when DAO table exists in the inspected context; False when it does not.
Public Function DaoTableExists(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
    'Intended use: Returns whether a named table exists in the supplied DAO database, treating only DAO 3265 as absence and preserving unexpected lookup failures.
    On Error GoTo NotFound
    Dim tdf As DAO.tableDef
    Set tdf = db.TableDefs(tableName)
    DaoTableExists = True
    Exit Function
NotFound:
    If Err.Number = 3265 Then
        Err.Clear
        Exit Function
    End If
    Err.Raise Err.Number, "DaoTableExists", Err.Description
End Function
' Purpose: Returns pipe-delimited schema field-order text from the supplied schema-definition collection in its existing authoritative order.
' Public role: Inter-module public API shared by modules that consume collection-based field schema definitions.
' Parameters:
'   schemaDefinitions (Input; Collection): Ordered schema-definition rows whose element 0 contains each field name.
' Returns: String: Pipe-delimited field names in the same order supplied by schemaDefinitions.
Public Function SchemaDefinitionFieldOrderText(ByVal schemaDefinitions As Collection) As String
    Dim fieldDefinition As Variant
    Dim valueText As String
    For Each fieldDefinition In schemaDefinitions
        If Len(valueText) > 0 Then valueText = valueText & "|"
        valueText = valueText & CStr(fieldDefinition(0))
    Next fieldDefinition
    SchemaDefinitionFieldOrderText = valueText
End Function




' Purpose: Ensures table def field order from text for tdf, field order text.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   tdf (Input; DAO.tableDef): DAO.TableDef representing the table whose fields, indexes, properties, or metadata are being inspected or modified.
'   fieldOrderText (Input; String): Input field order text for ensure table def field order from text; is tested by If Len(fieldOrderText) = 0; is passed to `Split`.
'   designChangeAlreadyPrepared (Input Optional; Boolean): Boolean state describing whether design change already prepared applies to this invocation; True selects the named condition/path and False selects its
'   complementary/default path. Default: False.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub EnsureTableDefFieldOrderFromText(ByVal tdf As DAO.tableDef, ByVal fieldOrderText As String, Optional ByVal designChangeAlreadyPrepared As Boolean = False)
    'Generic physical DAO field-order mechanic. The caller supplies the authoritative domain-owned field-order specification.
    On Error GoTo ErrHandler
    Dim desiredFields As Variant
    Dim desiredNames As Object
    Dim extraFieldNames As Collection
    Dim fld As DAO.Field
    Dim fieldName As String
    Dim tableNameText As String
    Dim fieldIndex As Long
    Dim expectedPosition As Long
    Dim changedCount As Long
    If tdf Is Nothing Then Err.Raise vbObjectError + 5902, "EnsureTableDefFieldOrderFromText", "TableDef reference is Nothing."
    tableNameText = tdf.Name
    If Len(fieldOrderText) = 0 Then Exit Sub
    ' INTENTIONAL KEYED LOOKUP: desiredNames provides keyed membership/lookup over source rows or state already read by EnsureTableDefFieldOrderFromText; it is a derived access path, not a second authority.
    ' INTENTIONAL ORDERED VIEW: extraFieldNames preserves the deterministic sequence required by EnsureTableDefFieldOrderFromText; keyed-only or unordered storage would not preserve the consumer semantics.
    Set desiredNames = CreateObject("Scripting.Dictionary")
    desiredNames.CompareMode = vbTextCompare
    Set extraFieldNames = New Collection
    desiredFields = Split(fieldOrderText, "|")
    If designChangeAlreadyPrepared = False Then tdf.fields.Refresh
    For fieldIndex = LBound(desiredFields) To UBound(desiredFields)
        PumpSetupResponsiveness
        fieldName = CStr(desiredFields(fieldIndex))
        If Len(fieldName) = 0 Then Err.Raise vbObjectError + 5903, "EnsureTableDefFieldOrderFromText", "The field-order definition contains a blank field name."
        If desiredNames.Exists(fieldName) Then Err.Raise vbObjectError + 5904, "EnsureTableDefFieldOrderFromText", "The field-order definition contains a duplicate field name: " & fieldName
        desiredNames.Add fieldName, True
        If SharedDao_TableDefFieldExistsSafe(tdf, fieldName) = False Then Err.Raise vbObjectError + 5905, "EnsureTableDefFieldOrderFromText", "The field-order definition references a missing field." & vbCrLf & "Table: " & tableNameText & vbCrLf & "Field: " & fieldName
    Next fieldIndex
    For Each fld In tdf.fields
        PumpSetupResponsiveness
        If desiredNames.Exists(fld.Name) = False Then extraFieldNames.Add fld.Name
    Next fld
    expectedPosition = 0
    For fieldIndex = LBound(desiredFields) To UBound(desiredFields)
        PumpSetupResponsiveness
        fieldName = CStr(desiredFields(fieldIndex))
        Set fld = tdf.fields(fieldName)
        If fld.ordinalPosition <> expectedPosition Then
            fld.ordinalPosition = expectedPosition
            changedCount = changedCount + 1
        End If
        If EnsureFieldColumnOrder(fld, tableNameText, expectedPosition + 1) Then changedCount = changedCount + 1
        expectedPosition = expectedPosition + 1
    Next fieldIndex
    For fieldIndex = 1 To extraFieldNames.count
        PumpSetupResponsiveness
        fieldName = CStr(extraFieldNames(fieldIndex))
        Set fld = tdf.fields(fieldName)
        If fld.ordinalPosition <> expectedPosition Then
            fld.ordinalPosition = expectedPosition
            changedCount = changedCount + 1
        End If
        If EnsureFieldColumnOrder(fld, tableNameText, expectedPosition + 1) Then changedCount = changedCount + 1
        expectedPosition = expectedPosition + 1
    Next fieldIndex
    If changedCount > 0 Then
        tdf.fields.Refresh
        CoreAddOperationNote "Repaired logical field order for table [" & tableNameText & "]."
    End If
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    If Len(tableNameText) = 0 Then tableNameText = "(unresolved)"
    Err.Raise originalErrNumber, "EnsureTableDefFieldOrderFromText", "Unable to enforce logical table field order." & vbCrLf & "Table: " & tableNameText & vbCrLf & "Description: " & originalErrDescription
End Sub


' Purpose: Ensures single field index for DB, table name.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   db (Input; DAO.Database): DAO.Database instance whose schema, data, queries, or properties the routine reads or modifies.
'   tableName (Input; String): Physical Access table name identifying the table whose schema, data, or metadata the routine processes.
'   indexName (Input; String): Exact DAO index name identifying the index to create, validate, or inspect.
'   fieldName (Input; String): Field name identifying the table/query field or bound control source the routine evaluates or modifies.
'   uniqueIndex (Input; Boolean): True creates/validates the index as unique; False permits duplicate key values.
'   primaryIndex (Input; Boolean): True marks the index as the table primary key; False creates/validates it as a non-primary index.
'   designChangeAlreadyPrepared (Input Optional; Boolean): Boolean state describing whether design change already prepared applies to this invocation; True selects the named condition/path and False selects its
'   complementary/default path. Default: False.
'   prepareOpenObjectForDesignChange (Input Optional; Boolean): When True, enables or requests open object for design change for this call; when False, the routine follows its normal non-forced/default path.
'   Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub EnsureSingleFieldIndex(ByVal db As DAO.Database, ByVal tableName As String, ByVal indexName As String, ByVal fieldName As String, ByVal uniqueIndex As Boolean, ByVal primaryIndex As Boolean, Optional ByVal designChangeAlreadyPrepared As Boolean = False, Optional ByVal prepareOpenObjectForDesignChange As Boolean = True)
    'Generic physical DAO index mechanic. The caller supplies the database and authoritative index specification; no domain batch state is consulted here.
    On Error GoTo ErrHandler
    Dim tdf As DAO.tableDef
    Dim idx As DAO.index
    If db Is Nothing Then Err.Raise vbObjectError + 5252, "EnsureSingleFieldIndex", "Database reference is Nothing."
    If designChangeAlreadyPrepared = False And prepareOpenObjectForDesignChange Then db.TableDefs.Refresh
    Set tdf = db.TableDefs(tableName)
    If designChangeAlreadyPrepared = False Then
        If prepareOpenObjectForDesignChange Then tdf.fields.Refresh
        tdf.Indexes.Refresh
    End If
    If SharedDao_TableDefSingleFieldIndexMatches(tdf, indexName, fieldName, uniqueIndex, primaryIndex) Then Exit Sub
    If designChangeAlreadyPrepared = False And prepareOpenObjectForDesignChange Then
        If SharedAccess_PrepareOpenObjectForDesignChange(acTable, tableName, "creating or repairing index [" & indexName & "] on table [" & tableName & "]") = False Then
            MarkSystemCommandCancelled
            Err.Raise vbObjectError + 9004, "EnsureSingleFieldIndex", "Index repair was cancelled because the affected open table was not approved for save and close."
        End If
        Set tdf = Nothing
        Set db = Nothing
        Set db = codeDb
        db.TableDefs.Refresh
        Set tdf = db.TableDefs(tableName)
        tdf.fields.Refresh
    End If
    If DAOFieldExists(tdf, fieldName) = False Then
        tdf.fields.Refresh
        If DAOFieldExists(tdf, fieldName) = False Then Err.Raise vbObjectError + 5254, "EnsureSingleFieldIndex", "The requested index field is not visible in the table definition." & vbCrLf & "Table: " & tableName & vbCrLf & "Index: " & indexName & vbCrLf & "Field: " & fieldName
    End If
    If SharedDao_TableDefSingleFieldIndexMatches(tdf, indexName, fieldName, uniqueIndex, primaryIndex) Then Exit Sub
    If primaryIndex Then DropConflictingPrimaryIndexes tdf, indexName, designChangeAlreadyPrepared = False
    DropIndexIfExists tdf, indexName, designChangeAlreadyPrepared = False
    Set idx = tdf.CreateIndex(indexName)
    idx.fields.Append idx.CreateField(fieldName)
    idx.Unique = uniqueIndex
    idx.Primary = primaryIndex
    tdf.Indexes.Append idx
    If designChangeAlreadyPrepared = False Then tdf.Indexes.Refresh
    If SharedDao_TableDefSingleFieldIndexMatches(tdf, indexName, fieldName, uniqueIndex, primaryIndex) = False Then
        tdf.Indexes.Refresh
        If SharedDao_TableDefSingleFieldIndexMatches(tdf, indexName, fieldName, uniqueIndex, primaryIndex) = False Then Err.Raise vbObjectError + 5253, "EnsureSingleFieldIndex", "The index was not visible or did not match after DAO append." & vbCrLf & "Table: " & tableName & vbCrLf & "Index: " & indexName
    End If
    If prepareOpenObjectForDesignChange Then CoreAddOperationNote "Created/repaired index [" & indexName & "] on table [" & tableName & "]."
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "EnsureSingleFieldIndex", "Could not create or repair index." & vbCrLf & "Table: " & tableName & vbCrLf & "Index: " & indexName & vbCrLf & "Field: " & fieldName & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

Private Sub DropConflictingPrimaryIndexes(ByVal tdf As DAO.tableDef, ByVal desiredIndexName As String, Optional ByVal refreshAfterDelete As Boolean = True)
    'Removes primary indexes that conflict with the desired primary-index definition.
    On Error GoTo ErrHandler
    Dim idx As DAO.index
    Dim namesToDrop As Collection
    Dim item As Variant
    ' INTENTIONAL FILTERED PARTITION: namesToDrop is a consumer-required subset/projection derived from the source enumerated by DropConflictingPrimaryIndexes; its membership is not independently authoritative.
    Set namesToDrop = New Collection
    For Each idx In tdf.Indexes
        PumpSetupResponsiveness
        If CBool(idx.Primary) Then
            If StrComp(idx.Name, desiredIndexName, vbTextCompare) <> 0 Then namesToDrop.Add idx.Name
        End If
    Next idx
    For Each item In namesToDrop
        PumpSetupResponsiveness
        tdf.Indexes.Delete CStr(item)
    Next item
    If namesToDrop.count > 0 And refreshAfterDelete Then tdf.Indexes.Refresh
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "DropConflictingPrimaryIndexes", "Could not delete conflicting primary index before rebuilding the requested primary index." & vbCrLf & "Table: " & tdf.Name & vbCrLf & "Desired index: " & desiredIndexName & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

Private Sub DropIndexIfExists(ByVal tdf As DAO.tableDef, ByVal indexName As String, Optional ByVal refreshAfterDelete As Boolean = True)
    'Removes the named DAO index when it exists on the supplied table definition.
    On Error GoTo ErrHandler
    If SharedDao_TableDefIndexExistsSafe(tdf, indexName) Then
        tdf.Indexes.Delete indexName
        If refreshAfterDelete Then tdf.Indexes.Refresh
    End If
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "DropIndexIfExists", "Could not delete invalid index before rebuilding it." & vbCrLf & "Table: " & tdf.Name & vbCrLf & "Index: " & indexName & vbCrLf & "Err.Description: " & originalErrDescription
End Sub





' Purpose: Displays a utility-branded message box after applying suppression and human-readable diagnostic handling.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   prompt (Input; Variant): Prompt text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   buttons (Input Optional; VbMsgBoxStyle): Message-box style/icon flags controlling the visual treatment and available choices for this UI operation. Default: vbOKOnly.
'   Title (Input Optional; Variant): Title text used as the title shown, stored, compared, or resolved by this operation.
' Returns: VbMsgBoxResult: Returns the button/result produced by the ADC message-box wrapper, including the configured suppression/default behavior when an actual popup is not shown.
Public Function UtilityMsgBox(ByVal prompt As Variant, Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As Variant) As VbMsgBoxResult
    'Displays a utility-branded message box after applying suppression and human-readable diagnostic handling.
    On Error GoTo UseOriginal
    Dim TitleText As String
    Dim displayText As String
    TitleText = UtilityLongName()
    If isMissing(Title) = False Then TitleText = ApplyDisplayBranding(CStr(Title))
    displayText = HumanReadablePopupText(CStr(prompt), TitleText, buttons)
    UtilityMsgBox = VBA.MsgBox(ApplyDisplayBranding(displayText), buttons, TitleText)
    Exit Function
UseOriginal:
    Err.Clear
    UtilityMsgBox = VBA.MsgBox(ApplyDisplayBranding(CStr(prompt)), buttons, TitleText)
End Function

' Purpose: Adds a Cancel suppression path while preserving non-button UtilityMsgBox style flags.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   requestedStyle (Input; VbMsgBoxStyle): Message-box style/icon flags controlling the visual treatment and available choices for this UI operation.
'   resultPrompt (Input; Boolean): True requests the routine’s result/confirmation prompt path; False suppresses that prompt and returns/continues through the non-prompt path.
' Returns: VbMsgBoxStyle: Returns the message-box style flags after adding the cancellation/suppression choice required by the current ADC prompt policy.
Public Function MsgBoxStyleWithSuppressionCancel(ByVal requestedStyle As VbMsgBoxStyle, ByVal resultPrompt As Boolean) As VbMsgBoxStyle
    'Adds a Cancel suppression path while preserving non-button UtilityMsgBox style flags.
    Dim styleValue As Long
    Dim nonButtonStyle As Long
    Dim buttonStyle As Long
    styleValue = CLng(requestedStyle)
    buttonStyle = styleValue And 7
    nonButtonStyle = styleValue And Not 7
    If resultPrompt Then
        Select Case buttonStyle
            Case vbOKOnly
                MsgBoxStyleWithSuppressionCancel = nonButtonStyle Or vbOKCancel
            Case vbYesNo
                MsgBoxStyleWithSuppressionCancel = nonButtonStyle Or vbYesNoCancel
            Case Else
                MsgBoxStyleWithSuppressionCancel = requestedStyle
        End Select
    Else
        MsgBoxStyleWithSuppressionCancel = nonButtonStyle Or vbOKCancel
    End If
End Function












' Purpose: Returns True when the current Access Development Console run has suppressed additional popup messages.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: Boolean: True when the current Access Development Console run has suppressed additional popup messages. False otherwise.
Public Function PopupMessagesSuppressedForCurrentRun() As Boolean
    'Returns True when the current Access Development Console run has suppressed additional popup messages.
    On Error GoTo ErrHandler
    PopupMessagesSuppressedForCurrentRun = CBool(TempVars(TEMPVAR_SUPPRESS_POPUPS).value)
    Exit Function
ErrHandler:
    PopupMessagesSuppressedForCurrentRun = False
End Function


' Purpose: Returns the safe current-run result when a result-bearing popup has already been suppressed.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   requestedStyle (Input; VbMsgBoxStyle): Message-box style/icon flags controlling the visual treatment and available choices for this UI operation.
' Returns: VbMsgBoxResult: Returns the button result ADC should assume when the specified popup is suppressed rather than displayed.
Public Function SuppressedPopupDefaultResult(ByVal requestedStyle As VbMsgBoxStyle) As VbMsgBoxResult
    'Returns the safe current-run result when a result-bearing popup has already been suppressed.
    Dim buttonStyle As Long
    buttonStyle = CLng(requestedStyle) And 7
    Select Case buttonStyle
        Case vbOKOnly
            SuppressedPopupDefaultResult = vbOK
        Case vbOKCancel, vbYesNoCancel, vbRetryCancel
            SuppressedPopupDefaultResult = vbCancel
        Case vbYesNo
            SuppressedPopupDefaultResult = vbNo
        Case vbAbortRetryIgnore
            SuppressedPopupDefaultResult = vbAbort
        Case Else
            SuppressedPopupDefaultResult = vbNo
    End Select
End Function


' Purpose: Returns vbExclamation for warning-bearing completion or preserves the caller-selected icon for clean completion.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   requestedIcon (Input; VbMsgBoxStyle): Message-box style/icon flags controlling the visual treatment and available choices for this UI operation.
'   warningDetected (Input; Boolean): Indicates whether warnings have already been detected for the operation; the routine preserves/updates warning handling based on this state.
' Returns: VbMsgBoxStyle: Returns the message-box icon/style appropriate for the supplied completion outcome and warning/error state.
Public Function CompletionPopupIcon(ByVal requestedIcon As VbMsgBoxStyle, ByVal warningDetected As Boolean) As VbMsgBoxStyle
    'Returns vbExclamation for warning-bearing completion or preserves the caller-selected icon for clean completion.
    If warningDetected Then
        CompletionPopupIcon = vbExclamation
    Else
        CompletionPopupIcon = requestedIcon
    End If
End Function





Private Sub ReplaceDaoPropertyPreservingOriginal(ByVal targetObject As Object, ByVal propertyName As String, ByVal requestedType As DAO.DataTypeEnum, ByVal requestedValue As Variant, ByVal originalType As DAO.DataTypeEnum, ByVal originalValue As Variant, ByVal sourceProcedure As String, ByVal targetDiagnosticContext As String, ByVal replacementReason As String, ByVal assignmentErrNumber As Long, ByVal assignmentErrSource As String, ByVal assignmentErrDescription As String, ByVal assignmentDaoErrorsText As String)
    'Replaces an incompatible DAO property while retaining enough original state to restore it if replacement fails.
    On Error GoTo ReplacementFailed
    Dim prp As DAO.Property
    Dim replacementStage As String
    Dim targetOwnerDatabase As DAO.Database
    Dim propertyExists As Boolean
    Dim lookupErrNumber As Long
    Dim lookupErrSource As String
    Dim lookupErrDescription As String
    replacementStage = "Delete existing DAO property"
    targetObject.properties.Delete propertyName
    RefreshDaoProperties targetObject
    Set prp = Nothing
    Set targetObject = Nothing
    DBEngine.Idle dbRefreshCache
    replacementStage = "Reacquire DAO property target after deletion through a fresh database handle"
    Set targetObject = ReacquireDaoPropertyTarget(targetObject, targetDiagnosticContext, sourceProcedure, targetOwnerDatabase)
    replacementStage = "Verify DAO property deletion"
    ReadDaoPropertyReference targetObject, propertyName, prp, propertyExists, lookupErrNumber, lookupErrSource, lookupErrDescription
    If propertyExists Then Err.Raise vbObjectError + 6890, sourceProcedure, "The DAO property remained present after deletion."
    If lookupErrNumber <> 0 And lookupErrNumber <> 3270 Then Err.Raise lookupErrNumber, lookupErrSource, lookupErrDescription
    replacementStage = "Create detached replacement DAO property with explicit text capacity"
    Set prp = CreateDetachedDaoProperty(targetObject, targetOwnerDatabase, propertyName, requestedType, requestedValue)
    replacementStage = "Append replacement DAO property"
    targetObject.properties.Append prp
    RefreshDaoProperties targetObject
    replacementStage = "Reacquire replacement DAO property"
    Set prp = targetObject.properties(propertyName)
    replacementStage = "Assign complete requested value to replacement DAO property"
    prp.value = requestedValue
    RefreshDaoProperties targetObject
    replacementStage = "Reacquire replacement DAO property for verification"
    Set prp = targetObject.properties(propertyName)
    replacementStage = "Verify replacement DAO property type and value"
    VerifyDaoPropertyValue prp, propertyName, requestedType, requestedValue, sourceProcedure
    Exit Sub
ReplacementFailed:
    Dim replacementErrNumber As Long
    Dim replacementErrSource As String
    Dim replacementErrDescription As String
    Dim replacementDaoErrorsText As String
    Dim restorationSucceeded As Boolean
    Dim restorationDetail As String
    replacementErrNumber = Err.Number
    replacementErrSource = Err.Source
    replacementErrDescription = Err.Description
    replacementDaoErrorsText = DaoEngineErrorsText()
    restorationSucceeded = TryRestoreDaoProperty(targetObject, propertyName, originalType, originalValue, sourceProcedure, targetDiagnosticContext, restorationDetail)
    Err.Raise replacementErrNumber, sourceProcedure, _
        "DAO property replacement failed; the original property restoration result is recorded below." & vbCrLf & _
        "Caller: " & sourceProcedure & vbCrLf & _
        DaoTargetContextText(targetDiagnosticContext) & _
        "Property: " & propertyName & vbCrLf & _
        "Replacement reason: " & replacementReason & vbCrLf & _
        "Requested DAO type: " & CStr(requestedType) & vbCrLf & _
        "Existing DAO type: " & CStr(originalType) & vbCrLf & _
        "Requested characters: " & DaoPropertyCharacterLengthText(requestedValue) & vbCrLf & _
        "Requested bytes: " & DaoPropertyByteLengthText(requestedValue) & vbCrLf & _
        "Original characters: " & DaoPropertyCharacterLengthText(originalValue) & vbCrLf & _
        "Original bytes: " & DaoPropertyByteLengthText(originalValue) & vbCrLf & _
        "Original assignment error: " & CStr(assignmentErrNumber) & vbCrLf & _
        "Original assignment source: " & assignmentErrSource & vbCrLf & _
        "Original assignment description: " & assignmentErrDescription & _
        IIf(Len(assignmentDaoErrorsText) > 0, vbCrLf & "Original assignment DBEngine.Errors:" & vbCrLf & assignmentDaoErrorsText, vbNullString) & vbCrLf & _
        "Replacement stage: " & replacementStage & vbCrLf & _
        "Restoration succeeded: " & CStr(restorationSucceeded) & vbCrLf & _
        "Restoration detail: " & restorationDetail & vbCrLf & _
        "Final DAO error number: " & CStr(replacementErrNumber) & vbCrLf & _
        "Final DAO error source: " & replacementErrSource & vbCrLf & _
        "Final DAO error description: " & replacementErrDescription & _
        IIf(Len(replacementDaoErrorsText) > 0, vbCrLf & "Replacement DBEngine.Errors:" & vbCrLf & replacementDaoErrorsText, vbNullString)
End Sub






Private Sub VerifyDaoPropertyValue(ByVal prp As DAO.Property, ByVal propertyName As String, ByVal requestedType As DAO.DataTypeEnum, ByVal requestedValue As Variant, ByVal sourceProcedure As String)
    'Verifies that a DAO property has the requested type and value and raises when the postcondition is not met.
    If prp Is Nothing Then Err.Raise vbObjectError + 6883, sourceProcedure, "The DAO property could not be reacquired for verification."
    If CLng(prp.Type) <> CLng(requestedType) Then Err.Raise vbObjectError + 6884, sourceProcedure, "The DAO property type did not verify after ensure/recreation." & vbCrLf & "Property: " & propertyName & vbCrLf & "Expected type: " & CStr(requestedType) & vbCrLf & "Actual type: " & CStr(prp.Type)
    If DaoPropertyValuesEqual(prp.value, requestedValue, requestedType) = False Then Err.Raise vbObjectError + 6885, sourceProcedure, "The DAO property value did not verify after ensure/recreation." & vbCrLf & "Property: " & propertyName & vbCrLf & "Expected value: " & DaoPropertyDiagnosticValue(requestedValue) & vbCrLf & "Actual value: " & DaoPropertyDiagnosticValue(prp.value)
End Sub

Private Sub RefreshDaoProperties(ByVal targetObject As Object)
    'Refreshes the Properties collection on a DAO object while safely tolerating unsupported refresh operations.
    On Error Resume Next
    targetObject.properties.Refresh
    Err.Clear
    On Error GoTo 0
End Sub


Private Function DaoPropertyByteArraysEqual(ByVal actualValue As Variant, ByVal requestedValue As Variant) As Boolean
    'Returns True when two DAO byte-array values contain identical bytes.
    On Error GoTo NotEqual
    Dim indexValue As Long
    If IsArray(actualValue) = False Or IsArray(requestedValue) = False Then Exit Function
    If LBound(actualValue) <> LBound(requestedValue) Or UBound(actualValue) <> UBound(requestedValue) Then Exit Function
    For indexValue = LBound(actualValue) To UBound(actualValue)
        If CByte(actualValue(indexValue)) <> CByte(requestedValue(indexValue)) Then Exit Function
    Next indexValue
    DaoPropertyByteArraysEqual = True
    Exit Function
NotEqual:
    DaoPropertyByteArraysEqual = False
End Function


' Purpose: Applies an already-resolved Access lookup specification to one DAO field. Domain modules remain responsible for deciding the lookup source and policy.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   fld (Input; DAO.Field): DAO.Field representing the field whose properties or metadata are being inspected or modified.
'   rowSourceType (Input; String): Access lookup RowSourceType text (for example Table/Query or Value List) to apply/validate on the field/control.
'   rowSourceText (Input; String): SQL/row-source text consumed by this procedure; it is parsed, compared, or assigned as coded without changing the caller's original string.
'   LimitToList (Input; Boolean): True restricts the Access lookup control/field to values present in its row source; False permits values outside the listed choices when Access allows them.
'   columnCount (Input; Long): Supplies the column count used to bound, report, or validate the number of items processed by this operation.
'   BoundColumn (Input; Long): One-based Access lookup column number whose value is stored as the control/field value.
'   columnWidthsText (Input; String): Input column widths text for apply lookup field properties; is tested by If (propertyMask And LOOKUP_FIELD_PROPERTY_COLUMN_WIDTHS) <> 0.
'   sourceProcedure (Input Optional; String): Procedure name recorded as the source of trace, error, status, or diagnostic information emitted by this call. Default: "ApplyLookupFieldProperties".
'   nothingErrorNumber (Input Optional; Long): Numeric identifier or ordinal for the nothing error item processed or reported by this operation. Default: vbObjectError + 6804.
'   blankNameErrorNumber (Input Optional; Long): Numeric identifier or ordinal for the blank name error item processed or reported by this operation. Default: vbObjectError + 6805.
'   propertyMask (Input Optional; LookupFieldPropertyMask): Bitmask selecting which lookup-field properties this call applies and verifies; the default applies the complete canonical lookup property set.
'   Default: LOOKUP_FIELD_PROPERTY_ALL.
' Returns: Boolean result of apply lookup field properties; True and False correspond to the affirmative and negative branches explicitly implemented by this function.
Public Function ApplyLookupFieldProperties(ByVal fld As DAO.Field, ByVal rowSourceType As String, ByVal rowSourceText As String, ByVal LimitToList As Boolean, ByVal columnCount As Long, ByVal BoundColumn As Long, ByVal columnWidthsText As String, Optional ByVal sourceProcedure As String = "ApplyLookupFieldProperties", Optional ByVal nothingErrorNumber As Long = vbObjectError + 6804, Optional ByVal blankNameErrorNumber As Long = vbObjectError + 6805, Optional ByVal propertyMask As LookupFieldPropertyMask = LOOKUP_FIELD_PROPERTY_ALL) As Boolean
    'Applies an already-resolved Access lookup specification to one DAO field. Domain modules remain responsible for deciding the lookup source and policy.
    Dim changed As Boolean
    If (propertyMask And LOOKUP_FIELD_PROPERTY_DISPLAY_CONTROL) <> 0 Then changed = EnsureDaoPropertyValue(fld, "DisplayControl", dbInteger, 111, sourceProcedure) Or changed
    If (propertyMask And LOOKUP_FIELD_PROPERTY_ROW_SOURCE_TYPE) <> 0 Then changed = EnsureLookupTextPropertyValue(fld, "RowSourceType", rowSourceType, sourceProcedure, nothingErrorNumber, blankNameErrorNumber) Or changed
    If (propertyMask And LOOKUP_FIELD_PROPERTY_ROW_SOURCE) <> 0 Then changed = EnsureLookupTextPropertyValue(fld, "RowSource", rowSourceText, sourceProcedure, nothingErrorNumber, blankNameErrorNumber) Or changed
    If (propertyMask And LOOKUP_FIELD_PROPERTY_BOUND_COLUMN) <> 0 Then changed = EnsureDaoPropertyValue(fld, "BoundColumn", dbInteger, BoundColumn, sourceProcedure) Or changed
    If (propertyMask And LOOKUP_FIELD_PROPERTY_COLUMN_COUNT) <> 0 Then changed = EnsureDaoPropertyValue(fld, "ColumnCount", dbInteger, columnCount, sourceProcedure) Or changed
    If (propertyMask And LOOKUP_FIELD_PROPERTY_LIMIT_TO_LIST) <> 0 Then changed = EnsureDaoPropertyValue(fld, "LimitToList", dbBoolean, LimitToList, sourceProcedure) Or changed
    If (propertyMask And LOOKUP_FIELD_PROPERTY_COLUMN_WIDTHS) <> 0 Then changed = EnsureLookupTextPropertyValue(fld, "ColumnWidths", columnWidthsText, sourceProcedure, nothingErrorNumber, blankNameErrorNumber) Or changed
    ApplyLookupFieldProperties = changed
End Function


Private Sub ReadDaoPropertyReference(ByVal targetObject As Object, ByVal propertyName As String, ByRef prp As DAO.Property, ByRef propertyExists As Boolean, ByRef lookupErrNumber As Long, ByRef lookupErrSource As String, ByRef lookupErrDescription As String)
    'Reads a DAO property reference and returns explicit existence and lookup-error details without raising expected missing-property errors.
    Set prp = Nothing
    propertyExists = False
    lookupErrNumber = 0
    lookupErrSource = vbNullString
    lookupErrDescription = vbNullString
    On Error Resume Next
    Err.Clear
    Set prp = targetObject.properties(propertyName)
    lookupErrNumber = Err.Number
    lookupErrSource = Err.Source
    lookupErrDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    propertyExists = (lookupErrNumber = 0 And Not prp Is Nothing)
End Sub



' Purpose: Removes one DAO database property when present, treats only property-not-found errors as already absent, and verifies deletion.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   db (Input; DAO.Database): DAO.Database instance whose schema, data, queries, or properties the routine reads or modifies.
'   propertyName (Input; String): Exact Access/DAO property name to read, compare, create, update, or fingerprint.
'   sourceProcedure (Input Optional; String): Procedure name recorded as the source of trace, error, status, or diagnostic information emitted by this call. Default: "EnsureDatabasePropertyRemoved".
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub EnsureDatabasePropertyRemoved(ByVal db As DAO.Database, ByVal propertyName As String, Optional ByVal sourceProcedure As String = "EnsureDatabasePropertyRemoved")
    'Removes one DAO database property when present, treats only property-not-found errors as already absent, and verifies deletion.
    On Error GoTo ErrHandler
    Dim propertyExists As Boolean
    Dim deleteErrNumber As Long
    Dim deleteErrSource As String
    Dim deleteErrDescription As String
    propertyName = Trim$(propertyName)
    sourceProcedure = Trim$(sourceProcedure)
    If Len(sourceProcedure) = 0 Then sourceProcedure = "EnsureDatabasePropertyRemoved"
    If db Is Nothing Then Err.Raise vbObjectError + 6894, sourceProcedure, "The DAO database property target is Nothing."
    If Len(propertyName) = 0 Then Err.Raise vbObjectError + 6895, sourceProcedure, "The DAO database property name to remove is blank."
    On Error Resume Next
    Err.Clear
    db.properties.Delete propertyName
    deleteErrNumber = Err.Number
    deleteErrSource = Err.Source
    deleteErrDescription = Err.Description
    Err.Clear
    On Error GoTo ErrHandler
    If deleteErrNumber <> 0 And deleteErrNumber <> 3265 And deleteErrNumber <> 3270 Then Err.Raise deleteErrNumber, deleteErrSource, deleteErrDescription
    db.properties.Refresh
    Call DatabaseTextPropertyValue(db, propertyName, propertyExists)
    If propertyExists Then Err.Raise vbObjectError + 6896, sourceProcedure, "The DAO database property remained present after deletion verification." & vbCrLf & "Property: " & propertyName
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, sourceProcedure, "Could not remove and verify the DAO database property." & vbCrLf & "Property: " & propertyName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Sub

' Purpose: Returns elapsed seconds from VBA Timer while handling a midnight rollover; an unset/nonpositive start timer returns zero.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   startTimer (Input; Single): Timing value used to calculate or report the start timer duration for this operation.
' Returns: Double: Returns elapsed seconds computed from the supplied timer values, including the routine's wraparound handling.
Public Function ElapsedSecondsFromTimer(ByVal startTimer As Single) As Double
    'Returns elapsed seconds from VBA Timer while handling a midnight rollover; an unset/nonpositive start timer returns zero.
    If startTimer <= 0 Then Exit Function
    ElapsedSecondsFromTimer = CDbl(Timer) - CDbl(startTimer)
    If ElapsedSecondsFromTimer < 0 Then ElapsedSecondsFromTimer = ElapsedSecondsFromTimer + 86400#
End Function


' ================================================================================================================================================================================
' COMMON HELPER FUNCTIONS - BEGIN
' ================================================================================================================================================================================
' Reusable parsers, normalizers, encoders, extractors, comparators, formatters, naming helpers, diagnostics, and DAO/Access utility functions centralized from ADC modules.

' Purpose: Resolves quoted text collection list from the supplied/current ADC state, using `SharedDao_SqlQuote` and `Next` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   values (Input; Collection): Caller-supplied Collection used to provide or receive the ordered set of items associated with this operation.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function QuotedTextCollectionList(ByVal values As Collection) As String
    'Builds a comma-delimited SQL-quoted list from the values in a Collection.
    Dim item As Variant
    Dim resultText As String
    For Each item In values
        PumpSetupResponsiveness
        If Len(resultText) > 0 Then resultText = resultText & ","
        resultText = resultText & SharedDao_SqlQuote(CStr(item))
    Next item
    QuotedTextCollectionList = resultText
End Function

' Purpose: Resolves part8 append list item from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   listText (Input; String): Input list text for part8 append list item; is tested by If Len(listText) = 0; is passed to `Part8AppendListItem`.
'   itemText (Input; String): Input item text for part8 append list item; is assigned to `Part8AppendListItem`; is passed to `Part8AppendListItem`.
' Returns: String: Returns the part 8 append list item text value. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function Part8AppendListItem(ByVal listText As String, ByVal itemText As String) As String
    'Appends one item to a comma-and-space-delimited list without adding a leading delimiter.
    If Len(listText) = 0 Then
        Part8AppendListItem = itemText
    Else
        Part8AppendListItem = listText & ", " & itemText
    End If
End Function

' Purpose: Resolves length prefixed variant from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   fieldValue (Input; Variant): Variant input representing field value; the routine accepts the caller-provided value in the form required by its existing validation and conversion logic.
' Returns: String: Returns the length prefixed variant text value. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function LengthPrefixedVariant(ByVal fieldValue As Variant) As String
    'Serializes a Variant as a null marker or a length-prefixed text value for stable comparison or hashing.
    Dim valueText As String
    If IsNull(fieldValue) Then
        LengthPrefixedVariant = "N;"
    Else
        valueText = CStr(fieldValue)
        LengthPrefixedVariant = "V" & CStr(Len(valueText)) & ":" & valueText & ";"
    End If
End Function

' Purpose: Appends line for existing text, new line.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   existingText (Input; String): Input existing text for append line; is assigned to `AppendLine`; is tested by ElseIf Len(existingText) = 0.
'   newLine (Input; String): Source-text line or line-derived value used while parsing, scanning, comparing, or documenting VBA text.
' Returns: String: Returns the original text with the requested line content appended using this helper's delimiter/newline and blank-value rules.
Public Function AppendLine(ByVal existingText As String, ByVal newLine As String) As String
    'Appends a nonblank line to existing text using CRLF and avoids adding an unnecessary leading line break.
    newLine = Trim$(newLine)
    If Len(newLine) = 0 Then
        AppendLine = existingText
    ElseIf Len(existingText) = 0 Then
        AppendLine = newLine
    Else
        AppendLine = existingText & vbCrLf & newLine
    End If
End Function

' Purpose: Returns the text inside the first balanced parenthesized expression, respecting quoted VBA strings.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   value (Input; String): Input value for text inside first parentheses; is passed to `InStr`; is passed to `MatchingCloseParenPosition`.
' Returns: String: Returns the resolved text inside first parentheses text in the canonical representation consumed by this module's caller; vbNullString represents no applicable value where the existing
'   contract permits it.
Public Function TextInsideFirstParentheses(ByVal value As String) As String
    'Returns the text inside the first balanced parenthesized expression, respecting quoted VBA strings.
    Dim openParenPos As Long
    Dim closeParenPos As Long
    openParenPos = InStr(1, value, "(", vbBinaryCompare)
    If openParenPos = 0 Then Exit Function
    closeParenPos = MatchingCloseParenPosition(value, openParenPos)
    If closeParenPos <= openParenPos Then Exit Function
    TextInsideFirstParentheses = Mid$(value, openParenPos + 1, closeParenPos - openParenPos - 1)
End Function

' Purpose: Splits comma-delimited text at top level only, preserving commas inside quoted strings or nested parentheses.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   value (Input; String): Input value for split top level comma text; is passed to `Len`; is passed to `ch`.
' Returns: Collection: Returns the collection of top-level comma-delimited items parsed from the supplied text while preserving commas that occur inside nested/quoted constructs.
Public Function SplitTopLevelCommaText(ByVal value As String) As Collection
    'Splits comma-delimited text at top level only, preserving commas inside quoted strings or nested parentheses.
    Dim result As Collection
    Dim index As Long
    Dim ch As String
    Dim tokenText As String
    Dim depth As Long
    Dim inString As Boolean
    ' INTENTIONAL ORDERED VIEW: result preserves the deterministic sequence required by SplitTopLevelCommaText; keyed-only or unordered storage would not preserve the consumer semantics.
    Set result = New Collection
    For index = 1 To Len(value)
        If (index And 255) = 0 Then PumpSetupResponsiveness
        ch = Mid$(value, index, 1)
        If ch = Chr$(34) Then
            tokenText = tokenText & ch
            If inString And index < Len(value) And Mid$(value, index + 1, 1) = Chr$(34) Then
                index = index + 1
                tokenText = tokenText & Chr$(34)
            Else
                inString = Not inString
            End If
        ElseIf Not inString And ch = "(" Then
            depth = depth + 1
            tokenText = tokenText & ch
        ElseIf Not inString And ch = ")" Then
            If depth > 0 Then depth = depth - 1
            tokenText = tokenText & ch
        ElseIf Not inString And depth = 0 And ch = "," Then
            result.Add Trim$(tokenText)
            tokenText = vbNullString
        Else
            tokenText = tokenText & ch
        End If
    Next index
    If Len(Trim$(tokenText)) > 0 Then result.Add Trim$(tokenText)
    Set SplitTopLevelCommaText = result
End Function

Private Function IsQuotedVbaStringLiteral(ByVal valueText As String) As Boolean
    'Returns True when text is enclosed by VBA double-quote delimiters.
    IsQuotedVbaStringLiteral = (Len(valueText) >= 2 And Left$(valueText, 1) = Chr$(34) And Right$(valueText, 1) = Chr$(34))
End Function

Private Function UnquoteVbaStringLiteral(ByVal valueText As String) As String
    'Removes VBA string-literal delimiters and converts doubled embedded quotes back to single quote characters.
    Dim innerText As String
    If IsQuotedVbaStringLiteral(valueText) = False Then
        UnquoteVbaStringLiteral = valueText
        Exit Function
    End If
    innerText = Mid$(valueText, 2, Len(valueText) - 2)
    UnquoteVbaStringLiteral = Replace(innerText, Chr$(34) & Chr$(34), Chr$(34))
End Function

Private Function IsHashDelimitedDateLiteral(ByVal valueText As String) As Boolean
    'Returns True when text is a valid VBA date literal enclosed in hash characters.
    If Len(valueText) < 3 Then Exit Function
    If Left$(valueText, 1) <> "#" Then Exit Function
    If Right$(valueText, 1) <> "#" Then Exit Function
    IsHashDelimitedDateLiteral = IsDate(Mid$(valueText, 2, Len(valueText) - 2))
End Function

' Purpose: Resolves empty to placeholder from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   valueText (Input; String): Input value text for empty to placeholder; is tested by If Len(Trim$(valueText)) = 0; is assigned to `EmptyToPlaceholder`.
'   placeholderText (Input; String): Input placeholder text for empty to placeholder; is assigned to `EmptyToPlaceholder`.
' Returns: String: Returns the empty to placeholder text value. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function EmptyToPlaceholder(ByVal valueText As String, ByVal placeholderText As String) As String
    'Returns a placeholder for blank or whitespace-only text and otherwise preserves the supplied value.
    If Len(Trim$(valueText)) = 0 Then
        EmptyToPlaceholder = placeholderText
    Else
        EmptyToPlaceholder = valueText
    End If
End Function

' Purpose: Resolves function return type from declaration from the supplied/current ADC state, using `MatchingCloseParenPosition` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   declarationText (Input; String): Input declaration text for function return type from declaration; is passed to `MatchingCloseParenPosition`; is passed to `afterParamsText`.
' Returns: String: Returns the function return type from declaration text value. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function FunctionReturnTypeFromDeclaration(ByVal declarationText As String) As String
    'Extracts a function return type from VBA declaration text and defaults to Variant when no As clause is present.
    Dim closeParenPos As Long
    Dim afterParamsText As String
    Dim asPos As Long
    closeParenPos = MatchingCloseParenPosition(declarationText, InStr(1, declarationText, "(", vbBinaryCompare))
    If closeParenPos > 0 Then
        afterParamsText = Trim$(Mid$(declarationText, closeParenPos + 1))
    Else
        afterParamsText = Trim$(declarationText)
    End If
    asPos = InStr(1, afterParamsText, " As ", vbTextCompare)
    If asPos > 0 Then
        FunctionReturnTypeFromDeclaration = Trim$(Mid$(afterParamsText, asPos + Len(" As ")))
    Else
        FunctionReturnTypeFromDeclaration = "Variant"
    End If
End Function

' Purpose: Returns True when a VBA argument declaration begins with Optional or ParamArray.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   argText (Input; String): Input arg text for procedure argument is optional; is assigned to `valueText`.
' Returns: Boolean: True when a VBA argument declaration begins with Optional or ParamArray. False otherwise.
Public Function ProcedureArgumentIsOptional(ByVal argText As String) As Boolean
    'Returns True when a VBA argument declaration begins with Optional or ParamArray.
    Dim valueText As String
    valueText = argText
    If (StrComp(Left$(valueText, Len("Optional ")), "Optional ", vbTextCompare) = 0) Then
        ProcedureArgumentIsOptional = True
    ElseIf (StrComp(Left$(valueText, Len("ParamArray ")), "ParamArray ", vbTextCompare) = 0) Then
        ProcedureArgumentIsOptional = True
    End If
End Function

' Purpose: Resolves parse procedure arguments from the supplied/current ADC state, using `End` and `If` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   argumentsText (Input; String): Input arguments text for parse procedure arguments; is tested by If Len(argumentsText) = 0; is passed to `Len`.
' Returns: Collection: Returns the ordered Collection assembled by the routine; callers receive the object reference and should treat its membership/order according to the routine contract.
Public Function ParseProcedureArguments(ByVal argumentsText As String) As Collection
    'Parses a VBA-style argument list into an ordered Collection while respecting strings and nested brackets.
    Dim args As Collection
    Dim index As Long
    Dim startPos As Long
    Dim depth As Long
    Dim inString As Boolean
    Dim ch As String
    Dim partText As String
    ' INTENTIONAL ORDERED VIEW: args preserves the deterministic sequence required by ParseProcedureArguments; keyed-only or unordered storage would not preserve the consumer semantics.
    Set args = New Collection
    argumentsText = Trim$(argumentsText)
    If Len(argumentsText) = 0 Then
        Set ParseProcedureArguments = args
        Exit Function
    End If
    startPos = 1
    For index = 1 To Len(argumentsText)
        If (index And 255) = 0 Then PumpSetupResponsiveness
        ch = Mid$(argumentsText, index, 1)
        If inString Then
            If ch = """" Then
                If index < Len(argumentsText) And Mid$(argumentsText, index + 1, 1) = """" Then
                    index = index + 1
                Else
                    inString = False
                End If
            End If
        Else
            Select Case ch
                Case """"
                    inString = True
                Case "(", "[", "{"
                    depth = depth + 1
                Case ")", "]", "}"
                    depth = depth - 1
                    If depth < 0 Then
                        Err.Raise vbObjectError + 5320, "ParseProcedureArguments", "TargetArguments has a closing bracket/parenthesis without a matching opening bracket/parenthesis." & vbCrLf & "Arguments: " & argumentsText & vbCrLf & "Character position: " & CStr(index)
                    End If
                Case ","
                    If depth = 0 Then
                        partText = Trim$(Mid$(argumentsText, startPos, index - startPos))
                        args.Add CoerceProcedureArgument(partText)
                        startPos = index + 1
                    End If
            End Select
        End If
    Next index
    If inString Then
        Err.Raise vbObjectError + 5321, "ParseProcedureArguments", "TargetArguments has an unterminated string literal." & vbCrLf & "Arguments: " & argumentsText
    End If
    If depth <> 0 Then
        Err.Raise vbObjectError + 5322, "ParseProcedureArguments", "TargetArguments has unbalanced parentheses/brackets/braces." & vbCrLf & "Arguments: " & argumentsText
    End If
    partText = Trim$(Mid$(argumentsText, startPos))
    If Len(partText) > 0 Or Right$(argumentsText, 1) = "," Then
        args.Add CoerceProcedureArgument(partText)
    End If
    Set ParseProcedureArguments = args
End Function

Private Function CoerceProcedureArgument(ByVal argumentText As String) As Variant
    'Converts one parsed VBA argument literal to its corresponding Variant value when its literal syntax is recognized.
    Dim valueText As String
    Dim dateText As String
    Dim numericValue As Double
    valueText = argumentText
    If Len(valueText) = 0 Then
        CoerceProcedureArgument = Empty
        Exit Function
    End If
    Select Case LCase$(valueText)
        Case "null"
            CoerceProcedureArgument = Null
            Exit Function
        Case "empty"
            CoerceProcedureArgument = Empty
            Exit Function
        Case "true"
            CoerceProcedureArgument = True
            Exit Function
        Case "false"
            CoerceProcedureArgument = False
            Exit Function
    End Select
    If Left$(valueText, 1) = """" Or Right$(valueText, 1) = """" Then
        If IsQuotedVbaStringLiteral(valueText) = False Then
            Err.Raise vbObjectError + 5323, "CoerceProcedureArgument", "Invalid VBA string literal in TargetArguments." & vbCrLf & "Argument: " & valueText
        End If
        CoerceProcedureArgument = UnquoteVbaStringLiteral(valueText)
        Exit Function
    End If
    If Left$(valueText, 1) = "#" Or Right$(valueText, 1) = "#" Then
        If IsHashDelimitedDateLiteral(valueText) = False Then
            Err.Raise vbObjectError + 5324, "CoerceProcedureArgument", "Invalid date literal in TargetArguments. Date literals must be wrapped in # characters." & vbCrLf & "Argument: " & valueText
        End If
        dateText = Mid$(valueText, 2, Len(valueText) - 2)
        If IsDate(dateText) = False Then
            Err.Raise vbObjectError + 5325, "CoerceProcedureArgument", "Invalid date value in TargetArguments." & vbCrLf & "Argument: " & valueText
        End If
        CoerceProcedureArgument = CDate(dateText)
        Exit Function
    End If
    If IsNumeric(valueText) Then
        numericValue = CDbl(valueText)
        If InStr(1, valueText, ".", vbBinaryCompare) = 0 And InStr(1, valueText, "E", vbTextCompare) = 0 And InStr(1, valueText, "D", vbTextCompare) = 0 Then
            If numericValue >= -2147483648# And numericValue <= 2147483647# Then
                CoerceProcedureArgument = CLng(valueText)
            Else
                CoerceProcedureArgument = CDbl(valueText)
            End If
        Else
            CoerceProcedureArgument = CDbl(valueText)
        End If
        Exit Function
    End If
    'Unquoted values are treated as literal text. No Eval call is used here.
    CoerceProcedureArgument = valueText
End Function

' Purpose: Returns True when VBA argument-signature text contains a ParamArray declaration.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   argumentSignature (Input; String): Receives or supplies the discovered argument signature for the target procedure, including its declared parameter list.
' Returns: Boolean: True when VBA argument-signature text contains a ParamArray declaration. False otherwise.
Public Function ProcedureSignatureHasParamArray(ByVal argumentSignature As String) As Boolean
    'Returns True when VBA argument-signature text contains a ParamArray declaration.
    Dim args As Collection
    Dim item As Variant
    Dim argText As String
    Set args = SplitTopLevelCommaText(argumentSignature)
    For Each item In args
        PumpSetupResponsiveness
        argText = CStr(item)
        If (StrComp(Left$(argText, Len("ParamArray ")), "ParamArray ", vbTextCompare) = 0) Then
            ProcedureSignatureHasParamArray = True
            Exit Function
        End If
    Next item
End Function

' Purpose: Extracts the callable procedure name from a RunCode-style expression.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   runCodeExpression (Input; String): Expression text representing the run code calculation, RunCode target, SQL fragment, or UI binding interpreted by this routine.
' Returns: String: Returns the exact Access/VBA expression text for procedure name from run code expression, including the quoting/callback syntax required by its generated UI or query consumer.
Public Function ProcedureNameFromRunCodeExpression(ByVal runCodeExpression As String) As String
    'Extracts the callable procedure name from a RunCode-style expression.
    'Intended use: Returns the Procedure Name From Run Code Expression result for direct, query, form, or required cross-module callers.
    Dim valueText As String
    Dim openParenPos As Long
    Dim firstSpacePos As Long
    Dim tokenText As String
    valueText = Trim$(runCodeExpression)
    If Len(valueText) = 0 Then Exit Function
    If Left$(valueText, 1) = "=" Then valueText = Trim$(Mid$(valueText, 2))
    openParenPos = InStr(1, valueText, "(", vbBinaryCompare)
    firstSpacePos = InStr(1, valueText, " ", vbBinaryCompare)
    If openParenPos > 0 Then
        tokenText = Trim$(Left$(valueText, openParenPos - 1))
    ElseIf firstSpacePos > 0 Then
        tokenText = Trim$(Left$(valueText, firstSpacePos - 1))
    Else
        tokenText = valueText
    End If
    ProcedureNameFromRunCodeExpression = tokenText
End Function

Private Function FunctionArgumentsTextFromCall(ByVal callText As String) As String
    'Extracts the balanced argument text from a function-call expression.
    Dim openParenPos As Long
    Dim closeParenPos As Long
    openParenPos = InStr(1, callText, "(", vbBinaryCompare)
    If openParenPos = 0 Then Exit Function
    closeParenPos = MatchingCloseParenPosition(callText, openParenPos)
    If closeParenPos = 0 Or closeParenPos <= openParenPos Then Exit Function
    FunctionArgumentsTextFromCall = Mid$(callText, openParenPos + 1, closeParenPos - openParenPos - 1)
End Function

' Purpose: Resolves first string argument from function call from the supplied/current ADC state, using `FunctionArgumentsTextFromCall` and `If` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   callText (Input; String): Input call text for first string argument from function call; is passed to `FunctionArgumentsTextFromCall`.
' Returns: String: Returns the first string argument from function call text value. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or
'   not resolved.
Public Function FirstStringArgumentFromFunctionCall(ByVal callText As String) As String
    'Returns the first argument of a function call when that argument is a quoted VBA string literal.
    Dim argsText As String
    Dim index As Long
    Dim ch As String
    Dim resultText As String
    Dim inString As Boolean
    argsText = FunctionArgumentsTextFromCall(callText)
    If Len(argsText) = 0 Then Exit Function
    For index = 1 To Len(argsText)
        If (index And 255) = 0 Then PumpSetupResponsiveness
        ch = Mid$(argsText, index, 1)
        If Not inString Then
            If ch = """" Then
                inString = True
            ElseIf ch <> " " And ch <> vbTab Then
                Exit Function
            End If
        Else
            If ch = """" Then
                If index < Len(argsText) And Mid$(argsText, index + 1, 1) = """" Then
                    resultText = resultText & """"
                    index = index + 1
                Else
                    FirstStringArgumentFromFunctionCall = resultText
                    Exit Function
                End If
            Else
                resultText = resultText & ch
            End If
        End If
    Next index
End Function

' Purpose: Resolves public callable procedure name from declaration from the supplied/current ADC state, using `Shared_String_CompactSpaces` and `StrComp` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   declarationText (Input; String): Input declaration text for public callable procedure name from declaration; is passed to `Shared_String_CompactSpaces`.
' Returns: String: Returns the resolved name produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function PublicCallableProcedureNameFromDeclaration(ByVal declarationText As String) As String
    'Extracts the procedure name from a public callable VBA declaration after removing trailing comments and compacting spacing.
    Dim lineText As String
    lineText = Shared_String_CompactSpaces(CodeLineWithoutTrailingComment(declarationText))
    If Len(lineText) = 0 Then Exit Function
    If StrComp(Left$(lineText, Len("Private ")), "Private ", vbTextCompare) = 0 Then Exit Function
    If StrComp(Left$(lineText, Len("Friend ")), "Friend ", vbTextCompare) = 0 Then Exit Function
    If StrComp(Left$(lineText, Len("Public ")), "Public ", vbTextCompare) = 0 Then lineText = Trim$(Mid$(lineText, Len("Public ") + 1))
    If StrComp(Left$(lineText, Len("Function ")), "Function ", vbTextCompare) = 0 Then
        PublicCallableProcedureNameFromDeclaration = ProcedureNameFromDeclarationRemainder(Mid$(lineText, Len("Function ") + 1))
    ElseIf StrComp(Left$(lineText, Len("Sub ")), "Sub ", vbTextCompare) = 0 Then
        PublicCallableProcedureNameFromDeclaration = ProcedureNameFromDeclarationRemainder(Mid$(lineText, Len("Sub ") + 1))
    End If
End Function

' Purpose: Returns True when a VBA code line can declare the requested procedure name at a valid identifier boundary.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   lineText (Input; String): Input line text for code line may declare procedure; is passed to `Shared_String_CompactSpaces`.
'   ProcedureName (Input; String): Identifies the VBA procedure being searched, validated, described, or invoked by its procedure name.
' Returns: Boolean: True when a VBA code line can declare the requested procedure name at a valid identifier boundary. False otherwise.
Public Function CodeLineMayDeclareProcedure(ByVal lineText As String, ByVal ProcedureName As String) As Boolean
    'Returns True when a VBA code line can declare the requested procedure name at a valid identifier boundary.
    Dim declarationLine As String
    Dim lowerLine As String
    Dim lowerName As String
    Dim publicFunctionPrefix As String
    Dim publicSubPrefix As String
    Dim functionPrefix As String
    Dim subPrefix As String
    declarationLine = Shared_String_CompactSpaces(CodeLineWithoutTrailingComment(lineText))
    If Len(declarationLine) = 0 Then Exit Function
    If Left$(declarationLine, 1) = "'" Then Exit Function
    If (StrComp(Left$(declarationLine, Len("Attribute ")), "Attribute ", vbTextCompare) = 0) Then Exit Function
    If (StrComp(Left$(declarationLine, Len("Private ")), "Private ", vbTextCompare) = 0) Then Exit Function
    If (StrComp(Left$(declarationLine, Len("Friend ")), "Friend ", vbTextCompare) = 0) Then Exit Function
    lowerLine = LCase$(declarationLine)
    lowerName = LCase$(ProcedureName)
    publicFunctionPrefix = LCase$("Public Function " & lowerName)
    publicSubPrefix = LCase$("Public Sub " & lowerName)
    functionPrefix = LCase$("Function " & lowerName)
    subPrefix = LCase$("Sub " & lowerName)
    If Left$(lowerLine, Len(publicFunctionPrefix)) = publicFunctionPrefix Then
        CodeLineMayDeclareProcedure = ProcedureNameBoundaryOk(declarationLine, Len(publicFunctionPrefix) + 1)
    ElseIf Left$(lowerLine, Len(publicSubPrefix)) = publicSubPrefix Then
        CodeLineMayDeclareProcedure = ProcedureNameBoundaryOk(declarationLine, Len(publicSubPrefix) + 1)
    ElseIf Left$(lowerLine, Len(functionPrefix)) = functionPrefix Then
        CodeLineMayDeclareProcedure = ProcedureNameBoundaryOk(declarationLine, Len(functionPrefix) + 1)
    ElseIf Left$(lowerLine, Len(subPrefix)) = subPrefix Then
        CodeLineMayDeclareProcedure = ProcedureNameBoundaryOk(declarationLine, Len(subPrefix) + 1)
    End If
End Function

Private Function ProcedureNameBoundaryOk(ByVal lineText As String, ByVal boundaryPosition As Long) As Boolean
    'Returns True when the character following a candidate procedure name is a valid declaration boundary.
    Dim ch As String
    If boundaryPosition > Len(lineText) Then
        ProcedureNameBoundaryOk = True
        Exit Function
    End If
    ch = Mid$(lineText, boundaryPosition, 1)
    ProcedureNameBoundaryOk = (ch = " " Or ch = "(")
End Function

' Purpose: Resolves procedure name from declaration remainder from the supplied/current ADC state, using `ElseIf` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   remainderText (Input; String): Input remainder text for procedure name from declaration remainder; is passed to `valueText`.
' Returns: String: Returns the resolved name produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function ProcedureNameFromDeclarationRemainder(ByVal remainderText As String) As String
    'Extracts the leading procedure identifier from declaration text that follows the procedure keyword.
    Dim valueText As String
    Dim openParenPos As Long
    Dim firstSpacePos As Long
    valueText = Trim$(remainderText)
    openParenPos = InStr(1, valueText, "(", vbBinaryCompare)
    firstSpacePos = InStr(1, valueText, " ", vbBinaryCompare)
    If openParenPos > 0 Then
        ProcedureNameFromDeclarationRemainder = Trim$(Left$(valueText, openParenPos - 1))
    ElseIf firstSpacePos > 0 Then
        ProcedureNameFromDeclarationRemainder = Trim$(Left$(valueText, firstSpacePos - 1))
    Else
        ProcedureNameFromDeclarationRemainder = valueText
    End If
End Function

' Purpose: Returns the first nonblank array-line index within the requested scan range, or -1 when none exists.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   codeLines (Input; Variant): Source-text line or line-derived value used while parsing, scanning, comparing, or documenting VBA text.
'   maxScanIndex (Input; Long): Supplies the max scan index used to bound, report, or validate the number of items processed by this operation.
' Returns: Long: Returns the first non blank line index line/item index used by the caller; the function's existing sentinel value indicates that no qualifying position was found.
Public Function FirstNonBlankLineIndex(ByVal codeLines As Variant, ByVal maxScanIndex As Long) As Long
    'Returns the first nonblank array-line index within the requested scan range, or -1 when none exists.
    Dim i As Long
    FirstNonBlankLineIndex = -1
    For i = LBound(codeLines) To maxScanIndex
        PumpSetupResponsiveness
        If Len(Trim$(CStr(codeLines(i)))) > 0 Then
            FirstNonBlankLineIndex = i
            Exit Function
        End If
    Next i
End Function

' Purpose: Finds the closing separator line of a managed module-header block within the supplied line range.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   codeLines (Input; Variant): Source-text line or line-derived value used while parsing, scanning, comparing, or documenting VBA text.
'   headerStartIndex (Input; Long): Index selecting the header start index item or position within the caller-supplied definition/source sequence.
'   maxScanIndex (Input; Long): Supplies the max scan index used to bound, report, or validate the number of items processed by this operation.
' Returns: Long: Returns the managed header end index line/item index used by the caller; the function's existing sentinel value indicates that no qualifying position was found.
Public Function ManagedHeaderEndIndex(ByVal codeLines As Variant, ByVal headerStartIndex As Long, ByVal maxScanIndex As Long) As Long
    'Finds the closing separator line of a managed module-header block within the supplied line range.
    Dim i As Long
    Dim lineText As String
    ManagedHeaderEndIndex = -1
    For i = headerStartIndex + 1 To maxScanIndex
        PumpSetupResponsiveness
        lineText = Trim$(CStr(codeLines(i)))
        If IsHeaderSeparatorLine(lineText) Then
            ManagedHeaderEndIndex = i
            Exit Function
        End If
    Next i
End Function

' Purpose: Extracts the component title from the recognized component-header line in a managed module header.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   codeLines (Input; Variant): Source-text line or line-derived value used while parsing, scanning, comparing, or documenting VBA text.
'   headerStartIndex (Input; Long): Index selecting the header start index item or position within the caller-supplied definition/source sequence.
'   maxScanIndex (Input; Long): Supplies the max scan index used to bound, report, or validate the number of items processed by this operation.
' Returns: String: Returns the canonical display text for header title from lines after applying the title/caption/label rules owned by this module; vbNullString represents no applicable display text where
'   supported.
Public Function HeaderTitleFromLines(ByVal codeLines As Variant, ByVal headerStartIndex As Long, ByVal maxScanIndex As Long) As String
    'Extracts the component title from the recognized component-header line in a managed module header.
    Dim i As Long
    Dim lineText As String
    For i = headerStartIndex To maxScanIndex
        PumpSetupResponsiveness
        lineText = Trim$(CStr(codeLines(i)))
        If IsComponentHeaderLine(lineText) Then
            HeaderTitleFromLines = Trim$(SharedText_AfterFirstColon(lineText))
            Exit Function
        End If
        If i > headerStartIndex And IsHeaderSeparatorLine(lineText) Then Exit Function
    Next i
End Function

' Purpose: Extracts a managed-header value and any comment-line continuations associated with a requested metadata prefix.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   codeLines (Input; Variant): Source-text line or line-derived value used while parsing, scanning, comparing, or documenting VBA text.
'   headerStartIndex (Input; Long): Index selecting the header start index item or position within the caller-supplied definition/source sequence.
'   headerEndIndex (Input; Long): Index selecting the header end index item or position within the caller-supplied definition/source sequence.
'   prefixText (Input; String): Input prefix text for header value by prefix; is tested by If (StrComp(Left$(lineText, Len(prefixText)), prefixText, vbTextCompare) = 0); is passed to `Len`.
' Returns: String: Returns the resolved header value by prefix text in the canonical representation consumed by this module's caller; vbNullString represents no applicable value where the existing contract
'   permits it.
Public Function HeaderValueByPrefix(ByVal codeLines As Variant, ByVal headerStartIndex As Long, ByVal headerEndIndex As Long, ByVal prefixText As String) As String
    'Extracts a managed-header value and any comment-line continuations associated with a requested metadata prefix.
    Dim i As Long
    Dim j As Long
    Dim lineText As String
    Dim nextLineText As String
    Dim valueText As String
    Dim continuationText As String
    If headerEndIndex < headerStartIndex Then Exit Function
    For i = headerStartIndex To headerEndIndex
        PumpSetupResponsiveness
        lineText = Trim$(CStr(codeLines(i)))
        If (StrComp(Left$(lineText, Len(prefixText)), prefixText, vbTextCompare) = 0) Then
            valueText = Trim$(Mid$(lineText, Len(prefixText) + 1))
            For j = i + 1 To headerEndIndex
                PumpSetupResponsiveness
                nextLineText = Trim$(CStr(codeLines(j)))
                If Len(nextLineText) = 0 Then
                    'Ignore blank lines inside a metadata block.
                ElseIf IsHeaderSeparatorLine(nextLineText) Then
                    Exit For
                ElseIf IsKnownHeaderPrefixLine(nextLineText) Then
                    Exit For
                ElseIf Left$(nextLineText, 1) = "'" Then
                    continuationText = HeaderCommentPayload(nextLineText)
                    If Len(continuationText) > 0 Then
                        If Len(valueText) > 0 Then valueText = valueText & vbCrLf
                        valueText = valueText & continuationText
                    End If
                Else
                    Exit For
                End If
            Next j
            HeaderValueByPrefix = Trim$(valueText)
            Exit Function
        End If
    Next i
End Function

Private Function HeaderCommentPayload(ByVal commentLineText As String) As String
    'Removes the leading apostrophe from a VBA comment line and returns its trimmed payload.
    Dim s As String
    s = commentLineText
    If Left$(s, 1) = "'" Then s = Mid$(s, 2)
    HeaderCommentPayload = Trim$(s)
End Function

' Purpose: Returns True when a line exactly matches the managed-header separator syntax.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   lineText (Input; String): Input line text for is header separator line; is passed to `StrComp`.
' Returns: Boolean: True when header separator line is true for the supplied/current context; False otherwise.
Public Function IsHeaderSeparatorLine(ByVal lineText As String) As Boolean
    'Returns True when a line exactly matches the managed-header separator syntax.
    IsHeaderSeparatorLine = (StrComp(lineText, "'====================================================================", vbBinaryCompare) = 0)
End Function

' Purpose: Returns True when a line begins with one of the recognized VBA component-header labels.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   lineText (Input; String): Input line text for is component header line; is assigned to `s`.
' Returns: Boolean: True when component header line is true for the supplied/current context; False otherwise.
Public Function IsComponentHeaderLine(ByVal lineText As String) As Boolean
    'Returns True when a line begins with one of the recognized VBA component-header labels.
    Dim s As String
    Dim prefixes As Variant
    Dim i As Long
    s = lineText
    prefixes = Array("' STANDARD MODULE:", "' CLASS MODULE:", "' FORM MODULE:", "' REPORT MODULE:", "' DOCUMENT MODULE:")
    For i = LBound(prefixes) To UBound(prefixes)
        PumpSetupResponsiveness
        If StrComp(Left$(s, Len(CStr(prefixes(i)))), CStr(prefixes(i)), vbTextCompare) = 0 Then
            IsComponentHeaderLine = True
            Exit Function
        End If
    Next i
End Function

Private Function IsKnownHeaderPrefixLine(ByVal lineText As String) As Boolean
    'Returns True when a line begins with a recognized managed-header metadata prefix.
    Dim s As String
    Dim prefixes As Variant
    Dim i As Long
    s = lineText
    If IsComponentHeaderLine(s) Then
        IsKnownHeaderPrefixLine = True
        Exit Function
    End If
    prefixes = Array("' VB_Name:", "' SHARED CODE:", "' SHARED FORM CODE:", "' SHARED TREEVIEW CODE:", "' SHARED TVW Code:", "' CLASSES USED:", "' DYNAMIC STRING REFERENCES:", "' Purpose:", "' Author:", "' Support:", "' Version:", "' Notes:", "' Core Rules Enforced Here:", "' Design Map:", "' Configuration:", "' Todo:", "' Module Level Bugs:", "' Public Subs:", "' Public Functions:", "' Public Properties:", "' Entry Points:")
    For i = LBound(prefixes) To UBound(prefixes)
        PumpSetupResponsiveness
        If StrComp(Left$(s, Len(CStr(prefixes(i)))), CStr(prefixes(i)), vbTextCompare) = 0 Then
            IsKnownHeaderPrefixLine = True
            Exit Function
        End If
    Next i
End Function

' Purpose: Resolves first identifier token from the supplied/current ADC state, using `If` and `SharedVba_IsIdentifierChar` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   valueText (Input; String): Input value text for first identifier token; is passed to `Len`; is passed to `ch`.
' Returns: String: Returns the proof/state token produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function FirstIdentifierToken(ByVal valueText As String) As String
    'Extracts the first VBA identifier token found in the supplied text.
    Dim i As Long
    Dim ch As String
    Dim tokenText As String
    Dim started As Boolean
    For i = 1 To Len(valueText)
        If (i And 255) = 0 Then PumpSetupResponsiveness
        ch = Mid$(valueText, i, 1)
        If started Then
            If SharedVba_IsIdentifierChar(ch) Then
                tokenText = tokenText & ch
            Else
                Exit For
            End If
        Else
            If IsIdentifierStartCharacter(ch) Then
                started = True
                tokenText = ch
            End If
        End If
    Next i
    FirstIdentifierToken = tokenText
End Function

' Purpose: Returns True when a character is valid as the first character of the identifier syntax used by the parser.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   ch (Input; String): Single source character being classified by the VBA source scanner.
' Returns: Boolean: True when identifier start character is true for the supplied/current context; False otherwise.
Public Function IsIdentifierStartCharacter(ByVal ch As String) As Boolean
    'Returns True when a character is valid as the first character of the identifier syntax used by the parser.
    If Len(ch) = 0 Then Exit Function
    IsIdentifierStartCharacter = ((ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or ch = "_")
End Function

Private Function MatchingCloseParenPosition(ByVal value As String, ByVal openParenPos As Long) As Long
    'Finds the closing parenthesis that balances a supplied opening parenthesis while respecting quoted strings.
    Dim index As Long
    Dim ch As String
    Dim depth As Long
    Dim inString As Boolean
    If openParenPos <= 0 Then Exit Function
    For index = openParenPos To Len(value)
        If (index And 255) = 0 Then PumpSetupResponsiveness
        ch = Mid$(value, index, 1)
        If ch = Chr$(34) Then
            If inString And index < Len(value) And Mid$(value, index + 1, 1) = Chr$(34) Then
                index = index + 1
            Else
                inString = Not inString
            End If
        ElseIf Not inString Then
            If ch = "(" Then
                depth = depth + 1
            ElseIf ch = ")" Then
                depth = depth - 1
                If depth = 0 Then
                    MatchingCloseParenPosition = index
                    Exit Function
                End If
            End If
        End If
    Next index
End Function

' Purpose: Returns True when a VBA source line ends with a continuation underscore after right-side whitespace is ignored.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   lineText (Input; String): Input line text for VBA line continues; is passed to `valueText`.
' Returns: Boolean: True when a VBA source line ends with a continuation underscore after right-side whitespace is ignored. False otherwise.
Public Function VbaLineContinues(ByVal lineText As String) As Boolean
    'Returns True when a VBA source line ends with a continuation underscore after right-side whitespace is ignored.
    Dim valueText As String
    valueText = RTrim$(lineText)
    VbaLineContinues = (Len(valueText) > 0 And Right$(valueText, 1) = "_")
End Function

' Purpose: Removes a trailing VBA continuation underscore and surrounding right-side whitespace.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   lineText (Input; String): Input line text for remove VBA line continuation; is passed to `valueText`.
' Returns: String: Returns the supplied source text with VBA line continuation removed according to this parser helper's rules.
Public Function RemoveVbaLineContinuation(ByVal lineText As String) As String
    'Removes a trailing VBA continuation underscore and surrounding right-side whitespace.
    Dim valueText As String
    valueText = RTrim$(lineText)
    If Len(valueText) > 0 And Right$(valueText, 1) = "_" Then
        valueText = Left$(valueText, Len(valueText) - 1)
    End If
    RemoveVbaLineContinuation = RTrim$(valueText)
End Function

' Purpose: Resolves code line without trailing comment from the supplied/current ADC state, using `If` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   lineText (Input; String): Input line text for code line without trailing comment; is passed to `Len`; is passed to `ch`.
' Returns: String: Returns the code line without trailing comment text value. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function CodeLineWithoutTrailingComment(ByVal lineText As String) As String
    'Removes a trailing apostrophe comment from a VBA code line while preserving apostrophes inside quoted strings.
    Dim index As Long
    Dim ch As String
    Dim inString As Boolean
    For index = 1 To Len(lineText)
        If (index And 255) = 0 Then PumpSetupResponsiveness
        ch = Mid$(lineText, index, 1)
        If ch = Chr$(34) Then
            If inString And index < Len(lineText) And Mid$(lineText, index + 1, 1) = Chr$(34) Then
                index = index + 1
            Else
                inString = Not inString
            End If
        ElseIf ch = "'" And Not inString Then
            CodeLineWithoutTrailingComment = RTrim$(Left$(lineText, index - 1))
            Exit Function
        End If
    Next index
    CodeLineWithoutTrailingComment = RTrim$(lineText)
End Function

' Purpose: Resolves and returns appended note from the supplied inputs and the authoritative ADC definitions/state used by this module.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   existingNotes (Input; String): Existing notes text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   newNoteText (Input; String): New note text text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
' Returns: String: Returns the original text with the requested ed note content appended using this helper's delimiter/newline and blank-value rules.
Public Function AppendedNote(ByVal existingNotes As String, ByVal newNoteText As String) As String
    'Appends a timestamped note to existing note text using CRLF when prior notes are present.
    Dim noteText As String
    noteText = Format$(Now, "yyyy-mm-dd hh:nn:ss") & " - " & newNoteText
    If Len(Trim$(existingNotes)) > 0 Then
        AppendedNote = existingNotes & vbCrLf & noteText
    Else
        AppendedNote = noteText
    End If
End Function

' Purpose: Computes the existing deterministic hexadecimal hash used for stable text fingerprints.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   value (Input; String): Input value examined or transformed by this procedure according to its documented stable text hash hex contract.
' Returns: String: Returns the deterministic stable text hash hex token/signature produced from the supplied state, allowing callers to compare exact state without repeating the underlying scan/build work.
Public Function StableTextHashHex(ByVal value As String) As String
    'Computes the existing deterministic hexadecimal hash used for stable text fingerprints.
    Dim byteIndex As Long
    Dim valueBytes() As Byte
    Dim hashValue As Double
    Dim chValue As Long
    hashValue = 5381#
    If Len(value) > 0 Then
        valueBytes = value
        For byteIndex = 0 To UBound(valueBytes) Step 2
            If byteIndex > 0 Then
                If (byteIndex And 8191) = 0 Then PumpSetupResponsiveness
            End If
            chValue = CLng(valueBytes(byteIndex)) + CLng(valueBytes(byteIndex + 1)) * 256&
            hashValue = hashValue * 33# + chValue
            hashValue = hashValue - Int(hashValue / 2147483647#) * 2147483647#
        Next byteIndex
    End If
    StableTextHashHex = Right$("00000000" & Hex$(CLng(hashValue)), 8)
End Function

' Purpose: Appends unique pipe delimited names for existing names, additional names.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   existingNames (Input; String): Delimited list/set of names already collected; used as the base when merging additional unique names.
'   additionalNames (Input; String): Delimited list/set of candidate names to merge into the existing name set without duplicating entries.
' Returns: String: Returns the original text with the requested unique pipe delimited names content appended using this helper's delimiter/newline and blank-value rules.
Public Function AppendUniquePipeDelimitedNames(ByVal existingNames As String, ByVal additionalNames As String) As String
    'Combines pipe-delimited names while trimming entries, removing blanks, and preserving only the first case-insensitive occurrence.
    Dim seen As Object
    Dim item As Variant
    Dim normalizedName As String
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    For Each item In Split(existingNames & IIf(Len(existingNames) > 0 And Len(additionalNames) > 0, "|", vbNullString) & additionalNames, "|")
        normalizedName = Trim$(CStr(item))
        If Len(normalizedName) > 0 Then
            If seen.Exists(normalizedName) = False Then
                seen.Add normalizedName, True
                AppendUniquePipeDelimitedNames = AppendUniquePipeDelimitedNames & IIf(Len(AppendUniquePipeDelimitedNames) = 0, vbNullString, "|") & normalizedName
            End If
        End If
    Next item
End Function

' Purpose: Resolves pipe delimited collection names from the supplied/current ADC state, using `AppendUniquePipeDelimitedNames` and `Next` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   names (Input; Collection): Caller-supplied Collection used to provide or receive the ordered set of items associated with this operation.
' Returns: String: Returns the resolved name produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function PipeDelimitedCollectionNames(ByVal names As Collection) As String
    'Converts a Collection of names to a unique pipe-delimited string using the shared name-appending rules.
    Dim item As Variant
    If names Is Nothing Then Exit Function
    For Each item In names
        PipeDelimitedCollectionNames = AppendUniquePipeDelimitedNames(PipeDelimitedCollectionNames, CStr(item))
    Next item
End Function

' Purpose: Returns True for ASCII VBA identifier characters and for non-ASCII source characters accepted by the dependency scanner.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   characterText (Input; String): Input character text for VBA source identifier character; is tested by If Len(characterText) <> 1; is tested by If SharedVba_IsIdentifierChar(characterText).
' Returns: Boolean: True when the supplied character is valid in the VBA identifier position tested by the scanner; False when it is not.
Public Function VbaSourceIdentifierCharacter(ByVal characterText As String) As Boolean
    'Returns True for ASCII VBA identifier characters and for non-ASCII source characters accepted by the dependency scanner.
    Dim characterCode As Long
    If Len(characterText) <> 1 Then Exit Function
    If SharedVba_IsIdentifierChar(characterText) Then VbaSourceIdentifierCharacter = True: Exit Function
    characterCode = AscW(characterText)
    If characterCode < 0 Then characterCode = characterCode + 65536
    VbaSourceIdentifierCharacter = (characterCode > 127)
End Function

' Purpose: Resolves leading identifier from the supplied/current ADC state, using `If` and `Next` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   valueText (Input; String): Input value text for leading identifier; is passed to `Len`; is passed to `currentCharacter`.
' Returns: String: Returns the leading identifier text value. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function LeadingIdentifier(ByVal valueText As String) As String
    'Returns the leading ASCII identifier token from trimmed text.
    Dim characterIndex As Long
    Dim currentCharacter As String
    valueText = Trim$(valueText)
    For characterIndex = 1 To Len(valueText)
        If (characterIndex And 255) = 0 Then PumpSetupResponsiveness
        currentCharacter = Mid$(valueText, characterIndex, 1)
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_", currentCharacter, vbBinaryCompare) = 0 Then Exit For
        LeadingIdentifier = LeadingIdentifier & currentCharacter
    Next characterIndex
End Function

' Purpose: Returns True when Rem begins a VBA comment at the supplied character position rather than appearing inside an identifier.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   sourceLine (Input; String): Source-text line or line-derived value used while parsing, scanning, comparing, or documenting VBA text.
'   characterIndex (Input; Long): Index selecting the character index item or position within the caller-supplied definition/source sequence.
' Returns: Boolean: True when Rem begins a VBA comment at the supplied character position rather than appearing inside an identifier. False otherwise.
Public Function VbaRemCommentStartsAt(ByVal sourceLine As String, ByVal characterIndex As Long) As Boolean
    'Returns True when Rem begins a VBA comment at the supplied character position rather than appearing inside an identifier.
    Dim precedingCharacter As String
    Dim followingCharacter As String
    If LCase$(Mid$(sourceLine, characterIndex, 3)) <> "rem" Then Exit Function
    If characterIndex > 1 Then
        precedingCharacter = Mid$(sourceLine, characterIndex - 1, 1)
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_", precedingCharacter, vbBinaryCompare) > 0 Then Exit Function
    End If
    If characterIndex + 3 <= Len(sourceLine) Then
        followingCharacter = Mid$(sourceLine, characterIndex + 3, 1)
        If followingCharacter <> " " And followingCharacter <> vbTab Then Exit Function
    End If
    VbaRemCommentStartsAt = True
End Function

' Purpose: Removes apostrophe or Rem comments from a VBA source line while preserving comment-like text inside quoted strings.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   sourceLine (Input; String): Source-text line or line-derived value used while parsing, scanning, comparing, or documenting VBA text.
' Returns: String: Returns the supplied source text with VBA comment removed according to this parser helper's rules.
Public Function StripVbaComment(ByVal sourceLine As String) As String
    'Removes apostrophe or Rem comments from a VBA source line while preserving comment-like text inside quoted strings.
    Dim characterIndex As Long
    Dim currentCharacter As String
    Dim insideString As Boolean
    For characterIndex = 1 To Len(sourceLine)
        If (characterIndex And 255) = 0 Then PumpSetupResponsiveness
        currentCharacter = Mid$(sourceLine, characterIndex, 1)
        If currentCharacter = Chr$(34) Then
            If insideString And characterIndex < Len(sourceLine) And Mid$(sourceLine, characterIndex + 1, 1) = Chr$(34) Then
                StripVbaComment = StripVbaComment & currentCharacter & currentCharacter
                characterIndex = characterIndex + 1
            Else
                insideString = Not insideString
                StripVbaComment = StripVbaComment & currentCharacter
            End If
        ElseIf insideString = False And VbaRemCommentStartsAt(sourceLine, characterIndex) Then
            Exit Function
        ElseIf currentCharacter = "'" And insideString = False Then
            Exit Function
        Else
            StripVbaComment = StripVbaComment & currentCharacter
        End If
    Next characterIndex
End Function

' Purpose: Appends delimited value for existing text, new value.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   existingText (Input; String): Input existing text for append delimited value; is tested by If Len(existingText) = 0; is passed to `AppendDelimitedValue`.
'   newValue (Input; String): Replacement text to write/apply when the target differs from the required value.
' Returns: String: Returns the original text with the requested delimited value content appended using this helper's delimiter/newline and blank-value rules.
Public Function AppendDelimitedValue(ByVal existingText As String, ByVal newValue As String) As String
    'Appends one value to existing text using a semicolon-and-space delimiter only when needed.
    If Len(existingText) = 0 Then
        AppendDelimitedValue = newValue
    Else
        AppendDelimitedValue = existingText & "; " & newValue
    End If
End Function

' Purpose: Returns a right-trimmed description prefix that fits the byte budget and prefers a nearby sentence or word boundary.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   authoritativeDescription (Input; String): Authoritative description text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   byteBudget (Input; Long): Maximum encoded byte budget available for the generated text/value; the routine truncates or rejects content according to its existing boundary behavior.
' Returns: String: Returns the canonical description text for managed description prefix within byte budget, suitable for the managed catalog/field/object metadata destination; vbNullString indicates that no
'   description applies where allowed.
Public Function ManagedDescriptionPrefixWithinByteBudget(ByVal authoritativeDescription As String, ByVal byteBudget As Long) As String
    'Returns a right-trimmed description prefix that fits the byte budget and prefers a nearby sentence or word boundary.
    Dim candidateText As String
    Dim maximumCharacters As Long
    Dim minimumBoundary As Long
    Dim boundaryPosition As Long
    Dim candidatePosition As Long
    If byteBudget <= 0 Then Exit Function
    maximumCharacters = byteBudget \ 2
    If maximumCharacters <= 0 Then Exit Function
    candidateText = Left$(authoritativeDescription, maximumCharacters)
    Do While LenB(candidateText) > byteBudget And Len(candidateText) > 0
        candidateText = Left$(candidateText, Len(candidateText) - 1)
    Loop
    minimumBoundary = (Len(candidateText) * 3) \ 4
    candidatePosition = InStrRev(candidateText, vbCrLf)
    If candidatePosition > boundaryPosition Then boundaryPosition = candidatePosition
    candidatePosition = InStrRev(candidateText, ". ")
    If candidatePosition > boundaryPosition Then boundaryPosition = candidatePosition + 1
    candidatePosition = InStrRev(candidateText, "! ")
    If candidatePosition > boundaryPosition Then boundaryPosition = candidatePosition + 1
    candidatePosition = InStrRev(candidateText, "? ")
    If candidatePosition > boundaryPosition Then boundaryPosition = candidatePosition + 1
    If boundaryPosition < minimumBoundary Then
        candidatePosition = InStrRev(candidateText, " ")
        If candidatePosition >= minimumBoundary Then boundaryPosition = candidatePosition
    End If
    If boundaryPosition >= minimumBoundary Then candidateText = Left$(candidateText, boundaryPosition)
    ManagedDescriptionPrefixWithinByteBudget = RTrim$(candidateText)
End Function

' Purpose: Resolves and returns normalized duplicate text key from the supplied inputs and the authoritative ADC definitions/state used by this module.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   valueText (Input; String): Input value text for normalized duplicate text key; is tested by If trimValue; is tested by If Len(valueText) = 0.
'   trimValue (Input; Boolean): True trims surrounding whitespace before the value is compared/stored; False preserves the supplied text exactly.
' Returns: String: Returns the normalized duplicate text key text after applying exactly the canonicalization rules implemented by this helper; vbNullString is preserved/returned when no usable text remains.
Public Function NormalizedDuplicateTextKey(ByVal valueText As String, ByVal trimValue As Boolean) As String
    'Optionally trims text and returns an uppercase key for case-insensitive duplicate grouping.
    If trimValue Then valueText = Trim$(valueText)
    If Len(valueText) = 0 Then Exit Function
    NormalizedDuplicateTextKey = UCase$(valueText)
End Function

' Purpose: Resolves trace file safe text from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   valueText (Input; Variant): Input value text for trace file safe text; is passed to `CStr`.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function TraceFileSafeText(ByVal valueText As Variant) As String
    'Preserves trace text exactly while converting Null to an empty string; the physical-log writer owns readable wrapping and embedded-line handling.
    TraceFileSafeText = CStr(Nz(valueText, vbNullString))
End Function

' Purpose: Returns the final file-name component from a path after normalizing slash direction.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   filePath (Input; String): Filesystem or database path identifying the file path location used by the operation.
' Returns: String: Returns the physical/logical name resolved for file name from path under the module's naming/definition rules; vbNullString indicates that no applicable name can be resolved where supported.
Public Function FileNameFromPath(ByVal filePath As String) As String
    'Returns the final file-name component from a path after normalizing slash direction.
    Dim separatorPosition As Long
    separatorPosition = InStrRev(filePath, "\")
    If separatorPosition > 0 Then
        FileNameFromPath = Mid$(filePath, separatorPosition + 1)
    Else
        FileNameFromPath = filePath
    End If
End Function

' Purpose: Resolves append recovery payload line from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   payloadText (Input; String): Input payload text for append recovery payload line; is tested by If Len(payloadText) = 0.
'   lineText (Input; String): Input line text for append recovery payload line; is tested by If Len(payloadText) = 0.
' Returns: String: Returns the append recovery payload line text value. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function AppendRecoveryPayloadLine(ByVal payloadText As String, ByVal lineText As String) As String
    'Appends one line to a recovery payload using CRLF only when prior payload text exists.
    If Len(payloadText) = 0 Then AppendRecoveryPayloadLine = lineText Else AppendRecoveryPayloadLine = payloadText & vbCrLf & lineText
End Function

Private Function UIDurableDecode(ByVal sourceText As String) As String
    'Decodes UI durable-state percent escapes back to tabs, line breaks, and percent signs.
    sourceText = Replace(sourceText, "%09", vbTab)
    sourceText = Replace(sourceText, "%0D%0A", vbCrLf)
    sourceText = Replace(sourceText, "%0D", vbCr)
    sourceText = Replace(sourceText, "%0A", vbLf)
    sourceText = Replace(sourceText, "%25", "%")
    UIDurableDecode = sourceText
End Function

' Purpose: Resolves UI durable encode from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   sourceText (Input; String): Input source text examined or transformed by this procedure according to its documented UI durable encode contract.
' Returns: String: Returns the UI durable encode text value. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function UIDurableEncode(ByVal sourceText As String) As String
    'Percent-encodes tabs, line breaks, and percent signs for UI durable-state storage.
    sourceText = Replace(sourceText, "%", "%25")
    sourceText = Replace(sourceText, vbCrLf, "%0D%0A")
    sourceText = Replace(sourceText, vbCr, "%0D")
    sourceText = Replace(sourceText, vbLf, "%0A")
    sourceText = Replace(sourceText, vbTab, "%09")
    UIDurableEncode = sourceText
End Function

' Purpose: Resolves UI durable note value from the supplied/current ADC state, using `StrComp` and `UIDurableDecode` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   notesText (Input; String): Notes text text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   keyName (Input; String): Canonical or physical name identifying the key targeted by this operation.
' Returns: String: Returns the UI durable note value text value. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function UIDurableNoteValue(ByVal notesText As String, ByVal keyName As String) As String
    'Finds a named key in line-oriented UI durable notes and returns its decoded value.
    Dim lineItem As Variant
    Dim prefixText As String
    prefixText = keyName & "="
    For Each lineItem In Split(notesText, vbCrLf)
        If StrComp(Left$(CStr(lineItem), Len(prefixText)), prefixText, vbTextCompare) = 0 Then
            UIDurableNoteValue = UIDurableDecode(Mid$(CStr(lineItem), Len(prefixText) + 1))
            Exit Function
        End If
    Next lineItem
End Function

' Purpose: Resolves null display text from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   valueText (Input; String): Input value text for null display text; is tested by If Len(Trim$(valueText)) = 0; is assigned to `NullDisplayText`.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function NullDisplayText(ByVal valueText As String) As String
    'Returns (not set) for blank or whitespace-only text and otherwise preserves the supplied display text.
    If Len(Trim$(valueText)) = 0 Then
        NullDisplayText = "(not set)"
    Else
        NullDisplayText = valueText
    End If
End Function

' Purpose: Trims field-input text and removes one matching pair of outer square brackets when present.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   value (Input; String): Input value for normalize field input; is tested by If Len(value) >= 2; is tested by If Left$(value, 1) = "[" And Right$(value, 1) = "]".
' Returns: String: Returns the normalized field input text after applying exactly the canonicalization rules implemented by this helper; vbNullString is preserved/returned when no usable text remains.
Public Function NormalizeFieldInput(ByVal value As String) As String
    'Trims field-input text and removes one matching pair of outer square brackets when present.
    value = Trim$(value)
    If Len(value) >= 2 Then
        If Left$(value, 1) = "[" And Right$(value, 1) = "]" Then value = Mid$(value, 2, Len(value) - 2)
    End If
    NormalizeFieldInput = Trim$(value)
End Function

' Purpose: Normalizes a configured database-role value to the canonical supported role text.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   ConfigOwnerDatabaseRole (Input; String): Identifies the database role whose configuration is authoritative for this operation, such as the CodeDb or HostDb role defined by ADC configuration.
' Returns: String: Returns the normalized config owner database role text after applying exactly the canonicalization rules implemented by this helper; vbNullString is preserved/returned when no usable text
'   remains.
Public Function NormalizedConfigOwnerDatabaseRole(ByVal ConfigOwnerDatabaseRole As String) As String
    'Normalizes a configured database-role value to the canonical supported role text.
    'Normalizes supported Config owner database-role names and aliases to the authoritative HostDb or CodeDb role.
    Select Case LCase$(Trim$(ConfigOwnerDatabaseRole))
        Case LCase$(CONFIG_OWNER_DATABASE_ROLE_HOST_DB), "hostdb"
            NormalizedConfigOwnerDatabaseRole = CONFIG_OWNER_DATABASE_ROLE_HOST_DB
        Case LCase$(CONFIG_OWNER_DATABASE_ROLE_CODE_DB), "code", "library", "lib", "referenced", "reference"
            NormalizedConfigOwnerDatabaseRole = CONFIG_OWNER_DATABASE_ROLE_CODE_DB
        Case Else
            NormalizedConfigOwnerDatabaseRole = vbNullString
    End Select
End Function

' Purpose: Normalizes a database path for stable case-insensitive path comparison.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   value (Input; String): Input value for normalized database path; is passed to `CStr`; is passed to `Replace`.
' Returns: String: Returns the normalized database path text after applying exactly the canonicalization rules implemented by this helper; vbNullString is preserved/returned when no usable text remains.
Public Function NormalizedDatabasePath(ByVal value As String) As String
    'Normalizes a database path for stable case-insensitive path comparison.
    'Intended use: Returns the Normalized Database Path result for direct, query, form, or required cross-module callers.
    On Error GoTo Fallback
    Dim fso As Object
    Dim pathText As String
    pathText = Trim$(value)
    pathText = Replace(pathText, "/", "\")
    If Len(pathText) = 0 Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    NormalizedDatabasePath = LCase$(fso.GetAbsolutePathName(pathText))
    Exit Function
Fallback:
    NormalizedDatabasePath = LCase$(Replace(Trim$(value), "/", "\"))
End Function

' Purpose: Extracts a database file base name from a full path and removes the extension.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   databasePath (Input; String): Filesystem or database path identifying the database path location used by the operation.
' Returns: String: Returns the physical/logical name resolved for database base name from path under the module's naming/definition rules; vbNullString indicates that no applicable name can be resolved where
'   supported.
Public Function DatabaseBaseNameFromPath(ByVal databasePath As String) As String
    'Extracts a database file base name from a full path and removes the extension.
    Dim slashPos As Long
    Dim dotPos As Long
    databasePath = Replace(Trim$(databasePath), "/", "\")
    slashPos = InStrRev(databasePath, "\")
    If slashPos > 0 Then databasePath = Mid$(databasePath, slashPos + 1)
    dotPos = InStrRev(databasePath, ".")
    If dotPos > 1 Then databasePath = Left$(databasePath, dotPos - 1)
    DatabaseBaseNameFromPath = databasePath
End Function

Private Function CommonLengthPrefixedSignaturePart(ByVal valueText As String) As String
    'Serializes one text value as a length-prefixed signature segment.
    CommonLengthPrefixedSignaturePart = CStr(Len(valueText)) & ":" & valueText & ";"
End Function
' Purpose: Resolves formatting insensitive query SQL key from the supplied/current ADC state, using `End` and `ElseIf` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   sqlText (Input; String): SQL text supplied to or resolved by the routine; it is treated as the query/action expression described by this parameter rather than as display text.
' Returns: String: Returns the SQL text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function FormattingInsensitiveQuerySqlKey(ByVal sqlText As String) As String
    'Produces a comparison key for query SQL while ignoring non-semantic formatting differences.
    ' Returns a stable SQL comparison key that ignores formatting-only differences while preserving token boundaries and quoted/bracketed content.
    Dim charIndex As Long
    Dim charText As String
    Dim quoteText As String
    Dim bracketContent As String
    Dim pendingWhitespace As Boolean
    Dim resultText As String
    Dim previousText As String
    Dim nextText As String
    sqlText = Replace(Replace(sqlText, vbCrLf, vbLf), vbCr, vbLf)
    charIndex = 1
    Do While charIndex <= Len(sqlText)
        PumpSetupResponsiveness
        charText = Mid$(sqlText, charIndex, 1)
        If Len(quoteText) > 0 Then
            resultText = resultText & charText
            If charText = quoteText Then
                If charIndex < Len(sqlText) And Mid$(sqlText, charIndex + 1, 1) = quoteText Then
                    resultText = resultText & quoteText
                    charIndex = charIndex + 1
                Else
                    quoteText = vbNullString
                End If
            End If
        ElseIf charText = "'" Or charText = Chr$(34) Then
            If pendingWhitespace Then
                previousText = Right$(resultText, 1)
                If SharedVba_IsIdentifierChar(previousText) Then resultText = resultText & " "
                pendingWhitespace = False
            End If
            quoteText = charText
            resultText = resultText & charText
        ElseIf charText = "[" Then
            bracketContent = vbNullString
            charIndex = charIndex + 1
            Do While charIndex <= Len(sqlText)
                PumpSetupResponsiveness
                charText = Mid$(sqlText, charIndex, 1)
                If charText = "]" Then
                    If charIndex < Len(sqlText) And Mid$(sqlText, charIndex + 1, 1) = "]" Then
                        bracketContent = bracketContent & "]"
                        charIndex = charIndex + 1
                    Else
                        Exit Do
                    End If
                Else
                    bracketContent = bracketContent & charText
                End If
                charIndex = charIndex + 1
            Loop
            If pendingWhitespace Then
                previousText = Right$(resultText, 1)
                If SharedVba_IsIdentifierChar(previousText) Then resultText = resultText & " "
                pendingWhitespace = False
            End If
            If SqlComparisonBracketContentIsSimpleIdentifier(bracketContent) Then
                resultText = resultText & LCase$(bracketContent)
            Else
                resultText = resultText & "[" & LCase$(bracketContent) & "]"
            End If
        ElseIf charText = ";" Then
            pendingWhitespace = False
        ElseIf charText = " " Or charText = vbTab Or charText = vbLf Then
            pendingWhitespace = True
        Else
            If pendingWhitespace Then
                previousText = Right$(resultText, 1)
                nextText = charText
                If SharedVba_IsIdentifierChar(previousText) And SharedVba_IsIdentifierChar(nextText) Then resultText = resultText & " "
                pendingWhitespace = False
            End If
            resultText = resultText & LCase$(charText)
        End If
        charIndex = charIndex + 1
    Loop
    FormattingInsensitiveQuerySqlKey = Trim$(resultText)
End Function

Private Function SqlComparisonBracketContentIsSimpleIdentifier(ByVal bracketContent As String) As Boolean
    'Tests whether bracketed SQL text contains only a simple identifier.
    Dim charIndex As Long
    If Len(bracketContent) = 0 Then Exit Function
    If Not (Left$(bracketContent, 1) Like "[A-Za-z_]") Then Exit Function
    For charIndex = 2 To Len(bracketContent)
        If (charIndex And 255) = 0 Then PumpSetupResponsiveness
        If Not (Mid$(bracketContent, charIndex, 1) Like "[A-Za-z0-9_]") Then Exit Function
    Next charIndex
    SqlComparisonBracketContentIsSimpleIdentifier = True
End Function

' Purpose: Derives the logical Navigation Pane group name by removing configured affixes.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   groupName (Input; String): Name selecting the logical Ribbon, form, catalog, or definition group affected by the operation.
'   systemMacroGroup (Input Optional; Boolean): Boolean state describing whether system macro group applies to this invocation; True selects the named condition/path and False selects its complementary/default
'   path. Default: False.
' Returns: String: Returns the physical/logical name resolved for navigation pane logical group name under the module's naming/definition rules; vbNullString indicates that no applicable name can be resolved
'   where supported.
Public Function NavigationPaneLogicalGroupName(ByVal groupName As String, Optional ByVal systemMacroGroup As Boolean = False) As String
    'Derives the logical Navigation Pane group name by removing configured affixes.
    Dim prefixText As String
    Dim postfixText As String
    Dim logicalName As String
    logicalName = groupName
    NavigationPaneGroupAffixes systemMacroGroup, prefixText, postfixText
    If Len(prefixText) > 0 Then
        If StrComp(Left$(logicalName, Len(prefixText)), prefixText, vbBinaryCompare) = 0 Then logicalName = Mid$(logicalName, Len(prefixText) + 1)
    End If
    If Len(postfixText) > 0 And Len(logicalName) >= Len(postfixText) Then
        If StrComp(Right$(logicalName, Len(postfixText)), postfixText, vbBinaryCompare) = 0 Then logicalName = Left$(logicalName, Len(logicalName) - Len(postfixText))
    End If
    NavigationPaneLogicalGroupName = logicalName
End Function

' Purpose: Returns True when supplied text contains only decimal digits.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   textValue (Input; String): Source text being normalized, parsed, compared, bounded, or converted by this helper; the function does not reinterpret unrelated caller state.
' Returns: Boolean: True when supplied text contains only decimal digits. False otherwise.
Public Function TextContainsDecimalDigitsOnly(ByVal textValue As String) As Boolean
    'Returns True when supplied text contains only decimal digits.
    'Intended use: Tests whether a naming token contains decimal digits only for shared naming transformations.
    Dim indexValue As Long
    Dim currentCharacter As String
    textValue = Trim$(textValue)
    If Len(textValue) = 0 Then Exit Function
    For indexValue = 1 To Len(textValue)
        currentCharacter = Mid$(textValue, indexValue, 1)
        If currentCharacter < "0" Or currentCharacter > "9" Then Exit Function
    Next indexValue
    TextContainsDecimalDigitsOnly = True
End Function

' Purpose: Returns True when text contains a recognized warning condition.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   sourceText (Input; String): Input source text for common text indicates warning; is passed to `CStr`.
' Returns: Boolean: True when text contains a recognized warning condition. False otherwise.
Public Function CommonTextIndicatesWarning(ByVal sourceText As String) As Boolean
    'Returns True when text contains a recognized warning condition.
    'Returns the authoritative warning classification for Access Development Console operation, table, and UI notes.
    Dim normalizedText As String
    normalizedText = LCase$(Trim$(sourceText))
    If Len(normalizedText) = 0 Then Exit Function
    If InStr(1, normalizedText, "could not", vbTextCompare) > 0 Then CommonTextIndicatesWarning = True
    If TextContainsFailedCondition(normalizedText) Then CommonTextIndicatesWarning = True
    If TextContainsFailureCondition(normalizedText) Then CommonTextIndicatesWarning = True
    If InStr(1, normalizedText, "err.", vbTextCompare) > 0 Then CommonTextIndicatesWarning = True
    If InStr(1, normalizedText, "error ", vbTextCompare) > 0 Then CommonTextIndicatesWarning = True
    If InStr(1, normalizedText, "not available", vbTextCompare) > 0 Then CommonTextIndicatesWarning = True
    If InStr(1, normalizedText, "not added", vbTextCompare) > 0 Then CommonTextIndicatesWarning = True
    If InStr(1, normalizedText, "not cleared", vbTextCompare) > 0 Then CommonTextIndicatesWarning = True
    If InStr(1, normalizedText, "not found", vbTextCompare) > 0 Then CommonTextIndicatesWarning = True
    If InStr(1, normalizedText, " is missing", vbTextCompare) > 0 Then CommonTextIndicatesWarning = True
    If InStr(1, normalizedText, " was missing", vbTextCompare) > 0 Then CommonTextIndicatesWarning = True
    If InStr(1, normalizedText, "did not contain", vbTextCompare) > 0 Then CommonTextIndicatesWarning = True
    If InStr(1, normalizedText, "skipped forms / setup notes", vbTextCompare) > 0 Then CommonTextIndicatesWarning = True
    If InStr(1, normalizedText, "skipped/missing", vbTextCompare) > 0 Then CommonTextIndicatesWarning = True
    If InStr(1, normalizedText, "skipped backup", vbTextCompare) > 0 Then CommonTextIndicatesWarning = True
    If InStr(1, normalizedText, "stopped early", vbTextCompare) > 0 Then CommonTextIndicatesWarning = True
    If MetricValueGreaterThanZero(normalizedText, "missing/skipped=") Then CommonTextIndicatesWarning = True
    If MetricValueGreaterThanZero(normalizedText, "missing/skipped:") Then CommonTextIndicatesWarning = True
    If MetricValueGreaterThanZero(normalizedText, "warnings=") Then CommonTextIndicatesWarning = True
    If MetricValueGreaterThanZero(normalizedText, "warnings:") Then CommonTextIndicatesWarning = True
    If MetricValueGreaterThanZero(normalizedText, "errors=") Then CommonTextIndicatesWarning = True
    If MetricValueGreaterThanZero(normalizedText, "errors:") Then CommonTextIndicatesWarning = True
    If MetricValueGreaterThanZero(normalizedText, "failed=") Then CommonTextIndicatesWarning = True
    If MetricValueGreaterThanZero(normalizedText, "failed:") Then CommonTextIndicatesWarning = True
    If MetricValueGreaterThanZero(normalizedText, "missing=") Then CommonTextIndicatesWarning = True
    If MetricValueGreaterThanZero(normalizedText, "missing:") Then CommonTextIndicatesWarning = True
End Function

Private Function TextContainsFailedCondition(ByVal normalizedText As String) As Boolean
    'Returns True when normalized text contains a recognized failed-condition phrase.
    Dim searchPosition As Long
    Dim failedPosition As Long
    Dim scanPosition As Long
    Dim digitText As String
    searchPosition = 1
    Do
        failedPosition = InStr(searchPosition, normalizedText, " failed", vbTextCompare)
        If failedPosition = 0 Then Exit Do
        scanPosition = failedPosition + Len(" failed")
        Do While scanPosition <= Len(normalizedText) And Mid$(normalizedText, scanPosition, 1) = " "
            scanPosition = scanPosition + 1
        Loop
        If scanPosition <= Len(normalizedText) And (Mid$(normalizedText, scanPosition, 1) = "=" Or Mid$(normalizedText, scanPosition, 1) = ":") Then
            scanPosition = scanPosition + 1
            Do While scanPosition <= Len(normalizedText) And Mid$(normalizedText, scanPosition, 1) = " "
                scanPosition = scanPosition + 1
            Loop
            digitText = vbNullString
            Do While scanPosition <= Len(normalizedText) And Mid$(normalizedText, scanPosition, 1) Like "#"
                digitText = digitText & Mid$(normalizedText, scanPosition, 1)
                scanPosition = scanPosition + 1
            Loop
            If Len(digitText) = 0 Then
                TextContainsFailedCondition = True
                Exit Function
            End If
            If CLng(digitText) > 0 Then
                TextContainsFailedCondition = True
                Exit Function
            End If
            searchPosition = scanPosition
        Else
            TextContainsFailedCondition = True
            Exit Function
        End If
    Loop
End Function

' Purpose: Resolves fingerprint token part from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   valueText (Input; String): Input value text for fingerprint token part; is passed to `CStr`.
' Returns: String: Returns the proof/state token produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function FingerprintTokenPart(ByVal valueText As String) As String
    'Encodes one text value as a stable fingerprint token segment.
    'Encodes one value as an unambiguous length-prefixed component for ADC fingerprint and state-token payloads.
    FingerprintTokenPart = CStr(Len(valueText)) & ":" & valueText & "|"
End Function

' Purpose: Computes the canonical ADC Adler-style hexadecimal hash for supplied text.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   textValue (Input; String): Source text being normalized, parsed, compared, bounded, or converted by this helper; the function does not reinterpret unrelated caller state.
' Returns: String: Returns the deterministic adler text hash token/signature produced from the supplied state, allowing callers to compare exact state without repeating the underlying scan/build work.
Public Function AdlerTextHash(ByVal textValue As String) As String
    'Computes the canonical ADC Adler-style hexadecimal hash for supplied text.
    'Intended use: Returns the canonical ADC Adler-style eight-hex-character text hash used by generated-form, ribbon, and comparison fingerprints while periodically pumping setup-control events.
    Dim a As Long
    Dim b As Long
    Dim byteIndex As Long
    Dim charValue As Long
    Dim textBytes() As Byte
    a = 1
    If Len(textValue) > 0 Then
        textBytes = textValue
        For byteIndex = 0 To UBound(textBytes) Step 2
            If (byteIndex And 255) = 0 Then PumpSetupResponsiveness
            charValue = CLng(textBytes(byteIndex)) + CLng(textBytes(byteIndex + 1)) * 256&
            a = (a + charValue) Mod 65521
            b = (b + a) Mod 65521
        Next byteIndex
    End If
    AdlerTextHash = Right$("0000" & Hex$(b), 4) & Right$("0000" & Hex$(a), 4)
End Function

' Purpose: Resolves first text difference position from the supplied/current ADC state, using `If` and `StrComp` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   leftText (Input; String): Input left text for first text difference position; is passed to `Len`; is tested by If StrComp(Mid$(leftText, charIndex, 1), Mid$(rightText, charIndex, 1), vbBinaryCompare) <> 0.
'   rightText (Input; String): Input right text for first text difference position; is tested by If Len(rightText) < compareLength; is tested by If StrComp(Mid$(leftText, charIndex, 1), Mid$(rightText, charIndex,
'     1), vbBinaryCompare) <> 0.
' Returns: Long: Returns the numeric first text difference position result; zero and other sentinel values retain the meanings established by the routine's existing implementation.
Public Function FirstTextDifferencePosition(ByVal leftText As String, ByVal rightText As String) As Long
    'Returns the first character position where two text values differ.
    'Intended use: Returns the first binary character position where two ADC-managed text values differ, or zero when identical, while periodically pumping setup-control events.
    Dim compareLength As Long
    Dim charIndex As Long
    compareLength = Len(leftText)
    If Len(rightText) < compareLength Then compareLength = Len(rightText)
    For charIndex = 1 To compareLength
        If (charIndex And 255) = 0 Then PumpSetupResponsiveness
        If StrComp(Mid$(leftText, charIndex, 1), Mid$(rightText, charIndex, 1), vbBinaryCompare) <> 0 Then
            FirstTextDifferencePosition = charIndex
            Exit Function
        End If
    Next charIndex
    If Len(leftText) <> Len(rightText) Then FirstTextDifferencePosition = compareLength + 1
End Function

' Purpose: Returns a bounded excerpt around a specified text-difference position.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   textValue (Input; String): Source text being normalized, parsed, compared, bounded, or converted by this helper; the function does not reinterpret unrelated caller state.
'   differencePosition (Input; Long): Numeric position locating the difference item within its source, UI layout, or Navigation Pane ordering.
' Returns: String: Returns formatted text comparison excerpt text intended for tracing, diagnostics, or user-facing reporting; vbNullString is returned when there is nothing applicable to report.
Public Function TextComparisonExcerpt(ByVal textValue As String, ByVal differencePosition As Long) As String
    'Returns a bounded excerpt around a specified text-difference position.
    'Intended use: Formats the surrounding text for ADC mismatch diagnostics using the canonical 45-character lead, 90-character window, and visible CR/LF markers.
    Dim startPosition As Long
    startPosition = differencePosition - 45
    If startPosition < 1 Then startPosition = 1
    TextComparisonExcerpt = Mid$(textValue, startPosition, 90)
    TextComparisonExcerpt = Replace(TextComparisonExcerpt, vbCr, "<CR>")
    TextComparisonExcerpt = Replace(TextComparisonExcerpt, vbLf, "<LF>")
End Function

' Purpose: Resolves clean progress scope label from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   scopeText (Input; String): Input scope text for clean progress scope label; is tested by If Right$(scopeText, 1) = ":"; is passed to `CleanProgressScopeLabel`.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function CleanProgressScopeLabel(ByVal scopeText As String) As String
    'Normalizes progress-scope label text for display.
    'Returns a trimmed progress scope label without a trailing colon.
    scopeText = Trim$(scopeText)
    If Right$(scopeText, 1) = ":" Then scopeText = Left$(scopeText, Len(scopeText) - 1)
    CleanProgressScopeLabel = Trim$(scopeText)
End Function

Private Function CanonicalTableQueryObjectName(ByVal namingCategory As String, ByVal objectName As String, ByVal reverseLookup As Boolean, ByRef mappedName As String) As Boolean
    'Resolves a table/query object name against the canonical configured physical-name map without altering the supplied name.
    Dim definitionData As Variant
    mappedName = vbNullString
    If Len(objectName) = 0 Then Exit Function
    For Each definitionData In CanonicalTableQueryNamingDefinitions()
        If StrComp(CStr(definitionData(0)), namingCategory, vbTextCompare) = 0 Then
            If reverseLookup Then
                If StrComp(CStr(definitionData(2)), objectName, vbTextCompare) = 0 Then mappedName = CStr(definitionData(1)): CanonicalTableQueryObjectName = True: Exit Function
            Else
                If StrComp(CStr(definitionData(1)), objectName, vbTextCompare) = 0 Then mappedName = CStr(definitionData(2)): CanonicalTableQueryObjectName = True: Exit Function
            End If
        End If
    Next definitionData
End Function

Private Function CanonicalGeneratedFormObjectName(ByVal objectName As String, ByVal reverseLookup As Boolean, ByRef mappedName As String, Optional ByRef ChatGPTUIForm As Boolean = False) As Boolean
    'Resolves a generated-form object name against the canonical configured physical-name map without altering the supplied name.
    'Bidirectional resolver over Defaults-owned generated-form naming data. ChatGPTUIForm=True appends the fixed _@ChatGptUI namespace suffix to the canonical physical stem.
    Dim definitionData As Variant
    Dim physicalName As String
    mappedName = vbNullString
    ChatGPTUIForm = False
    If Len(objectName) = 0 Then Exit Function
    For Each definitionData In CanonicalGeneratedFormNamingDefinitions()
        ChatGPTUIForm = CBool(definitionData(GENERATED_FORM_NAMING_CHATGPT_UI_FORM))
        physicalName = CStr(definitionData(GENERATED_FORM_NAMING_PHYSICAL_NAME))
        If ChatGPTUIForm Then physicalName = physicalName & DEFAULT_CHATGPT_UI_OBJECT_POSTFIX
        If reverseLookup Then
            If StrComp(physicalName, objectName, vbTextCompare) = 0 Then mappedName = CStr(definitionData(GENERATED_FORM_NAMING_LOGICAL_ID)): CanonicalGeneratedFormObjectName = True: Exit Function
        Else
            If StrComp(CStr(definitionData(GENERATED_FORM_NAMING_LOGICAL_ID)), objectName, vbTextCompare) = 0 Then mappedName = physicalName: CanonicalGeneratedFormObjectName = True: Exit Function
        End If
    Next definitionData
    ChatGPTUIForm = False
End Function

Private Function CanonicalReportObjectName(ByVal objectName As String, ByVal reverseLookup As Boolean, ByRef mappedName As String, Optional ByRef chatGPTUIReport As Boolean = False) As Boolean
    'Resolves a report object name against the canonical configured physical-name map without altering the supplied name.
    'Bidirectional resolver over Defaults-owned report naming data. ChatGPTUIReport=True appends the fixed _@ChatGptUI namespace suffix to the canonical physical stem.
    Dim definitionData As Variant
    Dim physicalName As String
    mappedName = vbNullString
    chatGPTUIReport = False
    If Len(objectName) = 0 Then Exit Function
    For Each definitionData In CanonicalReportNamingDefinitions()
        chatGPTUIReport = CBool(definitionData(REPORT_NAMING_CHATGPT_UI_REPORT))
        physicalName = CStr(definitionData(REPORT_NAMING_PHYSICAL_NAME))
        If chatGPTUIReport Then physicalName = physicalName & DEFAULT_CHATGPT_UI_OBJECT_POSTFIX
        If reverseLookup Then
            If StrComp(physicalName, objectName, vbTextCompare) = 0 Then mappedName = CStr(definitionData(REPORT_NAMING_LOGICAL_ID)): CanonicalReportObjectName = True: Exit Function
        Else
            If StrComp(CStr(definitionData(REPORT_NAMING_LOGICAL_ID)), objectName, vbTextCompare) = 0 Then mappedName = physicalName: CanonicalReportObjectName = True: Exit Function
        End If
    Next definitionData
    chatGPTUIReport = False
End Function
' Purpose: Derives a logical managed-object name from a configured physical name.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   namingCategory (Input; String): Category value classifying the naming item so the routine can select the correct naming, logging, or processing rules.
'   physicalObjectName (Input; String): Exact physical Access object name being mapped to/from its managed logical identity.
' Returns: String: Returns the physical/logical name resolved for logical object name from physical name under the module's naming/definition rules; vbNullString indicates that no applicable name can be
'   resolved where supported.
Public Function LogicalObjectNameFromPhysicalName(ByVal namingCategory As String, ByVal physicalObjectName As String) As String
    'Derives a logical managed-object name from a configured physical name.
    Dim prefixText As String
    Dim postfixText As String
    Dim logicalName As String
    Dim mappedName As String
    Dim ChatGPTUIForm As Boolean
    Dim chatGPTUIReport As Boolean
    logicalName = physicalObjectName
    If StrComp(namingCategory, NAMING_CATEGORY_GENERATED_FORM, vbTextCompare) = 0 Then
        If CanonicalGeneratedFormObjectName(logicalName, True, mappedName, ChatGPTUIForm) Then LogicalObjectNameFromPhysicalName = mappedName Else LogicalObjectNameFromPhysicalName = logicalName
        Exit Function
    ElseIf StrComp(namingCategory, NAMING_CATEGORY_REPORT, vbTextCompare) = 0 Then
        If CanonicalReportObjectName(logicalName, True, mappedName, chatGPTUIReport) Then LogicalObjectNameFromPhysicalName = mappedName Else LogicalObjectNameFromPhysicalName = logicalName
        Exit Function
    End If
    NamingAffixes namingCategory, prefixText, postfixText
    If Len(prefixText) > 0 Then
        If StrComp(Left$(logicalName, Len(prefixText)), prefixText, vbBinaryCompare) = 0 Then logicalName = Mid$(logicalName, Len(prefixText) + 1)
    End If
    If Len(postfixText) > 0 And Len(logicalName) >= Len(postfixText) Then
        If StrComp(Right$(logicalName, Len(postfixText)), postfixText, vbBinaryCompare) = 0 Then logicalName = Left$(logicalName, Len(logicalName) - Len(postfixText))
    End If
    If CanonicalTableQueryObjectName(namingCategory, logicalName, True, mappedName) Then
        LogicalObjectNameFromPhysicalName = mappedName
        Exit Function
    End If
    LogicalObjectNameFromPhysicalName = logicalName
End Function

' Purpose: Parses persisted ADC Boolean text into a Boolean using the accepted true/false spellings, returns DefaultValue for blank input, and raises an error for unsupported text.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   valueText (Input; String): Input value text for boolean text value; is tested by If Len(valueText) = 0.
'   DefaultValue (Input; Boolean): Fallback Boolean used when the setting/source has no usable persisted value; the routine returns/applies this value in that absence case.
' Returns: Boolean: Returns the Boolean interpretation of the supplied text/value according to ADC's accepted true/false representations and fallback behavior.
Public Function BooleanTextValue(ByVal valueText As String, ByVal DefaultValue As Boolean) As Boolean
    'Parses canonical persisted ADC Boolean text; storage/read ownership remains with the calling domain.
    'Canonical stateless interpretation grammar for persisted ADC Boolean Settings text; storage/read paths remain owned by their domains.
    valueText = LCase$(Trim$(valueText))
    If Len(valueText) = 0 Then
        BooleanTextValue = DefaultValue
        Exit Function
    End If
    Select Case valueText
        Case "true", "yes", "y", "1", "-1", "on"
            BooleanTextValue = True
        Case "false", "no", "n", "0", "off"
            BooleanTextValue = False
        Case Else
            Err.Raise vbObjectError + 6901, "BooleanTextValue", "Access Development Console Boolean setting text is invalid." & vbCrLf & "Stored value: " & valueText & vbCrLf & "Allowed values: True, Yes, Y, 1, -1, On, False, No, N, 0, or Off."
    End Select
End Function

' Purpose: Applies configured product branding substitutions to display text.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   sourceText (Input; String): Input source text examined or transformed by this procedure according to its documented apply display branding contract.
' Returns: String: Returns the resolved apply display branding text in the canonical representation consumed by this module's caller; vbNullString represents no applicable value where the existing contract
'   permits it.
Public Function ApplyDisplayBranding(ByVal sourceText As String) As String
    'Applies configured product branding substitutions to display text.
    Dim longName As String
    Dim shortName As String
    longName = UtilityLongName()
    shortName = UtilityShortName()
    sourceText = Replace(sourceText, DEFAULT_UTILITY_LONG_NAME, longName, 1, -1, vbTextCompare)
    sourceText = Replace(sourceText, DEFAULT_RIBBON_TAB_KEY_SYSTEM_1, shortName & "1", 1, -1, vbTextCompare)
    sourceText = Replace(sourceText, DEFAULT_RIBBON_TAB_KEY_SYSTEM_2, shortName & "2", 1, -1, vbTextCompare)
    sourceText = ReplaceWholeWord(sourceText, DEFAULT_UTILITY_SHORT_NAME, shortName)
    sourceText = ReplaceWholeWord(sourceText, "Database", "DB")
    sourceText = Replace(sourceText, " & ", "/", 1, -1, vbBinaryCompare)
    ApplyDisplayBranding = sourceText
End Function

' Purpose: Normalizes diagnostic text into a human-readable popup message.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   promptText (Input; String): Prompt text text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   TitleText (Input; String): Title text text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   buttons (Input; VbMsgBoxStyle): Message-box style/icon flags controlling the visual treatment and available choices for this UI operation.
' Returns: String: Returns the resolved human readable popup text text in the canonical representation consumed by this module's caller; vbNullString represents no applicable value where the existing contract
'   permits it.
Public Function HumanReadablePopupText(ByVal promptText As String, ByVal TitleText As String, ByVal buttons As VbMsgBoxStyle) As String
    'Normalizes diagnostic text into a human-readable popup message.
    On Error GoTo UseOriginal
    Dim sourceText As String
    Dim primaryDiagnosticText As String
    Dim operationText As String
    Dim reasonText As String
    Dim affectedText As String
    Dim stateText As String
    Dim actionText As String
    Dim resultText As String
    Dim iconValue As Long
    sourceText = Trim$(promptText)
    If Len(sourceText) = 0 Then GoTo UseOriginal
    If PopupTextContainsTechnicalDiagnostics(sourceText) = False Then
        HumanReadablePopupText = sourceText
        Exit Function
    End If
    primaryDiagnosticText = PopupPrimaryDiagnosticText(sourceText)
    If PopupTextIndicatesSuccessfulCompletion(primaryDiagnosticText) Then
        HumanReadablePopupText = PopupSuccessText(primaryDiagnosticText)
        Exit Function
    End If
    operationText = PopupOperationText(primaryDiagnosticText, TitleText)
    reasonText = PopupReasonText(primaryDiagnosticText)
    affectedText = PopupAffectedText(primaryDiagnosticText)
    iconValue = CLng(buttons) And &H70
    If iconValue = vbExclamation Then
        stateText = PopupStateText(primaryDiagnosticText, False)
        actionText = PopupActionText(primaryDiagnosticText, True)
        resultText = DEFAULT_UTILITY_SHORT_NAME & " reported a warning while processing " & operationText & "."
    Else
        stateText = PopupStateText(primaryDiagnosticText, True)
        actionText = PopupActionText(primaryDiagnosticText, False)
        resultText = DEFAULT_UTILITY_SHORT_NAME & " could not complete " & operationText & "."
    End If
    If Len(reasonText) > 0 Then resultText = resultText & vbCrLf & vbCrLf & reasonText
    If Len(affectedText) > 0 Then resultText = resultText & vbCrLf & vbCrLf & affectedText
    If Len(stateText) > 0 Then resultText = resultText & vbCrLf & vbCrLf & stateText
    If Len(actionText) > 0 Then resultText = resultText & vbCrLf & vbCrLf & actionText
    resultText = resultText & vbCrLf & vbCrLf & "Detailed technical information was written to TraceLog and the " & DEFAULT_UTILITY_LONG_NAME & " trace log file."
    HumanReadablePopupText = Left$(resultText, 900)
    Exit Function
UseOriginal:
    HumanReadablePopupText = promptText
End Function

Private Function ExtractSetupPartText(ByVal sourceText As String) As String
    'Extracts setup-part text from a diagnostic/status string.
    Dim startPosition As Long
    Dim endPosition As Long
    Dim candidateText As String
    startPosition = InStr(1, sourceText, "Setup Part ", vbTextCompare)
    If startPosition = 0 Then startPosition = InStr(1, sourceText, "Part ", vbTextCompare)
    If startPosition = 0 Then Exit Function
    endPosition = InStr(startPosition, sourceText, " > ", vbBinaryCompare)
    If endPosition = 0 Then endPosition = InStr(startPosition, sourceText, vbCrLf, vbBinaryCompare)
    If endPosition = 0 Then endPosition = InStr(startPosition, sourceText, " | ", vbBinaryCompare)
    If endPosition = 0 Then endPosition = startPosition + 40
    candidateText = Trim$(Mid$(sourceText, startPosition, endPosition - startPosition))
    If Len(candidateText) > 45 Then candidateText = Left$(candidateText, 45)
    If LCase$(Left$(candidateText, 5)) = "part " Then candidateText = "Setup " & candidateText
    ExtractSetupPartText = candidateText
End Function

Private Function FirstDiagnosticValue(ByVal sourceText As String, ByVal labelText As String) As String
    'Extracts the first value associated with a diagnostic label.
    Dim normalizedText As String
    Dim lineItems As Variant
    Dim lineIndex As Long
    Dim lineText As String
    Dim labelPrefix As String
    Dim valueText As String
    normalizedText = NormalizeDiagnosticSegments(sourceText)
    lineItems = Split(normalizedText, vbLf)
    labelPrefix = LCase$(labelText) & ":"
    For lineIndex = LBound(lineItems) To UBound(lineItems)
        lineText = Trim$(CStr(lineItems(lineIndex)))
        If LCase$(Left$(lineText, Len(labelPrefix))) = labelPrefix Then
            valueText = Trim$(Mid$(lineText, Len(labelPrefix) + 1))
            If Len(valueText) > 0 Then
                FirstDiagnosticValue = FirstSentence(valueText, 220)
                Exit Function
            End If
        End If
    Next lineIndex
End Function

Private Function NormalizeDiagnosticSegments(ByVal sourceText As String) As String
    'Normalizes diagnostic segment separators and whitespace.
    Dim normalizedText As String
    normalizedText = Replace(sourceText, vbCrLf, vbLf)
    normalizedText = Replace(normalizedText, vbCr, vbLf)
    normalizedText = Replace(normalizedText, " | ", vbLf)
    NormalizeDiagnosticSegments = normalizedText
End Function

Private Function MeaningfulFailureLine(ByVal sourceText As String) As String
    'Extracts the first meaningful failure line from diagnostic text.
    Dim lineItems As Variant
    Dim lineIndex As Long
    Dim lineText As String
    Dim selectedText As String
    lineItems = Split(NormalizeDiagnosticSegments(sourceText), vbLf)
    For lineIndex = UBound(lineItems) To LBound(lineItems) Step -1
        lineText = Trim$(CStr(lineItems(lineIndex)))
        If LineIsTechnicalLabel(lineText) = False And LineIndicatesFailure(lineText) Then
            selectedText = FirstSentence(lineText, 280)
            Exit For
        End If
    Next lineIndex
    MeaningfulFailureLine = selectedText
End Function

Private Function LineIndicatesFailure(ByVal lineText As String) As Boolean
    'Returns True when one diagnostic line indicates failure.
    Dim token As Variant
    If StrComp(lineText, DEFAULT_UTILITY_LONG_NAME & ": Setup failed.", vbTextCompare) = 0 Then Exit Function
    For Each token In Array("could not", "unable", "failed", "missing", "invalid", "conflict", "not found", "not installed", "not available", "rejected", "mismatch", "too large", "unmapped", "cannot")
        If InStr(1, lineText, CStr(token), vbTextCompare) > 0 Then
            LineIndicatesFailure = True
            Exit Function
        End If
    Next token
End Function

Private Function TextContainsFailureCondition(ByVal normalizedText As String) As Boolean
    'Returns True when diagnostic text contains a recognized failure condition.
    'Treats failure/failures metrics as warnings only when positive, ignores identifier-embedded Failure tokens, and keeps ordinary failure prose warning-bearing.
    Dim searchPosition As Long
    Dim failurePosition As Long
    Dim scanPosition As Long
    Dim digitText As String
    Dim beforeCharacter As String
    Dim afterCharacter As String
    searchPosition = 1
    Do
        failurePosition = InStr(searchPosition, normalizedText, "failure", vbTextCompare)
        If failurePosition = 0 Then Exit Do
        beforeCharacter = vbNullString
        If failurePosition > 1 Then beforeCharacter = Mid$(normalizedText, failurePosition - 1, 1)
        scanPosition = failurePosition + Len("failure")
        If scanPosition <= Len(normalizedText) And Mid$(normalizedText, scanPosition, 1) = "s" Then scanPosition = scanPosition + 1
        afterCharacter = vbNullString
        If scanPosition <= Len(normalizedText) Then afterCharacter = Mid$(normalizedText, scanPosition, 1)
        If beforeCharacter Like "[A-Za-z0-9_.]" Or SharedVba_IsIdentifierChar(afterCharacter) Then
            searchPosition = scanPosition
        Else
            Do While scanPosition <= Len(normalizedText) And Mid$(normalizedText, scanPosition, 1) = " "
                scanPosition = scanPosition + 1
            Loop
            If scanPosition <= Len(normalizedText) And (Mid$(normalizedText, scanPosition, 1) = "=" Or Mid$(normalizedText, scanPosition, 1) = ":") Then
                scanPosition = scanPosition + 1
                Do While scanPosition <= Len(normalizedText) And Mid$(normalizedText, scanPosition, 1) = " "
                    scanPosition = scanPosition + 1
                Loop
                digitText = vbNullString
                Do While scanPosition <= Len(normalizedText) And Mid$(normalizedText, scanPosition, 1) Like "#"
                    digitText = digitText & Mid$(normalizedText, scanPosition, 1)
                    scanPosition = scanPosition + 1
                Loop
                If Len(digitText) = 0 Then
                    TextContainsFailureCondition = True
                    Exit Function
                End If
                If CLng(digitText) > 0 Then
                    TextContainsFailureCondition = True
                    Exit Function
                End If
                searchPosition = scanPosition
            Else
                TextContainsFailureCondition = True
                Exit Function
            End If
        End If
    Loop
End Function

Private Function DaoEngineErrorsText() As String
    'Formats the current DAO engine error collection as readable text.
    On Error Resume Next
    Dim daoErrorItem As DAO.Error
    Dim resultText As String
    For Each daoErrorItem In DBEngine.Errors
        If Len(resultText) > 0 Then resultText = resultText & vbCrLf
        resultText = resultText & "Number=" & CStr(daoErrorItem.Number) & "; Source=" & daoErrorItem.Source & "; Description=" & daoErrorItem.Description
    Next daoErrorItem
    Err.Clear
    On Error GoTo 0
    DaoEngineErrorsText = resultText
End Function

Private Function LineIsTechnicalLabel(ByVal lineText As String) As Boolean
    'Returns True when a line begins with a recognized technical diagnostic label.
    Dim prefix As Variant
    If Len(lineText) = 0 Then
        LineIsTechnicalLabel = True
        Exit Function
    End If
    If Left$(lineText, 3) = "---" Then
        LineIsTechnicalLabel = True
        Exit Function
    End If
    For Each prefix In Array("Err.Number:", "Err.Source:", "Err.Description:", "Procedure:", "Database:", "CodeDb workflow:", "MC_Records table:", "CodeDb settings table:", "Diagnostics:", "Trace table:", "Trace run ID:", "Log file:", "Process:", "Elapsed time:", "Expected ", "Installed ", "Actual ", "Original ", "Requested ", "Existing ", "Property ", "Failure stage:", "Replacement ", "Restoration ", "DBEngine.Errors:", "Stage:", "Verification ", "Manifest ", "Caller:", "Object type:", "Table:", "Field:", "Module:", "Result:", "Rows ", "First failing module:", "What happened:", "Failed check:")
        If LCase$(Left$(lineText, Len(CStr(prefix)))) = LCase$(CStr(prefix)) Then
            LineIsTechnicalLabel = True
            Exit Function
        End If
    Next prefix
End Function

Private Function FirstSentence(ByVal sourceText As String, ByVal maxLength As Long) As String
    'Returns a cleaned and length-bounded first sentence from supplied text.
    Dim cleanedText As String
    Dim sentenceEnd As Long
    cleanedText = Trim$(Replace(Replace(sourceText, vbCr, " "), vbLf, " "))
    Do While InStr(cleanedText, "  ") > 0
        cleanedText = Replace(cleanedText, "  ", " ")
    Loop
    sentenceEnd = InStr(1, cleanedText, ". ", vbBinaryCompare)
    If sentenceEnd > 0 Then cleanedText = Left$(cleanedText, sentenceEnd)
    If Len(cleanedText) > maxLength Then cleanedText = Left$(cleanedText, maxLength - 3) & "..."
    FirstSentence = cleanedText
End Function

Private Function ReplaceWholeWord(ByVal sourceText As String, ByVal findText As String, ByVal replacementText As String) As String
    'Replaces whole-word occurrences in text without changing embedded substrings.
    Dim positionValue As Long
    Dim searchStart As Long
    Dim beforeChar As String
    Dim afterChar As String
    searchStart = 1
    Do
        positionValue = InStr(searchStart, sourceText, findText, vbTextCompare)
        If positionValue = 0 Then Exit Do
        beforeChar = vbNullString
        afterChar = vbNullString
        If positionValue > 1 Then beforeChar = Mid$(sourceText, positionValue - 1, 1)
        If positionValue + Len(findText) <= Len(sourceText) Then afterChar = Mid$(sourceText, positionValue + Len(findText), 1)
        If (Not SharedVba_IsIdentifierChar(beforeChar)) And (Not SharedVba_IsIdentifierChar(afterChar)) Then
            sourceText = Left$(sourceText, positionValue - 1) & replacementText & Mid$(sourceText, positionValue + Len(findText))
            searchStart = positionValue + Len(replacementText)
        Else
            searchStart = positionValue + Len(findText)
        End If
    Loop
    ReplaceWholeWord = sourceText
End Function

Private Function DaoPropertyCharacterLengthText(ByVal propertyValue As Variant) As String
    'Formats the character length of a DAO property value for diagnostics.
    On Error GoTo Unavailable
    If IsNull(propertyValue) Then DaoPropertyCharacterLengthText = "0": Exit Function
    If IsArray(propertyValue) Or IsObject(propertyValue) Then GoTo Unavailable
    DaoPropertyCharacterLengthText = CStr(Len(CStr(propertyValue)))
    Exit Function
Unavailable:
    DaoPropertyCharacterLengthText = "(unavailable)"
End Function

Private Function DaoPropertyByteLengthText(ByVal propertyValue As Variant) As String
    'Formats the byte length of a DAO property value for diagnostics.
    On Error GoTo Unavailable
    If IsNull(propertyValue) Then DaoPropertyByteLengthText = "0": Exit Function
    If IsArray(propertyValue) Or IsObject(propertyValue) Then GoTo Unavailable
    DaoPropertyByteLengthText = CStr(LenB(CStr(propertyValue)))
    Exit Function
Unavailable:
    DaoPropertyByteLengthText = "(unavailable)"
End Function

Private Function DaoPropertyDiagnosticValue(ByVal propertyValue As Variant) As String
    'Formats a DAO property value as bounded single-line diagnostic text.
    On Error GoTo Unavailable
    Dim valueText As String
    If IsNull(propertyValue) Then DaoPropertyDiagnosticValue = "(Null)": Exit Function
    If IsArray(propertyValue) Then DaoPropertyDiagnosticValue = "(array; VarType=" & CStr(VarType(propertyValue)) & ")": Exit Function
    valueText = CStr(propertyValue)
    valueText = Replace(valueText, vbCrLf, " | ")
    valueText = Replace(valueText, vbCr, " | ")
    valueText = Replace(valueText, vbLf, " | ")
    If Len(valueText) > 1000 Then valueText = Left$(valueText, 997) & "..."
    DaoPropertyDiagnosticValue = valueText
    Exit Function
Unavailable:
    DaoPropertyDiagnosticValue = "(unavailable; VarType=" & CStr(VarType(propertyValue)) & ")"
End Function

' Purpose: Resolves and returns join last collection items from the supplied inputs and the authoritative ADC definitions/state used by this module.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   values (Input; Collection): Caller-supplied value collection used as the source set for this operation; when passed ByRef the routine may append/update items according to its documented behavior.
'   delimiter (Input; String): Delimiter text inserted between items when the routine joins or formats the supplied values.
'   maxItems (Input Optional; Long): Upper bound limiting the max items quantity scanned, displayed, or processed by this call. Default: 40.
' Returns: String: Returns the selected collection items joined into one text value using the routine's ordering, item-count, and delimiter rules.
Public Function JoinLastCollectionItems(ByVal values As Collection, ByVal delimiter As String, Optional ByVal maxItems As Long = 40) As String
    'Joins the last requested Collection items and reports omitted earlier items.
    'Joins the last maxItems values from a Collection and reports how many earlier values were omitted.
    Dim index As Long
    Dim startIndex As Long
    Dim resultText As String
    If values Is Nothing Then Exit Function
    If values.count = 0 Then Exit Function
    If maxItems < 1 Then maxItems = 1
    startIndex = values.count - maxItems + 1
    If startIndex < 1 Then startIndex = 1
    For index = startIndex To values.count
        If Len(resultText) > 0 Then resultText = resultText & delimiter
        resultText = resultText & CStr(values(index))
    Next index
    If startIndex > 1 Then
        resultText = "... " & CStr(startIndex - 1) & " earlier trace step(s) omitted" & delimiter & resultText
    End If
    JoinLastCollectionItems = resultText
End Function

' Purpose: Resolves procedure call text from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   ProcedureName (Input; String): Canonical or physical name identifying the procedure targeted by this operation.
'   argumentsText (Input; String): Input arguments text for procedure call text; is tested by If Len(argumentsText) = 0; is passed to `ProcedureCallText`.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function ProcedureCallText(ByVal ProcedureName As String, ByVal argumentsText As String) As String
    'Builds a procedure-call expression from a procedure name and optional argument text.
    ProcedureName = Trim$(ProcedureName)
    argumentsText = Trim$(argumentsText)
    If Len(ProcedureName) = 0 Then Exit Function
    If Len(argumentsText) = 0 Then
        ProcedureCallText = ProcedureName
    Else
        ProcedureCallText = ProcedureName & "(" & argumentsText & ")"
    End If
End Function

' Purpose: Resolves clean description for caption from the supplied/current ADC state, using `Const` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   descriptionText (Input; String): Description text text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
' Returns: String: Returns the description text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function CleanDescriptionForCaption(ByVal descriptionText As String) As String
    'Normalizes description text for a single-line form caption and bounds the result to the caption safety limit.
    Const MAX_DESCRIPTION_CAPTION_LEN As Long = 1200
    Dim cleanedText As String
    cleanedText = Trim$(descriptionText)
    cleanedText = Replace(cleanedText, vbCrLf, " ")
    cleanedText = Replace(cleanedText, vbCr, " ")
    cleanedText = Replace(cleanedText, vbLf, " ")
    cleanedText = Replace(cleanedText, vbTab, " ")
    Do While InStr(1, cleanedText, "  ", vbBinaryCompare) > 0
        cleanedText = Replace(cleanedText, "  ", " ")
    Loop
    If Len(cleanedText) > MAX_DESCRIPTION_CAPTION_LEN Then cleanedText = Left$(cleanedText, MAX_DESCRIPTION_CAPTION_LEN - 3) & "..."
    CleanDescriptionForCaption = cleanedText
End Function

' ================================================================================================================================================================================

' Purpose: Builds one authoritative, action-aware duplicate target signature for acceptance, duplicate review, and health review.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   ConfigOwnerDatabaseRole (Input; String): Identifies the database role whose configuration is authoritative for this operation, such as the CodeDb or HostDb role defined by ADC configuration.
'   MacroAction (Input; String): Action value selecting the macro behavior the routine validates, dispatches, or persists.
'   TargetProcedureName (Input; String): Exact public VBA procedure name the search/validation/invocation logic must locate.
'   TargetArguments (Input; String): Argument data supplying the targets values/signature used to validate or invoke the target procedure.
'   TargetObjectName (Input; String): Exact Access object name whose action/metadata is being resolved or executed.
'   ConfirmationMode (Input; String): Mode value selecting the confirmation processing path for this invocation.
' Returns: String: Returns the deterministic target call signature token/signature produced from the supplied state, allowing callers to compare exact state without repeating the underlying scan/build work.
Public Function TargetCallSignature(ByVal ConfigOwnerDatabaseRole As String, ByVal MacroAction As String, ByVal TargetProcedureName As String, ByVal TargetArguments As String, ByVal TargetObjectName As String, ByVal ConfirmationMode As String) As String
    'Builds one authoritative, action-aware duplicate target signature for acceptance, duplicate review, and health review.
    Dim roleText As String
    Dim actionText As String
    Dim targetNameText As String
    roleText = UCase$(NormalizedConfigOwnerDatabaseRole(ConfigOwnerDatabaseRole))
    actionText = UCase$(Trim$(MacroAction))
    If Len(roleText) = 0 Or Len(actionText) = 0 Then Exit Function
    Select Case actionText
        Case UCase$(ACTION_RUNCODE)
            targetNameText = UCase$(Trim$(TargetProcedureName))
            If Len(targetNameText) = 0 Then Exit Function
            TargetCallSignature = CommonLengthPrefixedSignaturePart(roleText) & CommonLengthPrefixedSignaturePart(actionText) & CommonLengthPrefixedSignaturePart(targetNameText) & CommonLengthPrefixedSignaturePart(TargetArguments)
        Case UCase$(ACTION_OPENFORM), UCase$(ACTION_OPENTABLE), UCase$(ACTION_OPENQUERY), UCase$(ACTION_OPENREPORT), UCase$(ACTION_RUNMACRO)
            targetNameText = UCase$(TargetObjectName)
            If Len(targetNameText) = 0 Then Exit Function
            TargetCallSignature = CommonLengthPrefixedSignaturePart(roleText) & CommonLengthPrefixedSignaturePart(actionText) & CommonLengthPrefixedSignaturePart(targetNameText) & CommonLengthPrefixedSignaturePart(UCase$(Trim$(ConfirmationMode)))
        Case Else
            TargetCallSignature = CommonLengthPrefixedSignaturePart(roleText) & CommonLengthPrefixedSignaturePart(actionText) & CommonLengthPrefixedSignaturePart(UCase$(Trim$(TargetProcedureName))) & CommonLengthPrefixedSignaturePart(TargetArguments) & CommonLengthPrefixedSignaturePart(UCase$(TargetObjectName)) & CommonLengthPrefixedSignaturePart(UCase$(Trim$(ConfirmationMode)))
    End Select
End Function

' Purpose: Resolves naming performance summary from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function NamingPerformanceSummary() As String
    'Formats the current naming-performance counters as a diagnostic summary.
    NamingPerformanceSummary = _
        "Naming-policy snapshots built=" & CStr(mNamingSnapshotBuildCount) & _
        "; managed-name-only refreshes=" & CStr(mNamingManagedNameRefreshCount) & _
        "; invalidations or owned ends=" & CStr(mNamingSnapshotInvalidationCount) & _
        "; fallback naming database lookups=" & CStr(mNamingDatabaseLookupCount)
End Function

' Canonical managed-query no-op comparison authority. Direct lexical equivalence is authoritative when it matches; ACE temporary-QueryDef normalization is only an opportunistic second comparison because some valid
' named saved-query SQL cannot be reparsed through an anonymous QueryDef (for example compound OUTER JOIN predicates). A normalization failure must therefore fall back to the direct key rather than invalidate a
' valid saved query.
' Purpose: Resolves normalized query SQL for comparison from the supplied/current ADC state, using `FormattingInsensitiveQuerySqlKey` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   db (Input; DAO.Database): DAO Database object that defines the database context in which the routine reads, validates, or mutates objects.
'   querySql (Input; String): SQL text supplied to or resolved by the routine; it is treated as the query/action expression described by this parameter rather than as display text.
' Returns: String: Returns the SQL text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function NormalizedQuerySqlForComparison(ByVal db As DAO.Database, ByVal querySql As String) As String
    'Returns the DAO/ACE-normalized form of SQL used for semantic saved-query comparison.
    Dim qdf As DAO.queryDef
    Dim directKey As String
    directKey = FormattingInsensitiveQuerySqlKey(querySql)
    NormalizedQuerySqlForComparison = directKey
    On Error GoTo NormalizationUnavailable
    Set qdf = db.CreateQueryDef(vbNullString, querySql)
    NormalizedQuerySqlForComparison = FormattingInsensitiveQuerySqlKey(qdf.sql)
    On Error Resume Next
    If Not qdf Is Nothing Then qdf.Close
    Err.Clear
    On Error GoTo 0
    Exit Function
NormalizationUnavailable:
    'Comparison is best-effort only. Query owners prove writes through the actual named QueryDef they own; do not turn an anonymous-parser limitation into a setup failure.
    On Error Resume Next
    If Not qdf Is Nothing Then qdf.Close
    Err.Clear
    On Error GoTo 0
    NormalizedQuerySqlForComparison = directKey
End Function

' Purpose: Formats elapsed seconds for Access Development Console status, trace, and popup output.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   elapsedSeconds (Input; Double): Timing value used to calculate or report the elapsed seconds duration for this operation.
' Returns: String: Returns the resolved formatted elapsed seconds text in the canonical representation consumed by this module's caller; vbNullString represents no applicable value where the existing contract
'   permits it.
Public Function FormattedElapsedSeconds(ByVal elapsedSeconds As Double) As String
    'Formats elapsed seconds for Access Development Console status, trace, and popup output.
    Dim totalSeconds As Long
    Dim hoursValue As Long
    Dim minutesValue As Long
    Dim secondsValue As Long
    If elapsedSeconds < 0 Then elapsedSeconds = 0
    If elapsedSeconds < 60 Then
        FormattedElapsedSeconds = Format$(elapsedSeconds, "0.0") & " seconds"
        Exit Function
    End If
    totalSeconds = CLng(elapsedSeconds)
    hoursValue = totalSeconds \ 3600
    minutesValue = (totalSeconds Mod 3600) \ 60
    secondsValue = totalSeconds Mod 60
    If hoursValue > 0 Then
        FormattedElapsedSeconds = CStr(hoursValue) & " hr " & CStr(minutesValue) & " min " & CStr(secondsValue) & " sec"
    Else
        FormattedElapsedSeconds = CStr(minutesValue) & " min " & CStr(secondsValue) & " sec"
    End If
End Function

' Purpose: Resolves normalized object type code from the supplied/current ADC state, using `ObjectTypeDefinitionForCode` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   objectTypeCode (Input; String): Canonical code used to classify or select the object type code behavior.
' Returns: String: Returns the normalized object type code text value. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function NormalizedObjectTypeCode(ByVal objectTypeCode As String) As String
    'Returns the canonical ADC object-type code or blank when the supplied identity is unsupported.
    Dim Definition As Variant
    If ObjectTypeDefinitionForCode(objectTypeCode, Definition) Then NormalizedObjectTypeCode = CStr(Definition(0))
End Function

' Purpose: Returns the canonical ADC object-type code for a supported MSys type or blank for an unrelated MSys row.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   msysType (Input; Long): MSysObjects Type code supplied/inspected when mapping between system-table rows and Access object types.
' Returns: String: Returns the canonical object type code from m sys type value used by downstream policy/dispatch logic; vbNullString indicates that no managed value applies where the API permits absence.
Public Function ObjectTypeCodeFromMSysType(ByVal msysType As Long) As String
    'Returns the canonical ADC object-type code for a supported MSys type or blank for an unrelated MSys row.
    Dim Definition As Variant
    If ObjectTypeDefinitionForMSysType(msysType, Definition) Then ObjectTypeCodeFromMSysType = CStr(Definition(0))
End Function

' Purpose: Resolves supported M sys object types SQL list from the supplied/current ADC state, using `CreateObject` and `ObjectTypeDefinitionRows` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: String: Returns the SQL text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function SupportedMSysObjectTypesSqlList() As String
    'Returns the authoritative comma-separated MSys type set represented by ADC Access object types.
    Dim rowData As Variant
    Dim acceptedTypes As Variant
    Dim acceptedType As Variant
    Dim seen As Object
    Dim keyText As String
    ' INTENTIONAL KEYED LOOKUP: seen provides keyed membership/lookup over source rows or state already read by SupportedMSysObjectTypesSqlList; it is a derived access path, not a second authority.
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbBinaryCompare
    For Each rowData In ObjectTypeDefinitionRows()
        acceptedTypes = rowData(2)
        For Each acceptedType In acceptedTypes
            keyText = CStr(CLng(acceptedType))
            If seen.Exists(keyText) = False Then
                If Len(SupportedMSysObjectTypesSqlList) > 0 Then SupportedMSysObjectTypesSqlList = SupportedMSysObjectTypesSqlList & ","
                SupportedMSysObjectTypesSqlList = SupportedMSysObjectTypesSqlList & keyText
                seen.Add keyText, True
            End If
        Next acceptedType
    Next rowData
End Function

' Purpose: Builds the MSys type predicate for an Access object type; tables intentionally match local, ODBC, and linked table rows.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   objectType (Input; AcObjectType): Object-type value identifying the Access/VBA object class targeted by the operation.
'   fieldExpression (Input; String): Expression text representing the field calculation, RunCode target, SQL fragment, or UI binding interpreted by this routine.
' Returns: String: Returns the canonical m sys object type predicate value used by downstream policy/dispatch logic; vbNullString indicates that no managed value applies where the API permits absence.
Public Function MSysObjectTypePredicate(ByVal objectType As AcObjectType, ByVal fieldExpression As String) As String
    'Builds the MSys type predicate for an Access object type; tables intentionally match local, ODBC, and linked table rows.
    Dim Definition As Variant
    Dim acceptedTypes As Variant
    Dim acceptedType As Variant
    Dim typeList As String
    If ObjectTypeDefinitionForAccessType(objectType, Definition) = False Then Err.Raise vbObjectError + 6899, "MSysObjectTypePredicate", "Unsupported Access object type for MSys lookup: " & CStr(CLng(objectType))
    fieldExpression = Trim$(fieldExpression)
    If Len(fieldExpression) = 0 Then Err.Raise vbObjectError + 6900, "MSysObjectTypePredicate", "The MSys type field expression is blank."
    acceptedTypes = Definition(2)
    For Each acceptedType In acceptedTypes
        If Len(typeList) > 0 Then typeList = typeList & ","
        typeList = typeList & CStr(CLng(acceptedType))
    Next acceptedType
    If UBound(acceptedTypes) = LBound(acceptedTypes) Then
        MSysObjectTypePredicate = fieldExpression & "=" & typeList
    Else
        MSysObjectTypePredicate = fieldExpression & " IN (" & typeList & ")"
    End If
End Function

' Purpose: Builds the configured physical Navigation Pane group name from its logical group name.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   logicalGroupName (Input; String): Stable logical group identity used independently of the physical/display group name.
'   systemMacroGroup (Input Optional; Boolean): Boolean state describing whether system macro group applies to this invocation; True selects the named condition/path and False selects its complementary/default
'   path. Default: False.
' Returns: String: Returns the physical/logical name resolved for build navigation pane group name under the module's naming/definition rules; vbNullString indicates that no applicable name can be resolved
'   where supported.
Public Function BuildNavigationPaneGroupName(ByVal logicalGroupName As String, Optional ByVal systemMacroGroup As Boolean = False) As String
    'Builds the configured physical Navigation Pane group name from its logical group name.
    Dim prefixText As String
    Dim postfixText As String
    NavigationPaneGroupAffixes systemMacroGroup, prefixText, postfixText
    BuildNavigationPaneGroupName = prefixText & logicalGroupName & postfixText
End Function

Private Function NamingSnapshotManagedNameKey(ByVal namingCategory As String, ByVal logicalObjectName As String) As String
    'Builds the case-insensitive cache key for a managed naming-category/logical-name pair.
    NamingSnapshotManagedNameKey = UCase$(Trim$(namingCategory)) & "|" & UCase$(logicalObjectName)
End Function

' Purpose: Resolves bootstrap settings table name from the supplied/current ADC state, using `ExistingBootstrapTableName` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: String: Returns the resolved name produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function BootstrapSettingsTableName() As String
    'Returns the current physical Settings table name used during bootstrap operations.
    BootstrapSettingsTableName = ExistingBootstrapTableName(BOOTSTRAP_SETTINGS_TABLE_PROPERTY, DEFAULT_SETTINGS_TABLE_NAME)
End Function

Private Function NamingSettingText(ByVal SettingKey As String, ByVal defaultText As String) As String
    'Reads a naming-policy text setting and falls back to the supplied default when unavailable.
    On Error GoTo UseDefault
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim tableName As String
    If mNamingSnapshotActive And mNamingPolicyReady Then
        NamingSettingText = NamingPolicyText(SettingKey, defaultText)
        Exit Function
    End If
    tableName = BootstrapSettingsTableName()
    Set db = currentDb
    If DaoTableExists(db, tableName) = False Then GoTo UseDefault
    mNamingDatabaseLookupCount = mNamingDatabaseLookupCount + 1
    Set rs = db.OpenRecordset("SELECT TOP 1 [" & FIELD_SETTING_VALUE & "] FROM [" & Replace(tableName, "]", "]]") & "] WHERE [" & FIELD_SETTING_KEY & "]='" & Replace(SettingKey, "'", "''") & "';", dbOpenSnapshot)
    If rs.EOF Then GoTo UseDefault
    If IsNull(rs.fields(FIELD_SETTING_VALUE).value) Then NamingSettingText = vbNullString Else NamingSettingText = CStr(rs.fields(FIELD_SETTING_VALUE).value)
    rs.Close
    Exit Function
UseDefault:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    NamingSettingText = defaultText
End Function

Private Function ReadIdentitySetting(ByVal SettingKey As String, ByVal fallbackText As String) As String
    'Reads an identity setting from the current database and falls back when it is unavailable.
    On Error GoTo UseFallback
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Set db = currentDb
    If DaoTableExists(db, BootstrapSettingsTableName()) = False Then GoTo UseFallback
    sqlText = "SELECT TOP 1 [" & FIELD_SETTING_VALUE & "] FROM [" & Replace(BootstrapSettingsTableName(), "]", "]]") & "] WHERE [" & FIELD_SETTING_KEY & "]='" & Replace(SettingKey, "'", "''") & "'"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    If rs.EOF = False Then ReadIdentitySetting = Trim$(Nz(rs.fields(FIELD_SETTING_VALUE), vbNullString))
    rs.Close
    If Len(ReadIdentitySetting) = 0 Then ReadIdentitySetting = fallbackText
    Exit Function
UseFallback:
    ReadIdentitySetting = fallbackText
End Function

' Purpose: Resolves diagnostic date text from the supplied/current ADC state, using `IsDate` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   valueData (Input; Variant): Variant input representing value data; the routine accepts the caller-provided value in the form required by its existing validation and conversion logic.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function DiagnosticDateText(ByVal valueData As Variant) As String
    'Intended use: Formats ADC diagnostic date/value fields consistently for comparison and validation messages.
    If IsNull(valueData) Or IsEmpty(valueData) Then
        DiagnosticDateText = "(Null)"
    ElseIf IsDate(valueData) Then
        DiagnosticDateText = Format$(CDate(valueData), "yyyy-mm-dd hh:nn:ss")
    Else
        DiagnosticDateText = CStr(valueData)
    End If
End Function

Private Function NamingPolicyText(ByVal SettingKey As String, ByVal defaultText As String) As String
    'Returns the effective naming-policy text for a setting key, using snapshot state or the supplied default.
    Select Case UCase$(SettingKey)
        Case UCase$(SETTING_TABLE_OBJECT_PREFIX): NamingPolicyText = mNamingTableObjectPrefix
        Case UCase$(SETTING_TABLE_OBJECT_POSTFIX): NamingPolicyText = mNamingTableObjectPostfix
        Case UCase$(SETTING_QUERY_OBJECT_PREFIX): NamingPolicyText = mNamingQueryObjectPrefix
        Case UCase$(SETTING_QUERY_OBJECT_POSTFIX): NamingPolicyText = mNamingQueryObjectPostfix
        Case UCase$(SETTING_GENERATED_MACRO_OBJECT_PREFIX): NamingPolicyText = mNamingGeneratedMacroObjectPrefix
        Case UCase$(SETTING_GENERATED_MACRO_OBJECT_POSTFIX): NamingPolicyText = mNamingGeneratedMacroObjectPostfix
        Case UCase$(SETTING_SYSTEM_MACRO_OBJECT_PREFIX): NamingPolicyText = mNamingSystemMacroObjectPrefix
        Case UCase$(SETTING_SYSTEM_MACRO_OBJECT_POSTFIX): NamingPolicyText = mNamingSystemMacroObjectPostfix
        Case UCase$(SETTING_CUSTOM_FORM_OBJECT_PREFIX): NamingPolicyText = mNamingCustomFormObjectPrefix
        Case UCase$(SETTING_CUSTOM_FORM_OBJECT_POSTFIX): NamingPolicyText = mNamingCustomFormObjectPostfix
        Case UCase$(SETTING_NAVIGATION_PANE_GROUP_PREFIX): NamingPolicyText = mNamingNavigationPaneGroupPrefix
        Case UCase$(SETTING_NAVIGATION_PANE_GROUP_POSTFIX): NamingPolicyText = mNamingNavigationPaneGroupPostfix
        Case UCase$(SETTING_SYSTEM_MACRO_NAVIGATION_PANE_GROUP_PREFIX): NamingPolicyText = mNamingSystemMacroNavigationPaneGroupPrefix
        Case UCase$(SETTING_SYSTEM_MACRO_NAVIGATION_PANE_GROUP_POSTFIX): NamingPolicyText = mNamingSystemMacroNavigationPaneGroupPostfix
        Case Else: NamingPolicyText = defaultText
    End Select
End Function

' Purpose: Resolves msg box result text from the supplied/current ADC state, using `Unknown` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   resultValue (Input; VbMsgBoxResult): Receives the value returned to Access/Office through the callback contract.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function MsgBoxResultText(ByVal resultValue As VbMsgBoxResult) As String
    'Returns a stable diagnostic name for a VBA UtilityMsgBox result value.
    Select Case resultValue
        Case vbOK
            MsgBoxResultText = "vbOK"
        Case vbCancel
            MsgBoxResultText = "vbCancel"
        Case vbAbort
            MsgBoxResultText = "vbAbort"
        Case vbRetry
            MsgBoxResultText = "vbRetry"
        Case vbIgnore
            MsgBoxResultText = "vbIgnore"
        Case vbYes
            MsgBoxResultText = "vbYes"
        Case vbNo
            MsgBoxResultText = "vbNo"
        Case Else
            MsgBoxResultText = "Unknown(" & CStr(resultValue) & ")"
    End Select
End Function

' Purpose: Returns the standard Access Development Console completion wording for clean or warning-bearing success results.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   cleanSuccessText (Input; String): Input clean success text for completion success text; is tested by If InStr(1, cleanSuccessText, "completed with warning", vbTextCompare) > 0; is assigned to
'     `CompletionSuccessText`.
'   warningDetected (Input; Boolean): Indicates whether warnings have already been detected for the operation; the routine preserves/updates warning handling based on this state.
' Returns: String: Returns the resolved completion success text text in the canonical representation consumed by this module's caller; vbNullString represents no applicable value where the existing contract
'   permits it.
Public Function CompletionSuccessText(ByVal cleanSuccessText As String, ByVal warningDetected As Boolean) As String
    'Returns the standard Access Development Console completion wording for clean or warning-bearing success results.
    If warningDetected Then
        If InStr(1, cleanSuccessText, "completed with warning", vbTextCompare) > 0 Then
            CompletionSuccessText = cleanSuccessText
        Else
            CompletionSuccessText = Replace(cleanSuccessText, " completed successfully", " completed with warning(s)", 1, 1, vbTextCompare)
            If StrComp(CompletionSuccessText, cleanSuccessText, vbTextCompare) = 0 Then CompletionSuccessText = cleanSuccessText & " Completed with warning(s)."
        End If
    Else
        CompletionSuccessText = cleanSuccessText
    End If
End Function

Private Function DaoDiagnosticContextValue(ByVal targetDiagnosticContext As String, ByVal valueName As String) As String
    'Extracts a named value from the line-oriented DAO target diagnostic context.
    Dim contextLines() As String
    Dim lineIndex As Long
    Dim prefixText As String
    prefixText = valueName & ":"
    contextLines = Split(Replace(targetDiagnosticContext, vbCr, vbNullString), vbLf)
    For lineIndex = LBound(contextLines) To UBound(contextLines)
        If StrComp(Left$(Trim$(contextLines(lineIndex)), Len(prefixText)), prefixText, vbTextCompare) = 0 Then
            DaoDiagnosticContextValue = Trim$(Mid$(Trim$(contextLines(lineIndex)), Len(prefixText) + 1))
            Exit Function
        End If
    Next lineIndex
End Function

Private Function DaoTargetContextText(ByVal targetDiagnosticContext As String) As String
    'Returns a normalized DAO target diagnostic context suitable for diagnostic messages.
    targetDiagnosticContext = Trim$(targetDiagnosticContext)
    If Len(targetDiagnosticContext) > 0 Then DaoTargetContextText = targetDiagnosticContext & vbCrLf
End Function

' Purpose: Resolves fingerprint from payload from the supplied/current ADC state, using `StableTextHashHex` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   payloadText (Input; String): Input payload text for fingerprint from payload; is passed to `StableTextHashHex`.
' Returns: String: Returns the fingerprint from payload text value. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function FingerprintFromPayload(ByVal payloadText As String) As String
    'Returns the canonical ADC fingerprint for a completed payload as stable-hash text followed by the payload length.
    FingerprintFromPayload = StableTextHashHex(payloadText) & ":" & CStr(Len(payloadText))
End Function

' Purpose: Returns one DAO object property as fingerprint text, using an empty string for Null and <missing> for any read failure.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   target (Input; Object): Object instance whose property/metadata value is being read, ensured, or compared by the helper.
'   propertyName (Input; String): Exact Access/DAO property name to read, compare, create, update, or fingerprint.
' Returns: String: Returns the deterministic DAO object property fingerprint text token/signature produced from the supplied state, allowing callers to compare exact state without repeating the underlying
'   scan/build work.
Public Function DaoObjectPropertyFingerprintText(ByVal target As Object, ByVal propertyName As String) As String
    'Returns one DAO object property as fingerprint text, using an empty string for Null and <missing> for any read failure.
    On Error GoTo MissingProperty
    DaoObjectPropertyFingerprintText = CStr(Nz(target.properties(propertyName).value, vbNullString))
    Exit Function
MissingProperty:
    Err.Clear
    DaoObjectPropertyFingerprintText = "<missing>"
End Function

' Purpose: Validates that a generated form's physical name round-trips through the canonical generated-form naming contract before preparation or preflight continues.
' Public role: Inter-module public API. Generated-form planning/build code calls this routine directly before invoking shared form-preparation behavior.
' Parameters:
'   physicalFormName (Input; String): Canonical physical form name to validate against the generated-form definition map.
' Returns: None (Sub). Validation succeeds by returning normally; invalid or noncanonical identities raise the existing generated-form preparation errors.
Public Sub ValidateGeneratedFormPreparationIdentity(ByVal physicalFormName As String)
    Dim logicalName As String
    Dim canonicalPhysicalName As String
    Dim ChatGPTUIForm As Boolean
    If Len(physicalFormName) = 0 Then Err.Raise vbObjectError + 7860, "ValidateGeneratedFormPreparationIdentity", "The generated-form physical name is blank."
    If CanonicalGeneratedFormObjectName(physicalFormName, True, logicalName, ChatGPTUIForm) = False Then Err.Raise vbObjectError + 7861, "ValidateGeneratedFormPreparationIdentity", "The generated-form physical name does not resolve to a canonical generated-form definition." & vbCrLf & "Form: " & physicalFormName
    If CanonicalGeneratedFormObjectName(logicalName, False, canonicalPhysicalName, ChatGPTUIForm) = False Then Err.Raise vbObjectError + 7862, "ValidateGeneratedFormPreparationIdentity", "The generated-form logical identity does not resolve back to its canonical physical name." & vbCrLf & "Form: " & physicalFormName & vbCrLf & "Logical ID: " & logicalName
    If StrComp(canonicalPhysicalName, physicalFormName, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7863, "ValidateGeneratedFormPreparationIdentity", "The generated-form physical name does not exactly match its authoritative resolved physical identity." & vbCrLf & "Form: " & physicalFormName & vbCrLf & "Logical ID: " & logicalName & vbCrLf & "Authoritative physical name: " & canonicalPhysicalName
End Sub

' Purpose: Preserves the legacy public generated-form preparation-prefix API, including its historical validation error source, while delegating canonical identity validation to the explicit validator.
' Public role: Compatibility entry point for independent callers that may still invoke the historical Function; current ADC code does not use this return-value contract internally.
' Parameters:
'   physicalFormName (Input; String): Physical generated-form name to validate against the canonical generated-form identity map.
' Returns: String: Always vbNullString after successful validation. Known validation failures preserve the historical GeneratedFormPreparationRequiredPrefix Err.Source contract.
Public Function GeneratedFormPreparationRequiredPrefix(ByVal physicalFormName As String) As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim savedErrHelpFile As String
    Dim savedErrHelpContext As Long
    On Error GoTo ValidationError
    ValidateGeneratedFormPreparationIdentity physicalFormName
    GeneratedFormPreparationRequiredPrefix = vbNullString
    Exit Function
ValidationError:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    savedErrHelpFile = Err.HelpFile
    savedErrHelpContext = Err.HelpContext
    If savedErrNumber >= vbObjectError + 7860 And savedErrNumber <= vbObjectError + 7863 Then savedErrSource = "GeneratedFormPreparationRequiredPrefix"
    Err.Raise savedErrNumber, savedErrSource, savedErrDescription, savedErrHelpFile, savedErrHelpContext
End Function

' Purpose: Resolves a logical managed-object name to its configured physical database-object name.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   namingCategory (Input; String): Category value classifying the naming item so the routine can select the correct naming, logging, or processing rules.
'   logicalObjectName (Input; String): Stable managed logical object identity used to resolve the current physical Access object name.
' Returns: String: Returns the physical/logical name resolved for managed object physical name under the module's naming/definition rules; vbNullString indicates that no applicable name can be resolved where
'   supported.
Public Function ManagedObjectPhysicalName(ByVal namingCategory As String, ByVal logicalObjectName As String) As String
    'Resolves a logical managed-object name to its configured physical database-object name.
    Dim currentName As String
    Dim canonicalName As String
    Dim ChatGPTUIForm As Boolean
    Dim chatGPTUIReport As Boolean
    If StrComp(logicalObjectName, Trim$(logicalObjectName), vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 10140, "ManagedObjectPhysicalName", "A managed-object logical identity has noncanonical outer whitespace." & vbCrLf & "LogicalObjectName: [" & logicalObjectName & "]"
    If Len(logicalObjectName) = 0 Then Exit Function
    If StrComp(namingCategory, NAMING_CATEGORY_TABLE, vbTextCompare) = 0 Then
        If StrComp(logicalObjectName, DEFAULT_SETTINGS_TABLE_LOGICAL_ID, vbTextCompare) = 0 Then ManagedObjectPhysicalName = BootstrapSettingsTableName(): Exit Function
        If StrComp(logicalObjectName, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID, vbTextCompare) = 0 Then ManagedObjectPhysicalName = BootstrapManagedObjectsTableName(): Exit Function
    ElseIf StrComp(namingCategory, NAMING_CATEGORY_GENERATED_FORM, vbTextCompare) = 0 Then
        'Generated-form physical identity is canonical and classification-driven. Do not allow catalog overrides or mutable naming affixes to change it.
        If CanonicalGeneratedFormObjectName(logicalObjectName, False, canonicalName, ChatGPTUIForm) = False Then Err.Raise vbObjectError + 7811, "ManagedObjectPhysicalName", "Unknown ADC generated-form logical ID: " & logicalObjectName
        ManagedObjectPhysicalName = canonicalName
        Exit Function
    ElseIf StrComp(namingCategory, NAMING_CATEGORY_REPORT, vbTextCompare) = 0 Then
        'Managed-report physical identity is canonical and classification-driven. Do not allow catalog overrides or mutable naming affixes to change it.
        If CanonicalReportObjectName(logicalObjectName, False, canonicalName, chatGPTUIReport) = False Then Err.Raise vbObjectError + 7812, "ManagedObjectPhysicalName", "Unknown ADC managed-report logical ID: " & logicalObjectName
        ManagedObjectPhysicalName = canonicalName
        Exit Function
    End If
    currentName = ManagedObjectCatalogPhysicalName(namingCategory, logicalObjectName)
    If Len(currentName) > 0 Then ManagedObjectPhysicalName = currentName Else ManagedObjectPhysicalName = TargetManagedObjectPhysicalName(namingCategory, logicalObjectName)
End Function

' Purpose: Resolves the target physical name for a managed object using current naming rules.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   namingCategory (Input; String): Category value classifying the naming item so the routine can select the correct naming, logging, or processing rules.
'   logicalObjectName (Input; String): Stable managed logical object identity used to resolve the current physical Access object name.
' Returns: String: Returns the physical/logical name resolved for target managed object physical name under the module's naming/definition rules; vbNullString indicates that no applicable name can be resolved
'   where supported.
Public Function TargetManagedObjectPhysicalName(ByVal namingCategory As String, ByVal logicalObjectName As String) As String
    'Resolves the target physical name for a managed object using current naming rules.
    Dim prefixText As String
    Dim postfixText As String
    Dim canonicalName As String
    Dim ChatGPTUIForm As Boolean
    Dim chatGPTUIReport As Boolean
    Dim designatorText As String
    Dim separatorPosition As Long
    If StrComp(namingCategory, NAMING_CATEGORY_SYSTEM_MACRO, vbTextCompare) = 0 Then
        If Len(logicalObjectName) = 0 Then Exit Function
        separatorPosition = InStr(1, logicalObjectName, "|", vbBinaryCompare)
        If separatorPosition > 1 Then
            designatorText = Left$(logicalObjectName, separatorPosition - 1)
            If Len(designatorText) > 0 Then
                logicalObjectName = Mid$(logicalObjectName, separatorPosition + 1)
                If IncludeSystemMacroDesignatorInNavigationPaneTitle() Then logicalObjectName = designatorText & " " & logicalObjectName
            End If
        End If
    Else
        If StrComp(logicalObjectName, Trim$(logicalObjectName), vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 10141, "TargetManagedObjectPhysicalName", "A managed-object logical identity has noncanonical outer whitespace." & vbCrLf & "LogicalObjectName: [" & logicalObjectName & "]"
        If Len(logicalObjectName) = 0 Then Exit Function
    End If
    If StrComp(namingCategory, NAMING_CATEGORY_GENERATED_FORM, vbTextCompare) = 0 Then
        If CanonicalGeneratedFormObjectName(logicalObjectName, False, canonicalName, ChatGPTUIForm) = False Then Err.Raise vbObjectError + 7811, "TargetManagedObjectPhysicalName", "Unknown ADC generated-form logical ID: " & logicalObjectName
        TargetManagedObjectPhysicalName = canonicalName
        Exit Function
    ElseIf StrComp(namingCategory, NAMING_CATEGORY_REPORT, vbTextCompare) = 0 Then
        If CanonicalReportObjectName(logicalObjectName, False, canonicalName, chatGPTUIReport) = False Then Err.Raise vbObjectError + 7812, "TargetManagedObjectPhysicalName", "Unknown ADC managed-report logical ID: " & logicalObjectName
        TargetManagedObjectPhysicalName = canonicalName
        Exit Function
    End If
    If StrComp(namingCategory, NAMING_CATEGORY_TABLE, vbTextCompare) = 0 Or StrComp(namingCategory, NAMING_CATEGORY_QUERY, vbTextCompare) = 0 Then
        If CanonicalTableQueryObjectName(namingCategory, logicalObjectName, False, canonicalName) = False Then Err.Raise vbObjectError + 7810, "TargetManagedObjectPhysicalName", "Unknown ADC table/query logical ID: " & logicalObjectName & vbCrLf & "Naming category: " & namingCategory
        If StrComp(namingCategory, NAMING_CATEGORY_TABLE, vbTextCompare) = 0 And StrComp(logicalObjectName, DEFAULT_TERMINOLOGY_TABLE_LOGICAL_ID, vbTextCompare) = 0 Then TargetManagedObjectPhysicalName = canonicalName: Exit Function
        logicalObjectName = canonicalName
    End If
    NamingAffixes namingCategory, prefixText, postfixText
    TargetManagedObjectPhysicalName = prefixText & logicalObjectName & postfixText
End Function

' Purpose: Resolves managed object naming category from the supplied/current ADC state, using `CanonicalGeneratedFormObjectName` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   objectTypeCode (Input; String): Canonical code used to classify or select the object type code behavior.
'   categoryCode (Input; String): Canonical code used to classify or select the category code behavior.
'   physicalObjectName (Input; String): Canonical or physical name identifying the physical object targeted by this operation.
' Returns: String: Returns the managed object naming category text value. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function ManagedObjectNamingCategory(ByVal objectTypeCode As String, ByVal categoryCode As String, ByVal physicalObjectName As String) As String
    'Determines the naming category for a managed object from its type, category, and physical name.
    Dim mappedName As String
    Dim ChatGPTUIForm As Boolean
    Select Case UCase$(Trim$(objectTypeCode))
        Case "TABLE": ManagedObjectNamingCategory = NAMING_CATEGORY_TABLE
        Case "QUERY": ManagedObjectNamingCategory = NAMING_CATEGORY_QUERY
        Case "REPORT": ManagedObjectNamingCategory = NAMING_CATEGORY_REPORT
        Case "MACRO"
            If UCase$(Trim$(categoryCode)) = UCase$(ADCAT_CATEGORY_SYSTEM_MACRO) Then ManagedObjectNamingCategory = NAMING_CATEGORY_SYSTEM_MACRO Else ManagedObjectNamingCategory = NAMING_CATEGORY_GENERATED_MACRO
        Case "FORM"
            If CanonicalGeneratedFormObjectName(physicalObjectName, True, mappedName, ChatGPTUIForm) Then ManagedObjectNamingCategory = NAMING_CATEGORY_GENERATED_FORM
    End Select
End Function

' Purpose: Resolves bootstrap managed objects table name from the supplied/current ADC state, using `ExistingBootstrapTableName` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: String: Returns the resolved name produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function BootstrapManagedObjectsTableName() As String
    'Returns the current physical Managed Objects table name used during bootstrap operations.
    BootstrapManagedObjectsTableName = ExistingBootstrapTableName(BOOTSTRAP_MANAGED_OBJECTS_TABLE_PROPERTY, TargetManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID))
End Function

' Purpose: Resolves status bar progress text from the supplied/current ADC state, using `CoreQualifyProgressStatusText` and `StatusBarHierarchyText` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   statusText (Input; String): Status/progress message supplied for the current operation; the routine forwards it to the applicable UI/progress and trace/reporting path.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function StatusBarProgressText(ByVal statusText As String) As String
    'Applies the universal visible Part X of X contract before status-bar hierarchy formatting.
    On Error GoTo UseOriginal
    Dim qualifiedText As String
    statusText = Trim$(statusText)
    If Len(statusText) = 0 Then Exit Function
    qualifiedText = CoreQualifyProgressStatusText(statusText)
    statusText = qualifiedText
UseOriginal:
    StatusBarProgressText = StatusBarHierarchyText(statusText)
End Function

' Purpose: Resolves status bar hierarchy text from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   statusText (Input; String): Status/progress message supplied for the current operation; the routine forwards it to the applicable UI/progress and trace/reporting path.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function StatusBarHierarchyText(ByVal statusText As String) As String
    'Expands spaced hierarchy separators only for status-bar/progress display text.
    StatusBarHierarchyText = Replace(statusText, " > ", " >>> ", 1, -1, vbBinaryCompare)
End Function

Private Function ExistingBootstrapTableName(ByVal propertyName As String, ByVal fallbackText As String) As String
    'Returns the remembered bootstrap table name for a database property or the supplied fallback.
    On Error GoTo UseFallback
    Dim db As DAO.Database
    Dim candidateName As String
    Set db = currentDb
    candidateName = CStr(db.properties(propertyName).value)
    If Len(candidateName) > 0 Then
        If DaoTableExists(db, candidateName) Then ExistingBootstrapTableName = candidateName: Exit Function
    End If
UseFallback:
    ExistingBootstrapTableName = fallbackText
End Function

Private Function ManagedObjectCatalogPhysicalName(ByVal namingCategory As String, ByVal logicalObjectName As String) As String
    'Resolves a managed catalog object to its current physical name while preserving bootstrap behavior.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim tdf As DAO.tableDef
    Dim tableName As String
    If mNamingSnapshotActive Then
        If Not mNamingSnapshotManagedNames Is Nothing Then
            tableName = NamingSnapshotManagedNameKey(namingCategory, logicalObjectName)
            If mNamingSnapshotManagedNames.Exists(tableName) Then ManagedObjectCatalogPhysicalName = CStr(mNamingSnapshotManagedNames(tableName))
        End If
        Exit Function
    End If
    tableName = BootstrapManagedObjectsTableName()
    Set db = currentDb
    If DaoTableExists(db, tableName) = False Then Exit Function
    Set tdf = db.TableDefs(tableName)
    If DAOFieldExists(tdf, FIELD_LOGICAL_OBJECT_NAME) = False Then Exit Function
    If DAOFieldExists(tdf, FIELD_OBJECT_NAMING_CATEGORY) = False Then Exit Function
    mNamingDatabaseLookupCount = mNamingDatabaseLookupCount + 1
    Set rs = db.OpenRecordset("SELECT " & SQL_FIELD_OBJECT_NAME & " FROM [" & Replace(tableName, "]", "]]") & "] WHERE Nz([" & FIELD_MANAGED_OBJECT_STATUS & "],'" & ADCAT_STATUS_ACTIVE & "')='" & ADCAT_STATUS_ACTIVE & "' AND [" & FIELD_LOGICAL_OBJECT_NAME & "]='" & Replace(logicalObjectName, "'", "''") & "' AND [" & FIELD_OBJECT_NAMING_CATEGORY & "]='" & Replace(namingCategory, "'", "''") & "';", dbOpenSnapshot)
    If rs.EOF = False Then
        ManagedObjectCatalogPhysicalName = CStr(Nz(rs.fields(FIELD_OBJECT_NAME).value, vbNullString))
        rs.MoveNext
        If rs.EOF = False Then Err.Raise vbObjectError + 6802, "ManagedObjectCatalogPhysicalName", "More than one active managed-object row has the same naming category and logical name."
    End If
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
    Err.Raise savedErrNumber, IIf(Len(savedErrSource) = 0, "ManagedObjectCatalogPhysicalName", savedErrSource), savedErrDescription
End Function

Private Function UtilityLongName(Optional ByVal fallbackText As Variant) As String
    'Returns the effective long product name, using the supplied fallback when configuration is unavailable.
    Dim effectiveFallback As String
    If isMissing(fallbackText) Then
        effectiveFallback = DEFAULT_UTILITY_LONG_NAME
    Else
        effectiveFallback = CStr(fallbackText)
    End If
    UtilityLongName = ReadIdentitySetting("UtilityLongName", effectiveFallback)
End Function

Private Function UtilityShortName(Optional ByVal fallbackText As Variant) As String
    'Returns the effective short product name, using the supplied fallback when configuration is unavailable.
    Dim effectiveFallback As String
    If isMissing(fallbackText) Then
        effectiveFallback = DEFAULT_UTILITY_SHORT_NAME
    Else
        effectiveFallback = CStr(fallbackText)
    End If
    UtilityShortName = ReadIdentitySetting("UtilityShortName", effectiveFallback)
End Function

Private Function PopupPrimaryDiagnosticText(ByVal sourceText As String) As String
    'Keeps popup classification scoped to the current error chain. Trace summaries can contain historical words such as conflict, duplicate, mismatch, or failure that do not describe the current error.
    Dim cutPosition As Long
    Dim markerPosition As Long
    Dim markerText As Variant
    cutPosition = Len(sourceText) + 1
    For Each markerText In Array(vbCrLf & vbCrLf & "Diagnostics:", vbLf & vbLf & "Diagnostics:", vbCrLf & vbCrLf & DEFAULT_UTILITY_LONG_NAME & " table/object diagnostics:", vbLf & vbLf & DEFAULT_UTILITY_LONG_NAME & " table/object diagnostics:", vbCrLf & vbCrLf & "Trace steps captured before failure:", vbLf & vbLf & "Trace steps captured before failure:", vbCrLf & vbCrLf & "Recent trace steps:", vbLf & vbLf & "Recent trace steps:", vbCrLf & vbCrLf & "Table creation/maintenance operations:", vbLf & vbLf & "Table creation/maintenance operations:", vbCrLf & vbCrLf & "Process run:", vbLf & vbLf & "Process run:")
        markerPosition = InStr(1, sourceText, CStr(markerText), vbTextCompare)
        If markerPosition > 0 And markerPosition < cutPosition Then cutPosition = markerPosition
    Next markerText
    If cutPosition <= Len(sourceText) Then
        PopupPrimaryDiagnosticText = Trim$(Left$(sourceText, cutPosition - 1))
    Else
        PopupPrimaryDiagnosticText = sourceText
    End If
End Function

Private Function PopupSuccessText(ByVal sourceText As String) As String
    'Builds concise user-facing success text from a technical completion diagnostic.
    Dim normalizedText As String
    Dim firstBreak As Long
    normalizedText = Replace(Replace(sourceText, vbCrLf, vbLf), vbCr, vbLf)
    firstBreak = InStr(1, normalizedText, vbLf & vbLf, vbBinaryCompare)
    If firstBreak > 0 Then normalizedText = Left$(normalizedText, firstBreak - 1)
    normalizedText = Trim$(Replace(normalizedText, vbLf, vbCrLf))
    If Len(normalizedText) = 0 Or Len(normalizedText) > 350 Then normalizedText = DEFAULT_UTILITY_SHORT_NAME & " completed the requested operation successfully."
    PopupSuccessText = normalizedText
End Function

Private Function PopupOperationText(ByVal sourceText As String, ByVal TitleText As String) As String
    'Derives a concise operation description for a user-facing popup from diagnostic text and title context.
    Dim partText As String
    Dim cleanTitle As String
    partText = ExtractSetupPartText(sourceText)
    If Len(partText) > 0 Then
        PopupOperationText = partText
        Exit Function
    End If
    cleanTitle = Trim$(TitleText)
    cleanTitle = Replace(cleanTitle, DEFAULT_UTILITY_LONG_NAME, vbNullString, 1, -1, vbTextCompare)
    cleanTitle = Replace(cleanTitle, " - Error Detail", vbNullString, 1, -1, vbTextCompare)
    cleanTitle = Replace(cleanTitle, " Failed", vbNullString, 1, -1, vbTextCompare)
    cleanTitle = Replace(cleanTitle, " Error", vbNullString, 1, -1, vbTextCompare)
    cleanTitle = Replace(cleanTitle, ":", " ")
    cleanTitle = Trim$(cleanTitle)
    Do While InStr(cleanTitle, "  ") > 0
        cleanTitle = Replace(cleanTitle, "  ", " ")
    Loop
    If InStr(1, cleanTitle, "Module Update Guard", vbTextCompare) > 0 Then cleanTitle = "Module Guard verification"
    If Len(cleanTitle) = 0 Then cleanTitle = "the requested operation"
    PopupOperationText = cleanTitle
End Function

Private Function PopupReasonText(ByVal sourceText As String) As String
    'Extracts or derives the reason portion of a technical diagnostic for user-facing popup text.
    Dim tableName As String
    Dim fieldName As String
    Dim ModuleName As String
    Dim queryName As String
    Dim daoErrorText As String
    Dim failureCheck As String
    tableName = FirstDiagnosticValue(sourceText, "Table")
    fieldName = FirstDiagnosticValue(sourceText, "Field")
    ModuleName = FirstDiagnosticValue(sourceText, "First failing module")
    If Len(ModuleName) = 0 Then ModuleName = FirstDiagnosticValue(sourceText, "Module")
    If InStr(1, sourceText, "UNMAPPED FIELD DESCRIPTION:", vbTextCompare) > 0 Or InStr(1, sourceText, "no authoritative field description", vbTextCompare) > 0 Then
        If Len(tableName) > 0 And Len(fieldName) > 0 Then PopupReasonText = tableName & "." & fieldName & " has no authoritative field description." Else PopupReasonText = "An " & DEFAULT_UTILITY_SHORT_NAME & "-managed field has no authoritative field description."
        Exit Function
    End If
    If InStr(1, sourceText, "Current Version ID mismatch", vbTextCompare) > 0 Or InStr(1, sourceText, "HeaderMismatch", vbTextCompare) > 0 Or InStr(1, sourceText, "approved build", vbTextCompare) > 0 And InStr(1, sourceText, "mismatch", vbTextCompare) > 0 Then
        PopupReasonText = "The installed " & DEFAULT_UTILITY_SHORT_NAME & " module set does not match the approved build."
        If Len(ModuleName) > 0 Then PopupReasonText = PopupReasonText & " The first affected module is " & ModuleName & "."
        Exit Function
    End If
    If InStr(1, sourceText, "missing a required ADC module", vbTextCompare) > 0 Or InStr(1, sourceText, "required ADC module is not installed", vbTextCompare) > 0 Then
        PopupReasonText = "A required " & DEFAULT_UTILITY_SHORT_NAME & " module is not installed."
        If Len(ModuleName) > 0 Then PopupReasonText = PopupReasonText & " Missing module: " & ModuleName & "."
        Exit Function
    End If
    If InStr(1, sourceText, "Property value is too large", vbTextCompare) > 0 Or InStr(1, sourceText, "Err.Number: 3309", vbTextCompare) > 0 Then
        If Len(tableName) > 0 And Len(fieldName) > 0 Then PopupReasonText = "Access could not store the required description for " & tableName & "." & fieldName & "." Else PopupReasonText = "Access could not store a required managed-object description."
        Exit Function
    End If
    If InStr(1, sourceText, "active managed query could not expose its current physical definition", vbTextCompare) > 0 Then
        queryName = FirstDiagnosticValue(sourceText, "Query")
        daoErrorText = FirstDiagnosticValue(sourceText, "DAO error")
        PopupReasonText = "Access could not read the current physical definition of the managed saved query" & IIf(Len(queryName) = 0, ".", " [" & queryName & "].")
        If Len(daoErrorText) > 0 Then
            PopupReasonText = PopupReasonText & " DAO reported " & daoErrorText
            If Right$(PopupReasonText, 1) <> "." Then PopupReasonText = PopupReasonText & "."
        End If
        Exit Function
    End If
    If InStr(1, sourceText, "active managed table does not exist physically", vbTextCompare) > 0 Then
        PopupReasonText = "A managed table required by the current operation is not physically available in the database."
        Exit Function
    End If
    If PopupTextIndicatesCurrentConflict(sourceText) Then
        PopupReasonText = MeaningfulFailureLine(sourceText)
        If Len(PopupReasonText) = 0 Then PopupReasonText = DEFAULT_UTILITY_SHORT_NAME & " found a blocking conflict in the current managed definitions."
        Exit Function
    End If
    failureCheck = FirstDiagnosticValue(sourceText, "Failed check")
    If Len(failureCheck) > 0 Then
        PopupReasonText = "The required check failed: " & FirstSentence(failureCheck, 260)
        Exit Function
    End If
    PopupReasonText = MeaningfulFailureLine(sourceText)
    If Len(PopupReasonText) = 0 Then PopupReasonText = DEFAULT_UTILITY_SHORT_NAME & " encountered a condition that requires correction before the operation can continue."
End Function

Private Function PopupAffectedText(ByVal sourceText As String) As String
    'Extracts or derives the affected-object portion of a technical diagnostic for user-facing popup text.
    Dim tableName As String
    Dim fieldName As String
    Dim ModuleName As String
    Dim objectName As String
    tableName = FirstDiagnosticValue(sourceText, "Table")
    fieldName = FirstDiagnosticValue(sourceText, "Field")
    ModuleName = FirstDiagnosticValue(sourceText, "First failing module")
    If Len(ModuleName) = 0 Then ModuleName = FirstDiagnosticValue(sourceText, "Module")
    objectName = FirstDiagnosticValue(sourceText, "Object")
    If Len(tableName) > 0 And Len(fieldName) > 0 Then
        PopupAffectedText = "Affected field: " & tableName & "." & fieldName & "."
    ElseIf Len(ModuleName) > 0 Then
        PopupAffectedText = "Affected module: " & ModuleName & "."
    ElseIf Len(objectName) > 0 Then
        PopupAffectedText = "Affected object: " & objectName & "."
    ElseIf Len(tableName) > 0 Then
        PopupAffectedText = "Affected table: " & tableName & "."
    End If
End Function

Private Function PopupStateText(ByVal sourceText As String, ByVal useDefaultStoppedState As Boolean) As String
    'Builds the user-facing state description associated with a technical diagnostic.
    Dim rollbackFailureText As String
    Dim rollbackFailureCount As Long
    rollbackFailureText = FirstDiagnosticValue(sourceText, "Rollback operations failed")
    If Len(rollbackFailureText) > 0 Then rollbackFailureCount = CLng(Val(rollbackFailureText))
    If InStr(1, sourceText, "stopped before setup changed anything", vbTextCompare) > 0 Or InStr(1, sourceText, "No changes were made", vbTextCompare) > 0 Or InStr(1, sourceText, "Manifest write started: False", vbTextCompare) > 0 Then
        PopupStateText = DEFAULT_UTILITY_SHORT_NAME & " stopped before making setup changes."
    ElseIf rollbackFailureCount > 0 Or InStr(1, sourceText, "rollback failed", vbTextCompare) > 0 Or InStr(1, sourceText, "rollback partially failed", vbTextCompare) > 0 Or InStr(1, sourceText, "could not roll back", vbTextCompare) > 0 Then
        PopupStateText = DEFAULT_UTILITY_SHORT_NAME & " stopped, but rollback did not complete. Some changes may require manual review."
    ElseIf InStr(1, sourceText, "rollback completed successfully", vbTextCompare) > 0 Or InStr(1, sourceText, "rollback succeeded", vbTextCompare) > 0 Or InStr(1, sourceText, "Changes were rolled back", vbTextCompare) > 0 Or InStr(1, sourceText, "Setup/update rollback completed.", vbTextCompare) > 0 Then
        PopupStateText = DEFAULT_UTILITY_SHORT_NAME & " stopped and rolled back the incomplete changes."
    ElseIf InStr(1, sourceText, "post-commit work partially applied", vbTextCompare) > 0 Or InStr(1, sourceText, "committed before", vbTextCompare) > 0 Then
        PopupStateText = DEFAULT_UTILITY_SHORT_NAME & " stopped. Some earlier setup changes may remain and will be rechecked on the next run."
    ElseIf useDefaultStoppedState Then
        PopupStateText = DEFAULT_UTILITY_SHORT_NAME & " stopped before continuing."
    End If
End Function

Private Function PopupActionText(ByVal sourceText As String, ByVal warningMode As Boolean) As String
    'Builds the user-facing next-action description associated with a technical diagnostic.
    Dim explicitAction As String
    explicitAction = FirstDiagnosticValue(sourceText, "Next action")
    If Len(explicitAction) = 0 Then explicitAction = FirstDiagnosticValue(sourceText, "Corrective action")
    If Len(explicitAction) > 0 Then
        PopupActionText = FirstSentence(explicitAction, 260)
    ElseIf InStr(1, sourceText, "UNMAPPED FIELD DESCRIPTION:", vbTextCompare) > 0 Or InStr(1, sourceText, "no authoritative field description", vbTextCompare) > 0 Then
        PopupActionText = "Add the missing authoritative field description, then rerun setup."
    ElseIf InStr(1, sourceText, "Version ID mismatch", vbTextCompare) > 0 Or InStr(1, sourceText, "HeaderMismatch", vbTextCompare) > 0 Or InStr(1, sourceText, "approved build", vbTextCompare) > 0 And InStr(1, sourceText, "mismatch", vbTextCompare) > 0 Then
        PopupActionText = "If the installed ADC modules are one complete reviewed coordinated build, rerun Setup and use the Module Guard registration choice; otherwise replace them with one complete approved build."
    ElseIf InStr(1, sourceText, "missing a required ADC module", vbTextCompare) > 0 Or InStr(1, sourceText, "required ADC module is not installed", vbTextCompare) > 0 Then
        PopupActionText = "Install the required " & DEFAULT_UTILITY_SHORT_NAME & " component, then rerun the operation."
    ElseIf InStr(1, sourceText, "Property value is too large", vbTextCompare) > 0 Or InStr(1, sourceText, "Err.Number: 3309", vbTextCompare) > 0 Then
        PopupActionText = "Review the detailed property-repair entry, correct the metadata condition, and rerun setup."
    ElseIf InStr(1, sourceText, "active managed query could not expose its current physical definition", vbTextCompare) > 0 Then
        PopupActionText = "Review the detailed DAO query-definition error, make sure the affected saved query is present and readable, then rerun setup."
    ElseIf InStr(1, sourceText, "active managed table does not exist physically", vbTextCompare) > 0 Then
        PopupActionText = "Restore or recreate the reported managed table, then rerun setup."
    ElseIf PopupTextIndicatesCurrentConflict(sourceText) Then
        PopupActionText = "Resolve the reported current conflict before rerunning setup."
    ElseIf warningMode Then
        PopupActionText = "Review the warning and the detailed log before continuing if the issue affects your work."
    Else
        PopupActionText = "Review the detailed log, correct the reported condition, and rerun the operation."
    End If
End Function

' Purpose: Adds the standard Access Development Console popup suppression instructions to a prompt.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   promptText (Input; String): Prompt text text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   resultPrompt (Input; Boolean): True requests the routine’s result/confirmation prompt path; False suppresses that prompt and returns/continues through the non-prompt path.
' Returns: String: Returns the resolved suppressible popup prompt text in the canonical representation consumed by this module's caller; vbNullString represents no applicable value where the existing contract
'   permits it.
Public Function SuppressiblePopupPrompt(ByVal promptText As String, ByVal resultPrompt As Boolean) As String
    'Adds the standard Access Development Console popup suppression instructions to a prompt.
    Dim noteText As String
    If resultPrompt Then
        noteText = "Popup controls:" & vbCrLf & "Use the normal action buttons to answer this prompt." & vbCrLf & "Cancel suppresses further " & DEFAULT_UTILITY_LONG_NAME & " popups for this " & DEFAULT_UTILITY_LONG_NAME & " run and uses the safe cancel/no result for this prompt." & vbCrLf & "Status bar, Immediate Window, ADC_TraceLog, and physical trace log reporting continue."
    Else
        noteText = "Popup controls:" & vbCrLf & "OK closes this message." & vbCrLf & "Cancel suppresses further " & DEFAULT_UTILITY_LONG_NAME & " popups for this " & DEFAULT_UTILITY_LONG_NAME & " run." & vbCrLf & "Status bar, Immediate Window, ADC_TraceLog, and physical trace log reporting continue."
    End If
    SuppressiblePopupPrompt = promptText & vbCrLf & vbCrLf & noteText
End Function

' Purpose: Returns completion popup text with the standard warning-status prefix when the caller reports warning-bearing success.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   messageText (Input; String): Message text text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   warningDetected (Input; Boolean): Indicates whether warnings have already been detected for the operation; the routine preserves/updates warning handling based on this state.
' Returns: String: Returns formatted completion popup message text intended for tracing, diagnostics, or user-facing reporting; vbNullString is returned when there is nothing applicable to report.
Public Function CompletionPopupMessage(ByVal messageText As String, ByVal warningDetected As Boolean) As String
    'Returns completion popup text with the standard warning-status prefix when the caller reports warning-bearing success.
    If warningDetected Then
        If InStr(1, messageText, "Completion status: completed with warning", vbTextCompare) > 0 Then
            CompletionPopupMessage = messageText
        Else
            CompletionPopupMessage = "Completion status: completed with warning(s)." & vbCrLf & vbCrLf & messageText
        End If
    Else
        CompletionPopupMessage = messageText
    End If
End Function

' Purpose: Reads one DAO database text property, treats only property-not-found errors as absent, and raises other read failures.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   db (Input; DAO.Database): DAO.Database instance whose schema, data, queries, or properties the routine reads or modifies.
'   propertyName (Input; String): Exact Access/DAO property name to read, compare, create, update, or fingerprint.
'   propertyExists (Output; Boolean): Set True when the procedure determines that property exists applies; set False when that condition is not met.
' Returns: String: Returns the resolved database text property value text in the canonical representation consumed by this module's caller; vbNullString represents no applicable value where the existing
'   contract permits it.
Public Function DatabaseTextPropertyValue(ByVal db As DAO.Database, ByVal propertyName As String, ByRef propertyExists As Boolean) As String
    'Reads one DAO database text property, treats only property-not-found errors as absent, and raises other read failures.
    On Error GoTo ErrHandler
    Dim propertyValue As Variant
    propertyExists = False
    propertyValue = db.properties(propertyName).value
    propertyExists = True
    If IsNull(propertyValue) = False Then DatabaseTextPropertyValue = CStr(propertyValue)
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If savedErrNumber = 3270 Or savedErrNumber = 3265 Then
        Err.Clear
        propertyExists = False
        DatabaseTextPropertyValue = vbNullString
        Exit Function
    End If
    Err.Raise savedErrNumber, "DatabaseTextPropertyValue", "Could not read a database text property." & vbCrLf & "Property: " & propertyName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Function
' Purpose: Returns True when the supplied owner-database role resolves to a supported canonical role.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   ConfigOwnerDatabaseRole (Input; String): Identifies the database role whose configuration is authoritative for this operation, such as the CodeDb or HostDb role defined by ADC configuration.
' Returns: Boolean: True when valid config owner database role is true for the supplied/current context; False otherwise.
Public Function IsValidConfigOwnerDatabaseRole(ByVal ConfigOwnerDatabaseRole As String) As Boolean
    'Returns True when the supplied owner-database role resolves to a supported canonical role.
    IsValidConfigOwnerDatabaseRole = (Len(NormalizedConfigOwnerDatabaseRole(ConfigOwnerDatabaseRole)) > 0)
End Function

' Purpose: Returns True when two query SQL definitions normalize to the same DAO/ACE comparison form.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   db (Input; DAO.Database): DAO.Database instance whose schema, data, queries, or properties the routine reads or modifies.
'   existingSql (Input; String): SQL text or SQL-derived value used to define, validate, compare, or execute the existing SQL operation.
'   requestedSql (Input; String): SQL text or SQL-derived value used to define, validate, compare, or execute the requested SQL operation.
' Returns: Boolean: True when two query SQL definitions normalize to the same DAO/ACE comparison form. False otherwise.
Public Function QuerySqlEquivalent(ByVal db As DAO.Database, ByVal existingSql As String, ByVal requestedSql As String) As Boolean
    'Returns True when two query SQL definitions normalize to the same DAO/ACE comparison form.
    Dim existingKey As String
    Dim requestedKey As String
    existingKey = FormattingInsensitiveQuerySqlKey(existingSql)
    requestedKey = FormattingInsensitiveQuerySqlKey(requestedSql)
    If StrComp(existingKey, requestedKey, vbBinaryCompare) = 0 Then
        QuerySqlEquivalent = True
        Exit Function
    End If
    QuerySqlEquivalent = (StrComp(NormalizedQuerySqlForComparison(db, existingSql), NormalizedQuerySqlForComparison(db, requestedSql), vbBinaryCompare) = 0)
End Function

Private Function ObjectTypeDefinitionForCode(ByVal objectTypeCode As String, ByRef Definition As Variant) As Boolean
    'Finds the canonical object-type definition associated with an ADC object-type code.
    Dim rowData As Variant
    For Each rowData In ObjectTypeDefinitionRows()
        If StrComp(Trim$(objectTypeCode), CStr(rowData(0)), vbTextCompare) = 0 Then
            Definition = rowData
            ObjectTypeDefinitionForCode = True
            Exit Function
        End If
    Next rowData
End Function

' Purpose: Returns True when a Navigation Pane group name already has the configured prefix and postfix.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   groupName (Input; String): Name selecting the logical Ribbon, form, catalog, or definition group affected by the operation.
'   systemMacroGroup (Input Optional; Boolean): Boolean state describing whether system macro group applies to this invocation; True selects the named condition/path and False selects its complementary/default
'   path. Default: False.
' Returns: Boolean: True when a Navigation Pane group name already has the configured prefix and postfix. False otherwise.
Public Function NavigationPaneGroupUsesConfiguredAffixes(ByVal groupName As String, Optional ByVal systemMacroGroup As Boolean = False) As Boolean
    'Returns True when a Navigation Pane group name already has the configured prefix and postfix.
    Dim prefixText As String
    Dim postfixText As String
    NavigationPaneGroupAffixes systemMacroGroup, prefixText, postfixText
    If Len(prefixText) > 0 Then
        If StrComp(Left$(groupName, Len(prefixText)), prefixText, vbBinaryCompare) <> 0 Then Exit Function
    End If
    If Len(postfixText) > 0 Then
        If Len(groupName) < Len(postfixText) Then Exit Function
        If StrComp(Right$(groupName, Len(postfixText)), postfixText, vbBinaryCompare) <> 0 Then Exit Function
    End If
    NavigationPaneGroupUsesConfiguredAffixes = (Len(prefixText) > 0 Or Len(postfixText) > 0)
End Function

Private Function NamingSettingBoolean(ByVal SettingKey As String, ByVal DefaultValue As Boolean) As Boolean
    'Reads a naming-policy Boolean setting and falls back to the supplied default when unavailable.
    Dim valueText As String
    valueText = Trim$(NamingSettingText(SettingKey, IIf(DefaultValue, "True", "False")))
    NamingSettingBoolean = BooleanTextValue(valueText, DefaultValue)
End Function

' Purpose: Returns dictionary keys sorted case-insensitively; Empty is returned for Nothing or an empty dictionary.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   values (Input; Object): Caller-supplied value collection used as the source set for this operation; when passed ByRef the routine may append/update items according to its documented behavior.
' Returns: Variant: Returns a Variant array containing the dictionary keys in the routine’s defined sorted order; an empty result is returned when there are no keys to emit.
Public Function SortedDictionaryKeys(ByVal values As Object) As Variant
    'Returns dictionary keys sorted case-insensitively; Empty is returned for Nothing or an empty dictionary.
    Dim keys As Variant
    Dim i As Long
    Dim j As Long
    Dim swapValue As Variant
    If values Is Nothing Then Exit Function
    If values.count = 0 Then Exit Function
    keys = values.keys
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If StrComp(CStr(keys(i)), CStr(keys(j)), vbTextCompare) > 0 Then
                swapValue = keys(i)
                keys(i) = keys(j)
                keys(j) = swapValue
            End If
        Next j
    Next i
    SortedDictionaryKeys = keys
End Function

' Purpose: Returns True when the supplied database contains the requested Access object by type and name.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   db (Input; DAO.Database): DAO.Database instance whose schema, data, queries, or properties the routine reads or modifies.
'   objectType (Input; AcObjectType): Object-type value identifying the Access/VBA object class targeted by the operation.
'   objectName (Input; String): Physical Access object name identifying the database object targeted by the operation.
' Returns: Boolean: True when the supplied database contains the requested Access object by type and name. False otherwise.
Public Function DatabaseAccessObjectExistsByName(ByVal db As DAO.Database, ByVal objectType As AcObjectType, ByVal objectName As String) As Boolean
    'Returns True when the supplied database contains the requested Access object by type and name.
    On Error GoTo ErrHandler
    If Len(objectName) = 0 Then Exit Function
    Select Case objectType
        Case acTable
            DatabaseAccessObjectExistsByName = DaoTableExists(db, objectName)
        Case acQuery
            DatabaseAccessObjectExistsByName = SharedDao_QueryExists(db, objectName)
        Case acForm, acReport, acMacro
            DatabaseAccessObjectExistsByName = SharedDao_AccessObjectExists(db, objectType, objectName)
    End Select
    Exit Function
ErrHandler:
    DatabaseAccessObjectExistsByName = False
End Function

' Purpose: Checks the current Access database for an object of the requested type and name.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   objectType (Input; AcObjectType): Object-type value identifying the Access/VBA object class targeted by the operation.
'   objectName (Input; String): Physical Access object name identifying the database object targeted by the operation.
' Returns: Boolean: True when an Access object of the requested type and name exists in the current database; False when it does not.
Public Function AccessObjectExistsInCurrentDatabase(ByVal objectType As AcObjectType, ByVal objectName As String) As Boolean
    'Use the native DAO catalog first. HostDb AllQueries collection can be stale or can raise while Access is refreshing accepted/saved queries. Access CurrentProject fallbacks are used only when the resolved HostDb
    'is the current project, so CodeDb UI collections cannot create HostDb false positives.
    On Error GoTo ErrHandler
    Dim obj As accessObject
    Dim db As DAO.Database
    Dim hostDbIsCurrentProject As Boolean
    If Len(objectName) = 0 Then Exit Function
    Set db = currentDb
    hostDbIsCurrentProject = DatabaseIsCurrentProject(db)
    Select Case objectType
        Case acQuery
            db.QueryDefs.Refresh
            If SharedDao_QueryExists(db, objectName) Then
                AccessObjectExistsInCurrentDatabase = True
                Exit Function
            End If
            If hostDbIsCurrentProject Then
                If AccessProjectQueryExists(objectName) Then
                    AccessObjectExistsInCurrentDatabase = True
                    Exit Function
                End If
            End If
        Case acTable
            db.TableDefs.Refresh
            If DatabaseAccessObjectExistsByName(db, acTable, objectName) Then
                AccessObjectExistsInCurrentDatabase = True
                Exit Function
            End If
            If hostDbIsCurrentProject Then
                On Error Resume Next
                For Each obj In Application.CurrentData.AllTables
                    PumpSetupResponsiveness
                    If StrComp(obj.Name, objectName, vbTextCompare) = 0 Then
                        AccessObjectExistsInCurrentDatabase = True
                        Exit Function
                    End If
                Next obj
                On Error GoTo ErrHandler
            End If
        Case acForm
            If DatabaseAccessObjectExistsByName(db, acForm, objectName) Then
                AccessObjectExistsInCurrentDatabase = True
                Exit Function
            End If
            If hostDbIsCurrentProject Then
                On Error Resume Next
                For Each obj In Application.CurrentProject.AllForms
                    PumpSetupResponsiveness
                    If StrComp(obj.Name, objectName, vbTextCompare) = 0 Then
                        AccessObjectExistsInCurrentDatabase = True
                        Exit Function
                    End If
                Next obj
                On Error GoTo ErrHandler
            End If
        Case acReport
            If DatabaseAccessObjectExistsByName(db, acReport, objectName) Then
                AccessObjectExistsInCurrentDatabase = True
                Exit Function
            End If
            If hostDbIsCurrentProject Then
                On Error Resume Next
                For Each obj In Application.CurrentProject.AllReports
                    PumpSetupResponsiveness
                    If StrComp(obj.Name, objectName, vbTextCompare) = 0 Then
                        AccessObjectExistsInCurrentDatabase = True
                        Exit Function
                    End If
                Next obj
                On Error GoTo ErrHandler
            End If
        Case acMacro
            If DatabaseAccessObjectExistsByName(db, acMacro, objectName) Then
                AccessObjectExistsInCurrentDatabase = True
                Exit Function
            End If
            If hostDbIsCurrentProject Then
                On Error Resume Next
                For Each obj In Application.CurrentProject.AllMacros
                    PumpSetupResponsiveness
                    If StrComp(obj.Name, objectName, vbTextCompare) = 0 Then
                        AccessObjectExistsInCurrentDatabase = True
                        Exit Function
                    End If
                Next obj
                On Error GoTo ErrHandler
            End If
        Case Else
            Err.Raise vbObjectError + 5500, "AccessObjectExistsInCurrentDatabase", "Unsupported Access object type: " & CStr(objectType) & vbCrLf & "Object name: " & objectName & vbCrLf & "Supported Access object types in this helper: table, query, form, report, macro."
    End Select
    Exit Function
ErrHandler:
    'Final defensive fallback. Do not let an Access UI catalog refresh problem hide a real DAO object.
    On Error Resume Next
    Set db = currentDb
    Select Case objectType
        Case acQuery
            AccessObjectExistsInCurrentDatabase = SharedDao_QueryExists(db, objectName)
        Case acTable, acForm, acReport, acMacro
            AccessObjectExistsInCurrentDatabase = DatabaseAccessObjectExistsByName(db, objectType, objectName)
        Case Else
            AccessObjectExistsInCurrentDatabase = False
    End Select
End Function

Private Function AccessProjectQueryExists(ByVal queryName As String) As Boolean
    'Returns True when the current Access project contains a query with the supplied name.
    On Error GoTo ErrHandler
    Dim obj As accessObject
    If Len(queryName) = 0 Then Exit Function
    For Each obj In Application.CurrentData.AllQueries
        PumpSetupResponsiveness
        If StrComp(obj.Name, queryName, vbTextCompare) = 0 Then
            AccessProjectQueryExists = True
            Exit Function
        End If
    Next obj
    Exit Function
ErrHandler:
    AccessProjectQueryExists = False
End Function

' Purpose: Ensures field column order for fld, table name text.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   fld (Input; DAO.Field): DAO.Field representing the field whose properties or metadata are being inspected or modified.
'   tableNameText (Input; String): Input table name text for ensure field column order; is passed to `IIf`.
'   expectedColumnOrder (Input; Long): Expected ordinal position of the field/column used when validating that the physical schema matches the canonical definition.
' Returns: Boolean result of ensure field column order; True and False correspond to the affirmative and negative branches explicitly implemented by this function.
Public Function EnsureFieldColumnOrder(ByVal fld As DAO.Field, ByVal tableNameText As String, ByVal expectedColumnOrder As Long) As Boolean
    'Generic Access ColumnOrder mechanic; preserves the existing compatible DAO property type when possible.
    On Error GoTo ErrHandler
    Dim existingPropertyType As Long
    Dim lookupErrNumber As Long
    Dim lookupErrSource As String
    Dim lookupErrDescription As String
    Dim fieldNameText As String
    If fld Is Nothing Then Err.Raise vbObjectError + 6418, "EnsureFieldColumnOrder", "The field reference is Nothing."
    fieldNameText = fld.Name
    If expectedColumnOrder < 1 Or expectedColumnOrder > 32767 Then Err.Raise vbObjectError + 6419, "EnsureFieldColumnOrder", "The requested ColumnOrder is outside the supported Integer range."
    On Error Resume Next
    Err.Clear
    existingPropertyType = CLng(fld.properties("ColumnOrder").Type)
    lookupErrNumber = Err.Number
    lookupErrSource = Err.Source
    lookupErrDescription = Err.Description
    Err.Clear
    On Error GoTo ErrHandler
    If lookupErrNumber = 0 Then
        Select Case existingPropertyType
            Case dbInteger
                EnsureFieldColumnOrder = EnsureDaoPropertyValue(fld, "ColumnOrder", dbInteger, CInt(expectedColumnOrder))
            Case dbLong
                EnsureFieldColumnOrder = EnsureDaoPropertyValue(fld, "ColumnOrder", dbLong, expectedColumnOrder)
            Case Else
                EnsureFieldColumnOrder = EnsureDaoPropertyValue(fld, "ColumnOrder", dbInteger, CInt(expectedColumnOrder), "EnsureFieldColumnOrder")
        End Select
    ElseIf lookupErrNumber = 3270 Then
        EnsureFieldColumnOrder = EnsureDaoPropertyValue(fld, "ColumnOrder", dbInteger, CInt(expectedColumnOrder))
    Else
        Err.Raise lookupErrNumber, lookupErrSource, lookupErrDescription
    End If
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "EnsureFieldColumnOrder", "Unable to ensure Access datasheet field order." & vbCrLf & "Table: " & IIf(Len(Trim$(tableNameText)) = 0, "(unresolved)", tableNameText) & vbCrLf & "Field: " & IIf(Len(fieldNameText) = 0, "(unresolved)", fieldNameText) & vbCrLf & "Requested ColumnOrder: " & CStr(expectedColumnOrder) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Function

' Purpose: Checks whether DAO field exists in the database, collection, or metadata scope supplied to this function.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   tdf (Input; DAO.tableDef): DAO.TableDef representing the table whose fields, indexes, properties, or metadata are being inspected or modified.
'   fieldName (Input; String): Field name identifying the table/query field or bound control source the routine evaluates or modifies.
' Returns: Boolean: True when DAO field exists in the inspected context; False when it does not.
Public Function DAOFieldExists(ByVal tdf As DAO.tableDef, ByVal fieldName As String) As Boolean
    'Intended use: Tests whether a field exists in an already-resolved DAO TableDef while preserving unexpected DAO lookup errors.
    On Error GoTo ErrHandler
    Dim fld As DAO.Field
    Set fld = tdf.fields(fieldName)
    DAOFieldExists = True
    Exit Function
ErrHandler:
    If Err.Number = 3265 Then
        Err.Clear
        Exit Function
    End If
    Err.Raise Err.Number, "DaoFieldExists", Err.Description
End Function

Private Function PopupTextContainsTechnicalDiagnostics(ByVal sourceText As String) As Boolean
    'Returns True when popup source text contains markers indicating technical diagnostic content.
    Dim marker As Variant
    For Each marker In Array("Err.Number:", "Err.Source:", "Err.Description:", "DBEngine.Errors:", "Expected Version ID:", "Installed Version ID:", "Verification notes:", "Failure stage:", "Registration stage:", "Manifest write started:", "UNMAPPED FIELD DESCRIPTION:", "Installed module IDs:", "exact fingerprint", "VBE-stable fingerprint", "source line count")
        If InStr(1, sourceText, CStr(marker), vbTextCompare) > 0 Then
            PopupTextContainsTechnicalDiagnostics = True
            Exit Function
        End If
    Next marker
End Function

Private Function PopupTextIndicatesSuccessfulCompletion(ByVal sourceText As String) As Boolean
    'Returns True when popup source text describes successful completion rather than a warning or failure.
    If InStr(1, sourceText, "completed successfully", vbTextCompare) = 0 And InStr(1, sourceText, "registered and verified successfully", vbTextCompare) = 0 Then Exit Function
    If InStr(1, sourceText, "failed", vbTextCompare) > 0 Or InStr(1, sourceText, "could not", vbTextCompare) > 0 Or InStr(1, sourceText, "mismatch", vbTextCompare) > 0 Then Exit Function
    PopupTextIndicatesSuccessfulCompletion = True
End Function

Private Function PopupTextIndicatesCurrentConflict(ByVal sourceText As String) As Boolean
    'Uses only explicit current-error conflict phrases. Numeric summaries such as conflict=0 and unrelated historical trace text must never classify a popup as a conflict.
    Dim conflictPhrase As Variant
    For Each conflictPhrase In Array("identity conflicts were found", "conflicting active non-form dependency definitions", "managed-definition delta contains conflicts", "duplicate active SystemMacroDesignator values exist", "conflicting managed-object dependency definitions", "duplicate managed-form action payload conflicts", "CatalogOwner conflict:", "conflicting same-name rows require")
        If InStr(1, sourceText, CStr(conflictPhrase), vbTextCompare) > 0 Then
            PopupTextIndicatesCurrentConflict = True
            Exit Function
        End If
    Next conflictPhrase
End Function

Private Function MetricValueGreaterThanZero(ByVal sourceText As String, ByVal tokenText As String) As Boolean
    'Returns True when the first integer immediately following tokenText is greater than zero.
    Dim tokenPosition As Long
    Dim scanPosition As Long
    Dim digitText As String
    tokenPosition = InStr(1, sourceText, tokenText, vbTextCompare)
    If tokenPosition = 0 Then Exit Function
    scanPosition = tokenPosition + Len(tokenText)
    Do While scanPosition <= Len(sourceText) And Mid$(sourceText, scanPosition, 1) = " "
        scanPosition = scanPosition + 1
    Loop
    Do While scanPosition <= Len(sourceText) And Mid$(sourceText, scanPosition, 1) Like "#"
        digitText = digitText & Mid$(sourceText, scanPosition, 1)
        scanPosition = scanPosition + 1
    Loop
    If Len(digitText) = 0 Then Exit Function
    MetricValueGreaterThanZero = (CLng(digitText) > 0)
End Function

' Purpose: Ensures DAO property value for target object, property name.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   targetObject (Input; Object): Access/DAO/VBE object instance whose named property/state is being inspected, validated, or updated by this helper.
'   propertyName (Input; String): Exact Access/DAO property name to read, compare, create, update, or fingerprint.
'   requestedType (Input; DAO.DataTypeEnum): DAO DataTypeEnum requested for the target field/property and used to determine whether repair/coercion is required.
'   requestedValue (Input; Variant): Caller-requested value to validate/coerce/apply to the target property or field.
'   sourceProcedure (Input Optional; String): Procedure name recorded as the source of trace, error, status, or diagnostic information emitted by this call. Default: "EnsureDaoPropertyValue".
'   targetDiagnosticContext (Input Optional; String): Target diagnostic context text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'     Default: vbNullString.
'   repairOutcome (Output Optional; Variant): When supplied, receives the outcome text produced by the DAO property ensure/repair operation so callers can distinguish the implemented unchanged/create/repair
'     outcomes.
' Returns: Boolean result of ensure DAO property value; True and False correspond to the affirmative and negative branches explicitly implemented by this function.
Public Function EnsureDaoPropertyValue(ByVal targetObject As Object, ByVal propertyName As String, ByVal requestedType As DAO.DataTypeEnum, ByVal requestedValue As Variant, Optional ByVal sourceProcedure As String = "EnsureDaoPropertyValue", Optional ByVal targetDiagnosticContext As String = vbNullString, Optional ByRef repairOutcome As Variant) As Boolean
    'Coerces the requested Variant subtype, performs verified wrong-type or capacity repair when required, and restores the original property if replacement fails.
    On Error GoTo ErrHandler
    Dim coercedValue As Variant
    Dim outcomeText As String
    If requestedType = dbText Or requestedType = dbMemo Then
        If IsObject(requestedValue) Then Err.Raise vbObjectError + 6886, sourceProcedure, "A DAO text property cannot be assigned from an object value."
        If IsArray(requestedValue) Then Err.Raise vbObjectError + 6887, sourceProcedure, "A DAO text property cannot be assigned from an array value."
        If IsNull(requestedValue) Then Err.Raise vbObjectError + 6888, sourceProcedure, "A DAO text property cannot be assigned a Null value."
    End If
    Select Case requestedType
        Case dbBoolean
            coercedValue = CBool(requestedValue)
        Case dbByte
            coercedValue = CByte(requestedValue)
        Case dbInteger
            coercedValue = CInt(requestedValue)
        Case dbLong
            coercedValue = CLng(requestedValue)
        Case dbSingle
            coercedValue = CSng(requestedValue)
        Case dbDouble
            coercedValue = CDbl(requestedValue)
        Case dbCurrency
            coercedValue = CCur(requestedValue)
        Case dbDate
            coercedValue = CDate(requestedValue)
        Case dbText, dbMemo
            coercedValue = CStr(requestedValue)
            If requestedType = dbText Then
                If TypeOf targetObject Is DAO.Field Then
                    If StrComp(propertyName, DAO_DESCRIPTION_PROPERTY_NAME, vbTextCompare) = 0 Then
                        If Len(CStr(coercedValue)) > 255 Then Err.Raise vbObjectError + 6889, sourceProcedure, "The approved Access field Description exceeds the supported 255-character limit." & vbCrLf & "Field: " & CStr(targetObject.Name) & vbCrLf & "Property: " & propertyName & vbCrLf & "Requested characters: " & CStr(Len(CStr(coercedValue))) & vbCrLf & "Supported characters: 255" & vbCrLf & "Default/value source: (not supplied by the generic DAO property caller)"
                    End If
                End If
            End If
            If LenB(CStr(coercedValue)) > DAO_PROPERTY_TEXT_LIMIT_BYTES Then Err.Raise vbObjectError + 6889, sourceProcedure, "The approved DAO text property value exceeds the Access 2 KB property limit." & vbCrLf & "Property: " & propertyName & vbCrLf & "Requested characters: " & CStr(Len(CStr(coercedValue))) & vbCrLf & "Requested bytes: " & CStr(LenB(CStr(coercedValue))) & vbCrLf & "Supported bytes: " & CStr(DAO_PROPERTY_TEXT_LIMIT_BYTES)
        Case Else
            coercedValue = requestedValue
    End Select
    EnsureDaoPropertyValue = EnsureDaoPropertyValueCore(targetObject, propertyName, requestedType, coercedValue, sourceProcedure, targetDiagnosticContext, outcomeText)
    If isMissing(repairOutcome) = False Then repairOutcome = outcomeText
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, sourceProcedure, "Could not coerce, ensure, and verify the DAO property value." & vbCrLf & "Caller: " & sourceProcedure & vbCrLf & DaoTargetContextText(targetDiagnosticContext) & "Property: " & propertyName & vbCrLf & "Requested DAO type: " & CStr(requestedType) & vbCrLf & "Requested value: " & DaoPropertyDiagnosticValue(requestedValue) & vbCrLf & "Requested VarType: " & CStr(VarType(requestedValue)) & vbCrLf & "Requested characters: " & DaoPropertyCharacterLengthText(requestedValue) & vbCrLf & "Requested bytes: " & DaoPropertyByteLengthText(requestedValue) & vbCrLf & "Original error: " & CStr(savedErrNumber) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Original description: " & savedErrDescription
End Function

Private Function EnsureDaoPropertyValueCore(ByVal targetObject As Object, ByVal propertyName As String, ByVal requestedType As DAO.DataTypeEnum, ByVal requestedValue As Variant, ByVal sourceProcedure As String, ByVal targetDiagnosticContext As String, ByRef repairOutcome As String) As Boolean
    'Ensures a DAO property exists with the requested type and value, repairing incompatible existing properties when required.
    On Error GoTo ErrHandler
    Dim prp As DAO.Property
    Dim lookupErrNumber As Long
    Dim lookupErrSource As String
    Dim lookupErrDescription As String
    Dim propertyPreviouslyExisted As Boolean
    Dim existingTypeText As String
    Dim failureStage As String
    Dim wrongTypeReplacementAttempted As Boolean
    Dim capacityReplacementAttempted As Boolean
    Dim originalType As DAO.DataTypeEnum
    Dim originalValue As Variant
    Dim assignmentErrNumber As Long
    Dim assignmentErrSource As String
    Dim assignmentErrDescription As String
    Dim assignmentDaoErrorsText As String
    Dim detachedOwnerDatabase As DAO.Database
    repairOutcome = vbNullString
    If targetObject Is Nothing Then Err.Raise vbObjectError + 6880, sourceProcedure, "The DAO property target is Nothing."
    propertyName = Trim$(propertyName)
    If Len(propertyName) = 0 Then Err.Raise vbObjectError + 6881, sourceProcedure, "The DAO property name is blank."
    failureStage = "Read existing DAO property"
    ReadDaoPropertyReference targetObject, propertyName, prp, propertyPreviouslyExisted, lookupErrNumber, lookupErrSource, lookupErrDescription
    If lookupErrNumber <> 0 And lookupErrNumber <> 3270 Then Err.Raise lookupErrNumber, lookupErrSource, lookupErrDescription
    If TypeOf targetObject Is DAO.Relation Then
        If propertyPreviouslyExisted = False Then Err.Raise vbObjectError + 9260, sourceProcedure, "DAO.Relation does not support creation of user-defined properties. Relationship metadata must be represented by the canonical relationship definition and supported Relation properties."
        If CLng(prp.Type) <> CLng(requestedType) Then Err.Raise vbObjectError + 9261, sourceProcedure, "DAO.Relation does not support replacement of a property with a newly created user-defined property. Relationship metadata must be represented by the canonical relationship definition and supported Relation properties."
    End If
    If propertyPreviouslyExisted Then
        existingTypeText = CStr(prp.Type)
        originalType = prp.Type
        originalValue = prp.value
        If CLng(prp.Type) <> CLng(requestedType) Then
            wrongTypeReplacementAttempted = True
            failureStage = "Replace existing wrong-type DAO property"
            Set prp = Nothing
            ReplaceDaoPropertyPreservingOriginal targetObject, propertyName, requestedType, requestedValue, originalType, originalValue, sourceProcedure, targetDiagnosticContext, "WrongType", 0, vbNullString, vbNullString, vbNullString
            repairOutcome = "WrongTypeRecreated"
            EnsureDaoPropertyValueCore = True
            Exit Function
        End If
        failureStage = "Compare existing DAO property value"
        If DaoPropertyValuesEqual(prp.value, requestedValue, requestedType) Then Exit Function
        failureStage = "Assign requested value to existing DAO property"
        On Error Resume Next
        Err.Clear
        prp.value = requestedValue
        assignmentErrNumber = Err.Number
        assignmentErrSource = Err.Source
        assignmentErrDescription = Err.Description
        assignmentDaoErrorsText = DaoEngineErrorsText()
        Err.Clear
        On Error GoTo ErrHandler
        If assignmentErrNumber = 0 Then
            RefreshDaoProperties targetObject
            failureStage = "Reacquire existing DAO property for verification"
            Set prp = targetObject.properties(propertyName)
            failureStage = "Verify existing DAO property type and value"
            VerifyDaoPropertyValue prp, propertyName, requestedType, requestedValue, sourceProcedure
            repairOutcome = "Assigned"
            EnsureDaoPropertyValueCore = True
            Exit Function
        End If
        If IsDaoPropertyCapacityError(assignmentErrNumber, assignmentErrDescription, assignmentDaoErrorsText) And (requestedType = dbText Or requestedType = dbMemo) Then
            capacityReplacementAttempted = True
            failureStage = "Replace correct-type undersized DAO text property"
            Set prp = Nothing
            ReplaceDaoPropertyPreservingOriginal targetObject, propertyName, requestedType, requestedValue, originalType, originalValue, sourceProcedure, targetDiagnosticContext, "Capacity3309", assignmentErrNumber, assignmentErrSource, assignmentErrDescription, assignmentDaoErrorsText
            repairOutcome = "CapacityRecreated"
            EnsureDaoPropertyValueCore = True
            Exit Function
        End If
        Err.Raise assignmentErrNumber, assignmentErrSource, assignmentErrDescription
    End If
    failureStage = "Create missing detached DAO property"
    Set prp = CreateDetachedDaoProperty(targetObject, detachedOwnerDatabase, propertyName, requestedType, requestedValue)
    failureStage = "Append missing DAO property"
    targetObject.properties.Append prp
    RefreshDaoProperties targetObject
    failureStage = "Reacquire appended DAO property"
    Set prp = targetObject.properties(propertyName)
    failureStage = "Assign requested value to appended DAO property"
    prp.value = requestedValue
    RefreshDaoProperties targetObject
    failureStage = "Reacquire appended DAO property for verification"
    Set prp = targetObject.properties(propertyName)
    failureStage = "Verify appended DAO property type and value"
    VerifyDaoPropertyValue prp, propertyName, requestedType, requestedValue, sourceProcedure
    repairOutcome = "Created"
    EnsureDaoPropertyValueCore = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim handlerDaoErrorsText As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    handlerDaoErrorsText = DaoEngineErrorsText()
    Err.Raise savedErrNumber, sourceProcedure, _
        "Unable to ensure and verify the DAO property value." & vbCrLf & _
        "Caller: " & sourceProcedure & vbCrLf & _
        DaoTargetContextText(targetDiagnosticContext) & _
        "Property: " & propertyName & vbCrLf & _
        "Requested DAO type: " & CStr(requestedType) & vbCrLf & _
        "Requested value: " & DaoPropertyDiagnosticValue(requestedValue) & vbCrLf & _
        "Requested VarType: " & CStr(VarType(requestedValue)) & vbCrLf & _
        "Requested characters: " & DaoPropertyCharacterLengthText(requestedValue) & vbCrLf & _
        "Requested bytes: " & DaoPropertyByteLengthText(requestedValue) & vbCrLf & _
        "Property previously existed: " & CStr(propertyPreviouslyExisted) & vbCrLf & _
        "Existing DAO type: " & IIf(Len(existingTypeText) = 0, "(not read)", existingTypeText) & vbCrLf & _
        "Original value characters: " & DaoPropertyCharacterLengthText(originalValue) & vbCrLf & _
        "Original value bytes: " & DaoPropertyByteLengthText(originalValue) & vbCrLf & _
        "Wrong-type replacement attempted: " & CStr(wrongTypeReplacementAttempted) & vbCrLf & _
        "Capacity replacement attempted: " & CStr(capacityReplacementAttempted) & vbCrLf & _
        "Failure stage: " & IIf(Len(failureStage) = 0, "(not recorded)", failureStage) & vbCrLf & _
        "Original error: " & CStr(savedErrNumber) & vbCrLf & _
        "Original source: " & savedErrSource & vbCrLf & _
        "Original description: " & savedErrDescription & _
        IIf(Len(handlerDaoErrorsText) > 0, vbCrLf & "DBEngine.Errors:" & vbCrLf & handlerDaoErrorsText, vbNullString)
End Function

Private Function TryRestoreDaoProperty(ByVal targetObject As Object, ByVal propertyName As String, ByVal originalType As DAO.DataTypeEnum, ByVal originalValue As Variant, ByVal sourceProcedure As String, ByVal targetDiagnosticContext As String, ByRef restorationDetail As String) As Boolean
    'Attempts to restore a DAO property to its original type and value after a failed replacement.
    On Error GoTo RestoreFailed
    Dim prp As DAO.Property
    Dim propertyExists As Boolean
    Dim lookupErrNumber As Long
    Dim lookupErrSource As String
    Dim lookupErrDescription As String
    Dim targetOwnerDatabase As DAO.Database
    If Len(DaoDiagnosticContextValue(targetDiagnosticContext, "Table")) > 0 And Len(DaoDiagnosticContextValue(targetDiagnosticContext, "Field")) > 0 Then
        Set targetObject = ReacquireDaoPropertyTarget(targetObject, targetDiagnosticContext, sourceProcedure, targetOwnerDatabase)
    ElseIf targetObject Is Nothing Then
        Err.Raise vbObjectError + 6893, sourceProcedure, "The original DAO property could not be restored because its target is unavailable."
    End If
    ReadDaoPropertyReference targetObject, propertyName, prp, propertyExists, lookupErrNumber, lookupErrSource, lookupErrDescription
    If propertyExists Then
        If CLng(prp.Type) = CLng(originalType) Then
            If DaoPropertyValuesEqual(prp.value, originalValue, originalType) Then restorationDetail = "Original property remained intact.": TryRestoreDaoProperty = True: Exit Function
        End If
        targetObject.properties.Delete propertyName
        RefreshDaoProperties targetObject
    ElseIf lookupErrNumber <> 0 And lookupErrNumber <> 3270 Then
        Err.Raise lookupErrNumber, lookupErrSource, lookupErrDescription
    End If
    Set targetObject = Nothing
    DBEngine.Idle dbRefreshCache
    Set targetObject = ReacquireDaoPropertyTarget(targetObject, targetDiagnosticContext, sourceProcedure, targetOwnerDatabase)
    Set prp = CreateDetachedDaoProperty(targetObject, targetOwnerDatabase, propertyName, originalType, originalValue)
    targetObject.properties.Append prp
    RefreshDaoProperties targetObject
    Set prp = targetObject.properties(propertyName)
    prp.value = originalValue
    RefreshDaoProperties targetObject
    Set prp = targetObject.properties(propertyName)
    VerifyDaoPropertyValue prp, propertyName, originalType, originalValue, sourceProcedure
    restorationDetail = "Original property type and value were restored and verified."
    TryRestoreDaoProperty = True
    Exit Function
RestoreFailed:
    Dim restoreErrNumber As Long
    Dim restoreErrSource As String
    Dim restoreErrDescription As String
    Dim restoreDaoErrors As String
    restoreErrNumber = Err.Number
    restoreErrSource = Err.Source
    restoreErrDescription = Err.Description
    restoreDaoErrors = DaoEngineErrorsText()
    restorationDetail = "Restoration failed. Error " & CStr(restoreErrNumber) & " from " & restoreErrSource & ": " & restoreErrDescription & IIf(Len(restoreDaoErrors) > 0, " | DBEngine.Errors: " & Replace(restoreDaoErrors, vbCrLf, " | "), vbNullString)
End Function

Private Function CreateDetachedDaoProperty(ByVal targetObject As Object, ByVal ownerDatabase As DAO.Database, ByVal propertyName As String, ByVal requestedType As DAO.DataTypeEnum, ByVal requestedValue As Variant) As DAO.Property
    'Creates an unattached Property with explicit maximum field-Description capacity before the exact value is persisted.
    Dim prp As DAO.Property
    If Not (ownerDatabase Is Nothing) Then
        Set prp = ownerDatabase.CreateProperty()
    Else
        Set prp = targetObject.CreateProperty()
    End If
    prp.Name = propertyName
    prp.Type = requestedType
    If requestedType = dbText And TypeOf targetObject Is DAO.Field And StrComp(propertyName, DAO_DESCRIPTION_PROPERTY_NAME, vbTextCompare) = 0 Then
        prp.value = String$(255, " ")
    End If
    prp.value = requestedValue
    Set CreateDetachedDaoProperty = prp
End Function

Private Function ReacquireDaoPropertyTarget(ByVal targetObject As Object, ByVal targetDiagnosticContext As String, ByVal sourceProcedure As String, ByRef ownerDatabase As DAO.Database) As Object
    'Reacquires the live DAO object that owns a property after collection refresh or object invalidation.
    On Error GoTo ErrHandler
    Dim tableName As String
    Dim fieldName As String
    tableName = DaoDiagnosticContextValue(targetDiagnosticContext, "Table")
    fieldName = DaoDiagnosticContextValue(targetDiagnosticContext, "Field")
    If Len(tableName) > 0 And Len(fieldName) > 0 Then
        Set ownerDatabase = DBEngine.Workspaces(0).OpenDatabase(currentDb.Name, False, False)
        ownerDatabase.TableDefs.Refresh
        ownerDatabase.TableDefs(tableName).fields.Refresh
        Set ReacquireDaoPropertyTarget = ownerDatabase.TableDefs(tableName).fields(fieldName)
        Exit Function
    End If
    If targetObject Is Nothing Then Err.Raise vbObjectError + 6891, sourceProcedure, "The DAO property target could not be reacquired because the original target was released and no table/field context was supplied."
    Set ReacquireDaoPropertyTarget = targetObject
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, sourceProcedure, "Could not reacquire the DAO property target after deleting the existing property." & vbCrLf & DaoTargetContextText(targetDiagnosticContext) & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function IsDaoPropertyCapacityError(ByVal errorNumber As Long, ByVal errorDescription As String, Optional ByVal daoErrorsText As String = vbNullString) As Boolean
    'Returns True when an error represents a DAO property capacity or value-size failure.
    Dim normalizedDescription As String
    Dim normalizedDaoErrors As String
    normalizedDescription = errorDescription
    normalizedDaoErrors = daoErrorsText
    If errorNumber = 3309 Then IsDaoPropertyCapacityError = True: Exit Function
    If (errorNumber And &HFFFF&) = 3309 Then IsDaoPropertyCapacityError = True: Exit Function
    If InStr(1, normalizedDescription, "property value is too large", vbTextCompare) > 0 Then IsDaoPropertyCapacityError = True: Exit Function
    If InStr(1, normalizedDescription, "property setting cannot be larger than 2k", vbTextCompare) > 0 Then IsDaoPropertyCapacityError = True: Exit Function
    If InStr(1, normalizedDaoErrors, "number=3309", vbTextCompare) > 0 Then IsDaoPropertyCapacityError = True: Exit Function
    If InStr(1, normalizedDaoErrors, "property value is too large", vbTextCompare) > 0 Then IsDaoPropertyCapacityError = True: Exit Function
    If InStr(1, normalizedDaoErrors, "property setting cannot be larger than 2k", vbTextCompare) > 0 Then IsDaoPropertyCapacityError = True
End Function

Private Function DaoPropertyValuesEqual(ByVal actualValue As Variant, ByVal requestedValue As Variant, ByVal requestedType As DAO.DataTypeEnum) As Boolean
    'Returns True when an actual DAO property value equals the requested value under the requested DAO type.
    On Error GoTo NotEqual
    If IsNull(actualValue) Or IsNull(requestedValue) Then DaoPropertyValuesEqual = (IsNull(actualValue) And IsNull(requestedValue)): Exit Function
    If IsArray(actualValue) Or IsArray(requestedValue) Then DaoPropertyValuesEqual = DaoPropertyByteArraysEqual(actualValue, requestedValue): Exit Function
    Select Case requestedType
        Case dbBoolean
            DaoPropertyValuesEqual = (CBool(actualValue) = CBool(requestedValue))
        Case dbByte
            DaoPropertyValuesEqual = (CByte(actualValue) = CByte(requestedValue))
        Case dbInteger
            DaoPropertyValuesEqual = (CInt(actualValue) = CInt(requestedValue))
        Case dbLong
            DaoPropertyValuesEqual = (CLng(actualValue) = CLng(requestedValue))
        Case dbSingle, dbDouble
            DaoPropertyValuesEqual = (CDbl(actualValue) = CDbl(requestedValue))
        Case dbCurrency
            DaoPropertyValuesEqual = (CCur(actualValue) = CCur(requestedValue))
        Case dbDate
            DaoPropertyValuesEqual = (CDate(actualValue) = CDate(requestedValue))
        Case dbText, dbMemo
            DaoPropertyValuesEqual = (StrComp(CStr(actualValue), CStr(requestedValue), vbBinaryCompare) = 0)
        Case Else
            DaoPropertyValuesEqual = (StrComp(CStr(actualValue), CStr(requestedValue), vbBinaryCompare) = 0)
    End Select
    Exit Function
NotEqual:
    DaoPropertyValuesEqual = False
End Function

Private Function EnsureLookupTextPropertyValue(ByVal targetObject As Object, ByVal propertyName As String, ByVal propertyValue As String, ByVal sourceProcedure As String, ByVal nothingErrorNumber As Long, ByVal blankNameErrorNumber As Long) As Boolean
    'Preserves an existing supported text/memo type and routes any unsupported existing type through verified dbText replacement.
    On Error GoTo ErrHandler
    Dim prp As DAO.Property
    Dim requestedType As DAO.DataTypeEnum
    Dim lookupErrNumber As Long
    Dim lookupErrSource As String
    Dim lookupErrDescription As String
    If targetObject Is Nothing Then Err.Raise nothingErrorNumber, sourceProcedure, "The lookup-property target is Nothing."
    propertyName = Trim$(propertyName)
    If Len(propertyName) = 0 Then Err.Raise blankNameErrorNumber, sourceProcedure, "The lookup-property name is blank."
    On Error Resume Next
    Err.Clear
    Set prp = targetObject.properties(propertyName)
    lookupErrNumber = Err.Number
    lookupErrSource = Err.Source
    lookupErrDescription = Err.Description
    Err.Clear
    On Error GoTo ErrHandler
    If lookupErrNumber = 0 Then
        Select Case prp.Type
            Case dbText, dbMemo
                requestedType = prp.Type
            Case Else
                requestedType = dbText
        End Select
    ElseIf lookupErrNumber = 3270 Then
        requestedType = dbText
    Else
        Err.Raise lookupErrNumber, lookupErrSource, lookupErrDescription
    End If
    EnsureLookupTextPropertyValue = EnsureDaoPropertyValue(targetObject, propertyName, requestedType, propertyValue, sourceProcedure)
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, sourceProcedure, "Could not ensure the lookup text property without changing its existing DAO text type." & vbCrLf & "Property: " & propertyName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Function

' Purpose: Compares the supplied/current values required by database text property matches and reports whether they match under this function's exact comparison rules.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   db (Input; DAO.Database): DAO.Database instance whose schema, data, queries, or properties the routine reads or modifies.
'   propertyName (Input; String): Exact Access/DAO property name to read, compare, create, update, or fingerprint.
'   expectedValue (Input; String): Exact expected text used for comparison/verification against the current Access property, metadata, or source value.
'   propertyExists (Output; Boolean): Receives True when the named database text property exists and False when it is absent; the Function separately reports whether its value matches expectedValue.
' Returns: Boolean: True when the database text property matches condition is satisfied for the supplied or current ADC context; False otherwise.
Public Function DatabaseTextPropertyMatches(ByVal db As DAO.Database, ByVal propertyName As String, ByVal expectedValue As String, ByRef propertyExists As Boolean) As Boolean
    'Intended use: Tests whether a persisted database text property exists and exactly matches an expected fingerprint or state-token value.
    Dim storedValue As String
    storedValue = DatabaseTextPropertyValue(db, propertyName, propertyExists)
    If propertyExists = False Then Exit Function
    DatabaseTextPropertyMatches = (StrComp(storedValue, expectedValue, vbBinaryCompare) = 0)
End Function

' COMMON HELPER FUNCTIONS - END
' ================================================================================================================================================================================
