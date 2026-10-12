Attribute VB_Name = "x__UTIL_ADC_REPORTS"
' ================================================================================================================================================================================
' ChatGPT Rewrite Tracking Header
' ================================================================================================================================================================================
' ChatGPT Previous Version ID: CGPT-VF2P-55X7-5VRL-8V8Q-L3KL
' ChatGPT Baseline Version ID: CGPT-VF2P-55X7-5VRL-8V8Q-L3KL
' ChatGPT Version ID: CGPT-8VHX-76FK-GHB3-VGBG-G8SX
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
' STANDARD MODULE: Access Development Console Reports
' VB_Name: x__UTIL_ADC_REPORTS
' Entry Points:
'     PreflightManagedReportsForSetup
'     BuildAndVerifyManagedReportsForSetup
'     ReconcileManagedReportDomainForSetup
'     ManagedReportStateTokenForFinalProof
'     ManagedReportDependencyDefinitionsForSetup
' Purpose:
'     Sole physical Access Report authority for ADC. Owns the nine canonical report layout specifications, physical report construction/currentness, Phase-E preflight/build/reconciliation, and reusable same-run
'     proof.
'     It consumes existing authoritative tables/queries and never becomes an authority for the displayed domain data.
' Author: Christina Shonts
' Support: Access Development Console Setup Phase E and read-only reporting surfaces.
' Code Category: Access Development Console physical Reports.
' Requirements:
'     Microsoft Access VBA, DAO, x__UTIL_ADC_Core, x__UTIL_ADC_DEFAULTS, and the existing Phase-A saved-query foundation.
' Reference Dependencies: Access object model and DAO through the companion ADC modules.
' DB Object Dependencies:
'     Existing Settings, Candidates/Config, MacroCatalog_Sys, Health Review, Module Guard, MAC, and Add-In/Reference tables/queries established by their authoritative owners before Phase E. No saved query is created
'     here.
' Core Rules Enforced Here:
'     Reports are presentation objects only. Opening/building a report never scans, repairs, promotes, validates domain data, generates macros, changes Settings, repairs references, or performs work owned by another
'     Setup phase.
'     Managed Object Catalog identity/placement remains Defaults/Catalog authority; descriptions remain Phase G authority; Ribbon remains Phase H authority; Navigation Pane placement remains Phase I authority.
' Design Map:
'     E1 report preflight; E2 physical build/currentness proof; E3 logical/physical report-domain reconciliation.
' Code Status: Active Setup Phase-E physical Report authority.
' ================================================================================================================================================================================

Option Compare Database
Option Explicit

' Canonical physical-report definition row indexes.
Private Const REPORT_DEF_LOGICAL_NAME As Long = 0
Private Const REPORT_DEF_TITLE As Long = 1
Private Const REPORT_DEF_SOURCE_NAME As Long = 2
Private Const REPORT_DEF_SORT_FIELD As Long = 3
Private Const REPORT_DEF_SOURCE_KIND As Long = 4
Private Const REPORT_DEF_GROUP_FIELD As Long = 5
Private Const REPORT_DEF_LAYOUT_MODE As Long = 6
Private Const REPORT_DEF_PRESENTATION_SPEC As Long = 7

' Canonical field-presentation specification indexes.
Private Const REPORT_FIELD_NAME As Long = 0
Private Const REPORT_FIELD_CAPTION As Long = 1
Private Const REPORT_FIELD_ROW_KEY As Long = 2
Private Const REPORT_FIELD_PREFERRED_WIDTH As Long = 3
Private Const REPORT_FIELD_MIN_WIDTH As Long = 4
Private Const REPORT_FIELD_ALIGNMENT As Long = 5
Private Const REPORT_FIELD_FORMAT As Long = 6
Private Const REPORT_FIELD_CAN_GROW As Long = 7
Private Const REPORT_FIELD_VISIBLE As Long = 8
Private Const REPORT_FIELD_CONDITIONAL_MODE As Long = 9
Private Const REPORT_FIELD_SPEC_PART_COUNT As Long = 10

Private Const REPORT_SOURCE_QUERY As String = "Query"
Private Const REPORT_SOURCE_TABLE As String = "Table"
Private Const REPORT_LAYOUT_TABULAR As String = "TABULAR"
Private Const REPORT_LAYOUT_TABULAR_DETAIL As String = "TABULAR_DETAIL"
Private Const REPORT_LAYOUT_RECORD_DETAIL As String = "RECORD_DETAIL"
Private Const REPORT_ROW_PRIMARY As String = "P"
Private Const REPORT_CONDITIONAL_ALWAYS As String = "ALWAYS"
Private Const REPORT_CONDITIONAL_NONEMPTY As String = "NONEMPTY"
Private Const REPORT_FORMAT_YESNO As String = "YESNO"
Private Const REPORT_FORMAT_DATETIME As String = "DATETIME"
Private Const REPORT_FORMAT_DATE As String = "DATE"
Private Const REPORT_FORMAT_INTEGER As String = "INTEGER"
Private Const REPORT_ALIGN_LEFT As String = "L"
Private Const REPORT_ALIGN_CENTER As String = "C"
Private Const REPORT_ALIGN_RIGHT As String = "R"
Private Const REPORT_TEXT_ALIGN_LEFT As Long = 1
Private Const REPORT_TEXT_ALIGN_CENTER As Long = 2
Private Const REPORT_TEXT_ALIGN_RIGHT As Long = 3

' Physical layout/currentness contract.
Private Const REPORT_LAYOUT_VERSION As String = "ADC_REPORT_LAYOUT_V12"
Private Const REPORT_EXPECTED_COUNT As Long = 9
Private Const REPORT_PAGE_WIDTH As Long = 14500
Private Const REPORT_LEFT_MARGIN As Long = 180
Private Const REPORT_RIGHT_MARGIN As Long = 180
Private Const REPORT_TOP_MARGIN As Long = 240
Private Const REPORT_BOTTOM_MARGIN As Long = 240
Private Const REPORT_PAGE_HEADER_HEIGHT As Long = 870
Private Const REPORT_PAGE_HEADER_NOT_WITH_REPORT_HEADER As Long = 1 'Access Report.PageHeader setting: suppress Page Header only on pages that contain the Report Header.
Private Const REPORT_TITLE_TOP As Long = 30
Private Const REPORT_TITLE_HEIGHT As Long = 300
Private Const REPORT_COLUMN_HEADER_TOP As Long = 405
Private Const REPORT_COLUMN_HEADER_HEIGHT As Long = 390
Private Const REPORT_PRIMARY_CONTROL_TOP As Long = 15
Private Const REPORT_CONTROL_HEIGHT As Long = 345
Private Const REPORT_DETAIL_ROW_GAP As Long = 75
Private Const REPORT_DETAIL_LABEL_GAP As Long = 75
Private Const REPORT_RECORD_HEADER_HEIGHT As Long = 420
Private Const REPORT_RECORD_DETAIL_TOP As Long = 75
Private Const REPORT_RECORD_LABEL_HEIGHT As Long = 240
Private Const REPORT_RECORD_LABEL_VALUE_GAP As Long = 30
Private Const REPORT_RECORD_PRIMARY_VALUE_HEIGHT As Long = 480
Private Const REPORT_RECORD_VALUE_HEIGHT As Long = 360
Private Const REPORT_RECORD_GROWING_VALUE_HEIGHT As Long = 450
Private Const REPORT_RECORD_ROW_GAP As Long = 90
Private Const REPORT_RECORD_COLUMN_GAP As Long = 120
Private Const REPORT_RECORD_BOTTOM_PADDING As Long = 180
Private Const REPORT_GROUP_HEADER_HEIGHT As Long = 360
Private Const REPORT_GROUP_CONTROL_TOP As Long = 15
Private Const REPORT_FONT_SIZE_TITLE As Long = 15
Private Const REPORT_FONT_SIZE_GROUP As Long = 10
Private Const REPORT_FONT_SIZE_HEADER As Long = 9
Private Const REPORT_FONT_SIZE_DETAIL As Long = 9
Private Const REPORT_FONT_SIZE_IDENTITY As Long = 10
Private Const REPORT_MAX_DETAIL_LABEL_WIDTH As Long = 1800
Private Const REPORT_MIN_DETAIL_LABEL_WIDTH As Long = 760
Private Const REPORT_MIN_DETAIL_VALUE_WIDTH As Long = 650

' Same-run Phase-E proof. Lifetime is one active trace/database; it is invalidated automatically when either identity changes.
' INTENTIONAL OPERATION-SCOPED CONSISTENCY: E3 requires the same-run E2 token and freshly re-verifies each managed report's exact E2-owned physical currentness before reconciliation.
' J1 can consume the resulting E3 proof in the same trace/database context.
Private mReportProofTraceRunID As String
Private mReportProofDatabaseName As String
Private mReportPreflightToken As String
Private mReportBuildToken As String
Private mReportReconciledToken As String
' Report presentation caching is split by mutability. Parsed canonical presentation definitions are immutable and reusable for the active trace/database; source-field indexes and selected-field projections are
' mutable-source-derived and are rebuilt after each E2 responsiveness boundary and at E3/currentness verification entry.
Private mReportSourceFieldIndexByName As Object
Private mReportPresentationCacheByReport As Object
Private mReportSelectedFieldsByReport As Object

' ===============================================================================================================================================================================
' REPORT DEFINITIONS / PRESENTATION SPECIFICATIONS
' ===============================================================================================================================================================================

Private Function ReportFieldSpec(ByVal fieldName As String, ByVal captionText As String, ByVal rowKey As String, ByVal preferredWidth As Long, ByVal minWidth As Long, ByVal alignmentKey As String, Optional ByVal formatKey As String = vbNullString, Optional ByVal canGrow As Boolean = False, Optional ByVal isVisible As Boolean = True, Optional ByVal conditionalMode As String = REPORT_CONDITIONAL_ALWAYS) As String
    ReportFieldSpec = fieldName & "~" & captionText & "~" & rowKey & "~" & CStr(preferredWidth) & "~" & CStr(minWidth) & "~" & alignmentKey & "~" & formatKey & "~" & IIf(canGrow, "1", "0") & "~" & IIf(isVisible, "1", "0") & "~" & conditionalMode
End Function

Private Sub AppendReportFieldSpec(ByRef specText As String, ByVal fieldName As String, ByVal captionText As String, ByVal rowKey As String, ByVal preferredWidth As Long, ByVal minWidth As Long, ByVal alignmentKey As String, Optional ByVal formatKey As String = vbNullString, Optional ByVal canGrow As Boolean = False, Optional ByVal isVisible As Boolean = True, Optional ByVal conditionalMode As String = REPORT_CONDITIONAL_ALWAYS)
    If Len(specText) > 0 Then specText = specText & ";"
    specText = specText & ReportFieldSpec(fieldName, captionText, rowKey, preferredWidth, minWidth, alignmentKey, formatKey, canGrow, isVisible, conditionalMode)
End Sub

Private Function SettingsReportPresentationSpec() As String
    Dim s As String
    AppendReportFieldSpec s, FIELD_SETTING_KEY, "Setting Key", "P", 8000, 5600, "L"
    AppendReportFieldSpec s, FIELD_EFFECTIVE_VALUE, "Effective Value", "P", 6000, 4200, "L"
    AppendReportFieldSpec s, FIELD_SETTING_VALUE, "Setting Value", "D1", 4300, 3200, "L"
    AppendReportFieldSpec s, FIELD_MANAGED_OBJECT_FIELD_DEFAULT_VALUE, "Default Value", "D1", 4300, 3200, "L"
    AppendReportFieldSpec s, FIELD_MANAGED_OBJECT_FIELD_DATA_TYPE, "Type", "D1", 2200, 1700, "L"
    AppendReportFieldSpec s, "CustomizationState", "Customization State", "D1", 3000, 2300, "L"
    AppendReportFieldSpec s, FIELD_ACTIVE, "Active", "D2", 2000, 1500, "C", "YESNO"
    AppendReportFieldSpec s, FIELD_ADVANCED, "Advanced", "D2", 2100, 1600, "C", "YESNO"
    AppendReportFieldSpec s, FIELD_RESTART_REQUIRED, "Restart Required", "D2", 2600, 2000, "C", "YESNO"
    AppendReportFieldSpec s, FIELD_READ_ONLY, "Read Only", "D2", 2100, 1600, "C", "YESNO"
    AppendReportFieldSpec s, FIELD_REQUIRED, "Required", "D2", 2100, 1600, "C", "YESNO"
    AppendReportFieldSpec s, FIELD_LAST_UPDATED_ON, "Last Updated", "D3", 14000, 6000, "L", "DATETIME"
    AppendReportFieldSpec s, FIELD_DESCRIPTION, "Description", "D4", 7000, 5000, "L", vbNullString, True
    AppendReportFieldSpec s, FIELD_VALIDATION_NOTES, "Validation Notes", "D4", 7000, 5000, "L", vbNullString, True, True, REPORT_CONDITIONAL_NONEMPTY
    SettingsReportPresentationSpec = s
End Function

Private Function CandidatesReportPresentationSpec() As String
    Dim s As String
    AppendReportFieldSpec s, FIELD_CANDIDATE_MACRO_KEY, "Candidate Key", "P", 7600, 5400, "L"
    AppendReportFieldSpec s, FIELD_MACRO_TITLE, "Title", "P", 6400, 4600, "L"
    AppendReportFieldSpec s, FIELD_CANDIDATE_STATUS, "Candidate Status", "D1", 3400, 2600, "L"
    AppendReportFieldSpec s, FIELD_CONFIGURATION_STATUS, "Config Status", "D1", 3400, 2600, "L"
    AppendReportFieldSpec s, FIELD_MANUAL_REVIEW_STATUS, "Review Status", "D1", 3400, 2600, "L"
    AppendReportFieldSpec s, FIELD_REVIEW_PRIORITY, "Priority", "D1", 2200, 1700, "R", "INTEGER"
    AppendReportFieldSpec s, FIELD_MACRO_ACTION, "Action", "D2", 2800, 2200, "L"
    AppendReportFieldSpec s, FIELD_TARGET_PROCEDURE_NAME, "Target Procedure", "D2", 4500, 3400, "L"
    AppendReportFieldSpec s, FIELD_TARGET_OBJECT_NAME, "Target Object", "D2", 4400, 3300, "L"
    AppendReportFieldSpec s, FIELD_CANDIDATE_OWNER_DATABASE_ROLE, "Owner Role", "D2", 2500, 1900, "L"
    AppendReportFieldSpec s, FIELD_ACCEPTANCE_STATUS, "Acceptance", "D3", 3300, 2500, "L"
    AppendReportFieldSpec s, FIELD_LAST_VALIDATION_RESULT, "Validation", "D3", 3400, 2600, "L"
    AppendReportFieldSpec s, FIELD_LAST_GENERATED_RESULT, "Generation", "D3", 3400, 2600, "L"
    AppendReportFieldSpec s, FIELD_MACRO_SUBCATEGORY_KEY, "Subcategory", "D3", 3300, 2500, "L"
    AppendReportFieldSpec s, FIELD_CANDIDATE_NOTES, "Candidate Notes", "D4", 7000, 5000, "L", vbNullString, True, True, REPORT_CONDITIONAL_NONEMPTY
    AppendReportFieldSpec s, FIELD_MANUAL_REVIEW_NOTES, "Manual Review Notes", "D4", 7000, 5000, "L", vbNullString, True, True, REPORT_CONDITIONAL_NONEMPTY
    CandidatesReportPresentationSpec = s
End Function

Private Function ActiveMacrosReportPresentationSpec() As String
    Dim s As String
    AppendReportFieldSpec s, FIELD_MACRO_KEY, "Macro Key", "P", 4500, 3300, "L"
    AppendReportFieldSpec s, FIELD_MACRO_TITLE, "Title", "P", 9500, 6800, "L"
    AppendReportFieldSpec s, FIELD_ACTIVE, "Active", "D1", 1800, 1400, "C", "YESNO"
    AppendReportFieldSpec s, FIELD_SELECTED_FOR_RUNNER_MACRO_GENERATION, "Generate", "D1", 2200, 1700, "C", "YESNO"
    AppendReportFieldSpec s, FIELD_MACRO_ACTION, "Action", "D1", 3000, 2300, "L"
    AppendReportFieldSpec s, FIELD_CONFIGURATION_STATUS, "Config Status", "D1", 3400, 2600, "L"
    AppendReportFieldSpec s, FIELD_LAST_GENERATED_RESULT, "Generation", "D1", 3300, 2500, "L"
    AppendReportFieldSpec s, FIELD_TARGET_PROCEDURE_NAME, "Target Procedure", "D2", 5000, 3700, "L"
    AppendReportFieldSpec s, FIELD_TARGET_OBJECT_NAME, "Target Object", "D2", 4500, 3400, "L"
    AppendReportFieldSpec s, FIELD_CONFIG_OWNER_DATABASE_ROLE, "Owner Role", "D2", 3000, 2300, "L"
    AppendReportFieldSpec s, FIELD_GENERATED_MACRO_OBJECT_NAME, "Generated Macro", "D3", 4800, 3600, "L"
    AppendReportFieldSpec s, FIELD_NAVIGATION_PANE_GROUP, "Navigation Group", "D3", 4300, 3200, "L"
    AppendReportFieldSpec s, FIELD_RIBBON_TAB_LABEL, "Ribbon Tab", "D3", 3200, 2400, "L"
    AppendReportFieldSpec s, FIELD_RIBBON_GROUP_LABEL, "Ribbon Group", "D4", 5200, 3900, "L"
    AppendReportFieldSpec s, FIELD_RIBBON_TITLE, "Ribbon Title", "D4", 8500, 6100, "L"
    AppendReportFieldSpec s, FIELD_OBJECT_DESCRIPTION, "Description", "D5", 14000, 7000, "L", vbNullString, True, True, REPORT_CONDITIONAL_NONEMPTY
    ActiveMacrosReportPresentationSpec = s
End Function

Private Function DefaultMacrosReportPresentationSpec() As String
    Dim s As String
    AppendReportFieldSpec s, FIELD_SYSTEM_MACRO_DESIGNATOR, "Designator", "P", 2600, 2000, "L"
    AppendReportFieldSpec s, FIELD_SYSTEM_MACRO_KEY, "System Macro Key", "P", 4500, 3400, "L"
    AppendReportFieldSpec s, FIELD_DEFAULT_TITLE_TEXT, "Default Title", "P", 6800, 5000, "L"
    AppendReportFieldSpec s, FIELD_MACRO_CATEGORY_KEY, "Category", "D1", 3000, 2300, "L"
    AppendReportFieldSpec s, FIELD_MACRO_SUBCATEGORY_KEY, "Subcategory", "D1", 3200, 2400, "L"
    AppendReportFieldSpec s, FIELD_SYSTEM_MACRO_STATUS, "Status", "D1", 2600, 2000, "L"
    AppendReportFieldSpec s, FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER, "Sort", "D1", 1800, 1400, "R", "INTEGER"
    AppendReportFieldSpec s, FIELD_SYSTEM_MACRO_GROUP_KEY, "Macro Group", "D1", 3200, 2400, "L"
    AppendReportFieldSpec s, FIELD_RUN_CODE_WRAPPER_NAME, "RunCode Wrapper", "D2", 4700, 3500, "L"
    AppendReportFieldSpec s, "OwnerDatabaseRole", "Owner Role", "D2", 2800, 2100, "L"
    AppendReportFieldSpec s, FIELD_DEFAULT_SOURCE, "Default Source", "D2", 3500, 2700, "L"
    AppendReportFieldSpec s, FIELD_MACRO_TITLE_OVERRIDE, "Title Override", "D2", 3000, 2300, "L"
    AppendReportFieldSpec s, "RibbonGroupIdentity", "Ribbon Group", "D3", 6500, 4700, "L"
    AppendReportFieldSpec s, FIELD_RIBBON_TITLE, "Ribbon Title", "D3", 7500, 5400, "L"
    AppendReportFieldSpec s, FIELD_SYSTEM_MACRO_TARGET_TABLE_NAMES, "Target Tables", "D4", 14000, 7000, "L", vbNullString, True, True, REPORT_CONDITIONAL_NONEMPTY
    AppendReportFieldSpec s, FIELD_SYSTEM_MACRO_TARGET_QUERY_NAMES, "Target Queries", "D5", 14000, 7000, "L", vbNullString, True, True, REPORT_CONDITIONAL_NONEMPTY
    AppendReportFieldSpec s, FIELD_MACRO_DESCRIPTION, "Description", "D6", 14000, 7000, "L", vbNullString, True
    AppendReportFieldSpec s, FIELD_MANAGED_OBJECT_DEPENDENCY_NOTES, "Notes", "D7", 14000, 7000, "L", vbNullString, True, True, REPORT_CONDITIONAL_NONEMPTY
    DefaultMacrosReportPresentationSpec = s
End Function

Private Function MacroConfigurationReportPresentationSpec() As String
    Dim s As String
    AppendReportFieldSpec s, FIELD_MACRO_KEY, "Macro Key", "P", 4500, 3300, "L"
    AppendReportFieldSpec s, FIELD_MACRO_TITLE, "Title", "P", 9500, 6800, "L"
    AppendReportFieldSpec s, FIELD_CONFIGURATION_STATUS, "Config Status", "D1", 3400, 2600, "L"
    AppendReportFieldSpec s, FIELD_ACTIVE, "Active", "D1", 1800, 1400, "C", "YESNO"
    AppendReportFieldSpec s, FIELD_SELECTED_FOR_RUNNER_MACRO_GENERATION, "Generate", "D1", 2200, 1700, "C", "YESNO"
    AppendReportFieldSpec s, FIELD_MACRO_ACTION, "Action", "D1", 3000, 2300, "L"
    AppendReportFieldSpec s, FIELD_LAST_VALIDATION_RESULT, "Validation", "D1", 3300, 2500, "L"
    AppendReportFieldSpec s, FIELD_TARGET_PROCEDURE_NAME, "Target Procedure", "D2", 4200, 3200, "L"
    AppendReportFieldSpec s, FIELD_TARGET_PROCEDURE_MODULE_NAME, "Target Module", "D2", 3500, 2700, "L"
    AppendReportFieldSpec s, FIELD_TARGET_OBJECT_NAME, "Target Object", "D2", 4000, 3000, "L"
    AppendReportFieldSpec s, FIELD_CONFIG_OWNER_DATABASE_ROLE, "Owner Role", "D2", 2500, 1900, "L"
    AppendReportFieldSpec s, FIELD_LAST_GENERATED_RESULT, "Generation", "D3", 3200, 2400, "L"
    AppendReportFieldSpec s, FIELD_NAVIGATION_PANE_GROUP, "Navigation Group", "D3", 4200, 3200, "L"
    AppendReportFieldSpec s, FIELD_RIBBON_TAB_LABEL, "Ribbon Tab", "D3", 3000, 2300, "L"
    AppendReportFieldSpec s, FIELD_RIBBON_GROUP_LABEL, "Ribbon Group", "D3", 3200, 2400, "L"
    AppendReportFieldSpec s, FIELD_RIBBON_TITLE, "Ribbon Title", "D4", 14000, 7000, "L"
    AppendReportFieldSpec s, FIELD_OBJECT_DESCRIPTION, "Description", "D5", 14000, 7000, "L", vbNullString, True, True, REPORT_CONDITIONAL_NONEMPTY
    MacroConfigurationReportPresentationSpec = s
End Function

Private Function HealthReviewReportPresentationSpec() As String
    Dim s As String
    AppendReportFieldSpec s, FIELD_REVIEW_AREA, "Review Area", "P", 3500, 2700, "L"
    AppendReportFieldSpec s, FIELD_ISSUE_CODE, "Issue Code", "P", 3500, 2700, "L"
    AppendReportFieldSpec s, FIELD_DISPLAY_SUBJECT, "Subject", "P", 7000, 5100, "L"
    AppendReportFieldSpec s, FIELD_OBJECT_TYPE, "Object Type", "D1", 3000, 2300, "L"
    AppendReportFieldSpec s, FIELD_OBJECT_NAME, "Object", "D1", 6500, 4700, "L"
    AppendReportFieldSpec s, FIELD_MACRO_KEY, "Macro Key", "D1", 4500, 3400, "L"
    AppendReportFieldSpec s, FIELD_CONFIGURATION_STATUS, "Config Status", "D2", 4000, 3000, "L"
    AppendReportFieldSpec s, FIELD_CANDIDATE_STATUS, "Candidate Status", "D2", 4000, 3000, "L"
    AppendReportFieldSpec s, FIELD_CREATED_ON, "Created", "D2", 5000, 3700, "L", "DATETIME"
    AppendReportFieldSpec s, FIELD_EXPECTED_VALUE, "Expected", "D3", 14000, 7000, "L", vbNullString, True
    AppendReportFieldSpec s, FIELD_ACTUAL_VALUE, "Actual", "D4", 14000, 7000, "L", vbNullString, True
    AppendReportFieldSpec s, FIELD_DETAILS, "Details", "D5", 14000, 7000, "L", vbNullString, True
    AppendReportFieldSpec s, FIELD_RECOMMENDED_ACTION, "Recommended Action", "D6", 14000, 7000, "L", vbNullString, True, True, REPORT_CONDITIONAL_NONEMPTY
    HealthReviewReportPresentationSpec = s
End Function

Private Function ModuleVerificationReportPresentationSpec() As String
    Dim s As String
    AppendReportFieldSpec s, FIELD_MODULE_NAME, "Module", "P", 6000, 4400, "L"
    AppendReportFieldSpec s, FIELD_COMPONENT_TYPE, "Component Type", "P", 3200, 2400, "L"
    AppendReportFieldSpec s, FIELD_VERIFICATION_RESULT, "Verification", "P", 4600, 3400, "L"
    AppendReportFieldSpec s, FIELD_EXPECTED_VERSION_ID, "Expected Version", "D1", 7000, 5000, "L"
    AppendReportFieldSpec s, FIELD_INSTALLED_VERSION_ID, "Installed Version", "D1", 7000, 5000, "L"
    AppendReportFieldSpec s, FIELD_EXPECTED_PREVIOUS_VERSION_ID, "Expected Previous", "D2", 7000, 5000, "L"
    AppendReportFieldSpec s, FIELD_INSTALLED_PREVIOUS_VERSION_ID, "Installed Previous", "D2", 7000, 5000, "L"
    AppendReportFieldSpec s, FIELD_UPDATE_BATCH_ID, "Update Batch", "D3", 5000, 3700, "L"
    AppendReportFieldSpec s, FIELD_EXPECTED_VERSION_TIMESTAMP, "Expected Timestamp", "D3", 4500, 3400, "L", "DATETIME"
    AppendReportFieldSpec s, FIELD_INSTALLED_VERSION_TIMESTAMP, "Installed Timestamp", "D3", 4500, 3400, "L", "DATETIME"
    AppendReportFieldSpec s, FIELD_VERIFICATION_RUN_ID, "Run ID", "D4", 3500, 2700, "L"
    AppendReportFieldSpec s, FIELD_VERIFIED_ON, "Verified", "D4", 4200, 3200, "L", "DATETIME"
    AppendReportFieldSpec s, FIELD_MANAGED_OBJECT_DEPENDENCY_IS_REQUIRED, "Required", "D4", 2500, 1900, "C", "YESNO"
    AppendReportFieldSpec s, FIELD_BASELINE_ACCEPTED, "Baseline Accepted", "D4", 3000, 2300, "C", "YESNO"
    AppendReportFieldSpec s, FIELD_VERIFICATION_NOTES, "Verification Notes", "D5", 7000, 5000, "L", vbNullString, True, True, REPORT_CONDITIONAL_NONEMPTY
    AppendReportFieldSpec s, FIELD_EXPECTED_CHANGE_DESCRIPTION, "Expected Change", "D5", 7000, 5000, "L", vbNullString, True, True, REPORT_CONDITIONAL_NONEMPTY
    ModuleVerificationReportPresentationSpec = s
End Function

Private Function ManagedObjectsReportPresentationSpec() As String
    Dim s As String
    AppendReportFieldSpec s, FIELD_OBJECT_NAME, "Object", "P", 7000, 5000, "L"
    AppendReportFieldSpec s, FIELD_CHAT_GPTUI_REPORT, "ChatGPT UI Report", "D1", 2800, 2100, "C", "YESNO"
    AppendReportFieldSpec s, FIELD_CATEGORY_NAME, "Category", "P", 4000, 3000, "L"
    AppendReportFieldSpec s, FIELD_MANAGED_OBJECT_STATUS, "Object Status", "P", 3000, 2300, "L"
    AppendReportFieldSpec s, FIELD_CATALOG_REVIEW_STATUS, "Catalog Status", "D1", 3500, 2700, "L"
    AppendReportFieldSpec s, FIELD_RESOLVED_SETUP_STAGE, "Setup Stage", "D1", 3500, 2700, "L"
    AppendReportFieldSpec s, FIELD_LAST_SETUP_RESULT, "Setup Result", "D1", 3500, 2700, "L"
    AppendReportFieldSpec s, FIELD_RESOLVED_CONFIG_OWNER_DATABASE_ROLE, "Owner Role", "D1", 3500, 2700, "L"
    AppendReportFieldSpec s, FIELD_RESOLVED_SETUP_ACTION, "Setup Action", "D2", 4500, 3400, "L"
    AppendReportFieldSpec s, FIELD_CATEGORY_CODE, "Category Code", "D2", 3500, 2700, "L"
    AppendReportFieldSpec s, FIELD_RESOLVED_NAVIGATION_PANE_GROUP_SETTING_KEY, "Navigation Group Setting", "D2", 6000, 4400, "L"
    AppendReportFieldSpec s, FIELD_RIBBON_TAB_LABEL, "Ribbon Tab", "D3", 4000, 3000, "L"
    AppendReportFieldSpec s, FIELD_RIBBON_GROUP_LABEL, "Ribbon Group", "D3", 4500, 3400, "L"
    AppendReportFieldSpec s, FIELD_RIBBON_TITLE, "Ribbon Title", "D3", 5500, 4100, "L"
    AppendReportFieldSpec s, FIELD_LAST_SETUP_MESSAGE, "Last Setup Message", "D4", 14000, 7000, "L", vbNullString, True, True, REPORT_CONDITIONAL_NONEMPTY
    AppendReportFieldSpec s, FIELD_OBJECT_DESCRIPTION, "Description", "D5", 14000, 7000, "L", vbNullString, True
    AppendReportFieldSpec s, FIELD_MANAGED_OBJECT_DEPENDENCY_NOTES, "Notes", "D6", 14000, 7000, "L", vbNullString, True, True, REPORT_CONDITIONAL_NONEMPTY
    ManagedObjectsReportPresentationSpec = s
End Function

Private Function AddInReferenceInventoryReportPresentationSpec() As String
    Dim s As String
    AppendReportFieldSpec s, FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME, "Name", "P", 6000, 4400, "L"
    AppendReportFieldSpec s, FIELD_PROJECT_NAME, "Project", "P", 4000, 3000, "L"
    AppendReportFieldSpec s, FIELD_REFERENCE_NAME, "Reference", "P", 4000, 3000, "L"
    AppendReportFieldSpec s, FIELD_REFERENCE_KIND, "Kind", "D1", 3000, 2300, "L"
    AppendReportFieldSpec s, FIELD_REFERENCE_STATE, "State", "D1", 3000, 2300, "L"
    AppendReportFieldSpec s, FIELD_REFERENCE_VERSION, "Version", "D1", 3000, 2300, "L"
    AppendReportFieldSpec s, FIELD_REFERENCE_IDENTITY, "Reference Identity", "D1", 5000, 3700, "L"
    AppendReportFieldSpec s, FIELD_GUID, "GUID", "D2", 6500, 4700, "L"
    AppendReportFieldSpec s, FIELD_IS_REGISTERED, "Registered", "D2", 2500, 1900, "C", "YESNO"
    AppendReportFieldSpec s, FIELD_IS_CONNECTED, "Connected", "D2", 2500, 1900, "C", "YESNO"
    AppendReportFieldSpec s, FIELD_IS_LOADED, "Loaded", "D2", 2200, 1700, "C", "YESNO"
    AppendReportFieldSpec s, FIELD_MODULE_COUNT, "Modules", "D3", 2200, 1700, "R", "INTEGER"
    AppendReportFieldSpec s, FIELD_PUBLIC_PROCEDURE_COUNT, "Public Procedures", "D3", 3500, 2600, "R", "INTEGER"
    AppendReportFieldSpec s, FIELD_PROCEDURE_CONFLICT_COUNT, "Conflicts", "D3", 2800, 2100, "R", "INTEGER"
    AppendReportFieldSpec s, FIELD_DEPENDENCY_COUNT, "Dependencies", "D3", 2800, 2100, "R", "INTEGER"
    AppendReportFieldSpec s, FIELD_LAST_SCANNED_ON, "Last Scanned", "D3", 3500, 2700, "L", "DATETIME"
    AppendReportFieldSpec s, FIELD_FILE_PATH, "File Path", "D4", 14000, 7000, "L", vbNullString, True
    AppendReportFieldSpec s, FIELD_DIAGNOSTIC_TEXT, "Diagnostic", "D5", 14000, 7000, "L", vbNullString, True, True, REPORT_CONDITIONAL_NONEMPTY
    AddInReferenceInventoryReportPresentationSpec = s
End Function

Private Function ManagedReportDefinitions() As Collection
    ' AUTHORITATIVE REPORT CONSTRUCTION/PRESENTATION CACHE: this immutable ordered collection owns report source, grouping, layout mode, and field-presentation specifications; Defaults owns canonical report logical
    ' IDs and titles; DB_DEFINITIONS owns physical report and field/output names.
    ' Lifetime is the loaded VBA project; no runtime invalidation is required because the code-owned metadata changes only with a new module build. Catalog identity/placement remains Defaults-owned.
    Static rows As Collection
    If rows Is Nothing Then
        Set rows = New Collection
        rows.Add Array(CANON_REPORT_LOGICAL_ID_SETTINGS, CANON_REPORT_TITLE_SETTINGS, ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Settings_Resolved"), FIELD_SETTING_GROUP, REPORT_SOURCE_QUERY, FIELD_SETTING_GROUP, REPORT_LAYOUT_RECORD_DETAIL, SettingsReportPresentationSpec())
        rows.Add Array(CANON_REPORT_LOGICAL_ID_CANDIDATES, CANON_REPORT_TITLE_CANDIDATES, ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Config_CandidateReview"), FIELD_MACRO_CATEGORY_KEY, REPORT_SOURCE_QUERY, FIELD_MACRO_CATEGORY_KEY, REPORT_LAYOUT_RECORD_DETAIL, CandidatesReportPresentationSpec())
        rows.Add Array(CANON_REPORT_LOGICAL_ID_ACTIVE_MACROS, CANON_REPORT_TITLE_ACTIVE_MACROS, ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Macro_Report_Source"), FIELD_MACRO_CATEGORY_KEY, REPORT_SOURCE_QUERY, FIELD_MACRO_CATEGORY_KEY, REPORT_LAYOUT_RECORD_DETAIL, ActiveMacrosReportPresentationSpec())
        rows.Add Array(CANON_REPORT_LOGICAL_ID_DEFAULT_MACROS, CANON_REPORT_TITLE_DEFAULT_MACROS, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SystemMacros"), FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER, REPORT_SOURCE_TABLE, vbNullString, REPORT_LAYOUT_RECORD_DETAIL, DefaultMacrosReportPresentationSpec())
        rows.Add Array(CANON_REPORT_LOGICAL_ID_MACRO_CONFIGURATION, CANON_REPORT_TITLE_MACRO_CONFIGURATION, ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Macro_Report_Source"), FIELD_MACRO_CATEGORY_KEY, REPORT_SOURCE_QUERY, FIELD_MACRO_CATEGORY_KEY, REPORT_LAYOUT_RECORD_DETAIL, MacroConfigurationReportPresentationSpec())
        rows.Add Array(CANON_REPORT_LOGICAL_ID_HEALTH_REVIEW, CANON_REPORT_TITLE_HEALTH_REVIEW, ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_System_HealthReview"), FIELD_SEVERITY, REPORT_SOURCE_QUERY, FIELD_SEVERITY, REPORT_LAYOUT_RECORD_DETAIL, HealthReviewReportPresentationSpec())
        rows.Add Array(CANON_REPORT_LOGICAL_ID_MODULE_VERIFICATION, CANON_REPORT_TITLE_MODULE_VERIFICATION, ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_ModuleUpdates_Current"), FIELD_MODULE_NAME, REPORT_SOURCE_QUERY, vbNullString, REPORT_LAYOUT_RECORD_DETAIL, ModuleVerificationReportPresentationSpec())
        rows.Add Array(CANON_REPORT_LOGICAL_ID_MANAGED_OBJECTS, CANON_REPORT_TITLE_MANAGED_OBJECTS, ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_ManagedObjects_Review"), FIELD_RESOLVED_OBJECT_TYPE, REPORT_SOURCE_QUERY, FIELD_RESOLVED_OBJECT_TYPE, REPORT_LAYOUT_RECORD_DETAIL, ManagedObjectsReportPresentationSpec())
        rows.Add Array(CANON_REPORT_LOGICAL_ID_ADDIN_REFERENCE_INVENTORY, CANON_REPORT_TITLE_ADDIN_REFERENCE_INVENTORY, ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Reference_Report_Source"), FIELD_SOURCE_KIND, REPORT_SOURCE_QUERY, FIELD_SOURCE_KIND, REPORT_LAYOUT_RECORD_DETAIL, AddInReferenceInventoryReportPresentationSpec())
    End If
    Set ManagedReportDefinitions = rows
End Function

Public Function ManagedReportDependencyDefinitionsForSetup() As Collection
    'Read-only Phase-E projection for F1. Reports remains the sole authority for canonical report/source definitions; F1 consumes this projection without mutating report state.
    Dim rows As New Collection
    Dim definitions As Collection
    Dim definitionData As Variant
    Set definitions = ManagedReportDefinitions()
    For Each definitionData In definitions
        rows.Add Array(ManagedReportPhysicalName(CStr(definitionData(REPORT_DEF_LOGICAL_NAME))), CStr(definitionData(REPORT_DEF_SOURCE_NAME)), CStr(definitionData(REPORT_DEF_SOURCE_KIND)))
    Next definitionData
    Set ManagedReportDependencyDefinitionsForSetup = rows
End Function

Private Function ManagedReportPhysicalName(ByVal logicalName As String) As String
    ' Wrapper purpose: Private Reports-domain adapter that binds the generic managed-object name resolver to the Report naming category.
    ' Why retained: Retained to prevent REPORTS callers from supplying the wrong naming category and to keep Report-domain intent explicit at each call site.
    ManagedReportPhysicalName = ManagedObjectPhysicalName(NAMING_CATEGORY_REPORT, logicalName)
End Function

' ===============================================================================================================================================================================
' REPORT RECORD-SOURCE / PRESENTATION RESOLUTION PROCEDURES
' ===============================================================================================================================================================================

Private Sub ResetReportProofIfContextChanged()
    Dim activeTraceRunID As String
    Dim databaseName As String
    activeTraceRunID = GetActiveTraceRunID()
    databaseName = Trim$(GetHostDb().Name)
    If StrComp(mReportProofTraceRunID, activeTraceRunID, vbBinaryCompare) <> 0 Or StrComp(mReportProofDatabaseName, databaseName, vbTextCompare) <> 0 Then
        mReportProofTraceRunID = activeTraceRunID
        mReportProofDatabaseName = databaseName
        mReportPreflightToken = vbNullString
        mReportBuildToken = vbNullString
        mReportReconciledToken = vbNullString
        ResetReportSourceProjectionCaches
        Set mReportPresentationCacheByReport = Nothing
    End If
End Sub

Private Sub ResetReportSourceProjectionCaches()
    ' Mutable report-source schema can change after Setup responsiveness/event boundaries; discard only the caches derived from current source fields while retaining immutable parsed presentation definitions.
    Set mReportSourceFieldIndexByName = Nothing
    Set mReportSelectedFieldsByReport = Nothing
End Sub

Private Function ReportContextToken(ByVal purposeText As String, ByVal payloadText As String) As String
    ReportContextToken = purposeText & "|" & StableTextHashHex(REPORT_LAYOUT_VERSION & "|" & Trim$(GetHostDb().Name) & "|" & payloadText)
End Function

Private Function MacroReportPresentationRecordSourceSql(Optional ByVal fieldContractOnly As Boolean = False) As String
    ' PHASE-E PRESENTATION ONLY: enrich the A3 Config-only base projection with the already-complete Phase-B MAC placement rows. Both join predicates are field-to-field identities so ACE never receives a constant
    ' predicate inside the LEFT JOIN ON clause. This SQL is assigned directly to the two physical Reports; no saved query or second data authority is created here.
    Dim sqlText As String
    Dim baseSource As String
    Dim managedObjectsTable As String
    baseSource = ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Macro_Report_Source")
    managedObjectsTable = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID)
    sqlText = "SELECT q." & SQL_FIELD_ID & ",q." & SQL_FIELD_MACRO_KEY & ",q." & SQL_FIELD_MACRO_TITLE & ",q." & SQL_FIELD_MACRO_CATEGORY_KEY & ",q." & SQL_FIELD_MACRO_SUBCATEGORY_KEY & ",q." & SQL_FIELD_CONFIGURATION_STATUS & ",q." & SQL_FIELD_ACTIVE & ",q." & SQL_FIELD_SELECTED_FOR_RUNNER_MACRO_GENERATION & ",q." & SQL_FIELD_MACRO_ACTION & ",q." & SQL_FIELD_CONFIG_OWNER_DATABASE_ROLE & ","
    sqlText = sqlText & "q." & SQL_FIELD_TARGET_PROCEDURE_NAME & ",q." & SQL_FIELD_TARGET_PROCEDURE_MODULE_NAME & ",q." & SQL_FIELD_TARGET_OBJECT_NAME & ",q." & SQL_FIELD_LAST_VALIDATION_RESULT & ",q." & SQL_FIELD_LAST_GENERATED_RESULT & ",q." & SQL_FIELD_GENERATED_MACRO_OBJECT_NAME & ",q." & SQL_FIELD_NAVIGATION_PANE_GROUP & ",q." & SQL_FIELD_OBJECT_DESCRIPTION & ","
    sqlText = sqlText & "mo." & SQL_FIELD_RIBBON_TAB_LABEL & " AS " & SQL_FIELD_RIBBON_TAB_LABEL & ",mo." & SQL_FIELD_RIBBON_GROUP_LABEL & " AS " & SQL_FIELD_RIBBON_GROUP_LABEL & ",mo." & SQL_FIELD_RIBBON_TITLE & " AS " & SQL_FIELD_RIBBON_TITLE & ",mo." & SQL_FIELD_RIBBON_STATIC_TAG & " AS " & SQL_FIELD_RIBBON_STATIC_TAG & ",mo." & SQL_FIELD_NAVIGATION_PANE_GROUP_SETTING_KEY & " AS " & SQL_FIELD_RESOLVED_NAVIGATION_PANE_GROUP_SETTING_KEY & " "
    sqlText = sqlText & "FROM [" & baseSource & "] AS q LEFT JOIN [" & managedObjectsTable & "] AS mo ON q." & SQL_FIELD_GENERATED_MACRO_OBJECT_NAME & "=mo." & SQL_FIELD_OBJECT_NAME & " AND q." & SQL_FIELD_REPORT_PLACEMENT_OBJECT_TYPE & "=mo." & SQL_FIELD_OBJECT_TYPE & " "
    If fieldContractOnly Then sqlText = sqlText & "WHERE 1=0 "
    sqlText = sqlText & "ORDER BY " & OrderByPolicyText("REPORT_MACRO_CONFIG_JOINED") & ";"
    MacroReportPresentationRecordSourceSql = sqlText
End Function

Private Function SettingsReportPresentationRecordSourceSql(Optional ByVal fieldContractOnly As Boolean = False) As String
    ' REPORT PRESENTATION ONLY: preserve the Settings customization-state display as a field in the effective report RecordSource so every generated report control remains directly bound to a preflight-compiled
    ' field.
    Dim sqlText As String
    Dim baseSource As String
    baseSource = ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Settings_Resolved")
    sqlText = "SELECT s.*, IIf(Nz(s." & SQL_FIELD_EFFECTIVE_VALUE & ",'')=Nz(s." & SQL_FIELD_MANAGED_OBJECT_FIELD_DEFAULT_VALUE & ",''),'Default','Customized') AS [CustomizationState] FROM [" & Replace(baseSource, "]", "]]") & "] AS s "
    If fieldContractOnly Then sqlText = sqlText & "WHERE 1=0 "
    sqlText = sqlText & "ORDER BY " & OrderByPolicyText("CATALOG_SETTINGS_FORM") & ";"
    SettingsReportPresentationRecordSourceSql = sqlText
End Function

Private Function DefaultMacrosReportPresentationRecordSourceSql(Optional ByVal fieldContractOnly As Boolean = False) As String
    ' REPORT PRESENTATION ONLY: expose the three preexisting derived Default Macros display values as effective RecordSource fields instead of report-control expressions. E1 compiles this SQL before E2 creates any
    ' report controls.
    Dim sqlText As String
    Dim baseSource As String
    baseSource = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SystemMacros")
    sqlText = "SELECT s.*, " & SharedDao_SqlQuote(CONFIG_OWNER_DATABASE_ROLE_CODE_DB) & " AS [OwnerDatabaseRole], SystemMacroRibbonGroupIdentity(s." & SQL_FIELD_SYSTEM_MACRO_KEY & ") AS [RibbonGroupIdentity], SystemRibbonTitleForKey(s." & SQL_FIELD_SYSTEM_MACRO_KEY & ") AS " & SQL_FIELD_RIBBON_TITLE & " FROM [" & Replace(baseSource, "]", "]]") & "] AS s "
    If fieldContractOnly Then sqlText = sqlText & "WHERE 1=0 "
    sqlText = sqlText & "ORDER BY " & OrderByPolicyText("CATALOG_SYSTEM_MACRO") & ";"
    DefaultMacrosReportPresentationRecordSourceSql = sqlText
End Function

Private Function ReportEffectiveRecordSourceSql(ByVal definitionData As Variant, Optional ByVal fieldContractOnly As Boolean = False) As String
    Select Case CStr(definitionData(REPORT_DEF_LOGICAL_NAME))
        Case CANON_REPORT_LOGICAL_ID_SETTINGS
            ReportEffectiveRecordSourceSql = SettingsReportPresentationRecordSourceSql(fieldContractOnly)
        Case CANON_REPORT_LOGICAL_ID_ACTIVE_MACROS, CANON_REPORT_LOGICAL_ID_MACRO_CONFIGURATION
            ReportEffectiveRecordSourceSql = MacroReportPresentationRecordSourceSql(fieldContractOnly)
        Case CANON_REPORT_LOGICAL_ID_DEFAULT_MACROS
            ReportEffectiveRecordSourceSql = DefaultMacrosReportPresentationRecordSourceSql(fieldContractOnly)
    End Select
End Function

Private Function ReportRecordSourceText(ByVal definitionData As Variant) As String
    ReportRecordSourceText = ReportEffectiveRecordSourceSql(definitionData)
    If Len(ReportRecordSourceText) = 0 Then ReportRecordSourceText = CStr(definitionData(REPORT_DEF_SOURCE_NAME))
End Function

Private Function ReportRecordSourceSignatureText(ByVal definitionData As Variant) As String
    Dim sourceName As String
    Dim recordSourceText As String
    sourceName = CStr(definitionData(REPORT_DEF_SOURCE_NAME))
    recordSourceText = ReportRecordSourceText(definitionData)
    ReportRecordSourceSignatureText = sourceName
    If StrComp(recordSourceText, sourceName, vbBinaryCompare) <> 0 Then ReportRecordSourceSignatureText = ReportRecordSourceSignatureText & "|Effective=" & StableTextHashHex(recordSourceText)
End Function

Private Function ReportSourceExists(ByVal sourceName As String, ByVal sourceKind As String) As Boolean
    Dim db As DAO.Database
    Set db = GetHostDb()
    Select Case UCase$(sourceKind)
        Case UCase$(REPORT_SOURCE_QUERY)
            ReportSourceExists = SharedDao_QueryExists(db, sourceName)
        Case UCase$(REPORT_SOURCE_TABLE)
            ReportSourceExists = SharedDao_TableExists(db, sourceName)
        Case Else
            Err.Raise vbObjectError + 9870, "ReportSourceExists", "Unsupported report source kind: " & sourceKind
    End Select
End Function

Private Function ReportSourceCacheKey(ByVal definitionData As Variant) As String
    Dim sourceName As String
    Dim recordSourceText As String
    sourceName = CStr(definitionData(REPORT_DEF_SOURCE_NAME))
    recordSourceText = ReportRecordSourceText(definitionData)
    ReportSourceCacheKey = sourceName & "|" & StableTextHashHex(recordSourceText)
End Function

Private Sub EnsureReportSourceFieldIndex(ByVal definitionData As Variant)
    ' AUTHORITATIVE CONSTRUCTION POINT: builds the case-insensitive effective RecordSource field-membership index once per source within the current source-validation cache lifetime.
    Dim fieldIndex As Object
    Dim rs As DAO.Recordset
    Dim fld As DAO.Field
    Dim sourceName As String
    Dim recordSourceText As String
    Dim cacheKey As String
    sourceName = CStr(definitionData(REPORT_DEF_SOURCE_NAME))
    recordSourceText = ReportRecordSourceText(definitionData)
    cacheKey = ReportSourceCacheKey(definitionData)
    If mReportSourceFieldIndexByName Is Nothing Then
        Set mReportSourceFieldIndexByName = CreateObject("Scripting.Dictionary")
        mReportSourceFieldIndexByName.CompareMode = vbTextCompare
    End If
    If mReportSourceFieldIndexByName.Exists(cacheKey) Then Exit Sub
    Set fieldIndex = CreateObject("Scripting.Dictionary")
    fieldIndex.CompareMode = vbTextCompare
    If StrComp(recordSourceText, sourceName, vbBinaryCompare) = 0 Then
        Set rs = GetHostDb().OpenRecordset("SELECT * FROM [" & Replace(sourceName, "]", "]]") & "] WHERE 1=0;", dbOpenSnapshot)
    Else
        ' Compile the effective Phase-E presentation SQL against zero rows so E1 verifies every derived/enriched display field before any physical report is created.
        Set rs = GetHostDb().OpenRecordset(ReportEffectiveRecordSourceSql(definitionData, True), dbOpenSnapshot)
    End If
    For Each fld In rs.fields
        If fieldIndex.Exists(fld.Name) = False Then fieldIndex.Add fld.Name, True
    Next fld
    rs.Close
    Set rs = Nothing
    Set mReportSourceFieldIndexByName(cacheKey) = fieldIndex
End Sub

Private Function ReportAvailableFieldIndex(ByVal definitionData As Variant) As Object
    EnsureReportSourceFieldIndex definitionData
    Set ReportAvailableFieldIndex = mReportSourceFieldIndexByName(ReportSourceCacheKey(definitionData))
End Function

Private Function ReportPresentationCacheKey(ByVal definitionData As Variant) As String
    ReportPresentationCacheKey = ManagedReportPhysicalName(CStr(definitionData(REPORT_DEF_LOGICAL_NAME))) & "|" & StableTextHashHex(CStr(definitionData(REPORT_DEF_PRESENTATION_SPEC)))
End Function

Private Sub EnsureReportPresentationCache(ByVal definitionData As Variant)
    ' AUTHORITATIVE CONSTRUCTION POINT: parses each immutable presentation definition once and stores its ordered all-specs, row projections, and detail-row order for reuse by E1/E2/currentness proof.
    Dim reportCache As Object
    Dim allSpecs As New Collection
    Dim detailRows As New Collection
    Dim detailSeen As Object
    Dim rowSpecs As Collection
    Dim rowText As Variant
    Dim parts As Variant
    Dim rawSpec As String
    Dim rowKey As String
    Dim cacheKey As String
    cacheKey = ReportPresentationCacheKey(definitionData)
    If mReportPresentationCacheByReport Is Nothing Then
        Set mReportPresentationCacheByReport = CreateObject("Scripting.Dictionary")
        mReportPresentationCacheByReport.CompareMode = vbTextCompare
    End If
    If mReportPresentationCacheByReport.Exists(cacheKey) Then Exit Sub
    rawSpec = CStr(definitionData(REPORT_DEF_PRESENTATION_SPEC))
    If Len(rawSpec) = 0 Then Err.Raise vbObjectError + 9889, "EnsureReportPresentationCache", "A managed report has no field-presentation specification." & vbCrLf & "Report: " & CStr(definitionData(REPORT_DEF_LOGICAL_NAME))
    Set reportCache = CreateObject("Scripting.Dictionary")
    reportCache.CompareMode = vbTextCompare
    Set detailSeen = CreateObject("Scripting.Dictionary")
    detailSeen.CompareMode = vbTextCompare
    Set rowSpecs = New Collection
    Set reportCache("ROW|" & REPORT_ROW_PRIMARY) = rowSpecs
    For Each rowText In Split(rawSpec, ";")
        If Len(CStr(rowText)) > 0 Then
            parts = Split(CStr(rowText), "~")
            If UBound(parts) - LBound(parts) + 1 <> REPORT_FIELD_SPEC_PART_COUNT Then Err.Raise vbObjectError + 9890, "EnsureReportPresentationCache", "Malformed report field-presentation specification." & vbCrLf & "Report: " & CStr(definitionData(REPORT_DEF_LOGICAL_NAME)) & vbCrLf & "Specification: " & CStr(rowText)
            allSpecs.Add parts
            If ReportSpecBoolean(parts(REPORT_FIELD_VISIBLE)) Then
                rowKey = CStr(parts(REPORT_FIELD_ROW_KEY))
                If reportCache.Exists("ROW|" & rowKey) = False Then
                    Set rowSpecs = New Collection
                    Set reportCache("ROW|" & rowKey) = rowSpecs
                End If
                Set rowSpecs = reportCache("ROW|" & rowKey)
                rowSpecs.Add parts
                If StrComp(rowKey, REPORT_ROW_PRIMARY, vbBinaryCompare) <> 0 Then
                    If detailSeen.Exists(rowKey) = False Then
                        detailRows.Add rowKey
                        detailSeen.Add rowKey, True
                    End If
                End If
            End If
        End If
    Next rowText
    Set reportCache("ALL") = allSpecs
    Set reportCache("DETAIL_ROWS") = detailRows
    Set mReportPresentationCacheByReport(cacheKey) = reportCache
End Sub

Private Function ReportPresentationSpecs(ByVal definitionData As Variant) As Collection
    Dim reportCache As Object
    EnsureReportPresentationCache definitionData
    Set reportCache = mReportPresentationCacheByReport(ReportPresentationCacheKey(definitionData))
    Set ReportPresentationSpecs = reportCache("ALL")
End Function

Private Function ReportSpecsForRow(ByVal definitionData As Variant, ByVal rowKey As String) As Collection
    Dim reportCache As Object
    Dim rows As Collection
    Dim normalizedRowKey As String
    EnsureReportPresentationCache definitionData
    normalizedRowKey = rowKey
    Set reportCache = mReportPresentationCacheByReport(ReportPresentationCacheKey(definitionData))
    If reportCache.Exists("ROW|" & normalizedRowKey) = False Then
        Set rows = New Collection
        Set reportCache("ROW|" & normalizedRowKey) = rows
    End If
    Set ReportSpecsForRow = reportCache("ROW|" & normalizedRowKey)
End Function

Private Function ReportDetailRowKeys(ByVal definitionData As Variant) As Collection
    Dim reportCache As Object
    EnsureReportPresentationCache definitionData
    Set reportCache = mReportPresentationCacheByReport(ReportPresentationCacheKey(definitionData))
    Set ReportDetailRowKeys = reportCache("DETAIL_ROWS")
End Function

Private Function ReportSpecBoolean(ByVal specValue As Variant) As Boolean
    ReportSpecBoolean = (StrComp(CStr(specValue), "1", vbBinaryCompare) = 0 Or StrComp(CStr(specValue), "TRUE", vbTextCompare) = 0 Or StrComp(CStr(specValue), "YES", vbTextCompare) = 0)
End Function

Private Function ReportSelectedFields(ByVal definitionData As Variant) As Collection
    ' MUTABLE-SOURCE-DERIVED CACHE: selected is the displayed source-field projection for this report within the current source-validation cache lifetime and is rebuilt after cache invalidation.
    Dim selected As New Collection
    Dim specs As Collection
    Dim specData As Variant
    Dim reportName As String
    Dim availableIndex As Object
    Dim selectedIndex As Object
    Dim fieldName As String
    reportName = ManagedReportPhysicalName(CStr(definitionData(REPORT_DEF_LOGICAL_NAME)))
    If mReportSelectedFieldsByReport Is Nothing Then
        Set mReportSelectedFieldsByReport = CreateObject("Scripting.Dictionary")
        mReportSelectedFieldsByReport.CompareMode = vbTextCompare
    End If
    If mReportSelectedFieldsByReport.Exists(reportName) Then
        Set ReportSelectedFields = mReportSelectedFieldsByReport(reportName)
        Exit Function
    End If
    Set availableIndex = ReportAvailableFieldIndex(definitionData)
    Set selectedIndex = CreateObject("Scripting.Dictionary")
    selectedIndex.CompareMode = vbTextCompare
    Set specs = ReportPresentationSpecs(definitionData)
    For Each specData In specs
        If ReportSpecBoolean(specData(REPORT_FIELD_VISIBLE)) Then
            fieldName = CStr(specData(REPORT_FIELD_NAME))
            If Len(fieldName) = 0 Then Err.Raise vbObjectError + 9891, "ReportSelectedFields", "A visible report field has a blank source field name." & vbCrLf & "Report: " & reportName
            If availableIndex.Exists(fieldName) = False Then Err.Raise vbObjectError + 9892, "ReportSelectedFields", "A canonical report presentation field is not exposed by its authoritative RecordSource." & vbCrLf & "Report: " & reportName & vbCrLf & "Field: " & fieldName
            If selectedIndex.Exists(fieldName) Then Err.Raise vbObjectError + 9893, "ReportSelectedFields", "A report presentation field is defined more than once." & vbCrLf & "Report: " & reportName & vbCrLf & "Field: " & fieldName
            selected.Add fieldName
            selectedIndex.Add fieldName, True
        End If
    Next specData
    If selected.count = 0 Then Err.Raise vbObjectError + 9871, "ReportSelectedFields", "A report presentation specification exposed no visible source fields." & vbCrLf & "Source: " & CStr(definitionData(REPORT_DEF_SOURCE_NAME))
    Set mReportSelectedFieldsByReport(reportName) = selected
    Set ReportSelectedFields = selected
End Function

Private Function ReportSelectedFieldPayload(ByVal fields As Collection) As String
    Dim fieldName As Variant
    For Each fieldName In fields
        If Len(ReportSelectedFieldPayload) > 0 Then ReportSelectedFieldPayload = ReportSelectedFieldPayload & "|"
        ReportSelectedFieldPayload = ReportSelectedFieldPayload & CStr(fieldName)
    Next fieldName
End Function

Private Function ReportFilterExpression(ByVal definitionData As Variant) As String
    If StrComp(CStr(definitionData(REPORT_DEF_LOGICAL_NAME)), CANON_REPORT_LOGICAL_ID_ACTIVE_MACROS, vbBinaryCompare) = 0 Then ReportFilterExpression = SQL_FIELD_ACTIVE & "=True"
End Function

Private Function ReportResolvedGroupField(ByVal definitionData As Variant) As String
    Dim requestedGroupField As String
    Dim availableIndex As Object
    requestedGroupField = CStr(definitionData(REPORT_DEF_GROUP_FIELD))
    If Len(requestedGroupField) = 0 Then Exit Function
    Set availableIndex = ReportAvailableFieldIndex(definitionData)
    If availableIndex.Exists(requestedGroupField) = False Then Err.Raise vbObjectError + 9894, "ReportResolvedGroupField", "The canonical report group field is not exposed by the authoritative RecordSource." & vbCrLf & "Report: " & CStr(definitionData(REPORT_DEF_LOGICAL_NAME)) & vbCrLf & "Field: " & requestedGroupField
    ReportResolvedGroupField = requestedGroupField
End Function

Private Function ReportExpectedSignature(ByVal definitionData As Variant, ByVal selectedFields As Collection) As String
    ReportExpectedSignature = "ADC_REPORT|" & REPORT_LAYOUT_VERSION & "|" & StableTextHashHex(ManagedReportPhysicalName(CStr(definitionData(REPORT_DEF_LOGICAL_NAME))) & "|" & CStr(definitionData(REPORT_DEF_TITLE)) & "|" & ReportRecordSourceSignatureText(definitionData) & "|" & ReportSelectedFieldPayload(selectedFields) & "|Layout=" & CStr(definitionData(REPORT_DEF_LAYOUT_MODE)) & "|Presentation=" & StableTextHashHex(CStr(definitionData(REPORT_DEF_PRESENTATION_SPEC))) & "|Filter=" & ReportFilterExpression(definitionData) & "|Sort=" & CStr(definitionData(REPORT_DEF_SORT_FIELD)) & "|Group=" & ReportResolvedGroupField(definitionData))
End Function

Private Function ReportScaledSpecWidths(ByVal rowSpecs As Collection, Optional ByVal interFieldGap As Long = 0) As Variant
    Dim widths() As Long
    Dim mins() As Long
    Dim index As Long
    Dim totalWidth As Long
    Dim minimumTotalWidth As Long
    Dim availableWidth As Long
    Dim flexibleWidth As Long
    Dim desiredExtraWidth As Long
    Dim desiredFieldExtraWidth As Long
    Dim specData As Variant
    If rowSpecs.count <= 0 Then Err.Raise vbObjectError + 9895, "ReportScaledSpecWidths", "Cannot size an empty report row."
    If interFieldGap < 0 Then Err.Raise vbObjectError + 9910, "ReportScaledSpecWidths", "A report row cannot use a negative inter-field gap."
    ReDim widths(0 To rowSpecs.count - 1)
    ReDim mins(0 To rowSpecs.count - 1)
    index = 0
    For Each specData In rowSpecs
        widths(index) = CLng(specData(REPORT_FIELD_PREFERRED_WIDTH))
        mins(index) = CLng(specData(REPORT_FIELD_MIN_WIDTH))
        If widths(index) <= 0 Or mins(index) <= 0 Or mins(index) > widths(index) Then Err.Raise vbObjectError + 9896, "ReportScaledSpecWidths", "Invalid preferred/minimum report field width." & vbCrLf & "Field: " & CStr(specData(REPORT_FIELD_NAME))
        totalWidth = totalWidth + widths(index)
        minimumTotalWidth = minimumTotalWidth + mins(index)
        index = index + 1
    Next specData
    availableWidth = REPORT_PAGE_WIDTH - REPORT_LEFT_MARGIN - REPORT_RIGHT_MARGIN - (interFieldGap * (rowSpecs.count - 1))
    If availableWidth <= 0 Then Err.Raise vbObjectError + 9911, "ReportScaledSpecWidths", "A report row has no usable width after declared inter-field spacing."
    If minimumTotalWidth > availableWidth Then Err.Raise vbObjectError + 9897, "ReportScaledSpecWidths", "A canonical report row cannot fit within printable width at its declared readable minimum widths."
    If totalWidth > availableWidth Then
        flexibleWidth = availableWidth - minimumTotalWidth
        desiredExtraWidth = totalWidth - minimumTotalWidth
        For index = LBound(widths) To UBound(widths)
            desiredFieldExtraWidth = widths(index) - mins(index)
            If desiredExtraWidth > 0 Then widths(index) = mins(index) + CLng(Fix((CDbl(desiredFieldExtraWidth) * CDbl(flexibleWidth)) / CDbl(desiredExtraWidth))) Else widths(index) = mins(index)
        Next index
    End If
    ReportScaledSpecWidths = widths
End Function

Private Function ReportDetailLabelWidth(ByVal segmentWidth As Long) As Long
    Dim labelWidth As Long
    labelWidth = CLng(Fix(CDbl(segmentWidth) * 0.32))
    If labelWidth > REPORT_MAX_DETAIL_LABEL_WIDTH Then labelWidth = REPORT_MAX_DETAIL_LABEL_WIDTH
    If labelWidth < REPORT_MIN_DETAIL_LABEL_WIDTH Then labelWidth = REPORT_MIN_DETAIL_LABEL_WIDTH
    If segmentWidth - labelWidth - REPORT_DETAIL_LABEL_GAP < REPORT_MIN_DETAIL_VALUE_WIDTH Then labelWidth = segmentWidth - REPORT_DETAIL_LABEL_GAP - REPORT_MIN_DETAIL_VALUE_WIDTH
    If labelWidth < 0 Then labelWidth = 0
    ReportDetailLabelWidth = labelWidth
End Function


Private Function ReportUsesRecordDetail(ByVal definitionData As Variant) As Boolean
    ReportUsesRecordDetail = (StrComp(CStr(definitionData(REPORT_DEF_LAYOUT_MODE)), REPORT_LAYOUT_RECORD_DETAIL, vbBinaryCompare) = 0)
End Function





Private Function ReportExpectedRecordControlCount(ByVal definitionData As Variant) As Long
    Dim visualRows As Collection
    Dim rowSpecs As Collection
    Dim rowKey As Variant
    ReportExpectedRecordControlCount = 2
    Set visualRows = ReportVisualRowKeys(definitionData)
    For Each rowKey In visualRows
        Set rowSpecs = ReportSpecsForRow(definitionData, CStr(rowKey))
        ReportExpectedRecordControlCount = ReportExpectedRecordControlCount + (rowSpecs.count * 2)
    Next rowKey
End Function

Private Function ReportHeaderSectionHeight(ByVal definitionData As Variant) As Long
    If ReportUsesRecordDetail(definitionData) Then
        ReportHeaderSectionHeight = REPORT_RECORD_HEADER_HEIGHT
    Else
        ReportHeaderSectionHeight = REPORT_PAGE_HEADER_HEIGHT
    End If
End Function

Private Function ReportVisualRowKeys(ByVal definitionData As Variant) As Collection
    Dim rows As New Collection
    Dim detailRows As Collection
    Dim rowKey As Variant
    rows.Add REPORT_ROW_PRIMARY
    Set detailRows = ReportDetailRowKeys(definitionData)
    For Each rowKey In detailRows
        rows.Add CStr(rowKey)
    Next rowKey
    Set ReportVisualRowKeys = rows
End Function

Private Function ReportRecordValueHeight(ByVal definitionData As Variant, ByVal rowKey As String) As Long
    Dim rowSpecs As Collection
    Dim specData As Variant
    If StrComp(rowKey, REPORT_ROW_PRIMARY, vbBinaryCompare) = 0 Then
        ReportRecordValueHeight = REPORT_RECORD_PRIMARY_VALUE_HEIGHT
        Exit Function
    End If
    Set rowSpecs = ReportSpecsForRow(definitionData, rowKey)
    For Each specData In rowSpecs
        If ReportSpecBoolean(specData(REPORT_FIELD_CAN_GROW)) Then
            ReportRecordValueHeight = REPORT_RECORD_GROWING_VALUE_HEIGHT
            Exit Function
        End If
    Next specData
    ReportRecordValueHeight = REPORT_RECORD_VALUE_HEIGHT
End Function

Private Function ReportDetailSectionHeight(ByVal definitionData As Variant) As Long
    Dim detailRows As Collection
    Dim visualRows As Collection
    Dim rowKey As Variant
    If ReportUsesRecordDetail(definitionData) Then
        Set visualRows = ReportVisualRowKeys(definitionData)
        ReportDetailSectionHeight = REPORT_RECORD_DETAIL_TOP
        For Each rowKey In visualRows
            ReportDetailSectionHeight = ReportDetailSectionHeight + REPORT_RECORD_LABEL_HEIGHT + REPORT_RECORD_LABEL_VALUE_GAP + ReportRecordValueHeight(definitionData, CStr(rowKey))
        Next rowKey
        If visualRows.count > 1 Then ReportDetailSectionHeight = ReportDetailSectionHeight + ((visualRows.count - 1) * REPORT_RECORD_ROW_GAP)
        ReportDetailSectionHeight = ReportDetailSectionHeight + REPORT_RECORD_BOTTOM_PADDING
    Else
        Set detailRows = ReportDetailRowKeys(definitionData)
        ReportDetailSectionHeight = REPORT_PRIMARY_CONTROL_TOP + REPORT_CONTROL_HEIGHT + REPORT_DETAIL_ROW_GAP
        If detailRows.count > 0 Then ReportDetailSectionHeight = ReportDetailSectionHeight + detailRows.count * (REPORT_CONTROL_HEIGHT + REPORT_DETAIL_ROW_GAP)
    End If
End Function
Private Function ReportTextAlignment(ByVal alignmentKey As String) As Long
    Select Case alignmentKey
        Case REPORT_ALIGN_CENTER
            ReportTextAlignment = REPORT_TEXT_ALIGN_CENTER
        Case REPORT_ALIGN_RIGHT
            ReportTextAlignment = REPORT_TEXT_ALIGN_RIGHT
        Case Else
            ReportTextAlignment = REPORT_TEXT_ALIGN_LEFT
    End Select
End Function

Private Function ReportFormatText(ByVal formatKey As String) As String
    Select Case formatKey
        Case REPORT_FORMAT_YESNO
            ReportFormatText = "Yes/No"
        Case REPORT_FORMAT_DATETIME
            ReportFormatText = "m/d/yyyy h:nn AM/PM"
        Case REPORT_FORMAT_DATE
            ReportFormatText = "m/d/yyyy"
        Case REPORT_FORMAT_INTEGER
            ReportFormatText = "0"
        Case Else
            ReportFormatText = formatKey
    End Select
End Function

Private Function ReportConditionalLabelFormat(ByVal captionText As String) As String
    ' NONEMPTY labels stay directly bound to the same source field as their value control. Access text formats use the second section for zero-length/Null values, so the label itself shrinks away without a report
    ' event or duplicate runtime state.
    ReportConditionalLabelFormat = Chr$(34) & Replace(captionText, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34) & ";" & Chr$(34) & Chr$(34)
End Function

Private Sub ApplyReportLabelStyle(ByVal ctl As Access.control, ByVal headerStyle As Boolean)
    ctl.FontName = "Segoe UI"
    ctl.fontBold = True
    ctl.fontSize = REPORT_FONT_SIZE_HEADER
    ctl.BackStyle = 1
    ctl.BorderStyle = 0
    If headerStyle Then
        ctl.BackColor = RGB(232, 235, 238)
        ctl.ForeColor = RGB(42, 47, 52)
    Else
        ctl.BackColor = RGB(255, 255, 255)
        ctl.ForeColor = RGB(82, 88, 94)
    End If
End Sub

Private Sub ApplyReportFieldStyle(ByVal ctl As Access.control, ByVal specData As Variant, ByVal isPrimary As Boolean)
    ctl.FontName = "Segoe UI"
    ctl.fontSize = REPORT_FONT_SIZE_DETAIL
    ctl.BorderStyle = 0
    ctl.BackStyle = 0
    ctl.TextAlign = ReportTextAlignment(CStr(specData(REPORT_FIELD_ALIGNMENT)))
    ctl.Format = ReportFormatText(CStr(specData(REPORT_FIELD_FORMAT)))
    If isPrimary Then
        ctl.canGrow = False
        ctl.CanShrink = False
    Else
        ctl.canGrow = ReportSpecBoolean(specData(REPORT_FIELD_CAN_GROW))
        ctl.CanShrink = True
    End If
End Sub


Private Sub ValidateRecordDetailRowGeometry(ByVal definitionData As Variant)
    Dim visualRows As Collection
    Dim rowSpecs As Collection
    Dim rowKey As Variant
    Dim widths As Variant
    Dim index As Long
    Dim leftPos As Long
    Dim previousRight As Long
    Dim topPos As Long
    Dim valueTop As Long
    Dim rowBottom As Long
    Dim previousBottom As Long
    Set visualRows = ReportVisualRowKeys(definitionData)
    topPos = REPORT_RECORD_DETAIL_TOP
    previousBottom = -1
    For Each rowKey In visualRows
        Set rowSpecs = ReportSpecsForRow(definitionData, CStr(rowKey))
        widths = ReportScaledSpecWidths(rowSpecs, REPORT_RECORD_COLUMN_GAP)
        If previousBottom >= 0 And topPos < previousBottom + REPORT_RECORD_ROW_GAP Then Err.Raise vbObjectError + 9912, "ValidateRecordDetailRowGeometry", "RECORD_DETAIL rows overlap vertically; growable controls require non-overlapping design rows." & vbCrLf & "Report: " & CStr(definitionData(REPORT_DEF_LOGICAL_NAME)) & vbCrLf & "Row: " & CStr(rowKey)
        leftPos = REPORT_LEFT_MARGIN
        previousRight = -1
        For index = LBound(widths) To UBound(widths)
            If previousRight >= 0 And leftPos < previousRight + REPORT_RECORD_COLUMN_GAP Then Err.Raise vbObjectError + 9921, "ValidateRecordDetailRowGeometry", "RECORD_DETAIL controls overlap horizontally." & vbCrLf & "Report: " & CStr(definitionData(REPORT_DEF_LOGICAL_NAME)) & vbCrLf & "Row: " & CStr(rowKey)
            If leftPos < REPORT_LEFT_MARGIN Or leftPos + CLng(widths(index)) > REPORT_PAGE_WIDTH - REPORT_RIGHT_MARGIN Then Err.Raise vbObjectError + 9922, "ValidateRecordDetailRowGeometry", "RECORD_DETAIL control exceeds the printable report width." & vbCrLf & "Report: " & CStr(definitionData(REPORT_DEF_LOGICAL_NAME)) & vbCrLf & "Row: " & CStr(rowKey)
            previousRight = leftPos + CLng(widths(index))
            leftPos = previousRight + REPORT_RECORD_COLUMN_GAP
        Next index
        valueTop = topPos + REPORT_RECORD_LABEL_HEIGHT + REPORT_RECORD_LABEL_VALUE_GAP
        rowBottom = valueTop + ReportRecordValueHeight(definitionData, CStr(rowKey))
        previousBottom = rowBottom
        topPos = rowBottom + REPORT_RECORD_ROW_GAP
    Next rowKey
End Sub

' ===============================================================================================================================================================================
' REPORT PHYSICAL CURRENTNESS / BUILD VERIFICATION PROCEDURES
' ===============================================================================================================================================================================

Private Sub ValidateReportPresentationDependencies(ByVal definitionData As Variant)
    Dim layoutMode As String
    Dim specs As Collection
    Dim detailRows As Collection
    Dim primarySpecs As Collection
    Dim rowSpecs As Collection
    Dim specData As Variant
    Dim rowKey As Variant
    Dim availableIndex As Object
    Dim sortField As String
    Dim conditionalMode As String
    Dim alignmentKey As String
    layoutMode = CStr(definitionData(REPORT_DEF_LAYOUT_MODE))
    If StrComp(layoutMode, REPORT_LAYOUT_TABULAR, vbBinaryCompare) <> 0 And StrComp(layoutMode, REPORT_LAYOUT_TABULAR_DETAIL, vbBinaryCompare) <> 0 And StrComp(layoutMode, REPORT_LAYOUT_RECORD_DETAIL, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 9898, "ValidateReportPresentationDependencies", "Unsupported report layout mode: " & layoutMode
    Set specs = ReportPresentationSpecs(definitionData)
    Set primarySpecs = ReportSpecsForRow(definitionData, REPORT_ROW_PRIMARY)
    If primarySpecs.count = 0 Then Err.Raise vbObjectError + 9899, "ValidateReportPresentationDependencies", "A managed report has no visible primary fields." & vbCrLf & "Report: " & CStr(definitionData(REPORT_DEF_LOGICAL_NAME))
    ' Calculate widths here for the helper's strict width-contract validation; this validation procedure does not retain the calculated array.
    If ReportUsesRecordDetail(definitionData) Then
        Call ReportScaledSpecWidths(primarySpecs, REPORT_RECORD_COLUMN_GAP)
        ValidateRecordDetailRowGeometry definitionData
    Else
        Call ReportScaledSpecWidths(primarySpecs)
    End If
    Set detailRows = ReportDetailRowKeys(definitionData)
    If StrComp(layoutMode, REPORT_LAYOUT_TABULAR, vbBinaryCompare) = 0 And detailRows.count > 0 Then Err.Raise vbObjectError + 9900, "ValidateReportPresentationDependencies", "TABULAR report metadata cannot contain secondary/detail rows."
    For Each rowKey In detailRows
        Set rowSpecs = ReportSpecsForRow(definitionData, CStr(rowKey))
        If ReportUsesRecordDetail(definitionData) Then
            Call ReportScaledSpecWidths(rowSpecs, REPORT_RECORD_COLUMN_GAP)
        Else
            Call ReportScaledSpecWidths(rowSpecs)
        End If
    Next rowKey
    For Each specData In specs
        If ReportSpecBoolean(specData(REPORT_FIELD_VISIBLE)) Then
            rowKey = CStr(specData(REPORT_FIELD_ROW_KEY))
            If StrComp(CStr(rowKey), REPORT_ROW_PRIMARY, vbBinaryCompare) <> 0 And Left$(CStr(rowKey), 1) <> "D" Then Err.Raise vbObjectError + 9901, "ValidateReportPresentationDependencies", "Invalid report row placement key." & vbCrLf & "Field: " & CStr(specData(REPORT_FIELD_NAME)) & vbCrLf & "Row: " & CStr(rowKey)
            alignmentKey = CStr(specData(REPORT_FIELD_ALIGNMENT))
            If alignmentKey <> REPORT_ALIGN_LEFT And alignmentKey <> REPORT_ALIGN_CENTER And alignmentKey <> REPORT_ALIGN_RIGHT Then Err.Raise vbObjectError + 9902, "ValidateReportPresentationDependencies", "Invalid report alignment metadata." & vbCrLf & "Field: " & CStr(specData(REPORT_FIELD_NAME))
            conditionalMode = CStr(specData(REPORT_FIELD_CONDITIONAL_MODE))
            If conditionalMode <> REPORT_CONDITIONAL_ALWAYS And conditionalMode <> REPORT_CONDITIONAL_NONEMPTY Then Err.Raise vbObjectError + 9903, "ValidateReportPresentationDependencies", "Invalid report conditional-display metadata." & vbCrLf & "Field: " & CStr(specData(REPORT_FIELD_NAME))
            If conditionalMode = REPORT_CONDITIONAL_NONEMPTY And StrComp(CStr(rowKey), REPORT_ROW_PRIMARY, vbBinaryCompare) = 0 And ReportUsesRecordDetail(definitionData) = False Then Err.Raise vbObjectError + 9904, "ValidateReportPresentationDependencies", "NONEMPTY conditional display is supported on primary fields only in RECORD_DETAIL layout." & vbCrLf & "Field: " & CStr(specData(REPORT_FIELD_NAME))
        End If
    Next specData
    Set availableIndex = ReportAvailableFieldIndex(definitionData)
    sortField = CStr(definitionData(REPORT_DEF_SORT_FIELD))
    If Len(sortField) > 0 Then If availableIndex.Exists(sortField) = False Then Err.Raise vbObjectError + 9905, "ValidateReportPresentationDependencies", "The canonical report sort field is not exposed by the authoritative RecordSource." & vbCrLf & "Report: " & CStr(definitionData(REPORT_DEF_LOGICAL_NAME)) & vbCrLf & "Field: " & sortField
    Call ReportResolvedGroupField(definitionData)
    If StrComp(CStr(definitionData(REPORT_DEF_LOGICAL_NAME)), CANON_REPORT_LOGICAL_ID_ACTIVE_MACROS, vbBinaryCompare) = 0 Then
        If availableIndex.Exists("Active") = False Then Err.Raise vbObjectError + 9886, "ValidateReportPresentationDependencies", "Active Macros requires Active for its strict active-only report filter."
    End If
End Sub

Private Function ReportObjectExists(ByVal reportName As String) As Boolean
    ' Wrapper purpose: Private Reports-domain adapter that binds the generic Access-object existence check to HostDb and acReport.
    ' Why retained: Retained to prevent REPORTS callers from selecting the wrong database or object type and to keep Report-domain existence checks consistent.
    ReportObjectExists = SharedDao_AccessObjectExists(GetHostDb(), acReport, reportName)
End Function

Private Function ReportIsOpen(ByVal reportName As String) As Boolean
    ReportIsOpen = ((SysCmd(acSysCmdGetObjectState, acReport, reportName) And acObjStateOpen) <> 0)
End Function

Private Function ReportBackupPath(ByVal reportName As String) As String
    ReportBackupPath = Environ$("TEMP") & "\\ADC_E_Report_" & StableTextHashHex(reportName & "|" & Format$(Now, "yyyymmddhhnnss") & "|" & CStr(Timer)) & ".txt"
End Function

Private Sub VerifyReportBackupReadable(ByVal backupPath As String, ByVal reportName As String)
    Dim fileNumber As Integer
    If Len(Trim$(backupPath)) = 0 Or Len(Dir$(backupPath)) = 0 Then Err.Raise vbObjectError + 9872, "VerifyReportBackupReadable", "The existing managed report could not be backed up before replacement." & vbCrLf & "Report: " & reportName & vbCrLf & "Backup: " & backupPath
    fileNumber = FreeFile
    Open backupPath For Binary Access Read As #fileNumber
    If LOF(fileNumber) <= 0 Then Close #fileNumber: Err.Raise vbObjectError + 9873, "VerifyReportBackupReadable", "The managed-report rollback backup is empty." & vbCrLf & "Report: " & reportName & vbCrLf & "Backup: " & backupPath
    Close #fileNumber
End Sub

Private Function ManagedReportLayoutCurrent(ByVal rpt As Access.Report, ByVal definitionData As Variant, ByVal selectedFields As Collection, ByRef detailText As String) As Boolean
    On Error GoTo NotCurrent
    Dim titleCtl As Access.control
    Dim screenTitleCtl As Access.control
    Dim labelCtl As Access.control
    Dim screenLabelCtl As Access.control
    Dim fieldCtl As Access.control
    Dim primarySpecs As Collection
    Dim detailRows As Collection
    Dim visualRows As Collection
    Dim rowSpecs As Collection
    Dim specData As Variant
    Dim rowKey As Variant
    Dim widths As Variant
    Dim leftPos As Long
    Dim topPos As Long
    Dim valueTop As Long
    Dim rowValueHeight As Long
    Dim labelWidth As Long
    Dim index As Long
    Dim detailIndex As Long
    Dim groupField As String
    Dim expectedControlCount As Long
    Dim expectedSource As String
    Dim expectedFormat As String
    Dim expectedLabelFormat As String
    Dim conditionalMode As String
    Dim expectedSort As String
    Dim expectedHeaderHeight As Long
    Dim recordDetail As Boolean
    Set primarySpecs = ReportSpecsForRow(definitionData, REPORT_ROW_PRIMARY)
    Set detailRows = ReportDetailRowKeys(definitionData)
    groupField = ReportResolvedGroupField(definitionData)
    recordDetail = ReportUsesRecordDetail(definitionData)
    If recordDetail Then
        expectedControlCount = ReportExpectedRecordControlCount(definitionData)
    Else
        expectedControlCount = 2 + primarySpecs.count + (selectedFields.count * 2)
    End If
    If Len(groupField) > 0 Then expectedControlCount = expectedControlCount + 1
    If rpt.Controls.count <> expectedControlCount Then detailText = "Control-count mismatch: " & rpt.Name: Exit Function
    If rpt.Width <> REPORT_PAGE_WIDTH Then detailText = "Report-width mismatch: " & rpt.Name: Exit Function
    If rpt.AllowReportView = False Then detailText = "Report View is not allowed: " & rpt.Name: Exit Function
    If rpt.PageHeader <> REPORT_PAGE_HEADER_NOT_WITH_REPORT_HEADER Then detailText = "Page-header/report-header print interaction mismatch: " & rpt.Name: Exit Function
    expectedHeaderHeight = ReportHeaderSectionHeight(definitionData)
    If rpt.Section(acHeader).Height <> expectedHeaderHeight Then detailText = "Report-header screen-view height mismatch: " & rpt.Name: Exit Function
    If rpt.Section(acFooter).Height <> 0 Then detailText = "Report-footer height mismatch: " & rpt.Name: Exit Function
    If rpt.Section(acPageHeader).Height <> expectedHeaderHeight Then detailText = "Page-header height mismatch: " & rpt.Name: Exit Function
    If rpt.Section(acDetail).Height <> ReportDetailSectionHeight(definitionData) Then detailText = "Detail height mismatch: " & rpt.Name: Exit Function
    If rpt.Section(acDetail).canGrow = False Or rpt.Section(acDetail).CanShrink = False Then detailText = "Detail section growth/shrink contract mismatch: " & rpt.Name: Exit Function
    If rpt.Printer.Orientation <> acPRORLandscape Then detailText = "Orientation mismatch: " & rpt.Name: Exit Function
    If rpt.Printer.LeftMargin <> REPORT_LEFT_MARGIN Or rpt.Printer.RightMargin <> REPORT_RIGHT_MARGIN Or rpt.Printer.TopMargin <> REPORT_TOP_MARGIN Or rpt.Printer.BottomMargin <> REPORT_BOTTOM_MARGIN Then detailText = "Printer-margin mismatch: " & rpt.Name: Exit Function
    Set titleCtl = rpt.Controls("adc_report_title")
    Set screenTitleCtl = rpt.Controls("adc_screen_report_title")
    If StrComp(CStr(titleCtl.Caption), CStr(definitionData(REPORT_DEF_TITLE)), vbBinaryCompare) <> 0 Then detailText = "Report-title control mismatch: " & rpt.Name: Exit Function
    If StrComp(CStr(screenTitleCtl.Caption), CStr(definitionData(REPORT_DEF_TITLE)), vbBinaryCompare) <> 0 Then detailText = "Report View title control mismatch: " & rpt.Name: Exit Function
    If screenTitleCtl.Left <> REPORT_LEFT_MARGIN Or screenTitleCtl.Top <> REPORT_TITLE_TOP Or screenTitleCtl.Width <> REPORT_PAGE_WIDTH - REPORT_LEFT_MARGIN - REPORT_RIGHT_MARGIN Or screenTitleCtl.Height <> REPORT_TITLE_HEIGHT Then detailText = "Report View title layout mismatch: " & rpt.Name: Exit Function
    If titleCtl.Left <> REPORT_LEFT_MARGIN Or titleCtl.Top <> REPORT_TITLE_TOP Or titleCtl.Width <> REPORT_PAGE_WIDTH - REPORT_LEFT_MARGIN - REPORT_RIGHT_MARGIN Or titleCtl.Height <> REPORT_TITLE_HEIGHT Then detailText = "Printed title layout mismatch: " & rpt.Name: Exit Function
    If titleCtl.fontSize <> REPORT_FONT_SIZE_TITLE Or screenTitleCtl.fontSize <> REPORT_FONT_SIZE_TITLE Then detailText = "Report-title font mismatch: " & rpt.Name: Exit Function
    If recordDetail Then
        Set visualRows = ReportVisualRowKeys(definitionData)
        detailIndex = 0
        topPos = REPORT_RECORD_DETAIL_TOP
        For Each rowKey In visualRows
            Set rowSpecs = ReportSpecsForRow(definitionData, CStr(rowKey))
            widths = ReportScaledSpecWidths(rowSpecs, REPORT_RECORD_COLUMN_GAP)
            leftPos = REPORT_LEFT_MARGIN
            index = 0
            valueTop = topPos + REPORT_RECORD_LABEL_HEIGHT + REPORT_RECORD_LABEL_VALUE_GAP
            rowValueHeight = ReportRecordValueHeight(definitionData, CStr(rowKey))
            For Each specData In rowSpecs
                detailIndex = detailIndex + 1
                Set labelCtl = rpt.Controls("adc_record_label_" & Format$(detailIndex, "00"))
                Set fieldCtl = rpt.Controls("adc_record_field_" & Format$(detailIndex, "00"))
                expectedSource = "[" & Replace(CStr(specData(REPORT_FIELD_NAME)), "]", "]]") & "]"
                conditionalMode = CStr(specData(REPORT_FIELD_CONDITIONAL_MODE))
                If StrComp(conditionalMode, REPORT_CONDITIONAL_NONEMPTY, vbBinaryCompare) = 0 Then
                    If StrComp(typeName(labelCtl), "TextBox", vbTextCompare) <> 0 Then detailText = "Conditional record label type mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
                    If StrComp(Trim$(CStr(labelCtl.controlSource)), expectedSource, vbBinaryCompare) <> 0 Then detailText = "Conditional record label ControlSource mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
                    expectedLabelFormat = ReportConditionalLabelFormat(CStr(specData(REPORT_FIELD_CAPTION)))
                    If StrComp(Nz(labelCtl.Format, vbNullString), expectedLabelFormat, vbBinaryCompare) <> 0 Then detailText = "Conditional record label format mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
                    If labelCtl.canGrow Or labelCtl.CanShrink = False Then detailText = "Conditional record label shrink contract mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
                Else
                    If StrComp(typeName(labelCtl), "Label", vbTextCompare) <> 0 Then detailText = "Record label type mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
                    If StrComp(CStr(labelCtl.Caption), CStr(specData(REPORT_FIELD_CAPTION)), vbBinaryCompare) <> 0 Then detailText = "Record caption mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
                End If
                If labelCtl.Left <> leftPos Or labelCtl.Top <> topPos Or labelCtl.Width <> CLng(widths(index)) Or labelCtl.Height <> REPORT_RECORD_LABEL_HEIGHT Then detailText = "Record label layout mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
                If labelCtl.fontSize <> REPORT_FONT_SIZE_HEADER Or labelCtl.fontBold = False Then detailText = "Record label emphasis mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
                If StrComp(Trim$(CStr(fieldCtl.controlSource)), expectedSource, vbBinaryCompare) <> 0 Then detailText = "Record ControlSource mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
                If fieldCtl.Left <> leftPos Or fieldCtl.Top <> valueTop Or fieldCtl.Width <> CLng(widths(index)) Or fieldCtl.Height <> rowValueHeight Then detailText = "Record field layout mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
                If fieldCtl.canGrow <> ReportSpecBoolean(specData(REPORT_FIELD_CAN_GROW)) Or fieldCtl.CanShrink = False Then detailText = "Record growth contract mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
                If fieldCtl.TextAlign <> ReportTextAlignment(CStr(specData(REPORT_FIELD_ALIGNMENT))) Then detailText = "Record alignment mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
                expectedFormat = ReportFormatText(CStr(specData(REPORT_FIELD_FORMAT)))
                If StrComp(Nz(fieldCtl.Format, vbNullString), expectedFormat, vbBinaryCompare) <> 0 Then detailText = "Record format mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
                If StrComp(CStr(rowKey), REPORT_ROW_PRIMARY, vbBinaryCompare) = 0 Then
                    If fieldCtl.fontSize <> REPORT_FONT_SIZE_IDENTITY Or fieldCtl.fontBold = False Then detailText = "Primary record emphasis mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
                Else
                    If fieldCtl.fontSize <> REPORT_FONT_SIZE_DETAIL Or fieldCtl.fontBold Then detailText = "Secondary record emphasis mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
                End If
                If fieldCtl.Left + fieldCtl.Width > REPORT_PAGE_WIDTH - REPORT_RIGHT_MARGIN Then detailText = "Record control exceeds printable width: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
                leftPos = leftPos + CLng(widths(index)) + REPORT_RECORD_COLUMN_GAP
                index = index + 1
            Next specData
            topPos = valueTop + rowValueHeight + REPORT_RECORD_ROW_GAP
        Next rowKey
    Else
        widths = ReportScaledSpecWidths(primarySpecs)
        leftPos = REPORT_LEFT_MARGIN
        index = 0
        For Each specData In primarySpecs
            Set labelCtl = rpt.Controls("adc_primary_label_" & Format$(index + 1, "00"))
            Set screenLabelCtl = rpt.Controls("adc_screen_primary_label_" & Format$(index + 1, "00"))
            Set fieldCtl = rpt.Controls("adc_primary_field_" & Format$(index + 1, "00"))
            expectedSource = "[" & Replace(CStr(specData(REPORT_FIELD_NAME)), "]", "]]") & "]"
            If StrComp(CStr(labelCtl.Caption), CStr(specData(REPORT_FIELD_CAPTION)), vbBinaryCompare) <> 0 Then detailText = "Primary caption mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
            If StrComp(CStr(screenLabelCtl.Caption), CStr(specData(REPORT_FIELD_CAPTION)), vbBinaryCompare) <> 0 Then detailText = "Report View primary caption mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
            If StrComp(Trim$(CStr(fieldCtl.controlSource)), expectedSource, vbBinaryCompare) <> 0 Then detailText = "Primary ControlSource mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
            If labelCtl.Left <> leftPos Or screenLabelCtl.Left <> leftPos Or fieldCtl.Left <> leftPos Then detailText = "Primary left mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
            If labelCtl.Width <> CLng(widths(index)) Or screenLabelCtl.Width <> CLng(widths(index)) Or fieldCtl.Width <> CLng(widths(index)) Then detailText = "Primary width mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
            If screenLabelCtl.Top <> REPORT_COLUMN_HEADER_TOP Or screenLabelCtl.Height <> REPORT_COLUMN_HEADER_HEIGHT Then detailText = "Report View primary label layout mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
            If fieldCtl.Top <> REPORT_PRIMARY_CONTROL_TOP Or fieldCtl.Height <> REPORT_CONTROL_HEIGHT Then detailText = "Primary field layout mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
            If fieldCtl.canGrow Or fieldCtl.CanShrink Then detailText = "Primary growth contract mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
            If fieldCtl.TextAlign <> ReportTextAlignment(CStr(specData(REPORT_FIELD_ALIGNMENT))) Then detailText = "Primary alignment mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
            expectedFormat = ReportFormatText(CStr(specData(REPORT_FIELD_FORMAT)))
            If StrComp(Nz(fieldCtl.Format, vbNullString), expectedFormat, vbBinaryCompare) <> 0 Then detailText = "Primary format mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
            If fieldCtl.Left + fieldCtl.Width > REPORT_PAGE_WIDTH - REPORT_RIGHT_MARGIN Then detailText = "Primary control exceeds printable width: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
            leftPos = leftPos + CLng(widths(index))
            index = index + 1
        Next specData
        detailIndex = 0
        topPos = REPORT_PRIMARY_CONTROL_TOP + REPORT_CONTROL_HEIGHT + REPORT_DETAIL_ROW_GAP
        For Each rowKey In detailRows
            Set rowSpecs = ReportSpecsForRow(definitionData, CStr(rowKey))
            widths = ReportScaledSpecWidths(rowSpecs)
            leftPos = REPORT_LEFT_MARGIN
            index = 0
            For Each specData In rowSpecs
                detailIndex = detailIndex + 1
                Set labelCtl = rpt.Controls("adc_detail_label_" & Format$(detailIndex, "00"))
                Set fieldCtl = rpt.Controls("adc_detail_field_" & Format$(detailIndex, "00"))
                labelWidth = ReportDetailLabelWidth(CLng(widths(index)))
                expectedSource = "[" & Replace(CStr(specData(REPORT_FIELD_NAME)), "]", "]]") & "]"
                conditionalMode = CStr(specData(REPORT_FIELD_CONDITIONAL_MODE))
                If StrComp(conditionalMode, REPORT_CONDITIONAL_NONEMPTY, vbBinaryCompare) = 0 Then
                    If StrComp(typeName(labelCtl), "TextBox", vbTextCompare) <> 0 Then detailText = "Conditional detail label type mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
                    If StrComp(Trim$(CStr(labelCtl.controlSource)), expectedSource, vbBinaryCompare) <> 0 Then detailText = "Conditional detail label ControlSource mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
                    expectedLabelFormat = ReportConditionalLabelFormat(CStr(specData(REPORT_FIELD_CAPTION)))
                    If StrComp(Nz(labelCtl.Format, vbNullString), expectedLabelFormat, vbBinaryCompare) <> 0 Then detailText = "Conditional detail label format mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
                    If labelCtl.canGrow Or labelCtl.CanShrink = False Then detailText = "Conditional detail label shrink contract mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
                Else
                    If StrComp(typeName(labelCtl), "Label", vbTextCompare) <> 0 Then detailText = "Detail label type mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
                    If StrComp(CStr(labelCtl.Caption), CStr(specData(REPORT_FIELD_CAPTION)), vbBinaryCompare) <> 0 Then detailText = "Detail caption mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
                End If
                If StrComp(Trim$(CStr(fieldCtl.controlSource)), expectedSource, vbBinaryCompare) <> 0 Then detailText = "Detail ControlSource mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
                If labelCtl.Left <> leftPos Or fieldCtl.Left <> leftPos + labelWidth + REPORT_DETAIL_LABEL_GAP Then detailText = "Detail left mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
                If labelCtl.Width <> labelWidth Or fieldCtl.Width <> CLng(widths(index)) - labelWidth - REPORT_DETAIL_LABEL_GAP Then detailText = "Detail width mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
                If labelCtl.Top <> topPos Or fieldCtl.Top <> topPos Or labelCtl.Height <> REPORT_CONTROL_HEIGHT Or fieldCtl.Height <> REPORT_CONTROL_HEIGHT Then detailText = "Detail layout mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
                If fieldCtl.canGrow <> ReportSpecBoolean(specData(REPORT_FIELD_CAN_GROW)) Or fieldCtl.CanShrink = False Then detailText = "Detail growth contract mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
                If fieldCtl.TextAlign <> ReportTextAlignment(CStr(specData(REPORT_FIELD_ALIGNMENT))) Then detailText = "Detail alignment mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
                expectedFormat = ReportFormatText(CStr(specData(REPORT_FIELD_FORMAT)))
                If StrComp(Nz(fieldCtl.Format, vbNullString), expectedFormat, vbBinaryCompare) <> 0 Then detailText = "Detail format mismatch: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
                If fieldCtl.Left + fieldCtl.Width > REPORT_PAGE_WIDTH - REPORT_RIGHT_MARGIN Then detailText = "Detail control exceeds printable width: " & rpt.Name & " / " & CStr(specData(REPORT_FIELD_NAME)): Exit Function
                leftPos = leftPos + CLng(widths(index))
                index = index + 1
            Next specData
            topPos = topPos + REPORT_CONTROL_HEIGHT + REPORT_DETAIL_ROW_GAP
        Next rowKey
    End If
    If Len(groupField) > 0 Then
        If StrComp(Trim$(CStr(rpt.GroupLevel(0).controlSource)), groupField, vbTextCompare) <> 0 Then detailText = "Group-field mismatch: " & rpt.Name: Exit Function
        If rpt.Section(acGroupLevel1Header).Height <> REPORT_GROUP_HEADER_HEIGHT Then detailText = "Group-header height mismatch: " & rpt.Name: Exit Function
        Set fieldCtl = rpt.Controls("adc_group_value")
        If StrComp(Trim$(CStr(fieldCtl.controlSource)), "[" & Replace(groupField, "]", "]]") & "]", vbBinaryCompare) <> 0 Then detailText = "Group-header control mismatch: " & rpt.Name: Exit Function
        If fieldCtl.fontSize <> REPORT_FONT_SIZE_GROUP Then detailText = "Group-header font mismatch: " & rpt.Name: Exit Function
    Else
        expectedSort = CStr(definitionData(REPORT_DEF_SORT_FIELD))
        If Len(expectedSort) > 0 Then
            If StrComp(Trim$(Nz(rpt.OrderBy, vbNullString)), "[" & Replace(expectedSort, "]", "]]") & "]", vbBinaryCompare) <> 0 Or rpt.OrderByOn = False Then detailText = "Sort-order mismatch: " & rpt.Name: Exit Function
        End If
    End If
    ManagedReportLayoutCurrent = True
    Exit Function
NotCurrent:
    detailText = "Could not verify managed-report layout [" & rpt.Name & "]. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
End Function

Private Function ManagedReportCurrent(ByVal definitionData As Variant, ByVal selectedFields As Collection, ByRef detailText As String) As Boolean
    On Error GoTo NotCurrent
    Dim reportName As String
    Dim expectedSignature As String
    Dim rpt As Access.Report
    Dim openedHere As Boolean
    reportName = ManagedReportPhysicalName(CStr(definitionData(REPORT_DEF_LOGICAL_NAME)))
    If ReportObjectExists(reportName) = False Then detailText = "Missing report: " & reportName: Exit Function
    expectedSignature = ReportExpectedSignature(definitionData, selectedFields)
    If ReportIsOpen(reportName) = False Then DoCmd.OpenReport reportName, acViewDesign, , , acHidden: openedHere = True
    Set rpt = reports(reportName)
    If StrComp(Trim$(Nz(rpt.recordSource, vbNullString)), ReportRecordSourceText(definitionData), vbBinaryCompare) <> 0 Then detailText = "RecordSource mismatch: " & reportName: GoTo Cleanup
    If StrComp(Trim$(Nz(rpt.Caption, vbNullString)), CStr(definitionData(REPORT_DEF_TITLE)), vbBinaryCompare) <> 0 Then detailText = "Caption mismatch: " & reportName: GoTo Cleanup
    If StrComp(Trim$(Nz(rpt.Tag, vbNullString)), expectedSignature, vbBinaryCompare) <> 0 Then detailText = "Layout signature mismatch: " & reportName: GoTo Cleanup
    If StrComp(Trim$(Nz(rpt.Filter, vbNullString)), ReportFilterExpression(definitionData), vbBinaryCompare) <> 0 Then detailText = "Filter mismatch: " & reportName: GoTo Cleanup
    If rpt.FilterOnLoad <> (Len(ReportFilterExpression(definitionData)) > 0) Then detailText = "FilterOnLoad mismatch: " & reportName: GoTo Cleanup
    If ManagedReportLayoutCurrent(rpt, definitionData, selectedFields, detailText) = False Then GoTo Cleanup
    ManagedReportCurrent = True
Cleanup:
    On Error Resume Next
    Set rpt = Nothing
    If openedHere Then DoCmd.Close acReport, reportName, acSaveNo
    On Error GoTo 0
    Exit Function
NotCurrent:
    detailText = "Could not verify report [" & reportName & "]. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    Resume Cleanup
End Function

Private Sub BuildManagedReport(ByVal definitionData As Variant, ByVal selectedFields As Collection)
    On Error GoTo ErrHandler
    Dim reportName As String
    Dim sourceName As String
    Dim reportTitle As String
    Dim existingTarget As Boolean
    Dim rollbackTracking As Boolean
    Dim backupPath As String
    Dim tempName As String
    Dim rpt As Access.Report
    Dim ctl As Access.control
    Dim labelCtl As Access.control
    Dim screenLabelCtl As Access.control
    Dim primarySpecs As Collection
    Dim detailRows As Collection
    Dim visualRows As Collection
    Dim rowSpecs As Collection
    Dim specData As Variant
    Dim rowKey As Variant
    Dim widths As Variant
    Dim index As Long
    Dim detailIndex As Long
    Dim leftPos As Long
    Dim topPos As Long
    Dim valueTop As Long
    Dim rowValueHeight As Long
    Dim labelWidth As Long
    Dim groupField As String
    Dim conditionalMode As String
    Dim expectedSource As String
    Dim reportHeaderExists As Boolean
    Dim cleanupWarnings As String
    Dim recordDetail As Boolean
    reportName = ManagedReportPhysicalName(CStr(definitionData(REPORT_DEF_LOGICAL_NAME)))
    sourceName = ReportRecordSourceText(definitionData)
    reportTitle = CStr(definitionData(REPORT_DEF_TITLE))
    existingTarget = ReportObjectExists(reportName)
    rollbackTracking = SetupRollbackTrackingActive()
    If ReportIsOpen(reportName) Then DoCmd.Close acReport, reportName, acSaveNo
    If existingTarget Then
        backupPath = ReportBackupPath(reportName)
        Application.SaveAsText acReport, reportName, backupPath
        VerifyReportBackupReadable backupPath, reportName
    End If
    If rollbackTracking Then TrackSetupAccessObjectRollback SETUP_ROLLBACK_RESTORE_REPORT_ACTION, reportName, existingTarget, backupPath
    If existingTarget Then DoCmd.DeleteObject acReport, reportName
    Set rpt = CreateReport
    tempName = rpt.Name
    On Error Resume Next
    reportHeaderExists = (rpt.Section(acHeader).Height >= 0)
    Err.Clear
    On Error GoTo ErrHandler
    If reportHeaderExists = False Then DoCmd.RunCommand acCmdReportHdrFtr
    rpt.recordSource = sourceName
    rpt.Caption = reportTitle
    rpt.AllowReportView = True
    rpt.Tag = ReportExpectedSignature(definitionData, selectedFields)
    rpt.Filter = ReportFilterExpression(definitionData)
    rpt.FilterOnLoad = (Len(rpt.Filter) > 0)
    Set primarySpecs = ReportSpecsForRow(definitionData, REPORT_ROW_PRIMARY)
    Set detailRows = ReportDetailRowKeys(definitionData)
    recordDetail = ReportUsesRecordDetail(definitionData)
    rpt.Width = REPORT_PAGE_WIDTH
    rpt.PageHeader = REPORT_PAGE_HEADER_NOT_WITH_REPORT_HEADER
    rpt.Section(acHeader).Height = ReportHeaderSectionHeight(definitionData)
    rpt.Section(acHeader).BackColor = RGB(255, 255, 255)
    rpt.Section(acFooter).Height = 0
    rpt.Section(acPageHeader).Height = ReportHeaderSectionHeight(definitionData)
    rpt.Section(acPageHeader).BackColor = RGB(255, 255, 255)
    rpt.Section(acDetail).Height = ReportDetailSectionHeight(definitionData)
    rpt.Section(acDetail).canGrow = True
    rpt.Section(acDetail).CanShrink = True
    rpt.Section(acDetail).BackColor = RGB(255, 255, 255)
    groupField = ReportResolvedGroupField(definitionData)
    If Len(groupField) > 0 Then
        Call Application.CreateGroupLevel(tempName, groupField, True, False)
        rpt.Section(acGroupLevel1Header).Height = REPORT_GROUP_HEADER_HEIGHT
        rpt.Section(acGroupLevel1Header).BackColor = RGB(73, 79, 86)
        Set ctl = CreateReportControl(tempName, acTextBox, acGroupLevel1Header, , , REPORT_LEFT_MARGIN, REPORT_GROUP_CONTROL_TOP, REPORT_PAGE_WIDTH - REPORT_LEFT_MARGIN - REPORT_RIGHT_MARGIN, REPORT_CONTROL_HEIGHT)
        ctl.Name = "adc_group_value"
        ctl.controlSource = "[" & Replace(groupField, "]", "]]") & "]"
        ctl.FontName = "Segoe UI"
        ctl.fontBold = True
        ctl.fontSize = REPORT_FONT_SIZE_GROUP
        ctl.ForeColor = RGB(255, 255, 255)
        ctl.BackStyle = 0
        ctl.BorderStyle = 0
        ctl.canGrow = False
        ctl.CanShrink = False
    End If
    rpt.Printer.Orientation = acPRORLandscape
    rpt.Printer.LeftMargin = REPORT_LEFT_MARGIN
    rpt.Printer.RightMargin = REPORT_RIGHT_MARGIN
    rpt.Printer.TopMargin = REPORT_TOP_MARGIN
    rpt.Printer.BottomMargin = REPORT_BOTTOM_MARGIN
    Set labelCtl = CreateReportControl(tempName, acLabel, acPageHeader, , , REPORT_LEFT_MARGIN, REPORT_TITLE_TOP, REPORT_PAGE_WIDTH - REPORT_LEFT_MARGIN - REPORT_RIGHT_MARGIN, REPORT_TITLE_HEIGHT)
    labelCtl.Name = "adc_report_title"
    labelCtl.Caption = reportTitle
    labelCtl.FontName = "Segoe UI"
    labelCtl.fontBold = True
    labelCtl.fontSize = REPORT_FONT_SIZE_TITLE
    labelCtl.ForeColor = RGB(34, 39, 44)
    labelCtl.BackStyle = 0
    labelCtl.BorderStyle = 0
    Set screenLabelCtl = CreateReportControl(tempName, acLabel, acHeader, , , REPORT_LEFT_MARGIN, REPORT_TITLE_TOP, REPORT_PAGE_WIDTH - REPORT_LEFT_MARGIN - REPORT_RIGHT_MARGIN, REPORT_TITLE_HEIGHT)
    screenLabelCtl.Name = "adc_screen_report_title"
    screenLabelCtl.Caption = reportTitle
    screenLabelCtl.FontName = "Segoe UI"
    screenLabelCtl.fontBold = True
    screenLabelCtl.fontSize = REPORT_FONT_SIZE_TITLE
    screenLabelCtl.ForeColor = RGB(34, 39, 44)
    screenLabelCtl.BackStyle = 0
    screenLabelCtl.BorderStyle = 0
    If recordDetail Then
        Set visualRows = ReportVisualRowKeys(definitionData)
        detailIndex = 0
        topPos = REPORT_RECORD_DETAIL_TOP
        For Each rowKey In visualRows
            Set rowSpecs = ReportSpecsForRow(definitionData, CStr(rowKey))
            widths = ReportScaledSpecWidths(rowSpecs, REPORT_RECORD_COLUMN_GAP)
            leftPos = REPORT_LEFT_MARGIN
            index = 0
            valueTop = topPos + REPORT_RECORD_LABEL_HEIGHT + REPORT_RECORD_LABEL_VALUE_GAP
            rowValueHeight = ReportRecordValueHeight(definitionData, CStr(rowKey))
            For Each specData In rowSpecs
                detailIndex = detailIndex + 1
                expectedSource = "[" & Replace(CStr(specData(REPORT_FIELD_NAME)), "]", "]]") & "]"
                conditionalMode = CStr(specData(REPORT_FIELD_CONDITIONAL_MODE))
                If StrComp(conditionalMode, REPORT_CONDITIONAL_NONEMPTY, vbBinaryCompare) = 0 Then
                    Set labelCtl = CreateReportControl(tempName, acTextBox, acDetail, , , leftPos, topPos, CLng(widths(index)), REPORT_RECORD_LABEL_HEIGHT)
                    labelCtl.Name = "adc_record_label_" & Format$(detailIndex, "00")
                    labelCtl.controlSource = expectedSource
                    ApplyReportLabelStyle labelCtl, False
                    labelCtl.TextAlign = REPORT_TEXT_ALIGN_LEFT
                    labelCtl.Format = ReportConditionalLabelFormat(CStr(specData(REPORT_FIELD_CAPTION)))
                    labelCtl.canGrow = False
                    labelCtl.CanShrink = True
                Else
                    Set labelCtl = CreateReportControl(tempName, acLabel, acDetail, , , leftPos, topPos, CLng(widths(index)), REPORT_RECORD_LABEL_HEIGHT)
                    labelCtl.Name = "adc_record_label_" & Format$(detailIndex, "00")
                    labelCtl.Caption = CStr(specData(REPORT_FIELD_CAPTION))
                    ApplyReportLabelStyle labelCtl, False
                End If
                Set ctl = CreateReportControl(tempName, acTextBox, acDetail, , , leftPos, valueTop, CLng(widths(index)), rowValueHeight)
                ctl.Name = "adc_record_field_" & Format$(detailIndex, "00")
                ctl.controlSource = expectedSource
                ApplyReportFieldStyle ctl, specData, False
                ctl.fontBold = (StrComp(CStr(rowKey), REPORT_ROW_PRIMARY, vbBinaryCompare) = 0)
                If ctl.fontBold Then ctl.fontSize = REPORT_FONT_SIZE_IDENTITY
                leftPos = leftPos + CLng(widths(index)) + REPORT_RECORD_COLUMN_GAP
                index = index + 1
            Next specData
            topPos = valueTop + rowValueHeight + REPORT_RECORD_ROW_GAP
        Next rowKey
    Else
        widths = ReportScaledSpecWidths(primarySpecs)
        leftPos = REPORT_LEFT_MARGIN
        index = 0
        For Each specData In primarySpecs
            Set labelCtl = CreateReportControl(tempName, acLabel, acPageHeader, , , leftPos, REPORT_COLUMN_HEADER_TOP, CLng(widths(index)), REPORT_COLUMN_HEADER_HEIGHT)
            labelCtl.Name = "adc_primary_label_" & Format$(index + 1, "00")
            labelCtl.Caption = CStr(specData(REPORT_FIELD_CAPTION))
            ApplyReportLabelStyle labelCtl, True
            Set screenLabelCtl = CreateReportControl(tempName, acLabel, acHeader, , , leftPos, REPORT_COLUMN_HEADER_TOP, CLng(widths(index)), REPORT_COLUMN_HEADER_HEIGHT)
            screenLabelCtl.Name = "adc_screen_primary_label_" & Format$(index + 1, "00")
            screenLabelCtl.Caption = CStr(specData(REPORT_FIELD_CAPTION))
            ApplyReportLabelStyle screenLabelCtl, True
            Set ctl = CreateReportControl(tempName, acTextBox, acDetail, , , leftPos, REPORT_PRIMARY_CONTROL_TOP, CLng(widths(index)), REPORT_CONTROL_HEIGHT)
            ctl.Name = "adc_primary_field_" & Format$(index + 1, "00")
            ctl.controlSource = "[" & Replace(CStr(specData(REPORT_FIELD_NAME)), "]", "]]") & "]"
            ApplyReportFieldStyle ctl, specData, True
            leftPos = leftPos + CLng(widths(index))
            index = index + 1
        Next specData
        detailIndex = 0
        topPos = REPORT_PRIMARY_CONTROL_TOP + REPORT_CONTROL_HEIGHT + REPORT_DETAIL_ROW_GAP
        For Each rowKey In detailRows
            Set rowSpecs = ReportSpecsForRow(definitionData, CStr(rowKey))
            widths = ReportScaledSpecWidths(rowSpecs)
            leftPos = REPORT_LEFT_MARGIN
            index = 0
            For Each specData In rowSpecs
                detailIndex = detailIndex + 1
                labelWidth = ReportDetailLabelWidth(CLng(widths(index)))
                expectedSource = "[" & Replace(CStr(specData(REPORT_FIELD_NAME)), "]", "]]") & "]"
                conditionalMode = CStr(specData(REPORT_FIELD_CONDITIONAL_MODE))
                If StrComp(conditionalMode, REPORT_CONDITIONAL_NONEMPTY, vbBinaryCompare) = 0 Then
                    Set labelCtl = CreateReportControl(tempName, acTextBox, acDetail, , , leftPos, topPos, labelWidth, REPORT_CONTROL_HEIGHT)
                    labelCtl.Name = "adc_detail_label_" & Format$(detailIndex, "00")
                    labelCtl.controlSource = expectedSource
                    ApplyReportLabelStyle labelCtl, False
                    labelCtl.TextAlign = REPORT_TEXT_ALIGN_LEFT
                    labelCtl.Format = ReportConditionalLabelFormat(CStr(specData(REPORT_FIELD_CAPTION)))
                    labelCtl.canGrow = False
                    labelCtl.CanShrink = True
                Else
                    Set labelCtl = CreateReportControl(tempName, acLabel, acDetail, , , leftPos, topPos, labelWidth, REPORT_CONTROL_HEIGHT)
                    labelCtl.Name = "adc_detail_label_" & Format$(detailIndex, "00")
                    labelCtl.Caption = CStr(specData(REPORT_FIELD_CAPTION))
                    ApplyReportLabelStyle labelCtl, False
                End If
                Set ctl = CreateReportControl(tempName, acTextBox, acDetail, , , leftPos + labelWidth + REPORT_DETAIL_LABEL_GAP, topPos, CLng(widths(index)) - labelWidth - REPORT_DETAIL_LABEL_GAP, REPORT_CONTROL_HEIGHT)
                ctl.Name = "adc_detail_field_" & Format$(detailIndex, "00")
                ctl.controlSource = expectedSource
                ApplyReportFieldStyle ctl, specData, False
                leftPos = leftPos + CLng(widths(index))
                index = index + 1
            Next specData
            topPos = topPos + REPORT_CONTROL_HEIGHT + REPORT_DETAIL_ROW_GAP
        Next rowKey
    End If
    If Len(groupField) = 0 And Len(CStr(definitionData(REPORT_DEF_SORT_FIELD))) > 0 Then
        rpt.OrderBy = "[" & Replace(CStr(definitionData(REPORT_DEF_SORT_FIELD)), "]", "]]") & "]"
        rpt.OrderByOn = True
    End If
    DoCmd.Save acReport, tempName
    DoCmd.Close acReport, tempName, acSaveYes
    Set rpt = Nothing
    DoCmd.Rename reportName, acReport, tempName
    tempName = vbNullString
    If rollbackTracking = False And Len(backupPath) > 0 Then DeleteRollbackBackupFile backupPath, cleanupWarnings
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    Set ctl = Nothing
    Set labelCtl = Nothing
    Set screenLabelCtl = Nothing
    Set rpt = Nothing
    If Len(tempName) > 0 Then
        If (SysCmd(acSysCmdGetObjectState, acReport, tempName) And acObjStateOpen) <> 0 Then DoCmd.Close acReport, tempName, acSaveNo
        If SharedDao_AccessObjectExists(GetHostDb(), acReport, tempName) Then DoCmd.DeleteObject acReport, tempName
    End If
    If rollbackTracking = False And existingTarget And Len(backupPath) > 0 And Len(Dir$(backupPath)) > 0 Then
        If ReportObjectExists(reportName) Then DoCmd.DeleteObject acReport, reportName
        Application.LoadFromText acReport, reportName, backupPath
    End If
    On Error GoTo 0
    Err.Raise savedErrNumber, "BuildManagedReport", "Could not create or replace managed report [" & reportName & "]." & vbCrLf & "Source: " & sourceName & vbCrLf & "Original Err.Source: " & savedErrSource & vbCrLf & "Original Err.Description: " & savedErrDescription
End Sub

' ===============================================================================================================================================================================
' SETUP E-PHASE ORCHESTRATION / PROOF PROCEDURES
' ===============================================================================================================================================================================

' Purpose: Checks managed-report prerequisites and existing physical report state before the Setup phase that owns report generation proceeds.
' Public role: Public module entry point. This procedure is listed in the module header as an intentional externally callable ADC entry point and may be invoked directly by Access/VBA callers.
' Parameters:
'   detailText (Output Optional; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller. Default: vbNullString.
' Returns: Boolean: True when the condition represented by BuildAndVerifyManagedReportsForSetup is present under the routine's documented inputs/current state; False when that condition is absent.
Public Function PreflightManagedReportsForSetup(Optional ByRef detailText As String = vbNullString) As Boolean
    On Error GoTo ErrHandler
    Dim definitions As Collection
    Dim definitionData As Variant
    Dim selectedFields As Collection
    Dim seenReports As Object
    Dim sourceName As String
    Dim reportName As String
    Dim payloadText As String
    ResetReportProofIfContextChanged
    ResetReportSourceProjectionCaches
    Set mReportPresentationCacheByReport = Nothing
    mReportPreflightToken = vbNullString
    mReportBuildToken = vbNullString
    mReportReconciledToken = vbNullString
    Set definitions = ManagedReportDefinitions()
    If definitions.count <> REPORT_EXPECTED_COUNT Then Err.Raise vbObjectError + 9874, "PreflightManagedReportsForSetup", "The canonical report definition count is not nine." & vbCrLf & "Actual: " & CStr(definitions.count)
    Set seenReports = CreateObject("Scripting.Dictionary")
    seenReports.CompareMode = vbTextCompare
    For Each definitionData In definitions
        reportName = ManagedReportPhysicalName(CStr(definitionData(REPORT_DEF_LOGICAL_NAME)))
        If Len(Trim$(reportName)) = 0 Or Len(Trim$(CStr(definitionData(REPORT_DEF_TITLE)))) = 0 Then Err.Raise vbObjectError + 9875, "PreflightManagedReportsForSetup", "A report definition has a blank identity/title."
        If Len(reportName) > MAX_ACCESS_OBJECT_NAME_LENGTH Then Err.Raise vbObjectError + 9883, "PreflightManagedReportsForSetup", "A managed report name exceeds the Access object-name limit." & vbCrLf & "Report: " & reportName & vbCrLf & "Length: " & CStr(Len(reportName)) & vbCrLf & "Maximum: " & CStr(MAX_ACCESS_OBJECT_NAME_LENGTH)
        If seenReports.Exists(reportName) Then Err.Raise vbObjectError + 9876, "PreflightManagedReportsForSetup", "Duplicate managed report identity: " & reportName
        seenReports.Add reportName, True
        sourceName = CStr(definitionData(REPORT_DEF_SOURCE_NAME))
        If ReportSourceExists(sourceName, CStr(definitionData(REPORT_DEF_SOURCE_KIND))) = False Then Err.Raise vbObjectError + 9877, "PreflightManagedReportsForSetup", "A required managed-report record source is missing." & vbCrLf & "Report: " & reportName & vbCrLf & "Source: " & sourceName
        Set selectedFields = ReportSelectedFields(definitionData)
        ValidateReportPresentationDependencies definitionData
        payloadText = payloadText & reportName & "|BaseSource=" & sourceName & "|RecordSource=" & StableTextHashHex(ReportRecordSourceText(definitionData)) & "|" & ReportSelectedFieldPayload(selectedFields) & "|Layout=" & CStr(definitionData(REPORT_DEF_LAYOUT_MODE)) & "|Presentation=" & StableTextHashHex(CStr(definitionData(REPORT_DEF_PRESENTATION_SPEC))) & "|Filter=" & ReportFilterExpression(definitionData) & vbLf
        Set selectedFields = Nothing
    Next definitionData
    mReportPreflightToken = ReportContextToken("E1", payloadText)
    mReportBuildToken = vbNullString
    mReportReconciledToken = vbNullString
    detailText = "E1 report preflight passed for " & CStr(definitions.count) & " declarative report definitions and their existing authoritative record sources. No saved queries or domain scans were created/run."
    PreflightManagedReportsForSetup = True
    Exit Function
ErrHandler:
    detailText = "E1 report preflight failed. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
End Function

' Purpose: Resolves report preflight token for E2 from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   detailText (Output/in-out Optional; String): Carries human-readable diagnostic detail for the operation; when passed ByRef, the routine can return or augment the text for the caller. Default:
'   vbNullString.
' Returns: String: Returns the proof/state token produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function ReportPreflightTokenForE2(Optional ByRef detailText As String = vbNullString) As String
    ResetReportProofIfContextChanged
    If Len(mReportPreflightToken) = 0 Then detailText = "No same-run E1 report preflight proof is available. Run E1 first.": Exit Function
    ReportPreflightTokenForE2 = mReportPreflightToken
End Function

' Purpose: Builds and verify managed reports for setup.
' Public role: Public module entry point. This procedure is listed in the module header as an intentional externally callable ADC entry point and may be invoked directly by Access/VBA callers.
' Parameters:
'   detailText (Output Optional; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller. Default: vbNullString.
' Returns: Boolean: True on the function's affirmative/success branch for build and verify managed reports for setup; False on its explicit negative, unavailable, rejected, or handled-failure branch.
'   not-applicable path.
Public Function BuildAndVerifyManagedReportsForSetup(Optional ByRef detailText As String = vbNullString) As Boolean
    RequireSetupDomainMutationAllowed SETUP_DOMAIN_E2_PHYSICAL_REPORTS, "BuildAndVerifyManagedReportsForSetup"
    On Error GoTo ErrHandler
    Dim definitions As Collection
    Dim definitionData As Variant
    Dim selectedFields As Collection
    Dim currentDetail As String
    Dim reportName As String
    Dim builtCount As Long
    Dim matchedCount As Long
    Dim payloadText As String
    Dim descriptionChangeCount As Long
    ResetReportProofIfContextChanged
    If Len(mReportPreflightToken) = 0 Then Err.Raise vbObjectError + 9878, "BuildAndVerifyManagedReportsForSetup", "E2 requires the same-run E1 report preflight proof."
    Set definitions = ManagedReportDefinitions()
    For Each definitionData In definitions
        PumpSetupResponsiveness
        ResetReportSourceProjectionCaches
        reportName = ManagedReportPhysicalName(CStr(definitionData(REPORT_DEF_LOGICAL_NAME)))
        Set selectedFields = ReportSelectedFields(definitionData)
        currentDetail = vbNullString
        If ManagedReportCurrent(definitionData, selectedFields, currentDetail) Then
            matchedCount = matchedCount + 1
        Else
            CorePulseProgress "E2 building managed report [" & reportName & "]."
            BuildManagedReport definitionData, selectedFields
            currentDetail = vbNullString
            If ManagedReportCurrent(definitionData, selectedFields, currentDetail) = False Then Err.Raise vbObjectError + 9879, "BuildAndVerifyManagedReportsForSetup", "A managed report was not current after E2 build." & vbCrLf & currentDetail
            builtCount = builtCount + 1
        End If
        payloadText = payloadText & reportName & "|" & ReportExpectedSignature(definitionData, selectedFields) & vbLf
        Set selectedFields = Nothing
    Next definitionData
    If SetupDomainMutationWritesAllowedForCurrentPart(SETUP_DOMAIN_G1_DESCRIPTIONS) Then ReconcileStandaloneManagedReportDescriptions definitions, descriptionChangeCount
    mReportBuildToken = ReportContextToken("E2", payloadText)
    mReportReconciledToken = vbNullString
    detailText = "E2 physical reports current: " & CStr(definitions.count) & "; rebuilt: " & CStr(builtCount) & "; already current: " & CStr(matchedCount) & "; standalone descriptions changed: " & CStr(descriptionChangeCount) & "."
    BuildAndVerifyManagedReportsForSetup = True
    Exit Function
ErrHandler:
    mReportBuildToken = vbNullString
    mReportReconciledToken = vbNullString
    detailText = "E2 physical report build/currentness failed. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
End Function

Private Sub ReconcileStandaloneManagedReportDescriptions(ByVal definitions As Collection, ByRef descriptionChangeCount As Long)
    'E2 remains physical report construction only during numbered Setup. Direct standalone E2 completes the Description metadata for exactly the managed reports in its authoritative definition set.
    Dim definitionData As Variant
    Dim reportName As String
    Dim descriptionChanged As Variant
    descriptionChangeCount = 0
    For Each definitionData In definitions
        reportName = ManagedReportPhysicalName(CStr(definitionData(REPORT_DEF_LOGICAL_NAME)))
        descriptionChanged = False
        If ApplyAndVerifyManagedObjectDescriptionForStandaloneRepair(reportName, ADCAT_OBJECT_TYPE_REPORT, descriptionChanged) = False Then Err.Raise vbObjectError + 10128, "ReconcileStandaloneManagedReportDescriptions", "Managed-report Description reconciliation returned False." & vbCrLf & "Report: " & reportName
        If CBool(descriptionChanged) Then descriptionChangeCount = descriptionChangeCount + 1
    Next definitionData
End Sub

' Purpose: Resolves report build token for E3 from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   detailText (Output/in-out Optional; String): Carries human-readable diagnostic detail for the operation; when passed ByRef, the routine can return or augment the text for the caller. Default:
'   vbNullString.
' Returns: String: Returns the proof/state token produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function ReportBuildTokenForE3(Optional ByRef detailText As String = vbNullString) As String
    ResetReportProofIfContextChanged
    If Len(mReportBuildToken) = 0 Then detailText = "No same-run E2 physical-report proof is available. Run E2 first.": Exit Function
    ReportBuildTokenForE3 = mReportBuildToken
End Function

Private Function ManagedReportCatalogDefinitionCurrent(ByVal definitionData As Variant, ByRef detailText As String) As Boolean
    On Error GoTo ErrHandler
    Dim catalogDefinition As Variant
    Dim reportName As String
    reportName = ManagedReportPhysicalName(CStr(definitionData(REPORT_DEF_LOGICAL_NAME)))
    If TryGetManagedObjectDefinition(reportName, ADCAT_OBJECT_TYPE_REPORT, catalogDefinition) = False Then detailText = "Missing logical managed-report definition: " & reportName: Exit Function
    If StrComp(CStr(catalogDefinition(1)), ADCAT_CATEGORY_MANAGED_REPORT, vbTextCompare) <> 0 Then detailText = "Managed-report category mismatch: " & reportName: Exit Function
    If ManagedReportCatalogClassificationCurrent(reportName, detailText) = False Then Exit Function
    ManagedReportCatalogDefinitionCurrent = True
    Exit Function
ErrHandler:
    detailText = "Could not resolve logical managed-report definition [" & reportName & "]. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
End Function

' Purpose: Reconciles managed report domain for setup.
' Public role: Public module entry point. This procedure is listed in the module header as an intentional externally callable ADC entry point and may be invoked directly by Access/VBA callers.
' Parameters:
'   detailText (Output Optional; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller. Default: vbNullString.
' Returns: Boolean: True on the function's affirmative/success branch for reconcile managed report domain for setup; False on its explicit negative, unavailable, rejected, or handled-failure branch.
'   not-applicable path.
Public Function ReconcileManagedReportDomainForSetup(Optional ByRef detailText As String = vbNullString) As Boolean
    RequireSetupDomainMutationAllowed SETUP_DOMAIN_E3_REPORT_FINALIZATION, "ReconcileManagedReportDomainForSetup"
    On Error GoTo ErrHandler
    Dim definitions As Collection
    Dim definitionData As Variant
    Dim selectedFields As Collection
    Dim currentDetail As String
    Dim reportName As String
    Dim payloadText As String
    ResetReportProofIfContextChanged
    If Len(mReportBuildToken) = 0 Then Err.Raise vbObjectError + 9880, "ReconcileManagedReportDomainForSetup", "E3 requires the same-run E2 physical-report proof."
    Set definitions = ManagedReportDefinitions()
    ResetReportSourceProjectionCaches
    For Each definitionData In definitions
        reportName = ManagedReportPhysicalName(CStr(definitionData(REPORT_DEF_LOGICAL_NAME)))
        Set selectedFields = ReportSelectedFields(definitionData)
        currentDetail = vbNullString
        If ManagedReportCurrent(definitionData, selectedFields, currentDetail) = False Then Err.Raise vbObjectError + 9882, "ReconcileManagedReportDomainForSetup", "E3 requires every E2-managed physical report to remain exactly current through reconciliation." & vbCrLf & currentDetail
        currentDetail = vbNullString
        If ManagedReportCatalogDefinitionCurrent(definitionData, currentDetail) = False Then Err.Raise vbObjectError + 9881, "ReconcileManagedReportDomainForSetup", currentDetail
        payloadText = payloadText & reportName & vbLf
        Set selectedFields = Nothing
    Next definitionData
    mReportReconciledToken = ReportContextToken("E3", mReportBuildToken & "|" & payloadText)
    detailText = "E3 reconciled all " & CStr(definitions.count) & " managed reports after freshly confirming each report's exact E2-owned physical definition and authoritative logical managed-object definition remained current. Catalog authority remained in Phase B."
    ReconcileManagedReportDomainForSetup = True
    Exit Function
ErrHandler:
    mReportReconciledToken = vbNullString
    detailText = "E3 managed-report reconciliation failed. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
End Function

' Purpose: Checks whether managed reports already satisfies the exact state required by the current Setup run so the owning phase can determine whether repair is necessary.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   detailText (Input/output Optional; String): Receives or supplies diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller. Default: vbNullString.
' Returns: Boolean: True when managed reports matches the exact state required by the current Setup contract; False when the owning phase still needs repair, refresh, or reconciliation.
Public Function ManagedReportsCurrentForSetup(Optional ByRef detailText As String = vbNullString) As Boolean
    On Error GoTo ErrHandler
    Dim definitions As Collection
    Dim definitionData As Variant
    Dim selectedFields As Collection
    Dim currentDetail As String
    ResetReportProofIfContextChanged
    Set definitions = ManagedReportDefinitions()
    If definitions.count <> REPORT_EXPECTED_COUNT Then detailText = "Expected nine managed report definitions; found " & CStr(definitions.count) & ".": Exit Function
    ResetReportSourceProjectionCaches
    For Each definitionData In definitions
        If ReportSourceExists(CStr(definitionData(REPORT_DEF_SOURCE_NAME)), CStr(definitionData(REPORT_DEF_SOURCE_KIND))) = False Then detailText = "Missing report source: " & CStr(definitionData(REPORT_DEF_SOURCE_NAME)): Exit Function
        Set selectedFields = ReportSelectedFields(definitionData)
        currentDetail = vbNullString
        If ManagedReportCurrent(definitionData, selectedFields, currentDetail) = False Then detailText = currentDetail: Exit Function
        If ManagedReportCatalogDefinitionCurrent(definitionData, currentDetail) = False Then detailText = currentDetail: Exit Function
        Set selectedFields = Nothing
    Next definitionData
    detailText = "All nine managed reports are physically current and correspond to canonical logical report definitions."
    ManagedReportsCurrentForSetup = True
    Exit Function
ErrHandler:
    detailText = "Managed-report currentness verification failed. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
End Function

' Purpose: Resolves managed report state token for final proof from the supplied/current ADC state, using `End` and `ManagedReportsCurrentForSetup` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   detailText (Output/in-out Optional; String): Carries human-readable diagnostic detail for the operation; when passed ByRef, the routine can return or augment the text for the caller. Default:
'   vbNullString.
' Returns: String: Returns the proof/state token produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function ManagedReportStateTokenForFinalProof(Optional ByRef detailText As String = vbNullString) As String
    Dim verificationDetail As String
    ResetReportProofIfContextChanged
    If ManagedReportsCurrentForSetup(verificationDetail) = False Then detailText = verificationDetail: Exit Function
    If Len(mReportReconciledToken) > 0 Then
        ManagedReportStateTokenForFinalProof = ReportContextToken("REPORTS_CURRENT", mReportReconciledToken & "|" & verificationDetail)
        detailText = "Final Setup Verification freshly reverified the complete managed-report physical/catalog domain and bound that current state to the same-run E3 reconciliation proof. " & verificationDetail
    Else
        ManagedReportStateTokenForFinalProof = ReportContextToken("REPORTS_CURRENT", verificationDetail)
        detailText = "Standalone Final Setup Verification freshly verified the complete managed-report physical/catalog domain read-only because no same-run E3 proof was available. " & verificationDetail
    End If
End Function
