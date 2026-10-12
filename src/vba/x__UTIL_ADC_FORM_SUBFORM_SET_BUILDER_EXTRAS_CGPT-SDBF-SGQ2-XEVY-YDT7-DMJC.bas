Attribute VB_Name = "x__UTIL_ADC_FORM_SUBFORM_SET_BUILDER_EXTRAS"
' ================================================================================================================================================================================
' ChatGPT Rewrite Tracking Header
' ================================================================================================================================================================================
' ChatGPT Previous Version ID: CGPT-HM86-TXHP-5C4B-TMA4-3KP5
' ChatGPT Baseline Version ID: CGPT-HM86-TXHP-5C4B-TMA4-3KP5
' ChatGPT Version ID: CGPT-SDBF-SGQ2-XEVY-YDT7-DMJC
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
' STANDARD MODULE: Access Development Console Form Set Builder
' VB_Name: x__UTIL_ADC_FORM_SUBFORM_SET_BUILDER_EXTRAS
' Export Group: Form Builder Tools
' Entry Points:
'     - BuildADCFormSetFromValues
'     - TryBuildADCFormSetFromValues
' SHARED CODE:
'     SharedAccess_ObjectIsOpenStrict(ByVal objectType As AcObjectType, ByVal objectName As String) -> Boolean
' CLASSES USED: None
' DYNAMIC STRING REFERENCES:
'     Application.Run string reference: WriteADCDiagnostic "BEFORE", "Application.Run generic builder", "API=" & GENERIC_BUILDER_CALLER_API & "; RecordSource=" & recordSourceName & "; RequestedFormSetName=" &
'     requestedFormSetName & "; SkipRuntimeTest=True; CompactRoleSuffixes=True; OverwriteExistingWithoutPrompt=True"
'     Application.Run string reference: WriteADCDiagnostic "AFTER", "Application.Run generic builder", "API=" & GENERIC_BUILDER_CALLER_API
'     Event property string reference: If Len(Trim$(Nz(ctl.OnClick, vbNullString))) > 0 Then problems = problems & " - " & ctl.Name & " must remain inert, but OnClick is not blank. ActualOnClick='" & Nz(ctl.OnClick,
'     vbNullString) & "'." & vbCrLf
'     Dynamic-reference scan stopped after 2000 logical line(s) to prevent a long-running scan; review this module manually if needed.
' LastScanTimestamp: Saturday 8/15/26 8:58 AM
' Purpose:
'     Access Development Console-specific wrapper around x__FORM_SUBFORM_SET_BUILDER.
'     Builds the normal five generated forms, then applies and verifies Access Development Console-specific post-build decoration.
' Author: Christina Shonts
' Support:
' Code Category:
' Requirements:
'     - x__FORM_SUBFORM_SET_BUILDER, x__SHARED_CODE
'     - GeneratedFormSetBuilderCallerOutcome must be public and callable through Access Application.Run.
'     - Generated parent forms must expose GeneratedRuntimeHealthCheck.
' Reference Dependencies: DAO, Access object model
' DB Object Dependencies: Selected table/query record source
' Usage: Call BuildADCFormSetFromValues from the copied Access Development Console form-set generator form.
' Version: Retains the fastest measured wrapper behavior while using the generic builder's single public master switch to suppress or enable GeneratedFormSetDiagnostic.log reset and append work;
'     enabled resets delete and recreate the log instead of truncating it.
' Notes:
'     - The wrapper calls GeneratedFormSetBuilderCallerOutcome once, requires a Committed result, decorates and verifies the generated set, and optionally opens the Both form in Design view.
'     - No transaction-state variables or prepare/seal/commit/rollback APIs are used or retained in this module.
'     - This module does not modify the generic form-set builder.
'     - This module generates the same five-form set as the generic builder.
'     - The caller-oriented Try function returns True only after generic generation, pre-decoration existence and closed-state verification, Access Development Console decoration, persisted Both-form
'        ownership/reference verification, closed-state verification, and runtime health all succeed.
'     - The generic builder verifies exact Final ownership for all five forms before returning Committed. This wrapper does not reopen all five forms to repeat that scan; it requires all five to exist and remain
'       closed, then reverifies the Both form's exact Final tag in the same Design-view pass used for decoration and again with the persisted label, popup properties, and child references.
'     - Access Development Console post-build controls are destructive/idempotent and are recreated on every completed build.
'     - The main - Both form receives two rows of inert numbered Access Development Console command buttons in its Form Footer. Each row contains one button per native stacked visual column in the generated Single
'       form; buttons are named ADC1Btn, ADC2Btn, and so on, captioned 1, 2, 3, and so on, and are created with no OnClick action.
'     - The generated description label is added only to the main - Both form.
'     - Missing object descriptions produce a visible generated placeholder.
'     - Form headers are preferred for description placement; detail-section fallback resizes its reserved top band in either direction and repositions existing detail controls so the label cannot overlap generated
'       controls or retain obsolete blank space.
'     - The record source and description share one label whose height is measured with Access Label.SizeToFit using the actual saved font, margins, caption, and available width. Word wrapping is measured before the
'       final label is created, so the label can occupy as many rows as required without reserving an unnecessary extra row.
'     - Description width aligns to the existing SingleForm and DSForm horizontal span at design time, then expands and contracts with the parent form by anchoring to both horizontal edges; the widest existing
'       content span is used only as a fallback, with tight outer margins and light internal padding.
'     - The decorated Both form remains a centered nonmodal PopUp with AutoResize and FitToScreen enabled so the full description band is visible when the form first opens.
'     - The generic builder completes the five-form set before Access Development Console decoration; wrapper failures preserve the completed generic forms for inspection.
'     - Post-decoration verification checks the decorated Both form's ownership, subform references, label persistence and placement, two-row command-button count/names/captions/inert event state/geometry, overlap,
'       popup properties, all-form closed state, and generated runtime health without repeating unchanged all-form tag scans.
'     - Creating the Form Header can make Access expand the parent Width and create a default-height blank Form Footer after the generic builder has already aligned the parent and subforms. Decoration performs one
'       width-alignment save-close-reopen cycle, then the final visible Design-view handoff compacts the open parent Width to the generated controls' shared right edge and saves once without another close-reopen
'       cycle. Detail still compacts to the lowest generated control and the unused Footer remains zero.
'     - Any persisted-decoration verification failure now includes a complete geometry and property snapshot for the parent, SingleForm, DSForm, and generated description label; diagnostics are fail-soft and cannot
'       replace or hide the original verification failure.
' Core Rules Enforced Here:
'     - Do not weaken the generic builder overwrite confirmation, open-form protection, or final verification checks.
'     - Do not infer generic build success from DateModified timestamps.
'     - Do not preserve prior generated Access Development Console controls.
'     - Do not fail the build solely because an object description is missing.
'     - Do not allow a generated description label to overlap another control in its section.
'     - Keep the generated record source and description in one normalized caption, measure the exact required wrapped row count at the saved form width, and anchor the label to both horizontal edges so it uses all
'       runtime width made available to the parent form.
'     - Generate exactly two footer button rows with one inert numbered button per generated Single-form visual column in each row. Button numbering is row-major and continuous.
'     - Require SingleForm, DSForm, and the generated description label to persist on one common right edge within the two-twip Access rounding tolerance; allow only bounded nonclipping Access-normalized parent
'       Width beyond that shared edge; require Detail to end at the lowest generated control and the unused Form Footer to persist at zero height; reject clipping, excessive width expansion, anchor drift, control
'       movement, or retained bottom blank space.
'     - Preserve the original verification error and append exact persisted measurements and property-read failures so repeated layout faults can be diagnosed without weakening fail-closed replacement behavior.
'     - Preserve PopUp, AutoCenter, AutoResize, FitToScreen, and nonmodal behavior on the decorated Both form.
'     - A completed UI-facing build must open only the newly created Both form in Design view. It must not display a success pop-up; non-UI callers can suppress the open.
' Design Map:
'     - Run the caller-oriented generic five-form builder.
'     - Verify all five exact Final ownership tags and closed state.
'     - Apply the Access Development Console description header only to the completed main - Both form.
'     - Prefer a true form header sized to the measured multi-row label band; after persistence, accept the Access 2024 minimum Form Header height when Access normalizes a smaller requested band on save/reopen. If a
'       header is unavailable, reserve an equivalent detail-section band and remove obsolete excess band height.
'     - Count the generated Single form's separate native stacked layouts, then create two equally distributed footer-button rows across the same horizontal span as SingleForm and DSForm.
'     - Reassert centered auto-sized pop-up properties on the decorated Both form and verify them after persistence.
'     - Verify the decorated Both form in one Design-view pass, run the parent runtime health check, and finish with a lightweight all-form closed-state check.
' Configuration:
'     - Label placement, tight outer margins, light internal padding, sizing, caption limits, and two-row Access Development Console footer-button dimensions are declared in this module.
' Behavior:
'     - The lblGeneratedRecordSourceDescription control is deleted and recreated only on the main - Both form.
'     - ADC1Btn through ADC{2 x visual-column-count}Btn are deleted and recreated in two Form Footer rows on every Access Development Console build. Their captions are the matching numbers and their OnClick
'       properties remain blank.
'     - The record source is followed immediately by the description in the same normalized caption. The label height is derived from Access's own SizeToFit measurements and measured word wrapping; ordinary line-fit
'     tests use one width measurement, and prefix search is reserved for oversized unbreakable tokens. The label is horizontally anchored to both sides of the parent form.
'     - After Form Header and footer-button creation, the parent Width receives one real save-close-reopen alignment check, while Detail height and used Footer height retain bounded stabilization. If Access expands
'       the parent Width again, the final visible Design-view handoff sets the open parent Width back to the generated controls' shared right edge and saves once without closing and reopening it again.
'     - The UI-facing entry point opens the completed Both form in Design view instead of displaying a success message. The legacy success-message argument is retained only to preserve the existing public signature.
'     - If a form header section is unavailable, the module falls back to the detail section and shifts existing detail controls up or down by the exact band-height change required.
'     - When ENABLE_GENERATED_FORM_SET_DIAGNOSTIC_FILE_LOGGING=True in x__FORM_SUBFORM_SET_BUILDER, GeneratedFormSetDiagnostic.log is deleted and recreated before the first diagnostic entry so it
'       contains only the current run.
'       When False, this wrapper performs no reset or append work and leaves any existing file unchanged.
'     - The generic caller API is invoked with skipRuntimeTest=True. The wrapper retains its post-decoration runtime-health check, so the supported Access Development Console workflow still runtime-tests the actual
'       final decorated Both form once.
'     - Optional control properties fail soft with detailed Immediate Window diagnostics.
' Code Status:
' Todo:
'     - Attach project-specific procedures to the generated numbered buttons only in downstream edited forms; the generator intentionally leaves every OnClick property blank.
' Reporting Bugs:
' Module Level Bugs: None Known
' Public Subs:
'     BuildADCFormSetFromValues(ByVal recordSourceName As String, ByVal requestedFormSetName As String, Optional ByVal keyFieldName As String = FIELD_ID, Optional ByVal showADCSuccessMessage As
'     Boolean = True) -> None
' Public Functions:
'     FormNameFromSourceObject(ByVal sourceObjectText As String) -> String
'     ADCFormSetBuilderCallerOutcome(ByVal recordSourceName As String, ByVal requestedFormSetName As String, Optional ByVal keyFieldName As String = "") -> Variant
'     TryBuildADCFormSetFromValues(ByVal recordSourceName As String, ByVal requestedFormSetName As String, ByVal keyFieldName As String, ByRef completedBuildId As String, ByRef normalizedFormSetName As String,
'     ByRef outcomeCode As String, ByRef outcomeDetails As String, Optional ByVal openCommittedBothInDesignView As Boolean = True) -> Boolean
' Public Properties: None
' ================================================================================================================================================================================

Option Compare Database
Option Explicit

Private Const MODULE_NAME As String = "x__UTIL_ADC_FORM_SET_BUILDER"
Private Const DIAGNOSTIC_LOG_FILE_NAME As String = "GeneratedFormSetDiagnostic.log"
Private Const FORM_INDEX_BOTH As Long = 1
Private Const FORM_INDEX_CONT As Long = 2
Private Const FORM_INDEX_DS As Long = 3
Private Const FORM_INDEX_FORM As Long = 4
Private Const FORM_INDEX_SPLIT As Long = 5
Private Const SUFFIX_BOTH As String = " Both"
Private Const SUFFIX_CONT As String = " Cont"
Private Const SUFFIX_DS As String = " DS"
Private Const SUFFIX_FORM As String = " Form"
Private Const SUFFIX_SPLIT As String = " Split"
Private Const SINGLE_SUBFORM_CONTROL_NAME As String = "SingleForm"
Private Const DS_SUBFORM_CONTROL_NAME As String = "DSForm"
Private Const DESCRIPTION_LABEL_NAME As String = "lblGeneratedRecordSourceDescription"
Private Const DESCRIPTION_LABEL_LEFT_TWIPS As Long = 30
Private Const DESCRIPTION_LABEL_TOP_TWIPS As Long = 15
Private Const DESCRIPTION_LABEL_RIGHT_PADDING_TWIPS As Long = 30
Private Const DESCRIPTION_LABEL_BOTTOM_PADDING_TWIPS As Long = 15
' Access 2024 can persist a newly created Form Header no shorter than 585 twips after save/close/reopen even when a smaller height was accepted in Design view.
Private Const PERSISTED_FORM_HEADER_MINIMUM_TWIPS As Long = 585
Private Const DESCRIPTION_LABEL_INNER_HORIZONTAL_PADDING_TWIPS As Long = 30
Private Const DESCRIPTION_LABEL_INNER_VERTICAL_PADDING_TWIPS As Long = 15
Private Const DESCRIPTION_LABEL_HORIZONTAL_ANCHOR_BOTH As Long = 2
Private Const DECORATED_PARENT_RIGHT_EDGE_TOLERANCE_TWIPS As Long = 2
Private Const DECORATED_PARENT_BOTTOM_EDGE_TOLERANCE_TWIPS As Long = 2
Private Const MAX_DECORATED_PARENT_WIDTH_EXPANSION_TWIPS As Long = 1440
Private Const DECORATED_PARENT_BOTTOM_STABILIZATION_PASSES As Long = 3
Private Const ADC_BUTTON_NAME_PREFIX As String = "ADC"
Private Const ADC_BUTTON_NAME_SUFFIX As String = "Btn"
Private Const ADC_BUTTON_ROW_COUNT As Long = 2
Private Const ADC_BUTTON_HEIGHT_TWIPS As Long = 360
Private Const ADC_BUTTON_HORIZONTAL_GAP_TWIPS As Long = 45
Private Const ADC_BUTTON_VERTICAL_GAP_TWIPS As Long = 45
Private Const ADC_BUTTON_TOP_MARGIN_TWIPS As Long = 45
Private Const ADC_BUTTON_BOTTOM_MARGIN_TWIPS As Long = 45
Private Const ADC_BUTTON_FONT_SIZE As Long = 10
Private Const DESCRIPTION_MEASUREMENT_LABEL_NAME As String = "lblGeneratedDescriptionMeasurement"
Private Const DESCRIPTION_FONT_SIZE As Long = 10
Private Const DESCRIPTION_MISSING_TEXT As String = "[Missing object description]"
Private Const DESCRIPTION_UNAVAILABLE_TEXT As String = "[Object description unavailable]"
Private Const MAX_CONTROL_TIP_LEN As Long = 255
Private Const MAX_ACCESS_SECTION_HEIGHT_TWIPS As Long = 31680
Private Const GENERIC_TAG_PREFIX As String = "[GeneratedBy=x__FORM_SUBFORM_SET_BUILDER]"
Private Const GENERIC_FINAL_ROLE_TOKEN As String = "[Role=Final]"
Private Const GENERIC_BUILD_ID_PREFIX As String = "[BuildId="
Private Const GENERIC_FORM_NAME_PREFIX As String = "[FormName="
Private Const OUTCOME_COMMITTED As String = "Committed"
Private Const OUTCOME_CANCELLED As String = "Cancelled"
Private Const OUTCOME_REJECTED As String = "Rejected"
Private Const OUTCOME_GENERIC_FAILED As String = "GenericFailed"
Private Const OUTCOME_DECORATION_FAILED As String = "DecorationFailed"
Private Const EVENT_PROCEDURE_RESULT_METHOD As String = "GeneratedRuntimeHealthCheck"
Private Const GENERIC_BUILDER_CALLER_API As String = "GeneratedFormSetBuilderCallerOutcome"
Private Const GENERIC_BUILDER_STAMP_CONTRACT_API As String = "StampGeneratedFormSetBuildContract"
' Bump only when ADC decoration changes persisted physical output in a way that makes an existing decorated set stale; diagnostics/reporting changes do not belong in this contract.
Private Const ADC_DECORATION_PHYSICAL_CONTRACT_VERSION As String = "2026-08-07 ADC decoration contract v1"

' =============================================================================================================================================================================
' PUBLIC BUILD WRAPPERS / OUTCOME HANDOFF PROCEDURES
' =============================================================================================================================================================================

' Intended use: UI-facing Access Development Console entry point that opens the committed Both form in Design view and displays messages only for cancellation, rejection, failure, or post-commit open errors.
' Purpose: Builds ADC form set from values for record source name, requested form set name.
' Public role: Public ADC API. The routine is intentionally callable outside its defining procedure/module and may be used directly by Access/VBA or late-bound callers when that API surface is needed.
' Parameters:
'   recordSourceName (Input; String): Physical table/query name assigned or validated as the generated form/report RecordSource.
'   requestedFormSetName (Input; String): Logical/base form-set name requested from the form/subform builder; it drives the generated parent/child form names.
'   keyFieldName (Input Optional; String): Field name used as the record key/identity when constructing or binding the generated form set. Default: FIELD_ID.
'   showADCSuccessMessage (Input Optional; Boolean): True displays ADC success message; False suppresses that user-facing presentation without changing the underlying work. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub BuildADCFormSetFromValues(ByVal recordSourceName As String, ByVal requestedFormSetName As String, Optional ByVal keyFieldName As String = FIELD_ID, Optional ByVal showADCSuccessMessage As Boolean = True)
    On Error GoTo ErrHandler
    Dim completedBuildId As String
    Dim normalizedFormSetName As String
    Dim outcomeCode As String
    Dim outcomeDetails As String
    Dim buildCommitted As Boolean
    Dim messageStyle As VbMsgBoxStyle
    Dim traceLevel As String
    buildCommitted = TryBuildADCFormSetFromValues(recordSourceName, requestedFormSetName, keyFieldName, completedBuildId, normalizedFormSetName, outcomeCode, outcomeDetails, True)
    If buildCommitted Then
        If Not showADCSuccessMessage Then Debug.Print MODULE_NAME & ": Legacy success-message argument is ignored; committed UI builds open the Both form in Design view."
        Exit Sub
    End If
    Select Case outcomeCode
        Case OUTCOME_CANCELLED
            messageStyle = vbInformation
            traceLevel = "Info"
        Case OUTCOME_REJECTED
            messageStyle = vbExclamation
            traceLevel = "Warning"
        Case Else
            messageStyle = vbCritical
            traceLevel = "Error"
    End Select
    ShowFinalADCMessage outcomeDetails, messageStyle, "Access Development Console Form Set Builder", traceLevel, "BuildADCFormSetFromValues", outcomeCode, 0, vbNullString, vbNullString, normalizedFormSetName
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    DoCmd.Hourglass False
    ClearADCStatus
    On Error GoTo 0
    ShowFinalADCMessage "The Access Development Console form set was committed when applicable, but the final Both form could not be opened in Design view or the final outcome could not be reported." & vbCrLf & vbCrLf & "Form: " & normalizedFormSetName & SUFFIX_BOTH & vbCrLf & "Error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription, vbCritical, "Access Development Console Form Set Builder", "Error", "BuildADCFormSetFromValues", "Open committed Both form in Design view", savedErrNumber, savedErrSource, savedErrDescription, normalizedFormSetName
End Sub

Private Sub OpenCommittedADCBothFormInDesignView(ByVal formName As String)
    On Error GoTo ErrHandler
    If Len(Trim$(formName)) = 0 Then Err.Raise vbObjectError + 875, MODULE_NAME & ".OpenCommittedADCBothFormInDesignView", "The committed Both-form name is blank."
    If Not FormExists(formName) Then Err.Raise vbObjectError + 876, MODULE_NAME & ".OpenCommittedADCBothFormInDesignView", "The committed Both form was not found: " & formName
    If SharedAccess_ObjectIsOpenStrict(acForm, formName) Then Err.Raise vbObjectError + 877, MODULE_NAME & ".OpenCommittedADCBothFormInDesignView", "The committed Both form is unexpectedly already open, so the builder will not change or close its current view: " & formName
    WriteADCDiagnostic "BEFORE", "DoCmd.OpenForm final Both form", formName
    DoCmd.OpenForm formName, acDesign, , , , acWindowNormal
    WriteADCDiagnostic "AFTER", "DoCmd.OpenForm final Both form", formName
    If Not SharedAccess_ObjectIsOpenStrict(acForm, formName) Then Err.Raise vbObjectError + 878, MODULE_NAME & ".OpenCommittedADCBothFormInDesignView", "Access returned without leaving the committed Both form open in Design view: " & formName
    WriteADCDiagnostic "BEFORE", "DoCmd.SelectObject final Both form", formName
    DoCmd.SelectObject acForm, formName, False
    WriteADCDiagnostic "AFTER", "DoCmd.SelectObject final Both form", formName
    If Application.currentObjectType <> acForm Or StrComp(Application.currentObjectName, formName, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 879, MODULE_NAME & ".OpenCommittedADCBothFormInDesignView", "The committed Both form is open but Access did not activate it as the current Design-view object. Expected=" & formName & "; CurrentObjectType=" & CStr(Application.currentObjectType) & "; CurrentObjectName=" & Application.currentObjectName & "."
    WriteADCDiagnostic "BEFORE", "Compact final open Both form right edge", formName
    CompactFinalOpenBothFormRightEdge forms(formName), formName
    WriteADCDiagnostic "AFTER", "Compact final open Both form right edge", formName
    Debug.Print MODULE_NAME & ": Opened, activated, and right-edge compacted committed Access Development Console Both form in Design view: " & formName
    Exit Sub
ErrHandler:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Raise errNumber, MODULE_NAME & ".OpenCommittedADCBothFormInDesignView", "The Access Development Console form set was committed, but its Both form could not be opened in Design view. Form=" & formName & ". " & errDescription
End Sub

Private Sub CompactFinalOpenBothFormRightEdge(ByVal frm As Access.Form, ByVal formName As String)
    On Error GoTo ErrHandler
    Dim ctl As Access.control
    Dim desiredRight As Long
    Dim controlRight As Long
    Dim widthBefore As Long
    Dim widthAfter As Long
    desiredRight = CLng(frm.Controls(SINGLE_SUBFORM_CONTROL_NAME).Left) + CLng(frm.Controls(SINGLE_SUBFORM_CONTROL_NAME).Width)
    controlRight = CLng(frm.Controls(DS_SUBFORM_CONTROL_NAME).Left) + CLng(frm.Controls(DS_SUBFORM_CONTROL_NAME).Width)
    If controlRight > desiredRight Then desiredRight = controlRight
    controlRight = CLng(frm.Controls(DESCRIPTION_LABEL_NAME).Left) + CLng(frm.Controls(DESCRIPTION_LABEL_NAME).Width)
    If controlRight > desiredRight Then desiredRight = controlRight
    For Each ctl In frm.Controls
        If IsGeneratedADCButtonName(ctl.Name) Then
            controlRight = CLng(ctl.Left) + CLng(ctl.Width)
            If controlRight > desiredRight Then desiredRight = controlRight
        End If
    Next ctl
    widthBefore = CLng(frm.Width)
    If widthBefore < desiredRight - DECORATED_PARENT_RIGHT_EDGE_TOLERANCE_TWIPS Then Err.Raise vbObjectError + 883, MODULE_NAME & ".CompactFinalOpenBothFormRightEdge", "The final open Both form Width is smaller than the generated controls' shared right edge. Form=" & formName & "; Width=" & CStr(widthBefore) & "; RequiredRightEdge=" & CStr(desiredRight) & "."
    If widthBefore > desiredRight + DECORATED_PARENT_RIGHT_EDGE_TOLERANCE_TWIPS Then
        frm.Width = desiredRight
        DoCmd.Save acForm, formName
    End If
    widthAfter = CLng(frm.Width)
    If widthAfter < desiredRight - DECORATED_PARENT_RIGHT_EDGE_TOLERANCE_TWIPS Then Err.Raise vbObjectError + 884, MODULE_NAME & ".CompactFinalOpenBothFormRightEdge", "The final right-edge compaction clipped the generated controls. Form=" & formName & "; Width=" & CStr(widthAfter) & "; RequiredRightEdge=" & CStr(desiredRight) & "."
    Debug.Print MODULE_NAME & ": Final open Both form right edge compacted without another close-reopen cycle. Form=" & formName & "; WidthBefore=" & CStr(widthBefore) & "; WidthAfter=" & CStr(widthAfter) & "; GeneratedRightEdge=" & CStr(desiredRight)
    Exit Sub
ErrHandler:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Raise errNumber, MODULE_NAME & ".CompactFinalOpenBothFormRightEdge", "Unable to compact the final open Both form to the generated controls' right edge. Form=" & formName & ". " & errDescription
End Sub

Private Function TryOpenCommittedADCBothFormInDesignView(ByVal formName As String, ByRef failureDetails As String) As Boolean
    On Error GoTo ErrHandler
    failureDetails = vbNullString
    OpenCommittedADCBothFormInDesignView formName
    TryOpenCommittedADCBothFormInDesignView = True
    Exit Function
ErrHandler:
    failureDetails = "The Access Development Console form set was committed, but its Both form could not be opened and activated in Design view." & vbCrLf & vbCrLf & "Form: " & formName & vbCrLf & "Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    Debug.Print MODULE_NAME & ": Post-commit Design-view open failed. " & Replace(failureDetails, vbCrLf, " | ")
End Function

' Purpose: Performs ADC form set builder caller outcome by coordinating `TryBuildADCFormSetFromValues` and `End` while preserving this module's existing validation and error semantics.
' Public role: Public ADC API. The routine is intentionally callable outside its defining procedure/module and may be used directly by Access/VBA or late-bound callers when that API surface is needed.
' Parameters:
'   recordSourceName (Input; String): Physical table/query name assigned or validated as the generated form/report RecordSource.
'   requestedFormSetName (Input; String): Logical/base form-set name requested from the form/subform builder; it drives the generated parent/child form names.
'   keyFieldName (Input Optional; String): Field name used as the record key/identity when constructing or binding the generated form set. Default: "".
' Returns: Variant: Returns the form-set builder outcome payload expected by its caller, carrying the committed form/subform result state in the representation established by this helper.
Public Function ADCFormSetBuilderCallerOutcome(ByVal recordSourceName As String, ByVal requestedFormSetName As String, Optional ByVal keyFieldName As String = "") As Variant
    ' Wrapper purpose: Caller-oriented facade that converts the ADC builder's Boolean plus multiple ByRef outputs into one compact Variant-array outcome contract.
    ' Why retained: Retained for callers that need a single returned value without reproducing the builder's ByRef plumbing or losing the complete ADC outcome details.
    Dim completedBuildId As String
    Dim normalizedFormSetName As String
    Dim outcomeCode As String
    Dim outcomeDetails As String
    Dim buildCommitted As Boolean
    buildCommitted = TryBuildADCFormSetFromValues(recordSourceName, requestedFormSetName, keyFieldName, completedBuildId, normalizedFormSetName, outcomeCode, outcomeDetails, False)
    ADCFormSetBuilderCallerOutcome = Array(buildCommitted, completedBuildId, normalizedFormSetName, outcomeCode, outcomeDetails)
End Function

Private Function ADCGeneratedFormSetDecoratorContractToken(ByVal recordSourceName As String, ByVal requestedFormSetName As String, Optional ByVal keyFieldName As String = "") As String
    ' Wrapper purpose: Private ADC adapter that derives the physical-decoration compatibility token from normalized form-set inputs and the current record-source description label.
    ' Why retained: Retained to centralize the decorator compatibility contract so build/reuse checks do not duplicate token composition or drift from ADC decoration semantics.
    Dim descriptionText As String
    Dim descriptionStatus As String
    recordSourceName = Trim$(recordSourceName)
    requestedFormSetName = Trim$(requestedFormSetName)
    keyFieldName = NormalizeFieldInput(keyFieldName)
    If Len(keyFieldName) = 0 Then keyFieldName = FIELD_ID
    descriptionText = RecordSourceDescriptionForLabel(recordSourceName, descriptionStatus)
    ADCGeneratedFormSetDecoratorContractToken = ADC_DECORATION_PHYSICAL_CONTRACT_VERSION & "|" & requestedFormSetName & "|" & keyFieldName & "|" & HeaderLabelCaption(recordSourceName, descriptionText)
End Function

' Intended use: Caller-oriented API that returns True only when the generic form set and all Access Development Console decoration and verification steps commit.
' Purpose: Attempts to perform build ADC form set from values for record source name, requested form set name.
' Public role: Public ADC API. The routine is intentionally callable outside its defining procedure/module and may be used directly by Access/VBA or late-bound callers when that API surface is needed.
' Parameters:
'   recordSourceName (Input; String): Physical table/query name assigned or validated as the generated form/report RecordSource.
'   requestedFormSetName (Input; String): Logical/base form-set name requested from the form/subform builder; it drives the generated parent/child form names.
'   keyFieldName (Input; String): Field name used as the record key/identity when constructing or binding the generated form set.
'   completedBuildId (Output; String): Receives the resolved completed build ID produced from this procedure's validated inputs and current ADC state.
'   normalizedFormSetName (Output; String): Receives the resolved normalized form set name produced from this procedure's validated inputs and current ADC state.
'   outcomeCode (Output; String): Receives the resolved outcome code produced from this procedure's validated inputs and current ADC state.
'   outcomeDetails (Output; String): Receives outcome details used for caller diagnostics, verification output, or user-facing reporting.
'   openCommittedBothInDesignView (Input Optional; Boolean): True opens committed both in design view as part of this call; False leaves it unopened and follows the normal completion path. Default: True.
' Returns: Boolean: True when the attempt to build ADC form set from values completes successfully; False for the routine's defined normal unsuccessful/not-applicable outcome.
Public Function TryBuildADCFormSetFromValues(ByVal recordSourceName As String, ByVal requestedFormSetName As String, ByVal keyFieldName As String, ByRef completedBuildId As String, ByRef normalizedFormSetName As String, ByRef outcomeCode As String, ByRef outcomeDetails As String, Optional ByVal openCommittedBothInDesignView As Boolean = True) As Boolean
    On Error GoTo ErrHandler
    Dim formNames(1 To 5) As String
    Dim descriptionText As String
    Dim descriptionStatus As String
    Dim currentAction As String
    Dim startedAt As Date
    Dim genericBuildCompleted As Boolean
    Dim genericOutcomeCode As String
    Dim genericOutcomeDetails As String
    Dim ADCCompleted As Boolean
    Dim designOpenFailure As String
    Dim visualColumnCount As Long
    completedBuildId = vbNullString
    normalizedFormSetName = vbNullString
    outcomeCode = vbNullString
    outcomeDetails = vbNullString
    startedAt = Now
    recordSourceName = Trim$(recordSourceName)
    requestedFormSetName = Trim$(requestedFormSetName)
    keyFieldName = NormalizeFieldInput(keyFieldName)
    If Len(keyFieldName) = 0 Then keyFieldName = FIELD_ID
    ResetADCDiagnosticLog
    WriteADCDiagnostic "RUN", "TryBuildADCFormSetFromValues entered", "RecordSource=" & recordSourceName & "; RequestedFormSetName=" & requestedFormSetName & "; KeyFieldName=" & keyFieldName & "; OpenBothInDesignView=" & CStr(openCommittedBothInDesignView)
    Debug.Print String$(70, "=")
    Debug.Print MODULE_NAME & ": Access Development Console form-set build started " & Format$(startedAt, "yyyy-mm-dd hh:nn:ss")
    Debug.Print MODULE_NAME & ": RecordSource=" & recordSourceName
    Debug.Print MODULE_NAME & ": RequestedFormSetName=" & requestedFormSetName
    Debug.Print MODULE_NAME & ": KeyFieldName=" & keyFieldName
    currentAction = "Generate generic form set"
    SetADCStatus currentAction, requestedFormSetName
    ReadGenericBuilderOutcome recordSourceName, requestedFormSetName, keyFieldName, genericBuildCompleted, completedBuildId, normalizedFormSetName, genericOutcomeCode, genericOutcomeDetails
    WriteADCDiagnostic "COMPLETE", currentAction, "Committed=" & CStr(genericBuildCompleted) & "; BuildId=" & completedBuildId & "; FormSetName=" & normalizedFormSetName & "; Outcome=" & genericOutcomeCode
    Debug.Print MODULE_NAME & ": GenericBuilderOutcome=" & genericOutcomeCode
    Debug.Print MODULE_NAME & ": GenericBuilderBuildId=" & completedBuildId
    Debug.Print MODULE_NAME & ": GenericBuilderNormalizedFormSetName=" & normalizedFormSetName
    Debug.Print MODULE_NAME & ": GenericBuilderDetails=" & Replace(genericOutcomeDetails, vbCrLf, " | ")
    If Not genericBuildCompleted Then
        Select Case genericOutcomeCode
            Case OUTCOME_CANCELLED
                outcomeCode = OUTCOME_CANCELLED
            Case OUTCOME_REJECTED
                outcomeCode = OUTCOME_REJECTED
            Case Else
                outcomeCode = OUTCOME_GENERIC_FAILED
        End Select
        outcomeDetails = "The generic form set was not completed, so Access Development Console decoration was not applied." & vbCrLf & vbCrLf & "Generic outcome: " & genericOutcomeCode
        If Len(genericOutcomeDetails) > 0 Then outcomeDetails = outcomeDetails & vbCrLf & vbCrLf & genericOutcomeDetails
        Debug.Print MODULE_NAME & ": Access Development Console decoration skipped. Outcome=" & outcomeCode & "; Details=" & Replace(outcomeDetails, vbCrLf, " | ")
        ClearADCStatus
        DoCmd.Hourglass False
        Exit Function
    End If
    If StrComp(genericOutcomeCode, OUTCOME_COMMITTED, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 810, MODULE_NAME & ".TryBuildADCFormSetFromValues", "The generic builder returned success without the required Committed outcome. Outcome=" & genericOutcomeCode
    If Len(completedBuildId) = 0 Then Err.Raise vbObjectError + 811, MODULE_NAME & ".TryBuildADCFormSetFromValues", "The generic builder completed but returned no BuildId."
    If Len(normalizedFormSetName) = 0 Then Err.Raise vbObjectError + 812, MODULE_NAME & ".TryBuildADCFormSetFromValues", "The generic builder completed but returned no normalized form-set name. BuildId=" & completedBuildId
    MakeFormNames normalizedFormSetName, formNames
    currentAction = "Verify completed generic form set is closed"
    SetADCStatus currentAction, normalizedFormSetName
    VerifyAllGeneratedFormsClosed formNames, "Before Access Development Console decoration"
    WriteADCDiagnostic "COMPLETE", currentAction, normalizedFormSetName
    currentAction = "Count generated visual columns"
    SetADCStatus currentAction, formNames(FORM_INDEX_FORM)
    WriteADCDiagnostic "BEFORE", "GeneratedVisualColumnCount", formNames(FORM_INDEX_FORM)
    visualColumnCount = GeneratedVisualColumnCount(formNames(FORM_INDEX_FORM))
    WriteADCDiagnostic "AFTER", "GeneratedVisualColumnCount", "Form=" & formNames(FORM_INDEX_FORM) & "; Count=" & CStr(visualColumnCount)
    currentAction = "Read object description"
    SetADCStatus currentAction, recordSourceName
    descriptionText = RecordSourceDescriptionForLabel(recordSourceName, descriptionStatus)
    Debug.Print MODULE_NAME & ": DescriptionStatus=" & descriptionStatus
    currentAction = "Apply generated description header to main form"
    SetADCStatus currentAction, normalizedFormSetName
    ApplyDescriptionHeaderToMainForm formNames, recordSourceName, descriptionText, completedBuildId, visualColumnCount
    WriteADCDiagnostic "COMPLETE", currentAction, normalizedFormSetName
    currentAction = "Verify decorated generated form set"
    SetADCStatus currentAction, normalizedFormSetName
    VerifyDecoratedGeneratedFormSet formNames, completedBuildId, HeaderLabelCaption(recordSourceName, descriptionText), visualColumnCount
    WriteADCDiagnostic "COMPLETE", currentAction, normalizedFormSetName
    currentAction = "Stamp completed physical form-set contract"
    SetADCStatus currentAction, normalizedFormSetName
    Application.Run GENERIC_BUILDER_STAMP_CONTRACT_API, recordSourceName, normalizedFormSetName, keyFieldName, ADCGeneratedFormSetDecoratorContractToken(recordSourceName, normalizedFormSetName, keyFieldName), True
    WriteADCDiagnostic "COMPLETE", currentAction, normalizedFormSetName
    ADCCompleted = True
    outcomeCode = OUTCOME_COMMITTED
    outcomeDetails = "Access Development Console form-set build completed." & vbCrLf & vbCrLf & "Record source: " & recordSourceName & vbCrLf & "Form set name: " & normalizedFormSetName & vbCrLf & "Key field: " & keyFieldName & vbCrLf & "BuildId: " & completedBuildId & vbCrLf & "Description status: " & descriptionStatus & vbCrLf & "Elapsed seconds: " & Format$(DateDiff("s", startedAt, Now), "0")
    If openCommittedBothInDesignView Then
        currentAction = "Open completed Both form in Design view"
        SetADCStatus currentAction, normalizedFormSetName
        If Not TryOpenCommittedADCBothFormInDesignView(normalizedFormSetName & SUFFIX_BOTH, designOpenFailure) Then
            outcomeDetails = outcomeDetails & vbCrLf & vbCrLf & designOpenFailure
            ShowFinalADCMessage designOpenFailure, vbCritical, "Access Development Console Form Set Builder", "Error", "TryBuildADCFormSetFromValues", currentAction, 0, MODULE_NAME & ".OpenCommittedADCBothFormInDesignView", designOpenFailure, normalizedFormSetName
        End If
    End If
    Debug.Print MODULE_NAME & ": Access Development Console post-build customization completed " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Debug.Print MODULE_NAME & ": Elapsed seconds=" & Format$(DateDiff("s", startedAt, Now), "0")
    Debug.Print String$(70, "=")
    ClearADCStatus
    DoCmd.Hourglass False
    TryBuildADCFormSetFromValues = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    Dim savedErrSource As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    savedErrSource = Err.Source
    On Error Resume Next
    DoCmd.Hourglass False
    ClearADCStatus
    On Error GoTo 0
    If ADCCompleted Then
        outcomeCode = OUTCOME_COMMITTED
        outcomeDetails = "Access Development Console form-set build completed successfully, but final completion reporting failed." & vbCrLf & vbCrLf & "BuildId: " & completedBuildId & vbCrLf & "Form set name: " & normalizedFormSetName & vbCrLf & "Reporting error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription
        TryBuildADCFormSetFromValues = True
        Exit Function
    End If
    If genericBuildCompleted Then
        outcomeCode = OUTCOME_DECORATION_FAILED
    Else
        outcomeCode = OUTCOME_GENERIC_FAILED
    End If
    outcomeDetails = BuildFailureMessage(currentAction, recordSourceName, requestedFormSetName, normalizedFormSetName, keyFieldName, completedBuildId, savedErrSource, savedErrNumber, savedErrDescription, vbNullString)
    If genericBuildCompleted Then outcomeDetails = outcomeDetails & vbCrLf & vbCrLf & "The completed generic forms were preserved. No rollback or additional deletion was attempted."
    WriteADCDiagnostic "ERROR", "Build failed during " & currentAction, "RecordSource=" & recordSourceName & "; RequestedFormSetName=" & requestedFormSetName & "; NormalizedFormSetName=" & normalizedFormSetName & "; BuildId=" & completedBuildId & "; Error=" & CStr(savedErrNumber) & "; Source=" & savedErrSource & "; Description=" & savedErrDescription
    Debug.Print MODULE_NAME & ": Build failed during action: " & currentAction
    Debug.Print MODULE_NAME & ": RecordSource=" & recordSourceName
    Debug.Print MODULE_NAME & ": RequestedFormSetName=" & requestedFormSetName
    Debug.Print MODULE_NAME & ": NormalizedFormSetName=" & normalizedFormSetName
    Debug.Print MODULE_NAME & ": KeyFieldName=" & keyFieldName
    Debug.Print MODULE_NAME & ": BuildId=" & completedBuildId
    Debug.Print MODULE_NAME & ": Error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription
End Function

Private Sub ReadGenericBuilderOutcome(ByVal recordSourceName As String, ByVal requestedFormSetName As String, ByVal keyFieldName As String, ByRef buildCommitted As Boolean, ByRef completedBuildId As String, ByRef normalizedFormSetName As String, ByRef outcomeCode As String, ByRef outcomeDetails As String)
    On Error GoTo ErrHandler
    Dim outcomePacket As Variant
    WriteADCDiagnostic "BEFORE", "Application.Run generic builder", "API=" & GENERIC_BUILDER_CALLER_API & "; RecordSource=" & recordSourceName & "; RequestedFormSetName=" & requestedFormSetName & "; SkipRuntimeTest=True; CompactRoleSuffixes=True; OverwriteExistingWithoutPrompt=True"
    outcomePacket = Application.Run(GENERIC_BUILDER_CALLER_API, recordSourceName, requestedFormSetName, keyFieldName, True, True, True)
    WriteADCDiagnostic "AFTER", "Application.Run generic builder", "API=" & GENERIC_BUILDER_CALLER_API
    If Not IsArray(outcomePacket) Then Err.Raise vbObjectError + 836, MODULE_NAME & ".ReadGenericBuilderOutcome", "The generic builder caller API did not return an outcome array."
    If LBound(outcomePacket) <> 0 Or UBound(outcomePacket) < 4 Then Err.Raise vbObjectError + 837, MODULE_NAME & ".ReadGenericBuilderOutcome", "The generic builder caller API returned an incomplete outcome array."
    buildCommitted = CBool(outcomePacket(0))
    completedBuildId = CStr(outcomePacket(1))
    normalizedFormSetName = CStr(outcomePacket(2))
    outcomeCode = CStr(outcomePacket(3))
    outcomeDetails = CStr(outcomePacket(4))
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, MODULE_NAME & ".ReadGenericBuilderOutcome", "Unable to run generic builder caller API '" & GENERIC_BUILDER_CALLER_API & "'. Confirm that the current x__FORM_SUBFORM_SET_BUILDER replacement is installed in a standard module." & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Sub

Private Function BuildFailureMessage(ByVal currentAction As String, ByVal recordSourceName As String, ByVal requestedFormSetName As String, ByVal normalizedFormSetName As String, ByVal keyFieldName As String, ByVal completedBuildId As String, ByVal errSource As String, ByVal errNumber As Long, ByVal errDescription As String, ByVal restoreSummary As String) As String
    Dim result As String
    result = "Access Development Console form-set build failed." & vbCrLf & vbCrLf & "Action: " & currentAction & vbCrLf & "Record source: " & recordSourceName & vbCrLf & "Requested form set name: " & requestedFormSetName
    If Len(normalizedFormSetName) > 0 Then result = result & vbCrLf & "Normalized form set name: " & normalizedFormSetName
    result = result & vbCrLf & "Key field: " & keyFieldName
    If Len(completedBuildId) > 0 Then result = result & vbCrLf & "BuildId: " & completedBuildId
    result = result & vbCrLf & vbCrLf & "Source: " & errSource & vbCrLf & "Error " & CStr(errNumber) & ": " & errDescription
    If Len(restoreSummary) > 0 Then result = result & vbCrLf & vbCrLf & "Transaction recovery:" & vbCrLf & restoreSummary
    BuildFailureMessage = result
End Function

Private Function ExpectedFinalTag(ByVal formName As String, ByVal buildId As String) As String
    ExpectedFinalTag = GENERIC_TAG_PREFIX & GENERIC_FINAL_ROLE_TOKEN & GENERIC_BUILD_ID_PREFIX & buildId & "]" & GENERIC_FORM_NAME_PREFIX & formName & "]"
End Function

' =============================================================================================================================================================================
' DESCRIPTION DECORATION PROCEDURES
' =============================================================================================================================================================================

Private Sub ApplyDescriptionHeaderToMainForm(ByRef formNames() As String, ByVal recordSourceName As String, ByVal descriptionText As String, ByVal buildId As String, ByVal visualColumnCount As Long)
    SetADCStatus "Adding description header to main form", formNames(FORM_INDEX_BOTH)
    AddOrReplaceDescriptionHeader formNames(FORM_INDEX_BOTH), HeaderLabelCaption(recordSourceName, descriptionText), buildId, visualColumnCount
End Sub

Private Sub AddOrReplaceDescriptionHeader(ByVal formName As String, ByVal labelCaption As String, ByVal buildId As String, ByVal visualColumnCount As Long)
    On Error GoTo ErrHandler
    Dim frm As Access.Form
    Dim ctl As Access.control
    Dim measurementCtl As Access.label
    Dim targetSection As Long
    Dim labelLeft As Long
    Dim labelWidth As Long
    Dim labelHeight As Long
    Dim requiredBandHeight As Long
    Dim existingDetailBandHeight As Long
    Dim openedHere As Boolean
    If SharedAccess_ObjectIsOpenStrict(acForm, formName) Then Err.Raise vbObjectError + 820, MODULE_NAME & ".AddOrReplaceDescriptionHeader", "The generated Both form is unexpectedly open: " & formName
    WriteADCDiagnostic "BEFORE", "Open Both form hidden for decoration", formName
    DoCmd.OpenForm formName, acDesign, , , , acHidden
    openedHere = True
    Set frm = forms(formName)
    WriteADCDiagnostic "AFTER", "Open Both form hidden for decoration", formName
    If StrComp(Nz(frm.Tag, vbNullString), ExpectedFinalTag(formName, buildId), vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 821, MODULE_NAME & ".AddOrReplaceDescriptionHeader", "The generated Both form ownership tag changed before decoration."
    existingDetailBandHeight = ExistingDescriptionDetailBandHeight(frm)
    DeleteControlIfExists frm, formName, DESCRIPTION_LABEL_NAME
    DeleteControlIfExists frm, formName, DESCRIPTION_MEASUREMENT_LABEL_NAME
    ResolveDescriptionLabelHorizontalBounds frm, labelLeft, labelWidth
    WriteADCDiagnostic "BEFORE", "Create measurement label", formName
    Set measurementCtl = CreateControl(formName, acLabel, acDetail, , , 0, 0, labelWidth, 300)
    WriteADCDiagnostic "AFTER", "Create measurement label", formName
    measurementCtl.Name = DESCRIPTION_MEASUREMENT_LABEL_NAME
    ConfigureDescriptionLabel measurementCtl, labelCaption, False
    measurementCtl.Visible = False
    labelHeight = MeasuredDescriptionLabelHeight(measurementCtl, labelCaption, labelWidth)
    DeleteControl formName, DESCRIPTION_MEASUREMENT_LABEL_NAME
    Set measurementCtl = Nothing
    requiredBandHeight = DESCRIPTION_LABEL_TOP_TWIPS + labelHeight + DESCRIPTION_LABEL_BOTTOM_PADDING_TWIPS
    WriteADCDiagnostic "BEFORE", "DescriptionTargetSection", formName
    targetSection = DescriptionTargetSection(formName, frm, requiredBandHeight)
    WriteADCDiagnostic "AFTER", "DescriptionTargetSection", "Form=" & formName & "; Section=" & CStr(targetSection)
    If targetSection = acDetail Then PrepareDetailDescriptionBand frm, formName, existingDetailBandHeight, requiredBandHeight
    WriteADCDiagnostic "BEFORE", "Create persisted description label", formName
    Set ctl = CreateControl(formName, acLabel, targetSection, , , labelLeft, DESCRIPTION_LABEL_TOP_TWIPS, labelWidth, labelHeight)
    WriteADCDiagnostic "AFTER", "Create persisted description label", formName
    ctl.Name = DESCRIPTION_LABEL_NAME
    ConfigureDescriptionLabel ctl, labelCaption, True
    frm.PopUp = True
    frm.Modal = False
    frm.AutoCenter = True
    frm.AutoResize = True
    frm.FitToScreen = True
    ctl.properties("HorizontalAnchor").value = DESCRIPTION_LABEL_HORIZONTAL_ANCHOR_BOTH
    WriteADCDiagnostic "BEFORE", "AddOrReplaceADCButtons", "Form=" & formName & "; VisualColumns=" & CStr(visualColumnCount)
    AddOrReplaceADCButtons frm, formName, visualColumnCount
    WriteADCDiagnostic "AFTER", "AddOrReplaceADCButtons", formName
    WriteADCDiagnostic "BEFORE", "StabilizeDecoratedParentBottom", formName
    StabilizeDecoratedParentBottom frm, formName, targetSection
    WriteADCDiagnostic "AFTER", "StabilizeDecoratedParentBottom", formName
    WriteADCDiagnostic "BEFORE", "StabilizeDecoratedParentWidth", formName
    StabilizeDecoratedParentWidth frm, formName, ctl, visualColumnCount
    WriteADCDiagnostic "AFTER", "StabilizeDecoratedParentWidth", formName
    labelWidth = CLng(ctl.Width)
    labelHeight = CLng(ctl.Height)
    VerifyOpenDescriptionLabel frm, labelCaption, targetSection, labelLeft, labelWidth, labelHeight
    VerifyOpenADCButtons frm, visualColumnCount
    DoCmd.Save acForm, formName
    DoCmd.Close acForm, formName, acSaveYes
    openedHere = False
    Debug.Print MODULE_NAME & ": Added generated description header and footer buttons. Form=" & formName & "; Section=" & DescriptionSectionName(targetSection) & "; Left=" & CStr(labelLeft) & "; Width=" & CStr(labelWidth) & "; Height=" & CStr(labelHeight) & "; BandHeight=" & CStr(requiredBandHeight) & "; VisualColumns=" & CStr(visualColumnCount) & "; ButtonCount=" & CStr(visualColumnCount * ADC_BUTTON_ROW_COUNT) & "; HorizontalAnchor=Both; MeasuredMultiRowHeight=True; PopUp=True; AutoResize=True; FitToScreen=True"
    Exit Sub
ErrHandler:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    On Error Resume Next
    If openedHere Then DoCmd.Close acForm, formName, acSaveNo
    On Error GoTo 0
    Err.Raise errNumber, MODULE_NAME & ".AddOrReplaceDescriptionHeader", "Unable to add generated description header and footer buttons to '" & formName & "'. " & errDescription
End Sub

Private Sub ConfigureDescriptionLabel(ByVal ctl As Access.control, ByVal labelCaption As String, ByVal includeControlTip As Boolean)
    ctl.Caption = labelCaption
    ctl.fontSize = DESCRIPTION_FONT_SIZE
    ctl.BackStyle = 0
    ctl.SpecialEffect = 0
    TrySetControlProperty ctl, "AutoSize", False
    TrySetControlProperty ctl, "LeftMargin", DESCRIPTION_LABEL_INNER_HORIZONTAL_PADDING_TWIPS
    TrySetControlProperty ctl, "RightMargin", DESCRIPTION_LABEL_INNER_HORIZONTAL_PADDING_TWIPS
    TrySetControlProperty ctl, "TopMargin", DESCRIPTION_LABEL_INNER_VERTICAL_PADDING_TWIPS
    TrySetControlProperty ctl, "BottomMargin", DESCRIPTION_LABEL_INNER_VERTICAL_PADDING_TWIPS
    If includeControlTip Then TrySetControlProperty ctl, "ControlTipText", Left$(labelCaption, MAX_CONTROL_TIP_LEN)
End Sub

Private Function DescriptionTargetSection(ByVal formName As String, ByVal frm As Access.Form, ByVal requiredBandHeight As Long) As Long
    If TryPrepareHeaderSection(frm, requiredBandHeight) Then
        DescriptionTargetSection = acHeader
        Exit Function
    End If
    On Error Resume Next
    DoCmd.SelectObject acForm, formName, False
    DoCmd.RunCommand acCmdFormHdrFtr
    If Err.Number <> 0 Then
        Debug.Print MODULE_NAME & ": Unable to create a form header for " & formName & ". Error " & CStr(Err.Number) & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    If TryPrepareHeaderSection(frm, requiredBandHeight) Then
        DescriptionTargetSection = acHeader
        Exit Function
    End If
    Debug.Print MODULE_NAME & ": Form header section was unavailable for " & formName & "; using a reserved detail-section description band. RequiredBandHeight=" & CStr(requiredBandHeight)
    DescriptionTargetSection = acDetail
End Function

Private Function TryPrepareHeaderSection(ByVal frm As Access.Form, ByVal requiredHeight As Long) As Boolean
    On Error GoTo ExitHere
    With frm.Section(acHeader)
        .Visible = True
        .Height = requiredHeight
    End With
    TryPrepareHeaderSection = True
ExitHere:
    Err.Clear
End Function

Private Function ExistingDescriptionDetailBandHeight(ByVal frm As Access.Form) As Long
    On Error GoTo ExitHere
    Dim ctl As Access.control
    Set ctl = frm.Controls(DESCRIPTION_LABEL_NAME)
    If CLng(ctl.Section) = acDetail Then ExistingDescriptionDetailBandHeight = CLng(ctl.Top) + CLng(ctl.Height) + DESCRIPTION_LABEL_BOTTOM_PADDING_TWIPS
ExitHere:
    Set ctl = Nothing
    Err.Clear
End Function

Private Sub PrepareDetailDescriptionBand(ByVal frm As Access.Form, ByVal formName As String, ByVal existingBandHeight As Long, ByVal requiredBandHeight As Long)
    On Error GoTo ErrHandler
    Dim detailHeight As Long
    Dim newDetailHeight As Long
    Dim shiftAmount As Long
    Dim movedControlCount As Long
    detailHeight = CLng(frm.Section(acDetail).Height)
    shiftAmount = requiredBandHeight - existingBandHeight
    newDetailHeight = detailHeight + shiftAmount
    If newDetailHeight < requiredBandHeight Then Err.Raise vbObjectError + 822, MODULE_NAME & ".PrepareDetailDescriptionBand", "The detail section on '" & formName & "' cannot be reduced to the required description-band height. RequiredBandHeight=" & CStr(requiredBandHeight) & "; CalculatedDetailHeight=" & CStr(newDetailHeight) & "."
    If newDetailHeight > MAX_ACCESS_SECTION_HEIGHT_TWIPS Then Err.Raise vbObjectError + 839, MODULE_NAME & ".PrepareDetailDescriptionBand", "The detail section on '" & formName & "' cannot be expanded by " & CStr(shiftAmount) & " twips without exceeding the Access section-height limit."
    If shiftAmount > 0 Then
        frm.Section(acDetail).Height = newDetailHeight
        movedControlCount = ShiftDetailControlsVertically(frm, shiftAmount)
    ElseIf shiftAmount < 0 Then
        movedControlCount = ShiftDetailControlsVertically(frm, shiftAmount)
        frm.Section(acDetail).Height = newDetailHeight
    End If
    Debug.Print MODULE_NAME & ": Prepared detail description band. Form=" & formName & "; ExistingBandHeight=" & CStr(existingBandHeight) & "; RequiredBandHeight=" & CStr(requiredBandHeight) & "; ShiftAmount=" & CStr(shiftAmount) & "; MovedControls=" & CStr(movedControlCount) & "; DetailHeightBefore=" & CStr(detailHeight) & "; DetailHeightAfter=" & CStr(newDetailHeight)
    Exit Sub
ErrHandler:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Raise errNumber, MODULE_NAME & ".PrepareDetailDescriptionBand", "Unable to reserve detail-section description space on '" & formName & "'. " & errDescription
End Sub

Private Function ShiftDetailControlsVertically(ByVal frm As Access.Form, ByVal offsetTwips As Long) As Long
    On Error GoTo ErrHandler
    Dim ctl As Access.control
    Dim movedCount As Long
    For Each ctl In frm.Controls
        If CLng(ctl.Section) = acDetail Then
            ctl.Top = CLng(ctl.Top) + offsetTwips
            movedCount = movedCount + 1
        End If
    Next ctl
    ShiftDetailControlsVertically = movedCount
    Exit Function
ErrHandler:
    Dim errNumber As Long
    Dim errDescription As String
    Dim controlName As String
    errNumber = Err.Number
    errDescription = Err.Description
    controlName = SafeControlName(ctl)
    Err.Raise errNumber, MODULE_NAME & ".ShiftDetailControlsVertically", "Unable to move detail control '" & controlName & "' vertically by " & CStr(offsetTwips) & " twips. " & errDescription
End Function

Private Sub ResolveDescriptionLabelHorizontalBounds(ByVal frm As Access.Form, ByRef labelLeft As Long, ByRef labelWidth As Long)
    On Error GoTo ErrHandler
    Dim singleLeft As Long
    Dim singleRight As Long
    Dim dsLeft As Long
    Dim dsRight As Long
    Dim spanLeft As Long
    Dim spanRight As Long
    Dim hasSingleSpan As Boolean
    Dim hasDSSpan As Boolean
    hasSingleSpan = TryGetControlHorizontalSpan(frm, SINGLE_SUBFORM_CONTROL_NAME, singleLeft, singleRight)
    hasDSSpan = TryGetControlHorizontalSpan(frm, DS_SUBFORM_CONTROL_NAME, dsLeft, dsRight)
    If hasSingleSpan Or hasDSSpan Then
        If hasSingleSpan Then
            spanLeft = singleLeft
            spanRight = singleRight
        Else
            spanLeft = dsLeft
            spanRight = dsRight
        End If
        If hasDSSpan Then
            If dsLeft < spanLeft Then spanLeft = dsLeft
            If dsRight > spanRight Then spanRight = dsRight
        End If
    Else
        spanLeft = DESCRIPTION_LABEL_LEFT_TWIPS
        spanRight = ExistingHorizontalContentRightEdge(frm) - DESCRIPTION_LABEL_RIGHT_PADDING_TWIPS
    End If
    If spanLeft < 0 Then spanLeft = 0
    If spanRight > CLng(frm.Width) Then frm.Width = spanRight
    If spanRight <= spanLeft Then Err.Raise vbObjectError + 823, MODULE_NAME & ".ResolveDescriptionLabelHorizontalBounds", "The form does not have usable horizontal space for the generated description label. Left=" & CStr(spanLeft) & "; Right=" & CStr(spanRight)
    labelLeft = spanLeft
    labelWidth = spanRight - spanLeft
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, MODULE_NAME & ".ResolveDescriptionLabelHorizontalBounds", "Unable to calculate generated description-label bounds. " & Err.Description
End Sub

Private Function GeneratedVisualColumnCount(ByVal formName As String) As Long
    On Error GoTo ErrHandler
    Dim frm As Access.Form
    Dim ctl As Access.control
    Dim layoutIds As Collection
    Dim fallbackLefts As Collection
    Dim openedHere As Boolean
    Dim layoutId As Long
    If SharedAccess_ObjectIsOpenStrict(acForm, formName) Then Err.Raise vbObjectError + 861, MODULE_NAME & ".GeneratedVisualColumnCount", "The generated Single form must be closed before its visual columns are counted: " & formName
    Set layoutIds = New Collection
    Set fallbackLefts = New Collection
    DoCmd.OpenForm formName, acDesign, , , , acHidden
    openedHere = True
    Set frm = forms(formName)
    For Each ctl In frm.Controls
        If CLng(ctl.Section) = acDetail Then
            If IsGeneratedBoundFieldControl(ctl) Then
                If TryReadControlLongProperty(ctl, "LayoutID", layoutId) Then
                    If layoutId > 0 Then AddUniqueLongCollectionValue layoutIds, layoutId
                End If
                AddUniqueLongCollectionValue fallbackLefts, CLng(ctl.Left)
            End If
        End If
    Next ctl
    DoCmd.Close acForm, formName, acSaveNo
    openedHere = False
    Set frm = Nothing
    If layoutIds.count > 0 Then
        GeneratedVisualColumnCount = layoutIds.count
    Else
        GeneratedVisualColumnCount = fallbackLefts.count
    End If
    If GeneratedVisualColumnCount <= 0 Then Err.Raise vbObjectError + 862, MODULE_NAME & ".GeneratedVisualColumnCount", "No generated bound-control visual columns were found on form '" & formName & "'."
    Debug.Print MODULE_NAME & ": Counted generated Single-form visual columns. Form=" & formName & "; NativeLayoutColumns=" & CStr(layoutIds.count) & "; FallbackLeftColumns=" & CStr(fallbackLefts.count) & "; SelectedColumnCount=" & CStr(GeneratedVisualColumnCount)
    Exit Function
ErrHandler:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    On Error Resume Next
    If openedHere Then DoCmd.Close acForm, formName, acSaveNo
    On Error GoTo 0
    Err.Raise errNumber, MODULE_NAME & ".GeneratedVisualColumnCount", "Unable to count the generated visual columns on Single form '" & formName & "'. " & errDescription
End Function

Private Function IsGeneratedBoundFieldControl(ByVal ctl As Access.control) As Boolean
    On Error GoTo ExitHere
    Dim controlSourceText As String
    controlSourceText = Trim$(Nz(ctl.controlSource, vbNullString))
    If Len(controlSourceText) = 0 Then Exit Function
    If Left$(controlSourceText, 1) = "=" Then Exit Function
    IsGeneratedBoundFieldControl = True
ExitHere:
    Err.Clear
End Function

Private Function TryReadControlLongProperty(ByVal ctl As Access.control, ByVal propertyName As String, ByRef propertyValue As Long) As Boolean
    On Error GoTo ExitHere
    propertyValue = CLng(ctl.properties(propertyName).value)
    TryReadControlLongProperty = True
ExitHere:
    Err.Clear
End Function

Private Sub AddUniqueLongCollectionValue(ByVal values As Collection, ByVal value As Long)
    On Error Resume Next
    values.Add value, CStr(value)
    Err.Clear
    On Error GoTo 0
End Sub

' =============================================================================================================================================================================
' ADC FOOTER-BUTTON CONSTRUCTION PROCEDURES
' =============================================================================================================================================================================

Private Sub AddOrReplaceADCButtons(ByVal frm As Access.Form, ByVal formName As String, ByVal visualColumnCount As Long)
    On Error GoTo ErrHandler
    Dim ctl As Access.control
    Dim buttonNumber As Long
    Dim totalButtonCount As Long
    Dim buttonLeft As Long
    Dim buttonTop As Long
    Dim buttonWidth As Long
    Dim footerHeight As Long
    If visualColumnCount <= 0 Then Err.Raise vbObjectError + 863, MODULE_NAME & ".AddOrReplaceADCButtons", "Visual column count must be greater than zero."
    DeleteGeneratedADCButtons frm, formName
    footerHeight = ExpectedADCButtonFooterHeight()
    On Error GoTo FooterUnavailable
    frm.Section(acFooter).Visible = True
    frm.Section(acFooter).Height = footerHeight
    On Error GoTo ErrHandler
    totalButtonCount = visualColumnCount * ADC_BUTTON_ROW_COUNT
    For buttonNumber = 1 To totalButtonCount
        CalculateADCButtonBounds frm, visualColumnCount, buttonNumber, buttonLeft, buttonTop, buttonWidth
        Set ctl = CreateControl(formName, acCommandButton, acFooter, , , buttonLeft, buttonTop, buttonWidth, ADC_BUTTON_HEIGHT_TWIPS)
        ctl.Name = ADCButtonName(buttonNumber)
        ctl.Caption = CStr(buttonNumber)
        ctl.OnClick = vbNullString
        ctl.fontSize = ADC_BUTTON_FONT_SIZE
        ctl.TabStop = False
        ctl.Default = False
        ctl.Cancel = False
        Set ctl = Nothing
    Next buttonNumber
    frm.Section(acFooter).Height = footerHeight
    Debug.Print MODULE_NAME & ": Created inert Access Development Console footer buttons. Form=" & formName & "; VisualColumns=" & CStr(visualColumnCount) & "; Rows=" & CStr(ADC_BUTTON_ROW_COUNT) & "; ButtonCount=" & CStr(totalButtonCount) & "; FirstButton=ADC1Btn; LastButton=" & ADCButtonName(totalButtonCount) & "; FooterHeight=" & CStr(footerHeight)
    Exit Sub
FooterUnavailable:
    Dim footerErrNumber As Long
    Dim footerErrDescription As String
    footerErrNumber = Err.Number
    footerErrDescription = Err.Description
    On Error GoTo ErrHandler
    Err.Raise vbObjectError + 864, MODULE_NAME & ".AddOrReplaceADCButtons", "The generated Both form does not provide a usable Form Footer for the required two rows of Access Development Console buttons. Original error " & CStr(footerErrNumber) & ": " & footerErrDescription
ErrHandler:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Raise errNumber, MODULE_NAME & ".AddOrReplaceADCButtons", "Unable to create the generated Access Development Console footer buttons on '" & formName & "'. " & errDescription
End Sub

Private Sub DeleteGeneratedADCButtons(ByVal frm As Access.Form, ByVal formName As String)
    On Error GoTo ErrHandler
    Dim ctl As Access.control
    Dim controlIndex As Long
    Dim controlName As String
    For controlIndex = frm.Controls.count - 1 To 0 Step -1
        Set ctl = frm.Controls(controlIndex)
        controlName = ctl.Name
        If IsGeneratedADCButtonName(controlName) Then
            Set ctl = Nothing
            DeleteControl formName, controlName
            Debug.Print MODULE_NAME & ": Deleted existing generated Access Development Console button " & controlName & " from " & formName
        Else
            Set ctl = Nothing
        End If
    Next controlIndex
    Exit Sub
ErrHandler:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Raise errNumber, MODULE_NAME & ".DeleteGeneratedADCButtons", "Unable to remove prior generated Access Development Console buttons from '" & formName & "'. LastControl='" & controlName & "'. " & errDescription
End Sub

Private Function ADCButtonName(ByVal buttonNumber As Long) As String
    ADCButtonName = ADC_BUTTON_NAME_PREFIX & CStr(buttonNumber) & ADC_BUTTON_NAME_SUFFIX
End Function

Private Function IsGeneratedADCButtonName(ByVal controlName As String) As Boolean
    Dim numericText As String
    Dim characterIndex As Long
    If Len(controlName) <= Len(ADC_BUTTON_NAME_PREFIX) + Len(ADC_BUTTON_NAME_SUFFIX) Then Exit Function
    If StrComp(Left$(controlName, Len(ADC_BUTTON_NAME_PREFIX)), ADC_BUTTON_NAME_PREFIX, vbBinaryCompare) <> 0 Then Exit Function
    If StrComp(Right$(controlName, Len(ADC_BUTTON_NAME_SUFFIX)), ADC_BUTTON_NAME_SUFFIX, vbBinaryCompare) <> 0 Then Exit Function
    numericText = Mid$(controlName, Len(ADC_BUTTON_NAME_PREFIX) + 1, Len(controlName) - Len(ADC_BUTTON_NAME_PREFIX) - Len(ADC_BUTTON_NAME_SUFFIX))
    If Len(numericText) = 0 Or Len(numericText) > 9 Then Exit Function
    For characterIndex = 1 To Len(numericText)
        If Mid$(numericText, characterIndex, 1) < "0" Or Mid$(numericText, characterIndex, 1) > "9" Then Exit Function
    Next characterIndex
    If Val(numericText) <= 0 Then Exit Function
    IsGeneratedADCButtonName = True
End Function

Private Sub CalculateADCButtonBounds(ByVal frm As Access.Form, ByVal visualColumnCount As Long, ByVal buttonNumber As Long, ByRef buttonLeft As Long, ByRef buttonTop As Long, ByRef buttonWidth As Long)
    On Error GoTo ErrHandler
    Dim spanLeft As Long
    Dim spanRight As Long
    Dim spanWidth As Long
    Dim usableWidth As Long
    Dim baseWidth As Long
    Dim remainderWidth As Long
    Dim rowIndex As Long
    Dim ColumnIndex As Long
    Dim extraBefore As Long
    ADCButtonHorizontalSpan frm, spanLeft, spanRight
    spanWidth = spanRight - spanLeft
    usableWidth = spanWidth - ((visualColumnCount - 1) * ADC_BUTTON_HORIZONTAL_GAP_TWIPS)
    If usableWidth < visualColumnCount Then Err.Raise vbObjectError + 865, MODULE_NAME & ".CalculateADCButtonBounds", "The parent form span is too narrow for " & CStr(visualColumnCount) & " generated button columns. SpanWidth=" & CStr(spanWidth) & "; HorizontalGap=" & CStr(ADC_BUTTON_HORIZONTAL_GAP_TWIPS) & "."
    rowIndex = (buttonNumber - 1) \ visualColumnCount
    ColumnIndex = (buttonNumber - 1) Mod visualColumnCount
    If rowIndex < 0 Or rowIndex >= ADC_BUTTON_ROW_COUNT Then Err.Raise vbObjectError + 866, MODULE_NAME & ".CalculateADCButtonBounds", "Button number is outside the required two-row range. ButtonNumber=" & CStr(buttonNumber) & "; VisualColumnCount=" & CStr(visualColumnCount) & "."
    baseWidth = usableWidth \ visualColumnCount
    remainderWidth = usableWidth Mod visualColumnCount
    If ColumnIndex < remainderWidth Then
        buttonWidth = baseWidth + 1
        extraBefore = ColumnIndex
    Else
        buttonWidth = baseWidth
        extraBefore = remainderWidth
    End If
    buttonLeft = spanLeft + (ColumnIndex * (baseWidth + ADC_BUTTON_HORIZONTAL_GAP_TWIPS)) + extraBefore
    buttonTop = ADC_BUTTON_TOP_MARGIN_TWIPS + (rowIndex * (ADC_BUTTON_HEIGHT_TWIPS + ADC_BUTTON_VERTICAL_GAP_TWIPS))
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, MODULE_NAME & ".CalculateADCButtonBounds", "Unable to calculate geometry for generated Access Development Console button " & CStr(buttonNumber) & ". " & Err.Description
End Sub

Private Sub ADCButtonHorizontalSpan(ByVal frm As Access.Form, ByRef spanLeft As Long, ByRef spanRight As Long)
    On Error GoTo ErrHandler
    Dim singleCtl As Access.control
    Dim dsCtl As Access.control
    Dim singleRight As Long
    Dim dsRight As Long
    Set singleCtl = frm.Controls(SINGLE_SUBFORM_CONTROL_NAME)
    Set dsCtl = frm.Controls(DS_SUBFORM_CONTROL_NAME)
    singleRight = CLng(singleCtl.Left) + CLng(singleCtl.Width)
    dsRight = CLng(dsCtl.Left) + CLng(dsCtl.Width)
    If Abs(CLng(singleCtl.Left) - CLng(dsCtl.Left)) > DECORATED_PARENT_RIGHT_EDGE_TOLERANCE_TWIPS Or Abs(singleRight - dsRight) > DECORATED_PARENT_RIGHT_EDGE_TOLERANCE_TWIPS Then Err.Raise vbObjectError + 867, MODULE_NAME & ".ADCButtonHorizontalSpan", "SingleForm and DSForm do not provide one shared horizontal span for generated buttons. SingleLeft=" & CStr(singleCtl.Left) & "; SingleRight=" & CStr(singleRight) & "; DSLeft=" & CStr(dsCtl.Left) & "; DSRight=" & CStr(dsRight) & "."
    spanLeft = CLng(singleCtl.Left)
    spanRight = singleRight
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, MODULE_NAME & ".ADCButtonHorizontalSpan", "Unable to resolve the generated Access Development Console button span. " & Err.Description
End Sub

Private Function ExpectedADCButtonFooterHeight() As Long
    ExpectedADCButtonFooterHeight = ADC_BUTTON_TOP_MARGIN_TWIPS + (ADC_BUTTON_ROW_COUNT * ADC_BUTTON_HEIGHT_TWIPS) + ((ADC_BUTTON_ROW_COUNT - 1) * ADC_BUTTON_VERTICAL_GAP_TWIPS) + ADC_BUTTON_BOTTOM_MARGIN_TWIPS
End Function

Private Sub VerifyOpenADCButtons(ByVal frm As Access.Form, ByVal visualColumnCount As Long)
    Dim problems As String
    problems = ADCButtonProblems(frm, visualColumnCount)
    If Len(problems) > 0 Then Err.Raise vbObjectError + 868, MODULE_NAME & ".VerifyOpenADCButtons", "Generated Access Development Console button verification failed before save:" & vbCrLf & problems
End Sub

Private Function ADCButtonProblems(ByVal frm As Access.Form, ByVal visualColumnCount As Long) As String
    On Error GoTo ErrHandler
    Dim ctl As Access.control
    Dim buttonNumber As Long
    Dim totalButtonCount As Long
    Dim expectedLeft As Long
    Dim expectedTop As Long
    Dim expectedWidth As Long
    Dim expectedFooterHeight As Long
    Dim actualGeneratedCount As Long
    Dim problems As String
    If visualColumnCount <= 0 Then
        ADCButtonProblems = "  - Expected visual column count is not greater than zero." & vbCrLf
        Exit Function
    End If
    totalButtonCount = visualColumnCount * ADC_BUTTON_ROW_COUNT
    expectedFooterHeight = ExpectedADCButtonFooterHeight()
    actualGeneratedCount = CountGeneratedADCButtons(frm)
    If actualGeneratedCount <> totalButtonCount Then problems = problems & "  - Generated Access Development Console button count is incorrect. Expected=" & CStr(totalButtonCount) & "; Actual=" & CStr(actualGeneratedCount) & "; VisualColumns=" & CStr(visualColumnCount) & "; Rows=" & CStr(ADC_BUTTON_ROW_COUNT) & "." & vbCrLf
    For buttonNumber = 1 To totalButtonCount
        If Not ControlExistsOnFormInstance(frm, ADCButtonName(buttonNumber)) Then
            problems = problems & "  - Required generated button is missing: " & ADCButtonName(buttonNumber) & "." & vbCrLf
        Else
            Set ctl = frm.Controls(ADCButtonName(buttonNumber))
            CalculateADCButtonBounds frm, visualColumnCount, buttonNumber, expectedLeft, expectedTop, expectedWidth
            If CLng(ctl.controlType) <> acCommandButton Then problems = problems & "  - " & ctl.Name & " is not a command button. ActualControlType=" & CStr(ctl.controlType) & "." & vbCrLf
            If CLng(ctl.Section) <> acFooter Then problems = problems & "  - " & ctl.Name & " is not in the Form Footer. ActualSection=" & CStr(ctl.Section) & "." & vbCrLf
            If StrComp(CStr(ctl.Caption), CStr(buttonNumber), vbBinaryCompare) <> 0 Then problems = problems & "  - " & ctl.Name & " has the wrong caption. Expected='" & CStr(buttonNumber) & "'; Actual='" & CStr(ctl.Caption) & "'." & vbCrLf
            If Len(Trim$(Nz(ctl.OnClick, vbNullString))) > 0 Then problems = problems & "  - " & ctl.Name & " must remain inert, but OnClick is not blank. ActualOnClick='" & Nz(ctl.OnClick, vbNullString) & "'." & vbCrLf
            If Abs(CLng(ctl.Left) - expectedLeft) > 2 Or Abs(CLng(ctl.Top) - expectedTop) > 2 Or Abs(CLng(ctl.Width) - expectedWidth) > 2 Or Abs(CLng(ctl.Height) - ADC_BUTTON_HEIGHT_TWIPS) > 2 Then problems = problems & "  - " & ctl.Name & " geometry is incorrect. ExpectedLeft=" & CStr(expectedLeft) & "; ActualLeft=" & CStr(ctl.Left) & "; ExpectedTop=" & CStr(expectedTop) & "; ActualTop=" & CStr(ctl.Top) & "; ExpectedWidth=" & CStr(expectedWidth) & "; ActualWidth=" & CStr(ctl.Width) & "; ExpectedHeight=" & CStr(ADC_BUTTON_HEIGHT_TWIPS) & "; ActualHeight=" & CStr(ctl.Height) & "." & vbCrLf
            Set ctl = Nothing
        End If
    Next buttonNumber
    If Abs(CLng(frm.Section(acFooter).Height) - expectedFooterHeight) > DECORATED_PARENT_BOTTOM_EDGE_TOLERANCE_TWIPS Then problems = problems & "  - Form Footer height does not match the exact two-row button band. ExpectedFooterHeight=" & CStr(expectedFooterHeight) & "; ActualFooterHeight=" & CStr(frm.Section(acFooter).Height) & "; Tolerance=" & CStr(DECORATED_PARENT_BOTTOM_EDGE_TOLERANCE_TWIPS) & "." & vbCrLf
    ADCButtonProblems = problems
    Exit Function
ErrHandler:
    Err.Raise Err.Number, MODULE_NAME & ".ADCButtonProblems", "Unable to verify generated Access Development Console footer buttons. " & Err.Description
End Function

Private Function CountGeneratedADCButtons(ByVal frm As Access.Form) As Long
    Dim ctl As Access.control
    For Each ctl In frm.Controls
        If IsGeneratedADCButtonName(ctl.Name) Then CountGeneratedADCButtons = CountGeneratedADCButtons + 1
    Next ctl
End Function

Private Function RequiredGeneratedFooterHeight(ByVal frm As Access.Form, ByRef unexpectedFooterControlCount As Long) As Long
    On Error GoTo ErrHandler
    Dim ctl As Access.control
    Dim controlBottom As Long
    For Each ctl In frm.Controls
        If CLng(ctl.Section) = acFooter Then
            If IsGeneratedADCButtonName(ctl.Name) Then
                controlBottom = CLng(ctl.Top) + CLng(ctl.Height)
                If controlBottom > RequiredGeneratedFooterHeight Then RequiredGeneratedFooterHeight = controlBottom
            Else
                unexpectedFooterControlCount = unexpectedFooterControlCount + 1
            End If
        End If
    Next ctl
    If RequiredGeneratedFooterHeight > 0 Then RequiredGeneratedFooterHeight = RequiredGeneratedFooterHeight + ADC_BUTTON_BOTTOM_MARGIN_TWIPS
    Exit Function
ErrHandler:
    Err.Raise Err.Number, MODULE_NAME & ".RequiredGeneratedFooterHeight", "Unable to calculate the generated footer-button height. " & Err.Description
End Function

' =============================================================================================================================================================================
' GEOMETRY / MEASUREMENT / STABILIZATION PROCEDURES
' =============================================================================================================================================================================

Private Sub StabilizeDecoratedParentBottom(ByVal frm As Access.Form, ByVal formName As String, ByVal descriptionSection As Long)
    On Error GoTo ErrHandler
    Dim ctl As Access.control
    Dim singleCtl As Access.control
    Dim dsCtl As Access.control
    Dim expectedSingleTop As Long
    Dim expectedSingleHeight As Long
    Dim expectedDSTop As Long
    Dim expectedDSHeight As Long
    Dim requiredDetailHeight As Long
    Dim requiredFooterHeight As Long
    Dim detailHeightBefore As Long
    Dim detailHeightAfter As Long
    Dim footerHeightBefore As Long
    Dim footerHeightAfter As Long
    Dim unexpectedFooterControlCount As Long
    Dim controlBottom As Long
    Dim passNumber As Long
    Set singleCtl = frm.Controls(SINGLE_SUBFORM_CONTROL_NAME)
    Set dsCtl = frm.Controls(DS_SUBFORM_CONTROL_NAME)
    expectedSingleTop = CLng(singleCtl.Top)
    expectedSingleHeight = CLng(singleCtl.Height)
    expectedDSTop = CLng(dsCtl.Top)
    expectedDSHeight = CLng(dsCtl.Height)
    For Each ctl In frm.Controls
        If CLng(ctl.Section) = acDetail Then
            controlBottom = CLng(ctl.Top) + CLng(ctl.Height)
            If controlBottom > requiredDetailHeight Then requiredDetailHeight = controlBottom
        End If
    Next ctl
    If requiredDetailHeight <= 0 Then Err.Raise vbObjectError + 869, MODULE_NAME & ".StabilizeDecoratedParentBottom", "The decorated parent Detail section contains no measurable controls."
    If descriptionSection = acHeader Then
        requiredFooterHeight = RequiredGeneratedFooterHeight(frm, unexpectedFooterControlCount)
        If unexpectedFooterControlCount > 0 Then Err.Raise vbObjectError + 870, MODULE_NAME & ".StabilizeDecoratedParentBottom", "The generated Both form contains " & CStr(unexpectedFooterControlCount) & " unexpected Form Footer control(s), so the two-row generated button band cannot be compacted safely."
        If requiredFooterHeight <= 0 Then Err.Raise vbObjectError + 871, MODULE_NAME & ".StabilizeDecoratedParentBottom", "The generated Both form does not contain the required Access Development Console footer buttons."
    End If
    detailHeightBefore = CLng(frm.Section(acDetail).Height)
    If descriptionSection = acHeader Then footerHeightBefore = CLng(frm.Section(acFooter).Height)
    For passNumber = 1 To DECORATED_PARENT_BOTTOM_STABILIZATION_PASSES
        WriteADCDiagnostic "PASS", "StabilizeDecoratedParentBottom", "Form=" & formName & "; Pass=" & CStr(passNumber)
        If Abs(CLng(frm.Section(acDetail).Height) - requiredDetailHeight) > DECORATED_PARENT_BOTTOM_EDGE_TOLERANCE_TWIPS Then frm.Section(acDetail).Height = requiredDetailHeight
        If descriptionSection = acHeader Then
            frm.Section(acFooter).Visible = True
            If Abs(CLng(frm.Section(acFooter).Height) - requiredFooterHeight) > DECORATED_PARENT_BOTTOM_EDGE_TOLERANCE_TWIPS Then frm.Section(acFooter).Height = requiredFooterHeight
        End If
        DoCmd.Save acForm, formName
        detailHeightAfter = CLng(frm.Section(acDetail).Height)
        If descriptionSection = acHeader Then footerHeightAfter = CLng(frm.Section(acFooter).Height)
        If Abs(CLng(singleCtl.Top) - expectedSingleTop) > DECORATED_PARENT_BOTTOM_EDGE_TOLERANCE_TWIPS Or Abs(CLng(singleCtl.Height) - expectedSingleHeight) > DECORATED_PARENT_BOTTOM_EDGE_TOLERANCE_TWIPS Then Err.Raise vbObjectError + 872, MODULE_NAME & ".StabilizeDecoratedParentBottom", "SingleForm moved or resized while the decorated parent bottom was compacted. ExpectedTop=" & CStr(expectedSingleTop) & "; ActualTop=" & CStr(singleCtl.Top) & "; ExpectedHeight=" & CStr(expectedSingleHeight) & "; ActualHeight=" & CStr(singleCtl.Height) & "."
        If Abs(CLng(dsCtl.Top) - expectedDSTop) > DECORATED_PARENT_BOTTOM_EDGE_TOLERANCE_TWIPS Or Abs(CLng(dsCtl.Height) - expectedDSHeight) > DECORATED_PARENT_BOTTOM_EDGE_TOLERANCE_TWIPS Then Err.Raise vbObjectError + 873, MODULE_NAME & ".StabilizeDecoratedParentBottom", "DSForm moved or resized while the decorated parent bottom was compacted. ExpectedTop=" & CStr(expectedDSTop) & "; ActualTop=" & CStr(dsCtl.Top) & "; ExpectedHeight=" & CStr(expectedDSHeight) & "; ActualHeight=" & CStr(dsCtl.Height) & "."
        If Abs(detailHeightAfter - requiredDetailHeight) <= DECORATED_PARENT_BOTTOM_EDGE_TOLERANCE_TWIPS And (descriptionSection <> acHeader Or Abs(footerHeightAfter - requiredFooterHeight) <= DECORATED_PARENT_BOTTOM_EDGE_TOLERANCE_TWIPS) Then
            Debug.Print MODULE_NAME & ": Decorated parent bottom stabilized. Form=" & formName & "; DetailHeightBefore=" & CStr(detailHeightBefore) & "; RequiredDetailHeight=" & CStr(requiredDetailHeight) & "; PersistedDetailHeight=" & CStr(detailHeightAfter) & "; FooterHeightBefore=" & CStr(footerHeightBefore) & "; RequiredFooterHeight=" & CStr(requiredFooterHeight) & "; PersistedFooterHeight=" & CStr(footerHeightAfter) & "; Pass=" & CStr(passNumber)
            Exit Sub
        End If
        Debug.Print MODULE_NAME & ": Decorated parent bottom stabilization retry required. Form=" & formName & "; RequiredDetailHeight=" & CStr(requiredDetailHeight) & "; PersistedDetailHeight=" & CStr(detailHeightAfter) & "; RequiredFooterHeight=" & CStr(requiredFooterHeight) & "; PersistedFooterHeight=" & CStr(footerHeightAfter) & "; Pass=" & CStr(passNumber)
    Next passNumber
    Err.Raise vbObjectError + 874, MODULE_NAME & ".StabilizeDecoratedParentBottom", "Access retained incorrect bottom geometry after the bounded stabilization passes. Form=" & formName & "; DetailHeightBefore=" & CStr(detailHeightBefore) & "; RequiredDetailHeight=" & CStr(requiredDetailHeight) & "; PersistedDetailHeight=" & CStr(detailHeightAfter) & "; FooterHeightBefore=" & CStr(footerHeightBefore) & "; RequiredFooterHeight=" & CStr(requiredFooterHeight) & "; PersistedFooterHeight=" & CStr(footerHeightAfter) & "; Passes=" & CStr(DECORATED_PARENT_BOTTOM_STABILIZATION_PASSES) & "."
    Exit Sub
ErrHandler:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Raise errNumber, MODULE_NAME & ".StabilizeDecoratedParentBottom", "Unable to compact and stabilize the decorated parent bottom for '" & formName & "'. " & errDescription
End Sub

Private Sub StabilizeDecoratedParentWidth(ByRef frm As Access.Form, ByVal formName As String, ByRef descriptionLabel As Access.control, ByVal visualColumnCount As Long)
    On Error GoTo ErrHandler
    Dim singleCtl As Access.control
    Dim dsCtl As Access.control
    Dim expectedSingleLeft As Long
    Dim expectedSingleHeight As Long
    Dim expectedDSLeft As Long
    Dim expectedDSHeight As Long
    Dim singleRight As Long
    Dim dsRight As Long
    Dim labelRight As Long
    Dim commonRight As Long
    Dim parentWidth As Long
    Dim widthBefore As Long
    Dim parentExcess As Long
    Dim actionText As String
    Set singleCtl = frm.Controls(SINGLE_SUBFORM_CONTROL_NAME)
    Set dsCtl = frm.Controls(DS_SUBFORM_CONTROL_NAME)
    expectedSingleLeft = CLng(singleCtl.Left)
    expectedSingleHeight = CLng(singleCtl.Height)
    expectedDSLeft = CLng(dsCtl.Left)
    expectedDSHeight = CLng(dsCtl.Height)
    widthBefore = CLng(frm.Width)
    Set descriptionLabel = frm.Controls(DESCRIPTION_LABEL_NAME)
    If Abs(CLng(singleCtl.Left) - expectedSingleLeft) > DECORATED_PARENT_RIGHT_EDGE_TOLERANCE_TWIPS Or Abs(CLng(singleCtl.Height) - expectedSingleHeight) > DECORATED_PARENT_RIGHT_EDGE_TOLERANCE_TWIPS Then Err.Raise vbObjectError + 852, MODULE_NAME & ".StabilizeDecoratedParentWidth", "SingleForm moved horizontally or changed height before the one-cycle width persistence check. ExpectedLeft=" & CStr(expectedSingleLeft) & "; ActualLeft=" & CStr(singleCtl.Left) & "; ExpectedHeight=" & CStr(expectedSingleHeight) & "; ActualHeight=" & CStr(singleCtl.Height) & "."
    If Abs(CLng(dsCtl.Left) - expectedDSLeft) > DECORATED_PARENT_RIGHT_EDGE_TOLERANCE_TWIPS Or Abs(CLng(dsCtl.Height) - expectedDSHeight) > DECORATED_PARENT_RIGHT_EDGE_TOLERANCE_TWIPS Then Err.Raise vbObjectError + 853, MODULE_NAME & ".StabilizeDecoratedParentWidth", "DSForm moved horizontally or changed height before the one-cycle width persistence check. ExpectedLeft=" & CStr(expectedDSLeft) & "; ActualLeft=" & CStr(dsCtl.Left) & "; ExpectedHeight=" & CStr(expectedDSHeight) & "; ActualHeight=" & CStr(dsCtl.Height) & "."
    singleRight = CLng(singleCtl.Left) + CLng(singleCtl.Width)
    dsRight = CLng(dsCtl.Left) + CLng(dsCtl.Width)
    labelRight = CLng(descriptionLabel.Left) + CLng(descriptionLabel.Width)
    If Abs(singleRight - dsRight) > DECORATED_PARENT_RIGHT_EDGE_TOLERANCE_TWIPS Or Abs(singleRight - labelRight) > DECORATED_PARENT_RIGHT_EDGE_TOLERANCE_TWIPS Then Err.Raise vbObjectError + 851, MODULE_NAME & ".StabilizeDecoratedParentWidth", "The decorated controls do not share one right edge before the one-cycle width persistence check. SingleRight=" & CStr(singleRight) & "; DSRight=" & CStr(dsRight) & "; DescriptionRight=" & CStr(labelRight) & "."
    commonRight = singleRight
    parentWidth = CLng(frm.Width)
    If parentWidth - commonRight > MAX_DECORATED_PARENT_WIDTH_EXPANSION_TWIPS Then Err.Raise vbObjectError + 882, MODULE_NAME & ".StabilizeDecoratedParentWidth", "Access expanded the parent Width beyond the safe normalization allowance before persistence. ParentWidth=" & CStr(parentWidth) & "; CommonControlRightEdge=" & CStr(commonRight) & "; Expansion=" & CStr(parentWidth - commonRight) & "; MaximumAllowedExpansion=" & CStr(MAX_DECORATED_PARENT_WIDTH_EXPANSION_TWIPS) & "."
    If parentWidth < commonRight Then
        frm.Width = commonRight
        actionText = "Expanded parent Width to the generated controls"
    ElseIf parentWidth - commonRight > DECORATED_PARENT_RIGHT_EDGE_TOLERANCE_TWIPS Then
        singleCtl.Width = parentWidth - CLng(singleCtl.Left)
        dsCtl.Width = parentWidth - CLng(dsCtl.Left)
        descriptionLabel.Width = parentWidth - CLng(descriptionLabel.Left)
        ResizeDescriptionBandForCurrentWidth frm, formName, descriptionLabel
        AddOrReplaceADCButtons frm, formName, visualColumnCount
        actionText = "Expanded generated controls and button rows to the current parent Width"
    Else
        actionText = "Verified current right-edge alignment"
    End If
    WriteADCDiagnostic "BEFORE", "Single width persistence save", "Form=" & formName & "; Action=" & actionText & "; WidthBefore=" & CStr(widthBefore)
    DoCmd.Save acForm, formName
    WriteADCDiagnostic "AFTER", "Single width persistence save", formName
    WriteADCDiagnostic "BEFORE", "Single width persistence close", formName
    DoCmd.Close acForm, formName, acSaveYes
    Set descriptionLabel = Nothing
    Set singleCtl = Nothing
    Set dsCtl = Nothing
    Set frm = Nothing
    WriteADCDiagnostic "AFTER", "Single width persistence close", formName
    WriteADCDiagnostic "BEFORE", "Single width persistence reopen", formName
    DoCmd.OpenForm formName, acDesign, , , , acHidden
    Set frm = forms(formName)
    Set descriptionLabel = frm.Controls(DESCRIPTION_LABEL_NAME)
    Set singleCtl = frm.Controls(SINGLE_SUBFORM_CONTROL_NAME)
    Set dsCtl = frm.Controls(DS_SUBFORM_CONTROL_NAME)
    WriteADCDiagnostic "AFTER", "Single width persistence reopen", formName
    If Abs(CLng(singleCtl.Left) - expectedSingleLeft) > DECORATED_PARENT_RIGHT_EDGE_TOLERANCE_TWIPS Or Abs(CLng(singleCtl.Height) - expectedSingleHeight) > DECORATED_PARENT_RIGHT_EDGE_TOLERANCE_TWIPS Then Err.Raise vbObjectError + 852, MODULE_NAME & ".StabilizeDecoratedParentWidth", "SingleForm moved horizontally or changed height after the one-cycle width persistence check. ExpectedLeft=" & CStr(expectedSingleLeft) & "; ActualLeft=" & CStr(singleCtl.Left) & "; ExpectedHeight=" & CStr(expectedSingleHeight) & "; ActualHeight=" & CStr(singleCtl.Height) & "."
    If Abs(CLng(dsCtl.Left) - expectedDSLeft) > DECORATED_PARENT_RIGHT_EDGE_TOLERANCE_TWIPS Or Abs(CLng(dsCtl.Height) - expectedDSHeight) > DECORATED_PARENT_RIGHT_EDGE_TOLERANCE_TWIPS Then Err.Raise vbObjectError + 853, MODULE_NAME & ".StabilizeDecoratedParentWidth", "DSForm moved horizontally or changed height after the one-cycle width persistence check. ExpectedLeft=" & CStr(expectedDSLeft) & "; ActualLeft=" & CStr(dsCtl.Left) & "; ExpectedHeight=" & CStr(expectedDSHeight) & "; ActualHeight=" & CStr(dsCtl.Height) & "."
    parentWidth = CLng(frm.Width)
    singleRight = CLng(singleCtl.Left) + CLng(singleCtl.Width)
    dsRight = CLng(dsCtl.Left) + CLng(dsCtl.Width)
    labelRight = CLng(descriptionLabel.Left) + CLng(descriptionLabel.Width)
    commonRight = singleRight
    If dsRight > commonRight Then commonRight = dsRight
    If labelRight > commonRight Then commonRight = labelRight
    parentExcess = parentWidth - commonRight
    WriteADCDiagnostic "VERIFY", "Single width persistence result", "Form=" & formName & "; ParentWidth=" & CStr(parentWidth) & "; SingleRight=" & CStr(singleRight) & "; DSRight=" & CStr(dsRight) & "; DescriptionRight=" & CStr(labelRight) & "; ParentExcess=" & CStr(parentExcess)
    If Abs(singleRight - dsRight) > DECORATED_PARENT_RIGHT_EDGE_TOLERANCE_TWIPS Or Abs(singleRight - labelRight) > DECORATED_PARENT_RIGHT_EDGE_TOLERANCE_TWIPS Then Err.Raise vbObjectError + 855, MODULE_NAME & ".StabilizeDecoratedParentWidth", "The decorated controls did not retain one shared right edge after the one-cycle persistence check. Form=" & formName & "; SingleRight=" & CStr(singleRight) & "; DSRight=" & CStr(dsRight) & "; DescriptionRight=" & CStr(labelRight) & "."
    If parentExcess < -DECORATED_PARENT_RIGHT_EDGE_TOLERANCE_TWIPS Then Err.Raise vbObjectError + 856, MODULE_NAME & ".StabilizeDecoratedParentWidth", "The persisted parent Width clips the generated controls after the one-cycle persistence check. Form=" & formName & "; ParentWidth=" & CStr(parentWidth) & "; CommonControlRightEdge=" & CStr(commonRight) & "; Difference=" & CStr(parentExcess) & "."
    If parentExcess > MAX_DECORATED_PARENT_WIDTH_EXPANSION_TWIPS Then Err.Raise vbObjectError + 882, MODULE_NAME & ".StabilizeDecoratedParentWidth", "Access expanded the persisted parent Width beyond the safe normalization allowance. Form=" & formName & "; ParentWidth=" & CStr(parentWidth) & "; CommonControlRightEdge=" & CStr(commonRight) & "; Expansion=" & CStr(parentExcess) & "; MaximumAllowedExpansion=" & CStr(MAX_DECORATED_PARENT_WIDTH_EXPANSION_TWIPS) & "."
    VerifyOpenADCButtons frm, visualColumnCount
    If parentExcess > DECORATED_PARENT_RIGHT_EDGE_TOLERANCE_TWIPS Then
        Debug.Print MODULE_NAME & ": Access retained bounded right-side space after the single width persistence cycle. Form=" & formName & "; ParentWidth=" & CStr(parentWidth) & "; CommonControlRightEdge=" & CStr(commonRight) & "; Difference=" & CStr(parentExcess)
    Else
        Debug.Print MODULE_NAME & ": Decorated parent Width and generated horizontal spans persisted on one right edge after a single save-close-reopen cycle. Form=" & formName & "; ParentWidth=" & CStr(parentWidth)
    End If
    Exit Sub
ErrHandler:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Raise errNumber, MODULE_NAME & ".StabilizeDecoratedParentWidth", "Unable to complete the one-cycle decorated-parent Width persistence check for '" & formName & "'. " & errDescription
End Sub

Private Sub ResizeDescriptionBandForCurrentWidth(ByVal frm As Access.Form, ByVal formName As String, ByVal descriptionLabel As Access.control)
    On Error GoTo ErrHandler
    Dim labelControl As Access.label
    Dim oldBandHeight As Long
    Dim newBandHeight As Long
    Dim newLabelHeight As Long
    Dim shiftAmount As Long
    Dim detailHeight As Long
    Set labelControl = descriptionLabel
    oldBandHeight = CLng(descriptionLabel.Top) + CLng(descriptionLabel.Height) + DESCRIPTION_LABEL_BOTTOM_PADDING_TWIPS
    newLabelHeight = MeasuredDescriptionLabelHeight(labelControl, CStr(descriptionLabel.Caption), CLng(descriptionLabel.Width))
    newBandHeight = CLng(descriptionLabel.Top) + newLabelHeight + DESCRIPTION_LABEL_BOTTOM_PADDING_TWIPS
    If CLng(descriptionLabel.Section) = acHeader Then
        descriptionLabel.Height = newLabelHeight
        frm.Section(acHeader).Visible = True
        frm.Section(acHeader).Height = newBandHeight
    ElseIf CLng(descriptionLabel.Section) = acDetail Then
        shiftAmount = newBandHeight - oldBandHeight
        detailHeight = CLng(frm.Section(acDetail).Height)
        If detailHeight + shiftAmount < newBandHeight Then Err.Raise vbObjectError + 879, MODULE_NAME & ".ResizeDescriptionBandForCurrentWidth", "The Detail section cannot be resized to the remeasured description band. CurrentDetailHeight=" & CStr(detailHeight) & "; Shift=" & CStr(shiftAmount) & "; RequiredBandHeight=" & CStr(newBandHeight) & "."
        If detailHeight + shiftAmount > MAX_ACCESS_SECTION_HEIGHT_TWIPS Then Err.Raise vbObjectError + 880, MODULE_NAME & ".ResizeDescriptionBandForCurrentWidth", "The remeasured description band would exceed the Access Detail-section height limit. CurrentDetailHeight=" & CStr(detailHeight) & "; Shift=" & CStr(shiftAmount) & "."
        If shiftAmount > 0 Then
            frm.Section(acDetail).Height = detailHeight + shiftAmount
            ShiftDetailControlsExceptDescription frm, shiftAmount
        ElseIf shiftAmount < 0 Then
            ShiftDetailControlsExceptDescription frm, shiftAmount
            frm.Section(acDetail).Height = detailHeight + shiftAmount
        End If
        descriptionLabel.Height = newLabelHeight
    Else
        Err.Raise vbObjectError + 881, MODULE_NAME & ".ResizeDescriptionBandForCurrentWidth", "The generated description label is not in the Form Header or Detail section. ActualSection=" & CStr(descriptionLabel.Section) & "."
    End If
    Debug.Print MODULE_NAME & ": Remeasured description band after horizontal expansion. Form=" & formName & "; LabelWidth=" & CStr(descriptionLabel.Width) & "; OldBandHeight=" & CStr(oldBandHeight) & "; NewBandHeight=" & CStr(newBandHeight) & "; LabelHeight=" & CStr(newLabelHeight)
    Exit Sub
ErrHandler:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Raise errNumber, MODULE_NAME & ".ResizeDescriptionBandForCurrentWidth", "Unable to resize the generated description band after the persisted parent Width changed on '" & formName & "'. " & errDescription
End Sub

Private Sub ShiftDetailControlsExceptDescription(ByVal frm As Access.Form, ByVal offsetTwips As Long)
    On Error GoTo ErrHandler
    Dim ctl As Access.control
    For Each ctl In frm.Controls
        If CLng(ctl.Section) = acDetail And StrComp(ctl.Name, DESCRIPTION_LABEL_NAME, vbBinaryCompare) <> 0 Then ctl.Top = CLng(ctl.Top) + offsetTwips
    Next ctl
    Exit Sub
ErrHandler:
    Dim errNumber As Long
    Dim errDescription As String
    Dim controlName As String
    errNumber = Err.Number
    errDescription = Err.Description
    On Error Resume Next
    controlName = SafeControlName(ctl)
    On Error GoTo 0
    Err.Raise errNumber, MODULE_NAME & ".ShiftDetailControlsExceptDescription", "Unable to shift Detail control '" & controlName & "' by " & CStr(offsetTwips) & " twips while resizing the generated description band. " & errDescription
End Sub

Private Function MeasuredDescriptionLabelHeight(ByVal ctl As Access.label, ByVal labelCaption As String, ByVal labelWidth As Long) As Long
    On Error GoTo ErrHandler
    Dim savedCaption As String
    Dim savedWidth As Long
    Dim savedHeight As Long
    Dim lineCount As Long
    Dim oneLineHeight As Long
    Dim twoLineHeight As Long
    Dim lineAdvance As Long
    Dim measuredHeight As Long
    savedCaption = CStr(ctl.Caption)
    savedWidth = CLng(ctl.Width)
    savedHeight = CLng(ctl.Height)
    If labelWidth <= 0 Then Err.Raise vbObjectError + 842, MODULE_NAME & ".MeasuredDescriptionLabelHeight", "The generated description label has no usable text width. LabelWidth=" & CStr(labelWidth) & "."
    lineCount = MeasuredWrappedDescriptionLineCount(ctl, labelCaption, labelWidth)
    oneLineHeight = MeasuredLabelHeightForText(ctl, "Ag")
    twoLineHeight = MeasuredLabelHeightForText(ctl, "Ag" & vbCrLf & "Ag")
    lineAdvance = twoLineHeight - oneLineHeight
    If lineCount <= 0 Then Err.Raise vbObjectError + 843, MODULE_NAME & ".MeasuredDescriptionLabelHeight", "The generated description measurement returned no display lines."
    If oneLineHeight <= 0 Or lineAdvance <= 0 Then Err.Raise vbObjectError + 844, MODULE_NAME & ".MeasuredDescriptionLabelHeight", "Access returned unusable label height measurements. OneLineHeight=" & CStr(oneLineHeight) & "; TwoLineHeight=" & CStr(twoLineHeight) & "."
    measuredHeight = oneLineHeight + ((lineCount - 1) * lineAdvance)
    If DESCRIPTION_LABEL_TOP_TWIPS + measuredHeight + DESCRIPTION_LABEL_BOTTOM_PADDING_TWIPS > MAX_ACCESS_SECTION_HEIGHT_TWIPS Then Err.Raise vbObjectError + 845, MODULE_NAME & ".MeasuredDescriptionLabelHeight", "The generated description requires a label band taller than the Access section-height limit. CaptionLength=" & CStr(Len(labelCaption)) & "; Lines=" & CStr(lineCount) & "; MeasuredHeight=" & CStr(measuredHeight) & "."
    ctl.Caption = savedCaption
    ctl.Width = savedWidth
    ctl.Height = savedHeight
    MeasuredDescriptionLabelHeight = measuredHeight
    Exit Function
ErrHandler:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    On Error Resume Next
    ctl.Caption = savedCaption
    ctl.Width = savedWidth
    ctl.Height = savedHeight
    On Error GoTo 0
    Err.Raise errNumber, MODULE_NAME & ".MeasuredDescriptionLabelHeight", "Unable to measure the generated description label. " & errDescription
End Function

Private Function MeasuredWrappedDescriptionLineCount(ByVal ctl As Access.label, ByVal labelCaption As String, ByVal labelWidth As Long) As Long
    Dim normalizedCaption As String
    Dim paragraphs As Variant
    Dim paragraphValue As Variant
    Dim lineCount As Long
    normalizedCaption = Replace$(labelCaption, vbCrLf, vbLf)
    normalizedCaption = Replace$(normalizedCaption, vbCr, vbLf)
    paragraphs = Split(normalizedCaption, vbLf)
    For Each paragraphValue In paragraphs
        lineCount = lineCount + MeasuredParagraphLineCount(ctl, CStr(paragraphValue), labelWidth)
    Next paragraphValue
    MeasuredWrappedDescriptionLineCount = lineCount
End Function

Private Function MeasuredParagraphLineCount(ByVal ctl As Access.label, ByVal paragraphText As String, ByVal labelWidth As Long) As Long
    Dim sourceText As String
    Dim sourcePosition As Long
    Dim tokenText As String
    Dim currentLine As String
    Dim candidateLine As String
    Dim prefixLength As Long
    Dim lineCount As Long
    sourceText = RTrim$(paragraphText)
    If Len(sourceText) = 0 Then
        MeasuredParagraphLineCount = 1
        Exit Function
    End If
    sourcePosition = 1
    Do While sourcePosition <= Len(sourceText)
        tokenText = NextDescriptionWrapToken(sourceText, sourcePosition)
        If Len(currentLine) = 0 Then
            candidateLine = LTrim$(tokenText)
        Else
            candidateLine = currentLine & tokenText
        End If
        If MeasuredTextFitsWidth(ctl, candidateLine, labelWidth) Then
            currentLine = candidateLine
        Else
            If Len(currentLine) > 0 Then
                lineCount = lineCount + 1
                currentLine = LTrim$(tokenText)
            Else
                currentLine = LTrim$(tokenText)
            End If
            Do While Len(currentLine) > 0 And Not MeasuredTextFitsWidth(ctl, currentLine, labelWidth)
                prefixLength = MaximumMeasuredPrefixLength(ctl, currentLine, labelWidth)
                If prefixLength <= 0 Then Err.Raise vbObjectError + 846, MODULE_NAME & ".MeasuredParagraphLineCount", "Unable to fit even one description character within the generated label width."
                lineCount = lineCount + 1
                currentLine = LTrim$(Mid$(currentLine, prefixLength + 1))
            Loop
        End If
    Loop
    If Len(currentLine) > 0 Then lineCount = lineCount + 1
    If lineCount = 0 Then lineCount = 1
    MeasuredParagraphLineCount = lineCount
End Function

Private Function NextDescriptionWrapToken(ByVal sourceText As String, ByRef sourcePosition As Long) As String
    Dim tokenStart As Long
    tokenStart = sourcePosition
    Do While sourcePosition <= Len(sourceText) And Mid$(sourceText, sourcePosition, 1) = " "
        sourcePosition = sourcePosition + 1
    Loop
    Do While sourcePosition <= Len(sourceText) And Mid$(sourceText, sourcePosition, 1) <> " "
        sourcePosition = sourcePosition + 1
    Loop
    NextDescriptionWrapToken = Mid$(sourceText, tokenStart, sourcePosition - tokenStart)
End Function

Private Function MeasuredTextFitsWidth(ByVal ctl As Access.label, ByVal textValue As String, ByVal labelWidth As Long) As Boolean
    If Len(textValue) = 0 Then
        MeasuredTextFitsWidth = True
    Else
        MeasuredTextFitsWidth = (MeasuredLabelWidthForText(ctl, textValue) <= labelWidth)
    End If
End Function

Private Function MaximumMeasuredPrefixLength(ByVal ctl As Access.label, ByVal textValue As String, ByVal labelWidth As Long) As Long
    Dim lowValue As Long
    Dim highValue As Long
    Dim middleValue As Long
    Dim bestValue As Long
    Dim probeValue As Long
    If Len(textValue) = 0 Then Exit Function
    probeValue = 1
    Do
        If MeasuredLabelWidthForText(ctl, Left$(textValue, probeValue)) <= labelWidth Then
            bestValue = probeValue
            If probeValue = Len(textValue) Then
                MaximumMeasuredPrefixLength = bestValue
                Exit Function
            End If
            If probeValue > Len(textValue) \ 2 Then
                probeValue = Len(textValue)
            Else
                probeValue = probeValue * 2
            End If
        Else
            Exit Do
        End If
    Loop
    lowValue = bestValue + 1
    highValue = probeValue - 1
    Do While lowValue <= highValue
        middleValue = (lowValue + highValue) \ 2
        If MeasuredLabelWidthForText(ctl, Left$(textValue, middleValue)) <= labelWidth Then
            bestValue = middleValue
            lowValue = middleValue + 1
        Else
            highValue = middleValue - 1
        End If
    Loop
    MaximumMeasuredPrefixLength = bestValue
End Function

Private Function MeasuredLabelWidthForText(ByVal ctl As Access.label, ByVal textValue As String) As Long
    ctl.Caption = textValue
    ctl.Width = 1
    ctl.Height = 1
    ctl.SizeToFit
    MeasuredLabelWidthForText = CLng(ctl.Width)
End Function

Private Function MeasuredLabelHeightForText(ByVal ctl As Access.label, ByVal textValue As String) As Long
    ctl.Caption = textValue
    ctl.Width = 1
    ctl.Height = 1
    ctl.SizeToFit
    MeasuredLabelHeightForText = CLng(ctl.Height)
End Function

Private Function TryGetControlHorizontalSpan(ByVal frm As Access.Form, ByVal controlName As String, ByRef controlLeft As Long, ByRef controlRight As Long) As Boolean
    On Error GoTo ExitHere
    Dim ctl As Access.control
    Set ctl = frm.Controls(controlName)
    controlLeft = CLng(ctl.Left)
    controlRight = controlLeft + CLng(ctl.Width)
    TryGetControlHorizontalSpan = controlRight > controlLeft
ExitHere:
    Set ctl = Nothing
    Err.Clear
End Function

Private Function ExistingHorizontalContentRightEdge(ByVal frm As Access.Form) As Long
    On Error GoTo ErrHandler
    Dim ctl As Access.control
    Dim controlRightEdge As Long
    Dim rightEdge As Long
    rightEdge = CLng(frm.Width)
    For Each ctl In frm.Controls
        If StrComp(ctl.Name, DESCRIPTION_LABEL_NAME, vbBinaryCompare) <> 0 Then
            controlRightEdge = CLng(ctl.Left) + CLng(ctl.Width)
            If controlRightEdge > rightEdge Then rightEdge = controlRightEdge
        End If
    Next ctl
    ExistingHorizontalContentRightEdge = rightEdge
    Exit Function
ErrHandler:
    Err.Raise Err.Number, MODULE_NAME & ".ExistingHorizontalContentRightEdge", "Unable to determine the existing horizontal form span. " & Err.Description
End Function

Private Sub VerifyOpenDescriptionLabel(ByVal frm As Access.Form, ByVal expectedCaption As String, ByVal expectedSection As Long, ByVal expectedLeft As Long, ByVal expectedWidth As Long, ByVal expectedHeight As Long)
    On Error GoTo ErrHandler
    Dim ctl As Access.control
    Dim overlapProblems As String
    Set ctl = frm.Controls(DESCRIPTION_LABEL_NAME)
    If CLng(ctl.controlType) <> acLabel Then Err.Raise vbObjectError + 824, MODULE_NAME & ".VerifyOpenDescriptionLabel", "The generated description control is not a label."
    If StrComp(CStr(ctl.Caption), expectedCaption, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 825, MODULE_NAME & ".VerifyOpenDescriptionLabel", "The generated description label caption does not match the requested caption."
    If CLng(ctl.Section) <> expectedSection Then Err.Raise vbObjectError + 826, MODULE_NAME & ".VerifyOpenDescriptionLabel", "The generated description label was created in the wrong section."
    If Abs(CLng(ctl.Left) - expectedLeft) > 2 Or Abs(CLng(ctl.Top) - DESCRIPTION_LABEL_TOP_TWIPS) > 2 Then Err.Raise vbObjectError + 827, MODULE_NAME & ".VerifyOpenDescriptionLabel", "The generated description label position does not match the requested position."
    If Abs(CLng(ctl.Width) - expectedWidth) > 2 Or Abs(CLng(ctl.Height) - expectedHeight) > 2 Then Err.Raise vbObjectError + 828, MODULE_NAME & ".VerifyOpenDescriptionLabel", "The generated description label dimensions do not match the calculated dimensions."
    If CLng(ctl.properties("HorizontalAnchor").value) <> DESCRIPTION_LABEL_HORIZONTAL_ANCHOR_BOTH Then Err.Raise vbObjectError + 838, MODULE_NAME & ".VerifyOpenDescriptionLabel", "The generated description label is not anchored to both horizontal edges."
    If expectedSection = acHeader Then
        If Abs(CLng(frm.Section(acHeader).Height) - (DESCRIPTION_LABEL_TOP_TWIPS + expectedHeight + DESCRIPTION_LABEL_BOTTOM_PADDING_TWIPS)) > 2 Then Err.Raise vbObjectError + 841, MODULE_NAME & ".VerifyOpenDescriptionLabel", "The generated form header retains excess or insufficient vertical space."
    End If
    overlapProblems = DescriptionLabelOverlapProblems(frm, ctl)
    If Len(overlapProblems) > 0 Then Err.Raise vbObjectError + 829, MODULE_NAME & ".VerifyOpenDescriptionLabel", overlapProblems
    Exit Sub
ErrHandler:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    Err.Raise errNumber, MODULE_NAME & ".VerifyOpenDescriptionLabel", "Unable to verify the generated description label before saving. " & errDescription
End Sub

Private Function DescriptionLabelOverlapProblems(ByVal frm As Access.Form, ByVal descriptionLabel As Access.control) As String
    Dim ctl As Access.control
    Dim problems As String
    For Each ctl In frm.Controls
        If StrComp(ctl.Name, descriptionLabel.Name, vbBinaryCompare) <> 0 Then
            If CLng(ctl.Section) = CLng(descriptionLabel.Section) Then
                If ControlsOverlap(descriptionLabel, ctl) Then problems = problems & "Generated description label overlaps control '" & ctl.Name & "'." & vbCrLf
            End If
        End If
    Next ctl
    DescriptionLabelOverlapProblems = problems
End Function

Private Function ControlsOverlap(ByVal firstControl As Access.control, ByVal secondControl As Access.control) As Boolean
    Dim firstRight As Long
    Dim firstBottom As Long
    Dim secondRight As Long
    Dim secondBottom As Long
    firstRight = CLng(firstControl.Left) + CLng(firstControl.Width)
    firstBottom = CLng(firstControl.Top) + CLng(firstControl.Height)
    secondRight = CLng(secondControl.Left) + CLng(secondControl.Width)
    secondBottom = CLng(secondControl.Top) + CLng(secondControl.Height)
    ControlsOverlap = (CLng(firstControl.Left) < secondRight And firstRight > CLng(secondControl.Left) And CLng(firstControl.Top) < secondBottom And firstBottom > CLng(secondControl.Top))
End Function

' =============================================================================================================================================================================
' PERSISTED VERIFICATION / DIAGNOSTIC PROCEDURES
' =============================================================================================================================================================================

Private Sub VerifyDecoratedGeneratedFormSet(ByRef formNames() As String, ByVal buildId As String, ByVal expectedCaption As String, ByVal visualColumnCount As Long)
    WriteADCDiagnostic "BEFORE", "VerifyPersistedDescriptionLabel", formNames(FORM_INDEX_BOTH)
    VerifyPersistedDescriptionLabel formNames(FORM_INDEX_BOTH), expectedCaption, buildId, formNames(FORM_INDEX_FORM), formNames(FORM_INDEX_DS), visualColumnCount
    WriteADCDiagnostic "AFTER", "VerifyPersistedDescriptionLabel", formNames(FORM_INDEX_BOTH)
    WriteADCDiagnostic "BEFORE", "VerifyAllGeneratedFormsClosed after decoration", buildId
    VerifyAllGeneratedFormsClosed formNames, "After Access Development Console decoration"
    WriteADCDiagnostic "AFTER", "VerifyAllGeneratedFormsClosed after decoration", buildId
    WriteADCDiagnostic "BEFORE", "VerifyParentRuntimeHealth", formNames(FORM_INDEX_BOTH)
    VerifyParentRuntimeHealth formNames(FORM_INDEX_BOTH)
    WriteADCDiagnostic "AFTER", "VerifyParentRuntimeHealth", formNames(FORM_INDEX_BOTH)
    WriteADCDiagnostic "BEFORE", "VerifyAllGeneratedFormsClosed after runtime", buildId
    VerifyAllGeneratedFormsClosed formNames, "After Access Development Console runtime verification"
    WriteADCDiagnostic "AFTER", "VerifyAllGeneratedFormsClosed after runtime", buildId
    Debug.Print MODULE_NAME & ": Decorated generated form set verification passed. BuildId=" & buildId
End Sub

Private Sub VerifyPersistedDescriptionLabel(ByVal formName As String, ByVal expectedCaption As String, ByVal buildId As String, ByVal expectedSingleFormName As String, ByVal expectedDSFormName As String, ByVal visualColumnCount As Long)
    On Error GoTo ErrHandler
    Dim frm As Access.Form
    Dim ctl As Access.label
    Dim openedHere As Boolean
    Dim problems As String
    Dim persistedSection As Long
    Dim persistedWidth As Long
    Dim persistedHeight As Long
    Dim expectedLeft As Long
    Dim expectedWidth As Long
    Dim expectedHeight As Long
    Dim expectedHeaderHeight As Long
    Dim singleRaw As String
    Dim dsRaw As String
    Dim singleParsed As String
    Dim dsParsed As String
    Dim diagnosticText As String
    If SharedAccess_ObjectIsOpenStrict(acForm, formName) Then Err.Raise vbObjectError + 830, MODULE_NAME & ".VerifyPersistedDescriptionLabel", "The Both form must be closed before persisted decoration verification: " & formName
    DoCmd.OpenForm formName, acDesign, , , , acHidden
    openedHere = True
    Set frm = forms(formName)
    If StrComp(Nz(frm.Tag, vbNullString), ExpectedFinalTag(formName, buildId), vbBinaryCompare) <> 0 Then problems = problems & "  - Decorated Both form ownership tag does not match the completed build." & vbCrLf
    Set ctl = frm.Controls(DESCRIPTION_LABEL_NAME)
    If CLng(ctl.controlType) <> acLabel Then problems = problems & "  - Generated description control is not a label." & vbCrLf
    If StrComp(CStr(ctl.Caption), expectedCaption, vbBinaryCompare) <> 0 Then problems = problems & "  - Generated description caption does not match the expected caption." & vbCrLf
    ResolveDescriptionLabelHorizontalBounds frm, expectedLeft, expectedWidth
    expectedHeight = MeasuredDescriptionLabelHeight(ctl, expectedCaption, expectedWidth)
    If Abs(CLng(ctl.Left) - expectedLeft) > 2 Or Abs(CLng(ctl.Top) - DESCRIPTION_LABEL_TOP_TWIPS) > 2 Then problems = problems & "  - Generated description label position is incorrect. ExpectedLeft=" & CStr(expectedLeft) & "; ActualLeft=" & CStr(ctl.Left) & "." & vbCrLf
    If Abs(CLng(ctl.Width) - expectedWidth) > 2 Or Abs(CLng(ctl.Height) - expectedHeight) > 2 Then problems = problems & "  - Generated description label dimensions do not match the calculated persisted dimensions. ExpectedWidth=" & CStr(expectedWidth) & "; ActualWidth=" & CStr(ctl.Width) & "; ExpectedHeight=" & CStr(expectedHeight) & "; ActualHeight=" & CStr(ctl.Height) & "." & vbCrLf
    If CLng(ctl.properties("HorizontalAnchor").value) <> DESCRIPTION_LABEL_HORIZONTAL_ANCHOR_BOTH Then problems = problems & "  - Generated description label is not anchored to both horizontal edges. ActualHorizontalAnchor=" & CStr(ctl.properties("HorizontalAnchor").value) & "." & vbCrLf
    If CLng(frm.Controls(SINGLE_SUBFORM_CONTROL_NAME).properties("HorizontalAnchor").value) <> DESCRIPTION_LABEL_HORIZONTAL_ANCHOR_BOTH Then problems = problems & "  - SingleForm is not anchored to both horizontal edges. ActualHorizontalAnchor=" & CStr(frm.Controls(SINGLE_SUBFORM_CONTROL_NAME).properties("HorizontalAnchor").value) & "." & vbCrLf
    If CLng(frm.Controls(DS_SUBFORM_CONTROL_NAME).properties("HorizontalAnchor").value) <> DESCRIPTION_LABEL_HORIZONTAL_ANCHOR_BOTH Then problems = problems & "  - DSForm is not anchored to both horizontal edges. ActualHorizontalAnchor=" & CStr(frm.Controls(DS_SUBFORM_CONTROL_NAME).properties("HorizontalAnchor").value) & "." & vbCrLf
    If CLng(frm.ScrollBars) <> 0 Then problems = problems & "  - Decorated Both form ScrollBars must remain disabled. ActualScrollBars=" & CStr(frm.ScrollBars) & "." & vbCrLf
    If CBool(frm.RecordSelectors) Then problems = problems & "  - Decorated Both form RecordSelectors must remain disabled." & vbCrLf
    If CBool(frm.NavigationButtons) Then problems = problems & "  - Decorated Both form NavigationButtons must remain disabled." & vbCrLf
    If Not CBool(frm.PopUp) Then problems = problems & "  - Decorated Both form PopUp must be True." & vbCrLf
    If CBool(frm.Modal) Then problems = problems & "  - Decorated Both form Modal must be False." & vbCrLf
    If Not CBool(frm.AutoCenter) Then problems = problems & "  - Decorated Both form AutoCenter must be True." & vbCrLf
    If Not CBool(frm.AutoResize) Then problems = problems & "  - Decorated Both form AutoResize must be True." & vbCrLf
    If Not CBool(frm.FitToScreen) Then problems = problems & "  - Decorated Both form FitToScreen must be True." & vbCrLf
    If CLng(ctl.Section) = acHeader Then
        expectedHeaderHeight = DESCRIPTION_LABEL_TOP_TWIPS + expectedHeight + DESCRIPTION_LABEL_BOTTOM_PADDING_TWIPS
        If expectedHeaderHeight < PERSISTED_FORM_HEADER_MINIMUM_TWIPS Then expectedHeaderHeight = PERSISTED_FORM_HEADER_MINIMUM_TWIPS
        If Abs(CLng(frm.Section(acHeader).Height) - expectedHeaderHeight) > 2 Then problems = problems & "  - Generated form header height does not match the persisted Access geometry. ExpectedHeaderHeight=" & CStr(expectedHeaderHeight) & "; ActualHeaderHeight=" & CStr(frm.Section(acHeader).Height) & "." & vbCrLf
    ElseIf CLng(ctl.Section) <> acDetail Then
        problems = problems & "  - Generated description label is not in the Form Header or Detail section." & vbCrLf
    End If
    singleRaw = Nz(frm.Controls(SINGLE_SUBFORM_CONTROL_NAME).sourceObject, vbNullString)
    dsRaw = Nz(frm.Controls(DS_SUBFORM_CONTROL_NAME).sourceObject, vbNullString)
    singleParsed = FormNameFromSourceObject(singleRaw)
    dsParsed = FormNameFromSourceObject(dsRaw)
    If StrComp(singleParsed, expectedSingleFormName, vbTextCompare) <> 0 Then problems = problems & "  - SingleForm does not reference expected form '" & expectedSingleFormName & "'. Actual SourceObject='" & singleRaw & "'." & vbCrLf
    If StrComp(dsParsed, expectedDSFormName, vbTextCompare) <> 0 Then problems = problems & "  - DSForm does not reference expected form '" & expectedDSFormName & "'. Actual SourceObject='" & dsRaw & "'." & vbCrLf
    problems = problems & DecoratedHorizontalRightEdgeProblems(frm, ctl)
    problems = problems & DecoratedVerticalBottomProblems(frm, ctl)
    problems = problems & ADCButtonProblems(frm, visualColumnCount)
    problems = problems & DescriptionLabelOverlapProblems(frm, ctl)
    If Len(problems) > 0 Then diagnosticText = PersistedDescriptionVerificationDiagnostics(frm, ctl, expectedCaption, buildId, expectedLeft, expectedWidth, expectedHeight, expectedSingleFormName, expectedDSFormName, singleRaw, dsRaw, visualColumnCount)
    persistedSection = CLng(ctl.Section)
    persistedWidth = CLng(ctl.Width)
    persistedHeight = CLng(ctl.Height)
    DoCmd.Close acForm, formName, acSaveNo
    openedHere = False
    Set ctl = Nothing
    Set frm = Nothing
    If Len(problems) > 0 Then Err.Raise vbObjectError + 831, MODULE_NAME & ".VerifyPersistedDescriptionLabel", "Persisted decorated Both-form verification failed for '" & formName & "':" & vbCrLf & problems & vbCrLf & diagnosticText
    Debug.Print MODULE_NAME & ": Persisted decorated Both form verified in one design-view pass. Form=" & formName & "; Section=" & DescriptionSectionName(persistedSection) & "; Left=" & CStr(expectedLeft) & "; Width=" & CStr(persistedWidth) & "; Height=" & CStr(persistedHeight) & "; SingleSource='" & singleParsed & "'; DSSource='" & dsParsed & "'; VisualColumns=" & CStr(visualColumnCount) & "; ButtonCount=" & CStr(visualColumnCount * ADC_BUTTON_ROW_COUNT)
    Exit Sub
ErrHandler:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    On Error Resume Next
    If openedHere Then DoCmd.Close acForm, formName, acSaveNo
    On Error GoTo 0
    Err.Raise errNumber, MODULE_NAME & ".VerifyPersistedDescriptionLabel", "Unable to verify the persisted decorated Both form '" & formName & "'. " & errDescription
End Sub

Private Function DecoratedHorizontalRightEdgeProblems(ByVal frm As Access.Form, ByVal descriptionLabel As Access.label) As String
    On Error GoTo ErrHandler
    Dim singleRight As Long
    Dim dsRight As Long
    Dim labelRight As Long
    Dim commonRight As Long
    Dim parentRight As Long
    Dim parentExcess As Long
    Dim problems As String
    singleRight = CLng(frm.Controls(SINGLE_SUBFORM_CONTROL_NAME).Left) + CLng(frm.Controls(SINGLE_SUBFORM_CONTROL_NAME).Width)
    dsRight = CLng(frm.Controls(DS_SUBFORM_CONTROL_NAME).Left) + CLng(frm.Controls(DS_SUBFORM_CONTROL_NAME).Width)
    labelRight = CLng(descriptionLabel.Left) + CLng(descriptionLabel.Width)
    commonRight = singleRight
    If dsRight > commonRight Then commonRight = dsRight
    If labelRight > commonRight Then commonRight = labelRight
    parentRight = CLng(frm.Width)
    If Abs(singleRight - dsRight) > DECORATED_PARENT_RIGHT_EDGE_TOLERANCE_TWIPS Then problems = problems & "  - SingleForm and DSForm do not share the same persisted right edge. SingleRight=" & CStr(singleRight) & "; DSRight=" & CStr(dsRight) & "; Tolerance=" & CStr(DECORATED_PARENT_RIGHT_EDGE_TOLERANCE_TWIPS) & "." & vbCrLf
    If Abs(singleRight - labelRight) > DECORATED_PARENT_RIGHT_EDGE_TOLERANCE_TWIPS Then problems = problems & "  - Generated description label does not share the subform controls' persisted right edge. SingleRight=" & CStr(singleRight) & "; LabelRight=" & CStr(labelRight) & "; Tolerance=" & CStr(DECORATED_PARENT_RIGHT_EDGE_TOLERANCE_TWIPS) & "." & vbCrLf
    parentExcess = parentRight - commonRight
    If parentExcess < -DECORATED_PARENT_RIGHT_EDGE_TOLERANCE_TWIPS Then problems = problems & "  - Decorated parent Width clips the generated controls. ParentWidth=" & CStr(parentRight) & "; CommonControlRightEdge=" & CStr(commonRight) & "; Difference=" & CStr(parentExcess) & "; Tolerance=" & CStr(DECORATED_PARENT_RIGHT_EDGE_TOLERANCE_TWIPS) & "." & vbCrLf
    If parentExcess > MAX_DECORATED_PARENT_WIDTH_EXPANSION_TWIPS Then problems = problems & "  - Decorated parent Width exceeds the bounded Access-normalization allowance. ParentWidth=" & CStr(parentRight) & "; CommonControlRightEdge=" & CStr(commonRight) & "; Difference=" & CStr(parentExcess) & "; MaximumAllowedExpansion=" & CStr(MAX_DECORATED_PARENT_WIDTH_EXPANSION_TWIPS) & "." & vbCrLf
    DecoratedHorizontalRightEdgeProblems = problems
    Exit Function
ErrHandler:
    Err.Raise Err.Number, MODULE_NAME & ".DecoratedHorizontalRightEdgeProblems", "Unable to verify the decorated parent form's persisted horizontal right-edge alignment. " & Err.Description
End Function

Private Function DecoratedVerticalBottomProblems(ByVal frm As Access.Form, ByVal descriptionLabel As Access.label) As String
    On Error GoTo ErrHandler
    Dim ctl As Access.control
    Dim requiredDetailHeight As Long
    Dim requiredFooterHeight As Long
    Dim detailHeight As Long
    Dim footerHeight As Long
    Dim unexpectedFooterControlCount As Long
    Dim controlBottom As Long
    Dim problems As String
    For Each ctl In frm.Controls
        If CLng(ctl.Section) = acDetail Then
            controlBottom = CLng(ctl.Top) + CLng(ctl.Height)
            If controlBottom > requiredDetailHeight Then requiredDetailHeight = controlBottom
        End If
    Next ctl
    detailHeight = CLng(frm.Section(acDetail).Height)
    If requiredDetailHeight <= 0 Then problems = problems & "  - Decorated parent Detail section contains no measurable controls." & vbCrLf
    If requiredDetailHeight > 0 And Abs(detailHeight - requiredDetailHeight) > DECORATED_PARENT_BOTTOM_EDGE_TOLERANCE_TWIPS Then problems = problems & "  - Decorated parent Detail section does not end at the lowest generated control. DetailHeight=" & CStr(detailHeight) & "; RequiredDetailHeight=" & CStr(requiredDetailHeight) & "; Difference=" & CStr(detailHeight - requiredDetailHeight) & "; Tolerance=" & CStr(DECORATED_PARENT_BOTTOM_EDGE_TOLERANCE_TWIPS) & "." & vbCrLf
    If CLng(descriptionLabel.Section) = acHeader Then
        requiredFooterHeight = RequiredGeneratedFooterHeight(frm, unexpectedFooterControlCount)
        footerHeight = CLng(frm.Section(acFooter).Height)
        If unexpectedFooterControlCount > 0 Then problems = problems & "  - Decorated parent Form Footer contains " & CStr(unexpectedFooterControlCount) & " unexpected non-generated control(s)." & vbCrLf
        If requiredFooterHeight <= 0 Then problems = problems & "  - Decorated parent Form Footer contains no generated Access Development Console buttons." & vbCrLf
        If requiredFooterHeight > 0 And Abs(footerHeight - requiredFooterHeight) > DECORATED_PARENT_BOTTOM_EDGE_TOLERANCE_TWIPS Then problems = problems & "  - Decorated parent Form Footer does not end at the generated button band. FooterHeight=" & CStr(footerHeight) & "; RequiredHeight=" & CStr(requiredFooterHeight) & "; Difference=" & CStr(footerHeight - requiredFooterHeight) & "; Tolerance=" & CStr(DECORATED_PARENT_BOTTOM_EDGE_TOLERANCE_TWIPS) & "." & vbCrLf
    End If
    DecoratedVerticalBottomProblems = problems
    Exit Function
ErrHandler:
    Err.Raise Err.Number, MODULE_NAME & ".DecoratedVerticalBottomProblems", "Unable to verify the decorated parent form's persisted bottom-edge alignment. " & Err.Description
End Function

Private Function PersistedDescriptionVerificationDiagnostics(ByVal frm As Access.Form, ByVal descriptionLabel As Access.label, ByVal expectedCaption As String, ByVal buildId As String, ByVal expectedLeft As Long, ByVal expectedWidth As Long, ByVal expectedHeight As Long, ByVal expectedSingleFormName As String, ByVal expectedDSFormName As String, ByVal singleRaw As String, ByVal dsRaw As String, ByVal visualColumnCount As Long) As String
    On Error GoTo DiagnosticFailure
    Dim singleCtl As Access.control
    Dim dsCtl As Access.control
    Dim singleRight As Long
    Dim dsRight As Long
    Dim labelRight As Long
    Dim commonRight As Long
    Dim parentWidth As Long
    Dim parentExcess As Long
    Dim widthAccepted As Boolean
    Dim hintText As String
    Dim resultText As String
    Set singleCtl = frm.Controls(SINGLE_SUBFORM_CONTROL_NAME)
    Set dsCtl = frm.Controls(DS_SUBFORM_CONTROL_NAME)
    singleRight = CLng(singleCtl.Left) + CLng(singleCtl.Width)
    dsRight = CLng(dsCtl.Left) + CLng(dsCtl.Width)
    labelRight = CLng(descriptionLabel.Left) + CLng(descriptionLabel.Width)
    commonRight = singleRight
    If dsRight > commonRight Then commonRight = dsRight
    If labelRight > commonRight Then commonRight = labelRight
    parentWidth = CLng(frm.Width)
    parentExcess = parentWidth - commonRight
    widthAccepted = (parentExcess >= -DECORATED_PARENT_RIGHT_EDGE_TOLERANCE_TWIPS And parentExcess <= MAX_DECORATED_PARENT_WIDTH_EXPANSION_TWIPS)
    If Not widthAccepted Then hintText = "Diagnostic hint: the one-cycle parent-width persistence check found clipping or expansion beyond the bounded Access-normalization allowance."
    resultText = "Persisted decoration diagnostic snapshot:" & vbCrLf
    resultText = resultText & "  Form: Name='" & frm.Name & "'; Width=" & CStr(parentWidth) & "; CommonControlRightEdge=" & CStr(commonRight) & "; ParentDifference=" & CStr(parentExcess) & "; RequiredTolerance=" & CStr(DECORATED_PARENT_RIGHT_EDGE_TOLERANCE_TWIPS) & "; WidthAccepted=" & CStr(widthAccepted) & "." & vbCrLf
    resultText = resultText & "  Form behavior: PopUp=" & SafeObjectPropertyText(frm, "PopUp") & "; Modal=" & SafeObjectPropertyText(frm, "Modal") & "; AutoCenter=" & SafeObjectPropertyText(frm, "AutoCenter") & "; AutoResize=" & SafeObjectPropertyText(frm, "AutoResize") & "; FitToScreen=" & SafeObjectPropertyText(frm, "FitToScreen") & "; ScrollBars=" & SafeObjectPropertyText(frm, "ScrollBars") & "; RecordSelectors=" & SafeObjectPropertyText(frm, "RecordSelectors") & "; NavigationButtons=" & SafeObjectPropertyText(frm, "NavigationButtons") & "; BorderStyle=" & SafeObjectPropertyText(frm, "BorderStyle") & "." & vbCrLf
    resultText = resultText & "  Sections: HeaderHeight=" & SafeFormSectionHeightText(frm, acHeader) & "; PersistedHeaderMinimum=" & CStr(PERSISTED_FORM_HEADER_MINIMUM_TWIPS) & "; DetailHeight=" & SafeFormSectionHeightText(frm, acDetail) & "; RequiredDetailHeight=" & SafeRequiredDetailHeightText(frm) & "; FooterHeight=" & SafeFormSectionHeightText(frm, acFooter) & "; RequiredFooterHeight=" & SafeRequiredFooterHeightText(frm) & "." & vbCrLf
    resultText = resultText & "  SingleForm: " & ControlGeometryDiagnostic(singleCtl) & "; SourceObject='" & singleRaw & "'; ExpectedForm='" & expectedSingleFormName & "'." & vbCrLf
    resultText = resultText & "  DSForm: " & ControlGeometryDiagnostic(dsCtl) & "; SourceObject='" & dsRaw & "'; ExpectedForm='" & expectedDSFormName & "'." & vbCrLf
    resultText = resultText & "  DescriptionLabel: " & ControlGeometryDiagnostic(descriptionLabel) & "; CaptionLength=" & CStr(Len(CStr(descriptionLabel.Caption))) & "; ExpectedCaptionLength=" & CStr(Len(expectedCaption)) & "." & vbCrLf
    resultText = resultText & "  Expected label geometry: Left=" & CStr(expectedLeft) & "; Width=" & CStr(expectedWidth) & "; Right=" & CStr(expectedLeft + expectedWidth) & "; Height=" & CStr(expectedHeight) & "; Top=" & CStr(DESCRIPTION_LABEL_TOP_TWIPS) & "." & vbCrLf
    resultText = resultText & "  Right edges: Single=" & CStr(singleRight) & "; DS=" & CStr(dsRight) & "; DescriptionLabel=" & CStr(labelRight) & "; ParentWidth=" & CStr(parentWidth) & "." & vbCrLf
    resultText = resultText & "  Buttons: VisualColumns=" & CStr(visualColumnCount) & "; ExpectedCount=" & CStr(visualColumnCount * ADC_BUTTON_ROW_COUNT) & "; ActualGeneratedCount=" & CStr(CountGeneratedADCButtons(frm)) & "; Names=ADC1Btn through ADC" & CStr(visualColumnCount * ADC_BUTTON_ROW_COUNT) & "Btn; OnClickExpectedBlank=True." & vbCrLf
    resultText = resultText & "  Ownership: ExpectedTag='" & ExpectedFinalTag(frm.Name, buildId) & "'; ActualTag='" & Nz(frm.Tag, vbNullString) & "'."
    If Len(hintText) > 0 Then resultText = resultText & vbCrLf & "  " & hintText
    PersistedDescriptionVerificationDiagnostics = resultText
    Exit Function
DiagnosticFailure:
    PersistedDescriptionVerificationDiagnostics = "Persisted decoration diagnostic snapshot could not be completed. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
End Function

Private Function ControlGeometryDiagnostic(ByVal ctl As Object) As String
    On Error GoTo DiagnosticFailure
    ControlGeometryDiagnostic = "Name='" & SafeObjectPropertyText(ctl, "Name") & "'; Type=" & SafeObjectPropertyText(ctl, "ControlType") & "; Section=" & SafeObjectPropertyText(ctl, "Section") & "; Left=" & SafeObjectPropertyText(ctl, "Left") & "; Top=" & SafeObjectPropertyText(ctl, "Top") & "; Width=" & SafeObjectPropertyText(ctl, "Width") & "; Height=" & SafeObjectPropertyText(ctl, "Height") & "; Right=" & SafeControlRightEdgeText(ctl) & "; Bottom=" & SafeControlBottomEdgeText(ctl) & "; HorizontalAnchor=" & SafeObjectPropertyText(ctl, "HorizontalAnchor") & "; VerticalAnchor=" & SafeObjectPropertyText(ctl, "VerticalAnchor")
    Exit Function
DiagnosticFailure:
    ControlGeometryDiagnostic = "<geometry unavailable; Error " & CStr(Err.Number) & ": " & Err.Description & ">"
End Function

Private Function SafeObjectPropertyText(ByVal targetObject As Object, ByVal propertyName As String) As String
    On Error GoTo ReadFailure
    Dim value As Variant
    value = CallByName(targetObject, propertyName, VbGet)
    If IsNull(value) Then
        SafeObjectPropertyText = "<Null>"
    ElseIf IsEmpty(value) Then
        SafeObjectPropertyText = "<Empty>"
    Else
        SafeObjectPropertyText = CStr(value)
    End If
    Exit Function
ReadFailure:
    SafeObjectPropertyText = "<unavailable; Error " & CStr(Err.Number) & ": " & Err.Description & ">"
End Function

Private Function SafeControlRightEdgeText(ByVal ctl As Object) As String
    On Error GoTo ReadFailure
    SafeControlRightEdgeText = CStr(CLng(CallByName(ctl, "Left", VbGet)) + CLng(CallByName(ctl, "Width", VbGet)))
    Exit Function
ReadFailure:
    SafeControlRightEdgeText = "<unavailable; Error " & CStr(Err.Number) & ": " & Err.Description & ">"
End Function

Private Function SafeControlBottomEdgeText(ByVal ctl As Object) As String
    On Error GoTo ReadFailure
    SafeControlBottomEdgeText = CStr(CLng(CallByName(ctl, "Top", VbGet)) + CLng(CallByName(ctl, "Height", VbGet)))
    Exit Function
ReadFailure:
    SafeControlBottomEdgeText = "<unavailable; Error " & CStr(Err.Number) & ": " & Err.Description & ">"
End Function

Private Function SafeRequiredDetailHeightText(ByVal frm As Access.Form) As String
    On Error GoTo ReadFailure
    Dim ctl As Access.control
    Dim requiredHeight As Long
    Dim controlBottom As Long
    For Each ctl In frm.Controls
        If CLng(ctl.Section) = acDetail Then
            controlBottom = CLng(ctl.Top) + CLng(ctl.Height)
            If controlBottom > requiredHeight Then requiredHeight = controlBottom
        End If
    Next ctl
    SafeRequiredDetailHeightText = CStr(requiredHeight)
    Exit Function
ReadFailure:
    SafeRequiredDetailHeightText = "<unavailable; Error " & CStr(Err.Number) & ": " & Err.Description & ">"
End Function

Private Function SafeRequiredFooterHeightText(ByVal frm As Access.Form) As String
    On Error GoTo ReadFailure
    Dim unexpectedFooterControlCount As Long
    Dim requiredHeight As Long
    requiredHeight = RequiredGeneratedFooterHeight(frm, unexpectedFooterControlCount)
    SafeRequiredFooterHeightText = CStr(requiredHeight)
    If unexpectedFooterControlCount > 0 Then SafeRequiredFooterHeightText = SafeRequiredFooterHeightText & " (unexpected controls=" & CStr(unexpectedFooterControlCount) & ")"
    Exit Function
ReadFailure:
    SafeRequiredFooterHeightText = "<unavailable; Error " & CStr(Err.Number) & ": " & Err.Description & ">"
End Function

Private Function SafeFormSectionHeightText(ByVal frm As Access.Form, ByVal sectionValue As AcSection) As String
    On Error GoTo ReadFailure
    SafeFormSectionHeightText = CStr(CLng(frm.Section(sectionValue).Height))
    Exit Function
ReadFailure:
    SafeFormSectionHeightText = "<unavailable; Error " & CStr(Err.Number) & ": " & Err.Description & ">"
End Function

Private Sub VerifyAllGeneratedFormsClosed(ByRef formNames() As String, ByVal contextText As String)
    Dim i As Long
    Dim problems As String
    For i = LBound(formNames) To UBound(formNames)
        If Not FormExists(formNames(i)) Then
            problems = problems & "  - Missing form: " & formNames(i) & vbCrLf
        ElseIf SharedAccess_ObjectIsOpenStrict(acForm, formNames(i)) Then
            problems = problems & "  - Form is unexpectedly open: " & formNames(i) & vbCrLf
        End If
    Next i
    If Len(problems) > 0 Then Err.Raise vbObjectError + 833, MODULE_NAME & ".VerifyAllGeneratedFormsClosed", contextText & " closed-state verification failed:" & vbCrLf & problems
    Debug.Print MODULE_NAME & ": Verified all generated forms closed. Context=" & contextText
End Sub

Private Sub VerifyParentRuntimeHealth(ByVal bothFormName As String)
    On Error GoTo ErrHandler
    Dim healthResult As Variant
    Dim openedHere As Boolean
    If SharedAccess_ObjectIsOpenStrict(acForm, bothFormName) Then Err.Raise vbObjectError + 834, MODULE_NAME & ".VerifyParentRuntimeHealth", "The Both form must be closed before runtime verification: " & bothFormName
    DoCmd.OpenForm bothFormName, acNormal, , , , acHidden
    openedHere = True
    healthResult = CallByName(forms(bothFormName), EVENT_PROCEDURE_RESULT_METHOD, VbMethod)
    If Len(Trim$(CStr(healthResult))) > 0 Then Err.Raise vbObjectError + 835, MODULE_NAME & ".VerifyParentRuntimeHealth", CStr(healthResult)
    DoCmd.Close acForm, bothFormName, acSaveNo
    openedHere = False
    Debug.Print MODULE_NAME & ": Decorated parent runtime health verification passed. Form=" & bothFormName
    Exit Sub
ErrHandler:
    Dim errNumber As Long
    Dim errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    On Error Resume Next
    If openedHere Then DoCmd.Close acForm, bothFormName, acSaveNo
    On Error GoTo 0
    Err.Raise errNumber, MODULE_NAME & ".VerifyParentRuntimeHealth", "Unable to verify decorated parent runtime health for '" & bothFormName & "'. " & errDescription
End Sub


Private Function DescriptionSectionName(ByVal sectionValue As Long) As String
    Select Case sectionValue
        Case acHeader
            DescriptionSectionName = "Form Header"
        Case acDetail
            DescriptionSectionName = "Detail"
        Case Else
            DescriptionSectionName = CStr(sectionValue)
    End Select
End Function

Private Function SafeControlName(ByVal ctl As Access.control) As String
    On Error GoTo ExitHere
    If ctl Is Nothing Then Exit Function
    SafeControlName = ctl.Name
ExitHere:
    If Len(SafeControlName) = 0 Then SafeControlName = "[unknown control]"
End Function

Private Sub DeleteControlIfExists(ByVal frm As Access.Form, ByVal formName As String, ByVal controlName As String)
    If Not ControlExistsOnFormInstance(frm, controlName) Then Exit Sub
    DeleteControl formName, controlName
    Debug.Print MODULE_NAME & ": Deleted existing generated control " & controlName & " from " & formName
End Sub

Private Function ControlExistsOnFormInstance(ByVal frm As Access.Form, ByVal controlName As String) As Boolean
    On Error GoTo ExitHere
    Dim ctl As Access.control
    Set ctl = frm.Controls(controlName)
    ControlExistsOnFormInstance = True
ExitHere:
    Set ctl = Nothing
    Err.Clear
End Function

' =============================================================================================================================================================================
' SOURCE DESCRIPTION / OBJECT-EXISTENCE HELPERS
' =============================================================================================================================================================================

' Purpose: Resolves form name from source object from the supplied/current ADC state, using `End` and `StrComp` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   sourceObjectText (Input; String): Input source object text for form name from source object; is passed to `normalizedText`.
' Returns: String: Returns the resolved name produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function FormNameFromSourceObject(ByVal sourceObjectText As String) As String
    'Extracts the form name from a SourceObject value after removing a leading Form. qualifier.
    Dim normalizedText As String
    normalizedText = Trim$(sourceObjectText)
    If Len(normalizedText) = 0 Then Exit Function
    If Left$(normalizedText, 1) = "=" Then normalizedText = Trim$(Mid$(normalizedText, 2))
    If Len(normalizedText) >= 2 Then
        If Left$(normalizedText, 1) = Chr$(34) And Right$(normalizedText, 1) = Chr$(34) Then normalizedText = Mid$(normalizedText, 2, Len(normalizedText) - 2)
    End If
    If Len(normalizedText) >= 7 Then
        If StrComp(Left$(normalizedText, 7), "[Form].", vbTextCompare) = 0 Or StrComp(Left$(normalizedText, 7), "[Form]!", vbTextCompare) = 0 Then normalizedText = Mid$(normalizedText, 8)
    End If
    If Len(normalizedText) >= 5 Then
        If StrComp(Left$(normalizedText, 5), "Form.", vbTextCompare) = 0 Or StrComp(Left$(normalizedText, 5), "Form!", vbTextCompare) = 0 Then normalizedText = Mid$(normalizedText, 6)
    End If
    If Len(normalizedText) >= 6 Then
        If StrComp(Left$(normalizedText, 6), "Forms!", vbTextCompare) = 0 Then normalizedText = Mid$(normalizedText, 7)
    End If
    normalizedText = Trim$(normalizedText)
    If Len(normalizedText) >= 2 Then
        If Left$(normalizedText, 1) = "[" And Right$(normalizedText, 1) = "]" Then normalizedText = Mid$(normalizedText, 2, Len(normalizedText) - 2)
    End If
    If Len(normalizedText) >= 2 Then
        If Left$(normalizedText, 1) = Chr$(34) And Right$(normalizedText, 1) = Chr$(34) Then normalizedText = Mid$(normalizedText, 2, Len(normalizedText) - 2)
    End If
    FormNameFromSourceObject = Trim$(normalizedText)
End Function

Private Function HeaderLabelCaption(ByVal recordSourceName As String, ByVal descriptionText As String) As String
    HeaderLabelCaption = "Record source: " & recordSourceName & "    Description: " & CleanDescriptionForCaption(descriptionText)
End Function


Private Function RecordSourceDescriptionForLabel(ByVal recordSourceName As String, ByRef descriptionStatus As String) As String
    Dim rawDescription As String
    Dim lookupFailed As Boolean
    If tableExists(recordSourceName) Then
        rawDescription = TableDescriptionValue(recordSourceName, lookupFailed)
        If lookupFailed Then
            descriptionStatus = "Table description unavailable"
            RecordSourceDescriptionForLabel = DESCRIPTION_UNAVAILABLE_TEXT
        ElseIf Len(Trim$(rawDescription)) = 0 Then
            descriptionStatus = "Table description missing"
            RecordSourceDescriptionForLabel = DESCRIPTION_MISSING_TEXT
        Else
            descriptionStatus = "Table description found"
            RecordSourceDescriptionForLabel = rawDescription
        End If
        Exit Function
    End If
    If queryExists(recordSourceName) Then
        rawDescription = QueryDescriptionValue(recordSourceName, lookupFailed)
        If lookupFailed Then
            descriptionStatus = "Query description unavailable"
            RecordSourceDescriptionForLabel = DESCRIPTION_UNAVAILABLE_TEXT
        ElseIf Len(Trim$(rawDescription)) = 0 Then
            descriptionStatus = "Query description missing"
            RecordSourceDescriptionForLabel = DESCRIPTION_MISSING_TEXT
        Else
            descriptionStatus = "Query description found"
            RecordSourceDescriptionForLabel = rawDescription
        End If
        Exit Function
    End If
    descriptionStatus = "Record source description unavailable"
    RecordSourceDescriptionForLabel = DESCRIPTION_UNAVAILABLE_TEXT
End Function

Private Function TableDescriptionValue(ByVal tableName As String, ByRef lookupFailed As Boolean) As String
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Set db = currentDb
    Set tdf = db.TableDefs(tableName)
    TableDescriptionValue = Nz(tdf.properties("Description").value, vbNullString)
ExitHere:
    Set tdf = Nothing
    Set db = Nothing
    Exit Function
ErrHandler:
    If Err.Number <> 3270 Then
        lookupFailed = True
        Debug.Print MODULE_NAME & ": Unable to read table description for " & tableName & ". Error " & CStr(Err.Number) & ": " & Err.Description
    End If
    Resume ExitHere
End Function

Private Function QueryDescriptionValue(ByVal queryName As String, ByRef lookupFailed As Boolean) As String
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim qdf As DAO.queryDef
    Set db = currentDb
    Set qdf = db.QueryDefs(queryName)
    QueryDescriptionValue = Nz(qdf.properties("Description").value, vbNullString)
ExitHere:
    Set qdf = Nothing
    Set db = Nothing
    Exit Function
ErrHandler:
    If Err.Number <> 3270 Then
        lookupFailed = True
        Debug.Print MODULE_NAME & ": Unable to read query description for " & queryName & ". Error " & CStr(Err.Number) & ": " & Err.Description
    End If
    Resume ExitHere
End Function

Private Sub MakeFormNames(ByVal formSetName As String, ByRef formNames() As String)
    formNames(FORM_INDEX_BOTH) = formSetName & SUFFIX_BOTH
    formNames(FORM_INDEX_CONT) = formSetName & SUFFIX_CONT
    formNames(FORM_INDEX_DS) = formSetName & SUFFIX_DS
    formNames(FORM_INDEX_FORM) = formSetName & SUFFIX_FORM
    formNames(FORM_INDEX_SPLIT) = formSetName & SUFFIX_SPLIT
End Sub


Private Function tableExists(ByVal tableName As String) As Boolean
    On Error GoTo ExitHere
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Set db = currentDb
    Set tdf = db.TableDefs(tableName)
    tableExists = True
ExitHere:
    Set tdf = Nothing
    Set db = Nothing
    Err.Clear
End Function

Private Function queryExists(ByVal queryName As String) As Boolean
    On Error GoTo ExitHere
    Dim db As DAO.Database
    Dim qdf As DAO.queryDef
    Set db = currentDb
    Set qdf = db.QueryDefs(queryName)
    queryExists = True
ExitHere:
    Set qdf = Nothing
    Set db = Nothing
    Err.Clear
End Function

Private Function FormExists(ByVal formName As String) As Boolean
    Dim obj As accessObject
    For Each obj In CurrentProject.AllForms
        If StrComp(obj.Name, formName, vbTextCompare) = 0 Then
            FormExists = True
            Exit Function
        End If
    Next obj
End Function

Private Sub TrySetControlProperty(ByVal ctl As Access.control, ByVal propertyName As String, ByVal propertyValue As Variant)
    On Error Resume Next
    ctl.properties(propertyName).value = propertyValue
    If Err.Number <> 0 Then Debug.Print MODULE_NAME & ": Optional control property not applied. Target=" & SafeControlName(ctl) & "; Property=" & propertyName & "; AttemptedValue='" & CStr(propertyValue) & "'; Error " & CStr(Err.Number) & ": " & Err.Description
    Err.Clear
    On Error GoTo 0
End Sub

' =============================================================================================================================================================================
' DIAGNOSTIC / STATUS / POPUP INTEGRATION PROCEDURES
' =============================================================================================================================================================================

Private Sub ShowFinalADCMessage(ByVal messageText As String, ByVal buttons As VbMsgBoxStyle, ByVal TitleText As String, ByVal traceLevel As String, ByVal sourceProcedure As String, Optional ByVal stepName As String = vbNullString, Optional ByVal errNumber As Long = 0, Optional ByVal errSource As String = vbNullString, Optional ByVal errDescription As String = vbNullString, Optional ByVal objectName As String = vbNullString)
    WriteADCDiagnostic "POPUP", TitleText, "Procedure=" & sourceProcedure & "; Step=" & stepName & "; Message=" & messageText
    TraceFinalADCPopup messageText, TitleText, traceLevel, sourceProcedure, stepName, errNumber, errSource, errDescription, objectName
    UtilityMsgBox messageText, buttons, TitleText
End Sub

Private Sub TraceFinalADCPopup(ByVal messageText As String, ByVal TitleText As String, ByVal traceLevel As String, ByVal sourceProcedure As String, ByVal stepName As String, ByVal errNumber As Long, ByVal errSource As String, ByVal errDescription As String, ByVal objectName As String)
    On Error Resume Next
    Dim traceMessage As String
    Dim traceErrorNumber As Long
    Dim traceErrorDescription As String
    traceMessage = Replace$(messageText, vbCrLf, " | ")
    If Len(traceMessage) > 240 Then traceMessage = Left$(traceMessage, 237) & "..."
    Application.Run "LogTraceMessage", traceLevel, "Generation", sourceProcedure, traceMessage, messageText, stepName, errNumber, errSource, errDescription, objectName, "FormSet", "CodeDb", TitleText
    traceErrorNumber = Err.Number
    traceErrorDescription = Err.Description
    Err.Clear
    If traceErrorNumber <> 0 Then Debug.Print MODULE_NAME & ": Final popup could not be copied to Access Development Console trace. Error " & CStr(traceErrorNumber) & ": " & traceErrorDescription
    On Error GoTo 0
End Sub

Private Function ADCDiagnosticLogPath() As String
    On Error GoTo ExitHere
    Dim databaseFullName As String
    Dim separatorPosition As Long
    databaseFullName = CurrentProject.FullName
    separatorPosition = InStrRev(databaseFullName, "\")
    If separatorPosition > 0 Then
        ADCDiagnosticLogPath = Left$(databaseFullName, separatorPosition) & DIAGNOSTIC_LOG_FILE_NAME
    Else
        ADCDiagnosticLogPath = DIAGNOSTIC_LOG_FILE_NAME
    End If
ExitHere:
    If Len(ADCDiagnosticLogPath) = 0 Then ADCDiagnosticLogPath = DIAGNOSTIC_LOG_FILE_NAME
    Err.Clear
End Function

Private Sub ResetADCDiagnosticLog()
    If Not ENABLE_GENERATED_FORM_SET_DIAGNOSTIC_FILE_LOGGING Then Exit Sub
    On Error GoTo ResetFailure
    Dim fileNumber As Integer
    Dim logPath As String
    logPath = ADCDiagnosticLogPath()
    If Len(Dir$(logPath)) > 0 Then Kill logPath
    fileNumber = FreeFile
    Open logPath For Output As #fileNumber
    Close #fileNumber
    Exit Sub
ResetFailure:
    Dim resetErrNumber As Long
    Dim resetErrDescription As String
    resetErrNumber = Err.Number
    resetErrDescription = Err.Description
    On Error Resume Next
    If fileNumber > 0 Then Close #fileNumber
    Debug.Print MODULE_NAME & ": Diagnostic log reset failed. Error " & CStr(resetErrNumber) & ": " & resetErrDescription
    On Error GoTo 0
End Sub

Private Function ADCDiagnosticTimestamp() As String
    Dim timerValue As Double
    Dim milliseconds As Long
    timerValue = CDbl(Timer)
    milliseconds = CLng((timerValue - Fix(timerValue)) * 1000)
    If milliseconds < 0 Then milliseconds = 0
    If milliseconds > 999 Then milliseconds = 999
    ADCDiagnosticTimestamp = Format$(Now, "yyyy-mm-dd hh:nn:ss") & "." & Format$(milliseconds, "000")
End Function

Private Sub WriteADCDiagnostic(ByVal eventType As String, ByVal actionText As String, Optional ByVal detailText As String = vbNullString)
    If Not ENABLE_GENERATED_FORM_SET_DIAGNOSTIC_FILE_LOGGING Then Exit Sub
    On Error GoTo LogFailure
    Dim fileNumber As Integer
    Dim logLine As String
    logLine = ADCDiagnosticTimestamp() & " | " & MODULE_NAME & " | " & eventType & " | " & Replace$(actionText, vbCrLf, " | ")
    If Len(detailText) > 0 Then logLine = logLine & " | " & Replace$(detailText, vbCrLf, " | ")
    fileNumber = FreeFile
    Open ADCDiagnosticLogPath() For Append As #fileNumber
    Print #fileNumber, logLine
    Close #fileNumber
    Exit Sub
LogFailure:
    Dim logErrNumber As Long
    Dim logErrDescription As String
    logErrNumber = Err.Number
    logErrDescription = Err.Description
    On Error Resume Next
    If fileNumber > 0 Then Close #fileNumber
    Debug.Print MODULE_NAME & ": Diagnostic log write failed. Error " & CStr(logErrNumber) & ": " & logErrDescription
    On Error GoTo 0
End Sub

Private Sub SetADCStatus(ByVal actionText As String, Optional ByVal detailText As String = "")
    On Error Resume Next
    WriteADCDiagnostic "ACTION", actionText, detailText
    If Len(detailText) > 0 Then
        SysCmd acSysCmdSetStatus, ApplyDisplayBranding("Access Development Console Form Set Builder: " & actionText & " - " & detailText)
        Debug.Print MODULE_NAME & ": " & actionText & " - " & detailText
    Else
        SysCmd acSysCmdSetStatus, ApplyDisplayBranding("Access Development Console Form Set Builder: " & actionText)
        Debug.Print MODULE_NAME & ": " & actionText
    End If
    On Error GoTo 0
End Sub

Private Sub ClearADCStatus()
    On Error Resume Next
    SysCmd acSysCmdClearStatus
    On Error GoTo 0
End Sub
