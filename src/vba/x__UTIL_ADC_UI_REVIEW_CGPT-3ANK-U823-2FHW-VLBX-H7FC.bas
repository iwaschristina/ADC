Attribute VB_Name = "x__UTIL_ADC_UI_REVIEW"
' ================================================================================================================================================================================
' ChatGPT Rewrite Tracking Header
' ================================================================================================================================================================================
' ChatGPT Previous Version ID: CGPT-98KN-MKDL-YHUP-B3BH-PQAT
' ChatGPT Baseline Version ID: CGPT-98KN-MKDL-YHUP-B3BH-PQAT
' ChatGPT Version ID: CGPT-3ANK-U823-2FHW-VLBX-H7FC
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
' STANDARD MODULE: Access Development Console Review UI Utility
' VB_Name: x__UTIL_ADC_UI_REVIEW
' Purpose: Owns Add-In Review UI construction and callbacks plus form/report reference-review inspection and navigation.
' ================================================================================================================================================================================

Option Compare Database
Option Explicit

' ===============================================================================================================================================================================
' MODULE-LEVEL DECLARATIONS
' ===============================================================================================================================================================================

Public Const ADD_IN_REVIEW_FORM_VERSION As String = "2026-07-20 Add-In Review physical bound-control contract v2"

' Generated-form selection-update suspension state.
Private mAddInReviewSelectionUpdatesSuspended As Boolean

' ===============================================================================================================================================================================
' PROCEDURES
' ===============================================================================================================================================================================

' =============================================================================================================================================================================
' ADD-IN REVIEW FORM CONSTRUCTION / VALIDATION PROCEDURES
' =============================================================================================================================================================================

' Purpose: Requires add in review source fields for source name, field names.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   sourceName (Input; String): Name of the source object/query/table/control from which this operation reads or to which the generated definition is bound.
'   fieldNames (Input; Variant): Array/list of field names whose values/columns the UI or builder should include and order.
'   formName (Input; String): Physical Access form name identifying the form to create, open, configure, validate, or inspect.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub RequireAddInReviewSourceFields(ByVal sourceName As String, ByVal fieldNames As Variant, ByVal formName As String)
    Dim fieldItem As Variant
    For Each fieldItem In fieldNames
        SetupControlPump
        If SharedDao_DatabaseFieldExistsInSource(GetHostDb(), sourceName, CStr(fieldItem)) = False Then
            Err.Raise vbObjectError + 7306, "RequireAddInReviewSourceFields", "A required Add-In Review source field is missing." & vbCrLf & "Form: " & formName & vbCrLf & "Source: " & sourceName & vbCrLf & "Field: " & CStr(fieldItem)
        End If
    Next fieldItem
End Sub

' Purpose: Validates generated add in review child specialized for frm, form name.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   frm (Input; Access.Form): Access.Form instance whose controls or form properties the routine reads, configures, or validates.
'   formName (Input; String): Physical Access form name identifying the form to create, open, configure, validate, or inspect.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub ValidateGeneratedAddInReviewChildSpecialized(ByVal frm As Access.Form, ByVal formName As String)
    Dim childControl As Access.control
    Dim clickText As String
    Dim expectedActions As Collection
    Dim actionExpression As Variant
    Dim actionFound As Boolean
    If frm.allowEdits Or frm.allowAdditions Or frm.allowDeletions Then Err.Raise vbObjectError + 7296, "ValidateGeneratedAddInReviewChildSpecialized", "An Add-Ins review detail form is not read-only." & vbCrLf & "Form: " & formName
    ValidateAddInReviewChildControlSources frm, formName
    Set expectedActions = AddInReviewExpectedActionExpressions(formName)
    For Each actionExpression In expectedActions
        SetupControlPump
        actionFound = False
        For Each childControl In frm.Controls
            If childControl.controlType = acCommandButton Then
                clickText = Trim$(Nz(childControl.OnClick, vbNullString))
                If StrComp(clickText, CStr(actionExpression), vbBinaryCompare) = 0 Then
                    actionFound = True
                    Exit For
                End If
            End If
        Next childControl
        If actionFound = False Then Err.Raise vbObjectError + 7305, "ValidateGeneratedAddInReviewChildSpecialized", "An Add-Ins review detail form is missing a required filter or cross-navigation action." & vbCrLf & "Form: " & formName & vbCrLf & "Expected OnClick: " & CStr(actionExpression)
    Next actionExpression
End Sub

' Purpose: Builds the managed consolidated Add-In Review form from its canonical generated-form contract, hosts the ordered child review forms, saves the design, and reports built/skipped state to the caller.
' Public role: Inter-module public generated-form builder. The UI form-build dispatcher calls this routine to create the managed Add-In Review form and report built/skipped state.
' Parameters:
'   replaceExisting (Input; Boolean): When True, enables or requests existing for this call; when False, the routine follows its normal non-forced/default path.
'   builtForms (Input/output; Collection): Collection containing the built forms items that the routine consumes or updates as part of this operation.
'   skippedForms (Input/output; Collection): Supplies the caller-owned skipped-form collection and can receive a skip-reason entry when Add-In Review form preparation does not proceed.
' Returns: None (Sub). Results are returned through the documented ByRef parameters and/or the routine's database, UI, or module-state side effects.
Public Sub CreateChatGPTAddInReviewForm(ByVal replaceExisting As Boolean, ByRef builtForms As Collection, ByRef skippedForms As Collection)
    On Error GoTo ErrHandler
    Dim frm As Form
    Dim tabCtl As control
    Dim ctl As control
    Dim formName As String
    Dim controlDefinitions As Collection
    Dim callbackExpressions As Collection
    Dim addInReviewChildForms As Variant
    Dim addInReviewChildForm As Variant
    Dim addInReviewChildFormName As String
    Dim addInReviewChildIndex As Long
    formName = ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review")
    If PrepareChatGPTGeneratedFormForBuild(formName, replaceExisting, skippedForms) = False Then Exit Sub
    Set controlDefinitions = GeneratedFormContractControlDefinitions(formName)
    Set callbackExpressions = GeneratedFormContractCallbackExpressions(formName)
    Set frm = CreateBlankFormForDesign(formName)
    ApplyUnboundFormProperties frm, GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review")), DEFAULT_UI_ADD_IN_REVIEW_FORM_WIDTH, DEFAULT_UI_ADD_IN_REVIEW_FORM_HEIGHT
    EnsureHeaderFooterSections formName
    forms(formName).Section(acHeader).Height = 1320
    forms(formName).Section(acDetail).Height = 8500
    forms(formName).OnLoad = GeneratedFormContractOnLoadExpression(formName)
    With SharedFormDesign_AddTitleLabel(formName, acHeader, vbNullString, GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review")), DEFAULT_UI_FORM_LEFT_MARGIN, DEFAULT_UI_FORM_TOP_MARGIN, 6500, 420)
        .TextAlign = 1
    End With
    With SharedFormDesign_AddPlainLabel(formName, acHeader, vbNullString, "Read-only Reference Inventory and Loaded VBA Analysis for already loaded and accessible VBA projects. This interface never loads, connects, registers, or modifies an add-in or external project.", DEFAULT_UI_FORM_LEFT_MARGIN, 520, 13800, 390, DEFAULT_UI_FONT_SIZE)
        .TextAlign = 1
    End With
    Set ctl = ChatGPTHelpCreateTextBox(formName, GeneratedFormContractControlName(formName, 1, controlDefinitions), acHeader, vbNullString, DEFAULT_UI_FORM_LEFT_MARGIN, 950, 1500, 300, True, False)
    ctl.Visible = False
    With SharedFormDesign_AddPlainLabel(formName, acHeader, vbNullString, "Selected project", DEFAULT_UI_FORM_LEFT_MARGIN, 950, 1250, 260, DEFAULT_UI_FONT_SIZE)
        .TextAlign = 1
    End With
    Call ChatGPTHelpCreateTextBox(formName, GeneratedFormContractControlName(formName, 2, controlDefinitions), acHeader, vbNullString, 1500, 920, 2600, 330, True, False)
    Call ChatGPTHelpCreateTextBox(formName, GeneratedFormContractControlName(formName, 3, controlDefinitions), acHeader, vbNullString, 4200, 920, 3050, 330, True, False)
    SharedFormDesign_AddButton formName, acHeader, vbNullString, "Refresh Display", GeneratedFormContractCallbackExpression(formName, 1, callbackExpressions), 7460, 900, 1500, 360, "Requery the displayed records without running either scan."
    SharedFormDesign_AddButton formName, acHeader, vbNullString, "Inventory Scan", GeneratedFormContractCallbackExpression(formName, 2, callbackExpressions), 9020, 900, 1300, 360, "Run the existing passive loaded-project, VBE add-in, and COM add-in inventory refresh."
    SharedFormDesign_AddButton formName, acHeader, vbNullString, "Loaded VBA Analysis Scan", GeneratedFormContractCallbackExpression(formName, 3, callbackExpressions), 10380, 900, 1300, 360, "Analyze source evidence for all already loaded and inspectable VBA projects; ProjectRole remains classification metadata."
    SharedFormDesign_AddButton formName, acHeader, vbNullString, "Raw Data", GeneratedFormContractCallbackExpression(formName, 4, callbackExpressions), 11740, 900, 1200, 360, "Open the underlying read-only table or query for the selected review page."
    SharedFormDesign_AddButton formName, acHeader, vbNullString, "Clear Filters", GeneratedFormContractCallbackExpression(formName, 5, callbackExpressions), 12980, 900, 1300, 360, "Clear filters across the Add-Ins Review and restore selected-project links for the External-only Loaded VBA Analysis detail pages."
    Set tabCtl = SharedFormDesign_AddTabControl(formName, GeneratedFormContractControlName(formName, 4, controlDefinitions), DEFAULT_UI_TAB_LEFT, DEFAULT_UI_TAB_TOP, DEFAULT_UI_ADD_IN_REVIEW_TAB_WIDTH, DEFAULT_UI_TAB_HEIGHT + 400)
    addInReviewChildForms = AddInReviewChildFormNames()
    addInReviewChildIndex = 0
    For Each addInReviewChildForm In addInReviewChildForms
        addInReviewChildFormName = CStr(addInReviewChildForm)
        AddAddInReviewTabPage formName, tabCtl, addInReviewChildIndex, AddInReviewChildHostPageName(addInReviewChildFormName), AddInReviewChildHostPageCaption(addInReviewChildFormName), addInReviewChildFormName, AddInReviewChildHostPageDescription(addInReviewChildFormName), AddInReviewChildLinksSelectedProject(addInReviewChildFormName)
        addInReviewChildIndex = addInReviewChildIndex + 1
    Next addInReviewChildForm
    SaveAndCloseDesignForm formName
    builtForms.Add formName
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    SharedForm_CloseIfOpen formName, acSaveNo
    On Error GoTo 0
    Err.Raise savedErrNumber, "CreateChatGPTAddInReviewForm", "Could not create the consolidated Add-Ins review form." & vbCrLf & "Form: " & formName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub AddAddInReviewTabPage(ByVal hostFormName As String, ByVal tabCtl As control, ByVal pageIndex As Long, ByVal pageName As String, ByVal captionText As String, ByVal sourceFormName As String, ByVal descriptionText As String, ByVal linkToSelectedProject As Boolean)
    Dim pg As Page
    Dim subCtl As control
    Set pg = ConfigureTabPage(tabCtl, pageIndex, pageName, captionText)
    With SharedFormDesign_AddPlainLabel(hostFormName, acDetail, pg.Name, descriptionText, DEFAULT_UI_TAB_LEFT + DEFAULT_UI_FORM_LEFT_MARGIN, DEFAULT_UI_TAB_TOP + 500, tabCtl.Width - (2 * DEFAULT_UI_FORM_LEFT_MARGIN), 300, DEFAULT_UI_FONT_SIZE)
        .TextAlign = 1
    End With
    If SharedDao_AccessObjectExists(GetHostDb(), acForm, sourceFormName) = False Then Err.Raise vbObjectError + 7290, "AddAddInReviewTabPage", "A required Add-In Review source form is missing." & vbCrLf & "Form: " & sourceFormName
    SharedFormDesign_AddSubformControl hostFormName, acDetail, pg.Name, "sub_" & pageName, sourceFormName, DEFAULT_UI_TAB_LEFT + DEFAULT_UI_FORM_LEFT_MARGIN, DEFAULT_UI_TAB_TOP + 500 + 300 + DEFAULT_UI_SECTION_TITLE_CONTENT_GAP, tabCtl.Width - (2 * DEFAULT_UI_FORM_LEFT_MARGIN), DEFAULT_UI_TAB_HEIGHT - 300
    Set subCtl = forms(hostFormName).Controls("sub_" & pageName)
    If linkToSelectedProject Then
        subCtl.LinkMasterFields = GeneratedFormContractControlName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review"), 1)
        subCtl.LinkChildFields = FIELD_PROJECT_IDENTITY
    End If
End Sub

' Purpose: Validates add in review child control sources for frm, form name.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   frm (Input; Access.Form): Access.Form instance whose controls or form properties the routine reads, configures, or validates.
'   formName (Input; String): Physical Access form name identifying the form to create, open, configure, validate, or inspect.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub ValidateAddInReviewChildControlSources(ByVal frm As Access.Form, ByVal formName As String)
    On Error GoTo ErrHandler
    Dim sourceName As String
    Dim definitions As Collection
    Dim Definition As Variant
    Dim expectedSources As Object
    Dim actualCounts As Object
    Dim ctl As Access.control
    Dim fieldName As String
    Dim controlSourceText As String
    Dim actualBoundControlCount As Long
    Dim expectedCount As Long
    Dim validationStage As String
    sourceName = AddInReviewChildSourceName(formName)
    Set definitions = ContinuousListControlDefinitions(formName, sourceName, AddInReviewExpectedControlSources(formName), True)
    ' INTENTIONAL KEYED LOOKUP: expectedSources, actualCounts provides keyed membership/lookup over source rows or state already read by ValidateAddInReviewChildControlSources; it is a derived access path, not a
    ' second authority.
    Set expectedSources = CreateObject("Scripting.Dictionary")
    expectedSources.CompareMode = vbTextCompare
    Set actualCounts = CreateObject("Scripting.Dictionary")
    actualCounts.CompareMode = vbTextCompare
    For Each Definition In definitions
        fieldName = CStr(Definition(0))
        expectedSources.Add fieldName, True
    Next Definition
    validationStage = "Enumerate saved bound controls"
    For Each ctl In frm.Controls
        SetupControlPump
        Select Case ctl.controlType
            Case acTextBox, acCheckBox, acComboBox, acListBox
                controlSourceText = Trim$(Nz(ctl.controlSource, vbNullString))
                If Len(controlSourceText) = 0 Then Err.Raise vbObjectError + 7400, "ValidateAddInReviewChildControlSources", "An Add-In Review data control has a blank ControlSource." & vbCrLf & "Form: " & formName & vbCrLf & "Source: " & sourceName & vbCrLf & "Control: " & ctl.Name
                If expectedSources.Exists(controlSourceText) = False Then Err.Raise vbObjectError + 7401, "ValidateAddInReviewChildControlSources", "An Add-In Review detail form contains an unexpected bound ControlSource." & vbCrLf & "Form: " & formName & vbCrLf & "Source: " & sourceName & vbCrLf & "Control: " & ctl.Name & vbCrLf & "ControlSource: " & controlSourceText
                If CLng(ctl.Section) <> CLng(acDetail) Then Err.Raise vbObjectError + 7402, "ValidateAddInReviewChildControlSources", "A required Add-In Review bound control is not in the detail section." & vbCrLf & "Form: " & formName & vbCrLf & "Source: " & sourceName & vbCrLf & "Control: " & ctl.Name & vbCrLf & "ControlSource: " & controlSourceText & vbCrLf & "Section: " & CStr(ctl.Section)
                If actualCounts.Exists(controlSourceText) Then
                    actualCounts(controlSourceText) = CLng(actualCounts(controlSourceText)) + 1
                Else
                    actualCounts.Add controlSourceText, 1
                End If
                actualBoundControlCount = actualBoundControlCount + 1
        End Select
    Next ctl
    expectedCount = definitions.count
    For Each Definition In definitions
        SetupControlPump
        fieldName = CStr(Definition(0))
        validationStage = "Validate required ControlSource: " & fieldName
        SetChatGPTBuildContext "Strictly validating the Add-In Review interface.", formName, "Validate Add-In Review required bound controls.", "Source: " & sourceName & "; Required ControlSource: " & fieldName, sourceName
        If actualCounts.Exists(fieldName) = False Then Err.Raise vbObjectError + 7307, "ValidateAddInReviewChildControlSources", "An Add-In Review detail form is missing a required bound control." & vbCrLf & "Form: " & formName & vbCrLf & "Source: " & sourceName & vbCrLf & "ControlSource: " & fieldName & vbCrLf & "Validation stage: " & validationStage
        If CLng(actualCounts(fieldName)) <> 1 Then Err.Raise vbObjectError + 7403, "ValidateAddInReviewChildControlSources", "An Add-In Review required ControlSource is bound more than once." & vbCrLf & "Form: " & formName & vbCrLf & "Source: " & sourceName & vbCrLf & "ControlSource: " & fieldName & vbCrLf & "Bound controls: " & CStr(actualCounts(fieldName))
    Next Definition
    If actualBoundControlCount <> expectedCount Then Err.Raise vbObjectError + 7404, "ValidateAddInReviewChildControlSources", "The saved Add-In Review bound-control count does not match the authoritative contract." & vbCrLf & "Form: " & formName & vbCrLf & "Source: " & sourceName & vbCrLf & "Expected bound controls: " & CStr(expectedCount) & vbCrLf & "Actual bound controls: " & CStr(actualBoundControlCount)
    If StrComp(formName, ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review_Inventory"), vbTextCompare) = 0 Then
        If actualCounts.Exists(FIELD_SOURCE_KIND) = False Then Err.Raise vbObjectError + 7405, "ValidateAddInReviewChildControlSources", "The Inventory form does not contain a bound SourceKind control." & vbCrLf & "Form: " & formName & vbCrLf & "Source: " & sourceName
        If CLng(actualCounts(FIELD_SOURCE_KIND)) <> 1 Then Err.Raise vbObjectError + 7405, "ValidateAddInReviewChildControlSources", "The Inventory form does not contain exactly one bound SourceKind control." & vbCrLf & "Form: " & formName & vbCrLf & "Source: " & sourceName & vbCrLf & "Bound controls: " & CStr(actualCounts(FIELD_SOURCE_KIND))
    End If
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ValidateAddInReviewChildControlSources", "Could not validate the exact Add-In Review physical bound-control contract." & vbCrLf & "Form: " & formName & vbCrLf & "Source: " & sourceName & vbCrLf & "Validation stage: " & validationStage & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub RequeryAddInReviewAnalysisForms()
    Dim controlNames As Variant
    Dim controlItem As Variant
    controlNames = Array("sub_pgAddInReviewReferences", "sub_pgAddInReviewModules", "sub_pgAddInReviewProcedures", "sub_pgAddInReviewConflicts", "sub_pgAddInReviewDependencies")
    For Each controlItem In controlNames
        SetupControlPump
        forms(TargetManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review")).Controls(CStr(controlItem)).Form.FilterOn = False
        forms(TargetManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review")).Controls(CStr(controlItem)).Form.Requery
    Next controlItem
End Sub

' =============================================================================================================================================================================
' FORM / REPORT REFERENCE INSPECTION / NAVIGATION PROCEDURES
' =============================================================================================================================================================================

' Purpose: Resolves unwrap object reference from the supplied/current ADC state, using `StrComp` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   valueText (Input; String): Input value text for unwrap object reference; is tested by If StrComp(Left$(valueText, 5), "Form.", vbTextCompare) = 0; is tested by If StrComp(Left$(valueText, 7), "Report.",
'     vbTextCompare) = 0.
' Returns: String: Returns the unwrap object reference text value. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function UnwrapObjectReference(ByVal valueText As String) As String
    'Removes Form./Report. prefixes and matching square brackets from an Access object-reference string.
    valueText = Trim$(valueText)
    If StrComp(Left$(valueText, 5), "Form.", vbTextCompare) = 0 Then valueText = Mid$(valueText, 6)
    If StrComp(Left$(valueText, 7), "Report.", vbTextCompare) = 0 Then valueText = Mid$(valueText, 8)
    If Left$(valueText, 1) = "[" And Right$(valueText, 1) = "]" Then valueText = Mid$(valueText, 2, Len(valueText) - 2)
    UnwrapObjectReference = Trim$(valueText)
End Function

' Purpose: Finds broken form and report references.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub FindBrokenFormAndReportReferences(Optional ByVal showMessage As Boolean = True)
    'Read-only scanner for applicable non-generated Forms and Reports. Objects opened only for inspection are always closed with acSaveNo.
    On Error GoTo ErrHandler
    Dim findings As Collection
    Dim formObject As accessObject
    Dim reportObject As accessObject
    Dim summaryText As String
    Dim reviewedFormCount As Long
    Dim excludedFormCount As Long
    Dim reviewedReportCount As Long
    Dim excludedReportCount As Long
    Dim propertyCount As Long
    Dim traceRunOwnedByProcedure As Boolean
    traceRunOwnedByProcedure = BeginUITraceRunIfNeeded("UI.FormReportReferenceReview")
    If CodeDbIsHostDb() = False Then Err.Raise vbObjectError + 6698, "FindBrokenFormAndReportReferences", "Form and Report reference inspection currently requires CodeDb and HostDb to be the same active database."
    Set findings = New Collection
    PrintUIStatus "FindBrokenFormAndReportReferences", "Form and Report reference inspection started."
    For Each formObject In CurrentProject.AllForms
        SetupControlPump
        If FormExcludedFromReferenceReview(formObject.Name) Then
            excludedFormCount = excludedFormCount + 1
        Else
            reviewedFormCount = reviewedFormCount + 1
            PrintUIStatus "FindBrokenFormAndReportReferences", "Inspecting Form " & CStr(reviewedFormCount) & ": " & formObject.Name, formObject.Name, "FormReportReferenceReview"
            ReviewOneFormReferences formObject.Name, findings, propertyCount
        End If
    Next formObject
    For Each reportObject In CurrentProject.AllReports
        SetupControlPump
        If ReportExcludedFromReferenceReview(reportObject.Name) Then
            excludedReportCount = excludedReportCount + 1
        Else
            reviewedReportCount = reviewedReportCount + 1
            PrintUIStatus "FindBrokenFormAndReportReferences", "Inspecting Report " & CStr(reviewedReportCount) & ": " & reportObject.Name, reportObject.Name, "FormReportReferenceReview"
            ReviewOneReportReferences reportObject.Name, findings, propertyCount
        End If
    Next reportObject
    summaryText = "Broken Form and Report reference inspection complete." & vbCrLf & vbCrLf & _
        "Forms reviewed: " & CStr(reviewedFormCount) & vbCrLf & "Forms excluded: " & CStr(excludedFormCount) & vbCrLf & _
        "Reports reviewed: " & CStr(reviewedReportCount) & vbCrLf & "Reports excluded: " & CStr(excludedReportCount) & vbCrLf & _
        "Properties checked: " & CStr(propertyCount) & vbCrLf & "Findings: " & CStr(findings.count)
    If findings.count > 0 Then summaryText = summaryText & vbCrLf & vbCrLf & SharedCollection_Join(findings, vbCrLf)
    LogTraceMessage IIf(findings.count > 0, "Warning", "Info"), "FormReportReferenceReview", "FindBrokenFormAndReportReferences", "Form and Report reference inspection completed.", summaryText, vbNullString, 0, vbNullString, vbNullString, vbNullString, "FormReportSet", "CodeDb", "FormReportReferenceReview"
    PrintUIStatus "FindBrokenFormAndReportReferences", "Form and Report reference inspection complete. Forms=" & CStr(reviewedFormCount) & "; Reports=" & CStr(reviewedReportCount) & "; findings=" & CStr(findings.count) & "."
    If showMessage Then CoreShowMessage summaryText, IIf(findings.count > 0, vbExclamation, vbInformation), "Find Broken Form and Report References", "ShowADCMessage", "UI"
    CompleteUITraceRunIfOwned traceRunOwnedByProcedure, "Form and Report reference inspection complete."
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    LogTraceMessage "Error", "FormReportReferenceReview", "FindBrokenFormAndReportReferences", "Form and Report reference inspection failed.", savedErrDescription, vbNullString, savedErrNumber, savedErrSource, savedErrDescription, vbNullString, "FormReportSet", "CodeDb", "FormReportReferenceReview"
    If showMessage Then CoreShowErrorMessage "Could not inspect Form and Report references." & vbCrLf & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description:" & vbCrLf & savedErrDescription, vbCritical, "Find Broken Form and Report References", "FindBrokenFormAndReportReferences", "UI"
    CompleteUITraceRunIfOwned traceRunOwnedByProcedure, "Form and Report reference inspection failed.", False
    If showMessage Then Exit Sub
    Err.Raise savedErrNumber, "FindBrokenFormAndReportReferences", savedErrDescription
End Sub

Private Function FormExcludedFromReferenceReview(ByVal formName As String) As Boolean
    On Error GoTo ExitHere
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim definitionData As Variant
    If Len(formName) = 0 Then FormExcludedFromReferenceReview = True: Exit Function
    For Each definitionData In CanonicalGeneratedFormNamingDefinitions()
        If StrComp(formName, ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CStr(definitionData(GENERATED_FORM_NAMING_LOGICAL_ID))), vbTextCompare) = 0 Then FormExcludedFromReferenceReview = True: Exit Function
    Next definitionData
    If Left$(formName, 1) = "~" Then FormExcludedFromReferenceReview = True: Exit Function
    If CodeDbTableExists(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_FormDefinitions")) = False Then Exit Function
    If SharedDao_TableFieldExists(GetHostDb(), ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_FormDefinitions"), FIELD_FORM_OBJECT_NAME) = False Then Exit Function
    Set db = GetHostDb()
    sqlText = "SELECT TOP 1 " & SQL_FIELD_FORM_OBJECT_NAME & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_FormDefinitions") & "] WHERE " & SQL_FIELD_FORM_OBJECT_NAME & "=" & SharedDao_SqlQuote(formName)
    If SharedDao_TableFieldExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_FormDefinitions"), FIELD_CATALOG_OWNER) Then sqlText = sqlText & " AND (UCase(Trim(Nz(" & SQL_FIELD_CATALOG_OWNER & ",'')))='GENERATED'" Else sqlText = sqlText & " AND (False"
    If SharedDao_TableFieldExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_FormDefinitions"), FIELD_IS_GENERATED_BY_UTILITY) Then sqlText = sqlText & " OR Nz(" & SQL_FIELD_IS_GENERATED_BY_UTILITY & ",False)<>False"
    sqlText = sqlText & ");"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    FormExcludedFromReferenceReview = Not rs.EOF
ExitHere:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
End Function

Private Function ReportExcludedFromReferenceReview(ByVal reportName As String) As Boolean
    Dim definitionData As Variant
    If Len(reportName) = 0 Then ReportExcludedFromReferenceReview = True: Exit Function
    If Left$(reportName, 1) = "~" Then ReportExcludedFromReferenceReview = True: Exit Function
    For Each definitionData In CanonicalReportNamingDefinitions()
        If StrComp(reportName, ManagedObjectPhysicalName(NAMING_CATEGORY_REPORT, CStr(definitionData(0))), vbTextCompare) = 0 Then ReportExcludedFromReferenceReview = True: Exit Function
    Next definitionData
End Function

Private Sub ReviewOneFormReferences(ByVal formName As String, ByRef findings As Collection, ByRef propertyCount As Long)
    On Error GoTo ErrHandler
    Dim frm As Access.Form
    Dim openedForReview As Boolean
    Dim ctl As Access.control
    If SharedAccess_ObjectIsOpen(acForm, formName) Then
        Set frm = forms(formName)
    Else
        DoCmd.OpenForm formName, acDesign, , , , acHidden
        openedForReview = True
        Set frm = forms(formName)
    End If
    ReviewObjectPropertyReference formName, "Form", "RecordSource", CStr(Nz(frm.recordSource, vbNullString)), findings, propertyCount
    ReviewEventReferenceProperties formName, "Form", frm, findings, propertyCount
    For Each ctl In frm.Controls
        SetupControlPump
        ReviewControlReferenceProperties formName, ctl, findings, propertyCount
    Next ctl
    If openedForReview Then DoCmd.Close acForm, formName, acSaveNo
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    On Error Resume Next
    If openedForReview And SharedAccess_ObjectIsOpen(acForm, formName) Then DoCmd.Close acForm, formName, acSaveNo
    On Error GoTo 0
    findings.Add "Form [" & formName & "] | Review error " & CStr(savedErrNumber) & ": " & savedErrDescription
End Sub

Private Sub ReviewOneReportReferences(ByVal reportName As String, ByRef findings As Collection, ByRef propertyCount As Long)
    On Error GoTo ErrHandler
    Dim rpt As Access.Report
    Dim openedForReview As Boolean
    Dim ctl As Access.control
    If SharedAccess_ObjectIsOpen(acReport, reportName) Then
        Set rpt = reports(reportName)
    Else
        DoCmd.OpenReport reportName, acViewDesign, , , acHidden
        openedForReview = True
        Set rpt = reports(reportName)
    End If
    ReviewObjectPropertyReference reportName, "Report", "RecordSource", CStr(Nz(rpt.recordSource, vbNullString)), findings, propertyCount
    ReviewEventReferenceProperties reportName, "Report", rpt, findings, propertyCount
    For Each ctl In rpt.Controls
        SetupControlPump
        ReviewControlReferenceProperties reportName, ctl, findings, propertyCount
    Next ctl
    If openedForReview Then DoCmd.Close acReport, reportName, acSaveNo
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    On Error Resume Next
    If openedForReview And SharedAccess_ObjectIsOpen(acReport, reportName) Then DoCmd.Close acReport, reportName, acSaveNo
    On Error GoTo 0
    findings.Add "Report [" & reportName & "] | Review error " & CStr(savedErrNumber) & ": " & savedErrDescription
End Sub

Private Sub ReviewControlReferenceProperties(ByVal objectName As String, ByVal ctl As Access.control, ByRef findings As Collection, ByRef propertyCount As Long)
    On Error Resume Next
    ReviewObjectPropertyReference objectName, ctl.Name, "ControlSource", ctl.controlSource, findings, propertyCount
    ReviewObjectPropertyReference objectName, ctl.Name, "RowSource", ctl.rowSource, findings, propertyCount
    ReviewObjectPropertyReference objectName, ctl.Name, "SourceObject", ctl.sourceObject, findings, propertyCount
    ReviewEventReferenceProperties objectName, ctl.Name, ctl, findings, propertyCount
    On Error GoTo 0
End Sub

Private Sub ReviewEventReferenceProperties(ByVal objectName As String, ByVal ownerName As String, ByVal targetObject As Object, ByRef findings As Collection, ByRef propertyCount As Long)
    'Reviews every available On* event property for references to missing current macro objects.
    On Error Resume Next
    Dim propertyItem As Object
    Dim propertyValue As String
    For Each propertyItem In targetObject.properties
        SetupControlPump
        If Left$(CStr(propertyItem.Name), 2) = "On" Then
            Err.Clear
            propertyValue = CStr(Nz(propertyItem.value, vbNullString))
            If Err.Number = 0 Then ReviewObjectPropertyReference objectName, ownerName, CStr(propertyItem.Name), propertyValue, findings, propertyCount
        End If
    Next propertyItem
    On Error GoTo 0
End Sub

Private Sub ReviewObjectPropertyReference(ByVal objectName As String, ByVal ownerName As String, ByVal propertyName As String, ByVal propertyValue As String, ByRef findings As Collection, ByRef propertyCount As Long)
    propertyCount = propertyCount + 1
    propertyValue = Trim$(propertyValue)
    If Len(propertyValue) = 0 Then Exit Sub
    ReviewUnknownSimpleReference objectName, ownerName, propertyName, propertyValue, findings
End Sub

Private Sub ReviewUnknownSimpleReference(ByVal objectName As String, ByVal ownerName As String, ByVal propertyName As String, ByVal propertyValue As String, ByRef findings As Collection)
    Dim targetName As String
    Dim targetType As AcObjectType
    If StrComp(propertyName, "RecordSource", vbTextCompare) = 0 Or StrComp(propertyName, "RowSource", vbTextCompare) = 0 Then
        If LooksLikeSql(propertyValue) Then Exit Sub
        If StrComp(propertyName, "RowSource", vbTextCompare) = 0 And InStr(1, propertyValue, ";", vbBinaryCompare) > 0 Then Exit Sub
        targetName = UnwrapObjectReference(propertyValue)
        If Len(targetName) > 0 Then
            If SharedDao_AccessObjectExists(GetHostDb(), acTable, targetName) = False And SharedDao_AccessObjectExists(GetHostDb(), acQuery, targetName) = False Then findings.Add objectName & " | " & ownerName & "." & propertyName & " references missing table/query [" & targetName & "]."
        End If
    ElseIf StrComp(propertyName, "SourceObject", vbTextCompare) = 0 Then
        targetName = UnwrapObjectReference(propertyValue)
        If InStr(1, propertyValue, "Form.", vbTextCompare) = 1 Then
            targetType = acForm
        ElseIf InStr(1, propertyValue, "Report.", vbTextCompare) = 1 Then
            targetType = acReport
        Else
            Exit Sub
        End If
        If SharedDao_AccessObjectExists(GetHostDb(), targetType, targetName) = False Then findings.Add objectName & " | " & ownerName & "." & propertyName & " references missing object [" & propertyValue & "]."
    ElseIf Left$(propertyName, 2) = "On" Then
        If StrComp(propertyValue, "[Event Procedure]", vbTextCompare) = 0 Or StrComp(propertyValue, "[Embedded Macro]", vbTextCompare) = 0 Or Left$(propertyValue, 1) = "=" Then Exit Sub
        targetName = UnwrapObjectReference(propertyValue)
        If Len(targetName) > 0 Then
            If SharedDao_AccessObjectExists(GetHostDb(), acMacro, targetName) Then Exit Sub
            If InStr(1, targetName, ".", vbBinaryCompare) > 0 Then
                If SharedDao_AccessObjectExists(GetHostDb(), acMacro, Left$(targetName, InStr(1, targetName, ".", vbBinaryCompare) - 1)) Then Exit Sub
            End If
            findings.Add objectName & " | " & ownerName & "." & propertyName & " references missing macro [" & targetName & "]."
        End If
    End If
End Sub


' Purpose: Executes the chat gpt open reference location from current row managed-form/UI callback using the routine's existing workflow and reports results through its established return/side-effect
'   contract.
' Public role: Form-code/event-expression entry point. Generated or managed Access forms can invoke this Public routine from form code or an event expression.
' Parameters: None.
' Returns: Boolean: True when the chat gpt open reference location from current row condition holds under the routine's existing validation rules; False when it does not.
Public Function ChatGPT_OpenReferenceLocationFromCurrentRow() As Boolean
    ' Wrapper purpose: Generated review callback that reads the current reference row and opens or navigates to its stored source location.
    ' Why retained: Retained because the generated control expression needs a stable public Function while row interpretation and source navigation remain UI_REVIEW-owned.
    On Error GoTo ErrHandler
    Dim frm As Access.Form
    Dim pathText As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Set frm = Screen.ActiveControl.Parent
    If frm Is Nothing Then Err.Raise vbObjectError + 6743, "ChatGPT_OpenReferenceLocationFromCurrentRow", "The active generated reference form could not be resolved."
    If frm.Dirty Then frm.Dirty = False
    pathText = CurrentFormRecordText(frm, FIELD_FULL_PATH)
    If Len(Trim$(pathText)) = 0 Then pathText = CurrentFormRecordText(frm, FIELD_REGISTERED_PATH)
    If Len(Trim$(pathText)) = 0 Then Err.Raise vbObjectError + 6744, "ChatGPT_OpenReferenceLocationFromCurrentRow", "The current reference row does not contain a file path."
    Call OpenReferenceFileLocation(pathText, True)
    ChatGPT_OpenReferenceLocationFromCurrentRow = True
    Exit Function
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    LogTraceMessage "Error", "UI", "ChatGPT_OpenReferenceLocationFromCurrentRow", "A generated reference form could not open the selected file location.", "Path: " & pathText & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, "Open reference location", savedErrNumber, savedErrSource, savedErrDescription, pathText, "File", "CodeDb", "GeneratedFormCommand"
    CoreShowErrorMessage "Could not open the selected reference file location." & vbCrLf & vbCrLf & savedErrDescription, vbCritical, DEFAULT_UI_MESSAGE_TITLE, "ChatGPT_OpenReferenceLocationFromCurrentRow", "UI"
End Function

' =============================================================================================================================================================================
' ADD-IN REVIEW GENERATED-FORM CALLBACK / FILTER PROCEDURES
' =============================================================================================================================================================================

' Purpose: Initializes the Add-In Review form state, selections, filters, and detail display from the current review inventory.
' Public role: Form-code/event-expression entry point. Generated or managed Access forms can invoke this Public routine from form code or an event expression.
' Parameters: None.
' Returns: Boolean: True when chat gpt add in review initialize completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or not-applicable path.
Public Function ChatGPT_AddInReviewInitialize() As Boolean
    ' Wrapper purpose: Generated Add-In Review callback that initializes the review surface and its selection/display state.
    ' Why retained: Retained because the generated form event needs a stable public Function while initialization details and review state remain private.
    On Error GoTo ErrHandler
    If AddInReviewHostIsLoaded() = False Then Exit Function
    With forms(TargetManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review"))
        .Controls("txtSelectedProjectIdentity").value = Null
        .Controls("txtSelectedProjectName").value = "(none)"
        .Controls("txtSelectionStatus").value = "Select a loaded project for review. External projects populate the analysis detail tabs; all loaded-project roles remain available in the summary/project views."
    End With
    RestoreAddInReviewProjectLinks
    If ChatGPT_AddInReviewRefreshDisplay() = False Then Exit Function
    ChatGPT_AddInReviewInitialize = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    LogTraceMessage "Error", "UI", "ChatGPT_AddInReviewInitialize", "Could not initialize the consolidated Add-Ins review interface.", "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, "Initialize Add-Ins review", savedErrNumber, savedErrSource, savedErrDescription, TargetManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review"), "Form", "CodeDb", "AddInReview"
    CoreShowErrorMessage "Could not initialize the consolidated Add-Ins review interface." & vbCrLf & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, vbCritical, DEFAULT_UI_MESSAGE_TITLE, "ChatGPT_AddInReviewInitialize", "UI"
End Function

' Purpose: Applies the current Add-In Review selection and refreshes the dependent detail/display state.
' Public role: Form-code/event-expression entry point. Generated or managed Access forms can invoke this Public routine from form code or an event expression.
' Parameters:
'   projectIdentity (Input; Variant): Stable project identity used to select or correlate the exact ADC definition/state handled by this procedure.
'   displayName (Input; Variant): User-facing display name associated with the selected review/project/object item; Variant permits the Access Null/Empty state used by the UI.
'   statusValue (Input Optional; Variant): Status value to persist/report for the current operation or managed item; omission leaves the routine’s normal status derivation in effect.
' Returns: Boolean: True when chat gpt add in review refresh display completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or not-applicable
'   path.
Public Function ChatGPT_AddInReviewSelect(ByVal projectIdentity As Variant, ByVal displayName As Variant, Optional ByVal statusValue As Variant) As Boolean
    ' Wrapper purpose: Generated Add-In Review callback that accepts the selected project identity/display values and applies them to review state.
    ' Why retained: Retained because generated controls need a stable public Function contract while selection-state ownership and update rules remain private.
    On Error GoTo ErrHandler
    Dim identityText As String
    Dim nameText As String
    Dim statusText As String
    If AddInReviewHostIsLoaded() = False Then
        ChatGPT_AddInReviewSelect = True
        Exit Function
    End If
    If mAddInReviewSelectionUpdatesSuspended Then
        ChatGPT_AddInReviewSelect = True
        Exit Function
    End If
    identityText = Trim$(CStr(Nz(projectIdentity, vbNullString)))
    nameText = Trim$(CStr(Nz(displayName, vbNullString)))
    statusText = Trim$(CStr(Nz(statusValue, vbNullString)))
    With forms(TargetManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review"))
        If Len(identityText) = 0 Then
            .Controls("txtSelectedProjectIdentity").value = Null
            .Controls("txtSelectedProjectName").value = IIf(Len(nameText) = 0, "(no loaded project)", nameText)
            .Controls("txtSelectionStatus").value = "Loaded VBA Analysis unavailable: no reliable loaded-project relationship is stored for this selection."
        Else
            .Controls("txtSelectedProjectIdentity").value = identityText
            .Controls("txtSelectedProjectName").value = IIf(Len(nameText) = 0, identityText, nameText)
            If Len(statusText) = 0 Then statusText = "Loaded project selected."
            .Controls("txtSelectionStatus").value = statusText
        End If
    End With
    RestoreAddInReviewProjectLinks
    RequeryAddInReviewAnalysisForms
    ChatGPT_AddInReviewSelect = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    LogTraceMessage "Error", "UI", "ChatGPT_AddInReviewSelect", "Could not change the consolidated Add-Ins review selection.", "Project identity: " & UIVariantForMessage(projectIdentity) & vbCrLf & "Display name: " & UIVariantForMessage(displayName) & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, "Select Add-Ins review project", savedErrNumber, savedErrSource, savedErrDescription, TargetManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review"), "Form", "CodeDb", "AddInReview"
    CoreShowErrorMessage "Could not change the consolidated Add-Ins review selection." & vbCrLf & vbCrLf & "Project identity: " & UIVariantForMessage(projectIdentity) & vbCrLf & "Display name: " & UIVariantForMessage(displayName) & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, vbCritical, DEFAULT_UI_MESSAGE_TITLE, "ChatGPT_AddInReviewSelect", "UI"
End Function

' Purpose: Rebuilds the Add-In Review display from the current inventory, selection, and filter state.
' Public role: Form-code/event-expression entry point. Generated or managed Access forms can invoke this Public routine from form code or an event expression.
' Parameters: None.
' Returns: Boolean: True when chat gpt add in review open raw data completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or not-applicable path.
Public Function ChatGPT_AddInReviewRefreshDisplay() As Boolean
    ' Wrapper purpose: Generated Add-In Review callback that refreshes child surfaces while preserving and restoring the current project selection.
    ' Why retained: Retained because the generated Refresh control needs a stable public Function while child-form refresh and selection restoration remain private.
    On Error GoTo ErrHandler
    Dim controlNames As Variant
    Dim controlItem As Variant
    Dim selectedProjectIdentity As Variant
    Dim selectedProjectName As Variant
    Dim selectionStatus As Variant
    If AddInReviewHostIsLoaded() = False Then Exit Function
    With forms(TargetManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review"))
        selectedProjectIdentity = .Controls("txtSelectedProjectIdentity").value
        selectedProjectName = .Controls("txtSelectedProjectName").value
        selectionStatus = .Controls("txtSelectionStatus").value
    End With
    mAddInReviewSelectionUpdatesSuspended = True
    RestoreAddInReviewProjectLinks
    controlNames = Array("sub_pgAddInReviewInventory", "sub_pgAddInReviewVBE", "sub_pgAddInReviewCOM", "sub_pgAddInReviewProjects", "sub_pgAddInReviewReferences", "sub_pgAddInReviewModules", "sub_pgAddInReviewProcedures", "sub_pgAddInReviewConflicts", "sub_pgAddInReviewDependencies")
    For Each controlItem In controlNames
        SetupControlPump
        forms(TargetManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review")).Controls(CStr(controlItem)).Form.Requery
    Next controlItem
    With forms(TargetManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review"))
        .Controls("txtSelectedProjectIdentity").value = selectedProjectIdentity
        .Controls("txtSelectedProjectName").value = selectedProjectName
        .Controls("txtSelectionStatus").value = selectionStatus
    End With
    mAddInReviewSelectionUpdatesSuspended = False
    ChatGPT_AddInReviewRefreshDisplay = True
    Exit Function
ErrHandler:
    mAddInReviewSelectionUpdatesSuspended = False
    CoreShowErrorMessage "Could not refresh the Add-Ins review display without scanning." & vbCrLf & vbCrLf & "Err.Number: " & CStr(Err.Number) & vbCrLf & "Err.Source: " & Err.Source & vbCrLf & "Err.Description: " & Err.Description, vbCritical, DEFAULT_UI_MESSAGE_TITLE, "ChatGPT_AddInReviewRefreshDisplay", "UI"
End Function

' Purpose: Adds in review scan inventory chat gpt UI.
' Public role: Form-code/event-expression entry point. Generated or managed Access forms can invoke this Public routine from form code or an event expression.
' Parameters: None.
' Returns: Boolean: True when chat gpt add in review clear filters completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or not-applicable path.
Public Function AddInReviewScanInventory_ChatGptUI() As Boolean
    On Error GoTo ErrHandler
    If RunSystemCommand(SYS_MACRO_KEY_SCAN_DB_INVENTORY_ADD_INS_AND_LOADED_VBA_PROJECTS) = False Then Exit Function
    If ChatGPT_AddInReviewRefreshDisplay() = False Then Exit Function
    AddInReviewScanInventory_ChatGptUI = True
    Exit Function
ErrHandler:
    CoreShowErrorMessage "The Reference Inventory refresh failed. Review the error details below to identify the failing inventory step and its recovery state." & vbCrLf & vbCrLf & "Err.Number: " & CStr(Err.Number) & vbCrLf & "Err.Source: " & Err.Source & vbCrLf & "Err.Description: " & Err.Description, vbCritical, DEFAULT_UI_MESSAGE_TITLE, "AddInReviewScanInventory_ChatGptUI", "UI"
End Function

' Purpose: Adds in review scan loaded VBA analysis chat gpt UI.
' Public role: Form-code/event-expression entry point. Generated or managed Access forms can invoke this Public routine from form code or an event expression.
' Parameters: None.
' Returns: Boolean: True when add in review scan loaded VBA analysis chat gpt UI completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or
'   not-applicable path.
Public Function AddInReviewScanLoadedVBAAnalysis_ChatGptUI() As Boolean
    On Error GoTo ErrHandler
    If RunSystemCommand(SYS_MACRO_KEY_SCAN_DB_INVENTORY_LOADED_VBA_SOURCE) = False Then Exit Function
    If ChatGPT_AddInReviewRefreshDisplay() = False Then Exit Function
    AddInReviewScanLoadedVBAAnalysis_ChatGptUI = True
    Exit Function
ErrHandler:
    CoreShowErrorMessage "The Loaded VBA Analysis refresh failed. This workflow does not intentionally load unloaded add-ins. Review the error details below to identify whether the scan or display refresh failed." & vbCrLf & vbCrLf & "Err.Number: " & CStr(Err.Number) & vbCrLf & "Err.Source: " & Err.Source & vbCrLf & "Err.Description: " & Err.Description, vbCritical, DEFAULT_UI_MESSAGE_TITLE, "AddInReviewScanLoadedVBAAnalysis_ChatGptUI", "UI"
End Function

' Purpose: Opens the raw-data view associated with the current Add-In Review selection.
' Public role: Form-code/event-expression entry point. Generated or managed Access forms can invoke this Public routine from form code or an event expression.
' Parameters: None.
' Returns: Boolean: True when chat gpt add in review open raw data completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or not-applicable path.
Public Function ChatGPT_AddInReviewOpenRawData() As Boolean
    ' Wrapper purpose: Generated Add-In Review callback that maps the selected review tab to its underlying raw-data object and opens that object.
    ' Why retained: Retained because the generated control needs a stable public Function while tab-to-object mapping and opening logic remain UI_REVIEW-owned.
    On Error GoTo ErrHandler
    Dim pageIndex As Long
    Dim sourceName As String
    Dim sourceType As AcObjectType
    If AddInReviewHostIsLoaded() = False Then Exit Function
    pageIndex = CLng(Nz(forms(TargetManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review")).Controls("tabAddInReview").value, 0))
    Select Case pageIndex
        Case 0: sourceName = ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Review_Summary"): sourceType = acQuery
        Case 1: sourceName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_VBEAddIns"): sourceType = acTable
        Case 2: sourceName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_COMAddIns"): sourceType = acTable
        Case 3: sourceName = ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_References_Loaded_Projects"): sourceType = acQuery
        Case 4: sourceName = ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_ExternalProject_References"): sourceType = acQuery
        Case 5: sourceName = ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_ExternalProject_Modules"): sourceType = acQuery
        Case 6: sourceName = ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_ExternalProject_Public_Procedures"): sourceType = acQuery
        Case 7: sourceName = ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_ExternalProject_Procedure_Conflicts"): sourceType = acQuery
        Case 8: sourceName = ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_ExternalProject_Dependencies"): sourceType = acQuery
        Case Else
            Err.Raise vbObjectError + 7304, "ChatGPT_AddInReviewOpenRawData", "The selected Add-In Review page does not have a defined raw-data source." & vbCrLf & "Page index: " & CStr(pageIndex)
    End Select
    If sourceType = acTable Then
        DoCmd.OpenTable sourceName, acViewNormal, acReadOnly
    Else
        DoCmd.OpenQuery sourceName, acViewNormal, acReadOnly
    End If
    ChatGPT_AddInReviewOpenRawData = True
    Exit Function
ErrHandler:
    CoreShowErrorMessage "Could not open the underlying read-only Add-In Review data source." & vbCrLf & vbCrLf & "Source: " & IIf(Len(sourceName) = 0, "(not resolved)", sourceName) & vbCrLf & "Err.Number: " & CStr(Err.Number) & vbCrLf & "Err.Source: " & Err.Source & vbCrLf & "Err.Description: " & Err.Description, vbCritical, DEFAULT_UI_MESSAGE_TITLE, "ChatGPT_AddInReviewOpenRawData", "UI"
End Function

' Purpose: Clears Add-In Review filters and refreshes the display to the unfiltered review set.
' Public role: Form-code/event-expression entry point. Generated or managed Access forms can invoke this Public routine from form code or an event expression.
' Parameters: None.
' Returns: Boolean: True when chat gpt add in review show procedure conflicts completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or
'   not-applicable path.
Public Function ChatGPT_AddInReviewClearFilters() As Boolean
    ' Wrapper purpose: Generated Add-In Review callback that clears filters across all review child forms.
    ' Why retained: Retained because the generated Clear Filters control needs a stable public Function while child-form enumeration and reset mechanics remain private.
    On Error GoTo ErrHandler
    Dim controlNames As Variant
    Dim controlItem As Variant
    Dim childForm As Access.Form
    If AddInReviewHostIsLoaded() = False Then Exit Function
    mAddInReviewSelectionUpdatesSuspended = True
    RestoreAddInReviewProjectLinks
    controlNames = Array("sub_pgAddInReviewInventory", "sub_pgAddInReviewVBE", "sub_pgAddInReviewCOM", "sub_pgAddInReviewProjects", "sub_pgAddInReviewReferences", "sub_pgAddInReviewModules", "sub_pgAddInReviewProcedures", "sub_pgAddInReviewConflicts", "sub_pgAddInReviewDependencies")
    For Each controlItem In controlNames
        SetupControlPump
        Set childForm = forms(TargetManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review")).Controls(CStr(controlItem)).Form
        childForm.FilterOn = False
        childForm.OrderByOn = False
        childForm.Requery
    Next controlItem
    mAddInReviewSelectionUpdatesSuspended = False
    ChatGPT_AddInReviewClearFilters = True
    Exit Function
ErrHandler:
    mAddInReviewSelectionUpdatesSuspended = False
    CoreShowErrorMessage "Could not clear the Add-Ins review filters." & vbCrLf & vbCrLf & "Err.Number: " & CStr(Err.Number) & vbCrLf & "Err.Source: " & Err.Source & vbCrLf & "Err.Description: " & Err.Description, vbCritical, DEFAULT_UI_MESSAGE_TITLE, "ChatGPT_AddInReviewClearFilters", "UI"
End Function

Private Sub RestoreAddInReviewProjectLinks()
    Dim controlNames As Variant
    Dim controlItem As Variant
    If AddInReviewHostIsLoaded() = False Then Exit Sub
    controlNames = Array("sub_pgAddInReviewReferences", "sub_pgAddInReviewModules", "sub_pgAddInReviewProcedures", "sub_pgAddInReviewConflicts", "sub_pgAddInReviewDependencies")
    For Each controlItem In controlNames
        SetupControlPump
        With forms(TargetManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review")).Controls(CStr(controlItem))
            .LinkMasterFields = "txtSelectedProjectIdentity"
            .LinkChildFields = FIELD_PROJECT_IDENTITY
        End With
    Next controlItem
End Sub

Private Function AddInReviewHostIsLoaded() As Boolean
    On Error Resume Next
    AddInReviewHostIsLoaded = CurrentProject.AllForms(TargetManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review")).IsLoaded
    On Error GoTo 0
End Function

' Purpose: Shows the procedure-level detail associated with the selected module in Add-In Review.
' Public role: Form-code/event-expression entry point. Generated or managed Access forms can invoke this Public routine from form code or an event expression.
' Parameters:
'   moduleIdentity (Input; Variant): Stable module identity used to select or correlate the exact ADC definition/state handled by this procedure.
' Returns: Boolean: True when chat gpt add in review show module procedures completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or
'   not-applicable path.
Public Function ChatGPT_AddInReviewShowModuleProcedures(ByVal moduleIdentity As Variant) As Boolean
    ' Wrapper purpose: Fixed-argument Add-In Review callback that selects the Procedures page and filters the generic detail surface by ModuleIdentity.
    ' Why retained: Retained because generated controls need a stable semantic callback and should not supply page indexes, child-control names, or filter-field details directly.
    ChatGPT_AddInReviewShowModuleProcedures = FilterAddInReviewDetail(6, "sub_pgAddInReviewProcedures", FIELD_MODULE_IDENTITY, moduleIdentity, True)
End Function

' Purpose: Shows dependency detail associated with the selected module in Add-In Review.
' Public role: Form-code/event-expression entry point. Generated or managed Access forms can invoke this Public routine from form code or an event expression.
' Parameters:
'   moduleIdentity (Input; Variant): Stable module identity used to select or correlate the exact ADC definition/state handled by this procedure.
' Returns: Boolean: True when chat gpt add in review show module dependencies completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or
'   not-applicable path.
Public Function ChatGPT_AddInReviewShowModuleDependencies(ByVal moduleIdentity As Variant) As Boolean
    ' Wrapper purpose: Fixed-argument Add-In Review callback that selects the Dependencies page and filters the generic detail surface by SourceModuleIdentity.
    ' Why retained: Retained because generated controls need a stable semantic callback and should not supply page indexes, child-control names, or filter-field details directly.
    ChatGPT_AddInReviewShowModuleDependencies = FilterAddInReviewDetail(8, "sub_pgAddInReviewDependencies", FIELD_SOURCE_MODULE_IDENTITY, moduleIdentity, True)
End Function

' Purpose: Shows conflict detail associated with the selected procedure in Add-In Review.
' Public role: Form-code/event-expression entry point. Generated or managed Access forms can invoke this Public routine from form code or an event expression.
' Parameters:
'   procedureNameKey (Input; Variant): Canonical key identifying the procedure name definition, setting, or managed item to use.
' Returns: Boolean: True when chat gpt add in review show procedure conflicts completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or
'   not-applicable path.
Public Function ChatGPT_AddInReviewShowProcedureConflicts(ByVal procedureNameKey As Variant) As Boolean
    ' Wrapper purpose: Generated Add-In Review callback that configures and displays the conflicts surface for the selected procedure name.
    ' Why retained: Retained because the generated control needs a stable public Function while page selection, link-field reset, filtering, and errors remain UI_REVIEW-owned.
    On Error GoTo ErrHandler
    If AddInReviewHostIsLoaded() = False Then Exit Function
    With forms(TargetManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review")).Controls("sub_pgAddInReviewConflicts")
        .LinkMasterFields = vbNullString
        .LinkChildFields = vbNullString
    End With
    ChatGPT_AddInReviewShowProcedureConflicts = FilterAddInReviewDetail(7, "sub_pgAddInReviewConflicts", FIELD_PROCEDURE_NAME_KEY, procedureNameKey, False)
    Exit Function
ErrHandler:
    CoreShowErrorMessage "Could not navigate to all loaded External-project exposures for the selected procedure name." & vbCrLf & vbCrLf & "Err.Number: " & CStr(Err.Number) & vbCrLf & "Err.Source: " & Err.Source & vbCrLf & "Err.Description: " & Err.Description, vbCritical, DEFAULT_UI_MESSAGE_TITLE, "ChatGPT_AddInReviewShowProcedureConflicts", "UI"
End Function

Private Function FilterAddInReviewDetail(ByVal pageIndex As Long, ByVal subformControlName As String, ByVal fieldName As String, ByVal fieldValue As Variant, ByVal preserveProjectLink As Boolean) As Boolean
    On Error GoTo ErrHandler
    Dim valueText As String
    Dim childForm As Access.Form
    If AddInReviewHostIsLoaded() = False Then Exit Function
    valueText = Trim$(CStr(Nz(fieldValue, vbNullString)))
    If Len(valueText) = 0 Then Err.Raise vbObjectError + 7291, "FilterAddInReviewDetail", "The selected row does not contain the stable key required for this navigation." & vbCrLf & "Field: " & fieldName
    If preserveProjectLink Then RestoreAddInReviewProjectLinks
    forms(TargetManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review")).Controls("tabAddInReview").value = pageIndex
    Set childForm = forms(TargetManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review")).Controls(subformControlName).Form
    childForm.Filter = "[" & fieldName & "]='" & Replace(valueText, "'", "''") & "'"
    childForm.FilterOn = True
    FilterAddInReviewDetail = True
    Exit Function
ErrHandler:
    CoreShowErrorMessage "Could not complete Add-Ins review cross-navigation." & vbCrLf & vbCrLf & "Stable key field: " & fieldName & vbCrLf & "Stable key value: " & UIVariantForMessage(fieldValue) & vbCrLf & "Err.Number: " & CStr(Err.Number) & vbCrLf & "Err.Source: " & Err.Source & vbCrLf & "Err.Description: " & Err.Description, vbCritical, DEFAULT_UI_MESSAGE_TITLE, "FilterAddInReviewDetail", "UI"
End Function
