Attribute VB_Name = "x__UTIL_ADC_UI_FORMS"
' ================================================================================================================================================================================
' ChatGPT Rewrite Tracking Header
' ================================================================================================================================================================================
' ChatGPT Previous Version ID: CGPT-RXGA-TJHJ-2JSL-9KH2-SLL7
' ChatGPT Baseline Version ID: CGPT-RXGA-TJHJ-2JSL-9KH2-SLL7
' ChatGPT Version ID: CGPT-VRFQ-NCV6-27W9-6NBA-YBC4
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
' STANDARD MODULE: Access Development Console Generated Forms UI Utility
' VB_Name: x__UTIL_ADC_UI_FORMS
' Purpose: Owns generated-form planning, construction, validation, runtime form policies, generated-form catalog integration, deletion UI, and Report Error form behavior.
' ================================================================================================================================================================================

Option Compare Database
Option Explicit

' ===============================================================================================================================================================================
' MODULE-LEVEL DECLARATIONS
' ===============================================================================================================================================================================

' Setup generated-form preflight plan and proof state.
Private Enum GeneratedFormPreflightScopeState
    GENERATED_FORM_PREFLIGHT_SCOPE_UNKNOWN = 0
    GENERATED_FORM_PREFLIGHT_SCOPE_ALL = 1
    GENERATED_FORM_PREFLIGHT_SCOPE_CHATGPT_UI_FORMS = 2
    GENERATED_FORM_PREFLIGHT_SCOPE_SUPPORT_FORMS = 3
    GENERATED_FORM_PREFLIGHT_SCOPE_LOOKUP_EDITORS = 4
End Enum
Private mSetupGeneratedFormPreflightReady As Boolean
Private mSetupGeneratedFormPreflightReplaceExisting As Boolean
Private mSetupGeneratedFormPreflightScope As GeneratedFormPreflightScopeState
' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION:
' Captures the exact physical generated-form names checked during Setup preflight so the build phase consumes the preflighted scope rather than rediscovering it. The Setup preflight trace/scope state bounds
' lifetime; UI owns the mutable Collection.
Private mSetupGeneratedFormPreflightPhysicalNames As Collection
Private mSetupGeneratedFormCompleteSetFastPath As Boolean
Private mSetupGeneratedFormCompleteSetToken As String
Private mSetupGeneratedFormPreflightTraceRunID As String

' Delete-object scope selection and launch-dialog timing state.
Private mDeleteObjectScopeChoice As ObjectDeletionScope
Private mDeleteObjectScopeCallbackReached As Boolean
' Macro 65 launch diagnostics are operation-scoped; they separate ADC preparation from modal user decision time.
Private mDeleteScopePromptStartTimer As Single
Private mDeleteScopeDialogReadyTimer As Single
Private mDeleteScopeDialogReady As Boolean
Private mDeleteScopePreparationElapsed As Double

' Generated-form build behavior flags; callers combine only the behaviors required by their operation.
Public Enum GeneratedFormBuildOptions
    GENERATED_FORM_BUILD_OPTION_NONE = 0
    GENERATED_FORM_BUILD_OPTION_GROUP_NAVIGATION_PANE = 1
    GENERATED_FORM_BUILD_OPTION_REQUIRE_NO_SKIPPED_FORMS = 2
    GENERATED_FORM_BUILD_OPTION_REGISTER_GENERATED_FORMS = 4
    GENERATED_FORM_BUILD_OPTION_PRESERVE_SETUP_PREFLIGHT_PLAN = 8
    GENERATED_FORM_BUILD_OPTIONS_STANDARD = 5
End Enum

' Generated-form same-run D2/D3 proof and build-currentness cache state.
Private mGeneratedFormSameRunD2ProofTraceRunID As String
Private mGeneratedFormSameRunD2ProofScope As String
Private mGeneratedFormSameRunD2ProofToken As String
Private mGeneratedFormSameRunD2DiscoveredUserFormCount As Long
Private mGeneratedFormSameRunD3ProofTraceRunID As String
Private mGeneratedFormSameRunD3ProofScope As String
Private mGeneratedFormSameRunD3ProofToken As String
' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION:
' Records generated forms independently verified as already current during one build pass; the signature/source/existence dictionaries are operation caches keyed over canonical generated-form definitions and
' physical sources, and mChatGPTGeneratedFormsRequiringBuild is the distinct build-plan membership set. Existing build/preflight reset paths invalidate them; none is a second generated-form authority and callers do
' not receive the mutable caches.
Private mChatGPTVerifiedExistingForms As Collection
Private mChatGPTGeneratedFormSignatureCurrentCache As Object
Private mChatGPTGeneratedFormSourceFingerprintCache As Object
Private mChatGPTGeneratedFormLookupEditorPolicyFingerprintCache As Object
Private mChatGPTGeneratedFormObjectExistsCache As Object
Private mChatGPTGeneratedFormSourceFingerprintCacheHits As Long
Private mChatGPTGeneratedFormSourceFingerprintCacheMisses As Long
Private mChatGPTGeneratedFormObjectExistsCacheHits As Long
Private mChatGPTGeneratedFormObjectExistsCacheMisses As Long
Private mChatGPTGeneratedFormDesignCreateCount As Long
Private mChatGPTGeneratedFormDesignSaveCloseCount As Long
Private mChatGPTGeneratedFormsUsePopups As Boolean
Private mChatGPTGeneratedFormsRequiringBuild As Object
' Shared generated-form replacement recovery state.
Private mSharedPreparedGeneratedFormName As String

' Generated-form catalog row contract.
Private Type TGeneratedFormCatalogRow
    formName As String
    formDisplayName As String
    FormDescription As String
    formType As String
    NavigationPaneGroupKey As String
    ChatGPTUIForm As Boolean
    SortOrder As Long
End Type

' Foundation modes, form-builder geometry, UI text, and build/progress state.
Private Const FOUNDATION_MODE_FULL As String = "Full"
Private Const FOUNDATION_MODE_REVIEW_ONLY As String = "ReviewOnly"
Private Const FOUNDATION_MODE_SKIP As String = "Skip"
Private Const FORM_BUILD_SCOPE_ALL As String = "All"
Private Const FORM_BUILD_SCOPE_CHATGPT_UI_FORMS As String = "ChatGPTUIForms"
Private Const FORM_BUILD_SCOPE_SUPPORT_FORMS As String = "ADCSupportForms"
Private Const FORM_BUILD_SCOPE_LOOKUP_EDITORS As String = "LookupEditors"
Private Const CONTINUOUS_LIST_LAYOUT_VERSION As String = "2026-08-07 Date/Time width v1"
Private Const SPLIT_POPUP_LAYOUT_VERSION As String = "2026-08-07 datasheet bottom v1"

' Generated-form builder diagnostic context.
Private mChatGPTBuildStep As String
Private mChatGPTBuildForm As String
Private mChatGPTBuildAction As String
Private mChatGPTBuildDetail As String
Private mChatGPTBuildSourceName As String
Private mChatGPTBuildFieldName As String
Private mChatGPTBuildControlName As String
Private mChatGPTBuildSectionName As String
Private mChatGPTBuildParentName As String
Private mChatGPTBuildBounds As String

Private Const ERR_FORM_RESTORE_FAILED As Long = vbObjectError + 7211

Private Const ERR_GENERATED_CATALOG_ROLLBACK_FAILED As Long = vbObjectError + 7219

Private Const FORM_BUILDER_VERSION As String = "2026-08-20 lookup-list editors v2"
Private Const GENERATED_FORM_LOOKUP_EDITOR_BINDING_CONTRACT_VERSION As String = "2026-08-25 deterministic lookup-editor binding v1"
Private Const LOOKUP_EDITOR_EXACT_CONTROL_CONTRACT_VERSION As String = "2026-08-21 lookup-list referenced-value add-delete contract v5"
Private Const LOOKUP_EDITOR_DELETE_GUARD_EXPRESSION As String = "=GuardLookupEditorDelete()"
Private Const LOOKUP_EDITOR_CODE_FIELD_EDIT_POLICY_EXPRESSION As String = "=ApplyLookupEditorCodeFieldEditPolicy()"
Private Const LOOKUP_EDITOR_RUNTIME_POLICY_EXPRESSION As String = "=ApplyLookupEditorRuntimePolicy()"
Private Const DELETE_OBJECT_SCOPE_DIALOG_VERSION As String = "2026-08-17 Macro 65 explicit-selection contract v3"
' REGRESSION CONTRACT: bump DELETE_OBJECT_SCOPE_DIALOG_VERSION whenever the Delete ADC Objects selector physical builder contract changes, including controls, callbacks, OnLoad behavior, materially significant
' layout/dimensions, or scope choices.

Private Const UTILITY_LAZY_LOAD_FORM_VERSION As String = "2026-08-10 selected-tab lazy subform load v1"
Private Const GENERATED_FORM_BUILD_SIGNATURE_PROPERTY As String = "ADCBuildSignature"
Private Const GENERATED_FORM_DEFINITION_HASH_PROPERTY As String = "ADCDefinitionHash"
Private Const GENERATED_FORM_COMPLETE_SET_TOKEN_PROPERTY As String = "GeneratedFormCompleteSetToken"
Private Const GENERATED_FORM_PHYSICAL_STATE_TOKEN_PROPERTY As String = "GeneratedFormPhysicalStateToken"
Private Const GENERATED_FORM_CHATGPT_UI_COMPLETE_SET_TOKEN_PROPERTY As String = "ADCGeneratedFormChatGPTUICompleteSetToken"
Private Const GENERATED_FORM_CHATGPT_UI_PHYSICAL_STATE_TOKEN_PROPERTY As String = "ADCGeneratedFormChatGPTUIPhysicalStateToken"
Private Const GENERATED_FORM_SUPPORT_COMPLETE_SET_TOKEN_PROPERTY As String = "ADCGeneratedFormADCSupportCompleteSetToken"
Private Const GENERATED_FORM_SUPPORT_PHYSICAL_STATE_TOKEN_PROPERTY As String = "ADCGeneratedFormADCSupportPhysicalStateToken"
Private Const GENERATED_FORM_LOOKUP_EDITORS_COMPLETE_SET_TOKEN_PROPERTY As String = "ADCGeneratedFormLookupEditorsCompleteSetToken"
Private Const GENERATED_FORM_LOOKUP_EDITORS_PHYSICAL_STATE_TOKEN_PROPERTY As String = "ADCGeneratedFormLookupEditorsPhysicalStateToken"

' Report Error capture UI state.
Private Type TReportErrorGuid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
Private Type TReportErrorPictureDescription
    Size As Long
    PictureType As Long
    bitmapHandle As LongPtr
    PaletteHandle As LongPtr
End Type

' Windows API declarations used by error capture.
Private Declare PtrSafe Function ReportErrorGetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal windowHandle As LongPtr, ByRef windowRect As TReportErrorRect) As Long
Private Declare PtrSafe Function ReportErrorGetWindowDC Lib "user32" Alias "GetWindowDC" (ByVal windowHandle As LongPtr) As LongPtr
Private Declare PtrSafe Function ReportErrorReleaseDC Lib "user32" Alias "ReleaseDC" (ByVal windowHandle As LongPtr, ByVal deviceContext As LongPtr) As Long
Private Declare PtrSafe Function ReportErrorCreateCompatibleDC Lib "gdi32" Alias "CreateCompatibleDC" (ByVal deviceContext As LongPtr) As LongPtr
Private Declare PtrSafe Function ReportErrorDeleteDC Lib "gdi32" Alias "DeleteDC" (ByVal deviceContext As LongPtr) As Long
Private Declare PtrSafe Function ReportErrorCreateCompatibleBitmap Lib "gdi32" Alias "CreateCompatibleBitmap" (ByVal deviceContext As LongPtr, ByVal bitmapWidth As Long, ByVal bitmapHeight As Long) As LongPtr
Private Declare PtrSafe Function ReportErrorSelectObject Lib "gdi32" Alias "SelectObject" (ByVal deviceContext As LongPtr, ByVal objectHandle As LongPtr) As LongPtr
Private Declare PtrSafe Function ReportErrorDeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal objectHandle As LongPtr) As Long
Private Declare PtrSafe Function ReportErrorBitBlt Lib "gdi32" Alias "BitBlt" (ByVal destinationDC As LongPtr, ByVal destinationX As Long, ByVal destinationY As Long, ByVal bitmapWidth As Long, ByVal bitmapHeight As Long, ByVal sourceDC As LongPtr, ByVal sourceX As Long, ByVal sourceY As Long, ByVal rasterOperation As Long) As Long
Private Declare PtrSafe Function ReportErrorOleCreatePictureIndirect Lib "oleaut32" Alias "OleCreatePictureIndirect" (ByRef pictureDescription As TReportErrorPictureDescription, ByRef interfaceID As TReportErrorGuid, ByVal pictureOwnsHandle As Long, ByRef pictureObject As stdole.IPictureDisp) As Long

' Error-report capture constants and submitted description state.
Private Const REPORT_ERROR_PICTURE_TYPE_BITMAP As Long = 1
Private Const REPORT_ERROR_RASTER_COPY As Long = &HCC0020
Private Const REPORT_ERROR_CAPTURE_LAYERED_WINDOWS As Long = &H40000000
Private Const REPORT_ERROR_PNG_FORMAT_ID As String = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"

' ===============================================================================================================================================================================
' PROCEDURES
' ===============================================================================================================================================================================

' ===============================================================================================================================================================================
' GENERATED-FORM PREFLIGHT / RECOVERY / BUILD-CONTEXT PROCEDURES
' ===============================================================================================================================================================================

Private Function ResetPersistentGeneratedLauncherShellToBlank(ByVal frm As Form) As Long
    'CreateForm can inherit the configured Access Form Template. Strip template controls/data/module state so a newly created persistent launcher starts as a minimal blank shell.
    Dim removedControlCount As Long
    If frm Is Nothing Then Err.Raise vbObjectError + 9180, "ResetPersistentGeneratedLauncherShellToBlank", "The new persistent launcher form is missing."
    Do While frm.Controls.count > 0
        DeleteControl frm.Name, frm.Controls(0).Name
        removedControlCount = removedControlCount + 1
    Loop
    frm.recordSource = vbNullString
    frm.Filter = vbNullString
    frm.OrderBy = vbNullString
    If frm.HasModule Then frm.HasModule = False
    ResetPersistentGeneratedLauncherShellToBlank = removedControlCount
End Function

' Purpose: Ensures persistent generated launcher forms for setup.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   detailsText (Output Optional; String): Receives details text used for caller diagnostics, verification output, or user-facing reporting. Default: vbNullString.
' Returns: None (Sub). Results are returned through the documented ByRef parameters and/or the routine's database, UI, or module-state side effects.
Public Sub EnsurePersistentGeneratedLauncherFormsForSetup(Optional ByRef detailsText As String = vbNullString)
    'Forms-phase authority for required persistent generated launcher shells. Both forms are physically CodeDb-owned; existing designs are preserved intact.
    RequireSetupDomainMutationAllowed SETUP_DOMAIN_D1_PERSISTENT_LAUNCHER_FORMS, "EnsurePersistentGeneratedLauncherFormsForSetup"
    On Error GoTo ErrHandler
    Dim item As Variant
    Dim formName As String
    Dim temporaryFormName As String
    Dim frm As Form
    Dim createdCount As Long
    Dim preservedCount As Long
    Dim removedTemplateControlCount As Long
    For Each item In ManagedPersistentGeneratedFormNames()
        PumpSetupResponsiveness
        formName = CStr(item)
        If Len(formName) = 0 Then Err.Raise vbObjectError + 9181, "EnsurePersistentGeneratedLauncherFormsForSetup", "A canonical persistent generated form name is blank."
        If Len(formName) > MAX_ACCESS_OBJECT_NAME_LENGTH Then Err.Raise vbObjectError + 9183, "EnsurePersistentGeneratedLauncherFormsForSetup", "A canonical persistent generated form name exceeds the Access object-name limit." & vbCrLf & "Form: " & formName & vbCrLf & "Length: " & CStr(Len(formName)) & vbCrLf & "Maximum: " & CStr(MAX_ACCESS_OBJECT_NAME_LENGTH)
        If SharedDao_AccessObjectExists(ADCDb(), acForm, formName) Then
            preservedCount = preservedCount + 1
            AddTraceNote "Persistent generated launcher form already exists and was preserved unchanged: " & formName
            LogTraceMessage TRACE_LEVEL_INFO, "Setup", "EnsurePersistentGeneratedLauncherFormsForSetup", "Persistent generated launcher form already existed and was preserved.", "Form: " & formName & vbCrLf & "PhysicalOwnerDatabaseRole: CodeDb" & vbCrLf & "CatalogOwner: " & DEFAULT_CATALOG_OWNER_PERSISTENT_GENERATED, "D1 persistent launcher forms", 0, vbNullString, vbNullString, formName, "Form", "CodeDb", "PersistentGenerated"
        Else
            Set frm = CreateForm(ADCDb().Name)
            temporaryFormName = frm.Name
            removedTemplateControlCount = ResetPersistentGeneratedLauncherShellToBlank(frm)
            DoCmd.Close acForm, temporaryFormName, acSaveYes
            DoCmd.Rename formName, acForm, temporaryFormName
            temporaryFormName = vbNullString
            If SharedDao_AccessObjectExists(ADCDb(), acForm, formName) = False Then Err.Raise vbObjectError + 9182, "EnsurePersistentGeneratedLauncherFormsForSetup", "The persistent generated launcher shell was not created in CodeDb." & vbCrLf & "Form: " & formName & vbCrLf & "CodeDb: " & ADCDb().Name
            createdCount = createdCount + 1
            AddTraceNote "Persistent generated launcher form created as a blank CodeDb-owned shell: " & formName
            LogTraceMessage TRACE_LEVEL_INFO, "Setup", "EnsurePersistentGeneratedLauncherFormsForSetup", "Persistent generated launcher form created.", "Form: " & formName & vbCrLf & "PhysicalOwnerDatabaseRole: CodeDb" & vbCrLf & "CatalogOwner: " & DEFAULT_CATALOG_OWNER_PERSISTENT_GENERATED & vbCrLf & "Inherited template controls removed: " & CStr(removedTemplateControlCount), "D1 persistent launcher forms", 0, vbNullString, vbNullString, formName, "Form", "CodeDb", "PersistentGenerated"
        End If
        CompleteStandaloneGeneratedFormDescriptionIfAllowed formName
    Next item
    detailsText = "Persistent generated launcher shells: created=" & CStr(createdCount) & "; preserved=" & CStr(preservedCount) & "."
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Len(temporaryFormName) > 0 Then
        DoCmd.Close acForm, temporaryFormName, acSaveNo
        If SharedDao_AccessObjectExists(ADCDb(), acForm, temporaryFormName) Then DoCmd.DeleteObject acForm, temporaryFormName
    End If
    On Error GoTo 0
    Err.Raise savedErrNumber, "EnsurePersistentGeneratedLauncherFormsForSetup", "Could not ensure required persistent generated launcher forms." & vbCrLf & "Form: " & IIf(Len(formName) = 0, "(not set)", formName) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Function GeneratedFormPersistedCompactProofAvailable(ByVal normalizedBuildScope As String, ByRef storedCompleteToken As String, ByRef storedPhysicalToken As String, ByRef detailText As String) As Boolean
    Dim db As DAO.Database
    Dim completePropertyExists As Boolean
    Dim physicalPropertyExists As Boolean
    storedCompleteToken = vbNullString
    storedPhysicalToken = vbNullString
    detailText = vbNullString
    If GeneratedFormAuthoritativeCount(normalizedBuildScope) <= 0 Then detailText = "No authoritative generated-form identities are available for build scope [" & normalizedBuildScope & "].": Exit Function
    Set db = GetHostDb()
    storedCompleteToken = DatabaseTextPropertyValue(db, GeneratedFormCompleteSetTokenPropertyName(normalizedBuildScope), completePropertyExists)
    storedPhysicalToken = DatabaseTextPropertyValue(db, GeneratedFormPhysicalStateTokenPropertyName(normalizedBuildScope), physicalPropertyExists)
    If completePropertyExists = False Or Len(storedCompleteToken) = 0 Or physicalPropertyExists = False Or Len(storedPhysicalToken) = 0 Then
        detailText = "Persisted generated-form compact proof is absent or incomplete; skipped impossible compact comparison."
        Exit Function
    End If
    GeneratedFormPersistedCompactProofAvailable = True
End Function

Private Sub PrepareGeneratedFormExpectedPreflightEvidence(ByVal preflightDb As DAO.Database, ByVal normalizedBuildScope As String)
    Dim helpTopics As Collection
    Dim systemMacroDefinitions As Collection
    Dim systemMacroDefinitionIndex As Object
    preflightDb.TableDefs.Refresh
    preflightDb.QueryDefs.Refresh
    Set mChatGPTGeneratedFormSourceFingerprintCache = Nothing
    Set mChatGPTGeneratedFormLookupEditorPolicyFingerprintCache = Nothing
    Set mChatGPTGeneratedFormObjectExistsCache = Nothing
    mSetupGeneratedFormPreflightTraceRunID = GetActiveOrLastTraceRunID()
    ClearSetupGeneratedFormHelpProofState
    If StrComp(normalizedBuildScope, FORM_BUILD_SCOPE_ALL, vbTextCompare) = 0 Then
        Set systemMacroDefinitions = x__UTIL_ADC_DEFAULTS.GetSystemMacroDefinitions()
        Set systemMacroDefinitionIndex = ChatGPTSystemMacroDefinitionIndex(systemMacroDefinitions)
        Set helpTopics = ChatGPTHelpTopics(systemMacroDefinitions, systemMacroDefinitionIndex)
        ValidateChatGPTHelpDefinitions helpTopics, systemMacroDefinitions, systemMacroDefinitionIndex
        PublishSetupGeneratedFormHelpProof ChatGPTHelpFixedTopicSignature(helpTopics)
    End If
End Sub

' Purpose: Performs preflight generated forms for setup by coordinating `NormalizeChatGPTBuildScope` and `CodeDbIsHostDb` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   replaceExisting (Input Optional; Boolean): When True, enables or requests existing for this call; when False, the routine follows its normal non-forced/default path. Default: True.
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: False.
'   buildScope (Input Optional; String): Scope value limiting which build items or operations are included in this call. Default: FORM_BUILD_SCOPE_ALL.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub PreflightGeneratedFormsForSetup(Optional ByVal replaceExisting As Boolean = True, Optional ByVal showMessage As Boolean = False, Optional ByVal buildScope As String = FORM_BUILD_SCOPE_ALL)
    'Mutation-free setup preflight; caches the exact physical-build set for the immediately following apply phase.
    On Error GoTo ErrHandler
    Dim normalizedBuildScope As String
    Dim physicalStateToken As String
    Dim compactDetail As String
    Dim totalPhysicalBuildCount As Long
    Dim preflightDb As DAO.Database
    Dim storedCompleteToken As String
    Dim storedPhysicalToken As String
    Dim compactProofAvailable As Boolean
    Dim expectedEvidencePrepared As Boolean
    ClearSetupGeneratedFormPreflightProofState
    normalizedBuildScope = NormalizeChatGPTBuildScope(buildScope)
    If CodeDbIsHostDb() = False Then Err.Raise vbObjectError + 6750, "PreflightGeneratedFormsForSetup", "Generated-form setup preflight requires CodeDb and HostDb to be the same active database."
    Set preflightDb = GetHostDb()
    compactProofAvailable = GeneratedFormPersistedCompactProofAvailable(normalizedBuildScope, storedCompleteToken, storedPhysicalToken, compactDetail)
    If compactProofAvailable Then
        PrepareGeneratedFormExpectedPreflightEvidence preflightDb, normalizedBuildScope
        expectedEvidencePrepared = True
        If GeneratedFormCompactPreflightMatches(normalizedBuildScope, mSetupGeneratedFormCompleteSetToken, physicalStateToken, compactDetail, storedCompleteToken, storedPhysicalToken) Then
            Set mSetupGeneratedFormPreflightPhysicalNames = New Collection
            mSetupGeneratedFormCompleteSetFastPath = True
            LogTraceMessage TRACE_LEVEL_INFO, "UI", "PreflightGeneratedFormsForSetup", "Generated-form compact setup preflight accepted before exhaustive source/control preflight.", compactDetail, "Generated-form compact preflight", 0, vbNullString, vbNullString, DEFAULT_CHATGPT_UI_DISPLAY_PREFIX, "FormSet", "CodeDb", "FormGeneration"
        End If
    Else
        AddTraceNote compactDetail
    End If
    If mSetupGeneratedFormCompleteSetFastPath = False Then
        If expectedEvidencePrepared = False Then
            PrepareGeneratedFormExpectedPreflightEvidence preflightDb, normalizedBuildScope
        End If
        PreflightGeneratedFormBuildPlan normalizedBuildScope
        Set mSetupGeneratedFormPreflightPhysicalNames = ChatGPTPhysicalFormChangeNamesRequiringBuild(replaceExisting, normalizedBuildScope)
        If mSetupGeneratedFormPreflightPhysicalNames.count = 0 Then mSetupGeneratedFormCompleteSetFastPath = GeneratedFormCompleteSetTokenMatches(normalizedBuildScope, mSetupGeneratedFormCompleteSetToken)
        If Len(mSetupGeneratedFormCompleteSetToken) = 0 Then mSetupGeneratedFormCompleteSetToken = GeneratedFormCompleteSetToken(normalizedBuildScope)
    End If
    mSetupGeneratedFormPreflightReplaceExisting = replaceExisting
    mSetupGeneratedFormPreflightScope = GeneratedFormPreflightScopeStateFromText(normalizedBuildScope)
    mSetupGeneratedFormPreflightReady = True
    totalPhysicalBuildCount = mSetupGeneratedFormPreflightPhysicalNames.count
    LogTraceMessage TRACE_LEVEL_INFO, "UI", "PreflightGeneratedFormsForSetup", "Generated-form setup preflight completed.", "Build scope: " & normalizedBuildScope & vbCrLf & "ADC-owned authoritative forms: " & CStr(GeneratedFormAuthoritativeCount(normalizedBuildScope)) & vbCrLf & "Forms requiring physical build: " & CStr(totalPhysicalBuildCount), "Generated-form setup preflight", 0, vbNullString, vbNullString, DEFAULT_CHATGPT_UI_DISPLAY_PREFIX, "FormSet", "CodeDb", "FormGeneration"
    If showMessage Then
        CoreShowMessage _
            "Generated-form setup preflight completed." & vbCrLf & "Forms requiring physical build: " & CStr(totalPhysicalBuildCount), _
            vbInformation, _
            DEFAULT_UI_MESSAGE_TITLE, _
            "ShowADCMessage", _
            "UI"
    End If
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    ClearSetupGeneratedFormPreflightProofState
    If showMessage Then
        CoreShowErrorMessage "Generated-form setup preflight failed." & vbCrLf & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, vbCritical, "Generated Form Setup Preflight", "PreflightGeneratedFormsForSetup", "UI"
        Exit Sub
    End If
    Err.Raise savedErrNumber, "PreflightGeneratedFormsForSetup", "Generated-form setup preflight failed." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & savedErrDescription
End Sub

' Purpose: Applies generated forms from setup preflight for build scope.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: False.
'   buildScope (Input Optional; String): Scope value limiting which build items or operations are included in this call. Default: FORM_BUILD_SCOPE_ALL.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub ApplyGeneratedFormsFromSetupPreflight(Optional ByVal showMessage As Boolean = False, Optional ByVal buildScope As String = FORM_BUILD_SCOPE_ALL)
    'D2/D3 consume their requested partition from the one D1 All-scope plan; neither partition may reconstruct its own plan. Normalize first so every accepted spelling is subject to the same fail-closed owner check.
    Dim preflightDetail As String
    Dim normalizedBuildScope As String
    normalizedBuildScope = NormalizeChatGPTBuildScope(buildScope)
    If StrComp(normalizedBuildScope, FORM_BUILD_SCOPE_CHATGPT_UI_FORMS, vbTextCompare) = 0 Then
        RequireSetupDomainMutationAllowed SETUP_DOMAIN_D2_CHATGPT_UI_FORMS, "ApplyGeneratedFormsFromSetupPreflight"
    ElseIf StrComp(normalizedBuildScope, FORM_BUILD_SCOPE_SUPPORT_FORMS, vbTextCompare) = 0 Then
        RequireSetupDomainMutationAllowed SETUP_DOMAIN_D3_SUPPORT_FORMS, "ApplyGeneratedFormsFromSetupPreflight"
    Else
        Err.Raise vbObjectError + 6770, "ApplyGeneratedFormsFromSetupPreflight", "Setup partition application supports only ChatGPTUIForms or ADCSupportForms. Requested scope=" & normalizedBuildScope
    End If
    If SetupGeneratedFormPreflightReadyForSetup(preflightDetail) = False Then Err.Raise vbObjectError + 6765, "ApplyGeneratedFormsFromSetupPreflight", "Setup D2/D3 requires the successful same-session D1 generated-form plan." & vbCrLf & preflightDetail
    If mSetupGeneratedFormPreflightScope <> GENERATED_FORM_PREFLIGHT_SCOPE_ALL Then Err.Raise vbObjectError + 6769, "ApplyGeneratedFormsFromSetupPreflight", "Setup D2/D3 requires D1's authoritative All-scope generated-form plan. Retained scope=" & GeneratedFormPreflightScopeStateText(mSetupGeneratedFormPreflightScope)
    BuildGeneratedForms True, showMessage, FOUNDATION_MODE_SKIP, normalizedBuildScope, GENERATED_FORM_BUILD_OPTION_REQUIRE_NO_SKIPPED_FORMS Or GENERATED_FORM_BUILD_OPTION_PRESERVE_SETUP_PREFLIGHT_PLAN
End Sub

' Purpose: Resolves parse generated form recovery state from the supplied/current ADC state, using `CreateObject` and `UIRecoveryBackupReadable` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   payloadText (Input; String): Input payload text for parse generated form recovery state; is tested by If Len(payloadText) = 0; is passed to `Split`.
' Returns: Collection: Returns the ordered Collection assembled by the routine; callers receive the object reference and should treat its membership/order according to the routine contract.
Public Function ParseGeneratedFormRecoveryState(ByVal payloadText As String) As Collection
    On Error GoTo ErrHandler
    ' INTENTIONAL ORDERED VIEW: physicalForms, backupEntries, originalForms, result preserves the deterministic sequence required by ParseGeneratedFormRecoveryState; keyed-only or unordered storage would not
    ' preserve the consumer semantics.
    Dim physicalForms As New Collection
    Dim backupEntries As New Collection
    Dim originalForms As New Collection
    Dim result As New Collection
    Dim seenPhysical As Object
    Dim seenBackup As Object
    Dim seenOriginal As Object
    Dim lineItem As Variant
    Dim parts As Variant
    Dim rowType As String
    Dim formName As String
    Dim backupPath As String
    Dim rollbackTable As String
    Dim versionSeen As Boolean
    Dim rollbackSeen As Boolean
    Dim backupFailure As String
    Dim item As Variant
    ' INTENTIONAL KEYED LOOKUP: seenPhysical, seenBackup, seenOriginal provides keyed membership/lookup over source rows or state already read by ParseGeneratedFormRecoveryState; it is a derived access path, not a
    ' second authority.
    Set seenPhysical = CreateObject("Scripting.Dictionary")
    Set seenBackup = CreateObject("Scripting.Dictionary")
    Set seenOriginal = CreateObject("Scripting.Dictionary")
    seenPhysical.CompareMode = vbTextCompare
    seenBackup.CompareMode = vbTextCompare
    seenOriginal.CompareMode = vbTextCompare
    If Len(payloadText) = 0 Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "ParseGeneratedFormRecoveryState", "The generated-form recovery payload is blank."
    For Each lineItem In Split(payloadText, vbCrLf)
        SetupControlPump
        If Len(CStr(lineItem)) = 0 Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "ParseGeneratedFormRecoveryState", "The generated-form recovery payload contains a blank row."
        parts = Split(CStr(lineItem), Chr$(30))
        rowType = UCase$(CStr(parts(0)))
        Select Case rowType
            Case "VERSION"
                If versionSeen Or UBound(parts) <> 1 Or StrComp(CStr(parts(1)), "1", vbBinaryCompare) <> 0 Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "ParseGeneratedFormRecoveryState", "The generated-form recovery payload has an unsupported or duplicate version row."
                versionSeen = True
            Case "BACKUP"
                If UBound(parts) <> 2 Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "ParseGeneratedFormRecoveryState", "A generated-form BACKUP row has an invalid field count."
                formName = Trim$(CStr(parts(1)))
                backupPath = Trim$(CStr(parts(2)))
                If Len(formName) = 0 Or Len(backupPath) = 0 Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "ParseGeneratedFormRecoveryState", "A generated-form BACKUP row contains a blank form name or backup path."
                If seenBackup.Exists(formName) Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "ParseGeneratedFormRecoveryState", "The generated-form recovery payload contains a duplicate BACKUP row." & vbCrLf & "Form: " & formName
                backupFailure = vbNullString
                If UIRecoveryBackupReadable(backupPath, backupFailure) = False Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "ParseGeneratedFormRecoveryState", backupFailure
                seenBackup.Add formName, True
                backupEntries.Add Array(formName, backupPath)
            Case "PHYSICAL"
                If UBound(parts) <> 1 Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "ParseGeneratedFormRecoveryState", "A generated-form PHYSICAL row has an invalid field count."
                formName = Trim$(CStr(parts(1)))
                If Len(formName) = 0 Or seenPhysical.Exists(formName) Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "ParseGeneratedFormRecoveryState", "A generated-form PHYSICAL row is blank or duplicated."
                seenPhysical.Add formName, True
                physicalForms.Add formName
            Case "ORIGINAL"
                If UBound(parts) <> 1 Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "ParseGeneratedFormRecoveryState", "A generated-form ORIGINAL row has an invalid field count."
                formName = Trim$(CStr(parts(1)))
                If Len(formName) = 0 Or seenOriginal.Exists(formName) Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "ParseGeneratedFormRecoveryState", "A generated-form ORIGINAL row is blank or duplicated."
                seenOriginal.Add formName, True
                originalForms.Add formName
            Case "ROLLBACKTABLE"
                If rollbackSeen Or UBound(parts) <> 1 Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "ParseGeneratedFormRecoveryState", "The generated-form recovery payload has an invalid or duplicate ROLLBACKTABLE row."
                rollbackSeen = True
                rollbackTable = Trim$(CStr(parts(1)))
            Case Else
                Err.Raise ERR_UI_RECOVERY_BLOCKED, "ParseGeneratedFormRecoveryState", "The generated-form recovery payload contains an unrecognized row type: " & rowType
        End Select
    Next lineItem
    If versionSeen = False Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "ParseGeneratedFormRecoveryState", "The generated-form recovery payload has no recognized version row."
    If physicalForms.count = 0 Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "ParseGeneratedFormRecoveryState", "The generated-form recovery payload contains no PHYSICAL rows."
    If rollbackSeen = False Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "ParseGeneratedFormRecoveryState", "The generated-form recovery payload has no ROLLBACKTABLE row."
    For Each item In originalForms
        SetupControlPump
        If seenPhysical.Exists(CStr(item)) = False Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "ParseGeneratedFormRecoveryState", "An ORIGINAL form is not present in the PHYSICAL form set." & vbCrLf & "Form: " & CStr(item)
    Next item
    For Each item In backupEntries
        SetupControlPump
        If seenPhysical.Exists(CStr(item(0))) = False Or seenOriginal.Exists(CStr(item(0))) = False Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "ParseGeneratedFormRecoveryState", "A BACKUP form is not present in both the PHYSICAL and ORIGINAL form sets." & vbCrLf & "Form: " & CStr(item(0))
    Next item
    If Len(rollbackTable) > 0 Then
        If SharedDao_TableExists(GetHostDb(), rollbackTable) = False Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "ParseGeneratedFormRecoveryState", "The generated-form catalog rollback table does not exist." & vbCrLf & "Table: " & rollbackTable
    End If
    result.Add physicalForms, "PhysicalForms"
    result.Add backupEntries, "BackupEntries"
    result.Add originalForms, "OriginalForms"
    result.Add rollbackTable, "CatalogRollbackTable"
    Set ParseGeneratedFormRecoveryState = result
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "ParseGeneratedFormRecoveryState", Err.Description
End Function

' Purpose: Restores generated form recovery state for rollback state.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   rollbackState (Input; Collection): State data describing the current or prior rollback condition used for recovery, rollback, UI, or lifecycle decisions.
' Returns: String: Returns the resolved restore generated form recovery state text in the canonical representation consumed by this module's caller; vbNullString represents no applicable value where the
'   existing contract permits it.
Public Function RestoreGeneratedFormRecoveryState(ByVal rollbackState As Collection) As String
    Dim recoveryClearDetail As String
    'Restores generated forms and generated-form catalog rows captured by the UI durable recovery boundary.
    On Error GoTo ErrHandler
    Dim physicalForms As Collection
    Dim backupEntries As Collection
    Dim originalForms As Collection
    Dim catalogRollbackTable As String
    If rollbackState Is Nothing Then
        RestoreGeneratedFormRecoveryState = "No generated-form rollback state was available."
        Exit Function
    End If
    Set physicalForms = rollbackState("PhysicalForms")
    Set backupEntries = rollbackState("BackupEntries")
    Set originalForms = rollbackState("OriginalForms")
    catalogRollbackTable = CStr(rollbackState("CatalogRollbackTable"))
    RestoreGeneratedFormRecoveryState = RestoreChatGPTGeneratedFormsAfterFailure(physicalForms, backupEntries, originalForms, catalogRollbackTable, False)
    If ClearUIRecoveryFailure("Generated-form restoration", recoveryClearDetail) = False Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "RestoreGeneratedFormRecoveryState", recoveryClearDetail
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    RegisterUIRecoveryFailure "Generated-form restoration", "Generated-form rollback did not complete." & vbCrLf & "Error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription, Nothing, rollbackState
    Err.Raise savedErrNumber, "RestoreGeneratedFormRecoveryState", savedErrDescription
End Function

Private Function GeneratedFormCodeDbOnlyMessage(ByVal actionText As String) As String
    GeneratedFormCodeDbOnlyMessage = actionText & " is CodeDb-only for " & DEFAULT_UTILITY_LONG_NAME & " generated-form operation." & vbCrLf & vbCrLf
    GeneratedFormCodeDbOnlyMessage = GeneratedFormCodeDbOnlyMessage & "CodeDb: " & codeDb.Name & vbCrLf
    GeneratedFormCodeDbOnlyMessage = GeneratedFormCodeDbOnlyMessage & "HostDb: " & GetHostDb().Name & vbCrLf & vbCrLf
    GeneratedFormCodeDbOnlyMessage = GeneratedFormCodeDbOnlyMessage & "HostDb should contain only MC_Records and generated macro objects. ADC generated forms remain internal to CodeDb."
End Function

' Purpose: Sets chat gpt build context for step text, form name.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   stepText (Input; String): Input step text for set chat GPT build context; is assigned to `mChatGPTBuildStep`.
'   formName (Input Optional; String): Physical Access form name identifying the form to create, open, configure, validate, or inspect. Default: vbNullString.
'   actionText (Input Optional; String): Input action text for set chat GPT build context; is assigned to `mChatGPTBuildAction`. Default: vbNullString.
'   detailText (Input Optional; String): Receives or supplies diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller. Default: vbNullString.
'   sourceName (Input Optional; String): Name of the source object/query/table/control from which this operation reads or to which the generated definition is bound. Default: vbNullString.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub SetChatGPTBuildContext(ByVal stepText As String, Optional ByVal formName As String = vbNullString, Optional ByVal actionText As String = vbNullString, Optional ByVal detailText As String = vbNullString, Optional ByVal sourceName As String = vbNullString)
    mChatGPTBuildStep = stepText
    mChatGPTBuildForm = formName
    mChatGPTBuildAction = actionText
    mChatGPTBuildDetail = detailText
    mChatGPTBuildSourceName = sourceName
    mChatGPTBuildFieldName = vbNullString
    mChatGPTBuildControlName = vbNullString
    mChatGPTBuildSectionName = vbNullString
    mChatGPTBuildParentName = vbNullString
    mChatGPTBuildBounds = vbNullString
End Sub

' Purpose: Invalidates the generated-form signature-currentness cache without exposing its mutable dictionary to other modules.
' Public role: Inter-module public API. UI orchestration calls this routine when a form-state change invalidates cached generated-form currentness.
' Parameters: None.
' Returns: None (Sub). The cache is released so the next currentness check rebuilds it from authoritative form state.
Public Sub InvalidateGeneratedFormSignatureCurrentCache()
    Set mChatGPTGeneratedFormSignatureCurrentCache = Nothing
End Sub

' Purpose: Returns the read-only generated-form build context needed by progress logging without exposing mutable module state.
' Public role: Inter-module public API. UI progress code reads the current form and detail text through this routine.
' Parameters:
'   formName (Output/in-out; String): Receives the current generated-form build target, or vbNullString when no form is active.
'   detailText (Output/in-out; String): Receives the current generated-form build detail text, or vbNullString when no detail is active.
' Returns: None (Sub). Values are copied from UI_FORMS-owned diagnostic state.
Public Sub GetChatGPTBuildProgressContext(ByRef formName As String, ByRef detailText As String)
    formName = mChatGPTBuildForm
    detailText = mChatGPTBuildDetail
End Sub

Private Function ChatGPTCollectionContainsText(ByVal values As Collection, ByVal targetText As String) As Boolean
    Dim item As Variant
    If values Is Nothing Then Exit Function
    For Each item In values
        SetupControlPump
        If StrComp(CStr(item), targetText, vbTextCompare) = 0 Then
            ChatGPTCollectionContainsText = True
            Exit Function
        End If
    Next item
End Function

Private Sub ClearChatGPTBuildContext()
    mChatGPTBuildStep = vbNullString
    mChatGPTBuildForm = vbNullString
    mChatGPTBuildAction = vbNullString
    mChatGPTBuildDetail = vbNullString
    mChatGPTBuildSourceName = vbNullString
    mChatGPTBuildFieldName = vbNullString
    mChatGPTBuildControlName = vbNullString
    mChatGPTBuildSectionName = vbNullString
    mChatGPTBuildParentName = vbNullString
    mChatGPTBuildBounds = vbNullString
    Set mChatGPTGeneratedFormsRequiringBuild = Nothing
End Sub

' ===============================================================================================================================================================================
' GENERIC GENERATED-FORM CONSTRUCTION / VALIDATION PROCEDURES
' ===============================================================================================================================================================================

Private Function ExistingLookupOrderByText(ByVal sourceName As String, ByVal orderFields As Variant) As String
    Dim fieldName As Variant
    Dim orderText As String
    For Each fieldName In orderFields
        SetupControlPump
        If SharedDao_DatabaseFieldExistsInSource(GetHostDb(), sourceName, CStr(fieldName)) Then
            If Len(orderText) > 0 Then orderText = orderText & ", "
            orderText = orderText & SortableFieldOrderExpression(sourceName, CStr(fieldName))
        End If
    Next fieldName
    If Len(orderText) > 0 Then ExistingLookupOrderByText = " ORDER BY " & orderText
End Function


Private Sub SetChatGPTControlBuildContext(ByVal formName As String, ByVal sourceName As String, ByVal fieldName As String, ByVal controlName As String, ByVal sectionConstant As AcSection, ByVal parentName As String, ByVal boundsText As String, ByVal actionText As String)
    Dim detailText As String
    detailText = "Source: " & sourceName & "; Field: " & fieldName & "; Control: " & controlName
    detailText = detailText & "; Section: " & SharedFormDesign_SectionDisplayName(sectionConstant)
    detailText = detailText & "; Parent/Page: " & NullDisplayText(parentName) & "; Bounds: " & boundsText
    SetChatGPTBuildContext mChatGPTBuildStep, formName, actionText, detailText
    mChatGPTBuildSourceName = sourceName
    mChatGPTBuildFieldName = fieldName
    mChatGPTBuildControlName = controlName
    mChatGPTBuildSectionName = SharedFormDesign_SectionDisplayName(sectionConstant)
    mChatGPTBuildParentName = parentName
    mChatGPTBuildBounds = boundsText
End Sub

' Purpose: Performs chat GPT help create text box by coordinating `CreateControl` and `End` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   formName (Input; String): Physical Access form name identifying the form to create, open, configure, validate, or inspect.
'   controlName (Input; String): Exact Access control name to create, find, configure, or use as the callback/binding target.
'   sectionConstant (Input; AcSection): Access form section constant selecting the section in which the textbox is created (for example Detail, Form Header, or Form Footer).
'   parentName (Input; String): Name of the parent container/control used when creating or locating the generated child control.
'   leftPos (Input; Long): Numeric position locating the left pos item within its source, UI layout, or Navigation Pane ordering.
'   topPos (Input; Long): Numeric position locating the top pos item within its source, UI layout, or Navigation Pane ordering.
'   widthValue (Input; Long): Requested Access control/layout width, in the units expected by the generated-form/report builder.
'   heightValue (Input; Long): Requested Access control/layout height, in the units expected by the generated-form/report builder.
'   lockedValue (Input; Boolean): True configures the generated control as locked/read-only; False leaves the control editable subject to the form policy.
'   multilineValue (Input; Boolean): True configures the generated textbox for multi-line content; False uses the single-line control behavior.
' Returns: Control: Returns the Access textbox control created/configured on the specified Help form section; errors are raised if the control cannot be created as required.
Public Function ChatGPTHelpCreateTextBox(ByVal formName As String, ByVal controlName As String, ByVal sectionConstant As AcSection, ByVal parentName As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal widthValue As Long, ByVal heightValue As Long, ByVal lockedValue As Boolean, ByVal multilineValue As Boolean) As control
    Dim ctl As control
    Set ctl = CreateControl(formName, acTextBox, sectionConstant, parentName, vbNullString, leftPos, topPos, widthValue, heightValue)
    ctl.Name = controlName
    ctl.Locked = lockedValue
    ctl.TabStop = Not lockedValue
    ctl.fontSize = DEFAULT_UI_READING_FONT_SIZE
    If multilineValue Then
        ctl.ScrollBars = 2
        ctl.EnterKeyBehavior = True
    End If
    Set ChatGPTHelpCreateTextBox = ctl
End Function

Private Function BuildExactGeneratedListSelectSql(ByVal formName As String, ByVal sourceName As String, ByVal fieldNames As Variant, ByVal orderFields As Variant, Optional ByVal whereText As String = vbNullString) As String
    Dim selectText As String
    Dim orderText As String
    Dim fieldItem As Variant
    Dim fieldName As String
    Dim exactFields As Collection
    Set exactFields = ExactGeneratedControlFieldCollection(formName, sourceName, fieldNames)
    For Each fieldItem In exactFields
        If Len(selectText) > 0 Then selectText = selectText & ","
        selectText = selectText & SharedDao_BracketName(CStr(fieldItem))
    Next fieldItem
    If IsArray(orderFields) Then
        For Each fieldItem In orderFields
            fieldName = Trim$(CStr(fieldItem))
            If Len(fieldName) = 0 Then Err.Raise vbObjectError + 7381, "BuildExactGeneratedListSelectSql", "The exact generated-form order definition contains a blank field." & vbCrLf & "Form: " & formName & vbCrLf & "Source: " & sourceName
            If SharedDao_DatabaseFieldExistsInSource(GetHostDb(), sourceName, fieldName) = False Then Err.Raise vbObjectError + 7382, "BuildExactGeneratedListSelectSql", "An exact generated-form order field does not exist in the authoritative source." & vbCrLf & "Form: " & formName & vbCrLf & "Source: " & sourceName & vbCrLf & "Field: " & fieldName
            If Len(orderText) > 0 Then orderText = orderText & ","
            orderText = orderText & SortableFieldOrderExpression(sourceName, fieldName)
        Next fieldItem
    End If
    BuildExactGeneratedListSelectSql = "SELECT " & selectText & " FROM " & SharedDao_BracketName(sourceName)
    If Len(Trim$(whereText)) > 0 Then BuildExactGeneratedListSelectSql = BuildExactGeneratedListSelectSql & " WHERE " & whereText
    If Len(orderText) > 0 Then BuildExactGeneratedListSelectSql = BuildExactGeneratedListSelectSql & " ORDER BY " & orderText
    BuildExactGeneratedListSelectSql = BuildExactGeneratedListSelectSql & ";"
End Function

Private Function TableOrQueryExists(ByVal objectName As String) As Boolean
    If Len(objectName) = 0 Then Exit Function
    If mChatGPTGeneratedFormObjectExistsCache Is Nothing Then
        Set mChatGPTGeneratedFormObjectExistsCache = CreateObject("Scripting.Dictionary")
        mChatGPTGeneratedFormObjectExistsCache.CompareMode = vbTextCompare
    End If
    If mChatGPTGeneratedFormObjectExistsCache.Exists(objectName) Then
        mChatGPTGeneratedFormObjectExistsCacheHits = mChatGPTGeneratedFormObjectExistsCacheHits + 1
        TableOrQueryExists = CBool(mChatGPTGeneratedFormObjectExistsCache(objectName))
        Exit Function
    End If
    mChatGPTGeneratedFormObjectExistsCacheMisses = mChatGPTGeneratedFormObjectExistsCacheMisses + 1
    TableOrQueryExists = SharedDao_TableExists(GetHostDb(), objectName) Or SharedDao_QueryExists(GetHostDb(), objectName)
    mChatGPTGeneratedFormObjectExistsCache(objectName) = TableOrQueryExists
End Function

Private Function OrderFieldsRequireSortableProjection(ByVal sourceName As String, ByVal orderFields As Variant) As Boolean
    Dim fieldItem As Variant
    Dim fieldName As String
    If IsArray(orderFields) = False Then Exit Function
    For Each fieldItem In orderFields
        fieldName = Trim$(CStr(fieldItem))
        If Len(fieldName) > 0 Then
            If StrComp(SortableFieldOrderExpression(sourceName, fieldName), SharedDao_BracketName(fieldName), vbBinaryCompare) <> 0 Then
                OrderFieldsRequireSortableProjection = True
                Exit Function
            End If
        End If
    Next fieldItem
End Function

Private Function BuildSelectSql(ByVal sourceName As String, ByVal preferredFields As Variant, ByVal whereText As String, ByVal orderFields As Variant) As String
    Dim omittedWhereReason As String
    Dim sqlText As String
    Dim orderText As String
    If OrderFieldsRequireSortableProjection(sourceName, orderFields) Then
        'The shared builder expects its order-field contract in array form; this branch appends the canonical sortable ORDER BY after the unordered base SELECT is built.
        sqlText = SharedDao_DatabaseBuildSelectSql(GetHostDb(), sourceName, preferredFields, whereText, Array(vbNullString), omittedWhereReason)
        orderText = ExistingLookupOrderByText(sourceName, orderFields)
        If Len(orderText) > 0 And Len(sqlText) > 0 Then
            sqlText = Trim$(sqlText)
            If Right$(sqlText, 1) = ";" Then sqlText = Left$(sqlText, Len(sqlText) - 1)
            sqlText = sqlText & orderText & ";"
        End If
        BuildSelectSql = sqlText
    Else
        BuildSelectSql = SharedDao_DatabaseBuildSelectSql(GetHostDb(), sourceName, preferredFields, whereText, orderFields, omittedWhereReason)
    End If
    If Len(omittedWhereReason) > 0 Then SetChatGPTBuildContext mChatGPTBuildStep, mChatGPTBuildForm, "Build SELECT SQL", omittedWhereReason
End Function

' Purpose: Resolves and returns continuous list control definitions from the supplied inputs and the authoritative ADC definitions/state used by this module.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   formName (Input; String): Physical Access form name identifying the form to create, open, configure, validate, or inspect.
'   sourceName (Input; String): Name of the source object/query/table/control from which this operation reads or to which the generated definition is bound.
'   preferredFields (Input; Variant): Ordered array/list of preferred field names used to choose and order controls/columns when the generated UI is built.
'   useExactPreferredFields (Input; Boolean): True selects/uses exact preferred fields; False follows the routine's alternate or default source/path.
' Returns: Collection: Returns the collection/object containing the canonical continuous list control definitions selected by the supplied filters/context; an empty collection/object is used where the
'   implementation defines no matches.
Public Function ContinuousListControlDefinitions(ByVal formName As String, ByVal sourceName As String, ByVal preferredFields As Variant, ByVal useExactPreferredFields As Boolean) As Collection
    ' INTENTIONAL FILTERED/TRANSFORMED PROJECTION: rows derives only the controls applicable to the selected source fields and adds UI lookup/width/locking metadata; the source schema and Defaults definitions remain
    ' authoritative.
    Dim rows As New Collection
    Dim fields As Collection
    Dim seenControlNames As Object
    Dim fieldItem As Variant
    Dim fieldName As String
    Dim captionText As String
    Dim widthValue As Long
    Dim controlTypeValue As Long
    Dim lookupDefinition As Variant
    Dim valueListText As String
    Dim lockedValue As Boolean
    Dim projectedControlName As String
    Dim expectedFields As Variant
    Dim db As DAO.Database
    Dim sourceFieldDefinitions As Object
    Set db = GetHostDb()
    If SharedDao_TableExists(db, sourceName) Then
        Set sourceFieldDefinitions = db.TableDefs(sourceName).fields
    ElseIf SharedDao_QueryExists(db, sourceName) Then
        Set sourceFieldDefinitions = db.QueryDefs(sourceName).fields
    End If
    If useExactPreferredFields Then
        Set fields = ExactGeneratedControlFieldCollection(formName, sourceName, preferredFields)
    Else
        Set fields = SharedDao_DatabaseExistingFields(GetHostDb(), sourceName, preferredFields)
    End If
    ' INTENTIONAL KEYED LOOKUP: seenControlNames provides keyed membership/lookup over source rows or state already read by ContinuousListControlDefinitions; it is a derived access path, not a second authority.
    Set seenControlNames = CreateObject("Scripting.Dictionary")
    seenControlNames.CompareMode = vbTextCompare
    For Each fieldItem In fields
        SetupControlPump
        fieldName = Trim$(CStr(fieldItem))
        captionText = DisplayCaption(fieldName)
        widthValue = ListColumnWidth(fieldName)
        If Not sourceFieldDefinitions Is Nothing Then
            If sourceFieldDefinitions.item(fieldName).Type = dbDate And widthValue < DEFAULT_UI_CONTROL_WIDTH Then widthValue = DEFAULT_UI_CONTROL_WIDTH
        End If
        If widthValue > (DEFAULT_UI_LIST_FORM_RIGHT_EDGE - DEFAULT_UI_LIST_LEFT_MARGIN) Then widthValue = DEFAULT_UI_LIST_FORM_RIGHT_EDGE - 180
        controlTypeValue = ControlTypeForField(fieldName)
        lookupDefinition = Empty
        If controlTypeValue = acComboBox Then
            lookupDefinition = GeneratedFormLookupDefinition(fieldName, sourceName)
            If IsArray(lookupDefinition) = False Then
                valueListText = ValueListForField(fieldName)
                If Len(valueListText) = 0 Then
                    controlTypeValue = acTextBox
                Else
                    lookupDefinition = Array(GENERATED_FORM_LOOKUP_KIND_VALUE_LIST, valueListText, vbNullString, vbNullString, vbNullString, Empty, vbNullString, False)
                End If
            End If
        End If
        lockedValue = ShouldLockField(fieldName, False)
        If LookupEditorCodeFieldLocked(formName, fieldName) Then lockedValue = True
        projectedControlName = SharedFormDesign_SafeControlName("lst_" & fieldName)
        If Len(projectedControlName) = 0 Then Err.Raise vbObjectError + 7406, "ContinuousListControlDefinitions", "A generated list-control definition produced a blank control name." & vbCrLf & "Form: " & formName & vbCrLf & "Source: " & sourceName & vbCrLf & "Field: " & fieldName
        If seenControlNames.Exists(projectedControlName) Then Err.Raise vbObjectError + 7407, "ContinuousListControlDefinitions", "Generated list-control definitions produce duplicate control names." & vbCrLf & "Form: " & formName & vbCrLf & "Source: " & sourceName & vbCrLf & "Control: " & projectedControlName
        If widthValue <= 0 Then Err.Raise vbObjectError + 7408, "ContinuousListControlDefinitions", "A generated list-control definition has a nonpositive width." & vbCrLf & "Form: " & formName & vbCrLf & "Source: " & sourceName & vbCrLf & "Field: " & fieldName
        seenControlNames.Add projectedControlName, True
        rows.Add Array(fieldName, captionText, widthValue, controlTypeValue, lockedValue, projectedControlName, CLng(acDetail), lookupDefinition)
    Next fieldItem
    If useExactPreferredFields Then
        expectedFields = preferredFields
        If rows.count <> UBound(expectedFields) - LBound(expectedFields) + 1 Then Err.Raise vbObjectError + 7409, "ContinuousListControlDefinitions", "The authoritative exact control-definition count is incomplete." & vbCrLf & "Form: " & formName & vbCrLf & "Source: " & sourceName
    End If
    Set ContinuousListControlDefinitions = rows
End Function

Private Function ExactGeneratedControlFieldCollection(ByVal formName As String, ByVal sourceName As String, ByVal fieldNames As Variant) As Collection
    ' INTENTIONAL FILTERED PARTITION: result is a consumer-required subset/projection derived from the source enumerated by ExactGeneratedControlFieldCollection; its membership is not independently authoritative.
    ' INTENTIONAL KEYED LOOKUP: seen provides keyed membership/lookup over source rows or state already read by ExactGeneratedControlFieldCollection; it is a derived access path, not a second authority.
    Dim result As New Collection
    Dim seen As Object
    Dim fieldItem As Variant
    Dim fieldName As String
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    If IsArray(fieldNames) = False Then Err.Raise vbObjectError + 7376, "ExactGeneratedControlFieldCollection", "The exact generated-control field definition is not an array." & vbCrLf & "Form: " & formName & vbCrLf & "Source: " & sourceName
    For Each fieldItem In fieldNames
        SetupControlPump
        fieldName = Trim$(CStr(fieldItem))
        If Len(fieldName) = 0 Then Err.Raise vbObjectError + 7377, "ExactGeneratedControlFieldCollection", "The exact generated-control field definition contains a blank field." & vbCrLf & "Form: " & formName & vbCrLf & "Source: " & sourceName
        If seen.Exists(fieldName) Then Err.Raise vbObjectError + 7378, "ExactGeneratedControlFieldCollection", "The exact generated-control field definition contains a duplicate field." & vbCrLf & "Form: " & formName & vbCrLf & "Source: " & sourceName & vbCrLf & "Field: " & fieldName
        If SharedDao_DatabaseFieldExistsInSource(GetHostDb(), sourceName, fieldName) = False Then Err.Raise vbObjectError + 7379, "ExactGeneratedControlFieldCollection", "The exact generated-control field does not exist in the authoritative source." & vbCrLf & "Form: " & formName & vbCrLf & "Source: " & sourceName & vbCrLf & "Field: " & fieldName
        seen.Add fieldName, True
        result.Add fieldName
    Next fieldItem
    If result.count = 0 Then Err.Raise vbObjectError + 7380, "ExactGeneratedControlFieldCollection", "The exact generated-control field definition is empty." & vbCrLf & "Form: " & formName & vbCrLf & "Source: " & sourceName
    Set ExactGeneratedControlFieldCollection = result
End Function

Private Sub AddFieldBlock(ByVal formName As String, ByVal sourceName As String, ByVal fieldList As Variant, ByVal sectionConstant As AcSection, ByVal parentName As String, ByVal leftStart As Long, ByVal topStart As Long, ByVal columnCount As Long, ByVal controlPrefix As String, ByVal headerMode As Boolean)
    On Error GoTo ErrHandler
    Dim existing As Collection
    Dim fieldName As Variant
    Dim colIndex As Long
    Dim leftPos As Long
    Dim topPos As Long
    Dim rowTop() As Long
    Dim colWidth As Long
    Dim ctlType As AcControlType
    Dim ControlHeight As Long
    Dim ControlWidth As Long
    Dim occupiedHeight As Long
    Dim ctl As control
    Dim blockDetail As String
    blockDetail = "Source: " & sourceName & "; Section: " & SharedFormDesign_SectionDisplayName(sectionConstant)
    blockDetail = blockDetail & "; Parent/Page: " & NullDisplayText(parentName) & "; Prefix: " & controlPrefix
    blockDetail = blockDetail & "; Start: " & CStr(leftStart) & "," & CStr(topStart) & "; Columns: " & CStr(columnCount) & "; HeaderMode: " & CStr(headerMode)
    SetChatGPTBuildContext mChatGPTBuildStep, formName, "Resolve fields for generated field block.", blockDetail
    Set existing = SharedDao_DatabaseExistingFields(GetHostDb(), sourceName, fieldList)
    If existing.count = 0 Then Exit Sub
    If columnCount < 1 Then columnCount = 1
    ReDim rowTop(0 To columnCount - 1)
    colWidth = DEFAULT_UI_LABEL_WIDTH + DEFAULT_UI_CONTROL_WIDTH + DEFAULT_UI_COLUMN_GAP
    For colIndex = 0 To columnCount - 1
        SetupControlPump
        rowTop(colIndex) = topStart
    Next colIndex
    colIndex = 0
    For Each fieldName In existing
        SetupControlPump
        ctlType = ControlTypeForField(CStr(fieldName))
        ControlHeight = HeightForField(CStr(fieldName), headerMode)
        ControlWidth = WidthForField(CStr(fieldName), headerMode)
        occupiedHeight = ControlHeight
        If DEFAULT_UI_LABEL_HEIGHT + 30 > occupiedHeight Then occupiedHeight = DEFAULT_UI_LABEL_HEIGHT + 30
        If ctlType = acCheckBox And DEFAULT_UI_CONTROL_HEIGHT + 30 > occupiedHeight Then occupiedHeight = DEFAULT_UI_CONTROL_HEIGHT + 30
        If IsMultilineField(CStr(fieldName)) And headerMode = False Then
            colIndex = 0
            leftPos = leftStart
            topPos = MaxRowTop(rowTop)
            AddSingleFieldControl formName, sourceName, CStr(fieldName), ctlType, sectionConstant, parentName, leftPos, topPos, DEFAULT_UI_WIDE_CONTROL_WIDTH, ControlHeight, controlPrefix, headerMode, ctl
            SetAllRowTops rowTop, topPos + occupiedHeight + DEFAULT_UI_ROW_GAP
        Else
            leftPos = leftStart + (colIndex * colWidth)
            topPos = rowTop(colIndex)
            AddSingleFieldControl formName, sourceName, CStr(fieldName), ctlType, sectionConstant, parentName, leftPos, topPos, ControlWidth, ControlHeight, controlPrefix, headerMode, ctl
            rowTop(colIndex) = topPos + occupiedHeight + DEFAULT_UI_ROW_GAP
            colIndex = colIndex + 1
            If colIndex >= columnCount Then colIndex = 0
        End If
    Next fieldName
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If Len(Trim$(savedErrSource)) = 0 Then savedErrSource = "AddFieldBlock"
    Err.Raise savedErrNumber, savedErrSource, "Could not create a generated-form field block." & vbCrLf & blockDetail & vbCrLf & "Current field: " & NullDisplayText(CStr(fieldName)) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Original description: " & savedErrDescription
End Sub

Private Sub AddSingleFieldControl(ByVal formName As String, ByVal sourceName As String, ByVal fieldName As String, ByVal ctlType As AcControlType, ByVal sectionConstant As AcSection, ByVal parentName As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal ControlWidth As Long, ByVal ControlHeight As Long, ByVal controlPrefix As String, ByVal headerMode As Boolean, ByRef ctl As control)
    On Error GoTo ErrHandler
    Dim labelText As String
    Dim labelWidthValue As Long
    Dim controlLeft As Long
    Dim controlTop As Long
    Dim controlName As String
    Dim labelBounds As String
    Dim controlBounds As String
    labelText = DisplayCaption(fieldName)
    labelWidthValue = DEFAULT_UI_LABEL_WIDTH
    controlLeft = leftPos + labelWidthValue
    controlTop = topPos
    controlName = SharedFormDesign_SafeControlName(controlPrefix & "_" & fieldName)
    labelBounds = BoundsForMessage(leftPos, topPos + 30, labelWidthValue - DEFAULT_UI_LABEL_CONTROL_GAP, DEFAULT_UI_LABEL_HEIGHT)
    If ctlType = acCheckBox Then
        controlBounds = BoundsForMessage(controlLeft, controlTop + 30, 360, DEFAULT_UI_CONTROL_HEIGHT)
    Else
        controlBounds = BoundsForMessage(controlLeft, controlTop, ControlWidth, ControlHeight)
    End If
    SetChatGPTControlBuildContext formName, sourceName, fieldName, controlName, sectionConstant, parentName, labelBounds, "Create generated field label."
    With SharedFormDesign_AddPlainLabel(formName, sectionConstant, parentName, labelText, leftPos, topPos + 30, labelWidthValue - DEFAULT_UI_LABEL_CONTROL_GAP, DEFAULT_UI_LABEL_HEIGHT, DEFAULT_UI_FONT_SIZE)
        .TextAlign = 1
    End With
    SetChatGPTControlBuildContext formName, sourceName, fieldName, controlName, sectionConstant, parentName, controlBounds, "Create generated bound field control."
    If ctlType = acCheckBox Then
        Set ctl = SharedFormDesign_AddBoundControl(formName, sourceName, fieldName, ctlType, sectionConstant, parentName, controlLeft, controlTop + 30, 360, DEFAULT_UI_CONTROL_HEIGHT, controlPrefix)
    Else
        Set ctl = SharedFormDesign_AddBoundControl(formName, sourceName, fieldName, ctlType, sectionConstant, parentName, controlLeft, controlTop, ControlWidth, ControlHeight, controlPrefix)
    End If
    If ctlType = acComboBox Then
        SetChatGPTControlBuildContext formName, sourceName, fieldName, controlName, sectionConstant, parentName, controlBounds, "Configure generated combo-box lookup."
        ConfigureComboBox ctl, fieldName, sourceName
    End If
    If IsMultilineField(fieldName) Then
        SetChatGPTControlBuildContext formName, sourceName, fieldName, controlName, sectionConstant, parentName, controlBounds, "Configure generated multiline text box."
        SharedFormDesign_ConfigureMultilineTextBox ctl
    End If
    If ShouldLockField(fieldName, headerMode) Then
        SetChatGPTControlBuildContext formName, sourceName, fieldName, controlName, sectionConstant, parentName, controlBounds, "Apply generated locked-control style."
        SharedFormDesign_ApplyLockedControlStyle ctl, DEFAULT_UI_COLOR_LOCKED_BACK
    End If
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim failureText As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If Len(Trim$(savedErrSource)) = 0 Then savedErrSource = "AddSingleFieldControl"
    failureText = "Could not create or configure a generated field control."
    failureText = failureText & vbCrLf & "Form: " & formName & "; Source: " & sourceName
    failureText = failureText & vbCrLf & "Field: " & fieldName & "; Control: " & controlName
    failureText = failureText & vbCrLf & "Control type: " & SharedFormDesign_ControlTypeDisplayName(ctlType)
    failureText = failureText & vbCrLf & "Section: " & SharedFormDesign_SectionDisplayName(sectionConstant) & "; Parent/Page: " & NullDisplayText(parentName)
    failureText = failureText & vbCrLf & "Label bounds: " & labelBounds & "; Control bounds: " & controlBounds
    failureText = failureText & vbCrLf & "Build action: " & NullDisplayText(mChatGPTBuildAction)
    failureText = failureText & vbCrLf & "Original source: " & savedErrSource
    failureText = failureText & vbCrLf & "Original description: " & savedErrDescription
    Err.Raise savedErrNumber, savedErrSource, failureText
End Sub

Private Function ConfigureDistinctFieldLookup(ByVal ctl As control, ByVal tableName As String, ByVal fieldName As String, Optional ByVal columnWidthsText As String = "2.5;2.5") As Boolean
    On Error GoTo ErrHandler
    Dim sqlText As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    If SharedDao_TableExists(GetHostDb(), tableName) = False Then Exit Function
    If SharedDao_TableFieldExists(GetHostDb(), tableName, fieldName) = False Then Exit Function
    sqlText = "SELECT DISTINCT [" & fieldName & "], [" & fieldName & "] FROM " & SharedDao_BracketName(tableName) & " WHERE Len(Nz([" & fieldName & "],''))>0 ORDER BY [" & fieldName & "];"
    SharedFormDesign_ConfigureTwoColumnLookup ctl, sqlText
    ctl.columnWidths = columnWidthsText
    ConfigureDistinctFieldLookup = True
    Exit Function
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ConfigureDistinctFieldLookup", "Could not configure the generated distinct-field lookup." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function ExistingFieldCountForBuild(ByVal sourceName As String, ByVal fieldList As Variant) As Long
    On Error GoTo ErrHandler
    Dim existing As Collection
    Set existing = SharedDao_DatabaseExistingFields(GetHostDb(), sourceName, fieldList)
    ExistingFieldCountForBuild = existing.count
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If Len(Trim$(savedErrSource)) = 0 Then savedErrSource = "ExistingFieldCountForBuild"
    Err.Raise savedErrNumber, savedErrSource, "Could not count existing generated-form source fields." & vbCrLf & "Source: " & sourceName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Original description: " & savedErrDescription
End Function

Private Function RequiredFieldBlockSectionHeight(ByVal sourceName As String, ByVal fieldList As Variant, ByVal topStart As Long, ByVal columnCount As Long, ByVal headerMode As Boolean) As Long
    On Error GoTo ErrHandler
    Dim existing As Collection
    Dim fieldName As Variant
    Dim colIndex As Long
    Dim topPos As Long
    Dim rowTop() As Long
    Dim fieldHeight As Long
    Dim ctlType As AcControlType
    Dim itemHeight As Long
    Dim itemBottom As Long
    Dim maxBottom As Long
    Set existing = SharedDao_DatabaseExistingFields(GetHostDb(), sourceName, fieldList)
    If columnCount < 1 Then columnCount = 1
    ReDim rowTop(0 To columnCount - 1)
    For colIndex = 0 To columnCount - 1
        SetupControlPump
        rowTop(colIndex) = topStart
    Next colIndex
    colIndex = 0
    maxBottom = topStart
    For Each fieldName In existing
        SetupControlPump
        fieldHeight = HeightForField(CStr(fieldName), headerMode)
        ctlType = ControlTypeForField(CStr(fieldName))
        itemHeight = fieldHeight
        If DEFAULT_UI_LABEL_HEIGHT + 30 > itemHeight Then itemHeight = DEFAULT_UI_LABEL_HEIGHT + 30
        If ctlType = acCheckBox And DEFAULT_UI_CONTROL_HEIGHT + 30 > itemHeight Then itemHeight = DEFAULT_UI_CONTROL_HEIGHT + 30
        If IsMultilineField(CStr(fieldName)) And headerMode = False Then
            colIndex = 0
            topPos = MaxRowTop(rowTop)
            itemBottom = topPos + itemHeight
            If itemBottom > maxBottom Then maxBottom = itemBottom
            SetAllRowTops rowTop, topPos + itemHeight + DEFAULT_UI_ROW_GAP
        Else
            topPos = rowTop(colIndex)
            itemBottom = topPos + itemHeight
            If itemBottom > maxBottom Then maxBottom = itemBottom
            rowTop(colIndex) = topPos + itemHeight + DEFAULT_UI_ROW_GAP
            colIndex = colIndex + 1
            If colIndex >= columnCount Then colIndex = 0
        End If
    Next fieldName
    RequiredFieldBlockSectionHeight = maxBottom + DEFAULT_UI_FORM_SECTION_BOTTOM_MARGIN
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If Len(Trim$(savedErrSource)) = 0 Then savedErrSource = "RequiredFieldBlockSectionHeight"
    Err.Raise savedErrNumber, savedErrSource, "Could not calculate the required generated-form field-block height." & vbCrLf & "Source: " & sourceName & vbCrLf & "Top: " & CStr(topStart) & vbCrLf & "Columns: " & CStr(columnCount) & vbCrLf & "Header mode: " & CStr(headerMode) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Original description: " & savedErrDescription
End Function

Private Function GeneratedFormCompleteSetTokenPropertyName(ByVal normalizedBuildScope As String) As String
    Select Case UCase$(normalizedBuildScope)
        Case UCase$(FORM_BUILD_SCOPE_ALL): GeneratedFormCompleteSetTokenPropertyName = GENERATED_FORM_COMPLETE_SET_TOKEN_PROPERTY
        Case UCase$(FORM_BUILD_SCOPE_CHATGPT_UI_FORMS): GeneratedFormCompleteSetTokenPropertyName = GENERATED_FORM_CHATGPT_UI_COMPLETE_SET_TOKEN_PROPERTY
        Case UCase$(FORM_BUILD_SCOPE_SUPPORT_FORMS): GeneratedFormCompleteSetTokenPropertyName = GENERATED_FORM_SUPPORT_COMPLETE_SET_TOKEN_PROPERTY
        Case UCase$(FORM_BUILD_SCOPE_LOOKUP_EDITORS): GeneratedFormCompleteSetTokenPropertyName = GENERATED_FORM_LOOKUP_EDITORS_COMPLETE_SET_TOKEN_PROPERTY
        Case Else: Err.Raise vbObjectError + 6767, "GeneratedFormCompleteSetTokenPropertyName", "Unsupported generated-form proof scope: " & normalizedBuildScope
    End Select
End Function

Private Function GeneratedFormPhysicalStateTokenPropertyName(ByVal normalizedBuildScope As String) As String
    Select Case UCase$(normalizedBuildScope)
        Case UCase$(FORM_BUILD_SCOPE_ALL): GeneratedFormPhysicalStateTokenPropertyName = GENERATED_FORM_PHYSICAL_STATE_TOKEN_PROPERTY
        Case UCase$(FORM_BUILD_SCOPE_CHATGPT_UI_FORMS): GeneratedFormPhysicalStateTokenPropertyName = GENERATED_FORM_CHATGPT_UI_PHYSICAL_STATE_TOKEN_PROPERTY
        Case UCase$(FORM_BUILD_SCOPE_SUPPORT_FORMS): GeneratedFormPhysicalStateTokenPropertyName = GENERATED_FORM_SUPPORT_PHYSICAL_STATE_TOKEN_PROPERTY
        Case UCase$(FORM_BUILD_SCOPE_LOOKUP_EDITORS): GeneratedFormPhysicalStateTokenPropertyName = GENERATED_FORM_LOOKUP_EDITORS_PHYSICAL_STATE_TOKEN_PROPERTY
        Case Else: Err.Raise vbObjectError + 6768, "GeneratedFormPhysicalStateTokenPropertyName", "Unsupported generated-form proof scope: " & normalizedBuildScope
    End Select
End Function

Private Function GeneratedFormIsInBuildScope(ByVal formName As String, ByVal normalizedBuildScope As String) As Boolean
    Select Case UCase$(normalizedBuildScope)
        Case UCase$(FORM_BUILD_SCOPE_ALL)
            GeneratedFormIsInBuildScope = True
        Case UCase$(FORM_BUILD_SCOPE_CHATGPT_UI_FORMS)
            GeneratedFormIsInBuildScope = GeneratedFormChatGPTUIClassification(formName)
        Case UCase$(FORM_BUILD_SCOPE_SUPPORT_FORMS)
            GeneratedFormIsInBuildScope = Not GeneratedFormChatGPTUIClassification(formName)
        Case UCase$(FORM_BUILD_SCOPE_LOOKUP_EDITORS)
            GeneratedFormIsInBuildScope = LookupEditorGeneratedFormIsDefined(formName)
    End Select
End Function

Private Function GeneratedFormNamesForBuildScope(ByVal normalizedBuildScope As String) As Collection
    Dim result As New Collection
    Dim formNames As Variant
    Dim formItem As Variant
    Dim formName As String
    formNames = GeneratedFormNames()
    For Each formItem In formNames
        formName = CStr(formItem)
        If GeneratedFormIsInBuildScope(formName, normalizedBuildScope) Then result.Add formName
    Next formItem
    Set GeneratedFormNamesForBuildScope = result
End Function

Private Function GeneratedFormBuildNamesForScope(ByVal normalizedBuildScope As String) As Collection
    Dim result As New Collection
    Dim formNames As Collection
    Dim formItem As Variant
    Set formNames = BuildGeneratedFormNameList()
    For Each formItem In formNames
        If GeneratedFormIsInBuildScope(CStr(formItem), normalizedBuildScope) Then result.Add CStr(formItem)
    Next formItem
    Set GeneratedFormBuildNamesForScope = result
End Function

Private Function FilterGeneratedFormNamesForScope(ByVal sourceNames As Collection, ByVal normalizedBuildScope As String) As Collection
    Dim result As New Collection
    Dim item As Variant
    If sourceNames Is Nothing Then Set FilterGeneratedFormNamesForScope = result: Exit Function
    For Each item In sourceNames
        If GeneratedFormIsInBuildScope(CStr(item), normalizedBuildScope) Then result.Add CStr(item)
    Next item
    Set FilterGeneratedFormNamesForScope = result
End Function

Private Function GeneratedFormStaticBuildScopeCount(ByVal normalizedBuildScope As String) As Long
    Dim definitionData As Variant
    Dim editorLogicalIDs As Object
    Select Case UCase$(normalizedBuildScope)
        Case UCase$(FORM_BUILD_SCOPE_ALL)
            GeneratedFormStaticBuildScopeCount = CanonicalGeneratedFormNamingDefinitions().count
        Case UCase$(FORM_BUILD_SCOPE_CHATGPT_UI_FORMS), UCase$(FORM_BUILD_SCOPE_SUPPORT_FORMS)
            GeneratedFormStaticBuildScopeCount = GeneratedFormNamesForBuildScope(normalizedBuildScope).count
        Case UCase$(FORM_BUILD_SCOPE_LOOKUP_EDITORS)
            Set editorLogicalIDs = CreateObject("Scripting.Dictionary")
            editorLogicalIDs.CompareMode = vbTextCompare
            For Each definitionData In LookupEditorPolicyDefinitions()
                If StrComp(CStr(definitionData(LOOKUP_EDITOR_DEF_POLICY)), LOOKUP_EDITOR_POLICY_GENERATED, vbTextCompare) = 0 Then
                    If Len(CStr(definitionData(LOOKUP_EDITOR_DEF_EDITOR_LOGICAL_ID))) > 0 Then editorLogicalIDs(CStr(definitionData(LOOKUP_EDITOR_DEF_EDITOR_LOGICAL_ID))) = True
                End If
            Next definitionData
            GeneratedFormStaticBuildScopeCount = editorLogicalIDs.count
    End Select
End Function

Private Function GeneratedFormAuthoritativeCount(Optional ByVal buildScope As String = FORM_BUILD_SCOPE_ALL) As Long
    Dim formNames As Collection
    Set formNames = GeneratedFormNamesForBuildScope(buildScope)
    GeneratedFormAuthoritativeCount = formNames.count
End Function

Private Function GeneratedFormCompactPreflightMatches(ByVal normalizedBuildScope As String, ByRef completeSetToken As String, ByRef physicalStateToken As String, ByRef detailText As String, Optional ByVal suppliedStoredCompleteToken As Variant, Optional ByVal suppliedStoredPhysicalToken As Variant) As Boolean
    On Error GoTo NotCurrent
    Dim db As DAO.Database
    Dim completePropertyExists As Boolean
    Dim physicalPropertyExists As Boolean
    Dim storedCompleteToken As String
    Dim storedPhysicalToken As String
    Set db = GetHostDb()
    If isMissing(suppliedStoredCompleteToken) Then
        storedCompleteToken = DatabaseTextPropertyValue(db, GeneratedFormCompleteSetTokenPropertyName(normalizedBuildScope), completePropertyExists)
        If completePropertyExists = False Or Len(storedCompleteToken) = 0 Then detailText = "Persisted generated-form expected-definition proof is absent; skipped impossible compact comparison.": Exit Function
    Else
        storedCompleteToken = CStr(suppliedStoredCompleteToken)
        If Len(storedCompleteToken) = 0 Then Exit Function
    End If
    If isMissing(suppliedStoredPhysicalToken) Then
        storedPhysicalToken = DatabaseTextPropertyValue(db, GeneratedFormPhysicalStateTokenPropertyName(normalizedBuildScope), physicalPropertyExists)
        If physicalPropertyExists = False Or Len(storedPhysicalToken) = 0 Then detailText = "Persisted generated-form physical proof is absent; skipped impossible compact comparison.": Exit Function
    Else
        storedPhysicalToken = CStr(suppliedStoredPhysicalToken)
        If Len(storedPhysicalToken) = 0 Then Exit Function
    End If
    CorePulseProgress "Generated-form state check: calculating the complete expected-definition token."
    completeSetToken = GeneratedFormCompleteSetToken(normalizedBuildScope)
    If StrComp(storedCompleteToken, completeSetToken, vbBinaryCompare) <> 0 Then Exit Function
    CorePulseProgress "Generated-form expected definitions matched; calculating the compact physical-state token."
    physicalStateToken = GeneratedFormPhysicalStateToken(normalizedBuildScope, detailText)
    If Len(physicalStateToken) = 0 Then Exit Function
    If StrComp(storedPhysicalToken, physicalStateToken, vbBinaryCompare) <> 0 Then Exit Function
    detailText = "Generated-form expected-definition and physical-state tokens matched. Full form-by-form preflight was skipped."
    CorePulseProgress "Generated-form expected-definition and physical-state tokens matched."
    GeneratedFormCompactPreflightMatches = True
    Exit Function
NotCurrent:
    detailText = "Generated-form compact preflight could not prove the set current. Full preflight will run. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    Err.Clear
End Function

Private Function GeneratedFormCompleteSetToken(ByVal normalizedBuildScope As String) As String
    Dim formNames As Collection
    Dim formItem As Variant
    Dim formName As String
    Dim payloadText As String
    payloadText = "ADCGeneratedFormCompleteSetV1|" & FORM_BUILDER_VERSION & "|" & normalizedBuildScope & "|"
    Set formNames = GeneratedFormNamesForBuildScope(normalizedBuildScope)
    For Each formItem In formNames
        SetupControlPump
        formName = CStr(formItem)
        payloadText = payloadText & CStr(Len(formName)) & ":" & formName & "|" & GeneratedFormExpectedBuildSignature(formName) & vbLf
    Next formItem
    GeneratedFormCompleteSetToken = AdlerTextHash(payloadText) & ":" & CStr(Len(payloadText))
End Function

Private Function GeneratedFormCompleteSetTokenMatches(ByVal normalizedBuildScope As String, ByRef currentToken As String) As Boolean
    Dim db As DAO.Database
    Dim storedToken As String
    Dim propertyExists As Boolean
    Set db = GetHostDb()
    storedToken = DatabaseTextPropertyValue(db, GeneratedFormCompleteSetTokenPropertyName(normalizedBuildScope), propertyExists)
    If propertyExists = False Or Len(storedToken) = 0 Then Exit Function
    currentToken = GeneratedFormCompleteSetToken(normalizedBuildScope)
    GeneratedFormCompleteSetTokenMatches = (StrComp(storedToken, currentToken, vbBinaryCompare) = 0)
End Function

Private Function GeneratedFormPhysicalStateToken(ByVal normalizedBuildScope As String, ByRef detailText As String, Optional ByRef discoveredUserFormCount As Variant) As String
    'Compact mutation-free physical gate. During Setup, consume the one Catalog-owned physical-form snapshot; direct/non-Setup builds retain the HostDb scan path. Any missing form or unavailable metadata returns
    'blank and forces exhaustive preflight.
    On Error GoTo NotCurrent
    Dim expectedNames As Object
    Dim formNames As Collection
    Dim formItem As Variant
    Dim formName As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim payloadText As String
    Dim foundCount As Long
    Dim physicalInventory As Collection
    Dim rowData As Variant
    Dim setupInventoryAvailable As Boolean
    If isMissing(discoveredUserFormCount) = False Then discoveredUserFormCount = CLng(0)
    ' INTENTIONAL KEYED LOOKUP: expectedNames provides keyed membership/lookup over source rows or state already read by GeneratedFormPhysicalStateToken; it is a derived access path, not a second authority.
    Set expectedNames = CreateObject("Scripting.Dictionary")
    expectedNames.CompareMode = vbTextCompare
    Set formNames = GeneratedFormNamesForBuildScope(normalizedBuildScope)
    For Each formItem In formNames
        SetupControlPump
        formName = CStr(formItem)
        If expectedNames.Exists(formName) = False Then expectedNames.Add formName, True
    Next formItem
    payloadText = "ADCGeneratedFormPhysicalStateV1|" & FORM_BUILDER_VERSION & "|" & normalizedBuildScope & "|"
    If CodeDbIsHostDb() And TraceRunActive() Then setupInventoryAvailable = TryGetPhysicalManagedFormInventoryForSetup(physicalInventory)
    If setupInventoryAvailable Then
        For Each rowData In physicalInventory
            SetupControlPump
            formName = CStr(rowData(0))
            If expectedNames.Exists(formName) Then
                payloadText = payloadText & CStr(Len(formName)) & ":" & formName & "|" & Format$(CDate(rowData(5)), "yyyy-mm-dd hh:nn:ss") & "|" & Format$(CDate(rowData(6)), "yyyy-mm-dd hh:nn:ss") & "|" & CStr(CLng(Nz(rowData(7), 0))) & vbLf
                foundCount = foundCount + 1
            ElseIf isMissing(discoveredUserFormCount) = False Then
                If StrComp(CStr(rowData(3)), DEFAULT_CATALOG_OWNER_USER, vbTextCompare) = 0 Then discoveredUserFormCount = CLng(discoveredUserFormCount) + 1
            End If
        Next rowData
    Else
        Set db = GetHostDb()
        Set rs = db.OpenRecordset("SELECT [Name],[DateCreate],[DateUpdate],[Flags] FROM [" & MSYS_OBJECTS_TABLE & "] WHERE [Type]=" & CStr(MSYS_OBJECT_TYPE_FORM) & " ORDER BY [Name];", dbOpenSnapshot)
        Do While rs.EOF = False
            SetupControlPump
            formName = CStr(Nz(rs.fields("Name").value, vbNullString))
            If expectedNames.Exists(formName) Then
                payloadText = payloadText & CStr(Len(formName)) & ":" & formName & "|" & Format$(CDate(rs.fields("DateCreate").value), "yyyy-mm-dd hh:nn:ss") & "|" & Format$(CDate(rs.fields("DateUpdate").value), "yyyy-mm-dd hh:nn:ss") & "|" & CStr(Nz(rs.fields("Flags").value, 0)) & vbLf
                foundCount = foundCount + 1
            ElseIf isMissing(discoveredUserFormCount) = False Then
                If PhysicalFormIsDiscoveryCandidate(formName) Then
                    If ManagedGeneratedFormIsDefined(formName) = False And ManagedPersistentGeneratedFormIsDefined(formName) = False Then discoveredUserFormCount = CLng(discoveredUserFormCount) + 1
                End If
            End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
    If foundCount <> expectedNames.count Then
        detailText = "Generated-form compact state found " & CStr(foundCount) & " of " & CStr(expectedNames.count) & " required forms."
        Exit Function
    End If
    GeneratedFormPhysicalStateToken = AdlerTextHash(payloadText) & ":" & CStr(Len(payloadText))
    Exit Function
NotCurrent:
    detailText = "Generated-form compact physical state was unavailable; exhaustive preflight is required. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    On Error GoTo 0
End Function

Private Function GeneratedFormPreflightPlan(ByVal normalizedBuildScope As String) As Collection
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: rows holds planned/recovery work or independently observed validation evidence for GeneratedFormPreflightPlan; keep it distinct from the source authority.
    Dim rows As New Collection
    Dim formNames As Collection
    Dim formItem As Variant
    Dim formName As String
    Dim sourceName As String
    Dim sourceTypeText As String
    Dim requiredControlSources As Variant
    Set formNames = GeneratedFormNamesForBuildScope(normalizedBuildScope)
    For Each formItem In formNames
        SetupControlPump
        formName = CStr(formItem)
        sourceName = GeneratedFormSourceNameForSignature(formName)
        sourceTypeText = GeneratedFormExpectedSourceType(formName)
        requiredControlSources = GeneratedFormPreflightRequiredControlSources(formName)
        rows.Add Array(formName, sourceName, sourceTypeText, requiredControlSources)
    Next formItem
    Set GeneratedFormPreflightPlan = rows
End Function

Private Function GeneratedFormPreflightRequiredControlSources(ByVal formName As String) As Variant
    If AddInReviewChildFormIsDefinedDefault(formName) Then
        GeneratedFormPreflightRequiredControlSources = AddInReviewExpectedControlSources(formName)
    ElseIf LookupEditorGeneratedFormIsDefined(formName) Then
        GeneratedFormPreflightRequiredControlSources = GeneratedFormListFields(formName)
    Else
        GeneratedFormPreflightRequiredControlSources = Empty
    End If
End Function

Private Function GeneratedFormBuildSignatureIsCurrent(ByVal formName As String) As Boolean
    'Uses the deterministic builder/catalog/source signature as the authoritative rebuild gate. The physical hash remains stamped for diagnostics but does not make an unchanged form stale merely because custom
    'document properties changed its exported text.
    On Error GoTo NotCurrent
    Dim signatureExists As Boolean
    Dim storedSignature As String
    If mChatGPTGeneratedFormSignatureCurrentCache Is Nothing Then
        Set mChatGPTGeneratedFormSignatureCurrentCache = CreateObject("Scripting.Dictionary")
        mChatGPTGeneratedFormSignatureCurrentCache.CompareMode = vbTextCompare
    End If
    If mChatGPTGeneratedFormSignatureCurrentCache.Exists(formName) Then
        GeneratedFormBuildSignatureIsCurrent = CBool(mChatGPTGeneratedFormSignatureCurrentCache(formName))
        Exit Function
    End If
    If SharedDao_AccessObjectExists(GetHostDb(), acForm, formName) = False Then GoTo CacheResult
    storedSignature = GeneratedFormDocumentTextProperty(formName, GENERATED_FORM_BUILD_SIGNATURE_PROPERTY, signatureExists)
    If signatureExists = False Then GoTo CacheResult
    If StrComp(storedSignature, GeneratedFormExpectedBuildSignature(formName), vbBinaryCompare) <> 0 Then GoTo CacheResult
    GeneratedFormBuildSignatureIsCurrent = True
CacheResult:
    mChatGPTGeneratedFormSignatureCurrentCache(formName) = GeneratedFormBuildSignatureIsCurrent
    Exit Function
NotCurrent:
    Err.Clear
    If Not mChatGPTGeneratedFormSignatureCurrentCache Is Nothing Then mChatGPTGeneratedFormSignatureCurrentCache(formName) = False
End Function

Private Function GeneratedFormDocumentTextProperty(ByVal formName As String, ByVal propertyName As String, ByRef propertyExists As Boolean) As String
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim docs As DAO.Documents
    Dim doc As DAO.Document
    Dim prp As DAO.Property
    Dim firstAccessErrNumber As Long
    propertyExists = False
    Set db = GetHostDb()
    Set docs = db.Containers(FORM_DAO_CONTAINER_NAME).Documents
    On Error Resume Next
    Set doc = docs(formName)
    firstAccessErrNumber = Err.Number
    Err.Clear
    On Error GoTo ErrHandler
    If doc Is Nothing Then
        'Refresh only after direct document access proves the DAO Documents collection is stale/missing the newly created or replaced form.
        docs.Refresh
        Set doc = docs(formName)
    End If
    For Each prp In doc.properties
        SetupControlPump
        If StrComp(prp.Name, propertyName, vbTextCompare) = 0 Then
            propertyExists = True
            If IsNull(prp.value) = False Then GeneratedFormDocumentTextProperty = CStr(prp.value)
            Exit Function
        End If
    Next prp
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "GeneratedFormDocumentTextProperty", "Could not read generated-form signature metadata." & vbCrLf & "Form: " & formName & vbCrLf & "Property: " & propertyName & vbCrLf & "Initial direct-access error: " & CStr(firstAccessErrNumber) & vbCrLf & "Description: " & Err.Description
End Function

Private Function GeneratedFormExpectedBuildSignature(ByVal formName As String) As String
    Dim rowData As TGeneratedFormCatalogRow
    Dim sourceName As String
    Dim signatureText As String
    Dim lookupEditorPolicyFingerprint As String
    If GeneratedFormCatalogRow(formName, rowData) = False Then Err.Raise vbObjectError + 6727, "GeneratedFormExpectedBuildSignature", "No authoritative generated-form catalog row exists." & vbCrLf & "Form: " & formName
    sourceName = GeneratedFormSourceNameForSignature(formName)
    signatureText = GeneratedFormSignatureVersion(formName)
    If StrComp(rowData.formType, DEFAULT_FORM_TYPE_CONTINUOUS, vbTextCompare) = 0 Then signatureText = signatureText & "|" & CONTINUOUS_LIST_LAYOUT_VERSION
    If Len(Trim$(GeneratedSplitPopupDescription(formName))) > 0 Then signatureText = signatureText & "|" & SPLIT_POPUP_LAYOUT_VERSION
    signatureText = signatureText & "|" & rowData.formName & "|" & rowData.formDisplayName & "|" & rowData.FormDescription & "|" & rowData.formType & "|" & rowData.NavigationPaneGroupKey & "|" & CStr(rowData.SortOrder) & "|" & sourceName & "|" & GeneratedFormSourceFingerprint(sourceName)
    lookupEditorPolicyFingerprint = GeneratedFormLookupEditorPolicyFingerprint(sourceName)
    If Len(lookupEditorPolicyFingerprint) > 0 Then signatureText = signatureText & "|" & GENERATED_FORM_LOOKUP_EDITOR_BINDING_CONTRACT_VERSION & "|" & lookupEditorPolicyFingerprint
    GeneratedFormExpectedBuildSignature = AdlerTextHash(signatureText) & ":" & CStr(Len(signatureText))
End Function

Private Function GeneratedFormSignatureVersion(ByVal formName As String) As String
    GeneratedFormSignatureVersion = FORM_BUILDER_VERSION
    If LookupEditorGeneratedFormIsDefined(formName) Then GeneratedFormSignatureVersion = GeneratedFormSignatureVersion & "|" & LOOKUP_EDITOR_EXACT_CONTROL_CONTRACT_VERSION
    If StrComp(formName, ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Help"), vbTextCompare) = 0 Then GeneratedFormSignatureVersion = GeneratedFormSignatureVersion & "|" & HELP_FORM_CONTRACT_VERSION & "|" & HELP_FIXED_TOPIC_VERSION & "|" & ChatGPTHelpFixedTopicSignature()
    If StrComp(formName, ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Delete_ADC_Objects"), vbTextCompare) = 0 Then GeneratedFormSignatureVersion = GeneratedFormSignatureVersion & "|" & DELETE_OBJECT_SCOPE_DIALOG_VERSION
    If GeneratedFormIsAddInReviewForm(formName) Then GeneratedFormSignatureVersion = GeneratedFormSignatureVersion & "|" & ADD_IN_REVIEW_FORM_VERSION
    If StrComp(formName, ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Utility"), vbTextCompare) = 0 Then GeneratedFormSignatureVersion = GeneratedFormSignatureVersion & "|" & UTILITY_LAZY_LOAD_FORM_VERSION
End Function

Private Function GeneratedFormSourceFingerprint(ByVal sourceName As String) As String
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim qdf As DAO.queryDef
    Dim tdf As DAO.tableDef
    Dim fld As DAO.Field
    Dim sourceText As String
    If Len(sourceName) = 0 Then
        GeneratedFormSourceFingerprint = "UNBOUND"
        Exit Function
    End If
    If mChatGPTGeneratedFormSourceFingerprintCache Is Nothing Then
        Set mChatGPTGeneratedFormSourceFingerprintCache = CreateObject("Scripting.Dictionary")
        mChatGPTGeneratedFormSourceFingerprintCache.CompareMode = vbTextCompare
    End If
    If mChatGPTGeneratedFormSourceFingerprintCache.Exists(sourceName) Then
        mChatGPTGeneratedFormSourceFingerprintCacheHits = mChatGPTGeneratedFormSourceFingerprintCacheHits + 1
        GeneratedFormSourceFingerprint = CStr(mChatGPTGeneratedFormSourceFingerprintCache(sourceName))
        Exit Function
    End If
    mChatGPTGeneratedFormSourceFingerprintCacheMisses = mChatGPTGeneratedFormSourceFingerprintCacheMisses + 1
    Set db = GetHostDb()
    If SharedDao_QueryExists(db, sourceName) Then
        Set qdf = db.QueryDefs(sourceName)
        sourceText = "QUERY|" & sourceName & "|" & qdf.sql
    ElseIf SharedDao_TableExists(db, sourceName) Then
        Set tdf = db.TableDefs(sourceName)
        sourceText = "TABLE|" & sourceName
        For Each fld In tdf.fields
            SetupControlPump
            sourceText = sourceText & "|" & fld.Name & ":" & CStr(fld.Type) & ":" & CStr(fld.Size) & ":" & CStr(fld.Attributes)
        Next fld
    Else
        sourceText = "MISSING|" & sourceName
    End If
    GeneratedFormSourceFingerprint = AdlerTextHash(sourceText) & ":" & CStr(Len(sourceText))
    mChatGPTGeneratedFormSourceFingerprintCache(sourceName) = GeneratedFormSourceFingerprint
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "GeneratedFormSourceFingerprint", "Could not fingerprint a generated-form source." & vbCrLf & "Source: " & sourceName & vbCrLf & "Description: " & Err.Description
End Function

Private Function GeneratedFormLookupEditorPolicyFingerprint(ByVal sourceName As String) As String
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim qdf As DAO.queryDef
    Dim tdf As DAO.tableDef
    Dim sourceFields As DAO.fields
    Dim fld As DAO.Field
    Dim definitionData As Variant
    Dim lookupSourceTable As String
    Dim fingerprintText As String
    If Len(sourceName) = 0 Then Exit Function
    If mChatGPTGeneratedFormLookupEditorPolicyFingerprintCache Is Nothing Then
        Set mChatGPTGeneratedFormLookupEditorPolicyFingerprintCache = CreateObject("Scripting.Dictionary")
        mChatGPTGeneratedFormLookupEditorPolicyFingerprintCache.CompareMode = vbTextCompare
    End If
    If mChatGPTGeneratedFormLookupEditorPolicyFingerprintCache.Exists(sourceName) Then
        GeneratedFormLookupEditorPolicyFingerprint = CStr(mChatGPTGeneratedFormLookupEditorPolicyFingerprintCache(sourceName))
        Exit Function
    End If
    Set db = GetHostDb()
    If SharedDao_QueryExists(db, sourceName) Then
        Set qdf = db.QueryDefs(sourceName)
        Set sourceFields = qdf.fields
    ElseIf SharedDao_TableExists(db, sourceName) Then
        Set tdf = db.TableDefs(sourceName)
        Set sourceFields = tdf.fields
    Else
        mChatGPTGeneratedFormLookupEditorPolicyFingerprintCache(sourceName) = vbNullString
        Exit Function
    End If
    For Each fld In sourceFields
        SetupControlPump
        definitionData = GeneratedFormLookupDefinition(fld.Name, sourceName)
        If IsArray(definitionData) Then
            If StrComp(CStr(definitionData(0)), GENERATED_FORM_LOOKUP_KIND_TWO_COLUMN, vbTextCompare) = 0 Then
                lookupSourceTable = CStr(definitionData(1))
                fingerprintText = fingerprintText & "|" & CStr(Len(fld.Name)) & ":" & fld.Name & "|" & CStr(Len(lookupSourceTable)) & ":" & lookupSourceTable & LookupEditorPolicyFingerprintForPhysicalSource(lookupSourceTable)
            End If
        End If
    Next fld
    If Len(fingerprintText) > 0 Then GeneratedFormLookupEditorPolicyFingerprint = AdlerTextHash(fingerprintText) & ":" & CStr(Len(fingerprintText))
    mChatGPTGeneratedFormLookupEditorPolicyFingerprintCache(sourceName) = GeneratedFormLookupEditorPolicyFingerprint
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "GeneratedFormLookupEditorPolicyFingerprint", "Could not fingerprint generated-form lookup-editor policy." & vbCrLf & "Source: " & sourceName & vbCrLf & "Description: " & Err.Description
End Function

' Purpose: Resolves and returns generated form source name for signature from the supplied inputs and the authoritative ADC definitions/state used by this module.
' Public role: Public ADC API. The routine is intentionally callable outside its defining procedure/module and may be used directly by Access/VBA or late-bound callers when that API surface is needed.
' Parameters:
'   formName (Input; String): Physical Access form name identifying the form to create, open, configure, validate, or inspect.
' Returns: String: Returns the physical/logical name resolved for generated form source name for signature under the module's naming/definition rules; vbNullString indicates that no applicable name can be
'   resolved where supported.
Public Function GeneratedFormSourceNameForSignature(ByVal formName As String) As String
    Dim sourceCandidates As Variant
    sourceCandidates = GeneratedFormSourceCandidates(formName)
    If IsArray(sourceCandidates) = False Then Exit Function
    If GeneratedFormSourceUsesFirstExisting(formName) Then
        GeneratedFormSourceNameForSignature = FirstExistingSource(sourceCandidates)
    Else
        GeneratedFormSourceNameForSignature = CStr(sourceCandidates(LBound(sourceCandidates)))
    End If
End Function

Private Sub PreflightGeneratedFormBuildPlan(ByVal normalizedBuildScope As String)
    Dim plannedForms As Collection
    Dim seenFormNames As Object
    Dim item As Variant
    Dim formName As String
    Dim sourceName As String
    Dim sourceTypeText As String
    Dim requiredControlSources As Variant
    Dim authoritativeCount As Long
    Dim addInReviewNames As Variant
    Dim addInReviewCount As Long
    Set plannedForms = GeneratedFormPreflightPlan(normalizedBuildScope)
    authoritativeCount = GeneratedFormAuthoritativeCount(normalizedBuildScope)
    If plannedForms.count <> authoritativeCount Then Err.Raise vbObjectError + 7390, "PreflightGeneratedFormBuildPlan", "Generated-form preflight and builder inventories do not agree." & vbCrLf & "Preflight forms: " & CStr(plannedForms.count) & vbCrLf & "Builder forms: " & CStr(authoritativeCount)
    addInReviewNames = AddInReviewChildFormNames()
    If IsArray(addInReviewNames) Then addInReviewCount = UBound(addInReviewNames) - LBound(addInReviewNames) + 1
    ' INTENTIONAL KEYED LOOKUP: seenFormNames is the preflight duplicate-detection index over the planned form names; the build plan remains the ordered source and this set is not a second plan authority.
    Set seenFormNames = CreateObject("Scripting.Dictionary")
    seenFormNames.CompareMode = vbTextCompare
    For Each item In plannedForms
        SetupControlPump
        formName = CStr(item(0))
        sourceName = CStr(item(1))
        sourceTypeText = CStr(item(2))
        requiredControlSources = item(3)
        SetChatGPTBuildContext "Preflighting generated-form build plan.", formName, "Validate generated-form source and prospective control contract.", "Source: " & IIf(Len(sourceName) = 0, "(unbound)", sourceName) & "; Source type: " & sourceTypeText
        mChatGPTBuildSourceName = sourceName
        If seenFormNames.Exists(formName) Then Err.Raise vbObjectError + 6730, "PreflightGeneratedFormBuildPlan", "Generated-form preflight failed." & vbCrLf & "Form: " & formName & vbCrLf & "Validation stage: Validate unique planned form name" & vbCrLf & "Problem: The form name appears more than once in the generated-form build plan."
        seenFormNames.Add formName, True
        ValidateGeneratedFormPreparationIdentity formName
        ValidateOneGeneratedFormPreflightSource formName, sourceName, sourceTypeText, requiredControlSources
        RequireSafeWorkContinue "PreflightGeneratedFormBuildPlan", "Form [" & formName & "]"
    Next item
    If seenFormNames.count <> authoritativeCount Then Err.Raise vbObjectError + 7392, "PreflightGeneratedFormBuildPlan", "Generated-form preflight did not validate every authoritative form in the selected build scope." & vbCrLf & "Validated forms: " & CStr(seenFormNames.count) & vbCrLf & "Authoritative forms: " & CStr(authoritativeCount) & vbCrLf & "Build scope: " & normalizedBuildScope
    LogTraceMessage TRACE_LEVEL_INFO, "UI", "PreflightGeneratedFormBuildPlan", "Generated-form build plan preflight completed successfully.", "Forms checked: " & CStr(plannedForms.count) & vbCrLf & "Builder forms: " & CStr(authoritativeCount) & vbCrLf & "Add-In Review child forms: " & CStr(addInReviewCount) & vbCrLf & "Build scope: " & normalizedBuildScope, "Preflight generated-form plan", 0, vbNullString, vbNullString, DEFAULT_CHATGPT_UI_DISPLAY_PREFIX, "FormSet", "CodeDb", "FormGeneration"
End Sub

Private Sub ValidateGeneratedFormPlannedControlSources(ByVal formName As String, ByVal sourceName As String, ByVal requiredControlSources As Variant)
    On Error GoTo ErrHandler
    Dim definitions As Collection
    Dim Definition As Variant
    Dim expectedCount As Long
    Dim fieldName As String
    Dim projectedControlName As String
    Dim sourceKindPlanned As Boolean
    Dim sqlText As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim fld As DAO.Field
    Dim fieldFound As Boolean
    Dim validationStage As String
    validationStage = "Build prospective physical control definitions"
    Set definitions = ContinuousListControlDefinitions(formName, sourceName, requiredControlSources, True)
    expectedCount = UBound(requiredControlSources) - LBound(requiredControlSources) + 1
    If definitions.count <> expectedCount Then Err.Raise vbObjectError + 7394, "ValidateGeneratedFormPlannedControlSources", "The generated-control plan does not include every required ControlSource." & vbCrLf & "Form: " & formName & vbCrLf & "Source: " & sourceName & vbCrLf & "Required controls: " & CStr(expectedCount) & vbCrLf & "Planned controls: " & CStr(definitions.count)
    validationStage = "Open exact prospective RecordSource"
    sqlText = BuildExactGeneratedListSelectSql(formName, sourceName, requiredControlSources, GeneratedFormOrderFields(formName))
    Set db = GetHostDb()
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    For Each Definition In definitions
        SetupControlPump
        fieldName = CStr(Definition(0))
        projectedControlName = CStr(Definition(5))
        validationStage = "Validate prospective ControlSource: " & fieldName
        SetChatGPTBuildContext "Preflighting complete generated-form build plan.", formName, "Validate prospective exact bound-control contract.", "Source: " & sourceName & "; Required ControlSource: " & fieldName & "; Projected control: " & projectedControlName
        mChatGPTBuildSourceName = sourceName
        fieldFound = False
        Set fld = Nothing
        Err.Clear
        On Error Resume Next
        Set fld = rs.fields(fieldName)
        fieldFound = (Err.Number = 0 And Not fld Is Nothing)
        Err.Clear
        On Error GoTo ErrHandler
        If fieldFound = False Then Err.Raise vbObjectError + 7395, "ValidateGeneratedFormPlannedControlSources", "The exact generated SELECT does not expose a required planned ControlSource." & vbCrLf & "Form: " & formName & vbCrLf & "Source: " & sourceName & vbCrLf & "ControlSource: " & fieldName & vbCrLf & "Projected control: " & projectedControlName
        If Len(projectedControlName) = 0 Then Err.Raise vbObjectError + 7416, "ValidateGeneratedFormPlannedControlSources", "A prospective exact-contract control has a blank projected name." & vbCrLf & "Form: " & formName & vbCrLf & "Source: " & sourceName & vbCrLf & "ControlSource: " & fieldName
        If CLng(Definition(6)) <> CLng(acDetail) Then Err.Raise vbObjectError + 7417, "ValidateGeneratedFormPlannedControlSources", "A prospective exact-contract bound control is not assigned to the detail section." & vbCrLf & "Form: " & formName & vbCrLf & "Source: " & sourceName & vbCrLf & "ControlSource: " & fieldName
        If StrComp(formName, ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review_Inventory"), vbTextCompare) = 0 And StrComp(fieldName, FIELD_SOURCE_KIND, vbBinaryCompare) = 0 Then sourceKindPlanned = True
    Next Definition
    If StrComp(formName, ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review_Inventory"), vbTextCompare) = 0 And sourceKindPlanned = False Then Err.Raise vbObjectError + 7396, "ValidateGeneratedFormPlannedControlSources", "The Inventory form plan is missing the required SourceKind control." & vbCrLf & "Form: " & formName & vbCrLf & "Source: " & sourceName
    rs.Close
    Set rs = Nothing
    Set db = Nothing
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
    Set db = Nothing
    On Error GoTo 0
    Err.Raise savedErrNumber, "ValidateGeneratedFormPlannedControlSources", "Could not validate the exact generated-control plan." & vbCrLf & "Form: " & formName & vbCrLf & "Source: " & sourceName & vbCrLf & "Validation stage: " & validationStage & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub ValidateOneGeneratedFormPreflightSource(ByVal formName As String, ByVal sourceName As String, ByVal expectedSourceType As String, ByVal requiredControlSources As Variant)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim fld As DAO.Field
    Dim actualSourceType As String
    Dim validationStage As String
    Dim usableFieldCount As Long
    expectedSourceType = Trim$(expectedSourceType)
    validationStage = "Validate planned form name"
    If Len(formName) = 0 Then Err.Raise vbObjectError + 6731, "ValidateOneGeneratedFormPreflightSource", "The planned form name is blank."
    If Len(formName) > MAX_ACCESS_OBJECT_NAME_LENGTH Then Err.Raise vbObjectError + 6732, "ValidateOneGeneratedFormPreflightSource", "The planned form name exceeds the Access object-name limit." & vbCrLf & "Form: " & formName & vbCrLf & "Length: " & CStr(Len(formName)) & vbCrLf & "Maximum: " & CStr(MAX_ACCESS_OBJECT_NAME_LENGTH)
    If StrComp(expectedSourceType, GENERATED_FORM_SOURCE_TYPE_UNBOUND, vbTextCompare) = 0 Then
        validationStage = "Validate unbound form definition"
        If Len(sourceName) > 0 Then Err.Raise vbObjectError + 7393, "ValidateOneGeneratedFormPreflightSource", "An unbound generated form unexpectedly has a source definition."
        Exit Sub
    End If
    validationStage = "Resolve source object"
    If Len(sourceName) = 0 Then Err.Raise vbObjectError + 6733, "ValidateOneGeneratedFormPreflightSource", "The resolved source name is blank."
    Set db = GetHostDb()
    If SharedDao_TableExists(db, sourceName) Then
        actualSourceType = "Table"
    ElseIf SharedDao_QueryExists(db, sourceName) Then
        actualSourceType = "Query"
    Else
        Err.Raise vbObjectError + 6734, "ValidateOneGeneratedFormPreflightSource", "The required source object does not exist."
    End If
    validationStage = "Validate source object type"
    If StrComp(expectedSourceType, GENERATED_FORM_SOURCE_TYPE_TABLE_OR_QUERY, vbTextCompare) <> 0 Then
        If StrComp(actualSourceType, expectedSourceType, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 6735, "ValidateOneGeneratedFormPreflightSource", "The source object type does not match the generated-form plan."
    End If
    validationStage = "Open source as snapshot"
    Set rs = db.OpenRecordset(sourceName, dbOpenSnapshot)
    validationStage = "Enumerate usable source fields"
    For Each fld In rs.fields
        SetupControlPump
        If Len(Trim$(fld.Name)) > 0 Then usableFieldCount = usableFieldCount + 1
    Next fld
    If usableFieldCount = 0 Then Err.Raise vbObjectError + 6736, "ValidateOneGeneratedFormPreflightSource", "The source opened but exposed no usable fields."
    rs.Close
    Set rs = Nothing
    If IsArray(requiredControlSources) Then
        validationStage = "Validate required generated controls"
        ValidateGeneratedFormPlannedControlSources formName, sourceName, requiredControlSources
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
    Err.Raise savedErrNumber, "ValidateOneGeneratedFormPreflightSource", "Generated-form preflight failed." & vbCrLf & "Form: " & IIf(Len(formName) = 0, "(blank)", formName) & vbCrLf & "Source: " & IIf(Len(sourceName) = 0, "(unbound)", sourceName) & vbCrLf & "Expected source type: " & IIf(Len(expectedSourceType) = 0, "(blank)", expectedSourceType) & vbCrLf & "Actual source type: " & IIf(Len(actualSourceType) = 0, "(unresolved)", actualSourceType) & vbCrLf & "Validation stage: " & validationStage & vbCrLf & "DAO error number: " & CStr(savedErrNumber) & vbCrLf & "DAO error source: " & savedErrSource & vbCrLf & "DAO error description: " & savedErrDescription
End Sub

Private Sub ClearSetupGeneratedFormPreflightProofState()
    mSetupGeneratedFormPreflightReady = False
    mSetupGeneratedFormPreflightReplaceExisting = False
    mSetupGeneratedFormPreflightScope = GENERATED_FORM_PREFLIGHT_SCOPE_UNKNOWN
    mSetupGeneratedFormCompleteSetFastPath = False
    mSetupGeneratedFormCompleteSetToken = vbNullString
    mSetupGeneratedFormPreflightTraceRunID = vbNullString
    ClearSetupGeneratedFormHelpProofState
    Set mSetupGeneratedFormPreflightPhysicalNames = Nothing
End Sub

' Purpose: Reports whether UI_FORMS still owns a complete, same-trace D1 generated-form preflight plan, optionally requiring one exact normalized build scope.
' Public role: Read-only inter-module proof API. D2/D3 and Help-proof consumers use it to validate D1 state without reading or mutating UI_FORMS module variables directly.
' Parameters:
'   detailText (Output/in-out; String): Receives the first reason the retained D1 plan is unavailable or, on success, a summary of the retained scope and physical-change count.
'   expectedBuildScope (Input Optional; String): Optional generated-form scope that the retained D1 plan must match. vbNullString accepts any valid retained D1 scope.
' Returns: Boolean: True only when the D1 plan is ready, its physical-name collection exists, the trace run is still current, and any requested scope matches.
Public Function SetupGeneratedFormPreflightReadyForSetup(ByRef detailText As String, Optional ByVal expectedBuildScope As String = vbNullString) As Boolean
    detailText = vbNullString
    If mSetupGeneratedFormPreflightReady = False Then
        detailText = "D1 generated-form preflight plan is not available in the current setup session."
        Exit Function
    End If
    If mSetupGeneratedFormPreflightPhysicalNames Is Nothing Then
        detailText = "D1 generated-form preflight plan is marked ready but its physical-change collection is missing."
        Exit Function
    End If
    If Len(GetActiveOrLastTraceRunID()) = 0 Or StrComp(mSetupGeneratedFormPreflightTraceRunID, GetActiveOrLastTraceRunID(), vbBinaryCompare) <> 0 Then
        detailText = "D1 generated-form preflight plan belongs to a different or inactive setup trace run."
        Exit Function
    End If
    If Len(Trim$(expectedBuildScope)) > 0 Then
        If mSetupGeneratedFormPreflightScope <> GeneratedFormPreflightScopeStateFromText(NormalizeChatGPTBuildScope(expectedBuildScope)) Then
            detailText = "D1 generated-form preflight plan belongs to a different generated-form build scope. Retained=" & GeneratedFormPreflightScopeStateText(mSetupGeneratedFormPreflightScope) & "; expected=" & NormalizeChatGPTBuildScope(expectedBuildScope) & "."
            Exit Function
        End If
    End If
    SetupGeneratedFormPreflightReadyForSetup = True
    detailText = "D1 generated-form preflight plan is ready. Build scope=" & GeneratedFormPreflightScopeStateText(mSetupGeneratedFormPreflightScope) & "; ADC-owned physical changes=" & CStr(mSetupGeneratedFormPreflightPhysicalNames.count) & "."
End Function

Private Function GeneratedFormPostCommitToken(Optional ByVal buildScope As String = FORM_BUILD_SCOPE_ALL) As String
    'Returns a fail-closed token only when the complete expected form set and current physical form state match the last verified generated-form build.
    On Error GoTo ErrHandler
    Dim normalizedBuildScope As String
    Dim completeSetToken As String
    Dim physicalStateToken As String
    Dim detailText As String
    Dim payloadText As String
    normalizedBuildScope = NormalizeChatGPTBuildScope(buildScope)
    If CodeDbIsHostDb() = False Then Err.Raise vbObjectError + 6762, "GeneratedFormPostCommitToken", "Generated-form post-commit state can only be tokenized when CodeDb and HostDb are the same active database."
    If GeneratedFormCompactPreflightMatches(normalizedBuildScope, completeSetToken, physicalStateToken, detailText) = False Then Err.Raise vbObjectError + 6763, "GeneratedFormPostCommitToken", "The generated-form domain is not proven current." & vbCrLf & detailText
    payloadText = "ADCGeneratedFormPostCommitV3|" & normalizedBuildScope & "|" & completeSetToken & "|" & physicalStateToken
    GeneratedFormPostCommitToken = AdlerTextHash(payloadText) & ":" & CStr(Len(payloadText))
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "GeneratedFormPostCommitToken", "Could not calculate the exact generated-form post-commit token." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Sub ClearGeneratedFormSameRunD2Proof()
    mGeneratedFormSameRunD2ProofTraceRunID = vbNullString
    mGeneratedFormSameRunD2ProofScope = vbNullString
    mGeneratedFormSameRunD2ProofToken = vbNullString
    mGeneratedFormSameRunD2DiscoveredUserFormCount = 0
End Sub

Private Sub ClearGeneratedFormSameRunD3Proof()
    mGeneratedFormSameRunD3ProofTraceRunID = vbNullString
    mGeneratedFormSameRunD3ProofScope = vbNullString
    mGeneratedFormSameRunD3ProofToken = vbNullString
End Sub

Private Sub PublishGeneratedFormPartitionProof(ByVal normalizedBuildScope As String, ByVal completeSetToken As String, ByVal physicalStateToken As String, ByVal discoveredUserFormCount As Long)
    Dim traceRunID As String
    Dim payloadText As String
    traceRunID = GetActiveOrLastTraceRunID()
    If Len(traceRunID) = 0 Or Len(completeSetToken) = 0 Or Len(physicalStateToken) = 0 Then Exit Sub
    payloadText = "ADCGeneratedFormPartitionPhysicalProofV1|" & traceRunID & "|" & normalizedBuildScope & "|" & completeSetToken & "|" & physicalStateToken & "|" & CStr(discoveredUserFormCount)
    Select Case StrComp(normalizedBuildScope, FORM_BUILD_SCOPE_CHATGPT_UI_FORMS, vbTextCompare)
        Case 0
            ClearGeneratedFormSameRunD2Proof
            mGeneratedFormSameRunD2ProofTraceRunID = traceRunID
            mGeneratedFormSameRunD2ProofScope = normalizedBuildScope
            mGeneratedFormSameRunD2ProofToken = AdlerTextHash(payloadText) & ":" & CStr(Len(payloadText))
            mGeneratedFormSameRunD2DiscoveredUserFormCount = discoveredUserFormCount
        Case Else
            If StrComp(normalizedBuildScope, FORM_BUILD_SCOPE_SUPPORT_FORMS, vbTextCompare) = 0 Then
                ClearGeneratedFormSameRunD3Proof
                mGeneratedFormSameRunD3ProofTraceRunID = traceRunID
                mGeneratedFormSameRunD3ProofScope = normalizedBuildScope
                mGeneratedFormSameRunD3ProofToken = AdlerTextHash(payloadText) & ":" & CStr(Len(payloadText))
            End If
    End Select
End Sub

' Purpose: Attempts to perform get generated form same run D2 proof.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   proofTraceRunID (Output; String): Receives the resolved proof trace run ID produced from this procedure's validated inputs and current ADC state.
'   proofToken (Output; String): Receives the resolved proof token produced from this procedure's validated inputs and current ADC state.
'   detailText (Output; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
' Returns: Boolean: True when the attempt to get generated form same run D2 proof completes successfully; False for the routine's defined normal unsuccessful/not-applicable outcome.
Public Function TryGetGeneratedFormSameRunD2Proof(ByRef proofTraceRunID As String, ByRef proofToken As String, ByRef detailText As String) As Boolean
    Dim currentTraceRunID As String
    proofTraceRunID = vbNullString
    proofToken = vbNullString
    detailText = vbNullString
    currentTraceRunID = GetActiveOrLastTraceRunID()
    If Len(currentTraceRunID) = 0 Or StrComp(currentTraceRunID, mGeneratedFormSameRunD2ProofTraceRunID, vbBinaryCompare) <> 0 Then detailText = "No retained same-run D2 physical-form proof belongs to the active trace run.": Exit Function
    If StrComp(mGeneratedFormSameRunD2ProofScope, FORM_BUILD_SCOPE_CHATGPT_UI_FORMS, vbTextCompare) <> 0 Then detailText = "The retained D2 physical-form proof does not cover the ChatGPT UI Form partition.": Exit Function
    If Len(mGeneratedFormSameRunD2ProofToken) = 0 Then detailText = "The retained same-run D2 physical-form proof is incomplete.": Exit Function
    detailText = "Retained same-run D2 proof covers the ChatGPT UI Form partition only. Discovered User forms observed during D2=" & CStr(mGeneratedFormSameRunD2DiscoveredUserFormCount) & "."
    proofTraceRunID = mGeneratedFormSameRunD2ProofTraceRunID
    proofToken = mGeneratedFormSameRunD2ProofToken
    detailText = detailText & " This proof does not claim ADC Support Forms are current."
    TryGetGeneratedFormSameRunD2Proof = True
End Function

' Purpose: Attempts to perform get generated form same run D3 proof.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   proofTraceRunID (Output; String): Receives the resolved proof trace run ID produced from this procedure's validated inputs and current ADC state.
'   proofToken (Output; String): Receives the resolved proof token produced from this procedure's validated inputs and current ADC state.
'   detailText (Output; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
' Returns: Boolean: True when the attempt to get generated form same run D3 proof completes successfully; False for the routine's defined normal unsuccessful/not-applicable outcome.
Public Function TryGetGeneratedFormSameRunD3Proof(ByRef proofTraceRunID As String, ByRef proofToken As String, ByRef detailText As String) As Boolean
    Dim currentTraceRunID As String
    proofTraceRunID = vbNullString
    proofToken = vbNullString
    detailText = vbNullString
    currentTraceRunID = GetActiveOrLastTraceRunID()
    If Len(currentTraceRunID) = 0 Or StrComp(currentTraceRunID, mGeneratedFormSameRunD3ProofTraceRunID, vbBinaryCompare) <> 0 Then detailText = "No retained same-run D3 ADC Support Form proof belongs to the active trace run.": Exit Function
    If StrComp(mGeneratedFormSameRunD3ProofScope, FORM_BUILD_SCOPE_SUPPORT_FORMS, vbTextCompare) <> 0 Then detailText = "The retained D3 physical-form proof does not cover the ADC Support Form partition.": Exit Function
    If Len(mGeneratedFormSameRunD3ProofToken) = 0 Then detailText = "The retained same-run D3 ADC Support Form proof is incomplete.": Exit Function
    proofTraceRunID = mGeneratedFormSameRunD3ProofTraceRunID
    proofToken = mGeneratedFormSameRunD3ProofToken
    detailText = "Retained same-run D3 proof covers the ADC Support Form partition only; it does not claim ChatGPT UI Forms are current."
    TryGetGeneratedFormSameRunD3Proof = True
End Function

' Purpose: Finalizes generated form combined same run proof for d4 for expected D2 proof token, expected D3 proof token.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   expectedD2ProofToken (Input; String): Proof/version token identifying the exact expected D2 proof state that was previously validated in this setup run.
'   expectedD3ProofToken (Input; String): Proof/version token identifying the exact expected D3 proof state that was previously validated in this setup run.
'   detailText (Output; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
' Returns: Boolean: True on the function's affirmative/success branch for finalize generated form combined same run proof for D4; False on its explicit negative, unavailable, rejected, or handled-failure branch.
'   not-applicable path.
Public Function FinalizeGeneratedFormCombinedSameRunProofForD4(ByVal expectedD2ProofToken As String, ByVal expectedD3ProofToken As String, ByRef detailText As String) As Boolean
    ' D4 consumes, but does not reconstruct, the one D1 plan plus the independent D2/D3 partition proofs. Only their combination may publish the complete generated-form persisted proof.
    On Error GoTo ErrHandler
    Dim preflightDetail As String
    Dim traceRunID As String
    Dim allCompleteSetToken As String
    Dim allPhysicalStateToken As String
    Dim currentPlanToken As String
    Dim physicalDetail As String
    Dim allCount As Long
    Dim chatGPTUICount As Long
    Dim supportCount As Long
    Dim allFormNames As Collection
    detailText = vbNullString
    traceRunID = GetActiveOrLastTraceRunID()
    If Len(traceRunID) = 0 Then detailText = "D4 combined generated-form proof requires an active Setup trace run.": Exit Function
    If SetupGeneratedFormPreflightReadyForSetup(preflightDetail) = False Then detailText = preflightDetail: Exit Function
    If mSetupGeneratedFormPreflightScope <> GENERATED_FORM_PREFLIGHT_SCOPE_ALL Then detailText = "D4 requires D1's complete All-scope generated-form plan. Retained scope=" & GeneratedFormPreflightScopeStateText(mSetupGeneratedFormPreflightScope): Exit Function
    If StrComp(mSetupGeneratedFormPreflightTraceRunID, traceRunID, vbBinaryCompare) <> 0 Then detailText = "D1 generated-form plan belongs to a different Setup trace run.": Exit Function
    If StrComp(mGeneratedFormSameRunD2ProofTraceRunID, traceRunID, vbBinaryCompare) <> 0 Or StrComp(mGeneratedFormSameRunD3ProofTraceRunID, traceRunID, vbBinaryCompare) <> 0 Then detailText = "D2 and D3 generated-form proofs must both belong to the active D4 trace run.": Exit Function
    If StrComp(Trim$(expectedD2ProofToken), mGeneratedFormSameRunD2ProofToken, vbBinaryCompare) <> 0 Or Len(mGeneratedFormSameRunD2ProofToken) = 0 Then detailText = "D4 received a D2 ChatGPT UI Form proof token that does not match UI's retained same-run proof.": Exit Function
    If StrComp(Trim$(expectedD3ProofToken), mGeneratedFormSameRunD3ProofToken, vbBinaryCompare) <> 0 Or Len(mGeneratedFormSameRunD3ProofToken) = 0 Then detailText = "D4 received a D3 ADC Support Form proof token that does not match UI's retained same-run proof.": Exit Function
    If StrComp(mGeneratedFormSameRunD2ProofScope, FORM_BUILD_SCOPE_CHATGPT_UI_FORMS, vbTextCompare) <> 0 Or StrComp(mGeneratedFormSameRunD3ProofScope, FORM_BUILD_SCOPE_SUPPORT_FORMS, vbTextCompare) <> 0 Then detailText = "D2/D3 retained proof scopes do not match the authoritative generated-form partitions.": Exit Function
    allCount = GeneratedFormAuthoritativeCount(FORM_BUILD_SCOPE_ALL)
    chatGPTUICount = GeneratedFormAuthoritativeCount(FORM_BUILD_SCOPE_CHATGPT_UI_FORMS)
    supportCount = GeneratedFormAuthoritativeCount(FORM_BUILD_SCOPE_SUPPORT_FORMS)
    If allCount <= 0 Or chatGPTUICount + supportCount <> allCount Then detailText = "D2+D3 do not partition the complete authoritative generated-form roster exactly once. All=" & CStr(allCount) & "; ChatGPTUI=" & CStr(chatGPTUICount) & "; ADCSupport=" & CStr(supportCount): Exit Function
    currentPlanToken = GeneratedFormCompleteSetToken(FORM_BUILD_SCOPE_ALL)
    If Len(mSetupGeneratedFormCompleteSetToken) = 0 Or StrComp(currentPlanToken, mSetupGeneratedFormCompleteSetToken, vbBinaryCompare) <> 0 Then detailText = "The current complete generated-form definition token no longer matches the authoritative D1 plan.": Exit Function
    allCompleteSetToken = currentPlanToken
    allPhysicalStateToken = GeneratedFormPhysicalStateToken(FORM_BUILD_SCOPE_ALL, physicalDetail)
    If Len(allPhysicalStateToken) = 0 Then detailText = "D4 could not prove the complete generated-form physical population after D2+D3." & IIf(Len(Trim$(physicalDetail)) = 0, vbNullString, vbCrLf & physicalDetail): Exit Function
    Set allFormNames = GeneratedFormNamesForBuildScope(FORM_BUILD_SCOPE_ALL)
    ValidateOneGeneratedFormContract ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Help"), suppliedRequiredRelatedFormNames:=allFormNames
    Call EnsureDaoPropertyValue(GetHostDb(), GENERATED_FORM_COMPLETE_SET_TOKEN_PROPERTY, dbText, allCompleteSetToken, "FinalizeGeneratedFormCombinedSameRunProofForD4")
    Call EnsureDaoPropertyValue(GetHostDb(), GENERATED_FORM_PHYSICAL_STATE_TOKEN_PROPERTY, dbText, allPhysicalStateToken, "FinalizeGeneratedFormCombinedSameRunProofForD4")
    detailText = "D4 accepted one D1 All-scope plan plus same-run D2 ChatGPT UI Form and D3 ADC Support Form proofs; their disjoint union covers all " & CStr(allCount) & " authoritative generated forms. Complete persisted proof refreshed."
    ClearSetupGeneratedFormPreflightProofState
    FinalizeGeneratedFormCombinedSameRunProofForD4 = True
    Exit Function
ErrHandler:
    detailText = "Could not finalize the combined D2+D3 generated-form proof for D4. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
End Function

Private Sub AppendChatGPTGeneratedFormRollbackFailure(ByRef failureText As String, ByVal detailText As String)
    If Len(failureText) > 0 Then failureText = failureText & vbCrLf
    failureText = failureText & detailText
End Sub

Private Sub ApplyGeneratedFormDuplicateReviewCacheInvalidationEventsToOpenDesignForm(ByVal formName As String)
    'Applies cache invalidation while a generated form is already open in Design view, avoiding a second design-save pass after creation.
    On Error GoTo ErrHandler
    Dim frm As Form
    Dim recordSourceText As String
    Set frm = forms(formName)
    recordSourceText = Trim$(Nz(frm.recordSource, vbNullString))
    If GeneratedFormRecordSourceCanEditConfig(recordSourceText) Then
        frm.AfterUpdate = "=InvalidateDuplicateReviewCache()"
        frm.AfterInsert = "=InvalidateDuplicateReviewCache()"
        frm.OnDelete = "=InvalidateDuplicateReviewCache()"
        frm.AfterDelConfirm = "=InvalidateDuplicateReviewCache()"
    End If
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ApplyGeneratedFormDuplicateReviewCacheInvalidationEventsToOpenDesignForm", "Could not apply duplicate-review cache invalidation events before saving a generated form." & vbCrLf & "Form: " & formName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Function BuildGeneratedFormRecoveryState(ByVal physicalFormNames As Collection, ByVal backupEntries As Collection, ByVal originalFormNames As Collection, ByVal formCatalogRollbackTable As String) As Collection
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: result packages the physical-form scope, backup entries, original-form set, and catalog rollback table as one recovery payload; it exists only for the
    ' generated-form rollback lifecycle.
    Dim result As New Collection
    result.Add physicalFormNames, "PhysicalForms"
    result.Add backupEntries, "BackupEntries"
    result.Add originalFormNames, "OriginalForms"
    result.Add formCatalogRollbackTable, "CatalogRollbackTable"
    Set BuildGeneratedFormRecoveryState = result
End Function

Private Sub CleanupChatGPTGeneratedFormRollbackFiles(ByVal backupEntries As Collection, ByVal formCatalogRollbackTable As String, Optional ByVal requireVerifiedCleanup As Boolean = False)
    Dim backupEntry As Variant
    Dim db As DAO.Database
    Dim backupPath As String
    Dim failedPaths As String
    If requireVerifiedCleanup = False Then
        On Error Resume Next
        If Not (backupEntries Is Nothing) Then
            For Each backupEntry In backupEntries
                SetupControlPump
                DeleteRollbackBackupFile CStr(backupEntry(1)), failedPaths
            Next backupEntry
        End If
        If Len(Trim$(formCatalogRollbackTable)) > 0 Then
            Set db = GetHostDb()
            db.TableDefs.Refresh
            If SharedDao_TableExists(db, formCatalogRollbackTable) Then db.TableDefs.Delete formCatalogRollbackTable
        End If
        Exit Sub
    End If
    On Error GoTo ErrHandler
    If Not (backupEntries Is Nothing) Then
        For Each backupEntry In backupEntries
            SetupControlPump
            backupPath = CStr(backupEntry(1))
            DeleteRollbackBackupFile backupPath, failedPaths
        Next backupEntry
        If Len(failedPaths) > 0 Then Err.Raise ERR_FORM_RESTORE_FAILED, "CleanupChatGPTGeneratedFormRollbackFiles", "One or more generated-form rollback backup files could not be deleted and verified." & vbCrLf & failedPaths
    End If
    If Len(Trim$(formCatalogRollbackTable)) > 0 Then
        Set db = GetHostDb()
        db.TableDefs.Refresh
        If UITableExistsStrict(db, formCatalogRollbackTable) Then db.TableDefs.Delete formCatalogRollbackTable
        db.TableDefs.Refresh
        If UITableExistsStrict(db, formCatalogRollbackTable) Then Err.Raise ERR_GENERATED_CATALOG_ROLLBACK_FAILED, "CleanupChatGPTGeneratedFormRollbackFiles", "The generated-form catalog rollback table still exists after deletion." & vbCrLf & "Table: " & formCatalogRollbackTable
    End If
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "CleanupChatGPTGeneratedFormRollbackFiles", "Verified generated-form rollback-resource cleanup failed." & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Sub

Private Sub BuildOneCanonicalGeneratedForm(ByVal formName As String, ByVal replaceExisting As Boolean, ByRef builtForms As Collection, ByRef skippedForms As Collection)
    Dim sourceName As String
    sourceName = GeneratedFormSourceNameForSignature(formName)
    SetChatGPTBuildContext "Creating canonical generated form.", formName, "Create the generated form through its UI-owned implementation.", "Display name: " & GeneratedFormDisplayName(formName) & "; Source: " & IIf(Len(sourceName) = 0, "(unbound)", sourceName)
    UpdateChatGPTProgress "Creating generated form: " & GeneratedFormDisplayName(formName) & "."
    'The D1/D2 physical-change plan is authoritative for every canonical builder, including Setup Control special cases; forms outside that plan are verified existing state and must not be physically touched.
    If Not mChatGPTGeneratedFormsRequiringBuild Is Nothing Then
        If mChatGPTGeneratedFormsRequiringBuild.Exists(formName) = False Then
            AddVerifiedExistingFormNote formName
            Exit Sub
        End If
    End If
    If AddInReviewChildFormIsDefinedDefault(formName) Then
        CreateChatGPTAdministrativeListForm formName, AddInReviewChildFormCaption(formName), AddInReviewChildSourceName(formName), AddInReviewExpectedControlSources(formName), AddInReviewChildOrderFields(formName), AddInReviewChildActionProfile(formName), GeneratedFormListAllowEdits(formName), replaceExisting, builtForms, skippedForms, AddInReviewChildOnCurrentExpression(formName), True, True
        Exit Sub
    End If
    Select Case UCase$(formName)
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Manager")): CreateChatGPTManagerForm replaceExisting, builtForms, skippedForms
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Help")): CreateChatGPTHelpForm replaceExisting, builtForms, skippedForms
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Report_Error_Dialog")): CreateReportErrorDialogForm replaceExisting, builtForms, skippedForms
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Delete_ADC_Objects")): CreateDeleteObjectsDialogForm replaceExisting, builtForms, skippedForms
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_RUNNING)): EnsureSetupControlForm formName, SETUP_CONTROL_RUNNING_FORM_TAG, True: builtForms.Add formName
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_PAUSED)): EnsureSetupControlForm formName, SETUP_CONTROL_PAUSED_FORM_TAG, False: builtForms.Add formName
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SCAN_PROGRESS)): EnsureSetupControlForm formName, SCAN_PROGRESS_FORM_TAG, True, True: builtForms.Add formName
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Configuration_Detail")): CreateChatGPTConfigurationDetailForm replaceExisting, builtForms, skippedForms
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Configuration_List"))
            CreateConfigurationListLikeForm _
                ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Configuration_List"), _
                GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Configuration_List")), _
                GeneratedFormSourceNameForSignature(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Configuration_List")), _
                vbNullString, _
                replaceExisting, _
                builtForms, _
                skippedForms, _
                True
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Candidate_Review")): CreateChatGPTCandidateReviewForm replaceExisting, builtForms, skippedForms
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Settings")): CreateChatGPTSettingsForm replaceExisting, builtForms, skippedForms
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Health_Review")): CreateChatGPTHealthReviewForm replaceExisting, builtForms, skippedForms
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "LauncherControlsReviewForm")): CreateChatGPTCustomMacroLauncherReviewForm replaceExisting, builtForms, skippedForms
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Catalog_Review")): CreateChatGPTCatalogReviewForm replaceExisting, builtForms, skippedForms
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "System_Macros")): CreateChatGPTSystemMacrosForm replaceExisting, builtForms, skippedForms
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "NavigationPane_Groups")): CreateChatGPTNavigationPaneGroupsForm replaceExisting, builtForms, skippedForms
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Form_Types")): CreateChatGPTFormTypesForm replaceExisting, builtForms, skippedForms
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Forms")): CreateChatGPTFormsCatalogForm replaceExisting, builtForms, skippedForms
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Active_References"))
            CreateChatGPTAdministrativeListForm _
                ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Active_References"), _
                GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Active_References")), _
                GeneratedFormSourceNameForSignature(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Active_References")), _
                GeneratedFormListFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Active_References")), _
                GeneratedFormOrderFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Active_References")), _
                GeneratedFormListActionProfile(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Active_References")), _
                GeneratedFormListAllowEdits(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Active_References")), _
                replaceExisting, _
                builtForms, _
                skippedForms
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Available_References"))
            CreateChatGPTAdministrativeListForm _
                ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Available_References"), _
                GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Available_References")), _
                GeneratedFormSourceNameForSignature(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Available_References")), _
                GeneratedFormListFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Available_References")), _
                GeneratedFormOrderFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Available_References")), _
                GeneratedFormListActionProfile(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Available_References")), _
                GeneratedFormListAllowEdits(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Available_References")), _
                replaceExisting, _
                builtForms, _
                skippedForms
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Reference_Manager"))
            CreateChatGPTAdministrativeListForm _
                ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Reference_Manager"), _
                GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Reference_Manager")), _
                GeneratedFormSourceNameForSignature(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Reference_Manager")), _
                GeneratedFormListFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Reference_Manager")), _
                GeneratedFormOrderFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Reference_Manager")), _
                GeneratedFormListActionProfile(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Reference_Manager")), _
                GeneratedFormListAllowEdits(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Reference_Manager")), _
                replaceExisting, _
                builtForms, _
                skippedForms
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Updates_Current"))
            CreateChatGPTAdministrativeListForm _
                ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Updates_Current"), _
                GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Updates_Current")), _
                GeneratedFormSourceNameForSignature(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Updates_Current")), _
                GeneratedFormListFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Updates_Current")), _
                GeneratedFormOrderFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Updates_Current")), _
                GeneratedFormListActionProfile(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Updates_Current")), _
                GeneratedFormListAllowEdits(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Updates_Current")), _
                replaceExisting, _
                builtForms, _
                skippedForms
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Updates_History"))
            CreateChatGPTAdministrativeListForm _
                ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Updates_History"), _
                GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Updates_History")), _
                GeneratedFormSourceNameForSignature(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Updates_History")), _
                GeneratedFormListFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Updates_History")), _
                GeneratedFormOrderFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Updates_History")), _
                GeneratedFormListActionProfile(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Updates_History")), _
                GeneratedFormListAllowEdits(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Updates_History")), _
                replaceExisting, _
                builtForms, _
                skippedForms
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Verification"))
            CreateChatGPTAdministrativeListForm _
                ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Verification"), _
                GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Verification")), _
                GeneratedFormSourceNameForSignature(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Verification")), _
                GeneratedFormListFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Verification")), _
                GeneratedFormOrderFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Verification")), _
                GeneratedFormListActionProfile(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Verification")), _
                GeneratedFormListAllowEdits(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Verification")), _
                replaceExisting, _
                builtForms, _
                skippedForms
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Configuration_Split"))
            CreateSplitPopupForm _
                ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Configuration_Split"), _
                GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Configuration_Split")), _
                GeneratedFormSourceNameForSignature(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Configuration_List")), _
                vbNullString, _
                ConfigurationSplitPopupFields(), _
                GeneratedFormOrderFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Configuration_Split")), _
                replaceExisting, _
                builtForms, _
                skippedForms, _
                True, _
                GeneratedSplitPopupDescription(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Configuration_Split"))
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Candidate_Split")): CreateChatGPTCandidateSplitPopupForm replaceExisting, builtForms, skippedForms
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Settings_Split"))
            CreateSplitPopupForm _
                ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Settings_Split"), _
                GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Settings_Split")), _
                GeneratedFormSourceNameForSignature(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Settings_Split")), _
                vbNullString, _
                SettingsSplitPopupFields(), _
                GeneratedFormOrderFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Settings_Split")), _
                replaceExisting, _
                builtForms, _
                skippedForms, _
                False, _
                GeneratedSplitPopupDescription(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Settings_Split"))
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Health_Split"))
            CreateSplitPopupForm _
                ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Health_Split"), _
                GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Health_Split")), _
                GeneratedFormSourceNameForSignature(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Health_Split")), _
                vbNullString, _
                HealthSplitPopupFields(), _
                GeneratedFormOrderFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Health_Split")), _
                replaceExisting, _
                builtForms, _
                skippedForms, _
                False, _
                GeneratedSplitPopupDescription(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Health_Split"))
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "LauncherControls_Split"))
            CreateSplitPopupForm _
                ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "LauncherControls_Split"), _
                GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "LauncherControls_Split")), _
                GeneratedFormSourceNameForSignature(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "LauncherControls_Split")), _
                vbNullString, _
                CustomMacroLauncherSplitPopupFields(), _
                GeneratedFormOrderFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "LauncherControls_Split")), _
                replaceExisting, _
                builtForms, _
                skippedForms, _
                False, _
                GeneratedSplitPopupDescription(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "LauncherControls_Split"))
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Catalog_Split"))
            CreateSplitPopupForm _
                ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Catalog_Split"), _
                GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Catalog_Split")), _
                GeneratedFormSourceNameForSignature(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Catalog_Split")), _
                vbNullString, _
                CatalogSplitPopupFields(), _
                GeneratedFormOrderFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Catalog_Split")), _
                replaceExisting, _
                builtForms, _
                skippedForms, _
                False, _
                GeneratedSplitPopupDescription(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Catalog_Split"))
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Forms_Split"))
            CreateSplitPopupForm _
                ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Forms_Split"), _
                GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Forms_Split")), _
                ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_FormDefinitions"), _
                vbNullString, _
                FormsCatalogSplitPopupFields(), _
                GeneratedFormOrderFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Forms_Split")), _
                replaceExisting, _
                builtForms, _
                skippedForms, _
                False, _
                GeneratedSplitPopupDescription(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Forms_Split"))
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review")): CreateChatGPTAddInReviewForm replaceExisting, builtForms, skippedForms
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Utility")): CreateChatGPTTabbedUtilityForm replaceExisting, builtForms, skippedForms
        Case Else
            If LookupEditorGeneratedFormIsDefined(formName) Then
                CreateConfigurationListLikeForm formName, GeneratedFormDisplayName(formName), GeneratedFormSourceNameForSignature(formName), vbNullString, replaceExisting, builtForms, skippedForms, True
            Else
                Err.Raise vbObjectError + 7397, "BuildOneCanonicalGeneratedForm", "A canonical generated form has no UI build implementation." & vbCrLf & "Form: " & formName
            End If
    End Select
End Sub

Private Function GeneratedFormPhysicalDefinitionHash(ByVal formName As String) As String
    On Error GoTo ErrHandler
    Dim exportPath As String
    Dim definitionText As String
    exportPath = Environ$("TEMP") & "\ADC_FormSignature_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & Format$(CLng(Timer * 1000), "000000000") & ".txt"
    Application.SaveAsText acForm, formName, exportPath
    definitionText = NormalizedGeneratedFormDefinitionText(ReadGeneratedFormDefinitionText(exportPath))
    GeneratedFormPhysicalDefinitionHash = AdlerTextHash(definitionText) & ":" & CStr(Len(definitionText))
    SharedFile_DeleteIfExists exportPath
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    On Error Resume Next
    If Len(exportPath) > 0 Then SharedFile_DeleteIfExists exportPath
    On Error GoTo 0
    Err.Raise savedErrNumber, "GeneratedFormPhysicalDefinitionHash", "Could not fingerprint a generated form definition." & vbCrLf & "Form: " & formName & vbCrLf & "Description: " & savedErrDescription
End Function

Private Function GeneratedFormRecordSourceCanEditConfig(ByVal recordSourceText As String) As Boolean
    Dim sourceNames As Variant
    Dim sourceItem As Variant
    Dim sourceName As String
    Dim bracketedSourceName As String
    If Len(recordSourceText) = 0 Then Exit Function
    sourceNames = Array( _
        ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config"), _
        ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Config_Management"), _
        ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Config_CandidateReview"), _
        ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Config_DuplicateRows"))
    For Each sourceItem In sourceNames
        sourceName = CStr(sourceItem)
        bracketedSourceName = SharedDao_BracketName(sourceName)
        If StrComp(recordSourceText, sourceName, vbTextCompare) = 0 Then
            GeneratedFormRecordSourceCanEditConfig = True
            Exit Function
        End If
        If StrComp(recordSourceText, bracketedSourceName, vbTextCompare) = 0 Then
            GeneratedFormRecordSourceCanEditConfig = True
            Exit Function
        End If
        If InStr(1, recordSourceText, " FROM " & bracketedSourceName, vbTextCompare) > 0 Or InStr(1, recordSourceText, " JOIN " & bracketedSourceName, vbTextCompare) > 0 Then
            GeneratedFormRecordSourceCanEditConfig = True
            Exit Function
        End If
    Next sourceItem
End Function

Private Function NormalizedGeneratedFormDefinitionText(ByVal definitionText As String) As String
    Dim lines As Variant
    Dim lineItem As Variant
    Dim normalizedText As String
    definitionText = Replace$(definitionText, vbCrLf, vbLf)
    definitionText = Replace$(definitionText, vbCr, vbLf)
    lines = Split(definitionText, vbLf)
    For Each lineItem In lines
        SetupControlPump
        If InStr(1, CStr(lineItem), GENERATED_FORM_BUILD_SIGNATURE_PROPERTY, vbTextCompare) = 0 And InStr(1, CStr(lineItem), GENERATED_FORM_DEFINITION_HASH_PROPERTY, vbTextCompare) = 0 Then
            normalizedText = normalizedText & RTrim$(CStr(lineItem)) & vbLf
        End If
    Next lineItem
    NormalizedGeneratedFormDefinitionText = normalizedText
End Function

' Purpose: Determines whether a managed generated form still requires design-time rebuild work, invokes the shared replacement/preparation contract when needed, and records a skip reason when preparation cannot
'   proceed.
' Public role: Inter-module public generated-form build API. Generated-form builders call this routine before design work; Public visibility is required for the shared form-build contract.
' Parameters:
'   formName (Input; String): Physical Access form name identifying the form to create, open, configure, validate, or inspect.
'   replaceExisting (Input; Boolean): When True, enables or requests existing for this call; when False, the routine follows its normal non-forced/default path.
'   skippedForms (Input/output; Collection): Supplies the caller-owned skipped-form collection and can receive a new skip-reason entry when preparation does not proceed.
'   ignoreScopedBuildPlan (Input Optional; Boolean): When True, bypasses only the active scoped-build membership filter so the pre-D Report Error bootstrap can materialize its one emergency form independently.
' Returns: Boolean: True only when the form has been successfully prepared and the caller should continue building it; False when no rebuild is required or shared preparation does not proceed.
Public Function PrepareChatGPTGeneratedFormForBuild(ByVal formName As String, ByVal replaceExisting As Boolean, ByRef skippedForms As Collection, Optional ByVal ignoreScopedBuildPlan As Boolean = False) As Boolean
    If Not mChatGPTGeneratedFormsRequiringBuild Is Nothing And ignoreScopedBuildPlan = False Then
        If mChatGPTGeneratedFormsRequiringBuild.Exists(formName) = False Then
            AddVerifiedExistingFormNote formName
            Exit Function
        End If
    ElseIf replaceExisting And GeneratedFormBuildSignatureIsCurrent(formName) Then
        AddVerifiedExistingFormNote formName
        Exit Function
    End If
    If Len(mSharedPreparedGeneratedFormName) > 0 Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "PrepareChatGPTGeneratedFormForBuild", "A prior shared generated-form replacement is still active and must be recovered before another form is prepared." & vbCrLf & "Prepared form: " & mSharedPreparedGeneratedFormName & vbCrLf & "Requested form: " & formName
    ValidateGeneratedFormPreparationIdentity formName
    If SharedFormDesign_PrepareGeneratedForm(formName, replaceExisting, vbNullString) Then
        mSharedPreparedGeneratedFormName = formName
        PrepareChatGPTGeneratedFormForBuild = True
        Exit Function
    End If
    If replaceExisting = False And SharedDao_AccessObjectExists(GetHostDb(), acForm, formName) Then
        AddVerifiedExistingFormNote formName
    Else
        AddSkippedFormNote skippedForms, formName, "Generated form preparation returned False. ReplaceExisting=" & CStr(replaceExisting) & "."
    End If
End Function

Private Sub PrepareChatGPTGeneratedFormRollbackState(ByVal saveExistingDefinitions As Boolean, ByVal physicalFormNames As Collection, ByRef backupEntries As Collection, ByRef originalFormNames As Collection, ByRef formCatalogRollbackTable As String, Optional ByVal includeFormCatalogRollback As Boolean = True)
    On Error GoTo ErrHandler
    Dim formItem As Variant
    Dim backupEntry As Variant
    Dim backupPath As String
    Dim backupIndex As Long
    Dim db As DAO.Database
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim cleanupErrNumber As Long
    Dim cleanupErrSource As String
    Dim cleanupErrDescription As String
    Dim rollbackCatalogRowCount As Long
    Dim physicalFormNameSqlList As String
    Dim formCatalogTableName As String
    Dim sourceTableDef As DAO.tableDef
    Dim sourceField As DAO.Field
    Dim rollbackFieldSqlList As String
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: backupEntries, originalFormNames holds planned/recovery work or independently observed validation evidence for PrepareChatGPTGeneratedFormRollbackState; keep it
    ' distinct from the source authority.
    Set backupEntries = New Collection
    Set originalFormNames = New Collection
    formCatalogRollbackTable = vbNullString
    If physicalFormNames Is Nothing Then Err.Raise vbObjectError + 6712, "PrepareChatGPTGeneratedFormRollbackState", "No physical generated-form change scope was supplied."
    physicalFormNameSqlList = ChatGPTFormNameCollectionSqlList(physicalFormNames, "PrepareChatGPTGeneratedFormRollbackState")
    EnsureCoreRecoveryStorageReady
    For Each formItem In physicalFormNames
        SetupControlPump
        If SharedDao_AccessObjectExists(GetHostDb(), acForm, CStr(formItem)) Then
            originalFormNames.Add CStr(formItem)
            If saveExistingDefinitions Then
                backupIndex = backupIndex + 1
                backupPath = Environ$("TEMP") & "\ADC_GeneratedFormBackup_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & Format$(CLng(Timer * 1000), "000000000") & "_" & CStr(backupIndex) & ".txt"
                If CurrentProject.AllForms(CStr(formItem)).IsLoaded Then Err.Raise vbObjectError + 6726, "PrepareChatGPTGeneratedFormRollbackState", "An affected generated form is still open after the required design-change preparation." & vbCrLf & "Form: " & CStr(formItem)
                Application.SaveAsText acForm, CStr(formItem), backupPath
                backupEntry = Array(CStr(formItem), backupPath)
                backupEntries.Add backupEntry
            End If
        End If
    Next formItem
    formCatalogTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_FormDefinitions")
    If includeFormCatalogRollback And SharedDao_TableExists(GetHostDb(), formCatalogTableName) Then
        formCatalogRollbackTable = "z__ADC_FormCatalogRollback_" & Format$(Now, "hhnnss") & "_" & Format$(CLng(Timer * 1000), "000000000")
        Set db = GetHostDb()
        Set sourceTableDef = db.TableDefs(formCatalogTableName)
        sourceTableDef.fields.Refresh
        For Each sourceField In sourceTableDef.fields
            SetupControlPump
            If StrComp(sourceField.Name, FIELD_FORM_DISPLAY_NAME_CALC, vbTextCompare) <> 0 Then
                If Len(rollbackFieldSqlList) > 0 Then rollbackFieldSqlList = rollbackFieldSqlList & ","
                rollbackFieldSqlList = rollbackFieldSqlList & SharedDao_BracketName(sourceField.Name)
            End If
        Next sourceField
        If Len(rollbackFieldSqlList) = 0 Then Err.Raise ERR_GENERATED_CATALOG_ROLLBACK_FAILED, "PrepareChatGPTGeneratedFormRollbackState", "The generated form catalog has no writable fields available for rollback snapshot creation."
        db.Execute "SELECT " & rollbackFieldSqlList & " INTO [" & formCatalogRollbackTable & "] FROM [" & formCatalogTableName & "] WHERE " & SQL_FIELD_FORM_OBJECT_NAME & " IN (" & physicalFormNameSqlList & ");", dbFailOnError
        db.TableDefs.Refresh
        If SharedDao_TableExists(db, formCatalogRollbackTable) = False Then Err.Raise ERR_GENERATED_CATALOG_ROLLBACK_FAILED, "PrepareChatGPTGeneratedFormRollbackState", "The generated-form catalog rollback table was not visible after creation and TableDefs refresh." & vbCrLf & "Database: " & db.Name & vbCrLf & "Rollback table: " & formCatalogRollbackTable & vbCrLf & "TableDefs count: " & CStr(db.TableDefs.count)
        ApplyGeneratedFormRollbackTableDescriptions db, formCatalogTableName, formCatalogRollbackTable
        rollbackCatalogRowCount = ChatGPTRecordsetCount(db, "SELECT Count(*) AS RowCountValue FROM [" & formCatalogRollbackTable & "];")
    End If
    Dim preparedRecoveryState As Collection
    Set preparedRecoveryState = BuildGeneratedFormRecoveryState(physicalFormNames, backupEntries, originalFormNames, formCatalogRollbackTable)
    PrepareUIDurableRecovery "Generated-form restoration", "Generated-form backups and catalog rollback state were prepared before replacement. A reset or interruption requires verified restoration before later work." & vbCrLf & "Catalog rollback table: " & IIf(Len(formCatalogRollbackTable) = 0, "(not required)", formCatalogRollbackTable) & vbCrLf & "Catalog rollback rows: " & CStr(rollbackCatalogRowCount), UI_RECOVERY_TYPE_GENERATED_FORMS, SerializeGeneratedFormRecoveryState(preparedRecoveryState)
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    Err.Clear
    CleanupChatGPTGeneratedFormRollbackFiles backupEntries, formCatalogRollbackTable, True
    cleanupErrNumber = Err.Number
    cleanupErrSource = Err.Source
    cleanupErrDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    Err.Raise savedErrNumber, "PrepareChatGPTGeneratedFormRollbackState", "Could not prepare generated-form rollback state before replacement." & vbCrLf & "Form: " & CStr(formItem) & vbCrLf & "Backup path: " & backupPath & vbCrLf & "Catalog rollback table: " & formCatalogRollbackTable & vbCrLf & "Rollback-resource cleanup: " & IIf(cleanupErrNumber = 0, "deleted and verified", "FAILED") & IIf(cleanupErrNumber = 0, vbNullString, vbCrLf & "Cleanup Err.Number: " & CStr(cleanupErrNumber) & vbCrLf & "Cleanup Err.Source: " & cleanupErrSource & vbCrLf & "Cleanup Err.Description: " & cleanupErrDescription) & vbCrLf & "Original Err.Source: " & savedErrSource & vbCrLf & "Original Err.Description: " & savedErrDescription
End Sub

Private Sub ApplyGeneratedFormRollbackTableDescriptions(ByVal db As DAO.Database, ByVal sourceTableName As String, ByVal rollbackTableName As String)
    'The rollback snapshot is deliberately outside the managed catalog and can be retained by durable recovery, so UI owns complete metadata at creation.
    On Error GoTo ErrHandler
    Dim rollbackTdf As DAO.tableDef
    Dim rollbackField As DAO.Field
    Dim descriptionText As String
    Set rollbackTdf = db.TableDefs(rollbackTableName)
    Call EnsureDaoPropertyValue(rollbackTdf, DAO_DESCRIPTION_PROPERTY_NAME, dbText, Left$("Durable generated-form recovery snapshot of affected rows from [" & sourceTableName & "]. UI uses it only to restore the pre-replacement form catalog after failure; it is not the active form-definition authority and is deleted after verified cleanup.", 255), "ApplyGeneratedFormRollbackTableDescriptions")
    For Each rollbackField In rollbackTdf.fields
        descriptionText = "Recovery snapshot of [" & sourceTableName & "].[" & rollbackField.Name & "] for generated-form catalog rollback. The value represents pre-replacement state and must be consumed only by the verified UI restoration path."
        Call EnsureDaoPropertyValue(rollbackField, DAO_DESCRIPTION_PROPERTY_NAME, dbText, Left$(descriptionText, 255), "ApplyGeneratedFormRollbackTableDescriptions")
    Next rollbackField
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ApplyGeneratedFormRollbackTableDescriptions", "Could not apply and verify developer descriptions on generated-form rollback table [" & rollbackTableName & "]." & vbCrLf & savedErrDescription
End Sub

Private Function ReadGeneratedFormDefinitionText(ByVal filePath As String) As String
    On Error GoTo ErrHandler
    Dim fileNumber As Integer
    Dim fileLength As Long
    fileNumber = FreeFile
    Open filePath For Binary Access Read As #fileNumber
    fileLength = LOF(fileNumber)
    If fileLength > 0 Then
        ReadGeneratedFormDefinitionText = Space$(fileLength)
        Get #fileNumber, , ReadGeneratedFormDefinitionText
    End If
    Close #fileNumber
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    On Error Resume Next
    If fileNumber > 0 Then Close #fileNumber
    On Error GoTo 0
    Err.Raise savedErrNumber, "ReadGeneratedFormDefinitionText", "Could not read an exported generated-form definition." & vbCrLf & "File: " & filePath & vbCrLf & "Description: " & savedErrDescription
End Function

Private Sub RefreshGeneratedFormBuildSignatures(ByVal formNames As Collection)
    Dim formItem As Variant
    If formNames Is Nothing Then Exit Sub
    For Each formItem In formNames
        SetupControlPump
        StampGeneratedFormBuildSignature CStr(formItem)
    Next formItem
End Sub

Private Sub RequireChatGPTGeneratedFormRestoredAndLoadable(ByVal formName As String)
    On Error GoTo ErrHandler
    If SharedDao_AccessObjectExists(GetHostDb(), acForm, formName) = False Then Err.Raise vbObjectError + 6727, "RequireChatGPTGeneratedFormRestoredAndLoadable", "The original generated form was not restored." & vbCrLf & "Form: " & formName
    DoCmd.OpenForm formName, acDesign, , , , acHidden
    DoCmd.Close acForm, formName, acSaveNo
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If CurrentProject.AllForms(formName).IsLoaded Then DoCmd.Close acForm, formName, acSaveNo
    On Error GoTo 0
    Err.Raise savedErrNumber, "RequireChatGPTGeneratedFormRestoredAndLoadable", "The restored generated form could not be opened and closed in Design view." & vbCrLf & "Form: " & formName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Sub

Private Function RestoreChatGPTGeneratedFormsAfterFailure(ByVal physicalFormNames As Collection, ByVal backupEntries As Collection, ByVal originalFormNames As Collection, ByVal formCatalogRollbackTable As String, Optional ByVal preserveRollbackResources As Boolean = False) As String
    Dim recoveryClearDetail As String
    Dim formItem As Variant
    Dim backupEntry As Variant
    Dim restoredCount As Long
    Dim removedNewCount As Long
    Dim failureText As String
    Dim catalogRestoreText As String
    Dim summaryText As String
    Dim itemFailureText As String
    Dim recoveryState As Collection
    Dim unresolvedWorkspace As DAO.workspace
    EnsureCoreRecoveryStorageReady
    Set recoveryState = BuildGeneratedFormRecoveryState(physicalFormNames, backupEntries, originalFormNames, formCatalogRollbackTable)
    If Not backupEntries Is Nothing Then
        For Each backupEntry In backupEntries
            SetupControlPump
            itemFailureText = vbNullString
            If RestoreFormFromBackupByVerifiedSwap(CStr(backupEntry(0)), CStr(backupEntry(1)), itemFailureText) Then
                restoredCount = restoredCount + 1
            Else
                AppendChatGPTGeneratedFormRollbackFailure failureText, itemFailureText
            End If
        Next backupEntry
    End If
    If Len(failureText) = 0 Then
        If RestoreChatGPTGeneratedFormCatalogRowsTransactional(physicalFormNames, formCatalogRollbackTable, catalogRestoreText, unresolvedWorkspace) = False Then AppendChatGPTGeneratedFormRollbackFailure failureText, catalogRestoreText
    Else
        catalogRestoreText = "Generated form catalog restoration was deferred because one or more original forms were not restored and verified."
    End If
    If unresolvedWorkspace Is Nothing And Len(failureText) = 0 Then
        If Not physicalFormNames Is Nothing Then
            For Each formItem In physicalFormNames
                SetupControlPump
                If ChatGPTCollectionContainsText(originalFormNames, CStr(formItem)) = False Then
                    On Error Resume Next
                    Err.Clear
                    If SharedDao_AccessObjectExists(GetHostDb(), acForm, CStr(formItem)) Then
                        SharedFormDesign_CloseFormIfOpen CStr(formItem), acSaveNo
                        If Err.Number = 0 Then DoCmd.DeleteObject acForm, CStr(formItem)
                    End If
                    If Err.Number = 0 Then
                        removedNewCount = removedNewCount + 1
                    Else
                        AppendChatGPTGeneratedFormRollbackFailure failureText, "Could not remove newly created form [" & CStr(formItem) & "]. Error " & CStr(Err.Number) & ": " & Err.Description
                    End If
                    On Error GoTo 0
                End If
            Next formItem
        End If
    End If
    If Not originalFormNames Is Nothing Then
        For Each formItem In originalFormNames
            SetupControlPump
            On Error Resume Next
            Err.Clear
            RequireChatGPTGeneratedFormRestoredAndLoadable CStr(formItem)
            If Err.Number <> 0 Then AppendChatGPTGeneratedFormRollbackFailure failureText, "Restored form verification failed for [" & CStr(formItem) & "]. Error " & CStr(Err.Number) & ": " & Err.Description
            On Error GoTo 0
        Next formItem
    End If
    summaryText = "Original forms restored and verified: " & CStr(restoredCount) & vbCrLf & "Newly created forms removed: " & CStr(removedNewCount) & vbCrLf & catalogRestoreText
    RestoreChatGPTGeneratedFormsAfterFailure = summaryText
    If Len(failureText) > 0 Then
        RegisterUIRecoveryFailure "Generated-form restoration", summaryText & vbCrLf & "Rollback failures:" & vbCrLf & failureText & vbCrLf & "Form-definition backup files and the catalog rollback table were preserved for verified retry.", unresolvedWorkspace, recoveryState, UI_RECOVERY_TYPE_GENERATED_FORMS
        Err.Raise vbObjectError + 6725, "RestoreChatGPTGeneratedFormsAfterFailure", summaryText & vbCrLf & "Rollback failures:" & vbCrLf & failureText
    End If
    If ClearUIRecoveryFailure("Generated-form restoration", recoveryClearDetail) = False Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "RestoreChatGPTGeneratedFormsAfterFailure", recoveryClearDetail
    If preserveRollbackResources = False Then
        On Error Resume Next
        Err.Clear
        CleanupChatGPTGeneratedFormRollbackFiles backupEntries, formCatalogRollbackTable, True
        If Err.Number <> 0 Then
            summaryText = summaryText & vbCrLf & "Rollback-resource cleanup warning. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
            LogTraceMessage TRACE_LEVEL_WARNING, "UIRecovery", "RestoreChatGPTGeneratedFormsAfterFailure", "Generated-form restoration completed, but rollback-resource cleanup reported a warning.", summaryText, "Cleanup generated-form rollback resources", 0, vbNullString, vbNullString, DEFAULT_CHATGPT_UI_DISPLAY_PREFIX, "FormSet", "CodeDb", "Recovery"
        End If
        Err.Clear
        On Error GoTo 0
        RestoreChatGPTGeneratedFormsAfterFailure = summaryText
    End If
End Function

' Purpose: Resolves serialize generated form recovery state from the supplied/current ADC state, using `AppendRecoveryPayloadLine` and `Next` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   recoveryState (Input; Collection): Caller-supplied Collection used to provide or receive the ordered set of items associated with this operation.
' Returns: String: Returns the serialize generated form recovery state text value. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or
'   not resolved.
Public Function SerializeGeneratedFormRecoveryState(ByVal recoveryState As Collection) As String
    Dim physicalForms As Collection
    Dim backupEntries As Collection
    Dim originalForms As Collection
    Dim item As Variant
    SerializeGeneratedFormRecoveryState = "VERSION" & Chr$(30) & "1"
    If recoveryState Is Nothing Then Exit Function
    Set physicalForms = recoveryState("PhysicalForms")
    Set backupEntries = recoveryState("BackupEntries")
    Set originalForms = recoveryState("OriginalForms")
    For Each item In backupEntries
        SetupControlPump
        SerializeGeneratedFormRecoveryState = AppendRecoveryPayloadLine(SerializeGeneratedFormRecoveryState, "BACKUP" & Chr$(30) & CStr(item(0)) & Chr$(30) & CStr(item(1)))
    Next item
    For Each item In physicalForms
        SetupControlPump
        SerializeGeneratedFormRecoveryState = AppendRecoveryPayloadLine(SerializeGeneratedFormRecoveryState, "PHYSICAL" & Chr$(30) & CStr(item))
    Next item
    For Each item In originalForms
        SetupControlPump
        SerializeGeneratedFormRecoveryState = AppendRecoveryPayloadLine(SerializeGeneratedFormRecoveryState, "ORIGINAL" & Chr$(30) & CStr(item))
    Next item
    SerializeGeneratedFormRecoveryState = AppendRecoveryPayloadLine(SerializeGeneratedFormRecoveryState, "ROLLBACKTABLE" & Chr$(30) & CStr(recoveryState("CatalogRollbackTable")))
End Function

' Purpose: Performs stamp generated form build signature by coordinating `GeneratedFormExpectedBuildSignature` and `GeneratedFormPhysicalDefinitionHash` while preserving this module's existing validation and error
'     semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   formName (Input; String): Physical Access form name identifying the form to create, open, configure, validate, or inspect.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub StampGeneratedFormBuildSignature(ByVal formName As String)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim doc As DAO.Document
    Dim expectedSignature As String
    Dim definitionHash As String
    expectedSignature = GeneratedFormExpectedBuildSignature(formName)
    definitionHash = GeneratedFormPhysicalDefinitionHash(formName)
    Set db = GetHostDb()
    db.Containers(FORM_DAO_CONTAINER_NAME).Documents.Refresh
    Set doc = db.Containers(FORM_DAO_CONTAINER_NAME).Documents(formName)
    EnsureDaoPropertyValue doc, GENERATED_FORM_BUILD_SIGNATURE_PROPERTY, dbText, expectedSignature, "StampGeneratedFormBuildSignature"
    EnsureDaoPropertyValue doc, GENERATED_FORM_DEFINITION_HASH_PROPERTY, dbText, definitionHash, "StampGeneratedFormBuildSignature"
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "StampGeneratedFormBuildSignature", "Could not stamp generated-form build metadata." & vbCrLf & "Form: " & formName & vbCrLf & "Description: " & Err.Description
End Sub

Private Function TryRestoreChatGPTGeneratedFormsAfterFailure(ByVal physicalFormNames As Collection, ByVal backupEntries As Collection, ByVal originalFormNames As Collection, ByVal formCatalogRollbackTable As String, ByRef rollbackErrNumber As Long, ByRef rollbackErrDescription As String) As String
    rollbackErrNumber = 0
    rollbackErrDescription = vbNullString
    On Error GoTo ErrHandler
    TryRestoreChatGPTGeneratedFormsAfterFailure = RestoreChatGPTGeneratedFormsAfterFailure(physicalFormNames, backupEntries, originalFormNames, formCatalogRollbackTable)
    Exit Function
ErrHandler:
    Dim recoveryState As Collection
    rollbackErrNumber = Err.Number
    rollbackErrDescription = Err.Description
    If mUIRecoveryBlocked = False Then
        Set recoveryState = BuildGeneratedFormRecoveryState(physicalFormNames, backupEntries, originalFormNames, formCatalogRollbackTable)
        On Error Resume Next
        RegisterUIRecoveryFailure "Generated-form restoration", "Generated-form recovery terminated unexpectedly. Error " & CStr(rollbackErrNumber) & ": " & rollbackErrDescription, Nothing, recoveryState, UI_RECOVERY_TYPE_GENERATED_FORMS
        If Err.Number <> 0 Then rollbackErrDescription = rollbackErrDescription & vbCrLf & "Durable recovery registration also failed. Error " & CStr(Err.Number) & ": " & Err.Description
        On Error GoTo 0
    End If
    TryRestoreChatGPTGeneratedFormsAfterFailure = "FAILED. Error " & CStr(rollbackErrNumber) & ": " & rollbackErrDescription
End Function

Private Sub ValidateGeneratedFormSet(Optional ByVal suppliedHelpTopics As Variant, Optional ByVal suppliedSystemMacroDefinitions As Variant, Optional ByVal suppliedSystemMacroDefinitionIndex As Variant, Optional ByVal buildScope As String = FORM_BUILD_SCOPE_ALL)
    On Error GoTo ErrHandler
    Dim formNames As Collection
    Dim formItem As Variant
    Dim formName As String
    Dim validatedCount As Long
    If CodeDbIsHostDb() = False Then Err.Raise vbObjectError + 7420, "ValidateGeneratedFormSet", "Generated-form validation requires CodeDb and HostDb to be the same active database."
    Set formNames = GeneratedFormNamesForBuildScope(buildScope)
    For Each formItem In formNames
        SetupControlPump
        formName = CStr(formItem)
        SetChatGPTBuildContext "Strictly validating generated forms.", formName, "Validate the canonical generated-form contract.", "Build scope: " & buildScope
        ValidateOneGeneratedFormContract formName, suppliedHelpTopics, suppliedSystemMacroDefinitions, suppliedSystemMacroDefinitionIndex, formNames
        validatedCount = validatedCount + 1
    Next formItem
    LogTraceMessage TRACE_LEVEL_INFO, "UI", "ValidateGeneratedFormSet", "All generated forms in the selected build scope passed the shared validation framework.", "Forms validated: " & CStr(validatedCount) & vbCrLf & "Build scope: " & buildScope, "Validate generated forms", 0, vbNullString, vbNullString, DEFAULT_CHATGPT_UI_DISPLAY_PREFIX, "FormSet", "CodeDb", "FormGeneration"
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ValidateGeneratedFormSet", "Generated-form set validation failed." & vbCrLf & "Form: " & IIf(Len(formName) = 0, "(not selected)", formName) & vbCrLf & "Build scope: " & buildScope & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

' Purpose: Performs generated form record source for lookup editor proof by coordinating `SharedDao_AccessObjectExists` and `End` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   formName (Input; String): Physical Access form name identifying the form to create, open, configure, validate, or inspect.
' Returns: String: Returns the resolved generated form record source for lookup editor proof text in the canonical representation consumed by this module's caller; vbNullString represents no applicable value
'   where the existing contract permits it.
Public Function GeneratedFormRecordSourceForLookupEditorProof(ByVal formName As String) As String
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim frm As Access.Form
    Dim openedHere As Boolean
    If Len(formName) = 0 Then Err.Raise vbObjectError + 7446, "GeneratedFormRecordSourceForLookupEditorProof", "A generated-form name is required for lookup-editor RecordSource proof."
    Set db = GetHostDb()
    If SharedDao_AccessObjectExists(db, acForm, formName) = False Then Err.Raise vbObjectError + 7447, "GeneratedFormRecordSourceForLookupEditorProof", "The generated form required for lookup-editor RecordSource proof is missing." & vbCrLf & "Form: " & formName
    If CurrentProject.AllForms(formName).IsLoaded = False Then
        DoCmd.OpenForm formName, acDesign, , , , acHidden
        openedHere = True
    End If
    Set frm = forms(formName)
    GeneratedFormRecordSourceForLookupEditorProof = Trim$(Nz(frm.recordSource, vbNullString))
    If openedHere Then DoCmd.Close acForm, formName, acSaveNo
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If openedHere And CurrentProject.AllForms(formName).IsLoaded Then DoCmd.Close acForm, formName, acSaveNo
    On Error GoTo 0
    Err.Raise savedErrNumber, "GeneratedFormRecordSourceForLookupEditorProof", "Could not read the physical generated-form RecordSource for lookup-editor proof." & vbCrLf & "Form: " & formName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Sub ValidateOneGeneratedFormContract(ByVal formName As String, Optional ByVal suppliedHelpTopics As Variant, Optional ByVal suppliedSystemMacroDefinitions As Variant, Optional ByVal suppliedSystemMacroDefinitionIndex As Variant, Optional ByVal suppliedRequiredRelatedFormNames As Variant)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim frm As Access.Form
    Dim openedHere As Boolean
    Dim expectedSourceType As String
    Dim expectedSourceName As String
    Dim actualRecordSource As String
    Dim expectedOnOpen As String
    Dim expectedOnLoad As String
    Dim expectedOnCurrent As String
    If Len(formName) = 0 Then Err.Raise vbObjectError + 7421, "ValidateOneGeneratedFormContract", "A canonical generated-form name is blank."
    Set db = GetHostDb()
    If SharedDao_AccessObjectExists(db, acForm, formName) = False Then Err.Raise vbObjectError + 7422, "ValidateOneGeneratedFormContract", "A canonical generated form is missing." & vbCrLf & "Form: " & formName
    If CurrentProject.AllForms(formName).IsLoaded = False Then
        DoCmd.OpenForm formName, acDesign, , , , acHidden
        openedHere = True
    End If
    Set frm = forms(formName)
    expectedSourceType = GeneratedFormExpectedSourceType(formName)
    expectedSourceName = GeneratedFormSourceNameForSignature(formName)
    actualRecordSource = Trim$(Nz(frm.recordSource, vbNullString))
    If StrComp(expectedSourceType, GENERATED_FORM_SOURCE_TYPE_UNBOUND, vbTextCompare) = 0 Then
        If Len(actualRecordSource) > 0 Then Err.Raise vbObjectError + 7423, "ValidateOneGeneratedFormContract", "A canonical unbound generated form has a RecordSource." & vbCrLf & "Form: " & formName & vbCrLf & "RecordSource: " & actualRecordSource
    Else
        If Len(expectedSourceName) = 0 Then Err.Raise vbObjectError + 7424, "ValidateOneGeneratedFormContract", "A canonical bound generated form has no resolved authoritative source." & vbCrLf & "Form: " & formName
        If Len(actualRecordSource) = 0 Then Err.Raise vbObjectError + 7425, "ValidateOneGeneratedFormContract", "A canonical bound generated form has a blank RecordSource." & vbCrLf & "Form: " & formName & vbCrLf & "Expected source: " & expectedSourceName
        If InStr(1, actualRecordSource, expectedSourceName, vbTextCompare) = 0 Then Err.Raise vbObjectError + 7426, "ValidateOneGeneratedFormContract", "A canonical generated form does not use its authoritative source." & vbCrLf & "Form: " & formName & vbCrLf & "Expected source: " & expectedSourceName & vbCrLf & "Actual RecordSource: " & actualRecordSource
    End If
    expectedOnOpen = GeneratedFormContractOnOpenExpression(formName)
    If Len(expectedOnOpen) > 0 Then
        If StrComp(Trim$(Nz(frm.OnOpen, vbNullString)), expectedOnOpen, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7434, "ValidateOneGeneratedFormContract", "A generated form OnOpen event does not match its canonical contract." & vbCrLf & "Form: " & formName & vbCrLf & "Expected OnOpen: " & expectedOnOpen & vbCrLf & "Actual OnOpen: " & Nz(frm.OnOpen, vbNullString)
    End If
    expectedOnLoad = GeneratedFormContractOnLoadExpression(formName)
    If Len(expectedOnLoad) > 0 Then
        If StrComp(Trim$(Nz(frm.OnLoad, vbNullString)), expectedOnLoad, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7427, "ValidateOneGeneratedFormContract", "A generated form OnLoad event does not match its canonical contract." & vbCrLf & "Form: " & formName & vbCrLf & "Expected OnLoad: " & expectedOnLoad & vbCrLf & "Actual OnLoad: " & Nz(frm.OnLoad, vbNullString)
    End If
    expectedOnCurrent = GeneratedFormOnCurrentExpression(formName)
    If Len(expectedOnCurrent) > 0 Then
        If StrComp(Trim$(Nz(frm.OnCurrent, vbNullString)), expectedOnCurrent, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7428, "ValidateOneGeneratedFormContract", "A generated form OnCurrent event does not match its canonical contract." & vbCrLf & "Form: " & formName & vbCrLf & "Expected OnCurrent: " & expectedOnCurrent & vbCrLf & "Actual OnCurrent: " & Nz(frm.OnCurrent, vbNullString)
    End If
    ValidateGeneratedFormDeclaredControls frm, formName
    ValidateGeneratedFormDeclaredCallbacks frm, formName
    ValidateGeneratedFormLookupEditorMetadata frm, formName, expectedSourceName
    If LookupEditorGeneratedFormIsDefined(formName) Then ValidateLookupEditorGeneratedFormSpecialized frm, formName
    If AddInReviewChildFormIsDefinedDefault(formName) Then ValidateGeneratedAddInReviewChildSpecialized frm, formName
    If StrComp(formName, ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review"), vbTextCompare) = 0 Then ValidateGeneratedFormAddInHostSpecialized frm
    If StrComp(formName, ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Help"), vbTextCompare) = 0 Then ValidateGeneratedFormHelpSpecialized db, suppliedHelpTopics, suppliedSystemMacroDefinitions, suppliedSystemMacroDefinitionIndex, suppliedRequiredRelatedFormNames
    If openedHere Then DoCmd.Close acForm, formName, acSaveNo
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If openedHere And CurrentProject.AllForms(formName).IsLoaded Then DoCmd.Close acForm, formName, acSaveNo
    On Error GoTo 0
    Err.Raise savedErrNumber, "ValidateOneGeneratedFormContract", "Generated-form contract validation failed." & vbCrLf & "Form: " & formName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub ValidateGeneratedFormLookupEditorMetadata(ByVal frm As Access.Form, ByVal formName As String, ByVal sourceName As String)
    Dim ctl As Access.control
    Dim fieldName As String
    Dim definitionData As Variant
    Dim lookupKind As String
    Dim policyCode As String
    Dim expectedEditorForm As String
    Dim expectedEditorAssigned As Boolean
    For Each ctl In frm.Controls
        If ctl.controlType = acComboBox Then
            fieldName = Trim$(Nz(ctl.controlSource, vbNullString))
            If Len(fieldName) > 0 Then
                definitionData = GeneratedFormLookupDefinition(fieldName, sourceName)
                If IsArray(definitionData) Then
                    lookupKind = CStr(definitionData(0))
                    expectedEditorForm = vbNullString
                    expectedEditorAssigned = False
                    If StrComp(lookupKind, GENERATED_FORM_LOOKUP_KIND_TWO_COLUMN, vbTextCompare) = 0 Then
                        expectedEditorForm = LookupEditorPhysicalFormNameForSource(CStr(definitionData(1)), policyCode)
                        If StrComp(policyCode, LOOKUP_EDITOR_POLICY_NONE, vbTextCompare) <> 0 Then
                            If Len(Trim$(expectedEditorForm)) = 0 Then Err.Raise vbObjectError + 7461, "ValidateGeneratedFormLookupEditorMetadata", "A two-column generated lookup has authoritative editor policy but no canonical editor form identity." & vbCrLf & "Form: " & formName & vbCrLf & "Control: " & ctl.Name & vbCrLf & "Field: " & fieldName & vbCrLf & "Lookup source: " & CStr(definitionData(1)) & vbCrLf & "Policy: " & policyCode
                            expectedEditorAssigned = True
                        End If
                    End If
                    If expectedEditorAssigned Then
                        If ctl.AllowValueListEdits = False Then Err.Raise vbObjectError + 7435, "ValidateGeneratedFormLookupEditorMetadata", "A generated combo box with an authoritative lookup editor does not allow native Edit List Items." & vbCrLf & "Form: " & formName & vbCrLf & "Control: " & ctl.Name & vbCrLf & "Field: " & fieldName & vbCrLf & "Expected editor: " & expectedEditorForm
                        If StrComp(Trim$(Nz(ctl.ListItemsEditForm, vbNullString)), expectedEditorForm, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 7436, "ValidateGeneratedFormLookupEditorMetadata", "A generated combo box does not use the authoritative lookup editor." & vbCrLf & "Form: " & formName & vbCrLf & "Control: " & ctl.Name & vbCrLf & "Field: " & fieldName & vbCrLf & "Expected editor: " & expectedEditorForm & vbCrLf & "Actual editor: " & Nz(ctl.ListItemsEditForm, vbNullString)
                    Else
                        If ctl.AllowValueListEdits Or Len(Trim$(Nz(ctl.ListItemsEditForm, vbNullString))) > 0 Then Err.Raise vbObjectError + 7437, "ValidateGeneratedFormLookupEditorMetadata", "A generated combo box without authoritative editor intent has native Edit List Items metadata assigned." & vbCrLf & "Form: " & formName & vbCrLf & "Control: " & ctl.Name & vbCrLf & "Field: " & fieldName & vbCrLf & "Lookup kind: " & lookupKind
                    End If
                End If
            End If
        End If
    Next ctl
End Sub

Private Sub ValidateLookupEditorGeneratedFormSpecialized(ByVal frm As Access.Form, ByVal formName As String)
    Dim definitionData As Variant
    Dim codeFieldName As String
    Dim allowCodeEdit As Boolean
    Dim codeControlFound As Boolean
    Dim ctl As Access.control
    definitionData = LookupEditorGeneratedFormDefinition(formName)
    If IsArray(definitionData) = False Then Exit Sub
    If frm.defaultView <> acDefViewContinuous Then Err.Raise vbObjectError + 7438, "ValidateLookupEditorGeneratedFormSpecialized", "A generated lookup editor is not a continuous form." & vbCrLf & "Form: " & formName
    If frm.allowEdits = False Then Err.Raise vbObjectError + 7439, "ValidateLookupEditorGeneratedFormSpecialized", "A generated lookup editor is unexpectedly read-only." & vbCrLf & "Form: " & formName
    If frm.allowAdditions <> CBool(definitionData(LOOKUP_EDITOR_DEF_ALLOW_ADDITIONS)) Then Err.Raise vbObjectError + 7444, "ValidateLookupEditorGeneratedFormSpecialized", "A generated lookup editor has the wrong AllowAdditions policy." & vbCrLf & "Form: " & formName
    If frm.allowDeletions <> CBool(definitionData(LOOKUP_EDITOR_DEF_ALLOW_DELETIONS)) Then Err.Raise vbObjectError + 7445, "ValidateLookupEditorGeneratedFormSpecialized", "A generated lookup editor has the wrong AllowDeletions policy." & vbCrLf & "Form: " & formName
    If StrComp(Trim$(Nz(frm.OnLoad, vbNullString)), LOOKUP_EDITOR_RUNTIME_POLICY_EXPRESSION, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7457, "ValidateLookupEditorGeneratedFormSpecialized", "A generated lookup editor does not apply the persisted runtime edit policy on load." & vbCrLf & "Form: " & formName & vbCrLf & "Expected OnLoad: " & LOOKUP_EDITOR_RUNTIME_POLICY_EXPRESSION & vbCrLf & "Actual OnLoad: " & Nz(frm.OnLoad, vbNullString)
    If StrComp(Trim$(Nz(frm.OnDelete, vbNullString)), LOOKUP_EDITOR_DELETE_GUARD_EXPRESSION, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7449, "ValidateLookupEditorGeneratedFormSpecialized", "A generated lookup editor does not use the canonical referenced-value delete guard." & vbCrLf & "Form: " & formName & vbCrLf & "Expected OnDelete: " & LOOKUP_EDITOR_DELETE_GUARD_EXPRESSION & vbCrLf & "Actual OnDelete: " & Nz(frm.OnDelete, vbNullString)
    If StrComp(Trim$(Nz(frm.OnCurrent, vbNullString)), LOOKUP_EDITOR_CODE_FIELD_EDIT_POLICY_EXPRESSION, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7452, "ValidateLookupEditorGeneratedFormSpecialized", "A generated lookup editor does not use the canonical current-record key/delete policy." & vbCrLf & "Form: " & formName & vbCrLf & "Expected OnCurrent: " & LOOKUP_EDITOR_CODE_FIELD_EDIT_POLICY_EXPRESSION & vbCrLf & "Actual OnCurrent: " & Nz(frm.OnCurrent, vbNullString)
    If StrComp(Trim$(Nz(frm.AfterInsert, vbNullString)), LOOKUP_EDITOR_CODE_FIELD_EDIT_POLICY_EXPRESSION, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7453, "ValidateLookupEditorGeneratedFormSpecialized", "A generated lookup editor does not reapply the canonical current-record key/delete policy after insertion." & vbCrLf & "Form: " & formName & vbCrLf & "Expected AfterInsert: " & LOOKUP_EDITOR_CODE_FIELD_EDIT_POLICY_EXPRESSION & vbCrLf & "Actual AfterInsert: " & Nz(frm.AfterInsert, vbNullString)
    codeFieldName = Trim$(CStr(definitionData(LOOKUP_EDITOR_DEF_CODE_FIELD)))
    allowCodeEdit = CBool(definitionData(LOOKUP_EDITOR_DEF_ALLOW_CODE_EDIT))
    If Len(codeFieldName) = 0 Then Exit Sub
    For Each ctl In frm.Controls
        If ctl.controlType = acTextBox Or ctl.controlType = acComboBox Or ctl.controlType = acCheckBox Then
            If StrComp(Trim$(Nz(ctl.controlSource, vbNullString)), codeFieldName, vbTextCompare) = 0 Then
                codeControlFound = True
                If allowCodeEdit = False Then
                    If ctl.Locked = False Then Err.Raise vbObjectError + 7442, "ValidateLookupEditorGeneratedFormSpecialized", "A protected lookup-editor code field is editable." & vbCrLf & "Form: " & formName & vbCrLf & "Field: " & codeFieldName & vbCrLf & "Control: " & ctl.Name
                End If
                Exit For
            End If
        End If
    Next ctl
    If codeControlFound = False Then Err.Raise vbObjectError + 7443, "ValidateLookupEditorGeneratedFormSpecialized", "A generated lookup editor is missing its authoritative code/key control." & vbCrLf & "Form: " & formName & vbCrLf & "Field: " & codeFieldName
End Sub

' Purpose: Applies the persisted lookup-editor code-field and deletion policy to the current generated lookup-editor form record.
' Public role: Form-code/event-expression public entry point. Access form code or an event property can invoke this routine by name; its Public visibility is part of that UI contract.
' Parameters: None.
' Returns: Boolean: True when the current generated lookup-editor record policy is applied successfully; False when the callback handles an error and cannot complete the policy application.
Public Function ApplyLookupEditorCodeFieldEditPolicy() As Boolean
    On Error GoTo ErrHandler
    Dim frm As Access.Form
    Dim definitionData As Variant
    Dim codeFieldName As String
    Dim ctl As Access.control
    Dim fieldName As String
    Dim codeControlFound As Boolean
    Dim allowCodeEdit As Boolean, allowDeletionPolicy As Boolean, logicalID As String
    Set frm = Application.CodeContextObject
    LogPhysicalTraceMessageOnly "Info", "LookupEditorRuntime", "ApplyLookupEditorCodeFieldEditPolicy", "START", "Form: " & frm.Name, "Apply lookup-editor current-record policy", 0, vbNullString, vbNullString, frm.Name, "Form", "CodeDb", "LookupEditor"
    definitionData = LookupEditorGeneratedFormDefinition(frm.Name)
    If IsArray(definitionData) = False Then Err.Raise vbObjectError + 7454, "ApplyLookupEditorCodeFieldEditPolicy", "The lookup-editor current-record policy was invoked outside a canonical generated lookup editor." & vbCrLf & "Form: " & frm.Name
    logicalID = CStr(definitionData(LOOKUP_EDITOR_DEF_SOURCE_LOGICAL_ID))
    allowCodeEdit = LookupEditorPersistedBooleanPolicy(logicalID, "AllowCodeEdits")
    allowDeletionPolicy = LookupEditorPersistedBooleanPolicy(logicalID, "AllowDeletions")
    codeFieldName = Trim$(CStr(definitionData(LOOKUP_EDITOR_DEF_CODE_FIELD)))
    For Each ctl In frm.Controls
        Select Case ctl.controlType
            Case acTextBox, acComboBox, acCheckBox
                fieldName = Trim$(Nz(ctl.controlSource, vbNullString))
                If StrComp(fieldName, codeFieldName, vbTextCompare) = 0 Then
                    codeControlFound = True
                    ctl.Locked = (Not frm.newRecord) And (Not allowCodeEdit)
                End If
        End Select
    Next ctl
    If codeControlFound = False Then Err.Raise vbObjectError + 7455, "ApplyLookupEditorCodeFieldEditPolicy", "The lookup-editor current-record policy could not find the authoritative key control." & vbCrLf & "Form: " & frm.Name & vbCrLf & "Key field: " & codeFieldName
    frm.allowDeletions = allowDeletionPolicy And (Not frm.newRecord)
    LogPhysicalTraceMessageOnly "Info", "LookupEditorRuntime", "ApplyLookupEditorCodeFieldEditPolicy", "DONE", "Form: " & frm.Name & vbCrLf & "AllowDeletions: " & CStr(frm.allowDeletions), "Apply lookup-editor current-record policy", 0, vbNullString, vbNullString, frm.Name, "Form", "CodeDb", "LookupEditor"
    ApplyLookupEditorCodeFieldEditPolicy = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim traceFormName As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not frm Is Nothing Then traceFormName = frm.Name
    On Error GoTo 0
    LogPhysicalTraceMessageOnly "Error", "LookupEditorRuntime", "ApplyLookupEditorCodeFieldEditPolicy", "ERROR", "Form: " & IIf(Len(traceFormName) = 0, "(unavailable)", traceFormName), "Apply lookup-editor current-record policy", savedErrNumber, savedErrSource, savedErrDescription, traceFormName, "Form", "CodeDb", "LookupEditor"
    On Error Resume Next
    If Not frm Is Nothing Then frm.allowDeletions = False
    On Error GoTo 0
    CoreShowErrorMessage "ADC could not apply the lookup-editor current-record policy. Deletion has been disabled for safety." & vbCrLf & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, vbCritical, "Lookup Editor Policy", "ApplyLookupEditorCodeFieldEditPolicy", "UI"
End Function

Private Function LookupEditorPersistedBooleanPolicy(ByVal logicalID As String, ByVal fieldName As String) As Boolean
    On Error GoTo failed
    Dim rs As DAO.Recordset
    Set rs = ADCDb().OpenRecordset("SELECT " & SQL_FIELD_EDITOR_MODE & ", [" & fieldName & "] AS [PolicyValue] FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LookupDefinitions") & "] WHERE " & SQL_FIELD_LOOKUP_LOGICAL_ID & "='" & Replace(logicalID, "'", "''") & "';", dbOpenSnapshot)
    If rs.EOF Then GoTo failed
    If StrComp(Trim$(Nz(rs.fields(FIELD_EDITOR_MODE), vbNullString)), "Editable", vbTextCompare) <> 0 Then GoTo failed
    LookupEditorPersistedBooleanPolicy = CBool(Nz(rs!PolicyValue, False))
    rs.Close
    Exit Function
failed:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    LookupEditorPersistedBooleanPolicy = False
End Function

' Purpose: Validates whether deletion of the current generated lookup-editor record is allowed and blocks the form delete event when policy or dependency rules prohibit it.
' Public role: Form-code/event-expression public entry point. Access form code or an event property can invoke this routine by name; its Public visibility is part of that UI contract.
' Parameters: None.
' Returns: Boolean: True when the current record deletion is permitted to proceed; False when policy/dependency checks block or cancel the delete.
Public Function GuardLookupEditorDelete() As Boolean
    On Error GoTo ErrHandler
    Dim frm As Access.Form
    Dim definitionData As Variant
    Dim sourceTableName As String
    Dim codeFieldName As String
    Dim rowValues As Object
    Dim ctl As Access.control
    Dim fieldName As String
    Dim usageDetail As String
    Set frm = Application.CodeContextObject
    definitionData = LookupEditorGeneratedFormDefinition(frm.Name)
    If IsArray(definitionData) = False Then Err.Raise vbObjectError + 7450, "GuardLookupEditorDelete", "The delete guard was invoked outside a canonical generated lookup editor." & vbCrLf & "Form: " & frm.Name
    sourceTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, CStr(definitionData(LOOKUP_EDITOR_DEF_SOURCE_LOGICAL_ID)))
    codeFieldName = Trim$(CStr(definitionData(LOOKUP_EDITOR_DEF_CODE_FIELD)))
    Set rowValues = CreateObject("Scripting.Dictionary")
    rowValues.CompareMode = vbTextCompare
    For Each ctl In frm.Controls
        Select Case ctl.controlType
            Case acTextBox, acComboBox, acCheckBox
                fieldName = Trim$(Nz(ctl.controlSource, vbNullString))
                If Len(fieldName) > 0 And Left$(fieldName, 1) <> "=" Then rowValues(fieldName) = ctl.value
        End Select
    Next ctl
    If Len(codeFieldName) = 0 Or rowValues.Exists(codeFieldName) = False Then Err.Raise vbObjectError + 7451, "GuardLookupEditorDelete", "The lookup editor delete guard could not read the authoritative lookup key from the current row." & vbCrLf & "Form: " & frm.Name & vbCrLf & "Source table: " & sourceTableName & vbCrLf & "Key field: " & codeFieldName
    If LookupRecordReferenceUsage(sourceTableName, rowValues, usageDetail) Then
        DoCmd.CancelEvent
        CoreShowMessage "This lookup value cannot be deleted because it is currently used by another ADC record." & vbCrLf & vbCrLf & usageDetail, vbExclamation, "Lookup Value In Use", "GuardLookupEditorDelete", "UI"
        Exit Function
    End If
    GuardLookupEditorDelete = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    DoCmd.CancelEvent
    On Error GoTo 0
    CoreShowErrorMessage "The lookup value was not deleted because ADC could not verify that it is unused." & vbCrLf & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, vbCritical, "Lookup Delete Check", "GuardLookupEditorDelete", "UI"
End Function

Private Sub ValidateGeneratedFormDeclaredControls(ByVal frm As Access.Form, ByVal formName As String)
    Dim definitions As Collection
    Dim definitionData As Variant
    Dim ctl As Access.control
    Dim controlName As String
    Dim requireUnbound As Boolean
    Dim requireValueList As Boolean
    Dim requireLocked As Boolean
    Set definitions = GeneratedFormContractControlDefinitions(formName)
    For Each definitionData In definitions
        SetupControlPump
        controlName = CStr(definitionData(0))
        requireUnbound = CBool(definitionData(1))
        requireValueList = CBool(definitionData(2))
        requireLocked = CBool(definitionData(3))
        Set ctl = Nothing
        On Error Resume Next
        Set ctl = frm.Controls(controlName)
        On Error GoTo 0
        If ctl Is Nothing Then Err.Raise vbObjectError + 7429, "ValidateGeneratedFormDeclaredControls", "A generated form is missing a control required by its canonical contract." & vbCrLf & "Form: " & formName & vbCrLf & "Control: " & controlName
        If requireUnbound Then
            If Len(Trim$(Nz(ctl.controlSource, vbNullString))) > 0 Then Err.Raise vbObjectError + 7430, "ValidateGeneratedFormDeclaredControls", "A generated-form control required to remain unbound has a ControlSource." & vbCrLf & "Form: " & formName & vbCrLf & "Control: " & controlName & vbCrLf & "ControlSource: " & Nz(ctl.controlSource, vbNullString)
        End If
        If requireValueList Then
            If StrComp(Trim$(Nz(ctl.rowSourceType, vbNullString)), "Value List", vbTextCompare) <> 0 Then Err.Raise vbObjectError + 7431, "ValidateGeneratedFormDeclaredControls", "A generated-form control does not use the required Value List RowSourceType." & vbCrLf & "Form: " & formName & vbCrLf & "Control: " & controlName
        End If
        If requireLocked Then
            If ctl.Locked = False Then Err.Raise vbObjectError + 7432, "ValidateGeneratedFormDeclaredControls", "A generated-form control required to remain read-only is unlocked." & vbCrLf & "Form: " & formName & vbCrLf & "Control: " & controlName
        End If
    Next definitionData
End Sub

Private Sub ValidateGeneratedFormDeclaredCallbacks(ByVal frm As Access.Form, ByVal formName As String)
    Dim expressions As Collection
    Dim expectedExpression As Variant
    Dim foundExpressions As Object
    Dim ctl As Access.control
    Dim clickText As String
    Set expressions = GeneratedFormContractCallbackExpressions(formName)
    If expressions.count = 0 Then Exit Sub
    ' INTENTIONAL KEYED LOOKUP: foundExpressions provides keyed membership/lookup over source rows or state already read by ValidateGeneratedFormDeclaredCallbacks; it is a derived access path, not a second
    ' authority.
    Set foundExpressions = CreateObject("Scripting.Dictionary")
    foundExpressions.CompareMode = vbBinaryCompare
    For Each expectedExpression In expressions
        foundExpressions.Add CStr(expectedExpression), False
    Next expectedExpression
    For Each ctl In frm.Controls
        SetupControlPump
        If ctl.controlType = acCommandButton Then
            clickText = Trim$(Nz(ctl.OnClick, vbNullString))
            If foundExpressions.Exists(clickText) Then foundExpressions(clickText) = True
        End If
    Next ctl
    For Each expectedExpression In expressions
        If foundExpressions(CStr(expectedExpression)) = False Then Err.Raise vbObjectError + 7433, "ValidateGeneratedFormDeclaredCallbacks", "A generated form is missing a callback required by its canonical contract." & vbCrLf & "Form: " & formName & vbCrLf & "OnClick: " & CStr(expectedExpression)
    Next expectedExpression
End Sub

Private Sub ValidateGeneratedFormAddInHostSpecialized(ByVal frm As Access.Form)
    Dim addInReviewChildForms As Variant
    Dim addInReviewChildForm As Variant
    Dim subformControlName As String
    Dim ctl As Access.control
    Dim selectedProjectControlName As String
    If SharedDao_QueryExists(GetHostDb(), ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Review_Summary")) = False Then Err.Raise vbObjectError + 7293, "ValidateGeneratedFormAddInHostSpecialized", "The required Add-Ins review summary query is missing." & vbCrLf & "Query: " & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Review_Summary")
    selectedProjectControlName = GeneratedFormContractControlName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review"), 1)
    addInReviewChildForms = AddInReviewChildFormNames()
    For Each addInReviewChildForm In addInReviewChildForms
        SetupControlPump
        subformControlName = AddInReviewChildSubformControlName(CStr(addInReviewChildForm))
        Set ctl = Nothing
        On Error Resume Next
        Set ctl = frm.Controls(subformControlName)
        On Error GoTo 0
        If ctl Is Nothing Then Err.Raise vbObjectError + 7300, "ValidateGeneratedFormAddInHostSpecialized", "A required Add-Ins review subform control is missing." & vbCrLf & "Control: " & subformControlName
        If InStr(1, Nz(ctl.sourceObject, vbNullString), CStr(addInReviewChildForm), vbTextCompare) = 0 Then Err.Raise vbObjectError + 7301, "ValidateGeneratedFormAddInHostSpecialized", "An Add-Ins review subform control does not use its authoritative generated source form." & vbCrLf & "Control: " & subformControlName & vbCrLf & "Expected form: " & CStr(addInReviewChildForm) & vbCrLf & "Actual SourceObject: " & Nz(ctl.sourceObject, vbNullString)
        If AddInReviewChildLinksSelectedProject(CStr(addInReviewChildForm)) Then
            If StrComp(Nz(ctl.LinkMasterFields, vbNullString), selectedProjectControlName, vbTextCompare) <> 0 Or StrComp(Nz(ctl.LinkChildFields, vbNullString), FIELD_PROJECT_IDENTITY, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 7302, "ValidateGeneratedFormAddInHostSpecialized", "An Add-Ins review subform is not linked through the stable ProjectIdentity key." & vbCrLf & "Control: " & subformControlName
        End If
    Next addInReviewChildForm
End Sub

' Purpose: Checks whether generated form partition already satisfies the exact state required by the current Setup run so the owning phase can determine whether repair is necessary.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   buildScope (Input; String): Scope value limiting which build items or operations are included in this call.
'   detailText (Output; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
' Returns: Boolean: True when the condition represented by UtilityLazyLoadSelectedTab is present under the routine's documented inputs/current state; False when that condition is absent.
Public Function GeneratedFormPartitionCurrentForSetup(ByVal buildScope As String, ByRef detailText As String) As Boolean
    'Mutation-free current-state proof for one required Setup-generated-form partition. UI remains the authoritative generated-form proof owner.
    On Error GoTo ErrHandler
    Dim normalizedBuildScope As String
    Dim completeSetToken As String
    Dim physicalStateToken As String
    normalizedBuildScope = NormalizeChatGPTBuildScope(buildScope)
    If StrComp(normalizedBuildScope, FORM_BUILD_SCOPE_CHATGPT_UI_FORMS, vbTextCompare) <> 0 And _
       StrComp(normalizedBuildScope, FORM_BUILD_SCOPE_SUPPORT_FORMS, vbTextCompare) <> 0 Then
        detailText = "Generated-form partition currentness supports only ChatGPTUIForms or ADCSupportForms. Requested scope=" & normalizedBuildScope
        Exit Function
    End If
    If GeneratedFormCompactPreflightMatches(normalizedBuildScope, completeSetToken, physicalStateToken, detailText) = False Then Exit Function
    detailText = "Generated-form partition is current. Scope=" & normalizedBuildScope & "; complete-set token=" & completeSetToken & "; physical-state token=" & physicalStateToken & "."
    GeneratedFormPartitionCurrentForSetup = True
    Exit Function
ErrHandler:
    detailText = "Could not verify the required generated-form partition for Setup. Scope=" & buildScope & ". Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
End Function

' Purpose: Checks whether generated forms already satisfies the exact state required by the current Setup run so the owning phase can determine whether repair is necessary.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   detailText (Input/output; String): Receives or supplies diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
' Returns: Boolean: True when generated forms matches the exact state required by the current Setup contract; False when the owning phase still needs repair, refresh, or reconciliation.
Public Function GeneratedFormsCurrentForSetup(ByRef detailText As String) As Boolean
    ' Full-current-state verification path used when D4 cannot rely solely on same-run in-memory evidence. Each authoritative partition owns an independent persisted proof.
    On Error GoTo ErrHandler
    Dim chatGPTUIToken As String
    Dim supportToken As String
    Dim allCount As Long
    Dim chatGPTUICount As Long
    Dim supportCount As Long
    Dim allFormNames As Collection
    chatGPTUIToken = GeneratedFormPostCommitToken(FORM_BUILD_SCOPE_CHATGPT_UI_FORMS)
    If Len(chatGPTUIToken) = 0 Then detailText = "The ChatGPT UI Form partition is not proven current.": Exit Function
    supportToken = GeneratedFormPostCommitToken(FORM_BUILD_SCOPE_SUPPORT_FORMS)
    If Len(supportToken) = 0 Then detailText = "The ADC Support Form partition is not proven current.": Exit Function
    allCount = GeneratedFormAuthoritativeCount(FORM_BUILD_SCOPE_ALL)
    chatGPTUICount = GeneratedFormAuthoritativeCount(FORM_BUILD_SCOPE_CHATGPT_UI_FORMS)
    supportCount = GeneratedFormAuthoritativeCount(FORM_BUILD_SCOPE_SUPPORT_FORMS)
    If allCount <= 0 Or chatGPTUICount + supportCount <> allCount Then
        detailText = "The authoritative generated-form classification does not partition the complete required roster exactly once. All=" & CStr(allCount) & "; ChatGPTUI=" & CStr(chatGPTUICount) & "; ADCSupport=" & CStr(supportCount)
        Exit Function
    End If
    Set allFormNames = GeneratedFormNamesForBuildScope(FORM_BUILD_SCOPE_ALL)
    ValidateOneGeneratedFormContract ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Help"), suppliedRequiredRelatedFormNames:=allFormNames
    detailText = "Both required generated-form partitions are current and their disjoint ChatGPTUIForm Boolean classification covers the complete authoritative roster. ChatGPT UI Forms=" & CStr(chatGPTUICount) & "; ADC Support Forms=" & CStr(supportCount) & "; Total=" & CStr(allCount) & "."
    GeneratedFormsCurrentForSetup = True
    Exit Function
ErrHandler:
    detailText = "Could not verify both generated-form partitions required by Setup D4. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
End Function

Private Function ChatGPTPhysicalFormChangeNamesRequiringBuild(ByVal replaceExisting As Boolean, Optional ByVal normalizedBuildScope As String = FORM_BUILD_SCOPE_ALL) As Collection
    'Returns only forms in scope whose deterministic signature or physical definition requires creation or replacement.
    Dim result As Collection
    Dim formNames As Collection
    Dim formItem As Variant
    Dim formName As String
    ' INTENTIONAL FILTERED PARTITION: result is a consumer-required subset/projection derived from the source enumerated by ChatGPTPhysicalFormChangeNamesRequiringBuild; its membership is not independently
    ' authoritative.
    Set result = New Collection
    Set formNames = GeneratedFormNamesForBuildScope(normalizedBuildScope)
    For Each formItem In formNames
        SetupControlPump
        formName = CStr(formItem)
        If SharedDao_AccessObjectExists(GetHostDb(), acForm, formName) = False Then
            result.Add formName
        ElseIf replaceExisting And GeneratedFormBuildSignatureIsCurrent(formName) = False Then
            result.Add formName
        End If
    Next formItem
    Set ChatGPTPhysicalFormChangeNamesRequiringBuild = result
End Function

' Purpose: Loads or refreshes the content required by the currently selected utility tab when Access invokes the generated form callback.
' Public role: Form-code/event-expression public entry point. Access form code or an event property can invoke this routine by name; its Public visibility is part of that UI contract.
' Parameters: None.
' Returns: Boolean: True when the selected utility tab is loaded/refreshed successfully; False when the callback handles an error or cannot complete the requested load.
Public Function UtilityLazyLoadSelectedTab() As Boolean
    On Error GoTo ErrHandler
    Dim frm As Access.Form
    Dim tabCtl As Access.control
    Dim subCtl As Access.control
    Dim formName As String
    Dim pageName As String
    Dim controlName As String
    Dim sourceFormName As String
    formName = ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Utility")
    If CurrentProject.AllForms(formName).IsLoaded = False Then Exit Function
    Set frm = forms(formName)
    Set tabCtl = frm.Controls("tabADCUtility")
    pageName = tabCtl.pages(CLng(tabCtl.value)).Name
    controlName = "sub_" & pageName
    On Error Resume Next
    Set subCtl = frm.Controls(controlName)
    On Error GoTo ErrHandler
    If subCtl Is Nothing Then
        UtilityLazyLoadSelectedTab = True
        Exit Function
    End If
    If Len(Trim$(Nz(subCtl.sourceObject, vbNullString))) > 0 Then
        UtilityLazyLoadSelectedTab = True
        Exit Function
    End If
    sourceFormName = Trim$(Nz(subCtl.Tag, vbNullString))
    If Len(sourceFormName) = 0 Then Err.Raise vbObjectError + 7440, "UtilityLazyLoadSelectedTab", "The selected Utility tab has no stored source-form identity." & vbCrLf & "Page: " & pageName & vbCrLf & "Control: " & controlName
    If SharedDao_AccessObjectExists(GetHostDb(), acForm, sourceFormName) = False Then Err.Raise vbObjectError + 7441, "UtilityLazyLoadSelectedTab", "The selected Utility tab source form does not exist." & vbCrLf & "Page: " & pageName & vbCrLf & "Source form: " & sourceFormName
    subCtl.sourceObject = "Form." & sourceFormName
    LogTraceMessage TRACE_LEVEL_INFO, "UI", "UtilityLazyLoadSelectedTab", "Loaded the selected ChatGPT Utility tab on demand.", "Page: " & pageName & " | Source form: " & sourceFormName, "Load Utility tab", 0, vbNullString, vbNullString, sourceFormName, "Form", "CodeDb", "GeneratedFormCommand"
    UtilityLazyLoadSelectedTab = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    LogTraceMessage TRACE_LEVEL_ERROR, "UI", "UtilityLazyLoadSelectedTab", "Could not load the selected ChatGPT Utility tab.", "Page: " & pageName & " | Control: " & controlName & " | Source form: " & sourceFormName & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, "Load Utility tab", savedErrNumber, savedErrSource, savedErrDescription, sourceFormName, "Form", "CodeDb", "GeneratedFormCommand"
    CoreShowErrorMessage "Could not load the selected ChatGPT Utility tab." & vbCrLf & vbCrLf & savedErrDescription, vbCritical, DEFAULT_UI_MESSAGE_TITLE, "UtilityLazyLoadSelectedTab", "UI"
End Function

Private Sub AddEmbeddedFormTabPage(ByVal hostFormName As String, ByVal tabCtl As control, ByVal pageIndex As Long, ByVal pageName As String, ByVal captionText As String, ByVal sourceFormName As String, ByVal descriptionText As String)
    Dim pg As Page
    Dim subCtl As Access.control
    Dim leftPos As Long
    Dim topPos As Long
    Dim widthValue As Long
    Dim heightValue As Long
    Set pg = ConfigureTabPage(tabCtl, pageIndex, pageName, captionText)
    leftPos = DEFAULT_UI_TAB_LEFT + DEFAULT_UI_FORM_LEFT_MARGIN
    topPos = DEFAULT_UI_TAB_TOP + 500
    widthValue = DEFAULT_UI_TAB_WIDTH - (2 * DEFAULT_UI_FORM_LEFT_MARGIN)
    heightValue = DEFAULT_UI_TAB_HEIGHT - 780
    With SharedFormDesign_AddPlainLabel(hostFormName, acDetail, pg.Name, descriptionText, leftPos, topPos, widthValue, 260, DEFAULT_UI_FONT_SIZE)
        .TextAlign = 1
    End With
    If SharedDao_AccessObjectExists(GetHostDb(), acForm, sourceFormName) Then
        Set subCtl = SharedFormDesign_AddSubformControl(hostFormName, acDetail, pg.Name, "sub_" & pageName, sourceFormName, leftPos, topPos + 260 + DEFAULT_UI_SECTION_TITLE_CONTENT_GAP, widthValue, heightValue)
        subCtl.Tag = sourceFormName
        subCtl.sourceObject = vbNullString
    Else
        With SharedFormDesign_AddPlainLabel(hostFormName, acDetail, pg.Name, "Source form not found: " & sourceFormName, leftPos, topPos + 260 + DEFAULT_UI_SECTION_TITLE_CONTENT_GAP, widthValue, 260, DEFAULT_UI_FONT_SIZE)
            .TextAlign = 1
        End With
        SharedFormDesign_AddButton hostFormName, acDetail, pg.Name, "Build / Open Source Form", "=ChatGPT_OpenFormByName(""" & sourceFormName & """)", leftPos, topPos + 260 + DEFAULT_UI_SECTION_TITLE_CONTENT_GAP + 260 + DEFAULT_UI_BUTTON_ROW_GAP, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Build missing generated forms if possible, then open this source form."
    End If
End Sub

Private Sub AddSkippedFormNote(ByRef skippedForms As Collection, ByVal formName As String, ByVal reasonText As String)
    If skippedForms Is Nothing Then Exit Sub
    skippedForms.Add formName & ": " & reasonText
End Sub

Private Sub AddVerifiedExistingFormNote(ByVal formName As String)
    If mChatGPTVerifiedExistingForms Is Nothing Then ResetChatGPTVerifiedExistingForms
    mChatGPTVerifiedExistingForms.Add formName
End Sub

Private Sub ApplyBoundFormProperties(ByVal frm As Form, ByVal captionText As String, ByVal recordSourceText As String, ByVal widthValue As Long, ByVal heightValue As Long, ByVal defaultViewValue As Long)
    On Error GoTo ErrHandler
    frm.Caption = captionText
    frm.recordSource = recordSourceText
    frm.Width = widthValue
    frm.defaultView = defaultViewValue
    frm.RecordSelectors = True
    frm.NavigationButtons = True
    frm.DividingLines = True
    frm.AutoCenter = True
    frm.ScrollBars = 2
    frm.Section(acDetail).Height = heightValue
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ApplyBoundFormProperties", "Could not apply required bound-form properties." & vbCrLf & "Caption: " & captionText & vbCrLf & "RecordSource: " & recordSourceText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Sub

Private Sub ApplySplitPopupFormProperties(ByVal frm As Form)
    On Error GoTo ErrHandler
    frm.Modal = False
    frm.AutoCenter = True
    frm.RecordSelectors = True
    frm.NavigationButtons = True
    frm.AllowFilters = True
    frm.FilterOnLoad = False
    frm.OrderByOnLoad = False
    CallByName frm, "SplitFormDatasheet", VbLet, 1
    CallByName frm, "SplitFormOrientation", VbLet, 1
    CallByName frm, "SplitFormSplitterBar", VbLet, True
    CallByName frm, "SplitFormSize", VbLet, DEFAULT_UI_SPLIT_POPUP_DATASHEET_HEIGHT
    CallByName frm, "SplitFormPrinting", VbLet, 0
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ApplySplitPopupFormProperties", "Could not apply required split/popup form properties." & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Sub

' Purpose: Applies unbound form properties for frm, caption text.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   frm (Input; Form): Access.Form instance whose controls or form properties the routine reads, configures, or validates.
'   captionText (Input; String): Caption text text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   widthValue (Input; Long): Requested Access control/layout width, in the units expected by the generated-form/report builder.
'   heightValue (Input; Long): Requested Access control/layout height, in the units expected by the generated-form/report builder.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub ApplyUnboundFormProperties(ByVal frm As Form, ByVal captionText As String, ByVal widthValue As Long, ByVal heightValue As Long)
    On Error GoTo ErrHandler
    frm.Caption = captionText
    frm.Width = widthValue
    frm.defaultView = 0
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.DividingLines = False
    frm.AutoCenter = True
    frm.ScrollBars = 2
    frm.Section(acDetail).Height = heightValue
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ApplyUnboundFormProperties", "Could not apply required unbound-form properties." & vbCrLf & "Caption: " & captionText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Sub

Private Function BoundsForMessage(ByVal leftPos As Long, ByVal topPos As Long, ByVal widthValue As Long, ByVal heightValue As Long) As String
    BoundsForMessage = "Left=" & CStr(leftPos) & "; Top=" & CStr(topPos) & "; Width=" & CStr(widthValue) & "; Height=" & CStr(heightValue) & "; Right=" & CStr(leftPos + widthValue) & "; Bottom=" & CStr(topPos + heightValue)
End Function

' ===============================================================================================================================================================================
' BUILD ORCHESTRATION / SPECIALIZED GENERATED-FORM BUILDERS
' ===============================================================================================================================================================================

' Purpose: Builds generated forms for replace existing, foundation mode.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   replaceExisting (Input Optional; Boolean): When True, enables or requests existing for this call; when False, the routine follows its normal non-forced/default path. Default: True.
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
'   foundationMode (Input Optional; String): Mode value selecting the foundation processing path for this invocation. Default: FOUNDATION_MODE_FULL.
'   buildScope (Input Optional; String): Scope value limiting which build items or operations are included in this call. Default: FORM_BUILD_SCOPE_ALL.
'   buildOptions (Input Optional; GeneratedFormBuildOptions): Option flags selecting the optional build behaviors enabled for this invocation. Default: GENERATED_FORM_BUILD_OPTIONS_STANDARD.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub BuildGeneratedForms(Optional ByVal replaceExisting As Boolean = True, Optional ByVal showMessage As Boolean = True, Optional ByVal foundationMode As String = FOUNDATION_MODE_FULL, Optional ByVal buildScope As String = FORM_BUILD_SCOPE_ALL, Optional ByVal buildOptions As GeneratedFormBuildOptions = GENERATED_FORM_BUILD_OPTIONS_STANDARD)
    Dim recoveryClearDetail As String
    Dim groupGeneratedForms As Boolean
    Dim requireNoSkippedForms As Boolean
    Dim registerGeneratedForms As Boolean
    Dim preserveSetupPreflightPlan As Boolean
    'Intent: Builds the complete generated Access Development Console form set under the UI-owned durable recovery boundary.
    On Error GoTo ErrHandler
    groupGeneratedForms = ((buildOptions And GENERATED_FORM_BUILD_OPTION_GROUP_NAVIGATION_PANE) <> 0)
    requireNoSkippedForms = ((buildOptions And GENERATED_FORM_BUILD_OPTION_REQUIRE_NO_SKIPPED_FORMS) <> 0)
    registerGeneratedForms = ((buildOptions And GENERATED_FORM_BUILD_OPTION_REGISTER_GENERATED_FORMS) <> 0)
    preserveSetupPreflightPlan = ((buildOptions And GENERATED_FORM_BUILD_OPTION_PRESERVE_SETUP_PREFLIGHT_PLAN) <> 0)
    RequireNoUnfinishedUIRecovery "BuildGeneratedForms"
    If Len(mSharedPreparedGeneratedFormName) > 0 Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "BuildGeneratedForms", "A shared generated-form replacement remains active and must be recovered before a new form-build run starts." & vbCrLf & "Prepared form: " & mSharedPreparedGeneratedFormName
    Dim builtForms As Collection
    Dim skippedForms As Collection
    Dim navNotes As Collection
    Dim registeredForms As Collection
    Dim navigationPaneForms As Collection
    Dim summaryText As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim detailMessage As String
    Dim normalizedFoundationMode As String
    Dim normalizedBuildScope As String
    Dim progressTotal As Long
    Dim traceRunOwnedByProcedure As Boolean
    Dim sameDatabaseBuildAllowed As Boolean
    Dim formBackupEntries As Collection
    Dim originalGeneratedForms As Collection
    Dim generatedFormChangesStarted As Boolean
    Dim generatedFormOperationCommitted As Boolean
    Dim rollbackDetail As String
    Dim rollbackErrNumber As Long
    Dim rollbackErrDescription As String
    Dim generatedFormRollbackAttempted As Boolean
    Dim generatedFormRollbackDetail As String
    Dim originalFailureText As String
    Dim formCatalogRollbackTable As String
    Dim physicalFormChangeNames As Collection
    Dim buildStartedTimer As Single
    Dim phaseStartedTimer As Single
    Dim foundationElapsedSeconds As Double
    Dim physicalBuildElapsedSeconds As Double
    Dim registrationAndValidationElapsedSeconds As Double
    Dim formTimingText As String
    Dim requiredBuildItem As Variant
    Dim physicalBuildRequired As Boolean
    Dim completeSetFastPath As Boolean
    Dim completeSetToken As String
    Dim sharedPreparedFormName As String
    Dim sharedAbortErrNumber As Long
    Dim sharedAbortErrDescription As String
    Dim sharedAbortDetail As String
    Dim physicalStateDetail As String
    Dim physicalStateToken As String
    Dim helpProofDetail As String
    Dim helpProofReused As Boolean
    Dim helpTopics As Collection
    Dim systemMacroDefinitions As Collection
    Dim systemMacroDefinitionIndex As Object
    Dim generatedFormBuildNames As Collection
    Dim generatedFormBuildItem As Variant
    Dim oneFormStartedTimer As Single
    Dim managedFormItem As Variant
    Dim descriptionChanged As Variant
    Dim standaloneDescriptionChangeCount As Long
    Dim standaloneDescriptionReconciliationPerformed As Boolean
    Dim oneFormCreateStart As Long
    Dim oneFormSaveStart As Long
    Dim oneFormFingerprintHitStart As Long
    Dim oneFormFingerprintMissStart As Long
    Dim oneFormExistsHitStart As Long
    Dim oneFormExistsMissStart As Long
    Dim discoveredUserFormCount As Variant
    Dim namingSnapshotOwned As Boolean
    Dim setupControlVisibilitySuspendedForPhysicalChange As Boolean
    buildStartedTimer = Timer
    normalizedFoundationMode = NormalizeChatGPTFoundationMode(foundationMode)
    normalizedBuildScope = NormalizeChatGPTBuildScope(buildScope)
    If StrComp(normalizedBuildScope, FORM_BUILD_SCOPE_CHATGPT_UI_FORMS, vbTextCompare) = 0 Then
        ClearGeneratedFormSameRunD2Proof
    ElseIf StrComp(normalizedBuildScope, FORM_BUILD_SCOPE_SUPPORT_FORMS, vbTextCompare) = 0 Then
        ClearGeneratedFormSameRunD3Proof
    ElseIf StrComp(normalizedBuildScope, FORM_BUILD_SCOPE_ALL, vbTextCompare) = 0 Then
        ClearGeneratedFormSameRunD2Proof
        ClearGeneratedFormSameRunD3Proof
    End If
    'Use the static canonical naming roster for the progress maximum so the physical build roster is not resolved through database-backed naming before foundation repair completes.
    progressTotal = GeneratedFormStaticBuildScopeCount(normalizedBuildScope) + 11
    sameDatabaseBuildAllowed = CodeDbIsHostDb()
    traceRunOwnedByProcedure = BeginUITraceRunIfNeeded("UI.FormGeneration")
    If sameDatabaseBuildAllowed = False Then
        If requireNoSkippedForms Then Err.Raise vbObjectError + 6714, "BuildGeneratedForms", "Strict generated-form rebuilding requires CodeDb and HostDb to be the same active database."
        detailMessage = GeneratedFormCodeDbOnlyMessage("Generating " & DEFAULT_UTILITY_SHORT_NAME & " generated forms")
        LogTraceMessage TRACE_LEVEL_INFO, "UI", "BuildGeneratedForms", "Generated-form operation skipped because CodeDb and HostDb are different.", detailMessage, vbNullString, 0, vbNullString, vbNullString, vbNullString, "FormSet", "CodeDb", "FormGeneration"
        If showMessage Then
            CoreShowMessage _
                detailMessage, _
                vbInformation, _
                DEFAULT_UI_MESSAGE_TITLE, _
                "ShowADCMessage", _
                "UI"
        End If
        CompleteUITraceRunIfOwned traceRunOwnedByProcedure, "Generated-form operation skipped."
        Exit Sub
    End If
    LogTraceMessage TRACE_LEVEL_INFO, "UI", "BuildGeneratedForms", "Generated-form operation started.", "Replace existing: " & CStr(replaceExisting) & vbCrLf & "Foundation mode: " & normalizedFoundationMode & vbCrLf & "Group generated forms in Navigation Pane: " & CStr(groupGeneratedForms) & vbCrLf & "Build scope: " & normalizedBuildScope, vbNullString, 0, vbNullString, vbNullString, vbNullString, "FormSet", "CodeDb", "FormGeneration"
    ' INTENTIONAL FILTERED PARTITION: builtForms, skippedForms is a consumer-required subset/projection derived from the source enumerated by BuildGeneratedForms; its membership is not independently authoritative.
    ' INTENTIONAL ORDERED VIEW: navNotes preserves the deterministic sequence required by BuildGeneratedForms; keyed-only or unordered storage would not preserve the consumer semantics.
    Set builtForms = New Collection
    Set skippedForms = New Collection
    Set navNotes = New Collection
    ResetChatGPTVerifiedExistingForms
    Set mChatGPTGeneratedFormSignatureCurrentCache = CreateObject("Scripting.Dictionary")
    mChatGPTGeneratedFormSignatureCurrentCache.CompareMode = vbTextCompare
    Set mChatGPTGeneratedFormSourceFingerprintCache = CreateObject("Scripting.Dictionary")
    mChatGPTGeneratedFormSourceFingerprintCache.CompareMode = vbTextCompare
    Set mChatGPTGeneratedFormLookupEditorPolicyFingerprintCache = CreateObject("Scripting.Dictionary")
    mChatGPTGeneratedFormLookupEditorPolicyFingerprintCache.CompareMode = vbTextCompare
    Set mChatGPTGeneratedFormObjectExistsCache = CreateObject("Scripting.Dictionary")
    mChatGPTGeneratedFormObjectExistsCache.CompareMode = vbTextCompare
    mChatGPTGeneratedFormSourceFingerprintCacheHits = 0
    mChatGPTGeneratedFormSourceFingerprintCacheMisses = 0
    mChatGPTGeneratedFormObjectExistsCacheHits = 0
    mChatGPTGeneratedFormObjectExistsCacheMisses = 0
    mChatGPTGeneratedFormDesignCreateCount = 0
    mChatGPTGeneratedFormDesignSaveCloseCount = 0
    ClearChatGPTBuildContext
    SetChatGPTBuildContext "Validating " & DEFAULT_UTILITY_LONG_NAME & " form-builder dependencies.", vbNullString, "Validate required main-module functions before changing generated forms.", vbNullString
    SetChatGPTBuildContext "Starting form generation.", vbNullString, "Initialize collections and progress meter.", vbNullString
    BeginChatGPTProgress DEFAULT_UI_BUILD_PROGRESS_TITLE, progressTotal
    SetChatGPTBuildContext "Preparing " & DEFAULT_UTILITY_LONG_NAME & " foundation.", vbNullString, "Run only the table/query/catalog setup procedures still required by this caller.", "Foundation mode: " & normalizedFoundationMode
    phaseStartedTimer = Timer
    If StrComp(normalizedFoundationMode, FOUNDATION_MODE_SKIP, vbTextCompare) = 0 Then
        UpdateChatGPTProgress "Skipping " & DEFAULT_UTILITY_SHORT_NAME & " foundation repair; caller already completed prerequisites."
        LogTraceMessage TRACE_LEVEL_STATUS, "UI", "BuildGeneratedForms", "Skipped generated-form foundation repair because caller already completed prerequisites.", "Foundation mode: " & normalizedFoundationMode, "Prepare form foundation", 0, vbNullString, vbNullString, DEFAULT_CHATGPT_UI_DISPLAY_PREFIX, "FormSet", "CodeDb", "FormGeneration"
    Else
        UpdateChatGPTProgress "Preparing " & DEFAULT_UTILITY_SHORT_NAME & " supporting objects. Foundation mode: " & normalizedFoundationMode & "."
        ChatGPT_EnsureFoundation normalizedFoundationMode
    End If
    foundationElapsedSeconds = ElapsedSecondsFromTimer(phaseStartedTimer)
    ' INTENTIONAL OPERATION SNAPSHOT: foundation repair may mutate naming inputs. Reuse Setup/caller naming state when already active; otherwise acquire one COMMON snapshot after foundation and retain it through
    ' planning, build, verification, and terminal reporting.
    If NamingSnapshotIsReady() = False Then
        BeginNamingSnapshot
        namingSnapshotOwned = True
    End If
    Set generatedFormBuildNames = GeneratedFormBuildNamesForScope(normalizedBuildScope)
    If generatedFormBuildNames.count <> GeneratedFormAuthoritativeCount(normalizedBuildScope) Or generatedFormBuildNames.count + 11 <> progressTotal Then Err.Raise vbObjectError + 7865, "BuildGeneratedForms", "The selected generated-form naming roster and physical build roster do not contain the same number of forms." & vbCrLf & "Build scope: " & normalizedBuildScope & vbCrLf & "Canonical naming forms: " & CStr(progressTotal - 11) & vbCrLf & "Authoritative physical forms: " & CStr(GeneratedFormAuthoritativeCount(normalizedBuildScope)) & vbCrLf & "Physical build forms: " & CStr(generatedFormBuildNames.count)
    If StrComp(normalizedBuildScope, FORM_BUILD_SCOPE_ALL, vbTextCompare) = 0 Then
        helpProofReused = SetupGeneratedFormHelpProofCurrent(normalizedBuildScope, helpProofDetail)
        If helpProofReused Then
            LogTraceMessage TRACE_LEVEL_INFO, "UI", "BuildGeneratedForms", "Reused the generated Help-definition proof from Setup D1.", helpProofDetail, "Validate generated Help definitions", 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Help"), "Form", "CodeDb", "Help"
        Else
            Set systemMacroDefinitions = x__UTIL_ADC_DEFAULTS.GetSystemMacroDefinitions()
            Set systemMacroDefinitionIndex = ChatGPTSystemMacroDefinitionIndex(systemMacroDefinitions)
            Set helpTopics = ChatGPTHelpTopics(systemMacroDefinitions, systemMacroDefinitionIndex)
            ValidateChatGPTHelpDefinitions helpTopics, systemMacroDefinitions, systemMacroDefinitionIndex
        End If
    End If
    If mSetupGeneratedFormPreflightReady And _
       mSetupGeneratedFormPreflightReplaceExisting = replaceExisting And _
       Len(GetActiveOrLastTraceRunID()) > 0 And StrComp(mSetupGeneratedFormPreflightTraceRunID, GetActiveOrLastTraceRunID(), vbBinaryCompare) = 0 And _
       (mSetupGeneratedFormPreflightScope = GeneratedFormPreflightScopeStateFromText(normalizedBuildScope) Or mSetupGeneratedFormPreflightScope = GENERATED_FORM_PREFLIGHT_SCOPE_ALL) Then
        If mSetupGeneratedFormPreflightScope = GeneratedFormPreflightScopeStateFromText(normalizedBuildScope) Then
            Set physicalFormChangeNames = mSetupGeneratedFormPreflightPhysicalNames
            completeSetFastPath = mSetupGeneratedFormCompleteSetFastPath
            completeSetToken = mSetupGeneratedFormCompleteSetToken
        Else
            Set physicalFormChangeNames = FilterGeneratedFormNamesForScope(mSetupGeneratedFormPreflightPhysicalNames, normalizedBuildScope)
            completeSetFastPath = (mSetupGeneratedFormCompleteSetFastPath And physicalFormChangeNames.count = 0)
            completeSetToken = GeneratedFormCompleteSetToken(normalizedBuildScope)
        End If
        If preserveSetupPreflightPlan = False Then ClearSetupGeneratedFormPreflightProofState
        UpdateChatGPTProgress "Using the requested partition of the authoritative generated-form change plan from Setup D1."
    Else
        If preserveSetupPreflightPlan Then Err.Raise vbObjectError + 6766, "BuildGeneratedForms", "Setup D2/D3 refused to reconstruct a generated-form plan independently of D1."
        ClearSetupGeneratedFormPreflightProofState
        SetChatGPTBuildContext "Validating unchanged authoritative generated forms.", vbNullString, "Fail before any physical form changes when a form excluded from the selected build scope is missing.", "Build scope: " & normalizedBuildScope
        SetChatGPTBuildContext "Preflighting complete generated-form build plan.", vbNullString, "Validate every planned bound form source and form-name identity before rollback preparation or physical form changes.", "Build scope: " & normalizedBuildScope
        UpdateChatGPTProgress "Preflighting every generated-form source before the first physical form change."
        PreflightGeneratedFormBuildPlan normalizedBuildScope
        Set physicalFormChangeNames = ChatGPTPhysicalFormChangeNamesRequiringBuild(replaceExisting, normalizedBuildScope)
    End If
    Set mChatGPTGeneratedFormsRequiringBuild = CreateObject("Scripting.Dictionary")
    mChatGPTGeneratedFormsRequiringBuild.CompareMode = vbTextCompare
    For Each requiredBuildItem In physicalFormChangeNames
        SetupControlPump
        mChatGPTGeneratedFormsRequiringBuild(CStr(requiredBuildItem)) = True
    Next requiredBuildItem
    If mSetupControlActive Then
        If mSetupControlScanMode Then
            setupControlVisibilitySuspendedForPhysicalChange = mChatGPTGeneratedFormsRequiringBuild.Exists(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SCAN_PROGRESS))
        Else
            setupControlVisibilitySuspendedForPhysicalChange = mChatGPTGeneratedFormsRequiringBuild.Exists(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_RUNNING)) Or mChatGPTGeneratedFormsRequiringBuild.Exists(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CANON_GENERATED_FORM_LOGICAL_ID_SETUP_CONTROL_PAUSED))
        End If
        If setupControlVisibilitySuspendedForPhysicalChange Then
            mSetupControlVisibilitySuspended = True
            LogTraceMessage TRACE_LEVEL_INFO, "SetupControl", "BuildGeneratedForms", "Automatic Setup Control visibility was suspended for the D2 physical-change boundary.", "The active controller family contains a form in the authoritative D1/D2 physical-change plan; event pumping remains active while automatic reopen and Pause transition are deferred until commit or rollback completes.", "D2 Setup Control isolation", 0, vbNullString, vbNullString, vbNullString, "SetupControl", "CodeDb", "FormGeneration"
        End If
    End If
    physicalBuildRequired = (physicalFormChangeNames.count > 0)
    If physicalBuildRequired Then
        If CodeDbIsHostDb() And TraceRunActive() Then InvalidatePhysicalManagedFormInventoryForSetup
        mChatGPTGeneratedFormsUsePopups = SettingBool(SETTING_GENERATED_FORMS_USE_POPUPS, DEFAULT_GENERATED_FORMS_USE_POPUPS)
    End If
    If physicalFormChangeNames.count > 0 Then
            ReportBlockingPromptWait "BuildGeneratedForms", "saving and closing generated forms selected for physical rebuild"
            If SharedAccess_PrepareOpenObjectsForDesignChange(acForm, physicalFormChangeNames, "building or updating the selected generated " & DEFAULT_UTILITY_LONG_NAME & " forms") = False Then
                ReportBlockingPromptResult "BuildGeneratedForms", "saving and closing generated forms selected for physical rebuild", False
                MarkSystemCommandCancelled
                LogTraceMessage TRACE_LEVEL_INFO, "UI", "BuildGeneratedForms", "Generated-form rebuild cancelled before physical form changes because affected open forms were not approved for save and close.", "Build scope: " & normalizedBuildScope, vbNullString, 0, vbNullString, vbNullString, DEFAULT_CHATGPT_UI_DISPLAY_PREFIX, "FormSet", "CodeDb", "FormGeneration"
                Err.Raise vbObjectError + 6696, "BuildGeneratedForms", "Generated-form rebuild was cancelled because affected open forms were not approved for save and close."
            End If
            ReportBlockingPromptResult "BuildGeneratedForms", "saving and closing generated forms selected for physical rebuild", True
            SetChatGPTBuildContext "Preparing generated-form rollback state.", vbNullString, "Inventory only the physical forms the selected build scope can change and save restorable definitions before replacement.", "Replace existing: " & CStr(replaceExisting)
            PrintUIStatus "BuildGeneratedForms", "Preparing scoped generated-form rollback state before the first form change.", DEFAULT_CHATGPT_UI_DISPLAY_PREFIX, "FormGeneration"
            PrepareChatGPTGeneratedFormRollbackState replaceExisting, physicalFormChangeNames, formBackupEntries, originalGeneratedForms, formCatalogRollbackTable, registerGeneratedForms
            generatedFormChangesStarted = True
    Else
        ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: formBackupEntries holds planned/recovery work or independently observed validation evidence for BuildGeneratedForms; keep it distinct from the source authority.
        ' INTENTIONAL FILTERED PARTITION: originalGeneratedForms is a consumer-required subset/projection derived from the source enumerated by BuildGeneratedForms; its membership is not independently authoritative.
        Set formBackupEntries = New Collection
        Set originalGeneratedForms = New Collection
        formCatalogRollbackTable = vbNullString
        PrintUIStatus "BuildGeneratedForms", "All requested generated forms have current deterministic signatures; rollback preparation and physical regeneration were not required.", DEFAULT_CHATGPT_UI_DISPLAY_PREFIX, "FormGeneration"
    End If
    phaseStartedTimer = Timer
    If physicalBuildRequired = False Then
        UpdateChatGPTProgress "Skipping the ADC-owned physical generated-form creation loop because all ADC-owned deterministic signatures and physical forms are current."
        LogTraceMessage TRACE_LEVEL_STATUS, "UI", "BuildGeneratedForms", "Skipped the ADC-owned generated-form creation-call loop because complete preflight found no ADC-owned physical changes.", "ADC-owned forms requiring physical build: 0" & vbCrLf & "Build scope: " & normalizedBuildScope, "Skip physical form build loop", 0, vbNullString, vbNullString, DEFAULT_CHATGPT_UI_DISPLAY_PREFIX, "FormSet", "CodeDb", "FormGeneration"
        GoTo PhysicalBuildComplete
    End If
    For Each generatedFormBuildItem In generatedFormBuildNames
        SetupControlPump
        oneFormStartedTimer = Timer
        oneFormCreateStart = mChatGPTGeneratedFormDesignCreateCount
        oneFormSaveStart = mChatGPTGeneratedFormDesignSaveCloseCount
        oneFormFingerprintHitStart = mChatGPTGeneratedFormSourceFingerprintCacheHits
        oneFormFingerprintMissStart = mChatGPTGeneratedFormSourceFingerprintCacheMisses
        oneFormExistsHitStart = mChatGPTGeneratedFormObjectExistsCacheHits
        oneFormExistsMissStart = mChatGPTGeneratedFormObjectExistsCacheMisses
        BuildOneCanonicalGeneratedForm CStr(generatedFormBuildItem), replaceExisting, builtForms, skippedForms
        formTimingText = "Form: " & CStr(generatedFormBuildItem) & vbCrLf & _
            "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(oneFormStartedTimer)) & vbCrLf & _
            "Design creates: " & CStr(mChatGPTGeneratedFormDesignCreateCount - oneFormCreateStart) & vbCrLf & _
            "Design save/closes: " & CStr(mChatGPTGeneratedFormDesignSaveCloseCount - oneFormSaveStart) & vbCrLf & _
            "Source fingerprint cache hits/misses: " & CStr(mChatGPTGeneratedFormSourceFingerprintCacheHits - oneFormFingerprintHitStart) & "/" & CStr(mChatGPTGeneratedFormSourceFingerprintCacheMisses - oneFormFingerprintMissStart) & vbCrLf & _
            "Source existence cache hits/misses: " & CStr(mChatGPTGeneratedFormObjectExistsCacheHits - oneFormExistsHitStart) & "/" & CStr(mChatGPTGeneratedFormObjectExistsCacheMisses - oneFormExistsMissStart)
        LogTraceMessage TRACE_LEVEL_STATUS, "FormGenerationTiming", "BuildGeneratedForms", "Generated-form build call completed.", formTimingText, "Build one generated form", 0, vbNullString, vbNullString, CStr(generatedFormBuildItem), "Form", "CodeDb", "FormGenerationTiming"
    Next generatedFormBuildItem
PhysicalBuildComplete:
    If requireNoSkippedForms And skippedForms.count > 0 Then Err.Raise vbObjectError + 6715, "BuildGeneratedForms", "Strict generated-form rebuilding stopped because one or more required forms were skipped." & vbCrLf & SharedCollection_Join(skippedForms, vbCrLf)
    Set registeredForms = ChatGPTFormsForRegistration(normalizedBuildScope)
    physicalBuildElapsedSeconds = ElapsedSecondsFromTimer(phaseStartedTimer)
    phaseStartedTimer = Timer
    If registerGeneratedForms Then
        SetChatGPTBuildContext "Registering generated forms.", vbNullString, "Upsert generated form rows in UI_Frm_Registry for the selected authoritative partition.", "Built form count: " & CStr(builtForms.count) & "; Verified existing form count: " & CStr(ChatGPTVerifiedExistingFormsCount())
        UpdateChatGPTProgress "Registering and verifying every authoritative generated form in UI_Frm_Registry."
        RegisterGeneratedFormsInFormsCatalog registeredForms, navNotes
    Else
        SetChatGPTBuildContext "Deferring numbered Setup generated-form registry reconciliation to D4.", vbNullString, "Direct/runtime callers may request partition registration; during numbered Setup D2/D3 own physical form construction only and D4 exclusively owns complete UI_Frm_Registry reconciliation.", "Authoritative physical form count: " & CStr(registeredForms.count)
        UpdateChatGPTProgress "Physical generated forms verified; numbered Setup defers complete UI_Frm_Registry reconciliation to D4."
        navNotes.Add "D2/D3 own physical generated-form construction only during numbered Setup; D4 owns complete managed-form registry reconciliation."
    End If
    RefreshGeneratedFormBuildSignatures builtForms
    If completeSetFastPath Then
        If registerGeneratedForms Then
            navNotes.Add "Generated-form complete-set token matched; physical rebuilding was skipped and the direct/runtime caller still performed its requested UI_Frm_Registry registration/verification pass."
            LogTraceMessage TRACE_LEVEL_INFO, "UI", "BuildGeneratedForms", "Generated-form complete-set fast path used for physical forms.", "Forms registered/verified: " & CStr(registeredForms.count) & vbCrLf & "Complete-set token: " & completeSetToken, "Register unchanged generated forms", 0, vbNullString, vbNullString, DEFAULT_CHATGPT_UI_DISPLAY_PREFIX, "FormSet", "CodeDb", "FormGeneration"
        Else
            navNotes.Add "Generated-form partition token matched; physical rebuilding was skipped and the current partition metadata remains subject to D4 complete reconciliation."
            LogTraceMessage TRACE_LEVEL_INFO, "UI", "BuildGeneratedForms", "Setup D2/D3 generated-form partition complete-set fast path used.", "Physical forms verified: " & CStr(registeredForms.count) & vbCrLf & "Complete-set token: " & completeSetToken & vbCrLf & "UI_Frm_Registry metadata: deferred to D4 complete reconciliation", "Verify unchanged generated forms", 0, vbNullString, vbNullString, DEFAULT_CHATGPT_UI_DISPLAY_PREFIX, "FormSet", "CodeDb", "FormGeneration"
        End If
    End If
    SetChatGPTBuildContext "Delegating generated-form Navigation Pane placement.", vbNullString, "For this direct/runtime grouping request, Review owns physical generated-form placement; Setup D2/D3 always defer placement to I3.", "Authoritative form count: " & CStr(registeredForms.count) & "; GroupGeneratedForms: " & CStr(groupGeneratedForms)
    If groupGeneratedForms Then
        UpdateChatGPTProgress "Applying and verifying all active generated-form Navigation Pane placement through Review."
        Set navigationPaneForms = ActiveGeneratedFormsForNavigationPaneVerification(navNotes, registeredForms)
        ApplyGeneratedFormsNavigationPane navigationPaneForms, navNotes
        SetChatGPTBuildContext "Refreshing Navigation Pane.", vbNullString, "Refresh database window after Review-owned generated-form placement.", vbNullString
        UpdateChatGPTProgress "Refreshing the Navigation Pane."
        RefreshNavigationPaneAfterGrouping
    Else
        UpdateChatGPTProgress "Skipping generated-form Navigation Pane grouping; Setup defers all physical placement to I3."
        navNotes.Add "Generated-form Navigation Pane grouping was skipped because I3 owns the authoritative complete-domain physical placement pass."
        LogTraceMessage TRACE_LEVEL_STATUS, "UI", "BuildGeneratedForms", "Generated-form Navigation Pane grouping skipped for Setup D2/D3.", "GroupGeneratedForms: " & CStr(groupGeneratedForms), "Defer form grouping to I3", 0, vbNullString, vbNullString, DEFAULT_CHATGPT_UI_DISPLAY_PREFIX, "FormSet", "CodeDb", "FormGeneration"
    End If
    If SetupDomainMutationWritesAllowedForCurrentPart(SETUP_DOMAIN_G1_DESCRIPTIONS) Then
        standaloneDescriptionReconciliationPerformed = True
        SetChatGPTBuildContext "Reconciling standalone generated-form descriptions.", vbNullString, "Apply and verify canonical managed-object Descriptions only for the selected generated-form partition.", "Authoritative form count: " & CStr(registeredForms.count)
        For Each managedFormItem In registeredForms
            descriptionChanged = False
            If ApplyAndVerifyManagedObjectDescriptionForStandaloneRepair(CStr(managedFormItem), ADCAT_OBJECT_TYPE_FORM, descriptionChanged) = False Then Err.Raise vbObjectError + 10123, "BuildGeneratedForms", "Managed generated-form Description reconciliation returned False." & vbCrLf & "Form: " & CStr(managedFormItem)
            If CBool(descriptionChanged) Then standaloneDescriptionChangeCount = standaloneDescriptionChangeCount + 1
        Next managedFormItem
        navNotes.Add "Standalone managed-form descriptions reconciled for the selected build partition; changed: " & CStr(standaloneDescriptionChangeCount) & "."
    End If
    completeSetToken = GeneratedFormCompleteSetToken(normalizedBuildScope)
    Call EnsureDaoPropertyValue(GetHostDb(), GeneratedFormCompleteSetTokenPropertyName(normalizedBuildScope), dbText, completeSetToken, "BuildGeneratedForms")
    physicalStateToken = GeneratedFormPhysicalStateToken(normalizedBuildScope, physicalStateDetail, discoveredUserFormCount)
    If Len(physicalStateToken) > 0 Then Call EnsureDaoPropertyValue(GetHostDb(), GeneratedFormPhysicalStateTokenPropertyName(normalizedBuildScope), dbText, physicalStateToken, "BuildGeneratedForms")
    SetChatGPTBuildContext "Strictly validating generated forms.", DEFAULT_CHATGPT_UI_DISPLAY_PREFIX, "Validate every canonical generated form through the shared contract framework.", "Authoritative forms: " & CStr(registeredForms.count)
    UpdateChatGPTProgress "Strictly validating every canonical generated form."
    If StrComp(normalizedBuildScope, FORM_BUILD_SCOPE_ALL, vbTextCompare) = 0 Then
        If systemMacroDefinitions Is Nothing Then Set systemMacroDefinitions = x__UTIL_ADC_DEFAULTS.GetSystemMacroDefinitions()
        If systemMacroDefinitionIndex Is Nothing Then Set systemMacroDefinitionIndex = ChatGPTSystemMacroDefinitionIndex(systemMacroDefinitions)
        ValidateGeneratedFormSet helpTopics, systemMacroDefinitions, systemMacroDefinitionIndex, normalizedBuildScope
    Else
        ValidateGeneratedFormSet buildScope:=normalizedBuildScope
    End If
    PublishGeneratedFormPartitionProof normalizedBuildScope, completeSetToken, physicalStateToken, discoveredUserFormCount
    PrintUIStatus "BuildGeneratedForms", "Final verification complete. Forms checked: " & CStr(registeredForms.count) & "; rebuilt: " & CStr(builtForms.count) & "; unchanged: " & CStr(ChatGPTVerifiedExistingFormsCount()) & ".", DEFAULT_CHATGPT_UI_DISPLAY_PREFIX, "FormGeneration"
    registrationAndValidationElapsedSeconds = ElapsedSecondsFromTimer(phaseStartedTimer)
    formTimingText = "Generated-form build timing:" & vbCrLf & _
        " - Foundation preparation: " & FormattedElapsedSeconds(foundationElapsedSeconds) & vbCrLf & _
        " - Physical form build/verification: " & FormattedElapsedSeconds(physicalBuildElapsedSeconds) & vbCrLf & _
        " - " & IIf(registerGeneratedForms, "Registry integration and validation", "Physical proof and validation") & ": " & FormattedElapsedSeconds(registrationAndValidationElapsedSeconds) & vbCrLf & _
        " - Total: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(buildStartedTimer)) & vbCrLf & _
        " - Design creates/save-closes: " & CStr(mChatGPTGeneratedFormDesignCreateCount) & "/" & CStr(mChatGPTGeneratedFormDesignSaveCloseCount) & vbCrLf & _
        " - Source fingerprint cache hits/misses: " & CStr(mChatGPTGeneratedFormSourceFingerprintCacheHits) & "/" & CStr(mChatGPTGeneratedFormSourceFingerprintCacheMisses) & vbCrLf & _
        " - Source existence cache hits/misses: " & CStr(mChatGPTGeneratedFormObjectExistsCacheHits) & "/" & CStr(mChatGPTGeneratedFormObjectExistsCacheMisses)
    LogTraceMessage TRACE_LEVEL_INFO, "FormGenerationTiming", "BuildGeneratedForms", "Generated-form phase timing.", formTimingText, "Complete generated-form build", 0, vbNullString, vbNullString, DEFAULT_CHATGPT_UI_DISPLAY_PREFIX, "FormSet", "CodeDb", "FormGenerationTiming"
    summaryText = "Generated or refreshed " & CStr(builtForms.count) & " ADC generated forms." & vbCrLf & vbCrLf & formTimingText
    summaryText = summaryText & vbCrLf & "Verified existing generated forms left unchanged: " & CStr(ChatGPTVerifiedExistingFormsCount())
    summaryText = summaryText & vbCrLf & "Form builder version: " & FORM_BUILDER_VERSION
    summaryText = summaryText & vbCrLf & "Foundation mode: " & normalizedFoundationMode
    If standaloneDescriptionReconciliationPerformed Then summaryText = summaryText & vbCrLf & "Standalone managed-form descriptions changed: " & CStr(standaloneDescriptionChangeCount)
    summaryText = summaryText & vbCrLf & "Build scope: " & normalizedBuildScope
    summaryText = summaryText & vbCrLf & "Generated-form Navigation Pane grouping requested: " & CStr(groupGeneratedForms)
    summaryText = summaryText & vbCrLf & "Strict no-skipped-forms mode: " & CStr(requireNoSkippedForms)
    summaryText = summaryText & vbCrLf & "UI_Frm_Registry integration in this build: " & IIf(registerGeneratedForms, "enabled for direct/runtime caller", "not requested")
    If builtForms.count > 0 Then summaryText = summaryText & vbCrLf & vbCrLf & "Generated/refreshed forms:" & vbCrLf & SharedCollection_Join(builtForms, vbCrLf)
    If ChatGPTVerifiedExistingFormsCount() > 0 Then summaryText = summaryText & vbCrLf & vbCrLf & "Verified existing generated forms:" & vbCrLf & SharedCollection_Join(mChatGPTVerifiedExistingForms, vbCrLf)
    If skippedForms.count > 0 Then summaryText = summaryText & vbCrLf & vbCrLf & "Skipped forms / setup notes:" & vbCrLf & SharedCollection_Join(skippedForms, vbCrLf)
    If navNotes.count > 0 Then summaryText = summaryText & vbCrLf & vbCrLf & "Navigation Pane grouping notes:" & vbCrLf & SharedCollection_Join(navNotes, vbCrLf)
    If Len(mSharedPreparedGeneratedFormName) > 0 Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "BuildGeneratedForms", "Generated-form verification reached the commit boundary while shared replacement recovery still owns a prepared form." & vbCrLf & "Prepared form: " & mSharedPreparedGeneratedFormName
    generatedFormOperationCommitted = True
    If setupControlVisibilitySuspendedForPhysicalChange Then
        mSetupControlVisibilitySuspended = False
        setupControlVisibilitySuspendedForPhysicalChange = False
        SetupControlUpdateStatus mSetupControlStatusText, True, True, False
        LogTraceMessage TRACE_LEVEL_INFO, "SetupControl", "BuildGeneratedForms", "Automatic Setup Control visibility was restored after the D2 physical-change boundary committed.", "The appropriate current Setup Control surface may now be opened normally.", "D2 Setup Control isolation complete", 0, vbNullString, vbNullString, vbNullString, "SetupControl", "CodeDb", "FormGeneration"
    End If
    If generatedFormChangesStarted Then
        If ClearUIRecoveryFailure("Generated-form restoration", recoveryClearDetail) = False Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "BuildGeneratedForms", recoveryClearDetail
        On Error Resume Next
        Err.Clear
        CleanupChatGPTGeneratedFormRollbackFiles formBackupEntries, formCatalogRollbackTable, True
        If Err.Number <> 0 Then
            summaryText = summaryText & vbCrLf & vbCrLf & "Rollback-resource cleanup warning. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
            LogTraceMessage TRACE_LEVEL_WARNING, "UI", "BuildGeneratedForms", "Generated-form build completed, but rollback-resource cleanup reported a warning.", summaryText, "Cleanup generated-form rollback resources", 0, vbNullString, vbNullString, DEFAULT_CHATGPT_UI_DISPLAY_PREFIX, "FormSet", "CodeDb", "FormGeneration"
        End If
        Err.Clear
        On Error GoTo ErrHandler
    End If
    LogTraceMessage TRACE_LEVEL_INFO, "UI", "BuildGeneratedForms", "ChatGPT " & DEFAULT_UTILITY_LONG_NAME & " generated-form operation completed.", summaryText, vbNullString, 0, vbNullString, vbNullString, vbNullString, "FormSet", "CodeDb", "FormGeneration"
    On Error Resume Next
    EndChatGPTProgress
    ClearChatGPTBuildContext
    CompleteUITraceRunIfOwned traceRunOwnedByProcedure, "ChatGPT " & DEFAULT_UTILITY_LONG_NAME & " generated-form operation complete."
    On Error GoTo 0
    ReportCoreActionSuccess "BuildGeneratedForms", "UI", DEFAULT_UTILITY_SHORT_NAME & " generated forms verified/refreshed successfully.", summaryText, "Build ADC generated forms", DEFAULT_CHATGPT_UI_DISPLAY_PREFIX, "FormSet", "CodeDb", "UI", showMessage, DEFAULT_UI_MESSAGE_TITLE, IIf(skippedForms.count > 0, vbExclamation, vbInformation), False, "UI", True, visibleStatusMessage:="Generated forms verified/refreshed successfully."
    If namingSnapshotOwned Then InvalidateNamingSnapshot
    namingSnapshotOwned = False
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If generatedFormOperationCommitted Then
        savedErrDescription = "Generated-form build committed successfully, but rollback-resource cleanup or recovery-state clearing failed." & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
        savedErrSource = "BuildGeneratedForms.PostCommitCleanup"
    End If
    If generatedFormOperationCommitted = False And Len(mSharedPreparedGeneratedFormName) > 0 Then
        sharedPreparedFormName = mSharedPreparedGeneratedFormName
        On Error Resume Next
        Err.Clear
        SharedFormDesign_AbortGeneratedFormReplacement sharedPreparedFormName
        sharedAbortErrNumber = Err.Number
        sharedAbortErrDescription = Err.Description
        On Error GoTo 0
        If sharedAbortErrNumber = 0 Then
            mSharedPreparedGeneratedFormName = vbNullString
            sharedAbortDetail = "Shared generated-form replacement was aborted before " & DEFAULT_UTILITY_SHORT_NAME & " outer rollback. Form: " & sharedPreparedFormName
        Else
            sharedAbortDetail = "Shared generated-form replacement abort failed. Form: " & sharedPreparedFormName & vbCrLf & "Err.Number: " & CStr(sharedAbortErrNumber) & vbCrLf & "Err.Description: " & sharedAbortErrDescription
        End If
    End If
    If sharedAbortErrNumber <> 0 Then
        If rollbackErrNumber = 0 Then rollbackErrNumber = sharedAbortErrNumber
        If Len(rollbackErrDescription) > 0 Then rollbackErrDescription = rollbackErrDescription & vbCrLf & vbCrLf
        rollbackErrDescription = rollbackErrDescription & sharedAbortDetail
        If Len(rollbackDetail) > 0 Then rollbackDetail = rollbackDetail & vbCrLf & vbCrLf
        rollbackDetail = rollbackDetail & sharedAbortDetail
    ElseIf Len(sharedAbortDetail) > 0 Then
        If Len(rollbackDetail) > 0 Then rollbackDetail = rollbackDetail & vbCrLf
        rollbackDetail = rollbackDetail & sharedAbortDetail
    End If
    'Keep ownership of the prepared durable recovery record during inline rollback. Success clears it; failure transfers it to a blocked recovery record.
    If SystemCommandWasCancelled() Or IsCancellationError(savedErrNumber) Then
        If generatedFormChangesStarted And generatedFormOperationCommitted = False And sharedAbortErrNumber = 0 Then
            generatedFormRollbackAttempted = True
            generatedFormRollbackDetail = TryRestoreChatGPTGeneratedFormsAfterFailure(physicalFormChangeNames, formBackupEntries, originalGeneratedForms, formCatalogRollbackTable, rollbackErrNumber, rollbackErrDescription)
            If Len(Trim$(generatedFormRollbackDetail)) > 0 Then rollbackDetail = SharedText_AppendNonBlankLine(rollbackDetail, generatedFormRollbackDetail)
        End If
        If rollbackErrNumber = 0 Then
            If setupControlVisibilitySuspendedForPhysicalChange Then
                mSetupControlVisibilitySuspended = False
                setupControlVisibilitySuspendedForPhysicalChange = False
                SetupControlUpdateStatus mSetupControlStatusText, True, True, False
            End If
            EndChatGPTProgress
            ClearChatGPTBuildContext
            LogTraceMessage TRACE_LEVEL_INFO, "UI", "BuildGeneratedForms", "Generated-form rebuild cancelled before continuing with the affected design changes.", "Build scope: " & IIf(Len(normalizedBuildScope) = 0, buildScope, normalizedBuildScope), vbNullString, 0, vbNullString, vbNullString, DEFAULT_CHATGPT_UI_DISPLAY_PREFIX, "FormSet", "CodeDb", "FormGeneration"
            CompleteUITraceRunIfOwned traceRunOwnedByProcedure, "Generated-form rebuild cancelled.", False
            If showMessage Then
                CoreShowMessage _
                    "Generated-form rebuild was cancelled before continuing with the affected design changes.", _
                    vbInformation, _
                    DEFAULT_UI_MESSAGE_TITLE, _
                    "ShowADCMessage", _
                    "UI"
                On Error Resume Next
                If namingSnapshotOwned Then InvalidateNamingSnapshot
                namingSnapshotOwned = False
                On Error GoTo 0
                Exit Sub
            End If
            On Error Resume Next
            If namingSnapshotOwned Then InvalidateNamingSnapshot
            namingSnapshotOwned = False
            On Error GoTo 0
            Err.Raise savedErrNumber, "BuildGeneratedForms", savedErrDescription
        End If
    End If
    If generatedFormChangesStarted And generatedFormOperationCommitted = False And generatedFormRollbackAttempted = False And sharedAbortErrNumber = 0 Then
        generatedFormRollbackDetail = TryRestoreChatGPTGeneratedFormsAfterFailure(physicalFormChangeNames, formBackupEntries, originalGeneratedForms, formCatalogRollbackTable, rollbackErrNumber, rollbackErrDescription)
        If Len(Trim$(generatedFormRollbackDetail)) > 0 Then rollbackDetail = SharedText_AppendNonBlankLine(rollbackDetail, generatedFormRollbackDetail)
    End If
    If rollbackErrNumber = 0 Then
        If setupControlVisibilitySuspendedForPhysicalChange Then
            mSetupControlVisibilitySuspended = False
            setupControlVisibilitySuspendedForPhysicalChange = False
            SetupControlUpdateStatus mSetupControlStatusText, True, True, False
        End If
    ElseIf setupControlVisibilitySuspendedForPhysicalChange Then
        LogTraceMessage TRACE_LEVEL_WARNING, "SetupControl", "BuildGeneratedForms", "Automatic Setup Control visibility remains suspended because generated-form rollback did not complete.", "The affected Setup Control surface remains isolated while the blocked recovery error propagates. The owning Setup failure lifecycle will tear down Setup Control without reopening the uncertain form.", "D2 Setup Control isolation retained", rollbackErrNumber, "BuildGeneratedForms.Rollback", rollbackErrDescription, vbNullString, "SetupControl", "CodeDb", "Recovery"
    End If
    If rollbackErrNumber <> 0 Then
        originalFailureText = "Original generated-form operation error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription
        savedErrNumber = vbObjectError + 6726
        savedErrSource = "BuildGeneratedForms.Rollback"
        savedErrDescription = "Generated-form rollback did not complete and requires recovery." & vbCrLf & rollbackErrDescription & vbCrLf & vbCrLf & originalFailureText
    End If
    FailChatGPTProgress "ChatGPT " & DEFAULT_UTILITY_SHORT_NAME & " generated-form operation failed. Review the Immediate Window trace."
    detailMessage = BuildChatGPTErrorMessage("Could not generate the requested ADC generated-form scope.", savedErrNumber, savedErrSource, savedErrDescription)
    detailMessage = detailMessage & vbCrLf & "Foundation mode: " & IIf(Len(normalizedFoundationMode) = 0, foundationMode, normalizedFoundationMode) & vbCrLf & "Build scope: " & IIf(Len(normalizedBuildScope) = 0, buildScope, normalizedBuildScope) & vbCrLf & "Group generated forms in Navigation Pane: " & CStr(groupGeneratedForms)
    If Len(Trim$(rollbackDetail)) > 0 Then detailMessage = detailMessage & vbCrLf & vbCrLf & "Generated-form rollback:" & vbCrLf & rollbackDetail
    DebugPrintChatGPTBuildError detailMessage, savedErrNumber, savedErrSource, savedErrDescription
    LogTraceMessage TRACE_LEVEL_ERROR, "UI", "BuildGeneratedForms", "ChatGPT " & DEFAULT_UTILITY_LONG_NAME & " generated-form operation failed.", detailMessage, mChatGPTBuildStep, savedErrNumber, savedErrSource, savedErrDescription, mChatGPTBuildForm, "Form", "CodeDb", "FormGeneration"
    ClearChatGPTBuildContext
    If showMessage Then
        CoreShowErrorMessage _
            detailMessage, _
            vbCritical, _
            DEFAULT_UI_MESSAGE_TITLE, _
            "BuildGeneratedForms", _
            "UI"
    End If
    CompleteUITraceRunIfOwned traceRunOwnedByProcedure, "ChatGPT " & DEFAULT_UTILITY_LONG_NAME & " generated-form operation failed.", False
    On Error Resume Next
    If namingSnapshotOwned Then InvalidateNamingSnapshot
    namingSnapshotOwned = False
    On Error GoTo 0
    If showMessage Then Exit Sub
    Err.Raise savedErrNumber, "BuildGeneratedForms", detailMessage
End Sub

Private Function ChatGPTBuildFormStateForMessage(ByVal formName As String) As String
    On Error Resume Next
    Dim frm As Access.Form
    Dim stateText As String
    If Len(Trim$(formName)) = 0 Then
        ChatGPTBuildFormStateForMessage = "Form name is not set."
        Exit Function
    End If
    Set frm = forms(formName)
    If frm Is Nothing Then
        ChatGPTBuildFormStateForMessage = "Not currently loaded in the Forms collection."
        Exit Function
    End If
    stateText = "Loaded=True"
    stateText = stateText & "; CurrentView=" & CStr(frm.CurrentView)
    stateText = stateText & "; Width=" & CStr(frm.Width)
    stateText = stateText & "; HeaderHeight=" & CStr(frm.Section(acHeader).Height)
    stateText = stateText & "; DetailHeight=" & CStr(frm.Section(acDetail).Height)
    stateText = stateText & "; ControlCount=" & CStr(frm.Controls.count)
    stateText = stateText & "; RecordSource=" & NullDisplayText(CStr(frm.recordSource))
    If Err.Number <> 0 Then
        stateText = stateText & "; StateReadError=" & CStr(Err.Number) & ": " & Err.Description
        Err.Clear
    End If
    ChatGPTBuildFormStateForMessage = stateText
    Set frm = Nothing
    On Error GoTo 0
End Function

Private Function ChatGPTFormNameCollectionSqlList(ByVal formNames As Collection, ByVal callerName As String) As String
    Dim formItem As Variant
    Dim formName As String
    Dim resultText As String
    If formNames Is Nothing Then Err.Raise vbObjectError + 6740, callerName, "No generated-form name collection was supplied."
    For Each formItem In formNames
        SetupControlPump
        formName = CStr(formItem)
        If Len(formName) = 0 Then Err.Raise vbObjectError + 6741, callerName, "The generated-form change scope contains a blank form name."
        If Len(resultText) > 0 Then resultText = resultText & ","
        resultText = resultText & SharedDao_SqlQuote(formName)
    Next formItem
    If Len(resultText) = 0 Then Err.Raise vbObjectError + 6742, callerName, "The generated-form change scope is empty."
    ChatGPTFormNameCollectionSqlList = resultText
End Function

Private Function ChatGPTFormsForRegistration(Optional ByVal normalizedBuildScope As String = FORM_BUILD_SCOPE_ALL) As Collection
    'Returns the managed-generated catalog set in the selected physical build scope after verifying existence. GeneratedFormNames remains the ADC-owned physical-builder inventory.
    Dim allFormNames As Collection
    Dim formNames As New Collection
    Dim formItem As Variant
    Dim rowData As TGeneratedFormCatalogRow
    Set allFormNames = ManagedGeneratedFormNames()
    For Each formItem In allFormNames
        SetupControlPump
        If GeneratedFormIsInBuildScope(CStr(formItem), normalizedBuildScope) Then
            If SharedDao_AccessObjectExists(GetHostDb(), acForm, CStr(formItem)) = False Then Err.Raise vbObjectError + 6710, "ChatGPTFormsForRegistration", "An authoritative managed-generated form in the selected build scope is missing and cannot be registered or grouped." & vbCrLf & "Form: " & CStr(formItem) & vbCrLf & "Build scope: " & normalizedBuildScope
            If GeneratedFormCatalogRow(CStr(formItem), rowData) = False Then Err.Raise vbObjectError + 6711, "ChatGPTFormsForRegistration", "No authoritative managed-generated form definition exists." & vbCrLf & "Form: " & CStr(formItem)
            formNames.Add CStr(formItem)
        End If
    Next formItem
    If formNames.count <> GeneratedFormAuthoritativeCount(normalizedBuildScope) Then Err.Raise vbObjectError + 6775, "ChatGPTFormsForRegistration", "The selected generated-form build scope does not map one-to-one to the managed generated-form registry." & vbCrLf & "Build scope: " & normalizedBuildScope & vbCrLf & "Scoped physical forms: " & CStr(GeneratedFormAuthoritativeCount(normalizedBuildScope)) & vbCrLf & "Scoped managed forms: " & CStr(formNames.count)
    Set ChatGPTFormsForRegistration = formNames
End Function

Private Function ChatGPTVerifiedExistingFormsCount() As Long
    If mChatGPTVerifiedExistingForms Is Nothing Then Exit Function
    ChatGPTVerifiedExistingFormsCount = mChatGPTVerifiedExistingForms.count
End Function

' Purpose: Creates blank form for design for form name.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   formName (Input; String): Physical Access form name identifying the form to create, open, configure, validate, or inspect.
' Returns: Form: Returns the newly created blank Access Form opened for design-time construction by the generated-form builder.
Public Function CreateBlankFormForDesign(ByVal formName As String) As Form
    On Error GoTo ErrHandler
    Dim frm As Form
    SetChatGPTBuildContext mChatGPTBuildStep, formName, "Create blank design form.", "Target form: " & formName
    Set frm = SharedFormDesign_CreateBlankFormForDesign(formName)
    With frm.DefaultControl(acLabel)
        .FontName = DEFAULT_UI_FONT_NAME
        .fontSize = DEFAULT_UI_FONT_SIZE
        .TextAlign = 1
    End With
    With frm.DefaultControl(acTextBox)
        .FontName = DEFAULT_UI_FONT_NAME
        .fontSize = DEFAULT_UI_FONT_SIZE
    End With
    With frm.DefaultControl(acListBox)
        .FontName = DEFAULT_UI_FONT_NAME
        .fontSize = DEFAULT_UI_FONT_SIZE
    End With
    With frm.DefaultControl(acComboBox)
        .FontName = DEFAULT_UI_FONT_NAME
        .fontSize = DEFAULT_UI_FONT_SIZE
    End With
    With frm.DefaultControl(acCommandButton)
        .FontName = DEFAULT_UI_FONT_NAME
        .fontSize = DEFAULT_UI_FONT_SIZE
    End With
    With frm.DefaultControl(acTabCtl)
        .FontName = DEFAULT_UI_FONT_NAME
        .fontSize = DEFAULT_UI_FONT_SIZE
    End With
    frm.PopUp = mChatGPTGeneratedFormsUsePopups
    Set CreateBlankFormForDesign = frm
    mChatGPTGeneratedFormDesignCreateCount = mChatGPTGeneratedFormDesignCreateCount + 1
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "CreateBlankFormForDesign", BuildChatGPTChildErrorMessage("Could not create blank generated form.", savedErrSource, savedErrDescription, "Target form: " & formName)
End Function

Private Sub CreateChatGPTAdministrativeListForm(ByVal formName As String, ByVal captionText As String, ByVal sourceName As String, ByVal fields As Variant, ByVal orderFields As Variant, ByVal actionProfile As String, ByVal allowEdits As Boolean, ByVal replaceExisting As Boolean, ByRef builtForms As Collection, ByRef skippedForms As Collection, Optional ByVal onCurrentExpression As String = vbNullString, Optional ByVal requiredForm As Boolean = False, Optional ByVal requireAllPreferredFields As Boolean = False)
    On Error GoTo ErrHandler
    Dim sqlText As String
    If Len(Trim$(sourceName)) = 0 Or TableOrQueryExists(sourceName) = False Then
        If requiredForm Then
            RaiseRequiredFormBuildIssue formName, "Required administration source was not found: " & sourceName
        Else
            AddSkippedFormNote skippedForms, formName, "Administration source was not found: " & sourceName
        End If
        Exit Sub
    End If
    If requireAllPreferredFields Then RequireAddInReviewSourceFields sourceName, fields, formName
    If requireAllPreferredFields Then
        sqlText = BuildExactGeneratedListSelectSql(formName, sourceName, fields, orderFields)
    Else
        sqlText = BuildSelectSql(sourceName, fields, vbNullString, orderFields)
    End If
    If Len(sqlText) = 0 Then
        If requiredForm Then
            RaiseRequiredFormBuildIssue formName, "Source [" & sourceName & "] did not contain any expected administration fields."
        Else
            AddSkippedFormNote skippedForms, formName, "Source [" & sourceName & "] did not contain any expected administration fields."
        End If
        Exit Sub
    End If
    If PrepareChatGPTGeneratedFormForBuild(formName, replaceExisting, skippedForms) = False Then Exit Sub
    CreateContinuousListForm formName, captionText, sqlText, sourceName, fields, False, actionProfile, allowEdits, onCurrentExpression, requireAllPreferredFields
    If requireAllPreferredFields And AddInReviewChildFormIsDefinedDefault(formName) Then
        SetChatGPTBuildContext "Creating generated administration list forms.", formName, "Validate the saved generated-form contract.", "Source: " & sourceName
        mChatGPTBuildSourceName = sourceName
        ValidateOneGeneratedFormContract formName
    End If
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
    Err.Raise savedErrNumber, "CreateChatGPTAdministrativeListForm", "Could not create a generated administration list form." & vbCrLf & "Form: " & formName & vbCrLf & "Source: " & sourceName & vbCrLf & "Error source: " & savedErrSource & vbCrLf & "Error description: " & savedErrDescription
End Sub

Private Sub CreateChatGPTConfigurationDetailForm(ByVal replaceExisting As Boolean, ByRef builtForms As Collection, ByRef skippedForms As Collection)
    On Error GoTo ErrHandler
    Dim frm As Form
    Dim formName As String
    Dim sourceName As String
    Dim tabCtl As control
    Dim headerFieldList As Variant
    Dim headerFieldCount As Long
    Dim requiredHeaderHeight As Long
    Dim layoutDetail As String
    formName = ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Configuration_Detail")
    SetChatGPTBuildContext mChatGPTBuildStep, formName, "Resolve configuration-detail source and header layout.", vbNullString
    sourceName = GeneratedFormSourceNameForSignature(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Configuration_List"))
    If Len(sourceName) = 0 Then RaiseRequiredFormBuildIssue formName, "Required configuration source was not found. Expected " & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Config_Management") & " or " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "."
    headerFieldList = HeaderFields()
    headerFieldCount = ExistingFieldCountForBuild(sourceName, headerFieldList)
    If headerFieldCount = 0 Then RaiseRequiredFormBuildIssue formName, "The resolved configuration source contains none of the required fixed-header fields. Source: " & sourceName
    requiredHeaderHeight = RequiredFieldBlockSectionHeight(sourceName, headerFieldList, 880, 3, True)
    If requiredHeaderHeight < DEFAULT_UI_CONFIG_DETAIL_HEADER_MIN_HEIGHT Then requiredHeaderHeight = DEFAULT_UI_CONFIG_DETAIL_HEADER_MIN_HEIGHT
    layoutDetail = "Source: " & sourceName & "; Existing header fields: " & CStr(headerFieldCount)
    layoutDetail = layoutDetail & "; Required header height: " & CStr(requiredHeaderHeight)
    layoutDetail = layoutDetail & "; Detail height: 8600; Tab bounds: " & BoundsForMessage(DEFAULT_UI_TAB_LEFT, DEFAULT_UI_TAB_TOP, DEFAULT_UI_TAB_WIDTH, DEFAULT_UI_TAB_HEIGHT)
    SetChatGPTBuildContext mChatGPTBuildStep, formName, "Prepare generated configuration-detail form.", layoutDetail
    If PrepareChatGPTGeneratedFormForBuild(formName, replaceExisting, skippedForms) = False Then Exit Sub
    SetChatGPTBuildContext mChatGPTBuildStep, formName, "Create blank configuration-detail form.", layoutDetail
    Set frm = CreateBlankFormForDesign(formName)
    SetChatGPTBuildContext mChatGPTBuildStep, formName, "Apply bound configuration-detail form properties.", "RecordSource: " & sourceName & "; Width: " & CStr(DEFAULT_UI_FORM_WIDTH) & "; Initial detail height: " & CStr(DEFAULT_UI_DETAIL_FORM_HEIGHT)
    ApplyBoundFormProperties frm, GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Configuration_Detail")), sourceName, DEFAULT_UI_FORM_WIDTH, DEFAULT_UI_DETAIL_FORM_HEIGHT, 0
    SetChatGPTBuildContext mChatGPTBuildStep, formName, "Ensure configuration-detail header and footer sections.", vbNullString
    EnsureHeaderFooterSections formName
    SetChatGPTBuildContext mChatGPTBuildStep, formName, "Reserve configuration-detail section heights before control creation.", layoutDetail
    forms(formName).Section(acHeader).Height = requiredHeaderHeight
    forms(formName).Section(acDetail).Height = 8600
    SetChatGPTBuildContext mChatGPTBuildStep, formName, "Create configuration-detail header title.", BoundsForMessage(DEFAULT_UI_FORM_LEFT_MARGIN, DEFAULT_UI_FORM_TOP_MARGIN, 5000, 360)
    With SharedFormDesign_AddTitleLabel(formName, acHeader, vbNullString, "Configuration Detail", DEFAULT_UI_FORM_LEFT_MARGIN, DEFAULT_UI_FORM_TOP_MARGIN, 5000, 360)
        .TextAlign = 1
    End With
    SetChatGPTBuildContext mChatGPTBuildStep, formName, "Create configuration-detail header description.", BoundsForMessage(DEFAULT_UI_FORM_LEFT_MARGIN, 500, 11000, 260)
    With SharedFormDesign_AddPlainLabel(formName, acHeader, vbNullString, "Fixed header fields stay visible while tab pages group the detail fields.", DEFAULT_UI_FORM_LEFT_MARGIN, 500, 11000, 260, DEFAULT_UI_FONT_SIZE)
        .TextAlign = 1
    End With
    SetChatGPTBuildContext mChatGPTBuildStep, formName, "Create configuration-detail fixed-header field block.", layoutDetail
    AddFieldBlock formName, sourceName, headerFieldList, acHeader, vbNullString, DEFAULT_UI_FORM_LEFT_MARGIN, 880, 3, "hdr", True
    SetChatGPTBuildContext mChatGPTBuildStep, formName, "Create configuration-detail tab control.", "Control: tabConfigurationDetail; Bounds: " & BoundsForMessage(DEFAULT_UI_TAB_LEFT, DEFAULT_UI_TAB_TOP, DEFAULT_UI_TAB_WIDTH, DEFAULT_UI_TAB_HEIGHT)
    Set tabCtl = SharedFormDesign_AddTabControl(formName, "tabConfigurationDetail", DEFAULT_UI_TAB_LEFT, DEFAULT_UI_TAB_TOP, DEFAULT_UI_TAB_WIDTH, DEFAULT_UI_TAB_HEIGHT)
    SetChatGPTBuildContext mChatGPTBuildStep, formName, "Configure configuration-detail tab pages and field blocks.", "Tab control: " & tabCtl.Name & "; Expected pages: 4"
    ConfigureConfigurationDetailTabs formName, sourceName, tabCtl
    SetChatGPTBuildContext mChatGPTBuildStep, formName, "Verify completed configuration-detail layout before save.", layoutDetail
    VerifyConfigurationDetailFormLayout formName, sourceName, headerFieldList, requiredHeaderHeight
    SetChatGPTBuildContext mChatGPTBuildStep, formName, "Save and close verified configuration-detail form.", layoutDetail
    SaveAndCloseDesignForm formName
    builtForms.Add formName
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim failureText As String
    Dim liveFormState As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    liveFormState = ChatGPTBuildFormStateForMessage(formName)
    On Error Resume Next
    SharedForm_CloseIfOpen formName, acSaveNo
    On Error GoTo 0
    If Len(Trim$(savedErrSource)) = 0 Then savedErrSource = "CreateChatGPTConfigurationDetailForm"
    failureText = "Could not create the generated configuration-detail form."
    failureText = failureText & vbCrLf & "Form: " & NullDisplayText(formName)
    failureText = failureText & vbCrLf & "Source: " & NullDisplayText(sourceName)
    failureText = failureText & vbCrLf & "Existing header field count: " & CStr(headerFieldCount)
    failureText = failureText & vbCrLf & "Required header height: " & CStr(requiredHeaderHeight)
    failureText = failureText & vbCrLf & "Build action: " & NullDisplayText(mChatGPTBuildAction)
    failureText = failureText & vbCrLf & "Build detail: " & NullDisplayText(mChatGPTBuildDetail)
    failureText = failureText & vbCrLf & "Live form state before cleanup: " & liveFormState
    failureText = failureText & vbCrLf & "Original source: " & savedErrSource
    failureText = failureText & vbCrLf & "Original description: " & savedErrDescription
    Err.Raise savedErrNumber, savedErrSource, failureText
End Sub

Private Sub CreateChatGPTFormTypesForm(ByVal replaceExisting As Boolean, ByRef builtForms As Collection, ByRef skippedForms As Collection)
    On Error GoTo ErrHandler
    Dim formName As String
    Dim sourceName As String
    Dim sqlText As String
    Dim fields As Variant
    Dim orderFields As Variant
    formName = ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Form_Types")
    sourceName = GeneratedFormSourceNameForSignature(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Form_Types"))
    If Len(sourceName) = 0 Then
        AddSkippedFormNote skippedForms, formName, "Source was not found. " & "Expected " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_FormTypes") & "."
        Exit Sub
    End If
    If TableOrQueryExists(sourceName) = False Then
        AddSkippedFormNote skippedForms, formName, "Source [" & sourceName & "] was not found. " & "Expected " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_FormTypes") & "."
        Exit Sub
    End If
    fields = GeneratedFormListFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Form_Types"))
    orderFields = GeneratedFormOrderFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Form_Types"))
    sqlText = BuildSelectSql(sourceName, fields, vbNullString, orderFields)
    If Len(sqlText) = 0 Then
        AddSkippedFormNote skippedForms, formName, "Source [" & sourceName & "] did not contain any expected fields."
        Exit Sub
    End If
    If PrepareChatGPTGeneratedFormForBuild(formName, replaceExisting, skippedForms) = False Then Exit Sub
    CreateContinuousListForm formName, GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Form_Types")), sqlText, sourceName, fields, GeneratedFormIncludeOpenButton(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Form_Types")), GeneratedFormListActionProfile(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Form_Types")), GeneratedFormListAllowEdits(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Form_Types"))
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
    Err.Raise savedErrNumber, "CreateChatGPTFormTypesForm", savedErrDescription & vbCrLf & "Error source: " & savedErrSource
End Sub

Private Sub CreateChatGPTCustomMacroLauncherReviewForm(ByVal replaceExisting As Boolean, ByRef builtForms As Collection, ByRef skippedForms As Collection)
    On Error GoTo ErrHandler
    Dim formName As String
    Dim sourceName As String
    Dim sqlText As String
    Dim fields As Variant
    Dim orderFields As Variant
    formName = ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "LauncherControlsReviewForm")
    sourceName = GeneratedFormSourceNameForSignature(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "LauncherControlsReviewForm"))
    If Len(sourceName) = 0 Then
        AddSkippedFormNote skippedForms, formName, "Source was not found. " & "Expected " & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_LauncherControls_Review") & " or " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LauncherControlsReview") & "."
        Exit Sub
    End If
    If TableOrQueryExists(sourceName) = False Then
        AddSkippedFormNote skippedForms, formName, "Source [" & sourceName & "] was not found. " & "Expected " & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_LauncherControls_Review") & " or " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LauncherControlsReview") & "."
        Exit Sub
    End If
    fields = GeneratedFormListFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "LauncherControlsReviewForm"))
    orderFields = GeneratedFormOrderFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "LauncherControlsReviewForm"))
    sqlText = BuildSelectSql(sourceName, fields, vbNullString, orderFields)
    If Len(sqlText) = 0 Then
        AddSkippedFormNote skippedForms, formName, "Source [" & sourceName & "] did not contain any expected fields."
        Exit Sub
    End If
    If PrepareChatGPTGeneratedFormForBuild(formName, replaceExisting, skippedForms) = False Then Exit Sub
    CreateContinuousListForm formName, GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "LauncherControlsReviewForm")), sqlText, sourceName, fields, GeneratedFormIncludeOpenButton(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "LauncherControlsReviewForm")), GeneratedFormListActionProfile(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "LauncherControlsReviewForm")), GeneratedFormListAllowEdits(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "LauncherControlsReviewForm"))
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
    Err.Raise savedErrNumber, "CreateChatGPTCustomMacroLauncherReviewForm", savedErrDescription & vbCrLf & "Error source: " & savedErrSource
End Sub

' Purpose: Performs ensure report error dialog available by coordinating `ManagedObjectPhysicalName` and `SharedDao_AccessObjectExists` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: Boolean: True on the function's affirmative/success branch for ensure report error dialog available; False on its explicit negative, unavailable, rejected, or handled-failure branch.
'   path.
Public Function EnsureReportErrorDialogAvailable() As Boolean
    ' PRE-D ERROR-REPORTING BOOTSTRAP EXCEPTION: Setup can fail before D1/D2, so the canonical Report Error form must be independently materializable by the terminal error path before the normal generated-form
    ' phase. This exception is limited to Report Error; D1-D4 retain ordinary generated-form lifecycle ownership and D3 later verifies/converges this ADC Support Form before D4 complete reconciliation.
    ' Intended use: ensure the one Setup-managed Report Error form exists and persists its required modal/popup presentation before the terminal error workflow tries to open it, including failures that occur before
    ' the normal generated-form phase has run.
    On Error GoTo ErrHandler
    Dim formName As String
    Dim builtForms As Collection
    Dim skippedForms As Collection
    Dim effectiveUsePopups As Boolean
    Dim openedForPresentationRepair As Boolean
    formName = ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Report_Error_Dialog")
    effectiveUsePopups = SettingBool(SETTING_GENERATED_FORMS_USE_POPUPS, DEFAULT_GENERATED_FORMS_USE_POPUPS)
    If SharedDao_AccessObjectExists(GetHostDb(), acForm, formName) Then
        If CurrentProject.AllForms(formName).IsLoaded = False Then
            DoCmd.OpenForm formName, acDesign, , , , acHidden
            openedForPresentationRepair = True
            If forms(formName).Modal = False Or forms(formName).PopUp <> effectiveUsePopups Then
                forms(formName).Modal = True
                forms(formName).PopUp = effectiveUsePopups
                DoCmd.Close acForm, formName, acSaveYes
            Else
                DoCmd.Close acForm, formName, acSaveNo
            End If
            openedForPresentationRepair = False
        End If
        CompleteStandaloneGeneratedFormDescriptionIfAllowed formName
        EnsureReportErrorDialogAvailable = True
        Exit Function
    End If
    mChatGPTGeneratedFormsUsePopups = effectiveUsePopups
    ' INTENTIONAL FILTERED PARTITION: builtForms, skippedForms is a consumer-required subset/projection derived from the source enumerated by EnsureReportErrorDialogAvailable; its membership is not independently
    ' authoritative.
    Set builtForms = New Collection
    Set skippedForms = New Collection
    CreateReportErrorDialogForm False, builtForms, skippedForms, True
    If SharedDao_AccessObjectExists(GetHostDb(), acForm, formName) = False Then Err.Raise vbObjectError + 9147, "EnsureReportErrorDialogAvailable", "The canonical Report Error form builder returned without creating the required form." & vbCrLf & "Form: " & formName
    CompleteStandaloneGeneratedFormDescriptionIfAllowed formName
    EnsureReportErrorDialogAvailable = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If openedForPresentationRepair Then DoCmd.Close acForm, formName, acSaveNo
    On Error GoTo 0
    Err.Raise savedErrNumber, "EnsureReportErrorDialogAvailable", "Could not ensure the managed Report Error form was available with its required modal/popup presentation." & vbCrLf & "Form: " & formName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Sub CompleteStandaloneGeneratedFormDescriptionIfAllowed(ByVal formName As String)
    'Emergency/pre-D and direct D1 form creation may complete managed metadata outside numbered Setup. During D1-D3 the central guard keeps physical Description mutation deferred to G1.
    Dim descriptionChanged As Variant
    If SetupDomainMutationWritesAllowedForCurrentPart(SETUP_DOMAIN_G1_DESCRIPTIONS) = False Then Exit Sub
    descriptionChanged = False
    If ApplyAndVerifyManagedObjectDescriptionForStandaloneRepair(formName, ADCAT_OBJECT_TYPE_FORM, descriptionChanged) = False Then Err.Raise vbObjectError + 10130, "CompleteStandaloneGeneratedFormDescriptionIfAllowed", "Managed generated-form Description reconciliation returned False." & vbCrLf & "Form: " & formName
End Sub

Private Sub CreateReportErrorDialogForm(ByVal replaceExisting As Boolean, ByRef builtForms As Collection, ByRef skippedForms As Collection, Optional ByVal ignoreScopedBuildPlan As Boolean = False)
    On Error GoTo ErrHandler
    Dim frm As Form
    Dim ctl As control
    Dim formName As String
    Dim controlDefinitions As Collection
    Dim callbackExpressions As Collection
    formName = ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Report_Error_Dialog")
    If PrepareChatGPTGeneratedFormForBuild(formName, replaceExisting, skippedForms, ignoreScopedBuildPlan) = False Then Exit Sub
    Set controlDefinitions = GeneratedFormContractControlDefinitions(formName)
    Set callbackExpressions = GeneratedFormContractCallbackExpressions(formName)
    Set frm = CreateBlankFormForDesign(formName)
    ApplyUnboundFormProperties frm, GeneratedFormDisplayName(formName), 9000, 5600
    frm.ScrollBars = 0
    frm.AutoResize = True
    frm.Modal = True
    frm.BorderStyle = 3
    frm.MinMaxButtons = 0
    frm.CloseButton = False
    With SharedFormDesign_AddPlainLabel(formName, acDetail, vbNullString, "Describe the error. The description is required. Select Submit Report to hide this dialog and capture the visible Access application window.", DEFAULT_UI_FORM_LEFT_MARGIN, DEFAULT_UI_FORM_TOP_MARGIN, 8300, 600, DEFAULT_UI_FONT_SIZE)
        .TextAlign = 1
    End With
    Set ctl = CreateControl(formName, acTextBox, acDetail, vbNullString, vbNullString, DEFAULT_UI_FORM_LEFT_MARGIN, DEFAULT_UI_FORM_TOP_MARGIN + 600 + DEFAULT_UI_ROW_GAP, 8300, 3700)
    ctl.Name = GeneratedFormContractControlName(formName, 1, controlDefinitions)
    ctl.fontSize = DEFAULT_UI_READING_FONT_SIZE
    ctl.ScrollBars = 2
    ctl.EnterKeyBehavior = True
    ctl.TabIndex = 0
    ctl.statusBarText = "Enter the required error description."
    Set ctl = SharedFormDesign_AddButton(formName, acDetail, vbNullString, "Submit Report", GeneratedFormContractCallbackExpression(formName, 1, callbackExpressions), 5850, DEFAULT_UI_FORM_TOP_MARGIN + 600 + DEFAULT_UI_ROW_GAP + 3700 + DEFAULT_UI_BUTTON_ROW_GAP, 1450, 450, "Submit the report, capture the Access window, and append the report to retained diagnostics.")
    ctl.Name = "cmdSubmit"
    ctl.Default = True
    ctl.TabIndex = 1
    Set ctl = SharedFormDesign_AddButton(formName, acDetail, vbNullString, "Cancel", GeneratedFormContractCallbackExpression(formName, 2, callbackExpressions), 7450, DEFAULT_UI_FORM_TOP_MARGIN + 600 + DEFAULT_UI_ROW_GAP + 3700 + DEFAULT_UI_BUTTON_ROW_GAP, 1150, 450, "Close the Report Error form without creating a report.")
    ctl.Name = "cmdCancel"
    ctl.Cancel = True
    ctl.TabIndex = 2
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
    Err.Raise savedErrNumber, "CreateReportErrorDialogForm", "Could not create the managed Report Error dialog." & vbCrLf & "Form: " & formName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub CreateDeleteObjectsDialogForm(ByVal replaceExisting As Boolean, ByRef builtForms As Collection, ByRef skippedForms As Collection)
    On Error GoTo ErrHandler
    Dim frm As Form
    Dim ctl As control
    Dim formName As String
    Dim callbackExpressions As Collection
    formName = ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Delete_ADC_Objects")
    If PrepareChatGPTGeneratedFormForBuild(formName, replaceExisting, skippedForms) = False Then Exit Sub
    Set callbackExpressions = GeneratedFormContractCallbackExpressions(formName)
    Set frm = CreateBlankFormForDesign(formName)
    ApplyUnboundFormProperties frm, GeneratedFormDisplayName(formName), 6800, 3500
    frm.OnLoad = GeneratedFormContractOnLoadExpression(formName)
    frm.ScrollBars = 0
    frm.PopUp = True
    frm.Modal = True
    frm.BorderStyle = 3
    frm.ControlBox = False
    Set ctl = CreateControl(formName, acLabel, acDetail, vbNullString, vbNullString, 450, 350, 5900, 420)
    With ctl
        .TextAlign = 1
        .Caption = "What do you want to delete?"
        .FontName = DEFAULT_UI_FONT_NAME
        .fontSize = DEFAULT_UI_FONT_SIZE
        .fontBold = True
    End With
    Set ctl = CreateControl(formName, acCommandButton, acDetail, vbNullString, vbNullString, 450, 1050, 1750, 525)
    ctl.Name = "cmdDeleteTables"
    ctl.Caption = "Tables"
    ctl.FontName = DEFAULT_UI_FONT_NAME
    ctl.fontSize = DEFAULT_UI_FONT_SIZE
    ctl.OnClick = GeneratedFormContractCallbackExpression(formName, 1, callbackExpressions)
    Set ctl = CreateControl(formName, acCommandButton, acDetail, vbNullString, vbNullString, 2500, 1050, 1750, 525)
    ctl.Name = "cmdDeleteQueries"
    ctl.Caption = "Queries"
    ctl.FontName = DEFAULT_UI_FONT_NAME
    ctl.fontSize = DEFAULT_UI_FONT_SIZE
    ctl.OnClick = GeneratedFormContractCallbackExpression(formName, 2, callbackExpressions)
    Set ctl = CreateControl(formName, acCommandButton, acDetail, vbNullString, vbNullString, 4550, 1050, 1750, 525)
    ctl.Name = "cmdDeleteGeneratedForms"
    ctl.Caption = "Generated Forms"
    ctl.FontName = DEFAULT_UI_FONT_NAME
    ctl.fontSize = DEFAULT_UI_FONT_SIZE
    ctl.OnClick = GeneratedFormContractCallbackExpression(formName, 3, callbackExpressions)
    Set ctl = CreateControl(formName, acCommandButton, acDetail, vbNullString, vbNullString, 450, 1750, 1750, 525)
    ctl.Name = "cmdDeleteMacros"
    ctl.Caption = "Macros"
    ctl.FontName = DEFAULT_UI_FONT_NAME
    ctl.fontSize = DEFAULT_UI_FONT_SIZE
    ctl.OnClick = GeneratedFormContractCallbackExpression(formName, 4, callbackExpressions)
    Set ctl = CreateControl(formName, acCommandButton, acDetail, vbNullString, vbNullString, 2500, 1750, 1750, 525)
    ctl.Name = "cmdDeleteReports"
    ctl.Caption = "Reports"
    ctl.FontName = DEFAULT_UI_FONT_NAME
    ctl.fontSize = DEFAULT_UI_FONT_SIZE
    ctl.OnClick = GeneratedFormContractCallbackExpression(formName, 5, callbackExpressions)
    Set ctl = CreateControl(formName, acCommandButton, acDetail, vbNullString, vbNullString, 4550, 1750, 1750, 525)
    ctl.Name = "cmdDeleteAll"
    ctl.Caption = "All"
    ctl.FontName = DEFAULT_UI_FONT_NAME
    ctl.fontSize = DEFAULT_UI_FONT_SIZE
    ctl.OnClick = GeneratedFormContractCallbackExpression(formName, 6, callbackExpressions)
    Set ctl = CreateControl(formName, acCommandButton, acDetail, vbNullString, vbNullString, 2875, 2450, 1050, 525)
    ctl.Name = "cmdCancel"
    ctl.Caption = "Cancel"
    ctl.FontName = DEFAULT_UI_FONT_NAME
    ctl.fontSize = DEFAULT_UI_FONT_SIZE
    ctl.Cancel = True
    ctl.OnClick = GeneratedFormContractCallbackExpression(formName, 7, callbackExpressions)
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
    Err.Raise savedErrNumber, "CreateDeleteObjectsDialogForm", "Could not create the managed Delete ADC Objects dialog." & vbCrLf & "Form: " & formName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub CreateChatGPTManagerForm(ByVal replaceExisting As Boolean, ByRef builtForms As Collection, ByRef skippedForms As Collection)
    On Error GoTo ErrHandler
    Dim frm As Form
    Dim formName As String
    Dim topPos As Long
    Dim left1 As Long
    Dim left2 As Long
    Dim left3 As Long
    formName = ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Manager")
    left1 = DEFAULT_UI_FORM_LEFT_MARGIN
    left2 = left1 + DEFAULT_UI_BUTTON_WIDTH + DEFAULT_UI_COLUMN_GAP
    left3 = left2 + DEFAULT_UI_BUTTON_WIDTH + DEFAULT_UI_COLUMN_GAP
    If PrepareChatGPTGeneratedFormForBuild(formName, replaceExisting, skippedForms) = False Then Exit Sub
    Set frm = CreateBlankFormForDesign(formName)
    ApplyUnboundFormProperties frm, GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Manager")), DEFAULT_UI_FORM_WIDTH, DEFAULT_UI_MANAGER_FORM_HEIGHT
    EnsureHeaderFooterSections formName
    forms(formName).Section(acHeader).Height = 900
    forms(formName).Section(acDetail).Height = 15500
    With SharedFormDesign_AddTitleLabel(formName, acHeader, vbNullString, GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Manager")), DEFAULT_UI_FORM_LEFT_MARGIN, DEFAULT_UI_FORM_TOP_MARGIN, 9000, 420)
        .TextAlign = 1
    End With
    With SharedFormDesign_AddPlainLabel(formName, acHeader, vbNullString, "Generated management forms, split/popup launchers, filter helpers, and shortcuts for the " & DEFAULT_UTILITY_LONG_NAME & " utility.", DEFAULT_UI_FORM_LEFT_MARGIN, DEFAULT_UI_FORM_TOP_MARGIN + 420 + DEFAULT_UI_SECTION_TITLE_CONTENT_GAP, 11000, 260, DEFAULT_UI_FONT_SIZE)
        .TextAlign = 1
    End With
    topPos = DEFAULT_UI_FORM_TOP_MARGIN
    With SharedFormDesign_AddSectionLabel(formName, acDetail, vbNullString, "Forms", left1, topPos, 5000, DEFAULT_UI_SECTION_TITLE_HEIGHT)
        .TextAlign = 1
    End With
    topPos = topPos + DEFAULT_UI_SECTION_TITLE_HEIGHT + DEFAULT_UI_SECTION_TITLE_CONTENT_GAP
    SharedFormDesign_AddButton formName, acDetail, vbNullString, "Tabbed Utility", "=ChatGPT_OpenFormByName(""" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Utility") & """)", left1, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Open the original single-form tabbed utility shell."
    SharedFormDesign_AddButton formName, acDetail, vbNullString, "Help", "=ChatGPT_OpenFormByName(""" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Help") & """)", left2, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Open searchable " & DEFAULT_UTILITY_LONG_NAME & " Help and the database reference."
    topPos = topPos + DEFAULT_UI_BUTTON_HEIGHT + DEFAULT_UI_BUTTON_ROW_GAP
    SharedFormDesign_AddButton formName, acDetail, vbNullString, "Configuration Detail", "=ChatGPT_OpenFormByName(""" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Configuration_Detail") & """)", left1, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Open the fixed-header/tabbed configuration editor."
    SharedFormDesign_AddButton formName, acDetail, vbNullString, "Configuration List", "=ChatGPT_OpenFormByName(""" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Configuration_List") & """)", left2, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Open the continuous configuration list."
    SharedFormDesign_AddButton formName, acDetail, vbNullString, "Candidate Review", "=ChatGPT_OpenFormByName(""" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Candidate_Review") & """)", left3, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Open candidate-review rows for review and acceptance."
    topPos = topPos + DEFAULT_UI_BUTTON_HEIGHT + DEFAULT_UI_BUTTON_ROW_GAP
    SharedFormDesign_AddButton formName, acDetail, vbNullString, "ADC_Settings", "=ChatGPT_OpenFormByName(""" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Settings") & """)", left1, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Open " & DEFAULT_UTILITY_LONG_NAME & " settings."
    SharedFormDesign_AddButton formName, acDetail, vbNullString, "Health Review", "=ChatGPT_OpenFormByName(""" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Health_Review") & """)", left2, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Open health-review findings."
    SharedFormDesign_AddButton formName, acDetail, vbNullString, "Custom Macro Launcher Review", "=ChatGPT_OpenFormByName(""" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "LauncherControlsReviewForm") & """)", left3, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Open Custom Macro Launcher review rows."
    topPos = topPos + DEFAULT_UI_BUTTON_HEIGHT + DEFAULT_UI_BUTTON_ROW_GAP
    SharedFormDesign_AddButton formName, acDetail, vbNullString, "Catalog Review", "=ChatGPT_OpenFormByName(""" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Catalog_Review") & """)", left1, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Open [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID) & "] catalog review rows."
    SharedFormDesign_AddButton formName, acDetail, vbNullString, "System Macros", "=ChatGPT_OpenFormByName(""" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "System_Macros") & """)", left2, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Open table-driven built-in system macro rows."
    SharedFormDesign_AddButton formName, acDetail, vbNullString, "Navigation Pane Groups", "=ChatGPT_OpenFormByName(""" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "NavigationPane_Groups") & """)", left3, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Open editable Navigation Pane group catalog rows."
    topPos = topPos + DEFAULT_UI_BUTTON_HEIGHT + DEFAULT_UI_BUTTON_ROW_GAP
    SharedFormDesign_AddButton formName, acDetail, vbNullString, "Form Types", "=ChatGPT_OpenFormByName(""" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Form_Types") & """)", left1, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Open form type lookup rows."
    SharedFormDesign_AddButton formName, acDetail, vbNullString, "Forms Catalog", "=ChatGPT_OpenFormByName(""" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Forms") & """)", left2, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Open relevant " & DEFAULT_UTILITY_LONG_NAME & " form rows."
    topPos = topPos + DEFAULT_UI_BUTTON_HEIGHT + DEFAULT_UI_SECTION_GAP
    With SharedFormDesign_AddSectionLabel(formName, acDetail, vbNullString, "References and Update Guards", left1, topPos, 5000, DEFAULT_UI_SECTION_TITLE_HEIGHT)
        .TextAlign = 1
    End With
    topPos = topPos + DEFAULT_UI_SECTION_TITLE_HEIGHT + DEFAULT_UI_SECTION_TITLE_CONTENT_GAP
    SharedFormDesign_AddButton formName, acDetail, vbNullString, "Active References", "=ChatGPT_OpenFormByName(""" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Active_References") & """)", left1, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Open the generated active-reference inventory and loaded-project match status."
    SharedFormDesign_AddButton formName, acDetail, vbNullString, "Available References", "=ChatGPT_OpenFormByName(""" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Available_References") & """)", left2, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Open the generated available-reference inventory."
    SharedFormDesign_AddButton formName, acDetail, vbNullString, "Loaded VBA Projects", "=ChatGPT_OpenFormByName(""" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review_Projects") & """)", left3, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Open the reusable loaded-project inventory with source-analysis status and counts."
    topPos = topPos + DEFAULT_UI_BUTTON_HEIGHT + DEFAULT_UI_BUTTON_ROW_GAP
    SharedFormDesign_AddButton formName, acDetail, vbNullString, "Reference Manager", "=ChatGPT_OpenFormByName(""" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Reference_Manager") & """)", left1, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Open the generated reference-management view."
    SharedFormDesign_AddButton formName, acDetail, vbNullString, "Current Module Manifest", "=ChatGPT_OpenFormByName(""" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Updates_Current") & """)", left2, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Open expected-current module versions and verification status."
    topPos = topPos + DEFAULT_UI_BUTTON_HEIGHT + DEFAULT_UI_BUTTON_ROW_GAP
    SharedFormDesign_AddButton formName, acDetail, vbNullString, "Module Update History", "=ChatGPT_OpenFormByName(""" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Updates_History") & """)", left1, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Open append-only module update history."
    SharedFormDesign_AddButton formName, acDetail, vbNullString, "Module Verification", "=ChatGPT_OpenFormByName(""" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Verification") & """)", left2, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Open expected-versus-installed verification snapshots."
    SharedFormDesign_AddButton formName, acDetail, vbNullString, "Verify Module Guard", "=ChatGPT_VerifyModuleGuardFromGeneratedForm()", left3, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Run the module guard manually even when automatic enforcement is disabled."
    topPos = topPos + DEFAULT_UI_BUTTON_HEIGHT + DEFAULT_UI_BUTTON_ROW_GAP
    SharedFormDesign_AddButton formName, acDetail, vbNullString, "Repair Module Guard", "=ChatGPT_RunProcedure(""RepairModuleUpdateGuard"")", left1, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Create or repair module guard tables and queries."
    topPos = topPos + DEFAULT_UI_BUTTON_HEIGHT + DEFAULT_UI_SECTION_GAP
    With SharedFormDesign_AddSectionLabel(formName, acDetail, vbNullString, "Split / Popup Forms", left1, topPos, 5000, DEFAULT_UI_SECTION_TITLE_HEIGHT)
        .TextAlign = 1
    End With
    topPos = topPos + DEFAULT_UI_SECTION_TITLE_HEIGHT + DEFAULT_UI_SECTION_TITLE_CONTENT_GAP
    SharedFormDesign_AddButton formName, acDetail, vbNullString, "Configuration Split", "=ChatGPT_OpenFormByName(""" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Configuration_Split") & """)", left1, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Open the split-view configuration form."
    SharedFormDesign_AddButton formName, acDetail, vbNullString, "Candidate Split", "=ChatGPT_OpenFormByName(""" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Candidate_Split") & """)", left2, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Open the split-view candidate review form."
    SharedFormDesign_AddButton formName, acDetail, vbNullString, "Settings Split", "=ChatGPT_OpenFormByName(""" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Settings_Split") & """)", left3, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Open the split-view settings form."
    topPos = topPos + DEFAULT_UI_BUTTON_HEIGHT + DEFAULT_UI_BUTTON_ROW_GAP
    SharedFormDesign_AddButton formName, acDetail, vbNullString, "Health Split", "=ChatGPT_OpenFormByName(""" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Health_Split") & """)", left1, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Open the split-view health review form."
    SharedFormDesign_AddButton formName, acDetail, vbNullString, "Launcher Split", "=ChatGPT_OpenFormByName(""" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "LauncherControls_Split") & """)", left2, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Open the split-view Custom Macro Launcher review form."
    SharedFormDesign_AddButton formName, acDetail, vbNullString, "Catalog Split", "=ChatGPT_OpenFormByName(""" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Catalog_Split") & """)", left3, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Open the split-view catalog review form."
    topPos = topPos + DEFAULT_UI_BUTTON_HEIGHT + DEFAULT_UI_BUTTON_ROW_GAP
    SharedFormDesign_AddButton formName, acDetail, vbNullString, "Forms Catalog Split", "=ChatGPT_OpenFormByName(""" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Forms_Split") & """)", left1, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Open the split-view forms catalog form."
    topPos = topPos + DEFAULT_UI_BUTTON_HEIGHT + DEFAULT_UI_SECTION_GAP
    With SharedFormDesign_AddSectionLabel(formName, acDetail, vbNullString, "Filter Helpers", left1, topPos, 5000, DEFAULT_UI_SECTION_TITLE_HEIGHT)
        .TextAlign = 1
    End With
    topPos = topPos + DEFAULT_UI_SECTION_TITLE_HEIGHT + DEFAULT_UI_SECTION_TITLE_CONTENT_GAP
    SharedFormDesign_AddButton formName, acDetail, vbNullString, "Filter Current Selection", "=ChatGPT_FilterCurrentSelection()", left1, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Apply Access Filter By Selection to the active split/popup form."
    SharedFormDesign_AddButton formName, acDetail, vbNullString, "Clear Current Filter", "=ChatGPT_ClearCurrentFormFilter()", left2, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Clear the currently open Access form filter and sort."
    topPos = topPos + DEFAULT_UI_BUTTON_HEIGHT + DEFAULT_UI_SECTION_GAP
    With SharedFormDesign_AddSectionLabel(formName, acDetail, vbNullString, "Utility Actions", left1, topPos, 5000, DEFAULT_UI_SECTION_TITLE_HEIGHT)
        .TextAlign = 1
    End With
    topPos = topPos + DEFAULT_UI_SECTION_TITLE_HEIGHT + DEFAULT_UI_SECTION_TITLE_CONTENT_GAP
    AddCoreSystemMacroUtilityActionButtons formName, vbNullString, left1, left2, left3, topPos, False
    forms(formName).Section(acDetail).Height = topPos + DEFAULT_UI_BUTTON_HEIGHT + DEFAULT_UI_FORM_SECTION_BOTTOM_MARGIN
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
    Err.Raise savedErrNumber, "CreateChatGPTManagerForm", savedErrDescription & vbCrLf & "Error source: " & savedErrSource
End Sub

Private Sub CreateChatGPTTabbedUtilityForm(ByVal replaceExisting As Boolean, ByRef builtForms As Collection, ByRef skippedForms As Collection)
    On Error GoTo ErrHandler
    Dim frm As Form
    Dim formName As String
    Dim tabCtl As control
    formName = ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Utility")
    If PrepareChatGPTGeneratedFormForBuild(formName, replaceExisting, skippedForms) = False Then Exit Sub
    Set frm = CreateBlankFormForDesign(formName)
    ApplyUnboundFormProperties frm, GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Utility")), DEFAULT_UI_FORM_WIDTH, DEFAULT_UI_UTILITY_FORM_HEIGHT
    EnsureHeaderFooterSections formName
    forms(formName).Section(acHeader).Height = 900
    forms(formName).Section(acDetail).Height = 8800
    With SharedFormDesign_AddTitleLabel(formName, acHeader, vbNullString, GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Utility")), DEFAULT_UI_FORM_LEFT_MARGIN, DEFAULT_UI_FORM_TOP_MARGIN, 9000, 420)
        .TextAlign = 1
    End With
    With SharedFormDesign_AddPlainLabel(formName, acHeader, vbNullString, "Single generated utility shell. Existing generated forms are embedded on tabs and remain available as separate forms.", DEFAULT_UI_FORM_LEFT_MARGIN, DEFAULT_UI_FORM_TOP_MARGIN + 420 + DEFAULT_UI_SECTION_TITLE_CONTENT_GAP, 11400, 260, DEFAULT_UI_FONT_SIZE)
        .TextAlign = 1
    End With
    Set tabCtl = SharedFormDesign_AddTabControl(formName, "tabADCUtility", DEFAULT_UI_TAB_LEFT, DEFAULT_UI_TAB_TOP, DEFAULT_UI_TAB_WIDTH, DEFAULT_UI_TAB_HEIGHT + 600)
    ConfigureTabbedUtilityShell formName, tabCtl
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
    Err.Raise savedErrNumber, "CreateChatGPTTabbedUtilityForm", savedErrDescription & vbCrLf & "Error source: " & savedErrSource
End Sub

Private Sub CreateConfigurationListLikeForm(ByVal formName As String, ByVal captionText As String, ByVal sourceName As String, ByVal whereText As String, ByVal replaceExisting As Boolean, ByRef builtForms As Collection, ByRef skippedForms As Collection, Optional ByVal requiredForm As Boolean = False)
    On Error GoTo ErrHandler
    Dim sqlText As String
    Dim fields As Variant
    Dim orderFields As Variant
    Dim enforceExactLookupEditorContract As Boolean
    If Len(sourceName) = 0 Then
        If requiredForm Then
            RaiseRequiredFormBuildIssue formName, "Required configuration/candidate source was not found. Expected " & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Config_Management") & ", " & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Config_CandidateReview") & ", or " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_Config") & "."
        Else
            AddSkippedFormNote skippedForms, formName, "Source was not found."
        End If
        Exit Sub
    End If
    If TableOrQueryExists(sourceName) = False Then
        If requiredForm Then
            RaiseRequiredFormBuildIssue formName, "Required source [" & sourceName & "] was not found."
        Else
            AddSkippedFormNote skippedForms, formName, "Source [" & sourceName & "] was not found."
        End If
        Exit Sub
    End If
    fields = GeneratedFormListFields(formName)
    orderFields = GeneratedFormOrderFields(formName)
    enforceExactLookupEditorContract = LookupEditorGeneratedFormIsDefined(formName)
    If enforceExactLookupEditorContract Then
        sqlText = BuildExactGeneratedListSelectSql(formName, sourceName, fields, orderFields, whereText)
    Else
        sqlText = BuildSelectSql(sourceName, fields, whereText, orderFields)
    End If
    If Len(sqlText) = 0 Then
        If requiredForm Then
            RaiseRequiredFormBuildIssue formName, "Required source [" & sourceName & "] did not contain any expected fields."
        Else
            AddSkippedFormNote skippedForms, formName, "Source [" & sourceName & "] did not contain any expected fields."
        End If
        Exit Sub
    End If
    If PrepareChatGPTGeneratedFormForBuild(formName, replaceExisting, skippedForms) = False Then Exit Sub
    CreateContinuousListForm formName, captionText, sqlText, sourceName, fields, GeneratedFormIncludeOpenButton(formName), GeneratedFormListActionProfile(formName), GeneratedFormListAllowEdits(formName), useExactPreferredFields:=enforceExactLookupEditorContract
    If enforceExactLookupEditorContract Then
        SetChatGPTBuildContext "Creating canonical generated lookup editor.", formName, "Validate the saved exact lookup-editor contract.", "Source: " & sourceName
        mChatGPTBuildSourceName = sourceName
        ValidateOneGeneratedFormContract formName
    End If
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
    Err.Raise savedErrNumber, "CreateConfigurationListLikeForm", savedErrDescription & vbCrLf & "Error source: " & savedErrSource
End Sub

Private Sub CreateContinuousListForm(ByVal formName As String, ByVal captionText As String, ByVal recordSourceText As String, ByVal fieldSourceName As String, ByVal preferredFields As Variant, ByVal includeOpenButton As Boolean, Optional ByVal actionProfile As String = vbNullString, Optional ByVal allowEdits As Boolean = True, Optional ByVal onCurrentExpression As String = vbNullString, Optional ByVal useExactPreferredFields As Boolean = False)
    On Error GoTo ErrHandler
    Const LIST_HEADER_TOP_FIRST As Long = 650
    Const LIST_OPEN_COLUMN_WIDTH As Long = 700
    Dim frm As Form
    Dim definitions As Collection
    Dim Definition As Variant
    Dim createdControlNames As Object
    Dim leftPos As Long
    Dim headerTop As Long
    Dim detailTop As Long
    Dim headerBottom As Long
    Dim detailBottom As Long
    Dim fieldName As String
    Dim labelText As String
    Dim labelWidthValue As Long
    Dim controlLeft As Long
    Dim widthValue As Long
    Dim pairWidth As Long
    Dim maxControlWidth As Long
    Dim controlTypeValue As Long
    Dim lockedValue As Boolean
    Dim projectedControlName As String
    Dim ctl As control
    Dim lbl As control
    Dim detailText As String
    Dim enforceExactControlContract As Boolean
    enforceExactControlContract = useExactPreferredFields
    Set definitions = ContinuousListControlDefinitions(formName, fieldSourceName, preferredFields, useExactPreferredFields)
    ' INTENTIONAL KEYED LOOKUP: createdControlNames provides keyed membership/lookup over source rows or state already read by CreateContinuousListForm; it is a derived access path, not a second authority.
    Set createdControlNames = CreateObject("Scripting.Dictionary")
    createdControlNames.CompareMode = vbTextCompare
    Set frm = CreateBlankFormForDesign(formName)
    ApplyBoundFormProperties frm, captionText, recordSourceText, DEFAULT_UI_FORM_WIDTH, DEFAULT_UI_LIST_FORM_HEIGHT, 1
    frm.allowEdits = allowEdits
    frm.allowAdditions = LookupEditorAllowAdditions(formName)
    frm.allowDeletions = LookupEditorAllowDeletions(formName)
    If LookupEditorGeneratedFormIsDefined(formName) Then frm.OnLoad = LOOKUP_EDITOR_RUNTIME_POLICY_EXPRESSION
    EnsureHeaderFooterSections formName
    forms(formName).Section(acHeader).Height = 980
    forms(formName).Section(acDetail).Height = 420
    With SharedFormDesign_AddTitleLabel(formName, acHeader, vbNullString, captionText, DEFAULT_UI_LIST_LEFT_MARGIN, DEFAULT_UI_FORM_TOP_MARGIN, 8500, 360)
        .TextAlign = 1
    End With
    AddGeneratedListActionButtons formName, actionProfile
    leftPos = DEFAULT_UI_LIST_LEFT_MARGIN
    If Len(Trim$(actionProfile)) > 0 Then headerTop = DEFAULT_LIST_ACTION_TOP + DEFAULT_LIST_ACTION_HEIGHT + DEFAULT_UI_BUTTON_ROW_GAP Else headerTop = LIST_HEADER_TOP_FIRST
    detailTop = DEFAULT_UI_LIST_TOP_MARGIN
    headerBottom = 980
    detailBottom = 420
    If includeOpenButton Then
        SetChatGPTBuildContext mChatGPTBuildStep, formName, "Create continuous-list open button", "Left: " & CStr(leftPos) & "; HeaderTop: " & CStr(headerTop) & "; DetailTop: " & CStr(detailTop)
        Set lbl = SharedFormDesign_AddPlainLabel(formName, acHeader, vbNullString, "Open", leftPos, headerTop, LIST_OPEN_COLUMN_WIDTH, DEFAULT_UI_LABEL_HEIGHT, DEFAULT_UI_FONT_SIZE)
        lbl.TextAlign = 1
        SharedFormDesign_AddButton formName, acDetail, vbNullString, "Open", "=ChatGPT_OpenConfigurationDetailFromList(" & SQL_FIELD_ID & ")", leftPos, detailTop, LIST_OPEN_COLUMN_WIDTH, DEFAULT_UI_CONTROL_HEIGHT, "Open this row in the ChatGPT_ detail form."
        If headerTop + DEFAULT_UI_LABEL_HEIGHT + DEFAULT_UI_FORM_SECTION_BOTTOM_MARGIN > headerBottom Then headerBottom = headerTop + DEFAULT_UI_LABEL_HEIGHT + DEFAULT_UI_FORM_SECTION_BOTTOM_MARGIN
        If detailTop + DEFAULT_UI_CONTROL_HEIGHT + DEFAULT_UI_FORM_SECTION_BOTTOM_MARGIN > detailBottom Then detailBottom = detailTop + DEFAULT_UI_CONTROL_HEIGHT + DEFAULT_UI_FORM_SECTION_BOTTOM_MARGIN
        leftPos = leftPos + LIST_OPEN_COLUMN_WIDTH + DEFAULT_UI_LIST_COLUMN_GAP
    End If
    labelWidthValue = DEFAULT_UI_LABEL_WIDTH - DEFAULT_UI_LIST_LABEL_CONTROL_GAP
    For Each Definition In definitions
        SetupControlPump
        fieldName = CStr(Definition(0))
        labelText = CStr(Definition(1))
        widthValue = CLng(Definition(2))
        controlTypeValue = CLng(Definition(3))
        lockedValue = CBool(Definition(4))
        projectedControlName = CStr(Definition(5))
        If controlTypeValue = acCheckBox Then
            If widthValue - 120 < 120 Then widthValue = 240
        End If
        pairWidth = labelWidthValue + DEFAULT_UI_LIST_LABEL_CONTROL_GAP + widthValue
        If leftPos > DEFAULT_UI_LIST_LEFT_MARGIN Then
            If leftPos + pairWidth > DEFAULT_UI_LIST_FORM_RIGHT_EDGE Then
                leftPos = DEFAULT_UI_LIST_LEFT_MARGIN
                detailTop = detailTop + DEFAULT_UI_CONTROL_HEIGHT + DEFAULT_UI_LIST_ROW_GAP
            End If
        End If
        If leftPos + pairWidth > DEFAULT_UI_LIST_FORM_RIGHT_EDGE Then
            maxControlWidth = DEFAULT_UI_LIST_FORM_RIGHT_EDGE - leftPos - labelWidthValue - DEFAULT_UI_LIST_LABEL_CONTROL_GAP
            If maxControlWidth < widthValue Then widthValue = maxControlWidth
            pairWidth = labelWidthValue + DEFAULT_UI_LIST_LABEL_CONTROL_GAP + widthValue
        End If
        controlLeft = leftPos + labelWidthValue + DEFAULT_UI_LIST_LABEL_CONTROL_GAP
        detailText = "Field: " & fieldName & "; Source: " & fieldSourceName
        detailText = detailText & "; Projected control: " & projectedControlName
        detailText = detailText & "; LabelLeft: " & CStr(leftPos) & "; LabelWidth: " & CStr(labelWidthValue)
        detailText = detailText & "; ControlLeft: " & CStr(controlLeft) & "; ControlWidth: " & CStr(widthValue) & "; DetailTop: " & CStr(detailTop)
        detailText = detailText & "; PairWidth: " & CStr(pairWidth) & "; RightEdgeLimit: " & CStr(DEFAULT_UI_LIST_FORM_RIGHT_EDGE)
        SetChatGPTBuildContext mChatGPTBuildStep, formName, "Create continuous-list field controls", detailText
        Set lbl = SharedFormDesign_AddPlainLabel(formName, acDetail, vbNullString, labelText, leftPos, detailTop + 30, labelWidthValue, DEFAULT_UI_LABEL_HEIGHT, DEFAULT_UI_FONT_SIZE)
        lbl.TextAlign = 1
        If controlTypeValue = acCheckBox Then
            Set ctl = SharedFormDesign_AddBoundControl(formName, fieldSourceName, fieldName, acCheckBox, acDetail, vbNullString, controlLeft + 60, detailTop + 30, widthValue - 120, 240, "lst")
        Else
            Set ctl = SharedFormDesign_AddBoundControl(formName, fieldSourceName, fieldName, controlTypeValue, acDetail, vbNullString, controlLeft, detailTop, widthValue, DEFAULT_UI_CONTROL_HEIGHT, "lst")
            If controlTypeValue = acComboBox Then ConfigureComboBox ctl, fieldName, fieldSourceName, Definition(7)
        End If
        If enforceExactControlContract Then EnforceExactBoundControlDefinition ctl, formName, fieldSourceName, Definition, createdControlNames
        If lockedValue Then SharedFormDesign_ApplyLockedControlStyle ctl, DEFAULT_UI_COLOR_LOCKED_BACK
        If detailTop + DEFAULT_UI_CONTROL_HEIGHT + DEFAULT_UI_FORM_SECTION_BOTTOM_MARGIN > detailBottom Then detailBottom = detailTop + DEFAULT_UI_CONTROL_HEIGHT + DEFAULT_UI_FORM_SECTION_BOTTOM_MARGIN
        leftPos = leftPos + pairWidth + DEFAULT_UI_LIST_COLUMN_GAP
    Next Definition
    forms(formName).Section(acHeader).Height = headerBottom
    forms(formName).Section(acDetail).Height = detailBottom
    forms(formName).Width = DEFAULT_UI_FORM_WIDTH
    forms(formName).OnOpen = GeneratedFormContractOnOpenExpression(formName)
    If LookupEditorGeneratedFormIsDefined(formName) Then
        forms(formName).OnDelete = LOOKUP_EDITOR_DELETE_GUARD_EXPRESSION
        forms(formName).OnCurrent = LOOKUP_EDITOR_CODE_FIELD_EDIT_POLICY_EXPRESSION
        forms(formName).AfterInsert = LOOKUP_EDITOR_CODE_FIELD_EDIT_POLICY_EXPRESSION
    ElseIf Len(Trim$(onCurrentExpression)) > 0 Then
        forms(formName).OnCurrent = onCurrentExpression
    End If
    If useExactPreferredFields And AddInReviewChildFormIsDefinedDefault(formName) Then
        SetChatGPTBuildContext "Creating Add-In Review detail forms.", formName, "Validate Add-In Review required bound controls before save.", "Source: " & fieldSourceName & "; Planned bound controls: " & CStr(definitions.count)
        mChatGPTBuildSourceName = fieldSourceName
        ValidateAddInReviewChildControlSources forms(formName), formName
    End If
    SaveAndCloseDesignForm formName
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
    Err.Raise savedErrNumber, "CreateContinuousListForm", "Could not create continuous-list form." & vbCrLf & "Form: " & formName & vbCrLf & "Caption: " & captionText & vbCrLf & "Field source: " & fieldSourceName & vbCrLf & "RecordSource: " & recordSourceText & vbCrLf & "Build detail: " & NullDisplayText(mChatGPTBuildDetail) & vbCrLf & "Error source: " & savedErrSource & vbCrLf & "Error description: " & savedErrDescription
End Sub

' Purpose: Applies the persisted runtime add/edit/delete policy to the current generated lookup-editor form.
' Public role: Form-code/event-expression public entry point. Access form code or an event property can invoke this routine by name; its Public visibility is part of that UI contract.
' Parameters: None.
' Returns: Boolean: True when the current generated lookup editor receives its persisted runtime add/edit/delete policy; False when the callback handles an error and cannot complete it.
Public Function ApplyLookupEditorRuntimePolicy() As Boolean
    On Error GoTo failed
    Dim frm As Form, ctl As control, db As DAO.Database, rs As DAO.Recordset
    Dim p As Variant, policyDefinition As Variant, logicalID As String, expectedForm As String, modeText As String, sourceText As String
    Set frm = Application.CodeContextObject
    LogPhysicalTraceMessageOnly "Info", "LookupEditorRuntime", "ApplyLookupEditorRuntimePolicy", "START", "Form: " & frm.Name, "Apply persisted lookup-editor runtime policy", 0, vbNullString, vbNullString, frm.Name, "Form", "CodeDb", "LookupEditor"
    frm.allowEdits = False
    frm.allowAdditions = False
    frm.allowDeletions = False
    For Each ctl In frm.Controls
        If ctl.controlType = acTextBox Or ctl.controlType = acComboBox Or ctl.controlType = acCheckBox Then ctl.Locked = True
    Next ctl
    For Each p In LookupEditorPolicyDefinitions()
        expectedForm = ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CStr(p(2)))
        If StrComp(frm.Name, expectedForm, vbTextCompare) = 0 Then logicalID = CStr(p(0)): policyDefinition = p: Exit For
    Next p
    If Len(logicalID) = 0 Then Err.Raise vbObjectError + 7458, "ApplyLookupEditorRuntimePolicy", "The runtime lookup policy was invoked outside a canonical lookup editor. Form: " & frm.Name
    Set db = ADCDb()
    Set rs = db.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LookupDefinitions") & "] WHERE " & SQL_FIELD_LOOKUP_LOGICAL_ID & "='" & Replace(logicalID, "'", "''") & "';", dbOpenSnapshot)
    If rs.EOF Then Err.Raise vbObjectError + 7459, "ApplyLookupEditorRuntimePolicy", "The persisted lookup policy row is missing. Lookup: " & logicalID
    modeText = Trim$(Nz(rs.fields(FIELD_EDITOR_MODE), vbNullString))
    If StrComp(modeText, "Editable", vbTextCompare) <> 0 And StrComp(modeText, "ReadOnly", vbTextCompare) <> 0 Then Err.Raise vbObjectError + 7460, "ApplyLookupEditorRuntimePolicy", "The persisted EditorMode is invalid. Lookup: " & logicalID & "; value: " & modeText
    frm.allowEdits = (StrComp(modeText, "Editable", vbTextCompare) = 0)
    frm.allowAdditions = frm.allowEdits And CBool(Nz(rs!allowAdditions, False))
    frm.allowDeletions = frm.allowEdits And CBool(Nz(rs!allowDeletions, False))
    For Each ctl In frm.Controls
        If ctl.controlType = acTextBox Or ctl.controlType = acComboBox Or ctl.controlType = acCheckBox Then
            sourceText = LCase$(Trim$(Nz(ctl.controlSource, vbNullString)))
            If frm.allowEdits = False Then ctl.Locked = True Else ctl.Locked = LookupEditorRuntimeControlLocked(policyDefinition, sourceText, frm.newRecord, rs)
        End If
    Next ctl
    If frm.allowEdits Then frm.Caption = Replace(frm.Caption, " - FINALIZED - READ ONLY", vbNullString) & " - EDITABLE LOOKUP" Else frm.Caption = Replace(frm.Caption, " - EDITABLE LOOKUP", vbNullString) & " - FINALIZED - READ ONLY"
    rs.Close
    LogPhysicalTraceMessageOnly "Info", "LookupEditorRuntime", "ApplyLookupEditorRuntimePolicy", "DONE", "Form: " & frm.Name & vbCrLf & "AllowEdits: " & CStr(frm.allowEdits) & vbCrLf & "AllowAdditions: " & CStr(frm.allowAdditions) & vbCrLf & "AllowDeletions: " & CStr(frm.allowDeletions), "Apply persisted lookup-editor runtime policy", 0, vbNullString, vbNullString, frm.Name, "Form", "CodeDb", "LookupEditor"
    ApplyLookupEditorRuntimePolicy = True
    Exit Function
failed:
    Dim failureText As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim traceFormName As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not frm Is Nothing Then traceFormName = frm.Name
    On Error GoTo 0
    failureText = "ADC could not apply the persisted lookup-editor policy. The editor remains read-only for safety." & vbCrLf & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
    LogPhysicalTraceMessageOnly "Error", "LookupEditorRuntime", "ApplyLookupEditorRuntimePolicy", "ERROR", "Form: " & IIf(Len(traceFormName) = 0, "(unavailable)", traceFormName), "Apply persisted lookup-editor runtime policy", savedErrNumber, savedErrSource, savedErrDescription, traceFormName, "Form", "CodeDb", "LookupEditor"
    On Error Resume Next
    If Not frm Is Nothing Then
        frm.allowEdits = False: frm.allowAdditions = False: frm.allowDeletions = False
        For Each ctl In frm.Controls
            If ctl.controlType = acTextBox Or ctl.controlType = acComboBox Or ctl.controlType = acCheckBox Then ctl.Locked = True
        Next ctl
    End If
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
    CoreShowErrorMessage failureText, vbCritical, "Lookup Editor Policy", "ApplyLookupEditorRuntimePolicy", "UI"
    ApplyLookupEditorRuntimePolicy = False
End Function

Private Function LookupEditorRuntimeControlLocked(ByVal policyDefinition As Variant, ByVal sourceFieldName As String, ByVal newRecord As Boolean, ByVal rs As DAO.Recordset) As Boolean
    Dim codeFieldName As String
    sourceFieldName = LCase$(Trim$(sourceFieldName))
    If Len(sourceFieldName) = 0 Or Left$(sourceFieldName, 1) = "=" Then Exit Function
    codeFieldName = Trim$(CStr(policyDefinition(LOOKUP_EDITOR_DEF_CODE_FIELD)))
    If StrComp(sourceFieldName, codeFieldName, vbTextCompare) = 0 Then
        LookupEditorRuntimeControlLocked = (Not newRecord) And (Not CBool(Nz(rs.fields(FIELD_ALLOW_CODE_EDITS), False)))
    ElseIf sourceFieldName = LCase$(FIELD_LOOKUP_TABLE_NAME) Or sourceFieldName = LCase$(FIELD_LOOKUP_EDITOR_FORM_NAME) Then
        LookupEditorRuntimeControlLocked = True
    ElseIf sourceFieldName = LCase$(FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER) Then
        LookupEditorRuntimeControlLocked = Not CBool(Nz(rs.fields(FIELD_ALLOW_SORT_ORDER_EDITS), True))
    ElseIf sourceFieldName = LCase$(FIELD_LOOKUP_ACTIVE) Or sourceFieldName = LCase$(FIELD_ACTIVE) Then
        LookupEditorRuntimeControlLocked = Not CBool(Nz(rs.fields(FIELD_ALLOW_ACTIVE_EDITS), True))
    ElseIf sourceFieldName = LCase$(FIELD_MANAGED_OBJECT_DEPENDENCY_NOTES) Or InStr(1, sourceFieldName, "description", vbTextCompare) > 0 Then
        LookupEditorRuntimeControlLocked = Not CBool(Nz(rs.fields(FIELD_ALLOW_DESCRIPTION_EDITS), True))
    ElseIf sourceFieldName = LCase$(FIELD_LOOKUP_DISPLAY_NAME) Or Right$(sourceFieldName, 4) = "name" Or Right$(sourceFieldName, 5) = "title" Then
        LookupEditorRuntimeControlLocked = Not CBool(Nz(rs.fields(FIELD_ALLOW_DISPLAY_NAME_EDITS), True))
    Else
        LookupEditorRuntimeControlLocked = False
    End If
End Function

Private Sub CreateSplitPopupForm(ByVal formName As String, ByVal captionText As String, ByVal sourceName As String, ByVal whereText As String, ByVal preferredFields As Variant, ByVal orderFields As Variant, ByVal replaceExisting As Boolean, ByRef builtForms As Collection, ByRef skippedForms As Collection, ByVal requiredForm As Boolean, ByVal descriptionText As String)
    On Error GoTo ErrHandler
    Dim frm As Form
    Dim sqlText As String
    If Len(sourceName) = 0 Then
        If requiredForm Then
            RaiseRequiredFormBuildIssue formName, "Required source was not found for split popup form."
        Else
            AddSkippedFormNote skippedForms, formName, "Source was not found for split popup form."
        End If
        Exit Sub
    End If
    If TableOrQueryExists(sourceName) = False Then
        If requiredForm Then
            RaiseRequiredFormBuildIssue formName, "Required source [" & sourceName & "] was not found for split popup form."
        Else
            AddSkippedFormNote skippedForms, formName, "Source [" & sourceName & "] was not found for split popup form."
        End If
        Exit Sub
    End If
    sqlText = BuildSelectSql(sourceName, preferredFields, whereText, orderFields)
    If Len(sqlText) = 0 Then
        If requiredForm Then
            RaiseRequiredFormBuildIssue formName, "Required source [" & sourceName & "] did not contain any expected split popup fields."
        Else
            AddSkippedFormNote skippedForms, formName, "Source [" & sourceName & "] did not contain any expected split popup fields."
        End If
        Exit Sub
    End If
    If PrepareChatGPTGeneratedFormForBuild(formName, replaceExisting, skippedForms) = False Then Exit Sub
    Set frm = CreateBlankFormForDesign(formName)
    ApplyBoundFormProperties frm, captionText, sqlText, DEFAULT_UI_SPLIT_POPUP_FORM_WIDTH, DEFAULT_UI_SPLIT_POPUP_HEIGHT, DEFAULT_VIEW_SPLIT_FORM
    ApplySplitPopupFormProperties frm
    EnsureHeaderFooterSections formName
    forms(formName).Section(acHeader).Height = DEFAULT_UI_SPLIT_POPUP_HEADER_HEIGHT
    forms(formName).Section(acDetail).Height = DEFAULT_UI_SPLIT_POPUP_DETAIL_HEIGHT
    With SharedFormDesign_AddTitleLabel(formName, acHeader, vbNullString, captionText, DEFAULT_UI_FORM_LEFT_MARGIN, DEFAULT_UI_FORM_TOP_MARGIN, 7000, 330)
        .TextAlign = 1
    End With
    With SharedFormDesign_AddPlainLabel(formName, acHeader, vbNullString, descriptionText, DEFAULT_UI_FORM_LEFT_MARGIN, 470, 9400, 240, DEFAULT_UI_FONT_SIZE)
        .TextAlign = 1
    End With
    SharedFormDesign_AddButton formName, acHeader, vbNullString, "Filter By Selection", "=ChatGPT_FilterCurrentSelection()", 10100, 180, 1500, 330, "Apply Access Filter By Selection to the active datasheet/control."
    SharedFormDesign_AddButton formName, acHeader, vbNullString, "Clear Filter", "=ChatGPT_ClearCurrentFormFilter()", 11720, 180, 1320, 330, "Clear the currently open Access form filter and sort."
    AddFieldBlock formName, sourceName, preferredFields, acDetail, vbNullString, DEFAULT_UI_FORM_LEFT_MARGIN, DEFAULT_UI_FORM_TOP_MARGIN, 2, "spl", False
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
    Err.Raise savedErrNumber, "CreateSplitPopupForm", "Could not create split popup form." & vbCrLf & "Form: " & formName & vbCrLf & "Caption: " & captionText & vbCrLf & "Field source: " & sourceName & vbCrLf & "Where: " & NullDisplayText(whereText) & vbCrLf & "Build detail: " & NullDisplayText(mChatGPTBuildDetail) & vbCrLf & "Error source: " & savedErrSource & vbCrLf & "Error description: " & savedErrDescription
End Sub

Private Sub RaiseRequiredFormBuildIssue(ByVal formName As String, ByVal reasonText As String)
    Err.Raise vbObjectError + 6601, "BuildGeneratedForms", "Required generated form could not be built." & vbCrLf & "Form: " & formName & vbCrLf & "Reason: " & reasonText
End Sub

Private Sub ResetChatGPTVerifiedExistingForms()
    Set mChatGPTVerifiedExistingForms = New Collection
End Sub

Private Function RestoreFormFromBackupByVerifiedSwap(ByVal formName As String, ByVal backupPath As String, ByRef failureText As String) As Boolean
    On Error GoTo ErrHandler
    Dim temporaryFormName As String
    Dim verificationText As String
    Dim originalDeleted As Boolean
    If UIRecoveryBackupReadable(backupPath, failureText) = False Then Exit Function
    temporaryFormName = UniqueRecoveryFormName("zADCRecTmp_")
    Application.LoadFromText acForm, temporaryFormName, backupPath
    If VerifyRestoredForm(temporaryFormName, verificationText) = False Then Err.Raise ERR_FORM_RESTORE_FAILED, "RestoreFormFromBackupByVerifiedSwap", verificationText
    If SharedAccess_ObjectIsOpen(acForm, formName) Then DoCmd.Close acForm, formName, acSaveNo
    If SharedDao_AccessObjectExists(GetHostDb(), acForm, formName) Then
        DoCmd.DeleteObject acForm, formName
        originalDeleted = True
    End If
    DoCmd.Rename formName, acForm, temporaryFormName
    verificationText = vbNullString
    If VerifyRestoredForm(formName, verificationText) = False Then Err.Raise ERR_FORM_RESTORE_FAILED, "RestoreFormFromBackupByVerifiedSwap", verificationText
    RestoreFormFromBackupByVerifiedSwap = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    Dim recoveryText As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    On Error Resume Next
    If SharedDao_AccessObjectExists(GetHostDb(), acForm, temporaryFormName) Then recoveryText = " The verified temporary form remains preserved as [" & temporaryFormName & "]."
    If originalDeleted And SharedDao_AccessObjectExists(GetHostDb(), acForm, formName) = False Then recoveryText = recoveryText & " The original object name is currently absent, but the verified temporary form and source backup remain available for retry."
    On Error GoTo 0
    failureText = "Verified form swap failed. Form: " & formName & "; Backup: " & backupPath & "; Error " & CStr(savedErrNumber) & ": " & savedErrDescription & recoveryText
End Function

' Purpose: Performs the public save and close design form operation using the routine's existing ADC workflow and side effects.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   formName (Input; String): Canonical or physical name identifying the form targeted by this operation.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub SaveAndCloseDesignForm(ByVal formName As String)
    On Error GoTo ErrHandler
    If Len(mSharedPreparedGeneratedFormName) > 0 Then
        If StrComp(mSharedPreparedGeneratedFormName, formName, vbBinaryCompare) <> 0 Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "SaveAndCloseDesignForm", "The generated form being saved does not match the exact form retained by shared replacement recovery." & vbCrLf & "Prepared form: " & mSharedPreparedGeneratedFormName & vbCrLf & "Save target: " & formName
    End If
    SetChatGPTBuildContext mChatGPTBuildStep, formName, "Apply generated-form event metadata before the original design save.", vbNullString
    ApplyGeneratedFormDuplicateReviewCacheInvalidationEventsToOpenDesignForm formName
    SetChatGPTBuildContext mChatGPTBuildStep, formName, "Save and close generated form.", vbNullString
    SharedFormDesign_SaveAndCloseDesignForm formName
    mChatGPTGeneratedFormDesignSaveCloseCount = mChatGPTGeneratedFormDesignSaveCloseCount + 1
    If StrComp(mSharedPreparedGeneratedFormName, formName, vbBinaryCompare) = 0 Then mSharedPreparedGeneratedFormName = vbNullString
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "SaveAndCloseDesignForm", BuildChatGPTChildErrorMessage("Could not save and close generated form.", savedErrSource, savedErrDescription, "Form: " & formName)
End Sub

Private Sub VerifyConfigurationDetailFormLayout(ByVal formName As String, ByVal sourceName As String, ByVal headerFieldList As Variant, ByVal requiredHeaderHeight As Long)
    On Error GoTo ErrHandler
    Dim frm As Access.Form
    Dim tabCtl As Access.control
    Dim ctl As Access.control
    Dim existing As Collection
    Dim fieldName As Variant
    Dim controlName As String
    Dim headerHeight As Long
    Dim issueText As String
    Set frm = forms(formName)
    headerHeight = frm.Section(acHeader).Height
    If headerHeight < requiredHeaderHeight Then issueText = issueText & "Header height is smaller than required. Actual=" & CStr(headerHeight) & "; Required=" & CStr(requiredHeaderHeight) & vbCrLf
    Set tabCtl = frm.Controls("tabConfigurationDetail")
    If tabCtl.pages.count < 4 Then issueText = issueText & "Tab page count is smaller than required. Actual=" & CStr(tabCtl.pages.count) & "; Required=4" & vbCrLf
    Set existing = SharedDao_DatabaseExistingFields(GetHostDb(), sourceName, headerFieldList)
    For Each fieldName In existing
        SetupControlPump
        controlName = SharedFormDesign_SafeControlName("hdr_" & CStr(fieldName))
        Set ctl = Nothing
        On Error Resume Next
        Set ctl = frm.Controls(controlName)
        Err.Clear
        On Error GoTo ErrHandler
        If ctl Is Nothing Then
            issueText = issueText & "Missing fixed-header control: " & controlName & "; Field=" & CStr(fieldName) & vbCrLf
        ElseIf ctl.Top + ctl.Height > headerHeight Then
            issueText = issueText & "Fixed-header control exceeds header section: " & controlName & "; Bottom=" & CStr(ctl.Top + ctl.Height) & "; HeaderHeight=" & CStr(headerHeight) & vbCrLf
        End If
    Next fieldName
    If Len(issueText) > 0 Then Err.Raise vbObjectError + 6730, "VerifyConfigurationDetailFormLayout", "Configuration-detail layout verification failed." & vbCrLf & issueText
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If Len(Trim$(savedErrSource)) = 0 Then savedErrSource = "VerifyConfigurationDetailFormLayout"
    Err.Raise savedErrNumber, savedErrSource, "Could not verify the generated configuration-detail layout." & vbCrLf & "Form: " & formName & vbCrLf & "Source: " & sourceName & vbCrLf & "Required header height: " & CStr(requiredHeaderHeight) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Original description: " & savedErrDescription
End Sub

Private Function VerifyRestoredForm(ByVal formName As String, ByRef failureText As String) As Boolean
    On Error GoTo ErrHandler
    Dim verificationPath As String
    If SharedDao_AccessObjectExists(GetHostDb(), acForm, formName) = False Then Err.Raise vbObjectError + 7216, "VerifyRestoredForm", "The restored form does not exist."
    verificationPath = Environ$("TEMP") & "\ADC_FormReferenceRestoreVerify_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & Format$(CLng(Timer * 1000), "000000000") & ".txt"
    Application.SaveAsText acForm, formName, verificationPath
    If Len(Dir$(verificationPath)) = 0 Then Err.Raise vbObjectError + 7217, "VerifyRestoredForm", "The restored form could not be exported for verification."
    SharedFile_DeleteIfExists verificationPath
    VerifyRestoredForm = True
    Exit Function
ErrHandler:
    failureText = "Restored-form verification failed. Form: " & formName & "; Error " & CStr(Err.Number) & ": " & Err.Description
    On Error Resume Next
    If Len(verificationPath) > 0 Then SharedFile_DeleteIfExists verificationPath
    On Error GoTo 0
End Function

Private Sub AddCoreSystemMacroUtilityActionButtons(ByVal formName As String, ByVal parentName As String, ByVal left1 As Long, ByVal left2 As Long, ByVal left3 As Long, ByRef topPos As Long, Optional ByVal includeHelp As Boolean = False)
    Dim systemMacroDefinitions As Collection
    Dim actionDefinitions As Collection
    Dim actionDefinition As Variant
    Dim actionKey As String
    Dim actionIndex As Long
    Dim actionLeft As Long
    Set systemMacroDefinitions = x__UTIL_ADC_DEFAULTS.GetSystemMacroDefinitions()
    Set actionDefinitions = SystemMacroUtilityActionDefinitions(systemMacroDefinitions)
    For Each actionDefinition In actionDefinitions
        actionKey = CStr(actionDefinition(0))
        actionIndex = actionIndex + 1
        If actionIndex > 1 And ((actionIndex - 1) Mod 3) = 0 Then topPos = topPos + DEFAULT_UI_BUTTON_HEIGHT + DEFAULT_UI_BUTTON_ROW_GAP
        Select Case ((actionIndex - 1) Mod 3) + 1
            Case 1: actionLeft = left1
            Case 2: actionLeft = left2
            Case 3: actionLeft = left3
        End Select
        SharedFormDesign_AddButton formName, acDetail, parentName, SystemMacroUtilityActionCaption(actionKey, actionDefinition), "=RunSystemCommand(""" & actionKey & """)", actionLeft, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, SystemMacroUtilityActionTooltip(actionKey, actionDefinition)
    Next actionDefinition
    If includeHelp Then
        If actionIndex = 0 Then actionLeft = left1 Else actionLeft = Choose((actionIndex Mod 3) + 1, left1, left2, left3)
        SharedFormDesign_AddButton formName, acDetail, parentName, "Help", "=ChatGPT_OpenFormByName(""" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Help") & """)", actionLeft, topPos, DEFAULT_UI_BUTTON_WIDTH, DEFAULT_UI_BUTTON_HEIGHT, "Open searchable " & DEFAULT_UTILITY_LONG_NAME & " Help and the database reference."
    End If
End Sub

Private Sub CreateChatGPTSystemMacrosForm(ByVal replaceExisting As Boolean, ByRef builtForms As Collection, ByRef skippedForms As Collection)
    On Error GoTo ErrHandler
    Dim formName As String
    Dim sourceName As String
    Dim sqlText As String
    Dim fields As Variant
    Dim orderFields As Variant
    formName = ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "System_Macros")
    sourceName = GeneratedFormSourceNameForSignature(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "System_Macros"))
    If Len(sourceName) = 0 Then
        AddSkippedFormNote skippedForms, formName, "Source was not found. " & "Expected " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SystemMacros") & "."
        Exit Sub
    End If
    If TableOrQueryExists(sourceName) = False Then
        AddSkippedFormNote skippedForms, formName, "Source [" & sourceName & "] was not found. " & "Expected " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SystemMacros") & "."
        Exit Sub
    End If
    fields = GeneratedFormListFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "System_Macros"))
    orderFields = GeneratedFormOrderFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "System_Macros"))
    sqlText = BuildSelectSql(sourceName, fields, vbNullString, orderFields)
    If Len(sqlText) = 0 Then
        AddSkippedFormNote skippedForms, formName, "Source [" & sourceName & "] did not contain any expected fields."
        Exit Sub
    End If
    If PrepareChatGPTGeneratedFormForBuild(formName, replaceExisting, skippedForms) = False Then Exit Sub
    CreateContinuousListForm formName, GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "System_Macros")), sqlText, sourceName, fields, GeneratedFormIncludeOpenButton(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "System_Macros")), GeneratedFormListActionProfile(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "System_Macros")), GeneratedFormListAllowEdits(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "System_Macros"))
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
    Err.Raise savedErrNumber, "CreateChatGPTSystemMacrosForm", savedErrDescription & vbCrLf & "Error source: " & savedErrSource
End Sub

Private Function FirstExistingSource(ByVal names As Variant) As String
    Dim item As Variant
    For Each item In names
        SetupControlPump
        If Len(Trim$(CStr(item))) > 0 Then
            If TableOrQueryExists(CStr(item)) Then
                FirstExistingSource = CStr(item)
                Exit Function
            End If
        End If
    Next item
End Function

Private Function GeneratedFormChatGPTUIClassification(ByVal formName As String) As Boolean
    'Returns the authoritative ChatGPTUIForm value for one canonical generated physical form name.
    Dim definitionData As Variant
    Dim canonicalPhysicalName As String
    If Len(formName) = 0 Then Exit Function
    For Each definitionData In CanonicalGeneratedFormNamingDefinitions()
        canonicalPhysicalName = ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, CStr(definitionData(GENERATED_FORM_NAMING_LOGICAL_ID)))
        If StrComp(canonicalPhysicalName, formName, vbTextCompare) = 0 Then
            GeneratedFormChatGPTUIClassification = CBool(definitionData(GENERATED_FORM_NAMING_CHATGPT_UI_FORM))
            Exit Function
        End If
    Next definitionData
End Function

Private Function GeneratedFormCatalogRow(ByVal formName As String, ByRef rowData As TGeneratedFormCatalogRow) As Boolean
    Dim definitionData As Variant
    definitionData = GeneratedFormCatalogDefinition(formName)
    If IsArray(definitionData) = False Then Exit Function
    rowData.formName = CStr(definitionData(0))
    rowData.formDisplayName = CStr(definitionData(1))
    rowData.FormDescription = CStr(definitionData(2))
    rowData.formType = CStr(definitionData(3))
    rowData.NavigationPaneGroupKey = CStr(definitionData(4))
    rowData.ChatGPTUIForm = GeneratedFormChatGPTUIClassification(rowData.formName)
    rowData.SortOrder = CLng(definitionData(5))
    GeneratedFormCatalogRow = True
End Function

Private Sub CreateChatGPTCatalogReviewForm(ByVal replaceExisting As Boolean, ByRef builtForms As Collection, ByRef skippedForms As Collection)
    On Error GoTo ErrHandler
    Dim formName As String
    Dim sourceName As String
    Dim sqlText As String
    Dim fields As Variant
    Dim orderFields As Variant
    formName = ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Catalog_Review")
    sourceName = GeneratedFormSourceNameForSignature(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Catalog_Review"))
    If Len(sourceName) = 0 Then
        AddSkippedFormNote skippedForms, formName, "Source was not found. " & "Expected " & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_ManagedObjects_Review") & "."
        Exit Sub
    End If
    If TableOrQueryExists(sourceName) = False Then
        AddSkippedFormNote skippedForms, formName, "Source [" & sourceName & "] was not found. " & "Expected " & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_ManagedObjects_Review") & "."
        Exit Sub
    End If
    fields = GeneratedFormListFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Catalog_Review"))
    orderFields = GeneratedFormOrderFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Catalog_Review"))
    sqlText = BuildSelectSql(sourceName, fields, vbNullString, orderFields)
    If Len(sqlText) = 0 Then
        AddSkippedFormNote skippedForms, formName, "Source [" & sourceName & "] did not contain any expected fields."
        Exit Sub
    End If
    If PrepareChatGPTGeneratedFormForBuild(formName, replaceExisting, skippedForms) = False Then Exit Sub
    CreateContinuousListForm formName, GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Catalog_Review")), sqlText, sourceName, fields, GeneratedFormIncludeOpenButton(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Catalog_Review")), GeneratedFormListActionProfile(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Catalog_Review")), GeneratedFormListAllowEdits(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Catalog_Review"))
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
    Err.Raise savedErrNumber, "CreateChatGPTCatalogReviewForm", savedErrDescription & vbCrLf & "Error source: " & savedErrSource
End Sub

Private Sub CreateChatGPTFormsCatalogForm(ByVal replaceExisting As Boolean, ByRef builtForms As Collection, ByRef skippedForms As Collection)
    On Error GoTo ErrHandler
    Dim formName As String
    Dim sourceName As String
    Dim sqlText As String
    Dim fields As Variant
    Dim orderFields As Variant
    formName = ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Forms")
    sourceName = GeneratedFormSourceNameForSignature(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Forms"))
    If Len(sourceName) = 0 Then
        AddSkippedFormNote skippedForms, formName, "Source was not found. " & "Expected " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_FormDefinitions") & "."
        Exit Sub
    End If
    If TableOrQueryExists(sourceName) = False Then
        AddSkippedFormNote skippedForms, formName, "Source [" & sourceName & "] was not found. " & "Expected " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_FormDefinitions") & "."
        Exit Sub
    End If
    fields = GeneratedFormListFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Forms"))
    orderFields = GeneratedFormOrderFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Forms"))
    sqlText = BuildSelectSql(sourceName, fields, vbNullString, orderFields)
    If Len(sqlText) = 0 Then
        AddSkippedFormNote skippedForms, formName, "Source [" & sourceName & "] did not contain any expected fields."
        Exit Sub
    End If
    If PrepareChatGPTGeneratedFormForBuild(formName, replaceExisting, skippedForms) = False Then Exit Sub
    CreateContinuousListForm formName, GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Forms")), sqlText, sourceName, fields, GeneratedFormIncludeOpenButton(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Forms")), GeneratedFormListActionProfile(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Forms")), GeneratedFormListAllowEdits(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Forms"))
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
    Err.Raise savedErrNumber, "CreateChatGPTFormsCatalogForm", savedErrDescription & vbCrLf & "Error source: " & savedErrSource
End Sub

' ===============================================================================================================================================================================
' GENERATED-FORM CATALOG REGISTRATION / RECONCILIATION PROCEDURES
' ===============================================================================================================================================================================

Private Function GeneratedFormCatalogRowDiffers(ByVal rs As DAO.Recordset, ByRef rowData As TGeneratedFormCatalogRow, ByVal hasSourceObjectName As Boolean, ByVal hasIsGeneratedByUtility As Boolean, ByVal hasOpenAfterSetup As Boolean, ByVal hasDefaultSource As Boolean, ByVal hasNotes As Boolean) As Boolean
    'Compares authoritative desired-state fields; timestamps are not changes.
    If StrComp(CStr(Nz(rs.fields(FIELD_FORM_DISPLAY_NAME_BASE).value, vbNullString)), rowData.formDisplayName, vbBinaryCompare) <> 0 Then GeneratedFormCatalogRowDiffers = True: Exit Function
    If StrComp(CStr(Nz(rs.fields(FIELD_FORM_DESCRIPTION).value, vbNullString)), rowData.FormDescription, vbBinaryCompare) <> 0 Then GeneratedFormCatalogRowDiffers = True: Exit Function
    If StrComp(CStr(Nz(rs.fields(FIELD_FORM_TYPE).value, vbNullString)), rowData.formType, vbTextCompare) <> 0 Then GeneratedFormCatalogRowDiffers = True: Exit Function
    If StrComp(CStr(Nz(rs.fields(FIELD_NAVIGATION_PANE_GROUP_KEY).value, vbNullString)), rowData.NavigationPaneGroupKey, vbTextCompare) <> 0 Then GeneratedFormCatalogRowDiffers = True: Exit Function
    If StrComp(CStr(Nz(rs.fields(FIELD_CONFIG_OWNER_DATABASE_ROLE).value, vbNullString)), CONFIG_OWNER_DATABASE_ROLE_HOST_DB, vbTextCompare) <> 0 Then GeneratedFormCatalogRowDiffers = True: Exit Function
    If StrComp(CStr(Nz(rs.fields(FIELD_FORM_STATUS).value, vbNullString)), DEFAULT_GENERATED_FORM_STATUS, vbTextCompare) <> 0 Then GeneratedFormCatalogRowDiffers = True: Exit Function
    If StrComp(CStr(Nz(rs.fields(FIELD_CATALOG_OWNER).value, vbNullString)), DEFAULT_CATALOG_OWNER_GENERATED, vbTextCompare) <> 0 Then GeneratedFormCatalogRowDiffers = True: Exit Function
    If IsNull(rs.fields(FIELD_CHAT_GPTUI_FORM).value) Or CBool(Nz(rs.fields(FIELD_CHAT_GPTUI_FORM).value, False)) <> rowData.ChatGPTUIForm Then GeneratedFormCatalogRowDiffers = True: Exit Function
    If CLng(Nz(rs.fields(FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER).value, 0)) <> rowData.SortOrder Then GeneratedFormCatalogRowDiffers = True: Exit Function
    If hasSourceObjectName Then If StrComp(CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_SOURCE_OBJECT_NAME).value, vbNullString)), rowData.formName, vbTextCompare) <> 0 Then GeneratedFormCatalogRowDiffers = True: Exit Function
    If hasIsGeneratedByUtility Then If CBool(Nz(rs.fields(FIELD_IS_GENERATED_BY_UTILITY).value, False)) = False Then GeneratedFormCatalogRowDiffers = True: Exit Function
    If hasOpenAfterSetup Then If CBool(Nz(rs.fields(FIELD_OPEN_AFTER_SETUP).value, False)) <> GeneratedFormOpenAfterSetup(rowData.formName) Then GeneratedFormCatalogRowDiffers = True: Exit Function
    If hasDefaultSource Then If StrComp(CStr(Nz(rs.fields(FIELD_DEFAULT_SOURCE).value, vbNullString)), DEFAULT_UI_BUILDER_SOURCE, vbBinaryCompare) <> 0 Then GeneratedFormCatalogRowDiffers = True: Exit Function
    If hasNotes Then If StrComp(CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_DEPENDENCY_NOTES).value, vbNullString)), DEFAULT_GENERATED_FORM_CATALOG_NOTES, vbBinaryCompare) <> 0 Then GeneratedFormCatalogRowDiffers = True: Exit Function
End Function

' Purpose: Registers generated forms in forms catalog for built forms, nav notes.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   builtForms (Input; Collection): Collection containing the built forms items that the routine consumes or updates as part of this operation.
'   navNotes (Input/output; Collection): Collection supplying the nav notes consumed by this procedure; items are read in the collection's existing order unless the body explicitly keys or sorts them.
'   catalogChangeCount (Output Optional; Variant): Receives the catalog change count produced by this operation.
' Returns: None (Sub). Results are returned through the documented ByRef parameters and/or the routine's database, UI, or module-state side effects.
Public Sub RegisterGeneratedFormsInFormsCatalog(ByVal builtForms As Collection, ByRef navNotes As Collection, Optional ByRef catalogChangeCount As Variant)
    'During numbered Setup D2/D3 own physical generated-form construction only; D4 owns complete managed-form registry reconciliation. Direct/runtime form builds remain allowed outside an active Setup run.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim tableName As String
    Dim item As Variant
    Dim rowData As TGeneratedFormCatalogRow
    Dim registeredCount As Long
    Dim currentFormName As String
    Dim hasIsGeneratedByUtility As Boolean
    Dim changedCount As Long
    Dim expectedCount As Long
    catalogChangeCount = CLng(0)
    If Not builtForms Is Nothing Then expectedCount = builtForms.count
    If SetupDomainMutationWritesAllowedForCurrentPart(SETUP_DOMAIN_D4_MANAGED_FORM_LOGICAL_STATE) = False Then Err.Raise vbObjectError + 7871, "RegisterGeneratedFormsInFormsCatalog", "Form Definitions metadata mutation during numbered Setup is owned exclusively by D4; D2/D3 own physical generated-form construction only."
    If builtForms Is Nothing Then Exit Sub
    If builtForms.count = 0 Then Exit Sub
    Set db = GetHostDb()
    tableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_FormDefinitions")
    If SharedDao_TableExists(db, tableName) = False Then Err.Raise vbObjectError + 6685, "RegisterGeneratedFormsInFormsCatalog", "Required generated-form catalog table was not found." & vbCrLf & "Table: " & tableName
    If RequiredFormCatalogFieldsExist(tableName) = False Then Err.Raise vbObjectError + 6686, "RegisterGeneratedFormsInFormsCatalog", "The generated-form catalog table is missing one or more required fields." & vbCrLf & "Table: " & tableName
    hasIsGeneratedByUtility = SharedDao_TableFieldExists(db, tableName, FIELD_IS_GENERATED_BY_UTILITY)
    For Each item In builtForms
        SetupControlPump
        currentFormName = CStr(item)
        mChatGPTBuildForm = currentFormName
        mChatGPTBuildAction = "Register generated form catalog row."
        mChatGPTBuildDetail = "Registering form " & CStr(registeredCount + 1) & " of " & CStr(builtForms.count) & ": " & currentFormName
        If registeredCount = 0 Or (registeredCount + 1) Mod 5 = 0 Or registeredCount + 1 = builtForms.count Then UpdateChatGPTProgress mChatGPTBuildDetail
        If GeneratedFormCatalogRow(currentFormName, rowData) = False Then Err.Raise vbObjectError + 6687, "RegisterGeneratedFormsInFormsCatalog", "No authoritative generated-form catalog definition exists." & vbCrLf & "Form: " & currentFormName
        If UpsertGeneratedFormCatalogRow(db, tableName, rowData) Then changedCount = changedCount + 1
        VerifyGeneratedFormCatalogRow db, tableName, rowData.formName, hasIsGeneratedByUtility
        registeredCount = registeredCount + 1
        RequireSafeWorkContinue "RegisterGeneratedFormsInFormsCatalog", "Form [" & currentFormName & "]"
    Next item
    If registeredCount <> builtForms.count Then Err.Raise vbObjectError + 6688, "RegisterGeneratedFormsInFormsCatalog", "Not every authoritative generated form was registered and verified in the catalog." & vbCrLf & "Expected: " & CStr(builtForms.count) & vbCrLf & "Verified: " & CStr(registeredCount)
    mChatGPTBuildForm = vbNullString
    mChatGPTBuildAction = "Generated form catalog registration complete."
    mChatGPTBuildDetail = "Registered and verified " & CStr(registeredCount) & " authoritative generated forms."
    catalogChangeCount = changedCount
    navNotes.Add "Registered and verified all " & CStr(registeredCount) & " authoritative generated form rows in [" & tableName & "] with their form-specific NavigationPaneGroupKey values. Catalog rows changed: " & CStr(changedCount) & "."
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    catalogChangeCount = changedCount
    If IsCancellationError(savedErrNumber) Then Err.Raise savedErrNumber, savedErrSource, savedErrDescription
    If Not navNotes Is Nothing Then navNotes.Add "Generated form catalog registration failed for [" & IIf(Len(currentFormName) = 0, "(not set)", currentFormName) & "]. Error " & CStr(savedErrNumber) & ": " & savedErrDescription
    Err.Raise savedErrNumber, "RegisterGeneratedFormsInFormsCatalog", "Could not register and verify all authoritative generated forms in the required catalog." & vbCrLf & "Form: " & IIf(Len(currentFormName) = 0, "(not set)", currentFormName) & vbCrLf & "Registered before failure: " & CStr(registeredCount) & " of " & CStr(expectedCount) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

' Purpose: Deletes noncurrent generated form definitions for DB, canonical generated names.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   db (Input; DAO.Database): DAO.Database instance whose schema, data, queries, or properties the routine reads or modifies.
'   canonicalGeneratedNames (Input; Collection): Collection containing the canonical generated names items that the routine consumes or updates as part of this operation.
' Returns: Long: Returns the number of delete noncurrent generated form definitions items matched, processed, selected, deleted, or reviewed by this operation; zero represents no qualifying items.
Public Function DeleteNoncurrentGeneratedFormDefinitions(ByVal db As DAO.Database, ByVal canonicalGeneratedNames As Collection) As Long
    ' The current generated-form collection is authoritative. Delete only proven utility-generated UI_Frm_Registry metadata that is no longer current; never delete physical forms here. D4 owns the surrounding
    ' complete transaction/reconciliation, and this path must not gain a hardcoded historical generated-form list.
    On Error GoTo ErrHandler
    If SetupDomainMutationWritesAllowedForCurrentPart(SETUP_DOMAIN_D4_MANAGED_FORM_LOGICAL_STATE) = False Then Err.Raise vbObjectError + 7872, "DeleteNoncurrentGeneratedFormDefinitions", "Form Definitions metadata mutation during numbered Setup is owned exclusively by D4; D2/D3 own physical generated-form construction only."
    Dim tableName As String
    Dim canonicalFormNameList As String
    If db Is Nothing Then Err.Raise vbObjectError + 6768, "DeleteNoncurrentGeneratedFormDefinitions", "The database context is missing."
    If canonicalGeneratedNames Is Nothing Then Err.Raise vbObjectError + 6769, "DeleteNoncurrentGeneratedFormDefinitions", "The canonical generated-form collection is missing."
    If canonicalGeneratedNames.count = 0 Then Err.Raise vbObjectError + 6770, "DeleteNoncurrentGeneratedFormDefinitions", "The canonical generated-form collection is empty; noncurrent generated UI_Frm_Registry cannot be deleted safely."
    tableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_FormDefinitions")
    If SharedDao_TableExists(db, tableName) = False Then Err.Raise vbObjectError + 6771, "DeleteNoncurrentGeneratedFormDefinitions", "The generated-form catalog table is missing." & vbCrLf & "Table: " & tableName
    If SharedDao_TableFieldExists(db, tableName, FIELD_CATALOG_OWNER) = False Or SharedDao_TableFieldExists(db, tableName, FIELD_IS_GENERATED_BY_UTILITY) = False Then Err.Raise vbObjectError + 6772, "DeleteNoncurrentGeneratedFormDefinitions", "Generated-form ownership fields are missing from UI_Frm_Registry." & vbCrLf & "Table: " & tableName
    canonicalFormNameList = ChatGPTFormNameCollectionSqlList(canonicalGeneratedNames, "DeleteNoncurrentGeneratedFormDefinitions")
    db.Execute "DELETE FROM [" & tableName & "] WHERE Nz([" & FIELD_CATALOG_OWNER & "],'')<>" & SharedDao_SqlQuote(DEFAULT_CATALOG_OWNER_PERSISTENT_GENERATED) & " AND (Nz([" & FIELD_CATALOG_OWNER & "],'')=" & SharedDao_SqlQuote(DEFAULT_CATALOG_OWNER_GENERATED) & " OR Nz(" & SQL_FIELD_IS_GENERATED_BY_UTILITY & ",False)=True) AND Nz(" & SQL_FIELD_FORM_OBJECT_NAME & ",'') NOT IN (" & canonicalFormNameList & ");", dbFailOnError
    DeleteNoncurrentGeneratedFormDefinitions = db.RecordsAffected
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "DeleteNoncurrentGeneratedFormDefinitions", "Could not delete noncurrent utility-generated UI_Frm_Registry metadata." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

' Purpose: Deletes only stale ADC-discovered user UI_Frm_Registry metadata. Physical forms are never deleted here; the authoritative exact-prefix predicate owns membership.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   db (Input; DAO.Database): DAO.Database instance whose schema, data, queries, or properties the routine reads or modifies.
'   authoritativeUserNames (Input; Collection): Collection containing the authoritative user names items that the routine consumes or updates as part of this operation.
' Returns: Long: Returns the number of delete noncurrent discovered user form definitions items matched, processed, selected, deleted, or reviewed by this operation; zero represents no qualifying items.
Public Function DeleteNoncurrentDiscoveredUserFormDefinitions(ByVal db As DAO.Database, ByVal authoritativeUserNames As Collection) As Long
    'Deletes only stale ADC-discovered user UI_Frm_Registry metadata. Physical forms are never deleted here; the authoritative exact-prefix predicate owns membership.
    On Error GoTo ErrHandler
    If SetupDomainMutationWritesAllowedForCurrentPart(SETUP_DOMAIN_D4_MANAGED_FORM_LOGICAL_STATE) = False Then Err.Raise vbObjectError + 7873, "DeleteNoncurrentDiscoveredUserFormDefinitions", "Form Definitions metadata mutation during numbered Setup is owned exclusively by D4; D2/D3 own physical generated-form construction only."
    Dim tableName As String
    Dim rs As DAO.Recordset
    Dim keepNames As Object
    Dim item As Variant
    Dim formName As String
    Dim deletedCount As Long
    If db Is Nothing Then Err.Raise vbObjectError + 6773, "DeleteNoncurrentDiscoveredUserFormDefinitions", "The database context is missing."
    If authoritativeUserNames Is Nothing Then Err.Raise vbObjectError + 6774, "DeleteNoncurrentDiscoveredUserFormDefinitions", "The authoritative discovered-user collection is missing."
    tableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_FormDefinitions")
    If SharedDao_TableExists(db, tableName) = False Then Exit Function
    ' INTENTIONAL KEYED LOOKUP: keepNames is an exact-name index over the authoritative discovered-user collection for this one cleanup pass; it is not an independent membership authority.
    Set keepNames = CreateObject("Scripting.Dictionary")
    keepNames.CompareMode = vbBinaryCompare
    For Each item In authoritativeUserNames
        formName = CStr(item)
        If keepNames.Exists(formName) = False Then keepNames.Add formName, True
    Next item
    Set rs = db.OpenRecordset("SELECT " & SQL_FIELD_ID & "," & SQL_FIELD_FORM_OBJECT_NAME & " FROM [" & tableName & "] WHERE Nz([" & FIELD_CATALOG_OWNER & "],'')=" & SharedDao_SqlQuote(DEFAULT_CATALOG_OWNER_USER) & " ORDER BY " & SQL_FIELD_ID & ";", dbOpenDynaset)
    Do While rs.EOF = False
        formName = CStr(Nz(rs.fields(FIELD_FORM_OBJECT_NAME).value, vbNullString))
        If PhysicalFormIsDiscoveryCandidate(formName) Then
            If keepNames.Exists(formName) = False Then
                rs.Delete
                deletedCount = deletedCount + 1
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    DeleteNoncurrentDiscoveredUserFormDefinitions = deletedCount
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
    Err.Raise savedErrNumber, "DeleteNoncurrentDiscoveredUserFormDefinitions", "Could not delete stale discovered-user UI_Frm_Registry metadata." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Sub RequireChatGPTGeneratedFormCatalogRollbackMatches(ByVal db As DAO.Database, ByVal physicalFormNames As Collection, ByVal rollbackTableName As String)
    Dim backupRs As DAO.Recordset
    Dim currentRs As DAO.Recordset
    Dim backupField As DAO.Field
    Dim currentField As DAO.Field
    Dim currentFieldNames As Object
    Dim backupCount As Long
    Dim currentCount As Long
    Dim physicalFormNameSqlList As String
    Dim currentTableName As String
    currentTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_FormDefinitions")
    If SharedDao_TableExists(db, currentTableName) = False Then Err.Raise vbObjectError + 6728, "RequireChatGPTGeneratedFormCatalogRollbackMatches", "The generated-form catalog table is missing during rollback verification."
    If SharedDao_TableExists(db, rollbackTableName) = False Then Err.Raise vbObjectError + 6729, "RequireChatGPTGeneratedFormCatalogRollbackMatches", "The generated-form catalog rollback table is missing during rollback verification." & vbCrLf & "Table: " & rollbackTableName
    physicalFormNameSqlList = ChatGPTFormNameCollectionSqlList(physicalFormNames, "RequireChatGPTGeneratedFormCatalogRollbackMatches")
    backupCount = ChatGPTRecordsetCount(db, "SELECT Count(*) AS RowCountValue FROM [" & rollbackTableName & "];")
    currentCount = ChatGPTRecordsetCount(db, "SELECT Count(*) AS RowCountValue FROM [" & currentTableName & "] WHERE " & SQL_FIELD_FORM_OBJECT_NAME & " IN (" & physicalFormNameSqlList & ");")
    If currentCount <> backupCount Then Err.Raise vbObjectError + 6730, "RequireChatGPTGeneratedFormCatalogRollbackMatches", "Generated-form catalog rollback row counts do not match." & vbCrLf & "Current generated rows: " & CStr(currentCount) & vbCrLf & "Rollback rows: " & CStr(backupCount)
    Set backupRs = db.OpenRecordset("SELECT * FROM [" & rollbackTableName & "] ORDER BY " & SQL_FIELD_ID & ";", dbOpenSnapshot)
    Set currentRs = db.OpenRecordset("SELECT * FROM [" & currentTableName & "] WHERE " & SQL_FIELD_FORM_OBJECT_NAME & " IN (" & physicalFormNameSqlList & ") ORDER BY " & SQL_FIELD_ID & ";", dbOpenSnapshot)
    Set currentFieldNames = CreateObject("Scripting.Dictionary")
    currentFieldNames.CompareMode = vbTextCompare
    For Each currentField In currentRs.fields
        If currentFieldNames.Exists(currentField.Name) = False Then currentFieldNames.Add currentField.Name, True
    Next currentField
    Do While Not backupRs.EOF
        SetupControlPump
        If currentRs.EOF Then Err.Raise vbObjectError + 6731, "RequireChatGPTGeneratedFormCatalogRollbackMatches", "A generated-form catalog rollback ID is missing after restoration." & vbCrLf & "ID: " & CStr(backupRs.fields(FIELD_ID))
        If CLng(currentRs.fields(FIELD_ID)) <> CLng(backupRs.fields(FIELD_ID)) Then Err.Raise vbObjectError + 6731, "RequireChatGPTGeneratedFormCatalogRollbackMatches", "A generated-form catalog rollback ID is missing or out of sequence after restoration." & vbCrLf & "Expected ID: " & CStr(backupRs.fields(FIELD_ID)) & vbCrLf & "Current ID: " & CStr(currentRs.fields(FIELD_ID))
        For Each backupField In backupRs.fields
            SetupControlPump
            If currentFieldNames.Exists(backupField.Name) = False Then Err.Raise vbObjectError + 6732, "RequireChatGPTGeneratedFormCatalogRollbackMatches", "A generated-form catalog rollback field is missing after restoration." & vbCrLf & "Field: " & backupField.Name
            If ChatGPTVariantValuesEqual(currentRs.fields(backupField.Name).value, backupField.value) = False Then Err.Raise vbObjectError + 6733, "RequireChatGPTGeneratedFormCatalogRollbackMatches", "A generated-form catalog value does not match its rollback snapshot." & vbCrLf & "ID: " & CStr(backupRs.fields(FIELD_ID)) & vbCrLf & "Field: " & backupField.Name
        Next backupField
        currentRs.MoveNext
        backupRs.MoveNext
    Loop
    If currentRs.EOF = False Then Err.Raise vbObjectError + 6730, "RequireChatGPTGeneratedFormCatalogRollbackMatches", "Generated-form catalog rollback verification found extra restored rows after the ordered batch comparison."
    currentRs.Close
    backupRs.Close
End Sub

Private Function RequiredFormCatalogFieldsExist(ByVal tableName As String) As Boolean
    RequiredFormCatalogFieldsExist = _
        SharedDao_TableFieldExists(GetHostDb(), tableName, FIELD_FORM_OBJECT_NAME) And _
        SharedDao_TableFieldExists(GetHostDb(), tableName, FIELD_FORM_DISPLAY_NAME_BASE_PREFIX) And _
        SharedDao_TableFieldExists(GetHostDb(), tableName, FIELD_FORM_DISPLAY_NAME_BASE) And _
        SharedDao_TableFieldExists(GetHostDb(), tableName, FIELD_FORM_DISPLAY_NAME_BASE_POSTFIX) And _
        SharedDao_TableFieldExists(GetHostDb(), tableName, FIELD_FORM_DISPLAY_NAME_CALC) And _
        SharedDao_TableFieldExists(GetHostDb(), tableName, FIELD_CHAT_GPTUI_FORM) And _
        SharedDao_TableFieldExists(GetHostDb(), tableName, FIELD_FORM_DESCRIPTION) And _
        SharedDao_TableFieldExists(GetHostDb(), tableName, FIELD_FORM_TYPE) And _
        SharedDao_TableFieldExists(GetHostDb(), tableName, FIELD_NAVIGATION_PANE_GROUP_KEY) And _
        SharedDao_TableFieldExists(GetHostDb(), tableName, FIELD_CONFIG_OWNER_DATABASE_ROLE) And _
        SharedDao_TableFieldExists(GetHostDb(), tableName, FIELD_FORM_STATUS) And _
        SharedDao_TableFieldExists(GetHostDb(), tableName, FIELD_CATALOG_OWNER) And _
        SharedDao_TableFieldExists(GetHostDb(), tableName, FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER)
End Function

Private Function RestoreChatGPTGeneratedFormCatalogRowsTransactional(ByVal physicalFormNames As Collection, ByVal formCatalogRollbackTable As String, ByRef resultText As String, ByRef unresolvedWorkspace As DAO.workspace) As Boolean
    On Error GoTo ErrHandler
    Dim ws As DAO.workspace
    Dim db As DAO.Database
    Dim destinationTableDef As DAO.tableDef
    Dim destinationField As DAO.Field
    Dim transactionStarted As Boolean
    Dim formsCatalogExists As Boolean
    Dim currentGeneratedRowCount As Long
    Dim physicalFormNameSqlList As String
    Dim destinationTableName As String
    Dim restoreFieldSqlList As String
    If Len(Trim$(formCatalogRollbackTable)) = 0 Then resultText = "No generated form catalog rollback table was required.": RestoreChatGPTGeneratedFormCatalogRowsTransactional = True: Exit Function
    physicalFormNameSqlList = ChatGPTFormNameCollectionSqlList(physicalFormNames, "RestoreChatGPTGeneratedFormCatalogRowsTransactional")
    destinationTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_FormDefinitions")
    Set db = GetHostDb()
    db.TableDefs.Refresh
    If SharedDao_TableExists(db, formCatalogRollbackTable) = False Then
        formsCatalogExists = SharedDao_TableExists(db, destinationTableName)
        If formsCatalogExists Then currentGeneratedRowCount = ChatGPTRecordsetCount(db, "SELECT Count(*) AS RowCountValue FROM [" & destinationTableName & "] WHERE " & SQL_FIELD_FORM_OBJECT_NAME & " IN (" & physicalFormNameSqlList & ");")
        Err.Raise ERR_GENERATED_CATALOG_ROLLBACK_FAILED, "RestoreChatGPTGeneratedFormCatalogRowsTransactional", "The generated form catalog rollback table is missing after a TableDefs refresh." & vbCrLf & "Database: " & db.Name & vbCrLf & "Rollback table: " & formCatalogRollbackTable & vbCrLf & "TableDefs count: " & CStr(db.TableDefs.count) & vbCrLf & "Forms catalog exists: " & CStr(formsCatalogExists) & vbCrLf & "Current generated-form catalog rows: " & CStr(currentGeneratedRowCount) & vbCrLf & "Recovery remains blocked because catalog restoration cannot be verified without its snapshot."
    End If
    Set destinationTableDef = db.TableDefs(destinationTableName)
    destinationTableDef.fields.Refresh
    For Each destinationField In destinationTableDef.fields
        SetupControlPump
        If StrComp(destinationField.Name, FIELD_FORM_DISPLAY_NAME_CALC, vbTextCompare) <> 0 Then
            If SharedDao_DatabaseTableFieldExistsSafe(db, formCatalogRollbackTable, destinationField.Name) = False Then Err.Raise ERR_GENERATED_CATALOG_ROLLBACK_FAILED, "RestoreChatGPTGeneratedFormCatalogRowsTransactional", "The generated form catalog rollback table is missing a writable destination field." & vbCrLf & "Rollback table: " & formCatalogRollbackTable & vbCrLf & "Field: " & destinationField.Name
            If Len(restoreFieldSqlList) > 0 Then restoreFieldSqlList = restoreFieldSqlList & ","
            restoreFieldSqlList = restoreFieldSqlList & SharedDao_BracketName(destinationField.Name)
        End If
    Next destinationField
    If Len(restoreFieldSqlList) = 0 Then Err.Raise ERR_GENERATED_CATALOG_ROLLBACK_FAILED, "RestoreChatGPTGeneratedFormCatalogRowsTransactional", "The generated form catalog has no writable fields available for rollback restoration."
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    transactionStarted = True
    db.Execute "DELETE FROM [" & destinationTableName & "] WHERE " & SQL_FIELD_FORM_OBJECT_NAME & " IN (" & physicalFormNameSqlList & ");", dbFailOnError
    db.Execute "INSERT INTO [" & destinationTableName & "] (" & restoreFieldSqlList & ") SELECT " & restoreFieldSqlList & " FROM [" & formCatalogRollbackTable & "];", dbFailOnError
    RequireChatGPTGeneratedFormCatalogRollbackMatches db, physicalFormNames, formCatalogRollbackTable
    ws.CommitTrans
    transactionStarted = False
    resultText = "Generated form catalog rows restored transactionally and verified."
    RestoreChatGPTGeneratedFormCatalogRowsTransactional = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    Dim rollbackErrNumber As Long
    Dim rollbackErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    If transactionStarted Then
        On Error Resume Next
        Err.Clear
        ws.Rollback
        rollbackErrNumber = Err.Number
        rollbackErrDescription = Err.Description
        On Error GoTo 0
        If rollbackErrNumber <> 0 Then
            Set unresolvedWorkspace = ws
            resultText = "Generated form catalog restoration failed, and its transaction could not be rolled back. Original error " & CStr(savedErrNumber) & ": " & savedErrDescription & "; rollback error " & CStr(rollbackErrNumber) & ": " & rollbackErrDescription
            Exit Function
        End If
    End If
    resultText = "Generated form catalog restoration failed and was rolled back. Error " & CStr(savedErrNumber) & ": " & savedErrDescription
End Function

Private Function UpsertGeneratedFormCatalogRow(ByVal db As DAO.Database, ByVal tableName As String, ByRef rowData As TGeneratedFormCatalogRow) As Boolean
    'Updates only rows whose authoritative catalog values differ.
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim isNewRow As Boolean
    Dim rowChanged As Boolean
    Dim hasDefaultedOn As Boolean
    Dim hasLastDefaultRepairOn As Boolean
    Dim hasSourceObjectName As Boolean
    Dim hasIsGeneratedByUtility As Boolean
    Dim hasOpenAfterSetup As Boolean
    Dim hasLastVerifiedOn As Boolean
    Dim hasDefaultSource As Boolean
    Dim hasNotes As Boolean
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    If db Is Nothing Then Err.Raise vbObjectError + 6696, "UpsertGeneratedFormCatalogRow", "The generated-form catalog update database is not set."
    hasDefaultedOn = SharedDao_TableFieldExists(db, tableName, FIELD_DEFAULTED_ON)
    hasLastDefaultRepairOn = SharedDao_TableFieldExists(db, tableName, FIELD_LAST_DEFAULT_REPAIR_ON)
    hasSourceObjectName = SharedDao_TableFieldExists(db, tableName, FIELD_MANAGED_OBJECT_FIELD_SOURCE_OBJECT_NAME)
    hasIsGeneratedByUtility = SharedDao_TableFieldExists(db, tableName, FIELD_IS_GENERATED_BY_UTILITY)
    hasOpenAfterSetup = SharedDao_TableFieldExists(db, tableName, FIELD_OPEN_AFTER_SETUP)
    hasLastVerifiedOn = SharedDao_TableFieldExists(db, tableName, FIELD_LAST_VERIFIED_ON)
    hasDefaultSource = SharedDao_TableFieldExists(db, tableName, FIELD_DEFAULT_SOURCE)
    hasNotes = SharedDao_TableFieldExists(db, tableName, FIELD_MANAGED_OBJECT_DEPENDENCY_NOTES)
    sqlText = "SELECT * FROM " & SharedDao_BracketName(tableName) & " WHERE " & SQL_FIELD_FORM_OBJECT_NAME & "=" & SharedDao_SqlQuote(rowData.formName) & ";"
    Set rs = db.OpenRecordset(sqlText, dbOpenDynaset)
    isNewRow = rs.EOF
    If isNewRow Then
        rs.AddNew
        rs.fields(FIELD_FORM_OBJECT_NAME).value = rowData.formName
        If hasDefaultedOn Then rs.fields(FIELD_DEFAULTED_ON).value = Now
        rowChanged = True
    Else
        rowChanged = GeneratedFormCatalogRowDiffers(rs, rowData, hasSourceObjectName, hasIsGeneratedByUtility, hasOpenAfterSetup, hasDefaultSource, hasNotes)
        If rowChanged = False Then
            rs.Close
            Set rs = Nothing
            Exit Function
        End If
        rs.Edit
        If hasLastDefaultRepairOn Then rs.fields(FIELD_LAST_DEFAULT_REPAIR_ON).value = Now
    End If
    rs.fields(FIELD_FORM_DISPLAY_NAME_BASE).value = rowData.formDisplayName
    rs.fields(FIELD_FORM_DESCRIPTION).value = rowData.FormDescription
    rs.fields(FIELD_FORM_TYPE).value = rowData.formType
    rs.fields(FIELD_NAVIGATION_PANE_GROUP_KEY).value = rowData.NavigationPaneGroupKey
    rs.fields(FIELD_CONFIG_OWNER_DATABASE_ROLE).value = CONFIG_OWNER_DATABASE_ROLE_HOST_DB
    SetRecordsetFieldValueWithSource rs, FIELD_FORM_STATUS, DEFAULT_GENERATED_FORM_STATUS, "DEFAULT_GENERATED_FORM_STATUS"
    SetRecordsetFieldValueWithSource rs, FIELD_CATALOG_OWNER, DEFAULT_CATALOG_OWNER_GENERATED, "DEFAULT_CATALOG_OWNER_GENERATED"
    rs.fields(FIELD_CHAT_GPTUI_FORM).value = rowData.ChatGPTUIForm
    rs.fields(FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER).value = rowData.SortOrder
    If hasSourceObjectName Then rs.fields(FIELD_MANAGED_OBJECT_FIELD_SOURCE_OBJECT_NAME).value = rowData.formName
    If hasIsGeneratedByUtility Then rs.fields(FIELD_IS_GENERATED_BY_UTILITY).value = True
    If hasOpenAfterSetup Then SetRecordsetFieldValueWithSource rs, FIELD_OPEN_AFTER_SETUP, GeneratedFormOpenAfterSetup(rowData.formName), "GeneratedFormOpenAfterSetup"
    If hasLastVerifiedOn Then rs.fields(FIELD_LAST_VERIFIED_ON).value = Now
    If hasDefaultSource Then SetRecordsetFieldValueWithSource rs, FIELD_DEFAULT_SOURCE, DEFAULT_UI_BUILDER_SOURCE, "DEFAULT_UI_BUILDER_SOURCE"
    If hasNotes Then SetRecordsetFieldValueWithSource rs, FIELD_MANAGED_OBJECT_DEPENDENCY_NOTES, DEFAULT_GENERATED_FORM_CATALOG_NOTES, "DEFAULT_GENERATED_FORM_CATALOG_NOTES"
    rs.Update
    UpsertGeneratedFormCatalogRow = True
    rs.Close
    Set rs = Nothing
    Exit Function
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
    Err.Raise savedErrNumber, "UpsertGeneratedFormCatalogRow", "Could not update the generated-form catalog row." & vbCrLf & "Form: " & rowData.formName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Sub VerifyGeneratedFormCatalogRow(ByVal db As DAO.Database, ByVal tableName As String, ByVal formName As String, Optional ByVal requireGeneratedClassification As Boolean = False)
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim rowData As TGeneratedFormCatalogRow
    Dim sqlText As String
    Dim generatedClassificationSelect As String
    If db Is Nothing Then Err.Raise vbObjectError + 6697, "VerifyGeneratedFormCatalogRow", "The generated-form catalog verification database is not set."
    If GeneratedFormCatalogRow(formName, rowData) = False Then Err.Raise vbObjectError + 6698, "VerifyGeneratedFormCatalogRow", "The authoritative generated-form definition was not found."
    If requireGeneratedClassification Then
        If SharedDao_TableFieldExists(db, tableName, FIELD_IS_GENERATED_BY_UTILITY) = False Then Err.Raise vbObjectError + 6766, "VerifyGeneratedFormCatalogRow", "IsGeneratedByUtility is required for exact same-run D2 proof publication."
        generatedClassificationSelect = ", " & SQL_FIELD_IS_GENERATED_BY_UTILITY
    End If
    sqlText = "SELECT " & SQL_FIELD_FORM_OBJECT_NAME & ", [" & FIELD_NAVIGATION_PANE_GROUP_KEY & "], [" & FIELD_CONFIG_OWNER_DATABASE_ROLE & "], " & SQL_FIELD_FORM_STATUS & ", [" & FIELD_CATALOG_OWNER & "], " & SQL_FIELD_CHAT_GPTUI_FORM & generatedClassificationSelect & " FROM " & SharedDao_BracketName(tableName) & " WHERE " & SQL_FIELD_FORM_OBJECT_NAME & "=" & SharedDao_SqlQuote(formName) & ";"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    If rs.EOF Then Err.Raise vbObjectError + 6689, "VerifyGeneratedFormCatalogRow", "The generated-form catalog row was not found after upsert."
    If StrComp(Trim$(CStr(Nz(rs.fields(FIELD_NAVIGATION_PANE_GROUP_KEY).value, vbNullString))), rowData.NavigationPaneGroupKey, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 6690, "VerifyGeneratedFormCatalogRow", "NavigationPaneGroupKey does not match the authoritative generated-form definition." & vbCrLf & "Expected: " & rowData.NavigationPaneGroupKey
    If StrComp(Trim$(CStr(Nz(rs.fields(FIELD_CONFIG_OWNER_DATABASE_ROLE).value, vbNullString))), CONFIG_OWNER_DATABASE_ROLE_HOST_DB, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 6691, "VerifyGeneratedFormCatalogRow", "ConfigOwnerDatabaseRole is not HostDb."
    If StrComp(Trim$(CStr(Nz(rs.fields(FIELD_FORM_STATUS).value, vbNullString))), ADCAT_STATUS_ACTIVE, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 6692, "VerifyGeneratedFormCatalogRow", "FormStatus is not Active."
    If StrComp(Trim$(CStr(Nz(rs.fields(FIELD_CATALOG_OWNER).value, vbNullString))), DEFAULT_CATALOG_OWNER_GENERATED, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 6693, "VerifyGeneratedFormCatalogRow", "CatalogOwner is not Generated."
    If IsNull(rs.fields(FIELD_CHAT_GPTUI_FORM).value) Or CBool(Nz(rs.fields(FIELD_CHAT_GPTUI_FORM).value, False)) <> rowData.ChatGPTUIForm Then Err.Raise vbObjectError + 7870, "VerifyGeneratedFormCatalogRow", "ChatGPTUIForm does not match the authoritative generated-form classification or is Null." & vbCrLf & "Expected explicit Boolean: " & CStr(rowData.ChatGPTUIForm)
    If requireGeneratedClassification Then If CBool(Nz(rs.fields(FIELD_IS_GENERATED_BY_UTILITY).value, False)) = False Then Err.Raise vbObjectError + 6767, "VerifyGeneratedFormCatalogRow", "IsGeneratedByUtility is not True."
    rs.MoveNext
    If Not rs.EOF Then Err.Raise vbObjectError + 6694, "VerifyGeneratedFormCatalogRow", "More than one catalog row exists for the generated form."
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
    Err.Raise savedErrNumber, "VerifyGeneratedFormCatalogRow", "Could not verify durable generated-form catalog metadata." & vbCrLf & "Table: " & tableName & vbCrLf & "Form: " & formName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

' ===============================================================================================================================================================================
' SHARED CONFIGURATION / SPECIALIZED FORM HELPERS
' ===============================================================================================================================================================================

Private Function NormalizeChatGPTBuildScope(ByVal buildScope As String) As String
    Dim cleanedScope As String
    cleanedScope = UCase$(Replace$(Replace$(Trim$(buildScope), "_", vbNullString), "-", vbNullString))
    Select Case cleanedScope
        Case "", "ALL", "FULL"
            NormalizeChatGPTBuildScope = FORM_BUILD_SCOPE_ALL
        Case "CHATGPTUIFORM", "CHATGPTUIFORMS"
            NormalizeChatGPTBuildScope = FORM_BUILD_SCOPE_CHATGPT_UI_FORMS
        Case "ADCSUPPORTFORM", "ADCSUPPORTFORMS"
            NormalizeChatGPTBuildScope = FORM_BUILD_SCOPE_SUPPORT_FORMS
        Case "LOOKUPEDITOR", "LOOKUPEDITORS"
            NormalizeChatGPTBuildScope = FORM_BUILD_SCOPE_LOOKUP_EDITORS
        Case Else
            Err.Raise vbObjectError + 6605, "NormalizeChatGPTBuildScope", "Unsupported " & DEFAULT_UTILITY_LONG_NAME & " generated-form build scope: " & buildScope & ". Use All, ChatGPTUIForms, ADCSupportForms, or LookupEditors."
    End Select
End Function

Private Function GeneratedFormPreflightScopeStateFromText(ByVal normalizedBuildScope As String) As GeneratedFormPreflightScopeState
    Select Case normalizedBuildScope
        Case FORM_BUILD_SCOPE_ALL
            GeneratedFormPreflightScopeStateFromText = GENERATED_FORM_PREFLIGHT_SCOPE_ALL
        Case FORM_BUILD_SCOPE_CHATGPT_UI_FORMS
            GeneratedFormPreflightScopeStateFromText = GENERATED_FORM_PREFLIGHT_SCOPE_CHATGPT_UI_FORMS
        Case FORM_BUILD_SCOPE_SUPPORT_FORMS
            GeneratedFormPreflightScopeStateFromText = GENERATED_FORM_PREFLIGHT_SCOPE_SUPPORT_FORMS
        Case FORM_BUILD_SCOPE_LOOKUP_EDITORS
            GeneratedFormPreflightScopeStateFromText = GENERATED_FORM_PREFLIGHT_SCOPE_LOOKUP_EDITORS
        Case Else
            Err.Raise vbObjectError + 6605, "GeneratedFormPreflightScopeStateFromText", "Unsupported normalized generated-form build scope: " & normalizedBuildScope
    End Select
End Function

Private Function GeneratedFormPreflightScopeStateText(ByVal scopeState As GeneratedFormPreflightScopeState) As String
    Select Case scopeState
        Case GENERATED_FORM_PREFLIGHT_SCOPE_ALL
            GeneratedFormPreflightScopeStateText = FORM_BUILD_SCOPE_ALL
        Case GENERATED_FORM_PREFLIGHT_SCOPE_CHATGPT_UI_FORMS
            GeneratedFormPreflightScopeStateText = FORM_BUILD_SCOPE_CHATGPT_UI_FORMS
        Case GENERATED_FORM_PREFLIGHT_SCOPE_SUPPORT_FORMS
            GeneratedFormPreflightScopeStateText = FORM_BUILD_SCOPE_SUPPORT_FORMS
        Case GENERATED_FORM_PREFLIGHT_SCOPE_LOOKUP_EDITORS
            GeneratedFormPreflightScopeStateText = FORM_BUILD_SCOPE_LOOKUP_EDITORS
        Case Else
            GeneratedFormPreflightScopeStateText = "<none>"
    End Select
End Function

Private Function UITableExistsStrict(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
    On Error GoTo NotFoundOrFailed
    If db Is Nothing Then Err.Raise ERR_GENERATED_CATALOG_ROLLBACK_FAILED, "UITableExistsStrict", "No database object was supplied for strict rollback-table inspection."
    tableName = Trim$(tableName)
    If Len(tableName) = 0 Then Err.Raise ERR_GENERATED_CATALOG_ROLLBACK_FAILED, "UITableExistsStrict", "A blank table name was supplied for strict rollback-table inspection."
    UITableExistsStrict = DaoTableExists(db, tableName)
    Exit Function
NotFoundOrFailed:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "UITableExistsStrict", "Rollback-table inspection failed and was not treated as absence." & vbCrLf & "Table: " & tableName & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Function

Private Function ConfigureTwoColumnTableLookup(ByVal ctl As control, ByVal tableName As String, ByVal boundFieldName As String, ByVal displayFieldName As String, ByVal activeFieldName As String, ByVal orderFields As Variant, Optional ByVal columnWidthsText As String = "2.5;2.5", Optional ByVal limitToListValue As Boolean = True) As Boolean
    On Error GoTo ErrHandler
    Dim sqlText As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    If SharedDao_TableExists(GetHostDb(), tableName) = False Then Exit Function
    If SharedDao_TableFieldExists(GetHostDb(), tableName, boundFieldName) = False Then Exit Function
    If SharedDao_TableFieldExists(GetHostDb(), tableName, displayFieldName) = False Then Exit Function
    sqlText = "SELECT [" & boundFieldName & "], [" & displayFieldName & "] FROM " & SharedDao_BracketName(tableName)
    If Len(Trim$(activeFieldName)) > 0 Then
        If SharedDao_TableFieldExists(GetHostDb(), tableName, activeFieldName) Then sqlText = sqlText & " WHERE Nz([" & activeFieldName & "], False)<>False"
    End If
    sqlText = sqlText & ExistingLookupOrderByText(tableName, orderFields)
    sqlText = sqlText & ";"
    SharedFormDesign_ConfigureTwoColumnLookup ctl, sqlText
    ctl.columnWidths = NormalizeLookupColumnWidthsForAccess(columnWidthsText, 2, 1, limitToListValue)
    ctl.LimitToList = limitToListValue
    ConfigureTwoColumnTableLookup = True
    Exit Function
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ConfigureTwoColumnTableLookup", "Could not configure the generated table lookup." & vbCrLf & "Table: " & tableName & vbCrLf & "Bound field: " & boundFieldName & vbCrLf & "Display field: " & displayFieldName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function ActiveGeneratedFormsForNavigationPaneVerification(ByRef navNotes As Collection, Optional ByVal suppliedFormNames As Variant) As Collection
    'Returns and verifies the complete managed-generated catalog set. Physical build ownership remains separate.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim formNames As Collection
    Dim formItem As Variant
    Set db = GetHostDb()
    If SharedDao_TableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_FormDefinitions")) = False Then Err.Raise vbObjectError + 6701, "ActiveGeneratedFormsForNavigationPaneVerification", "Required generated-form catalog table was not found."
    If RequiredFormCatalogFieldsExist(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_FormDefinitions")) = False Then Err.Raise vbObjectError + 6702, "ActiveGeneratedFormsForNavigationPaneVerification", "Generated-form catalog is missing required fields."
    If isMissing(suppliedFormNames) = False Then
        If IsObject(suppliedFormNames) Then Set formNames = suppliedFormNames
    End If
    If formNames Is Nothing Then Set formNames = ManagedGeneratedFormNames()
    For Each formItem In formNames
        SetupControlPump
        If SharedDao_AccessObjectExists(db, acForm, CStr(formItem)) = False Then Err.Raise vbObjectError + 6703, "ActiveGeneratedFormsForNavigationPaneVerification", "An authoritative managed-generated form does not exist." & vbCrLf & "Form: " & CStr(formItem)
        VerifyGeneratedFormCatalogRow db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_FormDefinitions"), CStr(formItem)
    Next formItem
    If formNames.count = 0 Then Err.Raise vbObjectError + 6704, "ActiveGeneratedFormsForNavigationPaneVerification", "No authoritative managed-generated form definitions were available for Navigation Pane verification."
    navNotes.Add "Verified durable catalog metadata for all " & CStr(formNames.count) & " authoritative managed-generated forms before physical Navigation Pane group verification."
    Set ActiveGeneratedFormsForNavigationPaneVerification = formNames
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ActiveGeneratedFormsForNavigationPaneVerification", "Could not prepare the authoritative generated-form set for Navigation Pane verification." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Sub CreateChatGPTNavigationPaneGroupsForm(ByVal replaceExisting As Boolean, ByRef builtForms As Collection, ByRef skippedForms As Collection)
    On Error GoTo ErrHandler
    Dim formName As String
    Dim sourceName As String
    Dim sqlText As String
    Dim fields As Variant
    Dim orderFields As Variant
    formName = ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "NavigationPane_Groups")
    sourceName = GeneratedFormSourceNameForSignature(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "NavigationPane_Groups"))
    If Len(sourceName) = 0 Then
        AddSkippedFormNote skippedForms, formName, "Source was not found. " & "Expected " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups") & "."
        Exit Sub
    End If
    If TableOrQueryExists(sourceName) = False Then
        AddSkippedFormNote skippedForms, formName, "Source [" & sourceName & "] was not found. " & "Expected " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups") & "."
        Exit Sub
    End If
    fields = GeneratedFormListFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "NavigationPane_Groups"))
    orderFields = GeneratedFormOrderFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "NavigationPane_Groups"))
    sqlText = BuildSelectSql(sourceName, fields, vbNullString, orderFields)
    If Len(sqlText) = 0 Then
        AddSkippedFormNote skippedForms, formName, "Source [" & sourceName & "] did not contain any expected fields."
        Exit Sub
    End If
    If PrepareChatGPTGeneratedFormForBuild(formName, replaceExisting, skippedForms) = False Then Exit Sub
    CreateContinuousListForm formName, GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "NavigationPane_Groups")), sqlText, sourceName, fields, GeneratedFormIncludeOpenButton(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "NavigationPane_Groups")), GeneratedFormListActionProfile(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "NavigationPane_Groups")), GeneratedFormListAllowEdits(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "NavigationPane_Groups"))
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
    Err.Raise savedErrNumber, "CreateChatGPTNavigationPaneGroupsForm", savedErrDescription & vbCrLf & "Error source: " & savedErrSource
End Sub

Private Sub AddGeneratedListActionButtons(ByVal formName As String, ByVal actionProfile As String)
    Dim definitions As Collection
    Dim definitionData As Variant
    Dim leftPos As Long
    Set definitions = GeneratedListActionDefinitions(actionProfile)
    leftPos = DEFAULT_LIST_ACTION_LEFT
    For Each definitionData In definitions
        SharedFormDesign_AddButton formName, acHeader, vbNullString, CStr(definitionData(0)), CStr(definitionData(1)), leftPos, DEFAULT_LIST_ACTION_TOP, DEFAULT_LIST_ACTION_WIDTH, DEFAULT_LIST_ACTION_HEIGHT, CStr(definitionData(2))
        leftPos = leftPos + DEFAULT_LIST_ACTION_WIDTH + DEFAULT_LIST_ACTION_GAP
    Next definitionData
End Sub

Private Sub AddUtilityShellActionButtons(ByVal formName As String, ByVal parentName As String)
    Dim topPos As Long
    Dim left1 As Long
    Dim left2 As Long
    Dim left3 As Long
    left1 = DEFAULT_UI_TAB_LEFT + DEFAULT_UI_FORM_LEFT_MARGIN
    left2 = left1 + DEFAULT_UI_BUTTON_WIDTH + DEFAULT_UI_COLUMN_GAP
    left3 = left2 + DEFAULT_UI_BUTTON_WIDTH + DEFAULT_UI_COLUMN_GAP
    topPos = DEFAULT_UI_TAB_TOP + 540
    With SharedFormDesign_AddSectionLabel(formName, acDetail, parentName, "Utility Actions", left1, topPos, 5000, DEFAULT_UI_SECTION_TITLE_HEIGHT)
        .TextAlign = 1
    End With
    topPos = topPos + DEFAULT_UI_SECTION_TITLE_HEIGHT + DEFAULT_UI_SECTION_TITLE_CONTENT_GAP
    AddCoreSystemMacroUtilityActionButtons formName, parentName, left1, left2, left3, topPos, True
End Sub

Private Function BuildChatGPTChildErrorMessage(ByVal headlineText As String, ByVal errSource As String, ByVal errDescription As String, ByVal detailText As String) As String
    Dim msg As String
    msg = headlineText
    If Len(Trim$(detailText)) > 0 Then msg = msg & vbCrLf & detailText
    msg = msg & vbCrLf & "Error source: " & NullDisplayText(errSource)
    msg = msg & vbCrLf & "Error description: " & NullDisplayText(errDescription)
    msg = msg & TableDiagnosticsLocationForMessage()
    BuildChatGPTChildErrorMessage = msg
End Function

Private Function BuildChatGPTErrorMessage(ByVal headlineText As String, ByVal errNumber As Long, ByVal errSource As String, ByVal errDescription As String) As String
    Dim msg As String
    msg = headlineText
    msg = msg & vbCrLf & "Form builder version: " & FORM_BUILDER_VERSION
    msg = msg & vbCrLf & "Module: x__UTIL_ADC_UI_FORMS"
    msg = msg & vbCrLf & vbCrLf & "Build context:"
    msg = msg & vbCrLf & "Step: " & NullDisplayText(mChatGPTBuildStep)
    msg = msg & vbCrLf & "Form: " & NullDisplayText(mChatGPTBuildForm)
    msg = msg & vbCrLf & "Action: " & NullDisplayText(mChatGPTBuildAction)
    msg = msg & vbCrLf & "Detail: " & NullDisplayText(mChatGPTBuildDetail)
    msg = msg & vbCrLf & "Source: " & NullDisplayText(mChatGPTBuildSourceName)
    msg = msg & vbCrLf & "Field: " & NullDisplayText(mChatGPTBuildFieldName)
    msg = msg & vbCrLf & "Control: " & NullDisplayText(mChatGPTBuildControlName)
    msg = msg & vbCrLf & "Section: " & NullDisplayText(mChatGPTBuildSectionName)
    msg = msg & vbCrLf & "Parent/Page: " & NullDisplayText(mChatGPTBuildParentName)
    msg = msg & vbCrLf & "Bounds: " & NullDisplayText(mChatGPTBuildBounds)
    msg = msg & vbCrLf & "Live form state: " & ChatGPTBuildFormStateForMessage(mChatGPTBuildForm)
    msg = msg & vbCrLf & vbCrLf & "Original error:"
    msg = msg & vbCrLf & "Number: " & CStr(errNumber)
    msg = msg & vbCrLf & "Source: " & NullDisplayText(errSource)
    msg = msg & vbCrLf & "Description: " & NullDisplayText(errDescription)
    msg = msg & TableDiagnosticsLocationForMessage()
    msg = msg & ProcessRunSummaryForMessage("BuildGeneratedForms")
    msg = msg & vbCrLf & vbCrLf & "Troubleshooting:"
    msg = msg & vbCrLf & "- The failing form/action/field/control/bounds above identify the exact generation boundary."
    msg = msg & vbCrLf & "- The Immediate Window receives this full diagnostic even when durable trace-table writes are unavailable."
    msg = msg & vbCrLf & "- Partially created ChatGPT_ forms remain protected by the generated-form rollback workflow."
    BuildChatGPTErrorMessage = msg
End Function

Private Sub DebugPrintChatGPTBuildError(ByVal messageText As String, ByVal errNumber As Long, ByVal errSource As String, ByVal errDescription As String)
    On Error Resume Next
    Debug.Print String$(96, "=")
    Debug.Print DEFAULT_UTILITY_LONG_NAME & " generated-form failure"
    Debug.Print "Timestamp: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Debug.Print "Form builder version: " & FORM_BUILDER_VERSION
    Debug.Print "Step: " & NullDisplayText(mChatGPTBuildStep)
    Debug.Print "Form: " & NullDisplayText(mChatGPTBuildForm)
    Debug.Print "Action: " & NullDisplayText(mChatGPTBuildAction)
    Debug.Print "Detail: " & NullDisplayText(mChatGPTBuildDetail)
    Debug.Print "Source: " & NullDisplayText(mChatGPTBuildSourceName)
    Debug.Print "Field: " & NullDisplayText(mChatGPTBuildFieldName)
    Debug.Print "Control: " & NullDisplayText(mChatGPTBuildControlName)
    Debug.Print "Section: " & NullDisplayText(mChatGPTBuildSectionName)
    Debug.Print "Parent/Page: " & NullDisplayText(mChatGPTBuildParentName)
    Debug.Print "Bounds: " & NullDisplayText(mChatGPTBuildBounds)
    Debug.Print "Live form state: " & ChatGPTBuildFormStateForMessage(mChatGPTBuildForm)
    Debug.Print "Original error number: " & CStr(errNumber)
    Debug.Print "Original error source: " & NullDisplayText(errSource)
    Debug.Print "Original error description: " & NullDisplayText(errDescription)
    Debug.Print "Full diagnostic message:"
    Debug.Print messageText
    Debug.Print String$(96, "=")
    LogTraceMessage TRACE_LEVEL_DETAIL, "UI", "DebugPrintChatGPTBuildError", "ChatGPT " & DEFAULT_UTILITY_LONG_NAME & " form-generation failure detail recorded.", messageText, mChatGPTBuildStep, errNumber, errSource, errDescription, mChatGPTBuildForm, "Form", "CodeDb", "FormGeneration"
    On Error GoTo 0
End Sub

Private Function ChatGPTRecordsetCount(ByVal db As DAO.Database, ByVal sqlText As String) As Long
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    ChatGPTRecordsetCount = CLng(rs!RowCountValue)
    rs.Close
End Function

Private Function ChatGPTVariantValuesEqual(ByVal leftValue As Variant, ByVal rightValue As Variant) As Boolean
    If IsNull(leftValue) Or IsNull(rightValue) Then
        ChatGPTVariantValuesEqual = (IsNull(leftValue) And IsNull(rightValue))
        Exit Function
    End If
    If IsDate(leftValue) And IsDate(rightValue) Then
        ChatGPTVariantValuesEqual = (CDbl(CDate(leftValue)) = CDbl(CDate(rightValue)))
        Exit Function
    End If
    If IsNumeric(leftValue) And IsNumeric(rightValue) Then
        ChatGPTVariantValuesEqual = (CDec(leftValue) = CDec(rightValue))
        Exit Function
    End If
    ChatGPTVariantValuesEqual = (StrComp(CStr(leftValue), CStr(rightValue), vbBinaryCompare) = 0)
End Function

Private Sub ChatGPT_EnsureFoundation(Optional ByVal foundationMode As String = FOUNDATION_MODE_FULL)
    'Direct generated-form builds cannot assume RunFullSetup already refreshed table and retained form-design source-query dependencies.
    Select Case foundationMode
        Case FOUNDATION_MODE_FULL
            ChatGPT_RunFoundationStep "RepairTables", DEFAULT_UTILITY_LONG_NAME & " live table and generated-form foundation repair"
            LogTraceMessage TRACE_LEVEL_STATUS, "UI", "ChatGPT_EnsureFoundation", "Skipped separate review/form-design foundation repair because table repair already applies it.", "Foundation mode: " & foundationMode, "Prepare form foundation", 0, vbNullString, vbNullString, DEFAULT_CHATGPT_UI_DISPLAY_PREFIX, "FormSet", "CodeDb", "FormGeneration"
        Case FOUNDATION_MODE_REVIEW_ONLY
            ChatGPT_RunFoundationStep "RepairNonFormQueryFoundationForSetup", DEFAULT_UTILITY_LONG_NAME & " review/form-design query foundation repair"
        Case FOUNDATION_MODE_SKIP
            LogTraceMessage TRACE_LEVEL_STATUS, "UI", "ChatGPT_EnsureFoundation", "Foundation repair skipped by caller.", "Foundation mode: " & foundationMode, "Prepare form foundation", 0, vbNullString, vbNullString, DEFAULT_CHATGPT_UI_DISPLAY_PREFIX, "FormSet", "CodeDb", "FormGeneration"
        Case Else
            Err.Raise vbObjectError + 6603, "ChatGPT_EnsureFoundation", "Unsupported " & DEFAULT_UTILITY_LONG_NAME & " generated-form foundation mode: " & foundationMode
    End Select
End Sub

Private Sub ChatGPT_RunFoundationStep(ByVal ProcedureName As String, ByVal stepDescription As String)
    Dim errNumber As Long
    Dim errDescription As String
    If SharedRun_TryPublicProcedure1(ProcedureName, False, errNumber, errDescription) Then Exit Sub
    Err.Raise vbObjectError + 6602, "ChatGPT_EnsureFoundation", stepDescription & " failed or was unavailable." & vbCrLf & "Required public procedure: " & ProcedureName & vbCrLf & "Error " & CStr(errNumber) & ": " & errDescription
End Sub

Private Function NormalizeChatGPTFoundationMode(ByVal foundationMode As String) As String
    Dim cleanedMode As String
    cleanedMode = UCase$(Replace$(Replace$(Trim$(foundationMode), "_", vbNullString), "-", vbNullString))
    Select Case cleanedMode
        Case "", "FULL", "TRUE", "YES"
            NormalizeChatGPTFoundationMode = FOUNDATION_MODE_FULL
        Case "REVIEWONLY", "REVIEW", "QUERIESONLY", "QUERYONLY"
            NormalizeChatGPTFoundationMode = FOUNDATION_MODE_REVIEW_ONLY
        Case "SKIP", "NONE", "NO", "FALSE"
            NormalizeChatGPTFoundationMode = FOUNDATION_MODE_SKIP
        Case Else
            Err.Raise vbObjectError + 6604, "NormalizeChatGPTFoundationMode", "Unsupported " & DEFAULT_UTILITY_LONG_NAME & " generated-form foundation mode: " & foundationMode & ". Use Full, ReviewOnly, or Skip."
    End Select
End Function

Private Sub ConfigureComboBox(ByVal ctl As control, ByVal fieldName As String, Optional ByVal sourceName As String = vbNullString, Optional ByVal suppliedLookupDefinition As Variant)
    On Error GoTo ErrHandler
    Dim valueListText As String
    Dim controlName As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    controlName = ctl.Name
    If isMissing(suppliedLookupDefinition) Then
        If ConfigureLookupComboBox(ctl, fieldName, sourceName) Then Exit Sub
    Else
        If ConfigureLookupComboBox(ctl, fieldName, sourceName, suppliedLookupDefinition) Then Exit Sub
    End If
    If LookupFieldRequiresTableOrQuerySource(fieldName) Then Err.Raise vbObjectError + 5801, "ConfigureComboBox", "The required lookup source could not be configured for field [" & fieldName & "]."
    valueListText = ValueListForField(fieldName)
    If Len(valueListText) = 0 Then Err.Raise vbObjectError + 5802, "ConfigureComboBox", "No lookup source or static value list is defined for combo field [" & fieldName & "]."
    ctl.rowSourceType = "Value List"
    ctl.rowSource = valueListText
    ctl.LimitToList = False
    ctl.columnCount = 1
    ctl.BoundColumn = 1
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ConfigureComboBox", BuildChatGPTChildErrorMessage("Could not configure generated combo box.", savedErrSource, savedErrDescription, "Control: " & controlName & "; Field: " & fieldName & "; Source: " & sourceName)
End Sub

Private Sub ConfigureConfigurationDetailTabs(ByVal formName As String, ByVal sourceName As String, ByVal tabCtl As control)
    Dim pg As Page
    Set pg = ConfigureTabPage(tabCtl, 0, "pgSummaryReview", "Summary / Review")
    AddFieldBlock formName, sourceName, SummaryReviewFields(), acDetail, pg.Name, DEFAULT_UI_TAB_LEFT + DEFAULT_UI_FORM_LEFT_MARGIN, DEFAULT_UI_TAB_TOP + 620, 2, "sum", False
    Set pg = ConfigureTabPage(tabCtl, 1, "pgTarget", "Target")
    AddFieldBlock formName, sourceName, TargetFields(), acDetail, pg.Name, DEFAULT_UI_TAB_LEFT + DEFAULT_UI_FORM_LEFT_MARGIN, DEFAULT_UI_TAB_TOP + 620, 2, "tgt", False
    Set pg = ConfigureTabPage(tabCtl, 2, "pgValidationGeneration", "Validation / Generation")
    AddFieldBlock formName, sourceName, ValidationGenerationFields(), acDetail, pg.Name, DEFAULT_UI_TAB_LEFT + DEFAULT_UI_FORM_LEFT_MARGIN, DEFAULT_UI_TAB_TOP + 620, 1, "val", False
    Set pg = ConfigureTabPage(tabCtl, 3, "pgAdminCleanup", "Admin / Cleanup")
    AddFieldBlock formName, sourceName, AdminCleanupFields(), acDetail, pg.Name, DEFAULT_UI_TAB_LEFT + DEFAULT_UI_FORM_LEFT_MARGIN, DEFAULT_UI_TAB_TOP + 620, 1, "adm", False
End Sub

Private Function ConfigureLiveAccessObjectLookup(ByVal ctl As control, ByVal rowSourceText As String, Optional ByVal columnWidthsText As String = "2.5;4", Optional ByVal limitToListValue As Boolean = False, Optional ByVal columnCountValue As Long = 2, Optional ByVal boundColumnValue As Long = 1) As Boolean
    On Error GoTo ErrHandler
    Dim controlName As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    controlName = ctl.Name
    rowSourceText = Trim$(rowSourceText)
    columnWidthsText = Trim$(columnWidthsText)
    If Len(rowSourceText) = 0 Then Exit Function
    If Len(columnWidthsText) = 0 Then columnWidthsText = "2.5;4"
    ctl.rowSourceType = "Table/Query"
    ctl.rowSource = rowSourceText
    If columnCountValue < 1 Then columnCountValue = 1
    If boundColumnValue < 1 Or boundColumnValue > columnCountValue Then boundColumnValue = 1
    ctl.columnCount = columnCountValue
    ctl.BoundColumn = boundColumnValue
    ctl.columnWidths = columnWidthsText
    ctl.LimitToList = limitToListValue
    ConfigureLiveAccessObjectLookup = True
    Exit Function
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ConfigureLiveAccessObjectLookup", "Could not configure a live Access-object lookup." & vbCrLf & "Control: " & controlName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Sub ConfigureLookupEditorForComboBox(ByVal ctl As control, ByVal lookupSourceTable As String)
    On Error GoTo ErrHandler
    Dim policyCode As String
    Dim editorFormName As String
    editorFormName = LookupEditorPhysicalFormNameForSource(lookupSourceTable, policyCode)
    If StrComp(policyCode, LOOKUP_EDITOR_POLICY_NONE, vbTextCompare) = 0 Then
        ctl.AllowValueListEdits = False
        ctl.ListItemsEditForm = vbNullString
        Exit Sub
    End If
    If Len(Trim$(editorFormName)) = 0 Then Err.Raise vbObjectError + 7462, "ConfigureLookupEditorForComboBox", "Authoritative lookup-editor policy does not resolve to a canonical editor form identity." & vbCrLf & "Lookup source: " & lookupSourceTable & vbCrLf & "Policy: " & policyCode
    'D2 owns the generated combo-box contract, so persist the authoritative editor identity even when D3 has not created that editor form yet.
    'D3 owns physical support-form creation; D4 separately owns table-field binding.
    ctl.AllowValueListEdits = True
    ctl.ListItemsEditForm = editorFormName
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ConfigureLookupEditorForComboBox", "Could not configure native Edit List Items metadata for generated combo box." & vbCrLf & "Lookup source: " & lookupSourceTable & vbCrLf & "Editor form: " & editorFormName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Function ConfigureLookupComboBox(ByVal ctl As control, ByVal fieldName As String, Optional ByVal sourceName As String = vbNullString, Optional ByVal suppliedDefinitionData As Variant) As Boolean
    On Error GoTo ErrHandler
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim definitionData As Variant
    Dim lookupKind As String
    If isMissing(suppliedDefinitionData) Then
        definitionData = GeneratedFormLookupDefinition(fieldName, sourceName)
    Else
        definitionData = suppliedDefinitionData
    End If
    If IsArray(definitionData) = False Then Exit Function
    lookupKind = CStr(definitionData(0))
    ctl.AllowValueListEdits = False
    ctl.ListItemsEditForm = vbNullString
    Select Case lookupKind
        Case GENERATED_FORM_LOOKUP_KIND_TWO_COLUMN
            ConfigureLookupComboBox = ConfigureTwoColumnTableLookup(ctl, CStr(definitionData(1)), CStr(definitionData(2)), CStr(definitionData(3)), CStr(definitionData(4)), definitionData(5), CStr(definitionData(6)), CBool(definitionData(7)))
            If ConfigureLookupComboBox Then ConfigureLookupEditorForComboBox ctl, CStr(definitionData(1))
        Case GENERATED_FORM_LOOKUP_KIND_LIVE_TARGET
            ConfigureLookupComboBox = ConfigureLiveAccessObjectLookup(ctl, LiveTargetObjectsLookupRowSource(), CStr(definitionData(6)), CBool(definitionData(7)))
        Case GENERATED_FORM_LOOKUP_KIND_LIVE_TABLE_QUERY
            ConfigureLookupComboBox = ConfigureLiveAccessObjectLookup(ctl, LiveTableAndQueryLookupRowSource(), CStr(definitionData(6)), CBool(definitionData(7)))
        Case GENERATED_FORM_LOOKUP_KIND_LIVE_ONE_COLUMN
            ConfigureLookupComboBox = ConfigureLiveAccessObjectLookup(ctl, CStr(definitionData(1)), CStr(definitionData(6)), CBool(definitionData(7)), 1, 1)
        Case GENERATED_FORM_LOOKUP_KIND_VALUE_LIST
            If Len(Trim$(CStr(definitionData(1)))) > 0 Then
                ctl.rowSourceType = "Value List"
                ctl.rowSource = CStr(definitionData(1))
                ctl.LimitToList = CBool(definitionData(7))
                ctl.columnCount = 1
                ctl.BoundColumn = 1
                ConfigureLookupComboBox = True
            End If
        Case GENERATED_FORM_LOOKUP_KIND_DISTINCT_FIELD
            ConfigureLookupComboBox = ConfigureDistinctFieldLookup(ctl, CStr(definitionData(1)), CStr(definitionData(2)))
    End Select
    Exit Function
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ConfigureLookupComboBox", "Could not configure the generated lookup for field [" & fieldName & "]." & vbCrLf & "Source: " & sourceName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

' Purpose: Configures tab page for tab ctl, page index.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   tabCtl (Input; control): Access TabControl instance on which the requested Page is created/found/configured.
'   pageIndex (Input; Long): Index selecting the page index item or position within the caller-supplied definition/source sequence.
'   pageName (Input; String): Exact tab-page name to create/find/configure on the generated TabControl.
'   captionText (Input; String): Caption text text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
' Returns: Page: Returns the configured Access tab Page selected/created for the requested page name/caption and layout contract.
Public Function ConfigureTabPage(ByVal tabCtl As control, ByVal pageIndex As Long, ByVal pageName As String, ByVal captionText As String) As Page
    On Error GoTo ErrHandler
    Dim detailText As String
    detailText = "TabControl: " & tabCtl.Name & "; PageIndex: " & CStr(pageIndex) & "; PageName: " & pageName & "; Caption: " & captionText
    SetChatGPTBuildContext mChatGPTBuildStep, mChatGPTBuildForm, "Configure tab page", detailText
    Set ConfigureTabPage = SharedFormDesign_ConfigureTabPage(tabCtl, pageIndex, pageName, captionText)
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ConfigureTabPage", BuildChatGPTChildErrorMessage("Could not configure tab page.", savedErrSource, savedErrDescription, detailText)
End Function

Private Sub ConfigureTabbedUtilityShell(ByVal formName As String, ByVal tabCtl As control)
    Dim pg As Page
    Set pg = ConfigureTabPage(tabCtl, 0, "pgUtilityActions", "Actions")
    AddUtilityShellActionButtons formName, pg.Name
    AddEmbeddedFormTabPage formName, tabCtl, 1, "pgConfigurationDetail", "Configuration Detail", ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Configuration_Detail"), "Fixed-header and tabbed configuration editor."
    AddEmbeddedFormTabPage formName, tabCtl, 2, "pgConfigurationList", "Configuration List", ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Configuration_List"), "Continuous-list configuration review."
    AddEmbeddedFormTabPage formName, tabCtl, 3, "pgCandidateReview", "Candidate Review", ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Candidate_Review"), "Candidate-review rows for review and acceptance."
    AddEmbeddedFormTabPage formName, tabCtl, 4, "pgSettings", "ADC_Settings", ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Settings"), DEFAULT_UTILITY_LONG_NAME & " settings."
    AddEmbeddedFormTabPage formName, tabCtl, 5, "pgHealth", CANON_GENERATED_FORM_RIBBON_TITLE_HEALTH_REVIEW, ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Health_Review"), "Non-destructive health-review findings."
    AddEmbeddedFormTabPage formName, tabCtl, 6, "pgLauncher", "Launcher", ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "LauncherControlsReviewForm"), "Custom Macro Launcher review rows."
    AddEmbeddedFormTabPage formName, tabCtl, 7, "pgCatalog", CANON_GENERATED_FORM_RIBBON_TITLE_CATALOG_REVIEW, ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Catalog_Review"), "[" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID) & "] catalog review rows."
    AddEmbeddedFormTabPage formName, tabCtl, 8, "pgSystemMacros", "System Macros", ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "System_Macros"), "Table-driven built-in system macro rows."
    AddEmbeddedFormTabPage formName, tabCtl, 9, "pgNavGroups", "Navigation Pane Groups", ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "NavigationPane_Groups"), "Editable Navigation Pane group catalog rows."
    AddEmbeddedFormTabPage formName, tabCtl, 10, "pgFormTypes", "Form Types", ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Form_Types"), "Form type lookup rows."
    AddEmbeddedFormTabPage formName, tabCtl, 11, "pgFormsCatalog", "Forms Catalog", ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Forms"), "Relevant " & DEFAULT_UTILITY_LONG_NAME & " form rows."
    AddEmbeddedFormTabPage formName, tabCtl, 12, "pgActiveReferences", "Active References", ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Active_References"), "Active reference inventory with loaded-project match status, scan, and file-location actions."
    AddEmbeddedFormTabPage formName, tabCtl, 13, "pgAvailableReferences", "Available References", ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Available_References"), "Available registered type-library inventory."
    AddEmbeddedFormTabPage formName, tabCtl, 14, "pgLoadedVBAProjects", "Loaded VBA Projects", ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review_Projects"), "Reusable loaded-project inventory with source-analysis status and counts."
    AddEmbeddedFormTabPage formName, tabCtl, 15, "pgReferenceManager", "Reference Manager", ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Reference_Manager"), "Requested active-reference state and guarded apply/discard actions."
    AddEmbeddedFormTabPage formName, tabCtl, 16, "pgModuleCurrent", "Module Current", ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Updates_Current"), "Expected-current module manifest and latest verification state."
    AddEmbeddedFormTabPage formName, tabCtl, 17, "pgModuleHistory", "Module History", ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Updates_History"), "Append-only module update history."
    AddEmbeddedFormTabPage formName, tabCtl, 18, "pgModuleVerification", "Module Verification", ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Verification"), "Expected-versus-installed verification snapshots."
    AddEmbeddedFormTabPage formName, tabCtl, 19, "pgHelp", "Help", ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Help"), "Searchable " & DEFAULT_UTILITY_LONG_NAME & " Help and table-backed database reference."
    tabCtl.value = 0
    tabCtl.OnChange = "=UtilityLazyLoadSelectedTab()"
End Sub

Private Sub CreateChatGPTCandidateReviewForm(ByVal replaceExisting As Boolean, ByRef builtForms As Collection, ByRef skippedForms As Collection)
    On Error GoTo ErrHandler
    Dim sourceName As String
    Dim whereText As String
    sourceName = GeneratedFormSourceNameForSignature(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Candidate_Review"))
    If Len(sourceName) > 0 Then
        If SharedDao_DatabaseFieldExistsInSource(GetHostDb(), sourceName, FIELD_ROW_KIND) Then whereText = SQL_FIELD_ROW_KIND & "=" & SharedDao_SqlQuote(ROW_KIND_CANDIDATE)
    End If
    CreateConfigurationListLikeForm ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Candidate_Review"), GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Candidate_Review")), sourceName, whereText, replaceExisting, builtForms, skippedForms, True
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "CreateChatGPTCandidateReviewForm", "Could not create the candidate review form." & vbCrLf & "Source: " & NullDisplayText(sourceName) & vbCrLf & "Where: " & NullDisplayText(whereText) & vbCrLf & "Error source: " & savedErrSource & vbCrLf & "Error description: " & savedErrDescription
End Sub

Private Sub CreateChatGPTCandidateSplitPopupForm(ByVal replaceExisting As Boolean, ByRef builtForms As Collection, ByRef skippedForms As Collection)
    On Error GoTo ErrHandler
    Dim sourceName As String
    Dim whereText As String
    sourceName = GeneratedFormSourceNameForSignature(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Candidate_Review"))
    If Len(sourceName) > 0 Then
        If SharedDao_DatabaseFieldExistsInSource(GetHostDb(), sourceName, FIELD_ROW_KIND) Then whereText = SQL_FIELD_ROW_KIND & "=" & SharedDao_SqlQuote(ROW_KIND_CANDIDATE)
    End If
    CreateSplitPopupForm ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Candidate_Split"), GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Candidate_Split")), sourceName, whereText, CandidateSplitPopupFields(), GeneratedFormOrderFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Candidate_Split")), replaceExisting, builtForms, skippedForms, True, GeneratedSplitPopupDescription(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Candidate_Split"))
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "CreateChatGPTCandidateSplitPopupForm", "Could not create the candidate split popup form." & vbCrLf & "Source: " & NullDisplayText(sourceName) & vbCrLf & "Where: " & NullDisplayText(whereText) & vbCrLf & "Error source: " & savedErrSource & vbCrLf & "Error description: " & savedErrDescription
End Sub

Private Sub CreateChatGPTHealthReviewForm(ByVal replaceExisting As Boolean, ByRef builtForms As Collection, ByRef skippedForms As Collection)
    On Error GoTo ErrHandler
    Dim formName As String
    Dim sourceName As String
    Dim sqlText As String
    Dim fields As Variant
    Dim orderFields As Variant
    formName = ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Health_Review")
    sourceName = GeneratedFormSourceNameForSignature(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Health_Review"))
    If Len(sourceName) = 0 Then
        AddSkippedFormNote skippedForms, formName, "Source was not found. " & "Expected " & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_System_HealthReview") & " or " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReview") & "."
        Exit Sub
    End If
    If TableOrQueryExists(sourceName) = False Then
        AddSkippedFormNote skippedForms, formName, "Source [" & sourceName & "] was not found. " & "Expected " & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_System_HealthReview") & " or " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReview") & "."
        Exit Sub
    End If
    fields = GeneratedFormListFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Health_Review"))
    orderFields = GeneratedFormOrderFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Health_Review"))
    sqlText = BuildSelectSql(sourceName, fields, vbNullString, orderFields)
    If Len(sqlText) = 0 Then
        AddSkippedFormNote skippedForms, formName, "Source [" & sourceName & "] did not contain any expected fields."
        Exit Sub
    End If
    If PrepareChatGPTGeneratedFormForBuild(formName, replaceExisting, skippedForms) = False Then Exit Sub
    CreateContinuousListForm formName, GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Health_Review")), sqlText, sourceName, fields, GeneratedFormIncludeOpenButton(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Health_Review")), GeneratedFormListActionProfile(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Health_Review")), GeneratedFormListAllowEdits(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Health_Review"))
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
    Err.Raise savedErrNumber, "CreateChatGPTHealthReviewForm", savedErrDescription & vbCrLf & "Error source: " & savedErrSource
End Sub

Private Sub CreateChatGPTSettingsForm(ByVal replaceExisting As Boolean, ByRef builtForms As Collection, ByRef skippedForms As Collection)
    On Error GoTo ErrHandler
    Dim formName As String
    Dim sourceName As String
    Dim sqlText As String
    Dim fields As Variant
    Dim orderFields As Variant
    formName = ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Settings")
    sourceName = GeneratedFormSourceNameForSignature(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Settings"))
    If Len(sourceName) = 0 Then
        AddSkippedFormNote skippedForms, formName, "Source was not found. " & "Expected " & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Settings_Resolved") & " or " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID) & "."
        Exit Sub
    End If
    If TableOrQueryExists(sourceName) = False Then
        AddSkippedFormNote skippedForms, formName, "Source [" & sourceName & "] was not found. " & "Expected " & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Settings_Resolved") & " or " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID) & "."
        Exit Sub
    End If
    fields = GeneratedFormListFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Settings"))
    orderFields = GeneratedFormOrderFields(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Settings"))
    sqlText = BuildSelectSql(sourceName, fields, vbNullString, orderFields)
    If Len(sqlText) = 0 Then
        AddSkippedFormNote skippedForms, formName, "Source [" & sourceName & "] did not contain any expected fields."
        Exit Sub
    End If
    If PrepareChatGPTGeneratedFormForBuild(formName, replaceExisting, skippedForms) = False Then Exit Sub
    CreateContinuousListForm formName, GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Settings")), sqlText, sourceName, fields, GeneratedFormIncludeOpenButton(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Settings")), GeneratedFormListActionProfile(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Settings")), GeneratedFormListAllowEdits(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Settings"))
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
    Err.Raise savedErrNumber, "CreateChatGPTSettingsForm", savedErrDescription & vbCrLf & "Error source: " & savedErrSource
End Sub

Private Sub EnforceExactBoundControlDefinition(ByVal ctl As Access.control, ByVal formName As String, ByVal sourceName As String, ByVal Definition As Variant, ByRef createdControlNames As Object)
    On Error GoTo ErrHandler
    Dim fieldName As String
    Dim expectedControlName As String
    Dim expectedControlType As Long
    Dim expectedSection As Long
    Dim actualControlName As String
    fieldName = CStr(Definition(0))
    expectedControlType = CLng(Definition(3))
    expectedControlName = CStr(Definition(5))
    expectedSection = CLng(Definition(6))
    If ctl Is Nothing Then Err.Raise vbObjectError + 7410, "EnforceExactBoundControlDefinition", "The shared form-control builder returned no control."
    ctl.controlSource = fieldName
    ctl.Visible = True
    actualControlName = Trim$(ctl.Name)
    If StrComp(actualControlName, expectedControlName, vbTextCompare) <> 0 Then
        ctl.Name = expectedControlName
        actualControlName = Trim$(ctl.Name)
    End If
    If Len(actualControlName) = 0 Then Err.Raise vbObjectError + 7411, "EnforceExactBoundControlDefinition", "The generated bound control has a blank name."
    If createdControlNames.Exists(actualControlName) Then Err.Raise vbObjectError + 7412, "EnforceExactBoundControlDefinition", "The generated exact-contract form contains duplicate bound-control names." & vbCrLf & "Control: " & actualControlName
    createdControlNames.Add actualControlName, True
    If StrComp(Trim$(Nz(ctl.controlSource, vbNullString)), fieldName, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7413, "EnforceExactBoundControlDefinition", "The generated control did not retain its exact authoritative ControlSource." & vbCrLf & "Expected: " & fieldName & vbCrLf & "Actual: " & Nz(ctl.controlSource, vbNullString)
    If CLng(ctl.controlType) <> expectedControlType Then Err.Raise vbObjectError + 7414, "EnforceExactBoundControlDefinition", "The generated control type does not match the authoritative definition." & vbCrLf & "Expected type: " & CStr(expectedControlType) & vbCrLf & "Actual type: " & CStr(ctl.controlType)
    If CLng(ctl.Section) <> expectedSection Then Err.Raise vbObjectError + 7415, "EnforceExactBoundControlDefinition", "The generated control section does not match the authoritative definition." & vbCrLf & "Expected section: " & CStr(expectedSection) & vbCrLf & "Actual section: " & CStr(ctl.Section)
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "EnforceExactBoundControlDefinition", "Could not enforce an exact generated bound-control definition." & vbCrLf & "Form: " & formName & vbCrLf & "Source: " & sourceName & vbCrLf & "Field: " & fieldName & vbCrLf & "Projected control: " & expectedControlName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

' Purpose: Ensures header footer sections for form name.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   formName (Input; String): Physical Access form name identifying the form to create, open, configure, validate, or inspect.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub EnsureHeaderFooterSections(ByVal formName As String)
    On Error GoTo AddSections
    SetChatGPTBuildContext mChatGPTBuildStep, formName, "Verify form header/footer sections.", vbNullString
    forms(formName).Section(acHeader).Visible = True
    forms(formName).Section(acFooter).Visible = True
    Exit Sub
AddSections:
    Err.Clear
    On Error GoTo ErrHandler
    SetChatGPTBuildContext mChatGPTBuildStep, formName, "Create form header/footer sections.", "Running acCmdFormHdrFtr."
    DoCmd.SelectObject acForm, formName, False
    DoCmd.RunCommand acCmdFormHdrFtr
    forms(formName).Section(acHeader).Visible = True
    forms(formName).Section(acFooter).Visible = True
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "EnsureHeaderFooterSections", BuildChatGPTChildErrorMessage("Could not ensure form header/footer sections.", savedErrSource, savedErrDescription, "Form: " & formName)
End Sub

Private Function MaxRowTop(ByRef values() As Long) As Long
    Dim index As Long
    Dim resultValue As Long
    For index = LBound(values) To UBound(values)
        SetupControlPump
        If values(index) > resultValue Then resultValue = values(index)
    Next index
    MaxRowTop = resultValue
End Function

Private Sub SetAllRowTops(ByRef values() As Long, ByVal newValue As Long)
    Dim index As Long
    For index = LBound(values) To UBound(values)
        SetupControlPump
        values(index) = newValue
    Next index
End Sub

' ===============================================================================================================================================================================
' REPORT ERROR / DELETION / GENERATED-FORM RUNTIME CALLBACKS
' ===============================================================================================================================================================================

Private Function CompleteReportedError(ByVal formName As String, ByVal descriptionText As String, ByVal showSuccessMessage As Boolean) As Boolean
    On Error GoTo ErrHandler
    Dim screenshotPath As String
    Dim bitmapPath As String
    Dim captureWidth As Long
    Dim captureHeight As Long
    Dim traceLogID As Long
    If Len(Trim$(descriptionText)) = 0 Then Err.Raise vbObjectError + 9130, "CompleteReportedError", "The confirmed error report did not contain a description."
    If CurrentProject.AllForms(formName).IsLoaded Then forms(formName).Visible = False
    DoEvents
    ReportErrorSleep 250
    bitmapPath = ReportErrorTemporaryFilePath(".bmp")
    screenshotPath = ReportErrorCaptureAccessWindow(bitmapPath, captureWidth, captureHeight)
    traceLogID = LogReportedError(descriptionText, screenshotPath, captureWidth, captureHeight)
    If CurrentProject.AllForms(formName).IsLoaded Then DoCmd.Close acForm, formName, acSaveNo
    ReportErrorDeleteTemporaryFile screenshotPath
    If StrComp(bitmapPath, screenshotPath, vbTextCompare) <> 0 Then ReportErrorDeleteTemporaryFile bitmapPath
    If showSuccessMessage Then CoreShowMessage "The error report was saved." & vbCrLf & vbCrLf & "TraceLog ID: " & CStr(traceLogID) & vbCrLf & "Screenshot table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_TraceLogScreenshots"), vbInformation, "Report Error", "CompleteReportedError", "UI"
    CompleteReportedError = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Len(formName) > 0 Then
        If CurrentProject.AllForms(formName).IsLoaded Then DoCmd.Close acForm, formName, acSaveNo
    End If
    ReportErrorDeleteTemporaryFile screenshotPath
    If StrComp(bitmapPath, screenshotPath, vbTextCompare) <> 0 Then ReportErrorDeleteTemporaryFile bitmapPath
    On Error GoTo 0
    If savedErrNumber = 0 Then savedErrNumber = vbObjectError + 9131
    If Len(Trim$(savedErrSource)) = 0 Then savedErrSource = "CompleteReportedError"
    Err.Raise savedErrNumber, "CompleteReportedError", "The error report could not be completed." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

' Purpose: Performs report error dialog submit by coordinating `ManagedObjectPhysicalName` and `CoreShowMessage` while preserving this module's existing validation and error semantics.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters: None.
' Returns: Boolean: True when report error dialog cancel completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or not-applicable path.
Public Function ReportErrorDialogSubmit() As Boolean
    ' Wrapper purpose: Generated Report Error dialog callback that reads the dialog state and completes the Submit action contract.
    ' Why retained: Retained because the generated button expression needs a stable public Function entry point while dialog-state extraction and completion remain UI_FORMS-owned.
    On Error GoTo ErrHandler
    Dim formName As String
    Dim descriptionControlName As String
    Dim descriptionText As String
    formName = ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Report_Error_Dialog")
    If CurrentProject.AllForms(formName).IsLoaded = False Then Exit Function
    descriptionControlName = GeneratedFormContractControlName(formName, 1)
    descriptionText = CStr(Nz(forms(formName).Controls(descriptionControlName).value, vbNullString))
    If Len(Trim$(descriptionText)) = 0 Then
        CoreShowMessage "A description of the error is required before the report can continue.", vbExclamation, "Report Error", "ReportErrorDialogSubmit", "UI"
        forms(formName).Controls(descriptionControlName).SetFocus
        Exit Function
    End If
    ReportErrorDialogSubmit = CompleteReportedError(formName, descriptionText, True)
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    CoreShowErrorMessage "The error-report workflow could not be completed." & vbCrLf & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, vbCritical, "Report Error", "ReportErrorDialogSubmit", "UI"
    ReportErrorDialogSubmit = False
End Function

' Purpose: Performs report error dialog cancel by coordinating `ManagedObjectPhysicalName` and `End` while preserving this module's existing validation and error semantics.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters: None.
' Returns: Boolean: True when report error dialog cancel completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or not-applicable path.
Public Function ReportErrorDialogCancel() As Boolean
    ' Wrapper purpose: Generated Report Error dialog callback that performs the Cancel/close action and returns Access-compatible Boolean success.
    ' Why retained: Retained because the generated button expression needs a stable public Function entry point rather than direct access to dialog implementation details.
    On Error GoTo ErrHandler
    Dim formName As String
    formName = ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Report_Error_Dialog")
    If CurrentProject.AllForms(formName).IsLoaded Then DoCmd.Close acForm, formName, acSaveNo
    ReportErrorDialogCancel = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    CoreShowErrorMessage "Could not close the Report Error form." & vbCrLf & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, vbCritical, "Report Error", "ReportErrorDialogCancel", "UI"
    ReportErrorDialogCancel = False
End Function

Private Function ReportErrorCaptureAccessWindow(ByVal bitmapPath As String, ByRef captureWidth As Long, ByRef captureHeight As Long) As String
    On Error GoTo ErrHandler
    Dim windowHandle As LongPtr
    Dim windowRect As TReportErrorRect
    Dim windowDC As LongPtr
    Dim memoryDC As LongPtr
    Dim bitmapHandle As LongPtr
    Dim previousObject As LongPtr
    Dim pictureObject As stdole.IPictureDisp
    Dim pngPath As String
    windowHandle = Application.hWndAccessApp
    If windowHandle = 0 Then Err.Raise vbObjectError + 9133, "ReportErrorCaptureAccessWindow", "The Access application window handle is unavailable."
    If ReportErrorGetWindowRect(windowHandle, windowRect) = 0 Then Err.Raise vbObjectError + 9134, "ReportErrorCaptureAccessWindow", "The Access application window bounds could not be read."
    captureWidth = windowRect.Right - windowRect.Left
    captureHeight = windowRect.Bottom - windowRect.Top
    If captureWidth <= 0 Or captureHeight <= 0 Then Err.Raise vbObjectError + 9135, "ReportErrorCaptureAccessWindow", "The Access application window has invalid capture dimensions."
    windowDC = ReportErrorGetWindowDC(windowHandle)
    If windowDC = 0 Then Err.Raise vbObjectError + 9136, "ReportErrorCaptureAccessWindow", "The Access application window device context could not be acquired."
    memoryDC = ReportErrorCreateCompatibleDC(windowDC)
    If memoryDC = 0 Then Err.Raise vbObjectError + 9137, "ReportErrorCaptureAccessWindow", "A compatible screenshot device context could not be created."
    bitmapHandle = ReportErrorCreateCompatibleBitmap(windowDC, captureWidth, captureHeight)
    If bitmapHandle = 0 Then Err.Raise vbObjectError + 9138, "ReportErrorCaptureAccessWindow", "A compatible screenshot bitmap could not be created."
    previousObject = ReportErrorSelectObject(memoryDC, bitmapHandle)
    If previousObject = 0 Then Err.Raise vbObjectError + 9139, "ReportErrorCaptureAccessWindow", "The screenshot bitmap could not be selected into the capture context."
    If ReportErrorBitBlt(memoryDC, 0, 0, captureWidth, captureHeight, windowDC, 0, 0, REPORT_ERROR_RASTER_COPY Or REPORT_ERROR_CAPTURE_LAYERED_WINDOWS) = 0 Then Err.Raise vbObjectError + 9140, "ReportErrorCaptureAccessWindow", "The Access application window could not be copied into the screenshot bitmap."
    ReportErrorSelectObject memoryDC, previousObject
    previousObject = 0
    Set pictureObject = ReportErrorPictureFromBitmap(bitmapHandle)
    bitmapHandle = 0
    If Len(Dir$(bitmapPath)) > 0 Then Kill bitmapPath
    SavePicture pictureObject, bitmapPath
    Set pictureObject = Nothing
    pngPath = Left$(bitmapPath, Len(bitmapPath) - 4) & ".png"
    If ReportErrorTryConvertBitmapToPng(bitmapPath, pngPath) Then
        ReportErrorDeleteTemporaryFile bitmapPath
        ReportErrorCaptureAccessWindow = pngPath
    Else
        ReportErrorCaptureAccessWindow = bitmapPath
    End If
    ReportErrorDeleteDC memoryDC
    memoryDC = 0
    ReportErrorReleaseDC windowHandle, windowDC
    windowDC = 0
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If previousObject <> 0 And memoryDC <> 0 Then ReportErrorSelectObject memoryDC, previousObject
    If bitmapHandle <> 0 Then ReportErrorDeleteObject bitmapHandle
    If memoryDC <> 0 Then ReportErrorDeleteDC memoryDC
    If windowDC <> 0 Then ReportErrorReleaseDC windowHandle, windowDC
    Set pictureObject = Nothing
    On Error GoTo 0
    Err.Raise savedErrNumber, "ReportErrorCaptureAccessWindow", "Could not capture the Access application window." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function ReportErrorPictureFromBitmap(ByVal bitmapHandle As LongPtr) As stdole.IPictureDisp
    Dim Description As TReportErrorPictureDescription
    Dim interfaceID As TReportErrorGuid
    Dim resultCode As Long
    Description.Size = LenB(Description)
    Description.PictureType = REPORT_ERROR_PICTURE_TYPE_BITMAP
    Description.bitmapHandle = bitmapHandle
    interfaceID.Data1 = &H7BF80981
    interfaceID.Data2 = &HBF32
    interfaceID.Data3 = &H101A
    interfaceID.Data4(0) = &H8B
    interfaceID.Data4(1) = &HBB
    interfaceID.Data4(2) = &H0
    interfaceID.Data4(3) = &HAA
    interfaceID.Data4(4) = &H0
    interfaceID.Data4(5) = &H30
    interfaceID.Data4(6) = &HC
    interfaceID.Data4(7) = &HAB
    resultCode = ReportErrorOleCreatePictureIndirect(Description, interfaceID, True, ReportErrorPictureFromBitmap)
    If resultCode <> 0 Then Err.Raise vbObjectError + 9141, "ReportErrorPictureFromBitmap", "The captured bitmap could not be converted into an OLE picture. Result: " & CStr(resultCode)
    If ReportErrorPictureFromBitmap Is Nothing Then Err.Raise vbObjectError + 9142, "ReportErrorPictureFromBitmap", "The captured bitmap conversion returned no OLE picture."
End Function

Private Function ReportErrorTryConvertBitmapToPng(ByVal bitmapPath As String, ByVal pngPath As String) As Boolean
    On Error GoTo ExitHere
    Dim imageFile As Object
    Dim imageProcess As Object
    Dim convertedImage As Object
    If Len(Dir$(pngPath)) > 0 Then Kill pngPath
    Set imageFile = CreateObject("WIA.ImageFile")
    imageFile.LoadFile bitmapPath
    Set imageProcess = CreateObject("WIA.ImageProcess")
    imageProcess.Filters.Add imageProcess.FilterInfos("Convert").FilterID
    imageProcess.Filters(1).properties("FormatID").value = REPORT_ERROR_PNG_FORMAT_ID
    Set convertedImage = imageProcess.Apply(imageFile)
    convertedImage.SaveFile pngPath
    ReportErrorTryConvertBitmapToPng = (Len(Dir$(pngPath)) > 0)
ExitHere:
    If ReportErrorTryConvertBitmapToPng = False Then ReportErrorDeleteTemporaryFile pngPath
    Set convertedImage = Nothing
    Set imageProcess = Nothing
    Set imageFile = Nothing
End Function

Private Function ReportErrorTemporaryFilePath(ByVal extensionText As String) As String
    Dim folderPath As String
    Dim fileName As String
    folderPath = Environ$("TEMP")
    If Len(Trim$(folderPath)) = 0 Then folderPath = CurrentProject.Path
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    Randomize
    fileName = "ADC_Report_Error_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & Right$("000000" & CStr(Int(999999 * Rnd) + 1), 6) & extensionText
    ReportErrorTemporaryFilePath = folderPath & fileName
End Function

Private Sub ReportErrorDeleteTemporaryFile(ByVal filePath As String)
    On Error Resume Next
    If Len(Trim$(filePath)) > 0 Then
        If Len(Dir$(filePath)) > 0 Then Kill filePath
    End If
    On Error GoTo 0
End Sub

' Purpose: Performs prompt object deletion scope by coordinating `ManagedObjectPhysicalName` and `FormattedElapsedSeconds` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   selectorExistedInitially (Output Optional; Boolean): Set True when the procedure determines that selector existed initially applies; set False when that condition is not met. Default: False.
'   preparationShowSeconds (Output Optional; Double): Receives the resolved preparation show seconds produced from this procedure's validated inputs and current ADC state. Default: 0.
'   userInteractionSeconds (Output Optional; Double): Receives the resolved user interaction seconds produced from this procedure's validated inputs and current ADC state. Default: 0.
'   selectedScopeText (Output Optional; String): Receives the resolved selected scope text produced from this procedure's validated inputs and current ADC state. Default: vbNullString.
' Returns: ObjectDeletionScope: Returns the object/collection representing the deletion scope selected or confirmed by the user/caller.
Public Function PromptObjectDeletionScope(Optional ByRef selectorExistedInitially As Boolean = False, Optional ByRef preparationShowSeconds As Double = 0, Optional ByRef userInteractionSeconds As Double = 0, Optional ByRef selectedScopeText As String = vbNullString) As ObjectDeletionScope
    'System Macro #65 modal scope selector. Setup owns creation/currentness; operational launch only opens the existing Setup-created selector and never validates, repairs, or recreates it.
    On Error GoTo ErrHandler
    Dim formName As String
    Dim stageTimer As Single
    Dim openTimer As Single
    Dim openElapsed As Double
    Dim initialFormExists As Boolean
    Dim userInteractionElapsed As Double
    Dim selectionText As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    mDeleteObjectScopeChoice = DELETE_SCOPE_NOT_SELECTED
    mDeleteObjectScopeCallbackReached = False
    mDeleteScopePromptStartTimer = Timer
    mDeleteScopeDialogReadyTimer = 0
    mDeleteScopeDialogReady = False
    mDeleteScopePreparationElapsed = 0
    stageTimer = Timer
    formName = ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Delete_ADC_Objects")
    LogTraceMessage TRACE_LEVEL_INFO, "Performance", "PromptObjectDeletionScope", "Macro 65 selector physical-name resolution completed.", "Form=" & formName & "; elapsed=" & FormattedElapsedSeconds(ElapsedSecondsFromTimer(stageTimer)), "Macro 65 selector name", 0, vbNullString, vbNullString, formName, "Form", "HostDb", "SystemMacro65"
    stageTimer = Timer
    initialFormExists = SharedDao_AccessObjectExists(GetHostDb(), acForm, formName)
    selectorExistedInitially = initialFormExists
    LogTraceMessage TRACE_LEVEL_INFO, "Performance", "PromptObjectDeletionScope", "Macro 65 deletion dialog form existence check completed.", "Exists=" & CStr(initialFormExists) & "; elapsed=" & FormattedElapsedSeconds(ElapsedSecondsFromTimer(stageTimer)), "Macro 65 form existence check", 0, vbNullString, vbNullString, formName, "Form", "HostDb", "SystemMacro65"
    If initialFormExists = False Then
        LogTraceMessage TRACE_LEVEL_ERROR, "DeleteObjects", "PromptObjectDeletionScope", "Setup-owned Delete ADC Objects selector is missing; operational recreation is not permitted.", "Form=" & formName & "; Currentness check=False; Selector recreation=False; Required owner=Setup generated-form lifecycle", "Macro 65 selector missing", vbObjectError + 9149, "PromptObjectDeletionScope", "The Setup-created Delete ADC Objects selection dialog is missing.", formName, "Form", "HostDb", "SystemMacro65"
        Err.Raise vbObjectError + 9149, "PromptObjectDeletionScope", "The Delete ADC Objects selection dialog is missing. Run ADC Setup to restore the generated forms." & vbCrLf & "Form: " & formName
    End If
    LogTraceMessage TRACE_LEVEL_INFO, "Performance", "PromptObjectDeletionScope", "Macro 65 selector launch path ready.", "Selector existed initially=True; currentness check=False; selector recreation=False; generic generated-form signature check=False; full System Macro definition lookup used=False; Setup Part definition cache used=False", "Macro 65 selector launch path", 0, vbNullString, vbNullString, formName, "Form", "HostDb", "SystemMacro65"
    openTimer = Timer
    DoCmd.OpenForm formName, acNormal, , , , acDialog
    openElapsed = ElapsedSecondsFromTimer(openTimer)
    If mDeleteScopeDialogReady Then userInteractionElapsed = ElapsedSecondsFromTimer(mDeleteScopeDialogReadyTimer)
    If mDeleteObjectScopeCallbackReached = False Then
        LogTraceMessage TRACE_LEVEL_ERROR, "DeleteObjects", "PromptObjectDeletionScope", "Delete ADC Objects selector closed without a registered callback result.", "Form=" & formName & "; Stored selector value=" & CStr(CLng(mDeleteObjectScopeChoice)) & "; Callback reached=False; Form closed=True; Dispatch skipped=True", "Macro 65 selector contract", vbObjectError + 9150, "PromptObjectDeletionScope", "No valid deletion scope callback was registered.", formName, "Form", "HostDb", "SystemMacro65"
        Err.Raise vbObjectError + 9150, "PromptObjectDeletionScope", "The Delete ADC Objects selection dialog closed without returning a valid deletion scope. No objects were deleted."
    End If
    Select Case mDeleteObjectScopeChoice
        Case DELETE_SCOPE_TABLES: selectionText = "Tables"
        Case DELETE_SCOPE_QUERIES: selectionText = "Queries"
        Case DELETE_SCOPE_GENERATED_FORMS: selectionText = "Generated Forms"
        Case DELETE_SCOPE_MACROS: selectionText = "Macros"
        Case DELETE_SCOPE_REPORTS: selectionText = "Reports"
        Case DELETE_SCOPE_ALL: selectionText = "All"
        Case Else: selectionText = "Cancel"
    End Select
    LogTraceMessage TRACE_LEVEL_INFO, "Performance", "PromptObjectDeletionScope", "Macro 65 scope selected.", "Selection=" & selectionText & "; DoCmd.OpenForm/dialog elapsed=" & FormattedElapsedSeconds(openElapsed) & "; user interaction elapsed=" & FormattedElapsedSeconds(userInteractionElapsed), "Macro 65 user interaction", 0, vbNullString, vbNullString, formName, "Form", "HostDb", "SystemMacro65"
    preparationShowSeconds = mDeleteScopePreparationElapsed
    userInteractionSeconds = userInteractionElapsed
    selectedScopeText = selectionText
    PromptObjectDeletionScope = mDeleteObjectScopeChoice
    Exit Function
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "PromptObjectDeletionScope", "Could not display the Delete ADC Objects scope dialog." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

' Purpose: Evaluates the readiness conditions for delete object scope dialog and reports whether all required conditions are currently satisfied.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters: None.
' Returns: Boolean: True when delete object scope dialog is ready for the next dependent operation; False when a prerequisite is not yet satisfied.
Public Function DeleteObjectScopeDialogReady() As Boolean
    'Generated Delete ADC Objects form OnLoad callback. Trace writes remain part of launch preparation; user-decision timing starts only after they complete.
    Dim preparationElapsed As Double
    preparationElapsed = ElapsedSecondsFromTimer(mDeleteScopePromptStartTimer)
    LogTraceMessage TRACE_LEVEL_INFO, "Performance", "DeleteObjectScopeDialogReady", "Macro 65 form open/show operation completed.", "Preparation elapsed=" & FormattedElapsedSeconds(preparationElapsed), "Macro 65 form show", 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Delete_ADC_Objects"), "Form", "HostDb", "SystemMacro65"
    LogTraceMessage TRACE_LEVEL_INFO, "Performance", "DeleteObjectScopeDialogReady", "Macro 65 scope dialog became ready for user interaction.", "Preparation elapsed=" & FormattedElapsedSeconds(preparationElapsed), "Macro 65 dialog ready", 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Delete_ADC_Objects"), "Form", "HostDb", "SystemMacro65"
    mDeleteScopeDialogReadyTimer = Timer
    mDeleteScopePreparationElapsed = ElapsedSecondsFromTimer(mDeleteScopePromptStartTimer)
    mDeleteScopeDialogReady = True
    DeleteObjectScopeDialogReady = True
End Function

' Purpose: Selects delete object scope for selected scope.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   selectedScope (Input; Long): Scope value limiting which selected items or operations are included in this call.
' Returns: Boolean: True on the function's affirmative/success branch for select delete object scope; False on its explicit negative, unavailable, rejected, or handled-failure branch.
Public Function SelectDeleteObjectScope(ByVal selectedScope As Long) As Boolean
    'Expression callback used only by the canonical System Macro #65 modal choice form. The callback records authoritative state before the form closes so callback failure cannot masquerade as Cancel.
    Dim formName As String
    Select Case selectedScope
        Case DELETE_SCOPE_TABLES, DELETE_SCOPE_QUERIES, DELETE_SCOPE_GENERATED_FORMS, DELETE_SCOPE_MACROS, DELETE_SCOPE_REPORTS, DELETE_SCOPE_ALL, DELETE_SCOPE_CANCEL
            mDeleteObjectScopeChoice = selectedScope
            mDeleteObjectScopeCallbackReached = True
        Case Else
            mDeleteObjectScopeChoice = DELETE_SCOPE_NOT_SELECTED
            mDeleteObjectScopeCallbackReached = False
    End Select
    formName = ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Delete_ADC_Objects")
    On Error Resume Next
    If SharedAccess_ObjectIsOpen(acForm, formName) Then DoCmd.Close acForm, formName, acSaveNo
    On Error GoTo 0
    SelectDeleteObjectScope = True
End Function

' Purpose: Prepares generated form deletion plan for include delete objects selector, build scope.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   plannedCountResult (Output Optional; Long): Receives the planned count result produced by the operation so the caller can inspect the outcome without re-running the work. Default: 0.
'   includeDeleteObjectsSelector (Input Optional; Boolean): True includes delete objects selector in this operation; False excludes that optional content while leaving the rest of the operation unchanged.
'   Default: False.
'   buildScope (Input Optional; String): Scope value limiting which build items or operations are included in this call. Default: FORM_BUILD_SCOPE_ALL.
' Returns: Collection: Returns the prepared deletion-plan object/collection describing the exact managed objects eligible for the requested deletion scope; absence follows the routine's existing no-work
'   representation.
Public Function PrepareGeneratedFormDeletionPlan(Optional ByRef plannedCountResult As Long = 0, Optional ByVal includeDeleteObjectsSelector As Boolean = False, Optional ByVal buildScope As String = FORM_BUILD_SCOPE_ALL) As Collection
    'UI-builder authority for one operation-scoped generated-form deletion plan derived directly from DEFAULTS.CanonicalGeneratedFormNamingDefinitions() and filtered by the existing generated-form build-scope
    'classifier. Generated-Forms cleanup preserves the Setup-created Delete ADC Objects operational selector by default; terminal All deletion explicitly includes it. Full generated-form implementation definitions
    'are intentionally not constructed for deletion planning. The Report Error form remains last when it belongs to the selected scope so diagnostics stay available through the rest of deletion.
    On Error GoTo ErrHandler
    Dim namingDefinitions As Collection
    Dim namingDefinition As Variant
    Dim persistentItem As Variant
    Dim deletionNames As Collection
    Dim formName As String
    Dim logicalFormID As String
    Dim reportErrorFormName As String
    Dim normalizedBuildScope As String
    Dim planTimer As Single
    plannedCountResult = 0
    RequireNoUnfinishedUIRecovery "PrepareGeneratedFormDeletionPlan"
    If CodeDbIsHostDb() = False Then Err.Raise vbObjectError + 6698, "PrepareGeneratedFormDeletionPlan", "Generated-form deletion planning requires CodeDb and HostDb to be the same database."
    planTimer = Timer
    normalizedBuildScope = NormalizeChatGPTBuildScope(buildScope)
    Set namingDefinitions = CanonicalGeneratedFormNamingDefinitions()
    ' INTENTIONAL NAME-ONLY PROJECTION: deletionNames is derived directly from the Defaults-owned canonical generated-form naming roster. Deletion planning does not construct builder implementation metadata or the
    ' full managed-object definition registry.
    Set deletionNames = New Collection
    For Each namingDefinition In namingDefinitions
        SetupControlPump
        logicalFormID = CStr(namingDefinition(GENERATED_FORM_NAMING_LOGICAL_ID))
        formName = ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, logicalFormID)
        If GeneratedFormIsInBuildScope(formName, normalizedBuildScope) Then
            If StrComp(logicalFormID, CANON_GENERATED_FORM_LOGICAL_ID_REPORT_ERROR_DIALOG, vbTextCompare) = 0 Then
                reportErrorFormName = formName
            ElseIf StrComp(logicalFormID, CANON_GENERATED_FORM_LOGICAL_ID_DELETE_OBJECTS, vbTextCompare) = 0 And includeDeleteObjectsSelector = False Then
                LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "PrepareGeneratedFormDeletionPlan", "DELETE PLAN EXCLUDE", "Name=" & formName & "; Classification=OperationalSelector; Reason=Preserved for Generated Forms cleanup", "Setup-created Delete ADC Objects operational selector preserved so Macro 65 remains usable after Generated Forms cleanup.", 0, vbNullString, vbNullString, formName, "Form", "CodeDb", "DeleteForms"
            Else
                deletionNames.Add formName
            End If
        End If
    Next namingDefinition
    If Len(reportErrorFormName) > 0 Then deletionNames.Add reportErrorFormName
    For Each persistentItem In Array(CANON_PERSISTENT_FORM_NAME_CODEDB_CUSTOM_MACROS, CANON_PERSISTENT_FORM_NAME_HOSTDB_CUSTOM_MACROS)
        LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "PrepareGeneratedFormDeletionPlan", "DELETE PLAN EXCLUDE", "Name=" & CStr(persistentItem) & "; Classification=" & DEFAULT_CATALOG_OWNER_PERSISTENT_GENERATED, "Persistent Generated form preserved outside the ordinary generated-form deletion roster.", 0, vbNullString, vbNullString, CStr(persistentItem), "Form", "CodeDb", "DeleteForms"
    Next persistentItem
    plannedCountResult = deletionNames.count
    LogTraceMessage TRACE_LEVEL_INFO, "Performance", "PrepareGeneratedFormDeletionPlan", "Macro 65 generated-form deletion name-only plan prepared.", "Canonical naming rows read=" & CStr(namingDefinitions.count) & "; build scope=" & normalizedBuildScope & "; deletion names retained=" & CStr(deletionNames.count) & "; Delete Objects selector included=" & CStr(includeDeleteObjectsSelector) & "; GeneratedFormDefinitions constructed=False; full managed-object definition registry constructed=False; elapsed=" & FormattedElapsedSeconds(ElapsedSecondsFromTimer(planTimer)), "Macro 65 generated-form deletion plan", 0, vbNullString, vbNullString, vbNullString, "FormSet", "CodeDb", "SystemMacro65"
    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "PrepareGeneratedFormDeletionPlan", "DELETE PLAN", "Scope=Generated Forms; BuildScope=" & normalizedBuildScope & "; Planned=" & CStr(deletionNames.count), "Macro 65 generated-form deletion plan", 0, vbNullString, vbNullString, vbNullString, "FormSet", "CodeDb", "DeleteForms"
    If deletionNames.count = 0 Then LogTraceMessage TRACE_LEVEL_WARNING, "DeleteObjects", "PrepareGeneratedFormDeletionPlan", "DELETE PLAN EMPTY", "Scope=Generated Forms; BuildScope=" & normalizedBuildScope, "Macro 65 generated-form deletion plan", 0, vbNullString, vbNullString, vbNullString, "FormSet", "CodeDb", "DeleteForms"
    Set PrepareGeneratedFormDeletionPlan = deletionNames
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "PrepareGeneratedFormDeletionPlan", "Could not prepare the authoritative generated-form deletion plan." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

' Purpose: Deletes generated forms for confirm destructive action, prepared deletion plan.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
'   confirmDestructiveAction (Input Optional; Boolean): When True, enables or requests destructive action for this call; when False, the routine follows its normal non-forced/default path. Default: True.
'   plannedCountResult (Output Optional; Long): Receives the planned count result produced by the operation so the caller can inspect the outcome without re-running the work. Default: 0.
'   attemptedCountResult (Output Optional; Long): Receives the attempted count result produced by the operation so the caller can inspect the outcome without re-running the work. Default: 0.
'   deletedCountResult (Output Optional; Long): Receives the deleted count result produced by the operation so the caller can inspect the outcome without re-running the work. Default: 0.
'   skippedCountResult (Output Optional; Long): Receives the skipped count result produced by the operation so the caller can inspect the outcome without re-running the work. Default: 0.
'   failedCountResult (Output Optional; Long): Receives the failed count result produced by the operation so the caller can inspect the outcome without re-running the work. Default: 0.
'   operationSucceededResult (Output Optional; Boolean): Receives the operation succeeded result produced by the operation so the caller can inspect the outcome without re-running the work. Default: False.
'   preparedDeletionPlan (Input Optional; Variant): Previously prepared deletion-plan payload describing the exact objects/actions approved for deletion; omitted means the routine must follow its normal
'   prepare/resolve path.
'   buildScope (Input Optional; String): Scope value limiting which build items or operations are included in this call. Default: FORM_BUILD_SCOPE_ALL.
' Returns: None (Sub). Results are returned through the documented ByRef parameters and/or the routine's database, UI, or module-state side effects.
Public Sub DeleteGeneratedForms(Optional ByVal showMessage As Boolean = True, Optional ByVal confirmDestructiveAction As Boolean = True, Optional ByRef plannedCountResult As Long = 0, Optional ByRef attemptedCountResult As Long = 0, Optional ByRef deletedCountResult As Long = 0, Optional ByRef skippedCountResult As Long = 0, Optional ByRef failedCountResult As Long = 0, Optional ByRef operationSucceededResult As Boolean = False, Optional ByVal preparedDeletionPlan As Variant, Optional ByVal buildScope As String = FORM_BUILD_SCOPE_ALL)
    'Intent: Executes the UI-builder-owned generated-form deletion plan prepared from DEFAULTS.CanonicalGeneratedFormNamingDefinitions() for the selected generated-form scope. Ordinary Generated Forms cleanup
    'preserves the Delete ADC Objects operational selector; a supplied plan may include it and is narrowed through the same authoritative scope classifier when a non-All scope is requested.
    On Error GoTo ErrHandler
    RequireNoUnfinishedUIRecovery "DeleteGeneratedForms"
    Dim item As Variant
    Dim deletedCount As Long
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim messageText As String
    Dim response As VbMsgBoxResult
    Dim traceRunOwnedByProcedure As Boolean
    Dim deletionNames As Collection
    Dim normalizedBuildScope As String
    Dim trackedProcessStarted As Boolean
    plannedCountResult = 0
    attemptedCountResult = 0
    deletedCountResult = 0
    skippedCountResult = 0
    failedCountResult = 0
    operationSucceededResult = False
    normalizedBuildScope = NormalizeChatGPTBuildScope(buildScope)
    traceRunOwnedByProcedure = BeginUITraceRunIfNeeded("UI.DeleteForms")
    If CodeDbIsHostDb() = False Then
        messageText = x__UTIL_ADC_UI_FORMS.GeneratedFormCodeDbOnlyMessage("Deleting ChatGPT UI Forms")
        LogTraceMessage "Info", "UI", "DeleteGeneratedForms", "Deleting ADC generated forms skipped because CodeDb and HostDb are different.", messageText, vbNullString, 0, vbNullString, vbNullString, vbNullString, "FormSet", "CodeDb", "DeleteForms"
        If showMessage Then CoreShowMessage messageText, vbInformation, DEFAULT_UI_MESSAGE_TITLE, "ShowADCMessage", "UI"
        CompleteUITraceRunIfOwned traceRunOwnedByProcedure, "Delete forms skipped."
        Exit Sub
    End If
    If confirmDestructiveAction Then
        response = CoreShowMessageResult("Delete generated Access Development Console forms in scope [" & normalizedBuildScope & "]?" & vbCrLf & vbCrLf & "This deletes the canonical generated-form set in the selected scope except the Delete ADC Objects operational selector, which remains available so this command can be run again. If any affected generated forms are open, Yes also saves pending changes, closes only those forms, and verifies closure before deletion. The deleted forms can be rebuilt by running BuildGeneratedForms." & vbCrLf & vbCrLf & "Tables, queries, system macros, and reports are not deleted." & vbCrLf & vbCrLf & "Continue?", vbCritical Or vbYesNo Or vbDefaultButton2, "Delete Generated Forms", "ShowADCMessageResult", "UI")
        If response <> vbYes Then
            messageText = "Delete generated Access Development Console forms was cancelled before any forms were deleted."
            LogTraceMessage "Info", "UI", "DeleteGeneratedForms", messageText, "User response: " & CStr(response), vbNullString, 0, vbNullString, vbNullString, vbNullString, "FormSet", "CodeDb", "DeleteForms"
            If showMessage Then CoreShowMessage messageText, vbInformation, "Delete Generated Forms", "ShowADCMessage", "UI"
            MarkSystemCommandCancelled
            CompleteUITraceRunIfOwned traceRunOwnedByProcedure, "Delete forms cancelled.", False
            Exit Sub
        End If
    End If
    If isMissing(preparedDeletionPlan) Then
        Set deletionNames = PrepareGeneratedFormDeletionPlan(plannedCountResult, False, normalizedBuildScope)
    ElseIf IsObject(preparedDeletionPlan) Then
        Set deletionNames = preparedDeletionPlan
        If deletionNames Is Nothing Then Err.Raise vbObjectError + 6699, "DeleteGeneratedForms", "The supplied generated-form deletion plan is Nothing."
        If StrComp(normalizedBuildScope, FORM_BUILD_SCOPE_ALL, vbTextCompare) <> 0 Then Set deletionNames = FilterGeneratedFormNamesForScope(deletionNames, normalizedBuildScope)
        plannedCountResult = deletionNames.count
        LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "DeleteGeneratedForms", "Prepared deletion plan reused", "Scope=Generated Forms; BuildScope=" & normalizedBuildScope & "; Planned=" & CStr(plannedCountResult), "Macro 65 generated-form deletion plan reuse", 0, vbNullString, vbNullString, vbNullString, "FormSet", "CodeDb", "DeleteForms"
    Else
        Err.Raise vbObjectError + 6700, "DeleteGeneratedForms", "The supplied generated-form deletion plan is not an object."
    End If
    LogTraceMessage "Info", "UI", "DeleteGeneratedForms", "Deleting planned generated Access Development Console form set.", "Build scope: " & normalizedBuildScope & "; destructive confirmation required: " & CStr(confirmDestructiveAction) & "; prepared plan count: " & CStr(plannedCountResult), vbNullString, 0, vbNullString, vbNullString, vbNullString, "FormSet", "CodeDb", "DeleteForms"
    If SharedAccess_PrepareOpenObjectsForDesignChange(acForm, deletionNames, "deleting the planned generated Access Development Console form set", confirmDestructiveAction) = False Then
        MarkSystemCommandCancelled
        CompleteUITraceRunIfOwned traceRunOwnedByProcedure, "Delete forms cancelled during open-form preparation.", False
        Exit Sub
    End If
    BeginTrackedProcess "Deleting generated Access Development Console forms"
    trackedProcessStarted = True
    For Each item In deletionNames
        DoEvents
        On Error Resume Next
        Application.SysCmd acSysCmdSetStatus, "Deleting generated form: " & CStr(item)
        On Error GoTo ErrHandler
        attemptedCountResult = attemptedCountResult + 1
        LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "DeleteGeneratedForms", "DELETE ATTEMPT", "Type=Form; Name=" & CStr(item), "Macro 65 generated-form delete attempt", 0, vbNullString, vbNullString, CStr(item), "Form", "CodeDb", "DeleteForms"
        If SharedDao_AccessObjectExists(GetHostDb(), acForm, CStr(item)) Then
            If SharedAccess_ObjectIsOpen(acForm, CStr(item)) Then Err.Raise vbObjectError + 6697, "DeleteGeneratedForms", "The affected generated form remained open after design-change preparation."
            DoCmd.DeleteObject acForm, CStr(item)
            deletedCount = deletedCount + 1
            deletedCountResult = deletedCount
            LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "DeleteGeneratedForms", "DELETE SUCCESS", "Type=Form; Name=" & CStr(item), "Macro 65 generated-form delete success", 0, vbNullString, vbNullString, CStr(item), "Form", "CodeDb", "DeleteForms"
        Else
            skippedCountResult = skippedCountResult + 1
            LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "DeleteGeneratedForms", "DELETE SKIP", "Type=Form; Name=" & CStr(item) & "; Reason=Not found at deletion time", "Macro 65 generated-form delete skip", 0, vbNullString, vbNullString, CStr(item), "Form", "CodeDb", "DeleteForms"
        End If
    Next item
    On Error Resume Next
    Application.SysCmd acSysCmdClearStatus
    On Error GoTo ErrHandler
    RefreshNavigationPaneAfterGrouping
    LogTraceMessage "Info", "UI", "DeleteGeneratedForms", "Deleted planned generated Access Development Console form set.", "Deleted count: " & CStr(deletedCount), vbNullString, 0, vbNullString, vbNullString, vbNullString, "FormSet", "CodeDb", "DeleteForms"
    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "DeleteGeneratedForms", "DELETE ADC OBJECTS COMPLETE", "Planned=" & CStr(plannedCountResult) & "; Attempted=" & CStr(attemptedCountResult) & "; Deleted=" & CStr(deletedCountResult) & "; Skipped=" & CStr(skippedCountResult) & "; Failed=" & CStr(failedCountResult), "Macro 65 generated-form deletion summary", 0, vbNullString, vbNullString, vbNullString, "FormSet", "CodeDb", "DeleteForms"
    ReportCoreActionSuccess "DeleteGeneratedForms", "UI", "Generated " & DEFAULT_UTILITY_SHORT_NAME & " forms deleted successfully.", "Deleted " & CStr(deletedCount) & " generated Access Development Console forms.", "Delete generated forms", DEFAULT_CHATGPT_UI_DISPLAY_PREFIX, "FormSet", "CodeDb", "UI", showMessage, DEFAULT_UI_MESSAGE_TITLE, vbInformation, False, "UI", True, visibleStatusMessage:="Generated forms deleted successfully."
    EndTrackedProcess
    trackedProcessStarted = False
    operationSucceededResult = True
    CompleteUITraceRunIfOwned traceRunOwnedByProcedure, "Deleted planned generated Access Development Console form set."
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    operationSucceededResult = False
    On Error Resume Next
    Application.SysCmd acSysCmdClearStatus
    If trackedProcessStarted Then
        CaptureTrackedProcessSummary
        EndTrackedProcess
        trackedProcessStarted = False
    End If
    On Error GoTo 0
    messageText = "Could not delete the planned generated Access Development Console form set." & vbCrLf & vbCrLf & "Procedure: DeleteGeneratedForms" & vbCrLf & "Current form: " & CStr(item) & vbCrLf & "Deleted count before failure: " & CStr(deletedCount) & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description:" & vbCrLf & savedErrDescription & TableDiagnosticsLocationForMessage() & ProcessRunSummaryForMessage("DeleteGeneratedForms")
    failedCountResult = failedCountResult + 1
    LogTraceMessage TRACE_LEVEL_ERROR, "DeleteObjects", "DeleteGeneratedForms", "DELETE FAILED", "Type=Form; Name=" & CStr(item) & "; Error=" & CStr(savedErrNumber) & "; Description=" & savedErrDescription, "Macro 65 generated-form delete failure", savedErrNumber, savedErrSource, savedErrDescription, CStr(item), "Form", "CodeDb", "DeleteForms"
    LogTraceMessage "Error", "UI", "DeleteGeneratedForms", "Could not delete planned generated Access Development Console form set.", messageText, CStr(item), savedErrNumber, savedErrSource, savedErrDescription, CStr(item), "Form", "CodeDb", "DeleteForms"
    If showMessage Then
        CoreShowErrorMessage messageText, vbCritical, DEFAULT_UI_MESSAGE_TITLE, "DeleteGeneratedForms", "UI"
        CompleteUITraceRunIfOwned traceRunOwnedByProcedure, "Delete forms failed.", False
        Exit Sub
    End If
    CompleteUITraceRunIfOwned traceRunOwnedByProcedure, "Delete forms failed.", False
    Err.Raise savedErrNumber, "DeleteGeneratedForms", messageText
End Sub

' Purpose: Runs Module Guard verification from the generated form entry point and presents the existing verification result to the user.
' Public role: Form-code/event-expression entry point. Generated or managed Access forms can invoke this Public routine from form code or an event expression.
' Parameters: None.
' Returns: Boolean: True when chat gpt verify module guard from generated form completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or
'   not-applicable path.
Public Function ChatGPT_VerifyModuleGuardFromGeneratedForm() As Boolean
    ' Wrapper purpose: Generated-form callback exposing Module Guard verification through an Access expression-compatible Boolean Function and handling the disabled-guard case.
    ' Why retained: Retained because generated controls cannot call the underlying verification workflow directly and need a stable UI-facing callback contract.
    On Error GoTo ErrHandler
    If ModuleUpdateGuardEnabled() = False Then
        CoreShowMessage "The module update guard is currently disabled." & vbCrLf & vbCrLf & "Use Repair Guard when needed. Verification is available after EnableModuleUpdateGuard is enabled through the Module Guard registration/setup workflow.", vbInformation, "Access Development Console Module Update Guard", "ChatGPT_VerifyModuleGuardFromGeneratedForm", "UI"
        Exit Function
    End If
    Call VerifyModuleUpdateGuard(False)
    ChatGPT_VerifyModuleGuardFromGeneratedForm = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    LogTraceMessage "Error", "UI", "ChatGPT_VerifyModuleGuardFromGeneratedForm", "The generated form could not verify the module update guard.", "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription, "Verify module guard", savedErrNumber, savedErrSource, savedErrDescription, "ModuleUpdateGuard", "Guard", "CodeDb", "GeneratedFormCommand"
    CoreShowErrorMessage "Could not verify the Access Development Console module update guard." & vbCrLf & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Description: " & savedErrDescription, vbCritical, "Access Development Console Module Update Guard", "ChatGPT_VerifyModuleGuardFromGeneratedForm", "UI"
End Function
