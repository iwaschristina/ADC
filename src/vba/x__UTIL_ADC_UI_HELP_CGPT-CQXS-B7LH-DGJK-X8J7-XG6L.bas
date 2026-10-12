Attribute VB_Name = "x__UTIL_ADC_UI_HELP"
' ================================================================================================================================================================================
' ChatGPT Rewrite Tracking Header
' ================================================================================================================================================================================
' ChatGPT Previous Version ID: CGPT-D5HF-T5DE-M9NU-74RJ-WA8A
' ChatGPT Baseline Version ID: CGPT-D5HF-T5DE-M9NU-74RJ-WA8A
' ChatGPT Version ID: CGPT-CQXS-B7LH-DGJK-X8J7-XG6L
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
' STANDARD MODULE: Access Development Console Help UI Utility
' VB_Name: x__UTIL_ADC_UI_HELP
' Purpose: Owns generated Help definitions and form construction plus Help search, navigation, history, related-target handling, and managed-object Help presentation.
' ================================================================================================================================================================================

Option Compare Database
Option Explicit

' ===============================================================================================================================================================================
' MODULE-LEVEL DECLARATIONS
' ===============================================================================================================================================================================

' Setup generated-form Help proof and reuse instrumentation.
Private mSetupGeneratedFormHelpSignature As String
Private mGeneratedFormHelpDefinitionValidationCount As Long
Private mGeneratedFormHelpProofReuseCount As Long

' Help search, list-action profiles, and help-form navigation state.
'Revision token for the generated Help form's required control/interface contract. Change it when that contract changes in a way that should invalidate an existing generated Help form. This is not a change log.
Public Const HELP_FORM_CONTRACT_VERSION As String = "HelpFormContractV1"
'Revision token for fixed generated-Help content. Change it when fixed Help content changes in a way that should invalidate an existing generated Help form. This is not a change log.
Public Const HELP_FIXED_TOPIC_VERSION As String = "HelpFixedTopicsV3"

' Help search limits and help-form navigation state.
Private Const HELP_SEARCH_MAX_RESULTS As Long = 150
Private Const HELP_VALUE_LIST_MAX_LENGTH As Long = 30000
' The active Help form reference is event-scoped; mutable navigation state is stored on the invoking form instance.
Private mChatGPTHelpForm As Access.Form

' ===============================================================================================================================================================================
' PROCEDURES
' ===============================================================================================================================================================================

' =============================================================================================================================================================================
' HELP DEFINITION / VALIDATION / PROOF PROCEDURES
' =============================================================================================================================================================================

Private Sub ChatGPTHelpAddTopic(ByVal rows As Collection, ByVal sectionKey As String, ByVal topicKey As String, ByVal TitleText As String, ByVal bodyText As String, ByVal relatedTarget As String)
    rows.Add Array(sectionKey, topicKey, TitleText, bodyText, relatedTarget)
End Sub

' Purpose: Resolves chat GPT system macro definition index from the supplied/current ADC state, using `isMissing` and `IsObject` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   suppliedSystemMacroDefinitions (Input Optional; Variant): Variant input representing supplied system macro definitions; the routine accepts the caller-provided value in the form required by its
'   existing validation and conversion logic.
' Returns: Object: Returns the resolved object reference. `Nothing` remains the no-object result where the existing routine permits an unresolved or unavailable object.
Public Function ChatGPTSystemMacroDefinitionIndex(Optional ByVal suppliedSystemMacroDefinitions As Variant) As Object
    Dim definitions As Collection
    Dim Definition As Variant
    Dim keyIndex As Object
    Dim systemMacroKey As String
    If isMissing(suppliedSystemMacroDefinitions) = False Then
        If IsObject(suppliedSystemMacroDefinitions) Then Set definitions = suppliedSystemMacroDefinitions
    End If
    If definitions Is Nothing Then Set definitions = GetSystemMacroDefinitions()
    ' INTENTIONAL KEYED LOOKUP: keyIndex provides keyed membership/lookup over source rows or state already read by ChatGPTSystemMacroDefinitionIndex; it is a derived access path, not a second authority.
    Set keyIndex = CreateObject("Scripting.Dictionary")
    keyIndex.CompareMode = vbTextCompare
    For Each Definition In definitions
        systemMacroKey = Trim$(CStr(Definition(0)))
        If Len(systemMacroKey) = 0 Then Err.Raise vbObjectError + 7367, "ChatGPTSystemMacroDefinitionIndex", "A canonical System Macro definition has a blank key."
        If keyIndex.Exists(systemMacroKey) Then Err.Raise vbObjectError + 7368, "ChatGPTSystemMacroDefinitionIndex", "The canonical System Macro definitions contain a duplicate key." & vbCrLf & "SystemMacroKey: " & systemMacroKey
        keyIndex.Add systemMacroKey, Definition
    Next Definition
    Set ChatGPTSystemMacroDefinitionIndex = keyIndex
End Function

Private Function ChatGPTHelpCommandTargetKnown(ByVal commandKey As String, Optional ByVal suppliedSystemMacroDefinitionIndex As Variant) As Boolean
    Dim systemMacroDefinitionIndex As Object
    Dim Definition As Variant
    If Len(commandKey) = 0 Then Exit Function
    If isMissing(suppliedSystemMacroDefinitionIndex) = False Then
        If IsObject(suppliedSystemMacroDefinitionIndex) Then Set systemMacroDefinitionIndex = suppliedSystemMacroDefinitionIndex
    End If
    If systemMacroDefinitionIndex Is Nothing Then
        ChatGPTHelpCommandTargetKnown = (Len(Trim$(SystemMacroCatalogTitleByKey(commandKey, True))) > 0)
        Exit Function
    End If
    If systemMacroDefinitionIndex.Exists(commandKey) = False Then Exit Function
    Definition = systemMacroDefinitionIndex(commandKey)
    ChatGPTHelpCommandTargetKnown = (Len(Trim$(SystemMacroCalculatedTitle(CStr(Definition(1)), CStr(Definition(2)), CStr(Definition(0)), Definition))) > 0)
End Function

' Purpose: Resolves chat GPT help fixed topic signature from the supplied/current ADC state, using `IsObject` and `ChatGPTHelpTopics` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   helpTopics (Input Optional; Variant): Variant input representing help topics; the routine accepts the caller-provided value in the form required by its existing validation and conversion logic.
' Returns: String: Returns the chat gpt help fixed topic signature text value. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function ChatGPTHelpFixedTopicSignature(Optional ByVal helpTopics As Variant) As String
    Dim topics As Collection
    Dim topic As Variant
    Dim payloadText As String
    If IsObject(helpTopics) Then Set topics = helpTopics
    If topics Is Nothing Then Set topics = ChatGPTHelpTopics()
    payloadText = "ADCGeneratedHelpFixedTopicsV1|" & HELP_FIXED_TOPIC_VERSION & "|" & CStr(topics.count)
    For Each topic In topics
        SetupControlPump
        payloadText = payloadText & "|" & CStr(topic(0)) & "|" & CStr(topic(1)) & "|" & CStr(topic(2)) & "|" & CStr(topic(3)) & "|" & CStr(topic(4))
    Next topic
    ChatGPTHelpFixedTopicSignature = AdlerTextHash(payloadText) & ":" & CStr(Len(payloadText))
End Function

Private Function ChatGPTHelpKnownSectionKey(ByVal sectionKey As String) As Boolean
    Select Case UCase$(Trim$(sectionKey))
        Case "START", "SETUP", "CONFIG", "SCAN", "REVIEW", "GENERATE", "HEALTH", "OBJECTS", "FORMS", "RIBBON", "GUARD", "DIAG", "REFS", "ADDINS", "FAQ", "DBREF"
            ChatGPTHelpKnownSectionKey = True
    End Select
End Function

' Purpose: Resolves chat GPT help topics from the supplied/current ADC state, using `isMissing` and `IsObject` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   suppliedSystemMacroDefinitions (Input Optional; Variant): Variant input representing supplied system macro definitions; the routine accepts the caller-provided value in the form required by its
'   existing validation and conversion logic.
'   suppliedSystemMacroDefinitionIndex (Input Optional; Variant): Numeric supplied system macro definition index value used by the routine for identity, ordering, limits, or counts as indicated by the
'   parameter name.
' Returns: Collection: Returns the ordered Collection assembled by the routine; callers receive the object reference and should treat its membership/order according to the routine contract.
Public Function ChatGPTHelpTopics(Optional ByVal suppliedSystemMacroDefinitions As Variant, Optional ByVal suppliedSystemMacroDefinitionIndex As Variant) As Collection
    Dim rows As Collection
    Dim bodyText As String
    Dim addInCommandKeys As Collection
    Dim addInGroupLabel As String
    Dim addInGroupOrder As Long
    Dim addInTabKey As String
    Dim addInTabLabel As String
    Dim helpRibbonGroupLabel As String
    Dim helpRibbonGroupOrder As Long
    Dim helpRibbonTabKey As String
    Dim helpRibbonTabLabel As String
    Dim groupDefinitionIndex As Object
    Dim systemMacroDefinitions As Collection
    Dim systemMacroDefinitionIndex As Object
    ' CANONICAL AUTHORITY: rows owns the in-memory definition set constructed by ChatGPTHelpTopics; downstream consumers derive views from this set rather than create a competing authority.
    Set rows = New Collection
    Set groupDefinitionIndex = SystemRibbonGroupDefinitionIndex()
    If isMissing(suppliedSystemMacroDefinitions) = False Then
        If IsObject(suppliedSystemMacroDefinitions) Then Set systemMacroDefinitions = suppliedSystemMacroDefinitions
    End If
    If isMissing(suppliedSystemMacroDefinitionIndex) = False Then
        If IsObject(suppliedSystemMacroDefinitionIndex) Then Set systemMacroDefinitionIndex = suppliedSystemMacroDefinitionIndex
    End If
    If systemMacroDefinitions Is Nothing Then Set systemMacroDefinitions = x__UTIL_ADC_DEFAULTS.GetSystemMacroDefinitions()
    If systemMacroDefinitionIndex Is Nothing Then Set systemMacroDefinitionIndex = ChatGPTSystemMacroDefinitionIndex(systemMacroDefinitions)
    Set addInCommandKeys = SystemMacroKeysForRibbonPlacement(DEFAULT_SYSTEM_RIBBON_GROUP_ADD_INS, DEFAULT_SYSTEM_RIBBON_GROUP_ADD_INS_SORT, DEFAULT_RIBBON_TAB_KEY_SYSTEM_1, groupDefinitionIndex, systemMacroDefinitions)
    If addInCommandKeys.count > 0 Then
        SystemRibbonPlacementForKey CStr(addInCommandKeys(1)), addInTabKey, addInTabLabel, addInGroupLabel, addInGroupOrder, groupDefinitionIndex, systemMacroDefinitionIndex(CStr(addInCommandKeys(1)))
    End If
    HelpDocumentSystemRibbonPlacement helpRibbonTabKey, helpRibbonTabLabel, helpRibbonGroupLabel, helpRibbonGroupOrder, groupDefinitionIndex
    bodyText = vbNullString
    bodyText = bodyText & DEFAULT_UTILITY_LONG_NAME & " maintains the CodeDb schema, Managed Object Catalog, generated forms, system and Help macros, Navigation Pane placement, the ADC ribbon,"
    bodyText = bodyText & " configuration and candidate workflows, reference and add-in inventories, loaded VBA project analysis, Module Guard evidence, diagnostic history, recovery state, and field/query"
    bodyText = bodyText & " order history."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "This generated Help center combines fixed workflow guidance with a table-backed Database Reference. Use Search to find a task, workflow term, object, field, status, or safety"
    bodyText = bodyText & " boundary."
    ChatGPTHelpAddTopic rows, "START", "HOME", "What " & DEFAULT_UTILITY_LONG_NAME & " Does", bodyText, "COMMAND:" & SYS_MACRO_KEY_RUN_FULL_SETUP
    bodyText = vbNullString
    bodyText = bodyText & "1. Import the complete coordinated module batch."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "2. Run RegisterApprovedBuild through the Module Guard administrative workflow to register the exact approved coordinated build."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "3. Run ADC Setup for authoritative current-state installation reconciliation and verification."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "4. Run passive reference, add-in, or loaded-project scans when updated discovery evidence is needed."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "5. Review generated evidence and candidate rows."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "6. Resolve and validate intended actions."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "7. Generate required macros."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "8. Begin troubleshooting with Health Review, TraceLog, the physical log, Help, and Report Error rather than repeated blind reruns."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Setup and update verify current state and resume from durable boundaries."
    ChatGPTHelpAddTopic rows, "START", "WORKFLOW", "The Normal Workflow", bodyText, "COMMAND:" & SYS_MACRO_KEY_RUN_FULL_SETUP
    bodyText = vbNullString
    bodyText = bodyText & "Fixed workflow topics are maintained by the Help owner in x__UTIL_ADC_UI_HELP. Their complete section, topic, title, body, and related-action definitions contribute to the"
    bodyText = bodyText & " Help form's deterministic build signature. Adding or changing a topic therefore marks Help stale during a full generated-form build."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "The Database Reference remains authoritative for database-specific objects and fields. Missing documentation remains visibly missing; Help does not invent replacement prose at"
    bodyText = bodyText & " runtime."
    ChatGPTHelpAddTopic rows, "START", "CURRENT", "How This Help Stays Current", bodyText, "T:DBREF:INTRO"
    bodyText = vbNullString
    bodyText = bodyText & "Full Setup performs the complete CodeDb reconciliation: managed tables and queries, catalog and field documentation, generated forms, system and Help macros, Navigation Pane"
    bodyText = bodyText & " placement, the standard " & DEFAULT_UTILITY_SHORT_NAME & " ribbon, post-commit verification, recovery-state checks, and final Module Guard and foundation fingerprints."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Existing configuration, user-owned documentation, explicit overrides, and intentionally blank user values are preserved unless the authoritative current repair, reconciliation, rename, normalization, restore, or rebuild operation requires the change."
    ChatGPTHelpAddTopic rows, "SETUP", "FULL", "Run Full ADC Setup", bodyText, "COMMAND:" & SYS_MACRO_KEY_RUN_FULL_SETUP
    bodyText = vbNullString
    bodyText = bodyText & "The current setup contains "
    bodyText = bodyText & CStr(ImplementedSetupPartCount())
    bodyText = bodyText & " substantive parts. A part represents real work or an actual safety boundary; skipped duplicate work remains an operation note rather than a fake part. Every setup status message"
    bodyText = bodyText & " uses Part X of X."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Durable boundaries protect generated forms, catalog changes, Navigation Pane application, ribbon staging and commit, recovery state, Module Guard state, and final fingerprints."
    bodyText = bodyText & " Nested generated-form, Add-In Review, and ribbon work remains subordinate to its parent setup part."
    ChatGPTHelpAddTopic rows, "SETUP", "PARTS", "Setup Parts and Safety Boundaries", bodyText, "COMMAND:" & SYS_MACRO_KEY_RUN_FULL_SETUP
    bodyText = vbNullString
    bodyText = bodyText & "The running setup controller provides Pause and Cancel. Requests can be accepted while setup is running, but they are honored only at safe checkpoints. Pause opens the paused"
    bodyText = bodyText & " controller with Resume and Cancel; Resume restores the running controller. Cancel rolls back the active rollback-tracked part while preserving completed parts."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "An indivisible Access operation may finish before a pending request reaches a checkpoint. The heartbeat repeats the current work and elapsed time after a quiet interval. The"
    bodyText = bodyText & " throttled event pump keeps controls responsive; clicking an unrelated area of the dialog must not be required to restart setup. Add-In Review work must not hide or prematurely close the"
    bodyText = bodyText & " controller."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "If the interface appears frozen, the heartbeat and trace show whether an indivisible operation is still running or a safe checkpoint has not yet been reached."
    ChatGPTHelpAddTopic rows, "SETUP", "RESPONSIVE", "Pause, Resume, Cancel, and Heartbeat", bodyText, "COMMAND:" & SYS_MACRO_KEY_RUN_FULL_SETUP
    bodyText = vbNullString
    bodyText = bodyText & "Managed object names resolve through explicit logical identity, object type, naming category, and canonical naming rules. Generated forms classified ChatGPTUIForm=True use the fixed"
    bodyText = bodyText & " _@ChatGptUI physical object-name suffix; other canonical generated forms keep their unsuffixed physical names. Display captions remain separate from physical object identity."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Physical names are separate from ribbon captions and display labels. The Help_Terms table is " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_TERMINOLOGY_TABLE_LOGICAL_ID) & " and its editable terminology reference query is " & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Terminology_Review") & ". Edits through that updateable query directly update Help_Terms. Tables never inherit ordered system-macro prefixes."
    ChatGPTHelpAddTopic rows, "SETUP", "NAMES", "Managed Object Names", bodyText, "COMMAND:" & SYS_MACRO_KEY_APPLY_OBJECT_NAMING_RULES
    bodyText = vbNullString
    bodyText = bodyText & "Configuration rows define macro identity, displayed title, action type, target, arguments, ownership role, generation behavior, active state, and review status. Candidate rows"
    bodyText = bodyText & " remain separate until accepted."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Use the generated configuration forms for supported edits. Structural ownership and utility-owned defaults should be changed only through their owning"
    bodyText = bodyText & " owning current repair, reconciliation, rename, normalization, restore, or rebuild workflow."
    ChatGPTHelpAddTopic rows, "CONFIG", "ROWS", "Configuration Rows", bodyText, "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Configuration_Detail")
    bodyText = vbNullString
    bodyText = bodyText & "RunCode targets call approved public VBA functions. OpenForm, OpenTable, OpenQuery, OpenReport, RunMacro, and related action types operate on Access objects."
    bodyText = bodyText & " ConfigOwnerDatabaseRole identifies whether CodeDb or HostDb owns the configuration and target resolution."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Validation checks target existence, procedure visibility, required arguments, ambiguity, unsupported object actions, and ownership before macro generation."
    ChatGPTHelpAddTopic rows, "CONFIG", "TARGETS", "Target Routing and Ownership", bodyText, "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Configuration_Detail")
    bodyText = vbNullString
    bodyText = bodyText & "Candidate scanning discovers supported public procedures, form actions, and Access objects that may be useful as commands. Reference scans, Reference Inventory add-in inventory, and Loaded VBA Analysis"
    bodyText = bodyText & " loaded-project analysis are separate read-only evidence workflows."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "A scan does not approve candidates, generate macros, change references, connect add-ins, load projects, or modify external source."
    ChatGPTHelpAddTopic rows, "SCAN", "PURPOSE", "When to Scan", bodyText, "COMMAND:" & SYS_MACRO_KEY_SCAN_RUNNER_MACRO_CANDIDATES
    bodyText = vbNullString
    bodyText = bodyText & "Review discovery source, target type, procedure or object identity, ownership role, status, required arguments, and diagnostics. Missing or partial evidence should remain explicit"
    bodyText = bodyText & " rather than being converted into unsupported conclusions."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Rescan only when the underlying environment changed or new evidence is required. Prior valid snapshots are preserved where the owning workflow provides snapshot protection."
    ChatGPTHelpAddTopic rows, "SCAN", "RESULTS", "Understanding Scan Results", bodyText, "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Candidate_Review")
    bodyText = vbNullString
    bodyText = bodyText & "Accept only candidates that represent intended, supported commands. Acceptance creates or updates configuration metadata; it does not replace target code, load missing projects, or"
    bodyText = bodyText & " prove that every dependency is valid."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "After acceptance, resolve Runner Macro targets and validate them before generation. Use the Database Reference when ownership, edit policy, or lifecycle is unclear."
    ChatGPTHelpAddTopic rows, "REVIEW", "ACCEPT", "Accepting Candidates", bodyText, "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Candidate_Review")
    bodyText = vbNullString
    bodyText = bodyText & "Ignore candidates that should not become commands. Defer candidates that may be valid after another dependency is available. Block rows that should not proceed. Use Manual Review"
    bodyText = bodyText & " when evidence is incomplete, ambiguous, inaccessible, or requires a deliberate human decision."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "These states preserve review history and prevent uncertain evidence from becoming automatic action."
    ChatGPTHelpAddTopic rows, "REVIEW", "DECISIONS", "Ignore, Defer, Block, or Manual Review", bodyText, "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Candidate_Review")
    bodyText = vbNullString
    bodyText = bodyText & "Generate Selected Runner Macros creates active CodeDb Runner Macro objects whose configuration rows have Selected for Runner Macro Generation enabled. It validates the target and preserves existing fail-closed"
    bodyText = bodyText & " rules before replacement."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Use the normal default workflow unless a reviewed configuration intentionally requires broader generation."
    ChatGPTHelpAddTopic rows, "GENERATE", "SELECTED", "Generate Selected Runner Macros", bodyText, "COMMAND:" & SYS_MACRO_KEY_GENERATE_SELECTED_RUNNER_MACROS
    bodyText = vbNullString
    bodyText = bodyText & "Generate All Configured Runner Macros creates every otherwise-eligible active configured Runner Macro regardless of Selected for Runner Macro Generation. Use it only when intentionally requesting the complete active"
    bodyText = bodyText & " configuration set."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Generation does not repair missing target code, infer unsupported arguments, or bypass validation."
    ChatGPTHelpAddTopic rows, "GENERATE", "ALL", "Generate All Configured Runner Macros", bodyText, "COMMAND:" & SYS_MACRO_KEY_GENERATE_ALL_CONFIGURED_RUNNER_MACROS
    bodyText = vbNullString
    bodyText = bodyText & "Health Review performs non-destructive checks across configuration, catalog state, managed objects and fields, setup consistency, generated output, Navigation Pane and ribbon"
    bodyText = bodyText & " placement, reference evidence, recovery blockers, and known setup risks."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Use the first original failing finding and its preceding stage to identify the correct owner. A final wrapper error may describe propagation rather than the originating defect."
    ChatGPTHelpAddTopic rows, "HEALTH", "REVIEW", "Health Review", bodyText, "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Health_Review")
    bodyText = vbNullString
    bodyText = bodyText & "Validation confirms that configured calls, object targets, arguments, ownership, catalog rows, field descriptions, generated forms, system macros, placement, and Module Guard"
    bodyText = bodyText & " evidence meet the current contract."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Fail-closed validation must not accept placeholder descriptions, arbitrary versions, incomplete command definitions, stale fingerprints, or merely existing properties when their"
    bodyText = bodyText & " authoritative values are wrong."
    ChatGPTHelpAddTopic rows, "HEALTH", "VALIDATE", "Validation Before Mutation", bodyText, "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Health_Review")
    bodyText = vbNullString
    bodyText = bodyText & "MAC is the authoritative catalog for " & DEFAULT_UTILITY_SHORT_NAME & "-managed tables, queries, forms, reports, and macros. It stores stable logical identity independently of the current physical"
    bodyText = bodyText & " Access object name, ownership, object description, purpose, edit policy, lifecycle, setup and repair policy, placement, documentation source, search terms, and verification"
    bodyText = bodyText & " expectations."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Physical object existence alone does not make catalog metadata authoritative."
    ChatGPTHelpAddTopic rows, "OBJECTS", "CATALOG", "MAC Catalog", bodyText, "OBJECT:TABLE:" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID)
    bodyText = vbNullString
    bodyText = bodyText & "MAC_Fields stores field identity, DAO type and size, required/default/index roles, FieldDescriptionLong for unrestricted Help/Database Reference/search text, FieldDescriptionShort for the"
    bodyText = bodyText & " Access DAO Field.Description property and other <=255-character uses, plus usage guidance, edit policy, lifecycle, lookup source, allowed-value source, display order, and verification status."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "MAC_Fields_Vals stores controlled values. Placeholder text such as UNMAPPED FIELD DESCRIPTION, TODO, Unknown, or Description unavailable remains invalid for " & DEFAULT_UTILITY_SHORT_NAME & "-owned"
    bodyText = bodyText & " authoritative descriptions."
    ChatGPTHelpAddTopic rows, "OBJECTS", "FIELDS", "Managed Field Documentation", bodyText, "OBJECT:TABLE:" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFields")
    bodyText = vbNullString
    bodyText = bodyText & "Field Order is the canonical physical order for fields in " & DEFAULT_UTILITY_SHORT_NAME & "-managed tables. Setup and repair derive it from the current owning schema definitions and apply the supported DAO field-order metadata so repaired or recreated tables remain deterministic."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Query Order is the explicit output-field order of " & DEFAULT_UTILITY_SHORT_NAME & "-managed saved queries. Managed query SQL names output fields intentionally instead of using wildcard SELECT output; dependent query families also retain deterministic build order so prerequisites are created before dependent queries."
    ChatGPTHelpAddTopic rows, "OBJECTS", "ORDER", "Field Order and Query Order", bodyText, "T:DBREF:INTRO"
    bodyText = vbNullString
    bodyText = bodyText & "MAC_Deps records what each managed object depends on and what depends on it. Dependency type, required status, purpose, failure policy, and stored keys support"
    bodyText = bodyText & " Database Reference navigation and validation."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Dependency evidence documents relationships; it does not authorize automatic modification of external projects or shared utilities."
    ChatGPTHelpAddTopic rows, "OBJECTS", "DEPENDENCIES", "Dependencies", bodyText, "OBJECT:TABLE:" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectDependencies")
    bodyText = vbNullString
    bodyText = bodyText & "Generated objects, " & DEFAULT_UTILITY_SHORT_NAME & " utility-owned objects, user-owned objects, external/shared dependencies, and external VBA projects or add-ins have different modification rules."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & DEFAULT_UTILITY_SHORT_NAME & "-prefixed hand-built forms may use catalog placement while remaining user-owned. User-owned forms are not rebuilt or deleted as generated forms. The generalized form-builder"
    bodyText = bodyText & " utility remains external/shared and is excluded from " & DEFAULT_UTILITY_SHORT_NAME & " managed-object naming updates and Module Guard unless explicitly approved. Reviewing an add-in never gives " & DEFAULT_UTILITY_SHORT_NAME & " ownership over it."
    ChatGPTHelpAddTopic rows, "OBJECTS", "OWNERSHIP", "Ownership Boundaries", bodyText, "T:FAQ:EDIT"
    bodyText = vbNullString
    bodyText = bodyText & "The canonical managed-object registry is the authority for the generated form set. UI_Frm_Registry stores catalog metadata for those canonical forms. Generated-form physical identity comes from the Defaults-owned canonical stem; ChatGPTUIForm=True appends the fixed"
    bodyText = bodyText & " _@ChatGptUI suffix and ChatGPTUIForm=False keeps the canonical stem unsuffixed. Logical identity remains stable and separate from physical naming."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "The generated owner registers each authoritative form, its display name, description, type, Navigation Pane group, sort order, source, and deterministic signature."
    ChatGPTHelpAddTopic rows, "FORMS", "GENERATED", "Generated Forms", bodyText, "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Manager")
    bodyText = vbNullString
    bodyText = bodyText & "Generated-form naming resolves the Defaults-owned canonical physical stem and, when ChatGPTUIForm=True, appends the fixed _@ChatGptUI suffix. It does not apply mutable generated-form prefix/postfix settings, readable-name rewriting, underscore replacement, or capitalization-boundary splitting."
    bodyText = bodyText & " Hardcoded historical compatibility names are not part of the current naming authority."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Physical form identity determines required captions and registration. Ribbon captions and form display names are presentation values, not substitutes for the managed logical"
    bodyText = bodyText & " identity."
    ChatGPTHelpAddTopic rows, "FORMS", "NAMING", "Generated-Form Naming", bodyText, "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Manager")
    bodyText = vbNullString
    bodyText = bodyText & "The builder preflights required sources, compares deterministic catalog, source, fixed-topic, and physical signatures, prepares durable rollback state, preserves current forms, and"
    bodyText = bodyText & " rebuilds only stale forms where possible."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "It registers every authoritative generated form, repairs references after approved renames, verifies Navigation Pane placement, excludes user-owned forms from generated"
    bodyText = bodyText & " replacement, and uses a no-op path when an unchanged second run proves the complete set is current."
    ChatGPTHelpAddTopic rows, "FORMS", "REBUILD", "Rebuilding Generated Forms", bodyText, "COMMAND:" & SYS_MACRO_KEY_BUILD_CHATGPT_UI_FORMS
    bodyText = vbNullString
    bodyText = bodyText & "Generated-form Ribbon button titles use their explicitly stored RibbonTitle values without appending Record Source text or applying title normalization. Record Source and source metadata remain available through the existing caption, ScreenTip, and SuperTip presentation paths."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Configurable naming must not break caption resolution because the required caption is determined from the physical form identity and its actual source."
    ChatGPTHelpAddTopic rows, "FORMS", "CAPTIONS", "Record Source and Sources Captions", bodyText, "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Manager")
    bodyText = vbNullString
    bodyText = bodyText & "Only the standard " & DEFAULT_UTILITY_SHORT_NAME & " Ribbon is generated. ApplyADCRibbonAfterBuild controls whether a completed Ribbon build automatically selects it as the database CustomRibbonID. The Ribbon is still fully generated and maintained when automatic application is disabled. Explicit Apply and Disable commands remain available."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Buttons use database-backed ScreenTip and SuperTip callbacks. A database reopen is required when the configured startup Ribbon changes, and when ADC Ribbon XML changes while ADC is already configured as the startup Ribbon."
    ChatGPTHelpAddTopic rows, "RIBBON", "RIBBON", "The Standard " & DEFAULT_UTILITY_SHORT_NAME & " Ribbon", bodyText, "COMMAND:" & SYS_MACRO_KEY_ADD_TABS_TO_ACCESS_RIBBON
    bodyText = vbNullString
    bodyText = bodyText & helpRibbonGroupLabel & " is on " & helpRibbonTabKey & ". Loaded-project system commands remain under ADC1 > Projects, while Loaded Table, Loaded Query, Links Table, and Links Query remain under " & DEFAULT_RIBBON_TAB_LABEL_ADMINISTRATION & " >"
    bodyText = bodyText & " Projects."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    If addInCommandKeys.count > 0 Then bodyText = bodyText & "Remaining Add-Ins scan/work commands stay under " & addInTabLabel & " > " & addInGroupLabel & " at group order " & CStr(addInGroupOrder) & ". The Add-Ins Review generated form is exposed directly as the current review surface."
    bodyText = bodyText & " Settings remains separate from System Tools. When present, X Combo Forms follows every nonterminal " & DEFAULT_RIBBON_TAB_LABEL_FORMS & " group; when present, X Sort is the final " & DEFAULT_RIBBON_TAB_LABEL_FORMS & " group, and " & DEFAULT_RIBBON_TAB_LABEL_DESIGN & " remains the final tab where"
    bodyText = bodyText & " required."
    ChatGPTHelpAddTopic rows, "RIBBON", "LAYOUT", "Current Ribbon Organization", bodyText, "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review")
    bodyText = vbNullString
    bodyText = bodyText & "Forms and reports identify Record Source or Sources. Queries identify source objects. Tables and macros retain the relevant object-name caption. Full descriptive text comes from"
    bodyText = bodyText & " managed metadata."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Ribbon callbacks resolve current data through UI_Ribbon_Generated_BtnTargets so presentation stays synchronized with current managed descriptions rather than fixed XML prose. Each object uses an explicitly assigned permanent Ribbon tag stored in Defaults; the tag is not recalculated from object name, type, caption, placement, or build order, so Ribbon rebuilds cannot remap an existing item."
    ChatGPTHelpAddTopic rows, "RIBBON", "TOOLTIPS", "ScreenTips and SuperTips", bodyText, "COMMAND:" & SYS_MACRO_KEY_ADD_TABS_TO_ACCESS_RIBBON
    bodyText = vbNullString
    bodyText = bodyText & "UI_Groups stores configured groups. MAC stores object-specific placement. Setup applies catalog-driven placement; the visible Navigation Pane is not the"
    bodyText = bodyText & " authority."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Add-In Review generated forms belong in the configured Add-ins group. Existing Projects placement remains separate. Strict validation enforces the final approved combinations and rejects"
    bodyText = bodyText & " mixed or unauthorized placement."
    ChatGPTHelpAddTopic rows, "RIBBON", "NAV", "Navigation Pane Placement", bodyText, "OBJECT:TABLE:" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups")
    bodyText = vbNullString
    bodyText = bodyText & "Built-in numbered physical macro names use _N Title: an underscore, the authoritative numeric macro identifier, one space, and no dash. Examples include _0 Run Full ADC Setup and _9025 Repair Part Report repair completion. Setup Part physical macro names instead use the canonical Part identity, such as _A1 Validate Configuration and Object Names, _B1 Build/Repair Managed Object Catalog Structure, and _J1 Verify and Persist Final Setup State."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Managed-object names that do not match the current approved naming configuration are reconciled through the approved object-name update workflow. Unrelated user overrides remain preserved."
    ChatGPTHelpAddTopic rows, "RIBBON", "SYSTEM_MACRO_NAMES", "System and Help Macro Names", bodyText, "COMMAND:" & SYS_MACRO_KEY_APPLY_OBJECT_NAMING_RULES
    bodyText = vbNullString
    bodyText = bodyText & "Module Guard prevents setup or repair from accepting the wrong VBA component, stale baselines, mixed coordinated batches, incomplete installation, overwritten modules, or"
    bodyText = bodyText & " unapproved local edits."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "It verifies actual VBProject identity, managed-header VB_Name, Version ID, Previous Version ID, localized timestamp, exact fingerprint, supported VBE-stable fingerprint, source"
    bodyText = bodyText & " line count, append-only history, verification snapshots, and the current guarded inventory."
    ChatGPTHelpAddTopic rows, "GUARD", "PURPOSE", "Why the Module Update Guard Exists", bodyText, "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Verification")
    bodyText = vbNullString
    bodyText = bodyText & "RegisterApprovedBuild is the explicit Module Guard registration boundary. It validates the complete embedded module inventory before writing manifest or history rows and records the exact"
    bodyText = bodyText & " installed source as authoritative."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Structural repair does not authorize a source mismatch. Arbitrary versions are never accepted, existing guarded modules are not removed, and Add-In Review changes are registered only"
    bodyText = bodyText & " after the final affected-module set is known."
    ChatGPTHelpAddTopic rows, "GUARD", "REGISTRATION", "Approved Build Registration", bodyText, "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Updates_Current")
    bodyText = vbNullString
    bodyText = bodyText & "Read the first failing module and reason before importing anything else. Confirm component name, header identity, current and previous Version IDs, timestamp, fingerprints, and"
    bodyText = bodyText & " line count against the coordinated delivery."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Do not weaken or bypass Module Guard to make setup pass. Replace only the approved affected modules, with Module Guard last."
    ChatGPTHelpAddTopic rows, "GUARD", "RECOVERY", "Resolving a Guard Failure", bodyText, "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Module_Updates_History")
    bodyText = vbNullString
    bodyText = bodyText & "User-facing pop-ups summarize what " & DEFAULT_UTILITY_SHORT_NAME & " was trying to do, what prevented completion, whether " & DEFAULT_UTILITY_SHORT_NAME & " stopped, skipped work, rolled back, or may have left changes in place, and the next"
    bodyText = bodyText & " required action."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Fingerprints, SQL, procedure chains, raw state dumps, and dense diagnostics remain in TraceLog and the physical trace log. One authoritative pop-up is shown for a handled failure;"
    bodyText = bodyText & " lower-level procedures preserve and propagate technical evidence."
    ChatGPTHelpAddTopic rows, "DIAG", "POPUPS", "Human-Readable Pop-Ups", bodyText, "T:DIAG:TRACE"
    bodyText = vbNullString
    bodyText = bodyText & "TraceLog and " & DEFAULT_UTILITY_LONG_NAME & " Trace.log retain detailed operation, stage, object, error, rollback, validation, fingerprint, and completion evidence. Immediate Window"
    bodyText = bodyText & " milestones and status-bar progress or heartbeat provide current execution context."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Troubleshooting starts with the first original error and the preceding stage. The final wrapper message may not identify the root cause."
    ChatGPTHelpAddTopic rows, "DIAG", "TRACE", "Trace, Physical Log, and Immediate Window", bodyText, "OBJECT:TABLE:TraceLog"
    bodyText = vbNullString
    bodyText = bodyText & "Open the managed modal Report Error form directly. A nonblank multiline description is required. Cancel closes without creating a report. Submit Report hides the dialog before capture, captures"
    bodyText = bodyText & " the visible Access window, uses PNG when available and BMP as fallback, appends description and metadata to retained TraceLog history and the existing physical log without clearing prior log content, and stores the screenshot through"
    bodyText = bodyText & " TraceLog_Screenshots. Setup owns creation and ownership verification of the form; direct Navigation Pane or Ribbon opening uses the same report workflow."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "The TraceLog ID is reported only after complete successful storage. Generated labels use dimensions and line breaks rather than an invalid WordWrap assignment."
    ChatGPTHelpAddTopic rows, "DIAG", "REPORT", "Report Error", bodyText, "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Report_Error_Dialog")
    bodyText = vbNullString
    bodyText = bodyText & "Recover Pending " & DEFAULT_UTILITY_SHORT_NAME & " Work is the fail-closed recovery command. It detects supported blockers actually present, handles Core, Catalog, Review, and UI recovery in dependency-safe"
    bodyText = bodyText & " order, retries only applicable boundaries, verifies that every blocker cleared, and preserves evidence."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "A retry returning is not sufficient proof of success. New mutating work remains blocked while durable recovery is unfinished."
    ChatGPTHelpAddTopic rows, "DIAG", "RECOVERY", "Recover Pending " & DEFAULT_UTILITY_SHORT_NAME & " Work", bodyText, "COMMAND:" & SYS_MACRO_KEY_RESOLVE_PENDING_SETUP_RECOVERY
    bodyText = vbNullString
    bodyText = bodyText & "1. Preserve the popup, TraceLog row, physical log, and Report Error screenshot when available."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "2. Identify the first failing stage, module, object, field, or recovery boundary."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "3. Check Health Review, Module Guard verification, Help, and the Database Reference."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "4. Correct the owning source, metadata, name, reference, or physical object."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "5. Use the narrow repair or recovery command when one exists."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "6. Run full Setup only when complete reconciliation is required."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Do not repeatedly rerun replacement work while a durable rollback or recovery blocker remains unresolved."
    ChatGPTHelpAddTopic rows, "DIAG", "FAILURE", "Troubleshooting Order", bodyText, "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Health_Review")
    bodyText = vbNullString
    bodyText = bodyText & "Active References inventories the current database Application.References collection, including broken state, reference kind, priority, GUID/version, path and file metadata,"
    bodyText = bodyText & " registration provenance, health status, and loaded-project match evidence. Available References inventories registered type libraries that may be activated."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Scans are read-only and do not change membership or order."
    ChatGPTHelpAddTopic rows, "REFS", "ACTIVE", "Active and Available References", bodyText, "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Active_References")
    bodyText = vbNullString
    bodyText = bodyText & "A Project reference and a loaded VBA project are related but not identical. DBInv_VBARefs_Active represents references selected by the current database; DBInv_VBAPrjs_Ld represents"
    bodyText = bodyText & " projects currently exposed by Application.VBE.VBProjects. DBInv_VBARefs_Active_ReferencedPrjLinks uses exact normalized path where possible and unique name only as weaker evidence."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Loaded-project system commands remain under ADC1 > Projects; table and query evidence remains under " & DEFAULT_RIBBON_TAB_LABEL_ADMINISTRATION & " > Projects."
    ChatGPTHelpAddTopic rows, "REFS", "PROJECTS", "Project References and Loaded Projects", bodyText, "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review_Projects")
    bodyText = vbNullString
    bodyText = bodyText & "Loaded VBA Projects records VBE order, project name, CodeDb/HostDb/external role, active-project state, component count, accessibility, source readability, protection, project"
    bodyText = bodyText & " type, path, scan status, and diagnostics."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "The inventory does not load an unloaded project, read inaccessible source, modify external projects, or change active references."
    ChatGPTHelpAddTopic rows, "REFS", "LOADED", "Loaded VBA Projects", bodyText, "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review_Projects")
    bodyText = vbNullString
    bodyText = bodyText & "Reference Manager stages RequestedActive changes beside original active state, explains protected or ineligible rows, and applies only approved changes through guarded procedures."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Setup, Repair, and passive scans do not change active references. Rescan after external library installation, registration, file movement, or a deliberate"
    bodyText = bodyText & " Apply operation."
    ChatGPTHelpAddTopic rows, "REFS", "MANAGER", "Reference Manager", bodyText, "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Reference_Manager")
    bodyText = vbNullString
    bodyText = bodyText & "Reference Inventory passively inventories loaded VBA projects, registered VBE add-ins, registered COM add-ins, and the combined inventory. Loaded VBA Analysis performs read-only analysis of all inspectable loaded VBA projects:"
    bodyText = bodyText & " references, modules, procedures, and dependency evidence. External-only procedure and conflict projections are derived from that canonical inventory. Add-In Review presents those existing results in one generated"
    bodyText = bodyText & " consolidated review interface."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Add-In Review does not replace the Reference Inventory or Loaded VBA Analysis data model. Individual query commands remain available; the consolidated form is the normal review entry point."
    ChatGPTHelpAddTopic rows, "ADDINS", "OVERVIEW", "Add-Ins Inventory and Review Workflow", bodyText, "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review")
    bodyText = vbNullString
    bodyText = bodyText & "Reference Inventory distinguishes registered or connected VBE add-ins, registered COM add-ins, loaded VBA projects, active project references, combined inventory records, registered but"
    bodyText = bodyText & " unavailable add-ins, unmatched add-ins or projects, and inaccessible loaded projects."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "It does not connect or disconnect VBE add-ins, enable or disable COM add-ins, register or unregister add-ins, load unloaded add-ins, modify references, or modify external projects."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Authoritative Reference Inventory storage is "
    bodyText = bodyText & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects")
    bodyText = bodyText & ", "
    bodyText = bodyText & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_VBEAddIns")
    bodyText = bodyText & ", and "
    bodyText = bodyText & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_COMAddIns")
    bodyText = bodyText & ". The combined read-only query is "
    bodyText = bodyText & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Inventory")
    bodyText = bodyText & "."
    ChatGPTHelpAddTopic rows, "ADDINS", "INVENTORY", "Reference Inventory", bodyText, "COMMAND:" & SYS_MACRO_KEY_SCAN_DB_INVENTORY_ADD_INS_AND_LOADED_VBA_PROJECTS
    bodyText = vbNullString
    bodyText = bodyText & "Loaded VBA Analysis analyzes all loaded projects that can be inspected. The canonical tables inventory project references, modules, procedure declarations, and supported dependency evidence across Current CodeDb, HostDb, and External roles."
    bodyText = bodyText & " External-only review queries then derive public-procedure, conflict, and duplicate evidence for externally callable names from that canonical inventory."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "It does not load unloaded add-ins, open unloaded external source, modify source, add or remove references, rename modules or procedures, resolve conflicts automatically, or infer"
    bodyText = bodyText & " unsupported dependencies. Locked or inaccessible projects and partial evidence remain explicit, and prior completed snapshots are preserved by the existing workflow when refresh"
    bodyText = bodyText & " fails."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Authoritative Loaded VBA Analysis tables are "
    bodyText = bodyText & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjectReferences")
    bodyText = bodyText & ", "
    bodyText = bodyText & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAModules")
    bodyText = bodyText & ", "
    bodyText = bodyText & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProcedures")
    bodyText = bodyText & ", and "
    bodyText = bodyText & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBADependencies")
    bodyText = bodyText & ". Review queries are "
    bodyText = bodyText & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_ExternalProject_References")
    bodyText = bodyText & ", "
    bodyText = bodyText & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_ExternalProject_Modules")
    bodyText = bodyText & ", "
    bodyText = bodyText & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_ExternalProject_Public_Procedures")
    bodyText = bodyText & ", "
    bodyText = bodyText & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_ExternalProject_Procedure_Conflicts")
    bodyText = bodyText & ", and "
    bodyText = bodyText & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_ExternalProject_Dependencies")
    bodyText = bodyText & "."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Duplicate procedure conflicts are derived when the same externally callable ProcedureNameKey occurs more than once across the External-only callable-procedure projection derived from the canonical loaded-project inventory. Conflict review is evidence only; it does not rename or alter procedures, modules, or projects."
    ChatGPTHelpAddTopic rows, "ADDINS", "ANALYSIS", "Loaded VBA Analysis", bodyText, "COMMAND:" & SYS_MACRO_KEY_SCAN_DB_INVENTORY_LOADED_VBA_SOURCE
    bodyText = vbNullString
    bodyText = bodyText & "The primary generated form is "
    bodyText = bodyText & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review")
    bodyText = bodyText & ". It presents supported evidence for detected add-ins and loaded projects, VBE versus COM identity, loaded/registered/available/inaccessible/unmatched status, project and add-in"
    bodyText = bodyText & " names, path, ProgID, GUID, description, connection and registration evidence, and Loaded VBA Analysis reference, module, procedure, conflict, and dependency counts."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "The interface is read-only and reports only evidence actually stored by Reference Inventory and Loaded VBA Analysis. Reference, module, public-procedure, dependency, and Loaded VBA Analysis status/count evidence can represent all loaded-project roles; ProcedureConflictCount remains an External-project conflict-analysis metric."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Its consolidated authoritative query is "
    bodyText = bodyText & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Review_Summary")
    bodyText = bodyText & "."
    ChatGPTHelpAddTopic rows, "ADDINS", "CONSOLIDATED", "Add-In Review", bodyText, "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review")
    bodyText = vbNullString
    bodyText = bodyText & "The consolidated interface includes read-only pages for Combined Inventory, VBE Add-Ins, COM Add-Ins, and Loaded Projects. Their generated forms are "
    bodyText = bodyText & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review_Inventory")
    bodyText = bodyText & ", "
    bodyText = bodyText & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review_VBE")
    bodyText = bodyText & ", "
    bodyText = bodyText & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review_COM")
    bodyText = bodyText & ", and "
    bodyText = bodyText & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review_Projects")
    bodyText = bodyText & "."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Selecting an item establishes Add-In Review context only when a reliable stored relationship exists. Selection never loads, connects, enables, disables, or registers anything; missing"
    bodyText = bodyText & " relationships remain visibly unresolved."
    ChatGPTHelpAddTopic rows, "ADDINS", "INVENTORYPAGES", "Inventory Review Pages", bodyText, "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review_Inventory")
    bodyText = vbNullString
    bodyText = bodyText & "Loaded VBA Analysis scans and retains source evidence for all inspectable loaded projects across Current CodeDb, HostDb, and External roles. The Add-In Review analysis detail pages are deliberately External-only projections over that canonical inventory; their RecordSources remain the existing ExternalProject_* queries rather than the raw LoadedVBA* tables. The detail pages show Project References, Modules, Procedures, Conflicts, and Dependencies through "
    bodyText = bodyText & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review_Project_References")
    bodyText = bodyText & ", "
    bodyText = bodyText & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review_Modules")
    bodyText = bodyText & ", "
    bodyText = bodyText & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review_Procedures")
    bodyText = bodyText & ", "
    bodyText = bodyText & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review_Conflicts")
    bodyText = bodyText & ", and "
    bodyText = bodyText & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review_Dependencies")
    bodyText = bodyText & "."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "External-project detail pages filter through the current stable ProjectIdentity. CodeDb and HostDb projects can still show source-analysis status/counts in Loaded Projects and consolidated summary views, but selecting them does not convert the External-only tabs into raw all-project detail views. Module and procedure actions use stored ModuleIdentity and ProcedureNameKey where applicable. Changing selection"
    bodyText = bodyText & " refreshes dependent views without reopening the manager. Unloaded or inaccessible projects show an accurate unavailable state; " & DEFAULT_UTILITY_SHORT_NAME & " does not load them to populate details."
    ChatGPTHelpAddTopic rows, "ADDINS", "ANALYSISPAGES", "Analysis Detail Pages", bodyText, "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review_Project_References")
    bodyText = vbNullString
    bodyText = bodyText & "Implemented display filters include the standard Filter By Selection and Clear Filter actions on every Add-In Review detail page. The Modules page also provides Procedures and"
    bodyText = bodyText & " Dependencies navigation; Procedures provides Conflicts and Dependencies; Conflicts provides All Exposures."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Filters affect display only. They do not modify scan results, add-in state, references, modules, procedures, conflicts, or dependencies. Clear Filters restores the complete"
    bodyText = bodyText & " applicable view."
    ChatGPTHelpAddTopic rows, "ADDINS", "FILTERS", "Filtering and Search", bodyText, "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review")
    bodyText = vbNullString
    bodyText = bodyText & "Supported navigation follows stable stored keys among add-in and loaded project, project and references, project and modules, module and procedures, procedure and conflicts, module"
    bodyText = bodyText & " or procedure and dependencies, and conflict exposures."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Stable ProjectIdentity, ModuleIdentity, and ProcedureNameKey are preferred. Display text is not a substitute for a reliable key. Missing relationships produce an explicit"
    bodyText = bodyText & " limitation, and navigation changes review context only."
    ChatGPTHelpAddTopic rows, "ADDINS", "NAVIGATION", "Cross-Navigation", bodyText, "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review")
    bodyText = vbNullString
    bodyText = bodyText & "Refresh Display requeries stored results and does not scan. Inventory Scan calls the existing passive add-in inventory procedure and does not change add-in state. Loaded VBA Analysis Scan calls"
    bodyText = bodyText & " the existing loaded-project analysis and remains limited to loaded and accessible projects. Raw Data opens the underlying read-only source for the current page."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "The form does not duplicate scan logic. Failures use the owning diagnostic path, and valid prior snapshots are preserved according to the existing Reference Inventory or Loaded VBA Analysis workflow."
    ChatGPTHelpAddTopic rows, "ADDINS", "REFRESH", "Display Refresh and Scan Actions", bodyText, "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review")
    bodyText = vbNullString
    bodyText = bodyText & "The Add-Ins workflow does not connect or disconnect VBE add-ins, enable or disable COM add-ins, change registration, load an unloaded add-in, open unloaded external source, modify"
    bodyText = bodyText & " external VBA projects, add or remove references, rename modules or procedures, automatically resolve conflicts, delete add-ins or analysis rows as repair, infer unsupported"
    bodyText = bodyText & " dependencies, or replace a valid completed snapshot with partial failed data."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Diagnostic evidence supports review; it does not authorize external modification."
    ChatGPTHelpAddTopic rows, "ADDINS", "SAFETY", "Add-Ins Safety Boundaries", bodyText, "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review")
    bodyText = vbNullString
    If addInCommandKeys.count > 0 Then bodyText = bodyText & "The remaining Add-Ins system commands are work commands rather than object-open shortcuts: " & ChatGPTHelpSystemMacroTitleList(addInCommandKeys, systemMacroDefinitionIndex) & ". They remain under " & addInTabLabel & " > " & addInGroupLabel & " at group order " & CStr(addInGroupOrder) & "."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "The consolidated Add-Ins Review generated form is opened directly and " & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Review_Summary") & " supplies its summary data. Explicit scan commands remain available when fresh inventory evidence is needed, and duplicate ribbon targets are rejected."
    ChatGPTHelpAddTopic rows, "ADDINS", "COMMANDS", "Commands and Placement", bodyText, "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review")
    bodyText = vbNullString
    bodyText = bodyText & "Add-In Review forms are centrally generated and registered. Setup verifies Reference Inventory and Loaded VBA Analysis sources, creates missing Add-In Review forms, repairs properties, controls, event wiring, catalog"
    bodyText = bodyText & " rows, and placement, preserves current forms, rebuilds only stale forms, rolls back incomplete generation, places the forms in the configured Add-ins Navigation Pane group, and"
    bodyText = bodyText & " preserves parent setup progress, Pause, Cancel, pending-request, and heartbeat behavior."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "No manual one-off form creation competes with the generated-form owner."
    ChatGPTHelpAddTopic rows, "ADDINS", "SETUP", "Generated Ownership and Setup", bodyText, "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review")
    bodyText = vbNullString
    bodyText = bodyText & "Configuration, settings, review decisions, documentation guidance, requested reference state, and explicitly user-owned metadata may be edited through supported forms."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Utility-owned structural metadata, generated physical objects, Module Guard history, and setup policy fields should change only through their owning"
    bodyText = bodyText & " workflow. The Database Reference shows edit policy and lifecycle guidance."
    ChatGPTHelpAddTopic rows, "FAQ", "EDIT", "What Can I Edit?", bodyText, "T:DBREF:INTRO"
    bodyText = vbNullString
    bodyText = bodyText & "Setup synchronizes object and field documentation from authoritative managed definitions and verified physical metadata. A blank or missing authoritative description remains"
    bodyText = bodyText & " visible as missing; Help does not create a blank DAO property or invent substitute prose."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "For " & DEFAULT_UTILITY_SHORT_NAME & "-owned fields, placeholders such as UNMAPPED FIELD DESCRIPTION remain invalid and stop strict validation until the actual mapping is supplied."
    ChatGPTHelpAddTopic rows, "FAQ", "MISSING", "Why Is Documentation Marked Missing?", bodyText, "T:DBREF:ISSUES"
    bodyText = vbNullString
    bodyText = bodyText & "Controlled terms are stored in the canonical " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_TERMINOLOGY_TABLE_LOGICAL_ID) & " table. Open " & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Terminology_Review") & " directly for the managed editable terminology reference view; edits through the query update the authoritative Help_Terms rows directly, and no terminology system macro is required."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "ADC_Terminology is not a valid current ADC object name and must not be created or referenced."
    ChatGPTHelpAddTopic rows, "FAQ", "TERMS", "Where Are Controlled Terms Defined?", bodyText, "OBJECT:QUERY:" & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Terminology_Review")
    bodyText = vbNullString
    bodyText = bodyText & "The Database Reference lists active " & DEFAULT_UTILITY_SHORT_NAME & "-managed objects. Select an object to see purpose, ownership, edit policy, lifecycle, repair owner, workflow, Navigation Pane and ribbon"
    bodyText = bodyText & " placement, documentation source, schema verification, dependencies, and fields. Select a field to see structural metadata, description, usage guidance, lookup details, allowed"
    bodyText = bodyText & " values, and verification state."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "It reads MAC, MAC_Fields, MAC_Fields_Vals, and MAC_Deps. Live DAO metadata is harvested and synchronized during setup, not while Help"
    bodyText = bodyText & " is being browsed."
    ChatGPTHelpAddTopic rows, "DBREF", "INTRO", "Using the Database Reference", bodyText, "OBJECT:TABLE:" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID)
    bodyText = vbNullString
    bodyText = bodyText & "Search for MissingDescription, StructureUnavailable, an object or field name, Report Error, Module Guard, Add-in, external project, Fingerprint, Recovery, Record Source, Sources,"
    bodyText = bodyText & " or Ownership to locate evidence."
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & vbCrLf
    bodyText = bodyText & "Inactive field-documentation rows remain stored with inactivity metadata so schema history is not silently discarded. Normal object and field lists show active rows only. If a required documentation"
    bodyText = bodyText & " table is missing, Help directs the user to Setup rather than fabricating data."
    ChatGPTHelpAddTopic rows, "DBREF", "ISSUES", "Missing, Inactive, and Unavailable Documentation", bodyText, "OBJECT:TABLE:" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFields")
    Set ChatGPTHelpTopics = rows
End Function

' Purpose: Validates chat gpt help definitions for help topics, supplied system macro definitions.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   helpTopics (Input Optional; Variant): Help-topic payload supplied by the caller, normally the array/list representation consumed by the Help UI; omitted/Empty allows the routine to resolve its normal topic
'   source.
'   suppliedSystemMacroDefinitions (Input Optional; Variant): Caller-supplied supplied system macro definitions data used as the authoritative definition input for this call, avoiding an unnecessary reload when
'   already available.
'   suppliedSystemMacroDefinitionIndex (Input Optional; Variant): Caller-supplied supplied system macro definition index data used as the authoritative definition input for this call, avoiding an unnecessary
'   reload when already available.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub ValidateChatGPTHelpDefinitions(Optional ByVal helpTopics As Variant, Optional ByVal suppliedSystemMacroDefinitions As Variant, Optional ByVal suppliedSystemMacroDefinitionIndex As Variant)
    On Error GoTo ErrHandler
    mGeneratedFormHelpDefinitionValidationCount = mGeneratedFormHelpDefinitionValidationCount + 1
    Dim topics As Collection
    Dim topicsByKey As Object
    Dim topic As Variant
    Dim topicKey As String
    Dim sectionKey As String
    Dim fixedTopicKey As String
    Dim TitleText As String
    Dim bodyText As String
    Dim relatedTarget As String
    Dim targetParts As Variant
    Dim relatedTopicKey As String
    Dim commandKey As String
    Dim objectTypeText As String
    Dim requiredAddInKeys As Variant
    Dim requiredSearchTerms As Variant
    Dim requiredItem As Variant
    Dim aggregateText As String
    Dim addInTopicCount As Long
    Dim setupPartTopic As Variant
    Dim commandTopic As Variant
    Dim consolidatedTopic As Variant
    Dim systemMacroDefinitions As Collection
    Dim systemMacroDefinitionIndex As Object
    If IsObject(helpTopics) Then Set topics = helpTopics
    If isMissing(suppliedSystemMacroDefinitions) = False Then
        If IsObject(suppliedSystemMacroDefinitions) Then Set systemMacroDefinitions = suppliedSystemMacroDefinitions
    End If
    If isMissing(suppliedSystemMacroDefinitionIndex) = False Then
        If IsObject(suppliedSystemMacroDefinitionIndex) Then Set systemMacroDefinitionIndex = suppliedSystemMacroDefinitionIndex
    End If
    If topics Is Nothing Then
        If systemMacroDefinitions Is Nothing Then Set systemMacroDefinitions = x__UTIL_ADC_DEFAULTS.GetSystemMacroDefinitions()
        If systemMacroDefinitionIndex Is Nothing Then Set systemMacroDefinitionIndex = ChatGPTSystemMacroDefinitionIndex(systemMacroDefinitions)
        Set topics = ChatGPTHelpTopics(systemMacroDefinitions, systemMacroDefinitionIndex)
    ElseIf systemMacroDefinitionIndex Is Nothing Then
        If systemMacroDefinitions Is Nothing Then Set systemMacroDefinitions = x__UTIL_ADC_DEFAULTS.GetSystemMacroDefinitions()
        Set systemMacroDefinitionIndex = ChatGPTSystemMacroDefinitionIndex(systemMacroDefinitions)
    End If
    ' INTENTIONAL KEYED LOOKUP: topicsByKey provides keyed membership/lookup over source rows or state already read by ValidateChatGPTHelpDefinitions; it is a derived access path, not a second authority.
    Set topicsByKey = CreateObject("Scripting.Dictionary")
    topicsByKey.CompareMode = vbTextCompare
    For Each topic In topics
        SetupControlPump
        If IsArray(topic) = False Then Err.Raise vbObjectError + 7330, "ValidateChatGPTHelpDefinitions", "A fixed Help topic definition is not an array."
        If LBound(topic) <> 0 Or UBound(topic) < 4 Then Err.Raise vbObjectError + 7331, "ValidateChatGPTHelpDefinitions", "A fixed Help topic definition does not contain section, key, title, body, and related-target values."
        sectionKey = UCase$(Trim$(CStr(topic(0))))
        fixedTopicKey = UCase$(Trim$(CStr(topic(1))))
        TitleText = Trim$(CStr(topic(2)))
        bodyText = Trim$(CStr(topic(3)))
        relatedTarget = Trim$(CStr(topic(4)))
        If ChatGPTHelpKnownSectionKey(sectionKey) = False Then Err.Raise vbObjectError + 7332, "ValidateChatGPTHelpDefinitions", "A fixed Help topic uses an unknown section." & vbCrLf & "Section: " & sectionKey & vbCrLf & "Topic: " & fixedTopicKey
        If Len(fixedTopicKey) = 0 Or Len(TitleText) = 0 Or Len(bodyText) = 0 Then Err.Raise vbObjectError + 7333, "ValidateChatGPTHelpDefinitions", "A fixed Help topic has a blank key, title, or body." & vbCrLf & "Section: " & sectionKey & vbCrLf & "Topic: " & fixedTopicKey
        topicKey = sectionKey & "|" & fixedTopicKey
        If topicsByKey.Exists(topicKey) Then Err.Raise vbObjectError + 7334, "ValidateChatGPTHelpDefinitions", "A fixed Help topic key is duplicated." & vbCrLf & "Topic: " & topicKey
        topicsByKey.Add topicKey, topic
        If sectionKey = "ADDINS" Then addInTopicCount = addInTopicCount + 1
        aggregateText = aggregateText & vbCrLf & TitleText & vbCrLf & bodyText
    Next topic
    For Each topic In topics
        SetupControlPump
        relatedTarget = Trim$(CStr(topic(4)))
        If Len(relatedTarget) > 0 Then
            targetParts = Split(relatedTarget, ":")
            Select Case UCase$(Trim$(CStr(targetParts(0))))
                Case "T"
                    If UBound(targetParts) < 2 Then Err.Raise vbObjectError + 7335, "ValidateChatGPTHelpDefinitions", "A fixed-topic related target is incomplete." & vbCrLf & "Target: " & relatedTarget
                    relatedTopicKey = UCase$(Trim$(CStr(targetParts(1)))) & "|" & UCase$(Trim$(CStr(targetParts(2))))
                    If topicsByKey.Exists(relatedTopicKey) = False Then Err.Raise vbObjectError + 7336, "ValidateChatGPTHelpDefinitions", "A fixed-topic related target does not exist." & vbCrLf & "Target: " & relatedTarget
                Case "FORM"
                    If UBound(targetParts) < 1 Or Len(Trim$(Mid$(relatedTarget, InStr(1, relatedTarget, ":", vbBinaryCompare) + 1))) = 0 Then Err.Raise vbObjectError + 7337, "ValidateChatGPTHelpDefinitions", "A form related target is incomplete." & vbCrLf & "Target: " & relatedTarget
                Case "COMMAND"
                    commandKey = Trim$(Mid$(relatedTarget, InStr(1, relatedTarget, ":", vbBinaryCompare) + 1))
                    If Len(commandKey) = 0 Then Err.Raise vbObjectError + 7337, "ValidateChatGPTHelpDefinitions", "A command related target is incomplete." & vbCrLf & "Target: " & relatedTarget
                    If ChatGPTHelpCommandTargetKnown(commandKey, systemMacroDefinitionIndex) = False Then Err.Raise vbObjectError + 7365, "ValidateChatGPTHelpDefinitions", "A command related target is not part of the authoritative command set." & vbCrLf & "Target: " & relatedTarget
                Case "OBJECT"
                    If UBound(targetParts) < 2 Or Len(Trim$(CStr(targetParts(1)))) = 0 Or Len(Trim$(CStr(targetParts(2)))) = 0 Then Err.Raise vbObjectError + 7338, "ValidateChatGPTHelpDefinitions", "An object related target is incomplete." & vbCrLf & "Target: " & relatedTarget
                    objectTypeText = NormalizedObjectTypeCode(CStr(targetParts(1)))
                    If Len(objectTypeText) = 0 Then Err.Raise vbObjectError + 7366, "ValidateChatGPTHelpDefinitions", "An object related target uses an unsupported Access object type." & vbCrLf & "Target: " & relatedTarget
                Case Else
                    Err.Raise vbObjectError + 7339, "ValidateChatGPTHelpDefinitions", "A fixed Help topic uses an unsupported related-target type." & vbCrLf & "Target: " & relatedTarget
            End Select
        End If
    Next topic
    requiredAddInKeys = Array("OVERVIEW", "INVENTORY", "ANALYSIS", "CONSOLIDATED", "INVENTORYPAGES", "ANALYSISPAGES", "FILTERS", "NAVIGATION", "REFRESH", "SAFETY", "COMMANDS", "SETUP")
    For Each requiredItem In requiredAddInKeys
        If topicsByKey.Exists("ADDINS|" & CStr(requiredItem)) = False Then Err.Raise vbObjectError + 7340, "ValidateChatGPTHelpDefinitions", "A required Add-Ins Help topic is missing." & vbCrLf & "Topic: ADDINS|" & CStr(requiredItem)
    Next requiredItem
    If addInTopicCount <> UBound(requiredAddInKeys) - LBound(requiredAddInKeys) + 1 Then Err.Raise vbObjectError + 7341, "ValidateChatGPTHelpDefinitions", "The Add-Ins Help section contains an unexpected topic count." & vbCrLf & "Expected: " & CStr(UBound(requiredAddInKeys) - LBound(requiredAddInKeys) + 1) & vbCrLf & "Actual: " & CStr(addInTopicCount)
    setupPartTopic = topicsByKey("SETUP|PARTS")
    If InStr(1, CStr(setupPartTopic(3)), CStr(ImplementedSetupPartCount()) & " substantive parts", vbTextCompare) = 0 Then Err.Raise vbObjectError + 7342, "ValidateChatGPTHelpDefinitions", "The Setup Parts topic does not use the current authoritative setup-part total." & vbCrLf & "Expected total: " & CStr(ImplementedSetupPartCount())
    commandTopic = topicsByKey("ADDINS|COMMANDS")
    If StrComp(Trim$(CStr(commandTopic(4))), "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review"), vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7343, "ValidateChatGPTHelpDefinitions", "The Add-Ins Commands topic does not target the authoritative consolidated Add-Ins Review form."
    If InStr(1, CStr(commandTopic(3)), ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Review_Summary"), vbBinaryCompare) = 0 Or InStr(1, CStr(commandTopic(3)), "opened directly", vbTextCompare) = 0 Then Err.Raise vbObjectError + 7344, "ValidateChatGPTHelpDefinitions", "The Add-Ins Commands topic does not document the authoritative Add-In Review query and direct-form opening contract."
    consolidatedTopic = topicsByKey("ADDINS|CONSOLIDATED")
    If StrComp(Trim$(CStr(consolidatedTopic(4))), "FORM:" & ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Add_In_Review"), vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7345, "ValidateChatGPTHelpDefinitions", "The consolidated Add-Ins topic does not target the current generated Add-In Review form."
    requiredSearchTerms = Array("Add-in", "Add-ins", "VBE add-in", "COM add-in", "Loaded project", "External project", "Combined inventory", "Project reference", "Module", "Public procedure", "Procedure conflict", "Duplicate procedure", "Dependency", "Reference Inventory", "Loaded VBA Analysis", "Add-In Review", "Refresh", "Rescan", "Registered", "Connected", "Loaded", "Unavailable", "Inaccessible", "Unmatched", "Filter", "Clear filters", "Cross-navigation", "Pause", "Cancel", "Pending", "Frozen", "Heartbeat", "Part X of X", "Report Error", "Screenshot", "Module Guard", "Fingerprint", "Recovery", "Record Source", "Sources", "Field Order", "Query Order", "Ownership")
    For Each requiredItem In requiredSearchTerms
        If InStr(1, aggregateText, CStr(requiredItem), vbTextCompare) = 0 Then Err.Raise vbObjectError + 7346, "ValidateChatGPTHelpDefinitions", "A required fixed-topic Help search term is missing." & vbCrLf & "Search term: " & CStr(requiredItem)
    Next requiredItem
    LogTraceMessage TRACE_LEVEL_INFO, "UI", "ValidateChatGPTHelpDefinitions", "Generated Help fixed-topic definitions passed strict validation.", "Topics: " & CStr(topics.count) & vbCrLf & "Add-Ins topics: " & CStr(addInTopicCount) & vbCrLf & "Setup parts documented: " & CStr(ImplementedSetupPartCount()) & vbCrLf & "Fixed-topic signature: " & ChatGPTHelpFixedTopicSignature(topics), "Validate generated Help definitions", 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Help"), "Form", "CodeDb", "Help"
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ValidateChatGPTHelpDefinitions", "Generated Help fixed-topic validation failed." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

' Purpose: Clears the retained D1 generated-Help definition proof owned by this module.
' Public role: Inter-module public API. UI_FORMS calls this when D1 proof state is reset or invalidated.
' Parameters: None.
' Returns: None (Sub). The retained Help signature is cleared; validation/reuse counters remain lifetime instrumentation.
Public Sub ClearSetupGeneratedFormHelpProofState()
    mSetupGeneratedFormHelpSignature = vbNullString
End Sub

' Purpose: Publishes the validated D1 generated-Help definition signature owned by this module.
' Public role: Inter-module public API. UI_FORMS calls this only after strict D1 Help-definition validation succeeds.
' Parameters:
'   helpSignature (Input; String): Nonblank fixed-topic signature proving the exact generated-Help definitions validated during D1.
' Returns: None (Sub). Invalid blank proof input raises instead of publishing an ambiguous state.
Public Sub PublishSetupGeneratedFormHelpProof(ByVal helpSignature As String)
    helpSignature = Trim$(helpSignature)
    If Len(helpSignature) = 0 Then Err.Raise vbObjectError + 7347, "PublishSetupGeneratedFormHelpProof", "A blank generated Help-definition signature cannot be published."
    mSetupGeneratedFormHelpSignature = helpSignature
End Sub

' Purpose: Verifies that the retained D1 generated-Help signature is still reusable for the requested generated-form build scope.
' Public role: Read-only inter-module proof API used by generated-form build code; UI_FORMS remains the owner of D1 preflight state and exposes that state only through its readiness API.
' Parameters:
'   normalizedBuildScope (Input; String): Normalized generated-form scope that must match the retained D1 preflight plan.
'   detailText (Output/in-out; String): Receives the D1-plan/proof failure reason or the successful Help-proof reuse detail and instrumentation counts.
' Returns: Boolean: True only when the same-run D1 plan is current for the requested scope and the current fixed-topic Help signature still matches the D1 signature.
Public Function SetupGeneratedFormHelpProofCurrent(ByVal normalizedBuildScope As String, ByRef detailText As String) As Boolean
    Dim currentSignature As String
    If x__UTIL_ADC_UI_FORMS.SetupGeneratedFormPreflightReadyForSetup(detailText, normalizedBuildScope) = False Then Exit Function
    If Len(mSetupGeneratedFormHelpSignature) = 0 Then detailText = "The D1 Help-definition proof is unavailable.": Exit Function
    currentSignature = ChatGPTHelpFixedTopicSignature()
    If StrComp(currentSignature, mSetupGeneratedFormHelpSignature, vbBinaryCompare) <> 0 Then detailText = "The generated Help fixed-topic signature changed after D1.": Exit Function
    mGeneratedFormHelpProofReuseCount = mGeneratedFormHelpProofReuseCount + 1
    detailText = "Reused D1 generated Help-definition validation. Signature=" & currentSignature & "; full validations=" & CStr(mGeneratedFormHelpDefinitionValidationCount) & "; D1-to-D2 reuses=" & CStr(mGeneratedFormHelpProofReuseCount)
    SetupGeneratedFormHelpProofCurrent = True
End Function

' =============================================================================================================================================================================
' HELP FORM CONSTRUCTION / SPECIALIZED VALIDATION PROCEDURES
' =============================================================================================================================================================================

Private Function ChatGPTHelpCreateListBox(ByVal formName As String, ByVal controlName As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal widthValue As Long, ByVal heightValue As Long, ByVal onClickExpression As String) As control
    Dim ctl As control
    Set ctl = CreateControl(formName, acListBox, acDetail, vbNullString, vbNullString, leftPos, topPos, widthValue, heightValue)
    ctl.Name = controlName
    ctl.rowSourceType = "Value List"
    ctl.columnCount = 2
    ctl.BoundColumn = 1
    ctl.columnWidths = "0;2.75"
    ctl.OnClick = onClickExpression
    ctl.fontSize = DEFAULT_UI_READING_FONT_SIZE
    Set ChatGPTHelpCreateListBox = ctl
End Function

' Purpose: Builds the managed Help form from its canonical generated-form contract, wires its Access event/control expressions, saves the completed design, and reports the form through builtForms or skippedForms.
' Public role: Inter-module public generated-form builder. The UI form-build dispatcher calls this routine to create the managed Help form and report built/skipped state.
' Parameters:
'   replaceExisting (Input; Boolean): When True, enables or requests existing for this call; when False, the routine follows its normal non-forced/default path.
'   builtForms (Input/output; Collection): Collection containing the built forms items that the routine consumes or updates as part of this operation.
'   skippedForms (Input/output; Collection): Supplies the caller-owned skipped-form collection and can receive a skip-reason entry when Help-form preparation does not proceed.
' Returns: None (Sub). Results are returned through the documented ByRef parameters and/or the routine's database, UI, or module-state side effects.
Public Sub CreateChatGPTHelpForm(ByVal replaceExisting As Boolean, ByRef builtForms As Collection, ByRef skippedForms As Collection)
    On Error GoTo ErrHandler
    Dim frm As Form
    Dim formName As String
    Dim ctl As control
    Dim controlDefinitions As Collection
    Dim callbackExpressions As Collection
    formName = ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Help")
    If PrepareChatGPTGeneratedFormForBuild(formName, replaceExisting, skippedForms) = False Then Exit Sub
    Set controlDefinitions = GeneratedFormContractControlDefinitions(formName)
    Set callbackExpressions = GeneratedFormContractCallbackExpressions(formName)
    Set frm = CreateBlankFormForDesign(formName)
    ApplyUnboundFormProperties frm, GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Help")), DEFAULT_UI_HELP_FORM_WIDTH, DEFAULT_UI_HELP_FORM_HEIGHT
    EnsureHeaderFooterSections formName
    forms(formName).Section(acHeader).Height = 950
    forms(formName).Section(acDetail).Height = 8650
    forms(formName).OnLoad = GeneratedFormContractOnLoadExpression(formName)
    With SharedFormDesign_AddTitleLabel(formName, acHeader, vbNullString, GeneratedFormDisplayName(ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Help")), DEFAULT_UI_FORM_LEFT_MARGIN, DEFAULT_UI_FORM_TOP_MARGIN, 4800, 420)
        .TextAlign = 1
    End With
    With SharedFormDesign_AddPlainLabel(formName, acHeader, vbNullString, "Search current " & DEFAULT_UTILITY_SHORT_NAME & " workflows, the complete Add-Ins inventory and review workflow, or the table-backed reference for managed objects and fields.", DEFAULT_UI_FORM_LEFT_MARGIN, 540, 8700, 260, DEFAULT_UI_FONT_SIZE)
        .TextAlign = 1
    End With
    With SharedFormDesign_AddPlainLabel(formName, acHeader, vbNullString, "Search", 9250, 175, 750, 260, DEFAULT_UI_FONT_SIZE)
        .TextAlign = 1
    End With
    Set ctl = ChatGPTHelpCreateTextBox(formName, GeneratedFormContractControlName(formName, 1, controlDefinitions), acHeader, vbNullString, 10020, 120, 3100, 330, False, False)
    ctl.ControlTipText = "Search fixed Help topics plus table-backed object and field documentation."
    SharedFormDesign_AddButton formName, acHeader, vbNullString, "Search", GeneratedFormContractCallbackExpression(formName, 1, callbackExpressions), 13260, 100, 1050, 360, "Search Help and database-reference data."
    SharedFormDesign_AddButton formName, acHeader, vbNullString, "Clear", GeneratedFormContractCallbackExpression(formName, 2, callbackExpressions), 14340, 100, 1050, 360, "Clear the search and return to Getting Started."
    SharedFormDesign_AddButton formName, acHeader, vbNullString, "Home", GeneratedFormContractCallbackExpression(formName, 3, callbackExpressions), 9250, 520, 1050, 330, "Return to the Help home topic."
    SharedFormDesign_AddButton formName, acHeader, vbNullString, "Back", GeneratedFormContractCallbackExpression(formName, 4, callbackExpressions), 10330, 520, 1050, 330, "Return to the previous Help item."
    SharedFormDesign_AddButton formName, acHeader, vbNullString, "Forward", GeneratedFormContractCallbackExpression(formName, 5, callbackExpressions), 11410, 520, 1050, 330, "Move forward in Help history."
    SharedFormDesign_AddButton formName, acHeader, vbNullString, "Previous", GeneratedFormContractCallbackExpression(formName, 6, callbackExpressions), 12490, 520, 1050, 330, "Open the previous topic, object, or field in the current list."
    SharedFormDesign_AddButton formName, acHeader, vbNullString, "Next", GeneratedFormContractCallbackExpression(formName, 7, callbackExpressions), 13570, 520, 850, 330, "Open the next topic, object, or field in the current list."
    SharedFormDesign_AddButton formName, acHeader, vbNullString, "Open Related", GeneratedFormContractCallbackExpression(formName, 8, callbackExpressions), 14450, 520, 940, 330, "Open the related form, object, or system command when one is available."
    With SharedFormDesign_AddSectionLabel(formName, acDetail, vbNullString, "Sections", DEFAULT_UI_FORM_LEFT_MARGIN, DEFAULT_UI_FORM_TOP_MARGIN, 2800, 330)
        .TextAlign = 1
    End With
    Call ChatGPTHelpCreateListBox(formName, GeneratedFormContractControlName(formName, 2, controlDefinitions), DEFAULT_UI_FORM_LEFT_MARGIN, DEFAULT_UI_FORM_TOP_MARGIN + 330 + DEFAULT_UI_SECTION_TITLE_CONTENT_GAP, 2900, 7600, "=ChatGPT_HelpSelectSection()")
    With SharedFormDesign_AddSectionLabel(formName, acDetail, vbNullString, "Topics / Objects", 3260, 180, 3900, 330)
        .TextAlign = 1
    End With
    Call ChatGPTHelpCreateListBox(formName, GeneratedFormContractControlName(formName, 3, controlDefinitions), 3260, 560, 4100, 3800, "=ChatGPT_HelpSelectTopic()")
    With SharedFormDesign_AddSectionLabel(formName, acDetail, vbNullString, "Fields", 3260, 4520, 3900, 330)
        .TextAlign = 1
    End With
    Call ChatGPTHelpCreateListBox(formName, GeneratedFormContractControlName(formName, 4, controlDefinitions), 3260, 4900, 4100, 3260, "=ChatGPT_HelpSelectField()")
    With SharedFormDesign_AddSectionLabel(formName, acDetail, vbNullString, "Topic", 7500, 180, 7500, 330)
        .TextAlign = 1
    End With
    Set ctl = ChatGPTHelpCreateTextBox(formName, GeneratedFormContractControlName(formName, 5, controlDefinitions), acDetail, vbNullString, 7500, 560, 7800, 400, True, False)
    ctl.fontSize = DEFAULT_UI_TITLE_FONT_SIZE
    ctl.fontBold = True
    Set ctl = ChatGPTHelpCreateTextBox(formName, GeneratedFormContractControlName(formName, 6, controlDefinitions), acDetail, vbNullString, 7500, 1000, 7800, 320, True, False)
    ctl.FontItalic = True
    Set ctl = ChatGPTHelpCreateTextBox(formName, GeneratedFormContractControlName(formName, 7, controlDefinitions), acDetail, vbNullString, 7500, 1380, 7800, 6780, True, True)
    ctl.fontSize = DEFAULT_UI_READING_FONT_SIZE
    Set ctl = ChatGPTHelpCreateTextBox(formName, GeneratedFormContractControlName(formName, 8, controlDefinitions), acDetail, vbNullString, 7500, 8200, 100, 100, True, False)
    ctl.Visible = False
    Set ctl = ChatGPTHelpCreateTextBox(formName, GeneratedFormContractControlName(formName, 9, controlDefinitions), acDetail, vbNullString, 7620, 8200, 100, 100, True, False)
    ctl.Visible = False
    Set ctl = ChatGPTHelpCreateTextBox(formName, GeneratedFormContractControlName(formName, 10, controlDefinitions), acDetail, vbNullString, 7740, 8200, 100, 100, True, False)
    ctl.Visible = False
    Set ctl = ChatGPTHelpCreateTextBox(formName, GeneratedFormContractControlName(formName, 11, controlDefinitions), acDetail, vbNullString, 7860, 8200, 100, 100, True, False)
    ctl.Visible = False
    Set ctl = ChatGPTHelpCreateTextBox(formName, GeneratedFormContractControlName(formName, 12, controlDefinitions), acDetail, vbNullString, 7980, 8200, 100, 100, True, False)
    ctl.Visible = False
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
    Err.Raise savedErrNumber, "CreateChatGPTHelpForm", savedErrDescription & vbCrLf & "Error source: " & savedErrSource
End Sub

' Purpose: Validates generated form help specialized for DB, supplied help topics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   db (Input; DAO.Database): DAO.Database instance whose schema, data, queries, or properties the routine reads or modifies.
'   suppliedHelpTopics (Input Optional; Variant): Caller-supplied Help-topic array/list to use instead of reloading the default topic source; omitted/Empty follows normal resolution.
'   suppliedSystemMacroDefinitions (Input Optional; Variant): Caller-supplied supplied system macro definitions data used as the authoritative definition input for this call, avoiding an unnecessary reload when
'   already available.
'   suppliedSystemMacroDefinitionIndex (Input Optional; Variant): Caller-supplied supplied system macro definition index data used as the authoritative definition input for this call, avoiding an unnecessary
'   reload when already available.
'   suppliedRequiredRelatedFormNames (Input Optional; Variant): Caller-supplied Collection of physical generated-form names whose Help FORM targets must exist in the current validation partition; omitted
'   means all related forms are required.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub ValidateGeneratedFormHelpSpecialized(ByVal db As DAO.Database, Optional ByVal suppliedHelpTopics As Variant, Optional ByVal suppliedSystemMacroDefinitions As Variant, Optional ByVal suppliedSystemMacroDefinitionIndex As Variant, Optional ByVal suppliedRequiredRelatedFormNames As Variant)
    Dim helpTopics As Collection
    Dim systemMacroDefinitions As Collection
    Dim systemMacroDefinitionIndex As Object
    Dim requiredRelatedFormNames As Collection
    If isMissing(suppliedHelpTopics) = False Then
        If IsObject(suppliedHelpTopics) Then Set helpTopics = suppliedHelpTopics
    End If
    If isMissing(suppliedSystemMacroDefinitions) = False Then
        If IsObject(suppliedSystemMacroDefinitions) Then Set systemMacroDefinitions = suppliedSystemMacroDefinitions
    End If
    If isMissing(suppliedSystemMacroDefinitionIndex) = False Then
        If IsObject(suppliedSystemMacroDefinitionIndex) Then Set systemMacroDefinitionIndex = suppliedSystemMacroDefinitionIndex
    End If
    If isMissing(suppliedRequiredRelatedFormNames) = False Then
        If IsObject(suppliedRequiredRelatedFormNames) Then Set requiredRelatedFormNames = suppliedRequiredRelatedFormNames
    End If
    If helpTopics Is Nothing Then
        If systemMacroDefinitions Is Nothing Then Set systemMacroDefinitions = x__UTIL_ADC_DEFAULTS.GetSystemMacroDefinitions()
        If systemMacroDefinitionIndex Is Nothing Then Set systemMacroDefinitionIndex = ChatGPTSystemMacroDefinitionIndex(systemMacroDefinitions)
        Set helpTopics = ChatGPTHelpTopics(systemMacroDefinitions, systemMacroDefinitionIndex)
    ElseIf systemMacroDefinitionIndex Is Nothing Then
        If systemMacroDefinitions Is Nothing Then Set systemMacroDefinitions = x__UTIL_ADC_DEFAULTS.GetSystemMacroDefinitions()
        Set systemMacroDefinitionIndex = ChatGPTSystemMacroDefinitionIndex(systemMacroDefinitions)
    End If
    ValidateChatGPTHelpDefinitions helpTopics, systemMacroDefinitions, systemMacroDefinitionIndex
    If db Is Nothing Then Err.Raise vbObjectError + 7347, "ValidateGeneratedFormHelpSpecialized", "No database was supplied for generated Help validation."
    If SharedDao_TableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID)) = False Then Err.Raise vbObjectError + 7348, "ValidateGeneratedFormHelpSpecialized", "The generated Help Database Reference source is missing." & vbCrLf & "Table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID)
    If SharedDao_TableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFields")) = False Then Err.Raise vbObjectError + 7349, "ValidateGeneratedFormHelpSpecialized", "The generated Help field-documentation source is missing." & vbCrLf & "Table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFields")
    If SharedDao_TableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFieldValues")) = False Then Err.Raise vbObjectError + 7356, "ValidateGeneratedFormHelpSpecialized", "The generated Help allowed-value documentation source is missing." & vbCrLf & "Table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFieldValues")
    If SharedDao_TableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectDependencies")) = False Then Err.Raise vbObjectError + 7357, "ValidateGeneratedFormHelpSpecialized", "The generated Help dependency-documentation source is missing." & vbCrLf & "Table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectDependencies")
    ValidateChatGPTHelpRelatedTargets db, helpTopics, systemMacroDefinitions, systemMacroDefinitionIndex, requiredRelatedFormNames
End Sub

Private Sub ValidateChatGPTHelpRelatedTargets(ByVal db As DAO.Database, Optional ByVal helpTopics As Variant, Optional ByVal suppliedSystemMacroDefinitions As Variant, Optional ByVal suppliedSystemMacroDefinitionIndex As Variant, Optional ByVal suppliedRequiredRelatedFormNames As Variant)
    On Error GoTo ErrHandler
    Dim topics As Collection
    Dim topic As Variant
    Dim relatedTarget As String
    Dim targetParts As Variant
    Dim targetType As String
    Dim targetName As String
    Dim objectTypeText As String
    Dim objectTypeCode As AcObjectType
    Dim validatedCount As Long
    Dim deferredFormTargetCount As Long
    Dim systemMacroDefinitions As Collection
    Dim systemMacroDefinitionIndex As Object
    Dim requiredRelatedFormNames As Collection
    Dim requiredRelatedFormItem As Variant
    Dim relatedFormRequired As Boolean
    If db Is Nothing Then Err.Raise vbObjectError + 7358, "ValidateChatGPTHelpRelatedTargets", "No database was supplied for Help related-target validation."
    If IsObject(helpTopics) Then Set topics = helpTopics
    If isMissing(suppliedSystemMacroDefinitions) = False Then
        If IsObject(suppliedSystemMacroDefinitions) Then Set systemMacroDefinitions = suppliedSystemMacroDefinitions
    End If
    If isMissing(suppliedSystemMacroDefinitionIndex) = False Then
        If IsObject(suppliedSystemMacroDefinitionIndex) Then Set systemMacroDefinitionIndex = suppliedSystemMacroDefinitionIndex
    End If
    If isMissing(suppliedRequiredRelatedFormNames) = False Then
        If IsObject(suppliedRequiredRelatedFormNames) Then Set requiredRelatedFormNames = suppliedRequiredRelatedFormNames
    End If
    If topics Is Nothing Then
        If systemMacroDefinitions Is Nothing Then Set systemMacroDefinitions = x__UTIL_ADC_DEFAULTS.GetSystemMacroDefinitions()
        If systemMacroDefinitionIndex Is Nothing Then Set systemMacroDefinitionIndex = ChatGPTSystemMacroDefinitionIndex(systemMacroDefinitions)
        Set topics = ChatGPTHelpTopics(systemMacroDefinitions, systemMacroDefinitionIndex)
    ElseIf systemMacroDefinitionIndex Is Nothing Then
        If systemMacroDefinitions Is Nothing Then Set systemMacroDefinitions = x__UTIL_ADC_DEFAULTS.GetSystemMacroDefinitions()
        Set systemMacroDefinitionIndex = ChatGPTSystemMacroDefinitionIndex(systemMacroDefinitions)
    End If
    For Each topic In topics
        SetupControlPump
        relatedTarget = Trim$(CStr(topic(4)))
        If Len(relatedTarget) > 0 Then
            targetParts = Split(relatedTarget, ":")
            targetType = UCase$(Trim$(CStr(targetParts(0))))
            Select Case targetType
                Case "T"
                    validatedCount = validatedCount + 1
                Case "COMMAND"
                    targetName = Trim$(Mid$(relatedTarget, InStr(1, relatedTarget, ":", vbBinaryCompare) + 1))
                    If ChatGPTHelpCommandTargetKnown(targetName, systemMacroDefinitionIndex) = False Then Err.Raise vbObjectError + 7359, "ValidateChatGPTHelpRelatedTargets", "A generated Help related command is not part of the authoritative command set." & vbCrLf & "Topic: " & CStr(topic(0)) & "|" & CStr(topic(1)) & vbCrLf & "Command: " & targetName
                    validatedCount = validatedCount + 1
                Case "FORM"
                    targetName = Trim$(Mid$(relatedTarget, InStr(1, relatedTarget, ":", vbBinaryCompare) + 1))
                    relatedFormRequired = True
                    If Not requiredRelatedFormNames Is Nothing Then
                        relatedFormRequired = False
                        For Each requiredRelatedFormItem In requiredRelatedFormNames
                            If StrComp(CStr(requiredRelatedFormItem), targetName, vbTextCompare) = 0 Then
                                relatedFormRequired = True
                                Exit For
                            End If
                        Next requiredRelatedFormItem
                    End If
                    If relatedFormRequired Then
                        If SharedDao_AccessObjectExists(db, acForm, targetName) = False Then Err.Raise vbObjectError + 7360, "ValidateChatGPTHelpRelatedTargets", "A generated Help related form does not exist after generated-form reconciliation." & vbCrLf & "Topic: " & CStr(topic(0)) & "|" & CStr(topic(1)) & vbCrLf & "Form: " & targetName
                        validatedCount = validatedCount + 1
                    Else
                        deferredFormTargetCount = deferredFormTargetCount + 1
                    End If
                Case "OBJECT"
                    If UBound(targetParts) < 2 Then Err.Raise vbObjectError + 7361, "ValidateChatGPTHelpRelatedTargets", "A generated Help object target is incomplete." & vbCrLf & "Target: " & relatedTarget
                    objectTypeText = NormalizedObjectTypeCode(CStr(targetParts(1)))
                    targetName = Trim$(CStr(targetParts(2)))
                    If Len(objectTypeText) = 0 Then Err.Raise vbObjectError + 7362, "ValidateChatGPTHelpRelatedTargets", "A generated Help object target uses an unsupported Access object type." & vbCrLf & "Target: " & relatedTarget
                    objectTypeCode = AccessObjectTypeFromCode(objectTypeText)
                    If SharedDao_AccessObjectExists(db, objectTypeCode, targetName) = False Then Err.Raise vbObjectError + 7363, "ValidateChatGPTHelpRelatedTargets", "A generated Help related object does not exist after setup reconciliation." & vbCrLf & "Topic: " & CStr(topic(0)) & "|" & CStr(topic(1)) & vbCrLf & "Object type: " & objectTypeText & vbCrLf & "Object: " & targetName
                    validatedCount = validatedCount + 1
                Case Else
                    Err.Raise vbObjectError + 7364, "ValidateChatGPTHelpRelatedTargets", "A generated Help related target uses an unsupported target type." & vbCrLf & "Target: " & relatedTarget
            End Select
        End If
    Next topic
    LogTraceMessage TRACE_LEVEL_INFO, "UI", "ValidateChatGPTHelpRelatedTargets", "Generated Help related targets passed strict validation.", "Topics inspected: " & CStr(topics.count) & vbCrLf & "Related targets validated: " & CStr(validatedCount) & vbCrLf & "Related form targets deferred to another generated-form partition: " & CStr(deferredFormTargetCount), "Validate generated Help related targets", 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Help"), "Form", "CodeDb", "Help"
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ValidateChatGPTHelpRelatedTargets", "Generated Help related-target validation failed." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Function ChatGPTHelpSystemMacroTitleList(ByVal systemMacroKeys As Collection, Optional ByVal suppliedSystemMacroDefinitionIndex As Variant) As String
    Dim item As Variant
    Dim Definition As Variant
    Dim systemMacroDefinitionIndex As Object
    Dim TitleText As String
    Dim resultText As String
    If systemMacroKeys Is Nothing Then Exit Function
    If isMissing(suppliedSystemMacroDefinitionIndex) = False Then
        If IsObject(suppliedSystemMacroDefinitionIndex) Then Set systemMacroDefinitionIndex = suppliedSystemMacroDefinitionIndex
    End If
    For Each item In systemMacroKeys
        If systemMacroDefinitionIndex Is Nothing Then
            TitleText = Trim$(SystemMacroDefaultTitleText(CStr(item)))
        ElseIf systemMacroDefinitionIndex.Exists(CStr(item)) Then
            Definition = systemMacroDefinitionIndex(CStr(item))
            TitleText = Trim$(CStr(Definition(2)))
        Else
            TitleText = vbNullString
        End If
        If Len(TitleText) > 0 Then
            If Len(resultText) > 0 Then resultText = resultText & ", "
            resultText = resultText & TitleText
        End If
    Next item
    ChatGPTHelpSystemMacroTitleList = resultText
End Function

Private Sub ChatGPTHelpShowManagedObject(ByVal managedObjectID As Long)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim fieldsRs As DAO.Recordset
    Dim dependencyRs As DAO.Recordset
    Dim contentText As String
    Dim rowSourceText As String
    Dim objectName As String
    Dim objectType As String
    Dim displayName As String
    Dim managedObjectsTableName As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Set db = GetHostDb()
    managedObjectsTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID)
    Set rs = db.OpenRecordset("SELECT * FROM [" & managedObjectsTableName & "] WHERE " & SQL_FIELD_ID & "=" & CStr(managedObjectID) & ";", dbOpenSnapshot)
    If rs.EOF Then Err.Raise vbObjectError + 6781, "ChatGPTHelpShowManagedObject", "The selected " & managedObjectsTableName & " row no longer exists."
    objectName = CStr(Nz(rs.fields(FIELD_OBJECT_NAME).value, vbNullString))
    objectType = CStr(Nz(rs.fields(FIELD_OBJECT_TYPE).value, vbNullString))
    displayName = CStr(Nz(rs.fields(FIELD_OBJECT_DISPLAY_NAME).value, objectName))
    contentText = ChatGPTHelpLabeledText("Object", objectName)
    contentText = contentText & ChatGPTHelpLabeledText("Type", objectType)
    contentText = contentText & ChatGPTHelpLabeledText("Description", IIf(Len(CStr(Nz(rs.fields(FIELD_OBJECT_DESCRIPTION).value, vbNullString))) > 0, CStr(Nz(rs.fields(FIELD_OBJECT_DESCRIPTION).value, vbNullString)), "Description missing."))
    contentText = contentText & ChatGPTHelpLabeledText("Usage", CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_USAGE_GUIDANCE).value, vbNullString)))
    contentText = contentText & ChatGPTHelpLabeledText("Edit policy", CStr(Nz(rs.fields(FIELD_EDIT_POLICY).value, vbNullString)))
    contentText = contentText & ChatGPTHelpLabeledText("Lifecycle", CStr(Nz(rs.fields(FIELD_LIFECYCLE_POLICY).value, vbNullString)))
    contentText = contentText & ChatGPTHelpLabeledText("Repair owner", CStr(Nz(rs.fields(FIELD_REPAIR_OWNER).value, vbNullString)))
    contentText = contentText & ChatGPTHelpLabeledText("Deletion policy", CStr(Nz(rs.fields(FIELD_DELETION_POLICY).value, vbNullString)))
    contentText = contentText & ChatGPTHelpLabeledText("Primary workflow", CStr(Nz(rs.fields(FIELD_PRIMARY_WORKFLOW).value, vbNullString)))
    contentText = contentText & ChatGPTHelpLabeledText("Category / owner", CStr(Nz(rs.fields(FIELD_CATEGORY_CODE).value, vbNullString)) & " / " & CStr(Nz(rs.fields(FIELD_CATALOG_OWNER).value, vbNullString)))
    contentText = contentText & ChatGPTHelpLabeledText("Navigation Pane setting", CStr(Nz(rs.fields(FIELD_NAVIGATION_PANE_GROUP_SETTING_KEY).value, vbNullString)))
    contentText = contentText & ChatGPTHelpLabeledText("Ribbon placement", CStr(Nz(rs.fields(FIELD_RIBBON_TAB_LABEL).value, vbNullString)) & " > " & CStr(Nz(rs.fields(FIELD_RIBBON_GROUP_LABEL).value, vbNullString)))
    contentText = contentText & ChatGPTHelpLabeledText("Documentation", CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_DEPENDENCY_DOCUMENTATION_STATUS).value, vbNullString)) & "; source: " & CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_DOCUMENTATION_SOURCE).value, vbNullString)))
    contentText = contentText & ChatGPTHelpLabeledText("Schema verified", ChatGPTHelpDateText(rs.fields(FIELD_MANAGED_OBJECT_FIELD_SCHEMA_VERIFIED_ON).value))
    rs.Close
    If SharedDao_TableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFields")) Then
        Set fieldsRs = db.OpenRecordset("SELECT " & SQL_FIELD_ID & "," & SQL_FIELD_MANAGED_OBJECT_FIELD_NAME & "," & SQL_FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME & "," & SQL_FIELD_MANAGED_OBJECT_FIELD_DATA_TYPE & "," & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_DOCUMENTATION_STATUS & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFields") & "] WHERE " & SQL_FIELD_MANAGED_OBJECT_ID & "=" & CStr(managedObjectID) & " AND " & SQL_FIELD_MANAGED_OBJECT_FIELD_IS_ACTIVE & "=True ORDER BY " & SQL_FIELD_MANAGED_OBJECT_FIELD_ORDINAL & "," & SQL_FIELD_MANAGED_OBJECT_FIELD_NAME & ";", dbOpenSnapshot)
        Do While Not fieldsRs.EOF
            SetupControlPump
            rowSourceText = ChatGPTHelpValueListPair(rowSourceText, "F:" & CStr(fieldsRs.fields(FIELD_ID).value), CStr(Nz(fieldsRs.fields(FIELD_MANAGED_OBJECT_FIELD_NAME).value, vbNullString)) & " - " & CStr(Nz(fieldsRs.fields(FIELD_MANAGED_OBJECT_FIELD_DATA_TYPE).value, vbNullString)) & ChatGPTHelpStatusSuffix(CStr(Nz(fieldsRs.fields(FIELD_MANAGED_OBJECT_DEPENDENCY_DOCUMENTATION_STATUS).value, vbNullString))))
            fieldsRs.MoveNext
        Loop
        fieldsRs.Close
    End If
    ChatGPTHelpSetListRows "lstHelpFields", rowSourceText
    contentText = contentText & ChatGPTHelpObjectDependenciesText(db, managedObjectID)
    ChatGPTHelpDisplay displayName, "Database Reference > " & objectType, contentText, ChatGPTHelpObjectTarget(objectType, objectName)
    ChatGPTHelpSelectListKey "lstHelpSections", "DBREF"
    ChatGPTHelpSelectListKey "lstHelpTopics", "O:" & CStr(managedObjectID)
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not fieldsRs Is Nothing Then fieldsRs.Close
    If Not dependencyRs Is Nothing Then dependencyRs.Close
    On Error GoTo 0
    Err.Raise savedErrNumber, "ChatGPTHelpShowManagedObject", "Could not display the selected documented object." & vbCrLf & "ManagedObjectID: " & CStr(managedObjectID) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Sub

Private Sub ChatGPTHelpShowManagedObjectField(ByVal managedObjectFieldID As Long)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim valuesRs As DAO.Recordset
    Dim contentText As String
    Dim objectName As String
    Dim fieldName As String
    Dim valuesText As String
    Dim fieldSizeText As String
    Dim validationText As String
    Dim lookupText As String
    Dim managedObjectsTableName As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Set db = GetHostDb()
    managedObjectsTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID)
    Set rs = db.OpenRecordset("SELECT f.*,o." & SQL_FIELD_OBJECT_NAME & ",o." & SQL_FIELD_OBJECT_TYPE & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFields") & "] AS f INNER JOIN [" & managedObjectsTableName & "] AS o ON f." & SQL_FIELD_MANAGED_OBJECT_ID & "=o." & SQL_FIELD_ID & " WHERE f." & SQL_FIELD_ID & "=" & CStr(managedObjectFieldID) & ";", dbOpenSnapshot)
    If rs.EOF Then Err.Raise vbObjectError + 6782, "ChatGPTHelpShowManagedObjectField", "The selected MAC_Fields row no longer exists."
    objectName = CStr(Nz(rs.fields(FIELD_OBJECT_NAME).value, vbNullString))
    fieldName = CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_NAME).value, vbNullString))
    contentText = ChatGPTHelpLabeledText("Object", objectName)
    contentText = contentText & ChatGPTHelpLabeledText("Field", fieldName)
    contentText = contentText & ChatGPTHelpLabeledText("Description", IIf(Len(CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_DESCRIPTION_LONG).value, vbNullString))) > 0, CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_DESCRIPTION_LONG).value, vbNullString)), "Description missing."))
    contentText = contentText & ChatGPTHelpLabeledText("Usage", CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_USAGE_GUIDANCE).value, vbNullString)))
    fieldSizeText = CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_ACTUAL_SIZE).value, 0))
    If CLng(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_ACTUAL_SIZE).value, 0)) > 0 Then fieldSizeText = " / " & fieldSizeText Else fieldSizeText = vbNullString
    validationText = CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_VALIDATION_TEXT).value, vbNullString))
    lookupText = CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_LOOKUP_ROW_SOURCE).value, vbNullString))
    contentText = contentText & ChatGPTHelpLabeledText("Type / size", CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_DATA_TYPE).value, vbNullString)) & fieldSizeText)
    contentText = contentText & ChatGPTHelpLabeledText("Order", CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_ORDINAL).value, 0)))
    contentText = contentText & ChatGPTHelpLabeledText("Required / zero length", CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_DEPENDENCY_IS_REQUIRED).value, False)) & " / " & CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_ALLOWS_ZERO_LENGTH).value, False)))
    contentText = contentText & ChatGPTHelpLabeledText("Primary key / indexed / unique", CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_IS_PRIMARY_KEY).value, False)) & " / " & CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_IS_INDEXED).value, False)) & " / " & CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_IS_UNIQUE).value, False)))
    contentText = contentText & ChatGPTHelpLabeledText("Default", CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_DEFAULT_VALUE).value, vbNullString)))
    contentText = contentText & ChatGPTHelpLabeledText("Validation", CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_VALIDATION_RULE).value, vbNullString)) & IIf(Len(validationText) > 0, " - " & validationText, vbNullString))
    contentText = contentText & ChatGPTHelpLabeledText("Lookup", CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_LOOKUP_ROW_SOURCE_TYPE).value, vbNullString)) & IIf(Len(lookupText) > 0, ": " & lookupText, vbNullString))
    contentText = contentText & ChatGPTHelpLabeledText("Direct edit supported / structure system-maintained", CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_IS_USER_EDITABLE).value, False)) & " / " & CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_IS_SYSTEM_MAINTAINED).value, False)))
    contentText = contentText & ChatGPTHelpLabeledText("Value preservation", CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_PRESERVATION_POLICY).value, vbNullString)))
    contentText = contentText & ChatGPTHelpLabeledText("Documentation", CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_DEPENDENCY_DOCUMENTATION_STATUS).value, vbNullString)) & "; source: " & CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_DOCUMENTATION_SOURCE).value, vbNullString)))
    contentText = contentText & ChatGPTHelpLabeledText("Schema verified", ChatGPTHelpDateText(rs.fields(FIELD_MANAGED_OBJECT_FIELD_SCHEMA_VERIFIED_ON).value))
    rs.Close
    If SharedDao_TableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFieldValues")) Then
        Set valuesRs = db.OpenRecordset("SELECT " & SQL_FIELD_MANAGED_OBJECT_FIELD_VALUE_STORED_VALUE & "," & SQL_FIELD_MANAGED_OBJECT_FIELD_VALUE_DISPLAY_VALUE & "," & SQL_FIELD_MANAGED_OBJECT_FIELD_VALUE_DESCRIPTION & "," & SQL_FIELD_MANAGED_OBJECT_FIELD_VALUE_IS_DEFAULT & "," & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_DOCUMENTATION_STATUS & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFieldValues") & "] WHERE " & SQL_FIELD_MANAGED_OBJECT_FIELD_VALUE_MANAGED_OBJECT_FIELD_ID & "=" & CStr(managedObjectFieldID) & " AND " & SQL_FIELD_MANAGED_OBJECT_FIELD_IS_ACTIVE & "=True ORDER BY " & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER & "," & SQL_FIELD_MANAGED_OBJECT_FIELD_VALUE_DISPLAY_VALUE & ";", dbOpenSnapshot)
        Do While Not valuesRs.EOF
            SetupControlPump
            valuesText = valuesText & IIf(Len(valuesText) = 0, vbNullString, vbCrLf) & "- " & CStr(Nz(valuesRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_STORED_VALUE).value, vbNullString))
            If StrComp(CStr(Nz(valuesRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_DISPLAY_VALUE).value, vbNullString)), CStr(Nz(valuesRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_STORED_VALUE).value, vbNullString)), vbTextCompare) <> 0 Then valuesText = valuesText & " (" & CStr(Nz(valuesRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_DISPLAY_VALUE).value, vbNullString)) & ")"
            If Nz(valuesRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_IS_DEFAULT).value, False) Then valuesText = valuesText & " [default]"
            If Len(CStr(Nz(valuesRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_DESCRIPTION).value, vbNullString))) > 0 Then valuesText = valuesText & ": " & CStr(valuesRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_DESCRIPTION).value)
            valuesRs.MoveNext
        Loop
        valuesRs.Close
    End If
    If Len(valuesText) > 0 Then contentText = contentText & vbCrLf & "Allowed values" & vbCrLf & valuesText & vbCrLf
    ChatGPTHelpDisplay fieldName, "Database Reference > " & objectName & " > Field", contentText, vbNullString
    ChatGPTHelpSelectListKey "lstHelpFields", "F:" & CStr(managedObjectFieldID)
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not valuesRs Is Nothing Then valuesRs.Close
    On Error GoTo 0
    Err.Raise savedErrNumber, "ChatGPTHelpShowManagedObjectField", "Could not display the selected documented field." & vbCrLf & "ManagedObjectFieldID: " & CStr(managedObjectFieldID) & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Sub

' =============================================================================================================================================================================
' GENERATED HELP PUBLIC CALLBACK PROCEDURES
' =============================================================================================================================================================================

' Purpose: Initializes the generated Help Center form, loads its navigation state, and renders the initial help content.
' Public role: Form-code/event-expression entry point. Generated or managed Access forms can invoke this Public routine from form code or an event expression.
' Parameters: None.
' Returns: Boolean: True when chat gpt help select field completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or not-applicable path.
Public Function ChatGPT_HelpInitialize() As Boolean
    ' Wrapper purpose: Generated Help Center callback that captures the Help form, resets per-instance state, loads sections, and navigates to the Start section.
    ' Why retained: Retained because the generated form's event expression needs a stable public Function while Help state and navigation internals remain private.
    On Error GoTo ErrHandler
    ChatGPTHelpCaptureFormReference
    ChatGPTHelpResetInstanceState
    ChatGPTHelpLoadSections
    ChatGPTHelpSelectSectionKey "START", True
    ChatGPT_HelpInitialize = True
CleanExit:
    ChatGPTHelpReleaseFormReference
    Exit Function
ErrHandler:
    ChatGPTHelpShowRuntimeError "Initialize Help", Err.Number, Err.Source, Err.Description, "ChatGPT_HelpInitialize"
    Resume CleanExit
End Function

' Purpose: Applies the section selected in the Help Center and refreshes the dependent topic/field choices and displayed content.
' Public role: Form-code/event-expression entry point. Generated or managed Access forms can invoke this Public routine from form code or an event expression.
' Parameters: None.
' Returns: Boolean: True when chat gpt help search completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or not-applicable path.
Public Function ChatGPT_HelpSelectSection() As Boolean
    ' Wrapper purpose: Generated Help Center callback that converts the current section-list selection into Help-engine section navigation.
    ' Why retained: Retained because the generated list event needs a stable public Function while list interpretation and navigation internals remain private.
    On Error GoTo ErrHandler
    Dim sectionKey As String
    ChatGPTHelpCaptureFormReference
    sectionKey = ChatGPTHelpSelectedListKey("lstHelpSections")
    If Len(sectionKey) = 0 Then GoTo CleanExit
    ChatGPTHelpSelectSectionKey sectionKey, True
    ChatGPT_HelpSelectSection = True
CleanExit:
    ChatGPTHelpReleaseFormReference
    Exit Function
ErrHandler:
    ChatGPTHelpShowRuntimeError "Select Help section", Err.Number, Err.Source, Err.Description, "ChatGPT_HelpSelectSection"
    Resume CleanExit
End Function

' Purpose: Applies the topic selected in the Help Center and refreshes the dependent field choices and displayed content.
' Public role: Form-code/event-expression entry point. Generated or managed Access forms can invoke this Public routine from form code or an event expression.
' Parameters: None.
' Returns: Boolean: True when chat gpt help clear search completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or not-applicable path.
Public Function ChatGPT_HelpSelectTopic() As Boolean
    ' Wrapper purpose: Generated Help Center callback that converts the current topic-list selection into Help-item navigation.
    ' Why retained: Retained because the generated list event needs a stable public Function while topic selection and Help navigation internals remain private.
    On Error GoTo ErrHandler
    Dim itemKey As String
    ChatGPTHelpCaptureFormReference
    itemKey = ChatGPTHelpSelectedListKey("lstHelpTopics")
    If Len(itemKey) = 0 Then GoTo CleanExit
    ChatGPTHelpLoadKey itemKey, True
    ChatGPT_HelpSelectTopic = True
CleanExit:
    ChatGPTHelpReleaseFormReference
    Exit Function
ErrHandler:
    ChatGPTHelpShowRuntimeError "Select Help topic", Err.Number, Err.Source, Err.Description, "ChatGPT_HelpSelectTopic"
    Resume CleanExit
End Function

' Purpose: Applies the field selected in the Help Center and refreshes the displayed help content for that selection.
' Public role: Form-code/event-expression entry point. Generated or managed Access forms can invoke this Public routine from form code or an event expression.
' Parameters: None.
' Returns: Boolean: True when chat gpt help go home completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or not-applicable path.
Public Function ChatGPT_HelpSelectField() As Boolean
    ' Wrapper purpose: Generated Help Center callback that converts the current field-list selection into Help-item navigation.
    ' Why retained: Retained because the generated list event needs a stable public Function while field selection and Help navigation internals remain private.
    On Error GoTo ErrHandler
    Dim itemKey As String
    ChatGPTHelpCaptureFormReference
    itemKey = ChatGPTHelpSelectedListKey("lstHelpFields")
    If Len(itemKey) = 0 Then GoTo CleanExit
    ChatGPTHelpLoadKey itemKey, True
    ChatGPT_HelpSelectField = True
CleanExit:
    ChatGPTHelpReleaseFormReference
    Exit Function
ErrHandler:
    ChatGPTHelpShowRuntimeError "Select documented field", Err.Number, Err.Source, Err.Description, "ChatGPT_HelpSelectField"
    Resume CleanExit
End Function

' Purpose: Runs the Help Center search using the current search controls and displays the matching help results.
' Public role: Form-code/event-expression entry point. Generated or managed Access forms can invoke this Public routine from form code or an event expression.
' Parameters: None.
' Returns: Boolean: True when chat gpt help back completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or not-applicable path.
Public Function ChatGPT_HelpSearch() As Boolean
    ' Wrapper purpose: Generated Help Center callback that reads the search control and invokes the private Help search workflow.
    ' Why retained: Retained because the generated Search control needs a stable public Function entry point while search implementation and Help state remain private.
    On Error GoTo ErrHandler
    Dim searchText As String
    ChatGPTHelpCaptureFormReference
    searchText = Trim$(CStr(Nz(mChatGPTHelpForm.Controls("txtHelpSearch").value, vbNullString)))
    If Len(searchText) = 0 Then
        mChatGPTHelpForm.Controls("txtHelpSearch").value = Null
        ChatGPTHelpSelectSectionKey "START", True
        ChatGPT_HelpSearch = True
        GoTo CleanExit
    End If
    ChatGPTHelpLoadSearchResults searchText
    ChatGPT_HelpSearch = True
CleanExit:
    ChatGPTHelpReleaseFormReference
    Exit Function
ErrHandler:
    ChatGPTHelpShowRuntimeError "Search Help", Err.Number, Err.Source, Err.Description, "ChatGPT_HelpSearch"
    Resume CleanExit
End Function

' Purpose: Clears the Help Center search criteria/results and restores the normal navigation/content state.
' Public role: Form-code/event-expression entry point. Generated or managed Access forms can invoke this Public routine from form code or an event expression.
' Parameters: None.
' Returns: Boolean: True when chat gpt help forward completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or not-applicable path.
Public Function ChatGPT_HelpClearSearch() As Boolean
    ' Wrapper purpose: Generated Help Center callback that clears search UI state and restores the Start section.
    ' Why retained: Retained because the generated control needs a stable public Function entry point while search reset and navigation internals remain private.
    On Error GoTo ErrHandler
    ChatGPTHelpCaptureFormReference
    mChatGPTHelpForm.Controls("txtHelpSearch").value = Null
    ChatGPTHelpSelectSectionKey "START", True
    ChatGPT_HelpClearSearch = True
CleanExit:
    ChatGPTHelpReleaseFormReference
    Exit Function
ErrHandler:
    ChatGPTHelpShowRuntimeError "Clear Help search", Err.Number, Err.Source, Err.Description, "ChatGPT_HelpClearSearch"
    Resume CleanExit
End Function

' Purpose: Returns the Help Center to its home/root view and synchronizes the navigation controls to that location.
' Public role: Form-code/event-expression entry point. Generated or managed Access forms can invoke this Public routine from form code or an event expression.
' Parameters: None.
' Returns: Boolean: True when chat gpt help go home completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or not-applicable path.
Public Function ChatGPT_HelpGoHome() As Boolean
    ' Wrapper purpose: Generated Help Center callback that returns navigation to the Help Start section.
    ' Why retained: Retained because the generated Home control needs a stable public Function entry point while Help navigation implementation remains private.
    On Error GoTo ErrHandler
    ChatGPTHelpCaptureFormReference
    ChatGPTHelpSelectSectionKey "START", True
    ChatGPT_HelpGoHome = True
CleanExit:
    ChatGPTHelpReleaseFormReference
    Exit Function
ErrHandler:
    ChatGPTHelpShowRuntimeError "Open Help home", Err.Number, Err.Source, Err.Description, "ChatGPT_HelpGoHome"
    Resume CleanExit
End Function

' Purpose: Moves the Help Center to the preceding item in its navigation history when a prior item is available.
' Public role: Form-code/event-expression entry point. Generated or managed Access forms can invoke this Public routine from form code or an event expression.
' Parameters: None.
' Returns: Boolean: True when chat gpt help previous completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or not-applicable path.
Public Function ChatGPT_HelpBack() As Boolean
    ' Wrapper purpose: Generated Help Center callback that navigates backward through the private Help history state.
    ' Why retained: Retained because the generated Back control needs a stable public Function entry point without direct access to private history management.
    On Error GoTo ErrHandler
    Dim historyItems As Collection
    Dim historyIndex As Long
    ChatGPTHelpCaptureFormReference
    Set historyItems = ChatGPTHelpHistoryItems()
    historyIndex = ChatGPTHelpHistoryIndexValue()
    If historyIndex <= 1 Or historyItems.count = 0 Then GoTo CleanExit
    historyIndex = historyIndex - 1
    ChatGPTHelpSetHistoryIndexValue historyIndex
    ChatGPTHelpSetNavigatingHistoryValue True
    ChatGPTHelpLoadKey CStr(historyItems(historyIndex)), False
    ChatGPTHelpSetNavigatingHistoryValue False
    ChatGPT_HelpBack = True
CleanExit:
    ChatGPTHelpReleaseFormReference
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not mChatGPTHelpForm Is Nothing Then ChatGPTHelpSetNavigatingHistoryValue False
    On Error GoTo 0
    ChatGPTHelpShowRuntimeError "Navigate Help back", savedErrNumber, savedErrSource, savedErrDescription, "ChatGPT_HelpBack"
    Resume CleanExit
End Function

' Purpose: Moves the Help Center to the next item in its navigation history when a forward item is available.
' Public role: Form-code/event-expression entry point. Generated or managed Access forms can invoke this Public routine from form code or an event expression.
' Parameters: None.
' Returns: Boolean: True when chat gpt help open related completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or not-applicable path.
Public Function ChatGPT_HelpForward() As Boolean
    ' Wrapper purpose: Generated Help Center callback that navigates forward through the private Help history state.
    ' Why retained: Retained because the generated Forward control needs a stable public Function entry point without direct access to private history management.
    On Error GoTo ErrHandler
    Dim historyItems As Collection
    Dim historyIndex As Long
    ChatGPTHelpCaptureFormReference
    Set historyItems = ChatGPTHelpHistoryItems()
    historyIndex = ChatGPTHelpHistoryIndexValue()
    If historyItems.count = 0 Or historyIndex >= historyItems.count Then GoTo CleanExit
    historyIndex = historyIndex + 1
    ChatGPTHelpSetHistoryIndexValue historyIndex
    ChatGPTHelpSetNavigatingHistoryValue True
    ChatGPTHelpLoadKey CStr(historyItems(historyIndex)), False
    ChatGPTHelpSetNavigatingHistoryValue False
    ChatGPT_HelpForward = True
CleanExit:
    ChatGPTHelpReleaseFormReference
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not mChatGPTHelpForm Is Nothing Then ChatGPTHelpSetNavigatingHistoryValue False
    On Error GoTo 0
    ChatGPTHelpShowRuntimeError "Navigate Help forward", savedErrNumber, savedErrSource, savedErrDescription, "ChatGPT_HelpForward"
    Resume CleanExit
End Function

' Purpose: Moves the Help Center to the previous help item in the current ordered navigation sequence.
' Public role: Form-code/event-expression entry point. Generated or managed Access forms can invoke this Public routine from form code or an event expression.
' Parameters: None.
' Returns: Boolean: True when chat gpt help previous completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or not-applicable path.
Public Function ChatGPT_HelpPrevious() As Boolean
    ' Wrapper purpose: Generated Help Center callback that moves to the previous item in the active Help list.
    ' Why retained: Retained because the generated Previous control needs a stable public Function entry point while list navigation implementation remains private.
    On Error GoTo ErrHandler
    ChatGPTHelpCaptureFormReference
    ChatGPTHelpMoveWithinCurrentList -1
    ChatGPT_HelpPrevious = True
CleanExit:
    ChatGPTHelpReleaseFormReference
    Exit Function
ErrHandler:
    ChatGPTHelpShowRuntimeError "Open previous Help item", Err.Number, Err.Source, Err.Description, "ChatGPT_HelpPrevious"
    Resume CleanExit
End Function

' Purpose: Moves the Help Center to the next help item in the current ordered navigation sequence.
' Public role: Form-code/event-expression entry point. Generated or managed Access forms can invoke this Public routine from form code or an event expression.
' Parameters: None.
' Returns: Boolean: True when chat gpt help next completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or not-applicable path.
Public Function ChatGPT_HelpNext() As Boolean
    ' Wrapper purpose: Generated Help Center callback that moves to the next item in the active Help list.
    ' Why retained: Retained because the generated Next control needs a stable public Function entry point while list navigation implementation remains private.
    On Error GoTo ErrHandler
    ChatGPTHelpCaptureFormReference
    ChatGPTHelpMoveWithinCurrentList 1
    ChatGPT_HelpNext = True
CleanExit:
    ChatGPTHelpReleaseFormReference
    Exit Function
ErrHandler:
    ChatGPTHelpShowRuntimeError "Open next Help item", Err.Number, Err.Source, Err.Description, "ChatGPT_HelpNext"
    Resume CleanExit
End Function

' Purpose: Opens the related help item selected by the current Help Center record/control and records that navigation in history.
' Public role: Form-code/event-expression entry point. Generated or managed Access forms can invoke this Public routine from form code or an event expression.
' Parameters: None.
' Returns: Boolean: True when chat gpt help open related completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or not-applicable path.
Public Function ChatGPT_HelpOpenRelated() As Boolean
    ' Wrapper purpose: Generated Help Center callback that reads the related-target control and navigates to the referenced Help content.
    ' Why retained: Retained because the generated Related control needs a stable public Function entry point while target parsing and Help navigation remain private.
    On Error GoTo ErrHandler
    Dim targetText As String
    Dim parts As Variant
    ChatGPTHelpCaptureFormReference
    targetText = Trim$(CStr(Nz(mChatGPTHelpForm.Controls("txtHelpRelatedTarget").value, vbNullString)))
    If Len(targetText) = 0 Then
        CoreShowMessage "This Help item does not have a related action.", vbInformation, "Access Development Console Help", "ChatGPT_HelpOpenRelated", "UI"
        ChatGPT_HelpOpenRelated = True
        GoTo CleanExit
    End If
    parts = Split(targetText, ":")
    Select Case UCase$(CStr(parts(0)))
        Case "FORM"
            If ChatGPT_OpenFormByName(Mid$(targetText, 6)) = False Then GoTo CleanExit
        Case "COMMAND"
            If RunSystemCommand(Mid$(targetText, 9)) = False Then GoTo CleanExit
        Case "OBJECT"
            ChatGPTHelpOpenDocumentedObject parts
        Case "T", "O", "F"
            ChatGPTHelpLoadKey targetText, True
        Case Else
            Err.Raise vbObjectError + 6780, "ChatGPT_HelpOpenRelated", "Unsupported related target: " & targetText
    End Select
    ChatGPT_HelpOpenRelated = True
CleanExit:
    ChatGPTHelpReleaseFormReference
    Exit Function
ErrHandler:
    ChatGPTHelpShowRuntimeError "Open related Help target", Err.Number, Err.Source, Err.Description, "ChatGPT_HelpOpenRelated"
    Resume CleanExit
End Function

' =============================================================================================================================================================================
' HELP INSTANCE / HISTORY / NAVIGATION / RENDERING HELPERS
' =============================================================================================================================================================================

' Purpose: Resolves chat GPT help list text from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   valueText (Input; String): Input value text examined or transformed by this procedure according to its documented chat GPT help list text contract.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function ChatGPTHelpListText(ByVal valueText As String) As String
    'Converts semicolons and line breaks to list-safe display separators and returns trimmed Help-list text.
    valueText = Replace(valueText, ";", ",")
    valueText = Replace(valueText, vbCr, " ")
    valueText = Replace(valueText, vbLf, " ")
    ChatGPTHelpListText = Trim$(valueText)
End Function

Private Sub ChatGPTHelpCaptureFormReference()
    Dim candidate As Object
    Dim canonicalFormName As String
    Set mChatGPTHelpForm = Nothing
    On Error GoTo CodeContextUnavailable
    Set candidate = Application.CodeContextObject
    On Error GoTo 0
    If Not candidate Is Nothing Then
        If TypeOf candidate Is Access.Form Then
            If ChatGPTHelpFormImplementsContract(candidate) Then
                Set mChatGPTHelpForm = candidate
                Exit Sub
            End If
            Err.Raise vbObjectError + 6779, "ChatGPTHelpCaptureFormReference", "The form that invoked the Help event does not implement the Access Development Console Help control contract. Form: " & candidate.Name
        End If
    End If
CodeContextUnavailable:
    On Error GoTo 0
    Set candidate = Nothing
    On Error GoTo ActiveFormUnavailable
    Set candidate = Screen.ActiveForm
    On Error GoTo 0
    If Not candidate Is Nothing Then
        If TypeOf candidate Is Access.Form Then
            If ChatGPTHelpFormImplementsContract(candidate) Then
                Set mChatGPTHelpForm = candidate
                Exit Sub
            End If
            Err.Raise vbObjectError + 6779, "ChatGPTHelpCaptureFormReference", "The active form does not implement the Access Development Console Help control contract. Form: " & candidate.Name
        End If
    End If
ActiveFormUnavailable:
    On Error GoTo 0
    canonicalFormName = ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Help")
    Set candidate = Nothing
    On Error GoTo CanonicalUnavailable
    If CurrentProject.AllForms(canonicalFormName).IsLoaded Then Set candidate = forms(canonicalFormName)
    On Error GoTo 0
    If Not candidate Is Nothing Then
        If ChatGPTHelpFormImplementsContract(candidate) Then
            Set mChatGPTHelpForm = candidate
            Exit Sub
        End If
    End If
CanonicalUnavailable:
    On Error GoTo 0
    Err.Raise vbObjectError + 6779, "ChatGPTHelpCaptureFormReference", "No open form implementing the Access Development Console Help control contract could be resolved."
End Sub

Private Function ChatGPTHelpFormImplementsContract(ByVal candidateForm As Access.Form) As Boolean
    Dim requiredDefinitions As Collection
    Dim definitionData As Variant
    Dim actualControls As Object
    Dim ctl As control
    Dim canonicalFormName As String
    If candidateForm Is Nothing Then Exit Function
    canonicalFormName = ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Help")
    Set requiredDefinitions = GeneratedFormContractControlDefinitions(canonicalFormName)
    Set actualControls = CreateObject("Scripting.Dictionary")
    actualControls.CompareMode = vbTextCompare
    For Each ctl In candidateForm.Controls
        If Len(Trim$(ctl.Name)) > 0 Then actualControls(ctl.Name) = True
    Next ctl
    For Each definitionData In requiredDefinitions
        If actualControls.Exists(CStr(definitionData(0))) = False Then Exit Function
    Next definitionData
    ChatGPTHelpFormImplementsContract = True
End Function

Private Sub ChatGPTHelpReleaseFormReference()
    Set mChatGPTHelpForm = Nothing
End Sub

Private Sub ChatGPTHelpResetInstanceState()
    mChatGPTHelpForm.Controls("txtHelpHistoryState").value = vbNullString
    mChatGPTHelpForm.Controls("txtHelpHistoryIndex").value = 0
    mChatGPTHelpForm.Controls("txtHelpNavigatingHistory").value = False
End Sub

Private Function ChatGPTHelpHistoryItems() As Collection
    Dim items As New Collection
    Dim historyText As String
    Dim parts As Variant
    Dim item As Variant
    historyText = CStr(Nz(mChatGPTHelpForm.Controls("txtHelpHistoryState").value, vbNullString))
    If Len(historyText) > 0 Then
        parts = Split(historyText, vbLf)
        For Each item In parts
            If Len(CStr(item)) > 0 Then items.Add CStr(item)
        Next item
    End If
    Set ChatGPTHelpHistoryItems = items
End Function

Private Sub ChatGPTHelpWriteHistoryItems(ByVal items As Collection)
    Dim item As Variant
    Dim historyText As String
    For Each item In items
        If Len(historyText) > 0 Then historyText = historyText & vbLf
        historyText = historyText & CStr(item)
    Next item
    mChatGPTHelpForm.Controls("txtHelpHistoryState").value = historyText
End Sub

Private Function ChatGPTHelpHistoryIndexValue() As Long
    ChatGPTHelpHistoryIndexValue = CLng(Nz(mChatGPTHelpForm.Controls("txtHelpHistoryIndex").value, 0))
End Function

Private Sub ChatGPTHelpSetHistoryIndexValue(ByVal indexValue As Long)
    mChatGPTHelpForm.Controls("txtHelpHistoryIndex").value = indexValue
End Sub

Private Function ChatGPTHelpNavigatingHistoryValue() As Boolean
    ChatGPTHelpNavigatingHistoryValue = CBool(Nz(mChatGPTHelpForm.Controls("txtHelpNavigatingHistory").value, False))
End Function

Private Sub ChatGPTHelpSetNavigatingHistoryValue(ByVal navigating As Boolean)
    mChatGPTHelpForm.Controls("txtHelpNavigatingHistory").value = navigating
End Sub

Private Sub ChatGPTHelpLoadSections()
    Dim rowSourceText As String
    rowSourceText = ChatGPTHelpValueListPair(rowSourceText, "START", "Getting Started")
    rowSourceText = ChatGPTHelpValueListPair(rowSourceText, "SETUP", "Setup and Updates")
    rowSourceText = ChatGPTHelpValueListPair(rowSourceText, "CONFIG", "Configuration")
    rowSourceText = ChatGPTHelpValueListPair(rowSourceText, "SCAN", "Scanning Candidates")
    rowSourceText = ChatGPTHelpValueListPair(rowSourceText, "REVIEW", "Reviewing and Accepting Candidates")
    rowSourceText = ChatGPTHelpValueListPair(rowSourceText, "GENERATE", "Generating Macros")
    rowSourceText = ChatGPTHelpValueListPair(rowSourceText, "HEALTH", "Health and Validation")
    rowSourceText = ChatGPTHelpValueListPair(rowSourceText, "OBJECTS", "Managed Objects")
    rowSourceText = ChatGPTHelpValueListPair(rowSourceText, "FORMS", "Generated Forms")
    rowSourceText = ChatGPTHelpValueListPair(rowSourceText, "RIBBON", "Ribbon and Navigation Pane")
    rowSourceText = ChatGPTHelpValueListPair(rowSourceText, "GUARD", "Module Update Guard")
    rowSourceText = ChatGPTHelpValueListPair(rowSourceText, "DIAG", "Diagnostics and Troubleshooting")
    rowSourceText = ChatGPTHelpValueListPair(rowSourceText, "REFS", "Reference Management")
    rowSourceText = ChatGPTHelpValueListPair(rowSourceText, "ADDINS", "Add-Ins Review")
    rowSourceText = ChatGPTHelpValueListPair(rowSourceText, "FAQ", "Common Questions")
    rowSourceText = ChatGPTHelpValueListPair(rowSourceText, "DBREF", "Database Reference")
    ChatGPTHelpSetListRows "lstHelpSections", rowSourceText
End Sub

Private Sub ChatGPTHelpSelectSectionKey(ByVal sectionKey As String, ByVal selectFirst As Boolean)
    Dim rowSourceText As String
    Dim topics As Collection
    Dim topic As Variant
    Dim firstKey As String
    ChatGPTHelpSelectListKey "lstHelpSections", sectionKey
    ChatGPTHelpSetListRows "lstHelpFields", vbNullString
    If UCase$(sectionKey) = "DBREF" Then
        ChatGPTHelpLoadDatabaseObjects vbNullString
        If selectFirst Then ChatGPTHelpSelectFirstAndLoad "lstHelpTopics"
        Exit Sub
    End If
    Set topics = ChatGPTHelpTopics()
    For Each topic In topics
        SetupControlPump
        If StrComp(CStr(topic(0)), sectionKey, vbTextCompare) = 0 Then
            If Len(firstKey) = 0 Then firstKey = "T:" & CStr(topic(0)) & ":" & CStr(topic(1))
            rowSourceText = ChatGPTHelpValueListPair(rowSourceText, "T:" & CStr(topic(0)) & ":" & CStr(topic(1)), CStr(topic(2)))
        End If
    Next topic
    ChatGPTHelpSetListRows "lstHelpTopics", rowSourceText
    If selectFirst And Len(firstKey) > 0 Then
        ChatGPTHelpSelectListKey "lstHelpTopics", firstKey
        ChatGPTHelpLoadKey firstKey, True
    End If
End Sub

Private Sub ChatGPTHelpLoadKey(ByVal itemKey As String, ByVal addHistory As Boolean)
    Dim parts As Variant
    If Len(itemKey) = 0 Then Exit Sub
    parts = Split(itemKey, ":")
    Select Case UCase$(CStr(parts(0)))
        Case "T"
            If UBound(parts) < 2 Then Exit Sub
            ChatGPTHelpShowFixedTopic CStr(parts(1)), CStr(parts(2))
        Case "O"
            If UBound(parts) < 1 Then Exit Sub
            ChatGPTHelpShowManagedObject CLng(parts(1))
        Case "F"
            If UBound(parts) < 1 Then Exit Sub
            ChatGPTHelpShowManagedObjectField CLng(parts(1))
        Case Else
            Exit Sub
    End Select
    mChatGPTHelpForm.Controls("txtHelpCurrentKey").value = itemKey
    If addHistory And ChatGPTHelpNavigatingHistoryValue() = False Then ChatGPTHelpAddHistory itemKey
End Sub

Private Sub ChatGPTHelpShowFixedTopic(ByVal sectionKey As String, ByVal topicKey As String)
    Dim topics As Collection
    Dim topic As Variant
    Set topics = ChatGPTHelpTopics()
    For Each topic In topics
        SetupControlPump
        If StrComp(CStr(topic(0)), sectionKey, vbTextCompare) = 0 And StrComp(CStr(topic(1)), topicKey, vbTextCompare) = 0 Then
            ChatGPTHelpDisplay CStr(topic(2)), ChatGPTHelpSectionCaption(sectionKey), CStr(topic(3)), CStr(topic(4))
            ChatGPTHelpSelectListKey "lstHelpSections", sectionKey
            ChatGPTHelpSelectListKey "lstHelpTopics", "T:" & sectionKey & ":" & topicKey
            ChatGPTHelpSetListRows "lstHelpFields", vbNullString
            Exit Sub
        End If
    Next topic
End Sub

Private Sub ChatGPTHelpLoadDatabaseObjects(ByVal searchText As String)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim criteriaText As String
    Dim rowSourceText As String
    Dim managedObjectsTableName As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Set db = GetHostDb()
    managedObjectsTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID)
    If SharedDao_TableExists(db, managedObjectsTableName) = False Then
        ChatGPTHelpSetListRows "lstHelpTopics", vbNullString
        ChatGPTHelpDisplay "Database Reference Is Not Ready", "Database Reference", managedObjectsTableName & " is missing. Run Full ADC Setup to create and synchronize the documentation tables.", "COMMAND:RunFullSetup"
        Exit Sub
    End If
    criteriaText = "Nz(" & SQL_FIELD_IS_DOCUMENTATION_ACTIVE & ",True)=True AND UCase(Trim(Nz(" & SQL_FIELD_MANAGED_OBJECT_STATUS & ",'Active')))='ACTIVE'"
    If Len(searchText) > 0 Then criteriaText = criteriaText & " AND (Nz(" & SQL_FIELD_OBJECT_NAME & ",'') Like " & SharedDao_SqlQuote("*" & searchText & "*") & " OR Nz(" & SQL_FIELD_OBJECT_DISPLAY_NAME & ",'') Like " & SharedDao_SqlQuote("*" & searchText & "*") & " OR Nz(" & SQL_FIELD_OBJECT_DESCRIPTION & ",'') Like " & SharedDao_SqlQuote("*" & searchText & "*") & " OR Nz(" & SQL_FIELD_MANAGED_OBJECT_FIELD_SEARCH_KEYWORDS & ",'') Like " & SharedDao_SqlQuote("*" & searchText & "*") & " OR Nz(" & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_DOCUMENTATION_STATUS & ",'') Like " & SharedDao_SqlQuote("*" & searchText & "*") & ")"
    sqlText = "SELECT " & SQL_FIELD_ID & "," & SQL_FIELD_OBJECT_NAME & "," & SQL_FIELD_OBJECT_TYPE & "," & SQL_FIELD_OBJECT_DISPLAY_NAME & "," & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_DOCUMENTATION_STATUS & " FROM [" & managedObjectsTableName & "] WHERE " & criteriaText & " ORDER BY " & SQL_FIELD_OBJECT_TYPE & "," & SQL_FIELD_OBJECT_NAME & ";"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    Do While Not rs.EOF
        SetupControlPump
        rowSourceText = ChatGPTHelpValueListPair(rowSourceText, "O:" & CStr(rs.fields(FIELD_ID).value), CStr(Nz(rs.fields(FIELD_OBJECT_TYPE).value, vbNullString)) & " - " & CStr(Nz(rs.fields(FIELD_OBJECT_NAME).value, vbNullString)) & ChatGPTHelpStatusSuffix(CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_DEPENDENCY_DOCUMENTATION_STATUS).value, vbNullString))))
        rs.MoveNext
    Loop
    rs.Close
    ChatGPTHelpSetListRows "lstHelpTopics", rowSourceText
    If Len(rowSourceText) = 0 Then ChatGPTHelpDisplay "No Matching Objects", "Database Reference", "No active Access Development Console-owned object documentation matched the current filter.", vbNullString
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
    Err.Raise savedErrNumber, "ChatGPTHelpLoadDatabaseObjects", "Could not load the table-backed object reference." & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Sub

Private Function ChatGPTHelpObjectDependenciesText(ByVal db As DAO.Database, ByVal managedObjectID As Long) As String
    On Error GoTo ExitHere
    Dim rs As DAO.Recordset
    Dim resultText As String
    Dim managedObjectsTableName As String
    managedObjectsTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID)
    If SharedDao_TableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectDependencies")) = False Then Exit Function
    Set rs = db.OpenRecordset("SELECT d." & SQL_FIELD_DEPENDENCY_NAME & ",d." & SQL_FIELD_DEPENDENCY_OBJECT_TYPE & ",d." & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_PURPOSE & ",d." & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_DETAIL_DESCRIPTION & ",d." & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_IS_REQUIRED & ",d." & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_FAILURE_IMPACT & ",o." & SQL_FIELD_OBJECT_NAME & " AS [DependsOnName] FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectDependencies") & "] AS d LEFT JOIN [" & managedObjectsTableName & "] AS o ON d." & SQL_FIELD_DEPENDS_ON_MANAGED_OBJECT_ID & "=o." & SQL_FIELD_ID & " WHERE d." & SQL_FIELD_MANAGED_OBJECT_ID & "=" & CStr(managedObjectID) & " AND UCase(Trim(Nz(d." & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_STATUS & ",'ACTIVE')))='ACTIVE' ORDER BY d." & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER & ",d." & SQL_FIELD_DEPENDENCY_NAME & ";", dbOpenSnapshot)
    Do While Not rs.EOF
        SetupControlPump
        resultText = resultText & IIf(Len(resultText) = 0, vbCrLf & "Depends on" & vbCrLf, vbNullString) & "- " & CStr(Nz(rs.fields("DependsOnName").value, Nz(rs.fields(FIELD_DEPENDENCY_NAME).value, vbNullString)))
        If Len(CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_DEPENDENCY_PURPOSE).value, vbNullString))) > 0 Then resultText = resultText & ": " & CStr(rs.fields(FIELD_MANAGED_OBJECT_DEPENDENCY_PURPOSE).value)
        If Nz(rs.fields(FIELD_MANAGED_OBJECT_DEPENDENCY_IS_REQUIRED).value, False) Then resultText = resultText & " [required]"
        resultText = resultText & vbCrLf
        rs.MoveNext
    Loop
    rs.Close
    Set rs = db.OpenRecordset("SELECT o." & SQL_FIELD_OBJECT_NAME & ",d." & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_PURPOSE & ",d." & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_IS_REQUIRED & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectDependencies") & "] AS d INNER JOIN [" & managedObjectsTableName & "] AS o ON d." & SQL_FIELD_MANAGED_OBJECT_ID & "=o." & SQL_FIELD_ID & " WHERE d." & SQL_FIELD_DEPENDS_ON_MANAGED_OBJECT_ID & "=" & CStr(managedObjectID) & " AND UCase(Trim(Nz(d." & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_STATUS & ",'ACTIVE')))='ACTIVE' ORDER BY o." & SQL_FIELD_OBJECT_NAME & ";", dbOpenSnapshot)
    Do While Not rs.EOF
        SetupControlPump
        resultText = resultText & IIf(InStr(1, resultText, vbCrLf & "Used by" & vbCrLf, vbBinaryCompare) = 0, vbCrLf & "Used by" & vbCrLf, vbNullString) & "- " & CStr(Nz(rs.fields(FIELD_OBJECT_NAME).value, vbNullString))
        If Len(CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_DEPENDENCY_PURPOSE).value, vbNullString))) > 0 Then resultText = resultText & ": " & CStr(rs.fields(FIELD_MANAGED_OBJECT_DEPENDENCY_PURPOSE).value)
        If Nz(rs.fields(FIELD_MANAGED_OBJECT_DEPENDENCY_IS_REQUIRED).value, False) Then resultText = resultText & " [required]"
        resultText = resultText & vbCrLf
        rs.MoveNext
    Loop
    rs.Close
    ChatGPTHelpObjectDependenciesText = resultText
ExitHere:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Err.Clear
End Function

Private Sub ChatGPTHelpLoadSearchResults(ByVal searchText As String)
    On Error GoTo ErrHandler
    Dim topics As Collection
    Dim topic As Variant
    Dim rowSourceText As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim likeText As String
    Dim managedObjectsTableName As String
    Dim resultCount As Long
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Set topics = ChatGPTHelpTopics()
    For Each topic In topics
        SetupControlPump
        If InStr(1, CStr(topic(2)) & " " & CStr(topic(3)) & " " & ChatGPTHelpSectionCaption(CStr(topic(0))), searchText, vbTextCompare) > 0 Then
            rowSourceText = ChatGPTHelpValueListPair(rowSourceText, "T:" & CStr(topic(0)) & ":" & CStr(topic(1)), "Help - " & CStr(topic(2)))
            resultCount = resultCount + 1
        End If
    Next topic
    Set db = GetHostDb()
    managedObjectsTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID)
    likeText = "*" & searchText & "*"
    If SharedDao_TableExists(db, managedObjectsTableName) Then
        sqlText = "SELECT " & SQL_FIELD_ID & "," & SQL_FIELD_OBJECT_NAME & "," & SQL_FIELD_OBJECT_TYPE & "," & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_DOCUMENTATION_STATUS & " FROM [" & managedObjectsTableName & "] WHERE Nz(" & SQL_FIELD_IS_DOCUMENTATION_ACTIVE & ",True)=True AND (Nz(" & SQL_FIELD_OBJECT_NAME & ",'') Like " & SharedDao_SqlQuote(likeText) & " OR Nz(" & SQL_FIELD_OBJECT_DISPLAY_NAME & ",'') Like " & SharedDao_SqlQuote(likeText) & " OR Nz(" & SQL_FIELD_OBJECT_DESCRIPTION & ",'') Like " & SharedDao_SqlQuote(likeText) & " OR Nz(" & SQL_FIELD_MANAGED_OBJECT_FIELD_SEARCH_KEYWORDS & ",'') Like " & SharedDao_SqlQuote(likeText) & " OR Nz(" & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_DOCUMENTATION_STATUS & ",'') Like " & SharedDao_SqlQuote(likeText) & ") ORDER BY " & SQL_FIELD_OBJECT_TYPE & "," & SQL_FIELD_OBJECT_NAME & ";"
        Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
        Do While Not rs.EOF And resultCount < HELP_SEARCH_MAX_RESULTS
            SetupControlPump
            rowSourceText = ChatGPTHelpValueListPair(rowSourceText, "O:" & CStr(rs.fields(FIELD_ID).value), "Object - " & CStr(Nz(rs.fields(FIELD_OBJECT_TYPE).value, vbNullString)) & " - " & CStr(Nz(rs.fields(FIELD_OBJECT_NAME).value, vbNullString)) & ChatGPTHelpStatusSuffix(CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_DEPENDENCY_DOCUMENTATION_STATUS).value, vbNullString))))
            resultCount = resultCount + 1
            rs.MoveNext
        Loop
        rs.Close
    End If
    If SharedDao_TableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFields")) Then
        sqlText = "SELECT f." & SQL_FIELD_ID & ",f." & SQL_FIELD_MANAGED_OBJECT_FIELD_NAME & ",f." & SQL_FIELD_MANAGED_OBJECT_FIELD_DATA_TYPE & ",f." & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_DOCUMENTATION_STATUS & ",o." & SQL_FIELD_OBJECT_NAME & _
            " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFields") & "] AS f INNER JOIN [" & managedObjectsTableName & "] AS o ON f." & SQL_FIELD_MANAGED_OBJECT_ID & "=o." & SQL_FIELD_ID & _
            " WHERE f." & SQL_FIELD_MANAGED_OBJECT_FIELD_IS_ACTIVE & "=True AND (Nz(f." & SQL_FIELD_MANAGED_OBJECT_FIELD_NAME & ",'') Like " & SharedDao_SqlQuote(likeText) & _
            " OR Nz(f." & SQL_FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME & ",'') Like " & SharedDao_SqlQuote(likeText) & _
            " OR Nz(f." & SQL_FIELD_MANAGED_OBJECT_FIELD_DESCRIPTION_LONG & ",'') Like " & SharedDao_SqlQuote(likeText) & _
            " OR Nz(f." & SQL_FIELD_MANAGED_OBJECT_FIELD_USAGE_GUIDANCE & ",'') Like " & SharedDao_SqlQuote(likeText) & _
            " OR Nz(f." & SQL_FIELD_MANAGED_OBJECT_FIELD_SEARCH_KEYWORDS & ",'') Like " & SharedDao_SqlQuote(likeText) & _
            " OR Nz(f." & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_DOCUMENTATION_STATUS & ",'') Like " & SharedDao_SqlQuote(likeText) & _
            ") ORDER BY o." & SQL_FIELD_OBJECT_NAME & ",f." & SQL_FIELD_MANAGED_OBJECT_FIELD_ORDINAL & ",f." & SQL_FIELD_MANAGED_OBJECT_FIELD_NAME & ";"
        Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
        Do While Not rs.EOF And resultCount < HELP_SEARCH_MAX_RESULTS
            SetupControlPump
            rowSourceText = ChatGPTHelpValueListPair(rowSourceText, "F:" & CStr(rs.fields(FIELD_ID).value), "Field - " & CStr(Nz(rs.fields(FIELD_OBJECT_NAME).value, vbNullString)) & "." & CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_NAME).value, vbNullString)) & ChatGPTHelpStatusSuffix(CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_DEPENDENCY_DOCUMENTATION_STATUS).value, vbNullString))))
            resultCount = resultCount + 1
            rs.MoveNext
        Loop
        rs.Close
    End If
    ChatGPTHelpSetListRows "lstHelpFields", vbNullString
    ChatGPTHelpSetListRows "lstHelpTopics", rowSourceText
    ChatGPTHelpDisplay "Search Results", "Search: " & searchText, IIf(Len(rowSourceText) > 0, "Select a result to open the matching Help topic, object, or field.", "No Help or database-reference rows matched the search."), vbNullString
    If Len(rowSourceText) > 0 Then ChatGPTHelpSelectFirstAndLoad "lstHelpTopics"
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
    Err.Raise savedErrNumber, "ChatGPTHelpLoadSearchResults", "Could not search fixed Help and table-backed database-reference rows." & vbCrLf & "Search: " & searchText & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Sub

Private Sub ChatGPTHelpDisplay(ByVal TitleText As String, ByVal breadcrumbText As String, ByVal contentText As String, ByVal relatedTarget As String)
    mChatGPTHelpForm.Controls("txtHelpTitle").value = TitleText
    mChatGPTHelpForm.Controls("txtHelpBreadcrumb").value = breadcrumbText
    mChatGPTHelpForm.Controls("txtHelpContent").value = contentText
    mChatGPTHelpForm.Controls("txtHelpRelatedTarget").value = relatedTarget
End Sub

Private Sub ChatGPTHelpAddHistory(ByVal itemKey As String)
    Dim historyItems As Collection
    Dim revised As Collection
    Dim historyIndex As Long
    Dim index As Long
    Set historyItems = ChatGPTHelpHistoryItems()
    historyIndex = ChatGPTHelpHistoryIndexValue()
    If historyIndex > 0 And historyIndex <= historyItems.count Then
        If StrComp(CStr(historyItems(historyIndex)), itemKey, vbTextCompare) = 0 Then Exit Sub
    End If
    If historyIndex < historyItems.count Then
        Set revised = New Collection
        For index = 1 To historyIndex
            SetupControlPump
            revised.Add historyItems(index)
        Next index
        Set historyItems = revised
    End If
    historyItems.Add itemKey
    ChatGPTHelpWriteHistoryItems historyItems
    ChatGPTHelpSetHistoryIndexValue historyItems.count
End Sub

Private Sub ChatGPTHelpMoveWithinCurrentList(ByVal offsetValue As Long)
    Dim ctl As control
    Dim targetIndex As Long
    Dim itemKey As String
    Set ctl = mChatGPTHelpForm.Controls("lstHelpFields")
    If ctl.listIndex >= 0 Then
        targetIndex = ctl.listIndex + offsetValue
        If targetIndex >= 0 And targetIndex < ctl.ListCount Then
            ctl.value = ctl.Column(0, targetIndex)
            itemKey = CStr(ctl.Column(0, targetIndex))
            ChatGPTHelpLoadKey itemKey, True
            Exit Sub
        End If
    End If
    Set ctl = mChatGPTHelpForm.Controls("lstHelpTopics")
    If ctl.listIndex < 0 Then Exit Sub
    targetIndex = ctl.listIndex + offsetValue
    If targetIndex < 0 Or targetIndex >= ctl.ListCount Then Exit Sub
    ctl.value = ctl.Column(0, targetIndex)
    itemKey = CStr(ctl.Column(0, targetIndex))
    ChatGPTHelpLoadKey itemKey, True
End Sub

Private Sub ChatGPTHelpSelectFirstAndLoad(ByVal controlName As String)
    Dim ctl As control
    Set ctl = mChatGPTHelpForm.Controls(controlName)
    If ctl.ListCount = 0 Then Exit Sub
    ctl.value = ctl.Column(0, 0)
    ChatGPTHelpLoadKey CStr(ctl.Column(0, 0)), True
End Sub

Private Function ChatGPTHelpSelectedListKey(ByVal controlName As String) As String
    Dim ctl As control
    Set ctl = mChatGPTHelpForm.Controls(controlName)
    If ctl.listIndex < 0 Then Exit Function
    ChatGPTHelpSelectedListKey = CStr(ctl.Column(0, ctl.listIndex))
End Function

Private Sub ChatGPTHelpSelectListKey(ByVal controlName As String, ByVal itemKey As String)
    Dim ctl As control
    Dim index As Long
    Set ctl = mChatGPTHelpForm.Controls(controlName)
    For index = 0 To ctl.ListCount - 1
        SetupControlPump
        If StrComp(CStr(ctl.Column(0, index)), itemKey, vbTextCompare) = 0 Then
            ctl.value = ctl.Column(0, index)
            Exit Sub
        End If
    Next index
End Sub

Private Sub ChatGPTHelpSetListRows(ByVal controlName As String, ByVal rowSourceText As String)
    With mChatGPTHelpForm.Controls(controlName)
        .rowSourceType = "Value List"
        .rowSource = rowSourceText
        .Requery
    End With
End Sub

Private Function ChatGPTHelpValueListPair(ByVal currentText As String, ByVal itemKey As String, ByVal displayText As String) As String
    Dim pairText As String
    displayText = Left$(ChatGPTHelpListText(displayText), 240)
    pairText = ChatGPTHelpListText(itemKey) & ";" & displayText
    If Len(currentText) > 0 Then pairText = ";" & pairText
    If Len(currentText) + Len(pairText) > HELP_VALUE_LIST_MAX_LENGTH Then
        ChatGPTHelpValueListPair = currentText
    Else
        ChatGPTHelpValueListPair = currentText & pairText
    End If
End Function


Private Function ChatGPTHelpSectionCaption(ByVal sectionKey As String) As String
    Select Case UCase$(sectionKey)
        Case "START": ChatGPTHelpSectionCaption = "Getting Started"
        Case "SETUP": ChatGPTHelpSectionCaption = "Setup and Updates"
        Case "CONFIG": ChatGPTHelpSectionCaption = "Configuration"
        Case "SCAN": ChatGPTHelpSectionCaption = "Scanning Candidates"
        Case "REVIEW": ChatGPTHelpSectionCaption = "Reviewing and Accepting Candidates"
        Case "GENERATE": ChatGPTHelpSectionCaption = "Generating Macros"
        Case "HEALTH": ChatGPTHelpSectionCaption = "Health and Validation"
        Case "OBJECTS": ChatGPTHelpSectionCaption = "Managed Objects"
        Case "FORMS": ChatGPTHelpSectionCaption = "Generated Forms"
        Case "RIBBON": ChatGPTHelpSectionCaption = "Ribbon and Navigation Pane"
        Case "GUARD": ChatGPTHelpSectionCaption = "Module Update Guard"
        Case "DIAG": ChatGPTHelpSectionCaption = "Diagnostics and Troubleshooting"
        Case "REFS": ChatGPTHelpSectionCaption = "Reference Management"
        Case "ADDINS": ChatGPTHelpSectionCaption = "Add-Ins Review"
        Case "FAQ": ChatGPTHelpSectionCaption = "Common Questions"
        Case "DBREF": ChatGPTHelpSectionCaption = "Database Reference"
        Case Else: ChatGPTHelpSectionCaption = sectionKey
    End Select
End Function

Private Function ChatGPTHelpLabeledText(ByVal labelText As String, ByVal valueText As String) As String
    If Len(Trim$(valueText)) = 0 Then Exit Function
    ChatGPTHelpLabeledText = labelText & vbCrLf & valueText & vbCrLf & vbCrLf
End Function

Private Function ChatGPTHelpStatusSuffix(ByVal statusText As String) As String
    If Len(Trim$(statusText)) = 0 Or StrComp(statusText, "Complete", vbTextCompare) = 0 Then Exit Function
    ChatGPTHelpStatusSuffix = " [" & statusText & "]"
End Function

Private Function ChatGPTHelpDateText(ByVal value As Variant) As String
    If IsNull(value) Or IsDate(value) = False Then Exit Function
    ChatGPTHelpDateText = Format$(CDate(value), "yyyy-mm-dd hh:nn:ss")
End Function

Private Function ChatGPTHelpObjectTarget(ByVal objectType As String, ByVal objectName As String) As String
    Select Case UCase$(Trim$(objectType))
        Case "TABLE", "QUERY", "FORM", "REPORT", "MACRO": ChatGPTHelpObjectTarget = "OBJECT:" & UCase$(Trim$(objectType)) & ":" & objectName
    End Select
End Function

Private Sub ChatGPTHelpOpenDocumentedObject(ByVal parts As Variant)
    Dim objectType As String
    Dim objectName As String
    If UBound(parts) < 2 Then Exit Sub
    objectType = UCase$(CStr(parts(1)))
    objectName = CStr(parts(2))
    Select Case objectType
        Case "TABLE": DoCmd.OpenTable objectName, acViewNormal, acReadOnly
        Case "QUERY"
            If StrComp(objectName, ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Terminology_Review"), vbTextCompare) = 0 Then
                DoCmd.OpenQuery objectName, acViewNormal, acEdit
            Else
                DoCmd.OpenQuery objectName, acViewNormal, acReadOnly
            End If
        Case "FORM": DoCmd.OpenForm objectName, acNormal
        Case "REPORT": DoCmd.OpenReport objectName, acViewReport
        Case "MACRO": DoCmd.RunMacro objectName
    End Select
End Sub

Private Sub ChatGPTHelpShowRuntimeError(ByVal operationText As String, ByVal errorNumber As Long, ByVal errorSource As String, ByVal errorDescription As String, ByVal sourceProcedure As String)
    Dim helpFormName As String
    helpFormName = ManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Help")
    If Not mChatGPTHelpForm Is Nothing Then helpFormName = mChatGPTHelpForm.Name
    LogTraceMessage "Error", "Help", sourceProcedure, "Generated Help operation failed.", "Operation: " & operationText & vbCrLf & "Err.Number: " & CStr(errorNumber) & vbCrLf & "Err.Source: " & errorSource & vbCrLf & "Err.Description: " & errorDescription, operationText, errorNumber, errorSource, errorDescription, helpFormName, "Form", "CodeDb", "Help"
    CoreShowErrorMessage operationText & " could not complete." & vbCrLf & vbCrLf & "Err.Number: " & CStr(errorNumber) & vbCrLf & "Err.Source: " & errorSource & vbCrLf & "Err.Description: " & errorDescription, vbCritical, "Access Development Console Help", sourceProcedure, "UI"
End Sub
