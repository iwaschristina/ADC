Attribute VB_Name = "x__UTIL_ADC_SETUP"
' ================================================================================================================================================================================
' ChatGPT Rewrite Tracking Header
' ================================================================================================================================================================================
' ChatGPT Previous Version ID: CGPT-UQSX-MMPG-S6SU-PL98-24W6
' ChatGPT Baseline Version ID: CGPT-UQSX-MMPG-S6SU-PL98-24W6
' ChatGPT Version ID: CGPT-4KNW-Y3NJ-3X3R-X8W6-DXPQ
' ChatGPT Version Timestamp: Thursday 9/3/26 7:57 PM
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
' STANDARD MODULE: Access Development Console Setup
' VB_Name: x__UTIL_ADC_SETUP
' Entry Points:
'     RunSetupThroughPhaseJ
' SHARED CODE: SharedText_AppendNonBlankLine(ByVal existingText As String, ByVal lineText As String) -> String
' CLASSES USED: None
' DYNAMIC STRING REFERENCES: None
' LastScanTimestamp: Saturday 8/15/26 8:47 AM
' Purpose:
'     Owns the resequenced Access Development Console Setup execution for A1 through J1 after the unnumbered Setup Preflight completes, including standalone Part execution with read-only predecessor verification,
'     same-session phase plans, durable A-C / D-G / H / I rollback-group boundaries, and J1 final verification/persistence.
' Author: Christina Shonts
' Support: Access Development Console CodeDb Setup workflows.
' Code Category: Access Development Console Setup orchestration.
' Requirements:
'     Microsoft Access VBA, DAO, x__UTIL_ADC_Core, x__UTIL_ADC_MOC, x__UTIL_ADC_TABLES, x__UTIL_ADC_DB_INVENTORY_VBA, x__UTIL_ADC_REPORTS,
' Reference Dependencies: Access object model and DAO through the companion ADC modules.
' DB Object Dependencies:
'     Uses the Tables, Queries, and DB Inventory state established by A1-A4, the Phase-B Managed Object Catalog plus System Macro and Help Document Macro catalog state, the physical Access macro domain
'     verified/materialized by C1-C3, the D1-D4 managed-form and complete logical catalog state, E1-E3 physical managed-report state, F1 managed-object dependency state, G1 exact managed-description state,
'     H1-H3 committed ADC Ribbon state, I1-I3 complete physical Navigation Pane state, and Module Guard state verified by J1.
' Usage:
'     Called by RunFullSetup after unnumbered Preflight has resolved Module Guard, recovery state, Setup Control and rollback/checkpoint infrastructure.
' Version:
'     Adds new Setup Phase F/F1 managed-object dependency reconciliation and shifts existing Descriptions, Ribbon, Navigation Pane, and Final Verification to G, H, I, and J.
'     Removes stale module-header documentation for the deleted I1 same-run completed-part proof cache; executable behavior and Setup-phase ownership are unchanged.
'     Restores fresh I1 lookup-editor integration verification after responsiveness/event boundaries instead of carrying D4 physical binding state into the second verifier.
'     Routes Managed Object Catalog references to x__UTIL_ADC_MOC and physical lifecycle references to x__UTIL_ADC_MOC_LIFECYCLE after the responsibility split; runtime and Setup-phase behavior are unchanged.
'     Passes A2 schema-repair preflight evidence as one typed record so the known/value/summary state remains atomic across the phase boundary.
'     Documented the intentional pre-D bootstrap exceptions for Setup Control and Report Error forms while preserving D1-D4 as the ordinary generated-form lifecycle.
'     Completed the A-I phase model by inserting E1-E3 Reports between Forms and Descriptions, shifting Setup System Macro numbering, and integrating report proof into prerequisites, rollback, and I1.
'     Managed-header dependency metadata updated because x__UTIL_ADC_DEFAULTS is now a standard module; executable behavior unchanged.
'     I1 performs fresh read-only predecessor verification of A-H current state and does not rely on a same-run completed-part cache.
' Notes:
'     J1 is the final numbered Setup Part. It proves committed A-I outcomes and Module Guard state read-only, stages only J1 final completion evidence, and never repairs a failed predecessor domain.
' Core Rules Enforced Here:
'     One owner per setup domain; standalone Parts verify predecessors read-only and never silently execute earlier Parts. D1-D4 own the ordinary generated-form lifecycle; D2 and D3 own their disjoint generated-form
'     partitions and D4 owns complete managed-form reconciliation and logical catalog finalization.
'     The Setup Control forms and Report Error dialog are explicit pre-D bootstrap exceptions because Setup must support pause/cancel/progress and terminal error reporting before D can run. E1-E3 are the sole
'     physical managed-report Setup authority; F1 solely owns managed-object dependency discovery/reconciliation; G1 is the sole Setup description write authority. H1-H3 own Ribbon plan/stage/commit without
'     physical Navigation Pane placement; I1-I3 are the sole Setup authority for physical Navigation Pane group creation and complete-domain placement; J1 proves/finalizes without repair. Full Setup commits after
'     C3, G1, H3, and I3 according to durable rollback boundaries.
' Design Map:
'     A1-A4 Tables, Queries, and DB Inventory; B1-B5 Managed Object Catalog; C1-C3 System Macro and Help Document Macro; D1-D4 Forms and complete logical Managed Object Catalog; E1-E3 Reports; F1 Managed Object
'     Dependencies; G1 Descriptions; H1-H3 Ribbon; I1-I3 Navigation Pane; J1 Final Setup Verification.
' Configuration:
'     Existing ADC physical-name settings are authoritative; invalid or conflicting names fail A1 without automatic historical-name fallback or aliasing.
' Behavior:
'     Executes the 28 canonical Parts in order. D2 and D3 consume their disjoint partitions from the same-session D1 generated-form plan; D4 consumes exact same-run D2+D3 proof when available or authoritative
'     full-current-state generated-form proof otherwise; E2/E3 consume same-session report evidence; H2/H3 consume same-session Ribbon state; I3 consumes the same-session I2 complete placement plan.
'     J1 freshly revalidates predecessor-owned current state read-only, including reports, dependencies, descriptions, Ribbon, Navigation Pane, and lookup-editor physical integration, before staging only J1-owned
'     completion evidence.
' Code Status: Active A1-J1 Setup orchestrator; J1 is the final numbered Setup Part.
' Todo: None.
' Reporting Bugs: None currently recorded in this header.
' Module Level Bugs: None currently recorded in this header.
' Public Subs:
'     RunSetupThroughPhaseJ(Optional ByVal showMessage As Boolean = False, Optional ByVal moduleGuardVerificationDetails As String = vbNullString) -> None
' Public Functions:
'     RunSetupPartByIdentity(ByVal partIdentity As String, Optional ByVal showMessage As Boolean = True) -> Boolean
'     RunSetupPhaseByIdentity(ByVal phaseIdentity As String, Optional ByVal showMessage As Boolean = True) -> Boolean
' Public Properties: None
' ================================================================================================================================================================================

Option Compare Database
Option Explicit

' ===============================================================================================================================================================================
' MODULE-LEVEL DECLARATIONS
' ===============================================================================================================================================================================

' D2/D3 generated-form proof cache.
Private mSetupD2ProofTraceRunID As String
Private mSetupD2ProofToken As String
Private mSetupD3ProofTraceRunID As String
Private mSetupD3ProofToken As String
' Operation-scoped sequencing proof: set only after a Part passes its exit contract, then consumed by the next Part entry contract so broad shared helpers cannot advance Setup out of canonical order.
Private mSetupLastContractCompletedPartIdentity As String

' ===============================================================================================================================================================================
' PROCEDURES
' ===============================================================================================================================================================================

' Purpose: Runs Setup through Phase J using the supplied Module Guard verification details.
' Public role: Public module entry point. This procedure is listed in the module header as an intentional externally callable ADC entry point and may be invoked directly by Access/VBA callers.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: False.
'   moduleGuardVerificationDetails (Input Optional; String): Receives or supplies module guard verification details used for caller diagnostics, verification output, or user-facing reporting. Default:
'   vbNullString.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub RunSetupThroughPhaseJ(Optional ByVal showMessage As Boolean = False, Optional ByVal moduleGuardVerificationDetails As String = vbNullString)
    On Error GoTo ErrHandler
    Dim timingText As String
    Dim catalogContextStarted As Boolean
    CoreBeginSetupRunThroughPhaseJ moduleGuardVerificationDetails
    ExecuteSetupThroughPhase "J", timingText, catalogContextStarted
    CoreCompleteSetupRunThroughPhaseJ timingText, showMessage
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If catalogContextStarted Then
        On Error Resume Next
        EndCatalogSetupContext
        On Error GoTo 0
    End If
    CoreHandleSetupRunThroughPhaseJFailure savedErrNumber, savedErrSource, savedErrDescription
End Sub

Private Sub VerifySetupJ1FinalState(ByRef finalProofToken As String, ByRef moduleGuardStateToken As String, ByRef detailText As String)
    On Error GoTo A1NotCurrent
    Dim checkDetail As String
    Dim ribbonPredecessorToken As String
    Dim catalogToken As String
    Dim editorBindingToken As String
    Dim reportToken As String
    Dim dependencyToken As String
    Dim ribbonToken As String
    Dim navigationPaneToken As String
    Dim payloadText As String
    finalProofToken = vbNullString
    moduleGuardStateToken = vbNullString
    detailText = vbNullString
    checkDetail = vbNullString
    ValidateConfiguredPhysicalNamesForSetup "VerifySetupJ1FinalState.A1", False
    On Error GoTo 0
    detailText = SharedText_AppendNonBlankLine(detailText, "A1: current configured physical names proven.")
    checkDetail = vbNullString
    If CoreDataTableFoundationCurrentForSetup(checkDetail) = False Then RaiseSetupJ1FinalStateFailure "A2", checkDetail
    detailText = SharedText_AppendNonBlankLine(detailText, "A2: current core-table foundation proven.")
    checkDetail = vbNullString
    If NonFormQueryFoundationCurrentForSetup(checkDetail) = False Then RaiseSetupJ1FinalStateFailure "A3", checkDetail
    If FormDesignQueryFoundationCurrentForSetup(checkDetail) = False Then RaiseSetupJ1FinalStateFailure "A3", checkDetail
    detailText = SharedText_AppendNonBlankLine(detailText, "A3: current non-form and FormDesign query foundations proven.")
    checkDetail = vbNullString
    If ReferenceInventoryFoundationCurrentForSetup(checkDetail) = False Then RaiseSetupJ1FinalStateFailure "A4", checkDetail
    detailText = SharedText_AppendNonBlankLine(detailText, "A4: current reference/Add-In inventory proven.")
    checkDetail = vbNullString
    If CatalogPrerequisiteCurrentForSetup("C3", checkDetail) = False Then RaiseSetupJ1FinalStateFailure "B5", checkDetail
    detailText = SharedText_AppendNonBlankLine(detailText, "B1-B5: current authoritative catalog state proven.")
    checkDetail = vbNullString
    If PhaseC3CurrentForSetup(checkDetail) = False Then RaiseSetupJ1FinalStateFailure "C3", checkDetail
    detailText = SharedText_AppendNonBlankLine(detailText, "C3: current physical System Macro and Help Document Macro state proven.")
    checkDetail = vbNullString
    If GeneratedFormsCurrentForSetup(checkDetail) = False Then RaiseSetupJ1FinalStateFailure "D2/D3", checkDetail
    detailText = SharedText_AppendNonBlankLine(detailText, "D2/D3: both generated-form partitions proven current; together they cover the D1 authoritative plan.")
    checkDetail = vbNullString
    If ManagedFormDomainCurrentForSetup(checkDetail) = False Then RaiseSetupJ1FinalStateFailure "D4", checkDetail
    detailText = SharedText_AppendNonBlankLine(detailText, "D4: managed-form domain proof independently proven current.")
    checkDetail = vbNullString
    catalogToken = CompleteCatalogStateTokenForSetup(checkDetail)
    If Len(catalogToken) = 0 Then RaiseSetupJ1FinalStateFailure "complete catalog derived state", checkDetail
    detailText = SharedText_AppendNonBlankLine(detailText, "Complete catalog: shared logical state is current after all owner-phase mutations through G1.")
    checkDetail = vbNullString
    If D4LookupEditorBindingCurrentForSetup(checkDetail, editorBindingToken) = False Then RaiseSetupJ1FinalStateFailure "D4", checkDetail
    detailText = SharedText_AppendNonBlankLine(detailText, "D4: exact AllowValueListEdits/ListItemsEditForm binding proof is current.")
    checkDetail = vbNullString
    If ValidateLookupEditorIntegrationForJ1(checkDetail) = False Then RaiseSetupJ1FinalStateFailure "B2/D2/D3/D4", checkDetail
    detailText = SharedText_AppendNonBlankLine(detailText, "B2/D2/D3/D4: lookup-editor policy/intent, physical field/form/property bindings, shared-source resolution, and ExistingManagedFrm integration freshly proven current.")
    checkDetail = vbNullString
    reportToken = x__UTIL_ADC_REPORTS.ManagedReportStateTokenForFinalProof(checkDetail)
    If Len(reportToken) = 0 Then RaiseSetupJ1FinalStateFailure "E3", checkDetail
    detailText = SharedText_AppendNonBlankLine(detailText, "E3: exact managed-report physical/catalog reconciliation state proven current.")
    checkDetail = vbNullString
    If ManagedObjectDependenciesCurrentForSetup(checkDetail, dependencyToken) = False Then RaiseSetupJ1FinalStateFailure "F1", checkDetail
    detailText = SharedText_AppendNonBlankLine(detailText, "F1: generated managed-object dependency graph is current; manual dependency rows remain outside F1 ownership.")
    checkDetail = vbNullString
    ribbonPredecessorToken = RibbonDependencyTokenForSetup(checkDetail)
    If Len(ribbonPredecessorToken) = 0 Then RaiseSetupJ1FinalStateFailure "G1", checkDetail
    detailText = SharedText_AppendNonBlankLine(detailText, "G1: exact managed-description state proven current.")
    checkDetail = vbNullString
    ribbonToken = RibbonCommittedStateTokenForJ1(ribbonPredecessorToken, checkDetail, catalogToken)
    If Len(ribbonToken) = 0 Then RaiseSetupJ1FinalStateFailure "H3", checkDetail
    detailText = SharedText_AppendNonBlankLine(detailText, "H3: committed ADC Ribbon exact state and configured Ribbon-application policy proven current.")
    checkDetail = vbNullString
    navigationPaneToken = CompleteNavigationPanePlacementTokenForSetup(checkDetail, True)
    If Len(navigationPaneToken) = 0 Then RaiseSetupJ1FinalStateFailure "I3", checkDetail
    detailText = SharedText_AppendNonBlankLine(detailText, "I3: complete managed-object physical Navigation Pane placement proven current.")
    checkDetail = vbNullString
    moduleGuardStateToken = ModuleUpdateGuardFinalStateTokenForI1(checkDetail)
    If Len(moduleGuardStateToken) = 0 Then RaiseSetupJ1FinalStateFailure "Preflight/Module Guard", checkDetail
    detailText = SharedText_AppendNonBlankLine(detailText, "Module Guard: final infrastructure/manifest/enable-setting state proven without changing the user's enabled/disabled choice.")
    payloadText = "ADCSetupJ1FinalStateV1|A1-A4=Current|B1-B5=Current|C3=Current|CompleteCatalog=" & catalogToken & "|D4LookupEditorBinding=" & editorBindingToken & "|E3=" & reportToken & "|F1Dependencies=" & dependencyToken & "|G1=" & ribbonPredecessorToken & "|H3=" & ribbonToken & "|I3=" & navigationPaneToken & "|ModuleGuard=" & moduleGuardStateToken
    finalProofToken = FingerprintFromPayload(payloadText)
    detailText = SharedText_AppendNonBlankLine(detailText, "J1 final Setup verification token: " & finalProofToken)
    Exit Sub
A1NotCurrent:
    checkDetail = "A1 configuration/physical-name validation failed. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    On Error GoTo 0
    RaiseSetupJ1FinalStateFailure "A1", checkDetail
End Sub

Private Sub RaiseSetupJ1FinalStateFailure(ByVal ownerIdentity As String, ByVal failureDetail As String)
    Err.Raise vbObjectError + 9592, "VerifySetupJ1FinalState", "Setup J1 refused to persist final completion because the committed state owned by " & ownerIdentity & " is not current." & IIf(Len(Trim$(failureDetail)) = 0, vbNullString, vbCrLf & failureDetail) & vbCrLf & "J1 performed no repair for the failed predecessor domain."
End Sub

' Intended use: Delegate Setup Part identity normalization to the Defaults-owned canonical Setup Part sequence.
' Intended use: Execute one canonical implemented Setup Part after verifying its required predecessors without running earlier Parts automatically.
' Purpose: Runs setup part by identity for part identity.
' Public role: Public setup entry point. It is intentionally callable by ADC orchestration and by direct Access/VBA callers that need to run this setup scope.
' Parameters:
'   partIdentity (Input; String): Canonical Setup Part identity (for example A2 or G1) used to select, validate, and report the exact setup part being operated on.
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
' Returns: Boolean: True on the function's affirmative/success branch for run setup part by identity; False on its explicit negative, unavailable, rejected, or handled-failure branch.
Public Function RunSetupPartByIdentity(ByVal partIdentity As String, Optional ByVal showMessage As Boolean = True) As Boolean
    On Error GoTo ErrHandler
    Dim timingText As String
    Dim guardVerificationDetails As String
    Dim traceRunOwnedByProcedure As Boolean
    Dim catalogContextStarted As Boolean
    Dim lifecycleStarted As Boolean
    partIdentity = NormalizedImplementedSetupPartIdentity(partIdentity)
    CoreBeginStandaloneSetupPartRun partIdentity, guardVerificationDetails, traceRunOwnedByProcedure
    lifecycleStarted = True
    mSetupLastContractCompletedPartIdentity = vbNullString
    RequireSetupPartEntryContract partIdentity
    If Left$(partIdentity, 1) = "B" Then
        BeginCatalogStandalonePartContextForSetup partIdentity
        catalogContextStarted = True
    End If
    ExecuteSetupPartByIdentity partIdentity, timingText
    PersistCatalogOwnerProofBeforeExitIfRequired partIdentity
    RequireSetupPartExitContract partIdentity
    mSetupLastContractCompletedPartIdentity = partIdentity
    If catalogContextStarted Then
        EndCatalogSetupContext
        catalogContextStarted = False
    End If
    CoreCompleteStandaloneSetupPartRun partIdentity, timingText, showMessage, traceRunOwnedByProcedure
    RunSetupPartByIdentity = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim finalErrNumber As Long
    Dim finalErrSource As String
    Dim finalErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If catalogContextStarted Then
        On Error Resume Next
        EndCatalogSetupContext
        catalogContextStarted = False
        On Error GoTo 0
    End If
    If lifecycleStarted Then
        On Error Resume Next
        Err.Clear
        CoreHandleStandaloneSetupPartFailure partIdentity, savedErrNumber, savedErrSource, savedErrDescription, traceRunOwnedByProcedure
        finalErrNumber = Err.Number
        finalErrSource = Err.Source
        finalErrDescription = Err.Description
        On Error GoTo 0
        If finalErrNumber = 0 Then
            RunSetupPartByIdentity = False
            Exit Function
        End If
    Else
        finalErrNumber = savedErrNumber
        finalErrSource = "RunSetupPartByIdentity"
        finalErrDescription = "Could not start standalone Setup " & IIf(Len(partIdentity) = 0, "(unknown)", partIdentity) & "." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
    End If
    If showMessage Then
        CoreShowErrorMessage "Standalone Setup " & IIf(Len(partIdentity) = 0, "(unknown)", partIdentity) & " failed." & vbCrLf & vbCrLf & FailedSetupPhaseContextForMessage() & vbCrLf & vbCrLf & "Err.Number: " & CStr(finalErrNumber) & vbCrLf & "Err.Source: " & finalErrSource & vbCrLf & "Err.Description: " & finalErrDescription, vbCritical, "Standalone Setup Failed", "RunSetupPartByIdentity", "Setup"
        RunSetupPartByIdentity = False
        Exit Function
    End If
    Err.Raise finalErrNumber, finalErrSource, finalErrDescription
End Function

' Purpose: Runs setup phase by identity for phase identity.
' Public role: Public setup entry point. It is intentionally callable by ADC orchestration and by direct Access/VBA callers that need to run this setup scope.
' Parameters:
'   phaseIdentity (Input; String): One-letter Setup Phase selector. The routine trims and uppercases it, requires A through J, and uses it to select the implemented Setup Parts belonging to that phase.
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
' Returns: Boolean: True on the function's affirmative/success branch for run setup phase by identity; False on its explicit negative, unavailable, rejected, or handled-failure branch.
Public Function RunSetupPhaseByIdentity(ByVal phaseIdentity As String, Optional ByVal showMessage As Boolean = True) As Boolean
    On Error GoTo ErrHandler
    Dim implementedParts As Variant
    Dim partItem As Variant
    Dim phasePlan As New Collection
    Dim partIdentity As String
    Dim firstPartIdentity As String
    Dim operationIdentity As String
    Dim timingText As String
    Dim guardVerificationDetails As String
    Dim traceRunOwnedByProcedure As Boolean
    Dim catalogContextStarted As Boolean
    Dim lifecycleStarted As Boolean
    Dim phasePartCount As Long
    phaseIdentity = UCase$(Trim$(phaseIdentity))
    If Len(phaseIdentity) <> 1 Or InStr(1, "ABCDEFGHIJ", phaseIdentity, vbBinaryCompare) = 0 Then Err.Raise vbObjectError + 9596, "RunSetupPhaseByIdentity", "Unsupported Setup Phase identity: " & IIf(Len(phaseIdentity) = 0, "(blank)", phaseIdentity)
    implementedParts = ImplementedSetupPartIdentities()
    For Each partItem In implementedParts
        partIdentity = CStr(partItem)
        If StrComp(Left$(partIdentity, 1), phaseIdentity, vbBinaryCompare) = 0 Then phasePlan.Add partIdentity
    Next partItem
    phasePartCount = phasePlan.count
    If phasePartCount = 0 Then Err.Raise vbObjectError + 9597, "RunSetupPhaseByIdentity", "No canonical implemented Setup Parts belong to Phase " & phaseIdentity & "."
    firstPartIdentity = CStr(phasePlan(1))
    operationIdentity = "Phase " & phaseIdentity
    CoreBeginStandaloneSetupPartRun firstPartIdentity, guardVerificationDetails, traceRunOwnedByProcedure, operationIdentity, phasePartCount
    lifecycleStarted = True
    mSetupLastContractCompletedPartIdentity = vbNullString
    RequireSetupPartEntryContract firstPartIdentity
    If phaseIdentity = "B" Then
        BeginCatalogSetupContext
        catalogContextStarted = True
    End If
    For Each partItem In phasePlan
        partIdentity = CStr(partItem)
        If StrComp(partIdentity, firstPartIdentity, vbBinaryCompare) <> 0 Then RequireSetupPartEntryContract partIdentity
        ExecuteSetupPartByIdentity partIdentity, timingText
        PersistCatalogOwnerProofBeforeExitIfRequired partIdentity
        RequireSetupPartExitContract partIdentity
        mSetupLastContractCompletedPartIdentity = partIdentity
    Next partItem
    If catalogContextStarted Then
        EndCatalogSetupContext
        catalogContextStarted = False
    End If
    CoreCompleteStandaloneSetupPartRun firstPartIdentity, timingText, showMessage, traceRunOwnedByProcedure, operationIdentity
    RunSetupPhaseByIdentity = True
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim finalErrNumber As Long
    Dim finalErrSource As String
    Dim finalErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If catalogContextStarted Then
        On Error Resume Next
        EndCatalogSetupContext
        catalogContextStarted = False
        On Error GoTo 0
    End If
    If lifecycleStarted Then
        On Error Resume Next
        Err.Clear
        CoreHandleStandaloneSetupPartFailure firstPartIdentity, savedErrNumber, savedErrSource, savedErrDescription, traceRunOwnedByProcedure, operationIdentity
        finalErrNumber = Err.Number
        finalErrSource = Err.Source
        finalErrDescription = Err.Description
        On Error GoTo 0
        If finalErrNumber = 0 Then
            RunSetupPhaseByIdentity = False
            Exit Function
        End If
    Else
        finalErrNumber = savedErrNumber
        finalErrSource = "RunSetupPhaseByIdentity"
        finalErrDescription = "Could not start Setup " & IIf(Len(operationIdentity) = 0, "Phase (unknown)", operationIdentity) & "." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
    End If
    If showMessage Then
        CoreShowErrorMessage "Setup " & IIf(Len(operationIdentity) = 0, "Phase (unknown)", operationIdentity) & " failed." & vbCrLf & vbCrLf & FailedSetupPhaseContextForMessage() & vbCrLf & vbCrLf & "Err.Number: " & CStr(finalErrNumber) & vbCrLf & "Err.Source: " & finalErrSource & vbCrLf & "Err.Description: " & finalErrDescription, vbCritical, "Setup Phase Failed", "RunSetupPhaseByIdentity", "Setup"
        RunSetupPhaseByIdentity = False
        Exit Function
    End If
    Err.Raise finalErrNumber, finalErrSource, finalErrDescription
End Function

Private Function NormalizedImplementedSetupPartIdentity(ByVal partIdentity As String) As String
    Dim requestedIdentity As String
    requestedIdentity = partIdentity
    partIdentity = ImplementedSetupPartIdentity(partIdentity)
    If Len(partIdentity) = 0 Then Err.Raise vbObjectError + 9585, "NormalizedImplementedSetupPartIdentity", "Unsupported Setup Part identity: " & IIf(Len(Trim$(requestedIdentity)) = 0, "(blank)", requestedIdentity)
    NormalizedImplementedSetupPartIdentity = partIdentity
End Function

Private Sub ExecuteSetupPartByIdentity(ByVal partIdentity As String, ByRef timingText As String)
    Select Case partIdentity
        Case "A1": ExecuteSetupPartA1 timingText
        Case "A2": ExecuteSetupPartA2 timingText
        Case "A3": ExecuteSetupPartA3 timingText
        Case "A4": ExecuteSetupPartA4 timingText
        Case "B1": ExecuteSetupPartB1 timingText
        Case "B2": ExecuteSetupPartB2 timingText
        Case "B3": ExecuteSetupPartB3 timingText
        Case "B4": ExecuteSetupPartB4 timingText
        Case "B5": ExecuteSetupPartB5 timingText
        Case "C1": ExecuteSetupPartC1 timingText
        Case "C2": ExecuteSetupPartC2 timingText
        Case "C3": ExecuteSetupPartC3 timingText
        Case "D1": ExecuteSetupPartD1 timingText
        Case "D2": ExecuteSetupPartD2 timingText
        Case "D3": ExecuteSetupPartD3 timingText
        Case "D4": ExecuteSetupPartD4 timingText
        Case "E1": ExecuteSetupPartE1 timingText
        Case "E2": ExecuteSetupPartE2 timingText
        Case "E3": ExecuteSetupPartE3 timingText
        Case "F1": ExecuteSetupPartF1 timingText
        Case "G1": ExecuteSetupPartG1 timingText
        Case "H1": ExecuteSetupPartH1 timingText
        Case "H2": ExecuteSetupPartH2 timingText
        Case "H3": ExecuteSetupPartH3 timingText
        Case "I1": ExecuteSetupPartI1 timingText
        Case "I2": ExecuteSetupPartI2 timingText
        Case "I3": ExecuteSetupPartI3 timingText
        Case "J1": ExecuteSetupPartJ1 timingText
        Case Else: Err.Raise vbObjectError + 9595, "ExecuteSetupPartByIdentity", "Canonical Setup Part has no execution handler: " & partIdentity
    End Select
End Sub

Private Sub ExecuteSetupPartA1(ByRef timingText As String)
    Dim phaseStartedTimer As Single
    CoreBeginSetupPhasePart "A1", "Tables, Queries, and DB Inventory", 1, 4, "Validate Configuration and Object Names"
    phaseStartedTimer = Timer
    CoreBeginSetupNamingSnapshot
    ValidateConfiguredPhysicalNamesForSetup "ExecuteSetupPartA1"
    timingText = SharedText_AppendNonBlankLine(timingText, "A1 - Validate configuration and physical names: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseStartedTimer)))
End Sub

Private Sub ExecuteSetupPartA2(ByRef timingText As String)
    Dim phaseStartedTimer As Single
    Dim schemaRepairDetailsText As String
    Dim namingInputsChanged As Boolean
    Dim preflightSchemaEvidenceReady As Boolean
    Dim preflightSchemaRepairNeeded As Boolean
    Dim preflightSchemaRepairSummary As String
    Dim schemaRepairEvidence As TSchemaRepairPreflightEvidence
    CoreBeginSetupPhasePart "A2", "Tables, Queries, and DB Inventory", 2, 4, "Build/Repair Core ADC Tables"
    phaseStartedTimer = Timer
    ClearTableRepairWarningResult
    preflightSchemaEvidenceReady = TryGetSetupPreflightA2SchemaEvidence(preflightSchemaRepairNeeded, preflightSchemaRepairSummary)
    schemaRepairEvidence.IsKnown = preflightSchemaEvidenceReady
    schemaRepairEvidence.repairNeeded = preflightSchemaRepairNeeded
    schemaRepairEvidence.Summary = preflightSchemaRepairSummary
    If preflightSchemaEvidenceReady Then AddTraceNote "A2 received exact same-run TABLES-owned schema evidence from Setup Preflight."
    If RepairCoreDataTableFoundationForSetup(schemaRepairEvidence, False, schemaRepairDetailsText, namingInputsChanged, True) = False Then Err.Raise vbObjectError + 6286, "ExecuteSetupPartA2", "A2 core CodeDb data-table foundation repair returned False." & schemaRepairDetailsText
    AddTableRepairWarningToParent
    If namingInputsChanged Then CoreMarkSetupNamingSnapshotDirty "Setup A2", True
    CoreEnsureSetupNamingSnapshotCurrent
    timingText = SharedText_AppendNonBlankLine(timingText, "A2 - Repair core data tables: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseStartedTimer)))
End Sub

Private Sub ExecuteSetupPartA3(ByRef timingText As String)
    Dim phaseStartedTimer As Single
    Dim queryRepairDetailsText As String
    CoreBeginSetupPhasePart "A3", "Tables, Queries, and DB Inventory", 3, 4, "Build/Repair Core and UI Support Queries"
    phaseStartedTimer = Timer
    If RepairNonFormQueryFoundationForSetup(False, queryRepairDetailsText) = False Then Err.Raise vbObjectError + 6289, "ExecuteSetupPartA3", "A3 complete non-form saved-query foundation repair returned False." & queryRepairDetailsText
    timingText = SharedText_AppendNonBlankLine(timingText, "A3 - Repair all non-form saved-query foundations: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseStartedTimer)))
End Sub

Private Sub ExecuteSetupPartA4(ByRef timingText As String)
    Dim phaseStartedTimer As Single
    CoreBeginSetupPhasePart "A4", "Tables, Queries, and DB Inventory", 4, 4, "Build/Repair Reference and Add-In Inventory"
    phaseStartedTimer = Timer
    RepairReferenceInventoryFoundationForSetup False
    timingText = SharedText_AppendNonBlankLine(timingText, "A4 - Repair reference/Add-In inventory foundation: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseStartedTimer)))
End Sub

Private Sub ExecuteSetupPartB1(ByRef timingText As String)
    Dim phaseStartedTimer As Single
    Dim catalogDetailsText As String
    Dim namingInputsChanged As Boolean
    CoreBeginSetupPhasePart "B1", "Managed Object Catalog", 1, 5, "Build/Repair Managed Object Catalog Structure"
    phaseStartedTimer = Timer
    catalogDetailsText = vbNullString
    If RepairCatalogStructureForSetup(False, catalogDetailsText, namingInputsChanged) = False Then Err.Raise vbObjectError + 6291, "ExecuteSetupPartB1", "B1 Managed Object Catalog structure repair returned False." & vbCrLf & catalogDetailsText
    If namingInputsChanged Then CoreMarkSetupNamingSnapshotDirty "Setup B1"
    CoreEnsureSetupNamingSnapshotCurrent
    timingText = SharedText_AppendNonBlankLine(timingText, "B1 - Repair Managed Object Catalog structure: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseStartedTimer)))
End Sub

Private Sub ExecuteSetupPartB2(ByRef timingText As String)
    Dim phaseStartedTimer As Single
    Dim catalogDetailsText As String
    CoreBeginSetupPhasePart "B2", "Managed Object Catalog", 2, 5, "Build/Repair Catalog Reference Data, Default Values, and Lookup Metadata"
    phaseStartedTimer = Timer
    catalogDetailsText = vbNullString
    If RepairCatalogReferenceDataForSetup(False, catalogDetailsText) = False Then Err.Raise vbObjectError + 6292, "ExecuteSetupPartB2", "B2 catalog reference/default data and lookup metadata repair returned False." & vbCrLf & catalogDetailsText
    timingText = SharedText_AppendNonBlankLine(timingText, "B2 - Repair catalog reference/default data and lookup metadata: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseStartedTimer)))
End Sub

Private Sub ExecuteSetupPartB3(ByRef timingText As String)
    Dim phaseStartedTimer As Single
    Dim catalogDetailsText As String
    Dim coreNamingInputsChanged As Boolean
    Dim presentationNamingInputsChanged As Boolean
    CoreBeginSetupPhasePart "B3", "Managed Object Catalog", 3, 5, "Reconcile Core and UI Managed Definitions"
    phaseStartedTimer = Timer
    catalogDetailsText = vbNullString
    If RepairCoreManagedObjectDefinitionsForSetup(False, catalogDetailsText, coreNamingInputsChanged) = False Then Err.Raise vbObjectError + 6293, "ExecuteSetupPartB3", "B3 core managed-object definition reconciliation returned False." & vbCrLf & catalogDetailsText
    catalogDetailsText = vbNullString
    If RepairNonFormPresentationDefinitionsForSetup(False, catalogDetailsText, presentationNamingInputsChanged) = False Then Err.Raise vbObjectError + 6295, "ExecuteSetupPartB3", "B3 non-form presentation-metadata definition reconciliation returned False." & vbCrLf & catalogDetailsText
    If coreNamingInputsChanged Or presentationNamingInputsChanged Then CoreMarkSetupNamingSnapshotDirty "Setup B3"
    CoreEnsureSetupNamingSnapshotCurrent
    timingText = SharedText_AppendNonBlankLine(timingText, "B3 - Reconcile core non-form managed definitions: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseStartedTimer)))
End Sub

Private Sub ExecuteSetupPartB4(ByRef timingText As String)
    Dim phaseStartedTimer As Single
    Dim catalogDetailsText As String
    Dim namingInputsChanged As Boolean
    CoreBeginSetupPhasePart "B4", "Managed Object Catalog", 4, 5, "Reconcile System Macro and Help Document Macro Definitions"
    phaseStartedTimer = Timer
    catalogDetailsText = vbNullString
    If RepairSystemHelpMacroDefinitionsForSetup(False, catalogDetailsText, namingInputsChanged) = False Then Err.Raise vbObjectError + 6296, "ExecuteSetupPartB4", "B4 System Macro and Help Document Macro source and managed-definition reconciliation returned False." & vbCrLf & catalogDetailsText
    If namingInputsChanged Then CoreMarkSetupNamingSnapshotDirty "Setup B4"
    CoreEnsureSetupNamingSnapshotCurrent
    timingText = SharedText_AppendNonBlankLine(timingText, "B4 - Reconcile System Macro and Help Document Macro source and managed definitions: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseStartedTimer)))
End Sub

Private Sub ExecuteSetupPartB5(ByRef timingText As String)
    Dim phaseStartedTimer As Single
    Dim catalogDetailsText As String
    CoreBeginSetupPhasePart "B5", "Managed Object Catalog", 5, 5, "Build/Repair Catalog Support Queries and Validate Managed Object Catalog"
    phaseStartedTimer = Timer
    catalogDetailsText = vbNullString
    If RepairManagedDefinitionSupportQueriesForSetup(False, catalogDetailsText) = False Then Err.Raise vbObjectError + 6297, "ExecuteSetupPartB5", "B5 managed-definition support-query finalization returned False." & vbCrLf & catalogDetailsText
    catalogDetailsText = vbNullString
    If ValidateManagedObjectPlacementResolverCoveragePhase(False, catalogDetailsText, False) = False Then Err.Raise vbObjectError + 6299, "ExecuteSetupPartB5", "B5 non-form placement resolver coverage validation returned False." & vbCrLf & catalogDetailsText
    catalogDetailsText = vbNullString
    If ValidateManagedObjectCatalogForSetup(False, catalogDetailsText) = False Then Err.Raise vbObjectError + 6300, "ExecuteSetupPartB5", "B5 non-form catalog validation returned False." & vbCrLf & catalogDetailsText
    timingText = SharedText_AppendNonBlankLine(timingText, "B5 - Finalize and validate the non-form catalog: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseStartedTimer)))
End Sub

Private Sub ExecuteSetupPartC1(ByRef timingText As String)
    Dim phaseStartedTimer As Single
    Dim catalogDetailsText As String
    CoreBeginSetupPhasePart "C1", "Physical System and Help Document Macros", 1, 3, "Verify Physical System and Help Document Macros"
    phaseStartedTimer = Timer
    catalogDetailsText = vbNullString
    PhaseCVerifyPhysicalSystemHelpMacroState catalogDetailsText, CoreSetupHelpDocumentSynchronizationScope()
    timingText = SharedText_AppendNonBlankLine(timingText, "C1 - Verify physical System Macro and Help Document Macro state: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseStartedTimer)))
End Sub

Private Sub ExecuteSetupPartC2(ByRef timingText As String)
    Dim phaseStartedTimer As Single
    Dim catalogDetailsText As String
    CoreBeginSetupPhasePart "C2", "Physical System and Help Document Macros", 2, 3, "Build Required System and Help Document Macros"
    phaseStartedTimer = Timer
    catalogDetailsText = vbNullString
    PhaseCBuildRequiredSystemHelpMacros catalogDetailsText
    timingText = SharedText_AppendNonBlankLine(timingText, "C2 - Build required System Macro and Help Document Macros: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseStartedTimer)))
End Sub

Private Sub ExecuteSetupPartC3(ByRef timingText As String)
    Dim phaseStartedTimer As Single
    Dim catalogDetailsText As String
    CoreBeginSetupPhasePart "C3", "Physical System and Help Document Macros", 3, 3, "Finalize and Verify System and Help Document Macros"
    phaseStartedTimer = Timer
    catalogDetailsText = vbNullString
    PhaseCFinalizePhysicalSystemHelpMacroState catalogDetailsText
    timingText = SharedText_AppendNonBlankLine(timingText, "C3 - Finalize and verify physical macro state: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseStartedTimer)))
End Sub

Private Sub ExecuteSetupPartD1(ByRef timingText As String)
    Dim phaseStartedTimer As Single
    Dim persistentLauncherDetails As String
    mSetupD2ProofTraceRunID = vbNullString
    mSetupD2ProofToken = vbNullString
    mSetupD3ProofTraceRunID = vbNullString
    mSetupD3ProofToken = vbNullString
    phaseStartedTimer = Timer
    CoreBeginSetupPhasePart "D1", "Form Creation and Reconciliation", 1, 4, "Preflight Form Creation"
    x__UTIL_ADC_UI_FORMS.EnsurePersistentGeneratedLauncherFormsForSetup persistentLauncherDetails
    If Len(persistentLauncherDetails) > 0 Then AddTraceNote persistentLauncherDetails
    PreflightGeneratedFormsForSetup True, False, "All"
    timingText = SharedText_AppendNonBlankLine(timingText, "D1 - Preflight Form Creation: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseStartedTimer)))
End Sub

Private Sub ExecuteSetupPartD2(ByRef timingText As String)
    Dim phaseStartedTimer As Single
    Dim proofDetail As String
    mSetupD2ProofTraceRunID = vbNullString
    mSetupD2ProofToken = vbNullString
    phaseStartedTimer = Timer
    CoreBeginSetupPhasePart "D2", "Form Creation and Reconciliation", 2, 4, "Build/Verify ChatGPT UI Forms"
    ApplyGeneratedFormsFromSetupPreflight False, "ChatGPTUIForms"
    If TryGetGeneratedFormSameRunD2Proof(mSetupD2ProofTraceRunID, mSetupD2ProofToken, proofDetail) Then
        AddTraceNote "D2 retained exact same-run ChatGPT UI Form partition proof for D4. TraceRunID=" & mSetupD2ProofTraceRunID & "; token=" & mSetupD2ProofToken
    Else
        Err.Raise vbObjectError + 9570, "ExecuteSetupPartD2", "D2 completed without the required same-run ChatGPT UI Form partition proof. " & proofDetail
    End If
    timingText = SharedText_AppendNonBlankLine(timingText, "D2 - Build/Verify ChatGPT UI Forms: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseStartedTimer)))
End Sub

Private Sub ExecuteSetupPartD3(ByRef timingText As String)
    Dim phaseStartedTimer As Single
    Dim proofDetail As String
    mSetupD3ProofTraceRunID = vbNullString
    mSetupD3ProofToken = vbNullString
    phaseStartedTimer = Timer
    CoreBeginSetupPhasePart "D3", "Form Creation and Reconciliation", 3, 4, "Build/Verify ADC Support Forms"
    ApplyGeneratedFormsFromSetupPreflight False, "ADCSupportForms"
    If TryGetGeneratedFormSameRunD3Proof(mSetupD3ProofTraceRunID, mSetupD3ProofToken, proofDetail) Then
        AddTraceNote "D3 retained exact same-run ADC Support Form partition proof for D4. TraceRunID=" & mSetupD3ProofTraceRunID & "; token=" & mSetupD3ProofToken
    Else
        Err.Raise vbObjectError + 9571, "ExecuteSetupPartD3", "D3 completed without the required same-run ADC Support Form partition proof. " & proofDetail
    End If
    timingText = SharedText_AppendNonBlankLine(timingText, "D3 - Build/Verify ADC Support Forms: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseStartedTimer)))
End Sub

Private Sub ExecuteSetupPartD4(ByRef timingText As String)
    Dim phaseStartedTimer As Single
    Dim detailText As String
    Dim currentTraceRunID As String
    Dim sameRunD2ProofAvailable As Boolean
    Dim sameRunD3ProofAvailable As Boolean
    phaseStartedTimer = Timer
    CoreBeginSetupPhasePart "D4", "Form Creation and Reconciliation", 4, 4, "Reconcile Managed Forms and Finalize Logical Catalog"
    currentTraceRunID = GetActiveOrLastTraceRunID()
    If Len(currentTraceRunID) = 0 Then Err.Raise vbObjectError + 9572, "ExecuteSetupPartD4", "D4 requires an active Setup trace run."
    sameRunD2ProofAvailable = (StrComp(currentTraceRunID, mSetupD2ProofTraceRunID, vbBinaryCompare) = 0 And Len(mSetupD2ProofToken) > 0)
    sameRunD3ProofAvailable = (StrComp(currentTraceRunID, mSetupD3ProofTraceRunID, vbBinaryCompare) = 0 And Len(mSetupD3ProofToken) > 0)
    detailText = vbNullString
    If sameRunD2ProofAvailable And sameRunD3ProofAvailable Then
        If FinalizeGeneratedFormCombinedSameRunProofForD4(mSetupD2ProofToken, mSetupD3ProofToken, detailText) = False Then Err.Raise vbObjectError + 9575, "ExecuteSetupPartD4", detailText
    Else
        If GeneratedFormsCurrentForSetup(detailText) = False Then Err.Raise vbObjectError + 9575, "ExecuteSetupPartD4", detailText
        detailText = "D4 used authoritative full-current-state generated-form verification because exact same-run D2+D3 partition proof was unavailable for the active trace." & IIf(Len(Trim$(detailText)) = 0, vbNullString, vbCrLf & detailText)
    End If
    If Len(Trim$(detailText)) > 0 Then AddTraceNote detailText
    detailText = vbNullString
    If RepairManagedFormDomainForSetup(False, detailText) = False Then Err.Raise vbObjectError + 9576, "ExecuteSetupPartD4", detailText
    detailText = vbNullString
    If ReconcileLookupEditorBindingsForSetup(detailText) = False Then Err.Raise vbObjectError + 9601, "ExecuteSetupPartD4", detailText
    If Len(Trim$(detailText)) > 0 Then AddTraceNote detailText
    detailText = vbNullString
    If FinalizeCompleteCatalogStateForSetup(False, detailText) = False Then Err.Raise vbObjectError + 9587, "ExecuteSetupPartD4", detailText
    timingText = SharedText_AppendNonBlankLine(timingText, "D4 - Reconcile Managed Forms and Finalize Logical Catalog: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseStartedTimer)))
End Sub

Private Sub ExecuteSetupPartE1(ByRef timingText As String)
    Dim phaseStartedTimer As Single
    Dim detailText As String
    phaseStartedTimer = Timer
    CoreBeginSetupPhasePart "E1", "Reports", 1, 3, "Preflight ChatGPT UI Reports"
    If x__UTIL_ADC_REPORTS.PreflightManagedReportsForSetup(detailText) = False Then Err.Raise vbObjectError + 9596, "ExecuteSetupPartE1", detailText
    timingText = SharedText_AppendNonBlankLine(timingText, "E1 - Preflight ChatGPT UI Reports: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseStartedTimer)))
End Sub

Private Sub ExecuteSetupPartE2(ByRef timingText As String)
    Dim phaseStartedTimer As Single
    Dim detailText As String
    phaseStartedTimer = Timer
    CoreBeginSetupPhasePart "E2", "Reports", 2, 3, "Build/Repair and Verify ChatGPT UI Reports"
    If Len(x__UTIL_ADC_REPORTS.ReportPreflightTokenForE2(detailText)) = 0 Then Err.Raise vbObjectError + 9597, "ExecuteSetupPartE2", detailText
    detailText = vbNullString
    If x__UTIL_ADC_REPORTS.BuildAndVerifyManagedReportsForSetup(detailText) = False Then Err.Raise vbObjectError + 9598, "ExecuteSetupPartE2", detailText
    timingText = SharedText_AppendNonBlankLine(timingText, "E2 - Build/Repair and Verify ChatGPT UI Reports: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseStartedTimer)))
End Sub

Private Sub ExecuteSetupPartE3(ByRef timingText As String)
    Dim phaseStartedTimer As Single
    Dim detailText As String
    phaseStartedTimer = Timer
    CoreBeginSetupPhasePart "E3", "Reports", 3, 3, "Reconcile and Verify ChatGPT UI Reports"
    If Len(x__UTIL_ADC_REPORTS.ReportBuildTokenForE3(detailText)) = 0 Then Err.Raise vbObjectError + 9599, "ExecuteSetupPartE3", detailText
    detailText = vbNullString
    If x__UTIL_ADC_REPORTS.ReconcileManagedReportDomainForSetup(detailText) = False Then Err.Raise vbObjectError + 9600, "ExecuteSetupPartE3", detailText
    timingText = SharedText_AppendNonBlankLine(timingText, "E3 - Reconcile managed-report physical domain: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseStartedTimer)))
End Sub

Private Sub ExecuteSetupPartF1(ByRef timingText As String)
    Dim phaseStartedTimer As Single
    Dim detailText As String
    Dim dependencyToken As String
    phaseStartedTimer = Timer
    CoreBeginSetupPhasePart "F1", "Managed Object Dependencies", 1, 1, "Build/Repair and Verify Managed Object Dependencies"
    If ReconcileManagedObjectDependenciesForSetup(detailText, dependencyToken) = False Then Err.Raise vbObjectError + 9611, "ExecuteSetupPartF1", detailText
    If Len(Trim$(detailText)) > 0 Then AddTraceNote detailText
    timingText = SharedText_AppendNonBlankLine(timingText, "F1 - Build/repair and verify managed-object dependencies: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseStartedTimer)))
End Sub

Private Sub ExecuteSetupPartG1(ByRef timingText As String)
    Dim phaseStartedTimer As Single
    Dim detailText As String
    phaseStartedTimer = Timer
    CoreBeginSetupPhasePart "G1", "Object and Field Descriptions", 1, 1, "Apply and Verify Object and Field Descriptions"
    If ApplyAndVerifyManagedDescriptionsForSetup(False, detailText) = False Then Err.Raise vbObjectError + 9588, "ExecuteSetupPartG1", detailText
    timingText = SharedText_AppendNonBlankLine(timingText, "G1 - Apply and verify managed descriptions: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseStartedTimer)))
End Sub

Private Sub ExecuteSetupPartH1(ByRef timingText As String)
    Dim phaseStartedTimer As Single
    phaseStartedTimer = Timer
    CoreBeginSetupPhasePart "H1", "ADC Ribbon", 1, 3, "Prepare ADC Ribbon Build"
    PrepareCustomRibbonBuild False, True
    timingText = SharedText_AppendNonBlankLine(timingText, "H1 - Prepare Ribbon plan: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseStartedTimer)))
End Sub

Private Sub ExecuteSetupPartH2(ByRef timingText As String)
    Dim phaseStartedTimer As Single
    Dim detailText As String
    phaseStartedTimer = Timer
    CoreBeginSetupPhasePart "H2", "ADC Ribbon", 2, 3, "Build and Validate Staged ADC Ribbon"
    If RibbonSetupPlanReadyForH2(detailText) = False Then Err.Raise vbObjectError + 9590, "ExecuteSetupPartH2", detailText
    BuildStandardRibbonPhase
    timingText = SharedText_AppendNonBlankLine(timingText, "H2 - Build and validate staged Ribbon: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseStartedTimer)))
End Sub

Private Sub ExecuteSetupPartH3(ByRef timingText As String)
    Dim phaseStartedTimer As Single
    Dim detailText As String
    phaseStartedTimer = Timer
    CoreBeginSetupPhasePart "H3", "ADC Ribbon", 3, 3, "Commit Validated ADC Ribbon"
    If RibbonSetupBuildReadyForH3(detailText) = False Then Err.Raise vbObjectError + 9591, "ExecuteSetupPartH3", detailText
    CommitCustomRibbonBuild
    timingText = SharedText_AppendNonBlankLine(timingText, "H3 - Commit validated Ribbon: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseStartedTimer)))
End Sub

Private Sub ExecuteSetupPartI1(ByRef timingText As String)
    Dim phaseStartedTimer As Single
    CoreBeginSetupPhasePart "I1", "Navigation Pane Configuration", 1, 3, "Create/Verify Configured Navigation Pane Groups"
    phaseStartedTimer = Timer
    EnsureConfiguredNavigationPaneGroupsExist "I1", True
    timingText = SharedText_AppendNonBlankLine(timingText, "I1 - Create configured Navigation Pane groups: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseStartedTimer)))
End Sub

Private Sub ExecuteSetupPartI2(ByRef timingText As String)
    Dim phaseStartedTimer As Single
    CoreBeginSetupPhasePart "I2", "Navigation Pane Configuration", 2, 3, "Preflight Managed Object Navigation Pane Placement"
    phaseStartedTimer = Timer
    PreflightManagedObjectsNavigationPaneForSetup
    timingText = SharedText_AppendNonBlankLine(timingText, "I2 - Preflight complete managed-object Navigation Pane placement: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseStartedTimer)))
End Sub

Private Sub ExecuteSetupPartI3(ByRef timingText As String)
    Dim phaseStartedTimer As Single
    CoreBeginSetupPhasePart "I3", "Navigation Pane Configuration", 3, 3, "Apply and Verify Managed Object Navigation Pane Placement"
    phaseStartedTimer = Timer
    ApplyPreflightedManagedObjectsNavigationPaneForSetup False
    timingText = SharedText_AppendNonBlankLine(timingText, "I3 - Apply and verify complete managed-object Navigation Pane placement: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseStartedTimer)))
End Sub

Private Sub ExecuteSetupPartJ1(ByRef timingText As String)
    Dim phaseStartedTimer As Single
    Dim finalProofToken As String
    Dim moduleGuardStateToken As String
    Dim detailText As String
    phaseStartedTimer = Timer
    CoreBeginSetupPhasePart "J1", "Final Setup Verification", 1, 1, "Verify and Persist Final Setup State"
    VerifySetupJ1FinalState finalProofToken, moduleGuardStateToken, detailText
    CoreStageFinalSetupStateForJ1 finalProofToken, moduleGuardStateToken, detailText
    timingText = SharedText_AppendNonBlankLine(timingText, "J1 - Verify and persist final Setup state: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseStartedTimer)))
End Sub

' Intended use: Return the number of canonical implemented Setup Parts without requiring callers to maintain a separate literal count.
Private Sub ExecuteSetupThroughPhase(ByVal targetPhase As String, ByRef timingText As String, ByRef catalogContextStarted As Boolean)
    Dim partIdentities As Variant
    Dim partItem As Variant
    Dim partIdentity As String
    If Len(targetPhase) <> 1 Or InStr(1, "CDEFGHIJ", targetPhase, vbBinaryCompare) = 0 Then Err.Raise vbObjectError + 9594, "ExecuteSetupThroughPhase", "Unsupported Setup target phase: " & IIf(Len(targetPhase) = 0, "(blank)", targetPhase)
    partIdentities = ImplementedSetupPartIdentities()
    mSetupLastContractCompletedPartIdentity = vbNullString
    For Each partItem In partIdentities
        partIdentity = CStr(partItem)
        If StrComp(Left$(partIdentity, 1), targetPhase, vbBinaryCompare) > 0 Then Exit For
        RequireSetupPartEntryContract partIdentity
        If partIdentity = "B1" Then
            BeginCatalogSetupContext
            catalogContextStarted = True
        End If
        ExecuteSetupPartByIdentity partIdentity, timingText
        PersistCatalogOwnerProofBeforeExitIfRequired partIdentity
        RequireSetupPartExitContract partIdentity
        mSetupLastContractCompletedPartIdentity = partIdentity
        If partIdentity = "B5" Then
            EndCatalogSetupContext
            catalogContextStarted = False
        End If
        CommitSetupRollbackGroupBoundaryIfRequired partIdentity, targetPhase
    Next partItem
End Sub

Private Sub CommitSetupRollbackGroupBoundaryIfRequired(ByVal completedPartIdentity As String, ByVal targetPhase As String)
    Select Case completedPartIdentity
        Case "C3"
            If StrComp(targetPhase, "C", vbBinaryCompare) > 0 Then CoreCommitSetupRollbackGroupBoundary "A1-C3 tables, queries, DB Inventory, Managed Object Catalog, and macros", "D1-G1 forms, reports, managed-object dependencies, and Object and Field Descriptions"
        Case "G1"
            If StrComp(targetPhase, "G", vbBinaryCompare) > 0 Then
                CoreCommitSetupRollbackGroupBoundary "D1-G1 forms, reports, managed-object dependencies, and Object and Field Descriptions", "H1-H3 ADC Ribbon"
            End If
        Case "H3"
            If StrComp(targetPhase, "H", vbBinaryCompare) > 0 Then CoreCommitSetupRollbackGroupBoundary "H1-H3 ADC Ribbon", "I1-I3 Navigation Pane"
        Case "I3"
            If StrComp(targetPhase, "I", vbBinaryCompare) > 0 Then CoreCommitSetupRollbackGroupBoundary "I1-I3 Navigation Pane", "J1 final verification and evidence"
    End Select
End Sub

Private Sub PersistCatalogOwnerProofBeforeExitIfRequired(ByVal completedPartIdentity As String)
    'B1-B4 each persist their own durable cumulative catalog proof while the owning Part is still active. The exit contract then validates that persisted proof before Setup may advance.
    Select Case UCase$(Trim$(completedPartIdentity))
        Case "B1", "B2", "B3", "B4"
            PersistCatalogStandalonePartProofForSetup UCase$(Trim$(completedPartIdentity))
    End Select
End Sub

Private Function ImmediateImplementedPredecessorIdentity(ByVal requestedPartIdentity As String) As String
    Dim partIdentities As Variant
    Dim itemIndex As Long
    requestedPartIdentity = UCase$(Trim$(requestedPartIdentity))
    partIdentities = ImplementedSetupPartIdentities()
    For itemIndex = LBound(partIdentities) To UBound(partIdentities)
        If StrComp(CStr(partIdentities(itemIndex)), requestedPartIdentity, vbBinaryCompare) = 0 Then
            If itemIndex > LBound(partIdentities) Then ImmediateImplementedPredecessorIdentity = CStr(partIdentities(itemIndex - 1))
            Exit Function
        End If
    Next itemIndex
End Function

Private Sub RequireSetupPartEntryContract(ByVal requestedPartIdentity As String)
    Dim expectedPredecessor As String
    requestedPartIdentity = UCase$(Trim$(requestedPartIdentity))
    expectedPredecessor = ImmediateImplementedPredecessorIdentity(requestedPartIdentity)
    If Len(mSetupLastContractCompletedPartIdentity) > 0 Then
        If StrComp(mSetupLastContractCompletedPartIdentity, expectedPredecessor, vbBinaryCompare) <> 0 Then
            Err.Raise vbObjectError + 10160, "RequireSetupPartEntryContract", "Setup Part ordering violation." & vbCrLf & "Requested Part: " & requestedPartIdentity & vbCrLf & "Required immediate predecessor: " & IIf(Len(expectedPredecessor) = 0, "(none)", expectedPredecessor) & vbCrLf & "Last Part that passed its exit contract: " & mSetupLastContractCompletedPartIdentity & vbCrLf & "Execution refused before the requested Part began."
        End If
        Exit Sub
    End If
    If requestedPartIdentity <> "A1" Then RequirePersistedSetupPartPrerequisites requestedPartIdentity
End Sub

Private Sub RaiseSetupPartExitContractFailure(ByVal completedPartIdentity As String, ByVal detailText As String)
    Err.Raise vbObjectError + 10161, "RequireSetupPartExitContract", "Setup " & completedPartIdentity & " completed its execution routine but failed its owning-part exit contract. Setup will not advance to the next Part." & IIf(Len(Trim$(detailText)) = 0, vbNullString, vbCrLf & detailText)
End Sub

Private Sub RequireSetupPartExitContract(ByVal completedPartIdentity As String)
    Dim detailText As String
    Dim tokenText As String
    Dim rowCount As Long
    completedPartIdentity = UCase$(Trim$(completedPartIdentity))
    Select Case completedPartIdentity
        Case "A1"
            On Error GoTo A1ExitNotCurrent
            ValidateConfiguredPhysicalNamesForSetup "RequireSetupPartExitContract.A1", False
            On Error GoTo 0
        Case "A2"
            If CoreDataTableFoundationCurrentForSetup(detailText) = False Then RaiseSetupPartExitContractFailure completedPartIdentity, detailText
        Case "A3"
            If NonFormQueryFoundationCurrentForSetup(detailText) = False Then RaiseSetupPartExitContractFailure completedPartIdentity, detailText
            detailText = vbNullString
            If FormDesignQueryFoundationCurrentForSetup(detailText) = False Then RaiseSetupPartExitContractFailure completedPartIdentity, detailText
        Case "A4"
            If ReferenceInventoryFoundationCurrentForSetup(detailText) = False Then RaiseSetupPartExitContractFailure completedPartIdentity, detailText
        Case "B1"
            If CatalogPrerequisiteCurrentForSetup("B2", detailText) = False Then RaiseSetupPartExitContractFailure completedPartIdentity, detailText
        Case "B2"
            If CatalogPrerequisiteCurrentForSetup("B3", detailText) = False Then RaiseSetupPartExitContractFailure completedPartIdentity, detailText
        Case "B3"
            If CatalogPrerequisiteCurrentForSetup("B4", detailText) = False Then RaiseSetupPartExitContractFailure completedPartIdentity, detailText
        Case "B4"
            If CatalogPrerequisiteCurrentForSetup("B5", detailText) = False Then RaiseSetupPartExitContractFailure completedPartIdentity, detailText
        Case "B5"
            If CatalogPrerequisiteCurrentForSetup("C1", detailText) = False Then RaiseSetupPartExitContractFailure completedPartIdentity, detailText
        Case "C1", "C2"
            'C1/C2 are operation-scoped plan/mutation stages. Their execution routines already validate the C1 plan source and mutation targets; C3 owns the durable physical-macro proof.
        Case "C3"
            If PhaseC3CurrentForSetup(detailText) = False Then RaiseSetupPartExitContractFailure completedPartIdentity, detailText
        Case "D1"
            If SetupGeneratedFormPreflightReadyForSetup(detailText) = False Then RaiseSetupPartExitContractFailure completedPartIdentity, detailText
        Case "D2"
            If Len(mSetupD2ProofToken) = 0 Or Len(mSetupD2ProofTraceRunID) = 0 Then RaiseSetupPartExitContractFailure completedPartIdentity, "D2 same-run generated-form partition proof is unavailable."
        Case "D3"
            If Len(mSetupD3ProofToken) = 0 Or Len(mSetupD3ProofTraceRunID) = 0 Then RaiseSetupPartExitContractFailure completedPartIdentity, "D3 same-run generated-form partition proof is unavailable."
        Case "D4"
            If ManagedFormDomainCurrentForSetup(detailText) = False Then RaiseSetupPartExitContractFailure completedPartIdentity, detailText
            detailText = vbNullString
            If D4LookupEditorBindingCurrentForSetup(detailText) = False Then RaiseSetupPartExitContractFailure completedPartIdentity, detailText
            detailText = vbNullString
            If Len(CompleteCatalogStateTokenForSetup(detailText)) = 0 Then RaiseSetupPartExitContractFailure completedPartIdentity, detailText
        Case "E1"
            If Len(x__UTIL_ADC_REPORTS.ReportPreflightTokenForE2(detailText)) = 0 Then RaiseSetupPartExitContractFailure completedPartIdentity, detailText
        Case "E2"
            If Len(x__UTIL_ADC_REPORTS.ReportBuildTokenForE3(detailText)) = 0 Then RaiseSetupPartExitContractFailure completedPartIdentity, detailText
        Case "E3"
            If x__UTIL_ADC_REPORTS.ManagedReportsCurrentForSetup(detailText) = False Then RaiseSetupPartExitContractFailure completedPartIdentity, detailText
        Case "F1"
            ' ReconcileManagedObjectDependenciesForSetup already performs F1's immediate exact read-back proof against the same authoritative expected snapshot used for mutation.
            ' Do not rebuild live dependency evidence at the F1 exit boundary.
        Case "G1"
            If Len(RibbonDependencyTokenForSetup(detailText)) = 0 Then RaiseSetupPartExitContractFailure completedPartIdentity, detailText
        Case "H1"
            If RibbonSetupPlanReadyForH2(detailText) = False Then RaiseSetupPartExitContractFailure completedPartIdentity, detailText
        Case "H2"
            If RibbonSetupBuildReadyForH3(detailText) = False Then RaiseSetupPartExitContractFailure completedPartIdentity, detailText
        Case "H3"
            If Len(RibbonCommittedStateTokenForJ1(vbNullString, detailText)) = 0 Then RaiseSetupPartExitContractFailure completedPartIdentity, detailText
        Case "I1"
            If ConfiguredNavigationPaneGroupsCurrentForSetup(detailText, True) = False Then RaiseSetupPartExitContractFailure completedPartIdentity, detailText
        Case "I2"
            If SetupNavigationPanePlanReady(tokenText, rowCount) = False Then RaiseSetupPartExitContractFailure completedPartIdentity, "I2 did not retain the required same-session Navigation Pane plan."
        Case "I3"
            If Len(CompleteNavigationPanePlacementTokenForSetup(detailText, True)) = 0 Then RaiseSetupPartExitContractFailure completedPartIdentity, detailText
        Case "J1"
            'J1 performs the complete final read-only predecessor verification before Core stages final state.
        Case Else
            Err.Raise vbObjectError + 10162, "RequireSetupPartExitContract", "Unsupported Setup Part exit contract: " & completedPartIdentity
    End Select
    Exit Sub
A1ExitNotCurrent:
    detailText = "A1 exit verification failed. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    On Error GoTo 0
    RaiseSetupPartExitContractFailure completedPartIdentity, detailText
End Sub

Private Sub RequirePersistedSetupPartPrerequisites(ByVal requestedPartIdentity As String)
    Dim detailText As String
    Dim requestedRank As Long
    Dim catalogRequestedIdentity As String
    Dim planToken As String
    Dim planRowCount As Long
    requestedRank = ImplementedSetupPartRank(requestedPartIdentity)
    If requestedRank >= ImplementedSetupPartRank("A2") Then
        On Error GoTo A1NotCurrent
        ValidateConfiguredPhysicalNamesForSetup "RunSetupPartByIdentity." & requestedPartIdentity & ".PrerequisiteA1"
        On Error GoTo 0
    End If
    If requestedRank >= ImplementedSetupPartRank("A3") Then
        detailText = vbNullString
        If CoreDataTableFoundationCurrentForSetup(detailText) = False Then RaiseStandalonePrerequisiteFailure requestedPartIdentity, "A2", detailText
    End If
    If requestedRank >= ImplementedSetupPartRank("A4") Then
        detailText = vbNullString
        If NonFormQueryFoundationCurrentForSetup(detailText) = False Then RaiseStandalonePrerequisiteFailure requestedPartIdentity, "A3", detailText
        detailText = vbNullString
        If FormDesignQueryFoundationCurrentForSetup(detailText) = False Then RaiseStandalonePrerequisiteFailure requestedPartIdentity, "A3", detailText
    End If
    If requestedRank >= ImplementedSetupPartRank("B1") Then
        detailText = vbNullString
        If ReferenceInventoryFoundationCurrentForSetup(detailText) = False Then RaiseStandalonePrerequisiteFailure requestedPartIdentity, "A4", detailText
    End If
    If requestedRank >= ImplementedSetupPartRank("B2") Then
        If requestedRank > ImplementedSetupPartRank("C3") Then catalogRequestedIdentity = "C3" Else catalogRequestedIdentity = requestedPartIdentity
        detailText = vbNullString
        If CatalogPrerequisiteCurrentForSetup(catalogRequestedIdentity, detailText) = False Then RaiseStandalonePrerequisiteFailure requestedPartIdentity, IIf(requestedRank > ImplementedSetupPartRank("C3"), "B5", ImmediateCatalogPredecessorIdentity(requestedPartIdentity)), detailText
    End If
    If requestedRank >= ImplementedSetupPartRank("D1") Then
        detailText = vbNullString
        If PhaseC3CurrentForSetup(detailText) = False Then RaiseStandalonePrerequisiteFailure requestedPartIdentity, "C3", detailText
    End If
    If requestedPartIdentity = "D2" Or requestedPartIdentity = "D3" Then
        detailText = vbNullString
        If SetupGeneratedFormPreflightReadyForSetup(detailText) = False Then RaiseStandalonePrerequisiteFailure requestedPartIdentity, "D1", detailText
    End If
    If requestedPartIdentity = "D3" Then
        detailText = vbNullString
        If GeneratedFormPartitionCurrentForSetup("ChatGPTUIForms", detailText) = False Then RaiseStandalonePrerequisiteFailure requestedPartIdentity, "D2", detailText
    End If
    If requestedPartIdentity = "D4" Then
        detailText = vbNullString
        If GeneratedFormsCurrentForSetup(detailText) = False Then RaiseStandalonePrerequisiteFailure requestedPartIdentity, "D2/D3", detailText
    End If
    If requestedRank >= ImplementedSetupPartRank("E1") Then
        detailText = vbNullString
        If ManagedFormDomainCurrentForSetup(detailText) = False Then RaiseStandalonePrerequisiteFailure requestedPartIdentity, "D4", detailText
        detailText = vbNullString
        If D4LookupEditorBindingCurrentForSetup(detailText) = False Then RaiseStandalonePrerequisiteFailure requestedPartIdentity, "D4", detailText
        detailText = vbNullString
        If Len(CompleteCatalogStateTokenForSetup(detailText)) = 0 Then RaiseStandalonePrerequisiteFailure requestedPartIdentity, "D4", detailText
    End If
    If requestedPartIdentity = "E2" Then
        detailText = vbNullString
        If Len(x__UTIL_ADC_REPORTS.ReportPreflightTokenForE2(detailText)) = 0 Then RaiseStandalonePrerequisiteFailure requestedPartIdentity, "E1", detailText
    End If
    If requestedPartIdentity = "E3" Then
        detailText = vbNullString
        If Len(x__UTIL_ADC_REPORTS.ReportBuildTokenForE3(detailText)) = 0 Then RaiseStandalonePrerequisiteFailure requestedPartIdentity, "E2", detailText
    End If
    If requestedRank >= ImplementedSetupPartRank("F1") Then
        detailText = vbNullString
        If x__UTIL_ADC_REPORTS.ManagedReportsCurrentForSetup(detailText) = False Then RaiseStandalonePrerequisiteFailure requestedPartIdentity, "E3", detailText
    End If
    If requestedRank >= ImplementedSetupPartRank("G1") Then
        detailText = vbNullString
        If ManagedObjectDependenciesCurrentForSetup(detailText, planToken) = False Then RaiseStandalonePrerequisiteFailure requestedPartIdentity, "F1", detailText
    End If
    If requestedRank >= ImplementedSetupPartRank("H1") Then
        detailText = vbNullString
        If Len(RibbonDependencyTokenForSetup(detailText)) = 0 Then RaiseStandalonePrerequisiteFailure requestedPartIdentity, "G1", detailText
    End If
    If requestedPartIdentity = "H2" Then
        detailText = vbNullString
        If RibbonSetupPlanReadyForH2(detailText) = False Then RaiseStandalonePrerequisiteFailure requestedPartIdentity, "H1", detailText
    End If
    If requestedPartIdentity = "H3" Then
        detailText = vbNullString
        If RibbonSetupBuildReadyForH3(detailText) = False Then RaiseStandalonePrerequisiteFailure requestedPartIdentity, "H2", detailText
    End If
    If requestedRank >= ImplementedSetupPartRank("I1") Then
        detailText = vbNullString
        If Len(RibbonCommittedStateTokenForJ1(vbNullString, detailText)) = 0 Then RaiseStandalonePrerequisiteFailure requestedPartIdentity, "H3", detailText
    End If
    If requestedRank >= ImplementedSetupPartRank("I2") Then
        detailText = vbNullString
        If ConfiguredNavigationPaneGroupsCurrentForSetup(detailText, True) = False Then RaiseStandalonePrerequisiteFailure requestedPartIdentity, "I1", detailText
    End If
    If requestedPartIdentity = "I3" Then
        planToken = vbNullString
        planRowCount = 0
        If SetupNavigationPanePlanReady(planToken, planRowCount) = False Then RaiseStandalonePrerequisiteFailure requestedPartIdentity, "I2", "I3 requires the successful same-session I2 plan; I2 is never rerun implicitly."
    End If
    If requestedPartIdentity = "J1" Then
        detailText = vbNullString
        If Len(CompleteNavigationPanePlacementTokenForSetup(detailText, True)) = 0 Then RaiseStandalonePrerequisiteFailure requestedPartIdentity, "I3", detailText
    End If
    Exit Sub
A1NotCurrent:
    detailText = "A1 configuration/physical-name validation failed. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    On Error GoTo 0
    RaiseStandalonePrerequisiteFailure requestedPartIdentity, "A1", detailText
End Sub

Private Function ImmediateCatalogPredecessorIdentity(ByVal requestedPartIdentity As String) As String
    Select Case UCase$(Trim$(requestedPartIdentity))
        Case "B2": ImmediateCatalogPredecessorIdentity = "B1"
        Case "B3": ImmediateCatalogPredecessorIdentity = "B2"
        Case "B4": ImmediateCatalogPredecessorIdentity = "B3"
        Case "B5": ImmediateCatalogPredecessorIdentity = "B4"
        Case "C1", "C2", "C3": ImmediateCatalogPredecessorIdentity = "B5"
        Case Else: ImmediateCatalogPredecessorIdentity = "B1-B5"
    End Select
End Function

Private Sub RaiseStandalonePrerequisiteFailure(ByVal requestedPartIdentity As String, ByVal missingPartIdentity As String, ByVal detailText As String)
    Err.Raise vbObjectError + 9587, "RequirePersistedSetupPartPrerequisites", "Setup " & requestedPartIdentity & " refused to run because required predecessor " & missingPartIdentity & " is not current." & IIf(Len(Trim$(detailText)) = 0, vbNullString, vbCrLf & detailText) & vbCrLf & "No predecessor repair was executed automatically."
End Sub
