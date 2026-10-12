Attribute VB_Name = "x__UTIL_ADC_UI_RIBBON"
' ================================================================================================================================================================================
' ChatGPT Rewrite Tracking Header
' ================================================================================================================================================================================
' ChatGPT Previous Version ID: CGPT-5H3R-VVJ2-7HFB-MEPW-GLQ9
' ChatGPT Baseline Version ID: CGPT-5H3R-VVJ2-7HFB-MEPW-GLQ9
' ChatGPT Version ID: CGPT-9WVR-6K6W-9PR2-U9E8-AUG6
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
' STANDARD MODULE: Access Development Console Ribbon UI Utility
' VB_Name: x__UTIL_ADC_UI_RIBBON
' Purpose: Owns the ADC Ribbon H1-H3 plan/build/commit lifecycle, Ribbon persistence and verification, manifest validation, callbacks, tooltips, and terminal Ribbon teardown.
' ================================================================================================================================================================================

Option Compare Database
Option Explicit

' ===============================================================================================================================================================================
' MODULE-LEVEL DECLARATIONS
' ===============================================================================================================================================================================

' Ribbon table fields, callbacks, tabs, groups, limits, fingerprints, and runtime phase state.
Private Const RIBBON_NAME_ As String = "ADC"
Private Const RIBBON_PHASE_MODE_BUILD As String = "Build"
Private Const RIBBON_PHASE_MODE_VERIFY_ONLY As String = "VerifyOnly"
Private Const RIBBON_PHASE_MODE_SKIP As String = "Skip"
Private Const RIBBON_CUSTOM_UI_OPEN As String = "<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""RibbonOnLoad"">"
Private Const RIBBON_CALLBACK_OPEN_OBJECT As String = "RibbonOpenManagedObject"
Private Const RIBBON_CALLBACK_GET_SCREENTIP As String = "RibbonGetScreenTip"
Private Const RIBBON_CALLBACK_GET_SUPERTIP As String = "RibbonGetSuperTip"
Private Const RIBBON_CONTROL_TAG_STANDARD_PREFIX As String = "N"
Private Const RIBBON_TAG_DELIMITER As String = "|"
Private Const RIBBON_ACCESSUI_TAG_TERMINATOR As String = ";"
Private Const RIBBON_XML_INDENT_1 As String = "  "
Private Const RIBBON_XML_INDENT_2 As String = "    "
Private Const RIBBON_XML_INDENT_3 As String = "      "
Private Const RIBBON_XML_INDENT_4 As String = "        "
Private Const RIBBON_MACRO_IMAGE_MSO As String = "CreateMacro"
Private Const RIBBON_TABLE_IMAGE_MSO As String = "CreateTable"
Private Const RIBBON_QUERY_IMAGE_MSO As String = "CreateQueryInDesignView"
Private Const RIBBON_FORM_IMAGE_MSO As String = "CreateForm"
Private Const RIBBON_REPORT_IMAGE_MSO As String = "CreateReport"
Private Const RIBBON_RESTART_REQUIRED_PROPERTY As String = "ADCRibbonRestartRequired"
Private Const RIBBON_RESTART_REQUIRED_ON_PROPERTY As String = "ADCRibbonRestartRequiredOn"
Private Const RIBBON_LOADED_ON_PROPERTY As String = "ADCRibbonLoadedOn"
Private Const RIBBON_DEPENDENCY_TOKEN_PROPERTY As String = "ADCRibbonDependencyToken"
Private Const RIBBON_STANDARD_INPUT_FINGERPRINT_PROPERTY As String = "ADCRibbonStandardInputFingerprint"
Private Const RIBBON_STANDARD_XML_HASH_PROPERTY As String = "ADCRibbonStandardXmlHash"
Private Const RIBBON_STANDARD_TARGET_MAP_HASH_PROPERTY As String = "ADCRibbonStandardTargetMapHash"
Private Const RIBBON_STANDARD_BUTTON_COUNT_PROPERTY As String = "ADCRibbonStandardButtonCount"
Private Const RIBBON_FINGERPRINT_VERSION As String = "2026-08-14 system macro target captions v13"
Private Const RIBBON_MAX_BUTTONS_PER_GROUP As Long = 900
Private Const RIBBON_STANDARD_MAX_ID_ATTRIBUTE_LENGTH As Long = 45
Private Const RIBBON_SHORT_MAX_LABEL_LENGTH As Long = 48
Private Const RIBBON_SHORT_MAX_ID_ATTRIBUTE_LENGTH As Long = 48
Private Const RIBBON_MAX_BUILDER_TEXT_ATTRIBUTE_LENGTH As Long = 255
Private Const RIBBON_SKIP_REASON_DUPLICATE As String = "Duplicate"
Private Const RIBBON_SKIP_REASON_MISSING As String = "Missing"
Private Const RIBBON_SKIP_REASON_UNSUPPORTED As String = "Unsupported"
Private Const RIBBON_SKIP_REASON_BLANK_INVALID As String = "BlankInvalid"
' Loaded Ribbon callback reference, terminal teardown, and tooltip-error state.
Private mRibbonUI As IRibbonUI
' Becomes True only after Delete All has committed to terminal teardown; Ribbon callbacks remain inert for the rest of the session.
Private mTerminalTeardownActive As Boolean
Private mRibbonLastTooltipCallbackErrorKey As String
' INTENTIONAL RIBBON OPERATION CACHES / SNAPSHOT INDEXES:
' Caption-source and System Macro target dictionaries memoize metadata derived from the current catalog/MC_Sys state. The catalog source Collection is the coherent ordered Ribbon snapshot; exact/alias dictionaries
' are keyed indexes over those same rows. H1 initialization and source-currentness invalidation clear these caches before fresh reads; consumers must not mutate the module-owned caches.
Private mRibbonCaptionSourceCache As Object
Private mRibbonSystemMacroTargetCache As Object
Private mRibbonCatalogMetadataCache As Object
Private mRibbonCatalogAliasMetadataCache As Object
Private mRibbonCatalogSourceRows As Collection
Private mRibbonCatalogPreflightRowCount As Long
Private mRibbonCatalogSnapshotElapsedSeconds As Double
' INTENTIONAL FILTERED PROJECTION CACHE:
' Holds the configured Ribbon-group names resolved for the current Ribbon phase token so repeated consumers share the same ordered group projection. The Ribbon phase reset clears it; it is derived from canonical
' group definitions and is not independently authoritative.
Private mRibbonConfiguredGroupCache As Collection
' Ribbon target-insert batch state.
Private mRibbonTargetInsertRs As DAO.Recordset
' INTENTIONAL KEYED LOOKUP:
' During one Ribbon target-insert batch, maps each static button key to its target identity to reject conflicting duplicate ownership without rescanning staged rows. Begin/end batch owns the dictionary lifetime; it
' is transient validation state only.
Private mRibbonStaticTargetOwners As Object
' Ribbon progress-display state.
Public mRibbonProgressActive As Boolean
Private mRibbonProgressStep As Long
Private mRibbonProgressMax As Long
Private mRibbonProgressScopeLabel As String
Private mRibbonLastStatusText As String
' Active H1-H3 Ribbon phase transaction, plan, token, and timing state.
Public mRibbonPhasePrepared As Boolean
Private mRibbonPhaseToken As String
Private mRibbonPhaseStandardMode As String
Private mRibbonPhaseStandardCompleted As Boolean
Private mRibbonPhaseBuildRequired As Boolean
Private mRibbonPhaseStandardRequired As Boolean
Private mRibbonPhaseStandardBuilt As Boolean
Private mRibbonPhaseFastPathUsed As Boolean
Private mRibbonPhaseSkippedSameDatabase As Boolean
Private mRibbonPhaseShowMessage As Boolean
Private mRibbonPhaseApplyAfterBuild As Boolean
Private mRibbonPhaseApplyAfterBuildCaptured As Boolean
Private mRibbonPhaseTraceRunOwned As Boolean
Public mRibbonPhaseWorkspace As DAO.workspace
Public mRibbonPhaseTransactionStarted As Boolean
Private mRibbonPhaseMutationCommitted As Boolean
Private mRibbonPhaseRibbonXml As String
Private mRibbonPhaseDependencyToken As String
Private mRibbonPhaseCatalogToken As String
Private mRibbonPhasePredecessorToken As String
Private mRibbonPhaseGroupFingerprint As String
Private mRibbonPhaseStandardInputFingerprint As String
Private mRibbonPhaseReadinessDetail As String
Private mRibbonPhaseButtonCount As Long
Private mRibbonPhaseExpectedStandardCount As Long
' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION:
' Stores the prepared ADC Ribbon rows for the active Ribbon phase after preflight, preserving the checked deterministic row order through build/verification. The phase token/reset path bounds validity; rows are
' derived from canonical catalog/group sources and are not a second authority.
Private mRibbonPhaseStandardRows As Collection
Private mRibbonPhaseSourceDetail As String
Private mRibbonPhaseRowsPrepared As Boolean
Private mRibbonSourceSealActive As Boolean
Private mRibbonPhaseStandardChanged As Boolean
Private mRibbonPhaseRestartMarkerSaved As Boolean
Private mRibbonPhaseStartedTimer As Single
Private mRibbonPhaseReadinessElapsedSeconds As Double
Private mRibbonPhaseStandardElapsedSeconds As Double
Private mRibbonPhaseCommitElapsedSeconds As Double
' Same-run H3 proof identity. Mutable Ribbon source rows are always reread by later currentness consumers.
Private mRibbonSameRunH3ProofTraceRunID As String
Private mRibbonSameRunH3ProofHostDbName As String
Private mRibbonSameRunH3DependencyToken As String
Private mRibbonSameRunH3CatalogToken As String
Private mRibbonSameRunH3PredecessorToken As String

Private Const UI_RIBBON_POST_COMMIT_VERIFY_KEY As String = "ADC.UI.RibbonPostCommitVerifyV1"
Public Const UI_RIBBON_COMMIT_MARKER_KEY As String = "ADC.UI.RibbonCommitMarkerV1"
Private Const UI_RIBBON_COMMIT_STATUS_COMMIT_PENDING As String = "CommitPending"
Private Const UI_RIBBON_COMMIT_STATUS_COMMITTED As String = "Committed"
Private Const UI_RIBBON_COMMIT_STATUS_POST_COMMIT_COMPLETE As String = "PostCommitComplete"

Public Const ERR_RIBBON_ROLLBACK_FAILED As Long = vbObjectError + 7212

' ===============================================================================================================================================================================
' PROCEDURES
' ===============================================================================================================================================================================

' ===============================================================================================================================================================================
' RIBBON TRANSACTION RECOVERY / PERSISTENCE / MANIFEST VALIDATION
' ===============================================================================================================================================================================

' Purpose: Resolves retry ribbon transaction recovery from the supplied/current ADC state, using `ResolveUIRecoveryWorkspaceTransaction` and `VerifyRibbonPersistenceStorageAfterTransactionResolution` where required
'     by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: String: Returns the retry ribbon transaction recovery text value. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function RetryRibbonTransactionRecovery() As String
    On Error GoTo ErrHandler
    Dim transactionResolutionText As String
    Dim persistenceVerificationText As String
    transactionResolutionText = ResolveUIRecoveryWorkspaceTransaction("ribbon persistence recovery")
    persistenceVerificationText = VerifyRibbonPersistenceStorageAfterTransactionResolution()
    RetryRibbonTransactionRecovery = transactionResolutionText & vbCrLf & persistenceVerificationText & vbCrLf & _
        "The ribbon rollback recovery boundary is resolved. Normal setup can now repair missing Navigation Pane groups and independently rebuild or reverify the ribbon domain."
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, IIf(Len(Trim$(savedErrSource)) = 0, "RetryRibbonTransactionRecovery", savedErrSource), savedErrDescription
End Function

Private Function VerifyRibbonPersistenceStorageAfterTransactionResolution() As String
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim ribbonTableState As String
    Dim targetTableState As String
    Dim rowCount As Long
    Set db = GetHostDb()
    If SharedDao_DatabaseTableExistsSafe(db, DB_TABLE_USYS_RIBBONS) Then
        Set rs = db.OpenRecordset("SELECT Count(*) AS " & SQL_FIELD_ROW_COUNT & " FROM " & SharedDao_BracketName(DB_TABLE_USYS_RIBBONS) & ";", dbOpenSnapshot)
        rowCount = CLng(rs.fields(FIELD_ROW_COUNT).value)
        rs.Close
        Set rs = Nothing
        ribbonTableState = "present and readable; rows=" & CStr(rowCount)
    Else
        ribbonTableState = "absent; normal setup will create it before ribbon persistence"
    End If
    If SharedDao_DatabaseTableExistsSafe(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_RibbonButtonTargets")) Then
        Set rs = db.OpenRecordset("SELECT Count(*) AS " & SQL_FIELD_ROW_COUNT & " FROM " & SharedDao_BracketName(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_RibbonButtonTargets")) & ";", dbOpenSnapshot)
        rowCount = CLng(rs.fields(FIELD_ROW_COUNT).value)
        rs.Close
        Set rs = Nothing
        targetTableState = "present and readable; rows=" & CStr(rowCount)
    Else
        targetTableState = "absent; normal setup will create it before ribbon persistence"
    End If
    Set db = Nothing
    VerifyRibbonPersistenceStorageAfterTransactionResolution = "Ribbon persistence storage is readable after transaction resolution." & vbCrLf & "USysRibbons: " & ribbonTableState & vbCrLf & "UI_Ribbon_Generated_BtnTargets: " & targetTableState & vbCrLf & "Current ribbon-source freshness was intentionally deferred to the normal setup path so a repairable missing Navigation Pane group cannot keep the completed rollback boundary blocked."
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
    Set db = Nothing
    On Error GoTo 0
    Err.Raise savedErrNumber, "VerifyRibbonPersistenceStorageAfterTransactionResolution", "Ribbon persistence storage could not be read after transaction resolution." & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Function

Private Function RibbonManifestRowIsCanonicalSystemMacro(ByVal objectTypeCode As String, ByVal objectName As String, ByRef systemMacroDefinitions As Collection) As Boolean
    If StrComp(objectTypeCode, ADCAT_OBJECT_TYPE_MACRO, vbTextCompare) <> 0 Then Exit Function
    If systemMacroDefinitions Is Nothing Then Set systemMacroDefinitions = x__UTIL_ADC_DEFAULTS.GetSystemMacroDefinitions()
    RibbonManifestRowIsCanonicalSystemMacro = (Len(RibbonSystemMacroDesignatorDefault(objectName, systemMacroDefinitions)) > 0)
End Function

Private Function RibbonCatalogMetadataKey(ByVal objectName As String, ByVal objectTypeCode As String) As String
    RibbonCatalogMetadataKey = UCase$(NormalizedObjectTypeCode(objectTypeCode)) & "|" & UCase$(objectName)
End Function

Private Function RibbonTryExactCatalogMetadata(ByVal objectName As String, ByVal objectTypeCode As String, ByRef metadata As Variant) As Boolean
    Dim cacheKey As String
    If mRibbonCatalogMetadataCache Is Nothing Then Exit Function
    cacheKey = RibbonCatalogMetadataKey(objectName, objectTypeCode)
    If mRibbonCatalogMetadataCache.Exists(cacheKey) = False Then Exit Function
    metadata = mRibbonCatalogMetadataCache(cacheKey)
    RibbonTryExactCatalogMetadata = True
End Function

Private Function RibbonTryDisplayCatalogMetadata(ByVal objectName As String, ByVal objectTypeCode As String, ByRef metadata As Variant) As Boolean
    Dim cacheKey As String
    If RibbonTryExactCatalogMetadata(objectName, objectTypeCode, metadata) Then
        RibbonTryDisplayCatalogMetadata = True
        Exit Function
    End If
    If mRibbonCatalogAliasMetadataCache Is Nothing Then Exit Function
    cacheKey = RibbonCatalogMetadataKey(objectName, objectTypeCode)
    If mRibbonCatalogAliasMetadataCache.Exists(cacheKey) = False Then Exit Function
    metadata = mRibbonCatalogAliasMetadataCache(cacheKey)
    RibbonTryDisplayCatalogMetadata = True
End Function

Private Function RibbonAuthoritativeButtonTitle(ByVal objectTypeCode As String, ByVal objectName As String) As String
    'Ribbon button titles come from the stored managed-object RibbonTitle exactly; canonical System Macros may add their configured designator as the sole post-default title modification.
    Dim metadata As Variant
    If mRibbonCatalogMetadataCache Is Nothing Then PrepareRibbonCatalogSourceSnapshot
    If RibbonTryDisplayCatalogMetadata(objectName, objectTypeCode, metadata) = False Then Exit Function
    RibbonAuthoritativeButtonTitle = CStr(metadata(6))
    If StrComp(objectTypeCode, ADCAT_OBJECT_TYPE_MACRO, vbTextCompare) = 0 And CLng(metadata(2)) > 0 Then
        RibbonAuthoritativeButtonTitle = RibbonSystemMacroDisplayedTitleDefault(RibbonAuthoritativeButtonTitle, objectName)
    End If
End Function

Private Function RibbonCatalogFingerprintDiagnosticText(ByVal db As DAO.Database) As String
    On Error GoTo ErrHandler
    Dim propertyNames As Variant
    Dim propertyName As Variant
    Dim resultText As String
    propertyNames = CatalogDiagnosticPropertyNames()
    resultText = vbCrLf & "Catalog database-property snapshot:"
    For Each propertyName In propertyNames
        resultText = resultText & vbCrLf & "  " & CStr(propertyName) & "=" & RibbonDatabasePropertyDiagnosticText(db, CStr(propertyName))
    Next propertyName
    RibbonCatalogFingerprintDiagnosticText = resultText
    Exit Function
ErrHandler:
    RibbonCatalogFingerprintDiagnosticText = vbCrLf & "Catalog database-property snapshot failed. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
End Function

Private Function RequiredSystemManifestCountSummary(ByVal systemGroupCounts As Object, ByVal misplacedCount As Long, ByVal unexpectedCount As Long, ByVal systemRibbonGroupDefinitions As Collection) As String
    Dim Definition As Variant
    Dim groupIdentity As String
    Dim tabKey As String
    Dim tabLabel As String
    Dim groupLabel As String
    Dim groupOrder As Long
    Dim detailText As String
    For Each Definition In systemRibbonGroupDefinitions
        groupIdentity = CStr(Definition(0))
        groupLabel = CStr(Definition(1))
        groupOrder = CLng(Definition(2))
        tabKey = CStr(Definition(3))
        tabLabel = GetRibbonTabLabelDefault(tabKey)
        If Len(detailText) > 0 Then detailText = detailText & vbCrLf
        detailText = detailText & tabLabel & " > " & groupLabel & " (" & CStr(groupOrder) & ") rows: " & CStr(SystemManifestGroupCount(systemGroupCounts, groupIdentity))
    Next Definition
    If Len(detailText) > 0 Then detailText = detailText & vbCrLf
    detailText = detailText & "Recognized System Ribbon group rows on wrong tabs: " & CStr(misplacedCount)
    detailText = detailText & vbCrLf & "Unexpected groups on ADC1/ADC2: " & CStr(unexpectedCount)
    RequiredSystemManifestCountSummary = detailText
End Function

Private Sub AppendSystemManifestLayoutDiagnostic(ByRef diagnosticText As String, ByRef diagnosticRowCount As Long, ByRef diagnosticTruncated As Boolean, ByVal issueText As String, ByVal tabKey As String, ByVal tabLabel As String, ByVal groupLabel As String, ByVal buttonLabel As String, ByVal objectTypeCode As String, ByVal objectName As String, ByVal expectedTab As String)
    Const MAX_DIAGNOSTIC_ROWS As Long = 25
    If diagnosticRowCount >= MAX_DIAGNOSTIC_ROWS Then
        diagnosticTruncated = True
        Exit Sub
    End If
    diagnosticRowCount = diagnosticRowCount + 1
    diagnosticText = diagnosticText & IIf(Len(diagnosticText) = 0, vbNullString, vbCrLf) & " - Issue=" & issueText & "; ActualTabKey=" & IIf(Len(tabKey) = 0, "(blank)", tabKey) & "; ActualTabLabel=" & IIf(Len(tabLabel) = 0, "(blank)", tabLabel) & "; Group=" & IIf(Len(groupLabel) = 0, "(blank)", groupLabel) & "; ExpectedTab=" & expectedTab & "; Button=" & IIf(Len(buttonLabel) = 0, "(blank)", buttonLabel) & "; ObjectType=" & IIf(Len(objectTypeCode) = 0, "(blank)", objectTypeCode) & "; Object=" & IIf(Len(Trim$(objectName)) = 0, "(blank)", objectName)
End Sub

Private Function SystemManifestGroupCount(ByVal systemGroupCounts As Object, ByVal groupIdentity As String) As Long
    If systemGroupCounts Is Nothing Then Exit Function
    If Len(groupIdentity) = 0 Then Exit Function
    If systemGroupCounts.Exists(groupIdentity) Then SystemManifestGroupCount = CLng(systemGroupCounts(groupIdentity))
End Function

Private Sub CountRequiredSystemManifestGroup(ByVal tabKey As String, ByVal tabLabel As String, ByVal groupLabel As String, ByVal buttonLabel As String, ByVal objectTypeCode As String, ByVal objectName As String, ByVal systemGroupCounts As Object, ByVal systemGroupIdentityByPlacement As Object, ByVal systemExpectedTabsByVisibleGroup As Object, ByRef systemMacroDefinitions As Collection, ByRef unexpectedCount As Long, ByRef misplacedCount As Long, ByRef diagnosticText As String, ByRef diagnosticRowCount As Long, ByRef diagnosticTruncated As Boolean)
    'Counts final visible System Ribbon groups and diagnoses layout only for authoritative built-in System Macro rows.
    Dim normalizedTab As String
    Dim groupIdentity As String
    Dim expectedTabs As String
    Dim visibleGroupKey As String
    Dim placementKey As String
    Dim canonicalSystemMacroRow As Boolean
    normalizedTab = UCase$(tabKey)
    visibleGroupKey = RibbonSystemManifestVisibleGroupKey(groupLabel)
    placementKey = normalizedTab & "|" & visibleGroupKey
    If systemGroupIdentityByPlacement.Exists(placementKey) Then groupIdentity = CStr(systemGroupIdentityByPlacement(placementKey))
    If Len(groupIdentity) > 0 Then
        If systemGroupCounts.Exists(groupIdentity) = False Then systemGroupCounts.Add groupIdentity, CLng(0)
        systemGroupCounts(groupIdentity) = CLng(systemGroupCounts(groupIdentity)) + 1
        Exit Sub
    End If
    canonicalSystemMacroRow = RibbonManifestRowIsCanonicalSystemMacro(objectTypeCode, objectName, systemMacroDefinitions)
    If canonicalSystemMacroRow = False Then Exit Sub
    If systemExpectedTabsByVisibleGroup.Exists(visibleGroupKey) Then
        misplacedCount = misplacedCount + 1
        expectedTabs = CStr(systemExpectedTabsByVisibleGroup(visibleGroupKey))
        AppendSystemManifestLayoutDiagnostic diagnosticText, diagnosticRowCount, diagnosticTruncated, "Recognized System Ribbon group on wrong tab", tabKey, tabLabel, groupLabel, buttonLabel, objectTypeCode, objectName, expectedTabs
        Exit Sub
    End If
    If normalizedTab = DEFAULT_RIBBON_TAB_KEY_SYSTEM_1 Or normalizedTab = DEFAULT_RIBBON_TAB_KEY_SYSTEM_2 Then
        unexpectedCount = unexpectedCount + 1
        AppendSystemManifestLayoutDiagnostic diagnosticText, diagnosticRowCount, diagnosticTruncated, "Unexpected System Ribbon group", tabKey, tabLabel, groupLabel, buttonLabel, objectTypeCode, objectName, "(no canonical System Ribbon group)"
    End If
End Sub

' ===============================================================================================================================================================================
' RIBBON TEXT / LABEL / XML HELPERS
' ===============================================================================================================================================================================

' Purpose: Resolves ribbon XML attribute text from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   valueText (Input; String): Input value text examined or transformed by this procedure according to its documented ribbon XML attribute text contract.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function RibbonXmlAttributeText(ByVal valueText As String) As String
    'Escapes XML-special characters so text can be embedded safely in a Ribbon XML attribute.
    valueText = Replace(valueText, "&", "&amp;")
    valueText = Replace(valueText, "<", "&lt;")
    valueText = Replace(valueText, ">", "&gt;")
    valueText = Replace(valueText, Chr$(34), "&quot;")
    valueText = Replace(valueText, "'", "&apos;")
    RibbonXmlAttributeText = valueText
End Function

' Purpose: Resolves and returns ribbon limited error text from the supplied inputs and the authoritative ADC definitions/state used by this module.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   valueText (Input; String): Input value text examined or transformed by this procedure according to its documented ribbon limited error text contract.
'   maxLength (Input; Long): Upper bound limiting the max length quantity scanned, displayed, or processed by this call.
' Returns: String: Returns the resolved ribbon limited error text text in the canonical representation consumed by this module's caller; vbNullString represents no applicable value where the existing contract
'   permits it.
Public Function RibbonLimitedErrorText(ByVal valueText As String, ByVal maxLength As Long) As String
    'Flattens line breaks and truncates error text to the requested maximum length with an ellipsis.
    valueText = Replace(valueText, vbCrLf, " ")
    valueText = Replace(valueText, vbCr, " ")
    valueText = Replace(valueText, vbLf, " ")
    If maxLength <= 0 Then maxLength = 500
    If Len(valueText) > maxLength Then valueText = Left$(valueText, maxLength - 3) & "..."
    RibbonLimitedErrorText = valueText
End Function

' Purpose: Resolves ribbon group name key from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   groupName (Input; String): Canonical or physical name identifying the group targeted by this operation.
' Returns: String: Returns the resolved name produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function RibbonGroupNameKey(ByVal groupName As String) As String
    'Returns an uppercase trimmed key for case-insensitive Ribbon group-name comparison.
    RibbonGroupNameKey = UCase$(Trim$(groupName))
End Function

' Purpose: Resolves ribbon button label key from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   labelText (Input; String): Input label text for ribbon button label key; is passed to `RibbonButtonLabelKey`.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function RibbonButtonLabelKey(ByVal labelText As String) As String
    'Returns an uppercase key for case-insensitive Ribbon button-label comparison.
    RibbonButtonLabelKey = UCase$(labelText)
End Function

' Purpose: Resolves and returns ribbon limited caption text from the supplied inputs and the authoritative ADC definitions/state used by this module.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   captionText (Input; String): Caption text text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine.
'   maxLength (Input; Long): Upper bound limiting the max length quantity scanned, displayed, or processed by this call.
' Returns: String: Returns the canonical display text for ribbon limited caption text after applying the title/caption/label rules owned by this module; vbNullString represents no applicable display text where
'   supported.
Public Function RibbonLimitedCaptionText(ByVal captionText As String, ByVal maxLength As Long) As String
    'Truncates caption text to the requested maximum length and adds an ellipsis when truncation is required.
    If maxLength <= 3 Or Len(captionText) <= maxLength Then
        RibbonLimitedCaptionText = captionText
    Else
        RibbonLimitedCaptionText = Left$(captionText, maxLength - 3) & "..."
    End If
End Function

' Purpose: Resolves ribbon trim label from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   labelText (Input; String): Input label text for ribbon trim label; is tested by If Len(labelText) > DEFAULT_RIBBON_TITLE_MAX_LENGTH; is assigned to `RibbonTrimLabel`.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function RibbonTrimLabel(ByVal labelText As String) As String
    'Trims a Ribbon label and bounds it to the configured maximum title length.
    labelText = Trim$(labelText)
    If Len(labelText) > DEFAULT_RIBBON_TITLE_MAX_LENGTH Then labelText = Left$(labelText, DEFAULT_RIBBON_TITLE_MAX_LENGTH - 3) & "..."
    RibbonTrimLabel = labelText
End Function

' ===============================================================================================================================================================================
' RIBBON SOURCE DISCOVERY / SQL PARSING HELPERS
' ===============================================================================================================================================================================

' Purpose: Returns True when a SQL token is one of the reserved source-parsing keywords.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   tokenText (Input; String): Input token text for ribbon SQL source token is keyword; is normalized to upper case before use.
' Returns: Boolean: True when a SQL token is one of the reserved source-parsing keywords. False otherwise.
Public Function RibbonSqlSourceTokenIsKeyword(ByVal tokenText As String) As Boolean
    'Returns True when a SQL token is one of the reserved source-parsing keywords.
    Select Case UCase$(tokenText)
        Case "SELECT", "DISTINCT", "DISTINCTROW", "TOP", "ALL", "FROM", "JOIN", "INNER", "LEFT", "RIGHT", "FULL", "OUTER", "CROSS", "WHERE", "GROUP", "ORDER", "HAVING", "UNION", "PIVOT", "TRANSFORM", "PARAMETERS", "AS", "ON"
            RibbonSqlSourceTokenIsKeyword = True
    End Select
End Function

' Purpose: Returns True when text begins like a supported Access SQL statement.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   sourceText (Input; String): Input source text for ribbon text looks like SQL; is passed to `normalizedText`.
' Returns: Boolean: True when text begins like a supported Access SQL statement. False otherwise.
Public Function RibbonTextLooksLikeSql(ByVal sourceText As String) As Boolean
    'Returns True when text begins like a supported Access SQL statement.
    Dim normalizedText As String
    normalizedText = UCase$(sourceText)
    RibbonTextLooksLikeSql = (Left$(normalizedText, 6) = "SELECT" Or Left$(normalizedText, 9) = "TRANSFORM" Or Left$(normalizedText, 10) = "PARAMETERS")
End Function

Private Sub RibbonAddQuerySourceObjects(ByRef sourceNames As Object, ByRef orderedNames As Collection, ByVal queryName As String)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim qdf As DAO.queryDef
    Dim fld As DAO.Field
    Dim sourceName As String
    Set db = GetHostDb()
    Set qdf = db.QueryDefs(queryName)
    RibbonAddSqlSourceNames sourceNames, orderedNames, CStr(qdf.sql), queryName
    If orderedNames.count > 0 Then Exit Sub
    For Each fld In qdf.fields
        RibbonSetupControlPump
        sourceName = vbNullString
        On Error Resume Next
        sourceName = Trim$(CStr(fld.properties("SourceTable").value))
        On Error GoTo ErrHandler
        RibbonRememberSourceName sourceNames, orderedNames, sourceName, queryName
    Next fld
    Exit Sub
ErrHandler:
    On Error GoTo 0
End Sub

Private Sub RibbonAddRecordSourceObjects(ByRef sourceNames As Object, ByRef orderedNames As Collection, ByVal recordSourceText As String, ByVal ownerObjectName As String)
    recordSourceText = Trim$(recordSourceText)
    If Len(recordSourceText) = 0 Then Exit Sub
    If RibbonTextLooksLikeSql(recordSourceText) Then
        RibbonAddSqlSourceNames sourceNames, orderedNames, recordSourceText, ownerObjectName
    Else
        RibbonRememberSourceName sourceNames, orderedNames, recordSourceText, ownerObjectName
    End If
End Sub

Private Sub RibbonAddSqlSourceNames(ByRef sourceNames As Object, ByRef orderedNames As Collection, ByVal sqlText As String, ByVal ownerObjectName As String)
    Dim tokens As Collection
    Dim tokenIndex As Long
    Dim tokenText As String
    Dim tokenUpper As String
    Dim nextToken As String
    Dim expectSource As Boolean
    Dim inFromClause As Boolean
    Set tokens = RibbonSqlTokens(sqlText)
    For tokenIndex = 1 To tokens.count
        RibbonSetupControlPump
        tokenText = CStr(tokens(tokenIndex))
        tokenUpper = UCase$(tokenText)
        Select Case tokenUpper
            Case "FROM", "JOIN"
                expectSource = True
                inFromClause = True
            Case "WHERE", "GROUP", "ORDER", "HAVING", "UNION", "PIVOT", "TRANSFORM", "WITH", ";"
                inFromClause = False
                expectSource = False
            Case ","
                If inFromClause Then expectSource = True
            Case "("
                If expectSource And tokenIndex < tokens.count Then
                    nextToken = UCase$(CStr(tokens(tokenIndex + 1)))
                    If nextToken = "SELECT" Or nextToken = "TRANSFORM" Or nextToken = "PARAMETERS" Then expectSource = False
                End If
            Case ")", "INNER", "LEFT", "RIGHT", "FULL", "OUTER", "CROSS", "AS", "ON"
            Case Else
                If expectSource Then
                    If tokenIndex + 2 <= tokens.count And CStr(tokens(tokenIndex + 1)) = "." Then tokenText = CStr(tokens(tokenIndex + 2))
                    RibbonRememberSourceName sourceNames, orderedNames, tokenText, ownerObjectName
                    expectSource = False
                End If
        End Select
    Next tokenIndex
End Sub

Private Function RibbonCatalogFormSource(ByVal formName As String) As String
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Set db = GetHostDb()
    If SharedDao_TableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_FormDefinitions")) = False Then Exit Function
    If SharedDao_TableFieldExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_FormDefinitions"), FIELD_FORM_OBJECT_NAME) = False Or SharedDao_TableFieldExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_FormDefinitions"), FIELD_MANAGED_OBJECT_FIELD_SOURCE_OBJECT_NAME) = False Then Exit Function
    sqlText = "SELECT TOP 1 " & SQL_FIELD_MANAGED_OBJECT_FIELD_SOURCE_OBJECT_NAME & " FROM " & SharedDao_BracketName(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_FormDefinitions")) & " WHERE " & SQL_FIELD_FORM_OBJECT_NAME & "=" & SharedDao_SqlQuote(formName)
    If SharedDao_TableFieldExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_FormDefinitions"), FIELD_FORM_STATUS) Then sqlText = sqlText & " AND Nz(" & SQL_FIELD_FORM_STATUS & ",'Active')='Active'"
    sqlText = sqlText & " ORDER BY " & OrderByPolicyText("UI_FORM_CATALOG_SOURCE_MATCH") & ";"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    If rs.EOF = False Then
        RibbonCatalogFormSource = Trim$(CStr(Nz(rs.fields(FIELD_MANAGED_OBJECT_FIELD_SOURCE_OBJECT_NAME).value, vbNullString)))
        If StrComp(RibbonCatalogFormSource, formName, vbTextCompare) = 0 Then RibbonCatalogFormSource = vbNullString
    End If
    rs.Close
    Exit Function
ErrHandler:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
End Function

Private Function RibbonFormRecordSource(ByVal formName As String) As String
    On Error GoTo ErrHandler
    Dim openedByHelper As Boolean
    Dim wasLoaded As Boolean
    RibbonFormRecordSource = RibbonCatalogFormSource(formName)
    If Len(RibbonFormRecordSource) > 0 Then Exit Function
    wasLoaded = CurrentProject.AllForms(formName).IsLoaded
    If wasLoaded = False Then DoCmd.OpenForm formName, acDesign, , , , acHidden: openedByHelper = True
    RibbonFormRecordSource = Trim$(CStr(Nz(forms(formName).recordSource, vbNullString)))
    If openedByHelper Then DoCmd.Close acForm, formName, acSaveNo
    Exit Function
ErrHandler:
    On Error Resume Next
    If openedByHelper Then DoCmd.Close acForm, formName, acSaveNo
    On Error GoTo 0
End Function

Private Function RibbonObjectSourceSummary(ByVal objectTypeCode As String, ByVal objectName As String, ByRef sourceCount As Long) As String
    Dim cacheKey As String
    Dim cachedValue As Variant
    Dim sourceNames As Object
    Dim orderedNames As Collection
    cacheKey = UCase$(objectTypeCode) & RIBBON_TAG_DELIMITER & UCase$(objectName)
    If mRibbonCaptionSourceCache Is Nothing Then Set mRibbonCaptionSourceCache = CreateObject("Scripting.Dictionary"): mRibbonCaptionSourceCache.CompareMode = vbTextCompare
    If mRibbonCaptionSourceCache.Exists(cacheKey) Then
        cachedValue = mRibbonCaptionSourceCache(cacheKey)
        sourceCount = CLng(cachedValue(0))
        RibbonObjectSourceSummary = CStr(cachedValue(1))
        Exit Function
    End If
    ' INTENTIONAL KEYED LOOKUP: sourceNames provides keyed membership/lookup over source rows or state already read by RibbonObjectSourceSummary; it is a derived access path, not a second authority.
    ' INTENTIONAL ORDERED VIEW: orderedNames preserves the deterministic sequence required by RibbonObjectSourceSummary; keyed-only or unordered storage would not preserve the consumer semantics.
    Set sourceNames = CreateObject("Scripting.Dictionary")
    sourceNames.CompareMode = vbTextCompare
    Set orderedNames = New Collection
    Select Case UCase$(objectTypeCode)
        Case UCase$(ADCAT_OBJECT_TYPE_FORM)
            RibbonAddRecordSourceObjects sourceNames, orderedNames, RibbonFormRecordSource(objectName), objectName
        Case UCase$(ADCAT_OBJECT_TYPE_REPORT)
            RibbonAddRecordSourceObjects sourceNames, orderedNames, RibbonReportRecordSource(objectName), objectName
        Case UCase$(ADCAT_OBJECT_TYPE_QUERY)
            RibbonAddQuerySourceObjects sourceNames, orderedNames, objectName
    End Select
    sourceCount = orderedNames.count
    RibbonObjectSourceSummary = SharedCollection_Join(orderedNames, ", ")
    mRibbonCaptionSourceCache.Add cacheKey, Array(sourceCount, RibbonObjectSourceSummary)
End Function

Private Sub RibbonRememberSourceName(ByRef sourceNames As Object, ByRef orderedNames As Collection, ByVal sourceName As String, ByVal ownerObjectName As String)
    Dim normalizedName As String
    sourceName = Trim$(sourceName)
    If Len(sourceName) = 0 Then Exit Sub
    If Left$(sourceName, 1) = "[" And Right$(sourceName, 1) = "]" Then sourceName = Mid$(sourceName, 2, Len(sourceName) - 2)
    sourceName = Trim$(sourceName)
    If Len(sourceName) = 0 Then Exit Sub
    If StrComp(sourceName, ownerObjectName, vbTextCompare) = 0 Then Exit Sub
    If RibbonSqlSourceTokenIsKeyword(sourceName) Then Exit Sub
    normalizedName = UCase$(sourceName)
    If sourceNames.Exists(normalizedName) Then Exit Sub
    sourceNames.Add normalizedName, True
    orderedNames.Add sourceName
End Sub

Private Function RibbonReportRecordSource(ByVal reportName As String) As String
    On Error GoTo ErrHandler
    Dim openedByHelper As Boolean
    Dim wasLoaded As Boolean
    wasLoaded = CurrentProject.AllReports(reportName).IsLoaded
    If wasLoaded = False Then DoCmd.OpenReport reportName, acViewDesign, , , acHidden: openedByHelper = True
    RibbonReportRecordSource = Trim$(CStr(Nz(reports(reportName).recordSource, vbNullString)))
    If openedByHelper Then DoCmd.Close acReport, reportName, acSaveNo
    Exit Function
ErrHandler:
    On Error Resume Next
    If openedByHelper Then DoCmd.Close acReport, reportName, acSaveNo
    On Error GoTo 0
End Function


Private Function RibbonSqlTokens(ByVal sqlText As String) As Collection
    ' INTENTIONAL ORDERED VIEW: tokens preserves the deterministic sequence required by RibbonSqlTokens; keyed-only or unordered storage would not preserve the consumer semantics.
    Dim tokens As New Collection
    Dim tokenText As String
    Dim characterText As String
    Dim nextCharacter As String
    Dim indexValue As Long
    Dim inString As Boolean
    Dim inBracket As Boolean
    For indexValue = 1 To Len(sqlText)
        If (indexValue And 255) = 0 Then RibbonSetupControlPump
        characterText = Mid$(sqlText, indexValue, 1)
        If inString Then
            If characterText = """" Then
                If indexValue < Len(sqlText) Then nextCharacter = Mid$(sqlText, indexValue + 1, 1) Else nextCharacter = vbNullString
                If nextCharacter = """" Then indexValue = indexValue + 1 Else inString = False
            End If
        ElseIf inBracket Then
            If characterText = "]" Then
                If indexValue < Len(sqlText) Then nextCharacter = Mid$(sqlText, indexValue + 1, 1) Else nextCharacter = vbNullString
                If nextCharacter = "]" Then tokenText = tokenText & "]": indexValue = indexValue + 1 Else inBracket = False: If Len(tokenText) > 0 Then tokens.Add tokenText: tokenText = vbNullString
            Else
                tokenText = tokenText & characterText
            End If
        Else
            Select Case characterText
                Case """"
                    If Len(tokenText) > 0 Then tokens.Add tokenText: tokenText = vbNullString
                    inString = True
                Case "["
                    If Len(tokenText) > 0 Then tokens.Add tokenText: tokenText = vbNullString
                    inBracket = True
                Case " ", vbTab, vbCr, vbLf
                    If Len(tokenText) > 0 Then tokens.Add tokenText: tokenText = vbNullString
                Case "(", ")", ",", ".", ";"
                    If Len(tokenText) > 0 Then tokens.Add tokenText: tokenText = vbNullString
                    tokens.Add characterText
                Case Else
                    tokenText = tokenText & characterText
            End Select
        End If
    Next indexValue
    If Len(tokenText) > 0 Then tokens.Add tokenText
    Set RibbonSqlTokens = tokens
End Function

' ===============================================================================================================================================================================
' RIBBON STATUS / TRANSACTION RECOVERY / PERSISTENCE / MANIFEST VALIDATION
' ===============================================================================================================================================================================

Private Function RibbonObjectExists(ByVal objectTypeCode As String, ByVal objectName As String) As Boolean
    On Error GoTo ErrHandler
    Dim hostDb As DAO.Database
    Dim normalizedType As String
    normalizedType = NormalizedObjectTypeCode(objectTypeCode)
    If Len(normalizedType) = 0 Or Len(objectName) = 0 Then Exit Function
    Set hostDb = GetHostDb()
    Select Case UCase$(normalizedType)
        Case UCase$(ADCAT_OBJECT_TYPE_MACRO)
            RibbonObjectExists = SharedDao_AccessObjectExists(hostDb, acMacro, objectName)
        Case UCase$(ADCAT_OBJECT_TYPE_TABLE)
            RibbonObjectExists = SharedDao_TableExists(hostDb, objectName)
        Case UCase$(ADCAT_OBJECT_TYPE_QUERY)
            RibbonObjectExists = SharedDao_QueryExists(hostDb, objectName)
        Case UCase$(ADCAT_OBJECT_TYPE_FORM)
            RibbonObjectExists = SharedDao_AccessObjectExists(hostDb, acForm, objectName)
        Case UCase$(ADCAT_OBJECT_TYPE_REPORT)
            RibbonObjectExists = SharedDao_AccessObjectExists(hostDb, acReport, objectName)
    End Select
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "RibbonObjectExists", "Could not verify the physical managed object." & vbCrLf & "Object: " & objectName & vbCrLf & "Expected type: " & objectTypeCode & vbCrLf & "Original Err.Source: " & savedErrSource & vbCrLf & "Original Err.Description: " & savedErrDescription
End Function

Private Function RibbonAllowedImageMso(ByVal imageName As String) As Boolean
    Select Case imageName
        Case RIBBON_MACRO_IMAGE_MSO, RIBBON_TABLE_IMAGE_MSO, RIBBON_QUERY_IMAGE_MSO, RIBBON_FORM_IMAGE_MSO, RIBBON_REPORT_IMAGE_MSO
            RibbonAllowedImageMso = True
    End Select
End Function

Private Function RibbonStatusPartPrefixText() As String
    Dim scopeText As String
    scopeText = mRibbonProgressScopeLabel
    If Len(scopeText) > 0 Then RibbonStatusPartPrefixText = scopeText & " "
End Function

' Purpose: Resolves and returns ribbon status part text from the supplied inputs and the authoritative ADC definitions/state used by this module.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   statusText (Input; String): Status/progress message supplied for the current operation; the routine forwards it to the applicable UI/progress and trace/reporting path.
'   advancePart (Input Optional; Boolean): True advances the tracked Setup part after successful completion; False performs the work without advancing the Setup-part state. Default: True.
' Returns: String: Returns the resolved ribbon status part text text in the canonical representation consumed by this module's caller; vbNullString represents no applicable value where the existing contract
'   permits it.
Public Function RibbonStatusPartText(ByVal statusText As String, Optional ByVal advancePart As Boolean = True) As String
    Dim partNumber As Long
    If Len(statusText) = 0 Then Exit Function
    If mRibbonProgressActive = False Or mRibbonProgressMax <= 0 Then
        RibbonStatusPartText = statusText
        Exit Function
    End If
    partNumber = mRibbonProgressStep
    If advancePart Then partNumber = partNumber + 1
    If partNumber < 1 Then partNumber = 1
    If partNumber > mRibbonProgressMax Then partNumber = mRibbonProgressMax
    RibbonStatusPartText = RibbonStatusPartPrefixText() & "Part " & CStr(partNumber) & " of " & CStr(mRibbonProgressMax) & " > " & statusText
    If advancePart Then mRibbonProgressStep = partNumber
End Function


Private Function RibbonTextOccurrenceCount(ByVal sourceText As String, ByVal searchText As String) As Long
    Dim searchPosition As Long
    If Len(searchText) = 0 Then Exit Function
    searchPosition = 1
    Do
        RibbonSetupControlPump
        searchPosition = InStr(searchPosition, sourceText, searchText, vbTextCompare)
        If searchPosition = 0 Then Exit Do
        RibbonTextOccurrenceCount = RibbonTextOccurrenceCount + 1
        searchPosition = searchPosition + Len(searchText)
    Loop
End Function

Private Function RibbonUiManifestGroupDiagnostic(ByVal ribbonName As String) As String
    'Returns the persisted ADC UI tab group inventory for required-group validation diagnostics.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim detailText As String
    Set db = GetHostDb()
    sqlText = "SELECT [" & FIELD_RIBBON_TAB_KEY & "], [" & FIELD_RIBBON_TAB_LABEL & "], [" & FIELD_RIBBON_GROUP_LABEL & "], Count(*) AS " & SQL_FIELD_ROW_COUNT & " FROM " & SharedDao_BracketName(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_RIBBON_BUTTON_TARGETS_TABLE_LOGICAL_ID)) & " WHERE [" & FIELD_RIBBON_NAME & "]=" & SharedDao_SqlQuote(ribbonName) & " AND [" & FIELD_RIBBON_TAB_KEY & "]=" & SharedDao_SqlQuote(DEFAULT_RIBBON_TAB_KEY_UI) & " GROUP BY [" & FIELD_RIBBON_TAB_KEY & "], [" & FIELD_RIBBON_TAB_LABEL & "], [" & FIELD_RIBBON_GROUP_LABEL & "] ORDER BY " & OrderByPolicyText("UI_RIBBON_GROUP_DIAGNOSTIC") & ";"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    Do While rs.EOF = False
        If Len(detailText) > 0 Then detailText = detailText & vbCrLf
        detailText = detailText & "Tab key: " & CStr(rs.fields(FIELD_RIBBON_TAB_KEY)) & "; tab label: " & CStr(Nz(rs.fields(FIELD_RIBBON_TAB_LABEL), vbNullString)) & "; group: " & CStr(Nz(rs.fields(FIELD_RIBBON_GROUP_LABEL), vbNullString)) & "; rows: " & CStr(rs!rowCount)
        rs.MoveNext
    Loop
    If Len(detailText) = 0 Then detailText = "(No persisted rows were found for tab key " & DEFAULT_RIBBON_TAB_KEY_UI & ".)"
    RibbonUiManifestGroupDiagnostic = detailText
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    On Error GoTo 0
    Exit Function
ErrHandler:
    RibbonUiManifestGroupDiagnostic = "(Could not collect " & DEFAULT_UTILITY_SHORT_NAME & " UI group inventory. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description & ")"
    Resume CleanExit
End Function

Private Function StoredRibbonXml(ByVal db As DAO.Database, ByVal ribbonName As String, ByRef ribbonExists As Boolean) As String
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim matchingRowCount As Long
    ribbonExists = False
    sqlText = "SELECT " & SharedDao_BracketName(FIELD_RIBBON_XML) & " FROM " & SharedDao_BracketName(DB_TABLE_USYS_RIBBONS) & " WHERE " & SharedDao_BracketName(FIELD_RIBBON_NAME) & "=" & SharedDao_SqlQuote(ribbonName) & ";"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    If rs.EOF = False Then
        rs.MoveLast
        matchingRowCount = rs.RecordCount
        If matchingRowCount <> 1 Then Err.Raise vbObjectError + 7278, "StoredRibbonXml", "USysRibbons contains more than one row for the same CustomRibbonID. The Ribbon must be rebuilt so H2 can collapse the duplicate rows to one authoritative stored XML payload." & vbCrLf & "Ribbon: " & ribbonName & vbCrLf & "Rows: " & CStr(matchingRowCount)
        rs.MoveFirst
        ribbonExists = True
        StoredRibbonXml = CStr(Nz(rs.fields(FIELD_RIBBON_XML).value, vbNullString))
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
    Err.Raise savedErrNumber, "StoredRibbonXml", "Could not read stored " & DEFAULT_UTILITY_LONG_NAME & " ribbon XML." & vbCrLf & "Ribbon: " & ribbonName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function UISerializeRibbonCommitPayload(ByVal commitStatus As String, ByVal ribbonXmlChanged As Boolean, ByVal applyAfterBuild As Boolean) As String
    UISerializeRibbonCommitPayload = "Version=4" & vbCrLf & _
        "CommitStatus=" & UIDurableEncode(commitStatus) & vbCrLf & _
        "PreparationToken=" & UIDurableEncode(mRibbonPhaseToken) & vbCrLf & _
        "StandardMode=" & UIDurableEncode(mRibbonPhaseStandardMode) & vbCrLf & _
        "StandardCompleted=" & CStr(mRibbonPhaseStandardCompleted) & vbCrLf & _
        "StandardFingerprint=" & UIDurableEncode(mRibbonPhaseStandardInputFingerprint) & vbCrLf & _
        "StandardButtonCount=" & CStr(mRibbonPhaseButtonCount) & vbCrLf & _
        "StandardXmlHash=" & UIDurableEncode(AdlerTextHash(mRibbonPhaseRibbonXml) & ":" & CStr(Len(mRibbonPhaseRibbonXml))) & vbCrLf & _
        "RibbonXmlChanged=" & CStr(ribbonXmlChanged) & vbCrLf & _
        "ApplyAfterBuild=" & CStr(applyAfterBuild) & vbCrLf & _
        "RecordedOn=" & UIDurableEncode(Format$(Now, "yyyy-mm-dd hh:nn:ss"))
End Function

Private Function ClearVerifiedRibbonTransactionRecoveryOnly(ByRef failureDetail As String) As Boolean
    'Clears only a verified ribbon-transaction recovery blocker during the normal committed-success path.
    On Error GoTo ErrHandler
    Dim durableOperation As String
    Dim durableDetail As String
    Dim durableType As String
    Dim durablePayload As String
    Dim durableStatus As String
    Dim originalOperation As String
    Dim originalDetail As String
    Dim originalPayload As String
    Dim checkOperation As String
    Dim checkDetail As String
    Dim checkType As String
    Dim checkPayload As String
    Dim checkStatus As String
    Dim durableExists As Boolean
    Dim deletionCompleted As Boolean
    Dim restoreErrNumber As Long
    Dim restoreErrSource As String
    Dim restoreErrDescription As String
    durableExists = UIDurableRecoveryRecordExists(durableOperation, durableDetail, durableType, durablePayload, durableStatus)
    If durableExists Then
        originalOperation = durableOperation
        originalDetail = durableDetail
        originalPayload = durablePayload
        If Len(Trim$(durableType)) = 0 Then
            failureDetail = "A durable recovery blocker exists, but its recovery type is blank. The committed ribbon marker was retained."
            Exit Function
        End If
        If StrComp(Trim$(durableType), UI_RECOVERY_TYPE_RIBBON_TRANSACTION, vbTextCompare) <> 0 Then
            failureDetail = "A non-ribbon durable recovery blocker is present and was not cleared." & vbCrLf & "Operation: " & durableOperation & vbCrLf & "Recovery type: " & durableType & vbCrLf & "Status: " & durableStatus
            Exit Function
        End If
        DeleteUIDurableRecoveryRecord
        deletionCompleted = True
        If UIDurableRecoveryRecordExists(checkOperation, checkDetail, checkType, checkPayload, checkStatus) Then
            mUIRecoveryBlocked = True
            mUIRecoveryOperationName = checkOperation
            mUIRecoveryDetail = checkDetail
            mUIRecoveryType = checkType
            mUIRecoveryPayload = checkPayload
            mUIDurableRecoveryOwned = False
            If Len(Trim$(checkType)) = 0 Then
                failureDetail = "A durable recovery blocker still exists after deletion, but its recovery type is blank. The committed ribbon marker was retained." & vbCrLf & "Operation: " & checkOperation & vbCrLf & "Status: " & checkStatus & vbCrLf & "Detail: " & checkDetail
            ElseIf StrComp(Trim$(checkType), UI_RECOVERY_TYPE_RIBBON_TRANSACTION, vbTextCompare) <> 0 Then
                failureDetail = "A durable recovery blocker still exists after deletion, but it is no longer a RibbonTransaction record. The committed ribbon marker was retained." & vbCrLf & "Operation: " & checkOperation & vbCrLf & "Recovery type: " & checkType & vbCrLf & "Status: " & checkStatus & vbCrLf & "Detail: " & checkDetail
            Else
                failureDetail = "The ribbon-transaction recovery blocker still exists after deletion. The committed ribbon marker was retained." & vbCrLf & "Operation: " & checkOperation & vbCrLf & "Status: " & checkStatus & vbCrLf & "Detail: " & checkDetail
            End If
            Exit Function
        End If
    ElseIf mUIRecoveryBlocked Then
        If Len(Trim$(mUIRecoveryOperationName)) = 0 Or InStr(1, mUIRecoveryOperationName, "Ribbon", vbTextCompare) = 0 Then
            failureDetail = "In-memory UI recovery is blocked, but ribbon ownership could not be established." & vbCrLf & "Operation: " & IIf(Len(Trim$(mUIRecoveryOperationName)) = 0, "(blank)", mUIRecoveryOperationName) & vbCrLf & "Detail: " & mUIRecoveryDetail
            Exit Function
        End If
    End If
    mUIRecoveryBlocked = False
    mUIRecoveryOperationName = vbNullString
    mUIRecoveryDetail = vbNullString
    mUIRecoveryType = vbNullString
    mUIRecoveryPayload = vbNullString
    mUIDurableRecoveryOwned = False
    Set mUIRecoveryWorkspace = Nothing
    ClearVerifiedRibbonTransactionRecoveryOnly = True
    Exit Function
RestoreBlockedRecord:
    On Error Resume Next
    If Len(Trim$(originalOperation)) = 0 Then originalOperation = "Ribbon recovery verification"
    If Len(Trim$(originalDetail)) = 0 Then originalDetail = failureDetail
    Err.Clear
    PersistUIDurableRecoveryRecord originalOperation, originalDetail & vbCrLf & "Partial cleanup detail: " & failureDetail, UI_RECOVERY_TYPE_RIBBON_TRANSACTION, originalPayload, RECOVERY_STATUS_BLOCKED
    restoreErrNumber = Err.Number
    restoreErrSource = Err.Source
    restoreErrDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    mUIRecoveryBlocked = True
    mUIRecoveryOperationName = originalOperation
    mUIRecoveryDetail = originalDetail & vbCrLf & "Partial cleanup detail: " & failureDetail
    mUIRecoveryType = UI_RECOVERY_TYPE_RIBBON_TRANSACTION
    mUIRecoveryPayload = originalPayload
    mUIDurableRecoveryOwned = False
    If restoreErrNumber <> 0 Then
        failureDetail = failureDetail & vbCrLf & "Equivalent blocked recovery record restoration also failed. Error " & CStr(restoreErrNumber) & " from " & restoreErrSource & ": " & restoreErrDescription
    End If
    Exit Function
ErrHandler:
    failureDetail = "Could not clear and verify the ribbon-transaction recovery blocker after a successful ribbon commit." & vbCrLf & "Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    If deletionCompleted Then GoTo RestoreBlockedRecord
End Function

Private Sub CountRequiredRibbonManifestTarget(ByVal objectName As String, ByVal objectTypeCode As String, ByRef healthReviewCount As Long, ByRef systemHealthReviewCount As Long, ByRef reviewAreaCount As Long, ByRef severityCount As Long, ByRef setupCount As Long, ByRef traceCount As Long)
    'Counts the six required targets by stable logical identity after reversing the current canonical physical target name.
    Dim objectKey As String
    If StrComp(objectTypeCode, ADCAT_OBJECT_TYPE_TABLE, vbTextCompare) = 0 Then
        objectKey = UCase$(Trim$(LogicalObjectNameFromPhysicalName(NAMING_CATEGORY_TABLE, objectName)))
    ElseIf StrComp(objectTypeCode, ADCAT_OBJECT_TYPE_QUERY, vbTextCompare) = 0 Then
        objectKey = UCase$(Trim$(LogicalObjectNameFromPhysicalName(NAMING_CATEGORY_QUERY, objectName)))
    Else
        Exit Sub
    End If
    Select Case objectKey
        Case "ADC_HEALTHREVIEW": healthReviewCount = healthReviewCount + 1
        Case "ADC_SYSTEM_HEALTHREVIEW": systemHealthReviewCount = systemHealthReviewCount + 1
        Case "ADC_HEALTHREVIEWAREA_LOOKUP": reviewAreaCount = reviewAreaCount + 1
        Case "ADC_HEALTHSEVERITY_LOOKUP": severityCount = severityCount + 1
        Case "ADC_SETUPRESULTS": setupCount = setupCount + 1
        Case "ADC_TRACELOG": traceCount = traceCount + 1
    End Select
End Sub

Private Function RibbonSystemManifestVisibleGroupKey(ByVal groupLabel As String) As String
    RibbonSystemManifestVisibleGroupKey = UCase$(RibbonDefaultGroupLabel(groupLabel))
End Function

Private Sub BuildRibbonSystemManifestGroupIndexes(ByVal systemRibbonGroupDefinitions As Collection, ByRef systemGroupIdentityByPlacement As Object, ByRef systemExpectedTabsByVisibleGroup As Object)
    Dim Definition As Variant
    Dim seenLabelTabs As Object
    Dim tabKey As String
    Dim visibleGroupKey As String
    Dim placementKey As String
    Dim labelTabKey As String
    Dim expectedTabs As String
    ' INTENTIONAL KEYED LOOKUP: systemGroupIdentityByPlacement, systemExpectedTabsByVisibleGroup, seenLabelTabs provides keyed membership/lookup over source rows or state already read by
    ' BuildRibbonSystemManifestGroupIndexes; it is a derived access path, not a second authority.
    Set systemGroupIdentityByPlacement = CreateObject("Scripting.Dictionary")
    systemGroupIdentityByPlacement.CompareMode = vbBinaryCompare
    Set systemExpectedTabsByVisibleGroup = CreateObject("Scripting.Dictionary")
    systemExpectedTabsByVisibleGroup.CompareMode = vbBinaryCompare
    Set seenLabelTabs = CreateObject("Scripting.Dictionary")
    seenLabelTabs.CompareMode = vbBinaryCompare
    For Each Definition In systemRibbonGroupDefinitions
        tabKey = CStr(Definition(3))
        visibleGroupKey = RibbonSystemManifestVisibleGroupKey(CStr(Definition(1)))
        placementKey = UCase$(tabKey) & "|" & visibleGroupKey
        If systemGroupIdentityByPlacement.Exists(placementKey) Then Err.Raise vbObjectError + 7261, "BuildRibbonSystemManifestGroupIndexes", "More than one canonical System Ribbon group resolves to the same final visible tab/group identity." & vbCrLf & "Tab key: " & tabKey & vbCrLf & "Visible group: " & RibbonDefaultGroupLabel(CStr(Definition(1)))
        systemGroupIdentityByPlacement.Add placementKey, CStr(Definition(0))
        labelTabKey = visibleGroupKey & "|" & UCase$(tabKey)
        If seenLabelTabs.Exists(labelTabKey) = False Then
            seenLabelTabs.Add labelTabKey, True
            expectedTabs = vbNullString
            If systemExpectedTabsByVisibleGroup.Exists(visibleGroupKey) Then expectedTabs = CStr(systemExpectedTabsByVisibleGroup(visibleGroupKey))
            If Len(expectedTabs) > 0 Then expectedTabs = expectedTabs & " | "
            systemExpectedTabsByVisibleGroup(visibleGroupKey) = expectedTabs & tabKey
        End If
    Next Definition
End Sub

Private Sub DeleteRibbonCommitMarker()
    Dim operationName As String
    Dim detailText As String
    Dim payloadText As String
    DeleteDurableRecoveryMarker UI_RIBBON_COMMIT_MARKER_KEY
    If CoreDurableRecoveryMarkerExists(UI_RIBBON_COMMIT_MARKER_KEY, operationName, detailText, payloadText) Then Err.Raise ERR_UI_DURABLE_RECOVERY_FAILED, "DeleteRibbonCommitMarker", "The durable ribbon commit marker still exists after deletion verification."
End Sub

Private Sub PersistRibbonCommitMarker(ByVal commitStatus As String, ByVal ribbonXmlChanged As Boolean, ByVal applyAfterBuild As Boolean)
    'Persists the ADC-Ribbon commit handoff through the current Core recovery store, including the H1-captured application policy.
    Dim payloadText As String
    Dim checkOperation As String
    Dim checkDetail As String
    Dim checkPayload As String
    If Len(commitStatus) = 0 Then Err.Raise vbObjectError + 7206, "PersistRibbonCommitMarker", "Ribbon commit-marker status is blank."
    payloadText = UISerializeRibbonCommitPayload(commitStatus, ribbonXmlChanged, applyAfterBuild)
    PersistDurableRecoveryMarker UI_RIBBON_COMMIT_MARKER_KEY, "Ribbon commit handoff", "ADC ribbon transaction requires verified post-commit completion.", payloadText
    If CoreDurableRecoveryMarkerExists(UI_RIBBON_COMMIT_MARKER_KEY, checkOperation, checkDetail, checkPayload) = False Then Err.Raise ERR_UI_DURABLE_RECOVERY_FAILED, "PersistRibbonCommitMarker", "The durable ribbon commit marker was absent during write verification."
    If StrComp(checkPayload, payloadText, vbBinaryCompare) <> 0 Then Err.Raise ERR_UI_DURABLE_RECOVERY_FAILED, "PersistRibbonCommitMarker", "The durable ribbon commit marker payload did not match after write verification."
End Sub

' Purpose: Resolves a ADC-Ribbon commit handoff from the current Core durable-recovery store using the application policy captured by the original H1 operation.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub ResolvePendingRibbonCommitMarker()
    'Resolves a ADC-Ribbon commit handoff from the current Core durable-recovery store using the application policy captured by the original H1 operation.
    On Error GoTo ErrHandler
    Dim markerOperation As String
    Dim markerDetail As String
    Dim markerPayload As String
    Dim markerVersion As String
    Dim markerStatus As String
    Dim ribbonXmlChangedText As String
    Dim applyAfterBuildText As String
    Dim expectedCountText As String
    Dim expectedXmlHash As String
    Dim ribbonXmlChanged As Boolean
    Dim applyAfterBuild As Boolean
    Dim expectedCount As Long
    Dim db As DAO.Database
    Dim storedXml As String
    Dim storedXmlExists As Boolean
    Dim actualXmlHash As String
    Dim propertyExists As Boolean
    Dim currentRibbonId As String
    Dim verifyRibbonId As String
    Dim customRibbonChanged As Boolean
    Dim ribbonConfigured As Boolean
    Dim RestartRequired As Boolean
    Dim restartWarningText As String
    Dim recoveryClearDetail As String
    If CoreDurableRecoveryMarkerExists(UI_RIBBON_COMMIT_MARKER_KEY, markerOperation, markerDetail, markerPayload) = False Then Exit Sub
    markerVersion = UIDurableNoteValue(markerPayload, "Version")
    markerStatus = UIDurableNoteValue(markerPayload, "CommitStatus")
    ribbonXmlChangedText = UIDurableNoteValue(markerPayload, "RibbonXmlChanged")
    applyAfterBuildText = UIDurableNoteValue(markerPayload, "ApplyAfterBuild")
    expectedCountText = UIDurableNoteValue(markerPayload, "StandardButtonCount")
    expectedXmlHash = UIDurableNoteValue(markerPayload, "StandardXmlHash")
    If StrComp(markerVersion, "4", vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7208, "ResolvePendingRibbonCommitMarker", "Unsupported ADC-Ribbon commit marker version: " & IIf(Len(markerVersion) = 0, "(blank)", markerVersion)
    If TryParseUIDurableBoolean(ribbonXmlChangedText, ribbonXmlChanged) = False Then Err.Raise vbObjectError + 7209, "ResolvePendingRibbonCommitMarker", "The durable ADC-Ribbon commit marker has an invalid RibbonXmlChanged value."
    If TryParseUIDurableBoolean(applyAfterBuildText, applyAfterBuild) = False Then Err.Raise vbObjectError + 7209, "ResolvePendingRibbonCommitMarker", "The durable ADC-Ribbon commit marker has an invalid ApplyAfterBuild value."
    If Len(expectedCountText) = 0 Or IsNumeric(expectedCountText) = False Then Err.Raise vbObjectError + 7209, "ResolvePendingRibbonCommitMarker", "The durable ADC-Ribbon commit marker has an invalid button count."
    expectedCount = CLng(expectedCountText)
    Set db = GetHostDb()
    storedXml = StoredRibbonXml(db, RIBBON_NAME_, storedXmlExists)
    actualXmlHash = AdlerTextHash(storedXml) & ":" & CStr(Len(storedXml))
    If storedXmlExists = False Or Len(Trim$(storedXml)) = 0 Or StrComp(actualXmlHash, expectedXmlHash, vbBinaryCompare) <> 0 Then
        mUIRecoveryBlocked = True
        mUIRecoveryOperationName = "Ambiguous ribbon commit recovery"
        mUIRecoveryType = UI_RECOVERY_TYPE_RIBBON_TRANSACTION
        mUIRecoveryDetail = "A durable ADC-Ribbon commit marker exists, but the committed ADC ribbon XML does not match the staged transaction. The marker was retained."
        PersistUIDurableRecoveryRecord mUIRecoveryOperationName, mUIRecoveryDetail, UI_RECOVERY_TYPE_RIBBON_TRANSACTION, vbNullString, RECOVERY_STATUS_BLOCKED
        Err.Raise vbObjectError + 7212, "ResolvePendingRibbonCommitMarker", mUIRecoveryDetail
    End If
    ValidateRibbonXmlForAccess storedXml
    ValidateRibbonRuntimeManifest RIBBON_NAME_, expectedCount
    If StrComp(markerStatus, UI_RIBBON_COMMIT_STATUS_COMMIT_PENDING, vbTextCompare) = 0 Then
        PersistRibbonCommitMarker UI_RIBBON_COMMIT_STATUS_COMMITTED, ribbonXmlChanged, applyAfterBuild
        markerStatus = UI_RIBBON_COMMIT_STATUS_COMMITTED
    End If
    If StrComp(markerStatus, UI_RIBBON_COMMIT_STATUS_COMMITTED, vbTextCompare) <> 0 And StrComp(markerStatus, UI_RIBBON_COMMIT_STATUS_POST_COMMIT_COMPLETE, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 7207, "ResolvePendingRibbonCommitMarker", "Unsupported ADC-Ribbon commit status: " & markerStatus
    currentRibbonId = DatabaseTextPropertyValue(db, "CustomRibbonID", propertyExists)
    ribbonConfigured = (propertyExists And StrComp(currentRibbonId, RIBBON_NAME_, vbTextCompare) = 0)
    If applyAfterBuild Then
        customRibbonChanged = Not ribbonConfigured
        If customRibbonChanged Then EnsureDaoPropertyValue db, "CustomRibbonID", dbText, RIBBON_NAME_, "ResolvePendingRibbonCommitMarker"
        verifyRibbonId = DatabaseTextPropertyValue(db, "CustomRibbonID", propertyExists)
        If propertyExists = False Or StrComp(verifyRibbonId, RIBBON_NAME_, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 7220, "ResolvePendingRibbonCommitMarker", "CustomRibbonID verification failed while completing durable H3 recovery with ApplyAfterBuild=True."
        ribbonConfigured = True
    End If
    RestartRequired = customRibbonChanged Or (ribbonXmlChanged And ribbonConfigured)
    If RestartRequired Then
        If TryMarkRibbonRestartRequired(restartWarningText) = False Then Err.Raise vbObjectError + 7210, "ResolvePendingRibbonCommitMarker", "Committed ADC-Ribbon changes were verified, but the required database-reopen marker could not be restored." & vbCrLf & restartWarningText
    End If
    If ClearVerifiedRibbonTransactionRecoveryOnly(recoveryClearDetail) = False Then Err.Raise vbObjectError + 7213, "ResolvePendingRibbonCommitMarker", recoveryClearDetail
    DeleteRibbonCommitMarker
    LogTraceMessage TRACE_LEVEL_INFO, "RibbonRecovery", "ResolvePendingRibbonCommitMarker", "Completed pending ADC-Ribbon post-commit recovery.", "Marker status: " & markerStatus & vbCrLf & "Ribbon XML changed: " & CStr(ribbonXmlChanged) & vbCrLf & "ApplyAfterBuild: " & CStr(applyAfterBuild) & vbCrLf & "ADC configured as startup Ribbon: " & CStr(ribbonConfigured) & vbCrLf & "Restart required: " & CStr(RestartRequired) & vbCrLf & "Standard buttons verified: " & CStr(expectedCount), "Resolve committed ribbon state", 0, vbNullString, vbNullString, RIBBON_NAME_, "Ribbon", "CodeDb", "RibbonRecovery"
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ResolvePendingRibbonCommitMarker", savedErrDescription
End Sub

Private Function RibbonExpectedManagedPlacement(ByVal objectTypeCode As String, ByVal objectName As String, ByRef expectedTabKey As String, ByRef expectedGroupLabel As String) As Boolean
    'Single final-placement resolution path for runtime ribbon validation.
    Dim normalizedObjectType As String
    Dim expectedTabLabel As String
    Dim expectedGroupOrder As Long
    Dim isGenerated As Boolean
    normalizedObjectType = objectTypeCode
    expectedTabKey = vbNullString
    expectedGroupLabel = vbNullString
    If ManagedObjectPlacementRibbon(objectName, normalizedObjectType, expectedTabKey, expectedTabLabel, expectedGroupLabel, expectedGroupOrder) Then
        RibbonExpectedManagedPlacement = True
        Exit Function
    End If
    If StrComp(normalizedObjectType, ADCAT_OBJECT_TYPE_FORM, vbTextCompare) = 0 Then
        If TryResolveFormsRibbonPlacement(objectName, normalizedObjectType, isGenerated, expectedGroupLabel, expectedGroupOrder) Then
            expectedTabKey = DEFAULT_RIBBON_TAB_KEY_FORMS
            RibbonExpectedManagedPlacement = True
        End If
    End If
End Function

Private Function RibbonExpectedManifestGroupForDiagnostics(ByVal tabKey As String, ByVal objectTypeCode As String, ByVal objectName As String) As String
    'Returns the expected final group from the same final-placement authority used for normal validation.
    Dim expectedTabKey As String
    Dim expectedGroupLabel As String
    If RibbonExpectedManagedPlacement(objectTypeCode, objectName, expectedTabKey, expectedGroupLabel) Then
        If StrComp(tabKey, expectedTabKey, vbTextCompare) = 0 Then RibbonExpectedManifestGroupForDiagnostics = expectedGroupLabel
    End If
End Function

Private Function RibbonSafeId(ByVal valueText As String) As String
    Dim i As Long
    Dim ch As String
    Dim resultText As String
    valueText = Trim$(valueText)
    For i = 1 To Len(valueText)
        If (i And 255) = 0 Then RibbonSetupControlPump
        ch = Mid$(valueText, i, 1)
        If ch Like "[A-Za-z0-9]" Then
            resultText = resultText & ch
        Else
            resultText = resultText & "_"
        End If
    Next i
    If Len(resultText) = 0 Then resultText = "ADC"
    If Not Left$(resultText, 1) Like "[A-Za-z]" Then resultText = "adc_" & resultText
    If Len(resultText) > 80 Then resultText = Left$(resultText, 80)
    RibbonSafeId = resultText
End Function


Private Function TryMarkRibbonRestartRequired(ByRef warningText As String) As Boolean
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim markedOnText As String
    markedOnText = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Set db = GetHostDb()
    Call EnsureDaoPropertyValue(db, RIBBON_RESTART_REQUIRED_PROPERTY, dbText, "True", "TryMarkRibbonRestartRequired")
    Call EnsureDaoPropertyValue(db, RIBBON_RESTART_REQUIRED_ON_PROPERTY, dbText, markedOnText, "TryMarkRibbonRestartRequired")
    MarkRibbonRestartRequiredForCurrentRun
    TryMarkRibbonRestartRequired = True
    On Error Resume Next
    LogTraceMessage TRACE_LEVEL_INFO, "Ribbon", "TryMarkRibbonRestartRequired", "Database reopen required before the newly generated ribbon is visible.", "Restart-required property: " & RIBBON_RESTART_REQUIRED_PROPERTY & vbCrLf & "Marked on: " & markedOnText, "Ribbon restart required", 0, vbNullString, vbNullString, RIBBON_NAME_, "Ribbon", "CodeDb", "Ribbon"
    On Error GoTo 0
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    warningText = "Restart-required property persistence failed after the ribbon changes were committed." & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
    On Error Resume Next
    LogTraceMessage TRACE_LEVEL_WARNING, "Ribbon", "TryMarkRibbonRestartRequired", "Ribbon changes were committed, but the restart-required property could not be persisted.", warningText, "Persist ribbon restart requirement", savedErrNumber, savedErrSource, savedErrDescription, RIBBON_NAME_, "Ribbon", "CodeDb", "Ribbon"
    On Error GoTo 0
End Function

Private Sub ValidateRequiredRibbonManifestCounts(ByVal ribbonName As String, ByVal healthReviewCount As Long, ByVal systemHealthReviewCount As Long, ByVal reviewAreaCount As Long, ByVal severityCount As Long, ByVal setupCount As Long, ByVal traceCount As Long)
    'Requires exactly one of each confirmed ADC Admin Health/Diagnostics target in each persisted ribbon manifest.
    If healthReviewCount = 1 And systemHealthReviewCount = 1 And reviewAreaCount = 1 And severityCount = 1 And setupCount = 1 And traceCount = 1 Then Exit Sub
    Err.Raise vbObjectError + 7167, "ValidateRequiredRibbonManifestCounts", _
        "A persisted ribbon manifest does not contain exactly one of each required " & DEFAULT_RIBBON_TAB_LABEL_ADMINISTRATION & " Health/Diagnostics target." & vbCrLf & _
        "Ribbon: " & ribbonName & vbCrLf & _
        ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReview") & "/Table: " & CStr(healthReviewCount) & vbCrLf & _
        ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_System_HealthReview") & "/Query: " & CStr(systemHealthReviewCount) & vbCrLf & _
        ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReviewArea_Lookup") & "/Table: " & CStr(reviewAreaCount) & vbCrLf & _
        ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthSeverity_Lookup") & "/Table: " & CStr(severityCount) & vbCrLf & _
        ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SetupResults") & "/Table: " & CStr(setupCount) & vbCrLf & _
        ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_TraceLog") & "/Table: " & CStr(traceCount)
End Sub

Private Sub ValidateRibbonBuilderTextAttributeLengths(ByVal ribbonXml As String)
    ValidateRibbonTextAttributeLength ribbonXml, "id", RIBBON_SHORT_MAX_ID_ATTRIBUTE_LENGTH, "short-tooltip"
    ValidateRibbonTextAttributeLength ribbonXml, "label", RIBBON_SHORT_MAX_LABEL_LENGTH, "short-tooltip"
    ValidateRibbonTextAttributeLength ribbonXml, "tag", RIBBON_SHORT_MAX_ID_ATTRIBUTE_LENGTH, "short-tooltip"
    ValidateRibbonTextAttributeLength ribbonXml, "getScreentip", RIBBON_SHORT_MAX_ID_ATTRIBUTE_LENGTH, "short-tooltip"
    ValidateRibbonTextAttributeLength ribbonXml, "getSupertip", RIBBON_SHORT_MAX_ID_ATTRIBUTE_LENGTH, "short-tooltip"
    ValidateRibbonTextAttributeLength ribbonXml, "imageMso", RIBBON_SHORT_MAX_ID_ATTRIBUTE_LENGTH, "short-tooltip"
    ValidateRibbonTextAttributeLength ribbonXml, "onAction", RIBBON_SHORT_MAX_ID_ATTRIBUTE_LENGTH, "short-tooltip"
End Sub

Private Sub ValidateRibbonImageMsoValues(ByVal ribbonXml As String)
    Dim searchStart As Long
    Dim attrStart As Long
    Dim valueStart As Long
    Dim valueEnd As Long
    Dim imageName As String
    searchStart = 1
    Do
        RibbonSetupControlPump
        attrStart = InStr(searchStart, ribbonXml, "imageMso=""", vbTextCompare)
        If attrStart = 0 Then Exit Do
        valueStart = attrStart + Len("imageMso=""")
        valueEnd = InStr(valueStart, ribbonXml, """", vbBinaryCompare)
        If valueEnd = 0 Then Err.Raise vbObjectError + 7103, "ValidateRibbonImageMsoValues", "Generated " & DEFAULT_UTILITY_LONG_NAME & " ribbon XML has an unterminated imageMso attribute."
        imageName = Mid$(ribbonXml, valueStart, valueEnd - valueStart)
        If RibbonAllowedImageMso(imageName) = False Then Err.Raise vbObjectError + 7104, "ValidateRibbonImageMsoValues", "Generated " & DEFAULT_UTILITY_LONG_NAME & " ribbon XML contains an unapproved imageMso value." & vbCrLf & "imageMso: " & imageName
        searchStart = valueEnd + 1
    Loop
End Sub

Private Sub AppendRibbonManifestValidationIssue(ByRef issueText As String, ByRef issueCount As Long, ByVal validationStage As String, ByVal manifestRowNumber As Long, ByVal buttonKey As String, ByVal tabKey As String, ByVal tabLabel As String, ByVal groupLabel As String, ByVal buttonLabel As String, ByVal screenTipCaption As String, ByVal imageMso As String, ByVal objectTypeCode As String, ByVal objectName As String, ByVal errorNumber As Long, ByVal errorSource As String, ByVal errorDescription As String)
    Dim detailText As String
    issueCount = issueCount + 1
    detailText = "Validation issue " & CStr(issueCount) & vbCrLf & "Stage: " & validationStage
    If manifestRowNumber > 0 Then
        detailText = detailText & vbCrLf & "Manifest row: " & CStr(manifestRowNumber)
        detailText = detailText & vbCrLf & "Button key: " & buttonKey
        detailText = detailText & vbCrLf & "Tab key: " & tabKey
        detailText = detailText & vbCrLf & "Tab label: " & tabLabel
        detailText = detailText & vbCrLf & "Group: " & groupLabel
        detailText = detailText & vbCrLf & "Button title: " & buttonLabel
        detailText = detailText & vbCrLf & "ScreenTip caption: " & screenTipCaption
        detailText = detailText & vbCrLf & "ImageMso: " & imageMso
        detailText = detailText & vbCrLf & "Object type: " & objectTypeCode
        detailText = detailText & vbCrLf & "Object: " & objectName
    End If
    detailText = detailText & vbCrLf & "Err.Number: " & CStr(errorNumber)
    detailText = detailText & vbCrLf & "Err.Source: " & errorSource
    detailText = detailText & vbCrLf & "Err.Description: " & errorDescription
    If Len(issueText) > 0 Then issueText = issueText & vbCrLf & vbCrLf
    issueText = issueText & detailText
    On Error Resume Next
    LogTraceMessage TRACE_LEVEL_ERROR, "RibbonManifest", "ValidateRibbonRuntimeManifest", "Ribbon runtime manifest validation issue captured; validation will continue.", detailText, validationStage, errorNumber, errorSource, errorDescription, objectName, objectTypeCode, "CodeDb", "Ribbon"
    On Error GoTo 0
End Sub

Private Sub ValidateRibbonRuntimeManifest(ByVal ribbonName As String, ByVal expectedButtonCount As Long, Optional ByVal suppliedSystemRibbonGroupDefinitions As Variant)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim manifestRows As Collection
    Dim manifestKeys As Object
    Dim manifestButtonKeys As Object
    Dim manifestObjectPlacements As Object
    Dim rowData As Variant
    Dim sqlText As String
    Dim actualButtonCount As Long
    Dim rowValidationInProgress As Boolean
    Dim failureStage As String
    Dim buttonKey As String
    Dim tabKey As String
    Dim tabLabel As String
    Dim groupLabel As String
    Dim buttonLabel As String
    Dim imageMso As String
    Dim screenTipCaption As String
    Dim objectTypeCode As String
    Dim objectName As String
    Dim manifestObjectIdentity As String
    Dim manifestPlacement As String
    Dim existingPlacement As String
    Dim healthReviewCount As Long
    Dim systemHealthReviewCount As Long
    Dim reviewAreaCount As Long
    Dim severityCount As Long
    Dim setupCount As Long
    Dim traceCount As Long
    Dim uiSettingsCount As Long
    Dim uiCustomMacroLauncherCount As Long
    Dim uiNavigationPaneCount As Long
    Dim uiRibbonCount As Long
    Dim uiFormDefinitionsCount As Long
    Dim systemGroupCounts As Object
    Dim systemRibbonGroupDefinitions As Collection
    Dim systemMacroDefinitions As Collection
    Dim systemGroupIdentityByPlacement As Object
    Dim systemExpectedTabsByVisibleGroup As Object
    Dim systemUnexpectedCount As Long
    Dim systemMisplacedCount As Long
    Dim systemLayoutDiagnosticText As String
    Dim systemLayoutDiagnosticRowCount As Long
    Dim systemLayoutDiagnosticTruncated As Boolean
    Dim validationIssueCount As Long
    Dim validationIssueText As String
    Dim capturedErrNumber As Long
    Dim capturedErrSource As String
    Dim capturedErrDescription As String
    If isMissing(suppliedSystemRibbonGroupDefinitions) = False Then
        If IsObject(suppliedSystemRibbonGroupDefinitions) Then Set systemRibbonGroupDefinitions = suppliedSystemRibbonGroupDefinitions
    End If
    If systemRibbonGroupDefinitions Is Nothing Then Set systemRibbonGroupDefinitions = x__UTIL_ADC_DEFAULTS.GetSystemRibbonGroupDefinitions()
    ' OPERATION-SCOPED CONSISTENCY: Resolve the canonical System Macro definitions once for this manifest validation so row classification and aggregate group-count validation reuse the same authoritative definition
    ' set and never depend on a row-branch side effect.
    Set systemMacroDefinitions = x__UTIL_ADC_DEFAULTS.GetSystemMacroDefinitions()
    BuildRibbonSystemManifestGroupIndexes systemRibbonGroupDefinitions, systemGroupIdentityByPlacement, systemExpectedTabsByVisibleGroup
    failureStage = "Open persisted runtime manifest"
    Set db = GetHostDb()
    sqlText = "SELECT " & SQL_FIELD_BUTTON_KEY & ", [" & FIELD_RIBBON_TAB_KEY & "], [" & FIELD_RIBBON_TAB_LABEL & "], [" & FIELD_RIBBON_GROUP_LABEL & "], " & SQL_FIELD_BUTTON_LABEL & ", " & SQL_FIELD_SCREEN_TIP_CAPTION & ", " & SQL_FIELD_IMAGE_MSO & ", " & SQL_FIELD_OBJECT_TYPE_CODE & ", " & SQL_FIELD_OBJECT_NAME & " FROM " & SharedDao_BracketName(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_RIBBON_BUTTON_TARGETS_TABLE_LOGICAL_ID)) & " WHERE [" & FIELD_RIBBON_NAME & "]=" & SharedDao_SqlQuote(ribbonName) & " ORDER BY " & OrderByPolicyText("UI_RIBBON_RUNTIME_MANIFEST") & ";"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    ' INTENTIONAL OPERATION SNAPSHOT: manifestRows captures one coherent source state for this invocation of ValidateRibbonRuntimeManifest; validity ends with the operation rather than becoming a live data
    ' authority.
    ' INTENTIONAL KEYED LOOKUP: manifestKeys, manifestButtonKeys, manifestObjectPlacements, systemGroupCounts provides keyed membership/lookup over source rows or state already read by ValidateRibbonRuntimeManifest;
    ' it is a derived access path, not a second authority.
    Set manifestRows = New Collection
    Set manifestKeys = CreateObject("Scripting.Dictionary")
    manifestKeys.CompareMode = vbBinaryCompare
    Set manifestButtonKeys = CreateObject("Scripting.Dictionary")
    manifestButtonKeys.CompareMode = vbBinaryCompare
    Set manifestObjectPlacements = CreateObject("Scripting.Dictionary")
    manifestObjectPlacements.CompareMode = vbTextCompare
    Set systemGroupCounts = CreateObject("Scripting.Dictionary")
    systemGroupCounts.CompareMode = vbTextCompare
    failureStage = "Copy persisted runtime manifest into stable values"
    Do While rs.EOF = False
        manifestRows.Add Array(Trim$(CStr(Nz(rs!buttonKey, vbNullString))), Trim$(CStr(Nz(rs.fields(FIELD_RIBBON_TAB_KEY), vbNullString))), Trim$(CStr(Nz(rs.fields(FIELD_RIBBON_TAB_LABEL), vbNullString))), Trim$(CStr(Nz(rs.fields(FIELD_RIBBON_GROUP_LABEL), vbNullString))), Trim$(CStr(Nz(rs!buttonLabel, vbNullString))), Trim$(CStr(Nz(rs!screenTipCaption, vbNullString))), Trim$(CStr(Nz(rs!imageMso, vbNullString))), Trim$(CStr(Nz(rs!objectTypeCode, vbNullString))), CStr(Nz(rs!objectName, vbNullString)))
        manifestKeys(UCase$(NormalizedObjectTypeCode(CStr(Nz(rs!objectTypeCode, vbNullString)))) & "|" & CStr(Nz(rs!objectName, vbNullString))) = True
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    failureStage = "Validate copied runtime manifest rows"
    For Each rowData In manifestRows
        actualButtonCount = actualButtonCount + 1
        buttonKey = CStr(rowData(0))
        tabKey = CStr(rowData(1))
        tabLabel = CStr(rowData(2))
        groupLabel = CStr(rowData(3))
        buttonLabel = CStr(rowData(4))
        screenTipCaption = CStr(rowData(5))
        imageMso = CStr(rowData(6))
        objectTypeCode = NormalizedObjectTypeCode(CStr(rowData(7)))
        objectName = CStr(rowData(8))
        rowValidationInProgress = True
        If manifestButtonKeys.Exists(buttonKey) Then
            AppendRibbonManifestValidationIssue validationIssueText, validationIssueCount, "Validate unique static Ribbon tag", actualButtonCount, buttonKey, tabKey, tabLabel, groupLabel, buttonLabel, screenTipCaption, imageMso, objectTypeCode, objectName, vbObjectError + 7272, "ValidateRibbonRuntimeManifest", "The persisted runtime manifest contains the same static Ribbon tag more than once. Every Ribbon item must have a unique tag."
        Else
            manifestButtonKeys.Add buttonKey, True
        End If
        manifestObjectIdentity = RibbonCatalogMetadataKey(objectName, objectTypeCode)
        manifestPlacement = UCase$(tabKey) & RIBBON_TAG_DELIMITER & UCase$(groupLabel)
        If manifestObjectPlacements.Exists(manifestObjectIdentity) Then
            existingPlacement = CStr(manifestObjectPlacements(manifestObjectIdentity))
            If StrComp(existingPlacement, manifestPlacement, vbBinaryCompare) <> 0 Then
                AppendRibbonManifestValidationIssue validationIssueText, validationIssueCount, "Validate single Ribbon-group ownership", actualButtonCount, buttonKey, tabKey, tabLabel, groupLabel, buttonLabel, screenTipCaption, imageMso, objectTypeCode, objectName, vbObjectError + 7273, "ValidateRibbonRuntimeManifest", "The same managed object appears in more than one Ribbon group. Existing placement: " & existingPlacement & "; additional placement: " & manifestPlacement
            Else
                AppendRibbonManifestValidationIssue validationIssueText, validationIssueCount, "Validate one Ribbon item per object", actualButtonCount, buttonKey, tabKey, tabLabel, groupLabel, buttonLabel, screenTipCaption, imageMso, objectTypeCode, objectName, vbObjectError + 7274, "ValidateRibbonRuntimeManifest", "The same managed object appears more than once in the persisted runtime manifest."
            End If
        Else
            manifestObjectPlacements.Add manifestObjectIdentity, manifestPlacement
        End If
        capturedErrNumber = 0
        capturedErrSource = vbNullString
        capturedErrDescription = vbNullString
        On Error Resume Next
        Err.Clear
        ValidateRibbonRuntimeManifestRow buttonKey, tabKey, tabLabel, groupLabel, buttonLabel, screenTipCaption, objectTypeCode, objectName
        capturedErrNumber = Err.Number
        capturedErrSource = Err.Source
        capturedErrDescription = Err.Description
        Err.Clear
        On Error GoTo ErrHandler
        If capturedErrNumber <> 0 Then
            AppendRibbonManifestValidationIssue validationIssueText, validationIssueCount, "Per-row validation", actualButtonCount, buttonKey, tabKey, tabLabel, groupLabel, buttonLabel, screenTipCaption, imageMso, objectTypeCode, objectName, capturedErrNumber, capturedErrSource, capturedErrDescription
        End If
        CountRequiredRibbonManifestTarget objectName, objectTypeCode, healthReviewCount, systemHealthReviewCount, reviewAreaCount, severityCount, setupCount, traceCount
        CountRequiredUiManifestGroup tabKey, groupLabel, uiSettingsCount, uiCustomMacroLauncherCount, uiNavigationPaneCount, uiRibbonCount, uiFormDefinitionsCount
        CountRequiredSystemManifestGroup tabKey, tabLabel, groupLabel, buttonLabel, objectTypeCode, objectName, systemGroupCounts, systemGroupIdentityByPlacement, systemExpectedTabsByVisibleGroup, systemMacroDefinitions, systemUnexpectedCount, systemMisplacedCount, systemLayoutDiagnosticText, systemLayoutDiagnosticRowCount, systemLayoutDiagnosticTruncated
        rowValidationInProgress = False
    Next rowData
    failureStage = "Validate runtime manifest totals and required groups"
    If actualButtonCount <> expectedButtonCount Then
        AppendRibbonManifestValidationIssue validationIssueText, validationIssueCount, "Validate manifest button count", 0, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, ribbonName, vbObjectError + 7130, "ValidateRibbonRuntimeManifest", "The persisted ribbon runtime manifest button count does not match the generated XML button count." & vbCrLf & "Ribbon: " & ribbonName & vbCrLf & "Expected: " & CStr(expectedButtonCount) & vbCrLf & "Actual: " & CStr(actualButtonCount)
    End If
    On Error Resume Next
    Err.Clear
    ValidateRequiredRibbonManifestCounts ribbonName, healthReviewCount, systemHealthReviewCount, reviewAreaCount, severityCount, setupCount, traceCount
    capturedErrNumber = Err.Number
    capturedErrSource = Err.Source
    capturedErrDescription = Err.Description
    Err.Clear
    On Error GoTo ErrHandler
    If capturedErrNumber <> 0 Then AppendRibbonManifestValidationIssue validationIssueText, validationIssueCount, "Validate required Ribbon targets", 0, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, ribbonName, capturedErrNumber, capturedErrSource, capturedErrDescription
    On Error Resume Next
    Err.Clear
    ValidateRequiredUiManifestGroups ribbonName, uiSettingsCount, uiCustomMacroLauncherCount, uiNavigationPaneCount, uiRibbonCount, uiFormDefinitionsCount
    capturedErrNumber = Err.Number
    capturedErrSource = Err.Source
    capturedErrDescription = Err.Description
    Err.Clear
    On Error GoTo ErrHandler
    If capturedErrNumber <> 0 Then AppendRibbonManifestValidationIssue validationIssueText, validationIssueCount, "Validate required UI groups", 0, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, ribbonName, capturedErrNumber, capturedErrSource, capturedErrDescription
    On Error Resume Next
    Err.Clear
    ValidateRequiredSystemManifestGroups ribbonName, systemGroupCounts, systemUnexpectedCount, systemMisplacedCount, systemLayoutDiagnosticText, systemLayoutDiagnosticRowCount, systemLayoutDiagnosticTruncated, systemRibbonGroupDefinitions, systemMacroDefinitions
    capturedErrNumber = Err.Number
    capturedErrSource = Err.Source
    capturedErrDescription = Err.Description
    Err.Clear
    On Error GoTo ErrHandler
    If capturedErrNumber <> 0 Then AppendRibbonManifestValidationIssue validationIssueText, validationIssueCount, "Validate canonical System Ribbon layout", 0, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, ribbonName, capturedErrNumber, capturedErrSource, capturedErrDescription
    failureStage = "Verify complete active managed-object ribbon representation"
    On Error Resume Next
    Err.Clear
    ValidateManagedObjectRibbonManifestCoverage ribbonName, manifestKeys
    capturedErrNumber = Err.Number
    capturedErrSource = Err.Source
    capturedErrDescription = Err.Description
    Err.Clear
    On Error GoTo ErrHandler
    If capturedErrNumber <> 0 Then AppendRibbonManifestValidationIssue validationIssueText, validationIssueCount, "Verify complete active managed-object Ribbon representation", 0, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, ribbonName, capturedErrNumber, capturedErrSource, capturedErrDescription
    If validationIssueCount > 0 Then
        failureStage = "Aggregate runtime manifest validation"
        Err.Raise vbObjectError + 7260, "ValidateRibbonRuntimeManifest", "The generated ribbon runtime manifest contains " & CStr(validationIssueCount) & " validation issue(s). Every persisted row and each post-scan contract were evaluated in this run." & vbCrLf & vbCrLf & validationIssueText
    End If
    LogTraceMessage TRACE_LEVEL_INFO, "RibbonManifest", "ValidateRibbonRuntimeManifest", "Ribbon runtime manifest validated successfully.", "Ribbon: " & ribbonName & vbCrLf & "Buttons validated: " & CStr(actualButtonCount) & vbCrLf & "Complete active managed-object representation: verified" & vbCrLf & "Canonical System Ribbon group layout: verified", "Validate ribbon manifest", 0, vbNullString, vbNullString, ribbonName, "Ribbon", "CodeDb", "Ribbon"
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim detailText As String
    Dim expectedFinalGroup As String
    Dim expectedContextTitle As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    expectedFinalGroup = RibbonExpectedManifestGroupForDiagnostics(tabKey, objectTypeCode, objectName)
    Err.Clear
    expectedContextTitle = RibbonAuthoritativeButtonTitle(objectTypeCode, objectName)
    Err.Clear
    On Error GoTo 0
    detailText = "The generated ribbon runtime manifest failed validation." & vbCrLf & "Ribbon: " & ribbonName & vbCrLf & "Failure stage: " & failureStage & vbCrLf & "Failure phase: " & IIf(rowValidationInProgress, "Per-row validation", "Manifest read or post-scan validation") & vbCrLf & "Manifest row reached: " & IIf(rowValidationInProgress, CStr(actualButtonCount), "(not applicable)")
    detailText = detailText & vbCrLf & "Rows validated before failure: " & CStr(IIf(rowValidationInProgress And actualButtonCount > 0, actualButtonCount - 1, actualButtonCount)) & vbCrLf & "Expected XML button count: " & CStr(expectedButtonCount) & vbCrLf & "Persisted manifest row count reached: " & CStr(actualButtonCount)
    If rowValidationInProgress Then
        detailText = detailText & vbCrLf & "Failing row context:" & vbCrLf & "Button key: " & buttonKey & vbCrLf & "Tab key: " & tabKey & vbCrLf & "Tab label: " & tabLabel & vbCrLf & "Group: " & groupLabel & vbCrLf & "Button title: " & buttonLabel & vbCrLf & "ScreenTip caption: " & screenTipCaption & vbCrLf & "ImageMso: " & imageMso
        detailText = detailText & vbCrLf & "Object type: " & objectTypeCode & vbCrLf & "Normalized object type: " & objectTypeCode & vbCrLf & "Object: " & objectName & vbCrLf & "Normalized object key: " & DefaultObjectNameComparisonKey(objectName)
        detailText = detailText & vbCrLf & "Expected final group: " & expectedFinalGroup & vbCrLf & "Expected authoritative title: " & expectedContextTitle
    Else
        detailText = detailText & vbCrLf & "Failing row context: not applicable. All persisted rows were scanned when aggregate validation was active; otherwise the failure occurred during manifest read or post-scan validation."
    End If
    detailText = detailText & vbCrLf & vbCrLf & "System Ribbon counters at failure:" & vbCrLf & RequiredSystemManifestCountSummary(systemGroupCounts, systemMisplacedCount, systemUnexpectedCount, systemRibbonGroupDefinitions)
    If Len(systemLayoutDiagnosticText) > 0 Then detailText = detailText & vbCrLf & vbCrLf & "Affected system-layout rows:" & vbCrLf & systemLayoutDiagnosticText
    detailText = detailText & vbCrLf & "Manifest SQL: " & sqlText & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
    On Error Resume Next
    LogTraceMessage TRACE_LEVEL_ERROR, "RibbonManifest", "ValidateRibbonRuntimeManifest", "Ribbon runtime manifest validation failed after scanning " & CStr(actualButtonCount) & " row(s).", detailText, "Validate ribbon manifest", savedErrNumber, savedErrSource, savedErrDescription, objectName, objectTypeCode, "CodeDb", "Ribbon"
    On Error GoTo 0
    Err.Raise savedErrNumber, "ValidateRibbonRuntimeManifest", detailText
End Sub

Private Sub ValidateRibbonRuntimeManifestRow(ByVal buttonKey As String, ByVal tabKey As String, ByVal tabLabel As String, ByVal groupLabel As String, ByVal buttonLabel As String, ByVal screenTipCaption As String, ByVal objectTypeCode As String, ByVal objectName As String)
    Dim normalizedTab As String
    Dim normalizedGroup As String
    Dim expectedButtonKey As String
    Dim expectedGroup As String
    Dim expectedTab As String
    Dim expectedTitle As String
    Dim expectedScreenTipCaption As String
    Dim targetMetadata As Variant
    Dim targetTableNames As String
    Dim targetQueryNames As String
    Dim systemMacroKey As String
    Dim titleGroupLabel As String
    normalizedTab = UCase$(tabKey)
    normalizedGroup = UCase$(groupLabel)
    expectedButtonKey = RibbonStaticButtonKey(objectTypeCode, objectName)
    If Len(buttonKey) = 0 Then Err.Raise vbObjectError + 7275, "ValidateRibbonRuntimeManifestRow", "A persisted Ribbon item has a blank static tag key." & vbCrLf & "Object: " & objectName
    If StrComp(buttonKey, expectedButtonKey, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7276, "ValidateRibbonRuntimeManifestRow", "A persisted Ribbon item does not use the permanent tag stored for its managed-object identity." & vbCrLf & "Object type: " & objectTypeCode & vbCrLf & "Object: " & objectName & vbCrLf & "Expected tag: " & expectedButtonKey & vbCrLf & "Actual tag: " & buttonKey
    If Len(buttonLabel) = 0 Then Err.Raise vbObjectError + 7131, "ValidateRibbonRuntimeManifestRow", "A Ribbon button title is blank."
    If Len(screenTipCaption) = 0 Then Err.Raise vbObjectError + 7251, "ValidateRibbonRuntimeManifestRow", "A persisted Ribbon ScreenTip caption is blank." & vbCrLf & "Object: " & objectName
    If StrComp(objectTypeCode, ADCAT_OBJECT_TYPE_MACRO, vbTextCompare) = 0 Then
        If mRibbonSystemMacroTargetCache Is Nothing Then PrepareRibbonCatalogSourceSnapshot
        If mRibbonSystemMacroTargetCache.Exists(objectName) Then
            targetMetadata = mRibbonSystemMacroTargetCache(objectName)
            targetTableNames = CStr(targetMetadata(0))
            targetQueryNames = CStr(targetMetadata(1))
            systemMacroKey = CStr(targetMetadata(2))
            expectedScreenTipCaption = RibbonGeneratedScreenTipCaption(objectTypeCode, objectName)
            If StrComp(screenTipCaption, expectedScreenTipCaption, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7308, "ValidateRibbonRuntimeManifestRow", "A persisted canonical System Macro ScreenTip does not match the current generated ScreenTip authority." & vbCrLf & "System Macro key: " & systemMacroKey & vbCrLf & "System Macro title: " & objectName & vbCrLf & "Expected ScreenTip: " & expectedScreenTipCaption & vbCrLf & "Actual ScreenTip: " & screenTipCaption & vbCrLf & "Target tables: " & IIf(Len(targetTableNames) = 0, "(none)", targetTableNames) & vbCrLf & "Target queries: " & IIf(Len(targetQueryNames) = 0, "(none)", targetQueryNames)
        End If
    End If
    RibbonExpectedManagedPlacement objectTypeCode, objectName, expectedTab, expectedGroup
    titleGroupLabel = expectedGroup
    If Len(titleGroupLabel) = 0 Then titleGroupLabel = groupLabel
    expectedTitle = RibbonAuthoritativeButtonTitle(objectTypeCode, objectName)
    If Len(expectedTitle) = 0 Then Err.Raise vbObjectError + 7162, "ValidateRibbonRuntimeManifestRow", "No authoritative Ribbon button title was resolved." & vbCrLf & "Object: " & objectName & vbCrLf & "Object type: " & objectTypeCode & vbCrLf & "Tab: " & tabKey & vbCrLf & "Group: " & groupLabel
    If StrComp(buttonLabel, expectedTitle, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7162, "ValidateRibbonRuntimeManifestRow", "A persisted Ribbon title does not match the authoritative title." & vbCrLf & "Object: " & objectName & vbCrLf & "Expected title: " & expectedTitle & vbCrLf & "Actual title: " & buttonLabel
    If Len(expectedTab) > 0 And StrComp(normalizedTab, expectedTab, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 7165, "ValidateRibbonRuntimeManifestRow", "A managed Ribbon object is assigned to the wrong final tab." & vbCrLf & "Object: " & objectName & vbCrLf & "Expected tab key: " & expectedTab & vbCrLf & "Actual tab key: " & tabKey
    If normalizedTab = UCase$(DEFAULT_RIBBON_TAB_KEY_ADMINISTRATION) And StrComp(tabLabel, DEFAULT_RIBBON_TAB_LABEL_ADMINISTRATION, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7163, "ValidateRibbonRuntimeManifestRow", "The " & DEFAULT_RIBBON_TAB_LABEL_ADMINISTRATION & " runtime-manifest row does not use the confirmed tab label." & vbCrLf & "Expected tab label: " & DEFAULT_RIBBON_TAB_LABEL_ADMINISTRATION & vbCrLf & "Actual tab label: " & tabLabel
    If normalizedTab = DEFAULT_RIBBON_TAB_KEY_SYSTEM_1 And StrComp(tabLabel, DEFAULT_RIBBON_TAB_LABEL_SYSTEM_1, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7227, "ValidateRibbonRuntimeManifestRow", "An ADC1 runtime-manifest row does not use the confirmed tab label." & vbCrLf & "Expected tab label: " & DEFAULT_RIBBON_TAB_LABEL_SYSTEM_1 & vbCrLf & "Actual tab label: " & tabLabel
    If normalizedTab = DEFAULT_RIBBON_TAB_KEY_SYSTEM_2 And StrComp(tabLabel, DEFAULT_RIBBON_TAB_LABEL_SYSTEM_2, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7228, "ValidateRibbonRuntimeManifestRow", "An ADC2 runtime-manifest row does not use the confirmed tab label." & vbCrLf & "Expected tab label: " & DEFAULT_RIBBON_TAB_LABEL_SYSTEM_2 & vbCrLf & "Actual tab label: " & tabLabel
    If normalizedTab = DEFAULT_RIBBON_TAB_KEY_UI And StrComp(tabLabel, DEFAULT_RIBBON_TAB_LABEL_UI, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7176, "ValidateRibbonRuntimeManifestRow", "The " & DEFAULT_RIBBON_TAB_LABEL_UI & " runtime-manifest row does not use the confirmed tab label." & vbCrLf & "Expected tab label: " & DEFAULT_RIBBON_TAB_LABEL_UI & vbCrLf & "Actual tab label: " & tabLabel
    If normalizedTab = UCase$(DEFAULT_RIBBON_TAB_KEY_FORMS) And StrComp(tabLabel, DEFAULT_RIBBON_TAB_LABEL_FORMS, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7180, "ValidateRibbonRuntimeManifestRow", "An " & DEFAULT_RIBBON_TAB_LABEL_FORMS & " runtime-manifest row does not use the confirmed tab label." & vbCrLf & "Expected tab label: " & DEFAULT_RIBBON_TAB_LABEL_FORMS & vbCrLf & "Actual tab label: " & tabLabel
    If StrComp(objectName, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SystemMacros"), vbTextCompare) = 0 And StrComp(objectTypeCode, ADCAT_OBJECT_TYPE_TABLE, vbTextCompare) = 0 Then
        If normalizedTab <> DEFAULT_RIBBON_TAB_KEY_UI Or normalizedGroup <> UCase$(RibbonManagedGroupLabel("SETUP")) Then Err.Raise vbObjectError + 7134, "ValidateRibbonRuntimeManifestRow", "The [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SystemMacros") & "] table is not assigned to " & DEFAULT_RIBBON_TAB_LABEL_UI & " > Setup."
    End If
    If Len(expectedGroup) > 0 Then
        expectedGroup = RibbonDefaultGroupLabel(expectedGroup)
        If StrComp(groupLabel, expectedGroup, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 7136, "ValidateRibbonRuntimeManifestRow", "A managed Ribbon object is assigned to the wrong final group." & vbCrLf & "Object: " & objectName & vbCrLf & "Expected group: " & expectedGroup & vbCrLf & "Actual group: " & groupLabel
    End If
    If normalizedTab = UCase$(DEFAULT_RIBBON_TAB_KEY_FORMS) Then
        If Len(expectedGroup) = 0 Then Err.Raise vbObjectError + 7182, "ValidateRibbonRuntimeManifestRow", "An " & DEFAULT_RIBBON_TAB_LABEL_FORMS & " object does not have an authoritative catalog-resolved Ribbon group." & vbCrLf & "Group: " & groupLabel & vbCrLf & "Object: " & objectName
    End If
End Sub

Private Sub ValidateRibbonTextAttributeLength(ByVal ribbonXml As String, ByVal attributeName As String, ByVal maxLength As Long, ByVal ribbonModeName As String)
    Dim searchStart As Long
    Dim attrStart As Long
    Dim valueStart As Long
    Dim valueEnd As Long
    Dim valueText As String
    searchStart = 1
    Do
        RibbonSetupControlPump
        attrStart = InStr(searchStart, ribbonXml, attributeName & "=""", vbTextCompare)
        If attrStart = 0 Then Exit Do
        valueStart = attrStart + Len(attributeName & "=""")
        valueEnd = InStr(valueStart, ribbonXml, """", vbBinaryCompare)
        If valueEnd = 0 Then Err.Raise vbObjectError + 7105, "ValidateRibbonTextAttributeLength", "Generated " & DEFAULT_UTILITY_LONG_NAME & " ribbon XML has an unterminated " & attributeName & " attribute."
        valueText = Mid$(ribbonXml, valueStart, valueEnd - valueStart)
        If Len(valueText) > maxLength Then Err.Raise vbObjectError + 7106, "ValidateRibbonTextAttributeLength", "Generated " & DEFAULT_UTILITY_LONG_NAME & " " & ribbonModeName & " ribbon XML contains an attribute value that exceeds its approved limit." & vbCrLf & "Attribute: " & attributeName & vbCrLf & "Length: " & CStr(Len(valueText)) & vbCrLf & "Limit: " & CStr(maxLength)
        searchStart = valueEnd + 1
    Loop
End Sub

Private Sub ValidateRibbonXmlForAccess(ByVal ribbonXml As String, Optional ByVal requireBuilderSafeTextAttributes As Boolean = False)
    Const PROC_NAME As String = "ValidateRibbonXmlForAccess"
    If Len(Trim$(ribbonXml)) = 0 Then Err.Raise vbObjectError + 7100, PROC_NAME, "Generated " & DEFAULT_UTILITY_LONG_NAME & " ribbon XML is blank."
    If InStr(1, ribbonXml, " size=""small""", vbTextCompare) > 0 Then Err.Raise vbObjectError + 7101, PROC_NAME, "Generated " & DEFAULT_UTILITY_LONG_NAME & " ribbon XML contains invalid Access ribbon button size: small. Use normal or large."
    If InStr(1, ribbonXml, "<customUI", vbTextCompare) = 0 Then Err.Raise vbObjectError + 7102, PROC_NAME, "Generated " & DEFAULT_UTILITY_LONG_NAME & " ribbon XML is missing customUI root."
    ValidateRibbonImageMsoValues ribbonXml
    ValidateRibbonTooltipCallbackContract ribbonXml
    If requireBuilderSafeTextAttributes Then
        ValidateRibbonBuilderTextAttributeLengths ribbonXml
    Else
        ValidateRibbonTextAttributeLength ribbonXml, "id", RIBBON_STANDARD_MAX_ID_ATTRIBUTE_LENGTH, "standard"
    End If
End Sub

' ===============================================================================================================================================================================
' RIBBON PHASE PREPARATION / XML CONSTRUCTION PROCEDURES
' ===============================================================================================================================================================================

Private Function ResolveRibbonTabLabel(ByVal tabKey As String, Optional ByVal configuredTabLabel As String = vbNullString) As String
    If Len(configuredTabLabel) > 0 Then
        ResolveRibbonTabLabel = RibbonTrimLabel(configuredTabLabel)
        Exit Function
    End If
    ResolveRibbonTabLabel = RibbonTrimLabel(GetRibbonTabLabelDefault(tabKey))
End Function

Private Sub RequireRibbonPhaseToken(ByVal sourceProcedure As String)
    RequireRibbonPhasePrepared sourceProcedure
    If Len(mRibbonPhaseToken) = 0 Then Err.Raise vbObjectError + 7201, sourceProcedure, "Ribbon readiness state has no preparation token."
End Sub

' Purpose: Executes Setup H2 for the standard ADC Ribbon: validates prerequisites, stages target-map/XML changes in the owned recovery transaction, and records build/validation state for the later commit phase.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub BuildStandardRibbonPhase()
    RequireSetupDomainMutationAllowed SETUP_DOMAIN_H2_RIBBON_BUILD, "BuildStandardRibbonPhase"
    ' Part H2 owner: builds and validates the complete ADC Ribbon in one rollback-staged transaction.
    On Error GoTo ErrHandler
    Dim phaseStartedTimer As Single
    Dim systemRibbonGroupDefinitions As Collection
    Dim ribbonSchemaRepairAllowed As Boolean
    Dim ribbonSchemaDetail As String
    RequireRibbonPhaseToken "BuildStandardRibbonPhase"
    phaseStartedTimer = Timer
    PrintRibbonStageStatus 2, "H2 build and validate staged ADC Ribbon."
    If mRibbonPhaseSkippedSameDatabase Then
        mRibbonPhaseStandardCompleted = True
        Exit Sub
    End If
    ribbonSchemaRepairAllowed = SetupDomainMutationWritesAllowedForCurrentPart(SETUP_DOMAIN_A2_PHYSICAL_SCHEMA)
    If ribbonSchemaRepairAllowed Then
        EnsureUSysRibbonsTableSchema
    ElseIf USysRibbonsTableSchemaCurrentForSetup(ribbonSchemaDetail) = False Then
        Err.Raise vbObjectError + 9631, "BuildStandardRibbonPhase", "H2 refused to consume stale A2-owned USysRibbons physical schema. H2 does not repair A2." & vbCrLf & ribbonSchemaDetail
    End If
    If mRibbonPhaseStandardRequired = False Then
        PrintRibbonStatus "H2 verifying the H1 ADC Ribbon no-op plan against current source and persisted state."
        RequireRibbonPhaseSourcePlanCurrent "BuildStandardRibbonPhase", True
        mRibbonPhaseStandardCompleted = True
        mRibbonPhaseStandardElapsedSeconds = ElapsedSecondsFromTimer(phaseStartedTimer)
        Exit Sub
    End If
    If mRibbonPhaseTransactionStarted Then Err.Raise vbObjectError + 7192, "BuildStandardRibbonPhase", "A Ribbon staging transaction is already active."
    PrepareUIDurableRecovery "Ribbon transaction rollback", "H2 started the ADC Ribbon persistence transaction. An interruption requires verified rollback before later work.", UI_RECOVERY_TYPE_RIBBON_TRANSACTION, vbNullString
    Set mRibbonPhaseWorkspace = DBEngine.Workspaces(0)
    mRibbonPhaseWorkspace.BeginTrans
    mRibbonPhaseTransactionStarted = True
    PrintRibbonStatus "H2 staging ADC Ribbon target map and XML."
    RequireRibbonPhaseSourcePlanCurrent "BuildStandardRibbonPhase", True
    ClearRibbonButtonTargets RIBBON_NAME_
    ' H1 stores the expected source-row count; H2 must count generated buttons from zero.
    mRibbonPhaseButtonCount = 0
    Set systemRibbonGroupDefinitions = x__UTIL_ADC_DEFAULTS.GetSystemRibbonGroupDefinitions()
    mRibbonPhaseRibbonXml = BuildRibbonXml(mRibbonPhaseButtonCount, systemRibbonGroupDefinitions)
    If mRibbonPhaseButtonCount <> mRibbonPhaseExpectedStandardCount Then Err.Raise vbObjectError + 7264, "BuildStandardRibbonPhase", "The generated ADC Ribbon button count does not match the H1 source-row plan." & vbCrLf & "Expected source rows: " & CStr(mRibbonPhaseExpectedStandardCount) & vbCrLf & "Generated buttons: " & CStr(mRibbonPhaseButtonCount)
    PrintRibbonStatus "H2 validating staged runtime manifest. Buttons=" & CStr(mRibbonPhaseButtonCount) & "."
    ValidateRibbonRuntimeManifest RIBBON_NAME_, mRibbonPhaseButtonCount, systemRibbonGroupDefinitions
    PrintRibbonStatus "H2 validating staged Ribbon XML."
    ValidateRibbonXmlForAccess mRibbonPhaseRibbonXml
    mRibbonPhaseStandardChanged = UpsertRibbonXml(RIBBON_NAME_, mRibbonPhaseRibbonXml)
    mRibbonPhaseStandardBuilt = True
    mRibbonPhaseStandardCompleted = True
    mRibbonPhaseStandardElapsedSeconds = ElapsedSecondsFromTimer(phaseStartedTimer)
    LogTraceMessage TRACE_LEVEL_INFO, "RibbonTiming", "BuildStandardRibbonPhase", "H2 staged ADC Ribbon build and validation completed.", "Buttons: " & CStr(mRibbonPhaseButtonCount) & vbCrLf & "XML changed: " & CStr(mRibbonPhaseStandardChanged) & vbCrLf & "Elapsed: " & FormattedElapsedSeconds(mRibbonPhaseStandardElapsedSeconds), "H2 staged Ribbon", 0, vbNullString, vbNullString, RIBBON_NAME_, "Ribbon", "CodeDb", "H2"
    Exit Sub
ErrHandler:
    RaiseRibbonPhaseFailure "BuildStandardRibbonPhase", Err.Number, Err.Source, Err.Description
End Sub

Private Sub RequireRibbonPhasePrepared(ByVal sourceProcedure As String)
    If mRibbonPhasePrepared = False Then Err.Raise vbObjectError + 7197, sourceProcedure, "Ribbon readiness verification has not completed for the current staged ribbon operation."
End Sub

' Purpose: Clears and reinitializes ribbon phase state, coordinating `End` as the implementation requires.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub ResetRibbonPhaseState()
    If mRibbonPhaseTransactionStarted Then Exit Sub
    EndRibbonButtonTargetInsertBatch
    mRibbonPhasePrepared = False
    mRibbonPhaseToken = vbNullString
    mRibbonPhaseStandardMode = vbNullString
    mRibbonPhaseStandardCompleted = False
    mRibbonPhaseBuildRequired = False
    mRibbonPhaseStandardRequired = False
    mRibbonPhaseStandardBuilt = False
    mRibbonPhaseFastPathUsed = False
    mRibbonPhaseSkippedSameDatabase = False
    mRibbonPhaseShowMessage = False
    mRibbonPhaseApplyAfterBuild = False
    mRibbonPhaseApplyAfterBuildCaptured = False
    mRibbonPhaseTraceRunOwned = False
    mRibbonPhaseMutationCommitted = False
    Set mRibbonPhaseWorkspace = Nothing
    mRibbonPhaseRibbonXml = vbNullString
    mRibbonPhaseDependencyToken = vbNullString
    mRibbonPhaseCatalogToken = vbNullString
    mRibbonPhasePredecessorToken = vbNullString
    mRibbonPhaseGroupFingerprint = vbNullString
    mRibbonPhaseStandardInputFingerprint = vbNullString
    mRibbonPhaseReadinessDetail = vbNullString
    mRibbonPhaseButtonCount = 0
    mRibbonPhaseExpectedStandardCount = 0
    Set mRibbonPhaseStandardRows = Nothing
    Set mRibbonSystemMacroTargetCache = Nothing
    Set mRibbonCatalogMetadataCache = Nothing
    Set mRibbonCatalogAliasMetadataCache = Nothing
    Set mRibbonCatalogSourceRows = Nothing
    mRibbonCatalogPreflightRowCount = 0
    mRibbonCatalogSnapshotElapsedSeconds = 0
    Set mRibbonConfiguredGroupCache = Nothing
    mRibbonPhaseSourceDetail = vbNullString
    mRibbonPhaseRowsPrepared = False
    mRibbonSourceSealActive = False
    mRibbonPhaseStandardChanged = False
    mRibbonPhaseRestartMarkerSaved = False
    mRibbonPhaseStartedTimer = 0
    mRibbonPhaseReadinessElapsedSeconds = 0
    mRibbonPhaseStandardElapsedSeconds = 0
    mRibbonPhaseCommitElapsedSeconds = 0
End Sub

Private Sub InvalidateRibbonSameRunH3Proof()
    mRibbonSameRunH3ProofTraceRunID = vbNullString
    mRibbonSameRunH3ProofHostDbName = vbNullString
    mRibbonSameRunH3DependencyToken = vbNullString
    mRibbonSameRunH3CatalogToken = vbNullString
    mRibbonSameRunH3PredecessorToken = vbNullString
End Sub

Private Sub PublishRibbonSameRunH3Proof()
    Dim activeTraceRunID As String
    activeTraceRunID = GetActiveTraceRunID()
    If Len(activeTraceRunID) = 0 Then InvalidateRibbonSameRunH3Proof: Exit Sub
    If Len(mRibbonPhaseDependencyToken) = 0 Or Len(mRibbonPhaseCatalogToken) = 0 Or Len(mRibbonPhasePredecessorToken) = 0 Then Err.Raise vbObjectError + 7266, "PublishRibbonSameRunH3Proof", "The committed same-run H3 proof is incomplete."
    mRibbonSameRunH3ProofTraceRunID = activeTraceRunID
    mRibbonSameRunH3ProofHostDbName = GetHostDb().Name
    mRibbonSameRunH3DependencyToken = mRibbonPhaseDependencyToken
    mRibbonSameRunH3CatalogToken = mRibbonPhaseCatalogToken
    mRibbonSameRunH3PredecessorToken = mRibbonPhasePredecessorToken
End Sub

Private Function TryRibbonSameRunH3ProofInputs(ByVal prevalidatedPredecessorToken As String, ByVal prevalidatedCatalogToken As String, ByRef expectedDependencyToken As String, ByRef detailText As String) As Boolean
    Dim activeTraceRunID As String
    activeTraceRunID = GetActiveTraceRunID()
    If Len(activeTraceRunID) = 0 Or StrComp(activeTraceRunID, mRibbonSameRunH3ProofTraceRunID, vbBinaryCompare) <> 0 Then detailText = "The retained H3 proof belongs to a different or inactive setup trace run.": Exit Function
    If StrComp(GetHostDb().Name, mRibbonSameRunH3ProofHostDbName, vbTextCompare) <> 0 Then detailText = "The retained H3 proof belongs to a different HostDb.": Exit Function
    If Len(prevalidatedPredecessorToken) = 0 Or StrComp(prevalidatedPredecessorToken, mRibbonSameRunH3PredecessorToken, vbBinaryCompare) <> 0 Then detailText = "The retained H3 proof does not match J1's current D4/E3/G1 predecessor token.": Exit Function
    If Len(prevalidatedCatalogToken) = 0 Or StrComp(prevalidatedCatalogToken, mRibbonSameRunH3CatalogToken, vbBinaryCompare) <> 0 Then detailText = "The retained H3 proof does not match J1's current complete-catalog token.": Exit Function
    If Len(mRibbonSameRunH3DependencyToken) = 0 Then detailText = "The retained H3 proof is incomplete.": Exit Function
    expectedDependencyToken = mRibbonSameRunH3DependencyToken
    detailText = "Same-run H3 lineage/dependency proof accepted after trace-run, HostDb, D4/E3/G1 predecessor, and catalog identity checks; mutable Ribbon source rows must still be reread."
    TryRibbonSameRunH3ProofInputs = True
End Function

Private Sub PrintRibbonFailureDiagnosticSnapshot(ByVal diagnosticText As String)
    On Error Resume Next
    Debug.Print DEFAULT_UTILITY_SHORT_NAME & " ribbon failure diagnostic snapshot:"
    Debug.Print diagnosticText
    On Error GoTo 0
End Sub

Private Sub RaiseRibbonPhaseFailure(ByVal sourceProcedure As String, ByVal savedErrNumber As Long, ByVal savedErrSource As String, ByVal savedErrDescription As String)
    Dim rollbackText As String
    Dim rollbackErrNumber As Long
    Dim rollbackErrSource As String
    Dim rollbackErrDescription As String
    Dim failureStageText As String
    Dim diagnosticText As String
    Dim messageText As String
    Dim compactDescription As String
    Dim traceRunOwned As Boolean
    Dim recoveryPersistErrNumber As Long
    Dim recoveryPersistErrSource As String
    Dim recoveryPersistErrDescription As String
    Dim recoveryClearDetail As String
    Dim showMessage As Boolean
    traceRunOwned = mRibbonPhaseTraceRunOwned
    showMessage = mRibbonPhaseShowMessage
    If mRibbonPhaseMutationCommitted Then
        rollbackText = "The ribbon XML and target-map transaction was already committed; no rollback was attempted."
        savedErrDescription = "The ADC Ribbon database changes committed successfully, but post-commit state persistence, reporting, or cleanup failed." & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Original description: " & savedErrDescription
        savedErrSource = sourceProcedure & ".PostCommit"
    ElseIf mRibbonPhaseTransactionStarted Then
        On Error Resume Next
        Err.Clear
        mRibbonPhaseWorkspace.Rollback
        rollbackErrNumber = Err.Number
        rollbackErrSource = Err.Source
        rollbackErrDescription = Err.Description
        On Error GoTo 0
        If rollbackErrNumber = 0 Then
            rollbackText = "Ribbon XML and target-map changes were rolled back."
            mRibbonPhaseTransactionStarted = False
            If ClearUIRecoveryFailure("Ribbon transaction rollback", recoveryClearDetail) = False Then rollbackText = rollbackText & vbCrLf & recoveryClearDetail
        Else
            rollbackText = "Ribbon rollback failed. Error " & CStr(rollbackErrNumber) & " from " & rollbackErrSource & ": " & rollbackErrDescription
            On Error Resume Next
            Err.Clear
            RegisterUIRecoveryFailure "Ribbon transaction rollback", rollbackText & vbCrLf & "Original ribbon-build error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription, mRibbonPhaseWorkspace
            recoveryPersistErrNumber = Err.Number
            recoveryPersistErrSource = Err.Source
            recoveryPersistErrDescription = Err.Description
            On Error GoTo 0
            If recoveryPersistErrNumber <> 0 Then
                mUIRecoveryBlocked = True
                savedErrNumber = ERR_UI_DURABLE_RECOVERY_FAILED
                savedErrSource = sourceProcedure
                savedErrDescription = "The staged ribbon transaction could not be rolled back, and the durable blocked-recovery record could not be persisted. The workspace transaction state is unresolved and later UI, setup, reconciliation, generated-object, Ribbon, and managed-object naming work is blocked." & vbCrLf & rollbackText & vbCrLf & "Durable recovery persistence failure: Error " & CStr(recoveryPersistErrNumber) & " from " & IIf(Len(Trim$(recoveryPersistErrSource)) = 0, "(blank source)", recoveryPersistErrSource) & ": " & recoveryPersistErrDescription
            Else
                savedErrNumber = ERR_RIBBON_ROLLBACK_FAILED
                savedErrSource = sourceProcedure
                savedErrDescription = "The staged ribbon transaction could not be rolled back. The workspace transaction state is unresolved and later UI, setup, reconciliation, generated-object, Ribbon, and managed-object naming work is blocked." & vbCrLf & rollbackText
            End If
        End If
    ElseIf mUIDurableRecoveryOwned And mRibbonPhaseMutationCommitted = False Then
        If ClearUIRecoveryFailure("Ribbon transaction rollback", recoveryClearDetail) = False Then savedErrDescription = savedErrDescription & vbCrLf & recoveryClearDetail
    End If
    If IsCancellationError(savedErrNumber) Then
        EndRibbonStatusParts
        If mRibbonPhaseTransactionStarted = False Then ResetRibbonPhaseState
        Err.Raise savedErrNumber, savedErrSource, savedErrDescription
    End If
    failureStageText = mRibbonLastStatusText
    diagnosticText = RibbonFailureDiagnosticSnapshot(sourceProcedure, savedErrNumber, savedErrSource, savedErrDescription, failureStageText)
    PrintRibbonFailureDiagnosticSnapshot diagnosticText
    If mUIRecoveryBlocked = False Then
        If mRibbonPhaseMutationCommitted Then
            PrintRibbonStatus "ADC Ribbon changes committed, but post-commit completion failed. Error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription
        Else
            PrintRibbonStatus "ADC Ribbon phase failed. Error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription
        End If
    End If
    compactDescription = Replace(savedErrDescription, vbCrLf, " | ")
    If Len(compactDescription) > 1600 Then compactDescription = Left$(compactDescription, 1600) & "..."
    messageText = IIf(mRibbonPhaseMutationCommitted, "The " & DEFAULT_UTILITY_LONG_NAME & " ADC Ribbon committed, but post-commit completion failed.", "Could not build the " & DEFAULT_UTILITY_LONG_NAME & " ADC Ribbon.") & vbCrLf & vbCrLf & "Procedure: " & sourceProcedure
    messageText = messageText & vbCrLf & "Failure stage: " & IIf(Len(Trim$(failureStageText)) = 0, "(not recorded)", failureStageText)
    messageText = messageText & vbCrLf & "Error: " & CStr(savedErrNumber) & " from " & IIf(Len(Trim$(savedErrSource)) = 0, "(blank source)", savedErrSource)
    messageText = messageText & vbCrLf & "Description: " & compactDescription
    messageText = messageText & vbCrLf & "Transaction rollback: " & IIf(Len(Trim$(rollbackText)) = 0, "No live ribbon replacement transaction had started.", rollbackText)
    messageText = messageText & vbCrLf & "UI recovery blocked: " & CStr(mUIRecoveryBlocked)
    messageText = messageText & TableDiagnosticsLocationForMessage()
    If mUIRecoveryBlocked Then
        On Error Resume Next
        Debug.Print DEFAULT_UTILITY_SHORT_NAME & " UI recovery blocker: " & Replace(messageText, vbCrLf, " | ")
        On Error GoTo 0
    End If
    On Error Resume Next
    LogTraceMessage TRACE_LEVEL_DETAIL, "RibbonDiagnostics", sourceProcedure, "Ribbon failure diagnostic snapshot captured.", diagnosticText, failureStageText, savedErrNumber, savedErrSource, savedErrDescription, RIBBON_NAME_, "Ribbon", "CodeDb", "RibbonDiagnostics"
    LogTraceMessage TRACE_LEVEL_ERROR, "Ribbon", sourceProcedure, IIf(mRibbonPhaseMutationCommitted, "ADC Ribbon changes committed, but post-commit completion failed.", "Could not build the " & DEFAULT_UTILITY_LONG_NAME & " ADC Ribbon."), messageText, failureStageText, savedErrNumber, savedErrSource, savedErrDescription, RIBBON_NAME_, "Ribbon", "CodeDb", "Ribbon"
    On Error GoTo 0
    CompleteUITraceRunIfOwned traceRunOwned, IIf(mRibbonPhaseMutationCommitted, "ADC Ribbon changes committed; post-commit completion failed.", "ADC Ribbon build failed."), False
    EndRibbonStatusParts
    If showMessage Then
        On Error Resume Next
        CoreShowErrorMessage _
            messageText, _
            vbCritical, _
            DEFAULT_UTILITY_SHORT_NAME & " Ribbon Build Failed", _
            sourceProcedure, _
            "UI"
        On Error GoTo 0
    End If
    If mRibbonPhaseTransactionStarted = False Then ResetRibbonPhaseState
    If showMessage Then Exit Sub
    Err.Raise savedErrNumber, sourceProcedure, messageText
End Sub

Private Sub PrintRibbonStageStatus(ByVal stageNumber As Long, ByVal statusText As String)
    If mRibbonProgressActive And mRibbonProgressMax > 0 Then
        If stageNumber < 1 Then stageNumber = 1
        If stageNumber > mRibbonProgressMax Then stageNumber = mRibbonProgressMax
        mRibbonProgressStep = stageNumber
    End If
    PrintRibbonStatus statusText
End Sub

Private Sub PrintRibbonStatus(ByVal statusText As String)
    If mRibbonSourceSealActive Then Exit Sub
    On Error Resume Next
    Dim rawStatusText As String
    Dim stageStatusText As String
    rawStatusText = Trim$(statusText)
    If Len(rawStatusText) = 0 Then Exit Sub
    mRibbonLastStatusText = rawStatusText
    stageStatusText = RibbonStatusPartText(rawStatusText, False)
    If CoreParentProgressActive() Then
        statusText = CoreParentNestedProgressText(stageStatusText, vbNullString)
    Else
        statusText = stageStatusText
    End If
    SetUIStatusBarText statusText
    LogTraceMessage TRACE_LEVEL_STATUS, "Ribbon", "PrintRibbonStatus", statusText, stageStatusText, "Ribbon status", 0, vbNullString, vbNullString, vbNullString, vbNullString, "CodeDb", "Ribbon"
    If mSetupControlActive Then
        RibbonSetupControlPump
    Else
        DoEvents
    End If
End Sub

Private Sub RibbonSetupControlPump()
    If mRibbonSourceSealActive Then Exit Sub
    SetupControlPump
End Sub

Private Sub RibbonRequireSafeWorkContinue(ByVal workerRoutine As String, ByVal logicalItem As String)
    If mRibbonSourceSealActive Then Exit Sub
    RequireSafeWorkContinue workerRoutine, logicalItem
End Sub

' Purpose: Performs the public ribbon on load operation using the routine's existing ADC workflow and side effects.
' Public role: Ribbon callback entry point. Office/Access invokes this Public routine from Ribbon XML; keeping it Public is required for callback resolution.
' Parameters:
'   ribbon (Input; IRibbonUI): Office Ribbon UI object supplied by Access when the custom Ribbon loads; retained/used by the routine for Ribbon invalidation and refresh operations.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub RibbonOnLoad(ByVal ribbon As IRibbonUI)
    'Intended use: retain the loaded ribbon object for callback invalidation, clear the persisted ribbon-restart requirement, and surface ribbon-state property-write failures through the terminal ADC error workflow
    'without re-raising into Office's onLoad callback.
    If mTerminalTeardownActive Then Exit Sub
    On Error Resume Next
    LogPhysicalTraceMessageOnly TRACE_LEVEL_INFO, "Ribbon", "RibbonOnLoad", "DATABASE STARTUP RIBBON ONLOAD ENTERED.", "TerminalTeardownActive=False", "Ribbon startup diagnostics", 0, vbNullString, vbNullString, RIBBON_NAME_, "Ribbon", "CodeDb", "RibbonStartup"
    Err.Clear
    Dim db As DAO.Database
    Dim loadedOnText As String
    Dim restartErrNumber As Long
    Dim restartErrSource As String
    Dim restartErrDescription As String
    Dim loadedErrNumber As Long
    Dim loadedErrSource As String
    Dim loadedErrDescription As String
    Set mRibbonUI = ribbon
    mRibbonLastTooltipCallbackErrorKey = vbNullString
    loadedOnText = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    LogPhysicalTraceMessageOnly TRACE_LEVEL_INFO, "Ribbon", "RibbonOnLoad", "Ribbon startup HostDb resolution starting.", "LoadedOn=" & loadedOnText, "Ribbon startup diagnostics", 0, vbNullString, vbNullString, RIBBON_NAME_, "Ribbon", "CodeDb", "RibbonStartup"
    Err.Clear
    Set db = GetHostDb()
    LogPhysicalTraceMessageOnly TRACE_LEVEL_INFO, "Ribbon", "RibbonOnLoad", "Ribbon startup HostDb resolution returned.", "HostDbAvailable=" & IIf(db Is Nothing, "False", "True"), "Ribbon startup diagnostics", 0, vbNullString, vbNullString, RIBBON_NAME_, "Ribbon", "CodeDb", "RibbonStartup"
    Err.Clear
    LogPhysicalTraceMessageOnly TRACE_LEVEL_INFO, "Ribbon", "RibbonOnLoad", "Ribbon restart-required property update starting.", "Property=" & RIBBON_RESTART_REQUIRED_PROPERTY, "Ribbon startup diagnostics", 0, vbNullString, vbNullString, RIBBON_NAME_, "Ribbon", "CodeDb", "RibbonStartup"
    Err.Clear
    Call EnsureDaoPropertyValue(db, RIBBON_RESTART_REQUIRED_PROPERTY, dbText, "False", "RibbonOnLoad")
    restartErrNumber = Err.Number
    restartErrSource = Err.Source
    restartErrDescription = Err.Description
    LogPhysicalTraceMessageOnly IIf(restartErrNumber = 0, TRACE_LEVEL_INFO, TRACE_LEVEL_WARNING), "Ribbon", "RibbonOnLoad", "Ribbon restart-required property update returned.", "Err.Number=" & CStr(restartErrNumber) & "; Err.Source=" & restartErrSource & "; Err.Description=" & restartErrDescription, "Ribbon startup diagnostics", restartErrNumber, restartErrSource, restartErrDescription, RIBBON_NAME_, "Ribbon", "CodeDb", "RibbonStartup"
    Err.Clear
    LogPhysicalTraceMessageOnly TRACE_LEVEL_INFO, "Ribbon", "RibbonOnLoad", "Ribbon loaded-on property update starting.", "Property=" & RIBBON_LOADED_ON_PROPERTY, "Ribbon startup diagnostics", 0, vbNullString, vbNullString, RIBBON_NAME_, "Ribbon", "CodeDb", "RibbonStartup"
    Err.Clear
    Call EnsureDaoPropertyValue(db, RIBBON_LOADED_ON_PROPERTY, dbText, loadedOnText, "RibbonOnLoad")
    loadedErrNumber = Err.Number
    loadedErrSource = Err.Source
    loadedErrDescription = Err.Description
    LogPhysicalTraceMessageOnly IIf(loadedErrNumber = 0, TRACE_LEVEL_INFO, TRACE_LEVEL_WARNING), "Ribbon", "RibbonOnLoad", "Ribbon loaded-on property update returned.", "Err.Number=" & CStr(loadedErrNumber) & "; Err.Source=" & loadedErrSource & "; Err.Description=" & loadedErrDescription, "Ribbon startup diagnostics", loadedErrNumber, loadedErrSource, loadedErrDescription, RIBBON_NAME_, "Ribbon", "CodeDb", "RibbonStartup"
    Err.Clear
    If restartErrNumber = 0 And loadedErrNumber = 0 Then
        LogTraceMessage TRACE_LEVEL_INFO, "Ribbon", "RibbonOnLoad", DEFAULT_UTILITY_LONG_NAME & " ribbon loaded; restart-required state cleared and callback invalidation handle retained.", "Ribbon loaded on: " & loadedOnText, "Ribbon onLoad", 0, vbNullString, vbNullString, RIBBON_NAME_, "Ribbon", "CodeDb", "Ribbon"
    Else
        LogTraceMessage TRACE_LEVEL_ERROR, "Ribbon", "RibbonOnLoad", DEFAULT_UTILITY_LONG_NAME & " ribbon loaded, but one or more ribbon-state properties could not be updated.", "Ribbon loaded on: " & loadedOnText & vbCrLf & "Restart flag Err.Number: " & CStr(restartErrNumber) & vbCrLf & "Restart flag Err.Source: " & restartErrSource & vbCrLf & "Restart flag Err.Description: " & restartErrDescription & vbCrLf & "Loaded-on Err.Number: " & CStr(loadedErrNumber) & vbCrLf & "Loaded-on Err.Source: " & loadedErrSource & vbCrLf & "Loaded-on Err.Description: " & loadedErrDescription, "Ribbon onLoad property update", IIf(restartErrNumber <> 0, restartErrNumber, loadedErrNumber), IIf(restartErrNumber <> 0, restartErrSource, loadedErrSource), IIf(restartErrNumber <> 0, restartErrDescription, loadedErrDescription), RIBBON_NAME_, "Ribbon", "CodeDb", "Ribbon"
        CoreShowErrorMessage DEFAULT_UTILITY_LONG_NAME & " ribbon loaded, but one or more ribbon-state properties could not be updated." & vbCrLf & vbCrLf & "Restart flag Err.Number: " & CStr(restartErrNumber) & vbCrLf & "Restart flag Err.Source: " & restartErrSource & vbCrLf & "Restart flag Err.Description: " & restartErrDescription & vbCrLf & "Loaded-on Err.Number: " & CStr(loadedErrNumber) & vbCrLf & "Loaded-on Err.Source: " & loadedErrSource & vbCrLf & "Loaded-on Err.Description: " & loadedErrDescription, vbCritical, DEFAULT_UI_MESSAGE_TITLE, "RibbonOnLoad", "Ribbon"
    End If
    FlushTraceLogFile
End Sub

' ===============================================================================================================================================================================
' RIBBON TARGET MAP / SOURCE / FINGERPRINT / COMMIT LIFECYCLE PROCEDURES
' ===============================================================================================================================================================================

Private Sub ClearRibbonButtonTargets(ByVal ribbonName As String)
    On Error GoTo ErrHandler
    GetHostDb().Execute "DELETE FROM " & SharedDao_BracketName(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_RIBBON_BUTTON_TARGETS_TABLE_LOGICAL_ID)) & " WHERE [" & FIELD_RIBBON_NAME & "]=" & SharedDao_SqlQuote(ribbonName) & ";", dbFailOnError
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ClearRibbonButtonTargets", "Could not clear the " & DEFAULT_UTILITY_LONG_NAME & " ribbon button target map." & vbCrLf & "Ribbon: " & ribbonName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub BeginRibbonButtonTargetInsertBatch()
    On Error GoTo ErrHandler
    EndRibbonButtonTargetInsertBatch
    Set mRibbonStaticTargetOwners = CreateObject("Scripting.Dictionary")
    mRibbonStaticTargetOwners.CompareMode = vbBinaryCompare
    Set mRibbonTargetInsertRs = GetHostDb().OpenRecordset(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_RIBBON_BUTTON_TARGETS_TABLE_LOGICAL_ID), dbOpenDynaset)
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    EndRibbonButtonTargetInsertBatch
    On Error GoTo 0
    Err.Raise savedErrNumber, "BeginRibbonButtonTargetInsertBatch", "Could not open the " & DEFAULT_UTILITY_LONG_NAME & " ribbon runtime-manifest append batch." & vbCrLf & "Table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_RIBBON_BUTTON_TARGETS_TABLE_LOGICAL_ID) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub EndRibbonButtonTargetInsertBatch()
    On Error Resume Next
    If Not mRibbonTargetInsertRs Is Nothing Then mRibbonTargetInsertRs.Close
    Set mRibbonTargetInsertRs = Nothing
    Set mRibbonStaticTargetOwners = Nothing
    On Error GoTo 0
End Sub

Private Function RibbonMappedButtonPresentation(ByVal controlTag As String, ByRef ribbonName As String, ByRef ribbonTabKey As String, ByRef objectTypeCode As String, ByRef objectName As String, ByRef screenTipCaption As String) As Boolean
    If mTerminalTeardownActive Then Exit Function
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim buttonKey As String
    Dim requestedRibbonName As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    If RibbonParseControlTag(controlTag, requestedRibbonName, buttonKey) = False Then Exit Function
    Set db = GetHostDb()
    If SharedDao_TableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_RIBBON_BUTTON_TARGETS_TABLE_LOGICAL_ID)) = False Then Exit Function
    sqlText = "SELECT TOP 1 [" & FIELD_RIBBON_NAME & "], [" & FIELD_RIBBON_TAB_KEY & "], " & SQL_FIELD_OBJECT_TYPE_CODE & ", " & SQL_FIELD_OBJECT_NAME & ", " & SQL_FIELD_SCREEN_TIP_CAPTION & " FROM " & SharedDao_BracketName(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_RIBBON_BUTTON_TARGETS_TABLE_LOGICAL_ID)) & " WHERE " & SQL_FIELD_BUTTON_KEY & "=" & SharedDao_SqlQuote(buttonKey)
    If Len(requestedRibbonName) > 0 Then sqlText = sqlText & " AND [" & FIELD_RIBBON_NAME & "]=" & SharedDao_SqlQuote(requestedRibbonName)
    sqlText = sqlText & " ORDER BY " & OrderByPolicyText("UI_RIBBON_MAPPED_PRESENTATION") & ";"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    If rs.EOF = False Then
        ribbonName = Trim$(CStr(Nz(rs.fields(FIELD_RIBBON_NAME).value, vbNullString)))
        ribbonTabKey = Trim$(CStr(Nz(rs.fields(FIELD_RIBBON_TAB_KEY).value, vbNullString)))
        objectTypeCode = Trim$(CStr(Nz(rs.fields(FIELD_OBJECT_TYPE_CODE).value, vbNullString)))
        objectName = CStr(Nz(rs.fields(FIELD_OBJECT_NAME).value, vbNullString))
        screenTipCaption = Trim$(CStr(Nz(rs.fields(FIELD_SCREEN_TIP_CAPTION).value, vbNullString)))
        RibbonMappedButtonPresentation = (Len(objectTypeCode) > 0 And Len(objectName) > 0)
    End If
    rs.Close
    Exit Function
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
    Err.Raise savedErrNumber, "RibbonMappedButtonPresentation", "Could not resolve the ribbon button target used by the action and tooltip callbacks." & vbCrLf & "Tag: " & controlTag & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function UpsertRibbonXml(ByVal ribbonName As String, ByVal ribbonXml As String) As Boolean
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim existingXml As String
    Dim matchingRowCount As Long
    Set db = GetHostDb()
    sqlText = "SELECT * FROM " & SharedDao_BracketName(DB_TABLE_USYS_RIBBONS) & " WHERE [" & FIELD_RIBBON_NAME & "]=" & SharedDao_SqlQuote(ribbonName) & ";"
    Set rs = db.OpenRecordset(sqlText, dbOpenDynaset)
    If rs.EOF Then
        rs.AddNew
        rs.fields(FIELD_RIBBON_NAME).value = ribbonName
        rs.fields(FIELD_RIBBON_XML).value = ribbonXml
        rs.Update
        UpsertRibbonXml = True
    Else
        rs.MoveLast
        matchingRowCount = rs.RecordCount
        rs.MoveFirst
        If matchingRowCount > 1 Then
            rs.Close
            Set rs = Nothing
            db.Execute "DELETE FROM " & SharedDao_BracketName(DB_TABLE_USYS_RIBBONS) & " WHERE [" & FIELD_RIBBON_NAME & "]=" & SharedDao_SqlQuote(ribbonName) & ";", dbFailOnError
            Set rs = db.OpenRecordset(sqlText, dbOpenDynaset)
            rs.AddNew
            rs.fields(FIELD_RIBBON_NAME).value = ribbonName
            rs.fields(FIELD_RIBBON_XML).value = ribbonXml
            rs.Update
            UpsertRibbonXml = True
            LogTraceMessage TRACE_LEVEL_WARNING, "RibbonPersistence", "UpsertRibbonXml", "Duplicate USysRibbons rows for the same CustomRibbonID were collapsed to one authoritative row.", "Ribbon: " & ribbonName & vbCrLf & "Rows replaced: " & CStr(matchingRowCount), "Repair duplicate RibbonName rows", 0, vbNullString, vbNullString, ribbonName, "Ribbon", "CodeDb", "RibbonPersistence"
        Else
            existingXml = CStr(Nz(rs.fields(FIELD_RIBBON_XML).value, vbNullString))
            If StrComp(existingXml, ribbonXml, vbBinaryCompare) <> 0 Then
                LogTraceMessage TRACE_LEVEL_DETAIL, "RibbonComparison", "UpsertRibbonXml", "Stored and generated ribbon XML differ.", UITextMismatchDetail(existingXml, ribbonXml), "Compare ribbon XML", 0, vbNullString, vbNullString, ribbonName, "Ribbon", "CodeDb", "RibbonComparison"
                rs.Edit
                rs.fields(FIELD_RIBBON_XML).value = ribbonXml
                rs.Update
                UpsertRibbonXml = True
            End If
        End If
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
    Err.Raise savedErrNumber, "UpsertRibbonXml", "Could not save " & DEFAULT_UTILITY_LONG_NAME & " custom UI XML." & vbCrLf & "Ribbon: " & ribbonName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function BuildRibbonXml(ByRef buttonCount As Long, ByVal systemGroupDefinitions As Collection) As String
    On Error GoTo ErrHandler
    Dim xmlText As String
    Dim sourceRows As Collection
    Dim rowData As Variant
    Dim currentTab As String
    Dim currentTabLabel As String
    Dim currentGroup As String
    Dim tabKey As String
    Dim tabLabel As String
    Dim groupKey As String
    Dim groupID As String
    Dim groupStaticTag As String
    Dim groupLabel As String
    Dim groupSortOrder As Long
    Dim currentRibbonGroupLabel As String
    Dim groupButtonCount As Long
    Dim groupSegmentNumber As Long
    Dim itemID As Long
    Dim currentObjectName As String
    Dim currentObjectType As String
    Dim currentButtonLabel As String
    Dim currentNormalizedTab As String
    Dim currentNormalizedGroup As String
    Dim currentNormalizedObjectType As String
    Dim usedButtonLabels As Collection
    Dim authoritativeLabelOwners As Collection
    Dim usedCustomIds As Collection
    Dim usedStaticTags As Collection
    Dim sourceDetailText As String
    Dim defaultTitleCount As Long
    ' INTENTIONAL KEYED LOOKUP: usedCustomIds, usedStaticTags provides keyed membership/lookup over source rows or state already read by BuildRibbonXml; it is a derived access path, not a second authority.
    Set usedCustomIds = New Collection
    Set usedStaticTags = New Collection
    RibbonAppendXmlLine xmlText, RIBBON_CUSTOM_UI_OPEN
    RibbonAppendXmlLine xmlText, RIBBON_XML_INDENT_1 & "<ribbon>"
    RibbonAppendXmlLine xmlText, RIBBON_XML_INDENT_2 & "<tabs>"
    Set sourceRows = RibbonPhaseRows(sourceDetailText)
    ValidateRibbonTabOrder sourceRows
    PrintRibbonStatus "Ribbon source rows ready: " & CStr(sourceRows.count) & "."
    ValidateAuthoritativeRibbonTitles sourceRows
    ValidateAndLogRibbonObjectIcons sourceRows
    BeginRibbonButtonTargetInsertBatch
    For Each rowData In sourceRows
        itemID = itemID + 1
        If itemID = 1 Or itemID Mod DEFAULT_RIBBON_PROGRESS_PRINT_INTERVAL = 0 Or itemID = sourceRows.count Then PrintRibbonStatus "Ribbon XML still working. Building button " & CStr(itemID) & " of " & CStr(sourceRows.count) & "."
        tabKey = CStr(rowData(0))
        tabLabel = CStr(rowData(1))
        groupLabel = CStr(rowData(2))
        groupSortOrder = CLng(rowData(6))
        currentObjectType = CStr(rowData(3))
        currentObjectName = CStr(rowData(4))
        groupKey = RibbonGroupCompareKey(tabKey, groupLabel)
        If StrComp(tabKey, currentTab, vbTextCompare) <> 0 Then
            If Len(currentGroup) > 0 Then RibbonAppendXmlLine xmlText, RIBBON_XML_INDENT_4 & "</group>"
            If Len(currentTab) > 0 Then RibbonAppendXmlLine xmlText, RIBBON_XML_INDENT_3 & "</tab>"
            RegisterRibbonCustomId usedCustomIds, RibbonRegularTabId(tabKey), "Tab", tabKey & " / " & tabLabel
            RibbonAppendXmlLine xmlText, RibbonTabOpenXml(tabKey, tabLabel, usedStaticTags)
            currentTab = tabKey
            currentTabLabel = tabLabel
            currentGroup = vbNullString
            groupButtonCount = 0
            groupSegmentNumber = 0
            Set usedButtonLabels = Nothing
            Set authoritativeLabelOwners = Nothing
        End If
        If StrComp(groupKey, currentGroup, vbTextCompare) <> 0 Then
            If Len(currentGroup) > 0 Then RibbonAppendXmlLine xmlText, RIBBON_XML_INDENT_4 & "</group>"
            currentGroup = groupKey
            groupButtonCount = 0
            groupSegmentNumber = 1
            currentRibbonGroupLabel = RibbonGroupSegmentLabel(groupLabel, groupSegmentNumber)
            groupStaticTag = RibbonStaticGroupTagDefault(tabKey, groupLabel, groupSortOrder, groupSegmentNumber, systemGroupDefinitions)
            groupID = RibbonGroupId(groupStaticTag)
            ' INTENTIONAL KEYED LOOKUP: usedButtonLabels, authoritativeLabelOwners provides keyed membership/lookup over source rows or state already read by BuildRibbonXml; it is a derived access path, not a second
            ' authority.
            Set usedButtonLabels = New Collection
            Set authoritativeLabelOwners = New Collection
            RegisterRibbonCustomId usedCustomIds, groupID, "Group", tabKey & " / " & currentRibbonGroupLabel
            RibbonAppendXmlLine xmlText, RibbonGroupOpenXml(groupID, groupStaticTag, tabKey, currentRibbonGroupLabel, groupLabel, groupSegmentNumber, usedStaticTags)
        ElseIf groupButtonCount >= RIBBON_MAX_BUTTONS_PER_GROUP Then
            If Len(currentGroup) > 0 Then RibbonAppendXmlLine xmlText, RIBBON_XML_INDENT_4 & "</group>"
            groupButtonCount = 0
            groupSegmentNumber = groupSegmentNumber + 1
            currentRibbonGroupLabel = RibbonGroupSegmentLabel(groupLabel, groupSegmentNumber)
            groupStaticTag = RibbonStaticGroupTagDefault(tabKey, groupLabel, groupSortOrder, groupSegmentNumber, systemGroupDefinitions)
            groupID = RibbonGroupId(groupStaticTag)
            ' INTENTIONAL KEYED LOOKUP: usedButtonLabels, authoritativeLabelOwners provides keyed membership/lookup over source rows or state already read by BuildRibbonXml; it is a derived access path, not a second
            ' authority.
            Set usedButtonLabels = New Collection
            Set authoritativeLabelOwners = New Collection
            RegisterRibbonCustomId usedCustomIds, groupID, "Group", tabKey & " / " & currentRibbonGroupLabel
            RibbonAppendXmlLine xmlText, RibbonGroupOpenXml(groupID, groupStaticTag, tabKey, currentRibbonGroupLabel, groupLabel, groupSegmentNumber, usedStaticTags)
        End If
        currentNormalizedTab = UCase$(tabKey)
        currentNormalizedGroup = UCase$(RibbonDefaultGroupLabel(groupLabel))
        currentNormalizedObjectType = currentObjectType
        currentButtonLabel = CStr(rowData(8))
        If Len(currentButtonLabel) = 0 Then currentButtonLabel = RibbonAuthoritativeButtonTitle(currentObjectType, currentObjectName)
        If Len(currentButtonLabel) = 0 Then Err.Raise vbObjectError + 7154, "BuildRibbonXml", "No authoritative Ribbon button title was resolved." & vbCrLf & "Tab: " & tabKey & vbCrLf & "Group: " & groupLabel & vbCrLf & "Object type: " & currentObjectType & vbCrLf & "Object: " & currentObjectName
        RegisterAuthoritativeRibbonLabel usedButtonLabels, authoritativeLabelOwners, currentButtonLabel, currentNormalizedTab, currentNormalizedGroup, currentNormalizedObjectType, currentObjectName
        defaultTitleCount = defaultTitleCount + 1
        RibbonAppendXmlLine xmlText, RibbonButtonXml(tabKey, tabLabel, currentRibbonGroupLabel, currentObjectType, currentObjectName, currentButtonLabel, usedCustomIds, usedStaticTags)
        groupButtonCount = groupButtonCount + 1
        buttonCount = buttonCount + 1
        RibbonRequireSafeWorkContinue "BuildRibbonXml", currentObjectType & " [" & currentObjectName & "]"
    Next rowData
    EndRibbonButtonTargetInsertBatch
    If Len(currentGroup) > 0 Then RibbonAppendXmlLine xmlText, RIBBON_XML_INDENT_4 & "</group>"
    If Len(currentTab) > 0 Then RibbonAppendXmlLine xmlText, RIBBON_XML_INDENT_3 & "</tab>"
    RibbonAppendXmlLine xmlText, RIBBON_XML_INDENT_2 & "</tabs>"
    RibbonAppendXmlLine xmlText, RIBBON_XML_INDENT_1 & "</ribbon>"
    RibbonAppendXmlLine xmlText, "</customUI>"
    PrintRibbonStatus "Ribbon XML generated. Buttons=" & CStr(buttonCount) & "."
    LogTraceMessage TRACE_LEVEL_INFO, "RibbonTitle", "BuildRibbonXml", "Authoritative Ribbon button titles resolved and validated.", "Titles resolved: " & CStr(defaultTitleCount) & vbCrLf & "UI performed no title abbreviation, contextual rewriting, generated-form source suffixing, or uniqueness repair; the configured canonical System Macro designator remains the only permitted post-default title modification.", "Generate ribbon titles", 0, vbNullString, vbNullString, vbNullString, "RibbonManifest", "CodeDb", "Ribbon"
    BuildRibbonXml = xmlText
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    EndRibbonButtonTargetInsertBatch
    On Error GoTo 0
    Err.Raise savedErrNumber, "BuildRibbonXml", "Could not build " & DEFAULT_UTILITY_LONG_NAME & " custom UI XML." & vbCrLf & "Current tab key: " & tabKey & vbCrLf & "Current tab label: " & IIf(Len(tabLabel) = 0, currentTabLabel, tabLabel) & vbCrLf & "Current group: " & groupLabel & vbCrLf & "Current object: " & currentObjectName & vbCrLf & "Current type: " & currentObjectType & vbCrLf & "Current label: " & currentButtonLabel & vbCrLf & "Source detail: " & RibbonLimitedErrorText(sourceDetailText, 1200) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Sub RibbonAppendXmlLine(ByRef xmlText As String, ByVal lineText As String)
    If Len(xmlText) > 0 Then xmlText = xmlText & vbCrLf
    xmlText = xmlText & lineText
End Sub


Private Function RibbonSourceRows(ByRef sourceDetailText As String) As Collection
    On Error GoTo ErrHandler
    Dim rows As Collection
    Dim stageText As String
    ' INTENTIONAL FILTERED PARTITION: rows is the Ribbon projection of the complete logical Managed Object Catalog; physical Navigation Pane membership is not a Ribbon source or authority.
    Set rows = New Collection
    stageText = "Validate current System Macro Ribbon metadata without physical Navigation Pane mutation"
    ' The final seal reuses the immediately completed metadata validation and rereads mutable source rows without invoking the cross-module responsiveness path.
    If mRibbonSourceSealActive = False Then ValidateSystemRibbonGroupOverridesForRibbonBuild "Read-only Part H ribbon-row collection"
    sourceDetailText = "Ribbon source is the complete logical catalog plus configured Ribbon metadata; physical Navigation Pane state is deferred to Phase I."
    stageText = "Add catalog-managed primary Ribbon placements"
    RibbonAddCatalogManagedObjectRows rows, sourceDetailText
    stageText = "Validate complete catalog Ribbon coverage"
    ValidateRibbonCatalogCoverage rows
    stageText = "Verify one consistent derived group order and globally order tab/group/object rows"
    Set rows = RibbonRowsWithVerifiedGroupOrder(rows)
    Set rows = RibbonGloballyOrderedRows(rows)
    RibbonAppendSourceDetail sourceDetailText, "total ribbon rows=" & CStr(rows.count)
    If rows.count = 0 Then Err.Raise vbObjectError + 6662, "RibbonSourceRows", "No managed-object Ribbon rows were generated."
    Set RibbonSourceRows = rows
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "RibbonSourceRows", "Could not collect the complete " & DEFAULT_UTILITY_LONG_NAME & " managed-object Ribbon inventory from logical catalog metadata." & vbCrLf & "Stage: " & stageText & vbCrLf & "Source detail: " & IIf(Len(sourceDetailText) = 0, "(none)", sourceDetailText) & vbCrLf & "Original Err.Source: " & savedErrSource & vbCrLf & "Original Err.Description: " & savedErrDescription
End Function

Private Function RibbonQueryDefinitionFinding(ByVal db As DAO.Database, ByVal objectName As String) As String
    On Error GoTo ErrHandler
    Dim qdf As DAO.queryDef
    Set qdf = db.QueryDefs(objectName)
    qdf.fields.Refresh
    Set qdf = Nothing
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    Set qdf = Nothing
    RibbonQueryDefinitionFinding = "Cataloged query has an invalid definition or unresolved dependency." & vbCrLf & "Object: " & objectName & vbCrLf & "Error " & CStr(savedErrNumber) & ": " & savedErrDescription
End Function

Private Function RibbonPreflightFindingsText(ByVal findings As Collection) As String
    Dim itemValue As Variant
    Dim findingNumber As Long
    Dim resultText As String
    resultText = "Blocking finding 1 of " & CStr(findings.count) & ":" & vbCrLf & CStr(findings(1))
    If findings.count > 1 Then resultText = resultText & vbCrLf & vbCrLf & "Additional blocking findings: " & CStr(findings.count - 1)
    For findingNumber = 2 To findings.count
        RibbonSetupControlPump
        itemValue = findings(findingNumber)
        resultText = resultText & vbCrLf & vbCrLf & "Finding " & CStr(findingNumber) & " of " & CStr(findings.count) & ":" & vbCrLf & CStr(itemValue)
    Next findingNumber
    RibbonPreflightFindingsText = resultText
End Function

Private Sub LogRibbonPreflightFindings(ByVal findings As Collection, ByVal stageText As String, ByVal rowNumber As Long, ByVal rowCount As Long)
    Dim itemValue As Variant
    Dim findingNumber As Long
    Dim findingText As String
    Dim summaryText As String
    For Each itemValue In findings
        RibbonSetupControlPump
        findingNumber = findingNumber + 1
        findingText = CStr(itemValue)
        summaryText = "Blocking finding " & CStr(findingNumber) & " of " & CStr(findings.count) & " > " & RibbonPreflightImmediateText(findingText)
        LogTraceMessage TRACE_LEVEL_ERROR, "Preflight", "ValidateRibbonManagedObjectInventoryPreflight", summaryText, findingText & vbCrLf & "Stage: " & stageText & vbCrLf & "Catalog row: " & CStr(rowNumber) & " of " & CStr(rowCount), stageText, vbObjectError + 6665, "ValidateRibbonManagedObjectInventoryPreflight", findingText, vbNullString, "ManagedObjectCatalog", "CodeDb", "UI"
    Next itemValue
End Sub

Private Function RibbonPreflightImmediateText(ByVal sourceText As String) As String
    Dim resultText As String
    resultText = Trim$(sourceText)
    resultText = Replace(resultText, vbCrLf, " | ")
    resultText = Replace(resultText, vbLf, " | ")
    resultText = Replace(resultText, vbCr, " | ")
    Do While InStr(1, resultText, "  ", vbBinaryCompare) > 0
        RibbonSetupControlPump
        resultText = Replace(resultText, "  ", " ")
    Loop
    If Len(resultText) > 700 Then resultText = Left$(resultText, 697) & "..."
    RibbonPreflightImmediateText = resultText
End Function

Private Sub RibbonAppendSourceDetail(ByRef sourceDetailText As String, ByVal detailText As String)
    If Len(detailText) = 0 Then Exit Sub
    If Len(sourceDetailText) > 0 Then
        sourceDetailText = sourceDetailText & "; " & detailText
    Else
        sourceDetailText = detailText
    End If
End Sub

Private Function RibbonConfiguredGroupNames() As Collection
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim groups As Collection
    Dim seenGroupNames As Collection
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim groupName As String
    Dim groupSortOrder As Long
    Dim systemGroupFlag As Long
    Dim macroGroupFlag As Long
    Dim groupTabKey As String
    Dim groupTabLabel As String
    Dim groupRibbonLabel As String
    Dim sortBucketExpression As String
    If Len(mRibbonPhaseToken) > 0 Then
        If Not mRibbonConfiguredGroupCache Is Nothing Then
            Set RibbonConfiguredGroupNames = mRibbonConfiguredGroupCache
            Exit Function
        End If
    End If
    ' INTENTIONAL FILTERED PARTITION: groups, seenGroupNames is a consumer-required subset/projection derived from the source enumerated by RibbonConfiguredGroupNames; its membership is not independently
    ' authoritative.
    Set groups = New Collection
    Set seenGroupNames = New Collection
    sortBucketExpression = RibbonConfiguredGroupSortBucketExpression("G.[" & FIELD_RIBBON_TAB_KEY & "]")
    sqlText = "SELECT G.[" & FIELD_NAVIGATION_PANE_GROUP_NAME & "], "
    sqlText = sqlText & "Nz(G." & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER & ",9999) AS [BaseGroupSortOrder], "
    sqlText = sqlText & "Nz(G.[" & FIELD_RIBBON_GROUP_SORT_ORDER & "],Nz(G." & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER & ",9999)) AS [EffectiveRibbonGroupSortOrder], "
    sqlText = sqlText & "Nz(G.[" & FIELD_RIBBON_TAB_KEY & "],'') AS [ConfiguredRibbonTabKey], "
    sqlText = sqlText & "Nz(G.[" & FIELD_RIBBON_TAB_LABEL & "],'') AS [ConfiguredRibbonTabLabel], "
    sqlText = sqlText & "Nz(G.[" & FIELD_RIBBON_GROUP_LABEL & "],'') AS [ConfiguredRibbonGroupLabel], "
    sqlText = sqlText & "IIf(UCase(Nz(G." & SQL_FIELD_GROUP_PURPOSE & ",''))=" & SharedDao_SqlQuote(UCase$(NAVIGATION_PANE_GROUP_PURPOSE_SYSTEM_MACROS)) & ",1,0) AS [RibbonSystemGroupFlag], "
    sqlText = sqlText & "IIf(UCase(Nz(G.[" & FIELD_OBJECT_TYPE_SCOPE & "],''))='MACRO' Or UCase(Nz(G." & SQL_FIELD_GROUP_PURPOSE & ",'')) In (" & SharedDao_SqlQuote(UCase$(NAVIGATION_PANE_GROUP_PURPOSE_SYSTEM_MACROS)) & "," & SharedDao_SqlQuote(UCase$(NAVIGATION_PANE_GROUP_PURPOSE_GENERATED_MACROS)) & "),1,0) AS [RibbonMacroGroupFlag], "
    sqlText = sqlText & sortBucketExpression & " AS [RibbonTabSortBucket], "
    sqlText = sqlText & "Nz(G.[" & FIELD_NAVIGATION_PANE_GROUP_KEY & "],'') AS [ConfiguredNavigationPaneGroupKey], Nz(G." & SQL_FIELD_ID & ",0) AS [NavigationPaneGroupID] "
    sqlText = sqlText & "FROM " & SharedDao_BracketName(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups")) & " AS G "
    sqlText = sqlText & "WHERE Len(Trim(Nz(G.[" & FIELD_NAVIGATION_PANE_GROUP_NAME & "],'')))>0 "
    sqlText = sqlText & "AND Nz(G." & SQL_FIELD_LOOKUP_ACTIVE & ",False)<>False "
    sqlText = sqlText & "AND Nz(G.[" & FIELD_SHOW_ON_RIBBON & "],True)<>False "
    sqlText = sqlText & "AND Nz(G.[" & FIELD_RIBBON_GROUP_VISIBLE & "],True)<>False "
    sqlText = sqlText & "AND UCase(Trim(Nz(G.[" & FIELD_NAVIGATION_PANE_GROUP_NAME & "],'')))<>'ACCESS DEVELOPMENT CONSOLE - SORT' "
    sqlText = sqlText & "AND UCase(Trim(Nz(G.[" & FIELD_RIBBON_GROUP_LABEL & "],'')))<>'SORT' "
    sqlText = sqlText & "ORDER BY " & RibbonConfiguredGroupOrderByText() & ";"
    Set db = GetHostDb()
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    Do While rs.EOF = False
        RibbonSetupControlPump
        groupName = CStr(rs.fields(FIELD_NAVIGATION_PANE_GROUP_NAME).value)
        If RibbonGroupNameAlreadySelected(seenGroupNames, groupName) = False Then
            RibbonMarkGroupNameSelected seenGroupNames, groupName
            groupSortOrder = CLng(rs.fields("EffectiveRibbonGroupSortOrder").value)
            systemGroupFlag = CLng(rs.fields("RibbonSystemGroupFlag").value)
            macroGroupFlag = CLng(rs.fields("RibbonMacroGroupFlag").value)
            groupTabKey = RibbonTabKeyForGroup(groupName, systemGroupFlag, CStr(rs.fields("ConfiguredRibbonTabKey").value))
            groupTabLabel = ResolveRibbonTabLabel(groupTabKey, CStr(rs.fields("ConfiguredRibbonTabLabel").value))
            groupRibbonLabel = CStr(rs.fields("ConfiguredRibbonGroupLabel").value)
            If Len(Trim$(groupRibbonLabel)) = 0 Then groupRibbonLabel = groupName
            groups.Add Array(groupName, groupSortOrder, groupTabKey, groupTabLabel, groupRibbonLabel, systemGroupFlag, macroGroupFlag)
        End If
        rs.MoveNext
    Loop
    rs.Close
    If Len(mRibbonPhaseToken) > 0 Then Set mRibbonConfiguredGroupCache = groups
    Set RibbonConfiguredGroupNames = groups
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
    Err.Raise savedErrNumber, "RibbonConfiguredGroupNames", "Could not read configured Navigation Pane group names for the " & DEFAULT_UTILITY_LONG_NAME & " ribbon." & vbCrLf & "SQL:" & vbCrLf & sqlText & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function RibbonConfiguredGroupFingerprint() As String
    Dim groupRows As Collection
    Dim groupData As Variant
    Dim groupPayload As String
    Set groupRows = RibbonConfiguredGroupNames()
    groupPayload = "PartFRibbonGroupsV2|"
    For Each groupData In groupRows
        groupPayload = groupPayload & FingerprintTokenPart(CStr(groupData(0))) & FingerprintTokenPart(CStr(groupData(1))) & FingerprintTokenPart(CStr(groupData(2))) & FingerprintTokenPart(CStr(groupData(3))) & FingerprintTokenPart(CStr(groupData(4))) & FingerprintTokenPart(CStr(groupData(5))) & FingerprintTokenPart(CStr(groupData(6))) & vbLf
    Next groupData
    RibbonConfiguredGroupFingerprint = AdlerTextHash(groupPayload) & ":" & CStr(Len(groupPayload))
End Function

Private Function RibbonGroupNameAlreadySelected(ByVal selectedGroupNames As Collection, ByVal groupName As String) As Boolean
    On Error GoTo NotFound
    Dim valueText As String
    valueText = selectedGroupNames(RibbonGroupNameKey(groupName))
    RibbonGroupNameAlreadySelected = True
    Exit Function
NotFound:
    RibbonGroupNameAlreadySelected = False
End Function

Private Sub RibbonMarkGroupNameSelected(ByRef selectedGroupNames As Collection, ByVal groupName As String)
    On Error Resume Next
    ' INTENTIONAL FILTERED PARTITION: selectedGroupNames is a consumer-required subset/projection derived from the source enumerated by RibbonMarkGroupNameSelected; its membership is not independently authoritative.
    If selectedGroupNames Is Nothing Then Set selectedGroupNames = New Collection
    If RibbonGroupNameAlreadySelected(selectedGroupNames, groupName) = False Then selectedGroupNames.Add groupName, RibbonGroupNameKey(groupName)
End Sub


Private Function RibbonTabKeyForGroup(ByVal groupName As String, ByVal systemGroupFlag As Long, Optional ByVal configuredTabKey As String = vbNullString) As String
    Dim useSystemTab As Boolean
    configuredTabKey = Trim$(configuredTabKey)
    If StrComp(configuredTabKey, DEFAULT_RIBBON_TAB_KEY_SYSTEM_LOGICAL, vbTextCompare) = 0 Then configuredTabKey = RibbonDefaultTabKeyForGroup(True)
    useSystemTab = (systemGroupFlag <> 0 Or NavigationPaneGroupUsesConfiguredAffixes(groupName, True))
    If Len(configuredTabKey) > 0 Then RibbonTabKeyForGroup = RibbonSafeId(configuredTabKey) Else RibbonTabKeyForGroup = RibbonSafeId(RibbonDefaultTabKeyForGroup(useSystemTab))
    If Len(RibbonTabKeyForGroup) = 0 Then RibbonTabKeyForGroup = RibbonSafeId(RibbonDefaultTabKeyForGroup(False))
End Function

Private Function RibbonRecordsetText(ByVal rs As DAO.Recordset, ByVal fieldName As String) As String
    On Error GoTo ErrHandler
    RibbonRecordsetText = CStr(Nz(rs.fields(fieldName).value, vbNullString))
    Exit Function
ErrHandler:
    RibbonRecordsetText = vbNullString
End Function

Private Function RibbonAddRowIfObjectExists(ByRef rows As Collection, ByVal placementIndex As Object, ByVal groupTabKey As String, ByVal groupTabLabel As String, ByVal groupName As String, ByVal objectTypeCode As String, ByVal objectName As String, ByVal ribbonTooltip As String, ByRef skipReason As String, Optional ByVal groupSortOrder As Long = 9999, Optional ByVal objectSortOrder As Long = 9999, Optional ByVal authoritativeRibbonTitle As String = vbNullString) As Boolean
    Dim normalizedType As String
    Dim identityKey As String
    Dim existingPlacement As Variant
    skipReason = vbNullString
    groupTabKey = RibbonSafeId(groupTabKey)
    If Len(groupTabKey) = 0 Then groupTabKey = RibbonTabKeyForGroup(groupName, 0)
    If Len(groupTabLabel) = 0 Then groupTabLabel = ResolveRibbonTabLabel(groupTabKey)
    If Len(objectName) = 0 Or Len(groupName) = 0 Then
        skipReason = RIBBON_SKIP_REASON_BLANK_INVALID
        Exit Function
    End If
    normalizedType = objectTypeCode
    If Len(normalizedType) = 0 Then
        skipReason = RIBBON_SKIP_REASON_UNSUPPORTED
        Exit Function
    End If
    If RibbonObjectExists(normalizedType, objectName) = False Then
        skipReason = RIBBON_SKIP_REASON_MISSING
        Exit Function
    End If
    identityKey = RibbonCatalogMetadataKey(objectName, normalizedType)
    If placementIndex.Exists(identityKey) Then
        existingPlacement = placementIndex(identityKey)
        If StrComp(CStr(existingPlacement(0)), groupTabKey, vbTextCompare) = 0 And StrComp(CStr(existingPlacement(1)), groupName, vbTextCompare) = 0 Then
            skipReason = RIBBON_SKIP_REASON_DUPLICATE
            Exit Function
        End If
        Err.Raise vbObjectError + 7270, "RibbonAddRowIfObjectExists", "A managed Ribbon object is assigned to more than one Ribbon group. Each Ribbon item must have exactly one authoritative group." & vbCrLf & "Object type: " & normalizedType & vbCrLf & "Object: " & objectName & vbCrLf & "Existing tab/group: " & CStr(existingPlacement(0)) & " / " & CStr(existingPlacement(1)) & vbCrLf & "Incoming tab/group: " & groupTabKey & " / " & groupName
    End If
    placementIndex.Add identityKey, Array(groupTabKey, groupName)
    rows.Add Array(groupTabKey, groupTabLabel, groupName, normalizedType, objectName, ribbonTooltip, groupSortOrder, objectSortOrder, authoritativeRibbonTitle)
    RibbonAddRowIfObjectExists = True
End Function

Private Function RibbonTabOpenXml(ByVal tabKey As String, ByVal tabLabel As String, ByRef usedStaticTags As Collection) As String
    Dim tabID As String
    Dim staticTag As String
    tabID = RibbonRegularTabId(tabKey)
    staticTag = RibbonControlTag(RibbonStaticTabTagDefault(tabKey))
    RegisterRibbonStaticTag usedStaticTags, staticTag, "Tab", tabKey & " / " & tabLabel
    RibbonTabOpenXml = RIBBON_XML_INDENT_3 & "<tab id=""" & RibbonXmlAttributeText(tabID) & """ label=""" & RibbonXmlAttributeText(RibbonVisibleSeparatorDefault(RibbonManagedObjectAbbreviationDefault(ResolveRibbonTabLabel(tabKey, tabLabel)))) & """ tag=""" & RibbonXmlAttributeText(staticTag) & """>"
End Function

Private Function RibbonGroupOpenXml(ByVal groupID As String, ByVal groupStaticTag As String, ByVal tabKey As String, ByVal groupLabel As String, ByVal sourceGroupLabel As String, ByVal groupSegmentNumber As Long, ByRef usedStaticTags As Collection) As String
    Dim controlTag As String
    controlTag = RibbonControlTag(groupStaticTag)
    RegisterRibbonStaticTag usedStaticTags, controlTag, "Group", tabKey & " / " & sourceGroupLabel & " / segment " & CStr(groupSegmentNumber)
    RibbonGroupOpenXml = RIBBON_XML_INDENT_4 & "<group id=""" & RibbonXmlAttributeText(groupID) & """ label=""" & RibbonXmlAttributeText(RibbonManagedObjectAbbreviationDefault(RibbonDefaultGroupLabel(groupLabel))) & """ tag=""" & RibbonXmlAttributeText(controlTag) & """>"
End Function

Private Function RibbonRowsWithVerifiedGroupOrder(ByVal sourceRows As Collection) As Collection
    'Ribbon rows are derived projections. Conflicting orders for the same tab/group are invalid; no source may override another.
    Dim groupOrders As Collection
    Dim rowData As Variant
    Dim groupKey As String
    Dim existingOrder As Long
    Dim currentOrder As Long
    Dim groupFound As Boolean
    Dim rowNumber As Long
    ' INTENTIONAL ORDERED VIEW: groupOrders preserves the deterministic sequence required by RibbonRowsWithVerifiedGroupOrder; keyed-only or unordered storage would not preserve the consumer semantics.
    Set groupOrders = New Collection
    For Each rowData In sourceRows
        rowNumber = rowNumber + 1
        If rowNumber = 1 Or rowNumber Mod 50 = 0 Or rowNumber = sourceRows.count Then RibbonSetupControlPump
        groupKey = "K" & UCase$(CStr(rowData(0))) & "|" & UCase$(CStr(rowData(2)))
        currentOrder = CLng(rowData(6))
        If currentOrder <= 0 Then Err.Raise vbObjectError + 7190, "RibbonRowsWithVerifiedGroupOrder", "A derived Ribbon row has no valid group order." & vbCrLf & "Tab/group: " & groupKey & vbCrLf & "Order: " & CStr(currentOrder)
        groupFound = False
        On Error Resume Next
        existingOrder = CLng(groupOrders(groupKey))
        groupFound = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0
        If groupFound = False Then
            groupOrders.Add currentOrder, groupKey
        ElseIf existingOrder <> currentOrder Then
            Err.Raise vbObjectError + 7191, "RibbonRowsWithVerifiedGroupOrder", "Conflicting derived Ribbon group orders exist for the same tab/group. No source may override another; correct the canonical group definition that produced the inconsistent projection." & vbCrLf & "Tab/group: " & groupKey & vbCrLf & "Orders: " & CStr(existingOrder) & " and " & CStr(currentOrder)
        End If
    Next rowData
    Set RibbonRowsWithVerifiedGroupOrder = sourceRows
End Function

Private Function RibbonGloballyOrderedRows(ByVal sourceRows As Collection) As Collection
    'Stably groups all rows by one global tab order so each tab is contiguous, ADC Settings follows ADC Admin, and ADC Design remains last.
    Dim orderedRows As Collection
    Dim rowData As Variant
    Dim existingRow As Variant
    Dim positionValue As Long
    Dim inserted As Boolean
    Dim sourceRowNumber As Long
    ' INTENTIONAL ORDERED VIEW: orderedRows preserves the deterministic sequence required by RibbonGloballyOrderedRows; keyed-only or unordered storage would not preserve the consumer semantics.
    Set orderedRows = New Collection
    For Each rowData In sourceRows
        sourceRowNumber = sourceRowNumber + 1
        If sourceRowNumber = 1 Or sourceRowNumber Mod 50 = 0 Or sourceRowNumber = sourceRows.count Then RibbonSetupControlPump
        inserted = False
        For positionValue = 1 To orderedRows.count
            existingRow = orderedRows(positionValue)
            If RibbonCompareTabOrder(rowData, existingRow) < 0 Then
                orderedRows.Add rowData, Before:=positionValue
                inserted = True
                Exit For
            End If
        Next positionValue
        If inserted = False Then orderedRows.Add rowData
    Next rowData
    Set RibbonGloballyOrderedRows = orderedRows
End Function

Private Sub ValidateRibbonTabOrder(ByVal rows As Collection)
    'Requires contiguous tab/group blocks, one consistent group order, any present X Combo Forms/X Sort terminal groups after all nonterminal ADC Forms groups, ADC Settings immediately after ADC Admin, and ADC
    'Design as the final nonempty tab.
    Dim rowData As Variant
    Dim seenTabs As Collection
    Dim seenGroups As Collection
    Dim groupOrders As Collection
    Dim tabKey As String
    Dim groupKey As String
    Dim previousTabKey As String
    Dim previousGroupKey As String
    Dim designRowCount As Long
    Dim uiRowCount As Long
    Dim formsRowCount As Long
    Dim xComboFormsGroupOrder As Long
    Dim xSortGroupOrder As Long
    Dim highestOtherFormsGroupOrder As Long
    Dim currentGroupOrder As Long
    Dim existingGroupOrder As Long
    Dim requiredPreviousTabKey As String
    Dim terminalGroups As Variant
    Dim rowNumber As Long
    ' INTENTIONAL ORDERED VIEW: seenTabs, seenGroups, groupOrders preserves the deterministic sequence required by ValidateRibbonTabOrder; keyed-only or unordered storage would not preserve the consumer semantics.
    Set seenTabs = New Collection
    Set seenGroups = New Collection
    Set groupOrders = New Collection
    terminalGroups = RibbonFormsTerminalGroupSequence()
    For Each rowData In rows
        rowNumber = rowNumber + 1
        If rowNumber = 1 Or rowNumber Mod 50 = 0 Or rowNumber = rows.count Then RibbonSetupControlPump
        tabKey = UCase$(CStr(rowData(0)))
        groupKey = tabKey & "|" & UCase$(CStr(rowData(2)))
        currentGroupOrder = CLng(rowData(6))
        If Len(tabKey) = 0 Then Err.Raise vbObjectError + 7168, "ValidateRibbonTabOrder", "A ribbon row has a blank tab key." & vbCrLf & "Object: " & CStr(rowData(4))
        If Len(CStr(rowData(2))) = 0 Then Err.Raise vbObjectError + 7173, "ValidateRibbonTabOrder", "A ribbon row has a blank group label." & vbCrLf & "Tab key: " & tabKey & vbCrLf & "Object: " & CStr(rowData(4))
        If StrComp(tabKey, previousTabKey, vbBinaryCompare) <> 0 Then
            On Error Resume Next
            seenTabs.Add tabKey, tabKey
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                Err.Raise vbObjectError + 7169, "ValidateRibbonTabOrder", "A ribbon tab appears in more than one non-contiguous block." & vbCrLf & "Tab key: " & tabKey & vbCrLf & "Object: " & CStr(rowData(4))
            End If
            On Error GoTo 0
            requiredPreviousTabKey = UCase$(RibbonRequiredPreviousTabKey(tabKey))
            If Len(requiredPreviousTabKey) > 0 And StrComp(previousTabKey, requiredPreviousTabKey, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7177, "ValidateRibbonTabOrder", DEFAULT_RIBBON_TAB_LABEL_UI & " is not immediately after " & DEFAULT_RIBBON_TAB_LABEL_ADMINISTRATION & "." & vbCrLf & "Previous tab key: " & IIf(Len(previousTabKey) = 0, "(none)", previousTabKey)
            previousTabKey = tabKey
            previousGroupKey = vbNullString
        End If
        If StrComp(groupKey, previousGroupKey, vbBinaryCompare) <> 0 Then
            On Error Resume Next
            seenGroups.Add groupKey, groupKey
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                Err.Raise vbObjectError + 7174, "ValidateRibbonTabOrder", "A ribbon group appears in more than one non-contiguous block." & vbCrLf & "Tab/group: " & groupKey & vbCrLf & "Object: " & CStr(rowData(4))
            End If
            On Error GoTo 0
            previousGroupKey = groupKey
        End If
        On Error Resume Next
        existingGroupOrder = CLng(groupOrders(groupKey))
        If Err.Number <> 0 Then
            Err.Clear
            groupOrders.Add currentGroupOrder, groupKey
        ElseIf existingGroupOrder <> currentGroupOrder Then
            On Error GoTo 0
            Err.Raise vbObjectError + 7175, "ValidateRibbonTabOrder", "A ribbon group has conflicting sort-order values after normalization." & vbCrLf & "Tab/group: " & groupKey & vbCrLf & "First order: " & CStr(existingGroupOrder) & vbCrLf & "Current order: " & CStr(currentGroupOrder) & vbCrLf & "Object: " & CStr(rowData(4))
        End If
        On Error GoTo 0
        If tabKey = UCase$(DEFAULT_RIBBON_TAB_KEY_DESIGN) Then designRowCount = designRowCount + 1
        If tabKey = DEFAULT_RIBBON_TAB_KEY_UI Then uiRowCount = uiRowCount + 1
        If tabKey = UCase$(DEFAULT_RIBBON_TAB_KEY_FORMS) Then
            formsRowCount = formsRowCount + 1
            Select Case UCase$(CStr(rowData(2)))
                Case UCase$(CStr(terminalGroups(0)))
                    xComboFormsGroupOrder = currentGroupOrder
                Case UCase$(CStr(terminalGroups(1)))
                    xSortGroupOrder = currentGroupOrder
                Case Else
                    If currentGroupOrder > highestOtherFormsGroupOrder Then highestOtherFormsGroupOrder = currentGroupOrder
            End Select
        End If
    Next rowData
    If uiRowCount = 0 Then Err.Raise vbObjectError + 7178, "ValidateRibbonTabOrder", "The required " & DEFAULT_RIBBON_TAB_LABEL_UI & " ribbon tab has no rows."
    If formsRowCount > 0 Then
        If xComboFormsGroupOrder > 0 Then
            If highestOtherFormsGroupOrder >= xComboFormsGroupOrder Then Err.Raise vbObjectError + 7184, "ValidateRibbonTabOrder", "X Combo Forms is not after every nonterminal " & DEFAULT_RIBBON_TAB_LABEL_FORMS & " group." & vbCrLf & "Highest nonterminal group order: " & CStr(highestOtherFormsGroupOrder) & vbCrLf & "X Combo Forms order: " & CStr(xComboFormsGroupOrder)
        End If
        If xSortGroupOrder > 0 Then
            If highestOtherFormsGroupOrder >= xSortGroupOrder Then Err.Raise vbObjectError + 7184, "ValidateRibbonTabOrder", "X Sort is not after every nonterminal " & DEFAULT_RIBBON_TAB_LABEL_FORMS & " group." & vbCrLf & "Highest nonterminal group order: " & CStr(highestOtherFormsGroupOrder) & vbCrLf & "X Sort order: " & CStr(xSortGroupOrder)
            If xComboFormsGroupOrder > 0 And xComboFormsGroupOrder >= xSortGroupOrder Then Err.Raise vbObjectError + 7184, "ValidateRibbonTabOrder", "X Combo Forms is not before X Sort." & vbCrLf & "X Combo Forms order: " & CStr(xComboFormsGroupOrder) & vbCrLf & "X Sort order: " & CStr(xSortGroupOrder)
        End If
    End If
    If designRowCount = 0 Then Err.Raise vbObjectError + 7171, "ValidateRibbonTabOrder", "The required " & DEFAULT_RIBBON_TAB_LABEL_DESIGN & " ribbon tab has no rows and cannot be generated as the final tab."
    If StrComp(previousTabKey, UCase$(DEFAULT_RIBBON_TAB_KEY_DESIGN), vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7172, "ValidateRibbonTabOrder", DEFAULT_RIBBON_TAB_LABEL_DESIGN & " is not the final ribbon tab." & vbCrLf & "Final tab key: " & previousTabKey
End Sub

Private Sub RegisterAuthoritativeRibbonLabel(ByRef usedButtonLabels As Collection, ByRef labelOwners As Collection, ByVal labelText As String, ByVal normalizedTab As String, ByVal normalizedGroup As String, ByVal normalizedObjectType As String, ByVal objectName As String)
    'Preserves exact authoritative titles while allowing only the confirmed Health Review table/System Health Review query duplicate.
    Dim labelKey As String
    Dim firstOwner As String
    Dim currentOwner As String
    ' INTENTIONAL KEYED LOOKUP: usedButtonLabels, labelOwners provides keyed membership/lookup over source rows or state already read by RegisterAuthoritativeRibbonLabel; it is a derived access path, not a second
    ' authority.
    If usedButtonLabels Is Nothing Then Set usedButtonLabels = New Collection
    If labelOwners Is Nothing Then Set labelOwners = New Collection
    labelKey = RibbonButtonLabelKey(labelText)
    currentOwner = RibbonAuthoritativeOwnerKey(normalizedObjectType, objectName)
    If RibbonRememberUniqueButtonLabel(usedButtonLabels, labelText) Then
        labelOwners.Add currentOwner, labelKey
        Exit Sub
    End If
    On Error Resume Next
    firstOwner = CStr(labelOwners(labelKey))
    On Error GoTo 0
    If RibbonApprovedAuthoritativeDuplicate(normalizedTab, normalizedGroup, labelText, firstOwner, currentOwner) Then
        LogTraceMessage TRACE_LEVEL_INFO, "RibbonTitle", "RegisterAuthoritativeRibbonLabel", "Approved exact duplicate authoritative ribbon title retained.", "Tab: " & normalizedTab & vbCrLf & "Group: " & normalizedGroup & vbCrLf & "Title: " & labelText & vbCrLf & "First object: " & firstOwner & vbCrLf & "Second object: " & currentOwner, "Approve duplicate ribbon title", 0, vbNullString, vbNullString, objectName, normalizedObjectType, "CodeDb", "Ribbon"
        Exit Sub
    End If
    Err.Raise vbObjectError + 7166, "RegisterAuthoritativeRibbonLabel", "An unapproved duplicate authoritative ribbon title was found in the same ribbon group." & vbCrLf & "Tab: " & normalizedTab & vbCrLf & "Group: " & normalizedGroup & vbCrLf & "Title: " & labelText & vbCrLf & "First object: " & IIf(Len(firstOwner) = 0, "(not recorded)", firstOwner) & vbCrLf & "Current object: " & currentOwner
End Sub

Private Function RibbonAuthoritativeOwnerKey(ByVal normalizedObjectType As String, ByVal objectName As String) As String
    'Uses exact case-insensitive object identity; punctuation is significant for approved duplicate ownership.
    RibbonAuthoritativeOwnerKey = UCase$(normalizedObjectType) & "|" & UCase$(objectName)
End Function

Private Function RibbonApprovedAuthoritativeDuplicate(ByVal normalizedTab As String, ByVal normalizedGroup As String, ByVal labelText As String, ByVal firstOwner As String, ByVal currentOwner As String) As Boolean
    'Allows only the exact HealthReview/Table and HealthReview_Current/Query identities to share Review under ADC Admin > Health.
    Dim healthTableOwner As String
    Dim healthQueryOwner As String
    If normalizedTab <> UCase$(DEFAULT_RIBBON_TAB_KEY_ADMINISTRATION) Or normalizedGroup <> UCase$(RibbonManagedGroupLabel("HEALTH")) Then Exit Function
    If StrComp(labelText, "Review", vbBinaryCompare) <> 0 Then Exit Function
    healthTableOwner = RibbonAuthoritativeOwnerKey(ADCAT_OBJECT_TYPE_TABLE, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_HealthReview"))
    healthQueryOwner = RibbonAuthoritativeOwnerKey(ADCAT_OBJECT_TYPE_QUERY, ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_System_HealthReview"))
    RibbonApprovedAuthoritativeDuplicate = _
        ((firstOwner = healthTableOwner And currentOwner = healthQueryOwner) Or _
        (firstOwner = healthQueryOwner And currentOwner = healthTableOwner))
End Function

Private Function RibbonRememberUniqueButtonLabel(ByRef usedButtonLabels As Collection, ByVal labelText As String) As Boolean
    Dim labelKey As String
    labelKey = RibbonButtonLabelKey(labelText)
    If Len(labelKey) = 0 Then Exit Function
    If SharedCollection_KeyExists(usedButtonLabels, labelKey) Then Exit Function
    usedButtonLabels.Add True, labelKey
    RibbonRememberUniqueButtonLabel = True
End Function


Private Function RibbonButtonImageMso(ByVal objectTypeCode As String) As String
    Select Case UCase$(objectTypeCode)
        Case UCase$(ADCAT_OBJECT_TYPE_MACRO)
            RibbonButtonImageMso = RIBBON_MACRO_IMAGE_MSO
        Case UCase$(ADCAT_OBJECT_TYPE_TABLE)
            RibbonButtonImageMso = RIBBON_TABLE_IMAGE_MSO
        Case UCase$(ADCAT_OBJECT_TYPE_QUERY)
            RibbonButtonImageMso = RIBBON_QUERY_IMAGE_MSO
        Case UCase$(ADCAT_OBJECT_TYPE_FORM)
            RibbonButtonImageMso = RIBBON_FORM_IMAGE_MSO
        Case UCase$(ADCAT_OBJECT_TYPE_REPORT)
            RibbonButtonImageMso = RIBBON_REPORT_IMAGE_MSO
        Case Else
            Err.Raise vbObjectError + 7121, "RibbonButtonImageMso", "No ribbon icon is defined because the object type is not Macro, Form, Report, Query, or Table." & vbCrLf & "ObjectType: " & objectTypeCode
    End Select
End Function

Private Sub ValidateAuthoritativeRibbonTitles(ByVal rows As Collection)
    'Validates every authoritative resolved title before icon checks and XML construction so collisions fail immediately.
    Dim rowData As Variant
    Dim currentGroupKey As String
    Dim rowGroupKey As String
    Dim groupButtonCount As Long
    Dim normalizedTab As String
    Dim normalizedGroup As String
    Dim normalizedObjectType As String
    Dim objectName As String
    Dim authoritativeLabel As String
    Dim usedButtonLabels As Collection
    Dim authoritativeLabelOwners As Collection
    Dim rowNumber As Long
    For Each rowData In rows
        rowNumber = rowNumber + 1
        If rowNumber = 1 Or rowNumber Mod 50 = 0 Or rowNumber = rows.count Then RibbonSetupControlPump
        normalizedTab = UCase$(CStr(rowData(0)))
        normalizedGroup = UCase$(RibbonDefaultGroupLabel(CStr(rowData(2))))
        normalizedObjectType = CStr(rowData(3))
        objectName = CStr(rowData(4))
        rowGroupKey = RibbonGroupCompareKey(CStr(rowData(0)), CStr(rowData(2)))
        If StrComp(rowGroupKey, currentGroupKey, vbTextCompare) <> 0 Or groupButtonCount >= RIBBON_MAX_BUTTONS_PER_GROUP Then
            currentGroupKey = rowGroupKey
            groupButtonCount = 0
            ' INTENTIONAL KEYED LOOKUP: usedButtonLabels, authoritativeLabelOwners provides keyed membership/lookup over source rows or state already read by ValidateAuthoritativeRibbonTitles; it is a derived access
            ' path, not a second authority.
            Set usedButtonLabels = New Collection
            Set authoritativeLabelOwners = New Collection
        End If
        authoritativeLabel = CStr(rowData(8))
        If Len(authoritativeLabel) = 0 Then authoritativeLabel = RibbonAuthoritativeButtonTitle(CStr(rowData(3)), objectName)
        If Len(authoritativeLabel) = 0 Then Err.Raise vbObjectError + 7166, "ValidateAuthoritativeRibbonTitles", "No authoritative Ribbon title was resolved." & vbCrLf & "Object: " & objectName & vbCrLf & "Object type: " & normalizedObjectType
        RegisterAuthoritativeRibbonLabel usedButtonLabels, authoritativeLabelOwners, authoritativeLabel, normalizedTab, normalizedGroup, normalizedObjectType, objectName
        groupButtonCount = groupButtonCount + 1
    Next rowData
End Sub

Private Sub ValidateAndLogRibbonObjectIcons(ByVal rows As Collection)
    'Validates every ribbon object identity and writes one aggregate success record instead of one trace row per valid object.
    Dim rowData As Variant
    Dim rowNumber As Long
    Dim verifiedCount As Long
    Dim tabLabel As String
    Dim groupLabel As String
    Dim objectTypeCode As String
    Dim objectName As String
    Dim imageMsoText As String
    For Each rowData In rows
        rowNumber = rowNumber + 1
        If rowNumber = 1 Or rowNumber Mod 50 = 0 Or rowNumber = rows.count Then RibbonSetupControlPump
        tabLabel = CStr(rowData(1))
        groupLabel = CStr(rowData(2))
        objectTypeCode = CStr(rowData(3))
        objectName = CStr(rowData(4))
        If Len(objectTypeCode) = 0 Then Err.Raise vbObjectError + 7122, "ValidateAndLogRibbonObjectIcons", "A ribbon object does not have one of the five supported object identities." & vbCrLf & "Row: " & CStr(rowNumber) & vbCrLf & "Tab: " & tabLabel & vbCrLf & "Group: " & groupLabel & vbCrLf & "Object: " & objectName & vbCrLf & "ObjectType: " & CStr(rowData(3))
        If RibbonObjectExists(objectTypeCode, objectName) = False Then Err.Raise vbObjectError + 7123, "ValidateAndLogRibbonObjectIcons", "A ribbon object's authoritative type does not match a physical Access object." & vbCrLf & "Row: " & CStr(rowNumber) & vbCrLf & "Tab: " & tabLabel & vbCrLf & "Group: " & groupLabel & vbCrLf & "Object: " & objectName & vbCrLf & "ObjectType: " & objectTypeCode
        imageMsoText = RibbonButtonImageMso(objectTypeCode)
        If RibbonAllowedImageMso(imageMsoText) = False Then Err.Raise vbObjectError + 7124, "ValidateAndLogRibbonObjectIcons", "A ribbon object's imageMso is not one of the five approved object icons." & vbCrLf & "Row: " & CStr(rowNumber) & vbCrLf & "Tab: " & tabLabel & vbCrLf & "Group: " & groupLabel & vbCrLf & "Object: " & objectName & vbCrLf & "ObjectType: " & objectTypeCode & vbCrLf & "imageMso: " & imageMsoText
        verifiedCount = verifiedCount + 1
    Next rowData
    LogTraceMessage TRACE_LEVEL_INFO, "RibbonIcon", "ValidateAndLogRibbonObjectIcons", "Ribbon object icon validation completed successfully.", "Rows checked: " & CStr(rowNumber) & vbCrLf & "Rows verified: " & CStr(verifiedCount) & vbCrLf & "Failures: 0" & vbCrLf & "Per-object success trace rows were suppressed; failures retain full row, tab, group, object, type, and imageMso diagnostics.", "Ribbon icon validation complete", 0, vbNullString, vbNullString, vbNullString, "RibbonManifest", "CodeDb", "Ribbon"
End Sub

Private Sub RegisterRibbonCustomId(ByRef usedCustomIds As Collection, ByVal customId As String, ByVal elementType As String, ByVal elementContext As String)
    Dim customIdKey As String
    ' INTENTIONAL KEYED LOOKUP: usedCustomIds provides keyed membership/lookup over source rows or state already read by RegisterRibbonCustomId; it is a derived access path, not a second authority.
    If usedCustomIds Is Nothing Then Set usedCustomIds = New Collection
    If Len(customId) = 0 Then Err.Raise vbObjectError + 7277, "RegisterRibbonCustomId", "A Ribbon element reached XML generation without a CustomID." & vbCrLf & "Element: " & elementType & vbCrLf & "Context: " & elementContext
    customIdKey = UCase$(customId)
    On Error Resume Next
    usedCustomIds.Add elementType & " | " & elementContext, customIdKey
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Err.Raise vbObjectError + 7277, "RegisterRibbonCustomId", "Ribbon generation attempted to emit a duplicate CustomID. AccessUI Ribbon & Tree Builder requires every tab, group, and button CustomID to be unique." & vbCrLf & "CustomID: " & customId & vbCrLf & "Element: " & elementType & vbCrLf & "Context: " & elementContext
    End If
    On Error GoTo 0
End Sub

Private Sub RegisterRibbonStaticTag(ByRef usedStaticTags As Collection, ByVal staticTag As String, ByVal elementType As String, ByVal elementContext As String)
    Dim staticTagKey As String
    ' INTENTIONAL KEYED LOOKUP: usedStaticTags provides keyed membership/lookup over source rows or state already read by RegisterRibbonStaticTag; it is a derived access path, not a second authority.
    If usedStaticTags Is Nothing Then Set usedStaticTags = New Collection
    If Len(staticTag) = 0 Then Err.Raise vbObjectError + 7279, "RegisterRibbonStaticTag", "A Ribbon element reached XML generation without its required static tag." & vbCrLf & "Element: " & elementType & vbCrLf & "Context: " & elementContext
    staticTagKey = UCase$(staticTag)
    On Error Resume Next
    usedStaticTags.Add elementType & " | " & elementContext, staticTagKey
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Err.Raise vbObjectError + 7279, "RegisterRibbonStaticTag", "Ribbon generation attempted to emit the same static tag more than once. Every generated tab, group, and button must have one unique explicitly assigned permanent tag." & vbCrLf & "Tag: " & staticTag & vbCrLf & "Element: " & elementType & vbCrLf & "Context: " & elementContext
    End If
    On Error GoTo 0
End Sub

Private Function RibbonGroupCompareKey(ByVal tabKey As String, ByVal groupLabel As String) As String
    RibbonGroupCompareKey = UCase$(tabKey & RIBBON_TAG_DELIMITER & groupLabel)
End Function

Private Function RibbonGroupId(ByVal staticTag As String) As String
    'Ribbon XML IDs are permanent identity, never order. Keep this derived only from the Defaults-authored permanent group tag.
    If Len(staticTag) = 0 Then Err.Raise vbObjectError + 7257, "RibbonGroupId", "A Ribbon group reached ID generation without its permanent static tag."
    RibbonGroupId = "g_" & staticTag
    If Len(RibbonGroupId) > RIBBON_STANDARD_MAX_ID_ATTRIBUTE_LENGTH Then Err.Raise vbObjectError + 7257, "RibbonGroupId", "Generated regular-ribbon group ID exceeds the approved limit." & vbCrLf & "Tag: " & staticTag & vbCrLf & "ID: " & RibbonGroupId & vbCrLf & "Length: " & CStr(Len(RibbonGroupId)) & vbCrLf & "Limit: " & CStr(RIBBON_STANDARD_MAX_ID_ATTRIBUTE_LENGTH)
End Function

Private Function RibbonButtonId(ByVal staticTag As String, ByVal objectName As String) As String
    'Ribbon XML IDs are permanent identity, never order or placement. Moving, inserting, deleting, or reordering controls must not renumber surviving buttons.
    If Len(staticTag) = 0 Then Err.Raise vbObjectError + 7258, "RibbonButtonId", "A Ribbon button reached ID generation without its permanent static tag." & vbCrLf & "Object: " & objectName
    RibbonButtonId = "b_" & staticTag
    If Len(RibbonButtonId) > RIBBON_STANDARD_MAX_ID_ATTRIBUTE_LENGTH Then Err.Raise vbObjectError + 7259, "RibbonButtonId", "Generated regular-ribbon button ID exceeds the approved limit." & vbCrLf & "Object: " & objectName & vbCrLf & "Tag: " & staticTag & vbCrLf & "ID: " & RibbonButtonId & vbCrLf & "Length: " & CStr(Len(RibbonButtonId)) & vbCrLf & "Limit: " & CStr(RIBBON_STANDARD_MAX_ID_ATTRIBUTE_LENGTH)
End Function

Private Function RibbonStaticButtonKey(ByVal objectTypeCode As String, ByVal objectName As String) As String
    'Returns the permanent tag persisted with the managed-object identity. Ribbon rebuilds never independently calculate or renumber it.
    Dim metadata As Variant
    If mRibbonCatalogMetadataCache Is Nothing Then PrepareRibbonCatalogSourceSnapshot
    If RibbonTryDisplayCatalogMetadata(objectName, objectTypeCode, metadata) = False Then Err.Raise vbObjectError + 7281, "RibbonStaticButtonKey", "No active Managed Object Catalog row exists for this Ribbon target." & vbCrLf & "Object type: " & objectTypeCode & vbCrLf & "Object: " & objectName
    RibbonStaticButtonKey = CStr(metadata(5))
    If Len(RibbonStaticButtonKey) = 0 Then Err.Raise vbObjectError + 7282, "RibbonStaticButtonKey", "The managed-object RibbonStaticTag is blank. Run the current catalog repair before rebuilding the Ribbon." & vbCrLf & "Object type: " & objectTypeCode & vbCrLf & "Object: " & objectName
End Function

Private Function RibbonTargetIsUserOwnedForm(ByVal objectTypeCode As String, ByVal objectName As String, ByRef canonicalObjectTitle As String) As Boolean
    'Consumes the catalog's authoritative ownership classification; Ribbon code does not infer user ownership from form names.
    Dim metadata As Variant
    canonicalObjectTitle = vbNullString
    If StrComp(objectTypeCode, ADCAT_OBJECT_TYPE_FORM, vbTextCompare) <> 0 Then Exit Function
    If mRibbonCatalogMetadataCache Is Nothing Then PrepareRibbonCatalogSourceSnapshot
    If RibbonTryDisplayCatalogMetadata(objectName, objectTypeCode, metadata) = False Then Exit Function
    canonicalObjectTitle = CStr(metadata(8))
    RibbonTargetIsUserOwnedForm = (StrComp(CStr(metadata(7)), DEFAULT_CATALOG_OWNER_USER, vbTextCompare) = 0)
End Function

Private Function RibbonGroupSegmentLabel(ByVal groupLabel As String, ByVal groupSegmentNumber As Long) As String
    Dim labelText As String
    Dim suffixText As String
    labelText = RibbonDefaultGroupLabel(groupLabel)
    If groupSegmentNumber <= 1 Then
        RibbonGroupSegmentLabel = labelText
        Exit Function
    End If
    suffixText = " " & CStr(groupSegmentNumber)
    If Len(labelText) + Len(suffixText) > DEFAULT_RIBBON_TITLE_MAX_LENGTH Then labelText = Left$(labelText, DEFAULT_RIBBON_TITLE_MAX_LENGTH - Len(suffixText) - 3) & "..."
    RibbonGroupSegmentLabel = labelText & suffixText
End Function

Private Function RibbonRegularTabId(ByVal tabKey As String) As String
    'Existing tab IDs use an explicit canonical tab-code mapping and are independent of tab/group/button order; preserve that stable authority.
    RibbonRegularTabId = "t_" & RibbonRegularTabCode(tabKey)
End Function

Private Function RibbonRegularTabCode(ByVal tabKey As String) As String
    Dim safeKey As String
    safeKey = UCase$(tabKey)
    Select Case safeKey
        Case DEFAULT_RIBBON_TAB_KEY_SYSTEM_1
            RibbonRegularTabCode = "A1"
        Case DEFAULT_RIBBON_TAB_KEY_SYSTEM_2
            RibbonRegularTabCode = "A2"
        Case UCase$(DEFAULT_RIBBON_TAB_KEY_DATA)
            RibbonRegularTabCode = "DAT"
        Case UCase$(DEFAULT_RIBBON_TAB_KEY_ADMINISTRATION)
            RibbonRegularTabCode = "ADM"
        Case DEFAULT_RIBBON_TAB_KEY_UI
            RibbonRegularTabCode = "SET"
        Case UCase$(DEFAULT_RIBBON_TAB_KEY_FORMS)
            RibbonRegularTabCode = "FRM"
        Case UCase$(DEFAULT_RIBBON_TAB_KEY_DESIGN)
            RibbonRegularTabCode = "DSG"
        Case UCase$(DEFAULT_RIBBON_TAB_KEY_OBJECTS)
            RibbonRegularTabCode = "OBJ"
        Case Else
            If Len(safeKey) = 0 Then safeKey = "TAB"
            RibbonRegularTabCode = "X_" & Left$(safeKey, 18) & "_" & AdlerTextHash(UCase$(tabKey))
    End Select
End Function

Private Sub InvalidateRibbonSourceSnapshot()
    Set mRibbonPhaseStandardRows = Nothing
    mRibbonPhaseSourceDetail = vbNullString
    mRibbonPhaseRowsPrepared = False
    Set mRibbonCatalogSourceRows = Nothing
    mRibbonCatalogPreflightRowCount = 0
    mRibbonCatalogSnapshotElapsedSeconds = 0
    Set mRibbonSystemMacroTargetCache = Nothing
    Set mRibbonCatalogMetadataCache = Nothing
    Set mRibbonCatalogAliasMetadataCache = Nothing
    Set mRibbonConfiguredGroupCache = Nothing
    Set mRibbonCaptionSourceCache = Nothing
End Sub

Private Sub PrepareRibbonPhaseRows()
    On Error GoTo ErrHandler
    Dim standardRows As Collection
    Dim sourceDetailText As String
    Dim startedTimer As Single
    If mRibbonPhaseRowsPrepared Then Exit Sub
    startedTimer = Timer
    PrintRibbonStatus "Collect and validate the ordered ADC-Ribbon snapshot."
    Set standardRows = RibbonSourceRows(sourceDetailText)
    Set mRibbonPhaseStandardRows = standardRows
    mRibbonPhaseSourceDetail = sourceDetailText & " | ADC ribbon rows=" & CStr(standardRows.count) & _
        " | shared catalog load=" & FormattedElapsedSeconds(mRibbonCatalogSnapshotElapsedSeconds)
    mRibbonPhaseRowsPrepared = True
    LogTraceMessage TRACE_LEVEL_INFO, "RibbonTiming", "PrepareRibbonPhaseRows", "ADC ribbon-row snapshot prepared.", _
        mRibbonPhaseSourceDetail & vbCrLf & "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(startedTimer)), _
        "Prepare ADC ribbon rows", 0, vbNullString, vbNullString, RIBBON_NAME_, "RibbonRows", "CodeDb", "RibbonTiming"
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    InvalidateRibbonSourceSnapshot
    Err.Raise savedErrNumber, "PrepareRibbonPhaseRows", "Could not prepare the " & DEFAULT_UTILITY_LONG_NAME & " ADC ribbon-row snapshot." & vbCrLf & _
        "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Sub

Private Function RibbonPhaseRows(ByRef sourceDetailText As String) As Collection
    PrepareRibbonPhaseRows
    sourceDetailText = mRibbonPhaseSourceDetail
    Set RibbonPhaseRows = mRibbonPhaseStandardRows
End Function

Private Sub RequireRibbonSourceSealCurrent(ByVal sourceProcedure As String, ByVal expectedInputFingerprint As String, ByVal expectedButtonCount As Long, ByVal expectedGroupFingerprint As String)
    ' Final source/dependency seal: rebuilds mutable Ribbon source rows and configured-group state without Ribbon-local event pumping, then returns without another yield.
    Dim priorSealState As Boolean
    Dim currentInputFingerprint As String
    Dim currentButtonCount As Long
    Dim currentGroupFingerprint As String
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    priorSealState = mRibbonSourceSealActive
    On Error GoTo ErrHandler
    mRibbonSourceSealActive = True
    InvalidateRibbonSourceSnapshot
    currentInputFingerprint = RibbonRowsFingerprint(currentButtonCount)
    currentGroupFingerprint = RibbonConfiguredGroupFingerprint()
    If currentButtonCount <> expectedButtonCount Then Err.Raise vbObjectError + 7370, sourceProcedure, "The current ADC Ribbon source-row count changed during final source/dependency sealing." & vbCrLf & "Expected count: " & CStr(expectedButtonCount) & vbCrLf & "Current count: " & CStr(currentButtonCount)
    If StrComp(currentInputFingerprint, expectedInputFingerprint, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7371, sourceProcedure, "The current ADC Ribbon source-row fingerprint changed during final source/dependency sealing."
    If StrComp(currentGroupFingerprint, expectedGroupFingerprint, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7372, sourceProcedure, "The configured ADC Ribbon-group fingerprint changed during final source/dependency sealing."
    mRibbonSourceSealActive = priorSealState
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    mRibbonSourceSealActive = priorSealState
    Err.Raise savedErrNumber, sourceProcedure, savedErrDescription
End Sub

Private Sub RequireRibbonPhaseSourcePlanCurrent(ByVal sourceProcedure As String, ByVal verifyPersistedDecision As Boolean)
    Dim currentInputFingerprint As String
    Dim currentButtonCount As Long
    Dim currentDependencyToken As String
    Dim currentDependencyDetail As String
    Dim currentGroupFingerprint As String
    Dim currentXml As String
    Dim currentStateDetail As String
    Dim currentExactState As Boolean
    InvalidateRibbonSourceSnapshot
    currentInputFingerprint = RibbonRowsFingerprint(currentButtonCount)
    If currentButtonCount <> mRibbonPhaseExpectedStandardCount Then Err.Raise vbObjectError + 7353, sourceProcedure, "The current ADC Ribbon source-row count no longer matches the H1 plan." & vbCrLf & "H1 count: " & CStr(mRibbonPhaseExpectedStandardCount) & vbCrLf & "Current count: " & CStr(currentButtonCount)
    If StrComp(currentInputFingerprint, mRibbonPhaseStandardInputFingerprint, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7354, sourceProcedure, "The current ADC Ribbon source-row fingerprint no longer matches the H1 plan."
    If verifyPersistedDecision Then
        currentExactState = TryRibbonTopLevelFastSkip(currentDependencyToken, currentInputFingerprint, currentButtonCount, currentXml, currentStateDetail, currentGroupFingerprint)
    Else
        currentDependencyToken = CurrentRibbonDependencyToken(currentDependencyDetail, currentGroupFingerprint)
    End If
    If StrComp(currentDependencyToken, mRibbonPhaseDependencyToken, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7355, sourceProcedure, "The current ADC Ribbon dependency state no longer matches the H1 plan." & vbCrLf & IIf(Len(currentDependencyDetail) > 0, currentDependencyDetail, currentStateDetail)
    If verifyPersistedDecision Then
        If currentExactState = mRibbonPhaseStandardRequired Then
            Err.Raise vbObjectError + 7369, sourceProcedure, "The current ADC Ribbon exact-state decision no longer matches the H1 plan." & vbCrLf & "H1 build required: " & CStr(mRibbonPhaseStandardRequired) & vbCrLf & "Current exact state: " & CStr(currentExactState) & vbCrLf & currentStateDetail
        End If
    End If
    RequireRibbonSourceSealCurrent sourceProcedure, mRibbonPhaseStandardInputFingerprint, mRibbonPhaseExpectedStandardCount, mRibbonPhaseGroupFingerprint
End Sub

' Purpose: Resolves the committed Ribbon state token for J1 by combining same-run H3 lineage with fresh mutable-source and persisted exact-state verification.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   prevalidatedPredecessorToken (Input; String): Proof/state token representing the prevalidated predecessor token prerequisite or resulting state; callers use it to prove that coordinated setup work
'   still refers to the same validated state.
'   detailText (Output/in-out; String): Carries human-readable diagnostic detail for the operation; when passed ByRef, the routine can return or augment the text for the caller.
'   prevalidatedCatalogToken (Input Optional; String): Proof/state token representing the prevalidated catalog token prerequisite or resulting state; callers use it to prove that coordinated setup work
'   still refers to the same validated state. Default: vbNullString.
' Returns: String: Returns the proof/state token produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function RibbonCommittedStateTokenForJ1(ByVal prevalidatedPredecessorToken As String, ByRef detailText As String, Optional ByVal prevalidatedCatalogToken As String = vbNullString) As String
    ' J1 read-only proof: uses same-run H3 lineage when available and rereads mutable Ribbon source rows plus committed exact state.
    ' The current Settings-owned automatic-application requirement is applied without mutating CustomRibbonID.
    On Error GoTo NotCurrent
    Dim db As DAO.Database
    Dim propertyExists As Boolean
    Dim dependencyToken As String
    Dim expectedDependencyToken As String
    Dim standardInputFingerprint As String
    Dim standardButtonCount As Long
    Dim currentGroupFingerprint As String
    Dim standardXml As String
    Dim currentRibbonId As String
    Dim applyAfterBuild As Boolean
    Dim payloadText As String
    Dim sameRunDetail As String
    Dim sameRunFailureDetail As String
    Dim sameRunProofUsed As Boolean
    Set db = GetHostDb()
    detailText = vbNullString
    applyAfterBuild = SettingBool(SETTING_APPLY_RIBBON_AFTER_BUILD, DEFAULT_APPLY_RIBBON_AFTER_BUILD)
    If TryRibbonSameRunH3ProofInputs(prevalidatedPredecessorToken, prevalidatedCatalogToken, expectedDependencyToken, sameRunDetail) Then
        InvalidateRibbonSourceSnapshot
        standardInputFingerprint = vbNullString
        standardButtonCount = 0
        If TryRibbonTopLevelFastSkip(dependencyToken, standardInputFingerprint, standardButtonCount, standardXml, detailText, currentGroupFingerprint, prevalidatedPredecessorToken, prevalidatedCatalogToken) Then
            If StrComp(dependencyToken, expectedDependencyToken, vbBinaryCompare) = 0 Then
                sameRunProofUsed = True
            Else
                sameRunFailureDetail = "The current Ribbon dependency token does not match the retained same-run H3 token."
            End If
        Else
            sameRunFailureDetail = detailText
        End If
    Else
        sameRunFailureDetail = sameRunDetail
    End If
    If sameRunProofUsed = False Then
        dependencyToken = vbNullString
        standardInputFingerprint = vbNullString
        standardButtonCount = 0
        standardXml = vbNullString
        detailText = vbNullString
        InvalidateRibbonSourceSnapshot
        If TryRibbonTopLevelFastSkip(dependencyToken, standardInputFingerprint, standardButtonCount, standardXml, detailText, currentGroupFingerprint, prevalidatedPredecessorToken, prevalidatedCatalogToken) = False Then
            If Len(sameRunFailureDetail) > 0 Then detailText = "Same-run H3 proof was unavailable or failed: " & sameRunFailureDetail & vbCrLf & "Full committed-state fallback also failed: " & detailText
            Exit Function
        End If
    End If
    If applyAfterBuild Then
        currentRibbonId = DatabaseTextPropertyValue(db, "CustomRibbonID", propertyExists)
        If propertyExists = False Then detailText = "H3 committed Ribbon state is exact, but ApplyADCRibbonAfterBuild=True and the database CustomRibbonID property is missing.": Exit Function
        If StrComp(currentRibbonId, RIBBON_NAME_, vbTextCompare) <> 0 Then detailText = "H3 committed Ribbon state is exact, but ApplyADCRibbonAfterBuild=True and CustomRibbonID is not set to [" & RIBBON_NAME_ & "]. Current value: [" & currentRibbonId & "].": Exit Function
    End If
    RequireRibbonSourceSealCurrent "RibbonCommittedStateTokenForJ1", standardInputFingerprint, standardButtonCount, currentGroupFingerprint
    If applyAfterBuild Then
        payloadText = "ADCRibbonCommittedStateJ1V1|" & dependencyToken & "|" & standardInputFingerprint & "|" & CStr(standardButtonCount) & "|" & AdlerTextHash(standardXml) & ":" & CStr(Len(standardXml)) & "|ApplyAfterBuild=True|CustomRibbonID=ADC"
    Else
        payloadText = "ADCRibbonCommittedStateJ1V1|" & dependencyToken & "|" & standardInputFingerprint & "|" & CStr(standardButtonCount) & "|" & AdlerTextHash(standardXml) & ":" & CStr(Len(standardXml)) & "|ApplyAfterBuild=False|CustomRibbonIDRequirement=None"
    End If
    RibbonCommittedStateTokenForJ1 = AdlerTextHash(payloadText) & ":" & CStr(Len(payloadText))
    If sameRunProofUsed Then
        detailText = "H3 committed ADC Ribbon exact state and configured Ribbon-application policy are current. ApplyAfterBuild=" & CStr(applyAfterBuild) & ". Same-run H3 lineage was accepted; current source rows, configured-group dependency state, XML, and runtime-manifest evidence were reread and verified."
    Else
        detailText = "H3 committed ADC Ribbon exact state and configured Ribbon-application policy are current. ApplyAfterBuild=" & CStr(applyAfterBuild) & ". Full source/dependency fallback verification was used."
    End If
    Exit Function
NotCurrent:
    If IsCancellationError(Err.Number) Then Err.Raise Err.Number, Err.Source, Err.Description
    detailText = "Could not prove the committed H3 Ribbon state for J1. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    Err.Clear
End Function

Private Sub EndRibbonStatusParts()
    mRibbonProgressActive = False
    mRibbonProgressStep = 0
    mRibbonProgressMax = 0
    mRibbonProgressScopeLabel = vbNullString
    EndTrackedProcess
End Sub

Private Function NewRibbonPhaseToken() As String
    Randomize
    NewRibbonPhaseToken = Format$(Now, "yyyymmddhhnnss") & "-" & Right$("000000" & CStr(Int(999999 * Rnd) + 1), 6)
End Function

Private Sub PersistRibbonPostCommitVerificationFailure(ByVal errorNumber As Long, ByVal errorSource As String, ByVal errorDescription As String)
    'Persists the committed standard-Ribbon verification context before staged in-memory state is cleared.
    Dim detailText As String
    Dim payloadText As String
    detailText = "Committed standard-Ribbon state requires verification retry."
    payloadText = "ExpectedStandardCount=" & CStr(mRibbonPhaseExpectedStandardCount) & vbCrLf
    payloadText = payloadText & "StagedStandardCount=" & CStr(mRibbonPhaseButtonCount) & vbCrLf
    payloadText = payloadText & "StandardXmlHash=" & AdlerTextHash(mRibbonPhaseRibbonXml) & ":" & CStr(Len(mRibbonPhaseRibbonXml)) & vbCrLf
    payloadText = payloadText & "ApplyAfterBuildCaptured=" & CStr(mRibbonPhaseApplyAfterBuildCaptured) & vbCrLf
    If mRibbonPhaseApplyAfterBuildCaptured Then payloadText = payloadText & "ApplyAfterBuild=" & CStr(mRibbonPhaseApplyAfterBuild) & vbCrLf
    payloadText = payloadText & "ErrorNumber=" & CStr(errorNumber) & vbCrLf
    payloadText = payloadText & "ErrorSource=" & errorSource & vbCrLf
    payloadText = payloadText & "ErrorDescription=" & errorDescription
    PersistDurableRecoveryMarker UI_RIBBON_POST_COMMIT_VERIFY_KEY, "CommitCustomRibbonBuild.PostCommitVerify", detailText, payloadText
End Sub

Private Function StagedRibbonTimingText() As String
    StagedRibbonTimingText = "Ribbon phase timing:" & vbCrLf & " - Readiness and fast-path decision: " & FormattedElapsedSeconds(mRibbonPhaseReadinessElapsedSeconds)
    StagedRibbonTimingText = StagedRibbonTimingText & vbCrLf & " - ADC ribbon build/validation: " & IIf(mRibbonPhaseStandardRequired, FormattedElapsedSeconds(mRibbonPhaseStandardElapsedSeconds), "skipped")
    StagedRibbonTimingText = StagedRibbonTimingText & vbCrLf & " - Atomic commit and fingerprint persistence: " & FormattedElapsedSeconds(mRibbonPhaseCommitElapsedSeconds)
    StagedRibbonTimingText = StagedRibbonTimingText & vbCrLf & " - Total: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(mRibbonPhaseStartedTimer))
End Function

Private Function RibbonDiagnosticSettingValue(ByVal db As DAO.Database, ByVal SettingKey As String) As String
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    If Len(SettingKey) = 0 Then
        RibbonDiagnosticSettingValue = "(no effective setting key)"
        Exit Function
    End If
    If SharedDao_TableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID)) = False Then
        RibbonDiagnosticSettingValue = "(Settings table missing)"
        Exit Function
    End If
    If SharedDao_TableFieldExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID), FIELD_SETTING_KEY) = False Or SharedDao_TableFieldExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID), FIELD_SETTING_VALUE) = False Then
        RibbonDiagnosticSettingValue = "(Settings fields missing)"
        Exit Function
    End If
    Set rs = db.OpenRecordset("SELECT TOP 1 [" & FIELD_SETTING_VALUE & "] FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_SETTINGS_TABLE_LOGICAL_ID) & "] WHERE [" & FIELD_SETTING_KEY & "]=" & SharedDao_SqlQuote(SettingKey) & ";", dbOpenSnapshot)
    If rs.EOF Then
        RibbonDiagnosticSettingValue = "(setting row missing)"
    ElseIf IsNull(rs.fields(FIELD_SETTING_VALUE).value) Then
        RibbonDiagnosticSettingValue = "(Null)"
    ElseIf Len(Trim$(CStr(rs.fields(FIELD_SETTING_VALUE).value))) = 0 Then
        RibbonDiagnosticSettingValue = "(blank)"
    Else
        RibbonDiagnosticSettingValue = CStr(rs.fields(FIELD_SETTING_VALUE).value)
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
    Set rs = Nothing
    On Error GoTo 0
    RibbonDiagnosticSettingValue = "(lookup failed: " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription & ")"
End Function

Private Function RibbonDiagnosticNavigationGroupText(ByVal db As DAO.Database, ByVal SettingKey As String) As String
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    If Len(SettingKey) = 0 Then
        RibbonDiagnosticNavigationGroupText = "(no effective setting key)"
        Exit Function
    End If
    If SharedDao_TableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups")) = False Then
        RibbonDiagnosticNavigationGroupText = "(" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups") & " missing)"
        Exit Function
    End If
    Set rs = db.OpenRecordset("SELECT TOP 1 [" & FIELD_NAVIGATION_PANE_GROUP_KEY & "],[" & FIELD_NAVIGATION_PANE_GROUP_NAME & "]," & SQL_FIELD_LOOKUP_ACTIVE & ",[" & FIELD_SHOW_ON_RIBBON & "],[" & FIELD_RIBBON_TAB_KEY & "],[" & FIELD_RIBBON_TAB_LABEL & "],[" & FIELD_RIBBON_GROUP_LABEL & "],[" & FIELD_RIBBON_GROUP_SORT_ORDER & "],[" & FIELD_RIBBON_GROUP_VISIBLE & "] FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups") & "] WHERE [" & FIELD_SETTING_KEY & "]=" & SharedDao_SqlQuote(SettingKey) & ";", dbOpenSnapshot)
    If rs.EOF Then
        RibbonDiagnosticNavigationGroupText = "(group row missing for setting key)"
    Else
        RibbonDiagnosticNavigationGroupText = "Key=" & CStr(Nz(rs.fields(FIELD_NAVIGATION_PANE_GROUP_KEY).value, vbNullString)) & "; Name=" & CStr(Nz(rs.fields(FIELD_NAVIGATION_PANE_GROUP_NAME).value, vbNullString)) & "; LookupActive=" & CStr(Nz(rs.fields(FIELD_LOOKUP_ACTIVE).value, False)) & "; ShowOnRibbon=" & CStr(Nz(rs.fields(FIELD_SHOW_ON_RIBBON).value, False)) & "; TabKey=" & CStr(Nz(rs.fields(FIELD_RIBBON_TAB_KEY).value, vbNullString)) & "; TabLabel=" & CStr(Nz(rs.fields(FIELD_RIBBON_TAB_LABEL).value, vbNullString)) & "; Group=" & CStr(Nz(rs.fields(FIELD_RIBBON_GROUP_LABEL).value, vbNullString)) & "; GroupOrder=" & CStr(Nz(rs.fields(FIELD_RIBBON_GROUP_SORT_ORDER).value, 0)) & "; Visible=" & CStr(Nz(rs.fields(FIELD_RIBBON_GROUP_VISIBLE).value, False))
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
    Set rs = Nothing
    On Error GoTo 0
    RibbonDiagnosticNavigationGroupText = "(lookup failed: " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription & ")"
End Function

Private Function RibbonDiagnosticPhysicalObjectText(ByVal objectTypeCode As String, ByVal objectName As String) As String
    On Error GoTo ErrHandler
    RibbonDiagnosticPhysicalObjectText = "Exists=" & CStr(RibbonObjectExists(objectTypeCode, objectName))
    Exit Function
ErrHandler:
    RibbonDiagnosticPhysicalObjectText = "Check failed: " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
End Function

Private Function RibbonDatabasePropertyDiagnosticText(ByVal db As DAO.Database, ByVal propertyName As String) As String
    On Error Resume Next
    Dim valueText As String
    Err.Clear
    valueText = CStr(Nz(db.properties(propertyName).value, vbNullString))
    If Err.Number <> 0 Then
        RibbonDatabasePropertyDiagnosticText = "(missing; error " & CStr(Err.Number) & ")"
        Err.Clear
    Else
        RibbonDatabasePropertyDiagnosticText = "Length=" & CStr(Len(valueText)) & "; Prefix=" & IIf(Len(valueText) = 0, "(blank)", Left$(valueText, 80))
    End If
    On Error GoTo 0
End Function

Private Sub RequireRibbonRowForApply(ByVal db As DAO.Database, ByVal ribbonName As String)
    If RibbonRowExists(db, ribbonName) Then Exit Sub
    Err.Raise vbObjectError + 6657, "AddTabsToAccessRibbon", "The " & DEFAULT_UTILITY_LONG_NAME & " ADC Ribbon was not found in USysRibbons after the build step." & vbCrLf & "Ribbon: " & ribbonName & vbCrLf & "Database: " & db.Name
End Sub

Private Function RibbonRowExists(ByVal db As DAO.Database, ByVal ribbonName As String) As Boolean
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim sqlText As String
    If SharedDao_TableExists(db, DB_TABLE_USYS_RIBBONS) = False Then Exit Function
    If SharedDao_TableFieldExists(db, DB_TABLE_USYS_RIBBONS, FIELD_RIBBON_NAME) = False Then Exit Function
    sqlText = "SELECT [" & FIELD_RIBBON_NAME & "] FROM " & SharedDao_BracketName(DB_TABLE_USYS_RIBBONS) & " WHERE [" & FIELD_RIBBON_NAME & "]=" & SharedDao_SqlQuote(ribbonName) & ";"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    RibbonRowExists = (rs.EOF = False)
    rs.Close
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
    Err.Raise savedErrNumber, "RibbonRowExists", "Could not verify the " & DEFAULT_UTILITY_LONG_NAME & " ADC Ribbon row." & vbCrLf & "Ribbon: " & ribbonName & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Sub UpsertRibbonButtonTarget(ByVal ribbonName As String, ByVal buttonKey As String, ByVal ribbonTabKey As String, ByVal ribbonTabLabel As String, ByVal ribbonGroupLabel As String, ByVal buttonLabel As String, ByVal screenTipCaption As String, ByVal imageMso As String, ByVal objectTypeCode As String, ByVal objectName As String)
    'The caller clears the selected ribbon manifest before generation, so every target row is appended through one reusable dynaset.
    On Error GoTo ErrHandler
    Dim normalizedTargetIdentity As String
    Dim existingTargetIdentity As String
    If mRibbonTargetInsertRs Is Nothing Or mRibbonStaticTargetOwners Is Nothing Then Err.Raise vbObjectError + 6695, "UpsertRibbonButtonTarget", "The " & DEFAULT_UTILITY_LONG_NAME & " ribbon target append batch is not open."
    normalizedTargetIdentity = UCase$(objectTypeCode) & RIBBON_TAG_DELIMITER & UCase$(objectName)
    If mRibbonStaticTargetOwners.Exists(buttonKey) Then
        existingTargetIdentity = CStr(mRibbonStaticTargetOwners(buttonKey))
        If StrComp(existingTargetIdentity, normalizedTargetIdentity, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7261, "UpsertRibbonButtonTarget", "Static Ribbon tag collision detected; generation stopped before an object could be mapped to another object's persisted tag." & vbCrLf & "Tag: " & buttonKey & vbCrLf & "Existing target: " & existingTargetIdentity & vbCrLf & "Incoming target: " & normalizedTargetIdentity
        Err.Raise vbObjectError + 7271, "UpsertRibbonButtonTarget", "The same static Ribbon tag was emitted more than once. Each Ribbon item must have one unique tag and each managed object may appear only once." & vbCrLf & "Tag: " & buttonKey & vbCrLf & "Target: " & normalizedTargetIdentity & vbCrLf & "Incoming tab/group: " & ribbonTabKey & " / " & ribbonGroupLabel
    Else
        mRibbonStaticTargetOwners.Add buttonKey, normalizedTargetIdentity
    End If
    With mRibbonTargetInsertRs
        .AddNew
        .fields(FIELD_RIBBON_NAME).value = ribbonName
        .fields(FIELD_BUTTON_KEY).value = buttonKey
        .fields(FIELD_RIBBON_TAB_KEY).value = ribbonTabKey
        .fields(FIELD_RIBBON_TAB_LABEL).value = ribbonTabLabel
        .fields(FIELD_RIBBON_GROUP_LABEL).value = ribbonGroupLabel
        .fields(FIELD_BUTTON_LABEL).value = buttonLabel
        .fields(FIELD_SCREEN_TIP_CAPTION).value = screenTipCaption
        .fields(FIELD_IMAGE_MSO).value = imageMso
        .fields(FIELD_OBJECT_TYPE_CODE).value = objectTypeCode
        .fields(FIELD_OBJECT_NAME).value = objectName
        .fields(FIELD_UPDATED_AT).value = Now()
        .Update
    End With
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "UpsertRibbonButtonTarget", "Could not append an " & DEFAULT_UTILITY_LONG_NAME & " ribbon runtime-manifest row through the active batch." & vbCrLf & "Ribbon: " & ribbonName & vbCrLf & "Button key: " & buttonKey & vbCrLf & "Tab: " & ribbonTabLabel & vbCrLf & "Group: " & ribbonGroupLabel & vbCrLf & "Object: " & objectName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub ValidateRibbonGeneratedTitleBeforeManifestPersist(ByVal tabKey As String, ByVal objectTypeCode As String, ByVal objectName As String, ByVal generatedLabel As String)
    Dim expectedTitle As String
    expectedTitle = RibbonAuthoritativeButtonTitle(objectTypeCode, objectName)
    If Len(expectedTitle) = 0 Then Err.Raise vbObjectError + 7241, "ValidateRibbonGeneratedTitleBeforeManifestPersist", "No authoritative Ribbon button title was resolved before manifest persistence." & vbCrLf & "Tab: " & tabKey & vbCrLf & "Object type: " & objectTypeCode & vbCrLf & "Object: " & objectName
    If StrComp(generatedLabel, expectedTitle, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7241, "ValidateRibbonGeneratedTitleBeforeManifestPersist", "The generated Ribbon title does not match the authoritative title before manifest persistence." & vbCrLf & "Tab: " & tabKey & vbCrLf & "Object type: " & objectTypeCode & vbCrLf & "Object: " & objectName & vbCrLf & "Expected title: " & expectedTitle & vbCrLf & "Actual title: " & generatedLabel
End Sub

Private Sub ValidateRibbonDependencies()
    If SharedDao_TableExists(GetHostDb(), ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups")) = False Then Err.Raise vbObjectError + 6655, "ValidateRibbonDependencies", "The Navigation Pane groups table is missing: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups")
    RequireRibbonNavigationPaneGroupField FIELD_NAVIGATION_PANE_GROUP_KEY
    RequireRibbonNavigationPaneGroupField FIELD_NAVIGATION_PANE_GROUP_NAME
    RequireRibbonNavigationPaneGroupField FIELD_SETTING_KEY
    RequireRibbonNavigationPaneGroupField FIELD_OBJECT_TYPE_SCOPE
    RequireRibbonNavigationPaneGroupField FIELD_GROUP_PURPOSE
    RequireRibbonNavigationPaneGroupField FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER
    RequireRibbonNavigationPaneGroupField FIELD_LOOKUP_ACTIVE
    RequireRibbonNavigationPaneGroupField FIELD_SHOW_ON_RIBBON
    RequireRibbonNavigationPaneGroupField FIELD_RIBBON_TAB_KEY
    RequireRibbonNavigationPaneGroupField FIELD_RIBBON_TAB_LABEL
    RequireRibbonNavigationPaneGroupField FIELD_RIBBON_GROUP_LABEL
    RequireRibbonNavigationPaneGroupField FIELD_RIBBON_GROUP_SORT_ORDER
    RequireRibbonNavigationPaneGroupField FIELD_RIBBON_GROUP_VISIBLE
End Sub

Private Function RibbonControlTag(ByVal buttonTargetKey As String) As String
    RibbonControlTag = RIBBON_CONTROL_TAG_STANDARD_PREFIX & RIBBON_TAG_DELIMITER & buttonTargetKey & RIBBON_ACCESSUI_TAG_TERMINATOR
End Function


Private Function RibbonTargetMapFingerprintForRibbon(ByVal db As DAO.Database, ByVal ribbonName As String, ByRef buttonCount As Long) As String
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim payloadText As String
    Dim objectTypeCode As String
    Dim objectName As String
    Dim actualScreenTipCaption As String
    Dim expectedScreenTipCaption As String
    Dim targetMetadata As Variant
    Dim targetTableNames As String
    Dim targetQueryNames As String
    Dim systemMacroKey As String
    sqlText = "SELECT [" & FIELD_RIBBON_NAME & "]," & SQL_FIELD_BUTTON_KEY & ",[" & FIELD_RIBBON_TAB_KEY & "],[" & FIELD_RIBBON_TAB_LABEL & "],[" & FIELD_RIBBON_GROUP_LABEL & "]," & SQL_FIELD_BUTTON_LABEL & "," & SQL_FIELD_SCREEN_TIP_CAPTION & "," & SQL_FIELD_IMAGE_MSO & "," & SQL_FIELD_OBJECT_TYPE_CODE & "," & SQL_FIELD_OBJECT_NAME & " FROM " & SharedDao_BracketName(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_RIBBON_BUTTON_TARGETS_TABLE_LOGICAL_ID))
    sqlText = sqlText & " WHERE [" & FIELD_RIBBON_NAME & "]=" & SharedDao_SqlQuote(ribbonName)
    sqlText = sqlText & " ORDER BY " & SQL_FIELD_BUTTON_KEY & ";"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    Do While rs.EOF = False
        objectTypeCode = NormalizedObjectTypeCode(CStr(Nz(rs.fields(FIELD_OBJECT_TYPE_CODE).value, vbNullString)))
        objectName = CStr(Nz(rs.fields(FIELD_OBJECT_NAME).value, vbNullString))
        actualScreenTipCaption = CStr(Nz(rs.fields(FIELD_SCREEN_TIP_CAPTION).value, vbNullString))
        If StrComp(objectTypeCode, ADCAT_OBJECT_TYPE_MACRO, vbTextCompare) = 0 Then
            If mRibbonSystemMacroTargetCache Is Nothing Then PrepareRibbonCatalogSourceSnapshot
            If mRibbonSystemMacroTargetCache.Exists(objectName) Then
                targetMetadata = mRibbonSystemMacroTargetCache(objectName)
                targetTableNames = CStr(targetMetadata(0))
                targetQueryNames = CStr(targetMetadata(1))
                systemMacroKey = CStr(targetMetadata(2))
                expectedScreenTipCaption = RibbonGeneratedScreenTipCaption(objectTypeCode, objectName)
                If StrComp(actualScreenTipCaption, expectedScreenTipCaption, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7309, "RibbonTargetMapFingerprintForRibbon", "The persisted canonical System Macro ScreenTip is not semantically current and cannot be accepted by the Ribbon fast path." & vbCrLf & "System Macro key: " & systemMacroKey & vbCrLf & "System Macro title: " & objectName & vbCrLf & "Expected ScreenTip: " & expectedScreenTipCaption & vbCrLf & "Actual ScreenTip: " & actualScreenTipCaption & vbCrLf & "Target tables: " & IIf(Len(targetTableNames) = 0, "(none)", targetTableNames) & vbCrLf & "Target queries: " & IIf(Len(targetQueryNames) = 0, "(none)", targetQueryNames)
            End If
        End If
        payloadText = payloadText & FingerprintTokenPart(CStr(Nz(rs.fields(FIELD_RIBBON_NAME).value, vbNullString)))
        payloadText = payloadText & FingerprintTokenPart(CStr(Nz(rs.fields(FIELD_BUTTON_KEY).value, vbNullString)))
        payloadText = payloadText & FingerprintTokenPart(CStr(Nz(rs.fields(FIELD_RIBBON_TAB_KEY).value, vbNullString)))
        payloadText = payloadText & FingerprintTokenPart(CStr(Nz(rs.fields(FIELD_RIBBON_TAB_LABEL).value, vbNullString)))
        payloadText = payloadText & FingerprintTokenPart(CStr(Nz(rs.fields(FIELD_RIBBON_GROUP_LABEL).value, vbNullString)))
        payloadText = payloadText & FingerprintTokenPart(CStr(Nz(rs.fields(FIELD_BUTTON_LABEL).value, vbNullString)))
        payloadText = payloadText & FingerprintTokenPart(CStr(Nz(rs.fields(FIELD_SCREEN_TIP_CAPTION).value, vbNullString)))
        payloadText = payloadText & FingerprintTokenPart(CStr(Nz(rs.fields(FIELD_IMAGE_MSO).value, vbNullString)))
        payloadText = payloadText & FingerprintTokenPart(CStr(Nz(rs.fields(FIELD_OBJECT_TYPE_CODE).value, vbNullString)))
        payloadText = payloadText & FingerprintTokenPart(CStr(Nz(rs.fields(FIELD_OBJECT_NAME).value, vbNullString))) & vbLf
        buttonCount = buttonCount + 1
        rs.MoveNext
    Loop
    rs.Close
    RibbonTargetMapFingerprintForRibbon = AdlerTextHash(payloadText) & ":" & CStr(Len(payloadText))
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
    Err.Raise savedErrNumber, "RibbonTargetMapFingerprintForRibbon", "Could not fingerprint one persisted " & DEFAULT_UTILITY_LONG_NAME & " ribbon runtime manifest." & vbCrLf & "Ribbon: " & ribbonName & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function CurrentRibbonDependencyToken(ByRef detailText As String, ByRef resolvedGroupFingerprint As String, Optional ByVal prevalidatedPredecessorToken As String = vbNullString, Optional ByVal prevalidatedCatalogToken As String = vbNullString, Optional ByVal captureStagedEvidence As Boolean = False) As String
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim propertyExists As Boolean
    Dim catalogToken As String
    Dim systemMacroFingerprint As String
    Dim predecessorToken As String
    Dim predecessorDetail As String
    Dim groupFingerprint As String
    Dim payloadText As String
    Set db = GetHostDb()
    If Len(prevalidatedCatalogToken) > 0 Then
        catalogToken = prevalidatedCatalogToken
    Else
        catalogToken = RibbonCatalogValidationToken()
    End If
    systemMacroFingerprint = DatabaseTextPropertyValue(db, SYSTEM_MACRO_SET_FINGERPRINT_PROPERTY, propertyExists)
    If propertyExists = False Or Len(Trim$(systemMacroFingerprint)) = 0 Then Err.Raise vbObjectError + 7225, "CurrentRibbonDependencyToken", "The validated System Macro and Help Document Macro-set fingerprint is missing."
    If Len(prevalidatedPredecessorToken) > 0 Then
        predecessorToken = prevalidatedPredecessorToken
        predecessorDetail = "The caller supplied an already-proven complete Ribbon predecessor token from the authoritative orchestration path; duplicate predecessor verification was skipped."
    Else
        predecessorToken = RibbonDependencyTokenForSetup(predecessorDetail)
        If Len(predecessorToken) = 0 Then Err.Raise vbObjectError + 7242, "CurrentRibbonDependencyToken", predecessorDetail
    End If
    groupFingerprint = RibbonConfiguredGroupFingerprint()
    resolvedGroupFingerprint = groupFingerprint
    payloadText = FingerprintTokenPart(RIBBON_FINGERPRINT_VERSION)
    payloadText = payloadText & FingerprintTokenPart(RIBBON_FINGERPRINT_VERSION)
    payloadText = payloadText & FingerprintTokenPart(catalogToken)
    payloadText = payloadText & FingerprintTokenPart(systemMacroFingerprint)
    payloadText = payloadText & FingerprintTokenPart(predecessorToken)
    payloadText = payloadText & FingerprintTokenPart(groupFingerprint)
    payloadText = payloadText & FingerprintTokenPart(CStr(IncludeSystemMacroDesignatorInRibbonTitle()))
    CurrentRibbonDependencyToken = AdlerTextHash(payloadText) & ":" & CStr(Len(payloadText))
    If captureStagedEvidence Then
        mRibbonPhaseCatalogToken = catalogToken
        mRibbonPhasePredecessorToken = predecessorToken
        mRibbonPhaseGroupFingerprint = groupFingerprint
    End If
    detailText = "Complete catalog=current; System Macro and Help Document Macro fingerprint=current; D4/E3/G1 predecessor state=current; Ribbon group configuration=current; physical Navigation Pane state is not a Ribbon dependency."
    Exit Function
ErrHandler:
    If IsCancellationError(Err.Number) Then Err.Raise Err.Number, Err.Source, Err.Description
    Err.Raise Err.Number, "CurrentRibbonDependencyToken", "Could not calculate the Part G top-level Ribbon dependency token." & vbCrLf & Err.Description
End Function

Private Sub PersistRibbonDependencyToken()
    On Error GoTo PersistWarning
    Dim db As DAO.Database
    Dim detailText As String
    Dim groupFingerprint As String
    If Len(mRibbonPhaseDependencyToken) = 0 Then mRibbonPhaseDependencyToken = CurrentRibbonDependencyToken(detailText, groupFingerprint, vbNullString, vbNullString, True)
    Set db = GetHostDb()
    Call EnsureDaoPropertyValue(db, RIBBON_DEPENDENCY_TOKEN_PROPERTY, dbText, mRibbonPhaseDependencyToken, "PersistRibbonDependencyToken")
    Exit Sub
PersistWarning:
    LogTraceMessage TRACE_LEVEL_WARNING, "Ribbon", "PersistRibbonDependencyToken", "Ribbon state was verified, but the optional top-level dependency token could not be persisted.", "Err.Number: " & CStr(Err.Number) & vbCrLf & "Err.Source: " & Err.Source & vbCrLf & "Err.Description: " & Err.Description, "Persist ribbon dependency token", Err.Number, Err.Source, Err.Description, RIBBON_NAME_, "Ribbon", "CodeDb", "Ribbon"
    Err.Clear
End Sub

Private Sub BeginRibbonStatusParts(ByVal maxParts As Long)
    BeginTrackedProcess "Build " & DEFAULT_UTILITY_LONG_NAME & " ADC Ribbon"
    mRibbonProgressActive = True
    mRibbonProgressScopeLabel = "Ribbon Build"
    mRibbonProgressStep = 0
    mRibbonProgressMax = maxParts
End Sub

Private Function RibbonFailureDiagnosticSnapshot(ByVal sourceProcedure As String, ByVal savedErrNumber As Long, ByVal savedErrSource As String, ByVal savedErrDescription As String, ByVal failureStageText As String) As String
    On Error GoTo SnapshotErr
    Dim db As DAO.Database
    Dim hostDb As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim resultText As String
    Dim managedObjectsExists As Boolean
    Dim navigationGroupsExists As Boolean
    Dim ribbonTableExists As Boolean
    Dim targetTableExists As Boolean
    Dim managedObjectFieldsExists As Boolean
    Dim managedObjectFieldValuesExists As Boolean
    Dim snapshotErrNumber As Long
    Dim snapshotErrSource As String
    Dim snapshotErrDescription As String
    resultText = "Ribbon failure diagnostic snapshot:"
    resultText = resultText & vbCrLf & "Source procedure: " & sourceProcedure
    resultText = resultText & vbCrLf & "Failure stage: " & IIf(Len(Trim$(failureStageText)) = 0, "(not recorded)", failureStageText)
    resultText = resultText & vbCrLf & "Original error: " & CStr(savedErrNumber) & " from " & IIf(Len(Trim$(savedErrSource)) = 0, "(blank source)", savedErrSource)
    resultText = resultText & vbCrLf & "Original description: " & IIf(Len(Trim$(savedErrDescription)) = 0, "(blank description)", savedErrDescription)
    resultText = resultText & vbCrLf & "Phase flags: Prepared=" & CStr(mRibbonPhasePrepared) & "; BuildRequired=" & CStr(mRibbonPhaseBuildRequired) & "; StandardRequired=" & CStr(mRibbonPhaseStandardRequired) & "; StandardBuilt=" & CStr(mRibbonPhaseStandardBuilt) & "; FastPathUsed=" & CStr(mRibbonPhaseFastPathUsed)
    resultText = resultText & vbCrLf & "Recovery flags: TransactionStarted=" & CStr(mRibbonPhaseTransactionStarted) & "; DurableRecoveryOwned=" & CStr(mUIDurableRecoveryOwned) & "; RecoveryBlocked=" & CStr(mUIRecoveryBlocked)
    resultText = resultText & vbCrLf & "Button counts: Standard generated=" & CStr(mRibbonPhaseButtonCount) & "; Standard expected=" & CStr(mRibbonPhaseExpectedStandardCount)
    If Len(Trim$(mRibbonPhaseReadinessDetail)) > 0 Then resultText = resultText & vbCrLf & "Readiness detail: " & Left$(Replace(mRibbonPhaseReadinessDetail, vbCrLf, " | "), 4000)
    If Len(Trim$(mRibbonPhaseSourceDetail)) > 0 Then resultText = resultText & vbCrLf & "Ribbon-row source detail: " & Left$(Replace(mRibbonPhaseSourceDetail, vbCrLf, " | "), 6000)
    resultText = resultText & vbCrLf & "CodeDb: " & codeDb.Name
    Set hostDb = GetHostDb()
    resultText = resultText & vbCrLf & "HostDb: " & hostDb.Name
    Set db = ADCDb()
    managedObjectsExists = SharedDao_TableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID))
    navigationGroupsExists = SharedDao_TableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups"))
    ribbonTableExists = SharedDao_TableExists(hostDb, DB_TABLE_USYS_RIBBONS)
    targetTableExists = SharedDao_TableExists(hostDb, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_RIBBON_BUTTON_TARGETS_TABLE_LOGICAL_ID))
    managedObjectFieldsExists = SharedDao_TableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFields"))
    managedObjectFieldValuesExists = SharedDao_TableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFieldValues"))
    resultText = resultText & vbCrLf & "Required tables: MAC=" & CStr(managedObjectsExists) & "; UI_Groups=" & CStr(navigationGroupsExists) & "; MAC_Fields=" & CStr(managedObjectFieldsExists) & "; MAC_Fields_Vals=" & CStr(managedObjectFieldValuesExists) & "; USysRibbons=" & CStr(ribbonTableExists) & "; UI_Ribbon_Generated_BtnTargets=" & CStr(targetTableExists)
    If managedObjectsExists Then
        resultText = resultText & RibbonManagedObjectPlacementDiagnosticText(db)
        resultText = resultText & vbCrLf & "Catalog-owned complete placement preflight snapshot:" & vbCrLf & ManagedObjectPlacementDiagnosticText(50)
        resultText = resultText & RibbonCatalogFingerprintDiagnosticText(db)
        resultText = resultText & vbCrLf & "Authoritative ADC1/ADC2 System Macro and Help Document Macro placement snapshot:" & vbCrLf & SystemRibbonGroupOverrideDiagnosticText(50)
    End If
    If navigationGroupsExists Then
        If SharedDao_TableFieldExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups"), FIELD_LOOKUP_ACTIVE) And SharedDao_TableFieldExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups"), FIELD_SHOW_ON_RIBBON) And SharedDao_TableFieldExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups"), FIELD_RIBBON_GROUP_VISIBLE) Then
            Set rs = db.OpenRecordset("SELECT Count(*) AS [ConfiguredCount], Sum(IIf(Nz([" & FIELD_SHOW_ON_RIBBON & "],True)<>False AND Nz([" & FIELD_RIBBON_GROUP_VISIBLE & "],True)<>False,1,0)) AS [VisibleCount] FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups") & "] WHERE Nz(" & SQL_FIELD_LOOKUP_ACTIVE & ",True)<>False;", dbOpenSnapshot)
            resultText = resultText & vbCrLf & "Navigation Pane ribbon-group summary: Active configured=" & CStr(rs.fields("ConfiguredCount").value) & "; Visible=" & CStr(Nz(rs.fields("VisibleCount").value, 0))
            rs.Close
            Set rs = Nothing
        Else
            resultText = resultText & vbCrLf & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups") & " ribbon-control fields are incomplete."
        End If
    End If
    If ribbonTableExists Then
        If SharedDao_TableFieldExists(hostDb, DB_TABLE_USYS_RIBBONS, FIELD_RIBBON_NAME) And SharedDao_TableFieldExists(hostDb, DB_TABLE_USYS_RIBBONS, FIELD_RIBBON_XML) Then
            sqlText = "SELECT [" & FIELD_RIBBON_NAME & "], Len(Nz([" & FIELD_RIBBON_XML & "],'')) AS [XmlLength] FROM [" & DB_TABLE_USYS_RIBBONS & "] WHERE [" & FIELD_RIBBON_NAME & "]=" & SharedDao_SqlQuote(RIBBON_NAME_) & ";"
            Set rs = hostDb.OpenRecordset(sqlText, dbOpenSnapshot)
            If rs.EOF Then resultText = resultText & vbCrLf & "Persisted ADC ribbon row is missing."
            Do While rs.EOF = False
                resultText = resultText & vbCrLf & "Persisted ribbon row: Name=" & CStr(rs.fields(FIELD_RIBBON_NAME).value) & "; XmlLength=" & CStr(rs.fields("XmlLength").value)
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing
        Else
            resultText = resultText & vbCrLf & "USysRibbons required fields are incomplete."
        End If
    End If
    If targetTableExists Then
        If SharedDao_TableFieldExists(hostDb, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_RIBBON_BUTTON_TARGETS_TABLE_LOGICAL_ID), FIELD_RIBBON_NAME) Then
            Set rs = hostDb.OpenRecordset("SELECT [" & FIELD_RIBBON_NAME & "], Count(*) AS [TargetCount] FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_RIBBON_BUTTON_TARGETS_TABLE_LOGICAL_ID) & "] GROUP BY [" & FIELD_RIBBON_NAME & "] ORDER BY " & OrderByPolicyText("UI_RIBBON_FAILURE_DIAGNOSTIC") & ";", dbOpenSnapshot)
            If rs.EOF Then resultText = resultText & vbCrLf & "Ribbon target-map rows: none."
            Do While rs.EOF = False
                resultText = resultText & vbCrLf & "Ribbon target-map row count: Ribbon=" & CStr(Nz(rs.fields(FIELD_RIBBON_NAME).value, vbNullString)) & "; Targets=" & CStr(rs.fields("TargetCount").value)
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing
        Else
            resultText = resultText & vbCrLf & "UI_Ribbon_Generated_BtnTargets is missing the RibbonName field."
        End If
    End If
    RibbonFailureDiagnosticSnapshot = resultText
    Exit Function
SnapshotErr:
    snapshotErrNumber = Err.Number
    snapshotErrSource = Err.Source
    snapshotErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    On Error GoTo 0
    resultText = resultText & vbCrLf & "Diagnostic snapshot collection failed. Error " & CStr(snapshotErrNumber) & " from " & snapshotErrSource & ": " & snapshotErrDescription
    RibbonFailureDiagnosticSnapshot = resultText
End Function

Private Function RibbonParseControlTag(ByVal controlTag As String, ByRef ribbonName As String, ByRef buttonKey As String) As Boolean
    Dim delimiterPosition As Long
    Dim prefixText As String
    controlTag = Trim$(controlTag)
    Do While Len(controlTag) > 0 And StrComp(Right$(controlTag, 1), RIBBON_ACCESSUI_TAG_TERMINATOR, vbBinaryCompare) = 0
        controlTag = RTrim$(Left$(controlTag, Len(controlTag) - Len(RIBBON_ACCESSUI_TAG_TERMINATOR)))
    Loop
    ribbonName = vbNullString
    buttonKey = vbNullString
    If Len(controlTag) = 0 Then Exit Function
    delimiterPosition = InStr(1, controlTag, RIBBON_TAG_DELIMITER, vbBinaryCompare)
    If delimiterPosition > 0 Then
        prefixText = Trim$(Left$(controlTag, delimiterPosition - 1))
        If StrComp(prefixText, RIBBON_CONTROL_TAG_STANDARD_PREFIX, vbTextCompare) <> 0 Then Exit Function
        buttonKey = Trim$(Mid$(controlTag, delimiterPosition + Len(RIBBON_TAG_DELIMITER)))
    Else
        buttonKey = controlTag
    End If
    ribbonName = RIBBON_NAME_
    RibbonParseControlTag = (Len(buttonKey) > 0)
End Function

Private Function RibbonButtonXml(ByVal tabKey As String, ByVal tabLabel As String, ByVal groupLabel As String, ByVal objectTypeCode As String, ByVal objectName As String, ByVal buttonLabel As String, ByRef usedCustomIds As Collection, ByRef usedStaticTags As Collection) As String
    Dim xmlText As String
    Dim labelText As String
    Dim buttonID As String
    Dim buttonTargetKey As String
    Dim controlTag As String
    Dim imageMsoText As String
    Dim screenTipCaption As String
    Dim resolvedObjectTypeCode As String
    Dim canonicalObjectTitle As String
    resolvedObjectTypeCode = objectTypeCode
    If Len(resolvedObjectTypeCode) = 0 Then Err.Raise vbObjectError + 7120, "RibbonButtonXml", "A Ribbon row reached XML generation without one of the five supported object types." & vbCrLf & "Object: " & objectName & vbCrLf & "ObjectType: " & objectTypeCode
    labelText = buttonLabel
    If Len(labelText) = 0 Then Err.Raise vbObjectError + 7121, "RibbonButtonXml", "A Ribbon row reached XML generation without its authoritative button title." & vbCrLf & "Object: " & objectName
    imageMsoText = RibbonButtonImageMso(resolvedObjectTypeCode)
    buttonTargetKey = RibbonStaticButtonKey(resolvedObjectTypeCode, objectName)
    If RibbonTargetIsUserOwnedForm(resolvedObjectTypeCode, objectName, canonicalObjectTitle) Then
        ' User-owned form CustomID consumes the same ObjectDisplayName-derived key persisted as RibbonStaticTag.
        ' The N|key; XML tag wrapper remains separate because the existing callback parser requires that contract.
        buttonID = buttonTargetKey
        If Len(buttonID) > RIBBON_STANDARD_MAX_ID_ATTRIBUTE_LENGTH Then Err.Raise vbObjectError + 7311, "RibbonButtonXml", "A user-owned form's canonical ObjectDisplayName produces a Ribbon CustomID longer than the approved Ribbon limit." & vbCrLf & "Form: " & objectName & vbCrLf & "ObjectDisplayName: " & canonicalObjectTitle & vbCrLf & "Derived Ribbon key: " & buttonTargetKey & vbCrLf & "Length: " & CStr(Len(buttonID)) & vbCrLf & "Limit: " & CStr(RIBBON_STANDARD_MAX_ID_ATTRIBUTE_LENGTH)
    Else
        buttonID = RibbonButtonId(buttonTargetKey, objectName)
    End If
    RegisterRibbonCustomId usedCustomIds, buttonID, "Button", resolvedObjectTypeCode & " / " & objectName
    controlTag = RibbonControlTag(buttonTargetKey)
    RegisterRibbonStaticTag usedStaticTags, controlTag, "Button", resolvedObjectTypeCode & " / " & objectName
    screenTipCaption = RibbonGeneratedScreenTipCaption(resolvedObjectTypeCode, objectName)
    xmlText = RIBBON_XML_INDENT_4 & RIBBON_XML_INDENT_1 & "<button id=""" & RibbonXmlAttributeText(buttonID)
    xmlText = xmlText & """ label=""" & RibbonXmlAttributeText(labelText)
    If Len(imageMsoText) > 0 Then xmlText = xmlText & """ imageMso=""" & RibbonXmlAttributeText(imageMsoText)
    xmlText = xmlText & """ size=""normal"" onAction=""" & RIBBON_CALLBACK_OPEN_OBJECT
    ValidateRibbonGeneratedTitleBeforeManifestPersist tabKey, resolvedObjectTypeCode, objectName, labelText
    UpsertRibbonButtonTarget RIBBON_NAME_, buttonTargetKey, tabKey, tabLabel, groupLabel, labelText, screenTipCaption, imageMsoText, resolvedObjectTypeCode, objectName
    xmlText = xmlText & """ tag=""" & RibbonXmlAttributeText(controlTag)
    xmlText = xmlText & """ getScreentip=""" & RIBBON_CALLBACK_GET_SCREENTIP
    xmlText = xmlText & """ getSupertip=""" & RIBBON_CALLBACK_GET_SUPERTIP & """"
    RibbonButtonXml = xmlText & "/>"
End Function

Private Function TryRibbonTopLevelFastSkip(ByRef dependencyToken As String, ByRef standardInputFingerprint As String, ByRef standardButtonCount As Long, ByRef standardXml As String, ByRef detailText As String, ByRef resolvedGroupFingerprint As String, Optional ByVal prevalidatedPredecessorToken As String = vbNullString, Optional ByVal prevalidatedCatalogToken As String = vbNullString, Optional ByVal captureStagedEvidence As Boolean = False) As Boolean
    ' Fail-closed Part G fast path for the single ADC Ribbon implementation. Prerequisite proof errors propagate; only comparison against persisted Ribbon state is fail-soft.
    Dim db As DAO.Database
    Dim propertyExists As Boolean
    Dim currentDependencyDetail As String
    Dim storedDependencyToken As String
    Dim storedStandardInputFingerprint As String
    Dim storedStandardXmlHash As String
    Dim storedStandardTargetMapHash As String
    Dim storedStandardCountText As String
    Dim standardXmlExists As Boolean
    Dim actualStandardTargetMapHash As String
    Dim actualStandardCount As Long
    Dim currentStandardCount As Long
    Set db = GetHostDb()
    dependencyToken = CurrentRibbonDependencyToken(currentDependencyDetail, resolvedGroupFingerprint, prevalidatedPredecessorToken, prevalidatedCatalogToken, captureStagedEvidence)
    If Len(standardInputFingerprint) = 0 Or standardButtonCount <= 0 Then
        standardInputFingerprint = RibbonRowsFingerprint(currentStandardCount)
        standardButtonCount = currentStandardCount
    Else
        currentStandardCount = standardButtonCount
    End If
    On Error GoTo NotCurrent
    If SharedDao_DatabaseTableExistsSafe(db, DB_TABLE_USYS_RIBBONS) = False Then detailText = "USysRibbons is missing; ADC Ribbon build required.": Exit Function
    If SharedDao_DatabaseTableExistsSafe(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_RIBBON_BUTTON_TARGETS_TABLE_LOGICAL_ID)) = False Then detailText = "UI_Ribbon_Generated_BtnTargets is missing; Part G cannot prove a current Ribbon state.": Exit Function
    storedDependencyToken = DatabaseTextPropertyValue(db, RIBBON_DEPENDENCY_TOKEN_PROPERTY, propertyExists)
    If propertyExists = False Then detailText = "The top-level ribbon dependency token has not been seeded. " & currentDependencyDetail: Exit Function
    If StrComp(storedDependencyToken, dependencyToken, vbBinaryCompare) <> 0 Then detailText = "The top-level ribbon dependency token changed. " & currentDependencyDetail: Exit Function
    storedStandardInputFingerprint = DatabaseTextPropertyValue(db, RIBBON_STANDARD_INPUT_FINGERPRINT_PROPERTY, propertyExists)
    If propertyExists = False Or Len(storedStandardInputFingerprint) = 0 Then detailText = "Stored ADC-Ribbon input fingerprint is missing.": Exit Function
    If StrComp(storedStandardInputFingerprint, standardInputFingerprint, vbBinaryCompare) <> 0 Then detailText = "The resolved ADC-Ribbon source-row fingerprint changed.": Exit Function
    storedStandardXmlHash = DatabaseTextPropertyValue(db, RIBBON_STANDARD_XML_HASH_PROPERTY, propertyExists)
    If propertyExists = False Then detailText = "Stored ADC-Ribbon XML hash is missing.": Exit Function
    storedStandardTargetMapHash = DatabaseTextPropertyValue(db, RIBBON_STANDARD_TARGET_MAP_HASH_PROPERTY, propertyExists)
    If propertyExists = False Then detailText = "Stored ADC-Ribbon runtime-manifest hash is missing.": Exit Function
    storedStandardCountText = DatabaseTextPropertyValue(db, RIBBON_STANDARD_BUTTON_COUNT_PROPERTY, propertyExists)
    If propertyExists = False Or IsNumeric(storedStandardCountText) = False Then detailText = "Stored ADC-Ribbon button count is missing or invalid.": Exit Function
    If CLng(storedStandardCountText) <> currentStandardCount Then detailText = "Stored ADC-Ribbon button count changed.": Exit Function
    standardXml = StoredRibbonXml(db, RIBBON_NAME_, standardXmlExists)
    If standardXmlExists = False Or Len(Trim$(standardXml)) = 0 Then detailText = "Stored ADC-Ribbon XML is missing or blank.": Exit Function
    If StrComp(AdlerTextHash(standardXml) & ":" & CStr(Len(standardXml)), storedStandardXmlHash, vbBinaryCompare) <> 0 Then detailText = "Stored ADC-Ribbon XML changed outside the validated state.": Exit Function
    actualStandardTargetMapHash = RibbonTargetMapFingerprintForRibbon(db, RIBBON_NAME_, actualStandardCount)
    If StrComp(actualStandardTargetMapHash, storedStandardTargetMapHash, vbBinaryCompare) <> 0 Then detailText = "Stored ADC-Ribbon runtime manifest changed outside the validated state.": Exit Function
    If actualStandardCount <> currentStandardCount Then detailText = "Stored ADC-Ribbon runtime-manifest count changed.": Exit Function
    detailText = "Part G exact-state fast path succeeded: dependency token, resolved source rows, standard XML, and runtime manifest match the last validated build. " & currentDependencyDetail
    TryRibbonTopLevelFastSkip = True
    Exit Function
NotCurrent:
    If IsCancellationError(Err.Number) Then Err.Raise Err.Number, Err.Source, Err.Description
    detailText = "Top-level ribbon fast-path comparison could not prove an unchanged validated Ribbon state. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    Err.Clear
End Function

' Purpose: Builds and verifies the standard ADC Ribbon, applies it as the database CustomRibbonID, records the required reopen marker, and optionally reports the result to the user.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub AddTabsToAccessRibbon(Optional ByVal showMessage As Boolean = True)
    'Builds/verifies and applies only the standard Access Development Console ribbon.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim messageText As String
    Dim traceRunOwnedByProcedure As Boolean
    Dim restartWarningText As String
    Dim restartMarkerSaved As Boolean
    Dim appliedRibbonId As String
    Dim customRibbonPropertyExists As Boolean
    BeginTrackedProcess "Apply " & DEFAULT_UTILITY_LONG_NAME & " ADC Ribbon"
    traceRunOwnedByProcedure = BeginUITraceRunIfNeeded("UI.RibbonApply")
    PrintRibbonStatus "Standard ADC Ribbon apply started."
    If CodeDbIsHostDb() = False Then
        messageText = ChatGPTSameDbOnlyMessage("Applying the " & DEFAULT_UTILITY_LONG_NAME & " ADC Ribbon")
        If showMessage Then
            CoreShowMessage _
                messageText, _
                vbInformation, _
                DEFAULT_UI_MESSAGE_TITLE, _
                "ShowADCMessage", _
                "UI"
        End If
        CompleteUITraceRunIfOwned traceRunOwnedByProcedure, "ADC Ribbon apply skipped."
        EndTrackedProcess
        Exit Sub
    End If
    BuildRibbon False
    Set db = GetHostDb()
    RequireRibbonRowForApply db, RIBBON_NAME_
    Call EnsureDaoPropertyValue(db, "CustomRibbonID", dbText, RIBBON_NAME_, "AddTabsToAccessRibbon")
    appliedRibbonId = DatabaseTextPropertyValue(db, "CustomRibbonID", customRibbonPropertyExists)
    If customRibbonPropertyExists = False Or StrComp(appliedRibbonId, RIBBON_NAME_, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 7220, "AddTabsToAccessRibbon", "CustomRibbonID verification failed after applying the standard " & DEFAULT_UTILITY_LONG_NAME & " ribbon."
    restartMarkerSaved = TryMarkRibbonRestartRequired(restartWarningText)
    If restartMarkerSaved = False Then Err.Raise vbObjectError + 7222, "AddTabsToAccessRibbon", "The " & DEFAULT_UTILITY_LONG_NAME & " ADC ribbon was applied and verified, but the required database-reopen marker could not be persisted." & vbCrLf & restartWarningText
    messageText = DEFAULT_UTILITY_LONG_NAME & " ADC ribbon applied. Close and reopen the database to load it."
    LogTraceMessage TRACE_LEVEL_INFO, "Ribbon", "AddTabsToAccessRibbon", "Standard ADC Ribbon applied.", messageText, "Apply ADC Ribbon", 0, vbNullString, vbNullString, RIBBON_NAME_, "Ribbon", "CodeDb", "Ribbon"
    If showMessage Then
        CoreShowMessage _
            messageText, _
            vbInformation, _
            DEFAULT_UI_MESSAGE_TITLE, _
            "ShowADCMessage", _
            "UI"
    End If
    CompleteUITraceRunIfOwned traceRunOwnedByProcedure, "Standard ADC Ribbon apply complete."
    EndTrackedProcess
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    messageText = "The " & DEFAULT_UTILITY_LONG_NAME & " standard ADC Ribbon apply operation did not complete cleanly." & vbCrLf & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description:" & vbCrLf & savedErrDescription & TableDiagnosticsLocationForMessage()
    LogTraceMessage TRACE_LEVEL_ERROR, "Ribbon", "AddTabsToAccessRibbon", "Standard ADC Ribbon apply operation did not complete cleanly.", messageText, vbNullString, savedErrNumber, savedErrSource, savedErrDescription, RIBBON_NAME_, "Ribbon", "CodeDb", "Ribbon"
    CompleteUITraceRunIfOwned traceRunOwnedByProcedure, "ADC Ribbon apply failed.", False
    EndTrackedProcess
    If showMessage Then
        CoreShowErrorMessage _
            messageText, _
            vbCritical, _
            DEFAULT_UI_MESSAGE_TITLE, _
            "AddTabsToAccessRibbon", _
            "UI"
        Exit Sub
    End If
    Err.Raise savedErrNumber, "AddTabsToAccessRibbon", messageText
End Sub

Private Function RibbonRowsFingerprint(ByRef rowCount As Long) As String
    On Error GoTo ErrHandler
    Dim rows As Collection
    Dim rowData As Variant
    Dim targetMetadata As Variant
    Dim itemIndex As Long
    Dim payloadText As String
    Dim sourceDetailText As String
    Dim targetTableNames As String
    Dim targetQueryNames As String
    Set rows = RibbonPhaseRows(sourceDetailText)
    If mRibbonSystemMacroTargetCache Is Nothing Then PrepareRibbonCatalogSourceSnapshot
    payloadText = RIBBON_FINGERPRINT_VERSION & "|Standard|"
    For Each rowData In rows
        For itemIndex = LBound(rowData) To UBound(rowData)
            payloadText = payloadText & FingerprintTokenPart(CStr(Nz(rowData(itemIndex), vbNullString)))
        Next itemIndex
        targetTableNames = vbNullString
        targetQueryNames = vbNullString
        If StrComp(CStr(rowData(3)), ADCAT_OBJECT_TYPE_MACRO, vbTextCompare) = 0 Then
            If mRibbonSystemMacroTargetCache.Exists(CStr(rowData(4))) Then
                targetMetadata = mRibbonSystemMacroTargetCache(CStr(rowData(4)))
                targetTableNames = CStr(targetMetadata(0))
                targetQueryNames = CStr(targetMetadata(1))
            End If
        End If
        payloadText = payloadText & FingerprintTokenPart(targetTableNames) & FingerprintTokenPart(targetQueryNames) & vbLf
    Next rowData
    rowCount = rows.count
    RibbonRowsFingerprint = AdlerTextHash(payloadText) & ":" & CStr(Len(payloadText))
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "RibbonRowsFingerprint", "Could not fingerprint the resolved ADC-Ribbon input rows." & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Function

Private Sub PersistStagedRibbonFastPathState()
    'Strictly verifies the committed ADC ribbon, then persists only ADC-Ribbon fast-path evidence.
    On Error GoTo StrictVerificationError
    Dim db As DAO.Database
    Dim standardTargetMapHash As String
    Dim actualStandardCount As Long
    Set db = GetHostDb()
    standardTargetMapHash = RibbonTargetMapFingerprintForRibbon(db, RIBBON_NAME_, actualStandardCount)
    If actualStandardCount <> mRibbonPhaseButtonCount Then Err.Raise vbObjectError + 7198, "PersistStagedRibbonFastPathState", "The ADC-Ribbon runtime-manifest count did not match the staged button count."
    If mRibbonPhaseExpectedStandardCount <> mRibbonPhaseButtonCount Then Err.Raise vbObjectError + 7200, "PersistStagedRibbonFastPathState", "The ADC-Ribbon source-row count changed during the staged build."
    On Error GoTo OptionalStateWarning
    Call EnsureDaoPropertyValue(db, RIBBON_STANDARD_INPUT_FINGERPRINT_PROPERTY, dbText, mRibbonPhaseStandardInputFingerprint, "PersistStagedRibbonFastPathState")
    Call EnsureDaoPropertyValue(db, RIBBON_STANDARD_XML_HASH_PROPERTY, dbText, AdlerTextHash(mRibbonPhaseRibbonXml) & ":" & CStr(Len(mRibbonPhaseRibbonXml)), "PersistStagedRibbonFastPathState")
    Call EnsureDaoPropertyValue(db, RIBBON_STANDARD_TARGET_MAP_HASH_PROPERTY, dbText, standardTargetMapHash, "PersistStagedRibbonFastPathState")
    Call EnsureDaoPropertyValue(db, RIBBON_STANDARD_BUTTON_COUNT_PROPERTY, dbText, CStr(mRibbonPhaseButtonCount), "PersistStagedRibbonFastPathState")
    If Len(mRibbonPhaseDependencyToken) > 0 Then Call EnsureDaoPropertyValue(db, RIBBON_DEPENDENCY_TOKEN_PROPERTY, dbText, mRibbonPhaseDependencyToken, "PersistStagedRibbonFastPathState")
    Exit Sub
StrictVerificationError:
    Dim strictErrNumber As Long
    Dim strictErrSource As String
    Dim strictErrDescription As String
    strictErrNumber = Err.Number
    strictErrSource = Err.Source
    strictErrDescription = Err.Description
    PersistRibbonPostCommitVerificationFailure strictErrNumber, strictErrSource, strictErrDescription
    Err.Raise strictErrNumber, "PersistStagedRibbonFastPathState", "Committed ADC-Ribbon verification failed after the ribbon transaction committed." & vbCrLf & strictErrDescription
OptionalStateWarning:
    LogTraceMessage TRACE_LEVEL_WARNING, "Ribbon", "PersistStagedRibbonFastPathState", "Committed ADC ribbon verified, but optional fast-path evidence could not be persisted.", "Err.Number: " & CStr(Err.Number) & vbCrLf & "Err.Source: " & Err.Source & vbCrLf & "Err.Description: " & Err.Description, "Persist optional ribbon fast-path state", Err.Number, Err.Source, Err.Description, RIBBON_NAME_, "Ribbon", "CodeDb", "Ribbon"
    Err.Clear
End Sub

Private Function ClearRibbonPostCommitVerificationFailure(ByRef failureDetail As String) As Boolean
    On Error GoTo ErrHandler
    Dim operationName As String
    Dim detailText As String
    Dim payloadText As String
    DeleteDurableRecoveryMarker UI_RIBBON_POST_COMMIT_VERIFY_KEY
    If CoreDurableRecoveryMarkerExists(UI_RIBBON_POST_COMMIT_VERIFY_KEY, operationName, detailText, payloadText) Then
        failureDetail = "The ADC-Ribbon post-commit verification marker still exists after deletion."
        Exit Function
    End If
    ClearRibbonPostCommitVerificationFailure = True
    Exit Function
ErrHandler:
    failureDetail = "Could not clear the ADC-Ribbon post-commit verification marker. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
End Function

' Purpose: Builds or verifies only the standard Access Development Console ribbon.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
'   prevalidatedCatalogToken (Input Optional; String): Proof/version token identifying the exact prevalidated catalog state that was previously validated in this setup run. Default: vbNullString.
'   prevalidatedCatalogTraceRunID (Input Optional; String): Trace-run ID of catalog state already validated by the predecessor stage and supplied to avoid redundant validation. Default: vbNullString.
'   prevalidatedHelpG1Token (Input Optional; String): Proof/version token identifying the exact prevalidated help G1 state that was previously validated in this setup run. Default: vbNullString.
'   prevalidatedHelpG1TraceRunID (Input Optional; String): Trace-run ID of the same-run G1 Help/description proof already validated by the predecessor stage. Default: vbNullString.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub BuildRibbon(Optional ByVal showMessage As Boolean = True, Optional ByVal prevalidatedCatalogToken As String = vbNullString, Optional ByVal prevalidatedCatalogTraceRunID As String = vbNullString, Optional ByVal prevalidatedHelpG1Token As String = vbNullString, Optional ByVal prevalidatedHelpG1TraceRunID As String = vbNullString)
    'Builds or verifies only the standard Access Development Console ribbon.
    On Error GoTo ErrHandler
    PrepareCustomRibbonBuild showMessage, False, prevalidatedCatalogToken, prevalidatedCatalogTraceRunID, prevalidatedHelpG1Token, prevalidatedHelpG1TraceRunID
    BuildStandardRibbonPhase
    CommitCustomRibbonBuild
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    If showMessage Then
        CoreShowErrorMessage _
            savedErrDescription, _
            vbCritical, _
            DEFAULT_UI_MESSAGE_TITLE, _
            "BuildRibbon", _
            "UI"
        Exit Sub
    End If
    Err.Raise savedErrNumber, "BuildRibbon", savedErrDescription
End Sub

' Purpose: Prepares custom ribbon build for coordinated setup, prevalidated catalog token.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: False.
'   coordinatedSetup (Input Optional; Boolean): True indicates the call is running inside coordinated numbered Setup and must obey its phase ownership/rollback rules; False indicates standalone execution.
'   Default: False.
'   prevalidatedCatalogToken (Input Optional; String): Proof/version token identifying the exact prevalidated catalog state that was previously validated in this setup run. Default: vbNullString.
'   prevalidatedCatalogTraceRunID (Input Optional; String): Trace-run ID of catalog state already validated by the predecessor stage and supplied to avoid redundant validation. Default: vbNullString.
'   prevalidatedHelpG1Token (Input Optional; String): Proof/version token identifying the exact prevalidated help G1 state that was previously validated in this setup run. Default: vbNullString.
'   prevalidatedHelpG1TraceRunID (Input Optional; String): Trace-run ID of the same-run G1 Help/description proof already validated by the predecessor stage. Default: vbNullString.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub PrepareCustomRibbonBuild(Optional ByVal showMessage As Boolean = False, Optional ByVal coordinatedSetup As Boolean = False, Optional ByVal prevalidatedCatalogToken As String = vbNullString, Optional ByVal prevalidatedCatalogTraceRunID As String = vbNullString, Optional ByVal prevalidatedHelpG1Token As String = vbNullString, Optional ByVal prevalidatedHelpG1TraceRunID As String = vbNullString)
    ' Part H1 owner: read-only prerequisite verification and exact standard-Ribbon build plan. Physical Navigation Pane state is not read or mutated; TraceLog table writes are suppressed.
    On Error GoTo ErrHandler
    Dim readinessStartedTimer As Single
    Dim topLevelDetail As String
    Dim catalogDetail As String
    Dim catalogToken As String
    Dim recoveryDetail As String
    Dim traceTableWritesSuppressed As Boolean
    Dim sameRunCatalogHandoffAccepted As Boolean
    Dim authoritativeSameRunCatalogHandoffToken As String
    Dim authoritativeSameRunCatalogHandoffTraceRunID As String
    Dim authoritativeSameRunCatalogHandoffDetail As String
    Dim sameRunOperationalPredecessorProofAccepted As Boolean
    Dim authoritativeOperationalPredecessorProofToken As String
    Dim authoritativeOperationalPredecessorProofTraceRunID As String
    Dim authoritativeOperationalPredecessorProofDetail As String
    Dim authoritativeOperationalCatalogToken As String
    If mRibbonPhaseTransactionStarted Then Err.Raise vbObjectError + 7191, "PrepareCustomRibbonBuild", "A prior staged ribbon transaction is still active."
    If UIRecoveryBlockedForSetupPreflight(recoveryDetail) Then Err.Raise ERR_UI_RECOVERY_BLOCKED, "PrepareCustomRibbonBuild", recoveryDetail
    BeginTraceTableWriteSuppression
    traceTableWritesSuppressed = True
    InvalidateRibbonSameRunH3Proof
    ResetRibbonPhaseState
    mRibbonPhaseApplyAfterBuild = SettingBool(SETTING_APPLY_RIBBON_AFTER_BUILD, DEFAULT_APPLY_RIBBON_AFTER_BUILD)
    mRibbonPhaseApplyAfterBuildCaptured = True
    Set mRibbonCaptionSourceCache = Nothing
    mRibbonPhaseToken = NewRibbonPhaseToken()
    mRibbonPhaseShowMessage = showMessage
    mRibbonPhaseStartedTimer = Timer
    readinessStartedTimer = Timer
    BeginRibbonStatusParts 3
    mRibbonPhaseTraceRunOwned = BeginUITraceRunIfNeeded("UI.RibbonBuild")
    PrintRibbonStageStatus 1, "H1 read-only Ribbon prerequisite and exact-state plan."
    If CodeDbIsHostDb() = False Then
        mRibbonPhaseSkippedSameDatabase = True
        mRibbonPhaseStandardMode = RIBBON_PHASE_MODE_SKIP
        mRibbonPhaseStandardCompleted = True
        mRibbonPhasePrepared = True
        mRibbonPhaseReadinessDetail = "H1 skipped same-database Ribbon work because CodeDb is not HostDb."
        mRibbonPhaseReadinessElapsedSeconds = ElapsedSecondsFromTimer(readinessStartedTimer)
        EndTraceTableWriteSuppression
        traceTableWritesSuppressed = False
        Exit Sub
    End If
    If coordinatedSetup And (Len(Trim$(prevalidatedHelpG1Token)) > 0 Or Len(Trim$(prevalidatedHelpG1TraceRunID)) > 0) Then Err.Raise vbObjectError + 7320, "PrepareCustomRibbonBuild", "An operational prevalidated predecessor proof cannot be supplied to the coordinated Setup H1 path."
    If coordinatedSetup = False And Len(Trim$(prevalidatedHelpG1Token)) > 0 Then
        If Len(Trim$(prevalidatedHelpG1TraceRunID)) = 0 Then Err.Raise vbObjectError + 7321, "PrepareCustomRibbonBuild", "A prevalidated operational predecessor token was supplied without its trace identity."
        If HelpDocumentMacroBuildRepairSameRunFullRibbonPredecessorProof(authoritativeOperationalPredecessorProofToken, authoritativeOperationalPredecessorProofTraceRunID, authoritativeOperationalPredecessorProofDetail, authoritativeOperationalCatalogToken) = False Then Err.Raise vbObjectError + 7322, "PrepareCustomRibbonBuild", "Review did not confirm a complete Build/Repair Help Document Macros predecessor proof from the current trace run." & vbCrLf & authoritativeOperationalPredecessorProofDetail
        If StrComp(Trim$(prevalidatedHelpG1TraceRunID), authoritativeOperationalPredecessorProofTraceRunID, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7323, "PrepareCustomRibbonBuild", "The supplied operational predecessor proof trace identity does not match Review's authoritative same-run proof."
        If StrComp(Trim$(prevalidatedHelpG1Token), authoritativeOperationalPredecessorProofToken, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7324, "PrepareCustomRibbonBuild", "The supplied operational predecessor token does not match Review's authoritative same-run complete proof."
        If Len(authoritativeOperationalCatalogToken) = 0 Then Err.Raise vbObjectError + 7350, "PrepareCustomRibbonBuild", "The accepted operational predecessor proof did not expose its authoritative complete-catalog component."
        sameRunOperationalPredecessorProofAccepted = True
    ElseIf Len(Trim$(prevalidatedHelpG1TraceRunID)) > 0 Then
        Err.Raise vbObjectError + 7325, "PrepareCustomRibbonBuild", "An operational predecessor trace identity was supplied without the corresponding proof token."
    End If
    If coordinatedSetup Then
        ' C3, D4, and G1 already prove physical macros, the complete logical catalog, and managed descriptions.
        ' The coordinated Setup path consumes those exact predecessor proofs and performs only Ribbon-specific source resolution, coverage, and delta planning; physical Navigation Pane placement remains deferred to
        ' H.
    Else
        ' H1 owns this direct Ribbon dependency schema check even when Build/Repair Help Document Macros supplies a complete same-run catalog/G1 predecessor proof. The proof eliminates duplicate
        ' catalog/placement/inventory work; it does not replace Ribbon-owned UI_Groups field requirements.
        ValidateRibbonDependencies
    End If
    If SharedDao_DatabaseTableExistsSafe(GetHostDb(), ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_RIBBON_BUTTON_TARGETS_TABLE_LOGICAL_ID)) = False Then Err.Raise vbObjectError + 7188, "PrepareCustomRibbonBuild", "UI_Ribbon_Generated_BtnTargets is missing. Part G does not repair earlier setup tables."
    If sameRunOperationalPredecessorProofAccepted Then
        catalogToken = authoritativeOperationalCatalogToken
        catalogDetail = "Complete catalog token reused from Review's accepted same-run Build/Repair Help Document Macros predecessor proof."
    Else
        catalogToken = CompleteCatalogStateTokenForSetup(catalogDetail)
        If Len(catalogToken) = 0 Then Err.Raise vbObjectError + 7243, "PrepareCustomRibbonBuild", catalogDetail
    End If
    If coordinatedSetup = False And Len(Trim$(prevalidatedCatalogToken)) > 0 Then
        If Len(Trim$(prevalidatedCatalogTraceRunID)) = 0 Then Err.Raise vbObjectError + 7273, "PrepareCustomRibbonBuild", "A same-run complete-catalog handoff token was supplied without its trace identity."
        If CompleteCatalogSameRunProofForRibbon(authoritativeSameRunCatalogHandoffToken, authoritativeSameRunCatalogHandoffTraceRunID, authoritativeSameRunCatalogHandoffDetail, catalogToken, Trim$(prevalidatedCatalogTraceRunID)) = False Then Err.Raise vbObjectError + 7272, "PrepareCustomRibbonBuild", "Catalog did not confirm the exact current complete-catalog handoff for this trace run." & vbCrLf & authoritativeSameRunCatalogHandoffDetail
        If StrComp(Trim$(prevalidatedCatalogTraceRunID), authoritativeSameRunCatalogHandoffTraceRunID, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7273, "PrepareCustomRibbonBuild", "The supplied complete-catalog handoff trace identity does not match Catalog's authoritative same-run handoff."
        If StrComp(Trim$(prevalidatedCatalogToken), authoritativeSameRunCatalogHandoffToken, vbBinaryCompare) <> 0 Or StrComp(authoritativeSameRunCatalogHandoffToken, catalogToken, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7274, "PrepareCustomRibbonBuild", "The supplied same-run complete-catalog handoff does not match Catalog's authoritative current-run state that H1 is about to consume."
        sameRunCatalogHandoffAccepted = True
    ElseIf Len(Trim$(prevalidatedCatalogTraceRunID)) > 0 Then
        Err.Raise vbObjectError + 7275, "PrepareCustomRibbonBuild", "A same-run complete-catalog handoff trace identity was supplied without the corresponding complete-catalog handoff token."
    End If
    If coordinatedSetup = False And sameRunCatalogHandoffAccepted = False And sameRunOperationalPredecessorProofAccepted = False Then
        If ValidateManagedObjectPlacementMetadata(False, False, True) = False Then Err.Raise vbObjectError + 8763, "PrepareCustomRibbonBuild", ManagedObjectPlacementDiagnosticText(50)
    End If
    PrepareRibbonCatalogSourceSnapshot
    If coordinatedSetup = False And sameRunCatalogHandoffAccepted = False And sameRunOperationalPredecessorProofAccepted = False Then ValidateRibbonManagedObjectInventoryPreflight
    LogTraceMessage TRACE_LEVEL_INFO, "RibbonTiming", "PrepareCustomRibbonBuild", "H1 prerequisite mode selected.", _
        "Mode=" & IIf(coordinatedSetup, "CoordinatedSetup", IIf(sameRunCatalogHandoffAccepted And sameRunOperationalPredecessorProofAccepted, "OperationalSameRunCatalogHandoffAndFullPredecessorProof", IIf(sameRunCatalogHandoffAccepted, "OperationalSameRunCatalogHandoff", IIf(sameRunOperationalPredecessorProofAccepted, "OperationalFullPredecessorProof", "Standalone")))) & vbCrLf & _
        "Full placement validation executed=" & CStr(coordinatedSetup = False And sameRunCatalogHandoffAccepted = False And sameRunOperationalPredecessorProofAccepted = False) & vbCrLf & _
        "Managed inventory preflight executed=" & CStr(coordinatedSetup = False And sameRunCatalogHandoffAccepted = False And sameRunOperationalPredecessorProofAccepted = False) & vbCrLf & _
        "Complete operational predecessor proof accepted=" & CStr(sameRunOperationalPredecessorProofAccepted), _
        "H1 prerequisite mode", 0, vbNullString, vbNullString, RIBBON_NAME_, "Ribbon", "CodeDb", "H1"
    PrepareRibbonPhaseRows
    mRibbonPhaseStandardInputFingerprint = RibbonRowsFingerprint(mRibbonPhaseExpectedStandardCount)
    mRibbonPhaseButtonCount = mRibbonPhaseExpectedStandardCount
    If TryRibbonTopLevelFastSkip(mRibbonPhaseDependencyToken, mRibbonPhaseStandardInputFingerprint, mRibbonPhaseExpectedStandardCount, mRibbonPhaseRibbonXml, topLevelDetail, mRibbonPhaseGroupFingerprint, IIf(sameRunOperationalPredecessorProofAccepted, authoritativeOperationalPredecessorProofToken, vbNullString), catalogToken, True) Then
        mRibbonPhaseFastPathUsed = True
        mRibbonPhaseBuildRequired = False
        mRibbonPhaseStandardRequired = False
        mRibbonPhaseStandardMode = RIBBON_PHASE_MODE_VERIFY_ONLY
        mRibbonPhaseStandardCompleted = True
        mRibbonPhaseButtonCount = mRibbonPhaseExpectedStandardCount
        mRibbonPhaseReadinessDetail = topLevelDetail
    Else
        If Len(mRibbonPhaseDependencyToken) = 0 Then mRibbonPhaseDependencyToken = CurrentRibbonDependencyToken(catalogDetail, mRibbonPhaseGroupFingerprint, IIf(sameRunOperationalPredecessorProofAccepted, authoritativeOperationalPredecessorProofToken, vbNullString), catalogToken, True)
        mRibbonPhaseStandardRequired = True
        mRibbonPhaseStandardMode = RIBBON_PHASE_MODE_BUILD
        mRibbonPhaseBuildRequired = True
        mRibbonPhaseReadinessDetail = "H1 determined that the ADC Ribbon requires a staged rebuild." & vbCrLf & topLevelDetail
    End If
    mRibbonPhasePrepared = True
    mRibbonPhaseReadinessElapsedSeconds = ElapsedSecondsFromTimer(readinessStartedTimer)
    LogTraceMessage TRACE_LEVEL_INFO, "RibbonTiming", "PrepareCustomRibbonBuild", "H1 read-only Ribbon plan completed.", mRibbonPhaseReadinessDetail & vbCrLf & "Expected buttons: " & CStr(mRibbonPhaseExpectedStandardCount) & vbCrLf & "Apply ADC Ribbon after build: " & CStr(mRibbonPhaseApplyAfterBuild) & vbCrLf & "Staged dependency evidence retained: Catalog=" & CStr(Len(mRibbonPhaseCatalogToken) > 0) & "; D4/E3/G1 predecessor=" & CStr(Len(mRibbonPhasePredecessorToken) > 0) & "; RibbonGroups=" & CStr(Len(mRibbonPhaseGroupFingerprint) > 0), "H1 Ribbon plan", 0, vbNullString, vbNullString, RIBBON_NAME_, "Ribbon", "CodeDb", "H1"
    EndTraceTableWriteSuppression
    traceTableWritesSuppressed = False
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If traceTableWritesSuppressed Then EndTraceTableWriteSuppression
    EndRibbonStatusParts
    ResetRibbonPhaseState
    On Error GoTo 0
    If showMessage And IsCancellationError(savedErrNumber) = False Then
        CoreShowErrorMessage _
            savedErrDescription, _
            vbCritical, _
            DEFAULT_UI_MESSAGE_TITLE, _
            "PrepareCustomRibbonBuild", _
            "UI"
        Exit Sub
    End If
    Err.Raise savedErrNumber, IIf(Len(Trim$(savedErrSource)) = 0, "PrepareCustomRibbonBuild", savedErrSource), savedErrDescription
End Sub

' Purpose: Performs commit custom ribbon build by coordinating `RequireRibbonPhaseToken` and `PrintRibbonStageStatus` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub CommitCustomRibbonBuild()
    RequireSetupDomainMutationAllowed SETUP_DOMAIN_H3_RIBBON_COMMIT, "CommitCustomRibbonBuild"
    ' Part H3 owner: commits only H2-validated state, enforces the H1-captured automatic-application policy for CustomRibbonID, then finalizes exact-state evidence.
    On Error GoTo ErrHandler
    Dim showMessage As Boolean
    Dim traceRunOwned As Boolean
    Dim phaseStartedTimer As Single
    Dim ribbonXmlChanged As Boolean
    Dim customRibbonChanged As Boolean
    Dim propertyExists As Boolean
    Dim currentRibbonId As String
    Dim verifyRibbonId As String
    Dim ribbonConfigured As Boolean
    Dim RestartRequired As Boolean
    Dim restartWarningText As String
    Dim recoveryClearDetail As String
    Dim postCommitClearDetail As String
    Dim successMessage As String
    Dim statusSuccessMessage As String
    Dim visibleStatusSuccessMessage As String
    RequireRibbonPhaseToken "CommitCustomRibbonBuild"
    If mRibbonPhaseApplyAfterBuildCaptured = False Then Err.Raise vbObjectError + 7267, "CommitCustomRibbonBuild", "H3 refused to continue because H1 did not capture ApplyADCRibbonAfterBuild for this Ribbon operation."
    showMessage = mRibbonPhaseShowMessage
    traceRunOwned = mRibbonPhaseTraceRunOwned
    phaseStartedTimer = Timer
    PrintRibbonStageStatus 3, "H3 commit validated ADC Ribbon and enforce captured application policy."
    If mRibbonPhaseStandardCompleted = False Then Err.Raise vbObjectError + 7202, "CommitCustomRibbonBuild", "H2 did not complete for the current H1 plan."
    If mRibbonPhaseSkippedSameDatabase Then
        CompleteUITraceRunIfOwned traceRunOwned, "Part G same-database Ribbon work skipped."
        EndRibbonStatusParts
        ResetRibbonPhaseState
        Exit Sub
    End If
    RequireRibbonPhaseSourcePlanCurrent "CommitCustomRibbonBuild", (mRibbonPhaseStandardRequired = False)
    If mRibbonPhaseStandardRequired Then
        If mRibbonPhaseStandardBuilt = False Then Err.Raise vbObjectError + 7194, "CommitCustomRibbonBuild", "H1 required a ADC Ribbon rebuild, but H2 did not complete it."
        If mRibbonPhaseTransactionStarted = False Then Err.Raise vbObjectError + 7196, "CommitCustomRibbonBuild", "H2 staged Ribbon work has no transaction to commit."
        ribbonXmlChanged = mRibbonPhaseStandardChanged
        PersistRibbonCommitMarker UI_RIBBON_COMMIT_STATUS_COMMIT_PENDING, ribbonXmlChanged, mRibbonPhaseApplyAfterBuild
        mRibbonPhaseWorkspace.CommitTrans
        mRibbonPhaseTransactionStarted = False
        mRibbonPhaseMutationCommitted = True
        PersistRibbonCommitMarker UI_RIBBON_COMMIT_STATUS_COMMITTED, ribbonXmlChanged, mRibbonPhaseApplyAfterBuild
        PersistStagedRibbonFastPathState
    Else
        PersistRibbonDependencyToken
    End If
    currentRibbonId = DatabaseTextPropertyValue(GetHostDb(), "CustomRibbonID", propertyExists)
    ribbonConfigured = (propertyExists And StrComp(currentRibbonId, RIBBON_NAME_, vbTextCompare) = 0)
    If mRibbonPhaseApplyAfterBuild Then
        customRibbonChanged = Not ribbonConfigured
        If customRibbonChanged Then
            EnsureDaoPropertyValue GetHostDb(), "CustomRibbonID", dbText, RIBBON_NAME_, "CommitCustomRibbonBuild"
            mRibbonPhaseMutationCommitted = True
        End If
        verifyRibbonId = DatabaseTextPropertyValue(GetHostDb(), "CustomRibbonID", propertyExists)
        If propertyExists = False Or StrComp(verifyRibbonId, RIBBON_NAME_, vbTextCompare) <> 0 Then
            PersistRibbonPostCommitVerificationFailure vbObjectError + 7220, "CommitCustomRibbonBuild", "CustomRibbonID verification failed after H3 applied the standard " & DEFAULT_UTILITY_SHORT_NAME & " Ribbon with ApplyADCRibbonAfterBuild=True."
            Err.Raise vbObjectError + 7220, "CommitCustomRibbonBuild", "CustomRibbonID verification failed after H3 applied the standard " & DEFAULT_UTILITY_SHORT_NAME & " Ribbon."
        End If
        ribbonConfigured = True
    End If
    If mRibbonPhaseStandardRequired Then
        If Not mRibbonUI Is Nothing Then mRibbonUI.Invalidate
    End If
    RestartRequired = customRibbonChanged Or (ribbonXmlChanged And ribbonConfigured)
    If RestartRequired Then
        mRibbonPhaseRestartMarkerSaved = TryMarkRibbonRestartRequired(restartWarningText)
        If mRibbonPhaseRestartMarkerSaved = False Then
            PersistRibbonPostCommitVerificationFailure vbObjectError + 7211, "CommitCustomRibbonBuild", "H3 committed the ADC Ribbon but the required database-reopen marker could not be persisted." & vbCrLf & restartWarningText
            Err.Raise vbObjectError + 7211, "CommitCustomRibbonBuild", "H3 committed the ADC Ribbon but the required database-reopen marker could not be persisted." & vbCrLf & restartWarningText
        End If
    Else
        mRibbonPhaseRestartMarkerSaved = True
    End If
    If mRibbonPhaseApplyAfterBuild Then
        If customRibbonChanged Then
            successMessage = DEFAULT_UTILITY_LONG_NAME & " ADC Ribbon committed and automatically applied; database reopen required."
            statusSuccessMessage = DEFAULT_UTILITY_SHORT_NAME & " ADC Ribbon committed and automatically applied; database reopen required."
            visibleStatusSuccessMessage = "ADC Ribbon committed and automatically applied; database reopen required."
        ElseIf ribbonXmlChanged And ribbonConfigured Then
            successMessage = DEFAULT_UTILITY_LONG_NAME & " ADC Ribbon committed; ADC was already the configured database Ribbon, so a database reopen is required."
            statusSuccessMessage = DEFAULT_UTILITY_SHORT_NAME & " ADC Ribbon committed; ADC was already the configured database Ribbon, so a database reopen is required."
            visibleStatusSuccessMessage = "ADC Ribbon committed; ADC was already the configured database Ribbon, so a database reopen is required."
        Else
            successMessage = DEFAULT_UTILITY_LONG_NAME & " ADC Ribbon exact state and automatic-application policy verified current."
            statusSuccessMessage = DEFAULT_UTILITY_SHORT_NAME & " ADC Ribbon exact state and automatic-application policy verified current."
            visibleStatusSuccessMessage = "ADC Ribbon exact state and automatic-application policy verified current."
        End If
    Else
        If ribbonXmlChanged And ribbonConfigured Then
            successMessage = DEFAULT_UTILITY_LONG_NAME & " ADC Ribbon committed; automatic application skipped because ApplyADCRibbonAfterBuild=False, but ADC was already the configured database Ribbon so a database reopen is required."
            statusSuccessMessage = DEFAULT_UTILITY_SHORT_NAME & " ADC Ribbon committed; automatic application skipped because ApplyADCRibbonAfterBuild=False, but ADC was already the configured database Ribbon so a database reopen is required."
            visibleStatusSuccessMessage = "ADC Ribbon committed; automatic application skipped because ApplyADCRibbonAfterBuild=False, but ADC was already the configured database Ribbon so a database reopen is required."
        ElseIf mRibbonPhaseStandardRequired Then
            successMessage = DEFAULT_UTILITY_LONG_NAME & " ADC Ribbon committed; automatic application skipped because ApplyADCRibbonAfterBuild=False."
            statusSuccessMessage = DEFAULT_UTILITY_SHORT_NAME & " ADC Ribbon committed; automatic application skipped because ApplyADCRibbonAfterBuild=False."
            visibleStatusSuccessMessage = "ADC Ribbon committed; automatic application skipped because ApplyADCRibbonAfterBuild=False."
        Else
            successMessage = DEFAULT_UTILITY_LONG_NAME & " ADC Ribbon exact state verified current; automatic application skipped because ApplyADCRibbonAfterBuild=False."
            statusSuccessMessage = DEFAULT_UTILITY_SHORT_NAME & " ADC Ribbon exact state verified current; automatic application skipped because ApplyADCRibbonAfterBuild=False."
            visibleStatusSuccessMessage = "ADC Ribbon exact state verified current; automatic application skipped because ApplyADCRibbonAfterBuild=False."
        End If
    End If
    If ClearRibbonPostCommitVerificationFailure(postCommitClearDetail) = False Then Err.Raise vbObjectError + 7214, "CommitCustomRibbonBuild", postCommitClearDetail
    If mRibbonPhaseStandardRequired Then
        PersistRibbonCommitMarker UI_RIBBON_COMMIT_STATUS_POST_COMMIT_COMPLETE, ribbonXmlChanged, mRibbonPhaseApplyAfterBuild
        If ClearVerifiedRibbonTransactionRecoveryOnly(recoveryClearDetail) = False Then Err.Raise vbObjectError + 7215, "CommitCustomRibbonBuild", recoveryClearDetail
        DeleteRibbonCommitMarker
    End If
    PublishRibbonSameRunH3Proof
    mRibbonPhaseCommitElapsedSeconds = ElapsedSecondsFromTimer(phaseStartedTimer)
    ReportCoreActionSuccess "CommitCustomRibbonBuild", "Ribbon", statusSuccessMessage, "ApplyAfterBuild: " & CStr(mRibbonPhaseApplyAfterBuild) & vbCrLf & "Standard buttons: " & CStr(mRibbonPhaseButtonCount) & vbCrLf & StagedRibbonTimingText(), "H3 commit ADC Ribbon", RIBBON_NAME_, "Ribbon", "CodeDb", "UI", showMessage, DEFAULT_UI_MESSAGE_TITLE, vbInformation, False, "UI", True, visibleStatusMessage:=visibleStatusSuccessMessage
    CompleteUITraceRunIfOwned traceRunOwned, successMessage
    EndRibbonStatusParts
    ResetRibbonPhaseState
    Exit Sub
ErrHandler:
    RaiseRibbonPhaseFailure "CommitCustomRibbonBuild", Err.Number, Err.Source, Err.Description
End Sub

' Purpose: Initializes terminal teardown, coordinating `DisableRibbon` and `End` as the implementation requires.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub BeginTerminalTeardown()
    'The terminal state is set before startup Ribbon removal so the currently loaded ADC Ribbon cannot enter database-backed callbacks during teardown.
    On Error GoTo ErrHandler
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    mTerminalTeardownActive = True
    DisableRibbon False
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    mTerminalTeardownActive = False
    Err.Raise savedErrNumber, "BeginTerminalTeardown", "Could not establish terminal ADC teardown before destructive deletion." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

' Purpose: Disables ribbon, coordinating `RequireNoUnfinishedUIRecovery` and `CodeDbIsHostDb` as the implementation requires.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub DisableRibbon(Optional ByVal showMessage As Boolean = True)
    'Disables only the standard ADC startup ribbon; any unrelated CustomRibbonID is preserved.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim currentRibbonId As String
    Dim propertyExists As Boolean
    Dim restartWarningText As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim messageText As String
    RequireNoUnfinishedUIRecovery "DisableRibbon"
    If CodeDbIsHostDb() = False Then Err.Raise vbObjectError + 7221, "DisableRibbon", ChatGPTSameDbOnlyMessage("Disabling the " & DEFAULT_UTILITY_LONG_NAME & " ADC Ribbon")
    Set db = GetHostDb()
    currentRibbonId = DatabaseTextPropertyValue(db, "CustomRibbonID", propertyExists)
    If propertyExists And StrComp(currentRibbonId, RIBBON_NAME_, vbTextCompare) = 0 Then
        EnsureDatabasePropertyRemoved db, "CustomRibbonID", "DisableRibbon"
        If TryMarkRibbonRestartRequired(restartWarningText) = False Then Err.Raise vbObjectError + 7223, "DisableRibbon", "The " & DEFAULT_UTILITY_LONG_NAME & " ribbon was disabled, but the required database-reopen marker could not be persisted." & vbCrLf & restartWarningText
    End If
    If showMessage Then
        CoreShowMessage _
            IIf(propertyExists And StrComp(currentRibbonId, RIBBON_NAME_, vbTextCompare) = 0, DEFAULT_UTILITY_LONG_NAME & " ribbon disabled. Close and reopen the database.", DEFAULT_UTILITY_LONG_NAME & " ribbon was not the active database ribbon; no setting was changed."), _
            vbInformation, _
            DEFAULT_UI_MESSAGE_TITLE, _
            "ShowADCMessage", _
            "UI"
    End If
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    messageText = "The " & DEFAULT_UTILITY_LONG_NAME & " ribbon disable operation did not complete cleanly." & vbCrLf & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description:" & vbCrLf & savedErrDescription
    If showMessage Then
        CoreShowErrorMessage _
            messageText, _
            vbCritical, _
            DEFAULT_UI_MESSAGE_TITLE, _
            "DisableRibbon", _
            "UI"
        Exit Sub
    End If
    Err.Raise savedErrNumber, "DisableRibbon", messageText
End Sub

' Purpose: Checks whether ribbon setup plan has all predecessor proof/state required before Setup H2 may proceed.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   detailText (Output; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
' Returns: Boolean: True when ribbon setup plan has all predecessor proof/state required before Setup H2; False while any required predecessor condition is missing or stale.
Public Function RibbonSetupPlanReadyForH2(ByRef detailText As String) As Boolean
    If mRibbonPhasePrepared = False Or Len(mRibbonPhaseToken) = 0 Then
        detailText = "H2 requires a successful same-session H1 Ribbon plan."
        Exit Function
    End If
    If mRibbonPhaseTransactionStarted Then
        detailText = "H2 cannot start because a Ribbon staging transaction is already active."
        Exit Function
    End If
    detailText = "Same-session H1 Ribbon plan is ready. Build required=" & CStr(mRibbonPhaseBuildRequired) & "."
    RibbonSetupPlanReadyForH2 = True
End Function

' Purpose: Checks whether ribbon setup build has all predecessor proof/state required before Setup H3 may proceed.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   detailText (Output; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
' Returns: Boolean: True when ribbon setup build has all predecessor proof/state required before Setup H3; False while any required predecessor condition is missing or stale.
Public Function RibbonSetupBuildReadyForH3(ByRef detailText As String) As Boolean
    If mRibbonPhasePrepared = False Or Len(mRibbonPhaseToken) = 0 Then
        detailText = "H3 requires the same-session H1/H2 Ribbon state."
        Exit Function
    End If
    If mRibbonPhaseStandardCompleted = False Then
        detailText = "H3 requires H2 to complete ADC Ribbon build/validation or the validated H2 no-op."
        Exit Function
    End If
    If mRibbonPhaseStandardRequired And mRibbonPhaseStandardBuilt = False Then
        detailText = "H3 requires the ADC Ribbon staged build that H1 marked required."
        Exit Function
    End If
    If mRibbonPhaseStandardRequired And mRibbonPhaseTransactionStarted = False Then
        detailText = "H3 requires H2's active staged Ribbon transaction."
        Exit Function
    End If
    detailText = "Same-session H2 Ribbon state is ready for H3 commit."
    RibbonSetupBuildReadyForH3 = True
End Function

Private Sub RequireRibbonNavigationPaneGroupField(ByVal fieldName As String)
    If SharedDao_TableFieldExists(GetHostDb(), ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups"), fieldName) = False Then Err.Raise vbObjectError + 6656, "ValidateRibbonDependencies", "The Navigation Pane groups table is missing a ribbon grouping field." & vbCrLf & "Table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_NavigationPaneGroups") & vbCrLf & "Field: " & fieldName & vbCrLf & "Run " & DEFAULT_UTILITY_LONG_NAME & ": Setup/Table Repair first."
End Sub

Private Sub CountRequiredUiManifestGroup(ByVal tabKey As String, ByVal groupLabel As String, ByRef settingsCount As Long, ByRef launcherControlsCount As Long, ByRef navigationPaneCount As Long, ByRef ribbonCount As Long, ByRef formDefinitionsCount As Long)
    If StrComp(tabKey, DEFAULT_RIBBON_TAB_KEY_UI, vbTextCompare) <> 0 Then Exit Sub
    Select Case UCase$(groupLabel)
        Case UCase$(RibbonManagedGroupLabel("SETTINGS")): settingsCount = settingsCount + 1
        Case UCase$(RibbonManagedGroupLabel("LAUNCHER_CONTROLS")): launcherControlsCount = launcherControlsCount + 1
        Case UCase$(RibbonManagedGroupLabel("NAVIGATION_PANE")): navigationPaneCount = navigationPaneCount + 1
        Case UCase$(RibbonManagedGroupLabel("RIBBON")): ribbonCount = ribbonCount + 1
        Case UCase$(RibbonManagedGroupLabel("FORM_DEFINITIONS")): formDefinitionsCount = formDefinitionsCount + 1
    End Select
End Sub

Private Sub ValidateManagedObjectRibbonManifestCoverage(ByVal ribbonName As String, ByVal manifestKeys As Object)
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim objectName As String
    Dim objectTypeCode As String
    Dim normalizedType As String
    Dim expectedExistencePolicy As String
    Dim setupAction As String
    Dim optionalReason As String
    Dim keyText As String
    Dim missingText As String
    Dim missingCount As Long
    Dim optionalAbsentCount As Long
    Dim policyExcludedCount As Long
    Set db = ADCDb()
    sqlText = "SELECT m." & SQL_FIELD_OBJECT_NAME & ",m." & SQL_FIELD_OBJECT_TYPE & "," & _
              ResolvedExpectedExistencePolicySql("o", "c") & " " & "AS " & SQL_FIELD_RESOLVED_EXPECTED_EXISTENCE_POLICY & "," & _
              ResolvedSetupActionSql("o", "c") & " " & "AS " & SQL_FIELD_RESOLVED_SETUP_ACTION & " " & _
              "FROM ([" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID) & "] AS m INNER JOIN [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ObjectCategories") & "] AS c ON m.[" & FIELD_CATEGORY_CODE & "]=c.[" & FIELD_CATEGORY_CODE & "]) LEFT JOIN [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECT_OVERRIDES_TABLE_LOGICAL_ID) & "] AS o ON m." & SQL_FIELD_ID & "=o.[" & FIELD_MANAGED_OBJECT_ID & "] " & _
              "WHERE Nz(m.[" & FIELD_MANAGED_OBJECT_STATUS & "]," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & " AND Nz(c." & SQL_FIELD_OBJECT_CATEGORY_STATUS & "," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & " " & _
              "AND UCase(Trim(Nz(m." & SQL_FIELD_OBJECT_TYPE & ",''))) In ('TABLE','QUERY','FORM','REPORT','MACRO') " & _
              "ORDER BY " & OrderByPolicyText("RIBBON_MANIFEST_COVERAGE") & ";"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    Do While rs.EOF = False
        objectName = CStr(Nz(rs.fields(FIELD_OBJECT_NAME).value, vbNullString))
        objectTypeCode = Trim$(CStr(rs.fields(FIELD_OBJECT_TYPE).value))
        normalizedType = NormalizedObjectTypeCode(objectTypeCode)
        expectedExistencePolicy = Trim$(CStr(Nz(rs.fields(FIELD_RESOLVED_EXPECTED_EXISTENCE_POLICY).value, vbNullString)))
        setupAction = Trim$(CStr(Nz(rs.fields(FIELD_RESOLVED_SETUP_ACTION).value, vbNullString)))
        optionalReason = vbNullString
        If Len(Trim$(objectName)) = 0 Or Len(normalizedType) = 0 Then
            missingCount = missingCount + 1
            missingText = SharedText_AppendNonBlankLine(missingText, "Blank or unsupported managed-object ribbon identity. Type=" & objectTypeCode & "; Object=" & objectName)
        ElseIf ManagedObjectIsRibbonExcludedByPolicy(objectName, objectTypeCode) Then
            policyExcludedCount = policyExcludedCount + 1
        ElseIf RibbonObjectExists(normalizedType, objectName) = False Then
            optionalReason = vbNullString
            If Len(optionalReason) = 0 Then optionalReason = ManagedObjectOptionalMissingReason(expectedExistencePolicy, setupAction)
            If Len(optionalReason) > 0 Then
                optionalAbsentCount = optionalAbsentCount + 1
            Else
                missingCount = missingCount + 1
                missingText = SharedText_AppendNonBlankLine(missingText, objectTypeCode & " " & objectName & " (required physical object missing)")
            End If
        Else
            keyText = UCase$(normalizedType) & "|" & objectName
            If manifestKeys.Exists(keyText) = False Then
                missingCount = missingCount + 1
                missingText = SharedText_AppendNonBlankLine(missingText, objectTypeCode & " " & objectName)
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    If missingCount > 0 Then Err.Raise vbObjectError + 7248, "ValidateManagedObjectRibbonManifestCoverage", "One or more required active managed objects are not represented in the persisted ribbon manifest." & vbCrLf & "Ribbon: " & ribbonName & vbCrLf & "Missing count: " & CStr(missingCount) & vbCrLf & missingText
    LogTraceMessage TRACE_LEVEL_INFO, "RibbonManifest", "ValidateManagedObjectRibbonManifestCoverage", "Complete active managed-object ribbon representation verified.", "Ribbon: " & ribbonName & vbCrLf & "Optional uninstalled rows skipped: " & CStr(optionalAbsentCount) & vbCrLf & "Explicit Ribbon-policy exclusions skipped: " & CStr(policyExcludedCount), "Verify managed-object ribbon coverage", 0, vbNullString, vbNullString, ribbonName, "Ribbon", "CodeDb", "Ribbon"
End Sub

' ===============================================================================================================================================================================
' RIBBON RUNTIME CALLBACK / TOOLTIP / CATALOG VALIDATION PROCEDURES
' ===============================================================================================================================================================================

' Purpose: Performs the public ribbon open managed object operation using the routine's existing ADC workflow and side effects.
' Public role: Ribbon callback entry point. Office/Access invokes this Public routine from Ribbon XML; keeping it Public is required for callback resolution.
' Parameters:
'   control (Input; IRibbonControl): Office Ribbon control that triggered the callback; its Id/Tag identifies the requested Ribbon action or metadata lookup.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub RibbonOpenManagedObject(ByVal control As IRibbonControl)
    'Intent: Ribbon callback that resolves the clicked managed-object control and opens the configured Access object.
    If mTerminalTeardownActive Then Exit Sub
    On Error GoTo ErrHandler
    Dim tagText As String
    Dim objectTypeCode As String
    Dim objectName As String
    Dim messageText As String
    Dim ignoredRibbonName As String
    Dim ignoredRibbonTabKey As String
    Dim ignoredScreenTipCaption As String
    Dim resolutionTimer As Single
    Dim openTimer As Single
    Dim resolutionElapsedSeconds As Double
    Dim openElapsedSeconds As Double
    Dim reportDiagnosticText As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim failureStage As String
    failureStage = "Read ribbon control tag"
    resolutionTimer = Timer
    tagText = CStr(control.Tag)
    failureStage = "Resolve ribbon target"
    If RibbonMappedButtonPresentation(tagText, ignoredRibbonName, ignoredRibbonTabKey, objectTypeCode, objectName, ignoredScreenTipCaption) = False Then Err.Raise vbObjectError + 6650, "RibbonOpenManagedObject", "Ribbon button tag is invalid or unmapped: " & tagText
    resolutionElapsedSeconds = ElapsedSecondsFromTimer(resolutionTimer)
    On Error Resume Next
    LogTraceMessage TRACE_LEVEL_INFO, "Ribbon", "RibbonOpenManagedObject", "Opening " & DEFAULT_UTILITY_LONG_NAME & " ribbon object.", "Tag: " & tagText & vbCrLf & "Object type: " & objectTypeCode & vbCrLf & "Object name: " & objectName & vbCrLf & "Ribbon target resolution elapsed: " & FormattedElapsedSeconds(resolutionElapsedSeconds), vbNullString, 0, vbNullString, vbNullString, objectName, objectTypeCode, "CodeDb", "Ribbon"
    Err.Clear
    On Error GoTo ErrHandler
    openTimer = Timer
    Select Case UCase$(objectTypeCode)
        Case UCase$(ADCAT_OBJECT_TYPE_MACRO)
            failureStage = "DoCmd.RunMacro"
            DoCmd.RunMacro objectName
        Case UCase$(ADCAT_OBJECT_TYPE_TABLE)
            failureStage = "DoCmd.OpenTable"
            DoCmd.OpenTable objectName, acViewNormal
        Case UCase$(ADCAT_OBJECT_TYPE_QUERY)
            failureStage = "DoCmd.OpenQuery"
            DoCmd.OpenQuery objectName, acViewNormal
        Case UCase$(ADCAT_OBJECT_TYPE_FORM)
            failureStage = "DoCmd.OpenForm"
            DoCmd.OpenForm objectName, acNormal
        Case UCase$(ADCAT_OBJECT_TYPE_REPORT)
            failureStage = "DoCmd.OpenReport acViewReport"
            On Error Resume Next
            LogTraceMessage TRACE_LEVEL_INFO, "ReportOpen", "RibbonOpenManagedObject", "Ribbon report open dispatch starting.", "Report: " & objectName & vbCrLf & "Requested view: Report View (acViewReport)" & vbCrLf & "Tag: " & tagText & vbCrLf & "Ribbon target resolution elapsed: " & FormattedElapsedSeconds(resolutionElapsedSeconds), "Open managed report", 0, vbNullString, vbNullString, objectName, ADCAT_OBJECT_TYPE_REPORT, "CodeDb", "Performance"
            Err.Clear
            On Error GoTo ErrHandler
            openTimer = Timer
            DoCmd.OpenReport objectName, acViewReport
            openElapsedSeconds = ElapsedSecondsFromTimer(openTimer)
            reportDiagnosticText = "Report: " & objectName & vbCrLf & "Requested view: Report View (acViewReport)" & vbCrLf & "DoCmd.OpenReport elapsed: " & FormattedElapsedSeconds(openElapsedSeconds)
            On Error Resume Next
            If CurrentProject.AllReports(objectName).IsLoaded Then
                reportDiagnosticText = reportDiagnosticText & vbCrLf & "Loaded after return: True" & vbCrLf & "CurrentView: " & CStr(reports(objectName).CurrentView) & vbCrLf & "AllowReportView: " & CStr(reports(objectName).AllowReportView) & vbCrLf & "RecordSource: " & Nz(reports(objectName).recordSource, vbNullString) & vbCrLf & "FilterOn: " & CStr(reports(objectName).FilterOn) & vbCrLf & "Filter: " & Nz(reports(objectName).Filter, vbNullString) & vbCrLf & "OrderByOn: " & CStr(reports(objectName).OrderByOn) & vbCrLf & "OrderBy: " & Nz(reports(objectName).OrderBy, vbNullString)
            Else
                reportDiagnosticText = reportDiagnosticText & vbCrLf & "Loaded after return: False"
            End If
            If Err.Number <> 0 Then
                reportDiagnosticText = reportDiagnosticText & vbCrLf & "Post-open diagnostic read failed without failing the report open. Error " & CStr(Err.Number) & ": " & Err.Description
                Err.Clear
            End If
            On Error Resume Next
            LogTraceMessage TRACE_LEVEL_INFO, "ReportOpen", "RibbonOpenManagedObject", "Ribbon report open returned successfully.", reportDiagnosticText, "Open managed report", 0, vbNullString, vbNullString, objectName, ADCAT_OBJECT_TYPE_REPORT, "CodeDb", "Performance"
            Err.Clear
            On Error GoTo ErrHandler
        Case Else
            Err.Raise vbObjectError + 6651, "RibbonOpenManagedObject", "Unsupported ribbon object type: " & objectTypeCode
    End Select
    Exit Sub
ErrHandler:
    If mTerminalTeardownActive Then Exit Sub
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    openElapsedSeconds = 0
    If openTimer <> 0 Then openElapsedSeconds = ElapsedSecondsFromTimer(openTimer)
    messageText = "Could not open the " & DEFAULT_UTILITY_LONG_NAME & " ribbon object." & vbCrLf & vbCrLf & "Procedure: RibbonOpenManagedObject" & vbCrLf & "Failure stage: " & failureStage & vbCrLf & "Tag: " & tagText & vbCrLf & "Object type: " & objectTypeCode & vbCrLf & "Object name: " & objectName & vbCrLf & "Open-call elapsed before failure: " & FormattedElapsedSeconds(openElapsedSeconds) & vbCrLf & "Error " & CStr(savedErrNumber) & ": " & savedErrDescription & TableDiagnosticsLocationForMessage()
    LogTraceMessage TRACE_LEVEL_ERROR, "Ribbon", "RibbonOpenManagedObject", "Could not open the " & DEFAULT_UTILITY_LONG_NAME & " ribbon object.", messageText, failureStage, savedErrNumber, savedErrSource, savedErrDescription, objectName, objectTypeCode, "CodeDb", "Ribbon"
    CoreShowErrorMessage _
        messageText, _
        vbCritical, _
        DEFAULT_UI_MESSAGE_TITLE, _
        "RibbonOpenManagedObject", _
        "UI"
End Sub

Private Sub RibbonAddCatalogManagedObjectRows(ByRef rows As Collection, ByRef sourceDetailText As String)
    On Error GoTo ErrHandler
    Dim rowData As Variant
    Dim rowNumber As Long
    Dim rowCount As Long
    Dim objectName As String
    Dim objectTypeCode As String
    Dim tooltipText As String
    Dim tabKey As String
    Dim tabLabel As String
    Dim groupLabel As String
    Dim managedRibbonTitle As String
    Dim skipReason As String
    Dim expectedExistencePolicy As String
    Dim setupAction As String
    Dim optionalReason As String
    Dim addedCount As Long
    Dim optionalAbsentCount As Long
    Dim policyExcludedCount As Long
    Dim stageText As String
    Dim placementIndex As Object
    If mRibbonCatalogSourceRows Is Nothing Then PrepareRibbonCatalogSourceSnapshot
    rowCount = mRibbonCatalogSourceRows.count
    stageText = "Add catalog ribbon placements from shared H1 source snapshot"
    ' INTENTIONAL OPERATION-SCOPED KEYED LOOKUP: one identity -> tab/group index replaces repeated scans of the growing ordered row collection; mRibbonCatalogSourceRows remains the source authority.
    Set placementIndex = CreateObject("Scripting.Dictionary")
    placementIndex.CompareMode = vbTextCompare
    For Each rowData In mRibbonCatalogSourceRows
        rowNumber = rowNumber + 1
        objectName = CStr(rowData(0))
        objectTypeCode = CStr(rowData(1))
        expectedExistencePolicy = CStr(rowData(3))
        setupAction = CStr(rowData(4))
        tooltipText = CStr(rowData(5))
        tabKey = CStr(rowData(7))
        tabLabel = CStr(rowData(8))
        groupLabel = CStr(rowData(9))
        stageText = "Add shared catalog ribbon row " & CStr(rowNumber) & " of " & CStr(rowCount)
        If rowNumber = 1 Or rowNumber Mod 50 = 0 Or rowNumber = rowCount Then
            PrintRibbonStatus "Collect shared ribbon catalog row " & CStr(rowNumber) & " of " & CStr(rowCount) & ": " & objectTypeCode & " [" & objectName & "]."
            RibbonRequireSafeWorkContinue "RibbonAddCatalogManagedObjectRows", objectTypeCode & " [" & objectName & "]"
        End If
        If ManagedObjectIsRibbonExcludedByPolicy(objectName, objectTypeCode) Then
            policyExcludedCount = policyExcludedCount + 1
        Else
            If Len(tooltipText) = 0 Then tooltipText = RibbonTooltipForObject(objectName, objectTypeCode)
            managedRibbonTitle = CStr(rowData(17))
            If StrComp(objectTypeCode, ADCAT_OBJECT_TYPE_MACRO, vbTextCompare) = 0 And CLng(rowData(14)) > 0 Then managedRibbonTitle = RibbonSystemMacroDisplayedTitleDefault(managedRibbonTitle, objectName)
            skipReason = vbNullString
            If RibbonAddRowIfObjectExists(rows, placementIndex, tabKey, tabLabel, groupLabel, objectTypeCode, objectName, tooltipText, skipReason, CLng(rowData(10)), CLng(rowData(11)), managedRibbonTitle) = False Then
                optionalReason = vbNullString
                If StrComp(skipReason, RIBBON_SKIP_REASON_MISSING, vbTextCompare) = 0 Then
                    If Len(optionalReason) = 0 Then optionalReason = ManagedObjectOptionalMissingReason(expectedExistencePolicy, setupAction)
                End If
                If Len(optionalReason) > 0 Then
                    optionalAbsentCount = optionalAbsentCount + 1
                    LogTraceMessage TRACE_LEVEL_DETAIL, "Ribbon", "RibbonAddCatalogManagedObjectRows", "Optional cataloged Ribbon target is not installed; no Ribbon row was emitted.", "Object: " & objectName & vbCrLf & "Type: " & objectTypeCode & vbCrLf & "Tab: " & tabLabel & vbCrLf & "Group: " & groupLabel & vbCrLf & "Optional reason: " & optionalReason, stageText, 0, vbNullString, vbNullString, objectName, objectTypeCode, "CodeDb", "Ribbon"
                Else
                    Err.Raise vbObjectError + 6663, "RibbonAddCatalogManagedObjectRows", "A cataloged managed object could not be added to its primary ribbon placement." & vbCrLf & "Object: " & objectName & vbCrLf & "Type: " & objectTypeCode & vbCrLf & "Tab: " & tabLabel & vbCrLf & "Group: " & groupLabel & vbCrLf & "Reason: " & skipReason
                End If
            Else
                addedCount = addedCount + 1
            End If
        End If
    Next rowData
    RibbonAppendSourceDetail sourceDetailText, "shared catalog snapshot rows=" & CStr(rowCount) & "; catalog primary placements added=" & CStr(addedCount) & "; optional uninstalled rows skipped=" & CStr(optionalAbsentCount) & "; explicit hidden-form ribbon exclusions=" & CStr(policyExcludedCount) & "; universal managed-object ribbon coverage enforced=True except authorized optional absence and explicit hidden-form policy; duplicate placement lookup=single operation-scoped keyed index"
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim diagnosticText As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    diagnosticText = "Could not add catalog-managed ribbon rows from the shared H1 source snapshot." & vbCrLf & "Stage: " & stageText & vbCrLf & "Catalog row: " & CStr(rowNumber) & " of " & CStr(rowCount) & vbCrLf & "Object: " & IIf(Len(objectName) = 0, "(none)", objectName) & vbCrLf & "Type: " & IIf(Len(objectTypeCode) = 0, "(none)", objectTypeCode) & vbCrLf & "Tab: " & IIf(Len(tabLabel) = 0, "(none)", tabLabel) & vbCrLf & "Group: " & IIf(Len(groupLabel) = 0, "(none)", groupLabel) & vbCrLf & "Original Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Original Err.Source: " & savedErrSource & vbCrLf & "Original Err.Description: " & savedErrDescription
    Err.Raise savedErrNumber, "RibbonAddCatalogManagedObjectRows", diagnosticText
End Sub

Private Sub ValidateRibbonManagedObjectInventoryPreflight()
    ' Validates active cataloged Ribbon targets and authoritative placement metadata. Physical Navigation Pane membership is intentionally excluded; Phase I owns that proof.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rowData As Variant
    Dim rowNumber As Long
    Dim rowCount As Long
    Dim objectName As String
    Dim objectTypeCode As String
    Dim normalizedType As String
    Dim groupName As String
    Dim queryFinding As String
    Dim stageText As String
    Dim findings As Collection
    Dim expectedExistencePolicy As String
    Dim setupAction As String
    Dim optionalReason As String
    Dim optionalMissingCount As Long
    Dim physicalObjectExists As Boolean
    Dim releaseSnapshot As Boolean
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: findings retains independent Ribbon-preflight evidence; it does not own catalog or Navigation Pane state.
    Set findings = New Collection
    If mRibbonCatalogSourceRows Is Nothing Then
        PrepareRibbonCatalogSourceSnapshot
        releaseSnapshot = (Len(mRibbonPhaseToken) = 0)
    End If
    stageText = "Validate shared managed-object logical catalog snapshot"
    Set db = ADCDb()
    rowCount = mRibbonCatalogPreflightRowCount
    For Each rowData In mRibbonCatalogSourceRows
        If SharedValue_NzLong(rowData(15), 0) <> 0 Then
            RibbonSetupControlPump
            rowNumber = rowNumber + 1
            objectName = CStr(rowData(0))
            objectTypeCode = CStr(rowData(1))
            normalizedType = objectTypeCode
            expectedExistencePolicy = CStr(rowData(3))
            setupAction = CStr(rowData(4))
            groupName = vbNullString
            physicalObjectExists = False
            stageText = "Validate shared catalog row " & CStr(rowNumber) & " of " & CStr(rowCount) & " > normalize object metadata"
            If Len(objectName) = 0 Or Len(normalizedType) = 0 Then
                findings.Add "Catalog row has a blank object name or unsupported object type." & vbCrLf & "Object: " & IIf(Len(objectName) = 0, "(blank)", objectName) & vbCrLf & "Type: " & IIf(Len(objectTypeCode) = 0, "(blank)", objectTypeCode)
            Else
                stageText = "Validate shared catalog row " & CStr(rowNumber) & " of " & CStr(rowCount) & " > verify physical object"
                physicalObjectExists = RibbonObjectExists(normalizedType, objectName)
                If physicalObjectExists = False Then
                    optionalReason = vbNullString
                    If Len(optionalReason) = 0 Then optionalReason = ManagedObjectOptionalMissingReason(expectedExistencePolicy, setupAction)
                    If Len(optionalReason) > 0 Then
                        optionalMissingCount = optionalMissingCount + 1
                        LogTraceMessage TRACE_LEVEL_DETAIL, "Preflight", "ValidateRibbonManagedObjectInventoryPreflight", "Optional cataloged Ribbon target is not installed; physical target verification was skipped.", "Object: " & objectName & vbCrLf & "Expected type: " & objectTypeCode & vbCrLf & "Category: " & CStr(rowData(2)) & vbCrLf & "Expected existence: " & expectedExistencePolicy & vbCrLf & "Setup action: " & setupAction & vbCrLf & "Optional reason: " & optionalReason, stageText, 0, vbNullString, vbNullString, objectName, objectTypeCode, "CodeDb", "UI"
                    Else
                        findings.Add "Required cataloged Ribbon target is missing or exists under the wrong Access object type." & vbCrLf & "Object: " & objectName & vbCrLf & "Expected type: " & objectTypeCode & vbCrLf & "Category: " & CStr(rowData(2)) & vbCrLf & "Expected existence: " & expectedExistencePolicy & vbCrLf & "Setup action: " & setupAction
                    End If
                Else
                    If StrComp(normalizedType, ADCAT_OBJECT_TYPE_QUERY, vbTextCompare) = 0 Then
                        stageText = "Validate shared catalog row " & CStr(rowNumber) & " of " & CStr(rowCount) & " > verify query definition"
                        queryFinding = RibbonQueryDefinitionFinding(db, objectName)
                        If Len(queryFinding) > 0 Then findings.Add queryFinding
                    End If
                    stageText = "Validate shared catalog row " & CStr(rowNumber) & " of " & CStr(rowCount) & " > resolve Navigation Pane placement metadata"
                    groupName = ManagedObjectDefaultNavigationPaneGroupName(objectName, objectTypeCode)
                    If Len(groupName) = 0 Then findings.Add "Cataloged object has no resolved Navigation Pane placement metadata." & vbCrLf & "Object: " & objectName & vbCrLf & "Type: " & objectTypeCode
                End If
            End If
            RibbonRequireSafeWorkContinue "ValidateRibbonManagedObjectInventoryPreflight", objectTypeCode & " [" & objectName & "]"
        End If
    Next rowData
    If optionalMissingCount > 0 Then LogTraceMessage TRACE_LEVEL_INFO, "Preflight", "ValidateRibbonManagedObjectInventoryPreflight", "Managed-object Ribbon preflight classification completed from the shared H1 logical source snapshot.", "Optional uninstalled rows skipped=" & CStr(optionalMissingCount) & vbCrLf & "Catalog rows checked=" & CStr(rowCount) & vbCrLf & "Physical Navigation Pane membership checked=False", stageText, 0, vbNullString, vbNullString, vbNullString, "ManagedObjectCatalog", "CodeDb", "UI"
    If findings.count > 0 Then
        LogRibbonPreflightFindings findings, stageText, rowNumber, rowCount
        Err.Raise vbObjectError + 6665, "ValidateRibbonManagedObjectInventoryPreflight", RibbonPreflightFindingsText(findings)
    End If
    If releaseSnapshot Then
        Set mRibbonCatalogSourceRows = Nothing
        mRibbonCatalogPreflightRowCount = 0
        Set mRibbonSystemMacroTargetCache = Nothing
        Set mRibbonCatalogMetadataCache = Nothing
        Set mRibbonCatalogAliasMetadataCache = Nothing
    End If
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim diagnosticText As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If releaseSnapshot Then
        Set mRibbonCatalogSourceRows = Nothing
        mRibbonCatalogPreflightRowCount = 0
        Set mRibbonSystemMacroTargetCache = Nothing
        Set mRibbonCatalogMetadataCache = Nothing
        Set mRibbonCatalogAliasMetadataCache = Nothing
    End If
    If IsCancellationError(savedErrNumber) Then Err.Raise savedErrNumber, savedErrSource, savedErrDescription
    If InStr(1, savedErrDescription, "Blocking finding 1 of", vbTextCompare) > 0 Then Err.Raise savedErrNumber, "ValidateRibbonManagedObjectInventoryPreflight", savedErrDescription
    diagnosticText = "Could not validate the managed-object Ribbon inventory from the logical catalog." & vbCrLf & "Stage: " & stageText & vbCrLf & "Catalog row: " & CStr(rowNumber) & " of " & CStr(rowCount) & vbCrLf & "Object: " & IIf(Len(objectName) = 0, "(none)", objectName) & vbCrLf & "Type: " & IIf(Len(objectTypeCode) = 0, "(none)", objectTypeCode) & vbCrLf & "Original Err.Source: " & savedErrSource & vbCrLf & "Original Err.Description: " & savedErrDescription
    Err.Raise savedErrNumber, "ValidateRibbonManagedObjectInventoryPreflight", diagnosticText
End Sub

Private Sub ValidateRequiredSystemManifestGroups(ByVal ribbonName As String, ByVal systemGroupCounts As Object, ByVal unexpectedCount As Long, ByVal misplacedCount As Long, ByVal diagnosticText As String, ByVal diagnosticRowCount As Long, ByVal diagnosticTruncated As Boolean, ByVal systemRibbonGroupDefinitions As Collection, ByVal systemMacroDefinitions As Collection)
    Dim Definition As Variant
    Dim groupIdentity As String
    Dim expectedMacroCount As Long
    Dim actualGroupCount As Long
    Dim missingRequiredGroup As Boolean
    Dim detailText As String
    Dim canonicalCounts As Object
    Set canonicalCounts = SystemMacroCountsByRibbonGroupIdentity(systemMacroDefinitions)
    For Each Definition In systemRibbonGroupDefinitions
        groupIdentity = CStr(Definition(0))
        expectedMacroCount = 0
        If canonicalCounts.Exists(groupIdentity) Then expectedMacroCount = CLng(canonicalCounts(groupIdentity))
        actualGroupCount = SystemManifestGroupCount(systemGroupCounts, groupIdentity)
        If expectedMacroCount > 0 And actualGroupCount = 0 Then
            missingRequiredGroup = True
            Exit For
        End If
    Next Definition
    If missingRequiredGroup = False And unexpectedCount = 0 And misplacedCount = 0 Then Exit Sub
    detailText = "A persisted ribbon manifest does not contain the canonical System Ribbon layout." & vbCrLf
    detailText = detailText & "Ribbon: " & ribbonName & vbCrLf
    detailText = detailText & RequiredSystemManifestCountSummary(systemGroupCounts, misplacedCount, unexpectedCount, systemRibbonGroupDefinitions)
    detailText = detailText & vbCrLf & "Affected rows captured: " & CStr(diagnosticRowCount)
    If diagnosticTruncated Then detailText = detailText & vbCrLf & "Affected-row list was truncated after 25 rows."
    If Len(diagnosticText) > 0 Then detailText = detailText & vbCrLf & vbCrLf & "Affected System Ribbon rows:" & vbCrLf & diagnosticText
    Err.Raise vbObjectError + 7190, "ValidateRequiredSystemManifestGroups", detailText
End Sub

Private Sub ValidateRequiredUiManifestGroups(ByVal ribbonName As String, ByVal settingsCount As Long, ByVal launcherControlsCount As Long, ByVal navigationPaneCount As Long, ByVal ribbonCount As Long, ByVal formDefinitionsCount As Long)
    If settingsCount = 0 Or launcherControlsCount = 0 Or navigationPaneCount = 0 Or ribbonCount = 0 Or formDefinitionsCount = 0 Then Err.Raise vbObjectError + 7186, "ValidateRequiredUiManifestGroups", "A persisted ribbon manifest does not contain all five required current " & DEFAULT_UTILITY_SHORT_NAME & " UI groups. Group labels must match the current authoritative Ribbon definition." & vbCrLf & "Ribbon: " & ribbonName & vbCrLf & "Settings rows: " & CStr(settingsCount) & vbCrLf & "Custom Macro Launcher rows: " & CStr(launcherControlsCount) & vbCrLf & "Navigation Pane rows: " & CStr(navigationPaneCount) & vbCrLf & "Ribbon rows: " & CStr(ribbonCount) & vbCrLf & "Form Definitions rows: " & CStr(formDefinitionsCount) & vbCrLf & vbCrLf & "Persisted " & DEFAULT_UTILITY_SHORT_NAME & " UI group inventory:" & vbCrLf & RibbonUiManifestGroupDiagnostic(ribbonName)
    ValidateRequiredUIManifestInventory ribbonName
End Sub

Private Sub ValidateRibbonTooltipCallbackContract(ByVal ribbonXml As String)
    Dim buttonCount As Long
    Dim screenTipCallbackCount As Long
    Dim superTipCallbackCount As Long
    buttonCount = RibbonTextOccurrenceCount(ribbonXml, "<button ")
    screenTipCallbackCount = RibbonTextOccurrenceCount(ribbonXml, " getScreentip=""" & RIBBON_CALLBACK_GET_SCREENTIP & """")
    superTipCallbackCount = RibbonTextOccurrenceCount(ribbonXml, " getSupertip=""" & RIBBON_CALLBACK_GET_SUPERTIP & """")
    If InStr(1, ribbonXml, " screentip=""", vbTextCompare) > 0 Or InStr(1, ribbonXml, " supertip=""", vbTextCompare) > 0 Then Err.Raise vbObjectError + 7107, "ValidateRibbonTooltipCallbackContract", "Generated ribbon XML contains literal ScreenTip or SuperTip text. Tooltip text must remain database-backed through callbacks."
    If buttonCount <= 0 Then Err.Raise vbObjectError + 7108, "ValidateRibbonTooltipCallbackContract", "Generated ribbon XML contains no button controls."
    If screenTipCallbackCount <> buttonCount Or superTipCallbackCount <> buttonCount Then Err.Raise vbObjectError + 7109, "ValidateRibbonTooltipCallbackContract", "Generated ribbon XML does not assign both tooltip callbacks to every button." & vbCrLf & "Buttons: " & CStr(buttonCount) & vbCrLf & "getScreentip callbacks: " & CStr(screenTipCallbackCount) & vbCrLf & "getSupertip callbacks: " & CStr(superTipCallbackCount)
End Sub

Private Function RibbonTooltipForObject(ByVal objectName As String, ByVal objectTypeCode As String) As String
    On Error GoTo ErrHandler
    Dim cachedMetadata As Variant
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim tooltipText As String
    If Len(objectName) = 0 Then Exit Function
    If Not mRibbonCatalogMetadataCache Is Nothing Then
        If RibbonTryDisplayCatalogMetadata(objectName, objectTypeCode, cachedMetadata) Then
            tooltipText = CStr(cachedMetadata(3))
            If Len(tooltipText) = 0 Then tooltipText = CStr(cachedMetadata(4))
        End If
        RibbonTooltipForObject = tooltipText
        Exit Function
    End If
    If SharedDao_TableExists(GetHostDb(), ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID)) = False Then Exit Function
    If SharedDao_TableFieldExists(GetHostDb(), ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID), FIELD_OBJECT_NAME) = False Then Exit Function
    Set db = GetHostDb()
    sqlText = RibbonTooltipSql(objectName, objectTypeCode)
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    If rs.EOF = False Then
        tooltipText = Trim$(RibbonRecordsetText(rs, FIELD_RIBBON_TOOLTIP))
        If Len(tooltipText) = 0 Then tooltipText = Trim$(RibbonRecordsetText(rs, FIELD_OBJECT_DESCRIPTION))
    End If
    rs.Close
    RibbonTooltipForObject = tooltipText
    Exit Function
ErrHandler:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
    RibbonTooltipForObject = vbNullString
End Function

Private Function RibbonTooltipSql(ByVal objectName As String, ByVal objectTypeCode As String) As String
    Dim selectText As String
    Dim whereText As String
    Dim orderText As String
    Dim managedObjectsTableName As String
    Dim overrideTableName As String
    managedObjectsTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID)
    overrideTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECT_OVERRIDES_TABLE_LOGICAL_ID)
    selectText = "m." & SQL_FIELD_OBJECT_NAME
    If SharedDao_TableFieldExists(GetHostDb(), managedObjectsTableName, FIELD_RIBBON_TOOLTIP) Then selectText = selectText & ", m." & SQL_FIELD_RIBBON_TOOLTIP
    If SharedDao_TableFieldExists(GetHostDb(), managedObjectsTableName, FIELD_OBJECT_DESCRIPTION) Then selectText = selectText & ", m." & SQL_FIELD_OBJECT_DESCRIPTION
    whereText = "(m." & SQL_FIELD_OBJECT_NAME & "=" & SharedDao_SqlQuote(objectName) & " OR o.[" & FIELD_TARGET_OBJECT_NAME_OVERRIDE & "]=" & SharedDao_SqlQuote(objectName) & ")"
    If Len(objectTypeCode) > 0 And SharedDao_TableFieldExists(GetHostDb(), managedObjectsTableName, FIELD_OBJECT_TYPE) Then whereText = whereText & " AND UCase(Nz(m." & SQL_FIELD_OBJECT_TYPE & ",''))=" & SharedDao_SqlQuote(UCase$(objectTypeCode))
    If SharedDao_TableFieldExists(GetHostDb(), managedObjectsTableName, FIELD_MANAGED_OBJECT_STATUS) Then whereText = whereText & " AND Nz(m.[" & FIELD_MANAGED_OBJECT_STATUS & "]," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ""
    orderText = RibbonMetadataCandidateOrderByText(objectName, SharedDao_TableFieldExists(GetHostDb(), managedObjectsTableName, FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER))
    RibbonTooltipSql = "SELECT TOP 1 " & selectText & " FROM [" & managedObjectsTableName & "] AS m LEFT JOIN [" & overrideTableName & "] AS o ON m." & SQL_FIELD_ID & "=o.[" & FIELD_MANAGED_OBJECT_ID & "] WHERE " & whereText & " ORDER BY " & orderText & ";"
End Function

' Purpose: Performs the public ribbon get screen tip operation using the routine's existing ADC workflow and side effects.
' Public role: Ribbon callback entry point. Office/Access invokes this Public routine from Ribbon XML; keeping it Public is required for callback resolution.
' Parameters:
'   control (Input; IRibbonControl): Office Ribbon control that triggered the callback; its Id/Tag identifies the requested Ribbon action or metadata lookup.
'   returnedVal (Output/in-out; Variant): Receives the value returned to Access/Office through the callback contract.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects and the documented ByRef parameter(s): returnedVal.
Public Sub RibbonGetScreenTip(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    'Returns the current database-backed ScreenTip for one generated ribbon button.
    On Error GoTo ErrHandler
    Dim ribbonName As String
    Dim ribbonTabKey As String
    Dim objectTypeCode As String
    Dim objectName As String
    Dim tooltipText As String
    Dim screenTipCaption As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    If RibbonResolveTooltipContext(control, ribbonName, ribbonTabKey, objectTypeCode, objectName, screenTipCaption) = False Then
        returnedVal = vbNullString
        Exit Sub
    End If
    If Len(screenTipCaption) > 0 Then
        returnedVal = RibbonVisibleSeparatorDefault(screenTipCaption)
    Else
        tooltipText = RibbonTooltipForObject(objectName, objectTypeCode)
        returnedVal = RibbonVisibleSeparatorDefault(RibbonButtonScreenTip(ribbonTabKey, objectTypeCode, objectName, tooltipText))
    End If
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    returnedVal = vbNullString
    ReportRibbonTooltipCallbackFailure "RibbonGetScreenTip", control, savedErrNumber, savedErrSource, savedErrDescription
End Sub

' Purpose: Performs the public ribbon get super tip operation using the routine's existing ADC workflow and side effects.
' Public role: Ribbon callback entry point. Office/Access invokes this Public routine from Ribbon XML; keeping it Public is required for callback resolution.
' Parameters:
'   control (Input; IRibbonControl): Office Ribbon control that triggered the callback; its Id/Tag identifies the requested Ribbon action or metadata lookup.
'   returnedVal (Output/in-out; Variant): Receives the value returned to Access/Office through the callback contract.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects and the documented ByRef parameter(s): returnedVal.
Public Sub RibbonGetSuperTip(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    'Returns the current database-backed SuperTip for one generated ribbon button.
    On Error GoTo ErrHandler
    Dim ribbonName As String
    Dim ribbonTabKey As String
    Dim objectTypeCode As String
    Dim objectName As String
    Dim tooltipText As String
    Dim screenTipCaption As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    If RibbonResolveTooltipContext(control, ribbonName, ribbonTabKey, objectTypeCode, objectName, screenTipCaption) = False Then
        returnedVal = vbNullString
        Exit Sub
    End If
    tooltipText = RibbonTooltipForObject(objectName, objectTypeCode)
    If Len(tooltipText) = 0 Then tooltipText = RibbonButtonTooltip(objectTypeCode, objectName, vbNullString)
    returnedVal = RibbonVisibleSeparatorDefault(tooltipText)
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    returnedVal = vbNullString
    ReportRibbonTooltipCallbackFailure "RibbonGetSuperTip", control, savedErrNumber, savedErrSource, savedErrDescription
End Sub

Private Function RibbonGeneratedScreenTipCaption(ByVal objectTypeCode As String, ByVal objectName As String) As String
    Dim sourceText As String
    Dim sourceCount As Long
    Dim captionText As String
    Dim maxLength As Long
    Dim targetMetadata As Variant
    Dim targetTableText As String
    Dim targetQueryText As String
    Dim targetTableCount As Long
    Dim targetQueryCount As Long
    captionText = Trim$(objectName)
    Select Case UCase$(objectTypeCode)
        Case UCase$(ADCAT_OBJECT_TYPE_FORM), UCase$(ADCAT_OBJECT_TYPE_REPORT), UCase$(ADCAT_OBJECT_TYPE_QUERY)
            sourceText = RibbonObjectSourceSummary(objectTypeCode, objectName, sourceCount)
            If Len(sourceText) > 0 Then
                If StrComp(objectTypeCode, ADCAT_OBJECT_TYPE_QUERY, vbTextCompare) = 0 Or sourceCount > 1 Then
                    captionText = captionText & " | Sources: " & sourceText
                Else
                    captionText = captionText & " | Record Source: " & sourceText
                End If
            End If
        Case UCase$(ADCAT_OBJECT_TYPE_MACRO)
            If mRibbonSystemMacroTargetCache Is Nothing Then PrepareRibbonCatalogSourceSnapshot
            If mRibbonSystemMacroTargetCache.Exists(objectName) Then
                targetMetadata = mRibbonSystemMacroTargetCache(objectName)
                targetTableText = RibbonSystemMacroTargetListCaption(CStr(targetMetadata(0)), targetTableCount)
                targetQueryText = RibbonSystemMacroTargetListCaption(CStr(targetMetadata(1)), targetQueryCount)
                If targetTableCount > 0 Then captionText = captionText & " | " & IIf(targetTableCount = 1, "Table: ", "Tables: ") & targetTableText
                If targetQueryCount > 0 Then captionText = captionText & " | " & IIf(targetQueryCount = 1, "Query: ", "Queries: ") & targetQueryText
            End If
    End Select
    maxLength = RIBBON_MAX_BUILDER_TEXT_ATTRIBUTE_LENGTH
    RibbonGeneratedScreenTipCaption = RibbonLimitedCaptionText(captionText, maxLength)
End Function

Private Function RibbonSystemMacroTargetListCaption(ByVal targetNames As String, ByRef targetCount As Long) As String
    Dim orderedNames As Collection
    Dim namePart As Variant
    Dim normalizedName As String
    ' INTENTIONAL ORDERED VIEW: orderedNames preserves the deterministic sequence required by RibbonSystemMacroTargetListCaption; keyed-only or unordered storage would not preserve the consumer semantics.
    Set orderedNames = New Collection
    targetCount = 0
    If Len(targetNames) = 0 Then Exit Function
    For Each namePart In Split(targetNames, "|")
        normalizedName = Trim$(CStr(namePart))
        If Len(normalizedName) > 0 Then
            orderedNames.Add normalizedName
            targetCount = targetCount + 1
        End If
    Next namePart
    RibbonSystemMacroTargetListCaption = SharedCollection_Join(orderedNames, ", ")
End Function

Private Function RibbonResolveTooltipContext(ByVal control As IRibbonControl, ByRef ribbonName As String, ByRef ribbonTabKey As String, ByRef objectTypeCode As String, ByRef objectName As String, ByRef screenTipCaption As String) As Boolean
    Dim tagText As String
    tagText = CStr(control.Tag)
    RibbonResolveTooltipContext = RibbonMappedButtonPresentation(tagText, ribbonName, ribbonTabKey, objectTypeCode, objectName, screenTipCaption)
End Function

Private Function RibbonButtonScreenTip(ByVal tabKey As String, ByVal objectTypeCode As String, ByVal objectName As String, ByVal tooltipText As String) As String
    'Outside ADC1 and ADC2, the tooltip caption is the exact Access object name.
    Dim screenTipText As String
    If StrComp(tabKey, DEFAULT_RIBBON_TAB_KEY_SYSTEM_1, vbTextCompare) = 0 Or StrComp(tabKey, DEFAULT_RIBBON_TAB_KEY_SYSTEM_2, vbTextCompare) = 0 Then
        If StrComp(NormalizedObjectTypeCode(objectTypeCode), ADCAT_OBJECT_TYPE_MACRO, vbTextCompare) = 0 Then screenTipText = objectName
        If Len(screenTipText) = 0 Then screenTipText = RibbonTrimLabel(tooltipText)
    Else
        screenTipText = objectName
    End If
    RibbonButtonScreenTip = screenTipText
End Function

Private Sub ReportRibbonTooltipCallbackFailure(ByVal sourceProcedure As String, ByVal control As IRibbonControl, ByVal errorNumber As Long, ByVal errorSource As String, ByVal errorDescription As String)
    'Reports one distinct callback failure through the terminal ADC error workflow while suppressing repeated reports of the same tooltip failure.
    If mTerminalTeardownActive Then Exit Sub
    On Error Resume Next
    Dim tagText As String
    Dim errorKey As String
    tagText = CStr(control.Tag)
    errorKey = sourceProcedure & RIBBON_TAG_DELIMITER & tagText & RIBBON_TAG_DELIMITER & CStr(errorNumber) & RIBBON_TAG_DELIMITER & errorSource & RIBBON_TAG_DELIMITER & errorDescription
    If StrComp(errorKey, mRibbonLastTooltipCallbackErrorKey, vbBinaryCompare) = 0 Then Exit Sub
    mRibbonLastTooltipCallbackErrorKey = errorKey
    Debug.Print sourceProcedure & " failed for ribbon tag [" & tagText & "]. Error " & CStr(errorNumber) & " from " & errorSource & ": " & errorDescription
    LogTraceMessage TRACE_LEVEL_ERROR, "Ribbon", sourceProcedure, "Ribbon tooltip callback could not resolve its database-backed value.", "Tag: " & tagText & vbCrLf & "Error " & CStr(errorNumber) & " from " & errorSource & ": " & errorDescription, "Resolve ribbon tooltip callback", errorNumber, errorSource, errorDescription, tagText, "RibbonControl", "CodeDb", "Ribbon"
    CoreShowErrorMessage "The ribbon tooltip callback could not resolve its database-backed value." & vbCrLf & vbCrLf & "Tag: " & tagText & vbCrLf & "Err.Number: " & CStr(errorNumber) & vbCrLf & "Err.Source: " & errorSource & vbCrLf & "Err.Description: " & errorDescription, vbCritical, DEFAULT_UI_MESSAGE_TITLE, sourceProcedure, "Ribbon"
    On Error GoTo 0
End Sub

Private Function RibbonButtonTooltip(ByVal objectTypeCode As String, ByVal labelText As String, ByVal ribbonTooltip As String) As String
    'Uses the complete catalog description as the supertip body.
    Dim tooltipText As String
    tooltipText = ribbonTooltip
    If Len(tooltipText) = 0 Then tooltipText = "Open " & LCase$(ManagedObjectRibbonTitleObjectTypeLabel(objectTypeCode)) & " target: " & labelText
    RibbonButtonTooltip = tooltipText
End Function

Private Sub LoadRibbonSystemMacroTargetSnapshot(ByVal db As DAO.Database)
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim systemMacroKey As String
    Dim systemMacroTitle As String
    Dim defaultTitleText As String
    Dim macroTitleOverride As String
    Dim targetTableNames As String
    Dim targetQueryNames As String
    Dim systemMacroDesignator As String
    Set mRibbonSystemMacroTargetCache = CreateObject("Scripting.Dictionary")
    mRibbonSystemMacroTargetCache.CompareMode = vbTextCompare
    sqlText = "SELECT [" & FIELD_SYSTEM_MACRO_KEY & "],[" & FIELD_SYSTEM_MACRO_DESIGNATOR & "],[" & FIELD_DEFAULT_TITLE_TEXT & "],[" & FIELD_MACRO_TITLE_OVERRIDE & "],[" & FIELD_SYSTEM_MACRO_TARGET_TABLE_NAMES & "],[" & FIELD_SYSTEM_MACRO_TARGET_QUERY_NAMES & "] "
    sqlText = sqlText & "FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_SystemMacros") & "] "
    sqlText = sqlText & "WHERE Nz([" & FIELD_SYSTEM_MACRO_STATUS & "]," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & " "
    sqlText = sqlText & "ORDER BY Nz(" & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER & ",9999), [" & FIELD_SYSTEM_MACRO_DESIGNATOR & "], [" & FIELD_DEFAULT_TITLE_TEXT & "], [" & FIELD_SYSTEM_MACRO_KEY & "];"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    Do While rs.EOF = False
        systemMacroKey = Trim$(CStr(Nz(rs.fields(FIELD_SYSTEM_MACRO_KEY).value, vbNullString)))
        systemMacroDesignator = CStr(Nz(rs.fields(FIELD_SYSTEM_MACRO_DESIGNATOR).value, vbNullString))
        macroTitleOverride = CStr(Nz(rs.fields(FIELD_MACRO_TITLE_OVERRIDE).value, vbNullString))
        If Len(macroTitleOverride) > 0 Then
            systemMacroTitle = macroTitleOverride
        Else
            defaultTitleText = CStr(Nz(rs.fields(FIELD_DEFAULT_TITLE_TEXT).value, vbNullString))
            systemMacroTitle = SystemMacroCalculatedTitle(systemMacroDesignator, defaultTitleText, systemMacroKey)
        End If
        targetTableNames = Trim$(CStr(Nz(rs.fields(FIELD_SYSTEM_MACRO_TARGET_TABLE_NAMES).value, vbNullString)))
        targetQueryNames = Trim$(CStr(Nz(rs.fields(FIELD_SYSTEM_MACRO_TARGET_QUERY_NAMES).value, vbNullString)))
        If Len(systemMacroTitle) > 0 Then
            If mRibbonSystemMacroTargetCache.Exists(systemMacroTitle) Then Err.Raise vbObjectError + 7286, "LoadRibbonSystemMacroTargetSnapshot", "Two active MC_Sys rows resolve to the same physical macro title." & vbCrLf & "Macro: " & systemMacroTitle
            mRibbonSystemMacroTargetCache.Add systemMacroTitle, Array(targetTableNames, targetQueryNames, systemMacroKey)
        End If
        rs.MoveNext
    Loop
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
    Set mRibbonSystemMacroTargetCache = Nothing
    On Error GoTo 0
    Err.Raise savedErrNumber, "LoadRibbonSystemMacroTargetSnapshot", "Could not load persisted System Macro table/query target metadata for Ribbon caption generation." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub PrepareRibbonCatalogSourceSnapshot()
    ' Loads the active managed-object snapshot used by normal H1/H2 Ribbon work and by the final H2/H3/J1 source-currentness seal.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim rowData As Variant
    Dim objectName As String
    Dim TargetObjectName As String
    Dim objectTypeCode As String
    Dim exactKey As String
    Dim aliasKey As String
    Dim metadata As Variant
    Dim primaryPlacementFlag As Long
    Dim builtInSystemMacroFlag As Long
    Dim activeCategoryFlag As Long
    Dim loadedRowCount As Long
    Dim startedTimer As Single
    If Not mRibbonCatalogSourceRows Is Nothing And Not mRibbonCatalogMetadataCache Is Nothing And Not mRibbonCatalogAliasMetadataCache Is Nothing And Not mRibbonSystemMacroTargetCache Is Nothing Then Exit Sub
    startedTimer = Timer
    Set mRibbonCatalogSourceRows = New Collection
    mRibbonCatalogPreflightRowCount = 0
    Set mRibbonCatalogMetadataCache = CreateObject("Scripting.Dictionary")
    mRibbonCatalogMetadataCache.CompareMode = vbTextCompare
    Set mRibbonCatalogAliasMetadataCache = CreateObject("Scripting.Dictionary")
    mRibbonCatalogAliasMetadataCache.CompareMode = vbTextCompare
    Set db = ADCDb()
    sqlText = "SELECT m." & SQL_FIELD_OBJECT_NAME & ",o.[" & FIELD_TARGET_OBJECT_NAME_OVERRIDE & "] AS [" & FIELD_TARGET_OBJECT_NAME_OVERRIDE & "],m." & SQL_FIELD_OBJECT_TYPE & ",m.[" & FIELD_CATEGORY_CODE & "],m." & SQL_FIELD_RIBBON_TOOLTIP & ",m." & SQL_FIELD_OBJECT_DESCRIPTION & ",m." & SQL_FIELD_RIBBON_TITLE & ",m.[" & FIELD_RIBBON_STATIC_TAG & "],m.[" & FIELD_CATALOG_OWNER & "],m." & SQL_FIELD_OBJECT_DISPLAY_NAME & ","
    sqlText = sqlText & "m.[" & FIELD_RIBBON_TAB_KEY & "],m.[" & FIELD_RIBBON_TAB_LABEL & "],m.[" & FIELD_RIBBON_GROUP_LABEL & "],m.[" & FIELD_RIBBON_GROUP_SORT_ORDER & "],Nz(m." & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER & ",9999) AS [EffectiveSortOrder],"
    sqlText = sqlText & ResolvedExpectedExistencePolicySql("o", "c") & " AS " & SQL_FIELD_RESOLVED_EXPECTED_EXISTENCE_POLICY & ","
    sqlText = sqlText & ResolvedSetupActionSql("o", "c") & " AS " & SQL_FIELD_RESOLVED_SETUP_ACTION & ","
    sqlText = sqlText & "IIf(Len(Trim(Nz(m.[" & FIELD_RIBBON_TAB_KEY & "],'')))>0 AND Len(Trim(Nz(m.[" & FIELD_RIBBON_TAB_LABEL & "],'')))>0 AND Len(Trim(Nz(m.[" & FIELD_RIBBON_GROUP_LABEL & "],'')))>0 AND Nz(m.[" & FIELD_RIBBON_GROUP_SORT_ORDER & "],0)>0,1,0) AS [PrimaryPlacementFlag],"
    sqlText = sqlText & "IIf(UCase(Trim(Nz(m.[" & FIELD_CATEGORY_CODE & "],'')))=" & SharedDao_SqlQuote(UCase$(ADCAT_CATEGORY_SYSTEM_MACRO)) & ",1,0) AS [BuiltInSystemMacroFlag],"
    sqlText = sqlText & "IIf(Not IsNull(c.[" & FIELD_CATEGORY_CODE & "]) AND Nz(c." & SQL_FIELD_OBJECT_CATEGORY_STATUS & "," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ",1,0) AS [ActiveCategoryFlag] "
    sqlText = sqlText & "FROM ([" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID) & "] AS m LEFT JOIN [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ObjectCategories") & "] AS c ON m.[" & FIELD_CATEGORY_CODE & "]=c.[" & FIELD_CATEGORY_CODE & "]) LEFT JOIN [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECT_OVERRIDES_TABLE_LOGICAL_ID) & "] AS o ON m." & SQL_FIELD_ID & "=o.[" & FIELD_MANAGED_OBJECT_ID & "] "
    sqlText = sqlText & "WHERE Nz(m.[" & FIELD_MANAGED_OBJECT_STATUS & "]," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & " AND UCase(Trim(Nz(m." & SQL_FIELD_OBJECT_TYPE & ",''))) In ('TABLE','QUERY','FORM','REPORT','MACRO') "
    sqlText = sqlText & "ORDER BY Nz(m." & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER & ",9999), m." & SQL_FIELD_OBJECT_TYPE & ", m." & SQL_FIELD_OBJECT_NAME & ", Nz(o.[" & FIELD_TARGET_OBJECT_NAME_OVERRIDE & "],'');"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    Do While rs.EOF = False
        loadedRowCount = loadedRowCount + 1
        objectName = CStr(Nz(rs.fields(FIELD_OBJECT_NAME).value, vbNullString))
        TargetObjectName = CStr(Nz(rs.fields(FIELD_TARGET_OBJECT_NAME_OVERRIDE).value, vbNullString))
        objectTypeCode = NormalizedObjectTypeCode(CStr(rs.fields(FIELD_OBJECT_TYPE).value))
        If loadedRowCount = 1 Or loadedRowCount Mod 50 = 0 Then
            PrintRibbonStatus "Load shared H1 ribbon catalog row " & CStr(loadedRowCount) & ": " & objectTypeCode & " [" & objectName & "]."
            RibbonRequireSafeWorkContinue "PrepareRibbonCatalogSourceSnapshot", objectTypeCode & " [" & objectName & "]"
        End If
        primaryPlacementFlag = CLng(rs.fields("PrimaryPlacementFlag").value)
        builtInSystemMacroFlag = CLng(rs.fields("BuiltInSystemMacroFlag").value)
        activeCategoryFlag = CLng(rs.fields("ActiveCategoryFlag").value)
        If primaryPlacementFlag <> 0 Then
            If Len(Trim$(CStr(Nz(rs.fields(FIELD_RIBBON_STATIC_TAG).value, vbNullString)))) = 0 Then Err.Raise vbObjectError + 7284, "PrepareRibbonCatalogSourceSnapshot", "An active managed-object Ribbon placement has no permanent RibbonStaticTag. Run current catalog repair before Ribbon generation." & vbCrLf & "Object: " & objectName & vbCrLf & "Type: " & objectTypeCode
            If InStr(CStr(rs.fields(FIELD_RIBBON_STATIC_TAG).value), "|") > 0 Or InStr(CStr(rs.fields(FIELD_RIBBON_STATIC_TAG).value), ";") > 0 Then Err.Raise vbObjectError + 7285, "PrepareRibbonCatalogSourceSnapshot", "A permanent RibbonStaticTag contains a reserved delimiter." & vbCrLf & "Object: " & objectName & vbCrLf & "Tag: " & CStr(rs.fields(FIELD_RIBBON_STATIC_TAG).value)
        End If
        'Shared source-row layout: 0 name; 1 type; 2 category; 3 expected existence; 4 setup action; 5 tooltip; 6 description; 7-9 tab/group text; 10 group order; 11 object order; 12 target override; 13 placement
        'flag; 14 system-macro flag; 15 active-category preflight flag; 16 permanent Ribbon tag; 17 managed Ribbon title; 18 CatalogOwner; 19 canonical ObjectDisplayName.
        rowData = Array(objectName, objectTypeCode, _
            Trim$(CStr(Nz(rs.fields(FIELD_CATEGORY_CODE).value, vbNullString))), _
            Trim$(CStr(Nz(rs.fields(FIELD_RESOLVED_EXPECTED_EXISTENCE_POLICY).value, vbNullString))), _
            Trim$(CStr(Nz(rs.fields(FIELD_RESOLVED_SETUP_ACTION).value, vbNullString))), _
            Trim$(CStr(Nz(rs.fields(FIELD_RIBBON_TOOLTIP).value, vbNullString))), _
            Trim$(CStr(Nz(rs.fields(FIELD_OBJECT_DESCRIPTION).value, vbNullString))), _
            Trim$(CStr(Nz(rs.fields(FIELD_RIBBON_TAB_KEY).value, vbNullString))), _
            Trim$(CStr(Nz(rs.fields(FIELD_RIBBON_TAB_LABEL).value, vbNullString))), _
            Trim$(CStr(Nz(rs.fields(FIELD_RIBBON_GROUP_LABEL).value, vbNullString))), _
            CLng(Nz(rs.fields(FIELD_RIBBON_GROUP_SORT_ORDER).value, 9999)), _
            CLng(rs.fields("EffectiveSortOrder").value), _
            TargetObjectName, primaryPlacementFlag, builtInSystemMacroFlag, activeCategoryFlag, _
            Trim$(CStr(Nz(rs.fields(FIELD_RIBBON_STATIC_TAG).value, vbNullString))), _
            CStr(Nz(rs.fields(FIELD_RIBBON_TITLE).value, vbNullString)), _
            Trim$(CStr(Nz(rs.fields(FIELD_CATALOG_OWNER).value, vbNullString))), _
            Trim$(CStr(Nz(rs.fields(FIELD_OBJECT_DISPLAY_NAME).value, vbNullString))))
        mRibbonCatalogSourceRows.Add rowData
        If activeCategoryFlag <> 0 Then mRibbonCatalogPreflightRowCount = mRibbonCatalogPreflightRowCount + 1
        If Len(objectName) > 0 And Len(objectTypeCode) > 0 Then
            exactKey = RibbonCatalogMetadataKey(objectName, objectTypeCode)
            If mRibbonCatalogMetadataCache.Exists(exactKey) Then
                metadata = mRibbonCatalogMetadataCache(exactKey)
                metadata(0) = CLng(metadata(0)) + 1
                metadata(1) = CLng(metadata(1)) + primaryPlacementFlag
                metadata(2) = CLng(metadata(2)) + builtInSystemMacroFlag
                If StrComp(CStr(metadata(5)), CStr(rowData(16)), vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 7283, "PrepareRibbonCatalogSourceSnapshot", "Duplicate active managed-object identities have different permanent Ribbon tags." & vbCrLf & "Object: " & objectName & vbCrLf & "Type: " & objectTypeCode
                mRibbonCatalogMetadataCache(exactKey) = metadata
            Else
                metadata = Array(1&, primaryPlacementFlag, builtInSystemMacroFlag, CStr(rowData(5)), CStr(rowData(6)), CStr(rowData(16)), CStr(rowData(17)), CStr(rowData(18)), CStr(rowData(19)))
                mRibbonCatalogMetadataCache.Add exactKey, metadata
            End If
            If Len(TargetObjectName) > 0 Then
                aliasKey = RibbonCatalogMetadataKey(TargetObjectName, objectTypeCode)
                If mRibbonCatalogAliasMetadataCache.Exists(aliasKey) = False Then
                    metadata = Array(0&, 0&, 0&, CStr(rowData(5)), CStr(rowData(6)), CStr(rowData(16)), CStr(rowData(17)), CStr(rowData(18)), CStr(rowData(19)))
                    mRibbonCatalogAliasMetadataCache.Add aliasKey, metadata
                End If
            End If
        End If
        rs.MoveNext
    Loop
    If loadedRowCount > 0 And loadedRowCount Mod 50 <> 0 Then RibbonRequireSafeWorkContinue "PrepareRibbonCatalogSourceSnapshot", "completed " & CStr(loadedRowCount) & " rows"
    rs.Close
    Set rs = Nothing
    LoadRibbonSystemMacroTargetSnapshot db
    mRibbonCatalogSnapshotElapsedSeconds = ElapsedSecondsFromTimer(startedTimer)
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
    Set mRibbonCatalogSourceRows = Nothing
    mRibbonCatalogPreflightRowCount = 0
    Set mRibbonSystemMacroTargetCache = Nothing
    Set mRibbonCatalogMetadataCache = Nothing
    Set mRibbonCatalogAliasMetadataCache = Nothing
    On Error GoTo 0
    Err.Raise savedErrNumber, "PrepareRibbonCatalogSourceSnapshot", "Could not load the shared read-only managed-object Ribbon source snapshot used by Part G." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Function RibbonCatalogValidationToken() As String
    ' Part G consumes the active complete-catalog exact token returned by the same current-state handoff; it does not consult alternate state.
    On Error GoTo ErrHandler
    Dim detailText As String
    RibbonCatalogValidationToken = CompleteCatalogStateTokenForSetup(detailText)
    If Len(RibbonCatalogValidationToken) = 0 Then Err.Raise vbObjectError + 7240, "RibbonCatalogValidationToken", detailText
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "RibbonCatalogValidationToken", "Could not resolve the exact active catalog token required for Part H Ribbon verification." & vbCrLf & Err.Description
End Function

Private Sub ValidateRibbonCatalogCoverage(ByVal rows As Collection)
    On Error GoTo ErrHandler
    Dim rowData As Variant
    Dim ribbonRowData As Variant
    Dim ribbonIdentities As Object
    Dim objectName As String
    Dim objectTypeCode As String
    Dim categoryCode As String
    Dim expectedExistencePolicy As String
    Dim setupAction As String
    Dim optionalReason As String
    Dim checkedCount As Long
    Dim optionalAbsentCount As Long
    Dim policyExcludedCount As Long
    ' INTENTIONAL KEYED LOOKUP: ribbonIdentities provides keyed membership/lookup over source rows or state already read by ValidateRibbonCatalogCoverage; it is a derived access path, not a second authority.
    Set ribbonIdentities = CreateObject("Scripting.Dictionary")
    ribbonIdentities.CompareMode = vbTextCompare
    For Each ribbonRowData In rows
        ribbonIdentities(RibbonCatalogMetadataKey(CStr(ribbonRowData(4)), CStr(ribbonRowData(3)))) = True
    Next ribbonRowData
    If mRibbonCatalogSourceRows Is Nothing Then PrepareRibbonCatalogSourceSnapshot
    PrintRibbonStatus "Validate complete managed-object ribbon coverage from the shared H1 source snapshot."
    For Each rowData In mRibbonCatalogSourceRows
        checkedCount = checkedCount + 1
        objectName = CStr(rowData(0))
        objectTypeCode = CStr(rowData(1))
        categoryCode = CStr(rowData(2))
        expectedExistencePolicy = CStr(rowData(3))
        setupAction = CStr(rowData(4))
        If checkedCount = 1 Or checkedCount Mod 50 = 0 Then
            PrintRibbonStatus "Validate shared ribbon catalog row " & CStr(checkedCount) & ": " & objectTypeCode & " [" & objectName & "]."
            RibbonRequireSafeWorkContinue "ValidateRibbonCatalogCoverage", objectTypeCode & " [" & objectName & "]"
        End If
        If ManagedObjectIsRibbonExcludedByPolicy(objectName, objectTypeCode) Then
            policyExcludedCount = policyExcludedCount + 1
        ElseIf ribbonIdentities.Exists(RibbonCatalogMetadataKey(objectName, objectTypeCode)) = False Then
            optionalReason = vbNullString
            If RibbonObjectExists(objectTypeCode, objectName) = False Then
                If Len(optionalReason) = 0 Then optionalReason = ManagedObjectOptionalMissingReason(expectedExistencePolicy, setupAction)
            End If
            If Len(optionalReason) > 0 Then
                optionalAbsentCount = optionalAbsentCount + 1
            Else
                Err.Raise vbObjectError + 6671, "ValidateRibbonCatalogCoverage", "A required cataloged managed object has no ribbon placement." & vbCrLf & "Object: " & objectName & vbCrLf & "Type: " & objectTypeCode & vbCrLf & "Category: " & categoryCode
            End If
        End If
    Next rowData
    If checkedCount > 0 And checkedCount Mod 50 <> 0 Then RibbonRequireSafeWorkContinue "ValidateRibbonCatalogCoverage", "completed " & CStr(checkedCount) & " rows"
    PrintRibbonStatus "Managed-object ribbon coverage validation completed from the shared source snapshot: rows=" & CStr(checkedCount) & "; optional uninstalled rows skipped=" & CStr(optionalAbsentCount) & "; explicit hidden-form exclusions=" & CStr(policyExcludedCount) & "."
    LogTraceMessage TRACE_LEVEL_INFO, "Ribbon", "ValidateRibbonCatalogCoverage", "Complete managed-object ribbon coverage verified from the shared H1 source snapshot.", "Catalog rows checked=" & CStr(checkedCount) & vbCrLf & "Generated ribbon identities indexed once=" & CStr(ribbonIdentities.count) & vbCrLf & "Optional uninstalled rows skipped=" & CStr(optionalAbsentCount) & vbCrLf & "Explicit hidden-form ribbon exclusions=" & CStr(policyExcludedCount) & vbCrLf & "Built-in system macros use their catalog-authoritative ADC1/ADC2 group assignments.", vbNullString, 0, vbNullString, vbNullString, vbNullString, "ManagedObjectCatalog", "CodeDb", "UI"
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ValidateRibbonCatalogCoverage", savedErrDescription
End Sub

Private Function RibbonManagedObjectPlacementDiagnosticText(ByVal db As DAO.Database) As String
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim resultText As String
    Dim objectID As Long
    Dim objectName As String
    Dim objectTypeCode As String
    Dim categoryCode As String
    Dim catalogOwner As String
    Dim defaultSource As String
    Dim assignedSettingKey As String
    Dim effectiveSettingKey As String
    Dim tabKey As String
    Dim tabLabel As String
    Dim groupLabel As String
    Dim groupSortOrder As Long
    Dim lastDefaultRepairText As String
    Dim settingValueText As String
    Dim navGroupText As String
    Dim physicalExistsText As String
    Dim rowNumber As Long
    Dim totalIncomplete As Long
    If SharedDao_TableFieldExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID), FIELD_RIBBON_TAB_KEY) = False Or SharedDao_TableFieldExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID), FIELD_RIBBON_TAB_LABEL) = False Or SharedDao_TableFieldExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID), FIELD_RIBBON_GROUP_LABEL) = False Or SharedDao_TableFieldExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID), FIELD_RIBBON_GROUP_SORT_ORDER) = False Then
        RibbonManagedObjectPlacementDiagnosticText = vbCrLf & "Managed-object placement fields are incomplete in the table schema; detailed placement diagnostics could not be queried."
        Exit Function
    End If
    sqlText = "SELECT Count(*) AS [ActiveCount], Sum(IIf(Len(Trim(Nz([" & FIELD_RIBBON_TAB_KEY & "],'')))=0 OR Len(Trim(Nz([" & FIELD_RIBBON_TAB_LABEL & "],'')))=0 OR Len(Trim(Nz([" & FIELD_RIBBON_GROUP_LABEL & "],'')))=0 OR Nz([" & FIELD_RIBBON_GROUP_SORT_ORDER & "],0)<=0,1,0)) AS [IncompleteCount] FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID) & "] WHERE Nz([" & FIELD_MANAGED_OBJECT_STATUS & "]," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & " AND UCase(Trim(Nz(" & SQL_FIELD_OBJECT_TYPE & ",''))) In ('TABLE','QUERY','FORM','REPORT','MACRO');"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    totalIncomplete = CLng(Nz(rs.fields("IncompleteCount").value, 0))
    resultText = resultText & vbCrLf & "Managed-object placement summary: Active=" & CStr(rs.fields("ActiveCount").value) & "; Incomplete=" & CStr(totalIncomplete)
    rs.Close
    Set rs = Nothing
    sqlText = "SELECT " & SQL_FIELD_OBJECT_TYPE & ", [" & FIELD_CATEGORY_CODE & "], Count(*) AS " & SQL_FIELD_ROW_COUNT & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID) & "] WHERE Nz([" & FIELD_MANAGED_OBJECT_STATUS & "]," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & " AND UCase(Trim(Nz(" & SQL_FIELD_OBJECT_TYPE & ",''))) In ('TABLE','QUERY','FORM','REPORT','MACRO') AND (Len(Trim(Nz([" & FIELD_RIBBON_TAB_KEY & "],'')))=0 OR Len(Trim(Nz([" & FIELD_RIBBON_TAB_LABEL & "],'')))=0 OR Len(Trim(Nz([" & FIELD_RIBBON_GROUP_LABEL & "],'')))=0 OR Nz([" & FIELD_RIBBON_GROUP_SORT_ORDER & "],0)<=0) GROUP BY " & SQL_FIELD_OBJECT_TYPE & ", [" & FIELD_CATEGORY_CODE & "] ORDER BY " & OrderByPolicyText("UI_RIBBON_PLACEMENT_DIAGNOSTIC_GROUP") & ";"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    If rs.EOF Then resultText = resultText & vbCrLf & "Incomplete placement groups: none."
    Do While rs.EOF = False
        resultText = resultText & vbCrLf & "Incomplete placement group: Type=" & CStr(rs.fields(FIELD_OBJECT_TYPE).value) & "; Category=" & CStr(Nz(rs.fields(FIELD_CATEGORY_CODE).value, vbNullString)) & "; Rows=" & CStr(rs.fields(FIELD_ROW_COUNT).value)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    If totalIncomplete = 0 Then
        RibbonManagedObjectPlacementDiagnosticText = resultText
        Exit Function
    End If
    sqlText = "SELECT TOP 25 " & SQL_FIELD_ID & "," & SQL_FIELD_OBJECT_NAME & "," & SQL_FIELD_OBJECT_TYPE & ",[" & FIELD_CATEGORY_CODE & "],[" & FIELD_CATALOG_OWNER & "]," & SQL_FIELD_DEFAULT_SOURCE & ",[" & FIELD_NAVIGATION_PANE_GROUP_SETTING_KEY & "],[" & FIELD_RIBBON_TAB_KEY & "],[" & FIELD_RIBBON_TAB_LABEL & "],[" & FIELD_RIBBON_GROUP_LABEL & "],[" & FIELD_RIBBON_GROUP_SORT_ORDER & "],[" & FIELD_LAST_DEFAULT_REPAIR_ON & "] FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID) & "] WHERE Nz([" & FIELD_MANAGED_OBJECT_STATUS & "]," & _
        SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & " AND UCase(Trim(Nz(" & SQL_FIELD_OBJECT_TYPE & ",''))) In ('TABLE','QUERY','FORM','REPORT','MACRO') AND (Len(Trim(Nz([" & FIELD_NAVIGATION_PANE_GROUP_SETTING_KEY & "],'')))=0 OR Len(Trim(Nz([" & FIELD_RIBBON_TAB_KEY & "],'')))=0 OR Len(Trim(Nz([" & FIELD_RIBBON_TAB_LABEL & "],'')))=0 OR Len(Trim(Nz([" & FIELD_RIBBON_GROUP_LABEL & "],'')))=0 OR Nz([" & FIELD_RIBBON_GROUP_SORT_ORDER & "],0)<=0) ORDER BY " & OrderByPolicyText("UI_RIBBON_PLACEMENT_DIAGNOSTIC_ROWS") & ";"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    Do While rs.EOF = False
        rowNumber = rowNumber + 1
        objectID = CLng(Nz(rs.fields(FIELD_ID).value, 0))
        objectName = CStr(Nz(rs.fields(FIELD_OBJECT_NAME).value, vbNullString))
        objectTypeCode = Trim$(CStr(rs.fields(FIELD_OBJECT_TYPE).value))
        categoryCode = Trim$(CStr(Nz(rs.fields(FIELD_CATEGORY_CODE).value, vbNullString)))
        catalogOwner = Trim$(CStr(Nz(rs.fields(FIELD_CATALOG_OWNER).value, vbNullString)))
        defaultSource = Trim$(CStr(Nz(rs.fields(FIELD_DEFAULT_SOURCE).value, vbNullString)))
        assignedSettingKey = Trim$(CStr(Nz(rs.fields(FIELD_NAVIGATION_PANE_GROUP_SETTING_KEY).value, vbNullString)))
        effectiveSettingKey = assignedSettingKey
        tabKey = Trim$(CStr(Nz(rs.fields(FIELD_RIBBON_TAB_KEY).value, vbNullString)))
        tabLabel = Trim$(CStr(Nz(rs.fields(FIELD_RIBBON_TAB_LABEL).value, vbNullString)))
        groupLabel = Trim$(CStr(Nz(rs.fields(FIELD_RIBBON_GROUP_LABEL).value, vbNullString)))
        groupSortOrder = CLng(Nz(rs.fields(FIELD_RIBBON_GROUP_SORT_ORDER).value, 0))
        lastDefaultRepairText = DiagnosticDateText(rs.fields(FIELD_LAST_DEFAULT_REPAIR_ON).value)
        settingValueText = RibbonDiagnosticSettingValue(db, effectiveSettingKey)
        navGroupText = RibbonDiagnosticNavigationGroupText(db, effectiveSettingKey)
        physicalExistsText = RibbonDiagnosticPhysicalObjectText(objectTypeCode, objectName)
        resultText = resultText & vbCrLf & "Incomplete placement " & CStr(rowNumber) & " of " & CStr(totalIncomplete) & ": ID=" & CStr(objectID) & "; Object=" & objectName & "; Type=" & objectTypeCode & "; Category=" & categoryCode & "; CatalogOwner=" & catalogOwner & "; DefaultSource=" & defaultSource & "; LastDefaultRepairOn=" & lastDefaultRepairText
        resultText = resultText & vbCrLf & "  Placement fields: TabKey=" & IIf(Len(tabKey) = 0, "(blank)", tabKey) & "; TabLabel=" & IIf(Len(tabLabel) = 0, "(blank)", tabLabel) & "; Group=" & IIf(Len(groupLabel) = 0, "(blank)", groupLabel) & "; GroupOrder=" & CStr(groupSortOrder)
        resultText = resultText & vbCrLf & "  Navigation assignment: SettingKey=" & IIf(Len(assignedSettingKey) = 0, "(blank)", assignedSettingKey) & "; SettingValue=" & settingValueText
        resultText = resultText & vbCrLf & "  Navigation group: " & navGroupText & "; Physical object: " & physicalExistsText
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    If totalIncomplete > rowNumber Then resultText = resultText & vbCrLf & "Incomplete placement rows omitted from detail: " & CStr(totalIncomplete - rowNumber)
    RibbonManagedObjectPlacementDiagnosticText = resultText
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
    RibbonManagedObjectPlacementDiagnosticText = resultText & vbCrLf & "Managed-object placement diagnostic collection failed. Error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription
End Function
