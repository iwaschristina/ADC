Attribute VB_Name = "x__UTIL_ADC_DB_REFERENCE"
' ================================================================================================================================================================================
' ChatGPT Rewrite Tracking Header
' ================================================================================================================================================================================
' ChatGPT Previous Version ID: CGPT-5HAB-GK64-9HC5-GLAE-L6Y4
' ChatGPT Baseline Version ID: CGPT-5HAB-GK64-9HC5-GLAE-L6Y4
' ChatGPT Version ID: CGPT-6LTV-S5NN-EDJS-5QHM-EFKT
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
' STANDARD MODULE: Access Development Console Database Reference
' VB_Name: x__UTIL_ADC_DB_REFERENCE
' Entry Points:
'     SynchronizeNonFormDatabaseReferenceFieldIndexForSetup
'     ValidateNonFormDatabaseReferenceFieldIndex
' SHARED CODE:
'     SharedDao_QueryExists(ByVal db As DAO.Database, ByVal queryName As String) -> Boolean
'     SharedDao_RecordsetHasFieldSafe(ByVal rs As DAO.Recordset, ByVal fieldName As String) -> Boolean
'     SharedDao_SqlQuote(ByVal value As Variant) -> String
'     SharedDao_TableExists(ByVal db As DAO.Database, ByVal tableName As String) -> Boolean
'     SharedText_AppendNonBlankLine(ByVal existingText As String, ByVal lineText As String) -> String
'     SharedVba_IsIdentifierChar(ByVal value As String) -> Boolean
' CLASSES USED: None
' DYNAMIC STRING REFERENCES: None identified in the extracted Database Reference subsystem.
' LastScanTimestamp: Wednesday 8/26/26 12:13 AM (split-time static review).
' Purpose:
'     Owns the Database Reference implementation extracted from Catalog: B5 synchronization of non-form table/query physical-field evidence and controlled values, plus read-only prerequisite validation of the
'     completed B5 evidence before C1-C3. Catalog remains the phase orchestrator, managed-object rollback-SQL authority, documentation-status vocabulary owner, and dirty-domain/proof-state owner.
' Author: Christina Shonts
' Support: Access Development Console Phase-B managed-object Database Reference synchronization and downstream prerequisite verification.
' Code Category: Access Development Console Database Reference authority.
' Requirements: Microsoft Access, DAO, the companion Access Development Console modules, and x__SHARED_CODE.
' Reference Dependencies:
'     x__UTIL_ADC_MOC for canonical documentation-status values and managed-object rollback SQL; Core/Common/Defaults/Tables and shared DAO/text/VBA helpers for existing synchronization behavior.
' DB Object Dependencies:
'     Active non-form managed-object catalog rows, ADC_ManagedObjectFields, ADC_ManagedObjectFieldValues, and the physical CodeDb table/query definitions represented by those rows.
' Usage:
'     Catalog calls the synchronization entry only from its B5 validation path and the validation entry only from its read-only C1-C3 prerequisite checks. Do not invoke synchronization from later phases.
' Version:
'     Moves the non-form reference simple-identifier helper into Database Reference ownership; reference behavior and B5 ownership are unchanged.
'     Routes Managed Object Catalog references to x__UTIL_ADC_MOC and physical lifecycle references to x__UTIL_ADC_MOC_LIFECYCLE after the responsibility split; runtime and Setup-phase behavior are unchanged.
'     Created by extracting the cohesive 33-procedure non-form Database Reference subsystem and its six instrumentation variables from x__UTIL_ADC_MOC without changing setup-phase ownership.
' Notes:
'     B5 Database Reference synchronization uses reusable field dynasets, in-memory identity maps, dependency-aware query evidence reuse, delta-only structural writes, controlled-value synchronization, and one
'     retirement pass; query fields capture direct source identity and inherit cached source-table metadata without enumerating expensive QueryDef.Field custom-property collections; final fail-closed validation
'     covers documentation and controlled-value integrity.
'     B5 reuses stored query field evidence only when every stored field structural fingerprint recomputes exactly, the rows reproduce the parent physical fingerprint, and the parent proof postdates every current
'     physical table/query definition; otherwise it falls back to direct QueryDef.Fields enumeration.
'     Synchronization reports whether inventory evidence changed; Catalog publishes that change into its private dirty-domain state after the call, including the failure path when synchronization already
'     reached its prior mutation-detection point.
' Core Rules Enforced Here:
'     B5 may synchronize table/query Database Reference evidence. The C1-C3 entry is validation-only. Form metadata remains outside this subsystem, and Health Review remains owned by its later diagnostics workflow.
' Design Map:
'     Public synchronization/validation entry points lead into private object, field, controlled-value, staging, physical-definition, source-metadata, and postcondition helpers kept together with their module state.
' Configuration: Existing managed-object definitions and physical Access table/query metadata remain authoritative.
' Behavior: Preserves the extracted B5 transaction, staging, rollback-registration, evidence-reuse, validation, tracing, and responsiveness behavior.
' Code Status: Active Database Reference authority.
' Todo: None.
' Reporting Bugs: None currently recorded in this header.
' Module Level Bugs: None currently recorded in this header.
' Public Subs:
'     SynchronizeNonFormDatabaseReferenceFieldIndexForSetup(ByVal db As DAO.Database, ByRef detailText As String, ByRef inventoryChanged As Boolean) -> None
'     ValidateNonFormDatabaseReferenceFieldIndex(ByVal db As DAO.Database) -> None
' Public Functions:
'     NonFormReferenceIsSimpleIdentifier(ByVal textValue As String) -> Boolean
' Public Properties: None
' ================================================================================================================================================================================

Option Compare Database
Option Explicit

' ===============================================================================================================================================================================
' MODULE-LEVEL DECLARATIONS
' ===============================================================================================================================================================================

' B5 Database Reference instrumentation retained with the synchronization implementation that owns and reports these counters.
Private mB5DatabaseReferenceControlledValueFieldCount As Long
Private mB5DatabaseReferenceLookupRowSourceOpenCount As Long
Private mB5DatabaseReferenceLookupRowSourceCacheHitCount As Long
Private mB5DatabaseReferenceQueryControlledValueSkipCount As Long
Private mB5DatabaseReferenceFormDesignMinimalPropertyCount As Long
Private mB5DatabaseReferenceQueryMinimalPropertyCount As Long

' ===============================================================================================================================================================================
' PROCEDURES
' ===============================================================================================================================================================================

' Purpose: Synchronizes and postcondition-validates the Phase-B non-form table/query Database Reference physical-field and controlled-value index immediately before B5 completes.
' Public role: Inter-module B5 implementation API called only by Catalog phase orchestration; it performs Database Reference work but does not own Catalog dirty-domain state.
' Parameters:
'   db (Input; DAO.Database): CodeDb database whose active non-form MAC rows and physical table/query definitions are synchronized.
'   detailText (Output; String): Receives synchronization, validation, staging, reuse, and timing detail for the owning Catalog phase report.
'   inventoryChanged (Output; Boolean): True only after synchronization determines that object, field, controlled-value, or retirement changes occurred; Catalog uses this to invalidate its inventory proof.
' Returns: None (Sub). Raises on synchronization, postcondition-validation, rollback, or cleanup failure.
Public Sub SynchronizeNonFormDatabaseReferenceFieldIndexForSetup(ByVal db As DAO.Database, ByRef detailText As String, ByRef inventoryChanged As Boolean)
    'Builds the Phase-B table/query Database Reference physical-field index immediately before B5 validates it.
    'This is intentionally non-form. It synchronizes structural field evidence plus controlled-value documentation; Health Review findings remain owned by the later diagnostics workflow.
    On Error GoTo ErrHandler
    Dim parentRs As DAO.Recordset
    Dim parentBookmarks As Object
    Dim fieldRs As DAO.Recordset
    Dim fieldBookmarksByObjectID As Object
    Dim objects As Collection
    Dim rowData As Variant
    Dim ws As DAO.workspace
    Dim transactionStarted As Boolean
    Dim objectCount As Long
    Dim synchronizedFieldCount As Long
    Dim inactiveFieldCount As Long
    Dim changedObjectCount As Long
    Dim changedFieldCount As Long
    Dim reusedQueryEvidenceCount As Long
    Dim synchronizedValueCount As Long
    Dim changedValueCount As Long
    Dim physicalUpdatedByName As Object
    Dim physicalTypeByName As Object
    Dim physicalNameByKey As Object
    Dim querySqlByName As Object
    Dim activeValueFieldIDs As Object
    Dim valueRs As DAO.Recordset
    Dim valuePrimaryBookmarkByIdentity As Object
    Dim valuePrimaryIDByIdentity As Object
    Dim valueBookmarksByFieldID As Object
    Dim indexedFieldRowCount As Long
    Dim indexedValueRowCount As Long
    Dim sourceFieldMetadataCache As Object
    Dim sourceTableDefinitionCache As Object
    Dim rowSourceResultCache As Object
    Dim rowSourceOpenFailureCache As Object
    Dim coldBulkMode As Boolean
    Dim nextColdFieldID As Long
    Dim coldStageTablesPrepared As Boolean
    Dim fieldStageTableName As String
    Dim valueStageTableName As String
    Dim coldBulkDetail As String
    Dim coldCleanupDetail As String
    Dim phaseTimer As Single
    Dim parentLoadTimer As Single
    Dim parentLoadDetail As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim validationDetail As String
    Dim synchronizationDetail As String
    inventoryChanged = False
    phaseTimer = Timer
    parentLoadTimer = Timer
    mB5DatabaseReferenceControlledValueFieldCount = 0
    mB5DatabaseReferenceLookupRowSourceOpenCount = 0
    mB5DatabaseReferenceLookupRowSourceCacheHitCount = 0
    mB5DatabaseReferenceQueryControlledValueSkipCount = 0
    mB5DatabaseReferenceFormDesignMinimalPropertyCount = 0
    mB5DatabaseReferenceQueryMinimalPropertyCount = 0
    ' INTENTIONAL FILTERED PARTITION: objects is a consumer-required subset/projection derived from the source enumerated by SynchronizeNonFormDatabaseReferenceFieldIndexForSetup; its membership is not independently
    ' authoritative.
    ' INTENTIONAL KEYED LOOKUP: parentBookmarks provides keyed membership/lookup over source rows or state already read by SynchronizeNonFormDatabaseReferenceFieldIndexForSetup; it is a derived access path, not a
    ' second authority.
    Set objects = New Collection
    Set parentBookmarks = CreateObject("Scripting.Dictionary")
    parentBookmarks.CompareMode = vbBinaryCompare
    Set parentRs = db.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID) & "] WHERE Nz([" & FIELD_MANAGED_OBJECT_STATUS & "]," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & " AND UCase(Trim(Nz(" & SQL_FIELD_OBJECT_TYPE & ",''))) IN ('TABLE','QUERY') ORDER BY " & SQL_FIELD_OBJECT_TYPE & "," & SQL_FIELD_OBJECT_NAME & "," & SQL_FIELD_ID & ";", dbOpenDynaset)
    Do While parentRs.EOF = False
        PumpSetupResponsiveness
        objects.Add Array(CLng(parentRs.fields(FIELD_ID).value), CStr(Nz(parentRs.fields(FIELD_OBJECT_NAME).value, vbNullString)), CStr(parentRs.fields(FIELD_OBJECT_TYPE).value))
        parentBookmarks.Add CStr(CLng(parentRs.fields(FIELD_ID).value)), parentRs.Bookmark
        parentRs.MoveNext
    Loop
    If objects.count = 0 Then Err.Raise vbObjectError + 9572, "SynchronizeNonFormDatabaseReferenceFieldIndexForSetup", "No active non-form table/query MAC rows were available for Database Reference synchronization."
    parentLoadDetail = "Database Reference parent MAC load: one shared dynaset; rows=" & CStr(objects.count) & "; elapsed=" & FormattedElapsedSeconds(ElapsedSecondsFromTimer(parentLoadTimer))
    'TABLES may have created or repaired physical objects through its retained Workspace database handle earlier in the same setup run. Refresh this retained Catalog handle once before the authoritative B5
    'physical-definition snapshot so named DAO lookups cannot use stale TableDefs/QueryDefs collections.
    db.TableDefs.Refresh
    db.QueryDefs.Refresh
    NonFormReferenceBuildPhysicalDefinitionMaps db, physicalUpdatedByName, physicalTypeByName, physicalNameByKey, querySqlByName
    ' INTENTIONAL KEYED LOOKUP: sourceFieldMetadataCache, sourceTableDefinitionCache, rowSourceResultCache provides keyed membership/lookup over source rows or state already read by
    ' SynchronizeNonFormDatabaseReferenceFieldIndexForSetup; it is a derived access path, not a second authority.
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: rowSourceOpenFailureCache holds planned/recovery work or independently observed validation evidence for SynchronizeNonFormDatabaseReferenceFieldIndexForSetup; keep
    ' it distinct from the source authority.
    Set sourceFieldMetadataCache = CreateObject("Scripting.Dictionary")
    sourceFieldMetadataCache.CompareMode = vbTextCompare
    Set sourceTableDefinitionCache = CreateObject("Scripting.Dictionary")
    sourceTableDefinitionCache.CompareMode = vbTextCompare
    Set rowSourceResultCache = CreateObject("Scripting.Dictionary")
    rowSourceResultCache.CompareMode = vbTextCompare
    Set rowSourceOpenFailureCache = CreateObject("Scripting.Dictionary")
    rowSourceOpenFailureCache.CompareMode = vbTextCompare
    coldBulkMode = (NonFormReferenceTableRowCount(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFields")) = 0 And NonFormReferenceTableRowCount(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFieldValues")) = 0)
    If coldBulkMode Then
        NonFormReferencePrepareColdStageTables db, fieldStageTableName, valueStageTableName
        nextColdFieldID = 1
        coldStageTablesPrepared = True
    End If
    Set ws = DBEngine.Workspaces(0)
    'Cold mode writes only disposable stage tables during physical discovery. Keep DAO QueryDef/RowSource reads outside a Workspace transaction; transact only the final canonical append and validation.
    If coldBulkMode = False Then
        ws.BeginTrans
        transactionStarted = True
    End If
    If coldBulkMode Then
        NonFormReferenceOpenColdStageIndexes db, fieldStageTableName, valueStageTableName, fieldRs, fieldBookmarksByObjectID, valueRs, activeValueFieldIDs, valuePrimaryBookmarkByIdentity, valuePrimaryIDByIdentity, valueBookmarksByFieldID
    Else
        NonFormReferenceOpenSharedFieldIndex db, fieldRs, fieldBookmarksByObjectID, indexedFieldRowCount
        NonFormReferenceOpenSharedFieldValueIndex db, valueRs, activeValueFieldIDs, valuePrimaryBookmarkByIdentity, valuePrimaryIDByIdentity, valueBookmarksByFieldID, indexedValueRowCount
    End If
    For Each rowData In objects
        PumpSetupResponsiveness
        objectCount = objectCount + 1
        SynchronizeNonFormDatabaseReferenceObjectForSetup db, parentRs, parentBookmarks, fieldRs, fieldBookmarksByObjectID, CLng(rowData(0)), CStr(rowData(1)), CStr(rowData(2)), physicalUpdatedByName, physicalTypeByName, physicalNameByKey, querySqlByName, valueRs, activeValueFieldIDs, valuePrimaryBookmarkByIdentity, valuePrimaryIDByIdentity, valueBookmarksByFieldID, sourceFieldMetadataCache, sourceTableDefinitionCache, rowSourceResultCache, rowSourceOpenFailureCache, coldBulkMode, nextColdFieldID, synchronizedFieldCount, synchronizedValueCount, inactiveFieldCount, changedObjectCount, changedFieldCount, changedValueCount, reusedQueryEvidenceCount
        PulseProgressEvery objectCount, 10, "B5 Database Reference physical schema index " & CStr(objectCount) & " of " & CStr(objects.count) & ": [" & CStr(rowData(1)) & "]"
    Next rowData
    fieldRs.Close
    Set fieldRs = Nothing
    valueRs.Close
    Set valueRs = Nothing
    If coldBulkMode Then
        ws.BeginTrans
        transactionStarted = True
        NonFormReferenceAppendColdStageRows db, fieldStageTableName, valueStageTableName, coldBulkDetail
    End If
    ValidateNonFormDatabaseReferenceFieldIndexFromSnapshot db, parentRs, validationDetail
    ws.CommitTrans
    transactionStarted = False
    parentRs.Close
    Set parentRs = Nothing
    If coldStageTablesPrepared Then NonFormReferenceCleanupColdStageTables db, fieldStageTableName, valueStageTableName, coldCleanupDetail
    If Len(coldCleanupDetail) > 0 Then
        AddTraceNote coldCleanupDetail
        Err.Raise vbObjectError + 9614, "SynchronizeNonFormDatabaseReferenceFieldIndexForSetup", coldCleanupDetail
    End If
    inventoryChanged = (changedObjectCount > 0 Or changedFieldCount > 0 Or changedValueCount > 0 Or inactiveFieldCount > 0)
    detailText = SharedText_AppendNonBlankLine(detailText, parentLoadDetail)
    detailText = SharedText_AppendNonBlankLine(detailText, validationDetail)
    detailText = SharedText_AppendNonBlankLine(detailText, coldBulkDetail)
    synchronizationDetail = "Database Reference synchronization: mode=" & IIf(coldBulkMode, "cold-two-stage-unindexed-stage-bulk-append", "incremental-one-pass-dynaset-delta") & "; objects=" & CStr(objectCount) & "; fields=" & CStr(synchronizedFieldCount) & "; controlled values=" & CStr(synchronizedValueCount) & _
        "; query field evidence reused=" & CStr(reusedQueryEvidenceCount) & "; object rows changed=" & CStr(changedObjectCount) & "; field rows changed=" & CStr(changedFieldCount) & "; field-value rows changed=" & CStr(changedValueCount) & "; inactive field rows=" & CStr(inactiveFieldCount) & _
        "; MAC recordset opens=1; MAC_Fields working recordset opens=1; preexisting field rows indexed=" & CStr(indexedFieldRowCount) & "; MAC_Fields_Vals working recordset opens=1; preexisting field-value rows indexed=" & CStr(indexedValueRowCount) & "; per-object field recordset opens=0; per-field value recordset opens=0" & _
        "; controlled-value fields synchronized=" & CStr(mB5DatabaseReferenceControlledValueFieldCount) & "; query-field controlled-value enumerations skipped=" & CStr(mB5DatabaseReferenceQueryControlledValueSkipCount) & "; lookup RowSource recordset opens=" & CStr(mB5DatabaseReferenceLookupRowSourceOpenCount) & "; RowSource cache hits=" & CStr(mB5DatabaseReferenceLookupRowSourceCacheHitCount) & "; unique cached RowSources=" & CStr(rowSourceResultCache.count) & "; unavailable RowSources cached=" & CStr(rowSourceOpenFailureCache.count) & "; query fields using direct-source minimal metadata=" & CStr(mB5DatabaseReferenceQueryMinimalPropertyCount) & "; unique cached source-field metadata rows=" & CStr(sourceFieldMetadataCache.count) & "; FormDesign query fields=" & CStr(mB5DatabaseReferenceFormDesignMinimalPropertyCount) & "; elapsed=" & FormattedElapsedSeconds(ElapsedSecondsFromTimer(phaseTimer))
    detailText = SharedText_AppendNonBlankLine(detailText, synchronizationDetail)
    LogTraceMessage TRACE_LEVEL_INFO, "ManagedObjectCatalog", "SynchronizeNonFormDatabaseReferenceFieldIndexForSetup", "Non-form Database Reference physical schema index synchronized and verified.", detailText, "Database Reference schema index", 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFields"), "CatalogValidationNonForm", "CodeDb", "ManagedObjectCatalog"
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not fieldRs Is Nothing Then fieldRs.Close
    If Not valueRs Is Nothing Then valueRs.Close
    If Not parentRs Is Nothing Then parentRs.Close
    If transactionStarted Then ws.Rollback
    On Error GoTo 0
    If coldStageTablesPrepared Then NonFormReferenceCleanupColdStageTables db, fieldStageTableName, valueStageTableName, coldCleanupDetail
    Err.Raise savedErrNumber, "SynchronizeNonFormDatabaseReferenceFieldIndexForSetup", "Could not synchronize the Phase-B non-form Database Reference physical schema index." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription & IIf(Len(coldCleanupDetail) > 0, vbCrLf & coldCleanupDetail, vbNullString)
End Sub

Private Sub SynchronizeNonFormDatabaseReferenceObjectForSetup( _
    ByVal db As DAO.Database, ByVal parentRs As DAO.Recordset, ByVal parentBookmarks As Object, ByVal fieldRs As DAO.Recordset, ByVal fieldBookmarksByObjectID As Object, _
    ByVal managedObjectID As Long, ByVal objectName As String, ByVal objectType As String, ByVal physicalUpdatedByName As Object, ByVal physicalTypeByName As Object, ByVal physicalNameByKey As Object, ByVal querySqlByName As Object, _
    ByVal valueRs As DAO.Recordset, ByVal activeValueFieldIDs As Object, ByVal valuePrimaryBookmarkByIdentity As Object, ByVal valuePrimaryIDByIdentity As Object, ByVal valueBookmarksByFieldID As Object, _
    ByVal sourceFieldMetadataCache As Object, ByVal sourceTableDefinitionCache As Object, ByVal rowSourceResultCache As Object, ByVal rowSourceOpenFailureCache As Object, _
    ByVal coldBulkMode As Boolean, ByRef nextColdFieldID As Long, ByRef synchronizedFieldCount As Long, ByRef synchronizedValueCount As Long, ByRef inactiveFieldCount As Long, ByRef changedObjectCount As Long, ByRef changedFieldCount As Long, ByRef changedValueCount As Long, ByRef reusedQueryEvidenceCount As Long)
    On Error GoTo ErrHandler
    Dim tdf As DAO.tableDef
    Dim qdf As DAO.queryDef
    Dim fld As DAO.Field
    Dim selectedBookmarks As Object
    Dim selectedIDs As Object
    Dim objectFieldBookmarks As Object
    Dim seenPhysicalFields As Object
    Dim indexFlagsByField As Object
    Dim structuralSignature As String
    Dim objectFingerprint As String
    Dim parentChanged As Boolean
    Dim fieldDiscoveryErrNumber As Long
    Dim fieldDiscoveryErrDescription As String
    Dim existingFieldName As String
    Dim existingFieldID As Long
    Dim existingSelectedID As Long
    Dim rowBookmarkKey As Variant
    Dim objectFieldCount As Long
    Dim objectTimer As Single
    Dim objectElapsed As Double
    Dim queryEvidenceReused As Boolean
    Dim queryFieldExposureTimer As Single
    Dim queryFieldExposureElapsed As Double
    Dim queryFieldCollectionCount As Long
    Dim timingDetail As String
    Dim fieldMissingDescription As Boolean
    Dim categoryCode As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    objectTimer = Timer
    objectType = UCase$(Trim$(objectType))
    If managedObjectID <= 0 Or Len(objectName) = 0 Then Err.Raise vbObjectError + 9573, "SynchronizeNonFormDatabaseReferenceObjectForSetup", "A Database Reference object has an invalid managed identity."
    If fieldRs Is Nothing Or fieldBookmarksByObjectID Is Nothing Then Err.Raise vbObjectError + 9577, "SynchronizeNonFormDatabaseReferenceObjectForSetup", "The shared MAC_Fields synchronization index is unavailable."
    ' INTENTIONAL KEYED LOOKUP: selectedBookmarks, selectedIDs, objectFieldBookmarks provides keyed membership/lookup over source rows or state already read by SynchronizeNonFormDatabaseReferenceObjectForSetup; it
    ' is a derived access path, not a second authority.
    Set selectedBookmarks = CreateObject("Scripting.Dictionary")
    selectedBookmarks.CompareMode = vbTextCompare
    Set selectedIDs = CreateObject("Scripting.Dictionary")
    selectedIDs.CompareMode = vbTextCompare
    If fieldBookmarksByObjectID.Exists(CStr(managedObjectID)) Then
        Set objectFieldBookmarks = fieldBookmarksByObjectID(CStr(managedObjectID))
    Else
        Set objectFieldBookmarks = CreateObject("Scripting.Dictionary")
        objectFieldBookmarks.CompareMode = vbBinaryCompare
        fieldBookmarksByObjectID.Add CStr(managedObjectID), objectFieldBookmarks
    End If
    For Each rowBookmarkKey In objectFieldBookmarks.keys
        fieldRs.Bookmark = objectFieldBookmarks(CStr(rowBookmarkKey))
        existingFieldName = CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_NAME).value, vbNullString))
        existingFieldID = CLng(Nz(fieldRs.fields(FIELD_ID).value, 0))
        If Len(existingFieldName) > 0 And existingFieldID > 0 Then
            If selectedIDs.Exists(existingFieldName) Then
                existingSelectedID = CLng(selectedIDs(existingFieldName))
                If existingFieldID < existingSelectedID Then
                    selectedIDs(existingFieldName) = existingFieldID
                    selectedBookmarks(existingFieldName) = fieldRs.Bookmark
                End If
            Else
                selectedIDs.Add existingFieldName, existingFieldID
                selectedBookmarks.Add existingFieldName, fieldRs.Bookmark
            End If
        End If
    Next rowBookmarkKey
    ' INTENTIONAL KEYED LOOKUP: seenPhysicalFields, indexFlagsByField provides keyed membership/lookup over source rows or state already read by SynchronizeNonFormDatabaseReferenceObjectForSetup; it is a derived
    ' access path, not a second authority.
    Set seenPhysicalFields = CreateObject("Scripting.Dictionary")
    seenPhysicalFields.CompareMode = vbTextCompare
    Set indexFlagsByField = CreateObject("Scripting.Dictionary")
    indexFlagsByField.CompareMode = vbTextCompare
    If parentRs Is Nothing Or parentBookmarks Is Nothing Then Err.Raise vbObjectError + 9577, "SynchronizeNonFormDatabaseReferenceObjectForSetup", "The shared MAC synchronization snapshot is unavailable."
    If parentBookmarks.Exists(CStr(managedObjectID)) = False Then Err.Raise vbObjectError + 9577, "SynchronizeNonFormDatabaseReferenceObjectForSetup", "The managed-object row disappeared from the shared Database Reference synchronization snapshot." & vbCrLf & FIELD_MANAGED_OBJECT_ID & ": " & CStr(managedObjectID)
    parentRs.Bookmark = parentBookmarks(CStr(managedObjectID))
    If CLng(Nz(parentRs.fields(FIELD_ID).value, 0)) <> managedObjectID Then Err.Raise vbObjectError + 9577, "SynchronizeNonFormDatabaseReferenceObjectForSetup", "The shared MAC bookmark no longer identifies the expected row." & vbCrLf & "ManagedObjectID: " & CStr(managedObjectID)
    categoryCode = CStr(Nz(parentRs.fields(FIELD_CATEGORY_CODE).value, vbNullString))
    Select Case objectType
        Case "TABLE"
            On Error Resume Next
            Set tdf = db.TableDefs(objectName)
            fieldDiscoveryErrNumber = Err.Number
            fieldDiscoveryErrDescription = Err.Description
            On Error GoTo ErrHandler
            If fieldDiscoveryErrNumber <> 0 Or tdf Is Nothing Then Err.Raise vbObjectError + 9574, "SynchronizeNonFormDatabaseReferenceObjectForSetup", "The active managed table does not exist physically." & vbCrLf & "Table: " & objectName & IIf(fieldDiscoveryErrNumber = 0, vbNullString, vbCrLf & "DAO error: " & CStr(fieldDiscoveryErrNumber) & " - " & fieldDiscoveryErrDescription)
            If sourceTableDefinitionCache.Exists(UCase$(objectName)) = False Then sourceTableDefinitionCache.Add UCase$(objectName), tdf
            Set indexFlagsByField = NonFormReferenceBuildIndexFlagMap(tdf)
            structuralSignature = "TABLE|" & objectName
            For Each fld In tdf.fields
                PumpSetupResponsiveness
                structuralSignature = structuralSignature & "|" & fld.Name & ":" & CStr(fld.Type) & ":" & CStr(fld.Size)
                SynchronizeNonFormDatabaseReferenceFieldForSetup db, fieldRs, objectFieldBookmarks, selectedBookmarks, selectedIDs, seenPhysicalFields, indexFlagsByField, valueRs, activeValueFieldIDs, valuePrimaryBookmarkByIdentity, valuePrimaryIDByIdentity, valueBookmarksByFieldID, sourceFieldMetadataCache, sourceTableDefinitionCache, rowSourceResultCache, rowSourceOpenFailureCache, coldBulkMode, nextColdFieldID, managedObjectID, objectName, categoryCode, fld, True, synchronizedValueCount, changedFieldCount, changedValueCount, fieldMissingDescription
                synchronizedFieldCount = synchronizedFieldCount + 1
                objectFieldCount = objectFieldCount + 1
            Next fld
        Case "QUERY"
            On Error Resume Next
            Set qdf = db.QueryDefs(objectName)
            fieldDiscoveryErrNumber = Err.Number
            fieldDiscoveryErrDescription = Err.Description
            On Error GoTo ErrHandler
            If fieldDiscoveryErrNumber <> 0 Or qdf Is Nothing Then Err.Raise vbObjectError + 9575, "SynchronizeNonFormDatabaseReferenceObjectForSetup", "The active managed query could not expose its current physical definition." & vbCrLf & "Query: " & objectName & IIf(fieldDiscoveryErrNumber = 0, vbNullString, vbCrLf & "DAO error: " & CStr(fieldDiscoveryErrNumber) & " - " & fieldDiscoveryErrDescription)
            queryEvidenceReused = NonFormReferenceCanReuseQueryFieldEvidence(qdf, parentRs, fieldRs, objectFieldBookmarks, physicalUpdatedByName, physicalTypeByName, physicalNameByKey, querySqlByName, objectName, objectFieldCount, fieldMissingDescription)
            If queryEvidenceReused Then
                synchronizedFieldCount = synchronizedFieldCount + objectFieldCount
                reusedQueryEvidenceCount = reusedQueryEvidenceCount + 1
                SynchronizeNonFormDatabaseReferenceStoredFieldValuesForSetup db, fieldRs, objectFieldBookmarks, selectedIDs, valueRs, activeValueFieldIDs, valuePrimaryBookmarkByIdentity, valuePrimaryIDByIdentity, valueBookmarksByFieldID, rowSourceResultCache, rowSourceOpenFailureCache, synchronizedValueCount, changedValueCount
            Else
                fieldMissingDescription = False
                structuralSignature = "QUERY|" & objectName & "|Type=" & CStr(qdf.Type)
                queryFieldExposureTimer = Timer
                queryFieldCollectionCount = qdf.fields.count
                queryFieldExposureElapsed = ElapsedSecondsFromTimer(queryFieldExposureTimer)
                For Each fld In qdf.fields
                    PumpSetupResponsiveness
                    structuralSignature = structuralSignature & "|" & fld.Name & ":" & CStr(fld.Type) & ":" & CStr(fld.Size)
                    SynchronizeNonFormDatabaseReferenceFieldForSetup db, fieldRs, objectFieldBookmarks, selectedBookmarks, selectedIDs, seenPhysicalFields, indexFlagsByField, valueRs, activeValueFieldIDs, valuePrimaryBookmarkByIdentity, valuePrimaryIDByIdentity, valueBookmarksByFieldID, sourceFieldMetadataCache, sourceTableDefinitionCache, rowSourceResultCache, rowSourceOpenFailureCache, coldBulkMode, nextColdFieldID, managedObjectID, objectName, categoryCode, fld, False, synchronizedValueCount, changedFieldCount, changedValueCount, fieldMissingDescription
                    synchronizedFieldCount = synchronizedFieldCount + 1
                    objectFieldCount = objectFieldCount + 1
                Next fld
            End If
        Case Else
            Err.Raise vbObjectError + 9576, "SynchronizeNonFormDatabaseReferenceObjectForSetup", "Only TABLE and QUERY rows belong to the Phase-B Database Reference field-index synchronization." & vbCrLf & "Object: " & objectName & vbCrLf & "Type: " & objectType
    End Select
    If queryEvidenceReused = False Then
        If coldBulkMode = False Then MarkNonFormDatabaseReferenceFieldRowsInactiveForSetup db, fieldRs, objectFieldBookmarks, selectedIDs, seenPhysicalFields, activeValueFieldIDs, inactiveFieldCount, changedFieldCount, changedValueCount
        objectFingerprint = FingerprintFromPayload(structuralSignature)
        parentChanged = (StrComp(CStr(Nz(parentRs.fields(FIELD_MANAGED_OBJECT_FIELD_SCHEMA_FINGERPRINT).value, vbNullString)), objectFingerprint, vbBinaryCompare) <> 0 Or IsNull(parentRs.fields(FIELD_MANAGED_OBJECT_FIELD_SCHEMA_VERIFIED_ON).value))
        If parentChanged Then
            If SetupRollbackTrackingActive() Then TrackSetupManagedObjectRowRollback managedObjectID, True, ManagedObjectRollbackSql(parentRs)
            parentRs.Edit
            parentRs.fields(FIELD_MANAGED_OBJECT_FIELD_SCHEMA_FINGERPRINT).value = Left$(objectFingerprint, 255)
            parentRs.fields(FIELD_MANAGED_OBJECT_FIELD_SCHEMA_VERIFIED_ON).value = Now()
            If SharedDao_RecordsetHasFieldSafe(parentRs, FIELD_IS_DOCUMENTATION_ACTIVE) Then parentRs.fields(FIELD_IS_DOCUMENTATION_ACTIVE).value = True
            parentRs.Update
            changedObjectCount = changedObjectCount + 1
        End If
    End If
    NonFormReferenceUpdateObjectDocumentationStatusForSetup parentRs, fieldMissingDescription, parentChanged, changedObjectCount
    objectElapsed = ElapsedSecondsFromTimer(objectTimer)
    If objectElapsed >= 1# Then
        timingDetail = "Object=" & objectName & "; type=" & objectType & "; fields=" & CStr(objectFieldCount) & "; elapsed=" & FormattedElapsedSeconds(objectElapsed)
        If objectType = "QUERY" And queryEvidenceReused = False Then timingDetail = timingDetail & "; QueryDef.Fields exposure=" & FormattedElapsedSeconds(queryFieldExposureElapsed) & "; exposed field count=" & CStr(queryFieldCollectionCount)
        LogTraceMessage TRACE_LEVEL_INFO, "SetupTiming", "SynchronizeNonFormDatabaseReferenceObjectForSetup", "B5 Database Reference physical object timing.", timingDetail, "B5 Database Reference slow physical object", 0, vbNullString, vbNullString, objectName, "CatalogValidationNonForm", "CodeDb", "Performance"
    End If
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "SynchronizeNonFormDatabaseReferenceObjectForSetup", "Database Reference synchronization failed for a non-form managed object." & vbCrLf & "Object: " & objectName & vbCrLf & "Type: " & objectType & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Function NonFormDatabaseReferenceFieldRowIsCurrent( _
    ByVal fieldRs As DAO.Recordset, ByVal fld As DAO.Field, ByVal displayName As String, ByVal sourceObjectName As String, ByVal sourceFieldName As String, _
    ByVal dataTypeName As String, ByVal isRequired As Boolean, ByVal allowsZeroLength As Boolean, ByVal isPrimaryKey As Boolean, ByVal isIndexed As Boolean, _
    ByVal isUnique As Boolean, ByVal DefaultValue As String, ByVal validationRule As String, ByVal validationText As String, ByVal rowSourceType As String, _
    ByVal rowSource As String, ByVal lookupBoundColumn As Long, ByVal lookupColumnCount As Long, ByVal columnWidths As String, ByVal defaultDescriptionShortText As String, _
    ByVal defaultDescriptionLongText As String, ByVal structuralFingerprint As String, ByVal requiredFieldSizeDefault As Variant, ByRef existingDescriptionShort As String, ByRef existingDescriptionLong As String, _
    ByRef existingUsageGuidance As String, ByRef documentationSource As String, ByRef documentationStatus As String, ByRef schemaEvidenceChanged As Boolean) As Boolean
    'Compares one existing Database Reference field row with the already-derived B5 physical-field state; the caller retains row ownership, persistence, and child-value synchronization.
    Dim rowChanged As Boolean
    existingDescriptionShort = CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_DESCRIPTION_SHORT).value, vbNullString))
    existingDescriptionLong = CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_DESCRIPTION_LONG).value, vbNullString))
    existingUsageGuidance = CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_USAGE_GUIDANCE).value, vbNullString))
    documentationSource = CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_DOCUMENTATION_SOURCE).value, vbNullString))
    documentationStatus = CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_DOCUMENTATION_STATUS).value, vbNullString))
    If CLng(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_ORDINAL).value, -1)) <> fld.ordinalPosition Then rowChanged = True
    If StrComp(CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME).value, vbNullString)), displayName, vbBinaryCompare) <> 0 Then rowChanged = True
    If StrComp(CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_SOURCE_OBJECT_NAME).value, vbNullString)), sourceObjectName, vbBinaryCompare) <> 0 Then rowChanged = True
    If StrComp(CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_SOURCE_FIELD_NAME).value, vbNullString)), sourceFieldName, vbBinaryCompare) <> 0 Then rowChanged = True
    If StrComp(CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_DATA_TYPE).value, vbNullString)), dataTypeName, vbBinaryCompare) <> 0 Then rowChanged = True
    If CLng(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_ACTUAL_SIZE).value, -1)) <> fld.Size Then rowChanged = True
    If IsNull(requiredFieldSizeDefault) = False Then If IsNull(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_REQUIRED_SIZE).value) Then rowChanged = True
    If CBool(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_IS_REQUIRED).value, False)) <> isRequired Then rowChanged = True
    If CBool(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_ALLOWS_ZERO_LENGTH).value, False)) <> allowsZeroLength Then rowChanged = True
    If CBool(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_IS_PRIMARY_KEY).value, False)) <> isPrimaryKey Then rowChanged = True
    If CBool(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_IS_INDEXED).value, False)) <> isIndexed Then rowChanged = True
    If CBool(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_IS_UNIQUE).value, False)) <> isUnique Then rowChanged = True
    If StrComp(CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_DEFAULT_VALUE).value, vbNullString)), DefaultValue, vbBinaryCompare) <> 0 Then rowChanged = True
    If StrComp(CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALIDATION_RULE).value, vbNullString)), validationRule, vbBinaryCompare) <> 0 Then rowChanged = True
    If StrComp(CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALIDATION_TEXT).value, vbNullString)), validationText, vbBinaryCompare) <> 0 Then rowChanged = True
    If StrComp(CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_LOOKUP_ROW_SOURCE_TYPE).value, vbNullString)), rowSourceType, vbBinaryCompare) <> 0 Then rowChanged = True
    If StrComp(CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_LOOKUP_ROW_SOURCE).value, vbNullString)), rowSource, vbBinaryCompare) <> 0 Then rowChanged = True
    If CLng(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_LOOKUP_BOUND_COLUMN).value, 0)) <> lookupBoundColumn Then rowChanged = True
    If CLng(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_LOOKUP_COLUMN_COUNT).value, 0)) <> lookupColumnCount Then rowChanged = True
    If StrComp(CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_LOOKUP_COLUMN_WIDTHS).value, vbNullString)), columnWidths, vbBinaryCompare) <> 0 Then rowChanged = True
    If Len(existingDescriptionShort) = 0 And Len(defaultDescriptionShortText) > 0 Then rowChanged = True
    If Len(existingDescriptionLong) = 0 And Len(defaultDescriptionLongText) > 0 Then rowChanged = True
    If Len(existingUsageGuidance) = 0 Then rowChanged = True
    If Len(CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_PRESERVATION_POLICY).value, vbNullString))) = 0 Then rowChanged = True
    If Len(CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_SEARCH_KEYWORDS).value, vbNullString))) = 0 Then rowChanged = True
    If Len(documentationSource) = 0 Then
        rowChanged = True
    ElseIf StrComp(documentationSource, "ADC B5 physical schema index", vbTextCompare) = 0 Then
        rowChanged = True
    End If
    If Len(existingDescriptionLong) > 0 Or Len(defaultDescriptionLongText) > 0 Then
        If StrComp(documentationStatus, DOCUMENTATION_STATUS_COMPLETE, vbTextCompare) <> 0 Then rowChanged = True
    ElseIf StrComp(documentationStatus, DOCUMENTATION_STATUS_MISSING_DESCRIPTION, vbTextCompare) <> 0 Then
        rowChanged = True
    End If
    schemaEvidenceChanged = (StrComp(CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_SCHEMA_FINGERPRINT).value, vbNullString)), structuralFingerprint, vbBinaryCompare) <> 0 Or IsNull(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_SCHEMA_VERIFIED_ON).value))
    If schemaEvidenceChanged Then rowChanged = True
    If CBool(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_IS_ACTIVE).value, False)) = False Then rowChanged = True
    If Not IsNull(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_RETIRED_ON).value) Then rowChanged = True
    If Len(CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_RETIREMENT_REASON).value, vbNullString))) > 0 Then rowChanged = True
    NonFormDatabaseReferenceFieldRowIsCurrent = Not rowChanged
End Function

Private Sub SynchronizeNonFormDatabaseReferenceFieldForSetup(ByVal db As DAO.Database, ByVal fieldRs As DAO.Recordset, ByVal objectFieldBookmarks As Object, ByVal selectedBookmarks As Object, ByVal selectedIDs As Object, ByVal seenPhysicalFields As Object, ByVal indexFlagsByField As Object, ByVal valueRs As DAO.Recordset, ByVal activeValueFieldIDs As Object, ByVal valuePrimaryBookmarkByIdentity As Object, ByVal valuePrimaryIDByIdentity As Object, ByVal valueBookmarksByFieldID As Object, ByVal sourceFieldMetadataCache As Object, ByVal sourceTableDefinitionCache As Object, ByVal rowSourceResultCache As Object, ByVal rowSourceOpenFailureCache As Object, ByVal coldBulkMode As Boolean, ByRef nextColdFieldID As Long, ByVal managedObjectID As Long, ByVal objectName As String, ByVal categoryCode As String, ByVal fld As DAO.Field, ByVal isTableField As Boolean, ByRef synchronizedValueCount As Long, ByRef changedFieldCount As Long, ByRef changedValueCount As Long, ByRef fieldMissingDescription As Boolean)
    On Error GoTo ErrHandler
    Dim selectedID As Long
    Dim isNew As Boolean
    Dim rowChanged As Boolean
    Dim schemaEvidenceChanged As Boolean
    Dim displayName As String
    Dim sourceObjectName As String
    Dim sourceFieldName As String
    Dim dataTypeName As String
    Dim rowSourceType As String
    Dim rowSource As String
    Dim DefaultValue As String
    Dim validationRule As String
    Dim validationText As String
    Dim columnWidths As String
    Dim structuralText As String
    Dim structuralFingerprint As String
    Dim isRequired As Boolean
    Dim allowsZeroLength As Boolean
    Dim isPrimaryKey As Boolean
    Dim isIndexed As Boolean
    Dim isUnique As Boolean
    Dim lookupBoundColumn As Long
    Dim lookupColumnCount As Long
    Dim descriptionText As String
    Dim defaultDescriptionShortText As String
    Dim defaultDescriptionLongText As String
    Dim defaultDescriptionShortSource As String
    Dim defaultDescriptionLongSource As String
    Dim defaultDescriptionResolutionSource As Variant
    Dim existingDescriptionShort As String
    Dim existingDescriptionLong As String
    Dim existingUsageGuidance As String
    Dim documentationSource As String
    Dim documentationStatus As String
    Dim valuePreservationPolicy As String
    Dim searchKeywords As String
    Dim desiredDocumentationSource As String
    Dim requiredFieldSizeDefault As Variant
    Dim indexFlags As Long
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    NonFormReferenceReadFieldProperties fld, isTableField, categoryCode, descriptionText, displayName, rowSourceType, rowSource, sourceObjectName, sourceFieldName, isRequired, allowsZeroLength, DefaultValue, validationRule, validationText, lookupBoundColumn, lookupColumnCount, columnWidths
    If Len(displayName) = 0 Then displayName = ReadableDisplayName(fld.Name)
    If Len(sourceObjectName) = 0 Then sourceObjectName = objectName
    If Len(sourceFieldName) = 0 And isTableField Then sourceFieldName = fld.Name
    If isTableField = False And Len(sourceObjectName) > 0 And Len(sourceFieldName) > 0 Then
        If StrComp(sourceObjectName, objectName, vbTextCompare) <> 0 Then NonFormReferenceApplySourceFieldMetadata db, sourceFieldMetadataCache, sourceTableDefinitionCache, sourceObjectName, sourceFieldName, descriptionText, displayName, rowSourceType, rowSource, isRequired, allowsZeroLength, DefaultValue, validationRule, validationText, lookupBoundColumn, lookupColumnCount, columnWidths
    End If
    defaultDescriptionShortText = descriptionText
    defaultDescriptionLongText = descriptionText
    defaultDescriptionShortSource = "DAO Field.Description"
    defaultDescriptionLongSource = "DAO Field.Description"
    If isTableField And Len(Trim$(defaultDescriptionShortText)) = 0 Then
        defaultDescriptionResolutionSource = vbNullString
        defaultDescriptionShortText = ManagedObjectCatalogFieldDescriptionShort(objectName, fld.Name, defaultDescriptionResolutionSource)
        defaultDescriptionShortSource = CStr(defaultDescriptionResolutionSource)
        If Len(Trim$(defaultDescriptionShortSource)) = 0 Then defaultDescriptionShortSource = "ManagedObjectCatalogFieldDescriptionShort"
        If FieldDescriptionTextIsPlaceholder(defaultDescriptionShortText) Then defaultDescriptionShortText = vbNullString
    End If
    If isTableField And Len(Trim$(defaultDescriptionLongText)) = 0 Then
        defaultDescriptionResolutionSource = vbNullString
        defaultDescriptionLongText = ManagedObjectCatalogFieldDescriptionLong(objectName, fld.Name, defaultDescriptionResolutionSource)
        defaultDescriptionLongSource = CStr(defaultDescriptionResolutionSource)
        If Len(Trim$(defaultDescriptionLongSource)) = 0 Then defaultDescriptionLongSource = "ManagedObjectCatalogFieldDescriptionLong"
        If FieldDescriptionTextIsPlaceholder(defaultDescriptionLongText) Then defaultDescriptionLongText = vbNullString
    End If
    dataTypeName = NonFormReferenceDaoTypeName(fld.Type)
    requiredFieldSizeDefault = Null
    If isTableField Then requiredFieldSizeDefault = DefaultManagedFieldReqSize(objectName, fld.Name)
    valuePreservationPolicy = DefaultFieldValuePreservationPolicy(isTableField)
    searchKeywords = objectName & "; " & fld.Name & "; " & displayName & "; " & dataTypeName
    desiredDocumentationSource = DefaultFieldDocumentationSource(Len(descriptionText) > 0)
    If isTableField Then
        If Not indexFlagsByField Is Nothing Then
            If indexFlagsByField.Exists(fld.Name) Then indexFlags = CLng(indexFlagsByField(fld.Name))
            isIndexed = ((indexFlags And 1) <> 0)
            isPrimaryKey = ((indexFlags And 2) <> 0)
            isUnique = ((indexFlags And 4) <> 0)
        End If
    End If
    structuralText = objectName & "|" & fld.Name & "|" & CStr(fld.ordinalPosition) & "|" & CStr(fld.Type) & "|" & CStr(fld.Size) & "|" & CStr(isRequired) & "|" & CStr(allowsZeroLength) & "|" & DefaultValue & "|" & validationRule & "|" & rowSourceType & "|" & rowSource
    structuralFingerprint = FingerprintFromPayload(structuralText)
    seenPhysicalFields(fld.Name) = True
    isNew = (selectedIDs.Exists(fld.Name) = False)
    If isNew Then
        fieldRs.AddNew
        If coldBulkMode Then
            selectedID = nextColdFieldID
            nextColdFieldID = nextColdFieldID + 1
            fieldRs.fields(FIELD_ID).value = selectedID
        End If
        rowChanged = True
        schemaEvidenceChanged = True
        existingDescriptionShort = vbNullString
        existingDescriptionLong = vbNullString
        existingUsageGuidance = vbNullString
        documentationSource = desiredDocumentationSource
        If Len(defaultDescriptionLongText) > 0 Then
            documentationStatus = DOCUMENTATION_STATUS_COMPLETE
        Else
            documentationStatus = DOCUMENTATION_STATUS_MISSING_DESCRIPTION
        End If
    Else
        fieldRs.Bookmark = selectedBookmarks(fld.Name)
        selectedID = CLng(selectedIDs(fld.Name))
        rowChanged = Not NonFormDatabaseReferenceFieldRowIsCurrent(fieldRs, fld, displayName, sourceObjectName, sourceFieldName, dataTypeName, isRequired, allowsZeroLength, _
            isPrimaryKey, isIndexed, isUnique, DefaultValue, validationRule, validationText, rowSourceType, rowSource, lookupBoundColumn, lookupColumnCount, columnWidths, _
            defaultDescriptionShortText, defaultDescriptionLongText, structuralFingerprint, requiredFieldSizeDefault, existingDescriptionShort, existingDescriptionLong, existingUsageGuidance, _
            documentationSource, documentationStatus, schemaEvidenceChanged)
    End If
    If rowChanged = False Then
        If Len(Trim$(existingDescriptionLong)) = 0 And Len(Trim$(defaultDescriptionLongText)) = 0 Then fieldMissingDescription = True
        SynchronizeNonFormDatabaseReferenceFieldValuesForSetup db, valueRs, activeValueFieldIDs, valuePrimaryBookmarkByIdentity, valuePrimaryIDByIdentity, valueBookmarksByFieldID, selectedID, dataTypeName, rowSourceType, rowSource, lookupColumnCount, DefaultValue, isTableField, coldBulkMode, rowSourceResultCache, rowSourceOpenFailureCache, synchronizedValueCount, changedValueCount
        Exit Sub
    End If
    If isNew = False Then fieldRs.Edit
    fieldRs.fields(FIELD_MANAGED_OBJECT_ID).value = managedObjectID
    If isNew Then fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_NAME).value = fld.Name
    fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_ORDINAL).value = fld.ordinalPosition
    fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME).value = displayName
    fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_SOURCE_OBJECT_NAME).value = sourceObjectName
    If Len(sourceFieldName) > 0 Then
        fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_SOURCE_FIELD_NAME).value = sourceFieldName
    Else
        fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_SOURCE_FIELD_NAME).value = Null
    End If
    fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_DATA_TYPE).value = dataTypeName
    fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_ACTUAL_SIZE).value = fld.Size
    If IsNull(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_REQUIRED_SIZE).value) And IsNull(requiredFieldSizeDefault) = False Then fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_REQUIRED_SIZE).value = CLng(requiredFieldSizeDefault)
    fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_IS_REQUIRED).value = isRequired
    fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_ALLOWS_ZERO_LENGTH).value = allowsZeroLength
    fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_IS_PRIMARY_KEY).value = isPrimaryKey
    fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_IS_INDEXED).value = isIndexed
    fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_IS_UNIQUE).value = isUnique
    If Len(DefaultValue) > 0 Then
        SetRecordsetFieldIfExists fieldRs, FIELD_MANAGED_OBJECT_FIELD_DEFAULT_VALUE, DefaultValue, "DAO Field.DefaultValue"
    Else
        fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_DEFAULT_VALUE).value = Null
    End If
    If Len(validationRule) > 0 Then
        fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALIDATION_RULE).value = validationRule
    Else
        fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALIDATION_RULE).value = Null
    End If
    If Len(validationText) > 0 Then
        SetRecordsetFieldIfExists fieldRs, FIELD_MANAGED_OBJECT_FIELD_VALIDATION_TEXT, validationText, "DAO Field.ValidationText"
    Else
        fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALIDATION_TEXT).value = Null
    End If
    If Len(rowSourceType) > 0 Then
        fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_LOOKUP_ROW_SOURCE_TYPE).value = rowSourceType
    Else
        fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_LOOKUP_ROW_SOURCE_TYPE).value = Null
    End If
    If Len(rowSource) > 0 Then
        fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_LOOKUP_ROW_SOURCE).value = rowSource
    Else
        fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_LOOKUP_ROW_SOURCE).value = Null
    End If
    fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_LOOKUP_BOUND_COLUMN).value = lookupBoundColumn
    fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_LOOKUP_COLUMN_COUNT).value = lookupColumnCount
    If Len(columnWidths) > 0 Then
        fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_LOOKUP_COLUMN_WIDTHS).value = columnWidths
    Else
        fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_LOOKUP_COLUMN_WIDTHS).value = Null
    End If
    If Len(existingDescriptionShort) = 0 And Len(defaultDescriptionShortText) > 0 Then SetRecordsetFieldIfExists fieldRs, FIELD_MANAGED_OBJECT_FIELD_DESCRIPTION_SHORT, defaultDescriptionShortText, defaultDescriptionShortSource
    If Len(existingDescriptionLong) = 0 And Len(defaultDescriptionLongText) > 0 Then SetRecordsetFieldIfExists fieldRs, FIELD_MANAGED_OBJECT_FIELD_DESCRIPTION_LONG, defaultDescriptionLongText, defaultDescriptionLongSource
    If Len(existingUsageGuidance) = 0 Then
        If Len(existingDescriptionLong) > 0 Then
            fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_USAGE_GUIDANCE).value = existingDescriptionLong
        ElseIf Len(defaultDescriptionLongText) > 0 Then
            SetRecordsetFieldValueWithSource fieldRs, FIELD_MANAGED_OBJECT_FIELD_USAGE_GUIDANCE, defaultDescriptionLongText, "ManagedObjectCatalogFieldDescriptionLong"
        Else
            fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_USAGE_GUIDANCE).value = "Use this field through the owning " & DEFAULT_UTILITY_LONG_NAME & " form, query, or maintenance workflow."
        End If
    End If
    If isNew Then
        fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_IS_USER_EDITABLE).value = False
        fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_IS_SYSTEM_MAINTAINED).value = True
    End If
    If Len(CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_PRESERVATION_POLICY).value, vbNullString))) = 0 Then SetRecordsetFieldValueWithSource fieldRs, FIELD_MANAGED_OBJECT_FIELD_VALUE_PRESERVATION_POLICY, valuePreservationPolicy, "DefaultFieldValuePreservationPolicy"
    If Len(CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_SEARCH_KEYWORDS).value, vbNullString))) = 0 Then fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_SEARCH_KEYWORDS).value = searchKeywords
    If isNew Or Len(documentationSource) = 0 Or StrComp(documentationSource, "ADC B5 physical schema index", vbTextCompare) = 0 Then SetRecordsetFieldValueWithSource fieldRs, FIELD_MANAGED_OBJECT_FIELD_DOCUMENTATION_SOURCE, desiredDocumentationSource, "DefaultFieldDocumentationSource"
    If Len(existingDescriptionLong) > 0 Or Len(defaultDescriptionLongText) > 0 Then
        fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_DOCUMENTATION_STATUS).value = DOCUMENTATION_STATUS_COMPLETE
    Else
        fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_DOCUMENTATION_STATUS).value = DOCUMENTATION_STATUS_MISSING_DESCRIPTION
    End If
    If schemaEvidenceChanged Then
        fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_SCHEMA_FINGERPRINT).value = Left$(structuralFingerprint, 255)
        fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_SCHEMA_VERIFIED_ON).value = Now()
    End If
    fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_IS_ACTIVE).value = True
    fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_RETIRED_ON).value = Null
    fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_RETIREMENT_REASON).value = Null
    fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_DOCUMENTATION_UPDATED_ON).value = Now()
    fieldRs.Update
    If isNew Then
        selectedIDs.Add fld.Name, selectedID
        If coldBulkMode = False Then
            fieldRs.Bookmark = fieldRs.LastModified
            selectedID = CLng(fieldRs.fields(FIELD_ID).value)
            selectedIDs(fld.Name) = selectedID
            selectedBookmarks.Add fld.Name, fieldRs.Bookmark
            objectFieldBookmarks.Add CStr(selectedID), fieldRs.Bookmark
        End If
    End If
    changedFieldCount = changedFieldCount + 1
    If Len(Trim$(existingDescriptionLong)) = 0 And Len(Trim$(defaultDescriptionLongText)) = 0 Then fieldMissingDescription = True
    SynchronizeNonFormDatabaseReferenceFieldValuesForSetup db, valueRs, activeValueFieldIDs, valuePrimaryBookmarkByIdentity, valuePrimaryIDByIdentity, valueBookmarksByFieldID, selectedID, dataTypeName, rowSourceType, rowSource, lookupColumnCount, DefaultValue, isTableField, coldBulkMode, rowSourceResultCache, rowSourceOpenFailureCache, synchronizedValueCount, changedValueCount
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "SynchronizeNonFormDatabaseReferenceFieldForSetup", "Could not synchronize a Database Reference field row." & vbCrLf & "Object: " & objectName & vbCrLf & "Field: " & fld.Name & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub MarkNonFormDatabaseReferenceFieldRowsInactiveForSetup(ByVal db As DAO.Database, ByVal fieldRs As DAO.Recordset, ByVal objectFieldBookmarks As Object, ByVal selectedIDs As Object, ByVal seenPhysicalFields As Object, ByVal activeValueFieldIDs As Object, ByRef inactiveFieldCount As Long, ByRef changedFieldCount As Long, ByRef changedValueCount As Long)
    If objectFieldBookmarks Is Nothing Then Exit Sub
    Dim fieldName As String
    Dim rowID As Long
    Dim inactiveReason As String
    Dim isMissing As Boolean
    Dim rowBookmarkKey As Variant
    For Each rowBookmarkKey In objectFieldBookmarks.keys
        fieldRs.Bookmark = objectFieldBookmarks(CStr(rowBookmarkKey))
        PumpSetupResponsiveness
        If CBool(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_IS_ACTIVE).value, False)) Then
            fieldName = CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_NAME).value, vbNullString))
            rowID = CLng(Nz(fieldRs.fields(FIELD_ID).value, 0))
            inactiveReason = vbNullString
            isMissing = False
            If Len(fieldName) = 0 Or seenPhysicalFields.Exists(fieldName) = False Then
                inactiveReason = "Field was not present during the latest Phase-B physical schema synchronization."
                isMissing = True
            ElseIf selectedIDs.Exists(fieldName) And rowID <> CLng(selectedIDs(fieldName)) Then
                inactiveReason = "Duplicate field identity marked inactive during current physical schema synchronization."
            End If
            If Len(inactiveReason) > 0 Then
                fieldRs.Edit
                fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_IS_ACTIVE).value = False
                If IsNull(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_RETIRED_ON).value) Then fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_RETIRED_ON).value = Now()
                If Len(CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_RETIREMENT_REASON).value, vbNullString))) = 0 Then fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_RETIREMENT_REASON).value = inactiveReason
                fieldRs.Update
                If rowID > 0 Then
                    If activeValueFieldIDs.Exists(CStr(rowID)) Then
                        db.Execute "UPDATE [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFieldValues") & "] SET [" & FIELD_MANAGED_OBJECT_FIELD_VALUE_IS_ACTIVE & "]=False,[" & FIELD_MANAGED_OBJECT_FIELD_VALUE_DOCUMENTATION_UPDATED_ON & "]=Now() WHERE [" & FIELD_MANAGED_OBJECT_FIELD_VALUE_MANAGED_OBJECT_FIELD_ID & "]=" & CStr(rowID) & " AND [" & FIELD_MANAGED_OBJECT_FIELD_VALUE_IS_ACTIVE & "]=True;", dbFailOnError
                        changedValueCount = changedValueCount + db.RecordsAffected
                        activeValueFieldIDs.Remove CStr(rowID)
                    End If
                End If
                If isMissing Then
                    inactiveFieldCount = inactiveFieldCount + 1
                Else
                    changedFieldCount = changedFieldCount + 1
                End If
            End If
        End If
    Next rowBookmarkKey
End Sub

Private Sub SynchronizeNonFormDatabaseReferenceStoredFieldValuesForSetup(ByVal db As DAO.Database, ByVal fieldRs As DAO.Recordset, ByVal objectFieldBookmarks As Object, ByVal selectedIDs As Object, ByVal valueRs As DAO.Recordset, ByVal activeValueFieldIDs As Object, ByVal valuePrimaryBookmarkByIdentity As Object, ByVal valuePrimaryIDByIdentity As Object, ByVal valueBookmarksByFieldID As Object, ByVal rowSourceResultCache As Object, ByVal rowSourceOpenFailureCache As Object, ByRef synchronizedValueCount As Long, ByRef changedValueCount As Long)
    If objectFieldBookmarks Is Nothing Then Exit Sub
    Dim fieldName As String
    Dim rowID As Long
    Dim rowBookmarkKey As Variant
    For Each rowBookmarkKey In objectFieldBookmarks.keys
        fieldRs.Bookmark = objectFieldBookmarks(CStr(rowBookmarkKey))
        PumpSetupResponsiveness
        fieldName = CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_NAME).value, vbNullString))
        rowID = CLng(Nz(fieldRs.fields(FIELD_ID).value, 0))
        If CBool(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_IS_ACTIVE).value, False)) And Len(fieldName) > 0 Then
            If selectedIDs.Exists(fieldName) Then
                If rowID = CLng(selectedIDs(fieldName)) Then SynchronizeNonFormDatabaseReferenceFieldValuesForSetup db, valueRs, activeValueFieldIDs, valuePrimaryBookmarkByIdentity, valuePrimaryIDByIdentity, valueBookmarksByFieldID, rowID, CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_DATA_TYPE).value, vbNullString)), CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_LOOKUP_ROW_SOURCE_TYPE).value, vbNullString)), CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_LOOKUP_ROW_SOURCE).value, vbNullString)), CLng(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_LOOKUP_COLUMN_COUNT).value, 0)), CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_DEFAULT_VALUE).value, vbNullString)), False, False, rowSourceResultCache, rowSourceOpenFailureCache, synchronizedValueCount, changedValueCount
            End If
        End If
    Next rowBookmarkKey
End Sub

Private Sub SynchronizeNonFormDatabaseReferenceFieldValuesForSetup(ByVal db As DAO.Database, ByVal valueRs As DAO.Recordset, ByVal activeValueFieldIDs As Object, ByVal valuePrimaryBookmarkByIdentity As Object, ByVal valuePrimaryIDByIdentity As Object, ByVal valueBookmarksByFieldID As Object, ByVal managedObjectFieldID As Long, ByVal dataTypeName As String, ByVal rowSourceType As String, ByVal rowSource As String, ByVal lookupColumnCount As Long, ByVal DefaultValue As String, ByVal allowControlledValueEnumeration As Boolean, ByVal coldBulkMode As Boolean, ByVal rowSourceResultCache As Object, ByVal rowSourceOpenFailureCache As Object, ByRef synchronizedValueCount As Long, ByRef changedValueCount As Long)
    On Error GoTo ErrHandler
    Dim desired As Object
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim cacheKey As String
    Dim values As Variant
    Dim valueIndex As Long
    Dim stepSize As Long
    Dim SortOrder As Long
    Dim rowCount As Long
    Dim storedValue As String
    Dim displayValue As String
    Dim openErrNumber As Long
    Dim openErrDescription As String
    Dim cachedRows As Collection
    Dim cachedRow As Variant
    If managedObjectFieldID <= 0 Then Exit Sub
    ' INTENTIONAL KEYED LOOKUP: desired provides keyed membership/lookup over source rows or state already read by SynchronizeNonFormDatabaseReferenceFieldValuesForSetup; it is a derived access path, not a second
    ' authority.
    Set desired = CreateObject("Scripting.Dictionary")
    desired.CompareMode = vbTextCompare
    If allowControlledValueEnumeration = False Then
        If StrComp(Trim$(dataTypeName), "Boolean", vbTextCompare) = 0 Or Len(Trim$(rowSource)) > 0 Or activeValueFieldIDs.Exists(CStr(managedObjectFieldID)) Then mB5DatabaseReferenceQueryControlledValueSkipCount = mB5DatabaseReferenceQueryControlledValueSkipCount + 1
        If activeValueFieldIDs.Exists(CStr(managedObjectFieldID)) Then NonFormReferenceApplyDesiredFieldValues valueRs, activeValueFieldIDs, valuePrimaryBookmarkByIdentity, valuePrimaryIDByIdentity, valueBookmarksByFieldID, managedObjectFieldID, desired, coldBulkMode, synchronizedValueCount, changedValueCount
        Exit Sub
    End If
    If StrComp(Trim$(dataTypeName), "Boolean", vbTextCompare) <> 0 And Len(Trim$(rowSource)) = 0 Then
        If activeValueFieldIDs.Exists(CStr(managedObjectFieldID)) = False Then Exit Sub
    End If
    mB5DatabaseReferenceControlledValueFieldCount = mB5DatabaseReferenceControlledValueFieldCount + 1
    If StrComp(Trim$(dataTypeName), "Boolean", vbTextCompare) = 0 Then
        NonFormReferenceAddDesiredFieldValue desired, "False", "False", "The Boolean field is not selected or enabled.", 10, NonFormReferenceValueIsDefault("False", DefaultValue), "DAO Boolean type"
        NonFormReferenceAddDesiredFieldValue desired, "True", "True", "The Boolean field is selected or enabled.", 20, NonFormReferenceValueIsDefault("True", DefaultValue), "DAO Boolean type"
    ElseIf UCase$(Trim$(rowSourceType)) = "VALUE LIST" And Len(Trim$(rowSource)) > 0 Then
        values = Split(rowSource, ";")
        stepSize = lookupColumnCount
        If stepSize < 1 Then stepSize = 1
        valueIndex = LBound(values)
        Do While valueIndex <= UBound(values)
            SortOrder = SortOrder + 1
            storedValue = Trim$(CStr(values(valueIndex)))
            displayValue = storedValue
            If stepSize > 1 And valueIndex + 1 <= UBound(values) Then displayValue = Trim$(CStr(values(valueIndex + 1)))
            If Len(storedValue) > 0 Then NonFormReferenceAddDesiredFieldValue desired, storedValue, displayValue, vbNullString, SortOrder, NonFormReferenceValueIsDefault(storedValue, DefaultValue), "Access Value List"
            valueIndex = valueIndex + stepSize
        Loop
    ElseIf Len(Trim$(rowSource)) > 0 Then
        sqlText = Trim$(rowSource)
        If UCase$(Left$(sqlText, 6)) <> "SELECT" Then
            If SharedDao_TableExists(db, sqlText) Or SharedDao_QueryExists(db, sqlText) Then
                sqlText = "SELECT * FROM [" & Replace(sqlText, "]", "]]") & "];"
            Else
                AddTraceNote "Database Reference field-value rows retained because the lookup RowSource object is unavailable. ManagedObjectFieldID=" & CStr(managedObjectFieldID) & "; RowSource=" & rowSource
                Exit Sub
            End If
        End If
        cacheKey = sqlText
        If rowSourceOpenFailureCache.Exists(cacheKey) Then
            mB5DatabaseReferenceLookupRowSourceCacheHitCount = mB5DatabaseReferenceLookupRowSourceCacheHitCount + 1
            Exit Sub
        End If
        If rowSourceResultCache.Exists(cacheKey) Then
            mB5DatabaseReferenceLookupRowSourceCacheHitCount = mB5DatabaseReferenceLookupRowSourceCacheHitCount + 1
            Set cachedRows = rowSourceResultCache(cacheKey)
        Else
            ' INTENTIONAL OPERATION SNAPSHOT: cachedRows captures one coherent source state for this invocation of SynchronizeNonFormDatabaseReferenceFieldValuesForSetup; validity ends with the operation rather than
            ' becoming a live data authority.
            Set cachedRows = New Collection
            mB5DatabaseReferenceLookupRowSourceOpenCount = mB5DatabaseReferenceLookupRowSourceOpenCount + 1
            On Error Resume Next
            Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
            openErrNumber = Err.Number
            openErrDescription = Err.Description
            Err.Clear
            On Error GoTo ErrHandler
            If openErrNumber <> 0 Or rs Is Nothing Then
                rowSourceOpenFailureCache.Add cacheKey, CStr(openErrNumber) & ": " & openErrDescription
                AddTraceNote "Database Reference field-value rows retained because the lookup RowSource could not be opened. ManagedObjectFieldID=" & CStr(managedObjectFieldID) & "; Error " & CStr(openErrNumber) & ": " & openErrDescription
                Exit Sub
            End If
            Do While Not rs.EOF And rowCount < 200
                PumpSetupResponsiveness
                rowCount = rowCount + 1
                storedValue = CStr(Nz(rs.fields(0).value, vbNullString))
                displayValue = storedValue
                If rs.fields.count > 1 Then displayValue = CStr(Nz(rs.fields(1).value, storedValue))
                If Len(storedValue) > 0 Then cachedRows.Add Array(storedValue, displayValue, rowCount)
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing
            rowSourceResultCache.Add cacheKey, cachedRows
        End If
        For Each cachedRow In cachedRows
            NonFormReferenceAddDesiredFieldValue desired, CStr(cachedRow(0)), CStr(cachedRow(1)), vbNullString, CLng(cachedRow(2)), NonFormReferenceValueIsDefault(CStr(cachedRow(0)), DefaultValue), "Lookup RowSource"
        Next cachedRow
    End If
    NonFormReferenceApplyDesiredFieldValues valueRs, activeValueFieldIDs, valuePrimaryBookmarkByIdentity, valuePrimaryIDByIdentity, valueBookmarksByFieldID, managedObjectFieldID, desired, coldBulkMode, synchronizedValueCount, changedValueCount
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
    Err.Raise savedErrNumber, "SynchronizeNonFormDatabaseReferenceFieldValuesForSetup", "Controlled-value Database Reference synchronization failed." & vbCrLf & "ManagedObjectFieldID: " & CStr(managedObjectFieldID) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub NonFormReferenceUpdateObjectDocumentationStatusForSetup(ByVal parentRs As DAO.Recordset, ByVal fieldMissingDescription As Boolean, ByVal alreadyCountedAsChanged As Boolean, ByRef changedObjectCount As Long)
    Dim expectedStatus As String
    Dim objectMissingDescription As Boolean
    Dim rowChanged As Boolean
    Dim objectName As String
    Dim objectType As String
    Dim objectDisplayName As String
    Dim ObjectDescription As String
    Dim categoryCode As String
    Dim catalogOwner As String
    Dim commonName As String
    Dim keywords As String
    Dim usageGuidance As String
    Dim editPolicy As String
    Dim lifecyclePolicy As String
    Dim repairOwner As String
    Dim deletionPolicy As String
    Dim primaryWorkflow As String
    Dim enforceCanonicalEditPolicy As Boolean
    objectName = CStr(Nz(parentRs.fields(FIELD_OBJECT_NAME).value, vbNullString))
    objectType = Trim$(CStr(Nz(parentRs.fields(FIELD_OBJECT_TYPE).value, vbNullString)))
    objectDisplayName = Trim$(CStr(Nz(parentRs.fields(FIELD_OBJECT_DISPLAY_NAME).value, vbNullString)))
    ObjectDescription = Trim$(CStr(Nz(parentRs.fields(FIELD_OBJECT_DESCRIPTION).value, vbNullString)))
    categoryCode = Trim$(CStr(Nz(parentRs.fields(FIELD_CATEGORY_CODE).value, vbNullString)))
    catalogOwner = Trim$(CStr(Nz(parentRs.fields(FIELD_CATALOG_OWNER).value, vbNullString)))
    commonName = objectDisplayName
    If Len(commonName) = 0 Then commonName = ReadableDisplayName(objectName)
    keywords = objectName & "; " & objectType & "; " & categoryCode & "; " & catalogOwner
    usageGuidance = DefaultObjectUsageGuidance(ObjectDescription)
    editPolicy = DefaultObjectEditPolicy(catalogOwner, objectName)
    enforceCanonicalEditPolicy = (StrComp(objectType, "Query", vbTextCompare) = 0 And StrComp(objectName, ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_Terminology_Review"), vbTextCompare) = 0)
    lifecyclePolicy = DefaultObjectLifecyclePolicy()
    repairOwner = DefaultObjectRepairOwner()
    deletionPolicy = DefaultObjectDeletionPolicy()
    primaryWorkflow = NonFormReferenceDocumentationWorkflowForCategory(categoryCode)
    objectMissingDescription = (Len(ObjectDescription) = 0)
    If objectMissingDescription Or fieldMissingDescription Then
        expectedStatus = DOCUMENTATION_STATUS_MISSING_DESCRIPTION
    Else
        expectedStatus = DOCUMENTATION_STATUS_COMPLETE
    End If
    If Len(CStr(Nz(parentRs.fields(FIELD_COMMON_NAME).value, vbNullString))) = 0 Then rowChanged = True
    If Len(CStr(Nz(parentRs.fields(FIELD_MANAGED_OBJECT_FIELD_SEARCH_KEYWORDS).value, vbNullString))) = 0 Then rowChanged = True
    If Len(CStr(Nz(parentRs.fields(FIELD_MANAGED_OBJECT_FIELD_USAGE_GUIDANCE).value, vbNullString))) = 0 Then rowChanged = True
    If Len(CStr(Nz(parentRs.fields(FIELD_EDIT_POLICY).value, vbNullString))) = 0 Then rowChanged = True
    If enforceCanonicalEditPolicy And StrComp(CStr(Nz(parentRs.fields(FIELD_EDIT_POLICY).value, vbNullString)), editPolicy, vbBinaryCompare) <> 0 Then rowChanged = True
    If Len(CStr(Nz(parentRs.fields(FIELD_LIFECYCLE_POLICY).value, vbNullString))) = 0 Then rowChanged = True
    If Len(CStr(Nz(parentRs.fields(FIELD_REPAIR_OWNER).value, vbNullString))) = 0 Then rowChanged = True
    If Len(CStr(Nz(parentRs.fields(FIELD_DELETION_POLICY).value, vbNullString))) = 0 Then rowChanged = True
    If Len(CStr(Nz(parentRs.fields(FIELD_PRIMARY_WORKFLOW).value, vbNullString))) = 0 Then rowChanged = True
    If StrComp(CStr(Nz(parentRs.fields(FIELD_MANAGED_OBJECT_DEPENDENCY_DOCUMENTATION_STATUS).value, vbNullString)), expectedStatus, vbTextCompare) <> 0 Then rowChanged = True
    If CBool(Nz(parentRs.fields(FIELD_IS_DOCUMENTATION_ACTIVE).value, False)) = False Then rowChanged = True
    If Len(CStr(Nz(parentRs.fields(FIELD_MANAGED_OBJECT_FIELD_DOCUMENTATION_SOURCE).value, vbNullString))) = 0 Then rowChanged = True
    If rowChanged = False Then Exit Sub
    If SetupRollbackTrackingActive() And alreadyCountedAsChanged = False Then TrackSetupManagedObjectRowRollback CLng(parentRs.fields(FIELD_ID).value), True, ManagedObjectRollbackSql(parentRs)
    parentRs.Edit
    If Len(CStr(Nz(parentRs.fields(FIELD_COMMON_NAME).value, vbNullString))) = 0 Then parentRs.fields(FIELD_COMMON_NAME).value = commonName
    If Len(CStr(Nz(parentRs.fields(FIELD_MANAGED_OBJECT_FIELD_SEARCH_KEYWORDS).value, vbNullString))) = 0 Then parentRs.fields(FIELD_MANAGED_OBJECT_FIELD_SEARCH_KEYWORDS).value = keywords
    If Len(CStr(Nz(parentRs.fields(FIELD_MANAGED_OBJECT_FIELD_USAGE_GUIDANCE).value, vbNullString))) = 0 Then SetRecordsetFieldValueWithSource parentRs, FIELD_MANAGED_OBJECT_FIELD_USAGE_GUIDANCE, usageGuidance, "DefaultObjectUsageGuidance"
    If Len(CStr(Nz(parentRs.fields(FIELD_EDIT_POLICY).value, vbNullString))) = 0 Or (enforceCanonicalEditPolicy And StrComp(CStr(Nz(parentRs.fields(FIELD_EDIT_POLICY).value, vbNullString)), editPolicy, vbBinaryCompare) <> 0) Then SetRecordsetFieldValueWithSource parentRs, FIELD_EDIT_POLICY, editPolicy, "DefaultObjectEditPolicy"
    If Len(CStr(Nz(parentRs.fields(FIELD_LIFECYCLE_POLICY).value, vbNullString))) = 0 Then SetRecordsetFieldValueWithSource parentRs, FIELD_LIFECYCLE_POLICY, lifecyclePolicy, "DefaultObjectLifecyclePolicy"
    If Len(CStr(Nz(parentRs.fields(FIELD_REPAIR_OWNER).value, vbNullString))) = 0 Then SetRecordsetFieldValueWithSource parentRs, FIELD_REPAIR_OWNER, repairOwner, "DefaultObjectRepairOwner"
    If Len(CStr(Nz(parentRs.fields(FIELD_DELETION_POLICY).value, vbNullString))) = 0 Then SetRecordsetFieldValueWithSource parentRs, FIELD_DELETION_POLICY, deletionPolicy, "DefaultObjectDeletionPolicy"
    If Len(CStr(Nz(parentRs.fields(FIELD_PRIMARY_WORKFLOW).value, vbNullString))) = 0 Then SetRecordsetFieldValueWithSource parentRs, FIELD_PRIMARY_WORKFLOW, primaryWorkflow, "NonFormReferenceDocumentationWorkflowForCategory"
    parentRs.fields(FIELD_MANAGED_OBJECT_DEPENDENCY_DOCUMENTATION_STATUS).value = expectedStatus
    parentRs.fields(FIELD_IS_DOCUMENTATION_ACTIVE).value = True
    If Len(CStr(Nz(parentRs.fields(FIELD_MANAGED_OBJECT_FIELD_DOCUMENTATION_SOURCE).value, vbNullString))) = 0 Then SetRecordsetFieldValueWithSource parentRs, FIELD_MANAGED_OBJECT_FIELD_DOCUMENTATION_SOURCE, DefaultObjectDocumentationSource(), "DefaultObjectDocumentationSource"
    parentRs.fields(FIELD_MANAGED_OBJECT_FIELD_DOCUMENTATION_UPDATED_ON).value = Now()
    parentRs.Update
    If alreadyCountedAsChanged = False Then changedObjectCount = changedObjectCount + 1
End Sub

Private Sub NonFormReferenceApplySourceFieldMetadata(ByVal db As DAO.Database, ByVal sourceFieldMetadataCache As Object, ByVal sourceTableDefinitionCache As Object, ByVal sourceObjectName As String, ByVal sourceFieldName As String, ByRef descriptionText As String, ByRef displayName As String, ByRef rowSourceType As String, ByRef rowSource As String, ByRef isRequired As Boolean, ByRef allowsZeroLength As Boolean, ByRef DefaultValue As String, ByRef validationRule As String, ByRef validationText As String, ByRef lookupBoundColumn As Long, ByRef lookupColumnCount As Long, ByRef columnWidths As String)
    On Error GoTo MissingMetadata
    Dim tdf As DAO.tableDef
    Dim sourceField As DAO.Field
    Dim cacheKey As String
    Dim sourceObjectKey As String
    Dim metadata As Variant
    Dim sourceDescription As String
    Dim sourceDisplayName As String
    Dim sourceRowSourceType As String
    Dim sourceRowSource As String
    Dim sourceRequired As Boolean
    Dim sourceAllowsZeroLength As Boolean
    Dim sourceDefaultValue As String
    Dim sourceValidationRule As String
    Dim sourceValidationText As String
    Dim sourceBoundColumn As Long
    Dim sourceColumnCount As Long
    Dim sourceColumnWidths As String
    Dim ignoredSourceObjectName As String
    Dim ignoredSourceFieldName As String
    cacheKey = UCase$(sourceObjectName) & "|" & UCase$(Trim$(sourceFieldName))
    If sourceFieldMetadataCache.Exists(cacheKey) Then
        metadata = sourceFieldMetadataCache(cacheKey)
    Else
        sourceObjectKey = UCase$(sourceObjectName)
        If sourceTableDefinitionCache.Exists(sourceObjectKey) Then
            If IsObject(sourceTableDefinitionCache(sourceObjectKey)) = False Then sourceFieldMetadataCache(cacheKey) = Empty: Exit Sub
            Set tdf = sourceTableDefinitionCache(sourceObjectKey)
        Else
            On Error Resume Next
            Set tdf = db.TableDefs(sourceObjectName)
            If Err.Number <> 0 Or tdf Is Nothing Then
                Err.Clear
                sourceTableDefinitionCache.Add sourceObjectKey, False
                sourceFieldMetadataCache(cacheKey) = Empty
                On Error GoTo 0
                Exit Sub
            End If
            On Error GoTo MissingMetadata
            sourceTableDefinitionCache.Add sourceObjectKey, tdf
        End If
        Set sourceField = tdf.fields(sourceFieldName)
        NonFormReferenceReadFieldProperties sourceField, True, vbNullString, sourceDescription, sourceDisplayName, sourceRowSourceType, sourceRowSource, ignoredSourceObjectName, ignoredSourceFieldName, sourceRequired, sourceAllowsZeroLength, sourceDefaultValue, sourceValidationRule, sourceValidationText, sourceBoundColumn, sourceColumnCount, sourceColumnWidths
        metadata = Array(sourceDescription, sourceDisplayName, sourceRowSourceType, sourceRowSource, sourceRequired, sourceAllowsZeroLength, sourceDefaultValue, sourceValidationRule, sourceValidationText, sourceBoundColumn, sourceColumnCount, sourceColumnWidths)
        sourceFieldMetadataCache(cacheKey) = metadata
    End If
    If IsArray(metadata) = False Then Exit Sub
    If Len(descriptionText) = 0 Then descriptionText = CStr(metadata(0))
    If Len(displayName) = 0 Then displayName = CStr(metadata(1))
    rowSourceType = CStr(metadata(2))
    rowSource = CStr(metadata(3))
    isRequired = CBool(metadata(4))
    allowsZeroLength = CBool(metadata(5))
    DefaultValue = CStr(metadata(6))
    validationRule = CStr(metadata(7))
    validationText = CStr(metadata(8))
    lookupBoundColumn = CLng(metadata(9))
    lookupColumnCount = CLng(metadata(10))
    columnWidths = CStr(metadata(11))
    Exit Sub
MissingMetadata:
    Err.Clear
    If Len(cacheKey) > 0 Then sourceFieldMetadataCache(cacheKey) = Empty
End Sub

Private Function NonFormReferenceTableRowCount(ByVal db As DAO.Database, ByVal tableName As String) As Long
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT Count(*) " & "AS " & SQL_FIELD_ROW_COUNT & " FROM [" & Replace(tableName, "]", "]]") & "];", dbOpenSnapshot)
    NonFormReferenceTableRowCount = CLng(rs.fields(FIELD_ROW_COUNT).value)
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
    Err.Raise savedErrNumber, "NonFormReferenceTableRowCount", "Could not count Database Reference rows in [" & tableName & "]." & vbCrLf & savedErrDescription
End Function

Private Sub NonFormReferencePrepareColdStageTables(ByVal db As DAO.Database, ByRef fieldStageTableName As String, ByRef valueStageTableName As String)
    Dim stageSuffix As String
    stageSuffix = Format$(Now(), "yymmddhhnnss") & "_" & Right$("00000000" & CStr(CLng(Timer * 100#)), 8)
    fieldStageTableName = Left$("z__ADC_B5_FieldStage_" & stageSuffix, 64)
    valueStageTableName = Left$("z__ADC_B5_ValueStage_" & stageSuffix, 64)
    On Error GoTo ErrHandler
    NonFormReferenceCreateColdStageTable db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFields"), fieldStageTableName
    NonFormReferenceCreateColdStageTable db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFieldValues"), valueStageTableName
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    Dim cleanupDetail As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    NonFormReferenceCleanupColdStageTables db, fieldStageTableName, valueStageTableName, cleanupDetail
    Err.Raise savedErrNumber, "NonFormReferencePrepareColdStageTables", "Could not prepare the cold B5 staging tables." & vbCrLf & savedErrDescription & IIf(Len(cleanupDetail) > 0, vbCrLf & cleanupDetail, vbNullString)
End Sub

Private Sub NonFormReferenceCreateColdStageTable(ByVal db As DAO.Database, ByVal targetTableName As String, ByVal stageTableName As String)
    On Error GoTo ErrHandler
    Dim targetTdf As DAO.tableDef
    Dim stageTdf As DAO.tableDef
    Dim targetField As DAO.Field
    Dim stageField As DAO.Field
    Set targetTdf = db.TableDefs(targetTableName)
    Set stageTdf = db.CreateTableDef(stageTableName)
    stageTdf.Attributes = dbHiddenObject
    For Each targetField In targetTdf.fields
        If targetField.Type = dbText Or targetField.Type = dbBinary Then
            Set stageField = stageTdf.CreateField(targetField.Name, targetField.Type, targetField.Size)
        Else
            Set stageField = stageTdf.CreateField(targetField.Name, targetField.Type)
        End If
        'Transient B5 stage IDs are ordinary Long values so the cold path can assign deterministic temporary field identities without LastModified bookmark reads.
        stageTdf.fields.Append stageField
    Next targetField
    db.TableDefs.Append stageTdf
    db.TableDefs.Refresh
    Set stageTdf = db.TableDefs(stageTableName)
    Call EnsureDaoPropertyValue(stageTdf, DAO_DESCRIPTION_PROPERTY_NAME, dbText, Left$("Transient hidden B5 staging table for an atomic cold rebuild of [" & targetTableName & "]. It holds operation-local replacement rows only, is not catalog or runtime authority, and must be deleted after append or rollback cleanup.", 255), "NonFormReferenceCreateColdStageTable")
    For Each stageField In stageTdf.fields
        Call EnsureDaoPropertyValue(stageField, DAO_DESCRIPTION_PROPERTY_NAME, dbText, Left$("Transient B5 staging copy of [" & targetTableName & "].[" & stageField.Name & "]. Values exist only to assemble and verify the cold replacement set before the authoritative target-table append; do not edit or consume as durable metadata.", 255), "NonFormReferenceCreateColdStageTable")
    Next stageField
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "NonFormReferenceCreateColdStageTable", "Could not create the unindexed B5 staging table [" & stageTableName & "] from [" & targetTableName & "]." & vbCrLf & Err.Description
End Sub

Private Sub NonFormReferenceOpenColdStageIndexes(ByVal db As DAO.Database, ByVal fieldStageTableName As String, ByVal valueStageTableName As String, ByRef fieldRs As DAO.Recordset, ByRef fieldBookmarksByObjectID As Object, ByRef valueRs As DAO.Recordset, ByRef activeValueFieldIDs As Object, ByRef valuePrimaryBookmarkByIdentity As Object, ByRef valuePrimaryIDByIdentity As Object, ByRef valueBookmarksByFieldID As Object)
    ' INTENTIONAL KEYED LOOKUP: fieldBookmarksByObjectID, activeValueFieldIDs, valuePrimaryBookmarkByIdentity, valuePrimaryIDByIdentity, valueBookmarksByFieldID provides keyed membership/lookup over source rows or
    ' state already read by NonFormReferenceOpenColdStageIndexes; it is a derived access path, not a second authority.
    Set fieldBookmarksByObjectID = CreateObject("Scripting.Dictionary")
    fieldBookmarksByObjectID.CompareMode = vbBinaryCompare
    Set activeValueFieldIDs = CreateObject("Scripting.Dictionary")
    activeValueFieldIDs.CompareMode = vbTextCompare
    Set valuePrimaryBookmarkByIdentity = CreateObject("Scripting.Dictionary")
    valuePrimaryBookmarkByIdentity.CompareMode = vbTextCompare
    Set valuePrimaryIDByIdentity = CreateObject("Scripting.Dictionary")
    valuePrimaryIDByIdentity.CompareMode = vbTextCompare
    Set valueBookmarksByFieldID = CreateObject("Scripting.Dictionary")
    valueBookmarksByFieldID.CompareMode = vbTextCompare
    Set fieldRs = db.OpenRecordset("SELECT * FROM [" & Replace(fieldStageTableName, "]", "]]") & "];", dbOpenDynaset)
    Set valueRs = db.OpenRecordset("SELECT * FROM [" & Replace(valueStageTableName, "]", "]]") & "];", dbOpenDynaset)
End Sub

Private Sub NonFormReferenceAppendColdStageRows(ByVal db As DAO.Database, ByVal fieldStageTableName As String, ByVal valueStageTableName As String, ByRef detailText As String)
    On Error GoTo ErrHandler
    Dim fieldTargetColumns As String
    Dim fieldSelectColumns As String
    Dim valueTargetColumns As String
    Dim valueSelectColumns As String
    Dim separator As String
    Dim fieldSql As String
    Dim valueSql As String
    Dim stageFieldCount As Long
    Dim stageValueCount As Long
    Dim insertedFieldCount As Long
    Dim insertedValueCount As Long
    Dim targetTdf As DAO.tableDef
    Dim targetField As DAO.Field
    stageFieldCount = NonFormReferenceTableRowCount(db, fieldStageTableName)
    stageValueCount = NonFormReferenceTableRowCount(db, valueStageTableName)
    If NonFormReferenceTableRowCount(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFields")) <> 0 Or NonFormReferenceTableRowCount(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFieldValues")) <> 0 Then Err.Raise vbObjectError + 9581, "NonFormReferenceAppendColdStageRows", "The cold B5 target tables were no longer empty before the staged append."
    Set targetTdf = db.TableDefs(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFields"))
    For Each targetField In targetTdf.fields
        If StrComp(targetField.Name, FIELD_ID, vbTextCompare) <> 0 Then
            fieldTargetColumns = fieldTargetColumns & separator & "[" & Replace(targetField.Name, "]", "]]") & "]"
            fieldSelectColumns = fieldSelectColumns & separator & "[S].[" & Replace(targetField.Name, "]", "]]") & "]"
            separator = ","
        End If
    Next targetField
    fieldSql = "INSERT INTO [" & Replace(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFields"), "]", "]]") & "] (" & fieldTargetColumns & ") SELECT " & fieldSelectColumns & " FROM [" & Replace(fieldStageTableName, "]", "]]") & "] AS [S];"
    db.Execute fieldSql, dbFailOnError
    insertedFieldCount = db.RecordsAffected
    If insertedFieldCount <> stageFieldCount Then Err.Raise vbObjectError + 9582, "NonFormReferenceAppendColdStageRows", "The cold B5 field append count did not match the staged field count." & vbCrLf & "Staged: " & CStr(stageFieldCount) & vbCrLf & "Inserted: " & CStr(insertedFieldCount)
    separator = vbNullString
    Set targetTdf = db.TableDefs(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFieldValues"))
    For Each targetField In targetTdf.fields
        If StrComp(targetField.Name, FIELD_ID, vbTextCompare) <> 0 Then
            valueTargetColumns = valueTargetColumns & separator & "[" & Replace(targetField.Name, "]", "]]") & "]"
            If StrComp(targetField.Name, FIELD_MANAGED_OBJECT_FIELD_VALUE_MANAGED_OBJECT_FIELD_ID, vbTextCompare) = 0 Then
                valueSelectColumns = valueSelectColumns & separator & "[TF]." & SQL_FIELD_ID
            Else
                valueSelectColumns = valueSelectColumns & separator & "[V].[" & Replace(targetField.Name, "]", "]]") & "]"
            End If
            separator = ","
        End If
    Next targetField
    valueSql = "INSERT INTO [" & Replace(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFieldValues"), "]", "]]") & "] (" & valueTargetColumns & ") " & _
        "SELECT " & valueSelectColumns & " FROM ([" & Replace(valueStageTableName, "]", "]]") & "] AS [V] INNER JOIN [" & Replace(fieldStageTableName, "]", "]]") & "] AS [SF] ON [V].[" & FIELD_MANAGED_OBJECT_FIELD_VALUE_MANAGED_OBJECT_FIELD_ID & "]=[SF]." & SQL_FIELD_ID & ") " & _
        "INNER JOIN [" & Replace(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFields"), "]", "]]") & "] AS [TF] ON [TF].[" & FIELD_MANAGED_OBJECT_ID & "]=[SF].[" & FIELD_MANAGED_OBJECT_ID & "] AND [TF].[" & FIELD_MANAGED_OBJECT_FIELD_NAME & "]=[SF].[" & FIELD_MANAGED_OBJECT_FIELD_NAME & "];"
    db.Execute valueSql, dbFailOnError
    insertedValueCount = db.RecordsAffected
    If insertedValueCount <> stageValueCount Then Err.Raise vbObjectError + 9583, "NonFormReferenceAppendColdStageRows", "The cold B5 field-value append count did not match the staged value count." & vbCrLf & "Staged: " & CStr(stageValueCount) & vbCrLf & "Inserted: " & CStr(insertedValueCount)
    detailText = "Cold B5 two-stage apply: desired rows were written sequentially to unindexed transient tables with deterministic temporary field IDs and no per-row LastModified reads, then appended to the authoritative catalog in two set-based statements. Staged fields=" & CStr(stageFieldCount) & "; inserted fields=" & CStr(insertedFieldCount) & "; staged field values=" & CStr(stageValueCount) & "; inserted field values=" & CStr(insertedValueCount) & "; target append statements=2."
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "NonFormReferenceAppendColdStageRows", "Could not apply the cold B5 staged field and controlled-value rows to the authoritative catalog." & vbCrLf & Err.Description
End Sub

Private Sub NonFormReferenceCleanupColdStageTables(ByVal db As DAO.Database, ByVal fieldStageTableName As String, ByVal valueStageTableName As String, ByRef detailText As String)
    Dim cleanupErrors As String
    detailText = vbNullString
    On Error Resume Next
    If Len(valueStageTableName) > 0 Then
        If NonFormReferenceStageTableExistsStrict(db, valueStageTableName) Then db.TableDefs.Delete valueStageTableName
        If Err.Number <> 0 Then cleanupErrors = cleanupErrors & " [" & valueStageTableName & ": " & CStr(Err.Number) & " - " & Err.Description & "]"
        Err.Clear
        db.TableDefs.Refresh
        If NonFormReferenceStageTableExistsStrict(db, valueStageTableName) Then cleanupErrors = cleanupErrors & " [" & valueStageTableName & ": table still exists after deletion]"
        If Err.Number <> 0 Then cleanupErrors = cleanupErrors & " [" & valueStageTableName & " verification: " & CStr(Err.Number) & " - " & Err.Description & "]"
        Err.Clear
    End If
    If Len(fieldStageTableName) > 0 Then
        If NonFormReferenceStageTableExistsStrict(db, fieldStageTableName) Then db.TableDefs.Delete fieldStageTableName
        If Err.Number <> 0 Then cleanupErrors = cleanupErrors & " [" & fieldStageTableName & ": " & CStr(Err.Number) & " - " & Err.Description & "]"
        Err.Clear
        db.TableDefs.Refresh
        If NonFormReferenceStageTableExistsStrict(db, fieldStageTableName) Then cleanupErrors = cleanupErrors & " [" & fieldStageTableName & ": table still exists after deletion]"
        If Err.Number <> 0 Then cleanupErrors = cleanupErrors & " [" & fieldStageTableName & " verification: " & CStr(Err.Number) & " - " & Err.Description & "]"
        Err.Clear
    End If
    db.TableDefs.Refresh
    On Error GoTo 0
    If Len(cleanupErrors) > 0 Then detailText = "Cold B5 staging-table cleanup was incomplete; the hidden transient tables are not catalog authorities and will not be reused." & cleanupErrors
End Sub

Private Function NonFormReferenceStageTableExistsStrict(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
    'Returns False only for DAO collection-not-found; every inspection failure is raised so cleanup cannot claim an unverified absence.
    On Error GoTo InspectFailed
    Dim tdf As DAO.tableDef
    db.TableDefs.Refresh
    Set tdf = db.TableDefs(tableName)
    NonFormReferenceStageTableExistsStrict = True
    Exit Function
InspectFailed:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    If savedErrNumber = 3265 Then Err.Clear: Exit Function
    Err.Raise savedErrNumber, "NonFormReferenceStageTableExistsStrict", "Could not verify whether B5 staging table [" & tableName & "] exists." & vbCrLf & savedErrDescription
End Function

Private Sub NonFormReferenceOpenSharedFieldIndex(ByVal db As DAO.Database, ByRef fieldRs As DAO.Recordset, ByRef fieldBookmarksByObjectID As Object, ByRef indexedFieldRowCount As Long)
    On Error GoTo ErrHandler
    Dim managedObjectID As Long
    Dim objectKey As String
    Dim rowID As Long
    Dim rowBookmarks As Object
    ' INTENTIONAL KEYED LOOKUP: fieldBookmarksByObjectID provides keyed membership/lookup over source rows or state already read by NonFormReferenceOpenSharedFieldIndex; it is a derived access path, not a second
    ' authority.
    Set fieldBookmarksByObjectID = CreateObject("Scripting.Dictionary")
    fieldBookmarksByObjectID.CompareMode = vbBinaryCompare
    Set fieldRs = db.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFields") & "];", dbOpenDynaset)
    Do While fieldRs.EOF = False
        PumpSetupResponsiveness
        indexedFieldRowCount = indexedFieldRowCount + 1
        managedObjectID = CLng(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_ID).value, 0))
        rowID = CLng(Nz(fieldRs.fields(FIELD_ID).value, 0))
        If managedObjectID > 0 And rowID > 0 Then
            objectKey = CStr(managedObjectID)
            If fieldBookmarksByObjectID.Exists(objectKey) Then
                Set rowBookmarks = fieldBookmarksByObjectID(objectKey)
            Else
                ' INTENTIONAL KEYED LOOKUP: rowBookmarks provides keyed membership/lookup over source rows or state already read by NonFormReferenceOpenSharedFieldIndex; it is a derived access path, not a second
                ' authority.
                Set rowBookmarks = CreateObject("Scripting.Dictionary")
                rowBookmarks.CompareMode = vbBinaryCompare
                fieldBookmarksByObjectID.Add objectKey, rowBookmarks
            End If
            rowBookmarks.Add CStr(rowID), fieldRs.Bookmark
        End If
        fieldRs.MoveNext
    Loop
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not fieldRs Is Nothing Then fieldRs.Close
    Set fieldRs = Nothing
    On Error GoTo 0
    Err.Raise savedErrNumber, "NonFormReferenceOpenSharedFieldIndex", "Could not open/index the shared MAC_Fields synchronization recordset." & vbCrLf & savedErrDescription
End Sub

Private Sub NonFormReferenceBuildPhysicalDefinitionMaps(ByVal db As DAO.Database, ByRef physicalUpdatedByName As Object, ByRef physicalTypeByName As Object, ByRef physicalNameByKey As Object, ByRef querySqlByName As Object)
    On Error GoTo ErrHandler
    Dim tdf As DAO.tableDef
    Dim qdf As DAO.queryDef
    Dim objectKey As String
    Dim objectUpdated As Date
    Dim propertyFound As Boolean
    ' INTENTIONAL KEYED LOOKUP: physicalUpdatedByName, physicalTypeByName, physicalNameByKey, querySqlByName provides keyed membership/lookup over source rows or state already read by
    ' NonFormReferenceBuildPhysicalDefinitionMaps; it is a derived access path, not a second authority.
    Set physicalUpdatedByName = CreateObject("Scripting.Dictionary")
    physicalUpdatedByName.CompareMode = vbTextCompare
    Set physicalTypeByName = CreateObject("Scripting.Dictionary")
    physicalTypeByName.CompareMode = vbTextCompare
    Set physicalNameByKey = CreateObject("Scripting.Dictionary")
    physicalNameByKey.CompareMode = vbTextCompare
    Set querySqlByName = CreateObject("Scripting.Dictionary")
    querySqlByName.CompareMode = vbTextCompare
    For Each tdf In db.TableDefs
        PumpSetupResponsiveness
        If Left$(tdf.Name, 1) <> "~" And StrComp(Left$(tdf.Name, 4), "MSys", vbTextCompare) <> 0 Then
            objectKey = UCase$(tdf.Name)
            objectUpdated = NonFormReferenceDaoPropertyDate(tdf, "LastUpdated", propertyFound)
            If propertyFound = False Then objectUpdated = 0
            physicalUpdatedByName(objectKey) = objectUpdated
            physicalTypeByName(objectKey) = "TABLE"
            physicalNameByKey(objectKey) = tdf.Name
        End If
    Next tdf
    For Each qdf In db.QueryDefs
        PumpSetupResponsiveness
        If Left$(qdf.Name, 1) <> "~" Then
            objectKey = UCase$(qdf.Name)
            objectUpdated = NonFormReferenceDaoPropertyDate(qdf, "LastUpdated", propertyFound)
            If propertyFound = False Then objectUpdated = 0
            physicalUpdatedByName(objectKey) = objectUpdated
            physicalTypeByName(objectKey) = "QUERY"
            physicalNameByKey(objectKey) = qdf.Name
            querySqlByName(objectKey) = qdf.sql
        End If
    Next qdf
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "NonFormReferenceBuildPhysicalDefinitionMaps", "Could not build the Database Reference physical-definition freshness map." & vbCrLf & Err.Description
End Sub

Private Function NonFormReferenceCanReuseQueryFieldEvidence(ByVal qdf As DAO.queryDef, ByVal parentRs As DAO.Recordset, ByVal fieldRs As DAO.Recordset, ByVal objectFieldBookmarks As Object, ByVal physicalUpdatedByName As Object, ByVal physicalTypeByName As Object, ByVal physicalNameByKey As Object, ByVal querySqlByName As Object, ByVal objectName As String, ByRef evidenceFieldCount As Long, ByRef fieldMissingDescription As Boolean) As Boolean
    On Error GoTo CannotReuse
    Dim parentFingerprint As String
    Dim parentVerifiedOn As Date
    Dim visiting As Object
    Dim freshnessCache As Object
    If parentRs Is Nothing Then Exit Function
    If parentRs.EOF Then Exit Function
    parentFingerprint = Trim$(CStr(Nz(parentRs.fields(FIELD_MANAGED_OBJECT_FIELD_SCHEMA_FINGERPRINT).value, vbNullString)))
    If Len(parentFingerprint) = 0 Or IsNull(parentRs.fields(FIELD_MANAGED_OBJECT_FIELD_SCHEMA_VERIFIED_ON).value) Then Exit Function
    parentVerifiedOn = CDate(parentRs.fields(FIELD_MANAGED_OBJECT_FIELD_SCHEMA_VERIFIED_ON).value)
    ' INTENTIONAL KEYED LOOKUP: visiting, freshnessCache provides keyed membership/lookup over source rows or state already read by NonFormReferenceCanReuseQueryFieldEvidence; it is a derived access path, not a
    ' second authority.
    Set visiting = CreateObject("Scripting.Dictionary")
    visiting.CompareMode = vbTextCompare
    Set freshnessCache = CreateObject("Scripting.Dictionary")
    freshnessCache.CompareMode = vbTextCompare
    If NonFormReferenceDefinitionTreeCurrent(objectName, parentVerifiedOn, physicalUpdatedByName, physicalTypeByName, physicalNameByKey, querySqlByName, visiting, freshnessCache) = False Then Exit Function
    If NonFormReferenceStoredFieldEvidenceMatchesParent(fieldRs, objectFieldBookmarks, objectName, "QUERY|" & objectName & "|Type=" & CStr(qdf.Type), parentFingerprint, evidenceFieldCount, fieldMissingDescription) = False Then Exit Function
    NonFormReferenceCanReuseQueryFieldEvidence = True
    Exit Function
CannotReuse:
    NonFormReferenceCanReuseQueryFieldEvidence = False
End Function

Private Function NonFormReferenceDefinitionTreeCurrent(ByVal objectName As String, ByVal verifiedOn As Date, ByVal physicalUpdatedByName As Object, ByVal physicalTypeByName As Object, ByVal physicalNameByKey As Object, ByVal querySqlByName As Object, ByVal visiting As Object, ByVal freshnessCache As Object) As Boolean
    On Error GoTo NotCurrent
    Dim objectKey As String
    Dim cacheKey As String
    Dim objectType As String
    Dim objectUpdated As Date
    Dim sqlText As String
    Dim dependencyKey As Variant
    Dim dependencyCount As Long
    Dim hasSourceClause As Boolean
    objectKey = UCase$(objectName)
    If Len(objectKey) = 0 Then Exit Function
    cacheKey = objectKey & "|" & CStr(CDbl(verifiedOn))
    If freshnessCache.Exists(cacheKey) Then
        NonFormReferenceDefinitionTreeCurrent = CBool(freshnessCache(cacheKey))
        Exit Function
    End If
    If physicalUpdatedByName Is Nothing Or physicalTypeByName Is Nothing Or physicalNameByKey Is Nothing Then Exit Function
    If physicalUpdatedByName.Exists(objectKey) = False Or physicalTypeByName.Exists(objectKey) = False Then Exit Function
    objectUpdated = CDate(physicalUpdatedByName(objectKey))
    If objectUpdated = 0 Or objectUpdated > verifiedOn Then GoTo NotCurrent
    objectType = UCase$(CStr(physicalTypeByName(objectKey)))
    If objectType = "TABLE" Then
        freshnessCache(cacheKey) = True
        NonFormReferenceDefinitionTreeCurrent = True
        Exit Function
    End If
    If objectType <> "QUERY" Then GoTo NotCurrent
    If querySqlByName Is Nothing Then GoTo NotCurrent
    If querySqlByName.Exists(objectKey) = False Then GoTo NotCurrent
    If visiting.Exists(objectKey) Then GoTo NotCurrent
    visiting.Add objectKey, True
    sqlText = CStr(querySqlByName(objectKey))
    hasSourceClause = NonFormReferenceSqlHasSourceClause(sqlText)
    For Each dependencyKey In physicalUpdatedByName.keys
        If StrComp(CStr(dependencyKey), objectKey, vbTextCompare) <> 0 Then
            If physicalNameByKey.Exists(CStr(dependencyKey)) Then
                If NonFormReferenceSqlMentionsObject(sqlText, CStr(physicalNameByKey(CStr(dependencyKey)))) Then
                    dependencyCount = dependencyCount + 1
                    If NonFormReferenceDefinitionTreeCurrent(CStr(physicalNameByKey(CStr(dependencyKey))), verifiedOn, physicalUpdatedByName, physicalTypeByName, physicalNameByKey, querySqlByName, visiting, freshnessCache) = False Then GoTo NotCurrent
                End If
            End If
        End If
    Next dependencyKey
    If hasSourceClause And dependencyCount = 0 Then GoTo NotCurrent
    visiting.Remove objectKey
    freshnessCache(cacheKey) = True
    NonFormReferenceDefinitionTreeCurrent = True
    Exit Function
NotCurrent:
    On Error Resume Next
    If Not visiting Is Nothing Then
        If visiting.Exists(objectKey) Then visiting.Remove objectKey
    End If
    If Not freshnessCache Is Nothing Then
        If Len(cacheKey) > 0 Then freshnessCache(cacheKey) = False
    End If
    On Error GoTo 0
End Function

Private Function NonFormReferenceSqlHasSourceClause(ByVal sqlText As String) As Boolean
    Dim normalizedSql As String
    normalizedSql = UCase$(sqlText)
    normalizedSql = Replace(normalizedSql, vbCr, " ")
    normalizedSql = Replace(normalizedSql, vbLf, " ")
    normalizedSql = Replace(normalizedSql, vbTab, " ")
    Do While InStr(1, normalizedSql, "  ", vbBinaryCompare) > 0
        normalizedSql = Replace(normalizedSql, "  ", " ")
    Loop
    NonFormReferenceSqlHasSourceClause = (InStr(1, normalizedSql, " FROM ", vbBinaryCompare) > 0 Or InStr(1, normalizedSql, " JOIN ", vbBinaryCompare) > 0)
End Function

' Purpose: Returns True when the supplied text is a simple VBA identifier made only of identifier characters.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   textValue (Input; String): Source text being normalized, parsed, compared, bounded, or converted by this helper; the function does not reinterpret unrelated caller state.
' Returns: Boolean: True when the supplied text is a simple VBA identifier made only of identifier characters. False otherwise.
Public Function NonFormReferenceIsSimpleIdentifier(ByVal textValue As String) As Boolean
    'Returns True when the supplied text is a simple VBA identifier made only of identifier characters.
    Dim position As Long
    Dim currentCharacter As String
    If Len(textValue) = 0 Then Exit Function
    For position = 1 To Len(textValue)
        currentCharacter = Mid$(textValue, position, 1)
        If SharedVba_IsIdentifierChar(currentCharacter) = False Then Exit Function
    Next position
    NonFormReferenceIsSimpleIdentifier = True
End Function

Private Function NonFormReferenceSqlMentionsObject(ByVal sqlText As String, ByVal objectName As String) As Boolean
    Dim bracketedName As String
    Dim normalizedSql As String
    Dim normalizedName As String
    Dim searchPosition As Long
    Dim matchPosition As Long
    Dim beforeCharacter As String
    Dim afterCharacter As String
    If Len(objectName) = 0 Then Exit Function
    bracketedName = "[" & Replace(objectName, "]", "]]") & "]"
    If InStr(1, sqlText, bracketedName, vbTextCompare) > 0 Then
        NonFormReferenceSqlMentionsObject = True
        Exit Function
    End If
    If NonFormReferenceIsSimpleIdentifier(objectName) = False Then Exit Function
    normalizedSql = UCase$(sqlText)
    normalizedName = UCase$(objectName)
    searchPosition = 1
    Do
        matchPosition = InStr(searchPosition, normalizedSql, normalizedName, vbBinaryCompare)
        If matchPosition = 0 Then Exit Do
        beforeCharacter = vbNullString
        afterCharacter = vbNullString
        If matchPosition > 1 Then beforeCharacter = Mid$(normalizedSql, matchPosition - 1, 1)
        If matchPosition + Len(normalizedName) <= Len(normalizedSql) Then afterCharacter = Mid$(normalizedSql, matchPosition + Len(normalizedName), 1)
        If SharedVba_IsIdentifierChar(beforeCharacter) = False And SharedVba_IsIdentifierChar(afterCharacter) = False Then
            NonFormReferenceSqlMentionsObject = True
            Exit Function
        End If
        searchPosition = matchPosition + Len(normalizedName)
    Loop
End Function

Private Function NonFormReferenceStoredFieldEvidenceMatchesParent(ByVal fieldRs As DAO.Recordset, ByVal objectFieldBookmarks As Object, ByVal objectName As String, ByVal signaturePrefix As String, ByVal parentFingerprint As String, ByRef evidenceFieldCount As Long, ByRef fieldMissingDescription As Boolean) As Boolean
    On Error GoTo CannotReuse
    Dim fieldStoredFingerprint As String
    Dim fieldCalculatedFingerprint As String
    Dim fieldStructuralText As String
    Dim fieldName As String
    Dim dataTypeName As String
    Dim daoType As Long
    Dim ordinalValue As Long
    Dim rowID As Long
    Dim keyText As String
    Dim keys As Variant
    Dim keyValue As Variant
    Dim rowBookmarkKey As Variant
    Dim pieces As Object
    Dim seenNames As Object
    Dim storedFingerprint As String
    ' INTENTIONAL KEYED LOOKUP: pieces, seenNames provides keyed membership/lookup over source rows or state already read by NonFormReferenceStoredFieldEvidenceMatchesParent; it is a derived access path, not a
    ' second authority.
    Set pieces = CreateObject("Scripting.Dictionary")
    pieces.CompareMode = vbBinaryCompare
    Set seenNames = CreateObject("Scripting.Dictionary")
    seenNames.CompareMode = vbTextCompare
    If objectFieldBookmarks Is Nothing Then Exit Function
    For Each rowBookmarkKey In objectFieldBookmarks.keys
        fieldRs.Bookmark = objectFieldBookmarks(CStr(rowBookmarkKey))
        PumpSetupResponsiveness
        If CBool(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_IS_ACTIVE).value, False)) Then
            fieldName = Trim$(CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_NAME).value, vbNullString)))
            If Len(fieldName) = 0 Or seenNames.Exists(fieldName) Then Exit Function
            If Len(Trim$(CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_DESCRIPTION_LONG).value, vbNullString)))) = 0 Then fieldMissingDescription = True
            fieldStoredFingerprint = Trim$(CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_SCHEMA_FINGERPRINT).value, vbNullString)))
            If Len(fieldStoredFingerprint) = 0 Or IsNull(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_SCHEMA_VERIFIED_ON).value) Then Exit Function
            dataTypeName = CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_DATA_TYPE).value, vbNullString))
            daoType = NonFormReferenceDaoTypeFromName(dataTypeName)
            If daoType = -32768 Then Exit Function
            ordinalValue = CLng(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_ORDINAL).value, -1))
            If ordinalValue < 0 Then Exit Function
            fieldStructuralText = objectName & "|" & fieldName & "|" & CStr(ordinalValue) & "|" & CStr(daoType) & "|" & CStr(CLng(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_ACTUAL_SIZE).value, 0))) & "|" & CStr(CBool(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_IS_REQUIRED).value, False))) & "|" & CStr(CBool(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_ALLOWS_ZERO_LENGTH).value, False))) & "|" & CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_DEFAULT_VALUE).value, vbNullString)) & "|" & CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALIDATION_RULE).value, vbNullString)) & "|" & CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_LOOKUP_ROW_SOURCE_TYPE).value, vbNullString)) & "|" & CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_LOOKUP_ROW_SOURCE).value, vbNullString))
            fieldCalculatedFingerprint = FingerprintFromPayload(fieldStructuralText)
            If StrComp(fieldCalculatedFingerprint, fieldStoredFingerprint, vbBinaryCompare) <> 0 Then Exit Function
            rowID = CLng(Nz(fieldRs.fields(FIELD_ID).value, 0))
            keyText = Format$(ordinalValue, "000000") & "|" & UCase$(fieldName) & "|" & Format$(rowID, "0000000000")
            pieces.Add keyText, "|" & fieldName & ":" & CStr(daoType) & ":" & CStr(CLng(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_ACTUAL_SIZE).value, 0)))
            seenNames.Add fieldName, True
        End If
    Next rowBookmarkKey
    evidenceFieldCount = pieces.count
    If evidenceFieldCount = 0 Then Exit Function
    keys = SortedDictionaryKeys(pieces)
    For Each keyValue In keys
        signaturePrefix = signaturePrefix & CStr(pieces(CStr(keyValue)))
    Next keyValue
    storedFingerprint = FingerprintFromPayload(signaturePrefix)
    If StrComp(storedFingerprint, parentFingerprint, vbBinaryCompare) <> 0 Then Exit Function
    NonFormReferenceStoredFieldEvidenceMatchesParent = True
    Exit Function
CannotReuse:
    NonFormReferenceStoredFieldEvidenceMatchesParent = False
End Function

Private Function NonFormReferenceDaoPropertyDate(ByVal target As Object, ByVal propertyName As String, ByRef propertyFound As Boolean) As Date
    On Error GoTo MissingProperty
    propertyFound = True
    NonFormReferenceDaoPropertyDate = CDate(target.properties(propertyName).value)
    Exit Function
MissingProperty:
    propertyFound = False
    NonFormReferenceDaoPropertyDate = 0
End Function

Private Function NonFormReferenceDaoTypeFromName(ByVal typeName As String) As Long
    Select Case UCase$(Trim$(typeName))
        Case "BOOLEAN": NonFormReferenceDaoTypeFromName = dbBoolean
        Case "BYTE": NonFormReferenceDaoTypeFromName = dbByte
        Case "INTEGER": NonFormReferenceDaoTypeFromName = dbInteger
        Case "LONG INTEGER": NonFormReferenceDaoTypeFromName = dbLong
        Case "SINGLE": NonFormReferenceDaoTypeFromName = dbSingle
        Case "DOUBLE": NonFormReferenceDaoTypeFromName = dbDouble
        Case "CURRENCY": NonFormReferenceDaoTypeFromName = dbCurrency
        Case "DATE/TIME": NonFormReferenceDaoTypeFromName = dbDate
        Case "SHORT TEXT": NonFormReferenceDaoTypeFromName = dbText
        Case "LONG BINARY": NonFormReferenceDaoTypeFromName = dbLongBinary
        Case "LONG TEXT": NonFormReferenceDaoTypeFromName = dbMemo
        Case "GUID": NonFormReferenceDaoTypeFromName = dbGUID
        Case "DECIMAL": NonFormReferenceDaoTypeFromName = dbDecimal
        Case Else: NonFormReferenceDaoTypeFromName = -32768
    End Select
End Function

Private Function NonFormReferenceBuildIndexFlagMap(ByVal tdf As DAO.tableDef) As Object
    On Error GoTo ErrHandler
    Dim result As Object
    Dim idx As DAO.index
    Dim idxField As DAO.Field
    Dim fieldName As String
    Dim flags As Long
    ' INTENTIONAL KEYED LOOKUP: result provides keyed membership/lookup over source rows or state already read by NonFormReferenceBuildIndexFlagMap; it is a derived access path, not a second authority.
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    For Each idx In tdf.Indexes
        PumpSetupResponsiveness
        For Each idxField In idx.fields
            fieldName = idxField.Name
            flags = 0
            If result.Exists(fieldName) Then flags = CLng(result(fieldName))
            flags = flags Or 1
            If idx.Primary Then flags = flags Or 2
            If idx.Unique Then flags = flags Or 4
            result(fieldName) = flags
        Next idxField
    Next idx
    Set NonFormReferenceBuildIndexFlagMap = result
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "NonFormReferenceBuildIndexFlagMap", "Could not build the per-table Database Reference index-flag map." & vbCrLf & "Table: " & tdf.Name & vbCrLf & "Description: " & Err.Description
End Function

Private Sub NonFormReferenceOpenSharedFieldValueIndex(ByVal db As DAO.Database, ByRef valueRs As DAO.Recordset, ByRef activeValueFieldIDs As Object, ByRef valuePrimaryBookmarkByIdentity As Object, ByRef valuePrimaryIDByIdentity As Object, ByRef valueBookmarksByFieldID As Object, ByRef indexedValueRowCount As Long)
    On Error GoTo ErrHandler
    Dim fieldID As Long
    Dim fieldKey As String
    Dim storedValue As String
    Dim identityKey As String
    Dim rowBookmarks As Object
    ' INTENTIONAL KEYED LOOKUP: activeValueFieldIDs, valuePrimaryBookmarkByIdentity, valuePrimaryIDByIdentity, valueBookmarksByFieldID provides keyed membership/lookup over source rows or state already read by
    ' NonFormReferenceOpenSharedFieldValueIndex; it is a derived access path, not a second authority.
    Set activeValueFieldIDs = CreateObject("Scripting.Dictionary")
    activeValueFieldIDs.CompareMode = vbTextCompare
    Set valuePrimaryBookmarkByIdentity = CreateObject("Scripting.Dictionary")
    valuePrimaryBookmarkByIdentity.CompareMode = vbTextCompare
    Set valuePrimaryIDByIdentity = CreateObject("Scripting.Dictionary")
    valuePrimaryIDByIdentity.CompareMode = vbTextCompare
    Set valueBookmarksByFieldID = CreateObject("Scripting.Dictionary")
    valueBookmarksByFieldID.CompareMode = vbTextCompare
    Set valueRs = db.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFieldValues") & "];", dbOpenDynaset)
    Do While valueRs.EOF = False
        PumpSetupResponsiveness
        indexedValueRowCount = indexedValueRowCount + 1
        fieldID = CLng(Nz(valueRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_MANAGED_OBJECT_FIELD_ID).value, 0))
        If fieldID > 0 Then
            fieldKey = CStr(fieldID)
            If valueBookmarksByFieldID.Exists(fieldKey) Then
                Set rowBookmarks = valueBookmarksByFieldID(fieldKey)
            Else
                ' INTENTIONAL KEYED LOOKUP: rowBookmarks provides keyed membership/lookup over source rows or state already read by NonFormReferenceOpenSharedFieldValueIndex; it is a derived access path, not a
                ' second authority.
                Set rowBookmarks = CreateObject("Scripting.Dictionary")
                rowBookmarks.CompareMode = vbBinaryCompare
                valueBookmarksByFieldID.Add fieldKey, rowBookmarks
            End If
            rowBookmarks.Add CStr(CLng(Nz(valueRs.fields(FIELD_ID).value, 0))), valueRs.Bookmark
            If CBool(Nz(valueRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_IS_ACTIVE).value, False)) Then activeValueFieldIDs(fieldKey) = True
            storedValue = CStr(Nz(valueRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_STORED_VALUE).value, vbNullString))
            If Len(storedValue) > 0 Then
                identityKey = fieldKey & "|" & storedValue
                If valuePrimaryBookmarkByIdentity.Exists(identityKey) Then
                    If CLng(Nz(valueRs.fields(FIELD_ID).value, 0)) < CLng(valuePrimaryIDByIdentity(identityKey)) Then
                        valuePrimaryBookmarkByIdentity(identityKey) = valueRs.Bookmark
                        valuePrimaryIDByIdentity(identityKey) = CLng(Nz(valueRs.fields(FIELD_ID).value, 0))
                    End If
                Else
                    valuePrimaryBookmarkByIdentity.Add identityKey, valueRs.Bookmark
                    valuePrimaryIDByIdentity.Add identityKey, CLng(Nz(valueRs.fields(FIELD_ID).value, 0))
                End If
            End If
        End If
        valueRs.MoveNext
    Loop
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not valueRs Is Nothing Then valueRs.Close
    Set valueRs = Nothing
    On Error GoTo 0
    Err.Raise savedErrNumber, "NonFormReferenceOpenSharedFieldValueIndex", "Could not open/index the shared MAC_Fields_Vals synchronization recordset." & vbCrLf & savedErrDescription
End Sub

Private Sub NonFormReferenceAddDesiredFieldValue(ByVal desired As Object, ByVal storedValue As String, ByVal displayValue As String, ByVal valueDescription As String, ByVal SortOrder As Long, ByVal isDefault As Boolean, ByVal documentationSource As String)
    storedValue = Left$(storedValue, 255)
    If Len(Trim$(storedValue)) = 0 Then Exit Sub
    displayValue = Left$(displayValue, 255)
    If Len(Trim$(displayValue)) = 0 Then displayValue = storedValue
    desired(storedValue) = Array(displayValue, valueDescription, SortOrder, isDefault, documentationSource)
End Sub

Private Sub NonFormReferenceApplyDesiredFieldValues(ByVal valueRs As DAO.Recordset, ByVal activeValueFieldIDs As Object, ByVal valuePrimaryBookmarkByIdentity As Object, ByVal valuePrimaryIDByIdentity As Object, ByVal valueBookmarksByFieldID As Object, ByVal managedObjectFieldID As Long, ByVal desired As Object, ByVal coldBulkMode As Boolean, ByRef synchronizedValueCount As Long, ByRef changedValueCount As Long)
    On Error GoTo ErrHandler
    Dim fieldKey As String
    Dim identityKey As String
    Dim rowBookmarks As Object
    Dim key As Variant
    Dim rowData As Variant
    Dim storedValue As String
    Dim rowID As Long
    Dim rowChanged As Boolean
    Dim currentDescription As String
    Dim expectedStatus As String
    Dim rowBookmarkKey As Variant
    If valueRs Is Nothing Then Err.Raise vbObjectError + 9578, "NonFormReferenceApplyDesiredFieldValues", "The shared MAC_Fields_Vals synchronization recordset is unavailable."
    fieldKey = CStr(managedObjectFieldID)
    For Each key In desired.keys
        synchronizedValueCount = synchronizedValueCount + 1
        rowData = desired(CStr(key))
        identityKey = fieldKey & "|" & CStr(key)
        rowChanged = False
        If valuePrimaryBookmarkByIdentity.Exists(identityKey) Then
            valueRs.Bookmark = valuePrimaryBookmarkByIdentity(identityKey)
            currentDescription = CStr(Nz(valueRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_DESCRIPTION).value, vbNullString))
            expectedStatus = IIf(Len(currentDescription) > 0 Or Len(CStr(rowData(1))) > 0, DOCUMENTATION_STATUS_COMPLETE, DOCUMENTATION_STATUS_MISSING_DESCRIPTION)
            If StrComp(CStr(Nz(valueRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_DISPLAY_VALUE).value, vbNullString)), CStr(rowData(0)), vbBinaryCompare) <> 0 Then rowChanged = True
            If Len(currentDescription) = 0 And Len(CStr(rowData(1))) > 0 Then rowChanged = True
            If CLng(Nz(valueRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_SORT_ORDER).value, 0)) <> CLng(rowData(2)) Then rowChanged = True
            If CBool(Nz(valueRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_IS_DEFAULT).value, False)) <> CBool(rowData(3)) Then rowChanged = True
            If CBool(Nz(valueRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_IS_ACTIVE).value, False)) = False Then rowChanged = True
            If StrComp(CStr(Nz(valueRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_DOCUMENTATION_STATUS).value, vbNullString)), expectedStatus, vbTextCompare) <> 0 Then rowChanged = True
            If StrComp(CStr(Nz(valueRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_DOCUMENTATION_SOURCE).value, vbNullString)), CStr(rowData(4)), vbBinaryCompare) <> 0 Then rowChanged = True
            If rowChanged Then valueRs.Edit
        Else
            valueRs.AddNew
            valueRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_MANAGED_OBJECT_FIELD_ID).value = managedObjectFieldID
            valueRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_STORED_VALUE).value = CStr(key)
            currentDescription = vbNullString
            expectedStatus = IIf(Len(CStr(rowData(1))) > 0, DOCUMENTATION_STATUS_COMPLETE, DOCUMENTATION_STATUS_MISSING_DESCRIPTION)
            rowChanged = True
        End If
        If rowChanged Then
            valueRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_DISPLAY_VALUE).value = CStr(rowData(0))
            If Len(currentDescription) = 0 And Len(CStr(rowData(1))) > 0 Then valueRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_DESCRIPTION).value = CStr(rowData(1))
            valueRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_SORT_ORDER).value = CLng(rowData(2))
            valueRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_IS_DEFAULT).value = CBool(rowData(3))
            valueRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_IS_ACTIVE).value = True
            valueRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_DOCUMENTATION_STATUS).value = expectedStatus
            valueRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_DOCUMENTATION_SOURCE).value = CStr(rowData(4))
            valueRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_DOCUMENTATION_UPDATED_ON).value = Now()
            valueRs.Update
            changedValueCount = changedValueCount + 1
            If valuePrimaryBookmarkByIdentity.Exists(identityKey) = False And coldBulkMode = False Then
                valueRs.Bookmark = valueRs.LastModified
                rowID = CLng(Nz(valueRs.fields(FIELD_ID).value, 0))
                valuePrimaryBookmarkByIdentity.Add identityKey, valueRs.Bookmark
                valuePrimaryIDByIdentity.Add identityKey, rowID
                If valueBookmarksByFieldID.Exists(fieldKey) Then
                    Set rowBookmarks = valueBookmarksByFieldID(fieldKey)
                Else
                    ' INTENTIONAL KEYED LOOKUP: rowBookmarks provides keyed membership/lookup over source rows or state already read by NonFormReferenceApplyDesiredFieldValues; it is a derived access path, not a
                    ' second authority.
                    Set rowBookmarks = CreateObject("Scripting.Dictionary")
                    rowBookmarks.CompareMode = vbBinaryCompare
                    valueBookmarksByFieldID.Add fieldKey, rowBookmarks
                End If
                rowBookmarks.Add CStr(rowID), valueRs.Bookmark
            End If
        End If
    Next key
    If valueBookmarksByFieldID.Exists(fieldKey) Then
        Set rowBookmarks = valueBookmarksByFieldID(fieldKey)
        For Each rowBookmarkKey In rowBookmarks.keys
            valueRs.Bookmark = rowBookmarks(CStr(rowBookmarkKey))
            storedValue = CStr(Nz(valueRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_STORED_VALUE).value, vbNullString))
            rowID = CLng(Nz(valueRs.fields(FIELD_ID).value, 0))
            If CBool(Nz(valueRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_IS_ACTIVE).value, False)) Then
                identityKey = fieldKey & "|" & storedValue
                If desired.Exists(storedValue) = False Then
                    valueRs.Edit
                    valueRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_IS_ACTIVE).value = False
                    valueRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_DOCUMENTATION_UPDATED_ON).value = Now()
                    valueRs.Update
                    changedValueCount = changedValueCount + 1
                ElseIf valuePrimaryIDByIdentity.Exists(identityKey) Then
                    If rowID <> CLng(valuePrimaryIDByIdentity(identityKey)) Then
                        valueRs.Edit
                        valueRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_IS_ACTIVE).value = False
                        valueRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_DOCUMENTATION_UPDATED_ON).value = Now()
                        valueRs.Update
                        changedValueCount = changedValueCount + 1
                    End If
                End If
            End If
        Next rowBookmarkKey
    End If
    If desired.count > 0 Then
        activeValueFieldIDs(fieldKey) = True
    ElseIf activeValueFieldIDs.Exists(fieldKey) Then
        activeValueFieldIDs.Remove fieldKey
    End If
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "NonFormReferenceApplyDesiredFieldValues", "Shared controlled-value Database Reference synchronization failed." & vbCrLf & "ManagedObjectFieldID: " & CStr(managedObjectFieldID) & vbCrLf & Err.Description
End Sub

Private Function NonFormReferenceValueIsDefault(ByVal storedValue As String, ByVal DefaultValue As String) As Boolean
    Dim normalizedStored As String
    Dim normalizedDefault As String
    normalizedStored = UCase$(Trim$(storedValue))
    normalizedDefault = UCase$(Replace(Replace(Trim$(DefaultValue), "=", vbNullString), Chr$(34), vbNullString))
    If normalizedStored = "FALSE" Then
        NonFormReferenceValueIsDefault = (normalizedDefault = "FALSE" Or normalizedDefault = "0")
    ElseIf normalizedStored = "TRUE" Then
        NonFormReferenceValueIsDefault = (normalizedDefault = "TRUE" Or normalizedDefault = "-1")
    Else
        NonFormReferenceValueIsDefault = (StrComp(normalizedStored, normalizedDefault, vbTextCompare) = 0)
    End If
End Function

Private Sub NonFormReferenceReadFieldProperties(ByVal fld As DAO.Field, ByVal isTableField As Boolean, ByVal categoryCode As String, ByRef descriptionText As String, ByRef displayName As String, ByRef rowSourceType As String, ByRef rowSource As String, ByRef sourceObjectName As String, ByRef sourceFieldName As String, ByRef isRequired As Boolean, ByRef allowsZeroLength As Boolean, ByRef DefaultValue As String, ByRef validationRule As String, ByRef validationText As String, ByRef lookupBoundColumn As Long, ByRef lookupColumnCount As Long, ByRef columnWidths As String)
    'Table fields own custom DAO metadata. Query fields own only their exposed identity/type/order; source-table metadata is inherited separately without enumerating QueryDef.Field.Properties.
    Dim prp As DAO.Property
    Dim propertyName As String
    On Error Resume Next
    If isTableField = False Then
        sourceObjectName = CStr(Nz(fld.sourceTable, vbNullString))
        sourceFieldName = CStr(Nz(fld.sourceField, vbNullString))
        Err.Clear
        On Error GoTo 0
        mB5DatabaseReferenceQueryMinimalPropertyCount = mB5DatabaseReferenceQueryMinimalPropertyCount + 1
        If StrComp(Trim$(categoryCode), ADCAT_CATEGORY_FORM_QUERY, vbTextCompare) = 0 Then mB5DatabaseReferenceFormDesignMinimalPropertyCount = mB5DatabaseReferenceFormDesignMinimalPropertyCount + 1
        Exit Sub
    End If
    isRequired = CBool(Nz(fld.Required, False))
    allowsZeroLength = CBool(Nz(fld.allowZeroLength, False))
    DefaultValue = CStr(Nz(fld.DefaultValue, vbNullString))
    validationRule = CStr(Nz(fld.validationRule, vbNullString))
    validationText = CStr(Nz(fld.validationText, vbNullString))
    Err.Clear
    For Each prp In fld.properties
        propertyName = UCase$(prp.Name)
        Select Case propertyName
            Case UCase$(DAO_DESCRIPTION_PROPERTY_NAME)
                descriptionText = CStr(Nz(prp.value, vbNullString))
            Case "CAPTION"
                displayName = CStr(Nz(prp.value, vbNullString))
            Case "REQUIRED"
                isRequired = CBool(Nz(prp.value, False))
            Case "ALLOWZEROLENGTH"
                allowsZeroLength = CBool(Nz(prp.value, False))
            Case "DEFAULTVALUE"
                DefaultValue = CStr(Nz(prp.value, vbNullString))
            Case "VALIDATIONRULE"
                validationRule = CStr(Nz(prp.value, vbNullString))
            Case "VALIDATIONTEXT"
                validationText = CStr(Nz(prp.value, vbNullString))
            Case "ROWSOURCETYPE"
                rowSourceType = CStr(Nz(prp.value, vbNullString))
            Case "ROWSOURCE"
                rowSource = CStr(Nz(prp.value, vbNullString))
            Case "BOUNDCOLUMN"
                lookupBoundColumn = CLng(Nz(prp.value, 0))
            Case "COLUMNCOUNT"
                lookupColumnCount = CLng(Nz(prp.value, 0))
            Case "COLUMNWIDTHS"
                columnWidths = CStr(Nz(prp.value, vbNullString))
        End Select
        Err.Clear
    Next prp
    On Error GoTo 0
End Sub

Private Function NonFormReferenceDaoTypeName(ByVal daoType As Long) As String
    Select Case daoType
        Case dbBoolean: NonFormReferenceDaoTypeName = "Boolean"
        Case dbByte: NonFormReferenceDaoTypeName = "Byte"
        Case dbInteger: NonFormReferenceDaoTypeName = "Integer"
        Case dbLong: NonFormReferenceDaoTypeName = "Long Integer"
        Case dbSingle: NonFormReferenceDaoTypeName = "Single"
        Case dbDouble: NonFormReferenceDaoTypeName = "Double"
        Case dbCurrency: NonFormReferenceDaoTypeName = "Currency"
        Case dbDate: NonFormReferenceDaoTypeName = "Date/Time"
        Case dbText: NonFormReferenceDaoTypeName = "Short Text"
        Case dbLongBinary: NonFormReferenceDaoTypeName = "Long Binary"
        Case dbMemo: NonFormReferenceDaoTypeName = "Long Text"
        Case dbGUID: NonFormReferenceDaoTypeName = "GUID"
        Case dbDecimal: NonFormReferenceDaoTypeName = "Decimal"
        Case Else: NonFormReferenceDaoTypeName = "DAO Type " & CStr(daoType)
    End Select
End Function

' Purpose: Performs the read-only Database Reference postcondition validation used when C1-C3 require proof that completed B5 non-form table/query evidence remains current.
' Public role: Inter-module prerequisite-verification API called by Catalog; it validates B5 output without creating, repairing, or resynchronizing Database Reference data.
' Parameters:
'   db (Input; DAO.Database): CodeDb database containing the completed B5 managed-object, field, and controlled-value evidence to validate.
' Returns: None (Sub). Successful return means the Database Reference invariants passed; failures are raised to the owning prerequisite check.
Public Sub ValidateNonFormDatabaseReferenceFieldIndex(ByVal db As DAO.Database)
    On Error GoTo ErrHandler
    Dim parentRs As DAO.Recordset
    Dim detailText As String
    Set parentRs = db.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID) & "] WHERE Nz([" & FIELD_MANAGED_OBJECT_STATUS & "]," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & " AND UCase(Trim(Nz(" & SQL_FIELD_OBJECT_TYPE & ",''))) IN ('TABLE','QUERY') ORDER BY " & SQL_FIELD_OBJECT_TYPE & "," & SQL_FIELD_OBJECT_NAME & "," & SQL_FIELD_ID & ";", dbOpenSnapshot)
    ValidateNonFormDatabaseReferenceFieldIndexFromSnapshot db, parentRs, detailText
    parentRs.Close
    Set parentRs = Nothing
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not parentRs Is Nothing Then parentRs.Close
    On Error GoTo 0
    Err.Raise savedErrNumber, "ValidateNonFormDatabaseReferenceFieldIndex", savedErrDescription
End Sub

Private Sub ValidateNonFormDatabaseReferenceFieldIndexFromSnapshot(ByVal db As DAO.Database, ByVal parentRs As DAO.Recordset, ByRef detailText As String)
    On Error GoTo ErrHandler
    Dim fieldRs As DAO.Recordset
    Dim valueRs As DAO.Recordset
    Dim parentIDs As Object
    Dim parentObjectTypeByID As Object
    Dim fieldByID As Object
    Dim activeFieldKeys As Object
    Dim duplicateFieldKeys As Object
    Dim activeValueKeys As Object
    Dim duplicateValueKeys As Object
    Dim booleanFieldState As Object
    Dim valueListFieldHasValue As Object
    Dim invalidObjectCount As Long
    Dim invalidObjectDocumentationCount As Long
    Dim invalidFieldCount As Long
    Dim duplicateFieldCount As Long
    Dim invalidDocumentationCount As Long
    Dim orphanActiveValueCount As Long
    Dim duplicateActiveValueCount As Long
    Dim invalidActiveValueDocumentationCount As Long
    Dim missingBooleanFalseCount As Long
    Dim missingBooleanTrueCount As Long
    Dim invalidBooleanExtraCount As Long
    Dim valueListWithoutValuesCount As Long
    Dim queryFieldActiveValueCount As Long
    Dim parentID As Long
    Dim fieldID As Long
    Dim fieldName As String
    Dim fieldKey As String
    Dim storedValue As String
    Dim valueKey As String
    Dim dataTypeName As String
    Dim rowSourceType As String
    Dim rowSource As String
    Dim isActiveField As Boolean
    Dim isTableField As Boolean
    Dim fieldData As Variant
    Dim keyValue As Variant
    Dim stateValue As Long
    Dim parentRowCount As Long
    Dim fieldRowCount As Long
    Dim valueRowCount As Long
    Dim operationTimer As Single
    operationTimer = Timer
    ' INTENTIONAL OPERATION SNAPSHOT: parentIDs, booleanFieldState, valueListFieldHasValue captures one coherent source state for this invocation of ValidateNonFormDatabaseReferenceFieldIndexFromSnapshot; validity
    ' ends with the operation rather than becoming a live data authority.
    ' INTENTIONAL KEYED LOOKUP: parentObjectTypeByID, fieldByID, activeFieldKeys, duplicateFieldKeys, activeValueKeys, duplicateValueKeys provides keyed membership/lookup over source rows or state already read by
    ' ValidateNonFormDatabaseReferenceFieldIndexFromSnapshot; it is a derived access path, not a second authority.
    Set parentIDs = CreateObject("Scripting.Dictionary")
    parentIDs.CompareMode = vbBinaryCompare
    Set parentObjectTypeByID = CreateObject("Scripting.Dictionary")
    parentObjectTypeByID.CompareMode = vbBinaryCompare
    Set fieldByID = CreateObject("Scripting.Dictionary")
    fieldByID.CompareMode = vbBinaryCompare
    Set activeFieldKeys = CreateObject("Scripting.Dictionary")
    activeFieldKeys.CompareMode = vbBinaryCompare
    Set duplicateFieldKeys = CreateObject("Scripting.Dictionary")
    duplicateFieldKeys.CompareMode = vbBinaryCompare
    Set activeValueKeys = CreateObject("Scripting.Dictionary")
    activeValueKeys.CompareMode = vbBinaryCompare
    Set duplicateValueKeys = CreateObject("Scripting.Dictionary")
    duplicateValueKeys.CompareMode = vbBinaryCompare
    Set booleanFieldState = CreateObject("Scripting.Dictionary")
    booleanFieldState.CompareMode = vbBinaryCompare
    Set valueListFieldHasValue = CreateObject("Scripting.Dictionary")
    valueListFieldHasValue.CompareMode = vbBinaryCompare
    If parentRs Is Nothing Then Err.Raise vbObjectError + 9569, "ValidateNonFormDatabaseReferenceFieldIndexFromSnapshot", "The Database Reference parent-row snapshot is unavailable."
    If Not (parentRs.BOF And parentRs.EOF) Then
        parentRs.MoveFirst
        Do While parentRs.EOF = False
            PumpSetupResponsiveness
            parentID = CLng(Nz(parentRs.fields(FIELD_ID).value, 0))
            If parentID > 0 Then
                parentIDs(CStr(parentID)) = True
                parentObjectTypeByID(CStr(parentID)) = UCase$(Trim$(CStr(Nz(parentRs.fields(FIELD_OBJECT_TYPE).value, vbNullString))))
                parentRowCount = parentRowCount + 1
                If Len(Trim$(CStr(Nz(parentRs.fields(FIELD_MANAGED_OBJECT_FIELD_SCHEMA_FINGERPRINT).value, vbNullString)))) = 0 Or IsNull(parentRs.fields(FIELD_MANAGED_OBJECT_FIELD_SCHEMA_VERIFIED_ON).value) Then invalidObjectCount = invalidObjectCount + 1
                If CBool(Nz(parentRs.fields(FIELD_IS_DOCUMENTATION_ACTIVE).value, False)) = False Or Len(Trim$(CStr(Nz(parentRs.fields(FIELD_MANAGED_OBJECT_DEPENDENCY_DOCUMENTATION_STATUS).value, vbNullString)))) = 0 Or Len(Trim$(CStr(Nz(parentRs.fields(FIELD_MANAGED_OBJECT_FIELD_DOCUMENTATION_SOURCE).value, vbNullString)))) = 0 Then invalidObjectDocumentationCount = invalidObjectDocumentationCount + 1
            End If
            parentRs.MoveNext
        Loop
    End If
    Set fieldRs = db.OpenRecordset("SELECT " & SQL_FIELD_ID & ",[" & FIELD_MANAGED_OBJECT_ID & "],[" & FIELD_MANAGED_OBJECT_FIELD_NAME & "],[" & FIELD_MANAGED_OBJECT_FIELD_IS_ACTIVE & "],[" & FIELD_MANAGED_OBJECT_FIELD_SCHEMA_FINGERPRINT & "],[" & FIELD_MANAGED_OBJECT_FIELD_SCHEMA_VERIFIED_ON & "],[" & FIELD_MANAGED_OBJECT_FIELD_DOCUMENTATION_STATUS & "],[" & FIELD_MANAGED_OBJECT_FIELD_DOCUMENTATION_SOURCE & "],[" & FIELD_MANAGED_OBJECT_FIELD_DATA_TYPE & "],[" & FIELD_MANAGED_OBJECT_FIELD_LOOKUP_ROW_SOURCE_TYPE & "],[" & FIELD_MANAGED_OBJECT_FIELD_LOOKUP_ROW_SOURCE & "] FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFields") & "] ORDER BY [" & FIELD_MANAGED_OBJECT_ID & "]," & SQL_FIELD_ID & ";", dbOpenSnapshot)
    Do While fieldRs.EOF = False
        PumpSetupResponsiveness
        parentID = CLng(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_ID).value, 0))
        If parentIDs.Exists(CStr(parentID)) Then
            fieldRowCount = fieldRowCount + 1
            fieldID = CLng(Nz(fieldRs.fields(FIELD_ID).value, 0))
            fieldName = Trim$(CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_NAME).value, vbNullString)))
            isActiveField = CBool(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_IS_ACTIVE).value, False))
            dataTypeName = UCase$(Trim$(CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_DATA_TYPE).value, vbNullString))))
            rowSourceType = UCase$(Trim$(CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_LOOKUP_ROW_SOURCE_TYPE).value, vbNullString))))
            rowSource = Trim$(CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_LOOKUP_ROW_SOURCE).value, vbNullString)))
            isTableField = (StrComp(CStr(parentObjectTypeByID(CStr(parentID))), "TABLE", vbBinaryCompare) = 0)
            fieldByID(CStr(fieldID)) = Array(isActiveField, dataTypeName, rowSourceType, rowSource, isTableField)
            If isActiveField Then
                If Len(fieldName) = 0 Or Len(Trim$(CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_SCHEMA_FINGERPRINT).value, vbNullString)))) = 0 Or IsNull(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_SCHEMA_VERIFIED_ON).value) Then invalidFieldCount = invalidFieldCount + 1
                fieldKey = CStr(parentID) & "|" & UCase$(fieldName)
                If activeFieldKeys.Exists(fieldKey) Then
                    If duplicateFieldKeys.Exists(fieldKey) = False Then duplicateFieldKeys.Add fieldKey, True
                Else
                    activeFieldKeys.Add fieldKey, fieldID
                End If
                If Len(Trim$(CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_DOCUMENTATION_STATUS).value, vbNullString)))) = 0 Or Len(Trim$(CStr(Nz(fieldRs.fields(FIELD_MANAGED_OBJECT_FIELD_DOCUMENTATION_SOURCE).value, vbNullString)))) = 0 Then invalidDocumentationCount = invalidDocumentationCount + 1
                If isTableField And StrComp(dataTypeName, "BOOLEAN", vbBinaryCompare) = 0 Then booleanFieldState(CStr(fieldID)) = 0
                If isTableField And StrComp(rowSourceType, "VALUE LIST", vbBinaryCompare) = 0 And Len(rowSource) > 0 Then valueListFieldHasValue(CStr(fieldID)) = False
            End If
        End If
        fieldRs.MoveNext
    Loop
    fieldRs.Close
    Set fieldRs = Nothing
    Set valueRs = db.OpenRecordset("SELECT " & SQL_FIELD_ID & ",[" & FIELD_MANAGED_OBJECT_FIELD_VALUE_MANAGED_OBJECT_FIELD_ID & "],[" & FIELD_MANAGED_OBJECT_FIELD_VALUE_STORED_VALUE & "],[" & FIELD_MANAGED_OBJECT_FIELD_VALUE_IS_ACTIVE & "],[" & FIELD_MANAGED_OBJECT_FIELD_VALUE_DOCUMENTATION_STATUS & "],[" & FIELD_MANAGED_OBJECT_FIELD_VALUE_DOCUMENTATION_SOURCE & "] FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ManagedObjectFieldValues") & "] WHERE [" & FIELD_MANAGED_OBJECT_FIELD_VALUE_IS_ACTIVE & "]=True ORDER BY [" & FIELD_MANAGED_OBJECT_FIELD_VALUE_MANAGED_OBJECT_FIELD_ID & "]," & SQL_FIELD_ID & ";", dbOpenSnapshot)
    Do While valueRs.EOF = False
        PumpSetupResponsiveness
        fieldID = CLng(Nz(valueRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_MANAGED_OBJECT_FIELD_ID).value, 0))
        If fieldByID.Exists(CStr(fieldID)) Then
            valueRowCount = valueRowCount + 1
            fieldData = fieldByID(CStr(fieldID))
            isActiveField = CBool(fieldData(0))
            isTableField = CBool(fieldData(4))
            If isActiveField = False Then
                orphanActiveValueCount = orphanActiveValueCount + 1
            ElseIf isTableField = False Then
                queryFieldActiveValueCount = queryFieldActiveValueCount + 1
            Else
                storedValue = UCase$(Trim$(CStr(Nz(valueRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_STORED_VALUE).value, vbNullString))))
                valueKey = CStr(fieldID) & "|" & storedValue
                If activeValueKeys.Exists(valueKey) Then
                    If duplicateValueKeys.Exists(valueKey) = False Then duplicateValueKeys.Add valueKey, True
                Else
                    activeValueKeys.Add valueKey, True
                End If
                If Len(Trim$(CStr(Nz(valueRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_DOCUMENTATION_STATUS).value, vbNullString)))) = 0 Or Len(Trim$(CStr(Nz(valueRs.fields(FIELD_MANAGED_OBJECT_FIELD_VALUE_DOCUMENTATION_SOURCE).value, vbNullString)))) = 0 Then invalidActiveValueDocumentationCount = invalidActiveValueDocumentationCount + 1
                If StrComp(CStr(fieldData(1)), "BOOLEAN", vbBinaryCompare) = 0 Then
                    stateValue = CLng(booleanFieldState(CStr(fieldID)))
                    Select Case storedValue
                        Case "FALSE": stateValue = stateValue Or 1
                        Case "TRUE": stateValue = stateValue Or 2
                        Case Else: invalidBooleanExtraCount = invalidBooleanExtraCount + 1
                    End Select
                    booleanFieldState(CStr(fieldID)) = stateValue
                End If
                If valueListFieldHasValue.Exists(CStr(fieldID)) Then valueListFieldHasValue(CStr(fieldID)) = True
            End If
        End If
        valueRs.MoveNext
    Loop
    valueRs.Close
    Set valueRs = Nothing
    duplicateFieldCount = duplicateFieldKeys.count
    duplicateActiveValueCount = duplicateValueKeys.count
    For Each keyValue In booleanFieldState.keys
        stateValue = CLng(booleanFieldState(CStr(keyValue)))
        If (stateValue And 1) = 0 Then missingBooleanFalseCount = missingBooleanFalseCount + 1
        If (stateValue And 2) = 0 Then missingBooleanTrueCount = missingBooleanTrueCount + 1
    Next keyValue
    For Each keyValue In valueListFieldHasValue.keys
        If CBool(valueListFieldHasValue(CStr(keyValue))) = False Then valueListWithoutValuesCount = valueListWithoutValuesCount + 1
    Next keyValue
    detailText = "Database Reference postcondition validation: parent rows reused=" & CStr(parentRowCount) & "; field rows scanned once=" & CStr(fieldRowCount) & "; active value rows scanned once=" & CStr(valueRowCount) & "; active query-field value rows=" & CStr(queryFieldActiveValueCount) & "; repeated COUNT queries=0; elapsed=" & FormattedElapsedSeconds(ElapsedSecondsFromTimer(operationTimer))
    If invalidObjectCount > 0 Then Err.Raise vbObjectError + 9569, "ValidateNonFormDatabaseReferenceFieldIndex", "Active non-form table/query Database Reference rows lack current physical schema verification evidence." & vbCrLf & "Rows: " & CStr(invalidObjectCount)
    If invalidObjectDocumentationCount > 0 Then Err.Raise vbObjectError + 9590, "ValidateNonFormDatabaseReferenceFieldIndex", "Active non-form Database Reference object rows lack synchronized documentation-state metadata." & vbCrLf & "Rows: " & CStr(invalidObjectDocumentationCount)
    If invalidFieldCount > 0 Then Err.Raise vbObjectError + 9570, "ValidateNonFormDatabaseReferenceFieldIndex", "Active non-form Database Reference field rows lack current physical field-index verification evidence." & vbCrLf & "Rows: " & CStr(invalidFieldCount)
    If duplicateFieldCount > 0 Then Err.Raise vbObjectError + 9571, "ValidateNonFormDatabaseReferenceFieldIndex", "Duplicate active Database Reference field identities remain for non-form table/query objects." & vbCrLf & "Duplicate identities: " & CStr(duplicateFieldCount)
    If invalidDocumentationCount > 0 Then Err.Raise vbObjectError + 9591, "ValidateNonFormDatabaseReferenceFieldIndex", "Active non-form Database Reference field rows lack synchronized documentation-state metadata." & vbCrLf & "Rows: " & CStr(invalidDocumentationCount)
    If orphanActiveValueCount > 0 Then Err.Raise vbObjectError + 9592, "ValidateNonFormDatabaseReferenceFieldIndex", "Active Database Reference field-value rows remain attached to inactive non-form fields." & vbCrLf & "Rows: " & CStr(orphanActiveValueCount)
    If queryFieldActiveValueCount > 0 Then Err.Raise vbObjectError + 9597, "ValidateNonFormDatabaseReferenceFieldIndex", "Active controlled-value rows remain attached to query-derived fields. Controlled-value authority belongs only to physical table fields." & vbCrLf & "Rows: " & CStr(queryFieldActiveValueCount)
    If duplicateActiveValueCount > 0 Then Err.Raise vbObjectError + 9593, "ValidateNonFormDatabaseReferenceFieldIndex", "Duplicate active Database Reference controlled-value identities remain for non-form fields." & vbCrLf & "Duplicate identities: " & CStr(duplicateActiveValueCount)
    If invalidActiveValueDocumentationCount > 0 Then Err.Raise vbObjectError + 9594, "ValidateNonFormDatabaseReferenceFieldIndex", "Active Database Reference field-value rows lack synchronized documentation-state metadata." & vbCrLf & "Rows: " & CStr(invalidActiveValueDocumentationCount)
    If missingBooleanFalseCount > 0 Or missingBooleanTrueCount > 0 Or invalidBooleanExtraCount > 0 Then Err.Raise vbObjectError + 9595, "ValidateNonFormDatabaseReferenceFieldIndex", "Boolean Database Reference controlled-value synchronization is incomplete." & vbCrLf & "Missing False sets: " & CStr(missingBooleanFalseCount) & vbCrLf & "Missing True sets: " & CStr(missingBooleanTrueCount) & vbCrLf & "Unexpected active Boolean values: " & CStr(invalidBooleanExtraCount)
    If valueListWithoutValuesCount > 0 Then Err.Raise vbObjectError + 9596, "ValidateNonFormDatabaseReferenceFieldIndex", "One or more active Access Value List fields have no synchronized Database Reference values." & vbCrLf & "Fields: " & CStr(valueListWithoutValuesCount)
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not fieldRs Is Nothing Then fieldRs.Close
    If Not valueRs Is Nothing Then valueRs.Close
    On Error GoTo 0
    If Len(detailText) = 0 Then detailText = "Database Reference postcondition validation failed after " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(operationTimer))
    Err.Raise savedErrNumber, "ValidateNonFormDatabaseReferenceFieldIndex", savedErrDescription
End Sub
