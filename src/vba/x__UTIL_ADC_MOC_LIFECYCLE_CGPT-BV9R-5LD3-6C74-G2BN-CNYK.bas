Attribute VB_Name = "x__UTIL_ADC_MOC_LIFECYCLE"
' ================================================================================================================================================================================
' ChatGPT Rewrite Tracking Header
' ================================================================================================================================================================================
' ChatGPT Previous Version ID: CGPT-APDR-CKS7-HMY7-D4HF-ZW93
' ChatGPT Baseline Version ID: CGPT-APDR-CKS7-HMY7-D4HF-ZW93
' ChatGPT Version ID: CGPT-BV9R-5LD3-6C74-G2BN-CNYK
' ChatGPT Version Timestamp: Thursday 9/3/26 6:51 PM
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
' STANDARD MODULE: Access Development Console Managed Object Catalog Lifecycle
' VB_Name: x__UTIL_ADC_MOC_LIFECYCLE
' Entry Points:
'     ApplyObjectNamingRules
'     DeleteSystemOwnedObjects
' SHARED CODE:
'     Uses existing Access Development Console public/common/shared DAO, tracing, UI, naming, and setup-state services referenced by the moved procedures.
' CLASSES USED: None
' DYNAMIC STRING REFERENCES:
'     Existing dynamic references inside the moved lifecycle procedures are preserved unchanged.
' LastScanTimestamp: Wednesday 8/26/26 2:11 AM
' Purpose:
'     Owns physical Managed Object Catalog lifecycle operations: managed-object rename/preflight/verification and reference repair, catalog-driven table/query/report deletion, relationship cleanup,
'     and the standalone name-based destructive reset fallback including deferred Macro 65 completion.
' Author: Christina Shonts
' Support: Access Development Console CodeDb utility workflows.
' Code Category: Access Development Console Managed Object Catalog physical lifecycle authority.
' Requirements: Microsoft Access, DAO, the companion Access Development Console modules, and x__SHARED_CODE.
' Reference Dependencies:
'     x__UTIL_ADC_MOC supplies managed catalog/recovery/database-context authority; Core/Common/Defaults/Tables/UI modules and shared helpers retain their existing responsibilities.
' DB Object Dependencies:
'     Operates on current CodeDb-owned Access Development Console tables, queries, reports, forms, macros, relationships, and USysRibbons only through the preserved lifecycle rules and deletion sets.
' Usage:
'     Import with the companion Access Development Console modules. Public lifecycle entry points retain their existing names and behavior.
' Version:
'     Delete All now preserves the complete TraceLog physical namespace while continuing to delete every other object in its existing scope; standalone fallback follows the same preservation rule.
'     Setup creation/repair ownership and phase boundaries are unchanged.
'     Adds durable physical trace-file checkpoints around standalone deferred Macro 65 existence, deletion, completion presentation, and state cleanup without changing ownership, ordering, or Setup-phase boundaries.
'     Moves ValidateTerminologyCanonicalNaming from the declarations area to the naming-rule implementation it directly supports; procedure body, behavior, and Setup-phase ownership are unchanged.
'     Moves managed-object rename token replacement and its token-character helper into MOC Lifecycle ownership; rename behavior is unchanged.
'     Created by extracting the complete physical managed-object lifecycle and standalone destructive-reset subsystem from x__UTIL_ADC_CATALOG while renaming the retained catalog authority to
' Notes:
'     Procedure bodies and their existing comments were moved together. Setup-phase ownership, sequencing, destructive confirmation, rollback, recovery gates, and terminal teardown behavior are unchanged.
' Core Rules Enforced Here:
'     Preserve destructive safety, catalog recovery gating, explicit CodeDb/HostDb ownership, rename preflight/verification, relationship cleanup, survivor verification, and terminal Macro 65 handling.
' Design Map:
'     MOC owns catalog definitions/state/recovery; this module owns physical rename/delete/reset lifecycle behavior and calls MOC directly only for catalog recovery readiness and current catalog database resolution.
' Configuration: Existing settings, canonical definitions, and catalog rows remain authoritative.
' Behavior:
'     Existing lifecycle behavior is preserved; this split changes module ownership only.
' Code Status: Active canonical Managed Object Catalog lifecycle authority.
' Todo: None.
' Reporting Bugs: None currently recorded in this header.
' Module Level Bugs: None currently recorded in this header.
' Public Subs:
'     DeleteAllDBObjectsByName() -> None
'     PreviewDBObjectsByName() -> None
' Public Functions:
'     ReplaceManagedObjectNameToken(ByVal sourceText As String, ByVal oldName As String, ByVal newName As String) -> String
'     SystemOwnedDeletionPathEligible(ByVal deletionScope As ObjectDeletionScope, ByRef approvedStandaloneFallback As Boolean, ByRef reasonCode As String, ByRef detailText As String) -> Boolean
'     DeleteSystemOwnedObjects(ByVal deletionScope As ObjectDeletionScope, Optional ByVal showMessage As Boolean = True, Optional ByVal confirmDestructiveAction As Boolean = True) -> Boolean
'     ApplyObjectNamingRules(Optional ByVal showMessage As Boolean = True) -> Boolean
'     ApplyManagedObjectRenameMap(ByVal sourceText As String, ByVal renames As Collection) -> String
'     ConfirmDBObjectsByNameDeletion(Optional ByVal showCancellationMessage As Boolean = True) -> Boolean
'     StandaloneTryDeleteDeferredInvokingMacro65(ByVal macroName As String, ByRef failureText As String, ByRef wasMissing As Boolean) -> Boolean
'     StandaloneDeferredMacroDeletionCompleted(ByVal macroDeleted As Boolean, ByVal macroMissing As Boolean, ByVal finalFailureText As String) -> Boolean
'     DeleteDBObjectsByName(Optional ByVal requireTypedConfirmation As Boolean = True, Optional ByVal showCompletionMessage As Boolean = True, Optional ByVal completeOwnedTraceAfterFinalMacro As Boolean = False,
'     Optional ByVal traceSuccessMessage As String = vbNullString, Optional ByVal traceFailureMessage As String = vbNullString) -> Boolean
' Public Properties: None
' ================================================================================================================================================================================

Option Compare Database
Option Explicit

' Managed-object lifecycle diagnostics.
Private Const DELETE_PREFLIGHT_DIAGNOSTIC_INTERVAL_SECONDS As Double = 5#

' Standalone/name-based destructive deletion engine. This remains independent of a functioning Managed Object Catalog table even though Catalog owns both deletion strategies.

' Standalone destructive reset utility for ADC database objects.
' Purpose: Delete ADC-owned Access database objects without depending on MAC, Setup completion, Module Guard state, or the normal ADC deletion workflow.
' Authority boundary: Object names come from the code-owned canonical managed definitions resolved through x__UTIL_ADC_COMMON plus stale/current system-macro rows and generated macro names already recorded in
'   MC_Records when those tables are readable.
' Form safety: ONLY forms returned by ManagedGeneratedFormNames are deleted. Every other form is preserved, including user/custom forms.

Private Const RIBBON_NAME As String = "ADC"
' Standalone owns only the result summary for its terminal deletion plan. Core owns the post-RunCode timing needed to remove the currently invoking Macro 65 object.
Private mStandaloneDeferredMacroDeleteName As String
Private mStandaloneDeferredMacroDeleteShowCompletion As Boolean
Private mStandaloneDeferredMacroDeleteDeletedCount As Long
Private mStandaloneDeferredMacroDeleteMissingCount As Long
Private mStandaloneDeferredMacroDeleteFailedCount As Long
Private mStandaloneDeferredMacroDeletePreservedFormCount As Long
Private mStandaloneDeferredMacroDeleteFailureText As String

' ===============================================================================================================================================================================
' CONSOLIDATED MANAGED-OBJECT LIFECYCLE PROCEDURES
' ===============================================================================================================================================================================

Private Function TryAccessObjectTypeFromCatalogText(ByVal objectTypeText As String, ByRef objectType As AcObjectType) As Boolean
    Select Case UCase$(objectTypeText)
        Case UCase$(ADCAT_OBJECT_TYPE_MACRO)
            objectType = acMacro
            TryAccessObjectTypeFromCatalogText = True
        Case UCase$(ADCAT_OBJECT_TYPE_QUERY)
            objectType = acQuery
            TryAccessObjectTypeFromCatalogText = True
        Case UCase$(ADCAT_OBJECT_TYPE_TABLE)
            objectType = acTable
            TryAccessObjectTypeFromCatalogText = True
        Case UCase$(ADCAT_OBJECT_TYPE_REPORT)
            objectType = acReport
            TryAccessObjectTypeFromCatalogText = True
        Case UCase$(ADCAT_OBJECT_TYPE_FORM)
            objectType = acForm
            TryAccessObjectTypeFromCatalogText = True
    End Select
End Function

Private Function ManagedPhysicalObjectNamesForRename(ByVal db As DAO.Database, ByVal renames As Collection, ByVal objectTypeCode As String) As Object
    Dim names As Object
    Dim rs As DAO.Recordset
    Dim managedTableName As String
    Dim currentName As String
    Dim resolvedName As String
    ' INTENTIONAL KEYED LOOKUP: names provides keyed membership/lookup over source rows or state already read by ManagedPhysicalObjectNamesForRename; it is a derived access path, not a second authority.
    Set names = CreateObject("Scripting.Dictionary")
    names.CompareMode = vbTextCompare
    managedTableName = ConfiguredRenameTarget(renames, "Table", BootstrapManagedObjectsTableName())
    If DaoTableExists(db, managedTableName) = False Then Err.Raise vbObjectError + 8950, "ManagedPhysicalObjectNamesForRename", "The Managed Object Catalog table is unavailable during catalog-scoped reference repair." & vbCrLf & "Expected table: " & managedTableName
    Set rs = db.OpenRecordset("SELECT " & SQL_FIELD_OBJECT_NAME & " FROM [" & Replace(managedTableName, "]", "]]") & "] WHERE Nz(" & SQL_FIELD_MANAGED_OBJECT_STATUS & ",'Active')='Active' AND UCase(Nz(" & SQL_FIELD_OBJECT_TYPE & ",''))=" & SharedDao_SqlQuote(UCase$(objectTypeCode)) & ";", dbOpenSnapshot)
    Do While rs.EOF = False
        PumpSetupResponsiveness
        currentName = CStr(Nz(rs!objectName, vbNullString))
        resolvedName = ConfiguredRenameTarget(renames, objectTypeCode, currentName)
        If Len(resolvedName) > 0 Then If names.Exists(resolvedName) = False Then names.Add resolvedName, True
        rs.MoveNext
    Loop
    rs.Close
    If StrComp(objectTypeCode, "Table", vbTextCompare) = 0 Then If names.Exists(managedTableName) = False Then names.Add managedTableName, True
    Set ManagedPhysicalObjectNamesForRename = names
End Function

Private Sub PreflightSingleTableRename(ByVal db As DAO.Database, ByVal oldName As String, ByVal newName As String, ByVal databaseRole As String)
    If SharedDao_QueryExists(db, newName) Then Err.Raise vbObjectError + 8920, "PreflightSingleTableRename", "The destination table name is already used by a query." & vbCrLf & "Database role: " & databaseRole & vbCrLf & "Destination: " & newName
    If DaoTableExists(db, oldName) And DaoTableExists(db, newName) Then Err.Raise vbObjectError + 8921, "PreflightSingleTableRename", "Both the source and destination table names exist." & vbCrLf & "Database role: " & databaseRole & vbCrLf & "Source: " & oldName & vbCrLf & "Destination: " & newName
End Sub

Private Function RenameTableInDatabase(ByVal db As DAO.Database, ByVal oldName As String, ByVal newName As String) As Long
    Dim tdf As DAO.tableDef
    If DaoTableExists(db, oldName) = False Then Exit Function
    If DaoTableExists(db, newName) Then Err.Raise vbObjectError + 8924, "RenameTableInDatabase", "The destination table already exists: " & newName
    Set tdf = db.TableDefs(oldName)
    tdf.Name = newName
    db.TableDefs.Refresh
    If DaoTableExists(db, newName) = False Or DaoTableExists(db, oldName) Then Err.Raise vbObjectError + 8925, "RenameTableInDatabase", "The table rename could not be verified." & vbCrLf & "Source: " & oldName & vbCrLf & "Destination: " & newName
    RenameTableInDatabase = 1
End Function

Private Function ObjectNameIsInRetainedTraceLogNamespace(ByVal objectName As String) As Boolean
    'Delete All preservation boundary: retain the TraceLog namespace root and every underscore-delimited physical descendant. This changes deletion eligibility only; Setup remains the sole owner of creation/repair.
    objectName = Trim$(objectName)
    If StrComp(objectName, CANON_TABLE_PHYSICAL_TRACELOG, vbTextCompare) = 0 Then
        ObjectNameIsInRetainedTraceLogNamespace = True
    ElseIf Len(objectName) > Len(CANON_TABLE_PHYSICAL_TRACELOG) Then
        ObjectNameIsInRetainedTraceLogNamespace = (StrComp(Left$(objectName, Len(CANON_TABLE_PHYSICAL_TRACELOG) + 1), CANON_TABLE_PHYSICAL_TRACELOG & "_", vbTextCompare) = 0)
    End If
End Function

Private Function DeletionNamesWithoutRetainedTraceLogNamespace(ByVal deletionNames As Collection, ByVal objectTypeText As String) As Collection
    'Operation-scoped projection used only by terminal Delete All so the established object owner still performs deletion from its own prepared plan.
    Dim filteredNames As New Collection
    Dim item As Variant
    Dim objectName As String
    If deletionNames Is Nothing Then Err.Raise vbObjectError + 6313, "DeletionNamesWithoutRetainedTraceLogNamespace", "The deletion plan is unavailable."
    For Each item In deletionNames
        objectName = CStr(item)
        If ObjectNameIsInRetainedTraceLogNamespace(objectName) Then
            LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "DeletionNamesWithoutRetainedTraceLogNamespace", "DELETE PLAN EXCLUDE", "Name=" & objectName & "; Classification=TraceLogNamespace; Reason=Retained by Delete All policy", "TraceLog namespace is intentionally preserved by terminal Delete All without changing the owning Setup phase.", 0, vbNullString, vbNullString, objectName, objectTypeText, "CodeDb", "SystemMacro65"
        Else
            filteredNames.Add objectName
        End If
    Next item
    Set DeletionNamesWithoutRetainedTraceLogNamespace = filteredNames
End Function

' Purpose: Performs system owned deletion path eligible by coordinating `End` and `RequireNoUnfinishedCatalogRecovery` while preserving this module's existing validation and error semantics.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   deletionScope (Input; ObjectDeletionScope): Scope value limiting which deletion items or operations are included in this call.
'   approvedStandaloneFallback (Output; Boolean): Set True when the procedure determines that approved standalone fallback applies; set False when that condition is not met.
'   reasonCode (Output; String): Receives the resolved reason code produced from this procedure's validated inputs and current ADC state.
'   detailText (Output; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
' Returns: Boolean: True when the system owned deletion path eligible condition is satisfied for the supplied or current ADC context; False otherwise.
Public Function SystemOwnedDeletionPathEligible(ByVal deletionScope As ObjectDeletionScope, ByRef approvedStandaloneFallback As Boolean, ByRef reasonCode As String, ByRef detailText As String) As Boolean
    'Catalog-owned read-only routing authority for System Macro 65 and direct catalog deletion callers.
    Dim managedObjectsTableName As String
    approvedStandaloneFallback = False
    reasonCode = vbNullString
    detailText = vbNullString
    Select Case deletionScope
        Case DELETE_SCOPE_TABLES, DELETE_SCOPE_QUERIES, DELETE_SCOPE_REPORTS, DELETE_SCOPE_ALL
        Case Else
            Err.Raise vbObjectError + 6312, "SystemOwnedDeletionPathEligible", "Unsupported catalog deletion scope: " & CStr(CLng(deletionScope))
    End Select
    'Recovery and execution-context safety gates are authoritative and may never be converted to standalone fallback.
    RequireNoUnfinishedCatalogRecovery "SystemOwnedDeletionPathEligible"
    If CodeDbIsHostDb() = False Then Err.Raise vbObjectError + 6293, "SystemOwnedDeletionPathEligible", "ADC object deletion was stopped before confirmation because this destructive workflow requires CodeDb and HostDb to be the same database." & vbCrLf & "CodeDb workflow: " & CONFIG_OWNER_DATABASE_ROLE_CODE_DB
    managedObjectsTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID)
    If CodeDbTableExists(managedObjectsTableName) = False Then
        approvedStandaloneFallback = True
        reasonCode = "ManagedObjectCatalogMissing"
        detailText = "The authoritative Managed Object Catalog infrastructure is unavailable because the resolved catalog table does not exist." & vbCrLf & "Resolved table: " & managedObjectsTableName
        Exit Function
    End If
    'An existing-but-unopenable catalog is not broadly classified as absent infrastructure. Treat it as a genuine safety/runtime failure.
    If TableCanOpen(managedObjectsTableName) = False Then
        Err.Raise vbObjectError + 6215, "SystemOwnedDeletionPathEligible", "The managed object catalog table exists but cannot be opened, so deletion was stopped before any ADC object was touched." & vbCrLf & "Table: " & managedObjectsTableName & vbCrLf & TableOpenFailureDetail(managedObjectsTableName)
    End If
    reasonCode = "CatalogDeletionEligible"
    detailText = "The authoritative Managed Object Catalog deletion path is available." & vbCrLf & "Resolved table: " & managedObjectsTableName
    SystemOwnedDeletionPathEligible = True
End Function

' Purpose: Deletes system owned objects for deletion scope, confirm destructive action.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   deletionScope (Input; ObjectDeletionScope): Scope value limiting which deletion items or operations are included in this call.
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
'   confirmDestructiveAction (Input Optional; Boolean): When True, enables or requests destructive action for this call; when False, the routine follows its normal non-forced/default path. Default: True.
' Returns: Boolean result of delete system owned objects; True and False correspond to the affirmative and negative branches explicitly implemented by this function.
Public Function DeleteSystemOwnedObjects(ByVal deletionScope As ObjectDeletionScope, Optional ByVal showMessage As Boolean = True, Optional ByVal confirmDestructiveAction As Boolean = True) As Boolean
    'System Macro #65 destructive coordinator. Catalog authority owns the table/query/report snapshot; All delegates macro and generated-form deletion to their established owners before consuming the retained
    'catalog snapshot, with tables last.
    On Error GoTo ErrHandler
    Dim deletedCount As Long
    Dim skippedCount As Long
    Dim failedCount As Long
    Dim detailText As String
    Dim response As VbMsgBoxResult
    Dim typedText As String
    Dim progressStarted As Boolean
    Dim messageText As String
    Dim objectsToDelete As Collection
    Dim deletionManifestText As String
    Dim terminalTeardownStarted As Boolean
    Dim traceTableWriteSuppressionStarted As Boolean
    Dim failureDetailText As String
    Dim traceLogTableName As String
    Dim traceLogScreenshotsTableName As String
    Dim relationshipPlannedCount As Long
    Dim relationshipAttemptedCount As Long
    Dim relationshipDeletedCount As Long
    Dim relationshipSkippedCount As Long
    Dim relationshipFailedCount As Long
    Dim survivorCount As Long
    Dim survivorDetailText As String
    Dim preflightStartTimer As Single
    Dim stageStartTimer As Single
    Dim snapshotElapsed As Double
    Dim manifestElapsed As Double
    Dim traceLogLookupElapsed As Double
    Dim totalPreflightElapsed As Double
    Dim otherPreflightElapsed As Double
    Dim scopeIncludesTables As Boolean
    Dim scopeIncludesGeneratedForms As Boolean
    Dim scopeIncludesMacros As Boolean
    Dim scopeIncludesReports As Boolean
    Dim scopeLabel As String
    Dim planItem As Variant
    Dim plannedTableCount As Long
    Dim plannedQueryCount As Long
    Dim plannedReportCount As Long
    Dim macroPlannedCount As Long
    Dim macroAttemptedCount As Long
    Dim macroDeletedCount As Long
    Dim macroSkippedCount As Long
    Dim macroFailedCount As Long
    Dim formPlannedCount As Long
    Dim formAttemptedCount As Long
    Dim formDeletedCount As Long
    Dim formSkippedCount As Long
    Dim formFailedCount As Long
    Dim macroDeletionPlan As Collection
    Dim formDeletionPlan As Collection
    Dim macroOperationSucceeded As Boolean
    Dim formOperationSucceeded As Boolean
    Dim aggregatePlannedCount As Long
    Dim aggregateAttemptedCount As Long
    Dim aggregateDeletedCount As Long
    Dim aggregateSkippedCount As Long
    Dim aggregateFailedCount As Long
    Dim overallSucceeded As Boolean
    Dim macro65Deferred As Boolean
    Dim catalogDeletionEligible As Boolean
    Dim approvedStandaloneFallback As Boolean
    Dim deletionEligibilityReason As String
    Dim deletionEligibilityDetail As String
    Dim preserveTraceLogNamespace As Boolean
    macroOperationSucceeded = True
    formOperationSucceeded = True
    Select Case deletionScope
        Case DELETE_SCOPE_TABLES
            scopeIncludesTables = True
            scopeLabel = "Tables"
        Case DELETE_SCOPE_QUERIES
            scopeLabel = "Queries"
        Case DELETE_SCOPE_REPORTS
            scopeIncludesReports = True
            scopeLabel = "Reports"
        Case DELETE_SCOPE_ALL
            scopeIncludesTables = True
            scopeIncludesGeneratedForms = True
            scopeIncludesMacros = True
            scopeIncludesReports = True
            preserveTraceLogNamespace = True
            scopeLabel = "All"
        Case Else
            Err.Raise vbObjectError + 6309, "DeleteSystemOwnedObjects", "Unsupported catalog deletion scope: " & CStr(CLng(deletionScope))
    End Select
    catalogDeletionEligible = SystemOwnedDeletionPathEligible(deletionScope, approvedStandaloneFallback, deletionEligibilityReason, deletionEligibilityDetail)
    If catalogDeletionEligible = False Then
        Err.Raise vbObjectError + 6214, "DeleteSystemOwnedObjects", deletionEligibilityDetail
    End If
    LogTraceMessage "Info", "ManagedObjectCatalog", "DeleteSystemOwnedObjects", "ADC managed-object deletion started.", "Scope: " & scopeLabel, vbNullString, 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID), "Table", "CodeDb", "DeleteSystemOwnedObjects"
    If confirmDestructiveAction Then
        Select Case deletionScope
            Case DELETE_SCOPE_TABLES
                messageText = "This will delete ADC system-owned/generated tables. Queries, generated forms, System Macros and Help Document Macros, and reports will remain."
            Case DELETE_SCOPE_QUERIES
                messageText = "This will delete ADC system-owned/generated saved queries. Tables, generated forms, System Macros and Help Document Macros, and reports will remain."
            Case DELETE_SCOPE_REPORTS
                messageText = "This will delete ADC-managed reports. Tables, queries, generated forms, and System Macros and Help Document Macros will remain."
            Case DELETE_SCOPE_ALL
                messageText = "This will delete ADC system-owned/generated tables, saved queries, the canonical generated-form set except the TraceLog namespace, ADC System Macros and Help Document Macros, and ADC-managed reports." & vbCrLf & vbCrLf & "The TraceLog namespace will remain. Unrelated user-owned objects will also remain."
        End Select
        messageText = messageText & vbCrLf & vbCrLf & "No backup will be created. This action is destructive. Continue?"
        response = CoreShowMessageResult(messageText, vbCritical Or vbYesNo Or vbDefaultButton2, "Delete All ADC Objects", "ShowADCMessageResult", scopeLabel)
        If response <> vbYes Then
            LogTraceMessage "Info", "ManagedObjectCatalog", "DeleteSystemOwnedObjects", "Delete All ADC Objects cancelled before typed confirmation.", "Scope: " & scopeLabel & vbCrLf & "User response: " & CStr(response), vbNullString, 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID), "Table", "CodeDb", "DeleteSystemOwnedObjects"
            MarkSystemCommandCancelled
            Exit Function
        End If
        typedText = InputBox("Type DELETE ALL to delete the selected ADC objects." & vbCrLf & vbCrLf & "Scope: " & scopeLabel & vbCrLf & "No backup will be created.", "Confirm Delete All ADC Objects")
        If StrComp(Trim$(typedText), "DELETE ALL", vbBinaryCompare) <> 0 Then
            messageText = "Delete All ADC Objects was cancelled because the confirmation text did not match DELETE ALL."
            LogTraceMessage "Info", "ManagedObjectCatalog", "DeleteSystemOwnedObjects", messageText, "Scope: " & scopeLabel & vbCrLf & "Typed confirmation text did not match the required phrase.", vbNullString, 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID), "Table", "CodeDb", "DeleteSystemOwnedObjects"
            If showMessage Then CoreShowMessage messageText, vbInformation, "Delete All ADC Objects", "ShowADCMessage", scopeLabel
            MarkSystemCommandCancelled
            Exit Function
        End If
    End If
    progressStarted = CoreBeginOperationProgressIfNeeded("Prepare ADC object deletion", 1)
    CorePulseProgress "Preparing ADC object deletion | building deletion snapshot"
    preflightStartTimer = Timer
    stageStartTimer = preflightStartTimer
    Set objectsToDelete = SystemOwnedObjectDeletionSnapshot(preflightStartTimer, deletionScope)
    snapshotElapsed = ElapsedSecondsFromTimer(stageStartTimer)
    For Each planItem In objectsToDelete
        Select Case UCase$(Trim$(CStr(planItem(1))))
            Case UCase$(ADCAT_OBJECT_TYPE_TABLE): plannedTableCount = plannedTableCount + 1
            Case UCase$(ADCAT_OBJECT_TYPE_QUERY): plannedQueryCount = plannedQueryCount + 1
            Case UCase$(ADCAT_OBJECT_TYPE_REPORT): plannedReportCount = plannedReportCount + 1
        End Select
    Next planItem
    LogTraceMessage "Info", "ManagedObjectCatalog", "DeleteSystemOwnedObjects", "Deletion preflight | build catalog deletion snapshot | complete", "Scope: " & scopeLabel & vbCrLf & "Stage elapsed: " & FormattedElapsedSeconds(snapshotElapsed) & vbCrLf & "Objects retained: " & CStr(objectsToDelete.count), "Deletion preflight snapshot", 0, vbNullString, vbNullString, vbNullString, vbNullString, "CodeDb", "DeleteSystemOwnedObjects"
    If scopeIncludesMacros Then
        Set macroDeletionPlan = PrepareSystemHelpMacroDeletionPlan(macroPlannedCount)
        If preserveTraceLogNamespace Then
            Set macroDeletionPlan = DeletionNamesWithoutRetainedTraceLogNamespace(macroDeletionPlan, ADCAT_OBJECT_TYPE_MACRO)
            macroPlannedCount = macroDeletionPlan.count
        End If
    End If
    If scopeIncludesGeneratedForms Then
        Set formDeletionPlan = x__UTIL_ADC_UI_FORMS.PrepareGeneratedFormDeletionPlan(formPlannedCount, True)
        If preserveTraceLogNamespace Then
            Set formDeletionPlan = DeletionNamesWithoutRetainedTraceLogNamespace(formDeletionPlan, ADCAT_OBJECT_TYPE_FORM)
            formPlannedCount = formDeletionPlan.count
        End If
    End If
    aggregatePlannedCount = objectsToDelete.count + macroPlannedCount + formPlannedCount
    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "DeleteSystemOwnedObjects", "DELETE PLAN", "Scope=" & scopeLabel & "; Tables=" & CStr(plannedTableCount) & "; Queries=" & CStr(plannedQueryCount) & "; Reports=" & CStr(plannedReportCount) & "; Macros=" & CStr(macroPlannedCount) & "; GeneratedForms=" & CStr(formPlannedCount) & "; Total=" & CStr(aggregatePlannedCount), "Macro 65 aggregate deletion plan before destructive work", 0, vbNullString, vbNullString, vbNullString, "ManagedObjectSet", "CodeDb", "DeleteSystemOwnedObjects"
    If aggregatePlannedCount = 0 Then LogTraceMessage TRACE_LEVEL_WARNING, "DeleteObjects", "DeleteSystemOwnedObjects", "DELETE PLAN EMPTY", "Scope=" & scopeLabel & "; AggregatePlanned=0", "Macro 65 aggregate deletion plan", 0, vbNullString, vbNullString, vbNullString, "ManagedObjectSet", "CodeDb", "DeleteSystemOwnedObjects"
    CorePulseProgress "Preparing ADC object deletion | building retained deletion manifest"
    stageStartTimer = Timer
    PreflightSystemOwnedObjectDeletion objectsToDelete, deletionManifestText, preflightStartTimer
    manifestElapsed = ElapsedSecondsFromTimer(stageStartTimer)
    If scopeIncludesTables And preserveTraceLogNamespace = False Then
        CorePulseProgress "Preparing ADC object deletion | resolving TraceLog name"
        stageStartTimer = Timer
        traceLogTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_TraceLog")
        traceLogScreenshotsTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_TraceLogScreenshots")
        traceLogLookupElapsed = ElapsedSecondsFromTimer(stageStartTimer)
    End If
    LogTraceMessage "Info", "ManagedObjectCatalog", "DeleteSystemOwnedObjects", "ADC catalog deletion preflight completed; deletion manifest retained.", "Scope: " & scopeLabel & vbCrLf & deletionManifestText, "Retain deletion manifest", 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID), "Table", "CodeDb", "DeleteSystemOwnedObjects"
    totalPreflightElapsed = ElapsedSecondsFromTimer(preflightStartTimer)
    otherPreflightElapsed = totalPreflightElapsed - snapshotElapsed - manifestElapsed - traceLogLookupElapsed
    If otherPreflightElapsed < 0 Then otherPreflightElapsed = 0
    LogTraceMessage "Info", "ManagedObjectCatalog", "DeleteSystemOwnedObjects", "Macro 65 deletion preflight diagnostics complete.", _
        "Scope: " & scopeLabel & vbCrLf & "Snapshot elapsed: " & FormattedElapsedSeconds(snapshotElapsed) & vbCrLf & "Manifest elapsed: " & FormattedElapsedSeconds(manifestElapsed) & vbCrLf & _
        "TraceLog-name lookup elapsed: " & FormattedElapsedSeconds(traceLogLookupElapsed) & vbCrLf & "Other preflight elapsed: " & FormattedElapsedSeconds(otherPreflightElapsed) & vbCrLf & "Total preflight elapsed: " & FormattedElapsedSeconds(totalPreflightElapsed) & vbCrLf & _
        "Objects in catalog deletion snapshot: " & CStr(objectsToDelete.count), "Deletion preflight summary", 0, vbNullString, vbNullString, vbNullString, vbNullString, "CodeDb", "DeleteSystemOwnedObjects"
    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "DeleteSystemOwnedObjects", "DELETE DISPATCH", "Scope=" & scopeLabel, "Macro 65 deletion dispatch", 0, vbNullString, vbNullString, vbNullString, "ManagedObjectSet", "CodeDb", "DeleteSystemOwnedObjects"
    If scopeIncludesMacros Then
        CorePulseProgress "Deleting ADC System Macros and Help Document Macros before catalog teardown"
        macroOperationSucceeded = DeleteSystemHelpMacros(False, False, macroPlannedCount, macroAttemptedCount, macroDeletedCount, macroSkippedCount, macroFailedCount, macroDeletionPlan)
        If macroOperationSucceeded = False And SystemCommandWasCancelled() = False Then Err.Raise vbObjectError + 6307, "DeleteSystemOwnedObjects", "The ADC System Macro and Help Document Macro owner did not complete deletion."
    End If
    If scopeIncludesGeneratedForms Then
        CorePulseProgress "Deleting generated forms before catalog teardown"
        formOperationSucceeded = False
        x__UTIL_ADC_UI_FORMS.DeleteGeneratedForms False, False, formPlannedCount, formAttemptedCount, formDeletedCount, formSkippedCount, formFailedCount, formOperationSucceeded, formDeletionPlan
        If SystemCommandWasCancelled() Then
            If progressStarted Then CoreCompleteSetupProgress
            Exit Function
        End If
        If formOperationSucceeded = False Then Err.Raise vbObjectError + 6310, "DeleteSystemOwnedObjects", "The generated-form deletion owner did not complete deletion."
    End If
    If scopeIncludesTables Then
        If preserveTraceLogNamespace Then
            LogTraceMessage "Info", "ManagedObjectCatalog", "DeleteSystemOwnedObjects", "Terminal table teardown is beginning from the retained in-memory snapshot.", "Reports and queries in the retained snapshot are deleted before tables. The TraceLog namespace is retained, so database-backed TraceLog writes remain active. Ribbon callbacks will be made inert and the ADC startup CustomRibbonID will be disabled before destructive table deletion starts.", "Delete preflighted snapshot", 0, vbNullString, vbNullString, vbNullString, vbNullString, "CodeDb", "DeleteSystemOwnedObjects"
        Else
            LogTraceMessage "Info", "ManagedObjectCatalog", "DeleteSystemOwnedObjects", "Terminal table teardown is beginning from the retained in-memory snapshot.", "Reports and queries in the retained snapshot are deleted before tables. TraceLog is deferred until ADC-backed progress is closed. Ribbon callbacks will be made inert and the ADC startup CustomRibbonID will be disabled before destructive table deletion starts.", "Delete preflighted snapshot", 0, vbNullString, vbNullString, vbNullString, vbNullString, "CodeDb", "DeleteSystemOwnedObjects"
            BeginTerminalTraceTableTeardown
            traceTableWriteSuppressionStarted = True
        End If
        x__UTIL_ADC_UI_RIBBON.BeginTerminalTeardown
        terminalTeardownStarted = True
        If preserveTraceLogNamespace Then
            LogTraceMessage "Info", "ManagedObjectCatalog", "DeleteSystemOwnedObjects", "Terminal teardown state is active and the ADC startup Ribbon has been disabled.", "TraceLog namespace objects and database-backed TraceLog writes remain active by Delete All policy.", "Begin terminal teardown", 0, vbNullString, vbNullString, vbNullString, "Ribbon", "CodeDb", "DeleteSystemOwnedObjects"
        Else
            LogTraceMessage "Info", "ManagedObjectCatalog", "DeleteSystemOwnedObjects", "Terminal teardown state is active and the ADC startup Ribbon has been disabled.", "TraceLog table writes are suppressed for the terminal path; external trace-file output remains active.", "Begin terminal teardown", 0, vbNullString, vbNullString, vbNullString, "Ribbon", "CodeDb", "DeleteSystemOwnedObjects"
        End If
    End If
    DeleteSystemOwnedObjectsCore objectsToDelete, traceLogScreenshotsTableName, traceLogTableName, progressStarted, deletedCount, skippedCount, failedCount, detailText, failureDetailText, scopeIncludesTables, relationshipPlannedCount, relationshipAttemptedCount, relationshipDeletedCount, relationshipSkippedCount, relationshipFailedCount
    VerifySystemOwnedDeletionSurvivors objectsToDelete, macroDeletionPlan, formDeletionPlan, survivorCount, survivorDetailText
    If survivorCount > 0 Then
        detailText = SharedText_AppendNonBlankLine(detailText, survivorDetailText)
        failureDetailText = SharedText_AppendNonBlankLine(failureDetailText, survivorDetailText)
    End If
    aggregateAttemptedCount = deletedCount + skippedCount + failedCount + macroAttemptedCount + formAttemptedCount
    aggregateDeletedCount = deletedCount + macroDeletedCount + formDeletedCount
    aggregateSkippedCount = skippedCount + macroSkippedCount + formSkippedCount
    aggregateFailedCount = failedCount + macroFailedCount + formFailedCount
    macro65Deferred = CoreDeferredMacro65DeletionRequired()
    overallSucceeded = (aggregateFailedCount = 0 And relationshipFailedCount = 0 And survivorCount = 0 And macroOperationSucceeded And formOperationSucceeded)
    If overallSucceeded Then
        messageText = "Delete All ADC Objects completed successfully."
    Else
        messageText = "Delete All ADC Objects completed with failures."
    End If
    messageText = messageText & vbCrLf & vbCrLf & "Scope: " & scopeLabel & vbCrLf & "Planned: " & CStr(aggregatePlannedCount) & vbCrLf & "Attempted: " & CStr(aggregateAttemptedCount + IIf(macro65Deferred, 1, 0)) & vbCrLf & "Deleted: " & CStr(aggregateDeletedCount + IIf(macro65Deferred And overallSucceeded, 1, 0)) & vbCrLf & "Skipped: " & CStr(aggregateSkippedCount) & vbCrLf & "Failed: " & CStr(aggregateFailedCount) & vbCrLf & "Survivors: " & CStr(survivorCount)
    If scopeIncludesTables Then messageText = messageText & vbCrLf & "Relationships deleted: " & CStr(relationshipDeletedCount) & " of " & CStr(relationshipPlannedCount) & "; relationship failures: " & CStr(relationshipFailedCount)
    If scopeIncludesReports And plannedReportCount = 0 Then messageText = messageText & vbCrLf & "Reports: no active ADC-managed report rows were present in the catalog snapshot."
    If preserveTraceLogNamespace Then messageText = messageText & vbCrLf & "TraceLog namespace: retained."
    messageText = messageText & DeleteSystemOwnedObjectsDetailsForMessage(detailText)
    On Error Resume Next
    Application.SysCmd acSysCmdClearStatus
    If macro65Deferred And overallSucceeded Then
        CoreConfigureDeferredMacro65NormalCompletion showMessage, messageText, "Delete All ADC Objects"
    ElseIf showMessage Then
        If overallSucceeded Then
            VBA.MsgBox messageText, vbInformation, "Delete All ADC Objects"
        Else
            VBA.MsgBox messageText, vbCritical, "Delete All ADC Objects"
        End If
    End If
    On Error GoTo 0
    If overallSucceeded Then
        LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "DeleteSystemOwnedObjects", IIf(macro65Deferred, "DELETE ALL ADC OBJECTS ORDINARY WORK COMPLETE", "DELETE ALL ADC OBJECTS COMPLETE"), "Scope=" & scopeLabel & "; Planned=" & CStr(aggregatePlannedCount) & "; Attempted=" & CStr(aggregateAttemptedCount) & "; Deleted=" & CStr(aggregateDeletedCount) & "; Skipped=" & CStr(aggregateSkippedCount) & "; Failed=0; Survivors=" & CStr(survivorCount) & "; DeferredInvokingMacro=" & CStr(macro65Deferred) & "; RelationshipsFailed=" & CStr(relationshipFailedCount) & "; CatalogPlanned=" & CStr(objectsToDelete.count) & "; MacroPlanned=" & CStr(macroPlannedCount) & "; GeneratedFormPlanned=" & CStr(formPlannedCount), "Macro 65 deletion final totals", 0, vbNullString, vbNullString, vbNullString, "ManagedObjectSet", "CodeDb", "DeleteSystemOwnedObjects"
    Else
        LogTraceMessage TRACE_LEVEL_ERROR, "DeleteObjects", "DeleteSystemOwnedObjects", "DELETE ALL ADC OBJECTS COMPLETE WITH FAILURES", "Scope=" & scopeLabel & "; Planned=" & CStr(aggregatePlannedCount) & "; Attempted=" & CStr(aggregateAttemptedCount) & "; Deleted=" & CStr(aggregateDeletedCount) & "; Skipped=" & CStr(aggregateSkippedCount) & "; Failed=" & CStr(aggregateFailedCount) & "; Survivors=" & CStr(survivorCount) & "; RelationshipsFailed=" & CStr(relationshipFailedCount) & "; CatalogPlanned=" & CStr(objectsToDelete.count) & "; MacroPlanned=" & CStr(macroPlannedCount) & "; GeneratedFormPlanned=" & CStr(formPlannedCount), "Macro 65 deletion final totals", 0, vbNullString, vbNullString, vbNullString, "ManagedObjectSet", "CodeDb", "DeleteSystemOwnedObjects"
    End If
    DeleteSystemOwnedObjects = overallSucceeded
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If terminalTeardownStarted Then
        messageText = "Could not complete ADC object deletion after terminal table teardown had begun." & vbCrLf & vbCrLf & "Deleted catalog objects before failure: " & CStr(deletedCount) & vbCrLf & "Skipped catalog objects before failure: " & CStr(skippedCount) & vbCrLf & "Failed catalog objects before failure: " & CStr(failedCount) & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description:" & vbCrLf & savedErrDescription
        On Error Resume Next
        LogTraceMessage "Error", "ManagedObjectCatalog", "DeleteSystemOwnedObjects", "Terminal ADC object deletion failed.", messageText & DeleteSystemOwnedObjectsDetailsForMessage(failureDetailText), "Terminal delete failure", savedErrNumber, savedErrSource, savedErrDescription, vbNullString, vbNullString, "CodeDb", "DeleteSystemOwnedObjects"
        FlushTraceLogFile
        If progressStarted Then CoreFailSetupProgress
        Application.SysCmd acSysCmdClearStatus
        If showMessage Then VBA.MsgBox messageText & DeleteSystemOwnedObjectsDetailsForMessage(failureDetailText), vbCritical, "Delete All ADC Objects"
        On Error GoTo 0
        DeleteSystemOwnedObjects = False
        Exit Function
    End If
    If traceTableWriteSuppressionStarted Then EndTraceTableWriteSuppression
    If progressStarted Then CoreFailSetupProgress
    messageText = "Could not delete the selected ADC objects." & vbCrLf & vbCrLf & "Procedure: DeleteSystemOwnedObjects" & vbCrLf & "Scope: " & scopeLabel & vbCrLf & "Deleted catalog objects before failure: " & CStr(deletedCount) & vbCrLf & "Skipped catalog objects before failure: " & CStr(skippedCount) & vbCrLf & "Failed catalog objects before failure: " & CStr(failedCount) & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description:" & vbCrLf & savedErrDescription & TableDiagnosticsLocationForMessage() & CoreOperationSummaryForMessage()
    LogTraceMessage "Error", "ManagedObjectCatalog", "DeleteSystemOwnedObjects", "ADC object deletion failed.", messageText, vbNullString, savedErrNumber, savedErrSource, savedErrDescription, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID), "Table", "CodeDb", "DeleteSystemOwnedObjects"
    If showMessage Then
        CoreShowErrorMessage messageText, vbCritical, "Delete All ADC Objects", "DeleteSystemOwnedObjects", scopeLabel
        Exit Function
    End If
    Err.Raise savedErrNumber, "DeleteSystemOwnedObjects", messageText
End Function

Private Sub PreflightSystemOwnedObjectDeletion(ByVal objectsToDelete As Collection, ByRef deletionManifestText As String, ByVal overallPreflightTimer As Single)
    On Error GoTo ErrHandler
    Dim item As Variant
    Dim objectName As String
    Dim objectTypeText As String
    Dim manifestTimer As Single
    Dim lastProgressTimer As Single
    Dim processedCount As Long
    Dim totalCount As Long
    Dim manifestElapsed As Double
    If objectsToDelete Is Nothing Then Err.Raise vbObjectError + 6304, "PreflightSystemOwnedObjectDeletion", "The system-owned object deletion snapshot is unavailable, so destructive work was stopped."
    totalCount = objectsToDelete.count
    manifestTimer = Timer
    lastProgressTimer = manifestTimer
    LogTraceMessage "Info", "ManagedObjectCatalog", "PreflightSystemOwnedObjectDeletion", "Deletion preflight | manifest | starting", "Objects expected: " & CStr(totalCount) & vbCrLf & "Total elapsed preflight: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(overallPreflightTimer)) & vbCrLf & "Naming diagnostics: " & NamingPerformanceSummary(), "Deletion preflight manifest", 0, vbNullString, vbNullString, vbNullString, vbNullString, "CodeDb", "DeleteSystemOwnedObjects"
    deletionManifestText = "System-owned object deletion manifest captured before destructive work:"
    LogTraceMessage "Info", "ManagedObjectCatalog", "PreflightSystemOwnedObjectDeletion", "Deletion preflight | manifest | building", "Objects expected: " & CStr(totalCount), "Deletion preflight manifest", 0, vbNullString, vbNullString, vbNullString, vbNullString, "CodeDb", "DeleteSystemOwnedObjects"
    For Each item In objectsToDelete
        PumpSetupResponsiveness
        objectName = Trim$(CStr(item(0)))
        objectTypeText = Trim$(CStr(item(1)))
        deletionManifestText = SharedText_AppendNonBlankLine(deletionManifestText, objectTypeText & ": " & objectName)
        processedCount = processedCount + 1
        If ElapsedSecondsFromTimer(lastProgressTimer) >= DELETE_PREFLIGHT_DIAGNOSTIC_INTERVAL_SECONDS Then
            manifestElapsed = ElapsedSecondsFromTimer(manifestTimer)
            CorePulseProgress "Preparing system-owned object deletion | manifest " & CStr(processedCount) & " of " & CStr(totalCount) & " objects"
            LogTraceMessage "Info", "ManagedObjectCatalog", "PreflightSystemOwnedObjectDeletion", "Deletion preflight | manifest", "Objects processed: " & CStr(processedCount) & " of " & CStr(totalCount) & vbCrLf & "Current object: " & objectTypeText & ": " & objectName & vbCrLf & "Stage elapsed: " & FormattedElapsedSeconds(manifestElapsed) & vbCrLf & "Overall preflight elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(overallPreflightTimer)) & vbCrLf & "Naming diagnostics: " & NamingPerformanceSummary(), "Deletion preflight manifest", 0, vbNullString, vbNullString, objectName, objectTypeText, "CodeDb", "DeleteSystemOwnedObjects"
            lastProgressTimer = Timer
        End If
    Next item
    manifestElapsed = ElapsedSecondsFromTimer(manifestTimer)
    LogTraceMessage "Info", "ManagedObjectCatalog", "PreflightSystemOwnedObjectDeletion", "Deletion preflight | manifest | complete", "Objects processed: " & CStr(processedCount) & vbCrLf & "Manifest characters: " & CStr(Len(deletionManifestText)) & vbCrLf & "Stage elapsed: " & FormattedElapsedSeconds(manifestElapsed) & vbCrLf & "Overall preflight elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(overallPreflightTimer)) & vbCrLf & "Naming diagnostics: " & NamingPerformanceSummary(), "Deletion preflight manifest", 0, vbNullString, vbNullString, vbNullString, vbNullString, "CodeDb", "DeleteSystemOwnedObjects"
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "PreflightSystemOwnedObjectDeletion", "System-owned object deletion preflight failed before any object was deleted." & vbCrLf & "Object: " & objectName & vbCrLf & "Object type: " & objectTypeText & vbCrLf & "Err.Description: " & originalErrDescription
End Sub

Private Sub DeleteSystemOwnedObjectsCore(ByVal objectsToDelete As Collection, ByVal traceLogScreenshotsTableName As String, ByVal traceLogTableName As String, ByRef progressStarted As Boolean, ByRef deletedCount As Long, ByRef skippedCount As Long, ByRef failedCount As Long, ByRef detailText As String, ByRef failureDetailText As String, ByVal terminalTableTeardown As Boolean, ByRef relationshipPlannedCount As Long, ByRef relationshipAttemptedCount As Long, ByRef relationshipDeletedCount As Long, ByRef relationshipSkippedCount As Long, ByRef relationshipFailedCount As Long)
    On Error GoTo ErrHandler
    Dim objectTypeItem As Variant
    Dim item As Variant
    Dim itemIndex As Long
    Dim firstItemIndex As Long
    Dim lastItemIndex As Long
    Dim itemIndexStep As Long
    Dim objectName As String
    Dim objectTypeText As String
    Dim deferredTraceLogScreenshots As Boolean
    Dim deferredTraceLog As Boolean
    Dim relationshipTeardownPerformed As Boolean
    Dim summaryText As String
    If objectsToDelete Is Nothing Then Err.Raise vbObjectError + 6305, "DeleteSystemOwnedObjectsCore", "The ADC catalog deletion snapshot is unavailable, so deletion was stopped."
    traceLogScreenshotsTableName = Trim$(traceLogScreenshotsTableName)
    traceLogTableName = Trim$(traceLogTableName)
    For Each objectTypeItem In Array(ADCAT_OBJECT_TYPE_REPORT, ADCAT_OBJECT_TYPE_QUERY, ADCAT_OBJECT_TYPE_TABLE)
        If StrComp(CStr(objectTypeItem), ADCAT_OBJECT_TYPE_TABLE, vbTextCompare) = 0 Then
            If terminalTableTeardown And relationshipTeardownPerformed = False Then
                DeleteSystemOwnedRelationshipsForSnapshot objectsToDelete, relationshipPlannedCount, relationshipAttemptedCount, relationshipDeletedCount, relationshipSkippedCount, relationshipFailedCount, detailText, failureDetailText
                relationshipTeardownPerformed = True
            End If
            firstItemIndex = objectsToDelete.count
            lastItemIndex = 1
            itemIndexStep = -1
        Else
            firstItemIndex = 1
            lastItemIndex = objectsToDelete.count
            itemIndexStep = 1
        End If
        For itemIndex = firstItemIndex To lastItemIndex Step itemIndexStep
            item = objectsToDelete(itemIndex)
            objectName = CStr(item(0))
            objectTypeText = CStr(item(1))
            If StrComp(objectTypeText, CStr(objectTypeItem), vbTextCompare) = 0 Then
                If terminalTableTeardown And StrComp(objectTypeText, ADCAT_OBJECT_TYPE_TABLE, vbTextCompare) = 0 And Len(traceLogScreenshotsTableName) > 0 And StrComp(objectName, traceLogScreenshotsTableName, vbTextCompare) = 0 Then
                    deferredTraceLogScreenshots = True
                ElseIf terminalTableTeardown And StrComp(objectTypeText, ADCAT_OBJECT_TYPE_TABLE, vbTextCompare) = 0 And Len(traceLogTableName) > 0 And StrComp(objectName, traceLogTableName, vbTextCompare) = 0 Then
                    deferredTraceLog = True
                Else
                    DeleteSystemOwnedObjectSnapshotItem objectName, objectTypeText, deletedCount, skippedCount, failedCount, detailText, failureDetailText
                End If
            End If
        Next itemIndex
    Next objectTypeItem
    summaryText = "Deleted: " & CStr(deletedCount) & vbCrLf & "Skipped: " & CStr(skippedCount) & vbCrLf & "Failed: " & CStr(failedCount)
    If relationshipTeardownPerformed Then summaryText = summaryText & vbCrLf & "Relationships planned: " & CStr(relationshipPlannedCount) & vbCrLf & "Relationships deleted: " & CStr(relationshipDeletedCount) & vbCrLf & "Relationship failures: " & CStr(relationshipFailedCount)
    If deferredTraceLogScreenshots Then summaryText = summaryText & vbCrLf & "Deferred: TraceLog_Screenshots=True"
    If deferredTraceLog Then summaryText = summaryText & vbCrLf & "Deferred: TraceLog=True"
    If Len(Trim$(failureDetailText)) > 0 Then summaryText = summaryText & vbCrLf & "Failure details:" & vbCrLf & failureDetailText
    LogTraceMessage "Info", "ManagedObjectCatalog", "DeleteSystemOwnedObjectsCore", "Bulk ADC catalog-object deletion completed.", summaryText, "Deletion summary before deferred trace-table removal", 0, vbNullString, vbNullString, traceLogTableName, IIf(terminalTableTeardown, "Table", "ManagedObjectSet"), "CodeDb", "DeleteSystemOwnedObjects"
    FlushTraceLogFile
    If progressStarted Then
        CoreCompleteSetupProgress
        progressStarted = False
    End If
    If deferredTraceLogScreenshots Then DeleteSystemOwnedObjectSnapshotItem traceLogScreenshotsTableName, ADCAT_OBJECT_TYPE_TABLE, deletedCount, skippedCount, failedCount, detailText, failureDetailText
    If deferredTraceLog Then DeleteSystemOwnedObjectSnapshotItem traceLogTableName, ADCAT_OBJECT_TYPE_TABLE, deletedCount, skippedCount, failedCount, detailText, failureDetailText
    If deferredTraceLogScreenshots Or deferredTraceLog Then
        summaryText = "Deleted: " & CStr(deletedCount) & vbCrLf & "Skipped: " & CStr(skippedCount) & vbCrLf & "Failed: " & CStr(failedCount)
        If Len(Trim$(failureDetailText)) > 0 Then summaryText = summaryText & vbCrLf & "Failure details:" & vbCrLf & failureDetailText
        LogTraceMessage "Info", "ManagedObjectCatalog", "DeleteSystemOwnedObjectsCore", "Terminal ADC trace-table deletion completed.", summaryText, "Terminal deletion final summary", 0, vbNullString, vbNullString, vbNullString, vbNullString, "CodeDb", "DeleteSystemOwnedObjects"
        FlushTraceLogFile
    End If
    Exit Sub
ErrHandler:
    Dim originalErrNumber As Long
    Dim originalErrSource As String
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrSource = Err.Source
    originalErrDescription = Err.Description
    Err.Raise originalErrNumber, "DeleteSystemOwnedObjectsCore", "Unable to delete the current MAC catalog snapshot." & vbCrLf & "Object: " & objectName & vbCrLf & "Object type: " & objectTypeText & vbCrLf & "Err.Source: " & originalErrSource & vbCrLf & "Description: " & originalErrDescription
End Sub

Private Sub DeleteSystemOwnedRelationshipsForSnapshot(ByVal objectsToDelete As Collection, ByRef plannedCount As Long, ByRef attemptedCount As Long, ByRef deletedCount As Long, ByRef skippedCount As Long, ByRef failedCount As Long, ByRef detailText As String, ByRef failureDetailText As String)
    On Error GoTo ErrHandler
    Dim tableNames As Object
    Dim item As Variant
    Dim Definition As Variant
    Dim definitions As Collection
    Dim db As DAO.Database
    Dim relationshipName As String
    Dim primaryTable As String
    Dim foreignTable As String
    Dim namingSnapshotStartedHere As Boolean
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    Set tableNames = CreateObject("Scripting.Dictionary")
    tableNames.CompareMode = vbTextCompare
    For Each item In objectsToDelete
        If StrComp(Trim$(CStr(item(1))), ADCAT_OBJECT_TYPE_TABLE, vbTextCompare) = 0 Then tableNames(Trim$(CStr(item(0)))) = True
    Next item
    On Error Resume Next
    Application.SysCmd acSysCmdSetStatus, "Preparing ADC relationship teardown..."
    On Error GoTo ErrHandler
    'Build canonical relationship definitions under one complete naming snapshot so repeated physical-name resolution does not reopen MAC for each lookup.
    If NamingSnapshotIsReady() = False Then
        BeginNamingSnapshot
        namingSnapshotStartedHere = True
    End If
    Set definitions = CanonicalRelationshipDefinitions()
    If namingSnapshotStartedHere Then
        InvalidateNamingSnapshot
        namingSnapshotStartedHere = False
    End If
    Set db = ResolveCatalogDb()
    db.Relations.Refresh
    On Error Resume Next
    Application.SysCmd acSysCmdSetStatus, "Deleting ADC relationships..."
    On Error GoTo ErrHandler
    For Each Definition In definitions
        relationshipName = Trim$(CStr(Definition(0)))
        primaryTable = Trim$(CStr(Definition(1)))
        foreignTable = Trim$(CStr(Definition(3)))
        If tableNames.Exists(primaryTable) And tableNames.Exists(foreignTable) Then
            If RelationshipExistsByName(db, relationshipName) Then
                plannedCount = plannedCount + 1
                If RelationshipMatchesDeletionDefinition(db.Relations(relationshipName), primaryTable, foreignTable) Then
                    DeleteSystemOwnedRelationshipSnapshotItem db, relationshipName, primaryTable, foreignTable, attemptedCount, deletedCount, skippedCount, failedCount, detailText, failureDetailText
                Else
                    failedCount = failedCount + 1
                    failureDetailText = SharedText_AppendNonBlankLine(failureDetailText, "Refused relationship teardown because the live relationship endpoints do not match the canonical ADC definition: " & relationshipName)
                    LogTraceMessage TRACE_LEVEL_ERROR, "DeleteObjects", "DeleteSystemOwnedRelationshipsForSnapshot", "DELETE RELATIONSHIP FAILED", "Name=" & relationshipName & vbCrLf & "ExpectedPrimaryTable=" & primaryTable & vbCrLf & "ExpectedForeignTable=" & foreignTable & vbCrLf & "ActualPrimaryTable=" & db.Relations(relationshipName).Table & vbCrLf & "ActualForeignTable=" & db.Relations(relationshipName).foreignTable & vbCrLf & "Error=0" & vbCrLf & "Source=DeleteSystemOwnedRelationshipsForSnapshot" & vbCrLf & "Description=Live relationship endpoints do not match the canonical ADC relationship definition; relationship was preserved.", "Macro 65 ADC relationship ownership mismatch", 0, "DeleteSystemOwnedRelationshipsForSnapshot", "Live relationship endpoints do not match the canonical ADC relationship definition; relationship was preserved.", relationshipName, "Relationship", "CodeDb", "DeleteSystemOwnedObjects"
                End If
            End If
        End If
    Next Definition
    'Refresh once after the successful batch instead of once per relationship; delete-error recovery retains its defensive refresh.
    db.Relations.Refresh
    LogTraceMessage IIf(failedCount = 0, TRACE_LEVEL_INFO, TRACE_LEVEL_ERROR), "DeleteObjects", "DeleteSystemOwnedRelationshipsForSnapshot", "DELETE RELATIONSHIP SUMMARY", "Planned=" & CStr(plannedCount) & vbCrLf & "Attempted=" & CStr(attemptedCount) & vbCrLf & "Deleted=" & CStr(deletedCount) & vbCrLf & "Skipped=" & CStr(skippedCount) & vbCrLf & "Failed=" & CStr(failedCount), "Macro 65 ADC relationship teardown summary", 0, vbNullString, vbNullString, vbNullString, "RelationshipSet", "CodeDb", "DeleteSystemOwnedObjects"
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    On Error Resume Next
    If namingSnapshotStartedHere Then InvalidateNamingSnapshot
    On Error GoTo 0
    Err.Raise savedErrNumber, "DeleteSystemOwnedRelationshipsForSnapshot", "Could not prepare or execute the ADC-owned relationship teardown from the retained deletion snapshot. " & savedErrDescription
End Sub

Private Sub DeleteSystemOwnedRelationshipSnapshotItem(ByVal db As DAO.Database, ByVal relationshipName As String, ByVal primaryTable As String, ByVal foreignTable As String, ByRef attemptedCount As Long, ByRef deletedCount As Long, ByRef skippedCount As Long, ByRef failedCount As Long, ByRef detailText As String, ByRef failureDetailText As String)
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim refreshErrNumber As Long
    Dim refreshErrSource As String
    Dim refreshErrDescription As String
    Dim failureText As String
    attemptedCount = attemptedCount + 1
    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "DeleteSystemOwnedRelationshipSnapshotItem", "DELETE RELATIONSHIP ATTEMPT", "Name=" & relationshipName & vbCrLf & "PrimaryTable=" & primaryTable & vbCrLf & "ForeignTable=" & foreignTable, "Macro 65 ADC relationship delete attempt", 0, vbNullString, vbNullString, relationshipName, "Relationship", "CodeDb", "DeleteSystemOwnedObjects"
    On Error GoTo DeleteFailed
    db.Relations.Delete relationshipName
    deletedCount = deletedCount + 1
    detailText = SharedText_AppendNonBlankLine(detailText, "Deleted relationship: " & relationshipName & " (" & primaryTable & " -> " & foreignTable & ")")
    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "DeleteSystemOwnedRelationshipSnapshotItem", "DELETE RELATIONSHIP SUCCESS", "Name=" & relationshipName & vbCrLf & "PrimaryTable=" & primaryTable & vbCrLf & "ForeignTable=" & foreignTable, "Macro 65 ADC relationship delete success", 0, vbNullString, vbNullString, relationshipName, "Relationship", "CodeDb", "DeleteSystemOwnedObjects"
    Exit Sub
DeleteFailed:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If savedErrNumber = 3011 Or savedErrNumber = 3265 Then
        Err.Clear
        On Error Resume Next
        db.Relations.Refresh
        refreshErrNumber = Err.Number
        refreshErrSource = Err.Source
        refreshErrDescription = Err.Description
        Err.Clear
        On Error GoTo 0
        If refreshErrNumber = 0 Then
            skippedCount = skippedCount + 1
            detailText = SharedText_AppendNonBlankLine(detailText, "Skipped relationship already unavailable: " & relationshipName)
            LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "DeleteSystemOwnedRelationshipSnapshotItem", "DELETE RELATIONSHIP SKIP", "Name=" & relationshipName & vbCrLf & "PrimaryTable=" & primaryTable & vbCrLf & "ForeignTable=" & foreignTable, "Macro 65 ADC relationship already unavailable", savedErrNumber, savedErrSource, savedErrDescription, relationshipName, "Relationship", "CodeDb", "DeleteSystemOwnedObjects"
            Exit Sub
        End If
        savedErrNumber = refreshErrNumber
        savedErrSource = refreshErrSource
        savedErrDescription = "Relationship delete reported the relationship already unavailable, but the DAO.Relations collection could not be refreshed afterward. " & refreshErrDescription
    End If
    failedCount = failedCount + 1
    failureText = "Failed relationship: " & relationshipName & " (" & primaryTable & " -> " & foreignTable & "); Error=" & CStr(savedErrNumber) & "; Source=" & savedErrSource & "; Description=" & savedErrDescription
    detailText = SharedText_AppendNonBlankLine(detailText, failureText)
    failureDetailText = SharedText_AppendNonBlankLine(failureDetailText, failureText)
    LogTraceMessage TRACE_LEVEL_ERROR, "DeleteObjects", "DeleteSystemOwnedRelationshipSnapshotItem", "DELETE RELATIONSHIP FAILED", "Name=" & relationshipName & vbCrLf & "PrimaryTable=" & primaryTable & vbCrLf & "ForeignTable=" & foreignTable & vbCrLf & "Error=" & CStr(savedErrNumber) & vbCrLf & "Source=" & savedErrSource & vbCrLf & "Description=" & savedErrDescription, "Macro 65 ADC relationship delete failure", savedErrNumber, savedErrSource, savedErrDescription, relationshipName, "Relationship", "CodeDb", "DeleteSystemOwnedObjects"
End Sub

Private Function RelationshipMatchesDeletionDefinition(ByVal relationship As DAO.Relation, ByVal primaryTable As String, ByVal foreignTable As String) As Boolean
    If relationship Is Nothing Then Exit Function
    RelationshipMatchesDeletionDefinition = (StrComp(Trim$(relationship.Table), primaryTable, vbTextCompare) = 0 And StrComp(Trim$(relationship.foreignTable), foreignTable, vbTextCompare) = 0)
End Function

Private Function RelationshipExistsByName(ByVal db As DAO.Database, ByVal relationshipName As String) As Boolean
    On Error GoTo NotFound
    Dim rel As DAO.Relation
    Set rel = db.Relations(relationshipName)
    RelationshipExistsByName = True
    Exit Function
NotFound:
    If Err.Number = 3011 Or Err.Number = 3265 Then
        Err.Clear
        Exit Function
    End If
    Err.Raise Err.Number, "RelationshipExistsByName", Err.Description
End Function

Private Sub VerifySystemOwnedDeletionSurvivors(ByVal objectsToDelete As Collection, ByVal macroDeletionPlan As Collection, ByVal formDeletionPlan As Collection, ByRef survivorCount As Long, ByRef survivorDetailText As String)
    On Error GoTo ErrHandler
    Dim item As Variant
    Dim objectName As String
    Dim objectTypeText As String
    survivorCount = 0
    survivorDetailText = vbNullString
    If Not objectsToDelete Is Nothing Then
        For Each item In objectsToDelete
            objectName = Trim$(CStr(item(0)))
            objectTypeText = Trim$(CStr(item(1)))
            If DeletionSnapshotObjectExists(objectName, objectTypeText) Then RecordDeleteSurvivor objectName, objectTypeText, survivorCount, survivorDetailText
        Next item
    End If
    If Not macroDeletionPlan Is Nothing Then
        For Each item In macroDeletionPlan
            objectName = Trim$(CStr(item))
            If AccessObjectExistsInCurrentDatabase(acMacro, objectName) Then
                If CoreIsDeferredInvokingMacro65Name(objectName) Then
                    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "VerifySystemOwnedDeletionSurvivors", "DELETE VERIFY DEFERRED", "Type=Macro; Name=" & objectName & "; Reason=Invoking Macro 65 is scheduled for post-RunCode deletion", "Macro 65 terminal survivor verification", 0, vbNullString, vbNullString, objectName, ADCAT_OBJECT_TYPE_MACRO, "CodeDb", "DeleteSystemOwnedObjects"
                Else
                    RecordDeleteSurvivor objectName, ADCAT_OBJECT_TYPE_MACRO, survivorCount, survivorDetailText
                End If
            End If
        Next item
    End If
    If Not formDeletionPlan Is Nothing Then
        For Each item In formDeletionPlan
            objectName = Trim$(CStr(item))
            If AccessObjectExistsInCurrentDatabase(acForm, objectName) Then RecordDeleteSurvivor objectName, ADCAT_OBJECT_TYPE_FORM, survivorCount, survivorDetailText
        Next item
    End If
    LogTraceMessage IIf(survivorCount = 0, TRACE_LEVEL_INFO, TRACE_LEVEL_ERROR), "DeleteObjects", "VerifySystemOwnedDeletionSurvivors", "DELETE VERIFY COMPLETE", "Survivors=" & CStr(survivorCount), "Macro 65 terminal survivor verification from retained deletion plans", 0, vbNullString, vbNullString, vbNullString, "ManagedObjectSet", "CodeDb", "DeleteSystemOwnedObjects"
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "VerifySystemOwnedDeletionSurvivors", "Could not perform terminal survivor verification from the retained deletion plans. " & Err.Description
End Sub

Private Function DeletionSnapshotObjectExists(ByVal objectName As String, ByVal objectTypeText As String) As Boolean
    Select Case UCase$(objectTypeText)
        Case UCase$(ADCAT_OBJECT_TYPE_TABLE)
            DeletionSnapshotObjectExists = AccessObjectExistsInCurrentDatabase(acTable, objectName)
        Case UCase$(ADCAT_OBJECT_TYPE_QUERY)
            DeletionSnapshotObjectExists = AccessObjectExistsInCurrentDatabase(acQuery, objectName)
        Case UCase$(ADCAT_OBJECT_TYPE_REPORT)
            DeletionSnapshotObjectExists = AccessObjectExistsInCurrentDatabase(acReport, objectName)
        Case UCase$(ADCAT_OBJECT_TYPE_FORM)
            DeletionSnapshotObjectExists = AccessObjectExistsInCurrentDatabase(acForm, objectName)
        Case UCase$(ADCAT_OBJECT_TYPE_MACRO)
            DeletionSnapshotObjectExists = AccessObjectExistsInCurrentDatabase(acMacro, objectName)
        Case Else
            Err.Raise vbObjectError + 6311, "DeletionSnapshotObjectExists", "Unsupported object type in retained deletion verification: " & objectTypeText
    End Select
End Function

Private Sub RecordDeleteSurvivor(ByVal objectName As String, ByVal objectTypeText As String, ByRef survivorCount As Long, ByRef survivorDetailText As String)
    survivorCount = survivorCount + 1
    survivorDetailText = SharedText_AppendNonBlankLine(survivorDetailText, "Survivor " & objectTypeText & ": " & objectName)
    LogTraceMessage TRACE_LEVEL_ERROR, "DeleteObjects", "RecordDeleteSurvivor", "DELETE VERIFY SURVIVOR", "Type=" & objectTypeText & vbCrLf & "Name=" & objectName, "Macro 65 terminal survivor verification", 0, vbNullString, vbNullString, objectName, objectTypeText, "CodeDb", "DeleteSystemOwnedObjects"
End Sub

Private Sub DeleteSystemOwnedObjectSnapshotItem(ByVal objectName As String, ByVal objectTypeText As String, ByRef deletedCount As Long, ByRef skippedCount As Long, ByRef failedCount As Long, ByRef detailText As String, ByRef failureDetailText As String)
    Dim errorText As String
    Dim failureText As String
    Dim deletionErrNumber As Long
    Dim deletionErrSource As String
    Dim deletionErrDescription As String
    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "DeleteSystemOwnedObjectSnapshotItem", "DELETE ATTEMPT", "Type=" & objectTypeText & "; Name=" & objectName, "Macro 65 catalog delete attempt", 0, vbNullString, vbNullString, objectName, objectTypeText, "CodeDb", "DeleteSystemOwnedObjects"
    On Error Resume Next
    Application.SysCmd acSysCmdSetStatus, "Deleting " & objectTypeText & ": " & objectName
    On Error GoTo 0
    errorText = vbNullString
    deletionErrNumber = 0
    deletionErrSource = vbNullString
    deletionErrDescription = vbNullString
    If DeleteOneSystemOwnedObject(objectName, objectTypeText, errorText, deletionErrNumber, deletionErrSource, deletionErrDescription) Then
        deletedCount = deletedCount + 1
        detailText = SharedText_AppendNonBlankLine(detailText, "Deleted " & objectTypeText & ": " & objectName)
        LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "DeleteSystemOwnedObjectSnapshotItem", "DELETE SUCCESS", "Type=" & objectTypeText & "; Name=" & objectName, "Macro 65 catalog delete success", 0, vbNullString, vbNullString, objectName, objectTypeText, "CodeDb", "DeleteSystemOwnedObjects"
    ElseIf Len(Trim$(errorText)) = 0 Then
        skippedCount = skippedCount + 1
        detailText = SharedText_AppendNonBlankLine(detailText, "Skipped " & objectTypeText & ": " & objectName)
        LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "DeleteSystemOwnedObjectSnapshotItem", "DELETE SKIP", "Type=" & objectTypeText & "; Name=" & objectName & "; Reason=Deletion owner reported no deletion and no error", "Macro 65 catalog delete skip", 0, vbNullString, vbNullString, objectName, objectTypeText, "CodeDb", "DeleteSystemOwnedObjects"
    ElseIf Left$(errorText, 8) = "Skipped:" Or InStr(1, errorText, "already unavailable", vbTextCompare) > 0 Then
        skippedCount = skippedCount + 1
        detailText = SharedText_AppendNonBlankLine(detailText, errorText)
        LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "DeleteSystemOwnedObjectSnapshotItem", "DELETE SKIP", "Type=" & objectTypeText & "; Name=" & objectName & "; Reason=" & errorText, "Macro 65 catalog delete skip", 0, vbNullString, vbNullString, objectName, objectTypeText, "CodeDb", "DeleteSystemOwnedObjects"
    Else
        failedCount = failedCount + 1
        failureText = "Failed " & objectTypeText & ": " & objectName & "; " & errorText
        detailText = SharedText_AppendNonBlankLine(detailText, failureText)
        failureDetailText = SharedText_AppendNonBlankLine(failureDetailText, failureText)
        LogTraceMessage TRACE_LEVEL_ERROR, "DeleteObjects", "DeleteSystemOwnedObjectSnapshotItem", "DELETE FAILED", "Type=" & objectTypeText & "; Name=" & objectName & "; Error=" & CStr(deletionErrNumber) & "; Source=" & deletionErrSource & "; Description=" & deletionErrDescription, "Macro 65 catalog delete failure", deletionErrNumber, deletionErrSource, deletionErrDescription, objectName, objectTypeText, "CodeDb", "DeleteSystemOwnedObjects"
    End If
    PumpSetupResponsiveness
End Sub

Private Function SystemOwnedObjectDeletionSnapshot(ByVal overallPreflightTimer As Single, ByVal deletionScope As ObjectDeletionScope) As Collection
    On Error GoTo ErrHandler
    Dim rows As Collection
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim objectTypeFilter As String
    Dim objectName As String
    Dim objectTypeText As String
    Dim managedObjectsTableName As String
    Dim snapshotTimer As Single
    Dim boundaryTimer As Single
    Dim enumerationTimer As Single
    Dim lastProgressTimer As Single
    Dim rowsScanned As Long
    Dim rowsRetained As Long
    Dim boundaryElapsed As Double
    Dim enumerationElapsed As Double
    Dim snapshotElapsed As Double
    ' INTENTIONAL OPERATION SNAPSHOT: rows captures one coherent source state for this invocation of SystemOwnedObjectDeletionSnapshot; validity ends with the operation rather than becoming a live data authority.
    Set rows = New Collection
    snapshotTimer = Timer
    LogTraceMessage "Info", "ManagedObjectCatalog", "SystemOwnedObjectDeletionSnapshot", "Deletion preflight | snapshot | starting", "Total elapsed preflight: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(overallPreflightTimer)) & vbCrLf & "Naming diagnostics: " & NamingPerformanceSummary(), "Deletion preflight snapshot", 0, vbNullString, vbNullString, vbNullString, vbNullString, "CodeDb", "DeleteSystemOwnedObjects"
    CorePulseProgress "Preparing system-owned object deletion | building deletion snapshot"
    LogTraceMessage "Info", "ManagedObjectCatalog", "SystemOwnedObjectDeletionSnapshot", "Deletion preflight | snapshot | resolving MAC table", vbNullString, "Deletion preflight snapshot", 0, vbNullString, vbNullString, vbNullString, "Table", "CodeDb", "DeleteSystemOwnedObjects"
    boundaryTimer = Timer
    managedObjectsTableName = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID)
    boundaryElapsed = ElapsedSecondsFromTimer(boundaryTimer)
    LogTraceMessage "Info", "ManagedObjectCatalog", "SystemOwnedObjectDeletionSnapshot", "Deletion preflight | snapshot | MAC table resolved", "Resolved table: " & managedObjectsTableName & vbCrLf & "Boundary elapsed: " & FormattedElapsedSeconds(boundaryElapsed) & vbCrLf & "Snapshot elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(snapshotTimer)) & vbCrLf & "Overall preflight elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(overallPreflightTimer)) & vbCrLf & "Naming diagnostics: " & NamingPerformanceSummary(), "Deletion preflight snapshot", 0, vbNullString, vbNullString, managedObjectsTableName, "Table", "CodeDb", "DeleteSystemOwnedObjects"
    sqlText = "SELECT m." & SQL_FIELD_OBJECT_NAME & ", m." & SQL_FIELD_OBJECT_TYPE & ", o.[" & FIELD_TARGET_OBJECT_NAME_OVERRIDE & "] AS [" & FIELD_TARGET_OBJECT_NAME_OVERRIDE & "], m." & SQL_FIELD_MANAGED_OBJECT_STATUS & ", m." & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER & " FROM [" & managedObjectsTableName & "] AS m LEFT JOIN [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECT_OVERRIDES_TABLE_LOGICAL_ID) & "] AS o ON m." & SQL_FIELD_ID & "=o.[" & FIELD_MANAGED_OBJECT_ID & "] "
    sqlText = sqlText & "WHERE Len(Nz(m." & SQL_FIELD_OBJECT_NAME & ",''))>0 "
    sqlText = sqlText & "AND Nz(m." & SQL_FIELD_MANAGED_OBJECT_STATUS & "," & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & ")=" & SharedDao_SqlQuote(ADCAT_STATUS_ACTIVE) & " "
    sqlText = sqlText & "AND UCase(Trim(Nz(m.[" & FIELD_CATALOG_OWNER & "],''))) IN (" & SharedDao_SqlQuote(UCase$(ADCAT_OWNER_SYSTEM_DEFAULT)) & "," & SharedDao_SqlQuote(UCase$(DEFAULT_CATALOG_OWNER_GENERATED)) & ") "
    Select Case deletionScope
        Case DELETE_SCOPE_TABLES
            objectTypeFilter = SharedDao_SqlQuote(UCase$(ADCAT_OBJECT_TYPE_TABLE))
        Case DELETE_SCOPE_QUERIES
            objectTypeFilter = SharedDao_SqlQuote(UCase$(ADCAT_OBJECT_TYPE_QUERY))
        Case DELETE_SCOPE_REPORTS
            objectTypeFilter = SharedDao_SqlQuote(UCase$(ADCAT_OBJECT_TYPE_REPORT))
        Case DELETE_SCOPE_ALL
            objectTypeFilter = SharedDao_SqlQuote(UCase$(ADCAT_OBJECT_TYPE_REPORT)) & "," & SharedDao_SqlQuote(UCase$(ADCAT_OBJECT_TYPE_QUERY)) & "," & SharedDao_SqlQuote(UCase$(ADCAT_OBJECT_TYPE_TABLE))
        Case Else
            Err.Raise vbObjectError + 6308, "SystemOwnedObjectDeletionSnapshot", "Unsupported catalog snapshot scope: " & CStr(CLng(deletionScope))
    End Select
    sqlText = sqlText & "AND UCase(Trim(Nz(m." & SQL_FIELD_OBJECT_TYPE & ",''))) IN (" & objectTypeFilter & ") "
    sqlText = sqlText & "ORDER BY m." & SQL_FIELD_OBJECT_TYPE & ", m." & SQL_FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER & ", m." & SQL_FIELD_OBJECT_NAME & ";"
    LogTraceMessage "Info", "ManagedObjectCatalog", "SystemOwnedObjectDeletionSnapshot", "Deletion preflight | snapshot | opening recordset", "Table: " & managedObjectsTableName & vbCrLf & "Snapshot elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(snapshotTimer)), "Deletion preflight snapshot", 0, vbNullString, vbNullString, managedObjectsTableName, "Table", "CodeDb", "DeleteSystemOwnedObjects"
    boundaryTimer = Timer
    Set rs = ResolveCatalogDb().OpenRecordset(sqlText, dbOpenSnapshot)
    boundaryElapsed = ElapsedSecondsFromTimer(boundaryTimer)
    LogTraceMessage "Info", "ManagedObjectCatalog", "SystemOwnedObjectDeletionSnapshot", "Deletion preflight | snapshot | recordset opened", "Recordset-open elapsed: " & FormattedElapsedSeconds(boundaryElapsed) & vbCrLf & "Snapshot elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(snapshotTimer)) & vbCrLf & "Overall preflight elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(overallPreflightTimer)) & vbCrLf & "Naming diagnostics: " & NamingPerformanceSummary(), "Deletion preflight snapshot", 0, vbNullString, vbNullString, managedObjectsTableName, "Table", "CodeDb", "DeleteSystemOwnedObjects"
    enumerationTimer = Timer
    lastProgressTimer = enumerationTimer
    LogTraceMessage "Info", "ManagedObjectCatalog", "SystemOwnedObjectDeletionSnapshot", "Deletion preflight | snapshot | enumerating rows", "Snapshot elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(snapshotTimer)), "Deletion preflight snapshot", 0, vbNullString, vbNullString, managedObjectsTableName, "Table", "CodeDb", "DeleteSystemOwnedObjects"
    Do While rs.EOF = False
        PumpSetupResponsiveness
        rowsScanned = rowsScanned + 1
        objectName = CStr(Nz(rs.fields(FIELD_TARGET_OBJECT_NAME_OVERRIDE).value, vbNullString))
        If Len(Trim$(objectName)) = 0 Then objectName = CStr(rs.fields(FIELD_OBJECT_NAME).value)
        objectTypeText = CStr(rs.fields(FIELD_OBJECT_TYPE).value)
        If Len(Trim$(objectName)) > 0 Then
            If deletionScope = DELETE_SCOPE_ALL And ObjectNameIsInRetainedTraceLogNamespace(objectName) Then
                LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "SystemOwnedObjectDeletionSnapshot", "DELETE PLAN EXCLUDE", "Name=" & objectName & "; Type=" & objectTypeText & "; Classification=TraceLogNamespace; Reason=Retained by Delete All policy", "TraceLog namespace is intentionally excluded from the catalog-backed Delete All snapshot.", 0, vbNullString, vbNullString, objectName, objectTypeText, "CodeDb", "SystemMacro65"
            Else
                rows.Add Array(objectName, objectTypeText)
                rowsRetained = rowsRetained + 1
            End If
        End If
        If ElapsedSecondsFromTimer(lastProgressTimer) >= DELETE_PREFLIGHT_DIAGNOSTIC_INTERVAL_SECONDS Then
            enumerationElapsed = ElapsedSecondsFromTimer(enumerationTimer)
            CorePulseProgress "Preparing system-owned object deletion | snapshot " & CStr(rowsScanned) & " rows scanned"
            LogTraceMessage "Info", "ManagedObjectCatalog", "SystemOwnedObjectDeletionSnapshot", "Deletion preflight | snapshot enumeration", "Rows scanned: " & CStr(rowsScanned) & vbCrLf & "Rows retained: " & CStr(rowsRetained) & vbCrLf & "Current object type: " & objectTypeText & vbCrLf & "Current object: " & objectName & vbCrLf & "Stage elapsed: " & FormattedElapsedSeconds(enumerationElapsed) & vbCrLf & "Overall preflight elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(overallPreflightTimer)) & vbCrLf & "Naming diagnostics: " & NamingPerformanceSummary(), "Deletion preflight snapshot", 0, vbNullString, vbNullString, objectName, objectTypeText, "CodeDb", "DeleteSystemOwnedObjects"
            lastProgressTimer = Timer
        End If
        rs.MoveNext
    Loop
    enumerationElapsed = ElapsedSecondsFromTimer(enumerationTimer)
    LogTraceMessage "Info", "ManagedObjectCatalog", "SystemOwnedObjectDeletionSnapshot", "Deletion preflight | snapshot | enumeration complete", "Rows scanned: " & CStr(rowsScanned) & vbCrLf & "Rows retained: " & CStr(rowsRetained) & vbCrLf & "Enumeration elapsed: " & FormattedElapsedSeconds(enumerationElapsed) & vbCrLf & "Overall preflight elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(overallPreflightTimer)) & vbCrLf & "Naming diagnostics: " & NamingPerformanceSummary(), "Deletion preflight snapshot", 0, vbNullString, vbNullString, managedObjectsTableName, "Table", "CodeDb", "DeleteSystemOwnedObjects"
    rs.Close
    snapshotElapsed = ElapsedSecondsFromTimer(snapshotTimer)
    LogTraceMessage "Info", "ManagedObjectCatalog", "SystemOwnedObjectDeletionSnapshot", "Deletion preflight | snapshot | complete", "Rows scanned: " & CStr(rowsScanned) & vbCrLf & "Rows added to deletion collection: " & CStr(rowsRetained) & vbCrLf & "Elapsed: " & FormattedElapsedSeconds(snapshotElapsed) & vbCrLf & "Overall preflight elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(overallPreflightTimer)) & vbCrLf & "Naming diagnostics: " & NamingPerformanceSummary(), "Deletion preflight snapshot", 0, vbNullString, vbNullString, managedObjectsTableName, "Table", "CodeDb", "DeleteSystemOwnedObjects"
    Set SystemOwnedObjectDeletionSnapshot = rows
    Exit Function
ErrHandler:
    Dim handlerErrNumber As Long
    Dim handlerErrDescription As String
    handlerErrNumber = Err.Number
    handlerErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
    Err.Raise handlerErrNumber, "SystemOwnedObjectDeletionSnapshot", "Unable to build a snapshot of the selected active system-owned/generated catalog objects before deletion." & vbCrLf & "Table: " & IIf(Len(managedObjectsTableName) > 0, managedObjectsTableName, "(unresolved)") & vbCrLf & "Description: " & handlerErrDescription
End Function

Private Function DeleteOneSystemOwnedObject(ByVal objectName As String, ByVal objectTypeText As String, ByRef errorText As String, ByRef errorNumberResult As Long, ByRef errorSourceResult As String, ByRef errorDescriptionResult As String) As Boolean
    On Error GoTo ErrHandler
    Dim objectType As AcObjectType
    Dim db As DAO.Database
    objectName = Trim$(objectName)
    objectTypeText = Trim$(objectTypeText)
    errorText = vbNullString
    errorNumberResult = 0
    errorSourceResult = vbNullString
    errorDescriptionResult = vbNullString
    If Len(objectName) = 0 Then
        errorText = "Skipped: MAC row has a blank object name."
        Exit Function
    End If
    If TryAccessObjectTypeFromCatalogText(objectTypeText, objectType) = False Then
        errorText = "Skipped: unsupported managed object type " & objectTypeText & " for object " & objectName & "."
        Exit Function
    End If
    If objectType = acTable Then
        Set db = ResolveCatalogDb()
        If TryDeleteTableDef(db, objectName, errorText, errorNumberResult, errorSourceResult, errorDescriptionResult) Then errorText = vbNullString
        DeleteOneSystemOwnedObject = (Len(Trim$(errorText)) = 0)
    ElseIf objectType = acQuery Then
        Set db = ResolveCatalogDb()
        If TryDeleteQueryDef(db, objectName, errorText, errorNumberResult, errorSourceResult, errorDescriptionResult) Then errorText = vbNullString
        DeleteOneSystemOwnedObject = (Len(Trim$(errorText)) = 0)
    Else
        If objectType = acForm Then
            If AccessObjectExistsInCurrentDatabase(acForm, objectName) Then
                If CurrentProject.AllForms(objectName).IsLoaded Then DoCmd.Close acForm, objectName, acSaveNo
            End If
        ElseIf objectType = acReport Then
            If AccessObjectExistsInCurrentDatabase(acReport, objectName) Then
                If CurrentProject.AllReports(objectName).IsLoaded Then DoCmd.Close acReport, objectName, acSaveNo
            End If
        End If
        DeleteOneSystemOwnedObject = TryDeleteAccessObject(objectType, objectName, errorText, errorNumberResult, errorSourceResult, errorDescriptionResult)
    End If
    Exit Function
ErrHandler:
    errorNumberResult = Err.Number
    errorSourceResult = Err.Source
    errorDescriptionResult = Err.Description
    errorText = "Error " & CStr(errorNumberResult) & "; Source: " & errorSourceResult & "; Description: " & errorDescriptionResult
End Function

Private Function DeleteSystemOwnedObjectsDetailsForMessage(ByVal detailText As String) As String
    detailText = Trim$(detailText)
    If Len(detailText) = 0 Then Exit Function
    If Len(detailText) > 3500 Then detailText = Left$(detailText, 3500) & vbCrLf & "...details truncated."
    DeleteSystemOwnedObjectsDetailsForMessage = vbCrLf & vbCrLf & "Details:" & vbCrLf & detailText
End Function

' Purpose: Applies configured naming rules transactionally to ADC-managed physical objects, repairs dependent references, SQL, and lookup sources, verifies the renamed state, and rolls back on failure.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
' Returns: Boolean: True on the function's affirmative/success branch for apply object naming rules; False on its explicit negative, unavailable, rejected, or handled-failure branch.
Public Function ApplyObjectNamingRules(Optional ByVal showMessage As Boolean = True) As Boolean
    RequireNoUnfinishedCatalogRecovery "ApplyObjectNamingRules"
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim renames As Collection
    Dim physicalRenames As Collection
    Dim reversedRenames As Collection
    Dim ws As DAO.workspace
    Dim transactionStarted As Boolean
    Dim changedCount As Long
    Dim referenceChangeCount As Long
    Dim settingsNameBefore As String
    Dim managedObjectsNameBefore As String
    Dim settingsNameAfter As String
    Dim managedObjectsNameAfter As String
    Set db = ResolveCatalogDb()
    Set renames = ManagedObjectRenameMap()
    If renames.count = 0 Then
        ValidateTerminologyCanonicalNaming
        ApplyObjectNamingRules = True
        If showMessage Then ReportCoreActionSuccess "ApplyObjectNamingRules", "ObjectNaming", "ADC-managed physical object names already match the configured naming rules.", "Objects renamed: 0", "Apply ADC object naming rules", vbNullString, "ManagedObjectSet", "CodeDb", "ObjectNaming", True, "Apply ADC Object Naming Rules", vbInformation, True, "Tables", False
        Exit Function
    End If
    settingsNameBefore = BootstrapSettingsTableName()
    managedObjectsNameBefore = BootstrapManagedObjectsTableName()
    Set physicalRenames = PhysicalManagedObjectRenameMap(renames)
    PreflightManagedObjectRenames db, physicalRenames, "CodeDb"
    If physicalRenames.count > 0 Then PrepareOpenObjectsForManagedNameUpdate db, db, physicalRenames
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    transactionStarted = True
    changedCount = RenameManagedObjectsInDatabase(db, physicalRenames)
    referenceChangeCount = RepairRenamedObjectReferences(db, renames)
    referenceChangeCount = referenceChangeCount + RepairRenamedQuerySql(db, renames)
    referenceChangeCount = referenceChangeCount + RepairRenamedLookupRowSources(db, renames)
    settingsNameAfter = ConfiguredRenameTarget(renames, "Table", settingsNameBefore)
    managedObjectsNameAfter = ConfiguredRenameTarget(renames, "Table", managedObjectsNameBefore)
    SetNamingBootstrapTableNames settingsNameAfter, managedObjectsNameAfter
    VerifyManagedObjectRenames physicalRenames
    ValidateTerminologyCanonicalNaming
    ws.CommitTrans
    transactionStarted = False
    ApplyObjectNamingRules = True
    If showMessage Then ReportCoreActionSuccess "ApplyObjectNamingRules", "ObjectNaming", "ADC-managed physical object names were updated successfully.", "Physical objects renamed: " & CStr(changedCount) & vbCrLf & "Managed references repaired: " & CStr(referenceChangeCount), "Apply ADC object naming rules", vbNullString, "ManagedObjectSet", "CodeDb", "ObjectNaming", True, "Apply ADC Object Naming Rules", vbInformation, True, "Tables", False
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim rollbackText As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If transactionStarted Then ws.Rollback
    SetNamingBootstrapTableNames settingsNameBefore, managedObjectsNameBefore
    Set reversedRenames = ReverseRenameMap(physicalRenames)
    RenameManagedObjectsInDatabase db, reversedRenames
    If Err.Number <> 0 Then rollbackText = "Physical-name rollback error " & CStr(Err.Number) & ": " & Err.Description: Err.Clear
    On Error GoTo 0
    If Len(rollbackText) > 0 Then savedErrDescription = savedErrDescription & vbCrLf & vbCrLf & rollbackText
    If showMessage Then
        CoreShowErrorMessage "Could not apply ADC object naming rules to managed physical object names. The operation stopped and a rollback was attempted. Original physical names may have been restored only partially; review the error details and current object names before retrying." & vbCrLf & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description:" & vbCrLf & savedErrDescription, vbCritical, "Apply ADC Object Naming Rules", "ApplyObjectNamingRules", "Tables"
        Exit Function
    End If
    Err.Raise savedErrNumber, "ApplyObjectNamingRules", savedErrDescription
End Function

Private Sub ValidateTerminologyCanonicalNaming()
    Dim db As DAO.Database
    Set db = ResolveCatalogDb()
    If DaoTableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_TERMINOLOGY_TABLE_LOGICAL_ID)) = False Then Err.Raise vbObjectError + 8969, "ValidateTerminologyCanonicalNaming", "The canonical Help_Terms table is missing." & vbCrLf & "Expected table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_TERMINOLOGY_TABLE_LOGICAL_ID)
End Sub

Private Function ManagedObjectRenameMap() As Collection
    ' INTENTIONAL FILTERED PARTITION: items is a consumer-required subset/projection derived from the source enumerated by ManagedObjectRenameMap; its membership is not independently authoritative.
    Dim items As New Collection
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim currentName As String
    Dim targetName As String
    Dim objectTypeCode As String
    Dim namingCategory As String
    Set db = ResolveCatalogDb()
    Set rs = db.OpenRecordset("SELECT " & SQL_FIELD_OBJECT_NAME & "," & SQL_FIELD_OBJECT_TYPE & "," & SQL_FIELD_LOGICAL_OBJECT_NAME & "," & SQL_FIELD_OBJECT_NAMING_CATEGORY & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID) & "] WHERE Nz(" & SQL_FIELD_MANAGED_OBJECT_STATUS & ",'Active')='Active' AND Len(Nz(" & SQL_FIELD_LOGICAL_OBJECT_NAME & ",''))>0 AND Len(Nz(" & SQL_FIELD_OBJECT_NAMING_CATEGORY & ",''))>0 ORDER BY IIf(" & SQL_FIELD_LOGICAL_OBJECT_NAME & "=" & SharedDao_SqlQuote(DEFAULT_MANAGED_OBJECTS_TABLE_LOGICAL_ID) & ",1,0), " & SQL_FIELD_OBJECT_NAMING_CATEGORY & "," & SQL_FIELD_LOGICAL_OBJECT_NAME & ";", dbOpenSnapshot)
    Do While rs.EOF = False
        PumpSetupResponsiveness
        currentName = CStr(rs!objectName)
        objectTypeCode = CStr(rs!objectType)
        namingCategory = CStr(rs.fields(FIELD_OBJECT_NAMING_CATEGORY))
        targetName = TargetManagedObjectPhysicalName(namingCategory, CStr(rs!logicalObjectName))
        If StrComp(currentName, targetName, vbBinaryCompare) <> 0 Then AddUniqueManagedRename items, currentName, targetName, objectTypeCode
        rs.MoveNext
    Loop
    rs.Close
    Set ManagedObjectRenameMap = items
End Function

Private Function PhysicalManagedObjectRenameMap(ByVal renames As Collection) As Collection
    ' INTENTIONAL FILTERED PARTITION: physicalRenames is a consumer-required subset/projection derived from the source enumerated by PhysicalManagedObjectRenameMap; its membership is not independently authoritative.
    Dim physicalRenames As New Collection
    Dim item As Variant
    For Each item In renames
        PumpSetupResponsiveness
        If ConfiguredObjectExists(CStr(item(2)), CStr(item(0))) Then physicalRenames.Add item
    Next item
    Set PhysicalManagedObjectRenameMap = physicalRenames
End Function

Private Sub VerifyManagedObjectRenames(ByVal renames As Collection)
    Dim item As Variant
    For Each item In renames
        PumpSetupResponsiveness
        If ConfiguredObjectExists(CStr(item(2)), CStr(item(1))) = False Then Err.Raise vbObjectError + 8948, "VerifyManagedObjectRenames", "A configured target object is missing after the rename batch." & vbCrLf & "Type: " & CStr(item(2)) & vbCrLf & "Target: " & CStr(item(1))
        If ConfiguredObjectExists(CStr(item(2)), CStr(item(0))) Then Err.Raise vbObjectError + 8949, "VerifyManagedObjectRenames", "A source object remains after the rename batch." & vbCrLf & "Type: " & CStr(item(2)) & vbCrLf & "Source: " & CStr(item(0))
    Next item
End Sub

Private Sub PreflightManagedObjectRenames(ByVal db As DAO.Database, ByVal renames As Collection, ByVal databaseRole As String)
    Dim item As Variant
    Dim objectTypeCode As String
    Dim sourceName As String
    Dim targetName As String
    PreflightRenameDestinationSources renames, databaseRole
    For Each item In renames
        PumpSetupResponsiveness
        sourceName = CStr(item(0))
        targetName = CStr(item(1))
        objectTypeCode = CStr(item(2))
        If Len(targetName) = 0 Or Len(targetName) > 64 Then Err.Raise vbObjectError + 8919, "PreflightManagedObjectRenames", "A configured target object name is blank or exceeds 64 characters." & vbCrLf & "Type: " & objectTypeCode & vbCrLf & "Source: " & sourceName & vbCrLf & "Target: " & targetName
        Select Case UCase$(objectTypeCode)
            Case "TABLE": PreflightSingleTableRename db, sourceName, targetName, databaseRole
            Case "QUERY": PreflightSingleQueryRename db, sourceName, targetName, databaseRole
            Case "FORM", "REPORT", "MACRO": PreflightSingleAccessObjectRename objectTypeCode, sourceName, targetName, databaseRole
            Case Else: Err.Raise vbObjectError + 8918, "PreflightManagedObjectRenames", "Unsupported managed physical object type in the rename plan: " & objectTypeCode
        End Select
    Next item
End Sub

Private Function RenameManagedObjectsInDatabase(ByVal db As DAO.Database, ByVal renames As Collection) As Long
    Dim item As Variant
    Dim objectTypeCode As String
    For Each item In renames
        PumpSetupResponsiveness
        objectTypeCode = UCase$(CStr(item(2)))
        If objectTypeCode <> "TABLE" Then RenameManagedObjectsInDatabase = RenameManagedObjectsInDatabase + RenameConfiguredObject(db, CStr(item(2)), CStr(item(0)), CStr(item(1)))
    Next item
    For Each item In renames
        PumpSetupResponsiveness
        If StrComp(CStr(item(2)), "Table", vbTextCompare) = 0 Then RenameManagedObjectsInDatabase = RenameManagedObjectsInDatabase + RenameConfiguredObject(db, CStr(item(2)), CStr(item(0)), CStr(item(1)))
    Next item
End Function

Private Function ManagedObjectRenameFieldAppliesToItem(ByVal fieldName As String, ByVal renameItem As Variant) As Boolean
    Dim objectTypeCode As String
    objectTypeCode = CStr(renameItem(2))
    ManagedObjectRenameFieldAppliesToItem = True
    If StrComp(fieldName, FIELD_FORM_OBJECT_NAME, vbTextCompare) = 0 Then ManagedObjectRenameFieldAppliesToItem = (StrComp(objectTypeCode, "Form", vbTextCompare) = 0): Exit Function
    If StrComp(objectTypeCode, "Form", vbTextCompare) <> 0 Then Exit Function
    Select Case UCase$(Trim$(fieldName))
        Case UCase$(FIELD_OBJECT_DISPLAY_NAME), UCase$(FIELD_RIBBON_TITLE): ManagedObjectRenameFieldAppliesToItem = False
    End Select
End Function

' Purpose: Resolves replace managed object name token from the supplied/current ADC state, using `ObjectNameTokenCharacter` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   sourceText (Input; String): Input source text for replace managed object name token; is assigned to `resultText`.
'   oldName (Input; String): Canonical or physical name identifying the old targeted by this operation.
'   newName (Input; String): Canonical or physical name identifying the new targeted by this operation.
' Returns: String: Returns the proof/state token produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function ReplaceManagedObjectNameToken(ByVal sourceText As String, ByVal oldName As String, ByVal newName As String) As String
    'Replaces case-insensitive object-name tokens only at identifier boundaries while preserving surrounding text.
    Dim resultText As String
    Dim matchPosition As Long
    Dim searchPosition As Long
    Dim afterPosition As Long
    Dim beforeIsBoundary As Boolean
    Dim afterIsBoundary As Boolean
    resultText = sourceText
    searchPosition = 1
    Do
        PumpSetupResponsiveness
        matchPosition = InStr(searchPosition, resultText, oldName, vbTextCompare)
        If matchPosition = 0 Then Exit Do
        beforeIsBoundary = (matchPosition = 1)
        If beforeIsBoundary = False Then beforeIsBoundary = Not ObjectNameTokenCharacter(Mid$(resultText, matchPosition - 1, 1))
        afterPosition = matchPosition + Len(oldName)
        afterIsBoundary = (afterPosition > Len(resultText))
        If afterIsBoundary = False Then afterIsBoundary = Not ObjectNameTokenCharacter(Mid$(resultText, afterPosition, 1))
        If beforeIsBoundary And afterIsBoundary Then
            resultText = Left$(resultText, matchPosition - 1) & newName & Mid$(resultText, afterPosition)
            searchPosition = matchPosition + Len(newName)
        Else
            searchPosition = matchPosition + Len(oldName)
        End If
    Loop
    ReplaceManagedObjectNameToken = resultText
End Function

Private Function ObjectNameTokenCharacter(ByVal characterText As String) As Boolean
    'Returns True when a one-character string is valid inside the object-name token syntax used by token replacement.
    If Len(characterText) = 0 Then Exit Function
    ObjectNameTokenCharacter = (characterText Like "[A-Za-z0-9_]")
End Function

' Purpose: Applies managed object rename map for source text, renames.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   sourceText (Input; String): Input source text for apply managed object rename map; is assigned to `resultText`.
'   renames (Input; Collection): Collection containing the renames items that the routine consumes or updates as part of this operation.
' Returns: String: Returns the physical/logical name resolved for apply managed object rename map under the module's naming/definition rules; vbNullString indicates that no applicable name can be resolved where
'   supported.
Public Function ApplyManagedObjectRenameMap(ByVal sourceText As String, ByVal renames As Collection) As String
    Dim item As Variant
    Dim resultText As String
    resultText = sourceText
    For Each item In renames
        PumpSetupResponsiveness
        resultText = ReplaceManagedObjectNameToken(resultText, CStr(item(0)), CStr(item(1)))
    Next item
    ApplyManagedObjectRenameMap = resultText
End Function


Private Function RepairRenamedFieldRowSource(ByVal fld As DAO.Field, ByVal renames As Collection, ByVal tableName As String) As Long
    On Error GoTo propertyMissing
    Dim originalText As String
    Dim repairedText As String
    originalText = CStr(fld.properties("RowSource").value)
    repairedText = ApplyManagedObjectRenameMap(originalText, renames)
    If StrComp(originalText, repairedText, vbBinaryCompare) <> 0 Then
        If ApplyLookupFieldProperties(fld, vbNullString, repairedText, False, 0, 0, vbNullString, "RepairRenamedFieldRowSource", vbObjectError + 7750, vbObjectError + 7751, LOOKUP_FIELD_PROPERTY_ROW_SOURCE) Then RepairRenamedFieldRowSource = 1
    End If
    Exit Function
propertyMissing:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    If savedErrNumber = 3270 Then Exit Function
    Err.Raise savedErrNumber, "RepairRenamedFieldRowSource", "Could not repair lookup RowSource metadata." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fld.Name & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Sub AddUniqueManagedRename(ByVal renames As Collection, ByVal sourceName As String, ByVal targetName As String, ByVal objectTypeCode As String)
    Dim item As Variant
    sourceName = Trim$(sourceName)
    objectTypeCode = Trim$(objectTypeCode)
    If Len(sourceName) = 0 Or Len(targetName) = 0 Or Len(objectTypeCode) = 0 Then Err.Raise vbObjectError + 8965, "AddUniqueManagedRename", "A managed rename requires a source, target, and object type."
    If StrComp(sourceName, targetName, vbBinaryCompare) = 0 Then Exit Sub
    For Each item In renames
        If StrComp(CStr(item(0)), sourceName, vbTextCompare) = 0 And StrComp(CStr(item(2)), objectTypeCode, vbTextCompare) = 0 Then
            If StrComp(CStr(item(1)), targetName, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 8966, "AddUniqueManagedRename", "One managed object source resolved to conflicting destinations." & vbCrLf & "Object type: " & objectTypeCode & vbCrLf & "Source: " & sourceName & vbCrLf & "First target: " & CStr(item(1)) & vbCrLf & "Second target: " & targetName
            Exit Sub
        End If
    Next item
    renames.Add Array(sourceName, targetName, objectTypeCode)
End Sub

Private Function RepairRenamedObjectReferences(ByVal db As DAO.Database, ByVal renames As Collection) As Long
    Dim managedTables As Object
    Dim tableName As Variant
    Dim tdf As DAO.tableDef
    Dim item As Variant
    Dim fieldName As Variant
    Dim exactFields As Variant
    Dim containsFields As Variant
    Dim sqlText As String
    exactFields = Array(FIELD_OBJECT_NAME, FIELD_OBJECT_DISPLAY_NAME, FIELD_RIBBON_TITLE, FIELD_FORM_OBJECT_NAME, FIELD_MANAGED_OBJECT_FIELD_SOURCE_OBJECT_NAME, FIELD_TARGET_OBJECT_NAME, FIELD_SOURCE_OBJECT_NAME_OVERRIDE, FIELD_TARGET_OBJECT_NAME_OVERRIDE, FIELD_DEPENDENCY_NAME)
    containsFields = Array(FIELD_PAYLOAD_TEXT)
    Set managedTables = ManagedPhysicalObjectNamesForRename(db, renames, "Table")
    For Each tableName In managedTables.keys
        PumpSetupResponsiveness
        If DaoTableExists(db, CStr(tableName)) Then
            Set tdf = db.TableDefs(CStr(tableName))
            For Each fieldName In exactFields
                PumpSetupResponsiveness
                If SharedDao_TableDefFieldExistsSafe(tdf, CStr(fieldName)) Then
                    For Each item In renames
                        PumpSetupResponsiveness
                        If ManagedObjectRenameFieldAppliesToItem(CStr(fieldName), item) Then
                            sqlText = "UPDATE [" & Replace(CStr(tableName), "]", "]]") & "] SET [" & CStr(fieldName) & "]=" & SharedDao_SqlQuote(CStr(item(1))) & " WHERE [" & CStr(fieldName) & "]=" & SharedDao_SqlQuote(CStr(item(0))) & ";"
                            db.Execute sqlText, dbFailOnError
                            RepairRenamedObjectReferences = RepairRenamedObjectReferences + db.RecordsAffected
                        End If
                    Next item
                End If
            Next fieldName
            For Each fieldName In containsFields
                PumpSetupResponsiveness
                If SharedDao_TableDefFieldExistsSafe(tdf, CStr(fieldName)) Then RepairRenamedObjectReferences = RepairRenamedObjectReferences + RepairTokenTextFieldReferences(db, CStr(tableName), CStr(fieldName), renames)
            Next fieldName
        End If
    Next tableName
End Function

Private Function RepairTokenTextFieldReferences(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String, ByVal renames As Collection) As Long
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim originalText As String
    Dim repairedText As String
    Set rs = db.OpenRecordset("SELECT * FROM [" & Replace(tableName, "]", "]]") & "] WHERE Len(Nz([" & fieldName & "],''))>0;", dbOpenDynaset)
    Do While Not rs.EOF
        PumpSetupResponsiveness
        originalText = CStr(rs.fields(fieldName).value)
        repairedText = ApplyManagedObjectRenameMap(originalText, renames)
        If StrComp(originalText, repairedText, vbBinaryCompare) <> 0 Then
            rs.Edit
            rs.fields(fieldName).value = repairedText
            rs.Update
            RepairTokenTextFieldReferences = RepairTokenTextFieldReferences + 1
        End If
        rs.MoveNext
    Loop
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
    Err.Raise savedErrNumber, "RepairTokenTextFieldReferences", "Could not repair stored object-name references." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fieldName & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Sub PreflightSingleQueryRename(ByVal db As DAO.Database, ByVal oldName As String, ByVal newName As String, ByVal databaseRole As String)
    If DaoTableExists(db, newName) Then Err.Raise vbObjectError + 8922, "PreflightSingleQueryRename", "The destination query name is already used by a table." & vbCrLf & "Database role: " & databaseRole & vbCrLf & "Destination: " & newName
    If SharedDao_QueryExists(db, oldName) And SharedDao_QueryExists(db, newName) Then Err.Raise vbObjectError + 8923, "PreflightSingleQueryRename", "Both the source and destination query names exist." & vbCrLf & "Database role: " & databaseRole & vbCrLf & "Source: " & oldName & vbCrLf & "Destination: " & newName
End Sub

Private Function RenameQueryInDatabase(ByVal db As DAO.Database, ByVal oldName As String, ByVal newName As String) As Long
    Dim qdf As DAO.queryDef
    If SharedDao_QueryExists(db, oldName) = False Then Exit Function
    If SharedDao_QueryExists(db, newName) Then Err.Raise vbObjectError + 8926, "RenameQueryInDatabase", "The destination query already exists: " & newName
    Set qdf = db.QueryDefs(oldName)
    qdf.Name = newName
    db.QueryDefs.Refresh
    If SharedDao_QueryExists(db, newName) = False Or SharedDao_QueryExists(db, oldName) Then Err.Raise vbObjectError + 8927, "RenameQueryInDatabase", "The query rename could not be verified." & vbCrLf & "Source: " & oldName & vbCrLf & "Destination: " & newName
    RenameQueryInDatabase = 1
End Function

Private Function RepairRenamedQuerySql(ByVal db As DAO.Database, ByVal renames As Collection) As Long
    Dim managedNames As Object
    Dim queryName As Variant
    Dim qdf As DAO.queryDef
    Dim originalSql As String
    Dim repairedSql As String
    Set managedNames = ManagedPhysicalObjectNamesForRename(db, renames, "Query")
    For Each queryName In managedNames.keys
        PumpSetupResponsiveness
        If SharedDao_QueryExists(db, CStr(queryName)) Then
            Set qdf = db.QueryDefs(CStr(queryName))
            originalSql = qdf.sql
            repairedSql = ApplyManagedObjectRenameMap(originalSql, renames)
            If StrComp(originalSql, repairedSql, vbBinaryCompare) <> 0 Then
                qdf.sql = repairedSql
                RepairRenamedQuerySql = RepairRenamedQuerySql + 1
            End If
        End If
    Next queryName
End Function

Private Function ConfiguredRenameTarget(ByVal renames As Collection, ByVal objectTypeCode As String, ByVal sourceName As String) As String
    Dim item As Variant
    ConfiguredRenameTarget = sourceName
    For Each item In renames
        PumpSetupResponsiveness
        If StrComp(CStr(item(2)), objectTypeCode, vbTextCompare) = 0 And StrComp(CStr(item(0)), sourceName, vbTextCompare) = 0 Then ConfiguredRenameTarget = CStr(item(1)): Exit Function
    Next item
End Function

Private Sub PreflightRenameDestinationSources(ByVal renames As Collection, ByVal databaseRole As String)
    Dim sourceCountByDestination As Object
    Dim sourceNamesByDestination As Object
    Dim item As Variant
    Dim objectKind As String
    Dim sourceName As String
    Dim destinationName As String
    Dim destinationKey As String
    ' INTENTIONAL PLAN/PROOF LIFECYCLE SEPARATION: sourceCountByDestination, sourceNamesByDestination holds planned/recovery work or independently observed validation evidence for PreflightRenameDestinationSources;
    ' keep it distinct from the source authority.
    Set sourceCountByDestination = CreateObject("Scripting.Dictionary")
    Set sourceNamesByDestination = CreateObject("Scripting.Dictionary")
    sourceCountByDestination.CompareMode = vbTextCompare
    sourceNamesByDestination.CompareMode = vbTextCompare
    For Each item In renames
        PumpSetupResponsiveness
        sourceName = CStr(item(0))
        destinationName = CStr(item(1))
        objectKind = CStr(item(2))
        If ConfiguredObjectExists(objectKind, sourceName) = False Then Err.Raise vbObjectError + 8934, "PreflightRenameDestinationSources", "A managed rename source object is missing." & vbCrLf & "Database role: " & databaseRole & vbCrLf & "Object type: " & objectKind & vbCrLf & "Source: " & sourceName
        destinationKey = UCase$(objectKind) & "|" & UCase$(destinationName)
        If sourceCountByDestination.Exists(destinationKey) Then
            sourceCountByDestination(destinationKey) = CLng(sourceCountByDestination(destinationKey)) + 1
            sourceNamesByDestination(destinationKey) = CStr(sourceNamesByDestination(destinationKey)) & ", " & sourceName
        Else
            sourceCountByDestination.Add destinationKey, 1
            sourceNamesByDestination.Add destinationKey, sourceName
        End If
        If CLng(sourceCountByDestination(destinationKey)) > 1 Then Err.Raise vbObjectError + 8935, "PreflightRenameDestinationSources", "Multiple managed-object source names target the same destination." & vbCrLf & "Database role: " & databaseRole & vbCrLf & "Object type: " & objectKind & vbCrLf & "Destination: " & destinationName & vbCrLf & "Sources: " & CStr(sourceNamesByDestination(destinationKey))
    Next item
End Sub

Private Function RenameConfiguredObject(ByVal db As DAO.Database, ByVal objectTypeCode As String, ByVal oldName As String, ByVal newName As String) As Long
    Select Case UCase$(Trim$(objectTypeCode))
        Case "TABLE": RenameConfiguredObject = RenameTableInDatabase(db, oldName, newName)
        Case "QUERY": RenameConfiguredObject = RenameQueryInDatabase(db, oldName, newName)
        Case "FORM"
            If ConfiguredObjectExists("Form", oldName) = False Then Exit Function
            DoCmd.Rename newName, acForm, oldName
            RenameConfiguredObject = 1
        Case "REPORT"
            If ConfiguredObjectExists("Report", oldName) = False Then Exit Function
            DoCmd.Rename newName, acReport, oldName
            RenameConfiguredObject = 1
        Case "MACRO"
            If ConfiguredObjectExists("Macro", oldName) = False Then Exit Function
            DoCmd.Rename newName, acMacro, oldName
            RenameConfiguredObject = 1
        Case Else
            Err.Raise vbObjectError + 8929, "RenameConfiguredObject", "Unsupported managed physical object type: " & objectTypeCode
    End Select
    If ConfiguredObjectExists(objectTypeCode, newName) = False Or ConfiguredObjectExists(objectTypeCode, oldName) Then Err.Raise vbObjectError + 8930, "RenameConfiguredObject", "The managed-object rename could not be verified." & vbCrLf & "Type: " & objectTypeCode & vbCrLf & "Source: " & oldName & vbCrLf & "Destination: " & newName
End Function

Private Function ReverseRenameMap(ByVal renames As Collection) As Collection
    ' INTENTIONAL ORDERED VIEW: reversedItems preserves the deterministic sequence required by ReverseRenameMap; keyed-only or unordered storage would not preserve the consumer semantics.
    Dim reversedItems As New Collection
    Dim index As Long
    Dim item As Variant
    If renames Is Nothing Then Set ReverseRenameMap = reversedItems: Exit Function
    For index = renames.count To 1 Step -1
        PumpSetupResponsiveness
        item = renames(index)
        reversedItems.Add Array(CStr(item(1)), CStr(item(0)), CStr(item(2)))
    Next index
    Set ReverseRenameMap = reversedItems
End Function

Private Sub PreflightSingleAccessObjectRename(ByVal objectTypeCode As String, ByVal oldName As String, ByVal newName As String, ByVal databaseRole As String)
    If ConfiguredObjectExists(objectTypeCode, oldName) And ConfiguredObjectExists(objectTypeCode, newName) Then Err.Raise vbObjectError + 8928, "PreflightSingleAccessObjectRename", "Both the source and destination object names exist." & vbCrLf & "Database role: " & databaseRole & vbCrLf & "Object type: " & objectTypeCode & vbCrLf & "Source: " & oldName & vbCrLf & "Destination: " & newName
End Sub

Private Function RepairRenamedLookupRowSources(ByVal db As DAO.Database, ByVal renames As Collection) As Long
    Dim managedTables As Object
    Dim tableName As Variant
    Dim tdf As DAO.tableDef
    Dim fld As DAO.Field
    Set managedTables = ManagedPhysicalObjectNamesForRename(db, renames, "Table")
    For Each tableName In managedTables.keys
        PumpSetupResponsiveness
        If DaoTableExists(db, CStr(tableName)) Then
            Set tdf = db.TableDefs(CStr(tableName))
            For Each fld In tdf.fields
                PumpSetupResponsiveness
                RepairRenamedLookupRowSources = RepairRenamedLookupRowSources + RepairRenamedFieldRowSource(fld, renames, CStr(tableName))
            Next fld
        End If
    Next tableName
End Function


Private Function ConfiguredObjectExists(ByVal objectTypeCode As String, ByVal objectName As String) As Boolean
    On Error GoTo NotFound
    Dim ao As accessObject
    Select Case UCase$(Trim$(objectTypeCode))
        Case "TABLE": ConfiguredObjectExists = DaoTableExists(ResolveCatalogDb(), objectName)
        Case "QUERY": ConfiguredObjectExists = SharedDao_QueryExists(ResolveCatalogDb(), objectName)
        Case "FORM": Set ao = CurrentProject.AllForms(objectName): ConfiguredObjectExists = True
        Case "REPORT": Set ao = CurrentProject.AllReports(objectName): ConfiguredObjectExists = True
        Case "MACRO": Set ao = CurrentProject.AllMacros(objectName): ConfiguredObjectExists = True
    End Select
    Exit Function
NotFound:
    ConfiguredObjectExists = False
End Function

' ===============================================================================================================================================================================
' STANDALONE / NAME-BASED DELETION ENGINE
' ===============================================================================================================================================================================
' This fallback is owned by Catalog but intentionally does not require the Managed Object Catalog table to be readable or complete.

Private Sub TraceStandaloneDeletion(ByVal levelText As String, ByVal sourceProcedure As String, ByVal messageText As String, Optional ByVal detailText As String = vbNullString, Optional ByVal objectName As String = vbNullString, Optional ByVal objectType As String = vbNullString)
    On Error Resume Next
    LogTraceMessage levelText, "DeleteObjects", sourceProcedure, messageText, detailText, "Standalone ADC Reset", 0, vbNullString, vbNullString, objectName, objectType, "CodeDb", "SystemMacro65Standalone"
    On Error GoTo 0
End Sub

' Purpose: Performs confirm DB objects by name deletion by coordinating `TraceStandaloneDeletion` and `StrComp` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   showCancellationMessage (Input Optional; Boolean): True displays cancellation message; False suppresses that user-facing presentation without changing the underlying work. Default: True.
' Returns: Boolean: True on the function's affirmative/success branch for confirm DB objects by name deletion; False on its explicit negative, unavailable, rejected, or handled-failure branch.
'   path.
Public Function ConfirmDBObjectsByNameDeletion(Optional ByVal showCancellationMessage As Boolean = True) As Boolean
    Dim typedText As String
    TraceStandaloneDeletion "Info", "ConfirmDBObjectsByNameDeletion", "Standalone ADC Reset confirmation displayed."
    typedText = InputBox( _
        "This deletes ADC tables, queries, generated forms, reports, System Macros and Help Document Macros, " & _
        "and generated configuration macros by captured object name." & vbCrLf & vbCrLf & _
        "THE TRACELOG NAMESPACE IS PRESERVED." & vbCrLf & _
        "CUSTOM/USER FORMS ARE PRESERVED." & vbCrLf & vbCrLf & _
        "Type DELETE ALL to continue.", _
        "Standalone ADC Reset")
    If StrComp(Trim$(typedText), "DELETE ALL", vbBinaryCompare) <> 0 Then
        TraceStandaloneDeletion "Info", "ConfirmDBObjectsByNameDeletion", "Standalone ADC Reset confirmation cancelled.", "Typed confirmation did not match DELETE ALL."
        If showCancellationMessage Then MsgBox "Standalone ADC reset cancelled.", vbInformation, "Standalone ADC Reset"
        Exit Function
    End If
    TraceStandaloneDeletion "Info", "ConfirmDBObjectsByNameDeletion", "Standalone ADC Reset confirmation accepted."
    ConfirmDBObjectsByNameDeletion = True
End Function

' Do not delete. Intentional public convenience entry point for manual standalone deletion.
' Purpose: Provides the intentional manual standalone reset entry point that delegates to DeleteDBObjectsByName with typed confirmation and completion messaging enabled.
' Public role: Public ADC API. The routine is intentionally callable outside its defining procedure/module and may be used directly by Access/VBA or late-bound callers when that API surface is needed.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub DeleteAllDBObjectsByName()
    Call DeleteDBObjectsByName(True, True)
End Sub

' Do not delete. Intentional public diagnostic entry point for manually previewing standalone name-based deletion targets.
' Purpose: Performs preview DB objects by name by coordinating `BuildDeletionNameSets` and `PrintNameSet` while preserving this module's existing validation and error semantics.
' Public role: Public ADC API. The routine is intentionally callable outside its defining procedure/module and may be used directly by Access/VBA or late-bound callers when that API surface is needed.
' Parameters: None.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub PreviewDBObjectsByName()
    Dim tables As Object
    Dim queries As Object
    Dim forms As Object
    Dim reports As Object
    Dim systemMacros As Object
    Dim helpMacros As Object
    Dim configMacros As Object
    BuildDeletionNameSets tables, queries, forms, reports, systemMacros, helpMacros, configMacros
    Debug.Print String$(80, "=")
    Debug.Print "ADC standalone deletion preview"
    Debug.Print "TraceLog namespace: preserved"
    Debug.Print "Tables:        " & CStr(tables.count)
    Debug.Print "Queries:       " & CStr(queries.count)
    Debug.Print "Forms:         " & CStr(forms.count) & "  (generated forms only; custom forms preserved)"
    Debug.Print "Reports:       " & CStr(reports.count)
    Debug.Print "System macros: " & CStr(systemMacros.count)
    Debug.Print "Help macros:   " & CStr(helpMacros.count)
    Debug.Print "Config macros: " & CStr(configMacros.count)
    Debug.Print String$(80, "-")
    PrintNameSet "TABLE", tables
    PrintNameSet "QUERY", queries
    PrintNameSet "FORM", forms
    PrintNameSet "REPORT", reports
    PrintNameSet "SYSTEM MACRO", systemMacros
    PrintNameSet "HELP MACRO", helpMacros
    PrintNameSet "CONFIG MACRO", configMacros
    Debug.Print String$(80, "=")
End Sub

Private Sub ClearStandaloneDeferredMacroDeleteState()
    mStandaloneDeferredMacroDeleteName = vbNullString
    mStandaloneDeferredMacroDeleteShowCompletion = False
    mStandaloneDeferredMacroDeleteDeletedCount = 0
    mStandaloneDeferredMacroDeleteMissingCount = 0
    mStandaloneDeferredMacroDeleteFailedCount = 0
    mStandaloneDeferredMacroDeletePreservedFormCount = 0
    mStandaloneDeferredMacroDeleteFailureText = vbNullString
End Sub

' Purpose: Performs standalone try delete deferred invoking macro65 by coordinating `AccessObjectExists` and `End` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   macroName (Input; String): Physical Access macro name identifying the macro object to inspect, run, delete, or verify.
'   failureText (Output; String): Receives the resolved failure text produced from this procedure's validated inputs and current ADC state.
'   wasMissing (Output; Boolean): Set True when the procedure determines that was missing applies; set False when that condition is not met.
' Returns: Boolean: True when the condition represented by DeleteDBObjectsByName is present under the routine's documented inputs/current state; False when that condition is absent.
Public Function StandaloneTryDeleteDeferredInvokingMacro65(ByVal macroName As String, ByRef failureText As String, ByRef wasMissing As Boolean) As Boolean
    'Standalone name-based deletion retains ownership of its final Macro 65 object; Core invokes this only after RunCode has returned.
    On Error GoTo ErrHandler
    Dim objectExists As Boolean
    failureText = vbNullString
    wasMissing = False
    macroName = Trim$(macroName)
    If Len(macroName) = 0 Then Err.Raise vbObjectError + 9656, "StandaloneTryDeleteDeferredInvokingMacro65", "The deferred invoking Macro 65 name is blank."
    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "StandaloneTryDeleteDeferredInvokingMacro65", "Deferred Macro 65 pre-delete existence inspection starting.", "Owner=Standalone", macroName, 0, vbNullString, vbNullString, macroName, "Macro", "CodeDb", "SystemMacro65"
    FlushTraceLogFile
    objectExists = AccessObjectExists(acMacro, macroName)
    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "StandaloneTryDeleteDeferredInvokingMacro65", "Deferred Macro 65 pre-delete existence inspection returned.", "Owner=Standalone; Exists=" & CStr(objectExists), macroName, 0, vbNullString, vbNullString, macroName, "Macro", "CodeDb", "SystemMacro65"
    FlushTraceLogFile
    If objectExists = False Then
        wasMissing = True
        StandaloneTryDeleteDeferredInvokingMacro65 = True
        Exit Function
    End If
    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "StandaloneTryDeleteDeferredInvokingMacro65", "PHYSICAL MACRO DELETE STARTING.", "Owner=Standalone", macroName, 0, vbNullString, vbNullString, macroName, "Macro", "CodeDb", "SystemMacro65"
    FlushTraceLogFile
    DoCmd.DeleteObject acMacro, macroName
    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "StandaloneTryDeleteDeferredInvokingMacro65", "PHYSICAL MACRO DELETE RETURNED.", "Owner=Standalone", macroName, 0, vbNullString, vbNullString, macroName, "Macro", "CodeDb", "SystemMacro65"
    FlushTraceLogFile
    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "StandaloneTryDeleteDeferredInvokingMacro65", "Deferred Macro 65 post-delete existence inspection starting.", "Owner=Standalone", macroName, 0, vbNullString, vbNullString, macroName, "Macro", "CodeDb", "SystemMacro65"
    FlushTraceLogFile
    objectExists = AccessObjectExists(acMacro, macroName)
    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "StandaloneTryDeleteDeferredInvokingMacro65", "Deferred Macro 65 post-delete existence inspection returned.", "Owner=Standalone; Exists=" & CStr(objectExists), macroName, 0, vbNullString, vbNullString, macroName, "Macro", "CodeDb", "SystemMacro65"
    FlushTraceLogFile
    If objectExists Then
        failureText = "Access still reports the invoking Macro 65 object present after the standalone deletion attempt."
        Exit Function
    End If
    StandaloneTryDeleteDeferredInvokingMacro65 = True
    Exit Function
ErrHandler:
    failureText = "Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
End Function

' Purpose: Performs standalone deferred macro deletion completed by coordinating `AppendLine` and `End` while preserving this module's existing validation and error semantics.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   macroDeleted (Input; Boolean): Boolean state describing whether macro deleted applies to this invocation; True selects the named condition/path and False selects its complementary/default path.
'   macroMissing (Input; Boolean): Boolean state describing whether macro missing applies to this invocation; True selects the named condition/path and False selects its complementary/default path.
'   finalFailureText (Input; String): Input final failure text for standalone deferred macro deletion completed; is tested by If Len(Trim$(finalFailureText)) > 0; is passed to `AppendLine`.
' Returns: Boolean: True when the standalone deferred macro deletion completed condition is satisfied for the supplied or current ADC context; False otherwise.
Public Function StandaloneDeferredMacroDeletionCompleted(ByVal macroDeleted As Boolean, ByVal macroMissing As Boolean, ByVal finalFailureText As String) As Boolean
    'Receives the one post-RunCode Macro 65 result from Core and completes standalone-owned summary/presentation.
    On Error Resume Next
    Dim completionSucceeded As Boolean
    Dim messageText As String
    If macroDeleted Then mStandaloneDeferredMacroDeleteDeletedCount = mStandaloneDeferredMacroDeleteDeletedCount + 1
    If macroMissing Then mStandaloneDeferredMacroDeleteMissingCount = mStandaloneDeferredMacroDeleteMissingCount + 1
    If Len(Trim$(finalFailureText)) > 0 Then
        mStandaloneDeferredMacroDeleteFailedCount = mStandaloneDeferredMacroDeleteFailedCount + 1
        mStandaloneDeferredMacroDeleteFailureText = AppendLine(mStandaloneDeferredMacroDeleteFailureText, "Failed final invoking Macro 65 [" & mStandaloneDeferredMacroDeleteName & "]: " & finalFailureText)
    End If
    completionSucceeded = (mStandaloneDeferredMacroDeleteFailedCount = 0)
    TraceStandaloneDeletion IIf(completionSucceeded, "Info", "Error"), "StandaloneDeferredMacroDeletionCompleted", _
        IIf(completionSucceeded, "Standalone ADC Reset succeeded.", "Standalone ADC Reset completed with failures."), _
        "Deleted=" & CStr(mStandaloneDeferredMacroDeleteDeletedCount) & "; Missing=" & CStr(mStandaloneDeferredMacroDeleteMissingCount) & "; Failed=" & CStr(mStandaloneDeferredMacroDeleteFailedCount) & "; NonTargetFormsPreserved=" & CStr(mStandaloneDeferredMacroDeletePreservedFormCount), mStandaloneDeferredMacroDeleteName, "Macro"
    FlushTraceLogFile
    Debug.Print "Standalone ADC reset complete. Deleted=" & CStr(mStandaloneDeferredMacroDeleteDeletedCount) & _
                "; missing/skipped=" & CStr(mStandaloneDeferredMacroDeleteMissingCount) & _
                "; failed=" & CStr(mStandaloneDeferredMacroDeleteFailedCount) & _
                "; non-target forms preserved=" & CStr(mStandaloneDeferredMacroDeletePreservedFormCount)
    If Len(mStandaloneDeferredMacroDeleteFailureText) > 0 Then Debug.Print mStandaloneDeferredMacroDeleteFailureText
    If mStandaloneDeferredMacroDeleteShowCompletion Then
        LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "StandaloneDeferredMacroDeletionCompleted", "Standalone deferred Macro 65 completion message starting.", "CompletionSucceeded=" & CStr(completionSucceeded), mStandaloneDeferredMacroDeleteName, 0, vbNullString, vbNullString, mStandaloneDeferredMacroDeleteName, "Macro", "CodeDb", "SystemMacro65"
        FlushTraceLogFile
        messageText = "Standalone ADC reset completed." & vbCrLf & vbCrLf & _
                      "Deleted objects/relationships: " & CStr(mStandaloneDeferredMacroDeleteDeletedCount) & vbCrLf & _
                      "Already missing: " & CStr(mStandaloneDeferredMacroDeleteMissingCount) & vbCrLf & _
                      "Failures: " & CStr(mStandaloneDeferredMacroDeleteFailedCount) & vbCrLf & _
                      "Non-target forms preserved: " & CStr(mStandaloneDeferredMacroDeletePreservedFormCount) & vbCrLf & _
                      "TraceLog namespace: retained" & vbCrLf & vbCrLf & _
                      "Close and reopen the database before running Setup."
        If completionSucceeded Then
            MsgBox messageText, vbInformation, "Standalone ADC Reset"
        Else
            MsgBox messageText & vbCrLf & vbCrLf & "See the Immediate Window and trace log for failure details.", vbCritical, "Standalone ADC Reset"
        End If
        LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "StandaloneDeferredMacroDeletionCompleted", "Standalone deferred Macro 65 completion message returned.", "CompletionSucceeded=" & CStr(completionSucceeded), mStandaloneDeferredMacroDeleteName, 0, vbNullString, vbNullString, mStandaloneDeferredMacroDeleteName, "Macro", "CodeDb", "SystemMacro65"
        FlushTraceLogFile
    End If
    StandaloneDeferredMacroDeletionCompleted = completionSucceeded
    LogTraceMessage TRACE_LEVEL_INFO, "DeleteObjects", "StandaloneDeferredMacroDeletionCompleted", "Standalone deferred Macro 65 state cleanup starting.", "CompletionSucceeded=" & CStr(completionSucceeded), mStandaloneDeferredMacroDeleteName, 0, vbNullString, vbNullString, mStandaloneDeferredMacroDeleteName, "Macro", "CodeDb", "SystemMacro65"
    FlushTraceLogFile
    ClearStandaloneDeferredMacroDeleteState
    On Error GoTo 0
End Function

' Purpose: Deletes DB objects by name for require typed confirmation, show completion message.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   requireTypedConfirmation (Input Optional; Boolean): True requires typed confirmation before the operation may proceed; False allows the operation to continue without making that condition mandatory.
'   Default: True.
'   showCompletionMessage (Input Optional; Boolean): True displays completion message; False suppresses that user-facing presentation without changing the underlying work. Default: True.
'   completeOwnedTraceAfterFinalMacro (Input Optional; Boolean): Boolean state describing whether complete owned trace after final macro applies to this invocation; True selects the named condition/path and
'   False selects its complementary/default path. Default: False.
'   traceSuccessMessage (Input Optional; String): Trace success message text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine. Default:
'     vbNullString.
'   traceFailureMessage (Input Optional; String): Trace failure message text included in this procedure's user-facing or diagnostic output; blank handling follows the explicit branches in the routine. Default:
'     vbNullString.
' Returns: Boolean result of delete DB objects by name; True and False correspond to the affirmative and negative branches explicitly implemented by this function.
Public Function DeleteDBObjectsByName( _
    Optional ByVal requireTypedConfirmation As Boolean = True, _
    Optional ByVal showCompletionMessage As Boolean = True, _
    Optional ByVal completeOwnedTraceAfterFinalMacro As Boolean = False, _
    Optional ByVal traceSuccessMessage As String = vbNullString, _
    Optional ByVal traceFailureMessage As String = vbNullString) As Boolean
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim tables As Object
    Dim queries As Object
    Dim forms As Object
    Dim reports As Object
    Dim systemMacros As Object
    Dim helpMacros As Object
    Dim configMacros As Object
    Dim deletedCount As Long
    Dim missingCount As Long
    Dim failedCount As Long
    Dim preservedFormCount As Long
    Dim failureText As String
    Dim messageText As String
    Dim invokingMacroName As String
    Dim deferInvokingMacro As Boolean
    If CoreDeferredMacro65DeletionRequired() Or CoreDeferredMacro65DeletionPending() Then Err.Raise vbObjectError + 9655, "DeleteDBObjectsByName", "A prior terminal Macro 65 deletion is still pending. No new standalone destructive deletion was started."
    ClearStandaloneDeferredMacroDeleteState
    If requireTypedConfirmation Then
        If ConfirmDBObjectsByNameDeletion(showCompletionMessage) = False Then Exit Function
    End If
    TraceStandaloneDeletion "Info", "DeleteDBObjectsByName", "Standalone deletion-set construction started."
    Set db = currentDb
    BuildDeletionNameSets tables, queries, forms, reports, systemMacros, helpMacros, configMacros
    invokingMacroName = CoreInvokingMacro65ObjectName()
    If Len(invokingMacroName) = 0 Then Err.Raise vbObjectError + 9651, "DeleteDBObjectsByName", "The canonical System Macro 65 object name could not be resolved before standalone destructive deletion. No object was deleted."
    deferInvokingMacro = CoreShouldDeferInvokingMacro65Deletion(invokingMacroName)
    If deferInvokingMacro Then
        'Register the execution-context deferral while canonical System Macro identity/naming authority is still available.
        'The terminal reset may subsequently delete the catalog/naming tables that would be required to re-resolve this identity.
        CoreMarkInvokingMacro65DeletionDeferred invokingMacroName, "Standalone"
    End If
    preservedFormCount = CountFormsNotInDeletionSet(forms)
    TraceStandalonePersistentGeneratedFormPreservation
    TraceStandaloneDeletion "Info", "DeleteDBObjectsByName", "Standalone deletion-set construction completed.", _
        "Tables=" & CStr(tables.count) & "; Queries=" & CStr(queries.count) & "; GeneratedForms=" & CStr(forms.count) & "; Reports=" & CStr(reports.count) & "; SystemMacros=" & CStr(systemMacros.count) & "; HelpMacros=" & CStr(helpMacros.count) & "; GeneratedConfigMacros=" & CStr(configMacros.count) & "; DeferredInvokingMacro=" & IIf(deferInvokingMacro, invokingMacroName, "(none)")
    TraceStandaloneDeletion "Info", "DeleteDBObjectsByName", "TraceLog namespace retained for standalone Delete All; database-backed and physical trace-file logging remain available."
    TraceStandaloneDeletion "Info", "DeleteDBObjectsByName", "Terminal UI teardown started."
    x__UTIL_ADC_UI_RIBBON.BeginTerminalTeardown
    TraceStandaloneDeletion "Info", "DeleteDBObjectsByName", "Terminal UI teardown completed; ADC Ribbon disabled for terminal deletion."
    TraceStandaloneDeletion "Info", "DeleteDBObjectsByName", "Generated forms deletion started.", "Planned=" & CStr(forms.count)
    DeleteAccessObjectSet acForm, forms, deletedCount, missingCount, failedCount, failureText
    TraceStandaloneDeletion "Info", "DeleteDBObjectsByName", "Generated forms deletion completed."
    TraceStandaloneDeletion "Info", "DeleteDBObjectsByName", "Reports deletion started.", "Planned=" & CStr(reports.count)
    DeleteAccessObjectSet acReport, reports, deletedCount, missingCount, failedCount, failureText
    TraceStandaloneDeletion "Info", "DeleteDBObjectsByName", "Reports deletion completed."
    TraceStandaloneDeletion "Info", "DeleteDBObjectsByName", "System macros deletion started.", "Planned=" & CStr(systemMacros.count) & "; Invoking Macro 65 deferred=" & CStr(deferInvokingMacro)
    DeleteAccessObjectSet acMacro, systemMacros, deletedCount, missingCount, failedCount, failureText, IIf(deferInvokingMacro, invokingMacroName, vbNullString)
    TraceStandaloneDeletion "Info", "DeleteDBObjectsByName", IIf(deferInvokingMacro, "System macros deletion completed except deferred invoking Macro 65.", "System macros deletion completed.")
    TraceStandaloneDeletion "Info", "DeleteDBObjectsByName", "Help macros deletion started.", "Planned=" & CStr(helpMacros.count)
    DeleteAccessObjectSet acMacro, helpMacros, deletedCount, missingCount, failedCount, failureText, IIf(deferInvokingMacro, invokingMacroName, vbNullString)
    TraceStandaloneDeletion "Info", "DeleteDBObjectsByName", "Help macros deletion completed."
    TraceStandaloneDeletion "Info", "DeleteDBObjectsByName", "Generated Config macros deletion started.", "Planned=" & CStr(configMacros.count)
    DeleteAccessObjectSet acMacro, configMacros, deletedCount, missingCount, failedCount, failureText, IIf(deferInvokingMacro, invokingMacroName, vbNullString)
    TraceStandaloneDeletion "Info", "DeleteDBObjectsByName", "Generated Config macros deletion completed."
    TraceStandaloneDeletion "Info", "DeleteDBObjectsByName", "Queries deletion started.", "Planned=" & CStr(queries.count)
    DeleteQuerySet db, queries, deletedCount, missingCount, failedCount, failureText
    TraceStandaloneDeletion "Info", "DeleteDBObjectsByName", "Queries deletion completed."
    TraceStandaloneDeletion "Info", "DeleteDBObjectsByName", "Final terminal cleanup started.", "Relationships and ADC Ribbon persistence precede table deletion."
    DeleteRelationsForTargetTables db, tables, deletedCount, failedCount, failureText
    ResetRibbonPersistence db, deletedCount, failedCount, failureText
    TraceStandaloneDeletion "Info", "DeleteDBObjectsByName", "Tables deletion started.", "Planned=" & CStr(tables.count)
    DeleteTableSet db, tables, deletedCount, missingCount, failedCount, failureText
    TraceStandaloneDeletion "Info", "DeleteDBObjectsByName", "Tables deletion completed."
    If deferInvokingMacro Then
        TraceStandaloneDeletion "Info", "DeleteDBObjectsByName", "All ordinary standalone deletion completed; invoking Macro 65 remains deferred until its RunCode call returns.", _
            "Deleted=" & CStr(deletedCount) & "; Missing=" & CStr(missingCount) & "; Failed=" & CStr(failedCount), invokingMacroName, "Macro"
        mStandaloneDeferredMacroDeleteName = invokingMacroName
        mStandaloneDeferredMacroDeleteShowCompletion = showCompletionMessage
        mStandaloneDeferredMacroDeleteDeletedCount = deletedCount
        mStandaloneDeferredMacroDeleteMissingCount = missingCount
        mStandaloneDeferredMacroDeleteFailedCount = failedCount
        mStandaloneDeferredMacroDeletePreservedFormCount = preservedFormCount
        mStandaloneDeferredMacroDeleteFailureText = failureText
        CoreScheduleDeferredMacro65Deletion completeOwnedTraceAfterFinalMacro, traceSuccessMessage, traceFailureMessage
        DeleteDBObjectsByName = (failedCount = 0)
        TraceStandaloneDeletion IIf(failedCount = 0, "Info", "Warning"), "DeleteDBObjectsByName", "Standalone ADC Reset ordinary deletion completed; final Macro 65 deletion is pending outside the active RunCode stack.", _
            "DeletedBeforeFinalMacro=" & CStr(deletedCount) & "; Missing=" & CStr(missingCount) & "; FailedBeforeFinalMacro=" & CStr(failedCount) & "; NonTargetFormsPreserved=" & CStr(preservedFormCount), invokingMacroName, "Macro"
        FlushTraceLogFile
    Else
        DeleteDBObjectsByName = (failedCount = 0)
        TraceStandaloneDeletion IIf(failedCount = 0, "Info", "Error"), "DeleteDBObjectsByName", IIf(failedCount = 0, "Standalone ADC Reset succeeded.", "Standalone ADC Reset completed with failures."), _
            "Deleted=" & CStr(deletedCount) & "; Missing=" & CStr(missingCount) & "; Failed=" & CStr(failedCount) & "; NonTargetFormsPreserved=" & CStr(preservedFormCount)
        Debug.Print "Standalone ADC reset complete. Deleted=" & CStr(deletedCount) & "; missing/skipped=" & CStr(missingCount) & "; failed=" & CStr(failedCount) & "; non-target forms preserved=" & CStr(preservedFormCount)
        If Len(failureText) > 0 Then Debug.Print failureText
        If showCompletionMessage Then
            messageText = "Standalone ADC reset completed." & vbCrLf & vbCrLf & _
                          "Deleted objects/relationships: " & CStr(deletedCount) & vbCrLf & _
                          "Already missing: " & CStr(missingCount) & vbCrLf & _
                          "Failures: " & CStr(failedCount) & vbCrLf & _
                          "Non-target forms preserved: " & CStr(preservedFormCount) & vbCrLf & _
                          "TraceLog namespace: retained" & vbCrLf & vbCrLf & _
                          "Close and reopen the database before running Setup."
            If failedCount = 0 Then
                MsgBox messageText, vbInformation, "Standalone ADC Reset"
            Else
                MsgBox messageText & vbCrLf & vbCrLf & "See the Immediate Window and trace log for failure details.", vbCritical, "Standalone ADC Reset"
            End If
        End If
    End If
    Exit Function
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    'If terminal standalone work aborts before Core has scheduled the post-RunCode delete, cancel the registered deferral.
    'A failed destructive run must not later delete Macro 65 merely because its execution stack unwound.
    On Error Resume Next
    If CoreDeferredMacro65DeletionRequired() And CoreDeferredMacro65DeletionPending() = False Then
        If StrComp(CoreDeferredMacro65DeletionCompletionKind(), "Standalone", vbTextCompare) = 0 Then CoreCancelDeferredMacro65DeletionRequest
    End If
    On Error GoTo 0
    failureText = AppendLine(failureText, "Fatal error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription)
    TraceStandaloneDeletion "Error", "DeleteDBObjectsByName", "Standalone ADC Reset failed.", failureText
    Debug.Print failureText
    If showCompletionMessage Then
        MsgBox "Standalone ADC reset did not complete." & vbCrLf & vbCrLf & _
               "Err.Number: " & CStr(savedErrNumber) & vbCrLf & _
               "Err.Source: " & savedErrSource & vbCrLf & _
               "Err.Description: " & savedErrDescription & vbCrLf & vbCrLf & _
               "See the Immediate Window and trace log for details.", vbCritical, "Standalone ADC Reset"
    End If
End Function

Private Sub BuildDeletionNameSets( _
    ByRef tables As Object, _
    ByRef queries As Object, _
    ByRef forms As Object, _
    ByRef reports As Object, _
    ByRef systemMacros As Object, _
    ByRef helpMacros As Object, _
    ByRef configMacros As Object)
    Set tables = NewTextSet()
    Set queries = NewTextSet()
    Set forms = NewTextSet()
    Set reports = NewTextSet()
    Set systemMacros = NewTextSet()
    Set helpMacros = NewTextSet()
    Set configMacros = NewTextSet()
    AddStandardManagedDefinitions CoreSupportManagedObjectDefinitions(), tables, queries, reports
    AddStandardManagedDefinitions ReferenceManagedObjectDefinitions(), tables, queries, reports
    AddStandardManagedDefinitions ManagedObjectCatalogTableDefinitions(), tables, queries, reports
    AddStandardManagedDefinitions ModuleGuardManagedObjectDefinitions(), tables, queries, reports
    AddStandardManagedDefinitions ManagedObjectReviewQueryDefinitions(), tables, queries, reports
    AddStandardManagedDefinitions EditableLookupTableManagedObjectDefinitions(), tables, queries, reports
    AddSupersededLookupQueryNamesIfPresent queries
    AddStandardManagedDefinitions NavigationRibbonManagementManagedObjectDefinitions(), tables, queries, reports
    AddStandardManagedDefinitions FormDesignManagedObjectDefinitions(), tables, queries, reports
    AddStandardManagedDefinitions ReportManagedObjectDefinitions(), tables, queries, reports
    AddAddInReviewDefinitions AddInReviewFullPlacementDefinitions(), tables, queries, reports
    AddGeneratedFormNames ManagedGeneratedFormNames(), forms
    AddCanonicalSystemMacroNames systemMacros
    CaptureExistingSystemMacroNames systemMacros
    AddCanonicalHelpMacroNames helpMacros
    CaptureExistingHelpMacroNames helpMacros
    CaptureGeneratedConfigMacroNames configMacros
End Sub

Private Sub AddSupersededLookupQueryNamesIfPresent(ByVal queries As Object)
    'Covers the exact query-to-table transition from the supplied baseline when Delete All is run before Setup has converted the old QueryDefs.
    Dim item As Variant
    Dim objectName As String
    For Each item In EditableLookupTableManagedObjectDefinitions()
        objectName = Trim$(CStr(item(0)))
        If AccessObjectExists(acQuery, objectName) Then AddName queries, objectName
    Next item
    objectName = "MacroCatalog_Records_Lookup_RecordOwnerDBRole"
    If AccessObjectExists(acQuery, objectName) Then AddName queries, objectName
End Sub

Private Sub AddStandardManagedDefinitions( _
    ByVal definitions As Collection, _
    ByVal tables As Object, _
    ByVal queries As Object, _
    ByVal reports As Object)
    Dim item As Variant
    Dim objectName As String
    Dim objectType As String
    For Each item In definitions
        objectName = Trim$(CStr(item(0)))
        objectType = UCase$(Trim$(CStr(item(2))))
        Select Case objectType
            Case "TABLE": AddName tables, objectName
            Case "QUERY": AddName queries, objectName
            Case "REPORT": AddName reports, objectName
        End Select
    Next item
End Sub

Private Sub AddAddInReviewDefinitions( _
    ByVal definitions As Collection, _
    ByVal tables As Object, _
    ByVal queries As Object, _
    ByVal reports As Object)
    Dim item As Variant
    Dim objectName As String
    Dim objectType As String
    For Each item In definitions
        objectName = Trim$(CStr(item(0)))
        objectType = UCase$(Trim$(CStr(item(1))))
        Select Case objectType
            Case "TABLE": AddName tables, objectName
            Case "QUERY": AddName queries, objectName
            Case "REPORT": AddName reports, objectName
        End Select
    Next item
End Sub

Private Sub AddGeneratedFormNames(ByVal names As Collection, ByVal forms As Object)
    Dim item As Variant
    For Each item In names
        AddName forms, CStr(item)
    Next item
End Sub

Private Sub AddCanonicalSystemMacroNames(ByVal macros As Object)
    Dim definitions As Collection
    Dim Definition As Variant
    Dim macroName As String
    Set definitions = GetSystemMacroDefinitions()
    For Each Definition In definitions
        macroName = SystemMacroCalculatedTitle( _
                        CStr(Definition(1)), _
                        CStr(Definition(2)), _
                        CStr(Definition(0)), _
                        Definition)
        AddName macros, macroName
    Next Definition
End Sub

Private Sub AddCanonicalHelpMacroNames(ByVal macros As Object)
    Dim definitions As Collection
    Dim Definition As Variant
    Dim macroNumber As Long
    Dim macroName As String
    Set definitions = HelpDocumentDefinitions()
    For Each Definition In definitions
        macroNumber = GetHelpDocumentMacroNumber(CLng(Definition(1)))
        macroName = SystemMacroCalculatedTitle(CStr(macroNumber), CStr(Definition(2)))
        AddName macros, macroName
    Next Definition
End Sub

Private Sub CaptureExistingSystemMacroNames(ByVal macros As Object)
    On Error GoTo CleanExit
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim systemMacroKey As String
    Dim TitleText As String
    Dim macroTitleOverride As String
    Dim macroName As String
    Dim systemMacroDesignator As String
    Dim systemMacrosTableName As String
    Set db = currentDb
    systemMacrosTableName = ManagedObjectPhysicalName("Table", "ADC_SystemMacros")
    If TableExistsInDatabase(db, systemMacrosTableName) = False Then Exit Sub
    sqlText = "SELECT " & SQL_FIELD_SYSTEM_MACRO_KEY & ", " & SQL_FIELD_SYSTEM_MACRO_DESIGNATOR & ", " & SQL_FIELD_DEFAULT_TITLE_TEXT & ", " & SQL_FIELD_MACRO_TITLE_OVERRIDE & " FROM [" & systemMacrosTableName & "];"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    Do While rs.EOF = False
        systemMacroKey = Trim$(SharedText_Nz(rs.fields(FIELD_SYSTEM_MACRO_KEY).value))
        systemMacroDesignator = CStr(Nz(rs.fields(FIELD_SYSTEM_MACRO_DESIGNATOR).value, vbNullString))
        macroTitleOverride = Trim$(SharedText_Nz(rs.fields(FIELD_MACRO_TITLE_OVERRIDE).value))
        TitleText = Trim$(SharedText_Nz(rs.fields(FIELD_DEFAULT_TITLE_TEXT).value))
        If Len(macroTitleOverride) > 0 Then
            AddName macros, macroTitleOverride
        ElseIf Len(systemMacroDesignator) > 0 And Len(TitleText) > 0 Then
            On Error Resume Next
            macroName = SystemMacroCalculatedTitle(systemMacroDesignator, TitleText, systemMacroKey)
            On Error GoTo CleanExit
            AddName macros, macroName
        End If
        If Len(TitleText) > 0 Then AddLegacySetupPartMacroName macros, TitleText
        rs.MoveNext
    Loop
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
End Sub

Private Sub AddLegacySetupPartMacroName(ByVal macros As Object, ByVal TitleText As String)
    'Allows this reset to remove Setup Part macros from the immediately preceding phase roster even when their keys are no longer present in the current code-owned definition set.
    Dim dashPos As Long
    Dim identityText As String
    Dim descriptiveText As String
    Dim logicalTitle As String
    Dim macroName As String
    If Len(TitleText) < 9 Then Exit Sub
    If StrComp(Left$(TitleText, 6), "Setup ", vbTextCompare) <> 0 Then Exit Sub
    dashPos = InStr(7, TitleText, " - ", vbBinaryCompare)
    If dashPos = 0 Then Exit Sub
    identityText = Trim$(Mid$(TitleText, 7, dashPos - 7))
    descriptiveText = Trim$(Mid$(TitleText, dashPos + 3))
    If Len(identityText) = 0 Or Len(descriptiveText) = 0 Then Exit Sub
    logicalTitle = identityText & " " & descriptiveText
    On Error Resume Next
    macroName = TargetManagedObjectPhysicalName("SystemMacro", logicalTitle)
    On Error GoTo 0
    AddName macros, macroName
End Sub

Private Sub CaptureExistingHelpMacroNames(ByVal macros As Object)
    On Error GoTo CleanExit
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim helpTableName As String
    Dim TitleText As String
    Dim macroTitleOverride As String
    Dim macroName As String
    Dim helpNumber As Long
    Set db = currentDb
    helpTableName = ManagedObjectPhysicalName("Table", "ADC_HelpDocuments")
    If TableExistsInDatabase(db, helpTableName) = False Then Exit Sub
    Set rs = db.OpenRecordset("SELECT " & SQL_FIELD_HELP_DOCUMENT_NUMBER & ", " & SQL_FIELD_DEFAULT_TITLE_TEXT & ", " & SQL_FIELD_MACRO_TITLE_OVERRIDE & " FROM [" & helpTableName & "] WHERE Nz(" & SQL_FIELD_HELP_DOCUMENT_STATUS & ",'Active')='Active';", dbOpenSnapshot)
    Do While rs.EOF = False
        helpNumber = CLng(Nz(rs.fields(FIELD_HELP_DOCUMENT_NUMBER).value, 0))
        macroTitleOverride = Trim$(SharedText_Nz(rs.fields(FIELD_MACRO_TITLE_OVERRIDE).value))
        TitleText = Trim$(SharedText_Nz(rs.fields(FIELD_DEFAULT_TITLE_TEXT).value))
        If Len(macroTitleOverride) > 0 Then
            AddName macros, macroTitleOverride
        ElseIf helpNumber > 0 And Len(TitleText) > 0 Then
            macroName = SystemMacroCalculatedTitle(GetHelpDocumentMacroNumber(helpNumber), TitleText)
            AddName macros, macroName
        End If
        rs.MoveNext
    Loop
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
End Sub

Private Sub CaptureGeneratedConfigMacroNames(ByVal macros As Object)
    On Error GoTo CleanExit
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlText As String
    Dim macroName As String
    Dim configTableName As String
    Set db = currentDb
    configTableName = ManagedObjectPhysicalName("Table", "ADC_Config")
    If TableExistsInDatabase(db, configTableName) = False Then Exit Sub
    sqlText = "SELECT DISTINCT " & SQL_FIELD_GENERATED_MACRO_OBJECT_NAME & " FROM [" & configTableName & "] WHERE Len(Nz(" & SQL_FIELD_GENERATED_MACRO_OBJECT_NAME & ",''))>0;"
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    Do While rs.EOF = False
        macroName = Trim$(CStr(rs.fields(0).value))
        AddName macros, macroName
        rs.MoveNext
    Loop
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
End Sub

Private Sub DeleteAccessObjectSet( _
    ByVal objectType As AcObjectType, _
    ByVal names As Object, _
    ByRef deletedCount As Long, _
    ByRef missingCount As Long, _
    ByRef failedCount As Long, _
    ByRef failureText As String, _
    Optional ByVal deferredObjectName As String = vbNullString)
    Dim key As Variant
    Dim objectName As String
    Dim errorText As String
    For Each key In names.keys
        objectName = CStr(names(key))
        If Len(deferredObjectName) > 0 And StrComp(objectName, deferredObjectName, vbTextCompare) = 0 Then GoTo NextObject
        errorText = vbNullString
        TraceStandaloneDeletion "Info", "DeleteAccessObjectSet", "Standalone object deletion attempt.", vbNullString, objectName, AccessObjectTypeLabel(objectType)
        If DeleteOneAccessObject(objectType, objectName, errorText) Then
            deletedCount = deletedCount + 1
            Debug.Print "Deleted " & AccessObjectTypeLabel(objectType) & ": " & objectName
        ElseIf Len(errorText) = 0 Then
            missingCount = missingCount + 1
        Else
            failedCount = failedCount + 1
            failureText = AppendLine(failureText, "Failed " & AccessObjectTypeLabel(objectType) & " [" & objectName & "]: " & errorText)
        End If
NextObject:
        DoEvents
    Next key
End Sub

Private Function DeleteOneAccessObject( _
    ByVal objectType As AcObjectType, _
    ByVal objectName As String, _
    ByRef errorText As String) As Boolean
    On Error GoTo ErrHandler
    If AccessObjectExists(objectType, objectName) = False Then Exit Function
    On Error Resume Next
    Select Case objectType
        Case acForm: DoCmd.Close acForm, objectName, acSaveNo
        Case acReport: DoCmd.Close acReport, objectName, acSaveNo
        Case acQuery: DoCmd.Close acQuery, objectName, acSaveNo
    End Select
    Err.Clear
    On Error GoTo ErrHandler
    DoCmd.DeleteObject objectType, objectName
    DeleteOneAccessObject = True
    Exit Function
ErrHandler:
    errorText = "Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
End Function

Private Sub DeleteQuerySet( _
    ByVal db As DAO.Database, _
    ByVal names As Object, _
    ByRef deletedCount As Long, _
    ByRef missingCount As Long, _
    ByRef failedCount As Long, _
    ByRef failureText As String)
    Dim key As Variant
    Dim objectName As String
    For Each key In names.keys
        objectName = CStr(names(key))
        TraceStandaloneDeletion "Info", "DeleteQuerySet", "Standalone object deletion attempt.", vbNullString, objectName, "Query"
        If QueryExistsInDatabase(db, objectName) = False Then
            missingCount = missingCount + 1
        Else
            On Error Resume Next
            DoCmd.Close acQuery, objectName, acSaveNo
            Err.Clear
            db.QueryDefs.Delete objectName
            If Err.Number = 0 Then
                deletedCount = deletedCount + 1
                Debug.Print "Deleted QUERY: " & objectName
            Else
                failedCount = failedCount + 1
                failureText = AppendLine(failureText, "Failed QUERY [" & objectName & "]: Error " & CStr(Err.Number) & " - " & Err.Description)
            End If
            On Error GoTo 0
        End If
        DoEvents
    Next key
End Sub

Private Sub DeleteTableSet( _
    ByVal db As DAO.Database, _
    ByVal names As Object, _
    ByRef deletedCount As Long, _
    ByRef missingCount As Long, _
    ByRef failedCount As Long, _
    ByRef failureText As String)
    Dim key As Variant
    Dim objectName As String
    For Each key In names.keys
        objectName = CStr(names(key))
        TraceStandaloneDeletion "Info", "DeleteTableSet", "Standalone object deletion attempt.", vbNullString, objectName, "Table"
        If TableExistsInDatabase(db, objectName) = False Then
            missingCount = missingCount + 1
        Else
            On Error Resume Next
            db.TableDefs.Delete objectName
            If Err.Number = 0 Then
                deletedCount = deletedCount + 1
                Debug.Print "Deleted TABLE: " & objectName
            Else
                failedCount = failedCount + 1
                failureText = AppendLine(failureText, "Failed TABLE [" & objectName & "]: Error " & CStr(Err.Number) & " - " & Err.Description)
            End If
            On Error GoTo 0
        End If
        DoEvents
    Next key
End Sub

Private Sub DeleteRelationsForTargetTables( _
    ByVal db As DAO.Database, _
    ByVal targetTables As Object, _
    ByRef deletedCount As Long, _
    ByRef failedCount As Long, _
    ByRef failureText As String)
    Dim i As Long
    Dim relationName As String
    Dim primaryTable As String
    Dim foreignTable As String
    For i = db.Relations.count - 1 To 0 Step -1
        relationName = db.Relations(i).Name
        primaryTable = db.Relations(i).Table
        foreignTable = db.Relations(i).foreignTable
        If SetContains(targetTables, primaryTable) Or SetContains(targetTables, foreignTable) Then
            On Error Resume Next
            db.Relations.Delete relationName
            If Err.Number = 0 Then
                deletedCount = deletedCount + 1
                Debug.Print "Deleted RELATION: " & relationName
            Else
                failedCount = failedCount + 1
                failureText = AppendLine(failureText, "Failed RELATION [" & relationName & "]: Error " & CStr(Err.Number) & " - " & Err.Description)
            End If
            On Error GoTo 0
        End If
    Next i
End Sub

Private Sub ResetRibbonPersistence( _
    ByVal db As DAO.Database, _
    ByRef deletedCount As Long, _
    ByRef failedCount As Long, _
    ByRef failureText As String)
    On Error Resume Next
    Dim currentRibbon As String
    Dim rs As DAO.Recordset
    currentRibbon = CStr(Nz(db.properties("CustomRibbonID").value, vbNullString))
    If Err.Number = 0 Then
        If StrComp(Trim$(currentRibbon), RIBBON_NAME, vbTextCompare) = 0 Then
            Err.Clear
            db.properties.Delete "CustomRibbonID"
            If Err.Number = 0 Then
                deletedCount = deletedCount + 1
                Debug.Print "Removed database CustomRibbonID: " & RIBBON_NAME
            Else
                failedCount = failedCount + 1
                failureText = AppendLine(failureText, "Failed to remove CustomRibbonID: Error " & CStr(Err.Number) & " - " & Err.Description)
            End If
        End If
    End If
    Err.Clear
    'Discard ADC-only Ribbon fast-path/reopen evidence so the next Setup cannot mistake a deleted physical Ribbon state for a current one.
    RemoveDatabasePropertyIfExists db, "ADCRibbonRestartRequired", deletedCount, failedCount, failureText
    RemoveDatabasePropertyIfExists db, "ADCRibbonRestartRequiredOn", deletedCount, failedCount, failureText
    RemoveDatabasePropertyIfExists db, "ADCRibbonLoadedOn", deletedCount, failedCount, failureText
    RemoveDatabasePropertyIfExists db, "ADCRibbonDependencyToken", deletedCount, failedCount, failureText
    RemoveDatabasePropertyIfExists db, "ADCRibbonStandardInputFingerprint", deletedCount, failedCount, failureText
    RemoveDatabasePropertyIfExists db, "ADCRibbonStandardXmlHash", deletedCount, failedCount, failureText
    RemoveDatabasePropertyIfExists db, "ADCRibbonStandardTargetMapHash", deletedCount, failedCount, failureText
    RemoveDatabasePropertyIfExists db, "ADCRibbonStandardButtonCount", deletedCount, failedCount, failureText
    If TableExistsInDatabase(db, DB_TABLE_USYS_RIBBONS) Then
        db.Execute "DELETE FROM " & SharedDao_BracketName(DB_TABLE_USYS_RIBBONS) & " WHERE " & SQL_FIELD_RIBBON_NAME & "='ADC';", dbFailOnError
        If Err.Number = 0 Then
            If db.RecordsAffected > 0 Then
                deletedCount = deletedCount + db.RecordsAffected
                Debug.Print "Deleted ADC row(s) from USysRibbons: " & CStr(db.RecordsAffected)
            End If
        Else
            failedCount = failedCount + 1
            failureText = AppendLine(failureText, "Failed to delete ADC row from USysRibbons: Error " & CStr(Err.Number) & " - " & Err.Description)
            Err.Clear
        End If
        'Delete the table itself only when no unrelated Ribbon rows remain.
        Set rs = db.OpenRecordset("SELECT Count(*) " & "AS " & SQL_FIELD_CNT & " FROM " & SharedDao_BracketName(DB_TABLE_USYS_RIBBONS) & ";", dbOpenSnapshot)
        If Err.Number = 0 Then
            If CLng(rs.fields(FIELD_CNT).value) = 0 Then
                rs.Close
                Set rs = Nothing
                db.TableDefs.Delete DB_TABLE_USYS_RIBBONS
                If Err.Number = 0 Then
                    deletedCount = deletedCount + 1
                    Debug.Print "Deleted TABLE: USysRibbons"
                Else
                    failedCount = failedCount + 1
                    failureText = AppendLine(failureText, "Failed TABLE [USysRibbons]: Error " & CStr(Err.Number) & " - " & Err.Description)
                End If
            End If
        End If
        If Not rs Is Nothing Then rs.Close
    End If
    On Error GoTo 0
End Sub

Private Sub RemoveDatabasePropertyIfExists( _
    ByVal db As DAO.Database, _
    ByVal propertyName As String, _
    ByRef deletedCount As Long, _
    ByRef failedCount As Long, _
    ByRef failureText As String)
    On Error Resume Next
    Dim propertyValue As Variant
    propertyValue = db.properties(propertyName).value
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    db.properties.Delete propertyName
    If Err.Number = 0 Then
        deletedCount = deletedCount + 1
        Debug.Print "Removed database property: " & propertyName
    Else
        failedCount = failedCount + 1
        failureText = AppendLine(failureText, "Failed database property [" & propertyName & "]: Error " & CStr(Err.Number) & " - " & Err.Description)
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Function CountFormsNotInDeletionSet(ByVal deletionForms As Object) As Long
    Dim obj As accessObject
    For Each obj In CurrentProject.AllForms
        If SetContains(deletionForms, obj.Name) = False Then CountFormsNotInDeletionSet = CountFormsNotInDeletionSet + 1
    Next obj
End Function

Private Sub TraceStandalonePersistentGeneratedFormPreservation()
    Dim db As DAO.Database
    Dim formNames As Collection
    Dim item As Variant
    Dim formName As String
    Set db = codeDb
    Set formNames = ManagedPersistentGeneratedFormNames()
    For Each item In formNames
        formName = CStr(item)
        If DatabaseAccessObjectExistsByName(db, acForm, formName) Then
            TraceStandaloneDeletion "Info", "DeleteDBObjectsByName", "Persistent Generated form preserved by standalone deletion policy.", _
                "Name=" & formName & "; Classification=PersistentGenerated; DeleteTarget=False; PhysicalOwnerDatabaseRole=CodeDb", formName, "Form"
        End If
    Next item
End Sub

Private Function AccessObjectExists(ByVal objectType As AcObjectType, ByVal objectName As String) As Boolean
    On Error GoTo Missing
    Dim formObj As accessObject
    Dim reportObj As accessObject
    Dim macroObj As accessObject
    Dim queryObj As accessObject
    Select Case objectType
        Case acForm
            Set formObj = CurrentProject.AllForms(objectName)
        Case acReport
            Set reportObj = CurrentProject.AllReports(objectName)
        Case acMacro
            Set macroObj = CurrentProject.AllMacros(objectName)
        Case acQuery
            Set queryObj = CurrentData.AllQueries(objectName)
        Case Else
            Exit Function
    End Select
    AccessObjectExists = True
Missing:
End Function

Private Function TableExistsInDatabase(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
    On Error GoTo Missing
    Dim tdf As DAO.tableDef
    Set tdf = db.TableDefs(tableName)
    TableExistsInDatabase = True
Missing:
End Function

Private Function QueryExistsInDatabase(ByVal db As DAO.Database, ByVal queryName As String) As Boolean
    On Error GoTo Missing
    Dim qdf As DAO.queryDef
    Set qdf = db.QueryDefs(queryName)
    QueryExistsInDatabase = True
Missing:
End Function

Private Function NewTextSet() As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    Set NewTextSet = result
End Function

Private Sub AddName(ByVal targetSet As Object, ByVal objectName As String)
    'Standalone Delete All uses this single name-set insertion boundary, so namespace preservation applies uniformly to every deletion set.
    'The individual table, query, generated-form, report, and macro owners remain unchanged.
    objectName = Trim$(objectName)
    If Len(objectName) = 0 Then Exit Sub
    If ObjectNameIsInRetainedTraceLogNamespace(objectName) Then Exit Sub
    If targetSet.Exists(objectName) = False Then targetSet.Add objectName, objectName
End Sub

Private Function SetContains(ByVal targetSet As Object, ByVal objectName As String) As Boolean
    If targetSet Is Nothing Then Exit Function
    SetContains = targetSet.Exists(Trim$(objectName))
End Function


Private Function AccessObjectTypeLabel(ByVal objectType As AcObjectType) As String
    Select Case objectType
        Case acForm: AccessObjectTypeLabel = "FORM"
        Case acReport: AccessObjectTypeLabel = "REPORT"
        Case acMacro: AccessObjectTypeLabel = "MACRO"
        Case acQuery: AccessObjectTypeLabel = "QUERY"
        Case acTable: AccessObjectTypeLabel = "TABLE"
        Case Else: AccessObjectTypeLabel = CStr(objectType)
    End Select
End Function

Private Sub PrintNameSet(ByVal typeLabel As String, ByVal names As Object)
    Dim key As Variant
    For Each key In names.keys
        Debug.Print typeLabel & ": " & CStr(names(key))
    Next key
End Sub
