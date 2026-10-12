Attribute VB_Name = "x__UTIL_ADC_DB_INVENTORY_VBA"
' ================================================================================================================================================================================
' ChatGPT Rewrite Tracking Header
' ================================================================================================================================================================================
' ChatGPT Previous Version ID: CGPT-KE8L-CCU2-7YQR-CXJH-JAX6
' ChatGPT Baseline Version ID: CGPT-KE8L-CCU2-7YQR-CXJH-JAX6
' ChatGPT Version ID: CGPT-T2KT-YHYD-LSVS-LQXT-KSP6
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
' STANDARD MODULE: Access Development Console DB Inventory VBA
' VB_Name: x__UTIL_ADC_DB_INVENTORY_VBA
' Entry Points:
'     RepairReferenceInventoryFoundationForSetup
' SHARED CODE: SharedDao_SqlQuote(ByVal value As Variant) -> String
' CLASSES USED: None
' DYNAMIC STRING REFERENCES:
'     Dynamic-reference scan stopped after 2000 logical line(s) to prevent a long-running scan; review this module manually if needed.
' LastScanTimestamp: Saturday 8/15/26 8:46 AM
' Purpose:
'     Owns the complete References/Add-Ins domain: A4 structural foundation, active/available reference scans, loaded VBA-project and Add-In inventories, traditional Access add-in registration/entry-point inventory,
'     Reference Manager runtime data, staged reference changes, apply/discard workflows, and the project/module/procedure inventory used by this subsystem.
' Author: Christina Shonts
' Support: Active-reference scanning, registered type-library scanning, reference views, descriptions, lookup metadata, and system-macro command support.
' Code Category: Access Development Console DB Inventory VBA/reference inventory and diagnostics.
' Requirements: DAO is required,  Active-reference scanning uses the Access Application.References collection, Available-reference scanning uses the Windows StdRegProv registry provider through late binding,
'     Loaded-project scanning uses Application.VBE.VBProjects and records project metadata without inspecting project source code. Traditional Access add-in membership is discovered independently from the installed
'     Access Menu Add-Ins registry configuration; registered add-in databases may be opened read-only to inspect USysRegInfo through the single canonical registration parser,  VBE/COM add-in scanning reads Application.VBE.AddIns
'     and Application.COMAddIns through late-bound collection access and never assigns Connect or changes registration. Traditional Access registration expressions are read as data only and are never executed.
' Reference Dependencies: DAO, the Access object model, the VBA extensibility object model already used for loaded-project inventory, and the Office COMAddIns collection.
' DB Object Dependencies:
'     DBInventory_VBARefs_Active
'     DBInventory_VBARefs_Avail
'     DBInventory_VBARefs_Active_Resolved
'     DBInventory_VBARefs_Avail_Current
'     DBInventory_VBAPrjs_Loaded
'     DBInventory_VBARefs_Active_ReferencedPrjLinks
'     DBInventory_VBAPrjs_Loaded_Current
'     DBInventory_VBARefs_Active_ReferencedPrjLinks_Current
'     DBInventory_AddIns_VBE
'     DBInventory_AddIns_Access
'     DBInventory_AddIns_COM
'     DBInventory_AddIns+VBAPrjs
'     DBInventory_VBARefs_@LoadedPrjs
'     DBInventory_VBAModules_@LoadedPrjs
'     DBInventory_VBAProcs_@LoadedPrjs
'     DBInventory_VBADeps_@LoadedPrjs
'     DBInventory_VBARefs_@LoadedPrjs_@ExtPrjs
'     DBInventory_VBAModules_@LoadedPrjs_@ExtPrjs
'     DBInventory_VBAProcs_@LoadedPrjs_@ExtPrjs_Pub
'     DBInventory_VBAProcs_@LoadedPrjs_@ExtPrjs_Pub_Conflicts
'     DBInventory_VBADeps_@LoadedPrjs_@ExtPrjs
' Usage:
'     Run RepairReferenceInventoryFoundationForSetup to create or repair the reference/Add-In inventory structural tables and saved queries. Generated reference forms are owned by the centralized form phase.
'     Run ScanDBInventoryActiveReferences to refresh active references and, for standalone use, rebuild the active-reference project crosswalk.
'     Run ScanDBInventoryAvailableReferences to refresh only registered references that are not active in the current database.
'     Run ScanDBInventoryLoadedVBAProjects to refresh loaded VBE project metadata, rebuild the active-reference project crosswalk, and complete registered Access add-in enrichment for standalone use.
'     Run ScanDBInventoryAddInsAndLoadedVBAProjects to refresh loaded-project, installed traditional Access registration, VBE add-in, and COM add-in retained inventories without changing add-in state or executing
'     registration expressions; unseen identities become Stale.
' Version:
'     Shifts current managed-description proof references from F1 to G1 after insertion of Setup dependency Phase F.
'     Removes the invalid DAO.QueryDef.Attributes reference from the A4 query-family fingerprint; SQL and QueryDef.Type remain the A4-owned definition proof inputs.
'     Consolidates Reference/Add-In parsing and path helpers into DB Inventory VBA and adds major navigation sections without changing inventory behavior or Setup-phase ownership.
'     Removes 14 unused local declarations from ScanDBInventoryAvailableReferences after the per-registration work was delegated to
'     ScanAvailableReferenceVersionRegistration; scan behavior and Setup-phase ownership are unchanged.
'     Routes Managed Object Catalog references to x__UTIL_ADC_MOC and physical lifecycle references to x__UTIL_ADC_MOC_LIFECYCLE after the responsibility split; runtime and Setup-phase behavior are unchanged.
'     Makes the loaded-module RI pre-pass and child-writing pass consume one captured VBComponent roster, and limits A4 structural proof to A4-owned indexes so B5 relationship indexes cannot stale A4.
'     Removes A4 explicit FK indexes superseded by B5 relationship-backed indexes, invalidates affected A4 proofs, and materializes missing module parents before source child writes without changing
'     cache reuse semantics.
'     Narrows the module-internal reference failure helper while retaining reference-manager callback compatibility.
'     Repaired Access/ACE JOIN 3296 risks by joining outer queries to canonical current partitions, moving entry-point normalization out of INNER JOIN predicates, and surfacing the exact saved query in A4 failure-stage diagnostics.
'     Anchored DBInventory_AddIns+VBAPrjs UNION output types to canonical source fields without replacing not-applicable Null values, and extended A4 executable validation to DBInventory_AddIns+VBAPrjs alongside
'     DBInventory_AddIns+VBAPrjs+VBAProcs_ReviewSummary.
'     Updated reference/Add-In description-proof lifecycle wording for F1 description ownership under the reordered Setup architecture.
'     Separated authoritative scan refresh from dependent crosswalk/correlation finalization so combined scans defer only work that later phases supersede while standalone scans remain complete.
'     Managed-header dependency metadata updated because x__UTIL_ADC_DEFAULTS is now a standard module; executable behavior unchanged.
'     Traditional Access add-in membership now comes from the installed Access Menu Add-Ins registry source; USysRegInfo remains the single add-in-database registration parser and loaded VBE projects only enrich
'     registered rows with runtime project/procedure evidence.
'     Routed the staged Reference Manager Yes/No/Cancel prompt through Core popup policy while preserving the two-button apply confirmation.
'     Replaced reproduced Reference family identifiers with authoritative DEFAULTS-owned family-code accessors across inventory, repair, and proof-token paths.
'     Consolidated all live References/Add-Ins scans, Reference Manager workflows, staged changes, and related project/module/procedure inventory into this module.
'     Renamed from x__UTIL_ADC_REFERENCES to x__UTIL_ADC_DB_INVENTORY_VBA to make its existing DB Inventory VBA/reference/Add-In responsibility explicit; executable behavior is unchanged.
'     Updated current usage documentation to the surviving canonical reference-foundation and runtime inventory authority.
'     Reference-form setup preflight hardening. The complete generated reference-form change plan is now inspected before any reference table/query mutation, affected open forms use the shared save/close gate, and
'     inspection errors fail closed instead of silently closing forms or treating them as stale.
' Notes:
'     Active References means references selected for the current database VBA project.
'     Available References means registered type libraries found on the computer that are not active for the current database.
'     Loaded VBA Projects means the top-level projects currently exposed by Application.VBE.VBProjects, whether or not the current project references them.
'     VBE Add-ins means every add-in represented in the VBE Add-In Manager collection; COM Add-ins means every Office COM add-in available to Access through Application.COMAddIns. DBInventory_AddIns_Access contains
'     the installed traditional Access Menu Add-Ins registrations discovered from Access registry configuration; loaded external VBA projects may enrich those rows but never determine membership.
'     Scan tables are system-owned retained inventories. A successful scan marks prior current rows Stale and upserts rediscovered identities back to Current inside the same transaction; scan refresh never deletes them.
' Core Rules Enforced Here:
'     Active and available references remain distinct in storage, views, commands, and diagnostics.
'     A broken active reference remains visible in DBInventory_VBARefs_Active even when it cannot be matched to registration data.
'     Available-reference scanning does not add, remove, reorder, or repair active references.
'     Loaded-project scanning records metadata and accessibility only and never scans project source automatically; Source analysis is owned by ScanDBInventoryLoadedVBASource.
'     Add-in scanning is passive: it does not assign Connect, call COMAddIns.Update, change registration, load or unload add-ins, probe disconnected add-in Object properties, or scan external source.
'     Exact normalized file path is authoritative project-reference evidence; unique name-only resolution is retained but marked weaker.
'     A failed scan rolls back and preserves the prior current partition and retained history.
' Design Map:
'     DBInventory_VBARefs_Active retains observed reference identities; InventoryStatus identifies the current database reference order and health partition.
'     DBInventory_VBARefs_Avail stores registered non-active type-library versions and paths.
'     DBInventory_VBARefs_Active_Resolved and DBInventory_VBARefs_Avail_Current are separate read-only review views.
'     DBInventory_VBAPrjs_Loaded retains observed VBE project identities; InventoryStatus identifies the currently loaded partition.
'     DBInventory_VBARefs_Active_ReferencedPrjLinks records exact-path, name-only, unresolved, ambiguous, or not-applicable matching evidence.
'     DBInventory_AddIns_VBE and DBInventory_AddIns_COM retain separate passive registration/connection histories; DBInventory_AddIns+VBAPrjs combines them with DBInventory_VBAPrjs_Loaded for read-only review.
' Configuration: Registry roots include current-user and local-machine native and WOW6432Node TypeLib registrations.
' Behavior: Setup repairs structure only. Scanning occurs only through the dedicated scan commands.
' Code Status: Active canonical DB Inventory VBA/reference authority.
' Todo: None.
' Reporting Bugs: None currently recorded in this header.
' Module Level Bugs: None currently recorded in this header.
' Public Subs:
'     VerifyReferenceTableLookupDefinition(ByVal tdf As DAO.tableDef, ByVal Definition As Variant) -> None
'     RepairReferenceInventoryFoundationForSetup(Optional ByVal showMessage As Boolean = False) -> None
'     ScanDBInventoryActiveReferences(Optional ByVal showMessage As Boolean = True, Optional ByVal prepareReferenceInventory As Boolean = True, Optional ByVal rebuildProjectLinks As Boolean = True) -> None
'     ScanDBInventoryLoadedVBAProjects(Optional ByVal showMessage As Boolean = True, Optional ByVal prepareReferenceInventory As Boolean = True, Optional ByVal correlateRegisteredAccessAddIns As Boolean = True)
'     -> None
'     ScanDBInventoryAvailableReferences(Optional ByVal showMessage As Boolean = True, Optional ByVal prepareReferenceInventory As Boolean = True) -> None
'     ScanDBInventoryAddInsAndLoadedVBAProjects(Optional ByVal showMessage As Boolean = True, Optional ByVal prepareReferenceInventory As Boolean = True, Optional ByVal refreshLoadedVBAProjects As Boolean =
'     True, Optional ByVal correlateEntryPoints As Boolean = True) -> None
'     ScanDBInventoryLoadedVBASource(Optional ByVal showMessage As Boolean = True, Optional ByVal prepareReferenceInventory As Boolean = True, Optional ByVal refreshLoadedVBAProjects As Boolean = True, Optional
'     ByVal correlateAccessAddInsWithProjects As Boolean = True) -> None
'     OpenReferenceFileLocation(ByVal referencePath As String, Optional ByVal showMessage As Boolean = True) -> None
'     RefreshReferenceManager(Optional ByVal showMessage As Boolean = True, Optional ByVal prepareReferenceInventory As Boolean = True) -> None
'     ApplyReferenceManagerChanges(Optional ByVal showMessage As Boolean = True) -> None
'     DiscardReferenceManagerChanges(Optional ByVal showMessage As Boolean = True) -> None
' Public Functions:
'     ReferenceLocaleDecimalText(ByVal localeText As String) -> String
'     NormalizeReferenceGuid(ByVal guidText As String) -> String
'     ParseReferenceVersion(ByVal versionText As String, ByRef majorVersion As Long, ByRef minorVersion As Long) -> Boolean
'     ExpandReferencePath(ByVal pathText As String) -> String
'     NormalizeReferencePath(ByVal pathText As String) -> String
'     ReferenceFileBaseName(ByVal pathText As String) -> String
'     RegisteredAccessAddInProcedureName(ByVal expressionText As String) -> String
'     ReferencesConfiguredPhysicalObjectNamesForSetup() -> Collection
'     ReferenceInventoryFoundationCurrentForSetup(ByRef detailText As String) -> Boolean
'     ReferenceTableFieldOrderText(ByVal tableName As String) -> String
'     ReferenceManagerFormOpenInitialize() -> Boolean
' Public Properties: None
' ================================================================================================================================================================================

Option Compare Database
Option Explicit

' ===============================================================================================================================================================================
' MODULE-LEVEL DECLARATIONS
' ===============================================================================================================================================================================

' Reference structure-token property names and contract versions.
Private Const REFERENCE_INVENTORY_TOKEN_PROPERTY As String = "ADCReferenceInventoryStructureTokenV1"
Private Const REFERENCE_LOADED_VBA_ANALYSIS_TOKEN_PROPERTY As String = "ADCReferenceLoadedVBAAnalysisStructureTokenV1"
Private Const REFERENCE_CROSSWALK_TOKEN_PROPERTY As String = "ADCReferenceCrosswalkStructureTokenV1"
Private Const REFERENCE_QUERY_TOKEN_PROPERTY As String = "ADCReferenceQueryStructureTokenV1"
Private Const REFERENCE_INVENTORY_TOKEN_VERSION As String = "2026-08-28 A4-owned reference inventory physical schema without later-phase descriptions/lookup presentation v9"
Private Const REFERENCE_LOADED_VBA_ANALYSIS_TOKEN_VERSION As String = "2026-08-28 A4-owned loaded-VBA-analysis physical schema without later-phase descriptions/lookup presentation v9"
Private Const LOADED_VBA_SOURCE_SCAN_VERSION As Long = 2
Private Const REFERENCE_CROSSWALK_TOKEN_VERSION As String = "2026-08-28 A4-owned reference crosswalk physical schema without later-phase descriptions/lookup presentation v7"
Private Const REFERENCE_QUERY_TOKEN_VERSION As String = "2026-08-28 A4-owned reference query definition/executable-contract without later-phase Description DateUpdate v13"
Private Const INVENTORY_STATUS_CURRENT As String = "Current"
Private Const INVENTORY_STATUS_STALE As String = "Stale"

' Reference registry roots and TypeLib paths.
Private Const HKEY_CURRENT_USER As Long = &H80000001
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
Private Const REGISTRY_TYPELIB_NATIVE As String = "SOFTWARE\Classes\TypeLib"
Private Const REGISTRY_TYPELIB_WOW6432 As String = "SOFTWARE\Classes\WOW6432Node\TypeLib"

' ===============================================================================================================================================================================
' PROCEDURES
' ===============================================================================================================================================================================

' =============================================================================================================================================================================
' REFERENCE / ADD-IN FOUNDATION, SCHEMA, AND PREFLIGHT PROCEDURES
' =============================================================================================================================================================================

' Purpose: Resolves reference locale decimal text from the supplied/current ADC state, using `ElseIf` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   localeText (Input; String): Input locale text for reference locale decimal text; is tested by If Len(localeText) = 0; is tested by If UCase$(Left$(localeText, 2)) = "0X".
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function ReferenceLocaleDecimalText(ByVal localeText As String) As String
    'Converts decimal or hexadecimal reference-locale text to its decimal string representation, returning blank when parsing fails.
    On Error GoTo NotParsed
    If Len(localeText) = 0 Then Exit Function
    If UCase$(Left$(localeText, 2)) = "0X" Then
        ReferenceLocaleDecimalText = CStr(CLng("&H" & Mid$(localeText, 3)))
    ElseIf localeText Like "*[A-Fa-f]*" Then
        ReferenceLocaleDecimalText = CStr(CLng("&H" & localeText))
    Else
        ReferenceLocaleDecimalText = CStr(CLng(localeText))
    End If
    Exit Function
NotParsed:
    ReferenceLocaleDecimalText = vbNullString
End Function

' Purpose: Resolves normalize reference guid from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   guidText (Input; String): Input guid text for normalize reference guid; is tested by If Len(guidText) = 0; is tested by If Left$(guidText, 1) <> "{".
' Returns: String: Returns the normalize reference guid text value. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function NormalizeReferenceGuid(ByVal guidText As String) As String
    'Trims and uppercases a reference GUID and ensures enclosing braces are present.
    guidText = UCase$(Trim$(guidText))
    If Len(guidText) = 0 Then Exit Function
    If Left$(guidText, 1) <> "{" Then guidText = "{" & guidText
    If Right$(guidText, 1) <> "}" Then guidText = guidText & "}"
    NormalizeReferenceGuid = guidText
End Function

' Purpose: Parses a two-part dotted reference version into major and minor Long values.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   versionText (Input; String): Input version text for parse reference version; is passed to `Split`.
'   majorVersion (Output; Long): Receives the major version returned by `CLng` after this procedure completes its validation/resolution work.
'   minorVersion (Output; Long): Receives the minor version returned by `CLng` after this procedure completes its validation/resolution work.
' Returns: Boolean: True when the condition represented by LeadingIdentifier is present under the routine's documented inputs/current state; False when that condition is absent.
Public Function ParseReferenceVersion(ByVal versionText As String, ByRef majorVersion As Long, ByRef minorVersion As Long) As Boolean
    'Parses a two-part dotted reference version into major and minor Long values.
    Dim parts As Variant
    parts = Split(versionText, ".")
    If UBound(parts) <> 1 Then Exit Function
    If IsNumeric(parts(0)) = False Or IsNumeric(parts(1)) = False Then Exit Function
    majorVersion = CLng(parts(0))
    minorVersion = CLng(parts(1))
    ParseReferenceVersion = True
End Function

' Purpose: Resolves expand reference path from the supplied/current ADC state, using `CreateObject` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   pathText (Input; String): Path text identifying the file/location examined or transformed by this procedure; path normalization and existence checks are performed only where coded here.
' Returns: String: Returns the expand reference path text value. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function ExpandReferencePath(ByVal pathText As String) As String
    'Trims a reference path, expands environment variables, and removes matching outer double quotes.
    On Error GoTo ExitHere
    pathText = Trim$(pathText)
    If Len(pathText) = 0 Then Exit Function
    pathText = CreateObject("WScript.Shell").ExpandEnvironmentStrings(pathText)
    If Len(pathText) >= 2 And Left$(pathText, 1) = Chr$(34) And Right$(pathText, 1) = Chr$(34) Then pathText = Mid$(pathText, 2, Len(pathText) - 2)
    ExpandReferencePath = pathText
    Exit Function
ExitHere:
    ExpandReferencePath = pathText
End Function

' Purpose: Resolves normalize reference path from the supplied/current ADC state, using `ExpandReferencePath` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   pathText (Input; String): Path text identifying the file/location examined or transformed by this procedure; path normalization and existence checks are performed only where coded here.
' Returns: String: Returns the normalize reference path text value. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function NormalizeReferencePath(ByVal pathText As String) As String
    'Expands a reference path, normalizes path separators, trims it, and converts it to lowercase for comparison.
    pathText = ExpandReferencePath(pathText)
    pathText = Replace(pathText, "/", "\")
    NormalizeReferencePath = LCase$(Trim$(pathText))
End Function

' Purpose: Resolves reference file base name from the supplied/current ADC state, using `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   pathText (Input; String): Path text identifying the file/location examined or transformed by this procedure; path normalization and existence checks are performed only where coded here.
' Returns: String: Returns the resolved name produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function ReferenceFileBaseName(ByVal pathText As String) As String
    'Normalizes path separators and returns the final file-name component of a reference path.
    Dim slashPos As Long
    pathText = Replace(Trim$(pathText), "/", "\")
    slashPos = InStrRev(pathText, "\")
    If slashPos > 0 Then pathText = Mid$(pathText, slashPos + 1)
    ReferenceFileBaseName = pathText
End Function

' Purpose: Resolves registered access add in procedure name from the supplied/current ADC state, using `Next` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   expressionText (Input; String): Input expression text for registered access add in procedure name; is passed to `candidate`.
' Returns: String: Returns the resolved name produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not resolved.
Public Function RegisteredAccessAddInProcedureName(ByVal expressionText As String) As String
    'Extracts and validates a bare procedure name from the traditional Access add-in expression form =FunctionName().
    'Single narrow authority for direct traditional Access add-in expressions of the form =FunctionName().
    Dim candidate As String
    Dim ProcedureName As String
    Dim charIndex As Long
    Dim currentChar As String
    candidate = Trim$(expressionText)
    If Len(candidate) < 4 Or Left$(candidate, 1) <> "=" Then Exit Function
    candidate = Trim$(Mid$(candidate, 2))
    If Len(candidate) < 3 Or Right$(candidate, 2) <> "()" Then Exit Function
    ProcedureName = Trim$(Left$(candidate, Len(candidate) - 2))
    If Len(ProcedureName) = 0 Then Exit Function
    currentChar = Left$(ProcedureName, 1)
    If currentChar Like "[A-Za-z]" = False Then Exit Function
    For charIndex = 2 To Len(ProcedureName)
        currentChar = Mid$(ProcedureName, charIndex, 1)
        If currentChar Like "[A-Za-z0-9_]" = False Then Exit Function
    Next charIndex
    RegisteredAccessAddInProcedureName = ProcedureName
End Function

Private Function ReferenceTableFamilyStateToken(ByVal db As DAO.Database, ByVal versionText As String, ByVal tableNames As Collection, ByRef failureDetail As String) As String
    On Error GoTo ErrHandler
    Dim item As Variant
    Dim tdf As DAO.tableDef
    Dim tableName As String
    Dim payloadText As String
    Dim fld As DAO.Field
    Dim idx As DAO.index
    Dim indexField As DAO.Field
    Dim indexMetadata As Object
    Dim sortedIndexNames As Variant
    Dim indexName As String
    Dim indexPayload As String
    Dim indexPosition As Long
    Dim a4OwnedIndexCount As Long
    failureDetail = vbNullString
    payloadText = versionText & "|"
    For Each item In tableNames
        tableName = CStr(item)
        Set tdf = Nothing
        On Error Resume Next
        Set tdf = db.TableDefs(tableName)
        On Error GoTo ErrHandler
        If tdf Is Nothing Then
            failureDetail = "Missing table: " & tableName
            Exit Function
        End If
        payloadText = payloadText & FingerprintTokenPart(tdf.Name)
        payloadText = payloadText & FingerprintTokenPart(CStr(tdf.Attributes))
        payloadText = payloadText & FingerprintTokenPart(CStr(tdf.fields.count))
        For Each fld In tdf.fields
            payloadText = payloadText & FingerprintTokenPart(fld.Name)
            payloadText = payloadText & FingerprintTokenPart(CStr(fld.Type))
            payloadText = payloadText & FingerprintTokenPart(CStr(fld.Size))
            payloadText = payloadText & FingerprintTokenPart(CStr(fld.Attributes))
            payloadText = payloadText & FingerprintTokenPart(CStr(fld.Required))
            If fld.Type = dbText Or fld.Type = dbMemo Then
                payloadText = payloadText & FingerprintTokenPart(CStr(fld.allowZeroLength))
            Else
                payloadText = payloadText & FingerprintTokenPart("<N/A>")
            End If
            payloadText = payloadText & FingerprintTokenPart(CStr(Nz(fld.DefaultValue, vbNullString)))
            payloadText = payloadText & FingerprintTokenPart(CStr(Nz(fld.validationRule, vbNullString)))
            payloadText = payloadText & FingerprintTokenPart(CStr(Nz(fld.validationText, vbNullString)))
        Next fld
        'A4 structural proof owns physical schema/default/validation and its primary/unique index contracts. G1 Description and B2/D4 lookup-presentation metadata are intentionally excluded.
        'DAO non-unique Foreign indexes are B5 relationship support and are also excluded so later owner phases cannot invalidate A4 proof.
        ' INTENTIONAL KEYED LOOKUP: indexMetadata provides keyed membership/lookup over source rows or state already read by ReferenceTableFamilyStateToken; it is a derived access path, not a second authority.
        Set indexMetadata = CreateObject("Scripting.Dictionary")
        indexMetadata.CompareMode = vbTextCompare
        a4OwnedIndexCount = 0
        For Each idx In tdf.Indexes
            If idx.Foreign = False Or idx.Unique Or idx.Primary Then
                indexPayload = FingerprintTokenPart(CStr(idx.Unique)) & FingerprintTokenPart(CStr(idx.Primary))
                For Each indexField In idx.fields
                    indexPayload = indexPayload & FingerprintTokenPart(indexField.Name)
                Next indexField
                indexMetadata(idx.Name) = indexPayload
                a4OwnedIndexCount = a4OwnedIndexCount + 1
            End If
        Next idx
        payloadText = payloadText & FingerprintTokenPart(CStr(a4OwnedIndexCount))
        If indexMetadata.count > 0 Then
            sortedIndexNames = SortedDictionaryKeys(indexMetadata)
            For indexPosition = LBound(sortedIndexNames) To UBound(sortedIndexNames)
                indexName = CStr(sortedIndexNames(indexPosition))
                payloadText = payloadText & FingerprintTokenPart(indexName) & CStr(indexMetadata(indexName))
            Next indexPosition
        End If
    Next item
    ReferenceTableFamilyStateToken = FingerprintFromPayload(payloadText)
    Exit Function
ErrHandler:
    failureDetail = "Could not calculate reference table-family metadata. Table=" & tableName & "; Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
End Function

Private Function ReferenceFieldLookupStateText(ByVal fld As DAO.Field) As String
    Dim propertyNames As Variant
    Dim propertyName As Variant
    Dim prp As DAO.Property
    Dim partText As String
    propertyNames = Array("DisplayControl", "RowSourceType", "RowSource", "BoundColumn", "ColumnCount", "ColumnWidths", "LimitToList")
    For Each propertyName In propertyNames
        Set prp = Nothing
        On Error Resume Next
        Err.Clear
        Set prp = fld.properties(CStr(propertyName))
        If Err.Number = 0 And Not prp Is Nothing Then
            partText = partText & CStr(propertyName) & "=" & CStr(Nz(prp.value, vbNullString)) & "|"
        Else
            Err.Clear
            partText = partText & CStr(propertyName) & "=<absent>|"
        End If
        On Error GoTo 0
    Next propertyName
    ReferenceFieldLookupStateText = partText
End Function

Private Function ReferenceQueryFamilyStateToken(ByVal db As DAO.Database, ByVal versionText As String, ByVal queryNames As Collection, ByRef failureDetail As String) As String
    On Error GoTo ErrHandler
    Dim item As Variant
    Dim queryName As String
    Dim qdf As DAO.queryDef
    Dim validationRs As DAO.Recordset
    Dim payloadText As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim diagnosticText As String
    failureDetail = vbNullString
    payloadText = versionText & "|"
    For Each item In queryNames
        queryName = CStr(item)
        Set qdf = Nothing
        On Error Resume Next
        Set qdf = db.QueryDefs(queryName)
        On Error GoTo ErrHandler
        If qdf Is Nothing Then
            failureDetail = "Missing saved query: " & queryName
            Exit Function
        End If
        'A4 proof is definition-based. QueryDef SQL/type detect A4-owned query changes without coupling the proof to G1 Description-driven MSysObjects.DateUpdate changes.
        payloadText = payloadText & FingerprintTokenPart(qdf.Name)
        payloadText = payloadText & FingerprintTokenPart(CStr(qdf.Type))
        payloadText = payloadText & FingerprintTokenPart(CStr(qdf.sql))
        'ACE can persist these compound UNION QueryDefs while deferring incompatible output typing until execution, so A4 execute-validates the canonical Add-In UNION views before accepting the token.
        If StrComp(queryName, ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Inventory"), vbTextCompare) = 0 Or StrComp(queryName, ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Review_Summary"), vbTextCompare) = 0 Or StrComp(queryName, ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Reference_Report_Source"), vbTextCompare) = 0 Then
            Set validationRs = db.OpenRecordset("SELECT * FROM [" & queryName & "] WHERE 1=0;", dbOpenSnapshot)
            validationRs.Close
            Set validationRs = Nothing
        End If
    Next item
    ReferenceQueryFamilyStateToken = FingerprintFromPayload(payloadText)
    Exit Function
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    failureDetail = "Could not calculate or execute-validate reference query-family definition. Query=" & queryName & "; Error " & CStr(savedErrNumber) & " from " & savedErrSource & ": " & savedErrDescription
    If savedErrNumber = 3464 And StrComp(queryName, ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Review_Summary"), vbTextCompare) = 0 Then
        diagnosticText = ReferenceProjectIdentityJoinDiagnostics(db)
        If Len(diagnosticText) > 0 Then failureDetail = failureDetail & vbCrLf & diagnosticText
        Err.Clear
        diagnosticText = ReferenceAddInReviewUnionDiagnostics(db)
        If Len(diagnosticText) > 0 Then failureDetail = failureDetail & vbCrLf & diagnosticText
        Err.Clear
        diagnosticText = ReferenceAddInReviewValidationMethodDiagnostics(db, queryName)
        If Len(diagnosticText) > 0 Then failureDetail = failureDetail & vbCrLf & diagnosticText
        Err.Clear
    End If
    On Error Resume Next
    If Not validationRs Is Nothing Then validationRs.Close
    Set qdf = Nothing
End Function

Private Function ReferenceAddInReviewValidationMethodDiagnostics(ByVal db As DAO.Database, ByVal queryName As String) As String
    'Runs only after the saved DBInv_AddIns+VBAPrjs+VBAProcs_ReviewSummary validation wrapper fails with DAO 3464; successful A4 runs incur no probe work.
    On Error Resume Next
    Dim qdf As DAO.queryDef
    Dim probeRs As DAO.Recordset
    Dim escapedName As String
    Dim detailText As String
    Dim probeErrNumber As Long
    Dim probeErrSource As String
    Dim probeErrDescription As String
    escapedName = Replace(queryName, "]", "]]")
    detailText = queryName & " saved-QueryDef validation-method 3464 probes:"
    Err.Clear
    Set qdf = db.QueryDefs(queryName)
    If Err.Number <> 0 Or qdf Is Nothing Then
        ReferenceAddInReviewValidationMethodDiagnostics = detailText & vbCrLf & "Saved QueryDef lookup: FAIL Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
        Exit Function
    End If
    Set probeRs = Nothing
    Err.Clear
    Set probeRs = qdf.OpenRecordset(dbOpenSnapshot)
    probeErrNumber = Err.Number
    probeErrSource = Err.Source
    probeErrDescription = Err.Description
    If probeErrNumber = 0 And Not probeRs Is Nothing Then
        detailText = detailText & vbCrLf & "Direct QueryDef.OpenRecordset: PASS"
        probeRs.Close
    Else
        detailText = detailText & vbCrLf & "Direct QueryDef.OpenRecordset: FAIL Error " & CStr(probeErrNumber) & " from " & probeErrSource & ": " & probeErrDescription
    End If
    Set probeRs = Nothing
    Err.Clear
    Set probeRs = db.OpenRecordset("SELECT TOP 0 * FROM [" & escapedName & "];", dbOpenSnapshot)
    probeErrNumber = Err.Number
    probeErrSource = Err.Source
    probeErrDescription = Err.Description
    If probeErrNumber = 0 And Not probeRs Is Nothing Then
        detailText = detailText & vbCrLf & "SELECT TOP 0 wrapper: PASS"
        probeRs.Close
    Else
        detailText = detailText & vbCrLf & "SELECT TOP 0 wrapper: FAIL Error " & CStr(probeErrNumber) & " from " & probeErrSource & ": " & probeErrDescription
    End If
    Set probeRs = Nothing
    Err.Clear
    Set probeRs = db.OpenRecordset("SELECT * FROM [" & escapedName & "] WHERE 1=0;", dbOpenSnapshot)
    probeErrNumber = Err.Number
    probeErrSource = Err.Source
    probeErrDescription = Err.Description
    If probeErrNumber = 0 And Not probeRs Is Nothing Then
        detailText = detailText & vbCrLf & "SELECT * WHERE 1=0 wrapper: PASS"
        probeRs.Close
    Else
        detailText = detailText & vbCrLf & "SELECT * WHERE 1=0 wrapper: FAIL Error " & CStr(probeErrNumber) & " from " & probeErrSource & ": " & probeErrDescription
    End If
    Set probeRs = Nothing
    Set qdf = Nothing
    ReferenceAddInReviewValidationMethodDiagnostics = detailText
End Function

Private Function ReferenceProjectIdentityJoinDiagnostics(ByVal db As DAO.Database) As String
    'Runs only after DBInv_AddIns+VBAPrjs+VBAProcs_ReviewSummary executable validation fails with DAO 3464; successful A4 runs incur no probe work.
    On Error Resume Next
    Dim leftNames As Variant
    Dim rightNames As Variant
    Dim probeLabels As Variant
    Dim probeIndex As Long
    Dim leftName As String
    Dim rightName As String
    Dim leftMetadata As String
    Dim rightMetadata As String
    Dim probeResult As String
    Dim detailText As String
    Dim probeSql As String
    Dim leftRs As DAO.Recordset
    Dim rightRs As DAO.Recordset
    Dim joinRs As DAO.Recordset
    Dim probeErrNumber As Long
    Dim probeErrSource As String
    Dim probeErrDescription As String
    leftNames = Array(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AccessAddIns"), ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AccessAddIns"), ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AccessAddIns"), ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects"), ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects"))
    rightNames = Array(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects"), ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Project_Procedure_Counts"), ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Project_Conflict_Counts"), ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Project_Procedure_Counts"), ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Project_Conflict_Counts"))
    probeLabels = Array(CStr(leftNames(0)) & " -> " & CStr(rightNames(0)), CStr(leftNames(1)) & " -> " & CStr(rightNames(1)), CStr(leftNames(2)) & " -> " & CStr(rightNames(2)), CStr(leftNames(3)) & " -> " & CStr(rightNames(3)), CStr(leftNames(4)) & " -> " & CStr(rightNames(4)))
    detailText = "ProjectIdentity 3464 isolation probes:"
    For probeIndex = LBound(leftNames) To UBound(leftNames)
        leftName = CStr(leftNames(probeIndex))
        rightName = CStr(rightNames(probeIndex))
        leftMetadata = "unavailable"
        rightMetadata = "unavailable"
        probeResult = "not run"
        Set leftRs = Nothing
        Err.Clear
        Set leftRs = db.OpenRecordset("SELECT " & SQL_FIELD_PROJECT_IDENTITY & " FROM [" & Replace(leftName, "]", "]]") & "] WHERE 1=0;", dbOpenSnapshot)
        If Err.Number = 0 And Not leftRs Is Nothing Then
            leftMetadata = "DAO.Type=" & CStr(leftRs.fields(FIELD_PROJECT_IDENTITY).Type) & "; Size=" & CStr(leftRs.fields(FIELD_PROJECT_IDENTITY).Size)
            leftRs.Close
        Else
            leftMetadata = "metadata error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
        End If
        Set leftRs = Nothing
        Set rightRs = Nothing
        Err.Clear
        Set rightRs = db.OpenRecordset("SELECT " & SQL_FIELD_PROJECT_IDENTITY & " FROM [" & Replace(rightName, "]", "]]") & "] WHERE 1=0;", dbOpenSnapshot)
        If Err.Number = 0 And Not rightRs Is Nothing Then
            rightMetadata = "DAO.Type=" & CStr(rightRs.fields(FIELD_PROJECT_IDENTITY).Type) & "; Size=" & CStr(rightRs.fields(FIELD_PROJECT_IDENTITY).Size)
            rightRs.Close
        Else
            rightMetadata = "metadata error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
        End If
        Set rightRs = Nothing
        Set joinRs = Nothing
        Err.Clear
        probeSql = "SELECT L." & SQL_FIELD_PROJECT_IDENTITY & " AS " & SQL_FIELD_LEFT_PROJECT_IDENTITY & ",R." & SQL_FIELD_PROJECT_IDENTITY & " AS " & SQL_FIELD_RIGHT_PROJECT_IDENTITY & " FROM [" & Replace(leftName, "]", "]]") & "] AS L INNER JOIN [" & Replace(rightName, "]", "]]") & "] AS R ON L." & SQL_FIELD_PROJECT_IDENTITY & "=R." & SQL_FIELD_PROJECT_IDENTITY & " WHERE 1=0;"
        Set joinRs = db.OpenRecordset(probeSql, dbOpenSnapshot)
        probeErrNumber = Err.Number
        probeErrSource = Err.Source
        probeErrDescription = Err.Description
        If probeErrNumber = 0 And Not joinRs Is Nothing Then
            probeResult = "PASS"
            joinRs.Close
        Else
            probeResult = "FAIL Error " & CStr(probeErrNumber) & " from " & probeErrSource & ": " & probeErrDescription
        End If
        Set joinRs = Nothing
        detailText = detailText & vbCrLf & CStr(probeIndex + 1) & ". " & CStr(probeLabels(probeIndex)) & " | Left [" & leftName & "].ProjectIdentity: " & leftMetadata & " | Right [" & rightName & "].ProjectIdentity: " & rightMetadata & " | Probe: " & probeResult
    Next probeIndex
    ReferenceProjectIdentityJoinDiagnostics = detailText
End Function

Private Function ReferenceAddInReviewUnionDiagnostics(ByVal db As DAO.Database) As String
    'Runs only after DBInv_AddIns+VBAPrjs+VBAProcs_ReviewSummary executable validation fails with DAO 3464; successful A4 runs incur no branch or UNION probe work.
    On Error Resume Next
    Dim queryName As String
    Dim summarySql As String
    Dim summaryWithoutOrder As String
    Dim branchSql As Variant
    Dim branchZeroSql(0 To 3) As String
    Dim branchLabels As Variant
    Dim branchIndex As Long
    Dim branchCount As Long
    Dim hasTypeAnchor As Boolean
    Dim orderPosition As Long
    Dim probeRs As DAO.Recordset
    Dim probeErrNumber As Long
    Dim probeErrSource As String
    Dim probeErrDescription As String
    Dim detailText As String
    Dim branchResult As String
    Dim combinedSql As String
    Dim unionFailed As Boolean
    Dim branchFailed As Boolean
    Dim expressionDetail As String
    Dim fieldDetail As String
    Dim orderDetail As String
    Dim actualDetail As String
    queryName = ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Review_Summary")
    Err.Clear
    summarySql = Trim$(db.QueryDefs(queryName).sql)
    If Err.Number <> 0 Then
        ReferenceAddInReviewUnionDiagnostics = queryName & " branch/UNION diagnostics could not read the saved QueryDef SQL. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
        Exit Function
    End If
    If Right$(summarySql, 1) = ";" Then summarySql = Left$(summarySql, Len(summarySql) - 1)
    summarySql = Replace(summarySql, vbCr, " ")
    summarySql = Replace(summarySql, vbLf, " ")
    summarySql = Replace(summarySql, vbTab, " ")
    Do While InStr(1, summarySql, "  ", vbBinaryCompare) > 0
        summarySql = Replace(summarySql, "  ", " ")
    Loop
    orderPosition = InStrRev(summarySql, " ORDER BY ", -1, vbTextCompare)
    If orderPosition > 0 Then
        summaryWithoutOrder = Left$(summarySql, orderPosition - 1)
    Else
        summaryWithoutOrder = summarySql
    End If
    branchSql = Split(summaryWithoutOrder, " UNION ALL ", -1, vbTextCompare)
    branchCount = UBound(branchSql) - LBound(branchSql) + 1
    'The canonical query begins with one zero-row type-anchor SELECT. Failure diagnostics operate on the four real data branches, not on that schema-only anchor.
    hasTypeAnchor = (branchCount = 5 And InStr(1, CStr(branchSql(0)), "WHERE False", vbTextCompare) > 0)
    If hasTypeAnchor Then
        For branchIndex = 0 To 3
            branchSql(branchIndex) = branchSql(branchIndex + 1)
        Next branchIndex
    ElseIf branchCount <> 4 Then
        ReferenceAddInReviewUnionDiagnostics = queryName & " branch/UNION diagnostics expected 4 data branches, plus the optional canonical type anchor, but found " & CStr(branchCount) & " SELECT branches."
        Exit Function
    End If
    branchLabels = Array("VBE Add-in", "Access Add-in", "COM Add-in", "Loaded VBA Project")
    detailText = queryName & " 3464 branch/UNION isolation probes:"
    For branchIndex = 0 To 3
        branchZeroSql(branchIndex) = Trim$(CStr(branchSql(branchIndex))) & " AND 1=0"
        Set probeRs = Nothing
        Err.Clear
        Set probeRs = db.OpenRecordset(branchZeroSql(branchIndex) & ";", dbOpenSnapshot)
        probeErrNumber = Err.Number
        probeErrSource = Err.Source
        probeErrDescription = Err.Description
        If probeErrNumber = 0 And Not probeRs Is Nothing Then
            branchResult = "PASS"
            probeRs.Close
        Else
            branchResult = "FAIL Error " & CStr(probeErrNumber) & " from " & probeErrSource & ": " & probeErrDescription
            branchFailed = True
        End If
        Set probeRs = Nothing
        detailText = detailText & vbCrLf & CStr(branchIndex + 1) & ". Branch " & CStr(branchLabels(branchIndex)) & ": " & branchResult
        If branchResult <> "PASS" Then
            expressionDetail = ReferenceBranchExpressionDiagnostics(db, Trim$(CStr(branchSql(branchIndex))), CStr(branchLabels(branchIndex)))
            If Len(expressionDetail) > 0 Then detailText = detailText & vbCrLf & expressionDetail
        End If
    Next branchIndex
    If branchFailed = False Then
        combinedSql = branchZeroSql(0) & " UNION ALL " & branchZeroSql(1)
        detailText = detailText & vbCrLf & ReferenceUnionProbeResult(db, combinedSql, "Branches 1+2")
        If InStr(1, detailText, "Branches 1+2: FAIL", vbTextCompare) > 0 Then unionFailed = True
        combinedSql = combinedSql & " UNION ALL " & branchZeroSql(2)
        detailText = detailText & vbCrLf & ReferenceUnionProbeResult(db, combinedSql, "Branches 1+2+3")
        If InStr(1, detailText, "Branches 1+2+3: FAIL", vbTextCompare) > 0 Then unionFailed = True
        combinedSql = combinedSql & " UNION ALL " & branchZeroSql(3)
        detailText = detailText & vbCrLf & ReferenceUnionProbeResult(db, combinedSql, "Branches 1+2+3+4 without ORDER BY")
        If InStr(1, detailText, "Branches 1+2+3+4 without ORDER BY: FAIL", vbTextCompare) > 0 Then unionFailed = True
        If unionFailed Then
            fieldDetail = ReferenceBranchFieldTypeDiagnostics(db, branchZeroSql(0), branchZeroSql(1), branchZeroSql(2), branchZeroSql(3), branchLabels)
            If Len(fieldDetail) > 0 Then detailText = detailText & vbCrLf & fieldDetail
        Else
            detailText = detailText & vbCrLf & "All four branches and the complete UNION ALL pass without ORDER BY; the remaining executable difference is the final ORDER BY clause."
            orderDetail = ReferenceAddInReviewOrderByDiagnostics(db, combinedSql)
            If Len(orderDetail) > 0 Then detailText = detailText & vbCrLf & orderDetail
        End If
    End If
    actualDetail = ReferenceAddInReviewActualDataDiagnostics(db, branchSql, branchLabels, summarySql)
    If Len(actualDetail) > 0 Then detailText = detailText & vbCrLf & actualDetail
    ReferenceAddInReviewUnionDiagnostics = detailText
End Function

Private Function ReferenceAddInReviewActualDataDiagnostics(ByVal db As DAO.Database, ByVal branchSql As Variant, ByVal branchLabels As Variant, ByVal summarySql As String) As String
    'Runs only after DBInv_AddIns+VBAPrjs+VBAProcs_ReviewSummary validation has already failed with DAO 3464; it forces actual returned values so successful A4 runs incur no probe work.
    On Error Resume Next
    Dim dependencyNames As Variant
    Dim dependencyLabels As Variant
    Dim dependencyIndex As Long
    Dim branchIndex As Long
    Dim branchPassed(0 To 3) As Boolean
    Dim probePassed As Boolean
    Dim allBranchesPassed As Boolean
    Dim combinedSql As String
    Dim detailText As String
    Dim expressionDetail As String
    dependencyNames = Array(ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Project_Procedure_Counts"), ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Project_Conflict_Counts"))
    dependencyLabels = Array("Dependency " & CStr(dependencyNames(0)), "Dependency " & CStr(dependencyNames(1)))
    detailText = ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Review_Summary") & " actual-data 3464 isolation probes:"
    For dependencyIndex = LBound(dependencyNames) To UBound(dependencyNames)
        detailText = detailText & vbCrLf & ReferenceActualDataProbeResult(db, "SELECT * FROM [" & Replace(CStr(dependencyNames(dependencyIndex)), "]", "]]") & "]", CStr(dependencyLabels(dependencyIndex)), probePassed)
    Next dependencyIndex
    allBranchesPassed = True
    For branchIndex = 0 To 3
        detailText = detailText & vbCrLf & ReferenceActualDataProbeResult(db, Trim$(CStr(branchSql(branchIndex))), "Actual branch " & CStr(branchIndex + 1) & " " & CStr(branchLabels(branchIndex)), branchPassed(branchIndex))
        If branchPassed(branchIndex) = False Then
            allBranchesPassed = False
            expressionDetail = ReferenceBranchActualExpressionDiagnostics(db, Trim$(CStr(branchSql(branchIndex))), CStr(branchLabels(branchIndex)))
            If Len(expressionDetail) > 0 Then detailText = detailText & vbCrLf & expressionDetail
        End If
    Next branchIndex
    If allBranchesPassed Then
        combinedSql = Trim$(CStr(branchSql(0))) & " UNION ALL " & Trim$(CStr(branchSql(1)))
        detailText = detailText & vbCrLf & ReferenceActualDataProbeResult(db, combinedSql, "Actual branches 1+2", probePassed)
        If probePassed Then
            combinedSql = combinedSql & " UNION ALL " & Trim$(CStr(branchSql(2)))
            detailText = detailText & vbCrLf & ReferenceActualDataProbeResult(db, combinedSql, "Actual branches 1+2+3", probePassed)
        End If
        If probePassed Then
            combinedSql = combinedSql & " UNION ALL " & Trim$(CStr(branchSql(3)))
            detailText = detailText & vbCrLf & ReferenceActualDataProbeResult(db, combinedSql, "Actual branches 1+2+3+4 without ORDER BY", probePassed)
        End If
        If probePassed Then detailText = detailText & vbCrLf & ReferenceActualDataProbeResult(db, summarySql, "Actual complete production SQL including ORDER BY", probePassed)
    End If
    ReferenceAddInReviewActualDataDiagnostics = detailText
End Function

Private Function ReferenceActualDataProbeResult(ByVal db As DAO.Database, ByVal probeSql As String, ByVal probeLabel As String, ByRef probePassed As Boolean) As String
    On Error Resume Next
    Dim probeRs As DAO.Recordset
    Dim fld As DAO.Field
    Dim probeValue As Variant
    Dim rowCount As Long
    Dim probeErrNumber As Long
    Dim probeErrSource As String
    Dim probeErrDescription As String
    probePassed = False
    Set probeRs = Nothing
    Err.Clear
    Set probeRs = db.OpenRecordset(probeSql & ";", dbOpenSnapshot)
    probeErrNumber = Err.Number
    probeErrSource = Err.Source
    probeErrDescription = Err.Description
    If probeErrNumber <> 0 Or probeRs Is Nothing Then
        ReferenceActualDataProbeResult = probeLabel & ": FAIL opening recordset. Error " & CStr(probeErrNumber) & " from " & probeErrSource & ": " & probeErrDescription
        Exit Function
    End If
    Do While probeRs.EOF = False
        rowCount = rowCount + 1
        For Each fld In probeRs.fields
            Err.Clear
            probeValue = fld.value
            probeErrNumber = Err.Number
            probeErrSource = Err.Source
            probeErrDescription = Err.Description
            If probeErrNumber <> 0 Then
                ReferenceActualDataProbeResult = probeLabel & ": FAIL evaluating row " & CStr(rowCount) & ", field [" & fld.Name & "]. Error " & CStr(probeErrNumber) & " from " & probeErrSource & ": " & probeErrDescription
                probeRs.Close
                Set probeRs = Nothing
                Exit Function
            End If
        Next fld
        Err.Clear
        probeRs.MoveNext
        probeErrNumber = Err.Number
        probeErrSource = Err.Source
        probeErrDescription = Err.Description
        If probeErrNumber <> 0 Then
            ReferenceActualDataProbeResult = probeLabel & ": FAIL advancing after row " & CStr(rowCount) & ". Error " & CStr(probeErrNumber) & " from " & probeErrSource & ": " & probeErrDescription
            probeRs.Close
            Set probeRs = Nothing
            Exit Function
        End If
    Loop
    probePassed = True
    ReferenceActualDataProbeResult = probeLabel & ": PASS | Rows evaluated=" & CStr(rowCount)
    probeRs.Close
    Set probeRs = Nothing
End Function

Private Function ReferenceBranchActualExpressionDiagnostics(ByVal db As DAO.Database, ByVal branchSql As String, ByVal branchLabel As String) As String
    On Error Resume Next
    Dim fromPosition As Long
    Dim selectList As String
    Dim fromClause As String
    Dim expressions As Collection
    Dim expressionText As Variant
    Dim expressionIndex As Long
    Dim expressionPassed As Boolean
    Dim expressionResult As String
    Dim detailText As String
    fromPosition = ReferenceTopLevelFromPosition(branchSql)
    If fromPosition = 0 Then
        ReferenceBranchActualExpressionDiagnostics = "Actual-data expression isolation for branch " & branchLabel & " could not locate the top-level FROM clause."
        Exit Function
    End If
    selectList = Mid$(branchSql, Len("SELECT ") + 1, fromPosition - Len("SELECT ") - 1)
    fromClause = Mid$(branchSql, fromPosition)
    Set expressions = ReferenceTopLevelSelectExpressions(selectList)
    detailText = "Actual-data expression isolation for branch " & branchLabel & ":"
    For Each expressionText In expressions
        expressionIndex = expressionIndex + 1
        expressionResult = ReferenceActualDataProbeResult(db, "SELECT " & CStr(expressionText) & fromClause, "Expression " & CStr(expressionIndex), expressionPassed)
        detailText = detailText & vbCrLf & expressionResult & " | " & CStr(expressionText)
        If expressionPassed = False Then Exit For
    Next expressionText
    ReferenceBranchActualExpressionDiagnostics = detailText
End Function

Private Function ReferenceAddInReviewOrderByDiagnostics(ByVal db As DAO.Database, ByVal unionSql As String) As String
    'Runs only after every DBInv_AddIns+VBAPrjs+VBAProcs_ReviewSummary branch and the complete zero-row UNION ALL pass without ORDER BY following DAO 3464.
    On Error Resume Next
    Dim sortClauses As Variant
    Dim sortIndex As Long
    Dim detailText As String
    Dim fieldMetadata As String
    Dim probeRs As DAO.Recordset
    Dim probeErrNumber As Long
    Dim probeErrSource As String
    Dim probeErrDescription As String
    detailText = ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Review_Summary") & " ORDER BY 3464 isolation probes:"
    Set probeRs = Nothing
    Err.Clear
    Set probeRs = db.OpenRecordset(unionSql & ";", dbOpenSnapshot)
    probeErrNumber = Err.Number
    probeErrSource = Err.Source
    probeErrDescription = Err.Description
    If probeErrNumber = 0 And Not probeRs Is Nothing Then
        fieldMetadata = SQL_FIELD_SOURCE_KIND & " DAO.Type=" & CStr(probeRs.fields(FIELD_SOURCE_KIND).Type) & "; Size=" & CStr(probeRs.fields(FIELD_SOURCE_KIND).Size) & " | " & SQL_FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME & " DAO.Type=" & CStr(probeRs.fields(FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME).Type) & "; Size=" & CStr(probeRs.fields(FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME).Size) & " | " & SQL_FIELD_REVIEW_IDENTITY & " DAO.Type=" & CStr(probeRs.fields(FIELD_REVIEW_IDENTITY).Type) & "; Size=" & CStr(probeRs.fields(FIELD_REVIEW_IDENTITY).Size)
        probeRs.Close
    Else
        fieldMetadata = "complete UNION metadata unavailable: Error " & CStr(probeErrNumber) & " from " & probeErrSource & ": " & probeErrDescription
    End If
    Set probeRs = Nothing
    detailText = detailText & vbCrLf & "Complete UNION sort-field metadata: " & fieldMetadata
    sortClauses = Array(SQL_FIELD_SOURCE_KIND, SQL_FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME, SQL_FIELD_REVIEW_IDENTITY, SQL_FIELD_SOURCE_KIND & "," & SQL_FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME, SQL_FIELD_SOURCE_KIND & "," & SQL_FIELD_REVIEW_IDENTITY, SQL_FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME & "," & SQL_FIELD_REVIEW_IDENTITY, SQL_FIELD_SOURCE_KIND & "," & SQL_FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME & "," & SQL_FIELD_REVIEW_IDENTITY)
    For sortIndex = LBound(sortClauses) To UBound(sortClauses)
        detailText = detailText & vbCrLf & ReferenceUnionProbeResult(db, unionSql & " ORDER BY " & CStr(sortClauses(sortIndex)), "ORDER BY " & CStr(sortClauses(sortIndex)))
    Next sortIndex
    ReferenceAddInReviewOrderByDiagnostics = detailText
End Function

Private Function ReferenceUnionProbeResult(ByVal db As DAO.Database, ByVal probeSql As String, ByVal probeLabel As String) As String
    On Error Resume Next
    Dim probeRs As DAO.Recordset
    Dim probeErrNumber As Long
    Dim probeErrSource As String
    Dim probeErrDescription As String
    Set probeRs = Nothing
    Err.Clear
    Set probeRs = db.OpenRecordset(probeSql & ";", dbOpenSnapshot)
    probeErrNumber = Err.Number
    probeErrSource = Err.Source
    probeErrDescription = Err.Description
    If probeErrNumber = 0 And Not probeRs Is Nothing Then
        ReferenceUnionProbeResult = probeLabel & ": PASS"
        probeRs.Close
    Else
        ReferenceUnionProbeResult = probeLabel & ": FAIL Error " & CStr(probeErrNumber) & " from " & probeErrSource & ": " & probeErrDescription
    End If
    Set probeRs = Nothing
End Function

Private Function ReferenceBranchExpressionDiagnostics(ByVal db As DAO.Database, ByVal branchSql As String, ByVal branchLabel As String) As String
    On Error Resume Next
    Dim fromPosition As Long
    Dim selectList As String
    Dim fromClause As String
    Dim expressions As Collection
    Dim expressionText As Variant
    Dim expressionIndex As Long
    Dim probeRs As DAO.Recordset
    Dim probeErrNumber As Long
    Dim probeErrSource As String
    Dim probeErrDescription As String
    Dim detailText As String
    fromPosition = ReferenceTopLevelFromPosition(branchSql)
    If fromPosition = 0 Then
        ReferenceBranchExpressionDiagnostics = "Calculated-expression isolation for branch " & branchLabel & " could not locate the top-level FROM clause."
        Exit Function
    End If
    selectList = Mid$(branchSql, Len("SELECT ") + 1, fromPosition - Len("SELECT ") - 1)
    fromClause = Mid$(branchSql, fromPosition)
    Set expressions = ReferenceTopLevelSelectExpressions(selectList)
    detailText = "Calculated-expression isolation for branch " & branchLabel & ":"
    For Each expressionText In expressions
        expressionIndex = expressionIndex + 1
        Set probeRs = Nothing
        Err.Clear
        Set probeRs = db.OpenRecordset("SELECT " & CStr(expressionText) & fromClause & " AND 1=0;", dbOpenSnapshot)
        probeErrNumber = Err.Number
        probeErrSource = Err.Source
        probeErrDescription = Err.Description
        If probeErrNumber = 0 And Not probeRs Is Nothing Then
            detailText = detailText & vbCrLf & "Expression " & CStr(expressionIndex) & ": PASS | " & CStr(expressionText)
            probeRs.Close
        Else
            detailText = detailText & vbCrLf & "Expression " & CStr(expressionIndex) & ": FAIL Error " & CStr(probeErrNumber) & " from " & probeErrSource & ": " & probeErrDescription & " | " & CStr(expressionText)
        End If
        Set probeRs = Nothing
    Next expressionText
    ReferenceBranchExpressionDiagnostics = detailText
End Function

Private Function ReferenceTopLevelFromPosition(ByVal sqlText As String) As Long
    Dim charIndex As Long
    Dim parenDepth As Long
    Dim inString As Boolean
    Dim inBracket As Boolean
    Dim currentChar As String
    Dim nextChar As String
    For charIndex = 1 To Len(sqlText)
        currentChar = Mid$(sqlText, charIndex, 1)
        If inString Then
            If currentChar = "'" Then
                If charIndex < Len(sqlText) Then nextChar = Mid$(sqlText, charIndex + 1, 1) Else nextChar = vbNullString
                If nextChar = "'" Then
                    charIndex = charIndex + 1
                Else
                    inString = False
                End If
            End If
        ElseIf inBracket Then
            If currentChar = "]" Then inBracket = False
        Else
            Select Case currentChar
                Case "'": inString = True
                Case "[": inBracket = True
                Case "(": parenDepth = parenDepth + 1
                Case ")": If parenDepth > 0 Then parenDepth = parenDepth - 1
            End Select
            If parenDepth = 0 And inString = False And inBracket = False Then
                If StrComp(Mid$(sqlText, charIndex, 6), " FROM ", vbTextCompare) = 0 Then
                    ReferenceTopLevelFromPosition = charIndex
                    Exit Function
                End If
            End If
        End If
    Next charIndex
End Function

Private Function ReferenceTopLevelSelectExpressions(ByVal selectList As String) As Collection
    ' INTENTIONAL ORDERED VIEW: results preserves the deterministic sequence required by ReferenceTopLevelSelectExpressions; keyed-only or unordered storage would not preserve the consumer semantics.
    Dim results As New Collection
    Dim charIndex As Long
    Dim expressionStart As Long
    Dim parenDepth As Long
    Dim inString As Boolean
    Dim inBracket As Boolean
    Dim currentChar As String
    Dim nextChar As String
    expressionStart = 1
    For charIndex = 1 To Len(selectList)
        currentChar = Mid$(selectList, charIndex, 1)
        If inString Then
            If currentChar = "'" Then
                If charIndex < Len(selectList) Then nextChar = Mid$(selectList, charIndex + 1, 1) Else nextChar = vbNullString
                If nextChar = "'" Then
                    charIndex = charIndex + 1
                Else
                    inString = False
                End If
            End If
        ElseIf inBracket Then
            If currentChar = "]" Then inBracket = False
        Else
            Select Case currentChar
                Case "'": inString = True
                Case "[": inBracket = True
                Case "(": parenDepth = parenDepth + 1
                Case ")": If parenDepth > 0 Then parenDepth = parenDepth - 1
                Case ","
                    If parenDepth = 0 Then
                        results.Add Trim$(Mid$(selectList, expressionStart, charIndex - expressionStart))
                        expressionStart = charIndex + 1
                    End If
            End Select
        End If
    Next charIndex
    If expressionStart <= Len(selectList) Then results.Add Trim$(Mid$(selectList, expressionStart))
    Set ReferenceTopLevelSelectExpressions = results
End Function

Private Function ReferenceBranchFieldTypeDiagnostics(ByVal db As DAO.Database, ByVal branchOneSql As String, ByVal branchTwoSql As String, ByVal branchThreeSql As String, ByVal branchFourSql As String, ByVal branchLabels As Variant) As String
    On Error Resume Next
    Dim fieldMaps(0 To 3) As Object
    Dim branchZeroSql As Variant
    Dim fieldNames As Object
    Dim probeRs As DAO.Recordset
    Dim branchIndex As Long
    Dim fieldIndex As Long
    Dim fieldCount As Long
    Dim detailText As String
    Dim fieldName As String
    Dim metadataText As String
    Dim mismatchText As String
    branchZeroSql = Array(branchOneSql, branchTwoSql, branchThreeSql, branchFourSql)
    ' INTENTIONAL KEYED LOOKUP: fieldNames, fieldMaps provides keyed membership/lookup over source rows or state already read by ReferenceBranchFieldTypeDiagnostics; it is a derived access path, not a second
    ' authority.
    Set fieldNames = CreateObject("Scripting.Dictionary")
    For branchIndex = 0 To 3
        Set fieldMaps(branchIndex) = CreateObject("Scripting.Dictionary")
        Set probeRs = Nothing
        Err.Clear
        Set probeRs = db.OpenRecordset(branchZeroSql(branchIndex) & ";", dbOpenSnapshot)
        If Err.Number <> 0 Or probeRs Is Nothing Then
            ReferenceBranchFieldTypeDiagnostics = "Branch result-field type diagnostics could not reopen branch " & CStr(branchIndex + 1) & " (" & CStr(branchLabels(branchIndex)) & "). Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
            Exit Function
        End If
        If branchIndex = 0 Then fieldCount = probeRs.fields.count
        For fieldIndex = 0 To probeRs.fields.count - 1
            fieldName = probeRs.fields(fieldIndex).Name
            metadataText = "Type=" & CStr(probeRs.fields(fieldIndex).Type) & ";Size=" & CStr(probeRs.fields(fieldIndex).Size)
            fieldMaps(branchIndex).item(CStr(fieldIndex)) = metadataText
            If branchIndex = 0 Then fieldNames.item(CStr(fieldIndex)) = fieldName
        Next fieldIndex
        probeRs.Close
        Set probeRs = Nothing
    Next branchIndex
    detailText = "Branch result-field DAO type matrix (B1=VBE, B2=Access, B3=COM, B4=Loaded Project):"
    For fieldIndex = 0 To fieldCount - 1
        mismatchText = vbNullString
        If StrComp(CStr(fieldMaps(0).item(CStr(fieldIndex))), CStr(fieldMaps(1).item(CStr(fieldIndex))), vbBinaryCompare) <> 0 Or StrComp(CStr(fieldMaps(0).item(CStr(fieldIndex))), CStr(fieldMaps(2).item(CStr(fieldIndex))), vbBinaryCompare) <> 0 Or StrComp(CStr(fieldMaps(0).item(CStr(fieldIndex))), CStr(fieldMaps(3).item(CStr(fieldIndex))), vbBinaryCompare) <> 0 Then mismatchText = " | DAO metadata differs by branch"
        detailText = detailText & vbCrLf & CStr(fieldIndex + 1) & ". " & CStr(fieldNames.item(CStr(fieldIndex))) & " | B1 " & CStr(fieldMaps(0).item(CStr(fieldIndex))) & " | B2 " & CStr(fieldMaps(1).item(CStr(fieldIndex))) & " | B3 " & CStr(fieldMaps(2).item(CStr(fieldIndex))) & " | B4 " & CStr(fieldMaps(3).item(CStr(fieldIndex))) & mismatchText
    Next fieldIndex
    ReferenceBranchFieldTypeDiagnostics = detailText
End Function

Private Function ReferencePersistedTokenAvailable(ByVal db As DAO.Database, ByVal propertyName As String, ByRef storedToken As String) As Boolean
    Dim propertyExists As Boolean
    storedToken = DatabaseTextPropertyValue(db, propertyName, propertyExists)
    ReferencePersistedTokenAvailable = propertyExists And Len(storedToken) > 0
End Function

Private Sub AppendReferencePhysicalObjectNames(ByVal names As Collection, ByVal objectNames As Collection, ByVal objectTypeCode As String)
    Dim item As Variant
    For Each item In objectNames
        names.Add objectTypeCode & "|" & CStr(item)
    Next item
End Sub

Private Sub AddReferenceSchemaField(ByRef rows As Collection, ByVal fieldName As String, ByVal fieldType As DAO.DataTypeEnum, Optional ByVal fieldSize As Long = 0, Optional ByVal autoNumber As Boolean = False, Optional ByVal requiredField As Variant, Optional ByVal allowZeroLength As Variant, Optional ByVal validationRule As Variant, Optional ByVal validationText As Variant)
    Dim requiredValue As Variant
    Dim allowZeroLengthValue As Variant
    Dim validationRuleValue As Variant
    Dim validationTextValue As Variant
    If isMissing(requiredField) Then requiredValue = Null Else requiredValue = requiredField
    If isMissing(allowZeroLength) Then allowZeroLengthValue = Null Else allowZeroLengthValue = allowZeroLength
    If isMissing(validationRule) Then validationRuleValue = Null Else validationRuleValue = validationRule
    If isMissing(validationText) Then validationTextValue = Null Else validationTextValue = validationText
    rows.Add Array(fieldName, CLng(fieldType), fieldSize, autoNumber, requiredValue, allowZeroLengthValue, validationRuleValue, validationTextValue)
End Sub

Private Function ReferenceTableSchemaDefinitions(ByVal tableName As String) As Collection
    'Single physical schema-assembly authority for all reference and Add-In inventory table creation and repair; field-name identity literals come from x__UTIL_ADC_DB_DEFINITIONS.
    ' CANONICAL AUTHORITY: rows owns the in-memory definition set constructed by ReferenceTableSchemaDefinitions; downstream consumers derive views from this set rather than create a competing authority.
    Dim rows As New Collection
    Select Case UCase$(tableName)
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ReferenceManager"))
            AddReferenceSchemaField rows, FIELD_ID, dbLong, 0, True
            AddReferenceSchemaField rows, FIELD_REQUESTED_ACTIVE, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_ORIGINAL_ACTIVE, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_REFERENCE_SOURCE, dbText, 20, False
            AddReferenceSchemaField rows, FIELD_REFERENCE_PRIORITY, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_REFERENCE_NAME, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_REFERENCE_DESCRIPTION, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_REFERENCE_GUID, dbText, 50, False
            AddReferenceSchemaField rows, FIELD_MAJOR_VERSION, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_MINOR_VERSION, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_REFERENCE_IDENTITY, dbText, 255, False, True, False
            AddReferenceSchemaField rows, FIELD_CAN_CHANGE, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_PROTECTION_REASON, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_HEALTH_STATUS, dbText, 30, False
            AddReferenceSchemaField rows, FIELD_IS_BUILT_IN, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_IS_BROKEN, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_FILE_EXISTS, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_IS_REGISTERED, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_FULL_PATH, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_REGISTERED_PATH, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_REGISTRATION_SCOPE, dbText, 30, False
            AddReferenceSchemaField rows, FIELD_REGISTRATION_VIEW, dbText, 20, False
            AddReferenceSchemaField rows, FIELD_PLATFORM, dbText, 20, False
            AddReferenceSchemaField rows, FIELD_LAST_REFRESHED_ON, dbDate, 0, False
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects"))
            AddReferenceSchemaField rows, FIELD_ID, dbLong, 0, True
            AddReferenceSchemaField rows, FIELD_PROJECT_ORDER, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_PROJECT_NAME, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_PROJECT_ROLE, dbText, 50, False
            AddReferenceSchemaField rows, FIELD_IS_ACTIVE_VBE_PROJECT, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_COMPONENT_COUNT, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_COMPONENTS_READABLE, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_SOURCE_READABLE, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_PROTECTION_STATE, dbText, 30, False
            AddReferenceSchemaField rows, FIELD_PROTECTION_CODE, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_PROJECT_TYPE, dbText, 50, False
            AddReferenceSchemaField rows, FIELD_PROJECT_TYPE_CODE, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_SCAN_STATUS, dbText, 40, False
            AddReferenceSchemaField rows, FIELD_SOURCE_ANALYSIS_STATUS, dbText, 40, False
            AddReferenceSchemaField rows, FIELD_MODULE_ROSTER_SIGNATURE, dbText, 40, False
            AddReferenceSchemaField rows, FIELD_PROJECT_REFERENCE_COUNT, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_MODULE_INVENTORY_COUNT, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_PROCEDURE_INVENTORY_COUNT, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_DEPENDENCY_COUNT, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_LAST_SOURCE_SCANNED_ON, dbDate, 0, False
            AddReferenceSchemaField rows, FIELD_SOURCE_ANALYSIS_DIAGNOSTIC_TEXT, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_PROJECT_FILE_PATH, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_FILE_EXISTS, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_PROJECT_FILE_NAME, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_PROJECT_FILE_BASE_NAME, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_PROJECT_IDENTITY, dbText, 255, False, True, False
            AddReferenceSchemaField rows, FIELD_PROJECT_PATH_KEY, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_PROJECT_NAME_KEY, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_PROJECT_FILE_BASE_NAME_KEY, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_DIAGNOSTIC_TEXT, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_INVENTORY_STATUS, dbText, 20, False
            AddReferenceSchemaField rows, FIELD_LAST_SCAN_BATCH_ID, dbText, 100, False
            AddReferenceSchemaField rows, FIELD_LAST_SCANNED_ON, dbDate, 0, False
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_VBEAddIns")), UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_COMAddIns"))
            AddReferenceSchemaField rows, FIELD_ID, dbLong, 0, True
            AddReferenceSchemaField rows, FIELD_ADD_IN_ORDER, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_DESCRIPTION, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_PROG_ID, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_GUID, dbText, 100, False
            AddReferenceSchemaField rows, FIELD_IS_REGISTERED, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_IS_CONNECTED, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_OBJECT_AVAILABLE, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_SCAN_STATUS, dbText, 40, False
            AddReferenceSchemaField rows, FIELD_ADD_IN_IDENTITY, dbText, 255, False, True, False
            AddReferenceSchemaField rows, FIELD_DIAGNOSTIC_TEXT, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_INVENTORY_STATUS, dbText, 20, False
            AddReferenceSchemaField rows, FIELD_LAST_SCAN_BATCH_ID, dbText, 100, False
            AddReferenceSchemaField rows, FIELD_LAST_SCANNED_ON, dbDate, 0, False
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AccessAddIns"))
            AddReferenceSchemaField rows, FIELD_ID, dbLong, 0, True
            AddReferenceSchemaField rows, FIELD_ADD_IN_ORDER, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_ADD_IN_IDENTITY, dbText, 255, False, True, False
            AddReferenceSchemaField rows, FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_REGISTRATION_SUBKEY, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_ADD_IN_LIBRARY, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_ENTRY_POINT_EXPRESSION, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_ENTRY_POINT_PROCEDURE, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_ENTRY_POINT_EXISTS, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_ENTRY_POINT_EXTERNALLY_CALLABLE, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_PROJECT_IDENTITY, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_PROJECT_NAME, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_PROJECT_ROLE, dbText, 50, False
            AddReferenceSchemaField rows, FIELD_PROJECT_FILE_PATH, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_FILE_EXISTS, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_SCAN_STATUS, dbText, 40, False
            AddReferenceSchemaField rows, FIELD_DIAGNOSTIC_TEXT, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_INVENTORY_STATUS, dbText, 20, False
            AddReferenceSchemaField rows, FIELD_LAST_SCAN_BATCH_ID, dbText, 100, False
            AddReferenceSchemaField rows, FIELD_LAST_SCANNED_ON, dbDate, 0, False
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjectReferences"))
            AddReferenceSchemaField rows, FIELD_ID, dbLong, 0, True
            AddReferenceSchemaField rows, FIELD_PROJECT_NAME, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_PROJECT_ROLE, dbText, 50, False
            AddReferenceSchemaField rows, FIELD_REFERENCE_ORDER, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_REFERENCE_NAME, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_REFERENCE_DESCRIPTION, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_REFERENCE_GUID, dbText, 100, False
            AddReferenceSchemaField rows, FIELD_MAJOR_VERSION, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_MINOR_VERSION, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_REFERENCE_KIND, dbText, 50, False
            AddReferenceSchemaField rows, FIELD_REFERENCE_KIND_CODE, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_IS_BUILT_IN, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_IS_BROKEN, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_FULL_PATH, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_FILE_EXISTS, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_MATCH_STATUS, dbText, 50, False
            AddReferenceSchemaField rows, FIELD_MATCH_METHOD, dbText, 50, False
            AddReferenceSchemaField rows, FIELD_IS_AUTHORITATIVE_MATCH, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_REFERENCED_PROJECT_NAME, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_REFERENCED_PROJECT_ROLE, dbText, 50, False
            AddReferenceSchemaField rows, FIELD_PROJECT_IDENTITY, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_REFERENCE_IDENTITY, dbText, 255, False, True, False
            AddReferenceSchemaField rows, FIELD_REFERENCED_PROJECT_IDENTITY, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_DIAGNOSTIC_TEXT, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_INVENTORY_STATUS, dbText, 20, False
            AddReferenceSchemaField rows, FIELD_LAST_SCAN_BATCH_ID, dbText, 100, False
            AddReferenceSchemaField rows, FIELD_LAST_SCANNED_ON, dbDate, 0, False
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAModules"))
            AddReferenceSchemaField rows, FIELD_ID, dbLong, 0, True
            AddReferenceSchemaField rows, FIELD_PROJECT_NAME, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_PROJECT_ROLE, dbText, 50, False
            AddReferenceSchemaField rows, FIELD_MODULE_ORDER, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_MODULE_NAME, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_COMPONENT_TYPE, dbText, 50, False
            AddReferenceSchemaField rows, FIELD_COMPONENT_TYPE_CODE, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_SOURCE_READABLE, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_SOURCE_LINE_COUNT, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_DECLARATION_LINE_COUNT, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_PROCEDURE_COUNT, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_PUBLIC_PROCEDURE_COUNT, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_EXTERNALLY_CALLABLE_PROCEDURE_COUNT, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_HAS_OPTION_PRIVATE_MODULE, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_SOURCE_FINGERPRINT, dbText, 40, False
            AddReferenceSchemaField rows, FIELD_SCANNED_CHAT_GPTID, dbText, 40, False
            AddReferenceSchemaField rows, FIELD_SOURCE_SCAN_VERSION, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_SOURCE_SCAN_COMPLETED_ON, dbDate, 0, False
            AddReferenceSchemaField rows, FIELD_QUALIFIED_REFERENCE_TOKENS, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_HEADER_REFERENCE_DEPENDENCIES, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_HEADER_DB_OBJECT_DEPENDENCIES, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_SCAN_STATUS, dbText, 40, False
            AddReferenceSchemaField rows, FIELD_PROJECT_IDENTITY, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_MODULE_IDENTITY, dbText, 255, False, True, False
            AddReferenceSchemaField rows, FIELD_DIAGNOSTIC_TEXT, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_INVENTORY_STATUS, dbText, 20, False
            AddReferenceSchemaField rows, FIELD_LAST_SCAN_BATCH_ID, dbText, 100, False
            AddReferenceSchemaField rows, FIELD_LAST_SCANNED_ON, dbDate, 0, False
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProcedures"))
            AddReferenceSchemaField rows, FIELD_ID, dbLong, 0, True
            AddReferenceSchemaField rows, FIELD_PROJECT_NAME, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_PROJECT_ROLE, dbText, 50, False
            AddReferenceSchemaField rows, FIELD_MODULE_NAME, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_PROCEDURE_ORDER, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_PROCEDURE_NAME, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_PROCEDURE_KIND, dbText, 50, False
            AddReferenceSchemaField rows, FIELD_PROCEDURE_SCOPE, dbText, 40, False
            AddReferenceSchemaField rows, FIELD_IS_PUBLIC, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_IS_EXTERNALLY_CALLABLE, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_HAS_OPTION_PRIVATE_MODULE, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_DECLARATION_TEXT, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_PROJECT_IDENTITY, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_MODULE_IDENTITY, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_PROCEDURE_IDENTITY, dbText, 255, False, True, False
            AddReferenceSchemaField rows, FIELD_PROCEDURE_NAME_KEY, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_INVENTORY_STATUS, dbText, 20, False
            AddReferenceSchemaField rows, FIELD_LAST_SCAN_BATCH_ID, dbText, 100, False
            AddReferenceSchemaField rows, FIELD_LAST_SCANNED_ON, dbDate, 0, False
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBADependencies"))
            AddReferenceSchemaField rows, FIELD_ID, dbLong, 0, True
            AddReferenceSchemaField rows, FIELD_PROJECT_NAME, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_PROJECT_ROLE, dbText, 50, False
            AddReferenceSchemaField rows, FIELD_SOURCE_MODULE_NAME, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_DEPENDENCY_TYPE, dbText, 80, False
            AddReferenceSchemaField rows, FIELD_DEPENDENCY_NAME, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_TARGET_PROJECT_NAME, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_TARGET_MODULE_NAME, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_MATCH_STATUS, dbText, 80, False
            AddReferenceSchemaField rows, FIELD_IS_AUTHORITATIVE_MATCH, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_EVIDENCE_TEXT, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_PROJECT_IDENTITY, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_SOURCE_MODULE_IDENTITY, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_TARGET_PROJECT_IDENTITY, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_TARGET_MODULE_IDENTITY, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_DEPENDENCY_IDENTITY, dbText, 255, False, True, False
            AddReferenceSchemaField rows, FIELD_DIAGNOSTIC_TEXT, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_INVENTORY_STATUS, dbText, 20, False
            AddReferenceSchemaField rows, FIELD_LAST_SCAN_BATCH_ID, dbText, 100, False
            AddReferenceSchemaField rows, FIELD_LAST_SCANNED_ON, dbDate, 0, False
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActiveReferenceProjectLinks"))
            AddReferenceSchemaField rows, FIELD_ID, dbLong, 0, True
            AddReferenceSchemaField rows, FIELD_REFERENCE_PRIORITY, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_REFERENCE_NAME, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_REFERENCE_KIND, dbText, 30, False
            AddReferenceSchemaField rows, FIELD_MATCH_STATUS, dbText, 50, False
            AddReferenceSchemaField rows, FIELD_MATCH_METHOD, dbText, 50, False
            AddReferenceSchemaField rows, FIELD_IS_AUTHORITATIVE_MATCH, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_LOADED_PROJECT_NAME, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_LOADED_PROJECT_ROLE, dbText, 50, False
            AddReferenceSchemaField rows, FIELD_REFERENCE_FULL_PATH, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_LOADED_PROJECT_FILE_PATH, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_REFERENCE_IDENTITY, dbText, 255, False, True, False
            AddReferenceSchemaField rows, FIELD_LOADED_PROJECT_IDENTITY, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_MATCH_EVIDENCE, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_DIAGNOSTIC_TEXT, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_INVENTORY_STATUS, dbText, 20, False
            AddReferenceSchemaField rows, FIELD_LAST_SCAN_BATCH_ID, dbText, 100, False
            AddReferenceSchemaField rows, FIELD_LAST_SCANNED_ON, dbDate, 0, False
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActiveReferences"))
            AddReferenceSchemaField rows, FIELD_ID, dbLong, 0, True
            AddReferenceSchemaField rows, FIELD_REFERENCE_PRIORITY, dbLong, 0, False, True, Null, ">0", "ReferencePriority must be positive."
            AddReferenceSchemaField rows, FIELD_REFERENCE_NAME, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_REFERENCE_DESCRIPTION, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_REFERENCE_GUID, dbText, 50, False
            AddReferenceSchemaField rows, FIELD_MAJOR_VERSION, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_MINOR_VERSION, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_REFERENCE_KIND, dbText, 30, False
            AddReferenceSchemaField rows, FIELD_REFERENCE_KIND_CODE, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_REFERENCE_IDENTITY, dbText, 255, False, True, False
            AddReferenceSchemaField rows, FIELD_IS_BUILT_IN, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_IS_BROKEN, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_SCAN_STATUS, dbText, 30, False
            AddReferenceSchemaField rows, FIELD_FULL_PATH, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_FILE_EXISTS, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_FILE_NAME, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_FOLDER_PATH, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_FILE_EXTENSION, dbText, 30, False
            AddReferenceSchemaField rows, FIELD_FILE_SIZE_BYTES, dbDouble, 0, False
            AddReferenceSchemaField rows, FIELD_FILE_CREATED_ON, dbDate, 0, False
            AddReferenceSchemaField rows, FIELD_FILE_MODIFIED_ON, dbDate, 0, False
            AddReferenceSchemaField rows, FIELD_FILE_VERSION, dbText, 100, False
            AddReferenceSchemaField rows, FIELD_FILE_ATTRIBUTES, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_IS_REGISTERED, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_REGISTERED_PATH, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_REGISTRATION_SCOPE, dbText, 30, False
            AddReferenceSchemaField rows, FIELD_REGISTRATION_VIEW, dbText, 20, False
            AddReferenceSchemaField rows, FIELD_PLATFORM, dbText, 20, False
            AddReferenceSchemaField rows, FIELD_REGISTRY_LCID, dbText, 20, False
            AddReferenceSchemaField rows, FIELD_REGISTRY_KEY_PATH, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_TYPE_LIB_FLAGS, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_HELP_DIRECTORY, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_DATABASE_NAME, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_DATABASE_PATH, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_DIAGNOSTIC_TEXT, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_INVENTORY_STATUS, dbText, 20, False
            AddReferenceSchemaField rows, FIELD_LAST_SCAN_BATCH_ID, dbText, 100, False
            AddReferenceSchemaField rows, FIELD_LAST_SCANNED_ON, dbDate, 0, False
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AvailableReferences"))
            AddReferenceSchemaField rows, FIELD_ID, dbLong, 0, True
            AddReferenceSchemaField rows, FIELD_REFERENCE_NAME, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_REFERENCE_DESCRIPTION, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_REFERENCE_GUID, dbText, 50, False
            AddReferenceSchemaField rows, FIELD_MAJOR_VERSION, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_MINOR_VERSION, dbLong, 0, False
            AddReferenceSchemaField rows, FIELD_VERSION_TEXT, dbText, 30, False
            AddReferenceSchemaField rows, FIELD_LCID, dbText, 20, False
            AddReferenceSchemaField rows, FIELD_LCID_DECIMAL, dbText, 20, False
            AddReferenceSchemaField rows, FIELD_PLATFORM, dbText, 20, False
            AddReferenceSchemaField rows, FIELD_REFERENCE_IDENTITY, dbText, 255, False, True, False
            AddReferenceSchemaField rows, FIELD_SCAN_STATUS, dbText, 30, False
            AddReferenceSchemaField rows, FIELD_REGISTERED_PATH, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_FILE_EXISTS, dbBoolean, 0, False
            AddReferenceSchemaField rows, FIELD_FILE_NAME, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_FOLDER_PATH, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_FILE_EXTENSION, dbText, 30, False
            AddReferenceSchemaField rows, FIELD_FILE_SIZE_BYTES, dbDouble, 0, False
            AddReferenceSchemaField rows, FIELD_FILE_CREATED_ON, dbDate, 0, False
            AddReferenceSchemaField rows, FIELD_FILE_MODIFIED_ON, dbDate, 0, False
            AddReferenceSchemaField rows, FIELD_FILE_VERSION, dbText, 100, False
            AddReferenceSchemaField rows, FIELD_FILE_ATTRIBUTES, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_REGISTRATION_SCOPE, dbText, 30, False
            AddReferenceSchemaField rows, FIELD_REGISTRATION_VIEW, dbText, 20, False
            AddReferenceSchemaField rows, FIELD_REGISTRY_ROOT, dbText, 40, False
            AddReferenceSchemaField rows, FIELD_REGISTRY_KEY_PATH, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_TYPE_LIB_FLAGS, dbText, 255, False
            AddReferenceSchemaField rows, FIELD_HELP_DIRECTORY, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_DIAGNOSTIC_TEXT, dbMemo, 0, False
            AddReferenceSchemaField rows, FIELD_INVENTORY_STATUS, dbText, 20, False
            AddReferenceSchemaField rows, FIELD_LAST_SCAN_BATCH_ID, dbText, 100, False
            AddReferenceSchemaField rows, FIELD_LAST_SCANNED_ON, dbDate, 0, False
        Case Else
            Err.Raise vbObjectError + 9791, "ReferenceTableSchemaDefinitions", "No reference-inventory schema is registered for table: " & tableName
    End Select
    Set ReferenceTableSchemaDefinitions = rows
End Function

Private Sub ApplyReferenceSchemaFieldContract(ByVal fld As DAO.Field, ByVal fieldDefinition As Variant, ByVal tableName As String)
    On Error GoTo ErrHandler
    If UBound(fieldDefinition) >= 4 Then If IsNull(fieldDefinition(4)) = False Then If fld.Required <> CBool(fieldDefinition(4)) Then fld.Required = CBool(fieldDefinition(4))
    If UBound(fieldDefinition) >= 5 Then If IsNull(fieldDefinition(5)) = False And (fld.Type = dbText Or fld.Type = dbMemo) Then If fld.allowZeroLength <> CBool(fieldDefinition(5)) Then fld.allowZeroLength = CBool(fieldDefinition(5))
    If UBound(fieldDefinition) >= 6 Then If IsNull(fieldDefinition(6)) = False Then If StrComp(CStr(Nz(fld.validationRule, vbNullString)), CStr(fieldDefinition(6)), vbBinaryCompare) <> 0 Then fld.validationRule = CStr(fieldDefinition(6))
    If UBound(fieldDefinition) >= 7 Then If IsNull(fieldDefinition(7)) = False Then If StrComp(CStr(Nz(fld.validationText, vbNullString)), CStr(fieldDefinition(7)), vbBinaryCompare) <> 0 Then fld.validationText = CStr(fieldDefinition(7))
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "ApplyReferenceSchemaFieldContract", "Could not enforce the canonical reference-inventory field contract." & vbCrLf & "Table: " & tableName & vbCrLf & "Field: " & fld.Name & vbCrLf & "Err.Description: " & Err.Description
End Sub

Private Sub AppendReferenceSchemaFields(ByVal tdf As DAO.tableDef, ByVal definitions As Collection)
    Dim fieldDefinition As Variant
    For Each fieldDefinition In definitions
        If CBool(fieldDefinition(3)) Then
            AppendAutoNumberField tdf, CStr(fieldDefinition(0))
        Else
            AppendField tdf, CStr(fieldDefinition(0)), CLng(fieldDefinition(1)), CLng(fieldDefinition(2))
        End If
        ApplyReferenceSchemaFieldContract tdf.fields(CStr(fieldDefinition(0))), fieldDefinition, tdf.Name
    Next fieldDefinition
End Sub

Private Sub EnsureReferenceSchemaFields(ByRef tdf As DAO.tableDef, ByVal definitions As Collection)
    Dim db As DAO.Database
    Dim fieldDefinition As Variant
    Dim fieldName As String
    Dim expectedType As DAO.DataTypeEnum
    Dim expectedSize As Long
    Dim tableName As String
    Set db = ADCDb()
    tableName = tdf.Name
    For Each fieldDefinition In definitions
        fieldName = CStr(fieldDefinition(0))
        expectedType = CLng(fieldDefinition(1))
        expectedSize = CLng(fieldDefinition(2))
        If StrComp(fieldName, FIELD_PROJECT_IDENTITY, vbTextCompare) = 0 And expectedType = dbText Then
            If DAOFieldExists(tdf, fieldName) Then
                If tdf.fields(fieldName).Type <> dbText Then
                    db.Execute "ALTER TABLE [" & Replace(tableName, "]", "]]") & "] ALTER COLUMN [" & Replace(fieldName, "]", "]]") & "] TEXT(" & CStr(expectedSize) & ");", dbFailOnError
                    db.TableDefs.Refresh
                    Set tdf = db.TableDefs(tableName)
                End If
            End If
        End If
        EnsureField tdf, fieldName, expectedType, expectedSize, CBool(fieldDefinition(3))
        ApplyReferenceSchemaFieldContract tdf.fields(fieldName), fieldDefinition, tableName
    Next fieldDefinition
End Sub

Private Sub EnsureReferenceQueries(ByVal canonicalQueries As Collection, ByRef repairStage As String)
    Dim activeSql As String
    Dim availableSql As String
    Dim loadedProjectsSql As String
    Dim projectLinksSql As String
    Dim managerSql As String
    Dim addInInventorySql As String
    Dim addInProjectReferencesSql As String
    Dim addInModulesSql As String
    Dim addInProceduresSql As String
    Dim addInProcedureConflictsSql As String
    Dim addInDependenciesSql As String
    Dim addInProjectProcedureCountsSql As String
    Dim addInProjectConflictCountsSql As String
    Dim addInReviewSummarySql As String
    Dim addInReferenceReportSql As String
    Dim sqlByQuery As Object
    Dim item As Variant
    Dim queryName As String
    ' Access/ACE can reject constant/filter predicates inside OUTER JOIN ON expressions (Err 3296). Reuse canonical current-partition saved queries on the right side, then join only on ordinary field-to-field
    ' identities so unmatched left rows remain eligible.
    activeSql = "SELECT A." & SQL_FIELD_ID & ",A." & SQL_FIELD_REFERENCE_PRIORITY & ",A." & SQL_FIELD_REFERENCE_NAME & ",A." & SQL_FIELD_REFERENCE_DESCRIPTION & ",A." & SQL_FIELD_REFERENCE_GUID & ",A." & SQL_FIELD_MAJOR_VERSION & ",A." & SQL_FIELD_MINOR_VERSION & ",A." & SQL_FIELD_REFERENCE_KIND & ",A." & SQL_FIELD_REFERENCE_KIND_CODE & ",A." & SQL_FIELD_REFERENCE_IDENTITY & ",A." & SQL_FIELD_IS_BUILT_IN & ",A." & SQL_FIELD_IS_BROKEN & ",A." & SQL_FIELD_SCAN_STATUS & "," & _
                "L." & SQL_FIELD_MATCH_STATUS & " AS " & SQL_FIELD_LOADED_PROJECT_MATCH_STATUS & ",L." & SQL_FIELD_MATCH_METHOD & " AS " & SQL_FIELD_LOADED_PROJECT_MATCH_METHOD & ",L." & SQL_FIELD_IS_AUTHORITATIVE_MATCH & ",L." & SQL_FIELD_LOADED_PROJECT_NAME & ",L." & SQL_FIELD_LOADED_PROJECT_ROLE & ",L." & SQL_FIELD_LOADED_PROJECT_FILE_PATH & ",A." & SQL_FIELD_FULL_PATH & ",A." & SQL_FIELD_FILE_EXISTS & ",A." & SQL_FIELD_FILE_NAME & ",A." & SQL_FIELD_FOLDER_PATH & ",A." & SQL_FIELD_FILE_EXTENSION & "," & _
                "A." & SQL_FIELD_FILE_SIZE_BYTES & ",A." & SQL_FIELD_FILE_CREATED_ON & ",A." & SQL_FIELD_FILE_MODIFIED_ON & ",A." & SQL_FIELD_FILE_VERSION & ",A." & SQL_FIELD_FILE_ATTRIBUTES & ",A." & SQL_FIELD_IS_REGISTERED & ",A." & SQL_FIELD_REGISTERED_PATH & ",A." & SQL_FIELD_REGISTRATION_SCOPE & ",A." & SQL_FIELD_REGISTRATION_VIEW & ",A." & SQL_FIELD_PLATFORM & ",A." & SQL_FIELD_REGISTRY_LCID & ",A." & SQL_FIELD_REGISTRY_KEY_PATH & ",A." & SQL_FIELD_TYPE_LIB_FLAGS & ",A." & SQL_FIELD_HELP_DIRECTORY & ",A." & SQL_FIELD_DATABASE_NAME & ",A." & SQL_FIELD_DATABASE_PATH & "," & _
                "L." & SQL_FIELD_MATCH_EVIDENCE & " AS " & SQL_FIELD_LOADED_PROJECT_MATCH_EVIDENCE & ",L." & SQL_FIELD_DIAGNOSTIC_TEXT & " AS " & SQL_FIELD_LOADED_PROJECT_MATCH_DIAGNOSTIC & ",A." & SQL_FIELD_DIAGNOSTIC_TEXT & ",A." & SQL_FIELD_LAST_SCANNED_ON & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActiveReferences") & "] AS A LEFT JOIN [" & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_References_Project_Links") & "] AS L ON A." & SQL_FIELD_REFERENCE_IDENTITY & "=L." & SQL_FIELD_REFERENCE_IDENTITY & " WHERE A." & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " ORDER BY " & OrderByPolicyText("REFERENCE_ACTIVE") & ";"
    availableSql = "SELECT " & SQL_FIELD_ID & "," & SQL_FIELD_REFERENCE_NAME & "," & SQL_FIELD_REFERENCE_DESCRIPTION & "," & SQL_FIELD_REFERENCE_GUID & "," & SQL_FIELD_MAJOR_VERSION & "," & SQL_FIELD_MINOR_VERSION & "," & SQL_FIELD_VERSION_TEXT & "," & SQL_FIELD_LCID & "," & SQL_FIELD_LCID_DECIMAL & "," & SQL_FIELD_PLATFORM & "," & SQL_FIELD_REFERENCE_IDENTITY & "," & SQL_FIELD_SCAN_STATUS & "," & SQL_FIELD_REGISTERED_PATH & "," & SQL_FIELD_FILE_EXISTS & "," & SQL_FIELD_FILE_NAME & "," & SQL_FIELD_FOLDER_PATH & "," & SQL_FIELD_FILE_EXTENSION & "," & SQL_FIELD_FILE_SIZE_BYTES & "," & SQL_FIELD_FILE_CREATED_ON & "," & SQL_FIELD_FILE_MODIFIED_ON & "," & SQL_FIELD_FILE_VERSION & "," & SQL_FIELD_FILE_ATTRIBUTES & "," & SQL_FIELD_REGISTRATION_SCOPE & "," & SQL_FIELD_REGISTRATION_VIEW & "," & SQL_FIELD_REGISTRY_ROOT & "," & SQL_FIELD_REGISTRY_KEY_PATH & "," & SQL_FIELD_TYPE_LIB_FLAGS & "," & _
                SQL_FIELD_HELP_DIRECTORY & "," & SQL_FIELD_DIAGNOSTIC_TEXT & "," & SQL_FIELD_LAST_SCANNED_ON & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AvailableReferences") & "] WHERE " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " ORDER BY " & OrderByPolicyText("REFERENCE_AVAILABLE") & ";"
    loadedProjectsSql = "SELECT " & SQL_FIELD_ID & "," & SQL_FIELD_PROJECT_ORDER & "," & SQL_FIELD_PROJECT_NAME & "," & SQL_FIELD_PROJECT_ROLE & "," & SQL_FIELD_IS_ACTIVE_VBE_PROJECT & "," & SQL_FIELD_COMPONENT_COUNT & "," & SQL_FIELD_COMPONENTS_READABLE & "," & SQL_FIELD_SOURCE_READABLE & "," & SQL_FIELD_PROTECTION_STATE & "," & SQL_FIELD_PROTECTION_CODE & "," & SQL_FIELD_PROJECT_TYPE & "," & SQL_FIELD_PROJECT_TYPE_CODE & "," & SQL_FIELD_SCAN_STATUS & "," & SQL_FIELD_SOURCE_ANALYSIS_STATUS & "," & SQL_FIELD_MODULE_ROSTER_SIGNATURE & "," & SQL_FIELD_PROJECT_REFERENCE_COUNT & "," & SQL_FIELD_MODULE_INVENTORY_COUNT & "," & SQL_FIELD_PROCEDURE_INVENTORY_COUNT & "," & SQL_FIELD_DEPENDENCY_COUNT & "," & SQL_FIELD_LAST_SOURCE_SCANNED_ON & "," & SQL_FIELD_SOURCE_ANALYSIS_DIAGNOSTIC_TEXT & "," & SQL_FIELD_PROJECT_FILE_PATH & "," & SQL_FIELD_FILE_EXISTS & "," & SQL_FIELD_PROJECT_FILE_NAME & "," & _
                SQL_FIELD_PROJECT_FILE_BASE_NAME & "," & SQL_FIELD_PROJECT_IDENTITY & "," & SQL_FIELD_DIAGNOSTIC_TEXT & "," & SQL_FIELD_LAST_SCANNED_ON & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects") & "] WHERE " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " ORDER BY " & OrderByPolicyText("REFERENCE_LOADED_PROJECTS") & ";"
    projectLinksSql = "SELECT " & SQL_FIELD_ID & "," & SQL_FIELD_REFERENCE_PRIORITY & "," & SQL_FIELD_REFERENCE_NAME & "," & SQL_FIELD_REFERENCE_KIND & "," & SQL_FIELD_MATCH_STATUS & "," & SQL_FIELD_MATCH_METHOD & "," & SQL_FIELD_IS_AUTHORITATIVE_MATCH & "," & SQL_FIELD_LOADED_PROJECT_NAME & "," & SQL_FIELD_LOADED_PROJECT_ROLE & "," & SQL_FIELD_REFERENCE_FULL_PATH & "," & SQL_FIELD_LOADED_PROJECT_FILE_PATH & "," & SQL_FIELD_REFERENCE_IDENTITY & "," & SQL_FIELD_LOADED_PROJECT_IDENTITY & "," & SQL_FIELD_MATCH_EVIDENCE & "," & SQL_FIELD_DIAGNOSTIC_TEXT & "," & SQL_FIELD_LAST_SCANNED_ON & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActiveReferenceProjectLinks") & "] WHERE " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " ORDER BY " & OrderByPolicyText("REFERENCE_PROJECT_LINKS") & ";"
    managerSql = "SELECT " & SQL_FIELD_ID & "," & SQL_FIELD_REQUESTED_ACTIVE & "," & SQL_FIELD_ORIGINAL_ACTIVE & "," & SQL_FIELD_REFERENCE_SOURCE & "," & SQL_FIELD_REFERENCE_PRIORITY & "," & SQL_FIELD_REFERENCE_NAME & "," & SQL_FIELD_REFERENCE_DESCRIPTION & "," & SQL_FIELD_REFERENCE_GUID & "," & SQL_FIELD_MAJOR_VERSION & "," & SQL_FIELD_MINOR_VERSION & "," & SQL_FIELD_REFERENCE_IDENTITY & "," & SQL_FIELD_CAN_CHANGE & "," & SQL_FIELD_PROTECTION_REASON & "," & SQL_FIELD_HEALTH_STATUS & "," & SQL_FIELD_IS_BUILT_IN & "," & SQL_FIELD_IS_BROKEN & "," & SQL_FIELD_FILE_EXISTS & "," & SQL_FIELD_IS_REGISTERED & "," & SQL_FIELD_FULL_PATH & "," & SQL_FIELD_REGISTERED_PATH & "," & SQL_FIELD_REGISTRATION_SCOPE & "," & SQL_FIELD_REGISTRATION_VIEW & "," & SQL_FIELD_PLATFORM & "," & SQL_FIELD_LAST_REFRESHED_ON & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ReferenceManager") & _
                "] ORDER BY " & OrderByPolicyText("REFERENCE_MANAGER") & ";"
    ' ACE derives compound-UNION output typing from the leading SELECT. This zero-row type anchor uses canonical physical fields for every output column so later branches preserve Null for not-applicable values
    ' without ambiguous Variant typing.
    addInInventorySql = "SELECT V." & SQL_FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME & " AS " & SQL_FIELD_SOURCE_KIND & ",V." & SQL_FIELD_ADD_IN_ORDER & " AS " & SQL_FIELD_INVENTORY_ORDER & ",V." & SQL_FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME & ",V." & SQL_FIELD_DESCRIPTION & ",V." & SQL_FIELD_PROG_ID & ",V." & SQL_FIELD_GUID & ",A." & SQL_FIELD_ADD_IN_LIBRARY & ",A." & SQL_FIELD_ENTRY_POINT_EXPRESSION & ",A." & SQL_FIELD_ENTRY_POINT_PROCEDURE & ",A." & SQL_FIELD_ENTRY_POINT_EXISTS & ",A." & SQL_FIELD_ENTRY_POINT_EXTERNALLY_CALLABLE & ",V." & SQL_FIELD_IS_REGISTERED & ",V." & SQL_FIELD_IS_CONNECTED & ",V." & SQL_FIELD_IS_CONNECTED & " AS " & SQL_FIELD_IS_LOADED & ",V." & SQL_FIELD_OBJECT_AVAILABLE & ",A." & SQL_FIELD_PROJECT_ROLE & ",A." & SQL_FIELD_PROJECT_FILE_PATH & " AS " & SQL_FIELD_FILE_PATH & ",A." & SQL_FIELD_FILE_EXISTS & ",P." & SQL_FIELD_COMPONENT_COUNT & ",P." & _
                SQL_FIELD_SOURCE_READABLE & ",P." & SQL_FIELD_PROTECTION_STATE & ",V." & SQL_FIELD_SCAN_STATUS & ",V." & SQL_FIELD_DIAGNOSTIC_TEXT & ",V." & SQL_FIELD_LAST_SCANNED_ON & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_VBEAddIns") & "] AS V, [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AccessAddIns") & "] AS A, [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects") & "] AS P WHERE False UNION ALL " & _
                        "SELECT 'VBE Add-in' AS " & SQL_FIELD_SOURCE_KIND & "," & SQL_FIELD_ADD_IN_ORDER & " AS " & SQL_FIELD_INVENTORY_ORDER & "," & SQL_FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME & "," & SQL_FIELD_DESCRIPTION & "," & SQL_FIELD_PROG_ID & "," & SQL_FIELD_GUID & ",Null AS " & SQL_FIELD_ADD_IN_LIBRARY & ",Null AS " & SQL_FIELD_ENTRY_POINT_EXPRESSION & ",Null AS " & SQL_FIELD_ENTRY_POINT_PROCEDURE & ",Null AS " & SQL_FIELD_ENTRY_POINT_EXISTS & ",Null AS " & SQL_FIELD_ENTRY_POINT_EXTERNALLY_CALLABLE & "," & SQL_FIELD_IS_REGISTERED & "," & SQL_FIELD_IS_CONNECTED & "," & SQL_FIELD_IS_CONNECTED & " AS " & SQL_FIELD_IS_LOADED & "," & SQL_FIELD_OBJECT_AVAILABLE & ",Null AS " & SQL_FIELD_PROJECT_ROLE & ",Null AS " & SQL_FIELD_FILE_PATH & ",False AS " & SQL_FIELD_FILE_EXISTS & ",0 AS " & SQL_FIELD_COMPONENT_COUNT & ",False AS " & SQL_FIELD_SOURCE_READABLE & ",Null AS " & _
                                    SQL_FIELD_PROTECTION_STATE & "," & SQL_FIELD_SCAN_STATUS & "," & SQL_FIELD_DIAGNOSTIC_TEXT & "," & SQL_FIELD_LAST_SCANNED_ON & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_VBEAddIns") & "] WHERE " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " UNION ALL " & _
                        "SELECT 'Access Add-in' AS " & SQL_FIELD_SOURCE_KIND & "," & SQL_FIELD_ADD_IN_ORDER & " AS " & SQL_FIELD_INVENTORY_ORDER & "," & SQL_FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME & ",Null AS " & SQL_FIELD_DESCRIPTION & ",Null AS " & SQL_FIELD_PROG_ID & ",Null AS " & SQL_FIELD_GUID & "," & SQL_FIELD_ADD_IN_LIBRARY & "," & SQL_FIELD_ENTRY_POINT_EXPRESSION & "," & SQL_FIELD_ENTRY_POINT_PROCEDURE & "," & SQL_FIELD_ENTRY_POINT_EXISTS & "," & SQL_FIELD_ENTRY_POINT_EXTERNALLY_CALLABLE & ",Null AS " & SQL_FIELD_IS_REGISTERED & ",Null AS " & SQL_FIELD_IS_CONNECTED & ",IIf(Len(Nz(" & SQL_FIELD_PROJECT_IDENTITY & ",''))>0,True,False) AS " & SQL_FIELD_IS_LOADED & ",Null AS " & SQL_FIELD_OBJECT_AVAILABLE & "," & SQL_FIELD_PROJECT_ROLE & "," & SQL_FIELD_PROJECT_FILE_PATH & " AS " & SQL_FIELD_FILE_PATH & "," & SQL_FIELD_FILE_EXISTS & ",0 AS " & SQL_FIELD_COMPONENT_COUNT & _
                                    ",False AS " & SQL_FIELD_SOURCE_READABLE & ",Null AS " & SQL_FIELD_PROTECTION_STATE & "," & SQL_FIELD_SCAN_STATUS & "," & SQL_FIELD_DIAGNOSTIC_TEXT & "," & SQL_FIELD_LAST_SCANNED_ON & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AccessAddIns") & "] WHERE " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " UNION ALL " & _
                        "SELECT 'COM Add-in' AS " & SQL_FIELD_SOURCE_KIND & "," & SQL_FIELD_ADD_IN_ORDER & " AS " & SQL_FIELD_INVENTORY_ORDER & "," & SQL_FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME & "," & SQL_FIELD_DESCRIPTION & "," & SQL_FIELD_PROG_ID & "," & SQL_FIELD_GUID & ",Null AS " & SQL_FIELD_ADD_IN_LIBRARY & ",Null AS " & SQL_FIELD_ENTRY_POINT_EXPRESSION & ",Null AS " & SQL_FIELD_ENTRY_POINT_PROCEDURE & ",Null AS " & SQL_FIELD_ENTRY_POINT_EXISTS & ",Null AS " & SQL_FIELD_ENTRY_POINT_EXTERNALLY_CALLABLE & "," & SQL_FIELD_IS_REGISTERED & "," & SQL_FIELD_IS_CONNECTED & "," & SQL_FIELD_IS_CONNECTED & " AS " & SQL_FIELD_IS_LOADED & "," & SQL_FIELD_OBJECT_AVAILABLE & ",Null AS " & SQL_FIELD_PROJECT_ROLE & ",Null AS " & SQL_FIELD_FILE_PATH & ",False AS " & SQL_FIELD_FILE_EXISTS & ",0 AS " & SQL_FIELD_COMPONENT_COUNT & ",False AS " & SQL_FIELD_SOURCE_READABLE & ",Null AS " & _
                                    SQL_FIELD_PROTECTION_STATE & "," & SQL_FIELD_SCAN_STATUS & "," & SQL_FIELD_DIAGNOSTIC_TEXT & "," & SQL_FIELD_LAST_SCANNED_ON & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_COMAddIns") & "] WHERE " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " UNION ALL " & _
                        "SELECT 'Loaded VBA Project' AS " & SQL_FIELD_SOURCE_KIND & "," & SQL_FIELD_PROJECT_ORDER & " AS " & SQL_FIELD_INVENTORY_ORDER & "," & SQL_FIELD_PROJECT_NAME & " AS " & SQL_FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME & ",Null AS " & SQL_FIELD_DESCRIPTION & ",Null AS " & SQL_FIELD_PROG_ID & ",Null AS " & SQL_FIELD_GUID & ",Null AS " & SQL_FIELD_ADD_IN_LIBRARY & ",Null AS " & SQL_FIELD_ENTRY_POINT_EXPRESSION & ",Null AS " & SQL_FIELD_ENTRY_POINT_PROCEDURE & ",Null AS " & SQL_FIELD_ENTRY_POINT_EXISTS & ",Null AS " & SQL_FIELD_ENTRY_POINT_EXTERNALLY_CALLABLE & ",Null AS " & SQL_FIELD_IS_REGISTERED & ",Null AS " & SQL_FIELD_IS_CONNECTED & ",True AS " & SQL_FIELD_IS_LOADED & "," & SQL_FIELD_COMPONENTS_READABLE & " AS " & SQL_FIELD_OBJECT_AVAILABLE & "," & SQL_FIELD_PROJECT_ROLE & "," & SQL_FIELD_PROJECT_FILE_PATH & " AS " & SQL_FIELD_FILE_PATH & "," & SQL_FIELD_FILE_EXISTS & _
                                    "," & SQL_FIELD_COMPONENT_COUNT & "," & SQL_FIELD_SOURCE_READABLE & "," & SQL_FIELD_PROTECTION_STATE & "," & SQL_FIELD_SCAN_STATUS & "," & SQL_FIELD_DIAGNOSTIC_TEXT & "," & SQL_FIELD_LAST_SCANNED_ON & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects") & "] WHERE " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " ORDER BY " & OrderByPolicyText("REFERENCE_ADDIN_INVENTORY") & ";"
    addInProjectReferencesSql = "SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjectReferences") & "] WHERE " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " AND " & SQL_FIELD_PROJECT_ROLE & "='External' ORDER BY " & OrderByPolicyText("REFERENCE_PROJECT_REFERENCES") & ";"
    addInModulesSql = "SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAModules") & "] WHERE " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " AND " & SQL_FIELD_PROJECT_ROLE & "='External' ORDER BY " & OrderByPolicyText("REFERENCE_MODULES") & ";"
    addInProceduresSql = "SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProcedures") & "] WHERE " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " AND " & SQL_FIELD_PROJECT_ROLE & "='External' AND " & SQL_FIELD_IS_PUBLIC & "=True ORDER BY " & OrderByPolicyText("REFERENCE_PROCEDURES") & ";"
    addInProcedureConflictsSql = "SELECT P.*,'Conflict' AS " & SQL_FIELD_PROCEDURE_CONFLICT_STATUS & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProcedures") & "] AS P WHERE P." & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " AND P." & SQL_FIELD_PROJECT_ROLE & "='External' AND P." & SQL_FIELD_IS_EXTERNALLY_CALLABLE & "=True AND P." & SQL_FIELD_PROCEDURE_NAME_KEY & " IN (SELECT " & SQL_FIELD_PROCEDURE_NAME_KEY & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProcedures") & "] WHERE " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " AND " & SQL_FIELD_PROJECT_ROLE & "='External' AND " & SQL_FIELD_IS_EXTERNALLY_CALLABLE & "=True GROUP BY " & SQL_FIELD_PROCEDURE_NAME_KEY & " HAVING Count(*)>1) ORDER BY " & _
                OrderByPolicyText("REFERENCE_PROCEDURE_CONFLICTS") & ";"
    addInDependenciesSql = "SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBADependencies") & "] WHERE " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " AND " & SQL_FIELD_PROJECT_ROLE & "='External' ORDER BY " & OrderByPolicyText("REFERENCE_DEPENDENCIES") & ";"
    addInProjectProcedureCountsSql = "SELECT CStr(Nz(P." & SQL_FIELD_PROJECT_IDENTITY & ",'')) AS " & SQL_FIELD_PROJECT_IDENTITY & ",Sum(IIf(P." & SQL_FIELD_IS_PUBLIC & "=True,1,0)) AS " & SQL_FIELD_PUBLIC_PROCEDURE_COUNT & " FROM [" & _
                                      ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProcedures") & "] AS P WHERE P." & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " AND Len(Nz(P." & SQL_FIELD_PROJECT_IDENTITY & ",''))>0 GROUP BY CStr(Nz(P." & SQL_FIELD_PROJECT_IDENTITY & ",''));"
    addInProjectConflictCountsSql = "SELECT CStr(Nz(C." & SQL_FIELD_PROJECT_IDENTITY & ",'')) AS " & SQL_FIELD_PROJECT_IDENTITY & ",Count(*) AS " & SQL_FIELD_PROCEDURE_CONFLICT_COUNT & " FROM [" & _
                                     ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_ExternalProject_Procedure_Conflicts") & "] AS C WHERE Len(Nz(C." & SQL_FIELD_PROJECT_IDENTITY & ",''))>0 GROUP BY CStr(Nz(C." & SQL_FIELD_PROJECT_IDENTITY & ",''));"
    ' ACE derives compound-UNION output typing from the leading SELECT. This zero-row type anchor uses canonical physical fields for all 30 output ordinals so not-applicable Nulls in the four data branches cannot
    ' establish incompatible Variant types.
    addInReviewSummarySql = "SELECT V." & SQL_FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME & " AS " & SQL_FIELD_SOURCE_KIND & ",V." & SQL_FIELD_ADD_IN_IDENTITY & " AS " & SQL_FIELD_REVIEW_IDENTITY & ",A." & SQL_FIELD_PROJECT_IDENTITY & ",A." & SQL_FIELD_PROJECT_NAME & ",V." & SQL_FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME & " AS " & SQL_FIELD_ADD_IN_NAME & ",V." & SQL_FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME & ",V." & SQL_FIELD_DESCRIPTION & ",V." & SQL_FIELD_PROG_ID & ",V." & SQL_FIELD_GUID & ",A." & SQL_FIELD_ADD_IN_LIBRARY & ",A." & SQL_FIELD_ENTRY_POINT_EXPRESSION & ",A." & SQL_FIELD_ENTRY_POINT_PROCEDURE & "," & _
                            "A." & SQL_FIELD_ENTRY_POINT_EXISTS & ",A." & SQL_FIELD_ENTRY_POINT_EXTERNALLY_CALLABLE & ",V." & SQL_FIELD_IS_REGISTERED & ",V." & SQL_FIELD_IS_CONNECTED & ",V." & SQL_FIELD_IS_CONNECTED & " AS " & SQL_FIELD_IS_LOADED & ",V." & SQL_FIELD_OBJECT_AVAILABLE & ",V." & SQL_FIELD_SCAN_STATUS & " AS " & SQL_FIELD_REVIEW_STATUS & ",A." & SQL_FIELD_PROJECT_ROLE & ",A." & SQL_FIELD_PROJECT_FILE_PATH & " AS " & SQL_FIELD_FILE_PATH & ",A." & SQL_FIELD_FILE_EXISTS & ",P." & SQL_FIELD_PROJECT_REFERENCE_COUNT & "," & _
                            "P." & SQL_FIELD_MODULE_INVENTORY_COUNT & " AS " & SQL_FIELD_MODULE_COUNT & ",P." & SQL_FIELD_PROCEDURE_INVENTORY_COUNT & " AS " & SQL_FIELD_PUBLIC_PROCEDURE_COUNT & ",P." & SQL_FIELD_PROCEDURE_INVENTORY_COUNT & " AS " & SQL_FIELD_PROCEDURE_CONFLICT_COUNT & ",P." & SQL_FIELD_DEPENDENCY_COUNT & ",P." & SQL_FIELD_SOURCE_ANALYSIS_STATUS & ",V." & SQL_FIELD_DIAGNOSTIC_TEXT & ",V." & SQL_FIELD_LAST_SCANNED_ON & " FROM [" & _
                            ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_VBEAddIns") & "] AS V, [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AccessAddIns") & "] AS A, [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects") & "] AS P WHERE False UNION ALL " & _
                            "SELECT 'VBE Add-in' AS " & SQL_FIELD_SOURCE_KIND & ",'VBE|' & " & SQL_FIELD_ADD_IN_IDENTITY & " AS " & SQL_FIELD_REVIEW_IDENTITY & ",Null AS " & SQL_FIELD_PROJECT_IDENTITY & ",Null AS " & SQL_FIELD_PROJECT_NAME & "," & SQL_FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME & " AS " & SQL_FIELD_ADD_IN_NAME & "," & SQL_FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME & "," & SQL_FIELD_DESCRIPTION & "," & SQL_FIELD_PROG_ID & "," & SQL_FIELD_GUID & ",Null AS " & SQL_FIELD_ADD_IN_LIBRARY & ",Null AS " & SQL_FIELD_ENTRY_POINT_EXPRESSION & ",Null AS " & SQL_FIELD_ENTRY_POINT_PROCEDURE & ",Null AS " & SQL_FIELD_ENTRY_POINT_EXISTS & ",Null AS " & SQL_FIELD_ENTRY_POINT_EXTERNALLY_CALLABLE & "," & SQL_FIELD_IS_REGISTERED & "," & SQL_FIELD_IS_CONNECTED & "," & SQL_FIELD_IS_CONNECTED & " AS " & SQL_FIELD_IS_LOADED & "," & SQL_FIELD_OBJECT_AVAILABLE & "," & SQL_FIELD_SCAN_STATUS & " AS " & _
                                        SQL_FIELD_REVIEW_STATUS & ",Null AS " & SQL_FIELD_PROJECT_ROLE & ",Null AS " & SQL_FIELD_FILE_PATH & ",False AS " & SQL_FIELD_FILE_EXISTS & ",0 AS " & SQL_FIELD_PROJECT_REFERENCE_COUNT & ",0 AS " & SQL_FIELD_MODULE_COUNT & ",0 AS " & SQL_FIELD_PUBLIC_PROCEDURE_COUNT & ",0 AS " & SQL_FIELD_PROCEDURE_CONFLICT_COUNT & ",0 AS " & SQL_FIELD_DEPENDENCY_COUNT & ",Null AS " & SQL_FIELD_SOURCE_ANALYSIS_STATUS & "," & SQL_FIELD_DIAGNOSTIC_TEXT & "," & SQL_FIELD_LAST_SCANNED_ON & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_VBEAddIns") & "] WHERE " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " UNION ALL " & _
                            "SELECT 'Access Add-in' AS " & SQL_FIELD_SOURCE_KIND & ",'ACCESS|' & A." & SQL_FIELD_ADD_IN_IDENTITY & " AS " & SQL_FIELD_REVIEW_IDENTITY & ",A." & SQL_FIELD_PROJECT_IDENTITY & ",A." & SQL_FIELD_PROJECT_NAME & ",A." & SQL_FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME & " AS " & SQL_FIELD_ADD_IN_NAME & ",A." & SQL_FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME & ",Null AS " & SQL_FIELD_DESCRIPTION & ",Null AS " & SQL_FIELD_PROG_ID & ",Null AS " & SQL_FIELD_GUID & ",A." & SQL_FIELD_ADD_IN_LIBRARY & ",A." & SQL_FIELD_ENTRY_POINT_EXPRESSION & ",A." & SQL_FIELD_ENTRY_POINT_PROCEDURE & ",A." & SQL_FIELD_ENTRY_POINT_EXISTS & ",A." & SQL_FIELD_ENTRY_POINT_EXTERNALLY_CALLABLE & ",Null AS " & SQL_FIELD_IS_REGISTERED & ",Null AS " & SQL_FIELD_IS_CONNECTED & ",IIf(Len(Nz(A." & SQL_FIELD_PROJECT_IDENTITY & ",''))>0,True,False) AS " & SQL_FIELD_IS_LOADED & ",Null AS " & _
                                        SQL_FIELD_OBJECT_AVAILABLE & ",A." & SQL_FIELD_SCAN_STATUS & " AS " & SQL_FIELD_REVIEW_STATUS & ",A." & SQL_FIELD_PROJECT_ROLE & ",A." & SQL_FIELD_PROJECT_FILE_PATH & " AS " & SQL_FIELD_FILE_PATH & ",A." & SQL_FIELD_FILE_EXISTS & "," & _
                            "Nz(P." & SQL_FIELD_PROJECT_REFERENCE_COUNT & ",0) AS " & SQL_FIELD_PROJECT_REFERENCE_COUNT & ",Nz(P." & SQL_FIELD_MODULE_INVENTORY_COUNT & ",0) AS " & SQL_FIELD_MODULE_COUNT & ",Nz(PC." & SQL_FIELD_PUBLIC_PROCEDURE_COUNT & ",0) AS " & SQL_FIELD_PUBLIC_PROCEDURE_COUNT & ",Nz(CC." & SQL_FIELD_PROCEDURE_CONFLICT_COUNT & ",0) AS " & SQL_FIELD_PROCEDURE_CONFLICT_COUNT & ",Nz(P." & SQL_FIELD_DEPENDENCY_COUNT & ",0) AS " & SQL_FIELD_DEPENDENCY_COUNT & ",P." & SQL_FIELD_SOURCE_ANALYSIS_STATUS & ",A." & SQL_FIELD_DIAGNOSTIC_TEXT & ",IIf(P." & SQL_FIELD_LAST_SOURCE_SCANNED_ON & " Is Null,A." & SQL_FIELD_LAST_SCANNED_ON & ",P." & SQL_FIELD_LAST_SOURCE_SCANNED_ON & ") AS " & SQL_FIELD_LAST_SCANNED_ON & " FROM (([" & _
                            ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AccessAddIns") & "] AS A LEFT JOIN [" & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_References_Loaded_Projects") & "] AS P ON A." & SQL_FIELD_PROJECT_IDENTITY & "=P." & SQL_FIELD_PROJECT_IDENTITY & ") LEFT JOIN [" & _
                            ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Project_Procedure_Counts") & "] AS PC ON A." & SQL_FIELD_PROJECT_IDENTITY & "=PC." & SQL_FIELD_PROJECT_IDENTITY & ") LEFT JOIN [" & _
                            ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Project_Conflict_Counts") & "] AS CC ON A." & SQL_FIELD_PROJECT_IDENTITY & "=CC." & SQL_FIELD_PROJECT_IDENTITY & " WHERE A." & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " UNION ALL " & _
                            "SELECT 'COM Add-in' AS " & SQL_FIELD_SOURCE_KIND & ",'COM|' & " & SQL_FIELD_ADD_IN_IDENTITY & " AS " & SQL_FIELD_REVIEW_IDENTITY & ",Null AS " & SQL_FIELD_PROJECT_IDENTITY & ",Null AS " & SQL_FIELD_PROJECT_NAME & "," & SQL_FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME & " AS " & SQL_FIELD_ADD_IN_NAME & "," & SQL_FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME & "," & SQL_FIELD_DESCRIPTION & "," & SQL_FIELD_PROG_ID & "," & SQL_FIELD_GUID & ",Null AS " & SQL_FIELD_ADD_IN_LIBRARY & ",Null AS " & SQL_FIELD_ENTRY_POINT_EXPRESSION & ",Null AS " & SQL_FIELD_ENTRY_POINT_PROCEDURE & ",Null AS " & SQL_FIELD_ENTRY_POINT_EXISTS & ",Null AS " & SQL_FIELD_ENTRY_POINT_EXTERNALLY_CALLABLE & "," & SQL_FIELD_IS_REGISTERED & "," & SQL_FIELD_IS_CONNECTED & "," & SQL_FIELD_IS_CONNECTED & " AS " & SQL_FIELD_IS_LOADED & "," & SQL_FIELD_OBJECT_AVAILABLE & "," & SQL_FIELD_SCAN_STATUS & " AS " & _
                                        SQL_FIELD_REVIEW_STATUS & ",Null AS " & SQL_FIELD_PROJECT_ROLE & ",Null AS " & SQL_FIELD_FILE_PATH & ",False AS " & SQL_FIELD_FILE_EXISTS & ",0 AS " & SQL_FIELD_PROJECT_REFERENCE_COUNT & ",0 AS " & SQL_FIELD_MODULE_COUNT & ",0 AS " & SQL_FIELD_PUBLIC_PROCEDURE_COUNT & ",0 AS " & SQL_FIELD_PROCEDURE_CONFLICT_COUNT & ",0 AS " & SQL_FIELD_DEPENDENCY_COUNT & ",Null AS " & SQL_FIELD_SOURCE_ANALYSIS_STATUS & "," & SQL_FIELD_DIAGNOSTIC_TEXT & "," & SQL_FIELD_LAST_SCANNED_ON & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_COMAddIns") & "] WHERE " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " UNION ALL " & _
                            "SELECT 'Loaded VBA Project' AS " & SQL_FIELD_SOURCE_KIND & ",'PROJECT|' & P." & SQL_FIELD_PROJECT_IDENTITY & " AS " & SQL_FIELD_REVIEW_IDENTITY & ",P." & SQL_FIELD_PROJECT_IDENTITY & ",P." & SQL_FIELD_PROJECT_NAME & ",Null AS " & SQL_FIELD_ADD_IN_NAME & ",P." & SQL_FIELD_PROJECT_NAME & " AS " & SQL_FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME & ",Null AS " & SQL_FIELD_DESCRIPTION & ",Null AS " & SQL_FIELD_PROG_ID & ",Null AS " & SQL_FIELD_GUID & ",Null AS " & SQL_FIELD_ADD_IN_LIBRARY & ",Null AS " & SQL_FIELD_ENTRY_POINT_EXPRESSION & ",Null AS " & SQL_FIELD_ENTRY_POINT_PROCEDURE & ",Null AS " & SQL_FIELD_ENTRY_POINT_EXISTS & ",Null AS " & SQL_FIELD_ENTRY_POINT_EXTERNALLY_CALLABLE & ",Null AS " & SQL_FIELD_IS_REGISTERED & ",Null AS " & SQL_FIELD_IS_CONNECTED & ",True AS " & SQL_FIELD_IS_LOADED & ",P." & SQL_FIELD_COMPONENTS_READABLE & " AS " & _
                                        SQL_FIELD_OBJECT_AVAILABLE & "," & _
                            "IIf(Len(Nz(P." & SQL_FIELD_SOURCE_ANALYSIS_STATUS & ",''))=0,'Loaded','Loaded / ' & P." & SQL_FIELD_SOURCE_ANALYSIS_STATUS & ") AS " & SQL_FIELD_REVIEW_STATUS & ",P." & SQL_FIELD_PROJECT_ROLE & ",P." & SQL_FIELD_PROJECT_FILE_PATH & " AS " & SQL_FIELD_FILE_PATH & ",P." & SQL_FIELD_FILE_EXISTS & ",P." & SQL_FIELD_PROJECT_REFERENCE_COUNT & ",P." & SQL_FIELD_MODULE_INVENTORY_COUNT & " AS " & SQL_FIELD_MODULE_COUNT & ",Nz(PC." & SQL_FIELD_PUBLIC_PROCEDURE_COUNT & ",0) AS " & SQL_FIELD_PUBLIC_PROCEDURE_COUNT & ",Nz(CC." & SQL_FIELD_PROCEDURE_CONFLICT_COUNT & ",0) AS " & SQL_FIELD_PROCEDURE_CONFLICT_COUNT & ",P." & SQL_FIELD_DEPENDENCY_COUNT & ",P." & SQL_FIELD_SOURCE_ANALYSIS_STATUS & "," & _
                            "IIf(Len(Nz(P." & SQL_FIELD_SOURCE_ANALYSIS_DIAGNOSTIC_TEXT & ",''))>0,P." & SQL_FIELD_SOURCE_ANALYSIS_DIAGNOSTIC_TEXT & ",P." & SQL_FIELD_DIAGNOSTIC_TEXT & ") AS " & SQL_FIELD_DIAGNOSTIC_TEXT & ",IIf(P." & SQL_FIELD_LAST_SOURCE_SCANNED_ON & " Is Null,P." & SQL_FIELD_LAST_SCANNED_ON & ",P." & SQL_FIELD_LAST_SOURCE_SCANNED_ON & ") AS " & SQL_FIELD_LAST_SCANNED_ON & " FROM ([" & _
                            ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects") & "] AS P LEFT JOIN [" & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Project_Procedure_Counts") & "] AS PC ON P." & SQL_FIELD_PROJECT_IDENTITY & "=PC." & SQL_FIELD_PROJECT_IDENTITY & ") LEFT JOIN [" & _
                            ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Project_Conflict_Counts") & "] AS CC ON P." & SQL_FIELD_PROJECT_IDENTITY & "=CC." & SQL_FIELD_PROJECT_IDENTITY & " WHERE P." & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " ORDER BY " & OrderByPolicyText("REFERENCE_REVIEW_SUMMARY") & ";"
    ' Thin report-only projection over authoritative persisted Add-In/VBA-project summary plus the current Active References query. It never scans, repairs, loads, registers, or changes references.
    addInReferenceReportSql = "SELECT A." & SQL_FIELD_SOURCE_KIND & ",A." & SQL_FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME & ",A." & SQL_FIELD_PROJECT_NAME & ",A." & SQL_FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME & " AS " & SQL_FIELD_REFERENCE_NAME & ",A." & SQL_FIELD_SOURCE_KIND & " AS " & SQL_FIELD_REFERENCE_KIND & ",A." & SQL_FIELD_REVIEW_IDENTITY & " AS " & SQL_FIELD_REFERENCE_IDENTITY & ",A." & SQL_FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME & " AS " & SQL_FIELD_REFERENCE_VERSION & ",A." & SQL_FIELD_GUID & ",A." & SQL_FIELD_IS_REGISTERED & ",A." & SQL_FIELD_IS_CONNECTED & ",A." & SQL_FIELD_IS_LOADED & ",A." & SQL_FIELD_REVIEW_STATUS & " AS " & SQL_FIELD_REFERENCE_STATE & ",A." & SQL_FIELD_FILE_PATH & ",A." & SQL_FIELD_MODULE_COUNT & ",A." & SQL_FIELD_PUBLIC_PROCEDURE_COUNT & ",A." & SQL_FIELD_PROCEDURE_CONFLICT_COUNT & ",A." & SQL_FIELD_DEPENDENCY_COUNT & ",A." & SQL_FIELD_DIAGNOSTIC_TEXT & _
                ",A." & SQL_FIELD_LAST_SCANNED_ON & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Review_Summary") & "] AS A WHERE False UNION ALL " & _
                              "SELECT A." & SQL_FIELD_SOURCE_KIND & ",A." & SQL_FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME & ",A." & SQL_FIELD_PROJECT_NAME & ",Null AS " & SQL_FIELD_REFERENCE_NAME & ",Null AS " & SQL_FIELD_REFERENCE_KIND & ",Null AS " & SQL_FIELD_REFERENCE_IDENTITY & ",Null AS " & SQL_FIELD_REFERENCE_VERSION & ",A." & SQL_FIELD_GUID & ",A." & SQL_FIELD_IS_REGISTERED & ",A." & SQL_FIELD_IS_CONNECTED & ",A." & SQL_FIELD_IS_LOADED & ",A." & SQL_FIELD_REVIEW_STATUS & " AS " & SQL_FIELD_REFERENCE_STATE & ",A." & SQL_FIELD_FILE_PATH & ",A." & SQL_FIELD_MODULE_COUNT & ",A." & SQL_FIELD_PUBLIC_PROCEDURE_COUNT & ",A." & SQL_FIELD_PROCEDURE_CONFLICT_COUNT & ",A." & SQL_FIELD_DEPENDENCY_COUNT & ",A." & SQL_FIELD_DIAGNOSTIC_TEXT & ",A." & SQL_FIELD_LAST_SCANNED_ON & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Review_Summary") & "] AS A UNION ALL " & _
                              "SELECT 'Reference' AS " & SQL_FIELD_SOURCE_KIND & ",R." & SQL_FIELD_REFERENCE_NAME & " AS " & SQL_FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME & ",R." & SQL_FIELD_LOADED_PROJECT_NAME & " AS " & SQL_FIELD_PROJECT_NAME & ",R." & SQL_FIELD_REFERENCE_NAME & ",R." & SQL_FIELD_REFERENCE_KIND & ",R." & SQL_FIELD_REFERENCE_IDENTITY & ",CStr(Nz(R." & SQL_FIELD_MAJOR_VERSION & ",0)) & '.' & CStr(Nz(R." & SQL_FIELD_MINOR_VERSION & ",0)) & IIf(Len(Nz(R." & SQL_FIELD_FILE_VERSION & ",''))>0,' / ' & R." & SQL_FIELD_FILE_VERSION & ",'') AS " & SQL_FIELD_REFERENCE_VERSION & "," & _
                              "R." & SQL_FIELD_REFERENCE_GUID & " AS " & SQL_FIELD_GUID & ",R." & SQL_FIELD_IS_REGISTERED & ",Null AS " & SQL_FIELD_IS_CONNECTED & ",IIf(Len(Nz(R." & SQL_FIELD_LOADED_PROJECT_NAME & ",''))>0,True,False) AS " & SQL_FIELD_IS_LOADED & ",IIf(Nz(R." & SQL_FIELD_IS_BROKEN & ",False)<>False,'Broken',Nz(R." & SQL_FIELD_SCAN_STATUS & ",'')) AS " & SQL_FIELD_REFERENCE_STATE & ",R." & SQL_FIELD_FULL_PATH & " AS " & SQL_FIELD_FILE_PATH & ",0 AS " & SQL_FIELD_MODULE_COUNT & ",0 AS " & SQL_FIELD_PUBLIC_PROCEDURE_COUNT & ",0 AS " & SQL_FIELD_PROCEDURE_CONFLICT_COUNT & ",0 AS " & SQL_FIELD_DEPENDENCY_COUNT & ",IIf(Len(Nz(R." & SQL_FIELD_LOADED_PROJECT_MATCH_DIAGNOSTIC & ",''))>0,R." & SQL_FIELD_LOADED_PROJECT_MATCH_DIAGNOSTIC & ",R." & SQL_FIELD_DIAGNOSTIC_TEXT & ") AS " & SQL_FIELD_DIAGNOSTIC_TEXT & ",R." & SQL_FIELD_LAST_SCANNED_ON & " FROM [" & _
                                          ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_References_Active") & "] AS R ORDER BY " & OrderByPolicyText("REFERENCE_REPORT_INVENTORY") & ";"
    ' INTENTIONAL KEYED LOOKUP: sqlByQuery provides keyed membership/lookup over source rows or state already read by EnsureReferenceQueries; it is a derived access path, not a second authority.
    Set sqlByQuery = CreateObject("Scripting.Dictionary")
    sqlByQuery.CompareMode = vbTextCompare
    sqlByQuery.Add ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_References_Active"), activeSql
    sqlByQuery.Add ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_References_Available"), availableSql
    sqlByQuery.Add ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_References_Loaded_Projects"), loadedProjectsSql
    sqlByQuery.Add ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_References_Project_Links"), projectLinksSql
    sqlByQuery.Add ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_References_Manager"), managerSql
    sqlByQuery.Add ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Inventory"), addInInventorySql
    sqlByQuery.Add ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_ExternalProject_References"), addInProjectReferencesSql
    sqlByQuery.Add ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_ExternalProject_Modules"), addInModulesSql
    sqlByQuery.Add ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_ExternalProject_Public_Procedures"), addInProceduresSql
    sqlByQuery.Add ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_ExternalProject_Procedure_Conflicts"), addInProcedureConflictsSql
    sqlByQuery.Add ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_ExternalProject_Dependencies"), addInDependenciesSql
    sqlByQuery.Add ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Project_Procedure_Counts"), addInProjectProcedureCountsSql
    sqlByQuery.Add ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Project_Conflict_Counts"), addInProjectConflictCountsSql
    sqlByQuery.Add ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Review_Summary"), addInReviewSummarySql
    sqlByQuery.Add ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Reference_Report_Source"), addInReferenceReportSql
    If sqlByQuery.count <> canonicalQueries.count Then Err.Raise vbObjectError + 9728, "EnsureReferenceQueries", "Reference query implementation-map coverage does not match the canonical Reference query family." & vbCrLf & "Canonical queries: " & CStr(canonicalQueries.count) & vbCrLf & "SQL mappings: " & CStr(sqlByQuery.count)
    For Each item In canonicalQueries
        queryName = CStr(item)
        repairStage = "Reference inventory saved query [" & queryName & "]"
        If sqlByQuery.Exists(queryName) = False Then Err.Raise vbObjectError + 9729, "EnsureReferenceQueries", "The canonical Reference query family contains a query with no SQL implementation mapping." & vbCrLf & "Query: " & queryName
        EnsureSavedQuery queryName, CStr(sqlByQuery(queryName))
    Next item
End Sub

Private Sub EnsureReferenceLookupFields(ByVal tdf As DAO.tableDef)
    'A4 owns the physical lookup contract for reference/add-in tables. Every controlled vocabulary in this family is table-backed; A4 owns its Reference/Add-In physical table/query structure and explicitly A4-owned
    'lookup presentation contract; B2 records editor intent, D3 creates required ADC Support Form editors, and D4 applies the form-dependent editor binding.
    Dim definitions As Collection
    Dim Definition As Variant
    Set definitions = CatalogFieldLookupDefinitions(tdf.Name)
    For Each Definition In definitions
        If StrComp(CStr(Definition(0)), CATALOG_FIELD_LOOKUP_DEFINITION_KIND_TABLE, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 9731, "EnsureReferenceLookupFields", "Reference/add-in lookup fields must use the authoritative table-backed lookup contract." & vbCrLf & "Table: " & tdf.Name & vbCrLf & "Field: " & CStr(Definition(1)) & vbCrLf & "Lookup kind: " & CStr(Definition(0))
        Call EnsureFieldTableLookup(tdf, CStr(Definition(1)), CStr(Definition(2)), CStr(Definition(3)), CStr(Definition(4)), CBool(Definition(5)), CStr(Definition(8)))
        VerifyReferenceTableLookupDefinition tdf, Definition
    Next Definition
    ClearUnexpectedReferenceLookupFields tdf, definitions
End Sub

' Purpose: Verifies reference table lookup definition for tdf, definition.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   tdf (Input; DAO.tableDef): DAO.TableDef representing the table whose fields, indexes, properties, or metadata are being inspected or modified.
'   Definition (Input; Variant): Caller-supplied definition data used as the authoritative definition input for this call, avoiding an unnecessary reload when already available.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub VerifyReferenceTableLookupDefinition(ByVal tdf As DAO.tableDef, ByVal Definition As Variant)
    Dim fld As DAO.Field
    Dim expectedRowSource As String
    Dim expectedWidths As String
    Dim actualText As String
    Dim lookupTableName As String
    Dim fieldName As String
    Dim codeFieldName As String
    Dim displayFieldName As String
    Dim limitToListValue As Boolean
    fieldName = CStr(Definition(1))
    lookupTableName = CStr(Definition(2))
    codeFieldName = CStr(Definition(3))
    displayFieldName = CStr(Definition(4))
    limitToListValue = CBool(Definition(5))
    If DAOFieldExists(tdf, fieldName) = False Then Err.Raise vbObjectError + 9730, "VerifyReferenceTableLookupDefinition", "An authoritative Reference lookup field is missing." & vbCrLf & "Table: " & tdf.Name & vbCrLf & "Field: " & fieldName
    Set fld = tdf.fields(fieldName)
    expectedRowSource = TableLookupRowSource(lookupTableName, codeFieldName, displayFieldName, TableFieldExists(lookupTableName, FIELD_LOOKUP_ACTIVE), StrComp(codeFieldName, FIELD_SETTING_KEY, vbTextCompare) = 0, TableFieldExists(lookupTableName, FIELD_MANAGED_OBJECT_DEPENDENCY_SORT_ORDER))
    expectedWidths = NormalizeLookupColumnWidthsForAccess(CStr(Definition(8)), 2, 1, limitToListValue)
    If StrComp(ReferenceFieldPropertyText(fld, "RowSourceType"), "Table/Query", vbTextCompare) <> 0 Then GoTo Mismatch
    If StrComp(ReferenceFieldPropertyText(fld, "RowSource"), expectedRowSource, vbBinaryCompare) <> 0 Then GoTo Mismatch
    If CLng(Val(ReferenceFieldPropertyText(fld, "BoundColumn"))) <> 1 Then GoTo Mismatch
    If CLng(Val(ReferenceFieldPropertyText(fld, "ColumnCount"))) <> 2 Then GoTo Mismatch
    If StrComp(ReferenceFieldPropertyText(fld, "ColumnWidths"), expectedWidths, vbTextCompare) <> 0 Then GoTo Mismatch
    actualText = ReferenceFieldPropertyText(fld, "LimitToList")
    If CBool(Val(actualText) <> 0 Or StrComp(actualText, "True", vbTextCompare) = 0) <> limitToListValue Then GoTo Mismatch
    Exit Sub
Mismatch:
    Err.Raise vbObjectError + 9730, "VerifyReferenceTableLookupDefinition", "An authoritative table-backed Reference lookup did not verify after repair." & vbCrLf & "Table: " & tdf.Name & vbCrLf & "Field: " & fieldName & vbCrLf & "Lookup table: " & lookupTableName
End Sub

Private Function ReferenceFieldPropertyText(ByVal fld As DAO.Field, ByVal propertyName As String) As String
    On Error Resume Next
    Err.Clear
    ReferenceFieldPropertyText = CStr(Nz(fld.properties(propertyName).value, vbNullString))
    If Err.Number <> 0 Then ReferenceFieldPropertyText = vbNullString: Err.Clear
    On Error GoTo 0
End Function

Private Sub ClearUnexpectedReferenceLookupFields(ByVal tdf As DAO.tableDef, ByVal tableLookupDefinitions As Collection)
    Dim schemaDefinitions As Collection
    Dim expected As Object
    Dim Definition As Variant
    Dim schemaDefinition As Variant
    Dim fieldName As String
    Set expected = CreateObject("Scripting.Dictionary")
    expected.CompareMode = vbTextCompare
    For Each Definition In tableLookupDefinitions
        expected(CStr(Definition(1))) = True
    Next Definition
    Set schemaDefinitions = ReferenceTableSchemaDefinitions(tdf.Name)
    For Each schemaDefinition In schemaDefinitions
        fieldName = CStr(schemaDefinition(0))
        If expected.Exists(fieldName) = False Then ClearFieldLookupPresentationProperties tdf, fieldName
    Next schemaDefinition
End Sub

Private Sub ReportReferenceFailure(ByVal sourceProcedure As String, ByVal headline As String, ByVal showMessage As Boolean, Optional ByVal savedErrNumber As Long = 0, Optional ByVal savedErrSource As String = vbNullString, Optional ByVal savedErrDescription As String = vbNullString)
    Dim messageText As String
    If savedErrNumber = 0 Then savedErrNumber = Err.Number
    If Len(savedErrSource) = 0 Then savedErrSource = Err.Source
    If Len(savedErrDescription) = 0 Then savedErrDescription = Err.Description
    If savedErrNumber = 0 Then savedErrNumber = vbObjectError + 9705
    If Len(savedErrSource) = 0 Then savedErrSource = sourceProcedure
    If Len(savedErrDescription) = 0 Then savedErrDescription = headline
    If IsCancellationError(savedErrNumber) Then Err.Raise savedErrNumber, savedErrSource, savedErrDescription
    messageText = headline & vbCrLf & vbCrLf & "Procedure: " & sourceProcedure & vbCrLf & "Err.Number: " & CStr(savedErrNumber) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
    LogTraceMessage TRACE_LEVEL_ERROR, "References", sourceProcedure, headline, messageText, sourceProcedure, savedErrNumber, savedErrSource, savedErrDescription, vbNullString, "ReferenceInventory", "CodeDb", "References"
    If showMessage Then
        CoreShowErrorMessage messageText, vbCritical, DEFAULT_UTILITY_LONG_NAME & " References", sourceProcedure, "Tables"
        Exit Sub
    End If
    Err.Raise savedErrNumber, sourceProcedure, messageText
End Sub

' Purpose: Resolves references configured physical object names for setup from the supplied/current ADC state, using `AppendReferencePhysicalObjectNames` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters: None.
' Returns: Collection: Returns the ordered Collection assembled by the routine; callers receive the object reference and should treat its membership/order according to the routine contract.
Public Function ReferencesConfiguredPhysicalObjectNamesForSetup() As Collection
    ' INTENTIONAL FILTERED PARTITION: names is a consumer-required subset/projection derived from the source enumerated by ReferencesConfiguredPhysicalObjectNamesForSetup; its membership is not independently
    ' authoritative.
    Dim names As New Collection
    Dim inventoryNames As Collection
    Dim loadedVbaAnalysisNames As Collection
    Dim crosswalkNames As Collection
    Dim queryNames As Collection
    ResolveReferenceFamilies inventoryNames, loadedVbaAnalysisNames, crosswalkNames, queryNames
    AppendReferencePhysicalObjectNames names, inventoryNames, "TABLE"
    AppendReferencePhysicalObjectNames names, loadedVbaAnalysisNames, "TABLE"
    AppendReferencePhysicalObjectNames names, crosswalkNames, "TABLE"
    AppendReferencePhysicalObjectNames names, queryNames, "QUERY"
    Set ReferencesConfiguredPhysicalObjectNamesForSetup = names
End Function


' Purpose: Repairs reference inventory foundation for setup.
' Public role: Public module entry point. This procedure is listed in the module header as an intentional externally callable ADC entry point and may be invoked directly by Access/VBA callers.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: False.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub RepairReferenceInventoryFoundationForSetup(Optional ByVal showMessage As Boolean = False)
    RequireSetupDomainMutationAllowed SETUP_DOMAIN_A4_REFERENCE_INVENTORY, "RepairReferenceInventoryFoundationForSetup"
    On Error GoTo ErrHandler
    Dim repairStage As String
    Dim timingText As String
    Dim operationTimer As Single
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim completionDetail As String
    Dim db As DAO.Database
    Dim currentToken As String
    Dim storedToken As String
    Dim failureDetail As String
    Dim comparisonSkippedCount As Long
    Dim matchedFamilyCount As Long
    Dim dirtyFamilyCount As Long
    Dim inventoryNames As Collection
    Dim loadedVbaAnalysisNames As Collection
    Dim crosswalkNames As Collection
    Dim queryNames As Collection
    Dim familyMaterializationTimer As Single
    Dim familyMaterializationElapsedSeconds As Double
    Dim refreshTimer As Single
    Dim initialTableDefsRefreshElapsedSeconds As Double
    Dim initialQueryDefsRefreshElapsedSeconds As Double
    Dim descriptionDetail As String
    operationTimer = Timer
    CorePulseProgress "Reference inventory foundation | materializing canonical families"
    PumpSetupResponsiveness True
    familyMaterializationTimer = Timer
    'Materialize all four canonical Reference families in one managed-definition pass when no operation snapshot is active.
    ResolveReferenceFamilies inventoryNames, loadedVbaAnalysisNames, crosswalkNames, queryNames
    familyMaterializationElapsedSeconds = ElapsedSecondsFromTimer(familyMaterializationTimer)
    CorePulseProgress "Reference inventory foundation | canonical families ready"
    PumpSetupResponsiveness True
    Set db = ADCDb()
    CorePulseProgress "Reference inventory foundation | refreshing TableDefs metadata"
    refreshTimer = Timer
    db.TableDefs.Refresh
    initialTableDefsRefreshElapsedSeconds = ElapsedSecondsFromTimer(refreshTimer)
    PumpSetupResponsiveness True
    CorePulseProgress "Reference inventory foundation | refreshing QueryDefs metadata"
    refreshTimer = Timer
    db.QueryDefs.Refresh
    initialQueryDefsRefreshElapsedSeconds = ElapsedSecondsFromTimer(refreshTimer)
    PumpSetupResponsiveness True
    storedToken = vbNullString
    If ReferencePersistedTokenAvailable(db, REFERENCE_INVENTORY_TOKEN_PROPERTY, storedToken) Then
        failureDetail = vbNullString
        currentToken = ReferenceTableFamilyStateToken(db, REFERENCE_INVENTORY_TOKEN_VERSION, inventoryNames, failureDetail)
    Else
        currentToken = vbNullString
        comparisonSkippedCount = comparisonSkippedCount + 1
    End If
    If Len(currentToken) > 0 And Len(storedToken) > 0 And StrComp(storedToken, currentToken, vbBinaryCompare) = 0 Then
        matchedFamilyCount = matchedFamilyCount + 1
    Else
        dirtyFamilyCount = dirtyFamilyCount + 1
        repairStage = "Reference Inventory tables"
        RepairReferenceTableFamily inventoryNames
        db.TableDefs.Refresh
        failureDetail = vbNullString
        currentToken = ReferenceTableFamilyStateToken(db, REFERENCE_INVENTORY_TOKEN_VERSION, inventoryNames, failureDetail)
        If Len(currentToken) = 0 Then Err.Raise vbObjectError + 9720, "RepairReferenceInventoryFoundationForSetup", "Reference Inventory table verification did not produce a reusable token. " & failureDetail
        Call EnsureDaoPropertyValue(db, REFERENCE_INVENTORY_TOKEN_PROPERTY, dbText, currentToken, "RepairReferenceInventoryFoundationForSetup")
    End If
    storedToken = vbNullString
    If ReferencePersistedTokenAvailable(db, REFERENCE_LOADED_VBA_ANALYSIS_TOKEN_PROPERTY, storedToken) Then
        failureDetail = vbNullString
        currentToken = ReferenceTableFamilyStateToken(db, REFERENCE_LOADED_VBA_ANALYSIS_TOKEN_VERSION, loadedVbaAnalysisNames, failureDetail)
    Else
        currentToken = vbNullString
        comparisonSkippedCount = comparisonSkippedCount + 1
    End If
    If Len(currentToken) > 0 And Len(storedToken) > 0 And StrComp(storedToken, currentToken, vbBinaryCompare) = 0 Then
        matchedFamilyCount = matchedFamilyCount + 1
    Else
        dirtyFamilyCount = dirtyFamilyCount + 1
        repairStage = "Loaded VBA Analysis tables"
        RepairReferenceTableFamily loadedVbaAnalysisNames
        db.TableDefs.Refresh
        failureDetail = vbNullString
        currentToken = ReferenceTableFamilyStateToken(db, REFERENCE_LOADED_VBA_ANALYSIS_TOKEN_VERSION, loadedVbaAnalysisNames, failureDetail)
        If Len(currentToken) = 0 Then Err.Raise vbObjectError + 9721, "RepairReferenceInventoryFoundationForSetup", "Loaded VBA Analysis table verification did not produce a reusable token. " & failureDetail
        Call EnsureDaoPropertyValue(db, REFERENCE_LOADED_VBA_ANALYSIS_TOKEN_PROPERTY, dbText, currentToken, "RepairReferenceInventoryFoundationForSetup")
    End If
    storedToken = vbNullString
    If ReferencePersistedTokenAvailable(db, REFERENCE_CROSSWALK_TOKEN_PROPERTY, storedToken) Then
        failureDetail = vbNullString
        currentToken = ReferenceTableFamilyStateToken(db, REFERENCE_CROSSWALK_TOKEN_VERSION, crosswalkNames, failureDetail)
    Else
        currentToken = vbNullString
        comparisonSkippedCount = comparisonSkippedCount + 1
    End If
    If Len(currentToken) > 0 And Len(storedToken) > 0 And StrComp(storedToken, currentToken, vbBinaryCompare) = 0 Then
        matchedFamilyCount = matchedFamilyCount + 1
    Else
        dirtyFamilyCount = dirtyFamilyCount + 1
        repairStage = "Reference crosswalk and manager storage"
        RepairReferenceTableFamily crosswalkNames
        db.TableDefs.Refresh
        failureDetail = vbNullString
        currentToken = ReferenceTableFamilyStateToken(db, REFERENCE_CROSSWALK_TOKEN_VERSION, crosswalkNames, failureDetail)
        If Len(currentToken) = 0 Then Err.Raise vbObjectError + 9722, "RepairReferenceInventoryFoundationForSetup", "Reference crosswalk table verification did not produce a reusable token. " & failureDetail
        Call EnsureDaoPropertyValue(db, REFERENCE_CROSSWALK_TOKEN_PROPERTY, dbText, currentToken, "RepairReferenceInventoryFoundationForSetup")
    End If
    storedToken = vbNullString
    If ReferencePersistedTokenAvailable(db, REFERENCE_QUERY_TOKEN_PROPERTY, storedToken) Then
        failureDetail = vbNullString
        currentToken = ReferenceQueryFamilyStateToken(db, REFERENCE_QUERY_TOKEN_VERSION, queryNames, failureDetail)
    Else
        currentToken = vbNullString
        comparisonSkippedCount = comparisonSkippedCount + 1
    End If
    If Len(currentToken) > 0 And Len(storedToken) > 0 And StrComp(storedToken, currentToken, vbBinaryCompare) = 0 Then
        matchedFamilyCount = matchedFamilyCount + 1
    Else
        dirtyFamilyCount = dirtyFamilyCount + 1
        repairStage = "Reference inventory saved queries"
        EnsureReferenceQueries queryNames, repairStage
        db.QueryDefs.Refresh
        failureDetail = vbNullString
        currentToken = ReferenceQueryFamilyStateToken(db, REFERENCE_QUERY_TOKEN_VERSION, queryNames, failureDetail)
        If Len(currentToken) = 0 Then Err.Raise vbObjectError + 9723, "RepairReferenceInventoryFoundationForSetup", "Reference query verification did not produce a reusable token. " & failureDetail
        Call EnsureDaoPropertyValue(db, REFERENCE_QUERY_TOKEN_PROPERTY, dbText, currentToken, "RepairReferenceInventoryFoundationForSetup")
    End If
    timingText = "A4 structural families matched: " & CStr(matchedFamilyCount) & "; dirty/repaired: " & CStr(dirtyFamilyCount) & "; impossible compact comparisons skipped: " & CStr(comparisonSkippedCount) & vbCrLf & _
                 "Canonical family materialization: " & FormattedElapsedSeconds(familyMaterializationElapsedSeconds) & vbCrLf & _
                 "Initial TableDefs.Refresh: " & FormattedElapsedSeconds(initialTableDefsRefreshElapsedSeconds) & vbCrLf & _
                 "Initial QueryDefs.Refresh: " & FormattedElapsedSeconds(initialQueryDefsRefreshElapsedSeconds) & vbCrLf & _
                 "Live reference/project/Add-In scanning: not performed." & vbCrLf & _
                 "Generated reference forms: not inspected or mutated." & vbCrLf & _
                 "Elapsed: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(operationTimer))
    If comparisonSkippedCount > 0 Then AddTraceNote "A4 persisted family proof was absent for " & CStr(comparisonSkippedCount) & " family/families; impossible pre-repair current-token comparisons were skipped and fresh successor proof was persisted after repair."
    If SetupDomainMutationWritesAllowedForCurrentPart(SETUP_DOMAIN_G1_DESCRIPTIONS) Then
        ReconcileReferenceInventoryStandaloneDescriptions inventoryNames, loadedVbaAnalysisNames, crosswalkNames, queryNames, descriptionDetail
        timingText = timingText & vbCrLf & descriptionDetail
    End If
    completionDetail = "A4 repaired/verified reference and Add-In inventory infrastructure only." & vbCrLf & timingText
    LogTraceMessage TRACE_LEVEL_INFO, "References", "RepairReferenceInventoryFoundationForSetup", "A4 reference and Add-In inventory foundation verified.", completionDetail, "A4 - Repair reference/Add-In inventory foundation", 0, vbNullString, vbNullString, vbNullString, "ReferenceInventory", "CodeDb", "References"
    If showMessage Then ReportCoreActionSuccess "RepairReferenceInventoryFoundationForSetup", "References", "A4 reference and Add-In inventory foundation verified successfully.", completionDetail, "A4 - Repair reference/Add-In inventory foundation", vbNullString, "ReferenceInventory", "CodeDb", "References", True, "Reference and Add-in Inventory", vbInformation, True, "Tables", False
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    ReportReferenceFailure "RepairReferenceInventoryFoundationForSetup", "Could not create or repair the A4 reference/Add-In inventory foundation." & vbCrLf & "Failed stage: " & IIf(Len(repairStage) = 0, "(not established)", repairStage), showMessage, savedErrNumber, savedErrSource, savedErrDescription
End Sub

Private Sub ReconcileReferenceInventoryStandaloneDescriptions(ByVal inventoryNames As Collection, ByVal loadedVbaAnalysisNames As Collection, ByVal crosswalkNames As Collection, ByVal queryNames As Collection, ByRef detailText As String)
    'A4 remains structural-only during numbered Setup. A direct standalone A4 call completes only the managed Reference tables, fields, and queries that A4 owns before it reports success.
    On Error GoTo ErrHandler
    Dim tableNames As New Collection
    Dim item As Variant
    Dim descriptionChanged As Variant
    Dim fieldDescriptionChangeCount As Variant
    Dim objectDescriptionChanges As Long
    Dim fieldDescriptionChanges As Long
    Dim fieldDetail As String
    For Each item In inventoryNames
        tableNames.Add CStr(item)
    Next item
    For Each item In loadedVbaAnalysisNames
        tableNames.Add CStr(item)
    Next item
    For Each item In crosswalkNames
        tableNames.Add CStr(item)
    Next item
    For Each item In tableNames
        descriptionChanged = False
        If ApplyAndVerifyManagedObjectDescriptionForStandaloneRepair(CStr(item), ADCAT_OBJECT_TYPE_TABLE, descriptionChanged) = False Then Err.Raise vbObjectError + 10125, "ReconcileReferenceInventoryStandaloneDescriptions", "Reference managed-table Description reconciliation returned False." & vbCrLf & "Table: " & CStr(item)
        If CBool(descriptionChanged) Then objectDescriptionChanges = objectDescriptionChanges + 1
        fieldDescriptionChangeCount = CLng(0)
        fieldDetail = vbNullString
        If ApplyAndVerifyManagedTableFieldDescriptionsForStandaloneRepair(CStr(item), fieldDescriptionChangeCount, fieldDetail, True) = False Then Err.Raise vbObjectError + 10126, "ReconcileReferenceInventoryStandaloneDescriptions", fieldDetail
        fieldDescriptionChanges = fieldDescriptionChanges + CLng(Nz(fieldDescriptionChangeCount, 0))
    Next item
    For Each item In queryNames
        descriptionChanged = False
        If ApplyAndVerifyManagedObjectDescriptionForStandaloneRepair(CStr(item), ADCAT_OBJECT_TYPE_QUERY, descriptionChanged) = False Then Err.Raise vbObjectError + 10127, "ReconcileReferenceInventoryStandaloneDescriptions", "Reference managed-query Description reconciliation returned False." & vbCrLf & "Query: " & CStr(item)
        If CBool(descriptionChanged) Then objectDescriptionChanges = objectDescriptionChanges + 1
    Next item
    'A4 structural proof intentionally excludes standalone Description state, so this standalone completion does not rewrite the already-established A4 proof.
    detailText = "Standalone A4 managed descriptions reconciled; object descriptions changed=" & CStr(objectDescriptionChanges) & "; field descriptions changed=" & CStr(fieldDescriptionChanges) & "."
    Exit Sub
ErrHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "ReconcileReferenceInventoryStandaloneDescriptions", "Could not complete standalone A4 managed Description reconciliation." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub EnsureLoadedVBAProjectsTable()
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim schemaDefinitions As Collection
    Set db = ADCDb()
    Set schemaDefinitions = ReferenceTableSchemaDefinitions(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects"))
    If DaoTableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects")) = False Then
        Set tdf = db.CreateTableDef(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects"))
        AppendReferenceSchemaFields tdf, schemaDefinitions
        db.TableDefs.Append tdf
    End If
    Set tdf = db.TableDefs(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects"))
    EnsureReferenceSchemaFields tdf, schemaDefinitions
    EnsureTableDefFieldOrderFromText tdf, SchemaDefinitionFieldOrderText(schemaDefinitions)
    EnsureSingleFieldIndex db, tdf.Name, INDEX_PK_LOADED_VBA_PROJECTS, FIELD_ID, True, True
    EnsureSingleFieldIndex db, tdf.Name, INDEX_UX_LOADED_VBA_PROJECTS_IDENTITY, FIELD_PROJECT_IDENTITY, True, False
    EnsureSingleFieldIndex db, tdf.Name, INDEX_IX_LOADED_VBA_PROJECTS_PATH_KEY, FIELD_PROJECT_PATH_KEY, False, False
    EnsureSingleFieldIndex db, tdf.Name, INDEX_IX_LOADED_VBA_PROJECTS_NAME_KEY, FIELD_PROJECT_NAME_KEY, False, False
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "EnsureLoadedVBAProjectsTable", "Could not create or repair DBInv_VBAPrjs_Ld." & vbCrLf & Err.Description
End Sub

Private Sub EnsureAddInInventoryTable(ByVal tableName As String)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim schemaDefinitions As Collection
    Dim indexPrefix As String
    Set db = ADCDb()
    Set schemaDefinitions = ReferenceTableSchemaDefinitions(tableName)
    If DaoTableExists(db, tableName) = False Then
        Set tdf = db.CreateTableDef(tableName)
        AppendReferenceSchemaFields tdf, schemaDefinitions
        db.TableDefs.Append tdf
    End If
    Set tdf = db.TableDefs(tableName)
    EnsureReferenceSchemaFields tdf, schemaDefinitions
    EnsureTableDefFieldOrderFromText tdf, SchemaDefinitionFieldOrderText(schemaDefinitions)
    If StrComp(tableName, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_VBEAddIns"), vbTextCompare) = 0 Then indexPrefix = "VBE" Else indexPrefix = "COM"
    EnsureSingleFieldIndex db, tdf.Name, INDEX_PREFIX_PRIMARY_KEY & indexPrefix, FIELD_ID, True, True
    EnsureSingleFieldIndex db, tdf.Name, INDEX_PREFIX_UNIQUE_2 & indexPrefix & INDEX_SUFFIX_IDENTITY, FIELD_ADD_IN_IDENTITY, True, False
    EnsureSingleFieldIndex db, tdf.Name, INDEX_PREFIX_NONUNIQUE & indexPrefix & INDEX_SUFFIX_PROG_ID, FIELD_PROG_ID, False, False
    EnsureSingleFieldIndex db, tdf.Name, INDEX_PREFIX_NONUNIQUE & indexPrefix & INDEX_SUFFIX_GUID, FIELD_GUID, False, False
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "EnsureAddInInventoryTable", "Could not create or repair add-in inventory table [" & tableName & "]." & vbCrLf & Err.Description
End Sub

Private Sub EnsureAccessAddInInventoryTable()
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim schemaDefinitions As Collection
    Set db = ADCDb()
    Set schemaDefinitions = ReferenceTableSchemaDefinitions(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AccessAddIns"))
    If DaoTableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AccessAddIns")) = False Then
        Set tdf = db.CreateTableDef(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AccessAddIns"))
        AppendReferenceSchemaFields tdf, schemaDefinitions
        db.TableDefs.Append tdf
    End If
    Set tdf = db.TableDefs(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AccessAddIns"))
    EnsureReferenceSchemaFields tdf, schemaDefinitions
    EnsureTableDefFieldOrderFromText tdf, SchemaDefinitionFieldOrderText(schemaDefinitions)
    EnsureSingleFieldIndex db, tdf.Name, INDEX_PK_ACCESS_ADD_INS, FIELD_ID, True, True
    EnsureSingleFieldIndex db, tdf.Name, INDEX_UX_ACCESS_ADD_INS_IDENTITY, FIELD_ADD_IN_IDENTITY, True, False
    RemoveReferenceIndexIfPresent tdf, INDEX_IX_ACCESS_ADD_INS_PROJECT
    EnsureSingleFieldIndex db, tdf.Name, INDEX_IX_ACCESS_ADD_INS_PROCEDURE, FIELD_ENTRY_POINT_PROCEDURE, False, False
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "EnsureAccessAddInInventoryTable", "Could not create or repair Access add-in registration inventory." & vbCrLf & Err.Description
End Sub

Private Sub EnsureLoadedVBAProjectReferencesTable()
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim schemaDefinitions As Collection
    Set db = ADCDb()
    Set schemaDefinitions = ReferenceTableSchemaDefinitions(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjectReferences"))
    If DaoTableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjectReferences")) = False Then
        Set tdf = db.CreateTableDef(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjectReferences"))
        AppendReferenceSchemaFields tdf, schemaDefinitions
        db.TableDefs.Append tdf
    End If
    Set tdf = db.TableDefs(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjectReferences"))
    EnsureReferenceSchemaFields tdf, schemaDefinitions
    EnsureTableDefFieldOrderFromText tdf, SchemaDefinitionFieldOrderText(schemaDefinitions)
    EnsureSingleFieldIndex db, tdf.Name, INDEX_PK_LOADED_VBA_PROJECT_REFERENCES, FIELD_ID, True, True
    EnsureSingleFieldIndex db, tdf.Name, INDEX_UX_LOADED_VBA_PROJECT_REFERENCES_IDENTITY, FIELD_REFERENCE_IDENTITY, True, False
    RemoveReferenceIndexIfPresent tdf, INDEX_IX_LOADED_VBA_PROJECT_REFERENCES_PROJECT
End Sub

Private Sub EnsureLoadedVBAModulesTable()
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim schemaDefinitions As Collection
    Set db = ADCDb()
    Set schemaDefinitions = ReferenceTableSchemaDefinitions(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAModules"))
    If DaoTableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAModules")) = False Then
        Set tdf = db.CreateTableDef(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAModules"))
        AppendReferenceSchemaFields tdf, schemaDefinitions
        db.TableDefs.Append tdf
    End If
    Set tdf = db.TableDefs(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAModules"))
    EnsureReferenceSchemaFields tdf, schemaDefinitions
    EnsureTableDefFieldOrderFromText tdf, SchemaDefinitionFieldOrderText(schemaDefinitions)
    EnsureSingleFieldIndex db, tdf.Name, INDEX_PK_LOADED_VBA_MODULES, FIELD_ID, True, True
    EnsureSingleFieldIndex db, tdf.Name, INDEX_UX_LOADED_VBA_MODULES_IDENTITY, FIELD_MODULE_IDENTITY, True, False
    RemoveReferenceIndexIfPresent tdf, INDEX_IX_LOADED_VBA_MODULES_PROJECT
    EnsureSingleFieldIndex db, tdf.Name, INDEX_IX_LOADED_VBA_MODULES_NAME, FIELD_MODULE_NAME, False, False
End Sub

Private Sub EnsureLoadedVBAProceduresTable()
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim schemaDefinitions As Collection
    Set db = ADCDb()
    Set schemaDefinitions = ReferenceTableSchemaDefinitions(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProcedures"))
    If DaoTableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProcedures")) = False Then
        Set tdf = db.CreateTableDef(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProcedures"))
        AppendReferenceSchemaFields tdf, schemaDefinitions
        db.TableDefs.Append tdf
    End If
    Set tdf = db.TableDefs(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProcedures"))
    EnsureReferenceSchemaFields tdf, schemaDefinitions
    EnsureTableDefFieldOrderFromText tdf, SchemaDefinitionFieldOrderText(schemaDefinitions)
    EnsureSingleFieldIndex db, tdf.Name, INDEX_PK_LOADED_VBA_PROCEDURES, FIELD_ID, True, True
    EnsureSingleFieldIndex db, tdf.Name, INDEX_UX_LOADED_VBA_PROCEDURES_IDENTITY, FIELD_PROCEDURE_IDENTITY, True, False
    RemoveReferenceIndexIfPresent tdf, INDEX_IX_LOADED_VBA_PROCEDURES_PROJECT
    EnsureSingleFieldIndex db, tdf.Name, INDEX_IX_LOADED_VBA_PROCEDURES_NAME, FIELD_PROCEDURE_NAME_KEY, False, False
End Sub

Private Sub EnsureLoadedVBADependenciesTable()
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim schemaDefinitions As Collection
    Set db = ADCDb()
    Set schemaDefinitions = ReferenceTableSchemaDefinitions(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBADependencies"))
    If DaoTableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBADependencies")) = False Then
        Set tdf = db.CreateTableDef(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBADependencies"))
        AppendReferenceSchemaFields tdf, schemaDefinitions
        db.TableDefs.Append tdf
    End If
    Set tdf = db.TableDefs(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBADependencies"))
    EnsureReferenceSchemaFields tdf, schemaDefinitions
    EnsureTableDefFieldOrderFromText tdf, SchemaDefinitionFieldOrderText(schemaDefinitions)
    EnsureSingleFieldIndex db, tdf.Name, INDEX_PK_LOADED_VBA_DEPENDENCIES, FIELD_ID, True, True
    EnsureSingleFieldIndex db, tdf.Name, INDEX_UX_LOADED_VBA_DEPENDENCIES_IDENTITY, FIELD_DEPENDENCY_IDENTITY, True, False
    RemoveReferenceIndexIfPresent tdf, INDEX_IX_LOADED_VBA_DEPENDENCIES_PROJECT
    RemoveReferenceIndexIfPresent tdf, INDEX_IX_LOADED_VBA_DEPENDENCIES_MODULE
End Sub

Private Sub EnsureActiveReferenceProjectLinksTable()
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim schemaDefinitions As Collection
    Set db = ADCDb()
    Set schemaDefinitions = ReferenceTableSchemaDefinitions(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActiveReferenceProjectLinks"))
    If DaoTableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActiveReferenceProjectLinks")) = False Then
        Set tdf = db.CreateTableDef(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActiveReferenceProjectLinks"))
        AppendReferenceSchemaFields tdf, schemaDefinitions
        db.TableDefs.Append tdf
    End If
    Set tdf = db.TableDefs(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActiveReferenceProjectLinks"))
    EnsureReferenceSchemaFields tdf, schemaDefinitions
    EnsureTableDefFieldOrderFromText tdf, SchemaDefinitionFieldOrderText(schemaDefinitions)
    EnsureSingleFieldIndex db, tdf.Name, INDEX_PK_ACTIVE_REFERENCE_PROJECT_LINKS, FIELD_ID, True, True
    EnsureSingleFieldIndex db, tdf.Name, INDEX_UX_ACTIVE_REFERENCE_PROJECT_LINKS_REFERENCE, FIELD_REFERENCE_IDENTITY, True, False
    RemoveReferenceIndexIfPresent tdf, INDEX_IX_ACTIVE_REFERENCE_PROJECT_LINKS_PROJECT
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "EnsureActiveReferenceProjectLinksTable", "Could not create or repair DBInv_VBARefs_Active_ReferencedPrjLinks." & vbCrLf & Err.Description
End Sub

Private Sub EnsureActiveReferencesTable()
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim schemaDefinitions As Collection
    Set db = ADCDb()
    Set schemaDefinitions = ReferenceTableSchemaDefinitions(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActiveReferences"))
    If DaoTableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActiveReferences")) = False Then
        Set tdf = db.CreateTableDef(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActiveReferences"))
        AppendReferenceSchemaFields tdf, schemaDefinitions
        db.TableDefs.Append tdf
    End If
    Set tdf = db.TableDefs(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActiveReferences"))
    EnsureReferenceSchemaFields tdf, schemaDefinitions
    EnsureTableDefFieldOrderFromText tdf, SchemaDefinitionFieldOrderText(schemaDefinitions)
    EnsureSingleFieldIndex db, tdf.Name, INDEX_PK_ACTIVE_REFERENCES, FIELD_ID, True, True
    RemoveReferenceIndexIfPresent tdf, INDEX_UX_ACTIVE_REFERENCES_PRIORITY
    RemoveReferenceIndexIfPresent tdf, INDEX_IX_ACTIVE_REFERENCES_IDENTITY
    EnsureSingleFieldIndex db, tdf.Name, INDEX_IX_ACTIVE_REFERENCES_PRIORITY, FIELD_REFERENCE_PRIORITY, False, False
    EnsureSingleFieldIndex db, tdf.Name, INDEX_UX_ACTIVE_REFERENCES_IDENTITY, FIELD_REFERENCE_IDENTITY, True, False
    EnsureSingleFieldIndex db, tdf.Name, INDEX_IX_ACTIVE_REFERENCES_GUID, FIELD_REFERENCE_GUID, False, False
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "EnsureActiveReferencesTable", "Could not create or repair DBInv_VBARefs_Active." & vbCrLf & Err.Description
End Sub

Private Sub EnsureAvailableReferencesTable()
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim schemaDefinitions As Collection
    Set db = ADCDb()
    Set schemaDefinitions = ReferenceTableSchemaDefinitions(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AvailableReferences"))
    If DaoTableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AvailableReferences")) = False Then
        Set tdf = db.CreateTableDef(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AvailableReferences"))
        AppendReferenceSchemaFields tdf, schemaDefinitions
        db.TableDefs.Append tdf
    End If
    Set tdf = db.TableDefs(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AvailableReferences"))
    EnsureReferenceSchemaFields tdf, schemaDefinitions
    EnsureTableDefFieldOrderFromText tdf, SchemaDefinitionFieldOrderText(schemaDefinitions)
    EnsureSingleFieldIndex db, tdf.Name, INDEX_PK_AVAILABLE_REFERENCES, FIELD_ID, True, True
    EnsureSingleFieldIndex db, tdf.Name, INDEX_UX_AVAILABLE_REFERENCES_IDENTITY, FIELD_REFERENCE_IDENTITY, True, False
    EnsureSingleFieldIndex db, tdf.Name, INDEX_IX_AVAILABLE_REFERENCES_GUID, FIELD_REFERENCE_GUID, False, False
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "EnsureAvailableReferencesTable", "Could not create or repair DBInv_VBARefs_Avail." & vbCrLf & Err.Description
End Sub

' Purpose: Checks whether reference inventory foundation already satisfies the exact state required by the current Setup run so the owning phase can determine whether repair is necessary.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   detailText (Output; String): Receives diagnostic detail text used to explain validation, lookup, recovery, or failure results to the caller.
' Returns: Boolean: True when reference inventory foundation matches the exact state required by the current Setup contract; False when the owning phase still needs repair, refresh, or reconciliation.
Public Function ReferenceInventoryFoundationCurrentForSetup(ByRef detailText As String) As Boolean
    On Error GoTo NotCurrent
    Dim db As DAO.Database
    Dim currentToken As String
    Dim storedInventoryToken As String
    Dim storedLoadedVbaAnalysisToken As String
    Dim storedCrosswalkToken As String
    Dim storedQueryToken As String
    Dim failureDetail As String
    Dim matchedCount As Long
    Dim inventoryNames As Collection
    Dim loadedVbaAnalysisNames As Collection
    Dim crosswalkNames As Collection
    Dim queryNames As Collection
    detailText = vbNullString
    Set db = ADCDb()
    If ReferencePersistedTokenAvailable(db, REFERENCE_INVENTORY_TOKEN_PROPERTY, storedInventoryToken) = False Then detailText = "A4 Reference Inventory table family is unseeded; persisted proof is missing.": Exit Function
    If ReferencePersistedTokenAvailable(db, REFERENCE_LOADED_VBA_ANALYSIS_TOKEN_PROPERTY, storedLoadedVbaAnalysisToken) = False Then detailText = "A4 Loaded VBA Analysis table family is unseeded; persisted proof is missing.": Exit Function
    If ReferencePersistedTokenAvailable(db, REFERENCE_CROSSWALK_TOKEN_PROPERTY, storedCrosswalkToken) = False Then detailText = "A4 reference crosswalk family is unseeded; persisted proof is missing.": Exit Function
    If ReferencePersistedTokenAvailable(db, REFERENCE_QUERY_TOKEN_PROPERTY, storedQueryToken) = False Then detailText = "A4 saved-query family is unseeded; persisted proof is missing.": Exit Function
    ResolveReferenceFamilies inventoryNames, loadedVbaAnalysisNames, crosswalkNames, queryNames
    db.TableDefs.Refresh
    db.QueryDefs.Refresh
    currentToken = ReferenceTableFamilyStateToken(db, REFERENCE_INVENTORY_TOKEN_VERSION, inventoryNames, failureDetail)
    If Len(currentToken) = 0 Or StrComp(storedInventoryToken, currentToken, vbBinaryCompare) <> 0 Then detailText = "A4 Reference Inventory table family is stale. " & failureDetail: Exit Function
    matchedCount = matchedCount + 1
    failureDetail = vbNullString
    currentToken = ReferenceTableFamilyStateToken(db, REFERENCE_LOADED_VBA_ANALYSIS_TOKEN_VERSION, loadedVbaAnalysisNames, failureDetail)
    If Len(currentToken) = 0 Or StrComp(storedLoadedVbaAnalysisToken, currentToken, vbBinaryCompare) <> 0 Then detailText = "A4 Loaded VBA Analysis table family is stale. " & failureDetail: Exit Function
    matchedCount = matchedCount + 1
    failureDetail = vbNullString
    currentToken = ReferenceTableFamilyStateToken(db, REFERENCE_CROSSWALK_TOKEN_VERSION, crosswalkNames, failureDetail)
    If Len(currentToken) = 0 Or StrComp(storedCrosswalkToken, currentToken, vbBinaryCompare) <> 0 Then detailText = "A4 reference crosswalk family is stale. " & failureDetail: Exit Function
    matchedCount = matchedCount + 1
    failureDetail = vbNullString
    currentToken = ReferenceQueryFamilyStateToken(db, REFERENCE_QUERY_TOKEN_VERSION, queryNames, failureDetail)
    If Len(currentToken) = 0 Or StrComp(storedQueryToken, currentToken, vbBinaryCompare) <> 0 Then detailText = "A4 saved-query family is stale. " & failureDetail: Exit Function
    matchedCount = matchedCount + 1
    ReferenceInventoryFoundationCurrentForSetup = True
    detailText = "A4 reference/Add-In inventory foundation is current. Families matched=" & CStr(matchedCount) & "."
    Exit Function
NotCurrent:
    detailText = "A4 current-state verification failed. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
End Function

Private Sub RepairReferenceTableFamily(ByVal tableNames As Collection)
    Dim item As Variant
    For Each item In tableNames
        RepairReferenceTableByName CStr(item)
    Next item
End Sub

Private Sub RepairReferenceTableByName(ByVal tableName As String)
    Dim db As DAO.Database
    Select Case UCase$(tableName)
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActiveReferences"))
            EnsureActiveReferencesTable
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AvailableReferences"))
            EnsureAvailableReferencesTable
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects"))
            EnsureLoadedVBAProjectsTable
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_VBEAddIns")), UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_COMAddIns"))
            EnsureAddInInventoryTable tableName
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AccessAddIns"))
            EnsureAccessAddInInventoryTable
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjectReferences"))
            EnsureLoadedVBAProjectReferencesTable
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAModules"))
            EnsureLoadedVBAModulesTable
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProcedures"))
            EnsureLoadedVBAProceduresTable
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBADependencies"))
            EnsureLoadedVBADependenciesTable
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActiveReferenceProjectLinks"))
            EnsureActiveReferenceProjectLinksTable
        Case UCase$(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ReferenceManager"))
            EnsureReferenceManagerTable
        Case Else
            Err.Raise vbObjectError + 9727, "RepairReferenceTableByName", "The canonical Reference structural family contains a table with no repair implementation mapping." & vbCrLf & "Table: " & tableName
    End Select
    Set db = ADCDb()
    EnsureReferenceLookupFields db.TableDefs(tableName)
    If StrComp(tableName, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ReferenceManager"), vbTextCompare) <> 0 Then
        EnsureInventoryFreshnessIndex tableName
        InitializeInventoryFreshnessMetadata tableName
    End If
End Sub
' Purpose: Resolves reference table field order text from the supplied/current ADC state, using `SchemaDefinitionFieldOrderText` and `End` where required by the lookup path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   tableName (Input; String): Canonical or physical name identifying the table targeted by this operation.
' Returns: String: Returns the formatted text produced by this routine. `vbNullString` remains available where the existing routine uses an empty result to mean unavailable, not applicable, or not
'   resolved.
Public Function ReferenceTableFieldOrderText(ByVal tableName As String) As String
    ' Wrapper purpose: REFERENCES-owned facade returning canonical physical field order for a reference/Add-In inventory table from its authoritative schema definition.
    ' Why retained: Retained so DEFAULTS and other callers need not access private REFERENCES schema machinery or maintain separate field-order mirrors.
    'Intended use: Returns canonical physical field order for an ADC reference/Add-In inventory table directly from its authoritative schema definitions.
    ReferenceTableFieldOrderText = SchemaDefinitionFieldOrderText(ReferenceTableSchemaDefinitions(tableName))
End Function

Private Sub EnsureInventoryFreshnessIndex(ByVal tableName As String)
    'A4 physical-schema authority: current-state queries remain indexed even as retained stale history accumulates.
    Dim db As DAO.Database
    Set db = ADCDb()
    EnsureSingleFieldIndex db, tableName, Left$(INDEX_PREFIX_NONUNIQUE & tableName & INDEX_SUFFIX_INVENTORY_STATUS, 64), FIELD_INVENTORY_STATUS, False, False
End Sub

Private Function ReferenceScanBatchID(ByVal scannedOn As Date, ByVal inventoryName As String) As String
    'Operation-scoped consistency: groups rows written by the same retained-inventory scan without becoming a second current-state authority.
    ReferenceScanBatchID = Left$(Format$(scannedOn, "yyyymmdd_hhnnss") & "_" & UCase$(Trim$(inventoryName)), 100)
End Function

Private Sub InitializeInventoryFreshnessMetadata(ByVal tableName As String)
    'A4 schema migration only: rows from the last completed destructive snapshot remain current until the next live scan proves otherwise.
    Dim db As DAO.Database
    Set db = ADCDb()
    db.Execute "UPDATE [" & tableName & "] SET " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " WHERE Len(Nz(" & SQL_FIELD_INVENTORY_STATUS & ",''))=0;", dbFailOnError
End Sub

Private Sub PrepareInventoryRowsForScan(ByVal db As DAO.Database, ByVal tableName As String)
    'Retained-inventory lifecycle authority: a scan first makes the prior current partition stale; rediscovered identities are then updated back to Current in the same transaction.
    db.Execute "UPDATE [" & tableName & "] SET " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_STALE) & " WHERE Nz(" & SQL_FIELD_INVENTORY_STATUS & ",'')<>" & SharedDao_SqlQuote(INVENTORY_STATUS_STALE) & ";", dbFailOnError
End Sub

Private Function OpenInventoryUpsertRecordset(ByVal db As DAO.Database, ByVal tableName As String, ByVal identityIndexName As String) As DAO.Recordset
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset(tableName, dbOpenTable)
    rs.index = identityIndexName
    Set OpenInventoryUpsertRecordset = rs
End Function

Private Sub BeginInventoryUpsertRow(ByVal rs As DAO.Recordset, ByVal identityValue As String)
    rs.Seek "=", identityValue
    If rs.NoMatch Then rs.AddNew Else rs.Edit
End Sub

Private Function EnsureLoadedVBAModuleParentForChildWrites(ByVal moduleRows As DAO.Recordset, ByVal moduleIdentity As String, ByVal projectIdentity As String, ByVal projectName As String, ByVal projectRole As String, ByVal moduleOrder As Long, ByVal ModuleName As String, ByVal componentTypeCode As Long, ByVal scannedOn As Date) As Boolean
    'Source-scan ordering only: B5 RI requires first-seen module parents to exist before procedure/dependency children. Existing retained rows stay untouched so prior cache evidence remains authoritative.
    moduleRows.Seek "=", moduleIdentity
    If moduleRows.NoMatch = False Then Exit Function
    moduleRows.AddNew
    moduleRows.fields(FIELD_MODULE_IDENTITY).value = moduleIdentity
    moduleRows.fields(FIELD_PROJECT_IDENTITY).value = projectIdentity
    moduleRows.fields(FIELD_PROJECT_NAME).value = SharedValue_NormalizeAssignment(projectName)
    moduleRows.fields(FIELD_PROJECT_ROLE).value = projectRole
    moduleRows.fields(FIELD_MODULE_ORDER).value = moduleOrder
    moduleRows.fields(FIELD_MODULE_NAME).value = SharedValue_NormalizeAssignment(ModuleName)
    moduleRows.fields(FIELD_COMPONENT_TYPE_CODE).value = componentTypeCode
    moduleRows.fields(FIELD_COMPONENT_TYPE).value = VBComponentTypeText(componentTypeCode)
    moduleRows.fields(FIELD_INVENTORY_STATUS).value = INVENTORY_STATUS_CURRENT
    moduleRows.fields(FIELD_LAST_SCAN_BATCH_ID).value = ReferenceScanBatchID(scannedOn, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAModules"))
    moduleRows.fields(FIELD_LAST_SCANNED_ON).value = scannedOn
    moduleRows.Update
    EnsureLoadedVBAModuleParentForChildWrites = True
End Function

Private Sub RemoveReferenceIndexIfPresent(ByVal tdf As DAO.tableDef, ByVal indexName As String)
    On Error Resume Next
    tdf.Indexes.Delete indexName
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub EnsureReferenceManagerTable()
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim schemaDefinitions As Collection
    Set db = ADCDb()
    Set schemaDefinitions = ReferenceTableSchemaDefinitions(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ReferenceManager"))
    If DaoTableExists(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ReferenceManager")) = False Then
        Set tdf = db.CreateTableDef(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ReferenceManager"))
        AppendReferenceSchemaFields tdf, schemaDefinitions
        db.TableDefs.Append tdf
    End If
    Set tdf = db.TableDefs(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ReferenceManager"))
    If DAOFieldExists(tdf, FIELD_DIAGNOSTIC_TEXT) Then
        tdf.fields.Delete FIELD_DIAGNOSTIC_TEXT
        tdf.fields.Refresh
        CoreAddOperationNote "Removed obsolete field " & SQL_FIELD_DIAGNOSTIC_TEXT & " from canonical ADC-owned table [" & tdf.Name & "]."
    End If
    EnsureReferenceSchemaFields tdf, schemaDefinitions
    EnsureTableDefFieldOrderFromText tdf, SchemaDefinitionFieldOrderText(schemaDefinitions)
    EnsureSingleFieldIndex db, tdf.Name, INDEX_PK_REFERENCE_MANAGER, FIELD_ID, True, True
    EnsureSingleFieldIndex db, tdf.Name, INDEX_UX_REFERENCE_MANAGER_IDENTITY, FIELD_REFERENCE_IDENTITY, True, False
    EnsureSingleFieldIndex db, tdf.Name, INDEX_IX_REFERENCE_MANAGER_NAME, FIELD_REFERENCE_NAME, False, False
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "EnsureReferenceManagerTable", "Could not create or repair DBInv_VBARefs_Workspace." & vbCrLf & Err.Description
End Sub

Private Sub EnsureSavedQuery(ByVal queryName As String, ByVal sqlText As String)
    ' REFERENCES owns A4 query lifecycle/rollback. COMMON performs best-effort no-op comparison only; the authoritative write proof is the actual named QueryDef text accepted by ACE and then reread from this same
    ' query after refresh. Do not reparse persisted SQL through a second anonymous QueryDef.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim qdf As DAO.queryDef
    Dim priorSql As String
    Dim acceptedSql As String
    Dim existed As Boolean
    Dim sqlChanged As Boolean
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Set db = ADCDb()
    On Error Resume Next
    Set qdf = db.QueryDefs(queryName)
    existed = (Err.Number = 0)
    Err.Clear
    On Error GoTo ErrHandler
    If existed Then
        priorSql = qdf.sql
        sqlChanged = Not QuerySqlEquivalent(db, priorSql, sqlText)
        If sqlChanged Then qdf.sql = sqlText
    Else
        Set qdf = db.CreateQueryDef(queryName, sqlText)
        sqlChanged = True
    End If
    acceptedSql = qdf.sql
    Set qdf = Nothing
    db.QueryDefs.Refresh
    Set qdf = db.QueryDefs(queryName)
    If StrComp(FormattingInsensitiveQuerySqlKey(qdf.sql), FormattingInsensitiveQuerySqlKey(acceptedSql), vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 9701, "EnsureSavedQuery", "Saved query persisted SQL did not match the SQL accepted by ACE: " & queryName
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If existed And sqlChanged And Len(priorSql) > 0 Then db.QueryDefs(queryName).sql = priorSql
    If existed = False Then db.QueryDefs.Delete queryName
    On Error GoTo 0
    Err.Raise savedErrNumber, "EnsureSavedQuery", "Could not create or repair saved query [" & queryName & "]." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub EnsureField(ByVal tdf As DAO.tableDef, ByVal fieldName As String, ByVal fieldType As DAO.DataTypeEnum, Optional ByVal fieldSize As Long = 0, Optional ByVal autoNumber As Boolean = False)
    Dim fld As DAO.Field
    If DAOFieldExists(tdf, fieldName) = False Then
        If fieldType = dbText Then
            Set fld = tdf.CreateField(fieldName, fieldType, fieldSize)
        Else
            Set fld = tdf.CreateField(fieldName, fieldType)
        End If
        If autoNumber Then fld.Attributes = dbAutoIncrField
        tdf.fields.Append fld
        Exit Sub
    End If
    Set fld = tdf.fields(fieldName)
    If fld.Type <> fieldType Then Err.Raise vbObjectError + 9702, "EnsureField", "Reference inventory field type mismatch." & vbCrLf & "Table: " & tdf.Name & vbCrLf & "Field: " & fieldName
    If fieldType = dbText And fld.Size < fieldSize Then Err.Raise vbObjectError + 9703, "EnsureField", "Reference inventory text field is too short." & vbCrLf & "Table: " & tdf.Name & vbCrLf & "Field: " & fieldName & vbCrLf & "Required size: " & CStr(fieldSize) & vbCrLf & "Actual size: " & CStr(fld.Size)
    If autoNumber And (fld.Attributes And dbAutoIncrField) = 0 Then Err.Raise vbObjectError + 9704, "EnsureField", "Reference inventory ID field is not AutoNumber." & vbCrLf & "Table: " & tdf.Name
End Sub

Private Sub AppendAutoNumberField(ByVal tdf As DAO.tableDef, ByVal fieldName As String)
    Dim fld As DAO.Field
    Set fld = tdf.CreateField(fieldName, dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.fields.Append fld
End Sub

Private Sub AppendField(ByVal tdf As DAO.tableDef, ByVal fieldName As String, ByVal fieldType As DAO.DataTypeEnum, Optional ByVal fieldSize As Long = 0)
    If fieldType = dbText Then
        tdf.fields.Append tdf.CreateField(fieldName, fieldType, fieldSize)
    Else
        tdf.fields.Append tdf.CreateField(fieldName, fieldType)
    End If
End Sub

' =============================================================================================================================================================================
' REFERENCE / PROJECT INVENTORY SCAN PROCEDURES
' =============================================================================================================================================================================

' Purpose: Scans DB inventory active references for prepare reference inventory, rebuild project links.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
'   prepareReferenceInventory (Input Optional; Boolean): When True, enables or requests reference inventory for this call; when False, the routine follows its normal non-forced/default path. Default: True.
'   rebuildProjectLinks (Input Optional; Boolean): True rebuilds project links; False reuses the existing/current linkage when valid. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub ScanDBInventoryActiveReferences(Optional ByVal showMessage As Boolean = True, Optional ByVal prepareReferenceInventory As Boolean = True, Optional ByVal rebuildProjectLinks As Boolean = True)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim ws As DAO.workspace
    Dim rs As DAO.Recordset
    Dim referenceItem As Object
    Dim referenceIndex As Long
    Dim referenceCount As Long
    Dim scannedOn As Date
    Dim referenceIdentity As String
    Dim referenceName As String
    Dim referenceDescription As String
    Dim ReferenceGUID As String
    Dim fullPath As String
    Dim registrationScope As String
    Dim registrationView As String
    Dim platformText As String
    Dim registrationPath As String
    Dim registryLCID As String
    Dim registryKeyPath As String
    Dim typeLibFlags As String
    Dim helpDirectory As String
    Dim referenceKindLabel As String
    Dim scanStatus As String
    Dim fileNameText As String
    Dim folderPath As String
    Dim fileExtension As String
    Dim fileVersion As String
    Dim fileAttributes As String
    Dim majorVersion As Long
    Dim minorVersion As Long
    Dim referenceKindCode As Long
    Dim fileSizeBytes As Double
    Dim fileCreatedOn As Variant
    Dim fileModifiedOn As Variant
    Dim isBuiltIn As Boolean
    Dim isBroken As Boolean
    Dim fileExistsValue As Boolean
    Dim isRegistered As Boolean
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim operationTimer As Single
    Dim phaseTimer As Single
    Dim foundationElapsedSeconds As Double
    Dim preparationElapsedSeconds As Double
    Dim scanLoopElapsedSeconds As Double
    Dim finalizationElapsedSeconds As Double
    Dim timingText As String
    Dim completionDetail As String
    Dim foundationDetail As String
    operationTimer = Timer
    CorePulseProgress "Active references | preparing scan"
    PumpSetupResponsiveness True
    If SetupControlCancellationRequested() Then RequireSafeWorkContinue "ScanDBInventoryActiveReferences", "Scan progress boundary"
    If prepareReferenceInventory Then
        CorePulseProgress "Active references | verifying reference inventory foundation"
        phaseTimer = Timer
        If ReferenceInventoryFoundationCurrentForSetup(foundationDetail) = False Then
            Err.Raise vbObjectError + 9799, "ScanDBInventoryActiveReferences", "The A4 reference/Add-In inventory foundation is not current. Runtime operations do not repair setup-owned physical objects." & vbCrLf & foundationDetail
        End If
        foundationElapsedSeconds = ElapsedSecondsFromTimer(phaseTimer)
        CorePulseProgress "Active references | reference inventory foundation current | elapsed=" & FormattedElapsedSeconds(foundationElapsedSeconds)
        PumpSetupResponsiveness True
        If SetupControlCancellationRequested() Then RequireSafeWorkContinue "ScanDBInventoryActiveReferences", "Scan progress boundary"
    End If
    phaseTimer = Timer
    Set db = ADCDb()
    Set ws = DBEngine.Workspaces(0)
    scannedOn = Now
    referenceCount = Application.references.count
    CorePulseProgress "Active references | preparing retained inventory | total=" & CStr(referenceCount)
    ws.BeginTrans
    PrepareInventoryRowsForScan db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActiveReferences")
    Set rs = OpenInventoryUpsertRecordset(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActiveReferences"), INDEX_UX_ACTIVE_REFERENCES_IDENTITY)
    preparationElapsedSeconds = ElapsedSecondsFromTimer(phaseTimer)
    phaseTimer = Timer
    For referenceIndex = 1 To referenceCount
        PumpSetupResponsiveness
        If SetupControlCancellationRequested() Then RequireSafeWorkContinue "ScanDBInventoryActiveReferences", "Scan progress boundary"
        Set referenceItem = Application.references(referenceIndex)
        referenceName = SafeReferenceText(referenceItem, "Name")
        referenceDescription = SafeReferenceText(referenceItem, "Description")
        ReferenceGUID = NormalizeReferenceGuid(SafeReferenceText(referenceItem, "Guid"))
        majorVersion = SafeReferenceLong(referenceItem, "Major")
        minorVersion = SafeReferenceLong(referenceItem, "Minor")
        fullPath = SafeReferenceText(referenceItem, "FullPath")
        referenceKindCode = SafeReferenceLong(referenceItem, "Type")
        referenceKindLabel = ReferenceKindText(referenceKindCode)
        isBuiltIn = SafeReferenceBoolean(referenceItem, "BuiltIn", False)
        isBroken = SafeReferenceBoolean(referenceItem, "IsBroken", True)
        CorePulseProgress "Active references | " & CStr(referenceIndex) & "/" & CStr(referenceCount) & " | inspecting " & IIf(Len(referenceName) > 0, referenceName, "(unnamed reference)") & " | v" & CStr(majorVersion) & "." & CStr(minorVersion) & " | " & referenceKindLabel & " | broken=" & CStr(isBroken) & " | file=" & IIf(Len(fullPath) > 0, ReferenceFileBaseName(fullPath), "(none)")
        fileExistsValue = ReferenceFileExists(fullPath)
        CollectReferenceFileMetadata fullPath, fileNameText, folderPath, fileExtension, fileSizeBytes, fileCreatedOn, fileModifiedOn, fileVersion, fileAttributes
        registrationScope = vbNullString
        registrationView = vbNullString
        platformText = vbNullString
        registrationPath = vbNullString
        registryLCID = vbNullString
        registryKeyPath = vbNullString
        typeLibFlags = vbNullString
        helpDirectory = vbNullString
        isRegistered = FindReferenceRegistration(ReferenceGUID, majorVersion, minorVersion, fullPath, registrationScope, registrationView, platformText, registrationPath, registryLCID, registryKeyPath, typeLibFlags, helpDirectory)
        referenceIdentity = ActiveReferenceIdentity(ReferenceGUID, majorVersion, minorVersion, fullPath, referenceName)
        scanStatus = ActiveReferenceScanStatus(isBroken, fileExistsValue, isRegistered)
        PulseProgressEvery referenceIndex, 10, "Active references | " & CStr(referenceIndex) & "/" & CStr(referenceCount) & " | " & IIf(Len(referenceName) > 0, referenceName, "(unnamed reference)") & " | status=" & scanStatus & " | registered=" & CStr(isRegistered) & " | file exists=" & CStr(fileExistsValue), referenceCount, "ReferenceActiveScan", "References"
        BeginInventoryUpsertRow rs, referenceIdentity
        rs.fields(FIELD_REFERENCE_IDENTITY).value = referenceIdentity
        rs.fields(FIELD_REFERENCE_PRIORITY).value = referenceIndex
        rs.fields(FIELD_REFERENCE_NAME).value = SharedValue_NormalizeAssignment(referenceName)
        rs.fields(FIELD_REFERENCE_DESCRIPTION).value = SharedValue_NormalizeAssignment(referenceDescription)
        rs.fields(FIELD_REFERENCE_GUID).value = SharedValue_NormalizeAssignment(ReferenceGUID)
        rs.fields(FIELD_MAJOR_VERSION).value = majorVersion
        rs.fields(FIELD_MINOR_VERSION).value = minorVersion
        rs.fields(FIELD_REFERENCE_KIND_CODE).value = referenceKindCode
        rs.fields(FIELD_REFERENCE_KIND).value = SharedValue_NormalizeAssignment(referenceKindLabel)
        rs.fields(FIELD_FULL_PATH).value = SharedValue_NormalizeAssignment(fullPath)
        rs.fields(FIELD_FILE_NAME).value = SharedValue_NormalizeAssignment(fileNameText)
        rs.fields(FIELD_FOLDER_PATH).value = SharedValue_NormalizeAssignment(folderPath)
        rs.fields(FIELD_FILE_EXTENSION).value = SharedValue_NormalizeAssignment(fileExtension)
        rs.fields(FIELD_FILE_SIZE_BYTES).value = fileSizeBytes
        rs.fields(FIELD_FILE_CREATED_ON).value = fileCreatedOn
        rs.fields(FIELD_FILE_MODIFIED_ON).value = fileModifiedOn
        rs.fields(FIELD_FILE_VERSION).value = SharedValue_NormalizeAssignment(fileVersion)
        rs.fields(FIELD_FILE_ATTRIBUTES).value = SharedValue_NormalizeAssignment(fileAttributes)
        rs.fields(FIELD_IS_BUILT_IN).value = isBuiltIn
        rs.fields(FIELD_IS_BROKEN).value = isBroken
        rs.fields(FIELD_FILE_EXISTS).value = fileExistsValue
        rs.fields(FIELD_IS_REGISTERED).value = isRegistered
        rs.fields(FIELD_REGISTRATION_SCOPE).value = SharedValue_NormalizeAssignment(registrationScope)
        rs.fields(FIELD_REGISTRATION_VIEW).value = SharedValue_NormalizeAssignment(registrationView)
        rs.fields(FIELD_PLATFORM).value = SharedValue_NormalizeAssignment(platformText)
        rs.fields(FIELD_REGISTERED_PATH).value = SharedValue_NormalizeAssignment(registrationPath)
        rs.fields(FIELD_REGISTRY_LCID).value = SharedValue_NormalizeAssignment(registryLCID)
        rs.fields(FIELD_REGISTRY_KEY_PATH).value = SharedValue_NormalizeAssignment(registryKeyPath)
        rs.fields(FIELD_TYPE_LIB_FLAGS).value = SharedValue_NormalizeAssignment(typeLibFlags)
        rs.fields(FIELD_HELP_DIRECTORY).value = SharedValue_NormalizeAssignment(helpDirectory)
        rs.fields(FIELD_SCAN_STATUS).value = scanStatus
        rs.fields(FIELD_DIAGNOSTIC_TEXT).value = Null
        rs.fields(FIELD_DATABASE_NAME).value = DatabaseBaseNameFromPath(db.Name)
        rs.fields(FIELD_DATABASE_PATH).value = db.Name
        rs.fields(FIELD_INVENTORY_STATUS).value = INVENTORY_STATUS_CURRENT
        rs.fields(FIELD_LAST_SCAN_BATCH_ID).value = ReferenceScanBatchID(scannedOn, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActiveReferences"))
        rs.fields(FIELD_LAST_SCANNED_ON).value = scannedOn
        rs.Update
    Next referenceIndex
    scanLoopElapsedSeconds = ElapsedSecondsFromTimer(phaseTimer)
    phaseTimer = Timer
    rs.Close
    'Standalone scans rebuild the relationship immediately; combined parents that refresh DBInv_VBAPrjs_Ld next defer the rebuild until both authoritative inputs are current.
    If rebuildProjectLinks Then RebuildActiveReferenceProjectLinks db, scannedOn
    ws.CommitTrans
    finalizationElapsedSeconds = ElapsedSecondsFromTimer(phaseTimer)
    timingText = "Foundation verification: " & IIf(prepareReferenceInventory, FormattedElapsedSeconds(foundationElapsedSeconds), "reused by caller") & vbCrLf & _
                 "Scan preparation: " & FormattedElapsedSeconds(preparationElapsedSeconds) & vbCrLf & _
                 "Application.References scan loop: " & FormattedElapsedSeconds(scanLoopElapsedSeconds) & vbCrLf & _
                 "Finalization/crosswalk rebuild: " & IIf(rebuildProjectLinks, FormattedElapsedSeconds(finalizationElapsedSeconds), "deferred to caller") & vbCrLf & _
                 "Total: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(operationTimer))
    completionDetail = "Active references stored: " & CStr(referenceCount) & vbCrLf & IIf(rebuildProjectLinks, "Project-link crosswalk rebuilt from the current retained inventories.", "Project-link crosswalk rebuild deferred until the caller refreshes DBInv_VBAPrjs_Ld.") & vbCrLf & "Tables: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActiveReferences") & ", " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActiveReferenceProjectLinks") & vbCrLf & timingText
    LogTraceMessage "Info", "References", "ScanDBInventoryActiveReferences", "Active reference scan completed.", completionDetail, "Scan DB Inventory - Active References", 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActiveReferences"), "Table", "CodeDb", "References"
    If showMessage Then ReportCoreActionSuccess "ScanDBInventoryActiveReferences", "References", "Active reference scan completed successfully.", "Active references stored: " & CStr(referenceCount), "Scan DB Inventory - Active References", ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActiveReferences"), "Table", "CodeDb", "References", True, "Scan DB Inventory - Active References", vbInformation, True, "Tables", False
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not ws Is Nothing Then ws.Rollback
    On Error GoTo 0
    ReportReferenceFailure "ScanDBInventoryActiveReferences", "Could not scan active references. The prior current [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActiveReferences") & "] partition was preserved.", showMessage, savedErrNumber, savedErrSource, savedErrDescription
End Sub

' Purpose: Scans DB inventory loaded VBA projects for prepare reference inventory, correlate registered access add ins.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
'   prepareReferenceInventory (Input Optional; Boolean): When True, enables or requests reference inventory for this call; when False, the routine follows its normal non-forced/default path. Default: True.
'   correlateRegisteredAccessAddIns (Input Optional; Boolean): True performs registered access add ins correlation while building the result; False skips that optional correlation step. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub ScanDBInventoryLoadedVBAProjects(Optional ByVal showMessage As Boolean = True, Optional ByVal prepareReferenceInventory As Boolean = True, Optional ByVal correlateRegisteredAccessAddIns As Boolean = True)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim ws As DAO.workspace
    Dim rs As DAO.Recordset
    Dim vbProject As Object
    Dim activeProject As Object
    Dim projectIndex As Long
    Dim projectCount As Long
    Dim scannedOn As Date
    Dim projectIdentity As String
    Dim projectName As String
    Dim projectNameKey As String
    Dim projectFileName As String
    Dim projectFileBaseName As String
    Dim projectFileBaseNameKey As String
    Dim projectFilePath As String
    Dim projectPathKey As String
    Dim projectRole As String
    Dim protectionState As String
    Dim projectTypeText As String
    Dim scanStatus As String
    Dim diagnosticText As String
    Dim componentDiagnostic As String
    Dim componentCount As Long
    Dim protectionCode As Long
    Dim projectTypeCode As Long
    Dim componentsReadable As Boolean
    Dim sourceReadable As Boolean
    Dim isActiveProject As Boolean
    Dim fileExistsValue As Boolean
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim operationTimer As Single
    Dim phaseTimer As Single
    Dim foundationElapsedSeconds As Double
    Dim preparationElapsedSeconds As Double
    Dim scanLoopElapsedSeconds As Double
    Dim finalizationElapsedSeconds As Double
    Dim timingText As String
    Dim completionDetail As String
    Dim foundationDetail As String
    operationTimer = Timer
    CorePulseProgress "Loaded VBA projects | preparing scan"
    PumpSetupResponsiveness True
    If SetupControlCancellationRequested() Then RequireSafeWorkContinue "ScanDBInventoryLoadedVBAProjects", "Scan progress boundary"
    If prepareReferenceInventory Then
        CorePulseProgress "Loaded VBA projects | verifying reference inventory foundation"
        phaseTimer = Timer
        If ReferenceInventoryFoundationCurrentForSetup(foundationDetail) = False Then
            Err.Raise vbObjectError + 9799, "ScanDBInventoryLoadedVBAProjects", "The A4 reference/Add-In inventory foundation is not current. Runtime operations do not repair setup-owned physical objects." & vbCrLf & foundationDetail
        End If
        foundationElapsedSeconds = ElapsedSecondsFromTimer(phaseTimer)
        CorePulseProgress "Loaded VBA projects | reference inventory foundation current | elapsed=" & FormattedElapsedSeconds(foundationElapsedSeconds)
        PumpSetupResponsiveness True
        If SetupControlCancellationRequested() Then RequireSafeWorkContinue "ScanDBInventoryLoadedVBAProjects", "Scan progress boundary"
    End If
    phaseTimer = Timer
    Set db = ADCDb()
    Set ws = DBEngine.Workspaces(0)
    scannedOn = Now
    projectCount = Application.vbe.VBProjects.count
    CorePulseProgress "Loaded VBA projects | preparing retained inventory | total=" & CStr(projectCount)
    On Error Resume Next
    Set activeProject = Application.vbe.ActiveVBProject
    Err.Clear
    On Error GoTo ErrHandler
    ws.BeginTrans
    PrepareInventoryRowsForScan db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects")
    Set rs = OpenInventoryUpsertRecordset(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects"), INDEX_UX_LOADED_VBA_PROJECTS_IDENTITY)
    preparationElapsedSeconds = ElapsedSecondsFromTimer(phaseTimer)
    phaseTimer = Timer
    For projectIndex = 1 To projectCount
        PumpSetupResponsiveness
        If SetupControlCancellationRequested() Then RequireSafeWorkContinue "ScanDBInventoryLoadedVBAProjects", "Scan progress boundary"
        Set vbProject = Application.vbe.VBProjects(projectIndex)
        projectName = SafeVBProjectText(vbProject, "Name")
        projectNameKey = LoadedVBAProjectNameKey(projectName)
        projectFilePath = SafeVBProjectText(vbProject, "FileName")
        projectFileName = ReferenceFileBaseName(projectFilePath)
        projectFileBaseName = DatabaseBaseNameFromPath(projectFileName)
        projectFileBaseNameKey = LoadedVBAProjectNameKey(projectFileBaseName)
        projectPathKey = LoadedVBAProjectPathKey(projectFilePath)
        projectIdentity = LoadedVBAProjectIdentity(projectIndex, projectName, projectFilePath)
        projectRole = LoadedVBAProjectRole(projectFilePath)
        projectTypeCode = SafeVBProjectLong(vbProject, "Type", -1)
        projectTypeText = LoadedVBAProjectTypeText(projectTypeCode)
        protectionCode = SafeVBProjectLong(vbProject, "Protection", -1)
        protectionState = LoadedVBAProjectProtectionState(protectionCode)
        CorePulseProgress "Loaded VBA projects | " & CStr(projectIndex) & "/" & CStr(projectCount) & " | inspecting " & IIf(Len(projectName) > 0, projectName, "(unnamed project)") & " | role=" & projectRole & " | type=" & projectTypeText & " | protection=" & protectionState & " | file=" & IIf(Len(projectFileName) > 0, projectFileName, "(none)")
        componentDiagnostic = vbNullString
        componentsReadable = False
        componentCount = SafeVBProjectComponentCount(vbProject, componentsReadable, componentDiagnostic)
        sourceReadable = componentsReadable And protectionCode = 0
        isActiveProject = VBProjectObjectsMatch(vbProject, activeProject)
        fileExistsValue = ReferenceFileExists(projectFilePath)
        scanStatus = LoadedVBAProjectScanStatus(protectionCode, componentsReadable)
        PulseProgressEvery projectIndex, 10, "Loaded VBA projects | " & CStr(projectIndex) & "/" & CStr(projectCount) & " | " & IIf(Len(projectName) > 0, projectName, "(unnamed project)") & " | role=" & projectRole & " | components=" & CStr(componentCount) & " | source readable=" & CStr(sourceReadable) & " | status=" & scanStatus, projectCount, "ReferenceLoadedProjectScan", "References"
        diagnosticText = componentDiagnostic
        BeginInventoryUpsertRow rs, projectIdentity
        rs.fields(FIELD_PROJECT_IDENTITY).value = projectIdentity
        rs.fields(FIELD_PROJECT_ORDER).value = projectIndex
        rs.fields(FIELD_PROJECT_NAME).value = SharedValue_NormalizeAssignment(projectName)
        rs.fields(FIELD_PROJECT_NAME_KEY).value = SharedValue_NormalizeAssignment(projectNameKey)
        rs.fields(FIELD_PROJECT_FILE_NAME).value = SharedValue_NormalizeAssignment(projectFileName)
        rs.fields(FIELD_PROJECT_FILE_BASE_NAME).value = SharedValue_NormalizeAssignment(projectFileBaseName)
        rs.fields(FIELD_PROJECT_FILE_BASE_NAME_KEY).value = SharedValue_NormalizeAssignment(projectFileBaseNameKey)
        rs.fields(FIELD_PROJECT_FILE_PATH).value = SharedValue_NormalizeAssignment(projectFilePath)
        rs.fields(FIELD_PROJECT_PATH_KEY).value = SharedValue_NormalizeAssignment(projectPathKey)
        rs.fields(FIELD_PROJECT_ROLE).value = projectRole
        rs.fields(FIELD_IS_ACTIVE_VBE_PROJECT).value = isActiveProject
        rs.fields(FIELD_COMPONENT_COUNT).value = componentCount
        rs.fields(FIELD_COMPONENTS_READABLE).value = componentsReadable
        rs.fields(FIELD_SOURCE_READABLE).value = sourceReadable
        rs.fields(FIELD_PROTECTION_CODE).value = protectionCode
        rs.fields(FIELD_PROTECTION_STATE).value = protectionState
        rs.fields(FIELD_PROJECT_TYPE_CODE).value = projectTypeCode
        rs.fields(FIELD_PROJECT_TYPE).value = projectTypeText
        rs.fields(FIELD_FILE_EXISTS).value = fileExistsValue
        rs.fields(FIELD_SCAN_STATUS).value = scanStatus
        rs.fields(FIELD_DIAGNOSTIC_TEXT).value = SharedValue_NormalizeAssignment(diagnosticText)
        rs.fields(FIELD_INVENTORY_STATUS).value = INVENTORY_STATUS_CURRENT
        rs.fields(FIELD_LAST_SCAN_BATCH_ID).value = ReferenceScanBatchID(scannedOn, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects"))
        rs.fields(FIELD_LAST_SCANNED_ON).value = scannedOn
        rs.Update
    Next projectIndex
    scanLoopElapsedSeconds = ElapsedSecondsFromTimer(phaseTimer)
    phaseTimer = Timer
    rs.Close
    RebuildActiveReferenceProjectLinks db, scannedOn
    db.Execute "UPDATE [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects") & "] SET " & SQL_FIELD_SOURCE_ANALYSIS_STATUS & "='Not Scanned'," & SQL_FIELD_PROJECT_REFERENCE_COUNT & "=0," & SQL_FIELD_MODULE_INVENTORY_COUNT & "=0," & SQL_FIELD_PROCEDURE_INVENTORY_COUNT & "=0," & SQL_FIELD_DEPENDENCY_COUNT & "=0," & SQL_FIELD_LAST_SOURCE_SCANNED_ON & "=Null," & SQL_FIELD_SOURCE_ANALYSIS_DIAGNOSTIC_TEXT & "=Null WHERE " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " AND " & SQL_FIELD_SOURCE_ANALYSIS_STATUS & " Is Null;", dbFailOnError
    'Standalone Macro 61 completes registered Access add-in enrichment. Parent scans that are about to replace DBInv_AddIns_Access or loaded-project procedure evidence defer these correlations to their final
    'authoritative inventories.
    If correlateRegisteredAccessAddIns Then
        CorrelateTraditionalAccessAddInsWithLoadedProjects db
        CorrelateTraditionalAccessAddInEntryPoints db
    End If
    ws.CommitTrans
    finalizationElapsedSeconds = ElapsedSecondsFromTimer(phaseTimer)
    timingText = "Foundation verification: " & IIf(prepareReferenceInventory, FormattedElapsedSeconds(foundationElapsedSeconds), "reused by caller") & vbCrLf & _
                 "Scan preparation/retained inventory: " & FormattedElapsedSeconds(preparationElapsedSeconds) & vbCrLf & _
                 "VBE project enumeration loop: " & FormattedElapsedSeconds(scanLoopElapsedSeconds) & vbCrLf & _
                 "Finalization/crosswalk: " & FormattedElapsedSeconds(finalizationElapsedSeconds) & vbCrLf & _
                 "Total: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(operationTimer))
    completionDetail = "Loaded projects stored: " & CStr(projectCount) & vbCrLf & "The active-reference project crosswalk was rebuilt." & IIf(correlateRegisteredAccessAddIns, " Current registered Access add-ins were correlated with loaded projects and existing loaded-project entry-point evidence without controlling add-in membership.", " Registered Access add-in correlation was deferred because the parent scan will refresh a dependent authoritative inventory next.") & vbCrLf & "Tables: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects") & ", " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActiveReferenceProjectLinks") & vbCrLf & timingText
    LogTraceMessage "Info", "References", "ScanDBInventoryLoadedVBAProjects", "Loaded VBA project scan completed.", completionDetail, "Scan DB Inventory - Loaded VBA Projects", 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects"), "Table", "CodeDb", "References"
    If showMessage Then ReportCoreActionSuccess "ScanDBInventoryLoadedVBAProjects", "References", "Loaded VBA project scan completed successfully.", "Loaded projects stored: " & CStr(projectCount) & vbCrLf & "The active-reference project crosswalk was rebuilt without scanning loaded-project source. Registered Access add-ins were correlated only for runtime/load-state enrichment.", "Scan DB Inventory - Loaded VBA Projects", ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects"), "Table", "CodeDb", "References", True, "Scan DB Inventory - Loaded VBA Projects", vbInformation, True, "Tables", False
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not ws Is Nothing Then ws.Rollback
    On Error GoTo 0
    ReportReferenceFailure "ScanDBInventoryLoadedVBAProjects", "Could not scan loaded VBA projects. The prior current DBInv_VBAPrjs_Ld, project-link, and Access add-in correlation state was preserved.", showMessage, savedErrNumber, savedErrSource, savedErrDescription
End Sub

Private Sub UpdateLoadedVBAProjectSourceAnalysisSummary(ByVal db As DAO.Database, ByVal projectIdentity As String, ByVal scanStatus As String, ByVal moduleRosterSignature As String, ByVal referenceCount As Long, ByVal moduleCount As Long, ByVal procedureCount As Long, ByVal dependencyCount As Long, ByVal diagnosticText As String, ByVal scannedOn As Date)
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects") & "] WHERE " & SQL_FIELD_PROJECT_IDENTITY & "=" & SharedDao_SqlQuote(projectIdentity) & " AND " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & ";", dbOpenDynaset)
    If rs.EOF Then Err.Raise vbObjectError + 9744, "UpdateLoadedVBAProjectSourceAnalysisSummary", "DBInv_VBAPrjs_Ld does not contain the project being summarized: " & projectIdentity
    rs.Edit
    rs.fields(FIELD_SOURCE_ANALYSIS_STATUS).value = scanStatus
    rs.fields(FIELD_MODULE_ROSTER_SIGNATURE).value = SharedValue_NormalizeAssignment(moduleRosterSignature)
    rs.fields(FIELD_PROJECT_REFERENCE_COUNT).value = referenceCount
    rs.fields(FIELD_MODULE_INVENTORY_COUNT).value = moduleCount
    rs.fields(FIELD_PROCEDURE_INVENTORY_COUNT).value = procedureCount
    rs.fields(FIELD_DEPENDENCY_COUNT).value = dependencyCount
    rs.fields(FIELD_SOURCE_ANALYSIS_DIAGNOSTIC_TEXT).value = SharedValue_NormalizeAssignment(diagnosticText)
    rs.fields(FIELD_LAST_SOURCE_SCANNED_ON).value = scannedOn
    rs.Update
    rs.Close
End Sub

' Purpose: Scans DB inventory available references for prepare reference inventory.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
'   prepareReferenceInventory (Input Optional; Boolean): When True, enables or requests reference inventory for this call; when False, the routine follows its normal non-forced/default path. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub ScanDBInventoryAvailableReferences(Optional ByVal showMessage As Boolean = True, Optional ByVal prepareReferenceInventory As Boolean = True)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim ws As DAO.workspace
    Dim rs As DAO.Recordset
    Dim registryProvider As Object
    Dim activeKeys As Collection
    Dim activePaths As Collection
    Dim storedKeys As Collection
    Dim roots As Variant
    Dim rootInfo As Variant
    Dim guidKeys As Variant
    Dim guidItem As Variant
    Dim versionKeys As Variant
    Dim versionItem As Variant
    Dim guidText As String
    Dim versionText As String
    Dim activeIdentity As String
    Dim majorVersion As Long
    Dim minorVersion As Long
    Dim storedCount As Long
    Dim inspectedGuidCount As Long
    Dim scannedOn As Date
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim operationTimer As Single
    Dim phaseTimer As Single
    Dim foundationElapsedSeconds As Double
    Dim preparationElapsedSeconds As Double
    Dim scanLoopElapsedSeconds As Double
    Dim finalizationElapsedSeconds As Double
    Dim timingText As String
    Dim completionDetail As String
    Dim foundationDetail As String
    operationTimer = Timer
    CorePulseProgress "Available references | preparing scan"
    PumpSetupResponsiveness True
    If SetupControlCancellationRequested() Then RequireSafeWorkContinue "ScanDBInventoryAvailableReferences", "Scan progress boundary"
    If prepareReferenceInventory Then
        CorePulseProgress "Available references | verifying reference inventory foundation"
        phaseTimer = Timer
        If ReferenceInventoryFoundationCurrentForSetup(foundationDetail) = False Then
            Err.Raise vbObjectError + 9799, "ScanDBInventoryAvailableReferences", "The A4 reference/Add-In inventory foundation is not current. Runtime operations do not repair setup-owned physical objects." & vbCrLf & foundationDetail
        End If
        foundationElapsedSeconds = ElapsedSecondsFromTimer(phaseTimer)
        CorePulseProgress "Available references | reference inventory foundation current | elapsed=" & FormattedElapsedSeconds(foundationElapsedSeconds)
        PumpSetupResponsiveness True
        If SetupControlCancellationRequested() Then RequireSafeWorkContinue "ScanDBInventoryAvailableReferences", "Scan progress boundary"
    End If
    phaseTimer = Timer
    Set db = ADCDb()
    Set ws = DBEngine.Workspaces(0)
    CorePulseProgress "Available references | opening registry provider"
    PumpSetupResponsiveness True
    If SetupControlCancellationRequested() Then RequireSafeWorkContinue "ScanDBInventoryAvailableReferences", "Scan progress boundary"
    Set registryProvider = ReferenceRegistryProvider()
    CorePulseProgress "Available references | registry provider ready | loading active-reference identities"
    ResolveActiveReferenceKeys activeKeys, activePaths
    ' INTENTIONAL KEYED LOOKUP: storedKeys provides keyed membership/lookup over source rows or state already read by ScanDBInventoryAvailableReferences; it is a derived access path, not a second authority.
    Set storedKeys = New Collection
    roots = ReferenceRegistryRoots()
    scannedOn = Now
    CorePulseProgress "Available references | preparing registry scan | roots=" & CStr(UBound(roots) - LBound(roots) + 1) & " | active reference identities=" & CStr(activeKeys.count)
    ws.BeginTrans
    PrepareInventoryRowsForScan db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AvailableReferences")
    Set rs = OpenInventoryUpsertRecordset(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AvailableReferences"), INDEX_UX_AVAILABLE_REFERENCES_IDENTITY)
    preparationElapsedSeconds = ElapsedSecondsFromTimer(phaseTimer)
    phaseTimer = Timer
    For Each rootInfo In roots
        PumpSetupResponsiveness
        If SetupControlCancellationRequested() Then RequireSafeWorkContinue "ScanDBInventoryAvailableReferences", "Scan progress boundary"
        CorePulseProgress "Available references | root=" & CStr(rootInfo(2)) & " | view=" & CStr(rootInfo(3)) & " | GUIDs inspected=" & CStr(inspectedGuidCount) & " | rows stored=" & CStr(storedCount)
        guidKeys = RegistrySubkeys(registryProvider, CLng(rootInfo(0)), CStr(rootInfo(1)))
        If IsArray(guidKeys) Then
            For Each guidItem In guidKeys
                PumpSetupResponsiveness
                If SetupControlCancellationRequested() Then RequireSafeWorkContinue "ScanDBInventoryAvailableReferences", "Scan progress boundary"
                guidText = NormalizeReferenceGuid(CStr(guidItem))
                inspectedGuidCount = inspectedGuidCount + 1
                PulseProgressEvery inspectedGuidCount, 100, "Available references | GUID " & CStr(inspectedGuidCount) & " | " & CStr(rootInfo(2)) & "/" & CStr(rootInfo(3)) & " | " & CStr(guidItem) & " | rows stored=" & CStr(storedCount), 0, "ReferenceAvailableRegistrationScan", "References"
                versionKeys = RegistrySubkeys(registryProvider, CLng(rootInfo(0)), CStr(rootInfo(1)) & "\" & CStr(guidItem))
                If IsArray(versionKeys) Then
                    For Each versionItem In versionKeys
                        PumpSetupResponsiveness
                        If SetupControlCancellationRequested() Then RequireSafeWorkContinue "ScanDBInventoryAvailableReferences", "Scan progress boundary"
                        versionText = Trim$(CStr(versionItem))
                        If ParseReferenceVersion(versionText, majorVersion, minorVersion) Then
                            activeIdentity = ReferenceVersionIdentity(guidText, majorVersion, minorVersion)
                            If CollectionHasKey(activeKeys, activeIdentity) = False Then
                                ScanAvailableReferenceVersionRegistration registryProvider, rootInfo, CStr(guidItem), guidText, versionText, majorVersion, minorVersion, _
                                    activePaths, storedKeys, rs, scannedOn, storedCount
                            End If
                        End If
                    Next versionItem
                End If
            Next guidItem
        End If
    Next rootInfo
    scanLoopElapsedSeconds = ElapsedSecondsFromTimer(phaseTimer)
    phaseTimer = Timer
    rs.Close
    ws.CommitTrans
    finalizationElapsedSeconds = ElapsedSecondsFromTimer(phaseTimer)
    timingText = "Foundation verification: " & IIf(prepareReferenceInventory, FormattedElapsedSeconds(foundationElapsedSeconds), "reused by caller") & vbCrLf & _
                 "Registry/provider and active-reference preparation: " & FormattedElapsedSeconds(preparationElapsedSeconds) & vbCrLf & _
                 "Registry enumeration loop: " & FormattedElapsedSeconds(scanLoopElapsedSeconds) & vbCrLf & _
                 "Finalization/commit: " & FormattedElapsedSeconds(finalizationElapsedSeconds) & vbCrLf & _
                 "Total: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(operationTimer))
    completionDetail = "Registered type-library GUIDs checked: " & CStr(inspectedGuidCount) & vbCrLf & "Available non-active rows stored: " & CStr(storedCount) & vbCrLf & "Table: " & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AvailableReferences") & vbCrLf & timingText
    LogTraceMessage "Info", "References", "ScanDBInventoryAvailableReferences", "Available reference scan completed.", completionDetail, "Scan DB Inventory - Available References", 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AvailableReferences"), "Table", "CodeDb", "References"
    If showMessage Then ReportCoreActionSuccess "ScanDBInventoryAvailableReferences", "References", "Available reference scan completed successfully.", "Available non-active reference registrations stored: " & CStr(storedCount), "Scan DB Inventory - Available References", ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AvailableReferences"), "Table", "CodeDb", "References", True, "Scan DB Inventory - Available References", vbInformation, True, "Tables", False
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not ws Is Nothing Then ws.Rollback
    On Error GoTo 0
    ReportReferenceFailure "ScanDBInventoryAvailableReferences", "Could not scan available references. The prior current [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AvailableReferences") & "] partition was preserved.", showMessage, savedErrNumber, savedErrSource, savedErrDescription
End Sub

Private Sub ScanAvailableReferenceVersionRegistration( _
    ByVal registryProvider As Object, ByRef rootInfo As Variant, ByVal guidItem As String, ByVal guidText As String, _
    ByVal versionText As String, ByVal majorVersion As Long, ByVal minorVersion As Long, ByVal activePaths As Collection, _
    ByVal storedKeys As Collection, ByVal rs As DAO.Recordset, ByVal scannedOn As Date, ByRef storedCount As Long)
    'Processes the locale/platform subtree for one already-validated, non-active registered type-library version while the caller retains transaction and root/GUID/version orchestration.
    Dim localeKeys As Variant
    Dim localeItem As Variant
    Dim platformKeys As Variant
    Dim platformItem As Variant
    Dim localeText As String
    Dim localeDecimalText As String
    Dim platformText As String
    Dim versionPath As String
    Dim localePath As String
    Dim platformPath As String
    Dim registeredPath As String
    Dim descriptionText As String
    Dim referenceName As String
    Dim identityText As String
    Dim scanStatus As String
    Dim registryRootText As String
    Dim registryKeyPath As String
    Dim typeLibFlags As String
    Dim helpDirectory As String
    Dim fileNameText As String
    Dim folderPath As String
    Dim fileExtension As String
    Dim fileVersion As String
    Dim fileAttributes As String
    Dim fileSizeBytes As Double
    Dim fileCreatedOn As Variant
    Dim fileModifiedOn As Variant

    registryRootText = ReferenceRegistryRootDisplay(CLng(rootInfo(0)))
    versionPath = CStr(rootInfo(1)) & "\" & guidItem & "\" & versionText
    descriptionText = RegistryDefaultString(registryProvider, CLng(rootInfo(0)), versionPath)
    typeLibFlags = RegistryDefaultString(registryProvider, CLng(rootInfo(0)), versionPath & "\FLAGS")
    helpDirectory = ExpandReferencePath(RegistryDefaultString(registryProvider, CLng(rootInfo(0)), versionPath & "\HELPDIR"))
    localeKeys = RegistrySubkeys(registryProvider, CLng(rootInfo(0)), versionPath)
    If IsArray(localeKeys) = False Then Exit Sub

    For Each localeItem In localeKeys
        PumpSetupResponsiveness
        If SetupControlCancellationRequested() Then RequireSafeWorkContinue "ScanDBInventoryAvailableReferences", "Scan progress boundary"
        localeText = Trim$(CStr(localeItem))
        If IsReferenceLocaleKey(localeText) Then
            localeDecimalText = ReferenceLocaleDecimalText(localeText)
            localePath = versionPath & "\" & localeText
            platformKeys = RegistrySubkeys(registryProvider, CLng(rootInfo(0)), localePath)
            If IsArray(platformKeys) Then
                For Each platformItem In platformKeys
                    PumpSetupResponsiveness
                    If SetupControlCancellationRequested() Then RequireSafeWorkContinue "ScanDBInventoryAvailableReferences", "Scan progress boundary"
                    platformText = LCase$(Trim$(CStr(platformItem)))
                    If IsReferencePlatformKey(platformText) Then
                        platformPath = localePath & "\" & CStr(platformItem)
                        registeredPath = ExpandReferencePath(RegistryDefaultString(registryProvider, CLng(rootInfo(0)), platformPath))
                        registryKeyPath = registryRootText & "\" & platformPath
                        identityText = AvailableReferenceIdentity(guidText, majorVersion, minorVersion, localeText, platformText, CStr(rootInfo(2)), CStr(rootInfo(3)))
                        If CollectionHasKey(storedKeys, identityText) = False And CollectionHasKey(activePaths, ReferencePathIdentity(registeredPath)) = False Then
                            AddCollectionKey storedKeys, identityText
                            referenceName = descriptionText
                            If Len(Trim$(referenceName)) = 0 Then referenceName = ReferenceFileBaseName(registeredPath)
                            If Len(Trim$(referenceName)) = 0 Then referenceName = guidText
                            CollectReferenceFileMetadata registeredPath, fileNameText, folderPath, fileExtension, fileSizeBytes, fileCreatedOn, fileModifiedOn, fileVersion, fileAttributes
                            scanStatus = IIf(ReferenceFileExists(registeredPath), "Available", "Missing File")
                            BeginInventoryUpsertRow rs, identityText
                            rs.fields(FIELD_REFERENCE_IDENTITY).value = identityText
                            rs.fields(FIELD_REFERENCE_NAME).value = SharedValue_NormalizeAssignment(referenceName)
                            rs.fields(FIELD_REFERENCE_DESCRIPTION).value = SharedValue_NormalizeAssignment(descriptionText)
                            rs.fields(FIELD_REFERENCE_GUID).value = SharedValue_NormalizeAssignment(guidText)
                            rs.fields(FIELD_MAJOR_VERSION).value = majorVersion
                            rs.fields(FIELD_MINOR_VERSION).value = minorVersion
                            rs.fields(FIELD_VERSION_TEXT).value = versionText
                            rs.fields(FIELD_LCID).value = localeText
                            rs.fields(FIELD_LCID_DECIMAL).value = SharedValue_NormalizeAssignment(localeDecimalText)
                            rs.fields(FIELD_PLATFORM).value = platformText
                            rs.fields(FIELD_REGISTERED_PATH).value = SharedValue_NormalizeAssignment(registeredPath)
                            rs.fields(FIELD_FILE_NAME).value = SharedValue_NormalizeAssignment(fileNameText)
                            rs.fields(FIELD_FOLDER_PATH).value = SharedValue_NormalizeAssignment(folderPath)
                            rs.fields(FIELD_FILE_EXTENSION).value = SharedValue_NormalizeAssignment(fileExtension)
                            rs.fields(FIELD_FILE_SIZE_BYTES).value = fileSizeBytes
                            rs.fields(FIELD_FILE_CREATED_ON).value = fileCreatedOn
                            rs.fields(FIELD_FILE_MODIFIED_ON).value = fileModifiedOn
                            rs.fields(FIELD_FILE_VERSION).value = SharedValue_NormalizeAssignment(fileVersion)
                            rs.fields(FIELD_FILE_ATTRIBUTES).value = SharedValue_NormalizeAssignment(fileAttributes)
                            rs.fields(FIELD_FILE_EXISTS).value = ReferenceFileExists(registeredPath)
                            rs.fields(FIELD_REGISTRATION_SCOPE).value = CStr(rootInfo(2))
                            rs.fields(FIELD_REGISTRATION_VIEW).value = CStr(rootInfo(3))
                            rs.fields(FIELD_REGISTRY_ROOT).value = registryRootText
                            rs.fields(FIELD_REGISTRY_KEY_PATH).value = registryKeyPath
                            rs.fields(FIELD_TYPE_LIB_FLAGS).value = SharedValue_NormalizeAssignment(typeLibFlags)
                            rs.fields(FIELD_HELP_DIRECTORY).value = SharedValue_NormalizeAssignment(helpDirectory)
                            rs.fields(FIELD_SCAN_STATUS).value = scanStatus
                            rs.fields(FIELD_DIAGNOSTIC_TEXT).value = Null
                            rs.fields(FIELD_INVENTORY_STATUS).value = INVENTORY_STATUS_CURRENT
                            rs.fields(FIELD_LAST_SCAN_BATCH_ID).value = ReferenceScanBatchID(scannedOn, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AvailableReferences"))
                            rs.fields(FIELD_LAST_SCANNED_ON).value = scannedOn
                            rs.Update
                            storedCount = storedCount + 1
                        End If
                    End If
                Next platformItem
            End If
        End If
    Next localeItem
End Sub

Private Function ActiveReferenceIdentity(ByVal guidText As String, ByVal majorVersion As Long, ByVal minorVersion As Long, ByVal pathText As String, ByVal referenceName As String) As String
    ActiveReferenceIdentity = Left$("ACTIVE|" & ReferenceVersionIdentity(guidText, majorVersion, minorVersion) & "|" & ReferencePathIdentity(pathText) & "|" & UCase$(Trim$(referenceName)), 255)
End Function

Private Sub ResolveActiveReferenceKeys(ByRef identityKeys As Collection, ByRef pathKeys As Collection)
    Dim referenceItem As Object
    Dim referenceIndex As Long
    Dim guidText As String
    Dim majorVersion As Long
    Dim minorVersion As Long
    Dim fullPath As String
    ' INTENTIONAL OPERATION SNAPSHOT: identityKeys, pathKeys captures one coherent source state for this invocation of ResolveActiveReferenceKeys; validity ends with the operation rather than becoming a live data
    ' authority.
    Set identityKeys = New Collection
    Set pathKeys = New Collection
    For referenceIndex = 1 To Application.references.count
        PumpSetupResponsiveness
        Set referenceItem = Application.references(referenceIndex)
        guidText = NormalizeReferenceGuid(SafeReferenceText(referenceItem, "Guid"))
        majorVersion = SafeReferenceLong(referenceItem, "Major")
        minorVersion = SafeReferenceLong(referenceItem, "Minor")
        If Len(guidText) > 0 Then AddCollectionKey identityKeys, ReferenceVersionIdentity(guidText, majorVersion, minorVersion)
        fullPath = SafeReferenceText(referenceItem, "FullPath")
        If Len(Trim$(fullPath)) > 0 Then AddCollectionKey pathKeys, ReferencePathIdentity(fullPath)
    Next referenceIndex
End Sub

Private Sub RebuildActiveReferenceProjectLinks(ByVal db As DAO.Database, ByVal scannedOn As Date)
    Dim sourceRs As DAO.Recordset
    Dim targetRs As DAO.Recordset
    Dim referenceKindCode As Long
    Dim matchStatus As String
    Dim matchMethod As String
    Dim matchEvidence As String
    Dim loadedProjectIdentity As String
    Dim loadedProjectName As String
    Dim loadedProjectFilePath As String
    Dim loadedProjectRole As String
    Dim authoritativeMatch As Boolean
    PrepareInventoryRowsForScan db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActiveReferenceProjectLinks")
    Set sourceRs = db.OpenRecordset("SELECT " & SQL_FIELD_REFERENCE_IDENTITY & "," & SQL_FIELD_REFERENCE_PRIORITY & "," & SQL_FIELD_REFERENCE_NAME & "," & SQL_FIELD_REFERENCE_KIND_CODE & "," & SQL_FIELD_REFERENCE_KIND & "," & SQL_FIELD_FULL_PATH & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActiveReferences") & "] WHERE " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " ORDER BY " & SQL_FIELD_REFERENCE_PRIORITY & "," & SQL_FIELD_ID & ";", dbOpenSnapshot)
    Set targetRs = OpenInventoryUpsertRecordset(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActiveReferenceProjectLinks"), INDEX_UX_ACTIVE_REFERENCE_PROJECT_LINKS_REFERENCE)
    Do While sourceRs.EOF = False
        PumpSetupResponsiveness
        referenceKindCode = CLng(Nz(sourceRs.fields(FIELD_REFERENCE_KIND_CODE).value, -1))
        matchStatus = vbNullString
        matchMethod = vbNullString
        matchEvidence = vbNullString
        loadedProjectIdentity = vbNullString
        loadedProjectName = vbNullString
        loadedProjectFilePath = vbNullString
        loadedProjectRole = vbNullString
        authoritativeMatch = False
        If referenceKindCode = 1 Then
            ResolveLoadedVBAProjectMatch db, CStr(Nz(sourceRs.fields(FIELD_REFERENCE_NAME).value, vbNullString)), CStr(Nz(sourceRs.fields(FIELD_FULL_PATH).value, vbNullString)), matchStatus, matchMethod, matchEvidence, loadedProjectIdentity, loadedProjectName, loadedProjectFilePath, loadedProjectRole, authoritativeMatch
        Else
            matchStatus = "Not Applicable"
            matchMethod = "Reference kind"
            matchEvidence = vbNullString
        End If
        BeginInventoryUpsertRow targetRs, CStr(sourceRs.fields(FIELD_REFERENCE_IDENTITY).value)
        targetRs.fields(FIELD_REFERENCE_IDENTITY).value = sourceRs.fields(FIELD_REFERENCE_IDENTITY).value
        targetRs.fields(FIELD_REFERENCE_PRIORITY).value = sourceRs.fields(FIELD_REFERENCE_PRIORITY).value
        targetRs.fields(FIELD_REFERENCE_NAME).value = sourceRs.fields(FIELD_REFERENCE_NAME).value
        targetRs.fields(FIELD_REFERENCE_KIND).value = sourceRs.fields(FIELD_REFERENCE_KIND).value
        targetRs.fields(FIELD_REFERENCE_FULL_PATH).value = sourceRs.fields(FIELD_FULL_PATH).value
        targetRs.fields(FIELD_LOADED_PROJECT_IDENTITY).value = SharedValue_NormalizeAssignment(loadedProjectIdentity)
        targetRs.fields(FIELD_LOADED_PROJECT_NAME).value = SharedValue_NormalizeAssignment(loadedProjectName)
        targetRs.fields(FIELD_LOADED_PROJECT_FILE_PATH).value = SharedValue_NormalizeAssignment(loadedProjectFilePath)
        targetRs.fields(FIELD_LOADED_PROJECT_ROLE).value = SharedValue_NormalizeAssignment(loadedProjectRole)
        targetRs.fields(FIELD_MATCH_STATUS).value = matchStatus
        targetRs.fields(FIELD_MATCH_METHOD).value = matchMethod
        targetRs.fields(FIELD_MATCH_EVIDENCE).value = SharedValue_NormalizeAssignment(matchEvidence)
        targetRs.fields(FIELD_IS_AUTHORITATIVE_MATCH).value = authoritativeMatch
        targetRs.fields(FIELD_DIAGNOSTIC_TEXT).value = Null
        targetRs.fields(FIELD_INVENTORY_STATUS).value = INVENTORY_STATUS_CURRENT
        targetRs.fields(FIELD_LAST_SCAN_BATCH_ID).value = ReferenceScanBatchID(scannedOn, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActiveReferenceProjectLinks"))
        targetRs.fields(FIELD_LAST_SCANNED_ON).value = scannedOn
        targetRs.Update
        sourceRs.MoveNext
    Loop
    sourceRs.Close
    targetRs.Close
End Sub

Private Sub ResolveLoadedVBAProjectMatch(ByVal db As DAO.Database, ByVal referenceName As String, ByVal referenceFullPath As String, ByRef matchStatus As String, ByRef matchMethod As String, ByRef matchEvidence As String, ByRef loadedProjectIdentity As String, ByRef loadedProjectName As String, ByRef loadedProjectFilePath As String, ByRef loadedProjectRole As String, ByRef authoritativeMatch As Boolean)
    Dim rs As DAO.Recordset
    Dim referencePathKey As String
    Dim referenceNameKey As String
    Dim referenceFileBaseNameKey As String
    Dim exactMatchCount As Long
    Dim nameMatchCount As Long
    Dim loadedCount As Long
    Dim projectPathKey As String
    Dim projectNameKey As String
    Dim projectFileBaseNameKey As String
    Dim exactIdentity As String
    Dim exactName As String
    Dim exactPath As String
    Dim exactRole As String
    Dim nameIdentity As String
    Dim nameValue As String
    Dim namePath As String
    Dim nameRole As String
    referencePathKey = NormalizeReferencePath(referenceFullPath)
    referenceNameKey = LoadedVBAProjectNameKey(referenceName)
    referenceFileBaseNameKey = LoadedVBAProjectNameKey(DatabaseBaseNameFromPath(referenceFullPath))
    Set rs = db.OpenRecordset("SELECT " & SQL_FIELD_PROJECT_IDENTITY & "," & SQL_FIELD_PROJECT_NAME & "," & SQL_FIELD_PROJECT_FILE_PATH & "," & SQL_FIELD_PROJECT_ROLE & "," & SQL_FIELD_PROJECT_NAME_KEY & "," & SQL_FIELD_PROJECT_FILE_BASE_NAME_KEY & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects") & "] WHERE " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " ORDER BY " & SQL_FIELD_PROJECT_ORDER & "," & SQL_FIELD_ID & ";", dbOpenSnapshot)
    Do While rs.EOF = False
        PumpSetupResponsiveness
        loadedCount = loadedCount + 1
        projectPathKey = NormalizeReferencePath(CStr(Nz(rs.fields(FIELD_PROJECT_FILE_PATH).value, vbNullString)))
        projectNameKey = CStr(Nz(rs.fields(FIELD_PROJECT_NAME_KEY).value, vbNullString))
        projectFileBaseNameKey = CStr(Nz(rs.fields(FIELD_PROJECT_FILE_BASE_NAME_KEY).value, vbNullString))
        If Len(referencePathKey) > 0 And StrComp(referencePathKey, projectPathKey, vbBinaryCompare) = 0 Then
            exactMatchCount = exactMatchCount + 1
            If exactMatchCount = 1 Then
                exactIdentity = CStr(Nz(rs.fields(FIELD_PROJECT_IDENTITY).value, vbNullString))
                exactName = CStr(Nz(rs.fields(FIELD_PROJECT_NAME).value, vbNullString))
                exactPath = CStr(Nz(rs.fields(FIELD_PROJECT_FILE_PATH).value, vbNullString))
                exactRole = CStr(Nz(rs.fields(FIELD_PROJECT_ROLE).value, vbNullString))
            End If
        ElseIf LoadedVBAProjectNameMatches(referenceNameKey, referenceFileBaseNameKey, projectNameKey, projectFileBaseNameKey) Then
            nameMatchCount = nameMatchCount + 1
            If nameMatchCount = 1 Then
                nameIdentity = CStr(Nz(rs.fields(FIELD_PROJECT_IDENTITY).value, vbNullString))
                nameValue = CStr(Nz(rs.fields(FIELD_PROJECT_NAME).value, vbNullString))
                namePath = CStr(Nz(rs.fields(FIELD_PROJECT_FILE_PATH).value, vbNullString))
                nameRole = CStr(Nz(rs.fields(FIELD_PROJECT_ROLE).value, vbNullString))
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    If exactMatchCount = 1 Then
        matchStatus = "Resolved - Exact Path"
        matchMethod = "Exact normalized file path"
        matchEvidence = vbNullString
        loadedProjectIdentity = exactIdentity
        loadedProjectName = exactName
        loadedProjectFilePath = exactPath
        loadedProjectRole = exactRole
        authoritativeMatch = True
    ElseIf exactMatchCount > 1 Then
        matchStatus = "Ambiguous"
        matchMethod = "Exact normalized file path"
        matchEvidence = "Exact path matches: " & CStr(exactMatchCount)
    ElseIf nameMatchCount = 1 Then
        matchStatus = "Resolved - Name Only"
        matchMethod = "Unique project/file name"
        matchEvidence = vbNullString
        loadedProjectIdentity = nameIdentity
        loadedProjectName = nameValue
        loadedProjectFilePath = namePath
        loadedProjectRole = nameRole
    ElseIf nameMatchCount > 1 Then
        matchStatus = "Ambiguous"
        matchMethod = "Project/file name"
        matchEvidence = "Name matches: " & CStr(nameMatchCount)
    Else
        matchStatus = "Unresolved"
        matchMethod = "Path then name"
        matchEvidence = "Loaded projects examined: " & CStr(loadedCount)
    End If
End Sub

Private Function LoadedVBAProjectNameMatches(ByVal referenceNameKey As String, ByVal referenceFileBaseNameKey As String, ByVal projectNameKey As String, ByVal projectFileBaseNameKey As String) As Boolean
    If Len(referenceNameKey) > 0 Then
        If StrComp(referenceNameKey, projectNameKey, vbBinaryCompare) = 0 Or StrComp(referenceNameKey, projectFileBaseNameKey, vbBinaryCompare) = 0 Then LoadedVBAProjectNameMatches = True: Exit Function
    End If
    If Len(referenceFileBaseNameKey) > 0 Then
        If StrComp(referenceFileBaseNameKey, projectNameKey, vbBinaryCompare) = 0 Or StrComp(referenceFileBaseNameKey, projectFileBaseNameKey, vbBinaryCompare) = 0 Then LoadedVBAProjectNameMatches = True
    End If
End Function

Private Function LoadedVBAProjectIdentity(ByVal projectOrder As Long, ByVal projectName As String, ByVal projectFilePath As String) As String
    Dim normalizedPath As String
    normalizedPath = NormalizeReferencePath(projectFilePath)
    If Len(normalizedPath) > 0 Then
        LoadedVBAProjectIdentity = "P|" & LoadedVBAProjectPathFingerprint(normalizedPath)
    Else
        LoadedVBAProjectIdentity = Left$("N|" & LoadedVBAProjectNameKey(projectName) & "|" & CStr(projectOrder), 255)
    End If
End Function

Private Function LoadedVBAProjectPathFingerprint(ByVal normalizedPath As String) As String
    Dim indexValue As Long
    Dim sourceLength As Long
    Dim byteOffset As Long
    Dim reverseByteOffset As Long
    Dim sourceBytes() As Byte
    Dim forwardHash As Double
    Dim reverseHash As Double
    Dim forwardCharacter As Long
    Dim reverseCharacter As Long
    sourceLength = Len(normalizedPath)
    forwardHash = 5381#
    reverseHash = 5381#
    If sourceLength = 0 Then LoadedVBAProjectPathFingerprint = "0000150500001505": Exit Function
    sourceBytes = normalizedPath
    For indexValue = 0 To sourceLength - 1
        If (indexValue And 255) = 0 Then PumpSetupResponsiveness
        byteOffset = indexValue * 2
        forwardCharacter = CLng(sourceBytes(byteOffset)) + CLng(sourceBytes(byteOffset + 1)) * 256&
        forwardHash = forwardHash * 33# + forwardCharacter
        forwardHash = forwardHash - Int(forwardHash / 2147483647#) * 2147483647#
        reverseByteOffset = (sourceLength - indexValue - 1) * 2
        reverseCharacter = CLng(sourceBytes(reverseByteOffset)) + CLng(sourceBytes(reverseByteOffset + 1)) * 256&
        reverseHash = reverseHash * 33# + reverseCharacter
        reverseHash = reverseHash - Int(reverseHash / 2147483647#) * 2147483647#
    Next indexValue
    LoadedVBAProjectPathFingerprint = Right$("00000000" & Hex$(CLng(forwardHash)), 8) & Right$("00000000" & Hex$(CLng(reverseHash)), 8)
End Function

Private Function LoadedVBAProjectPathKey(ByVal pathText As String) As String
    LoadedVBAProjectPathKey = Left$(NormalizeReferencePath(pathText), 255)
End Function

Private Function LoadedVBAProjectNameKey(ByVal nameText As String) As String
    LoadedVBAProjectNameKey = Left$(LCase$(Trim$(nameText)), 255)
End Function

Private Function LoadedVBAProjectRole(ByVal projectFilePath As String) As String
    Dim codeDbPath As String
    Dim hostDbPath As String
    Dim projectPathKey As String
    On Error Resume Next
    codeDbPath = ADCDb().Name
    hostDbPath = GetHostDb().Name
    Err.Clear
    On Error GoTo 0
    projectPathKey = NormalizeReferencePath(projectFilePath)
    If Len(projectPathKey) = 0 Then
        LoadedVBAProjectRole = "Unknown"
    ElseIf StrComp(projectPathKey, NormalizeReferencePath(codeDbPath), vbBinaryCompare) = 0 Then
        LoadedVBAProjectRole = "Current CodeDb"
    ElseIf StrComp(projectPathKey, NormalizeReferencePath(hostDbPath), vbBinaryCompare) = 0 Then
        LoadedVBAProjectRole = "HostDb"
    Else
        LoadedVBAProjectRole = "External"
    End If
End Function

Private Function LoadedVBAProjectTypeText(ByVal projectTypeCode As Long) As String
    Select Case projectTypeCode
        Case 100: LoadedVBAProjectTypeText = "Host Project"
        Case 101: LoadedVBAProjectTypeText = "Standalone Project"
        Case Else: LoadedVBAProjectTypeText = "Unknown (" & CStr(projectTypeCode) & ")"
    End Select
End Function

Private Function LoadedVBAProjectProtectionState(ByVal protectionCode As Long) As String
    Select Case protectionCode
        Case 0: LoadedVBAProjectProtectionState = "Unlocked"
        Case 1: LoadedVBAProjectProtectionState = "Locked"
        Case Else: LoadedVBAProjectProtectionState = "Unknown"
    End Select
End Function

Private Function LoadedVBAProjectScanStatus(ByVal protectionCode As Long, ByVal componentsReadable As Boolean) As String
    If protectionCode = 1 Then
        LoadedVBAProjectScanStatus = "Locked"
    ElseIf componentsReadable = False Or protectionCode < 0 Then
        LoadedVBAProjectScanStatus = "Metadata Limited"
    Else
        LoadedVBAProjectScanStatus = "Loaded"
    End If
End Function

Private Function AvailableReferenceIdentity(ByVal guidText As String, ByVal majorVersion As Long, ByVal minorVersion As Long, ByVal localeText As String, ByVal platformText As String, ByVal scopeText As String, ByVal viewText As String) As String
    AvailableReferenceIdentity = Left$(ReferenceVersionIdentity(guidText, majorVersion, minorVersion) & "|" & UCase$(localeText) & "|" & UCase$(platformText) & "|" & UCase$(scopeText) & "|" & UCase$(viewText), 255)
End Function

Private Function ActiveReferenceScanStatus(ByVal isBroken As Boolean, ByVal fileExistsValue As Boolean, ByVal isRegistered As Boolean) As String
    If isBroken Then
        ActiveReferenceScanStatus = "Broken"
    ElseIf fileExistsValue = False Then
        ActiveReferenceScanStatus = "Missing File"
    ElseIf isRegistered = False Then
        ActiveReferenceScanStatus = "Active Unregistered"
    Else
        ActiveReferenceScanStatus = "Active"
    End If
End Function

Private Function FindActiveReferenceForManagerRecord(ByVal rs As DAO.Recordset) As Object
    Dim referenceItem As Object
    Dim referenceIndex As Long
    Dim targetGuid As String
    Dim targetPath As String
    Dim targetName As String
    Dim targetMajor As Long
    Dim targetMinor As Long
    targetGuid = NormalizeReferenceGuid(CStr(Nz(rs.fields(FIELD_REFERENCE_GUID).value, vbNullString)))
    targetPath = NormalizeReferencePath(CStr(Nz(rs.fields(FIELD_FULL_PATH).value, vbNullString)))
    targetName = Trim$(CStr(Nz(rs.fields(FIELD_REFERENCE_NAME).value, vbNullString)))
    targetMajor = CLng(Nz(rs.fields(FIELD_MAJOR_VERSION).value, 0))
    targetMinor = CLng(Nz(rs.fields(FIELD_MINOR_VERSION).value, 0))
    For referenceIndex = 1 To Application.references.count
        PumpSetupResponsiveness
        Set referenceItem = Application.references(referenceIndex)
        If Len(targetGuid) > 0 Then
            If StrComp(NormalizeReferenceGuid(SafeReferenceText(referenceItem, "Guid")), targetGuid, vbTextCompare) = 0 And SafeReferenceLong(referenceItem, "Major") = targetMajor And SafeReferenceLong(referenceItem, "Minor") = targetMinor Then Set FindActiveReferenceForManagerRecord = referenceItem: Exit Function
        ElseIf Len(targetPath) > 0 Then
            If StrComp(NormalizeReferencePath(SafeReferenceText(referenceItem, "FullPath")), targetPath, vbTextCompare) = 0 Then Set FindActiveReferenceForManagerRecord = referenceItem: Exit Function
        ElseIf Len(targetName) > 0 Then
            If StrComp(SafeReferenceText(referenceItem, "Name"), targetName, vbTextCompare) = 0 Then Set FindActiveReferenceForManagerRecord = referenceItem: Exit Function
        End If
    Next referenceIndex
End Function

Private Function FindActiveReferenceByGuid(ByVal ReferenceGUID As String) As Object
    Dim referenceItem As Object
    Dim referenceIndex As Long
    ReferenceGUID = NormalizeReferenceGuid(ReferenceGUID)
    If Len(ReferenceGUID) = 0 Then Exit Function
    For referenceIndex = 1 To Application.references.count
        PumpSetupResponsiveness
        Set referenceItem = Application.references(referenceIndex)
        If StrComp(NormalizeReferenceGuid(SafeReferenceText(referenceItem, "Guid")), ReferenceGUID, vbTextCompare) = 0 Then Set FindActiveReferenceByGuid = referenceItem: Exit Function
    Next referenceIndex
End Function

Private Function ActiveReferenceObjectKey(ByVal referenceItem As Object) As String
    Dim guidText As String
    Dim pathText As String
    guidText = NormalizeReferenceGuid(SafeReferenceText(referenceItem, "Guid"))
    If Len(guidText) > 0 Then
        ActiveReferenceObjectKey = ReferenceVersionIdentity(guidText, SafeReferenceLong(referenceItem, "Major"), SafeReferenceLong(referenceItem, "Minor"))
        Exit Function
    End If
    pathText = SafeReferenceText(referenceItem, "FullPath")
    If Len(NormalizeReferencePath(pathText)) > 0 Then
        ActiveReferenceObjectKey = ReferencePathIdentity(pathText)
    Else
        ActiveReferenceObjectKey = "N|" & UCase$(Trim$(SafeReferenceText(referenceItem, "Name")))
    End If
End Function

Private Function AddInDisplayName(ByVal sourceKind As String, ByVal addInOrder As Long, ByVal descriptionText As String, ByVal progIDText As String, ByVal guidText As String) As String
    If Len(Trim$(descriptionText)) > 0 Then
        AddInDisplayName = Trim$(descriptionText)
    ElseIf Len(Trim$(progIDText)) > 0 Then
        AddInDisplayName = Trim$(progIDText)
    ElseIf Len(Trim$(guidText)) > 0 Then
        AddInDisplayName = Trim$(guidText)
    Else
        AddInDisplayName = sourceKind & " #" & CStr(addInOrder)
    End If
End Function

Private Function BuildAddInIdentity(ByVal sourceKind As String, ByVal addInOrder As Long, ByVal guidText As String, ByVal progIDText As String, ByVal displayName As String) As String
    Dim identityText As String
    If Len(Trim$(guidText)) > 0 Then
        identityText = "GUID|" & UCase$(Trim$(guidText))
    ElseIf Len(Trim$(progIDText)) > 0 Then
        identityText = "PROGID|" & UCase$(Trim$(progIDText))
    Else
        identityText = "NAME|" & UCase$(Trim$(displayName)) & "|ORDER|" & CStr(addInOrder)
    End If
    BuildAddInIdentity = Left$(UCase$(sourceKind) & "|" & identityText, 255)
End Function

Private Function AddInScanStatus(ByVal isConnected As Boolean, ByVal connectReadable As Boolean) As String
    If connectReadable = False Then
        AddInScanStatus = "Metadata Limited"
    ElseIf isConnected Then
        AddInScanStatus = "Connected"
    Else
        AddInScanStatus = "Registered"
    End If
End Function

' =============================================================================================================================================================================
' LOADED VBA SOURCE / DEPENDENCY ANALYSIS PROCEDURES
' =============================================================================================================================================================================

' Purpose: Scans DB inventory add ins and loaded VBA projects for prepare reference inventory, refresh loaded VBA projects.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
'   prepareReferenceInventory (Input Optional; Boolean): When True, enables or requests reference inventory for this call; when False, the routine follows its normal non-forced/default path. Default: True.
'   refreshLoadedVBAProjects (Input Optional; Boolean): When True, enables or requests loaded VBA projects for this call; when False, the routine follows its normal non-forced/default path. Default: True.
'   correlateEntryPoints (Input Optional; Boolean): True performs entry points correlation while building the result; False skips that optional correlation step. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub ScanDBInventoryAddInsAndLoadedVBAProjects(Optional ByVal showMessage As Boolean = True, Optional ByVal prepareReferenceInventory As Boolean = True, Optional ByVal refreshLoadedVBAProjects As Boolean = True, Optional ByVal correlateEntryPoints As Boolean = True)
    On Error GoTo ErrHandler
    Dim loadedCount As Long
    Dim accessAddInCount As Long
    Dim vbeCount As Long
    Dim comCount As Long
    Dim operationTimer As Single
    Dim phaseTimer As Single
    Dim foundationElapsedSeconds As Double
    Dim loadedProjectsElapsedSeconds As Double
    Dim accessAddInsElapsedSeconds As Double
    Dim vbeAddInsElapsedSeconds As Double
    Dim comAddInsElapsedSeconds As Double
    Dim finalizationElapsedSeconds As Double
    Dim timingText As String
    Dim completionDetail As String
    Dim foundationDetail As String
    operationTimer = Timer
    CorePulseProgress "Add-in scan | preparing scan"
    PumpSetupResponsiveness True
    If SetupControlCancellationRequested() Then RequireSafeWorkContinue "ScanDBInventoryAddInsAndLoadedVBAProjects", "Scan progress boundary"
    If prepareReferenceInventory Then
        CorePulseProgress "Add-in scan | verifying reference inventory foundation"
        phaseTimer = Timer
        If ReferenceInventoryFoundationCurrentForSetup(foundationDetail) = False Then
            Err.Raise vbObjectError + 9799, "ScanDBInventoryAddInsAndLoadedVBAProjects", "The A4 reference/Add-In inventory foundation is not current. Runtime operations do not repair setup-owned physical objects." & vbCrLf & foundationDetail
        End If
        foundationElapsedSeconds = ElapsedSecondsFromTimer(phaseTimer)
        CorePulseProgress "Add-in scan | reference inventory foundation current | elapsed=" & FormattedElapsedSeconds(foundationElapsedSeconds) & " | step 1/4: loaded VBA projects"
        PumpSetupResponsiveness True
        If SetupControlCancellationRequested() Then RequireSafeWorkContinue "ScanDBInventoryAddInsAndLoadedVBAProjects", "Scan progress boundary"
    Else
        CorePulseProgress "Add-in scan | reference inventory foundation reused from parent operation | step 1/4: loaded VBA projects"
    End If
    'Standalone Macro 77 refreshes the authoritative loaded-project inventory here but defers Access add-in correlation until the replacement Add-In inventories are current; a coordinated parent scan can reuse its
    'just-completed loaded-project phase.
    If refreshLoadedVBAProjects Then
        phaseTimer = Timer
        ScanDBInventoryLoadedVBAProjects False, False, False
        loadedProjectsElapsedSeconds = ElapsedSecondsFromTimer(phaseTimer)
    End If
    loadedCount = InventoryTableRowCount(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects"))
    CorePulseProgress "Add-in scan | step 1/4 complete | loaded VBA projects=" & CStr(loadedCount) & " | elapsed=" & IIf(refreshLoadedVBAProjects, FormattedElapsedSeconds(loadedProjectsElapsedSeconds), "reused by caller") & " | step 2/4: installed Access add-ins"
    phaseTimer = Timer
    accessAddInCount = ScanTraditionalAccessAddInsCore(correlateEntryPoints)
    accessAddInsElapsedSeconds = ElapsedSecondsFromTimer(phaseTimer)
    CorePulseProgress "Add-in scan | step 2/4 complete | installed Access registrations=" & CStr(accessAddInCount) & " | elapsed=" & FormattedElapsedSeconds(accessAddInsElapsedSeconds) & " | step 3/4: VBE add-ins"
    phaseTimer = Timer
    vbeCount = ScanVBEAddInsCore()
    vbeAddInsElapsedSeconds = ElapsedSecondsFromTimer(phaseTimer)
    CorePulseProgress "Add-in scan | step 3/4 complete | VBE add-ins=" & CStr(vbeCount) & " | elapsed=" & FormattedElapsedSeconds(vbeAddInsElapsedSeconds) & " | step 4/4: COM add-ins"
    phaseTimer = Timer
    comCount = ScanCOMAddInsCore()
    comAddInsElapsedSeconds = ElapsedSecondsFromTimer(phaseTimer)
    CorePulseProgress "Add-in scan | step 4/4 complete | loaded projects=" & CStr(loadedCount) & " | Access registrations=" & CStr(accessAddInCount) & " | VBE add-ins=" & CStr(vbeCount) & " | COM add-ins=" & CStr(comCount) & " | elapsed=" & FormattedElapsedSeconds(comAddInsElapsedSeconds)
    phaseTimer = Timer
    completionDetail = "Loaded VBA projects: " & CStr(loadedCount) & vbCrLf & "Installed traditional Access add-in registrations: " & CStr(accessAddInCount) & vbCrLf & "VBE add-ins: " & CStr(vbeCount) & vbCrLf & "COM add-ins: " & CStr(comCount) & vbCrLf & "No add-in connection state, registration, reference, expression, or source was changed or executed."
    finalizationElapsedSeconds = ElapsedSecondsFromTimer(phaseTimer)
    timingText = "Foundation verification: " & IIf(prepareReferenceInventory, FormattedElapsedSeconds(foundationElapsedSeconds), "reused by caller") & vbCrLf & _
                 "Loaded VBA projects: " & IIf(refreshLoadedVBAProjects, FormattedElapsedSeconds(loadedProjectsElapsedSeconds), "reused by caller") & vbCrLf & _
                 "Installed Access add-ins: " & FormattedElapsedSeconds(accessAddInsElapsedSeconds) & vbCrLf & _
                 "VBE add-ins: " & FormattedElapsedSeconds(vbeAddInsElapsedSeconds) & vbCrLf & _
                 "Phase 4 COM add-ins: " & FormattedElapsedSeconds(comAddInsElapsedSeconds) & vbCrLf & _
                 "Finalization: " & FormattedElapsedSeconds(finalizationElapsedSeconds) & vbCrLf & _
                 "Total: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(operationTimer))
    completionDetail = completionDetail & vbCrLf & timingText
    LogTraceMessage "Info", "AddIns", "ScanDBInventoryAddInsAndLoadedVBAProjects", "Passive add-in inventory completed.", completionDetail, "Scan DB Inventory - Add-Ins and Loaded VBA Projects", 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Inventory"), "Query", "CodeDb", "AddIns"
    If showMessage Then ReportCoreActionSuccess "ScanDBInventoryAddInsAndLoadedVBAProjects", "AddIns", "Passive add-in inventory completed successfully.", "Loaded VBA projects: " & CStr(loadedCount) & vbCrLf & "Installed traditional Access add-in registrations: " & CStr(accessAddInCount) & vbCrLf & "VBE add-ins: " & CStr(vbeCount) & vbCrLf & "COM add-ins: " & CStr(comCount) & vbCrLf & "The scan did not connect, disconnect, register, load, unload, or execute any add-in entry-point expression.", "Scan DB Inventory - Add-Ins and Loaded VBA Projects", ManagedObjectPhysicalName(NAMING_CATEGORY_QUERY, "ADC_AddIn_Inventory"), "Query", "CodeDb", "AddIns", True, "Scan DB Inventory - Add-Ins and Loaded VBA Projects", vbInformation, True, "Tables", False
    Exit Sub
ErrHandler:
    ReportReferenceFailure "ScanDBInventoryAddInsAndLoadedVBAProjects", "Could not complete the passive add-in inventory. Each inventory that completed before the failure retained its committed current/stale partition; the failing inventory preserved its prior current/stale partition.", showMessage, Err.Number, Err.Source, Err.Description
End Sub

' Purpose: Scans DB inventory loaded VBA source for prepare reference inventory, refresh loaded VBA projects.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
'   prepareReferenceInventory (Input Optional; Boolean): When True, enables or requests reference inventory for this call; when False, the routine follows its normal non-forced/default path. Default: True.
'   refreshLoadedVBAProjects (Input Optional; Boolean): When True, enables or requests loaded VBA projects for this call; when False, the routine follows its normal non-forced/default path. Default: True.
'   correlateAccessAddInsWithProjects (Input Optional; Boolean): True performs access add ins with projects correlation while building the result; False skips that optional correlation step. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub ScanDBInventoryLoadedVBASource(Optional ByVal showMessage As Boolean = True, Optional ByVal prepareReferenceInventory As Boolean = True, Optional ByVal refreshLoadedVBAProjects As Boolean = True, Optional ByVal correlateAccessAddInsWithProjects As Boolean = True)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim ws As DAO.workspace
    Dim referenceRows As DAO.Recordset
    Dim moduleRows As DAO.Recordset
    Dim procedureRows As DAO.Recordset
    Dim dependencyRows As DAO.Recordset
    Dim vbProject As Object
    Dim projectIndex As Long
    Dim projectCount As Long
    Dim scannedProjectCount As Long
    Dim currentCodeDbProjectCount As Long
    Dim hostDbProjectCount As Long
    Dim externalProjectCount As Long
    Dim referenceCount As Long
    Dim moduleCount As Long
    Dim procedureCount As Long
    Dim dependencyCount As Long
    Dim modulesReusedByChatGPTID As Long
    Dim modulesReusedByFingerprint As Long
    Dim modulesSourceRescanned As Long
    Dim modulesRosterReconciled As Long
    Dim proceduresRetained As Long
    Dim proceduresRebuilt As Long
    Dim projectReferenceCount As Long
    Dim projectModuleCount As Long
    Dim projectProcedureCount As Long
    Dim projectDependencyCount As Long
    Dim projectIdentity As String
    Dim projectName As String
    Dim projectFilePath As String
    Dim projectRole As String
    Dim projectStatus As String
    Dim projectDiagnostic As String
    Dim projectModuleRosterSignature As String
    Dim scannedOn As Date
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim transactionActive As Boolean
    Dim operationTimer As Single
    Dim phaseTimer As Single
    Dim foundationElapsedSeconds As Double
    Dim loadedProjectsElapsedSeconds As Double
    Dim sourcePreparationElapsedSeconds As Double
    Dim sourceScanElapsedSeconds As Double
    Dim finalizationElapsedSeconds As Double
    Dim timingText As String
    Dim completionDetail As String
    Dim foundationDetail As String
    operationTimer = Timer
    CorePulseProgress "Loaded VBA source | preparing scan"
    PumpSetupResponsiveness True
    If SetupControlCancellationRequested() Then RequireSafeWorkContinue "ScanDBInventoryLoadedVBASource", "Scan progress boundary"
    If prepareReferenceInventory Then
        CorePulseProgress "Loaded VBA source | verifying reference inventory foundation"
        phaseTimer = Timer
        If ReferenceInventoryFoundationCurrentForSetup(foundationDetail) = False Then
            Err.Raise vbObjectError + 9799, "ScanDBInventoryLoadedVBASource", "The A4 reference/Add-In inventory foundation is not current. Runtime operations do not repair setup-owned physical objects." & vbCrLf & foundationDetail
        End If
        foundationElapsedSeconds = ElapsedSecondsFromTimer(phaseTimer)
        CorePulseProgress "Loaded VBA source | reference inventory foundation current | elapsed=" & FormattedElapsedSeconds(foundationElapsedSeconds) & " | Loaded VBA projects"
        PumpSetupResponsiveness True
        If SetupControlCancellationRequested() Then RequireSafeWorkContinue "ScanDBInventoryLoadedVBASource", "Scan progress boundary"
    Else
        CorePulseProgress "Loaded VBA source | reference inventory foundation reused from parent operation | Loaded VBA projects"
    End If
    Set db = ADCDb()
    'Standalone Macro 81 refreshes the authoritative loaded-project inventory here but defers Access add-in correlation until the all-loaded-project procedure inventory is current; a coordinated parent scan can
    'reuse its just-completed loaded-project phase.
    If refreshLoadedVBAProjects Then
        phaseTimer = Timer
        ScanDBInventoryLoadedVBAProjects False, False, False
        loadedProjectsElapsedSeconds = ElapsedSecondsFromTimer(phaseTimer)
    End If
    CorePulseProgress "Loaded VBA source | loaded VBA project inventory current | elapsed=" & IIf(refreshLoadedVBAProjects, FormattedElapsedSeconds(loadedProjectsElapsedSeconds), "reused by caller") & " | Source analysis"
    PumpSetupResponsiveness True
    If SetupControlCancellationRequested() Then RequireSafeWorkContinue "ScanDBInventoryLoadedVBASource", "Scan progress boundary"
    phaseTimer = Timer
    Set ws = DBEngine.Workspaces(0)
    scannedOn = Now
    projectCount = Application.vbe.VBProjects.count
    ws.BeginTrans
    transactionActive = True
    PrepareInventoryRowsForScan db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjectReferences")
    PrepareInventoryRowsForScan db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAModules")
    PrepareInventoryRowsForScan db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProcedures")
    PrepareInventoryRowsForScan db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBADependencies")
    db.Execute "UPDATE [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects") & "] SET " & SQL_FIELD_SOURCE_ANALYSIS_STATUS & "='Pending'," & SQL_FIELD_PROJECT_REFERENCE_COUNT & "=0," & SQL_FIELD_MODULE_INVENTORY_COUNT & "=0," & SQL_FIELD_PROCEDURE_INVENTORY_COUNT & "=0," & SQL_FIELD_DEPENDENCY_COUNT & "=0," & SQL_FIELD_LAST_SOURCE_SCANNED_ON & "=Null," & SQL_FIELD_SOURCE_ANALYSIS_DIAGNOSTIC_TEXT & "=Null WHERE " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & ";", dbFailOnError
    Set referenceRows = OpenInventoryUpsertRecordset(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjectReferences"), INDEX_UX_LOADED_VBA_PROJECT_REFERENCES_IDENTITY)
    Set moduleRows = OpenInventoryUpsertRecordset(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAModules"), INDEX_UX_LOADED_VBA_MODULES_IDENTITY)
    Set procedureRows = OpenInventoryUpsertRecordset(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProcedures"), INDEX_UX_LOADED_VBA_PROCEDURES_IDENTITY)
    Set dependencyRows = OpenInventoryUpsertRecordset(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBADependencies"), INDEX_UX_LOADED_VBA_DEPENDENCIES_IDENTITY)
    CorePulseProgress "Loaded VBA source | Source analysis | 0/" & CStr(projectCount) & " | projects inspected=0 | refs=0 | modules=0 | procedures=0 | dependencies=0"
    sourcePreparationElapsedSeconds = ElapsedSecondsFromTimer(phaseTimer)
    phaseTimer = Timer
    For projectIndex = 1 To projectCount
        PumpSetupResponsiveness
        If SetupControlCancellationRequested() Then RequireSafeWorkContinue "ScanDBInventoryLoadedVBASource", "Scan progress boundary"
        Set vbProject = Application.vbe.VBProjects(projectIndex)
        projectName = SafeVBProjectText(vbProject, "Name")
        projectFilePath = SafeVBProjectText(vbProject, "FileName")
        projectIdentity = LoadedVBAProjectIdentity(projectIndex, projectName, projectFilePath)
        projectRole = LoadedVBAProjectRole(projectFilePath)
        PulseProgressEvery projectIndex, 5, "Loaded VBA source | " & CStr(projectIndex) & "/" & CStr(projectCount) & " | " & IIf(Len(projectName) > 0, projectName, "(unnamed project)") & " | role=" & projectRole & " | totals refs=" & CStr(referenceCount) & " modules=" & CStr(moduleCount) & " procedures=" & CStr(procedureCount) & " dependencies=" & CStr(dependencyCount), projectCount, "ReferenceLoadedVBAProjectSourceScan", "References"
        projectReferenceCount = 0
        projectModuleCount = 0
        projectProcedureCount = 0
        projectDependencyCount = 0
        projectDiagnostic = vbNullString
        projectModuleRosterSignature = vbNullString
        ScanOneLoadedVBAProjectSource db, vbProject, projectIdentity, projectName, projectRole, scannedOn, referenceRows, moduleRows, procedureRows, dependencyRows, projectReferenceCount, projectModuleCount, projectProcedureCount, projectDependencyCount, projectStatus, projectDiagnostic, projectModuleRosterSignature, modulesReusedByChatGPTID, modulesReusedByFingerprint, modulesSourceRescanned, modulesRosterReconciled, proceduresRetained, proceduresRebuilt
        scannedProjectCount = scannedProjectCount + 1
        Select Case projectRole
            Case "Current CodeDb": currentCodeDbProjectCount = currentCodeDbProjectCount + 1
            Case "HostDb": hostDbProjectCount = hostDbProjectCount + 1
            Case "External": externalProjectCount = externalProjectCount + 1
        End Select
        referenceCount = referenceCount + projectReferenceCount
        moduleCount = moduleCount + projectModuleCount
        procedureCount = procedureCount + projectProcedureCount
        dependencyCount = dependencyCount + projectDependencyCount
        UpdateLoadedVBAProjectSourceAnalysisSummary db, projectIdentity, projectStatus, projectModuleRosterSignature, projectReferenceCount, projectModuleCount, projectProcedureCount, projectDependencyCount, projectDiagnostic, scannedOn
        CorePulseProgress "Loaded VBA source | " & CStr(projectIndex) & "/" & CStr(projectCount) & " complete | " & IIf(Len(projectName) > 0, projectName, "(unnamed project)") & " | status=" & projectStatus & " | project refs=" & CStr(projectReferenceCount) & " modules=" & CStr(projectModuleCount) & " procedures=" & CStr(projectProcedureCount) & " dependencies=" & CStr(projectDependencyCount) & " | totals refs=" & CStr(referenceCount) & " modules=" & CStr(moduleCount) & " procedures=" & CStr(procedureCount) & " dependencies=" & CStr(dependencyCount)
    Next projectIndex
    sourceScanElapsedSeconds = ElapsedSecondsFromTimer(phaseTimer)
    phaseTimer = Timer
    'Standalone Macro 81 finalizes registered Access add-in enrichment only after the all-loaded-project procedure inventory is current. Scan All DB Inventory can reuse the project correlation completed by its
    'immediately preceding Add-ins phase.
    If correlateAccessAddInsWithProjects Then CorrelateTraditionalAccessAddInsWithLoadedProjects db
    CorrelateTraditionalAccessAddInEntryPoints db
    referenceRows.Close
    moduleRows.Close
    procedureRows.Close
    dependencyRows.Close
    ws.CommitTrans
    transactionActive = False
    finalizationElapsedSeconds = ElapsedSecondsFromTimer(phaseTimer)
    timingText = "Foundation verification: " & IIf(prepareReferenceInventory, FormattedElapsedSeconds(foundationElapsedSeconds), "reused by caller") & vbCrLf & _
                 "Loaded VBA projects: " & IIf(refreshLoadedVBAProjects, FormattedElapsedSeconds(loadedProjectsElapsedSeconds), "reused by caller") & vbCrLf & _
                 "Source-analysis preparation: " & FormattedElapsedSeconds(sourcePreparationElapsedSeconds) & vbCrLf & _
                 "Project/source analysis loop: " & FormattedElapsedSeconds(sourceScanElapsedSeconds) & vbCrLf & _
                 "Finalization/commit: " & FormattedElapsedSeconds(finalizationElapsedSeconds) & vbCrLf & _
                 "Total: " & FormattedElapsedSeconds(ElapsedSecondsFromTimer(operationTimer))
    completionDetail = "Loaded projects considered: " & CStr(projectCount) & vbCrLf & _
                 "Projects inspected: " & CStr(scannedProjectCount) & vbCrLf & _
                 "Current CodeDb projects inspected: " & CStr(currentCodeDbProjectCount) & vbCrLf & _
                 "HostDb projects inspected: " & CStr(hostDbProjectCount) & vbCrLf & _
                 "External projects inspected: " & CStr(externalProjectCount) & vbCrLf & _
                 "Project references: " & CStr(referenceCount) & vbCrLf & _
                 "Modules examined: " & CStr(moduleCount) & vbCrLf & _
                 "Modules reused by ChatGPT ID: " & CStr(modulesReusedByChatGPTID) & vbCrLf & _
                 "Modules reused by fingerprint: " & CStr(modulesReusedByFingerprint) & vbCrLf & _
                 "Modules source-rescanned: " & CStr(modulesSourceRescanned) & vbCrLf & _
                 "Modules requiring roster-only dependency reconciliation: " & CStr(modulesRosterReconciled) & vbCrLf & _
                 "Procedures retained: " & CStr(proceduresRetained) & vbCrLf & _
                 "Procedures rebuilt: " & CStr(proceduresRebuilt) & vbCrLf & _
                 "Procedures current: " & CStr(procedureCount) & vbCrLf & _
                 "Dependencies current: " & CStr(dependencyCount) & vbCrLf & _
                 "No project, reference, module, or add-in state was changed." & vbCrLf & timingText
    LogTraceMessage "Info", "AddIns", "ScanDBInventoryLoadedVBASource", "Loaded VBA project source analysis completed.", completionDetail, "Scan DB Inventory - Loaded VBA Source", 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAModules"), "Table", "CodeDb", "AddIns"
    If showMessage Then ReportCoreActionSuccess "ScanDBInventoryLoadedVBASource", "AddIns", "Loaded VBA project source analysis completed successfully.", "Projects inspected: " & CStr(scannedProjectCount) & vbCrLf & "Project references: " & CStr(referenceCount) & vbCrLf & "Modules: " & CStr(moduleCount) & vbCrLf & "Public and private procedures: " & CStr(procedureCount) & vbCrLf & "Dependency evidence rows: " & CStr(dependencyCount), "Scan DB Inventory - Loaded VBA Source", ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAModules"), "Table", "CodeDb", "AddIns", True, "Scan DB Inventory - Loaded VBA Source", vbInformation, True, "Tables", False
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not referenceRows Is Nothing Then referenceRows.Close
    If Not moduleRows Is Nothing Then moduleRows.Close
    If Not procedureRows Is Nothing Then procedureRows.Close
    If Not dependencyRows Is Nothing Then dependencyRows.Close
    If transactionActive And Not ws Is Nothing Then ws.Rollback
    On Error GoTo 0
    ReportReferenceFailure "ScanDBInventoryLoadedVBASource", "Could not scan loaded VBA project source. The prior current source-analysis detail partitions and matching [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects") & "] summary state were preserved.", showMessage, savedErrNumber, savedErrSource, savedErrDescription
End Sub

Private Sub AddLoadedDependencyRow(ByVal dependencyRows As DAO.Recordset, ByVal projectIdentity As String, ByVal projectName As String, ByVal projectRole As String, ByVal sourceModuleIdentity As String, ByVal sourceModuleName As String, ByVal dependencyType As String, ByVal dependencyName As String, ByVal targetProjectIdentity As String, ByVal targetProjectName As String, ByVal targetModuleIdentity As String, ByVal targetModuleName As String, ByVal matchStatus As String, ByVal authoritativeMatch As Boolean, ByVal evidenceText As String, ByVal scannedOn As Date, Optional ByVal identitySeed As String = vbNullString)
    If Len(identitySeed) = 0 Then identitySeed = projectIdentity & "|" & sourceModuleIdentity & "|" & dependencyType & "|" & dependencyName & "|" & targetProjectIdentity & "|" & targetModuleIdentity & "|" & evidenceText
    identitySeed = "D|" & ReferenceScanTextFingerprint(identitySeed)
    BeginInventoryUpsertRow dependencyRows, identitySeed
    dependencyRows.fields(FIELD_DEPENDENCY_IDENTITY).value = identitySeed
    dependencyRows.fields(FIELD_PROJECT_IDENTITY).value = projectIdentity
    dependencyRows.fields(FIELD_PROJECT_NAME).value = SharedValue_NormalizeAssignment(projectName)
    dependencyRows.fields(FIELD_PROJECT_ROLE).value = projectRole
    dependencyRows.fields(FIELD_SOURCE_MODULE_IDENTITY).value = SharedValue_NormalizeAssignment(sourceModuleIdentity)
    dependencyRows.fields(FIELD_SOURCE_MODULE_NAME).value = SharedValue_NormalizeAssignment(sourceModuleName)
    dependencyRows.fields(FIELD_DEPENDENCY_TYPE).value = dependencyType
    dependencyRows.fields(FIELD_DEPENDENCY_NAME).value = SharedValue_NormalizeAssignment(dependencyName)
    dependencyRows.fields(FIELD_TARGET_PROJECT_IDENTITY).value = SharedValue_NormalizeAssignment(targetProjectIdentity)
    dependencyRows.fields(FIELD_TARGET_PROJECT_NAME).value = SharedValue_NormalizeAssignment(targetProjectName)
    dependencyRows.fields(FIELD_TARGET_MODULE_IDENTITY).value = SharedValue_NormalizeAssignment(targetModuleIdentity)
    dependencyRows.fields(FIELD_TARGET_MODULE_NAME).value = SharedValue_NormalizeAssignment(targetModuleName)
    dependencyRows.fields(FIELD_MATCH_STATUS).value = SharedValue_NormalizeAssignment(matchStatus)
    dependencyRows.fields(FIELD_IS_AUTHORITATIVE_MATCH).value = authoritativeMatch
    dependencyRows.fields(FIELD_EVIDENCE_TEXT).value = SharedValue_NormalizeAssignment(evidenceText)
    dependencyRows.fields(FIELD_DIAGNOSTIC_TEXT).value = Null
    dependencyRows.fields(FIELD_INVENTORY_STATUS).value = INVENTORY_STATUS_CURRENT
    dependencyRows.fields(FIELD_LAST_SCAN_BATCH_ID).value = ReferenceScanBatchID(scannedOn, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBADependencies"))
    dependencyRows.fields(FIELD_LAST_SCANNED_ON).value = scannedOn
    dependencyRows.Update
End Sub

Private Sub ScanOneLoadedVBAProjectSource(ByVal db As DAO.Database, ByVal vbProject As Object, ByVal projectIdentity As String, ByVal projectName As String, ByVal projectRole As String, ByVal scannedOn As Date, ByVal referenceRows As DAO.Recordset, ByVal moduleRows As DAO.Recordset, ByVal procedureRows As DAO.Recordset, ByVal dependencyRows As DAO.Recordset, ByRef referenceCount As Long, ByRef moduleCount As Long, ByRef procedureCount As Long, ByRef dependencyCount As Long, ByRef scanStatus As String, ByRef diagnosticText As String, ByRef moduleRosterSignature As String, ByRef modulesReusedByChatGPTID As Long, ByRef modulesReusedByFingerprint As Long, ByRef modulesSourceRescanned As Long, ByRef modulesRosterReconciled As Long, ByRef proceduresRetained As Long, ByRef proceduresRebuilt As Long)
    Dim referencesReadable As Boolean
    Dim componentsReadable As Boolean
    Dim anySourceReadable As Boolean
    Dim allSourceReadable As Boolean
    Dim protectionCode As Long
    protectionCode = SafeVBProjectLong(vbProject, "Protection", -1)
    CorePulseProgress "Loaded VBA project | " & projectName & " | references | starting"
    ScanLoadedProjectReferences db, vbProject, projectIdentity, projectName, projectRole, scannedOn, referenceRows, dependencyRows, referenceCount, dependencyCount, referencesReadable, diagnosticText
    CorePulseProgress "Loaded VBA project | " & projectName & " | references=" & CStr(referenceCount) & " | dependencies=" & CStr(dependencyCount) & " | modules: starting"
    ScanLoadedProjectModules db, vbProject, projectIdentity, projectName, projectRole, scannedOn, moduleRows, procedureRows, dependencyRows, moduleCount, procedureCount, dependencyCount, componentsReadable, anySourceReadable, allSourceReadable, diagnosticText, moduleRosterSignature, modulesReusedByChatGPTID, modulesReusedByFingerprint, modulesSourceRescanned, modulesRosterReconciled, proceduresRetained, proceduresRebuilt
    CorePulseProgress "Loaded VBA project | " & projectName & " | modules=" & CStr(moduleCount) & " | procedures=" & CStr(procedureCount) & " | dependencies=" & CStr(dependencyCount) & " | source readable=" & CStr(anySourceReadable)
    If referencesReadable = False And componentsReadable = False Then
        scanStatus = "Metadata Limited"
    ElseIf componentsReadable And moduleCount = 0 Then
        If referencesReadable Then scanStatus = "No Components" Else scanStatus = "No Components; Refs Limited"
    ElseIf allSourceReadable Then
        If referencesReadable Then scanStatus = "Source Scanned" Else scanStatus = "Source Scanned; Refs Limited"
    ElseIf anySourceReadable Then
        If referencesReadable Then scanStatus = "Source Partial" Else scanStatus = "Source Partial; Refs Limited"
    ElseIf protectionCode = 1 Then
        scanStatus = "Locked - Metadata Only"
    Else
        scanStatus = "Modules Metadata Only"
    End If
End Sub

Private Sub ScanLoadedProjectReferences(ByVal db As DAO.Database, ByVal vbProject As Object, ByVal projectIdentity As String, ByVal projectName As String, ByVal projectRole As String, ByVal scannedOn As Date, ByVal referenceRows As DAO.Recordset, ByVal dependencyRows As DAO.Recordset, ByRef storedCount As Long, ByRef dependencyCount As Long, ByRef referencesReadable As Boolean, ByRef diagnosticText As String)
    Dim projectReferences As Object
    Dim referenceItem As Object
    Dim referenceIndex As Long
    Dim referenceCount As Long
    Dim referenceName As String
    Dim referenceDescription As String
    Dim ReferenceGUID As String
    Dim referencePath As String
    Dim referenceIdentity As String
    Dim referenceKind As String
    Dim referenceKindCode As Long
    Dim majorVersion As Long
    Dim minorVersion As Long
    Dim isBuiltIn As Boolean
    Dim isBroken As Boolean
    Dim fileExistsValue As Boolean
    Dim matchStatus As String
    Dim matchMethod As String
    Dim matchEvidence As String
    Dim matchedProjectIdentity As String
    Dim matchedProjectName As String
    Dim matchedProjectPath As String
    Dim matchedProjectRole As String
    Dim authoritativeMatch As Boolean
    Dim readErrNumber As Long
    Dim readErrDescription As String
    On Error Resume Next
    Set projectReferences = CallByName(vbProject, "References", VbGet)
    readErrNumber = Err.Number
    readErrDescription = Err.Description
    Err.Clear
    If readErrNumber = 0 And Not projectReferences Is Nothing Then referenceCount = CLng(projectReferences.count)
    If Err.Number <> 0 Then
        readErrNumber = Err.Number
        readErrDescription = Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    If readErrNumber <> 0 Or projectReferences Is Nothing Then
        referencesReadable = False
        diagnosticText = SharedText_AppendNonBlankLine(diagnosticText, "Project references could not be read. Error " & CStr(readErrNumber) & ": " & readErrDescription)
        Exit Sub
    End If
    referencesReadable = True
    For referenceIndex = 1 To referenceCount
        PumpSetupResponsiveness
        Set referenceItem = Nothing
        readErrNumber = 0
        readErrDescription = vbNullString
        On Error Resume Next
        Set referenceItem = projectReferences.item(referenceIndex)
        readErrNumber = Err.Number
        readErrDescription = Err.Description
        Err.Clear
        On Error GoTo 0
        If readErrNumber <> 0 Or referenceItem Is Nothing Then
            referencesReadable = False
            diagnosticText = SharedText_AppendNonBlankLine(diagnosticText, "Reference " & CStr(referenceIndex) & " could not be read. Error " & CStr(readErrNumber) & ": " & readErrDescription)
            GoTo NextReference
        End If
        referenceName = SafeReferenceText(referenceItem, "Name")
        referenceDescription = SafeReferenceText(referenceItem, "Description")
        ReferenceGUID = NormalizeReferenceGuid(SafeReferenceText(referenceItem, "Guid"))
        referencePath = SafeReferenceText(referenceItem, "FullPath")
        majorVersion = SafeReferenceLong(referenceItem, "Major")
        minorVersion = SafeReferenceLong(referenceItem, "Minor")
        referenceKindCode = SafeReferenceLong(referenceItem, "Type")
        referenceKind = ReferenceKindText(referenceKindCode)
        isBuiltIn = SafeReferenceBoolean(referenceItem, "BuiltIn", False)
        isBroken = SafeReferenceBoolean(referenceItem, "IsBroken", False)
        fileExistsValue = ReferenceFileExists(referencePath)
        referenceIdentity = "R|" & ReferenceScanTextFingerprint(projectIdentity & "|" & ReferenceGUID & "|" & CStr(majorVersion) & "." & CStr(minorVersion) & "|" & NormalizeReferencePath(referencePath) & "|" & referenceName)
        matchStatus = vbNullString
        matchMethod = vbNullString
        matchEvidence = vbNullString
        matchedProjectIdentity = vbNullString
        matchedProjectName = vbNullString
        matchedProjectPath = vbNullString
        matchedProjectRole = vbNullString
        authoritativeMatch = False
        ResolveLoadedVBAProjectMatch db, referenceName, referencePath, matchStatus, matchMethod, matchEvidence, matchedProjectIdentity, matchedProjectName, matchedProjectPath, matchedProjectRole, authoritativeMatch
        PulseProgressEvery referenceIndex, 10, "Loaded VBA project | " & projectName & " | reference " & CStr(referenceIndex) & "/" & CStr(referenceCount) & " | " & IIf(Len(referenceName) > 0, referenceName, "(unnamed reference)") & " | match=" & matchStatus & " | stored=" & CStr(storedCount) & " | dependencies=" & CStr(dependencyCount), referenceCount, "ReferenceLoadedProjectReferenceScan", "References"
        BeginInventoryUpsertRow referenceRows, referenceIdentity
        referenceRows.fields(FIELD_REFERENCE_IDENTITY).value = referenceIdentity
        referenceRows.fields(FIELD_PROJECT_IDENTITY).value = projectIdentity
        referenceRows.fields(FIELD_PROJECT_NAME).value = SharedValue_NormalizeAssignment(projectName)
        referenceRows.fields(FIELD_PROJECT_ROLE).value = projectRole
        referenceRows.fields(FIELD_REFERENCE_ORDER).value = referenceIndex
        referenceRows.fields(FIELD_REFERENCE_NAME).value = SharedValue_NormalizeAssignment(referenceName)
        referenceRows.fields(FIELD_REFERENCE_DESCRIPTION).value = SharedValue_NormalizeAssignment(referenceDescription)
        referenceRows.fields(FIELD_REFERENCE_GUID).value = SharedValue_NormalizeAssignment(ReferenceGUID)
        referenceRows.fields(FIELD_MAJOR_VERSION).value = majorVersion
        referenceRows.fields(FIELD_MINOR_VERSION).value = minorVersion
        referenceRows.fields(FIELD_REFERENCE_KIND_CODE).value = referenceKindCode
        referenceRows.fields(FIELD_REFERENCE_KIND).value = referenceKind
        referenceRows.fields(FIELD_FULL_PATH).value = SharedValue_NormalizeAssignment(referencePath)
        referenceRows.fields(FIELD_FILE_EXISTS).value = fileExistsValue
        referenceRows.fields(FIELD_IS_BUILT_IN).value = isBuiltIn
        referenceRows.fields(FIELD_IS_BROKEN).value = isBroken
        referenceRows.fields(FIELD_MATCH_STATUS).value = SharedValue_NormalizeAssignment(matchStatus)
        referenceRows.fields(FIELD_MATCH_METHOD).value = SharedValue_NormalizeAssignment(matchMethod)
        referenceRows.fields(FIELD_IS_AUTHORITATIVE_MATCH).value = authoritativeMatch
        referenceRows.fields(FIELD_REFERENCED_PROJECT_IDENTITY).value = SharedValue_NormalizeAssignment(matchedProjectIdentity)
        referenceRows.fields(FIELD_REFERENCED_PROJECT_NAME).value = SharedValue_NormalizeAssignment(matchedProjectName)
        referenceRows.fields(FIELD_REFERENCED_PROJECT_ROLE).value = SharedValue_NormalizeAssignment(matchedProjectRole)
        referenceRows.fields(FIELD_DIAGNOSTIC_TEXT).value = Null
        referenceRows.fields(FIELD_INVENTORY_STATUS).value = INVENTORY_STATUS_CURRENT
        referenceRows.fields(FIELD_LAST_SCAN_BATCH_ID).value = ReferenceScanBatchID(scannedOn, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjectReferences"))
        referenceRows.fields(FIELD_LAST_SCANNED_ON).value = scannedOn
        referenceRows.Update
        AddLoadedDependencyRow dependencyRows, projectIdentity, projectName, projectRole, vbNullString, vbNullString, "Project Reference", referenceName, matchedProjectIdentity, matchedProjectName, vbNullString, vbNullString, matchStatus, authoritativeMatch, matchEvidence, scannedOn, referenceIdentity
        storedCount = storedCount + 1
        dependencyCount = dependencyCount + 1
NextReference:
    Next referenceIndex
End Sub

Private Function ManagerReferenceMatchKey(ByVal ReferenceGUID As String, ByVal majorVersion As Long, ByVal minorVersion As Long, ByVal pathText As String, ByVal referenceName As String) As String
    ReferenceGUID = NormalizeReferenceGuid(ReferenceGUID)
    If Len(ReferenceGUID) > 0 Then
        ManagerReferenceMatchKey = ReferenceVersionIdentity(ReferenceGUID, majorVersion, minorVersion)
    ElseIf Len(NormalizeReferencePath(pathText)) > 0 Then
        ManagerReferenceMatchKey = ReferencePathIdentity(pathText)
    Else
        ManagerReferenceMatchKey = "N|" & UCase$(Trim$(referenceName))
    End If
End Function

Private Function ManagerReferenceIdentity(ByVal sourceText As String, ByVal ReferenceGUID As String, ByVal majorVersion As Long, ByVal minorVersion As Long, ByVal pathText As String, ByVal referenceName As String, ByVal sourceID As Long) As String
    Dim identityText As String
    ReferenceGUID = NormalizeReferenceGuid(ReferenceGUID)
    If Len(ReferenceGUID) > 0 Then
        identityText = ReferenceVersionIdentity(ReferenceGUID, majorVersion, minorVersion)
    ElseIf Len(NormalizeReferencePath(pathText)) > 0 Then
        identityText = ReferencePathIdentity(pathText)
    Else
        identityText = "N|" & UCase$(Trim$(referenceName)) & "|" & CStr(sourceID)
    End If
    ManagerReferenceIdentity = Left$("M|" & UCase$(sourceText) & "|" & identityText & "|" & CStr(sourceID), 255)
End Function

Private Function ReferenceChangeProtectionReason(ByVal referenceName As String, ByVal isBuiltIn As Boolean) As String
    If isBuiltIn Then
        ReferenceChangeProtectionReason = "Access reports this reference as built in."
        Exit Function
    End If
    Select Case UCase$(Trim$(referenceName))
        Case "ACCESS", "VBA", "OFFICE", "DAO", "ACEDAO", "VBIDE", "STDOLE"
            ReferenceChangeProtectionReason = "This reference is required by Access, VBA, DAO, Office, or Access Development Console runtime code and is protected from programmatic deactivation."
    End Select
End Function

Private Sub CollectReferenceFileMetadata(ByVal pathText As String, ByRef fileNameText As String, ByRef folderPath As String, ByRef fileExtension As String, ByRef fileSizeBytes As Double, ByRef fileCreatedOn As Variant, ByRef fileModifiedOn As Variant, ByRef fileVersion As String, ByRef fileAttributes As String)
    On Error Resume Next
    Dim fso As Object
    Dim fileItem As Object
    Dim expandedPath As String
    expandedPath = ExpandReferencePath(pathText)
    fileNameText = vbNullString
    folderPath = vbNullString
    fileExtension = vbNullString
    fileSizeBytes = 0
    fileCreatedOn = Null
    fileModifiedOn = Null
    fileVersion = vbNullString
    fileAttributes = vbNullString
    If Len(Trim$(expandedPath)) = 0 Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    fileNameText = fso.GetFileName(expandedPath)
    folderPath = fso.GetParentFolderName(expandedPath)
    fileExtension = LCase$(fso.GetExtensionName(expandedPath))
    If fso.fileExists(expandedPath) = False Then Exit Sub
    Set fileItem = fso.GetFile(expandedPath)
    fileSizeBytes = CDbl(fileItem.Size)
    fileCreatedOn = fileItem.DateCreated
    fileModifiedOn = fileItem.DateLastModified
    fileVersion = CStr(fso.GetFileVersion(expandedPath))
    fileAttributes = ReferenceFileAttributesText(CLng(fileItem.Attributes))
    On Error GoTo 0
End Sub

Private Function ReferenceFileAttributesText(ByVal attributeValue As Long) As String
    Dim resultText As String
    If (attributeValue And 1) <> 0 Then resultText = AppendDelimitedValue(resultText, "ReadOnly")
    If (attributeValue And 2) <> 0 Then resultText = AppendDelimitedValue(resultText, "Hidden")
    If (attributeValue And 4) <> 0 Then resultText = AppendDelimitedValue(resultText, "System")
    If (attributeValue And 16) <> 0 Then resultText = AppendDelimitedValue(resultText, "Directory")
    If (attributeValue And 32) <> 0 Then resultText = AppendDelimitedValue(resultText, "Archive")
    If (attributeValue And 1024) <> 0 Then resultText = AppendDelimitedValue(resultText, "ReparsePoint")
    If (attributeValue And 2048) <> 0 Then resultText = AppendDelimitedValue(resultText, "Compressed")
    If Len(resultText) = 0 Then resultText = "Normal"
    ReferenceFileAttributesText = resultText
End Function

Private Function ReferenceKindText(ByVal kindCode As Long) As String
    Select Case kindCode
        Case 0: ReferenceKindText = "Type Library"
        Case 1: ReferenceKindText = "Project"
        Case 2: ReferenceKindText = "Control"
        Case Else: ReferenceKindText = "Unknown (" & CStr(kindCode) & ")"
    End Select
End Function


Private Function ReferenceRegistryRootDisplay(ByVal hiveValue As Long) As String
    Select Case hiveValue
        Case HKEY_CURRENT_USER: ReferenceRegistryRootDisplay = "HKEY_CURRENT_USER"
        Case HKEY_LOCAL_MACHINE: ReferenceRegistryRootDisplay = "HKEY_LOCAL_MACHINE"
        Case Else: ReferenceRegistryRootDisplay = "HKEY_" & Hex$(hiveValue)
    End Select
End Function

Private Function ReferenceRegistryProvider() As Object
    On Error GoTo ErrHandler
    Set ReferenceRegistryProvider = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "ReferenceRegistryProvider", "The Windows registry provider could not be opened." & vbCrLf & Err.Description
End Function

Private Function ReferenceRegistryRoots() As Variant
    ReferenceRegistryRoots = Array(Array(HKEY_CURRENT_USER, REGISTRY_TYPELIB_NATIVE, "Current User", "Native"), Array(HKEY_CURRENT_USER, REGISTRY_TYPELIB_WOW6432, "Current User", "WOW6432"), Array(HKEY_LOCAL_MACHINE, REGISTRY_TYPELIB_NATIVE, "Local Machine", "Native"), Array(HKEY_LOCAL_MACHINE, REGISTRY_TYPELIB_WOW6432, "Local Machine", "WOW6432"))
End Function

Private Function AccessMenuAddInRegistryRoots() As Variant
    ' CANONICAL AUTHORITY: this fixed root set enumerates the Access Menu Add-Ins registration locations for the current Office major version, including native, WOW6432, and Click-to-Run machine views.
    Dim officeVersion As String
    officeVersion = Trim$(Application.Version)
    AccessMenuAddInRegistryRoots = Array( _
        Array(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Office\" & officeVersion & "\Access\Menu Add-Ins"), _
        Array(HKEY_CURRENT_USER, "SOFTWARE\WOW6432Node\Microsoft\Office\" & officeVersion & "\Access\Menu Add-Ins"), _
        Array(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Office\" & officeVersion & "\Access\Menu Add-Ins"), _
        Array(HKEY_LOCAL_MACHINE, "SOFTWARE\WOW6432Node\Microsoft\Office\" & officeVersion & "\Access\Menu Add-Ins"), _
        Array(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\" & officeVersion & "\Access\Menu Add-Ins"), _
        Array(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\" & officeVersion & "\Access\Menu Add-Ins"))
End Function

Private Function FindReferenceRegistration(ByVal ReferenceGUID As String, ByVal majorVersion As Long, ByVal minorVersion As Long, ByVal activeFullPath As String, ByRef registrationScope As String, ByRef registrationView As String, ByRef platformText As String, ByRef registeredPath As String, ByRef registryLCID As String, ByRef registryKeyPath As String, ByRef typeLibFlags As String, ByRef helpDirectory As String) As Boolean
    On Error GoTo ErrHandler
    Dim provider As Object
    Dim roots As Variant
    Dim rootInfo As Variant
    Dim localeKeys As Variant
    Dim localeItem As Variant
    Dim platformKeys As Variant
    Dim platformItem As Variant
    Dim versionPath As String
    Dim localePath As String
    Dim platformPath As String
    Dim candidatePath As String
    Dim firstScope As String
    Dim firstView As String
    Dim firstPlatform As String
    Dim firstPath As String
    Dim firstLCID As String
    Dim firstKeyPath As String
    Dim firstTypeLibFlags As String
    Dim firstHelpDirectory As String
    Dim currentTypeLibFlags As String
    Dim currentHelpDirectory As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    If Len(ReferenceGUID) = 0 Then Exit Function
    Set provider = ReferenceRegistryProvider()
    roots = ReferenceRegistryRoots()
    For Each rootInfo In roots
        PumpSetupResponsiveness
        versionPath = CStr(rootInfo(1)) & "\" & ReferenceGUID & "\" & CStr(majorVersion) & "." & CStr(minorVersion)
        currentTypeLibFlags = RegistryDefaultString(provider, CLng(rootInfo(0)), versionPath & "\FLAGS")
        currentHelpDirectory = ExpandReferencePath(RegistryDefaultString(provider, CLng(rootInfo(0)), versionPath & "\HELPDIR"))
        localeKeys = RegistrySubkeys(provider, CLng(rootInfo(0)), versionPath)
        If IsArray(localeKeys) Then
            For Each localeItem In localeKeys
                PumpSetupResponsiveness
                If IsReferenceLocaleKey(CStr(localeItem)) Then
                    localePath = versionPath & "\" & CStr(localeItem)
                    platformKeys = RegistrySubkeys(provider, CLng(rootInfo(0)), localePath)
                    If IsArray(platformKeys) Then
                        For Each platformItem In platformKeys
                            PumpSetupResponsiveness
                            If IsReferencePlatformKey(CStr(platformItem)) Then
                                platformPath = localePath & "\" & CStr(platformItem)
                                candidatePath = ExpandReferencePath(RegistryDefaultString(provider, CLng(rootInfo(0)), platformPath))
                                If Len(firstPath) = 0 Then
                                    firstScope = CStr(rootInfo(2))
                                    firstView = CStr(rootInfo(3))
                                    firstPlatform = LCase$(CStr(platformItem))
                                    firstPath = candidatePath
                                    firstLCID = CStr(localeItem)
                                    firstKeyPath = ReferenceRegistryRootDisplay(CLng(rootInfo(0))) & "\" & platformPath
                                    firstTypeLibFlags = currentTypeLibFlags
                                    firstHelpDirectory = currentHelpDirectory
                                End If
                                If Len(activeFullPath) > 0 And StrComp(NormalizeReferencePath(candidatePath), NormalizeReferencePath(activeFullPath), vbTextCompare) = 0 Then
                                    registrationScope = CStr(rootInfo(2))
                                    registrationView = CStr(rootInfo(3))
                                    platformText = LCase$(CStr(platformItem))
                                    registeredPath = candidatePath
                                    registryLCID = CStr(localeItem)
                                    registryKeyPath = ReferenceRegistryRootDisplay(CLng(rootInfo(0))) & "\" & platformPath
                                    typeLibFlags = currentTypeLibFlags
                                    helpDirectory = currentHelpDirectory
                                    FindReferenceRegistration = True
                                    Exit Function
                                End If
                            End If
                        Next platformItem
                    End If
                End If
            Next localeItem
        End If
    Next rootInfo
    If Len(firstPath) > 0 Then
        registrationScope = firstScope
        registrationView = firstView
        platformText = firstPlatform
        registeredPath = firstPath
        registryLCID = firstLCID
        registryKeyPath = firstKeyPath
        typeLibFlags = firstTypeLibFlags
        helpDirectory = firstHelpDirectory
        FindReferenceRegistration = True
    End If
    Exit Function
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "FindReferenceRegistration", "Could not inspect type-library registration for an active reference." & vbCrLf & "GUID: " & ReferenceGUID & vbCrLf & "Version: " & CStr(majorVersion) & "." & CStr(minorVersion) & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Function SafeReferenceText(ByVal referenceItem As Object, ByVal propertyName As String) As String
    On Error GoTo ExitHere
    SafeReferenceText = CStr(CallByName(referenceItem, propertyName, VbGet))
ExitHere:
End Function

Private Function SafeReferenceLong(ByVal referenceItem As Object, ByVal propertyName As String) As Long
    On Error GoTo ExitHere
    SafeReferenceLong = CLng(CallByName(referenceItem, propertyName, VbGet))
ExitHere:
End Function

Private Function SafeReferenceBoolean(ByVal referenceItem As Object, ByVal propertyName As String, ByVal DefaultValue As Boolean) As Boolean
    On Error GoTo UseDefault
    SafeReferenceBoolean = CBool(CallByName(referenceItem, propertyName, VbGet))
    Exit Function
UseDefault:
    SafeReferenceBoolean = DefaultValue
End Function


Private Function IsReferenceLocaleKey(ByVal keyText As String) As Boolean
    On Error GoTo InvalidKey
    Dim localeValue As Long
    keyText = Trim$(keyText)
    If Len(keyText) = 0 Then Exit Function
    localeValue = CLng("&H" & keyText)
    IsReferenceLocaleKey = True
    Exit Function
InvalidKey:
    IsReferenceLocaleKey = False
End Function

Private Function IsReferencePlatformKey(ByVal keyText As String) As Boolean
    Select Case LCase$(Trim$(keyText))
        Case "win16", "win32", "win64": IsReferencePlatformKey = True
    End Select
End Function

Private Function ReferencePathIdentity(ByVal pathText As String) As String
    pathText = NormalizeReferencePath(pathText)
    If Len(pathText) > 0 Then ReferencePathIdentity = "P|" & Left$(pathText, 240)
End Function

Private Function ReferenceVersionIdentity(ByVal guidText As String, ByVal majorVersion As Long, ByVal minorVersion As Long) As String
    ReferenceVersionIdentity = "K" & guidText & "|" & CStr(majorVersion) & "." & CStr(minorVersion)
End Function


Private Function ReferenceFileExists(ByVal pathText As String) As Boolean
    On Error GoTo Missing
    pathText = ExpandReferencePath(pathText)
    If Len(pathText) = 0 Then Exit Function
    GetAttr pathText
    ReferenceFileExists = True
    Exit Function
Missing:
    ReferenceFileExists = False
End Function


' Purpose: Opens reference file location for reference path.
' Public role: Inter-module public API. Other ADC modules call this routine directly; its Public visibility is part of the module-to-module contract.
' Parameters:
'   referencePath (Input; String): Filesystem or database path identifying the reference path location used by the operation.
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub OpenReferenceFileLocation(ByVal referencePath As String, Optional ByVal showMessage As Boolean = True)
    On Error GoTo ErrHandler
    Dim expandedPath As String
    Dim folderPath As String
    Dim commandText As String
    expandedPath = ExpandReferencePath(referencePath)
    If Len(Trim$(expandedPath)) = 0 Then Err.Raise vbObjectError + 9730, "OpenReferenceFileLocation", "The selected reference does not contain a file path."
    folderPath = ReferenceFolderPath(expandedPath)
    If ReferenceFileExists(expandedPath) Then
        commandText = "explorer.exe /select,""" & expandedPath & """"
    ElseIf ReferenceFolderExists(folderPath) Then
        commandText = "explorer.exe """ & folderPath & """"
    Else
        Err.Raise vbObjectError + 9731, "OpenReferenceFileLocation", "Neither the reference file nor its containing folder currently exists." & vbCrLf & "Path: " & expandedPath
    End If
    shell commandText, vbNormalFocus
    Exit Sub
ErrHandler:
    ReportReferenceFailure "OpenReferenceFileLocation", "Could not open the selected reference file location.", showMessage, Err.Number, Err.Source, Err.Description
End Sub

Private Sub AddReferenceFromManagerRecord(ByVal rs As DAO.Recordset)
    On Error GoTo TryFile
    Dim ReferenceGUID As String
    Dim majorVersion As Long
    Dim minorVersion As Long
    Dim activationPath As String
    Dim guidErrNumber As Long
    Dim guidErrSource As String
    Dim guidErrDescription As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    ReferenceGUID = NormalizeReferenceGuid(CStr(Nz(rs.fields(FIELD_REFERENCE_GUID).value, vbNullString)))
    majorVersion = CLng(Nz(rs.fields(FIELD_MAJOR_VERSION).value, 0))
    minorVersion = CLng(Nz(rs.fields(FIELD_MINOR_VERSION).value, 0))
    activationPath = ReferenceManagerActivationPath(rs)
    If Len(ReferenceGUID) > 0 Then
        Application.references.AddFromGuid ReferenceGUID, majorVersion, minorVersion
        Exit Sub
    End If
TryFile:
    guidErrNumber = Err.Number
    guidErrSource = Err.Source
    guidErrDescription = Err.Description
    Err.Clear
    On Error GoTo ErrHandler
    If ReferenceFileExists(activationPath) = False Then
        If guidErrNumber <> 0 Then Err.Raise guidErrNumber, guidErrSource, guidErrDescription
        Err.Raise vbObjectError + 9714, "AddReferenceFromManagerRecord", "The activation path does not exist: " & IIf(Len(activationPath) = 0, "(blank)", activationPath)
    End If
    Application.references.AddFromFile activationPath
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "AddReferenceFromManagerRecord", "Could not activate reference " & ReferenceManagerRecordDisplay(rs) & "." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub RemoveReferenceFromManagerRecord(ByVal rs As DAO.Recordset)
    On Error GoTo ErrHandler
    Dim referenceItem As Object
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Set referenceItem = FindActiveReferenceForManagerRecord(rs)
    If referenceItem Is Nothing Then Err.Raise vbObjectError + 9715, "RemoveReferenceFromManagerRecord", "The active reference could not be found."
    If SafeReferenceBoolean(referenceItem, "BuiltIn", False) Then Err.Raise vbObjectError + 9716, "RemoveReferenceFromManagerRecord", "Built-in references cannot be removed."
    Application.references.Remove referenceItem
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Err.Raise savedErrNumber, "RemoveReferenceFromManagerRecord", "Could not deactivate reference " & ReferenceManagerRecordDisplay(rs) & "." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub RestoreOriginalReferenceState(ByRef rollbackDetail As String)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim originalKeys As Collection
    Dim referenceItem As Object
    Dim referenceIndex As Long
    Dim referenceKey As String
    Dim rs As DAO.Recordset
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Set db = ADCDb()
    ' INTENTIONAL OPERATION SNAPSHOT: originalKeys captures one coherent source state for this invocation of RestoreOriginalReferenceState; validity ends with the operation rather than becoming a live data
    ' authority.
    Set originalKeys = New Collection
    Set rs = db.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ReferenceManager") & "] WHERE " & SQL_FIELD_ORIGINAL_ACTIVE & "=True ORDER BY " & SQL_FIELD_REFERENCE_PRIORITY & "," & SQL_FIELD_ID & ";", dbOpenSnapshot)
    Do While rs.EOF = False
        PumpSetupResponsiveness
        AddCollectionKey originalKeys, ReferenceManagerRecordKey(rs)
        rs.MoveNext
    Loop
    rs.Close
    For referenceIndex = Application.references.count To 1 Step -1
        PumpSetupResponsiveness
        Set referenceItem = Application.references(referenceIndex)
        referenceKey = ActiveReferenceObjectKey(referenceItem)
        If SafeReferenceBoolean(referenceItem, "BuiltIn", False) = False Then
            If CollectionHasKey(originalKeys, referenceKey) = False Then Application.references.Remove referenceItem
        End If
    Next referenceIndex
    Set rs = db.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ReferenceManager") & "] WHERE " & SQL_FIELD_ORIGINAL_ACTIVE & "=True ORDER BY " & SQL_FIELD_REFERENCE_PRIORITY & "," & SQL_FIELD_ID & ";", dbOpenSnapshot)
    Do While rs.EOF = False
        PumpSetupResponsiveness
        If ReferenceManagerRecordIsCurrentlyActive(rs) = False Then AddReferenceFromManagerRecord rs
        rs.MoveNext
    Loop
    rs.Close
    rollbackDetail = "The original active-reference membership was restored. Any reference that had to be removed and re-added may appear later in the reference priority list."
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
    rollbackDetail = "Reference rollback was incomplete. Err.Number: " & CStr(savedErrNumber) & " | Err.Source: " & savedErrSource & " | Err.Description: " & savedErrDescription
End Sub

Private Function ReferenceFolderPath(ByVal pathText As String) As String
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    ReferenceFolderPath = fso.GetParentFolderName(ExpandReferencePath(pathText))
    On Error GoTo 0
End Function

Private Function ReferenceFolderExists(ByVal folderPath As String) As Boolean
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    ReferenceFolderExists = fso.FolderExists(folderPath)
    On Error GoTo 0
End Function

Private Sub ScanLoadedProjectModules(ByVal db As DAO.Database, ByVal vbProject As Object, ByVal projectIdentity As String, ByVal projectName As String, ByVal projectRole As String, ByVal scannedOn As Date, ByVal moduleRows As DAO.Recordset, ByVal procedureRows As DAO.Recordset, ByVal dependencyRows As DAO.Recordset, ByRef moduleCount As Long, ByRef procedureCount As Long, ByRef dependencyCount As Long, ByRef componentsReadable As Boolean, ByRef anySourceReadable As Boolean, ByRef allSourceReadable As Boolean, ByRef diagnosticText As String, ByRef moduleRosterSignature As String, ByRef modulesReusedByChatGPTID As Long, ByRef modulesReusedByFingerprint As Long, ByRef modulesSourceRescanned As Long, ByRef modulesRosterReconciled As Long, ByRef proceduresRetained As Long, ByRef proceduresRebuilt As Long)
    Dim components As Object
    Dim sourceComponent As Object
    Dim targetComponent As Object
    ' INTENTIONAL OPERATION-SCOPED KEYED LOOKUP: moduleRosterByName projects the already-enumerated VBComponents collection by module name for this project scan only, supporting roster signatures and
    ' qualified-reference resolution without repeated target scans. It is not a reusable cache or canonical source authority.
    Dim moduleRosterByName As Object
    ' INTENTIONAL OPERATION-SCOPED KEYED LOOKUP: newModuleParentIdentities tracks only parent rows inserted by this scan's RI pre-pass so they are not mistaken for reusable prior source inventory.
    Dim newModuleParentIdentities As Object
    'The first pass captures one authoritative component roster. Parent materialization, cache lookup, source scanning, and child writes all consume these same captured values.
    'Handled VBE read failures therefore cannot change ModuleIdentity between passes.
    Dim rosterComponents() As Object
    Dim rosterComponentNames() As String
    Dim rosterComponentTypeCodes() As Long
    Dim rosterModuleIdentities() As String
    Dim rosterReadErrNumbers() As Long
    Dim rosterReadErrDescriptions() As String
    Dim moduleIndex As Long
    Dim targetIndex As Long
    Dim componentCount As Long
    Dim targetComponentName As String
    Dim moduleIdentity As String
    Dim componentTypeCode As Long
    Dim readErrNumber As Long
    Dim readErrDescription As String
    Dim rosterChanged As Boolean
    Dim previousRosterSignature As String
    On Error Resume Next
    Set components = CallByName(vbProject, "VBComponents", VbGet)
    readErrNumber = Err.Number
    readErrDescription = Err.Description
    Err.Clear
    If readErrNumber = 0 And Not components Is Nothing Then componentCount = CLng(components.count)
    If Err.Number <> 0 Then
        readErrNumber = Err.Number
        readErrDescription = Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    If readErrNumber <> 0 Or components Is Nothing Then
        componentsReadable = False
        diagnosticText = SharedText_AppendNonBlankLine(diagnosticText, "Project modules could not be enumerated. Error " & CStr(readErrNumber) & ": " & readErrDescription)
        Exit Sub
    End If
    componentsReadable = True
    allSourceReadable = (componentCount > 0)
    Set moduleRosterByName = CreateObject("Scripting.Dictionary")
    moduleRosterByName.CompareMode = vbTextCompare
    Set newModuleParentIdentities = CreateObject("Scripting.Dictionary")
    newModuleParentIdentities.CompareMode = vbTextCompare
    If componentCount > 0 Then
        ReDim rosterComponents(1 To componentCount)
        ReDim rosterComponentNames(1 To componentCount)
        ReDim rosterComponentTypeCodes(1 To componentCount)
        ReDim rosterModuleIdentities(1 To componentCount)
        ReDim rosterReadErrNumbers(1 To componentCount)
        ReDim rosterReadErrDescriptions(1 To componentCount)
    End If
    For targetIndex = 1 To componentCount
        PumpSetupResponsiveness
        Set targetComponent = Nothing
        readErrNumber = 0
        readErrDescription = vbNullString
        On Error Resume Next
        Set targetComponent = components.item(targetIndex)
        readErrNumber = Err.Number
        readErrDescription = Err.Description
        Err.Clear
        If readErrNumber <> 0 Or targetComponent Is Nothing Then
            readErrNumber = 0
            readErrDescription = vbNullString
            Set targetComponent = components.item(targetIndex)
            readErrNumber = Err.Number
            readErrDescription = Err.Description
            Err.Clear
        End If
        On Error GoTo 0
        rosterReadErrNumbers(targetIndex) = readErrNumber
        rosterReadErrDescriptions(targetIndex) = readErrDescription
        If readErrNumber = 0 And Not targetComponent Is Nothing Then
            Set rosterComponents(targetIndex) = targetComponent
            targetComponentName = SafeVBComponentText(targetComponent, "Name")
            If Len(targetComponentName) = 0 Then targetComponentName = SafeVBComponentText(targetComponent, "Name")
            componentTypeCode = SafeVBComponentLong(targetComponent, "Type", -1)
            If componentTypeCode = -1 Then componentTypeCode = SafeVBComponentLong(targetComponent, "Type", -1)
            moduleIdentity = "M|" & ReferenceScanTextFingerprint(projectIdentity & "|" & targetComponentName & "|" & CStr(componentTypeCode))
            rosterComponentNames(targetIndex) = targetComponentName
            rosterComponentTypeCodes(targetIndex) = componentTypeCode
            rosterModuleIdentities(targetIndex) = moduleIdentity
            If Len(targetComponentName) > 0 Then moduleRosterByName(targetComponentName) = Array(targetComponentName, componentTypeCode, moduleIdentity)
            If EnsureLoadedVBAModuleParentForChildWrites(moduleRows, moduleIdentity, projectIdentity, projectName, projectRole, targetIndex, targetComponentName, componentTypeCode, scannedOn) Then newModuleParentIdentities(moduleIdentity) = True
        End If
    Next targetIndex
    moduleRosterSignature = LoadedVBAModuleRosterSignature(moduleRosterByName)
    previousRosterSignature = StoredLoadedVBAModuleRosterSignature(db, projectIdentity)
    rosterChanged = (Len(previousRosterSignature) = 0 Or StrComp(previousRosterSignature, moduleRosterSignature, vbBinaryCompare) <> 0)
    If rosterChanged Then
        CorePulseProgress "Loaded VBA project | " & projectName & " | module roster changed | qualified dependencies will be reconciled without reparsing reusable module source"
    Else
        CorePulseProgress "Loaded VBA project | " & projectName & " | module roster unchanged | reusable resolved dependencies can be retained"
    End If
    For moduleIndex = 1 To componentCount
        PumpSetupResponsiveness
        If SetupControlCancellationRequested() Then RequireSafeWorkContinue "ScanLoadedProjectModules", "Loaded VBA module boundary"
        Set sourceComponent = rosterComponents(moduleIndex)
        ScanLoadedProjectModule db, sourceComponent, rosterReadErrNumbers(moduleIndex), rosterReadErrDescriptions(moduleIndex), moduleRosterByName, newModuleParentIdentities, _
            projectIdentity, projectName, projectRole, scannedOn, moduleRows, procedureRows, dependencyRows, moduleIndex, componentCount, rosterComponentNames(moduleIndex), _
            rosterComponentTypeCodes(moduleIndex), rosterModuleIdentities(moduleIndex), rosterChanged, moduleCount, procedureCount, dependencyCount, anySourceReadable, allSourceReadable, diagnosticText, _
            modulesReusedByChatGPTID, modulesReusedByFingerprint, modulesSourceRescanned, modulesRosterReconciled, proceduresRetained, proceduresRebuilt
    Next moduleIndex
End Sub

Private Sub ScanLoadedProjectModule( _
    ByVal db As DAO.Database, ByVal sourceComponent As Object, ByVal readErrNumber As Long, ByVal readErrDescription As String, _
    ByVal moduleRosterByName As Object, ByVal newModuleParentIdentities As Object, ByVal projectIdentity As String, ByVal projectName As String, _
    ByVal projectRole As String, ByVal scannedOn As Date, ByVal moduleRows As DAO.Recordset, ByVal procedureRows As DAO.Recordset, ByVal dependencyRows As DAO.Recordset, _
    ByVal moduleIndex As Long, ByVal componentCount As Long, ByVal componentName As String, ByVal componentTypeCode As Long, ByVal moduleIdentity As String, ByVal rosterChanged As Boolean, _
    ByRef moduleCount As Long, ByRef procedureCount As Long, ByRef dependencyCount As Long, ByRef anySourceReadable As Boolean, ByRef allSourceReadable As Boolean, ByRef diagnosticText As String, _
    ByRef modulesReusedByChatGPTID As Long, ByRef modulesReusedByFingerprint As Long, ByRef modulesSourceRescanned As Long, ByRef modulesRosterReconciled As Long, _
    ByRef proceduresRetained As Long, ByRef proceduresRebuilt As Long)
    'Processes one component captured by ScanLoadedProjectModules after the authoritative roster pre-pass; source reuse/rescan, procedure/dependency reconciliation, and module-row persistence remain one lifecycle.
    Dim sourceModule As Object
    Dim priorModuleRows As DAO.Recordset
    Dim ComponentTypeText As String
    Dim sourceText As String
    Dim dependencyText As String
    Dim qualifiedReferenceTokens As String
    Dim headerReferenceDependencies As String
    Dim headerObjectDependencies As String
    Dim sourceLineCount As Long
    Dim declarationLineCount As Long
    Dim moduleProcedureCount As Long
    Dim modulePublicProcedureCount As Long
    Dim externallyCallableCount As Long
    Dim hasOptionPrivateModule As Boolean
    Dim componentSourceReadable As Boolean
    Dim moduleScanStatus As String
    Dim moduleDiagnostic As String
    Dim currentChatGPTID As String
    Dim priorScannedChatGPTID As String
    Dim priorSourceFingerprint As String
    Dim priorQualifiedReferenceTokens As String
    Dim priorSourceScanVersion As Long
    Dim priorSourceScanCompletedOn As Variant
    Dim priorModuleLastScannedOn As Variant
    Dim priorSourceReadable As Boolean
    Dim priorSourceLineCount As Long
    Dim priorDeclarationLineCount As Long
    Dim priorProcedureCount As Long
    Dim priorPublicProcedureCount As Long
    Dim priorExternallyCallableCount As Long
    Dim priorHasOptionPrivateModule As Boolean
    Dim priorHeaderReferenceDependencies As String
    Dim priorHeaderObjectDependencies As String
    Dim priorScanStatus As String
    Dim priorDiagnosticText As String
    Dim priorInventoryFound As Boolean
    Dim sourceFingerprint As String
    Dim reuseSourceInventory As Boolean
    Dim reuseByChatGPTID As Boolean
    Dim retainedProcedureRows As Long
    Dim retainedDependencyRows As Long
    Dim rebuiltDependencyRows As Long
    Dim moduleTimer As Single
    Dim stageTimer As Single
    Dim sourceAcquisitionSeconds As Double
    Dim headerInspectionSeconds As Double
    Dim procedureScanSeconds As Double
    Dim dependencyPreparationSeconds As Double
    Dim tokenExtractionSeconds As Double
    Dim dependencyResolutionSeconds As Double
    Dim fingerprintSeconds As Double
    Dim storageSeconds As Double
    Dim cacheReason As String
    If readErrNumber <> 0 Or sourceComponent Is Nothing Then
        allSourceReadable = False
        diagnosticText = SharedText_AppendNonBlankLine(diagnosticText, "Module " & CStr(moduleIndex) & " could not be read. Error " & CStr(readErrNumber) & ": " & readErrDescription)
        Exit Sub
    End If
    PulseProgressEvery moduleIndex, 10, "Loaded VBA project | " & projectName & " | module " & CStr(moduleIndex) & "/" & CStr(componentCount) & " | " & componentName & " | checking cached source identity", componentCount, "ReferenceLoadedProjectModuleScan", "References"
    ComponentTypeText = VBComponentTypeText(componentTypeCode)
    'Final RI safeguard: re-seek the exact captured ModuleIdentity immediately before any procedure/dependency path can write a child. This is normally a no-op after the roster pre-pass.
    If EnsureLoadedVBAModuleParentForChildWrites(moduleRows, moduleIdentity, projectIdentity, projectName, projectRole, moduleIndex, componentName, componentTypeCode, scannedOn) Then newModuleParentIdentities(moduleIdentity) = True
    sourceText = vbNullString
    dependencyText = vbNullString
    qualifiedReferenceTokens = vbNullString
    headerReferenceDependencies = vbNullString
    headerObjectDependencies = vbNullString
    sourceFingerprint = vbNullString
    sourceLineCount = 0
    declarationLineCount = 0
    moduleProcedureCount = 0
    modulePublicProcedureCount = 0
    externallyCallableCount = 0
    hasOptionPrivateModule = False
    componentSourceReadable = False
    moduleDiagnostic = vbNullString
    currentChatGPTID = vbNullString
    priorScannedChatGPTID = vbNullString
    priorSourceFingerprint = vbNullString
    priorQualifiedReferenceTokens = vbNullString
    priorSourceScanVersion = 0
    priorSourceScanCompletedOn = Null
    priorModuleLastScannedOn = Null
    priorSourceReadable = False
    priorSourceLineCount = 0
    priorDeclarationLineCount = 0
    priorProcedureCount = 0
    priorPublicProcedureCount = 0
    priorExternallyCallableCount = 0
    priorHasOptionPrivateModule = False
    priorHeaderReferenceDependencies = vbNullString
    priorHeaderObjectDependencies = vbNullString
    priorScanStatus = vbNullString
    priorDiagnosticText = vbNullString
    priorInventoryFound = False
    reuseSourceInventory = False
    reuseByChatGPTID = False
    retainedProcedureRows = 0
    retainedDependencyRows = 0
    rebuiltDependencyRows = 0
    sourceAcquisitionSeconds = 0
    headerInspectionSeconds = 0
    procedureScanSeconds = 0
    dependencyPreparationSeconds = 0
    tokenExtractionSeconds = 0
    dependencyResolutionSeconds = 0
    fingerprintSeconds = 0
    storageSeconds = 0
    cacheReason = vbNullString
    moduleTimer = Timer
    If newModuleParentIdentities.Exists(moduleIdentity) = False Then
        Set priorModuleRows = db.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAModules") & "] WHERE " & SQL_FIELD_MODULE_IDENTITY & "=" & SharedDao_SqlQuote(moduleIdentity) & ";", dbOpenSnapshot)
        If Not priorModuleRows.EOF Then
            priorInventoryFound = True
            priorScannedChatGPTID = CStr(Nz(priorModuleRows.fields(FIELD_SCANNED_CHAT_GPTID).value, vbNullString))
            priorSourceScanVersion = CLng(Nz(priorModuleRows.fields(FIELD_SOURCE_SCAN_VERSION).value, 0))
            priorSourceScanCompletedOn = priorModuleRows.fields(FIELD_SOURCE_SCAN_COMPLETED_ON).value
            priorModuleLastScannedOn = priorModuleRows.fields(FIELD_LAST_SCANNED_ON).value
            priorSourceReadable = CBool(Nz(priorModuleRows.fields(FIELD_SOURCE_READABLE).value, False))
            priorSourceFingerprint = CStr(Nz(priorModuleRows.fields(FIELD_SOURCE_FINGERPRINT).value, vbNullString))
            priorQualifiedReferenceTokens = CStr(Nz(priorModuleRows.fields(FIELD_QUALIFIED_REFERENCE_TOKENS).value, vbNullString))
            priorSourceLineCount = CLng(Nz(priorModuleRows.fields(FIELD_SOURCE_LINE_COUNT).value, 0))
            priorDeclarationLineCount = CLng(Nz(priorModuleRows.fields(FIELD_DECLARATION_LINE_COUNT).value, 0))
            priorProcedureCount = CLng(Nz(priorModuleRows.fields(FIELD_PROCEDURE_COUNT).value, 0))
            priorPublicProcedureCount = CLng(Nz(priorModuleRows.fields(FIELD_PUBLIC_PROCEDURE_COUNT).value, 0))
            priorExternallyCallableCount = CLng(Nz(priorModuleRows.fields(FIELD_EXTERNALLY_CALLABLE_PROCEDURE_COUNT).value, 0))
            priorHasOptionPrivateModule = CBool(Nz(priorModuleRows.fields(FIELD_HAS_OPTION_PRIVATE_MODULE).value, False))
            priorHeaderReferenceDependencies = CStr(Nz(priorModuleRows.fields(FIELD_HEADER_REFERENCE_DEPENDENCIES).value, vbNullString))
            priorHeaderObjectDependencies = CStr(Nz(priorModuleRows.fields(FIELD_HEADER_DB_OBJECT_DEPENDENCIES).value, vbNullString))
            priorScanStatus = CStr(Nz(priorModuleRows.fields(FIELD_SCAN_STATUS).value, vbNullString))
            priorDiagnosticText = CStr(Nz(priorModuleRows.fields(FIELD_DIAGNOSTIC_TEXT).value, vbNullString))
        End If
        priorModuleRows.Close
        Set priorModuleRows = Nothing
    End If
    Set sourceModule = Nothing
    readErrNumber = 0
    readErrDescription = vbNullString
    On Error Resume Next
    Set sourceModule = CallByName(sourceComponent, "CodeModule", VbGet)
    If Err.Number = 0 And Not sourceModule Is Nothing Then
        sourceLineCount = CLng(CallByName(sourceModule, "CountOfLines", VbGet))
        declarationLineCount = CLng(CallByName(sourceModule, "CountOfDeclswarationLines", VbGet))
    End If
    readErrNumber = Err.Number
    readErrDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If readErrNumber = 0 And Not sourceModule Is Nothing Then
        stageTimer = Timer
        currentChatGPTID = ModuleGuardChatGPTVersionIDFromComponent(sourceComponent)
        headerInspectionSeconds = ElapsedSecondsFromTimer(stageTimer)
        If Len(currentChatGPTID) > 0 And priorInventoryFound And priorSourceReadable And Not IsNull(priorSourceScanCompletedOn) And priorSourceScanVersion = LOADED_VBA_SOURCE_SCAN_VERSION And StrComp(priorScannedChatGPTID, currentChatGPTID, vbBinaryCompare) = 0 Then
            reuseSourceInventory = True
            reuseByChatGPTID = True
            cacheReason = "valid ChatGPT ID unchanged and source-scan version current"
        End If
        If reuseSourceInventory = False And Len(currentChatGPTID) = 0 Then
            stageTimer = Timer
            CorePulseProgress "Loaded VBA project | " & projectName & "." & componentName & " | no valid ChatGPT ID | reading source once for fingerprint fallback"
            On Error Resume Next
            If sourceLineCount > 0 Then sourceText = CStr(CallByName(sourceModule, "Lines", VbGet, 1, sourceLineCount)) Else sourceText = vbNullString
            readErrNumber = Err.Number
            readErrDescription = Err.Description
            Err.Clear
            On Error GoTo 0
            sourceAcquisitionSeconds = ElapsedSecondsFromTimer(stageTimer)
            If readErrNumber = 0 Then
                componentSourceReadable = True
                stageTimer = Timer
                sourceFingerprint = ReferenceScanTextFingerprint(sourceText, "Loaded VBA project | " & projectName & "." & componentName & " | fingerprinting source")
                fingerprintSeconds = ElapsedSecondsFromTimer(stageTimer)
                If priorInventoryFound And priorSourceReadable And Not IsNull(priorSourceScanCompletedOn) And priorSourceScanVersion = LOADED_VBA_SOURCE_SCAN_VERSION And Len(priorSourceFingerprint) > 0 And StrComp(priorSourceFingerprint, sourceFingerprint, vbBinaryCompare) = 0 Then
                    reuseSourceInventory = True
                    cacheReason = "no valid ChatGPT ID; source fingerprint unchanged and source-scan version current"
                End If
            End If
        End If
    End If
    If reuseSourceInventory Then
        anySourceReadable = True
        componentSourceReadable = True
        sourceLineCount = priorSourceLineCount
        declarationLineCount = priorDeclarationLineCount
        moduleProcedureCount = priorProcedureCount
        modulePublicProcedureCount = priorPublicProcedureCount
        externallyCallableCount = priorExternallyCallableCount
        hasOptionPrivateModule = priorHasOptionPrivateModule
        sourceFingerprint = priorSourceFingerprint
        qualifiedReferenceTokens = priorQualifiedReferenceTokens
        headerReferenceDependencies = priorHeaderReferenceDependencies
        headerObjectDependencies = priorHeaderObjectDependencies
        moduleScanStatus = priorScanStatus
        moduleDiagnostic = priorDiagnosticText
        retainedProcedureRows = ReactivateLoadedVBAProcedureRows(db, moduleIdentity, priorModuleLastScannedOn, scannedOn)
        proceduresRetained = proceduresRetained + retainedProcedureRows
        procedureCount = procedureCount + moduleProcedureCount
        If rosterChanged Then
            retainedDependencyRows = ReactivateLoadedVBADependencyRows(db, moduleIdentity, priorModuleLastScannedOn, scannedOn, False)
            rebuiltDependencyRows = ResolveQualifiedReferenceTokens(dependencyRows, moduleRosterByName, projectIdentity, projectName, projectRole, moduleIdentity, componentName, qualifiedReferenceTokens, scannedOn)
            dependencyCount = dependencyCount + retainedDependencyRows + rebuiltDependencyRows
            modulesRosterReconciled = modulesRosterReconciled + 1
        Else
            retainedDependencyRows = ReactivateLoadedVBADependencyRows(db, moduleIdentity, priorModuleLastScannedOn, scannedOn, True)
            dependencyCount = dependencyCount + retainedDependencyRows
        End If
        If reuseByChatGPTID Then modulesReusedByChatGPTID = modulesReusedByChatGPTID + 1 Else modulesReusedByFingerprint = modulesReusedByFingerprint + 1
        CorePulseProgress "Loaded VBA project | " & projectName & "." & componentName & " | unchanged" & IIf(reuseByChatGPTID, " " & currentChatGPTID, " by fingerprint") & " | reusing previous source inventory"
        LogLoadedVBASourceCacheDecision projectName, componentName, True, cacheReason, priorScannedChatGPTID, currentChatGPTID
    ElseIf Not sourceModule Is Nothing And readErrNumber = 0 Then
        If Len(sourceText) = 0 And sourceLineCount > 0 Then
            stageTimer = Timer
            CorePulseProgress "Loaded VBA project | " & projectName & "." & componentName & " | source inventory invalid | reading source for rescan"
            On Error Resume Next
            sourceText = CStr(CallByName(sourceModule, "Lines", VbGet, 1, sourceLineCount))
            readErrNumber = Err.Number
            readErrDescription = Err.Description
            Err.Clear
            On Error GoTo 0
            sourceAcquisitionSeconds = sourceAcquisitionSeconds + ElapsedSecondsFromTimer(stageTimer)
        ElseIf sourceLineCount = 0 Then
            readErrNumber = 0
            sourceText = vbNullString
        End If
        If readErrNumber = 0 Then
            componentSourceReadable = True
            anySourceReadable = True
            modulesSourceRescanned = modulesSourceRescanned + 1
            If priorInventoryFound = False Then
                cacheReason = "no previous source inventory exists"
            ElseIf priorSourceScanVersion <> LOADED_VBA_SOURCE_SCAN_VERSION Then
                cacheReason = "source-scan algorithm version changed"
            ElseIf IsNull(priorSourceScanCompletedOn) Then
                cacheReason = "previous source scan did not complete successfully"
            ElseIf Len(currentChatGPTID) > 0 Then
                cacheReason = "ChatGPT ID changed or was not previously recorded for this source inventory"
            Else
                cacheReason = "source fingerprint changed"
            End If
            LogLoadedVBASourceCacheDecision projectName, componentName, False, cacheReason, priorScannedChatGPTID, currentChatGPTID
            CorePulseProgress "Loaded VBA project | " & projectName & "." & componentName & " | rescanning source | enumerating procedures"
            hasOptionPrivateModule = SourceHasOptionPrivateModule(sourceText)
            headerReferenceDependencies = ManagedHeaderSectionText(sourceText, "Reference Dependencies")
            headerObjectDependencies = ManagedHeaderSectionText(sourceText, "DB Object Dependencies")
            stageTimer = Timer
            ScanProcedureDeclarations projectIdentity, projectName, projectRole, moduleIdentity, componentName, componentTypeCode, hasOptionPrivateModule, sourceText, scannedOn, procedureRows, moduleProcedureCount, modulePublicProcedureCount, externallyCallableCount
            procedureScanSeconds = ElapsedSecondsFromTimer(stageTimer)
            procedureCount = procedureCount + moduleProcedureCount
            proceduresRebuilt = proceduresRebuilt + moduleProcedureCount
            CorePulseProgress "Loaded VBA project | " & projectName & "." & componentName & " | preparing dependency source"
            stageTimer = Timer
            dependencyText = VbaTextForDependencyScan(sourceText, "Loaded VBA project | " & projectName & "." & componentName & " | preparing dependency source")
            dependencyPreparationSeconds = ElapsedSecondsFromTimer(stageTimer)
            CorePulseProgress "Loaded VBA project | " & projectName & "." & componentName & " | extracting qualified reference tokens"
            stageTimer = Timer
            qualifiedReferenceTokens = QualifiedReferenceTokensFromDependencyText(dependencyText, "Loaded VBA project | " & projectName & "." & componentName & " | extracting qualified reference tokens")
            tokenExtractionSeconds = ElapsedSecondsFromTimer(stageTimer)
            If Len(headerReferenceDependencies) > 0 Then AddLoadedDependencyRow dependencyRows, projectIdentity, projectName, projectRole, moduleIdentity, componentName, "Declared Reference Dependency", headerReferenceDependencies, vbNullString, vbNullString, vbNullString, vbNullString, "Declared", False, vbNullString, scannedOn: dependencyCount = dependencyCount + 1
            If Len(headerObjectDependencies) > 0 Then AddLoadedDependencyRow dependencyRows, projectIdentity, projectName, projectRole, moduleIdentity, componentName, "Declared DB Object Dependency", headerObjectDependencies, vbNullString, vbNullString, vbNullString, vbNullString, "Declared", False, vbNullString, scannedOn: dependencyCount = dependencyCount + 1
            CorePulseProgress "Loaded VBA project | " & projectName & "." & componentName & " | resolving module dependencies"
            stageTimer = Timer
            rebuiltDependencyRows = ResolveQualifiedReferenceTokens(dependencyRows, moduleRosterByName, projectIdentity, projectName, projectRole, moduleIdentity, componentName, qualifiedReferenceTokens, scannedOn)
            dependencyResolutionSeconds = ElapsedSecondsFromTimer(stageTimer)
            dependencyCount = dependencyCount + rebuiltDependencyRows
            If Len(sourceFingerprint) = 0 Then
                CorePulseProgress "Loaded VBA project | " & projectName & "." & componentName & " | fingerprinting source"
                stageTimer = Timer
                sourceFingerprint = ReferenceScanTextFingerprint(sourceText, "Loaded VBA project | " & projectName & "." & componentName & " | fingerprinting source")
                fingerprintSeconds = fingerprintSeconds + ElapsedSecondsFromTimer(stageTimer)
            End If
            moduleScanStatus = "Source Scanned"
        Else
            allSourceReadable = False
            moduleScanStatus = "Metadata Only"
            moduleDiagnostic = "Source text was not readable. Error " & CStr(readErrNumber) & ": " & readErrDescription
        End If
    Else
        allSourceReadable = False
        moduleScanStatus = "Metadata Only"
        moduleDiagnostic = "Source text was not readable. Error " & CStr(readErrNumber) & ": " & readErrDescription
    End If
    stageTimer = Timer
    BeginInventoryUpsertRow moduleRows, moduleIdentity
    moduleRows.fields(FIELD_MODULE_IDENTITY).value = moduleIdentity
    moduleRows.fields(FIELD_PROJECT_IDENTITY).value = projectIdentity
    moduleRows.fields(FIELD_PROJECT_NAME).value = SharedValue_NormalizeAssignment(projectName)
    moduleRows.fields(FIELD_PROJECT_ROLE).value = projectRole
    moduleRows.fields(FIELD_MODULE_ORDER).value = moduleIndex
    moduleRows.fields(FIELD_MODULE_NAME).value = SharedValue_NormalizeAssignment(componentName)
    moduleRows.fields(FIELD_COMPONENT_TYPE_CODE).value = componentTypeCode
    moduleRows.fields(FIELD_COMPONENT_TYPE).value = ComponentTypeText
    moduleRows.fields(FIELD_SOURCE_LINE_COUNT).value = sourceLineCount
    moduleRows.fields(FIELD_DECLARATION_LINE_COUNT).value = declarationLineCount
    moduleRows.fields(FIELD_PROCEDURE_COUNT).value = moduleProcedureCount
    moduleRows.fields(FIELD_PUBLIC_PROCEDURE_COUNT).value = modulePublicProcedureCount
    moduleRows.fields(FIELD_EXTERNALLY_CALLABLE_PROCEDURE_COUNT).value = externallyCallableCount
    moduleRows.fields(FIELD_HAS_OPTION_PRIVATE_MODULE).value = hasOptionPrivateModule
    moduleRows.fields(FIELD_SOURCE_READABLE).value = componentSourceReadable
    If componentSourceReadable Then moduleRows.fields(FIELD_SOURCE_FINGERPRINT).value = SharedValue_NormalizeAssignment(sourceFingerprint) Else moduleRows.fields(FIELD_SOURCE_FINGERPRINT).value = Null
    If componentSourceReadable And (reuseSourceInventory Or moduleScanStatus = "Source Scanned") Then
        If Len(currentChatGPTID) > 0 Then moduleRows.fields(FIELD_SCANNED_CHAT_GPTID).value = currentChatGPTID Else moduleRows.fields(FIELD_SCANNED_CHAT_GPTID).value = Null
        moduleRows.fields(FIELD_SOURCE_SCAN_VERSION).value = LOADED_VBA_SOURCE_SCAN_VERSION
        If reuseSourceInventory Then moduleRows.fields(FIELD_SOURCE_SCAN_COMPLETED_ON).value = priorSourceScanCompletedOn Else moduleRows.fields(FIELD_SOURCE_SCAN_COMPLETED_ON).value = scannedOn
        moduleRows.fields(FIELD_QUALIFIED_REFERENCE_TOKENS).value = SharedValue_NormalizeAssignment(qualifiedReferenceTokens)
    Else
        moduleRows.fields(FIELD_SCANNED_CHAT_GPTID).value = Null
        moduleRows.fields(FIELD_SOURCE_SCAN_VERSION).value = Null
        moduleRows.fields(FIELD_SOURCE_SCAN_COMPLETED_ON).value = Null
        moduleRows.fields(FIELD_QUALIFIED_REFERENCE_TOKENS).value = Null
    End If
    moduleRows.fields(FIELD_HEADER_REFERENCE_DEPENDENCIES).value = SharedValue_NormalizeAssignment(headerReferenceDependencies)
    moduleRows.fields(FIELD_HEADER_DB_OBJECT_DEPENDENCIES).value = SharedValue_NormalizeAssignment(headerObjectDependencies)
    moduleRows.fields(FIELD_SCAN_STATUS).value = moduleScanStatus
    moduleRows.fields(FIELD_DIAGNOSTIC_TEXT).value = SharedValue_NormalizeAssignment(moduleDiagnostic)
    moduleRows.fields(FIELD_INVENTORY_STATUS).value = INVENTORY_STATUS_CURRENT
    moduleRows.fields(FIELD_LAST_SCAN_BATCH_ID).value = ReferenceScanBatchID(scannedOn, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAModules"))
    moduleRows.fields(FIELD_LAST_SCANNED_ON).value = scannedOn
    moduleRows.Update
    storageSeconds = ElapsedSecondsFromTimer(stageTimer)
    moduleCount = moduleCount + 1
    If componentSourceReadable And reuseSourceInventory = False And moduleScanStatus = "Source Scanned" Then
        LogLoadedVBAModuleTiming projectName, componentName, sourceLineCount, moduleProcedureCount, sourceAcquisitionSeconds, headerInspectionSeconds, procedureScanSeconds, dependencyPreparationSeconds, tokenExtractionSeconds, dependencyResolutionSeconds, fingerprintSeconds, storageSeconds, ElapsedSecondsFromTimer(moduleTimer)
    End If
    CorePulseProgress "Loaded VBA project | " & projectName & "." & componentName & " complete | status=" & moduleScanStatus & " | lines=" & CStr(sourceLineCount) & " | module procedures=" & CStr(moduleProcedureCount) & " | total procedures=" & CStr(procedureCount) & " | dependencies=" & CStr(dependencyCount)
End Sub

Private Function StoredLoadedVBAModuleRosterSignature(ByVal db As DAO.Database, ByVal projectIdentity As String) As String
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT " & SQL_FIELD_MODULE_ROSTER_SIGNATURE & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects") & "] WHERE " & SQL_FIELD_PROJECT_IDENTITY & "=" & SharedDao_SqlQuote(projectIdentity) & " AND " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & ";", dbOpenSnapshot)
    If Not rs.EOF Then StoredLoadedVBAModuleRosterSignature = CStr(Nz(rs.fields(FIELD_MODULE_ROSTER_SIGNATURE).value, vbNullString))
    rs.Close
End Function

Private Function LoadedVBAModuleRosterSignature(ByVal moduleRosterByName As Object) As String
    ' INTENTIONAL ORDERED VIEW: keys is the operation-local sortable name view of the authoritative moduleRosterByName lookup used only to make the roster signature order-independent.
    Dim keys As Variant
    Dim leftIndex As Long
    Dim rightIndex As Long
    Dim swapValue As Variant
    Dim moduleInfo As Variant
    Dim signatureText As String
    If moduleRosterByName Is Nothing Then Exit Function
    If moduleRosterByName.count = 0 Then LoadedVBAModuleRosterSignature = ReferenceScanTextFingerprint(vbNullString): Exit Function
    keys = moduleRosterByName.keys
    For leftIndex = LBound(keys) To UBound(keys) - 1
        For rightIndex = leftIndex + 1 To UBound(keys)
            If StrComp(CStr(keys(leftIndex)), CStr(keys(rightIndex)), vbTextCompare) > 0 Then
                swapValue = keys(leftIndex)
                keys(leftIndex) = keys(rightIndex)
                keys(rightIndex) = swapValue
            End If
        Next rightIndex
    Next leftIndex
    For leftIndex = LBound(keys) To UBound(keys)
        moduleInfo = moduleRosterByName(CStr(keys(leftIndex)))
        signatureText = signatureText & UCase$(CStr(moduleInfo(0))) & "|" & CStr(CLng(moduleInfo(1))) & vbLf
    Next leftIndex
    LoadedVBAModuleRosterSignature = ReferenceScanTextFingerprint(signatureText)
End Function

Private Function ReactivateLoadedVBAProcedureRows(ByVal db As DAO.Database, ByVal moduleIdentity As String, ByVal priorScannedOn As Variant, ByVal scannedOn As Date) As Long
    Dim priorBatchID As String
    If IsNull(priorScannedOn) Or IsEmpty(priorScannedOn) Then Exit Function
    priorBatchID = ReferenceScanBatchID(CDate(priorScannedOn), ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProcedures"))
    db.Execute "UPDATE [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProcedures") & "] SET " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & ", " & SQL_FIELD_LAST_SCAN_BATCH_ID & "=" & SharedDao_SqlQuote(ReferenceScanBatchID(scannedOn, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProcedures"))) & ", " & SQL_FIELD_LAST_SCANNED_ON & "=Now() WHERE " & SQL_FIELD_MODULE_IDENTITY & "=" & SharedDao_SqlQuote(moduleIdentity) & " AND " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_STALE) & " AND " & SQL_FIELD_LAST_SCAN_BATCH_ID & "=" & SharedDao_SqlQuote(priorBatchID) & ";", dbFailOnError
    ReactivateLoadedVBAProcedureRows = db.RecordsAffected
End Function

Private Function ReactivateLoadedVBADependencyRows(ByVal db As DAO.Database, ByVal moduleIdentity As String, ByVal priorScannedOn As Variant, ByVal scannedOn As Date, ByVal includeQualifiedReferences As Boolean) As Long
    Dim priorBatchID As String
    Dim criteriaText As String
    If IsNull(priorScannedOn) Or IsEmpty(priorScannedOn) Then Exit Function
    priorBatchID = ReferenceScanBatchID(CDate(priorScannedOn), ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBADependencies"))
    criteriaText = SQL_FIELD_SOURCE_MODULE_IDENTITY & "=" & SharedDao_SqlQuote(moduleIdentity) & " AND " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_STALE) & " AND " & SQL_FIELD_LAST_SCAN_BATCH_ID & "=" & SharedDao_SqlQuote(priorBatchID)
    If includeQualifiedReferences = False Then criteriaText = criteriaText & " AND " & SQL_FIELD_DEPENDENCY_TYPE & "<>" & SharedDao_SqlQuote("Qualified Module Reference")
    db.Execute "UPDATE [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBADependencies") & "] SET " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & ", " & SQL_FIELD_LAST_SCAN_BATCH_ID & "=" & SharedDao_SqlQuote(ReferenceScanBatchID(scannedOn, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBADependencies"))) & ", " & SQL_FIELD_LAST_SCANNED_ON & "=Now() WHERE " & criteriaText & ";", dbFailOnError
    ReactivateLoadedVBADependencyRows = db.RecordsAffected
End Function

Private Function QualifiedReferenceTokensFromDependencyText(ByVal dependencyText As String, Optional ByVal progressContext As String = vbNullString) As String
    ' INTENTIONAL DISTINCT FILTERED PARTITION: tokenKeys and orderedTokens derive the unique qualified-identifier candidates from one normalized module source for persistence as the module's sole roster-independent
    ' qualified-reference projection. Their lifetime ends when this extraction returns.
    Dim tokenKeys As Object
    Dim orderedTokens As New Collection
    Dim characterIndex As Long
    Dim tokenStart As Long
    Dim suffixStart As Long
    Dim tokenText As String
    Dim suffixText As String
    Dim currentCharacter As String
    Dim precedingCharacter As String
    Dim tokenItem As Variant
    Dim heartbeatTimer As Single
    Set tokenKeys = CreateObject("Scripting.Dictionary")
    tokenKeys.CompareMode = vbTextCompare
    heartbeatTimer = Timer
    characterIndex = 1
    Do While characterIndex <= Len(dependencyText)
        If (characterIndex And 255) = 0 Then PumpSetupResponsiveness
        If Len(progressContext) > 0 And ElapsedSecondsFromTimer(heartbeatTimer) >= 5# Then
            If SetupControlCancellationRequested() Then RequireSafeWorkContinue "QualifiedReferenceTokensFromDependencyText", progressContext
            CorePulseProgress progressContext & " | character " & CStr(characterIndex) & "/" & CStr(Len(dependencyText))
            heartbeatTimer = Timer
        End If
        currentCharacter = Mid$(dependencyText, characterIndex, 1)
        If VbaSourceIdentifierCharacter(currentCharacter) Then
            tokenStart = characterIndex
            Do While characterIndex <= Len(dependencyText) And VbaSourceIdentifierCharacter(Mid$(dependencyText, characterIndex, 1))
                characterIndex = characterIndex + 1
            Loop
            If characterIndex <= Len(dependencyText) And Mid$(dependencyText, characterIndex, 1) = "." Then
                tokenText = Mid$(dependencyText, tokenStart, characterIndex - tokenStart)
                If Len(tokenText) > 0 And tokenKeys.Exists(tokenText) = False Then
                    tokenKeys.Add tokenText, True
                    orderedTokens.Add tokenText
                End If
                ' Preserve the prior matcher exactly when a non-ASCII character preceded a possible module-name suffix: the old ASCII boundary test treated that position as a valid identifier boundary.
                For suffixStart = 2 To Len(tokenText)
                    precedingCharacter = Mid$(tokenText, suffixStart - 1, 1)
                    If SharedVba_IsIdentifierChar(precedingCharacter) = False Then
                        suffixText = Mid$(tokenText, suffixStart)
                        If Len(suffixText) > 0 And tokenKeys.Exists(suffixText) = False Then
                            tokenKeys.Add suffixText, True
                            orderedTokens.Add suffixText
                        End If
                    End If
                Next suffixStart
            End If
        Else
            characterIndex = characterIndex + 1
        End If
    Loop
    For Each tokenItem In orderedTokens
        QualifiedReferenceTokensFromDependencyText = SharedText_AppendNonBlankLine(QualifiedReferenceTokensFromDependencyText, CStr(tokenItem))
    Next tokenItem
End Function


Private Function ResolveQualifiedReferenceTokens(ByVal dependencyRows As DAO.Recordset, ByVal moduleRosterByName As Object, ByVal projectIdentity As String, ByVal projectName As String, ByVal projectRole As String, ByVal sourceModuleIdentity As String, ByVal sourceModuleName As String, ByVal qualifiedReferenceTokens As String, ByVal scannedOn As Date) As Long
    ' INTENTIONAL FILTERED PARTITION VIEW: tokens is the operation-local line split of the module row's authoritative QualifiedReferenceTokens projection used only for current-roster resolution.
    Dim tokens As Variant
    Dim tokenItem As Variant
    Dim tokenText As String
    Dim targetInfo As Variant
    Dim heartbeatTimer As Single
    Dim tokenNumber As Long
    Dim tokenCount As Long
    If Len(Trim$(qualifiedReferenceTokens)) = 0 Then Exit Function
    tokens = Split(Replace(qualifiedReferenceTokens, vbCrLf, vbLf), vbLf)
    tokenCount = UBound(tokens) - LBound(tokens) + 1
    heartbeatTimer = Timer
    For Each tokenItem In tokens
        PumpSetupResponsiveness
        tokenNumber = tokenNumber + 1
        If ElapsedSecondsFromTimer(heartbeatTimer) >= 5# Then
            If SetupControlCancellationRequested() Then RequireSafeWorkContinue "ResolveQualifiedReferenceTokens", projectName & "." & sourceModuleName
            CorePulseProgress "Loaded VBA project | " & projectName & "." & sourceModuleName & " | resolving qualified dependencies | token " & CStr(tokenNumber) & "/" & CStr(tokenCount)
            heartbeatTimer = Timer
        End If
        tokenText = Trim$(CStr(tokenItem))
        If Len(tokenText) > 0 And StrComp(tokenText, sourceModuleName, vbTextCompare) <> 0 Then
            If moduleRosterByName.Exists(tokenText) Then
                targetInfo = moduleRosterByName(tokenText)
                AddLoadedDependencyRow dependencyRows, projectIdentity, projectName, projectRole, sourceModuleIdentity, sourceModuleName, "Qualified Module Reference", CStr(targetInfo(0)), projectIdentity, projectName, CStr(targetInfo(2)), CStr(targetInfo(0)), "Resolved - Exact Module", True, vbNullString, scannedOn
                ResolveQualifiedReferenceTokens = ResolveQualifiedReferenceTokens + 1
            End If
        End If
    Next tokenItem
End Function

Private Sub LogLoadedVBASourceCacheDecision(ByVal projectName As String, ByVal ModuleName As String, ByVal reused As Boolean, ByVal reasonText As String, ByVal oldChatGPTID As String, ByVal newChatGPTID As String)
    Dim detailText As String
    detailText = "Project: " & projectName & vbCrLf & "Module: " & ModuleName & vbCrLf & "Reason: " & reasonText
    If Len(oldChatGPTID) > 0 Or Len(newChatGPTID) > 0 Then detailText = detailText & vbCrLf & "Old scanned ChatGPT ID: " & IIf(Len(oldChatGPTID) = 0, "(none)", oldChatGPTID) & vbCrLf & "Current ChatGPT ID: " & IIf(Len(newChatGPTID) = 0, "(none)", newChatGPTID)
    LogTraceMessage "Info", "References", "ScanLoadedProjectModules", IIf(reused, "Module source inventory reused.", "Module source rescan required."), detailText, "Scan DB Inventory - Loaded VBA Source", 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAModules"), "Table", "CodeDb", "References"
End Sub

Private Sub LogLoadedVBAModuleTiming(ByVal projectName As String, ByVal ModuleName As String, ByVal sourceLineCount As Long, ByVal procedureCount As Long, ByVal sourceAcquisitionSeconds As Double, ByVal headerInspectionSeconds As Double, ByVal procedureScanSeconds As Double, ByVal dependencyPreparationSeconds As Double, ByVal tokenExtractionSeconds As Double, ByVal dependencyResolutionSeconds As Double, ByVal fingerprintSeconds As Double, ByVal storageSeconds As Double, ByVal totalSeconds As Double)
    Dim detailText As String
    detailText = "Project: " & projectName & vbCrLf & "Module: " & ModuleName & vbCrLf & "Lines: " & CStr(sourceLineCount) & vbCrLf & "Procedures: " & CStr(procedureCount) & vbCrLf & "Source acquisition: " & FormattedElapsedSeconds(sourceAcquisitionSeconds) & vbCrLf & "Header/ChatGPT ID inspection: " & FormattedElapsedSeconds(headerInspectionSeconds) & vbCrLf & "Procedure scan: " & FormattedElapsedSeconds(procedureScanSeconds) & vbCrLf & "Dependency source preparation: " & FormattedElapsedSeconds(dependencyPreparationSeconds) & vbCrLf & "Qualified-reference extraction: " & FormattedElapsedSeconds(tokenExtractionSeconds) & vbCrLf & "Dependency resolution: " & FormattedElapsedSeconds(dependencyResolutionSeconds) & vbCrLf & "Fingerprint: " & FormattedElapsedSeconds(fingerprintSeconds) & vbCrLf & "Storage: " & FormattedElapsedSeconds(storageSeconds) & vbCrLf & "Total: " & _
                FormattedElapsedSeconds(totalSeconds)
    LogTraceMessage "Info", "References", "ScanLoadedProjectModules", "Loaded VBA module timing.", detailText, "Scan DB Inventory - Loaded VBA Source", 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAModules"), "Table", "CodeDb", "References"
End Sub

Private Function SourceHasOptionPrivateModule(ByVal sourceText As String) As Boolean
    Dim lines As Variant
    Dim lineIndex As Long
    Dim workText As String
    lines = Split(Replace(sourceText, vbCrLf, vbLf), vbLf)
    For lineIndex = LBound(lines) To UBound(lines)
        PumpSetupResponsiveness
        workText = LCase$(Trim$(StripVbaComment(CStr(lines(lineIndex)))))
        If StrComp(workText, "option private module", vbBinaryCompare) = 0 Then SourceHasOptionPrivateModule = True: Exit Function
    Next lineIndex
End Function

Private Function SafeVBProjectText(ByVal projectItem As Object, ByVal propertyName As String) As String
    On Error GoTo Missing
    SafeVBProjectText = CStr(CallByName(projectItem, propertyName, VbGet))
    Exit Function
Missing:
    Err.Clear
End Function

Private Function SafeVBProjectLong(ByVal projectItem As Object, ByVal propertyName As String, ByVal DefaultValue As Long) As Long
    On Error GoTo Missing
    SafeVBProjectLong = CLng(CallByName(projectItem, propertyName, VbGet))
    Exit Function
Missing:
    Err.Clear
    SafeVBProjectLong = DefaultValue
End Function

Private Function SafeVBProjectComponentCount(ByVal projectItem As Object, ByRef componentsReadable As Boolean, ByRef diagnosticText As String) As Long
    On Error GoTo Missing
    SafeVBProjectComponentCount = CLng(projectItem.VBComponents.count)
    componentsReadable = True
    Exit Function
Missing:
    diagnosticText = "VBComponents collection could not be read. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    Err.Clear
End Function

Private Function VBProjectObjectsMatch(ByVal leftProject As Object, ByVal rightProject As Object) As Boolean
    On Error GoTo ExitHere
    If leftProject Is Nothing Or rightProject Is Nothing Then Exit Function
    VBProjectObjectsMatch = (leftProject Is rightProject)
ExitHere:
    Err.Clear
End Function

Private Sub ScanProcedureDeclarations(ByVal projectIdentity As String, ByVal projectName As String, ByVal projectRole As String, ByVal moduleIdentity As String, ByVal ModuleName As String, ByVal componentTypeCode As Long, ByVal hasOptionPrivateModule As Boolean, ByVal sourceText As String, ByVal scannedOn As Date, ByVal procedureRows As DAO.Recordset, ByRef procedureCount As Long, ByRef publicProcedureCount As Long, ByRef externallyCallableCount As Long)
    Dim lines As Variant
    Dim lineIndex As Long
    Dim declarationText As String
    Dim ProcedureName As String
    Dim procedureKind As String
    Dim procedureScope As String
    Dim isPublic As Boolean
    Dim isExternallyCallable As Boolean
    Dim procedureIdentity As String
    Dim heartbeatTimer As Single
    lines = Split(Replace(sourceText, vbCrLf, vbLf), vbLf)
    heartbeatTimer = Timer
    For lineIndex = LBound(lines) To UBound(lines)
        PumpSetupResponsiveness
        If ElapsedSecondsFromTimer(heartbeatTimer) >= 5# Then
            If SetupControlCancellationRequested() Then RequireSafeWorkContinue "ScanProcedureDeclarations", projectName & "." & ModuleName
            CorePulseProgress "Loaded VBA project | " & projectName & "." & ModuleName & " | enumerating procedures | line " & CStr(lineIndex - LBound(lines) + 1) & "/" & CStr(UBound(lines) - LBound(lines) + 1)
            heartbeatTimer = Timer
        End If
        declarationText = Trim$(CStr(lines(lineIndex)))
        If ProcedureDeclarationInfo(declarationText, componentTypeCode, hasOptionPrivateModule, ProcedureName, procedureKind, procedureScope, isPublic, isExternallyCallable) Then
            procedureIdentity = "P|" & ReferenceScanTextFingerprint(moduleIdentity & "|" & ProcedureName & "|" & procedureKind)
            BeginInventoryUpsertRow procedureRows, procedureIdentity
            procedureRows.fields(FIELD_PROCEDURE_IDENTITY).value = procedureIdentity
            procedureRows.fields(FIELD_PROJECT_IDENTITY).value = projectIdentity
            procedureRows.fields(FIELD_PROJECT_NAME).value = SharedValue_NormalizeAssignment(projectName)
            procedureRows.fields(FIELD_PROJECT_ROLE).value = projectRole
            procedureRows.fields(FIELD_MODULE_IDENTITY).value = moduleIdentity
            procedureRows.fields(FIELD_MODULE_NAME).value = SharedValue_NormalizeAssignment(ModuleName)
            procedureRows.fields(FIELD_PROCEDURE_ORDER).value = lineIndex + 1
            procedureRows.fields(FIELD_PROCEDURE_NAME).value = ProcedureName
            procedureRows.fields(FIELD_PROCEDURE_NAME_KEY).value = UCase$(ProcedureName)
            procedureRows.fields(FIELD_PROCEDURE_KIND).value = procedureKind
            procedureRows.fields(FIELD_PROCEDURE_SCOPE).value = procedureScope
            procedureRows.fields(FIELD_IS_PUBLIC).value = isPublic
            procedureRows.fields(FIELD_IS_EXTERNALLY_CALLABLE).value = isExternallyCallable
            procedureRows.fields(FIELD_HAS_OPTION_PRIVATE_MODULE).value = hasOptionPrivateModule
            procedureRows.fields(FIELD_DECLARATION_TEXT).value = Left$(declarationText, 1000)
            procedureRows.fields(FIELD_INVENTORY_STATUS).value = INVENTORY_STATUS_CURRENT
            procedureRows.fields(FIELD_LAST_SCAN_BATCH_ID).value = ReferenceScanBatchID(scannedOn, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProcedures"))
            procedureRows.fields(FIELD_LAST_SCANNED_ON).value = scannedOn
            procedureRows.Update
            procedureCount = procedureCount + 1
            PulseProgressEvery procedureCount, 25, "Loaded VBA project | " & projectName & "." & ModuleName & " | procedure " & ProcedureName & " | module procedures=" & CStr(procedureCount) & " | line=" & CStr(lineIndex + 1), 0, "ReferenceLoadedProjectProcedureScan", "References"
            If isPublic Then publicProcedureCount = publicProcedureCount + 1
            If isExternallyCallable Then externallyCallableCount = externallyCallableCount + 1
        End If
    Next lineIndex
End Sub

Private Function ProcedureDeclarationInfo(ByVal sourceLine As String, ByVal componentTypeCode As Long, ByVal hasOptionPrivateModule As Boolean, ByRef ProcedureName As String, ByRef procedureKind As String, ByRef procedureScope As String, ByRef isPublic As Boolean, ByRef isExternallyCallable As Boolean) As Boolean
    Dim workText As String
    Dim lowerText As String
    Dim nameStart As Long
    ProcedureName = vbNullString
    procedureKind = vbNullString
    procedureScope = vbNullString
    isPublic = False
    isExternallyCallable = False
    workText = Trim$(StripVbaComment(sourceLine))
    If Len(workText) = 0 Then Exit Function
    lowerText = LCase$(workText)
    procedureScope = "Implicit Public"
    If Left$(lowerText, 7) = "public " Then procedureScope = "Public": workText = Trim$(Mid$(workText, 8)): lowerText = LCase$(workText)
    If Left$(lowerText, 8) = "private " Then procedureScope = "Private": workText = Trim$(Mid$(workText, 9)): lowerText = LCase$(workText)
    If Left$(lowerText, 7) = "friend " Then procedureScope = "Friend": workText = Trim$(Mid$(workText, 8)): lowerText = LCase$(workText)
    If Left$(lowerText, 7) = "static " Then workText = Trim$(Mid$(workText, 8)): lowerText = LCase$(workText)
    If Left$(lowerText, 16) = "declare ptrsafe " Then workText = Trim$(Mid$(workText, 17)): lowerText = LCase$(workText): procedureKind = "Declare "
    If Left$(lowerText, 8) = "declare " Then workText = Trim$(Mid$(workText, 9)): lowerText = LCase$(workText): procedureKind = "Declare "
    If Left$(lowerText, 4) = "sub " Then procedureKind = procedureKind & "Sub": nameStart = 5
    If Left$(lowerText, 9) = "function " Then procedureKind = procedureKind & "Function": nameStart = 10
    If Left$(lowerText, 13) = "property get " Then procedureKind = "Property Get": nameStart = 14
    If Left$(lowerText, 13) = "property let " Then procedureKind = "Property Let": nameStart = 14
    If Left$(lowerText, 13) = "property set " Then procedureKind = "Property Set": nameStart = 14
    If nameStart = 0 Then Exit Function
    ProcedureName = LeadingIdentifier(Mid$(workText, nameStart))
    If Len(ProcedureName) = 0 Then Exit Function
    isPublic = (StrComp(procedureScope, "Private", vbTextCompare) <> 0 And StrComp(procedureScope, "Friend", vbTextCompare) <> 0)
    isExternallyCallable = (isPublic And componentTypeCode = 1 And hasOptionPrivateModule = False)
    ProcedureDeclarationInfo = True
End Function


Private Function VbaTextForDependencyScan(ByVal sourceText As String, Optional ByVal progressContext As String = vbNullString) As String
    Dim lines As Variant
    ' INTENTIONAL BOUNDED CONSTRUCTION BUFFER: normalizedLines holds one completed normalized result per physical source line so this routine can Join once instead of growing a module-sized string character by
    ' character.
    Dim normalizedLines() As String
    Dim lineIndex As Long
    Dim sourceLine As String
    Dim lineBuffer As String
    Dim outputLength As Long
    Dim characterIndex As Long
    Dim currentCharacter As String
    Dim insideString As Boolean
    Dim heartbeatTimer As Single
    lines = Split(Replace(sourceText, vbCrLf, vbLf), vbLf)
    ReDim normalizedLines(LBound(lines) To UBound(lines))
    heartbeatTimer = Timer
    For lineIndex = LBound(lines) To UBound(lines)
        PumpSetupResponsiveness
        If Len(progressContext) > 0 And ElapsedSecondsFromTimer(heartbeatTimer) >= 5# Then
            If SetupControlCancellationRequested() Then RequireSafeWorkContinue "VbaTextForDependencyScan", progressContext
            CorePulseProgress progressContext & " | line " & CStr(lineIndex - LBound(lines) + 1) & "/" & CStr(UBound(lines) - LBound(lines) + 1)
            heartbeatTimer = Timer
        End If
        sourceLine = CStr(lines(lineIndex))
        lineBuffer = Space$(Len(sourceLine))
        outputLength = Len(sourceLine)
        insideString = False
        characterIndex = 1
        Do While characterIndex <= Len(sourceLine)
            If (characterIndex And 255) = 0 Then PumpSetupResponsiveness
            currentCharacter = Mid$(sourceLine, characterIndex, 1)
            If currentCharacter = Chr$(34) Then
                If insideString And characterIndex < Len(sourceLine) And Mid$(sourceLine, characterIndex + 1, 1) = Chr$(34) Then
                    characterIndex = characterIndex + 2
                Else
                    insideString = Not insideString
                    characterIndex = characterIndex + 1
                End If
            ElseIf insideString = False And VbaRemCommentStartsAt(sourceLine, characterIndex) Then
                outputLength = characterIndex - 1
                Exit Do
            ElseIf currentCharacter = "'" And insideString = False Then
                outputLength = characterIndex - 1
                Exit Do
            ElseIf insideString Then
                characterIndex = characterIndex + 1
            Else
                Mid$(lineBuffer, characterIndex, 1) = LCase$(currentCharacter)
                characterIndex = characterIndex + 1
            End If
        Loop
        If outputLength > 0 Then normalizedLines(lineIndex) = Left$(lineBuffer, outputLength) Else normalizedLines(lineIndex) = vbNullString
    Next lineIndex
    VbaTextForDependencyScan = Join(normalizedLines, vbLf) & vbLf
End Function

Private Function ManagedHeaderSectionText(ByVal sourceText As String, ByVal sectionName As String) As String
    Dim lines As Variant
    Dim lineIndex As Long
    Dim sourceLine As String
    Dim commentText As String
    Dim sectionPrefix As String
    Dim collecting As Boolean
    lines = Split(Replace(sourceText, vbCrLf, vbLf), vbLf)
    sectionPrefix = LCase$(sectionName) & ":"
    For lineIndex = LBound(lines) To UBound(lines)
        PumpSetupResponsiveness
        sourceLine = Trim$(CStr(lines(lineIndex)))
        If Left$(sourceLine, 1) <> "'" Then
            If collecting Then Exit For
        Else
            commentText = Trim$(Mid$(sourceLine, 2))
            If collecting = False Then
                If LCase$(Left$(commentText, Len(sectionPrefix))) = sectionPrefix Then
                    collecting = True
                    commentText = Trim$(Mid$(commentText, Len(sectionPrefix) + 1))
                    If Len(commentText) > 0 Then ManagedHeaderSectionText = commentText
                End If
            Else
                If IsManagedHeaderTopLevelLine(commentText) Then Exit For
                If Len(commentText) > 0 Then ManagedHeaderSectionText = SharedText_AppendNonBlankLine(ManagedHeaderSectionText, commentText)
            End If
        End If
    Next lineIndex
End Function

Private Function SafeVBComponentText(ByVal componentItem As Object, ByVal propertyName As String) As String
    On Error GoTo ExitHere
    SafeVBComponentText = CStr(CallByName(componentItem, propertyName, VbGet))
ExitHere:
End Function

Private Function VBComponentTypeText(ByVal componentTypeCode As Long) As String
    Select Case componentTypeCode
        Case 1: VBComponentTypeText = "Standard Module"
        Case 2: VBComponentTypeText = "Class Module"
        Case 3: VBComponentTypeText = "UserForm"
        Case 11: VBComponentTypeText = "ActiveX Designer"
        Case 100: VBComponentTypeText = "Document Module"
        Case Else: VBComponentTypeText = "Unknown (" & CStr(componentTypeCode) & ")"
    End Select
End Function

Private Function ReferenceScanTextFingerprint(ByVal sourceText As String, Optional ByVal progressContext As String = vbNullString) As String
    Const HASH_MODULUS As Double = 2147483647#
    Dim forwardHash As Double
    Dim reverseHash As Double
    Dim characterIndex As Long
    Dim characterCode As Long
    Dim heartbeatTimer As Single
    forwardHash = 5381#
    reverseHash = 5381#
    heartbeatTimer = Timer
    For characterIndex = 1 To Len(sourceText)
        If (characterIndex And 255) = 0 Then PumpSetupResponsiveness
        If Len(progressContext) > 0 And ElapsedSecondsFromTimer(heartbeatTimer) >= 5# Then
            If SetupControlCancellationRequested() Then RequireSafeWorkContinue "ReferenceScanTextFingerprint", progressContext
            CorePulseProgress progressContext & " | character " & CStr(characterIndex) & "/" & CStr(Len(sourceText))
            heartbeatTimer = Timer
        End If
        characterCode = AscW(Mid$(sourceText, characterIndex, 1))
        If characterCode < 0 Then characterCode = characterCode + 65536
        forwardHash = (forwardHash * 33# + characterCode) - Int((forwardHash * 33# + characterCode) / HASH_MODULUS) * HASH_MODULUS
        characterCode = AscW(Mid$(sourceText, Len(sourceText) - characterIndex + 1, 1))
        If characterCode < 0 Then characterCode = characterCode + 65536
        reverseHash = (reverseHash * 33# + characterCode) - Int((reverseHash * 33# + characterCode) / HASH_MODULUS) * HASH_MODULUS
    Next characterIndex
    ReferenceScanTextFingerprint = Right$("00000000" & Hex$(CLng(forwardHash)), 8) & Right$("00000000" & Hex$(CLng(reverseHash)), 8)
End Function

Private Function RegistryDefaultString(ByVal registryProvider As Object, ByVal hiveValue As Long, ByVal keyPath As String) As String
    On Error GoTo NoValue
    Dim resultCode As Long
    Dim valueData As Variant
    resultCode = registryProvider.GetStringValue(hiveValue, keyPath, vbNullString, valueData)
    If resultCode <> 0 Then
        resultCode = registryProvider.GetExpandedStringValue(hiveValue, keyPath, vbNullString, valueData)
        If resultCode <> 0 Then Exit Function
    End If
    If IsNull(valueData) Or IsEmpty(valueData) Then Exit Function
    RegistryDefaultString = CStr(valueData)
    Exit Function
NoValue:
    RegistryDefaultString = vbNullString
End Function

Private Function RegistryNamedString(ByVal registryProvider As Object, ByVal hiveValue As Long, ByVal keyPath As String, ByVal valueName As String) As String
    On Error GoTo NoValue
    Dim resultCode As Long
    Dim valueData As Variant
    resultCode = registryProvider.GetStringValue(hiveValue, keyPath, valueName, valueData)
    If resultCode <> 0 Then
        resultCode = registryProvider.GetExpandedStringValue(hiveValue, keyPath, valueName, valueData)
        If resultCode <> 0 Then Exit Function
    End If
    If IsNull(valueData) Or IsEmpty(valueData) Then Exit Function
    RegistryNamedString = CStr(valueData)
    Exit Function
NoValue:
    RegistryNamedString = vbNullString
End Function

Private Function SafeAddInText(ByVal addInItem As Object, ByVal propertyName As String) As String
    On Error GoTo Missing
    SafeAddInText = CStr(CallByName(addInItem, propertyName, VbGet))
    Exit Function
Missing:
    Err.Clear
End Function

Private Function IsManagedHeaderTopLevelLine(ByVal commentText As String) As Boolean
    Dim labelText As Variant
    For Each labelText In Array("Purpose:", "Author:", "Support:", "Code Category:", "Requirements:", "Reference Dependencies:", "DB Object Dependencies:", "Usage:", "Version:", "Notes:", "Core Rules Enforced Here:", "Design Map:", "Configuration:", "Behavior:", "Code Status:", "Todo:", "Reporting Bugs:", "Module Level Bugs:", "Entry Points:")
        PumpSetupResponsiveness
        If LCase$(Left$(commentText, Len(CStr(labelText)))) = LCase$(CStr(labelText)) Then IsManagedHeaderTopLevelLine = True: Exit Function
    Next labelText
End Function

Private Function SafeVBComponentLong(ByVal componentItem As Object, ByVal propertyName As String, ByVal DefaultValue As Long) As Long
    On Error GoTo UseDefault
    SafeVBComponentLong = CLng(CallByName(componentItem, propertyName, VbGet))
    Exit Function
UseDefault:
    SafeVBComponentLong = DefaultValue
End Function

' =============================================================================================================================================================================
' TRADITIONAL / VBE / COM ADD-IN INVENTORY PROCEDURES
' =============================================================================================================================================================================

Private Function ReadTraditionalAccessAddInRegistrations(ByVal projectFilePath As String, ByRef usysRegInfoPresent As Boolean, ByRef diagnosticText As String) As Collection
    'Canonical bounded reader for traditional Access add-in registration metadata; each project database is opened once and its USysRegInfo rows are enumerated once.
    On Error GoTo ReadFailed
    ' INTENTIONAL OPERATION SNAPSHOT: registrations captures one coherent source state for this invocation of ReadTraditionalAccessAddInRegistrations; validity ends with the operation rather than becoming a live
    ' data authority.
    ' INTENTIONAL KEYED LOOKUP: registrationKeys provides keyed membership/lookup over source rows or state already read by ReadTraditionalAccessAddInRegistrations; it is a derived access path, not a second
    ' authority.
    Dim registrations As New Collection
    Dim externalDb As DAO.Database
    Dim rs As DAO.Recordset
    Dim registrationKeys As New Collection
    Dim seenKeys As Object
    Dim librariesByKey As Object
    Dim expressionsByKey As Object
    Dim subkeyText As String
    Dim valueName As String
    Dim rawValue As String
    Dim item As Variant
    projectFilePath = Trim$(projectFilePath)
    If Len(projectFilePath) = 0 Or ReferenceFileExists(projectFilePath) = False Then
        Set ReadTraditionalAccessAddInRegistrations = registrations
        Exit Function
    End If
    Set externalDb = DBEngine.Workspaces(0).OpenDatabase(projectFilePath, False, True)
    If DaoTableExists(externalDb, "USysRegInfo") = False Then GoTo ReadComplete
    usysRegInfoPresent = True
    ' INTENTIONAL KEYED LOOKUP: seenKeys, librariesByKey, expressionsByKey provides keyed membership/lookup over source rows or state already read by ReadTraditionalAccessAddInRegistrations; it is a derived access
    ' path, not a second authority.
    Set seenKeys = CreateObject("Scripting.Dictionary")
    seenKeys.CompareMode = vbTextCompare
    Set librariesByKey = CreateObject("Scripting.Dictionary")
    librariesByKey.CompareMode = vbTextCompare
    Set expressionsByKey = CreateObject("Scripting.Dictionary")
    expressionsByKey.CompareMode = vbTextCompare
    Set rs = externalDb.OpenRecordset("SELECT " & SQL_FIELD_SUBKEY & "," & SQL_FIELD_VAL_NAME & "," & SQL_FIELD_VALUE & " FROM [USysRegInfo] ORDER BY " & SQL_FIELD_SUBKEY & "," & SQL_FIELD_VAL_NAME & ";", dbOpenSnapshot)
    Do While rs.EOF = False
        PumpSetupResponsiveness
        subkeyText = Trim$(Nz(rs.fields(FIELD_SUBKEY).value, vbNullString))
        valueName = Trim$(Nz(rs.fields(FIELD_VAL_NAME).value, vbNullString))
        rawValue = Nz(rs.fields(FIELD_VALUE).value, vbNullString)
        If Len(subkeyText) > 0 Then
            If seenKeys.Exists(subkeyText) = False Then
                seenKeys.Add subkeyText, True
                registrationKeys.Add subkeyText
                librariesByKey.Add subkeyText, vbNullString
                expressionsByKey.Add subkeyText, vbNullString
            End If
            Select Case UCase$(valueName)
                Case "LIBRARY"
                    librariesByKey(subkeyText) = rawValue
                Case "EXPRESSION"
                    expressionsByKey(subkeyText) = rawValue
            End Select
        End If
        rs.MoveNext
    Loop
    For Each item In registrationKeys
        subkeyText = CStr(item)
        If Len(CStr(expressionsByKey(subkeyText))) > 0 Or InStr(1, subkeyText, "\Menu Add-ins\", vbTextCompare) > 0 Then
            registrations.Add Array(subkeyText, CStr(librariesByKey(subkeyText)), CStr(expressionsByKey(subkeyText)))
        End If
    Next item
ReadComplete:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not externalDb Is Nothing Then externalDb.Close
    On Error GoTo 0
    Set ReadTraditionalAccessAddInRegistrations = registrations
    Exit Function
ReadFailed:
    diagnosticText = "USysRegInfo registration metadata could not be read from the add-in database. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not externalDb Is Nothing Then externalDb.Close
    On Error GoTo 0
    Set ReadTraditionalAccessAddInRegistrations = registrations
End Function


Private Function IsTraditionalAccessAddInDatabasePath(ByVal projectFilePath As String) As Boolean
    Dim extensionPosition As Long
    Dim extensionText As String
    extensionPosition = InStrRev(projectFilePath, ".")
    If extensionPosition = 0 Then Exit Function
    extensionText = LCase$(Mid$(projectFilePath, extensionPosition + 1))
    IsTraditionalAccessAddInDatabasePath = extensionText = "accda" Or extensionText = "mda"
End Function

Private Function ResolveTraditionalAccessAddInLibraryPath(ByVal addInLibrary As String) As String
    Dim candidatePath As String
    Dim accessDirectory As String
    addInLibrary = ExpandReferencePath(addInLibrary)
    If Len(addInLibrary) = 0 Then Exit Function
    If StrComp(Left$(addInLibrary, 9), "|ACCDIR|\", vbTextCompare) = 0 Then
        accessDirectory = CStr(Application.SysCmd(acSysCmdAccessDir))
        candidatePath = accessDirectory & Mid$(addInLibrary, 10)
    ElseIf StrComp(Left$(addInLibrary, 7), "ACCDIR\", vbTextCompare) = 0 Then
        accessDirectory = CStr(Application.SysCmd(acSysCmdAccessDir))
        candidatePath = accessDirectory & Mid$(addInLibrary, 8)
    End If
    If Len(candidatePath) > 0 Then
        candidatePath = ExpandReferencePath(candidatePath)
        ResolveTraditionalAccessAddInLibraryPath = candidatePath
        Exit Function
    End If
    If ReferenceFileExists(addInLibrary) Then
        ResolveTraditionalAccessAddInLibraryPath = addInLibrary
        Exit Function
    End If
    If InStr(addInLibrary, "\") = 0 And InStr(addInLibrary, "/") = 0 Then
        candidatePath = ExpandReferencePath(Environ$("APPDATA") & "\Microsoft\AddIns\" & addInLibrary)
        If ReferenceFileExists(candidatePath) Then
            ResolveTraditionalAccessAddInLibraryPath = candidatePath
            Exit Function
        End If
    End If
    ResolveTraditionalAccessAddInLibraryPath = addInLibrary
End Function

Private Function ScanTraditionalAccessAddInsCore(Optional ByVal correlateEntryPoints As Boolean = True) As Long
    'Canonical membership authority for traditional Access add-ins: installed Access Menu Add-Ins registry registrations, independent of VBE project load state.
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim ws As DAO.workspace
    Dim accessAddInRows As DAO.Recordset
    Dim registryProvider As Object
    Dim registryRoots As Variant
    Dim rootInfo As Variant
    Dim registrationNames As Variant
    Dim registrationName As Variant
    Dim registrations As Collection
    Dim registration As Variant
    Dim registryPath As String
    Dim registrationSubkey As String
    Dim addInLibrary As String
    Dim addInFilePath As String
    Dim expressionText As String
    Dim sourceRegistrationName As String
    Dim ProcedureName As String
    Dim addInIdentity As String
    Dim scanStatus As String
    Dim diagnosticText As String
    Dim readDiagnostic As String
    Dim scannedOn As Date
    Dim fileExistsValue As Boolean
    Dim usysRegInfoPresent As Boolean
    Dim sourceRegistrationMatched As Boolean
    Dim transactionActive As Boolean
    Dim addInOrder As Long
    Dim namePosition As Long
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Set db = ADCDb()
    Set ws = DBEngine.Workspaces(0)
    Set registryProvider = ReferenceRegistryProvider()
    registryRoots = AccessMenuAddInRegistryRoots()
    scannedOn = Now
    ws.BeginTrans
    transactionActive = True
    PrepareInventoryRowsForScan db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AccessAddIns")
    Set accessAddInRows = OpenInventoryUpsertRecordset(db, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AccessAddIns"), INDEX_UX_ACCESS_ADD_INS_IDENTITY)
    For Each rootInfo In registryRoots
        PumpSetupResponsiveness
        If SetupControlCancellationRequested() Then RequireSafeWorkContinue "ScanDBInventoryAddInsAndLoadedVBAProjects", "Installed Access add-in registry scan boundary"
        registrationNames = RegistrySubkeys(registryProvider, CLng(rootInfo(0)), CStr(rootInfo(1)))
        If IsEmpty(registrationNames) = False Then
            For Each registrationName In registrationNames
                PumpSetupResponsiveness
                If SetupControlCancellationRequested() Then RequireSafeWorkContinue "ScanDBInventoryAddInsAndLoadedVBAProjects", "Installed Access add-in registration boundary"
                registryPath = CStr(rootInfo(1)) & "\" & CStr(registrationName)
                registrationSubkey = ReferenceRegistryRootDisplay(CLng(rootInfo(0))) & "\" & registryPath
                addInLibrary = RegistryNamedString(registryProvider, CLng(rootInfo(0)), registryPath, "Library")
                expressionText = RegistryNamedString(registryProvider, CLng(rootInfo(0)), registryPath, "Expression")
                addInFilePath = ResolveTraditionalAccessAddInLibraryPath(addInLibrary)
                fileExistsValue = ReferenceFileExists(addInFilePath)
                diagnosticText = vbNullString
                readDiagnostic = vbNullString
                sourceRegistrationMatched = False
                usysRegInfoPresent = False
                If fileExistsValue And IsTraditionalAccessAddInDatabasePath(addInFilePath) Then
                    Set registrations = ReadTraditionalAccessAddInRegistrations(addInFilePath, usysRegInfoPresent, readDiagnostic)
                    For Each registration In registrations
                        sourceRegistrationName = CStr(registration(0))
                        namePosition = InStrRev(sourceRegistrationName, "\")
                        If namePosition > 0 And namePosition < Len(sourceRegistrationName) Then sourceRegistrationName = Mid$(sourceRegistrationName, namePosition + 1)
                        If StrComp(Trim$(sourceRegistrationName), Trim$(CStr(registrationName)), vbTextCompare) = 0 Then
                            sourceRegistrationMatched = True
                            If Len(expressionText) = 0 Then expressionText = CStr(registration(2))
                            Exit For
                        End If
                    Next registration
                    If Len(readDiagnostic) > 0 Then diagnosticText = readDiagnostic
                    If sourceRegistrationMatched = False And usysRegInfoPresent And Len(diagnosticText) = 0 Then diagnosticText = "Installed registration was found, but the add-in database USysRegInfo did not contain a matching Menu Add-ins registration group."
                End If
                ProcedureName = RegisteredAccessAddInProcedureName(expressionText)
                If Len(addInLibrary) = 0 Or Len(expressionText) = 0 Then
                    scanStatus = "Registration Incomplete"
                ElseIf fileExistsValue = False Then
                    scanStatus = "Registered File Missing"
                    If Len(diagnosticText) = 0 Then diagnosticText = "The installed Access Menu Add-Ins registration points to an add-in database that is not currently accessible at the registered Library path."
                ElseIf Len(ProcedureName) = 0 Then
                    scanStatus = "Expression Unparsed"
                Else
                    scanStatus = "Registered"
                End If
                addInOrder = addInOrder + 1
                addInIdentity = "ACCESS|" & ReferenceScanTextFingerprint(UCase$(Trim$(CStr(registrationName))) & "|" & UCase$(Trim$(addInLibrary)))
                BeginInventoryUpsertRow accessAddInRows, addInIdentity
                accessAddInRows.fields(FIELD_ADD_IN_ORDER).value = addInOrder
                accessAddInRows.fields(FIELD_ADD_IN_IDENTITY).value = addInIdentity
                accessAddInRows.fields(FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME).value = SharedValue_NormalizeAssignment(CStr(registrationName))
                accessAddInRows.fields(FIELD_REGISTRATION_SUBKEY).value = registrationSubkey
                accessAddInRows.fields(FIELD_ADD_IN_LIBRARY).value = SharedValue_NormalizeAssignment(addInLibrary)
                accessAddInRows.fields(FIELD_ENTRY_POINT_EXPRESSION).value = SharedValue_NormalizeAssignment(expressionText)
                accessAddInRows.fields(FIELD_ENTRY_POINT_PROCEDURE).value = SharedValue_NormalizeAssignment(ProcedureName)
                accessAddInRows.fields(FIELD_ENTRY_POINT_EXISTS).value = Null
                accessAddInRows.fields(FIELD_ENTRY_POINT_EXTERNALLY_CALLABLE).value = Null
                accessAddInRows.fields(FIELD_PROJECT_IDENTITY).value = Null
                accessAddInRows.fields(FIELD_PROJECT_NAME).value = Null
                accessAddInRows.fields(FIELD_PROJECT_ROLE).value = Null
                accessAddInRows.fields(FIELD_PROJECT_FILE_PATH).value = SharedValue_NormalizeAssignment(addInFilePath)
                accessAddInRows.fields(FIELD_FILE_EXISTS).value = fileExistsValue
                accessAddInRows.fields(FIELD_SCAN_STATUS).value = scanStatus
                accessAddInRows.fields(FIELD_DIAGNOSTIC_TEXT).value = SharedValue_NormalizeAssignment(diagnosticText)
                accessAddInRows.fields(FIELD_INVENTORY_STATUS).value = INVENTORY_STATUS_CURRENT
                accessAddInRows.fields(FIELD_LAST_SCAN_BATCH_ID).value = ReferenceScanBatchID(scannedOn, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AccessAddIns"))
                accessAddInRows.fields(FIELD_LAST_SCANNED_ON).value = scannedOn
                accessAddInRows.Update
            Next registrationName
        End If
    Next rootInfo
    accessAddInRows.Close
    CorrelateTraditionalAccessAddInsWithLoadedProjects db
    If correlateEntryPoints Then CorrelateTraditionalAccessAddInEntryPoints db
    ws.CommitTrans
    transactionActive = False
    ScanTraditionalAccessAddInsCore = InventoryTableRowCount(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AccessAddIns"))
    Exit Function
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not accessAddInRows Is Nothing Then accessAddInRows.Close
    If transactionActive And Not ws Is Nothing Then ws.Rollback
    On Error GoTo 0
    Err.Raise savedErrNumber, "ScanTraditionalAccessAddInsCore", "Could not scan installed traditional Access add-in registrations. The prior current DBInv_AddIns_Access partition was preserved." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Function

Private Sub CorrelateTraditionalAccessAddInsWithLoadedProjects(ByVal db As DAO.Database)
    'Loaded-project inventory is enrichment only: the registered database path is normalized to the existing indexed ProjectPathKey and never defines DBInv_AddIns_Access membership.
    On Error GoTo ErrHandler
    Dim accessRows As DAO.Recordset
    Dim projectRows As DAO.Recordset
    Dim accessTable As String
    Dim projectTable As String
    Dim projectPathKey As String
    Dim savedErrNumber As Long
    Dim savedErrDescription As String
    accessTable = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AccessAddIns")
    projectTable = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects")
    Set accessRows = db.OpenRecordset("SELECT " & SQL_FIELD_ID & "," & SQL_FIELD_PROJECT_FILE_PATH & "," & SQL_FIELD_PROJECT_IDENTITY & "," & SQL_FIELD_PROJECT_NAME & "," & SQL_FIELD_PROJECT_ROLE & " FROM [" & accessTable & "] WHERE " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & ";", dbOpenDynaset)
    Do While accessRows.EOF = False
        projectPathKey = LoadedVBAProjectPathKey(Nz(accessRows.fields(FIELD_PROJECT_FILE_PATH).value, vbNullString))
        accessRows.Edit
        accessRows.fields(FIELD_PROJECT_IDENTITY).value = Null
        accessRows.fields(FIELD_PROJECT_NAME).value = Null
        accessRows.fields(FIELD_PROJECT_ROLE).value = Null
        If Len(projectPathKey) > 0 Then
            Set projectRows = db.OpenRecordset("SELECT TOP 1 " & SQL_FIELD_PROJECT_IDENTITY & "," & SQL_FIELD_PROJECT_NAME & "," & SQL_FIELD_PROJECT_ROLE & " FROM [" & projectTable & "] WHERE " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " AND " & SQL_FIELD_PROJECT_PATH_KEY & "=" & SharedDao_SqlQuote(projectPathKey) & " ORDER BY " & SQL_FIELD_PROJECT_ORDER & ";", dbOpenSnapshot)
            If projectRows.EOF = False Then
                accessRows.fields(FIELD_PROJECT_IDENTITY).value = projectRows.fields(FIELD_PROJECT_IDENTITY).value
                accessRows.fields(FIELD_PROJECT_NAME).value = projectRows.fields(FIELD_PROJECT_NAME).value
                accessRows.fields(FIELD_PROJECT_ROLE).value = projectRows.fields(FIELD_PROJECT_ROLE).value
            End If
            projectRows.Close
            Set projectRows = Nothing
        End If
        accessRows.Update
        accessRows.MoveNext
    Loop
    accessRows.Close
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not projectRows Is Nothing Then projectRows.Close
    If Not accessRows Is Nothing Then accessRows.Close
    On Error GoTo 0
    Err.Raise savedErrNumber, "CorrelateTraditionalAccessAddInsWithLoadedProjects", savedErrDescription
End Sub

Private Sub CorrelateTraditionalAccessAddInEntryPoints(ByVal db As DAO.Database)
    'Single set-based correlation authority. It reuses only complete DBInv_VBAProcs_@LdPrjs inventories and never reopens or rescans source modules.
    Dim accessTable As String
    Dim projectTable As String
    Dim procedureTable As String
    Dim completeStatusList As String
    accessTable = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AccessAddIns")
    projectTable = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProjects")
    procedureTable = ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_LoadedVBAProcedures")
    completeStatusList = "('Source Scanned','Source Scanned; Refs Limited','No Components','No Components; Refs Limited')"
    db.Execute "UPDATE [" & accessTable & "] AS A INNER JOIN [" & projectTable & "] AS V ON A." & SQL_FIELD_PROJECT_IDENTITY & "=V." & SQL_FIELD_PROJECT_IDENTITY & " SET A." & SQL_FIELD_ENTRY_POINT_EXISTS & "=False,A." & SQL_FIELD_ENTRY_POINT_EXTERNALLY_CALLABLE & "=Null WHERE A." & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " AND V." & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " AND Len(Nz(A." & SQL_FIELD_ENTRY_POINT_PROCEDURE & ",''))>0 AND V." & SQL_FIELD_SOURCE_ANALYSIS_STATUS & " In " & completeStatusList & ";", dbFailOnError
    db.Execute "UPDATE [" & accessTable & "] AS A INNER JOIN [" & procedureTable & "] AS P ON A." & SQL_FIELD_PROJECT_IDENTITY & "=P." & SQL_FIELD_PROJECT_IDENTITY & " SET A." & SQL_FIELD_ENTRY_POINT_EXISTS & "=True,A." & SQL_FIELD_ENTRY_POINT_EXTERNALLY_CALLABLE & "=False WHERE A." & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " AND P." & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " AND UCase(A." & SQL_FIELD_ENTRY_POINT_PROCEDURE & ")=P." & SQL_FIELD_PROCEDURE_NAME_KEY & " AND A." & SQL_FIELD_PROJECT_IDENTITY & " In (SELECT " & SQL_FIELD_PROJECT_IDENTITY & " FROM [" & projectTable & "] WHERE " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " AND " & SQL_FIELD_SOURCE_ANALYSIS_STATUS & " In " & completeStatusList & ");", dbFailOnError
    db.Execute "UPDATE [" & accessTable & "] AS A INNER JOIN [" & procedureTable & "] AS P ON A." & SQL_FIELD_PROJECT_IDENTITY & "=P." & SQL_FIELD_PROJECT_IDENTITY & " SET A." & SQL_FIELD_ENTRY_POINT_EXTERNALLY_CALLABLE & "=True WHERE A." & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " AND P." & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " AND UCase(A." & SQL_FIELD_ENTRY_POINT_PROCEDURE & ")=P." & SQL_FIELD_PROCEDURE_NAME_KEY & " AND A." & SQL_FIELD_PROJECT_IDENTITY & " In (SELECT " & SQL_FIELD_PROJECT_IDENTITY & " FROM [" & projectTable & "] WHERE " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " AND " & SQL_FIELD_SOURCE_ANALYSIS_STATUS & " In " & completeStatusList & ") AND P." & SQL_FIELD_IS_EXTERNALLY_CALLABLE & "=True;", dbFailOnError
End Sub

Private Function ScanVBEAddInsCore() As Long
    On Error GoTo ErrHandler
    Dim addIns As Object
    Set addIns = CallByName(Application.vbe, "AddIns", VbGet)
    ScanVBEAddInsCore = ScanAddInCollection("VBE Add-in", ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_VBEAddIns"), addIns)
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "ScanVBEAddInsCore", "The VBE AddIns collection could not be inventoried without changing its state." & vbCrLf & Err.Description
End Function

Private Function ScanCOMAddInsCore() As Long
    On Error GoTo ErrHandler
    Dim addIns As Object
    Set addIns = CallByName(Application, "COMAddIns", VbGet)
    ScanCOMAddInsCore = ScanAddInCollection("COM Add-in", ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_COMAddIns"), addIns)
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "ScanCOMAddInsCore", "The Office COMAddIns collection could not be inventoried without changing its state." & vbCrLf & Err.Description
End Function

Private Function ScanAddInCollection(ByVal sourceKind As String, ByVal tableName As String, ByVal addIns As Object) As Long
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim ws As DAO.workspace
    Dim rs As DAO.Recordset
    Dim addInItem As Object
    Dim addInIndex As Long
    Dim addInCount As Long
    Dim scannedOn As Date
    Dim addInIdentity As String
    Dim displayName As String
    Dim descriptionText As String
    Dim progIDText As String
    Dim guidText As String
    Dim scanStatus As String
    Dim diagnosticText As String
    Dim propertyDiagnostic As String
    Dim isConnected As Boolean
    Dim connectReadable As Boolean
    Dim objectAvailable As Boolean
    Dim objectReadable As Boolean
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    If addIns Is Nothing Then Err.Raise vbObjectError + 9730, "ScanAddInCollection", sourceKind & " collection is unavailable."
    addInCount = CLng(addIns.count)
    CorePulseProgress "Add-in inventory | " & sourceKind & " | 0/" & CStr(addInCount) & " | passive state read"
    Set db = ADCDb()
    Set ws = DBEngine.Workspaces(0)
    scannedOn = Now
    ws.BeginTrans
    PrepareInventoryRowsForScan db, tableName
    Set rs = OpenInventoryUpsertRecordset(db, tableName, IIf(StrComp(tableName, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_VBEAddIns"), vbTextCompare) = 0, "UX_VBE_Identity", "UX_COM_Identity"))
    For addInIndex = 1 To addInCount
        PumpSetupResponsiveness
        Set addInItem = addIns.item(addInIndex)
        descriptionText = SafeAddInText(addInItem, "Description")
        progIDText = Left$(SafeAddInText(addInItem, "ProgId"), 255)
        guidText = Left$(NormalizeReferenceGuid(SafeAddInText(addInItem, "Guid")), 100)
        displayName = Left$(AddInDisplayName(sourceKind, addInIndex, descriptionText, progIDText, guidText), 255)
        CorePulseProgress "Add-in inventory | " & sourceKind & " | " & CStr(addInIndex) & "/" & CStr(addInCount) & " | inspecting " & IIf(Len(displayName) > 0, displayName, IIf(Len(progIDText) > 0, progIDText, "(unnamed add-in)")) & " | ProgID=" & IIf(Len(progIDText) > 0, progIDText, "(none)")
        diagnosticText = vbNullString
        connectReadable = False
        objectReadable = False
        propertyDiagnostic = vbNullString
        isConnected = SafeAddInBoolean(addInItem, "Connect", False, connectReadable, propertyDiagnostic)
        If connectReadable = False Then diagnosticText = SharedText_AppendNonBlankLine(diagnosticText, propertyDiagnostic)
        If connectReadable And isConnected Then
            propertyDiagnostic = vbNullString
            objectAvailable = SafeAddInObjectAvailable(addInItem, objectReadable, propertyDiagnostic)
            If objectReadable = False Then diagnosticText = SharedText_AppendNonBlankLine(diagnosticText, propertyDiagnostic)
        Else
            objectAvailable = False
        End If
        scanStatus = AddInScanStatus(isConnected, connectReadable)
        PulseProgressEvery addInIndex, 10, "Add-in inventory | " & sourceKind & " | " & CStr(addInIndex) & "/" & CStr(addInCount) & " | " & IIf(Len(displayName) > 0, displayName, "(unnamed add-in)") & " | status=" & scanStatus & " | connected=" & CStr(isConnected) & " | object available=" & CStr(objectAvailable), addInCount, "ReferenceAddInInventoryScan", "References"
        addInIdentity = BuildAddInIdentity(sourceKind, addInIndex, guidText, progIDText, displayName)
        BeginInventoryUpsertRow rs, addInIdentity
        rs.fields(FIELD_ADD_IN_IDENTITY).value = addInIdentity
        rs.fields(FIELD_ADD_IN_ORDER).value = addInIndex
        rs.fields(FIELD_MANAGED_OBJECT_FIELD_DISPLAY_NAME).value = SharedValue_NormalizeAssignment(displayName)
        rs.fields(FIELD_DESCRIPTION).value = SharedValue_NormalizeAssignment(descriptionText)
        rs.fields(FIELD_PROG_ID).value = SharedValue_NormalizeAssignment(progIDText)
        rs.fields(FIELD_GUID).value = SharedValue_NormalizeAssignment(guidText)
        rs.fields(FIELD_IS_REGISTERED).value = True
        rs.fields(FIELD_IS_CONNECTED).value = isConnected
        rs.fields(FIELD_OBJECT_AVAILABLE).value = objectAvailable
        rs.fields(FIELD_SCAN_STATUS).value = scanStatus
        rs.fields(FIELD_DIAGNOSTIC_TEXT).value = SharedValue_NormalizeAssignment(diagnosticText)
        rs.fields(FIELD_INVENTORY_STATUS).value = INVENTORY_STATUS_CURRENT
        rs.fields(FIELD_LAST_SCAN_BATCH_ID).value = ReferenceScanBatchID(scannedOn, tableName)
        rs.fields(FIELD_LAST_SCANNED_ON).value = scannedOn
        rs.Update
    Next addInIndex
    rs.Close
    ws.CommitTrans
    ScanAddInCollection = addInCount
    LogTraceMessage "Info", "AddIns", "ScanAddInCollection", sourceKind & " inventory completed.", "Rows stored: " & CStr(addInCount) & vbCrLf & "Table: " & tableName & vbCrLf & "The collection was read passively; no Connect property was assigned.", "Scan " & LCase$(sourceKind) & " inventory", 0, vbNullString, vbNullString, tableName, "Table", "CodeDb", "AddIns"
    Exit Function
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not ws Is Nothing Then ws.Rollback
    On Error GoTo 0
    Err.Raise savedErrNumber, "ScanAddInCollection", "Could not refresh " & tableName & ". Its prior current retained-inventory partition was preserved." & vbCrLf & "Original source: " & savedErrSource & vbCrLf & "Description: " & savedErrDescription
End Function

Private Sub AddCollectionKey(ByVal values As Collection, ByVal keyText As String)
    On Error Resume Next
    values.Add True, keyText
End Sub

Private Function CollectionHasKey(ByVal values As Collection, ByVal keyText As String) As Boolean
    On Error GoTo Missing
    Dim valueItem As Variant
    valueItem = values(keyText)
    CollectionHasKey = True
    Exit Function
Missing:
    CollectionHasKey = False
End Function

' =============================================================================================================================================================================
' REFERENCE MANAGER RUNTIME / APPLY-DISCARD PROCEDURES
' =============================================================================================================================================================================

' Purpose: Initializes the generated Reference Manager form from current reference-inventory state when its form-open expression invokes this entry point.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters: None.
' Returns: Boolean: True when refresh reference manager completes its documented action normally; False when the routine follows its handled failure, cancellation, rejection, or not-applicable path.
Public Function ReferenceManagerFormOpenInitialize() As Boolean
    On Error GoTo ErrHandler
    Dim pendingCount As Long
    Dim response As VbMsgBoxResult
    Dim refreshManagerData As Boolean
    Dim foundationDetail As String
    If ReferenceInventoryFoundationCurrentForSetup(foundationDetail) = False Then
        Err.Raise vbObjectError + 9799, "ReferenceManagerFormOpenInitialize", "The A4 reference/Add-In inventory foundation is not current. Reference Manager initialization does not repair setup-owned physical objects." & vbCrLf & foundationDetail
    End If
    pendingCount = ReferenceManagerPendingChangeCount()
    If pendingCount > 0 Then
        response = CoreShowMessageResult(CStr(pendingCount) & " staged reference change(s) remain from the prior Reference Manager session." & vbCrLf & vbCrLf & "Yes: reopen the staged changes." & vbCrLf & "No: discard them and rescan both inventories." & vbCrLf & "Cancel: leave the manager closed.", vbQuestion Or vbYesNoCancel Or vbDefaultButton1, "Reference Manager", "ReferenceManagerFormOpenInitialize", "References")
        If response = vbCancel Then
            DoCmd.CancelEvent
            Exit Function
        End If
        refreshManagerData = (response = vbNo)
    Else
        refreshManagerData = True
    End If
    If refreshManagerData Then
        ScanDBInventoryActiveReferences False, False, False
        ScanDBInventoryLoadedVBAProjects False, False
        ScanDBInventoryAvailableReferences False, False
        PopulateReferenceManager
    End If
    ReferenceManagerFormOpenInitialize = True
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
    ReportReferenceFailure "ReferenceManagerFormOpenInitialize", "Could not initialize the Reference Manager while opening it.", True, savedErrNumber, savedErrSource, savedErrDescription
End Function

' Purpose: Refreshes reference manager for prepare reference inventory.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
'   prepareReferenceInventory (Input Optional; Boolean): When True, enables or requests reference inventory for this call; when False, the routine follows its normal non-forced/default path. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub RefreshReferenceManager(Optional ByVal showMessage As Boolean = True, Optional ByVal prepareReferenceInventory As Boolean = True)
    On Error GoTo ErrHandler
    Dim foundationDetail As String
    If prepareReferenceInventory Then
        If ReferenceInventoryFoundationCurrentForSetup(foundationDetail) = False Then
            Err.Raise vbObjectError + 9799, "RefreshReferenceManager", "The A4 reference/Add-In inventory foundation is not current. Reference Manager refresh does not repair setup-owned physical objects." & vbCrLf & foundationDetail
        End If
    End If
    ScanDBInventoryActiveReferences False, False, False
    ScanDBInventoryLoadedVBAProjects False, False
    ScanDBInventoryAvailableReferences False, False
    PopulateReferenceManager
    RequeryReferenceManagerForm
    If showMessage Then ReportCoreActionSuccess "RefreshReferenceManager", "References", "Reference Manager refreshed successfully.", "Active, loaded-project, and available reference inventories were rescanned and the staged manager list was rebuilt.", "Refresh Reference Manager", ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ReferenceManager"), "Table", "CodeDb", "References", True, "Reference Manager", vbInformation, True, "Tables", False
    Exit Sub
ErrHandler:
    ReportReferenceFailure "RefreshReferenceManager", "Could not refresh the Reference Manager.", showMessage, Err.Number, Err.Source, Err.Description
End Sub

Private Sub PopulateReferenceManager()
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim ws As DAO.workspace
    Dim sourceRs As DAO.Recordset
    Dim targetRs As DAO.Recordset
    Dim managerKeys As Collection
    Dim refreshedOn As Date
    Dim protectionReason As String
    Dim managerKey As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Set db = ADCDb()
    Set ws = DBEngine.Workspaces(0)
    ' INTENTIONAL KEYED LOOKUP: managerKeys provides keyed membership/lookup over source rows or state already read by PopulateReferenceManager; it is a derived access path, not a second authority.
    Set managerKeys = New Collection
    refreshedOn = Now
    ws.BeginTrans
    db.Execute "DELETE FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ReferenceManager") & "];", dbFailOnError
    Set targetRs = db.OpenRecordset(ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ReferenceManager"), dbOpenDynaset)
    Set sourceRs = db.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ActiveReferences") & "] WHERE " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " ORDER BY " & SQL_FIELD_REFERENCE_PRIORITY & "," & SQL_FIELD_ID & ";", dbOpenSnapshot)
    Do While sourceRs.EOF = False
        PumpSetupResponsiveness
        managerKey = ManagerReferenceMatchKey(CStr(Nz(sourceRs.fields(FIELD_REFERENCE_GUID).value, vbNullString)), CLng(Nz(sourceRs.fields(FIELD_MAJOR_VERSION).value, 0)), CLng(Nz(sourceRs.fields(FIELD_MINOR_VERSION).value, 0)), CStr(Nz(sourceRs.fields(FIELD_FULL_PATH).value, vbNullString)), CStr(Nz(sourceRs.fields(FIELD_REFERENCE_NAME).value, vbNullString)))
        AddCollectionKey managerKeys, managerKey
        protectionReason = ReferenceChangeProtectionReason(CStr(Nz(sourceRs.fields(FIELD_REFERENCE_NAME).value, vbNullString)), CBool(Nz(sourceRs.fields(FIELD_IS_BUILT_IN).value, False)))
        targetRs.AddNew
        targetRs.fields(FIELD_REFERENCE_IDENTITY).value = ManagerReferenceIdentity("Active", CStr(Nz(sourceRs.fields(FIELD_REFERENCE_GUID).value, vbNullString)), CLng(Nz(sourceRs.fields(FIELD_MAJOR_VERSION).value, 0)), CLng(Nz(sourceRs.fields(FIELD_MINOR_VERSION).value, 0)), CStr(Nz(sourceRs.fields(FIELD_FULL_PATH).value, vbNullString)), CStr(Nz(sourceRs.fields(FIELD_REFERENCE_NAME).value, vbNullString)), CLng(Nz(sourceRs.fields(FIELD_ID).value, 0)))
        targetRs.fields(FIELD_REQUESTED_ACTIVE).value = True
        targetRs.fields(FIELD_ORIGINAL_ACTIVE).value = True
        targetRs.fields(FIELD_REFERENCE_SOURCE).value = "Active"
        CopyReferenceManagerCommonFields sourceRs, targetRs, True
        targetRs.fields(FIELD_CAN_CHANGE).value = (Len(protectionReason) = 0)
        targetRs.fields(FIELD_PROTECTION_REASON).value = SharedValue_NormalizeAssignment(protectionReason)
        targetRs.fields(FIELD_HEALTH_STATUS).value = sourceRs.fields(FIELD_SCAN_STATUS).value
        targetRs.fields(FIELD_LAST_REFRESHED_ON).value = refreshedOn
        targetRs.Update
        sourceRs.MoveNext
    Loop
    sourceRs.Close
    Set sourceRs = db.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_AvailableReferences") & "] WHERE " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & " ORDER BY " & SQL_FIELD_REFERENCE_NAME & "," & SQL_FIELD_MAJOR_VERSION & " DESC," & SQL_FIELD_MINOR_VERSION & " DESC," & SQL_FIELD_FILE_EXISTS & " DESC," & SQL_FIELD_ID & ";", dbOpenSnapshot)
    Do While sourceRs.EOF = False
        PumpSetupResponsiveness
        managerKey = ManagerReferenceMatchKey(CStr(Nz(sourceRs.fields(FIELD_REFERENCE_GUID).value, vbNullString)), CLng(Nz(sourceRs.fields(FIELD_MAJOR_VERSION).value, 0)), CLng(Nz(sourceRs.fields(FIELD_MINOR_VERSION).value, 0)), CStr(Nz(sourceRs.fields(FIELD_REGISTERED_PATH).value, vbNullString)), CStr(Nz(sourceRs.fields(FIELD_REFERENCE_NAME).value, vbNullString)))
        If CollectionHasKey(managerKeys, managerKey) = False Then
            AddCollectionKey managerKeys, managerKey
            targetRs.AddNew
            targetRs.fields(FIELD_REFERENCE_IDENTITY).value = ManagerReferenceIdentity("Available", CStr(Nz(sourceRs.fields(FIELD_REFERENCE_GUID).value, vbNullString)), CLng(Nz(sourceRs.fields(FIELD_MAJOR_VERSION).value, 0)), CLng(Nz(sourceRs.fields(FIELD_MINOR_VERSION).value, 0)), CStr(Nz(sourceRs.fields(FIELD_REGISTERED_PATH).value, vbNullString)), CStr(Nz(sourceRs.fields(FIELD_REFERENCE_NAME).value, vbNullString)), CLng(Nz(sourceRs.fields(FIELD_ID).value, 0)))
            targetRs.fields(FIELD_REQUESTED_ACTIVE).value = False
            targetRs.fields(FIELD_ORIGINAL_ACTIVE).value = False
            targetRs.fields(FIELD_REFERENCE_SOURCE).value = "Available"
            CopyReferenceManagerCommonFields sourceRs, targetRs, False
            targetRs.fields(FIELD_CAN_CHANGE).value = (CBool(Nz(sourceRs.fields(FIELD_FILE_EXISTS).value, False)) Or Len(Trim$(CStr(Nz(sourceRs.fields(FIELD_REFERENCE_GUID).value, vbNullString)))) > 0)
            If CBool(targetRs.fields(FIELD_CAN_CHANGE).value) = False Then targetRs.fields(FIELD_PROTECTION_REASON).value = "This registration has no usable file path or GUID for activation."
            targetRs.fields(FIELD_HEALTH_STATUS).value = sourceRs.fields(FIELD_SCAN_STATUS).value
            targetRs.fields(FIELD_LAST_REFRESHED_ON).value = refreshedOn
            targetRs.Update
        End If
        sourceRs.MoveNext
    Loop
    sourceRs.Close
    targetRs.Close
    ws.CommitTrans
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not sourceRs Is Nothing Then sourceRs.Close
    If Not targetRs Is Nothing Then targetRs.Close
    If Not ws Is Nothing Then ws.Rollback
    On Error GoTo 0
    Err.Raise savedErrNumber, "PopulateReferenceManager", "Could not rebuild the Reference Manager staging rows." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub CopyReferenceManagerCommonFields(ByVal sourceRs As DAO.Recordset, ByVal targetRs As DAO.Recordset, ByVal activeSource As Boolean)
    targetRs.fields(FIELD_REFERENCE_NAME).value = sourceRs.fields(FIELD_REFERENCE_NAME).value
    targetRs.fields(FIELD_REFERENCE_DESCRIPTION).value = sourceRs.fields(FIELD_REFERENCE_DESCRIPTION).value
    targetRs.fields(FIELD_REFERENCE_GUID).value = sourceRs.fields(FIELD_REFERENCE_GUID).value
    targetRs.fields(FIELD_MAJOR_VERSION).value = sourceRs.fields(FIELD_MAJOR_VERSION).value
    targetRs.fields(FIELD_MINOR_VERSION).value = sourceRs.fields(FIELD_MINOR_VERSION).value
    targetRs.fields(FIELD_REGISTERED_PATH).value = sourceRs.fields(FIELD_REGISTERED_PATH).value
    targetRs.fields(FIELD_FILE_EXISTS).value = sourceRs.fields(FIELD_FILE_EXISTS).value
    targetRs.fields(FIELD_REGISTRATION_SCOPE).value = sourceRs.fields(FIELD_REGISTRATION_SCOPE).value
    targetRs.fields(FIELD_REGISTRATION_VIEW).value = sourceRs.fields(FIELD_REGISTRATION_VIEW).value
    targetRs.fields(FIELD_PLATFORM).value = sourceRs.fields(FIELD_PLATFORM).value
    If activeSource Then
        targetRs.fields(FIELD_REFERENCE_PRIORITY).value = sourceRs.fields(FIELD_REFERENCE_PRIORITY).value
        targetRs.fields(FIELD_FULL_PATH).value = sourceRs.fields(FIELD_FULL_PATH).value
        targetRs.fields(FIELD_IS_BUILT_IN).value = sourceRs.fields(FIELD_IS_BUILT_IN).value
        targetRs.fields(FIELD_IS_BROKEN).value = sourceRs.fields(FIELD_IS_BROKEN).value
        targetRs.fields(FIELD_IS_REGISTERED).value = sourceRs.fields(FIELD_IS_REGISTERED).value
    Else
        targetRs.fields(FIELD_REFERENCE_PRIORITY).value = Null
        targetRs.fields(FIELD_FULL_PATH).value = Null
        targetRs.fields(FIELD_IS_BUILT_IN).value = False
        targetRs.fields(FIELD_IS_BROKEN).value = False
        targetRs.fields(FIELD_IS_REGISTERED).value = True
    End If
End Sub

Private Function ReferenceManagerPendingChangeCount() As Long
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = ADCDb()
    Set rs = db.OpenRecordset("SELECT Count(*) AS " & SQL_FIELD_PENDING_COUNT & " FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ReferenceManager") & "] WHERE " & SQL_FIELD_REQUESTED_ACTIVE & "<>" & SQL_FIELD_ORIGINAL_ACTIVE & ";", dbOpenSnapshot)
    ReferenceManagerPendingChangeCount = CLng(rs.fields(FIELD_PENDING_COUNT).value)
    rs.Close
End Function

Private Sub RequeryReferenceManagerForm()
    If FormIsLoaded(TargetManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Reference_Manager")) Then forms(TargetManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Reference_Manager")).Requery
End Sub

' Purpose: Applies pending Reference Manager add/remove/activation changes transactionally, verifies the resulting VBA reference state, and performs rollback/recovery plus optional user-facing reporting on
'   failure.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine displays its normal user-facing status, completion, warning, or error message; False suppresses that presentation while preserving the
'   underlying operation. Default: True.
' Returns: None (Sub). The procedure communicates its result through its database, UI, trace, or module-state side effects and raises errors where its implementation requires it.
Public Sub ApplyReferenceManagerChanges(Optional ByVal showMessage As Boolean = True)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim addedCount As Long
    Dim removedCount As Long
    Dim pendingCount As Long
    Dim activationCount As Long
    Dim deactivationCount As Long
    Dim confirmationText As String
    Dim changesVerified As Boolean
    Dim referenceChangesStarted As Boolean
    Dim rollbackDetail As String
    Dim foundationDetail As String
    Dim failureHeadline As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    SaveReferenceManagerCurrentRecord
    Set db = ADCDb()
    ValidateReferenceManagerPlan activationCount, deactivationCount, confirmationText
    pendingCount = activationCount + deactivationCount
    If pendingCount = 0 Then
        If showMessage Then CoreShowMessage "There are no staged reference changes to apply.", vbInformation, "Reference Manager", "ShowADCMessage", "Tables"
        Exit Sub
    End If
    If ReferenceInventoryFoundationCurrentForSetup(foundationDetail) = False Then
        Err.Raise vbObjectError + 9799, "ApplyReferenceManagerChanges", "The A4 reference/Add-In inventory foundation is not current. Reference Manager apply does not repair setup-owned physical objects." & vbCrLf & foundationDetail
    End If
    If UtilityMsgBox(confirmationText, vbQuestion Or vbYesNo Or vbDefaultButton2, "Apply Reference Changes") <> vbYes Then Exit Sub
    Set rs = db.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ReferenceManager") & "] WHERE " & SQL_FIELD_REQUESTED_ACTIVE & "=True AND " & SQL_FIELD_ORIGINAL_ACTIVE & "=False ORDER BY " & SQL_FIELD_REFERENCE_NAME & "," & SQL_FIELD_MAJOR_VERSION & " DESC," & SQL_FIELD_MINOR_VERSION & " DESC," & SQL_FIELD_ID & ";", dbOpenSnapshot)
    Do While rs.EOF = False
        PumpSetupResponsiveness
        referenceChangesStarted = True
        AddReferenceFromManagerRecord rs
        addedCount = addedCount + 1
        rs.MoveNext
    Loop
    rs.Close
    Set rs = db.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ReferenceManager") & "] WHERE " & SQL_FIELD_REQUESTED_ACTIVE & "=False AND " & SQL_FIELD_ORIGINAL_ACTIVE & "=True ORDER BY " & SQL_FIELD_REFERENCE_PRIORITY & " DESC," & SQL_FIELD_ID & " DESC;", dbOpenSnapshot)
    Do While rs.EOF = False
        PumpSetupResponsiveness
        referenceChangesStarted = True
        RemoveReferenceFromManagerRecord rs
        removedCount = removedCount + 1
        rs.MoveNext
    Loop
    rs.Close
    VerifyReferenceManagerRequestedState
    changesVerified = True
    ScanDBInventoryActiveReferences False, False, False
    ScanDBInventoryLoadedVBAProjects False, False
    ScanDBInventoryAvailableReferences False, False
    PopulateReferenceManager
    RequeryReferenceManagerForm
    LogTraceMessage "Info", "References", "ApplyReferenceManagerChanges", "Reference changes applied and verified.", "Activated: " & CStr(addedCount) & vbCrLf & "Deactivated: " & CStr(removedCount), "Apply reference changes", 0, vbNullString, vbNullString, ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ReferenceManager"), "Table", "CodeDb", "References"
    If showMessage Then ReportCoreActionSuccess "ApplyReferenceManagerChanges", "References", "Reference changes applied successfully.", "Activated: " & CStr(addedCount) & vbCrLf & "Deactivated: " & CStr(removedCount) & vbCrLf & "The active, loaded-project, and available inventories were rescanned and verified.", "Apply reference changes", ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ReferenceManager"), "Table", "CodeDb", "References", True, "Reference Manager", vbInformation, True, "Tables", False
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If referenceChangesStarted And changesVerified = False Then RestoreOriginalReferenceState rollbackDetail
    On Error GoTo 0
    If changesVerified Then
        failureHeadline = "The reference changes were applied and verified, but the post-change inventory refresh failed."
    Else
        failureHeadline = "Could not apply the staged reference changes."
    End If
    If Len(rollbackDetail) > 0 Then savedErrDescription = savedErrDescription & vbCrLf & vbCrLf & "Rollback:" & vbCrLf & rollbackDetail
    ReportReferenceFailure "ApplyReferenceManagerChanges", failureHeadline, showMessage, savedErrNumber, savedErrSource, savedErrDescription
End Sub

' Purpose: Performs the public discard reference manager changes operation using the routine's existing ADC workflow and side effects.
' Public role: Public dynamic/form expression API. The name is referenced as text outside this module, so Public visibility is required for Access or generated-expression resolution.
' Parameters:
'   showMessage (Input Optional; Boolean): Controls whether the routine shows its user-facing completion, status, warning, or error message in addition to performing the underlying work. Default: True.
' Returns: None (Sub). Any caller-visible results are communicated through existing side effects.
Public Sub DiscardReferenceManagerChanges(Optional ByVal showMessage As Boolean = True)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    SaveReferenceManagerCurrentRecord
    Set db = ADCDb()
    db.Execute "UPDATE [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ReferenceManager") & "] SET " & SQL_FIELD_REQUESTED_ACTIVE & "=" & SQL_FIELD_ORIGINAL_ACTIVE & ";", dbFailOnError
    RequeryReferenceManagerForm
    If showMessage Then CoreShowMessage "All staged reference changes were discarded.", vbInformation, "Reference Manager", "ShowADCMessage", "Tables"
    Exit Sub
ErrHandler:
    ReportReferenceFailure "DiscardReferenceManagerChanges", "Could not discard the staged reference changes.", showMessage, Err.Number, Err.Source, Err.Description
End Sub

Private Sub ValidateReferenceManagerPlan(ByRef activationCount As Long, ByRef deactivationCount As Long, ByRef confirmationText As String)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim requestedGuids As Collection
    Dim guidKey As String
    Dim displayText As String
    Dim detailText As String
    Dim existingReference As Object
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Set db = ADCDb()
    ' INTENTIONAL FILTERED PARTITION: requestedGuids is a consumer-required subset/projection derived from the source enumerated by ValidateReferenceManagerPlan; its membership is not independently authoritative.
    Set requestedGuids = New Collection
    Set rs = db.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ReferenceManager") & "] WHERE " & SQL_FIELD_REQUESTED_ACTIVE & "<>" & SQL_FIELD_ORIGINAL_ACTIVE & " ORDER BY " & SQL_FIELD_REQUESTED_ACTIVE & " DESC," & SQL_FIELD_REFERENCE_NAME & "," & SQL_FIELD_MAJOR_VERSION & " DESC," & SQL_FIELD_MINOR_VERSION & " DESC," & SQL_FIELD_ID & ";", dbOpenSnapshot)
    Do While rs.EOF = False
        PumpSetupResponsiveness
        displayText = ReferenceManagerRecordDisplay(rs)
        If CBool(Nz(rs.fields(FIELD_CAN_CHANGE).value, False)) = False Then Err.Raise vbObjectError + 9710, "ValidateReferenceManagerPlan", "A protected reference has a staged change." & vbCrLf & displayText & vbCrLf & CStr(Nz(rs.fields(FIELD_PROTECTION_REASON).value, vbNullString))
        If CBool(rs.fields(FIELD_REQUESTED_ACTIVE).value) Then
            activationCount = activationCount + 1
            guidKey = UCase$(Trim$(CStr(Nz(rs.fields(FIELD_REFERENCE_GUID).value, vbNullString))))
            If Len(guidKey) > 0 Then
                If CollectionHasKey(requestedGuids, guidKey) Then Err.Raise vbObjectError + 9711, "ValidateReferenceManagerPlan", "More than one version of the same reference GUID is staged for activation." & vbCrLf & displayText
                AddCollectionKey requestedGuids, guidKey
                Set existingReference = FindActiveReferenceByGuid(guidKey)
                If Not existingReference Is Nothing Then Err.Raise vbObjectError + 9712, "ValidateReferenceManagerPlan", "A different version of this GUID is already active. Version replacement is blocked because programmatic remove/re-add can change reference priority. Deactivate the current version, refresh, and then activate the replacement." & vbCrLf & displayText
            End If
            If Len(guidKey) = 0 And ReferenceFileExists(ReferenceManagerActivationPath(rs)) = False Then Err.Raise vbObjectError + 9713, "ValidateReferenceManagerPlan", "The reference cannot be activated because it has neither a usable GUID nor an existing file path." & vbCrLf & displayText
            detailText = SharedText_AppendNonBlankLine(detailText, "Activate: " & displayText)
        Else
            deactivationCount = deactivationCount + 1
            detailText = SharedText_AppendNonBlankLine(detailText, "Deactivate: " & displayText)
        End If
        rs.MoveNext
    Loop
    rs.Close
    confirmationText = "Apply the staged reference changes?" & vbCrLf & vbCrLf & "Activate: " & CStr(activationCount) & vbCrLf & "Deactivate: " & CStr(deactivationCount)
    If Len(detailText) > 0 Then confirmationText = confirmationText & vbCrLf & vbCrLf & Left$(detailText, 3000)
    confirmationText = confirmationText & vbCrLf & vbCrLf & "New references are appended by the Access reference API. Existing retained reference order is not intentionally changed."
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
    Err.Raise savedErrNumber, "ValidateReferenceManagerPlan", "The staged reference plan is not safe to apply." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Sub VerifyReferenceManagerRequestedState()
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim isActive As Boolean
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Set db = ADCDb()
    Set rs = db.OpenRecordset("SELECT * FROM [" & ManagedObjectPhysicalName(NAMING_CATEGORY_TABLE, "ADC_ReferenceManager") & "] ORDER BY " & SQL_FIELD_ID & ";", dbOpenSnapshot)
    Do While rs.EOF = False
        PumpSetupResponsiveness
        isActive = ReferenceManagerRecordIsCurrentlyActive(rs)
        If isActive <> CBool(rs.fields(FIELD_REQUESTED_ACTIVE).value) Then Err.Raise vbObjectError + 9717, "VerifyReferenceManagerRequestedState", "The final active state does not match the requested state." & vbCrLf & ReferenceManagerRecordDisplay(rs) & vbCrLf & "Requested active: " & CStr(CBool(rs.fields(FIELD_REQUESTED_ACTIVE).value)) & vbCrLf & "Actual active: " & CStr(isActive)
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
ErrHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
    Err.Raise savedErrNumber, "VerifyReferenceManagerRequestedState", "Reference-state verification failed." & vbCrLf & "Err.Source: " & savedErrSource & vbCrLf & "Err.Description: " & savedErrDescription
End Sub

Private Function ReferenceManagerRecordIsCurrentlyActive(ByVal rs As DAO.Recordset) As Boolean
    Dim referenceItem As Object
    Set referenceItem = FindActiveReferenceForManagerRecord(rs)
    ReferenceManagerRecordIsCurrentlyActive = Not referenceItem Is Nothing
End Function

Private Function ReferenceManagerActivationPath(ByVal rs As DAO.Recordset) As String
    ReferenceManagerActivationPath = CStr(Nz(rs.fields(FIELD_REGISTERED_PATH).value, vbNullString))
    If ReferenceFileExists(ReferenceManagerActivationPath) = False Then ReferenceManagerActivationPath = CStr(Nz(rs.fields(FIELD_FULL_PATH).value, vbNullString))
End Function

Private Function ReferenceManagerRecordDisplay(ByVal rs As DAO.Recordset) As String
    ReferenceManagerRecordDisplay = CStr(Nz(rs.fields(FIELD_REFERENCE_NAME).value, "(unnamed reference)")) & " " & CStr(Nz(rs.fields(FIELD_MAJOR_VERSION).value, 0)) & "." & CStr(Nz(rs.fields(FIELD_MINOR_VERSION).value, 0))
End Function

Private Function ReferenceManagerRecordKey(ByVal rs As DAO.Recordset) As String
    Dim guidText As String
    Dim pathText As String
    guidText = NormalizeReferenceGuid(CStr(Nz(rs.fields(FIELD_REFERENCE_GUID).value, vbNullString)))
    If Len(guidText) > 0 Then
        ReferenceManagerRecordKey = ReferenceVersionIdentity(guidText, CLng(Nz(rs.fields(FIELD_MAJOR_VERSION).value, 0)), CLng(Nz(rs.fields(FIELD_MINOR_VERSION).value, 0)))
        Exit Function
    End If
    pathText = CStr(Nz(rs.fields(FIELD_FULL_PATH).value, vbNullString))
    If Len(NormalizeReferencePath(pathText)) = 0 Then pathText = CStr(Nz(rs.fields(FIELD_REGISTERED_PATH).value, vbNullString))
    If Len(NormalizeReferencePath(pathText)) > 0 Then
        ReferenceManagerRecordKey = ReferencePathIdentity(pathText)
    Else
        ReferenceManagerRecordKey = "N|" & UCase$(Trim$(CStr(Nz(rs.fields(FIELD_REFERENCE_NAME).value, vbNullString))))
    End If
End Function

Private Sub SaveReferenceManagerCurrentRecord()
    If FormIsLoaded(TargetManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Reference_Manager")) Then
        If forms(TargetManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Reference_Manager")).Dirty Then forms(TargetManagedObjectPhysicalName(NAMING_CATEGORY_GENERATED_FORM, "Reference_Manager")).Dirty = False
    End If
End Sub

Private Function FormIsLoaded(ByVal formName As String) As Boolean
    On Error GoTo Missing
    FormIsLoaded = CurrentProject.AllForms(formName).IsLoaded
    Exit Function
Missing:
    FormIsLoaded = False
End Function


Private Function RegistrySubkeys(ByVal registryProvider As Object, ByVal hiveValue As Long, ByVal keyPath As String) As Variant
    On Error GoTo NoKeys
    Dim values As Variant
    Dim resultCode As Long
    resultCode = registryProvider.EnumKey(hiveValue, keyPath, values)
    If resultCode <> 0 Then GoTo NoKeys
    If IsNull(values) Or IsEmpty(values) Then GoTo NoKeys
    RegistrySubkeys = values
    Exit Function
NoKeys:
    RegistrySubkeys = Empty
End Function

Private Function SafeAddInBoolean(ByVal addInItem As Object, ByVal propertyName As String, ByVal DefaultValue As Boolean, ByRef readable As Boolean, ByRef diagnosticText As String) As Boolean
    On Error GoTo Missing
    SafeAddInBoolean = CBool(CallByName(addInItem, propertyName, VbGet))
    readable = True
    Exit Function
Missing:
    diagnosticText = propertyName & " property could not be read. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    Err.Clear
    readable = False
    SafeAddInBoolean = DefaultValue
End Function

Private Function SafeAddInObjectAvailable(ByVal addInItem As Object, ByRef readable As Boolean, ByRef diagnosticText As String) As Boolean
    On Error GoTo Missing
    Dim exposedObject As Object
    Set exposedObject = CallByName(addInItem, "Object", VbGet)
    readable = True
    SafeAddInObjectAvailable = Not exposedObject Is Nothing
    Exit Function
Missing:
    diagnosticText = "Object property could not be read. Error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    Err.Clear
    readable = False
End Function

Private Function InventoryTableRowCount(ByVal tableName As String) As Long
    On Error GoTo Missing
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = ADCDb()
    Set rs = db.OpenRecordset("SELECT Count(*) AS " & SQL_FIELD_ROW_COUNT & " FROM [" & tableName & "] WHERE " & SQL_FIELD_INVENTORY_STATUS & "=" & SharedDao_SqlQuote(INVENTORY_STATUS_CURRENT) & ";", dbOpenSnapshot)
    InventoryTableRowCount = CLng(rs.fields(FIELD_ROW_COUNT).value)
    rs.Close
    Exit Function
Missing:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Err.Clear
End Function
