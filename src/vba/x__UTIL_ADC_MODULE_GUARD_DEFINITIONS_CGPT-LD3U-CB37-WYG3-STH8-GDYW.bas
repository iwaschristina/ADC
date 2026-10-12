Attribute VB_Name = "x__UTIL_ADC_MODULE_GUARD_DEFINITIONS"
' ================================================================================================================================================================================
' ChatGPT Rewrite Tracking Header
' ================================================================================================================================================================================
' ChatGPT Previous Version ID: CGPT-HUVA-BQ3R-NYBC-LFPC-MV99
' ChatGPT Baseline Version ID: CGPT-HUVA-BQ3R-NYBC-LFPC-MV99
' ChatGPT Version ID: CGPT-LD3U-CB37-WYG3-STH8-GDYW
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
' STANDARD MODULE: Access Development Console Module Guard Definitions
' VB_Name: x__UTIL_ADC_MODULE_GUARD_DEFINITIONS
' Entry Points: None
' SHARED CODE: None
' CLASSES USED: None
' DYNAMIC STRING REFERENCES: None.
' LastScanTimestamp: Saturday 8/29/26 10:23 AM
' Purpose:
'     Owns the intentionally change-prone approved-build identity and fixed guarded-module roster consumed by x__UTIL_ADC_MODULE_GUARD.
' Author: Christina Shonts
' Support: Access Development Console Module Guard coordinated-build definition authority.
' Code Category: Access Development Console update assurance definitions.
' Requirements:
'     Companion x__UTIL_ADC_MODULE_GUARD and the complete guarded ADC module set.
' Notes:
'     Advance MODULE_GUARD_APPROVED_CURRENT_BUILD_BATCH_ID and this module's rewrite-tracking Version ID for every coordinated build that changes another guarded module.
'     Change the roster only when a guarded module is added, removed, renamed, or changes expected VBA component type.
'     Per-module change descriptions are intentionally not part of the active build-definition contract; historical descriptions already stored in Module Guard tables remain preserved.
' Design Map:
'     Approved-build batch identity; guarded module names; expected VBA component types.
' Configuration:
'     This module is the source-controlled changing definition surface for Module Guard. The guard engine remains stable unless Module Guard behavior or policy changes.
' Behavior:
'     Module Guard reads this roster, then reads current/previous Version IDs and timestamps from each installed module's managed rewrite header.
' Code Status: Active.
' Todo: None.
' Reporting Bugs: None currently recorded in this header.
' Module Level Bugs: None currently recorded in this header.
' Public Subs: None
' Public Functions:
'     ModuleGuardBaselineRosterDefinitions() -> Collection
' Public Properties: None
' ================================================================================================================================================================================

Option Compare Database
Option Explicit

' ===============================================================================================================================================================================
' MODULE-LEVEL DECLARATIONS
' ===============================================================================================================================================================================

' Approved coordinated-build identity. Advance this value for every guarded build that changes another guarded module.
Public Const MODULE_GUARD_APPROVED_CURRENT_BUILD_BATCH_ID As String = "2026-09-03-F1-Shared-Expected-Dependency-Snapshot"

' ===============================================================================================================================================================================
' PROCEDURES
' ===============================================================================================================================================================================

Public Function ModuleGuardBaselineRosterDefinitions() As Collection
    'CANONICAL AUTHORITY: fixed guarded-module membership and expected VBA component type. Runtime Version IDs and timestamps are read from managed module headers by Module Guard.
    Dim rows As New Collection
    rows.Add Array("x__UTIL_ADC_Core", "Standard Module")
    rows.Add Array("x__UTIL_ADC_MOC", "Standard Module")
    rows.Add Array("x__UTIL_ADC_MOC_LIFECYCLE", "Standard Module")
    rows.Add Array("x__UTIL_ADC_DB_REFERENCE", "Standard Module")
    rows.Add Array("x__UTIL_ADC_COMMON", "Standard Module")
    rows.Add Array("x__UTIL_ADC_MACRO_RUNTIME", "Standard Module")
    rows.Add Array("x__UTIL_ADC_DEFAULTS", "Standard Module")
    rows.Add Array("x__UTIL_ADC_DB_DEFINITIONS", "Standard Module")
    rows.Add Array("x__UTIL_ADC_MODULE_GUARD", "Standard Module")
    rows.Add Array("x__UTIL_ADC_MODULE_GUARD_DEFINITIONS", "Standard Module")
    rows.Add Array("x__UTIL_ADC_DB_INVENTORY_VBA", "Standard Module")
    rows.Add Array("x__UTIL_ADC_DB_INVENTORY_CANDIDATES", "Standard Module")
    rows.Add Array("x__UTIL_ADC_REPORTS", "Standard Module")
    rows.Add Array("x__UTIL_ADC_REVIEW", "Standard Module")
    rows.Add Array("x__UTIL_ADC_SETUP", "Standard Module")
    rows.Add Array("x__UTIL_ADC_TABLES", "Standard Module")
    rows.Add Array("x__UTIL_ADC_UI", "Standard Module")
    rows.Add Array("x__UTIL_ADC_UI_FORMS", "Standard Module")
    rows.Add Array("x__UTIL_ADC_UI_RIBBON", "Standard Module")
    rows.Add Array("x__UTIL_ADC_UI_HELP", "Standard Module")
    rows.Add Array("x__UTIL_ADC_UI_REVIEW", "Standard Module")
    Set ModuleGuardBaselineRosterDefinitions = rows
End Function
