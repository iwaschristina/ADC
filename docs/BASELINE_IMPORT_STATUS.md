# ADC 3.0 VBA baseline import / verification record

Recorded: 2026-10-12 (America/New_York)  
**Current status: BLOCKED — authoritative full VBA export NOT YET PUBLISHED.**

## Definition of the complete active baseline (updated by project owner)

The complete set of authoritative VBA components whose names are **exactly `x__UTIL_ADC` or start with `x__UTIL_ADC_`** is the **full ADC 3.0 Git baseline**. It includes all matching standard modules, classes, document modules and form/report code-behind components (if any) present in the latest authoritative Access database; unrelated modules are excluded. Source must be extracted from one authoritative version, with its exact names, bytes, Module Guard contracts, protected/no-edit modules and CGPT version headers/chains preserved. A complete, verified VBA source manifest and GitHub readback hashes establish baseline completeness.

**An ACCDB binary, table data, saved queries, standalone macros, form/report layouts, embedded media, configuration/property exports, or other non-VBA database objects are _supplementary restoration assets_, not mandatory constituents of this defined VBA baseline.** Retain these as separate restoration materials when available. The full VBA baseline is not a claim that Git alone can restore the complete native Access database.

## Verified local authoritative material (retained outside GitHub)

| Item | Observed result |
|---|---|
| Authoritative package | `ADC_3.0_xlib2_NATIVE_BASELINE_20261012.zip` |
| ZIP size | `52759166` bytes |
| ZIP SHA-256 | `328d09a6570a5d721d94c0082b68f9ad9b8e81fe6254bf2fc3de8c18197074b0` |
| ZIP contents / CRC | 7 members; CRC passed |
| Original source database | `xlib2.accdb` |
| Database size | `252628992` bytes |
| Database SHA-256 | `dfb116bb18e6e0556067584ca13eae9bd259f07382a1a6c877dc61a159f8b659` |
| Database/member/manifest checksum comparison | Passed |
| Utility release | `ADC 3.0` in package metadata |
| Embedded DAO `ADCUtilityVersion` property | Not yet written to original database |
| Extracted current VBA files in authoritative ZIP | **0** |
| VBA component inventory / integrity verification | **Not available** |
| Native database modified by extraction | No |

No older module was substituted.

## Committed preparation utilities

The Windows-only exporter and execution instructions were committed to [`tools/`](../tools/); GitHub readback confirmed each Git blob matches the respective original generated file byte-for-byte. This exporter was **not executed here**; the environment has no Microsoft Access and the current source ZIP and sidecar manifest remain unavailable. Script preservation is not proof that the application was exported.

## GitHub active baseline publication

**Pending:** source files for every current VBA component and a verified component/integrity/version manifest.

**Actual blocking condition:** Microsoft Access/VBE automated extraction is unavailable in this environment. The native package includes an export script but no text export of the current VBA, so the exact authoritative code cannot be published or compared to a comprehensive module inventory here.

**Not a blocker to the VBA-only baseline:** The original ACCDB is 252,628,992 bytes and exceeds GitHub's direct file-size limit. Publishing an optional native backup requires verified Git LFS or equivalent storage. Since the project owner defined the **VBA source** as the full baseline, this does not prevent initializing a complete **VBA** baseline once authoritative exports are provided.

### Conditions required to certify the full VBA baseline

- [ ] Export all VBA components from the latest authoritative copied database (including standard, class, document, and form/report code-behind components), preserving original database bytes.
- [ ] Inventory and compare component type/name/count and exact source/bytes against the authoritative source database.
- [ ] Verify the current Module Guard, protected/no-edit modules including `x__UTIL_ADC_MODULE_GUARD` and `x__UTIL_ADC_NO_EDIT_REFERENCE`, CGPT identifiers and previous/baseline chains, dependencies, setup-phase ownership, and code integrity. Do not rewrite during initialization.
- [ ] Commit full source export and manifest to GitHub and independently read back all published files, verifying content hashes and counts.
- [ ] Report any unavailable Access compile/runtime/regression tests and restoration constraints; do not claim they ran.
- [ ] Mark GitHub's active VBA baseline complete only after all source requirements above pass.

Non-VBA exports and a native ACCDB backup are desirable for full Access application reconstruction but **not required** for certification of the Git-controlled full VBA baseline.

## Historical Library archival and cleanup

**Not started:** user-provided sequencing requires full active VBA-source import and independent remote verification first.

| Operation | Completed |
|---|---:|
| Historical files fully inventoried | No |
| Historical baseline archives committed | 0 |
| Historical originals independently verified against GitHub copies | 0 |
| Historical originals deleted from ChatGPT Library | 0 |

Accessible Library results in the conversation were partial. Library root pagination, some nested folders/surfaces and prior non-file-backed conversation assets were not exhaustively searched. Do not claim complete historical discovery. Preserve every historical original until its exact bytes are committed to `archive/legacy-baselines/`, independently downloaded, and compared by SHA-256 and size. Retain original relative paths and filenames. Do not change the active source tree or discard history. Any failure must leave Library originals untouched.
