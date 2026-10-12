# ADC 3.0 baseline import / verification record

Recorded: 2026-10-12 (America/New_York)  
**Overall status: BLOCKED — NOT INITIALIZED AS A COMPLETE ADC BASELINE.**

## Verified local authoritative material

| Item | Observed result |
|---|---|
| Local authoritative archive | `ADC_3.0_xlib2_NATIVE_BASELINE_20261012.zip` |
| Archive size | `52759166` bytes |
| Archive SHA-256 | `328d09a6570a5d721d94c0082b68f9ad9b8e81fe6254bf2fc3de8c18197074b0` |
| ZIP member count | `7` |
| ZIP CRC integrity | Passed |
| Database ZIP member | `xlib2.accdb` |
| Database size | `252628992` bytes |
| Database SHA-256 | `dfb116bb18e6e0556067584ca13eae9bd259f07382a1a6c877dc61a159f8b659` |
| Member/database SHA-256 compared with original upload and package manifest | All three matched |
| Release | `ADC 3.0`, package metadata only |
| Embedded `ADCUtilityVersion` DAO property stamped | **No — not yet applied** |
| Exported VBA modules in archive | `0` |
| Access forms/reports/macros/queries text exports in archive | `0` |
| Access compile / runtime verification | Not performed |

The source ZIP and original database were inspected locally; the database itself was **not modified**. No previous code was substituted.

## Active-source GitHub publication

**Not published:** complete VBA source, classes, form/report definitions, macros, queries, table schemas, dependencies, resources, settings/defaults and native database.

**Blocking conditions**

1. Microsoft Access/VBE/DAO Windows automation needed for faithful extraction of the source and Access objects is unavailable in this execution environment. The ZIP has only the database and support metadata/scripts; no usable current exported module/object set.
2. The original database is 252,628,992 bytes, exceeding normal GitHub per-file limits. A verified Git LFS remote or other approved binary storage solution is required. Git LFS was not available in the local execution environment or via the connected GitHub write tools.
3. GitHub connector text-file operations do not provide direct binary-file handoff from this local database to a verified remote LFS object.
4. Database object counts, module version chains, Module Guard definitions, protected source identities, and the setup-phase tests cannot be verified without exporting and inspecting the authoritative database.

This repository deliberately holds only documentation and tracking rules; it is **not** a backup from which the ADC application can be restored.

## Gate to consider initial import complete

- [ ] Export and enumerate all relevant Access and VBA objects from the **same** copied authoritative database, leaving original source bytes unchanged.
- [ ] Preserve, review, and compare exact names, protected modules, Module Guard authorities, version headers/chains, and all current setup-phase rules.
- [ ] Hash every export, record object scope/type/name, verify export counts and completeness against Access inventory.
- [ ] Publish all required source files, manifest, documentation, dependencies and the original native database backup.
- [ ] Independently fetch all published files plus **actual LFS payloads** where applicable, and verify hashes/sizes against the original.
- [ ] Verify a restoration and relevant Microsoft Access compile/regression tests; document any unavailable tests.
- [ ] Create a verified active-baseline commit; then (and only then) change this document to completed.

## Historical Library archive / deletion phase

**Not started, by design:** the user required independently verified active-source import before historical archival or ChatGPT Library deletion.

| Operation | Completed |
|---|---:|
| Historical baselines committed in `archive/legacy-baselines/` | 0 |
| Historical baselines independently fetched and hash-verified | 0 |
| Historical baselines deleted from ChatGPT Library | 0 |
| Complete historical baseline inventory | **Not established** |

Partial observations in the conversation do not constitute an exhaustive inventory of Library folders/ZIPs. No deletion of any source or historical baseline has been attempted. Library areas not fully inventoried include the root with pagination and all nested folders/surfaces, and older conversations that are not exposed as file-backed objects.

Do not infer archival completion or cleanup authorization from this preparation commit. Pending inventory must separately record originals, relative paths, filename collision handling, SHA-256/size, remote commit/path and deletion status, with private Library IDs only in the private reconciliation report.
