# Access Development Console (ADC)

**Utility release:** ADC 3.0  
**State:** Repository preparation only — **authoritative source import and verification are NOT complete**.

> Do not use the files currently committed to this repository as a restored, compiled, or independently verified ADC application. The complete authoritative source has **not** been imported.

## Source of truth

The latest completed baseline supplied in the originating conversation is the **ADC 3.0 database-native baseline**, based exclusively on the original `xlib2.accdb`. This release name is independent of the existing ChatGPT module revision identifiers.

| Property | Value |
|---|---|
| Utility release | `ADC 3.0` (package-level declaration) |
| Native source | `xlib2.accdb` |
| Original database size | `252628992` bytes |
| Original database SHA-256 | `dfb116bb18e6e0556067584ca13eae9bd259f07382a1a6c877dc61a159f8b659` |
| Original database in GitHub | **Not uploaded** |
| Complete exported Access/VBA source | **Not uploaded** |
| Baseline verification | **Pending** |

The delivered ADC 3.0 native package preserves `xlib2.accdb` unchanged; it declares release 3.0 in package metadata but has **not** stamped the `ADCUtilityVersion` DAO property into the original database. A Windows finalization script in the conversation package makes a separate copy and applies that property.

## Architecture and preservation rules

ADC is a Microsoft Access/VBA utility. Its authoritative architecture, naming, setup-phase boundaries, settings/defaults, source provenance and Module Guard definitions must be discovered **only from the latest ADC 3.0 database**, not reconstructed from earlier exports.

- `x__UTIL_ADC_MODULE_GUARD` is a protected integrity/contract module.
- `x__UTIL_ADC_NO_EDIT_REFERENCE` is protected and must not be changed.
- Other VBA modules, class modules, form/report code-behind, queries, macros, database object definitions, data-backed configuration, dependencies and resources must be exported and inventoried from the **same** authoritative database.
- Preserve exact Access object names and existing namespaces; avoid automatic rename, repair, refactoring or migration during baseline acquisition.
- Preserve module-level `CGPT-...` version, previous-version, baseline-version and timestamp tracking. These are **not** utility release versions.
- Setup-phase ownership and cross-phase change restrictions remain in effect, including any Module Guard contracts.

## Development environment

To complete a verified Git source baseline:

1. Use Windows with desktop Microsoft Access and Trust Center authorization to inspect/export VBA via VBE. Prefer performing extraction on a **copy** of the source database, and hash-checking the untouched original both before and afterward.
2. Export every VBA component and Access form, report, macro and query using supported Access APIs (e.g. `VBComponents.Export` and Access `SaveAsText`, where applicable). Preserve each object's actual name and scope, and build an object inventory from the same database.
3. Export or document table schemas, relationships, database properties, references, startup configuration and other non-text/required dependencies. The complete original ACCDB remains essential for objects not faithfully representable as text.
4. Verify counts, names, per-object hashes, protection rules, Module Guard manifests, version chains and required application objects against the original database; run Access compile and applicable setup/regression tests.
5. Configure and **verify** Git LFS or another explicitly approved repository-backed binary solution before adding the full 252 MB ACCDB backup. Verify the remotely retrieved bytes and SHA-256; a Git LFS pointer alone is not proof of storage.
6. Commit the resulting complete source and original database snapshot as the first verified **baseline commit**, without substituting older files. Update this README's status only after verification.

## Expected repository layout (not yet populated)

- `src/` — text exports of current VBA and Access object definitions, with a source manifest.
- `resources/` — current, applicable dependencies, resources and settings/defaults.
- `database/` — preserved authoritative ACCDB backup (requires verified large-file storage).
- `docs/` — setup, restoration, integrity and regression documentation.
- `archive/legacy-baselines/` — **only after** full active-source import and remote verification: untouched earlier archives and a reconciliation manifest.

These paths describe intended layout; they do **not** imply files exist.

## Restoring the application

**Restore is not yet possible from this repository alone.** Retain the original ADC 3.0 native baseline ZIP and `xlib2.accdb` independently. Once initial import is complete, use the verified database backup, the same-version text exports and the verified manifest; validate Access references, configuration, Module Guard rules, startup/setup phases and runtime behavior before relying on the restored utility.

## Historical archives and safety

No ChatGPT Library historical baseline is authorized for deletion *until* its exact original bytes are committed, retrieved independently from GitHub (including LFS payload where used), and compared by SHA-256 and file size. Do not put historical source files into the active tree or rewrite prior Git history.

Do not commit secrets or personal data to this public repository without reviewing the Access database, source, and exported data. Large-file storage and privacy review remain prerequisites.
