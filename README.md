# Access Development Console (ADC)

**Utility release:** ADC 3.0  
**State:** Repository preparation only — **complete authoritative VBA baseline not yet imported or verified**.

> Do not use the files currently committed to this repository as a restored, compiled, or independently verified ADC application. The complete authoritative source has **not** been imported.

## Source of truth

**Baseline definition:** The complete, exact VBA source exported from the latest authoritative ADC 3.0 `xlib2.accdb` is the **full version-controlled ADC baseline**. The database is a source/provenance and native restoration snapshot, **not a mandatory member of the Git-tracked VBA baseline**. The utility release name is independent of the existing ChatGPT module revision identifiers.

| Property | Value |
|---|---|
| Utility release | `ADC 3.0` (package-level declaration) |
| Native source | `xlib2.accdb` |
| Original database size | `252628992` bytes |
| Original database SHA-256 | `dfb116bb18e6e0556067584ca13eae9bd259f07382a1a6c877dc61a159f8b659` |
| Original database in GitHub | **Not uploaded; optional for VBA baseline** |
| Complete exported VBA source | **Not uploaded** |
| Baseline verification | **Pending** |

The delivered ADC 3.0 native package preserves `xlib2.accdb` unchanged; it declares release 3.0 in package metadata but has **not** stamped the `ADCUtilityVersion` DAO property into the original database. A Windows finalization script in the conversation package makes a separate copy and applies that property.

## Architecture and preservation rules

ADC is a Microsoft Access/VBA utility. Its authoritative architecture, naming, setup-phase boundaries, settings/defaults, source provenance and Module Guard definitions must be discovered **only from the latest ADC 3.0 database**, not reconstructed from earlier exports.

- `x__UTIL_ADC_MODULE_GUARD` is a protected integrity/contract module.
- `x__UTIL_ADC_NO_EDIT_REFERENCE` is protected and must not be changed.
- The full VBA baseline includes **every VBA standard module, class module, document module, and form/report code-behind component** from the authoritative database, including protected/no-edit modules. Export exact current source without code changes.
- Access-only objects (saved queries, macros, form/report layouts, table data/schema, embedded resources, database configuration) are separate **supplementary restoration assets**, not prerequisites for declaring the **VBA source baseline** complete. Preserve/export them when available, but never substitute older material.
- Preserve exact Access object names and existing namespaces; avoid automatic rename, repair, refactoring or migration during baseline acquisition.
- Preserve module-level `CGPT-...` version, previous-version, baseline-version and timestamp tracking. These are **not** utility release versions.
- Setup-phase ownership and cross-phase change restrictions remain in effect, including any Module Guard contracts.

## Development environment

To complete a verified **VBA** Git source baseline:

1. Use Windows with desktop Microsoft Access and Trust Center authorization to inspect/export VBA via VBE. Prefer performing extraction on a **copy** of the source database, and hash-checking the untouched original both before and afterward.
2. Export **every current VBA component**, including standard modules, classes, document modules and form/report code-behind, using supported Access/VBE APIs (e.g. `VBComponents.Export`). Preserve exact names and source bytes, and build a module inventory from the **same** database.
3. Independently verify VBA component counts, names, per-component hashes, Module Guard definitions, protected/no-edit modules and `CGPT-...` version chains against the current database. Run Access compile and applicable regression tests where possible; report tests not run.
4. Commit the **entire verified VBA source set** and its manifest as the active full baseline, without substituting older exports. Retrieve committed files from GitHub and verify their hashes before marking the baseline complete.
5. Optionally export Access-only objects (queries, macros, layouts, schema, configuration, resources) as additional reconstruction materials. Keep these clearly distinguished from the full **VBA** baseline.
6. Retain the unchanged database backup independently. Publishing it to GitHub is **optional** and requires independently verified large-file storage; the 252 MB binary and Git LFS do **not** block completing the VBA baseline.

## Expected repository layout (not yet populated)

- `src/` — **complete current VBA source exports** (standard/class/document/form/report code-behind) and integrity manifest.
- `resources/` — optional supplementary dependency/resources/settings exports.
- `database/` — optional authoritative ACCDB backup (requires verified large-file storage if published).
- `docs/` — setup, restoration, integrity and regression documentation.
- `archive/legacy-baselines/` — **only after** full active-source import and remote verification: untouched earlier archives and a reconciliation manifest.

These paths describe intended layout; they do **not** imply files exist.

## Restoring the application

**A full VBA baseline is not the same as a fully reconstructable Access database.** Retain the original ADC 3.0 native baseline ZIP and `xlib2.accdb` for native restoration. Once source import is complete, Git will preserve the complete VBA code/version history; reconstructing an entire Access application additionally needs the appropriate database and non-VBA objects. Verify Module Guard, references, startup/setup phases and runtime behavior before relying on a restored utility.

## Historical archives and safety

No ChatGPT Library historical baseline is authorized for deletion *until* the current full **VBA** source baseline is imported and verified, and that specific historical file's exact original bytes are committed, retrieved independently from GitHub (including LFS payload where used), and compared by SHA-256 and file size. Do not put historical source files into the active tree or rewrite prior Git history.

Do not commit secrets or personal data to this public repository without reviewing the source and any optional exported data. Large-file storage is needed only for optional large binaries or historical archives.
