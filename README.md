# Access Development Console (ADC)

**Utility release:** ADC 3.0  
**State:** **ADC 3.0 full VBA baseline published and independently verified**; native Access object restoration and runtime tests remain outside this source baseline.

> The complete current **ADC utility VBA** source is in Git, but these source files alone do not restore the Access database or establish that the application compiles and runs in Microsoft Access.

## Source of truth

**Baseline definition:** The complete, exact VBA source exported from the latest authoritative ADC 3.0 `xlib2.accdb` is the **full version-controlled ADC baseline**. The database is a source/provenance and native restoration snapshot, **not a mandatory member of the Git-tracked VBA baseline**. The utility release name is independent of the existing ChatGPT module revision identifiers.

| Property | Value |
|---|---|
| Utility release | `ADC 3.0` (package-level declaration) |
| Native source | `xlib2.accdb` |
| Original database size | `252628992` bytes |
| Original database SHA-256 | `dfb116bb18e6e0556067584ca13eae9bd259f07382a1a6c877dc61a159f8b659` |
| Original database in GitHub | **Not uploaded; optional for VBA baseline** |
| Complete ADC VBA source (`x__UTIL_ADC` family) | **24 modules in `src/vba/`** |
| VBA baseline verification | **Passed** — GitHub Actions full SHA-256/ZIP check and byte-for-byte source readback |

The delivered ADC 3.0 native package preserves `xlib2.accdb` unchanged; it declares release 3.0 in package metadata but has **not** stamped the `ADCUtilityVersion` DAO property into the original database. A Windows finalization script in the conversation package makes a separate copy and applies that property.

## Architecture and preservation rules

ADC is a Microsoft Access/VBA utility. Its authoritative architecture, naming, setup-phase boundaries, settings/defaults, source provenance and Module Guard definitions must be discovered **only from the latest ADC 3.0 database**, not reconstructed from earlier exports.

- `x__UTIL_ADC_MODULE_GUARD` is a protected integrity/contract module.
- `x__UTIL_ADC_NO_EDIT_REFERENCE` is protected and must not be changed.
- The full ADC **utility** VBA baseline contains **all and only** VBA components with the exact module name `x__UTIL_ADC` or names beginning `x__UTIL_ADC_` from the authoritative database; export every matching component, including protected/no-edit modules, without code changes. Do not include unrelated modules. The delivery ZIP is flat and source-only, like prior full-baseline archives; keep its source integrity manifest outside the ZIP.
- Access-only objects (saved queries, macros, form/report layouts, table data/schema, embedded resources, database configuration) are separate **supplementary restoration assets**, not prerequisites for declaring the **VBA source baseline** complete. Preserve/export them when available, but never substitute older material.
- Preserve exact Access object names and existing namespaces; avoid automatic rename, repair, refactoring or migration during baseline acquisition.
- Preserve module-level `CGPT-...` version, previous-version, baseline-version and timestamp tracking. These are **not** utility release versions.
- Setup-phase ownership and cross-phase change restrictions remain in effect, including any Module Guard contracts.

## Published ADC 3.0 baseline

The 24 VBA modules were extracted read-only from `xlib2.accdb`, without replacing any module from a historical ZIP. Their original names, `CGPT-...` tracking headers, non-UTF-8 bytes and protected Module Guard/no-edit contents are preserved in [`src/vba/`](src/vba/).

The source-only [`full-baseline ZIP`](releases/ADC_3.0_x__UTIL_ADC_CGPT_CURRENT_FULL_BASELINE_20261012.zip) contains **only 24 flat `.bas` files**. The separate [`SHA-256 manifest`](docs/ADC_3.0_Full_VBA_Baseline_SHA256_Manifest.csv) records exact byte sizes and checksums. The [GitHub Actions verification](.github/workflows/verify-adc-baseline.yml) passed after checking out the committed repository: all 24 module hashes, ZIP SHA-256 `3e84af277417d76cc434f23006ff4d9d6bd95aa8babdad03c5f5a122fa48da3c`, ZIP CRC and every ZIP entry matched. A Windows Access compile or native database restoration **has not** been performed.

## Baseline source export tools

- [`tools/Build-ADC-3.0-Full-VBA-Baseline.ps1`](tools/Build-ADC-3.0-Full-VBA-Baseline.ps1) exports the specified `x__UTIL_ADC`-family modules from a **disposable copy** of the authoritative `xlib2.accdb` on Windows with Microsoft Access. It checks the original database SHA-256 before and after extraction and builds a flat, source-only ZIP plus a separate JSON manifest.
- [`tools/ADC_3.0_Full_VBA_Baseline_EXPORT_INSTRUCTIONS.txt`](tools/ADC_3.0_Full_VBA_Baseline_EXPORT_INSTRUCTIONS.txt) contains Windows execution steps and prerequisites.
- Scripts are **not** application modules. The latest baseline was recovered read-only from the original Access database without using the Windows exporter. GitHub Actions verified that the published ZIP matches all 24 committed modules and their SHA-256 manifest; the example August 26 exports were never substituted.

## Development environment

For subsequent **VBA** baseline exports and verification:

1. Use Windows with desktop Microsoft Access and Trust Center authorization to inspect/export VBA via VBE. Prefer performing extraction on a **copy** of the source database, and hash-checking the untouched original both before and afterward.
2. Export **every current VBA component**, including standard modules, classes, document modules and form/report code-behind, using supported Access/VBE APIs (e.g. `VBComponents.Export`). Preserve exact names and source bytes, and build a module inventory from the **same** database.
3. Independently verify VBA component counts, names, per-component hashes, Module Guard definitions, protected/no-edit modules and `CGPT-...` version chains against the current database. Run Access compile and applicable regression tests where possible; report tests not run.
4. Commit the **entire verified VBA source set** and its manifest as the active full baseline, without substituting older exports. Retrieve committed files from GitHub and verify their hashes before marking the baseline complete.
5. Optionally export Access-only objects (queries, macros, layouts, schema, configuration, resources) as additional reconstruction materials. Keep these clearly distinguished from the full **VBA** baseline.
6. Retain the unchanged database backup independently. Publishing it to GitHub is **optional** and requires independently verified large-file storage; the 252 MB binary and Git LFS do **not** block completing the VBA baseline.

## Repository layout

- [`src/vba/`](src/vba/) — 24 current, byte-preserved `x__UTIL_ADC`-family `.bas` modules, including Module Guard and protected reference modules.
- `resources/` — optional supplementary dependency/resources/settings exports.
- `database/` — optional authoritative ACCDB backup (requires verified large-file storage if published).
- [`docs/`](docs/) — source integrity manifests, setup/restoration requirements and verification records.
- [`releases/`](releases/) — original flat source-only [ADC 3.0 full-baseline ZIP](releases/ADC_3.0_x__UTIL_ADC_CGPT_CURRENT_FULL_BASELINE_20261012.zip).
- [`.github/workflows/verify-adc-baseline.yml`](.github/workflows/verify-adc-baseline.yml) — validates checked-out source hashes against the manifest and verifies the binary ZIP.
- `archive/legacy-baselines/` — historic reference archives, subject to independent remote verification and file-specific cleanup gates.

Only the linked directories/files above are currently published. The optional `resources/`, `database/` and historical archive directory may be incomplete or absent.

## Restoring the application

**A full VBA baseline is not the same as a fully reconstructable Access database.** Retain the original ADC 3.0 native baseline ZIP and `xlib2.accdb` for native restoration. Git now preserves the complete ADC 3.0 utility VBA source and CGPT module revision history; reconstructing an entire Access application additionally needs the appropriate database and non-VBA objects. Verify Module Guard, references, startup/setup phases and runtime behavior before relying on a restored utility.

## Historical archives and safety

No ChatGPT Library historical baseline is authorized for deletion *until* the current full **VBA** source baseline is imported and verified, and that specific historical file's exact original bytes are committed, retrieved independently from GitHub (including LFS payload where used), and compared by SHA-256 and file size. Do not put historical source files into the active tree or rewrite prior Git history.

Do not commit secrets or personal data to this public repository without reviewing the source and any optional exported data. Large-file storage is needed only for optional large binaries or historical archives.
