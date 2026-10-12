# Contributing and baseline safety

This document codifies the project owner's rules for **ADC 3.0** and subsequent releases.

## Authoritative code and provenance

1. The **full ADC Git baseline is the complete VBA source export** (all standard, class, document, and form/report code-behind modules) from the latest authoritative database; the `.accdb` binary and other non-VBA objects are separate native restoration materials, not required for declaring the VBA baseline complete. Until the full VBA import is verified, this repository remains a preparation scaffold.
2. Do not substitute older ZIPs, module exports, backups or reconstructed code for missing current VBA. Preserve exact exported source and the original native database snapshot independently.
3. Preserve Access object names, VBA module names, classes, identifiers, namespaces, directory layout and explicitly prescribed abbreviations.
4. The utility release (`3.0`) and module rewrite tracking (`CGPT-...`) are separate. Change a module's CGPT version chain only when the module is legitimately rewritten in accordance with its existing rewrite rules.
5. Never change `x__UTIL_ADC_MODULE_GUARD`, `x__UTIL_ADC_NO_EDIT_REFERENCE` or any other protected/no-edit component without explicit authorization.

## Boundaries and change control

- Preserve Module Guard architecture, definitions, expected module identities, dependency contracts, integrity checks and version chains. Update Module Guard definitions when a legitimate, authorized change requires it, without altering protected source by default.
- Keep the established setup-phase ownership, responsibility boundaries, source-provenance rules and public signatures.
- Cross-phase changes require explicit authorization.
- Follow the current baseline's naming, static-literal, defaults, configuration and verification policies; never invent a replacement policy from a historical module.
- Avoid unrelated cleanup, repair, refactoring, optimization, renaming or migration.
- Preserve original comments with the code they document.

## Before each commit

1. Inspect the exact authoritative current source; confirm version chains and module integrity.
2. Apply only approved modifications and record affected Access objects.
3. Export and hash **every VBA component** from a copied authoritative Access database; inventory all code-bearing component names and types and verify the original database hash has not changed.
4. Validate compile/runtime behavior and relevant setup/regression boundaries on Microsoft Access where feasible; disclose anything not tested.
5. Verify Module Guard updates and all protected/no-edit modules against their authorized baselines.
6. Publish the complete **VBA** baseline, a full-VBA-baseline ZIP and a changed-module ZIP for code changes, with hashes and manifests. Non-VBA Access objects and the ACCDB backup are supplementary restoration materials, not prerequisites to the full VBA baseline. Preserve prior Git commits and history.
7. Verify committed contents by fetching the GitHub copies; never rely only on upload responses.

## Historical baselines

Store verified prior baselines in `archive/legacy-baselines/` **after the complete current VBA baseline** has been independently verified in GitHub. Keep old ZIPs and original filenames/bytes intact. Any necessary name collision resolution must use distinct containing directories, never file renaming. Keep historical code separate from active source and retain every previous commit.

Delete a source from ChatGPT Library only when its specific remotely stored GitHub copy has been committed, retrieved, and independently checked against both byte size and SHA-256. Lack of Git LFS download verification means **no deletion**.
