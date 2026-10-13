# ADC 3.0 VBA baseline — current verification status

Updated 2026-10-13. **Active VBA baseline published and verified. Historical cleanup incomplete.**

The authoritative VBA source consists of 24 unmodified `x__UTIL_ADC`-family `.bas` files extracted directly from the latest `xlib2.accdb`; it is separate from the native Access database backup. The `ADC 3.0` utility release does not replace per-module `CGPT-...` version identifiers.

- [Source](../src/vba/) — 24 modules, exact Git blob hashes and byte sizes independently checked against the original archive.
- [Original full-source ZIP](../releases/ADC_3.0_x__UTIL_ADC_CGPT_CURRENT_FULL_BASELINE_20261012.zip) — SHA-256 `3e84af277417d76cc434f23006ff4d9d6bd95aa8babdad03c5f5a122fa48da3c`.
- [SHA-256 manifest](ADC_3.0_Full_VBA_Baseline_SHA256_Manifest.csv) — 24 module checksums.
- GitHub Actions run `38314508442` succeeded on code verification commit `36da548beab8df2fcc00d6c197b1c68463106e11`, checking 24 module hashes, manifest, archive SHA-256, CRC and identical ZIP/source contents.
- Original Access database: 252,628,992 bytes, SHA-256 `dfb116bb18e6e0556067584ca13eae9bd259f07382a1a6c877dc61a159f8b659`; preserved independently and not modified.
- Protected `x__UTIL_ADC_MODULE_GUARD`, `x__UTIL_ADC_MODULE_GUARD_DEFINITIONS` and `x__UTIL_ADC_NO_EDIT_REFERENCE` source files are present and unchanged.
- **Not established:** Access compile/runtime tests, all non-VBA Access objects or full native database restoration.

Historical [archive manifest](../archive/legacy-baselines/manifest.csv): **1,924 ZIP candidates indexed**, not uploaded/verified/deleted. Further ADC-related files were observed in Library search and are not yet included; inventory is incomplete. Google Drive Library recursive browsing is unavailable. **No Library file has been deleted.** Before deleting any historical original, independently download its archived GitHub copy and compare both exact byte size and SHA-256 to the original. No archived legacy ZIP has yet passed that check.

The repository stores the verified active VBA baseline, not a verified complete native Microsoft Access application. Preserve existing commits and module-version chains; do not substitute older code.
