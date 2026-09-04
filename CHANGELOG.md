# Changelog

## 1.0.0 - 2026-09-04

- Replaced divergent pandas, Excel COM and ImportExcel implementations with one streaming Python engine.
- Added collision-free workbook and worksheet identities and immutable existing output.
- Preserved string identifiers without pandas type inference.
- Added an explicit conversion policy for encoding, delimiter, formulas, hidden sheets and formula safety.
- Added atomic workbook publication, per-workbook manifests, run manifests and output hashes.
- Added archive resource limits and `defusedxml` parser hardening.
- Added bounded central-directory preflight and a single stable handle for validation, hashing and parsing.
- Added partial-batch results and documented exit codes.
- Rebuilt the PowerShell module as a thin wrapper.
- Added tests, CI, packaging metadata, security guidance and release gates.
