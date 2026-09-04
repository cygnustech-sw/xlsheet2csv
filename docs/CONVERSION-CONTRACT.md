# Conversion contract

## Input

- `.xlsx` files only.
- Excel lock files beginning `~$` are ignored during directory discovery.
- Recursive discovery excludes the selected output root.
- Directory discovery ignores symlinks that resolve outside the requested input tree.
- Validation, hashing and worksheet parsing share one open input handle, so an atomic path replacement cannot substitute unchecked bytes or make the manifest describe a different workbook.
- Input archives must pass compressed-size, expanded-size and compression-ratio limits before openpyxl parses them.
- Archive member count and central-directory byte size are checked with a bounded preflight before Python materialises ZIP metadata or parses workbook XML. ZIP64 and multi-disk archives are rejected within this tool's lower XLSX size limits.
- XML parsing uses the installed `defusedxml` protection supported by openpyxl.

## Workbook and sheet identity

Each workbook directory contains a safe basename, the first 12 characters of the source SHA-256 hash and an eight-character hash of its path relative to the invocation root.

Each CSV begins with the one-based worksheet index. The original workbook path and worksheet name are retained in manifests. Sanitisation therefore cannot cause one valid export to overwrite another.

If the deterministic directory already exists, conversion fails for that workbook without changing the existing output.

## Cell values

- Strings remain strings; values such as `00123` are not re-inferred as numbers.
- Numeric cells are serialised from their underlying workbook value. Excel number-format display rules are not applied.
- Dates, datetimes and times use ISO 8601 text.
- Booleans use lowercase `true` and `false`.
- Empty cells become empty CSV fields. Trailing empty cells are omitted.
- Formula mode `values` reads cached results and never calculates formulas. An uncached formula may therefore be empty.
- Formula mode `formulas` writes formula text.
- Formula safety `escape`, the default, prefixes formula-like string values with an apostrophe. `preserve` disables this alteration and should be used only across a trusted downstream boundary.

## CSV

- UTF-8 without a byte-order mark by default; `utf-8-sig` is available.
- Comma delimiter and LF line endings by default.
- Python CSV quoting rules are used.
- Every selected sheet produces a file, including an empty file for an empty worksheet.

## Publication and failure

A workbook is written to a private staging directory and atomically renamed only after all selected worksheets and its manifest succeed. On failure, its staging directory is removed.

One bad workbook does not stop a batch. A unique run manifest records every success and failure. Partial success returns exit code `2`.
