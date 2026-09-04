# RVTools extraction example

Preserve the original RVTools workbook as evidence and extract selected worksheets into a separate destination:

```bash
xlsheet2csv ./original/RVTools.xlsx \
  --output-root ./extracted \
  --include vMetaData \
  --include vInfo \
  --include vHost \
  --include vDatastore
```

Retain the generated run and workbook manifests with the CSV files. They record the source SHA-256 hash, worksheet identity, output hashes and conversion policy.

This is raw extraction only. It does not establish RVTools compatibility, deduplicate inventory, preserve performance history, reconcile other collectors or produce assessment conclusions.
