# Security policy

Report vulnerabilities privately to `security@cygnustech.co.uk`. Do not attach real customer workbooks or generated data to public issues.

The current release line receives security fixes. Older source snapshots are unsupported.

## Security boundaries

- Workbooks are processed locally and are never uploaded by the tool.
- ZIP resource limits are checked before parsing.
- XML parser hardening is installed through `defusedxml`.
- Workbook and worksheet naming cannot silently overwrite another conversion.
- Existing completed output is immutable.
- Formula-like text is escaped by default.
- The PowerShell module delegates to the canonical Python engine.

These controls reduce risk; they do not make arbitrary workbooks safe to process with unlimited resources. Do not expose the CLI directly as a public upload service. A hosted implementation would require process isolation, stricter quotas, malware controls, authentication, retention rules and monitoring.
