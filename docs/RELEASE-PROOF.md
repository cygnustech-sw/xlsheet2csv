# Release proof record

| Gate | Required evidence | Status |
|---|---|---|
| Source integrity | Clean commit, version tag and reviewed diff | Pending |
| Python quality | Ruff and unittest pass on all supported Python versions | 13 local tests and Ruff pass on Python 3.12; CI matrix pending |
| Packaging | Wheel and source distribution build and install in a clean environment | Local clean-environment build/install pass on Python 3.12 |
| PowerShell | Manifest import and zero PSScriptAnalyzer warnings/errors | Local pass on PowerShell 7.6.5; Windows PowerShell 5.1 pending |
| Collision regression | Same basename, sanitised sheet names and existing-output cases pass | Local regression suite passes |
| Conversion corpus | Identifiers, Unicode, dates, formulae, empty/hidden sheets and malformed workbooks pass | Pending |
| Resource controls | File, expanded-size, compression-ratio and archive-member limits pass | Local regression suite passes |
| Cross-platform | Supported Windows, Linux and macOS matrix passes | Pending |
| Supply chain | Dependency review, SBOM and artefact checksums available | Pending |
| Documentation | Contract, security, support and examples match the release | Pending |

Do not describe a release as production-ready while any applicable gate remains pending.
