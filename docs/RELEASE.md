# Release process

1. Update the project and PowerShell module versions and changelog.
2. Run Ruff, unittest, Python package build and PSScriptAnalyzer.
3. Install the built wheel into a clean virtual environment and run `xlsheet2csv --version` plus a representative conversion.
4. Run the workbook corpus across every supported Python and operating-system combination.
5. Review dependency vulnerability results and generate an SBOM.
6. Build wheel, source distribution and any separately approved portable executable.
7. Sign distributable executable or PowerShell artefacts where applicable and publish SHA-256 checksums.
8. Verify the downloaded artefacts on clean systems before promotion on cygnustech.co.uk.

Do not provide a hosted workbook upload endpoint from this repository.
