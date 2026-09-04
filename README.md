# xlsheet2csv

xlsheet2csv is a local, deterministic XLSX worksheet extractor from Cygnus Tech. It writes every selected worksheet to a distinct CSV file and records what happened in machine-readable manifests with SHA-256 hashes.

It is intended for infrastructure exports, reporting pipelines and other work where silent overwrites or undocumented type inference are unacceptable.

## Safety and data handling

- Processing is local. The tool does not upload workbooks or contact a Cygnus service.
- Output directories include source-content and relative-path hashes, preventing same-name workbook collisions.
- Sheet filenames include their workbook index, preventing sanitised-name collisions.
- Existing deterministic output is never overwritten.
- Each workbook is published atomically after every selected sheet succeeds.
- Batch conversion continues after a bad workbook and exits `2` for partial success.
- Input ZIP size, expanded size, compression ratio, sheet, row and column limits are enforced.
- Formula-like text is escaped by default for safer use in spreadsheet applications.

Read the exact [conversion contract](docs/CONVERSION-CONTRACT.md) before using CSV output as an interchange format.

## Install

Python 3.10 or newer is required.

```bash
python -m pip install .

# For an isolated command installation
pipx install .
```

## Use

```bash
# One workbook; csv-export is created beside the source
xlsheet2csv report.xlsx

# Recursive batch with an explicit destination
xlsheet2csv ./exports --recurse --output-root ./csv

# Exact, case-insensitive sheet filters
xlsheet2csv report.xlsx --include vInfo --include vHost

# Preserve formula-like text exactly only when the downstream trust boundary permits it
xlsheet2csv report.xlsx --formula-safety preserve

# Machine-readable run result on stdout
xlsheet2csv report.xlsx --json
```

Exit codes are `0` for complete success, `1` when nothing succeeded or invocation failed, and `2` for partial batch success.

## PowerShell wrapper

The PowerShell module is a thin wrapper around the same Python engine; it does not implement a second conversion path.

```powershell
Import-Module .\powershell\Xlsheet2Csv\Xlsheet2Csv.psd1
Export-XlsxWorkbookToCsv -SourcePath C:\Data\report.xlsx -DestinationPath C:\Data\csv
```

Use `-PythonExecutable` when the package is installed into a particular virtual environment.

## RVTools and similar exports

xlsheet2csv performs raw worksheet extraction. It does not infer inventory meaning, validate an RVTools version, reconcile multiple sources or produce Cygnus assessment findings. See [the RVTools example](examples/rvtools-example.md).

## Development

```bash
python -m pip install -e '.[dev]'
ruff check .
python -m unittest discover -s tests -v
python -m build
```

```powershell
$findings = Invoke-ScriptAnalyzer -Path ./powershell -Recurse -Severity Warning,Error
if ($findings) { $findings | Format-Table; throw 'PSScriptAnalyzer findings detected.' }
```

## Licence

MIT. See [LICENSE](LICENSE).
