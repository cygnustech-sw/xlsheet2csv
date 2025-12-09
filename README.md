# xlsheet2csv

xlsheet2csv provides utilities to export all worksheets in one or more XLSX workbooks to individual CSV files.

Supported variants:

- PowerShell module using Excel COM (requires Microsoft Excel on Windows)
- PowerShell module using ImportExcel (no Excel dependency, cross-platform PowerShell)
- Python CLI using pandas and openpyxl (no Excel dependency, cross-platform)

## PowerShell usage

Import the module and run against a single file or a folder of workbooks:

```powershell
Import-Module "$PSScriptRoot/powershell/Xlsheet2Csv/Xlsheet2Csv.psm1"

# Single workbook, Excel COM backend, default csv-export folder next to source
Export-XlsxWorkbookToCsv -SourcePath "C:\Data\report.xlsx"

# Folder of workbooks, Excel COM backend, explicit destination root
Export-XlsxWorkbookToCsv -SourcePath "C:\Data\Reports" -DestinationPath "C:\Data\Exports"

# ImportExcel backend (no Excel required, module must be installed)
Export-XlsxWorkbookToCsv -SourcePath "C:\Data\Reports" -Backend ImportExcel
```

Logs:

- By default, each workbook export writes a log file into its own export folder.
- Use `-LogRoot` to redirect logs to a central folder if required.

## Python usage

From the `python` folder (after installing dependencies):

```bash
cd python
python -m venv .venv
source .venv/bin/activate  # or .venv\Scripts\activate on Windows
pip install -e .

xlsheet2csv /path/to/report_or_folder
```

Logs:

- By default, a log file is written into each workbook's export folder.
- Use `--log-root` to redirect logs to a central folder.

## Dependencies

PowerShell:

- Excel COM backend:
  - Windows
  - Windows PowerShell 5.1 or PowerShell 7
  - Microsoft Excel installed and licensed

- ImportExcel backend:
  - PowerShell 5.1 or 7
  - ImportExcel module installed (`Install-Module ImportExcel`)

Python:

- Python 3.9+
- `pandas`
- `openpyxl`

## Licensing

This project is released under the MIT License.

The author reserves the right to offer alternative licensing or commercial
distributions of xlsheet2csv in the future.
