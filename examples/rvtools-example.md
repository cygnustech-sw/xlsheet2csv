# RVTools example

Example pattern for using xlsheet2csv with RVTools exports.

## PowerShell

```powershell
Import-Module "$PSScriptRoot/../powershell/Xlsheet2Csv/Xlsheet2Csv.psm1"

$results = Export-XlsxWorkbookToCsv -SourcePath "C:\Data\RVToolsExports" -Recurse -Backend ImportExcel

foreach ($result in $results) {
    "$($result.WorkbookPath) -> $($result.OutputFolder)"
}
```

## Python

```bash
cd python
xlsheet2csv "C:\Data\RVToolsExports" -o "C:\Data\RVToolsCsv" --recurse
```
