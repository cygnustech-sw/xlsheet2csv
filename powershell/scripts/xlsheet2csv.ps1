# xlsheet2csv - Export XLSX worksheets to CSV files.
# MIT License - see LICENSE file in the repository root for details.

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$SourcePath,

    [string]$DestinationPath,

    [ValidateSet('ExcelCom','ImportExcel')]
    [string]$Backend = 'ExcelCom',

    [switch]$Recurse,

    [string]$DateFormat = 'dd-MM-yyyy_HHmm',

    [string[]]$IncludeSheets,
    [string[]]$ExcludeSheets,

    [switch]$Visible,

    [string]$LogRoot
)

$scriptRoot = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$modulePath = Join-Path -Path (Split-Path -Path $scriptRoot -Parent) -ChildPath 'Xlsheet2Csv/Xlsheet2Csv.psm1'

Import-Module $modulePath -ErrorAction Stop

Export-XlsxWorkbookToCsv `
    -SourcePath $SourcePath `
    -DestinationPath $DestinationPath `
    -Backend $Backend `
    -Recurse:$Recurse `
    -DateFormat $DateFormat `
    -IncludeSheets $IncludeSheets `
    -ExcludeSheets $ExcludeSheets `
    -Visible:$Visible `
    -LogRoot $LogRoot
