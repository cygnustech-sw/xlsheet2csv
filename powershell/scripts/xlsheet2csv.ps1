#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory, Position = 0)][string]$SourcePath,
    [string]$DestinationPath,
    [switch]$Recurse,
    [string[]]$IncludeSheets,
    [string[]]$ExcludeSheets,
    [ValidateSet('values', 'formulas')][string]$Formulas = 'values',
    [ValidateSet('escape', 'preserve')][string]$FormulaSafety = 'escape',
    [ValidateSet('include', 'exclude')][string]$HiddenSheets = 'include',
    [ValidateSet('utf-8', 'utf-8-sig')][string]$Encoding = 'utf-8',
    [ValidateSet('comma', 'semicolon', 'tab', 'pipe')][string]$Delimiter = 'comma',
    [string]$PythonExecutable
)

$modulePath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Xlsheet2Csv\Xlsheet2Csv.psd1'
Import-Module $modulePath -Force -ErrorAction Stop

$parameters = @{
    SourcePath     = $SourcePath
    Recurse        = $Recurse
    IncludeSheets  = $IncludeSheets
    ExcludeSheets  = $ExcludeSheets
    Formulas       = $Formulas
    FormulaSafety  = $FormulaSafety
    HiddenSheets   = $HiddenSheets
    Encoding       = $Encoding
    Delimiter      = $Delimiter
    WhatIf         = $WhatIfPreference
}
if ($DestinationPath) { $parameters.DestinationPath = $DestinationPath }
if ($PythonExecutable) { $parameters.PythonExecutable = $PythonExecutable }

Export-XlsxWorkbookToCsv @parameters
