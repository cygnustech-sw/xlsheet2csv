Set-StrictMode -Version 2.0

function Export-XlsxWorkbookToCsv {
    <#
    .SYNOPSIS
    Invokes the canonical xlsheet2csv Python engine from PowerShell.

    .DESCRIPTION
    Preserves the PowerShell entry point without maintaining a second conversion implementation.
    Install the Python package first with `python -m pip install xlsheet2csv` or provide
    -PythonExecutable to invoke `python -m xlsheet2csv`.
    #>
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

    $resolvedSource = (Resolve-Path -LiteralPath $SourcePath -ErrorAction Stop).Path
    $target = if ($DestinationPath) { $DestinationPath } else { "$resolvedSource -> csv-export" }
    if (-not $PSCmdlet.ShouldProcess($target, 'Extract XLSX worksheets using xlsheet2csv')) {
        return
    }

    $arguments = @($resolvedSource, '--json', '--formulas', $Formulas, '--formula-safety', $FormulaSafety, '--hidden-sheets', $HiddenSheets, '--encoding', $Encoding, '--delimiter', $Delimiter)
    if ($DestinationPath) { $arguments += @('--output-root', $DestinationPath) }
    if ($Recurse) { $arguments += '--recurse' }
    foreach ($sheet in @($IncludeSheets)) { if ($sheet) { $arguments += @('--include', $sheet) } }
    foreach ($sheet in @($ExcludeSheets)) { if ($sheet) { $arguments += @('--exclude', $sheet) } }

    if ($PythonExecutable) {
        $pythonCommand = Get-Command $PythonExecutable -CommandType Application -ErrorAction Stop
        $json = & $pythonCommand.Source -m xlsheet2csv @arguments
    } else {
        $command = Get-Command xlsheet2csv -CommandType Application -ErrorAction Stop
        $json = & $command.Source @arguments
    }
    $exitCode = $LASTEXITCODE

    $result = $null
    if ($json) {
        $result = $json | ConvertFrom-Json
    }
    if ($exitCode -ne 0) {
        $summary = if ($result) { "$($result.success_count) succeeded and $($result.failure_count) failed" } else { 'no JSON result was returned' }
        throw "xlsheet2csv failed with exit code ${exitCode}: $summary."
    }
    return $result
}

Export-ModuleMember -Function Export-XlsxWorkbookToCsv
