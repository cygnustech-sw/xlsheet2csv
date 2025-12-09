# xlsheet2csv - Export XLSX worksheets to CSV files.
# MIT License - see LICENSE file in the repository root for details.

# Xlsheet2Csv PowerShell module
# Provides Export-XlsxWorkbookToCsv with:
#   - Excel COM backend (requires Excel)
#   - ImportExcel backend (no Excel dependency)
# Logging:
#   - Default: export.log in each workbook's output folder
#   - Override: -LogRoot for centralised logs

using namespace System.IO

function Write-Xlsheet2CsvLog {
    param(
        [string]$Message,
        [string]$LogFile
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$timestamp] $Message"

    Write-Host $line
    if ($LogFile) {
        $logDir = Split-Path -Path $LogFile -Parent
        if ($logDir -and -not (Test-Path -LiteralPath $logDir)) {
            New-Item -ItemType Directory -Path $logDir | Out-Null
        }
        Add-Content -Path $LogFile -Value $line
    }
}

function New-Xlsheet2CsvOutputFolder {
    param(
        [string]$WorkbookPath,
        [string]$DestinationRoot,
        [string]$DateFormat
    )

    $baseName = [Path]::GetFileNameWithoutExtension($WorkbookPath)
    $safeBase = $baseName -replace '[\\/:*?"<>|]', '_'
    $timestamp = Get-Date -Format $DateFormat
    $folderName = "{0}_{1}" -f $safeBase, $timestamp
    $folderPath = Join-Path -Path $DestinationRoot -ChildPath $folderName

    if (-not (Test-Path -LiteralPath $folderPath)) {
        New-Item -ItemType Directory -Path $folderPath | Out-Null
    }

    return $folderPath
}

function Export-XlsxWorkbookToCsv_ExcelCom {
    param(
        [string]$WorkbookPath,
        [string]$OutputFolder,
        [string[]]$IncludeSheets,
        [string[]]$ExcludeSheets,
        [string]$LogFile,
        [switch]$Visible
    )

    Write-Xlsheet2CsvLog -Message "Excel COM backend: starting workbook '$WorkbookPath'" -LogFile $LogFile

    try {
        $excel = New-Object -ComObject Excel.Application
    } catch {
        throw "Failed to create Excel COM object. Ensure Excel is installed. $_"
    }

    $excel.Visible = [bool]$Visible
    $excel.DisplayAlerts = $false

    try {
        Write-Xlsheet2CsvLog -Message "Opening workbook..." -LogFile $LogFile
        $workbook = $excel.Workbooks.Open($WorkbookPath)

        $allSheets = @($workbook.Worksheets)
        Write-Xlsheet2CsvLog -Message ("Workbook has {0} sheet(s)." -f $allSheets.Count) -LogFile $LogFile

        $sheets = $allSheets

        if ($IncludeSheets -and $IncludeSheets.Count -gt 0) {
            $includeSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            foreach ($name in $IncludeSheets) { [void]$includeSet.Add($name) }
            $sheets = $sheets | Where-Object { $includeSet.Contains($_.Name) }
            Write-Xlsheet2CsvLog -Message ("Filtered to {0} sheet(s) via IncludeSheets." -f $sheets.Count) -LogFile $LogFile
        }

        if ($ExcludeSheets -and $ExcludeSheets.Count -gt 0) {
            $excludeSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            foreach ($name in $ExcludeSheets) { [void]$excludeSet.Add($name) }
            $sheets = $sheets | Where-Object { -not $excludeSet.Contains($_.Name) }
            Write-Xlsheet2CsvLog -Message ("After ExcludeSheets, {0} sheet(s) remain." -f $sheets.Count) -LogFile $LogFile
        }

        $sheetIndex = 0
        $csvFiles = @()
        $sheetNames = @()

        foreach ($sheet in $sheets) {
            $sheetIndex++
            $sheetName = $sheet.Name
            $sheetNames += $sheetName

            Write-Xlsheet2CsvLog -Message ("[{0}/{1}] Exporting sheet '{2}'..." -f $sheetIndex, $sheets.Count, $sheetName) -LogFile $LogFile

            $safeName = $sheetName -replace '[\\/:*?"<>|]', '_'
            $csvPath = Join-Path -Path $OutputFolder -ChildPath ("{0}.csv" -f $safeName)

            $newWb = $excel.Workbooks.Add()
            $targetSheet = $newWb.Worksheets.Item(1)

            $sheet.UsedRange.Copy($targetSheet.Range("A1"))

            Write-Xlsheet2CsvLog -Message ("Saving to '{0}'" -f $csvPath) -LogFile $LogFile
            $newWb.SaveAs($csvPath, 6) # 6 = xlCSV
            $newWb.Close($false)

            $csvFiles += $csvPath
        }

        Write-Xlsheet2CsvLog -Message "Closing workbook." -LogFile $LogFile
        $workbook.Close($false)

        return [pscustomobject]@{
            Backend        = 'ExcelCom'
            WorkbookPath   = $WorkbookPath
            OutputFolder   = $OutputFolder
            SheetsExported = $sheetNames
            CsvFiles       = $csvFiles
            LogFile        = $LogFile
        }
    }
    finally {
        if ($excel) {
            $excel.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
        Write-Xlsheet2CsvLog -Message "Excel COM backend: completed workbook." -LogFile $LogFile
    }
}

function Export-XlsxWorkbookToCsv_ImportExcel {
    param(
        [string]$WorkbookPath,
        [string]$OutputFolder,
        [string[]]$IncludeSheets,
        [string[]]$ExcludeSheets,
        [string]$LogFile
    )

    Write-Xlsheet2CsvLog -Message "ImportExcel backend: starting workbook '$WorkbookPath'" -LogFile $LogFile

    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        throw "ImportExcel module not found. Install it with 'Install-Module ImportExcel'."
    }

    Import-Module ImportExcel -ErrorAction Stop

    $sheetInfo = Get-ExcelSheetInfo -Path $WorkbookPath
    $allSheets = $sheetInfo.Name
    Write-Xlsheet2CsvLog -Message ("Workbook has {0} sheet(s)." -f $allSheets.Count) -LogFile $LogFile

    $sheets = $allSheets

    if ($IncludeSheets -and $IncludeSheets.Count -gt 0) {
        $includeSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($name in $IncludeSheets) { [void]$includeSet.Add($name) }
        $sheets = $sheets | Where-Object { $includeSet.Contains($_) }
        Write-Xlsheet2CsvLog -Message ("Filtered to {0} sheet(s) via IncludeSheets." -f $sheets.Count) -LogFile $LogFile
    }

    if ($ExcludeSheets -and $ExcludeSheets.Count -gt 0) {
        $excludeSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($name in $ExcludeSheets) { [void]$excludeSet.Add($name) }
        $sheets = $sheets | Where-Object { -not $excludeSet.Contains($_) }
        Write-Xlsheet2CsvLog -Message ("After ExcludeSheets, {0} sheet(s) remain." -f $sheets.Count) -LogFile $LogFile
    }

    $sheetIndex = 0
    $csvFiles = @()
    $sheetNames = @()

    foreach ($sheetName in $sheets) {
        $sheetIndex++
        $sheetNames += $sheetName

        Write-Xlsheet2CsvLog -Message ("[{0}/{1}] Exporting sheet '{2}'..." -f $sheetIndex, $sheets.Count, $sheetName) -LogFile $LogFile

        $safeName = $sheetName -replace '[\\/:*?"<>|]', '_'
        $csvPath = Join-Path -Path $OutputFolder -ChildPath ("{0}.csv" -f $safeName)

        $data = Import-Excel -Path $WorkbookPath -WorksheetName $sheetName
        Write-Xlsheet2CsvLog -Message ("Saving to '{0}'" -f $csvPath) -LogFile $LogFile
        $data | Export-Csv -Path $csvPath -NoTypeInformation

        $csvFiles += $csvPath
    }

    Write-Xlsheet2CsvLog -Message "ImportExcel backend: completed workbook." -LogFile $LogFile

    return [pscustomobject]@{
        Backend        = 'ImportExcel'
        WorkbookPath   = $WorkbookPath
        OutputFolder   = $OutputFolder
        SheetsExported = $sheetNames
        CsvFiles       = $csvFiles
        LogFile        = $LogFile
    }
}

function Export-XlsxWorkbookToCsv {
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

    $ErrorActionPreference = 'Stop'

    $resolved = Resolve-Path -Path $SourcePath -ErrorAction Stop
    $item = Get-Item -LiteralPath $resolved

    if (-not $DestinationPath) {
        if ($item.PSIsContainer) {
            $DestinationPath = Join-Path -Path $item.FullName -ChildPath 'csv-export'
        } else {
            $DestinationPath = Join-Path -Path $item.DirectoryName -ChildPath 'csv-export'
        }
    }

    if (-not (Test-Path -LiteralPath $DestinationPath)) {
        New-Item -ItemType Directory -Path $DestinationPath | Out-Null
    }

    if ($item.PSIsContainer) {
        $searchPath = $item.FullName
        $files = Get-ChildItem -Path $searchPath -Filter '*.xlsx' -File -Recurse:$Recurse |
            Where-Object { -not $_.Name.StartsWith('~$') }
    } else {
        $files = @($item)
    }

    if (-not $files -or $files.Count -eq 0) {
        Write-Warning "No .xlsx files found for SourcePath '$SourcePath'."
        return
    }

    $results = @()
    $index = 0
    foreach ($file in $files) {
        $index++
        $workbookPath = $file.FullName
        $outputFolder = New-Xlsheet2CsvOutputFolder -WorkbookPath $workbookPath -DestinationRoot $DestinationPath -DateFormat $DateFormat

        if ($LogRoot) {
            if (-not (Test-Path -LiteralPath $LogRoot)) {
                New-Item -ItemType Directory -Path $LogRoot | Out-Null
            }
            $safeName = ([Path]::GetFileNameWithoutExtension($workbookPath)) -replace '[\\/:*?"<>|]', '_'
            $timestamp = Get-Date -Format $DateFormat
            $logFile = Join-Path -Path $LogRoot -ChildPath ("{0}_{1}.log" -f $safeName, $timestamp)
        } else {
            $logFile = Join-Path -Path $outputFolder -ChildPath 'export.log'
        }

        Write-Xlsheet2CsvLog -Message ("[{0}/{1}] Processing workbook '{2}'" -f $index, $files.Count, $workbookPath) -LogFile $logFile

        switch ($Backend) {
            'ExcelCom' {
                $result = Export-XlsxWorkbookToCsv_ExcelCom -WorkbookPath $workbookPath -OutputFolder $outputFolder -IncludeSheets $IncludeSheets -ExcludeSheets $ExcludeSheets -LogFile $logFile -Visible:$Visible
            }
            'ImportExcel' {
                $result = Export-XlsxWorkbookToCsv_ImportExcel -WorkbookPath $workbookPath -OutputFolder $outputFolder -IncludeSheets $IncludeSheets -ExcludeSheets $ExcludeSheets -LogFile $logFile
            }
        }

        $results += $result
    }

    return $results
}

Export-ModuleMember -Function Export-XlsxWorkbookToCsv
