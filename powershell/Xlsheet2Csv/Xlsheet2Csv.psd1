@{
    RootModule        = 'Xlsheet2Csv.psm1'
    ModuleVersion     = '0.1.0'
    GUID              = '9f7ee1e6-7e9e-4e4c-8eab-3a1b870f46f5'
    Author            = 'xlsheet2csv contributors'
    CompanyName       = 'Community'
    Description       = 'Export XLSX worksheets to individual CSV files using Excel COM or ImportExcel.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @(
        'Export-XlsxWorkbookToCsv'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}
