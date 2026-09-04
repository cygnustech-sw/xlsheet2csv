@{
    RootModule        = 'Xlsheet2Csv.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = '9f7ee1e6-7e9e-4e4c-8eab-3a1b870f46f5'
    Author            = 'Cygnus Tech Consulting Ltd'
    CompanyName       = 'Cygnus Tech Consulting Ltd'
    Copyright         = '(c) 2026 Cygnus Tech Consulting Ltd. MIT.'
    Description       = 'PowerShell wrapper for the canonical xlsheet2csv Python CLI.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @('Export-XlsxWorkbookToCsv')
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
    PrivateData       = @{
        PSData = @{
            Tags       = @('XLSX', 'CSV', 'Conversion')
            LicenseUri = 'https://opensource.org/license/mit'
            ProjectUri = 'https://github.com/cygnustech-sw/xlsheet2csv'
        }
    }
}
