# Support policy

The source release supports Python 3.10 through 3.13 on currently supported Windows, Linux and macOS versions after the release matrix passes.

The optional PowerShell wrapper supports Windows PowerShell 5.1 and PowerShell 7 where the Python package is already installed. Excel and Excel COM automation are neither required nor supported.

Only `.xlsx` input is supported. Legacy `.xls`, macro-enabled `.xlsm`, encrypted workbooks and password-protected files are outside the contract.

For a defect report, include the tool version, Python version, operating system, redacted run manifest and a minimal synthetic workbook when possible. Do not submit customer data.
