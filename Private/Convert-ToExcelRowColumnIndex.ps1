Function Convert-ToExcelRowColumnIndex {
    Param(
        [validatepattern('[a-z]+\d+')]
        [parameter(Mandatory=$true,Position=0)]
        [String]$CellReference

    )
    $refrow = 0
    $refcolumn = 0
    [SpreadsheetLight.SLDocument]::WhatIsRowColumnIndex($CellReference,[ref]$refRow,[ref]$refColumn) | Out-Null
    New-Object PSObject -Property @{Row = $refRow;Column = $refColumn}
}