Function Convert-ToExcelColumnIndex {
    Param(
        [parameter(Mandatory=$true,Position=0)]
        [String]$ColumnName
    )

    [SpreadsheetLight.SLDocument]::WhatIsColumnIndex($ColumnName)
}