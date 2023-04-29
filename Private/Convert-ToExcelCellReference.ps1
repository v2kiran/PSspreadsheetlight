Function Convert-ToExcelCellReference {
    Param(
        [parameter(Mandatory=$true,Position=0)]
        [Int]$Row,
        [parameter(Mandatory=$true,Position=1)]
        [Int]$Column
    )

    $cReference = [SpreadsheetLight.SLDocument]::WhatIsCellReference($Row,$Column)
    Write-Output ($cReference + ":" + $cReference)
}