Function Convert-ToExcelRowColumnStats {
    Param(
        [parameter(Mandatory=$true,Position=0)]
        [String]$Range


    )
        $StartCellReference, $ENDCellReference = $Range -split ":"

        $refrow = $refrow1 = 0
        $refcolumn = $refcolumn1 = 0
        [SpreadsheetLight.SLDocument]::WhatIsRowColumnIndex($StartCellReference,[ref]$refRow,[ref]$refColumn) | Out-Null
        [SpreadsheetLight.SLDocument]::WhatIsRowColumnIndex($ENDCellReference,[ref]$refRow1,[ref]$refColumn1) | Out-Null


        $props = [Ordered]@{
            StartColumnName = [SpreadsheetLight.SLConvert]::ToColumnName($refColumn)
            StartColumnIndex = $refColumn
            StartRowIndex = $refRow
            EndColumnName = [SpreadsheetLight.SLConvert]::ToColumnName($refColumn1)
            EndColumnIndex = $refColumn1
            EndRowIndex = $refRow1 }

        New-Object PSobject -Property $props

}