Function Convert-ToExcelAbsoluteRange {
    Param(
        [parameter(Mandatory=$true,Position=0)]
        [String]$Range,
        [parameter(Mandatory=$false,Position=1)]
        [string]$WorkSheetName

    )
    $r1,$r2 = $Range -split ":"
    $RC1 = Convert-ToExcelRowColumnIndex -CellReference $r1
    $RC2 = Convert-ToExcelRowColumnIndex -CellReference $r2
    if($WorkSheetName)
    {
        [SpreadsheetLight.SLConvert]::ToCellRange($WorkSheetName,$RC1.Row,$RC1.Column,$RC2.Row,$RC2.Column,$true)
    }
    Else
    {
        [SpreadsheetLight.SLConvert]::ToCellRange($RC1.Row,$RC1.Column,$RC2.Row,$RC2.Column,$true)
    }

}