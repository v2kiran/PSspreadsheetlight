Function Convert-ToExcelRange {
    Param(
        [parameter(Mandatory=$true,Position=0)]
        [Int]$StartRowIndex,
        [parameter(Mandatory=$true,Position=1)]
        [Int]$StartColumnIndex,
        [parameter(Mandatory=$true,Position=2)]
        [Int]$EndRowIndex,
        [parameter(Mandatory=$true,Position=3)]
        [Int]$EndColumnIndex,
        [parameter(Mandatory=$false,Position=4)]
        [string]$WorkSheetName

    )

    if($WorkSheetName)
    {
        [SpreadsheetLight.SLConvert]::ToCellRange($WorkSheetName,$StartRowIndex,$StartColumnIndex,$ENDRowIndex,$ENDColumnIndex)
    }
    Else
    {
        [SpreadsheetLight.SLConvert]::ToCellRange($StartRowIndex,$StartColumnIndex,$ENDRowIndex,$ENDColumnIndex)
    }

}