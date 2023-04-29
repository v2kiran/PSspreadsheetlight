Function Set-SLCellFormat  {


    <#

.SYNOPSIS
    Apply string formatting to cells.

.DESCRIPTION
    Apply string formatting to cells.A single or a range of cells can be specified as input

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER CellReference
    The target cell for stringformat. Eg: J3.

.PARAMETER Range
    The target range for stringformat. Eg: J3:K4.

.PARAMETER FormatString
    The format to be set on a particular cell or cells Eg. 'd mm yyyy'.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLCellValue -WorksheetName sheet5 -CellReference D4 -value 567890789 -Verbose |
                Set-SLCellFormat -FormatString '000\-00\-0000' -Verbose |
                    Save-SLDocument

    Description
    -----------
    Apply stringformat to cell d4. Here the format string - '000\-00\-0000' uses the pattern for matching a "social security number".

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Set-SLCellFormat -WorksheetName sheet5 -Range j3:l5 -FormatString '000\-00\-0000' -Verbose | Save-SLDocument


    Description
    -----------
    Formatstring is applied to a range j3:l5.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLColumnValue -WorksheetName sheet5 -CellReference B3 -value @(123456789.12345,-123456789.12345,(get-date),12.3456,12.3456,123456789.12345) -Verbose
    PS C:\> $doc | Set-SLCellFormat -WorksheetName sheet5 -CellReference B3 -FormatString '#,##0.000' -Verbose
    PS C:\> $doc | Set-SLCellFormat -WorksheetName sheet5 -CellReference B4 -FormatString '$#,##0.00_);[Red]($#,##0.00)' -Verbose
    PS C:\> $doc | Set-SLCellFormat -WorksheetName sheet5 -CellReference B5 -FormatString 'd mmm yyyy' -Verbose
    PS C:\> $doc | Set-SLCellFormat -WorksheetName sheet5 -CellReference B6 -FormatString '0.00%' -Verbose
    PS C:\> $doc | Set-SLCellFormat -WorksheetName sheet5 -CellReference B7 -FormatString '# ??/??' -Verbose
    PS C:\> $doc | Set-SLCellFormat -WorksheetName sheet5 -CellReference B8 -FormatString '0.000E+00' -Verbose
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    Example showing different formats applied to values in the cell range B3:B8.

.INPUTS
   String,SpreadsheetLight.SLDocument

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    http://office.microsoft.com/en-us/excel-help/number-format-codes-HP005198679.aspx
    http://www.databison.com/custom-format-in-excel-how-to-format-numbers-and-text/
#>


    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true, Position = 0, ValueFromPipeLine = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [Alias('CurrentWorkSheetName')]
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Set-SLCellFormat :`tCellReference should specify values in following format. Eg: A1,B10,AB5..etc"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true, ParameterSetname = 'cell')]
        [string[]]$CellReference,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+:[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Set-SLCellFormat :`tRange should specify values in following format. Eg: A1:D10 or AB1:AD5"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true, ParameterSetname = 'Range')]
        [string]$Range,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Position = 3)]
        [string]$FormatString


    )

    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'cell')
            {
                Foreach ($cref in $CellReference)
                {
                    $SLStyle = $WorkBookInstance.GetCellStyle($cref)
                    $SLStyle.FormatCode = $FormatString


                    Write-Verbose ("Set-SLCellFormat :`tSetting FormatString style '{0}' on Cell '{1}'" -f $FormatString, $cref)
                    $WorkBookInstance.SetCellStyle($Cref, $SLStyle) | Out-Null
                }
                $WorkBookInstance | Add-Member NoteProperty CellReference $CellReference -Force
            }
            elseif ($PSCmdlet.ParameterSetName -eq 'Range')
            {
                Write-Verbose ("Set-SLCellFormat :`tSetting FormatString style '{0}' on CellRange '{1}'" -f $FormatString, $Range)
                $StartCellReference, $ENDCellReference = $Range -split ':'

                $SLStyle = $WorkBookInstance.CreateStyle()
                $SLStyle.FormatCode = $FormatString
                $WorkBookInstance.SetCellStyle($StartCellReference, $ENDCellReference, $SLStyle) | Out-Null

                $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
            }#if parameterset range

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }#select worksheet

    }#Process

}
