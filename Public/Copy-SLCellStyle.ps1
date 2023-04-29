Function Copy-SLCellStyle  {


    <#

.SYNOPSIS
    Copy a style from a cell and apply it to another cell or a range of cells.

.DESCRIPTION
    Copy a style from a cell and apply it to another cell or a range of cells.
    Note: style can only be copied from a cell and not from a range of cells so the source is always going to be a single cell.
    The target howver can be either a single cell or a range of cells.
    #known issue - style is not applied if the style being copied is applied to another cell in the same row or column.
    Eg: Copy style from G10 to G4 or from G10 to D10 will not work.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER FromCellReference
    The source cell that contains the style to be copied. Eg: A5 or AB10

.PARAMETER ToCellReference
    The target cell that needs to have the copied cellstyle. Eg: A5 or AB10

.PARAMETER Range
    The target cell range that needs to have the copied cellstyle. Eg: A5:B10 or AB10:AD20


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Copy-SLCellStyle -WorksheetName sheet5 -FromCellReference g9 -ToCellReference f3  -Verbose | Save-SLDocument

    Description
    -----------
    Copy cellstyle from G9 and apply it to cell F3.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Copy-SLCellStyle -WorksheetName sheet5 -FromCellReference b10 -Range f4:h6  -Verbose | Save-SLDocument

    Description
    -----------
    Copy style from cell 'g10' and apply to Cell Range 'f4:h6'.



.INPUTS
   String,SpreadsheetLight.SLDocument

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    N/A
#>



    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true, Position = 0, ValueFromPipeLine = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [Alias('CurrentWorkSheetName')]
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [Alias('CellReference')]
        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Copy-SLCellStyle :`tFromCellReference should specify values in following format. Eg: A1,B10,AB5..etc"; break }
            })]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, position = 2)]
        [string]$FromCellReference,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Copy-SLCellStyle :`tToCellReference should specify values in following format. Eg: A1,B10,AB5..etc"; break }
            })]
        [parameter(Mandatory = $true, Position = 3, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Singlecell')]
        [string]$ToCellReference,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+:[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Copy-SLCellStyle :`tRange should specify values in following format. Eg: A1:D10 or AB1:AD5"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'MultipleCells')]
        [String]$Range

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {

            if ($PSCmdlet.ParameterSetName -eq 'Singlecell')
            {
                Write-Verbose ("Copy-SLCellStyle :`tCopy style from cell '{0}' to Cell '{1}'" -f $FromCellReference, $ToCellReference)
                $WorkBookInstance.CopyCellStyle($FromCellReference, $ToCellReference) | Out-Null
                $WorkBookInstance | Add-Member NoteProperty CellReference $ToCellReference -Force
            }
            elseif ($PSCmdlet.ParameterSetName -eq 'MultipleCells')
            {
                Write-Verbose ("Copy-SLCellStyle :`tCopy style from cell '{0}' and apply to Cell Range '{1}'" -f $FromCellReference, $Range)
                $rowindex, $columnindex = $range -split ':'
                $WorkBookInstance.CopyCellStyle($FromCellReference, $rowindex, $columnindex) | Out-Null
                $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
            }
        }

        $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

    }#process

}
