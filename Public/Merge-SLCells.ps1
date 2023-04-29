Function Merge-SLCells  {


    <#

.SYNOPSIS
    Merge cells.

.DESCRIPTION
    Merge cells.No merging is done if it's just one cell.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER Range
    The range of cells to be merged. Eg: J3:K4.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Merge-SLCells -WorksheetName sheet5 -Range j3:k4 -Verbose | Save-SLDocument

    Description
    -----------
    Merge cells in the range 'j3:k4'. The content of the first cell is displayed in the merged cell.


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

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+:[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Merge-SLCells :`tRange should specify values in following format. Eg: A1:D10 or AB1:AD5"; break }
            })]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Position = 2, ParameterSetName = 'CellReference')]
        [string]$Range


    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'CellReference')
            {
                Write-Verbose ("Merge-SLCells :`tMerging cells in the range '{0}'" -f $range)
                $StartCellReference, $ENDCellReference = $range -split ':'
                $WorkBookInstance.MergeWorksheetCells($StartCellReference, $ENDCellReference) | Out-Null
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }#select sl-worksheet
    }#process

}
