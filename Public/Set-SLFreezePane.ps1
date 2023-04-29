Function Set-SLFreezePane  {


    <#

.SYNOPSIS
    set up Split pane.

.DESCRIPTION
    set up Split pane.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER NumberOfTopMostRows
    Number of top-most rows to keep in place.

.PARAMETER NumberOfLeftMostColumns
   Number of left-most columns to keep in place.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Set-SLFreezePane -WorksheetName sheet5 -NumberOfTopMostRows 3 -NumberOfLeftMostColumns 8 -Verbose  | Save-SLDocument


    Description
    -----------
    Top-left pane is '3' Rows high and '8' columns wide.


.INPUTS
   String,Int,SpreadsheetLight.SLDocument

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
        [parameter(Mandatory = $true, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true)]
        [int]$NumberOfTopMostRows,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true)]
        [int]$NumberOfLeftMostColumns

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            Write-Verbose ("Set-SLFreezePane :`tTop-left pane is '{0}' Rows high and '{1}' columns wide. " -f $NumberOfTopMostRows, $NumberOfLeftMostColumns)
            $WorkBookInstance.FreezePanes($NumberOfTopMostRows, $NumberOfLeftMostColumns) | Out-Null
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }
    }
    END
    {
    }

}
