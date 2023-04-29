Function UnGroup-SLRow  {


    <#

.SYNOPSIS
    UnGroup rows by index.

.DESCRIPTION
    UnGroup rows by index.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER StartRowIndex
    The rowIndex from which rows are to be ungrouped. Eg: 2.

.PARAMETER EndRowIndex
    The rowIndex upto which rows are to be ungrouped. Eg: 4.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | UnGroup-SLRow -WorksheetName sheet5 -StartRowIndex 4 -ENDRowIndex 6 -Verbose | Save-SLDocument


    Description
    -----------
    UnGroup Rows 4 to 6.


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
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true)]
        [int]$StartRowIndex,

        [parameter(Mandatory = $true, Position = 3)]
        [int]$ENDRowIndex


    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            Write-Verbose ("Group-SLRow :`tUnGrouping Rows '{0}' to '{1}' " -f $StartRowIndex, $ENDRowIndex)
            $WorkBookInstance.UnGroupRows($StartRowIndex, $ENDRowIndex) | Out-Null
            $WorkBookInstance.UnhideRow($StartRowIndex, $ENDRowIndex) | Out-Null
        }

        $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
    }
    END
    {
    }

}
