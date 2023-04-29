Function Remove-SLRow  {


    <#

.SYNOPSIS
    Delete rows by index.

.DESCRIPTION
    Delete rows by index.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER StartRowIndex
    The rowIndex from which rows are to be deleted. Eg: 2.

.PARAMETER NumberOfRows
    The number of rows to be deleted. Eg: 2.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Remove-SLRow -WorksheetName sheet5 -StartRowIndex 4 -NumberOfRows 2  -Verbose | Save-SLDocument


    Description
    -----------
    Delete 2 rows starting from row 4 and moving down.
    Note: The count starts from the startrowindex so row 4 in this case will be deleted.


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
        [int]$NumberOfRows


    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            Write-Verbose ("Remove-SLRow :`tDeleting '{0}' Rows starting from Row '{1}' " -f $NumberOfRows, $StartRowIndex)
            $WorkBookInstance.DeleteRow($StartRowIndex, $NumberOfRows) | Out-Null
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }
    }
    END
    {
    }

}
