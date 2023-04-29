Function Insert-SLRow  {


    <#

.SYNOPSIS
    Insert rows by index.

.DESCRIPTION
    Insert rows by index.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER StartRowIndex
    The rowIndex before which rows are to be inserted. Eg: 3.

.PARAMETER NumberOfRows
    The number of columns to be inserted. Eg: 2.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Insert-SLRow -WorksheetName sheet5 -StartRowIndex 4 -NumberOfRows 2  -Verbose | Save-SLDocument


    Description
    -----------
    Insert 2 columns before row 4.


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
            Write-Verbose ("Insert-SLRow :`tInserting '{0}' Rows before Row '{1}' " -f $NumberOfRows, $StartRowIndex)
            $WorkBookInstance.InsertRow($StartRowIndex, $NumberOfRows) | Out-Null
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }
    }
    END
    {
    }

}
