Function Collapse-SLRow  {


    <#

.SYNOPSIS
    Collapse rows by index.

.DESCRIPTION
    Collapse rows by index.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER RowIndex
    The row index of the row just after the group of rows you want to collapse.
    For example, this will be row 5 if rows 2 to 4 are grouped.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Collapse-SLRow -WorksheetName sheet5 -RowIndex 7  -Verbose | Save-SLDocument


    Description
    -----------
    Collapse Row 7.


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
        [int]$RowIndex

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            Write-Verbose ("Collapse-SLRow :`tCollapsing Row '{0}' " -f $RowIndex)
            $WorkBookInstance.CollapseRows($RowIndex) | Out-Null
        }

        $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

    }#process
    END
    {
    }

}
