Function Hide-SLRow  {


    <#

.SYNOPSIS
    Hide rows by index.

.DESCRIPTION
    Hide rows by index.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER RowIndex
    The rowIndex that specifies the row to be hidden. Eg: 2.

.PARAMETER StartRowIndex
    The rowIndex from which rows are to be hidden. Eg: 2.

.PARAMETER EndRowIndex
    The rowIndex upto which rows are to be hidden. Eg: 4.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Hide-SLRow -WorksheetName sheet5 -RowIndex 4  -Verbose | Save-SLDocument


    Description
    -----------
    Hide row 4.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Hide-SLRow -WorksheetName sheet5 -StartRowIndex 3 -ENDRowIndex 4  -Verbose | Save-SLDocument


    Description
    -----------
    Hide rows 3 & 4.

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

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleRow')]
        [int]$RowIndex,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'RangeofRows')]
        [int]$StartRowIndex,

        [parameter(Mandatory = $true, Position = 3, ValueFromPipelineByPropertyName = $true, Parametersetname = 'RangeofRows')]
        [int]$ENDRowIndex


    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'SingleRow')
            {
                Write-Verbose ("Hide-SLRow :`tHiding Row '{0}' from worksheet '{1}' " -f $RowIndex, $WorksheetName)
                $WorkBookInstance.HideRow($RowIndex) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'RangeofRows')
            {
                Write-Verbose ("Hide-SLRow :`tHiding Rows '{0}' to '{1}' " -f $StartRowIndex, $ENDRowIndex)
                $WorkBookInstance.HideRow($StartRowIndex, $ENDRowIndex) | Out-Null
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }#select-slworksheet

    }#process
    END
    {
    }

}
