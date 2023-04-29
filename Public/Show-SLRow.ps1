Function Show-SLRow  {


    <#

.SYNOPSIS
    UnHide rows by index.

.DESCRIPTION
    UnHide rows by index.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER RowIndex
    The rowIndex that specifies the row to be shown. Eg: 2.

.PARAMETER StartRowIndex
    The rowIndex from which rows are to be shown. Eg: 2.

.PARAMETER EndRowIndex
    The rowIndex upto which rows are to be shown. Eg: 4.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Show-SLRow -WorksheetName sheet5 -RowIndex 4  -Verbose | Save-SLDocument


    Description
    -----------
    UnHide row 4.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Show-SLRow -WorksheetName sheet5 -StartRowIndex 3 -ENDRowIndex 4  -Verbose | Save-SLDocument


    Description
    -----------
    UnHide rows 3 & 4.

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
                Write-Verbose ("Show-SLRow :`tUn-Hiding Row '{0}' from worksheet '{1}' " -f $RowIndex, $WorksheetName)
                $WorkBookInstance.UnhideRow($RowIndex) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'RangeofRows')
            {
                Write-Verbose ("Show-SLRow :`tUn-Hiding Rows '{0}' to '{1}' " -f $StartRowIndex, $ENDRowIndex)
                $WorkBookInstance.UnhideRow($StartRowIndex, $ENDRowIndex) | Out-Null
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }#select-slworksheet

    }#process
    END
    {
    }

}
