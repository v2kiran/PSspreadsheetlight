Function UnGroup-SLColumn  {


    <#

.SYNOPSIS
    UnGroup columns by name or index.

.DESCRIPTION
    UnGroup columns by name or index.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER StartColumnName
    The columnName from which columns are to be UnGrouped. Eg: B.

.PARAMETER EndColumnName
    The columnName upto which columns are to be UnGrouped. Eg: D.

.PARAMETER StartColumnIndex
    The columnIndex from which columns are to be UnGrouped. Eg: 3.

.PARAMETER EndColumnIndex
    The columnIndex upto which columns are to be UnGrouped. Eg: 5.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | UnGroup-SLColumn -WorksheetName sheet5 -StartColumnName F -ENDColumnName H  -Verbose | Save-SLDocument


    Description
    -----------
    UnGroup columns F to H.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | UnGroup-SLColumn -WorksheetName sheet5 -StartColumnIndex 6 -ENDColumnIndex 8  -Verbose | Save-SLDocument


    Description
    -----------
    UnGroup columns 6 to 8.


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

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Name')]
        [string]$StartColumnName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Name')]
        [string]$ENDColumnName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'index')]
        [int]$StartColumnIndex,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'index')]
        [int]$ENDColumnIndex



    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'Name')
            {
                Write-Verbose ("UnGroup-SLColumn :`tUnGrouping Columns '{0}' to '{1}' " -f $StartColumnName, $ENDColumnName)
                $WorkBookInstance.UngroupColumns($StartColumnName, $ENDColumnName) | Out-Null
                $WorkBookInstance.UnhideColumn($StartColumnName, $ENDColumnName) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Index')
            {
                Write-Verbose ("UnGroup-SLColumn :`tUnGrouping Columns '{0}' to '{1}' " -f $StartColumnIndex, $ENDColumnIndex)
                $WorkBookInstance.UngroupColumns($StartColumnIndex, $ENDColumnIndex) | Out-Null
                $WorkBookInstance.UnhideColumn($StartColumnIndex, $ENDColumnIndex) | Out-Null
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }#select-slworksheet
    }#process
    END
    {
    }

}
