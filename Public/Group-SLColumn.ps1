Function Group-SLColumn  {


    <#

.SYNOPSIS
    Group columns by name or index.

.DESCRIPTION
    Group columns by name or index.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER StartColumnName
    The columnName from which columns are to be Grouped. Eg: B.

.PARAMETER EndColumnName
    The columnName upto which columns are to be Grouped. Eg: D.

.PARAMETER StartColumnIndex
    The columnIndex from which columns are to be Grouped. Eg: 3.

.PARAMETER EndColumnIndex
    The columnIndex upto which columns are to be Grouped. Eg: 5.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Group-SLColumn -WorksheetName sheet5 -StartColumnName F -ENDColumnName H  -Verbose | Save-SLDocument


    Description
    -----------
    Group columns F to H.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Group-SLColumn -WorksheetName sheet5 -StartColumnIndex 6 -ENDColumnIndex 8  -Verbose | Save-SLDocument


    Description
    -----------
    Group columns 6 to 8.


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
                Write-Verbose ("Group-SLColumn :`tGrouping Columns '{0}' to '{1}' " -f $StartColumnName, $ENDColumnName)
                $WorkBookInstance.GroupColumns($StartColumnName, $ENDColumnName) | Out-Null
                $WorkBookInstance.CollapseColumns(((Convert-ToExcelColumnIndex $ENDColumnName) + 1)) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Index')
            {
                Write-Verbose ("Group-SLColumn :`tGrouping Columns '{0}' to '{1}' " -f $StartColumnIndex, $ENDColumnIndex)
                $WorkBookInstance.GroupColumns($StartColumnIndex, $ENDColumnIndex) | Out-Null
                $WorkBookInstance.CollapseColumns(( $ENDColumnIndex + 1)) | Out-Null
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }#select-slworksheet
    }#process
    END
    {
    }

}
