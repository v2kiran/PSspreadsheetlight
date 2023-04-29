Function Hide-SLColumn  {


    <#

.SYNOPSIS
    Hide columns by name or index.

.DESCRIPTION
    Hide columns by name or index.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER ColumnName
    The columnName to be hidden. Eg: B.

.PARAMETER ColumnIndex
    The columnIndex to be hidden. Eg: 3.

.PARAMETER StartColumnName
    The columnName from which columns are to be hidden. Eg: B.

.PARAMETER EndColumnName
    The columnName upto which columns are to be hidden. Eg: D.

.PARAMETER StartColumnIndex
    The columnIndex from which columns are to be hidden. Eg: 3.

.PARAMETER EndColumnIndex
    The columnIndex upto which columns are to be hidden. Eg: 5.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Hide-SLColumn -WorksheetName sheet5 -ColumnName B  -Verbose | Save-SLDocument


    Description
    -----------
    Hide column B.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Hide-SLColumn -WorksheetName sheet5 -ColumnIndex 3  -Verbose | Save-SLDocument


    Description
    -----------
    Hide column 3(column C).


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Hide-SLColumn -WorksheetName sheet5 -StartColumnName B -ENDColumnName C  -Verbose | Save-SLDocument


    Description
    -----------
    Hide columns B to C.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Hide-SLColumn -WorksheetName sheet5 -StartColumnIndex 4 -ENDColumnIndex 5  -Verbose | Save-SLDocument


    Description
    -----------
    Hide columns 4 to 5.


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

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleColumnName')]
        [string]$ColumnName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleColumnIndex')]
        [int]$ColumnIndex,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'RangeofColumnsName')]
        [string]$StartColumnName,

        [parameter(Mandatory = $true, Position = 3, ValueFromPipelineByPropertyName = $true, Parametersetname = 'RangeofColumnsName')]
        [string]$ENDColumnName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'RangeofColumnsIndex')]
        [int]$StartColumnIndex,

        [parameter(Mandatory = $true, Position = 3, ValueFromPipelineByPropertyName = $true, Parametersetname = 'RangeofColumnsIndex')]
        [int]$ENDColumnIndex


    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'SingleColumnIndex')
            {
                Write-Verbose ("Hide-SLColumn :`tHiding Column '{0}' from worksheet '{1}' " -f $ColumnIndex, $WorksheetName)
                $WorkBookInstance.HideColumn($ColumnIndex) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'RangeofColumnsIndex')
            {
                Write-Verbose ("Hide-SLColumn :`tHiding Columns '{0}' to '{1}' " -f $StartColumnIndex, $ENDColumnIndex)
                $WorkBookInstance.HideColumn($StartColumnIndex, $ENDColumnIndex ) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'SingleColumnName')
            {
                Write-Verbose ("Hide-SLColumn :`tHiding Column '{0}' from worksheet '{1}' " -f $ColumnName, $WorksheetName)
                $WorkBookInstance.HideColumn($ColumnName) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'RangeofColumnsName')
            {
                Write-Verbose ("Hide-SLColumn :`tHiding Columns '{0}' to '{1}' " -f $StartColumnName, $ENDColumnName)
                $WorkBookInstance.HideColumn($StartColumnName, $ENDColumnName ) | Out-Null
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }#select-slworksheet
    }#process
    END
    {
    }

}
