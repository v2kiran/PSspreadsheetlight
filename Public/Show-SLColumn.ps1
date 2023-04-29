Function Show-SLColumn  {


    <#

.SYNOPSIS
    Un-Hide columns by name or index.

.DESCRIPTION
    Un-Hide columns by name or index.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER ColumnName
    The columnName to be shown. Eg: B.

.PARAMETER ColumnIndex
    The columnIndex to be shown. Eg: 3.

.PARAMETER StartColumnName
    The columnName from which columns are to be shown. Eg: B.

.PARAMETER EndColumnName
    The columnName upto which columns are to be shown. Eg: D.

.PARAMETER StartColumnIndex
    The columnIndex from which columns are to be shown. Eg: 3.

.PARAMETER EndColumnIndex
    The columnIndex upto which columns are to be shown. Eg: 5.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Show-SLColumn -WorksheetName sheet5 -ColumnName B  -Verbose | Save-SLDocument


    Description
    -----------
    UnHide column B.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Show-SLColumn -WorksheetName sheet5 -ColumnIndex 3  -Verbose | Save-SLDocument


    Description
    -----------
    UnHide column 3(column C).


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Show-SLColumn -WorksheetName sheet5 -StartColumnName B -ENDColumnName C  -Verbose | Save-SLDocument


    Description
    -----------
    UnHide columns B to C.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Show-SLColumn -WorksheetName sheet5 -StartColumnIndex 4 -ENDColumnIndex 5  -Verbose | Save-SLDocument


    Description
    -----------
    UnHide columns 4 to 5.


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
                Write-Verbose ("Show-SLColumn :`tUn-Hiding Column '{0}' from worksheet '{1}' " -f $ColumnIndex, $WorksheetName)
                $WorkBookInstance.UnhideColumn($ColumnIndex) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'RangeofColumnsIndex')
            {
                Write-Verbose ("Show-SLColumn :`tUn-Hiding Columns '{0}' to '{1}' " -f $StartColumnIndex, $ENDColumnIndex)
                $WorkBookInstance.UnhideColumn($StartColumnIndex, $ENDColumnIndex ) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'SingleColumnName')
            {
                Write-Verbose ("Show-SLColumn :`tUn-Hiding Column '{0}' from worksheet '{1}' " -f $ColumnName, $WorksheetName)
                $WorkBookInstance.UnhideColumn($ColumnName) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'RangeofColumnsName')
            {
                Write-Verbose ("Show-SLColumn :`tUn-Hiding Columns '{0}' to '{1}' " -f $StartColumnName, $ENDColumnName)
                $WorkBookInstance.UnhideColumn($StartColumnName, $ENDColumnName ) | Out-Null
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }#select-slworksheet
    }#process
    END
    {
    }

}
