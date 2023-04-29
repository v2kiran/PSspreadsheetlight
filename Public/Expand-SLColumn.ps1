Function Expand-SLColumn  {


    <#

.SYNOPSIS
    Expand columns by name or index.

.DESCRIPTION
    Expand columns by name or index.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER ColumnName
    The column name (such as "A1") of the column just after the group of columns you want to expand.
    For example, this will be column E if columns B to D are grouped.

.PARAMETER ColumnIndex
    The column index of the column just after the group of columns you want to expand.
    For example, this will be column 5 if columns 2 to 4 are grouped.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Expand-SLColumn   -WorksheetName sheet5 -ColumnName I  -Verbose | Save-SLDocument


    Description
    -----------
    Expand column I.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Expand-SLColumn   -WorksheetName sheet5 -ColumnIndex 9  -Verbose | Save-SLDocument


    Description
    -----------
    Expand column 9.


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
        [parameter(Mandatory = $true, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'index')]
        [int]$ColumnIndex,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Name')]
        [string]$ColumnName

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'Name')
            {
                Write-Verbose ("Expand-SLColumn :`tExpanding column '{0}' " -f $ColumnName)
                $WorkBookInstance.ExpandColumns($ColumnName) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Index')
            {
                Write-Verbose ("Expand-SLColumn :`tExpanding column '{0}' " -f $ColumnIndex)
                $WorkBookInstance.ExpandColumns($ColumnIndex) | Out-Null
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }#select-slworksheet
    }#process
    END
    {
    }

}
