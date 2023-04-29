Function Insert-SLColumn  {


    <#

.SYNOPSIS
    Insert columns by name or index.

.DESCRIPTION
    Insert columns by name or index.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER StartColumnName
    The columnName before which columns are to be inserted. Eg: B.

.PARAMETER StartColumnIndex
    The columnIndex before which columns are to be inserted. Eg: 3.

.PARAMETER NumberOfColumns
    The number of columns to be inserted. Eg: 2.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Insert-SLColumn -WorksheetName sheet5 -StartColumnName C -NumberOfColumns 2  -Verbose | Save-SLDocument


    Description
    -----------
    Insert 2 columns before column C.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Insert-SLColumn -WorksheetName sheet5 -StartColumnIndex 3 -NumberOfColumns 2  -Verbose | Save-SLDocument


    Description
    -----------
    Insert 2 columns before column 3(column C).


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

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'index')]
        [int]$StartColumnIndex,

        [parameter(Mandatory = $true, Position = 3)]
        [int]$NumberOfColumns


    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'Name')
            {
                Write-Verbose ("Insert-SLColumn :`tInserting '{0}' columns before column '{1}' " -f $NumberOfColumns, $StartColumnName)
                $WorkBookInstance.InsertColumn($StartColumnName, $NumberOfColumns) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Index')
            {
                Write-Verbose ("Insert-SLColumn :`tInserting '{0}' columns before column '{1}' " -f $NumberOfColumns, $StartColumnIndex)
                $WorkBookInstance.InsertColumn($StartColumnIndex, $NumberOfColumns) | Out-Null
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }#select-slworksheet

    }#process
    END
    {
    }

}
