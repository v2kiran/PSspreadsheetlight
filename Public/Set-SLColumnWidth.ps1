Function Set-SLColumnWidth  {

    <#

.SYNOPSIS
    Set column width by name or index.

.DESCRIPTION
    Set column width by name or index.A single column or a range of columns can be specified as input.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.


.PARAMETER ColumnName
    The column name. Eg: A or B.

.PARAMETER ColumnIndex
    The column index. Eg: 2 or 5.

.PARAMETER StartColumnName
    Specifies the start of the column range. Eg: A.

.PARAMETER ENDColumnName
    Specifies the end of the column range. Eg: G.

.PARAMETER StartColumnIndex
    Specifies the start index of the column range. Eg: 1 .

.PARAMETER ENDColumnIndex
    Specifies the end index of the column range. Eg: 7.

.PARAMETER ColumnWidth
    Specifies the column width to be applied. Eg: 20.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLColumnWidth -WorksheetName sheet5 -ColumnName f -ColumnWidth 30 -Verbose | Save-SLDocument

    Description
    -----------
    Set columnwidth F to 30.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLColumnWidth -WorksheetName sheet5 -StartColumnName f -ENDColumnName h -ColumnWidth 30 -Verbose | Save-SLDocument

    Description
    -----------
    Set columnwidth of a range F - H to 30.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLColumnWidth -WorksheetName sheet5 -ColumnName f -ColumnWidth 30 -Verbose
    PS C:\> $doc | Set-SLColumnWidth -WorksheetName sheet5 -StartColumnIndex 7 -ENDColumnIndex 8  -ColumnWidth 15 -Verbose
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    Set columnwidth F to 30(Header column). Set columnwidth of column range 7-8 to 20(data columns)

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

        [Parameter(Mandatory = $true, Position = 2, ParameterSetName = 'SingleColumnName')]
        [String]$ColumnName,

        [parameter(Mandatory = $true, Position = 2, ParameterSetName = 'SingleColumnIndex')]
        [int]$ColumnIndex,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'MultiPleColumnName')]
        [string]$StartColumnName,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'MultiPleColumnName')]
        [string]$ENDColumnName,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'MultiPleColumnIndex')]
        [int]$StartColumnIndex,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'MultiPleColumnIndex')]
        [int]$ENDColumnIndex,

        [parameter(Mandatory = $true)]
        [Double]$ColumnWidth

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'SingleColumnName')
            {
                Write-Verbose ("Set-SLColumnWidth :`tSetting Width of Column '{0}' to '{1}'" -f $ColumnName, $ColumnWidth)
                $WorkBookInstance.SetColumnWidth($ColumnName, $ColumnWidth) | Out-Null
            }
            elseif ($PSCmdlet.ParameterSetName -eq 'SingleColumnIndex')
            {
                Write-Verbose ("Set-SLColumnWidth :`tSetting Width of Column '{0}' to '{1}'" -f $ColumnIndex, $ColumnWidth)
                $WorkBookInstance.SetColumnWidth($ColumnIndex, $ColumnWidth) | Out-Null
            }
            elseif ($PSCmdlet.ParameterSetName -eq 'MultiPleColumnName')
            {
                Write-Verbose ("Set-SLColumnWidth :`tSetting Width of Columns '{0}' to '{1}' to '{2}' " -f $StartColumnName, $ENDColumnName, $ColumnWidth)
                $WorkBookInstance.SetColumnWidth($StartColumnName, $ENDColumnName, $ColumnWidth) | Out-Null
            }
            elseif ($PSCmdlet.ParameterSetName -eq 'MultiPleColumnIndex')
            {
                Write-Verbose ("Set-SLColumnWidth :`tSetting Width of Columns '{0}' to '{1}' to '{2}' " -f $StartColumnIndex, $ENDColumnIndex, $ColumnWidth)
                $WorkBookInstance.SetColumnWidth($StartColumnIndex, $ENDColumnIndex, $ColumnWidth) | Out-Null
            }


            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }# select-slworksheet
    }#Process

}
