Function Set-SLAutoFitColumn  {


    <#

.SYNOPSIS
    Autofit columns by ColumnName or ColumnIndex.

.DESCRIPTION
    Autofit columns by ColumnName or ColumnIndex. A single or a range of columns by name or index can be specified as input.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER ColumnName
    The columnName to be autofit. Eg: A or G.

.PARAMETER ColumnIndex
    The columnIndex to be autofit. Eg: 1 or 5.

.PARAMETER StartColumnName
    Specifies the start of the autofit column range. Eg: A.

.PARAMETER EndColumnName
    Specifies the end of the autofit column range. Eg: G.

.PARAMETER StartColumnIndex
    Specifies the start of the autofit column range. Eg: 1.

.PARAMETER EndColumnIndex
    Specifies the end of the autofit column range. Eg: 5.

.PARAMETER MaximumColumnWidth
    Specifies the maximum column width for a column or range after autofit is applied. Eg: 10.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLAutoFitColumn -WorksheetName sheet5 -ColumnName F -Verbose | Save-SLDocument

    Description
    -----------
    Autofit column F by Name .


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLAutoFitColumn -WorksheetName sheet5 -StartColumnName F -ENDColumnName H -Verbose | Save-SLDocument

    Description
    -----------
    Autofit columns from F to H by Name.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLAutoFitColumn -WorksheetName sheet5 -StartColumnName F -ENDColumnName H -MaximumColumnWidth 10 -Verbose | Save-SLDocument

    Description
    -----------
    Autofit columns from F to H by Name and optionally set a maxcolumnwidth of 10.

.INPUTS
   String,Int,SpreadsheetLight.SLDocument

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    N/A
#>


    [CmdletBinding(DefaultParameterSetName = 'All')]
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

        [parameter(Mandatory = $false)]
        [Double]$MaximumColumnWidth

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'All')
            {
                Write-Verbose ("Set-SLAutoFitColumn :`tSetting autofit on all columns in worksheet '{0}'" -f $worksheetname)
                $WorkBookInstance.AutoFitColumn('A', 'DD') | Out-Null
            }

            elseif ($PSCmdlet.ParameterSetName -eq 'SingleColumnName')
            {
                if ($MaximumColumnWidth)
                {
                    Write-Verbose ("Set-SLAutoFitColumn :`tSetting autofit on column '{0}' with maxcolumnwidth of '{1}'" -f $ColumnName, $MaximumColumnWidth)
                    $WorkBookInstance.AutoFitColumn($ColumnName, $MaximumColumnWidth) | Out-Null
                }
                Else
                {
                    Write-Verbose ("Set-SLAutoFitColumn :`tSetting autofit on column '{0}'" -f $columnName)
                    $WorkBookInstance.AutoFitColumn($columnName) | Out-Null
                }
            }
            elseif ($PSCmdlet.ParameterSetName -eq 'SingleColumnIndex')
            {
                if ($MaximumColumnWidth)
                {
                    Write-Verbose ("Set-SLAutoFitColumn :`tSetting autofit on column '{0}' with maxcolumnwidth of '{1}'" -f $ColumnIndex, $MaximumColumnWidth)
                    $WorkBookInstance.AutoFitColumn($ColumnIndex, $MaximumColumnWidth) | Out-Null
                }
                Else
                {
                    Write-Verbose ("Set-SLAutoFitColumn :`tSetting autofit on column '{0}'" -f $ColumnIndex)
                    $WorkBookInstance.AutoFitColumn($ColumnIndex) | Out-Null
                }

            }
            elseif ($PSCmdlet.ParameterSetName -eq 'MultiPleColumnName')
            {
                if ($MaximumColumnWidth)
                {
                    Write-Verbose ("Set-SLAutoFitColumn :`tSetting autofit on columns from '{0}' to '{1}' with maxcolumnwidth of '{2}'" -f $StartColumnName, $ENDColumnName, $MaximumColumnWidth)
                    $WorkBookInstance.AutoFitColumn($StartColumnName, $ENDColumnName, $MaximumColumnWidth) | Out-Null
                }
                Else
                {
                    Write-Verbose ("Set-SLAutoFitColumn :`tSetting autofit on columns from '{0}' to '{1}'" -f $StartColumnName, $ENDColumnName)
                    $WorkBookInstance.AutoFitColumn($StartColumnName, $ENDColumnName) | Out-Null
                }

            }
            elseif ($PSCmdlet.ParameterSetName -eq 'MultiPleColumnIndex')
            {
                if ($MaximumColumnWidth)
                {
                    Write-Verbose ("Set-SLAutoFitColumn :`tSetting autofit on columns from '{0}' to '{1}' with maxcolumnwidth of '{2}'" -f $StartColumnIndex, $ENDColumnIndex, $MaximumColumnWidth)
                    $WorkBookInstance.AutoFitColumn($StartColumnIndex, $ENDColumnIndex, $MaximumColumnWidth) | Out-Null
                }
                Else
                {
                    Write-Verbose ("Set-SLAutoFitColumn :`tSetting autofit on columns from '{0}' to '{1}'" -f $StartColumnIndex, $ENDColumnIndex)
                    $WorkBookInstance.AutoFitColumn($StartColumnIndex, $ENDColumnIndex) | Out-Null
                }

            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        } # select-slworksheet
    }#process

}
