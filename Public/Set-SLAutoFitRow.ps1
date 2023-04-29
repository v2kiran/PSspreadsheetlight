Function Set-SLAutoFitRow  {


    <#

.SYNOPSIS
    Autofit rows by RowIndex.

.DESCRIPTION
    Autofit columns by RowIndex.A single row or a range of rows can be specified as input.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER RowIndex
    The row to be autofit. Eg: 2 or 5.

.PARAMETER StartRowIndex
    Specifies the start of the autofit row range. Eg: 2.

.PARAMETER EndRowIndex
    Specifies the end of the autofit row range. Eg: 5.

.PARAMETER MaximumRowHeight
    Specifies the maximum row height for a row or a range of rows after autofit is applied. Eg: 10.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLAutoFitRow -WorksheetName sheet5 -RowIndex 3 -Verbose | Save-SLDocument

    Description
    -----------
    Autofit row 3.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLAutoFitRow -WorksheetName sheet5 -StartRowIndex 4 -ENDRowIndex 6 -Verbose | Save-SLDocument

    Description
    -----------
    Autofit rows 4 to 6 by index.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLAutoFitRow -WorksheetName sheet5 -StartRowIndex 4 -ENDRowIndex 6 -MaximumRowHeight 20 -Verbose | Save-SLDocument

    Description
    -----------
    Autofit rows 4 to 6 by index and optionally set a MaximumRowHeight of 20.

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

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Position = 2, ParameterSetName = 'SingleRowIndex')]
        [int]$RowIndex,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'MultiPleRowIndex')]
        [int]$StartRowIndex,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'MultiPleRowIndex')]
        [int]$ENDRowIndex,

        [parameter(Mandatory = $false)]
        [Double]$MaximumRowHeight

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'All')
            {
                Write-Verbose ("Set-SLAutoFitRow :`tSetting autofit on the first 2000 rows in worksheet '{0}'" -f $worksheetname)
                $WorkBookInstance.AutoFitRow(1, 2000) | Out-Null
            }

            elseif ($PSCmdlet.ParameterSetName -eq 'SingleRowIndex')
            {
                if ($MaximumRowHeight)
                {
                    Write-Verbose ("Set-SLAutoFitRow :`tSetting autofit on Row '{0}' with MaximumRowHeight of '{1}'" -f $RowIndex, $MaximumRowHeight)
                    $WorkBookInstance.AutoFitRow($RowIndex, $MaximumRowHeight) | Out-Null
                }
                Else
                {
                    Write-Verbose ("Set-SLAutoFitRow :`tSetting autofit on Row '{0}'" -f $RowIndex)
                    $WorkBookInstance.AutoFitRow($RowIndex)
                }
            }

            elseif ($PSCmdlet.ParameterSetName -eq 'MultiPleRowIndex')
            {
                if ($MaximumRowHeight)
                {
                    Write-Verbose ("Set-SLAutoFitRow :`tSetting autofit on Rows '{0}' to '{1}' with MaximumRowHeight of '{2}'" -f $StartRowIndex, $ENDRowIndex, $MaximumRowHeight)
                    $WorkBookInstance.AutoFitRow($StartRowIndex, $ENDRowIndex, $MaximumRowHeight) | Out-Null
                }
                Else
                {
                    Write-Verbose ("Set-SLAutoFitRow :`tSetting autofit on Rows '{0}' to '{1}'" -f $StartRowIndex, $ENDRowIndex)
                    $WorkBookInstance.AutoFitRow($StartRowIndex, $ENDRowIndex) | Out-Null
                }
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }#select-slworksheet
    }#process

}
