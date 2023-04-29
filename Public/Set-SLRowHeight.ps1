Function Set-SLRowHeight  {


    <#

.SYNOPSIS
    Set Row height by index.

.DESCRIPTION
    Set Row height by index.A single row or a range of rows can be specified as input.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER RowIndex
    The row index. Eg: 2 or 5.

.PARAMETER StartRowIndex
    Specifies the start index of the row range. Eg: 1 .

.PARAMETER ENDRowIndex
    Specifies the end index of the row range. Eg: 7.

.PARAMETER RowHeight
    Specifies the row height to be applied. Eg: 20.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLRowHeight -WorksheetName sheet5 -RowIndex 3 -RowHeight 30 -Verbose | Save-SLDocument

    Description
    -----------
    Set Rowheight of row 3 to 30.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLRowHeight -WorksheetName sheet5 -StartRowIndex 4 -ENDRowIndex 6 -RowHeight 15 -Verbose | Save-SLDocument

    Description
    -----------
    Set Rowheight of a range 4 - 6 to 15.


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

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Position = 2, ParameterSetName = 'SingleRowIndex')]
        [int]$RowIndex,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'MultiPleRowIndex')]
        [int]$StartRowIndex,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'MultiPleRowIndex')]
        [int]$ENDRowIndex,

        [parameter(Mandatory = $true)]
        [Double]$RowHeight

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'SingleRowIndex')
            {
                Write-Verbose ("Set-SLRowHeight :`tSetting height of row '{0}' to '{1}'" -f $RowIndex, $RowHeight)
                $WorkBookInstance.SetRowHeight($RowIndex, $RowHeight) | Out-Null
            }

            elseif ($PSCmdlet.ParameterSetName -eq 'MultiPleRowIndex')
            {
                Write-Verbose ("Set-SLRowHeight :`tSetting height of rows '{0}' to '{1}' to '{2}'" -f $StartRowIndex, $ENDRowIndex, $RowHeight)
                $WorkBookInstance.SetRowHeight($StartRowIndex, $ENDRowIndex, $RowHeight) | Out-Null
            }
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }#select-slworksheet
    }#process

}
