Function Remove-SLPageBreak  {


    <#

.SYNOPSIS
    Remove pagebreaks.

.DESCRIPTION
    Remove pagebreaks.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER AboveRowIndex
    Row index.

.PARAMETER LeftofColumnIndex
    Column Index.

.PARAMETER All
    If specified Will remove all page breaks from a worksheet .

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Remove-SLPageBreak -WorksheetName sheet2 -AboveRowIndex 9 -Verbose  | Save-SLDocument


    Description
    -----------
    remove page break above row 9.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Remove-SLPageBreak -WorksheetName sheet2 -LeftofColumnIndex 5 -Verbose  | Save-SLDocument


    Description
    -----------
    remove page break to the left of column 5.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Remove-SLPageBreak -WorksheetName sheet2 -AboveRowIndex 6 -LeftofColumnIndex 6 -Verbose  | Save-SLDocument


    Description
    -----------
    remove page break above row 6 and to the left of column 6.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Remove-SLPageBreak -WorksheetName sheet2 -All  -Verbose  | Save-SLDocument


    Description
    -----------
    remove all page breaks in worksheet 'sheet2'.

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

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetname = 'Row')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetname = 'RowColumn')]
        [Int]$AboveRowIndex,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetname = 'Column')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetname = 'RowColumn')]
        [Int]$LeftofColumnIndex,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetname = 'all')]
        [Switch]$All

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'RowColumn')
            {
                Write-Verbose ("Remove-SLPageBreak :`tRemoving pagebreak above Row '{0}' and to the left of column '{1}'" -f $AboveRowIndex, $LeftofColumnIndex)
                $WorkBookInstance.RemovePageBreak($AboveRowIndex, $LeftofColumnIndex) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Row')
            {
                Write-Verbose ("Remove-SLPageBreak :`tRemoving pagebreak above Row '{0}'" -f $AboveRowIndex)
                $WorkBookInstance.RemovePageBreak($AboveRowIndex, -1) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Column')
            {
                Write-Verbose ("Remove-SLPageBreak :`tRemoving pagebreak to the left of column '{0}'" -f $LeftofColumnIndex)
                $WorkBookInstance.RemovePageBreak(-1, $LeftofColumnIndex) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'all')
            {
                Write-Verbose ("Remove-SLPageBreak :`tRemoving all pagebreaks in the worksheet '{0}'" -f $WorksheetName)
                $WorkBookInstance.RemoveAllPageBreaks() | Out-Null
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }#select-slworksheet

    }#process
    END
    {
    }

}
