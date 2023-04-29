Function Insert-SLPageBreak  {


    <#

.SYNOPSIS
    Insert pagebreaks.

.DESCRIPTION
    Insert pagebreaks.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER AboveRowIndex
    Row index.

.PARAMETER LeftofColumnIndex
    Column Index.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Insert-SLPageBreak -WorksheetName sheet2 -AboveRowIndex 9 -Verbose  | Save-SLDocument


    Description
    -----------
    Insert page break above row 9.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Insert-SLPageBreak -WorksheetName sheet2 -LeftofColumnIndex 5 -Verbose  | Save-SLDocument


    Description
    -----------
    Insert page break to the left of column 5.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Insert-SLPageBreak -WorksheetName sheet2 -AboveRowIndex 6 -LeftofColumnIndex 6 -Verbose  | Save-SLDocument


    Description
    -----------
    Insert page break above row 6 and to the left of column 6.

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
        [Int]$LeftofColumnIndex

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {

            if ($PSCmdlet.ParameterSetName -eq 'RowColumn')
            {
                Write-Verbose ("Insert-SLPageBreak :`tInsert pagebreak above Row '{0}' and to the left of column '{1}'" -f $AboveRowIndex, $LeftofColumnIndex)
                $WorkBookInstance.InsertPageBreak($AboveRowIndex, $LeftofColumnIndex) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Row')
            {
                Write-Verbose ("Insert-SLPageBreak :`tInsert pagebreak above Row '{0}'" -f $AboveRowIndex)
                $WorkBookInstance.InsertPageBreak($AboveRowIndex, -1) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Column')
            {
                Write-Verbose ("Insert-SLPageBreak :`tInsert pagebreak to the left of column '{0}'" -f $LeftofColumnIndex)
                $WorkBookInstance.InsertPageBreak(-1, $LeftofColumnIndex) | Out-Null
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }

    }#process
    END
    {
    }

}
