Function Set-SLColumnValue  {


    <#

.SYNOPSIS
    Set column values.

.DESCRIPTION
    Set column values..
    values cannot span multiple columns.Values are set on a single column moving from top to bottom until the value enumeration stops.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER CellReference
    The CellReference that specifies the start row and start column. Eg: A5 or AB10

.PARAMETER Value
    User can specify single or multiple values. Value assignment flow is from top to bottom.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLColumnValue -CellReference b3 -value Jan,feb,march -Verbose | Save-SLDocument

    VERBOSE: Set-SLColumnValue :	Setting value 'Jan' on cell 'b3'
    VERBOSE: Set-SLColumnValue :	Setting value 'feb' on cell 'b4'
    VERBOSE: Set-SLColumnValue :	Setting value 'march' on cell 'b5'

    Description
    -----------
    Since we specified 3 values(jan,feb & march) the cell values start from b3 and flows down to b5.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLColumnValue -CellReference b3 -value Jan,feb,march -Verbose |
                Set-SLBuiltinCellStyle -CellStyle Accent1 -Verbose |
                    Save-SLDocument

    Description
    -----------
    We build on the previous example by setting a cell style:'Accent1' on the cells whose values were set using 'Set-columnValue'.
    Note: Since we piped the output of Set-SLColumnValue we didnt have to specify a worksheetname or cell range with the 'Set-SLBuiltinCellStyle'
    cmdlet because those values are automatically mapped from the "SLdocument" object .

.INPUTS
   String,SpreadsheetLight.SLDocument

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    N/A
#>



    [CmdletBinding()]
    [OutputType([SpreadsheetLight.SLDocument])]
    param (
        [parameter(Mandatory = $true, Position = 0, ValueFromPipeLine = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [Alias('CurrentWorkSheetName')]
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Set-SLBuiltinCellStyle :`tCellReference should specify values in following format. Eg: A1,B10,AB5..etc"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true)]
        [string]$CellReference,

        [parameter(Mandatory = $true, Position = 3, ValueFromPipelineByPropertyName = $true)]
        $value
    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            $col = [regex]::Match($CellReference, '[a-zA-Z]') | Select-Object -ExpandProperty value
            [int]$Row = [regex]::Match($CellReference, '\d+') | Select-Object -ExpandProperty value

            $StartCellReference = $CellReference

            foreach ($val in $value)
            {
                $CellReference = $col + $row
                Write-Verbose ("Set-SLColumnValue :`tSetting value '{0}' on cell '{1}'" -f $val, $CellReference)

                $WorkBookInstance.SetCellValue($CellReference, $val) | Out-Null
                $row++
            }

            $Range = $StartCellReference + ':' + ($col + ($row - 1))

            $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }
    }#Process
    END
    {

    }

}
