Function Set-SLRowValue  {


    <#

.SYNOPSIS
    Set Row values.

.DESCRIPTION
    Set Row values..
    values cannot span multiple rows.Values are set on a single row moving from left to right until the value enumeration stops.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER CellReference
    The CellReference that specifies the start row and start column. Eg: A5 or AB10

.PARAMETER Value
    User can specify single or multiple values. Value assignment flow is from top to bottom.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLRowValue -CellReference b3 -value Jan,feb,march -Verbose | Save-SLDocument

    VERBOSE: Set-SLRowValue :	Setting value 'Jan' on cell 'b3'
    VERBOSE: Set-SLRowValue :	Setting value 'feb' on cell 'c3'
    VERBOSE: Set-SLRowValue :	Setting value 'march' on cell 'd3'

    Description
    -----------
    Since we specified 3 values(jan,feb & march) the cell values start from b3 and flow right as shown above.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLRowValue    -CellReference b3 -value FirstName,LastName,DepartMent -Verbose |
                Set-SLRowValue -CellReference b4 -value Jon,Doe,Sales -Verbose |
                Set-SLRowValue -CellReference b5 -value Zenedine,Zidanne,Football -Verbose |
                Set-SLRowValue -CellReference b6 -value Rahul,Dravid,Cricket -Verbose |
                    Save-SLDocument

    Description
    -----------
    Create a table with 3 rows and 3 columns.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
            $doc | Set-SLRowValue -CellReference b3 -value FirstName,LastName,DepartMent  | Set-SLBuiltinCellStyle -CellStyle Heading3
            $doc | Set-SLRowValue -CellReference b4 -value Jon,Doe,Sales  | Set-SLAlignMent -Vertical Top  | Set-SLBuiltinCellStyle -CellStyle ExplanatoryText
            $doc | Set-SLRowValue -CellReference b5 -value Zenedine,Zidanne,Football  | Set-SLAlignMent -Vertical Top  | Set-SLBuiltinCellStyle -CellStyle ExplanatoryText
            $doc | Set-SLRowValue -CellReference b6 -value Rahul,Dravid,Cricket  | Set-SLAlignMent -Vertical Top | Set-SLBuiltinCellStyle -CellStyle ExplanatoryText
            $doc | Save-SLDocument

    Description
    -----------
    We build on the previous example by applying some alignment and cellstyles to our table.
    Note the above result can be achieved using a gaint piepline but for the sake of legibility the task has been split into various steps.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLRowValue    -CellReference b3 -value FirstName,LastName,DepartMent -Verbose |
                Set-SLRowValue -CellReference b4 -value Jon,Doe,Sales -Verbose |
                Set-SLRowValue -CellReference b5 -value Zenedine,Zidanne,Football -Verbose |
                Set-SLRowValue -CellReference b6 -value Rahul,Dravid,Cricket -Verbose |
                     Set-SLTableStyle -Range B3:D6 -TableStyle Medium17 |
                        Save-SLDocument

    Description
    -----------
    Instead of styling individual rows and columns we can set a tablestyle by specifying the range.


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
                else { $false; Write-Warning "Set-SLRowValue :`tCellReference should specify values in following format. Eg: A1,B10,AB5..etc"; break }
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
            $colIndex = Convert-ToExcelColumnIndex -ColumnName $col

            foreach ($val in $value)
            {
                $CellReference = (Convert-ToExcelColumnName $colIndex) + $Row
                Write-Verbose ("Set-SLRowValue :`tSetting value '{0}' on cell '{1}'" -f $val, $CellReference)

                $WorkBookInstance.SetCellValue($CellReference, $val) | Out-Null
                $colIndex++
            }

            $Range = $StartCellReference + ':' + ((Convert-ToExcelColumnName -Index ($colIndex - 1)) + $row)

            $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }
    }#Process
    END
    {

    }

}
