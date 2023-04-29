Function Set-SLAlignMent  {


    <#

.SYNOPSIS
    Set text alignment settings on a single or a range of cells.

.DESCRIPTION
    Set text alignment settings on a single or a range of cells.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER CellReference
    The target cell that needs to have the specified alignment settings. Eg: A5 or AB10

.PARAMETER Range
    The target cell range that needs to have the specified alignment settings. Eg: A5:B10 or AB10:AD20

.PARAMETER Vertical
    Valid values for the Vertical alignment parameter is - 'Bottom','Center','Top','Justify','Distributed'.

.PARAMETER Horizontal
    Valid values for the Horizontal alignment parameter is - 'Left','Center','Right','Justify','Distributed'.

.PARAMETER TextRotation
    Specifies the rotation angle of the text, ranging from -90 degrees to 90 degrees.

.PARAMETER Indent
    Each indent value is 3 spaces so an indent value of 5 means 15 spaces wide.

.PARAMETER ShrinkToFit
    Specifies if the text in the cell should be shrunk to fit the cell.

.PARAMETER WrapText
    Specifies if the text in the cell should be wrapped.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLAlignMent -WorksheetName a -cellreference b3 -WrapText -Verbose | Save-SLDocument

    Description
    -----------
    Apply textwrap to cell B3


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLAlignMent -WorksheetName a -cellreference b3 -Vertical Top -WrapText  | Save-SLDocument

    Description
    -----------
    Top align cell content in B3 and then wrap text.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLAlignMent -WorksheetName a -cellreference b3 -indent 3 -TextRotation -80 | Save-SLDocument

    Description
    -----------
    Indent text by 9 spaces and set rotation at 80 degrees.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLAlignMent -WorksheetName a -Range D5:d16 -Horizontal Left -Vertical Center -indent 3 | Save-SLDocument

    Description
    -----------
    Here we apply multiple alignment settings settings to a range of cells.


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
                else { $false; Write-Warning 'CellReference should specify values in following format. Eg: A1,B10,AB5..etc'; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ParameterSetname = 'cell', ValueFromPipeLineByPropertyName = $true)]
        [string[]]$CellReference,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+:[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning 'Range should specify values in following format. Eg: A1:D10 or AB1:AD5'; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true, ParameterSetname = 'Range')]
        [string]$Range,

        [Validateset('Bottom', 'Center', 'Top', 'Justify', 'Distributed')]
        [parameter(Mandatory = $false, ValueFromPipeLineByPropertyName = $true)]
        [String]$Vertical,

        [Validateset('Left', 'Center', 'Right', 'Justify', 'Distributed')]
        [parameter(Mandatory = $false, ValueFromPipeLineByPropertyName = $true)]
        [String]$Horizontal,

        [Validaterange(-90, 90)]
        [parameter(Mandatory = $false, ValueFromPipeLineByPropertyName = $true)]
        [int]$TextRotation,

        [parameter(Mandatory = $false, ValueFromPipeLineByPropertyName = $true)]
        [int]$indent,

        [parameter(Mandatory = $false, ValueFromPipeLineByPropertyName = $true)]
        [switch]$ShrinkToFit,

        [parameter(Mandatory = $false, ValueFromPipeLineByPropertyName = $true)]
        [switch]$WrapText
    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'cell')
            {
                Foreach ($cref in $CellReference)
                {
                    $SLStyle = $WorkBookInstance.GetCellStyle($cref)

                    ## each indent is 3 spaces
                    $SLStyle.Alignment.Indent = $indent

                    if ($ShrinkToFit) { $SLStyle.Alignment.ShrinkToFit = $true }
                    if ($WrapText) { $SLStyle.Alignment.WrapText = $true }
                    if ($TextRotation) { $SLStyle.Alignment.TextRotation = $TextRotation }
                    if ($Vertical) { $SLStyle.Alignment.Vertical = [DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues]::$Vertical }
                    if ($Horizontal) { $SLStyle.Alignment.Horizontal = [DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues]::$Horizontal }

                    Write-Verbose ("Set-SLAlignMent :`tSetting Alignment options on cell '{0}'..." -f $cref)
                    $WorkBookInstance.SetCellStyle($Cref, $SLStyle) | Out-Null
                }
                $WorkBookInstance | Add-Member NoteProperty CellReference $CellReference -Force
            }

            elseif ($PSCmdlet.ParameterSetName -eq 'Range')
            {
                $rowindex, $columnindex = $range -split ':'
                Write-Verbose ("Set-SLAlignMent :`tSetting Alignment options on CellRange '{0}'..." -f $Range)

                $startrowcolumn = Convert-ToExcelRowColumnIndex -CellReference $rowindex
                $endrowcolumn = Convert-ToExcelRowColumnIndex -CellReference $columnindex
                $sRow = $startrowcolumn.Row
                $sColumn = $startrowcolumn.Column
                $eRow = $endrowcolumn.Row
                $eColumn = $endrowcolumn.Column

                $k = 0
                for ($i = $sColumn; $i -le $eColumn; $i++)
                {
                    $Cell = (Convert-ToExcelColumnName -index ($startrowcolumn.Column + $k)) + $sRow

                    $SLStyle = $WorkBookInstance.GetcellStyle($Cell)
                    ## each indent is 3 spaces
                    $SLStyle.Alignment.Indent = $indent

                    if ($ShrinkToFit) { $SLStyle.Alignment.ShrinkToFit = $true }
                    if ($WrapText) { $SLStyle.Alignment.WrapText = $true }
                    if ($TextRotation) { $SLStyle.Alignment.TextRotation = $TextRotation }
                    if ($Vertical) { $SLStyle.Alignment.Vertical = [DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues]::$Vertical }
                    if ($Horizontal) { $SLStyle.Alignment.Horizontal = [DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues]::$Horizontal }
                    $CRCol = ([regex]::Match($cell, '[a-zA-Z]+') | Select-Object -ExpandProperty value) + $erow
                    $WorkBookInstance.SetCellStyle($Cell, $CrCol, $SLStyle) | Out-Null

                    $k++
                }

                $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }#Select-slworksheet

    }#Process
    END
    {

    }

}
