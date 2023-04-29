Function Set-SLFont  {


    <#

.SYNOPSIS
    Set Font settings on a single or a range of cells.

.DESCRIPTION
    Set Font settings on a single or a range of cells.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER CellReference
    The target cell that needs to have the specified font settings. Eg: A5 or AB10

.PARAMETER Range
    The target cell range that needs to have the specified font settings. Eg: A5:B10 or AB10:AD20

.PARAMETER FontName
    Name of the font.

.PARAMETER FontSize
    Size of the font.

.PARAMETER FontColor
    Color of the font. Use tab completion or intellisense to select a possible value from a list provided by the parameter.

.PARAMETER Underline
    Specifies the underline formatting style of the font text.Valid values are:'Single','Double','SingleAccounting','DoubleAccounting','None'

.PARAMETER IsBold
    Specifies if the font text should be bold.

.PARAMETER IsItalic
    Specifies if the font text should be italic.

.PARAMETER IsStrikenthrough
    Specifies if the font text should have a strikethrough.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLFont -WorksheetName sheet1 -CellReference C15 -Underline Double -IsBold -IsStrikenThrough -Verbose | Save-SLDocument

    Description
    -----------
    Apply Underline,Bold & Strikethrough settings to cell C15


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx |
                Set-SLFont -WorksheetName sheet1 -Range g4:l5  -FontName "Segoe UI" -FontSize 13 -FontColor Chocolate -Verbose | Save-SLDocument

    Description
    -----------
    Apply font settings to a range of cells (g4:l5)


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLCellValue -WorksheetName sheet1 -CellReference B3 -value "Hello" -Verbose |
                Set-SLFont -Underline Double -IsBold -IsItalic -Verbose |
                    Save-SLDocument

    Description
    -----------
    Set the cell value of B3 to 'Hello' and then set the font settings. Notice how we did not have to specify the -worksheetname and -cellreference parameters
    for the 'Set-SLFont' function. This is because we already specified values for those parameters once for the 'Set-SLCellvalue' function so the output
    of this function becomes the input for Set-SLFont.


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
        [parameter(Mandatory = $true, Position = 1, ValueFromPipelineByPropertyName = $true)]
        [String]$WorksheetName,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Set-SLFont :`tCellReference should specify values in following format. Eg: A1,B10,AB5..etc"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true, ParameterSetname = 'cell')]
        [string[]]$CellReference,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+:[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Set-SLFont :`tRange should specify values in following format. Eg: A1:D10 or AB1:AD5"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true, ParameterSetname = 'Range')]
        [string]$Range,

        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, position = 3)]
        [string]$FontName,

        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, position = 4)]
        [System.UInt16]$FontSize,

        [Validateset('AliceBlue', 'AntiqueWhite', 'Aqua', 'Aquamarine', 'Azure', 'Beige', 'Bisque', 'Black', 'BlanchedAlmond', 'Blue', 'BlueViolet', 'Brown', 'BurlyWood', 'CadetBlue', 'Chartreuse', 'Chocolate', 'Coral', 'CornflowerBlue', 'Cornsilk',
            'Crimson', 'Cyan', 'DarkBlue', 'DarkCyan', 'DarkGoldenrod', 'DarkGray', 'DarkGreen', 'DarkKhaki', 'DarkMagenta', 'DarkOliveGreen', 'DarkOrange', 'DarkOrchid', 'DarkRed', 'DarkSalmon', 'DarkSeaGreen', 'DarkSlateBlue', 'DarkSlateGray',
            'DarkTurquoise', 'DarkViolet', 'DeepPink', 'DeepSkyBlue', 'DimGray', 'DodgerBlue', 'Firebrick', 'FloralWhite', 'ForestGreen', 'Fuchsia', 'Gainsboro', 'GhostWhite', 'Gold', 'Goldenrod', 'Gray', 'Green', 'GreenYellow', 'Honeydew', 'HotPink', 'IndianRed',
            'Indigo', 'Ivory', 'Khaki', 'LavENDer', 'LavENDerBlush', 'LawnGreen', 'LemonChiffon', 'LightBlue', 'LightCoral', 'LightCyan', 'LightGoldenrodYellow', 'LightGray', 'LightGreen', 'LightPink', 'LightSalmon', 'LightSeaGreen', 'LightSkyBlue', 'LightSlateGray',
            'LightSteelBlue', 'LightYellow', 'Lime', 'LimeGreen', 'Linen', 'Magenta', 'Maroon', 'MediumAquamarine', 'MediumBlue', 'MediumOrchid', 'MediumPurple', 'MediumSeaGreen', 'MediumSlateBlue', 'MediumSpringGreen', 'MediumTurquoise', 'MediumVioletRed', 'MidnightBlue'
            , 'MintCream', 'MistyRose', 'Moccasin', 'Name', 'NavajoWhite', 'Navy', 'OldLace', 'Olive', 'OliveDrab', 'Orange', 'OrangeRed', 'Orchid', 'PaleGoldenrod', 'PaleGreen', 'PaleTurquoise', 'PaleVioletRed', 'PapayaWhip', 'PeachPuff', 'Peru', 'Pink', 'Plum', 'PowderBlue',
            'Purple', 'Red', 'RosyBrown', 'RoyalBlue', 'SaddleBrown', 'Salmon', 'SandyBrown', 'SeaGreen', 'SeaShell', 'Sienna', 'Silver', 'SkyBlue', 'SlateBlue', 'SlateGray', 'Snow', 'SpringGreen', 'SteelBlue', 'Tan', 'Teal', 'Thistle', 'Tomato', 'Transparent', 'Turquoise',
            'Violet', 'Wheat', 'White', 'WhiteSmoke', 'Yellow', 'YellowGreen')]
        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, position = 5)]
        [String]$FontColor,

        [Validateset('Single', 'Double', 'SingleAccounting', 'DoubleAccounting', 'None')]
        [parameter(Mandatory = $false)]
        [String]$Underline,

        [switch]$IsBold,

        [switch]$IsItalic,

        [switch]$IsStrikenThrough


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

                    if ($isBold) { $SLStyle.Font.Bold = $true }
                    if ($isItalic) { $SLStyle.Font.Italic = $true }
                    if ($IsStrikenThrough) { $SLStyle.Font.Strike = $true }

                    if ($FontName) { $SLStyle.Font.FontName = $FontName }
                    if ($FontSize) { $SLStyle.Font.FontSize = $FontSize }
                    if ($FontColor) { $SLStyle.SetFontColor([System.Drawing.Color]::$FontColor) }
                    if ($Underline) { $SLStyle.Font.Underline = [DocumentFormat.OpenXml.Spreadsheet.UnderlineValues]::$Underline }

                    Write-Verbose ("Set-SLFont :`tSetting Font Style on Cell '{0}'" -f $cref)
                    $WorkBookInstance.SetCellStyle($Cref, $SLStyle) | Out-Null
                }
                $WorkBookInstance | Add-Member NoteProperty CellReference $CellReference -Force
            }
            elseif ($PSCmdlet.ParameterSetName -eq 'Range')
            {
                Write-Verbose ("Set-SLFont :`tSetting Font Style on Cell Range '{0}'" -f $Range)
                $rowindex, $columnindex = $range -split ':'

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
                    if ($isBold) { $SLStyle.Font.Bold = $true }
                    if ($isItalic) { $SLStyle.Font.Italic = $true }
                    if ($IsStrikenThrough) { $SLStyle.Font.Strike = $true }

                    if ($FontName) { $SLStyle.Font.FontName = $FontName }
                    if ($FontSize) { $SLStyle.Font.FontSize = $FontSize }
                    if ($FontColor) { $SLStyle.SetFontColor([System.Drawing.Color]::$FontColor) }
                    if ($Underline) { $SLStyle.Font.Underline = [DocumentFormat.OpenXml.Spreadsheet.UnderlineValues]::$Underline }
                    $CRCol = ([regex]::Match($cell, '[a-zA-Z]+') | Select-Object -ExpandProperty value) + $erow
                    $WorkBookInstance.SetCellStyle($Cell, $CrCol, $SLStyle) | Out-Null

                    $k++
                }

                $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
            }#if parameterset range

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }#select worksheet

    }#Process
    END
    {

    }

}
