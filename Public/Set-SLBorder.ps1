Function Set-SLBorder  {


    <#

.SYNOPSIS
    Set Border Style on a single or a range of cells.

.DESCRIPTION
    Set Border Style on a single or a range of cells.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER CellReference
    The target cell that needs to have the specified border settings. Eg: A5 or AB10

.PARAMETER Range
    The target cell range that needs to have the specified border settings. Eg: A5:B10 or AB10:AD20

.PARAMETER BorderStyle
    Valid values for the border style parameter are as follows:
    'Thick','Thin','Double','Dotted','Hair','Dashed','DashDot','DashDotDot','SlantDashDot','Medium','MediumDashDot','MediumDashDotDot','MediumDashed'.

.PARAMETER BorderColor
    Use tab completion or intellisense to select a possible value from a list provided by the parameter.

.PARAMETER LeftBorder
    Specify style settings that apply to the left border of a cell or cell range.

.PARAMETER RightBorder
    Specify style settings that apply to the right border of a cell or cell range.

.PARAMETER TopBorder
    Specify style settings that apply to the top border of a cell or cell range.

.PARAMETER BottomBorder
    Specify style settings that apply to the Bottom border of a cell or cell range.

.PARAMETER VerticalBorder
    Specify style settings that apply to the vertical border of a cell or cell range.

.PARAMETER HorizontalBorder
    Specify style settings that apply to the Horizontal border of a cell or cell range.

.PARAMETER CellBorder
    Specify style settings that apply to the whole cell i.e., Left,right,top,bottom.

.PARAMETER DiagonalBorder
    Specify style settings that apply to the DiagonalUp & DiagonalDown border of a cell or cell range.



.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLBorder -WorksheetName sheet1 -CellReference D3 -BorderStyle Double -BorderColor CadetBlue -CellBorder -Verbose | Save-SLDocument

    Description
    -----------
    Apply a border style to cell D3 using the switch '-Cellborder' which means apply the same style to all the borders of the cell D3.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx |
                    Set-SLBorder -WorksheetName sheet1 -Range d15:f24 -BorderStyle Double -BorderColor Blue -CellBorder |
                            Save-SLDocument

    Description
    -----------
    Apply border style to a range of cells (d15:f24)


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx |
                Set-SLBorder -WorksheetName sheet1 -Range d15:f24 -BorderStyle Double -BorderColor CadetBlue -LeftBorder  |
                    Set-SLBorder -BorderStyle Dashed -BorderColor Blue   -RightBorder |
                        Set-SLBorder -BorderStyle Dotted -BorderColor Orange -TopBorder   |
                            Set-SLBorder -BorderStyle Hair   -BorderColor Violet -BottomBorder   |
                                Save-SLDocument

    Description
    -----------
    Here we apply a different border style to each side of the cell range d15:f24.
    Notice that we had to specify the worksheetname and range parameters only once.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx |
                Set-SLBorder -WorksheetName sheet1 -Range d15:d24 -BorderStyle Double -BorderColor CadetBlue -LeftBorder  |
                    Set-SLBorder -Range e15:e24 -BorderStyle Dashed -BorderColor Blue   -RightBorder |
                        Set-SLBorder -Range f15:f24 -BorderStyle Dotted -BorderColor Orange -TopBorder   |
                            Set-SLBorder -Range g15:g24 -BorderStyle Hair   -BorderColor Violet -LeftBorder   |
                                 Save-SLDocument

    Description
    -----------
    Similar to the previous example except that in this case we specify a different border setting for different ranges.
    Notice that we had to specify the worksheetname parameter only once.


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
        [parameter(Mandatory = $true, Position = 1, ValueFromPipeLine = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [Alias('CurrentWorkSheetName')]
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Set-SLBorder :`tCellReference should specify values in following format. Eg: A1,B10,AB5..etc"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true, ParameterSetname = 'cell')]
        [string[]]$CellReference,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+:[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Set-SLBorder :`tRange should specify values in following format. Eg: A1:D10 or AB1:AD5"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true, ParameterSetname = 'Range')]
        [string]$Range,


        [Validateset('Thick', 'Thin', 'Double', 'Dotted', 'Hair', 'Dashed', 'DashDot', 'DashDotDot', 'SlantDashDot', 'Medium', 'MediumDashDot', 'MediumDashDotDot', 'MediumDashed')]
        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [String]$BorderStyle,


        [Validateset('AliceBlue', 'AntiqueWhite', 'Aqua', 'Aquamarine', 'Azure', 'Beige', 'Bisque', 'Black', 'BlanchedAlmond', 'Blue', 'BlueViolet', 'Brown', 'BurlyWood', 'CadetBlue', 'Chartreuse', 'Chocolate', 'Coral', 'CornflowerBlue', 'Cornsilk',
            'Crimson', 'Cyan', 'DarkBlue', 'DarkCyan', 'DarkGoldenrod', 'DarkGray', 'DarkGreen', 'DarkKhaki', 'DarkMagenta', 'DarkOliveGreen', 'DarkOrange', 'DarkOrchid', 'DarkRed', 'DarkSalmon', 'DarkSeaGreen', 'DarkSlateBlue', 'DarkSlateGray',
            'DarkTurquoise', 'DarkViolet', 'DeepPink', 'DeepSkyBlue', 'DimGray', 'DodgerBlue', 'Firebrick', 'FloralWhite', 'ForestGreen', 'Fuchsia', 'Gainsboro', 'GhostWhite', 'Gold', 'Goldenrod', 'Gray', 'Green', 'GreenYellow', 'Honeydew', 'HotPink', 'IndianRed',
            'Indigo', 'Ivory', 'Khaki', 'LavENDer', 'LavENDerBlush', 'LawnGreen', 'LemonChiffon', 'LightBlue', 'LightCoral', 'LightCyan', 'LightGoldenrodYellow', 'LightGray', 'LightGreen', 'LightPink', 'LightSalmon', 'LightSeaGreen', 'LightSkyBlue', 'LightSlateGray',
            'LightSteelBlue', 'LightYellow', 'Lime', 'LimeGreen', 'Linen', 'Magenta', 'Maroon', 'MediumAquamarine', 'MediumBlue', 'MediumOrchid', 'MediumPurple', 'MediumSeaGreen', 'MediumSlateBlue', 'MediumSpringGreen', 'MediumTurquoise', 'MediumVioletRed', 'MidnightBlue'
            , 'MintCream', 'MistyRose', 'Moccasin', 'Name', 'NavajoWhite', 'Navy', 'OldLace', 'Olive', 'OliveDrab', 'Orange', 'OrangeRed', 'Orchid', 'PaleGoldenrod', 'PaleGreen', 'PaleTurquoise', 'PaleVioletRed', 'PapayaWhip', 'PeachPuff', 'Peru', 'Pink', 'Plum', 'PowderBlue',
            'Purple', 'Red', 'RosyBrown', 'RoyalBlue', 'SaddleBrown', 'Salmon', 'SandyBrown', 'SeaGreen', 'SeaShell', 'Sienna', 'Silver', 'SkyBlue', 'SlateBlue', 'SlateGray', 'Snow', 'SpringGreen', 'SteelBlue', 'Tan', 'Teal', 'Thistle', 'Tomato', 'Transparent', 'Turquoise',
            'Violet', 'Wheat', 'White', 'WhiteSmoke', 'Yellow', 'YellowGreen')]
        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [String]$BorderColor,

        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [Switch]$LeftBorder,

        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [Switch]$RightBorder,

        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [Switch]$TopBorder,

        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [Switch]$BottomBorder,

        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [Switch]$VerticalBorder,

        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [Switch]$HoriZontalBorder,

        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [Switch]$CellBorder,

        [parameter(ValueFromPipelineByPropertyName = $true)]
        [switch]$DiagonalBorder





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

                    $BStyle = [DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues]::$BorderStyle
                    $BColor = [System.Drawing.Color]::$Bordercolor

                    if ($LeftBorder) { $SLStyle.SetLeftBorder($BStyle, $BColor) }
                    if ($RightBorder) { $SLStyle.SetRightBorder($BStyle, $BColor) }
                    if ($TopBorder) { $SLStyle.SetTopBorder($BStyle, $BColor) }
                    if ($BottomBorder) { $SLStyle.SetBottomBorder($BStyle, $BColor) }
                    if ($VerticalBorder) { $SLStyle.SetVerticalBorder($BStyle, $BColor) }
                    if ($HoriZontalBorder) { $SLStyle.SetHorizontalBorder($BStyle, $BColor) }

                    if ($DiagonalBorder)
                    {
                        $SLStyle.Border.DiagonalUp = $true
                        $SLStyle.Border.DiagonalDown = $true
                        $SLStyle.SetDiagonalBorder($BStyle, $BColor) | Out-Null
                    }

                    if ($CellBorder)
                    {
                        $SLStyle.SetLeftBorder($BStyle, $BColor)
                        $SLStyle.SetRightBorder($BStyle, $BColor)
                        $SLStyle.SetTopBorder($BStyle, $BColor)
                        $SLStyle.SetBottomBorder($BStyle, $BColor)
                    }

                    Write-Verbose ("Set-SLBorder :`tSetting Border Style on Cell '{0}'" -f $cref)
                    $WorkBookInstance.SetCellStyle($Cref, $SLStyle) | Out-Null
                }
                $WorkBookInstance | Add-Member NoteProperty CellReference $CellReference -Force
            }

            elseif ($PSCmdlet.ParameterSetName -eq 'Range')
            {

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

                    $BStyle = [DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues]::$BorderStyle
                    $BColor = [System.Drawing.Color]::$Bordercolor

                    if ($DiagonalBorder)
                    {
                        $SLStyle.Border.DiagonalUp = $true
                        $SLStyle.Border.DiagonalDown = $true
                        $SLStyle.SetDiagonalBorder($BStyle, $BColor) | Out-Null
                    }

                    if ($LeftBorder) { $SLStyle.SetLeftBorder($BStyle, $BColor) }
                    if ($RightBorder) { $SLStyle.SetRightBorder($BStyle, $BColor) }
                    if ($TopBorder) { $SLStyle.SetTopBorder($BStyle, $BColor) }
                    if ($BottomBorder) { $SLStyle.SetBottomBorder($BStyle, $BColor) }
                    if ($VerticalBorder) { $SLStyle.SetVerticalBorder($BStyle, $BColor) }
                    if ($HoriZontalBorder) { $SLStyle.SetHorizontalBorder($BStyle, $BColor) }

                    if ($CellBorder)
                    {
                        $SLStyle.SetLeftBorder($BStyle, $BColor)
                        $SLStyle.SetRightBorder($BStyle, $BColor)
                        $SLStyle.SetTopBorder($BStyle, $BColor)
                        $SLStyle.SetBottomBorder($BStyle, $BColor)
                    }

                    $CRCol = ([regex]::Match($cell, '[a-zA-Z]+') | Select-Object -ExpandProperty value) + $erow

                    Write-Verbose ("Set-SLBorder :`tSetting Border Style on Cell Range '{0}'" -f $Range)
                    $WorkBookInstance.SetCellStyle($Cell, $CrCol, $SLStyle) | Out-Null

                    $k++
                }

                $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
            }#if parameterset range

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }#If Select-SLWorksheet

    }#Process
    END
    {

    }

}
