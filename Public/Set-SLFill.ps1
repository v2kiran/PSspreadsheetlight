Function Set-SLFill  {


    <#

.SYNOPSIS
    Set Fill settings on a single or a range of cells.

.DESCRIPTION
    Set Fill settings on a single or a range of cells.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER TargetCellorRange
    The target cellreference or Range that needs to have the specified fill settings. Eg: A5 or A5:B10
    Due to the complexity involved in setting up the various fill methods the cellreference and range parameters have been combined as TargetcellorRage.

.PARAMETER Color
    The fill color to be set.

.PARAMETER ColorFromHTML
    The fill color from an HTML string such as '#12b1e6'.

.PARAMETER ThemeColor
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'Accent1Color','Accent2Color','Accent3Color','Accent4Color','Accent5Color',
    'Accent6Color','Dark1Color','Dark2Color','Light1Color',
    'Light2Color','Hyperlink','FollowedHyperlinkColor'

.PARAMETER Pattern
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'DarkDown','DarkGray','DarkGrid','DarkHorizontal','DarkTrellis',
    'DarkUp','DarkVertical','Gray0625','Gray125',
    'LightDown','LightGray','LightGrid','LightHorizontal',
    'LightTrellis','LightUp','LightVertical','MediumGray','None','Solid'

.PARAMETER ForeGroundColor
    The ForeGroundColor fill color to be set. Values are the same as the parameter 'color'.

.PARAMETER BackGroundColor
    The BackGroundColor fill color to be set. Values are the same as the parameter 'color'.

.PARAMETER ForeGroundThemeColor
    The ForeGroundThemeColor fill color to be set. Values are the same as the parameter 'Themecolor'.

.PARAMETER BackGroundThemeColor
    The BackGroundThemeColor fill color to be set. Values are the same as the parameter 'Themecolor'.

.PARAMETER GradientDirection
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'Corner1','Corner2','Corner3','Corner4','DiagonalDown1',
    'DiagonalDown2','DiagonalDown3','DiagonalUp1','DiagonalUp2',
    'DiagonalUp3','Horizontal1','Horizontal2','Horizontal3',
    'Vertical1','Vertical2','Vertical3','FromCenter'


.Example
    PS C:\> Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx | Set-SLFill -WorksheetName sheet6 -TargetCellorRange b2 -Color Aqua -Verbose | Save-SLDocument

    Description
    -----------
    Apply fill color Aqua to cell B2


.Example
    PS C:\> Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx | Set-SLFill -WorksheetName sheet6 -TargetCellorRange b3 -ThemeColor Accent2Color -Verbose | Save-SLDocument

    Description
    -----------
    Apply fill themecolor Accent2Color to cell B3.


.Example
    PS C:\> Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx |
                Set-SLFill -WorksheetName sheet6 -TargetCellorRange b4 -Pattern DarkDown -ForeGroundColor Aquamarine -BackGroundColor AliceBlue -Verbose |
                    Save-SLDocument

    Description
    -----------
    Apply pattern darkdown with two different Foreground and background colors to cell B4.

.Example
    PS C:\> Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx |
                Set-SLFill -WorksheetName sheet6 -TargetCellorRange b5 -Pattern DarkGray -ForeGroundThemeColor Accent1Color -BackGroundThemeColor Accent2Color -Verbose |
                    Save-SLDocument

    Description
    -----------
    Apply pattern darkgray with two different Foreground and background Themecolors to cell B5.

.Example
    PS C:\> Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx |
                Set-SLFill -WorksheetName sheet6 -TargetCellorRange b6 -Pattern DarkGrid -ForeGroundThemeColor Accent2Color -BackGroundColor Brown -Verbose |
                    Save-SLDocument

    Description
    -----------
    Apply pattern darkgrid with a themecolor and a regular color value to cell B6.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Copy-SLCellValue -WorksheetName sheet6 -Range g2:i8 -ToAnchorCellreference g10 -PasteSpecial Values -Verbose
    PS C:\> $doc | Set-SLFont -WorksheetName sheet6 -Range g10:i10 -FontName Tahoma -FontColor White -IsBold -Verbose |
                Set-SLAlignMent -Vertical Center -Horizontal Center |
                     Set-SLFill -ColorFromHTML '#12b1e6'  -Verbose
    PS C:\> $doc | Set-SLFont -WorksheetName sheet6 -Range g11:g16 -FontName Tahoma -FontColor Tan  -Verbose | Set-SLFill -Color Gray
    PS C:\> $doc | Set-SLFill -WorksheetName sheet6 -TargetCellorRange H11:I16  -Color LightGray  -Verbose
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    Copy the range g2:g8 and paste it at G10 filling up cells G10:I16.
    Set font and alignment settings on the header range G10:I16 and apply a fill color '#12b1e6'
    Set a different font and fill color on the first data column. Font Tahoma & color Tan
    To provide contrast apply a light background fill on the remaining data columns H11:I16
    Dont forget to save the document :).

.INPUTS
   String,SpreadsheetLight.SLDocument

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

        [Alias('CellReference', 'Range')]
        [parameter(Mandatory = $true, Position = 1, ValueFromPipelineByPropertyName = $true)]
        [string]$TargetCellorRange,

        [Validateset('AliceBlue', 'AntiqueWhite', 'Aqua', 'Aquamarine', 'Azure', 'Beige', 'Bisque', 'Black', 'BlanchedAlmond', 'Blue', 'BlueViolet', 'Brown', 'BurlyWood', 'CadetBlue', 'Chartreuse', 'Chocolate', 'Coral', 'CornflowerBlue', 'Cornsilk',
            'Crimson', 'Cyan', 'DarkBlue', 'DarkCyan', 'DarkGoldenrod', 'DarkGray', 'DarkGreen', 'DarkKhaki', 'DarkMagenta', 'DarkOliveGreen', 'DarkOrange', 'DarkOrchid', 'DarkRed', 'DarkSalmon', 'DarkSeaGreen', 'DarkSlateBlue', 'DarkSlateGray',
            'DarkTurquoise', 'DarkViolet', 'DeepPink', 'DeepSkyBlue', 'DimGray', 'DodgerBlue', 'Firebrick', 'FloralWhite', 'ForestGreen', 'Fuchsia', 'Gainsboro', 'GhostWhite', 'Gold', 'Goldenrod', 'Gray', 'Green', 'GreenYellow', 'Honeydew', 'HotPink', 'IndianRed',
            'Indigo', 'Ivory', 'Khaki', 'LavENDer', 'LavENDerBlush', 'LawnGreen', 'LemonChiffon', 'LightBlue', 'LightCoral', 'LightCyan', 'LightGoldenrodYellow', 'LightGray', 'LightGreen', 'LightPink', 'LightSalmon', 'LightSeaGreen', 'LightSkyBlue', 'LightSlateGray',
            'LightSteelBlue', 'LightYellow', 'Lime', 'LimeGreen', 'Linen', 'Magenta', 'Maroon', 'MediumAquamarine', 'MediumBlue', 'MediumOrchid', 'MediumPurple', 'MediumSeaGreen', 'MediumSlateBlue', 'MediumSpringGreen', 'MediumTurquoise', 'MediumVioletRed', 'MidnightBlue'
            , 'MintCream', 'MistyRose', 'Moccasin', 'Name', 'NavajoWhite', 'Navy', 'OldLace', 'Olive', 'OliveDrab', 'Orange', 'OrangeRed', 'Orchid', 'PaleGoldenrod', 'PaleGreen', 'PaleTurquoise', 'PaleVioletRed', 'PapayaWhip', 'PeachPuff', 'Peru', 'Pink', 'Plum', 'PowderBlue',
            'Purple', 'Red', 'RosyBrown', 'RoyalBlue', 'SaddleBrown', 'Salmon', 'SandyBrown', 'SeaGreen', 'SeaShell', 'Sienna', 'Silver', 'SkyBlue', 'SlateBlue', 'SlateGray', 'Snow', 'SpringGreen', 'SteelBlue', 'Tan', 'Teal', 'Thistle', 'Tomato', 'Transparent', 'Turquoise',
            'Violet', 'Wheat', 'White', 'WhiteSmoke', 'Yellow', 'YellowGreen')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'Pattern1Color')]
        [string]$Color,

        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'Pattern1ColorHtml')]
        [string]$ColorFromHTML,

        [Validateset('Accent1Color', 'Accent2Color', 'Accent3Color', 'Accent4Color', 'Accent5Color',
            'Accent6Color', 'Dark1Color', 'Dark2Color', 'Light1Color',
            'Light2Color', 'Hyperlink', 'FollowedHyperlinkColor')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'Pattern1Theme')]
        [string]$ThemeColor,

        [Validateset('DarkDown', 'DarkGray', 'DarkGrid', 'DarkHorizontal', 'DarkTrellis',
            'DarkUp', 'DarkVertical', 'Gray0625', 'Gray125',
            'LightDown', 'LightGray', 'LightGrid', 'LightHorizontal',
            'LightTrellis', 'LightUp', 'LightVertical', 'MediumGray', 'None', 'Solid')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'Pattern1Theme1Color')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'Pattern2Colors')]
        [parameter(Mandatory = $false, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'Pattern1Color1Theme')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'Pattern2ThemeColors')]
        [string]$Pattern,

        [Validateset('AliceBlue', 'AntiqueWhite', 'Aqua', 'Aquamarine', 'Azure', 'Beige', 'Bisque', 'Black', 'BlanchedAlmond', 'Blue', 'BlueViolet', 'Brown', 'BurlyWood', 'CadetBlue', 'Chartreuse', 'Chocolate', 'Coral', 'CornflowerBlue', 'Cornsilk',
            'Crimson', 'Cyan', 'DarkBlue', 'DarkCyan', 'DarkGoldenrod', 'DarkGray', 'DarkGreen', 'DarkKhaki', 'DarkMagenta', 'DarkOliveGreen', 'DarkOrange', 'DarkOrchid', 'DarkRed', 'DarkSalmon', 'DarkSeaGreen', 'DarkSlateBlue', 'DarkSlateGray',
            'DarkTurquoise', 'DarkViolet', 'DeepPink', 'DeepSkyBlue', 'DimGray', 'DodgerBlue', 'Firebrick', 'FloralWhite', 'ForestGreen', 'Fuchsia', 'Gainsboro', 'GhostWhite', 'Gold', 'Goldenrod', 'Gray', 'Green', 'GreenYellow', 'Honeydew', 'HotPink', 'IndianRed',
            'Indigo', 'Ivory', 'Khaki', 'LavENDer', 'LavENDerBlush', 'LawnGreen', 'LemonChiffon', 'LightBlue', 'LightCoral', 'LightCyan', 'LightGoldenrodYellow', 'LightGray', 'LightGreen', 'LightPink', 'LightSalmon', 'LightSeaGreen', 'LightSkyBlue', 'LightSlateGray',
            'LightSteelBlue', 'LightYellow', 'Lime', 'LimeGreen', 'Linen', 'Magenta', 'Maroon', 'MediumAquamarine', 'MediumBlue', 'MediumOrchid', 'MediumPurple', 'MediumSeaGreen', 'MediumSlateBlue', 'MediumSpringGreen', 'MediumTurquoise', 'MediumVioletRed', 'MidnightBlue'
            , 'MintCream', 'MistyRose', 'Moccasin', 'Name', 'NavajoWhite', 'Navy', 'OldLace', 'Olive', 'OliveDrab', 'Orange', 'OrangeRed', 'Orchid', 'PaleGoldenrod', 'PaleGreen', 'PaleTurquoise', 'PaleVioletRed', 'PapayaWhip', 'PeachPuff', 'Peru', 'Pink', 'Plum', 'PowderBlue',
            'Purple', 'Red', 'RosyBrown', 'RoyalBlue', 'SaddleBrown', 'Salmon', 'SandyBrown', 'SeaGreen', 'SeaShell', 'Sienna', 'Silver', 'SkyBlue', 'SlateBlue', 'SlateGray', 'Snow', 'SpringGreen', 'SteelBlue', 'Tan', 'Teal', 'Thistle', 'Tomato', 'Transparent', 'Turquoise',
            'Violet', 'Wheat', 'White', 'WhiteSmoke', 'Yellow', 'YellowGreen')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'Pattern1Color1Theme')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'Pattern2Colors')]
        [string]$ForeGroundColor,


        [Validateset('AliceBlue', 'AntiqueWhite', 'Aqua', 'Aquamarine', 'Azure', 'Beige', 'Bisque', 'Black', 'BlanchedAlmond', 'Blue', 'BlueViolet', 'Brown', 'BurlyWood', 'CadetBlue', 'Chartreuse', 'Chocolate', 'Coral', 'CornflowerBlue', 'Cornsilk',
            'Crimson', 'Cyan', 'DarkBlue', 'DarkCyan', 'DarkGoldenrod', 'DarkGray', 'DarkGreen', 'DarkKhaki', 'DarkMagenta', 'DarkOliveGreen', 'DarkOrange', 'DarkOrchid', 'DarkRed', 'DarkSalmon', 'DarkSeaGreen', 'DarkSlateBlue', 'DarkSlateGray',
            'DarkTurquoise', 'DarkViolet', 'DeepPink', 'DeepSkyBlue', 'DimGray', 'DodgerBlue', 'Firebrick', 'FloralWhite', 'ForestGreen', 'Fuchsia', 'Gainsboro', 'GhostWhite', 'Gold', 'Goldenrod', 'Gray', 'Green', 'GreenYellow', 'Honeydew', 'HotPink', 'IndianRed',
            'Indigo', 'Ivory', 'Khaki', 'LavENDer', 'LavENDerBlush', 'LawnGreen', 'LemonChiffon', 'LightBlue', 'LightCoral', 'LightCyan', 'LightGoldenrodYellow', 'LightGray', 'LightGreen', 'LightPink', 'LightSalmon', 'LightSeaGreen', 'LightSkyBlue', 'LightSlateGray',
            'LightSteelBlue', 'LightYellow', 'Lime', 'LimeGreen', 'Linen', 'Magenta', 'Maroon', 'MediumAquamarine', 'MediumBlue', 'MediumOrchid', 'MediumPurple', 'MediumSeaGreen', 'MediumSlateBlue', 'MediumSpringGreen', 'MediumTurquoise', 'MediumVioletRed', 'MidnightBlue'
            , 'MintCream', 'MistyRose', 'Moccasin', 'Name', 'NavajoWhite', 'Navy', 'OldLace', 'Olive', 'OliveDrab', 'Orange', 'OrangeRed', 'Orchid', 'PaleGoldenrod', 'PaleGreen', 'PaleTurquoise', 'PaleVioletRed', 'PapayaWhip', 'PeachPuff', 'Peru', 'Pink', 'Plum', 'PowderBlue',
            'Purple', 'Red', 'RosyBrown', 'RoyalBlue', 'SaddleBrown', 'Salmon', 'SandyBrown', 'SeaGreen', 'SeaShell', 'Sienna', 'Silver', 'SkyBlue', 'SlateBlue', 'SlateGray', 'Snow', 'SpringGreen', 'SteelBlue', 'Tan', 'Teal', 'Thistle', 'Tomato', 'Transparent', 'Turquoise',
            'Violet', 'Wheat', 'White', 'WhiteSmoke', 'Yellow', 'YellowGreen')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'Pattern1Theme1Color')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'Pattern2Colors')]
        [string]$BackGroundColor,


        [Validateset('Accent1Color', 'Accent2Color', 'Accent3Color', 'Accent4Color', 'Accent5Color',
            'Accent6Color', 'Dark1Color', 'Dark2Color', 'Light1Color',
            'Light2Color', 'Hyperlink', 'FollowedHyperlinkColor')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'Pattern1Theme1Color')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'Pattern2ThemeColors')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'GradientFill2ThemeColors')]
        [string]$ForeGroundThemeColor,

        [Validateset('Accent1Color', 'Accent2Color', 'Accent3Color', 'Accent4Color', 'Accent5Color',
            'Accent6Color', 'Dark1Color', 'Dark2Color', 'Light1Color',
            'Light2Color', 'Hyperlink', 'FollowedHyperlinkColor')]

        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'Pattern1Color1Theme')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'Pattern2ThemeColors')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'GradientFill2ThemeColors')]
        [string]$BackGroundThemeColor,

        [Validateset('Corner1', 'Corner2', 'Corner3', 'Corner4', 'DiagonalDown1',
            'DiagonalDown2', 'DiagonalDown3', 'DiagonalUp1', 'DiagonalUp2',
            'DiagonalUp3', 'Horizontal1', 'Horizontal2', 'Horizontal3',
            'Vertical1', 'Vertical2', 'Vertical3', 'FromCenter')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'GradientFill2Colors')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'GradientFill2ThemeColors')]
        [string]$GradientDirection


    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {

            Switch -Regex ($TargetCellorRange)
            {

                #CellReference
                '^[a-zA-Z]+\d+$'
                {
                    Write-Verbose ("Set-SLFill :`tTargetCellorRange is CellReference '{0}'" -f $TargetCellorRange)
                    $SLStyle = $WorkBookInstance.GetCellStyle($TargetCellorRange)
                    $isValidationTargetValid = $true
                    $isCellReference = $true
                    Break
                }

                #Range
                '[a-zA-Z]+\d+:[a-zA-Z]+\d+$'
                {
                    $startcellreference, $endcellreference = $TargetCellorRange -split ':'
                    Write-Verbose ("Set-SLFill :`tTargetCellorRange is CellRange '{0}'" -f $TargetCellorRange)
                    $SLStyle = $WorkBookInstance.CreateStyle()
                    $isValidationTargetValid = $true
                    $isRange = $true
                    Break
                }

                Default
                {
                    Write-Warning ("Set-SLDataValidation :`tYou must provide either a Cellreference Eg. C3 or a Range Eg. C3:G10")
                    $isValidationTargetValid = $false
                    Break
                }

            }#switch


            if ($PSCmdlet.ParameterSetName -eq 'Pattern1Theme' -and $isValidationTargetValid )
            {
                Write-Verbose ("Set-SLFill :`tPattern 'Solid' with ThemeColor '{0}' selected" -f $ThemeColor)
                $SLStyle.Fill.SetPatternType([DocumentFormat.OpenXml.Spreadsheet.PatternValues]::'Solid') | Out-Null
                $SLStyle.Fill.SetPatternForegroundColor([SpreadsheetLight.SLThemeColorIndexValues]::$ThemeColor) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Pattern1Color' -and $isValidationTargetValid )
            {
                Write-Verbose ("Set-SLFill :`tPattern 'Solid' with Color '{0}' selected" -f $Color)
                $SLStyle.Fill.SetPatternType([DocumentFormat.OpenXml.Spreadsheet.PatternValues]::'Solid') | Out-Null
                $SLStyle.Fill.SetPatternForegroundColor([System.Drawing.Color]::$Color) | Out-Null
            }


            if ($PSCmdlet.ParameterSetName -eq 'Pattern1ColorHtml' -and $isValidationTargetValid )
            {
                Write-Verbose ("Set-SLFill :`tPattern 'Solid' with HTML Color value '{0}' selected" -f $Color)
                $SLStyle.Fill.SetPatternType([DocumentFormat.OpenXml.Spreadsheet.PatternValues]::'Solid') | Out-Null
                $SLStyle.Fill.SetPatternForegroundColor([System.Drawing.ColorTranslator]::FromHtml($ColorFromHTML))
            }

            if ($PSCmdlet.ParameterSetName -eq 'Pattern2Colors' -and $isValidationTargetValid )
            {
                Write-Verbose ("Set-SLFill :`tPattern '{0}' with ForegroundColor '{1}' & BackGroundColor '{2}' selected" -f $pattern, $ForeGroundColor, $BackGroundColor)
                $SLStyle.Fill.SetPattern([DocumentFormat.OpenXml.Spreadsheet.PatternValues]::$pattern, [System.Drawing.Color]::$ForeGroundColor, [System.Drawing.Color]::$BackGroundColor   ) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Pattern2ThemeColors' -and $isValidationTargetValid )
            {
                Write-Verbose ("Set-SLFill :`tPattern '{0}' with ForegroundThemeColor '{1}' & BackGroundThemeColor '{2}' selected" -f $pattern, $ForeGroundThemeColor, $BackGroundThemeColor)
                $SLStyle.Fill.SetPattern([DocumentFormat.OpenXml.Spreadsheet.PatternValues]::$pattern, [SpreadsheetLight.SLThemeColorIndexValues]::$ForeGroundThemeColor, [SpreadsheetLight.SLThemeColorIndexValues]::$BackGroundThemeColor   ) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Pattern1Theme1Color' -and $isValidationTargetValid )
            {
                Write-Verbose ("Set-SLFill :`tPattern '{0}' with ForegroundThemeColor '{1}' & BackGroundColor '{2}' selected" -f $pattern, $ForeGroundThemeColor, $BackGroundColor)
                $SLStyle.Fill.SetPattern([DocumentFormat.OpenXml.Spreadsheet.PatternValues]::$pattern, [SpreadsheetLight.SLThemeColorIndexValues]::$ForeGroundThemeColor, [System.Drawing.Color]::$BackGroundColor   ) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Pattern1Color1Theme' -and $isValidationTargetValid )
            {
                Write-Verbose ("Set-SLFill :`tPattern '{0}' with ForegroundColor '{1}' & BackGroundThemeColor '{2}' selected" -f $pattern, $ForeGroundColor, $BackGroundThemeColor)
                $SLStyle.Fill.SetPattern([DocumentFormat.OpenXml.Spreadsheet.PatternValues]::$pattern, [System.Drawing.Color]::$ForeGroundColor, [SpreadsheetLight.SLThemeColorIndexValues]::$BackGroundThemeColor   ) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'GradientFill2ThemeColors' -and $isValidationTargetValid )
            {
                Write-Verbose ("Set-SLFill :`tGradientDirection '{0}' with ForeGroundThemeColor '{1}' & BackGroundThemeColor '{2}' selected" -f $GradientDirection, $ForeGroundThemeColor, $BackGroundThemeColor)
                $SLStyle.SetGradientFill([SpreadsheetLight.SLGradientShadingStyleValues]::$GradientDirection, [SpreadsheetLight.SLThemeColorIndexValues]::$ForeGroundThemeColor, [SpreadsheetLight.SLThemeColorIndexValues]::$BackGroundThemeColor) | Out-Null

            }

            if ($PSCmdlet.ParameterSetName -eq 'GradientFill2Colors' -and $isValidationTargetValid )
            {
                Write-Verbose ("Set-SLFill :`tGradientDirection '{0}' with ForeGroundColor '{1}' & BackGroundColor '{2}' selected" -f $GradientDirection, $ForeGroundColor, $BackGroundColor)
                $SLStyle.SetGradientFill([SpreadsheetLight.SLGradientShadingStyleValues]::$GradientDirection, [System.Drawing.Color]::$ForeGroundColor, [System.Drawing.Color]::$BackGroundColor) | Out-Null

            }


            if ($isValidationTargetValid)
            {

                If ($isCellReference)
                {
                    Write-Verbose ("Set-SLFill :`tAdding Fill style..")
                    $WorkBookInstance.SetCellStyle($TargetCellorRange, $SLStyle) | Out-Null
                    $WorkBookInstance | Add-Member NoteProperty CellReference $TargetCellorRange -Force
                }
                Elseif ($isRange)
                {
                    Write-Verbose ("Set-SLFill :`tAdding Fill style..")
                    $WorkBookInstance.SetCellStyle($startcellreference, $endcellreference, $SLStyle) | Out-Null
                    $WorkBookInstance | Add-Member NoteProperty Range $TargetCellorRange -Force
                }
                $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
            }

        }#select-slworksheet

    }#process

}
