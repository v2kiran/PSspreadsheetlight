Function Set-SLConditionalFormatColorScale  {


    <#

.SYNOPSIS
    Apply conditional formatting color scale to a range.

.DESCRIPTION
    Apply conditional formatting color scale to a range.
    Cells are shaded with gradations of two or three colors that correspond to minimum, midpoint, and maximum thresholds.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    This is the name of the worksheet that contains the cell range where formatting is to be applied.

.PARAMETER Range
    The range of cells where conditional formatting has to be applied.

.PARAMETER ColorScaleType
    Built-in color scale styles.
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'GreenYellowRed','RedYellowGreen','BlueYellowRed','RedYellowBlue','GreenWhiteRed','RedWhiteGreen','BlueWhiteRed','RedWhiteBlue','WhiteRed','RedWhite','GreenWhite','WhiteGreen','Yellow',
    'Red','RedYellow','GreenYellow','YellowGreen'

.PARAMETER ColorScaleMinType
    to be used with a custom color scale formatting style.
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
   'Value','Number','Percent','Formula','Percentile'

.PARAMETER MinValue
    the minimum value in the range.

.PARAMETER ColorScaleMinSystemColor
    Custom color for the minimum values.

.PARAMETER ColorScaleMaxType
    to be used with a custom color scale formatting style.
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
   'Value','Number','Percent','Formula','Percentile'

.PARAMETER MaxValue
    The maximum value in the range.

.PARAMETER ColorScaleMaxSystemColor
    Custom color for the maximum values.

.PARAMETER ColorScale2
    to be used with a custom 2colorscale formatting style.

.PARAMETER ColorScale3
    to be used with a custom 3colorscale formatting style.

.PARAMETER MidPointType
    to be used with a custom 3color scale formatting style.
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
   'Number','Percent','Formula','Percentile'

.PARAMETER MidPointValue
    The mid value in the range.

.PARAMETER MidPointColor
    Custom color for the mid values.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatColorScale -WorksheetName sheet7 -Range D4:D15 -ColorScaleType GreenYellowRed -Verbose
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    apply the built-in 3colorscale style 'GreenyellowRed' on range D4:D15.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatColorScale -WorksheetName sheet7 -Range F4:F15 -ColorScaleMinType Number -MinValue 12 -ColorScaleMinSystemColor Crimson -ColorScaleMaxType Number -MaxValue 99 -ColorScaleMaxSystemColor Yellow -ColorScale2
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    apply a custom 2colorscale style on range F4:F15.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatColorScale -WorksheetName sheet7 -Range h4:h15 -ColorScaleMinType Number -MinValue 12 -ColorScaleMinSystemColor Crimson -ColorScaleMaxType Number -MaxValue 99 -ColorScaleMaxSystemColor Yellow -MidPointType Number -MidPointValue 60 -MidPointColor Beige -ColorScale3
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    apply a custom 3colorscale style on range h4:h15.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> for($i=18 ;$i -le 89 ; $i++) { Set-SLConditionalFormatColorScale - -WorkBookInstance $doc -WorksheetName sheet7 -Range "C$($i):G$($i)" -ColorScaleType RedYellow -Verbose }
    PS C:\> $doc | Save-SLDocument


    Description
    -----------
    At times we may want to apply color scale formatting to individual rows instead of a range or rows.
    This example makes use of a for-loop to loop through rows 19 to 89 while applying the built-in style of 'RedYellow' on each row.


.INPUTS
   String,SpreadsheetLight.SLDocument

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    n\a

#>



    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true, Position = 0, ValueFromPipeLine = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [Alias('CurrentWorkSheetName')]
        [parameter(Mandatory = $true, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+:[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Set-SLConditionalFormatColorScale :`tRange should specify values in following format. Eg: A1:D10 or AB1:AD5"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true)]
        [string]$Range,

        [ValidateSet('GreenYellowRed', 'RedYellowGreen', 'BlueYellowRed', 'RedYellowBlue', 'GreenWhiteRed', 'RedWhiteGreen', 'BlueWhiteRed', 'RedWhiteBlue', 'WhiteRed', 'RedWhite', 'GreenWhite', 'WhiteGreen', 'Yellow',
            'Red', 'RedYellow', 'GreenYellow', 'YellowGreen')]
        [parameter(Mandatory = $True, Position = 3, ParameterSetName = 'Normal')]
        [string]$ColorScaleType,

        [ValidateSet('Value', 'Number', 'Percent', 'Formula', 'Percentile')]
        [parameter(Mandatory = $True, ParameterSetName = 'Custom3ColorScale')]
        [parameter(Mandatory = $True, ParameterSetName = 'Custom2ColorScale')]
        [String]$ColorScaleMinType,

        [parameter(Mandatory = $True, ParameterSetName = 'Custom3ColorScale')]
        [parameter(Mandatory = $True, ParameterSetName = 'Custom2ColorScale')]
        $MinValue,

        [Validateset('AliceBlue', 'AntiqueWhite', 'Aqua', 'Aquamarine', 'Azure', 'Beige', 'Bisque', 'Black', 'BlanchedAlmond', 'Blue', 'BlueViolet', 'Brown', 'BurlyWood', 'CadetBlue', 'Chartreuse', 'Chocolate', 'Coral', 'CornflowerBlue', 'Cornsilk',
            'Crimson', 'Cyan', 'DarkBlue', 'DarkCyan', 'DarkGoldenrod', 'DarkGray', 'DarkGreen', 'DarkKhaki', 'DarkMagenta', 'DarkOliveGreen', 'DarkOrange', 'DarkOrchid', 'DarkRed', 'DarkSalmon', 'DarkSeaGreen', 'DarkSlateBlue', 'DarkSlateGray',
            'DarkTurquoise', 'DarkViolet', 'DeepPink', 'DeepSkyBlue', 'DimGray', 'DodgerBlue', 'Firebrick', 'FloralWhite', 'ForestGreen', 'Fuchsia', 'Gainsboro', 'GhostWhite', 'Gold', 'Goldenrod', 'Gray', 'Green', 'GreenYellow', 'Honeydew', 'HotPink', 'IndianRed',
            'Indigo', 'Ivory', 'Khaki', 'LavENDer', 'LavENDerBlush', 'LawnGreen', 'LemonChiffon', 'LightBlue', 'LightCoral', 'LightCyan', 'LightGoldenrodYellow', 'LightGray', 'LightGreen', 'LightPink', 'LightSalmon', 'LightSeaGreen', 'LightSkyBlue', 'LightSlateGray',
            'LightSteelBlue', 'LightYellow', 'Lime', 'LimeGreen', 'Linen', 'Magenta', 'Maroon', 'MediumAquamarine', 'MediumBlue', 'MediumOrchid', 'MediumPurple', 'MediumSeaGreen', 'MediumSlateBlue', 'MediumSpringGreen', 'MediumTurquoise', 'MediumVioletRed', 'MidnightBlue'
            , 'MintCream', 'MistyRose', 'Moccasin', 'Name', 'NavajoWhite', 'Navy', 'OldLace', 'Olive', 'OliveDrab', 'Orange', 'OrangeRed', 'Orchid', 'PaleGoldenrod', 'PaleGreen', 'PaleTurquoise', 'PaleVioletRed', 'PapayaWhip', 'PeachPuff', 'Peru', 'Pink', 'Plum', 'PowderBlue',
            'Purple', 'Red', 'RosyBrown', 'RoyalBlue', 'SaddleBrown', 'Salmon', 'SandyBrown', 'SeaGreen', 'SeaShell', 'Sienna', 'Silver', 'SkyBlue', 'SlateBlue', 'SlateGray', 'Snow', 'SpringGreen', 'SteelBlue', 'Tan', 'Teal', 'Thistle', 'Tomato', 'Transparent', 'Turquoise',
            'Violet', 'Wheat', 'White', 'WhiteSmoke', 'Yellow', 'YellowGreen')]
        [parameter(Mandatory = $True, ParameterSetName = 'Custom3ColorScale')]
        [parameter(Mandatory = $True, ParameterSetName = 'Custom2ColorScale')]
        [string]$ColorScaleMinSystemColor,

        [ValidateSet('Value', 'Number', 'Percent', 'Formula', 'Percentile')]
        [parameter(Mandatory = $True, ParameterSetName = 'Custom3ColorScale')]
        [parameter(Mandatory = $True, ParameterSetName = 'Custom2ColorScale')]
        [String]$ColorScaleMaxType,

        [parameter(Mandatory = $True, ParameterSetName = 'Custom3ColorScale')]
        [parameter(Mandatory = $True, ParameterSetName = 'Custom2ColorScale')]
        $MaxValue,

        [Validateset('AliceBlue', 'AntiqueWhite', 'Aqua', 'Aquamarine', 'Azure', 'Beige', 'Bisque', 'Black', 'BlanchedAlmond', 'Blue', 'BlueViolet', 'Brown', 'BurlyWood', 'CadetBlue', 'Chartreuse', 'Chocolate', 'Coral', 'CornflowerBlue', 'Cornsilk',
            'Crimson', 'Cyan', 'DarkBlue', 'DarkCyan', 'DarkGoldenrod', 'DarkGray', 'DarkGreen', 'DarkKhaki', 'DarkMagenta', 'DarkOliveGreen', 'DarkOrange', 'DarkOrchid', 'DarkRed', 'DarkSalmon', 'DarkSeaGreen', 'DarkSlateBlue', 'DarkSlateGray',
            'DarkTurquoise', 'DarkViolet', 'DeepPink', 'DeepSkyBlue', 'DimGray', 'DodgerBlue', 'Firebrick', 'FloralWhite', 'ForestGreen', 'Fuchsia', 'Gainsboro', 'GhostWhite', 'Gold', 'Goldenrod', 'Gray', 'Green', 'GreenYellow', 'Honeydew', 'HotPink', 'IndianRed',
            'Indigo', 'Ivory', 'Khaki', 'LavENDer', 'LavENDerBlush', 'LawnGreen', 'LemonChiffon', 'LightBlue', 'LightCoral', 'LightCyan', 'LightGoldenrodYellow', 'LightGray', 'LightGreen', 'LightPink', 'LightSalmon', 'LightSeaGreen', 'LightSkyBlue', 'LightSlateGray',
            'LightSteelBlue', 'LightYellow', 'Lime', 'LimeGreen', 'Linen', 'Magenta', 'Maroon', 'MediumAquamarine', 'MediumBlue', 'MediumOrchid', 'MediumPurple', 'MediumSeaGreen', 'MediumSlateBlue', 'MediumSpringGreen', 'MediumTurquoise', 'MediumVioletRed', 'MidnightBlue'
            , 'MintCream', 'MistyRose', 'Moccasin', 'Name', 'NavajoWhite', 'Navy', 'OldLace', 'Olive', 'OliveDrab', 'Orange', 'OrangeRed', 'Orchid', 'PaleGoldenrod', 'PaleGreen', 'PaleTurquoise', 'PaleVioletRed', 'PapayaWhip', 'PeachPuff', 'Peru', 'Pink', 'Plum', 'PowderBlue',
            'Purple', 'Red', 'RosyBrown', 'RoyalBlue', 'SaddleBrown', 'Salmon', 'SandyBrown', 'SeaGreen', 'SeaShell', 'Sienna', 'Silver', 'SkyBlue', 'SlateBlue', 'SlateGray', 'Snow', 'SpringGreen', 'SteelBlue', 'Tan', 'Teal', 'Thistle', 'Tomato', 'Transparent', 'Turquoise',
            'Violet', 'Wheat', 'White', 'WhiteSmoke', 'Yellow', 'YellowGreen')]
        [parameter(Mandatory = $True, ParameterSetName = 'Custom3ColorScale')]
        [parameter(Mandatory = $True, ParameterSetName = 'Custom2ColorScale')]
        [string]$ColorScaleMaxSystemColor,


        [parameter(ParameterSetName = 'Custom2ColorScale')]
        [Switch]$ColorScale2,

        [parameter(ParameterSetName = 'Custom3ColorScale')]
        [Switch]$ColorScale3,

        [ValidateSet('Number', 'Percent', 'Formula', 'Percentile')]
        [parameter(Mandatory = $True, ParameterSetName = 'Custom3ColorScale')]
        [String]$MidPointType,

        [parameter(Mandatory = $True, ParameterSetName = 'Custom3ColorScale')]
        [String]$MidPointValue,

        [Validateset('AliceBlue', 'AntiqueWhite', 'Aqua', 'Aquamarine', 'Azure', 'Beige', 'Bisque', 'Black', 'BlanchedAlmond', 'Blue', 'BlueViolet', 'Brown', 'BurlyWood', 'CadetBlue', 'Chartreuse', 'Chocolate', 'Coral', 'CornflowerBlue', 'Cornsilk',
            'Crimson', 'Cyan', 'DarkBlue', 'DarkCyan', 'DarkGoldenrod', 'DarkGray', 'DarkGreen', 'DarkKhaki', 'DarkMagenta', 'DarkOliveGreen', 'DarkOrange', 'DarkOrchid', 'DarkRed', 'DarkSalmon', 'DarkSeaGreen', 'DarkSlateBlue', 'DarkSlateGray',
            'DarkTurquoise', 'DarkViolet', 'DeepPink', 'DeepSkyBlue', 'DimGray', 'DodgerBlue', 'Firebrick', 'FloralWhite', 'ForestGreen', 'Fuchsia', 'Gainsboro', 'GhostWhite', 'Gold', 'Goldenrod', 'Gray', 'Green', 'GreenYellow', 'Honeydew', 'HotPink', 'IndianRed',
            'Indigo', 'Ivory', 'Khaki', 'LavENDer', 'LavENDerBlush', 'LawnGreen', 'LemonChiffon', 'LightBlue', 'LightCoral', 'LightCyan', 'LightGoldenrodYellow', 'LightGray', 'LightGreen', 'LightPink', 'LightSalmon', 'LightSeaGreen', 'LightSkyBlue', 'LightSlateGray',
            'LightSteelBlue', 'LightYellow', 'Lime', 'LimeGreen', 'Linen', 'Magenta', 'Maroon', 'MediumAquamarine', 'MediumBlue', 'MediumOrchid', 'MediumPurple', 'MediumSeaGreen', 'MediumSlateBlue', 'MediumSpringGreen', 'MediumTurquoise', 'MediumVioletRed', 'MidnightBlue'
            , 'MintCream', 'MistyRose', 'Moccasin', 'Name', 'NavajoWhite', 'Navy', 'OldLace', 'Olive', 'OliveDrab', 'Orange', 'OrangeRed', 'Orchid', 'PaleGoldenrod', 'PaleGreen', 'PaleTurquoise', 'PaleVioletRed', 'PapayaWhip', 'PeachPuff', 'Peru', 'Pink', 'Plum', 'PowderBlue',
            'Purple', 'Red', 'RosyBrown', 'RoyalBlue', 'SaddleBrown', 'Salmon', 'SandyBrown', 'SeaGreen', 'SeaShell', 'Sienna', 'Silver', 'SkyBlue', 'SlateBlue', 'SlateGray', 'Snow', 'SpringGreen', 'SteelBlue', 'Tan', 'Teal', 'Thistle', 'Tomato', 'Transparent', 'Turquoise',
            'Violet', 'Wheat', 'White', 'WhiteSmoke', 'Yellow', 'YellowGreen')]
        [parameter(Mandatory = $True, ParameterSetName = 'Custom3ColorScale')]
        [String]$MidPointColor



    )
    PROCESS
    {

        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {

            $startcellreference, $ENDcellreference = $range -split ':'
            $ConditionalFormatting = New-Object SpreadsheetLight.SLConditionalFormatting($startcellreference, $ENDcellreference)

            if ($PSCmdlet.ParameterSetName -eq 'Normal')
            {
                Write-Verbose ("Set-SLConditionalFormatColorScale :`t Selected ColorScaleType is '{0}'" -f $ColorScaleType)
                $ConditionalFormatting.SetColorScale([SpreadsheetLight.SLConditionalFormatColorScaleValues]::$ColorScaleType) | Out-Null
            }


            if ($PSCmdlet.ParameterSetName -eq 'Custom2ColorScale')
            {
                $ConditionalFormatting.SetCustom2ColorScale([SLCFMinMax]::$ColorScaleMinType, $MinValue, [Color]::$ColorScaleMinSystemColor, [SLCFMinMax]::$ColorScaleMaxType, $MaxValue, [Color]::$ColorScaleMaxSystemColor) | Out-Null
            }



            if ($PSCmdlet.ParameterSetName -eq 'Custom3ColorScale')
            {
                $ConditionalFormatting.SetCustom3ColorScale([SLCFMinMax]::$ColorScaleMinType, $MinValue, [Color]::$ColorScaleMinSystemColor, [SpreadsheetLight.SLConditionalFormatRangeValues]::$MidPointType, $MidPointValue, [color]::$MidPointColor, [SLCFMinMax]::$ColorScaleMaxType, $MaxValue, [Color]::$ColorScaleMaxSystemColor) | Out-Null
            }

            Write-Verbose ("Set-SLConditionalFormatColorScale :`t Applying conditional formatting color scale on Range '{0}'" -f $Range)
            $WorkBookInstance.AddConditionalFormatting($ConditionalFormatting) | Out-Null

            $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }#select-slworksheet

    }#process
    END
    {
    }

}
