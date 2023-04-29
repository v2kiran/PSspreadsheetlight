Function Set-SLConditionalFormatHighLights  {


    <#

.SYNOPSIS
    Apply conditional formatting Highlights.

.DESCRIPTION
    Apply conditional formatting Highlights on cells.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    This is the name of the worksheet that contains the cell range where formatting is to be applied.

.PARAMETER Range
    The range of cells containing text to which conditional formatting has to be applied.

.PARAMETER StyleType
    Choose between excel's 'PresetStyle' or a 'CustomStyle'.

.PARAMETER PresetStyleValue
    Built-in Preset styles.
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'LightRedFillWithDarkRedText','YellowFillWithDarkYellowText','GreenFillWithDarkGreenText','LightRedFill','RedText','RedBorder'

.PARAMETER IsBetweenFirstValue
    Minumum value to be used when specifying a range that is 'Between' two values.

.PARAMETER IsBetweenLastValue
    Maximum value to be used when specifying a range that is 'Between' two values.

.PARAMETER IsNotBetweenFirstValue
    Minumum value to be used when specifying a range that is 'NOTBetween' two values.

.PARAMETER IsNotBetweenLastValue
    Maximum value to be used when specifying a range that is 'NOTBetween' two values.

.PARAMETER TopRankValue
    Top rank value to be used when specifying top and bottom ranks.

.PARAMETER BottomRankValue
    Bottom rank value to be used when specifying top and bottom ranks.

.PARAMETER IsPercent
    Specifies that values should be considered as numbers.

.PARAMETER IsItems
    Specifies that values should be considered as a percentage.

.PARAMETER GreaterThanValue
    Highlight values that are greater than this value.

.PARAMETER GreaterThanorEqualToValue
    Highlight values that are greater than or Equalto this value.

.PARAMETER LessThanValue
    Highlight values that are less than this value.

.PARAMETER LessThanorEqualToValue
    Highlight values that are less than or Equalto this value.

.PARAMETER EqualToValue
    Highlight values that are Equalto this value.

.PARAMETER NotEqualToValue
    Highlight values that are NOTEqualto this value.

.PARAMETER TextContainsString
    Highlight cells that contain this string.

.PARAMETER TextDoesNotContainString
    Highlight cells that DO NOT contain this string.

.PARAMETER TextENDsWithString
    Highlight cells that End with this string.

.PARAMETER TextBEGINsWithString
    Highlight cells that begin with this string.

.PARAMETER AverageType
    Built-in AverageType values.
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'Above','Below','EqualOrAbove','EqualOrBelow','OneStdDevAbove','OneStdDevBelow',
    'TwoStdDevAbove','TwoStdDevBelow','ThreeStdDevAbove','ThreeStdDevBelow'

.PARAMETER DateString
    Highlight cells that match the date specified by this value.

.PARAMETER FormulaString
    Highlight cells that match the criteria specified by a formula.

.PARAMETER HighLightDuplicateValues
    Highlight all duplicate values in a range.

.PARAMETER HighLightUniqueValues
    Highlight all unique values in a range.

.PARAMETER HighlightBlankCells
    Highlight all blank cells in a range.

.PARAMETER HighlightNonBlankCells
    Highlight all non-blank cells in a range.

.PARAMETER HighlightErrorCells
    Highlight all cells containing formula errors in a range.

.PARAMETER HighlightNonErrorCells
    Highlight all cells that dont contain formula errors in a range.

.PARAMETER FontColor
    Fontcolor to be specified when using a custom highlight style.

.PARAMETER FontIsBold
    Specify that the font is bold when using a custom highlight style.

.PARAMETER FontIsItalic
    Specify that the font is italic when using a custom highlight style.

.PARAMETER FontIsUnderlined
    Specify that the font is underlined when using a custom highlight style.

.PARAMETER FillColor
    Fill Color to be specified when using a custom highlight style.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatHighLights -WorksheetName sheet9 -Range a2:b9 -StyleType PresetStyle -PresetStyleValue LightRedFill -HighLightDuplicateValues | Save-SLDocument

    Description
    -----------
    Highlight Duplicate values


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatHighLights -WorksheetName sheet9 -Range a11:a21 -StyleType PresetStyle -PresetStyleValue RedBorder -HighlightBlankCells | Save-SLDocument

    Description
    -----------
    Highlight blanks cells.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatHighLights -WorksheetName sheet9 -Range a11:a21 -StyleType PresetStyle -PresetStyleValue RedText -HighlightNonBlankCells | Save-SLDocument

    Description
    -----------
    Highlight non-blank cells


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatHighLights -WorksheetName sheet9 -Range b11:b21 -StyleType PresetStyle -PresetStyleValue LightRedFillWithDarkRedText -HighlightErrorCells | Save-SLDocument

    Description
    -----------
    Highlight Error cells.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatHighLights -WorksheetName sheet9 -Range b11:b21 -StyleType PresetStyle -PresetStyleValue GreenFillWithDarkGreenText -HighlightNonErrorCells | Save-SLDocument

    Description
    -----------
    Highlight non-Error cells


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatHighLights -WorksheetName sheet9 -Range C2:C9 -StyleType PresetStyle -PresetStyleValue YellowFillWithDarkYellowText -IsBetweenFirstValue 200 -IsBetweenLastValue 400 | Save-SLDocument

    Description
    -----------
    Highlight values between 200 and 400.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatHighLights -WorksheetName sheet9 -Range C2:C9 -StyleType PresetStyle -PresetStyleValue GreenFillWithDarkGreenText -IsNotBetweenFirstValue 200 -IsNotBetweenLastValue  400 | Save-SLDocument

    Description
    -----------
    Highlight values NOT between 200 and 400


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatHighLights -WorksheetName sheet9 -Range c11:c21 -StyleType PresetStyle -PresetStyleValue GreenFillWithDarkGreenText -TopRankValue 25 -IsPercent | Save-SLDocument

    Description
    -----------
    Highlight TOP 25% of the values.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatHighLights -WorksheetName sheet9 -Range E2:E9 -StyleType PresetStyle -PresetStyleValue YellowFillWithDarkYellowText -TopRankValue 3 -IsItems | Save-SLDocument

    Description
    -----------
    Highlight TOP 3 items.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatHighLights -WorksheetName sheet9 -Range c11:c21 -StyleType PresetStyle -PresetStyleValue LightRedFillWithDarkRedText  -BottomRankValue 25 -IsPercent | Save-SLDocument

    Description
    -----------
    Highlight BOTTOM 25%.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatHighLights -WorksheetName sheet9 -Range E2:E9 -StyleType PresetStyle -PresetStyleValue LightRedFill -BottomRankValue 3 -IsItems | Save-SLDocument

    Description
    -----------
    Highlight BOTTOM 3 items


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatHighLights -WorksheetName sheet9 -Range E11:E21 -StyleType CustomStyle -GreaterThanValue 200 -FontColor Blue -FontIsBold -FontIsItalic -FontIsUnderlined -FillColor Yellow | Save-SLDocument

    Description
    -----------
    Highlight values Greaterthan 200 - Custom Style.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatHighLights -WorksheetName sheet9 -Range E11:E21 -StyleType CustomStyle -LessThanValue 11 -FontColor White -FontIsBold -FillColor Darkblue | Save-SLDocument

    Description
    -----------
    Highlight values LessThan 11 - Custom Style


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatHighLights -WorksheetName sheet9 -Range G2:G9 -StyleType PresetStyle -PresetStyleValue GreenFillWithDarkGreenText -TextENDsWithString bob | Save-SLDocument

    Description
    -----------
    Highlight cells that END with 'Bob'


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatHighLights -WorksheetName sheet9 -Range G2:G9 -StyleType PresetStyle -PresetStyleValue YellowFillWithDarkYellowText  -TextContainsString jones | Save-SLDocument

    Description
    -----------
    Highlight cells that Contain 'Jones'.


.INPUTS
   String,Int,Date,SpreadsheetLight.SLDocument

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
                else { $false; Write-Warning "Set-SLConditionalFormattingHighLights :`tRange should specify values in following format. Eg: A1:D10 or AB1:AD5"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true)]
        [string]$Range,

        [ValidateSet('CustomStyle', 'PresetStyle')]
        [parameter(Mandatory = $True)]
        [String]$StyleType,

        [ValidateSet('LightRedFillWithDarkRedText', 'YellowFillWithDarkYellowText', 'GreenFillWithDarkGreenText', 'LightRedFill', 'RedText', 'RedBorder')]
        [parameter(Mandatory = $false)]
        [string]$PresetStyleValue,

        [parameter(Mandatory = $True, ParameterSetName = 'Between')]
        [Double]$IsBetweenFirstValue,

        [parameter(Mandatory = $True, ParameterSetName = 'Between')]
        [Double]$IsBetweenLastValue,

        [parameter(Mandatory = $True, ParameterSetName = 'NotBetween')]
        [Double]$IsNotBetweenFirstValue,

        [parameter(Mandatory = $True, ParameterSetName = 'NotBetween')]
        [Double]$IsNotBetweenLastValue,

        [parameter(Mandatory = $True, ParameterSetName = 'TopRank')]
        [System.UInt32]$TopRankValue,

        [parameter(Mandatory = $True, ParameterSetName = 'BottomRank')]
        [System.UInt32]$BottomRankValue,

        [parameter(ParameterSetName = 'TopRank')]
        [parameter(ParameterSetName = 'BottomRank')]
        [Switch]$IsPercent,

        [parameter(ParameterSetName = 'TopRank')]
        [parameter(ParameterSetName = 'BottomRank')]
        [Switch]$IsItems,

        [parameter(Mandatory = $True, ParameterSetName = 'GreaterThan')]
        [Double]$GreaterThanValue,

        [parameter(Mandatory = $True, ParameterSetName = 'GreaterThanorEqualTo')]
        [Double]$GreaterThanorEqualToValue,

        [parameter(Mandatory = $True, ParameterSetName = 'LessThan')]
        [Double]$LessThanValue,

        [parameter(Mandatory = $True, ParameterSetName = 'LessThanorEqualTo')]
        [Double]$LessThanorEqualToValue,

        [parameter(Mandatory = $True, ParameterSetName = 'EqualTo')]
        [String]$EqualToValue,

        [parameter(Mandatory = $True, ParameterSetName = 'NotEqualTo')]
        [String]$NotEqualToValue,

        [parameter(Mandatory = $True, ParameterSetName = 'TextContains')]
        [String]$TextContainsString,

        [parameter(Mandatory = $True, ParameterSetName = 'TextDoesNotContain')]
        [String]$TextDoesNotContainString,

        [parameter(Mandatory = $True, ParameterSetName = 'TextENDsWith')]
        [String]$TextENDsWithString,

        [parameter(Mandatory = $True, ParameterSetName = 'TextBEGINsWith')]
        [String]$TextBEGINsWithString,

        [ValidateSet('Above', 'Below', 'EqualOrAbove', 'EqualOrBelow', 'OneStdDevAbove', 'OneStdDevBelow', 'TwoStdDevAbove', 'TwoStdDevBelow', 'ThreeStdDevAbove', 'ThreeStdDevBelow')]
        [parameter(Mandatory = $True, ParameterSetName = 'Average')]
        [String]$AverageType,

        [parameter(Mandatory = $True, ParameterSetName = 'Formula')]
        [String]$FormulaString,

        [parameter(Mandatory = $True, ParameterSetName = 'Date')]
        [String]$DateString,

        [parameter(Mandatory = $True, ParameterSetName = 'Duplicate')]
        [Switch]$HighLightDuplicateValues,

        [parameter(Mandatory = $True, ParameterSetName = 'Unique')]
        [Switch]$HighLightUniqueValues,

        [parameter(Mandatory = $True, ParameterSetName = 'Blank')]
        [Switch]$HighlightBlankCells,

        [parameter(Mandatory = $True, ParameterSetName = 'NonBlank')]
        [Switch]$HighlightNonBlankCells,

        [parameter(Mandatory = $True, ParameterSetName = 'Error')]
        [Switch]$HighlightErrorCells,

        [parameter(Mandatory = $True, ParameterSetName = 'NonError')]
        [Switch]$HighlightNonErrorCells,


        [parameter(Mandatory = $False)]
        [String]$FontColor,

        [Switch]$FontIsBold,

        [Switch]$FontIsItalic,

        [Switch]$FontIsUnderlined,

        [parameter(Mandatory = $False)]
        [String]$FillColor





    )
    PROCESS
    {

        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {

            $startcellreference, $ENDcellreference = $range -split ':'
            $ConditionalFormatting = New-Object SpreadsheetLight.SLConditionalFormatting($startcellreference, $ENDcellreference)

            If ($StyleType -eq 'CustomStyle')
            {
                $SLStyle = $WorkBookInstance.CreateStyle()

                if ($FontColor) { $SLStyle.SetFontColor([System.Drawing.Color]::$FontColor) | Out-Null }
                if ($FontIsBold) { $SLStyle.SetFontBold($true) | Out-Null }
                if ($FontIsItalic) { $SLStyle.SetFontItalic($true) | Out-Null }
                if ($FontIsUnderlined) { $SLStyle.SetFontUnderline([DocumentFormat.OpenXml.Spreadsheet.UnderlineValues]::'Single') | Out-Null }
                if ($FillColor)
                {
                    $SLStyle.Fill.SetPatternType([DocumentFormat.OpenXml.Spreadsheet.PatternValues]::'Solid') | Out-Null
                    $SLStyle.Fill.SetPatternBackgroundColor([System.Drawing.Color]::$FillColor) | Out-Null
                }

            }
            elseif ($StyleType -eq 'PresetStyle')
            {
                $PresetStyle = [SpreadsheetLight.SLHighlightCellsStyleValues]::$PresetStyleValue
            }


            if ($PSCmdlet.ParameterSetName -eq 'GreaterThan')
            {

                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsGreaterThan($False, $GreaterThanValue, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsGreaterThan($False, $GreaterThanValue, $SLStyle) | Out-Null }
            }

            if ($PSCmdlet.ParameterSetName -eq 'GreaterThanorEqualTo')
            {

                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsGreaterThan($True, $GreaterThanorEqualToValue, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsGreaterThan($True, $GreaterThanorEqualToValue, $SLStyle) | Out-Null }
            }

            ## // less than

            if ($PSCmdlet.ParameterSetName -eq 'LessThan')
            {

                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsLessThan($False, $LessThanValue, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsLessThan($False, $LessThanValue, $SLStyle) | Out-Null }
            }

            if ($PSCmdlet.ParameterSetName -eq 'LessThanorEqualTo')
            {

                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsLessThan($True, $LessThanorEqualToValue, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsLessThan($True, $LessThanorEqualToValue, $SLStyle) | Out-Null }
            }

            ## // Between

            if ($PSCmdlet.ParameterSetName -eq 'Between')
            {

                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsBetween($True, $IsBetweenFirstValue, $IsBetweenLastValue, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsBetween($True, $IsBetweenFirstValue, $IsBetweenLastValue, $SLStyle) | Out-Null }
            }

            if ($PSCmdlet.ParameterSetName -eq 'NotBetween')
            {

                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsBetween($False, $IsNotBetweenFirstValue, $IsNotBetweenLastValue, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsBetween($False, $IsNotBetweenFirstValue, $IsNotBetweenLastValue, $SLStyle) | Out-Null }
            }

            ## // Range

            if ($PSCmdlet.ParameterSetName -eq 'TopRank')
            {
                if ($IsPercent) { $percentoritems = $true }
                elseif ($IsItems) { $percentoritems = $false }
                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsInTopRange($True, $TopRankValue, $percentoritems, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsInTopRange($True, $TopRankValue, $percentoritems, $SLStyle) | Out-Null }
            }

            if ($PSCmdlet.ParameterSetName -eq 'BottomRank')
            {

                if ($IsPercent) { $percentoritems = $true }
                elseif ($IsItems) { $percentoritems = $false }
                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsInTopRange($False, $BottomRankValue, $percentoritems, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsInTopRange($False, $BottomRankValue, $percentoritems, $SLStyle) | Out-Null }
            }

            ## // Blank Cells

            if ($PSCmdlet.ParameterSetName -eq 'Blank')
            {
                if ($HighlightBlankCells)
                {
                    if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsContainingBlanks($True, $PresetStyle) | Out-Null }
                    elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsContainingBlanks($True, $SLStyle) | Out-Null }
                }
            }

            if ($PSCmdlet.ParameterSetName -eq 'NonBlank')
            {
                if ($HighlightNonBlankCells)
                {
                    if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsContainingBlanks($False, $PresetStyle) | Out-Null }
                    elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsContainingBlanks($False, $SLStyle) | Out-Null }
                }
            }

            ## // Error Cells

            if ($PSCmdlet.ParameterSetName -eq 'Error')
            {
                if ($HighlightErrorCells)
                {
                    if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsContainingErrors($True, $PresetStyle) | Out-Null }
                    elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsContainingErrors($True, $SLStyle) | Out-Null }
                }
            }

            if ($PSCmdlet.ParameterSetName -eq 'NonError')
            {
                if ($HighlightNonErrorCells)
                {
                    if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsContainingErrors($False, $PresetStyle) | Out-Null }
                    elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsContainingErrors($False, $SLStyle) | Out-Null }
                }
            }

            ## // Equal to

            if ($PSCmdlet.ParameterSetName -eq 'EqualTo')
            {

                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsEqual($True, $EqualToValue, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsEqual($True, $EqualToValue, $SLStyle) | Out-Null }
            }

            if ($PSCmdlet.ParameterSetName -eq 'NotEqualTo')
            {

                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsEqual($False, $NotEqualToValue, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsEqual($False, $NotEqualToValue, $SLStyle) | Out-Null }
            }

            ## // Text that contains

            if ($PSCmdlet.ParameterSetName -eq 'TextContains')
            {

                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsContainingText($True, $TextContainsString, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsContainingText($True, $TextContainsString, $SLStyle) | Out-Null }
            }

            if ($PSCmdlet.ParameterSetName -eq 'TextDoesNotContain')
            {

                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsContainingText($False, $TextDoesNotContainString, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsContainingText($False, $TextDoesNotContainString, $SLStyle) | Out-Null }
            }

            ## // Text that ENDs with

            if ($PSCmdlet.ParameterSetName -eq 'TextENDsWith')
            {

                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsENDingWith($TextENDsWithString, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsENDingWith($TextENDsWithString, $SLStyle) | Out-Null }
            }

            ## // Text that BEGINs with

            if ($PSCmdlet.ParameterSetName -eq 'TextBEGINsWith')
            {

                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsBEGINningWith($TextBEGINsWithString, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsBEGINningWith($TextBEGINsWithString, $SLStyle) | Out-Null }
            }

            ## // Average

            if ($PSCmdlet.ParameterSetName -eq 'Average')
            {

                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsAboveAverage([SpreadsheetLight.SLHighlightCellsAboveAverageValues]::$AverageType, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsAboveAverage([SpreadsheetLight.SLHighlightCellsAboveAverageValues]::$AverageType, $SLStyle) | Out-Null }
            }

            ## // DUplicates

            if ($PSCmdlet.ParameterSetName -eq 'Duplicate')
            {
                if ($HighLightDuplicateValues)
                {
                    if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsWithDuplicates($PresetStyle) | Out-Null }
                    elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsWithDuplicates($SLStyle) | Out-Null }
                }
            }

            ## // Unique

            if ($PSCmdlet.ParameterSetName -eq 'Unique')
            {
                if ($HighLightUniqueValues)
                {
                    if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsWithUniques($PresetStyle) | Out-Null }
                    elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsWithUniques($SLStyle) | Out-Null }
                }
            }

            ## // Cells with Formula

            if ($PSCmdlet.ParameterSetName -eq 'Formula')
            {

                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsWithFormula($FormulaString, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsWithFormula($FormulaString, $SLStyle) | Out-Null }
            }

            ## // Dates

            if ($PSCmdlet.ParameterSetName -eq 'Date')
            {

                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsWithDatesOccurring($DateString, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsWithDatesOccurring($DateString, $SLStyle) | Out-Null }
            }

            Write-Verbose ("Set-SLConditionalFormatIconSet :`t Applying conditional formatting IconSet on Range '{0}'" -f $Range)
            $WorkBookInstance.AddConditionalFormatting($ConditionalFormatting) | Out-Null

            $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }#select-slworksheet

    }#process
    END
    {
    }

}
