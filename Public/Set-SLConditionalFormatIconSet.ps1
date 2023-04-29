Function Set-SLConditionalFormatIconSet  {


    <#

.SYNOPSIS
    Apply conditional formatting Iconset on numbers.

.DESCRIPTION
    Apply conditional formatting Iconset on numbers.
    Based on the data users may select  3, 4 or 5iconsets to display data.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    This is the name of the worksheet that contains the cell range where formatting is to be applied.

.PARAMETER Range
    The range of cells containing text to which conditional formatting has to be applied.

.PARAMETER IconSet
    Built-in Iconset styles.
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'FiveArrows','FiveArrowsGray','FiveQuarters','FiveRating','FourArrows',
        'FourArrowsGray','FourRating','FourRedToBlack','FourTrafficLights',
        'ThreeArrows','ThreeArrowsGray','ThreeFlags','ThreeSigns',
        'ThreeSymbols','ThreeSymbols2','ThreeTrafficLights1','ThreeTrafficLights2'

.PARAMETER FiveIconSetType
    Use this to apply different formatting types on 5 different ranges.

.PARAMETER FourIconSetType
    Use this to apply different formatting types on 4 different ranges.

.PARAMETER ThreeIconSetType
    Use this to apply different formatting types on 4 different ranges.

.PARAMETER ReverseIconOrder
   Reverse the order of the icons displayed.

.PARAMETER ShowIconsOnly
    Will show just the icons instead of icons and numbers.


.PARAMETER GreaterThanOrEqual2
    True if values are to be greater than or equal to the 2nd range value.False if values are to be strictly greater than.

.PARAMETER SecondRangeValue
    The 2nd Range value.

.PARAMETER SecondRangeValueType
    Built-in Iconset format types.
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'Number','Percent','Formula','Percentile'


.PARAMETER ThirdRangeValueType
    Built-in Iconset format types.
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'Number','Percent','Formula','Percentile'

.PARAMETER FourthRangeValueType
    Built-in Iconset format types.
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'Number','Percent','Formula','Percentile'


.PARAMETER FifthRangeValueType
    Built-in Iconset format types.
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'Number','Percent','Formula','Percentile'

.PARAMETER GreaterThanOrEqual3
    True if values are to be greater than or equal to the 3rd range value.False if values are to be strictly greater than.

.PARAMETER ThirdRangeValue
    The 3rd range value.

.PARAMETER GreaterThanOrEqual4
    True if values are to be greater than or equal to the 4th range value.False if values are to be strictly greater than.

.PARAMETER FourthRangeValue
    The 4th range value.
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'Number','Percent','Formula','Percentile'

.PARAMETER GreaterThanOrEqual5
    True if values are to be greater than or equal to the 5th range value.False if values are to be strictly greater than.

.PARAMETER FifthRangeValue
    The 5th range value.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatIconSet -WorksheetName sheet7 -Range d4:d15 -IconSet ThreeSymbols -Verbose
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    Apply the conditional formatting Icon set 'ThreeSymbols' to the range d4:d15. Both icons and values are shown


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatIconSet -WorksheetName sheet7 -Range f4:f15 -FiveIconSetType FiveRating -ReverseIconOrder:$false -ShowIconOnly -GreaterThanOrEqual2 -SecondRangeValue 15 -SecondRangeValueType Percentile -GreaterThanOrEqual3 -ThirdRangeValue 35 -ThirdRangeValueType Percentile -GreaterThanOrEqual4 -FourthRangeValue 67 -FourthRangeValueType Percentile -GreaterThanOrEqual5 -FifthRangeValue 80 -FifthRangeValueType Percentile
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    Apply the conditional formatting Icon set 'FiveRating' to the range f4:f15. Only icons are shown.

.Example
    PS C:\> $IconSet5Params = @{

        WorkBookInstance = ($doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx)
        WorksheetName = 'sheet7'
        Range = 'f4:f15'
        FiveIconSetType = 'FiveRating'
        ReverseIconOrder = $false
        ShowIconOnly = $true
        GreaterThanOrEqual2 = $true
        SecondRangeValue = 15
        SecondRangeValueType = 'Percentile'
        GreaterThanOrEqual3 = $true
        ThirdRangeValue = 35
        ThirdRangeValueType = 'Percentile'
        GreaterThanOrEqual4 = $true
        FourthRangeValue = 67
        FourthRangeValueType = 'Percentile'
        GreaterThanOrEqual5 = $true
        FifthRangeValue = 80
        FifthRangeValueType = 'Percentile'
        Verbose = $true
}

    PS C:\> Set-SLConditionalFormatIconSet @IconSet5Params
    PS C:\> $doc | Save-SLDocument


    Description
    -----------
    Since the last example had a lot of parameters and values that scrolled off to the right this example will retain the same values\parameters
    but will use a different format of data input to the cmdlet which will make it easier to read.
    All the parameters and values required to run the cmdlet Set-SLConditionalFormatIconSet are stored in the variable - IconSet5Params
    which is a hashtable that contains Key\Value pairs.
    The keys are the parameters and the values are the parameter values.
    Note: You’ll notice a little trick here. The “@” sign is followed by the variable name "IconSet5Params", which doesn’t include the dollar sign.
    The “@” sign, when used as a splat operator says,
    “Take whatever characters come next and assume they’re a variable name. Assume that the variable contains a hashtable, and that the keys are parameter names.
    The above explanation of the @ 'splat' operator is a direct quote from Don jones :)



.Example

    PS C:\> $IconSet3Params = @{

        WorkBookInstance = (Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx)
        WorksheetName = 'sheet7'
        Range = 'h4:h15'
        ThreeIconSetType = 'ThreeTrafficLights1'
        ReverseIconOrder = $false
        ShowIconOnly = $false
        GreaterThanOrEqual2 = $true
        SecondRangeValue = 33
        SecondRangeValueType = 'Number'
        GreaterThanOrEqual3 = $true
        ThirdRangeValue = 82
        ThirdRangeValueType = 'Number'
        Verbose = $true
}


    PS C:\> $IconSet3ReverseIconsParams = @{

        WorkBookInstance = (Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx)
        WorksheetName = 'sheet7'
        Range = 'j4:j15'
        ThreeIconSetType = 'ThreeTrafficLights1'
        ReverseIconOrder = $true
        ShowIconOnly = $true
        GreaterThanOrEqual2 = $true
        SecondRangeValue = 33
        SecondRangeValueType = 'Number'
        GreaterThanOrEqual3 = $true
        ThirdRangeValue = 82
        ThirdRangeValueType = 'Number'
        Verbose = $true
}

    PS C:\> Set-SLConditionalFormatIconSet @IconSet3Params | Save-SLDocument
    PS C:\> Set-SLConditionalFormatIconSet @IconSet3ReverseIconsParams | Save-SLDocument


    Description
    -----------
    Here we apply conditional formatting twice on two different ranges. h4:h15 & J4:J15.
    The only difference between the 2 is that the second range J4;J15 has the icon order reversed.


.INPUTS
   String,Int,Bool,SpreadsheetLight.SLDocument

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
                else { $false; Write-Warning "Set-SLConditionalFormatIconSet :`tRange should specify values in following format. Eg: A1:D10 or AB1:AD5"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true)]
        [string]$Range,

        [Validateset('FiveArrows', 'FiveArrowsGray', 'FiveQuarters', 'FiveRating', 'FourArrows',
            'FourArrowsGray', 'FourRating', 'FourRedToBlack', 'FourTrafficLights',
            'ThreeArrows', 'ThreeArrowsGray', 'ThreeFlags', 'ThreeSigns',
            'ThreeSymbols', 'ThreeSymbols2', 'ThreeTrafficLights1', 'ThreeTrafficLights2')]
        [parameter(Mandatory = $true, Position = 3, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'IconSet')]
        [string]$IconSet,

        [Validateset('FiveArrows', 'FiveArrowsGray', 'FiveBoxes', 'FiveQuarters', 'FiveRating')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'CustomIconSet5')]
        [string]$FiveIconSetType,

        [Validateset('FourArrows', 'FourArrowsGray', 'FourRating', 'FourRedToBlack', 'FourTrafficLights')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'CustomIconSet4')]
        [string]$FourIconSetType,

        [Validateset('ThreeArrows', 'ThreeArrowsGray', 'ThreeFlags', 'ThreeSigns', 'ThreeStars', 'ThreeSymbols', 'ThreeSymbols2', 'ThreeTrafficLights1', 'ThreeTrafficLights2', 'ThreeTriangles')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'CustomIconSet3')]
        [string]$ThreeIconSetType,

        [parameter(ParameterSetName = 'CustomIconSet3')]
        [parameter(ParameterSetName = 'CustomIconSet4')]
        [parameter(ParameterSetName = 'CustomIconSet5')]
        [switch]$ReverseIconOrder,

        [parameter(ParameterSetName = 'CustomIconSet3')]
        [parameter(ParameterSetName = 'CustomIconSet4')]
        [parameter(ParameterSetName = 'CustomIconSet5')]
        [switch]$ShowIconOnly,

        [parameter(ParameterSetName = 'CustomIconSet3')]
        [parameter(ParameterSetName = 'CustomIconSet4')]
        [parameter(ParameterSetName = 'CustomIconSet5')]
        [switch]$GreaterThanOrEqual2,

        [parameter(ParameterSetName = 'CustomIconSet3')]
        [parameter(ParameterSetName = 'CustomIconSet4')]
        [parameter(ParameterSetName = 'CustomIconSet5')]
        [string]$SecondRangeValue,

        [Validateset('Number', 'Percent', 'Formula', 'Percentile')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'CustomIconSet3')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'CustomIconSet4')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'CustomIconSet5')]
        [string]$SecondRangeValueType,

        [Validateset('Number', 'Percent', 'Formula', 'Percentile')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'CustomIconSet3')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'CustomIconSet4')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'CustomIconSet5')]
        [string]$ThirdRangeValueType,

        [Validateset('Number', 'Percent', 'Formula', 'Percentile')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'CustomIconSet4')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'CustomIconSet5')]
        [string]$FourthRangeValueType,

        [Validateset('Number', 'Percent', 'Formula', 'Percentile')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'CustomIconSet5')]
        [string]$FifthRangeValueType,

        [parameter(ParameterSetName = 'CustomIconSet3')]
        [parameter(ParameterSetName = 'CustomIconSet4')]
        [parameter(ParameterSetName = 'CustomIconSet5')]
        [switch]$GreaterThanOrEqual3,

        [parameter(ParameterSetName = 'CustomIconSet3')]
        [parameter(ParameterSetName = 'CustomIconSet4')]
        [parameter(ParameterSetName = 'CustomIconSet5')]
        [string]$ThirdRangeValue,

        [parameter(ParameterSetName = 'CustomIconSet4')]
        [parameter(ParameterSetName = 'CustomIconSet5')]
        [switch]$GreaterThanOrEqual4,

        [parameter(ParameterSetName = 'CustomIconSet4')]
        [parameter(ParameterSetName = 'CustomIconSet5')]
        [string]$FourthRangeValue,

        [parameter(ParameterSetName = 'CustomIconSet5')]
        [switch]$GreaterThanOrEqual5,

        [parameter(ParameterSetName = 'CustomIconSet5')]
        [string]$FifthRangeValue


    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {

            $startcellreference, $ENDcellreference = $range -split ':'
            $ConditionalFormatting = New-Object SpreadsheetLight.SLConditionalFormatting($startcellreference, $ENDcellreference)

            if ($PSCmdlet.ParameterSetName -eq 'IconSet')
            {
                Write-Verbose ("Set-SLConditionalFormatIconSet :`t Selected Iconset is '{0}'" -f $IconSet)
                $ConditionalFormatting.SetIconSet([OLIconsetvalues]::$IconSet) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'CustomIconSet5')
            {
                Write-Verbose ("Set-SLConditionalFormatIconSet :`t Selected Iconset is '{0}'" -f $FiveIconSetType)
                $ConditionalFormatting.SetCustomIconSet([SpreadsheetLight.SLFiveIconSetValues]::$FiveIconSetType, $ReverseIconOrder, $ShowIconOnly, $GreaterThanOrEqual2, $SecondRangeValue, [SLCFRangeValues]::$SecondRangeValueType, $GreaterThanOrEqual3, $ThirdRangeValue, [SLCFRangeValues]::$ThirdRangeValueType, $GreaterThanOrEqual4, $FourthRangeValue, [SLCFRangeValues]::$FourthRangeValueType, $GreaterThanOrEqual5, $FifthRangeValue, [SLCFRangeValues]::$FifthRangeValueType ) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'CustomIconSet4')
            {
                Write-Verbose ("Set-SLConditionalFormatIconSet :`t Selected Iconset is '{0}'" -f $FourIconSetType)
                $ConditionalFormatting.SetCustomIconSet([SpreadsheetLight.SLFourIconSetValues]::$FourIconSetType, $ReverseIconOrder, $ShowIconOnly, $GreaterThanOrEqual2, $SecondRangeValue, [SLCFRangeValues]::$SecondRangeValueType, $GreaterThanOrEqual3, $ThirdRangeValue, [SLCFRangeValues]::$ThirdRangeValueType, $GreaterThanOrEqual4, $FourthRangeValue, [SLCFRangeValues]::$FourthRangeValueType ) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'CustomIconSet3')
            {
                Write-Verbose ("Set-SLConditionalFormatIconSet :`t Selected Iconset is '{0}'" -f $ThreeIconSetType)
                $ConditionalFormatting.SetCustomIconSet([SpreadsheetLight.SLThreeIconSetValues]::$ThreeIconSetType, $ReverseIconOrder, $ShowIconOnly, $GreaterThanOrEqual2, $SecondRangeValue, [SLCFRangeValues]::$SecondRangeValueType, $GreaterThanOrEqual3, $ThirdRangeValue, [SLCFRangeValues]::$ThirdRangeValueType ) | Out-Null
            }


            Write-Verbose ("Set-SLConditionalFormatIconSet :`t Applying conditional formatting IconSet on Range '{0}'" -f $Range)
            $WorkBookInstance.AddConditionalFormatting($ConditionalFormatting) | Out-Null

            $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru


        }
    }


}
