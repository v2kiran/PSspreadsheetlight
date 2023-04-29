Function Set-SLConditionalFormattingDataBars  {


    <#

.SYNOPSIS
    Set conditional formatting data bars on a given range of cells.

.DESCRIPTION
    Set conditional formatting data bars on a given range of cells.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
        This is the name of the worksheet that contains the cell range where formatting is to be applied.

.PARAMETER Range
    The range of cells where conditional formatting has to be applied.

.PARAMETER DataBarColor
    to be used with the parameterset 'normal'.
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'Blue','Green','Red','Orange','LightBlue','Purple'

.PARAMETER ThemeColor
    to be used with the parameterset 'CustomDataBar1'.
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'Light1Color','Dark1Color','Light2Color','Dark2Color','Accent1Color','Accent2Color','Accent3Color','Accent4Color','Accent5Color','Accent6Color','Hyperlink','FollowedHyperlinkColor'

.PARAMETER DataBarMinLength
    Set the minimum length of the databar.

.PARAMETER DataBarMaxLength
    Set the maximum length of the databar.

.PARAMETER DataBarType1
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'Value','Number','Percent','Formula','Percentile'

.PARAMETER DataBarType2
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'Value','Number','Percent','Formula','Percentile'

.PARAMETER MinValue
    This is the minimum value from which the databar will begin.

.PARAMETER MaxValue
    This is the maximum value at which the databar will end.

.PARAMETER Color
    Color of the databar.Can be used in place of themecolor.

.PARAMETER ShowDataBarOnly
    If used only the databar sans value will be shown.



.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\Databars.xlsx
    PS C:\> $doc | Set-SLConditionalFormattingDataBars -WorksheetName sheet1 -Range e4:h6 -DataBarColor Green -Verbose
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    apply conditional formatting on range e4:h6 with the databar color chosen as green.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\Databars.xlsx
    PS C:\> $doc | Set-SLConditionalFormattingDataBars -WorksheetName sheet1 -Range e8:e10 -DataBarMinLength 0 -DataBarMaxLength 100 -DataBarType1 Number -MinValue 0 -DataBarType2 Value -MaxValue 100 -ThemeColor Accent3Color
    PS C:\> $doc | Save-SLDocument


    Description
    -----------
    Custom databar formatting applied with accent color3 as the databar color.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\Databars.xlsx
    PS C:\> $doc | Set-SLConditionalFormattingDataBars -WorksheetName sheet7 -Range f8:f10 -DataBarMinLength 0 -DataBarMaxLength 80 -DataBarType1 Number -MinValue 0 -DataBarType2 Value -MaxValue 100 -ThemeColor Accent4Color -ShowDataBarOnly
    PS C:\> $doc | Save-SLDocument


    Description
    -----------
    Same as the previous example but here we make 2 changes.
    1 - the maximum databar length is changed from 100 to 80 and
    2 - The values are hidden showing just the databars.

.Example
    PS C:\> Get-SLDocument D:\ps\excel\Databars.xlsx  |
                Set-SLFill -WorksheetName sheet7 -TargetCellorRange h12:h14 -Color Black |
                    Set-SLFont -FontColor White |
            Set-SLConditionalFormattingDataBars -DataBarMinLength 0 -DataBarMaxLength 80 -DataBarType1 Number -MinValue 0 -DataBarType2 Value -MaxValue 100 -ThemeColor Accent4Color |
                Set-SLColumnWidth -ColumnName h -ColumnWidth 20 |
                    Save-SLDocument


    Description
    -----------
    At times it may be difficult to see where the bars end, because of the graduated coloring in the data bars,
    so here we apply a dark fill color to the cells, and then change the font to a light color
    Also we change the width of the column to 20 which makes it a little easier to see the differences in the databar lengths.
    Note: since we are piping data between cmdlets we can ignore specifying the values for some of the parameters such as 'worksheetname' and 'Range'.
    However the best practise would be to specify parameter names so that there is no cause for confusion or ambiguity.



.INPUTS
   String,SpreadsheetLight.SLDocument

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    http://www.excel-easy.com/examples/data-bars.html

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
                else { $false; Write-Warning "Set-SLConditionalFormattingDataBars :`tRange should specify values in following format. Eg: A1:D10 or AB1:AD5"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true)]
        [string]$Range,

        [ValidateSet('Blue', 'Green', 'Red', 'Orange', 'LightBlue', 'Purple')]
        [parameter(Mandatory = $True, Position = 3, ParameterSetName = 'Normal')]
        [string]$DataBarColor,

        [ValidateSet('Light1Color', 'Dark1Color', 'Light2Color', 'Dark2Color', 'Accent1Color', 'Accent2Color', 'Accent3Color', 'Accent4Color', 'Accent5Color', 'Accent6Color', 'Hyperlink', 'FollowedHyperlinkColor')]
        [parameter(Mandatory = $True, ParameterSetName = 'CustomDataBar1')]
        [string]$ThemeColor,

        [parameter(Mandatory = $True, ParameterSetName = 'CustomDataBar1')]
        [parameter(Mandatory = $True, ParameterSetName = 'CustomDataBar2')]
        [int]$DataBarMinLength,

        [parameter(Mandatory = $True, ParameterSetName = 'CustomDataBar1')]
        [parameter(Mandatory = $True, ParameterSetName = 'CustomDataBar2')]
        [int]$DataBarMaxLength,

        [ValidateSet('Value', 'Number', 'Percent', 'Formula', 'Percentile')]
        [parameter(Mandatory = $True, ParameterSetName = 'CustomDataBar1')]
        [parameter(Mandatory = $True, ParameterSetName = 'CustomDataBar2')]
        [string]$DataBarType1,

        [ValidateSet('Value', 'Number', 'Percent', 'Formula', 'Percentile')]
        [parameter(Mandatory = $True, ParameterSetName = 'CustomDataBar1')]
        [parameter(Mandatory = $True, ParameterSetName = 'CustomDataBar2')]
        [string]$DataBarType2,

        [parameter(Mandatory = $True, ParameterSetName = 'CustomDataBar1')]
        [parameter(Mandatory = $True, ParameterSetName = 'CustomDataBar2')]
        $MinValue,

        [parameter(Mandatory = $True, ParameterSetName = 'CustomDataBar1')]
        [parameter(Mandatory = $True, ParameterSetName = 'CustomDataBar2')]
        $MaxValue,

        [parameter(Mandatory = $True, ParameterSetName = 'CustomDataBar2')]
        [string]$Color,

        [parameter(Mandatory = $False, ParameterSetName = 'CustomDataBar1')]
        [parameter(Mandatory = $False, ParameterSetName = 'CustomDataBar2')]
        [Switch]$ShowDataBarOnly = $false
    )
    PROCESS
    {

        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {

            $startcellreference, $ENDcellreference = $range -split ':'
            $ConditionalFormatting = New-Object SpreadsheetLight.SLConditionalFormatting($startcellreference, $ENDcellreference)

            if ($PSCmdlet.ParameterSetName -eq 'Normal')
            {
                Write-Verbose ("Set-SLConditionalFormattingDataBars :`t Databar color is '{0}'" -f $DataBarColor)
                $ConditionalFormatting.SetDataBar([SpreadsheetLight.SLConditionalFormatDataBarValues]::$DataBarColor) | Out-Null
            }
            if ($PSCmdlet.ParameterSetName -eq 'CustomDataBar1')
            {
                #Write-Verbose ("Set-SLConditionalFormattingDataBars :`tData Range '{0}'. DataBarMinLength is '{1}'" -f $Range,$DataBarColor)
                $ConditionalFormatting.SetCustomDataBar($ShowDataBarOnly, $DataBarMinLength, $DataBarMaxLength, [SpreadsheetLight.SLConditionalFormatMinMaxValues]::$DataBarType1, $MinValue, [SpreadsheetLight.SLConditionalFormatMinMaxValues]::$DataBarType2, $MaxValue, [SpreadsheetLight.SLThemeColorIndexValues]::$ThemeColor) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'CustomDataBar2')
            {
                $ConditionalFormatting.SetCustomDataBar($ShowDataBarOnly, $DataBarMinLength, $DataBarMaxLength, [SpreadsheetLight.SLConditionalFormatMinMaxValues]::$DataBarType1, $MinValue, [SpreadsheetLight.SLConditionalFormatMinMaxValues]::$DataBarType2, $MaxValue, [System.Drawing.Color]::$Color) | Out-Null
            }


            Write-Verbose ("Set-SLConditionalFormattingDataBars :`tSetting conditional formatting on range '{0}'" -f $Range)
            $WorkBookInstance.AddConditionalFormatting($ConditionalFormatting) | Out-Null

            $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }#select-slworksheet

    }#process
    END
    {


    }

}
