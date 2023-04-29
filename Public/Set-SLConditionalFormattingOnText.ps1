Function Set-SLConditionalFormattingOnText  {


    <#

.SYNOPSIS
    Apply conditional formatting Iconset to text instead of numbers.

.DESCRIPTION
    Apply conditional formatting Iconset to text instead of numbers.
    Excel iconsets are applied on numbers and there is currently no built-in method to apply it on text or strings.
    This cmdlet takes a range containing text, inserts a new column before it and then applies conditional formatting on it.
    You can only apply text formatting on a column that has 3 or less unique values in a given column. Eg: "Working","Stopped","Disabled"

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    This is the name of the worksheet that contains the cell range where formatting is to be applied.

.PARAMETER Range
    The range of cells containing text to which conditional formatting has to be applied.

.PARAMETER IconSet
    Built-in Iconset styles.
    Use tab or intellisense to select from a range of possible values.
    Default value is - ThreeSymbols
    Possible values are:
    'ThreeArrows','ThreeArrowsGray','ThreeFlags','ThreeSigns','ThreeStars',
        'ThreeSymbols','ThreeSymbols2','ThreeTrafficLights1','ThreeTrafficLights2','ThreeTriangles'

.PARAMETER Properties
    String containing comma seperated text values. EG: "Working,Stopped,Disabled"

.PARAMETER IconColumnHeader
    The header text to be set for the new column contining icons. Default value is - "Icon"

.PARAMETER ReverseIconorder
    Reverses the order in which icons are applied.

.PARAMETER ShowIconsOnly
    Will show just the icons instead of icons and numbers.
    The default value is true.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormattingOnText -WorksheetName sheet1 -Range f4:f10 -Properties "working,stopped,disabled"
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    This will insert a column before column F and the conditional formatting Icon set 'ThreeSymbols' will be applied to the new column.
    The cell corresponding to value working will be 'Green', stopped will be 'yello\orange' and disabled in 'Red'.
    Note: Column F becomes column G.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $Services = Get-WmiObject -Class Win32_Service | Sort State,StartMode | Select __Server,Name,Displayname,State,StartMode
    PS C:\> $services | Export-SLDocument -WorkBookInstance $doc -WorksheetName sheet3 -AutofitColumns
    PS C:\> $doc | Save-SLDocument
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormattingOnText -WorksheetName sheet3 -Range e5:e187 -Properties "Running,Stopped"
    PS C:\> $doc | Save-SLDocument
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormattingOnText -WorksheetName sheet3 -Range g5:g187 -Properties "Auto,Manual,Disabled"
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    Servicedata is exported to sheet3, workbook is saved and closed. We then open the workbook to determine the data range for conditionalformatting.
    We apply conditional formatting on the state column which has 2 properties "Running" and "stopped" save and close.
    We open the document again to determine the datarange corresponding to the startmode column which has 3 properties "Auto","Manual" & "Disabled"



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
                $r1, $r2 = $_ -split ':'
                $r1_match = [regex]::Match($r1, '[a-zA-Z]+') | Select-Object -ExpandProperty value
                $r2_match = [regex]::Match($r2, '[a-zA-Z]+') | Select-Object -ExpandProperty value
                if ($r1_match -eq $r2_match) { $true }
                else { $false; Write-Warning "Set-SLConditionalFormattingOnText :`tVFormulacellRange should specify values that belong to the same column. Eg: A1:A10 or AB1:AB5"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true)]
        [string]$Range,

        [Validateset('ThreeArrows', 'ThreeArrowsGray', 'ThreeFlags', 'ThreeSigns', 'ThreeStars',
            'ThreeSymbols', 'ThreeSymbols2', 'ThreeTrafficLights1', 'ThreeTrafficLights2', 'ThreeTriangles')]
        [parameter(Mandatory = $false, Position = 2)]
        [string]$IconSet = 'ThreeSymbols',

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$Properties,

        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [string]$IconColumnHeader = 'Icon',

        [Switch]$ReverseIconorder = $true,
        [Switch]$ShowIconsOnly = $true





    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            $StartCellReference, $ENDCellReference = $Range -split ':'

            $RangeStats = Convert-ToExcelRowColumnStats -Range $Range
            $MatchedColumnIndex = $RangeStats.StartColumnIndex + 1
            $MatchedColumnName = Convert-ToExcelColumnName -Index $MatchedColumnIndex
            $MatchedcellReference = $MatchedColumnName + ($RangeStats.StartRowIndex + 1)

            $WorkBookInstance.InsertColumn($RangeStats.StartColumnIndex, 1) | Out-Null

            $NewColumnHeaderCellreference = $RangeStats.StartColumnName + $RangeStats.StartRowIndex
            $WorkBookInstance.SetCellValue("$NewColumnHeaderCellreference", $IconColumnHeader ) | Out-Null

            $NewIconColumnName = $RangeStats.StartColumnName
            $NewIconColumnIndex = $RangeStats.StartColumnIndex
            $NewIconRowIndex = $RangeStats.StartRowIndex + 1
            $endrowIndex = $RangeStats.EndRowIndex

            $Fproperties = (($Properties -split ',' | ForEach-Object { '"' + $_ + '"' } ) -join ',').ToString()

            for ($i = $NewIconRowIndex; $i -le $endrowIndex; $i++)
            {
                $MatchedcellReference = $MatchedColumnName + $NewIconRowIndex
                $WorkBookInstance.SetCellValue(($NewIconColumnName + $i), "=MATCH($MatchedcellReference,{$Fproperties},0)") | Out-Null
                $NewIconRowIndex++
            }


            $ConditionalFormatting = New-Object SpreadsheetLight.SLConditionalFormatting($startcellreference, $ENDcellreference)

            $ConditionalFormatting.SetCustomIconSet([SpreadsheetLight.SLThreeIconSetValues]::$IconSet, $ReverseIconorder, $ShowIconsOnly, $true, 2, [SpreadsheetLight.SLConditionalFormatRangeValues]::'Number', $true, 3, [SpreadsheetLight.SLConditionalFormatRangeValues]::'Number') | Out-Null

            Write-Verbose ("Set-SLConditionalFormatColorScale :`t Applying conditional formatting color scale on Range '{0}'" -f $Range)
            $WorkBookInstance.AddConditionalFormatting($ConditionalFormatting) | Out-Null

            $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }#select-slworksheet

    }#process

}
