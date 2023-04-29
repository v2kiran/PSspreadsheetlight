Function Set-SLWorksheetTabColor  {


    <#

.SYNOPSIS
    Sets the tab color of a worksheet.

.DESCRIPTION
    Sets the tab color of a worksheet.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER Color
    Any color that is to be used eg: Red.

.PARAMETER ThemeColor
    Theme color to be used. Valid values are:
    'Light1Color','Dark1Color','Light2Color','Dark2Color','Accent1Color','Accent2Color','Accent3Color',
    'Accent4Color','Accent5Color','Accent6Color','Hyperlink','FollowedHyperlinkColor'

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Set-SLWorksheetTabColor -WorksheetName sheet2 -Color Yellow   -Verbose  | Save-SLDocument


    Description
    -----------
    Set the tab color of sheet2 to yellow.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Set-SLWorksheetTabColor -WorksheetName sheet2 -ThemeColor Accent2Color   -Verbose  | Save-SLDocument



    Description
    -----------
    Set the tab color of sheet2 to Accent2Color.


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

        [parameter(Mandatory = $True, Position = 2, ParameterSetName = 'Color')]
        [string]$Color,

        [ValidateSet('Light1Color', 'Dark1Color', 'Light2Color', 'Dark2Color', 'Accent1Color', 'Accent2Color', 'Accent3Color', 'Accent4Color', 'Accent5Color', 'Accent6Color', 'Hyperlink', 'FollowedHyperlinkColor')]
        [parameter(Mandatory = $True, Position = 2, ParameterSetName = 'ThemeColor')]
        [string]$ThemeColor

    )
    PROCESS
    {

        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            $PageSettings = $WorkBookInstance.GetPageSettings()

            if ($PSCmdlet.ParameterSetName -eq 'Color')
            {
                Write-Verbose ("Set-SLWorksheetTabColor :`tSet worksheet '{0}' tab color to '{1}'" -f $WorksheetName, $Color)
                $PageSettings.TabColor = [System.Drawing.Color]::$color
            }
            if ($PSCmdlet.ParameterSetName -eq 'ThemeColor')
            {
                Write-Verbose ("Set-SLWorksheetTabColor :`tSet worksheet '{0}' tab color to '{1}'" -f $WorksheetName, $ThemeColor)
                $PageSettings.SetTabColor([SpreadsheetLight.SLThemeColorIndexValues]::$ThemeColor)
            }

            $WorkBookInstance.SetPageSettings($PageSettings) | Out-Null
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }

    }#process
    END
    {
    }

}
