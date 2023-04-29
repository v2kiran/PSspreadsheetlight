Function Set-SLBuiltinCellStyle  {


    <#

.SYNOPSIS
    Apply a style based on the built-in cellstyles.

.DESCRIPTION
    Apply a style based on the built-in cellstyles.
    Applying a cell style will replace any existing cell formatting except for text alignment.
    You may not want to use cell styles if you've added custom formatting to a cell or cells.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER CellReference
    The target cell that needs to have the specified cellstyle. Eg: A5 or AB10

.PARAMETER Range
    The target cell range that needs to have the specified cellstyle. Eg: A5:B10 or AB10:AD20

.PARAMETER CellStyle
    Use tab completion or intellisense to select a possible value from a list provided by the parameter.
    'Normal','Bad','Good','Neutral','Calculation','CheckCell','ExplanatoryText','Input','LinkedCell','Note','Output','WarningText',
        'Heading1','Heading2','Heading3','Heading4','Title','Total','Accent1','Accent2','Accent3','Accent4','Accent5','Accent6',
        'Accent1Percentage60','Accent2Percentage60','Accent3Percentage60','Accent4Percentage60','Accent5Percentage60','Accent6Percentage60',
        'Accent1Percentage40','Accent2Percentage40','Accent3Percentage40','Accent4Percentage40','Accent5Percentage40','Accent6Percentage40',
        'Accent1Percentage20','Accent2Percentage20','Accent3Percentage20','Accent4Percentage20','Accent5Percentage20','Accent6Percentage20',
        'Comma','Comma0','Currency','Currency0','Percentage'


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLBuiltinCellStyle -WorksheetName sheet1 -CellReference B6,C6,D6 -CellStyle Accent2 -Verbose | Save-SLDocument

    Description
    -----------
    Apply a cellstyle anmed 'Accent2' to cells B6,C6,D6.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLBuiltinCellStyle -WorksheetName sheet1 -Range G5:L5 -CellStyle Accent3 -Verbose
    PS C:\> $doc | Set-SLBuiltinCellStyle -WorksheetName sheet1 -Range G6:G7 -CellStyle Accent3Percentage60 -Verbose
    PS C:\> $doc | Set-SLBuiltinCellStyle -WorksheetName sheet1 -Range H6:L7 -CellStyle Accent3Percentage40 -Verbose
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    Apply different cellstyles to a set of cell ranges.
    Note: save-sldocument is called in the last step i.e., after we apply all styles the final step is to save the document.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx |
                Set-SLBuiltinCellStyle -WorksheetName sheet1 -Range G5:L5 -CellStyle Accent3 -Verbose |
                    Set-SLBuiltinCellStyle  -Range G6:G7 -CellStyle Accent3Percentage60 -Verbose |
                        Set-SLBuiltinCellStyle  -Range H6:L7 -CellStyle Accent3Percentage40 -Verbose |
                            Save-SLDocument

    Description
    -----------
    Same as the previous example except that here we use the pipe to apply various styles.
    Note: This is more efficient because it avoids the additional step of assigning the instance to a variable and then piping that variable to apply the style.


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
                else { $false; Write-Warning "Set-SLBuiltinCellStyle :`tCellReference should specify values in following format. Eg: A1,B10,AB5..etc"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true, ParameterSetname = 'cell')]
        [string[]]$CellReference,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+:[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Set-SLBuiltinCellStyle :`tRange should specify values in following format. Eg: A1:D10 or AB1:AD5"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true, ParameterSetname = 'Range')]
        [string]$Range,

        [ValidateSet('Normal', 'Bad', 'Good', 'Neutral', 'Calculation', 'CheckCell', 'ExplanatoryText', 'Input', 'LinkedCell', 'Note', 'Output', 'WarningText',
            'Heading1', 'Heading2', 'Heading3', 'Heading4', 'Title', 'Total', 'Accent1', 'Accent2', 'Accent3', 'Accent4', 'Accent5', 'Accent6',
            'Accent1Percentage60', 'Accent2Percentage60', 'Accent3Percentage60', 'Accent4Percentage60', 'Accent5Percentage60', 'Accent6Percentage60',
            'Accent1Percentage40', 'Accent2Percentage40', 'Accent3Percentage40', 'Accent4Percentage40', 'Accent5Percentage40', 'Accent6Percentage40',
            'Accent1Percentage20', 'Accent2Percentage20', 'Accent3Percentage20', 'Accent4Percentage20', 'Accent5Percentage20', 'Accent6Percentage20',
            'Comma', 'Comma0', 'Currency', 'Currency0', 'Percentage')]
        [parameter(Mandatory = $true, Position = 3, ValueFromPipelineByPropertyName = $true)]
        [string]$CellStyle

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {

            if ($PSCmdlet.ParameterSetName -eq 'cell')
            {
                Foreach ($cref in $CellReference)
                {
                    Try
                    {
                        Write-Verbose ("Set-SLBuiltinCellStyle :`tSetting Built-In CellStyle '{0}' on Cell '{1}'" -f $CellStyle, $cref)
                        $WorkBookInstance.ApplyNamedCellStyle($cref, [SpreadsheetLight.SLNamedCellStyleValues]::$CellStyle) | Out-Null
                    }
                    Catch
                    {
                        Write-Warning ("Set-SLBuiltinCellStyle :`tPlease check if the specified cellstyle is available on the version of excel installed ...'{0}'" -f $CellStyle)
                    }
                }
                $WorkBookInstance | Add-Member NoteProperty CellReference $CellReference -Force
            }
            if ($PSCmdlet.ParameterSetName -eq 'Range')
            {
                $rowindex, $columnindex = $range -split ':'
                Try
                {
                    Write-Verbose ("Set-SLBuiltinCellStyle :`tSetting Built-In CellStyle '{0}' on Range '{1}'" -f $CellStyle, $Range)
                    $WorkBookInstance.ApplyNamedCellStyle($rowindex, $columnindex, [SpreadsheetLight.SLNamedCellStyleValues]::$CellStyle) | Out-Null
                }
                Catch
                {
                    Write-Warning ("Set-SLBuiltinCellStyle :`tPlease check if the specified cellstyle is available on the version of excel installed ...'{0}'" -f $CellStyle)
                }
                $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }

    }#Process
    END
    {

    }

}
