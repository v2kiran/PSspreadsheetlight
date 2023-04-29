Function Set-SLCellValue  {


    <#

.SYNOPSIS
    Set a Cell value on a single or a range of cells.

.DESCRIPTION
    Set a Cell value on a single or a range of cells.
    Note: you can only set the same value on multiple cells.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER CellReference
    The target cell whose value has to be set. Eg: A5 or AB10

.PARAMETER Value
    The value to be set.

.PARAMETER Range
    The target cell range that needs to have the specified value. Eg: A5:B10 or AB10:AD20


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLCellValue -CellReference B5,C7  -value "Hello" -Verbose | Save-SLDocument

    Description
    -----------
    Set the value of cells B5 & C7 to "Hello"


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | New-SLFontStyle  -WorksheetName sheet1 -FontName arial -FontSize 10 -FontColor Blue -IsBold  -IsItalic -IsStrikenThrough -Verbose
    PS C:\> $doc | New-SLRichTextStyle  -WorksheetName sheet1 -Text Hello
    PS C:\> $doc | New-SLFontStyle  -WorksheetName sheet1 -FontName arial -FontSize 12 -FontColor red -IsBold -Verbose
    PS C:\> $doc | New-SLRichTextStyle  -WorksheetName sheet1 -Text World -Append
    PS C:\> $doc | Set-SLCellValue -WorksheetName sheet1 -CellReference B6 -SetRichTextStyle -Verbose
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    Set the string Hello Worls as rich text in cell B6.


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
        [parameter(Mandatory = $true, position = 0, ValueFromPipeLine = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [Alias('CurrentWorkSheetName')]
        [parameter(Mandatory = $true, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true)]
        [String[]]$CellReference,

        [parameter(Mandatory = $true, Position = 3, ValueFromPipeLineByPropertyName = $true, Parametersetname = 'Value')]
        $value,

        [parameter(Parametersetname = 'RichText')]
        [Switch]$SetRichTextStyle


    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'Singlecell')
            {
                Foreach ($cref in $CellReference)
                {
                    Write-Verbose ("Set-SLCellValue :`tSetting Cell Value '{0}' on Cell '{1}'" -f $Value, $cref)
                    $WorkBookInstance.SetCellValue($Cref, $value) | Out-Null
                }
            }

            if ($PSCmdlet.ParameterSetName -eq 'RichText')
            {
                If ($WorkBookInstance.RichTextStyle)
                {
                    Foreach ($cref in $CellReference)
                    {
                        Write-Verbose ("Set-SLCellValue :`tSetting RichText Style on Cell '{0}'" -f $cref)
                        $WorkBookInstance.SetCellValue($Cref, $WorkBookInstance.RichTextStyle.ToInlineString()) | Out-Null
                    }
                }
                Else
                {
                    Write-Warning ("Set-SLCellValue :`tUse the New-SLFontStyle & New-SLRichTextStyle cmdlets to create font/richtext styles and then apply that style on a cellreference")
                }
            }

            $WorkBookInstance | Add-Member NoteProperty CellReference $CellReference -Force
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force
            Write-Output $WorkBookInstance
        }


    }#process
    END
    {

    }

}
