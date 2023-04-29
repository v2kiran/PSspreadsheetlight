Function Copy-SLCellValue  {


    <#

.SYNOPSIS
    Copy a single or a range of cell values.

.DESCRIPTION
    Copy a single or a range of cell values. Source data can be on a worksheet that is different than the target.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER FromWorksheetName
    The worksheet containing the source data.

.PARAMETER FromCellReference
    the source cell containing the data to be copied Eg. A3

.PARAMETER Range
    The source data range to be copied Eg. A1:C3

.PARAMETER ToCellreference
    the target cell where data is to be copied to Eg. A3

.PARAMETER ToAnchorCellreference
    The cell reference of the target anchor cell, such as "A1".

.PARAMETER CutorCopy
    Specify whether data is to be copied or pasted

.PARAMETER PasteSpecial
    Specift special paste options such as:
    'Formatting','Formulas','Paste','Values','Transpose'


.Example
    PS C:\> Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx | Copy-SLCellValue -WorksheetName sheet4 -FromCellReference B2 -ToCellreference C2 -Verbose | Save-SLDocument

    Description
    -----------
    Copy cell B2 to C2.


.Example
    PS C:\> Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx |
                Copy-SLCellValue -WorksheetName sheet4 -FromCellReference B2 -ToCellreference D2 -PasteSpecial Formatting -Verbose |
                    Save-SLDocument

    Description
    -----------
    copy only formatting settings from B2 to D2.


.Example
    PS C:\> Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx |
                Copy-SLCellValue -WorksheetName sheet4 -FromCellReference C3 -ToCellreference F2 -CutorCopy Cut -Verbose |
                    Save-SLDocument

    Description
    -----------
    Cut cell C3 and paste it to F2.


.Example
    PS C:\> Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx |
                Copy-SLCellValue -WorksheetName sheet6 -FromWorksheetName sheet4 -FromCellReference B2 -ToCellreference E2 -PasteSpecial Values -Verbose |
                    Save-SLDocument

    Description
    -----------
    Copy B2 from sheet4 and paste it to E2 on sheet6

.Example
    PS C:\> Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx | Copy-SLCellValue -WorksheetName sheet4 -Range A9:C15 -ToAnchorCellreference E9 -Verbose  | Save-SLDocument


    Description
    -----------
    Copy range A9:C15 to E9

.Example
    PS C:\> Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx |
                Copy-SLCellValue -WorksheetName sheet6 -FromWorksheetName sheet4 -Range A9:A15 -ToAnchorCellreference J9 -PasteSpecial  Values -Verbose |
                    Save-SLDocument

    Description
    -----------
    Copy range A9:A15 from sheet4 and paste only the values (ignore any style settings applied to the range) to anchor cell J9 on sheet6.

.Example
    PS C:\> Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx |
                Copy-SLCellValue -WorksheetName sheet6 -FromWorksheetName sheet4 -Range A9:C9 -ToAnchorCellreference N9 -PasteSpecial  Transpose -Verbose |
                    Save-SLDocument

    Description
    -----------
    Copy range A9:A15 from sheet4 and transpose the values (convert row to column and viceversa) to anchor cell N9 on sheet6.

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

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Range-PasteSpecial-DifferentWorksheet')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Range-SimplyCopyPaste-DifferentWorksheet')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleCell-PasteSpecial-DifferentWorksheet')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleCell-SimplyCopyPaste-DifferentWorksheet')]
        [string]$FromWorksheetName,


        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Copy-SLCellValue :`tCellReference should specify values in following format. Eg: A1,B10,AB5..etc"; break }
            })]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleCell-PasteSpecial-DifferentWorksheet')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleCell-SimplyCopyPaste-DifferentWorksheet')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleCell-CutOrCopy')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleCell-PasteSpecial')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleCell-SimplyCopyPaste')]
        [string]$FromCellReference,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+:[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Copy-SLCellValue :`tRange should specify values in following format. Eg: A1:D10 or AB1:AD5"; break }
            })]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Range-PasteSpecial-DifferentWorksheet')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Range-SimplyCopyPaste-DifferentWorksheet')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Range-CutOrCopy')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Range-PasteSpecial')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Range-SimplyCopyPaste')]
        [string]$Range,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Copy-SLCellValue :`tCellReference should specify values in following format. Eg: A1,B10,AB5..etc"; break }
            })]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleCell-PasteSpecial-DifferentWorksheet')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleCell-SimplyCopyPaste-DifferentWorksheet')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleCell-CutOrCopy')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleCell-PasteSpecial')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleCell-SimplyCopyPaste')]
        [String]$ToCellreference,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Copy-SLCellValue :`tCellReference should specify values in following format. Eg: A1,B10,AB5..etc"; break }
            })]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Range-PasteSpecial-DifferentWorksheet')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Range-SimplyCopyPaste-DifferentWorksheet')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Range-CutOrCopy')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Range-PasteSpecial')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Range-SimplyCopyPaste')]
        [String]$ToAnchorCellreference,


        [ValidateSet('Cut', 'Copy')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Range-CutOrCopy')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleCell-CutOrCopy')]
        [String]$CutorCopy,

        [ValidateSet('Formatting', 'Formulas', 'Paste', 'Values', 'Transpose')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Range-PasteSpecial-DifferentWorksheet')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Range-PasteSpecial')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleCell-PasteSpecial-DifferentWorksheet')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleCell-PasteSpecial')]
        [String]$PasteSpecial

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            # paste - just plain paste - Retains Font, Alignment,Fill etc but not data validation
            # values - no formulas or formatting
            # Formulas - Paste values and formulas. Cell references are re-calculated
            # Formatting - only formatting no values - Retains Font, Alignment,Fill etc

            ### - #  SINGLECELL - SameWorksheet
            if ($PSCmdlet.ParameterSetName -eq 'SingleCell-SimplyCopyPaste')
            {
                Write-Verbose ("Copy-SLCellValue :`tCopy cell '{0}' to cell '{1}'" -f $FromCellReference, $ToCellreference)
                # copy one cell to another - Retains Font, Alignment,Fill etc but not data validation
                $WorkBookInstance.CopyCell($FromCellReference, $ToCellreference) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'SingleCell-PasteSpecial')
            {
                Write-Verbose ("Copy-SLCellValue :`tCopy cell '{0}' to cell '{1}' with PasteSpecial Option '{2}'" -f $FromCellReference, $ToCellreference, $PasteSpecial)
                # copy one cell to another with paste option
                $WorkBookInstance.CopyCell($FromCellReference, $ToCellreference, [SpreadsheetLight.SLPasteTypeValues]::$PasteSpecial) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'SingleCell-CutOrCopy')
            {
                if ($CutorCopy -eq 'Cut') { $cutcopyoption = $true }
                else { $cutcopyoption = $false }

                Write-Verbose ("Copy-SLCellValue :`t'{0}' cell '{1}' to cell '{2}'" -f $CutorCopy, $FromCellReference, $ToCellreference)
                # copy one cell to another - Retains Font, Alignment,Fill etc but not data validation
                $WorkBookInstance.CopyCell($FromCellReference, $ToCellreference, $cutcopyoption) | Out-Null
            }

            ### - #  SINGLECELL - DifferentWorksheet
            if ($PSCmdlet.ParameterSetName -eq 'SingleCell-SimplyCopyPaste-DifferentWorksheet')
            {
                Write-Verbose ("Copy-SLCellValue :`tCopy cell '{0}' from Worksheet '{1}' to cell '{2}' on worksheet '{3}'" -f $FromCellReference, $FromWorksheetName, $ToCellreference, $WorksheetName)
                # copy one cell to another - Retains Font, Alignment,Fill etc but not data validation
                $WorkBookInstance.CopyCellFromWorksheet($FromWorksheetName, $FromCellReference, $ToCellreference) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'SingleCell-PasteSpecial-DifferentWorksheet')
            {
                Write-Verbose ("Copy-SLCellValue :`tCopy cell '{0}' from Worksheet '{1}' to cell '{2}' on worksheet '{3}' with PasteSpecial Option '{4}' " -f $FromCellReference, $FromWorksheetName, $ToCellreference, $WorksheetName, $PasteSpecial)
                # copy one cell to another with paste option
                $WorkBookInstance.CopyCellFromWorksheet($FromWorksheetName, $FromCellReference, $ToCellreference, [SpreadsheetLight.SLPasteTypeValues]::$PasteSpecial) | Out-Null
            }


            ### - #  Range - SameWorksheet
            if ($PSCmdlet.ParameterSetName -eq 'Range-SimplyCopyPaste')
            {
                $StartCellReference, $ENDCellReference = $Range -split ':'

                Write-Verbose ("Copy-SLCellValue :`tCopy cellRange '{0}' to cell '{1}'" -f $Range, $ToAnchorCellreference)
                # copy one cell to another - Retains Font, Alignment,Fill etc but not data validation
                $WorkBookInstance.CopyCell($StartCellReference, $ENDCellReference, $ToAnchorCellreference) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Range-PasteSpecial')
            {
                $StartCellReference, $ENDCellReference = $Range -split ':'
                Write-Verbose ("Copy-SLCellValue :`tCopy cellRange '{0}' to cell '{1}' with PasteSpecial Option '{2}'" -f $Range, $ToAnchorCellreference, $PasteSpecial)
                # copy one cell to another with paste option
                $WorkBookInstance.CopyCell($StartCellReference, $ENDCellReference, $ToAnchorCellreference, [SpreadsheetLight.SLPasteTypeValues]::$PasteSpecial) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Range-CutOrCopy')
            {
                if ($CutorCopy -eq 'Cut') { $cutcopyoption = $true }
                else { $cutcopyoption = $false }

                $StartCellReference, $ENDCellReference = $Range -split ':'
                Write-Verbose ("Copy-SLCellValue :`t'{0}' cellRange '{1}' to cell '{2}'" -f $CutorCopy, $Range, $ToAnchorCellreference)
                # copy one cell to another - Retains Font, Alignment,Fill etc but not data validation
                $WorkBookInstance.CopyCell($StartCellReference, $ENDCellReference, $ToAnchorCellreference, $cutcopyoption) | Out-Null
            }

            ### - #  RANGE - DifferentWorksheet
            if ($PSCmdlet.ParameterSetName -eq 'Range-SimplyCopyPaste-DifferentWorksheet')
            {
                $StartCellReference, $ENDCellReference = $Range -split ':'
                Write-Verbose ("Copy-SLCellValue :`tCopy cellrange '{0}' from Worksheet '{1}' to cell '{2}' on worksheet '{3}'" -f $Range, $FromWorksheetName, $ToAnchorCellreference, $WorksheetName)
                # copy one cell to another - Retains Font, Alignment,Fill etc but not data validation
                $WorkBookInstance.CopyCellFromWorksheet($FromWorksheetName, $StartCellReference, $ENDCellReference, $ToAnchorCellreference) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Range-PasteSpecial-DifferentWorksheet')
            {
                $StartCellReference, $ENDCellReference = $Range -split ':'
                Write-Verbose ("Copy-SLCellValue :`tCopy cellrange '{0}' from Worksheet '{1}' to cell '{2}' on worksheet '{3}' with PasteSpecial Option '{4}' " -f $Range, $FromWorksheetName, $ToAnchorCellreference, $WorksheetName, $PasteSpecial)
                # copy one cell to another with paste option
                $WorkBookInstance.CopyCellFromWorksheet($FromWorksheetName, $StartCellReference, $ENDCellReference, $ToAnchorCellreference, [SpreadsheetLight.SLPasteTypeValues]::$PasteSpecial) | Out-Null
            }


            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }#select-slworksheet

    }#process
    END
    {
    }

}
