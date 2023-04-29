Function Set-SLDataValidation  {


    <#

.SYNOPSIS
    Add Datavalidation.

.DESCRIPTION
    Create drop-down lists or otherwise control the type of data that users enter on a worksheet.
    Apply data constraints on Integers,decimals,Date,Time,TextLength or custom forumulas.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER ValidationTarget
    The target cell or range of cells that need to have datavalidation.

.PARAMETER DataLookupRange
    The range that can be used to create a drop-down list on a cell or range of cells.

.PARAMETER DefinedName
    The definedname that can be used to create a drop-down list on a cell or range of cells.

.PARAMETER Decimal
    Restrict data entry to a decimal number  Example - "1.3".

.PARAMETER StartDecimal
    This is the minimum value for a decimal range.

.PARAMETER EndDecimal
    This is the maximum value for a decimal range.


.PARAMETER WholeNumber
    Restrict data entry to a wholenumber  Example - "3".

.PARAMETER StartWholeNumber
    This is the minimum value for a wholenumber range.

.PARAMETER EndWholeNumber
    This is the maximum value for a wholenumber range.

.PARAMETER Date
    Restrict data entry to a Date  Example - "12/25/2014".

.PARAMETER StartDate
    This is the minimum value for a Date range.

.PARAMETER EndDate
    This is the maximum value for a Date range.

.PARAMETER Time
    Restrict data entry to a Time  Example - "14:30:55".

.PARAMETER StartTime
    This is the minimum value for a Time range.

.PARAMETER EndTime
    This is the maximum value for a Time range.

.PARAMETER TextLength
    Restrict data entry to a TextLength  Example - "6".

.PARAMETER StartTextLength
    This is the minimum value for a TextLength range.

.PARAMETER EndTextLength
    This is the maximum value for a TextLength range.

.PARAMETER CustomFormula
    Restrict data entry to values that conform to a CustomFormula  Example - "=len(b3)".

.PARAMETER ValidationOperator
    The Operator to be used for validating data.
    Use tab or intellisense to select from a list of possible values.
    'Equal','NotEqual','GreaterThan','LessThan','GreaterThanOrEqual','LessThanOrEqual','Between'

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\DataValidation.xlsx
    PS C:\> $doc | Set-SLColumnValue -WorksheetName sheet1 -CellReference B3 -value @('Pete','Andre','Roger','Jimmy','Pat') -Verbose
    PS C:\> $doc | Set-SLDataValidation -WorksheetName sheet1 -ValidationTarget C3 -DataLookupRange B3:B7 -Verbose
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    Set column vlaues B3 to B7 and use that range to create a drop-down list in cell C3.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\DataValidation.xlsx
    PS C:\> $doc | Set-SLDataValidation -WorksheetName sheet1 -ValidationTarget D3:E4 -DataLookupRange B3:B7 -Verbose
    PS C:\> $doc | Save-SLDocument


    Description
    -----------
    Use a predefined range B3:B7(values we created in example 1 above) to create a drop-down list in a range of cells D3:E4.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\DataValidation.xlsx
    PS C:\> $doc | Set-SLColumnValue    -WorksheetName sheet2 -CellReference C3 -value @('Sampras','Agassi','Federer','Connors','Rafter') -Verbose
    PS C:\> $doc | New-SLDefinedName    -WorksheetName sheet2 -DefinedName LookupRange1 -Range C3:C7 -Verbose
    PS C:\> $doc | Set-SLDataValidation -WorksheetName sheet1 -ValidationTarget F3 -DefinedName LookupRange1 -Verbose
    PS C:\> $doc | Save-SLDocument


    Description
    -----------
    Set column vlaues C3 to C7 and use that range to create a DefinedName called 'lookuprange1' and finally use the DefinedName to create a drop-down list in cell F3.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\DataValidation.xlsx
    PS C:\> $doc | Set-SLDataValidation -WorksheetName sheet1 -ValidationTarget c4 -StartDecimal 1.2 -ENDDecimal 2.5 -ValidationOperator Between -Verbose
    PS C:\> $doc | Save-SLDocument


    Description
    -----------
    Set datavalidation on cell c4 to contain only values that are between 1.2 and 2.5.
    Note: if you omit the ValidationOperator in the command above the validation operator defaults to 'notbetween'
    so in effect the validation would then be all values that are not between 1.2 and 2.5

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\DataValidation.xlsx
    PS C:\> $doc | Set-SLDataValidation -WorksheetName sheet1 -ValidationTarget c5 -WholeNumber 5 -ValidationOperator Equal -Verbose
    PS C:\> $doc | Save-SLDocument


    Description
    -----------
    Set datavalidation on cell c5 to contain only value that is equal to 5.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\DataValidation.xlsx
    PS C:\> $doc | Set-SLColumnValue -WorksheetName sheet2 -CellReference D3 -value @(1,2,3)
    PS C:\> $doc | Set-SLColumnValue -WorksheetName sheet2 -CellReference E3 -value @(3,4,5)
    PS C:\> $doc | New-SLDefinedName    -WorksheetName sheet2 -DefinedName MinRangeValue -Range D3:D5
    PS C:\> $doc | New-SLDefinedName    -WorksheetName sheet2 -DefinedName MaxRangeValue -Range E3:E5
    PS C:\> $doc | Set-SLDataValidation -WorksheetName sheet1 -ValidationTarget C13 -StartWholeNumber '=MIN(MinRangeValue)' -EndWholeNumber '=MAX(MaxRangeValue)' -ValidationOperator Between
    PS C:\> $doc | Save-SLDocument


    Description
    -----------
    Set column values D3:D5 and also E3:E5.
    Create defined names for each of the above ranges
    Set datavalidation on C13 that makes use of the 2 defined names created above.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\DataValidation.xlsx
    PS C:\> $doc | Set-SLDataValidation -WorksheetName sheet1 -ValidationTarget C14 -StartDate '12/20/2014' -EndDate '12/25/2014' -ValidationOperator Between -Verbose
    PS C:\> $doc | Save-SLDocument


    Description
    -----------
    Restrict cell c14 to contain dates between '12/20/2014' & '12/25/2014'

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\DataValidation.xlsx
    PS C:\> $doc | Set-SLDataValidation -WorksheetName sheet1 -ValidationTarget C17 -Date '12/25/2014' -ValidationOperator Equal -Verbose
    PS C:\> $doc | Save-SLDocument


    Description
    -----------
    Restrict cell c17 to contain date equal to '12/25/2014'.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\DataValidation.xlsx
    PS C:\> $doc | Set-SLDataValidation -WorksheetName sheet1 -ValidationTarget C20 -Time 14:20:35 -ValidationOperator LessThan
    PS C:\> $doc | Save-SLDocument


    Description
    -----------
    Restrict cell c20 to contain time values that are lessthan 14:20:35.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\DataValidation.xlsx
    PS C:\> $doc | Set-SLDataValidation -WorksheetName sheet1 -ValidationTarget C23 -TextLength '=SUM(LEN(B3),LEN(F3))' -ValidationOperator LessThanOrEqual -Verbose
    PS C:\> $doc | Save-SLDocument


    Description
    -----------
    The forumala SUM(LEN(B3),LEN(F3)) --> compute the length of the cell value B3, compute the length of cell value F3 and add them up.
    Restrict cell c23 to contain time values that are lessthanoeEqual to the textlength obtained by the formula above.
    If the forumal SUM(LEN(B3),LEN(F3)) yeilded value 11 then the total length of the value in cell C23 cannot exceed 11.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\DataValidation.xlsx
    PS C:\> $doc | Set-SLDataValidation -WorksheetName sheet1 -ValidationTarget D14:D17 -CustomFormula 'COUNTIF($D$14:$D$17,D14) <= 1' -Verbose
    PS C:\> $doc | Save-SLDocument


    Description
    -----------
    The forumala 'COUNTIF($D$14:$D$17,D14) <= 1' --> count the occurrences of the value in cell D14, in the range $D$14:$D$17. The formula's result must be 1 or 0
    The net result is to prevent duplicate values from being entered in the range D14:D17.


.INPUTS
   String,Int,Double,SpreadsheetLight.SLDocument

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    http://www.contextures.com/xlDataVal07.html

#>



    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true, Position = 0, ValueFromPipeLine = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [Alias('CurrentWorkSheetName')]
        [parameter(Mandatory = $true, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true)]
        [string]$ValidationTarget,

        [Alias('Range')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'DataLookupRange')]
        [String]$DataLookupRange,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'NamedRange')]
        [String]$DefinedName,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Decimal')]
        [Double]$Decimal,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'StartEndDecimal')]
        [Double]$StartDecimal,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'StartEndDecimal')]
        [Double]$EndDecimal,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'WholeNumber')]
        [Int]$WholeNumber,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'StartEndWholeNumber')]
        $StartWholeNumber,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'StartEndWholeNumber')]
        $EndWholeNumber,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Date')]
        [String]$Date,

        [parameter(Mandatory = $True, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'StartEndDate')]
        [String]$StartDate,

        [parameter(Mandatory = $True, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'StartEndDate')]
        [String]$EndDate,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Time')]
        [String]$Time,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'StartEndTime')]
        [String]$StartTime,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'StartEndTime')]
        [String]$EndTime,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'TextLength')]
        $TextLength,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'StartEndTextLength')]
        $StartTextLength,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'StartEndTextLength')]
        $EndTextLength,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Custom')]
        [string]$CustomFormula,

        [ValidateSet('Equal', 'NotEqual', 'GreaterThan', 'LessThan', 'GreaterThanOrEqual', 'LessThanOrEqual', 'Between')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'WholeNumber')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'StartEndWholeNumber')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Date')]
        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'StartEndDecimal')]
        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'StartEndDate')]
        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'StartEndTime')]
        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'StartEndTextLength')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Time')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'TextLength')]
        [String]$ValidationOperator


    )
    PROCESS
    {

        ## -- ## Check if the referenced worksheet exists in the workbook and proceed only if true.
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {

            Switch -Regex ($ValidationTarget)
            {

                #CellReference
                '^[a-zA-Z]+\d+$'
                {
                    Write-Verbose ("Set-SLDataValidation :`tValidationTarget is Cell '{0}'" -f $ValidationTarget)
                    $DataValidation = $WorkBookInstance.CreateDataValidation($ValidationTarget)
                    $isValidationTargetValid = $true
                    Break
                }

                #Range
                '[a-zA-Z]+\d+:[a-zA-Z]+\d+$'
                {
                    $startcellreference, $endcellreference = $ValidationTarget -split ':'
                    Write-Verbose ("Set-SLDataValidation :`tValidationTarget is Cell Range '{0}'" -f $ValidationTarget)
                    $DataValidation = $WorkBookInstance.CreateDataValidation($startcellreference, $endcellreference)
                    $isValidationTargetValid = $true
                    Break
                }

                Default
                {
                    Write-Warning ("Set-SLDataValidation :`tYou must provide either a Cellreference Eg. C3 or a Range Eg. C3:G10")
                    $isValidationTargetValid = $false
                    Break
                }

            }#switch

            if ( ($PSCmdlet.ParameterSetName -eq 'DataLookupRange') -and $isValidationTargetValid)
            {
                Write-Verbose ("Set-SLDataValidation :`tDatalookup Range '{0}'" -f $DataLookupRange)
                $AbsoluteRange = Convert-ToExcelAbsoluteRange -Range $DataLookupRange

                #allowlist - DataSource, IgnoreBlank, InCellDropDown
                $DataValidation.AllowList($AbsoluteRange, $true, $true)
                $DataValidation.SetInputMessage('ValidationMessage', "Only Values in the Cell Range - $DataLookupRange are accepted") | Out-Null
                $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You Must enter a Value specified in the range - $range") | Out-Null

            }

            if ( ($PSCmdlet.ParameterSetName -eq 'NamedRange') -and $isValidationTargetValid)
            {

                if ((($WorkBookInstance.GetDefinedNames() | Select-Object -ExpandProperty Name) -contains $DefinedName))
                {
                    $DefinedNameText = $WorkBookInstance.GetDefinedNameText($DefinedName)
                    Write-Verbose ("Set-SLDataValidation :`tNamedRange '{0}' corresponds to '{1}'" -f $DefinedName, $DefinedNameText)

                    $DataValidation.AllowList("=$DefinedName", $true, $true)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values in the Cell Range - $DefinedNameText are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You Must enter a Value specified in the range - $range") | Out-Null
                }
                Else
                {
                    Write-Warning ("Set-SLDataValidation :`tSpecified Named Range '{0}' was not found in the workbook. Check spelling and try again." -f $DefinedName)
                    break
                }
            }

            if ( ($PSCmdlet.ParameterSetName -eq 'Decimal') -and $isValidationTargetValid)
            {

                If ($ValidationOperator -ne 'Between')
                {
                    Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should be '{1}' Date '{2}'" -f $ValidationTarget, $ValidationOperator, $Decimal)
                    $DataValidation.AllowDecimal([SpreadsheetLight.SLDataValidationSingleOperandValues]::$ValidationOperator, $Decimal, $true)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values that are $ValidationOperator $Decimal are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You Must enter a Value that is $ValidationOperator  Decimal $Decimal") | Out-Null
                }
                else
                {
                    Write-Warning ("Set-SLDataValidation :`tUse ValidationOperator 'Between' with Parameters 'StartDecimal' & EndDecimal' ")
                    Break
                }
            }

            if ($PSCmdlet.ParameterSetName -eq 'StartEndDecimal' -and $isValidationTargetValid)
            {
                If ($ValidationOperator -eq 'Between')
                {
                    Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should be between '{1}' & '{2}'" -f $ValidationTarget, $StartDecimal, $ENDDecimal)
                    $DataValidation.AllowDecimal($true, $StartDecimal, $ENDDecimal, $false)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values in the Decimal Range - $StartDecimal - $ENDDecimal are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You Must enter a Value specified in the range - $StartDecimal - $ENDDecimal") | Out-Null
                }
                else
                {
                    Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should NOT be between '{1}' & '{2}'" -f $ValidationTarget, $StartDecimal, $ENDDecimal)
                    $DataValidation.AllowDecimal($true, $StartDecimal, $ENDDecimal, $false)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values NOT in the Decimal Range - $StartDecimal - $ENDDecimal are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You Must enter a Value NOT in the range - $StartDecimal - $ENDDecimal") | Out-Null
                }
            }

            if ($PSCmdlet.ParameterSetName -eq 'WholeNumber' -and $isValidationTargetValid)
            {
                If ($ValidationOperator -ne 'Between')
                {
                    Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should be '{1}' to '{2}'" -f $ValidationTarget, $ValidationOperator, $WholeNumber)
                    $DataValidation.AllowWholeNumber([SpreadsheetLight.SLDataValidationSingleOperandValues]::$ValidationOperator, $WholeNumber, $true)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values $ValidationOperator - $WholeNumber are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "Validation Criteria: Value $ValidationOperator - $WholeNumber not met") | Out-Null
                }
                else
                {
                    Write-Warning ("Set-SLDataValidation :`tWhen Parameter 'Wholenumber' is used, the value of the Validationoperator must NOT be 'Between'. Use 'Between' with StartWholeNumber & EndWholeNumber ")
                    Break
                }
            }

            if ($PSCmdlet.ParameterSetName -eq 'StartEndWholeNumber' -and $isValidationTargetValid)
            {
                If ($ValidationOperator -eq 'Between')
                {
                    Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should be between '{1}' and '{2}'" -f $ValidationTarget, $StartWholeNumber, $EndWholeNumber)
                    $DataValidation.AllowWholeNumber($true, $StartWholeNumber, $EndWholeNumber, $true)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values $ValidationOperator : $StartWholeNumber-$EndWholeNumber are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You must enter a value that is between $StartWholeNumber-$EndWholeNumber") | Out-Null
                }
                else
                {
                    Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should NOT be between '{1}' and '{2}'" -f $ValidationTarget, $StartWholeNumber, $EndWholeNumber)
                    $DataValidation.AllowWholeNumber($false, $StartWholeNumber, $EndWholeNumber, $true)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values NOT between: $StartWholeNumber-$EndWholeNumber are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You must enter a value that is NOT between $StartWholeNumber-$EndWholeNumber") | Out-Null
                }
            }

            if ($PSCmdlet.ParameterSetName -eq 'Date' -and $isValidationTargetValid)
            {
                If ($ValidationOperator -ne 'Between')
                {
                    Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should be '{1}' Date '{2}'" -f $ValidationTarget, $ValidationOperator, $Date)
                    $DataValidation.AllowDate([SpreadsheetLight.SLDataValidationSingleOperandValues]::$ValidationOperator, $Date, $true)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values that are $ValidationOperator $date are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You Must enter a Value that is $ValidationOperator  Date $Date") | Out-Null
                }
                else
                {
                    Write-Warning ("Set-SLDataValidation :`tUse ValidationOperator 'Between' with Parameters 'StartDate' & EndDate' ")
                    Break
                }
            }

            if ($PSCmdlet.ParameterSetName -eq 'StartEndDate' -and $isValidationTargetValid)
            {
                If ($ValidationOperator -eq 'Between')
                {
                    Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should be between '{1}' & '{2}'" -f $ValidationTarget, $StartDate, $EndDate)
                    $DataValidation.AllowDate($true, $StartDate, $EndDate, $true)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values between Dates $StartDate & $EndDate are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You Must enter a Date that is between $StartDate & $EndDate") | Out-Null
                }
                Else
                {
                    Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should NOT be between '{1}' & '{2}'" -f $ValidationTarget, $StartDate, $EndDate)
                    $DataValidation.AllowDate($false, $StartDate, $EndDate, $true)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values NOT between Dates $StartDate & $EndDate are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You Must enter a Date that is NOT between $StartDate & $EndDate") | Out-Null

                }
            }

            if ($PSCmdlet.ParameterSetName -eq 'Time' -and $isValidationTargetValid)
            {

                If ($ValidationOperator -ne 'Between')
                {
                    Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should be '{1}' Time '{2}'" -f $ValidationTarget, $ValidationOperator, $Time)
                    $DataValidation.AllowTime([SpreadsheetLight.SLDataValidationSingleOperandValues]::$ValidationOperator, $Time, $true)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values that are $ValidationOperator $Time are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You Must enter a Value that is $ValidationOperator  Time $Time") | Out-Null
                }
                else
                {
                    Write-Warning ("Set-SLDataValidation :`tUse ValidationOperator 'Between' with Parameters 'StartTime' & EndTime' ")
                    Break
                }
            }

            if ($PSCmdlet.ParameterSetName -eq 'StartEndTime' -and $isValidationTargetValid)
            {
                If ($ValidationOperator -eq 'Between')
                {
                    Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should be between Time values '{1}' & '{2}'" -f $ValidationTarget, $StartTime, $EndTime)
                    $DataValidation.AllowTime($true, $StartTime, $EndTime, $true)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values between Times $StartTime & $EndTime are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You Must enter a Time that is between $StartTime & $EndTime") | Out-Null
                }
                Else
                {
                    Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should NOT be between Time values '{1}' & '{2}'" -f $ValidationTarget, $StartTime, $EndTime)
                    $DataValidation.AllowTime($false, $StartTime, $EndTime, $true)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values NOT between Times $StartTime & $EndTime are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You Must enter a Time that is NOT between $StartTime & $EndTime") | Out-Null

                }
            }



            if ($PSCmdlet.ParameterSetName -eq 'TextLength' -and $isValidationTargetValid)
            {

                If ($ValidationOperator -ne 'Between')
                {
                    Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should be '{1}' TextLength '{2}'" -f $ValidationTarget, $ValidationOperator, $Time)
                    $DataValidation.AllowTextLength([SpreadsheetLight.SLDataValidationSingleOperandValues]::$ValidationOperator, $TextLength, $true)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values that are $ValidationOperator $TextLength are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You Must enter a Value that is $ValidationOperator  TextLength $TextLength") | Out-Null
                }
                else
                {
                    Write-Warning ("Set-SLDataValidation :`tUse ValidationOperator 'Between' with Parameters 'StartTextLength' & EndTextLength' ")
                    Break
                }
            }

            if ($PSCmdlet.ParameterSetName -eq 'StartEndTextLength' -and $isValidationTargetValid)
            {
                If ($ValidationOperator -eq 'Between')
                {
                    Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should be between TextLength values '{1}' & '{2}'" -f $ValidationTarget, $StartTextLength, $EndTextLength)
                    $DataValidation.AllowTextLength($true, $StartTextLength, $EndTextLength, $true)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values between TextLengths $StartTextLength & $EndTextLength are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You Must enter a TextLength that is between $StartTextLength & $EndTextLength") | Out-Null
                }
                Else
                {
                    Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should NOT be between Time values '{1}' & '{2}'" -f $ValidationTarget, $StartTextLength, $EndTextLength)
                    $DataValidation.AllowTextLength($false, $StartTextLength, $EndTextLength, $true)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values NOT between TextLengths $StartTextLength & $EndTextLength are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You Must enter a TextLength that is NOT between $StartTextLength & $EndTextLength") | Out-Null

                }
            }

            if ($PSCmdlet.ParameterSetName -eq 'Custom' -and $isValidationTargetValid)
            {
                Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should conform to formula '{1}'" -f $ValidationTarget, $CustomFormula)
                $DataValidation.AllowCustom($CustomFormula, $true)
                $DataValidation.SetInputMessage('ValidationMessage', "Only Values that conform to the forumula  - $CustomFormula are accepted") | Out-Null
                $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You Must enter a Value that is valid for forumula - $CustomFormula") | Out-Null
            }

            if ($isValidationTargetValid)
            {
                Write-Verbose ("Set-SLDataValidation :`tAdding Datavalidation..")
                $WorkBookInstance.AddDataValidation($DataValidation) | Out-Null
                $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
            }


        }#END if select-slworksheet

    }#process
    END
    {
    }

}
