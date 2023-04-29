Function New-SLDefinedName  {


    <#

.SYNOPSIS
    Create a Defined Name for a cell reference, range, constant, formula, or table.

.DESCRIPTION
    A name is a meaningful shorthand that makes it easier to understand the purpose of a cell reference,
    constant, formula, or table, each of which may be difficult to comprehend at first glance.
    The following information shows common examples of names and how they can improve clarity and understanding.

    EXAMPLE TYPE	EXAMPLE WITH NO NAME	                EXAMPLE WITH A NAME
    Reference	    =SUM(C20:C30)	                        =SUM(FirstQuarterSales)
    Constant	    =PRODUCT(A5,8.3)	                    =PRODUCT(Price,WASalesTax)
    Formula	        =SUM(VLOOKUP(A1,B1:F20,5,FALSE), -G5)	=SUM(Inventory_Level,-Order_Amt)
    Table	        C4:G36	                                =TopSales06

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that contains the Range referenced by the defined name.

.PARAMETER DefinedName
    A userfriendly Name for a cell reference, range, constant, formula, or table.

.PARAMETER Range
    cellrange which would be the datasource for a defined name.
    To define a cellreference instead of a range use the range format like so: B3:B3

.PARAMETER Comment
    Comment that provides a short description of the defined name.

.PARAMETER Scope
    The name of the worksheet that the defined name is effective in.

.PARAMETER Force
    If the defined name to be created already exists in the workbook use the force switch to overwrite the existing value.
    By default the cmdlet will not overwrite an existing Defined Name.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | New-SLDefinedName -WorksheetName sheet1 -DefinedName DFName1 -Range B3:B7 -Verbose | Save-SLDocument


    Description
    -----------
    Create a New defined name 'DFName1'.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\DataValidation.xlsx |
                New-SLDefinedName -WorksheetName sheet1 -DefinedName DFName2 -Range B3:B7 -Comment "This range represents Athlete Names" |
                    Save-SLDocument



    Description
    -----------
    Create a New defined name 'DFName1'.Additionally specify a comment to describe the defined name.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\DataValidation.xlsx |
                New-SLDefinedName -WorksheetName sheet1 -DefinedName DFName3 -Range B3:B7 -Comment "This range represents Athlete Names" -Scope sheet2 |
                    Save-SLDocument



    Description
    -----------
    Create a New defined name 'DFName1'.Additionally specify a comment and scope.
    Because we specified 'sheet2' as the value for the scope parameter, the defined name 'DFName3' can only be used on worksheet named 'sheet2'.


.INPUTS
   String,SpreadsheetLight.SLDocument

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    http://office.microsoft.com/en-in/excel-help/define-and-use-names-in-formulas-HA010147120.aspx

#>


    [CmdletBinding(DefaultParameterSetName = 'All', SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param (
        [parameter(Mandatory = $true, Position = 0, ValueFromPipeLine = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [Alias('CurrentWorkSheetName')]
        [parameter(Mandatory = $true, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true)]
        [String]$DefinedName,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+:[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "New-SLDefinedName :`tRange should specify values in following format. Eg: A1:D10 or AB1:AD5"; break }
            })]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$Range,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetname = 'Comment')]
        [string]$Comment,

        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, ParameterSetname = 'Comment', HelpMessage = 'The name of the worksheet that the defined name is effective in')]
        [string]$Scope,

        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [Switch]$Force

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {

            $DefinedNames = $WorkBookInstance.GetDefinedNames() | Select-Object -ExpandProperty Name
            if ($DefinedNames -contains $DefinedName)
            {
                $DefinedNameExists = $true
            }
            Else
            {
                $DefinedNameExists = $false
            }
            $AbsoluteRange = Convert-ToExcelAbsoluteRange -Range $Range -WorkSheetName $WorksheetName

            If ($PSCmdlet.ParameterSetName -eq 'Comment')
            {
                if ($Scope)
                {
                    If ($DefinedNameExists -and $Force -and $PSCmdlet.ShouldPROCESS($DefinedName, 'OVERWRITE DEFINED NAME') )
                    {
                        Write-Verbose ("New-SLDefinedName :`tForce parameter specified. Overwriting existing DefinedName '{0}'" -f $DefinedName)
                        Write-Verbose ("New-SLDefinedName :`tDefinedName '{0}' Scope is '{1}'" -f $DefinedName, $Scope)
                        Write-Verbose ("New-SLDefinedName :`tDefinedName '{0}' Comment is '{1}'" -f $DefinedName, $Comment)
                        Write-Verbose ("New-SLDefinedName :`tCreating DefinedName '{0}' corresponding to Range '{1}'" -f $DefinedName, $Range)
                        $WorkBookInstance.SetDefinedName($DefinedName, $AbsoluteRange, $Comment, $Scope) | Out-Null
                        $WorkBookInstance | Add-Member NoteProperty DefinedName $DefinedName -Force
                        $WorkBookInstance | Add-Member NoteProperty DefinedNameRange $AbsoluteRange -Force
                    }
                    Elseif ($DefinedNameExists)
                    {
                        Write-Warning ("New-SLDefinedName :`tDefinedName '{0}' Already exists. Specify the '-Force' parameter to Overwrite" -f $DefinedName)
                    }
                    Else
                    {
                        Write-Verbose ("New-SLDefinedName :`tDefinedName '{0}' Scope is '{1}'" -f $DefinedName, $Scope)
                        Write-Verbose ("New-SLDefinedName :`tDefinedName '{0}' Comment is '{1}'" -f $DefinedName, $Comment)
                        Write-Verbose ("New-SLDefinedName :`tCreating DefinedName '{0}' corresponding to Range '{1}'" -f $DefinedName, $Range)
                        $WorkBookInstance.SetDefinedName($DefinedName, $AbsoluteRange, $Comment, $Scope) | Out-Null
                        $WorkBookInstance | Add-Member NoteProperty DefinedName $DefinedName -Force
                        $WorkBookInstance | Add-Member NoteProperty DefinedNameRange $AbsoluteRange -Force
                    }
                }
                Else
                {
                    If ($DefinedNameExists -and $Force -and $PSCmdlet.ShouldPROCESS($DefinedName, 'OVERWRITE DEFINED NAME') )
                    {
                        Write-Verbose ("New-SLDefinedName :`tForce parameter specified. Overwriting existing DefinedName '{0}'" -f $DefinedName)
                        Write-Verbose ("New-SLDefinedName :`tDefinedName '{0}' Comment is '{1}'" -f $DefinedName, $Comment)
                        Write-Verbose ("New-SLDefinedName :`tCreating DefinedName '{0}' corresponding to Range '{1}'" -f $DefinedName, $Range)
                        $WorkBookInstance.SetDefinedName($DefinedName, $AbsoluteRange, $Comment) | Out-Null
                        $WorkBookInstance | Add-Member NoteProperty DefinedName $DefinedName -Force
                        $WorkBookInstance | Add-Member NoteProperty DefinedNameRange $AbsoluteRange -Force
                    }
                    Elseif ($DefinedNameExists)
                    {
                        Write-Warning ("New-SLDefinedName :`tDefinedName '{0}' Already exists. Specify the '-Force' parameter to Overwrite" -f $DefinedName)
                    }
                    Else
                    {
                        Write-Verbose ("New-SLDefinedName :`tDefinedName '{0}' Comment is '{1}'" -f $DefinedName, $Comment)
                        Write-Verbose ("New-SLDefinedName :`tCreating DefinedName '{0}' corresponding to Range '{1}'" -f $DefinedName, $Range)
                        $WorkBookInstance.SetDefinedName($DefinedName, $AbsoluteRange, $Comment) | Out-Null
                        $WorkBookInstance | Add-Member NoteProperty DefinedName $DefinedName -Force
                        $WorkBookInstance | Add-Member NoteProperty DefinedNameRange $AbsoluteRange -Force
                    }
                }
            }
            elseIf ($PSCmdlet.ParameterSetName -eq 'All')
            {
                If ($DefinedNameExists -and $Force -and $PSCmdlet.ShouldPROCESS($DefinedName, 'OVERWRITE DEFINED NAME') )
                {
                    Write-Verbose ("New-SLDefinedName :`tForce parameter specified. Overwriting existing DefinedName '{0}'" -f $DefinedName)
                    Write-Verbose ("New-SLDefinedName :`tCreating DefinedName '{0}' corresponding to Range '{1}'" -f $DefinedName, $Range)
                    $WorkBookInstance.SetDefinedName($DefinedName, $AbsoluteRange) | Out-Null
                    $WorkBookInstance | Add-Member NoteProperty DefinedName $DefinedName -Force
                    $WorkBookInstance | Add-Member NoteProperty DefinedNameRange $AbsoluteRange -Force
                }
                Elseif ($DefinedNameExists)
                {
                    Write-Warning ("New-SLDefinedName :`tDefinedName '{0}' Already exists. Specify the '-Force' parameter to Overwrite" -f $DefinedName)
                }
                Else
                {
                    Write-Verbose ("New-SLDefinedName :`tCreating DefinedName '{0}' corresponding to Range '{1}'" -f $DefinedName, $Range)
                    $WorkBookInstance.SetDefinedName($DefinedName, $AbsoluteRange) | Out-Null
                    $WorkBookInstance | Add-Member NoteProperty DefinedName $DefinedName -Force
                    $WorkBookInstance | Add-Member NoteProperty DefinedNameRange $AbsoluteRange -Force

                }
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }

    }
    END
    {
    }

}
