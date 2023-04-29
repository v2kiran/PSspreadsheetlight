Function Hide-SLGridLines  {


    <#

.SYNOPSIS
    Removes Gridlines from one or more worksheets.

.DESCRIPTION
    Removes Gridlines from one or more worksheets.


.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet from which gridlines are to be removed.

.PARAMETER All
    If specified will remove gridlines from all worksheets within a workbook.
    The user does not need to use the worksheetname parameter in conjunction with the all parameter.



.Example
    PS C:\> Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx | Hide-SLGridLines -WorksheetName sheet1,sheet3 -Verbose | Save-SLDocument
    VERBOSE: sheet1 : is now selected
    VERBOSE: sheet1	: Removing Gridlines...
    VERBOSE: sheet3 : is now selected
    VERBOSE: sheet3	: Removing Gridlines...


    Description
    -----------
    Remove gridlines from sheet1 & sheet3 contained in MyFirstDoc.


.Example
    PS C:\> Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx | Hide-SLGridLines -All -Verbose  | Save-SLDocument

    VERBOSE: WorkSheet - Sheet1	: Removing Gridlines...
    VERBOSE: WorkSheet - sheet2	: Removing Gridlines...
    VERBOSE: WorkSheet - Sheet3	: Removing Gridlines...
    VERBOSE: WorkSheet - Sheet4	: Removing Gridlines...
    VERBOSE: WorkSheet - Sheet5	: Removing Gridlines...
    VERBOSE: WorkSheet - A	: Removing Gridlines...
    VERBOSE: WorkSheet - B	: Removing Gridlines...
    VERBOSE: WorkSheet - C	: Removing Gridlines...
    VERBOSE: Finsihed removing Gridlines from all worksheets contained in workbook : MyFirstDoc


    Description
    -----------
    Remove gridlines from all worksheets contained in MyFirstDoc.



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

        [parameter(Mandatory = $true, Position = 1, ValueFromPipeLine = $false, ParameterSetName = 'Named')]
        [string[]]$WorkSheetName,

        [parameter(ParameterSetName = 'All')]
        [Switch]$All


    )
    PROCESS
    {
        $sheets = $WorkBookInstance.GetSheetNames()
        $pagesettings = New-Object SpreadsheetLight.SLPageSettings
        $pagesettings.ShowGridLines = $false

        # ParameterSet 'All' - Will Remove gridlines from all worksheets contained in the specified workbook
        If ($PSCmdlet.ParameterSetName -eq 'All')
        {

            foreach ($w in $sheets)
            {
                $WorkBookInstance.SelectWorksheet($w) | Out-Null
                Write-Verbose ("Hide-SLGridLines :`tWorkSheet - '{0}'`t: Removing Gridlines..." -f $w )
                $WorkBookInstance.SetPageSettings($pagesettings, $w) | Out-Null
            }
            Write-Verbose "Hide-SLGridLines : Finsihed removing Gridlines from all worksheets contained in workbook : $($WorkBookInstance.WorkbookName)  "
        }

        # ParameterSet 'Named' - Will Remove gridlines specified worksheet contained in the specified workbook
        if ($PSCmdlet.ParameterSetName -eq 'Named')
        {

            foreach ($w in $WorksheetName)
            {
                if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $w -NoPassThru)
                {
                    Write-Verbose ("Hide-SLGridLines : '{0}'`t: Removing Gridlines..." -f $w )
                    $WorkBookInstance.SetPageSettings($pagesettings, $w) | Out-Null
                }
            }
        }

        $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

    }# process

    END
    {

    }

}
