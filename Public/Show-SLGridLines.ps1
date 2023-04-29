Function Show-SLGridLines  {


    <#

.SYNOPSIS
    Show Gridlines from one or more worksheets.

.DESCRIPTION
    Show Gridlines from one or more worksheets.


.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet where gridlines are to be shown.

.PARAMETER All
    If specified will show gridlines from all worksheets within a workbook( assuming they have been hidden ).
    The user does not need to use the worksheetname parameter in conjunction with the all parameter.



.Example
    PS C:\> Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx | Show-SLGridLines -WorksheetName sheet1,sheet3 -Verbose | Save-SLDocument

    VERBOSE: sheet1 : is now selected
    VERBOSE: sheet1	: Showing Gridlines...
    VERBOSE: sheet3 : is now selected
    VERBOSE: sheet3	: Showing Gridlines...


    Description
    -----------
    Show gridlines from sheet1 & sheet3 contained in MyFirstDoc.


.Example
    PS C:\> Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx | Show-SLGridLines -All -Verbose  | Save-SLDocument

    VERBOSE: Sheet1	: Showing Gridlines...
    VERBOSE: sheet2_renamed	: Showing Gridlines...
    VERBOSE: Sheet3	: Showing Gridlines...
    VERBOSE: Sheet4	: Showing Gridlines...
    VERBOSE: Sheet5	: Showing Gridlines...
    VERBOSE: A	: Showing Gridlines...
    VERBOSE: B	: Showing Gridlines...
    VERBOSE: C	: Showing Gridlines...
    VERBOSE: Finsihed Showing Gridlines for all worksheets contained in workbook : MyFirstDoc


    Description
    -----------
    Show gridlines from all worksheets contained in MyFirstDoc.



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
        [string[]]$WorksheetName,

        [parameter(ParameterSetName = 'All')]
        [Switch]$All



    )
    PROCESS
    {

        $sheets = $WorkBookInstance.GetSheetNames()
        $pagesettings = New-Object SpreadsheetLight.SLPageSettings
        $pagesettings.ShowGridLines = $true

        # ParameterSet 'All' - Will Show gridlines from all worksheets contained in the specified workbook
        If ($PSCmdlet.ParameterSetName -eq 'All')
        {

            foreach ($w in $sheets)
            {
                $WorkBookInstance.SelectWorksheet($w) | Out-Null
                Write-Verbose ("Show-SLGridLines : '{0}'`t: Showing Gridlines..." -f $w )
                $WorkBookInstance.SetPageSettings($pagesettings, $w) | Out-Null
            }
            Write-Verbose "Show-SLGridLines : Finsihed Showing Gridlines for all worksheets contained in workbook : $($WorkBookInstance.WorkbookName)  "
        }

        # ParameterSet 'Named' - Will Show gridlines specified worksheet contained in the specified workbook
        if ($PSCmdlet.ParameterSetName -eq 'Named')
        {

            foreach ($w in $WorksheetName)
            {
                if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $w -NoPassThru)
                {
                    Write-Verbose ("Show-SLGridLines : '{0}'`t: Showing Gridlines..." -f $w )
                    $WorkBookInstance.SetPageSettings($pagesettings, $w) | Out-Null
                }
            }
        }

        $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

    }# Process


    END
    {

    }

}
