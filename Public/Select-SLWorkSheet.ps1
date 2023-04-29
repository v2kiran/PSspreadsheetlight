Function Select-SLWorkSheet  {


    <#

.SYNOPSIS
    Select a worksheet from a workbook for editing.

.DESCRIPTION
    Select a worksheet from a workbook for editing.
    When a workbook contains multiple worksheets it is important that you use select-slworksheet to
    select the required worksheet if not more often than not you might find that the worksheet
    you made changes to is not the one you wanted.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    The worksheet that needs to be selected and made active for editing.

.PARAMETER NoPassThru
    Retunrs a boolean based on whether the given worksheet was selected or not.Does not pass the workbookinstance through the pipeline.
    By default the workbookinstance is passed through the pipeline so that other commands can work with it.


.Example
    PS C:\> $doc = Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx

    PS C:\> $Doc | Select-SLWorkSheet -WorksheetName sheet2 -Verbose
    VERBOSE: Worksheet :	sheet2 is now selected


    WorkbookName         : MyFirstDoc
    WorksheetName        : {Sheet1, Sheet2, Sheet3, Sheet4...}
    Path                 : D:\ps\Excel\MyFirstDoc.xlsx
    CurrentWorksheetName : sheet2
    DocumentProperties   : SpreadsheetLight.SLDocumentProperties


    Description
    -----------
    'MyFirstDoc' contains more than 4 worksheets.By default sheet5 which is the last worksheet is active.
     We use select-slworksheet to select sheet2 and make it active.



.INPUTS
   String

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    N/A

#>


    [CmdletBinding()]
    [OutputType([SpreadsheetLight.SLDocument])]
    param (
        [parameter(Mandatory = $true, Position = 1, ValueFromPipeLine = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [parameter(Mandatory = $true, Position = 0)]
        [String]$WorksheetName,

        [parameter(Mandatory = $false)]
        [Switch]$NoPassThru

    )
    PROCESS
    {
        if ($WorkBookInstance.GetSheetNames() -contains $WorksheetName)
        {
            $selected = $true
            $WorkBookInstance.SelectWorkSheet($WorksheetName) | Out-Null
            Write-Verbose ("Select-SLWorkSheet :`tWorksheet '{0}' is now selected" -f $WorksheetName)
        }
        Else
        {
            Write-Warning ("Select-SLWorkSheet : Could Not Find a WorkSheet Named '{0}'.Check the spelling and try again" -f $WorksheetName)
            $selected = $false
        }

        if ($NoPassThru)
        {
            return $selected
        }
        Else
        {
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }

    }
    END
    {

    }

}
