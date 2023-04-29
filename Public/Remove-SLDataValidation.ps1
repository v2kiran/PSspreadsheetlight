Function Remove-SLDataValidation  {


    <#

.SYNOPSIS
    Clear all data validation from a worksheet.

.DESCRIPTION
    Clear all data validation from a worksheet.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    This is the name of the worksheet that contains datavalidation to be removed.

.Example
    PS C:\> Get-SLDocument  D:\ps\Excel\Vlookup.xlsx | Remove-SLDataValidation -WorksheetName sheet1 -Verbose | Save-SLDocument


    Description
    -----------
    Removes all data validation entries form worksheet named 'sheet1'.


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
        [String]$WorksheetName

    )

    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            Write-Verbose ("Remove-SLDataValidation :`tRemoving all data validation entries from Worksheet '{0}'" -f $WorksheetName)
            $WorkBookInstance.ClearDataValidation()
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }

    }#process
    END
    {
    }

}
