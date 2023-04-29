Function UnProtect-SLWorksheet  {


    <#

.SYNOPSIS
    UnProtect Worksheet.

.DESCRIPTION
    UnProtect Worksheet. Settings enabled are as follows:
                EditObjects
                AutoFilter
                DeleteColumns
                DeleteRows
                FormatCells
                FormatColumns
                FormatRows
                InsertColumns
                InsertRows
                PivotTables
                SelectLockedCells
                SelectUnlockedCells
                Sort
                InsertHyperlinks

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | UnProtect-SLWorksheet -worksheet sheet2  -Verbose  | Save-SLDocument



    Description
    -----------
    UnProtect sheet2.



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
            Write-Verbose ("UnProtect-SLWorksheet :`tEnabling all Modify\filter\sort settings on worksheet '{0}'" -f $WorksheetName)
            $WorkBookInstance.UnprotectWorksheet() | Out-Null

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }

    }#process
    END
    {
    }

}
