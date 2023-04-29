Function Protect-SLWorksheet  {


    <#

.SYNOPSIS
    Protect Worksheet.

.DESCRIPTION
    Protect Worksheet. Settings disabled are as follows:
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

    Currently disabling or enabling individual settings don't work as expected and hence they are not made available.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Protect-SLWorksheet -worksheet sheet2  -Verbose  | Save-SLDocument



    Description
    -----------
    Protect sheet2.



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

            $SheetProtection = New-Object SpreadsheetLight.SLSheetProtection

            Write-Verbose ("Protect-SLWorksheet :`tDisabling all Modify/filter/sort settings on worksheet '{0}'" -f $WorksheetName)

            $SheetProtection.AllowEditObjects = $false
            $SheetProtection.AllowAutoFilter = $false
            $SheetProtection.AllowDeleteColumns = $false
            $SheetProtection.AllowDeleteRows = $false
            $SheetProtection.AllowFormatCells = $false
            $SheetProtection.AllowFormatColumns = $false
            $SheetProtection.AllowFormatRows = $false
            $SheetProtection.AllowInsertColumns = $false
            $SheetProtection.AllowInsertRows = $false
            $SheetProtection.AllowPivotTables = $false
            $SheetProtection.AllowSelectLockedCells = $false
            $SheetProtection.AllowSelectUnlockedCells = $false
            $SheetProtection.AllowSort = $false
            $SheetProtection.AllowInsertHyperlinks = $false

            $WorkBookInstance.ProtectWorksheet($SheetProtection) | Out-Null

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }

    }#process
    END
    {
    }

}
