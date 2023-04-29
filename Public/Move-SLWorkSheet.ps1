Function Move-SLWorkSheet  {


    <#

.SYNOPSIS
    Moves a worksheet.

.DESCRIPTION
    Moves a worksheet.Only one worksheet can be moved at a time.


.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet to be Moved.

.PARAMETER Position
    Position index. Use 1 for 1st position, 2 for 2nd position and so on.
    If there were 10 worksheets in a document and you specify the value of the position parameter as 20 for sheet1 the
    worksheet will then be moved to the last position in the document.



.Example
    PS C:\> Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx | Move-SLWorkSheet -worksheetName sheet1 -Position 100 -Verbose | Save-SLDocument

    VERBOSE: sheet1 : Worksheet is being Moved to Position : 100
    VERBOSE: All Done!

    Description
    -----------
    Sheet1 is moved from its current position to position 100, assuming of course there are 100 worksheets in the document.
    If there were 10 worksheets in a document and you specify the value of the position parameter as 20 for sheet1 the
    worksheet will then be moved to the last position in the document.


.INPUTS
   String,Int,SpreadsheetLight.SLDocument

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

        [parameter(Mandatory = $true, Position = 1, ValueFromPipeLine = $false)]
        [string]$WorkSheetName,

        [ValidateRange(1, 1000)]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true)]
        [System.UInt32]$Position


    )
    PROCESS
    {
        if ($WorkBookInstance.GetSheetNames() -contains $WorkSheetName)
        {
            Write-Verbose ("Move-SLWorkSheet :`tWorksheet '{0}' is being Moved to Position '{1}'" -f $worksheetName, $Position )
            $WorkBookInstance.MoveWorkSheet($worksheetName, $Position) | Out-Null
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }
        Else
        {
            Write-Warning ("Move-SLWorkSheet :  :`tWorksheet '{0}' Could not be Found. Check the spelling and try again." -f $WorkSheetName )
        }
    }
    END
    {
        Write-Verbose 'Move-SLWorkSheet : All Done!'
    }

}
