Function Add-SLWorkSheet  {


    <#

.SYNOPSIS
    Adds one or more worksheets to an Excel Document.

.DESCRIPTION
    Adds one or more worksheets to an Excel Document.If the specified worksheet to be added already exists in the workbook,
    a random number is appended to the worksheetname to make it unique and then created as a worksheet.


.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet(s) to be created.
    You can create more than one worksheet by specifying the names as a comma separated list.



.Example
    PS C:\> Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx | Add-SLWorkSheet -worksheetName A,B,C -Verbose | Save-SLDocument
    VERBOSE: A : Adding worksheet to excel document - MyFirstDoc...
    VERBOSE: B : Adding worksheet to excel document - MyFirstDoc...
    VERBOSE: C : Adding worksheet to excel document - MyFirstDoc...
    VERBOSE: All Done!

    Description
    -----------
    MyfirstDoc is piped to Add-SLWorksheet with 3 worksheets to be created - A,B,C

.Example
    PS C:\> $doc = Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx

    PS C:\> Add-SLWorkSheet -WorkBookInstance $doc -worksheetName A,B,C -Verbose | Save-SLDocument
    VERBOSE: A : Adding worksheet to excel document - MyFirstDoc...
    VERBOSE: B : Adding worksheet to excel document - MyFirstDoc...
    VERBOSE: C : Adding worksheet to excel document - MyFirstDoc...
    VERBOSE: All Done!

    Description
    -----------
    An instance of MyFirstDoc is stored in a variable named doc which is then passed as a named parameter to Add-slworksheet.



.Example
    PS C:\> Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx | Add-SLWorkSheet -worksheetName A,B,C -Verbose | Save-SLDocument
    VERBOSE: A workSheet Named : 'A' Already exists so creating a worksheet named A9
    VERBOSE: A9 : Adding worksheet to excel document - MyFirstDoc...
    VERBOSE: A workSheet Named : 'B' Already exists so creating a worksheet named B75
    VERBOSE: B75 : Adding worksheet to excel document - MyFirstDoc...
    VERBOSE: A workSheet Named : 'C' Already exists so creating a worksheet named C82
    VERBOSE: C82 : Adding worksheet to excel document - MyFirstDoc...
    VERBOSE: All Done!

    Description
    -----------
    Demonstrates what happens when a user specifies a worksheetname that already exists in the workbook.

.Example
    PS C:\> New-SLDocument -WorkbookName "Test" -WorksheetName "Sheet1" -Path D:\ps\Excel -PassThru -Verbose | Add-SLWorkSheet -worksheetName A,B,C -Verbose | Save-SLDocument
    VERBOSE: New document has been created at :	D:\ps\Excel\Test.xlsx
    VERBOSE: A : Adding worksheet to excel document - Test...
    VERBOSE: B : Adding worksheet to excel document - Test...
    VERBOSE: C : Adding worksheet to excel document - Test...
    VERBOSE: All Done!

    Description
    -----------
    New document named test is created and then Add-slworksheet is called to add additional worksheets A,B,C.
    Note: 'Passthru' parameter with New-SLDocument is required when you want to pass the document instance to the pipeline if not no object is passed.

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
        [parameter(Mandatory = $true, Position = 0, ValueFromPipeLine = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [parameter(Mandatory = $true, Position = 1, ValueFromPipeLine = $false)]
        [string[]]$worksheetName


    )

    PROCESS
    {
        Foreach ($w in $worksheetName)
        {
            if ($WorkBookInstance.GetSheetNames() -contains $w)
            {
                $Random = Get-Random -Maximum 100
                Write-Verbose ("Add-SLWorkSheet :`tA workSheet Named '{0}' already exists so creating a new worksheet named '{1}' " -f $w, ($w + $Random))
                $w = $w + $Random
            }
            Write-Verbose ("Add-SLWorkSheet : '{0}' :`tAdding worksheet to excel document - '{1}'..." -f $w, $($WorkBookInstance.WorkbookName) )
            $WorkBookInstance.AddWorksheet($w) | Out-Null
        }

        $WorkBookInstance | Add-Member NoteProperty WorksheetNames @($WorkBookInstance.GetSheetNames()) -Force -PassThru |
            Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

    }
    END
    {

        Write-Verbose 'Add-SLWorkSheet : All Done!'
    }

}
