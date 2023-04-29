Function Get-SLDocument  {


    <#

.SYNOPSIS
    Gets an excel Document from the specified path and creates an instance of it for editing.

.DESCRIPTION
    Gets an excel Document from the specified path and creates an instance of it for editing.

.PARAMETER WorksheetName
    If specified the worksheet will be selected for editing.

.PARAMETER Path
    Path where the document is located. Specify the complete path along with the file extension.


.Example
    PS C:\> $Doc = Get-SLDocument -Path D:\PS\Excel\MyFirstDoc.xlsx
    PS C:\> $Doc

    WorkbookName         : MyFirstDoc
    WorksheetName        : {Test1, Test2}
    CurrentWorksheetName : Test2
    Path                 : D:\PS\Excel\MyFirstDoc.xlsx
    DocumentProperties   : SpreadsheetLight.SLDocumentProperties

    Description
    -----------
    Gets an instance of the document named 'MyFirstDoc' and stores it in a variable named Doc.The last worksheet is made active and selected for editing.
    Note: When a workbook contains multiple worksheets and you dont specify a worksheetname
          by default the last worksheet will be selected as the active or current worksheet.

.Example
    PS C:\> $Doc = Get-SLDocument -Path D:\PS\Excel\MyFirstDoc.xlsx -WorksheetName Test1
    PS C:\> $Doc


    WorkbookName         : MyFirstDoc
    WorksheetName        : {Test1, Test2}
    CurrentWorksheetName : Test1
    Path                 : D:\PS\Excel\MyFirstDoc.xlsx
    DocumentProperties   : SpreadsheetLight.SLDocumentProperties

    Description
    -----------
    Gets an instance of the document named 'MyFirstDoc' and stores it in a variable named Doc.
    since test1 is passed as a value to the worksheetname parameter, it is selected as the active worksheet.



.Example
    PS C:\> $Doc = Get-SLDocument  D:\PS\Excel\MyFirstDoc.xlsx  Test1


    Description
    -----------
    Positional parameters.



.Example
    PS C:\> dir d:\excel

    Mode                LastWriteTime     Length Name
    ----                -------------     ------ ----
    -a---          6/4/2014   3:53 PM          0 contacts.txt
    -a---          6/4/2014   3:53 PM          0 image1.bmp
    -a---          6/3/2014   5:06 PM       4628 MyFirstDoc.xlsx
    -a---          6/3/2014   3:23 PM       4626 PositionalWorkbook.xlsx



    PS C:\> $doc = dir d:\excel -Filter myfirstdoc.xlsx | Get-SLDocument
    PS C:\> $Doc


    WorkbookName         : MyFirstDoc
    WorksheetName        : {Sheet1}
    CurrentWorksheetName : Sheet1
    Path                 : D:\excel\MyFirstDoc.xlsx
    DocumentProperties   : SpreadsheetLight.SLDocumentProperties


    Description
    -----------
    Use dir to get the required excel document and then pipe it to Get-SLDcoument and work with it.
    Note: you can pass in more than one document using dir. The result will be an array which be enumerated easily using foreach or foreach-object.


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

        [Alias('FullName')]
        [parameter(Mandatory = $true, Position = 0, ValueFromPipeLine = $true, ValueFromPipeLineByPropertyName = $true)]
        [String]$Path,

        [parameter(Mandatory = $false, Position = 1)]
        [String]$WorkSheetName

    )
    BEGIN
    {

    }
    PROCESS
    {
        if (-not (Test-Path $Path))
        {
            Write-Warning ("Get-SLDocument :`tCould Not Find a Workbook at the path specified '{0}'" -f $path)
            break
        }
        $WorkBookInstance = New-Object SpreadsheetLight.SLDocument($path)

        if ($worksheetName)
        {
            if ($WorkBookInstance.GetSheetNames() -contains $WorksheetName)
            {
                $WorkBookInstance.SelectWorkSheet($WorksheetName) | Out-Null
            }
            Else
            {
                Write-Warning ("Get-SLDocument :`tCould Not Find a workSheet Named '{0}'. Current worksheet is '{1}'" -f $WorksheetName, $WorkBookInstance.GetCurrentWorksheetName())
            }
        }


        $filestats = Get-Item $Path
        $sheetnames = New-Object System.Collections.ArrayList
        $sheetnames.AddRange($WorkBookInstance.getsheetnames()) | Out-Null


        $WorkBookInstance | Add-Member NoteProperty WorkbookName $filestats.BaseName
        $WorkBookInstance | Add-Member NoteProperty WorksheetNames $sheetnames
        $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName()
        #$WorkBookInstance | Add-Member ScriptProperty WorksheetStatistics     {$this.GetWorksheetStatistics()}
        $WorkBookInstance | Add-Member NoteProperty Path $path

        Write-Output $WorkBookInstance


    }# Process

    END
    {

    }

}
