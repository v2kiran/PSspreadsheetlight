Function Remove-SLWorkSheet  {


    <#

.SYNOPSIS
    Deletes one or more worksheets from an Excel Document.

.DESCRIPTION
    Deletes one or more worksheets from an Excel Document.If the specified worksheet to be added already exists in the workbook,
    a random number is appended to the worksheetname to make it unique and then created as a worksheet.
    Since delete is destructive operation, the document is backed up before any work is done on it.
    The backup location defaults to the current user's temp directory
    On my system the backup location is as follows:
        'C:\Users\kiran\AppData\Local\Temp\PowerPSExcel'


.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet(s) to be deleted.
    You can delete more than one worksheet by specifying the names as a comma separated list.



.Example
    PS C:\> Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx | Remove-SLWorkSheet -worksheetName A,B,C -Verbose |  Save-SLDocument
    VERBOSE: A : Deleting worksheet from Workbook MyFirstDoc
    VERBOSE: B : Deleting worksheet from Workbook MyFirstDoc
    VERBOSE: C : Deleting worksheet from Workbook MyFirstDoc
    VERBOSE: Performing CleanUP...
    VERBOSE: All Done!

    Description
    -----------
    MyfirstDoc is piped to Remove-SLWorkSheet with 3 worksheets to be deleted - A,B,C

.Example
    PS C:\> $doc = Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx

    PS C:\> Remove-SLWorkSheet -WorkBookInstance $doc -worksheetName A,B,C -Verbose | Save-SLDocument
    VERBOSE: A : Deleting worksheet to excel document MyFirstDoc
    VERBOSE: B : Deleting worksheet to excel document MyFirstDoc
    VERBOSE: C : Deleting worksheet to excel document MyFirstDoc
    VERBOSE: All Done!

    Description
    -----------
    An instance of MyFirstDoc is stored in a variable named doc which is then passed as a named parameter to Remove-SLWorkSheet.



.Example
    PS C:\> Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx | Remove-SLWorkSheet -worksheetName A,B,DoesNotExist -Verbose | Save-SLDocument
    VERBOSE: A : Deleting worksheet from Workbook MyFirstDoc
    VERBOSE: B : Deleting worksheet from Workbook MyFirstDoc
    WARNING: Could Not Find workSheet Named : 'DoesNotExist' .
    VERBOSE: Performing CleanUP...
    VERBOSE: All Done!

    Description
    -----------
    Demonstrates what happens when a user specifies a worksheetname that does not exist in the workbook.

.Example
    PS C:\> $doc = Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $WorksheetNames = $doc | List-SLWorkSheet -filter "a|b|c"
    PS C:\> Remove-SLWorkSheet -WorkBookInstance $Doc -worksheetName $WorksheetNames -Verbose | Save-SLDocument
    VERBOSE: A : Deleting worksheet from Workbook MyFirstDoc
    VERBOSE: B : Deleting worksheet from Workbook MyFirstDoc
    VERBOSE: C : Deleting worksheet from Workbook MyFirstDoc
    VERBOSE: Performing CleanUP...
    VERBOSE: All Done!

    Description
    -----------
    List-slworksheet is used to filter out the worksheet names and stored in a variable named 'Worksheetnames'.
    Delete-sldocument is then called with anmed parameters


.INPUTS
   String[],SpreadsheetLight.SLDocument

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
        [string[]]$WorkSheetName



    )
    BEGIN
    {
        #   need to add a temp worksheet because the currently selected worksheet cannot be deleted
        $random = Get-Random -Maximum 100

    }
    PROCESS
    {

        if ($WorkBookInstance.GetSheetNames() -notcontains $random )
        {

            $WorkBookInstance.AddWorksheet($random) | Out-Null
        }



        #Backup-SLDocument -WorkBookInstance $WorkBookInstance
        Foreach ($w in $worksheetName)
        {
            if ($WorkBookInstance.GetSheetNames() -contains $w )
            {
                if ($w -ne $random)
                {

                    Write-Verbose ('Remove-SLWorkSheet : {0} : Deleting worksheet from Workbook {1}' -f $w, $WorkBookInstance.workbookname )
                    $WorkBookInstance.DeleteWorkSheet($w) | Out-Null
                    $processfile = $true
                }

            }
            Else
            {
                Write-Warning ("Remove-SLWorkSheet :`tCould Not Find workSheet Named '{0}'. No Action Taken  " -f $w)
                $processfile = $false
            }
        }# foreach worksheet

        If ($processfile)
        {
            # Clean-UP: delete the temp sheet that was added in the PROCESS block
            Write-Verbose 'Remove-SLWorkSheet : Performing CleanUP...'
            $existingsheet = $WorkBookInstance.GetSheetNames() | Select-Object -First 1
            $WorkBookInstance.SelectWorksheet($existingsheet) | Out-Null
            if ($random -ne $null) { $WorkBookInstance.DeleteWorksheet($random) | Out-Null }
        }
        $WorkBookInstance | Add-Member NoteProperty WorksheetNames @($WorkBookInstance.GetSheetNames()) -Force -PassThru |
            Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

    }
    END
    {

        Write-Verbose 'Remove-SLWorkSheet : All Done!'
    }

}
