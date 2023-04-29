Function Import-CSVToSLDocument  {


    <#

.SYNOPSIS
    Import one or more CSV files into an excel document.

.DESCRIPTION
    Import one or more CSV files into an excel document.Worksheet names are automatically generated based on the csv filename.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER CSVFile
    The complete path to the CSV file along with extension.Multiple paths can be specified in a comma seperated list.

.PARAMETER ImportStartCell
    Marks the start of the cell within excel for the csvfile data.


.PARAMETER Force
    Use force to overwrite an existing worksheet.

.PARAMETER AutofitColumns
    Autofit Columns.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> Import-CSVToSLDocument -WorkBookInstance $doc -CSVFile "D:\workdocuments\CSVFiles\Disk\disk.csv" -AutofitColumns -Verbose | Save-SLDocument

    Description
    -----------
    An instance of MyFirstDoc is stored in a variable named doc.
    CSVfile named 'disk' is then imported into the existing document.


.Example
    PS C:\> $doc = New-SLDocument -WorkbookName Disk -Path D:\ps\Excel -PassThru
    PS C:\> Import-CSVToSLDocument -WorkBookInstance $doc -CSVFile "D:\workdocuments\CSVFiles\Disk\disk.csv" -AutofitColumns -Verbose | Save-SLDocument

    Description
    -----------
    Same as the first example except that in this case we create a new document named disk and then import the contents of the csv file.
    Note the use of the 'Passthru' parameter with New-SLDcoument. Required if you want to process the new document further.


.Example
    PS C:\> $doc = Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $CSVFIle_Disk    = "D:\workdocuments\CSVFiles\Disk\disk.csv"
    PS C:\> $CSVFIle_Service = "D:\workdocuments\CSVFiles\Service\Service.csv"
    PS C:\> Import-CSVToSLDocument -WorkBookInstance $doc -CSVFile $CSVFIle_Disk,$CSVFIle_Service -AutofitColumns -Verbose
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    An instance of MyFirstDoc is stored in a variable named doc.
    2 csv files named disk & service in different locations are passed in a comma seperated list
    Since we are importing the csv files to one single document we cannot call save-sldocument in the same pipeline.
    Instead we perform all the operations on the doucment and then finally call save to save all changes.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> dir D:\workdocuments -Filter *.csv -Exclude process.csv -Recurse | Import-CSVToSLDocument -WorkBookInstance $doc -AutofitColumns  -Verbose
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    We make use of the pipeline to pipe the CSVfile paths using 'DIR' along with the filter parameter.
    We are getting all csv files from the location d:\workdocuments with the exclusion of 'process.csv" and then pipe them Import-CSVToSLDocument.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> dir D:\workdocuments -Filter *.csv -Exclude process.csv -Recurse | Import-CSVToSLDocument -WorkBookInstance $doc -AutofitColumns  -Verbose | Set-SLTableStyle -TableStyle Dark10 -Verbose
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    We build on the last example by importing the csvfiles and additionally setting a tablestyle of 'Dark10' on each imported file.

.Example
    PS C:\> dir D:\workdocuments -Filter *.csv -Exclude process.csv -Recurse |
                ForEach-Object {Import-CSVToSLDocument -CSVFile $_.FullName -WorkBookInstance (New-SLDocument -WorkbookName $_.BaseName -Path D:\ps\Excel -PassThru) -AutofitColumns  -Verbose } |
                    Set-SLTableStyle -TableStyle Dark10 -Verbose |
                        Save-SLDocument

    Description
    -----------
    A oneliner to import all csv files with the exception of process.csv to new excel workbooks and also set a tablestyle of 'Dark10' on each imported csvfile.
    Since we are exporting the csvfiles to different workbooks we can use 'Save-SLDocument' as the last command on the pipeline.

.INPUTS
   String,SpreadsheetLight.SLDocument

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    N/A

#>


    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param (

        [ValidatenotnullorEmpty()]
        [parameter(Mandatory = $true, position = 1, valuefrompipeline = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [parameter(Mandatory = $true, Position = 0, ValueFromPipeLineByPropertyName = $true)]
        [Alias('FullName')]
        [String[]]$CSVFile,

        [parameter(Mandatory = $false, ValueFromPipeLineByPropertyName = $true)]
        [String]$ImportStartCell = 'B2',

        [parameter(Mandatory = $false)]
        [System.UInt32[]]$SkipColumns,

        [Switch]$Force,

        [Switch]$AutofitColumns = $true

    )
    BEGIN
    {
        $ImportOptions = New-Object Spreadsheetlight.SLTextImportOptions
        $ImportOptions.UseCommaDelimiter = $true
    }
    PROCESS
    {
        Backup-SLDocument -WorkBookInstance $WorkBookInstance
        Foreach ($CSV in $CSVFile)
        {
            $filestats = Get-Item $CSV
            $worksheetName = $filestats.basename
            Write-Verbose ("Import-CSVToSLDocument :`tProcessing CSV FIle '{0}'..." -f $filestats.Name)

            if ($SkipColumns)
            {
                foreach ($SkipColumn in $SkipColumns)
                {
                    Write-Verbose ("Import-SLTextFile :`tSkipping Import of columnID '{0}'..." -f $SkipColumn)
                    $ImportOptions.SkipColumn($SkipColumn)
                }
            }

            if ($WorkBookInstance.GetSheetNames() -notcontains $WorksheetName)
            {

                $ProcessFile = $true
                $WorkBookInstance.AddWorksheet($WorksheetName) | Out-Null
                $WorkBookInstance.ImportText($CSV, $ImportStartCell, $ImportOptions)
            }
            Else
            {
                if ($force -and $PSCmdlet.ShouldProcess($WorksheetName, 'OVERWRITE Worksheet'))
                {
                    $ProcessFile = $true
                    $WorkBookInstance.SelectWorkSheet($WorksheetName) | Out-Null
                    $WorkBookInstance.ImportText($CSV, $ImportStartCell, $ImportOptions)
                }
                Else
                {
                    $ProcessFile = $false
                    Write-Warning ("Import-CSVToSLDocument :`tSpecified Worksheet '{0}' already Exists. Please select a different name or use the '-Force' parameter to overwrite" -f $worksheetName)
                }
            }


            If ($ProcessFile)
            {
                ## Add Autofit to the Table - Optional
                if ($AutofitColumns)
                {
                    $WorkBookInstance.autofitcolumn('A', 'DD')
                }

                $stats = $WorkBookInstance.GetWorksheetStatistics()
                $Range = Convert-ToExcelRange -StartRowIndex $stats.StartRowIndex -StartColumnIndex $stats.StartColumnIndex -EndRowIndex $stats.ENDRowIndex -EndColumnIndex $stats.ENDColumnIndex

                $WorkBookInstance | Add-Member NoteProperty StartRowIndex $stats.StartRowIndex -Force
                $WorkBookInstance | Add-Member NoteProperty StartColumnIndex $stats.StartColumnIndex -Force
                $WorkBookInstance | Add-Member NoteProperty EndRowIndex $stats.ENDRowIndex -Force
                $WorkBookInstance | Add-Member NoteProperty EndColumnIndex $stats.ENDColumnIndex -Force
                $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
                $WorkBookInstance | Add-Member NoteProperty WorksheetNames $WorkBookInstance.GetSheetNames() -Force
                $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force
                Write-Output $WorkBookInstance

            }#If processFile
        }#Foreach CSVFile
    }#process

    END
    {

    }

}
