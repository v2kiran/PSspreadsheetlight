Function New-SLDocument  {

    <#

.SYNOPSIS
    Creates a new Excel Document or an instance of an excel document that can be piped to other commands in the module.

.DESCRIPTION
    Creates a new Excel Document or an instance of an excel document that can be piped to other commands in the module.
    2 parametersets included, the simplest is with no parameters which outputs a new instance of an excel
    document which can be piped to other cmdlets in the module.
    The 2nd parameterset allows you to name the workbook and optionally the worksheet.
    Note:Passthru parameter must be used if you want to pipe the instance to another cmdlet.


.PARAMETER WorkbookName
    Name of Excel Document to be created. There is no need to specify the .xlsx extension.

.PARAMETER WorksheetName
    Name of the Worksheet to be created.
    Only one worksheet can be created. If you want to create more, create a new instance of excel and then pipe that to Add-SLWorksheet

.PARAMETER Path
	Path where the excel document is to be created. You may specify a partial path without the filename or file extension.

.PARAMETER Force
	Use this to Overwrite an existing file.

.PARAMETER PassThru
	Use this parameter to pass the newly created document to the next cmdlet on the pipeline.


.Example
    PS C:\> New-SLDocument -WorkbookName MyFirstDoc -WorksheetName Testsheet1 -Path D:\PS\Excel  -Verbose

    VERBOSE: New document has been created at :	D:\PS\Excel\MyFirstDoc.xlsx

    Description
    -----------
    Creates a new excel document with a blank worksheet named Testsheet1

.Example
    PS C:\> New-SLDocument -WorkbookName MyFirstDoc -WorksheetName Overwritten -Path D:\PS\Excel  -Verbose -Force

    Confirm
    Are you sure you want to perform this action?
    Performing operation "OVERWRITE FILE" on Target "D:\PS\Excel\MyFirstDoc.xlsx".
    [Y] Yes  [A] Yes to All  [N] No  [L] No to All  [S] Suspend  [?] Help (default is "Y"): y

    VERBOSE: Force Switch specified. overwriting existing file
    VERBOSE: New document has been created at : D:\PS\Excel\MyFirstDoc.xlsx

    Description
    -----------
    Overwrites the excel file created in example 1 above via the use of the 'Force' parameter.
    Note: User will be prompted to confirm the 'Overwrite' action. The action will only succeed with an input of either 'Y' or 'A'


.Example
    PS C:\> New-SLDocument -WorkbookName MyFirstDoc -WorksheetName Overwritten -Path D:\PS\Excel  -Verbose -Force -Confirm:$false

    VERBOSE: Performing operation "OVERWRITE FILE" on Target "D:\PS\Excel\MyFirstDoc.xlsx".
    VERBOSE: Force Switch specified. overwriting existing file
    VERBOSE: New document has been created at : D:\PS\Excel\MyFirstDoc.xlsx

    Description
    -----------
    Overwrites the existing file but does not prompt the user for confirmation because -confirm is set to 'False', default is true.

.Example
    PS C:\> $Doc =  New-SLDocument -WorkbookName Passthru -WorksheetName Test1 -Path D:\PS\Excel  -PassThru
    PS C:\> $Doc


    WorkbookName         : Passthru
    WorksheetName        : {Test1}
    CurrentWorksheetName : Test1
    Path                 : D:\PS\Excel\Passthru.xlsx
    DocumentProperties   : SpreadsheetLight.SLDocumentProperties

    Description
    -----------
    creates a new document and stores an instance of that document in the variable $Doc which can be piped to other commands.

.Example
    PS C:\> New-SLDocument  PositionalWorkbook    Positionalworksheet    D:\PS\Excel  -Verbose

    VERBOSE: New document has been created at :	D:\PS\Excel\PositionalWorkbook.xlsx

    Description
    -----------
    Create a new document using positional parameters.

.Example
    PS C:\> $Doc = New-SLDocument

    PS C:\> $Doc

    WorksheetName                                                CurrentWorksheetName                                         DocumentProperties
    -------------                                                --------------------                                         ------------------
    {Sheet1}                                                     Sheet1                                                       SpreadsheetLight.SLDocumentProperties


    Description
    -----------
    Creates a new instance of an excel document which can be piped to other commands such as Add-SLWorksheet.
    This is the simplest and most useful way to use the cmdlet.

.INPUTS
   String

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    N/A

#>



    [CmdletBinding(DefaultParameterSetName = 'None', SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    [OutputType([SpreadsheetLight.SLDocument])]
    param (
        [parameter(Mandatory = $True, Position = 0, ParameterSetName = 'Named')]
        [String]$WorkbookName,

        [parameter(Mandatory = $false, Position = 1, ParameterSetName = 'Named')]
        [String]$WorksheetName,

        [parameter(Mandatory = $True, Position = 2, ParameterSetName = 'Named')]
        [String]$Path,

        [parameter(ParameterSetName = 'Named')]
        [Switch]$Force,

        [parameter(ParameterSetName = 'Named')]
        [Switch]$PassThru

    )
    PROCESS {}
    END
    {

        if ($PSCmdlet.ParameterSetName -eq 'Named')
        {

            $WorkBookInstance = New-Object SpreadsheetLight.SLDocument

            #Set the file path
            If ($WorkbookName -match '.xlsx')
            {
                $Fullpath = Join-Path $path $WorkbookName
            }
            Else
            {
                $Fullpath = (Join-Path $path $WorkbookName) + '.xlsx'
            }

            # If Parameter worksheetname is mentioned create workbook with the specified worksheetname
            if ($WorksheetName)
            {
                # Create Workbook with specified workbookname and worksheetname
                $WorkBookInstance.RenameWorksheet([SpreadsheetLight.SLDocument]::DefaultFirstSheetName, $WorksheetName) | Out-Null
            }

            if (Test-Path $Fullpath)
            {
                if ($Force -and $PSCmdlet.ShouldPROCESS($Fullpath, 'OVERWRITE FILE') )
                {
                    Write-Verbose ("New-SLDocument :`tForce Switch specified. overwriting existing file located at '{0}'" -f $Fullpath)
                    $WorkBookInstance.SaveAs($Fullpath)

                    if (Test-Path $Fullpath)
                    {
                        $IsFileCreated = $true
                        Write-Verbose ("New-SLDocument :`tNew document has been created at '{0}'" -f $Fullpath)
                    }
                    Else
                    {
                        Write-Warning ("New-SLDocument :`tFailed to create a new document at '{0}'" -f $Fullpath)
                    }
                }
                else
                {
                    Write-Warning ("New-SLDocument :`tSpecified WorkbookName '{0}' already exists at '{1}'. Use the '-Force' parameter to overwrite" -f $WorkbookName, $Fullpath)
                    $IsFileCreated = $false
                    $WorkBookInstance.Dispose()
                }

            }
            else
            {
                # Save the document
                $WorkBookInstance.SaveAs($Fullpath)
                if (Test-Path $Fullpath)
                {
                    $IsFileCreated = $true
                    Write-Verbose ("New-SLDocument :`tNew document has been created at '{0}'" -f $Fullpath)
                }
                Else
                {
                    Write-Warning ("New-SLDocument :`tFailed to create a new document at '{0}'" -f $Fullpath)
                }
            }

            if ($PassThru -and $IsFileCreated)
            {
                $WorkBookInstance = New-Object SpreadsheetLight.SLDocument($Fullpath)

                # Generate Worksheet statistics
                $sheetnames = New-Object System.Collections.ArrayList
                $sheetnames.addrange($WorkBookInstance.GetSheetNames()) | Out-Null
                $filestats = Get-Item $Fullpath

                # Add properties to sldocument
                $WorkBookInstance | Add-Member NoteProperty WorkbookName $filestats.BaseName
                $WorkBookInstance | Add-Member NoteProperty WorksheetNames $sheetnames
                $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName()
                $WorkBookInstance | Add-Member NoteProperty Path $Fullpath
                Write-Output $WorkBookInstance

            }


        }# Parameterset Named


        if ($PSCmdlet.ParameterSetName -eq 'None')
        {
            $WorkBookInstance = New-Object SpreadsheetLight.SLDocument

            # Generate Worksheet statistics
            $sheetnames = New-Object System.Collections.ArrayList
            $sheetnames.addrange($WorkBookInstance.GetSheetNames()) | Out-Null

            # Add properties to sldocument
            $WorkBookInstance | Add-Member NoteProperty WorkbookName 'Book1'
            $WorkBookInstance | Add-Member NoteProperty WorksheetNames $sheetnames
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName()
            Write-Output $WorkBookInstance
        }


    }#END


}
