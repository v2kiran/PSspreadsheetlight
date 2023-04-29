# Create Type Accelerators
[System.Type]$typeAcceleratorsType = [System.Management.Automation.PSObject].Assembly.GetType('System.Management.Automation.TypeAccelerators', $true, $true)
$typeAcceleratorsType::Add('OLPattern', [DocumentFormat.OpenXml.Spreadsheet.PatternValues])
$typeAcceleratorsType::Add('SLThemeColor', [SpreadsheetLight.SLThemeColorIndexValues])
$typeAcceleratorsType::Add('SLGradient', [SpreadsheetLight.SLGradientShadingStyleValues])
$typeAcceleratorsType::Add('SLCFRangeValues', [SpreadsheetLight.SLConditionalFormatRangeValues])
$typeAcceleratorsType::Add('SLCFMinMax', [SpreadsheetLight.SLConditionalFormatMinMaxValues])
$typeAcceleratorsType::Add('SLCFColorScale', [SpreadsheetLight.SLConditionalFormatColorScaleValues])
$typeAcceleratorsType::Add('OLIconSetValues', [DocumentFormat.OpenXml.Spreadsheet.IconSetValues])
# edit set-slcellvalue



Function New-SLDocument
{
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

Function Get-SLDocument
{

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

Function Save-SLDocument
{

    <#

.SYNOPSIS
    Saves an instance of an excel document.

.DESCRIPTION
    Saves an instance of an excel document.
    2 parametersets included. You can either save changes to an existing document or save to a new path.

.PARAMETER WorkbookName
    Name of Excel Document to be saved.

.PARAMETER Path
	Path where the excel document is to be saved. You have to specify the complete path along with the file extension.

.PARAMETER Force
	Use this to Overwrite an existing file at the path mentioned in the path parameter.


.Example
    PS C:\> Get-SLDocument -path D:\PS\Excel\MyFirstDoc.xlsx | Save-SLDocument -Verbose

    VERBOSE: Document has been Saved

    Description
    -----------
    Saves the excel document MyFirstDoc.

.Example
    PS C:\> Get-SLDocument -path D:\PS\Excel\MyFirstDoc.xlsx | Save-SLDocument -path D:\PS\Excel\MyFirstDoc-Duplicate.xlsx  -Verbose

    VERBOSE: Document has been Saved to :	D:\PS\Excel\MyFirstDoc-Duplicate.xlsx

    Description
    -----------
    Save the document MyFirstDoc as MyFirstDoc-Duplicate .


.Example
    PS C:\> Get-SLDocument -path D:\PS\Excel\MyFirstDoc.xlsx | Save-SLDocument -path D:\PS\Excel\MyFirstDoc-Duplicate.xlsx  -force -Verbose

    VERBOSE: Performing operation "OVERWRITE FILE" on Target "D:\PS\Excel\MyFirstDoc-Duplicate.xlsx".
    VERBOSE: Force Switch specified. overwriting existing file
    VERBOSE: Document has been Saved to :	D:\PS\Excel\MyFirstDoc-Duplicate.xlsx

    Description
    -----------
    Here we use the force switch to overwrite the existing file 'MyFirstDoc-Duplicate' which we created in the previous example.


.INPUTS
   String

.OUTPUTS
   No Output

.Link
    N/A

#>



    [CmdletBinding(DefaultParameterSetName = 'None', SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param (
        [parameter(Mandatory = $true, Position = 0, ValueFromPipeLine = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [parameter(Mandatory = $True, Position = 1, ParameterSetName = 'Path')]
        [String]$Path,

        [parameter(Position = 2, ParameterSetName = 'Path')]
        [Switch]$Force

    )
    PROCESS
    {
        If ($PSCmdlet.ParameterSetName -eq 'Path')
        {
            if (Test-Path $Path)
            {
                if ($Force -and $PSCmdlet.ShouldPROCESS($path, 'OVERWRITE FILE') )
                {
                    Write-Verbose ("Save-SLDocument :`tForce Switch specified. Overwriting existing file at '{0}'" -f $Path)
                    $WorkBookInstance.SaveAs($Path)
                    $IsFileSaved = $true
                }
                else
                {
                    Write-Warning ("Save-SLDocument :`tFile already exists at '{0}'. Use the -Force Parameter to overwrite" -f $Path)
                    $IsFileSaved = $false
                }
            }
            else
            {
                # Save the document
                $WorkBookInstance.SaveAs($path)
                $IsFileSaved = $true
                Write-Verbose ("Save-SLDocument :`tDocument has been Saved to '{0}'" -f $path)
            }

        }#Parametersetname path

        ## Parametersetname 'None'
        if ($PSCmdlet.ParameterSetName -eq 'None')
        {
            $WorkBookInstance.Save()
            Write-Verbose ("Save-SLDocument :`tDocument has been Saved")
        }

    }#process
}




Function Select-SLWorkSheet
{

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

Function List-SLWorkSheet
{

    <#

.SYNOPSIS
    List all worksheets contained in a workbook.

.DESCRIPTION
    List all worksheets contained in a workbook.
    This is intended to quickly provide a way for the user to determine what worksheets are contained in a workbook
    Note: The workbook instance is not passed through.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER Path
    The path to the excel document.

.PARAMETER Filter
    Simple Filter that supports regex matches.


.Example
    PS C:\> List-SLWorkSheet -Path D:\ps\Excel\MyFirstDoc.xlsx

    sheet1
    A
    B
    C
    D
    Sheet5

    Description
    -----------
    Lists all worksheets contained in the document named 'MyFirstDoc'


.Example
    PS C:\> $doc = Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | List-SLWorkSheet


    sheet1
    A
    B
    C
    D
    Sheet5


    Description
    -----------
    Get-SLDocument is used to get an instance of the document named 'MyFirstDoc' which is then piped to List-SLWorkSheet

.Example
    PS C:\> $doc = Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | List-SLWorkSheet -Filter "a|b"


    A
    B


    Description
    -----------
    List all worksheets that match either a or b.

.Example
    PS C:\> $doc = Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | List-SLWorkSheet -Filter "[^a|b]"

    sheet1
    C
    D
    Sheet5


    Description
    -----------
    List all worksheets that DONT match either a or b.

.Example
    PS C:\> $doc = Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | List-SLWorkSheet -Filter "sheet"

    sheet1
    Sheet5


    Description
    -----------
    List all worksheets that have the word 'sheet'.


.INPUTS
   SpreadsheetLight.SLDocument

.OUTPUTS
   String[]

.Link
    N/A
#>

    [CmdletBinding(defaultparametersetname = 'Instance')]
    [OutputType([string[]])]
    param (
        [parameter(Mandatory = $true, Position = 0, ValueFromPipeLine = $true, ParameterSetName = 'Instance')]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [parameter(Mandatory = $true, Position = 1, ValueFromPipeLine = $false, ParameterSetName = 'path')]
        [String]$Path,

        [parameter(Mandatory = $false, Position = 2, ValueFromPipeLine = $false)]
        [String]$Filter

    )
    PROCESS
    {

        If ($PSCmdlet.ParameterSetName -eq 'path')
        {
            $WorkBookInstance = New-Object SpreadsheetLight.SLDocument($path)
            if ($filter)
            {
                $WorkBookInstance.GetSheetNames() -match $Filter
            }
            Else
            {
                $WorkBookInstance.GetSheetNames()
            }
            $WorkBookInstance.Dispose() | Out-Null
        }

        If ($PSCmdlet.ParameterSetName -eq 'Instance')
        {
            if ($filter)
            {
                $WorkBookInstance.GetSheetNames() -match $Filter
            }
            Else
            {
                $WorkBookInstance.GetSheetNames()
            }
        }
    }
}






Function Add-SLWorkSheet
{

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



Function Remove-SLWorkSheet
{

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



Function Copy-SLWorkSheet
{

    <#

.SYNOPSIS
    Copies or duplicates a worksheet.

.DESCRIPTION
    Copies or duplicates a worksheet.Only one worksheet can be copied at a time. ,
    If you want to copy multiple worksheets use a looping construct.


.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet to be copied.

.PARAMETER NewworksheetName
    Name of the target or duplicate worksheet.

.PARAMETER Force
    Use force to overwrite an existing worksheet.

.Example
    PS C:\> Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx | Copy-SLWorkSheet -worksheetName Sheet2 -NewworksheetName Sheet2_Copy -Verbose | Save-SLDocument
    VERBOSE: Sheet2 : Worksheet is being copied to NewWorkSheetName : Sheet2_Copy
    VERBOSE: Performing CleanUP...
    VERBOSE: All Done!

    Description
    -----------
    An instance of MyfirstDoc is piped to Copy-SLWorkSheet with sheet2 being copied Sheet2_Copy.

.Example
    PS C:\> $doc = Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx

    PS C:\> Copy-SLWorkSheet -WorkBookInstance $doc -worksheetName Sheet2 -NewworksheetName Sheet2_Copy -Verbose | Save-SLDocument
    WARNING: sheet2_copy : Exists. No Action Taken.. Specify '-Force' parameter to overwrite..
    VERBOSE: Performing CleanUP...
    VERBOSE: All Done!

    Description
    -----------
    Demonstrates what happens when a user specifies a Newworksheetname that already exist's in the workbook.
    The target worksheet is not overwritten and instead you get a warning asking you to specify the force parameter.



.Example
    $doc = Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> Copy-SLWorkSheet -WorkBookInstance $doc -worksheetName sheet2 -NewworksheetName sheet2_copy -Verbose -Force | Save-SLDocument

    Confirm
    Are you sure you want to perform this action?
    Performing operation "OVERWRITE FILE" on Target "sheet2_copy".
    [Y] Yes  [A] Yes to All  [N] No  [L] No to All  [S] Suspend  [?] Help (default is "Y"): y
    VERBOSE: sheet2_copy : NewworksheetName specified exists but since Force Option is specified the existing worksheet
    will be overwritten with contents from : sheet2
    VERBOSE: Performing CleanUP...
    VERBOSE: All Done!

    Description
    -----------
    Demonstrates the use of the force parameter to overwrite an existing worksheet.
    The action is completed only if the user supplies either an 'y' or 'A' at the prompt.
    Use with caution.





.INPUTS
   String,SpreadsheetLight.SLDocument

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    N/A

#>




    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    [OutputType([SpreadsheetLight.SLDocument])]
    param (
        [parameter(Mandatory = $true, Position = 0, ValueFromPipeLine = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [parameter(Mandatory = $true, Position = 1, ValueFromPipeLine = $false)]
        [string]$WorkSheetName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true)]
        [string]$NeWworkSheetName,

        [parameter(Mandatory = $false)]
        [Switch]$Force


    )
    BEGIN
    {
        Clear-Variable Random -ErrorAction silentlycontinue
    }
    PROCESS
    {
        if ($WorkBookInstance.GetSheetNames() -contains $NewworksheetName)
        {
            if ($Force -and $PSCmdlet.ShouldPROCESS($NewworksheetName, 'OVERWRITE FILE') )
            {

                #need to add a temp worksheet because the currently selected worksheet cannot be deleted
                $random = Get-Random -Maximum 100
                $WorkBookInstance.AddWorksheet($random) | Out-Null

                Write-Verbose ("Copy-SLWorkSheet :`tSpecified New WorksheetName '{0}' already exists but since Force Option is specified the existing worksheet will be overwritten with contents from '{1}'" -f $NewworksheetName, $worksheetName )
                $WorkBookInstance.CopyWorkSheet($worksheetName, $NewWorksheetName) | Out-Null
                $processfile = $true
            }
            Else
            {
                Write-Warning ("Copy-SLWorkSheet :`tSpecified New WorksheetName '{0}' already exists so No Action Taken. Use the '-Force' parameter to overwrite" -f $NewworksheetName )
                $processfile = $false
            }

        }
        Else
        {
            #Need to add a temp worksheet because the currently selected worksheet cannot be deleted
            $random = Get-Random -Maximum 100
            $WorkBookInstance.AddWorksheet($random) | Out-Null

            Write-Verbose ("Copy-SLWorkSheet :`tWorksheet '{0}' is being copied to NewWorkSheet '{1}'" -f $worksheetName, $NewworksheetName )
            $WorkBookInstance.CopyWorkSheet($worksheetName, $NewWorksheetName) | Out-Null
            $processfile = $true
        }

        If ($processfile)
        {
            #clean up- delete the temp sheet that was added in the PROCESS block
            Write-Verbose 'Copy-SLWorkSheet : Performing CleanUP...'
            $existingsheet = $WorkBookInstance.GetSheetNames() | Where-Object { $_ -ne $random } | Select-Object -First 1
            $WorkBookInstance.SelectWorksheet($existingsheet) | Out-Null
            if ($random -ne $null) { $WorkBookInstance.DeleteWorksheet($random) | Out-Null }
        }
        $WorkBookInstance | Add-Member NoteProperty WorksheetNames @($WorkBookInstance.GetSheetNames()) -Force -PassThru |
            Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

    }
    END
    {
        Write-Verbose 'Copy-SLWorkSheet : All Done!'
    }

}

Function Rename-SLWorkSheet
{

    <#

.SYNOPSIS
    Renames a worksheet.

.DESCRIPTION
    Renames a worksheet..Only one worksheet can be renamed at a time.
    If you want to rename multiple worksheets use a looping construct.


.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet to be renamed.

.PARAMETER NewworksheetName
    New worksheetname.


.Example
    PS C:\> Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx | Rename-SLWorkSheet -WorkSheetName sheet2 -NeWworkSheetName sheet2_renamed -Verbose | Save-SLDocument

    VERBOSE: sheet2 : Worksheet is being renamed to NewWorkSheetName : sheet2_renamed
    VERBOSE: All Done!

    Description
    -----------
    MyfirstDoc is piped to Rename-SLWorkSheet with sheet2 being renamed as Sheet2_renamed.


.INPUTS
   String,SpreadsheetLight.SLDocument

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

        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true)]
        [string]$NeWworkSheetName


    )
    PROCESS
    {
        if ($WorkBookInstance.GetSheetNames() -contains $WorkSheetName)
        {

            if ($WorkBookInstance.GetSheetNames() -contains $NewWorksheetName)
            {
                Write-Warning ("Rename-SLWorkSheet :`tNewWorksheet '{0}' already Exists.Re-try the command with a different name" -f $NewWorksheetName )
            }
            Else
            {


                $WorkBookInstance.RenameWorkSheet($worksheetName, $NewWorksheetName) | Out-Null
                Write-Verbose ("Rename-SLWorkSheet :`tWorksheet '{0}' has being renamed to NewWorkSheetName '{1}'" -f $worksheetName, $NewworksheetName )

                $WorkBookInstance | Add-Member NoteProperty WorksheetNames @($WorkBookInstance.GetSheetNames()) -Force -PassThru |
                    Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
            }
        }
        Else
        {
            Write-Warning ("Rename-SLWorkSheet :  :`tSpecified Worksheet '{0}' Could not be Found. Check the spelling and try again." -f $WorkSheetName )
        }
    }
    END
    {
        Write-Verbose 'Rename-SLWorkSheet : All Done!'
    }
}

Function Move-SLWorkSheet
{

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


Function Set-SLDocumentMetadata
{

    <#

.SYNOPSIS
    Set document metadata.

.DESCRIPTION
    Set document metadata that helps identify a document and also to organise them by tags,comment or author.


.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER Title
    The title of the document.

.PARAMETER Author
    The creator of the document.

.PARAMETER Comment
    The summary or abstract of the contents of the document.

.PARAMETER Tags
    A word or set of words describing the document.Refers to keywords in excel.

.PARAMETER Category
    The category of the document eg: personal,business,financial etc.

.PARAMETER LastModifiedBy
    The document is last modified by this person.

.PARAMETER Subject
    The topic of the document.


.Example
    PS C:\> Get-SLDocument C:\temp\test.xlsx | Set-SLDocumentMetadata -Title mydoc -Author kiran -Comment "this is a test doc" -Tags "test;document" | Save-SLDocument


    Description
    -----------
    Set the title,author,comment and tags properties on the document named test.
    Note:Tags are seperated by semicolons.


.INPUTS
   String,SpreadsheetLight.SLDocument

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    N/A

#>


    #>
    [CmdletBinding()]
    [OutputType([SpreadsheetLight.SLDocument])]
    param (
        [parameter(Mandatory = $true, Position = 0, ValueFromPipeLine = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [parameter(Mandatory = $false)]
        [string]$Title,

        [parameter(Mandatory = $false)]
        [string]$Author,

        [parameter(Mandatory = $false)]
        [string]$Comment,

        [parameter(Mandatory = $false)]
        [string]$Tags,

        [parameter(Mandatory = $false)]
        [string]$Category,

        [parameter(Mandatory = $false)]
        [string]$LastModifiedBy,

        [parameter(Mandatory = $false)]
        [string]$Subject
    )
    PROCESS
    {
        if ($Author) { $WorkBookInstance.DocumentProperties.Creator = $Author }
        if ($Title) { $WorkBookInstance.DocumentProperties.Title = $Title }
        if ($Comment) { $WorkBookInstance.DocumentProperties.Description = $Comment }
        if ($Tags) { $WorkBookInstance.DocumentProperties.Keywords = $Tags }
        if ($Category) { $WorkBookInstance.DocumentProperties.Category = $Category }
        if ($LastModifiedBy) { $WorkBookInstance.DocumentProperties.LastModifiedBy = $LastModifiedBy }
        if ($Subject) { $WorkBookInstance.DocumentProperties.Subject = $Subject }

        Write-Output $WorkBookInstance
    }
}





Function Hide-SLGridLines
{

    <#

.SYNOPSIS
    Removes Gridlines from one or more worksheets.

.DESCRIPTION
    Removes Gridlines from one or more worksheets.


.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet from which gridlines are to be removed.

.PARAMETER All
    If specified will remove gridlines from all worksheets within a workbook.
    The user does not need to use the worksheetname parameter in conjunction with the all parameter.



.Example
    PS C:\> Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx | Hide-SLGridLines -WorksheetName sheet1,sheet3 -Verbose | Save-SLDocument
    VERBOSE: sheet1 : is now selected
    VERBOSE: sheet1	: Removing Gridlines...
    VERBOSE: sheet3 : is now selected
    VERBOSE: sheet3	: Removing Gridlines...


    Description
    -----------
    Remove gridlines from sheet1 & sheet3 contained in MyFirstDoc.


.Example
    PS C:\> Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx | Hide-SLGridLines -All -Verbose  | Save-SLDocument

    VERBOSE: WorkSheet - Sheet1	: Removing Gridlines...
    VERBOSE: WorkSheet - sheet2	: Removing Gridlines...
    VERBOSE: WorkSheet - Sheet3	: Removing Gridlines...
    VERBOSE: WorkSheet - Sheet4	: Removing Gridlines...
    VERBOSE: WorkSheet - Sheet5	: Removing Gridlines...
    VERBOSE: WorkSheet - A	: Removing Gridlines...
    VERBOSE: WorkSheet - B	: Removing Gridlines...
    VERBOSE: WorkSheet - C	: Removing Gridlines...
    VERBOSE: Finsihed removing Gridlines from all worksheets contained in workbook : MyFirstDoc


    Description
    -----------
    Remove gridlines from all worksheets contained in MyFirstDoc.



.INPUTS
   String,SpreadsheetLight.SLDocument

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

        [parameter(Mandatory = $true, Position = 1, ValueFromPipeLine = $false, ParameterSetName = 'Named')]
        [string[]]$WorkSheetName,

        [parameter(ParameterSetName = 'All')]
        [Switch]$All


    )
    PROCESS
    {
        $sheets = $WorkBookInstance.GetSheetNames()
        $pagesettings = New-Object SpreadsheetLight.SLPageSettings
        $pagesettings.ShowGridLines = $false

        # ParameterSet 'All' - Will Remove gridlines from all worksheets contained in the specified workbook
        If ($PSCmdlet.ParameterSetName -eq 'All')
        {

            foreach ($w in $sheets)
            {
                $WorkBookInstance.SelectWorksheet($w) | Out-Null
                Write-Verbose ("Hide-SLGridLines :`tWorkSheet - '{0}'`t: Removing Gridlines..." -f $w )
                $WorkBookInstance.SetPageSettings($pagesettings, $w) | Out-Null
            }
            Write-Verbose "Hide-SLGridLines : Finsihed removing Gridlines from all worksheets contained in workbook : $($WorkBookInstance.WorkbookName)  "
        }

        # ParameterSet 'Named' - Will Remove gridlines specified worksheet contained in the specified workbook
        if ($PSCmdlet.ParameterSetName -eq 'Named')
        {

            foreach ($w in $WorksheetName)
            {
                if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $w -NoPassThru)
                {
                    Write-Verbose ("Hide-SLGridLines : '{0}'`t: Removing Gridlines..." -f $w )
                    $WorkBookInstance.SetPageSettings($pagesettings, $w) | Out-Null
                }
            }
        }

        $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

    }# process

    END
    {

    }
}


Function Show-SLGridLines
{

    <#

.SYNOPSIS
    Show Gridlines from one or more worksheets.

.DESCRIPTION
    Show Gridlines from one or more worksheets.


.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet where gridlines are to be shown.

.PARAMETER All
    If specified will show gridlines from all worksheets within a workbook( assuming they have been hidden ).
    The user does not need to use the worksheetname parameter in conjunction with the all parameter.



.Example
    PS C:\> Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx | Show-SLGridLines -WorksheetName sheet1,sheet3 -Verbose | Save-SLDocument

    VERBOSE: sheet1 : is now selected
    VERBOSE: sheet1	: Showing Gridlines...
    VERBOSE: sheet3 : is now selected
    VERBOSE: sheet3	: Showing Gridlines...


    Description
    -----------
    Show gridlines from sheet1 & sheet3 contained in MyFirstDoc.


.Example
    PS C:\> Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx | Show-SLGridLines -All -Verbose  | Save-SLDocument

    VERBOSE: Sheet1	: Showing Gridlines...
    VERBOSE: sheet2_renamed	: Showing Gridlines...
    VERBOSE: Sheet3	: Showing Gridlines...
    VERBOSE: Sheet4	: Showing Gridlines...
    VERBOSE: Sheet5	: Showing Gridlines...
    VERBOSE: A	: Showing Gridlines...
    VERBOSE: B	: Showing Gridlines...
    VERBOSE: C	: Showing Gridlines...
    VERBOSE: Finsihed Showing Gridlines for all worksheets contained in workbook : MyFirstDoc


    Description
    -----------
    Show gridlines from all worksheets contained in MyFirstDoc.



.INPUTS
   String,SpreadsheetLight.SLDocument

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

        [parameter(Mandatory = $true, Position = 1, ValueFromPipeLine = $false, ParameterSetName = 'Named')]
        [string[]]$WorksheetName,

        [parameter(ParameterSetName = 'All')]
        [Switch]$All



    )
    PROCESS
    {

        $sheets = $WorkBookInstance.GetSheetNames()
        $pagesettings = New-Object SpreadsheetLight.SLPageSettings
        $pagesettings.ShowGridLines = $true

        # ParameterSet 'All' - Will Show gridlines from all worksheets contained in the specified workbook
        If ($PSCmdlet.ParameterSetName -eq 'All')
        {

            foreach ($w in $sheets)
            {
                $WorkBookInstance.SelectWorksheet($w) | Out-Null
                Write-Verbose ("Show-SLGridLines : '{0}'`t: Showing Gridlines..." -f $w )
                $WorkBookInstance.SetPageSettings($pagesettings, $w) | Out-Null
            }
            Write-Verbose "Show-SLGridLines : Finsihed Showing Gridlines for all worksheets contained in workbook : $($WorkBookInstance.WorkbookName)  "
        }

        # ParameterSet 'Named' - Will Show gridlines specified worksheet contained in the specified workbook
        if ($PSCmdlet.ParameterSetName -eq 'Named')
        {

            foreach ($w in $WorksheetName)
            {
                if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $w -NoPassThru)
                {
                    Write-Verbose ("Show-SLGridLines : '{0}'`t: Showing Gridlines..." -f $w )
                    $WorkBookInstance.SetPageSettings($pagesettings, $w) | Out-Null
                }
            }
        }

        $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

    }# Process


    END
    {

    }
}






Function Import-SLTextFile
{

    <#

.SYNOPSIS
    Import text documents into excel.

.DESCRIPTION
    Import text documents into excel.Text delimiter may be TAB,Space,Semicolon or fixedWidth
    Note: Import does not work properly when an imported column begins with an "=" operator.
    This is because excel will try to interpret those cells as formulas instead of text.


.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet where gridlines are to be shown.

.PARAMETER TextFile
    Path to the textfile that is to be imported.Only one path can be specified at a time.

.PARAMETER Delimiter
    Text delimiter. Use tab or intellisense to select from a list of values:TAB,Space,Semicolon or fixedWidth

.PARAMETER ImportStartCell
    Begin text import at this cellreference Eg. B2.
    Default value is B2.

.PARAMETER SkipColumns
    Column ID's to be skipped.More than one column ID can be specified.Eg. 2,3
    The column ID's passed as arguements to this parameter wont be imported into excel.

.PARAMETER Culture
    The culture to be used Eg. de-DE (German)

.PARAMETER DateColumnIndex
    The Column ID's that are to be formatted as dates. More than one column ID can be specified.Eg. 2,3
    To be used along with the 'DateFormat' parameter.

.PARAMETER ImportCustomDateFormat
    To be used when the text to be imported contains date in a non-standard format.

.PARAMETER AutofitColumns
    Autofit all columns.

.PARAMETER Force
    Use force to overwrite an existing worksheet in a workbook.


.Example
    PS C:\> Get-SLDocument -Path C:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $textfile = "C:\temp\delim.txt"
    PS C:\> $doc | Import-SLTextFile -TextFile $textfile -Delimiter Tab -ImportStartCell B2 -DateColumnIndex 7 -DateFormat YYYY-MM-dd -Verbose


    Description
    -----------
    Import delim.txt into an existing document 'MyFirstDoc'. Column 7 will be formatted as a date


.Example
    PS C:\> $doc = New-SLDocument -Path C:\ps\Excel\MyFirstDoc.xlsx -Passthru
    PS C:\> $textfile = "C:\temp\delim.txt"
    PS C:\> $doc | Import-SLTextFile -TextFile $textfile -Delimiter Tab -ImportStartCell B2 -DateColumnIndex 7 -DateFormat YYYY-MM-dd -SkipColumns 3,5 -Verbose


    Description
    -----------
    Import delim.txt into a new document 'MyFirstDoc'. Column 7 will be formatted as a date
    Columns 3 and 5 will be skipped.



.INPUTS
   String,SpreadsheetLight.SLDocument

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    N/A

#>


    [CmdletBinding()]
    param (

        [parameter(Mandatory = $true, Position = 1, ValueFromPipeline = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [Alias('FullName')]
        [parameter(Mandatory = $true, Position = 0, ValueFromPipeLineByPropertyName = $true)]
        [String]$TextFile,

        [ValidateSet('Tab', 'Space', 'SemiColon', 'FixedWidth')]
        [parameter(Mandatory = $true, Position = 2)]
        [string]$Delimiter,

        [parameter(Mandatory = $false)]
        [String]$ImportStartCell = 'B2',

        [parameter(Mandatory = $false)]
        [System.UInt32[]]$SkipColumns,

        [parameter(Mandatory = $false)]
        [String]$Culture,

        [parameter(Mandatory = $false)]
        [System.UInt32[]]$DateColumnIndex,

        [parameter(Mandatory = $false)]
        [String]$DateFormat,

        [parameter(Mandatory = $false)]
        [String]$ImportCustomDateFormat,

        [switch]$AutofitColumns = $true,

        [switch]$Force

    )
    BEGIN
    {
    }
    PROCESS
    {
        $filestats = Get-Item $TextFile
        $worksheetName = $filestats.basename
        Write-Verbose ("Import-SLTextFile :`tProcessing TEXT FIle '{0}'..." -f $filestats.Name)

        $ImportOptions = New-Object Spreadsheetlight.SLTextImportOptions

        switch ($delimiter)
        {
            'Comma' { $ImportOptions.UseCommaDelimiter = $true }
            'SemiColon' { $ImportOptions.UseSemicolonDelimiter = $true }
            'Space' { $ImportOptions.UseSpaceDelimiter = $true }
            'Tab' { $ImportOptions.UseTabDelimiter = $true }
            'FixedWidth' { $ImportOptions.DataFieldType = [SpreadsheetLight.SLTextImportDataFieldTypeValues]::'FixedWidth' }
        }


        if ($SkipColumns)
        {
            foreach ($SkipColumn in $SkipColumns)
            {
                Write-Verbose ("Import-SLTextFile :`tSkipping Import of columnID '{0}'..." -f $SkipColumn)
                $ImportOptions.SkipColumn($SkipColumn)
            }
        }

        if ($Culture)
        {
            Write-Verbose ("Import-SLTextFile :`tSetting culture to '{0}'..." -f $Culture)
            $ImportOptions.Culture = New-Object System.Globalization.CultureInfo($Culture)
        }

        If ($ImportCustomDateFormat)
        {
            Write-Verbose ("Import-SLTextFile :`tAdding custom date format '{0}'..." -f $ImportCustomDateFormat)
            $ImportOptions.AddCustomDateFormat($ImportCustomDateFormat) | Out-Null
        }


        if ($WorkBookInstance.GetSheetNames() -notcontains $WorksheetName)
        {

            $ProcessFile = $true
            $WorkBookInstance.AddWorksheet($WorksheetName) | Out-Null

        }
        Else
        {
            if ($force -and $PSCmdlet.ShouldProcess($WorksheetName, 'OVERWRITE Worksheet'))
            {
                $ProcessFile = $true
                $random = Get-Random -Maximum 100
                $WorkBookInstance.AddWorksheet($random) | Out-Null
                $WorkBookInstance.DeleteWorksheet($worksheetName) | Out-Null
                $WorkBookInstance.AddWorksheet($WorksheetName) | Out-Null
                $WorkBookInstance.DeleteWorksheet($random) | Out-Null
            }
            Else
            {
                $ProcessFile = $false
                Write-Warning ("Import-SLTextFile :`tSpecified Worksheet '{0}' already Exists. Please select a different name or use the '-Force' parameter to overwrite" -f $worksheetName)
            }
        }



        If ($ProcessFile)
        {
            Write-Verbose ("Import-SLTextFile :`tImporting data into Worksheet '{0}' " -f $worksheetName)
            $WorkBookInstance.ImportText($TextFile, $ImportStartCell, $ImportOptions)

            If ($DateFormat)
            {
                $SLStyle = $WorkBookInstance.CreateStyle()
                $SLStyle.FormatCode = $DateFormat
                foreach ($DateColumn in $DateColumnIndex)
                {
                    if ($ImportStartCell)
                    {
                        $AdjustedDateColumn = ((Convert-ToExcelColumnIndex -ColumnName B3) - 1 ) + $DateColumn
                        Write-Verbose ("Import-SLTextFile :`tSetting DateFormat '{0}' on ColumnID '{1}' " -f $DateFormat, $DateColumn)
                        $WorkBookInstance.SetColumnStyle($AdjustedDateColumn, $SLStyle) | Out-Null
                    }
                    Else
                    {
                        Write-Verbose ("Import-SLTextFile :`tSetting DateFormat '{0}' on ColumnID '{1}' " -f $DateFormat, $DateColumn)
                        $WorkBookInstance.SetColumnStyle($DateColumn, $SLStyle) | Out-Null
                    }
                }
            }


            ## Add Autofit to the Table - Optional
            if ($AutofitColumns)
            {
                $WorkBookInstance.autofitcolumn('A', 'DD')
            }

            $stats = $WorkBookInstance.GetWorksheetStatistics()

            $WorkBookInstance | Add-Member NoteProperty StartRowIndex $stats.StartRowIndex -Force
            $WorkBookInstance | Add-Member NoteProperty StartColumnIndex $stats.StartColumnIndex -Force
            $WorkBookInstance | Add-Member NoteProperty EndRowIndex $stats.ENDRowIndex -Force
            $WorkBookInstance | Add-Member NoteProperty EndColumnIndex $stats.ENDColumnIndex -Force
            $WorkBookInstance | Add-Member NoteProperty WorksheetNames $WorkBookInstance.GetSheetNames() -Force
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force
            Write-Output $WorkBookInstance

        }#IF processfile

    }#process
    END
    {
    }
}



Function Export-SLDocument
{

    <#

.SYNOPSIS
    Stores data in a datatable which is the input type that excel accepts.


.DESCRIPTION
    Stores data in a datatable which is the input type that excel accepts.
    Since there may be a possibility of overwriting existing data the workbook is backed up prior to processing the command.
    Location of backup --> $Env:temp\SLPSLib

.PARAMETER InputObject
    Data in the form of rows and columns. eg. output from the cmdlet 'Get-service'

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.Please note you cannot pipe a workbookinstance to this cmdlet. Instead use named parameter.

.PARAMETER WorksheetName
    Name of the Worksheet where data will be exported. Make sure this is blank if not existing data will be overwritten.

.PARAMETER StartRowIndex
    Row number which marks the start of the data table.Default value is 5


.PARAMETER StartColumnIndex
    Column number which marks the start of the data table.Default value is 2


.PARAMETER AutofitColumns
    Autofit all columns that contain data in the selected worksheet.

.PARAMETER ParseStringData
    For the most part powershell handles dataconversion to its proper datatype but it cannot help
    when data is explicitly cast as a string which gives rise to mismatch between data and datatype.
    Eg: $a = "12" stored as a string even though the value is an integer.

    Mismatched datatypes may also result due to poorly built functions that cast everything as a string.
    parsestringdata tries to coerce these string values into their respective datatypes(Integer,Double, or datetime)
    eg: $a = "25-07-2014" is a string so excel will store this as a string but when parsestring data is used the value of $a is stored as a datetime instead of a string.

    Note: In Excel Numbers are always right aligned and strings leftaligned.

.Example
    PS C:\> $doc = Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> Get-Service | Export-SLDocument -WorkBookInstance $doc -WorksheetName MyComp_Services -StartRowIndex 3 -StartColumnIndex 2 -AutofitColumns | Save-SLDocument

    Description
    -----------
    Get-Service is piped to an instance of 'MyFirstDoc'. The output is saved to a worksheet named 'MyComp_Services'.
    Note: A new worksheet will be created in case the specified worksheet dosen't exist.

.Example
    PS C:\> $doc = Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $Service = Get-Service
    PS C:\> $Process = Get-Process
    PS C:\> $Disk = Get-WmiObject -Class Win32_LogicalDisk -ComputerName "Localhost"
    PS C:\> Export-SLDocument -inputobject $Service -WorkBookInstance $doc -WorksheetName Service -AutofitColumns
    PS C:\> Export-SLDocument -inputobject $Process -WorkBookInstance $doc -WorksheetName Process -AutofitColumns
    PS C:\> Export-SLDocument -inputobject $Disk    -WorkBookInstance $doc -WorksheetName Disk    -AutofitColumns
    PS C:\> Save-SLDocument -WorkBookInstance $doc

    Description
    -----------
    An instance of MyFirstDoc is stored in a variable named doc.
    Service,process and diskdata from the localcomputer is then exported to worksheets 'service','process' & 'Disk' respectively.



.Example
    PS C:\> Export-SLDocument -inputobject (Get-EventLog -LogName System -Newest 5 | Select InstanceID,TimeGenerated,EntryType,Message) -WorkBookInstance (New-SLDocument -WorkbookName Eventlog -Path D:\ps\excel -PassThru) -WorksheetName System -AutofitColumns | Save-SLDocument


    Description
    -----------
    A one-liner to get the newest 5 entries from the system eventlog to a new workbook named 'Eventlog'.

.Example
    PS C:\> $ServiceDoc =  New-SLDocument -WorkbookName MyComp_Services -WorksheetName Service -Path D:\ps\Excel -Verbose -PassThru
    PS C:\> $ProcessDoc =  New-SLDocument -WorkbookName MyComp_Process  -WorksheetName Process -Path D:\ps\Excel -Verbose -PassThru
    PS C:\> $DiskDoc    =  New-SLDocument -WorkbookName MyComp_Disk     -WorksheetName Disk    -Path D:\ps\Excel -Verbose -PassThru
    PS C:\> $Service = Get-Service
    PS C:\> $Process = Get-Process
    PS C:\> $Disk = Get-WmiObject -Class Win32_LogicalDisk -ComputerName "Localhost"
    PS C:\> Export-SLDocument -inputobject $Service   -WorkBookInstance $ServiceDoc -WorksheetName Service  -AutofitColumns  | Save-SLDocument
    PS C:\> Export-SLDocument -inputobject $Process   -WorkBookInstance $ProcessDoc -WorksheetName Process  -AutofitColumns  | Save-SLDocument
    PS C:\> Export-SLDocument -inputobject $Disk      -WorkBookInstance $DiskDoc    -WorksheetName Disk     -AutofitColumns  | Save-SLDocument

    Description
    -----------
    3 new documents are created.One each for service,process and disk respectively.
    Export-SLDocument is then used to export the relevant data to each fo the workbooks.
    Note: 'Passthru' parameter with New-SLDocument is required when you want to store the document instance in a variable as shown above.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $Service = Get-Service | Group-Object -Property Status -AsHashTable -AsString
    PS C:\> $Running_Svcs = $service.Running | Select Name,DisplayName,Status
    PS C:\> $Stopped_Svcs = $service.Stopped | Select Name,DisplayName,Status
    PS C:\> Export-SLDocument -inputobject $Running_Svcs -WorkBookInstance $doc -WorksheetName Service -AutofitColumns -StartRowIndex 3 -StartColumnIndex 2
    PS C:\> Export-SLDocument -inputobject $Stopped_Svcs -WorkBookInstance $doc -WorksheetName Service -AutofitColumns -StartRowIndex 3 -StartColumnIndex 6
    PS C:\> Save-SLDocument -WorkBookInstance $doc

    Description
    -----------
    Group-Object is used to group by the status property .
    Export-SLDocument is then used to export the running services starting from column number 2 to 5.
    Stopped services are then exported to columns 6-8 in the same worksheet.


.INPUTS
   Object,SpreadsheetLight.SLDocument

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    N/A

#>


    [CmdletBinding()]
    [OutputType([SpreadsheetLight.SLDocument])]
    param (

        [parameter(Mandatory = $true, Position = 1, valuefrompipeline = $true)]
        $inputobject,

        [parameter(Mandatory = $true, Position = 0)]
        [SpreadsheetLight.SLDocument]
        $WorkBookInstance,

        [parameter(Mandatory = $true, Position = 2)]
        [String]
        $WorksheetName,

        [ValidateRange(1, 50)]
        [parameter(Mandatory = $false, Position = 3)]
        [int]
        $StartRowIndex = 1,

        [ValidateRange(1, 20)]
        [parameter(Mandatory = $false, Position = 4)]
        [int]
        $StartColumnIndex = 1,

        [switch]
        $AutofitColumns = $true,

        [Switch]
        $ParseStringData,

        [Switch]$ClearWorksheet,

        [string]
        $Path


    )
    BEGIN
    {
        $Data = @()
        $dt = New-Object System.Data.DataTable
    }
    PROCESS
    {
        $Data += $InputObject
    }
    END
    {
        Backup-SLDocument -WorkBookInstance $WorkBookInstance
        if ($WorkBookInstance.GetSheetNames() -notcontains $WorksheetName)
        {
            $WorkBookInstance.AddWorksheet($WorksheetName) | Out-Null
        }
        Else
        {
            $WorkBookInstance.SelectWorksheet($WorksheetName) | Out-Null
        }

        Write-Verbose "Export-SLDocument :`tCreating Datatable..."
        #region Create DataTable
        $dt = New-Object System.Data.DataTable
        $DataHeaders = @()
        $DateHeaders = @()

        #$DataHeaders += $Data[0] | Get-Member -MemberType Properties | select  -ExpandProperty name
        $DataHeaders += $Data[0].psobject.Properties | Select-Object -ExpandProperty name

        Write-Verbose "Export-SLDocument :`tAdding column Headers to Datatable..."
        ## Add datatable Columns
        ForEach ($d in $DataHeaders )
        {

            $DataColumn = $d
            try
            {
                $ErrorActionPreference = 'stop'
                if ([string]::IsNullOrEmpty($($data[0].$DataColumn)))
                {
                    $dt.columns.add($DataColumn, [String]) | Out-Null
                }
                else
                {
                    $Dtype = ($data[0].$DataColumn).gettype().name
                    Switch -regex  ( $Dtype )
                    {

                        'string'
                        {
                            if ( $parseStringData )
                            {
                                $ConvertedIntValue = ''
                                $ConvertedDoubleValue = ''
                                $Int = [Int]::TryParse($data[0].$DataColumn, [ref]$ConvertedIntValue)
                                $Double = [Double]::TryParse($data[0].$DataColumn, [ref]$ConvertedDoubleValue)
                                try
                                {

                                    $ConvertedDateValue = Get-Date -Date $data[0].$DataColumn -ErrorAction Stop
                                    $IsDateTime = $true
                                    $DateHeaders += $DataColumn
                                }
                                catch
                                {
                                    $IsDateTime = $false
                                }

                                if ($ConvertedIntValue -ne 0 -and $ConvertedDoubleValue -ne 0 )
                                {
                                    $dt.columns.add($DataColumn, [Int]) | Out-Null
                                }
                                elseif ($ConvertedIntValue -eq 0 -and $ConvertedDoubleValue -ne 0)
                                {
                                    $dt.columns.add($DataColumn, [Double]) | Out-Null
                                }
                                elseif ($IsDateTime)
                                {
                                    $dt.columns.add($DataColumn, [DateTime]) | Out-Null
                                }

                                else
                                {
                                    $dt.columns.add($DataColumn) | Out-Null
                                }
                                break;

                            }#Ifparsestringdatatype
                            Else
                            {
                                $dt.columns.add($DataColumn, [String]) | Out-Null
                                break;
                            }
                        }
                        'Double'
                        {
                            $dt.columns.add($DataColumn, [Double]) | Out-Null
                            break;
                        }
                        'Datetime'
                        {
                            $dt.columns.add($DataColumn, [DateTime]) | Out-Null
                            $DateHeaders += $DataColumn
                            break;
                        }

                        'Boolean'
                        {
                            $dt.columns.add($DataColumn, [System.Boolean]) | Out-Null
                            Break
                        }

                        'Byte\[\]'
                        {
                            $dt.columns.add($DataColumn, [System.Byte[]]) | Out-Null
                            $dt.Columns[$DataColumn].DataType = [System.String]
                            break;
                        }
                        'Byte'
                        {
                            $dt.columns.add($DataColumn, [System.Byte]) | Out-Null
                            Break
                        }

                        'char'
                        {
                            $dt.columns.add($DataColumn, [System.Char]) | Out-Null
                            break;
                        }
                        'Decimal'
                        {
                            $dt.columns.add($DataColumn, [System.Decimal]) | Out-Null
                            Break
                        }

                        'Guid'
                        {
                            $dt.columns.add($DataColumn, [System.Guid]) | Out-Null
                            break;
                        }
                        'Int16'
                        {
                            $dt.columns.add($DataColumn, [System.Int16]) | Out-Null
                            Break
                        }

                        'Int32'
                        {
                            $dt.columns.add($DataColumn, [System.Int32]) | Out-Null
                            break;
                        }
                        'Int64|long'
                        {
                            $dt.columns.add($DataColumn, [System.Int64]) | Out-Null
                            break;
                        }
                        'UInt16'
                        {
                            $dt.columns.add($DataColumn, [System.UInt16]) | Out-Null
                            Break
                        }

                        'UInt32'
                        {
                            $dt.columns.add($DataColumn, [System.UInt32]) | Out-Null
                            break;
                        }
                        'UInt64|long'
                        {
                            $dt.columns.add($DataColumn, [System.UInt64]) | Out-Null
                            Break
                        }

                        'Single'
                        {
                            $dt.columns.add($DataColumn, [System.Single]) | Out-Null
                            break;
                        }
                        'IntPtr'
                        {
                            $dt.columns.add($DataColumn, [System.IntPtr]) | Out-Null
                            $dt.Columns[$DataColumn].DataType = [System.Int64]
                            break;
                        }

                        Default
                        {
                            $dt.columns.add($DataColumn) | Out-Null

                        }
                    }#switch
                }#else

            }
            catch
            {
                $ErrorActionPreference = 'continue'
                if ($null -eq $Dtype)
                {
                    $dt.columns.add($DataColumn, [String]) | Out-Null
                }
                #Write-Warning $Error[0].Exception.Message
            }

        }# END foreach dataheaders

        Write-Verbose "Export-SLDocument :`tAdding Rows to Datatable..."
        ## Add datatable Rows
        for ($i = 0; $i -lt $data.count; $i++)
        {
            $row = $dt.NewRow()
            foreach ($dhead in $DataHeaders)
            {
                If ([string]::IsNullOrEmpty($Data[$i].$dhead))
                {
                    $row.Item($dhead) = [DBNull]::Value
                }
                Else
                {
                    Try
                    {
                        $ErrorActionPreference = 'Stop'
                        if ($Data[$i].$dhead.Gettype().name -match 'Intptr' )
                        {
                            $row.Item($dhead) = $Data[$i].$dhead.ToInt64()
                        }
                        Elseif ($Data[$i].$dhead.Gettype().basetype.name -eq 'array')
                        {
                            $row.Item($dhead) = $Data[$i].$dhead -join ','
                        }
                        Elseif ($Data[$i].$dhead.Gettype().name -match 'byte\[\]')
                        {
                            $row.Item($dhead) = $Data[$i].$dhead -join ','
                        }
                        Else
                        {
                            $row.Item($dhead) = $Data[$i].$dhead
                        }
                    }
                    Catch
                    {
                        Write-Warning ("Export-SLDocument :`tAn Error Occured...{0}" -f $Error[0].Exception.Message)
                        $ErrorActionPreference = 'Continue'
                    }
                }

            }

            $dt.Rows.Add($row)
        }

        #ENDregion Create DataTable
        Write-Verbose "Export-SLDocument :`tFinsihed creating the Datatable.Loading data into excel.."

        if ($ClearWorksheet)
        {
            $WorkBookInstance.ClearCellContent()
        }

        $WorkBookInstance.ImportDataTable($StartRowIndex, $StartColumnIndex, $dt, $true ) | Out-Null
        $WorkBookInstance | Add-Member NoteProperty DataTable $dt -Force
        $dt.Dispose()

        ##  Add dateformat to the date headers
        $stats = $WorkBookInstance.GetWorksheetStatistics()
        $dhrange = @()
        $DataHeaders = $DataHeaders | ForEach-Object { $_.ToString().ToUpper() }

        $DateHeaders |
            ForEach-Object {
                $h = $_.tostring().toupper()
                $dhcolumn = [array]::IndexOf($DataHeaders, $h)
                $dh = $dhcolumn + $StartColumnIndex
                $dhrange += [SpreadsheetLight.SLConvert]::ToCellRange( ($StartRowIndex + 1 ), $dh, $stats.ENDRowIndex, $dh )

            }


        $SLStyle = $WorkBookInstance.CreateStyle()
        $SLStyle.FormatCode = 'dd/MM/yyyy h:mm:ss AM/PM'


        $dhrange |
            ForEach-Object {
                $StartCellReference, $ENDCellReference = $_ -split ':'
                $WorkBookInstance.SetCellStyle($StartCellReference, $ENDCellReference, $SLStyle) | Out-Null
            }


        ## AutoFit Columns
        if ($AutofitColumns)
        {
            $WorkBookInstance.autofitcolumn('A', 'DD')
        }

        if ($path)
        {
            $WorkBookInstance.SaveAs($path)
            Write-Verbose ("Export-SLDocument :`tDocument has been Saved to path $Path")

        }
        else
        {
            $WorkBookInstance.Save()
            Write-Verbose ("Export-SLDocument :`tDocument has been Saved to path $($WorkBookInstance.Path)")
        }

        <#
        $HeaderRange = Convert-ToExcelRange -StartRowIndex $stats.StartRowIndex -StartColumnIndex $stats.StartColumnIndex -EndRowIndex $stats.StartRowIndex -EndColumnIndex $stats.ENDColumnIndex
        $DataRange = Convert-ToExcelRange -StartRowIndex ($stats.StartRowIndex + 1) -StartColumnIndex $stats.StartColumnIndex -EndRowIndex $stats.EndRowIndex -EndColumnIndex $stats.ENDColumnIndex
        $FirstDataColumn = Convert-ToExcelRange -StartRowIndex ($stats.StartRowIndex + 1) -StartColumnIndex $stats.StartColumnIndex -EndRowIndex $stats.EndRowIndex -EndColumnIndex $stats.StartColumnIndex


        $WorkBookInstance | Add-Member NoteProperty StartRowIndex $stats.StartRowIndex -Force
        $WorkBookInstance | Add-Member NoteProperty StartColumnIndex $stats.StartColumnIndex -Force
        $WorkBookInstance | Add-Member NoteProperty EndRowIndex $stats.ENDRowIndex -Force
        $WorkBookInstance | Add-Member NoteProperty EndColumnIndex $stats.ENDColumnIndex -Force
        $WorkBookInstance | Add-Member NoteProperty HeaderRange $HeaderRange -Force
        $WorkBookInstance | Add-Member NoteProperty DataRange $DataRange -Force
        $WorkBookInstance | Add-Member NoteProperty FirstDataColumn $FirstDataColumn -Force

        $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru #>

    }
    CLEAN {
        $WorkBookInstance.Dispose()
    }
}



Function Set-SLTableStyle
{

    <#

.SYNOPSIS
    Excel offers to style your data tables via some built-in styles. This cmdlet help the user choose a built-in table style.


.DESCRIPTION
    Excel offers to style your data tables via some built-in styles. This cmdlet helps the user choose a built-in table style.
    In order to set a tablestyle excel would need to know the startrowindex,startcolumnindex,endrowindex and endcolumnindex,
    or simply the range eg: A1:B10

    If you want to apply a style to existing data in a worksheet then you would need to obtain the table values namely
    startrowindex,startcolumnindex,endrowindex,endcolumnindex or the range and then feed those values to the parameters.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER TableStyle
    There are about 55 built-in styles to choose from, ranging from light-dark.
    While there is a way to set a table style there isnt however a method to remove an applied style.
    Use tab or intellisense to choose from a list of possible values:
    'light1','light2','light3','light4','light5','light6','light7','light8','light9','light10','light11','light12','light13','light14','light15','light16','light17'
    ,'light18','light19','light20','light21','Medium1','Medium2','Medium3','Medium4','Medium5','Medium6','Medium7','Medium8','Medium9','Medium10','Medium11','Medium12','Medium13','Medium14'
    ,'Medium15','Medium16','Medium17','Medium18','Medium19','Medium20','Medium21','Medium22','Medium23','Medium24','Medium25','Medium26','Medium27','Medium28'
    ,'Dark1','Dark2','Dark3','Dark4','Dark5','Dark6','Dark7','Dark8','Dark9','Dark10','Dark11'


.PARAMETER TotalRowFunction
    Choose from the following as one of the valid options for a totalrowfunction.
    Possible Values : 'Sum','Count','Average','Product','Maximum','Minimum','CountNumbers','StandardDeviation','Variance'
    Note: Excel 2007 does not contain the 'Product' function.

.PARAMETER TotalColumnIndex
    The data column to which the totalrowfunction has to be applied.Valid values start from 1 irrespective of the column from which the data table starts..

.PARAMETER TotalRowLabel
    The label text indicating the calculated value from the total row function Eg. 'Average Sales Revenue'.

.PARAMETER TotalRowLabelColumnIndex
    The column index to set the TotalRowLabel on Eg. 2 or 4 . values start from 1 irrespective of the column from which the data table starts.

.PARAMETER StartRowIndex
    Row number which marks the start of the data table.

.PARAMETER StartColumnIndex
    Column number which marks the start of the data table.

.PARAMETER EndRowIndex
    Row number which marks the end of the data table.

.PARAMETER EndColumnIndex
    Column number which marks the end of the data table.

.PARAMETER Range
    The range that constitutes the table data eg: A1:b10.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $Service = Get-Service | Select Name,DisplayName,Status
    PS C:\> Export-SLDocument -inputobject $Service -WorkBookInstance $doc -WorksheetName Sheet2 |
                        Set-SLTableStyle -WorksheetName sheet2 -TableStyle Dark10 | Save-SLDocument

    Description
    -----------
    An instance of MyFirstDoc is stored in a variable named doc.
    Service data is first exported to sheet2 and then a built-in tablestyle named 'Dark10' is set on the table.
    This can probably be condensed into a oneliner but for the sake of simplicity the activity has been broken down into several steps.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $Size = @{n='Size-GB';e={$_.size/1GB -as [INT]}}
    PS C:\> $FreeSpace = @{n='FreeSpace-GB';e={$_.FreeSpace/1GB -as [INT]}}
    PS C:\> $disk = Get-WmiObject -Class Win32_Logicaldisk | select SystemName,DeviceID,VolumeName,DriveType,$size,$FreeSpace
    PS C:\> Export-SLDocument -inputobject $Disk -WorkBookInstance $doc -WorksheetName Disk |
                    Set-SLTableStyle -WorksheetName disk -TableStyle Dark11  -TotalRowFunction Sum -TotalColumnIndex 6 | Save-SLDocument

    Description
    -----------
    Disk Information is exported to a worksheetnamed 'disk'. The 'sum' function is applied to the contents of the column 6(FreeSpace)


.Example
    PS C:\> $doc = Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $Service = Get-Service
    PS C:\> $Process = Get-Process
    PS C:\> $Disk = Get-WmiObject -Class Win32_LogicalDisk -ComputerName "Localhost"
    PS C:\> Export-SLDocument -inputobject $Service -WorkBookInstance $doc -WorksheetName Service -AutofitColumns | Set-SLTableStyle -WorksheetName Service -TableStyle Medium16
    PS C:\> Export-SLDocument -inputobject $Process -WorkBookInstance $doc -WorksheetName Process -AutofitColumns | Set-SLTableStyle -WorksheetName Process -TableStyle Medium16
    PS C:\> Export-SLDocument -inputobject $Disk    -WorkBookInstance $doc -WorksheetName Disk    -AutofitColumns | Set-SLTableStyle -WorksheetName Disk    -TableStyle Medium16
    PS C:\> Save-SLDocument -WorkBookInstance $doc

    Description
    -----------
    An instance of MyFirstDoc is stored in a variable named doc.
    Service,process and disk data from the localcomputer is then exported to worksheets
    'service','process' & 'Disk' with styles Medium16,17 & 28 applied respectively.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $Service = Get-Service | Group-Object -Property Status -AsHashTable -AsString
    PS C:\> $Running_Svcs = $service.Running | Select Name,DisplayName,Status
    PS C:\> $Stopped_Svcs = $service.Stopped | Select Name,DisplayName,Status
    PS C:\> Export-SLDocument -inputobject $Running_Svcs -WorkBookInstance $doc -WorksheetName Service -AutofitColumns -StartRowIndex 3 -StartColumnIndex 2
    PS C:\> Export-SLDocument -inputobject $Stopped_Svcs -WorkBookInstance $doc -WorksheetName Service -AutofitColumns -StartRowIndex 3 -StartColumnIndex 6
    PS C:\> Save-SLDocument   -WorkBookInstance $doc
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> Set-SLTableStyle -WorkBookInstance $doc -WorksheetName Service -TableStyle Light4 -StartRowIndex 3 -StartColumnIndex 2 -EndRowIndex 67  -EndColumnIndex 4
    PS C:\> Set-SLTableStyle -WorkBookInstance $doc -WorksheetName Service -TableStyle Light3 -StartRowIndex 3 -StartColumnIndex 6 -EndRowIndex 118 -EndColumnIndex 8
    PS C:\> Save-SLDocument  -WorkBookInstance $doc

    Description
    -----------
    Get-Service data is piped to group-object to be grouped by the status property and the results are stored in a variable named 'service'.
    export the running services starting from column number 2 to 5.
    Stopped services are exported to columns 6-8 in the same worksheet.
    Since we are applying two different table styles to the same worksheet we need to manually find out the table values which is why we
    save the document open it find out the start and end values for the running and stopped service ranges and then apply our style.
    Note: We cannot pipe  Export-SLDocument to set-sltablestyle because the start and end values are calculated for the entire worksheet
    and so the same style will be applied to both tables.


.INPUTS
   String,SpreadsheetLight.SLDocument

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    N/A

#>


    [CmdletBinding()]
    [OutputType([SpreadsheetLight.SLDocument])]
    param (
        [parameter(Mandatory = $true, position = 0, valuefrompipeline = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [Alias('CurrentWorksheetName')]
        [parameter(Mandatory = $true, Position = 1, ValueFromPipelineByPropertyName = $true)]
        [String]$WorksheetName,

        [Validateset('light1', 'light2', 'light3', 'light4', 'light5', 'light6', 'light7', 'light8', 'light9', 'light10', 'light11', 'light12', 'light13', 'light14', 'light15', 'light16', 'light17'
            , 'light18', 'light19', 'light20', 'light21', 'Medium1', 'Medium2', 'Medium3', 'Medium4', 'Medium5', 'Medium6', 'Medium7', 'Medium8', 'Medium9', 'Medium10', 'Medium11', 'Medium12', 'Medium13', 'Medium14'
            , 'Medium15', 'Medium16', 'Medium17', 'Medium18', 'Medium19', 'Medium20', 'Medium21', 'Medium22', 'Medium23', 'Medium24', 'Medium25', 'Medium26', 'Medium27', 'Medium28'
            , 'Dark1', 'Dark2', 'Dark3', 'Dark4', 'Dark5', 'Dark6', 'Dark7', 'Dark8', 'Dark9', 'Dark10', 'Dark11')]
        [Parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true)]
        [String]$TableStyle,

        [ValidateSet('Sum', 'Count', 'Average', 'Product', 'Maximum', 'Minimum', 'CountNumbers', 'StandardDeviation', 'Variance')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Range-TotalRowFunction')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Index-TotalRowFunction')]
        [string]$TotalRowFunction,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Range-TotalRowFunction')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Index-TotalRowFunction')]
        [UInt32]$TotalColumnIndex,

        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Range-TotalRowFunction')]
        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Index-TotalRowFunction')]
        [String]$TotalRowLabel,

        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Range-TotalRowFunction')]
        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Index-TotalRowFunction')]
        [int]$TotalRowLabelColumnIndex,

        [ValidateNotNullOrEmpty()]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Index-TotalRowFunction')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Index')]
        [UInt32]$StartRowIndex,

        [ValidateNotNullOrEmpty()]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Index-TotalRowFunction')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Index')]
        [UInt32]$StartColumnIndex,

        [ValidateNotNullOrEmpty()]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Index-TotalRowFunction')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Index')]
        [UInt32]$EndRowIndex,

        [ValidateNotNullOrEmpty()]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Index-TotalRowFunction')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Index')]
        [UInt32]$EndColumnIndex,

        [ValidateNotNullOrEmpty()]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Range-TotalRowFunction')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Range')]
        [String]$Range

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            If ($PSCmdlet.ParameterSetName -eq 'Range')
            {
                $RangeValue = $Range
                $startcellreference, $endcellreference = $Range -split ':'
                $SLTable = $WorkBookInstance.CreateTable($startcellreference, $endcellreference)

                Write-Verbose ("Set-SLTableStyle :`tSetting TableStyle '{0}' on CellRange '{1}' " -f $TableStyle, $Range)
                $SLtable.SetTableStyle([SpreadsheetLight.SLTableStyleTypeValues]::$TableStyle)


            }

            If ($PSCmdlet.ParameterSetName -eq 'Index')
            {
                $RangeValue = Convert-ToExcelRange -StartRowIndex $StartRowIndex -StartColumnIndex $StartColumnIndex -EndRowIndex $EndRowIndex -EndColumnIndex $EndColumnIndex
                $SLTable = $WorkBookInstance.CreateTable($StartRowIndex, $StartColumnIndex, $ENDRowIndex, $ENDColumnIndex)

                Write-Verbose ("Set-SLTableStyle :`tSetting TableStyle '{0}' on CellRange - StartRow/StartColumn '{1}':'{2}' & EndRow/EndColumn '{3}':'{4}' " -f $TableStyle, $StartRowIndex, $StartColumnIndex, $ENDRowIndex, $ENDColumnIndex)
                $SLtable.SetTableStyle([SpreadsheetLight.SLTableStyleTypeValues]::$TableStyle)

            }

            If ($PSCmdlet.ParameterSetName -eq 'Range-TotalRowFunction')
            {
                $RangeValue = $Range
                $startcellreference, $endcellreference = $Range -split ':'
                $SLTable = $WorkBookInstance.CreateTable($startcellreference, $endcellreference)

                Write-Verbose ("Set-SLTableStyle :`tSetting TableStyle '{0}' on CellRange '{1}' " -f $TableStyle, $Range)
                $SLtable.SetTableStyle([SpreadsheetLight.SLTableStyleTypeValues]::$TableStyle)

                # Setting TotalRowFunction
                $sltable.HasTotalRow = $true;

                If ($TotalRowLabel)
                {
                    $SLTable.SetTotalRowLabel($TotalRowLabelColumnIndex, $TotalRowLabel) | Out-Null
                }

                Write-Verbose ("Set-SLTableStyle : Setting TotalRowFunction - '{0}' on worksheet '{1}' column '{2}' " -f $TotalRowFunction, $WorksheetName, $TotalColumnIndex)
                $sltable.SetTotalRowFunction($TotalColumnIndex, [spreadsheetlight.SLTotalsRowFunctionValues]::$TotalRowFunction  ) | Out-Null
            }

            If ($PSCmdlet.ParameterSetName -eq 'Index-TotalRowFunction')
            {
                $RangeValue = Convert-ToExcelRange -StartRowIndex $StartRowIndex -StartColumnIndex $StartColumnIndex -EndRowIndex $EndRowIndex -EndColumnIndex $EndColumnIndex
                $SLTable = $WorkBookInstance.CreateTable($StartRowIndex, $StartColumnIndex, $ENDRowIndex, $ENDColumnIndex)

                Write-Verbose ("Set-SLTableStyle :`tSetting TableStyle '{0}' on CellRange - StartRow/StartColumn '{1}':'{2}' & EndRow/EndColumn '{3}':'{4}' " -f $TableStyle, $StartRowIndex, $StartColumnIndex, $ENDRowIndex, $ENDColumnIndex)
                $SLtable.SetTableStyle([SpreadsheetLight.SLTableStyleTypeValues]::$TableStyle)

                $sltable.HasTotalRow = $true;

                If ($TotalRowLabel)
                {
                    $SLTable.SetTotalRowLabel($TotalRowLabelColumnIndex, $TotalRowLabel) | Out-Null
                }

                Write-Verbose ("Set-SLTableStyle : Setting TotalRowFunction - '{0}' on worksheet '{1}' column '{2}' " -f $TotalRowFunction, $WorksheetName, $TotalColumnIndex)
                $sltable.SetTotalRowFunction($TotalColumnIndex, [spreadsheetlight.SLTotalsRowFunctionValues]::$TotalRowFunction  ) | Out-Null

            }

            $WorkBookInstance.InsertTable($SLTable) | Out-Null

            $WorkBookInstance | Add-Member NoteProperty Range $RangeValue -Force
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }#select-worksheet
    }
}



Function Import-CSVToSLDocument
{

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


Function Get-SLCellStyle
{

    <#

.SYNOPSIS
    Gets the various style settings applied to a cell.

.DESCRIPTION
    Gets the various style settings applied to a cell.The style settings can be either accessed by their name or as a property that is attached to the workbookinstance.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER CellReference
    The cell whose style settings are to be obtained.

.PARAMETER Alignment
    Display the alignment settings applied to the specified cell.

.PARAMETER Font
    Display the font settings applied to the specified cell.

.PARAMETER Fill
    Display the fill settings applied to the specified cell.

.PARAMETER Border
    Display the border settings applied to the specified cell.

.PARAMETER FormatCode
    Display the formatcode settings applied to the specified cell.

.PARAMETER Protection
    Display the protection settings applied to the specified cell.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Get-SLCellStyle -WorksheetName sheet6 -CellReference d6 -Alignment
    PS C:\> $doc | Get-SLCellStyle -WorksheetName sheet6 -CellReference d6 -Font
    PS C:\> $doc | Get-SLCellStyle -WorksheetName sheet6 -CellReference d6 -Fill
    PS C:\> $doc | Get-SLCellStyle -WorksheetName sheet6 -CellReference d6 -Border
    PS C:\> $doc | Get-SLCellStyle -WorksheetName sheet6 -CellReference d6 -FormatCode
    PS C:\> $doc | Get-SLCellStyle -WorksheetName sheet6 -CellReference d6 -Protection
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    Display the various style settings applied to cell d6 on sheet6.



.INPUTS
   String,SpreadsheetLight.SLDocument

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

        [Alias('CurrentWorkSheetName')]
        [parameter(Mandatory = $true, Position = 1, ValueFromPipelineByPropertyName = $true)]
        [String]$WorksheetName,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, position = 2, parametersetname = 'cell')]
        [string]$CellReference,

        [Switch]$Alignment,

        [Switch]$Font,

        [Switch]$Fill,

        [Switch]$Border,

        [Switch]$FormatCode,

        [Switch]$Protection

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {

            if ($PSCmdlet.ParameterSetName -eq 'cell')
            {

                $SLStyle = $WorkBookInstance.GetCellStyle($CellReference)

                $StyleHash = @{

                    Alignment  = $SLStyle.Alignment
                    Protection = $SLStyle.Protection
                    FormatCode = $SLStyle.FormatCode
                    Font       = $SLStyle.Font
                    Fill       = $SLStyle.Fill
                    Border     = $SLStyle.Border
                }

                if ($Alignment)
                {
                    $Alignment_props = $SLStyle.Alignment | Get-Member -MemberType Properties | Select-Object -ExpandProperty name
                    $Assigned_AlignMent_Props = $AlignMent_Props | Where-Object { $SLStyle.Alignment.$_ -ne $null }
                    $SLStyle.Alignment | Select-Object $Assigned_AlignMent_Props

                }
                if ($Font)
                {

                    $FontHTMLColor = '#' + $SLStyle.Font.FontColor.Name
                    $SLStyle.Font | Add-Member noteproperty FontHtmlColor $FontHTMLColor -Force
                    $Font_props = $SLStyle.Font | Get-Member -MemberType Properties | Select-Object -ExpandProperty name
                    $Assigned_Font_Props = $Font_Props | Where-Object { $SLStyle.Font.$_ -ne $null }
                    $SLStyle.Font | Select-Object $Assigned_Font_Props

                }

                if ($Fill)
                {
                    $ForegroundColor = '#' + $SLStyle.Fill.PatternForegroundColor.Name
                    $BackgroundColor = '#' + $SLStyle.Fill.PatternBackgroundColor.Name
                    $SLStyle.Fill | Add-Member noteproperty ForegroundColorHTML $ForegroundColor -Force
                    $SLStyle.Fill | Add-Member noteproperty BackgroundColorHTML $BackgroundColor -Force
                    $Fill_props = $SLStyle.fill | Get-Member -MemberType Properties | Select-Object -ExpandProperty name
                    $Fill_props_Gradient = $Fill_props | Where-Object { $_ -match 'gradient' }
                    $Assigned_Fill_Gradient_Props = $Fill_props_Gradient | Where-Object { $SLStyle.Fill.$_ -ne 0 }
                    if ($Assigned_Fill_Gradient_Props)
                    {
                        $SLStyle.fill | Select-Object ForegroundColorHTML, BackgroundColorHTML, PatternType, GradientType, $Assigned_Fill_Gradient_Props
                    }
                    Else
                    {
                        $SLStyle.fill | Select-Object ForegroundColorHTML, BackgroundColorHTML, PatternType, GradientType
                    }

                }
                if ($Border)
                {
                    $LeftBorderColor = '#' + $SLStyle.Border.LeftBorder.Color.Name
                    $RightBorderColor = '#' + $SLStyle.Border.RightBorder.Color.Name
                    $TopBorderColor = '#' + $SLStyle.Border.TopBorder.Color.Name
                    $BottomBorderColor = '#' + $SLStyle.Border.BottomBorder.Color.Name
                    $DiagonalBorderColor = '#' + $SLStyle.Border.DiagonalBorder.Color.Name
                    $VerticalBorderColor = '#' + $SLStyle.Border.VerticalBorder.Color.Name
                    $HorizontalBorderColor = '#' + $SLStyle.Border.HorizontalBorder.Color.Name

                    $LeftBorderStyle = $SLStyle.Border.LeftBorder.BorderStyle
                    $RightBorderStyle = $SLStyle.Border.RightBorder.BorderStyle
                    $TopBorderStyle = $SLStyle.Border.TopBorder.BorderStyle
                    $BottomBorderStyle = $SLStyle.Border.BottomBorder.BorderStyle
                    $DiagonalBorderStyle = $SLStyle.Border.DiagonalBorder.BorderStyle
                    $VerticalBorderStyle = $SLStyle.Border.VerticalBorder.BorderStyle
                    $HorizontalBorderStyle = $SLStyle.Border.HorizontalBorder.BorderStyle


                    $SLStyle.Border | Add-Member noteproperty Left ($LeftBorderStyle , $LeftBorderColor -join ',') -Force
                    $SLStyle.Border | Add-Member noteproperty Right ($RightBorderStyle , $RightBorderColor -join ',') -Force
                    $SLStyle.Border | Add-Member noteproperty Top ($TopBorderStyle , $TopBorderColor -join ',') -Force
                    $SLStyle.Border | Add-Member noteproperty Bottom ($BottomBorderStyle , $BottomBorderColor -join ',') -Force
                    $SLStyle.Border | Add-Member noteproperty Diagonal ($DiagonalBorderStyle , $DiagonalBorderColor -join ',') -Force
                    $SLStyle.Border | Add-Member noteproperty Vertical ($VerticalBorderStyle , $VerticalBorderColor -join ',') -Force
                    $SLStyle.Border | Add-Member noteproperty Horizontal ($HorizontalBorderStyle , $HorizontalBorderColor -join ',') -Force

                    $Border_Noteprops = $SLStyle.Border | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty name
                    $Assigned_Border_Noteprops = $Border_Noteprops | Where-Object { $SLStyle.Border.$_ -notmatch 'none' }

                    #$SLStyle.border | select Left,Right,Top,Bottom,Diagonal,Vertical,Horizontal
                    $SLStyle.Border | Select-Object $Assigned_Border_Noteprops

                }
                if ($FormatCode)
                {
                    $SLStyle.FormatCode

                }
                if ($Protection)
                {
                    $SLStyle.Protection

                }

                #Write-Verbose ("Set-SLFont :`tSetting Font Style on Cell '{0}'" -f $cref)

                $WorkBookInstance | Add-Member NoteProperty CellReference $CellReference -Force
                $WorkBookInstance | Add-Member NoteProperty Style $StyleHash -Force
            }#parameterset cell

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }#select worksheet

    }#Process
    END
    {

    }
}

Function Set-SLCellValue
{

    <#

.SYNOPSIS
    Set a Cell value on a single or a range of cells.

.DESCRIPTION
    Set a Cell value on a single or a range of cells.
    Note: you can only set the same value on multiple cells.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER CellReference
    The target cell whose value has to be set. Eg: A5 or AB10

.PARAMETER Value
    The value to be set.

.PARAMETER Range
    The target cell range that needs to have the specified value. Eg: A5:B10 or AB10:AD20


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLCellValue -CellReference B5,C7  -value "Hello" -Verbose | Save-SLDocument

    Description
    -----------
    Set the value of cells B5 & C7 to "Hello"


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | New-SLFontStyle  -WorksheetName sheet1 -FontName arial -FontSize 10 -FontColor Blue -IsBold  -IsItalic -IsStrikenThrough -Verbose
    PS C:\> $doc | New-SLRichTextStyle  -WorksheetName sheet1 -Text Hello
    PS C:\> $doc | New-SLFontStyle  -WorksheetName sheet1 -FontName arial -FontSize 12 -FontColor red -IsBold -Verbose
    PS C:\> $doc | New-SLRichTextStyle  -WorksheetName sheet1 -Text World -Append
    PS C:\> $doc | Set-SLCellValue -WorksheetName sheet1 -CellReference B6 -SetRichTextStyle -Verbose
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    Set the string Hello Worls as rich text in cell B6.


.INPUTS
   String,SpreadsheetLight.SLDocument

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    N/A

#>
    [CmdletBinding()]
    [OutputType([SpreadsheetLight.SLDocument])]
    param (
        [parameter(Mandatory = $true, position = 0, ValueFromPipeLine = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [Alias('CurrentWorkSheetName')]
        [parameter(Mandatory = $true, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true)]
        [String[]]$CellReference,

        [parameter(Mandatory = $true, Position = 3, ValueFromPipeLineByPropertyName = $true, Parametersetname = 'Value')]
        $value,

        [parameter(Parametersetname = 'RichText')]
        [Switch]$SetRichTextStyle


    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'Singlecell')
            {
                Foreach ($cref in $CellReference)
                {
                    Write-Verbose ("Set-SLCellValue :`tSetting Cell Value '{0}' on Cell '{1}'" -f $Value, $cref)
                    $WorkBookInstance.SetCellValue($Cref, $value) | Out-Null
                }
            }

            if ($PSCmdlet.ParameterSetName -eq 'RichText')
            {
                If ($WorkBookInstance.RichTextStyle)
                {
                    Foreach ($cref in $CellReference)
                    {
                        Write-Verbose ("Set-SLCellValue :`tSetting RichText Style on Cell '{0}'" -f $cref)
                        $WorkBookInstance.SetCellValue($Cref, $WorkBookInstance.RichTextStyle.ToInlineString()) | Out-Null
                    }
                }
                Else
                {
                    Write-Warning ("Set-SLCellValue :`tUse the New-SLFontStyle & New-SLRichTextStyle cmdlets to create font/richtext styles and then apply that style on a cellreference")
                }
            }

            $WorkBookInstance | Add-Member NoteProperty CellReference $CellReference -Force
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force
            Write-Output $WorkBookInstance
        }


    }#process
    END
    {

    }
}


Function Set-SLColumnValue
{

    <#

.SYNOPSIS
    Set column values.

.DESCRIPTION
    Set column values..
    values cannot span multiple columns.Values are set on a single column moving from top to bottom until the value enumeration stops.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER CellReference
    The CellReference that specifies the start row and start column. Eg: A5 or AB10

.PARAMETER Value
    User can specify single or multiple values. Value assignment flow is from top to bottom.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLColumnValue -CellReference b3 -value Jan,feb,march -Verbose | Save-SLDocument

    VERBOSE: Set-SLColumnValue :	Setting value 'Jan' on cell 'b3'
    VERBOSE: Set-SLColumnValue :	Setting value 'feb' on cell 'b4'
    VERBOSE: Set-SLColumnValue :	Setting value 'march' on cell 'b5'

    Description
    -----------
    Since we specified 3 values(jan,feb & march) the cell values start from b3 and flows down to b5.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLColumnValue -CellReference b3 -value Jan,feb,march -Verbose |
                Set-SLBuiltinCellStyle -CellStyle Accent1 -Verbose |
                    Save-SLDocument

    Description
    -----------
    We build on the previous example by setting a cell style:'Accent1' on the cells whose values were set using 'Set-columnValue'.
    Note: Since we piped the output of Set-SLColumnValue we didnt have to specify a worksheetname or cell range with the 'Set-SLBuiltinCellStyle'
    cmdlet because those values are automatically mapped from the "SLdocument" object .

.INPUTS
   String,SpreadsheetLight.SLDocument

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

        [Alias('CurrentWorkSheetName')]
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Set-SLBuiltinCellStyle :`tCellReference should specify values in following format. Eg: A1,B10,AB5..etc"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true)]
        [string]$CellReference,

        [parameter(Mandatory = $true, Position = 3, ValueFromPipelineByPropertyName = $true)]
        $value
    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            $col = [regex]::Match($CellReference, '[a-zA-Z]') | Select-Object -ExpandProperty value
            [int]$Row = [regex]::Match($CellReference, '\d+') | Select-Object -ExpandProperty value

            $StartCellReference = $CellReference

            foreach ($val in $value)
            {
                $CellReference = $col + $row
                Write-Verbose ("Set-SLColumnValue :`tSetting value '{0}' on cell '{1}'" -f $val, $CellReference)

                $WorkBookInstance.SetCellValue($CellReference, $val) | Out-Null
                $row++
            }

            $Range = $StartCellReference + ':' + ($col + ($row - 1))

            $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }
    }#Process
    END
    {

    }
}


Function Set-SLRowValue
{

    <#

.SYNOPSIS
    Set Row values.

.DESCRIPTION
    Set Row values..
    values cannot span multiple rows.Values are set on a single row moving from left to right until the value enumeration stops.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER CellReference
    The CellReference that specifies the start row and start column. Eg: A5 or AB10

.PARAMETER Value
    User can specify single or multiple values. Value assignment flow is from top to bottom.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLRowValue -CellReference b3 -value Jan,feb,march -Verbose | Save-SLDocument

    VERBOSE: Set-SLRowValue :	Setting value 'Jan' on cell 'b3'
    VERBOSE: Set-SLRowValue :	Setting value 'feb' on cell 'c3'
    VERBOSE: Set-SLRowValue :	Setting value 'march' on cell 'd3'

    Description
    -----------
    Since we specified 3 values(jan,feb & march) the cell values start from b3 and flow right as shown above.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLRowValue    -CellReference b3 -value FirstName,LastName,DepartMent -Verbose |
                Set-SLRowValue -CellReference b4 -value Jon,Doe,Sales -Verbose |
                Set-SLRowValue -CellReference b5 -value Zenedine,Zidanne,Football -Verbose |
                Set-SLRowValue -CellReference b6 -value Rahul,Dravid,Cricket -Verbose |
                    Save-SLDocument

    Description
    -----------
    Create a table with 3 rows and 3 columns.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
            $doc | Set-SLRowValue -CellReference b3 -value FirstName,LastName,DepartMent  | Set-SLBuiltinCellStyle -CellStyle Heading3
            $doc | Set-SLRowValue -CellReference b4 -value Jon,Doe,Sales  | Set-SLAlignMent -Vertical Top  | Set-SLBuiltinCellStyle -CellStyle ExplanatoryText
            $doc | Set-SLRowValue -CellReference b5 -value Zenedine,Zidanne,Football  | Set-SLAlignMent -Vertical Top  | Set-SLBuiltinCellStyle -CellStyle ExplanatoryText
            $doc | Set-SLRowValue -CellReference b6 -value Rahul,Dravid,Cricket  | Set-SLAlignMent -Vertical Top | Set-SLBuiltinCellStyle -CellStyle ExplanatoryText
            $doc | Save-SLDocument

    Description
    -----------
    We build on the previous example by applying some alignment and cellstyles to our table.
    Note the above result can be achieved using a gaint piepline but for the sake of legibility the task has been split into various steps.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLRowValue    -CellReference b3 -value FirstName,LastName,DepartMent -Verbose |
                Set-SLRowValue -CellReference b4 -value Jon,Doe,Sales -Verbose |
                Set-SLRowValue -CellReference b5 -value Zenedine,Zidanne,Football -Verbose |
                Set-SLRowValue -CellReference b6 -value Rahul,Dravid,Cricket -Verbose |
                     Set-SLTableStyle -Range B3:D6 -TableStyle Medium17 |
                        Save-SLDocument

    Description
    -----------
    Instead of styling individual rows and columns we can set a tablestyle by specifying the range.


.INPUTS
   String,SpreadsheetLight.SLDocument

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

        [Alias('CurrentWorkSheetName')]
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Set-SLRowValue :`tCellReference should specify values in following format. Eg: A1,B10,AB5..etc"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true)]
        [string]$CellReference,

        [parameter(Mandatory = $true, Position = 3, ValueFromPipelineByPropertyName = $true)]
        $value
    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            $col = [regex]::Match($CellReference, '[a-zA-Z]') | Select-Object -ExpandProperty value
            [int]$Row = [regex]::Match($CellReference, '\d+') | Select-Object -ExpandProperty value

            $StartCellReference = $CellReference
            $colIndex = Convert-ToExcelColumnIndex -ColumnName $col

            foreach ($val in $value)
            {
                $CellReference = (Convert-ToExcelColumnName $colIndex) + $Row
                Write-Verbose ("Set-SLRowValue :`tSetting value '{0}' on cell '{1}'" -f $val, $CellReference)

                $WorkBookInstance.SetCellValue($CellReference, $val) | Out-Null
                $colIndex++
            }

            $Range = $StartCellReference + ':' + ((Convert-ToExcelColumnName -Index ($colIndex - 1)) + $row)

            $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }
    }#Process
    END
    {

    }
}


Function Copy-SLCellValue
{

    <#

.SYNOPSIS
    Copy a single or a range of cell values.

.DESCRIPTION
    Copy a single or a range of cell values. Source data can be on a worksheet that is different than the target.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER FromWorksheetName
    The worksheet containing the source data.

.PARAMETER FromCellReference
    the source cell containing the data to be copied Eg. A3

.PARAMETER Range
    The source data range to be copied Eg. A1:C3

.PARAMETER ToCellreference
    the target cell where data is to be copied to Eg. A3

.PARAMETER ToAnchorCellreference
    The cell reference of the target anchor cell, such as "A1".

.PARAMETER CutorCopy
    Specify whether data is to be copied or pasted

.PARAMETER PasteSpecial
    Specift special paste options such as:
    'Formatting','Formulas','Paste','Values','Transpose'


.Example
    PS C:\> Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx | Copy-SLCellValue -WorksheetName sheet4 -FromCellReference B2 -ToCellreference C2 -Verbose | Save-SLDocument

    Description
    -----------
    Copy cell B2 to C2.


.Example
    PS C:\> Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx |
                Copy-SLCellValue -WorksheetName sheet4 -FromCellReference B2 -ToCellreference D2 -PasteSpecial Formatting -Verbose |
                    Save-SLDocument

    Description
    -----------
    copy only formatting settings from B2 to D2.


.Example
    PS C:\> Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx |
                Copy-SLCellValue -WorksheetName sheet4 -FromCellReference C3 -ToCellreference F2 -CutorCopy Cut -Verbose |
                    Save-SLDocument

    Description
    -----------
    Cut cell C3 and paste it to F2.


.Example
    PS C:\> Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx |
                Copy-SLCellValue -WorksheetName sheet6 -FromWorksheetName sheet4 -FromCellReference B2 -ToCellreference E2 -PasteSpecial Values -Verbose |
                    Save-SLDocument

    Description
    -----------
    Copy B2 from sheet4 and paste it to E2 on sheet6

.Example
    PS C:\> Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx | Copy-SLCellValue -WorksheetName sheet4 -Range A9:C15 -ToAnchorCellreference E9 -Verbose  | Save-SLDocument


    Description
    -----------
    Copy range A9:C15 to E9

.Example
    PS C:\> Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx |
                Copy-SLCellValue -WorksheetName sheet6 -FromWorksheetName sheet4 -Range A9:A15 -ToAnchorCellreference J9 -PasteSpecial  Values -Verbose |
                    Save-SLDocument

    Description
    -----------
    Copy range A9:A15 from sheet4 and paste only the values (ignore any style settings applied to the range) to anchor cell J9 on sheet6.

.Example
    PS C:\> Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx |
                Copy-SLCellValue -WorksheetName sheet6 -FromWorksheetName sheet4 -Range A9:C9 -ToAnchorCellreference N9 -PasteSpecial  Transpose -Verbose |
                    Save-SLDocument

    Description
    -----------
    Copy range A9:A15 from sheet4 and transpose the values (convert row to column and viceversa) to anchor cell N9 on sheet6.

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
        [String]$WorksheetName,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Range-PasteSpecial-DifferentWorksheet')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Range-SimplyCopyPaste-DifferentWorksheet')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleCell-PasteSpecial-DifferentWorksheet')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleCell-SimplyCopyPaste-DifferentWorksheet')]
        [string]$FromWorksheetName,


        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Copy-SLCellValue :`tCellReference should specify values in following format. Eg: A1,B10,AB5..etc"; break }
            })]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleCell-PasteSpecial-DifferentWorksheet')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleCell-SimplyCopyPaste-DifferentWorksheet')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleCell-CutOrCopy')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleCell-PasteSpecial')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleCell-SimplyCopyPaste')]
        [string]$FromCellReference,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+:[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Copy-SLCellValue :`tRange should specify values in following format. Eg: A1:D10 or AB1:AD5"; break }
            })]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Range-PasteSpecial-DifferentWorksheet')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Range-SimplyCopyPaste-DifferentWorksheet')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Range-CutOrCopy')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Range-PasteSpecial')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Range-SimplyCopyPaste')]
        [string]$Range,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Copy-SLCellValue :`tCellReference should specify values in following format. Eg: A1,B10,AB5..etc"; break }
            })]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleCell-PasteSpecial-DifferentWorksheet')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleCell-SimplyCopyPaste-DifferentWorksheet')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleCell-CutOrCopy')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleCell-PasteSpecial')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleCell-SimplyCopyPaste')]
        [String]$ToCellreference,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Copy-SLCellValue :`tCellReference should specify values in following format. Eg: A1,B10,AB5..etc"; break }
            })]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Range-PasteSpecial-DifferentWorksheet')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Range-SimplyCopyPaste-DifferentWorksheet')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Range-CutOrCopy')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Range-PasteSpecial')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Range-SimplyCopyPaste')]
        [String]$ToAnchorCellreference,


        [ValidateSet('Cut', 'Copy')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Range-CutOrCopy')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleCell-CutOrCopy')]
        [String]$CutorCopy,

        [ValidateSet('Formatting', 'Formulas', 'Paste', 'Values', 'Transpose')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Range-PasteSpecial-DifferentWorksheet')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Range-PasteSpecial')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleCell-PasteSpecial-DifferentWorksheet')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleCell-PasteSpecial')]
        [String]$PasteSpecial

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            # paste - just plain paste - Retains Font, Alignment,Fill etc but not data validation
            # values - no formulas or formatting
            # Formulas - Paste values and formulas. Cell references are re-calculated
            # Formatting - only formatting no values - Retains Font, Alignment,Fill etc

            ### - #  SINGLECELL - SameWorksheet
            if ($PSCmdlet.ParameterSetName -eq 'SingleCell-SimplyCopyPaste')
            {
                Write-Verbose ("Copy-SLCellValue :`tCopy cell '{0}' to cell '{1}'" -f $FromCellReference, $ToCellreference)
                # copy one cell to another - Retains Font, Alignment,Fill etc but not data validation
                $WorkBookInstance.CopyCell($FromCellReference, $ToCellreference) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'SingleCell-PasteSpecial')
            {
                Write-Verbose ("Copy-SLCellValue :`tCopy cell '{0}' to cell '{1}' with PasteSpecial Option '{2}'" -f $FromCellReference, $ToCellreference, $PasteSpecial)
                # copy one cell to another with paste option
                $WorkBookInstance.CopyCell($FromCellReference, $ToCellreference, [SpreadsheetLight.SLPasteTypeValues]::$PasteSpecial) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'SingleCell-CutOrCopy')
            {
                if ($CutorCopy -eq 'Cut') { $cutcopyoption = $true }
                else { $cutcopyoption = $false }

                Write-Verbose ("Copy-SLCellValue :`t'{0}' cell '{1}' to cell '{2}'" -f $CutorCopy, $FromCellReference, $ToCellreference)
                # copy one cell to another - Retains Font, Alignment,Fill etc but not data validation
                $WorkBookInstance.CopyCell($FromCellReference, $ToCellreference, $cutcopyoption) | Out-Null
            }

            ### - #  SINGLECELL - DifferentWorksheet
            if ($PSCmdlet.ParameterSetName -eq 'SingleCell-SimplyCopyPaste-DifferentWorksheet')
            {
                Write-Verbose ("Copy-SLCellValue :`tCopy cell '{0}' from Worksheet '{1}' to cell '{2}' on worksheet '{3}'" -f $FromCellReference, $FromWorksheetName, $ToCellreference, $WorksheetName)
                # copy one cell to another - Retains Font, Alignment,Fill etc but not data validation
                $WorkBookInstance.CopyCellFromWorksheet($FromWorksheetName, $FromCellReference, $ToCellreference) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'SingleCell-PasteSpecial-DifferentWorksheet')
            {
                Write-Verbose ("Copy-SLCellValue :`tCopy cell '{0}' from Worksheet '{1}' to cell '{2}' on worksheet '{3}' with PasteSpecial Option '{4}' " -f $FromCellReference, $FromWorksheetName, $ToCellreference, $WorksheetName, $PasteSpecial)
                # copy one cell to another with paste option
                $WorkBookInstance.CopyCellFromWorksheet($FromWorksheetName, $FromCellReference, $ToCellreference, [SpreadsheetLight.SLPasteTypeValues]::$PasteSpecial) | Out-Null
            }


            ### - #  Range - SameWorksheet
            if ($PSCmdlet.ParameterSetName -eq 'Range-SimplyCopyPaste')
            {
                $StartCellReference, $ENDCellReference = $Range -split ':'

                Write-Verbose ("Copy-SLCellValue :`tCopy cellRange '{0}' to cell '{1}'" -f $Range, $ToAnchorCellreference)
                # copy one cell to another - Retains Font, Alignment,Fill etc but not data validation
                $WorkBookInstance.CopyCell($StartCellReference, $ENDCellReference, $ToAnchorCellreference) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Range-PasteSpecial')
            {
                $StartCellReference, $ENDCellReference = $Range -split ':'
                Write-Verbose ("Copy-SLCellValue :`tCopy cellRange '{0}' to cell '{1}' with PasteSpecial Option '{2}'" -f $Range, $ToAnchorCellreference, $PasteSpecial)
                # copy one cell to another with paste option
                $WorkBookInstance.CopyCell($StartCellReference, $ENDCellReference, $ToAnchorCellreference, [SpreadsheetLight.SLPasteTypeValues]::$PasteSpecial) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Range-CutOrCopy')
            {
                if ($CutorCopy -eq 'Cut') { $cutcopyoption = $true }
                else { $cutcopyoption = $false }

                $StartCellReference, $ENDCellReference = $Range -split ':'
                Write-Verbose ("Copy-SLCellValue :`t'{0}' cellRange '{1}' to cell '{2}'" -f $CutorCopy, $Range, $ToAnchorCellreference)
                # copy one cell to another - Retains Font, Alignment,Fill etc but not data validation
                $WorkBookInstance.CopyCell($StartCellReference, $ENDCellReference, $ToAnchorCellreference, $cutcopyoption) | Out-Null
            }

            ### - #  RANGE - DifferentWorksheet
            if ($PSCmdlet.ParameterSetName -eq 'Range-SimplyCopyPaste-DifferentWorksheet')
            {
                $StartCellReference, $ENDCellReference = $Range -split ':'
                Write-Verbose ("Copy-SLCellValue :`tCopy cellrange '{0}' from Worksheet '{1}' to cell '{2}' on worksheet '{3}'" -f $Range, $FromWorksheetName, $ToAnchorCellreference, $WorksheetName)
                # copy one cell to another - Retains Font, Alignment,Fill etc but not data validation
                $WorkBookInstance.CopyCellFromWorksheet($FromWorksheetName, $StartCellReference, $ENDCellReference, $ToAnchorCellreference) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Range-PasteSpecial-DifferentWorksheet')
            {
                $StartCellReference, $ENDCellReference = $Range -split ':'
                Write-Verbose ("Copy-SLCellValue :`tCopy cellrange '{0}' from Worksheet '{1}' to cell '{2}' on worksheet '{3}' with PasteSpecial Option '{4}' " -f $Range, $FromWorksheetName, $ToAnchorCellreference, $WorksheetName, $PasteSpecial)
                # copy one cell to another with paste option
                $WorkBookInstance.CopyCellFromWorksheet($FromWorksheetName, $StartCellReference, $ENDCellReference, $ToAnchorCellreference, [SpreadsheetLight.SLPasteTypeValues]::$PasteSpecial) | Out-Null
            }


            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }#select-slworksheet

    }#process
    END
    {
    }
}


Function Set-SLAlignMent
{

    <#

.SYNOPSIS
    Set text alignment settings on a single or a range of cells.

.DESCRIPTION
    Set text alignment settings on a single or a range of cells.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER CellReference
    The target cell that needs to have the specified alignment settings. Eg: A5 or AB10

.PARAMETER Range
    The target cell range that needs to have the specified alignment settings. Eg: A5:B10 or AB10:AD20

.PARAMETER Vertical
    Valid values for the Vertical alignment parameter is - 'Bottom','Center','Top','Justify','Distributed'.

.PARAMETER Horizontal
    Valid values for the Horizontal alignment parameter is - 'Left','Center','Right','Justify','Distributed'.

.PARAMETER TextRotation
    Specifies the rotation angle of the text, ranging from -90 degrees to 90 degrees.

.PARAMETER Indent
    Each indent value is 3 spaces so an indent value of 5 means 15 spaces wide.

.PARAMETER ShrinkToFit
    Specifies if the text in the cell should be shrunk to fit the cell.

.PARAMETER WrapText
    Specifies if the text in the cell should be wrapped.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLAlignMent -WorksheetName a -cellreference b3 -WrapText -Verbose | Save-SLDocument

    Description
    -----------
    Apply textwrap to cell B3


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLAlignMent -WorksheetName a -cellreference b3 -Vertical Top -WrapText  | Save-SLDocument

    Description
    -----------
    Top align cell content in B3 and then wrap text.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLAlignMent -WorksheetName a -cellreference b3 -indent 3 -TextRotation -80 | Save-SLDocument

    Description
    -----------
    Indent text by 9 spaces and set rotation at 80 degrees.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLAlignMent -WorksheetName a -Range D5:d16 -Horizontal Left -Vertical Center -indent 3 | Save-SLDocument

    Description
    -----------
    Here we apply multiple alignment settings settings to a range of cells.


.INPUTS
   String,SpreadsheetLight.SLDocument

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

        [Alias('CurrentWorkSheetName')]
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning 'CellReference should specify values in following format. Eg: A1,B10,AB5..etc'; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ParameterSetname = 'cell', ValueFromPipeLineByPropertyName = $true)]
        [string[]]$CellReference,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+:[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning 'Range should specify values in following format. Eg: A1:D10 or AB1:AD5'; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true, ParameterSetname = 'Range')]
        [string]$Range,

        [Validateset('Bottom', 'Center', 'Top', 'Justify', 'Distributed')]
        [parameter(Mandatory = $false, ValueFromPipeLineByPropertyName = $true)]
        [String]$Vertical,

        [Validateset('Left', 'Center', 'Right', 'Justify', 'Distributed')]
        [parameter(Mandatory = $false, ValueFromPipeLineByPropertyName = $true)]
        [String]$Horizontal,

        [Validaterange(-90, 90)]
        [parameter(Mandatory = $false, ValueFromPipeLineByPropertyName = $true)]
        [int]$TextRotation,

        [parameter(Mandatory = $false, ValueFromPipeLineByPropertyName = $true)]
        [int]$indent,

        [parameter(Mandatory = $false, ValueFromPipeLineByPropertyName = $true)]
        [switch]$ShrinkToFit,

        [parameter(Mandatory = $false, ValueFromPipeLineByPropertyName = $true)]
        [switch]$WrapText
    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'cell')
            {
                Foreach ($cref in $CellReference)
                {
                    $SLStyle = $WorkBookInstance.GetCellStyle($cref)

                    ## each indent is 3 spaces
                    $SLStyle.Alignment.Indent = $indent

                    if ($ShrinkToFit) { $SLStyle.Alignment.ShrinkToFit = $true }
                    if ($WrapText) { $SLStyle.Alignment.WrapText = $true }
                    if ($TextRotation) { $SLStyle.Alignment.TextRotation = $TextRotation }
                    if ($Vertical) { $SLStyle.Alignment.Vertical = [DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues]::$Vertical }
                    if ($Horizontal) { $SLStyle.Alignment.Horizontal = [DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues]::$Horizontal }

                    Write-Verbose ("Set-SLAlignMent :`tSetting Alignment options on cell '{0}'..." -f $cref)
                    $WorkBookInstance.SetCellStyle($Cref, $SLStyle) | Out-Null
                }
                $WorkBookInstance | Add-Member NoteProperty CellReference $CellReference -Force
            }

            elseif ($PSCmdlet.ParameterSetName -eq 'Range')
            {
                $rowindex, $columnindex = $range -split ':'
                Write-Verbose ("Set-SLAlignMent :`tSetting Alignment options on CellRange '{0}'..." -f $Range)

                $startrowcolumn = Convert-ToExcelRowColumnIndex -CellReference $rowindex
                $endrowcolumn = Convert-ToExcelRowColumnIndex -CellReference $columnindex
                $sRow = $startrowcolumn.Row
                $sColumn = $startrowcolumn.Column
                $eRow = $endrowcolumn.Row
                $eColumn = $endrowcolumn.Column

                $k = 0
                for ($i = $sColumn; $i -le $eColumn; $i++)
                {
                    $Cell = (Convert-ToExcelColumnName -index ($startrowcolumn.Column + $k)) + $sRow

                    $SLStyle = $WorkBookInstance.GetcellStyle($Cell)
                    ## each indent is 3 spaces
                    $SLStyle.Alignment.Indent = $indent

                    if ($ShrinkToFit) { $SLStyle.Alignment.ShrinkToFit = $true }
                    if ($WrapText) { $SLStyle.Alignment.WrapText = $true }
                    if ($TextRotation) { $SLStyle.Alignment.TextRotation = $TextRotation }
                    if ($Vertical) { $SLStyle.Alignment.Vertical = [DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues]::$Vertical }
                    if ($Horizontal) { $SLStyle.Alignment.Horizontal = [DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues]::$Horizontal }
                    $CRCol = ([regex]::Match($cell, '[a-zA-Z]+') | Select-Object -ExpandProperty value) + $erow
                    $WorkBookInstance.SetCellStyle($Cell, $CrCol, $SLStyle) | Out-Null

                    $k++
                }

                $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }#Select-slworksheet

    }#Process
    END
    {

    }
}



Function Set-SLFont
{

    <#

.SYNOPSIS
    Set Font settings on a single or a range of cells.

.DESCRIPTION
    Set Font settings on a single or a range of cells.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER CellReference
    The target cell that needs to have the specified font settings. Eg: A5 or AB10

.PARAMETER Range
    The target cell range that needs to have the specified font settings. Eg: A5:B10 or AB10:AD20

.PARAMETER FontName
    Name of the font.

.PARAMETER FontSize
    Size of the font.

.PARAMETER FontColor
    Color of the font. Use tab completion or intellisense to select a possible value from a list provided by the parameter.

.PARAMETER Underline
    Specifies the underline formatting style of the font text.Valid values are:'Single','Double','SingleAccounting','DoubleAccounting','None'

.PARAMETER IsBold
    Specifies if the font text should be bold.

.PARAMETER IsItalic
    Specifies if the font text should be italic.

.PARAMETER IsStrikenthrough
    Specifies if the font text should have a strikethrough.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLFont -WorksheetName sheet1 -CellReference C15 -Underline Double -IsBold -IsStrikenThrough -Verbose | Save-SLDocument

    Description
    -----------
    Apply Underline,Bold & Strikethrough settings to cell C15


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx |
                Set-SLFont -WorksheetName sheet1 -Range g4:l5  -FontName "Segoe UI" -FontSize 13 -FontColor Chocolate -Verbose | Save-SLDocument

    Description
    -----------
    Apply font settings to a range of cells (g4:l5)


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLCellValue -WorksheetName sheet1 -CellReference B3 -value "Hello" -Verbose |
                Set-SLFont -Underline Double -IsBold -IsItalic -Verbose |
                    Save-SLDocument

    Description
    -----------
    Set the cell value of B3 to 'Hello' and then set the font settings. Notice how we did not have to specify the -worksheetname and -cellreference parameters
    for the 'Set-SLFont' function. This is because we already specified values for those parameters once for the 'Set-SLCellvalue' function so the output
    of this function becomes the input for Set-SLFont.


.INPUTS
   String,SpreadsheetLight.SLDocument

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

        [Alias('CurrentWorkSheetName')]
        [parameter(Mandatory = $true, Position = 1, ValueFromPipelineByPropertyName = $true)]
        [String]$WorksheetName,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Set-SLFont :`tCellReference should specify values in following format. Eg: A1,B10,AB5..etc"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true, ParameterSetname = 'cell')]
        [string[]]$CellReference,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+:[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Set-SLFont :`tRange should specify values in following format. Eg: A1:D10 or AB1:AD5"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true, ParameterSetname = 'Range')]
        [string]$Range,

        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, position = 3)]
        [string]$FontName,

        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, position = 4)]
        [System.UInt16]$FontSize,

        [Validateset('AliceBlue', 'AntiqueWhite', 'Aqua', 'Aquamarine', 'Azure', 'Beige', 'Bisque', 'Black', 'BlanchedAlmond', 'Blue', 'BlueViolet', 'Brown', 'BurlyWood', 'CadetBlue', 'Chartreuse', 'Chocolate', 'Coral', 'CornflowerBlue', 'Cornsilk',
            'Crimson', 'Cyan', 'DarkBlue', 'DarkCyan', 'DarkGoldenrod', 'DarkGray', 'DarkGreen', 'DarkKhaki', 'DarkMagenta', 'DarkOliveGreen', 'DarkOrange', 'DarkOrchid', 'DarkRed', 'DarkSalmon', 'DarkSeaGreen', 'DarkSlateBlue', 'DarkSlateGray',
            'DarkTurquoise', 'DarkViolet', 'DeepPink', 'DeepSkyBlue', 'DimGray', 'DodgerBlue', 'Firebrick', 'FloralWhite', 'ForestGreen', 'Fuchsia', 'Gainsboro', 'GhostWhite', 'Gold', 'Goldenrod', 'Gray', 'Green', 'GreenYellow', 'Honeydew', 'HotPink', 'IndianRed',
            'Indigo', 'Ivory', 'Khaki', 'LavENDer', 'LavENDerBlush', 'LawnGreen', 'LemonChiffon', 'LightBlue', 'LightCoral', 'LightCyan', 'LightGoldenrodYellow', 'LightGray', 'LightGreen', 'LightPink', 'LightSalmon', 'LightSeaGreen', 'LightSkyBlue', 'LightSlateGray',
            'LightSteelBlue', 'LightYellow', 'Lime', 'LimeGreen', 'Linen', 'Magenta', 'Maroon', 'MediumAquamarine', 'MediumBlue', 'MediumOrchid', 'MediumPurple', 'MediumSeaGreen', 'MediumSlateBlue', 'MediumSpringGreen', 'MediumTurquoise', 'MediumVioletRed', 'MidnightBlue'
            , 'MintCream', 'MistyRose', 'Moccasin', 'Name', 'NavajoWhite', 'Navy', 'OldLace', 'Olive', 'OliveDrab', 'Orange', 'OrangeRed', 'Orchid', 'PaleGoldenrod', 'PaleGreen', 'PaleTurquoise', 'PaleVioletRed', 'PapayaWhip', 'PeachPuff', 'Peru', 'Pink', 'Plum', 'PowderBlue',
            'Purple', 'Red', 'RosyBrown', 'RoyalBlue', 'SaddleBrown', 'Salmon', 'SandyBrown', 'SeaGreen', 'SeaShell', 'Sienna', 'Silver', 'SkyBlue', 'SlateBlue', 'SlateGray', 'Snow', 'SpringGreen', 'SteelBlue', 'Tan', 'Teal', 'Thistle', 'Tomato', 'Transparent', 'Turquoise',
            'Violet', 'Wheat', 'White', 'WhiteSmoke', 'Yellow', 'YellowGreen')]
        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, position = 5)]
        [String]$FontColor,

        [Validateset('Single', 'Double', 'SingleAccounting', 'DoubleAccounting', 'None')]
        [parameter(Mandatory = $false)]
        [String]$Underline,

        [switch]$IsBold,

        [switch]$IsItalic,

        [switch]$IsStrikenThrough


    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'cell')
            {
                Foreach ($cref in $CellReference)
                {
                    $SLStyle = $WorkBookInstance.GetCellStyle($cref)

                    if ($isBold) { $SLStyle.Font.Bold = $true }
                    if ($isItalic) { $SLStyle.Font.Italic = $true }
                    if ($IsStrikenThrough) { $SLStyle.Font.Strike = $true }

                    if ($FontName) { $SLStyle.Font.FontName = $FontName }
                    if ($FontSize) { $SLStyle.Font.FontSize = $FontSize }
                    if ($FontColor) { $SLStyle.SetFontColor([System.Drawing.Color]::$FontColor) }
                    if ($Underline) { $SLStyle.Font.Underline = [DocumentFormat.OpenXml.Spreadsheet.UnderlineValues]::$Underline }

                    Write-Verbose ("Set-SLFont :`tSetting Font Style on Cell '{0}'" -f $cref)
                    $WorkBookInstance.SetCellStyle($Cref, $SLStyle) | Out-Null
                }
                $WorkBookInstance | Add-Member NoteProperty CellReference $CellReference -Force
            }
            elseif ($PSCmdlet.ParameterSetName -eq 'Range')
            {
                Write-Verbose ("Set-SLFont :`tSetting Font Style on Cell Range '{0}'" -f $Range)
                $rowindex, $columnindex = $range -split ':'

                $startrowcolumn = Convert-ToExcelRowColumnIndex -CellReference $rowindex
                $endrowcolumn = Convert-ToExcelRowColumnIndex -CellReference $columnindex
                $sRow = $startrowcolumn.Row
                $sColumn = $startrowcolumn.Column
                $eRow = $endrowcolumn.Row
                $eColumn = $endrowcolumn.Column

                $k = 0
                for ($i = $sColumn; $i -le $eColumn; $i++)
                {
                    $Cell = (Convert-ToExcelColumnName -index ($startrowcolumn.Column + $k)) + $sRow

                    $SLStyle = $WorkBookInstance.GetcellStyle($Cell)
                    if ($isBold) { $SLStyle.Font.Bold = $true }
                    if ($isItalic) { $SLStyle.Font.Italic = $true }
                    if ($IsStrikenThrough) { $SLStyle.Font.Strike = $true }

                    if ($FontName) { $SLStyle.Font.FontName = $FontName }
                    if ($FontSize) { $SLStyle.Font.FontSize = $FontSize }
                    if ($FontColor) { $SLStyle.SetFontColor([System.Drawing.Color]::$FontColor) }
                    if ($Underline) { $SLStyle.Font.Underline = [DocumentFormat.OpenXml.Spreadsheet.UnderlineValues]::$Underline }
                    $CRCol = ([regex]::Match($cell, '[a-zA-Z]+') | Select-Object -ExpandProperty value) + $erow
                    $WorkBookInstance.SetCellStyle($Cell, $CrCol, $SLStyle) | Out-Null

                    $k++
                }

                $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
            }#if parameterset range

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }#select worksheet

    }#Process
    END
    {

    }
}

Function Set-SLFill
{

    <#

.SYNOPSIS
    Set Fill settings on a single or a range of cells.

.DESCRIPTION
    Set Fill settings on a single or a range of cells.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER TargetCellorRange
    The target cellreference or Range that needs to have the specified fill settings. Eg: A5 or A5:B10
    Due to the complexity involved in setting up the various fill methods the cellreference and range parameters have been combined as TargetcellorRage.

.PARAMETER Color
    The fill color to be set.

.PARAMETER ColorFromHTML
    The fill color from an HTML string such as '#12b1e6'.

.PARAMETER ThemeColor
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'Accent1Color','Accent2Color','Accent3Color','Accent4Color','Accent5Color',
    'Accent6Color','Dark1Color','Dark2Color','Light1Color',
    'Light2Color','Hyperlink','FollowedHyperlinkColor'

.PARAMETER Pattern
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'DarkDown','DarkGray','DarkGrid','DarkHorizontal','DarkTrellis',
    'DarkUp','DarkVertical','Gray0625','Gray125',
    'LightDown','LightGray','LightGrid','LightHorizontal',
    'LightTrellis','LightUp','LightVertical','MediumGray','None','Solid'

.PARAMETER ForeGroundColor
    The ForeGroundColor fill color to be set. Values are the same as the parameter 'color'.

.PARAMETER BackGroundColor
    The BackGroundColor fill color to be set. Values are the same as the parameter 'color'.

.PARAMETER ForeGroundThemeColor
    The ForeGroundThemeColor fill color to be set. Values are the same as the parameter 'Themecolor'.

.PARAMETER BackGroundThemeColor
    The BackGroundThemeColor fill color to be set. Values are the same as the parameter 'Themecolor'.

.PARAMETER GradientDirection
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'Corner1','Corner2','Corner3','Corner4','DiagonalDown1',
    'DiagonalDown2','DiagonalDown3','DiagonalUp1','DiagonalUp2',
    'DiagonalUp3','Horizontal1','Horizontal2','Horizontal3',
    'Vertical1','Vertical2','Vertical3','FromCenter'


.Example
    PS C:\> Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx | Set-SLFill -WorksheetName sheet6 -TargetCellorRange b2 -Color Aqua -Verbose | Save-SLDocument

    Description
    -----------
    Apply fill color Aqua to cell B2


.Example
    PS C:\> Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx | Set-SLFill -WorksheetName sheet6 -TargetCellorRange b3 -ThemeColor Accent2Color -Verbose | Save-SLDocument

    Description
    -----------
    Apply fill themecolor Accent2Color to cell B3.


.Example
    PS C:\> Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx |
                Set-SLFill -WorksheetName sheet6 -TargetCellorRange b4 -Pattern DarkDown -ForeGroundColor Aquamarine -BackGroundColor AliceBlue -Verbose |
                    Save-SLDocument

    Description
    -----------
    Apply pattern darkdown with two different Foreground and background colors to cell B4.

.Example
    PS C:\> Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx |
                Set-SLFill -WorksheetName sheet6 -TargetCellorRange b5 -Pattern DarkGray -ForeGroundThemeColor Accent1Color -BackGroundThemeColor Accent2Color -Verbose |
                    Save-SLDocument

    Description
    -----------
    Apply pattern darkgray with two different Foreground and background Themecolors to cell B5.

.Example
    PS C:\> Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx |
                Set-SLFill -WorksheetName sheet6 -TargetCellorRange b6 -Pattern DarkGrid -ForeGroundThemeColor Accent2Color -BackGroundColor Brown -Verbose |
                    Save-SLDocument

    Description
    -----------
    Apply pattern darkgrid with a themecolor and a regular color value to cell B6.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Copy-SLCellValue -WorksheetName sheet6 -Range g2:i8 -ToAnchorCellreference g10 -PasteSpecial Values -Verbose
    PS C:\> $doc | Set-SLFont -WorksheetName sheet6 -Range g10:i10 -FontName Tahoma -FontColor White -IsBold -Verbose |
                Set-SLAlignMent -Vertical Center -Horizontal Center |
                     Set-SLFill -ColorFromHTML '#12b1e6'  -Verbose
    PS C:\> $doc | Set-SLFont -WorksheetName sheet6 -Range g11:g16 -FontName Tahoma -FontColor Tan  -Verbose | Set-SLFill -Color Gray
    PS C:\> $doc | Set-SLFill -WorksheetName sheet6 -TargetCellorRange H11:I16  -Color LightGray  -Verbose
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    Copy the range g2:g8 and paste it at G10 filling up cells G10:I16.
    Set font and alignment settings on the header range G10:I16 and apply a fill color '#12b1e6'
    Set a different font and fill color on the first data column. Font Tahoma & color Tan
    To provide contrast apply a light background fill on the remaining data columns H11:I16
    Dont forget to save the document :).

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
        [String]$WorksheetName,

        [Alias('CellReference', 'Range')]
        [parameter(Mandatory = $true, Position = 1, ValueFromPipelineByPropertyName = $true)]
        [string]$TargetCellorRange,

        [Validateset('AliceBlue', 'AntiqueWhite', 'Aqua', 'Aquamarine', 'Azure', 'Beige', 'Bisque', 'Black', 'BlanchedAlmond', 'Blue', 'BlueViolet', 'Brown', 'BurlyWood', 'CadetBlue', 'Chartreuse', 'Chocolate', 'Coral', 'CornflowerBlue', 'Cornsilk',
            'Crimson', 'Cyan', 'DarkBlue', 'DarkCyan', 'DarkGoldenrod', 'DarkGray', 'DarkGreen', 'DarkKhaki', 'DarkMagenta', 'DarkOliveGreen', 'DarkOrange', 'DarkOrchid', 'DarkRed', 'DarkSalmon', 'DarkSeaGreen', 'DarkSlateBlue', 'DarkSlateGray',
            'DarkTurquoise', 'DarkViolet', 'DeepPink', 'DeepSkyBlue', 'DimGray', 'DodgerBlue', 'Firebrick', 'FloralWhite', 'ForestGreen', 'Fuchsia', 'Gainsboro', 'GhostWhite', 'Gold', 'Goldenrod', 'Gray', 'Green', 'GreenYellow', 'Honeydew', 'HotPink', 'IndianRed',
            'Indigo', 'Ivory', 'Khaki', 'LavENDer', 'LavENDerBlush', 'LawnGreen', 'LemonChiffon', 'LightBlue', 'LightCoral', 'LightCyan', 'LightGoldenrodYellow', 'LightGray', 'LightGreen', 'LightPink', 'LightSalmon', 'LightSeaGreen', 'LightSkyBlue', 'LightSlateGray',
            'LightSteelBlue', 'LightYellow', 'Lime', 'LimeGreen', 'Linen', 'Magenta', 'Maroon', 'MediumAquamarine', 'MediumBlue', 'MediumOrchid', 'MediumPurple', 'MediumSeaGreen', 'MediumSlateBlue', 'MediumSpringGreen', 'MediumTurquoise', 'MediumVioletRed', 'MidnightBlue'
            , 'MintCream', 'MistyRose', 'Moccasin', 'Name', 'NavajoWhite', 'Navy', 'OldLace', 'Olive', 'OliveDrab', 'Orange', 'OrangeRed', 'Orchid', 'PaleGoldenrod', 'PaleGreen', 'PaleTurquoise', 'PaleVioletRed', 'PapayaWhip', 'PeachPuff', 'Peru', 'Pink', 'Plum', 'PowderBlue',
            'Purple', 'Red', 'RosyBrown', 'RoyalBlue', 'SaddleBrown', 'Salmon', 'SandyBrown', 'SeaGreen', 'SeaShell', 'Sienna', 'Silver', 'SkyBlue', 'SlateBlue', 'SlateGray', 'Snow', 'SpringGreen', 'SteelBlue', 'Tan', 'Teal', 'Thistle', 'Tomato', 'Transparent', 'Turquoise',
            'Violet', 'Wheat', 'White', 'WhiteSmoke', 'Yellow', 'YellowGreen')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'Pattern1Color')]
        [string]$Color,

        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'Pattern1ColorHtml')]
        [string]$ColorFromHTML,

        [Validateset('Accent1Color', 'Accent2Color', 'Accent3Color', 'Accent4Color', 'Accent5Color',
            'Accent6Color', 'Dark1Color', 'Dark2Color', 'Light1Color',
            'Light2Color', 'Hyperlink', 'FollowedHyperlinkColor')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'Pattern1Theme')]
        [string]$ThemeColor,

        [Validateset('DarkDown', 'DarkGray', 'DarkGrid', 'DarkHorizontal', 'DarkTrellis',
            'DarkUp', 'DarkVertical', 'Gray0625', 'Gray125',
            'LightDown', 'LightGray', 'LightGrid', 'LightHorizontal',
            'LightTrellis', 'LightUp', 'LightVertical', 'MediumGray', 'None', 'Solid')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'Pattern1Theme1Color')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'Pattern2Colors')]
        [parameter(Mandatory = $false, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'Pattern1Color1Theme')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'Pattern2ThemeColors')]
        [string]$Pattern,

        [Validateset('AliceBlue', 'AntiqueWhite', 'Aqua', 'Aquamarine', 'Azure', 'Beige', 'Bisque', 'Black', 'BlanchedAlmond', 'Blue', 'BlueViolet', 'Brown', 'BurlyWood', 'CadetBlue', 'Chartreuse', 'Chocolate', 'Coral', 'CornflowerBlue', 'Cornsilk',
            'Crimson', 'Cyan', 'DarkBlue', 'DarkCyan', 'DarkGoldenrod', 'DarkGray', 'DarkGreen', 'DarkKhaki', 'DarkMagenta', 'DarkOliveGreen', 'DarkOrange', 'DarkOrchid', 'DarkRed', 'DarkSalmon', 'DarkSeaGreen', 'DarkSlateBlue', 'DarkSlateGray',
            'DarkTurquoise', 'DarkViolet', 'DeepPink', 'DeepSkyBlue', 'DimGray', 'DodgerBlue', 'Firebrick', 'FloralWhite', 'ForestGreen', 'Fuchsia', 'Gainsboro', 'GhostWhite', 'Gold', 'Goldenrod', 'Gray', 'Green', 'GreenYellow', 'Honeydew', 'HotPink', 'IndianRed',
            'Indigo', 'Ivory', 'Khaki', 'LavENDer', 'LavENDerBlush', 'LawnGreen', 'LemonChiffon', 'LightBlue', 'LightCoral', 'LightCyan', 'LightGoldenrodYellow', 'LightGray', 'LightGreen', 'LightPink', 'LightSalmon', 'LightSeaGreen', 'LightSkyBlue', 'LightSlateGray',
            'LightSteelBlue', 'LightYellow', 'Lime', 'LimeGreen', 'Linen', 'Magenta', 'Maroon', 'MediumAquamarine', 'MediumBlue', 'MediumOrchid', 'MediumPurple', 'MediumSeaGreen', 'MediumSlateBlue', 'MediumSpringGreen', 'MediumTurquoise', 'MediumVioletRed', 'MidnightBlue'
            , 'MintCream', 'MistyRose', 'Moccasin', 'Name', 'NavajoWhite', 'Navy', 'OldLace', 'Olive', 'OliveDrab', 'Orange', 'OrangeRed', 'Orchid', 'PaleGoldenrod', 'PaleGreen', 'PaleTurquoise', 'PaleVioletRed', 'PapayaWhip', 'PeachPuff', 'Peru', 'Pink', 'Plum', 'PowderBlue',
            'Purple', 'Red', 'RosyBrown', 'RoyalBlue', 'SaddleBrown', 'Salmon', 'SandyBrown', 'SeaGreen', 'SeaShell', 'Sienna', 'Silver', 'SkyBlue', 'SlateBlue', 'SlateGray', 'Snow', 'SpringGreen', 'SteelBlue', 'Tan', 'Teal', 'Thistle', 'Tomato', 'Transparent', 'Turquoise',
            'Violet', 'Wheat', 'White', 'WhiteSmoke', 'Yellow', 'YellowGreen')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'Pattern1Color1Theme')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'Pattern2Colors')]
        [string]$ForeGroundColor,


        [Validateset('AliceBlue', 'AntiqueWhite', 'Aqua', 'Aquamarine', 'Azure', 'Beige', 'Bisque', 'Black', 'BlanchedAlmond', 'Blue', 'BlueViolet', 'Brown', 'BurlyWood', 'CadetBlue', 'Chartreuse', 'Chocolate', 'Coral', 'CornflowerBlue', 'Cornsilk',
            'Crimson', 'Cyan', 'DarkBlue', 'DarkCyan', 'DarkGoldenrod', 'DarkGray', 'DarkGreen', 'DarkKhaki', 'DarkMagenta', 'DarkOliveGreen', 'DarkOrange', 'DarkOrchid', 'DarkRed', 'DarkSalmon', 'DarkSeaGreen', 'DarkSlateBlue', 'DarkSlateGray',
            'DarkTurquoise', 'DarkViolet', 'DeepPink', 'DeepSkyBlue', 'DimGray', 'DodgerBlue', 'Firebrick', 'FloralWhite', 'ForestGreen', 'Fuchsia', 'Gainsboro', 'GhostWhite', 'Gold', 'Goldenrod', 'Gray', 'Green', 'GreenYellow', 'Honeydew', 'HotPink', 'IndianRed',
            'Indigo', 'Ivory', 'Khaki', 'LavENDer', 'LavENDerBlush', 'LawnGreen', 'LemonChiffon', 'LightBlue', 'LightCoral', 'LightCyan', 'LightGoldenrodYellow', 'LightGray', 'LightGreen', 'LightPink', 'LightSalmon', 'LightSeaGreen', 'LightSkyBlue', 'LightSlateGray',
            'LightSteelBlue', 'LightYellow', 'Lime', 'LimeGreen', 'Linen', 'Magenta', 'Maroon', 'MediumAquamarine', 'MediumBlue', 'MediumOrchid', 'MediumPurple', 'MediumSeaGreen', 'MediumSlateBlue', 'MediumSpringGreen', 'MediumTurquoise', 'MediumVioletRed', 'MidnightBlue'
            , 'MintCream', 'MistyRose', 'Moccasin', 'Name', 'NavajoWhite', 'Navy', 'OldLace', 'Olive', 'OliveDrab', 'Orange', 'OrangeRed', 'Orchid', 'PaleGoldenrod', 'PaleGreen', 'PaleTurquoise', 'PaleVioletRed', 'PapayaWhip', 'PeachPuff', 'Peru', 'Pink', 'Plum', 'PowderBlue',
            'Purple', 'Red', 'RosyBrown', 'RoyalBlue', 'SaddleBrown', 'Salmon', 'SandyBrown', 'SeaGreen', 'SeaShell', 'Sienna', 'Silver', 'SkyBlue', 'SlateBlue', 'SlateGray', 'Snow', 'SpringGreen', 'SteelBlue', 'Tan', 'Teal', 'Thistle', 'Tomato', 'Transparent', 'Turquoise',
            'Violet', 'Wheat', 'White', 'WhiteSmoke', 'Yellow', 'YellowGreen')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'Pattern1Theme1Color')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'Pattern2Colors')]
        [string]$BackGroundColor,


        [Validateset('Accent1Color', 'Accent2Color', 'Accent3Color', 'Accent4Color', 'Accent5Color',
            'Accent6Color', 'Dark1Color', 'Dark2Color', 'Light1Color',
            'Light2Color', 'Hyperlink', 'FollowedHyperlinkColor')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'Pattern1Theme1Color')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'Pattern2ThemeColors')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'GradientFill2ThemeColors')]
        [string]$ForeGroundThemeColor,

        [Validateset('Accent1Color', 'Accent2Color', 'Accent3Color', 'Accent4Color', 'Accent5Color',
            'Accent6Color', 'Dark1Color', 'Dark2Color', 'Light1Color',
            'Light2Color', 'Hyperlink', 'FollowedHyperlinkColor')]

        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'Pattern1Color1Theme')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'Pattern2ThemeColors')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'GradientFill2ThemeColors')]
        [string]$BackGroundThemeColor,

        [Validateset('Corner1', 'Corner2', 'Corner3', 'Corner4', 'DiagonalDown1',
            'DiagonalDown2', 'DiagonalDown3', 'DiagonalUp1', 'DiagonalUp2',
            'DiagonalUp3', 'Horizontal1', 'Horizontal2', 'Horizontal3',
            'Vertical1', 'Vertical2', 'Vertical3', 'FromCenter')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'GradientFill2Colors')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'GradientFill2ThemeColors')]
        [string]$GradientDirection


    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {

            Switch -Regex ($TargetCellorRange)
            {

                #CellReference
                '^[a-zA-Z]+\d+$'
                {
                    Write-Verbose ("Set-SLFill :`tTargetCellorRange is CellReference '{0}'" -f $TargetCellorRange)
                    $SLStyle = $WorkBookInstance.GetCellStyle($TargetCellorRange)
                    $isValidationTargetValid = $true
                    $isCellReference = $true
                    Break
                }

                #Range
                '[a-zA-Z]+\d+:[a-zA-Z]+\d+$'
                {
                    $startcellreference, $endcellreference = $TargetCellorRange -split ':'
                    Write-Verbose ("Set-SLFill :`tTargetCellorRange is CellRange '{0}'" -f $TargetCellorRange)
                    $SLStyle = $WorkBookInstance.CreateStyle()
                    $isValidationTargetValid = $true
                    $isRange = $true
                    Break
                }

                Default
                {
                    Write-Warning ("Set-SLDataValidation :`tYou must provide either a Cellreference Eg. C3 or a Range Eg. C3:G10")
                    $isValidationTargetValid = $false
                    Break
                }

            }#switch


            if ($PSCmdlet.ParameterSetName -eq 'Pattern1Theme' -and $isValidationTargetValid )
            {
                Write-Verbose ("Set-SLFill :`tPattern 'Solid' with ThemeColor '{0}' selected" -f $ThemeColor)
                $SLStyle.Fill.SetPatternType([DocumentFormat.OpenXml.Spreadsheet.PatternValues]::'Solid') | Out-Null
                $SLStyle.Fill.SetPatternForegroundColor([SpreadsheetLight.SLThemeColorIndexValues]::$ThemeColor) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Pattern1Color' -and $isValidationTargetValid )
            {
                Write-Verbose ("Set-SLFill :`tPattern 'Solid' with Color '{0}' selected" -f $Color)
                $SLStyle.Fill.SetPatternType([DocumentFormat.OpenXml.Spreadsheet.PatternValues]::'Solid') | Out-Null
                $SLStyle.Fill.SetPatternForegroundColor([System.Drawing.Color]::$Color) | Out-Null
            }


            if ($PSCmdlet.ParameterSetName -eq 'Pattern1ColorHtml' -and $isValidationTargetValid )
            {
                Write-Verbose ("Set-SLFill :`tPattern 'Solid' with HTML Color value '{0}' selected" -f $Color)
                $SLStyle.Fill.SetPatternType([DocumentFormat.OpenXml.Spreadsheet.PatternValues]::'Solid') | Out-Null
                $SLStyle.Fill.SetPatternForegroundColor([System.Drawing.ColorTranslator]::FromHtml($ColorFromHTML))
            }

            if ($PSCmdlet.ParameterSetName -eq 'Pattern2Colors' -and $isValidationTargetValid )
            {
                Write-Verbose ("Set-SLFill :`tPattern '{0}' with ForegroundColor '{1}' & BackGroundColor '{2}' selected" -f $pattern, $ForeGroundColor, $BackGroundColor)
                $SLStyle.Fill.SetPattern([DocumentFormat.OpenXml.Spreadsheet.PatternValues]::$pattern, [System.Drawing.Color]::$ForeGroundColor, [System.Drawing.Color]::$BackGroundColor   ) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Pattern2ThemeColors' -and $isValidationTargetValid )
            {
                Write-Verbose ("Set-SLFill :`tPattern '{0}' with ForegroundThemeColor '{1}' & BackGroundThemeColor '{2}' selected" -f $pattern, $ForeGroundThemeColor, $BackGroundThemeColor)
                $SLStyle.Fill.SetPattern([DocumentFormat.OpenXml.Spreadsheet.PatternValues]::$pattern, [SpreadsheetLight.SLThemeColorIndexValues]::$ForeGroundThemeColor, [SpreadsheetLight.SLThemeColorIndexValues]::$BackGroundThemeColor   ) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Pattern1Theme1Color' -and $isValidationTargetValid )
            {
                Write-Verbose ("Set-SLFill :`tPattern '{0}' with ForegroundThemeColor '{1}' & BackGroundColor '{2}' selected" -f $pattern, $ForeGroundThemeColor, $BackGroundColor)
                $SLStyle.Fill.SetPattern([DocumentFormat.OpenXml.Spreadsheet.PatternValues]::$pattern, [SpreadsheetLight.SLThemeColorIndexValues]::$ForeGroundThemeColor, [System.Drawing.Color]::$BackGroundColor   ) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Pattern1Color1Theme' -and $isValidationTargetValid )
            {
                Write-Verbose ("Set-SLFill :`tPattern '{0}' with ForegroundColor '{1}' & BackGroundThemeColor '{2}' selected" -f $pattern, $ForeGroundColor, $BackGroundThemeColor)
                $SLStyle.Fill.SetPattern([DocumentFormat.OpenXml.Spreadsheet.PatternValues]::$pattern, [System.Drawing.Color]::$ForeGroundColor, [SpreadsheetLight.SLThemeColorIndexValues]::$BackGroundThemeColor   ) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'GradientFill2ThemeColors' -and $isValidationTargetValid )
            {
                Write-Verbose ("Set-SLFill :`tGradientDirection '{0}' with ForeGroundThemeColor '{1}' & BackGroundThemeColor '{2}' selected" -f $GradientDirection, $ForeGroundThemeColor, $BackGroundThemeColor)
                $SLStyle.SetGradientFill([SpreadsheetLight.SLGradientShadingStyleValues]::$GradientDirection, [SpreadsheetLight.SLThemeColorIndexValues]::$ForeGroundThemeColor, [SpreadsheetLight.SLThemeColorIndexValues]::$BackGroundThemeColor) | Out-Null

            }

            if ($PSCmdlet.ParameterSetName -eq 'GradientFill2Colors' -and $isValidationTargetValid )
            {
                Write-Verbose ("Set-SLFill :`tGradientDirection '{0}' with ForeGroundColor '{1}' & BackGroundColor '{2}' selected" -f $GradientDirection, $ForeGroundColor, $BackGroundColor)
                $SLStyle.SetGradientFill([SpreadsheetLight.SLGradientShadingStyleValues]::$GradientDirection, [System.Drawing.Color]::$ForeGroundColor, [System.Drawing.Color]::$BackGroundColor) | Out-Null

            }


            if ($isValidationTargetValid)
            {

                If ($isCellReference)
                {
                    Write-Verbose ("Set-SLFill :`tAdding Fill style..")
                    $WorkBookInstance.SetCellStyle($TargetCellorRange, $SLStyle) | Out-Null
                    $WorkBookInstance | Add-Member NoteProperty CellReference $TargetCellorRange -Force
                }
                Elseif ($isRange)
                {
                    Write-Verbose ("Set-SLFill :`tAdding Fill style..")
                    $WorkBookInstance.SetCellStyle($startcellreference, $endcellreference, $SLStyle) | Out-Null
                    $WorkBookInstance | Add-Member NoteProperty Range $TargetCellorRange -Force
                }
                $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
            }

        }#select-slworksheet

    }#process
}


Function Set-SLBorder
{

    <#

.SYNOPSIS
    Set Border Style on a single or a range of cells.

.DESCRIPTION
    Set Border Style on a single or a range of cells.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER CellReference
    The target cell that needs to have the specified border settings. Eg: A5 or AB10

.PARAMETER Range
    The target cell range that needs to have the specified border settings. Eg: A5:B10 or AB10:AD20

.PARAMETER BorderStyle
    Valid values for the border style parameter are as follows:
    'Thick','Thin','Double','Dotted','Hair','Dashed','DashDot','DashDotDot','SlantDashDot','Medium','MediumDashDot','MediumDashDotDot','MediumDashed'.

.PARAMETER BorderColor
    Use tab completion or intellisense to select a possible value from a list provided by the parameter.

.PARAMETER LeftBorder
    Specify style settings that apply to the left border of a cell or cell range.

.PARAMETER RightBorder
    Specify style settings that apply to the right border of a cell or cell range.

.PARAMETER TopBorder
    Specify style settings that apply to the top border of a cell or cell range.

.PARAMETER BottomBorder
    Specify style settings that apply to the Bottom border of a cell or cell range.

.PARAMETER VerticalBorder
    Specify style settings that apply to the vertical border of a cell or cell range.

.PARAMETER HorizontalBorder
    Specify style settings that apply to the Horizontal border of a cell or cell range.

.PARAMETER CellBorder
    Specify style settings that apply to the whole cell i.e., Left,right,top,bottom.

.PARAMETER DiagonalBorder
    Specify style settings that apply to the DiagonalUp & DiagonalDown border of a cell or cell range.



.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLBorder -WorksheetName sheet1 -CellReference D3 -BorderStyle Double -BorderColor CadetBlue -CellBorder -Verbose | Save-SLDocument

    Description
    -----------
    Apply a border style to cell D3 using the switch '-Cellborder' which means apply the same style to all the borders of the cell D3.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx |
                    Set-SLBorder -WorksheetName sheet1 -Range d15:f24 -BorderStyle Double -BorderColor Blue -CellBorder |
                            Save-SLDocument

    Description
    -----------
    Apply border style to a range of cells (d15:f24)


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx |
                Set-SLBorder -WorksheetName sheet1 -Range d15:f24 -BorderStyle Double -BorderColor CadetBlue -LeftBorder  |
                    Set-SLBorder -BorderStyle Dashed -BorderColor Blue   -RightBorder |
                        Set-SLBorder -BorderStyle Dotted -BorderColor Orange -TopBorder   |
                            Set-SLBorder -BorderStyle Hair   -BorderColor Violet -BottomBorder   |
                                Save-SLDocument

    Description
    -----------
    Here we apply a different border style to each side of the cell range d15:f24.
    Notice that we had to specify the worksheetname and range parameters only once.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx |
                Set-SLBorder -WorksheetName sheet1 -Range d15:d24 -BorderStyle Double -BorderColor CadetBlue -LeftBorder  |
                    Set-SLBorder -Range e15:e24 -BorderStyle Dashed -BorderColor Blue   -RightBorder |
                        Set-SLBorder -Range f15:f24 -BorderStyle Dotted -BorderColor Orange -TopBorder   |
                            Set-SLBorder -Range g15:g24 -BorderStyle Hair   -BorderColor Violet -LeftBorder   |
                                 Save-SLDocument

    Description
    -----------
    Similar to the previous example except that in this case we specify a different border setting for different ranges.
    Notice that we had to specify the worksheetname parameter only once.


.INPUTS
   String,SpreadsheetLight.SLDocument

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

        [Alias('CurrentWorkSheetName')]
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Set-SLBorder :`tCellReference should specify values in following format. Eg: A1,B10,AB5..etc"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true, ParameterSetname = 'cell')]
        [string[]]$CellReference,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+:[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Set-SLBorder :`tRange should specify values in following format. Eg: A1:D10 or AB1:AD5"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true, ParameterSetname = 'Range')]
        [string]$Range,


        [Validateset('Thick', 'Thin', 'Double', 'Dotted', 'Hair', 'Dashed', 'DashDot', 'DashDotDot', 'SlantDashDot', 'Medium', 'MediumDashDot', 'MediumDashDotDot', 'MediumDashed')]
        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [String]$BorderStyle,


        [Validateset('AliceBlue', 'AntiqueWhite', 'Aqua', 'Aquamarine', 'Azure', 'Beige', 'Bisque', 'Black', 'BlanchedAlmond', 'Blue', 'BlueViolet', 'Brown', 'BurlyWood', 'CadetBlue', 'Chartreuse', 'Chocolate', 'Coral', 'CornflowerBlue', 'Cornsilk',
            'Crimson', 'Cyan', 'DarkBlue', 'DarkCyan', 'DarkGoldenrod', 'DarkGray', 'DarkGreen', 'DarkKhaki', 'DarkMagenta', 'DarkOliveGreen', 'DarkOrange', 'DarkOrchid', 'DarkRed', 'DarkSalmon', 'DarkSeaGreen', 'DarkSlateBlue', 'DarkSlateGray',
            'DarkTurquoise', 'DarkViolet', 'DeepPink', 'DeepSkyBlue', 'DimGray', 'DodgerBlue', 'Firebrick', 'FloralWhite', 'ForestGreen', 'Fuchsia', 'Gainsboro', 'GhostWhite', 'Gold', 'Goldenrod', 'Gray', 'Green', 'GreenYellow', 'Honeydew', 'HotPink', 'IndianRed',
            'Indigo', 'Ivory', 'Khaki', 'LavENDer', 'LavENDerBlush', 'LawnGreen', 'LemonChiffon', 'LightBlue', 'LightCoral', 'LightCyan', 'LightGoldenrodYellow', 'LightGray', 'LightGreen', 'LightPink', 'LightSalmon', 'LightSeaGreen', 'LightSkyBlue', 'LightSlateGray',
            'LightSteelBlue', 'LightYellow', 'Lime', 'LimeGreen', 'Linen', 'Magenta', 'Maroon', 'MediumAquamarine', 'MediumBlue', 'MediumOrchid', 'MediumPurple', 'MediumSeaGreen', 'MediumSlateBlue', 'MediumSpringGreen', 'MediumTurquoise', 'MediumVioletRed', 'MidnightBlue'
            , 'MintCream', 'MistyRose', 'Moccasin', 'Name', 'NavajoWhite', 'Navy', 'OldLace', 'Olive', 'OliveDrab', 'Orange', 'OrangeRed', 'Orchid', 'PaleGoldenrod', 'PaleGreen', 'PaleTurquoise', 'PaleVioletRed', 'PapayaWhip', 'PeachPuff', 'Peru', 'Pink', 'Plum', 'PowderBlue',
            'Purple', 'Red', 'RosyBrown', 'RoyalBlue', 'SaddleBrown', 'Salmon', 'SandyBrown', 'SeaGreen', 'SeaShell', 'Sienna', 'Silver', 'SkyBlue', 'SlateBlue', 'SlateGray', 'Snow', 'SpringGreen', 'SteelBlue', 'Tan', 'Teal', 'Thistle', 'Tomato', 'Transparent', 'Turquoise',
            'Violet', 'Wheat', 'White', 'WhiteSmoke', 'Yellow', 'YellowGreen')]
        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [String]$BorderColor,

        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [Switch]$LeftBorder,

        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [Switch]$RightBorder,

        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [Switch]$TopBorder,

        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [Switch]$BottomBorder,

        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [Switch]$VerticalBorder,

        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [Switch]$HoriZontalBorder,

        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [Switch]$CellBorder,

        [parameter(ValueFromPipelineByPropertyName = $true)]
        [switch]$DiagonalBorder





    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'cell')
            {
                Foreach ($cref in $CellReference)
                {
                    $SLStyle = $WorkBookInstance.GetCellStyle($cref)

                    $BStyle = [DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues]::$BorderStyle
                    $BColor = [System.Drawing.Color]::$Bordercolor

                    if ($LeftBorder) { $SLStyle.SetLeftBorder($BStyle, $BColor) }
                    if ($RightBorder) { $SLStyle.SetRightBorder($BStyle, $BColor) }
                    if ($TopBorder) { $SLStyle.SetTopBorder($BStyle, $BColor) }
                    if ($BottomBorder) { $SLStyle.SetBottomBorder($BStyle, $BColor) }
                    if ($VerticalBorder) { $SLStyle.SetVerticalBorder($BStyle, $BColor) }
                    if ($HoriZontalBorder) { $SLStyle.SetHorizontalBorder($BStyle, $BColor) }

                    if ($DiagonalBorder)
                    {
                        $SLStyle.Border.DiagonalUp = $true
                        $SLStyle.Border.DiagonalDown = $true
                        $SLStyle.SetDiagonalBorder($BStyle, $BColor) | Out-Null
                    }

                    if ($CellBorder)
                    {
                        $SLStyle.SetLeftBorder($BStyle, $BColor)
                        $SLStyle.SetRightBorder($BStyle, $BColor)
                        $SLStyle.SetTopBorder($BStyle, $BColor)
                        $SLStyle.SetBottomBorder($BStyle, $BColor)
                    }

                    Write-Verbose ("Set-SLBorder :`tSetting Border Style on Cell '{0}'" -f $cref)
                    $WorkBookInstance.SetCellStyle($Cref, $SLStyle) | Out-Null
                }
                $WorkBookInstance | Add-Member NoteProperty CellReference $CellReference -Force
            }

            elseif ($PSCmdlet.ParameterSetName -eq 'Range')
            {

                $rowindex, $columnindex = $range -split ':'

                $startrowcolumn = Convert-ToExcelRowColumnIndex -CellReference $rowindex
                $endrowcolumn = Convert-ToExcelRowColumnIndex -CellReference $columnindex
                $sRow = $startrowcolumn.Row
                $sColumn = $startrowcolumn.Column
                $eRow = $endrowcolumn.Row
                $eColumn = $endrowcolumn.Column

                $k = 0
                for ($i = $sColumn; $i -le $eColumn; $i++)
                {
                    $Cell = (Convert-ToExcelColumnName -index ($startrowcolumn.Column + $k)) + $sRow
                    $SLStyle = $WorkBookInstance.GetcellStyle($Cell)

                    $BStyle = [DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues]::$BorderStyle
                    $BColor = [System.Drawing.Color]::$Bordercolor

                    if ($DiagonalBorder)
                    {
                        $SLStyle.Border.DiagonalUp = $true
                        $SLStyle.Border.DiagonalDown = $true
                        $SLStyle.SetDiagonalBorder($BStyle, $BColor) | Out-Null
                    }

                    if ($LeftBorder) { $SLStyle.SetLeftBorder($BStyle, $BColor) }
                    if ($RightBorder) { $SLStyle.SetRightBorder($BStyle, $BColor) }
                    if ($TopBorder) { $SLStyle.SetTopBorder($BStyle, $BColor) }
                    if ($BottomBorder) { $SLStyle.SetBottomBorder($BStyle, $BColor) }
                    if ($VerticalBorder) { $SLStyle.SetVerticalBorder($BStyle, $BColor) }
                    if ($HoriZontalBorder) { $SLStyle.SetHorizontalBorder($BStyle, $BColor) }

                    if ($CellBorder)
                    {
                        $SLStyle.SetLeftBorder($BStyle, $BColor)
                        $SLStyle.SetRightBorder($BStyle, $BColor)
                        $SLStyle.SetTopBorder($BStyle, $BColor)
                        $SLStyle.SetBottomBorder($BStyle, $BColor)
                    }

                    $CRCol = ([regex]::Match($cell, '[a-zA-Z]+') | Select-Object -ExpandProperty value) + $erow

                    Write-Verbose ("Set-SLBorder :`tSetting Border Style on Cell Range '{0}'" -f $Range)
                    $WorkBookInstance.SetCellStyle($Cell, $CrCol, $SLStyle) | Out-Null

                    $k++
                }

                $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
            }#if parameterset range

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }#If Select-SLWorksheet

    }#Process
    END
    {

    }
}



Function Set-SLBuiltinCellStyle
{

    <#

.SYNOPSIS
    Apply a style based on the built-in cellstyles.

.DESCRIPTION
    Apply a style based on the built-in cellstyles.
    Applying a cell style will replace any existing cell formatting except for text alignment.
    You may not want to use cell styles if you've added custom formatting to a cell or cells.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER CellReference
    The target cell that needs to have the specified cellstyle. Eg: A5 or AB10

.PARAMETER Range
    The target cell range that needs to have the specified cellstyle. Eg: A5:B10 or AB10:AD20

.PARAMETER CellStyle
    Use tab completion or intellisense to select a possible value from a list provided by the parameter.
    'Normal','Bad','Good','Neutral','Calculation','CheckCell','ExplanatoryText','Input','LinkedCell','Note','Output','WarningText',
        'Heading1','Heading2','Heading3','Heading4','Title','Total','Accent1','Accent2','Accent3','Accent4','Accent5','Accent6',
        'Accent1Percentage60','Accent2Percentage60','Accent3Percentage60','Accent4Percentage60','Accent5Percentage60','Accent6Percentage60',
        'Accent1Percentage40','Accent2Percentage40','Accent3Percentage40','Accent4Percentage40','Accent5Percentage40','Accent6Percentage40',
        'Accent1Percentage20','Accent2Percentage20','Accent3Percentage20','Accent4Percentage20','Accent5Percentage20','Accent6Percentage20',
        'Comma','Comma0','Currency','Currency0','Percentage'


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLBuiltinCellStyle -WorksheetName sheet1 -CellReference B6,C6,D6 -CellStyle Accent2 -Verbose | Save-SLDocument

    Description
    -----------
    Apply a cellstyle anmed 'Accent2' to cells B6,C6,D6.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLBuiltinCellStyle -WorksheetName sheet1 -Range G5:L5 -CellStyle Accent3 -Verbose
    PS C:\> $doc | Set-SLBuiltinCellStyle -WorksheetName sheet1 -Range G6:G7 -CellStyle Accent3Percentage60 -Verbose
    PS C:\> $doc | Set-SLBuiltinCellStyle -WorksheetName sheet1 -Range H6:L7 -CellStyle Accent3Percentage40 -Verbose
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    Apply different cellstyles to a set of cell ranges.
    Note: save-sldocument is called in the last step i.e., after we apply all styles the final step is to save the document.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx |
                Set-SLBuiltinCellStyle -WorksheetName sheet1 -Range G5:L5 -CellStyle Accent3 -Verbose |
                    Set-SLBuiltinCellStyle  -Range G6:G7 -CellStyle Accent3Percentage60 -Verbose |
                        Set-SLBuiltinCellStyle  -Range H6:L7 -CellStyle Accent3Percentage40 -Verbose |
                            Save-SLDocument

    Description
    -----------
    Same as the previous example except that here we use the pipe to apply various styles.
    Note: This is more efficient because it avoids the additional step of assigning the instance to a variable and then piping that variable to apply the style.


.INPUTS
   String,SpreadsheetLight.SLDocument

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

        [Alias('CurrentWorkSheetName')]
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Set-SLBuiltinCellStyle :`tCellReference should specify values in following format. Eg: A1,B10,AB5..etc"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true, ParameterSetname = 'cell')]
        [string[]]$CellReference,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+:[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Set-SLBuiltinCellStyle :`tRange should specify values in following format. Eg: A1:D10 or AB1:AD5"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true, ParameterSetname = 'Range')]
        [string]$Range,

        [ValidateSet('Normal', 'Bad', 'Good', 'Neutral', 'Calculation', 'CheckCell', 'ExplanatoryText', 'Input', 'LinkedCell', 'Note', 'Output', 'WarningText',
            'Heading1', 'Heading2', 'Heading3', 'Heading4', 'Title', 'Total', 'Accent1', 'Accent2', 'Accent3', 'Accent4', 'Accent5', 'Accent6',
            'Accent1Percentage60', 'Accent2Percentage60', 'Accent3Percentage60', 'Accent4Percentage60', 'Accent5Percentage60', 'Accent6Percentage60',
            'Accent1Percentage40', 'Accent2Percentage40', 'Accent3Percentage40', 'Accent4Percentage40', 'Accent5Percentage40', 'Accent6Percentage40',
            'Accent1Percentage20', 'Accent2Percentage20', 'Accent3Percentage20', 'Accent4Percentage20', 'Accent5Percentage20', 'Accent6Percentage20',
            'Comma', 'Comma0', 'Currency', 'Currency0', 'Percentage')]
        [parameter(Mandatory = $true, Position = 3, ValueFromPipelineByPropertyName = $true)]
        [string]$CellStyle

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {

            if ($PSCmdlet.ParameterSetName -eq 'cell')
            {
                Foreach ($cref in $CellReference)
                {
                    Try
                    {
                        Write-Verbose ("Set-SLBuiltinCellStyle :`tSetting Built-In CellStyle '{0}' on Cell '{1}'" -f $CellStyle, $cref)
                        $WorkBookInstance.ApplyNamedCellStyle($cref, [SpreadsheetLight.SLNamedCellStyleValues]::$CellStyle) | Out-Null
                    }
                    Catch
                    {
                        Write-Warning ("Set-SLBuiltinCellStyle :`tPlease check if the specified cellstyle is available on the version of excel installed ...'{0}'" -f $CellStyle)
                    }
                }
                $WorkBookInstance | Add-Member NoteProperty CellReference $CellReference -Force
            }
            if ($PSCmdlet.ParameterSetName -eq 'Range')
            {
                $rowindex, $columnindex = $range -split ':'
                Try
                {
                    Write-Verbose ("Set-SLBuiltinCellStyle :`tSetting Built-In CellStyle '{0}' on Range '{1}'" -f $CellStyle, $Range)
                    $WorkBookInstance.ApplyNamedCellStyle($rowindex, $columnindex, [SpreadsheetLight.SLNamedCellStyleValues]::$CellStyle) | Out-Null
                }
                Catch
                {
                    Write-Warning ("Set-SLBuiltinCellStyle :`tPlease check if the specified cellstyle is available on the version of excel installed ...'{0}'" -f $CellStyle)
                }
                $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }

    }#Process
    END
    {

    }
}

#known issue - style is not applied if the style being copied is applied to another cell in the same row
Function Copy-SLCellStyle
{

    <#

.SYNOPSIS
    Copy a style from a cell and apply it to another cell or a range of cells.

.DESCRIPTION
    Copy a style from a cell and apply it to another cell or a range of cells.
    Note: style can only be copied from a cell and not from a range of cells so the source is always going to be a single cell.
    The target howver can be either a single cell or a range of cells.
    #known issue - style is not applied if the style being copied is applied to another cell in the same row or column.
    Eg: Copy style from G10 to G4 or from G10 to D10 will not work.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER FromCellReference
    The source cell that contains the style to be copied. Eg: A5 or AB10

.PARAMETER ToCellReference
    The target cell that needs to have the copied cellstyle. Eg: A5 or AB10

.PARAMETER Range
    The target cell range that needs to have the copied cellstyle. Eg: A5:B10 or AB10:AD20


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Copy-SLCellStyle -WorksheetName sheet5 -FromCellReference g9 -ToCellReference f3  -Verbose | Save-SLDocument

    Description
    -----------
    Copy cellstyle from G9 and apply it to cell F3.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Copy-SLCellStyle -WorksheetName sheet5 -FromCellReference b10 -Range f4:h6  -Verbose | Save-SLDocument

    Description
    -----------
    Copy style from cell 'g10' and apply to Cell Range 'f4:h6'.



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
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [Alias('CellReference')]
        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Copy-SLCellStyle :`tFromCellReference should specify values in following format. Eg: A1,B10,AB5..etc"; break }
            })]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, position = 2)]
        [string]$FromCellReference,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Copy-SLCellStyle :`tToCellReference should specify values in following format. Eg: A1,B10,AB5..etc"; break }
            })]
        [parameter(Mandatory = $true, Position = 3, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Singlecell')]
        [string]$ToCellReference,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+:[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Copy-SLCellStyle :`tRange should specify values in following format. Eg: A1:D10 or AB1:AD5"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'MultipleCells')]
        [String]$Range

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {

            if ($PSCmdlet.ParameterSetName -eq 'Singlecell')
            {
                Write-Verbose ("Copy-SLCellStyle :`tCopy style from cell '{0}' to Cell '{1}'" -f $FromCellReference, $ToCellReference)
                $WorkBookInstance.CopyCellStyle($FromCellReference, $ToCellReference) | Out-Null
                $WorkBookInstance | Add-Member NoteProperty CellReference $ToCellReference -Force
            }
            elseif ($PSCmdlet.ParameterSetName -eq 'MultipleCells')
            {
                Write-Verbose ("Copy-SLCellStyle :`tCopy style from cell '{0}' and apply to Cell Range '{1}'" -f $FromCellReference, $Range)
                $rowindex, $columnindex = $range -split ':'
                $WorkBookInstance.CopyCellStyle($FromCellReference, $rowindex, $columnindex) | Out-Null
                $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
            }
        }

        $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

    }#process
}


Function Set-SLAutoFitColumn
{

    <#

.SYNOPSIS
    Autofit columns by ColumnName or ColumnIndex.

.DESCRIPTION
    Autofit columns by ColumnName or ColumnIndex. A single or a range of columns by name or index can be specified as input.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER ColumnName
    The columnName to be autofit. Eg: A or G.

.PARAMETER ColumnIndex
    The columnIndex to be autofit. Eg: 1 or 5.

.PARAMETER StartColumnName
    Specifies the start of the autofit column range. Eg: A.

.PARAMETER EndColumnName
    Specifies the end of the autofit column range. Eg: G.

.PARAMETER StartColumnIndex
    Specifies the start of the autofit column range. Eg: 1.

.PARAMETER EndColumnIndex
    Specifies the end of the autofit column range. Eg: 5.

.PARAMETER MaximumColumnWidth
    Specifies the maximum column width for a column or range after autofit is applied. Eg: 10.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLAutoFitColumn -WorksheetName sheet5 -ColumnName F -Verbose | Save-SLDocument

    Description
    -----------
    Autofit column F by Name .


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLAutoFitColumn -WorksheetName sheet5 -StartColumnName F -ENDColumnName H -Verbose | Save-SLDocument

    Description
    -----------
    Autofit columns from F to H by Name.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLAutoFitColumn -WorksheetName sheet5 -StartColumnName F -ENDColumnName H -MaximumColumnWidth 10 -Verbose | Save-SLDocument

    Description
    -----------
    Autofit columns from F to H by Name and optionally set a maxcolumnwidth of 10.

.INPUTS
   String,Int,SpreadsheetLight.SLDocument

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    N/A
#>


    [CmdletBinding(DefaultParameterSetName = 'All')]
    param (
        [parameter(Mandatory = $true, Position = 0, ValueFromPipeLine = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [Alias('CurrentWorkSheetName')]
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [Parameter(Mandatory = $true, Position = 2, ParameterSetName = 'SingleColumnName')]
        [String]$ColumnName,

        [parameter(Mandatory = $true, Position = 2, ParameterSetName = 'SingleColumnIndex')]
        [int]$ColumnIndex,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'MultiPleColumnName')]
        [string]$StartColumnName,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'MultiPleColumnName')]
        [string]$ENDColumnName,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'MultiPleColumnIndex')]
        [int]$StartColumnIndex,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'MultiPleColumnIndex')]
        [int]$ENDColumnIndex,

        [parameter(Mandatory = $false)]
        [Double]$MaximumColumnWidth

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'All')
            {
                Write-Verbose ("Set-SLAutoFitColumn :`tSetting autofit on all columns in worksheet '{0}'" -f $worksheetname)
                $WorkBookInstance.AutoFitColumn('A', 'DD') | Out-Null
            }

            elseif ($PSCmdlet.ParameterSetName -eq 'SingleColumnName')
            {
                if ($MaximumColumnWidth)
                {
                    Write-Verbose ("Set-SLAutoFitColumn :`tSetting autofit on column '{0}' with maxcolumnwidth of '{1}'" -f $ColumnName, $MaximumColumnWidth)
                    $WorkBookInstance.AutoFitColumn($ColumnName, $MaximumColumnWidth) | Out-Null
                }
                Else
                {
                    Write-Verbose ("Set-SLAutoFitColumn :`tSetting autofit on column '{0}'" -f $columnName)
                    $WorkBookInstance.AutoFitColumn($columnName) | Out-Null
                }
            }
            elseif ($PSCmdlet.ParameterSetName -eq 'SingleColumnIndex')
            {
                if ($MaximumColumnWidth)
                {
                    Write-Verbose ("Set-SLAutoFitColumn :`tSetting autofit on column '{0}' with maxcolumnwidth of '{1}'" -f $ColumnIndex, $MaximumColumnWidth)
                    $WorkBookInstance.AutoFitColumn($ColumnIndex, $MaximumColumnWidth) | Out-Null
                }
                Else
                {
                    Write-Verbose ("Set-SLAutoFitColumn :`tSetting autofit on column '{0}'" -f $ColumnIndex)
                    $WorkBookInstance.AutoFitColumn($ColumnIndex) | Out-Null
                }

            }
            elseif ($PSCmdlet.ParameterSetName -eq 'MultiPleColumnName')
            {
                if ($MaximumColumnWidth)
                {
                    Write-Verbose ("Set-SLAutoFitColumn :`tSetting autofit on columns from '{0}' to '{1}' with maxcolumnwidth of '{2}'" -f $StartColumnName, $ENDColumnName, $MaximumColumnWidth)
                    $WorkBookInstance.AutoFitColumn($StartColumnName, $ENDColumnName, $MaximumColumnWidth) | Out-Null
                }
                Else
                {
                    Write-Verbose ("Set-SLAutoFitColumn :`tSetting autofit on columns from '{0}' to '{1}'" -f $StartColumnName, $ENDColumnName)
                    $WorkBookInstance.AutoFitColumn($StartColumnName, $ENDColumnName) | Out-Null
                }

            }
            elseif ($PSCmdlet.ParameterSetName -eq 'MultiPleColumnIndex')
            {
                if ($MaximumColumnWidth)
                {
                    Write-Verbose ("Set-SLAutoFitColumn :`tSetting autofit on columns from '{0}' to '{1}' with maxcolumnwidth of '{2}'" -f $StartColumnIndex, $ENDColumnIndex, $MaximumColumnWidth)
                    $WorkBookInstance.AutoFitColumn($StartColumnIndex, $ENDColumnIndex, $MaximumColumnWidth) | Out-Null
                }
                Else
                {
                    Write-Verbose ("Set-SLAutoFitColumn :`tSetting autofit on columns from '{0}' to '{1}'" -f $StartColumnIndex, $ENDColumnIndex)
                    $WorkBookInstance.AutoFitColumn($StartColumnIndex, $ENDColumnIndex) | Out-Null
                }

            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        } # select-slworksheet
    }#process
}



Function Set-SLAutoFitRow
{

    <#

.SYNOPSIS
    Autofit rows by RowIndex.

.DESCRIPTION
    Autofit columns by RowIndex.A single row or a range of rows can be specified as input.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER RowIndex
    The row to be autofit. Eg: 2 or 5.

.PARAMETER StartRowIndex
    Specifies the start of the autofit row range. Eg: 2.

.PARAMETER EndRowIndex
    Specifies the end of the autofit row range. Eg: 5.

.PARAMETER MaximumRowHeight
    Specifies the maximum row height for a row or a range of rows after autofit is applied. Eg: 10.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLAutoFitRow -WorksheetName sheet5 -RowIndex 3 -Verbose | Save-SLDocument

    Description
    -----------
    Autofit row 3.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLAutoFitRow -WorksheetName sheet5 -StartRowIndex 4 -ENDRowIndex 6 -Verbose | Save-SLDocument

    Description
    -----------
    Autofit rows 4 to 6 by index.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLAutoFitRow -WorksheetName sheet5 -StartRowIndex 4 -ENDRowIndex 6 -MaximumRowHeight 20 -Verbose | Save-SLDocument

    Description
    -----------
    Autofit rows 4 to 6 by index and optionally set a MaximumRowHeight of 20.

.INPUTS
   String,Int,SpreadsheetLight.SLDocument

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    N/A
#>

    [CmdletBinding(DefaultParameterSetName = 'All')]
    param (
        [parameter(Mandatory = $true, Position = 0, ValueFromPipeLine = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [Alias('CurrentWorkSheetName')]
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Position = 2, ParameterSetName = 'SingleRowIndex')]
        [int]$RowIndex,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'MultiPleRowIndex')]
        [int]$StartRowIndex,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'MultiPleRowIndex')]
        [int]$ENDRowIndex,

        [parameter(Mandatory = $false)]
        [Double]$MaximumRowHeight

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'All')
            {
                Write-Verbose ("Set-SLAutoFitRow :`tSetting autofit on the first 2000 rows in worksheet '{0}'" -f $worksheetname)
                $WorkBookInstance.AutoFitRow(1, 2000) | Out-Null
            }

            elseif ($PSCmdlet.ParameterSetName -eq 'SingleRowIndex')
            {
                if ($MaximumRowHeight)
                {
                    Write-Verbose ("Set-SLAutoFitRow :`tSetting autofit on Row '{0}' with MaximumRowHeight of '{1}'" -f $RowIndex, $MaximumRowHeight)
                    $WorkBookInstance.AutoFitRow($RowIndex, $MaximumRowHeight) | Out-Null
                }
                Else
                {
                    Write-Verbose ("Set-SLAutoFitRow :`tSetting autofit on Row '{0}'" -f $RowIndex)
                    $WorkBookInstance.AutoFitRow($RowIndex)
                }
            }

            elseif ($PSCmdlet.ParameterSetName -eq 'MultiPleRowIndex')
            {
                if ($MaximumRowHeight)
                {
                    Write-Verbose ("Set-SLAutoFitRow :`tSetting autofit on Rows '{0}' to '{1}' with MaximumRowHeight of '{2}'" -f $StartRowIndex, $ENDRowIndex, $MaximumRowHeight)
                    $WorkBookInstance.AutoFitRow($StartRowIndex, $ENDRowIndex, $MaximumRowHeight) | Out-Null
                }
                Else
                {
                    Write-Verbose ("Set-SLAutoFitRow :`tSetting autofit on Rows '{0}' to '{1}'" -f $StartRowIndex, $ENDRowIndex)
                    $WorkBookInstance.AutoFitRow($StartRowIndex, $ENDRowIndex) | Out-Null
                }
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }#select-slworksheet
    }#process
}



Function Set-SLColumnWidth
{
    <#

.SYNOPSIS
    Set column width by name or index.

.DESCRIPTION
    Set column width by name or index.A single column or a range of columns can be specified as input.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.


.PARAMETER ColumnName
    The column name. Eg: A or B.

.PARAMETER ColumnIndex
    The column index. Eg: 2 or 5.

.PARAMETER StartColumnName
    Specifies the start of the column range. Eg: A.

.PARAMETER ENDColumnName
    Specifies the end of the column range. Eg: G.

.PARAMETER StartColumnIndex
    Specifies the start index of the column range. Eg: 1 .

.PARAMETER ENDColumnIndex
    Specifies the end index of the column range. Eg: 7.

.PARAMETER ColumnWidth
    Specifies the column width to be applied. Eg: 20.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLColumnWidth -WorksheetName sheet5 -ColumnName f -ColumnWidth 30 -Verbose | Save-SLDocument

    Description
    -----------
    Set columnwidth F to 30.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLColumnWidth -WorksheetName sheet5 -StartColumnName f -ENDColumnName h -ColumnWidth 30 -Verbose | Save-SLDocument

    Description
    -----------
    Set columnwidth of a range F - H to 30.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLColumnWidth -WorksheetName sheet5 -ColumnName f -ColumnWidth 30 -Verbose
    PS C:\> $doc | Set-SLColumnWidth -WorksheetName sheet5 -StartColumnIndex 7 -ENDColumnIndex 8  -ColumnWidth 15 -Verbose
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    Set columnwidth F to 30(Header column). Set columnwidth of column range 7-8 to 20(data columns)

.INPUTS
   String,Int,SpreadsheetLight.SLDocument

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
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [Parameter(Mandatory = $true, Position = 2, ParameterSetName = 'SingleColumnName')]
        [String]$ColumnName,

        [parameter(Mandatory = $true, Position = 2, ParameterSetName = 'SingleColumnIndex')]
        [int]$ColumnIndex,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'MultiPleColumnName')]
        [string]$StartColumnName,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'MultiPleColumnName')]
        [string]$ENDColumnName,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'MultiPleColumnIndex')]
        [int]$StartColumnIndex,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'MultiPleColumnIndex')]
        [int]$ENDColumnIndex,

        [parameter(Mandatory = $true)]
        [Double]$ColumnWidth

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'SingleColumnName')
            {
                Write-Verbose ("Set-SLColumnWidth :`tSetting Width of Column '{0}' to '{1}'" -f $ColumnName, $ColumnWidth)
                $WorkBookInstance.SetColumnWidth($ColumnName, $ColumnWidth) | Out-Null
            }
            elseif ($PSCmdlet.ParameterSetName -eq 'SingleColumnIndex')
            {
                Write-Verbose ("Set-SLColumnWidth :`tSetting Width of Column '{0}' to '{1}'" -f $ColumnIndex, $ColumnWidth)
                $WorkBookInstance.SetColumnWidth($ColumnIndex, $ColumnWidth) | Out-Null
            }
            elseif ($PSCmdlet.ParameterSetName -eq 'MultiPleColumnName')
            {
                Write-Verbose ("Set-SLColumnWidth :`tSetting Width of Columns '{0}' to '{1}' to '{2}' " -f $StartColumnName, $ENDColumnName, $ColumnWidth)
                $WorkBookInstance.SetColumnWidth($StartColumnName, $ENDColumnName, $ColumnWidth) | Out-Null
            }
            elseif ($PSCmdlet.ParameterSetName -eq 'MultiPleColumnIndex')
            {
                Write-Verbose ("Set-SLColumnWidth :`tSetting Width of Columns '{0}' to '{1}' to '{2}' " -f $StartColumnIndex, $ENDColumnIndex, $ColumnWidth)
                $WorkBookInstance.SetColumnWidth($StartColumnIndex, $ENDColumnIndex, $ColumnWidth) | Out-Null
            }


            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }# select-slworksheet
    }#Process
}


Function Set-SLRowHeight
{

    <#

.SYNOPSIS
    Set Row height by index.

.DESCRIPTION
    Set Row height by index.A single row or a range of rows can be specified as input.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER RowIndex
    The row index. Eg: 2 or 5.

.PARAMETER StartRowIndex
    Specifies the start index of the row range. Eg: 1 .

.PARAMETER ENDRowIndex
    Specifies the end index of the row range. Eg: 7.

.PARAMETER RowHeight
    Specifies the row height to be applied. Eg: 20.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLRowHeight -WorksheetName sheet5 -RowIndex 3 -RowHeight 30 -Verbose | Save-SLDocument

    Description
    -----------
    Set Rowheight of row 3 to 30.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLRowHeight -WorksheetName sheet5 -StartRowIndex 4 -ENDRowIndex 6 -RowHeight 15 -Verbose | Save-SLDocument

    Description
    -----------
    Set Rowheight of a range 4 - 6 to 15.


.INPUTS
   String,Int,SpreadsheetLight.SLDocument

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
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Position = 2, ParameterSetName = 'SingleRowIndex')]
        [int]$RowIndex,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'MultiPleRowIndex')]
        [int]$StartRowIndex,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'MultiPleRowIndex')]
        [int]$ENDRowIndex,

        [parameter(Mandatory = $true)]
        [Double]$RowHeight

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'SingleRowIndex')
            {
                Write-Verbose ("Set-SLRowHeight :`tSetting height of row '{0}' to '{1}'" -f $RowIndex, $RowHeight)
                $WorkBookInstance.SetRowHeight($RowIndex, $RowHeight) | Out-Null
            }

            elseif ($PSCmdlet.ParameterSetName -eq 'MultiPleRowIndex')
            {
                Write-Verbose ("Set-SLRowHeight :`tSetting height of rows '{0}' to '{1}' to '{2}'" -f $StartRowIndex, $ENDRowIndex, $RowHeight)
                $WorkBookInstance.SetRowHeight($StartRowIndex, $ENDRowIndex, $RowHeight) | Out-Null
            }
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }#select-slworksheet
    }#process
}

Function Merge-SLCells
{

    <#

.SYNOPSIS
    Merge cells.

.DESCRIPTION
    Merge cells.No merging is done if it's just one cell.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER Range
    The range of cells to be merged. Eg: J3:K4.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Merge-SLCells -WorksheetName sheet5 -Range j3:k4 -Verbose | Save-SLDocument

    Description
    -----------
    Merge cells in the range 'j3:k4'. The content of the first cell is displayed in the merged cell.


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
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+:[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Merge-SLCells :`tRange should specify values in following format. Eg: A1:D10 or AB1:AD5"; break }
            })]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Position = 2, ParameterSetName = 'CellReference')]
        [string]$Range


    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'CellReference')
            {
                Write-Verbose ("Merge-SLCells :`tMerging cells in the range '{0}'" -f $range)
                $StartCellReference, $ENDCellReference = $range -split ':'
                $WorkBookInstance.MergeWorksheetCells($StartCellReference, $ENDCellReference) | Out-Null
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }#select sl-worksheet
    }#process
}

Function UnMerge-SLCells
{

    <#

.SYNOPSIS
    UnMerge cells.

.DESCRIPTION
    UnMerge cells.No Unmerging is done if it's just one cell.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER Range
    The range of cells to be Unmerged. Eg: J3:K4.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | UnMerge-SLCells -WorksheetName sheet5 -Range j3:k4 -Verbose | Save-SLDocument

    Description
    -----------
    UnMerge cells in the range 'j3:k4'.


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
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+:[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Merge-SLCells :`tRange should specify values in following format. Eg: A1:D10 or AB1:AD5"; break }
            })]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Position = 2, ParameterSetName = 'CellReference')]
        [string]$Range


    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'CellReference')
            {
                Write-Verbose ("UnMerge-SLCells :`tUnMerging cells in the range '{0}'" -f $range)
                $StartCellReference, $ENDCellReference = $range -split ':'
                $WorkBookInstance.UnMergeWorksheetCells($StartCellReference, $ENDCellReference) | Out-Null
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }#select sl-worksheet
    }#process
}


Function Set-SLCellFormat
{

    <#

.SYNOPSIS
    Apply string formatting to cells.

.DESCRIPTION
    Apply string formatting to cells.A single or a range of cells can be specified as input

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER CellReference
    The target cell for stringformat. Eg: J3.

.PARAMETER Range
    The target range for stringformat. Eg: J3:K4.

.PARAMETER FormatString
    The format to be set on a particular cell or cells Eg. 'd mm yyyy'.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Set-SLCellValue -WorksheetName sheet5 -CellReference D4 -value 567890789 -Verbose |
                Set-SLCellFormat -FormatString '000\-00\-0000' -Verbose |
                    Save-SLDocument

    Description
    -----------
    Apply stringformat to cell d4. Here the format string - '000\-00\-0000' uses the pattern for matching a "social security number".

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Set-SLCellFormat -WorksheetName sheet5 -Range j3:l5 -FormatString '000\-00\-0000' -Verbose | Save-SLDocument


    Description
    -----------
    Formatstring is applied to a range j3:l5.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLColumnValue -WorksheetName sheet5 -CellReference B3 -value @(123456789.12345,-123456789.12345,(get-date),12.3456,12.3456,123456789.12345) -Verbose
    PS C:\> $doc | Set-SLCellFormat -WorksheetName sheet5 -CellReference B3 -FormatString '#,##0.000' -Verbose
    PS C:\> $doc | Set-SLCellFormat -WorksheetName sheet5 -CellReference B4 -FormatString '$#,##0.00_);[Red]($#,##0.00)' -Verbose
    PS C:\> $doc | Set-SLCellFormat -WorksheetName sheet5 -CellReference B5 -FormatString 'd mmm yyyy' -Verbose
    PS C:\> $doc | Set-SLCellFormat -WorksheetName sheet5 -CellReference B6 -FormatString '0.00%' -Verbose
    PS C:\> $doc | Set-SLCellFormat -WorksheetName sheet5 -CellReference B7 -FormatString '# ??/??' -Verbose
    PS C:\> $doc | Set-SLCellFormat -WorksheetName sheet5 -CellReference B8 -FormatString '0.000E+00' -Verbose
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    Example showing different formats applied to values in the cell range B3:B8.

.INPUTS
   String,SpreadsheetLight.SLDocument

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    http://office.microsoft.com/en-us/excel-help/number-format-codes-HP005198679.aspx
    http://www.databison.com/custom-format-in-excel-how-to-format-numbers-and-text/
#>


    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true, Position = 0, ValueFromPipeLine = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [Alias('CurrentWorkSheetName')]
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Set-SLCellFormat :`tCellReference should specify values in following format. Eg: A1,B10,AB5..etc"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true, ParameterSetname = 'cell')]
        [string[]]$CellReference,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+:[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Set-SLCellFormat :`tRange should specify values in following format. Eg: A1:D10 or AB1:AD5"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true, ParameterSetname = 'Range')]
        [string]$Range,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Position = 3)]
        [string]$FormatString


    )

    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'cell')
            {
                Foreach ($cref in $CellReference)
                {
                    $SLStyle = $WorkBookInstance.GetCellStyle($cref)
                    $SLStyle.FormatCode = $FormatString


                    Write-Verbose ("Set-SLCellFormat :`tSetting FormatString style '{0}' on Cell '{1}'" -f $FormatString, $cref)
                    $WorkBookInstance.SetCellStyle($Cref, $SLStyle) | Out-Null
                }
                $WorkBookInstance | Add-Member NoteProperty CellReference $CellReference -Force
            }
            elseif ($PSCmdlet.ParameterSetName -eq 'Range')
            {
                Write-Verbose ("Set-SLCellFormat :`tSetting FormatString style '{0}' on CellRange '{1}'" -f $FormatString, $Range)
                $StartCellReference, $ENDCellReference = $Range -split ':'

                $SLStyle = $WorkBookInstance.CreateStyle()
                $SLStyle.FormatCode = $FormatString
                $WorkBookInstance.SetCellStyle($StartCellReference, $ENDCellReference, $SLStyle) | Out-Null

                $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
            }#if parameterset range

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }#select worksheet

    }#Process
}

Function Insert-SLColumn
{

    <#

.SYNOPSIS
    Insert columns by name or index.

.DESCRIPTION
    Insert columns by name or index.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER StartColumnName
    The columnName before which columns are to be inserted. Eg: B.

.PARAMETER StartColumnIndex
    The columnIndex before which columns are to be inserted. Eg: 3.

.PARAMETER NumberOfColumns
    The number of columns to be inserted. Eg: 2.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Insert-SLColumn -WorksheetName sheet5 -StartColumnName C -NumberOfColumns 2  -Verbose | Save-SLDocument


    Description
    -----------
    Insert 2 columns before column C.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Insert-SLColumn -WorksheetName sheet5 -StartColumnIndex 3 -NumberOfColumns 2  -Verbose | Save-SLDocument


    Description
    -----------
    Insert 2 columns before column 3(column C).


.INPUTS
   String,Int,SpreadsheetLight.SLDocument

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
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Name')]
        [string]$StartColumnName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'index')]
        [int]$StartColumnIndex,

        [parameter(Mandatory = $true, Position = 3)]
        [int]$NumberOfColumns


    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'Name')
            {
                Write-Verbose ("Insert-SLColumn :`tInserting '{0}' columns before column '{1}' " -f $NumberOfColumns, $StartColumnName)
                $WorkBookInstance.InsertColumn($StartColumnName, $NumberOfColumns) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Index')
            {
                Write-Verbose ("Insert-SLColumn :`tInserting '{0}' columns before column '{1}' " -f $NumberOfColumns, $StartColumnIndex)
                $WorkBookInstance.InsertColumn($StartColumnIndex, $NumberOfColumns) | Out-Null
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }#select-slworksheet

    }#process
    END
    {
    }
}

Function Insert-SLRow
{

    <#

.SYNOPSIS
    Insert rows by index.

.DESCRIPTION
    Insert rows by index.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER StartRowIndex
    The rowIndex before which rows are to be inserted. Eg: 3.

.PARAMETER NumberOfRows
    The number of columns to be inserted. Eg: 2.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Insert-SLRow -WorksheetName sheet5 -StartRowIndex 4 -NumberOfRows 2  -Verbose | Save-SLDocument


    Description
    -----------
    Insert 2 columns before row 4.


.INPUTS
   String,Int,SpreadsheetLight.SLDocument

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
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true)]
        [int]$StartRowIndex,

        [parameter(Mandatory = $true, Position = 3)]
        [int]$NumberOfRows


    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            Write-Verbose ("Insert-SLRow :`tInserting '{0}' Rows before Row '{1}' " -f $NumberOfRows, $StartRowIndex)
            $WorkBookInstance.InsertRow($StartRowIndex, $NumberOfRows) | Out-Null
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }
    }
    END
    {
    }
}

Function Remove-SLRow
{

    <#

.SYNOPSIS
    Delete rows by index.

.DESCRIPTION
    Delete rows by index.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER StartRowIndex
    The rowIndex from which rows are to be deleted. Eg: 2.

.PARAMETER NumberOfRows
    The number of rows to be deleted. Eg: 2.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Remove-SLRow -WorksheetName sheet5 -StartRowIndex 4 -NumberOfRows 2  -Verbose | Save-SLDocument


    Description
    -----------
    Delete 2 rows starting from row 4 and moving down.
    Note: The count starts from the startrowindex so row 4 in this case will be deleted.


.INPUTS
   String,Int,SpreadsheetLight.SLDocument

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
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true)]
        [int]$StartRowIndex,

        [parameter(Mandatory = $true, Position = 3)]
        [int]$NumberOfRows


    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            Write-Verbose ("Remove-SLRow :`tDeleting '{0}' Rows starting from Row '{1}' " -f $NumberOfRows, $StartRowIndex)
            $WorkBookInstance.DeleteRow($StartRowIndex, $NumberOfRows) | Out-Null
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }
    }
    END
    {
    }
}


Function Remove-SLColumn
{

    <#

.SYNOPSIS
    Delete columns by name or index.

.DESCRIPTION
    Delete columns by name or index.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER StartColumnName
    The columnName from which columns are to be deleted. Eg: B.

.PARAMETER StartColumnIndex
    The columnIndex from which columns are to be deleted. Eg: 3.

.PARAMETER NumberOfColumns
    The number of columns to be deleted. Eg: 2.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Remove-SLColumn -WorksheetName sheet5 -StartColumnName C -NumberOfColumns 2  -Verbose | Save-SLDocument


    Description
    -----------
    Delete 2 columns starting from column C.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Remove-SLColumn -WorksheetName sheet5 -StartColumnIndex 3 -NumberOfColumns 2  -Verbose | Save-SLDocument


    Description
    -----------
    Delete 2 columns starting from column 3.


.INPUTS
   String,Int,SpreadsheetLight.SLDocument

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
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Name')]
        [string]$StartColumnName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'index')]
        [int]$StartColumnIndex,

        [parameter(Mandatory = $true, Position = 3)]
        [int]$NumberOfColumns


    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'Name')
            {
                Write-Verbose ("Remove-SLColumn :`tDeleting '{0}' columns starting from column '{1}' " -f $NumberOfColumns, $StartColumnName)
                $WorkBookInstance.DeleteColumn($StartColumnName, $NumberOfColumns) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Index')
            {
                Write-Verbose ("Remove-SLColumn :`tDeleting '{0}' columns starting from column '{1}' " -f $NumberOfColumns, $StartColumnIndex)
                $WorkBookInstance.DeleteColumn($StartColumnIndex, $NumberOfColumns) | Out-Null
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }#select-slworksheet

    }#process
    END
    {
    }
}

Function Hide-SLRow
{

    <#

.SYNOPSIS
    Hide rows by index.

.DESCRIPTION
    Hide rows by index.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER RowIndex
    The rowIndex that specifies the row to be hidden. Eg: 2.

.PARAMETER StartRowIndex
    The rowIndex from which rows are to be hidden. Eg: 2.

.PARAMETER EndRowIndex
    The rowIndex upto which rows are to be hidden. Eg: 4.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Hide-SLRow -WorksheetName sheet5 -RowIndex 4  -Verbose | Save-SLDocument


    Description
    -----------
    Hide row 4.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Hide-SLRow -WorksheetName sheet5 -StartRowIndex 3 -ENDRowIndex 4  -Verbose | Save-SLDocument


    Description
    -----------
    Hide rows 3 & 4.

.INPUTS
   String,Int,SpreadsheetLight.SLDocument

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
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleRow')]
        [int]$RowIndex,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'RangeofRows')]
        [int]$StartRowIndex,

        [parameter(Mandatory = $true, Position = 3, ValueFromPipelineByPropertyName = $true, Parametersetname = 'RangeofRows')]
        [int]$ENDRowIndex


    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'SingleRow')
            {
                Write-Verbose ("Hide-SLRow :`tHiding Row '{0}' from worksheet '{1}' " -f $RowIndex, $WorksheetName)
                $WorkBookInstance.HideRow($RowIndex) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'RangeofRows')
            {
                Write-Verbose ("Hide-SLRow :`tHiding Rows '{0}' to '{1}' " -f $StartRowIndex, $ENDRowIndex)
                $WorkBookInstance.HideRow($StartRowIndex, $ENDRowIndex) | Out-Null
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }#select-slworksheet

    }#process
    END
    {
    }
}


Function Hide-SLColumn
{

    <#

.SYNOPSIS
    Hide columns by name or index.

.DESCRIPTION
    Hide columns by name or index.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER ColumnName
    The columnName to be hidden. Eg: B.

.PARAMETER ColumnIndex
    The columnIndex to be hidden. Eg: 3.

.PARAMETER StartColumnName
    The columnName from which columns are to be hidden. Eg: B.

.PARAMETER EndColumnName
    The columnName upto which columns are to be hidden. Eg: D.

.PARAMETER StartColumnIndex
    The columnIndex from which columns are to be hidden. Eg: 3.

.PARAMETER EndColumnIndex
    The columnIndex upto which columns are to be hidden. Eg: 5.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Hide-SLColumn -WorksheetName sheet5 -ColumnName B  -Verbose | Save-SLDocument


    Description
    -----------
    Hide column B.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Hide-SLColumn -WorksheetName sheet5 -ColumnIndex 3  -Verbose | Save-SLDocument


    Description
    -----------
    Hide column 3(column C).


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Hide-SLColumn -WorksheetName sheet5 -StartColumnName B -ENDColumnName C  -Verbose | Save-SLDocument


    Description
    -----------
    Hide columns B to C.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Hide-SLColumn -WorksheetName sheet5 -StartColumnIndex 4 -ENDColumnIndex 5  -Verbose | Save-SLDocument


    Description
    -----------
    Hide columns 4 to 5.


.INPUTS
   String,Int,SpreadsheetLight.SLDocument

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
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleColumnName')]
        [string]$ColumnName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleColumnIndex')]
        [int]$ColumnIndex,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'RangeofColumnsName')]
        [string]$StartColumnName,

        [parameter(Mandatory = $true, Position = 3, ValueFromPipelineByPropertyName = $true, Parametersetname = 'RangeofColumnsName')]
        [string]$ENDColumnName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'RangeofColumnsIndex')]
        [int]$StartColumnIndex,

        [parameter(Mandatory = $true, Position = 3, ValueFromPipelineByPropertyName = $true, Parametersetname = 'RangeofColumnsIndex')]
        [int]$ENDColumnIndex


    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'SingleColumnIndex')
            {
                Write-Verbose ("Hide-SLColumn :`tHiding Column '{0}' from worksheet '{1}' " -f $ColumnIndex, $WorksheetName)
                $WorkBookInstance.HideColumn($ColumnIndex) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'RangeofColumnsIndex')
            {
                Write-Verbose ("Hide-SLColumn :`tHiding Columns '{0}' to '{1}' " -f $StartColumnIndex, $ENDColumnIndex)
                $WorkBookInstance.HideColumn($StartColumnIndex, $ENDColumnIndex ) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'SingleColumnName')
            {
                Write-Verbose ("Hide-SLColumn :`tHiding Column '{0}' from worksheet '{1}' " -f $ColumnName, $WorksheetName)
                $WorkBookInstance.HideColumn($ColumnName) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'RangeofColumnsName')
            {
                Write-Verbose ("Hide-SLColumn :`tHiding Columns '{0}' to '{1}' " -f $StartColumnName, $ENDColumnName)
                $WorkBookInstance.HideColumn($StartColumnName, $ENDColumnName ) | Out-Null
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }#select-slworksheet
    }#process
    END
    {
    }
}


Function Show-SLRow
{

    <#

.SYNOPSIS
    UnHide rows by index.

.DESCRIPTION
    UnHide rows by index.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER RowIndex
    The rowIndex that specifies the row to be shown. Eg: 2.

.PARAMETER StartRowIndex
    The rowIndex from which rows are to be shown. Eg: 2.

.PARAMETER EndRowIndex
    The rowIndex upto which rows are to be shown. Eg: 4.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Show-SLRow -WorksheetName sheet5 -RowIndex 4  -Verbose | Save-SLDocument


    Description
    -----------
    UnHide row 4.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Show-SLRow -WorksheetName sheet5 -StartRowIndex 3 -ENDRowIndex 4  -Verbose | Save-SLDocument


    Description
    -----------
    UnHide rows 3 & 4.

.INPUTS
   String,Int,SpreadsheetLight.SLDocument

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
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleRow')]
        [int]$RowIndex,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'RangeofRows')]
        [int]$StartRowIndex,

        [parameter(Mandatory = $true, Position = 3, ValueFromPipelineByPropertyName = $true, Parametersetname = 'RangeofRows')]
        [int]$ENDRowIndex


    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'SingleRow')
            {
                Write-Verbose ("Show-SLRow :`tUn-Hiding Row '{0}' from worksheet '{1}' " -f $RowIndex, $WorksheetName)
                $WorkBookInstance.UnhideRow($RowIndex) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'RangeofRows')
            {
                Write-Verbose ("Show-SLRow :`tUn-Hiding Rows '{0}' to '{1}' " -f $StartRowIndex, $ENDRowIndex)
                $WorkBookInstance.UnhideRow($StartRowIndex, $ENDRowIndex) | Out-Null
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }#select-slworksheet

    }#process
    END
    {
    }
}


Function Show-SLColumn
{

    <#

.SYNOPSIS
    Un-Hide columns by name or index.

.DESCRIPTION
    Un-Hide columns by name or index.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER ColumnName
    The columnName to be shown. Eg: B.

.PARAMETER ColumnIndex
    The columnIndex to be shown. Eg: 3.

.PARAMETER StartColumnName
    The columnName from which columns are to be shown. Eg: B.

.PARAMETER EndColumnName
    The columnName upto which columns are to be shown. Eg: D.

.PARAMETER StartColumnIndex
    The columnIndex from which columns are to be shown. Eg: 3.

.PARAMETER EndColumnIndex
    The columnIndex upto which columns are to be shown. Eg: 5.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Show-SLColumn -WorksheetName sheet5 -ColumnName B  -Verbose | Save-SLDocument


    Description
    -----------
    UnHide column B.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Show-SLColumn -WorksheetName sheet5 -ColumnIndex 3  -Verbose | Save-SLDocument


    Description
    -----------
    UnHide column 3(column C).


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Show-SLColumn -WorksheetName sheet5 -StartColumnName B -ENDColumnName C  -Verbose | Save-SLDocument


    Description
    -----------
    UnHide columns B to C.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Show-SLColumn -WorksheetName sheet5 -StartColumnIndex 4 -ENDColumnIndex 5  -Verbose | Save-SLDocument


    Description
    -----------
    UnHide columns 4 to 5.


.INPUTS
   String,Int,SpreadsheetLight.SLDocument

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
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleColumnName')]
        [string]$ColumnName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'SingleColumnIndex')]
        [int]$ColumnIndex,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'RangeofColumnsName')]
        [string]$StartColumnName,

        [parameter(Mandatory = $true, Position = 3, ValueFromPipelineByPropertyName = $true, Parametersetname = 'RangeofColumnsName')]
        [string]$ENDColumnName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'RangeofColumnsIndex')]
        [int]$StartColumnIndex,

        [parameter(Mandatory = $true, Position = 3, ValueFromPipelineByPropertyName = $true, Parametersetname = 'RangeofColumnsIndex')]
        [int]$ENDColumnIndex


    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'SingleColumnIndex')
            {
                Write-Verbose ("Show-SLColumn :`tUn-Hiding Column '{0}' from worksheet '{1}' " -f $ColumnIndex, $WorksheetName)
                $WorkBookInstance.UnhideColumn($ColumnIndex) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'RangeofColumnsIndex')
            {
                Write-Verbose ("Show-SLColumn :`tUn-Hiding Columns '{0}' to '{1}' " -f $StartColumnIndex, $ENDColumnIndex)
                $WorkBookInstance.UnhideColumn($StartColumnIndex, $ENDColumnIndex ) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'SingleColumnName')
            {
                Write-Verbose ("Show-SLColumn :`tUn-Hiding Column '{0}' from worksheet '{1}' " -f $ColumnName, $WorksheetName)
                $WorkBookInstance.UnhideColumn($ColumnName) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'RangeofColumnsName')
            {
                Write-Verbose ("Show-SLColumn :`tUn-Hiding Columns '{0}' to '{1}' " -f $StartColumnName, $ENDColumnName)
                $WorkBookInstance.UnhideColumn($StartColumnName, $ENDColumnName ) | Out-Null
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }#select-slworksheet
    }#process
    END
    {
    }
}

Function Group-SLRow
{

    <#

.SYNOPSIS
    Group rows by index.

.DESCRIPTION
    Group rows by index.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER StartRowIndex
    The rowIndex from which rows are to be grouped. Eg: 2.

.PARAMETER EndRowIndex
    The rowIndex upto which rows are to be grouped. Eg: 4.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Group-SLRow -WorksheetName sheet5 -StartRowIndex 4 -ENDRowIndex 6 -Verbose | Save-SLDocument


    Description
    -----------
    Group Rows 4 to 6.


.INPUTS
   String,Int,SpreadsheetLight.SLDocument

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
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true)]
        [int]$StartRowIndex,

        [parameter(Mandatory = $true, Position = 3)]
        [int]$ENDRowIndex


    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            Write-Verbose ("Group-SLRow :`tGrouping Rows '{0}' to '{1}' " -f $StartRowIndex, $ENDRowIndex)
            $WorkBookInstance.GroupRows($StartRowIndex, $ENDRowIndex) | Out-Null
            $WorkBookInstance.CollapseRows(($ENDRowIndex + 1 )) | Out-Null
        }

        $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
    }
    END
    {
    }
}


Function UnGroup-SLRow
{

    <#

.SYNOPSIS
    UnGroup rows by index.

.DESCRIPTION
    UnGroup rows by index.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER StartRowIndex
    The rowIndex from which rows are to be ungrouped. Eg: 2.

.PARAMETER EndRowIndex
    The rowIndex upto which rows are to be ungrouped. Eg: 4.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | UnGroup-SLRow -WorksheetName sheet5 -StartRowIndex 4 -ENDRowIndex 6 -Verbose | Save-SLDocument


    Description
    -----------
    UnGroup Rows 4 to 6.


.INPUTS
   String,Int,SpreadsheetLight.SLDocument

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
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true)]
        [int]$StartRowIndex,

        [parameter(Mandatory = $true, Position = 3)]
        [int]$ENDRowIndex


    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            Write-Verbose ("Group-SLRow :`tUnGrouping Rows '{0}' to '{1}' " -f $StartRowIndex, $ENDRowIndex)
            $WorkBookInstance.UnGroupRows($StartRowIndex, $ENDRowIndex) | Out-Null
            $WorkBookInstance.UnhideRow($StartRowIndex, $ENDRowIndex) | Out-Null
        }

        $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
    }
    END
    {
    }
}

Function Group-SLColumn
{

    <#

.SYNOPSIS
    Group columns by name or index.

.DESCRIPTION
    Group columns by name or index.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER StartColumnName
    The columnName from which columns are to be Grouped. Eg: B.

.PARAMETER EndColumnName
    The columnName upto which columns are to be Grouped. Eg: D.

.PARAMETER StartColumnIndex
    The columnIndex from which columns are to be Grouped. Eg: 3.

.PARAMETER EndColumnIndex
    The columnIndex upto which columns are to be Grouped. Eg: 5.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Group-SLColumn -WorksheetName sheet5 -StartColumnName F -ENDColumnName H  -Verbose | Save-SLDocument


    Description
    -----------
    Group columns F to H.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Group-SLColumn -WorksheetName sheet5 -StartColumnIndex 6 -ENDColumnIndex 8  -Verbose | Save-SLDocument


    Description
    -----------
    Group columns 6 to 8.


.INPUTS
   String,Int,SpreadsheetLight.SLDocument

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
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Name')]
        [string]$StartColumnName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Name')]
        [string]$ENDColumnName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'index')]
        [int]$StartColumnIndex,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'index')]
        [int]$ENDColumnIndex

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'Name')
            {
                Write-Verbose ("Group-SLColumn :`tGrouping Columns '{0}' to '{1}' " -f $StartColumnName, $ENDColumnName)
                $WorkBookInstance.GroupColumns($StartColumnName, $ENDColumnName) | Out-Null
                $WorkBookInstance.CollapseColumns(((Convert-ToExcelColumnIndex $ENDColumnName) + 1)) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Index')
            {
                Write-Verbose ("Group-SLColumn :`tGrouping Columns '{0}' to '{1}' " -f $StartColumnIndex, $ENDColumnIndex)
                $WorkBookInstance.GroupColumns($StartColumnIndex, $ENDColumnIndex) | Out-Null
                $WorkBookInstance.CollapseColumns(( $ENDColumnIndex + 1)) | Out-Null
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }#select-slworksheet
    }#process
    END
    {
    }
}

Function UnGroup-SLColumn
{

    <#

.SYNOPSIS
    UnGroup columns by name or index.

.DESCRIPTION
    UnGroup columns by name or index.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER StartColumnName
    The columnName from which columns are to be UnGrouped. Eg: B.

.PARAMETER EndColumnName
    The columnName upto which columns are to be UnGrouped. Eg: D.

.PARAMETER StartColumnIndex
    The columnIndex from which columns are to be UnGrouped. Eg: 3.

.PARAMETER EndColumnIndex
    The columnIndex upto which columns are to be UnGrouped. Eg: 5.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | UnGroup-SLColumn -WorksheetName sheet5 -StartColumnName F -ENDColumnName H  -Verbose | Save-SLDocument


    Description
    -----------
    UnGroup columns F to H.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | UnGroup-SLColumn -WorksheetName sheet5 -StartColumnIndex 6 -ENDColumnIndex 8  -Verbose | Save-SLDocument


    Description
    -----------
    UnGroup columns 6 to 8.


.INPUTS
   String,Int,SpreadsheetLight.SLDocument

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
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Name')]
        [string]$StartColumnName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Name')]
        [string]$ENDColumnName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'index')]
        [int]$StartColumnIndex,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'index')]
        [int]$ENDColumnIndex



    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'Name')
            {
                Write-Verbose ("UnGroup-SLColumn :`tUnGrouping Columns '{0}' to '{1}' " -f $StartColumnName, $ENDColumnName)
                $WorkBookInstance.UngroupColumns($StartColumnName, $ENDColumnName) | Out-Null
                $WorkBookInstance.UnhideColumn($StartColumnName, $ENDColumnName) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Index')
            {
                Write-Verbose ("UnGroup-SLColumn :`tUnGrouping Columns '{0}' to '{1}' " -f $StartColumnIndex, $ENDColumnIndex)
                $WorkBookInstance.UngroupColumns($StartColumnIndex, $ENDColumnIndex) | Out-Null
                $WorkBookInstance.UnhideColumn($StartColumnIndex, $ENDColumnIndex) | Out-Null
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }#select-slworksheet
    }#process
    END
    {
    }
}

Function Collapse-SLRow
{

    <#

.SYNOPSIS
    Collapse rows by index.

.DESCRIPTION
    Collapse rows by index.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER RowIndex
    The row index of the row just after the group of rows you want to collapse.
    For example, this will be row 5 if rows 2 to 4 are grouped.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Collapse-SLRow -WorksheetName sheet5 -RowIndex 7  -Verbose | Save-SLDocument


    Description
    -----------
    Collapse Row 7.


.INPUTS
   String,Int,SpreadsheetLight.SLDocument

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
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true)]
        [int]$RowIndex

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            Write-Verbose ("Collapse-SLRow :`tCollapsing Row '{0}' " -f $RowIndex)
            $WorkBookInstance.CollapseRows($RowIndex) | Out-Null
        }

        $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

    }#process
    END
    {
    }
}

Function Expand-SLRow
{

    <#

.SYNOPSIS
    Expand rows by index.

.DESCRIPTION
    Expand rows by index.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER RowIndex
    The row index of the row just after the group of rows you want to collapse.
    For example, this will be row 5 if rows 2 to 4 are grouped.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Expand-SLRow   -WorksheetName sheet5 -RowIndex 7  -Verbose | Save-SLDocument


    Description
    -----------
    Expand Row 7.


.INPUTS
   String,Int,SpreadsheetLight.SLDocument

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
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true)]
        [int]$RowIndex

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            Write-Verbose ("Expand-SLRow :`tExpanding Row '{0}' " -f $RowIndex)
            $WorkBookInstance.ExpandRows($RowIndex) | Out-Null
        }

        $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

    }#process
    END
    {
    }
}


Function Collapse-SLColumn
{

    <#

.SYNOPSIS
    Collapse columns by name or index.

.DESCRIPTION
    Collapse columns by name or index.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER ColumnName
    The column name (such as "A1") of the column just after the group of columns you want to collapse.
    For example, this will be column E if columns B to D are grouped.

.PARAMETER ColumnIndex
    The column index of the column just after the group of columns you want to collapse.
    For example, this will be column 5 if columns 2 to 4 are grouped.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Collapse-SLColumn -WorksheetName sheet5 -ColumnName I  -Verbose | Save-SLDocument


    Description
    -----------
    Collapse column I.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Collapse-SLColumn -WorksheetName sheet5 -ColumnIndex 9  -Verbose | Save-SLDocument


    Description
    -----------
    Collapse column 9.


.INPUTS
   String,Int,SpreadsheetLight.SLDocument

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
        [parameter(Mandatory = $false, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'index')]
        [int]$ColumnIndex,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Name')]
        [string]$ColumnName

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'Name')
            {
                Write-Verbose ("Collapse-SLColumn :`tCollapsing column '{0}' " -f $ColumnName)
                $WorkBookInstance.CollapseColumns($ColumnName) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Index')
            {
                Write-Verbose ("Collapse-SLColumn :`tCollapsing column '{0}' " -f $ColumnIndex)
                $WorkBookInstance.CollapseColumns($ColumnIndex) | Out-Null
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }#select-slworksheet
    }#process
    END
    {
    }
}


Function Expand-SLColumn
{

    <#

.SYNOPSIS
    Expand columns by name or index.

.DESCRIPTION
    Expand columns by name or index.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER ColumnName
    The column name (such as "A1") of the column just after the group of columns you want to expand.
    For example, this will be column E if columns B to D are grouped.

.PARAMETER ColumnIndex
    The column index of the column just after the group of columns you want to expand.
    For example, this will be column 5 if columns 2 to 4 are grouped.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Expand-SLColumn   -WorksheetName sheet5 -ColumnName I  -Verbose | Save-SLDocument


    Description
    -----------
    Expand column I.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Expand-SLColumn   -WorksheetName sheet5 -ColumnIndex 9  -Verbose | Save-SLDocument


    Description
    -----------
    Expand column 9.


.INPUTS
   String,Int,SpreadsheetLight.SLDocument

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
        [String]$WorksheetName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'index')]
        [int]$ColumnIndex,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true, Parametersetname = 'Name')]
        [string]$ColumnName

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'Name')
            {
                Write-Verbose ("Expand-SLColumn :`tExpanding column '{0}' " -f $ColumnName)
                $WorkBookInstance.ExpandColumns($ColumnName) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Index')
            {
                Write-Verbose ("Expand-SLColumn :`tExpanding column '{0}' " -f $ColumnIndex)
                $WorkBookInstance.ExpandColumns($ColumnIndex) | Out-Null
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }#select-slworksheet
    }#process
    END
    {
    }
}

Function Set-SLSplitPane
{

    <#

.SYNOPSIS
    set up Split pane.

.DESCRIPTION
    set up Split pane.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER NumberOfRows
    Number of top-most rows above the horizontal split line.

.PARAMETER NumberOfColumns
    Number of left-most columns left of the vertical split line.

.PARAMETER ShowRowColumnHeadings
    If included in the parameterlist row and column headings are shown. False otherwise.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Set-SLSplitPane -WorksheetName sheet5 -NumberOfRows 3 -NumberOfColumns 8 -ShowRowColumnHeadings -Verbose  | Save-SLDocument


    Description
    -----------
    Top-left pane is '3' Rows high and '8' columns wide. Headers shown - 'True'



.INPUTS
   String,Int,SpreadsheetLight.SLDocument

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
        [String]$WorksheetName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true)]
        [int]$NumberOfRows,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true)]
        [int]$NumberOfColumns,

        [parameter(Mandatory = $false, Position = 2)]
        [Switch]$ShowRowColumnHeadings


    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($ShowRowColumnHeadings) { $Headershown = $true } else { $Headershown = $false }
            Write-Verbose ("Set-SLSplitPane :`tTop-left pane is '{0}' Rows high and '{1}' columns wide. Headers shown - '{2}' " -f $NumberOfRows, $NumberOfColumns, $Headershown)

            $WorkBookInstance.SplitPanes($NumberOfRows, $NumberOfColumns, $Headershown) | Out-Null
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }
    }
    END
    {
    }
}

Function Set-SLFreezePane
{

    <#

.SYNOPSIS
    set up Split pane.

.DESCRIPTION
    set up Split pane.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER NumberOfTopMostRows
    Number of top-most rows to keep in place.

.PARAMETER NumberOfLeftMostColumns
   Number of left-most columns to keep in place.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Set-SLFreezePane -WorksheetName sheet5 -NumberOfTopMostRows 3 -NumberOfLeftMostColumns 8 -Verbose  | Save-SLDocument


    Description
    -----------
    Top-left pane is '3' Rows high and '8' columns wide.


.INPUTS
   String,Int,SpreadsheetLight.SLDocument

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
        [String]$WorksheetName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true)]
        [int]$NumberOfTopMostRows,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true)]
        [int]$NumberOfLeftMostColumns

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            Write-Verbose ("Set-SLFreezePane :`tTop-left pane is '{0}' Rows high and '{1}' columns wide. " -f $NumberOfTopMostRows, $NumberOfLeftMostColumns)
            $WorkBookInstance.FreezePanes($NumberOfTopMostRows, $NumberOfLeftMostColumns) | Out-Null
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }
    }
    END
    {
    }
}

Function Set-SLDataFilter
{

    <#

.SYNOPSIS
    Set Autofilter on a cellrange.

.DESCRIPTION
    Set Autofilter on a cellrange.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER Range
    cellrange which needs to be filtered.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Set-SLDataFilter -WorksheetName sheet5 -Range F3:H6 -Verbose  | Save-SLDocument


    Description
    -----------
    Filter data in the range F3:H6.


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
        [String]$WorksheetName,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+:[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Set-SLDataFilter :`tRange should specify values in following format. Eg: A1:D10 or AB1:AD5"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true)]
        [String]$Range


    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            Write-Verbose ("Set-SLDataFilter :`tSetting autofilter on Cellrange '{0}'. " -f $Range)
            $StartCellReference, $Endcellreference = $range -split ':'
            $WorkBookInstance.Filter($StartCellReference, $Endcellreference) | Out-Null

            $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }

    }
    END
    {
    }
}


Function Sort-SLData
{

    <#

.SYNOPSIS
    Sort data by row or column.

.DESCRIPTION
    Sort data by row or column.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER Range
    cellrange which needs to be sorted.

.PARAMETER ColumnNameToSortBy
    The column to be sorted Eg. A.

.PARAMETER RowIndexToSortBy
    The rowindex to be sorted Eg. 5.

.PARAMETER SortOrder
    Specify the sort order as either : 'Ascending or Descending'.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Sort-SLData -WorksheetName sheet5 -Range F4:H6 -ColumnNameToSortBy H -SortOrder ASCending  -Verbose  | Save-SLDocument


    Description
    -----------
    sort data in the range F4:H6 by column H in the Ascending order.


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
        [String]$WorksheetName,


        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+:[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Sort-SLData :`tRange should specify values in following format. Eg: A1:D10 or AB1:AD5"; break }
            })]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [String]$Range,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetname = 'ColumnSort')]
        [String]$ColumnNameToSortBy,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetname = 'RowSort')]
        [String]$RowIndexToSortBy,

        [ValidateSet('ASCending', 'DESCending')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [String]$SortOrder

    )
    PROCESS
    {

        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($SortOrder -eq 'AscENDing') { $SortOrderbool = $true }
            Else { $SortOrderbool = $false }

            if ($PSCmdlet.ParameterSetName -eq 'ColumnSort')
            {
                Write-Verbose ("Sort-SLData :`tSorting Cellrange '{0}' by the column '{1}' in the '{2}' order" -f $Range, $ColumnNameToSortBy, $SortOrder)
                $startcellreference, $ENDcellreference = $range -split ':'
                $WorkBookInstance.sort($startcellreference, $ENDcellreference, $ColumnNameToSortBy, $SortOrderbool)

                $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
            }

            Elseif ($PSCmdlet.ParameterSetName -eq 'RowSort')
            {
                <#
                Write-Verbose ("Sort-SLData :`tSorting Cellrange '{0}' by the RowIndex '{1}' in the '{2}' order" -f $Range,$RowIndexToSortBy,$SortOrder)
                $startcellreference,$ENDcellreference = $range -split ":"
                $WorkBookInstance.sort($startcellreference,$ENDcellreference,$RowIndexToSortBy,$SortOrderbool)

                $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
                #>
                Write-Warning "Sort-SLData :`tSorting by row is currently not working. Will be fixed shortly."
            }


            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }#select-slworksheet

    }#process
    END
    {
    }
}

Function Insert-SLPageBreak
{

    <#

.SYNOPSIS
    Insert pagebreaks.

.DESCRIPTION
    Insert pagebreaks.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER AboveRowIndex
    Row index.

.PARAMETER LeftofColumnIndex
    Column Index.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Insert-SLPageBreak -WorksheetName sheet2 -AboveRowIndex 9 -Verbose  | Save-SLDocument


    Description
    -----------
    Insert page break above row 9.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Insert-SLPageBreak -WorksheetName sheet2 -LeftofColumnIndex 5 -Verbose  | Save-SLDocument


    Description
    -----------
    Insert page break to the left of column 5.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Insert-SLPageBreak -WorksheetName sheet2 -AboveRowIndex 6 -LeftofColumnIndex 6 -Verbose  | Save-SLDocument


    Description
    -----------
    Insert page break above row 6 and to the left of column 6.

.INPUTS
   String,Int,SpreadsheetLight.SLDocument

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
        [String]$WorksheetName,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetname = 'Row')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetname = 'RowColumn')]
        [Int]$AboveRowIndex,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetname = 'Column')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetname = 'RowColumn')]
        [Int]$LeftofColumnIndex

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {

            if ($PSCmdlet.ParameterSetName -eq 'RowColumn')
            {
                Write-Verbose ("Insert-SLPageBreak :`tInsert pagebreak above Row '{0}' and to the left of column '{1}'" -f $AboveRowIndex, $LeftofColumnIndex)
                $WorkBookInstance.InsertPageBreak($AboveRowIndex, $LeftofColumnIndex) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Row')
            {
                Write-Verbose ("Insert-SLPageBreak :`tInsert pagebreak above Row '{0}'" -f $AboveRowIndex)
                $WorkBookInstance.InsertPageBreak($AboveRowIndex, -1) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Column')
            {
                Write-Verbose ("Insert-SLPageBreak :`tInsert pagebreak to the left of column '{0}'" -f $LeftofColumnIndex)
                $WorkBookInstance.InsertPageBreak(-1, $LeftofColumnIndex) | Out-Null
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }

    }#process
    END
    {
    }
}

Function Remove-SLPageBreak
{

    <#

.SYNOPSIS
    Remove pagebreaks.

.DESCRIPTION
    Remove pagebreaks.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER AboveRowIndex
    Row index.

.PARAMETER LeftofColumnIndex
    Column Index.

.PARAMETER All
    If specified Will remove all page breaks from a worksheet .

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Remove-SLPageBreak -WorksheetName sheet2 -AboveRowIndex 9 -Verbose  | Save-SLDocument


    Description
    -----------
    remove page break above row 9.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Remove-SLPageBreak -WorksheetName sheet2 -LeftofColumnIndex 5 -Verbose  | Save-SLDocument


    Description
    -----------
    remove page break to the left of column 5.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Remove-SLPageBreak -WorksheetName sheet2 -AboveRowIndex 6 -LeftofColumnIndex 6 -Verbose  | Save-SLDocument


    Description
    -----------
    remove page break above row 6 and to the left of column 6.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Remove-SLPageBreak -WorksheetName sheet2 -All  -Verbose  | Save-SLDocument


    Description
    -----------
    remove all page breaks in worksheet 'sheet2'.

.INPUTS
   String,Int,SpreadsheetLight.SLDocument

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
        [String]$WorksheetName,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetname = 'Row')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetname = 'RowColumn')]
        [Int]$AboveRowIndex,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetname = 'Column')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetname = 'RowColumn')]
        [Int]$LeftofColumnIndex,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetname = 'all')]
        [Switch]$All

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($PSCmdlet.ParameterSetName -eq 'RowColumn')
            {
                Write-Verbose ("Remove-SLPageBreak :`tRemoving pagebreak above Row '{0}' and to the left of column '{1}'" -f $AboveRowIndex, $LeftofColumnIndex)
                $WorkBookInstance.RemovePageBreak($AboveRowIndex, $LeftofColumnIndex) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Row')
            {
                Write-Verbose ("Remove-SLPageBreak :`tRemoving pagebreak above Row '{0}'" -f $AboveRowIndex)
                $WorkBookInstance.RemovePageBreak($AboveRowIndex, -1) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'Column')
            {
                Write-Verbose ("Remove-SLPageBreak :`tRemoving pagebreak to the left of column '{0}'" -f $LeftofColumnIndex)
                $WorkBookInstance.RemovePageBreak(-1, $LeftofColumnIndex) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'all')
            {
                Write-Verbose ("Remove-SLPageBreak :`tRemoving all pagebreaks in the worksheet '{0}'" -f $WorksheetName)
                $WorkBookInstance.RemoveAllPageBreaks() | Out-Null
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }#select-slworksheet

    }#process
    END
    {
    }
}


Function Protect-SLWorksheet
{

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

Function UnProtect-SLWorksheet
{

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

Function Set-SLWorksheetTabColor
{

    <#

.SYNOPSIS
    Sets the tab color of a worksheet.

.DESCRIPTION
    Sets the tab color of a worksheet.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER Color
    Any color that is to be used eg: Red.

.PARAMETER ThemeColor
    Theme color to be used. Valid values are:
    'Light1Color','Dark1Color','Light2Color','Dark2Color','Accent1Color','Accent2Color','Accent3Color',
    'Accent4Color','Accent5Color','Accent6Color','Hyperlink','FollowedHyperlinkColor'

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Set-SLWorksheetTabColor -WorksheetName sheet2 -Color Yellow   -Verbose  | Save-SLDocument


    Description
    -----------
    Set the tab color of sheet2 to yellow.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Set-SLWorksheetTabColor -WorksheetName sheet2 -ThemeColor Accent2Color   -Verbose  | Save-SLDocument



    Description
    -----------
    Set the tab color of sheet2 to Accent2Color.


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
        [String]$WorksheetName,

        [parameter(Mandatory = $True, Position = 2, ParameterSetName = 'Color')]
        [string]$Color,

        [ValidateSet('Light1Color', 'Dark1Color', 'Light2Color', 'Dark2Color', 'Accent1Color', 'Accent2Color', 'Accent3Color', 'Accent4Color', 'Accent5Color', 'Accent6Color', 'Hyperlink', 'FollowedHyperlinkColor')]
        [parameter(Mandatory = $True, Position = 2, ParameterSetName = 'ThemeColor')]
        [string]$ThemeColor

    )
    PROCESS
    {

        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            $PageSettings = $WorkBookInstance.GetPageSettings()

            if ($PSCmdlet.ParameterSetName -eq 'Color')
            {
                Write-Verbose ("Set-SLWorksheetTabColor :`tSet worksheet '{0}' tab color to '{1}'" -f $WorksheetName, $Color)
                $PageSettings.TabColor = [System.Drawing.Color]::$color
            }
            if ($PSCmdlet.ParameterSetName -eq 'ThemeColor')
            {
                Write-Verbose ("Set-SLWorksheetTabColor :`tSet worksheet '{0}' tab color to '{1}'" -f $WorksheetName, $ThemeColor)
                $PageSettings.SetTabColor([SpreadsheetLight.SLThemeColorIndexValues]::$ThemeColor)
            }

            $WorkBookInstance.SetPageSettings($PageSettings) | Out-Null
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }

    }#process
    END
    {
    }
}


Function New-SLDefinedName
{

    <#

.SYNOPSIS
    Create a Defined Name for a cell reference, range, constant, formula, or table.

.DESCRIPTION
    A name is a meaningful shorthand that makes it easier to understand the purpose of a cell reference,
    constant, formula, or table, each of which may be difficult to comprehend at first glance.
    The following information shows common examples of names and how they can improve clarity and understanding.

    EXAMPLE TYPE	EXAMPLE WITH NO NAME	                EXAMPLE WITH A NAME
    Reference	    =SUM(C20:C30)	                        =SUM(FirstQuarterSales)
    Constant	    =PRODUCT(A5,8.3)	                    =PRODUCT(Price,WASalesTax)
    Formula	        =SUM(VLOOKUP(A1,B1:F20,5,FALSE), -G5)	=SUM(Inventory_Level,-Order_Amt)
    Table	        C4:G36	                                =TopSales06

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that contains the Range referenced by the defined name.

.PARAMETER DefinedName
    A userfriendly Name for a cell reference, range, constant, formula, or table.

.PARAMETER Range
    cellrange which would be the datasource for a defined name.
    To define a cellreference instead of a range use the range format like so: B3:B3

.PARAMETER Comment
    Comment that provides a short description of the defined name.

.PARAMETER Scope
    The name of the worksheet that the defined name is effective in.

.PARAMETER Force
    If the defined name to be created already exists in the workbook use the force switch to overwrite the existing value.
    By default the cmdlet will not overwrite an existing Defined Name.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | New-SLDefinedName -WorksheetName sheet1 -DefinedName DFName1 -Range B3:B7 -Verbose | Save-SLDocument


    Description
    -----------
    Create a New defined name 'DFName1'.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\DataValidation.xlsx |
                New-SLDefinedName -WorksheetName sheet1 -DefinedName DFName2 -Range B3:B7 -Comment "This range represents Athlete Names" |
                    Save-SLDocument



    Description
    -----------
    Create a New defined name 'DFName1'.Additionally specify a comment to describe the defined name.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\DataValidation.xlsx |
                New-SLDefinedName -WorksheetName sheet1 -DefinedName DFName3 -Range B3:B7 -Comment "This range represents Athlete Names" -Scope sheet2 |
                    Save-SLDocument



    Description
    -----------
    Create a New defined name 'DFName1'.Additionally specify a comment and scope.
    Because we specified 'sheet2' as the value for the scope parameter, the defined name 'DFName3' can only be used on worksheet named 'sheet2'.


.INPUTS
   String,SpreadsheetLight.SLDocument

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    http://office.microsoft.com/en-in/excel-help/define-and-use-names-in-formulas-HA010147120.aspx

#>


    [CmdletBinding(DefaultParameterSetName = 'All', SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param (
        [parameter(Mandatory = $true, Position = 0, ValueFromPipeLine = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [Alias('CurrentWorkSheetName')]
        [parameter(Mandatory = $true, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true)]
        [String]$DefinedName,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+:[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "New-SLDefinedName :`tRange should specify values in following format. Eg: A1:D10 or AB1:AD5"; break }
            })]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$Range,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetname = 'Comment')]
        [string]$Comment,

        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, ParameterSetname = 'Comment', HelpMessage = 'The name of the worksheet that the defined name is effective in')]
        [string]$Scope,

        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [Switch]$Force

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {

            $DefinedNames = $WorkBookInstance.GetDefinedNames() | Select-Object -ExpandProperty Name
            if ($DefinedNames -contains $DefinedName)
            {
                $DefinedNameExists = $true
            }
            Else
            {
                $DefinedNameExists = $false
            }
            $AbsoluteRange = Convert-ToExcelAbsoluteRange -Range $Range -WorkSheetName $WorksheetName

            If ($PSCmdlet.ParameterSetName -eq 'Comment')
            {
                if ($Scope)
                {
                    If ($DefinedNameExists -and $Force -and $PSCmdlet.ShouldPROCESS($DefinedName, 'OVERWRITE DEFINED NAME') )
                    {
                        Write-Verbose ("New-SLDefinedName :`tForce parameter specified. Overwriting existing DefinedName '{0}'" -f $DefinedName)
                        Write-Verbose ("New-SLDefinedName :`tDefinedName '{0}' Scope is '{1}'" -f $DefinedName, $Scope)
                        Write-Verbose ("New-SLDefinedName :`tDefinedName '{0}' Comment is '{1}'" -f $DefinedName, $Comment)
                        Write-Verbose ("New-SLDefinedName :`tCreating DefinedName '{0}' corresponding to Range '{1}'" -f $DefinedName, $Range)
                        $WorkBookInstance.SetDefinedName($DefinedName, $AbsoluteRange, $Comment, $Scope) | Out-Null
                        $WorkBookInstance | Add-Member NoteProperty DefinedName $DefinedName -Force
                        $WorkBookInstance | Add-Member NoteProperty DefinedNameRange $AbsoluteRange -Force
                    }
                    Elseif ($DefinedNameExists)
                    {
                        Write-Warning ("New-SLDefinedName :`tDefinedName '{0}' Already exists. Specify the '-Force' parameter to Overwrite" -f $DefinedName)
                    }
                    Else
                    {
                        Write-Verbose ("New-SLDefinedName :`tDefinedName '{0}' Scope is '{1}'" -f $DefinedName, $Scope)
                        Write-Verbose ("New-SLDefinedName :`tDefinedName '{0}' Comment is '{1}'" -f $DefinedName, $Comment)
                        Write-Verbose ("New-SLDefinedName :`tCreating DefinedName '{0}' corresponding to Range '{1}'" -f $DefinedName, $Range)
                        $WorkBookInstance.SetDefinedName($DefinedName, $AbsoluteRange, $Comment, $Scope) | Out-Null
                        $WorkBookInstance | Add-Member NoteProperty DefinedName $DefinedName -Force
                        $WorkBookInstance | Add-Member NoteProperty DefinedNameRange $AbsoluteRange -Force
                    }
                }
                Else
                {
                    If ($DefinedNameExists -and $Force -and $PSCmdlet.ShouldPROCESS($DefinedName, 'OVERWRITE DEFINED NAME') )
                    {
                        Write-Verbose ("New-SLDefinedName :`tForce parameter specified. Overwriting existing DefinedName '{0}'" -f $DefinedName)
                        Write-Verbose ("New-SLDefinedName :`tDefinedName '{0}' Comment is '{1}'" -f $DefinedName, $Comment)
                        Write-Verbose ("New-SLDefinedName :`tCreating DefinedName '{0}' corresponding to Range '{1}'" -f $DefinedName, $Range)
                        $WorkBookInstance.SetDefinedName($DefinedName, $AbsoluteRange, $Comment) | Out-Null
                        $WorkBookInstance | Add-Member NoteProperty DefinedName $DefinedName -Force
                        $WorkBookInstance | Add-Member NoteProperty DefinedNameRange $AbsoluteRange -Force
                    }
                    Elseif ($DefinedNameExists)
                    {
                        Write-Warning ("New-SLDefinedName :`tDefinedName '{0}' Already exists. Specify the '-Force' parameter to Overwrite" -f $DefinedName)
                    }
                    Else
                    {
                        Write-Verbose ("New-SLDefinedName :`tDefinedName '{0}' Comment is '{1}'" -f $DefinedName, $Comment)
                        Write-Verbose ("New-SLDefinedName :`tCreating DefinedName '{0}' corresponding to Range '{1}'" -f $DefinedName, $Range)
                        $WorkBookInstance.SetDefinedName($DefinedName, $AbsoluteRange, $Comment) | Out-Null
                        $WorkBookInstance | Add-Member NoteProperty DefinedName $DefinedName -Force
                        $WorkBookInstance | Add-Member NoteProperty DefinedNameRange $AbsoluteRange -Force
                    }
                }
            }
            elseIf ($PSCmdlet.ParameterSetName -eq 'All')
            {
                If ($DefinedNameExists -and $Force -and $PSCmdlet.ShouldPROCESS($DefinedName, 'OVERWRITE DEFINED NAME') )
                {
                    Write-Verbose ("New-SLDefinedName :`tForce parameter specified. Overwriting existing DefinedName '{0}'" -f $DefinedName)
                    Write-Verbose ("New-SLDefinedName :`tCreating DefinedName '{0}' corresponding to Range '{1}'" -f $DefinedName, $Range)
                    $WorkBookInstance.SetDefinedName($DefinedName, $AbsoluteRange) | Out-Null
                    $WorkBookInstance | Add-Member NoteProperty DefinedName $DefinedName -Force
                    $WorkBookInstance | Add-Member NoteProperty DefinedNameRange $AbsoluteRange -Force
                }
                Elseif ($DefinedNameExists)
                {
                    Write-Warning ("New-SLDefinedName :`tDefinedName '{0}' Already exists. Specify the '-Force' parameter to Overwrite" -f $DefinedName)
                }
                Else
                {
                    Write-Verbose ("New-SLDefinedName :`tCreating DefinedName '{0}' corresponding to Range '{1}'" -f $DefinedName, $Range)
                    $WorkBookInstance.SetDefinedName($DefinedName, $AbsoluteRange) | Out-Null
                    $WorkBookInstance | Add-Member NoteProperty DefinedName $DefinedName -Force
                    $WorkBookInstance | Add-Member NoteProperty DefinedNameRange $AbsoluteRange -Force

                }
            }

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }

    }
    END
    {
    }
}


Function Get-SLDefinedName
{

    <#

.SYNOPSIS
    Lists defined names contained in an excel document.

.DESCRIPTION
    Lists defined names contained in an excel document.The properties associated with a defined name are:
    Name,Text,Comment & LocalsheetID(scope).

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER Filter
    Gets matching defined names. Filter can be a string or a regex pattern.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Get-SLDefinedName


    Description
    -----------
    Will list all defined names in document 'myfirstdoc'.

.Example

    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Get-SLDefinedName -Filter DFname


    Description
    -----------
    Will list all defined names matching the string 'dfname'.


.INPUTS
   String,SpreadsheetLight.SLDocument

.OUTPUTS
   String

.Link
    N/A

#>


    [CmdletBinding(DefaultParameterSetName = 'All')]
    param (
        [parameter(Mandatory = $true, Position = 0, ValueFromPipeLine = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [parameter(Mandatory = $true, Position = 1, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'DefinedName')]
        [String]$DefinedName

    )
    PROCESS
    {
        $DefinedNames = $WorkBookInstance.GetDefinedNames()

        if ($DefinedNames)
        {
            $DefinedNamesExist = $true
        }
        Else
        {
            $DefinedNamesExist = $false
            Write-Warning ("Get-SLDefinedName :`tThe specified workbook did not contain any DefinedNames")
            break
        }

        if ($PSCmdlet.ParameterSetName -eq 'DefinedName')
        {

            if ($DefinedNames.name -contains $DefinedName)
            {
                $DefinedNameExists = $true
                $DefinedNames | Where-Object { $_.name -eq $DefinedName } | Select-Object Name, Text, Comment, LocalSheetID
            }
            Else
            {
                $DefinedNameExists = $false
                Write-Warning ("Get-SLDefinedName :`tThe Defined Name '{0}' could not be found. Check spelling and try again." -f $DefinedName)
            }

        }

        if ($PSCmdlet.ParameterSetName -eq 'All')
        {
            $DefinedNames | Select-Object Name, Text, Comment, LocalSheetID
        }

        #$WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force  -PassThru

    }#process
    END
    {
    }
}


Function Remove-SLDefinedName
{

    <#

.SYNOPSIS
    Remove defined names contained in an excel document.

.DESCRIPTION
    Remove defined names contained in an excel document.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER DefinedName
    The defined name that has to be removed.

.PARAMETER RemoveAll
    Will remove all defined names from a workbook.Use with caution!

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Remove-SLDefinedName-DefinedName dfname2 -Verbose  | Save-SLDocument


    Description
    -----------
    Remove the definedname 'dfname2' from 'myfirstdoc'.


    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx | Remove-SLDefinedName -RemoveAll -Verbose | Save-SLDocument


    Description
    -----------
    Remove all defined names in workbook 'myfirstdoc'.


.INPUTS
   String,SpreadsheetLight.SLDocument

.OUTPUTS
   String

.Link
    N/A

#>

    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true, Position = 0, ValueFromPipeLine = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [parameter(Mandatory = $true, Position = 1, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'DefinedName')]
        [String]$DefinedName,

        [parameter(Mandatory = $true, Position = 1, ParameterSetName = 'RemoveAll')]
        [Switch]$RemoveAll

    )
    PROCESS
    {
        $DefinedNames = $WorkBookInstance.GetDefinedNames()

        if ($PSCmdlet.ParameterSetName -eq 'DefinedName')
        {
            $DefinedNameMatches = $DefinedNames | Where-Object { $_.name -eq $DefinedName } | Select-Object -ExpandProperty Name

            If ($DefinedNameMatches)
            {
                Write-Verbose ("Remove-SLDefinedName :`tRemoving Defined Name '{0}'.." -f $DefinedName)
                $WorkBookInstance.DeleteDefinedName($DefinedName) | Out-Null
                $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
            }
            Else
            {
                Write-Warning ("Remove-SLDefinedName :`tDefined Name '{0}' could not be found. Check spelling and try again." -f $DefinedName)
            }
        }

        if ($PSCmdlet.ParameterSetName -eq 'RemoveAll')
        {
            Write-Verbose ("Remove-SLDefinedName :`tRemoving all DefinedNames from the workbook..")
            $DefinedNames |
                ForEach-Object {
                    Write-Verbose ("Remove-SLDefinedName :`tRemoving DefinedName '{0}'" -f $_.Name)
                    $WorkBookInstance.DeleteDefinedName($_.name) | Out-Null
                }
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }

    }#process
    END
    {
    }
}


Function Set-SLVlookup
{

    <#

.SYNOPSIS
    Perform vlookup.Supports lookup from same or on different worksheets.

.DESCRIPTION
    Perform vlookup.Supports lookup from same or on different worksheets. The lookup worksheet(s) have to be from the same workbook.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER VRangeworksheetname
    This is the name of the worksheet that contains the lookup table.

.PARAMETER VRange
    The range that contains the lookup values eg: A1:D20.

.PARAMETER Vcellworksheetname
    This is the worksheetname that will host the vlookup formula.

.PARAMETER Vlookupcell
    This is the lookup cell reference. Example - "C10".

.PARAMETER VFormulacellRange
    This is the range containing the lookup formula. Example = "D10:D20". Note range must include cells from the same column

.PARAMETER DataColumn
    This is the datacolumn from the lookup table that contains the value(s) to be pulled.
    So if the lookup range is D1:G6 the datatable is 4 columns wide so count from 1(D) to G(4).

.Example
    PS C:\> $doc = Get-SLDocument -Path D:\ps\Excel\Vlookup.xlsx
    PS C:\> $doc | Set-SLVlookup -VRangeworksheetname OS -VRange E5:G7 -Vcellworksheetname disk -Vlookupcell A6 -VFormulacellRange H6:H11 -DataColumn 2 -Verbose
    PS C:\> $doc | Set-SLVlookup -VRangeworksheetname OS -VRange E5:G7 -Vcellworksheetname disk -Vlookupcell A6 -VFormulacellRange I6:I11 -DataColumn 3 -Verbose
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    Use vlookup to lookup a datatable(E5:G7) contained in worksheet 'disk' and dump the matching values into worksheet 'OS'.
    Note: since we are populating 2 columns H & I we need to use the vlookup cmdlet twice.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\Vlookup.xlsx  |
                Set-SLVlookup -VRangeworksheetname disk -VRange L6:M8 -Vcellworksheetname disk -Vlookupcell A6 -VFormulacellRange J6:J11 -DataColumn 2 -Verbose |
                    Save-SLDocument


    Description
    -----------
    Use vlookup to lookup and insert values in the worksheet 'disk'.


.INPUTS
   String,SpreadsheetLight.SLDocument

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    N/A

#>


    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true, Position = 1, Valuefrompipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, HelpMessage = 'This is the name of the worksheet that contains the lookup table')]
        [string]$VRangeworksheetname,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+:[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Set-SLVlookup :`tVRange should specify values in following format. Eg: A1:D10 or AB1:AD5"; break }
            })]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, HelpMessage = 'This is the lookup range. Example.. a1:c50')]
        [string]$VRange,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, HelpMessage = 'This is the worksheetname that will host the vlookup formula')]
        [string]$Vcellworksheetname,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Set-SLVlookup :`tVlookupcell should specify values in following format. Eg: A1,B10,AB5..etc"; break }
            })]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, HelpMessage = 'This is the lookup cell reference. Example - "C10"')]
        [string]$Vlookupcell,

        [ValidateScript({
                $r1, $r2 = $_ -split ':'
                $r1_match = [regex]::Match($r1, '[a-zA-Z]+') | Select-Object -ExpandProperty value
                $r2_match = [regex]::Match($r2, '[a-zA-Z]+') | Select-Object -ExpandProperty value
                if ($r1_match -eq $r2_match) { $true }
                else { $false; Write-Warning "Set-SLVlookup :`tVFormulacellRange should specify values that belong to the same column. Eg: A1:A10 or AB1:AB5"; break }
            })]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, HelpMessage = 'This is the Cell Range containing the lookup formula. Example = "D10:D20"')]
        [string]$VFormulacellRange,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, HelpMessage = 'This is the datacolumn from the lookup table that contains the value(s) to be pulled')]
        [int]$DataColumn


    )

    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $VRangeworksheetname -NoPassThru)
        {

            if ($VRangeworksheetname -ne $Vcellworksheetname)
            {
                $WorkBookInstance.SelectWorksheet($Vcellworksheetname) | Out-Null
                $nrange = Convert-ToExcelAbsoluteRange -Range $VRange -WorkSheetName $VRangeworksheetname
            }
            else
            {
                $nrange = Convert-ToExcelAbsoluteRange -Range $VRange
            }

            $r1, $r2 = $VFormulacellRange -split ':'
            $start = Convert-ToExcelRowColumnIndex -CellReference $r1 | Select-Object -ExpandProperty Row
            $END = Convert-ToExcelRowColumnIndex -CellReference $r2 | Select-Object -ExpandProperty Row
            $columnname = Convert-ToExcelColumnName -CellReference $r1
            $lookup = Convert-ToExcelColumnName -CellReference $Vlookupcell

            for ($i = $start; $i -le $END; $i++)
            {
                $cref = "$columnname$i"
                $lookup1 = "$lookup$i"

                Write-Verbose ("Set-SLVlookup :`tLookup cell '{0}', Lookup Range '{1}',Datacolumn '{2}'" -f $lookup1, $nrange, $datacolumn)
                $WorkBookInstance.SetCellValue($cref, "=vlookup($lookup1,$nrange,$datacolumn,$false)") | Out-Null
            }
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }
    }
    END
    {
    }
}


Function Set-SLDataValidation
{

    <#

.SYNOPSIS
    Add Datavalidation.

.DESCRIPTION
    Create drop-down lists or otherwise control the type of data that users enter on a worksheet.
    Apply data constraints on Integers,decimals,Date,Time,TextLength or custom forumulas.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER ValidationTarget
    The target cell or range of cells that need to have datavalidation.

.PARAMETER DataLookupRange
    The range that can be used to create a drop-down list on a cell or range of cells.

.PARAMETER DefinedName
    The definedname that can be used to create a drop-down list on a cell or range of cells.

.PARAMETER Decimal
    Restrict data entry to a decimal number  Example - "1.3".

.PARAMETER StartDecimal
    This is the minimum value for a decimal range.

.PARAMETER EndDecimal
    This is the maximum value for a decimal range.


.PARAMETER WholeNumber
    Restrict data entry to a wholenumber  Example - "3".

.PARAMETER StartWholeNumber
    This is the minimum value for a wholenumber range.

.PARAMETER EndWholeNumber
    This is the maximum value for a wholenumber range.

.PARAMETER Date
    Restrict data entry to a Date  Example - "12/25/2014".

.PARAMETER StartDate
    This is the minimum value for a Date range.

.PARAMETER EndDate
    This is the maximum value for a Date range.

.PARAMETER Time
    Restrict data entry to a Time  Example - "14:30:55".

.PARAMETER StartTime
    This is the minimum value for a Time range.

.PARAMETER EndTime
    This is the maximum value for a Time range.

.PARAMETER TextLength
    Restrict data entry to a TextLength  Example - "6".

.PARAMETER StartTextLength
    This is the minimum value for a TextLength range.

.PARAMETER EndTextLength
    This is the maximum value for a TextLength range.

.PARAMETER CustomFormula
    Restrict data entry to values that conform to a CustomFormula  Example - "=len(b3)".

.PARAMETER ValidationOperator
    The Operator to be used for validating data.
    Use tab or intellisense to select from a list of possible values.
    'Equal','NotEqual','GreaterThan','LessThan','GreaterThanOrEqual','LessThanOrEqual','Between'

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\DataValidation.xlsx
    PS C:\> $doc | Set-SLColumnValue -WorksheetName sheet1 -CellReference B3 -value @('Pete','Andre','Roger','Jimmy','Pat') -Verbose
    PS C:\> $doc | Set-SLDataValidation -WorksheetName sheet1 -ValidationTarget C3 -DataLookupRange B3:B7 -Verbose
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    Set column vlaues B3 to B7 and use that range to create a drop-down list in cell C3.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\DataValidation.xlsx
    PS C:\> $doc | Set-SLDataValidation -WorksheetName sheet1 -ValidationTarget D3:E4 -DataLookupRange B3:B7 -Verbose
    PS C:\> $doc | Save-SLDocument


    Description
    -----------
    Use a predefined range B3:B7(values we created in example 1 above) to create a drop-down list in a range of cells D3:E4.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\DataValidation.xlsx
    PS C:\> $doc | Set-SLColumnValue    -WorksheetName sheet2 -CellReference C3 -value @('Sampras','Agassi','Federer','Connors','Rafter') -Verbose
    PS C:\> $doc | New-SLDefinedName    -WorksheetName sheet2 -DefinedName LookupRange1 -Range C3:C7 -Verbose
    PS C:\> $doc | Set-SLDataValidation -WorksheetName sheet1 -ValidationTarget F3 -DefinedName LookupRange1 -Verbose
    PS C:\> $doc | Save-SLDocument


    Description
    -----------
    Set column vlaues C3 to C7 and use that range to create a DefinedName called 'lookuprange1' and finally use the DefinedName to create a drop-down list in cell F3.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\DataValidation.xlsx
    PS C:\> $doc | Set-SLDataValidation -WorksheetName sheet1 -ValidationTarget c4 -StartDecimal 1.2 -ENDDecimal 2.5 -ValidationOperator Between -Verbose
    PS C:\> $doc | Save-SLDocument


    Description
    -----------
    Set datavalidation on cell c4 to contain only values that are between 1.2 and 2.5.
    Note: if you omit the ValidationOperator in the command above the validation operator defaults to 'notbetween'
    so in effect the validation would then be all values that are not between 1.2 and 2.5

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\DataValidation.xlsx
    PS C:\> $doc | Set-SLDataValidation -WorksheetName sheet1 -ValidationTarget c5 -WholeNumber 5 -ValidationOperator Equal -Verbose
    PS C:\> $doc | Save-SLDocument


    Description
    -----------
    Set datavalidation on cell c5 to contain only value that is equal to 5.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\DataValidation.xlsx
    PS C:\> $doc | Set-SLColumnValue -WorksheetName sheet2 -CellReference D3 -value @(1,2,3)
    PS C:\> $doc | Set-SLColumnValue -WorksheetName sheet2 -CellReference E3 -value @(3,4,5)
    PS C:\> $doc | New-SLDefinedName    -WorksheetName sheet2 -DefinedName MinRangeValue -Range D3:D5
    PS C:\> $doc | New-SLDefinedName    -WorksheetName sheet2 -DefinedName MaxRangeValue -Range E3:E5
    PS C:\> $doc | Set-SLDataValidation -WorksheetName sheet1 -ValidationTarget C13 -StartWholeNumber '=MIN(MinRangeValue)' -EndWholeNumber '=MAX(MaxRangeValue)' -ValidationOperator Between
    PS C:\> $doc | Save-SLDocument


    Description
    -----------
    Set column values D3:D5 and also E3:E5.
    Create defined names for each of the above ranges
    Set datavalidation on C13 that makes use of the 2 defined names created above.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\DataValidation.xlsx
    PS C:\> $doc | Set-SLDataValidation -WorksheetName sheet1 -ValidationTarget C14 -StartDate '12/20/2014' -EndDate '12/25/2014' -ValidationOperator Between -Verbose
    PS C:\> $doc | Save-SLDocument


    Description
    -----------
    Restrict cell c14 to contain dates between '12/20/2014' & '12/25/2014'

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\DataValidation.xlsx
    PS C:\> $doc | Set-SLDataValidation -WorksheetName sheet1 -ValidationTarget C17 -Date '12/25/2014' -ValidationOperator Equal -Verbose
    PS C:\> $doc | Save-SLDocument


    Description
    -----------
    Restrict cell c17 to contain date equal to '12/25/2014'.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\DataValidation.xlsx
    PS C:\> $doc | Set-SLDataValidation -WorksheetName sheet1 -ValidationTarget C20 -Time 14:20:35 -ValidationOperator LessThan
    PS C:\> $doc | Save-SLDocument


    Description
    -----------
    Restrict cell c20 to contain time values that are lessthan 14:20:35.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\DataValidation.xlsx
    PS C:\> $doc | Set-SLDataValidation -WorksheetName sheet1 -ValidationTarget C23 -TextLength '=SUM(LEN(B3),LEN(F3))' -ValidationOperator LessThanOrEqual -Verbose
    PS C:\> $doc | Save-SLDocument


    Description
    -----------
    The forumala SUM(LEN(B3),LEN(F3)) --> compute the length of the cell value B3, compute the length of cell value F3 and add them up.
    Restrict cell c23 to contain time values that are lessthanoeEqual to the textlength obtained by the formula above.
    If the forumal SUM(LEN(B3),LEN(F3)) yeilded value 11 then the total length of the value in cell C23 cannot exceed 11.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\DataValidation.xlsx
    PS C:\> $doc | Set-SLDataValidation -WorksheetName sheet1 -ValidationTarget D14:D17 -CustomFormula 'COUNTIF($D$14:$D$17,D14) <= 1' -Verbose
    PS C:\> $doc | Save-SLDocument


    Description
    -----------
    The forumala 'COUNTIF($D$14:$D$17,D14) <= 1' --> count the occurrences of the value in cell D14, in the range $D$14:$D$17. The formula's result must be 1 or 0
    The net result is to prevent duplicate values from being entered in the range D14:D17.


.INPUTS
   String,Int,Double,SpreadsheetLight.SLDocument

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    http://www.contextures.com/xlDataVal07.html

#>



    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true, Position = 0, ValueFromPipeLine = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [Alias('CurrentWorkSheetName')]
        [parameter(Mandatory = $true, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true)]
        [string]$ValidationTarget,

        [Alias('Range')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'DataLookupRange')]
        [String]$DataLookupRange,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'NamedRange')]
        [String]$DefinedName,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Decimal')]
        [Double]$Decimal,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'StartEndDecimal')]
        [Double]$StartDecimal,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'StartEndDecimal')]
        [Double]$EndDecimal,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'WholeNumber')]
        [Int]$WholeNumber,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'StartEndWholeNumber')]
        $StartWholeNumber,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'StartEndWholeNumber')]
        $EndWholeNumber,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Date')]
        [String]$Date,

        [parameter(Mandatory = $True, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'StartEndDate')]
        [String]$StartDate,

        [parameter(Mandatory = $True, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'StartEndDate')]
        [String]$EndDate,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Time')]
        [String]$Time,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'StartEndTime')]
        [String]$StartTime,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'StartEndTime')]
        [String]$EndTime,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'TextLength')]
        $TextLength,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'StartEndTextLength')]
        $StartTextLength,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'StartEndTextLength')]
        $EndTextLength,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Custom')]
        [string]$CustomFormula,

        [ValidateSet('Equal', 'NotEqual', 'GreaterThan', 'LessThan', 'GreaterThanOrEqual', 'LessThanOrEqual', 'Between')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'WholeNumber')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'StartEndWholeNumber')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Date')]
        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'StartEndDecimal')]
        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'StartEndDate')]
        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'StartEndTime')]
        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'StartEndTextLength')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Time')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'TextLength')]
        [String]$ValidationOperator


    )
    PROCESS
    {

        ## -- ## Check if the referenced worksheet exists in the workbook and proceed only if true.
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {

            Switch -Regex ($ValidationTarget)
            {

                #CellReference
                '^[a-zA-Z]+\d+$'
                {
                    Write-Verbose ("Set-SLDataValidation :`tValidationTarget is Cell '{0}'" -f $ValidationTarget)
                    $DataValidation = $WorkBookInstance.CreateDataValidation($ValidationTarget)
                    $isValidationTargetValid = $true
                    Break
                }

                #Range
                '[a-zA-Z]+\d+:[a-zA-Z]+\d+$'
                {
                    $startcellreference, $endcellreference = $ValidationTarget -split ':'
                    Write-Verbose ("Set-SLDataValidation :`tValidationTarget is Cell Range '{0}'" -f $ValidationTarget)
                    $DataValidation = $WorkBookInstance.CreateDataValidation($startcellreference, $endcellreference)
                    $isValidationTargetValid = $true
                    Break
                }

                Default
                {
                    Write-Warning ("Set-SLDataValidation :`tYou must provide either a Cellreference Eg. C3 or a Range Eg. C3:G10")
                    $isValidationTargetValid = $false
                    Break
                }

            }#switch

            if ( ($PSCmdlet.ParameterSetName -eq 'DataLookupRange') -and $isValidationTargetValid)
            {
                Write-Verbose ("Set-SLDataValidation :`tDatalookup Range '{0}'" -f $DataLookupRange)
                $AbsoluteRange = Convert-ToExcelAbsoluteRange -Range $DataLookupRange

                #allowlist - DataSource, IgnoreBlank, InCellDropDown
                $DataValidation.AllowList($AbsoluteRange, $true, $true)
                $DataValidation.SetInputMessage('ValidationMessage', "Only Values in the Cell Range - $DataLookupRange are accepted") | Out-Null
                $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You Must enter a Value specified in the range - $range") | Out-Null

            }

            if ( ($PSCmdlet.ParameterSetName -eq 'NamedRange') -and $isValidationTargetValid)
            {

                if ((($WorkBookInstance.GetDefinedNames() | Select-Object -ExpandProperty Name) -contains $DefinedName))
                {
                    $DefinedNameText = $WorkBookInstance.GetDefinedNameText($DefinedName)
                    Write-Verbose ("Set-SLDataValidation :`tNamedRange '{0}' corresponds to '{1}'" -f $DefinedName, $DefinedNameText)

                    $DataValidation.AllowList("=$DefinedName", $true, $true)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values in the Cell Range - $DefinedNameText are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You Must enter a Value specified in the range - $range") | Out-Null
                }
                Else
                {
                    Write-Warning ("Set-SLDataValidation :`tSpecified Named Range '{0}' was not found in the workbook. Check spelling and try again." -f $DefinedName)
                    break
                }
            }

            if ( ($PSCmdlet.ParameterSetName -eq 'Decimal') -and $isValidationTargetValid)
            {

                If ($ValidationOperator -ne 'Between')
                {
                    Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should be '{1}' Date '{2}'" -f $ValidationTarget, $ValidationOperator, $Decimal)
                    $DataValidation.AllowDecimal([SpreadsheetLight.SLDataValidationSingleOperandValues]::$ValidationOperator, $Decimal, $true)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values that are $ValidationOperator $Decimal are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You Must enter a Value that is $ValidationOperator  Decimal $Decimal") | Out-Null
                }
                else
                {
                    Write-Warning ("Set-SLDataValidation :`tUse ValidationOperator 'Between' with Parameters 'StartDecimal' & EndDecimal' ")
                    Break
                }
            }

            if ($PSCmdlet.ParameterSetName -eq 'StartEndDecimal' -and $isValidationTargetValid)
            {
                If ($ValidationOperator -eq 'Between')
                {
                    Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should be between '{1}' & '{2}'" -f $ValidationTarget, $StartDecimal, $ENDDecimal)
                    $DataValidation.AllowDecimal($true, $StartDecimal, $ENDDecimal, $false)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values in the Decimal Range - $StartDecimal - $ENDDecimal are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You Must enter a Value specified in the range - $StartDecimal - $ENDDecimal") | Out-Null
                }
                else
                {
                    Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should NOT be between '{1}' & '{2}'" -f $ValidationTarget, $StartDecimal, $ENDDecimal)
                    $DataValidation.AllowDecimal($true, $StartDecimal, $ENDDecimal, $false)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values NOT in the Decimal Range - $StartDecimal - $ENDDecimal are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You Must enter a Value NOT in the range - $StartDecimal - $ENDDecimal") | Out-Null
                }
            }

            if ($PSCmdlet.ParameterSetName -eq 'WholeNumber' -and $isValidationTargetValid)
            {
                If ($ValidationOperator -ne 'Between')
                {
                    Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should be '{1}' to '{2}'" -f $ValidationTarget, $ValidationOperator, $WholeNumber)
                    $DataValidation.AllowWholeNumber([SpreadsheetLight.SLDataValidationSingleOperandValues]::$ValidationOperator, $WholeNumber, $true)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values $ValidationOperator - $WholeNumber are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "Validation Criteria: Value $ValidationOperator - $WholeNumber not met") | Out-Null
                }
                else
                {
                    Write-Warning ("Set-SLDataValidation :`tWhen Parameter 'Wholenumber' is used, the value of the Validationoperator must NOT be 'Between'. Use 'Between' with StartWholeNumber & EndWholeNumber ")
                    Break
                }
            }

            if ($PSCmdlet.ParameterSetName -eq 'StartEndWholeNumber' -and $isValidationTargetValid)
            {
                If ($ValidationOperator -eq 'Between')
                {
                    Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should be between '{1}' and '{2}'" -f $ValidationTarget, $StartWholeNumber, $EndWholeNumber)
                    $DataValidation.AllowWholeNumber($true, $StartWholeNumber, $EndWholeNumber, $true)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values $ValidationOperator : $StartWholeNumber-$EndWholeNumber are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You must enter a value that is between $StartWholeNumber-$EndWholeNumber") | Out-Null
                }
                else
                {
                    Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should NOT be between '{1}' and '{2}'" -f $ValidationTarget, $StartWholeNumber, $EndWholeNumber)
                    $DataValidation.AllowWholeNumber($false, $StartWholeNumber, $EndWholeNumber, $true)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values NOT between: $StartWholeNumber-$EndWholeNumber are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You must enter a value that is NOT between $StartWholeNumber-$EndWholeNumber") | Out-Null
                }
            }

            if ($PSCmdlet.ParameterSetName -eq 'Date' -and $isValidationTargetValid)
            {
                If ($ValidationOperator -ne 'Between')
                {
                    Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should be '{1}' Date '{2}'" -f $ValidationTarget, $ValidationOperator, $Date)
                    $DataValidation.AllowDate([SpreadsheetLight.SLDataValidationSingleOperandValues]::$ValidationOperator, $Date, $true)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values that are $ValidationOperator $date are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You Must enter a Value that is $ValidationOperator  Date $Date") | Out-Null
                }
                else
                {
                    Write-Warning ("Set-SLDataValidation :`tUse ValidationOperator 'Between' with Parameters 'StartDate' & EndDate' ")
                    Break
                }
            }

            if ($PSCmdlet.ParameterSetName -eq 'StartEndDate' -and $isValidationTargetValid)
            {
                If ($ValidationOperator -eq 'Between')
                {
                    Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should be between '{1}' & '{2}'" -f $ValidationTarget, $StartDate, $EndDate)
                    $DataValidation.AllowDate($true, $StartDate, $EndDate, $true)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values between Dates $StartDate & $EndDate are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You Must enter a Date that is between $StartDate & $EndDate") | Out-Null
                }
                Else
                {
                    Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should NOT be between '{1}' & '{2}'" -f $ValidationTarget, $StartDate, $EndDate)
                    $DataValidation.AllowDate($false, $StartDate, $EndDate, $true)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values NOT between Dates $StartDate & $EndDate are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You Must enter a Date that is NOT between $StartDate & $EndDate") | Out-Null

                }
            }

            if ($PSCmdlet.ParameterSetName -eq 'Time' -and $isValidationTargetValid)
            {

                If ($ValidationOperator -ne 'Between')
                {
                    Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should be '{1}' Time '{2}'" -f $ValidationTarget, $ValidationOperator, $Time)
                    $DataValidation.AllowTime([SpreadsheetLight.SLDataValidationSingleOperandValues]::$ValidationOperator, $Time, $true)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values that are $ValidationOperator $Time are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You Must enter a Value that is $ValidationOperator  Time $Time") | Out-Null
                }
                else
                {
                    Write-Warning ("Set-SLDataValidation :`tUse ValidationOperator 'Between' with Parameters 'StartTime' & EndTime' ")
                    Break
                }
            }

            if ($PSCmdlet.ParameterSetName -eq 'StartEndTime' -and $isValidationTargetValid)
            {
                If ($ValidationOperator -eq 'Between')
                {
                    Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should be between Time values '{1}' & '{2}'" -f $ValidationTarget, $StartTime, $EndTime)
                    $DataValidation.AllowTime($true, $StartTime, $EndTime, $true)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values between Times $StartTime & $EndTime are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You Must enter a Time that is between $StartTime & $EndTime") | Out-Null
                }
                Else
                {
                    Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should NOT be between Time values '{1}' & '{2}'" -f $ValidationTarget, $StartTime, $EndTime)
                    $DataValidation.AllowTime($false, $StartTime, $EndTime, $true)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values NOT between Times $StartTime & $EndTime are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You Must enter a Time that is NOT between $StartTime & $EndTime") | Out-Null

                }
            }



            if ($PSCmdlet.ParameterSetName -eq 'TextLength' -and $isValidationTargetValid)
            {

                If ($ValidationOperator -ne 'Between')
                {
                    Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should be '{1}' TextLength '{2}'" -f $ValidationTarget, $ValidationOperator, $Time)
                    $DataValidation.AllowTextLength([SpreadsheetLight.SLDataValidationSingleOperandValues]::$ValidationOperator, $TextLength, $true)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values that are $ValidationOperator $TextLength are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You Must enter a Value that is $ValidationOperator  TextLength $TextLength") | Out-Null
                }
                else
                {
                    Write-Warning ("Set-SLDataValidation :`tUse ValidationOperator 'Between' with Parameters 'StartTextLength' & EndTextLength' ")
                    Break
                }
            }

            if ($PSCmdlet.ParameterSetName -eq 'StartEndTextLength' -and $isValidationTargetValid)
            {
                If ($ValidationOperator -eq 'Between')
                {
                    Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should be between TextLength values '{1}' & '{2}'" -f $ValidationTarget, $StartTextLength, $EndTextLength)
                    $DataValidation.AllowTextLength($true, $StartTextLength, $EndTextLength, $true)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values between TextLengths $StartTextLength & $EndTextLength are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You Must enter a TextLength that is between $StartTextLength & $EndTextLength") | Out-Null
                }
                Else
                {
                    Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should NOT be between Time values '{1}' & '{2}'" -f $ValidationTarget, $StartTextLength, $EndTextLength)
                    $DataValidation.AllowTextLength($false, $StartTextLength, $EndTextLength, $true)
                    $DataValidation.SetInputMessage('ValidationMessage', "Only Values NOT between TextLengths $StartTextLength & $EndTextLength are accepted") | Out-Null
                    $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You Must enter a TextLength that is NOT between $StartTextLength & $EndTextLength") | Out-Null

                }
            }

            if ($PSCmdlet.ParameterSetName -eq 'Custom' -and $isValidationTargetValid)
            {
                Write-Verbose ("Set-SLDataValidation :`tValues in cell '{0}' should conform to formula '{1}'" -f $ValidationTarget, $CustomFormula)
                $DataValidation.AllowCustom($CustomFormula, $true)
                $DataValidation.SetInputMessage('ValidationMessage', "Only Values that conform to the forumula  - $CustomFormula are accepted") | Out-Null
                $DataValidation.SetErrorAlert([DocumentFormat.OpenXml.Spreadsheet.DataValidationErrorStyleValues]::'Stop', 'Data Input Error', "You Must enter a Value that is valid for forumula - $CustomFormula") | Out-Null
            }

            if ($isValidationTargetValid)
            {
                Write-Verbose ("Set-SLDataValidation :`tAdding Datavalidation..")
                $WorkBookInstance.AddDataValidation($DataValidation) | Out-Null
                $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
            }


        }#END if select-slworksheet

    }#process
    END
    {
    }
}


Function Remove-SLDataValidation
{

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

Function Set-SLConditionalFormattingDataBars
{

    <#

.SYNOPSIS
    Set conditional formatting data bars on a given range of cells.

.DESCRIPTION
    Set conditional formatting data bars on a given range of cells.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
        This is the name of the worksheet that contains the cell range where formatting is to be applied.

.PARAMETER Range
    The range of cells where conditional formatting has to be applied.

.PARAMETER DataBarColor
    to be used with the parameterset 'normal'.
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'Blue','Green','Red','Orange','LightBlue','Purple'

.PARAMETER ThemeColor
    to be used with the parameterset 'CustomDataBar1'.
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'Light1Color','Dark1Color','Light2Color','Dark2Color','Accent1Color','Accent2Color','Accent3Color','Accent4Color','Accent5Color','Accent6Color','Hyperlink','FollowedHyperlinkColor'

.PARAMETER DataBarMinLength
    Set the minimum length of the databar.

.PARAMETER DataBarMaxLength
    Set the maximum length of the databar.

.PARAMETER DataBarType1
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'Value','Number','Percent','Formula','Percentile'

.PARAMETER DataBarType2
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'Value','Number','Percent','Formula','Percentile'

.PARAMETER MinValue
    This is the minimum value from which the databar will begin.

.PARAMETER MaxValue
    This is the maximum value at which the databar will end.

.PARAMETER Color
    Color of the databar.Can be used in place of themecolor.

.PARAMETER ShowDataBarOnly
    If used only the databar sans value will be shown.



.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\Databars.xlsx
    PS C:\> $doc | Set-SLConditionalFormattingDataBars -WorksheetName sheet1 -Range e4:h6 -DataBarColor Green -Verbose
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    apply conditional formatting on range e4:h6 with the databar color chosen as green.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\Databars.xlsx
    PS C:\> $doc | Set-SLConditionalFormattingDataBars -WorksheetName sheet1 -Range e8:e10 -DataBarMinLength 0 -DataBarMaxLength 100 -DataBarType1 Number -MinValue 0 -DataBarType2 Value -MaxValue 100 -ThemeColor Accent3Color
    PS C:\> $doc | Save-SLDocument


    Description
    -----------
    Custom databar formatting applied with accent color3 as the databar color.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\Databars.xlsx
    PS C:\> $doc | Set-SLConditionalFormattingDataBars -WorksheetName sheet7 -Range f8:f10 -DataBarMinLength 0 -DataBarMaxLength 80 -DataBarType1 Number -MinValue 0 -DataBarType2 Value -MaxValue 100 -ThemeColor Accent4Color -ShowDataBarOnly
    PS C:\> $doc | Save-SLDocument


    Description
    -----------
    Same as the previous example but here we make 2 changes.
    1 - the maximum databar length is changed from 100 to 80 and
    2 - The values are hidden showing just the databars.

.Example
    PS C:\> Get-SLDocument D:\ps\excel\Databars.xlsx  |
                Set-SLFill -WorksheetName sheet7 -TargetCellorRange h12:h14 -Color Black |
                    Set-SLFont -FontColor White |
            Set-SLConditionalFormattingDataBars -DataBarMinLength 0 -DataBarMaxLength 80 -DataBarType1 Number -MinValue 0 -DataBarType2 Value -MaxValue 100 -ThemeColor Accent4Color |
                Set-SLColumnWidth -ColumnName h -ColumnWidth 20 |
                    Save-SLDocument


    Description
    -----------
    At times it may be difficult to see where the bars end, because of the graduated coloring in the data bars,
    so here we apply a dark fill color to the cells, and then change the font to a light color
    Also we change the width of the column to 20 which makes it a little easier to see the differences in the databar lengths.
    Note: since we are piping data between cmdlets we can ignore specifying the values for some of the parameters such as 'worksheetname' and 'Range'.
    However the best practise would be to specify parameter names so that there is no cause for confusion or ambiguity.



.INPUTS
   String,SpreadsheetLight.SLDocument

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    http://www.excel-easy.com/examples/data-bars.html

#>




    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true, Position = 0, ValueFromPipeLine = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [Alias('CurrentWorkSheetName')]
        [parameter(Mandatory = $true, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,


        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+:[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Set-SLConditionalFormattingDataBars :`tRange should specify values in following format. Eg: A1:D10 or AB1:AD5"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true)]
        [string]$Range,

        [ValidateSet('Blue', 'Green', 'Red', 'Orange', 'LightBlue', 'Purple')]
        [parameter(Mandatory = $True, Position = 3, ParameterSetName = 'Normal')]
        [string]$DataBarColor,

        [ValidateSet('Light1Color', 'Dark1Color', 'Light2Color', 'Dark2Color', 'Accent1Color', 'Accent2Color', 'Accent3Color', 'Accent4Color', 'Accent5Color', 'Accent6Color', 'Hyperlink', 'FollowedHyperlinkColor')]
        [parameter(Mandatory = $True, ParameterSetName = 'CustomDataBar1')]
        [string]$ThemeColor,

        [parameter(Mandatory = $True, ParameterSetName = 'CustomDataBar1')]
        [parameter(Mandatory = $True, ParameterSetName = 'CustomDataBar2')]
        [int]$DataBarMinLength,

        [parameter(Mandatory = $True, ParameterSetName = 'CustomDataBar1')]
        [parameter(Mandatory = $True, ParameterSetName = 'CustomDataBar2')]
        [int]$DataBarMaxLength,

        [ValidateSet('Value', 'Number', 'Percent', 'Formula', 'Percentile')]
        [parameter(Mandatory = $True, ParameterSetName = 'CustomDataBar1')]
        [parameter(Mandatory = $True, ParameterSetName = 'CustomDataBar2')]
        [string]$DataBarType1,

        [ValidateSet('Value', 'Number', 'Percent', 'Formula', 'Percentile')]
        [parameter(Mandatory = $True, ParameterSetName = 'CustomDataBar1')]
        [parameter(Mandatory = $True, ParameterSetName = 'CustomDataBar2')]
        [string]$DataBarType2,

        [parameter(Mandatory = $True, ParameterSetName = 'CustomDataBar1')]
        [parameter(Mandatory = $True, ParameterSetName = 'CustomDataBar2')]
        $MinValue,

        [parameter(Mandatory = $True, ParameterSetName = 'CustomDataBar1')]
        [parameter(Mandatory = $True, ParameterSetName = 'CustomDataBar2')]
        $MaxValue,

        [parameter(Mandatory = $True, ParameterSetName = 'CustomDataBar2')]
        [string]$Color,

        [parameter(Mandatory = $False, ParameterSetName = 'CustomDataBar1')]
        [parameter(Mandatory = $False, ParameterSetName = 'CustomDataBar2')]
        [Switch]$ShowDataBarOnly = $false
    )
    PROCESS
    {

        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {

            $startcellreference, $ENDcellreference = $range -split ':'
            $ConditionalFormatting = New-Object SpreadsheetLight.SLConditionalFormatting($startcellreference, $ENDcellreference)

            if ($PSCmdlet.ParameterSetName -eq 'Normal')
            {
                Write-Verbose ("Set-SLConditionalFormattingDataBars :`t Databar color is '{0}'" -f $DataBarColor)
                $ConditionalFormatting.SetDataBar([SpreadsheetLight.SLConditionalFormatDataBarValues]::$DataBarColor) | Out-Null
            }
            if ($PSCmdlet.ParameterSetName -eq 'CustomDataBar1')
            {
                #Write-Verbose ("Set-SLConditionalFormattingDataBars :`tData Range '{0}'. DataBarMinLength is '{1}'" -f $Range,$DataBarColor)
                $ConditionalFormatting.SetCustomDataBar($ShowDataBarOnly, $DataBarMinLength, $DataBarMaxLength, [SpreadsheetLight.SLConditionalFormatMinMaxValues]::$DataBarType1, $MinValue, [SpreadsheetLight.SLConditionalFormatMinMaxValues]::$DataBarType2, $MaxValue, [SpreadsheetLight.SLThemeColorIndexValues]::$ThemeColor) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'CustomDataBar2')
            {
                $ConditionalFormatting.SetCustomDataBar($ShowDataBarOnly, $DataBarMinLength, $DataBarMaxLength, [SpreadsheetLight.SLConditionalFormatMinMaxValues]::$DataBarType1, $MinValue, [SpreadsheetLight.SLConditionalFormatMinMaxValues]::$DataBarType2, $MaxValue, [System.Drawing.Color]::$Color) | Out-Null
            }


            Write-Verbose ("Set-SLConditionalFormattingDataBars :`tSetting conditional formatting on range '{0}'" -f $Range)
            $WorkBookInstance.AddConditionalFormatting($ConditionalFormatting) | Out-Null

            $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }#select-slworksheet

    }#process
    END
    {


    }
}


Function Set-SLConditionalFormatColorScale
{

    <#

.SYNOPSIS
    Apply conditional formatting color scale to a range.

.DESCRIPTION
    Apply conditional formatting color scale to a range.
    Cells are shaded with gradations of two or three colors that correspond to minimum, midpoint, and maximum thresholds.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    This is the name of the worksheet that contains the cell range where formatting is to be applied.

.PARAMETER Range
    The range of cells where conditional formatting has to be applied.

.PARAMETER ColorScaleType
    Built-in color scale styles.
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'GreenYellowRed','RedYellowGreen','BlueYellowRed','RedYellowBlue','GreenWhiteRed','RedWhiteGreen','BlueWhiteRed','RedWhiteBlue','WhiteRed','RedWhite','GreenWhite','WhiteGreen','Yellow',
    'Red','RedYellow','GreenYellow','YellowGreen'

.PARAMETER ColorScaleMinType
    to be used with a custom color scale formatting style.
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
   'Value','Number','Percent','Formula','Percentile'

.PARAMETER MinValue
    the minimum value in the range.

.PARAMETER ColorScaleMinSystemColor
    Custom color for the minimum values.

.PARAMETER ColorScaleMaxType
    to be used with a custom color scale formatting style.
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
   'Value','Number','Percent','Formula','Percentile'

.PARAMETER MaxValue
    The maximum value in the range.

.PARAMETER ColorScaleMaxSystemColor
    Custom color for the maximum values.

.PARAMETER ColorScale2
    to be used with a custom 2colorscale formatting style.

.PARAMETER ColorScale3
    to be used with a custom 3colorscale formatting style.

.PARAMETER MidPointType
    to be used with a custom 3color scale formatting style.
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
   'Number','Percent','Formula','Percentile'

.PARAMETER MidPointValue
    The mid value in the range.

.PARAMETER MidPointColor
    Custom color for the mid values.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatColorScale -WorksheetName sheet7 -Range D4:D15 -ColorScaleType GreenYellowRed -Verbose
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    apply the built-in 3colorscale style 'GreenyellowRed' on range D4:D15.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatColorScale -WorksheetName sheet7 -Range F4:F15 -ColorScaleMinType Number -MinValue 12 -ColorScaleMinSystemColor Crimson -ColorScaleMaxType Number -MaxValue 99 -ColorScaleMaxSystemColor Yellow -ColorScale2
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    apply a custom 2colorscale style on range F4:F15.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatColorScale -WorksheetName sheet7 -Range h4:h15 -ColorScaleMinType Number -MinValue 12 -ColorScaleMinSystemColor Crimson -ColorScaleMaxType Number -MaxValue 99 -ColorScaleMaxSystemColor Yellow -MidPointType Number -MidPointValue 60 -MidPointColor Beige -ColorScale3
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    apply a custom 3colorscale style on range h4:h15.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> for($i=18 ;$i -le 89 ; $i++) { Set-SLConditionalFormatColorScale - -WorkBookInstance $doc -WorksheetName sheet7 -Range "C$($i):G$($i)" -ColorScaleType RedYellow -Verbose }
    PS C:\> $doc | Save-SLDocument


    Description
    -----------
    At times we may want to apply color scale formatting to individual rows instead of a range or rows.
    This example makes use of a for-loop to loop through rows 19 to 89 while applying the built-in style of 'RedYellow' on each row.


.INPUTS
   String,SpreadsheetLight.SLDocument

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    n\a

#>



    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true, Position = 0, ValueFromPipeLine = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [Alias('CurrentWorkSheetName')]
        [parameter(Mandatory = $true, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+:[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Set-SLConditionalFormatColorScale :`tRange should specify values in following format. Eg: A1:D10 or AB1:AD5"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true)]
        [string]$Range,

        [ValidateSet('GreenYellowRed', 'RedYellowGreen', 'BlueYellowRed', 'RedYellowBlue', 'GreenWhiteRed', 'RedWhiteGreen', 'BlueWhiteRed', 'RedWhiteBlue', 'WhiteRed', 'RedWhite', 'GreenWhite', 'WhiteGreen', 'Yellow',
            'Red', 'RedYellow', 'GreenYellow', 'YellowGreen')]
        [parameter(Mandatory = $True, Position = 3, ParameterSetName = 'Normal')]
        [string]$ColorScaleType,

        [ValidateSet('Value', 'Number', 'Percent', 'Formula', 'Percentile')]
        [parameter(Mandatory = $True, ParameterSetName = 'Custom3ColorScale')]
        [parameter(Mandatory = $True, ParameterSetName = 'Custom2ColorScale')]
        [String]$ColorScaleMinType,

        [parameter(Mandatory = $True, ParameterSetName = 'Custom3ColorScale')]
        [parameter(Mandatory = $True, ParameterSetName = 'Custom2ColorScale')]
        $MinValue,

        [Validateset('AliceBlue', 'AntiqueWhite', 'Aqua', 'Aquamarine', 'Azure', 'Beige', 'Bisque', 'Black', 'BlanchedAlmond', 'Blue', 'BlueViolet', 'Brown', 'BurlyWood', 'CadetBlue', 'Chartreuse', 'Chocolate', 'Coral', 'CornflowerBlue', 'Cornsilk',
            'Crimson', 'Cyan', 'DarkBlue', 'DarkCyan', 'DarkGoldenrod', 'DarkGray', 'DarkGreen', 'DarkKhaki', 'DarkMagenta', 'DarkOliveGreen', 'DarkOrange', 'DarkOrchid', 'DarkRed', 'DarkSalmon', 'DarkSeaGreen', 'DarkSlateBlue', 'DarkSlateGray',
            'DarkTurquoise', 'DarkViolet', 'DeepPink', 'DeepSkyBlue', 'DimGray', 'DodgerBlue', 'Firebrick', 'FloralWhite', 'ForestGreen', 'Fuchsia', 'Gainsboro', 'GhostWhite', 'Gold', 'Goldenrod', 'Gray', 'Green', 'GreenYellow', 'Honeydew', 'HotPink', 'IndianRed',
            'Indigo', 'Ivory', 'Khaki', 'LavENDer', 'LavENDerBlush', 'LawnGreen', 'LemonChiffon', 'LightBlue', 'LightCoral', 'LightCyan', 'LightGoldenrodYellow', 'LightGray', 'LightGreen', 'LightPink', 'LightSalmon', 'LightSeaGreen', 'LightSkyBlue', 'LightSlateGray',
            'LightSteelBlue', 'LightYellow', 'Lime', 'LimeGreen', 'Linen', 'Magenta', 'Maroon', 'MediumAquamarine', 'MediumBlue', 'MediumOrchid', 'MediumPurple', 'MediumSeaGreen', 'MediumSlateBlue', 'MediumSpringGreen', 'MediumTurquoise', 'MediumVioletRed', 'MidnightBlue'
            , 'MintCream', 'MistyRose', 'Moccasin', 'Name', 'NavajoWhite', 'Navy', 'OldLace', 'Olive', 'OliveDrab', 'Orange', 'OrangeRed', 'Orchid', 'PaleGoldenrod', 'PaleGreen', 'PaleTurquoise', 'PaleVioletRed', 'PapayaWhip', 'PeachPuff', 'Peru', 'Pink', 'Plum', 'PowderBlue',
            'Purple', 'Red', 'RosyBrown', 'RoyalBlue', 'SaddleBrown', 'Salmon', 'SandyBrown', 'SeaGreen', 'SeaShell', 'Sienna', 'Silver', 'SkyBlue', 'SlateBlue', 'SlateGray', 'Snow', 'SpringGreen', 'SteelBlue', 'Tan', 'Teal', 'Thistle', 'Tomato', 'Transparent', 'Turquoise',
            'Violet', 'Wheat', 'White', 'WhiteSmoke', 'Yellow', 'YellowGreen')]
        [parameter(Mandatory = $True, ParameterSetName = 'Custom3ColorScale')]
        [parameter(Mandatory = $True, ParameterSetName = 'Custom2ColorScale')]
        [string]$ColorScaleMinSystemColor,

        [ValidateSet('Value', 'Number', 'Percent', 'Formula', 'Percentile')]
        [parameter(Mandatory = $True, ParameterSetName = 'Custom3ColorScale')]
        [parameter(Mandatory = $True, ParameterSetName = 'Custom2ColorScale')]
        [String]$ColorScaleMaxType,

        [parameter(Mandatory = $True, ParameterSetName = 'Custom3ColorScale')]
        [parameter(Mandatory = $True, ParameterSetName = 'Custom2ColorScale')]
        $MaxValue,

        [Validateset('AliceBlue', 'AntiqueWhite', 'Aqua', 'Aquamarine', 'Azure', 'Beige', 'Bisque', 'Black', 'BlanchedAlmond', 'Blue', 'BlueViolet', 'Brown', 'BurlyWood', 'CadetBlue', 'Chartreuse', 'Chocolate', 'Coral', 'CornflowerBlue', 'Cornsilk',
            'Crimson', 'Cyan', 'DarkBlue', 'DarkCyan', 'DarkGoldenrod', 'DarkGray', 'DarkGreen', 'DarkKhaki', 'DarkMagenta', 'DarkOliveGreen', 'DarkOrange', 'DarkOrchid', 'DarkRed', 'DarkSalmon', 'DarkSeaGreen', 'DarkSlateBlue', 'DarkSlateGray',
            'DarkTurquoise', 'DarkViolet', 'DeepPink', 'DeepSkyBlue', 'DimGray', 'DodgerBlue', 'Firebrick', 'FloralWhite', 'ForestGreen', 'Fuchsia', 'Gainsboro', 'GhostWhite', 'Gold', 'Goldenrod', 'Gray', 'Green', 'GreenYellow', 'Honeydew', 'HotPink', 'IndianRed',
            'Indigo', 'Ivory', 'Khaki', 'LavENDer', 'LavENDerBlush', 'LawnGreen', 'LemonChiffon', 'LightBlue', 'LightCoral', 'LightCyan', 'LightGoldenrodYellow', 'LightGray', 'LightGreen', 'LightPink', 'LightSalmon', 'LightSeaGreen', 'LightSkyBlue', 'LightSlateGray',
            'LightSteelBlue', 'LightYellow', 'Lime', 'LimeGreen', 'Linen', 'Magenta', 'Maroon', 'MediumAquamarine', 'MediumBlue', 'MediumOrchid', 'MediumPurple', 'MediumSeaGreen', 'MediumSlateBlue', 'MediumSpringGreen', 'MediumTurquoise', 'MediumVioletRed', 'MidnightBlue'
            , 'MintCream', 'MistyRose', 'Moccasin', 'Name', 'NavajoWhite', 'Navy', 'OldLace', 'Olive', 'OliveDrab', 'Orange', 'OrangeRed', 'Orchid', 'PaleGoldenrod', 'PaleGreen', 'PaleTurquoise', 'PaleVioletRed', 'PapayaWhip', 'PeachPuff', 'Peru', 'Pink', 'Plum', 'PowderBlue',
            'Purple', 'Red', 'RosyBrown', 'RoyalBlue', 'SaddleBrown', 'Salmon', 'SandyBrown', 'SeaGreen', 'SeaShell', 'Sienna', 'Silver', 'SkyBlue', 'SlateBlue', 'SlateGray', 'Snow', 'SpringGreen', 'SteelBlue', 'Tan', 'Teal', 'Thistle', 'Tomato', 'Transparent', 'Turquoise',
            'Violet', 'Wheat', 'White', 'WhiteSmoke', 'Yellow', 'YellowGreen')]
        [parameter(Mandatory = $True, ParameterSetName = 'Custom3ColorScale')]
        [parameter(Mandatory = $True, ParameterSetName = 'Custom2ColorScale')]
        [string]$ColorScaleMaxSystemColor,


        [parameter(ParameterSetName = 'Custom2ColorScale')]
        [Switch]$ColorScale2,

        [parameter(ParameterSetName = 'Custom3ColorScale')]
        [Switch]$ColorScale3,

        [ValidateSet('Number', 'Percent', 'Formula', 'Percentile')]
        [parameter(Mandatory = $True, ParameterSetName = 'Custom3ColorScale')]
        [String]$MidPointType,

        [parameter(Mandatory = $True, ParameterSetName = 'Custom3ColorScale')]
        [String]$MidPointValue,

        [Validateset('AliceBlue', 'AntiqueWhite', 'Aqua', 'Aquamarine', 'Azure', 'Beige', 'Bisque', 'Black', 'BlanchedAlmond', 'Blue', 'BlueViolet', 'Brown', 'BurlyWood', 'CadetBlue', 'Chartreuse', 'Chocolate', 'Coral', 'CornflowerBlue', 'Cornsilk',
            'Crimson', 'Cyan', 'DarkBlue', 'DarkCyan', 'DarkGoldenrod', 'DarkGray', 'DarkGreen', 'DarkKhaki', 'DarkMagenta', 'DarkOliveGreen', 'DarkOrange', 'DarkOrchid', 'DarkRed', 'DarkSalmon', 'DarkSeaGreen', 'DarkSlateBlue', 'DarkSlateGray',
            'DarkTurquoise', 'DarkViolet', 'DeepPink', 'DeepSkyBlue', 'DimGray', 'DodgerBlue', 'Firebrick', 'FloralWhite', 'ForestGreen', 'Fuchsia', 'Gainsboro', 'GhostWhite', 'Gold', 'Goldenrod', 'Gray', 'Green', 'GreenYellow', 'Honeydew', 'HotPink', 'IndianRed',
            'Indigo', 'Ivory', 'Khaki', 'LavENDer', 'LavENDerBlush', 'LawnGreen', 'LemonChiffon', 'LightBlue', 'LightCoral', 'LightCyan', 'LightGoldenrodYellow', 'LightGray', 'LightGreen', 'LightPink', 'LightSalmon', 'LightSeaGreen', 'LightSkyBlue', 'LightSlateGray',
            'LightSteelBlue', 'LightYellow', 'Lime', 'LimeGreen', 'Linen', 'Magenta', 'Maroon', 'MediumAquamarine', 'MediumBlue', 'MediumOrchid', 'MediumPurple', 'MediumSeaGreen', 'MediumSlateBlue', 'MediumSpringGreen', 'MediumTurquoise', 'MediumVioletRed', 'MidnightBlue'
            , 'MintCream', 'MistyRose', 'Moccasin', 'Name', 'NavajoWhite', 'Navy', 'OldLace', 'Olive', 'OliveDrab', 'Orange', 'OrangeRed', 'Orchid', 'PaleGoldenrod', 'PaleGreen', 'PaleTurquoise', 'PaleVioletRed', 'PapayaWhip', 'PeachPuff', 'Peru', 'Pink', 'Plum', 'PowderBlue',
            'Purple', 'Red', 'RosyBrown', 'RoyalBlue', 'SaddleBrown', 'Salmon', 'SandyBrown', 'SeaGreen', 'SeaShell', 'Sienna', 'Silver', 'SkyBlue', 'SlateBlue', 'SlateGray', 'Snow', 'SpringGreen', 'SteelBlue', 'Tan', 'Teal', 'Thistle', 'Tomato', 'Transparent', 'Turquoise',
            'Violet', 'Wheat', 'White', 'WhiteSmoke', 'Yellow', 'YellowGreen')]
        [parameter(Mandatory = $True, ParameterSetName = 'Custom3ColorScale')]
        [String]$MidPointColor



    )
    PROCESS
    {

        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {

            $startcellreference, $ENDcellreference = $range -split ':'
            $ConditionalFormatting = New-Object SpreadsheetLight.SLConditionalFormatting($startcellreference, $ENDcellreference)

            if ($PSCmdlet.ParameterSetName -eq 'Normal')
            {
                Write-Verbose ("Set-SLConditionalFormatColorScale :`t Selected ColorScaleType is '{0}'" -f $ColorScaleType)
                $ConditionalFormatting.SetColorScale([SpreadsheetLight.SLConditionalFormatColorScaleValues]::$ColorScaleType) | Out-Null
            }


            if ($PSCmdlet.ParameterSetName -eq 'Custom2ColorScale')
            {
                $ConditionalFormatting.SetCustom2ColorScale([SLCFMinMax]::$ColorScaleMinType, $MinValue, [Color]::$ColorScaleMinSystemColor, [SLCFMinMax]::$ColorScaleMaxType, $MaxValue, [Color]::$ColorScaleMaxSystemColor) | Out-Null
            }



            if ($PSCmdlet.ParameterSetName -eq 'Custom3ColorScale')
            {
                $ConditionalFormatting.SetCustom3ColorScale([SLCFMinMax]::$ColorScaleMinType, $MinValue, [Color]::$ColorScaleMinSystemColor, [SpreadsheetLight.SLConditionalFormatRangeValues]::$MidPointType, $MidPointValue, [color]::$MidPointColor, [SLCFMinMax]::$ColorScaleMaxType, $MaxValue, [Color]::$ColorScaleMaxSystemColor) | Out-Null
            }

            Write-Verbose ("Set-SLConditionalFormatColorScale :`t Applying conditional formatting color scale on Range '{0}'" -f $Range)
            $WorkBookInstance.AddConditionalFormatting($ConditionalFormatting) | Out-Null

            $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }#select-slworksheet

    }#process
    END
    {
    }
}

Function Set-SLConditionalFormattingOnText
{

    <#

.SYNOPSIS
    Apply conditional formatting Iconset to text instead of numbers.

.DESCRIPTION
    Apply conditional formatting Iconset to text instead of numbers.
    Excel iconsets are applied on numbers and there is currently no built-in method to apply it on text or strings.
    This cmdlet takes a range containing text, inserts a new column before it and then applies conditional formatting on it.
    You can only apply text formatting on a column that has 3 or less unique values in a given column. Eg: "Working","Stopped","Disabled"

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    This is the name of the worksheet that contains the cell range where formatting is to be applied.

.PARAMETER Range
    The range of cells containing text to which conditional formatting has to be applied.

.PARAMETER IconSet
    Built-in Iconset styles.
    Use tab or intellisense to select from a range of possible values.
    Default value is - ThreeSymbols
    Possible values are:
    'ThreeArrows','ThreeArrowsGray','ThreeFlags','ThreeSigns','ThreeStars',
        'ThreeSymbols','ThreeSymbols2','ThreeTrafficLights1','ThreeTrafficLights2','ThreeTriangles'

.PARAMETER Properties
    String containing comma seperated text values. EG: "Working,Stopped,Disabled"

.PARAMETER IconColumnHeader
    The header text to be set for the new column contining icons. Default value is - "Icon"

.PARAMETER ReverseIconorder
    Reverses the order in which icons are applied.

.PARAMETER ShowIconsOnly
    Will show just the icons instead of icons and numbers.
    The default value is true.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormattingOnText -WorksheetName sheet1 -Range f4:f10 -Properties "working,stopped,disabled"
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    This will insert a column before column F and the conditional formatting Icon set 'ThreeSymbols' will be applied to the new column.
    The cell corresponding to value working will be 'Green', stopped will be 'yello\orange' and disabled in 'Red'.
    Note: Column F becomes column G.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $Services = Get-WmiObject -Class Win32_Service | Sort State,StartMode | Select __Server,Name,Displayname,State,StartMode
    PS C:\> $services | Export-SLDocument -WorkBookInstance $doc -WorksheetName sheet3 -AutofitColumns
    PS C:\> $doc | Save-SLDocument
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormattingOnText -WorksheetName sheet3 -Range e5:e187 -Properties "Running,Stopped"
    PS C:\> $doc | Save-SLDocument
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormattingOnText -WorksheetName sheet3 -Range g5:g187 -Properties "Auto,Manual,Disabled"
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    Servicedata is exported to sheet3, workbook is saved and closed. We then open the workbook to determine the data range for conditionalformatting.
    We apply conditional formatting on the state column which has 2 properties "Running" and "stopped" save and close.
    We open the document again to determine the datarange corresponding to the startmode column which has 3 properties "Auto","Manual" & "Disabled"



.INPUTS
   String,SpreadsheetLight.SLDocument

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    n\a

#>




    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true, Position = 0, ValueFromPipeLine = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [Alias('CurrentWorkSheetName')]
        [parameter(Mandatory = $true, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [ValidateScript({
                $r1, $r2 = $_ -split ':'
                $r1_match = [regex]::Match($r1, '[a-zA-Z]+') | Select-Object -ExpandProperty value
                $r2_match = [regex]::Match($r2, '[a-zA-Z]+') | Select-Object -ExpandProperty value
                if ($r1_match -eq $r2_match) { $true }
                else { $false; Write-Warning "Set-SLConditionalFormattingOnText :`tVFormulacellRange should specify values that belong to the same column. Eg: A1:A10 or AB1:AB5"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true)]
        [string]$Range,

        [Validateset('ThreeArrows', 'ThreeArrowsGray', 'ThreeFlags', 'ThreeSigns', 'ThreeStars',
            'ThreeSymbols', 'ThreeSymbols2', 'ThreeTrafficLights1', 'ThreeTrafficLights2', 'ThreeTriangles')]
        [parameter(Mandatory = $false, Position = 2)]
        [string]$IconSet = 'ThreeSymbols',

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$Properties,

        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [string]$IconColumnHeader = 'Icon',

        [Switch]$ReverseIconorder = $true,
        [Switch]$ShowIconsOnly = $true





    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            $StartCellReference, $ENDCellReference = $Range -split ':'

            $RangeStats = Convert-ToExcelRowColumnStats -Range $Range
            $MatchedColumnIndex = $RangeStats.StartColumnIndex + 1
            $MatchedColumnName = Convert-ToExcelColumnName -Index $MatchedColumnIndex
            $MatchedcellReference = $MatchedColumnName + ($RangeStats.StartRowIndex + 1)

            $WorkBookInstance.InsertColumn($RangeStats.StartColumnIndex, 1) | Out-Null

            $NewColumnHeaderCellreference = $RangeStats.StartColumnName + $RangeStats.StartRowIndex
            $WorkBookInstance.SetCellValue("$NewColumnHeaderCellreference", $IconColumnHeader ) | Out-Null

            $NewIconColumnName = $RangeStats.StartColumnName
            $NewIconColumnIndex = $RangeStats.StartColumnIndex
            $NewIconRowIndex = $RangeStats.StartRowIndex + 1
            $endrowIndex = $RangeStats.EndRowIndex

            $Fproperties = (($Properties -split ',' | ForEach-Object { '"' + $_ + '"' } ) -join ',').ToString()

            for ($i = $NewIconRowIndex; $i -le $endrowIndex; $i++)
            {
                $MatchedcellReference = $MatchedColumnName + $NewIconRowIndex
                $WorkBookInstance.SetCellValue(($NewIconColumnName + $i), "=MATCH($MatchedcellReference,{$Fproperties},0)") | Out-Null
                $NewIconRowIndex++
            }


            $ConditionalFormatting = New-Object SpreadsheetLight.SLConditionalFormatting($startcellreference, $ENDcellreference)

            $ConditionalFormatting.SetCustomIconSet([SpreadsheetLight.SLThreeIconSetValues]::$IconSet, $ReverseIconorder, $ShowIconsOnly, $true, 2, [SpreadsheetLight.SLConditionalFormatRangeValues]::'Number', $true, 3, [SpreadsheetLight.SLConditionalFormatRangeValues]::'Number') | Out-Null

            Write-Verbose ("Set-SLConditionalFormatColorScale :`t Applying conditional formatting color scale on Range '{0}'" -f $Range)
            $WorkBookInstance.AddConditionalFormatting($ConditionalFormatting) | Out-Null

            $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }#select-slworksheet

    }#process
}


Function Set-SLConditionalFormatIconSet
{

    <#

.SYNOPSIS
    Apply conditional formatting Iconset on numbers.

.DESCRIPTION
    Apply conditional formatting Iconset on numbers.
    Based on the data users may select  3, 4 or 5iconsets to display data.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    This is the name of the worksheet that contains the cell range where formatting is to be applied.

.PARAMETER Range
    The range of cells containing text to which conditional formatting has to be applied.

.PARAMETER IconSet
    Built-in Iconset styles.
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'FiveArrows','FiveArrowsGray','FiveQuarters','FiveRating','FourArrows',
        'FourArrowsGray','FourRating','FourRedToBlack','FourTrafficLights',
        'ThreeArrows','ThreeArrowsGray','ThreeFlags','ThreeSigns',
        'ThreeSymbols','ThreeSymbols2','ThreeTrafficLights1','ThreeTrafficLights2'

.PARAMETER FiveIconSetType
    Use this to apply different formatting types on 5 different ranges.

.PARAMETER FourIconSetType
    Use this to apply different formatting types on 4 different ranges.

.PARAMETER ThreeIconSetType
    Use this to apply different formatting types on 4 different ranges.

.PARAMETER ReverseIconOrder
   Reverse the order of the icons displayed.

.PARAMETER ShowIconsOnly
    Will show just the icons instead of icons and numbers.


.PARAMETER GreaterThanOrEqual2
    True if values are to be greater than or equal to the 2nd range value.False if values are to be strictly greater than.

.PARAMETER SecondRangeValue
    The 2nd Range value.

.PARAMETER SecondRangeValueType
    Built-in Iconset format types.
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'Number','Percent','Formula','Percentile'


.PARAMETER ThirdRangeValueType
    Built-in Iconset format types.
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'Number','Percent','Formula','Percentile'

.PARAMETER FourthRangeValueType
    Built-in Iconset format types.
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'Number','Percent','Formula','Percentile'


.PARAMETER FifthRangeValueType
    Built-in Iconset format types.
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'Number','Percent','Formula','Percentile'

.PARAMETER GreaterThanOrEqual3
    True if values are to be greater than or equal to the 3rd range value.False if values are to be strictly greater than.

.PARAMETER ThirdRangeValue
    The 3rd range value.

.PARAMETER GreaterThanOrEqual4
    True if values are to be greater than or equal to the 4th range value.False if values are to be strictly greater than.

.PARAMETER FourthRangeValue
    The 4th range value.
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'Number','Percent','Formula','Percentile'

.PARAMETER GreaterThanOrEqual5
    True if values are to be greater than or equal to the 5th range value.False if values are to be strictly greater than.

.PARAMETER FifthRangeValue
    The 5th range value.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatIconSet -WorksheetName sheet7 -Range d4:d15 -IconSet ThreeSymbols -Verbose
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    Apply the conditional formatting Icon set 'ThreeSymbols' to the range d4:d15. Both icons and values are shown


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatIconSet -WorksheetName sheet7 -Range f4:f15 -FiveIconSetType FiveRating -ReverseIconOrder:$false -ShowIconOnly -GreaterThanOrEqual2 -SecondRangeValue 15 -SecondRangeValueType Percentile -GreaterThanOrEqual3 -ThirdRangeValue 35 -ThirdRangeValueType Percentile -GreaterThanOrEqual4 -FourthRangeValue 67 -FourthRangeValueType Percentile -GreaterThanOrEqual5 -FifthRangeValue 80 -FifthRangeValueType Percentile
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    Apply the conditional formatting Icon set 'FiveRating' to the range f4:f15. Only icons are shown.

.Example
    PS C:\> $IconSet5Params = @{

        WorkBookInstance = ($doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx)
        WorksheetName = 'sheet7'
        Range = 'f4:f15'
        FiveIconSetType = 'FiveRating'
        ReverseIconOrder = $false
        ShowIconOnly = $true
        GreaterThanOrEqual2 = $true
        SecondRangeValue = 15
        SecondRangeValueType = 'Percentile'
        GreaterThanOrEqual3 = $true
        ThirdRangeValue = 35
        ThirdRangeValueType = 'Percentile'
        GreaterThanOrEqual4 = $true
        FourthRangeValue = 67
        FourthRangeValueType = 'Percentile'
        GreaterThanOrEqual5 = $true
        FifthRangeValue = 80
        FifthRangeValueType = 'Percentile'
        Verbose = $true
}

    PS C:\> Set-SLConditionalFormatIconSet @IconSet5Params
    PS C:\> $doc | Save-SLDocument


    Description
    -----------
    Since the last example had a lot of parameters and values that scrolled off to the right this example will retain the same values\parameters
    but will use a different format of data input to the cmdlet which will make it easier to read.
    All the parameters and values required to run the cmdlet Set-SLConditionalFormatIconSet are stored in the variable - IconSet5Params
    which is a hashtable that contains Key\Value pairs.
    The keys are the parameters and the values are the parameter values.
    Note: You’ll notice a little trick here. The “@” sign is followed by the variable name "IconSet5Params", which doesn’t include the dollar sign.
    The “@” sign, when used as a splat operator says,
    “Take whatever characters come next and assume they’re a variable name. Assume that the variable contains a hashtable, and that the keys are parameter names.
    The above explanation of the @ 'splat' operator is a direct quote from Don jones :)



.Example

    PS C:\> $IconSet3Params = @{

        WorkBookInstance = (Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx)
        WorksheetName = 'sheet7'
        Range = 'h4:h15'
        ThreeIconSetType = 'ThreeTrafficLights1'
        ReverseIconOrder = $false
        ShowIconOnly = $false
        GreaterThanOrEqual2 = $true
        SecondRangeValue = 33
        SecondRangeValueType = 'Number'
        GreaterThanOrEqual3 = $true
        ThirdRangeValue = 82
        ThirdRangeValueType = 'Number'
        Verbose = $true
}


    PS C:\> $IconSet3ReverseIconsParams = @{

        WorkBookInstance = (Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx)
        WorksheetName = 'sheet7'
        Range = 'j4:j15'
        ThreeIconSetType = 'ThreeTrafficLights1'
        ReverseIconOrder = $true
        ShowIconOnly = $true
        GreaterThanOrEqual2 = $true
        SecondRangeValue = 33
        SecondRangeValueType = 'Number'
        GreaterThanOrEqual3 = $true
        ThirdRangeValue = 82
        ThirdRangeValueType = 'Number'
        Verbose = $true
}

    PS C:\> Set-SLConditionalFormatIconSet @IconSet3Params | Save-SLDocument
    PS C:\> Set-SLConditionalFormatIconSet @IconSet3ReverseIconsParams | Save-SLDocument


    Description
    -----------
    Here we apply conditional formatting twice on two different ranges. h4:h15 & J4:J15.
    The only difference between the 2 is that the second range J4;J15 has the icon order reversed.


.INPUTS
   String,Int,Bool,SpreadsheetLight.SLDocument

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    n\a

#>


    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true, Position = 0, ValueFromPipeLine = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [Alias('CurrentWorkSheetName')]
        [parameter(Mandatory = $true, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,


        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+:[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Set-SLConditionalFormatIconSet :`tRange should specify values in following format. Eg: A1:D10 or AB1:AD5"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true)]
        [string]$Range,

        [Validateset('FiveArrows', 'FiveArrowsGray', 'FiveQuarters', 'FiveRating', 'FourArrows',
            'FourArrowsGray', 'FourRating', 'FourRedToBlack', 'FourTrafficLights',
            'ThreeArrows', 'ThreeArrowsGray', 'ThreeFlags', 'ThreeSigns',
            'ThreeSymbols', 'ThreeSymbols2', 'ThreeTrafficLights1', 'ThreeTrafficLights2')]
        [parameter(Mandatory = $true, Position = 3, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'IconSet')]
        [string]$IconSet,

        [Validateset('FiveArrows', 'FiveArrowsGray', 'FiveBoxes', 'FiveQuarters', 'FiveRating')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'CustomIconSet5')]
        [string]$FiveIconSetType,

        [Validateset('FourArrows', 'FourArrowsGray', 'FourRating', 'FourRedToBlack', 'FourTrafficLights')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'CustomIconSet4')]
        [string]$FourIconSetType,

        [Validateset('ThreeArrows', 'ThreeArrowsGray', 'ThreeFlags', 'ThreeSigns', 'ThreeStars', 'ThreeSymbols', 'ThreeSymbols2', 'ThreeTrafficLights1', 'ThreeTrafficLights2', 'ThreeTriangles')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'CustomIconSet3')]
        [string]$ThreeIconSetType,

        [parameter(ParameterSetName = 'CustomIconSet3')]
        [parameter(ParameterSetName = 'CustomIconSet4')]
        [parameter(ParameterSetName = 'CustomIconSet5')]
        [switch]$ReverseIconOrder,

        [parameter(ParameterSetName = 'CustomIconSet3')]
        [parameter(ParameterSetName = 'CustomIconSet4')]
        [parameter(ParameterSetName = 'CustomIconSet5')]
        [switch]$ShowIconOnly,

        [parameter(ParameterSetName = 'CustomIconSet3')]
        [parameter(ParameterSetName = 'CustomIconSet4')]
        [parameter(ParameterSetName = 'CustomIconSet5')]
        [switch]$GreaterThanOrEqual2,

        [parameter(ParameterSetName = 'CustomIconSet3')]
        [parameter(ParameterSetName = 'CustomIconSet4')]
        [parameter(ParameterSetName = 'CustomIconSet5')]
        [string]$SecondRangeValue,

        [Validateset('Number', 'Percent', 'Formula', 'Percentile')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'CustomIconSet3')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'CustomIconSet4')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'CustomIconSet5')]
        [string]$SecondRangeValueType,

        [Validateset('Number', 'Percent', 'Formula', 'Percentile')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'CustomIconSet3')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'CustomIconSet4')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'CustomIconSet5')]
        [string]$ThirdRangeValueType,

        [Validateset('Number', 'Percent', 'Formula', 'Percentile')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'CustomIconSet4')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'CustomIconSet5')]
        [string]$FourthRangeValueType,

        [Validateset('Number', 'Percent', 'Formula', 'Percentile')]
        [parameter(Mandatory = $true, ValueFromPipeLineByPropertyName = $true, ParameterSetName = 'CustomIconSet5')]
        [string]$FifthRangeValueType,

        [parameter(ParameterSetName = 'CustomIconSet3')]
        [parameter(ParameterSetName = 'CustomIconSet4')]
        [parameter(ParameterSetName = 'CustomIconSet5')]
        [switch]$GreaterThanOrEqual3,

        [parameter(ParameterSetName = 'CustomIconSet3')]
        [parameter(ParameterSetName = 'CustomIconSet4')]
        [parameter(ParameterSetName = 'CustomIconSet5')]
        [string]$ThirdRangeValue,

        [parameter(ParameterSetName = 'CustomIconSet4')]
        [parameter(ParameterSetName = 'CustomIconSet5')]
        [switch]$GreaterThanOrEqual4,

        [parameter(ParameterSetName = 'CustomIconSet4')]
        [parameter(ParameterSetName = 'CustomIconSet5')]
        [string]$FourthRangeValue,

        [parameter(ParameterSetName = 'CustomIconSet5')]
        [switch]$GreaterThanOrEqual5,

        [parameter(ParameterSetName = 'CustomIconSet5')]
        [string]$FifthRangeValue


    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {

            $startcellreference, $ENDcellreference = $range -split ':'
            $ConditionalFormatting = New-Object SpreadsheetLight.SLConditionalFormatting($startcellreference, $ENDcellreference)

            if ($PSCmdlet.ParameterSetName -eq 'IconSet')
            {
                Write-Verbose ("Set-SLConditionalFormatIconSet :`t Selected Iconset is '{0}'" -f $IconSet)
                $ConditionalFormatting.SetIconSet([OLIconsetvalues]::$IconSet) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'CustomIconSet5')
            {
                Write-Verbose ("Set-SLConditionalFormatIconSet :`t Selected Iconset is '{0}'" -f $FiveIconSetType)
                $ConditionalFormatting.SetCustomIconSet([SpreadsheetLight.SLFiveIconSetValues]::$FiveIconSetType, $ReverseIconOrder, $ShowIconOnly, $GreaterThanOrEqual2, $SecondRangeValue, [SLCFRangeValues]::$SecondRangeValueType, $GreaterThanOrEqual3, $ThirdRangeValue, [SLCFRangeValues]::$ThirdRangeValueType, $GreaterThanOrEqual4, $FourthRangeValue, [SLCFRangeValues]::$FourthRangeValueType, $GreaterThanOrEqual5, $FifthRangeValue, [SLCFRangeValues]::$FifthRangeValueType ) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'CustomIconSet4')
            {
                Write-Verbose ("Set-SLConditionalFormatIconSet :`t Selected Iconset is '{0}'" -f $FourIconSetType)
                $ConditionalFormatting.SetCustomIconSet([SpreadsheetLight.SLFourIconSetValues]::$FourIconSetType, $ReverseIconOrder, $ShowIconOnly, $GreaterThanOrEqual2, $SecondRangeValue, [SLCFRangeValues]::$SecondRangeValueType, $GreaterThanOrEqual3, $ThirdRangeValue, [SLCFRangeValues]::$ThirdRangeValueType, $GreaterThanOrEqual4, $FourthRangeValue, [SLCFRangeValues]::$FourthRangeValueType ) | Out-Null
            }

            if ($PSCmdlet.ParameterSetName -eq 'CustomIconSet3')
            {
                Write-Verbose ("Set-SLConditionalFormatIconSet :`t Selected Iconset is '{0}'" -f $ThreeIconSetType)
                $ConditionalFormatting.SetCustomIconSet([SpreadsheetLight.SLThreeIconSetValues]::$ThreeIconSetType, $ReverseIconOrder, $ShowIconOnly, $GreaterThanOrEqual2, $SecondRangeValue, [SLCFRangeValues]::$SecondRangeValueType, $GreaterThanOrEqual3, $ThirdRangeValue, [SLCFRangeValues]::$ThirdRangeValueType ) | Out-Null
            }


            Write-Verbose ("Set-SLConditionalFormatIconSet :`t Applying conditional formatting IconSet on Range '{0}'" -f $Range)
            $WorkBookInstance.AddConditionalFormatting($ConditionalFormatting) | Out-Null

            $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru


        }
    }

}


Function Set-SLConditionalFormatHighLights
{

    <#

.SYNOPSIS
    Apply conditional formatting Highlights.

.DESCRIPTION
    Apply conditional formatting Highlights on cells.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    This is the name of the worksheet that contains the cell range where formatting is to be applied.

.PARAMETER Range
    The range of cells containing text to which conditional formatting has to be applied.

.PARAMETER StyleType
    Choose between excel's 'PresetStyle' or a 'CustomStyle'.

.PARAMETER PresetStyleValue
    Built-in Preset styles.
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'LightRedFillWithDarkRedText','YellowFillWithDarkYellowText','GreenFillWithDarkGreenText','LightRedFill','RedText','RedBorder'

.PARAMETER IsBetweenFirstValue
    Minumum value to be used when specifying a range that is 'Between' two values.

.PARAMETER IsBetweenLastValue
    Maximum value to be used when specifying a range that is 'Between' two values.

.PARAMETER IsNotBetweenFirstValue
    Minumum value to be used when specifying a range that is 'NOTBetween' two values.

.PARAMETER IsNotBetweenLastValue
    Maximum value to be used when specifying a range that is 'NOTBetween' two values.

.PARAMETER TopRankValue
    Top rank value to be used when specifying top and bottom ranks.

.PARAMETER BottomRankValue
    Bottom rank value to be used when specifying top and bottom ranks.

.PARAMETER IsPercent
    Specifies that values should be considered as numbers.

.PARAMETER IsItems
    Specifies that values should be considered as a percentage.

.PARAMETER GreaterThanValue
    Highlight values that are greater than this value.

.PARAMETER GreaterThanorEqualToValue
    Highlight values that are greater than or Equalto this value.

.PARAMETER LessThanValue
    Highlight values that are less than this value.

.PARAMETER LessThanorEqualToValue
    Highlight values that are less than or Equalto this value.

.PARAMETER EqualToValue
    Highlight values that are Equalto this value.

.PARAMETER NotEqualToValue
    Highlight values that are NOTEqualto this value.

.PARAMETER TextContainsString
    Highlight cells that contain this string.

.PARAMETER TextDoesNotContainString
    Highlight cells that DO NOT contain this string.

.PARAMETER TextENDsWithString
    Highlight cells that End with this string.

.PARAMETER TextBEGINsWithString
    Highlight cells that begin with this string.

.PARAMETER AverageType
    Built-in AverageType values.
    Use tab or intellisense to select from a range of possible values.
    Possible values are:
    'Above','Below','EqualOrAbove','EqualOrBelow','OneStdDevAbove','OneStdDevBelow',
    'TwoStdDevAbove','TwoStdDevBelow','ThreeStdDevAbove','ThreeStdDevBelow'

.PARAMETER DateString
    Highlight cells that match the date specified by this value.

.PARAMETER FormulaString
    Highlight cells that match the criteria specified by a formula.

.PARAMETER HighLightDuplicateValues
    Highlight all duplicate values in a range.

.PARAMETER HighLightUniqueValues
    Highlight all unique values in a range.

.PARAMETER HighlightBlankCells
    Highlight all blank cells in a range.

.PARAMETER HighlightNonBlankCells
    Highlight all non-blank cells in a range.

.PARAMETER HighlightErrorCells
    Highlight all cells containing formula errors in a range.

.PARAMETER HighlightNonErrorCells
    Highlight all cells that dont contain formula errors in a range.

.PARAMETER FontColor
    Fontcolor to be specified when using a custom highlight style.

.PARAMETER FontIsBold
    Specify that the font is bold when using a custom highlight style.

.PARAMETER FontIsItalic
    Specify that the font is italic when using a custom highlight style.

.PARAMETER FontIsUnderlined
    Specify that the font is underlined when using a custom highlight style.

.PARAMETER FillColor
    Fill Color to be specified when using a custom highlight style.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatHighLights -WorksheetName sheet9 -Range a2:b9 -StyleType PresetStyle -PresetStyleValue LightRedFill -HighLightDuplicateValues | Save-SLDocument

    Description
    -----------
    Highlight Duplicate values


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatHighLights -WorksheetName sheet9 -Range a11:a21 -StyleType PresetStyle -PresetStyleValue RedBorder -HighlightBlankCells | Save-SLDocument

    Description
    -----------
    Highlight blanks cells.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatHighLights -WorksheetName sheet9 -Range a11:a21 -StyleType PresetStyle -PresetStyleValue RedText -HighlightNonBlankCells | Save-SLDocument

    Description
    -----------
    Highlight non-blank cells


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatHighLights -WorksheetName sheet9 -Range b11:b21 -StyleType PresetStyle -PresetStyleValue LightRedFillWithDarkRedText -HighlightErrorCells | Save-SLDocument

    Description
    -----------
    Highlight Error cells.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatHighLights -WorksheetName sheet9 -Range b11:b21 -StyleType PresetStyle -PresetStyleValue GreenFillWithDarkGreenText -HighlightNonErrorCells | Save-SLDocument

    Description
    -----------
    Highlight non-Error cells


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatHighLights -WorksheetName sheet9 -Range C2:C9 -StyleType PresetStyle -PresetStyleValue YellowFillWithDarkYellowText -IsBetweenFirstValue 200 -IsBetweenLastValue 400 | Save-SLDocument

    Description
    -----------
    Highlight values between 200 and 400.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatHighLights -WorksheetName sheet9 -Range C2:C9 -StyleType PresetStyle -PresetStyleValue GreenFillWithDarkGreenText -IsNotBetweenFirstValue 200 -IsNotBetweenLastValue  400 | Save-SLDocument

    Description
    -----------
    Highlight values NOT between 200 and 400


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatHighLights -WorksheetName sheet9 -Range c11:c21 -StyleType PresetStyle -PresetStyleValue GreenFillWithDarkGreenText -TopRankValue 25 -IsPercent | Save-SLDocument

    Description
    -----------
    Highlight TOP 25% of the values.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatHighLights -WorksheetName sheet9 -Range E2:E9 -StyleType PresetStyle -PresetStyleValue YellowFillWithDarkYellowText -TopRankValue 3 -IsItems | Save-SLDocument

    Description
    -----------
    Highlight TOP 3 items.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatHighLights -WorksheetName sheet9 -Range c11:c21 -StyleType PresetStyle -PresetStyleValue LightRedFillWithDarkRedText  -BottomRankValue 25 -IsPercent | Save-SLDocument

    Description
    -----------
    Highlight BOTTOM 25%.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatHighLights -WorksheetName sheet9 -Range E2:E9 -StyleType PresetStyle -PresetStyleValue LightRedFill -BottomRankValue 3 -IsItems | Save-SLDocument

    Description
    -----------
    Highlight BOTTOM 3 items


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatHighLights -WorksheetName sheet9 -Range E11:E21 -StyleType CustomStyle -GreaterThanValue 200 -FontColor Blue -FontIsBold -FontIsItalic -FontIsUnderlined -FillColor Yellow | Save-SLDocument

    Description
    -----------
    Highlight values Greaterthan 200 - Custom Style.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatHighLights -WorksheetName sheet9 -Range E11:E21 -StyleType CustomStyle -LessThanValue 11 -FontColor White -FontIsBold -FillColor Darkblue | Save-SLDocument

    Description
    -----------
    Highlight values LessThan 11 - Custom Style


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatHighLights -WorksheetName sheet9 -Range G2:G9 -StyleType PresetStyle -PresetStyleValue GreenFillWithDarkGreenText -TextENDsWithString bob | Save-SLDocument

    Description
    -----------
    Highlight cells that END with 'Bob'


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\excel\MyFirstDoc.xlsx
    PS C:\> $doc | Set-SLConditionalFormatHighLights -WorksheetName sheet9 -Range G2:G9 -StyleType PresetStyle -PresetStyleValue YellowFillWithDarkYellowText  -TextContainsString jones | Save-SLDocument

    Description
    -----------
    Highlight cells that Contain 'Jones'.


.INPUTS
   String,Int,Date,SpreadsheetLight.SLDocument

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    n\a

#>








    [CmdletBinding()]
    param (

        [parameter(Mandatory = $true, Position = 0, ValueFromPipeLine = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [Alias('CurrentWorkSheetName')]
        [parameter(Mandatory = $true, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+:[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Set-SLConditionalFormattingHighLights :`tRange should specify values in following format. Eg: A1:D10 or AB1:AD5"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipeLineByPropertyName = $true)]
        [string]$Range,

        [ValidateSet('CustomStyle', 'PresetStyle')]
        [parameter(Mandatory = $True)]
        [String]$StyleType,

        [ValidateSet('LightRedFillWithDarkRedText', 'YellowFillWithDarkYellowText', 'GreenFillWithDarkGreenText', 'LightRedFill', 'RedText', 'RedBorder')]
        [parameter(Mandatory = $false)]
        [string]$PresetStyleValue,

        [parameter(Mandatory = $True, ParameterSetName = 'Between')]
        [Double]$IsBetweenFirstValue,

        [parameter(Mandatory = $True, ParameterSetName = 'Between')]
        [Double]$IsBetweenLastValue,

        [parameter(Mandatory = $True, ParameterSetName = 'NotBetween')]
        [Double]$IsNotBetweenFirstValue,

        [parameter(Mandatory = $True, ParameterSetName = 'NotBetween')]
        [Double]$IsNotBetweenLastValue,

        [parameter(Mandatory = $True, ParameterSetName = 'TopRank')]
        [System.UInt32]$TopRankValue,

        [parameter(Mandatory = $True, ParameterSetName = 'BottomRank')]
        [System.UInt32]$BottomRankValue,

        [parameter(ParameterSetName = 'TopRank')]
        [parameter(ParameterSetName = 'BottomRank')]
        [Switch]$IsPercent,

        [parameter(ParameterSetName = 'TopRank')]
        [parameter(ParameterSetName = 'BottomRank')]
        [Switch]$IsItems,

        [parameter(Mandatory = $True, ParameterSetName = 'GreaterThan')]
        [Double]$GreaterThanValue,

        [parameter(Mandatory = $True, ParameterSetName = 'GreaterThanorEqualTo')]
        [Double]$GreaterThanorEqualToValue,

        [parameter(Mandatory = $True, ParameterSetName = 'LessThan')]
        [Double]$LessThanValue,

        [parameter(Mandatory = $True, ParameterSetName = 'LessThanorEqualTo')]
        [Double]$LessThanorEqualToValue,

        [parameter(Mandatory = $True, ParameterSetName = 'EqualTo')]
        [String]$EqualToValue,

        [parameter(Mandatory = $True, ParameterSetName = 'NotEqualTo')]
        [String]$NotEqualToValue,

        [parameter(Mandatory = $True, ParameterSetName = 'TextContains')]
        [String]$TextContainsString,

        [parameter(Mandatory = $True, ParameterSetName = 'TextDoesNotContain')]
        [String]$TextDoesNotContainString,

        [parameter(Mandatory = $True, ParameterSetName = 'TextENDsWith')]
        [String]$TextENDsWithString,

        [parameter(Mandatory = $True, ParameterSetName = 'TextBEGINsWith')]
        [String]$TextBEGINsWithString,

        [ValidateSet('Above', 'Below', 'EqualOrAbove', 'EqualOrBelow', 'OneStdDevAbove', 'OneStdDevBelow', 'TwoStdDevAbove', 'TwoStdDevBelow', 'ThreeStdDevAbove', 'ThreeStdDevBelow')]
        [parameter(Mandatory = $True, ParameterSetName = 'Average')]
        [String]$AverageType,

        [parameter(Mandatory = $True, ParameterSetName = 'Formula')]
        [String]$FormulaString,

        [parameter(Mandatory = $True, ParameterSetName = 'Date')]
        [String]$DateString,

        [parameter(Mandatory = $True, ParameterSetName = 'Duplicate')]
        [Switch]$HighLightDuplicateValues,

        [parameter(Mandatory = $True, ParameterSetName = 'Unique')]
        [Switch]$HighLightUniqueValues,

        [parameter(Mandatory = $True, ParameterSetName = 'Blank')]
        [Switch]$HighlightBlankCells,

        [parameter(Mandatory = $True, ParameterSetName = 'NonBlank')]
        [Switch]$HighlightNonBlankCells,

        [parameter(Mandatory = $True, ParameterSetName = 'Error')]
        [Switch]$HighlightErrorCells,

        [parameter(Mandatory = $True, ParameterSetName = 'NonError')]
        [Switch]$HighlightNonErrorCells,


        [parameter(Mandatory = $False)]
        [String]$FontColor,

        [Switch]$FontIsBold,

        [Switch]$FontIsItalic,

        [Switch]$FontIsUnderlined,

        [parameter(Mandatory = $False)]
        [String]$FillColor





    )
    PROCESS
    {

        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {

            $startcellreference, $ENDcellreference = $range -split ':'
            $ConditionalFormatting = New-Object SpreadsheetLight.SLConditionalFormatting($startcellreference, $ENDcellreference)

            If ($StyleType -eq 'CustomStyle')
            {
                $SLStyle = $WorkBookInstance.CreateStyle()

                if ($FontColor) { $SLStyle.SetFontColor([System.Drawing.Color]::$FontColor) | Out-Null }
                if ($FontIsBold) { $SLStyle.SetFontBold($true) | Out-Null }
                if ($FontIsItalic) { $SLStyle.SetFontItalic($true) | Out-Null }
                if ($FontIsUnderlined) { $SLStyle.SetFontUnderline([DocumentFormat.OpenXml.Spreadsheet.UnderlineValues]::'Single') | Out-Null }
                if ($FillColor)
                {
                    $SLStyle.Fill.SetPatternType([DocumentFormat.OpenXml.Spreadsheet.PatternValues]::'Solid') | Out-Null
                    $SLStyle.Fill.SetPatternBackgroundColor([System.Drawing.Color]::$FillColor) | Out-Null
                }

            }
            elseif ($StyleType -eq 'PresetStyle')
            {
                $PresetStyle = [SpreadsheetLight.SLHighlightCellsStyleValues]::$PresetStyleValue
            }


            if ($PSCmdlet.ParameterSetName -eq 'GreaterThan')
            {

                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsGreaterThan($False, $GreaterThanValue, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsGreaterThan($False, $GreaterThanValue, $SLStyle) | Out-Null }
            }

            if ($PSCmdlet.ParameterSetName -eq 'GreaterThanorEqualTo')
            {

                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsGreaterThan($True, $GreaterThanorEqualToValue, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsGreaterThan($True, $GreaterThanorEqualToValue, $SLStyle) | Out-Null }
            }

            ## // less than

            if ($PSCmdlet.ParameterSetName -eq 'LessThan')
            {

                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsLessThan($False, $LessThanValue, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsLessThan($False, $LessThanValue, $SLStyle) | Out-Null }
            }

            if ($PSCmdlet.ParameterSetName -eq 'LessThanorEqualTo')
            {

                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsLessThan($True, $LessThanorEqualToValue, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsLessThan($True, $LessThanorEqualToValue, $SLStyle) | Out-Null }
            }

            ## // Between

            if ($PSCmdlet.ParameterSetName -eq 'Between')
            {

                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsBetween($True, $IsBetweenFirstValue, $IsBetweenLastValue, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsBetween($True, $IsBetweenFirstValue, $IsBetweenLastValue, $SLStyle) | Out-Null }
            }

            if ($PSCmdlet.ParameterSetName -eq 'NotBetween')
            {

                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsBetween($False, $IsNotBetweenFirstValue, $IsNotBetweenLastValue, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsBetween($False, $IsNotBetweenFirstValue, $IsNotBetweenLastValue, $SLStyle) | Out-Null }
            }

            ## // Range

            if ($PSCmdlet.ParameterSetName -eq 'TopRank')
            {
                if ($IsPercent) { $percentoritems = $true }
                elseif ($IsItems) { $percentoritems = $false }
                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsInTopRange($True, $TopRankValue, $percentoritems, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsInTopRange($True, $TopRankValue, $percentoritems, $SLStyle) | Out-Null }
            }

            if ($PSCmdlet.ParameterSetName -eq 'BottomRank')
            {

                if ($IsPercent) { $percentoritems = $true }
                elseif ($IsItems) { $percentoritems = $false }
                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsInTopRange($False, $BottomRankValue, $percentoritems, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsInTopRange($False, $BottomRankValue, $percentoritems, $SLStyle) | Out-Null }
            }

            ## // Blank Cells

            if ($PSCmdlet.ParameterSetName -eq 'Blank')
            {
                if ($HighlightBlankCells)
                {
                    if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsContainingBlanks($True, $PresetStyle) | Out-Null }
                    elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsContainingBlanks($True, $SLStyle) | Out-Null }
                }
            }

            if ($PSCmdlet.ParameterSetName -eq 'NonBlank')
            {
                if ($HighlightNonBlankCells)
                {
                    if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsContainingBlanks($False, $PresetStyle) | Out-Null }
                    elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsContainingBlanks($False, $SLStyle) | Out-Null }
                }
            }

            ## // Error Cells

            if ($PSCmdlet.ParameterSetName -eq 'Error')
            {
                if ($HighlightErrorCells)
                {
                    if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsContainingErrors($True, $PresetStyle) | Out-Null }
                    elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsContainingErrors($True, $SLStyle) | Out-Null }
                }
            }

            if ($PSCmdlet.ParameterSetName -eq 'NonError')
            {
                if ($HighlightNonErrorCells)
                {
                    if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsContainingErrors($False, $PresetStyle) | Out-Null }
                    elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsContainingErrors($False, $SLStyle) | Out-Null }
                }
            }

            ## // Equal to

            if ($PSCmdlet.ParameterSetName -eq 'EqualTo')
            {

                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsEqual($True, $EqualToValue, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsEqual($True, $EqualToValue, $SLStyle) | Out-Null }
            }

            if ($PSCmdlet.ParameterSetName -eq 'NotEqualTo')
            {

                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsEqual($False, $NotEqualToValue, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsEqual($False, $NotEqualToValue, $SLStyle) | Out-Null }
            }

            ## // Text that contains

            if ($PSCmdlet.ParameterSetName -eq 'TextContains')
            {

                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsContainingText($True, $TextContainsString, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsContainingText($True, $TextContainsString, $SLStyle) | Out-Null }
            }

            if ($PSCmdlet.ParameterSetName -eq 'TextDoesNotContain')
            {

                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsContainingText($False, $TextDoesNotContainString, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsContainingText($False, $TextDoesNotContainString, $SLStyle) | Out-Null }
            }

            ## // Text that ENDs with

            if ($PSCmdlet.ParameterSetName -eq 'TextENDsWith')
            {

                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsENDingWith($TextENDsWithString, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsENDingWith($TextENDsWithString, $SLStyle) | Out-Null }
            }

            ## // Text that BEGINs with

            if ($PSCmdlet.ParameterSetName -eq 'TextBEGINsWith')
            {

                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsBEGINningWith($TextBEGINsWithString, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsBEGINningWith($TextBEGINsWithString, $SLStyle) | Out-Null }
            }

            ## // Average

            if ($PSCmdlet.ParameterSetName -eq 'Average')
            {

                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsAboveAverage([SpreadsheetLight.SLHighlightCellsAboveAverageValues]::$AverageType, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsAboveAverage([SpreadsheetLight.SLHighlightCellsAboveAverageValues]::$AverageType, $SLStyle) | Out-Null }
            }

            ## // DUplicates

            if ($PSCmdlet.ParameterSetName -eq 'Duplicate')
            {
                if ($HighLightDuplicateValues)
                {
                    if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsWithDuplicates($PresetStyle) | Out-Null }
                    elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsWithDuplicates($SLStyle) | Out-Null }
                }
            }

            ## // Unique

            if ($PSCmdlet.ParameterSetName -eq 'Unique')
            {
                if ($HighLightUniqueValues)
                {
                    if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsWithUniques($PresetStyle) | Out-Null }
                    elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsWithUniques($SLStyle) | Out-Null }
                }
            }

            ## // Cells with Formula

            if ($PSCmdlet.ParameterSetName -eq 'Formula')
            {

                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsWithFormula($FormulaString, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsWithFormula($FormulaString, $SLStyle) | Out-Null }
            }

            ## // Dates

            if ($PSCmdlet.ParameterSetName -eq 'Date')
            {

                if ($StyleType -eq 'PresetStyle') { $ConditionalFormatting.HighlightCellsWithDatesOccurring($DateString, $PresetStyle) | Out-Null }
                elseif ($StyleType -eq 'CustomStyle') { $ConditionalFormatting.HighlightCellsWithDatesOccurring($DateString, $SLStyle) | Out-Null }
            }

            Write-Verbose ("Set-SLConditionalFormatIconSet :`t Applying conditional formatting IconSet on Range '{0}'" -f $Range)
            $WorkBookInstance.AddConditionalFormatting($ConditionalFormatting) | Out-Null

            $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }#select-slworksheet

    }#process
    END
    {
    }
}
