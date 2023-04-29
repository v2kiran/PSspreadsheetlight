Function Copy-SLWorkSheet  {


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
