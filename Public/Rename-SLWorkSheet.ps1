Function Rename-SLWorkSheet  {


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
