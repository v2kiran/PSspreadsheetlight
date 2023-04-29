Function Save-SLDocument  {


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
