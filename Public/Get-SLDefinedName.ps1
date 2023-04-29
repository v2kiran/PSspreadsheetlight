Function Get-SLDefinedName  {


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
