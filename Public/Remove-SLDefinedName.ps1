Function Remove-SLDefinedName  {


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
