Function List-SLWorkSheet  {


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
