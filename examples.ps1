Set-Location $PSScriptRoot
Import-Module .\PSspreadsheetlight.psd1 -Verbose -Force

$path = 'C:\gh\PSspreadsheetlight\samples'
#$doc = New-SLDocument -WorkbookName MyFirstDoc -WorksheetName service -Path $path -Verbose -PassThru -Force
$doc = Get-SLDocument -Path $path\MyFirstDoc.xlsx

Get-Service b* | Select-Object name, status, displayname |
    Export-SLDocument -WorkBookInstance $doc -WorksheetName service -ClearWorksheet


Get-ChildItem c:\temp | Select-Object name, FullName, Length |
    Export-SLDocument -WorkBookInstance $doc -WorksheetName service -ClearWorksheet





$c16
$c17

$st = Get-SLCellStyle -WorkBookInstance $doc -WorksheetName service -CellReference c16

Set-SLCellFormat -WorkBookInstance $doc -WorksheetName service -CellReference c17 -FormatString '0.00'
Save-SLDocument $doc

$st = Get-SLCellStyle -WorkBookInstance $doc -WorksheetName service -CellReference c17






$ModuleName = 'PSspreadsheetlight' #Specify the module name
$SplittedFunctionPath = 'C:\gh\PSspreadsheetlight\Public' #Specify Splitted function path



#Function to split the module and export it as functions
Function Insert-Content
{
    param ( [String]$Path )
    process
    {
        $( , $_; Get-Content $Path -ea SilentlyContinue) | Out-File $Path
    }
}

$FunctionName = Get-Command -Module $ModuleName
$path = $SplittedFunctionPath
Foreach ($Function in $FunctionName.Name)
{
    $def = (Get-Command $Function).definition
    @"
Function $Function  {
$def
}
"@ | set-Content $Path\$Function.ps1


}