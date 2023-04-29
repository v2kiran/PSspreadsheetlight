[cmdletbinding()]
param()

Write-Verbose $PSScriptRoot


$functionFolders = @('public', 'private')
ForEach ($folder in $functionFolders)
{
    $folderPath = Join-Path -Path "$PSScriptRoot" -ChildPath $folder
    If (Test-Path -Path $folderPath)
    {
        Write-Verbose -Message "Importing from $folder"
        ForEach ($function in @(Get-ChildItem -Path $folderPath -Filter '*.ps1' ))
        {
            Write-Verbose -Message "  Importing $($function.BaseName)"
            . $($function.FullName)
        }
    }
}

Export-ModuleMember -function (Get-ChildItem -Path "$PSScriptRoot\public\*.ps1").basename