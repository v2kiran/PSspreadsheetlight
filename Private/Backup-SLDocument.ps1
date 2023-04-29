Function Backup-SLDocument{

    [CmdletBinding(Defaultparametersetname='Default')]
    Param(
            [parameter(Mandatory=$true,Position=0)]
		    [SpreadsheetLight.SLDocument]$WorkBookInstance,

            [parameter(Mandatory=$true,Position=1,Parametersetname='Path')]
            [String]$Path
    )

          $DateTimeString = get-date -f  dd-MM-yyyy_hhmmss
          $DuplicateName = $WorkBookInstance.workbookname + '_' + $DateTimeString + '.xlsx'


         if($PSCmdlet.ParameterSetName -eq 'Default')
         {
              If($WorkBookInstance.path)
             {
                if(-not (Test-Path $env:temp\SLPSLib))
                {
                    Try
                    {
                        New-Item -Path $env:TEMP -Name SLPSLib -ItemType Directory -ErrorAction Stop | Out-Null
                    }
                    Catch
                    {
                        Write-Warning ("Backup-SLDocument :`tAn error occured while creating the Backup folder 'SLPSLIB' at '{0}'...{1}" -f $env:temp, $Error[0].Exception.Message)
                    }
                }
                Try
                {
                    Copy-Item $WorkBookInstance.path "$env:TEMP\SLPSLib\$DuplicateName" -ErrorAction Stop
                    Write-Verbose ("Backup-SLDocument :`tWorkbook - '{0}' is now backed up to '{1}'" -f $WorkBookInstance.workbookname,"$env:TEMP\SLPSLib\$DuplicateName")
                }
                catch
                {
                    Write-Warning ("Backup-SLDocument :`tAn error occured while copying the file...{0}" -f $Error[0].Exception.Message)
                }
             }
         }


         if($PSCmdlet.ParameterSetName -eq 'Path')
         {

             If($WorkBookInstance.path)
             {
                if(Test-Path $Path)
                {
                    $backuppath = Join-Path $Path -ChildPath $DuplicateName
                    Try
                    {
                        Copy-Item $WorkBookInstance.path $backuppath -ErrorAction Stop
                        Write-Verbose ("Backup-SLDocument :`tWorkbook - '{0}' is now backed up to '{1}'" -f $WorkBookInstance.workbookname,$backuppath)
                    }
                    catch
                    {
                        Write-Warning ("Backup-SLDocument :`tAn error occured while copying the file...{0}" -f $Error[0].Exception.Message)
                    }
                }
                Else
                {
                    Write-Warning ("Backup-SLDocument :`tCould not find Path...{0}.. Make sure the target directory is created" -f $Path)
                }
            }
         }#parameterset path
}