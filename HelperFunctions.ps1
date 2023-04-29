
Function Out-DataTable1 {
    [CmdletBinding()]
    [OutputType([SpreadsheetLight.SLDocument])]
	param (

        [parameter(Mandatory=$true,Position=0,valuefrompipeline=$true)]
         $inputobject,

        [Switch]$ParseStringData


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

        Write-Verbose "Out-DataTable :`tCreating Datatable..."
        #region Create DataTable
        $dt = New-Object System.Data.DataTable
        $DataHeaders = @()
        $DateHeaders = @()

        #$DataHeaders += $Data[0] | Get-Member -MemberType Properties | select  -ExpandProperty name
        $DataHeaders  += $Data[0].psobject.Properties | select -ExpandProperty name

        Write-Verbose "Out-DataTable :`tAdding column Headers to Datatable..."
        ## Add datatable Columns
        ForEach($d in $DataHeaders )
        {

             $DataColumn = $d
             try
             {
                $ErrorActionPreference = 'stop'
                if([string]::IsNullOrEmpty($($data[0].$DataColumn)))
                {
                    $dt.columns.add($DataColumn, [String]) | Out-Null
                }
                else
                {
                    $Dtype= ($data[0].$DataColumn).gettype().name
                     Switch -regex  ( $Dtype )
                     {

                        'string'
                        {
                            if( $parseStringData )
                            {
                                 $ConvertedIntValue = ""
                                 $ConvertedDoubleValue = ""
                                 $Int    = [Int]::TryParse($data[0].$DataColumn,[ref]$ConvertedIntValue)
                                 $Double = [Double]::TryParse($data[0].$DataColumn,[ref]$ConvertedDoubleValue)
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

                                 if($ConvertedIntValue -ne 0 -and $ConvertedDoubleValue -ne 0 )
                                 {
                                    $dt.columns.add($DataColumn, [Int]) | Out-Null
                                 }
                                 elseif($ConvertedIntValue -eq 0 -and $ConvertedDoubleValue -ne 0)
                                 {
                                    $dt.columns.add($DataColumn, [Double]) | Out-Null
                                 }
                                 elseif($IsDateTime)
                                 {
                                    $dt.columns.add($DataColumn, [DateTime]) | Out-Null
                                 }

                                 else{
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
            if($Dtype -eq $null)
            {
                $dt.columns.add($DataColumn, [String]) | Out-Null
            }
            #Write-Warning $Error[0].Exception.Message
            #continue
        }

        }# END foreach dataheaders

        Write-Verbose "Out-DataTable :`tAdding Rows to Datatable..."
        ## Add datatable Rows
        for($i = 0;$i -lt $data.count; $i++)
        {
            $row = $dt.NewRow()
            foreach($dhead in $DataHeaders)
            {
                If([string]::IsNullOrEmpty($Data[$i].$dhead))
                {
                    $row.Item($dhead) = [DBNull]::Value
                }
                Else
                {
                    Try
                    {
                        $ErrorActionPreference = 'Stop'
                        if($Data[$i].$dhead.Gettype().name -match 'Intptr' )
                        {
                            $row.Item($dhead) = $Data[$i].$dhead.ToInt64()
                        }
                        Elseif($Data[$i].$dhead.Gettype().basetype.name -eq 'array')
                        {
                          $row.Item($dhead) = $Data[$i].$dhead -join ','
                        }
                        Elseif($Data[$i].$dhead.Gettype().name -match 'byte\[\]')
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
                        Write-Warning ("Out-DataTable : An Error Occured...{0}" -f $Error[0].Exception.Message)
                        $ErrorActionPreference = 'Continue'
                    }
                }

            }

            $dt.Rows.Add($row)
        }

     #ENDregion Create DataTable

     #Write the Datatable Object
     Write-Output $dt

    }
}