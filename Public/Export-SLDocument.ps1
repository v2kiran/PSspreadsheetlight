Function Export-SLDocument  {


    <#

.SYNOPSIS
    Stores data in a datatable which is the input type that excel accepts.


.DESCRIPTION
    Stores data in a datatable which is the input type that excel accepts.
    Since there may be a possibility of overwriting existing data the workbook is backed up prior to processing the command.
    Location of backup --> $Env:temp\SLPSLib

.PARAMETER InputObject
    Data in the form of rows and columns. eg. output from the cmdlet 'Get-service'

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.Please note you cannot pipe a workbookinstance to this cmdlet. Instead use named parameter.

.PARAMETER WorksheetName
    Name of the Worksheet where data will be exported. Make sure this is blank if not existing data will be overwritten.

.PARAMETER StartRowIndex
    Row number which marks the start of the data table.Default value is 5


.PARAMETER StartColumnIndex
    Column number which marks the start of the data table.Default value is 2


.PARAMETER AutofitColumns
    Autofit all columns that contain data in the selected worksheet.

.PARAMETER ParseStringData
    For the most part powershell handles dataconversion to its proper datatype but it cannot help
    when data is explicitly cast as a string which gives rise to mismatch between data and datatype.
    Eg: $a = "12" stored as a string even though the value is an integer.

    Mismatched datatypes may also result due to poorly built functions that cast everything as a string.
    parsestringdata tries to coerce these string values into their respective datatypes(Integer,Double, or datetime)
    eg: $a = "25-07-2014" is a string so excel will store this as a string but when parsestring data is used the value of $a is stored as a datetime instead of a string.

    Note: In Excel Numbers are always right aligned and strings leftaligned.

.Example
    PS C:\> $doc = Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> Get-Service | Export-SLDocument -WorkBookInstance $doc -WorksheetName MyComp_Services -StartRowIndex 3 -StartColumnIndex 2 -AutofitColumns | Save-SLDocument

    Description
    -----------
    Get-Service is piped to an instance of 'MyFirstDoc'. The output is saved to a worksheet named 'MyComp_Services'.
    Note: A new worksheet will be created in case the specified worksheet dosen't exist.

.Example
    PS C:\> $doc = Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $Service = Get-Service
    PS C:\> $Process = Get-Process
    PS C:\> $Disk = Get-WmiObject -Class Win32_LogicalDisk -ComputerName "Localhost"
    PS C:\> Export-SLDocument -inputobject $Service -WorkBookInstance $doc -WorksheetName Service -AutofitColumns
    PS C:\> Export-SLDocument -inputobject $Process -WorkBookInstance $doc -WorksheetName Process -AutofitColumns
    PS C:\> Export-SLDocument -inputobject $Disk    -WorkBookInstance $doc -WorksheetName Disk    -AutofitColumns
    PS C:\> Save-SLDocument -WorkBookInstance $doc

    Description
    -----------
    An instance of MyFirstDoc is stored in a variable named doc.
    Service,process and diskdata from the localcomputer is then exported to worksheets 'service','process' & 'Disk' respectively.



.Example
    PS C:\> Export-SLDocument -inputobject (Get-EventLog -LogName System -Newest 5 | Select InstanceID,TimeGenerated,EntryType,Message) -WorkBookInstance (New-SLDocument -WorkbookName Eventlog -Path D:\ps\excel -PassThru) -WorksheetName System -AutofitColumns | Save-SLDocument


    Description
    -----------
    A one-liner to get the newest 5 entries from the system eventlog to a new workbook named 'Eventlog'.

.Example
    PS C:\> $ServiceDoc =  New-SLDocument -WorkbookName MyComp_Services -WorksheetName Service -Path D:\ps\Excel -Verbose -PassThru
    PS C:\> $ProcessDoc =  New-SLDocument -WorkbookName MyComp_Process  -WorksheetName Process -Path D:\ps\Excel -Verbose -PassThru
    PS C:\> $DiskDoc    =  New-SLDocument -WorkbookName MyComp_Disk     -WorksheetName Disk    -Path D:\ps\Excel -Verbose -PassThru
    PS C:\> $Service = Get-Service
    PS C:\> $Process = Get-Process
    PS C:\> $Disk = Get-WmiObject -Class Win32_LogicalDisk -ComputerName "Localhost"
    PS C:\> Export-SLDocument -inputobject $Service   -WorkBookInstance $ServiceDoc -WorksheetName Service  -AutofitColumns  | Save-SLDocument
    PS C:\> Export-SLDocument -inputobject $Process   -WorkBookInstance $ProcessDoc -WorksheetName Process  -AutofitColumns  | Save-SLDocument
    PS C:\> Export-SLDocument -inputobject $Disk      -WorkBookInstance $DiskDoc    -WorksheetName Disk     -AutofitColumns  | Save-SLDocument

    Description
    -----------
    3 new documents are created.One each for service,process and disk respectively.
    Export-SLDocument is then used to export the relevant data to each fo the workbooks.
    Note: 'Passthru' parameter with New-SLDocument is required when you want to store the document instance in a variable as shown above.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $Service = Get-Service | Group-Object -Property Status -AsHashTable -AsString
    PS C:\> $Running_Svcs = $service.Running | Select Name,DisplayName,Status
    PS C:\> $Stopped_Svcs = $service.Stopped | Select Name,DisplayName,Status
    PS C:\> Export-SLDocument -inputobject $Running_Svcs -WorkBookInstance $doc -WorksheetName Service -AutofitColumns -StartRowIndex 3 -StartColumnIndex 2
    PS C:\> Export-SLDocument -inputobject $Stopped_Svcs -WorkBookInstance $doc -WorksheetName Service -AutofitColumns -StartRowIndex 3 -StartColumnIndex 6
    PS C:\> Save-SLDocument -WorkBookInstance $doc

    Description
    -----------
    Group-Object is used to group by the status property .
    Export-SLDocument is then used to export the running services starting from column number 2 to 5.
    Stopped services are then exported to columns 6-8 in the same worksheet.


.INPUTS
   Object,SpreadsheetLight.SLDocument

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    N/A

#>


    [CmdletBinding()]
    [OutputType([SpreadsheetLight.SLDocument])]
    param (

        [parameter(Mandatory = $true, Position = 1, valuefrompipeline = $true)]
        $inputobject,

        [parameter(Mandatory = $true, Position = 0)]
        [SpreadsheetLight.SLDocument]
        $WorkBookInstance,

        [parameter(Mandatory = $true, Position = 2)]
        [String]
        $WorksheetName,

        [ValidateRange(1, 50)]
        [parameter(Mandatory = $false, Position = 3)]
        [int]
        $StartRowIndex = 1,

        [ValidateRange(1, 20)]
        [parameter(Mandatory = $false, Position = 4)]
        [int]
        $StartColumnIndex = 1,

        [switch]
        $AutofitColumns = $true,

        [Switch]
        $ParseStringData,

        [Switch]$ClearWorksheet,

        [string]
        $Path


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
        Backup-SLDocument -WorkBookInstance $WorkBookInstance
        if ($WorkBookInstance.GetSheetNames() -notcontains $WorksheetName)
        {
            $WorkBookInstance.AddWorksheet($WorksheetName) | Out-Null
        }
        Else
        {
            $WorkBookInstance.SelectWorksheet($WorksheetName) | Out-Null
        }

        Write-Verbose "Export-SLDocument :`tCreating Datatable..."
        #region Create DataTable
        $dt = New-Object System.Data.DataTable
        $DataHeaders = @()
        $DateHeaders = @()

        #$DataHeaders += $Data[0] | Get-Member -MemberType Properties | select  -ExpandProperty name
        $DataHeaders += $Data[0].psobject.Properties | Select-Object -ExpandProperty name

        Write-Verbose "Export-SLDocument :`tAdding column Headers to Datatable..."
        ## Add datatable Columns
        ForEach ($d in $DataHeaders )
        {

            $DataColumn = $d
            try
            {
                $ErrorActionPreference = 'stop'
                if ([string]::IsNullOrEmpty($($data[0].$DataColumn)))
                {
                    $dt.columns.add($DataColumn, [String]) | Out-Null
                }
                else
                {
                    $Dtype = ($data[0].$DataColumn).gettype().name
                    Switch -regex  ( $Dtype )
                    {

                        'string'
                        {
                            if ( $parseStringData )
                            {
                                $ConvertedIntValue = ''
                                $ConvertedDoubleValue = ''
                                $Int = [Int]::TryParse($data[0].$DataColumn, [ref]$ConvertedIntValue)
                                $Double = [Double]::TryParse($data[0].$DataColumn, [ref]$ConvertedDoubleValue)
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

                                if ($ConvertedIntValue -ne 0 -and $ConvertedDoubleValue -ne 0 )
                                {
                                    $dt.columns.add($DataColumn, [Int]) | Out-Null
                                }
                                elseif ($ConvertedIntValue -eq 0 -and $ConvertedDoubleValue -ne 0)
                                {
                                    $dt.columns.add($DataColumn, [Double]) | Out-Null
                                }
                                elseif ($IsDateTime)
                                {
                                    $dt.columns.add($DataColumn, [DateTime]) | Out-Null
                                }

                                else
                                {
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
                if ($null -eq $Dtype)
                {
                    $dt.columns.add($DataColumn, [String]) | Out-Null
                }
                #Write-Warning $Error[0].Exception.Message
            }

        }# END foreach dataheaders

        Write-Verbose "Export-SLDocument :`tAdding Rows to Datatable..."
        ## Add datatable Rows
        for ($i = 0; $i -lt $data.count; $i++)
        {
            $row = $dt.NewRow()
            foreach ($dhead in $DataHeaders)
            {
                If ([string]::IsNullOrEmpty($Data[$i].$dhead))
                {
                    $row.Item($dhead) = [DBNull]::Value
                }
                Else
                {
                    Try
                    {
                        $ErrorActionPreference = 'Stop'
                        if ($Data[$i].$dhead.Gettype().name -match 'Intptr' )
                        {
                            $row.Item($dhead) = $Data[$i].$dhead.ToInt64()
                        }
                        Elseif ($Data[$i].$dhead.Gettype().basetype.name -eq 'array')
                        {
                            $row.Item($dhead) = $Data[$i].$dhead -join ','
                        }
                        Elseif ($Data[$i].$dhead.Gettype().name -match 'byte\[\]')
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
                        Write-Warning ("Export-SLDocument :`tAn Error Occured...{0}" -f $Error[0].Exception.Message)
                        $ErrorActionPreference = 'Continue'
                    }
                }

            }

            $dt.Rows.Add($row)
        }

        #ENDregion Create DataTable
        Write-Verbose "Export-SLDocument :`tFinsihed creating the Datatable.Loading data into excel.."

        if ($ClearWorksheet)
        {
            $WorkBookInstance.ClearCellContent()
        }

        $WorkBookInstance.ImportDataTable($StartRowIndex, $StartColumnIndex, $dt, $true ) | Out-Null
        $WorkBookInstance | Add-Member NoteProperty DataTable $dt -Force
        $dt.Dispose()

        ##  Add dateformat to the date headers
        $stats = $WorkBookInstance.GetWorksheetStatistics()
        $dhrange = @()
        $DataHeaders = $DataHeaders | ForEach-Object { $_.ToString().ToUpper() }

        $DateHeaders |
            ForEach-Object {
                $h = $_.tostring().toupper()
                $dhcolumn = [array]::IndexOf($DataHeaders, $h)
                $dh = $dhcolumn + $StartColumnIndex
                $dhrange += [SpreadsheetLight.SLConvert]::ToCellRange( ($StartRowIndex + 1 ), $dh, $stats.ENDRowIndex, $dh )

            }


        $SLStyle = $WorkBookInstance.CreateStyle()
        $SLStyle.FormatCode = 'dd/MM/yyyy h:mm:ss AM/PM'


        $dhrange |
            ForEach-Object {
                $StartCellReference, $ENDCellReference = $_ -split ':'
                $WorkBookInstance.SetCellStyle($StartCellReference, $ENDCellReference, $SLStyle) | Out-Null
            }


        ## AutoFit Columns
        if ($AutofitColumns)
        {
            $WorkBookInstance.autofitcolumn('A', 'DD')
        }

        if ($path)
        {
            $WorkBookInstance.SaveAs($path)
            Write-Verbose ("Export-SLDocument :`tDocument has been Saved to path $Path")

        }
        else
        {
            $WorkBookInstance.Save()
            Write-Verbose ("Export-SLDocument :`tDocument has been Saved to path $($WorkBookInstance.Path)")
        }

        <#
        $HeaderRange = Convert-ToExcelRange -StartRowIndex $stats.StartRowIndex -StartColumnIndex $stats.StartColumnIndex -EndRowIndex $stats.StartRowIndex -EndColumnIndex $stats.ENDColumnIndex
        $DataRange = Convert-ToExcelRange -StartRowIndex ($stats.StartRowIndex + 1) -StartColumnIndex $stats.StartColumnIndex -EndRowIndex $stats.EndRowIndex -EndColumnIndex $stats.ENDColumnIndex
        $FirstDataColumn = Convert-ToExcelRange -StartRowIndex ($stats.StartRowIndex + 1) -StartColumnIndex $stats.StartColumnIndex -EndRowIndex $stats.EndRowIndex -EndColumnIndex $stats.StartColumnIndex


        $WorkBookInstance | Add-Member NoteProperty StartRowIndex $stats.StartRowIndex -Force
        $WorkBookInstance | Add-Member NoteProperty StartColumnIndex $stats.StartColumnIndex -Force
        $WorkBookInstance | Add-Member NoteProperty EndRowIndex $stats.ENDRowIndex -Force
        $WorkBookInstance | Add-Member NoteProperty EndColumnIndex $stats.ENDColumnIndex -Force
        $WorkBookInstance | Add-Member NoteProperty HeaderRange $HeaderRange -Force
        $WorkBookInstance | Add-Member NoteProperty DataRange $DataRange -Force
        $WorkBookInstance | Add-Member NoteProperty FirstDataColumn $FirstDataColumn -Force

        $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru #>

    }
    CLEAN {
        $WorkBookInstance.Dispose()
    }

}
