Function Set-SLTableStyle  {


    <#

.SYNOPSIS
    Excel offers to style your data tables via some built-in styles. This cmdlet help the user choose a built-in table style.


.DESCRIPTION
    Excel offers to style your data tables via some built-in styles. This cmdlet helps the user choose a built-in table style.
    In order to set a tablestyle excel would need to know the startrowindex,startcolumnindex,endrowindex and endcolumnindex,
    or simply the range eg: A1:B10

    If you want to apply a style to existing data in a worksheet then you would need to obtain the table values namely
    startrowindex,startcolumnindex,endrowindex,endcolumnindex or the range and then feed those values to the parameters.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER TableStyle
    There are about 55 built-in styles to choose from, ranging from light-dark.
    While there is a way to set a table style there isnt however a method to remove an applied style.
    Use tab or intellisense to choose from a list of possible values:
    'light1','light2','light3','light4','light5','light6','light7','light8','light9','light10','light11','light12','light13','light14','light15','light16','light17'
    ,'light18','light19','light20','light21','Medium1','Medium2','Medium3','Medium4','Medium5','Medium6','Medium7','Medium8','Medium9','Medium10','Medium11','Medium12','Medium13','Medium14'
    ,'Medium15','Medium16','Medium17','Medium18','Medium19','Medium20','Medium21','Medium22','Medium23','Medium24','Medium25','Medium26','Medium27','Medium28'
    ,'Dark1','Dark2','Dark3','Dark4','Dark5','Dark6','Dark7','Dark8','Dark9','Dark10','Dark11'


.PARAMETER TotalRowFunction
    Choose from the following as one of the valid options for a totalrowfunction.
    Possible Values : 'Sum','Count','Average','Product','Maximum','Minimum','CountNumbers','StandardDeviation','Variance'
    Note: Excel 2007 does not contain the 'Product' function.

.PARAMETER TotalColumnIndex
    The data column to which the totalrowfunction has to be applied.Valid values start from 1 irrespective of the column from which the data table starts..

.PARAMETER TotalRowLabel
    The label text indicating the calculated value from the total row function Eg. 'Average Sales Revenue'.

.PARAMETER TotalRowLabelColumnIndex
    The column index to set the TotalRowLabel on Eg. 2 or 4 . values start from 1 irrespective of the column from which the data table starts.

.PARAMETER StartRowIndex
    Row number which marks the start of the data table.

.PARAMETER StartColumnIndex
    Column number which marks the start of the data table.

.PARAMETER EndRowIndex
    Row number which marks the end of the data table.

.PARAMETER EndColumnIndex
    Column number which marks the end of the data table.

.PARAMETER Range
    The range that constitutes the table data eg: A1:b10.


.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $Service = Get-Service | Select Name,DisplayName,Status
    PS C:\> Export-SLDocument -inputobject $Service -WorkBookInstance $doc -WorksheetName Sheet2 |
                        Set-SLTableStyle -WorksheetName sheet2 -TableStyle Dark10 | Save-SLDocument

    Description
    -----------
    An instance of MyFirstDoc is stored in a variable named doc.
    Service data is first exported to sheet2 and then a built-in tablestyle named 'Dark10' is set on the table.
    This can probably be condensed into a oneliner but for the sake of simplicity the activity has been broken down into several steps.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $Size = @{n='Size-GB';e={$_.size/1GB -as [INT]}}
    PS C:\> $FreeSpace = @{n='FreeSpace-GB';e={$_.FreeSpace/1GB -as [INT]}}
    PS C:\> $disk = Get-WmiObject -Class Win32_Logicaldisk | select SystemName,DeviceID,VolumeName,DriveType,$size,$FreeSpace
    PS C:\> Export-SLDocument -inputobject $Disk -WorkBookInstance $doc -WorksheetName Disk |
                    Set-SLTableStyle -WorksheetName disk -TableStyle Dark11  -TotalRowFunction Sum -TotalColumnIndex 6 | Save-SLDocument

    Description
    -----------
    Disk Information is exported to a worksheetnamed 'disk'. The 'sum' function is applied to the contents of the column 6(FreeSpace)


.Example
    PS C:\> $doc = Get-SLDocument -Path D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $Service = Get-Service
    PS C:\> $Process = Get-Process
    PS C:\> $Disk = Get-WmiObject -Class Win32_LogicalDisk -ComputerName "Localhost"
    PS C:\> Export-SLDocument -inputobject $Service -WorkBookInstance $doc -WorksheetName Service -AutofitColumns | Set-SLTableStyle -WorksheetName Service -TableStyle Medium16
    PS C:\> Export-SLDocument -inputobject $Process -WorkBookInstance $doc -WorksheetName Process -AutofitColumns | Set-SLTableStyle -WorksheetName Process -TableStyle Medium16
    PS C:\> Export-SLDocument -inputobject $Disk    -WorkBookInstance $doc -WorksheetName Disk    -AutofitColumns | Set-SLTableStyle -WorksheetName Disk    -TableStyle Medium16
    PS C:\> Save-SLDocument -WorkBookInstance $doc

    Description
    -----------
    An instance of MyFirstDoc is stored in a variable named doc.
    Service,process and disk data from the localcomputer is then exported to worksheets
    'service','process' & 'Disk' with styles Medium16,17 & 28 applied respectively.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $Service = Get-Service | Group-Object -Property Status -AsHashTable -AsString
    PS C:\> $Running_Svcs = $service.Running | Select Name,DisplayName,Status
    PS C:\> $Stopped_Svcs = $service.Stopped | Select Name,DisplayName,Status
    PS C:\> Export-SLDocument -inputobject $Running_Svcs -WorkBookInstance $doc -WorksheetName Service -AutofitColumns -StartRowIndex 3 -StartColumnIndex 2
    PS C:\> Export-SLDocument -inputobject $Stopped_Svcs -WorkBookInstance $doc -WorksheetName Service -AutofitColumns -StartRowIndex 3 -StartColumnIndex 6
    PS C:\> Save-SLDocument   -WorkBookInstance $doc
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> Set-SLTableStyle -WorkBookInstance $doc -WorksheetName Service -TableStyle Light4 -StartRowIndex 3 -StartColumnIndex 2 -EndRowIndex 67  -EndColumnIndex 4
    PS C:\> Set-SLTableStyle -WorkBookInstance $doc -WorksheetName Service -TableStyle Light3 -StartRowIndex 3 -StartColumnIndex 6 -EndRowIndex 118 -EndColumnIndex 8
    PS C:\> Save-SLDocument  -WorkBookInstance $doc

    Description
    -----------
    Get-Service data is piped to group-object to be grouped by the status property and the results are stored in a variable named 'service'.
    export the running services starting from column number 2 to 5.
    Stopped services are exported to columns 6-8 in the same worksheet.
    Since we are applying two different table styles to the same worksheet we need to manually find out the table values which is why we
    save the document open it find out the start and end values for the running and stopped service ranges and then apply our style.
    Note: We cannot pipe  Export-SLDocument to set-sltablestyle because the start and end values are calculated for the entire worksheet
    and so the same style will be applied to both tables.


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
        [parameter(Mandatory = $true, position = 0, valuefrompipeline = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [Alias('CurrentWorksheetName')]
        [parameter(Mandatory = $true, Position = 1, ValueFromPipelineByPropertyName = $true)]
        [String]$WorksheetName,

        [Validateset('light1', 'light2', 'light3', 'light4', 'light5', 'light6', 'light7', 'light8', 'light9', 'light10', 'light11', 'light12', 'light13', 'light14', 'light15', 'light16', 'light17'
            , 'light18', 'light19', 'light20', 'light21', 'Medium1', 'Medium2', 'Medium3', 'Medium4', 'Medium5', 'Medium6', 'Medium7', 'Medium8', 'Medium9', 'Medium10', 'Medium11', 'Medium12', 'Medium13', 'Medium14'
            , 'Medium15', 'Medium16', 'Medium17', 'Medium18', 'Medium19', 'Medium20', 'Medium21', 'Medium22', 'Medium23', 'Medium24', 'Medium25', 'Medium26', 'Medium27', 'Medium28'
            , 'Dark1', 'Dark2', 'Dark3', 'Dark4', 'Dark5', 'Dark6', 'Dark7', 'Dark8', 'Dark9', 'Dark10', 'Dark11')]
        [Parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true)]
        [String]$TableStyle,

        [ValidateSet('Sum', 'Count', 'Average', 'Product', 'Maximum', 'Minimum', 'CountNumbers', 'StandardDeviation', 'Variance')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Range-TotalRowFunction')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Index-TotalRowFunction')]
        [string]$TotalRowFunction,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Range-TotalRowFunction')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Index-TotalRowFunction')]
        [UInt32]$TotalColumnIndex,

        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Range-TotalRowFunction')]
        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Index-TotalRowFunction')]
        [String]$TotalRowLabel,

        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Range-TotalRowFunction')]
        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Index-TotalRowFunction')]
        [int]$TotalRowLabelColumnIndex,

        [ValidateNotNullOrEmpty()]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Index-TotalRowFunction')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Index')]
        [UInt32]$StartRowIndex,

        [ValidateNotNullOrEmpty()]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Index-TotalRowFunction')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Index')]
        [UInt32]$StartColumnIndex,

        [ValidateNotNullOrEmpty()]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Index-TotalRowFunction')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Index')]
        [UInt32]$EndRowIndex,

        [ValidateNotNullOrEmpty()]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Index-TotalRowFunction')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Index')]
        [UInt32]$EndColumnIndex,

        [ValidateNotNullOrEmpty()]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Range-TotalRowFunction')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Range')]
        [String]$Range

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            If ($PSCmdlet.ParameterSetName -eq 'Range')
            {
                $RangeValue = $Range
                $startcellreference, $endcellreference = $Range -split ':'
                $SLTable = $WorkBookInstance.CreateTable($startcellreference, $endcellreference)

                Write-Verbose ("Set-SLTableStyle :`tSetting TableStyle '{0}' on CellRange '{1}' " -f $TableStyle, $Range)
                $SLtable.SetTableStyle([SpreadsheetLight.SLTableStyleTypeValues]::$TableStyle)


            }

            If ($PSCmdlet.ParameterSetName -eq 'Index')
            {
                $RangeValue = Convert-ToExcelRange -StartRowIndex $StartRowIndex -StartColumnIndex $StartColumnIndex -EndRowIndex $EndRowIndex -EndColumnIndex $EndColumnIndex
                $SLTable = $WorkBookInstance.CreateTable($StartRowIndex, $StartColumnIndex, $ENDRowIndex, $ENDColumnIndex)

                Write-Verbose ("Set-SLTableStyle :`tSetting TableStyle '{0}' on CellRange - StartRow/StartColumn '{1}':'{2}' & EndRow/EndColumn '{3}':'{4}' " -f $TableStyle, $StartRowIndex, $StartColumnIndex, $ENDRowIndex, $ENDColumnIndex)
                $SLtable.SetTableStyle([SpreadsheetLight.SLTableStyleTypeValues]::$TableStyle)

            }

            If ($PSCmdlet.ParameterSetName -eq 'Range-TotalRowFunction')
            {
                $RangeValue = $Range
                $startcellreference, $endcellreference = $Range -split ':'
                $SLTable = $WorkBookInstance.CreateTable($startcellreference, $endcellreference)

                Write-Verbose ("Set-SLTableStyle :`tSetting TableStyle '{0}' on CellRange '{1}' " -f $TableStyle, $Range)
                $SLtable.SetTableStyle([SpreadsheetLight.SLTableStyleTypeValues]::$TableStyle)

                # Setting TotalRowFunction
                $sltable.HasTotalRow = $true;

                If ($TotalRowLabel)
                {
                    $SLTable.SetTotalRowLabel($TotalRowLabelColumnIndex, $TotalRowLabel) | Out-Null
                }

                Write-Verbose ("Set-SLTableStyle : Setting TotalRowFunction - '{0}' on worksheet '{1}' column '{2}' " -f $TotalRowFunction, $WorksheetName, $TotalColumnIndex)
                $sltable.SetTotalRowFunction($TotalColumnIndex, [spreadsheetlight.SLTotalsRowFunctionValues]::$TotalRowFunction  ) | Out-Null
            }

            If ($PSCmdlet.ParameterSetName -eq 'Index-TotalRowFunction')
            {
                $RangeValue = Convert-ToExcelRange -StartRowIndex $StartRowIndex -StartColumnIndex $StartColumnIndex -EndRowIndex $EndRowIndex -EndColumnIndex $EndColumnIndex
                $SLTable = $WorkBookInstance.CreateTable($StartRowIndex, $StartColumnIndex, $ENDRowIndex, $ENDColumnIndex)

                Write-Verbose ("Set-SLTableStyle :`tSetting TableStyle '{0}' on CellRange - StartRow/StartColumn '{1}':'{2}' & EndRow/EndColumn '{3}':'{4}' " -f $TableStyle, $StartRowIndex, $StartColumnIndex, $ENDRowIndex, $ENDColumnIndex)
                $SLtable.SetTableStyle([SpreadsheetLight.SLTableStyleTypeValues]::$TableStyle)

                $sltable.HasTotalRow = $true;

                If ($TotalRowLabel)
                {
                    $SLTable.SetTotalRowLabel($TotalRowLabelColumnIndex, $TotalRowLabel) | Out-Null
                }

                Write-Verbose ("Set-SLTableStyle : Setting TotalRowFunction - '{0}' on worksheet '{1}' column '{2}' " -f $TotalRowFunction, $WorksheetName, $TotalColumnIndex)
                $sltable.SetTotalRowFunction($TotalColumnIndex, [spreadsheetlight.SLTotalsRowFunctionValues]::$TotalRowFunction  ) | Out-Null

            }

            $WorkBookInstance.InsertTable($SLTable) | Out-Null

            $WorkBookInstance | Add-Member NoteProperty Range $RangeValue -Force
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }#select-worksheet
    }

}
