Function Import-SLTextFile  {


    <#

.SYNOPSIS
    Import text documents into excel.

.DESCRIPTION
    Import text documents into excel.Text delimiter may be TAB,Space,Semicolon or fixedWidth
    Note: Import does not work properly when an imported column begins with an "=" operator.
    This is because excel will try to interpret those cells as formulas instead of text.


.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet where gridlines are to be shown.

.PARAMETER TextFile
    Path to the textfile that is to be imported.Only one path can be specified at a time.

.PARAMETER Delimiter
    Text delimiter. Use tab or intellisense to select from a list of values:TAB,Space,Semicolon or fixedWidth

.PARAMETER ImportStartCell
    Begin text import at this cellreference Eg. B2.
    Default value is B2.

.PARAMETER SkipColumns
    Column ID's to be skipped.More than one column ID can be specified.Eg. 2,3
    The column ID's passed as arguements to this parameter wont be imported into excel.

.PARAMETER Culture
    The culture to be used Eg. de-DE (German)

.PARAMETER DateColumnIndex
    The Column ID's that are to be formatted as dates. More than one column ID can be specified.Eg. 2,3
    To be used along with the 'DateFormat' parameter.

.PARAMETER ImportCustomDateFormat
    To be used when the text to be imported contains date in a non-standard format.

.PARAMETER AutofitColumns
    Autofit all columns.

.PARAMETER Force
    Use force to overwrite an existing worksheet in a workbook.


.Example
    PS C:\> Get-SLDocument -Path C:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $textfile = "C:\temp\delim.txt"
    PS C:\> $doc | Import-SLTextFile -TextFile $textfile -Delimiter Tab -ImportStartCell B2 -DateColumnIndex 7 -DateFormat YYYY-MM-dd -Verbose


    Description
    -----------
    Import delim.txt into an existing document 'MyFirstDoc'. Column 7 will be formatted as a date


.Example
    PS C:\> $doc = New-SLDocument -Path C:\ps\Excel\MyFirstDoc.xlsx -Passthru
    PS C:\> $textfile = "C:\temp\delim.txt"
    PS C:\> $doc | Import-SLTextFile -TextFile $textfile -Delimiter Tab -ImportStartCell B2 -DateColumnIndex 7 -DateFormat YYYY-MM-dd -SkipColumns 3,5 -Verbose


    Description
    -----------
    Import delim.txt into a new document 'MyFirstDoc'. Column 7 will be formatted as a date
    Columns 3 and 5 will be skipped.



.INPUTS
   String,SpreadsheetLight.SLDocument

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    N/A

#>


    [CmdletBinding()]
    param (

        [parameter(Mandatory = $true, Position = 1, ValueFromPipeline = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [Alias('FullName')]
        [parameter(Mandatory = $true, Position = 0, ValueFromPipeLineByPropertyName = $true)]
        [String]$TextFile,

        [ValidateSet('Tab', 'Space', 'SemiColon', 'FixedWidth')]
        [parameter(Mandatory = $true, Position = 2)]
        [string]$Delimiter,

        [parameter(Mandatory = $false)]
        [String]$ImportStartCell = 'B2',

        [parameter(Mandatory = $false)]
        [System.UInt32[]]$SkipColumns,

        [parameter(Mandatory = $false)]
        [String]$Culture,

        [parameter(Mandatory = $false)]
        [System.UInt32[]]$DateColumnIndex,

        [parameter(Mandatory = $false)]
        [String]$DateFormat,

        [parameter(Mandatory = $false)]
        [String]$ImportCustomDateFormat,

        [switch]$AutofitColumns = $true,

        [switch]$Force

    )
    BEGIN
    {
    }
    PROCESS
    {
        $filestats = Get-Item $TextFile
        $worksheetName = $filestats.basename
        Write-Verbose ("Import-SLTextFile :`tProcessing TEXT FIle '{0}'..." -f $filestats.Name)

        $ImportOptions = New-Object Spreadsheetlight.SLTextImportOptions

        switch ($delimiter)
        {
            'Comma' { $ImportOptions.UseCommaDelimiter = $true }
            'SemiColon' { $ImportOptions.UseSemicolonDelimiter = $true }
            'Space' { $ImportOptions.UseSpaceDelimiter = $true }
            'Tab' { $ImportOptions.UseTabDelimiter = $true }
            'FixedWidth' { $ImportOptions.DataFieldType = [SpreadsheetLight.SLTextImportDataFieldTypeValues]::'FixedWidth' }
        }


        if ($SkipColumns)
        {
            foreach ($SkipColumn in $SkipColumns)
            {
                Write-Verbose ("Import-SLTextFile :`tSkipping Import of columnID '{0}'..." -f $SkipColumn)
                $ImportOptions.SkipColumn($SkipColumn)
            }
        }

        if ($Culture)
        {
            Write-Verbose ("Import-SLTextFile :`tSetting culture to '{0}'..." -f $Culture)
            $ImportOptions.Culture = New-Object System.Globalization.CultureInfo($Culture)
        }

        If ($ImportCustomDateFormat)
        {
            Write-Verbose ("Import-SLTextFile :`tAdding custom date format '{0}'..." -f $ImportCustomDateFormat)
            $ImportOptions.AddCustomDateFormat($ImportCustomDateFormat) | Out-Null
        }


        if ($WorkBookInstance.GetSheetNames() -notcontains $WorksheetName)
        {

            $ProcessFile = $true
            $WorkBookInstance.AddWorksheet($WorksheetName) | Out-Null

        }
        Else
        {
            if ($force -and $PSCmdlet.ShouldProcess($WorksheetName, 'OVERWRITE Worksheet'))
            {
                $ProcessFile = $true
                $random = Get-Random -Maximum 100
                $WorkBookInstance.AddWorksheet($random) | Out-Null
                $WorkBookInstance.DeleteWorksheet($worksheetName) | Out-Null
                $WorkBookInstance.AddWorksheet($WorksheetName) | Out-Null
                $WorkBookInstance.DeleteWorksheet($random) | Out-Null
            }
            Else
            {
                $ProcessFile = $false
                Write-Warning ("Import-SLTextFile :`tSpecified Worksheet '{0}' already Exists. Please select a different name or use the '-Force' parameter to overwrite" -f $worksheetName)
            }
        }



        If ($ProcessFile)
        {
            Write-Verbose ("Import-SLTextFile :`tImporting data into Worksheet '{0}' " -f $worksheetName)
            $WorkBookInstance.ImportText($TextFile, $ImportStartCell, $ImportOptions)

            If ($DateFormat)
            {
                $SLStyle = $WorkBookInstance.CreateStyle()
                $SLStyle.FormatCode = $DateFormat
                foreach ($DateColumn in $DateColumnIndex)
                {
                    if ($ImportStartCell)
                    {
                        $AdjustedDateColumn = ((Convert-ToExcelColumnIndex -ColumnName B3) - 1 ) + $DateColumn
                        Write-Verbose ("Import-SLTextFile :`tSetting DateFormat '{0}' on ColumnID '{1}' " -f $DateFormat, $DateColumn)
                        $WorkBookInstance.SetColumnStyle($AdjustedDateColumn, $SLStyle) | Out-Null
                    }
                    Else
                    {
                        Write-Verbose ("Import-SLTextFile :`tSetting DateFormat '{0}' on ColumnID '{1}' " -f $DateFormat, $DateColumn)
                        $WorkBookInstance.SetColumnStyle($DateColumn, $SLStyle) | Out-Null
                    }
                }
            }


            ## Add Autofit to the Table - Optional
            if ($AutofitColumns)
            {
                $WorkBookInstance.autofitcolumn('A', 'DD')
            }

            $stats = $WorkBookInstance.GetWorksheetStatistics()

            $WorkBookInstance | Add-Member NoteProperty StartRowIndex $stats.StartRowIndex -Force
            $WorkBookInstance | Add-Member NoteProperty StartColumnIndex $stats.StartColumnIndex -Force
            $WorkBookInstance | Add-Member NoteProperty EndRowIndex $stats.ENDRowIndex -Force
            $WorkBookInstance | Add-Member NoteProperty EndColumnIndex $stats.ENDColumnIndex -Force
            $WorkBookInstance | Add-Member NoteProperty WorksheetNames $WorkBookInstance.GetSheetNames() -Force
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force
            Write-Output $WorkBookInstance

        }#IF processfile

    }#process
    END
    {
    }

}
