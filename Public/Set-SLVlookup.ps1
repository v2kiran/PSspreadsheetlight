Function Set-SLVlookup  {


    <#

.SYNOPSIS
    Perform vlookup.Supports lookup from same or on different worksheets.

.DESCRIPTION
    Perform vlookup.Supports lookup from same or on different worksheets. The lookup worksheet(s) have to be from the same workbook.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER VRangeworksheetname
    This is the name of the worksheet that contains the lookup table.

.PARAMETER VRange
    The range that contains the lookup values eg: A1:D20.

.PARAMETER Vcellworksheetname
    This is the worksheetname that will host the vlookup formula.

.PARAMETER Vlookupcell
    This is the lookup cell reference. Example - "C10".

.PARAMETER VFormulacellRange
    This is the range containing the lookup formula. Example = "D10:D20". Note range must include cells from the same column

.PARAMETER DataColumn
    This is the datacolumn from the lookup table that contains the value(s) to be pulled.
    So if the lookup range is D1:G6 the datatable is 4 columns wide so count from 1(D) to G(4).

.Example
    PS C:\> $doc = Get-SLDocument -Path D:\ps\Excel\Vlookup.xlsx
    PS C:\> $doc | Set-SLVlookup -VRangeworksheetname OS -VRange E5:G7 -Vcellworksheetname disk -Vlookupcell A6 -VFormulacellRange H6:H11 -DataColumn 2 -Verbose
    PS C:\> $doc | Set-SLVlookup -VRangeworksheetname OS -VRange E5:G7 -Vcellworksheetname disk -Vlookupcell A6 -VFormulacellRange I6:I11 -DataColumn 3 -Verbose
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    Use vlookup to lookup a datatable(E5:G7) contained in worksheet 'disk' and dump the matching values into worksheet 'OS'.
    Note: since we are populating 2 columns H & I we need to use the vlookup cmdlet twice.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\Vlookup.xlsx  |
                Set-SLVlookup -VRangeworksheetname disk -VRange L6:M8 -Vcellworksheetname disk -Vlookupcell A6 -VFormulacellRange J6:J11 -DataColumn 2 -Verbose |
                    Save-SLDocument


    Description
    -----------
    Use vlookup to lookup and insert values in the worksheet 'disk'.


.INPUTS
   String,SpreadsheetLight.SLDocument

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    N/A

#>


    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true, Position = 1, Valuefrompipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, HelpMessage = 'This is the name of the worksheet that contains the lookup table')]
        [string]$VRangeworksheetname,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+:[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Set-SLVlookup :`tVRange should specify values in following format. Eg: A1:D10 or AB1:AD5"; break }
            })]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, HelpMessage = 'This is the lookup range. Example.. a1:c50')]
        [string]$VRange,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, HelpMessage = 'This is the worksheetname that will host the vlookup formula')]
        [string]$Vcellworksheetname,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Set-SLVlookup :`tVlookupcell should specify values in following format. Eg: A1,B10,AB5..etc"; break }
            })]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, HelpMessage = 'This is the lookup cell reference. Example - "C10"')]
        [string]$Vlookupcell,

        [ValidateScript({
                $r1, $r2 = $_ -split ':'
                $r1_match = [regex]::Match($r1, '[a-zA-Z]+') | Select-Object -ExpandProperty value
                $r2_match = [regex]::Match($r2, '[a-zA-Z]+') | Select-Object -ExpandProperty value
                if ($r1_match -eq $r2_match) { $true }
                else { $false; Write-Warning "Set-SLVlookup :`tVFormulacellRange should specify values that belong to the same column. Eg: A1:A10 or AB1:AB5"; break }
            })]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, HelpMessage = 'This is the Cell Range containing the lookup formula. Example = "D10:D20"')]
        [string]$VFormulacellRange,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, HelpMessage = 'This is the datacolumn from the lookup table that contains the value(s) to be pulled')]
        [int]$DataColumn


    )

    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $VRangeworksheetname -NoPassThru)
        {

            if ($VRangeworksheetname -ne $Vcellworksheetname)
            {
                $WorkBookInstance.SelectWorksheet($Vcellworksheetname) | Out-Null
                $nrange = Convert-ToExcelAbsoluteRange -Range $VRange -WorkSheetName $VRangeworksheetname
            }
            else
            {
                $nrange = Convert-ToExcelAbsoluteRange -Range $VRange
            }

            $r1, $r2 = $VFormulacellRange -split ':'
            $start = Convert-ToExcelRowColumnIndex -CellReference $r1 | Select-Object -ExpandProperty Row
            $END = Convert-ToExcelRowColumnIndex -CellReference $r2 | Select-Object -ExpandProperty Row
            $columnname = Convert-ToExcelColumnName -CellReference $r1
            $lookup = Convert-ToExcelColumnName -CellReference $Vlookupcell

            for ($i = $start; $i -le $END; $i++)
            {
                $cref = "$columnname$i"
                $lookup1 = "$lookup$i"

                Write-Verbose ("Set-SLVlookup :`tLookup cell '{0}', Lookup Range '{1}',Datacolumn '{2}'" -f $lookup1, $nrange, $datacolumn)
                $WorkBookInstance.SetCellValue($cref, "=vlookup($lookup1,$nrange,$datacolumn,$false)") | Out-Null
            }
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }
    }
    END
    {
    }

}
