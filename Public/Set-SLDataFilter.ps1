Function Set-SLDataFilter  {


    <#

.SYNOPSIS
    Set Autofilter on a cellrange.

.DESCRIPTION
    Set Autofilter on a cellrange.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER Range
    cellrange which needs to be filtered.


.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Set-SLDataFilter -WorksheetName sheet5 -Range F3:H6 -Verbose  | Save-SLDocument


    Description
    -----------
    Filter data in the range F3:H6.


.INPUTS
   String,SpreadsheetLight.SLDocument

.OUTPUTS
   SpreadsheetLight.SLDocument

.Link
    N/A

#>


    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true, Position = 0, ValueFromPipeLine = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [Alias('CurrentWorkSheetName')]
        [parameter(Mandatory = $true, Position = 1, ValueFromPipeLineByPropertyName = $true)]
        [String]$WorksheetName,

        [ValidateScript({
                $MatchFound = [regex]::Match($_, '[a-zA-Z]+\d+:[a-zA-Z]+\d+') | Select-Object -ExpandProperty value
                if ($MatchFound) { $true }
                else { $false; Write-Warning "Set-SLDataFilter :`tRange should specify values in following format. Eg: A1:D10 or AB1:AD5"; break }
            })]
        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true)]
        [String]$Range


    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            Write-Verbose ("Set-SLDataFilter :`tSetting autofilter on Cellrange '{0}'. " -f $Range)
            $StartCellReference, $Endcellreference = $range -split ':'
            $WorkBookInstance.Filter($StartCellReference, $Endcellreference) | Out-Null

            $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }

    }
    END
    {
    }

}
