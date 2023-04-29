Function Sort-SLData  {


    <#

.SYNOPSIS
    Sort data by row or column.

.DESCRIPTION
    Sort data by row or column.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER Range
    cellrange which needs to be sorted.

.PARAMETER ColumnNameToSortBy
    The column to be sorted Eg. A.

.PARAMETER RowIndexToSortBy
    The rowindex to be sorted Eg. 5.

.PARAMETER SortOrder
    Specify the sort order as either : 'Ascending or Descending'.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Sort-SLData -WorksheetName sheet5 -Range F4:H6 -ColumnNameToSortBy H -SortOrder ASCending  -Verbose  | Save-SLDocument


    Description
    -----------
    sort data in the range F4:H6 by column H in the Ascending order.


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
                else { $false; Write-Warning "Sort-SLData :`tRange should specify values in following format. Eg: A1:D10 or AB1:AD5"; break }
            })]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [String]$Range,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetname = 'ColumnSort')]
        [String]$ColumnNameToSortBy,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetname = 'RowSort')]
        [String]$RowIndexToSortBy,

        [ValidateSet('ASCending', 'DESCending')]
        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [String]$SortOrder

    )
    PROCESS
    {

        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($SortOrder -eq 'AscENDing') { $SortOrderbool = $true }
            Else { $SortOrderbool = $false }

            if ($PSCmdlet.ParameterSetName -eq 'ColumnSort')
            {
                Write-Verbose ("Sort-SLData :`tSorting Cellrange '{0}' by the column '{1}' in the '{2}' order" -f $Range, $ColumnNameToSortBy, $SortOrder)
                $startcellreference, $ENDcellreference = $range -split ':'
                $WorkBookInstance.sort($startcellreference, $ENDcellreference, $ColumnNameToSortBy, $SortOrderbool)

                $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
            }

            Elseif ($PSCmdlet.ParameterSetName -eq 'RowSort')
            {
                <#
                Write-Verbose ("Sort-SLData :`tSorting Cellrange '{0}' by the RowIndex '{1}' in the '{2}' order" -f $Range,$RowIndexToSortBy,$SortOrder)
                $startcellreference,$ENDcellreference = $range -split ":"
                $WorkBookInstance.sort($startcellreference,$ENDcellreference,$RowIndexToSortBy,$SortOrderbool)

                $WorkBookInstance | Add-Member NoteProperty Range $Range -Force
                #>
                Write-Warning "Sort-SLData :`tSorting by row is currently not working. Will be fixed shortly."
            }


            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru

        }#select-slworksheet

    }#process
    END
    {
    }

}
