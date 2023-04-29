Function Set-SLSplitPane  {


    <#

.SYNOPSIS
    set up Split pane.

.DESCRIPTION
    set up Split pane.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER NumberOfRows
    Number of top-most rows above the horizontal split line.

.PARAMETER NumberOfColumns
    Number of left-most columns left of the vertical split line.

.PARAMETER ShowRowColumnHeadings
    If included in the parameterlist row and column headings are shown. False otherwise.

.Example
    PS C:\> Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx  | Set-SLSplitPane -WorksheetName sheet5 -NumberOfRows 3 -NumberOfColumns 8 -ShowRowColumnHeadings -Verbose  | Save-SLDocument


    Description
    -----------
    Top-left pane is '3' Rows high and '8' columns wide. Headers shown - 'True'



.INPUTS
   String,Int,SpreadsheetLight.SLDocument

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

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true)]
        [int]$NumberOfRows,

        [parameter(Mandatory = $true, Position = 2, ValueFromPipelineByPropertyName = $true)]
        [int]$NumberOfColumns,

        [parameter(Mandatory = $false, Position = 2)]
        [Switch]$ShowRowColumnHeadings


    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {
            if ($ShowRowColumnHeadings) { $Headershown = $true } else { $Headershown = $false }
            Write-Verbose ("Set-SLSplitPane :`tTop-left pane is '{0}' Rows high and '{1}' columns wide. Headers shown - '{2}' " -f $NumberOfRows, $NumberOfColumns, $Headershown)

            $WorkBookInstance.SplitPanes($NumberOfRows, $NumberOfColumns, $Headershown) | Out-Null
            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }
    }
    END
    {
    }

}
