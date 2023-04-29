# convert excel column number to name
Function Convert-ToExcelColumnName {
    [CmdletBinding(Defaultparametersetname='index')]
    Param(
        [parameter(Mandatory=$true,Position=0,Parametersetname='index')]
        [int]$Index,
        [parameter(Mandatory=$true,Position=0,Parametersetname='CellReference')]
        [String]$CellReference
    )

    if($PSCmdlet.ParameterSetName -eq 'index')
    {
        [SpreadsheetLight.SLDocument]::WhatIsColumnName($Index)
    }

    if($PSCmdlet.ParameterSetName -eq 'CellReference')
    {
        [regex]::Match($CellReference,'[a-zA-Z]+') | select -ExpandProperty value
    }
}