Function Get-SLCellStyle  {


    <#

.SYNOPSIS
    Gets the various style settings applied to a cell.

.DESCRIPTION
    Gets the various style settings applied to a cell.The style settings can be either accessed by their name or as a property that is attached to the workbookinstance.

.PARAMETER WorkBookInstance
    Instance of an excel document that can be used for editing.

.PARAMETER WorksheetName
    Name of the Worksheet that is to be processed.

.PARAMETER CellReference
    The cell whose style settings are to be obtained.

.PARAMETER Alignment
    Display the alignment settings applied to the specified cell.

.PARAMETER Font
    Display the font settings applied to the specified cell.

.PARAMETER Fill
    Display the fill settings applied to the specified cell.

.PARAMETER Border
    Display the border settings applied to the specified cell.

.PARAMETER FormatCode
    Display the formatcode settings applied to the specified cell.

.PARAMETER Protection
    Display the protection settings applied to the specified cell.

.Example
    PS C:\> $doc = Get-SLDocument D:\ps\Excel\MyFirstDoc.xlsx
    PS C:\> $doc | Get-SLCellStyle -WorksheetName sheet6 -CellReference d6 -Alignment
    PS C:\> $doc | Get-SLCellStyle -WorksheetName sheet6 -CellReference d6 -Font
    PS C:\> $doc | Get-SLCellStyle -WorksheetName sheet6 -CellReference d6 -Fill
    PS C:\> $doc | Get-SLCellStyle -WorksheetName sheet6 -CellReference d6 -Border
    PS C:\> $doc | Get-SLCellStyle -WorksheetName sheet6 -CellReference d6 -FormatCode
    PS C:\> $doc | Get-SLCellStyle -WorksheetName sheet6 -CellReference d6 -Protection
    PS C:\> $doc | Save-SLDocument

    Description
    -----------
    Display the various style settings applied to cell d6 on sheet6.



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
        [parameter(Mandatory = $true, Position = 0, ValueFromPipeLine = $true)]
        [SpreadsheetLight.SLDocument]$WorkBookInstance,

        [Alias('CurrentWorkSheetName')]
        [parameter(Mandatory = $true, Position = 1, ValueFromPipelineByPropertyName = $true)]
        [String]$WorksheetName,

        [parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, position = 2, parametersetname = 'cell')]
        [string]$CellReference,

        [Switch]$Alignment,

        [Switch]$Font,

        [Switch]$Fill,

        [Switch]$Border,

        [Switch]$FormatCode,

        [Switch]$Protection

    )
    PROCESS
    {
        if (Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {

            if ($PSCmdlet.ParameterSetName -eq 'cell')
            {

                $SLStyle = $WorkBookInstance.GetCellStyle($CellReference)

                $StyleHash = @{

                    Alignment  = $SLStyle.Alignment
                    Protection = $SLStyle.Protection
                    FormatCode = $SLStyle.FormatCode
                    Font       = $SLStyle.Font
                    Fill       = $SLStyle.Fill
                    Border     = $SLStyle.Border
                }

                if ($Alignment)
                {
                    $Alignment_props = $SLStyle.Alignment | Get-Member -MemberType Properties | Select-Object -ExpandProperty name
                    $Assigned_AlignMent_Props = $AlignMent_Props | Where-Object { $SLStyle.Alignment.$_ -ne $null }
                    $SLStyle.Alignment | Select-Object $Assigned_AlignMent_Props

                }
                if ($Font)
                {

                    $FontHTMLColor = '#' + $SLStyle.Font.FontColor.Name
                    $SLStyle.Font | Add-Member noteproperty FontHtmlColor $FontHTMLColor -Force
                    $Font_props = $SLStyle.Font | Get-Member -MemberType Properties | Select-Object -ExpandProperty name
                    $Assigned_Font_Props = $Font_Props | Where-Object { $SLStyle.Font.$_ -ne $null }
                    $SLStyle.Font | Select-Object $Assigned_Font_Props

                }

                if ($Fill)
                {
                    $ForegroundColor = '#' + $SLStyle.Fill.PatternForegroundColor.Name
                    $BackgroundColor = '#' + $SLStyle.Fill.PatternBackgroundColor.Name
                    $SLStyle.Fill | Add-Member noteproperty ForegroundColorHTML $ForegroundColor -Force
                    $SLStyle.Fill | Add-Member noteproperty BackgroundColorHTML $BackgroundColor -Force
                    $Fill_props = $SLStyle.fill | Get-Member -MemberType Properties | Select-Object -ExpandProperty name
                    $Fill_props_Gradient = $Fill_props | Where-Object { $_ -match 'gradient' }
                    $Assigned_Fill_Gradient_Props = $Fill_props_Gradient | Where-Object { $SLStyle.Fill.$_ -ne 0 }
                    if ($Assigned_Fill_Gradient_Props)
                    {
                        $SLStyle.fill | Select-Object ForegroundColorHTML, BackgroundColorHTML, PatternType, GradientType, $Assigned_Fill_Gradient_Props
                    }
                    Else
                    {
                        $SLStyle.fill | Select-Object ForegroundColorHTML, BackgroundColorHTML, PatternType, GradientType
                    }

                }
                if ($Border)
                {
                    $LeftBorderColor = '#' + $SLStyle.Border.LeftBorder.Color.Name
                    $RightBorderColor = '#' + $SLStyle.Border.RightBorder.Color.Name
                    $TopBorderColor = '#' + $SLStyle.Border.TopBorder.Color.Name
                    $BottomBorderColor = '#' + $SLStyle.Border.BottomBorder.Color.Name
                    $DiagonalBorderColor = '#' + $SLStyle.Border.DiagonalBorder.Color.Name
                    $VerticalBorderColor = '#' + $SLStyle.Border.VerticalBorder.Color.Name
                    $HorizontalBorderColor = '#' + $SLStyle.Border.HorizontalBorder.Color.Name

                    $LeftBorderStyle = $SLStyle.Border.LeftBorder.BorderStyle
                    $RightBorderStyle = $SLStyle.Border.RightBorder.BorderStyle
                    $TopBorderStyle = $SLStyle.Border.TopBorder.BorderStyle
                    $BottomBorderStyle = $SLStyle.Border.BottomBorder.BorderStyle
                    $DiagonalBorderStyle = $SLStyle.Border.DiagonalBorder.BorderStyle
                    $VerticalBorderStyle = $SLStyle.Border.VerticalBorder.BorderStyle
                    $HorizontalBorderStyle = $SLStyle.Border.HorizontalBorder.BorderStyle


                    $SLStyle.Border | Add-Member noteproperty Left ($LeftBorderStyle , $LeftBorderColor -join ',') -Force
                    $SLStyle.Border | Add-Member noteproperty Right ($RightBorderStyle , $RightBorderColor -join ',') -Force
                    $SLStyle.Border | Add-Member noteproperty Top ($TopBorderStyle , $TopBorderColor -join ',') -Force
                    $SLStyle.Border | Add-Member noteproperty Bottom ($BottomBorderStyle , $BottomBorderColor -join ',') -Force
                    $SLStyle.Border | Add-Member noteproperty Diagonal ($DiagonalBorderStyle , $DiagonalBorderColor -join ',') -Force
                    $SLStyle.Border | Add-Member noteproperty Vertical ($VerticalBorderStyle , $VerticalBorderColor -join ',') -Force
                    $SLStyle.Border | Add-Member noteproperty Horizontal ($HorizontalBorderStyle , $HorizontalBorderColor -join ',') -Force

                    $Border_Noteprops = $SLStyle.Border | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty name
                    $Assigned_Border_Noteprops = $Border_Noteprops | Where-Object { $SLStyle.Border.$_ -notmatch 'none' }

                    #$SLStyle.border | select Left,Right,Top,Bottom,Diagonal,Vertical,Horizontal
                    $SLStyle.Border | Select-Object $Assigned_Border_Noteprops

                }
                if ($FormatCode)
                {
                    $SLStyle.FormatCode

                }
                if ($Protection)
                {
                    $SLStyle.Protection

                }

                #Write-Verbose ("Set-SLFont :`tSetting Font Style on Cell '{0}'" -f $cref)

                $WorkBookInstance | Add-Member NoteProperty CellReference $CellReference -Force
                $WorkBookInstance | Add-Member NoteProperty Style $StyleHash -Force
            }#parameterset cell

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force -PassThru
        }#select worksheet

    }#Process
    END
    {

    }

}
