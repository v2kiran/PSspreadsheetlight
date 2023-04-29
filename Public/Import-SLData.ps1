Function Import-SLData {

    [CmdletBinding()]
    [OutputType([SpreadsheetLight.SLDocument])]
	param (
        [parameter(Mandatory=$true,Position=0,ValueFromPipeLine=$true)]
		[SpreadsheetLight.SLDocument]$WorkBookInstance,

        [Alias('CurrentWorkSheetName')]
        [parameter(Mandatory=$true,Position=1,ValueFromPipelineByPropertyName=$true)]
		[String]$WorksheetName,

        [parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,position=2,parametersetname='cell')]
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
        if(Select-SLWorkSheet -WorkBookInstance $WorkBookInstance -WorksheetName $WorksheetName -NoPassThru)
        {

            if($PSCmdlet.ParameterSetName -eq 'cell')
            {

                    $SLStyle =  $WorkBookInstance.GetCellStyle($CellReference)

                    $StyleHash = @{

                        Alignment = $SLStyle.Alignment
                        Protection = $SLStyle.Protection
                        FormatCode = $SLStyle.FormatCode
                        Font = $SLStyle.Font
                        Fill = $SLStyle.Fill
                        Border = $SLStyle.Border }

                if($Alignment)
                {
                    $Alignment_props = $SLStyle.Alignment | gm -MemberType Properties | select -ExpandProperty name
                    $Assigned_AlignMent_Props = $AlignMent_Props | ? {$SLStyle.Alignment.$_ -ne $null }
                    $SLStyle.Alignment | select $Assigned_AlignMent_Props

                }
                if($Font)
                {

                    $FontHTMLColor = '#' + $SLStyle.Font.FontColor.Name
                    $SLStyle.Font | Add-Member noteproperty FontHtmlColor $FontHTMLColor -Force
                    $Font_props = $SLStyle.Font | gm -MemberType Properties | select -ExpandProperty name
                    $Assigned_Font_Props = $Font_Props | ? {$SLStyle.Font.$_ -ne $null }
                    $SLStyle.Font | select $Assigned_Font_Props

                }

                if($Fill)
                {
                    $ForegroundColor = '#' + $SLStyle.Fill.PatternForegroundColor.Name
                    $BackgroundColor = '#' + $SLStyle.Fill.PatternBackgroundColor.Name
                    $SLStyle.Fill | Add-Member noteproperty ForegroundColorHTML $ForegroundColor -Force
                    $SLStyle.Fill | Add-Member noteproperty BackgroundColorHTML $BackgroundColor -Force
                    $Fill_props = $SLStyle.fill | gm -MemberType Properties | select -ExpandProperty name
                    $Fill_props_Gradient = $Fill_props | where {$_ -match 'gradient'}
                    $Assigned_Fill_Gradient_Props = $Fill_props_Gradient | where {$SLStyle.Fill.$_ -ne 0}
                    if($Assigned_Fill_Gradient_Props)
                    {
                        $SLStyle.fill | select ForegroundColorHTML,BackgroundColorHTML,PatternType,GradientType,$Assigned_Fill_Gradient_Props
                    }
                    Else
                    {
                        $SLStyle.fill | select ForegroundColorHTML,BackgroundColorHTML,PatternType,GradientType
                    }

                }
                if($Border)
                {
                    $LeftBorderColor = '#' + $SLStyle.Border.LeftBorder.Color.Name
                    $RightBorderColor = '#' + $SLStyle.Border.RightBorder.Color.Name
                    $TopBorderColor = '#' + $SLStyle.Border.TopBorder.Color.Name
                    $BottomBorderColor = '#' + $SLStyle.Border.BottomBorder.Color.Name
                    $DiagonalBorderColor = '#' + $SLStyle.Border.DiagonalBorder.Color.Name
                    $VerticalBorderColor = '#' + $SLStyle.Border.VerticalBorder.Color.Name
                    $HorizontalBorderColor = '#' + $SLStyle.Border.HorizontalBorder.Color.Name

                    $LeftBorderStyle =  $SLStyle.Border.LeftBorder.BorderStyle
                    $RightBorderStyle = $SLStyle.Border.RightBorder.BorderStyle
                    $TopBorderStyle = $SLStyle.Border.TopBorder.BorderStyle
                    $BottomBorderStyle = $SLStyle.Border.BottomBorder.BorderStyle
                    $DiagonalBorderStyle = $SLStyle.Border.DiagonalBorder.BorderStyle
                    $VerticalBorderStyle = $SLStyle.Border.VerticalBorder.BorderStyle
                    $HorizontalBorderStyle = $SLStyle.Border.HorizontalBorder.BorderStyle


                    $SLStyle.Border | Add-Member noteproperty Left ($LeftBorderStyle ,$LeftBorderColor -join ',') -Force
                    $SLStyle.Border | Add-Member noteproperty Right ($RightBorderStyle ,$RightBorderColor -join ',') -Force
                    $SLStyle.Border | Add-Member noteproperty Top ($TopBorderStyle ,$TopBorderColor -join ',') -Force
                    $SLStyle.Border | Add-Member noteproperty Bottom ($BottomBorderStyle ,$BottomBorderColor -join ',') -Force
                    $SLStyle.Border | Add-Member noteproperty Diagonal ($DiagonalBorderStyle ,$DiagonalBorderColor -join ',') -Force
                    $SLStyle.Border | Add-Member noteproperty Vertical ($VerticalBorderStyle ,$VerticalBorderColor -join ',') -Force
                    $SLStyle.Border | Add-Member noteproperty Horizontal ($HorizontalBorderStyle ,$HorizontalBorderColor -join ',') -Force

                    $Border_Noteprops = $SLStyle.Border | gm -MemberType NoteProperty | select -ExpandProperty name
                    $Assigned_Border_Noteprops = $Border_Noteprops | where {$SLStyle.Border.$_ -notmatch 'none'}

                    #$SLStyle.border | select Left,Right,Top,Bottom,Diagonal,Vertical,Horizontal
                    $SLStyle.Border | select $Assigned_Border_Noteprops

                }
                if($FormatCode)
                {
                    $SLStyle.FormatCode

                }
                if($Protection)
                {
                    $SLStyle.Protection

                }

                    #Write-Verbose ("Set-SLFont :`tSetting Font Style on Cell '{0}'" -f $cref)

                $WorkBookInstance | Add-Member NoteProperty CellReference  $CellReference -Force
                $WorkBookInstance | Add-Member NoteProperty Style           $StyleHash -Force
            }#parameterset cell

            $WorkBookInstance | Add-Member NoteProperty CurrentWorksheetName $WorkBookInstance.GetCurrentWorksheetName() -Force  -PassThru
        }#select worksheet

	}#Process
    END
    {

    }
}

