Attribute VB_Name = "Report"
Public Sub FeeReport_CommandButton()
'On Error GoTo ErrorHandler
Dim errMsg As String, user As String
Dim errOutlook As Object
Dim errEmail As Object
Dim wbName As String
Dim rg As Range, Cl As Range
Dim fee_rg As Range, fee_cl As Range
Dim CallForm As Integer
Dim wbDate As String
Dim wbPath As Variant, ProjectYear As String, fname As String
Dim Wb As Workbook, FeeWb As Workbook
Dim FileFolderExists As Boolean
Dim Count_Phase As Integer
ActiveWorkbook.Unprotect
With WorksheetFunction
    user = Join(.Transpose(.Index(ActiveWorkbook.UserStatus, 0, 1)), vbLf)
End With
With Worksheets(2)
    .Visible = True
    .Copy
    .Visible = Flase
End With
Count_Phase = 20 'Number of rows to begin with, max with all phases & services & optional fees
ProjectYear = "20" & Left(FeeReport.JobNumberBox, 2)
dtPath = CreateObject("WScript.Shell").specialfolders("Desktop") & "\"
wbDate = Replace(Date, "/", "-")
wbName = CStr(FeeReport.JobNumberBox.Value) & " - " & CStr(FeeReport.TitleBox.Value) & "_" & wbDate
wbPath = Application.GetSaveAsFilename(InitialFileName:=wbName, fileFilter:="Excel Files (*.xlsx), *.xlsx")
'"F:\proj\" & ProjectYear & "\" & JobNumberBox & "\1 Administration\Fee Estimates and Scoping\"
'check if file name already exists
If Not Dir(wbPath & ".xlsx", vbDirectory) = vbNullString Or _
Not Dir(dtPath & wbName & ".xlsx", vbDirectory) = vbNullString Then
    wbPath = Application.InputBox("A file with that name already exists," & vbNewLine & "please choose a new name:")
End If
'save file
If wbPath <> False Then
    ActiveWorkbook.SaveAs Filename:=wbPath & ".xlsx"
Else
    Exit Sub
End If
        ActiveWorkbook.Sheets.Add
        ActiveWorkbook.Sheets(1).Name = "Design Fee Report"
        ActiveWindow.DisplayGridlines = False
        With ActiveSheet.PageSetup
            .LeftHeader = "&""-,Bold""&12Water Main Design Fees"
            .LeftFooter = "&09&Z&F"
            .RightFooter = "&09&D &T"
            .Orientation = xlLandscape
        End With
' Project Information =======================================================================================================
With ActiveWorkbook.Sheets(1)
    .Range("A1").Value = "Project Title:"
    .Range("A2").Value = "Job Number:"
    .Range("A3").Value = "Agency / Client:"
    .Range("A4").Value = "Linear Feet:"
    .Range("A1:A4").Font.Bold = True
    .Range("B1").Value = FeeReport.TitleBox
    .Range("B2").Value = FeeReport.JobNumberBox
    .Range("B3").Value = FeeReport.AgencyBox
    .Range("B4").Value = CLng(FeeReport.LinearFeetBox)
    .Range("B4").NumberFormat = "#,##0"
    .Range("B4").HorizontalAlignment = xlLeft
    .Range("A6").Value = "Phase / Service"
    .Range("B6").Value = "Per LF"
    .Range("B6").HorizontalAlignment = xlRight
    .Range("C6").Value = "Total"
    .Range("C6").HorizontalAlignment = xlRight
    .Range("E6").Value = "Comments"
    .Range("E6").Font.Bold = True
    .Range("A6:C6").Font.Bold = True
    .Range("A21").Value = "Tota Design Fee:"
    .Range("B21").Value = FeeReport.TotalLFBox.Value + CLng(FeeReport.LengthAdjLF_Box.Value)
    .Range("C21").Value = FeeReport.TotalFeeBox.Value + CLng(FeeReport.LengthAdjTotal_Box.Value)
    .Range("A21:C21").Font.Bold = True
End With
' Length Adjustment Factor =============================================================================================================
If FeeReport.LengthAdjTotal_Box.Value <> "0" Then
        With ActiveWorkbook.Sheets(1)
            .Range("A20").Value = "Length Adjustment Factor"
            .Range("B20").Value = FeeReport.LengthAdjLF_Box.Value
            .Range("C20").Value = FeeReport.LengthAdjTotal_Box.Value
            .Range("A20:C20").Font.Italic = True
        End With
    Else
        Count_Phase = Count_Phase - 1
        Range("A20").EntireRow.Delete
End If
' Optional ============================================================================================================================
If FeeReport.AddFee3_TotalBox.Value <> "0" Then
    ActiveWorkbook.Sheets(1).Range("A19").Value = "Additional Fee 3"
    ActiveWorkbook.Sheets(1).Range("B19").Value = FeeReport.AddFee3_LFBox.Value
    ActiveWorkbook.Sheets(1).Range("C19").Value = FeeReport.AddFee3_TotalBox.Value
    ActiveWorkbook.Sheets(1).Range("E19").Value = FeeReport.AddFee3_TextBox.Value
Else
    Count_Phase = Count_Phase - 1
    Range("A19").EntireRow.Delete
End If

If FeeReport.AddFee2_TotalBox.Value <> "0" Then
    ActiveWorkbook.Sheets(1).Range("A18").Value = "Additional Fee 2"
    ActiveWorkbook.Sheets(1).Range("B18").Value = FeeReport.AddFee2_LFBox.Value
    ActiveWorkbook.Sheets(1).Range("C18").Value = FeeReport.AddFee2_TotalBox.Value
    ActiveWorkbook.Sheets(1).Range("E18").Value = FeeReport.AddFee2_TextBox.Value
Else
    Count_Phase = Count_Phase - 1
    Range("A18").EntireRow.Delete
End If

If FeeReport.AddFee1_TotalBox.Value <> "0" Then
    ActiveWorkbook.Sheets(1).Range("A17").Value = "Additional Fee 1"
    ActiveWorkbook.Sheets(1).Range("B17").Value = FeeReport.AddFee1_LFBox.Value
    ActiveWorkbook.Sheets(1).Range("C17").Value = FeeReport.AddFee1_TotalBox.Value
    ActiveWorkbook.Sheets(1).Range("E17").Value = FeeReport.AddFee1_TextBox.Value
Else
    Count_Phase = Count_Phase - 1
    Range("A17").EntireRow.Delete
End If

If FeeReport.Enve_TotalBox.Value <> "0" Then
    ActiveWorkbook.Sheets(1).Range("A16").Value = "Environmental Documents"
    ActiveWorkbook.Sheets(1).Range("B16").Value = FeeReport.Enve_LFBox.Value
    ActiveWorkbook.Sheets(1).Range("C16").Value = FeeReport.Enve_TotalBox.Value
    ActiveWorkbook.Sheets(1).Range("E16").Value = FeeReport.Enve_TextBox.Value
Else
    Count_Phase = Count_Phase - 1
    Range("A16").EntireRow.Delete
End If

If FeeReport.CS_TotalBox.Value <> "0" Then
    ActiveWorkbook.Sheets(1).Range("A15").Value = "Construction Support"
    ActiveWorkbook.Sheets(1).Range("B15").Value = FeeReport.CS_LFBox.Value
    ActiveWorkbook.Sheets(1).Range("C15").Value = FeeReport.CS_TotalBox.Value
    ActiveWorkbook.Sheets(1).Range("E15").Value = FeeReport.CS_TextBox.Value
Else
    Count_Phase = Count_Phase - 1
    Range("A15").EntireRow.Delete
End If
' Services ============================================================================================================================
If FeeReport.Pot_TotalBox.Value <> "0" Then
    ActiveWorkbook.Sheets(1).Range("A14").Value = "Potholing"
    ActiveWorkbook.Sheets(1).Range("B14").Value = FeeReport.Pot_LFBox.Value
    ActiveWorkbook.Sheets(1).Range("C14").Value = FeeReport.Pot_TotalBox.Value
    ActiveWorkbook.Sheets(1).Range("E14").Value = FeeReport.Pot_TextBox.Value
Else
    Count_Phase = Count_Phase - 1
    Range("A14").EntireRow.Delete
End If

If FeeReport.TC_TotalBox.Value <> "0" Then
    ActiveWorkbook.Sheets(1).Range("A13").Value = "Traffic Control"
    ActiveWorkbook.Sheets(1).Range("B13").Value = FeeReport.TC_LFBox.Value
    ActiveWorkbook.Sheets(1).Range("C13").Value = FeeReport.TC_TotalBox.Value
    ActiveWorkbook.Sheets(1).Range("E13").Value = FeeReport.TC_TextBox.Value
Else
    Count_Phase = Count_Phase - 1
    Range("A13").EntireRow.Delete
End If

If FeeReport.Geo_TotalBox.Value <> "0" Then
    ActiveWorkbook.Sheets(1).Range("A12").Value = "Geotechnical"
    ActiveWorkbook.Sheets(1).Range("B12").Value = FeeReport.Geo_LFBox.Value
    ActiveWorkbook.Sheets(1).Range("C12").Value = FeeReport.Geo_TotalBox.Value
    ActiveWorkbook.Sheets(1).Range("E12").Value = FeeReport.Geo_TextBox.Value
Else
    Count_Phase = Count_Phase - 1
    Range("A12").EntireRow.Delete
End If

If FeeReport.S_TotalBox.Value <> "0" Then
    ActiveWorkbook.Sheets(1).Range("A11").Value = "Survey"
    ActiveWorkbook.Sheets(1).Range("B11").Value = FeeReport.S_LFBox.Value
    ActiveWorkbook.Sheets(1).Range("C11").Value = FeeReport.S_TotalBox.Value
    ActiveWorkbook.Sheets(1).Range("E11").Value = FeeReport.S_TextBox.Value
Else
    Count_Phase = Count_Phase - 1
    Range("A11").EntireRow.Delete
End If
' Design Fees =======================================================================================================
If FeeReport.R_TotalBox.Value <> "0" Then
    ActiveWorkbook.Sheets(1).Range("A10").Value = "Reimbursables"
    ActiveWorkbook.Sheets(1).Range("B10").Value = FeeReport.R_LFBox.Value
    ActiveWorkbook.Sheets(1).Range("C10").Value = FeeReport.R_TotalBox.Value
    ActiveWorkbook.Sheets(1).Range("E10").Value = FeeReport.R_TextBox.Value
Else
    Count_Phase = Count_Phase - 1
    Range("A10").EntireRow.Delete
End If

If FeeReport.PM_TotalBox.Value <> "0" Then
    ActiveWorkbook.Sheets(1).Range("A9").Value = "Project Management"
    ActiveWorkbook.Sheets(1).Range("B9").Value = FeeReport.PM_LFBox.Value
    ActiveWorkbook.Sheets(1).Range("C9").Value = FeeReport.PM_TotalBox.Value
    ActiveWorkbook.Sheets(1).Range("E9").Value = FeeReport.PM_TextBox.Value
Else
    Count_Phase = Count_Phase - 1
    Range("A9").EntireRow.Delete
End If

If FeeReport.Design_TotalBox.Value <> "0" Then
    ActiveWorkbook.Sheets(1).Range("A8").Value = "Design"
    ActiveWorkbook.Sheets(1).Range("B8").Value = FeeReport.Design_LFBox.Value
    ActiveWorkbook.Sheets(1).Range("C8").Value = FeeReport.Design_TotalBox.Value
    ActiveWorkbook.Sheets(1).Range("E8").Value = FeeReport.Design_TextBox.Value
Else
    Count_Phase = Count_Phase - 1
    Range("A8").EntireRow.Delete
End If

If FeeReport.PD_TotalBox.Value <> "0" Then
    ActiveWorkbook.Sheets(1).Range("A7").Value = "Preliminary Design"
    ActiveWorkbook.Sheets(1).Range("B7").Value = FeeReport.PD_LFBox.Value
    ActiveWorkbook.Sheets(1).Range("C7").Value = FeeReport.PD_TotalBox.Value
    ActiveWorkbook.Sheets(1).Range("E7").Value = FeeReport.PD_TextBox.Value
Else
    Count_Phase = Count_Phase - 1
    Range("A7").EntireRow.Delete
End If

' Format / Clean Up =============================================================================================================
Cnt = 0
For i = 7 To 14
    If Not IsEmpty(Cells(i, 5)) Then
        Cnt = Cnt + 1
    End If
Next i
If Cnt = 0 Then
    ActiveWorkbook.Sheets(1).Range("E6").Value = ""
End If
'format as currency
Set rg = Range("C7:C25")
    For Each Cl In rg
            Cl.NumberFormat = " $#,##0"
    Next
'format as currency
Set fee_rg = Range("B7:B25")
    For Each fee_cl In fee_rg
            fee_cl.NumberFormat = " $#,##0.00"
    Next
'format column width
ActiveWorkbook.Sheets(1).Columns("A:A").ColumnWidth = 24
ActiveWorkbook.Sheets(1).Columns("B:C").ColumnWidth = 12
                       
' Analysis & Graphics =============================================================================================================
' Pie Chart
    If FeeReport.LengthAdjOn_OptionButton.Value = True Then
        PieRange = Count_Phase - 1
    Else
        PieRange = Count_Phase
    End If
    Range("C7:C" & PieRange).Select
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlPie
    ActiveChart.SetSourceData Source:=Range("'Design Fee Report'!$C$7:$C$" & PieRange)
    ActiveChart.SeriesCollection(1).Select
    ActiveChart.SeriesCollection(1).XValues = "='Design Fee Report'!$A$7:$A$" & PieRange
    ActiveChart.ApplyLayout (1)
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes("Chart 1").IncrementLeft -350
    ActiveSheet.Shapes("Chart 1").IncrementTop 365
    ActiveSheet.Shapes("Chart 1").ScaleWidth 1#, msoFalse, msoScaleFromTopLeft
    ActiveSheet.Shapes("Chart 1").ScaleHeight 2, msoFalse, _
        msoScaleFromTopLeft
    ActiveChart.SetElement (msoElementChartTitleNone)
    ActiveChart.PlotArea.Select
    ActiveChart.ClearToMatchStyle
    ActiveChart.ChartStyle = 15
    ActiveChart.ClearToMatchStyle
    ActiveSheet.Shapes("Chart 1").Line.Visible = msoFalse
    ActiveSheet.Shapes("Chart 1").Fill.Visible = msoFalse
' Scatter Plot
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlXYScatter
    ActiveChart.SetSourceData Source:=Range("'Project Information'!$F:$G")
    ActiveSheet.ChartObjects("Chart 2").Activate
    ActiveChart.ClearToMatchStyle
    ActiveChart.ChartStyle = 7
    ActiveChart.ClearToMatchStyle
    ActiveSheet.ChartObjects("Chart 2").Activate
    ActiveSheet.Shapes("Chart 2").Line.Visible = msoFalse
    ActiveSheet.ChartObjects("Chart 2").Activate
    ActiveSheet.Shapes("Chart 2").Fill.Visible = msoFalse
    ActiveSheet.ChartObjects("Chart 2").Activate
    ActiveChart.Axes(xlValue).MajorGridlines.Select
    Selection.Delete
    ActiveSheet.ChartObjects("Chart 2").Activate
    ActiveChart.Legend.Select
    Selection.Delete
    ActiveSheet.ChartObjects("Chart 2").Activate
    ActiveChart.SeriesCollection(1).Select
    With Selection
        .MarkerStyle = 2
        .MarkerSize = 3
    End With
    Selection.MarkerStyle = 8
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).DisplayUnit = xlThousands
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).DisplayUnit = xlThousands
    ActiveChart.Axes(xlValue).DisplayUnitLabel.Select
    Selection.Delete
    ActiveSheet.ChartObjects("Chart 2").Activate
    ActiveChart.Axes(xlCategory).DisplayUnitLabel.Select
    Selection.Delete
    ActiveSheet.ChartObjects("Chart 2").Activate
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    Selection.Caption = "Length (1,000 Ft)"
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleRotated)
    Selection.Caption = "Cost ($1,000)"
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Total Cost vs. Length"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "Total Cost vs. Length"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 21).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 21).Font
        .BaselineOffset = 0
        .Bold = msoTrue
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Fill.Transparency = 0
        .Fill.Solid
        .size = 12
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    'Add a linear trendline
    ActiveSheet.ChartObjects("Chart 2").Activate
    ActiveChart.SeriesCollection(1).Select
    ActiveChart.SeriesCollection(1).Trendlines.Add
    ActiveChart.SeriesCollection(1).Trendlines(1).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
    End With
    Selection.Format.Line.Transparency = 0.75
    'Add estimate to chart in red
    ActiveChart.ChartArea.Select
    ActiveSheet.ChartObjects("Chart 2").Activate
    ActiveChart.PlotArea.Select
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(2).Name = "=""Estimate"""
    ActiveChart.SeriesCollection(2).XValues = "='Design Fee Report'!$B$4"
    ActiveChart.SeriesCollection(2).Values = "='Design Fee Report'!$C$" & (Count_Phase + 1)
    ActiveChart.SeriesCollection(2).Select
    With Selection
        .MarkerStyle = 8
        .MarkerSize = 4
        .Format.Fill.ForeColor.RGB = RGB(195, 0, 0)
        .Format.Line.Visible = msoFalse
    End With
    ActiveSheet.Shapes("Chart 2").IncrementLeft -385
    ActiveSheet.Shapes("Chart 2").IncrementTop 120
    ActiveSheet.Shapes("Chart 2").ScaleWidth 0.8, msoFalse, msoScaleFromTopLeft
    ActiveSheet.Shapes("Chart 2").ScaleHeight 0.8, msoFalse, msoScaleFromBottomRight
    Worksheets(2).Visible = False
    
' Print =====================================================================================================
Application.GoTo Reference:=Range("A1"), Scroll:=True
'Print to PDF
    If FileFolderExists = True Then
        ActiveWorkbook.Worksheets("Design Fee Report").ExportAsFixedFormat _
        Filename:=wbPath & wbName & ".pdf", _
        Type:=xlTypePDF, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
        ActiveWorkbook.Close savechanges:=True
    Else
        ActiveWorkbook.Worksheets("Design Fee Report").ExportAsFixedFormat _
        Filename:=wbPath & ".pdf", _
        Type:=xlTypePDF, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
        ActiveWorkbook.Close savechanges:=True
    End If

CallForm = MsgBox("Fee report complete!" & vbCrLf & vbNewLine & "Do you want to return to the fee calculator?", vbYesNo + vbQuestion)
    If CallForm = vbYes Then
        ThisWorkbook.Activate
        FeeReport.UserForm_Initialize
        FeeReport.FeeReportPages.Pages(1).Enabled = True
        FeeReport.FeeReportPages.Value = 1
    Else
        Unload FeeReport
    End If
'Error Handling ============================================================================================
'ExitSub:
'    'any cleanup code goes here
'    Exit Sub
'ErrorHandler:
'    If Err.Number <> 0 Then
'            errMsg = user & " has encountered an error!" & vbNewLine & vbNewLine & _
'                "Error number: " & Str(Err.Number) & vbNewLine & _
'                "Source: " & Err.Source & vbNewLine & _
'                "Description: " & Err.Description
'            ' email error to developer
'            Set errOutlook = CreateObject("Outlook.Application")
'            Set errEmail = errOutlook.CreateItem(0)
'                errEmail.Subject = "Design Fee Calculator Error"
'                errEmail.body = errMsg
'                errEmail.Recipients.Add "itaia@cannoncorp.us"
'                errEmail.send
'            MsgBox errMsg
'            Debug.Print errMsg
'            Err.Clear
'        Resume ExitSub
'        Resume
'    End If
End Sub

