Attribute VB_Name = "EditProject"
Public Sub EditProject_CommandButton()
Dim CallForm As Integer, projectRow As Integer, t As Integer
Dim rg As Range, cell As Range
Dim pass As String
' find project row from lookup and match
Set Search = ActiveWorkbook.Sheets(2).Range("D:D").Find(What:=CStr(FeeReport.SearchComboBox.Value)) ', LookAt:=xlWhole, MatchCase:=False, SearchFormat:=False)
If Search Is Nothing Then
    MsgBox "Cannot complete edits," & vbNewLine & "project not found in database"
    Exit Sub
Else
    projectRow = Search.Row
End If
' input project info, total & LF phase & service fees
ActiveWorkbook.Sheets(2).Activate ' Project Information
            Cells(projectRow, 1).Value = projectRow - 1 'column A
            Cells(projectRow, 2).Value = "20" & Left(FeeReport.JobNumberBox, 2) 'column B
            Cells(projectRow, 3).Value = FeeReport.JobNumberBox 'column C
            Cells(projectRow, 3).HorizontalAlignment = xlRight 'right justify p#
            Cells(projectRow, 4).Value = FeeReport.TitleBox 'column D
            Cells(projectRow, 5).Value = FeeReport.AgencyBox 'column E
            Cells(projectRow, 6).Value = CInt(FeeReport.LinearFeetBox) 'column F
            Cells(projectRow, 7).Value = CDbl(FeeReport.TotalFeeBox.Value) 'column G
            Cells(projectRow, 8).Value = FeeReport.TotalLFBox.Value 'column H
            Cells(projectRow, 9).Value = FeeReport.UserName_Box.Value 'column I
            Set rg = Range("A2:O" & projectRow) 'Clear all zero values
            For Each cell In rg
                If cell.Value = "0" Then cell.Clear
            Next
ActiveWorkbook.Sheets(3).Activate ' Total Fees
            Cells(projectRow, 1).Value = projectRow - 1 'column A
            Cells(projectRow, 2).Value = CDbl(FeeReport.PD_TotalBox) 'column B
            Cells(projectRow, 3).Value = CDbl(FeeReport.Design_TotalBox) 'column C
            Cells(projectRow, 4).Value = CDbl(FeeReport.PM_TotalBox) 'column D
            Cells(projectRow, 5).Value = CDbl(FeeReport.R_TotalBox) 'column E
            Cells(projectRow, 6).Value = CDbl(FeeReport.S_TotalBox) 'column F
            Cells(projectRow, 7).Value = CDbl(FeeReport.Geo_TotalBox) 'column G
            Cells(projectRow, 8).Value = CDbl(FeeReport.TC_TotalBox) 'column H
            Cells(projectRow, 9).Value = CDbl(FeeReport.Pot_TotalBox) 'column I
            Cells(projectRow, 10).Value = CDbl(FeeReport.Pot_QuantityBox) 'column J
            Cells(projectRow, 11).Value = CDbl(FeeReport.CS_TotalBox) 'column K
            Cells(projectRow, 12).Value = CDbl(FeeReport.Enve_TotalBox) 'column L
            Cells(projectRow, 13).Value = CDbl(FeeReport.AddFee1_TotalBox) 'column M
            Cells(projectRow, 14).Value = CDbl(FeeReport.AddFee2_TotalBox) 'column N
            Cells(projectRow, 15).Value = CDbl(FeeReport.AddFee3_TotalBox) 'column O
            Set rg = Range("A2:O" & projectRow) 'Clear all zero values
            For Each cell In rg
                If cell.Value = "0" Then cell.Clear
            Next
ActiveWorkbook.Sheets(4).Activate 'LF Fees
            Cells(projectRow, 1).Value = projectRow - 1 'column A
            Cells(projectRow, 2).Value = FeeReport.PD_LFBox 'column B
            Cells(projectRow, 3).Value = FeeReport.Design_LFBox.Value 'column C
            Cells(projectRow, 4).Value = FeeReport.PM_LFBox.Value 'column D
            Cells(projectRow, 5).Value = FeeReport.R_LFBox.Value 'column E
            Cells(projectRow, 6).Value = FeeReport.S_LFBox.Value 'column F
            Cells(projectRow, 7).Value = FeeReport.Geo_LFBox.Value 'column G
            Cells(projectRow, 8).Value = FeeReport.TC_LFBox.Value 'column H
            Cells(projectRow, 9).Value = FeeReport.Pot_LFBox.Value 'column I
            If FeeReport.Pot_QuantityBox.Value <> "0" Then
                Cells(projectRow, 10).Value = FeeReport.Pot_TotalBox.Value / FeeReport.Pot_QuantityBox.Value 'column J
            End If
            Cells(projectRow, 11).Value = FeeReport.CS_LFBox.Value 'column K
            Cells(projectRow, 12).Value = FeeReport.Enve_LFBox.Value 'column L
            Cells(projectRow, 13).Value = FeeReport.AddFee1_LFBox 'column M
            Cells(projectRow, 14).Value = FeeReport.AddFee2_LFBox 'column N
            Cells(projectRow, 15).Value = FeeReport.AddFee3_LFBox 'column O
            Set rg = Range("A2:O" & projectRow) 'Clear all zero values
            For Each cell In rg
                If cell.Value = "0" Then cell.Clear
            Next
' input phase & service comments
ActiveWorkbook.Sheets(5).Activate 'Comments
            Cells(projectRow, 1).Value = projectRow - 1 'column A
            Cells(projectRow, 2).Value = FeeReport.PD_TextBox.Value 'column B
            Cells(projectRow, 3).Value = FeeReport.Design_TextBox.Value 'column C
            Cells(projectRow, 4).Value = FeeReport.PM_TextBox.Value 'column D
            Cells(projectRow, 5).Value = FeeReport.R_TextBox.Value 'column E
            Cells(projectRow, 6).Value = FeeReport.S_TextBox.Value 'column F
            Cells(projectRow, 7).Value = FeeReport.Geo_TextBox.Value 'column G
            Cells(projectRow, 8).Value = FeeReport.TC_TextBox.Value 'column H
            Cells(projectRow, 9).Value = FeeReport.Pot_TextBox.Value 'column I
            Cells(projectRow, 10).Value = FeeReport.CS_TextBox.Value 'column J
            Cells(projectRow, 11).Value = FeeReport.Enve_TextBox.Value 'column K
            Cells(projectRow, 12).Value = FeeReport.AddFee1_TextBox 'column L
            Cells(projectRow, 13).Value = FeeReport.AddFee2_TextBox 'column M
            Cells(projectRow, 14).Value = FeeReport.AddFee3_TextBox 'column N

CallForm = MsgBox("The database has been edited!" & vbCrLf _
& vbNewLine & "Do you wish to edit another project?", vbYesNo + vbQuestion)
If CallForm = vbYes Then
    ThisWorkbook.Activate
    FeeReport.UserForm_Initialize
Else
    Unload FeeReport
End If
End Sub

