Attribute VB_Name = "LookUp"
Public Sub SearchCommandButton()
On Error GoTo 0
Dim Search As Range
' search engine
Set Search = ActiveWorkbook.Sheets(2).Range("D:D").Find(What:=CStr(FeeReport.SearchComboBox.Value)) ', LookAt:=xlWhole, MatchCase:=False, SearchFormat:=False)
If Search Is Nothing Then
    MsgBox "Project not found in database," & vbNewLine & "please try the drop-down list"
    Exit Sub
Else
    projectRow = Search.Row
End If
' UserForm set up code
FeeReport.TitleBox.Locked = True
FeeReport.AgencyBox.Locked = True
FeeReport.JobNumberBox.Locked = True
FeeReport.LinearFeetBox.Locked = True
FeeReport.LengthAdjOff_OptionButton = True
FeeReport.LengthAdjOn_OptionButton.Locked = True
FeeReport.LengthAdjLF_Box.Value = "0"
FeeReport.LengthAdjLF_Box.Enabled = False
FeeReport.LengthAdjTotal_Box.Value = "0"
FeeReport.LengthAdjTotal_Box.Enabled = False
FeeReport.Edit_CommandButton.Enabled = True
' Find Project Information
FeeReport.JobNumberBox.Value = ActiveWorkbook.Sheets(2).Range("C" & projectRow)
FeeReport.TitleBox.Value = ActiveWorkbook.Sheets(2).Range("D" & projectRow)
FeeReport.AgencyBox.Value = ActiveWorkbook.Sheets(2).Range("E" & projectRow)
FeeReport.LinearFeetBox.Value = ActiveWorkbook.Sheets(2).Range("F" & projectRow)
' Preliminary Design
If ActiveWorkbook.Sheets(3).Range("B" & projectRow) <> "" Then
    FeeReport.PD_LumpSumOptionButton.Value = True
    FeeReport.PD_TotalBox.Locked = True
    FeeReport.PD_TotalBox = ActiveWorkbook.Sheets(3).Range("B" & projectRow)
Else
    FeeReport.PD_NAOptionButton.Value = True
End If
' Design Phase
If ActiveWorkbook.Sheets(3).Range("C" & projectRow) <> "" Then
    FeeReport.Design_LumpSumOptionButton.Value = True
    FeeReport.Design_TotalBox.Locked = True
    FeeReport.Design_TotalBox = ActiveWorkbook.Sheets(3).Range("C" & projectRow)
Else
    FeeReport.Design_NAOptionButton.Value = True
End If
' Project Management
If ActiveWorkbook.Sheets(3).Range("D" & projectRow) <> "" Then
    FeeReport.PM_LumpSumOptionButton.Value = True
    FeeReport.PM_TotalBox.Locked = True
    FeeReport.PM_TotalBox = ActiveWorkbook.Sheets(3).Range("D" & projectRow)
Else
    FeeReport.PM_NAOptionButton.Value = True
End If
' Reimbursables
If ActiveWorkbook.Sheets(3).Range("E" & projectRow) <> "" Then
    FeeReport.R_LumpSumOptionButton.Value = True
    FeeReport.R_TotalBox.Locked = True
    FeeReport.R_TotalBox = ActiveWorkbook.Sheets(3).Range("E" & projectRow)
Else
    FeeReport.R_NAOptionButton.Value = True
End If
' Survey
If ActiveWorkbook.Sheets(3).Range("F" & projectRow) <> "" Then
    FeeReport.S_LumpSumOptionButton.Value = True
    FeeReport.S_TotalBox.Locked = True
    FeeReport.S_TotalBox = ActiveWorkbook.Sheets(3).Range("F" & projectRow)
Else
    FeeReport.S_NAOptionButton.Value = True
End If
' Geotechnical
If ActiveWorkbook.Sheets(3).Range("G" & projectRow) <> "" Then
    FeeReport.Geo_LumpSumOptionButton.Value = True
    FeeReport.Geo_TotalBox.Locked = True
    FeeReport.Geo_TotalBox = ActiveWorkbook.Sheets(3).Range("G" & projectRow)
Else
    FeeReport.Geo_NAOptionButton.Value = True
End If
' Traffic Control
If ActiveWorkbook.Sheets(3).Range("H" & projectRow) <> "" Then
    FeeReport.TC_LumpSumOptionButton.Value = True
    FeeReport.TC_TotalBox.Locked = True
    FeeReport.TC_TotalBox = ActiveWorkbook.Sheets(3).Range("H" & projectRow)
Else
    FeeReport.TC_NAOptionButton.Value = True
End If
' Potholing
If ActiveWorkbook.Sheets(3).Range("I" & projectRow) <> "" Then
    FeeReport.Pot_LumpSumOptionButton.Value = True
    FeeReport.Pot_TotalBox.Locked = True
    FeeReport.Pot_TotalBox = ActiveWorkbook.Sheets(3).Range("I" & projectRow)
Else
    FeeReport.Pot_NAOptionButton.Value = True
End If
If ActiveWorkbook.Sheets(3).Range("J" & projectRow) <> "" Then
    FeeReport.Pot_QuantityBox = ActiveWorkbook.Sheets(3).Range("J" & projectRow)
    FeeReport.Pot_QuantityBox.Enabled = True
Else
    FeeReport.Pot_QuantityBox = "0"
End If
' Construction Support
If ActiveWorkbook.Sheets(3).Range("K" & projectRow) <> "" Then
    FeeReport.CS_LumpSumOptionButton.Value = True
    FeeReport.CS_TotalBox.Locked = True
    FeeReport.CS_TotalBox = ActiveWorkbook.Sheets(3).Range("K" & projectRow)
Else
    FeeReport.CS_NAOptionButton.Value = True
End If
' Environmental Documents
If ActiveWorkbook.Sheets(3).Range("L" & projectRow) <> "" Then
    FeeReport.Enve_LumpSumOptionButton.Value = True
    FeeReport.Enve_TotalBox.Locked = True
    FeeReport.Enve_TotalBox = ActiveWorkbook.Sheets(3).Range("L" & projectRow)
Else
    FeeReport.Enve_NAOptionButton.Value = True
End If
'Additional Fee 1
'FeeReport.AddFee1_TotalBox.Locked = True
'FeeReport.AddFee1_TextBox.Locked = True
If ActiveWorkbook.Sheets(3).Range("M" & projectRow) <> "" Then
    FeeReport.AddFee1_TotalBox = ActiveWorkbook.Sheets(3).Range("M" & projectRow)
Else
    FeeReport.AddFee1_TotalBox = "0"
    FeeReport.AddFee1_LFBox = "0"
End If
'Additional Fee 2
'FeeReport.AddFee2_TotalBox.Locked = True
'FeeReport.AddFee2_TextBox.Locked = True
If ActiveWorkbook.Sheets(3).Range("N" & projectRow) <> "" Then
    FeeReport.AddFee2_TotalBox = ActiveWorkbook.Sheets(3).Range("N" & projectRow)
Else
    FeeReport.AddFee2_TotalBox = "0"
    FeeReport.AddFee2_LFBox = "0"
End If
'Additional Fee 3
'FeeReport.AddFee3_TotalBox.Locked = True
'FeeReport.AddFee3_TextBox.Locked = True
If ActiveWorkbook.Sheets(3).Range("O" & projectRow) <> "" Then
    FeeReport.AddFee3_TotalBox = ActiveWorkbook.Sheets(3).Range("O" & projectRow)
Else
    FeeReport.AddFee3_TotalBox = "0"
    FeeReport.AddFee3_LFBox = "0"
End If
' Populate comments
FeeReport.PD_TextBox = ActiveWorkbook.Sheets(5).Range("B" & projectRow)
FeeReport.Design_TextBox = ActiveWorkbook.Sheets(5).Range("C" & projectRow)
FeeReport.PM_TextBox = ActiveWorkbook.Sheets(5).Range("D" & projectRow)
FeeReport.R_TextBox = ActiveWorkbook.Sheets(5).Range("E" & projectRow)
FeeReport.S_TextBox = ActiveWorkbook.Sheets(5).Range("F" & projectRow)
FeeReport.Geo_TextBox = ActiveWorkbook.Sheets(5).Range("G" & projectRow)
FeeReport.TC_TextBox = ActiveWorkbook.Sheets(5).Range("H" & projectRow)
FeeReport.Pot_TextBox = ActiveWorkbook.Sheets(5).Range("I" & projectRow)
FeeReport.CS_TextBox = ActiveWorkbook.Sheets(5).Range("J" & projectRow)
FeeReport.Enve_TextBox = ActiveWorkbook.Sheets(5).Range("K" & projectRow)
FeeReport.AddFee1_TextBox = ActiveWorkbook.Sheets(5).Range("L" & projectRow)
FeeReport.AddFee2_TextBox = ActiveWorkbook.Sheets(5).Range("M" & projectRow)
FeeReport.AddFee3_TextBox = ActiveWorkbook.Sheets(5).Range("N" & projectRow)
'Lock the search button so the file path goes to desktop
FeeReport.SearchCommandButton.Locked = True
End Sub
