Attribute VB_Name = "AddProject"
Public Sub AddProject_CommandButton()
'On Error GoTo ErrorHandler
Dim CallForm As Integer, LastRow As Integer, t As Integer
Dim rg As Range, cell As Range
Dim pass As String
' check if titel box is empty
If FeeReport.TitleBox.Value = "" Then
    FeeReport.TitleBox.BackColor = RGB(255, 219, 179)
End If
' check if job number box is empty
If FeeReport.JobNumberBox.Value = "" Then
    FeeReport.JobNumberBox.BackColor = RGB(255, 219, 179)
End If
' check if agency box is empty
If FeeReport.AgencyBox.Value = "" Then
    FeeReport.AgencyBox.BackColor = RGB(255, 219, 179)
End If
' warning message if project info has not been added
If FeeReport.TitleBox.Value = "" Or FeeReport.JobNumberBox.Value = "" Or FeeReport.AgencyBox.Value = "" Then
    FeeReport.FeeReportPages.Value = 1
    MsgBox "Please add project information"
    Exit Sub
ElseIf Not ActiveWorkbook.Sheets(2).Range("C:C").Find(What:=CStr(FeeReport.JobNumberBox.Value), _
    LookAt:=xlWhole, MatchCase:=False, SearchFormat:=False) Is Nothing Then
    MsgBox "This project already exists in the database!" & vbNewLine & "Please add a new project."
    Exit Sub
End If
' input project info, total & LF phase & service fees
LastRow = ActiveWorkbook.Sheets(2).Range("A" & Rows.Count).End(xlUp).Row + 1
ActiveWorkbook.Sheets(2).Activate ' Project Information
            Cells(LastRow, 1).Value = LastRow - 1 'column A
            Cells(LastRow, 2).Value = "20" & Left(FeeReport.JobNumberBox, 2) 'column B
            Cells(LastRow, 3).Value = FeeReport.JobNumberBox 'column C
            Cells(LastRow, 3).HorizontalAlignment = xlRight 'right justify p#
            Cells(LastRow, 4).Value = FeeReport.TitleBox 'column D
            Cells(LastRow, 5).Value = FeeReport.AgencyBox 'column E
            Cells(LastRow, 6).Value = CInt(FeeReport.LinearFeetBox) 'column F
            Cells(LastRow, 7).Value = CDbl(FeeReport.TotalFeeBox.Value) 'column G
            Cells(LastRow, 8).Value = FeeReport.TotalLFBox.Value 'column H
            Cells(LastRow, 9).Value = FeeReport.UserName_Box.Value 'column I
            Set rg = Range("A2:O" & LastRow) 'Clear all zero values
            For Each cell In rg
                If cell.Value = "0" Then cell.Clear
            Next
ActiveWorkbook.Sheets(3).Activate ' Total Fees
            Cells(LastRow, 1).Value = LastRow - 1 'column A
            Cells(LastRow, 2).Value = CDbl(FeeReport.PD_TotalBox) 'column B
            Cells(LastRow, 3).Value = CDbl(FeeReport.Design_TotalBox) 'column C
            Cells(LastRow, 4).Value = CDbl(FeeReport.PM_TotalBox) 'column D
            Cells(LastRow, 5).Value = CDbl(FeeReport.R_TotalBox) 'column E
            Cells(LastRow, 6).Value = CDbl(FeeReport.S_TotalBox) 'column F
            Cells(LastRow, 7).Value = CDbl(FeeReport.Geo_TotalBox) 'column G
            Cells(LastRow, 8).Value = CDbl(FeeReport.TC_TotalBox) 'column H
            Cells(LastRow, 9).Value = CDbl(FeeReport.Pot_TotalBox) 'column I
            Cells(LastRow, 10).Value = CDbl(FeeReport.Pot_QuantityBox) 'column J
            Cells(LastRow, 11).Value = CDbl(FeeReport.CS_TotalBox) 'column K
            Cells(LastRow, 12).Value = CDbl(FeeReport.Enve_TotalBox) 'column L
            Cells(LastRow, 13).Value = CDbl(FeeReport.AddFee1_TotalBox) 'column M
            Cells(LastRow, 14).Value = CDbl(FeeReport.AddFee2_TotalBox) 'column N
            Cells(LastRow, 15).Value = CDbl(FeeReport.AddFee3_TotalBox) 'column O
            Set rg = Range("A2:O" & LastRow) 'Clear all zero values
            For Each cell In rg
                If cell.Value = "0" Then cell.Clear
            Next
ActiveWorkbook.Sheets(4).Activate 'LF Fees
            Cells(LastRow, 1).Value = LastRow - 1 'column A
            Cells(LastRow, 2).Value = FeeReport.PD_LFBox 'column B
            Cells(LastRow, 3).Value = FeeReport.Design_LFBox.Value 'column C
            Cells(LastRow, 4).Value = FeeReport.PM_LFBox.Value 'column D
            Cells(LastRow, 5).Value = FeeReport.R_LFBox.Value 'column E
            Cells(LastRow, 6).Value = FeeReport.S_LFBox.Value 'column F
            Cells(LastRow, 7).Value = FeeReport.Geo_LFBox.Value 'column G
            Cells(LastRow, 8).Value = FeeReport.TC_LFBox.Value 'column H
            Cells(LastRow, 9).Value = FeeReport.Pot_LFBox.Value 'column I
            If FeeReport.Pot_QuantityBox.Value <> "0" Then
                Cells(LastRow, 10).Value = FeeReport.Pot_TotalBox.Value / FeeReport.Pot_QuantityBox.Value 'column J
            End If
            Cells(LastRow, 11).Value = FeeReport.CS_LFBox.Value 'column K
            Cells(LastRow, 12).Value = FeeReport.Enve_LFBox.Value 'column L
            Cells(LastRow, 13).Value = FeeReport.AddFee1_LFBox 'column M
            Cells(LastRow, 14).Value = FeeReport.AddFee2_LFBox 'column N
            Cells(LastRow, 15).Value = FeeReport.AddFee3_LFBox 'column O
            Set rg = Range("A2:O" & LastRow) 'Clear all zero values
            For Each cell In rg
                If cell.Value = "0" Then cell.Clear
            Next
' input phase & service comments
ActiveWorkbook.Sheets(5).Activate 'Comments
            Cells(LastRow, 1).Value = LastRow - 1 'column A
            Cells(LastRow, 2).Value = FeeReport.PD_TextBox.Value 'column B
            Cells(LastRow, 3).Value = FeeReport.Design_TextBox.Value 'column C
            Cells(LastRow, 4).Value = FeeReport.PM_TextBox.Value 'column D
            Cells(LastRow, 5).Value = FeeReport.R_TextBox.Value 'column E
            Cells(LastRow, 6).Value = FeeReport.S_TextBox.Value 'column F
            Cells(LastRow, 7).Value = FeeReport.Geo_TextBox.Value 'column G
            Cells(LastRow, 8).Value = FeeReport.TC_TextBox.Value 'column H
            Cells(LastRow, 9).Value = FeeReport.Pot_TextBox.Value 'column I
            Cells(LastRow, 10).Value = FeeReport.CS_TextBox.Value 'column J
            Cells(LastRow, 11).Value = FeeReport.Enve_TextBox.Value 'column K
            Cells(LastRow, 12).Value = FeeReport.AddFee1_TextBox 'column L
            Cells(LastRow, 13).Value = FeeReport.AddFee2_TextBox 'column M
            Cells(LastRow, 14).Value = FeeReport.AddFee3_TextBox 'column N
          
CallForm = MsgBox("Your project has been added to the database!" & vbCrLf _
& vbNewLine & "Do you wish to add another project?", vbYesNo + vbQuestion)
If CallForm = vbYes Then
    ThisWorkbook.Activate
    FeeReport.UserForm_Initialize
Else
    Unload FeeReport
End If
          
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

