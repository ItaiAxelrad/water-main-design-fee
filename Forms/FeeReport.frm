VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FeeReport 
   Caption         =   "Water Main Design Fee Calculator"
   ClientHeight    =   9285
   ClientLeft      =   48
   ClientTop       =   -26292
   ClientWidth     =   6312
   OleObjectBlob   =   "FeeReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FeeReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================================================================================================
' UserForm Activate/Terminate & Initialize
'=======================================================================================================================================
Public Sub UserForm_Activate()
Dim user As String
' Userform starting position in top right of application
    Me.StartUpPosition = 0
    Me.Top = Application.Top + 150
    Me.Left = Application.Left + Application.Width - Me.Width - 30
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    ActiveWindow.WindowState = xlMinimized
End Sub
' Exit UserForm
Private Sub userform_terminate()
    ActiveWindow.WindowState = xlMaximized
    'ActiveSheet.DisplayPageBreaks = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
End Sub

Public Sub UserForm_Initialize()
On Error Resume Next
Dim rng As Range
Dim MaxFee As Double
Dim c, t As Integer
Dim wbActive As String
Dim N, i, j As Long
Dim vTemp, vaItems As Variant
t = 2 ' set worksheet tab to "Project Information" tab
    wbActive = ActiveWorkbook.Name
    Workbooks(wbActive).Activate
    Worksheets(t).Activate
        'Start on first tab and disable the rest until linear feet input
        FeeReportPages.Value = 0
        FeeReportPages.Pages(1).Enabled = True
        FeeReportPages.Pages(2).Enabled = False
        FeeReportPages.Pages(3).Enabled = False
        FeeReportPages.Pages(4).Enabled = False
        FeeReportPages.Pages(5).Enabled = False
        'Begin with default average value for design
        FeeReport.PD_AverageOptionButton.Value = True
        FeeReport.Design_AverageOptionButton.Value = True
        FeeReport.PM_AverageOptionButton.Value = True
        FeeReport.R_AverageOptionButton.Value = True
        'Begin with no services selected
        FeeReport.S_NAOptionButton.Value = True
        FeeReport.Geo_NAOptionButton.Value = True
        FeeReport.Pot_NAOptionButton.Value = True
        FeeReport.TC_NAOptionButton.Value = True
        'Begin with no optional items selected and additional fees at 0
        FeeReport.CS_NAOptionButton.Value = True
        FeeReport.Enve_NAOptionButton.Value = True
        FeeReport.AddFee1_LFBox.Value = 0
        FeeReport.AddFee1_TotalBox.Value = 0
        FeeReport.AddFee1_TotalBox.Locked = False
        FeeReport.AddFee2_LFBox.Value = 0
        FeeReport.AddFee2_TotalBox.Value = 0
        FeeReport.AddFee2_TotalBox.Locked = False
        FeeReport.AddFee3_LFBox.Value = 0
        FeeReport.AddFee3_TotalBox.Value = 0
        FeeReport.AddFee3_TotalBox.Locked = False
        ' clear comments
        FeeReport.PD_TextBox = ""
        FeeReport.Design_TextBox = ""
        FeeReport.PM_TextBox = ""
        FeeReport.R_TextBox = ""
        FeeReport.S_TextBox = ""
        FeeReport.Geo_TextBox = ""
        FeeReport.TC_TextBox = ""
        FeeReport.Pot_TextBox = ""
        FeeReport.CS_TextBox = ""
        FeeReport.Enve_TextBox = ""
        FeeReport.AddFee1_TextBox = ""
        FeeReport.AddFee2_TextBox = ""
        FeeReport.AddFee3_TextBox = ""
        ' Set project info sheet to blanks
        Call TotalFeeCalc
        FeeReport.TitleBox.Locked = False
        FeeReport.AgencyBox.Locked = False
        FeeReport.JobNumberBox.Locked = False
        FeeReport.LinearFeetBox.Locked = False
        FeeReport.TitleBox = ""
        FeeReport.TitleBox.BackColor = &H80000005
        FeeReport.JobNumberBox = ""
        FeeReport.JobNumberBox.BackColor = &H80000005
        FeeReport.AgencyBox = ""
        FeeReport.AgencyBox.BackColor = &H80000005
        FeeReport.LinearFeetBox = ""
        'FeeReport.TitleBox.SetFocus
        FeeReport.LengthAdjOn_OptionButton.Locked = False
        FeeReport.LengthAdjOn_OptionButton.Value = True
        FeeReport.LengthAdjOff_OptionButton = True
        FeeReport_CommandButton.Enabled = False
        AddProject_CommandButton.Enabled = False
        Edit_CommandButton.Enabled = False
        Clear_CommandButton.Enabled = False
    ' Cover page info
    Worksheets(t).Activate
    FeeReport.PopulationLabel.Caption = Application.WorksheetFunction.Count(Columns("A"))
    FeeReport.AvgLengthLabel.Caption = Format(Round(Application.WorksheetFunction.Average(Columns("F")), 0), "#,##0")
    FeeReport.AvgFeeLabel.Caption = Format(Round(Application.WorksheetFunction.Average(Columns("G")), 0), "#,##0")
    FeeReport.AvgLFFeeLabel.Caption = Round(Application.WorksheetFunction.Average(Columns("H")), 2)
    'FeeReport.OptLabel.Caption = Format(Fee_Calc.optFactor.Value, "0.00")
    'populate search combobox
    With Sheets(t)
        N = .Cells(Rows.Count, 1).End(xlUp).Row
    End With
    SearchComboBox.Clear
    With SearchComboBox
        For i = 2 To N
            .AddItem Sheets(t).Cells(i, 4).Value
        Next i
    End With
    vaItems = SearchComboBox.List 'Put the items in a variant array
    For i = LBound(vaItems, 1) To UBound(vaItems, 1) - 1 'with VBA to sort the array
        For j = i + 1 To UBound(vaItems, 1)
            If vaItems(i, 0) > vaItems(j, 0) Then
                vTemp = vaItems(i, 0)
                vaItems(i, 0) = vaItems(j, 0)
                vaItems(j, 0) = vTemp
            End If
        Next j
    Next i
    MsgBox (vaItems & vaTemp & SearchComboBox.List)
    SearchComboBox.Clear 'Clear the listbox
    SearchComboBox.AddItem "Select Project" 'Add the sorted array back to the listbox
    For i = LBound(vaItems, 1) To UBound(vaItems, 1)
        SearchComboBox.AddItem vaItems(i, 0)
    Next i
    SearchComboBox.ListIndex = 0
    'Application.ScreenUpdating = True
    Call TotalLFBox_Change 'Set initial progress bar max fee to highest $/LF value
End Sub
'=Project info boxes background color and character limit============================================================================================
Public Sub TitleBox_Change()
    If FeeReport.TitleBox.Value = "" Then
        FeeReport.TitleBox.BackColor = RGB(255, 219, 179)
    Else
        FeeReport.TitleBox.BackColor = &H80000005
    End If
    If Len(TitleBox) > 50 Then TitleBox.Value = Left(TitleBox.Value, 50)
End Sub
Public Sub JobNumberBox_Change()
    If FeeReport.JobNumberBox.Value = "" Then
        FeeReport.JobNumberBox.BackColor = RGB(255, 219, 179)
    Else
        FeeReport.JobNumberBox.BackColor = &H80000005
    End If
    If Len(JobNumberBox) > 50 Then JobNumberBox.Value = Left(JobNumberBox.Value, 50)
End Sub
Public Sub AgencyBox_Change()
    If FeeReport.AgencyBox.Value = "" Then
        FeeReport.AgencyBox.BackColor = RGB(255, 219, 179)
    Else
        FeeReport.AgencyBox.BackColor = &H80000005
    End If
    If Len(AgencyBox) > 50 Then AgencyBox.Value = Left(AgencyBox.Value, 50)
End Sub
'=Progress Bar Controls=============================================================================================================================================
Public Sub TotalLFBox_Change()
    On Error Resume Next
    Dim MaxFee As Double
    Dim c As Integer
        Worksheets(2).Activate 'Project info tab
            MaxFee = Application.WorksheetFunction.Max(Columns("H")) 'Sub highAvg returns the largest avg based on col
            MaxFeeLabel.Caption = "$" & Round(MaxFee, 0) 'Displays largest value
            If CDbl(TotalLFBox.Value) <= (Round(MaxFee, 0) / 2) Then
                c = (255 * 2 * (TotalLFBox.Value / MaxFee))
                ProgressBarLabel.Width = (TotalLFBox.Value / MaxFee) * RangeMaxLabel.Width
                ProgressBarLabel.BackColor = RGB(c, 255, 0) 'green to yellow
            ElseIf CDbl(TotalLFBox.Value) > ((Round(MaxFee, 0) / 2)) And CDbl(TotalLFBox.Value) <= (Round(MaxFee, 0)) Then
                c = (255 * ((TotalLFBox.Value - (MaxFee / 2)) / (MaxFee / 2)))
                ProgressBarLabel.Width = (TotalLFBox.Value / MaxFee) * RangeMaxLabel.Width
                ProgressBarLabel.BackColor = RGB(255, 255 - c, 0) 'yellow to red
            Else
                ProgressBarLabel.Width = RangeMaxLabel.Width
                ProgressBarLabel.BackColor = RGB(255, 0, 0) 'full red
            End If
End Sub
'=Project Length=================================================================================================================
Public Sub LinearFeetBox_Change()
' Disable blank values and convert to integer in number format if non-zero
If FeeReport.LinearFeetBox.Value = "" Then
    FeeReport.LinearFeetBox.Value = "0"
ElseIf InStr(LinearFeetBox, ",") = 1 Then
    LinearFeetBox = Replace(LinearFeetBox, ",", "")
ElseIf FeeReport.LinearFeetBox.TextLength < 7 Then
    LinearFeetBox = Format(CLng(LinearFeetBox.Value), "#,##0")
Else
    LinearFeetBox.Value = Left(LinearFeetBox.Value, 7)
    LinearFeetBox = Format(CLng(LinearFeetBox.Value), "#,##0")
End If
' update all phase and service fees based on new project length
Call PD_TotalBox_Change
Call Design_TotalBox_Change
Call PM_TotalBox_Change
Call R_TotalBox_Change
Call S_TotalBox_Change
Call Geo_TotalBox_Change
Call TC_TotalBox_Change
Call Pot_TotalBox_Change
Call TotalFeeCalc
Call TotalLFCalc
Call LengthAdjTotal_Box_Change
' enable tabs and buttons once project length is entered
    If FeeReport.LinearFeetBox.Value <> "" And FeeReport.LinearFeetBox.Value <> "0" Then
        FeeReport_CommandButton.Enabled = True
        AddProject_CommandButton.Enabled = True
        Clear_CommandButton.Enabled = True
        FeeReportPages.Pages(2).Enabled = True
        FeeReportPages.Pages(3).Enabled = True
        FeeReportPages.Pages(4).Enabled = True
        FeeReportPages.Pages(5).Enabled = True
    End If
End Sub
Sub TotalLFCalc()
    FeeReport.TotalLFBox.Value = Round(CDbl(PD_LFBox.Value) + CDbl(Design_LFBox.Value) + CDbl(PM_LFBox.Value) + CDbl(R_LFBox.Value) _
                                + CDbl(S_LFBox.Value) + CDbl(Geo_LFBox.Value) + CDbl(Pot_LFBox.Value) + CDbl(TC_LFBox.Value), 2) _
                                + CDbl(CS_LFBox.Value) + CDbl(Enve_LFBox.Value) + CDbl(AddFee1_LFBox.Value) + CDbl(AddFee2_LFBox.Value) + CDbl(AddFee3_LFBox.Value)
End Sub
Public Sub LinearFeetBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Select Case KeyAscii
    Case 46, 48 To 57
    Case Else
        KeyAscii = 0
        'MsgBox "Only numbers allowed!"
End Select
End Sub
'=Length Adjustment Factors===============================================================================================================================
Public Sub LengthAdjOff_OptionButton_Click()
LengthAdjLF_Box.Value = "0"
LengthAdjLF_Box.Enabled = False
LengthAdjTotal_Box.Value = "0"
LengthAdjTotal_Box.Enabled = False
End Sub
Public Sub LengthAdjOn_OptionButton_Click()
LengthAdjLF_Box.Enabled = True
LengthAdjTotal_Box.Enabled = True
Call LengthAdjTotal_Box_Change
End Sub
Public Sub LengthAdjTotal_Box_Change()
On Error Resume Next
Dim AvgLength As Double
If LengthAdjOn_OptionButton = True Then
    LengthAdjLF_Box.Enabled = True
    LengthAdjTotal_Box.Enabled = True
    AvgLength = Round(WorksheetFunction.Average(ActiveWorkbook.Sheets(2).Range("F:F")), 0)
    LengthAdjTotal_Box.Value = Round(((AvgLength - (CLng(FeeReport.LinearFeetBox))) / (10 * AvgLength)) * CLng(FeeReport.TotalFeeBox.Value), 0)
    LengthAdjTotal_Box = Format(CLng(LengthAdjTotal_Box), "#,##0")
    'this needs to be fixed, only negative numbers!
    If Abs(LengthAdjTotal_Box) > Abs((0.25 * FeeReport.TotalFeeBox.Value)) Then
        LengthAdjTotal_Box.Value = Format(CLng((-0.25 * FeeReport.TotalFeeBox.Value)), "#,##0")
    End If
    FeeReport.LengthAdjLF_Box.Value = Round(FeeReport.LengthAdjTotal_Box.Value / FeeReport.LinearFeetBox, 2)
End If
End Sub
'=Total Fee=================================================================================================================================================
Public Sub TotalFeeBox_Change()
Call TotalLFCalc
Call LengthAdjTotal_Box_Change
End Sub
Public Sub TotalFeeCalc()
' Re-calculate the total design fee & convert fees to double
PD_Number = CDbl(Replace(FeeReport.PD_TotalBox.Value, ",", ""))
Design_Number = CDbl(Replace(FeeReport.Design_TotalBox.Value, ",", ""))
PM_Number = CDbl(Replace(FeeReport.PM_TotalBox.Value, ",", ""))
R_Number = CDbl(Replace(FeeReport.R_TotalBox.Value, ",", ""))
S_Number = CDbl(Replace(FeeReport.S_TotalBox.Value, ",", ""))
Geo_Number = CDbl(Replace(FeeReport.Geo_TotalBox.Value, ",", ""))
Pot_Number = CDbl(Replace(FeeReport.Pot_TotalBox.Value, ",", ""))
TC_Number = CDbl(Replace(FeeReport.TC_TotalBox.Value, ",", ""))
CS_Number = CDbl(Replace(FeeReport.CS_TotalBox.Value, ",", ""))
Enve_Number = CDbl(Replace(FeeReport.Enve_TotalBox.Value, ",", ""))
AddFee1_Number = CDbl(Replace(FeeReport.AddFee1_TotalBox.Value, ",", ""))
AddFee2_Number = CDbl(Replace(FeeReport.AddFee2_TotalBox.Value, ",", ""))
AddFee3_Number = CDbl(Replace(FeeReport.AddFee3_TotalBox.Value, ",", ""))
' Add all phase and service fees
FeeReport.TotalFeeBox.Value = PD_Number + Design_Number + PM_Number + R_Number _
                            + S_Number + Geo_Number + Pot_Number + TC_Number _
                            + CS_Number + Enve_Number + AddFee1_Number + AddFee2_Number + AddFee3_Number
' Format with comma sep. and no decimal
TotalFeeBox = Format(TotalFeeBox, "#,##0")
End Sub
'===================================================================================================================================================
' Login Page
'===================================================================================================================================================
Public Sub Login_CommandButton_Click()
    Dim uFind, pFind As Range
    ' find UserName and Password row number
    Set uFind = ActiveWorkbook.Sheets(6).Range("B:B").Find(FeeReport.UserName_Box.Value, LookAt:=xlWhole, MatchCase:=True, SearchFormat:=False)
        If Not uFind Is Nothing Then
            userRow = uFind.Row
            FeeReport.UserName_Box.BackColor = &H80000005
        Else
            FeeReport.UserName_Box.BackColor = RGB(255, 219, 179)
        End If
    Set pFind = ActiveWorkbook.Sheets(6).Range("C:C").Find(FeeReport.Password_Box.Value, LookAt:=xlWhole, MatchCase:=True, SearchFormat:=False)
        If Not pFind Is Nothing Then
            passRow = pFind.Row
            FeeReport.Password_Box.BackColor = &H80000005
        Else
            FeeReport.Password_Box.BackColor = RGB(255, 219, 179)
        End If
    ' check if username and password match
    If uFind Is Nothing Or pFind Is Nothing Then
        MsgBox "Incorrect login information," & vbNewLine & "please try again."
    Else
        If userRow = passRow Then
            FeeReportPages.Pages(1).Enabled = True
            FeeReportPages.Value = 1
        Else
            MsgBox "Incorrect login information," & vbNewLine & "please try again."
        End If
    End If
' -- Say hello to computer user (use login username instead) --
'With WorksheetFunction
'        UserName = Join(.Transpose(.Index(ActiveWorkbook.UserStatus, 0, 1)), vbLf)
'End With
'UserName = Left(user, InStr(user, " ")) ' first name only
'Application.Speech.Speak "Hello there, " & UserName
End Sub
Public Sub SignUp_CommandButton_Click()
    Dim uFind, pFind As Range
    LastRow = ActiveWorkbook.Sheets(6).Range("A" & Rows.Count).End(xlUp).Row + 1
    ' find if UserName or Password exist in database already
    Set uFind = ActiveWorkbook.Sheets(6).Range("B:B").Find(FeeReport.UserName_Box.Value, LookAt:=xlWhole, MatchCase:=True, SearchFormat:=False)
        If Not uFind Is Nothing Then
            userRow = uFind.Row
            FeeReport.UserName_Box.BackColor = &H80000005
        End If
    Set pFind = ActiveWorkbook.Sheets(6).Range("C:C").Find(FeeReport.Password_Box.Value, LookAt:=xlWhole, MatchCase:=True, SearchFormat:=False)
        If Not pFind Is Nothing Then
            passRow = pFind.Row
            FeeReport.Password_Box.BackColor = &H80000005
        End If
    ' check if username and password match
    If uFind Is Nothing Then
            ActiveWorkbook.Sheets(6).Activate ' Login Information
            Cells(LastRow, 1).Value = LastRow - 1 'column A
            Cells(LastRow, 2).Value = FeeReport.UserName_Box.Value 'column B
            Cells(LastRow, 3).Value = FeeReport.Password_Box.Value 'column C
            FeeReportPages.Pages(1).Enabled = True
            FeeReportPages.Value = 1
    Else
        MsgBox "Username already exists, please choose a differnet username."
    End If
End Sub
'===================================================================================================================================================
' Add Project
'===================================================================================================================================================
Public Sub AddProject_CommandButton_Click()
If MsgBox("Are you sure you want to add a project?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
Call AddProject.AddProject_CommandButton
End Sub
'================================================================================================================================================
' Fee Report
'================================================================================================================================================
Public Sub FeeReport_CommandButton_Click()
    Call Report.FeeReport_CommandButton
End Sub
'===============================================================================================================================================
' Project Database Lookup / Search
'===============================================================================================================================================
Public Sub SearchCommandButton_Click()
    Call LookUp.SearchCommandButton
End Sub
Public Sub SearchComboBox_Change() ' unlock search button upon searchbox change
    SearchCommandButton.Locked = False
End Sub
'Clear fees button
Public Sub Clear_CommandButton_Click()
        ThisWorkbook.Activate
        UserForm_Initialize
        FeeReportPages.Pages(1).Enabled = True
        FeeReportPages.Value = 1
        'ActiveWorkbook.Save
End Sub
'===================================================================================================================================================
' Edit Database
'===================================================================================================================================================
Public Sub Edit_CommandButton_Click()
    If MsgBox("Are you sure you want to edit this project?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    Call EditProject.EditProject_CommandButton
End Sub
'============================================================================================================================================================
' Phase Tab
'============================================================================================================================================================
'=Preliminary Design=========================================================================================================================================
Public Sub PD_AverageOptionButton_Click()
    On Error Resume Next
    Call Phase.PD_AverageOptionButton
    Call TotalFeeCalc
End Sub
Public Sub PD_HighOptionButton_Click()
    On Error Resume Next
    Call Phase.PD_HighOptionButton
    Call TotalFeeCalc
End Sub
Public Sub PD_LowOptionButton_Click()
    On Error Resume Next
    Call Phase.PD_LowOptionButton
    Call TotalFeeCalc
End Sub
Public Sub PD_LumpSumOptionButton_Click()
    On Error Resume Next
    Call Phase.PD_LumpSumOptionButton
    Call TotalFeeCalc
End Sub
Public Sub PD_NAOptionButton_Click()
    On Error Resume Next
    Call Phase.PD_NAOptionButton
    Call TotalFeeCalc
End Sub
Public Sub PD_TotalBox_Change()
    On Error GoTo 0
    Call Phase.PD_TotalBox
End Sub
Public Sub PD_TotalBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 46, 48 To 57
        Case Else
            KeyAscii = 0
            'MsgBox "Only numbers allowed!"
    End Select
End Sub
'=Design=====================================================================================================================================
Public Sub Design_AverageOptionButton_Click()
    On Error Resume Next
    Call Phase.Design_AverageOptionButton
    Call TotalFeeCalc
End Sub
Public Sub Design_HighOptionButton_Click()
    On Error Resume Next
    Call Phase.Design_HighOptionButton
    Call TotalFeeCalc
End Sub
Public Sub Design_LowOptionButton_Click()
    On Error Resume Next
    Call Phase.Design_LowOptionButton
    Call TotalFeeCalc
End Sub
Public Sub Design_LumpSumOptionButton_Click()
    On Error Resume Next
    Call Phase.Design_LumpSumOptionButton
    Call TotalFeeCalc
End Sub
Public Sub Design_NAOptionButton_Click()
    On Error Resume Next
    Call Phase.Design_NAOptionButton
    Call TotalFeeCalc
End Sub
Public Sub Design_TotalBox_Change()
    On Error GoTo 0
    Call Phase.Design_TotalBox
End Sub
Public Sub Design_TotalBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 46, 48 To 57
        Case Else
            KeyAscii = 0
            'MsgBox "Only numbers allowed!"
    End Select
End Sub
'=Project Management=====================================================================================================
Public Sub PM_AverageOptionButton_Click()
    On Error Resume Next
    Call Phase.PM_AverageOptionButton
    Call TotalFeeCalc
End Sub
Public Sub PM_HighOptionButton_Click()
    On Error Resume Next
    Call Phase.PM_HighOptionButton
    Call TotalFeeCalc
End Sub
Public Sub PM_LowOptionButton_Click()
    On Error Resume Next
    Call Phase.PM_LowOptionButton
    Call TotalFeeCalc
End Sub
Public Sub PM_LumpSumOptionButton_Click()
    On Error Resume Next
    Call Phase.PM_LumpSumOptionButton
    Call TotalFeeCalc
End Sub
Public Sub PM_NAOptionButton_Click()
    On Error Resume Next
    Call Phase.PM_NAOptionButton
    Call TotalFeeCalc
End Sub
Public Sub PM_TotalBox_Change()
    On Error GoTo 0
    Call Phase.PM_TotalBox
End Sub
Public Sub PM_TotalBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 46, 48 To 57
        Case Else
            KeyAscii = 0
            'MsgBox "Only numbers allowed!"
    End Select
End Sub
'=Reimbursables=========================================================================================================================
Public Sub R_AverageOptionButton_Click()
    On Error Resume Next
    Call Phase.R_AverageOptionButton
    Call TotalFeeCalc
End Sub
Public Sub R_HighOptionButton_Click()
    On Error Resume Next
    Call Phase.R_HighOptionButton
    Call TotalFeeCalc
End Sub
Public Sub R_LowOptionButton_Click()
    On Error Resume Next
    Call Phase.R_LowOptionButton
    Call TotalFeeCalc
End Sub
Public Sub R_LumpSumOptionButton_Click()
    On Error Resume Next
    Call Phase.R_LumpSumOptionButton
    Call TotalFeeCalc
End Sub
Public Sub R_NAOptionButton_Click()
    On Error Resume Next
    Call Phase.R_NAOptionButton
    Call TotalFeeCalc
End Sub
Public Sub R_TotalBox_Change()
    On Error GoTo 0
    Call Phase.R_TotalBox
End Sub
Public Sub R_TotalBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 46, 48 To 57
        Case Else
            KeyAscii = 0
            'MsgBox "Only numbers allowed!"
    End Select
End Sub
'============================================================================================================================================================
' Services Tab
'============================================================================================================================================================
'=Survey=========================================================================================================================================
Public Sub S_AverageOptionButton_Click()
    On Error Resume Next
    Call Service.S_AverageOptionButton
    Call TotalFeeCalc
End Sub
Public Sub S_HighOptionButton_Click()
    On Error Resume Next
    Call Service.S_HighOptionButton
    Call TotalFeeCalc
End Sub
Public Sub S_LowOptionButton_Click()
    On Error Resume Next
    Call Service.S_LowOptionButton
    Call TotalFeeCalc
End Sub
Public Sub S_LumpSumOptionButton_Click()
    On Error Resume Next
    Call Service.S_LumpSumOptionButton
    Call TotalFeeCalc
End Sub
Public Sub S_NAOptionButton_Click()
    On Error Resume Next
    Call Service.S_NAOptionButton
    Call TotalFeeCalc
End Sub
Public Sub S_TotalBox_Change()
    On Error GoTo 0
    Call Service.S_TotalBox
End Sub
Public Sub S_TotalBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 46, 48 To 57
        Case Else
            KeyAscii = 0
            'MsgBox "Only numbers allowed!"
    End Select
End Sub
'=Geotechnical==========================================================================================================================
Public Sub Geo_AverageOptionButton_Click()
    On Error Resume Next
    Call Service.Geo_AverageOptionButton
    Call TotalFeeCalc
End Sub
Public Sub Geo_HighOptionButton_Click()
    On Error Resume Next
    Call Service.Geo_HighOptionButton
    Call TotalFeeCalc
End Sub
Public Sub Geo_LowOptionButton_Click()
    On Error Resume Next
    Call Service.Geo_LowOptionButton
    Call TotalFeeCalc
End Sub
Public Sub Geo_LumpSumOptionButton_Click()
    On Error Resume Next
    Call Service.Geo_LumpSumOptionButton
    Call TotalFeeCalc
End Sub
Public Sub Geo_NAOptionButton_Click()
    On Error Resume Next
    Call Service.Geo_NAOptionButton
    Call TotalFeeCalc
End Sub
Public Sub Geo_TotalBox_Change()
    On Error GoTo 0
    Call Service.Geo_TotalBox
End Sub
Public Sub Geo_TotalBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 46, 48 To 57
        Case Else
            KeyAscii = 0
            'MsgBox "Only numbers allowed!"
    End Select
End Sub
'=Traffic Control=======================================================================================================================
Public Sub TC_AverageOptionButton_Click()
    On Error Resume Next
    Call Service.TC_AverageOptionButton
    Call TotalFeeCalc
End Sub
Public Sub TC_HighOptionButton_Click()
    On Error Resume Next
    Call Service.TC_HighOptionButton
    Call TotalFeeCalc
End Sub
Public Sub TC_LowOptionButton_Click()
    On Error Resume Next
    Call Service.TC_LowOptionButton
    Call TotalFeeCalc
End Sub
Public Sub TC_LumpSumOptionButton_Click()
    On Error Resume Next
    Call Service.TC_LumpSumOptionButton
    Call TotalFeeCalc
End Sub
Public Sub TC_NAOptionButton_Click()
    On Error Resume Next
    Call Service.TC_NAOptionButton
    Call TotalFeeCalc
End Sub
Public Sub TC_TotalBox_Change()
    On Error GoTo 0
    Call Service.TC_TotalBox
End Sub
Public Sub TC_TotalBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 46, 48 To 57
        Case Else
            KeyAscii = 0
            'MsgBox "Only numbers allowed!"
    End Select
End Sub
'=Potholing===========================================================================================================================================
Public Sub Pot_AverageOptionButton_Click()
    On Error Resume Next
    Call Service.Pot_AverageOptionButton
    Call TotalFeeCalc
End Sub
Public Sub Pot_HighOptionButton_Click()
    On Error Resume Next
    Call Service.Pot_HighOptionButton
    Call TotalFeeCalc
End Sub
Public Sub Pot_LowOptionButton_Click()
    On Error Resume Next
    Call Service.Pot_LowOptionButton
    Call TotalFeeCalc
End Sub
Public Sub Pot_LumpSumOptionButton_Click()
    On Error Resume Next
    Call Service.Pot_LumpSumOptionButton
    Call TotalFeeCalc
End Sub
Public Sub Pot_NAOptionButton_Click()
    On Error Resume Next
    Call Service.Pot_NAOptionButton
    Call TotalFeeCalc
End Sub
Public Sub Pot_QuantityOptionButton_Click()
    On Error Resume Next
    Call Service.Pot_QuantityOptionButton
    Call TotalFeeCalc
End Sub
Public Sub Pot_QuantityBox_Change()
    On Error GoTo 0
    Call Service.Pot_QuantityBox
End Sub
Public Sub Pot_TotalBox_Change()
    On Error GoTo 0
    Call Service.Pot_TotalBox
End Sub
Public Sub Pot_QuantityBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Select Case KeyAscii
    Case 46, 48 To 57
    Case Else
        KeyAscii = 0
        'MsgBox "Only numbers allowed!"
End Select
End Sub
Public Sub Pot_TotalBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 46, 48 To 57
        Case Else
            KeyAscii = 0
            'MsgBox "Only numbers allowed!"
    End Select
End Sub
'============================================================================================================================================================
' Optionals Tab
'============================================================================================================================================================
'=Construction Support=========================================================================================================================================
Public Sub CS_AverageOptionButton_Click()
    On Error Resume Next
    Call Optionals.CS_AverageOptionButton
    Call TotalFeeCalc
End Sub
Public Sub CS_HighOptionButton_Click()
    On Error Resume Next
    Call Optionals.CS_HighOptionButton
    Call TotalFeeCalc
End Sub
Public Sub CS_LowOptionButton_Click()
    On Error Resume Next
    Call Optionals.CS_LowOptionButton
    Call TotalFeeCalc
End Sub
Public Sub CS_LumpSumOptionButton_Click()
    On Error Resume Next
    Call Optionals.CS_LumpSumOptionButton
    Call TotalFeeCalc
End Sub
Public Sub CS_NAOptionButton_Click()
    On Error Resume Next
    Call Optionals.CS_NAOptionButton
    Call TotalFeeCalc
End Sub
Public Sub CS_TotalBox_Change()
    On Error GoTo 0
    Call Optionals.CS_TotalBox
End Sub
Public Sub CS_TotalBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 46, 48 To 57
        Case Else
            KeyAscii = 0
            'MsgBox "Only numbers allowed!"
    End Select
End Sub
'=Environmental Documents & Permitting========================================================================================================================
Public Sub Enve_AverageOptionButton_Click()
    On Error Resume Next
    Call Optionals.Enve_AverageOptionButton
    Call TotalFeeCalc
End Sub
Public Sub Enve_HighOptionButton_Click()
    On Error Resume Next
    Call Optionals.Enve_HighOptionButton
    Call TotalFeeCalc
End Sub
Public Sub Enve_LowOptionButton_Click()
    On Error Resume Next
    Call Optionals.Enve_LowOptionButton
    Call TotalFeeCalc
End Sub
Public Sub Enve_LumpSumOptionButton_Click()
    On Error Resume Next
    Call Optionals.Enve_LumpSumOptionButton
    Call TotalFeeCalc
End Sub
Public Sub Enve_NAOptionButton_Click()
    On Error Resume Next
    Call Optionals.Enve_NAOptionButton
    Call TotalFeeCalc
End Sub
Public Sub Enve_TotalBox_Change()
    On Error GoTo 0
    Call Optionals.Enve_TotalBox
End Sub
Public Sub Enve_TotalBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 46, 48 To 57
        Case Else
            KeyAscii = 0
            'MsgBox "Only numbers allowed!"
    End Select
End Sub
' Additional Fees =========================================================================================================================
Public Sub AddFee1_TotalBox_Change()
    FeeReport.AddFee1_LFBox.Locked = True
    If FeeReport.AddFee1_TotalBox <> "0" Then
        FeeReport.AddFee1_TotalBox = Format(FeeReport.AddFee1_TotalBox, "#,##0")
        FeeReport.AddFee1_LFBox = Round(FeeReport.AddFee1_TotalBox / FeeReport.LinearFeetBox, 2)
        Call TotalFeeCalc
    Else
        FeeReport.AddFee1_TotalBox.Value = 0
    End If
End Sub
Public Sub AddFee1_TotalBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 46, 48 To 57
        Case Else
            KeyAscii = 0
            'MsgBox "Only numbers allowed!"
    End Select
End Sub
Public Sub AddFee2_TotalBox_Change()
    FeeReport.AddFee2_LFBox.Locked = True
    If FeeReport.AddFee2_TotalBox <> "0" Then
        FeeReport.AddFee2_TotalBox = Format(FeeReport.AddFee2_TotalBox, "#,##0")
        FeeReport.AddFee2_LFBox = Round(FeeReport.AddFee2_TotalBox / FeeReport.LinearFeetBox, 2)
        Call TotalFeeCalc
    Else
        FeeReport.AddFee1_TotalBox.Value = 0
    End If
End Sub
Public Sub AddFee2_TotalBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 46, 48 To 57
        Case Else
            KeyAscii = 0
            'MsgBox "Only numbers allowed!"
    End Select
End Sub
Public Sub AddFee3_TotalBox_Change()
    FeeReport.AddFee3_LFBox.Locked = True
    If FeeReport.AddFee3_TotalBox <> "0" Then
        FeeReport.AddFee3_TotalBox = Format(FeeReport.AddFee3_TotalBox, "#,##0")
        FeeReport.AddFee3_LFBox = Round(FeeReport.AddFee3_TotalBox / FeeReport.LinearFeetBox, 2)
        Call TotalFeeCalc
    Else
        FeeReport.AddFee1_TotalBox.Value = 0
    End If
End Sub
Public Sub AddFee3_TotalBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 46, 48 To 57
        Case Else
            KeyAscii = 0
            'MsgBox "Only numbers allowed!"
    End Select
End Sub

'============================================================================================================================================================
' Comments Tab
'============================================================================================================================================================
Public Sub PD_TextBox_Change()
    If Len(PD_TextBox) > 50 Then PD_TextBox.Value = Left(PD_TextBox.Value, 50)
End Sub
Public Sub Design_TextBox_Change()
    If Len(Design_TextBox) > 50 Then Design_TextBox.Value = Left(Design_TextBox.Value, 50)
End Sub
Public Sub PM_TextBox_Change()
    If Len(PM_TextBox) > 50 Then PM_TextBox.Value = Left(PM_TextBox.Value, 50)
End Sub
Public Sub R_TextBox_Change()
    If Len(R_TextBox) > 50 Then R_TextBox.Value = Left(R_TextBox.Value, 50)
End Sub
Public Sub S_TextBox_Change()
    If Len(S_TextBox) > 50 Then S_TextBox.Value = Left(S_TextBox.Value, 50)
End Sub
Public Sub Geo_TextBox_Change()
    If Len(Geo_TextBox) > 50 Then Geo_TextBox.Value = Left(Geo_TextBox.Value, 50)
End Sub
Public Sub TC_TextBox_Change()
    If Len(TC_TextBox) > 50 Then TC_TextBox.Value = Left(TC_TextBox.Value, 50)
End Sub
Public Sub Pot_TextBox_Change()
    If Len(Pot_TextBox) > 50 Then Pot_TextBox.Value = Left(Pot_TextBox.Value, 50)
End Sub
Public Sub CS_TextBox_Change()
    If Len(CS_TextBox) > 50 Then CS_TextBox.Value = Left(CS_TextBox.Value, 50)
End Sub
Public Sub Enve_TextBox_Change()
    If Len(Enve_TextBox) > 50 Then Enve_TextBox.Value = Left(Enve_TextBox.Value, 50)
End Sub
