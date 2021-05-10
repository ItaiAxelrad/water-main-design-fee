Attribute VB_Name = "Phase"
'============================================================================================================================================================
' Phases
'============================================================================================================================================================
'=Preliminary Design=========================================================================================================================================
Public Sub PD_AverageOptionButton()
On Error Resume Next
FeeReport.PD_TotalBox.Locked = True
FeeReport.PD_LFBox.Locked = True
FeeReport.PD_TotalBox.Enabled = True
FeeReport.PD_LFBox.Enabled = True
    If FeeReport.PD_AverageOptionButton.Value = True Then
            FeeReport.PD_LFBox.Value = Fee_Calc.FeeCalc("PD", "Average")
    End If
FeeReport.PD_TotalBox = CDbl(FeeReport.PD_LFBox.Value) * CDbl(FeeReport.LinearFeetBox.Value)
End Sub

Public Sub PD_HighOptionButton()
FeeReport.PD_TotalBox.Locked = True
FeeReport.PD_LFBox.Locked = True
FeeReport.PD_TotalBox.Enabled = True
FeeReport.PD_LFBox.Enabled = True
    If FeeReport.PD_HighOptionButton.Value = True Then
            FeeReport.PD_LFBox.Value = Fee_Calc.FeeCalc("PD", "High")
    End If
FeeReport.PD_TotalBox = CDbl(FeeReport.PD_LFBox.Value) * CDbl(FeeReport.LinearFeetBox.Value)
End Sub

Public Sub PD_LowOptionButton()
FeeReport.PD_TotalBox.Locked = True
FeeReport.PD_LFBox.Locked = True
FeeReport.PD_TotalBox.Enabled = True
FeeReport.PD_LFBox.Enabled = True
    If FeeReport.PD_LowOptionButton.Value = True Then
            FeeReport.PD_LFBox.Value = Fee_Calc.FeeCalc("PD", "Low")
    End If
FeeReport.PD_TotalBox = CDbl(FeeReport.PD_LFBox.Value) * CDbl(FeeReport.LinearFeetBox.Value)
End Sub

Public Sub PD_LumpSumOptionButton()
FeeReport.PD_TotalBox.Locked = False
FeeReport.PD_LFBox.Locked = True
FeeReport.PD_TotalBox.Enabled = True
FeeReport.PD_LFBox.Enabled = True
    If FeeReport.PD_LumpSumOptionButton.Value = True Then
            FeeReport.PD_LFBox.Value = "0"
    End If
FeeReport.PD_TotalBox = "0"
End Sub

Public Sub PD_NAOptionButton()
FeeReport.PD_TotalBox.Locked = True
FeeReport.PD_LFBox.Locked = True
FeeReport.PD_TotalBox.Enabled = False
FeeReport.PD_LFBox.Enabled = False
    If FeeReport.PD_NAOptionButton.Value = True Then
            FeeReport.PD_LFBox.Value = "0"
    End If
FeeReport.PD_TotalBox = "0"
End Sub
Public Sub PD_TotalBox()
    On Error GoTo 0
    If FeeReport.PD_LumpSumOptionButton.Value = True And FeeReport.PD_TotalBox.Value <> "" Then
            FeeReport.PD_LFBox.Value = Round(FeeReport.PD_TotalBox.Value / FeeReport.LinearFeetBox.Value, 2)
            Call FeeReport.TotalFeeCalc
    ElseIf FeeReport.PD_LumpSumOptionButton.Value = True And FeeReport.PD_TotalBox.Value = "" Then
        FeeReport.PD_LFBox.Value = "0"
    ElseIf FeeReport.LinearFeetBox.Value <> "" Then
        FeeReport.PD_TotalBox.Value = FeeReport.PD_LFBox.Value * FeeReport.LinearFeetBox.Value
    End If
    FeeReport.PD_TotalBox.Value = Format(FeeReport.PD_TotalBox.Value, "#,##0")
End Sub

'=Design=========================================================================================================================================
Public Sub Design_AverageOptionButton()
On Error Resume Next
FeeReport.Design_TotalBox.Locked = True
FeeReport.Design_LFBox.Locked = True
FeeReport.Design_TotalBox.Enabled = True
FeeReport.Design_LFBox.Enabled = True
    If FeeReport.Design_AverageOptionButton.Value = True Then
            FeeReport.Design_LFBox.Value = Fee_Calc.FeeCalc("Design", "Average")
    End If
FeeReport.Design_TotalBox = CDbl(FeeReport.Design_LFBox.Value) * CDbl(FeeReport.LinearFeetBox.Value)
End Sub

Public Sub Design_HighOptionButton()
FeeReport.Design_TotalBox.Locked = True
FeeReport.Design_LFBox.Locked = True
FeeReport.Design_TotalBox.Enabled = True
FeeReport.Design_LFBox.Enabled = True
    If FeeReport.Design_HighOptionButton.Value = True Then
            FeeReport.Design_LFBox.Value = Fee_Calc.FeeCalc("Design", "High")
    End If
FeeReport.Design_TotalBox = CDbl(FeeReport.Design_LFBox.Value) * CDbl(FeeReport.LinearFeetBox.Value)
End Sub

Public Sub Design_LowOptionButton()
FeeReport.Design_TotalBox.Locked = True
FeeReport.Design_LFBox.Locked = True
FeeReport.Design_TotalBox.Enabled = True
FeeReport.Design_LFBox.Enabled = True
    If FeeReport.Design_LowOptionButton.Value = True Then
            FeeReport.Design_LFBox.Value = Fee_Calc.FeeCalc("Design", "Low")
    End If
FeeReport.Design_TotalBox = CDbl(FeeReport.Design_LFBox.Value) * CDbl(FeeReport.LinearFeetBox.Value)
End Sub

Public Sub Design_LumpSumOptionButton()
FeeReport.Design_TotalBox.Locked = False
FeeReport.Design_LFBox.Locked = True
FeeReport.Design_TotalBox.Enabled = True
FeeReport.Design_LFBox.Enabled = True
    If FeeReport.Design_LumpSumOptionButton.Value = True Then
            FeeReport.Design_LFBox.Value = "0"
    End If
FeeReport.Design_TotalBox = "0"
End Sub

Public Sub Design_NAOptionButton()
FeeReport.Design_TotalBox.Locked = True
FeeReport.Design_LFBox.Locked = True
FeeReport.Design_TotalBox.Enabled = False
FeeReport.Design_LFBox.Enabled = False
    If FeeReport.Design_NAOptionButton.Value = True Then
            FeeReport.Design_LFBox.Value = "0"
    End If
FeeReport.Design_TotalBox = "0"
End Sub
Public Sub Design_TotalBox()
    On Error GoTo 0
    If FeeReport.Design_LumpSumOptionButton.Value = True And FeeReport.Design_TotalBox.Value <> "" Then
            FeeReport.Design_LFBox.Value = Round(FeeReport.Design_TotalBox.Value / FeeReport.LinearFeetBox.Value, 2)
            Call FeeReport.TotalFeeCalc
    ElseIf FeeReport.Design_LumpSumOptionButton.Value = True And FeeReport.Design_TotalBox.Value = "" Then
        FeeReport.Design_LFBox.Value = "0"
    ElseIf FeeReport.LinearFeetBox.Value <> "" Then
        FeeReport.Design_TotalBox.Value = FeeReport.Design_LFBox.Value * FeeReport.LinearFeetBox.Value
    End If
    FeeReport.Design_TotalBox.Value = Format(FeeReport.Design_TotalBox.Value, "#,##0")
End Sub

'=Project Management=========================================================================================================================================

Public Sub PM_AverageOptionButton()
On Error Resume Next
FeeReport.PM_TotalBox.Locked = True
FeeReport.PM_LFBox.Locked = True
FeeReport.PM_TotalBox.Enabled = True
FeeReport.PM_LFBox.Enabled = True
    If FeeReport.PM_AverageOptionButton.Value = True Then
            FeeReport.PM_LFBox.Value = Fee_Calc.FeeCalc("PM", "Average")
    End If
FeeReport.PM_TotalBox = CDbl(FeeReport.PM_LFBox.Value) * CDbl(FeeReport.LinearFeetBox.Value)
End Sub

Public Sub PM_HighOptionButton()
FeeReport.PM_TotalBox.Locked = True
FeeReport.PM_LFBox.Locked = True
FeeReport.PM_TotalBox.Enabled = True
FeeReport.PM_LFBox.Enabled = True
    If FeeReport.PM_HighOptionButton.Value = True Then
            FeeReport.PM_LFBox.Value = Fee_Calc.FeeCalc("PM", "High")
    End If
FeeReport.PM_TotalBox = CDbl(FeeReport.PM_LFBox.Value) * CDbl(FeeReport.LinearFeetBox.Value)
End Sub

Public Sub PM_LowOptionButton()
FeeReport.PM_TotalBox.Locked = True
FeeReport.PM_LFBox.Locked = True
FeeReport.PM_TotalBox.Enabled = True
FeeReport.PM_LFBox.Enabled = True
    If FeeReport.PM_LowOptionButton.Value = True Then
            FeeReport.PM_LFBox.Value = Fee_Calc.FeeCalc("PM", "Low")
    End If
FeeReport.PM_TotalBox = CDbl(FeeReport.PM_LFBox.Value) * CDbl(FeeReport.LinearFeetBox.Value)
End Sub

Public Sub PM_LumpSumOptionButton()
FeeReport.PM_TotalBox.Locked = False
FeeReport.PM_LFBox.Locked = True
FeeReport.PM_TotalBox.Enabled = True
FeeReport.PM_LFBox.Enabled = True
    If FeeReport.PM_LumpSumOptionButton.Value = True Then
            FeeReport.PM_LFBox.Value = "0"
    End If
FeeReport.PM_TotalBox = "0"
End Sub

Public Sub PM_NAOptionButton()
FeeReport.PM_TotalBox.Locked = True
FeeReport.PM_LFBox.Locked = True
FeeReport.PM_TotalBox.Enabled = False
FeeReport.PM_LFBox.Enabled = False
    If FeeReport.PM_NAOptionButton.Value = True Then
            FeeReport.PM_LFBox.Value = "0"
    End If
FeeReport.PM_TotalBox = "0"
End Sub
Public Sub PM_TotalBox()
    On Error GoTo 0
    If FeeReport.PM_LumpSumOptionButton.Value = True And FeeReport.PM_TotalBox.Value <> "" Then
            FeeReport.PM_LFBox.Value = Round(FeeReport.PM_TotalBox.Value / FeeReport.LinearFeetBox.Value, 2)
            Call FeeReport.TotalFeeCalc
    ElseIf FeeReport.PM_LumpSumOptionButton.Value = True And FeeReport.PM_TotalBox.Value = "" Then
        FeeReport.PM_LFBox.Value = "0"
    ElseIf FeeReport.LinearFeetBox.Value <> "" Then
        FeeReport.PM_TotalBox.Value = FeeReport.PM_LFBox.Value * FeeReport.LinearFeetBox.Value
    End If
    FeeReport.PM_TotalBox.Value = Format(FeeReport.PM_TotalBox.Value, "#,##0")
End Sub

'=Reimbursables=========================================================================================================================================
Public Sub R_AverageOptionButton()
On Error Resume Next
FeeReport.R_TotalBox.Locked = True
FeeReport.R_LFBox.Locked = True
FeeReport.R_TotalBox.Enabled = True
FeeReport.R_LFBox.Enabled = True
    If FeeReport.R_AverageOptionButton.Value = True Then
            FeeReport.R_LFBox.Value = Fee_Calc.FeeCalc("R", "Average")
    End If
FeeReport.R_TotalBox = CDbl(FeeReport.R_LFBox.Value) * CDbl(FeeReport.LinearFeetBox.Value)
End Sub

Public Sub R_HighOptionButton()
FeeReport.R_TotalBox.Locked = True
FeeReport.R_LFBox.Locked = True
FeeReport.R_TotalBox.Enabled = True
FeeReport.R_LFBox.Enabled = True
    If FeeReport.R_HighOptionButton.Value = True Then
            FeeReport.R_LFBox.Value = Fee_Calc.FeeCalc("R", "High")
    End If
FeeReport.R_TotalBox = CDbl(FeeReport.R_LFBox.Value) * CDbl(FeeReport.LinearFeetBox.Value)
End Sub

Public Sub R_LowOptionButton()
FeeReport.R_TotalBox.Locked = True
FeeReport.R_LFBox.Locked = True
FeeReport.R_TotalBox.Enabled = True
FeeReport.R_LFBox.Enabled = True
    If FeeReport.R_LowOptionButton.Value = True Then
            FeeReport.R_LFBox.Value = Fee_Calc.FeeCalc("R", "Low")
    End If
FeeReport.R_TotalBox = CDbl(FeeReport.R_LFBox.Value) * CDbl(FeeReport.LinearFeetBox.Value)
End Sub

Public Sub R_LumpSumOptionButton()
FeeReport.R_TotalBox.Locked = False
FeeReport.R_LFBox.Locked = True
FeeReport.R_TotalBox.Enabled = True
FeeReport.R_LFBox.Enabled = True
    If FeeReport.R_LumpSumOptionButton.Value = True Then
            FeeReport.R_LFBox.Value = "0"
    End If
FeeReport.R_TotalBox = "0"
End Sub

Public Sub R_NAOptionButton()
FeeReport.R_TotalBox.Locked = True
FeeReport.R_LFBox.Locked = True
FeeReport.R_TotalBox.Enabled = False
FeeReport.R_LFBox.Enabled = False
    If FeeReport.R_NAOptionButton.Value = True Then
            FeeReport.R_LFBox.Value = "0"
    End If
FeeReport.R_TotalBox = "0"
End Sub
Public Sub R_TotalBox()
    On Error GoTo 0
    If FeeReport.R_LumpSumOptionButton.Value = True And FeeReport.R_TotalBox.Value <> "" Then
            FeeReport.R_LFBox.Value = Round(FeeReport.R_TotalBox.Value / FeeReport.LinearFeetBox.Value, 2)
            Call FeeReport.TotalFeeCalc
    ElseIf FeeReport.R_LumpSumOptionButton.Value = True And FeeReport.R_TotalBox.Value = "" Then
        FeeReport.R_LFBox.Value = "0"
    ElseIf FeeReport.LinearFeetBox.Value <> "" Then
        FeeReport.R_TotalBox.Value = FeeReport.R_LFBox.Value * FeeReport.LinearFeetBox.Value
    End If
    FeeReport.R_TotalBox.Value = Format(FeeReport.R_TotalBox.Value, "#,##0")
End Sub

