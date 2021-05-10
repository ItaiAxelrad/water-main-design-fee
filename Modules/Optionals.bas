Attribute VB_Name = "Optionals"
'============================================================================================================================================================
' Optionals
'============================================================================================================================================================
'=Construction Support=========================================================================================================================================
Public Sub CS_AverageOptionButton()
On Error Resume Next
FeeReport.CS_TotalBox.Locked = True
FeeReport.CS_LFBox.Locked = True
FeeReport.CS_TotalBox.Enabled = True
FeeReport.CS_LFBox.Enabled = True
    If FeeReport.CS_AverageOptionButton.Value = True Then
            FeeReport.CS_LFBox.Value = Fee_Calc.FeeCalc("CS", "Average")
    End If
FeeReport.CS_TotalBox = CDbl(FeeReport.CS_LFBox.Value) * CDbl(FeeReport.LinearFeetBox.Value)
End Sub

Public Sub CS_HighOptionButton()
FeeReport.CS_TotalBox.Locked = True
FeeReport.CS_LFBox.Locked = True
FeeReport.CS_TotalBox.Enabled = True
FeeReport.CS_LFBox.Enabled = True
    If FeeReport.CS_HighOptionButton.Value = True Then
            FeeReport.CS_LFBox.Value = Fee_Calc.FeeCalc("CS", "High")
    End If
FeeReport.CS_TotalBox = CDbl(FeeReport.CS_LFBox.Value) * CDbl(FeeReport.LinearFeetBox.Value)
End Sub

Public Sub CS_LowOptionButton()
FeeReport.CS_TotalBox.Locked = True
FeeReport.CS_LFBox.Locked = True
FeeReport.CS_TotalBox.Enabled = True
FeeReport.CS_LFBox.Enabled = True
    If FeeReport.CS_LowOptionButton.Value = True Then
            FeeReport.CS_LFBox.Value = Fee_Calc.FeeCalc("CS", "Low")
    End If
FeeReport.CS_TotalBox = CDbl(FeeReport.CS_LFBox.Value) * CDbl(FeeReport.LinearFeetBox.Value)
End Sub

Public Sub CS_LumpSumOptionButton()
FeeReport.CS_TotalBox.Locked = False
FeeReport.CS_LFBox.Locked = True
FeeReport.CS_TotalBox.Enabled = True
FeeReport.CS_LFBox.Enabled = True
    If FeeReport.CS_LumpSumOptionButton.Value = True Then
            FeeReport.CS_LFBox.Value = "0"
    End If
FeeReport.CS_TotalBox = "0"
End Sub

Public Sub CS_NAOptionButton()
FeeReport.CS_TotalBox.Locked = True
FeeReport.CS_LFBox.Locked = True
FeeReport.CS_TotalBox.Enabled = False
FeeReport.CS_LFBox.Enabled = False
    If FeeReport.CS_NAOptionButton.Value = True Then
            FeeReport.CS_LFBox.Value = "0"
    End If
FeeReport.CS_TotalBox = "0"
End Sub
Public Sub CS_TotalBox()
    On Error GoTo 0
    If FeeReport.CS_LumpSumOptionButton.Value = True And FeeReport.CS_TotalBox.Value <> "" Then
            FeeReport.CS_LFBox.Value = Round(FeeReport.CS_TotalBox.Value / FeeReport.LinearFeetBox.Value, 2)
            Call FeeReport.TotalFeeCalc
    ElseIf FeeReport.CS_LumpSumOptionButton.Value = True And FeeReport.CS_TotalBox.Value = "" Then
        FeeReport.CS_LFBox.Value = "0"
    ElseIf FeeReport.LinearFeetBox.Value <> "" Then
        FeeReport.CS_TotalBox.Value = FeeReport.CS_LFBox.Value * FeeReport.LinearFeetBox.Value
    End If
    FeeReport.CS_TotalBox.Value = Format(FeeReport.CS_TotalBox.Value, "#,##0")
End Sub
'=Environmental Documents=========================================================================================================================================
Public Sub Enve_AverageOptionButton()
On Error Resume Next
FeeReport.Enve_TotalBox.Locked = True
FeeReport.Enve_LFBox.Locked = True
FeeReport.Enve_TotalBox.Enabled = True
FeeReport.Enve_LFBox.Enabled = True
    If FeeReport.Enve_AverageOptionButton.Value = True Then
            FeeReport.Enve_LFBox.Value = Fee_Calc.FeeCalc("Enve", "Average")
    End If
FeeReport.Enve_TotalBox = CDbl(FeeReport.Enve_LFBox.Value) * CDbl(FeeReport.LinearFeetBox.Value)
End Sub

Public Sub Enve_HighOptionButton()
FeeReport.Enve_TotalBox.Locked = True
FeeReport.Enve_LFBox.Locked = True
FeeReport.Enve_TotalBox.Enabled = True
FeeReport.Enve_LFBox.Enabled = True
    If FeeReport.Enve_HighOptionButton.Value = True Then
            FeeReport.Enve_LFBox.Value = Fee_Calc.FeeCalc("Enve", "High")
    End If
FeeReport.Enve_TotalBox = CDbl(FeeReport.Enve_LFBox.Value) * CDbl(FeeReport.LinearFeetBox.Value)
End Sub

Public Sub Enve_LowOptionButton()
FeeReport.Enve_TotalBox.Locked = True
FeeReport.Enve_LFBox.Locked = True
FeeReport.Enve_TotalBox.Enabled = True
FeeReport.Enve_LFBox.Enabled = True
    If FeeReport.Enve_LowOptionButton.Value = True Then
            FeeReport.Enve_LFBox.Value = Fee_Calc.FeeCalc("Enve", "Low")
    End If
FeeReport.Enve_TotalBox = CDbl(FeeReport.Enve_LFBox.Value) * CDbl(FeeReport.LinearFeetBox.Value)
End Sub

Public Sub Enve_LumpSumOptionButton()
FeeReport.Enve_TotalBox.Locked = False
FeeReport.Enve_LFBox.Locked = True
FeeReport.Enve_TotalBox.Enabled = True
FeeReport.Enve_LFBox.Enabled = True
    If FeeReport.Enve_LumpSumOptionButton.Value = True Then
            FeeReport.Enve_LFBox.Value = "0"
    End If
FeeReport.Enve_TotalBox = "0"
End Sub

Public Sub Enve_NAOptionButton()
FeeReport.Enve_TotalBox.Locked = True
FeeReport.Enve_LFBox.Locked = True
FeeReport.Enve_TotalBox.Enabled = False
FeeReport.Enve_LFBox.Enabled = False
    If FeeReport.Enve_NAOptionButton.Value = True Then
            FeeReport.Enve_LFBox.Value = "0"
    End If
FeeReport.Enve_TotalBox = "0"
End Sub
Public Sub Enve_TotalBox()
    On Error GoTo 0
    If FeeReport.Enve_LumpSumOptionButton.Value = True And FeeReport.Enve_TotalBox.Value <> "" Then
            FeeReport.Enve_LFBox.Value = Round(FeeReport.Enve_TotalBox.Value / FeeReport.LinearFeetBox.Value, 2)
            Call FeeReport.TotalFeeCalc
    ElseIf FeeReport.Enve_LumpSumOptionButton.Value = True And FeeReport.Enve_TotalBox.Value = "" Then
        FeeReport.Enve_LFBox.Value = "0"
    ElseIf FeeReport.LinearFeetBox.Value <> "" Then
        FeeReport.Enve_TotalBox.Value = FeeReport.Enve_LFBox.Value * FeeReport.LinearFeetBox.Value
    End If
    FeeReport.Enve_TotalBox.Value = Format(FeeReport.Enve_TotalBox.Value, "#,##0")
End Sub
