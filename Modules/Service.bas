Attribute VB_Name = "Service"
'============================================================================================================================================================
' Service
'============================================================================================================================================================
'=Survey=========================================================================================================================================
Public Sub S_AverageOptionButton()
On Error Resume Next
FeeReport.S_TotalBox.Locked = True
FeeReport.S_LFBox.Locked = True
FeeReport.S_TotalBox.Enabled = True
FeeReport.S_LFBox.Enabled = True
    If FeeReport.S_AverageOptionButton.Value = True Then
            FeeReport.S_LFBox.Value = Fee_Calc.FeeCalc("S", "Average")
    End If
FeeReport.S_TotalBox = CDbl(FeeReport.S_LFBox.Value) * CDbl(FeeReport.LinearFeetBox.Value)
End Sub

Public Sub S_HighOptionButton()
FeeReport.S_TotalBox.Locked = True
FeeReport.S_LFBox.Locked = True
FeeReport.S_TotalBox.Enabled = True
FeeReport.S_LFBox.Enabled = True
    If FeeReport.S_HighOptionButton.Value = True Then
            FeeReport.S_LFBox.Value = Fee_Calc.FeeCalc("S", "High")
    End If
FeeReport.S_TotalBox = CDbl(FeeReport.S_LFBox.Value) * CDbl(FeeReport.LinearFeetBox.Value)
End Sub

Public Sub S_LowOptionButton()
FeeReport.S_TotalBox.Locked = True
FeeReport.S_LFBox.Locked = True
FeeReport.S_TotalBox.Enabled = True
FeeReport.S_LFBox.Enabled = True
    If FeeReport.S_LowOptionButton.Value = True Then
            FeeReport.S_LFBox.Value = Fee_Calc.FeeCalc("S", "Low")
    End If
FeeReport.S_TotalBox = CDbl(FeeReport.S_LFBox.Value) * CDbl(FeeReport.LinearFeetBox.Value)
End Sub

Public Sub S_LumpSumOptionButton()
FeeReport.S_TotalBox.Locked = False
FeeReport.S_LFBox.Locked = True
FeeReport.S_TotalBox.Enabled = True
FeeReport.S_LFBox.Enabled = True
    If FeeReport.S_LumpSumOptionButton.Value = True Then
            FeeReport.S_LFBox.Value = "0"
    End If
FeeReport.S_TotalBox = "0"
End Sub

Public Sub S_NAOptionButton()
FeeReport.S_TotalBox.Locked = True
FeeReport.S_LFBox.Locked = True
FeeReport.S_TotalBox.Enabled = False
FeeReport.S_LFBox.Enabled = False
    If FeeReport.S_NAOptionButton.Value = True Then
            FeeReport.S_LFBox.Value = "0"
    End If
FeeReport.S_TotalBox = "0"
End Sub
Public Sub S_TotalBox()
    On Error GoTo 0
    If FeeReport.S_LumpSumOptionButton.Value = True And FeeReport.S_TotalBox.Value <> "" Then
            FeeReport.S_LFBox.Value = Round(FeeReport.S_TotalBox.Value / FeeReport.LinearFeetBox.Value, 2)
            Call FeeReport.TotalFeeCalc
    ElseIf FeeReport.S_LumpSumOptionButton.Value = True And FeeReport.S_TotalBox.Value = "" Then
        FeeReport.S_LFBox.Value = "0"
    ElseIf FeeReport.LinearFeetBox.Value <> "" Then
        FeeReport.S_TotalBox.Value = FeeReport.S_LFBox.Value * FeeReport.LinearFeetBox.Value
    End If
    FeeReport.S_TotalBox.Value = Format(FeeReport.S_TotalBox.Value, "#,##0")
End Sub


'=Geo=========================================================================================================================================

Public Sub Geo_AverageOptionButton()
On Error Resume Next
FeeReport.Geo_TotalBox.Locked = True
FeeReport.Geo_LFBox.Locked = True
FeeReport.Geo_TotalBox.Enabled = True
FeeReport.Geo_LFBox.Enabled = True
    If FeeReport.Geo_AverageOptionButton.Value = True Then
            FeeReport.Geo_LFBox.Value = Fee_Calc.FeeCalc("Geo", "Average")
    End If
FeeReport.Geo_TotalBox = CDbl(FeeReport.Geo_LFBox.Value) * CDbl(FeeReport.LinearFeetBox.Value)
End Sub

Public Sub Geo_HighOptionButton()
FeeReport.Geo_TotalBox.Locked = True
FeeReport.Geo_LFBox.Locked = True
FeeReport.Geo_TotalBox.Enabled = True
FeeReport.Geo_LFBox.Enabled = True
    If FeeReport.Geo_HighOptionButton.Value = True Then
            FeeReport.Geo_LFBox.Value = Fee_Calc.FeeCalc("Geo", "High")
    End If
FeeReport.Geo_TotalBox = CDbl(FeeReport.Geo_LFBox.Value) * CDbl(FeeReport.LinearFeetBox.Value)
End Sub

Public Sub Geo_LowOptionButton()
FeeReport.Geo_TotalBox.Locked = True
FeeReport.Geo_LFBox.Locked = True
FeeReport.Geo_TotalBox.Enabled = True
FeeReport.Geo_LFBox.Enabled = True
    If FeeReport.Geo_LowOptionButton.Value = True Then
            FeeReport.Geo_LFBox.Value = Fee_Calc.FeeCalc("Geo", "Low")
    End If
FeeReport.Geo_TotalBox = CDbl(FeeReport.Geo_LFBox.Value) * CDbl(FeeReport.LinearFeetBox.Value)
End Sub

Public Sub Geo_LumpSumOptionButton()
FeeReport.Geo_TotalBox.Locked = False
FeeReport.Geo_LFBox.Locked = True
FeeReport.Geo_TotalBox.Enabled = True
FeeReport.Geo_LFBox.Enabled = True
    If FeeReport.Geo_LumpSumOptionButton.Value = True Then
            FeeReport.Geo_LFBox.Value = "0"
    End If
FeeReport.Geo_TotalBox = "0"
End Sub

Public Sub Geo_NAOptionButton()
FeeReport.Geo_TotalBox.Locked = True
FeeReport.Geo_LFBox.Locked = True
FeeReport.Geo_TotalBox.Enabled = False
FeeReport.Geo_LFBox.Enabled = False
    If FeeReport.Geo_NAOptionButton.Value = True Then
            FeeReport.Geo_LFBox.Value = "0"
    End If
FeeReport.Geo_TotalBox = "0"
End Sub
Public Sub Geo_TotalBox()
    On Error GoTo 0
    If FeeReport.Geo_LumpSumOptionButton.Value = True And FeeReport.Geo_TotalBox.Value <> "" Then
            FeeReport.Geo_LFBox.Value = Round(FeeReport.Geo_TotalBox.Value / FeeReport.LinearFeetBox.Value, 2)
            Call FeeReport.TotalFeeCalc
    ElseIf FeeReport.Geo_LumpSumOptionButton.Value = True And FeeReport.Geo_TotalBox.Value = "" Then
        FeeReport.Geo_LFBox.Value = "0"
    ElseIf FeeReport.LinearFeetBox.Value <> "" Then
        FeeReport.Geo_TotalBox.Value = FeeReport.Geo_LFBox.Value * FeeReport.LinearFeetBox.Value
    End If
    FeeReport.Geo_TotalBox.Value = Format(FeeReport.Geo_TotalBox.Value, "#,##0")
End Sub

'=Traffic Control=========================================================================================================================================

Public Sub TC_AverageOptionButton()
On Error Resume Next
FeeReport.TC_TotalBox.Locked = True
FeeReport.TC_LFBox.Locked = True
FeeReport.TC_TotalBox.Enabled = True
FeeReport.TC_LFBox.Enabled = True
    If FeeReport.TC_AverageOptionButton.Value = True Then
            FeeReport.TC_LFBox.Value = Fee_Calc.FeeCalc("TC", "Average")
    End If
FeeReport.TC_TotalBox = CDbl(FeeReport.TC_LFBox.Value) * CDbl(FeeReport.LinearFeetBox.Value)
End Sub

Public Sub TC_HighOptionButton()
FeeReport.TC_TotalBox.Locked = True
FeeReport.TC_LFBox.Locked = True
FeeReport.TC_TotalBox.Enabled = True
FeeReport.TC_LFBox.Enabled = True
    If FeeReport.TC_HighOptionButton.Value = True Then
            FeeReport.TC_LFBox.Value = Fee_Calc.FeeCalc("TC", "High")
    End If
FeeReport.TC_TotalBox = CDbl(FeeReport.TC_LFBox.Value) * CDbl(FeeReport.LinearFeetBox.Value)
End Sub

Public Sub TC_LowOptionButton()
FeeReport.TC_TotalBox.Locked = True
FeeReport.TC_LFBox.Locked = True
FeeReport.TC_TotalBox.Enabled = True
FeeReport.TC_LFBox.Enabled = True
    If FeeReport.TC_LowOptionButton.Value = True Then
            FeeReport.TC_LFBox.Value = Fee_Calc.FeeCalc("TC", "Low")
    End If
FeeReport.TC_TotalBox = CDbl(FeeReport.TC_LFBox.Value) * CDbl(FeeReport.LinearFeetBox.Value)
End Sub

Public Sub TC_LumpSumOptionButton()
FeeReport.TC_TotalBox.Locked = False
FeeReport.TC_LFBox.Locked = True
FeeReport.TC_TotalBox.Enabled = True
FeeReport.TC_LFBox.Enabled = True
    If FeeReport.TC_LumpSumOptionButton.Value = True Then
            FeeReport.TC_LFBox.Value = "0"
    End If
FeeReport.TC_TotalBox = "0"
End Sub

Public Sub TC_NAOptionButton()
FeeReport.TC_TotalBox.Locked = True
FeeReport.TC_LFBox.Locked = True
FeeReport.TC_TotalBox.Enabled = False
FeeReport.TC_LFBox.Enabled = False
    If FeeReport.TC_NAOptionButton.Value = True Then
            FeeReport.TC_LFBox.Value = "0"
    End If
FeeReport.TC_TotalBox = "0"
End Sub
Public Sub TC_TotalBox()
    On Error GoTo 0
    If FeeReport.TC_LumpSumOptionButton.Value = True And FeeReport.TC_TotalBox.Value <> "" Then
            FeeReport.TC_LFBox.Value = Round(FeeReport.TC_TotalBox.Value / FeeReport.LinearFeetBox.Value, 2)
            Call FeeReport.TotalFeeCalc
    ElseIf FeeReport.TC_LumpSumOptionButton.Value = True And FeeReport.TC_TotalBox.Value = "" Then
        FeeReport.TC_LFBox.Value = "0"
    ElseIf FeeReport.LinearFeetBox.Value <> "" Then
        FeeReport.TC_TotalBox.Value = FeeReport.TC_LFBox.Value * FeeReport.LinearFeetBox.Value
    End If
    FeeReport.TC_TotalBox.Value = Format(FeeReport.TC_TotalBox.Value, "#,##0")
End Sub

'=Potholes=========================================================================================================================================
Public Sub Pot_AverageOptionButton()
On Error Resume Next
FeeReport.Pot_TotalBox.Locked = True
FeeReport.Pot_LFBox.Locked = True
FeeReport.Pot_TotalBox.Enabled = True
FeeReport.Pot_LFBox.Enabled = True
    If FeeReport.Pot_AverageOptionButton.Value = True Then
            FeeReport.Pot_LFBox.Value = Fee_Calc.FeeCalc("Pot", "Average")
    End If
FeeReport.Pot_TotalBox = CDbl(FeeReport.Pot_LFBox.Value) * CDbl(FeeReport.LinearFeetBox.Value)
End Sub

Public Sub Pot_HighOptionButton()
FeeReport.Pot_TotalBox.Locked = True
FeeReport.Pot_LFBox.Locked = True
FeeReport.Pot_TotalBox.Enabled = True
FeeReport.Pot_LFBox.Enabled = True
    If FeeReport.Pot_HighOptionButton.Value = True Then
            FeeReport.Pot_LFBox.Value = Fee_Calc.FeeCalc("Pot", "High")
    End If
FeeReport.Pot_TotalBox = CDbl(FeeReport.Pot_LFBox.Value) * CDbl(FeeReport.LinearFeetBox.Value)
End Sub

Public Sub Pot_LowOptionButton()
FeeReport.Pot_TotalBox.Locked = True
FeeReport.Pot_LFBox.Locked = True
FeeReport.Pot_TotalBox.Enabled = True
FeeReport.Pot_LFBox.Enabled = True
    If FeeReport.Pot_LowOptionButton.Value = True Then
            FeeReport.Pot_LFBox.Value = Fee_Calc.FeeCalc("Pot", "Low")
    End If
FeeReport.Pot_TotalBox = CDbl(FeeReport.Pot_LFBox.Value) * CDbl(FeeReport.LinearFeetBox.Value)
End Sub

Public Sub Pot_LumpSumOptionButton()
FeeReport.Pot_TotalBox.Locked = False
FeeReport.Pot_LFBox.Locked = True
FeeReport.Pot_TotalBox.Enabled = True
FeeReport.Pot_LFBox.Enabled = True
    If FeeReport.Pot_LumpSumOptionButton.Value = True Then
            FeeReport.Pot_LFBox.Value = "0"
    End If
FeeReport.Pot_TotalBox = "0"
End Sub

Public Sub Pot_NAOptionButton()
FeeReport.Pot_TotalBox.Locked = True
FeeReport.Pot_LFBox.Locked = True
FeeReport.Pot_QuantityBox.Locked = True
FeeReport.Pot_TotalBox.Enabled = False
FeeReport.Pot_LFBox.Enabled = False
FeeReport.Pot_QuantityBox.Enabled = Flase
    If FeeReport.Pot_NAOptionButton.Value = True Then
            FeeReport.Pot_LFBox.Value = "0"
    End If
FeeReport.Pot_TotalBox = "0"
End Sub
Public Sub Pot_QuantityOptionButton()
FeeReport.Pot_TotalBox.Locked = True
FeeReport.Pot_LFBox.Locked = True
FeeReport.Pot_TotalBox.Enabled = True
FeeReport.Pot_LFBox.Enabled = True
FeeReport.Pot_QuantityBox.Enabled = True
FeeReport.Pot_QuantityBox.Locked = False
    If FeeReport.Pot_QuantityOptionButton.Value = True Then
        FeeReport.Pot_LFBox.Value = "0"
        FeeReport.Pot_TotalBox = "0"
        Call FeeReport.TotalFeeCalc
    End If
End Sub
Public Sub Pot_QuantityBox()
If FeeReport.Pot_QuantityOptionButton = True And FeeReport.Pot_QuantityBox.Value <> "" Or "0" Then
    Set potRange = Worksheets(4).Range("J:J")
    FeeReport.Pot_TotalBox = CDbl(FeeReport.Pot_QuantityBox.Value * Application.WorksheetFunction.Average(potRange))
End If
FeeReport.Pot_TotalBox.Value = Format(FeeReport.Pot_TotalBox, "#,##0")
End Sub
Public Sub Pot_TotalBox()
     On Error GoTo 0
     If FeeReport.Pot_LumpSumOptionButton.Value = True And FeeReport.Pot_TotalBox.Value <> "" Then
        FeeReport.Pot_LFBox.Value = Round(FeeReport.Pot_TotalBox.Value / FeeReport.LinearFeetBox.Value, 2)
        Call FeeReport.TotalFeeCalc
    ElseIf LinearFeetBox <> "" Then
        FeeReport.Pot_TotalBox.Value = Pot_LFBox.Value * FeeReport.LinearFeetBox.Value
    End If
    
    If FeeReport.Pot_QuantityOptionButton.Value = True And FeeReport.Pot_TotalBox.Value <> "" Then
        FeeReport.Pot_LFBox = Round(CDbl(FeeReport.Pot_TotalBox.Value) / CDbl(FeeReport.LinearFeetBox.Value), 2)
        Call FeeReport.TotalFeeCalc
    Else
        FeeReport.Pot_QuantityBox.Value = "0"
    End If
    FeeReport.Pot_TotalBox.Value = Format(FeeReport.Pot_TotalBox.Value, "#,##0")
End Sub
