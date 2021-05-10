Attribute VB_Name = "Fee_Calc"
Function FeeCalc(Task As String, opt As Variant)
    Dim MyArray() As Variant
    Dim i As Integer, t As Integer, rw As Integer, LastRow As Integer, size As Integer
    Dim cost, optFactor, LF, sig, avg, fee As Double
    Dim rFind As Range
    t = 4
    optFactor = 0.5
    Sheets(t).Activate
    Set rFind = ActiveWorkbook.Sheets(t).Range("A1:Z1").Find(What:=CStr(Task), LookAt:=xlWhole, MatchCase:=False, SearchFormat:=False)
    col = rFind.Column
    size = ActiveWorkbook.Sheets(t).Columns(col).Cells.SpecialCells(xlCellTypeConstants).Count - 1
    LastRow = ActiveWorkbook.Sheets(t).Range("A" & Rows.Count).End(xlUp).Row + 1
    ReDim MyArray(1 To size)
        For i = 1 To size
            For rw = 2 To LastRow
                If IsEmpty(Cells(rw, col).Value) = False Then
                    MyArray(i) = ActiveWorkbook.Sheets(t).Cells(rw, col).Value
                    i = i + 1
                Else
                End If
            Next rw
        Next i
    avg = WorksheetFunction.Average(MyArray)
    sig = WorksheetFunction.StDev(MyArray) 'only compatible w/ '07?
        If opt = "High" Then
            fee = CDbl(Round(avg + (optFactor * sig), 2)) ' high option button
        ElseIf opt = "Low" Then
            fee = CDbl(Round(avg + (-1 * optFactor * sig), 2)) ' low option button
        ElseIf opt = "Average" Then
            fee = CDbl(Round(avg, 2)) ' average option button
        Else
            MsgBox "No option selected"
        End If
    FeeCalc = fee
    Set rFind = Nothing
End Function




    
