Attribute VB_Name = "MEvalIf2"
' Turns a string formula into a “real” formula and evaluates it if the condition is met.
Function EvalIf2(Rng As Range, Cond As Range, Crit As Range)
    Dim i As Integer
    Dim j As Integer
    Dim sum As Integer
    sum = 0
    For i = 1 To Rng.Rows.Count
        For j = 1 To Rng.Columns.Count
            If Cond.Cells(i, j).Value = Crit.Value Then
                sum = sum + Eval(Rng.Cells(i, j))
            End If
        Next j
    Next i
    
    EvalIf2 = sum
End Function
