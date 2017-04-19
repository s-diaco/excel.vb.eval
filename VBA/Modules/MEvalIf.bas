Attribute VB_Name = "MEvalIf"
' Turns a string formula into a “real” formula and evaluates it if the condition is met.
Function EvalIf(Rng As Range, Cond As Range, Crit As Range)
    Dim i As Integer
    Dim j As Integer
    Dim sum As Integer
    sum = 0
    For i = 1 To Rng.Rows.Count
        For j = 1 To Rng.Columns.Count
            If Left(Cond.Cells(i, j), 2) = Left(Crit.Value, 2) Then
                sum = sum + Eval(Rng.Cells(i, j))
            End If
        Next j
    Next i
    
    EvalIf = sum
End Function
