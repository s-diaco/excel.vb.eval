Attribute VB_Name = "Module1"
' Turns a string formula into a “real” formula and evaluates it if the condition is met.
Function EvalIf(Rng As Range, Cond As Range, Crit As Range)
    Dim I As Integer
    Dim j As Integer
    Dim sum As Integer
    sum = 0
    For I = 1 To Rng.Rows.Count
        For j = 1 To Rng.Columns.Count
            If Left(Cond.Cells(I, j), 2) = Left(Crit.Value, 2) Then
                sum = sum + Eval(Rng.Cells(I, j))
            End If
        Next j
    Next I
    
    EvalIf = sum
End Function
