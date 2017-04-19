Attribute VB_Name = "MEval"
'
' Purpose   : Turns a string formula into a “real” formula and evaluates it.
' Comment   : [Rng] Formula in string format.
'
Function Eval(ByVal Rng As Range)

' Application.Volatile

    Dim strSearch As String
    Dim StartPt As Long
    Dim TimeNumber As Single
    Dim HourDigits As Integer
    Dim i As Integer
    Dim j As Integer
    Dim sum As Single
        
    sum = 0
    
' On Error GoTo BuildResult
    For i = 1 To Rng.Rows.Count
        For j = 1 To Rng.Columns.Count
            strSearch = Rng.Cells(i, j).Value
                       
FindTime:
            If strSearch Like "*#:##*" Then
                If (WorksheetFunction.Search(":", strSearch) - 2 > 0) Then
                    If Mid(strSearch, WorksheetFunction.Search(":", strSearch) - 2, 1) Like "#" Then
                        HourDigits = 2
                    Else
                        HourDigits = 1
                    End If
                Else
                    HourDigits = 1
                End If
                StartPt = WorksheetFunction.Search(":", strSearch) - HourDigits
                TimeNumber = TimeValue(Mid(strSearch, StartPt, HourDigits + 3))
                strSearch = WorksheetFunction.Substitute(strSearch, Mid(strSearch, StartPt, HourDigits + 3), CStr(TimeNumber))
                GoTo FindTime
            End If
            'check if the cell is not an empty or non-numeric cell.
            If strSearch Like "*#*" Then
                ' remove any non-numerinc trailing character
                Do While Not (Right(strSearch, 1) Like "#")
                    strSearch = Left(strSearch, Len(strSearch) - 1)
                Loop
                sum = sum + Evaluate(strSearch)
            End If
        Next j
    Next i
    
Eval = sum
End Function



