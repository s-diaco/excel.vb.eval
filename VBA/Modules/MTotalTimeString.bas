Attribute VB_Name = "MTotalTimeString"
'
' Purpose   : multiplies a string of times by a specified factor.
' Comment   : [Rng] Formula in string format. this function is copied from Eval
'
Function TotalTimeString(ByVal Rng As Range, ByVal Factor As Single)

' Application.Volatile

    Dim strSearch As String
    Dim strResult As String
    Dim StartPt As Long
    Dim TimeNumber As Single
    Dim HourDigits As Integer
    Dim i As Integer
    Dim j As Integer
    strResult = ""
    
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
                strSearch = Right(strSearch, Len(strSearch) - StartPt - 3)
                strResult = strResult & Application.Text(TimeNumber * Factor, "[h]:mm")
                If Len(strSearch) > 1 Then
                    strResult = strResult & Left(strSearch, 1)
                    strSearch = Right(strSearch, Len(strSearch) - 1)
                End If
                GoTo FindTime
            Else
                If strSearch Like "*#.#*" Then
                    ' there is just one time number in the cell
                    TimeNumber = strSearch
                    strResult = strResult & Application.Text(TimeNumber * Factor, "[h]:mm")
                End If
            End If
            'check if the cell is not an empty or non-numeric cell.
            If strSearch Like "*#*" Then
            ' remove any non-numerinc trailing character
            If Not (Right(strSearch, 1) Like "#") Then
                strSearch = Left(strSearch, Len(strSearch) - 1)
            End If
            
            End If
        Next j
    Next i
    
TotalTimeString = strResult
End Function




