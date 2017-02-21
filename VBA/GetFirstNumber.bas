Attribute VB_Name = "MGetFirstNumber"
' Gets the first number in a string formula.
Function GetFirstNumber(Rng As Range)
    Dim i As Long, pos As Long
    Dim sString As String

    sString = UCase$(Rng.Value)
    pos = Len(sString) + 1
    For i = 1 To Len(sString)
        Select Case Asc(Mid$(sString, i, 1))
        Case 48 To 57, 58
        Case Else: pos = i: Exit For
        End Select
    Next i

    GetFirstNumber = Left(Rng.Value, pos - 1)
End Function
