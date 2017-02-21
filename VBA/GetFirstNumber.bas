Attribute VB_Name = "Module2"
' Gets the first number in a string formula.
Function GetFirstNumber(Rng As Range)
    Dim I As Long, pos As Long
    Dim sString As String

    sString = UCase$(Rng.Value)
    pos = Len(sString) + 1
    For I = 1 To Len(sString)
        Select Case Asc(Mid$(sString, I, 1))
        Case 48 To 57, 58
        Case Else: pos = I: Exit For
        End Select
    Next I

    GetFirstNumber = Left(Rng.Value, pos - 1)
End Function
