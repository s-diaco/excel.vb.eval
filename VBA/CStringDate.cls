VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CstringDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim yr As Integer
Dim mn As Integer
Dim dy As Integer

Public Property Get Day() As Integer
    Day = dy
End Property

Public Property Get Month() As Integer
    Month = mn
End Property

Public Property Get Year() As String
    Year = yr
End Property

Public Sub NumberizeDate(BeadsDate As String, DateSeparator As String)
    Dim FstSep As Integer
    Dim SecSep As Integer
    yr = 0
    mn = 0
    dy = 0
    FstSep = 0
    SecSep = 0
    
    FstSep = InStr(1, BeadsDate, DateSeparator, vbTextCompare)
    If FstSep = 3 Then
        SecSep = InStr(FstSep + 1, BeadsDate, DateSeparator, vbTextCompare)
        If SecSep > 4 Then
            'two separator in the string
            If IsNumeric(Left(BeadsDate, 2)) And SecSep - FstSep > 1 Then
                yr = Left(BeadsDate, 2)
                If IsNumeric(Mid(BeadsDate, FstSep + 1, SecSep - FstSep - 1)) Then
                    mn = Mid(BeadsDate, FstSep + 1, SecSep - FstSep - 1)
                    If IsNumeric(Mid(BeadsDate, SecSep + 1)) Then
                        dy = Mid(BeadsDate, SecSep + 1)
                    End If
                End If
            End If
        Else
            ' one separator in the string
            If SecSep = 0 Then
                If IsNumeric(Left(BeadsDate, 2)) Then
                    yr = Left(BeadsDate, 2)
                    If IsNumeric(Mid(BeadsDate, FstSep + 1)) Then
                        mn = Mid(BeadsDate, FstSep + 1)
                    End If
                End If
            Else
            End If
        End If
    Else
        ' no separators in the string
        If FstSep = 0 And IsNumeric(BeadsDate) Then
            yr = CInt(BeadsDate)
        End If
    End If
End Sub
