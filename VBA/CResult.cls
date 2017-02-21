VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim vSheetName As String
Dim vRowNumber As Integer
Dim vSortIndex As String

Public Property Get RowNumber() As Integer
    RowNumber = vRowNumber
End Property

Public Property Let RowNumber(rn As Integer)
    vRowNumber = rn
End Property
Public Property Get SortIndex() As String
    SortIndex = vSortIndex
End Property

Public Property Let SortIndex(si As String)
    vSortIndex = si
End Property

Public Property Get SheetName() As String
    SheetName = vSheetName
End Property

Public Property Let SheetName(sn As String)
    'If Len(lp) <> 6 Then Err.Raise (xlErrValue) 'Raise error
    vSheetName = sn
End Property
