VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CInkSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim vSheetName As String
Dim vBeadsCol As String
Dim vDateCol As String
Dim vRetValCol As String
Dim vCriteriaCol As String
Dim vLastRow As Long
Dim vSerialCol As String
Dim vRotTimeCol As String
Dim vFileNameCol As String

Public Property Get BeadsCol() As String
    BeadsCol = vBeadsCol
End Property

Public Property Let BeadsCol(rn As String)
    vBeadsCol = rn
End Property

Public Property Get DateCol() As String
    DateCol = vDateCol
End Property

Public Property Let DateCol(si As String)
    vDateCol = si
End Property

Public Property Get SheetName() As String
    SheetName = vSheetName
End Property

Public Property Let SheetName(sn As String)
    'If Len(lp) <> 6 Then Err.Raise (xlErrValue) 'Raise error
    vSheetName = sn
End Property

Public Property Get RetValCol() As String
    RetValCol = vRetValCol
End Property

Public Property Let RetValCol(rvc As String)
    vRetValCol = rvc
End Property

Public Property Get CriteriaCol() As String
    CriteriaCol = vCriteriaCol
End Property

Public Property Let CriteriaCol(cc As String)
    vCriteriaCol = cc
End Property

Public Property Get LastRow() As String
    LastRow = vLastRow
End Property

Public Property Let LastRow(lr As String)
    vLastRow = lr
End Property

Public Property Get SerialCol() As String
    SerialCol = vSerialCol
End Property

Public Property Let SerialCol(sc As String)
    vSerialCol = sc
End Property

Public Property Get RotTimeCol() As String
    RotTimeCol = vRotTimeCol
End Property

Public Property Let RotTimeCol(rt As String)
    vRotTimeCol = rt
End Property

Public Property Get FileNameCol() As String
    FileNameCol = vFileNameCol
End Property

Public Property Let FileNameCol(fn As String)
    vFileNameCol = fn
End Property

