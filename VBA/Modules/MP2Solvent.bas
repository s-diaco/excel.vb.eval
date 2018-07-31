Attribute VB_Name = "MP2Solvent"
' returns nessessary p2 p100 solvent value using previous values.
Function p2hv130(NewSerial As Range, StSerial As Integer)
    Dim StRow, NewRow, BaseSolvent, AllStSolvent, iSheet As Integer
    Dim StDensityTemp, NewDensityTemp As String
    Dim StDen, NewDen As Double
    Dim NewTemp, StTemp As Integer
    Dim FullBatchWeight As Integer
    Dim Diff As Double
    Dim wb As Workbook
    Dim P1File As Workbook
    Dim ThisBook As Workbook
    Dim LRow As Long
    Dim ShName As String
    
    ' sec yellow blue cyan
    ' first yellow beige pink
    
    iSheet = ActiveSheet.Index
    ShName = ActiveSheet.Name
    NewRow = NewSerial.Row

    For Each wb In Workbooks
        If wb.Name = "2-i-tech(first part).xlsx" Then Set P1File = wb
        If wb.Name = "3-i-tech(second part).xlsx" Then Set ThisBook = wb
    Next
    
    ' Find the last row in serial numbers
    With ThisBook.Sheets(iSheet)
        LRow = .Range("D" & .Rows.Count).End(xlUp).Row
    End With
    StRow = Application.Match(StSerial, ThisBook.Sheets(iSheet).Range("D1:D" & LRow), 0)
    
    Select Case ShName
        Case "YELLOW"
            BaseSolvent = GetFirstNumber(ThisBook.Sheets(iSheet).Range("H" & StRow))
            AllStSolvent = Eval(ThisBook.Sheets(iSheet).Range("H" & StRow))
            StDensityTemp = P1File.Sheets(ShName).Range("V" & StRow)
            NewDensityTemp = P1File.Sheets(ShName).Range("V" & NewRow)
            FullBatchWeight = GetFirstNumber(ThisBook.Sheets(iSheet).Range("I" & NewRow))
        Case "BEIGE", "PINK"
            BaseSolvent = GetFirstNumber(ThisBook.Sheets(iSheet).Range("G" & StRow))
            AllStSolvent = Eval(ThisBook.Sheets(iSheet).Range("G" & StRow))
            StDensityTemp = P1File.Sheets(ShName).Range("V" & StRow)
            NewDensityTemp = P1File.Sheets(ShName).Range("V" & NewRow)
            FullBatchWeight = GetFirstNumber(ThisBook.Sheets(iSheet).Range("I" & NewRow))
        Case "cyan L.T"
            BaseSolvent = GetFirstNumber(ThisBook.Sheets(iSheet).Range("H" & StRow))
            AllStSolvent = Eval(ThisBook.Sheets(iSheet).Range("H" & StRow))
            StDensityTemp = P1File.Sheets("cyan").Range("W" & StRow)
            NewDensityTemp = P1File.Sheets("cyan").Range("W" & NewRow)
            FullBatchWeight = GetFirstNumber(ThisBook.Sheets(iSheet).Range("J" & NewRow))
        Case "dark_blue"
            BaseSolvent = GetFirstNumber(ThisBook.Sheets(iSheet).Range("G" & StRow))
            AllStSolvent = Eval(ThisBook.Sheets(iSheet).Range("G" & StRow))
            StDensityTemp = P1File.Sheets(ShName).Range("W" & StRow)
            NewDensityTemp = P1File.Sheets(ShName).Range("W" & NewRow)
            FullBatchWeight = GetFirstNumber(ThisBook.Sheets(iSheet).Range("J" & NewRow))
        Case "red BROWN"
            BaseSolvent = GetFirstNumber(ThisBook.Sheets(iSheet).Range("G" & StRow))
            AllStSolvent = Eval(ThisBook.Sheets(iSheet).Range("G" & StRow))
            StDensityTemp = P1File.Sheets("BROWN").Range("W" & StRow)
            NewDensityTemp = P1File.Sheets("BROWN").Range("W" & NewRow)
            FullBatchWeight = GetFirstNumber(ThisBook.Sheets(iSheet).Range("I" & NewRow))
        Case Else
            BaseSolvent = GetFirstNumber(ThisBook.Sheets(iSheet).Range("G" & StRow))
            AllStSolvent = Eval(ThisBook.Sheets(iSheet).Range("G" & StRow))
            StDensityTemp = P1File.Sheets(ShName).Range("W" & StRow)
            NewDensityTemp = P1File.Sheets(ShName).Range("W" & NewRow)
            FullBatchWeight = GetFirstNumber(ThisBook.Sheets(iSheet).Range("I" & NewRow))
        End Select
        
    StDen = Split(StDensityTemp, "/")(0)
    NewDen = Split(NewDensityTemp, "/")(0)
    StTemp = Split(Replace(StDensityTemp, "*", ""), "/")(1)
    ' Old version: StTemp = Left(Right(StDensityTemp, 3), 2)
    NewTemp = Split(Replace(NewDensityTemp, "*", ""), "/")(1)
    
    Diff = (((NewTemp - StTemp) / 1000) + NewDen - StDen) * 1.5 * FullBatchWeight
    
    p2hv130 = BaseSolvent & Format(AllStSolvent - BaseSolvent + Diff, "+0;-0")
End Function
