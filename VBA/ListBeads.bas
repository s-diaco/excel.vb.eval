Attribute VB_Name = "ListBeads"
' gets and sorts the added bead for each machine
' Depends on CInkSheet and CResult

Function ListBeads(ByVal SheetNames As Range, _
                    ByVal DateColumn As Range, _
                    ByVal BeadColumn As Range, _
                    ByVal FilterColumn As Range, _
                    ResultIndex As Integer, _
                    ByVal Criterion As Range, _
                    ByVal SerialColumn As Range, _
                    ByVal RetValueColumn As Range)
         ListBeads = "-"
         Dim AddedRows As Collection
         Set AddedRows = New Collection
         Dim BeadRow As CResult
         
         ' fill the values.
         Dim InkSheetCell As Range
         Dim i As Integer
         Dim n As Integer
         Dim InkSheetIndex As Integer
         InkSheetIndex = 1
         Dim InkSheet() As New CInkSheet
         ReDim InkSheet(SheetNames.Cells.Count) As New CInkSheet
         n = 1
         For Each InkSheetCell In SheetNames
         
            ' Create an inksheet instance.
            InkSheet(InkSheetIndex).DateCol = DateColumn.Cells(InkSheetIndex, 1)
            InkSheet(InkSheetIndex).BeadsCol = BeadColumn.Cells(InkSheetIndex, 1)
            InkSheet(InkSheetIndex).RetValCol = RetValueColumn.Cells(InkSheetIndex, 1)
            InkSheet(InkSheetIndex).CriteriaCol = FilterColumn.Cells(InkSheetIndex, 1)
            InkSheet(InkSheetIndex).SerialCol = SerialColumn.Cells(InkSheetIndex, 1)
            InkSheet(InkSheetIndex).SheetName = InkSheetCell.Value
            With Worksheets(InkSheet(InkSheetIndex).SheetName)
                InkSheet(InkSheetIndex).LastRow = .Cells(.Rows.Count, InkSheet(InkSheetIndex).SerialCol).End(xlUp).Row
                For i = 1 To InkSheet(InkSheetIndex).LastRow
                    If .Cells(i, InkSheet(InkSheetIndex).BeadsCol).Value <> "" _
                        And .Cells(i, InkSheet(InkSheetIndex).CriteriaCol).Value = Criterion.Cells(1, 1).Value Then
                        
                        ' Create an BeadRow instance
                        Set BeadRow = New CResult
                        BeadRow.SortIndex = Worksheets(InkSheet(InkSheetIndex).SheetName).Cells(i, InkSheet(InkSheetIndex).DateCol).Value
                        BeadRow.RowNumber = i
                        BeadRow.SheetName = InkSheet(InkSheetIndex).SheetName
                        AddedRows.Add BeadRow
                    End If
                Next i
            End With
            InkSheetIndex = InkSheetIndex + 1
         Next
         
         If ResultIndex > AddedRows.Count Then
            Exit Function
        End If
         
         ' Sort AddedRows list
         ' todo: change the algorithm to mergesort. it's faster.
         Dim j As Integer
         Dim vTemp As CResult
         Set vTemp = New CResult
         Dim SortedArray() As New CResult
         ReDim SortedArray(1 To AddedRows.Count) As New CResult
         For i = 1 To AddedRows.Count
            SortedArray(i).SortIndex = AddedRows(i).SortIndex
            SortedArray(i).SheetName = AddedRows(i).SheetName
            SortedArray(i).RowNumber = AddedRows(i).RowNumber
         Next i
         For i = 1 To AddedRows.Count - 1
            For j = i + 1 To AddedRows.Count
                If IsLaterThan(SortedArray(i).SortIndex, SortedArray(j).SortIndex, ".") Then
                    'store the lesser item
                    vTemp.RowNumber = SortedArray(j).RowNumber
                    vTemp.SheetName = SortedArray(j).SheetName
                    vTemp.SortIndex = SortedArray(j).SortIndex
                    'remove the lesser item
                    SortedArray(j).RowNumber = SortedArray(i).RowNumber
                    SortedArray(j).SheetName = SortedArray(i).SheetName
                    SortedArray(j).SortIndex = SortedArray(i).SortIndex
                    're-add the lesser item before the greater Item
                    SortedArray(i).RowNumber = vTemp.RowNumber
                    SortedArray(i).SheetName = vTemp.SheetName
                    SortedArray(i).SortIndex = vTemp.SortIndex
                End If
            Next j
         Next i
         
         ' Return the requested column value
         For i = LBound(InkSheet) To UBound(InkSheet)
            If InkSheet(i).SheetName = SortedArray(ResultIndex).SheetName Then Exit For
         Next i
         With Worksheets(SortedArray(ResultIndex).SheetName)
            ListBeads = .Cells(SortedArray(ResultIndex).RowNumber, InkSheet(i).RetValCol).Value
         End With
End Function

Function RotTimes(sDate As String, eDate As String, _
                    ByVal SheetNames As Range, _
                    ByVal DateColumn As Range, _
                    ByVal FilterColumn As Range, _
                    ByVal Criterion As Range, _
                    ByVal SerialColumn As Range, _
                    ByVal RotTimeColumn As Range)
         Dim BeadRow As CResult
         
         Dim sum As Double
         sum = 0
         If Not IsNumeric(Left(sDate, 2)) Then
            Exit Function
         End If
         
         ' fill the values.
         Dim InkSheetCell As Range
         Dim i As Integer
         Dim n As Integer
         Dim InkSheetIndex As Integer
         InkSheetIndex = 1
         Dim InkSheet() As New CInkSheet
         ReDim InkSheet(SheetNames.Cells.Count) As New CInkSheet
         n = 1
         For Each InkSheetCell In SheetNames
         
            ' Create an inksheet instance.
            InkSheet(InkSheetIndex).DateCol = DateColumn.Cells(InkSheetIndex, 1)
            InkSheet(InkSheetIndex).CriteriaCol = FilterColumn.Cells(InkSheetIndex, 1)
            InkSheet(InkSheetIndex).SerialCol = SerialColumn.Cells(InkSheetIndex, 1)
            InkSheet(InkSheetIndex).RotTimeCol = RotTimeColumn.Cells(InkSheetIndex, 1)
            InkSheet(InkSheetIndex).SheetName = InkSheetCell.Value
            With Worksheets(InkSheet(InkSheetIndex).SheetName)
                InkSheet(InkSheetIndex).LastRow = .Cells(.Rows.Count, InkSheet(InkSheetIndex).SerialCol).End(xlUp).Row
                For i = 1 To InkSheet(InkSheetIndex).LastRow
                    If Not IsLaterThan(sDate, .Cells(i, InkSheet(InkSheetIndex).DateCol).Value, ".") _
                        And .Cells(i, InkSheet(InkSheetIndex).CriteriaCol).Value = Criterion.Cells(1, 1).Value _
                        And (Not IsNumeric(Left(eDate, 2)) Or IsLaterThan(eDate, .Cells(i, InkSheet(InkSheetIndex).DateCol).Value, ".")) Then
                        sum = sum + Eval(Worksheets(InkSheet(InkSheetIndex).SheetName).Cells(i, InkSheet(InkSheetIndex).RotTimeCol))
                    End If
                Next i
            End With
            InkSheetIndex = InkSheetIndex + 1
         Next
    RotTimes = sum
End Function

Private Function IsLaterThan(FirstDate As String, SecDate As String, DateSeparator As String) As Boolean
    Dim fDate As CstringDate
    Set fDate = New CstringDate
    fDate.NumberizeDate FirstDate, DateSeparator
    Dim sDate As CstringDate
    Set sDate = New CstringDate
    sDate.NumberizeDate SecDate, DateSeparator
    IsLaterThan = False
    If fDate.Year > sDate.Year Then
        IsLaterThan = True
    Else
        If fDate.Year = sDate.Year Then
            If fDate.Month > sDate.Month Then
                IsLaterThan = True
            Else
                If fDate.Month = sDate.Month And fDate.Day > sDate.Day Then
                    IsLaterThan = True
                End If
            End If
        End If
    End If
End Function
