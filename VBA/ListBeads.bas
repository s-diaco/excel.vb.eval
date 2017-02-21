Attribute VB_Name = "Module4"
Sub ListBeads(ByVal refSheetName As String)

'get ref values

         ' Declare Current as a worksheet object variable.
         Dim Current As Worksheet

         ' Loop through all of the worksheets in the active workbook.
         For Each Current In Worksheets

            'add sheet values to array
            ' This line displays the worksheet name in a message box.
            MsgBox Current.Name
         Next
         
         'sort the arrary
         'insert results into the desired sheet
End Sub
