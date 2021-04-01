Attribute VB_Name = "Module12"
Sub êVïisku()
    
    Dim i As Long
    For i = 2 To LastRow(Sheet1, 2)
    
      For j = 2 To 5
        If Sheet1.Cells(i, j) = "" Then
          Sheet1.Range(Sheet1.Cells(i, 2), Sheet1.Cells(i, 5)) = 0
          Exit For
        End If
      Next j
    
      Sheet1.Cells(i, 2) = Application.WorksheetFunction.Round(Sheet1.Cells(i, 2) * 2.54 * 10, 0)
      Sheet1.Cells(i, 3) = Application.WorksheetFunction.Round(Sheet1.Cells(i, 3) * 2.54 * 10, 0)
      Sheet1.Cells(i, 4) = Application.WorksheetFunction.Round(Sheet1.Cells(i, 4) * 2.54 * 10, 0)
      Sheet1.Cells(i, 5) = Application.WorksheetFunction.Round(Sheet1.Cells(i, 5) * 453.6 / 10, 0)
    Next i
    
End Sub
Function LastRow(sheetobj As Worksheet, C)
    LastRow = sheetobj.Cells(sheetobj.Rows.Count, C).End(xlUp).Row
End Function
Function LastCol(sheetobj As Worksheet, r)
    LastCol = sheetobj.Cells(r, sheetobj.Columns.Count).End(xlToLeft).Column
End Function
