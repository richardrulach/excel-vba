Function MaxCol(iRow As Integer) As Integer
    MaxCol = ActiveSheet.Cells(iRow, ActiveSheet.Columns.Count).End(xlToLeft).Column
End Function

Function MaxRow(sCol As Variant) As Integer
    MaxRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, sCol).End(xlUp).Row
End Function