Sub AddSheet()

    Dim sWorkbook As String
    sWorkbook = "c:\development\excel\BookToEdit.xlsx"
    sWorkbookName = Right(sWorkbook, Len(sWorkbook) - InStrRev(sWorkbook, "\"))
    
    If Not IsOpen(sWorkbook) Then
        Workbooks.Open (sWorkbook)
    End If
    
    With Workbooks(sWorkbookName)
        .Sheets.Add
        .Sheets(1).Cells(1, 1) = 200
        .Sheets(1).Cells(2, 1) = 300
        .Sheets(1).Cells(3, 1) = "=sum(a1:a2)"
    End With
    
End Sub

Function IsOpen(sPath)
    
    Dim bReturn As Boolean
    bReturn = False
    
    Dim wb As Workbook
    For Each wb In Workbooks
        If LCase(wb.FullName) = LCase(sPath) Then
            bReturn = True
        End If
    Next
        
    IsOpen = bReturn
End Function
