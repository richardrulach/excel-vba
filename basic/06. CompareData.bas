' IDENTIFY THE SHEET, ROW AND COL OF THE FIRST CELLS OF THE AREAS TO BE COMPARED
Const A_Title = "A"
Const A_Sheet = 1
Const A_Row = 12
Const A_Col = 3
    
Const B_Title = "B"
Const B_Sheet = 2
Const B_Row = 3
Const B_Col = 7
    
Sub Compare_OnDataSetB()

    Dim EndCol, EndRow
    EndCol = GetLastCol(A_Sheet, A_Row, A_Col)
    EndRow = GetLastRow(A_Sheet, A_Row, A_Col)

    Dim rowCounter, colCounter
    
    For rowCounter = 0 To EndRow - A_Row
        For colCounter = 0 To EndCol - A_Col
        
            ' ADD COMMENT IF THEY ARE DIFFERENT
            With Sheets(B_Sheet).Cells(B_Row + rowCounter, B_Col + colCounter)
                If Not (.Comment Is Nothing) Then .Comment.Delete
            
                If Sheets(A_Sheet).Cells(A_Row + rowCounter, A_Col + colCounter).Value <> .Value Then
                        .Style = "Bad"
                        .AddComment
                        .Comment.Visible = False
                        .Comment.Text Text:="Column: " & Sheets(A_Sheet).Cells(A_Row - 1, A_Col + colCounter).Value & Chr(10) & _
                                            "Row: " & Sheets(A_Sheet).Cells(A_Row + rowCounter, A_Col).Value & Chr(10) & Chr(10) & _
                                            A_Title & ": " & Sheets(A_Sheet).Cells(A_Row + rowCounter, A_Col + colCounter).Value & Chr(10) & _
                                            B_Title & ": " & Sheets(B_Sheet).Cells(B_Row + rowCounter, B_Col + colCounter).Value & Chr(10)
                Else
                        .Style = "Normal"
                End If
            End With
        
        Next colCounter
    Next rowCounter

End Sub



Sub Compare_NewSheet()

    ' REPLACE ALL NULLS WITH EMPTY STRINGS
    Cells.Replace What:="NULL", Replacement:="", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    ' PARAMETERS FOR THE COMPARISOM
    Dim ResultsSheet
    Dim EndCol, EndRow
    
    Sheets.Add After:=Sheets(Sheets.Count)
    
    ResultsSheet = Sheets.Count

    EndCol = GetLastCol(A_Sheet, A_Row, A_Col)
    EndRow = GetLastRow(A_Sheet, A_Row, A_Col)
    
    Dim rowCounter, colCounter
    
    For rowCounter = 0 To EndRow - A_Row
        For colCounter = 0 To EndCol - A_Col
            
            ' SHOW THE RESULTS OF THE COMPARISOM
            Cells(rowCounter + 2, colCounter + 1).Select
            ActiveCell.FormulaR1C1 = _
                "=" & Sheets(A_Sheet).Name & "!R" & (A_Row + rowCounter) & "C" & (A_Col + colCounter) & _
                "=" & Sheets(B_Sheet).Name & "!R" & (B_Row + rowCounter) & "C" & (B_Col + colCounter)
                
            ' ADD COMMENT IF THEY ARE DIFFERENT
            If Sheets(A_Sheet).Cells(A_Row + rowCounter, A_Col + colCounter).Value <> _
                Sheets(B_Sheet).Cells(B_Row + rowCounter, B_Col + colCounter).Value Then
                With Cells(rowCounter + 2, colCounter + 1)
                    .AddComment
                    .Comment.Visible = False
                    .Comment.Text Text:="Column: " & Sheets(A_Sheet).Cells(A_Row - 1, A_Col + colCounter).Value & Chr(10) & _
                                            "Row: " & Sheets(A_Sheet).Cells(A_Row + rowCounter, A_Col).Value & Chr(10) & Chr(10) & _
                                            A_Title & ": " & Sheets(A_Sheet).Cells(A_Row + rowCounter, A_Col + colCounter).Value & Chr(10) & _
                                            B_Title & ": " & Sheets(B_Sheet).Cells(B_Row + rowCounter, B_Col + colCounter).Value & Chr(10)
                End With
            End If
                
        Next colCounter
    Next rowCounter
    
    Range(Cells(2, 1), Cells(EndRow - A_Row + 2, EndCol - A_Col + 1)).Select
    Call SetGreenForTrueRedForFalse
    
    Range("A1").Select
End Sub


Function GetLastCol(SheetVal, RowVal, ColVal)
    Dim LastCol As Integer
    With Sheets(SheetVal)
        LastCol = .Cells(RowVal - 1, .Columns.Count).End(xlToLeft).Column
    End With
    GetLastCol = LastCol
End Function


Function GetLastRow(SheetVal, RowVal, ColVal)
    Dim LastRow As Long
    With Sheets(SheetVal)
        LastRow = .Cells(.Rows.Count, ColVal).End(xlUp).Row
    End With
    GetLastRow = LastRow
End Function


' ADDS FORMATTING RULES FOR THE SELECTED CELLS
Sub SetGreenForTrueRedForFalse()
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=TRUE"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=FALSE"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub

