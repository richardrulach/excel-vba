

Sub CompareWorksheets()

    ' REPLACE ALL NULLS WITH EMPTY STRINGS
    Cells.Replace What:="NULL", Replacement:="", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    ' PARAMETERS FOR THE COMPARISOM
    Dim SourceSheet, SourceRow, SourceCol
    Dim DestinationSheet, DestinationRow, DestinationCol
    Dim ResultsSheet
    
    Dim EndCol, EndRow
    
    SourceSheet = 1
    SourceRow = 12
    SourceCol = 3
    
    DestinationSheet = 2
    DestinationRow = 3
    DestinationCol = 7

    Sheets.Add After:=Sheets(Sheets.Count)
    
    ResultsSheet = Sheets.Count

    EndCol = GetLastCol(SourceSheet, SourceRow, SourceCol)
    EndRow = GetLastRow(SourceSheet, SourceRow, SourceCol)
    
    Dim rowCounter, colCounter
    
    For rowCounter = 0 To EndRow - SourceRow
        For colCounter = 0 To EndCol - SourceCol
            
            Cells(rowCounter + 1, colCounter + 1).Select
            ActiveCell.FormulaR1C1 = "=" & Sheets(SourceSheet).Name & _
                "!R" & (SourceRow + rowCounter) & "C" & (SourceCol + colCounter) & "=" & Sheets(DestinationSheet).Name & _
                "!R" & (DestinationRow + rowCounter) & "C" & (DestinationCol + colCounter) & ""
                
        Next colCounter
    Next rowCounter
    
    Range(Cells(1, 1), Cells(EndRow - SourceRow + 1, EndCol - SourceCol + 1)).Select
    

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
    
    Range(Cells(1, 1), Cells(EndRow - SourceRow + 1, EndCol - SourceCol + 1)).Select
    
    
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


Function GetLastCol(SheetVal, RowVal, ColVal)
    GetLastCol = 4
End Function


Function GetLastRow(SheetVal, RowVal, ColVal)
    GetLastRow = 13
End Function
