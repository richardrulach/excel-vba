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
    SourceRow = 1
    SourceCol = 1
    
    DestinationSheet = 2
    DestinationRow = 1
    DestinationCol = 1

    Sheets.Add After:=Sheets(Sheets.Count)
    
    ResultsSheet = Sheets.Count

    EndCol = GetLastCol(SourceSheet, SourceRow, SourceCol)
    EndRow = GetLastRow(SourceSheet, SourceRow, SourceCol)
    
End Sub


Function GetLastCol(SheetVal, RowVal, ColVal)
    GetLastCol = 2
End Function


Function GetLastRow(SheetVal, RowVal, ColVal)
    GetEndRow = 2
End Function
