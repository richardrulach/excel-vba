''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Author:       RR
' Created:      01/12/2016
' Description:  Demonstrates the access and use of the main objects in Excel
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' Add sheets to the workbook with the names specified
Sub AddSheets()
    Dim answer
    answer = InputBox("How many sheets do you want?")
    
    If Not isValidNumber(answer) Then
        MsgBox "That is not a valid number"
        Exit Sub
    End If
    
    For I = 1 To answer
        Dim newSheetName
        newSheetName = InputBox("Type the new sheet name")
                
        Dim newSheet As Worksheet
        With ActiveWorkbook.Worksheets.Add(after:=ActiveSheet)
            .name = newSheetName
        End With
    Next
End Sub

Function isValidNumber(testValue)
    Dim bReturn As Boolean
    bReturn = True
    
    If (Not IsNumeric(testValue)) Then bReturn = False
    If testValue < 1 Then bReturn = False

    isValidNumber = bReturn
End Function

' CREATE A TABLE AND LINK A QUERY TABLE TO IT
Sub RunWebQuery()

    Set shFirstQtr = Workbooks(1).Worksheets(1)
    Set qtQtrResults = shFirstQtr.QueryTables _
        .Add(Connection:="URL;http://www.legislation.gov.uk/developer/formats/rdf", _
        Destination:=shFirstQtr.Cells(1, 1))
    
    With qtQtrResults
        .WebFormatting = xlNone
        .WebSelectionType = xlSpecifiedTables
        .WebTables = "1"
        .Refresh
    End With

End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Author:       RR
' Created:      01/12/2016
' Description:  Demonstrates the access and use of the main objects in Excel
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' Add sheets to the workbook with the names specified
Sub AddSheets()
    Dim answer
    answer = InputBox("How many sheets do you want?")
    
    If Not isValidNumber(answer) Then
        MsgBox "That is not a valid number"
        Exit Sub
    End If
    
    For I = 1 To answer
        Dim newSheetName
        newSheetName = InputBox("Type the new sheet name")
                
        Dim newSheet As Worksheet
        With ActiveWorkbook.Worksheets.Add(after:=ActiveSheet)
            .name = newSheetName
        End With
    Next
End Sub

Function isValidNumber(testValue)
    Dim bReturn As Boolean
    bReturn = True
    
    If (Not IsNumeric(testValue)) Then bReturn = False
    If testValue < 1 Then bReturn = False

    isValidNumber = bReturn
End Function

' CREATE A TABLE AND LINK A QUERY TABLE TO IT
Sub RunWebQuery()

    Set shFirstQtr = Workbooks(1).Worksheets(1)
    Set qtQtrResults = shFirstQtr.QueryTables _
        .Add(Connection:="URL;http://www.legislation.gov.uk/developer/formats/rdf", _
        Destination:=shFirstQtr.Cells(1, 1))
    
    With qtQtrResults
        .WebFormatting = xlNone
        .WebSelectionType = xlSpecifiedTables
        .WebTables = "1"
        .Refresh
    End With

End Sub
