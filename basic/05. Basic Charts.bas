''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Author:       RR
' Created:      01/12/2016
' Description:  Demonstrates generating charts in VBA
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub AddChart()
    
    Dim newSheet As Worksheet
    With ActiveWorkbook.Sheets.Add(After:=Sheets(1))
        .Select
        For I = 1 To 20
            .Cells(I, 1).Value = I ^ 3
        Next
    End With
    
    Set newSheet = ActiveWorkbook.ActiveSheet
    
    With Charts.Add(After:=newSheet)
        .ChartWizard Source:=newSheet.Range("A1:A20"), _
            Gallery:=xlLine, Title:="My numbers"
    End With

End Sub


