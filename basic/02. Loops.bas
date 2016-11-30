''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Author:       RR
' Created:      30/11/2016
' Description:  DEMONSTRATE THE VARIOUS LOOPING STATEMENTS IN EXCEL
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Const y = 5

' Always runs at least once
Sub DoLoopWhile()
    Dim x
    x = 1
    
    Sheets(1).Select
    Do
        ActiveSheet.Cells(x, y - 1).Value = x * 2
        x = x + 1
    Loop While x > y

End Sub


Sub DoWhileLoop()
    Dim x
    x = 1
    
    Sheets(1).Select
    Do While x < 10
        ActiveSheet.Cells(x, y - 2).Value = x * 2
        x = x + 1
    Loop

End Sub

' Always runs at least once
Sub DoLoopUntil()
    Dim x
    x = 1
    
    Dim newName
    Do
        newName = InputBox(Prompt:="Add another name", Title:="Names")
    Loop Until newName = ""

End Sub


Sub DoUntilLoop()
    Dim x
    x = 1
    
    Sheets(1).Select
    Do Until x > y
        ActiveSheet.Cells(x, y - 3).Value = x * 2
        x = x + 1
    Loop

End Sub

' essentially do-while-loop with different syntax
Sub WhileWend()
    Dim z
    z = 10
    
    While z > 0
        ActiveSheet.Cells(10, z) = (10 - z) ^ 3
        z = z - 1
    Wend
    
End Sub

' iterate through array
Sub ForEachNext()
    For Each x In Sheets
        MsgBox x.Name
    Next
End Sub

' define a series
Sub ForNext()
    For i = 1 To 10
        ActiveSheet.Cells(i, i).Value = "diagonal"
    Next
End Sub
