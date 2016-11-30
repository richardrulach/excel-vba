''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Author:       RR
' Created:      30/11/2016
' Description:  Demonstrate the use of conditionals in VBA
'               1 - IF THEN ELSE
'               2 - SELECT CASE
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub IF_THEN_ELSE()

    Call CLEAR_ALL
    
    Dim newValue As Object

    Sheets(1).Select
    If 2 = 4 Then ActiveSheet.Cells(1, 1).Value = "Home"
    
    If 1 = True Then ActiveSheet.Cells(1, 2).Value = "One is true"

    If IsNull(newValue) Then ActiveSheet.Cells(1, 3).Value = "Object is null if not initialised"
    
    If newValue Is Nothing Then ActiveSheet.Cells(1, 3).Value = "Object is nothing if not initialised"

    Dim name
    name = "Tom"

    If LCase(name) = "tom" Then
        ActiveSheet.Cells(2, 1).Value = "Is tom"
    Else
        ActiveSheet.Cells(2, 1).Value = "Is NOT tom"
    End If

    If 1 = 1 Then
    
    ElseIf 1 = 2 Then
    
    Else
    
    End If


End Sub

Sub CASE_STMT()

    Call CLEAR_ALL
    
    Dim x
    x = 12
    
    Select Case x
    Case 10
        ActiveSheet.Cells(3, 1) = 10
    Case 11
        ActiveSheet.Cells(3, 1) = 11
    Case 12
        ActiveSheet.Cells(3, 1) = 12
    Case Else
        ActiveSheet.Cells(3, 1) = 13
    End Select
    
End Sub


Sub CLEAR_ALL()
    Sheets(1).Range("a1:z26").ClearContents
End Sub
