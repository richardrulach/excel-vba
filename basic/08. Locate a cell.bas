Attribute VB_Name = "Module1"
Option Explicit

Public myBook As Workbook
Public mySheet As Worksheet

Sub TestFindData()
    Set myBook = ThisWorkbook
    Set mySheet = myBook.Sheets("Sheet1")
    
    Debug.Print CStr(GetPos(False, 1, "london", iIndex:=2))
    Debug.Print CStr(GetPos(True, 1, "jim", iIndex:=4))

    Debug.Print CStr(ColInt("A"))
    Debug.Print CStr(ColInt("B"))
    Debug.Print CStr(ColInt("E"))
    Debug.Print CStr(ColInt("AB"))
    Debug.Print CStr(ColInt("AAA"))

End Sub


Function GetPos(bRow As Boolean, iDimension As Integer, sSearch As String, Optional iIndex As Integer) As Integer
    Dim myRange As Range, sSet As String, iCount As Integer, sFirst As String
    
    On Error GoTo ErrHandler
    sSet = IIf(bRow, ColLetter(iDimension), CStr(iDimension)) & ":" & IIf(bRow, ColLetter(iDimension), CStr(iDimension))
    
    For iCount = 1 To iIndex
        If iCount = 1 Then
            Set myRange = mySheet.Range(sSet).Find(What:=sSearch, LookIn:=xlValues)
            If myRange Is Nothing Then
                GetPos = 0
                Exit Function
            Else
                sFirst = myRange.Address
            End If
        Else
            Set myRange = mySheet.Range(sSet).FindNext(myRange)
        End If
        
        If sFirst = myRange.Address And iCount <> 1 Then
            GetPos = 0
            Exit Function
        End If
    Next
    GetPos = IIf(bRow, myRange.Row, myRange.Column)
    Exit Function
    
ErrHandler:
    Call PrintError("GetPos")
    Debug.Print "sSet: " & sSet
    Debug.Print "iCount: " & CStr(iCount)
    Debug.Print "sFirst: " & sFirst
    Call PrintErrorClose
    Resume Next
    
End Function

Sub PrintError(sFunc As String)
    Debug.Print "***********************************"
    Debug.Print "ERRROR in " & sFunc
    Debug.Print "***********************************"
    Debug.Print "Number: " & Err.Number
    Debug.Print "Description: " & Err.Description
End Sub

Sub PrintErrorClose()
    Debug.Print "***********************************"
    Debug.Print " "
End Sub


Function ColLetter(i As Integer) As String
    ColLetter = Replace(Cells(1, i).Address(False, False), "1", "")
End Function

Function ColInt(s As String) As String
    ColInt = Range(s & "1").Column
End Function

