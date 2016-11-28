''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Author:       RR
' Created:      28/11/2016
' Description:  Downloads a csv file to the My Documents folder then inserts it into the
'               the file (in this code into the second spreadsheet in the workbook.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub ResetCard()
    Range("B3:J3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Call LoadData
End Sub

Sub LoadData()
    
    Dim myURL As String

    ' PROVIDE THE URL FOR DOWNLOADING THE CSV FILE FROM
    myURL = ""
    
    Dim WinHttpReq As Object
    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", myURL, False
    WinHttpReq.send
    
    myURL = WinHttpReq.responseBody
    
    Dim MyDocsPath As String
    MyDocsPath = Environ$("USERPROFILE") & "\My Documents\DownloadedOutstandingTasks.csv"
    
    If WinHttpReq.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.responseBody
        oStream.SaveToFile MyDocsPath, 2 ' 1 = no overwrite, 2 = overwrite
        oStream.Close
        Call CSV_Import(MyDocsPath)
    End If
End Sub


Sub CSV_Import(strFile)
    Dim ws As Worksheet
    
    Set ws = ActiveWorkbook.Sheets(2)
    
    With ws.QueryTables.Add(Connection:="TEXT;" & strFile, Destination:=ws.Range("B3"))
         .TextFileParseType = xlDelimited
         .TextFileCommaDelimiter = True
         .TextFileStartRow = 2
         .Refresh
    End With
End Sub




