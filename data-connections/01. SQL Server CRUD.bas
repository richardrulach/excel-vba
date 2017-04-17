Attribute VB_Name = "Module1"
Const CONNECTION_STRING = "driver={SQL Server};server=.;Trusted_Connection=yes;database=test1"

Sub SQL_CreateDemoTable()
    Dim sSQL As String
    sSQL = "create table demo( " & _
            "demoId int identity(1,1) NOT NULL, " & _
            "name nvarchar(20) NOT NULL, " & _
            "description nvarchar(200) NOT NULL, " & _
            "inserted datetime DEFAULT GETDATE() " & _
            ")"
    
    Dim cnn As ADODB.Connection
    Set cnn = New ADODB.Connection
    
    cnn.ConnectionString = CONNECTION_STRING
    cnn.Open
    cnn.Execute sSQL
    
    cnn.Close
    Set cnn = Nothing
End Sub

Sub SQL_DropDemoTable()
    Dim sSQL As String
    sSQL = "drop table demo"
    
    Dim cnn As ADODB.Connection
    Set cnn = New ADODB.Connection
    
    cnn.ConnectionString = CONNECTION_STRING
    cnn.Open
    cnn.Execute sSQL
    
    cnn.Close
    Set cnn = Nothing
End Sub

Sub SQL_Create()
    Dim sSQL As String
    sSQL = "insert into demo([name],[description]) values ('mike','general'),('sam','games developer')"
    
    Dim cnn As ADODB.Connection
    Set cnn = New ADODB.Connection
    
    cnn.ConnectionString = CONNECTION_STRING
    cnn.Open
    cnn.Execute sSQL
    
    cnn.Close
    Set cnn = Nothing
End Sub


Sub SQL_Read()
    Dim cnn As ADODB.Connection
    Set cnn = New ADODB.Connection
    
    cnn.ConnectionString = CONNECTION_STRING
    cnn.Open
    
    Dim rsPubs As ADODB.Recordset
    Set rsPubs = New ADODB.Recordset
    
    With rsPubs
        .ActiveConnection = cnn
        .Open "SELECT * FROM demo"
        Sheets(1).Range("A2:D14").Clear
        Sheets(1).Range("A2:D14").CopyFromRecordset rsPubs
        
        .Close
    End With
    
    Set rsPubs = Nothing
    
    cnn.Close
    Set cnn = Nothing
    
End Sub


Sub SQL_Update()
    Dim iRow As Integer, iCol As Integer
    For iRow = 2 To GetRowCount()
        Dim sql As String
        sql = "update demo set name = '" & Sheets(1).Cells(iRow, 2).Value & _
            "', description = '" & Sheets(1).Cells(iRow, 3).Value & _
            "' where demoid = " & Sheets(1).Cells(iRow, 1).Value
        Call RunSQL(sql)
    Next
End Sub

Sub SQL_Delete()
    RunSQL "delete demo"
End Sub

Sub SQL_Truncate()
    RunSQL "truncate table demo"
End Sub


Sub RunSQL(sql)
    Dim cnn As ADODB.Connection
    Set cnn = New ADODB.Connection
    cnn.ConnectionString = CONNECTION_STRING
    cnn.Open
    cnn.Execute sql
    cnn.Close
    Set cnn = Nothing
End Sub

Function GetColumnCount()
    Dim colCount As Integer, bFound As Boolean
    colCount = 0
    bFound = False
    
    Do While bFound = False
        If Len(CStr(Trim(Sheets(1).Cells(1, colCount + 1).Value))) = 0 Then
            bFound = True
        Else
            colCount = colCount + 1
        End If
    Loop

    GetColumnCount = colCount
End Function


Function GetRowCount()
    Dim rowCount As Integer, bFound As Boolean
    rowCount = 0
    bFound = False
    
    Do While bFound = False
        If Len(CStr(Trim(Sheets(1).Cells(rowCount + 1, 1).Value))) = 0 Then
            bFound = True
        Else
            rowCount = rowCount + 1
        End If
    Loop

    GetRowCount = rowCount
End Function

