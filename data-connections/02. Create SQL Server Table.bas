Attribute VB_Name = "Module1"
Const CONNECTION_STRING = "driver={SQL Server};server=.;Trusted_Connection=yes;database=test1"

Sub TestConnection()
    Dim cnn As ADODB.Connection
    Set cnn = New ADODB.Connection
    
    cnn.ConnectionString = CONNECTION_STRING
    cnn.Open

    MsgBox cnn.State
    
    
    cnn.Close
    Set cnn = Nothing
End Sub

Sub CreateAndLoadTable()
    
    Dim sTableName As String
    sTableName = InputBox("Enter new table name")
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' CREATE THE TABLE
    Dim sql_a As String, sql_b As String, sql_cols As String
    
    sql_a = "create table " & sTableName & "( " & _
            sTableName & "ID    int identity(1,1)   NOT NULL, " & _
            "rowNumber       int                 NOT NULL, "
    
    sql_b = ")"
    
    ' GET SIZE OF THE DATA
    Dim numCols As Integer, numRows As Integer, rowLen As Integer
    numCols = GetColumnCount()
    numRows = GetRowCount()
    
    ' CHECK ROW LENGTH AND WRITE COLUMN DEFINITION
    For iCol = 1 To numCols
                
        ' MINIMUM ROW SIZE
        rowLen = 100
        
        For iRow = 2 To numRows
            If rowLen < Len(Sheets(1).Cells(iRow, iCol).Value) Then
                rowLen = Len(Sheets(1).Cells(iRow, iCol).Value)
            End If
        Next
        
        If rowLen + 100 < 4001 Then
            sql_cols = sql_cols + "[" + Replace(Sheets(1).Cells(1, iCol).Value, vbLf, "") + "]" + " nvarchar(" & CStr(rowLen + 100) & ") NULL,"
        Else
            sql_cols = sql_cols + "[" + Replace(Sheets(1).Cells(1, iCol).Value, vbLf, "") + "]" + " ntext NULL,"
        End If
    Next
    
    If Len(sql_cols) > 0 Then sql_cols = Left(sql_cols, Len(sql_cols) - 1)
    
    Call RunSQL(sql_a & sql_cols & sql_b)
    ' END CREATING THE TABLE
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    

    Dim sqlInsert As String, newCount As Integer, newBatch As String
    sqlInsert = "INSERT INTO " & sTableName & " Values "
    
    For iRow = 2 To numRows
        If newCount > 999 Then
            newCount = 1
            newBatch = Left(newBatch, Len(newBatch) - 1)
            Call RunSQL(sqlInsert & newBatch)
            newBatch = ""
        End If
                
        Dim newRow As String
        newRow = "(" & CStr(iRow) & ","
    
        For iCol = 1 To numCols
            newRow = newRow & "'" & rep(Sheets(1).Cells(iRow, iCol).Value) & "',"
        Next
        
        newRow = Left(newRow, Len(newRow) - 1) & "),"
        newBatch = newBatch & newRow
            
        newCount = newCount + 1
    Next
    
    ' RUN THE FINAL BATCH
    newBatch = Left(newBatch, Len(newBatch) - 1)
    Call RunSQL(sqlInsert & newBatch)


End Sub

Sub RunSQL(sql)
    Dim cnn As ADODB.Connection
    Set cnn = New ADODB.Connection
    
    cnn.ConnectionString = CONNECTION_STRING
    
    'MsgBox sql
    
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

Function rep(s)
    rep = Replace(s, "'", "''")
End Function

