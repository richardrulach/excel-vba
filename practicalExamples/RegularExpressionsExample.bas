Attribute VB_Name = "Module1"


Function GetFirstMatch(source As String, pattern As String) As String
    
    Dim regex As RegExp
    Dim s As String
    Dim matches As MatchCollection
    Dim match As match
    
    Set regex = New RegExp
        
    regex.MultiLine = True
    regex.pattern = pattern
    regex.Global = True
    Set matches = regex.Execute(source)
    
    If matches.Count = 1 Then
        For Each match In matches
            s = match.Value
        Next
    ElseIf matches.Count = 0 Then
        s = "No match"
    Else
        s = "Too many matches"
    End If
    
    Set matches = Nothing
    Set match = Nothing
    Set regex = Nothing
    
    GetFirstMatch = s
    
End Function


Function GetAllMatches(source As String, pattern As String) As String
    
    Dim regex As RegExp
    Dim s As String
    Dim matches As MatchCollection
    Dim match As match
    
    Set regex = New RegExp
        
    regex.MultiLine = True
    regex.pattern = pattern
    regex.Global = True
    Set matches = regex.Execute(source)
    
    For Each match In matches
        s = s & match.Value & " " & Chr(10)
    Next
    
    Set matches = Nothing
    Set match = Nothing
    Set regex = Nothing
    
    GetAllMatches = s
    
End Function


Function GetFirstPostcode(s As String) As String
    GetFirstPostcode = GetFirstMatch(s, "[A-Z]{2}[0-9]{1,2} [0-9]{1,2}[A-Z]{2}")
End Function


Function GetFirstEmailAddress(s As String) As String
    GetFirstEmailAddress = GetFirstMatch(s, "\b[\w-\.]{1,}\@([\da-zA-Z-]{1,}\.){1,}[\da-zA-Z-]{2,3}\b")
End Function


Function GetAllPostcodes(s As String) As String
    GetAllPostcodes = GetAllMatches(s, "[A-Z]{2}[0-9]{1,2} [0-9]{1,2}[A-Z]{2}")
End Function


Function GetAllEmailAddresses(s As String) As String
    GetAllEmailAddresses = GetAllMatches(s, "\b[\w-\.]{1,}\@([\da-zA-Z-]{1,}\.){1,}[\da-zA-Z-]{2,3}\b")
End Function



