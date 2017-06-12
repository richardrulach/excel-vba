Attribute VB_Name = "Module1"
Option Explicit

Public myBook As Workbook
Public mySheet As Worksheet


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   MAIN SUBROUTINES
'
'       - SINGLE FILE
'           - OPENS FILE AND ADDS VALIDATION
'           - DOES NOT SAVE (TO ENABLE MANUAL CHECKING)
'
'       - MULTIPLE FILES
'           - ALL XLSX FILES IN A SINGLE DIRECTORY
'           - OPEN
'           - ADD VALIDATION
'           - SAVE AND CLOSE
'

Sub Main_SingleFile()
    
    Dim sFileName As String
    sFileName = SelectFile
    If Len(sFileName) = 0 Then Exit Sub

    Application.ScreenUpdating = False
    
    If IsOpen(sFileName) Then
        Set myBook = Workbooks.Open(Right(sFileName, Len(sFileName) - InStrRev(sFileName, "\")))
    Else
        Set myBook = Workbooks.Open(sFileName)
    End If

    If myBook Is Nothing Then
        Debug.Print "FAILED TO OPEN THE WORKBOOK"
    Else
        Call UpdateLMC
    End If

    Application.ScreenUpdating = True
End Sub

Sub Main_AllFilesInFolder()
    Dim fldr As FileDialog
    Dim sItem As String
    
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    fldr.Title = "Select a Folder"
    fldr.AllowMultiSelect = False
    fldr.InitialFileName = Left(ThisWorkbook.FullName, InStrRev(ThisWorkbook.FullName, "\"))
    
    If fldr.Show = -1 Then
        Dim fso As Scripting.FileSystemObject
        Set fso = New FileSystemObject
        Dim f As Scripting.File
        
        For Each f In fso.GetFolder(fldr.SelectedItems(1)).Files
            If LCase(Right(f.Path, Len(f.Path) - InStrRev(f.Path, "."))) = "xlsx" Then
                If IsOpen(f.Path) Then
                    Set myBook = Workbooks.Open(Right(f.Path, Len(f.Path) - InStrRev(f.Path, "\")))
                Else
                    Set myBook = Workbooks.Open(f.Path)
                End If
                Call UpdateLMC
                myBook.Save
                myBook.Close
            End If
        Next
        
        Set f = Nothing
        Set fso = Nothing
    End If
    
    Set fldr = Nothing
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'
'   MAIN SUBROUTINE FOR APPLYING THE VALIDATION TO A FILE
'
'

Sub UpdateLMC()
    
    Dim thisWB As Workbook
    Set thisWB = ThisWorkbook
        
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' TARGET SPREADSHEET VALIDATION
    If HasNewSheetsAlready(myBook) Then
        Exit Sub
    Else
        Debug.Print "CHECKED: New sheets do not exist in target spreadsheet"
    End If
    
    If Not HasValidSheets(myBook) Then
        Exit Sub
    Else
        Debug.Print "CHECKED: All sheets required for the report are in target spreadsheet"
    End If
        
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' COPY SHEETS
    With thisWB
        .Sheets("Error Check").Copy Before:=myBook.Sheets("Front Page")
        .Sheets("Contents").Copy After:=myBook.Sheets("Front Page")
        .Sheets("Ex Summary").Copy After:=myBook.Sheets("Contents")
        .Sheets("Definitions").Copy After:=myBook.Sheets("Fees & Expenses")
    End With


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' CHECK IF ATP AND SPO HAVE VALID DATA SETS
    '   - These will be used later to fix the values in the rows relating to ATP and SPO
    '   - At time of development this onlly applied to the H1 pool
    Dim bATP As Boolean, bSPO As Boolean
    bATP = IsATPvalid
    bSPO = IsSPOvalid

    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ADD VALIDATION
    '   - Breaks down into two parts:
    '       - Error Check Sheet (most complicated formulas)
    '       - Executive Summary Sheet (set of fixed formulas which update as the report is populated)
    With myBook
    
        .Sheets("Error Check").Cells(3, 3).Formula = GetFormula("Loan Pool Summary", "Total Book #", "G")
        .Sheets("Error Check").Cells(3, 4).Formula = GetFormula("Loan Characteristics", "Total Book #", "H")
        .Sheets("Error Check").Cells(3, 5).Formula = GetFormula("Interest Rates", "Total Book #", _
            ColLetter(GetPos("Interest Rates", False, 3, "Change") - 1))
            
        .Sheets("Error Check").Cells(3, 6).Formula = GetFormula("Index Breakdown", "Total", "B")
        .Sheets("Error Check").Cells(3, 7).Formula = GetFormula("Current LTV", "Total Book #", _
            ColLetter(GetPos("Current LTV", False, 3, "Change") - 1))
        
        .Sheets("Error Check").Cells(3, 8).Formula = GetFormula("Headline Roll Rates", "Current Month Total", "T")
        .Sheets("Error Check").Cells(3, 9).Formula = GetFormula("Interest Rate Breakdown", "Total Book #", _
            ColLetter(GetPos("Interest Rate Breakdown", False, 4, "Change") - 1))
        
        .Sheets("Error Check").Cells(3, 10).Formula = GetSumFormula("Headline Roll Rates", "Total Improvements", "B", 4)
        
        .Sheets("Error Check").Cells(7, 3).Formula = GetFormula("Loan Pool Summary", "Total Book £", "G")
        .Sheets("Error Check").Cells(7, 4).Formula = GetFormula("Loan Pool Summary", "Closing Balance", "G")
        .Sheets("Error Check").Cells(7, 5).Formula = GetFormula("Loan Characteristics", "Total Book £", "H")
        .Sheets("Error Check").Cells(7, 6).Formula = GetFormula("Index Breakdown", "Total", "C")
        .Sheets("Error Check").Cells(7, 7).Formula = GetFormula("Current LTV", "Total Book #", _
                    ColLetter(GetPos("Current LTV", False, 3, "Change") - 1), 2)
        
        .Sheets("Error Check").Cells(11, 3).Formula = GetFormula("Arrears Workout", "Total", "B")
        .Sheets("Error Check").Cells(11, 4).Formula = GetSumFormula("Arrears Breakdown", "Total Arrears", "G", 3)
        .Sheets("Error Check").Cells(11, 5).Formula = GetRowSumFormula("Headline Roll Rates", "Current Month Total", 5, 11)
        
        .Sheets("Error Check").Cells(11, 6).Formula = GetFormula("Arrears by Loan Size", "Total Arrears", "G")
        .Sheets("Error Check").Cells(11, 7).Formula = GetFormula("RFA Summary - Analysis", "Total", "B")
        
        .Sheets("Error Check").Cells(11, 8).Formula = GetFormulaWithColCheck("RFA Summary - Loan Balance", "Total", 4, "Loan Count")
        
        .Sheets("Error Check").Cells(11, 9).Formula = GetFormula("Index Breakdown", "Total", "D")
        
        .Sheets("Error Check").Cells(14, 3).Formula = GetSumFormula("Arrears Breakdown (2)", "Total Arrears", "G", 3)
        .Sheets("Error Check").Cells(14, 4).Formula = "N/A"
        .Sheets("Error Check").Cells(14, 5).Formula = GetFormula("Arrears by Loan Size", "Total Arrears", "G", 2)
        .Sheets("Error Check").Cells(14, 6).Formula = GetFormula("RFA Summary - Analysis", "Total", "D")
        
        .Sheets("Error Check").Cells(14, 7).Formula = GetFormulaWithColCheck("RFA Summary - Loan Balance", "Total", 4, "Total")
        
        .Sheets("Error Check").Cells(14, 8).Formula = GetFormula("Index Breakdown", "Total", "E")
        
        .Sheets("Error Check").Cells(18, 3).Formula = GetFormula("PTP Summary", "Net at Month End", "G")
        .Sheets("Error Check").Cells(18, 4).Formula = GetFormula("PTP Summary (2)", "Total", "B")
        .Sheets("Error Check").Cells(18, 5).Formula = GetFormula("PTP Summary", "Total", "B")
        .Sheets("Error Check").Cells(18, 6).Formula = GetFormula("PTP Summary", "Total", "B", 2)
        
        If bATP Then
            .Sheets("Error Check").Cells(22, 3).Formula = GetFormula("ATP Summary", "Total", "D", 2)
            .Sheets("Error Check").Cells(22, 4).Formula = GetFormula("ATP Summary", "Total", "G")
            .Sheets("Error Check").Cells(22, 5).Formula = GetFormula("ATP Summary", "Total", "B", 3)
            .Sheets("Error Check").Cells(22, 6).Formula = GetFormula("ATP Summary (2)", "Total", "B")
            .Sheets("Error Check").Cells(22, 7).Formula = GetFormula("ATP Summary (2)", "Total", "B", 2)
        Else
            .Sheets("Error Check").Cells(22, 3).Value = 0
            .Sheets("Error Check").Cells(22, 4).Value = 0
            .Sheets("Error Check").Cells(22, 5).Value = 0
            .Sheets("Error Check").Cells(22, 6).Value = 0
            .Sheets("Error Check").Cells(22, 7).Value = 0
        End If
        
        .Sheets("Error Check").Cells(26, 3).Formula = GetFormula("Litigation Summary", "Total Litigation", "G")
        .Sheets("Error Check").Cells(26, 4).Formula = GetFormula("Litigation Summary", "Carried Forward", "G")
        .Sheets("Error Check").Cells(26, 5).Formula = GetFormula("Litigation Summary (2)", "Total", "G")
        
        If bSPO Then
            .Sheets("Error Check").Cells(30, 3).Formula = GetFormula("SPO Summary", "Total", "D", 2)
            .Sheets("Error Check").Cells(30, 4).Formula = GetFormula("SPO Summary", "Total", _
                                ColLetter(GetPos("SPO Summary", False, 5, "Change") - 1))
            .Sheets("Error Check").Cells(30, 5).Formula = GetFormula("SPO Summary", "Total", "B", 3)
            .Sheets("Error Check").Cells(30, 6).Formula = GetFormula("SPO Summary (2)", "Total", "B")
            .Sheets("Error Check").Cells(30, 7).Formula = GetFormula("SPO Summary (2)", "Total", "B", 2)
        Else
            .Sheets("Error Check").Cells(30, 3).Value = 0
            .Sheets("Error Check").Cells(30, 4).Value = 0
            .Sheets("Error Check").Cells(30, 5).Value = 0
            .Sheets("Error Check").Cells(30, 6).Value = 0
            .Sheets("Error Check").Cells(30, 7).Value = 0
        End If
            
        .Sheets("Error Check").Cells(34, 3).Value = GetFormula("Repossession Summary", "Live Repos", "B")
        .Sheets("Error Check").Cells(34, 4).Value = GetFormula("Repossession Summary", "Carried Forward", "G")
        
        
        
        .Sheets("Error Check").Cells(38, 3).Formula = GetFormula("Arrears Workout", "Total", "C")
        .Sheets("Error Check").Cells(38, 4).Formula = GetFormula("RFA Summary - Workout", "Total", "B")
        .Sheets("Error Check").Cells(38, 5).Formula = GetSumFormula("PTP Summary (2)", "1 - 1.99 pmts", "B", 13)
        
        If bATP Then
            .Sheets("Error Check").Cells(41, 3).Formula = GetFormula("Arrears Workout", "Total", "D")
            .Sheets("Error Check").Cells(41, 4).Formula = GetFormula("RFA Summary - Workout", "Total", "D")
            .Sheets("Error Check").Cells(41, 5).Formula = GetSumFormulaWithStop("ATP Summary (2)", "1 - 1.99 pmts", "B", "Total", iItem1:=1, iItem2:=2) + GetCustom_SPO("SPO Summary (2)")
        Else
            .Sheets("Error Check").Cells(41, 3).Value = 0
            .Sheets("Error Check").Cells(41, 4).Value = 0
            .Sheets("Error Check").Cells(41, 5).Value = 0
        End If
        
        
        .Sheets("Error Check").Cells(44, 3).Formula = GetFormula("Arrears Workout", "Total", "E")
        .Sheets("Error Check").Cells(44, 4).Formula = GetFormula("RFA Summary - Workout", "Total", "E")
        
        .Sheets("Error Check").Cells(47, 3).Formula = GetFormula("Arrears Workout", "Total", "F")
        .Sheets("Error Check").Cells(47, 4).Formula = GetFormula("RFA Summary - Workout", "Total", "F")
        
        .Sheets("Error Check").Cells(50, 3).Formula = GetFormula("Arrears Workout", "Total", "G")
        .Sheets("Error Check").Cells(50, 4).Formula = GetFormula("RFA Summary - Workout", "Total", "G")
        .Sheets("Error Check").Cells(50, 5).Formula = "=+C26"
        
        .Sheets("Error Check").Cells(53, 3).Formula = GetFormula("Arrears Workout", "Total", "I")
        .Sheets("Error Check").Cells(53, 4).Formula = GetFormula("RFA Summary - Workout", "Total", "I")
        .Sheets("Error Check").Cells(53, 5).Formula = GetFormula("Litigation Summary", "Total Possession", "G")
        
        .Sheets("Error Check").Cells(56, 3).Formula = GetFormula("Arrears Workout", "Total", "J")
        .Sheets("Error Check").Cells(56, 4).Formula = GetFormula("RFA Summary - Workout", "Total", "J")
        
        .Sheets("Error Check").Cells(59, 3).Formula = GetFormula("Arrears Workout", "Total", "l")
        .Sheets("Error Check").Cells(59, 4).Formula = GetFormula("RFA Summary - Workout", "Total", "K")
        
        
        ' EXECUTIVE SUMMARY PAGE FORMULAS
        .Sheets("Ex Summary").Cells(11, 2).Formula = "='Arrears Breakdown'!A3"
        .Sheets("Ex Summary").Cells(15, 2).Formula = "=+'Headline Roll Rates'!A3"
        .Sheets("Ex Summary").Cells(20, 2).Formula = "='Roll Rate History'!A3"
        .Sheets("Ex Summary").Cells(24, 2).Formula = "=+'Arrears Workout'!A3"
        .Sheets("Ex Summary").Cells(29, 2).Formula = "=+'Arrears Strategy Summary'!A3"
        .Sheets("Ex Summary").Cells(34, 2).Formula = "='RFA Summary - Analysis'!A3"
        .Sheets("Ex Summary").Cells(38, 2).Formula = "='Vulnerable Customers'!A3"
        .Sheets("Ex Summary").Cells(42, 2).Formula = "=+'PTP Summary'!A3"
        .Sheets("Ex Summary").Cells(48, 2).Formula = "=+'ATP Summary'!A3"
        .Sheets("Ex Summary").Cells(52, 2).Formula = "=+'Failed & Canx ATPs'!A3"
        .Sheets("Ex Summary").Cells(56, 2).Formula = "=+'Rehab Mods'!A3"
        .Sheets("Ex Summary").Cells(60, 2).Formula = "=+'Exit Mods'!A3"
        .Sheets("Ex Summary").Cells(64, 2).Formula = "=+'Litigation Summary'!A3"
        .Sheets("Ex Summary").Cells(69, 2).Formula = "=+'Litigation Summary (4)'!A3"
        .Sheets("Ex Summary").Cells(73, 2).Formula = "='SPO Summary'!A3"
        .Sheets("Ex Summary").Cells(77, 2).Formula = "=+'Failed & Canx SPOs'!A4"
        .Sheets("Ex Summary").Cells(81, 2).Formula = "=+'Repossession Summary'!A3"
        .Sheets("Ex Summary").Cells(85, 2).Formula = "=+'Redemption Summary'!A3"
        .Sheets("Ex Summary").Cells(89, 2).Formula = "=+'Fees & Expenses'!A3"
    
        ' FIX FORMATTING ISSUE SO THAT VALIDATION INDICATION CELLS ARE SHOWN FULL COLOUR (RED OR GREEN)
        .Sheets(1).Range("K3,H7,J11,I14,G18,H22,F26,H30,E34,F38,F41,E44,E47,F50,F53,E56,E59").Interior.Pattern = xlNone
        .Sheets(1).Range("K3,H7,J11,I14,G18,H22,F26,H30,E34,F38,F41,E44,E47,F50,F53,E56,E59").Interior.TintAndShade = 0
        .Sheets(1).Range("K3,H7,J11,I14,G18,H22,F26,H30,E34,F38,F41,E44,E47,F50,F53,E56,E59").Interior.PatternTintAndShade = 0
    
        .Sheets(1).Select
    End With

    Application.ScreenUpdating = True
    
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ATP AND SPO CHECKS
'       - INITIALLY FOR H1 POOLS
'

Function IsATPvalid() As Boolean
    Dim rng As Range
    Dim bResult As Boolean
    Set rng = myBook.Sheets("ATP Summary").Range("A1:Z1000").Find(What:="#div/0", LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:= _
        False, SearchFormat:=False)

    If rng Is Nothing Then
        bResult = True
    Else
        bResult = False
    End If
    
    IsATPvalid = bResult
End Function

Function IsSPOvalid() As Boolean
    Dim rng As Range
    Dim bResult As Boolean
    Set rng = myBook.Sheets("SPO Summary").Range("A1:Z1000").Find(What:="#div/0", LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:= _
        False, SearchFormat:=False)

    If rng Is Nothing Then
        bResult = True
    Else
        bResult = False
    End If
    
    IsSPOvalid = bResult
End Function


    
        
Function GetCustom_SPO(sSheet As String)
    Dim iRowStart As Integer, iRowEnd As Integer, i As Integer
    iRowStart = GetPos(sSheet, True, 1, "SPOs by MIA")
    iRowEnd = GetPos(sSheet, True, 1, "Total", 2)

    Dim iSum As Integer
    iSum = 0
    
    For i = iRowStart + 1 To iRowEnd - 1
        Debug.Print "Val:"; myBook.Sheets(sSheet).Cells(i, 1).Value
        If IsNumeric(Left(myBook.Sheets(sSheet).Cells(i, 1).Value, 1)) Then
            Debug.Print "is numeric"
            If CInt(Left(myBook.Sheets(sSheet).Cells(i, 1).Value, 1)) >= 1 Then
                iSum = iSum + myBook.Sheets(sSheet).Cells(i, 2).Value
            End If
        End If
    Next

    GetCustom_SPO = "+" & CStr(iSum)
End Function
        
        
Function GetRowSumFormula(sSheet As String, sLookupInA As String, iColumnToStart, iNumberOfItems As Integer)
    Dim iRowStart As Integer
    iRowStart = GetPos(sSheet, True, 1, sLookupInA)
    GetRowSumFormula = "=SUM('" & sSheet & "'!E" & CStr(iRowStart) & ":S" & CStr(iRowStart) & ")"
End Function


Function GetFormula(sSheet As String, sLookupInA As String, sColumn As String, Optional iItem As Integer = 1)
    Dim sRow As String
    sRow = CStr(GetPos(sSheet, True, 1, sLookupInA, iItem))
    GetFormula = "=+'" & sSheet & "'!" & sColumn & sRow
End Function


Function GetSumFormula(sSheet As String, sLookupInA As String, sColumn As String, iNumberOfItems As Integer, Optional iItem As Integer = 1, Optional FirstChar = "=")
    Dim sRow As String
    sRow = CStr(GetPos(sSheet, True, 1, sLookupInA, iItem))
    GetSumFormula = FirstChar & "SUM('" & sSheet & "'!" & sColumn & sRow & ":" & sColumn & CStr(CInt(sRow) + iNumberOfItems - 1) & ")"
End Function


Function GetSumFormulaWithStop(sSheet As String, sLookupInA As String, sColumn As String, sStopText As String, Optional iItem1 As Integer = 1, Optional iItem2 As Integer = 1, Optional FirstChar = "=")
    Dim sRow As String, sRowEnd As String
    sRow = CStr(GetPos(sSheet, True, 1, sLookupInA, iItem1))
    sRowEnd = CStr(GetPos(sSheet, True, 1, sStopText, iItem2) - 1)
    GetSumFormulaWithStop = FirstChar & "SUM('" & sSheet & "'!" & sColumn & sRow & ":" & sColumn & sRowEnd & ")"
End Function


Function GetFormulaWithColCheck(sSheet As String, sLookupInA As String, iRowForColumnLookup As Integer, sColumnLookup As String, Optional iItem As Integer = 1)
    Dim sRow As String, sColumn As String, i As Integer
    sColumn = ColLetter(GetPos(sSheet, False, iRowForColumnLookup, sColumnLookup))
    sRow = GetPos(sSheet, True, 1, sLookupInA, iItem)
    GetFormulaWithColCheck = "=+'" & sSheet & "'!" & sColumn & sRow
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   DIRECTORY FUNCTIONS
'
'       - SELECT FILE
'           - DISPLAYS DIALOG TO GET SINGLE FILE
'
'       - IS OPEN
'           - CHECKS IF A FILE IS OPEN IN CURRENT WORKBOOKS

Function SelectFile()
    Dim intChoice As Integer
    Dim strPath As String
    strPath = ""
    
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
    Application.FileDialog(msoFileDialogOpen).InitialFileName = Left(ThisWorkbook.FullName, InStrRev(ThisWorkbook.FullName, "\"))
    
    intChoice = Application.FileDialog(msoFileDialogOpen).Show
    If intChoice <> 0 Then strPath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
    
    SelectFile = strPath
End Function


Function IsOpen(sPath)
    Dim wb As Workbook, bIsOpen As Boolean
    bIsOpen = False
    For Each wb In Workbooks
        If wb.FullName = sPath Then
            bIsOpen = True
        End If
        Debug.Print (wb.FullName)
    Next
    IsOpen = bIsOpen
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   FIND POSITION FUNCTION
'
'

Function GetPos(lSheet As String, bRow As Boolean, iDimension As Integer, sSearch As String, Optional iIndex As Integer = 1) As Integer
    Dim myRange As Range, iCount As Integer, iPos As Integer
    
    On Error GoTo ErrHandler
    
    iPos = 0
    For iCount = 1 To Limit(lSheet, bRow)
        If myBook.Sheets(lSheet).Cells(IIf(bRow, iCount, iDimension), IIf(bRow, iDimension, iCount)).Value = sSearch Then
            iPos = iCount
            If iIndex = 1 Then
                Exit For
            Else
                iIndex = iIndex - 1
            End If
        End If
    Next
    
    GetPos = iPos
    Exit Function
    
ErrHandler:
    Call PrintError("GetPos")
    Debug.Print "lSheet: " & lSheet
    Debug.Print "iCount: " & CStr(iCount)
    Call PrintErrorClose
    Resume Next
    
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   UTILITY SUBROUTINES / FUNCTIONS
'
'

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


Function Limit(sSheet As String, bRow As Boolean)
    If bRow Then
        Limit = myBook.Worksheets(sSheet).UsedRange.Rows.Count
    Else
        Limit = myBook.Worksheets(sSheet).UsedRange.Columns.Count
    End If
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   VALIDATION FUNCTIONS
'
'

Function HasNewSheetsAlready(wb As Workbook) As Boolean
    
    Dim ws As Worksheet
    Dim bHasAllSheets As Boolean
    Dim bSheetsAlready
    bSheetsAlready = False
    
    Dim d As Scripting.Dictionary
    Set d = New Dictionary
    
    d.Add "Error Check", 1
    d.Add "Contents", 1
    d.Add "Ex Summary", 1
    d.Add "Definitions", 1
    
    On Error Resume Next
    
    Dim sCheckName As String
    Dim sName As Variant
    
    For Each sName In d.Keys
        sCheckName = ""
        sCheckName = wb.Sheets(sName).Name
        If sCheckName <> "" Then
            Debug.Print "Already has sheet: " & sCheckName
            bSheetsAlready = True
        End If
    Next

    On Error GoTo 0

    HasNewSheetsAlready = bSheetsAlready
End Function


Function HasValidSheets(wb As Workbook) As Boolean
    
    Dim ws As Worksheet
    Dim bHasAllSheets As Boolean
    bHasAllSheets = True
    
    Dim d As Scripting.Dictionary
    Set d = New Dictionary
    
    d.Add "Front Page", 1
    d.Add "Loan Pool Summary", 1
    d.Add "Loan Characteristics", 1
    d.Add "Interest Rates", 1
    d.Add "Index Breakdown", 1
    d.Add "Teaser Rate", 1
    d.Add "Interest Rate Breakdown", 1
    d.Add "Current LTV", 1
    d.Add "Repayment Type", 1
    d.Add "Arrears Breakdown", 1
    d.Add "Arrears Breakdown (2)", 1
    d.Add "Headline Roll Rates", 1
    d.Add "Roll Rate History", 1
    d.Add "Arrears by Loan Size", 1
    d.Add "Arrears Workout", 1
    d.Add "Arrears Strategy Summary", 1
    d.Add "Payment Methods - Arrears", 1
    d.Add "RFA Summary - Analysis", 1
    d.Add "RFA Summary - Workout", 1
    d.Add "RFA Summary - Loan Balance", 1
    d.Add "RFA Summary - Graph", 1
    d.Add "Vulnerable Customers", 1
    d.Add "PTP Summary", 1
    d.Add "PTP Summary (2)", 1
    d.Add "ATP Summary", 1
    d.Add "ATP Summary (2)", 1
    d.Add "Failed & Canx ATPs", 1
    d.Add "Rehab Mods", 1
    d.Add "Rehab Mods (2)", 1
    d.Add "Exit Mods", 1
    d.Add "Exit Mods (2)", 1
    d.Add "Litigation Summary", 1
    d.Add "Litigation Summary (2)", 1
    d.Add "Litigation Summary (3)", 1
    d.Add "Litigation Summary (4)", 1
    d.Add "SPO Summary", 1
    d.Add "SPO Summary (2)", 1
    d.Add "Failed & Canx SPOs", 1
    d.Add "LPA Summary", 1
    d.Add "Repossession Summary", 1
    d.Add "Repo Timeline", 1
    d.Add "Repo Timeline (2)", 1
    d.Add "Price Acheived Analysis", 1
    d.Add "Loss Severity", 1
    d.Add "Cashflow Summary - WP", 1
    d.Add "Cashflow Summary (2) - WP", 1
    d.Add "Cashflow Summary (3) - WP", 1
    d.Add "Cashflow Summary - WOP", 1
    d.Add "Cashflow Summary (2) - WOP", 1
    d.Add "Cashflow Summary (3) - WOP", 1
    d.Add "Redemption Summary", 1
    d.Add "Fees & Expenses", 1
    
    On Error GoTo ErrHandler:
    
    Dim sCheckName As String
    Dim sName As Variant
    
    For Each sName In d.Keys
        sCheckName = wb.Sheets(sName).Name
    Next

    HasValidSheets = bHasAllSheets
    Exit Function
ErrHandler:
    Debug.Print "Could not find sheet: " & sName
    bHasAllSheets = False
    
    Resume Next

End Function

