' Launches the calculator function and does a calculation on it...
Sub Calc()
    Dim ReturnValue As Long
    Dim I
    ReturnValue = CInt(Shell("CALC.EXE", 1))    ' Run Calculator.
    
    ' Needs a two second pause before it is ready
    Application.Wait (Now + TimeValue("0:00:02"))
    AppActivate ReturnValue
    
    For I = 1 To 100    ' Set up counting loop.
        SendKeys I & "{+}", True    ' Send keystrokes to Calculator
    Next I    ' to add each value of I.
    SendKeys "=", True    ' Get grand total.
'    SendKeys "%{F4}", True    ' Send ALT+F4 to close Calculator.
End Sub
