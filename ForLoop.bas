Attribute VB_Name = "Module1"
Option Explicit
' This makes sure all variables are declared before using them — helps prevent spelling mistakes in variable names.

Sub ForLoop()
' Starts a new macro named "ForLoop".

Dim i As Integer
' Declares a variable "i" that will be used as a counter in the loop.

Dim total As Integer
' Declares a variable "total" to store the sum of numbers from the loop.

For i = 1 To 10 Step 2
' This is a For loop that starts with i = 1, goes up to 10, and increases by 2 every time.
' So it will run for i = 1, 3, 5, 7, 9.

    Sheet1.Range("a" & i).Value = i
    ' Puts the value of i into cells A1, A3, A5, A7, and A9 (because i increases by 2 each time).

    total = total + i
    ' Adds the current value of i to the total.
    ' For example, total = total + 1 in first loop, then +3, +5, etc.

Next
' Moves to the next loop step — increases i by 2 and repeats the code until i > 10.

Range("a11").Value = total
' After the loop ends, this writes the final total value into cell A11.
' Example: 1 + 3 + 5 + 7 + 9 = 25 ? A11 will show 25.

End Sub
' Ends the macro.

