Attribute VB_Name = "Module1"
Sub DoWhile()
' Start of the subroutine named "DoWhile"

Dim i As Integer
' Declare a variable "i" as an Integer to use as a counter

i = 1
' Initialize the counter variable "i" with value 1

Do While i <= 10
' Start a loop that will continue as long as "i" is less than or equal to 10

    Sheet1.Cells(i, 1).Value = 1
    ' Set the value of cells in column A (1st column) from row 1 to 10 to 1

    i = i + 1
    ' Increase the value of "i" by 1 each time the loop runs

Loop
' End of the loop — control goes back to "Do While" to check the condition again

End Sub
' End of the subroutine

