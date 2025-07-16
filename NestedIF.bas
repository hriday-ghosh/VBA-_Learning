Attribute VB_Name = "Module1"
Option Explicit
' This forces you to declare all variables before using them.
' It helps avoid errors due to misspelled variable names.

Sub NestedIF()
' Declares the start of a macro named "NestedIF".

Dim order As Integer
' Declares a variable named "order" that will hold an Integer value.
' This variable is used to store the numeric value of an order.

Dim Category As String
' Declares a variable named "Category" that will hold a text (String) value.
' This variable will be used to store the result category (A, B, or C).

order = Sheet1.Range("B2").Value
' Reads the value from cell B2 of the worksheet named "Sheet1" and stores it in the variable "order".

If order >= 100 Then
' Checks if the value of "order" is greater than or equal to 100.
' If this condition is True, then:

    Category = "A"
    ' The variable "Category" is assigned the value "A".

ElseIf order >= 90 Then
' If the first condition is False, then checks if "order" is greater than or equal to 90.
' If this is True:

    Category = "B"
    ' Assigns the value "B" to the variable "Category".

Else
' If neither of the above conditions are True, this block is executed:

    Category = "C"
    ' Assigns the value "C" to the variable "Category".

End If
' Ends the If...ElseIf...Else decision structure.

Sheet1.Range("C2") = Category
' Writes the value of "Category" (A, B, or C) into cell C2 of Sheet1.

End Sub
' Ends the Subroutine.

