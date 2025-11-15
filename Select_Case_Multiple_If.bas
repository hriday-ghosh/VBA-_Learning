Attribute VB_Name = "Module2"
Option Explicit
' Forces you to declare every variable (good practice)

Sub Look_Up()
' This is the start of your macro named Look_Up

Dim DATA As String
' Creates a variable called DATA which will store text (string)

DATA = Range("B41").Value
' Reads the value from cell B41 and stores it into DATA


' Select Case is used when you want to check many conditions for one value
Select Case DATA
' Here we are checking what value is stored in DATA


' ------------------------------- CASE 1 -------------------------------

Case "HHP"
' If DATA = "HHP", then this block will run

    If Range("B41").Value = "HHP" Then
    ' Extra checking (not required but still works)
        Range("C2:E32").Copy
        ' Copy the cells from C2 to E32
        Range("C42").PasteSpecial
        ' Paste the copied cells starting at C42
    End If


' ------------------------------- CASE 2 -------------------------------

Case "AV"
' If DATA = "AV", then this block will run

    If Range("B41").Value = "AV" Then
    ' Again double checking
        Range("H2:J32").Copy
        ' Copy the range H2 to J32
        Range("C42").PasteSpecial
        ' Paste starting at C42
    End If


' ------------------------------- CASE 3 -------------------------------

Case "WG"
' If DATA = "WG", then this block will run

    If Range("B41").Value = "WG" Then
        Range("M2:O32").Copy
        ' Copy the range M2 to O32
        Range("C42").PasteSpecial
        ' Paste starting at C42
    End If


' ------------------------------- CASE 4 -------------------------------

Case "eStore"
' If DATA = "eStore"

    If Range("B41").Value = "eStore" Then
        Range("R2:T32").Copy
        ' Copy R2 to T32
        Range("C42").PasteSpecial
        ' Paste at C42
    End If


' ------------------------------- CASE 5 (Special Case) -------------------------------

Case "Sales"
' If DATA = "Sales"

    If Range("B41").Value = "Sales" Then
        Range("W2:W32").Copy
        ' Copy column W (W2 to W32)
        Range("C42").PasteSpecial
        ' Paste into C42

        Range("X2:X32").Copy
        ' Copy column X
        Range("E42").PasteSpecial
        ' Paste into E42

        Range("D42:D72").Value = 0
        ' Set all cells in D42 to D72 to zero
    End If


End Select
' End of Select Case block

End Sub
' End of the macro

