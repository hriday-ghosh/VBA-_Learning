Attribute VB_Name = "Module1"
Option Explicit
'Forces you to declare all variables before using them. Helps avoid typos & bugs.

Sub DeleteBlankRows()
'This starts the macro named DeleteBlankRows

Dim i As Integer
'Variable "i" will be used as a counter in the loop (row number)
Dim Last_Row As Long
'Variable "Last_Row" will store the last used row number in column A

'--- Find the last used row in column A ---
Last_Row = Sheets("Data").Range("a" & Rows.Count).End(xlUp).Row
' Explanation:
' Sheets("Data") -> means we are working on the sheet named "Data"
' Rows.Count -> gives total number of rows in Excel (like 1048576)
' "a" & Rows.Count ? joins "A" with that number ? "A1048576"
' Range("A1048576").End(xlUp) ? goes upward from bottom until it finds a filled cell
' .Row -> gives that row number (example: 350 if data ends at row 350)

'Alternative way using Cells function:
'Last_Row = Sheets("Data").Cells(Rows.Count, 1).End(xlUp).Row

'--- Loop through all rows from 1 to the last row ---
For i = 1 To Last_Row Step 1
    ' Step 1 -> means move one row at a time (1, 2, 3, ...)

    'If the cell in column A (1st column) of the current row is blank
    If Sheets("Data").Cells(i, 1) = "" Then

        'Select the whole blank row
        Sheets("Data").Rows(i & ":" & i).Select
        ' Explanation:
        ' i & ":" & i -> joins the same row number twice, like "5:5" -> row 5 only

        'Delete the selected row and shift cells upward
        Selection.Delete shift:=xlUp
        ' Explanation:
        ' shift:=xlUp -> moves rows below up to fill the empty space

    End If

Next
'Moves to the next row and repeats the process

Sheets("Data").Cells(1, 1).Select
'After finishing, move cursor back to cell A1 for neatness

End Sub
'End of the macro
