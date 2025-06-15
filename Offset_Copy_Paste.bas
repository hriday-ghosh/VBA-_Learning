Attribute VB_Name = "Module1"
Sub SaveDate()
    Sheets("Data").Select
    ' Select the sheet called "Data"

    Range("A2").Select
    ' Select cell A2 on "Data"

    Range(Selection, Selection.End(xlToRight)).Select
    ' From A2, select all cells to the right until the last filled cell in that row

    Range(Selection, Selection.End(xlDown)).Select
    ' Then from that range, select all cells downward until the last filled row

    Selection.Copy
    ' Copy all these selected cells

    Sheets("Save").Select
    ' Select the sheet called "Save"

    ActiveSheet.Paste
    ' Paste the copied cells here (starting from the currently selected cell)

    Selection.End(xlDown).Select
    ' Move the selection downward to the last filled cell in this column

    ActiveCell.Offset(1, 0).Select
    ' Move the selection 1 row downward from the last filled cell
    ' (Here there is a typo; it should be "ActiveCell" instead of "ActiveCel")

    Sheets("Data").Select
    ' Go back to "Data" sheet

End Sub

