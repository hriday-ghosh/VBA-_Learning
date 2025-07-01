Attribute VB_Name = "Module1"
Sub COPY_DATE()
    ' Start the macro named COPY_DATE

    Sheets("Data").Select
    ' Go to the "Data" sheet

    Range("A1").Select
    ' Click on cell A1

    Range(Selection, Selection.End(xlToRight)).Select
    ' Select from A1 to the last column with data in row 1

    Range(Selection, Selection.End(xlDown)).Select
    ' Expand the selection down to the last row with data

    Selection.Copy
    ' Copy the selected data

    Sheets("SAVE").Select
    ' Go to the "SAVE" sheet

    Range("A1").Select
    ' Click on cell A1

    ActiveSheet.Paste
    ' Paste the data starting in A1

    Range("A1").Select
    ' Go back to A1

    Selection.End(xlDown).Select
    ' Go to the last row of pasted data

    ActiveCell.Offset(1, 0).Select
    ' Move one row below
   
    'Code for PATA Sheet Below
    Sheets("PATA").Select
    ' Switch to the "PATA" sheet

    Range("A1").Select
    ' Click on cell A1

    ActiveCell.Offset(1, 0).Select
    ' Move one row below

    Range(Selection, Selection.End(xlToRight)).Select
    ' Select from that cell to the end of the row

    Range(Selection, Selection.End(xlDown)).Select
    ' Select all rows of that data

    Selection.Copy
    ' Copy the data

    Sheets("SAVE").Select
    ' Go back to "SAVE" sheet

    ActiveSheet.Paste
    ' Paste the copied data below the previous one

End Sub
