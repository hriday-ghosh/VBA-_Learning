Attribute VB_Name = "Module1"
Sub Input_BOX()

    Dim Name As String
    Dim Age As Integer
    Dim ws As Worksheet

    ' Set the worksheet to the sheet named "SAVE"
    Set ws = ThisWorkbook.Sheets("SAVE")

    ' Get user input
    Name = InputBox("Enter your name")
    Age = InputBox("Enter your age") ' Corrected to prompt for age

    ' Place the data on the ranges in the "SAVE" sheet
    ws.Range("C1").Value = Name
    ws.Range("C2").Value = Age

    ' Notify the user
    MsgBox "Data Updated."

End Sub

