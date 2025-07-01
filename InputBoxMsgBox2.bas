Attribute VB_Name = "Module1"
Sub Input_BOXX()

    Dim Name As String
    Dim Age As Integer
   
    ' Get user input
    Name = InputBox("Enter your name")
    Age = InputBox("Enter your age") ' Corrected to prompt for age

    Sheet("SAVE").Select
   
    ' Place the data on the ranges in the "SAVE" sheet
    Range("C1") = Name
    Range("C2") = Age

    ' Notify the user
    MsgBox "Data Updated."

End Sub
