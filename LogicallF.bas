Attribute VB_Name = "Module1"
Sub LogicalAND()

    Dim Sales As Integer
    Dim AHT As Integer

    ' Assigning values from Sheet1 cells
    Sales = Sheet1.Range("B2").Value
    AHT = Sheet1.Range("C2").Value

    ' Applying logical AND condition
    If Sales >= 100 And AHT <= 150 Then
        Sheet1.Range("D2").Value = "Good Performance"
    Else
        Sheet1.Range("D2").Value = "Needs Improvement"
    End If

End Sub

