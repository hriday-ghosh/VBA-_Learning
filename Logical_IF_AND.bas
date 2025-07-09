Attribute VB_Name = "Module1"
Sub if_with_AND()

    ' Declare variables to store product name and sales quantity
    Dim Product_Name As String
    Dim Sale_Quantity As Integer

    ' Assign product name from cell C2
    Product_Name = Sheet1.Range("C2")
    
    ' Assign sales quantity from cell D2
    Sale_Quantity = Sheet1.Range("D2")

    ' Check if the product is "Headphone" and sales are 34 or more
    If Product_Name = "Headphone" And Sale_Quantity >= 34 Then
        ' If both conditions are true, mark as Good Seller
        Sheet1.Range("E2") = "Good Seller"
    Else
        ' If not, mark as Not Good
        Sheet1.Range("E2") = "Not Good"
    End If

End Sub

