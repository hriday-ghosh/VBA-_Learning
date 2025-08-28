Attribute VB_Name = "Module1"
Sub doUntil()

    ' This loop will keep running UNTIL the ActiveCell (currently selected cell) is blank ("")
    Do Until ActiveCell.Value = ""
    
        ' ActiveCell.Offset(1, 0).Select means:
        ' Move the selection to another cell based on "rows" and "columns" shift
        ' Offset(RowShift, ColumnShift)
        ' (1, 0) means move 1 row DOWN, and 0 column sideways (so stay in the same column)
        ' Example: If you are at A1 this moves to A2 then A3 then A4 … and so on
        
        ActiveCell.Offset(1, 0).Select

    Loop
    
End Sub

