Attribute VB_Name = "Module1"
'=========================================================
' VBA LOOPS - BEGINNER NOTES
'=========================================================

'#########################################################
' 1. FOR...NEXT LOOP
'#########################################################
'Use:
'Jokhon age thekei jano kotobar loop cholbe.

Dim i As Integer

For i = 1 To 5

    MsgBox i

Next i

'Output:
'1
'2
'3
'4
'5


'#########################################################
' 2. FOR EACH...NEXT LOOP
'#########################################################
'Use:
'Cells, Sheets, Workbook iterate korar jonno.

Dim Cell As Range

For Each Cell In Range("A1:A5")

    MsgBox Cell.Value

Next Cell


'#########################################################
' 3. DO WHILE LOOP
'#########################################################
'Use:
'Condition TRUE thaka porjonto loop cholbe.

Dim x As Integer

x = 1

Do While x <= 5

    MsgBox x

    x = x + 1

Loop


'#########################################################
' 4. DO UNTIL LOOP
'#########################################################
'Use:
'Condition FALSE thaka porjonto loop cholbe.

Dim y As Integer

y = 1

Do Until y > 5

    MsgBox y

    y = y + 1

Loop


'#########################################################
' 5. DO...LOOP WHILE
'#########################################################
'Use:
'Age ekbar run korbe
'Tarpor condition check korbe.

Dim z As Integer

z = 1

Do

    MsgBox z

    z = z + 1

Loop While z <= 5


'#########################################################
' 6. DO...LOOP UNTIL
'#########################################################
'Use:
'Age ekbar run korbe
'Tarpor condition check korbe.

Dim a As Integer

a = 1

Do

    MsgBox a

    a = a + 1

Loop Until a > 5


'#########################################################
' 7. NESTED LOOP
'#########################################################
'Use:
'Loop er vitore abar loop.

Dim RowNo As Integer
Dim ColNo As Integer

For RowNo = 1 To 3

    For ColNo = 1 To 2

        MsgBox "Row " & RowNo & _
               " Column " & ColNo

    Next ColNo

Next RowNo


'#########################################################
' 8. EXIT FOR
'#########################################################
'Use:
'Condition fulfill hole loop bondho.

Dim j As Integer

For j = 1 To 10

    If j = 5 Then Exit For

    MsgBox j

Next j


'#########################################################
' 9. EXIT DO
'#########################################################
'Use:
'Condition fulfill hole Do Loop bondho.

Dim k As Integer

k = 1

Do While k <= 10

    If k = 5 Then Exit Do

    MsgBox k

    k = k + 1

Loop


'#########################################################
' 10. STEP
'#########################################################
'Use:
'Increment ba decrement.

'Increment by 2

Dim n As Integer

For n = 2 To 10 Step 2

    MsgBox n

Next n


'Decrement

Dim m As Integer

For m = 10 To 1 Step -1

    MsgBox m

Next m


'=========================================================
' REAL EXCEL EXAMPLES
'=========================================================

'Example 1
'Color Blank Cells

Dim C As Range

For Each C In Range("A2:A20")

    If C.Value = "" Then

        C.Interior.Color = vbYellow

    End If

Next C


'---------------------------------------------------------

'Example 2
'Delete Blank Rows

Dim R As Long

For R = 100 To 2 Step -1

    If Cells(R, "A").Value = "" Then

        Rows(R).Delete

    End If

Next R


'---------------------------------------------------------

'Example 3
'Find First Blank Row

Dim RowNum As Long

For RowNum = 2 To 100

    If Cells(RowNum, "A").Value = "" Then

        MsgBox "Blank Row Found : " & RowNum

        Exit For

    End If

Next RowNum


'=========================================================
' INTERVIEW CHEAT SHEET
'=========================================================

'For...Next
'? Fixed number of iterations

'For Each...Next
'? Cells, Sheets, Workbook

'Do While
'? Loop until condition becomes False

'Do Until
'? Loop until condition becomes True

'Nested Loop
'? Table/Matrix Processing

'Exit For
'? Stop For Loop

'Exit Do
'? Stop Do Loop

'Step
'? Increase/Decrease by custom value
'=========================================================

