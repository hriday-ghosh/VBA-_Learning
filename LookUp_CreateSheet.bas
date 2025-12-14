Attribute VB_Name = "Module1"
'============================================================
' Macro 1: Look_Up
' Purpose: Based on the value in B41, copy a specific table
'          and paste it starting at C42.
'============================================================
Sub Look_Up()

    ' Declare a variable called DATA to store text
    Dim DATA As String
    
    ' Read the value from cell B41 (the selected category) and store it in DATA
    DATA = Range("B41").Value
    
    ' Check which category was selected in B41
    Select Case DATA
    
        ' If category is "HHP"
        Case "HHP"
            ' Double-check B41 is "HHP" (optional)
            If Range("B41").Value = "HHP" Then
                ' Copy cells C2 to E32
                Range("C2:E32").Copy
                ' Paste the copied data starting at C42
                Range("C42").PasteSpecial
            End If
            
        ' If category is "AV"
        Case "AV"
            If Range("B41").Value = "AV" Then
                ' Copy H2:J32 and paste at C42
                Range("H2:J32").Copy
                Range("C42").PasteSpecial
            End If
            
        ' If category is "WG"
        Case "WG"
            If Range("B41").Value = "WG" Then
                ' Copy M2:O32 and paste at C42
                Range("M2:O32").Copy
                Range("C42").PasteSpecial
            End If
            
        ' If category is "eStore"
        Case "eStore"
            If Range("B41").Value = "eStore" Then
                ' Copy R2:T32 and paste at C42
                Range("R2:T32").Copy
                Range("C42").PasteSpecial
            End If
            
        ' If category is "Sales"
        Case "Sales"
            If Range("B41").Value = "Sales" Then
                ' Copy W2:W32 and paste at C42
                Range("W2:W32").Copy
                Range("C42").PasteSpecial
                
                ' Copy X2:X32 and paste at E42
                Range("X2:X32").Copy
                Range("E42").PasteSpecial
                
                ' Fill D42:D72 with zero
                Range("D42:D72").Value = 0
            End If
            
    End Select

End Sub

'============================================================
' Macro 2: Delete
' Purpose: Clear the contents of the previously pasted table.
'============================================================
Sub Delete()

    ' Select cells C42:E72 and remove all data (values)
    ' Formatting, colors, and borders remain intact
    Range("C42:E72").ClearContents

End Sub

'============================================================
' Macro 3: Create_Sheet
' Purpose: Copy the last sheet of the workbook into a new workbook
'          and convert all formulas to values for CSV export.
'============================================================
Sub Create_Sheet()

    ' Declare variables
    
    ' Original workbook
    Dim wbSource As Workbook
    ' New workbook that will be created
    Dim wbNew As Workbook
    ' Last sheet in the source workbook
    Dim wsLast As Worksheet
    ' First sheet in the new workbook
    Dim wsNew As Worksheet
    ' Range to copy
    Dim rng As Range
    ' Date value (not currently used)
    Dim dt As Date
    ' Name of CSV file (not used currently)
    Dim fileName As String
    ' Folder path for saving (not used currently)
    Dim savePath As String

    ' Set reference to the current workbook
    Set wbSource = ThisWorkbook

    ' Identify the last worksheet in the workbook
    Set wsLast = wbSource.Sheets(wbSource.Sheets.Count)

    ' Copy the last sheet into a new workbook
    ' The Copy method automatically creates a new workbook
    wsLast.Copy

    ' Set reference to the new workbook
    Set wbNew = ActiveWorkbook

    ' Set reference to the first sheet in the new workbook
    Set wsNew = wbNew.Sheets(1)

    ' Activate the new sheet
    wsNew.Activate

    ' Move cursor to cell A1
    wsNew.Range("A1").Select

    ' Define the range A1:I15000 (covers almost all data)
    Set rng = wsNew.Range("A1:I15000")

    ' Copy the range
    rng.Copy

    ' Paste only the values back into the same range
    ' This removes any formulas but keeps the data
    rng.PasteSpecial Paste:=xlPasteValues

    ' Remove the dashed border around copied cells
    Application.CutCopyMode = False

    '============================================================
    ' The following lines are commented out, but would:
    ' 1. Read a date from B2
    ' 2. Build a CSV file name with the date
    ' 3. Define a folder path
    ' 4. Save the new workbook as CSV
    ' 5. Display a message box
    '============================================================
    ' dt = wsNew.Range("B2").Value
    ' fileName = "FC_NICE02B_102849782_SamsungPoland_" & Format(dt, "ddmmyy") & ".csv"
    ' wbNew.Windows(1).Caption = fileName
    ' savePath = "C:\Users\hriday.ghosh\OneDrive - Concentrix Corporation\Hriday\Samsung Poland\Upload Forecast Here\"
    ' Application.DisplayAlerts = False
    ' wbNew.SaveAs fileName:=savePath & fileName, FileFormat:=xlCSV
    ' Application.DisplayAlerts = True
    ' MsgBox "New CSV file saved as: " & savePath & fileName, vbInformation

End Sub

