Attribute VB_Name = "Module1"
Sub updated()

    ' ===== This macro reads shift data from the Roster sheet
    ' ===== matches it with Shift Database
    ' ===== and creates rows in Updated Schedule

    ' Turn off screen updating to avoid flickering
    Application.ScreenUpdating = False

    ' Turn off warning popups (like overwrite alerts)
    Application.DisplayAlerts = False

    ' Go to Roster sheet
    Sheets("Roster").Select

    ' Start from cell A3 (first agent ID)
    Range("A3").Select

    ' Z is used to move upward to get the date
    Z = 1

    ' Loop until Agent ID cell becomes blank
    Do Until ActiveCell.Value = ""

        ' Move to column C of the same row (shift column)
        Range("C" & ActiveCell.Row).Select

        ' Loop through shifts until an empty cell is found
        Do Until ActiveCell.Value = ""

            ' Store shift code (example: PTO, OFF, 600 etc.)
            Shiftx = ActiveCell.Value

            ' Store employee ID from column A
            IDX = Range("A" & ActiveCell.Row).Value

            ' Store employee name from column B
            nameX = Range("B" & ActiveCell.Row).Value

            ' Get the date from the row above
            DateX = ActiveCell.Offset(-Z, 0).Value

            ' Go to Shift Database sheet
            Sheets("Shift Database").Select

            ' Start searching shift from A2
            Range("A2").Select

            ' Find matching shift code in Shift Database
            Do Until ActiveCell.Value = Shiftx
                ActiveCell.Offset(1, 0).Select

                ' If shift not found, go back and skip
                If ActiveCell.Value = "" Then
                    GoTo Home
                End If
            Loop

            ' Move to column B of matched shift row
            Range("B" & ActiveCell.Row).Select

            ' ===== Special shift conditions =====

            ' If PTO shift, copy first 3 columns
            If Shiftx = "PTO" Then
                Range(Selection, Selection.Offset(0, 2)).Select
                Selection.Copy
                GoTo Home2
            End If

            ' If BH (Bank Holiday)
            If Shiftx = "BH" Then
                Range(Selection, Selection.Offset(0, 2)).Select
                Selection.Copy
                GoTo Home2
            End If

            ' If OFF day, skip copying
            If Shiftx = "OFF" Then
                GoTo Home3
            End If

            ' If shift duration = 600 minutes
            If Range("H" & ActiveCell.Row).Value = 600 Then
                Range(Selection, Selection.Offset(4, 2)).Select
                Selection.Copy
                GoTo Home2
            End If

            ' If shift duration = 700 minutes
            If Range("H" & ActiveCell.Row).Value = 700 Then
                Range(Selection, Selection.Offset(4, 2)).Select
                Selection.Copy
                GoTo Home2
            End If

            ' If shift duration = 400 minutes (09–13)
            If Range("H" & ActiveCell.Row).Value = 400 Then
                Range(Selection, Selection.Offset(2, 2)).Select
                Selection.Copy
                GoTo Home2
            End If

            ' If shift duration = 500 minutes (09–14)
            If Range("H" & ActiveCell.Row).Value = 500 Then
                Range(Selection, Selection.Offset(4, 2)).Select
                Selection.Copy
                GoTo Home2
            End If

            ' Example of custom exception – duration 620
            If Range("H" & ActiveCell.Row).Value = 620 Then
                Range(Selection, Selection.Offset(4, 2)).Select
                Selection.Copy
                GoTo Home2
            End If

            ' If Comp OFF shift
            If Shiftx = "Comp OFF" Then
                Range(Selection, Selection.Offset(0, 2)).Select
                Selection.Copy
                GoTo Home2
                'GoTo Home2
                'Stop what you are doing now and jump immediately to the line where Home: is written.
                'Emergency exit / End processing
                
            End If

            ' Default copy logic if no condition matched
            Range(Selection, Selection.Offset(6, 2)).Select
            Selection.Copy

Home2:
            
            
            ' ===== Paste data into Updated Schedule =====

            Sheets("Updated Schedule").Select

            ' Find first empty row in column A
            Range("A2").Select
            Do Until ActiveCell.Value = ""
                ActiveCell.Offset(1, 0).Select
            Loop

            ' Fill employee details
            Range("A" & ActiveCell.Row).Value = IDX
            Range("B" & ActiveCell.Row).Value = nameX
            Range("C" & ActiveCell.Row).Value = Format(DateX, "YYYYMMDD")
            Range("D" & ActiveCell.Row).Value = Format(DateX, "YYYYMMDD")

            ' Paste copied shift data
            Range("E" & ActiveCell.Row).PasteSpecial

            ' Move down to paste multiple rows if needed
            Range("E" & ActiveCell.Row + 1).Select
            Do Until ActiveCell = ""
                Range("A" & ActiveCell.Row).Value = IDX
                Range("B" & ActiveCell.Row).Value = nameX
                Range("C" & ActiveCell.Row).Value = Format(DateX, "YYYYMMDD")
                Range("D" & ActiveCell.Row).Value = Format(DateX, "YYYYMMDD")
                ActiveCell.Offset(1, 0).Select
            Loop

Home3:
            ' Go back to Roster and move to next shift column
            Sheets("Roster").Select
            ActiveCell.Offset(0, 1).Select

        Loop

        ' Move to next agent row
        Sheets("Roster").Select
        Range("A" & ActiveCell.Row).Select
        ActiveCell.Offset(1, 0).Select

        ' Increase date offset
        Z = Z + 1

    Loop

Home:
    ' Turn screen updating and alerts back on
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "Process is Done."

End Sub


Sub CreateSheet()

    ' ===== This macro creates a CSV file based on team name =====

    Dim wbSource As Workbook
    Dim wsLast As Worksheet
    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    Dim dtString As String
    Dim fileName As String
    Dim savePath As String
    Dim folderName As String
    Dim rosterSheet As Worksheet
    Dim teamName As String

    ' Reference current workbook
    Set wbSource = ThisWorkbook
    Set rosterSheet = wbSource.Sheets("Roster")

    ' Read team name from cell L5
    teamName = Trim(rosterSheet.Range("L5").Value)

    ' Decide folder based on team name
    Select Case teamName
        Case "BackOffice", "E-Promo", "Samsung HHP Cagliary ITA", _
             "HHP Tirana ALB", "Samsung T2 Cagliary", "VOC"
            folderName = teamName
        Case Else
            MsgBox "Invalid team name in Roster!", vbCritical
            Exit Sub
    End Select

    ' Get last worksheet
    Set wsLast = wbSource.Sheets(wbSource.Sheets.Count)

    ' Copy last sheet to new workbook
    wsLast.Copy
    Set wbNew = ActiveWorkbook
    Set wsNew = wbNew.Sheets(1)

    ' Delete column B
    wsNew.Columns("B").Delete

    ' Get date value from C2
    dtString = wsNew.Range("C2").Value

    ' Create CSV file name
    fileName = "Attendance_scheduling" & dtString & ".csv"

    ' Build save path
    savePath = "C:\\Users\\hriday.ghosh\\OneDrive - Concentrix Corporation\\Hriday\\Tristar Italy\\Schedule Uploads\\" & folderName & "\\"

    ' Save as CSV
    Application.DisplayAlerts = False
    wbNew.SaveAs fileName:=savePath & fileName, FileFormat:=xlCSV
    Application.DisplayAlerts = True

    MsgBox "New CSV file created successfully."

End Sub


Sub DeleteRegion()

    ' ===== Clears data from Updated Schedule =====

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim WS As Worksheet
    Set WS = Sheets("Updated Schedule")
    WS.Activate

    ' Select data starting from A2 till last row & column
    WS.Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select

    ' Clear only values, not formatting
    Selection.ClearContents

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "Process is Done."

End Sub


