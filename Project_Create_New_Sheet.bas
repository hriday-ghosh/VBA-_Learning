Attribute VB_Name = "Module1"
Sub Create_Sheet_Beginner()
' Sub starts a macro (subroutine)
' Create_Sheet_Beginner ? name of the macro
' () macro takes no inputs

    Dim wbSource As Workbook
    ' Dim ? used to declare (create) a variable
    ' wbSource ? name of the variable
    ' As Workbook ? this variable will store a workbook object

    Dim wbNew As Workbook
    ' Will hold the new workbook created after copying the sheet

    Dim wsLast As Worksheet
    ' Will store the last sheet from the source workbook

    Dim wsNew As Worksheet
    ' Will store the copied sheet inside the new workbook

    Dim rng As Range
    ' Will store a block of cells (A1 to I15000)

    Dim dt As Date
    ' Will store a date value (from cell B2)

    Dim fileName As String
    ' String text variable

    Dim savePath As String
    ' Will hold folder path text

    Set wbSource = ThisWorkbook
    ' Set ? used when assigning objects (Workbook, Sheet, Range)
    ' ThisWorkbook ? the workbook that contains THIS macro
    ' = assigns the workbook to wbSource

    Set wsLast = wbSource.Sheets(wbSource.Sheets.Count)
    ' wbSource.Sheets ? all sheets inside the workbook
    ' .Count ? gives number of sheets
    ' Sheets(number) ? picks the sheet at that position
    ' Therefore Sheets(Count) ? last sheet

    wsLast.Copy
    ' Copy ? copies the entire sheet
    ' When a sheet is copied without specifying destination, Excel creates a NEW workbook automatically

    Set wbNew = ActiveWorkbook
    ' ActiveWorkbook ? the workbook currently active (in front)
    ' After copying a sheet, Excel activates the new workbook
    ' So we store that workbook into wbNew

    Set wsNew = wbNew.Sheets(1)
    ' Sheets(1) ? first sheet inside the new workbook
    ' That is the copied sheet

    wsNew.Activate
    ' Activate ? brings the sheet into view (makes it active)

    wsNew.Range("A1").Select
    ' Range("A1") ? a single cell
    ' "A1" is a STRING
    ' Select ? highlight that cell

    Set rng = wsNew.Range("A1:I15000")
    ' Range("A1:I15000") ? large block of cells
    ' Set ? assign this Range object to rng

    rng.Copy
    ' Copy ? put cells into clipboard (copy operation)

    rng.PasteSpecial Paste:=xlPasteValues
    ' PasteSpecial ? advanced paste options
    ' Paste:= assigns a parameter
    ' xlPasteValues ? paste only values (no formulas, no formatting)

    Application.CutCopyMode = False
    ' Application ? refers to Excel itself
    ' CutCopyMode ? indicates whether Excel is in copy/paste mode
    ' False ? turn off copy mode (remove moving dashed border)

    dt = wsNew.Range("B2").Value
    ' .Value ? read the content of the cell
    ' dt now stores the date from B2

    fileName = "FC_NICE02B_102849782_SamsungPoland_" & Format(dt, "ddmmyy") & ".csv"
    ' "text" ? fixed text string
    ' & ? join text pieces together (concatenation)
    ' Format(value, "ddmmyy") ? convert date to ddmmyy form
    ' ".csv" ? file extension

    savePath = "C:\Users\hriday.ghosh\OneDrive - Concentrix Corporation\Hriday\Samsung Poland\Upload Forecast Here\"
    ' A fixed folder location stored as text

    Application.DisplayAlerts = False
    ' Turn off Excel warning pop-ups (Example: overwrite warning)

    wbNew.SaveAs fileName:=savePath & fileName, FileFormat:=xlCSV
    ' SaveAs ? save file with new name or location

End Sub



Sub Create_Sheet_Professional()

    Dim wbSource As Workbook
    ' This workbook contains the macro.

    Dim wbNew As Workbook
    ' This will store the newly created workbook.

    Dim wsLast As Worksheet
    ' This stores the last sheet from the source workbook.

    Dim wsNew As Worksheet
    ' This stores the copied sheet inside the new workbook.

    Dim dt As Date
    ' Stores the date picked from B2.

    Dim fileName As String
    ' Stores the final file name.

    Dim savePath As String
    ' Stores folder location where file is saved.


    Set wbSource = ThisWorkbook
    ' Source workbook is the one running this macro.

    Set wsLast = wbSource.Sheets(wbSource.Sheets.Count)
    ' Identify last sheet in workbook.

    wsLast.Copy
    ' Copy last sheet into a new workbook.

    Set wbNew = ActiveWorkbook
    ' The newly created workbook is now ActiveWorkbook.

    Set wsNew = wbNew.Sheets(1)
    ' The copied sheet becomes sheet 1.


    wsNew.Range("A1:I15000").Value = wsNew.Range("A1:I15000").Value
    ' Convert formulas to values without using Copy/Paste.


    dt = wsNew.Range("B2").Value
    ' Read date from B2.

    fileName = "FC_NICE02B_102849782_SamsungPoland_" & Format(dt, "ddmmyy") & ".csv"
    ' Build the file name.

    savePath = "C:\Users\hriday.ghosh\OneDrive - Concentrix Corporation\Hriday\Samsung Poland\Upload Forecast Here\"
    ' Folder path for saving the file.

    Application.DisplayAlerts = False
    ' Disable warnings.

    wbNew.SaveAs savePath & fileName, xlCSV
    ' Save workbook as CSV.

    Application.DisplayAlerts = True
    ' Enable warnings back.

    MsgBox "CSV file created successfully!", vbInformation
    ' Confirmation message.

End Sub



Sub Create_Sheet_Debug()

    MsgBox "Step 1: Declaring variables"

    Dim wbSource As Workbook
    ' Source workbook.

    Dim wbNew As Workbook
    ' New workbook.

    Dim wsLast As Worksheet
    ' Last sheet.

    Dim wsNew As Worksheet
    ' Copied sheet.

    Dim dt As Date
    ' Date value.

    Dim fileName As String
    ' Final file name.

    Dim savePath As String
    ' Folder path.


    MsgBox "Step 2: Setting source workbook"
    Set wbSource = ThisWorkbook


    MsgBox "Step 3: Finding last sheet"
    Set wsLast = wbSource.Sheets(wbSource.Sheets.Count)


    MsgBox "Step 4: Copying sheet"
    wsLast.Copy


    MsgBox "Step 5: Identifying new workbook"
    Set wbNew = ActiveWorkbook
    Set wsNew = wbNew.Sheets(1)


    MsgBox "Step 6: Converting formulas to values"
    wsNew.Range("A1:I15000").Value = wsNew.Range("A1:I15000").Value


    MsgBox "Step 7: Reading date"
    dt = wsNew.Range("B2").Value


    MsgBox "Step 8: Creating file name"
    fileName = "FC_NICE02B_102849782_SamsungPoland_" & Format(dt, "ddmmyy") & ".csv"


    MsgBox "Step 9: Setting save path"
    savePath = "C:\Users\hriday.ghosh\OneDrive - Concentrix Corporation\Hriday\Samsung Poland\Upload Forecast Here\"


    MsgBox "Step 10: Saving file"
    Application.DisplayAlerts = False
    wbNew.SaveAs savePath & fileName, xlCSV
    Application.DisplayAlerts = True


    MsgBox "Step 11: Done! File Saved Successfully."

End Sub



Sub Create_Sheet_Advanced()

    On Error GoTo ErrorHandler
    ' Enable error handling. If any error occurs, code jumps to ErrorHandler.


    Dim wbSource As Workbook
    ' Stores the workbook where this macro is written.

    Dim wbNew As Workbook
    ' Stores the new workbook created from sheet copy.

    Dim wsLast As Worksheet
    ' Stores the last worksheet of the source workbook.

    Dim wsNew As Worksheet
    ' Stores the copied worksheet.

    Dim dt As Date
    ' Stores the date extracted from B2.

    Dim fileName As String
    ' Stores the CSV file name.

    Dim savePath As String
    ' Stores the folder path for saving the file.



    Set wbSource = ThisWorkbook
    ' Assign the current workbook as the source workbook.



    Set wsLast = wbSource.Sheets(wbSource.Sheets.Count)
    ' Identify the last sheet based on sheet count.



    wsLast.Copy
    ' Copy the last sheet to a new workbook.



    Set wbNew = ActiveWorkbook
    ' The copied sheet creates a new active workbook. Store it.



    Set wsNew = wbNew.Sheets(1)
    ' The copied sheet is now the first sheet in the new workbook.



    wsNew.Range("A1:I15000").Value = wsNew.Range("A1:I15000").Value
    ' Convert entire range to values (remove formulas).



    dt = wsNew.Range("B2").Value
    ' Read the date value from cell B2.



    fileName = "FC_NICE02B_102849782_SamsungPoland_" & Format(dt, "ddmmyy") & ".csv"
    ' Build the file name using formatted date.



    savePath = "C:\Users\hriday.ghosh\OneDrive - Concentrix Corporation\Hriday\Samsung Poland\Upload Forecast Here\"
    ' Predefined folder location for saving the CSV file.



    Application.DisplayAlerts = False
    ' Turn off Excel warnings (CSV overwrite warning, format warning).



    wbNew.SaveAs savePath & fileName, xlCSV
    ' Save the file as CSV in the defined folder.



    Application.DisplayAlerts = True
    ' Re-enable Excel warnings.



    MsgBox "File saved successfully:" & vbCrLf & savePath & fileName, vbInformation
    ' Display success message.



CleanUp:
    ' This section ensures proper cleanup before exiting.



    On Error Resume Next
    ' Prevent errors during cleanup phase.



    If Not wbNew Is Nothing Then wbNew.Close SaveChanges:=False
    ' Close the new workbook if it is still open (extra safety).



    Exit Sub
    ' Exit macro to prevent running the ErrorHandler block on success.



ErrorHandler:
    ' Code jumps here if an error occurs anywhere above.



    Application.DisplayAlerts = True
    ' Ensure alerts are turned back on even after error.



    MsgBox "An error occurred:" & vbCrLf & Err.Description, vbCritical
    ' Show the error message to the user.



    Resume CleanUp
    ' Go to cleanup section to safely exit.

End Sub

Sub Create_Sheet_Beginner_AutoPath()

    ' If any error happens, go to the error section.
    On Error GoTo ErrorSection


    Dim mainWorkbook As Workbook
    ' This is the workbook where the macro is stored.

    Dim newWorkbook As Workbook
    ' This will be the new workbook created after copying a sheet.

    Dim lastSheet As Worksheet
    ' This will be the last sheet in the current workbook.

    Dim copiedSheet As Worksheet
    ' This will be the sheet that gets copied into the new workbook.

    Dim theDate As Date
    ' This will store the date from cell B2.

    Dim fileName As String
    ' This will store the CSV file name.

    Dim saveFolder As String
    ' This will store the folder where the CSV will be saved.



    Set mainWorkbook = ThisWorkbook
    ' Set the main workbook (the current file).



    Set lastSheet = mainWorkbook.Sheets(mainWorkbook.Sheets.Count)
    ' Find the last sheet in the workbook.



    lastSheet.Copy
    ' Copy the last sheet. Excel automatically creates a new workbook.



    Set newWorkbook = ActiveWorkbook
    ' The workbook created by copy becomes "active".



    Set copiedSheet = newWorkbook.Sheets(1)
    ' The copied sheet becomes sheet number 1 in the new workbook.



    copiedSheet.Range("A1:I15000").Value = copiedSheet.Range("A1:I15000").Value
    ' Convert formulas to values. This removes all formulas.



    theDate = copiedSheet.Range("B2").Value
    ' Pick the date from cell B2.



    fileName = "FC_NICE02B_102849782_SamsungPoland_" & Format(theDate, "ddmmyy") & ".csv"
    ' Make the file name using the date.



    saveFolder = mainWorkbook.Path & "\"
    ' Automatically get the folder where the Excel file is saved.



    Application.DisplayAlerts = False
    ' Turn off warning messages.



    newWorkbook.SaveAs saveFolder & fileName, xlCSV
    ' Save the new workbook as a CSV in the same folder.



    Application.DisplayAlerts = True
    ' Turn warnings back on.



    MsgBox "CSV File Saved Successfully!" & vbCrLf & saveFolder & fileName
    ' Show message telling where the file is saved.



CleanUp:
    ' This section handles closing and cleanup before ending.



    On Error Resume Next
    ' Ignore errors during cleanup.



    newWorkbook.Close SaveChanges:=False
    ' Close the temporary new workbook.



    Exit Sub
    ' Exit the macro so the error section does not run.



ErrorSection:
    ' This runs if any error happens in the macro.



    MsgBox "Something went wrong: " & Err.Description, vbCritical
    ' Show the error message.



    Resume CleanUp
    ' Go to cleanup section to safely exit.

End Sub

Sub OpenRawFile_AndLog()

    '1. Get the folder where THIS Excel file is saved
    Dim mainFolder As String
    mainFolder = ThisWorkbook.Path

    '2. Build the path of the file to open
    '?? Change the file name as per your need
    Dim targetFile As String
    targetFile = mainFolder & "\Raw\DataFile.xlsx"
    
    '3. Open the file
    Workbooks.Open targetFile

    '4. Create path for the log file
    Dim logFilePath As String
    logFilePath = mainFolder & "\Log.txt"

    '5. Prepare the log message
    Dim FSO As Object
    Dim logFile As Object
    Dim logMessage As String

    'Who ran the code (Excel ID) + Date + Time
    logMessage = Application.UserName & " ran the macro on " & Date & " at " & Time & vbCrLf

    '6. Create or open log file and write the message
    Set FSO = CreateObject("Scripting.FileSystemObject")

    'If file exists ? open & append
    If FSO.FileExists(logFilePath) Then
        Set logFile = FSO.OpenTextFile(logFilePath, 8) '8 = append mode
    Else
        Set logFile = FSO.CreateTextFile(logFilePath) 'create new file
    End If

    logFile.WriteLine logMessage
    logFile.Close

End Sub

Sub OpenRawFile_AndLog()
    'This starts the macro

    Dim mainFolder As String
    'This creates a variable named mainFolder to store a folder path as text

    mainFolder = ThisWorkbook.Path
    'This gets the folder where the current Excel file is saved
    'Example: "C:\Users\Hriday\Documents"

    Dim targetFile As String
    'This creates a variable named targetFile to store the full file path of the file we want to open

    targetFile = mainFolder & "\Raw\DataFile.xlsx"
    'This joins the folder path with the file location and name
    'Example result: "C:\Users\Hriday\Documents\Raw\DataFile.xlsx"

    Workbooks.Open targetFile
    'This opens the file written in targetFile

    Dim logFilePath As String
    'This creates a variable to store the log file’s location

    logFilePath = mainFolder & "\Log.txt"
    'This prepares the location where the log text file will be saved
    'Example: "C:\Users\Hriday\Documents\Log.txt"

    Dim FSO As Object
    'This creates a variable to use the File System Object (used to work with files/folders)

    Dim logFile As Object
    'This creates a variable to represent the log file itself

    Dim logMessage As String
    'This creates a variable to store the text we will write in Log.txt

    logMessage = Application.UserName & " ran the macro on " & Date & " at " & Time & vbCrLf
    'This prepares the message
    'Application.UserName = the Excel username
    'Date = today’s date
    'Time = current time
    'vbCrLf = moves to the next line after writing

    Set FSO = CreateObject("Scripting.FileSystemObject")
    'This creates a File System Object which lets VBA create/read/write files

    If FSO.FileExists(logFilePath) Then
        'This checks if the Log.txt file already exists

        Set logFile = FSO.OpenTextFile(logFilePath, 8)
        'If the log file exists, open it in "append" mode (8 means add new lines without deleting old ones)

    Else
        Set logFile = FSO.CreateTextFile(logFilePath)
        'If the log file does NOT exist, create a new Log.txt file
    End If

    logFile.WriteLine logMessage
    'This writes the log message into the Log.txt file

    logFile.Close
    'This closes the log file to save the content

End Sub
'This ends the macro

