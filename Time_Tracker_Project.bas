Sub TaskTrackerUpdate()
' This is the name of our macro (program inside Excel)

    Dim ws As Worksheet              ' ws = a sheet in Excel
    Dim i As Long                    ' i = number we will use for looping rows

    Set ws = ThisWorkbook.Sheets("Sheet1") 
    ' Tell VBA to work with "Sheet1" from this Excel file
    ' (change name if your sheet is different)

    ' We now start checking rows, from row 2 to row 10000
    For i = 2 To 10000              
        ' Look at column B in the current row
        ' Convert the word to CAPITAL so we don’t worry about small/big letters
        Select Case UCase(ws.Cells(i, "B").Value)

            Case "STARTED"          ' If column B has the word "STARTED" then:

                ' If Start Time (column C) is empty → write the current time
                If ws.Cells(i, "C").Value = "" Then
                    ws.Cells(i, "C").Value = Time
                End If

                ' If Start Date (column D) is empty → write today’s date
                If ws.Cells(i, "D").Value = "" Then
                    ws.Cells(i, "D").Value = Date
                End If

                ' Write "Still Working" in Progress column (G)
                ws.Cells(i, "G").Value = "Still Working"


            Case "COMPLETED"        ' If column B has the word "COMPLETED" then:

                ' If End Date (E) is empty → put today’s date
                If ws.Cells(i, "E").Value = "" Then
                    ws.Cells(i, "E").Value = Date
                End If

                ' If End Time (F) is empty → put current time
                If ws.Cells(i, "F").Value = "" Then
                    ws.Cells(i, "F").Value = Time
                End If

                ' Mark Progress (G) as "Task Completed"
                ws.Cells(i, "G").Value = "Task Completed"

                ' Now calculate how long the task took (Duration column H)
                If ws.Cells(i, "C").Value <> "" And ws.Cells(i, "F").Value <> "" Then
                    ws.Cells(i, "H").Value = _
                        (ws.Cells(i, "E").Value + ws.Cells(i, "F").Value) - _
                        (ws.Cells(i, "D").Value + ws.Cells(i, "C").Value)
                    ' Above: (End Date + End Time) - (Start Date + Start Time)

                    ' Format H to show hours:minutes:seconds
                    ws.Cells(i, "H").NumberFormat = "[h]:mm:ss"
                End If


            Case Else
                ' If column B is blank or has some other word → do nothing
        End Select
    Next i
    ' Loop ends when we reach row 10000

    ' After finishing, show a small message box
    MsgBox "Task Tracker Updated Successfully!", vbInformation

End Sub
' End of macro