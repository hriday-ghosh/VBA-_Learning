Sub TaskTrackerUpdate()
' This is the name of our macro (program inside Excel)

    Dim ws As Worksheet              
    ' ws = a sheet in Excel
    Dim i As Long                    
    ' i = number we will use for looping rows

    Set ws = ThisWorkbook.Sheets("Sheet1") 
    ' Tell VBA to work with "Sheet1" from this Excel file
    ' (change name if your sheet is different)

    For i = 2 To 10000
    ' We now start checking rows, from row 2 to row 10000

        Select Case UCase(ws.Cells(i, "B").Value)
        ' Look at column B in the current row
        ' Convert the word to CAPITAL so we don’t worry about small/big letters
        
            Case "STARTED"          
            ' If column B has the word "STARTED" then:

                If ws.Cells(i, "C").Value = "" Then
                ' If Start Time (column C) is empty → write the current time
                    ws.Cells(i, "C").Value = Time
                End If

                If ws.Cells(i, "D").Value = "" Then
                ' If Start Date (column D) is empty → write today’s date
                    ws.Cells(i, "D").Value = Date
                End If

                ws.Cells(i, "G").Value = "Still Working"
                ' Write "Still Working" in Progress column (G)

            Case "COMPLETED"        
            ' If column B has the word "COMPLETED" then:

                If ws.Cells(i, "E").Value = "" Then
                  ' If End Date (E) is empty Then put today’s date  
                    ws.Cells(i, "E").Value = Date
                End If

                
                If ws.Cells(i, "F").Value = "" Then
                ' If End Time (F) is empty Then put current time
                    ws.Cells(i, "F").Value = Time
                End If

                
                ws.Cells(i, "G").Value = "Task Completed"
                ' Mark Progress (G) as "Task Completed"

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

    
    MsgBox "Task Tracker Updated Successfully!", vbInformation
    ' After finishing, show a small message box

End Sub
' End of macro
