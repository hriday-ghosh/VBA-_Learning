Attribute VB_Name = "Module1"
Sub MBPerformanceUSV4Email()
' This is the start of the macro.
' Sub = Subroutine (a block of code that performs a task)

    Application.ScreenUpdating = False
    ' This stops Excel from refreshing the screen again and again.
    ' Makes the macro run faster and avoids screen flickering.

    Dim outlook As Object
    ' Variable to control Outlook application

    Dim weditor As Object
    ' Variable to control the email body editor (Word editor inside Outlook)

    Dim outlookMail As Object
    ' Variable to control a new Outlook email

    Dim Emailbody As Range
    ' Variable to store an Excel range (email body part 1)

    ' These ranges are named ranges in Excel
    ' They contain formatted email content
    Set Emailbody = Range("USPMEMAIL")
    Set Emailbody1 = Range("USPMEMAIL1")
    Set Emailbody2 = Range("USPMEMAIL3")

    ' Open Outlook application using VBA
    Set outlook = CreateObject("Outlook.Application")

    ' Create a new email item (0 means Mail Item)
    Set outlookMail = outlook.CreateItem(0)

    ' Start filling email details
    With outlookMail

        .SentOnBehalfOfName = Range("USPMFROM").Value
        ' Sends email on behalf of the email ID mentioned in Excel

        .Display
        ' Opens the email window (IMPORTANT before pasting body)

        .To = Range("USPMTO").Value
        ' Sets TO recipients from Excel cell

        .CC = Range("USPMCC").Value
        ' Sets CC recipients from Excel cell

        .Subject = Range("USPMUBJECT").Value
        ' Sets email subject from Excel cell

        .Attachments.Add Range("USPMATTACHMENT").Value
        ' Attaches first file (path stored in Excel)

        .Attachments.Add Range("RONAsMTD").Value
        ' Attaches second file

        .Attachments.Add Range("RONAsDAILY").Value
        ' Attaches third file

    End With
    ' End of email setup section

    On Error GoTo 0
    ' Resets error handling to default
    ' (Earlier errors, if any, are ignored)

    ' Get access to the email body editor (Word editor inside Outlook)
    Set weditor = outlook.ActiveInspector.WordEditor

    ' Copy first email body content from Excel
    Emailbody.Copy

    ' Paste it into Outlook email body
    weditor.Application.Selection.Paste

    ' Add blank lines (line breaks) for spacing
    Text = vbCrLf & vbCrLf & vbCrLf & vbCrLf

    ' Copy second part of email body
    Emailbody1.Copy

    ' Paste second content
    weditor.Application.Selection.Paste

    ' Add spacing again
    Text = vbCrLf & vbCrLf & vbCrLf & vbCrLf

    ' Copy third part of email body
    Emailbody2.Copy

    ' Paste third content
    weditor.Application.Selection.Paste

    ' Clear memory – good practice
    Set outlookMail = Nothing
    Set outlook = Nothing

End Sub
' End of macro

Sub Send_Performance_Email_Beginner()

    '--------------------------------------------------
    ' STEP 1: Speed up Excel by stopping screen refresh
    '--------------------------------------------------
    Application.ScreenUpdating = False

    '--------------------------------------------------
    ' STEP 2: Declare variables (easy-to-understand names)
    '--------------------------------------------------
    Dim OutlookApp As Object
    ' Outlook application
    Dim OutlookEmail As Object
    ' New email
    Dim EmailEditor As Object
    ' Email body editor (Word)

    Dim BodyPart1 As Range
    ' First email body section
    Dim BodyPart2 As Range
    ' Second email body section
    Dim BodyPart3 As Range
    ' Third email body section

    '--------------------------------------------------
    ' STEP 3: Assign Excel named ranges to variables
    '--------------------------------------------------
    Set BodyPart1 = Range("USPMEMAIL")
    Set BodyPart2 = Range("USPMEMAIL1")
    Set BodyPart3 = Range("USPMEMAIL3")

    '--------------------------------------------------
    ' STEP 4: Open Outlook and create a new email
    '--------------------------------------------------
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookEmail = OutlookApp.CreateItem(0)
    ' revealed as Mail Item

    '--------------------------------------------------
    ' STEP 5: Fill email details from Excel
    '--------------------------------------------------
    With OutlookEmail

        .Display
        ' Display is required before pasting formatted content

        .SentOnBehalfOfName = Range("USPMFROM").Value
        ' From email ID

        .To = Range("USPMTO").Value
        ' To recipients

        .CC = Range("USPMCC").Value
        ' CC recipients

        .Subject = Range("USPMUBJECT").Value
        ' Email subject

        .Attachments.Add Range("USPMATTACHMENT").Value
        ' First attachment

        .Attachments.Add Range("RONAsMTD").Value
        ' Second attachment

        .Attachments.Add Range("RONAsDAILY").Value
        ' Third attachment

    End With

    '--------------------------------------------------
    ' STEP 6: Get access to email body editor (Word)
    '--------------------------------------------------
    Set EmailEditor = OutlookApp.ActiveInspector.WordEditor

    '--------------------------------------------------
    ' STEP 7: Paste email body content from Excel
    '--------------------------------------------------
    BodyPart1.Copy
    EmailEditor.Application.Selection.Paste

    ' Add space between sections
    EmailEditor.Application.Selection.TypeParagraph
    EmailEditor.Application.Selection.TypeParagraph

    BodyPart2.Copy
    EmailEditor.Application.Selection.Paste

    EmailEditor.Application.Selection.TypeParagraph
    EmailEditor.Application.Selection.TypeParagraph

    BodyPart3.Copy
    EmailEditor.Application.Selection.Paste

    '--------------------------------------------------
    ' STEP 8: Clean up and turn screen updating back on
    '--------------------------------------------------
    Set OutlookEmail = Nothing
    Set OutlookApp = Nothing

    Application.ScreenUpdating = True

End Sub

