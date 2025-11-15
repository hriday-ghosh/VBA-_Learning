Attribute VB_Name = "Module3"
Sub GoTo_And_NoGoTo_Together()

    Dim score As Integer
    score = 45                      '
    'Example Number

    ' ------------------------------------------------------
    ' PART 1 - Using IF normally (No GoTo)
    ' ------------------------------------------------------
    
    If score >= 50 Then
    'Normal IF check
        MsgBox "PASS (Normal IF)"
        'Runs if score >= 50
    Else
        MsgBox "FAIL (Normal IF)"
        'Runs if score < 50
    End If


    ' ------------------------------------------------------
    ' PART 2 - Doing SAME CHECK but using GoTo
    ' ------------------------------------------------------

    If score >= 50 Then GoTo PassLabel
    ' Jump to PASS
    GoTo FailLabel
    ' If not =50, go to FAIL

PassLabel:
' Label for PASS
    MsgBox "PASS"
    ' Message for PASS
    GoTo EndCheck
    ' Jump to end (avoid FAIL part)


FailLabel:
' Label for FAIL
    MsgBox "FAIL"
    ' Message for FAIL


EndCheck:
' Final label for ending
    MsgBox "Check Complete"
    ' Common ending message

End Sub

Sub Grade_Check_GoTo()

    Dim marks As Integer
    marks = Range("A1").Value


    '==========================================
    ' LABELS (Placed on top but protected)
    '==========================================

GradeA:
    MsgBox "Grade A"
    GoTo EndCheck


GradeB:
    MsgBox "Grade B"
    GoTo EndCheck


FailLabel:
    MsgBox "Fail"
    GoTo EndCheck


EndCheck:
    ' Nothing here yet



    '==========================================
    ' START THE CHECKING HERE (important!)
    '==========================================

StartCheck:

    If marks >= 90 And marks <= 100 Then GoTo GradeA

    If marks >= 80 And marks < 90 Then GoTo GradeB

    GoTo FailLabel


End Sub

