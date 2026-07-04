Attribute VB_Name = "Module2"
Sub Example()

    'Step 1:
    'Macro start hocche
    MsgBox "Start"

    'Step 2:
    'Ekhon niche line execute korbe na
    'Soja "EndCode:" label e jump korbe
    GoTo EndCode

    'Step 3:
    'Ei line ta kokhono execute hobe na
    'Karon GoTo agei EndCode e niye geche
    MsgBox "Eta ar cholbe na"

'---------------------------------------------------
'Eta kono function noy
'Eta sudhu ekta Label (Address)
'GoTo ei jayga tei jump kore
'---------------------------------------------------
EndCode:

    'Step 4:
    'GoTo jump korar por ei line execute hobe
    MsgBox "End"

End Sub
