Sub update_report()

Dim empower_format_rng As String
Dim empower_msgbox As String

With empower_report

    Application.ScreenUpdating = True
    Worksheets("empower_report").Activate
    Application.ScreenUpdating = False
    
    ' This clear contents in a given range Excluding first row.
    Range(Cells(2, 3), Cells(Rows.Count, 3)).ClearContents
    
    ' This Big Nest to build Conditional Formatting using if Cell contains Specific Text
    empower_format_rng = getColRange("Empower Address 1")
    With Range(empower_format_rng)
        .FormatConditions.Add Type:=xlTextString, String:="Po Bo", _
                TextOperator:=xlContains
            .FormatConditions(.FormatConditions.Count).SetFirstPriority
           With .FormatConditions(1).Font
               .Color = -16727809
               .TintAndShade = 0
           End With
           With .FormatConditions(1).Interior
               .PatternColorIndex = xlAutomatic
               .Color = 255
               .TintAndShade = 0
           End With
           .FormatConditions(1).StopIfTrue = False
    End With
    
    ' This is a sneaky Part
    empower_msgbox = getColRange("Empower Address 1")
    'First Part: Delete string after certain Character
    For Each c In Range(empower_msgbox)
    If InStr(c.Value, "x") Then
        c.Value = Left(c.Value, InStr(c.Value, "x") - 1)
    End If
    Next
    ' Pop up a message box if the range contains specfic string
    If Not IsError(Application.Match("Po Bo", Range(empower_msgbox), 0)) Then
         MsgBox ("Warning!!!! ---PO Box--- detected in Empower Address 2 Please input plan ID address")
    End If
End With
End Sub
