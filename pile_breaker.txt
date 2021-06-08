'Created By Jake Ayoub 6/3/2021
'Updated 6/5/2021
'This function to organize complied data

Sub data_organizer()

    Dim col_rng As String
    
    With piled_data
        ' This function let you run the code from any worksheet in the workbook
        Application.ScreenUpdating = True
        Worksheets("piled_data").Activate
        Application.ScreenUpdating = False
               
        'insert column that splited values will be pasted into
        Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("A1").Value = "Dates"
        Range("B1").Value = "Acc_no"
        Range("C1").Value = "Client_Name"
        Range("D1").Value = "Street_Address"
        Range("E1").Value = "Apt/Unit_no"
        Range("F1").Value = "City"
        Range("G1").Value = "State"
        Range("H1").Value = "Status"
        Range("I1").Value = "Zip_code"
        
        
        col_rng = getColRange("Dates")
        ' This function will split the cell after "-"
        With Range(col_rng)
            .TextToColumns Destination:=Range("A2"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
            :="~", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
            1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
            , 1), Array(13, 1), Array(14, 1)), TrailingMinusNumbers:=True
        End With
        
    Cells.EntireColumn.AutoFit
    End With
End Sub



