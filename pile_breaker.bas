'Created By Jake Ayoub 6/3/2021
'Updated 6/11/2021
'This function to break complied data and migrate data to the designated worksheet

Sub data_organizer()
    
    Dim col_rng As String
    Dim boring_filter As String
    Dim existing_Sheet As String
     
    Dim date_col As String
    Dim date_rng As String
    
    Dim acc_col As String
    Dim acc_rng As String
    
    Dim ad1_col As String
    Dim ad1_rng As String
    
    Dim ad2_col As String
    Dim ad2_rng As String
    
    Dim city_col As String
    Dim city_rng As String
    
    Dim state_col As String
    Dim state_rng As String
    
    Dim zip_col As String
    Dim zip_rng As String
    
    With piled_data
        ' This function let you run the code from any worksheet in the workbook
        Application.ScreenUpdating = True
        Worksheets("piled_data").Activate
        Application.ScreenUpdating = False
               
        'insert column that splited values will be pasted into
        Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("A1").Value = "Date"
        Range("B1").Value = "Account Number"
        Range("C1").Value = "Name"
        Range("D1").Value = "Address 1"
        Range("E1").Value = "Address 2"
        Range("F1").Value = "City"
        Range("G1").Value = "State"
        Range("H1").Value = ""
        Range("I1").Value = "Zip"
        
        
        col_rng = getColRange("Date")
        ' This function will split the cell after "-"
        With Range(col_rng)
            .TextToColumns Destination:=Range("A2"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
            :="~", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
            1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
            , 1), Array(13, 1), Array(14, 1)), TrailingMinusNumbers:=True
        End With
        
        ' Add Filter to the worksheet and Expand all rows
        boring_filter = Range("A1:I1").AutoFilter
        Cells.EntireColumn.AutoFit
        
        ' this is used to transfer ranges from one sheet to another - reference from polish_data code
        date_col = getColStr("Date")
        date_rng = "" & date_col & ":" & date_col
        exisiting_Sheet = "onbase_data"

        Range(date_rng).Copy
        Worksheets(exisiting_Sheet).Range(date_rng).PasteSpecial xlPasteValues
        
        acc_col = getColStr("Account Number")
        acc_rng = "" & acc_col & ":" & acc_col

        Range(acc_rng).Copy
        Worksheets(exisiting_Sheet).Range(acc_rng).PasteSpecial xlPasteValues
        
        name_col = getColStr("Name")
        name_rng = "" & name_col & ":" & name_col

        Range(name_rng).Copy
        Worksheets(exisiting_Sheet).Range(name_rng).PasteSpecial xlPasteValues
        
        ad1_col = getColStr("Address 1")
        ad1_rng = "" & ad1_col & ":" & ad1_col

        Range(ad1_rng).Copy
        Worksheets(exisiting_Sheet).Range(ad1_rng).PasteSpecial xlPasteValues
        
        ad2_col = getColStr("Address 2")
        ad2_rng = "" & ad2_col & ":" & ad2_col

        Range(ad2_rng).Copy
        Worksheets(exisiting_Sheet).Range(ad2_rng).PasteSpecial xlPasteValues
        
        city_col = getColStr("City")
        city_rng = "" & city_col & ":" & city_col

        Range(city_rng).Copy
        Worksheets(exisiting_Sheet).Range(city_rng).PasteSpecial xlPasteValues
        
        state_col = getColStr("State")
        state_rng = "" & state_col & ":" & state_col

        Range(state_rng).Copy
        Worksheets(exisiting_Sheet).Range(state_rng).PasteSpecial xlPasteValues
        
        
        zip_col = getColStr("Zip")
        zip_rng = "" & zip_col & ":" & zip_col

        Range(zip_rng).Copy
        Worksheets(exisiting_Sheet).Range(zip_rng).PasteSpecial xlPasteValues
    End With
End Sub



