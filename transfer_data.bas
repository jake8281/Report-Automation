'Created By Jake Ayoub 6/3/2021
'Updated 6/4/2021

Sub migrate_data()

    Dim new_sheet As String
       
    Dim emp_acc_col As String
    Dim bos_acc_col As String
    Dim bos_address1_col As String
    Dim emp_address1_col As String
    Dim emp_address2_col As String
    Dim city_col As String
    Dim state_col As String
    Dim zip_col As String
    
    Dim emp_acc_rng As String
    Dim bos_acc_rng As String
    Dim bos_address1_rng As String
    Dim emp_address1_rng As String
    Dim emp_address2_rng As String
    Dim city_rng As String
    Dim state_rng As String
    Dim zip_rng As String
    
    With combine_report
        ' This function let you run the code from any worksheet in the workbook
        Application.ScreenUpdating = True
        Worksheets("combine_report").Activate
        Application.ScreenUpdating = False
        
        ' This piece used to provide the range and it returns column name.
        emp_acc_col = getColStr("Empower Account Number")
        bos_acc_col = getColStr("BOS Account number")
        bos_address1_col = getColStr("BOS Address 1")
        emp_address1_col = getColStr("Empower Address 1")
        emp_address2_col = getColStr("Empower Address 2")
        city_col = getColStr("Empower City")
        state_col = getColStr("Empower State")
        zip_col = getColStr("Empower Zip")
            
        ' This used to concatnate whole column
        emp_acc_rng = "" & emp_acc_col & ":" & emp_acc_col
        bos_acc_rng = "" & bos_acc_col & ":" & bos_acc_col
        bos_address1_rng = "" & bos_address1_col & ":" & bos_address1_col
        emp_address1_rng = "" & emp_address1_col & ":" & emp_address1_col
        emp_address2_rng = "" & emp_address2_col & ":" & emp_address2_col
        city_rng = "" & city_col & ":" & city_col
        state_rng = "" & state_col & ":" & state_col
        zip_rng = "" & zip_col & ":" & zip_col
        
        new_sheet = "empower_report"
        'This is used add new sheet
        With ThisWorkbook
            .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = new_sheet
        End With
    
        ThisWorkbook.Sheets("combine_report").Activate
        ' This is used to copy the range from one sheet to another and remove the formulas.
        With combine_report
            Range(emp_acc_rng).Copy
            Worksheets(new_sheet).Range(emp_acc_rng).PasteSpecial xlPasteValues
    
            Range(bos_acc_rng).Copy
            Worksheets(new_sheet).Range(bos_acc_rng).PasteSpecial xlPasteValues
    
            Range(bos_address1_rng).Copy
            Worksheets(new_sheet).Range(bos_address1_rng).PasteSpecial xlPasteValues
    
            Range(emp_address1_rng).Copy
            Worksheets(new_sheet).Range(emp_address1_rng).PasteSpecial xlPasteValues
            
            Range(emp_address2_rng).Copy
            Worksheets(new_sheet).Range(emp_address2_rng).PasteSpecial xlPasteValues
    
            Range(city_rng).Copy
            Worksheets(new_sheet).Range(city_rng).PasteSpecial xlPasteValues
    
            Range(state_rng).Copy
            Worksheets(new_sheet).Range(state_rng).PasteSpecial xlPasteValues
    
            Range(zip_rng).Copy
            Worksheets(new_sheet).Range(zip_rng).PasteSpecial xlPasteValues
        End With
    End With
End Sub
