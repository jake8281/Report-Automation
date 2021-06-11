'Created By Jake Ayoub 6/3/2021
'Updated 6/10/2021

Sub build_report()

    Dim myRow As Long
    Dim myCounter As Long

    Dim empower_rng As String
    Dim counter As Integer

    Dim address1_rng As String
    Dim address2_rng As String

    Dim bos_txt_fix As Range
    Dim bos_cell As Range
    Dim bos_selection As String

    Dim emp_txt_fix As Range
    Dim emp_cell As Range
    Dim emp_selection As String

    Dim emp_add As String
    Dim trim_row As Integer
    Dim trim_rng As Range
    Dim trimcell As Range

    Dim empcol_rng As String
    Dim boscol_rng As String
    Dim first_col  As Range
    Dim second_col  As Range
    Dim n As Long


    'code will run n times because it fixes a sneaky bug
    counter = 0
    While counter < 3
        counter = counter + 1
        With empower_report

            ' This function let you run the code from any worksheet in the workbook
            Application.ScreenUpdating = True
            Worksheets("empower_report").Activate
            Application.ScreenUpdating = False

            ' This to expand the entire sheet
            Cells.EntireColumn.AutoFit

            ' This delete the row if the cell contains zero in the first column:
            ' How many rows in worksheet
            myRow = 200
            ' loop through all the rows until the specified value,its counting from  bottom to top by deducting one in each step.
            ' (my row to 2) to exclude first row from being deleted if it have specified criteria
            For myCounter = myRow To 2 Step -1
              If Cells(myCounter, 1).Value = 0 Then
                  Rows(myCounter).Delete
              End If
            Next

            ' This clear contents in a given range if specific value exist
            empower_rng = getColRange("Empower Address 2")
            For Each Cell In Range(empower_rng)
                If Cell = 0 Then Cell.ClearContents
            Next Cell

            ' Sneaky Part range requires column name so we used colstr function to give it the range and it will return column name
            address1_rng = getColStr("Empower Address 1")
            address2_rng = getColStr("Empower Address 2")
            'This is used to remove string/value from one column (ex:address1)and add it to the end of string/value in another column(ex:Address2)
            For i = 2 To Range(address1_rng & Rows.Count).End(xlUp).Row
                Range(address1_rng & i).Value = Range(address1_rng & i).Value & " " & Range(address2_rng & i).Value
                Range(address2_rng & i).Clear
            Next i

            ' This to change the string in a given range into a proper case format.
            bos_selection = getColRange("BOS Address 1")
            Set bos_txt_fix = Range(bos_selection)
            For Each bos_cell In bos_txt_fix
                bos_cell.Value = WorksheetFunction.Proper(bos_cell.Value)
            Next

            emp_selection = getColRange("Empower Address 1")
            Set emp_txt_fix = Range(emp_selection)
            For Each emp_cell In emp_txt_fix
                emp_cell.Value = WorksheetFunction.Proper(emp_cell.Value)
            Next

            ' This to delete trail space in range
            emp_add = getColRange("Empower Address 1")
            ' To 'Find the last used cell in Col A
            trim_row = Range(emp_add).End(xlDown).Row
            '    MsgBox (trim_row)
            'Declare the range used by having the coordinates of rows and column till the last cell used.
            Set trim_rng = Range(Cells(2, 4), Cells(trim_row, 4))
            ' Loop through the range and remove any trailing space
            For Each trimcell In trim_rng
                trimcell = RTrim(trimcell)
            'Go to the next Cell
            Next trimcell

            
            ' This part to get the range for column C and D
            empcol_rng = getColRange("Empower Address 1")
            boscol_rng = getColRange("BOS Address 1")
            
            ' get reference for Column C and D
            Set first_col = Range(empcol_rng)
            Set second_col = Range(boscol_rng)
            ' This is used if Ranges are not aligned
            If first_col.Row <> second_col.Row Then
                Exit Sub
            End If
            
            ' This is used if Ranges are not the same size
            If first_col.Rows.Count <> second_col.Rows.Count Then
                Exit Sub
            End If
        
             'Loop the array
            For n = first_col.Rows.Count To 1 Step -1
                 'Detect if value/string in column C and D on the same row are equal.  If it matches delete the entire row
                If first_col.Cells(n, 1) = second_col.Cells(n, 1) Then
                    first_col.Rows(n).EntireRow.Delete
                End If
            Next
        End With
    Wend
End Sub




