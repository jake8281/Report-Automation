'Created By Jake Ayoub 6/3/2021
'Updated 6/4/2021

Sub restart_process()
    Dim new_sheet As String
    
    With piled_data
        ' This function let you run the code from any worksheet in the workbook
        Application.ScreenUpdating = True
        Worksheets("piled_data").Activate
        Application.ScreenUpdating = False
        ' This clear contents excluding first row
        Cells.ClearContents
    End With

    With onbase_data
        ' This function let you run the code from any worksheet in the workbook
        Application.ScreenUpdating = True
        Worksheets("onbase_data").Activate
        Application.ScreenUpdating = False
        ' This clear contents excluding first row
        Rows("2:" & Rows.Count).ClearContents
    End With
        
    With bulkDemo_data
        ' This function let you run the code from any worksheet in the workbook
        Application.ScreenUpdating = True
        Worksheets("bulkDemo_data").Activate
        Application.ScreenUpdating = False
            
        Rows("2:" & Rows.Count).ClearContents
    End With
            
     With combine_report
        new_sheet = "empower_report"
         ' This Function delete the sheet if it has already existed.
        For i = 1 To Worksheets.Count
            If Worksheets(i).Name = new_sheet Then
                Worksheets(i).Delete
            End If
        Next
     End With
    MsgBox ("Have a nice day!!!")
End Sub
