Option Explicit



Sub time_report():
    
    'launch the create_report sub-procedure automatically based on a particular date and time
    'Date format to be in US date format
    Application.OnTime DateValue("16/1/2022 4.00pm"), "create_report"
    
    'launch the create_report sub-procedure automatically based on a particular time daily
    Application.OnTime TimeValue("3:00:00pm"), "create_report"

End Sub



Sub create_report()

    'a procedure to create an excel copy of the records of current occupants
    
    Dim datetime_now As String
    Dim file_path_name As String
    Dim main_workbook_name As String
    
    
    'making a copy of the main data worksheet in the same workbook on another sheet
    Sheets("Data").Activate
    Sheets("Data").Select
    Sheets("Data").Copy After:=Sheets(2)
    Sheets("Data (2)").Select
    'renaming the copy as "DRF Records"
    Sheets("Data (2)").Name = "DRF Records"
    
    'copy and paste values only to remove formulas in the cells
    Sheets("DRF Records").Activate
    Range("A1:P1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    
    'Removing rows for workers who have already checked out
    Dim lastRow As Long
    Dim number_check_outs As Integer
    
    'Find the last row number
    Range("A4").Select
    Selection.End(xlDown).Select
    lastRow = ActiveCell.Row
    'alternative code in the event of empty spaces in between data
    'lastRow = Cells(Rows.count, 1).End(xlUp).Row
    'Find the total number of check-outs
    number_check_outs = Application.WorksheetFunction.CountA(Range(Cells(4, 8), Cells(lastRow, 8)))
  
    'sort the column values with non-blank cells first so that we can iterate through each row with check-outs to remove them
    ActiveWorkbook.Worksheets("DRF Records").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("DRF Records").AutoFilter.Sort.SortFields.Add2 Key _
        :=Range(Cells(3, 8), Cells(lastRow, 8)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("DRF Records").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'delete the rows with check-outs using a for loop
    Dim i As Integer
    For i = 1 To number_check_outs
        Rows("4:4").Delete
    Next i
    
    
    
    'Save the name of the current workbook (DRF tracker) to a variable for reference later
    main_workbook_name = ThisWorkbook.Name
    
    
    
    Sheets("DRF Records").Move
    
    'MsgBox ThisWorkbook.Name
    
    'find the current date and time to be used for the file name
    'Use the excel worksheet function to convert the date type value to string
    datetime_now = Application.WorksheetFunction.Text(Now, "DD-MM-YYYY hh-mm-ss")
    
    
    'defining the file path based on the active workbook location
    file_path_name = Workbooks(main_workbook_name).Path & "/DRF current occupants as at " & datetime_now
    
    'choose the default option for the macro-enabled workbook display alert when saving
    Application.DisplayAlerts = False
    
    'ThisWorkbook.SaveCopyAs file_path_name
    ' save the workbook as a normal excel workbook in the same directory
    ActiveWorkbook.SaveAs Filename:=file_path_name, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
  
    
End Sub


Sub main_procedure()
    'This is the main sub procedure for activating the user dialog box
    
    UserForm1.Show
    
End Sub






