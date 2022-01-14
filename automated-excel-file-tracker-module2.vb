Option Explicit

Dim main_workbook_name As String, datetime_now As String, file_path_name As String




Sub initialize_choices()
    
    
    Sheets.Add After:=Worksheets("Data")
    ActiveSheet.Name = "Temp"


    Worksheets("Data").Activate
    Range("E4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy Worksheets("Temp").Range("A1")
    Worksheets("Temp").Activate
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.RemoveDuplicates Columns:=1, Header:=xlNo
    
    

End Sub


Sub filter_report1(company As String, Optional status As String = "All")

    Sheets("Data").Activate
    Range("A1:P1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.AutoFilter Field:=5, Criteria1:= _
        company
    If status = "Ag+" Then Selection.AutoFilter Field:=9, Criteria1:=status
    If status = "C+" Then Selection.AutoFilter Field:=9, Criteria1:=status
    
    Range("A1").CurrentRegion.Select
    Selection.Copy
    Sheets.Add After:=Worksheets("Data")
    ActiveSheet.Name = company
    Range("A1").PasteSpecial
    Columns("A:P").EntireColumn.AutoFit
    
    'Save the name of the current workbook (DRF tracker) to a variable for reference later
    main_workbook_name = ThisWorkbook.Name
    
    
    
    Sheets(company).Move
    'MsgBox ThisWorkbook.Name
    
    'find the current date and time to be used for the file name
    'Use the excel worksheet function to convert the date type value to string
    datetime_now = Application.WorksheetFunction.Text(Now, "DD-MM-YYYY hh-mm-ss")
    
    'defining the file path based on the active workbook location
    file_path_name = Workbooks(main_workbook_name).Path & "/" & status & " occupants for " & company & " as at " & datetime_now
    
    'choose the default option for the macro-enabled workbook display alert when saving
    Application.DisplayAlerts = False
    
    'ThisWorkbook.SaveCopyAs file_path_name
    ' save the workbook as a normal excel workbook in the same directory
    ActiveWorkbook.SaveAs Filename:=file_path_name, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
    
End Sub


Sub filter_report2(companies, Optional status As String = "All")

    Sheets("Data").Activate
    Range("A1:P1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.AutoFilter Field:=5, Criteria1:= _
        companies, Operator:=xlFilterValues
    If status = "Ag+" Then Selection.AutoFilter Field:=9, Criteria1:=status
    If status = "C+" Then Selection.AutoFilter Field:=9, Criteria1:=status
    
    Range("A1").CurrentRegion.Select
    Selection.Copy
    Sheets.Add After:=Worksheets("Data")
    ActiveSheet.Name = "Multiple companies"
    Range("A1").PasteSpecial
    Columns("A:P").EntireColumn.AutoFit
    
    'Save the name of the current workbook (DRF tracker) to a variable for reference later
    main_workbook_name = ThisWorkbook.Name
    
    
    
    Sheets("Multiple companies").Move
    'MsgBox ThisWorkbook.Name
    
    'find the current date and time to be used for the file name
    'Use the excel worksheet function to convert the date type value to string
    datetime_now = Application.WorksheetFunction.Text(Now, "DD-MM-YYYY hh-mm-ss")
    
    'defining the file path based on the active workbook location
    file_path_name = Workbooks(main_workbook_name).Path & "/" & status & " occupants for multiple companies as at " & datetime_now
    
    'choose the default option for the macro-enabled workbook display alert when saving
    Application.DisplayAlerts = False
    
    'ThisWorkbook.SaveCopyAs file_path_name
    ' save the workbook as a normal excel workbook in the same directory
    ActiveWorkbook.SaveAs Filename:=file_path_name, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

End Sub
