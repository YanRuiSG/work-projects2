Option Explicit

Private Sub CancelButton_Click()

    Application.DisplayAlerts = False
    
    Worksheets("Temp").Delete
    
    Unload UserForm1 'unloads the dialog box from memory
    
    Application.DisplayAlerts = True
    
End Sub



Private Sub OKButton_Click()

    'Determine if more than 1 company is selected
    Dim counter As Integer
    Dim i As Integer
    Dim Msg As String, company_selected As String
    
    'count the number of companies selected
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) Then counter = counter + 1
    Next i
    
    
    Select Case counter
        Case 0
            MsgBox "No company is selected"
            'set the focus back to the dialog box in case the users wants to try again
            OKButton.SetFocus
        Case 1
            For i = 0 To ListBox1.ListCount - 1
                If ListBox1.Selected(i) Then
                    company_selected = ListBox1.List(i)
                End If
            Next i
            If AgPositiveButton.Value = True Then
                Call Module2.filter_report1(company_selected, "Ag+")
            ElseIf PCRpositiveButton.Value = True Then
                Call Module2.filter_report1(company_selected, "C+")
            Else
                Call Module2.filter_report1(company_selected)
            End If
        
            MsgBox "Company selected is " & company_selected
        Case Else
            'create the array of selected companies
            Dim x() As String
            ReDim x(1 To counter)
            Dim counter2 As Integer
            'Adding each selected company to the empty array
            For i = 0 To ListBox1.ListCount - 1
                If ListBox1.Selected(i) Then
                    Msg = Msg & ListBox1.List(i) & vbNewLine
                    counter2 = counter2 + 1
                    x(counter2) = ListBox1.List(i)
                End If
            Next i
            
            If AgPositiveButton.Value = True Then
                Call Module2.filter_report2(x, "Ag+")
            ElseIf PCRpositiveButton.Value = True Then
                Call Module2.filter_report2(x, "C+")
            Else
                Call Module2.filter_report2(x)
            End If
            
            
            MsgBox Msg
            MsgBox UBound(x) - LBound(x) + 1
    
    End Select
    
    
    Workbooks("Dormitory Recovery Facility Tracker").Activate
    
    Application.DisplayAlerts = False
    
    Worksheets("Temp").Delete
    
    Worksheets("Data").ShowAllData
    
    Unload UserForm1 'unloads the dialog box from memory
    
    Application.DisplayAlerts = True
    

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()


    Call Module2.initialize_choices
    
    Dim companies As Range
    Worksheets("Temp").Activate
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Set companies = Selection
    
    'MsgBox companies.Address
    ListBox1.RowSource = Selection.Address
    
    Worksheets("Data").Activate
    ActiveSheet.Range("F20").Select
    
    
End Sub


