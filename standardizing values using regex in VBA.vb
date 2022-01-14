Option Explicit


'define the function to be private for access only within the module
'function will not be available as user-defined function in worksheet

Private Function test_value(expression As String, value As String) As Boolean
    'A function to test for the string expression in the argument against a regular expression rule
    'It will return True if matches
    Dim regexObject As RegExp
    
    Set regexObject = New RegExp
    
    With regexObject
        .Pattern = expression
    End With
    
    test_value = regexObject.Test(value)

End Function


Function test_nationality(nationality As String) As String
    'A function to test and convert the nationality value given in the argument
    'It will be converted into a set of standardized nationality values
    If test_value("[d,D][e,E][s,S][h,H]|[b,B][a,A][n,N][g,G]", nationality) = True Then
        test_nationality = "Bangladesh"
    ElseIf test_value("[i,I][n,N][d,D]", nationality) = True Then
        test_nationality = "India"
    ElseIf test_value("[c,C][h,H]|[p,P][r,R][c,C]", nationality) = True Then
        test_nationality = "China"
    ElseIf test_value("[m,M][y,Y]|[b,B][u,U]", nationality) = True Then
        test_nationality = "Myanmar"
    ElseIf test_value("[t,T][h,H][a,A][i,I]", nationality) = True Then
        test_nationality = "Thailand"
    ElseIf test_value("[m,M][a,A][l,L][a,A][y,Y]", nationality) = True Then
        test_nationality = "Malaysia"
    ElseIf test_value("[v,V][i,I][e,E][t,T]", nationality) = True Then
        test_nationality = "Vietnam"
    ElseIf test_value("[s,S][i,I][n,N]", nationality) = True Then
        test_nationality = "Singaporean PR"
    ElseIf test_value("[f,F][i,I][p,P][i,I]", nationality) = True Then
        test_nationality = "Filipino"
    ElseIf test_value("[p,P][i,I][l,L][i,I]", nationality) = True Then
        test_nationality = "Filipino"
    Else
        test_nationality = nationality
    End If
    
End Function


Sub convert_values()

    'loop through each value in the Nationality column to standardize the values
    Dim cell As Range
    
    'Ignore any error, eg. errors like wrong value and data type
    On Error Resume Next
    
    For Each cell In Sheets("Data").Range("C:C").SpecialCells(xlCellTypeConstants)
        cell.value = test_nationality(cell.value)
    Next cell
    
End Sub
