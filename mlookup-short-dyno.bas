Attribute VB_Name = "mlookup"
Function mlookup(search_val As Variant, search_space As Range, return_space As Range, if_not_found)
'This function will only work on column data

Dim retval() 'Return value
Dim i 'Lower bound variable used in For loop
Dim j 'Upper bound variable used in For loop
j = search_space.Rows.Count 'initialize upper bound variable to how many rows are in our search space
Dim found 'Variable to keep track of how many matches we have
found = 0 'Initialize found to 0

On Error GoTo ErrorHandle
For i = WorksheetFunction.Match(search_val, search_space, 0) To j
'Using match to get us close we look record the first and any proceeding matches quitting as soon as we find something that doesn't match
    If search_space(i).Value = search_val Then 'If we found our value record it in retval
        ReDim Preserve retval(0 To found)
        retval(found) = return_space(i).Value
        found = found + 1
    Else
        Exit For
    End If
Next i

mlookup = retval 'return retval

Exit Function

ErrorHandle:
    mlookup = if_not_found

End Function
