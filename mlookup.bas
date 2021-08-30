Attribute VB_Name = "Module1"
Function mlookup(search_val As Variant, search_space As Range, return_space As Range, if_not_found)
'This function has only works on Column data

Dim screenUpdateState, statusBarState, calculationState, eventsState

'Save application settings
screenUpdateState = Application.ScreenUpdating
statusBarState = Application.DisplayStatusBar
calculationState = Application.Calculation
eventsState = Application.EnableEvents

'Set applications settings to enhance performance
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

Dim retval() 'Return value
Dim i 'Lower bound variable used in For loop
Dim j 'Upper bound variable used in For loop
j = search_space.Rows.Count 'initialize upper bound variable to how many rows are in our search space
Dim found 'Variable to keep track of how many matches we have
found = 0 'Initialize found to 0

On Error GoTo ErrorHandle
For i = WorksheetFunction.match(search_val, search_space, 0) To j
'Using match to get us close we look record the first and any proceeding matches quitting as soon as we find something that doesn't match
    If search_space(i).Value = search_val Then 'If we found our value record it in retval
        ReDim Preserve retval(0 To found)
        retval(found) = return_space(i).Value
        found = found + 1
    End If
Next i

'Restore Application settings
Application.ScreenUpdating = screenUpdateState
Application.DisplayStatusBar = statusBarState
Application.Calculation = calculationState
Application.EnableEvents = eventsStat
mlookup = retval 'return retval

Exit Function

ErrorHandle:
    'Restore Application settings
    Application.ScreenUpdating = screenUpdateState
    Application.DisplayStatusBar = statusBarState
    Application.Calculation = calculationState
    Application.EnableEvents = eventsStat
    mlookup = if_not_found

End Function
