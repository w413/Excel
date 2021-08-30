Attribute VB_Name = "Module3"
Private Sub Clear(SheetToClear)
'
' Clear Macro
' Clean selected sheet of all data and formatting
'

    Sheets(SheetToClear).Activate
    Cells.Clear
End Sub
Private Function getRowCount(sheet, row, col)
'
' getRowCount Macro
' Count rows with data, assumes first blank is end of data
' Counts downward
'

Dim total
total = 0
    
    Sheets(sheet).Activate
    Worksheets(sheet).Cells(row, col).Select
    While ActiveCell.FormulaR1C1 <> ""
        row = row + 1 'Increment the rows
        total = total + 1
        Worksheets(sheet).Cells(row, col).Select
    Wend
    getRowCount = total
    
End Function
Private Function getColCount(sheet, row, col)
'
' GetColCount Macro
' Counts cols with data, assumes first blank is end of data
' Counts righward
'
Dim total
total = 0

    Sheets(sheet).Activate
    Worksheets(sheet).Cells(row, col).Select
    While ActiveCell.FormulaR1C1 <> ""
        col = col + 1
        total = total + 1
        Worksheets(sheet).Cells(row, col).Select
    Wend
    getColCount = total
    
End Function
Private Sub drag(startRow, startCol, endRow, endCol)
'
' drag Macro
' drags the selection, startRow & startCol = top left endRow & endCol = bottom right
'
'

'
    Range(Cells(startRow, startCol), Cells(startRow, endCol)).AutoFill Destination:=Range(Cells(startRow, startCol), Cells(endRow, endCol)), Type:=xlFillDefault

End Sub
Private Sub stackCols(sheet)
'
' stackCols Macrp
' Stacks all Columns right of Column A into column B and copies Column A to end of sheet
'
Sheets(sheet).Activate
Dim row, col, rowTemp, colTemp, rowTotal
row = 1
col = 1
Cells(row, col).Select
While ActiveCell.FormulaR1C1 <> "" 'While loop to insert Column A in between each column
        Application.CutCopyMode = False
        Columns(col).Select 'Select the Column to paste
        Selection.Copy 'Copy the selection to paste
        col = col + 2 'Increment Column index to paste location
        Columns(col).Select 'Select the Column to paste location
        Selection.Insert Shift:=xlToRight 'Paste and shift columns to the right
        Cells(row, col + 2).Select 'Select the check cell so the while loop exits properly
    Wend 'Value of Col will be the value of last pasted column after this loop
    colTemp = col 'Save the ending column index
    col = 1

    'The next Two while loops can benefit from the following code
    ' Range(Selection, Selection.End(xlDown)).Select
    'This appears to select a range from current selection until the end of the data, further testing needed

    rowTotal = getRowCount("Paste Here", 1, 1) 'Call getRowCount which stores in global var rowTotal how many contiguous rows of data there are, starting from Row Col
    rowTemp = rowTotal
       
    While colTemp > 1 'While loop to consilidate all Columns into Column A and B
        Application.CutCopyMode = False
        Range(Cells(2, colTemp), Cells(rowTemp, colTemp + 1)).Select 'Select the range of cells to be cut, omittinig the header row
        Selection.Cut 'Cut the selection
        Cells(rowTemp + 1, 1).Select 'Select cells to paste the cut
        ActiveSheet.Paste 'Paste Cut cells into new selection
        rowTemp = rowTemp + rowTotal - 1 'Keep track of where we are in Column A
        colTemp = colTemp - 2 'Keep track of which Columns we have pasted
    Wend
End Sub
Sub removeRows(sheetName, col As Integer, contains)
'
' removeRows macro
' removes rows whose col contains
' Assumes data is sorted by Column A
'
    Dim infinityCheck As Integer
    infinityCheck = 0
    
    Sheets(sheetName).Activate
    
    rowTotal = getRowCount(sheetName, 1, 1)
    For row = 1 To rowTotal
        infinityCheck = infinityCheck + 1
        Worksheets(sheetName).Cells(row, col).Select
        If ActiveCell.FormulaR1C1 Like contains Then
            Worksheets(sheetName).Rows(row).Select
            Selection.Delete Shift:=xlUp
            row = row - 1
        End If
        If infinityCheck > rowTotal Then
            Exit For
        End If
    Next row
End Sub
Function mlookup(search_val As Variant, search_space As Range, return_space As Range, if_not_found)
'This function will only work on column data
' This function assumes data in search_space is sorted

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
Sub IterateCells()
' Macro prototype for iterating through all used cells
   For Each Cell In ActiveSheet.UsedRange.Cells
      'do some stuff
   Next

End Sub
Sub findAndReplace(sheet, find, replace)
' Simple find and replace Macro
' Searches entire sheet for find and replaces with replace
' May crash excel with large data
    Sheets(sheet).Activate
    For Each Cell In ActiveSheet.UsedRange.Cells
        If Cell.Value Like find Then
            Cell.Value = replace
        End If
    Next
End Sub
Sub saveHereAs(fname, ext)
'
' Work in progress
' Doesn't save a usable file
'
ActiveWorkbook.SaveAs Filename:=ActiveWorkbook.Path + fname + ext
End Sub

Private Sub SortAscend(sheet, col)
'
' SortColTwo Macro
' Sort all data in sheet in ascending order with respect to col
'
'

'
    ActiveWorkbook.Worksheets(sheet).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(sheet).Sort.SortFields.Add2 _
        Key:=Range(col), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(sheet).Sort
        .SetRange Cells
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

