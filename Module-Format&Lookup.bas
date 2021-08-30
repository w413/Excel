Attribute VB_Name = "Module2"
Sub FormatData()
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

Call format

'Restore Application settings
Application.ScreenUpdating = screenUpdateState
Application.DisplayStatusBar = statusBarState
Application.Calculation = calculationState
Application.EnableEvents = eventsState

End Sub
Sub format()
'
'Format data for lookups
'

Dim fName As String
Dim row, rowTotal, colTotal 'Define Column and Row indexes as well as temp variables to save notable indices
row = 1 'Set Row Index

' Select the sheet and first cell where data should start
Sheets("Paste Here").Activate
Worksheets("Paste Here").Cells(1, 1).Select

If ActiveCell.FormulaR1C1 <> "" Then 'As long as cell A1 is not empty do the following

    ' Before starting anything lets get some user data
    Dim userInput As String
    fName = Left(Worksheets("Paste Here").Cells(2, 1).FormulaR1C1, Len(Worksheets("Paste Here").Cells(2, 1).FormulaR1C1) - 3) + "_Lookups"
    strInput = InputBox("What would you like to call your file?", "File Name", fName)
    
    ' Using input create file name
    fName = strInput
    ext = 52
    fName = ActiveWorkbook.Path + "\" + fName + ".xlsm"
    
    'Stack columns and copy column one all the way down to correspond with stacked columns
    Call stackCols("Paste Here")
    
    'Sort Columns in descending order with respect to column A
    Call sortDescend("Paste Here", "A:A")
    
    ' Remove Empty Rows and No Equipment rows
    Call removeRows("Paste Here", 2, "")
    Call removeRows("Paste Here", 2, "No Equipment*")
    
    ' Move formatted data from Paste Here to Finished
    Application.CutCopyMode = False
    Sheets("Paste Here").Activate
    Worksheets("Paste Here").Range("A2:B2", Range("A2:B2").End(xlDown)).Cut Destination:=Sheets("Finished").Cells(2, 1)
    ' End move formatted data
    
    ' Colorise the raw data we moved
    Call ColorizeIt
    
    ' Pull Headers and functions from Headers & Formulas sheet
    Call getHeadersAndFuncs
    
    ' Clear Paste Here sheet
    Call Clear("Paste Here")
    
    'Update the values of rowTotal and rowTotal to drag formulas across the whole sheet
    colTotal = getColCount("Finished", 1, 3)
    rowTotal = getRowCount("Finished", 2, 1)
    
    'Drag formulas to all rows
    Call formulaDrag(2, 3, rowTotal + 1, colTotal + 1)
    
    ' Save changes to a new workbook
    Worksheets("Finished").Visible = True
    Worksheets("Finished").Copy
    ActiveWorkbook.SaveAs Filename:=fName, FileFormat:=ext
    ActiveWorkbook.Close Savechanges:=False
    Worksheets("Finished").Visible = False
    
    ' Clean up
    Call Clear("Finished")
    Sheets("Paste Here").Activate
    Worksheets("Paste Here").Range("A1").Select

Else
    MsgBox "Please Paste Scan Data on Sheet Paste Here Starting in Cell A1"
End If
    
End Sub
Private Sub Clear(SheetToClear)
'
' Clear Macro
' Clean selected sheet
'

'
    Sheets(SheetToClear).Activate
    Cells.Clear
End Sub
Private Sub SortColTwo()
'
' SortColTwo Macro
' Sort the Data by Column Two
' Recorded macro
'

'
    ActiveWorkbook.Worksheets("Paste Here").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Paste Here").Sort.SortFields.Add2 Key:=Range( _
        "B:B"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Paste Here").Sort
        .SetRange Range("A:B")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Private Sub getHeadersAndFuncs()
'
' GetHeaders Macro
'
' Static macro to retrieve headers from Sheet "Headers & Formulas" and place them in "Finished"
'
    Dim endCol, temp
    
    'Get headers
    endCol = getColCount("Headers & Formulas", 1, 1)
    Application.CutCopyMode = False
    Worksheets("Headers & Formulas").Range(Cells(1, 1), Cells(1, endCol)).Copy Destination:=Sheets("Finished").Range("A1")
    
    'Get formulas
    Application.CutCopyMode = False
    Worksheets("Headers & Formulas").Range(Cells(2, 3), Cells(2, endCol)).Copy Destination:=Sheets("Finished").Range("C2")
    
End Sub
Private Function getRowCount(sheet, row, col)
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
Dim total
total = 0

    Sheets(sheet).Activate
    Worksheets(sheet).Cells(row, col).Select
    While ActiveCell.FormulaR1C1 <> ""
        col = col + 1 'Increment the columns
        total = total + 1
        Worksheets(sheet).Cells(row, col).Select 'Select the new cell
    Wend
    getColCount = total
    
End Function
Private Sub ColorizeIt()
'
' ColorizeIt Macro
' Static macro to color scan data
'

'
    Sheets("Finished").Activate
    Worksheets("Finished").Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
Private Sub formulaDrag(startRow, startCol, endRow, endCol)
'
' formulaDrag Macro
'

'
    Range(Cells(startRow, startCol), Cells(startRow, endCol)).AutoFill Destination:=Range(Cells(startRow, startCol), Cells(endRow, endCol)), Type:=xlFillDefault

   ' Selection.AutoFill Destination:=Range(Cells(startRow, startCol), Cells(endRow, endCol)), Type:=xlFillDefault

End Sub
Private Sub stackCols(sheet)
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
    Dim infinityCheck As Integer
    infinityCheck = 0
    
    Sheets(sheetName).Activate
    
    rowTotal = getRowCount(sheetName, 1, 1)
    For row = 1 To rowTotal 'Find first row with empty column B
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
Private Sub sortDescend(sheet, col)
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


