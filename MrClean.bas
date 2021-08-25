Attribute VB_Name = "Module11"
Sub CleanJunk()
Dim regexVal, sheetName As String
Dim fname, ext

regexVal = "[A-Z][A-Z]=*"
sheetName = ActiveSheet.Name
'Value 52 corresponds to extension .xlsm which is an xlExcel12 or a binary workbook
ext = 50
fname = ActiveWorkbook.Path + "\" + sheetName + ".xlsb"
''Value 52 corresponds to extension .xlsm which is an xlOpenXMLWorkbook
'ext = 51
'fname = ActiveWorkbook.Path + "\" + sheetName + ".xlsx"
''Value 52 corresponds to extension .xlsm which is an xlOpenXMLWorkbookMacroEnabled
'ext = 52
'fname = ActiveWorkbook.Path + "\" + sheetName + ".xlsm"

Call SortCols(ActiveSheet.Name, "V:V")
Call removeRows(ActiveSheet.Name, 22, regexVal)
Range("B:B,D:D,F:P,R:Z,AC:AI,AK:AN,AR:BD,BG:BL,BN:BN,BP:BT,BV:BV,BX:CD,CG:CH,CJ:CU,CW:CX,CZ:DI,DK:EM,EO:EP,ER:EV,EX:FC,FE:FZ,GB:GT").Delete Shift:=xlToLeft
Call SortCols(ActiveSheet.Name, "A:A")

ActiveWorkbook.SaveAs Filename:=fname, FileFormat:=ext

End Sub
Private Sub SortCols(sheet, rng)
'
' SortColTwo Macro
' Sort the Data by Column Two
' Recorded macro
'

'
    ActiveWorkbook.Worksheets(sheet).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(sheet).Sort.SortFields.Add2 _
        Key:=Range(rng), SortOn:=xlSortOnValues, Order:=xlAscending, _
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
Sub removeRows(sheetName, col As Integer, contains)
    Dim infinityCheck As Integer
    infinityCheck = 0
    
    Sheets(sheetName).Activate
    
    rowTotal = getRowCount(sheetName, 1, col)
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
Private Function getRowCount(sheet, row, col)
Dim total
total = 0
    
    Sheets(sheet).Activate
    Worksheets(sheet).Cells(row, col).Select
    While ActiveCell.FormulaR1C1 <> "" 'Get the number of rows we are working with
        row = row + 1 'Increment the rows
        total = total + 1
        Worksheets(sheet).Cells(row, col).Select
    Wend
    getRowCount = total
    
End Function
