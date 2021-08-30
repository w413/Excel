Attribute VB_Name = "Module11"
Sub mrClean(infname As String, outfname As String, ext As String)
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

Call cleanMI_Junk(infname, outfname, ext)

'Restore Application settings
Application.ScreenUpdating = screenUpdateState
Application.DisplayStatusBar = statusBarState
Application.Calculation = calculationState
Application.EnableEvents = eventsState

End Sub
Private Sub cleanMI_Junk(MI_wbName As String, MI_saveName As String, ext As String)

Dim regexVal As String
Dim fname, extVal
Dim MI_wb As Workbook

regexVal = "[A-Z][A-Z]=*"

fname = ActiveWorkbook.path + "\" + MI_wbName + "." + ext
Set MI_wb = Workbooks.Open(fname) 'Open workbook name specified in cell B2 in Format.MI.Data
MI_wb.Activate 'Set the workbook as active

'Hopefully clean up junk information. The folloing two lines may need a more intellegent revision
Call SortCols(ActiveSheet.Name, "V:V")
Call removeRows(ActiveSheet.Name, 22, regexVal)

'Remove uneeded columns, these shouldn't change from export to export *Fingers crossed*
Range("B:B,D:D,F:P,R:Z,AC:AI,AK:AN,AR:BD,BG:BL,BN:BN,BP:BT,BV:BV,BX:CD,CG:CH,CJ:CU,CW:CX,CZ:DI,DK:EM,EO:EP,ER:EV,EX:FC,FE:FZ,GB:GT").Delete Shift:=xlToLeft
Call SortCols(ActiveSheet.Name, "Z:Z") 'Sort by serial number for lookup sheet
Call findAndReplace(ActiveSheet.Name, regexVal, "") 'Find and replace anything that matches the regex value specified above


'Comment and uncomment the following lines to change save extension
'Value 50 corresponds to extension .xlsb which is an xlExcel12 or a binary workbook
extVal = 50
fname = ActiveWorkbook.path + "\" + MI_saveName + ".xlsb"
''Value 52 corresponds to extension .xlsm which is an xlOpenXMLWorkbook
'extVal = 51
'fname = ActiveWorkbook.Path + "\" + MI_saveName + ".xlsx"
''Value 52 corresponds to extension .xlsm which is an xlOpenXMLWorkbookMacroEnabled
'extVal = 52
'fname = ActiveWorkbook.Path + "\" + MI_saveName + ".xlsm"

'Save changes to new workbook
MI_wb.SaveAs Filename:=fname, FileFormat:=extVal

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
