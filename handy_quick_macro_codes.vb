' Column Width Adjustment
Worksheets("working_sheet").Columns("C:AA").ColumnWidth = 24
Worksheets("working_sheet").Columns("J").ColumnWidth = 22.11

' Font Color
Columns(57).Font.Color = RGB(255, 255, 255)

'Copy Rows and Transpose
Sheets("Sheet1").Range("named_range").Copy
Sheets("copy_to_sheet").Range("J10").PasteSpecial xlPasteValues, Transpose:=True


Application.ScreenUpdating = False

'Calculation control
Application.Calculation = xlCalculationManual
Application.Calculation = xlCalculationAutomatic


'Sheet Hide/Unhide
Sheets("Sheet1").Visible = True
Sheets("Sheet1").Select


'Last Row
LastRow = Sheets("Sheet1").Range("E" & Rows.Count).End(xlUp).Row

'Last Column
colNum = Worksheets("Sheet1").Cells.SpecialCells(xlLastCell).Column

'Clear Contents of cells
Sheets("Sheet1").Range("F:F").ClearContents

'Add formulas
Sheets("Sheet1").Range("F2:F" & LastRow).Formula = "=IFERROR(VLOOKUP(E2,$N$2:$O$999,2,0),0)"


'Copy Paste Formattings
Sheets("Sheet1").Select
Sheets("Sheet1").Range("A9:BD9").PasteSpecial xlPasteFormats

'Copy Paste using Named Ranges
Sheets("Sheet1").Range("copy_from_named_range").Copy
Sheets("Sheet1").Range("copy_TO_named_range").PasteSpecial xlPasteValues

