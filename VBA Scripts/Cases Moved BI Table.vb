'Cases Moved BI
Sub Button8_Click()
Dim CellValue1 As String, CellValue2 As String, table As ListObject

CellValue1 = Sheets("Home").Range("F7").Value
CellValue2 = Sheets("Home").Range("F8").Value

If CellValue1 = "Pearmund Cellars" Then Set table = Sheets("PC Inventory").ListObjects("PCInv")
If CellValue1 = "Effingham Manor" Then Set table = Sheets("EF Inventory").ListObjects("EFInv")
If CellValue1 = "Vint Hill" Then Set table = Sheets("VH Inventory").ListObjects("VHInv")

Dim SumRng As Range, Crit1_Rng As Range, Crit1 As String, Crit2_Rng As Range, Crit3_Rng As Range, Crit2 As Variant, NumYears As Long

Set SumRng = table.ListColumns(5).DataBodyRange
Set Crit1_Rng = table.ListColumns(3).DataBodyRange
Set Crit2_Rng = table.ListColumns(4).DataBodyRange
Set Crit3_Rng = table.ListColumns(1).DataBodyRange
NumYears = Year(Date) - 2016

'January
Sheets("Home").Range("I8") = (Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CellValue2, Crit2_Rng, 1) / NumYears)
Sheets("Home").Range("I9") = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CellValue2, Crit2_Rng, 1, Crit3_Rng, Year(Date))

'February
Sheets("Home").Range("J8") = (Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CellValue2, Crit2_Rng, 2) / NumYears)
Sheets("Home").Range("J9") = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CellValue2, Crit2_Rng, 2, Crit3_Rng, Year(Date))

'March
Sheets("Home").Range("K8") = (Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CellValue2, Crit2_Rng, 3) / NumYears)
Sheets("Home").Range("K9") = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CellValue2, Crit2_Rng, 3, Crit3_Rng, Year(Date))

'April
Sheets("Home").Range("L8") = (Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CellValue2, Crit2_Rng, 4) / NumYears)
Sheets("Home").Range("L9") = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CellValue2, Crit2_Rng, 4, Crit3_Rng, Year(Date))

'May
Sheets("Home").Range("M8") = (Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CellValue2, Crit2_Rng, 5) / NumYears)
Sheets("Home").Range("M9") = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CellValue2, Crit2_Rng, 5, Crit3_Rng, Year(Date))

'June
Sheets("Home").Range("N8") = (Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CellValue2, Crit2_Rng, 6) / NumYears)
Sheets("Home").Range("N9") = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CellValue2, Crit2_Rng, 6, Crit3_Rng, Year(Date))

'July
Sheets("Home").Range("O8") = (Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CellValue2, Crit2_Rng, 7) / NumYears)
Sheets("Home").Range("O9") = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CellValue2, Crit2_Rng, 7, Crit3_Rng, Year(Date))

'August
Sheets("Home").Range("P8") = (Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CellValue2, Crit2_Rng, 8) / NumYears)
Sheets("Home").Range("P9") = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CellValue2, Crit2_Rng, 8, Crit3_Rng, Year(Date))

'September
Sheets("Home").Range("Q8") = (Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CellValue2, Crit2_Rng, 9) / NumYears)
Sheets("Home").Range("Q9") = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CellValue2, Crit2_Rng, 9, Crit3_Rng, Year(Date))

'October
Sheets("Home").Range("R8") = (Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CellValue2, Crit2_Rng, 10) / NumYears)
Sheets("Home").Range("R9") = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CellValue2, Crit2_Rng, 10, Crit3_Rng, Year(Date))

'November
Sheets("Home").Range("S8") = (Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CellValue2, Crit2_Rng, 11) / NumYears)
Sheets("Home").Range("S9") = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CellValue2, Crit2_Rng, 11, Crit3_Rng, Year(Date))

'December
Sheets("Home").Range("T8") = (Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CellValue2, Crit2_Rng, 12) / NumYears)
Sheets("Home").Range("T9") = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CellValue2, Crit2_Rng, 12, Crit3_Rng, Year(Date))

ActiveWorkbook.Save

End Sub
