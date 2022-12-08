'Top 5 and Bottom 5 Wines
Sub Button11_Click()
Application.ScreenUpdating = False
Dim CellValue1 As String, CellValue2 As String, table As ListObject, table2 As ListObject

CellValue1 = Sheets("Home").Range("F15").Value
CellValue2 = Sheets("Home").Range("F16").Value

If CellValue1 = "Pearmund Cellars" Then Set table = Sheets("PC Inventory").ListObjects("PCInv"): Set table2 = Sheets("PC Current Inv").ListObjects("PCCIAll")
If CellValue1 = "Effingham Manor" Then Set table = Sheets("EF Inventory").ListObjects("EFInv"): Set table2 = Sheets("EF Current Inv").ListObjects("EFCIAll")
If CellValue1 = "Vint Hill" Then Set table = Sheets("VH Inventory").ListObjects("VHInv"): Set table2 = Sheets("VH Current Inv").ListObjects("VHCIAll")

Sheets("PC Current Inv").Range("F2:F50").Copy

Sheets("WineRankCalc").Range("A1").PasteSpecial Paste:=xlPasteValues, SkipBlanks:=True
Sheets("WineRankCalc").Range("A1:A50").Value = Sheets("WineRankCalc").Range("A1:A50").Value

Dim NoRow As Integer
NoRow = Application.WorksheetFunction.CountA(Worksheets("WineRankCalc").Columns("A"))

Mon = Application.VLookup(CellValue2, Worksheets("Data Entry Options").Range("E2:F13"), 2, False)


Dim SumRng As Range, Crit1_Rng As Range, Crit1 As String, Crit2_Rng As Range, Crit3_Rng As Range, Yearv As Integer

Set SumRng = table.ListColumns(5).DataBodyRange
Set Crit1_Rng = table.ListColumns(3).DataBodyRange
Set Crit2_Rng = table.ListColumns(4).DataBodyRange
Set Crit3_Rng = table.ListColumns(1).DataBodyRange
Yearv = Year(Date)

For Counter = 1 To NoRow
Set refCell = Sheets("WineRankCalc").Range("A" & Counter)
Sheets("WineRankCalc").Cells(Counter, 2).Clear
Sheets("WineRankCalc").Cells(Counter, 2) = Application.WorksheetFunction.SumIfs(SumRng, Crit3_Rng, Yearv, Crit1_Rng, refCell, Crit2_Rng, Mon)

Next Counter

Worksheets("WineRankCalc").Range("A1:B" & NoRow).Sort Key1:=Worksheets("WineRankCalc").Range("B1"), Order1:=xlDescending, Header:=xlNo

Dim n As Integer
n = Application.WorksheetFunction.CountIf(Worksheets("WineRankCalc").Range("B1:B" & NoRow), ">0")
If n = 0 Then
n = 5
End If


'1
Sheets("Home").Range("H16") = Sheets("WineRankCalc").Range("A1")
Sheets("Home").Range("I16") = Sheets("WineRankCalc").Range("B1")
Sheets("Home").Range("K16") = Sheets("WineRankCalc").Range("A" & n)
Sheets("Home").Range("L16") = Sheets("WineRankCalc").Range("B" & n)
'2
Sheets("Home").Range("H17") = Sheets("WineRankCalc").Range("A2")
Sheets("Home").Range("I17") = Sheets("WineRankCalc").Range("B2")
Sheets("Home").Range("K17") = Sheets("WineRankCalc").Range("A" & n - 1)
Sheets("Home").Range("L17") = Sheets("WineRankCalc").Range("B" & n - 1)
'3
Sheets("Home").Range("H18") = Sheets("WineRankCalc").Range("A3")
Sheets("Home").Range("I18") = Sheets("WineRankCalc").Range("B3")
Sheets("Home").Range("K18") = Sheets("WineRankCalc").Range("A" & n - 2)
Sheets("Home").Range("L18") = Sheets("WineRankCalc").Range("B" & n - 2)
'4
Sheets("Home").Range("H19") = Sheets("WineRankCalc").Range("A4")
Sheets("Home").Range("I19") = Sheets("WineRankCalc").Range("B4")
Sheets("Home").Range("K19") = Sheets("WineRankCalc").Range("A" & n - 3)
Sheets("Home").Range("L19") = Sheets("WineRankCalc").Range("B" & n - 3)
'5
Sheets("Home").Range("H20") = Sheets("WineRankCalc").Range("A5")
Sheets("Home").Range("I20") = Sheets("WineRankCalc").Range("B5")
Sheets("Home").Range("K20") = Sheets("WineRankCalc").Range("A" & n - 4)
Sheets("Home").Range("L20") = Sheets("WineRankCalc").Range("B" & n - 4)
ActiveWorkbook.Save
Application.ScreenUpdating = True


End Sub
