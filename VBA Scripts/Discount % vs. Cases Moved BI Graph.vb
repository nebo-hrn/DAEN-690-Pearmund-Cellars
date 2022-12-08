'Cases Moved vs Discount Rate
Sub Button38_Click()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim CellValue1 As String, CellValue2 As String, table1 As ListObject, table2 As ListObject

CellValue1 = Sheets("Home").Range("P46").Value 'Winery
CellValue2 = Sheets("Home").Range("P47").Value 'Wine

If CellValue1 = "Pearmund Cellars" Then Set table1 = Sheets("PC Inventory").ListObjects("PCInv"): Set table2 = Sheets("Sales Data").ListObjects("PCSalesData")
If CellValue1 = "Effingham Manor" Then Set table1 = Sheets("EF Inventory").ListObjects("EFInv"): Set table2 = Sheets("Sales Data").ListObjects("EFSalesData")

If CellValue1 = "Vint Hill" Then
MsgBox "No Data for Vint Hill", vbCritical
Exit Sub
End If

table2.DataBodyRange.AutoFilter Field:=2, Criteria1:=CellValue2
start_year = Application.WorksheetFunction.Min(table2.ListColumns(1).DataBodyRange)

Sheets("Graphs").Delete
Sheets.Add(After:=Sheets("Home")).Name = "Graphs"

Sheets("Graphs").Range("A1") = "Year"
Sheets("Graphs").Range("B1") = "Cases Moved"
Sheets("Graphs").Range("C1") = "Discount Percentage"

Dim SumRng1 As Range, SumRng2 As Range, Crit1_Rng As Range, Crit2_Rng As Range, Crit3_Rng As Range, Crit4_Rng As Range
Set SumRng1 = table1.ListColumns(5).DataBodyRange 'Cases Moved
Set Crit1_Rng = table1.ListColumns(1).DataBodyRange 'Fiscal Year
Set Crit2_Rng = table1.ListColumns(3).DataBodyRange 'Wine

Set SumRng2 = table2.ListColumns(5).DataBodyRange 'Discount Percentage
Set Crit3_Rng = table2.ListColumns(1).DataBodyRange 'Fiscal Year
Set Crit4_Rng = table2.ListColumns(2).DataBodyRange 'Wine

Num = 2
For Counter = start_year To Year(Date) - 1
Sheets("Graphs").Range("A" & Num) = Counter
Sheets("Graphs").Range("B" & Num) = Application.WorksheetFunction.SumIfs(SumRng1, Crit1_Rng, Counter, Crit2_Rng, CellValue2)
Sheets("Graphs").Range("C" & Num) = FormatPercent(Application.WorksheetFunction.SumIfs(SumRng2, Crit3_Rng, Counter, Crit4_Rng, CellValue2))
Num = Num + 1
Next Counter

Dim Cht1 As Chart
Set Cht1 = Sheets("Graphs").Shapes.AddChart(Left:=0, Width:=1200, Top:=0, Height:=500).Chart
With Cht1
.SeriesCollection(1).Delete
'.SetSourceData Sheets("Graphs").Range("A1:C" & Num - 1)
.ChartType = xlLineMarkers
.SeriesCollection(1).XValues = Sheets("Graphs").Range("A2:A" & Num - 1)
.SeriesCollection(1).Values = Sheets("Graphs").Range("B2:B" & Num - 1)
.SeriesCollection(2).Values = Sheets("Graphs").Range("C2:C" & Num - 1)
.SeriesCollection(1).AxisGroup = 1
.SeriesCollection(2).AxisGroup = 2

.SeriesCollection(1).HasDataLabels = True
.SeriesCollection(2).HasDataLabels = True
.SeriesCollection(1).DataLabels.Position = xlLabelPositionAbove
.SeriesCollection(2).DataLabels.Position = xlLabelPositionBelow

.HasTitle = True
.ChartTitle.Text = "Discount % vs. Cases Moved: " & CellValue1 & " - " & CellValue2

.HasAxis(xlValue, xlSecondary) = True ' add the secondary axis
.Axes(xlValue, xlPrimary).HasTitle = True
.Axes(xlValue, xlSecondary).HasTitle = True
.Axes(xlValue, xlPrimary).AxisTitle.Text = "Cases Moved"
.Axes(xlValue, xlSecondary).AxisTitle.Text = "Discount Percentage"
.Axes(xlValue, xlPrimary).MinimumScale = Application.WorksheetFunction.Min(Sheets("Graphs").Range("B2:B" & Num)) - 50
.Axes(xlValue, xlSecondary).MinimumScale = Application.WorksheetFunction.Min(Sheets("Graphs").Range("C2:C" & Num)) - 0.01
.SetElement msoElementPrimaryCategoryAxisTitleBelowAxis
.Axes(xlCategory).AxisTitle.Text = "Year"
End With

ActiveWorkbook.Save

Application.ScreenUpdating = False
Application.DisplayAlerts = False
End Sub
