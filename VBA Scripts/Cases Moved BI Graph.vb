'Average Cases Moved Graph
Sub Button35_Click()
Application.DisplayAlerts = False
Application.ScreenUpdating = False

Sheets("Graphs").Delete
Sheets.Add(After:=Sheets("Home")).Name = "Graphs"

Dim Cht1 As Chart
Set Cht1 = Sheets("Graphs").Shapes.AddChart(Left:=0, Width:=1200, Top:=0, Height:=500).Chart
With Cht1
.SetSourceData Source:=Sheets("Home").Range("H7:T9")
.ChartType = xlLineMarkers
.HasTitle = True
.ChartTitle.Text = Sheets("Home").Range("F7") & ": " & Sheets("Home").Range("F8") & " - All-Time Average & Current Totals"

.SeriesCollection(1).HasDataLabels = True
.SeriesCollection(2).HasDataLabels = True
.SeriesCollection(1).DataLabels.Position = xlLabelPositionAbove
.SeriesCollection(2).DataLabels.Position = xlLabelPositionAbove
'Y Axis Title
.Axes(xlValue).HasTitle = True
.Axes(xlValue).AxisTitle.Text = "Cases Moved per Month"
        
'X Axis Title
.SetElement msoElementPrimaryCategoryAxisTitleBelowAxis
.Axes(xlCategory).AxisTitle.Text = "Month"
End With

ActiveWorkbook.Save

Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub
