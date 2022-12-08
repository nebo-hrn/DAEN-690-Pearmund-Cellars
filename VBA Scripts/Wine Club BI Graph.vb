'Wine Club Graph
Sub Button34_Click()
Application.DisplayAlerts = False
Application.ScreenUpdating = False

Sheets("Graphs").Delete
Sheets.Add(After:=Sheets("Home")).Name = "Graphs"


Dim Cht1 As Chart
Set Cht1 = Sheets("Graphs").Shapes.AddChart(Left:=0, Width:=1200, Top:=0, Height:=500).Chart
With Cht1
.SetSourceData Source:=Sheets("Home").Range("Q15:Y18")
.ChartType = xlColumnClustered
.HasTitle = True
.ChartTitle.Text = "Cases Moved Before & After Wine Club"
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
