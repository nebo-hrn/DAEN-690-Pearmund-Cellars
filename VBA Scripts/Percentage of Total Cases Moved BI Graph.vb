'Month to Month Graph
Sub Button37_Click()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim CellValue1 As String, CellValue2 As String, CellValue3 As String, CellValue4 As String, NumYears As Long, table As ListObject

CellValue1 = Sheets("Home").Range("M46").Value 'Winery
CellValue2 = Sheets("Home").Range("M47").Value 'Wine 1
CellValue3 = Sheets("Home").Range("M48").Value 'Wine 2
CellValue4 = Sheets("Home").Range("M49").Value 'Wine 3

If CellValue1 = "Pearmund Cellars" Then Set table = Sheets("PC Inventory").ListObjects("PCInv")
If CellValue1 = "Effingham Manor" Then Set table = Sheets("EF Inventory").ListObjects("EFInv")
If CellValue1 = "Vint Hill" Then Set table = Sheets("VH Inventory").ListObjects("VHInv")

Dim SumRng As Range, Crit1_Rng As Range, Crit2_Rng As Range, Crit3_Rng As Range, Crit4_Rng As Range
Set SumRng = table.ListColumns(5).DataBodyRange 'Cases Moved
Set Crit1_Rng = table.ListColumns(6).DataBodyRange 'Sale Date
Set Crit2_Rng = table.ListColumns(3).DataBodyRange 'Wine

Sheets("Graphs").Delete
Sheets.Add(After:=Sheets("Home")).Name = "Graphs"


F_date = DateSerial(2016, 1, 1)
NumMonth = DateDiff("m", F_date, Date)

If Len(CellValue2) > 0 And Len(CellValue3) = 0 And Len(CellValue4) = 0 Then
    Sheets("Graphs").Range("A1") = "Date"
    Sheets("Graphs").Range("B1") = CellValue2
    
    For Counter = 0 To NumMonth - 1
        MonthYear = DateAdd("m", Counter, F_date)
        Sheets("Graphs").Range("A" & Counter + 2) = MonthYear
        totalcases = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, MonthYear)
        winecases1 = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, MonthYear, Crit2_Rng, CellValue2)
        If totalcases = 0 Then
            Sheets("Graphs").Range("B" & Counter + 2) = FormatPercent(0)
        Else
            Sheets("Graphs").Range("B" & Counter + 2) = FormatPercent(winecases1 / totalcases)
        End If
    Next Counter
    wide = 1
End If

If Len(CellValue2) > 0 And Len(CellValue3) > 0 And Len(CellValue4) = 0 Then
    Sheets("Graphs").Range("A1") = "Year"
    Sheets("Graphs").Range("B1") = CellValue2
    Sheets("Graphs").Range("C1") = CellValue3

    For Counter = 0 To NumMonth - 1
        MonthYear = DateAdd("m", Counter, F_date)
        Sheets("Graphs").Range("A" & Counter + 2) = MonthYear
        totalcases = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, MonthYear)
        winecases1 = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, MonthYear, Crit2_Rng, CellValue2)
        winecases2 = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, MonthYear, Crit2_Rng, CellValue3)
        If totalcases = 0 Then
            Sheets("Graphs").Range("B" & Counter + 2) = FormatPercent(0)
            Sheets("Graphs").Range("C" & Counter + 2) = FormatPercent(0)
        Else
            Sheets("Graphs").Range("B" & Counter + 2) = FormatPercent(winecases1 / totalcases)
            Sheets("Graphs").Range("C" & Counter + 2) = FormatPercent(winecases2 / totalcases)
        End If
    Next Counter
    wide = 2
End If

If Len(CellValue2) > 0 And Len(CellValue3) > 0 And Len(CellValue4) > 0 Then
    Sheets("Graphs").Range("A1") = "Year"
    Sheets("Graphs").Range("B1") = CellValue2
    Sheets("Graphs").Range("C1") = CellValue3
    Sheets("Graphs").Range("D1") = CellValue4
    
    For Counter = 0 To NumMonth - 1
        MonthYear = DateAdd("m", Counter, F_date)
        Sheets("Graphs").Range("A" & Counter + 2) = MonthYear
        totalcases = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, MonthYear)
        winecases1 = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, MonthYear, Crit2_Rng, CellValue2)
        winecases2 = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, MonthYear, Crit2_Rng, CellValue3)
        winecases3 = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, MonthYear, Crit2_Rng, CellValue4)
        If totalcases = 0 Then
            Sheets("Graphs").Range("B" & Counter + 2) = FormatPercent(0)
            Sheets("Graphs").Range("C" & Counter + 2) = FormatPercent(0)
            Sheets("Graphs").Range("D" & Counter + 2) = FormatPercent(0)
        Else
            Sheets("Graphs").Range("B" & Counter + 2) = FormatPercent(winecases1 / totalcases)
            Sheets("Graphs").Range("C" & Counter + 2) = FormatPercent(winecases2 / totalcases)
            Sheets("Graphs").Range("D" & Counter + 2) = FormatPercent(winecases3 / totalcases)
        End If
    Next Counter
    wide = 3
End If

Dim Cht1 As Chart
Set Cht1 = Sheets("Graphs").Shapes.AddChart(Left:=0, Width:=1200, Top:=0, Height:=500).Chart
With Cht1
If wide = 1 Then .SetSourceData Source:=Sheets("Graphs").Range("B1:B" & Counter + 1)
If wide = 2 Then .SetSourceData Source:=Sheets("Graphs").Range("B1:C" & Counter + 1)
If wide = 3 Then .SetSourceData Source:=Sheets("Graphs").Range("B1:D" & Counter + 1)
.ChartType = xlColumnStacked
.HasTitle = True
.ChartTitle.Text = "Month to Month Percentage of Total Cases Moved - " & CellValue1
'Y Axis Title
.Axes(xlValue).HasTitle = True
.Axes(xlValue).AxisTitle.Text = "Percentage of Total Cases Moved"
        
'X Axis Title
.SetElement msoElementPrimaryCategoryAxisTitleBelowAxis
.Axes(xlCategory).AxisTitle.Text = "Month & Year"
.SeriesCollection(1).XValues = Sheets("Graphs").Range("A2:A" & Counter + 1)
End With

ActiveWorkbook.Save

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub
