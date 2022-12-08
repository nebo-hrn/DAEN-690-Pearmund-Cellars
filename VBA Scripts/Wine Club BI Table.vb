'Wine Club
Sub Button32_Click()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim CellValue1 As String, CellValue2 As String, CellValue3 As String, table1 As ListObject, table2 As ListObject, table3 As ListObject

CellValue1 = Sheets("Home").Range("O15").Value 'Winery
CellValue2 = Sheets("Home").Range("O16").Value 'Year
CellValue3 = Sheets("Home").Range("O17").Value 'Quarter

If CellValue1 = "Pearmund Cellars" Then Set table1 = Sheets("Descriptive Info").ListObjects("PCquarter"): Set table2 = Sheets("PC Inventory").ListObjects("PCInv"): Set table3 = Sheets("Wine Club").ListObjects("PCWineClub")
If CellValue1 = "Effingham Manor" Then Set table1 = Sheets("Descriptive Info").ListObjects("EFquarter"): Set table2 = Sheets("EF Inventory").ListObjects("EFInv"): Set table3 = Sheets("Wine Club").ListObjects("EFWineClub")

If CellValue1 = "Vint Hill" Then
MsgBox "No Data for Vint Hill", vbCritical
Exit Sub
End If

Qmonth = Application.VLookup(CellValue3, table1.DataBodyRange, 2, False)

Dim lo As ListObject
    For Each lo In Sheets("Wine Club").ListObjects
        lo.AutoFilter.ShowAllData
    Next lo
    

Dim Column1 As Range, Column2 As Range

Set Column1 = table3.ListColumns(2).DataBodyRange
Set Column2 = table3.ListColumns(3).DataBodyRange

Column1.AutoFilter Field:=2, Criteria1:=CellValue2
Column2.AutoFilter Field:=3, Criteria1:=Qmonth

wcdate = DateSerial(CInt(CellValue2), CInt(Qmonth), 1)

Counter = 16

Dim SumRng As Range, Crit1_Rng As Range, Crit2_Rng As Range, Crit3_Rng As Range, Crit4_Rng As Range
Set SumRng = table2.ListColumns(5).DataBodyRange 'Cases Moved
Set Crit1_Rng = table2.ListColumns(6).DataBodyRange 'Sale Date
Set Crit2_Rng = table2.ListColumns(2).DataBodyRange 'Vintage
Set Crit3_Rng = table2.ListColumns(3).DataBodyRange 'Wine

Workbooks("Capstone_Excel_Format.xlsm").Worksheets("Home").Range("Q15:Y21").Clear

Dim cl As Range
    For Each cl In table3.ListColumns(4).DataBodyRange.SpecialCells(xlCellTypeVisible)
        Vintage = cl
        wine = table3.ListColumns(5).Range(cl.Row)
        
        Sheets("Home").Range("Q" & Counter) = Vintage
        Sheets("Home").Range("R" & Counter) = wine
        Sheets("Home").Range("Q" & Counter).Font.Bold = True
        Sheets("Home").Range("R" & Counter).Font.Bold = True
        Sheets("Home").Range("Q" & Counter).Interior.ColorIndex = 15
        Sheets("Home").Range("R" & Counter).Interior.ColorIndex = 15
        Sheets("Home").Range("Q" & Counter + 3) = "All"
        Sheets("Home").Range("R" & Counter + 3) = wine
        
        Sheets("Home").Range("S15") = DateAdd("m", -3, wcdate)
        Sheets("Home").Range("T15") = DateAdd("m", -2, wcdate)
        Sheets("Home").Range("U15") = DateAdd("m", -1, wcdate)
        Sheets("Home").Range("V15") = DateAdd("m", 0, wcdate)
        Sheets("Home").Range("W15") = DateAdd("m", 1, wcdate)
        Sheets("Home").Range("X15") = DateAdd("m", 2, wcdate)
        Sheets("Home").Range("Y15") = DateAdd("m", 3, wcdate)
        
        Sheets("Home").Range("S" & Counter) = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CStr(Sheets("Home").Range("S15")), Crit2_Rng, CStr(cl), Crit3_Rng, CStr(wine))
        Sheets("Home").Range("T" & Counter) = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CStr(Sheets("Home").Range("T15")), Crit2_Rng, CStr(cl), Crit3_Rng, CStr(wine))
        Sheets("Home").Range("U" & Counter) = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CStr(Sheets("Home").Range("U15")), Crit2_Rng, CStr(cl), Crit3_Rng, CStr(wine))
        Sheets("Home").Range("V" & Counter) = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CStr(Sheets("Home").Range("V15")), Crit2_Rng, CStr(cl), Crit3_Rng, CStr(wine))
        Sheets("Home").Range("W" & Counter) = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CStr(Sheets("Home").Range("W15")), Crit2_Rng, CStr(cl), Crit3_Rng, CStr(wine))
        Sheets("Home").Range("X" & Counter) = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CStr(Sheets("Home").Range("X15")), Crit2_Rng, CStr(cl), Crit3_Rng, CStr(wine))
        Sheets("Home").Range("Y" & Counter) = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CStr(Sheets("Home").Range("Y15")), Crit2_Rng, CStr(cl), Crit3_Rng, CStr(wine))
        
        Sheets("Home").Range("S" & Counter).Font.Bold = True
        Sheets("Home").Range("T" & Counter).Font.Bold = True
        Sheets("Home").Range("U" & Counter).Font.Bold = True
        Sheets("Home").Range("V" & Counter).Font.Bold = True
        Sheets("Home").Range("W" & Counter).Font.Bold = True
        Sheets("Home").Range("X" & Counter).Font.Bold = True
        Sheets("Home").Range("Y" & Counter).Font.Bold = True
        
        Sheets("Home").Range("S" & Counter).HorizontalAlignment = xlCenter
        Sheets("Home").Range("T" & Counter).HorizontalAlignment = xlCenter
        Sheets("Home").Range("U" & Counter).HorizontalAlignment = xlCenter
        Sheets("Home").Range("V" & Counter).HorizontalAlignment = xlCenter
        Sheets("Home").Range("W" & Counter).HorizontalAlignment = xlCenter
        Sheets("Home").Range("X" & Counter).HorizontalAlignment = xlCenter
        Sheets("Home").Range("Y" & Counter).HorizontalAlignment = xlCenter
        
        Sheets("Home").Range("S" & Counter).Interior.ColorIndex = 15
        Sheets("Home").Range("T" & Counter).Interior.ColorIndex = 15
        Sheets("Home").Range("U" & Counter).Interior.ColorIndex = 15
        Sheets("Home").Range("V" & Counter).Interior.ColorIndex = 15
        Sheets("Home").Range("W" & Counter).Interior.ColorIndex = 15
        Sheets("Home").Range("X" & Counter).Interior.ColorIndex = 15
        Sheets("Home").Range("Y" & Counter).Interior.ColorIndex = 15
        
        Sheets("Home").Range("S" & Counter + 3) = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CStr(Sheets("Home").Range("S15")), Crit3_Rng, CStr(wine))
        Sheets("Home").Range("T" & Counter + 3) = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CStr(Sheets("Home").Range("T15")), Crit3_Rng, CStr(wine))
        Sheets("Home").Range("U" & Counter + 3) = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CStr(Sheets("Home").Range("U15")), Crit3_Rng, CStr(wine))
        Sheets("Home").Range("V" & Counter + 3) = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CStr(Sheets("Home").Range("V15")), Crit3_Rng, CStr(wine))
        Sheets("Home").Range("W" & Counter + 3) = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CStr(Sheets("Home").Range("W15")), Crit3_Rng, CStr(wine))
        Sheets("Home").Range("X" & Counter + 3) = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CStr(Sheets("Home").Range("X15")), Crit3_Rng, CStr(wine))
        Sheets("Home").Range("Y" & Counter + 3) = Application.WorksheetFunction.SumIfs(SumRng, Crit1_Rng, CStr(Sheets("Home").Range("Y15")), Crit3_Rng, CStr(wine))
        
        Sheets("Home").Range("S" & Counter + 3).HorizontalAlignment = xlCenter
        Sheets("Home").Range("T" & Counter + 3).HorizontalAlignment = xlCenter
        Sheets("Home").Range("U" & Counter + 3).HorizontalAlignment = xlCenter
        Sheets("Home").Range("V" & Counter + 3).HorizontalAlignment = xlCenter
        Sheets("Home").Range("W" & Counter + 3).HorizontalAlignment = xlCenter
        Sheets("Home").Range("X" & Counter + 3).HorizontalAlignment = xlCenter
        Sheets("Home").Range("Y" & Counter + 3).HorizontalAlignment = xlCenter
        
        Counter = Counter + 1
        
    Next cl
    
Sheets("Home").Range("Q15:Y" & Counter + 2).BorderAround LineStyle:=xlContinuous, Weight:=xlThick
Sheets("Home").Range("Q15:Y15").BorderAround LineStyle:=xlContinuous, Weight:=xlThick
Sheets("Home").Range("V15:V" & Counter + 2).Interior.ColorIndex = 6


ActiveWorkbook.Save


Application.ScreenUpdating = True
Application.DisplayAlerts = True


End Sub
