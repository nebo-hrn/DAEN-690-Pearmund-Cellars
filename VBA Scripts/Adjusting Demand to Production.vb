'Adjusting Demand
Sub Button16_Click()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim CellValue1 As String, table1 As ListObject, CellValue2 As Integer, CellValue3 As Integer

CellValue1 = Sheets("Home").Range("F25").Value
CellValue2 = Sheets("Home").Range("B29").Value
CellValue3 = Sheets("Home").Range("B30").Value
MidVal = (CellValue2 + CellValue3) / 2

If CellValue1 = "Pearmund Cellars" Then Set table1 = Sheets("PC Current Inv").ListObjects("PCCIAll")
If CellValue1 = "Effingham Manor" Then Set table1 = Sheets("EF Current Inv").ListObjects("EFCIAll")
If CellValue1 = "Vint Hill" Then Set table1 = Sheets("VH Current Inv").ListObjects("VHCIAll")

Dim NoRow As Integer
NoRow = Application.WorksheetFunction.CountA(Worksheets("Home").Range("H27:H41"))
NoRow2 = 27 + NoRow

Workbooks("Capstone_Excel_Format.xlsm").Worksheets("Home").Range("V27:AA41").Clear 'Clear adjustments

Dim wine_rng As String, predd_rng As String, prod_rng As String 'Set the columns to be written based on winery
If CellValue1 = "Pearmund Cellars" Then wine_rng = "A": predd_rng = "B": prod_rng = "C"
If CellValue1 = "Effingham Manor" Then wine_rng = "F": predd_rng = "G": prod_rng = "H"
If CellValue1 = "Vint Hill" Then wine_rng = "K": predd_rng = "L": prod_rng = "M"

Workbooks("Capstone_Excel_Format.xlsm").Worksheets("Previous Predictions").Range(wine_rng & "1:" & prod_rng & "20").Clear 'Clear Previous instance

Sheets("Previous Predictions").Range(wine_rng & 1).Font.Bold = True 'Bold the Winery
Sheets("Previous Predictions").Range(wine_rng & 1) = CellValue1 'Winery

Sheets("Previous Predictions").Range(predd_rng & 1) = "Last Predicted: " 'Last predicted string

Sheets("Previous Predictions").Range(prod_rng & 1).Font.Bold = True 'Bold the Date & Time
Sheets("Previous Predictions").Range(prod_rng & 1) = Format(Now, "mmmm d yyyy h:mm AM/PM") 'Date & Time Predicted

'Text formating for Previous Prediction Tables
Sheets("Previous Predictions").Range(wine_rng & 2).Font.Bold = True
Sheets("Previous Predictions").Range(wine_rng & 2) = "Wine"
Sheets("Previous Predictions").Range(predd_rng & 2).Font.Bold = True
Sheets("Previous Predictions").Range(predd_rng & 2) = "Predicted Demand"
Sheets("Previous Predictions").Range(prod_rng & 2).Font.Bold = True
Sheets("Previous Predictions").Range(prod_rng & 2) = "Adjusted Production"


'Start loop through all previously predicted wines
For Counter = 27 To (26 + NoRow)
refCell = Worksheets("Home").Range("H" & Counter)
CurInv = Application.VLookup(refCell, table1.DataBodyRange, 4, False)

Sheets("Previous Predictions").Range(predd_rng & Counter - 24).NumberFormat = "0" 'Formating demand to whole cases
Sheets("Previous Predictions").Range(wine_rng & Counter - 24) = refCell 'Write Wine
Sheets("Previous Predictions").Range(predd_rng & Counter - 24) = Sheets("Home").Range("U" & Counter) 'Write Predicted Demand
Sheets("Previous Predictions").Range(prod_rng & Counter - 24).NumberFormat = "0" 'Formating production to whole cases

If VarType(CurInv) = 10 Then
Sheets("Home").Range("V" & Counter) = 0
Else
Sheets("Home").Range("V" & Counter) = Application.VLookup(refCell, table1.DataBodyRange, 4, False)
End If

Sheets("Home").Range("W" & Counter) = Sheets("Home").Range("U" & Counter) + Sheets("Home").Range("V" & Counter)
Sheets("Home").Range("X" & Counter) = Sheets("Home").Range("W" & Counter) / (Sheets("Home").Range("U" & Counter) / 12)

Sheets("Home").Range("Y" & Counter).NumberFormat = "0"
Sheets("Home").Range("Y" & Counter).Interior.ColorIndex = 6
Sheets("Home").Range("Y" & Counter).Borders(xlEdgeLeft).LineStyle = xlContinuous
Sheets("Home").Range("Y" & Counter).Borders(xlEdgeLeft).Weight = xlThick

If Sheets("Home").Range("X" & Counter) < CellValue3 Then
Sheets("Home").Range("Y" & Counter) = (MidVal - Sheets("Home").Range("X" & Counter)) * (Sheets("Home").Range("U" & Counter) / 12) + Sheets("Home").Range("U" & Counter)
Sheets("Previous Predictions").Range(prod_rng & Counter - 24) = Sheets("Home").Range("Y" & Counter)
End If
If Sheets("Home").Range("X" & Counter) > CellValue2 Then
AdjVal = Sheets("Home").Range("U" & Counter) - (Sheets("Home").Range("X" & Counter) - MidVal) * (Sheets("Home").Range("U" & Counter) / 12)
If AdjVal > 0 Then
Sheets("Home").Range("Y" & Counter) = AdjVal
Sheets("Previous Predictions").Range(prod_rng & Counter - 24) = Sheets("Home").Range("Y" & Counter) 'Write adjusted value
Else
Sheets("Home").Range("Y" & Counter) = 0
Sheets("Previous Predictions").Range(prod_rng & Counter - 24) = Sheets("Home").Range("Y" & Counter) 'Write ajusted value
End If
End If
If Sheets("Home").Range("X" & Counter) > CellValue3 And Sheets("Home").Range("X" & Counter) < CellValue2 Then
Sheets("Home").Range("Y" & Counter) = Sheets("Home").Range("U" & Counter)
Sheets("Previous Predictions").Range(prod_rng & Counter - 24) = Sheets("Home").Range("Y" & Counter)
End If

Sheets("Home").Range("Z" & Counter) = (Sheets("Home").Range("V" & Counter) + Sheets("Home").Range("Y" & Counter)) / (Sheets("Home").Range("U" & Counter) / 12)
Sheets("Home").Range("AA" & Counter).Font.Bold = True
Sheets("Home").Range("AA" & Counter) = refCell
Next Counter

Sheets("Home").Range("V27:Z" & (26 + NoRow)).Borders(xlEdgeBottom).LineStyle = xlContinuous
Sheets("Home").Range("V27:Z" & (26 + NoRow)).Borders(xlEdgeBottom).Weight = xlThick
Sheets("Home").Range("V27:Z" & (26 + NoRow)).Borders(xlEdgeRight).LineStyle = xlContinuous
Sheets("Home").Range("V27:Z" & (26 + NoRow)).Borders(xlEdgeRight).Weight = xlThick


Worksheets("Previous Predictions").Range(wine_rng & "1:" & prod_rng & Counter - 25).Columns.AutoFit
Sheets("Previous Predictions").Range(wine_rng & "1:" & prod_rng & Counter - 25).BorderAround LineStyle:=xlContinuous, Weight:=xlThick
Sheets("Previous Predictions").Range(wine_rng & "1:" & prod_rng & "2").BorderAround LineStyle:=xlContinuous, Weight:=xlThick

ActiveWorkbook.Save

Application.ScreenUpdating = True
Application.DisplayAlerts = True

Application.Goto Range("N12"), Scroll:=True


End Sub
