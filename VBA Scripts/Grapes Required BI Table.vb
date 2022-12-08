'Grapes Required
Sub Button36_Click()
Application.DisplayAlerts = False
Application.ScreenUpdating = False

Sheets("Home").Range("F44") = Sheets("Home").Range("F25")

Dim table As ListObject

If Sheets("Home").Range("F44") = "Pearmund Cellars" Then Set table = Sheets("Descriptive Info").ListObjects("PCWineRecipes")
If Sheets("Home").Range("F44") = "Effingham Manor" Then Set table = Sheets("Descriptive Info").ListObjects("EFWineRecipes")
If Sheets("Home").Range("F44") = "Vint Hill" Then Set table = Sheets("Descriptive Info").ListObjects("VHWineRecipes")

Sheets("Descriptive Info").ListObjects("GrapeType").ListColumns(1).DataBodyRange.Copy
Sheets("Home").Range("H45").PasteSpecial Paste:=xlPasteValues

Counter = 45
Dim cl As Range
    For Each cl In Sheets("Descriptive Info").ListObjects("GrapeType").ListColumns(1).DataBodyRange
    Dim cl2 As Range
    Sum = 0
    Length = Application.WorksheetFunction.CountA(Worksheets("Home").Range("H27:H41"))
        For Each cl2 In Worksheets("Home").Range("H27:H" & 26 + Length)
            winecases = Application.WorksheetFunction.Index(Sheets("Home").Range("Y27:Y" & (26 + Length)), Application.Match(cl2, Sheets("Home").Range("AA27:AA" & (26 + Length)), 0))
            grapeper = Application.WorksheetFunction.Lookup(cl2, table.ListColumns(1).DataBodyRange, table.ListColumns(CStr(cl)).DataBodyRange)
            Sum = Sum + (winecases * grapeper)
        Next cl2
    Sheets("Home").Range("I" & Counter).NumberFormat = "0"
    Sheets("Home").Range("I" & Counter) = Sum
    Counter = Counter + 1
    
    Next cl
   
ActiveWorkbook.Save
  
Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub
