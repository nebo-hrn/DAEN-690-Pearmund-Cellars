'Function that makes the next block wait until the python is done
Sub ShellAndWait(pathFile As String)
    With CreateObject("WScript.Shell")
        .Run pathFile, 1, True
    End With
End Sub
'Predicting Demand
Sub Button12_Click()

ActiveWorkbook.Save
Application.ScreenUpdating = False
Set Source_workbook = Workbooks.Open("C:\Users\Nebojsa\Desktop\DAEN690\Demand.xlsx")
Workbooks("Demand.xlsx").Worksheets("Demand").Range("A1:Z33").Clear
Source_workbook.Close SaveChanges:=True

Dim args As String
args = "C:/Users/Nebojsa/Desktop/DAEN690/ARIMAbackend.py"

ShellAndWait ("C:\Programming\Python\python.exe" & " " & args)



Application.ScreenUpdating = False
Set Source_workbook = Workbooks.Open("C:\Users\Nebojsa\Desktop\DAEN690\Demand.xlsx")
Workbooks("Capstone_Excel_Format.xlsm").Worksheets("Home").Range("H26:U41").Clear
Workbooks("Demand.xlsx").Worksheets("Demand").Range("A1:N16").Copy
Workbooks("Capstone_Excel_Format.xlsm").Worksheets("Home").Range("H26:U41").PasteSpecial Paste:=xlPasteFormats
Workbooks("Capstone_Excel_Format.xlsm").Worksheets("Home").Range("H26:U41").PasteSpecial Paste:=xlPasteValues
Application.DisplayAlerts = False
Source_workbook.Close SaveChanges:=False
ActiveWorkbook.Save
Application.ScreenUpdating = True
Application.DisplayAlerts = True

Application.Goto Range("G12"), Scroll:=True



End Sub
