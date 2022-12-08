Private Sub CommandButton1_Click()
Dim sht As Worksheet, sht1 As Worksheet, lastrow As Long, table As ListObject

If optPC.Value = True Then Set table = Sheets("PC Current Inv").ListObjects("PCCInv")
If optEF.Value = True Then Set table = Sheets("PC Inventory").ListObjects("EFInv")
If optVH.Value = True Then Set table = Sheets("PC Inventory").ListObjects("VHInv")
'Based on what winery is selected data will go to the appropriate table

If optPC.Value = False And optEF.Value = False And optVH.Value = False Then
MsgBox "Please select a Winery for this data", vbCritical
Exit Sub
End If
'Making sure a winery is actually selected

If VBA.IsNumeric(txtFiscalYear.Value) = False Then
MsgBox "Only numeric values are accepted as a Fiscal Year", vbCritical
Exit Sub
End If

If VBA.IsNumeric(txtCasesMoved.Value) = False Then
MsgBox "Only numeric values are accepted in the Cases Moved", vbCritical
Exit Sub
End If
'Making sure values entered are numeric for text boxes

Dim addedRow As ListRow
Set addedRow = table.ListRows.Add()

With addedRow
    .Range(1) = txtFiscalYear.Value
    .Range(2) = cmbVintage.Value
    .Range(3) = cmbWine.Value
    .Range(4) = txtCasesMoved.Value
End With

End Sub
Private Sub CommandButton2_Click()
With Me
.txtFiscalYear = ""
.cmbVintage = ""
.cmbWine = ""
.txtCasesMoved = ""
.optPC = False
.optEF = False
.optVH = False
End With
ActiveWorkbook.Save
End Sub

Private Sub CommandButton3_Click()
Unload Me
ActiveWorkbook.Save
End Sub
Private Sub UserForm_Activate()

cmbVintage.List = Sheets("Data Entry Options").ListObjects("VintageType").ListColumns(1).DataBodyRange.Value

cmbWine.List = Sheets("Data Entry Options").ListObjects("WineType").ListColumns(1).DataBodyRange.Value

End Sub
