VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()

'
Dim table1 As ListObject
Set sh = ThisWorkbook.Sheets("Projet")
Set table1 = sh.ListObjects("Table1")


stop_macro = False

If Import.checked_Cat = False Then
MsgBox "Please select Category ", vbInformation
 
 Exit Sub

End If

If (table1.ListRows.Count > 1) Then
If MsgBox("Would you like to clear data?", vbYesNo) = vbYes Then

    Call Module1.ClearData
    Call Module1.ClearLatestUpdate

End If
End If


Call Import.Importation_SAP

If (stop_macro = False) Then
Call Import.Importation_STOCK(Import.last_Stock_File)
ThisWorkbook.RefreshAll
Call Import.CloseAllFiles
Call Module1.MettreAJourDate
ThisWorkbook.Activate
MsgBox ("Operation Done")

End If

End Sub




Private Sub CommandButton2_Click()

    Call Module2.Clear_IMA_logistica_data
    Call Module2.Open_IMA_Logistica_CSV_File
    Call Import.Convertir_csv_excel
    Call Module2.copyDataStockIMA


End Sub

Private Sub SUPPLY_Click()

End Sub
