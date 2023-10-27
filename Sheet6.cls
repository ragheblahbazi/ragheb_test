VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()

Call Module1.clear_date_from_last_date_file
Call Module1.rechercheChaine

End Sub

Private Sub CommandButton2_Click()
Call Module1.clear_date_from_last_date_file
End Sub

Private Sub Worksheet_Activate()

'Call Module1.fileLastDate
End Sub


