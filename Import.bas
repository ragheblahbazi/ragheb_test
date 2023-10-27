Attribute VB_Name = "Import"

'*****************************************************************************
'Date de creation: 19.09.2023
'Développé par: Ragheb LAHBAZI et Faustin LUANGA MUKELA
'dernier date de modification: 26.10.2023

'********************************************************************************


Dim derniereLigne As Long
Dim ListopenedFile As Dictionary
Dim col_msg As Collection
Public stop_macro As Boolean
Dim github As Integer


Sub Importation_SAP()

' procédure permerttant d'importer les différents fichiers(Supply,Production Delivery)

    Dim path_supplier_file As Variant
    Set col_msg = New Collection
    
    If ThisWorkbook.Sheets("Projet").OLEObjects("SUPPLY").Object.value = True And stop_macro = False Then
    
        col_msg.Add "Choose Supply Excel File"
        Call chooseFile(2, "Choose Supply Excel File")
    
    End If
    
    If ThisWorkbook.Sheets("Projet").OLEObjects("PRODUCTION").Object.value = True And stop_macro = False Then
        
        col_msg.Add "Choose Production Excel File"
        Call chooseFile(3, "Choose Production Excel File")
    
    End If
    
    If ThisWorkbook.Sheets("Projet").OLEObjects("DELIVERY").Object.value = True And stop_macro = False Then
    
        col_msg.Add "Choose Delivery Excel File"
        Call chooseFile(4, "Choose Delivery Excel File")
    
    End If
    
    
    ThisWorkbook.Sheets("Projet").Range("E1").EntireColumn.AutoFit


End Sub


Sub chooseFile(cat_num, msg)


' procédure permettant de selectionner les fichiers(Supply, Production, Delivery) dans le browser

    If Folder_check(ThisWorkbook.Sheets("Pilotage").Range("C" & cat_num)) = False Then
    
        ChDrive ("C:/")
        ChDir ("C:")
    
    Else
        ChDir ThisWorkbook.Sheets("Pilotage").Range("C" & cat_num)
    
    End If
    
    path_supplier_file = Application.GetOpenFilename(FileFilter:="Fichiers Excel (*.xls*),*xls*", Title:=msg, MultiSelect:=True) ' variable de type chaine de caractere qui stock le lien complet de fichier supplier
    
    If TypeName(path_supplier_file) = "Variant()" Then
    
        
        Select Case cat_num
        
        Case 2:
            Call ImportData_Supply(path_supplier_file)
        
        Case 3:
            Call ImportData_Production(path_supplier_file)
        
        Case 4:
            Call ImportData_Delivery(path_supplier_file)
        
        End Select
        
        Else
        
            If path_supplier_file = Faux Then
                stop_macro = True
                Exit Sub
        
        End If
    
    End If

End Sub
Sub ImportData_Supply(path_supplier_file)


'Procedure permettant d'importer les données Supply à notre Projet

    
    Dim black_list As Variant
    Dim dernierLigne_sap As Long
    Dim rowSap As Long
    Dim compteurJ As Long
    Set ListopenedFile = New Dictionary
    Set black_list = supprimer_projet()
    
    
    For k = LBound(path_supplier_file) To UBound(path_supplier_file)
    
        Workbooks.Open (path_supplier_file(k))  ' ouvrir le fichier
        ListopenedFile.Add Workbooks(Workbooks.Count).Name, Workbooks.Count
        
        'chercher la derniere ligne de la colonne A feuille Projet
        derniereLigne = ThisWorkbook.Sheets("Projet").Range("A" & Rows.Count).End(xlUp).row + 1  ' derniere  ligne
        'determiner  la derniere ligne de DE EXTRACT SAP FILE.....
        dernierLigne_sap = Workbooks(Workbooks.Count).Sheets(1).Range("J" & Rows.Count).End(xlUp).row
        compteurA = derniereLigne
        For compteurJ = 2 To dernierLigne_sap
        
            If (Workbooks(Workbooks.Count).Sheets(1).Range("N" & compteurJ) <> "") Then '  Condition pour verifier que order purshase différent de ""
            
            
                project_name = Nom_projet(compteurJ, Workbooks(Workbooks.Count).Sheets(1))
                
                If (black_list.Exists(UCase(project_name)) = False) Then
                    
                    ThisWorkbook.Sheets("Projet").Range("A" & compteurA).value = project_name
                    ThisWorkbook.Sheets("Projet").Range("B" & compteurA).value = Workbooks(Workbooks.Count).Sheets(1).Range("J" & compteurJ)
                    
                    ThisWorkbook.Sheets("Projet").Range("C" & compteurA).value = Workbooks(Workbooks.Count).Sheets(1).Range("K" & compteurJ)
                    ThisWorkbook.Sheets("Projet").Range("E" & compteurA).value = Workbooks(Workbooks.Count).Sheets(1).Range("A" & compteurJ)
                    ThisWorkbook.Sheets("Projet").Range("F" & compteurA).value = Workbooks(Workbooks.Count).Sheets(1).Range("I" & compteurJ)
                    ThisWorkbook.Sheets("Projet").Range("G" & compteurA).value = "IEMA"
                    ThisWorkbook.Sheets("Projet").Range("H" & compteurA).value = Workbooks(Workbooks.Count).Sheets(1).Range("M" & compteurJ)
                    ThisWorkbook.Sheets("Projet").Range("N" & compteurA).value = "SUPPLY"
                    ThisWorkbook.Sheets("Projet").Range("J" & compteurA).value = Workbooks(Workbooks.Count).Sheets(1).Range("L" & compteurJ)
                    
                    
                    compteurA = compteurA + 1
                End If
                
            
            End If
        Next compteurJ
    
    Next k

End Sub

Sub ImportData_Production(path_supplier_file)

'Procedure permettant d'importer les données Production à notre Projet
    
    Dim k As Long
    Dim supplybk As Workbook
    Dim sh As Worksheet
    Dim ligne As Long
    Dim lastLine As Long
    
    If (ListopenedFile Is Nothing) Then
    
        Set ListopenedFile = New Dictionary
    
    End If
    
    lastLine = ThisWorkbook.Sheets("Projet").Range("A" & Rows.Count).End(xlUp).row + 1
    
    For k = LBound(path_supplier_file) To UBound(path_supplier_file)
    
        Workbooks.Open (path_supplier_file(k))  ' ouvrir le fichier
        ListopenedFile.Add Workbooks(Workbooks.Count).Name, Workbooks.Count
        Set supplybk = Workbooks(Workbooks.Count)
        Set sh = supplybk.Sheets(1)
        ligne = 10
        
        While (sh.Range("B" & ligne).value <> "")
        
            ThisWorkbook.Sheets("Projet").Range("A" & lastLine) = sh.Range("C3")
            ThisWorkbook.Sheets("Projet").Range("B" & lastLine) = sh.Range("B" & ligne)
            ThisWorkbook.Sheets("Projet").Range("C" & lastLine) = sh.Range("D" & ligne)
            ThisWorkbook.Sheets("Projet").Range("E" & lastLine) = sh.Range("C4")
            ThisWorkbook.Sheets("Projet").Range("G" & lastLine) = "IEMA"
            ThisWorkbook.Sheets("Projet").Range("J" & lastLine) = sh.Range("E" & ligne)
            ThisWorkbook.Sheets("Projet").Range("F" & lastLine) = sh.Range("C" & ligne)
            ThisWorkbook.Sheets("Projet").Range("N" & lastLine) = "PRODUCTION"
            ligne = ligne + 1
            lastLine = lastLine + 1
            
        Wend
    
    Next k

End Sub


Sub ImportData_Delivery(path_supplier_file)

'Procedure permettant d'importer les données Delivery à notre Projet

    Dim k As Long
    Dim supplybk As Workbook
    Dim sh As Worksheet
    Dim ligne As Long
    Dim lastLine As Long
    Dim lastDelivery As Long
    Dim project_name As String
    Dim rng As Range
    lastLine = ThisWorkbook.Sheets("Projet").Range("A" & Rows.Count).End(xlUp).row + 1
    
    If (ListopenedFile Is Nothing) Then
    
        Set ListopenedFile = New Dictionary
    
    End If
    
    For k = LBound(path_supplier_file) To UBound(path_supplier_file)
    
        Workbooks.Open (path_supplier_file(k))  ' ouvrir le fichier
        ListopenedFile.Add Workbooks(Workbooks.Count).Name, Workbooks.Count
        
        Set deliverybk = Workbooks(Workbooks.Count)
        Set sh = deliverybk.Sheets(1)
        
        lastDelivery = sh.Range("A" & Rows.Count).End(xlUp).row
        
        For ligne = 2 To lastDelivery
        
            
            project_name = Import.Nom_projet(ligne, sh)
            ThisWorkbook.Sheets("Projet").Range("A" & lastLine) = project_name
            ThisWorkbook.Sheets("Projet").Range("B" & lastLine) = sh.Range("G" & ligne)
            ThisWorkbook.Sheets("Projet").Range("C" & lastLine) = sh.Range("H" & ligne)
            ThisWorkbook.Sheets("Projet").Range("E" & lastLine) = sh.Range("A" & ligne)
            ThisWorkbook.Sheets("Projet").Range("G" & lastLine) = "IEMA"
            ThisWorkbook.Sheets("Projet").Range("J" & lastLine) = sh.Range("I" & ligne)
            ThisWorkbook.Sheets("Projet").Range("F" & lastLine) = sh.Range("F" & ligne)
            ThisWorkbook.Sheets("Projet").Range("H" & lastLine) = sh.Range("N" & ligne)
            ThisWorkbook.Sheets("Projet").Range("N" & lastLine) = "DELIVERY"
            ' calcul criticité en fonction couleur  cas de delivery
            Set rng = sh.Range("G" & ligne)
            ThisWorkbook.Sheets("Projet").Range("M" & lastLine) = calcul_criticité_delivery(ligne, rng)
            
            MiseEnCouleur (lastLine)
            
            lastLine = lastLine + 1
        
        Next ligne
    
    Next k



End Sub
Function last_Stock_File()

'Fonction permettant de parcourir le fichier CSV


    Dim fs As Object
    Dim oFolder As Object
    Dim lastdate As String
    Dim dateNumber As Long
    Dim dateString As String
    Dim file_string As String
    Dim max_Date As Long
    
    
    max_Date = 0
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set oFolder = fs.GetFolder(ThisWorkbook.Sheets("Pilotage").Range("C5").value)
    
    For Each oFile In oFolder.Files
        
        If (oFile.Name Like "stock_med*" And InStr(oFile.Name, ".csv") <> 0) Then
        
            dateString = Mid(oFile.Name, Len("stock_med") + 1, 8)
            dateNumber = CLng(dateString)
            If (dateNumber > max_Date) Then
                max_Date = dateNumber
                file_string = oFile.Name
            End If
        End If
    Next oFile
    
    last_Stock_File = ThisWorkbook.Sheets("Pilotage").Range("C5") & "\" & file_string

End Function
Sub Importation_STOCK(path_stock_complet)

'Procedure permettant d'importer les données du Stock( Fichier CSV)
   
    
    If (ListopenedFile Is Nothing) Then
    
        Set ListopenedFile = New Dictionary
    
    End If
    
    Workbooks.Open Filename:=path_stock_complet
    ListopenedFile.Add Workbooks(Workbooks.Count).Name, Workbooks.Count
    
    Call Import.Convertir_csv_excel
    Call Import.recherche_Code_SAP
    Call Import.calcul_criticity

End Sub


Sub Convertir_csv_excel()
'Procedure permettant de convertir le fichier CSV en Excel
    Dim WK As Workbook
    Application.DisplayAlerts = False
    
    Set WK = ActiveWorkbook
        Columns("A:A").Select
        Selection.TextToColumns destination:=Range("A1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
            Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1 _
            ), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array _
            (20, 1), Array(21, 1), Array(22, 1), Array(23, 1), Array(24, 1), Array(25, 1), Array(26, 1), _
            Array(27, 1), Array(28, 1), Array(29, 1), Array(30, 1), Array(31, 1), Array(32, 1), Array( _
            33, 1), Array(34, 1)), TrailingMinusNumbers:=True
    
    Application.DisplayAlerts = True

End Sub



Sub recherche_Code_SAP()

'Procédure permettant de chercher le code SAP des differents fichiers et de calculer le somme des stocks IMA logistica pour chaque code SAP

    Dim newlastligne As Long
    Dim WK As Workbook
    Dim code_article As String
    Dim row As Long
    
    
    Dim RowD As Long
    newlastligne = ThisWorkbook.Sheets("Projet").Range("A" & Rows.Count).End(xlUp).row
    Set WK = Workbooks(Workbooks.Count)
    
    For i = 10 To newlastligne
        code_article = ThisWorkbook.Sheets("Projet").Range("B" & i).value
        If (Workbooks(Workbooks.Count).Sheets(1).Range("G:G").Find(code_article) Is Nothing) Then
        
            ThisWorkbook.Sheets("Projet").Range("K" & i) = 0
            ThisWorkbook.Sheets("Projet").Range("D" & i) = ""
            
            
            Else
            
                row = Workbooks(Workbooks.Count).Sheets(1).Range("G:G").Find(code_article).row
                ThisWorkbook.Sheets("Projet").Range("K" & i) = Workbooks(Workbooks.Count).Sheets(1).Range("J" & row)
                ThisWorkbook.Sheets("Projet").Range("D" & i) = Workbooks(Workbooks.Count).Sheets(1).Range("AJ" & row)
        End If
        
        If (ThisWorkbook.Sheets(2).Range("D:D").Find(code_article) Is Nothing) Then
        
            ThisWorkbook.Sheets("Projet").Range("L" & i).Formula = "=SUMIF('IMA Stock Logistica'!D:D,[@[SAP Code]],'IMA Stock Logistica'!F:F)"
        
        Else
        
            row1 = ThisWorkbook.Sheets(2).Range("D:D").Find(code_article).row
            ThisWorkbook.Sheets("Projet").Range("L" & i).Formula = "=SUMIF('IMA Stock Logistica'!D:D,[@[SAP Code]],'IMA Stock Logistica'!F:F)"
            
        End If
    Next i

End Sub


Sub calcul_criticity()

'Procédure permettant de calculer le niveau de criticité pour chaque code SAP

    Dim rng0 As Range
    Dim rng1 As Range
    Dim rng2 As Range
    Dim rng3 As Range
    Dim rngCase As Range
    Dim CaseTable As ListObject
    
    Dim derniere_ligne As Long
    Dim value As String
    derniere_ligne = ThisWorkbook.Sheets("Projet").Range("A" & Rows.Count).End(xlUp).row
    Set rng0 = ThisWorkbook.Sheets("Projet").Range("T9")
    Set rng1 = ThisWorkbook.Sheets("Projet").Range("U9")
    Set rng2 = ThisWorkbook.Sheets("Projet").Range("V9")
    Set rng3 = ThisWorkbook.Sheets("Projet").Range("W9")
    Set rngCase = ThisWorkbook.Sheets("Projet").Range("X9")
    Set CaseTable = ThisWorkbook.Sheets("Projet").ListObjects("CaseTable")
    
    If (derniere_ligne >= 9) Then
        For i = 9 To derniere_ligne
            CaseTable.ListRows.Add
            
            
            If (ThisWorkbook.Sheets("Projet").Range("N" & i) <> "DELIVERY") Then
            
                If (ThisWorkbook.Sheets("Projet").Range("T" & i).value = True) Then
                    
                    ThisWorkbook.Sheets("Projet").Range("M" & i) = 0
                    Call MiseEnCouleur(i)
                
                Else
                
                
                If (ThisWorkbook.Sheets("Projet").Range("U" & i).value = True) Then
                        ThisWorkbook.Sheets("Projet").Range("M" & i) = 1
                        Call MiseEnCouleur(i)
                    Else
                    
                        If (ThisWorkbook.Sheets("Projet").Range("V" & i).value = True) Then
                        
                            ThisWorkbook.Sheets("Projet").Range("M" & i) = 2
                            Call MiseEnCouleur(i)
                        Else
                            ThisWorkbook.Sheets("Projet").Range("M" & i) = 3
                            Call MiseEnCouleur(i)
                    
                End If
                End If
                
                End If
                
            
            
            End If
            
            
            
            
        Next i
    End If
    
    ThisWorkbook.Sheets("Projet").Rows(9).Hidden = True
    ThisWorkbook.Sheets("Projet").Range("M9") = ""
 End Sub
 
 
 
Sub MiseEnCouleur(ligne)

' Procedure permettant de mettre en couleur en fonction du niveau de criticité

    Dim Critic As String
    
    Critic = ThisWorkbook.Sheets("Projet").Range("M" & ligne).value
    
    Select Case Critic
    Case "0"
       
        ThisWorkbook.Sheets("Projet").Range("M" & ligne).Interior.ColorIndex = 43
        
    Case "1"
    
        ThisWorkbook.Sheets("Projet").Range("M" & ligne).Interior.ColorIndex = 6
        
    Case "2"
    
       ThisWorkbook.Sheets("Projet").Range("M" & ligne).Interior.ColorIndex = 44
        
    Case "3"
        
       ThisWorkbook.Sheets("Projet").Range("M" & ligne).Interior.ColorIndex = 3
       
    Case Else
    
         ThisWorkbook.Sheets("Projet").Range("M" & ligne).Interior.ColorIndex = -4142
         
    End Select
 
 
 End Sub


Function Exist(sap_code As String)

'Fonction permettant de vérifier  l'existance d'un code SAP dans la feuille Projet

    If (Not (ThisWorkbook.Sheets("Projet").Range("B:B").Find(sap_code) Is Nothing)) Then
    
        Exist = True
    
    Else
        Exist = False
    End If
    
End Function

Function Nom_projet(i As Long, sh As Worksheet)

' Fonction permettant de déterminer le nom du projet

    If (sh.Range("A" & i) Like "*[A-Za-z][A-Za-z][A-Za-z]##*") Then
        For j = 1 To Len(sh.Range("A" & i))
            
            If (Mid(sh.Range("A" & i), j, 5) Like "[A-Za-z][A-Za-z][A-Za-z]##") Then
            
                Nom_projet = Mid(sh.Range("A" & i), j, 5)
            
                GoTo Fin
            End If
    
        Next j
    
    End If
    
Fin:
End Function

Function supprimer_projet()

'Fonction qui permet de black listé les projets

Dim sh_pilotage As Worksheet
Dim black_list As Variant
Dim lasline As Long
Set black_list = CreateObject("Scripting.Dictionary")

Set sh_pilotage = ThisWorkbook.Sheets("Pilotage")
lastLine = sh_pilotage.Range("A" & Rows.Count).End(xlUp).row

If (lastLine >= 2) Then

For i = 2 To lastLine

black_list.Add UCase(sh_pilotage.Range("A" & i)), ""


Next i

Set supprimer_projet = black_list


End If


End Function


Function calcul_criticité_delivery(ligne As Long, rng As Range)

'Fonction permettant de calculer le niveau de criticité de la phase Delivery

Select Case code_couleur(rng)


Case 43
calcul_criticité_delivery = 0

Case Else
calcul_criticité_delivery = 3


End Select


End Function


Function Folder_check(path)

'Fonction permettant de verifier si un chemin existe dans le disque dur

Dim objFso
Set objFsp = CreateObject("Scripting.FileSystemObject")

Folder_check = objFsp.FolderExists(path)


End Function


Function code_couleur(rng As Range)

'Fonction permettant de determiner le code couleur du niveau de cricitité

code_couleur = rng.Interior.ColorIndex
End Function

Sub CloseAllFiles()

'Procédure permettant de fermer les fichiers après les avoir importés dans le fichier unique

Application.DisplayAlerts = False
Dim WK As Workbook

For Each WK In Workbooks

If (WK.Name <> ThisWorkbook.Name And ListopenedFile.Exists(WK.Name)) Then

    WK.Close


End If

Next WK

ListopenedFile.RemoveAll
Set ListopenedFile = Nothing

Application.DisplayAlerts = True

End Sub


Function checked_Cat()

'Fonction permettant de verifier les category qui sont selectionnés avant de lancer le calcul

Dim sh As Worksheet
Dim SupplyCheck As OLEObject
Dim ProductionCheck As OLEObject
Dim DeliveryCheck As OLEObject


Set sh = ThisWorkbook.Sheets("Projet")
Set SupplyCheck = sh.OLEObjects("SUPPLY")
Set ProductionCheck = sh.OLEObjects("PRODUCTION")
Set DeliveryCheck = sh.OLEObjects("DELIVERY")

checked_Cat = SupplyCheck.Object.value Or ProductionCheck.Object.value Or DeliveryCheck.Object.value


End Function
