Attribute VB_Name = "Module1"
Sub Bouton4_Cliquer()

'Procédure permettant d'afficher un message avant de supprimer les données

'Remove data

Dim reponse 'reserver une espace mémoire

reponse = MsgBox("Are you sure you want to clear all data?", vbYesNo)
If (reponse = vbYes) Then

Call Module1.ClearData
Call Module1.ClearLatestUpdate

End If

End Sub

Sub ClearData()
Attribute ClearData.VB_ProcData.VB_Invoke_Func = " \n14"

'Procédure permettant de resize le "table1" dans la feuille projet

Dim sh As Worksheet
Dim table1 As ListObject

Set sh = ThisWorkbook.Sheets("Projet")
Set table1 = sh.ListObjects("Table1")

table1.AutoFilter.ShowAllData

 Call Module1.resizetab




   Rows("10:" & Rows.Count).Select
    Selection.Delete Shift:=xlUp
    
   'Call Module1.clearCaseTable
End Sub


Sub resizetab()

'Procédure permettant de resize le "CaseTable" dans la feuille projet

Dim CaseTable As ListObject
Dim listrow As listrow
Set CaseTable = ThisWorkbook.Sheets("Projet").ListObjects("CaseTable")

CaseTable.Resize (ThisWorkbook.Sheets("Projet").Range("T8:AO9"))

End Sub
Sub clearCaseTable()

' Procédure permettant de supprimer le "CaseTable" à la fin du calcul

Dim CaseTable As ListObject
Dim listrow As listrow
Set CaseTable = ThisWorkbook.Sheets("Projet").ListObjects("CaseTable")

Set listrow = CaseTable.ListRows(CaseTable.ListRows.Count)
While (CaseTable.ListRows.Count > 1)
listrow.Delete
Set listrow = CaseTable.ListRows(CaseTable.ListRows.Count)
Wend

End Sub


Sub ClearLatestUpdate()

'Procédure permettant de supprimer les données une fois avoir appuyer sur le bouton "Clear Latest Update"

Range("H1").ClearContents

End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
 
    Range("C2").Select
End Sub



Sub fileLastDate()

'Procédure permettant de remplir la date de la dernière reception du fichier du fournisseur



Dim TabDate As ListObject
Dim sh As Worksheet
Dim ligne_project_name As Long
Dim folder_path As String
Dim date1  As String


Set fs = CreateObject("Scripting.FileSystemObject")
Set sh = ThisWorkbook.Sheets("Update List")
Set TabDate = sh.ListObjects("UpdateListTable")




For i = 2 To TabDate.ListRows.Count


If (TabDate.Range(i, 1) = "") Then
ligne_project_name = TabDate.Range(i, 1).End(xlUp).row
projectName = sh.Range("B" & ligne_project_name)


Else
projectName = TabDate.Range(i, 1)

End If

category = TabDate.Range(i, 5)
destination = TabDate.Range(i, 3)


If (projectName <> "" And category <> "" And destination <> "") Then

Select Case category

Case "SUPPLY"

folder_path = ThisWorkbook.Sheets("Pilotage").Range("C2").value
Set oFolder = fs.GetFolder(folder_path)

For Each oFile In oFolder.Files

If (oFile.Name Like "*.xls*" And UCase(oFile.Name) Like "*" & UCase(projectName) & "*" And UCase(oFile.Name) Like "*" & UCase(projectName) & "*" And UCase(oFile.Name) _
Like "*" & UCase(Trim(destination)) & "*" And UCase(oFile.Name) Like "*??-??-????*") Then

' TabDate.Range(i, 6).NumberFormat = "dd-mm-yyyy"
date1 = Mid(oFile.Name, InStr(oFile.Name, ".") - Len("00-00-0000"), Len("00-00-0000"))

TabDate.Range(i, 6) = Mid(date1, 4, 2) & "-" & Left(date1, 2) & "-" & Right(date1, 4)



Exit For
End If

Next oFile


Case "DELIVERED"

folder_path = ThisWorkbook.Sheets("Pilotage").Range("C4").value
Set oFolder = fs.GetFolder(folder_path)

For Each oFile In oFolder.Files

If (oFile.Name Like "*.xls*" And UCase(oFile.Name) Like "*" & UCase(projectName) & "*" And UCase(oFile.Name) Like "*" & UCase(projectName) & "*" And UCase(oFile.Name) _
Like "*" & UCase(Trim(destination)) & "*" And UCase(oFile.Name) Like "*??-??-????*") Then
 'TabDate.Range(i, 6).NumberFormat = "dd-mm-yyyy"
 
date1 = Mid(oFile.Name, InStr(oFile.Name, ".") - Len("00-00-0000"), Len("00-00-0000"))

TabDate.Range(i, 6) = Mid(date1, 4, 2) & "-" & Left(date1, 2) & "-" & Right(date1, 4)
 
Exit For
End If

Next oFile


Case "PRODUCTION"

folder_path = ThisWorkbook.Sheets("Pilotage").Range("C3").value
Set oFolder = fs.GetFolder(folder_path)

For Each oFile In oFolder.Files

If (oFile.Name Like "*.xls*" And UCase(oFile.Name) Like "*" & UCase(projectName) & "*" And UCase(oFile.Name) Like "*" & UCase(projectName) & "*" And UCase(oFile.Name) _
Like "*" & UCase(Trim(destination)) & "*" And UCase(oFile.Name) Like "*??-??-????*") Then
 'TabDate.Range(i, 6).NumberFormat = "dd-mm-yyyy"
date1 = Mid(oFile.Name, InStr(oFile.Name, ".") - Len("00-00-0000"), Len("00-00-0000"))

TabDate.Range(i, 6) = Mid(date1, 4, 2) & "-" & Left(date1, 2) & "-" & Right(date1, 4)
Exit For
End If

Next oFile

End Select


Else

 TabDate.Range(i, 6) = ""

End If

Next i


End Sub



Sub parcours_file(Dic As Dictionary, cat As String)

'Procédure permettant de parcourir le browser en allant recuperer les données des fichiers fournisseurs dans leur catégorie correspondante

Dim objFso
Dim path As String
Dim TabDate As ListObject
Dim sh As Worksheet


Set sh = ThisWorkbook.Sheets("Update List")
Set TabDate = sh.ListObjects("UpdateListTable")
Set objFsp = CreateObject("Scripting.FileSystemObject")

Select Case cat

Case "SUPPLY"
path = ThisWorkbook.Sheets("Pilotage").Range("C2").value
Set Folder = objFsp.GetFolder(path)

For Each File In Folder.Files

If (File.Name Like "*xls*") Then


Workbooks.Open (File.path)

LastLineA = Workbooks(Workbooks.Count).Sheets(1).Range("A" & Rows.Count).End(xlUp).row
For Each Key In Dic.Keys
For i = 2 To LastLineA
Debug.Print Workbooks(Workbooks.Count).Sheets(1).Range("A" & i).value
If (Workbooks(Workbooks.Count).Sheets(1).Range("A" & i).value Like Key) Then

'*****************************************************************************************************************

If (Workbooks(Workbooks.Count).Name Like "*.xls*" And UCase(Workbooks(Workbooks.Count).Name) Like "*??-??-????*") Then

' TabDate.Range(i, 6).NumberFormat = "dd-mm-yyyy"
date1 = Mid(Workbooks(Workbooks.Count).Name, InStr(Workbooks(Workbooks.Count).Name, ".") - Len("00-00-0000"), Len("00-00-0000"))

Debug.Print Dic(Key)
TabDate.DataBodyRange(Dic(Key), 6) = Mid(date1, 4, 2) & "-" & Left(date1, 2) & "-" & Right(date1, 4)
GoTo keynext

End If

'*********************************************************




End If
Next i
keynext:
Next Key

End If


Workbooks(Workbooks.Count).Close
Next File


'******************************************  production********************************



Case "PRODUCTION"
path = ThisWorkbook.Sheets("Pilotage").Range("C3").value
Set Folder = objFsp.GetFolder(path)





For Each Key In Dic.Keys

For Each File In Folder.Files

If (File.Name Like "*xls*") Then
If (File.Name Like "*.xls*" And UCase(File.Name) Like "*??-??-????*") Then

' TabDate.Range(i, 6).NumberFormat = "dd-mm-yyyy"
date1 = Mid(File.Name, InStr(File.Name, ".") - Len("00-00-0000"), Len("00-00-0000"))

If (UCase(File.Name) Like UCase(Key)) Then

TabDate.DataBodyRange(Dic(Key), 6) = Mid(date1, 4, 2) & "-" & Left(date1, 2) & "-" & Right(date1, 4)

End If
End If

'*********************************************************


End If



Next File
Next Key

Case "DELIVERED"
path = ThisWorkbook.Sheets("Pilotage").Range("C4").value
Set Folder = objFsp.GetFolder(path)

For Each File In Folder.Files

If (File.Name Like "*xls*") Then


Workbooks.Open (File.path)

LastLineA = Workbooks(Workbooks.Count).Sheets(1).Range("A" & Rows.Count).End(xlUp).row
For Each Key In Dic.Keys
For i = 2 To LastLineA
Debug.Print Workbooks(Workbooks.Count).Sheets(1).Range("A" & i).value
If (Workbooks(Workbooks.Count).Sheets(1).Range("A" & i).value Like Key) Then

'*****************************************************************************************************************

If (Workbooks(Workbooks.Count).Name Like "*.xls*" And UCase(Workbooks(Workbooks.Count).Name) Like "*??-??-????*") Then

' TabDate.Range(i, 6).NumberFormat = "dd-mm-yyyy"
date1 = Mid(Workbooks(Workbooks.Count).Name, InStr(Workbooks(Workbooks.Count).Name, ".xl") - Len("00-00-0000"), Len("00-00-0000"))

Debug.Print Dic(Key)
TabDate.DataBodyRange(Dic(Key), 6) = Mid(date1, 4, 2) & "-" & Left(date1, 2) & "-" & Right(date1, 4)
GoTo keynext1

End If

'*********************************************************


End If
Next i
keynext1:
Next Key

End If


Workbooks(Workbooks.Count).Close
Next File

'***************************************************************************************************


End Select
End Sub
Sub rechercheChaine()

'Procédure permettant d'aller chercher les chaines de caratère par projet pour les importer dans la "Update List"

Dim TabDate As ListObject
Dim sh As Worksheet
Dim row As listrow
Dim chaine_tabSupply As Dictionary
Dim chaine_tabDelivered As Dictionary
Dim chaine_tabProduction As Dictionary
Dim destination As String
Dim projectName As String
Dim categorie As Dictionary
Dim rowCategorie As Dictionary
Set sh = ThisWorkbook.Sheets("Update List")
Set TabDate = sh.ListObjects("UpdateListTable")
projectName = TabDate.DataBodyRange(1, 1)

Set chaine_tabSupply = New Dictionary
Set chaine_tabDelivered = New Dictionary
Set chaine_tabProduction = New Dictionary

Set categorie = New Dictionary
Set rowCategorie = New Dictionary
categorie.Add "SUPPLY", 1
categorie.Add "PRODUCTION", 2
categorie.Add "DELIVERED", 3

For Each row In TabDate.ListRows




    If (TabDate.DataBodyRange(row.Index, 1) <> "") Then
    
    
        projectName = TabDate.DataBodyRange(row.Index, 1)
        
    End If
  
    If (categorie.Exists(TabDate.DataBodyRange(row.Index, 5).value) = True) Then
    
    
    
        Select Case TabDate.DataBodyRange(row.Index, 2)
        
            Case "cabinet"
            destination = Trim(TabDate.DataBodyRange(row.Index, 3)) & "*" & "QE"
            
            Case Else
            
            destination = TabDate.DataBodyRange(row.Index, 3)
        
        End Select
        
        Select Case TabDate.DataBodyRange(row.Index, 5).value
        
        Case "SUPPLY"
        chaine_tabSupply.Add "*" & projectName & "*" & Trim(destination) & "*", row.Index
        
        If (Right(destination, 2) = "QE") Then
        
        chaine_tabSupply.Add "*" & projectName & "*" & Trim(TabDate.DataBodyRange(row.Index, 3)) & "*" & "CABINETS*", row.Index
        
        chaine_tabSupply.Add "*" & projectName & "*" & "CABINETS*" & Trim(TabDate.DataBodyRange(row.Index, 3)) & "*", row.Index
        
        End If
        
        Case "PRODUCTION"
          If (destination Like "*M##*QE*" And UCase(TabDate.DataBodyRange(row.Index, 2).value) = "CABINET") Then
          
           
           destination = Replace(destination, "M", "")
           destination = Replace(destination, "*QE", "")
           destination = "QE" & destination
                  
        End If
        
        
        If (UCase(TabDate.DataBodyRange(row.Index, 3).value) = "ETH" And UCase(TabDate.DataBodyRange(row.Index, 2).value) = "CABINET") Then
        
        destination = Replace(destination, "*QE", "")
       
        End If
        
      
           
        chaine_tabProduction.Add "*" & projectName & "*" & Trim(destination) & "*", row.Index
          
        
        
        Case "DELIVERED"
        chaine_tabDelivered.Add "*" & projectName & "*" & Trim(destination) & "*", row.Index
        
        If (Right(destination, 2) = "QE") Then
        chaine_tabDelivered.Add "*" & projectName & "*" & Trim(TabDate.DataBodyRange(row.Index, 3)) & "*" & "CABINETS*", row.Index
        chaine_tabDelivered.Add "*" & projectName & "*" & "CABINETS*" & Trim(TabDate.DataBodyRange(row.Index, 3)) & "*", row.Index
        
        End If
        
        
        
        End Select
    
    End If
    
Next row

Call Module1.parcours_file(chaine_tabSupply, "SUPPLY")
Call Module1.parcours_file(chaine_tabDelivered, "DELIVERED")
Call Module1.parcours_file(chaine_tabProduction, "PRODUCTION")

End Sub

Sub MettreAJourDate()

'Procédure permettant de mettre à jour la date une fois le calcul terminé

    Range("H1").value = Date
       
End Sub


Sub clear_date_from_last_date_file()

'Procédure permettant de supprimer les dates une fois après avoir appuyer sur le bouton "Update List"

Dim tabledate As ListObject
Dim sh As Worksheet
Dim lesdates As Range
Set sh = ThisWorkbook.Worksheets("Update List")
Set tabledate = sh.ListObjects("UpdateListTable")
Set lesdates = tabledate.DataBodyRange.Columns(6)
lesdates.ClearContents


End Sub

Sub Uncheck_Category()

'Procédure permettant d'enlever le "check" sur toute les catégories une fois avoir fermé le fichier


ThisWorkbook.Sheets("Projet").OLEObjects("SUPPLY").Object.value = False
ThisWorkbook.Sheets("Projet").OLEObjects("PRODUCTION").Object.value = False
ThisWorkbook.Sheets("Projet").OLEObjects("DELIVERY").Object.value = False


End Sub

