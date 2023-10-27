Attribute VB_Name = "Module2"
Sub Clear_IMA_logistica_data()
' Cette procedure permet de nettoyer les données dela feuille IMA  stock logistica


    ThisWorkbook.Sheets("IMA Stock Logistica").Cells.Clear
    
    

End Sub


Sub Open_IMA_Logistica_CSV_File()

Application.Workbooks.Open last_Stock_IMA_logistica_File


    

End Sub


Function last_Stock_IMA_logistica_File()

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
        Debug.Print oFile.Name
        If (oFile.Name Like "stock_1600*" And InStr(oFile.Name, ".csv") <> 0) Then
        
            dateString = Mid(oFile.Name, Len("stock1600") + 1, 8)
            dateNumber = CLng(dateString)
            If (dateNumber > max_Date) Then
                max_Date = dateNumber
                file_string = oFile.Name
            End If
        End If
    Next oFile
    
    last_Stock_IMA_logistica_File = ThisWorkbook.Sheets("Pilotage").Range("C5") & "\" & file_string

End Function
Sub copyDataStockIMA()
Dim sh As Worksheet
If (ActiveWorkbook.Name Like "stock_1600*") Then

Set sh = ActiveWorkbook.Sheets(1)
sh.Range("D:D").Copy ThisWorkbook.Sheets("IMA Stock Logistica").Range("D1")

sh.Range("N:S").Copy ThisWorkbook.Sheets("IMA Stock Logistica").Range("N1")

Else

End If

End Sub
