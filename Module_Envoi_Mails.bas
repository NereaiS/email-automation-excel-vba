Attribute VB_Name = "Module1"
Sub Envoi_Mails()

    Dim sh As Worksheet
    Dim i As Long
    Dim last_row As Long
    Dim OA As Object
    Dim msg As Object
    
    Set sh = ThisWorkbook.Sheets("Envoi mails")
    Set OA = CreateObject("Outlook.Application")
    
    ' Trouver la dernière ligne remplie dans la colonne A
    last_row = sh.Cells(sh.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To last_row
        
        If UCase(Trim(sh.Range("H" & i).Value)) <> "NON" And Trim(sh.Range("A" & i).Value) <> "" Then
            
            On Error GoTo ErreurEnvoi
            
            Set msg = OA.CreateItem(0)
            
            With msg
                .To = sh.Range("A" & i).Value
                .CC = sh.Range("B" & i).Value
                .BCC = sh.Range("C" & i).Value
                .Subject = sh.Range("D" & i).Value
                .Body = sh.Range("E" & i).Value
                
                If Trim(sh.Range("F" & i).Value) <> "" Then
                    .Attachments.Add sh.Range("F" & i).Value
                End If
                
                If Trim(sh.Range("G" & i).Value) <> "" Then
                    .Attachments.Add sh.Range("G" & i).Value
                End If
                
                .Send
            End With
            
            sh.Range("I" & i).Value = "Envoyé"
            GoTo SuiteBoucle
            
ErreurEnvoi:
            sh.Range("I" & i).Value = "Erreur"
            Err.Clear
            
SuiteBoucle:
            On Error GoTo 0
            Set msg = Nothing
            
        End If
        
    Next i
    
    MsgBox "Messages envoyés", vbInformation

End Sub

Sub EffacerD()
    ThisWorkbook.Sheets("Envoi mails").Range("D2:D100").ClearContents
End Sub

Sub EffacerE()
    ThisWorkbook.Sheets("Envoi mails").Range("E2:E100").ClearContents
End Sub

Sub EffacerF()
    ThisWorkbook.Sheets("Envoi mails").Range("F2:F100").ClearContents
End Sub

Sub EffacerG()
    ThisWorkbook.Sheets("Envoi mails").Range("G2:G100").ClearContents
End Sub

Sub EffacerH()
    ThisWorkbook.Sheets("Envoi mails").Range("H2:H100").ClearContents
End Sub

Sub EffacerI()
    ThisWorkbook.Sheets("Envoi mails").Range("I2:I100").ClearContents
End Sub

Sub Fichier()

    Dim file_path As Variant
    
    file_path = Application.GetOpenFilename(MultiSelect:=False)
    
    If file_path <> False Then
        Selection.Value = file_path
    End If

End Sub
