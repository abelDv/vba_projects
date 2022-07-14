Sub enviar_email()
    Dim ano, mes, repete As String
    
    If MsgBox("Você está certo de que já pode enviar todos os e-mails?", vbYesNo, "Está certo disso?") = vbNo Then
        Exit Sub
    End If
    
    repete = "sim"
    
    Do While repete = "sim"
        mes = InputBox("Por Gentileza, informe o mês de referência:(dois dígitos)", "Informe o mês!")
        ano = InputBox("Por Gentileza, informe o ano de referência:(quatro dígitos)", "Informe o ano!")

        If MsgBox("Você informou a referência " & mes & "/" & ano & ". Está correto?", vbYesNo, "Está correto?") = vbYes Then
            repete = "não"
        Else
            repete = "sim"
        End If
    Loop
    
    Dim Ulinha As Long
    Dim FSO As New FileSystemObject
    Dim Pasta As Folder
    Dim Arquivo As File
    Dim verba As String

    Ulinha = ultimaLinha("VERBAS - GRUPOS", "H")
 
    If FSO.FolderExists(ThisWorkbook.Path & "\Planilhas") Then
        Set Pasta = FSO.GetFolder(ThisWorkbook.Path & "\Planilhas")
    
        For i = 2 To Ulinha
            verba = Sheets("VERBAS - GRUPOS").Cells(i, 5).Value
            
            If Not IsEmpty(Sheets("VERBAS - GRUPOS").Cells(i, 8).Value) Then
                For Each Arquivo In Pasta.Files
                
                    Set objeto_outlook = CreateObject("Outlook.Application")
                    Set email = objeto_outlook.createitem(0)
                   
                    If Right(Arquivo.Name, 4) = "xlsx" Then
                        If Left(Arquivo.Name, 3) = verba Then
                            email.display
                            email.To = Conversion.CStr(Sheets("VERBAS - GRUPOS").Cells(i, 8).Value)
                            email.Subject = "Informações"
                            email.body = "Olá!" & Chr(10) & Chr(10) & "    Segue, em anexo, a planilha informativa relativa ao mês " & mes & "/" & ano & "."    
                            email.Attachments.Add (ThisWorkbook.Path & "\Planilhas\" & Arquivo.Name)
                            email.send
                        End If
                    End If
                Next
            End If
        Next i
        
    End If
    
    MsgBox "Os e-mails foram encaminhados com sucesso!", vbInformation, "Sucesso!"

End Sub