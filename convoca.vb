Sub Convocar()
    
    Dim ordemGeral, totalAprovados, qtdConv As Long
    
   
    If MsgBox("Você está prestes a convocar candidatos. Esta certo disso?", vbYesNoCancel, "Está certo?") = vbYes Then

        qtd = Conversion.CInt(Sheets("MENU").Range("G2").Value)
        ordemGeral2 = Conversion.CInt(Sheets("MENU").Range("L5").Value) + 1
        ordemGeral = Conversion.CInt(Sheets("MENU").Range("L4").Value) + 1
        totalAprovados = Conversion.CInt(Sheets("CONVOCADOS").Range("B" & (Conversion.CInt(Sheets("MENU").Range("K1").Value) + 4)).Value)
        
        Sheets("MENU").Range("Q2:V501").ClearContents
        Sheets("MENU").Range("X2:AA501").ClearContents
        Sheets("MENU").Range("AF2:AF501").ClearContents
        
        For i = 1 To qtd
            
            If totalAprovados >= ordemGeral2 Then

                If (right(ordemGeral, 1) = "3") Or (right(ordemGeral, 1) = "8") Then
                    Call convocarNegro(1, (i), False)

                ElseIf ((ordemGeral = 5) Or (right(ordemGeral, 1) = "1")) And (ordemGeral > 1) Then
                    Call convocarDeficiente(1, (i), False)

                Else
                    Call convocarAmpla(1, (i), False)

                End If
                
                ordemGeral2 = ordemGeral2 + 1
                ordemGeral = ordemGeral + 1

            Else
                MsgBox "Não há mais candidatos aprovados para este cargo!"
                Exit Sub
            End If
        
        Next
    
        MsgBox "Convocação realizada com sucesso!", vbInformation, "Sucesso!"
        
    End If
    
End Sub