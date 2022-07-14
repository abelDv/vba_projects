Private Sub CommandButton2_Click()
    Dim i As Integer
    Dim apagou As Boolean
       
    
    If MsgBox("Deseja realmente apagar as planilhas?", vbYesNo, "Deseja apagar?") = vbYes Then
        For i = 500 To 1000
            If Dir(ThisWorkbook.Path & "\Planilhas\" & i & ".xlsx") <> vbNullString Then
                apagou = True
                Kill (ThisWorkbook.Path & "\Planilhas\" & i & ".xlsx")
            End If
            
        Next i
        
        If apagou Then
            MsgBox "As planilhas foram apagadas com sucesso!", vbInformation, "Sucesso!"
        Else
            MsgBox "Não há planilhas para serem apagadas!", vbInformation, "Ok!"
        End If
    End If
    
End Sub