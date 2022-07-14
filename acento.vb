Function Acento(caract)

    'Acentos e caracteres especiais que serão buscados na string
    'Você pode definir outros caracteres nessa variável, mas
    ' precisará também colocar a letra correspondente em codiB
    codiA = "àáâãäèéêëìíîïòóôõöùúûüÀÁÂÃÄÈÉÊËÌÍÎÒÓÔÕÖÙÚÛÜçÇñÑ/"
    
    'Letras correspondentes para substituição
    codiB = "aaaaaeeeeiiiiooooouuuuAAAAAEEEEIIIOOOOOUUUUcCnN "
    
    'Armazena em temp a string recebida
    temp = caract
    
    'Loop que irá de andará a string letra a letra
    For i = 1 To Len(temp)
    
        'InStr buscará se a letra indice i de temp pertence a
        ' codiA e se existir retornará a posição dela
        p = InStr(codiA, Mid(temp, i, 1))
        
        'Substitui a letra de indice i em codiA pela sua
        ' correspondente em codiB
        If p > 0 Then Mid(temp, i, 1) = Mid(codiB, p, 1)
    Next
    
    'Retorna a nova string
    Acento = temp
    
End Function