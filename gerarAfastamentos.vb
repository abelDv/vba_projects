Sub geraAfastamentos()

    Dim i650, i1012, iE As Integer
    
    'Limpa Planilha
    Call Módulo2.limpaPlan(Plan2)

    'Verifica e abre as planilhas das verbas e a de estornos
    If Dir(ThisWorkbook.Path & "\650.xls") <> "" And Dir(ThisWorkbook.Path & "\1012.xls") <> "" And Dir(ThisWorkbook.Path & "\e.xls") <> "" Then
        'Abre a planilha 650.xls
        Excel.Application.Workbooks.Open (ThisWorkbook.Path & "\650.xls")
    
        'Abre a planilha 1012.xls
        Excel.Application.Workbooks.Open (ThisWorkbook.Path & "\1012.xls")

        'Abre a planilha E.xls
        Excel.Application.Workbooks.Open (ThisWorkbook.Path & "\E.xls")
    Else
        MsgBox "As planilhas não estão configuradas corretamente!" & Chr(13) & "Salve-as com os nomes: '650', '1012', 'E'."
        Exit Sub
    End If
    
    
    'Atribui o numero da ultima linha da planilha 650 para iE
    iE = Módulo2.ultimaLinha
    
    
   
    '--Copia a coluna das matrículas e cola na primeira coluna a esquerda
    'Ativa a planilha "650.xls"
    Windows("650.xls").Activate
    
    'Recorta a coluna G
    Columns("G:G").Cut
    
    'Cola o que foi recortado na coluna A
    Columns("A:A").Insert Shift:=xlToRight
    
    'Apaga a área de transferência
    Application.CutCopyMode = False
    
    'Atribui o numero da ultima linha da planilha à variável i
    i650 = Módulo2.ultimaLinha
    
    
    
    
    '--Copia a coluna das matrículas e cola na primeira coluna a esquerda
    'Ativa a planilha "1012.xls"
    Windows("1012.xls").Activate
    
    'Recorta a coluna G
    Columns("G:G").Cut
    
    'Cola o que foi recortado na coluna A
    Columns("A:A").Insert Shift:=xlToRight
    
    'Apaga a área de transferência
    Application.CutCopyMode = False
    
    
    
    
    '--Copia as linhas da planilha 1012.xls para a planilha 650.xls
    'Atribui o valor da ultima linha da planilha 1012.xls para a variavel i1012
    i1012 = Módulo2.ultimaLinha
    
    'Copia as linhas da planilha 1012.xls
    Rows("2:" & i1012).Copy
    
    'Ativa a planilha "650.xls"
    Windows("650.xls").Activate
    
    'Ativa a primeira celula (abaixo) em branco da planilha 650.xls
    Range("A" & i650 + 1).Activate
    
    'Cola
    ActiveCell.PasteSpecial xlPasteAll
    
    'Apaga a área de transferência
    Application.CutCopyMode = False
    
    'Atribui o numero da ultima linha da planilha à variável i
    i650 = Módulo2.ultimaLinha
     
        
        
    '--Cria a fórmula PROCV que fará a busca de afastamentos na planilha 650.xls
    'Ativa a planilha "650.xls"
    Windows("E.xls").Activate
        
    'Cria uma coluna (C)
    Columns("C:C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    'Insere o título à Coluna C
    Range("C1").Value = "CPF"
        
    'Insere a fórmula PROCV na planilha
    Range("C2").FormulaR1C1 = _
        "=TEXT(VLOOKUP(RC[-1],'[650.xls]650'!C1:C28,27,FALSE),""000.000.000"")&""-""&TEXT(VLOOKUP(RC[-1],'[650.xls]650'!C1:C28,28,FALSE),""00"")"
        

    'Cola a formula para as celulas restantes (abaixo)
    Range("C2").AutoFill Destination:=Range("C2:C" & iE)
    
    '--Filtra as celulas sem afastamento e as exclui da relação
    'Ativa a celula C1
    Range("C1").Activate
    
    'Filtra as colunas C
    Columns("C").AutoFilter
    
    'Filtra não afastados na coluna C
    ActiveSheet.Range("$C$1:$C$" & iE).AutoFilter Field:=1, Criteria1:="#N/D"
    

    'Deleta as linhas sem afastamento
    Rows("2:" & iE).Delete Shift:=xlUp
    
    'Apaga os filtros
    Selection.AutoFilter
    
    
    Columns("C:C").Copy
    
    'Ativa a celula C1
    Range("C1").Activate
    
    'Cola valores
    ActiveCell.PasteSpecial xlPasteValues
    
    'Apaga a área de transferência
    Application.CutCopyMode = False
    
    Columns("A:A").Delete
    Columns("C:C").Delete
    Columns("D:H").Delete
    Columns("E:H").Delete
    Columns("F:M").Delete
    
    Cells.Copy
    
    'Ativa a planilha "GERA AFASTAMENTOS IPREMU.xls"
    Windows("GERA AFASTAMENTOS IPREMU.xls").Activate

    Sheets("Afastados").Select
    Cells.PasteSpecial xlPasteAll
    Cells.PasteSpecial xlPasteFormats
    
    'Filtra as colunas C
    Columns("C").AutoFilter

    Call Módulo2.formata
    
    Call classifica
    
    Cells(1, 1).Value = "MAT"
    Cells(1, 3).Value = "VERBA"
    Cells(1, 4).Value = "VALOR"
    Cells(1, 5).Value = "NOME"
    
    
    Windows("650.xls").Close False
    Windows("1012.xls").Close False
    Windows("E.xls").Close False
            
    
    Call Módulo2.formata
    
    
    MsgBox "Obrigado por aguardar! A planilha está pronta!", vbOKOnly, "Sucesso!"
    
End Sub