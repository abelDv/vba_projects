Sub atualizaBanco()

    Dim i, i1, i2, i3 As Integer
    
    
   'Verifica e abre as planilhas 1, 2 e 3
    If Dir(ThisWorkbook.Path & "\1.xls") <> "" And Dir(ThisWorkbook.Path & "\2.xls") <> "" And Dir(ThisWorkbook.Path & "\3.xls") <> "" Then
        'Abre a planilha 1.xls
        Excel.Application.Workbooks.Open (ThisWorkbook.Path & "\1.xls")
    
        'Abre a planilha 2.xls
        Excel.Application.Workbooks.Open (ThisWorkbook.Path & "\2.xls")

        'Abre a planilha 3.xls
        Excel.Application.Workbooks.Open (ThisWorkbook.Path & "\3.xls")
    Else
        'Caixa de dialogo
        MsgBox "As planilhas apresentam inconsistência ou não foram salvas!" & Chr(13) & "Salve-as com os nomes: 1.xls, 2.xls e 3.xls."
        Exit Sub
    End If
        
    '------------------------
    ' CLASSIFICAR PLANILHAS
    '------------------------
    'Ativa a planilha "1.xls"
    Windows("1.xls").Activate
    
    'Atribui o numero da ultima linha da planilha 1 para a variavel i1
    i1 = Módulo3.ultimaLinha
    
    'Classifica a planilha
    Call Módulo3.classifica(1, i1, "A")


    'Ativa a planilha "2.xls"
    Windows("2.xls").Activate
    
    'Atribui o numero da ultima linha da planilha 1 para a variavel i1
    i2 = Módulo3.ultimaLinha
    
    'Classifica a planilha
    Call Módulo3.classifica(1, i2, "A")


    'Ativa a planilha "3.xls"
    Windows("3.xls").Activate
    
    'Atribui o numero da ultima linha da planilha 1 para a variavel i1
    i3 = Módulo3.ultimaLinha
    
    'Classifica a planilha
    Call Módulo3.classifica(1, i3, "C")
   '-------------------------
        
    Windows("macro.xlsm").Activate
    Sheets("banco").Cells.Clear
        
    '----------------------------------------------------
    ' COLAR TODAS AS PLANILHAS EM BANCO
    '-------------------------------------------------------
    Windows("1.xls").Activate
    Columns("A:W").Cut
    
    Windows("macro.xlsm").Activate
    Sheets("banco").Columns("A:A").Insert Shift:=xlToRight
  
  
    Windows("2.xls").Activate
    Columns("A:AG").Cut
    
    Windows("macro.xlsm").Activate
    Sheets("banco").Columns("A:A").Insert Shift:=xlToRight
    
    
    Windows("3.xls").Activate
    Columns("A:R").Cut
    
    Windows("macro.xlsm").Activate
    Sheets("banco").Columns("A:A").Insert Shift:=xlToRight
    '--------------------------------------------------------------
    
    '----------------------------
    ' EXCLUIR COLUNAS DUPLICADAS
    '----------------------------
    Dim j As Integer
    
    i = 1
    j = 1
    
    Do While Sheets("banco").Cells(1, i).Value <> ""
        
        If j < 80 Then
            For j = i + 1 To 80
                If Sheets("banco").Cells(1, i).Value = Sheets("banco").Cells(1, j).Value Then
                    Sheets("banco").Columns(j).Delete Shift:=xlToLeft
                End If
            Next j
        End If
        
        i = i + 1
        j = i
        
    Loop
        
    '-------------------------------------------------------------------------
    
    '---------------------------------------
    ' APLICAR FÓRMULAS NO FINAL DA PLANILHA
    '---------------------------------------
    
    Sheets("banco").Range("AU1").Value = "SECRETARIA"
    Sheets("banco").Range("AU2").FormulaR1C1 = _
        "=VLOOKUP(VALUE(LEFT(RC[-39],2)),SECRETARIAS!C[-46]:C[-45],2,FALSE)"
        
    Sheets("banco").Range("AV1").Value = "A E O"
    Sheets("banco").Range("AW1").Value = "A SEM O"
    Sheets("banco").Range("AX1").Value = "À E AO"
    Sheets("banco").Range("AY1").Value = "HOJE"
    Sheets("banco").Range("AV2").FormulaR1C1 = "=IF(RC39=""F"",""A"",""O"")"
    Sheets("banco").Range("AW2").FormulaR1C1 = "=IF(RC39=""F"",""A"","""")"
    Sheets("banco").Range("AX2").FormulaR1C1 = "=IF(RC39=""F"",""À"",""AO"")"
    Sheets("banco").Range("AY2").FormulaR1C1 = "=TODAY()"
    Sheets("banco").Range("AU2:AY2").Copy
    
    Sheets("banco").Activate
    
    Sheets("banco").Range("AU2").Activate
    Sheets("banco").Range("AU2:AY" & i1).Select
    ActiveSheet.Paste
    
    Application.CutCopyMode = False
    
    '-------------------------------------------------------------------------
        
    Sheets("MENU").Activate
        
    Windows("1.xls").Close False
    Windows("2.xls").Close False
    Windows("3.xls").Close False
        
   
    MsgBox "Obrigado por aguardar! A planilha está pronta!", vbOKOnly, "Sucesso!"
    
End Sub