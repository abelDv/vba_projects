Sub Chamar()
    Dim xlw As Excel.Workbook
    Dim Plan1, Plan2, Plan3 As String
    Dim lastLine As Integer
        
    Plan1 = "Plan1"
    Plan2 = "Plan2"
    Plan3 = "Plan3"
    
    Sheets(Plan2).Unprotect ("")
        
    Sheets(Plan1).Cells.ClearContents 'limpa a planilha
    
    Call migrar(Plan1, "$A$2", "PMU") 'migra txt para a planilha pmu
    
    lastLine = ultimaLinha(Plan1, "A") + 1
    
    Call migrar(Plan1, "$A$" & lastLine, "PS") 'migra txt para a planilha pmu
    
    Call transformaValor(Plan1)
    
    Call formataCelulas(Plan1)
    
    Call calculaVerbas(Plan1, Plan2, Plan3)
        
    Call incluirCab(Plan1, Plan2, Plan3) 'inclui cabecalho
        
    ActiveWorkbook.Connections("PMU").Delete
    
    ActiveWorkbook.Connections("PS").Delete
        
    Sheets(Plan2).Protect ("")

    MsgBox "Sua Planilha foi atualizada com sucesso!", vbInformation, "Sucesso!"


End Sub

Sub migrar(plan, inicio, arquivo)
    
    With Sheets(plan).QueryTables.Add(Connection:= _
        "TEXT;" & ThisWorkbook.Path & "\" & arquivo & ".txt", Destination:= _
        Sheets(plan).Range(inicio))
        .Name = arquivo
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 850
        .TextFileStartRow = 1
        .TextFileParseType = xlFixedWidth
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileFixedColumnWidths = Array(2, 2, 7, 1, 60, 10, 13, 13, 8, 11, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
End Sub

Sub incluirCab(plan, Plan2, Plan3)

    Sheets(plan).Cells(1, 1).Value = "EMP"
    Sheets(plan).Cells(1, 2).Value = "SEC"
    Sheets(plan).Cells(1, 3).Value = "MAT"
    Sheets(plan).Cells(1, 4).Value = "DIG"
    Sheets(plan).Cells(1, 5).Value = "NOME"
    Sheets(plan).Cells(1, 6).Value = "NASC"
    Sheets(plan).Cells(1, 7).Value = "VALOR"
    Sheets(plan).Cells(1, 8).Value = "FOLHA"
    Sheets(plan).Cells(1, 9).Value = "CPF"
    Sheets(plan).Cells(1, 10).Value = "VERBA"
    
    Sheets(Plan2).Cells(1, 1).Value = "EMP"
    Sheets(Plan2).Cells(1, 2).Value = "SEC"
    Sheets(Plan2).Cells(1, 3).Value = "MAT"
    Sheets(Plan2).Cells(1, 4).Value = "DIG"
    Sheets(Plan2).Cells(1, 5).Value = "NOME"
    Sheets(Plan2).Cells(1, 6).Value = "NASC"
    Sheets(Plan2).Cells(1, 7).Value = "VALOR"
    Sheets(Plan2).Cells(1, 8).Value = "FOLHA"
    Sheets(Plan2).Cells(1, 9).Value = "CPF"
    Sheets(Plan2).Cells(1, 10).Value = "VERBA"
    
    Sheets(Plan3).Cells(1, 1).Value = "VERBA"
    Sheets(Plan3).Cells(1, 2).Value = "VALOR"

End Sub

Sub incluirCabRestante(plan, Plan2, Plan3)
    Sheets(Plan2).Cells(1, 1).Value = "EMP"
    Sheets(Plan2).Cells(1, 2).Value = "SEC"
    Sheets(Plan2).Cells(1, 3).Value = "MAT"
    Sheets(Plan2).Cells(1, 4).Value = "DIG"
    Sheets(Plan2).Cells(1, 5).Value = "NOME"
    Sheets(Plan2).Cells(1, 6).Value = "NASC"
    Sheets(Plan2).Cells(1, 7).Value = "VALOR"
    Sheets(Plan2).Cells(1, 8).Value = "FOLHA"
    Sheets(Plan2).Cells(1, 9).Value = "CPF"
    Sheets(Plan2).Cells(1, 10).Value = "VERBA"
    
    Sheets(Plan3).Cells(1, 1).Value = "VERBA"
    Sheets(Plan3).Cells(1, 2).Value = "VALOR"
End Sub

Sub transformaValor(plan)

    Sheets(plan).Range("H2").FormulaR1C1 = "=IF(RC[-1]="""","""",VALUE(TEXT(RC[-1],""00\,00"")))"
    Sheets(plan).Range("H2").Copy
    Sheets(plan).Range("H2:H50000").PasteSpecial xlPasteFormulas
    Sheets(plan).Range("H:H").Copy
    Sheets(plan).Range("H:H").PasteSpecial xlPasteValues
    Sheets(plan).Range("G:G").Delete
    Sheets(plan).Range("J:J").Delete
        
End Sub

Sub formataCelulas(plan)

    Sheets(plan).Columns("I:I").NumberFormat = "000"".""000"".""000""-""00"
    Sheets(plan).Columns("G:G").NumberFormat = "$ #,##0.00"
    Sheets(plan).Cells.ColumnWidth = 50#
    Sheets(plan).Cells.EntireColumn.AutoFit
    Sheets(plan).Columns("B:B").NumberFormat = "00"
    
End Sub

Sub calculaVerbas(plan, Plan2, Plan3)
    Dim iPlan1, iPlan2, iPlan3, compara As Long
    Dim inicio As Integer
    Dim soma As Double
    
    iPlan1 = 2
    iPlan2 = 2
    iPlan3 = 2
    inicio = 2
    soma = 0#
    
    ActiveWorkbook.Worksheets(plan).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(plan).Sort.SortFields.Add Key:=Range("J2:J50000" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets(plan).Sort.SortFields.Add Key:=Range("I2:I50000" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets(plan).Sort
            .SetRange Range("A1:J50000")
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
    compara = Sheets(plan).Range("J2")
    
    Sheets(Plan2).Columns("A:J").ClearContents
    Sheets(Plan3).Columns("A:B").ClearContents
        
    Do
        
        If Sheets(plan).Range("J" & iPlan1).Value = compara Then
            soma = soma + Sheets(plan).Range("G" & iPlan1).Value
                        
            'Sheets(plan).Rows(iPlan1 & ":" & iPlan1).Copy
            'Sheets(plan2).Range("A" & iPlan2).PasteSpecial xlPasteValues
            'Application.CutCopyMode = False
            'iPlan2 = iPlan2 + 1
        
        ElseIf Sheets(plan).Range("J" & iPlan1) <> "" Then
            Sheets(plan).Range("A" & inicio & ":J" & iPlan1 - 1).Copy
            Sheets(Plan2).Range("A2").PasteSpecial xlPasteValues
            Sheets(Plan3).Range("A" & iPlan3).Value = compara
            Sheets(Plan3).Range("B" & iPlan3).Value = soma
            
            Call incluirCab(plan, Plan2, Plan3)
            
            Call CriaArquivo(Sheets(Plan2), ThisWorkbook.Path, compara)
            
            iPlan3 = iPlan3 + 1
            inicio = iPlan1
            
            soma = Sheets(plan).Range("G" & iPlan1).Value
            
            compara = Sheets(plan).Range("J" & iPlan1).Value
                
            Sheets(Plan2).Columns("A:J").ClearContents
            
            
            'Sheets(plan2).Range("F50000").FormulaR1C1 = "=SUM(R[-49998]C:R[-1]C)"
            'Sheets(plan2).Range("F50000").Copy
            'Sheets(plan3).Range("B" & iPlan3).PasteSpecial xlPasteValues
            'Sheets(plan3).Range("A" & iPlan3).Value = compara
            'iPlan3 = iPlan3 + 1
            'Sheets(plan2).Cells.Clear
            'iPlan2 = 2
            'compara = Sheets(plan).Range("I" & iPlan1).Value
            'iPlan1 = iPlan1 - 1
        Else
            Sheets(plan).Range("A" & inicio & ":J" & iPlan1 - 1).Copy
            Sheets(Plan2).Range("A2").PasteSpecial xlPasteValues
            Sheets(Plan3).Range("A" & iPlan3).Value = compara
            Sheets(Plan3).Range("B" & iPlan3).Value = soma
            
            Call incluirCab(plan, Plan2, Plan3)
            
            Call CriaArquivo(Sheets(Plan2), ThisWorkbook.Path, compara)
        End If
       
       iPlan1 = iPlan1 + 1
       
    Loop Until Sheets(plan).Range("J" & iPlan1 - 1) = ""
    
End Sub

Function ultimaLinha(plan, coluna) As Integer
    
    Dim rLast As Long
    
    rLast = Sheets(plan).Range(coluna & "65536").End(xlUp).Row
    
    ultimaLinha = rLast
    
End Function

Sub CriaArquivo(mPlan As Worksheet, mPathSave As String, verba)
Dim NovoArquivoXLS As Workbook
Dim sht As Worksheet

    'Cria um novo arquivo excel
    Set NovoArquivoXLS = Application.Workbooks.Add

    'Copia a planilha para o novo arquivo criado
    mPlan.Copy Before:=NovoArquivoXLS.Sheets(1)

    Sheets("Plan2 (2)").Cells.Copy
    
    Sheets("Plan2 (2)").Cells.PasteSpecial xlPasteValues
    
    'Salva o arquivo
    NovoArquivoXLS.SaveAs mPathSave & "\Planilhas\" & verba & ".xlsx"
    
    
    'Fecha o arquivo
    NovoArquivoXLS.Close

End Sub