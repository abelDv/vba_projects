Function EscreverNoWord()
    Dim objWord 'o aplicativo Microsoft Word
    Dim objDoc 'o documento do word em si (doc ou docx)
    Dim objSelection 'o cursor de operações
    Dim text As String
 
    Set objWord = CreateObject("Word.Application") ' coloca o word na memória
    Set objDoc = objWord.Documents.Add ' Word, crie um novo documento em branco
 
    objWord.Visible = True ' Word, mostre sua cara
 
    Set objSelection = objWord.Selection 'me dê a referência do cursor do mouse
 
    For Each cell In Sheets("Decreto").Range("A1:A6")
        text = text & cell.Value & Chr(13)
    Next cell
 
    objSelection.TypeText (text) 'Digitar texto
End Function

Sub criaDecreto()

    
    If MsgBox("Você está prestes a gerar os Decretos de Nomeação! Está certo disso?", vbYesNoCancel, "Está certo?") = vbYes Then

        Dim wdApp As Word.Application
        Dim wdDoc As Word.Document
        Dim wdRange1 As Word.Range
        Dim wdRange2 As Word.Range
        Dim wdRange3 As Word.Range
        Dim wdRange4 As Word.Range
        Dim wdRange5 As Word.Range
        Dim wdRange6 As Word.Range
        Dim wdRange7 As Word.Range
        Dim wdRange8 As Word.Range
        Dim wdRange9 As Word.Range
        Dim nomeArquivo As String
        
        Dim wbBook As Workbook
        Dim wsSheet As Worksheet
        
        Dim vaData As Variant
        
        Set wbBook = ThisWorkbook
        Set wsSheet = wbBook.Worksheets("Decreto")
        
        vaData = wsSheet.Range("W_Data").Value
        
        ' Instantiate the Word Objects.
        Set wdApp = New Word.Application
        Set wdDoc = wdApp.Documents.Open(wbBook.Path & "\Modelo Nomeia.docx")
        
        Sheets("MENU").Range("P1:V" & Sheets("MENU").Cells(Rows.Count, 17).End(xlUp).Row).Copy
        
        With wdDoc
            Set wdRange1 = .Bookmarks(Sheets("Decreto").Range("B1").Value).Range
            Set wdRange2 = .Bookmarks(Sheets("Decreto").Range("B2").Value).Range
            Set wdRange3 = .Bookmarks(Sheets("Decreto").Range("B3").Value).Range
            Set wdRange4 = .Bookmarks(Sheets("Decreto").Range("B4").Value).Range
            Set wdRange5 = .Bookmarks(Sheets("Decreto").Range("B5").Value).Range
            Set wdRange6 = .Bookmarks(Sheets("Decreto").Range("B6").Value).Range
            Set wdRange7 = .Bookmarks(Sheets("Decreto").Range("B7").Value).Range
            Set wdRange8 = .Bookmarks("tabela").Range
            Set wdRange9 = .Bookmarks(Sheets("Decreto").Range("B8").Value).Range
        End With
        
        ' Write values to the bookmarks.
        wdRange1.text = vaData(1, 1)
        wdRange2.text = vaData(2, 1)
        wdRange3.text = vaData(3, 1)
        wdRange4.text = vaData(4, 1)
        wdRange5.text = vaData(5, 1)
        wdRange6.text = vaData(6, 1)
        wdRange7.text = vaData(7, 1)
        wdRange9.text = vaData(8, 1)
        
        wdRange8.PasteExcelTable False, False, False
        
        wdRange8.Rows.AllowBreakAcrossPages = False
        wdRange8.Rows.HeadingFormat = True
        
        If Sheets("MENU").Range("G2").Value > 1 Then
            nomeArquivo = Left(Sheets("MENU").Range("V2").Value, 15) & " E OUTROS - " & Módulo1.Acento(Left(Sheets("MENU").Range("L3").Value, 20))
        ElseIf Sheets("MENU").Range("G2").Value = 2 Then
            nomeArquivo = Left(Sheets("MENU").Range("V2").Value, 15) & " E OUTRO - " & Módulo1.Acento(Left(Sheets("MENU").Range("L3").Value, 20))
        Else
            nomeArquivo = Left(Sheets("MENU").Range("V2").Value, 15) & " - " & Módulo1.Acento(Left(Sheets("MENU").Range("L3").Value, 20))
        End If
        
        With wdDoc
            .SaveAs wbBook.Path & "\ATOS\" & nomeArquivo & ".docx"
            .Close
        End With
        
        wdApp.Quit
        
        ' Release the objects from memory.
        Set wdRange1 = Nothing
        Set wdRange2 = Nothing
        Set wdRange3 = Nothing
        Set wdRange4 = Nothing
        Set wdRange5 = Nothing
        Set wdRange6 = Nothing
        Set wdRange7 = Nothing
        Set wdDoc = Nothing
        Set wdApp = Nothing
        
        Application.CutCopyMode = False

        MsgBox "Decreto criado com sucesso!", vbInformation, "Sucesso!!!"
    
    End If
    
End Sub