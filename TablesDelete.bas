Sub TablesDelete()
    Dim tbCurrent As Table
    Dim rgReplacement As Range
    
    Do While ActiveDocument.Tables.count > 0
        i = i + 1
        Set tbCurrent = ActiveDocument.Tables(1)
        If tbCurrent.NestingLevel = 1 Then
            Set rgReplacement = tbCurrent.Range.Next(wdParagraph, 1)
            rgReplacement.InsertParagraphBefore
            rgReplacement.InsertParagraphBefore
            rgReplacement.Paragraphs.First.Range.Text = "Tema 5 Tabla " & i
            rgReplacement.Paragraphs.First.Style = wdStyleBlockQuotation
            rgReplacement.Paragraphs.First.Range.Font.Size = 17
            rgReplacement.Paragraphs.First.Alignment = wdAlignParagraphCenter
            tbCurrent.Delete
        End If
    Loop
End Sub