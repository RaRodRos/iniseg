Attribute VB_Name = "InisegStory"
Option Explicit

Private Function ConversionParrafos()
'
' Iniseg4ConversionParrafos Macro
' Conversion de Word impreso a formato para Storyline
'
    ' Cambio del tamaño de Titulo 1 de 16 a 17
    
    Application.Run "RaMacros.LimpiarFindAndReplaceParameters"
    
    With ActiveDocument.Styles(wdStyleHeading1).Font
        .Name = "Swis721 Lt BT"
        .Size = 17
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = True
        .Color = -738148353
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Scaling = 100
        .Kerning = 0
        .Animation = wdAnimationNone
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    With ActiveDocument.Styles(wdStyleHeading1)
        .AutomaticallyUpdate = False
        .BaseStyle = "List Paragraph"
        .NextParagraphStyle = "Normal"
    End With

    ' Eliminar ALLCAPS de los títulos 2 y 3
    ActiveDocument.Styles(wdStyleHeading2).Font.AllCaps = False
    ActiveDocument.Styles(wdStyleHeading3).Font.AllCaps = False

    ' Poner el estilo quote centrado y sin espacio a derecha ni izquierda
    With ActiveDocument.Styles(wdStyleQuote).ParagraphFormat
        .LeftIndent = 0
        .RightIndent = 0
        .Alignment = wdAlignParagraphCenter
    End With

' Cambio de tamaño de parrafos de separacion

    ' Listas: 4 a 2
    With ActiveDocument.Range.Find
        .ClearFormatting
        .Font.Size = 4
        .Replacement.ClearFormatting
        .Replacement.Font.Size = 2
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With

    ' Parrafos normales: 5 a 4.
    With ActiveDocument.Range.Find
        .ClearFormatting
        .Font.Size = 5
        .Replacement.ClearFormatting
        .Replacement.Font.Size = 4
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With

    ' Titulos 2, 3 y 4: 8 a 6.
    With ActiveDocument.Range.Find
        .ClearFormatting
        .Font.Size = 8
        .Replacement.ClearFormatting
        .Replacement.Font.Size = 6
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With

' Titulos 1: 11 a 8

    ' Dar tamaño 8 a todos los párrafos tras los Heading 1
    With ActiveDocument.Range.Find
        .ClearFormatting
        .Style = ActiveDocument.Styles(wdStyleHeading1)
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Text = "(*@^13)"
        .Replacement.Text = "\1FISTRO"
        .Execute Replace:=wdReplaceAll
        .Style = ActiveDocument.Styles(wdStyleNormal)
        .Text = "FISTRO^13"
        .Replacement.Font.Size = 8
        .Replacement.Text = "^13"
        .Execute Replace:=wdReplaceAll
    End With

    ' Meter salto de página antes de cada Heading 1 y Title
    Dim prParrafoActual As Paragraph, index As Integer

    For index = 1 To ActiveDocument.Paragraphs.Count - 1
        With ActiveDocument.Paragraphs(index).Range
            If .Next(Unit:=wdParagraph, Count:=1).Paragraphs(1).OutlineLevel = 1 Then
                If .Previous(Unit:=wdParagraph, Count:=2).Style <> ActiveDocument.Styles(wdStyleTitle) Then
                    '.Collapse Direction:=wdCollapseEnd
                    ActiveDocument.Paragraphs(index).Range.InsertBreak Type:=wdPageBreak
                    'index = index + 1
                End If
            End If
        End With
    Next index

    Application.Run "RaMacros.LimpiarFindAndReplaceParameters"
    
End Function

Private Function ImagenesGrandes()
'
' ImagenesGrandes Function
'
' Hace que todas las imágenes sean enormes, para meterlas en el story
'
    Dim inlShape As InlineShape
        
    For Each inlShape In ActiveDocument.InlineShapes
        If inlShape.Type = wdInlineShapePicture Then inlShape.Width = CentimetersToPoints(29)
    Next inlShape

End Function

Private Function TitulosCon3Espacios()
'
' TitulosCon3Espacios Function
'
' Sustituye la tabulación en los títulos por 3 espacios
'
    With ActiveDocument.Range.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Text = "([0-9].)^t"
        .Replacement.Text = "\1   "
        .Style = ActiveDocument.Styles(wdStyleHeading1)
        .Execute Replace:=wdReplaceAll
        .Style = ActiveDocument.Styles(wdStyleHeading2)
        .Execute Replace:=wdReplaceAll
        .Style = ActiveDocument.Styles(wdStyleHeading3)
        .Execute Replace:=wdReplaceAll

    End With

    Application.Run "RaMacros.LimpiarFindAndReplaceParameters"

End Function

Sub Iniseg4FormatoStory()

    Application.Run "InisegLibro.InisegInterlineado"
    Application.Run "RaMacros.ListasATexto"
    Application.Run "ConversionParrafos"
    Application.Run "TitulosCon3Espacios"
    Application.Run "ImagenesGrandes"
    Application.Run "InisegLibro.InisegInterlineado"

End Sub

Sub Iniseg5NotasPieATexto()
'
' NotasPieATexto Macro
'
' Convierte las referencias de notas al pie al texto "NOTA_PIE-numNota"
    ' para poder automatizar externamente su conversión en el .story
'
    Dim lContadorNotas As Long
    Dim bSeguir As Boolean
    Dim oEstiloNota As Font
    Set oEstiloNota = New Font
    
    lContadorNotas = ActiveDocument.Footnotes.StartingNumber
    bSeguir = True

    With oEstiloNota
        .Name = "Swis721 Lt BT"
        .Bold = True
        .Color = -738148353
        .Superscript = True
    End With
    
    Application.Run "RaMacros.LimpiarFindAndReplaceParameters"
    
    Do While bSeguir = True
    
        With ActiveDocument.Range.Find
            .ClearFormatting
            .Text = "^2"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True
            .Replacement.Font = oEstiloNota
            .Replacement.Text = "NOTA_PIE-" & lContadorNotas
            .Execute Replace:=wdReplaceOne
            
            If .Found = True Then
                bSeguir = True
                lContadorNotas = lContadorNotas + 1
            Else
                bSeguir = False
            End If
            
        End With
        
    Loop

    Application.Run "RaMacros.LimpiarFindAndReplaceParameters"
    
End Sub

