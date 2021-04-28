Option Explicit

Sub Iniseg1Limpieza()
' Iniseg1Pre1Limpieza Macro
'
' Crea copia de seguridad del archivo original.
' Ejecuta limpieza (espacios, estilos innecesarios, etc.)
' Crea y deja abierto el archivo en formato libro para comenzar a darle los estilos
'
    Dim dcOriginalFile As Document, dcLibro As Document, stFileName As String, iDeleteAnswer As Integer

    Set dcOriginalFile = ActiveDocument
    stFileName = dcOriginalFile.FullName

	' Borrar contenido innecesario
    iDeleteAnswer = MsgBox("¿Borrar contenido hasta el punto seleccionado?", vbYesNoCancel, "Borrar contenido")
    If iDeleteAnswer = vbYes Then
        RaMacros.CopiaSeguridad dcOriginalFile, "0-", ""
        Selection.HomeKey Unit:=wdStory, Extend:=wdExtend
        Selection.Delete
    ElseIf iDeleteAnswer = vbCancel Then
        Exit Sub
    Else
        RaMacros.CopiaSeguridad dcOriginalFile, "0-", ""
    End If

	' Limpieza y creación de archivo nuevo
    If dcOriginalFile.CompatibilityMode < 15 Then dcOriginalFile.Convert
    Set dcLibro = Documents.Add("C:\Users\Ra\Documents\Plantillas personalizadas de Office\iniseg.dotm")
	Iniseg.InisegHeaderCopy dcOriginalFile, dcLibro, 1
    Iniseg.InisegAutoFormateo dcOriginalFile
    RaMacros.HyperlinksOnlyDomain dcOriginalFile
    RaMacros.LimpiarEspacios dcOriginalFile
    RaMacros.RemoveHeadAndFoot dcOriginalFile
    RaMacros.DeleteUnusedStyles dcOriginalFile

	' Copia de seguridad limpia
    RaMacros.SaveAsNewFile dcOriginalFile, "01-", "", True
    
	' Guarda el archivo con nombre original, preparado para el siguiente paso
    dcLibro.Content.FormattedText = dcOriginalFile.Content
    dcOriginalFile.Close wdDoNotSaveChanges
    dcLibro.SaveAs2 stFileName
    dcLibro.Activate

    Beep
    MsgBox "Corregir numeración de notas al pie, aplicar estilos y ejecutar Iniseg2"
End Sub





Sub Iniseg2LibroYStory()
' Iniseg2LibroYStory Macro
'
' Llama a las macros de InisegLibro e InisegStory y da un aviso para seguir trabajando
    ' Organizado de esta forma las macros de libro y story se pueden llamar por separado
'
    Dim dcLibro As Document, dcStory As Document

    Set dcLibro = InisegLibro(ActiveDocument, True)
	dcLibro.Save

    Set dcStory = InisegStory(dcLibro)
    dcStory.Close wdSaveChanges

    dcLibro.Activate

    Beep
    MsgBox "Revisar formato libro (viudas/huérfanas y tamaño de imágenes tablas) y exportar el material necesario"
End Sub










Function InisegLibro(dcLibro As Document, Optional bCompleto As Boolean = False)
' InisegLibro Function
'
' Realiza la limpieza necesaria y formatea correctamente
'
    RaMacros.LimpiezaBasica dcLibro
    RaMacros.TitulosNoPuntuacionFinal dcLibro
    RaMacros.TitulosQuitarNumeracion dcLibro
    RaMacros.HyperlinksFormatting dcLibro
    Iniseg.InisegComillas dcLibro
    RaMacros.DirectFormattingToStyles dcLibro
    Iniseg.InisegImagenes dcLibro
    Iniseg.InisegInterlineado dcLibro
    RaMacros.LimpiezaBasica dcLibro
    RaMacros.LimpiarFindAndReplaceParameters dcLibro
    Iniseg.InisegParrafosSeparacion dcLibro

    If Not bCompleto Then
        Beep
        MsgBox "Revisar formato libro (viudas/huérfanas, tamaño de imágenes/tablas) y exportar el material necesario"
    End If
    
    Set InisegLibro = dcLibro
End Function










Function InisegStory(dcLibro As Document)
' InisegStory Function
'
' Da el tamaño correcto a párrafos, imágenes y formatea marcas de pie de página
'
    Dim dcStory As Document

    Set dcStory = RaMacros.SaveAsNewFile(dcLibro, "2-", "", False)
    Iniseg.InisegInterlineado dcStory
    RaMacros.ListasATexto dcStory
    Iniseg.ConversionParrafos dcStory
    Iniseg.TitulosConTresEspacios dcStory
    ImagenesGrandes dcStory
    Iniseg.InisegInterlineado dcStory

    Beep
    If MsgBox("¿Procesar notas al pie de página?", vbYesNo, "Notas al pie") = vbYes Then
		Iniseg.InisegNotasPie dcStory
    End If
    
    Set InisegStory = dcStory
End Function











Sub InisegHeaderCopy(dcOriginalDocument As Document, _
                        dcObjectiveDocument As Document, _
                        Optional iHeaderSelection As Integer = 3)
' InisegHeaderCopy Sub
' Copia los encabezados de un archivo a otro según la opción que se le pase:
    ' iHeaderSelection = 1 => copia el encabezado de pág. impar en todos los encabezados
    ' iHeaderSelection = 2 => copia los de pág. impar y par
    ' iHeaderSelection = 3 => respeta el encabezado diferente de la primera pág.
' TODO
    ' GUI para seleccionar qué copiar, cómo y de qué archivo
'
    If iHeaderSelection > 3 Or iHeaderSelection < 1 Then
        Err.Raise Number:=513, Description:="iHeaderSelection out of range"
    End If

    Dim stOriginalHeader As String, iHeader As Integer

    If iHeaderSelection = 1 Then
        stOriginalHeader = Replace(dcOriginalDocument.Sections(1).Headers(1).Range.Text, vbLf, "")
        stOriginalHeader = Trim(Replace(stOriginalHeader, vbCr, ""))
    End If

    For iHeader = 1 To 3
        If iHeaderSelection > 1 Then
            If iHeaderSelection = 2 And iHeader = 2 Then
                stOriginalHeader = Replace(dcOriginalDocument.Sections(1).Headers(1).Range.Text, vbLf, "")
            Else
                stOriginalHeader = Replace(dcOriginalDocument.Sections(1).Headers(iHeader).Range.Text, vbLf, "")
            End If
            stOriginalHeader = Trim(Replace(stOriginalHeader, vbCr, ""))
        End If

        With dcObjectiveDocument.Sections(1).Headers(iHeader).Range.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Text = "Título libro"
            .Replacement.Text = stOriginalHeader
            .Execute Replace:=wdReplaceOne
        End With

        dcObjectiveDocument.Sections(1).Headers(iHeader).Range.Case = wdLowerCase
        dcObjectiveDocument.Sections(1).Headers(iHeader).Range.Case = wdTitleWord
    Next iHeader
End Sub





Sub InisegAutoFormateo(dcArgumentDocument As Document)
'
' InisegAutoFormateo Sub
    ' Convierte las URL de texto plano a hiperenlaces
    ' Da viñeta a las listas que no tienen
    ' Da estilo de lista a las listas
    ' Hace que los paréntesis tengan principio y cierre
    ' Convierte dos guiones seguidos en un guión largo
    '
        ' Cambian cosas que no se pueden desactivar:
            ' Borra párrafos vacíos
' TODO
    ' Recoger y devolver las propiedades con un bucle ForEach, usando ReDim al principio de cada ciclo
'
    Dim optAutoformatValores(14) As Boolean

    With Options

        optAutoformatValores(0) = .AutoFormatApplyBulletedLists
        optAutoformatValores(1) = .AutoFormatApplyFirstIndents
        optAutoformatValores(2) = .AutoFormatApplyHeadings
        optAutoformatValores(3) = .AutoFormatApplyLists
        optAutoformatValores(4) = .AutoFormatApplyOtherParas
        optAutoformatValores(5) = .AutoFormatPlainTextWordMail
        optAutoformatValores(6) = .AutoFormatMatchParentheses
        optAutoformatValores(7) = .AutoFormatPreserveStyles
        optAutoformatValores(8) = .AutoFormatReplaceFarEastDashes
        optAutoformatValores(9) = .AutoFormatReplaceFractions
        optAutoformatValores(10) = .AutoFormatReplaceHyperlinks
        optAutoformatValores(11) = .AutoFormatReplaceOrdinals
        optAutoformatValores(12) = .AutoFormatReplacePlainTextEmphasis
        optAutoformatValores(13) = .AutoFormatReplaceQuotes
        optAutoformatValores(14) = .AutoFormatReplaceSymbols

        .AutoFormatApplyBulletedLists = True
        .AutoFormatApplyFirstIndents = False
        .AutoFormatApplyHeadings = False
        .AutoFormatApplyLists = False
        .AutoFormatApplyOtherParas = False
        .AutoFormatPlainTextWordMail = False
        .AutoFormatMatchParentheses = True
        .AutoFormatPreserveStyles = True
        .AutoFormatReplaceFarEastDashes = False
        .AutoFormatReplaceFractions = False
        .AutoFormatReplaceOrdinals = False
        .AutoFormatReplaceHyperlinks = True
        .AutoFormatReplacePlainTextEmphasis = False
        .AutoFormatReplaceQuotes = False
        .AutoFormatReplaceSymbols = True

        dcArgumentDocument.Range.AutoFormat

        .AutoFormatApplyBulletedLists = optAutoformatValores(0)
        .AutoFormatApplyFirstIndents = optAutoformatValores(1)
        .AutoFormatApplyHeadings = optAutoformatValores(2)
        .AutoFormatApplyLists = optAutoformatValores(3)
        .AutoFormatApplyOtherParas = optAutoformatValores(4)
        .AutoFormatPlainTextWordMail = optAutoformatValores(5)
        .AutoFormatMatchParentheses = optAutoformatValores(6)
        .AutoFormatPreserveStyles = optAutoformatValores(7)
        .AutoFormatReplaceFarEastDashes = optAutoformatValores(8)
        .AutoFormatReplaceFractions = optAutoformatValores(9)
        .AutoFormatReplaceHyperlinks = optAutoformatValores(10)
        .AutoFormatReplaceOrdinals = optAutoformatValores(11)
        .AutoFormatReplacePlainTextEmphasis = optAutoformatValores(12)
        .AutoFormatReplaceQuotes = optAutoformatValores(13)
        .AutoFormatReplaceSymbols = optAutoformatValores(14)
    End With

End Sub





Sub InisegComillas(dcArgumentDocument As Document)
'
' InisegComillas Sub
'
' Quita la negrita y cursiva de las comillas
'
' Basada en RaMacros.ComillasRectasAInglesas
'
    Dim bSmtQt As Boolean
    bSmtQt = Options.AutoFormatAsYouTypeReplaceQuotes
    Options.AutoFormatAsYouTypeReplaceQuotes = True
    RaMacros.LimpiarFindAndReplaceParameters dcArgumentDocument
    
    With dcArgumentDocument.Range.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Replacement.Font.Bold = False
        .Replacement.Font.Italic = False
        .Replacement.Font.Underline = wdUnderlineNone
        .Text = """"
        .Replacement.Text = """"
        .Execute Replace:=wdReplaceAll
        .Text = "'"
        .Replacement.Text = "'"
        .Execute Replace:=wdReplaceAll
    End With
    
    Options.AutoFormatAsYouTypeReplaceQuotes = bSmtQt
    RaMacros.LimpiarFindAndReplaceParameters dcArgumentDocument

End Sub





Sub InisegImagenes(dcArgumentDocument As Document)
'
' InisegImagenes Sub
' Formatea más cómodamente las imágenes
    ' Las convierte de flotantes a inline (de shapes a inlineshapes)
    ' Impide que aparezcan deformadas (mismo % relativo al tamaño original en alto y ancho)
    ' Las centra
    ' Impide que superen el ancho de página
'
    Dim inlShape As InlineShape, shShape As Shape, sngRealPageWidth As Single, sngRealPageHeight As Single, _
        iIndex As Integer

    sngRealPageWidth = dcArgumentDocument.PageSetup.PageWidth - dcArgumentDocument.PageSetup.Gutter _
        - dcArgumentDocument.PageSetup.RightMargin - dcArgumentDocument.PageSetup.LeftMargin

    sngRealPageHeight = dcArgumentDocument.PageSetup.PageHeight _
        - dcArgumentDocument.PageSetup.TopMargin - dcArgumentDocument.PageSetup.BottomMargin _
        - dcArgumentDocument.PageSetup.FooterDistance - dcArgumentDocument.PageSetup.HeaderDistance

    ' Se convierten todas de inlineshapes a shapes
    'For Each inlShape In dcArgumentDocument.InlineShapes
    '    If inlShape.Type = wdInlineShapePicture Then inlShape.ConvertToShape
    'Next inlShape
'
    '' Se les da el formato correcto
    'For Each shShape In dcArgumentDocument.Shapes
    '    With shShape
    '        If .Type = msoPicture Then
    '            shShape.LockAnchor = True
    '            .RelativeVerticalPosition = wdRelativeVerticalPositionParagraph
    '            With .WrapFormat
    '                .AllowOverlap = False
    '                .DistanceTop = 8
    '                .DistanceBottom = 8
    '                .Type = wdWrapTopBottom
    '            End With
    '            .ScaleHeight 1, msoTrue, msoScaleFromBottomRight
    '            .ScaleWidth 1, msoTrue, msoScaleFromBottomRight
    '            .LockAspectRatio = msoTrue
    '            If .Width > sngRealPageWidth Then .Width = sngRealPageWidth
    '            .Left = wdShapeCenter
    '            .Top = 8
    '        End If
    '    End With
    'Next shShape

    ' Se convierten todas de shapes a inlineshapes
    ' For Each shShape In dcArgumentDocument.Shapes
    '     If shShape.Type = msoPicture Then shShape.ConvertToInlineShape
    ' Next shShape


    ' Se convierten todas de shapes a inlineshapes
        
    If dcArgumentDocument.Shapes.Count > 0 Then
    
        For iIndex = 1 To dcArgumentDocument.Shapes.Count
            With dcArgumentDocument.Shapes(iIndex)
                If .Type = msoPicture Then
                    .LockAnchor = True
                    .RelativeVerticalPosition = wdRelativeVerticalPositionParagraph
                    With .WrapFormat
                        .AllowOverlap = False
                        .DistanceTop = 8
                        .DistanceBottom = 8
                        .Type = wdWrapTopBottom
                    End With
                    .ConvertToInlineShape

                    ' Esto a lo mejor es una cagada, pero así evito bucles innecesarios
                    iIndex = iIndex - 1
                    
                End If
                
                If dcArgumentDocument.Shapes.Count = 0 Then Exit For
                
            End With
        Next iIndex
    End If

    ' Se les da el formato correcto
    For Each inlShape In dcArgumentDocument.InlineShapes
        With inlShape
            If .Type = wdInlineShapePicture Then
                .ScaleHeight = .ScaleWidth
                .LockAspectRatio = msoTrue
                ' If .Width / (.ScaleWidth / 100) > sngRealPageWidth Then .Width = sngRealPageWidth Else .ScaleWidth = 100
                .Width = sngRealPageWidth
                If .Height / (.ScaleHeight / 100) > sngRealPageHeight - 15 Then .Height = sngRealPageHeight - 15
                If .Range.Next(Unit:=wdCharacter, Count:=1).Text <> vbCr Then
                    .Range.InsertAfter vbCr
                End If
                ' .Range.InsertAfter vbCr
                ' .Range.Next(Unit:=wdParagraph, Count:=1).Style = wdStyleNormal
                ' .Range.Next(Unit:=wdParagraph, Count:=1).Font.Size = 5
                If .Range.Previous(Unit:=wdCharacter, Count:=1).Text <> vbCr Then
                    .Range.InsertBefore vbCr
                End If
                ' .Range.InsertBefore vbCr
                ' .Range.Previous(Unit:=wdParagraph, Count:=1).Style = wdStyleNormal
                ' .Range.Previous(Unit:=wdParagraph, Count:=1).Font.Size = 5
                .Range.Style = wdStyleNormal
                .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            End If
        End With
    Next inlShape

End Sub





Sub InisegInterlineado(dcArgumentDocument As Document)
'
' InterlineadoSinEspaciado Macro
'
' Interlineado de 1,15 sin espaciado entre párrafos
    ' Eliminar los espaciados verticales entre párrafos y aplica el interlineado correcto

    With dcArgumentDocument.Range.ParagraphFormat
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(1.15)
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
    
    RaMacros.LimpiarFindAndReplaceParameters dcArgumentDocument

End Sub





Sub InisegParrafosSeparacion(dcArgumentDocument As Document)
'
' ParrafosSeparacion Macro
'
' TODO
    ' Refactorizar con variables y recolocando el código
'
    With dcArgumentDocument.Range.Find
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




        .Text = "(*)^13"

        .Style = wdStyleHeading1
        .Replacement.Text = "SEP_11^13\1^13SEP_11^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleHeading2
        .Replacement.Text = "SEP_8^13\1^13SEP_8^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleHeading3
        .Replacement.Text = "SEP_8^13\1^13SEP_8^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleHeading4
        .Replacement.Text = "SEP_6^13\1^13SEP_6^13"
        .Execute Replace:=wdReplaceAll




        .Text = "(^13)"

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleNormal
        .Replacement.Text = "\1SEP_5^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleCaption
        .Replacement.Text = "\1SEP_5^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleQuote
        .Replacement.Text = "\1SEP_5^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleList
        .Replacement.Text = "\1SEP_4^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleList2
        .Replacement.Text = "\1SEP_4^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleList3
        .Replacement.Text = "\1SEP_4^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleListBullet
        .Replacement.Text = "\1SEP_4^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleListBullet2
        .Replacement.Text = "\1SEP_4^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleListBullet3
        .Replacement.Text = "\1SEP_4^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleListBullet4
        .Replacement.Text = "\1SEP_4^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleListBullet5
        .Replacement.Text = "\1SEP_4^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleListNumber
        .Replacement.Text = "\1SEP_4^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleListNumber2
        .Replacement.Text = "\1SEP_4^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleListNumber3
        .Replacement.Text = "\1SEP_4^13"
        .Execute Replace:=wdReplaceAll




        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "(SEP_[0-9]{1;2}^13)"
        .Replacement.Style = wdStyleNormal
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll




        .ClearFormatting
        .Replacement.ClearFormatting
        .Format = False
        .Text = "(SEP_4^13)(SEP_5^13)"
        .Replacement.Text = "\2"
        .Execute Replace:=wdReplaceAll

        .Text = "(SEP_4^13)(SEP_6^13)"
        .Replacement.Text = "\2"
        .Execute Replace:=wdReplaceAll

        .Text = "(SEP_4^13)(SEP_8^13)"
        .Replacement.Text = "\2"
        .Execute Replace:=wdReplaceAll

        .Text = "(SEP_4^13)(SEP_11^13)"
        .Replacement.Text = "\2"
        .Execute Replace:=wdReplaceAll

        .Text = "(SEP_5^13)(SEP_6^13)"
        .Replacement.Text = "\2"
        .Execute Replace:=wdReplaceAll

        .Text = "(SEP_5^13)(SEP_8^13)"
        .Replacement.Text = "\2"
        .Execute Replace:=wdReplaceAll

        .Text = "(SEP_5^13)(SEP_11^13)"
        .Replacement.Text = "\2"
        .Execute Replace:=wdReplaceAll

        .Text = "(SEP_6^13)(SEP_6^13)"
        .Replacement.Text = "\2"
        .Execute Replace:=wdReplaceAll

        .Text = "(SEP_6^13)(SEP_8^13)"
        .Replacement.Text = "\2"
        .Execute Replace:=wdReplaceAll

        .Text = "(SEP_6^13)(SEP_11^13)"
        .Replacement.Text = "\2"
        .Execute Replace:=wdReplaceAll

        .Text = "(SEP_8^13)(SEP_8^13)"
        .Replacement.Text = "\2"
        .Execute Replace:=wdReplaceAll

        .Text = "(SEP_8^13)(SEP_11^13)"
        .Replacement.Text = "\2"
        .Execute Replace:=wdReplaceAll

        .Text = "(SEP_11^13)(SEP_11^13)"
        .Replacement.Text = "\2"
        .Execute Replace:=wdReplaceAll

        .Text = "(SEP_11^13)(SEP_8^13)"
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll

        .Text = "(SEP_11^13)(SEP_6^13)"
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll

        .Text = "(SEP_8^13)(SEP_6^13)"
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll




        .ClearFormatting
        .Replacement.ClearFormatting
        .Format = True
        .Replacement.Font.Size = 4
        .Text = "SEP_4(^13)"
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Replacement.Font.Size = 5
        .Text = "SEP_5(^13)"
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Replacement.Font.Size = 6
        .Text = "SEP_6(^13)"
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Replacement.Font.Size = 8
        .Text = "SEP_8(^13)"
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Replacement.Font.Size = 11
        .Text = "SEP_11(^13)"
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll
    End With
    
    RaMacros.LimpiarFindAndReplaceParameters dcArgumentDocument

End Sub





Sub ConversionParrafos(dcArgumentDocument As Document)
'
' Iniseg4ConversionParrafos Macro
' Conversion de Word impreso a formato para Storyline
'
    ' Cambio del tamaño de Titulo 1 de 16 a 17
    
    RaMacros.LimpiarFindAndReplaceParameters dcArgumentDocument
    
    With dcArgumentDocument.Styles(wdStyleHeading1).Font
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

    ' Eliminar ALLCAPS de los títulos 2 y 3
    dcArgumentDocument.Styles(wdStyleHeading2).Font.AllCaps = False
    dcArgumentDocument.Styles(wdStyleHeading3).Font.AllCaps = False

    ' Poner el estilo quote centrado y sin espacio a derecha ni izquierda
    With dcArgumentDocument.Styles(wdStyleQuote).ParagraphFormat
        .LeftIndent = 0
        .RightIndent = 0
        .Alignment = wdAlignParagraphCenter
    End With

' Cambio de tamaño de parrafos de separacion

    ' Listas: 4 a 2
    With dcArgumentDocument.Range.Find
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
    With dcArgumentDocument.Range.Find
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
    With dcArgumentDocument.Range.Find
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
    With dcArgumentDocument.Range.Find
        .ClearFormatting
        .Style = wdStyleHeading1
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

        .Style = wdStyleNormal
        .Text = "FISTRO^13"
        .Replacement.Font.Size = 8
        .Replacement.Text = "^13"
        .Execute Replace:=wdReplaceAll
    End With

    ' Meter salto de página antes de cada Heading 1 y Title
    Dim prParrafoActual As Paragraph, index As Integer

    For index = 1 To dcArgumentDocument.Paragraphs.Count - 1
        With dcArgumentDocument.Paragraphs(index).Range
            If .Next(Unit:=wdParagraph, Count:=1).Paragraphs(1).OutlineLevel = 1 Then
                If .Previous(Unit:=wdParagraph, Count:=2).Style <> wdStyleTitle Then
                    '.Collapse Direction:=wdCollapseEnd
                    dcArgumentDocument.Paragraphs(index).Range.InsertBreak Type:=wdPageBreak
                    'index = index + 1
                End If
            End If
        End With
    Next index

    RaMacros.LimpiarFindAndReplaceParameters dcArgumentDocument
    
End Sub





Sub ImagenesGrandes(dcArgumentDocument As Document)
'
' ImagenesGrandes Sub
'
' Hace que todas las imágenes sean enormes, para meterlas en el story
'
    Dim inlShape As InlineShape
        
    For Each inlShape In dcArgumentDocument.InlineShapes
        If inlShape.Type = wdInlineShapePicture Then inlShape.Width = CentimetersToPoints(29)
    Next inlShape

End Sub





Sub TitulosConTresEspacios(dcArgumentDocument As Document)
'
' TitulosConTresEspacios Sub
'
' Sustituye la tabulación en los títulos por 3 espacios
'
    With dcArgumentDocument.Range.Find
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
        .Style = wdStyleHeading1
        .Execute Replace:=wdReplaceAll
        .Style = wdStyleHeading2
        .Execute Replace:=wdReplaceAll
        .Style = wdStyleHeading3
        .Execute Replace:=wdReplaceAll

    End With

    RaMacros.LimpiarFindAndReplaceParameters dcArgumentDocument

End Sub





Sub InisegNotasPie(dcArgumentDocument As Document)
'
' NotasPieATexto Sub
'
' Convierte las referencias de notas al pie al texto "NOTA_PIE-numNota"
    ' para poder automatizar externamente su conversión en el .story
'
    Dim lContadorNotas As Long
    Dim bSeguir As Boolean
    Dim oEstiloNota As Font
    Set oEstiloNota = New Font
    
    lContadorNotas = dcArgumentDocument.Footnotes.StartingNumber
    bSeguir = True

    With oEstiloNota
        .Name = "Swis721 Lt BT"
        .Bold = True
        .Color = -738148353
        .Superscript = True
    End With
    
    RaMacros.LimpiarFindAndReplaceParameters dcArgumentDocument
    
    Do While bSeguir = True
    
        With dcArgumentDocument.Range.Find
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

    RaMacros.LimpiarFindAndReplaceParameters dcArgumentDocument
    
End Sub





