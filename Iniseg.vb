Option Explicit

Sub uiInisegConversionLibro()
    ConversionLibro ActiveDocument, False
End Sub

Sub uiInisegConversionStory()
    ConversionStory ActiveDocument
End Sub
    
Sub Iniseg1Limpieza()
' Crea copia de seguridad del archivo original.
' Ejecuta limpieza (espacios, estilos innecesarios, etc.)
' Crea y deja abierto el archivo en formato libro para comenzar a darle los estilos
'
    Dim dcOriginalFile As Document, dcLibro As Document, stFileName As String, iDeleteAnswer As Integer, lEstilosBorrados As Long

    Set dcOriginalFile = ActiveDocument
    stFileName = dcOriginalFile.FullName

	' Borrar contenido innecesario
    iDeleteAnswer = MsgBox("¿Borrar contenido hasta el punto seleccionado?", vbYesNoCancel, "Borrar contenido")
    If iDeleteAnswer = vbYes Then
        RaMacros.CopySecurity dcOriginalFile, "0-", ""
        Selection.HomeKey Unit:=wdStory, Extend:=wdExtend
        Selection.Delete
    ElseIf iDeleteAnswer = vbCancel Then
        Exit Sub
    Else
        RaMacros.CopySecurity dcOriginalFile, "0-", ""
    End If

	' Limpieza y creación de archivo nuevo
    If dcOriginalFile.CompatibilityMode < 15 Then dcOriginalFile.Convert
    Set dcLibro = Documents.Add("C:\Users\Ra\Documents\Plantillas personalizadas de Office\iniseg.dotm")
	Iniseg.HeaderCopy dcOriginalFile, dcLibro, 1
    Iniseg.AutoFormateo dcOriginalFile
    RaMacros.HyperlinksOnlyDomain dcOriginalFile
    RaMacros.CleanSpaces dcOriginalFile
    RaMacros.HeadersFootersRemove dcOriginalFile
    lEstilosBorrados = RaMacros.StylesDeleteUnused(dcOriginalFile, False)

	' Copia de seguridad limpia
    RaMacros.SaveAsNewFile dcOriginalFile, "01-", "", True
    
	' Guarda el archivo con nombre original, preparado para el siguiente paso
    dcLibro.Content.FormattedText = dcOriginalFile.Content
    dcOriginalFile.Close wdDoNotSaveChanges
    dcLibro.SaveAs2 stFileName
    dcLibro.Activate

    Beep
    MsgBox lEstilosBorrados & " Estilos borrados" & vbCrLf & "Revisar numeración de notas al pie, aplicar estilos y ejecutar Iniseg2"
End Sub





Sub Iniseg2LibroYStory()
' Llama a las macros de ConversionLibro e ConversionStory y da un aviso para seguir trabajando
    ' Organizado de esta forma las macros de libro y story se pueden llamar por separado
'
    Dim dcLibro As Document, dcStory As Document

    Set dcLibro = Iniseg.ConversionLibro(ActiveDocument, True)
	dcLibro.Save

    Iniseg.ConversionStory dcLibro

    dcLibro.Activate

    Beep
    MsgBox "Revisar formato libro (viudas/huérfanas, tamaño de imágenes o tablas...), exportar material necesario y ejecutar iniseg 3"
End Sub







Sub Iniseg3PáginasBlancasVisibles()
	' Esta macro es una mala práctica y solo está para evitar confusiones por
		' falta de uniformidad en el uso de plantillas y estilos
    RaMacros.SectionsFillBlankPages ActiveDocument
End Sub









Function ConversionLibro(dcLibro As Document, Optional bCompleto As Boolean = False)
' Realiza la limpieza necesaria y formatea correctamente
'
	RaMacros.SaveAsNewFile dcLibro, "1-", "", True
    RaMacros.CleanBasic dcLibro
    
	RaMacros.HeadingsNoPunctuation dcLibro
    RaMacros.HeadingsNoNumeration dcLibro
	' Títulos con mayúsculas tipo título
	dcLibro.Styles(wdstyleheading1).Font.AllCaps = False
	RaMacros.HeadingsChangeCase dcLibro, 0, 4
	dcLibro.Styles(wdstyleheading1).Font.AllCaps = True
    
	RaMacros.HyperlinksFormatting dcLibro
    Iniseg.ComillasFormato dcLibro
    RaMacros.StylesNoDirectFormatting dcLibro
    
	Iniseg.ImagenesLibro dcLibro
    
	Iniseg.InterlineadoCorregido dcLibro
    RaMacros.CleanBasic dcLibro
    
	Iniseg.ParrafosSeparacionLibro dcLibro
	RaMacros.SectionBreakBeforeHeading dcLibro, False, 4, 1

    If Not bCompleto Then
        Beep
        MsgBox "Revisar formato libro (viudas/huérfanas, tamaño de imágenes/tablas), activar Iniseg3 y exportar el material necesario"
    End If
    
    Set ConversionLibro = dcLibro
End Function










Sub ConversionStory(dcLibro As Document)
' Da el tamaño correcto a párrafos, imágenes y formatea marcas de pie de página
'
    Dim dcStory As Document

    Set dcStory = RaMacros.SaveAsNewFile(dcLibro, "2-", "", False)
    Iniseg.InterlineadoCorregido dcStory
    RaMacros.ListsToText dcStory
    Iniseg.ParrafosConversionStory dcStory
    Iniseg.TitulosConTresEspacios dcStory
    Iniseg.ImagenesStory dcStory
    Iniseg.InterlineadoCorregido dcStory

    Beep
    If MsgBox("¿Procesar notas al pie de página?", vbYesNo, "Notas al pie") = vbYes Then
		Iniseg.NotasPieMarcadores dcStory
    End If
    
	If dcStory.Sections.Count > 1 Then
		RaMacros.SectionsExportEachToFiles dcStory,, "-tema_"
	End If

    dcStory.Close wdSaveChanges
End Sub











Sub HeaderCopy(dcOriginalDocument As Document, _
                dcObjectiveDocument As Document, _
                Optional iHeaderSelection As Integer = 3)
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





Sub AutoFormateo(dcArgumentDocument As Document)
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





Sub ComillasFormato(dcArgumentDocument As Document)
' Quita la negrita y cursiva de las comillas y las pasa a curvadas
'
' Basada en RaMacros.QuotesStraightToCurly
'
    Dim bSmtQt As Boolean
    bSmtQt = Options.AutoFormatAsYouTypeReplaceQuotes
    Options.AutoFormatAsYouTypeReplaceQuotes = True
    RaMacros.FindAndReplaceClearParameters
    
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
    RaMacros.FindAndReplaceClearParameters

End Sub





Sub ImagenesLibro(dcArgumentDocument As Document)
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





Sub ImagenesStory(dcArgumentDocument As Document)
' Hace que todas las imágenes sean enormes, para meterlas en el story
'
    Dim inlShape As InlineShape
        
    For Each inlShape In dcArgumentDocument.InlineShapes
        If inlShape.Type = wdInlineShapePicture Then inlShape.Width = CentimetersToPoints(29)
    Next inlShape

End Sub





Sub InterlineadoCorregido(dcArgumentDocument As Document)
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
    
    RaMacros.FindAndReplaceClearParameters

End Sub





Sub ParrafosSeparacionLibro(dcArgumentDocument As Document)
' TODO
    ' Refactorizar con variables y recolocando el código
'
	' Mete dos saltos de línea manuales en los Heading 1, entre "Tema N" y el nombre del tema
    With dcArgumentDocument.Range.Find
		.ClearFormatting
		.Replacement.ClearFormatting
		.Forward = True
		.Wrap = wdFindContinue
		.Format = False
		.MatchCase = False
		.MatchWholeWord = False
		.MatchWildcards = True
		.MatchSoundsLike = False
		.MatchAllWordForms = False

		' Elimina saltos manuales de página (innecesarios con los saltos de sección y revisión posteriores)
		.Text = "^m"
		.Replacement.Text = ""
		.Execute Replace:=wdReplaceAll

		.Format = True
		.style = wdstyleheading1
		.Text = "([tT][eE][mM][aA] [0-9]{1;2})"
		.Replacement.Text = "\1^l^l"
		.Execute Replace:=wdReplaceAll
    End With

		' Formatea los saltos de línea y les da tamaño 10
	RaMacros.CleanSpaces dcArgumentDocument
    With dcArgumentDocument.Range.Find
       .ClearFormatting
       .Replacement.ClearFormatting
       .Forward = True
       .Wrap = wdFindContinue
       .Format = True
       .MatchCase = False
       .MatchWholeWord = False
       .MatchWildcards = True
       .MatchSoundsLike = False
       .MatchAllWordForms = False
	   .style = wdstyleheading1
       .Replacement.ClearFormatting
       .Replacement.Font.Size = 10
       .Text = "[^13^l]{2;}"
       .Replacement.Text = "^l^l"
	   .Execute Replace:=wdReplaceAll
	End With


	' Párrafos de separación generales
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

		' Marcas para los títulos
        .Text = "(*)^13"

        .Style = wdstyleheading1
        .Replacement.Text = "\1^13SEP_11^13"
        .Execute Replace:=wdReplaceAll

        .Style = wdstyleheading2
        .Replacement.Text = "SEP_11^13\1^13SEP_11^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdstyleheading3
        .Replacement.Text = "SEP_8^13\1^13SEP_8^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdstyleheading4
        .Replacement.Text = "SEP_8^13\1^13SEP_8^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleHeading5
        .Replacement.Text = "SEP_6^13\1^13SEP_6^13"
        .Execute Replace:=wdReplaceAll



		' Marcas del resto de estilos

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleNormal
        .Replacement.Text = "SEP_5^13\1^13SEP_5^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleCaption
        .Replacement.Text = "SEP_5^13\1^13SEP_5^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleQuote
        .Replacement.Text = "SEP_5^13\1^13SEP_5^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleList
        .Replacement.Text = "SEP_4^13\1^13SEP_4^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleList2
        .Replacement.Text = "SEP_4^13\1^13SEP_4^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleList3
        .Replacement.Text = "SEP_4^13\1^13SEP_4^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleListBullet
        .Replacement.Text = "SEP_4^13\1^13SEP_4^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleListBullet2
        .Replacement.Text = "SEP_4^13\1^13SEP_4^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleListBullet3
        .Replacement.Text = "SEP_4^13\1^13SEP_4^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleListBullet4
        .Replacement.Text = "SEP_4^13\1^13SEP_4^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleListBullet5
        .Replacement.Text = "SEP_4^13\1^13SEP_4^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleListContinue
        .Replacement.Text = "SEP_4^13\1^13SEP_4^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleListContinue2
        .Replacement.Text = "SEP_4^13\1^13SEP_4^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleListContinue3
        .Replacement.Text = "SEP_4^13\1^13SEP_4^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleListContinue4
        .Replacement.Text = "SEP_4^13\1^13SEP_4^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleListContinue5
        .Replacement.Text = "SEP_4^13\1^13SEP_4^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleListNumber
        .Replacement.Text = "SEP_4^13\1^13SEP_4^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleListNumber2
        .Replacement.Text = "SEP_4^13\1^13SEP_4^13"
        .Execute Replace:=wdReplaceAll

        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = wdStyleListNumber3
        .Replacement.Text = "SEP_4^13\1^13SEP_4^13"
        .Execute Replace:=wdReplaceAll



		' Convertir marcas a estilo Normal
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "(SEP_[0-9]{1;2}^13)"
        .Replacement.Style = wdStyleNormal
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll



		' Seleccionar la marca de mayor tamaño, cuando coinciden 2
        .ClearFormatting
        .Replacement.ClearFormatting
        .Format = False

		' Prevalece la segunda
        .Text = "(SEP_4^13)(SEP_4^13)"
        .Replacement.Text = "\2"
        .Execute Replace:=wdReplaceAll

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

        .Text = "(SEP_5^13)(SEP_5^13)"
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


		' Prevalece la primera
        .Text = "(SEP_11^13)(SEP_8^13)"
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll

        .Text = "(SEP_11^13)(SEP_6^13)"
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll

        .Text = "(SEP_11^13)(SEP_5^13)"
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll

        .Text = "(SEP_11^13)(SEP_4^13)"
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll

        .Text = "(SEP_8^13)(SEP_6^13)"
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll

        .Text = "(SEP_8^13)(SEP_5^13)"
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll

        .Text = "(SEP_8^13)(SEP_4^13)"
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll

        .Text = "(SEP_6^13)(SEP_5^13)"
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll

        .Text = "(SEP_6^13)(SEP_4^13)"
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll

        .Text = "(SEP_5^13)(SEP_4^13)"
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll



		' Redimensionado de párrafo y borrado de marca
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
    RaMacros.FindAndReplaceClearParameters
End Sub





Sub ParrafosConversionStory(dcArgumentDocument As Document)
' Conversion de Word impreso a formato para Storyline
'
    RaMacros.FindAndReplaceClearParameters

	' Títulos 1 en minúsculas (para copiarlos más cómodamente a la primera diapositiva)
	dcArgumentDocument.Styles(wdstyleheading1).Font.AllCaps = False

    ' Cambio del tamaño de Titulo 2 de 16 a 17
    With dcArgumentDocument.Styles(wdstyleheading2).Font
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

    ' Eliminar ALLCAPS de los títulos 3 y 4
    dcArgumentDocument.Styles(wdstyleheading3).Font.AllCaps = False
    dcArgumentDocument.Styles(wdstyleheading4).Font.AllCaps = False

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

    ' Titulos 3, 4 y 5: 8 a 6.
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



' Titulos 2: 11 a 8
    ' Dar tamaño 8 a todos los párrafos tras los Heading 2
    With dcArgumentDocument.Range.Find
        .ClearFormatting
        .Style = wdstyleheading2
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

    RaMacros.FindAndReplaceClearParameters
End Sub





Sub TitulosConTresEspacios(dcArgumentDocument As Document)
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
        .Style = wdstyleheading2
        .Execute Replace:=wdReplaceAll
        .Style = wdstyleheading3
        .Execute Replace:=wdReplaceAll
        .Style = wdstyleheading4
        .Execute Replace:=wdReplaceAll

    End With

    RaMacros.FindAndReplaceClearParameters

End Sub





Sub NotasPieMarcadores(dcArgumentDocument As Document)
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
    
    RaMacros.FindAndReplaceClearParameters
    
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

    RaMacros.FindAndReplaceClearParameters
    
End Sub





