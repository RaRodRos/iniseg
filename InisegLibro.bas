Attribute VB_Name = "InisegLibro"
Option Explicit

Private Function InisegAutoFormateo()
'
' InisegAutoFormateo Function
' Convierte las URL de texto plano a hiperenlaces
' Da viñeta a las listas que no tienen
' Da estilo de lista a las listas
' Hace que los paréntesis tengan principio y cierre
' Convierte dos guiones seguidos en un guión largo
'
    ' Cambian cosas que no se pueden desactivar:
        ' Borra párrafos vacíos
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

        ActiveDocument.Range.AutoFormat

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

End Function

Private Function InisegInterlineado()
'
' InterlineadoSinEspaciado Macro
'
' Interlineado de 1,15 sin espaciado entre párrafos
    ' Eliminar los espaciados verticales entre párrafos y aplica el interlineado correcto

    With ActiveDocument.Range.ParagraphFormat
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(1.15)
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
    
    Application.Run "RaMacros.LimpiarFindAndReplaceParameters"

End Function

Private Function InisegComillas()
'
' InisegComillas Function
'
' Quita la negrita y cursiva de las comillas
'
' Basada en RaMacros.ComillasRectasAInglesas
'
    Dim bSmtQt As Boolean
    bSmtQt = Options.AutoFormatAsYouTypeReplaceQuotes
    Options.AutoFormatAsYouTypeReplaceQuotes = True
    Application.Run "RaMacros.LimpiarFindAndReplaceParameters"
    
    With ActiveDocument.Range.Find
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
    Application.Run "RaMacros.LimpiarFindAndReplaceParameters"

End Function

Private Function InisegImagenes()
'
' InisegImagenes Function
' Formatea más cómodamente las imágenes
    ' Las convierte de flotantes a inline (de shapes a inlineshapes)
    ' Impide que aparezcan deformadas (mismo % relativo al tamaño original en alto y ancho)
    ' Las centra
    ' Impide que superen el ancho de página
'
    Dim inlShape As InlineShape, shShape As Shape, sngRealPageWidth As Single, sngRealPageHeight As Single, _
        iIndex As Integer

    sngRealPageWidth = ActiveDocument.PageSetup.PageWidth - ActiveDocument.PageSetup.Gutter _
        - ActiveDocument.PageSetup.RightMargin - ActiveDocument.PageSetup.LeftMargin

    sngRealPageHeight = ActiveDocument.PageSetup.PageHeight _
        - ActiveDocument.PageSetup.TopMargin - ActiveDocument.PageSetup.BottomMargin _
        - ActiveDocument.PageSetup.FooterDistance - ActiveDocument.PageSetup.HeaderDistance

    ' Se convierten todas de inlineshapes a shapes
    'For Each inlShape In ActiveDocument.InlineShapes
    '    If inlShape.Type = wdInlineShapePicture Then inlShape.ConvertToShape
    'Next inlShape
'
    '' Se les da el formato correcto
    'For Each shShape In ActiveDocument.Shapes
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
    ' For Each shShape In ActiveDocument.Shapes
    '     If shShape.Type = msoPicture Then shShape.ConvertToInlineShape
    ' Next shShape


    ' Se convierten todas de shapes a inlineshapes
        
    If ActiveDocument.Shapes.Count > 0 Then
    
        For iIndex = 1 To ActiveDocument.Shapes.Count
            With ActiveDocument.Shapes(iIndex)
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
                
                If ActiveDocument.Shapes.Count = 0 Then Exit For
                
            End With
        Next iIndex
    End If

    ' Se les da el formato correcto
    For Each inlShape In ActiveDocument.InlineShapes
        With inlShape
            If .Type = wdInlineShapePicture Then
                .ScaleHeight = .ScaleWidth
                .LockAspectRatio = msoTrue
                If .Width / (.ScaleWidth / 100) > sngRealPageWidth Then .Width = sngRealPageWidth Else .ScaleWidth = 100
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

End Function

Sub Iniseg1Limpieza()
'
' Iniseg1Pre1Limpieza Macro
'
' Ejecuta limpieza de espacios innecesarios y estilos:
'
    Application.Run "InisegLibro.InisegAutoFormateo"
    Application.Run "RaMacros.HyperlinksOnlyDomain"
    Application.Run "RaMacros.LimpiarEspacios"
    Application.Run "RaMacros.LimpiarFindAndReplaceParameters"
    Application.Run "RaMacros.RemoveHeadAndFoot"
    Application.Run "DeleteUnusedStyles.DeleteUnusedStyles"

End Sub

Sub Iniseg2Formatos()
'
' Iniseg2Limpieza Macro
'
' Limpia puntos y formatea correctamente marcas a pie de página e hiperenlaces
' TODO:
    ' formatear marcas a pie de página
    ' Formatear hipervínculos

    Application.Run "RaMacros.TitulosQuitarPuntacionFinal"
    Application.Run "RaMacros.TitulosQuitarNumeracion"
    Application.Run "RaMacros.LimpiezaBasica"
    Application.Run "RaMacros.HipervinculosFormatear"
    Application.Run "InisegLibro.InisegComillas"
    Application.Run "RaMacros.DirectFormattingToStyles"
    Application.Run "InisegLibro.InisegImagenes"
    Application.Run "InisegLibro.InisegInterlineado"
    Application.Run "RaMacros.NoParrafosVacios"
    Application.Run "RaMacros.LimpiarFindAndReplaceParameters"
    
End Sub

Sub Iniseg3ParrafosSeparacion()
'
' ParrafosSeparacion Macro
'
    Application.Run "RaMacros.LimpiarFindAndReplaceParameters"
    
    With ActiveDocument.Range.Find
        .ClearFormatting
        .Style = ActiveDocument.Styles(wdStyleHeading1)
        .Replacement.ClearFormatting
        .Text = "(*)^13"
        .Replacement.Text = "SEP_11^13\1^13SEP_11^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .ClearFormatting
        .Style = ActiveDocument.Styles(wdStyleHeading2)
        .Replacement.ClearFormatting
        .Text = "(*)^13"
        .Replacement.Text = "SEP_8^13\1^13SEP_8^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .ClearFormatting
        .Style = ActiveDocument.Styles(wdStyleHeading3)
        .Replacement.ClearFormatting
        .Text = "(*)^13"
        .Replacement.Text = "SEP_8^13\1^13SEP_8^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .ClearFormatting
        .Style = ActiveDocument.Styles(wdStyleHeading4)
        .Replacement.ClearFormatting
        .Text = "(*)^13"
        .Replacement.Text = "SEP_6^13\1^13SEP_6^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .ClearFormatting
        .Style = ActiveDocument.Styles(wdStyleNormal)
        .Replacement.ClearFormatting
        .Text = "(^13)"
        .Replacement.Text = "\1SEP_5^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .ClearFormatting
        .Style = ActiveDocument.Styles(wdStyleCaption)
        .Replacement.ClearFormatting
        .Text = "(^13)"
        .Replacement.Text = "\1SEP_5^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .ClearFormatting
        .Style = ActiveDocument.Styles(wdStyleQuote)
        .Replacement.ClearFormatting
        .Text = "(^13)"
        .Replacement.Text = "\1SEP_5^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .ClearFormatting
        .Style = ActiveDocument.Styles(wdStyleList)
        .Replacement.ClearFormatting
        .Text = "(^13)"
        .Replacement.Text = "\1SEP_4^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .ClearFormatting
        .Style = ActiveDocument.Styles(wdStyleList2)
        .Replacement.ClearFormatting
        .Text = "(^13)"
        .Replacement.Text = "\1SEP_4^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .ClearFormatting
        .Style = ActiveDocument.Styles(wdStyleList3)
        .Replacement.ClearFormatting
        .Text = "(^13)"
        .Replacement.Text = "\1SEP_4^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .ClearFormatting
        .Style = ActiveDocument.Styles(wdStyleListBullet)
        .Replacement.ClearFormatting
        .Text = "(^13)"
        .Replacement.Text = "\1SEP_4^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .ClearFormatting
        .Style = ActiveDocument.Styles(wdStyleListBullet2)
        .Replacement.ClearFormatting
        .Text = "(^13)"
        .Replacement.Text = "\1SEP_4^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .ClearFormatting
        .Style = ActiveDocument.Styles(wdStyleListBullet3)
        .Replacement.ClearFormatting
        .Text = "(^13)"
        .Replacement.Text = "\1SEP_4^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .ClearFormatting
        .Style = ActiveDocument.Styles(wdStyleListBullet4)
        .Replacement.ClearFormatting
        .Text = "(^13)"
        .Replacement.Text = "\1SEP_4^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .ClearFormatting
        .Style = ActiveDocument.Styles(wdStyleListBullet5)
        .Replacement.ClearFormatting
        .Text = "(^13)"
        .Replacement.Text = "\1SEP_4^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .ClearFormatting
        .Style = ActiveDocument.Styles(wdStyleListNumber)
        .Replacement.ClearFormatting
        .Text = "(^13)"
        .Replacement.Text = "\1SEP_4^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .ClearFormatting
        .Style = ActiveDocument.Styles(wdStyleListNumber2)
        .Replacement.ClearFormatting
        .Text = "(^13)"
        .Replacement.Text = "\1SEP_4^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .ClearFormatting
        .Style = ActiveDocument.Styles(wdStyleListNumber3)
        .Replacement.ClearFormatting
        .Text = "(^13)"
        .Replacement.Text = "\1SEP_4^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Replacement.Style = ActiveDocument.Styles(wdStyleNormal)
        .Text = "(SEP_[0-9]{1;2}^13)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "(SEP_4^13)(SEP_5^13)"
        .Replacement.Text = "\2"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .Text = "(SEP_4^13)(SEP_6^13)"
        .Replacement.Text = "\2"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .Text = "(SEP_4^13)(SEP_8^13)"
        .Replacement.Text = "\2"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .Text = "(SEP_4^13)(SEP_11^13)"
        .Replacement.Text = "\2"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .Text = "(SEP_5^13)(SEP_6^13)"
        .Replacement.Text = "\2"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .Text = "(SEP_5^13)(SEP_8^13)"
        .Replacement.Text = "\2"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .Text = "(SEP_5^13)(SEP_11^13)"
        .Replacement.Text = "\2"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .Text = "(SEP_6^13)(SEP_6^13)"
        .Replacement.Text = "\2"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .Text = "(SEP_6^13)(SEP_8^13)"
        .Replacement.Text = "\2"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .Text = "(SEP_6^13)(SEP_11^13)"
        .Replacement.Text = "\2"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .Text = "(SEP_8^13)(SEP_8^13)"
        .Replacement.Text = "\2"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .Text = "(SEP_8^13)(SEP_11^13)"
        .Replacement.Text = "\2"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .Text = "(SEP_11^13)(SEP_11^13)"
        .Replacement.Text = "\2"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "(SEP_11^13)(SEP_8^13)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .Text = "(SEP_11^13)(SEP_6^13)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .Text = "(SEP_8^13)(SEP_6^13)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Replacement.Font.Size = 4
        .Text = "SEP_4(^13)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Replacement.Font.Size = 5
        .Text = "SEP_5(^13)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Replacement.Font.Size = 6
        .Text = "SEP_6(^13)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Replacement.Font.Size = 8
        .Text = "SEP_8(^13)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Range.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Replacement.Font.Size = 11
        .Text = "SEP_11(^13)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    
    Application.Run "RaMacros.LimpiarFindAndReplaceParameters"

End Sub



