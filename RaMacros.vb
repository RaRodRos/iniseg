Option Explicit

Function StyleInUse(Styname As String, dcArgumentDocument As Document) As Boolean
' StyleInUse Function
' Del mismo desarrollador que DeleteUnusedStyles
' Is Stryname used any of dcArgumentDocument's story
    Dim Stry As Range
    Dim Shp As Shape
    Dim txtFrame As TextFrame

    If Not dcArgumentDocument.Styles(Styname).InUse Then StyleInUse = False: Exit Function
    ' check if Currently used in a story?

    For Each Stry In dcArgumentDocument.StoryRanges
        If StoryInUse(Stry, dcArgumentDocument) Then
            If StyleInUseInRangeText(Stry, Styname) Then StyleInUse = True: Exit Function
            For Each Shp In Stry.ShapeRange
                Set txtFrame = Shp.TextFrame
                If Not txtFrame Is Nothing Then
                    If txtFrame.HasText Then
                        If txtFrame.TextRange.Characters.Count > 1 Then
                            If StyleInUseInRangeText(txtFrame.TextRange, Styname) Then
                                StyleInUse = True: Exit Function
                            End If
                        End If
                    End If
                End If
            Next Shp
        End If
    Next Stry
    StyleInUse = False ' Not currently in use.
End Function

Function StyleInUseInRangeText(rng As Range, Styname As String) As Boolean
' StyleInUseInRangeText Function
' Del mismo desarrollador que DeleteUnusedStyles
' Returns True if "Styname" is use in rng
    With rng.Find
        .ClearFormatting
        .ClearHitHighlight
        .Style = Styname
        .Format = True
        .Text = ""
        .Replacement.Text = ""
        .Wrap = wdFindContinue
        StyleInUseInRangeText = .Execute
    End With
End Function

Function StoryInUse(Stry As Range, dcArgumentDocument As Document) As Boolean
' StoryInUse Function
' Del mismo desarrollador que DeleteUnusedStyles
' Note: this will mark even the always-existing stories as not in use if they're empty
    If Not Stry.StoryLength > 1 Then StoryInUse = False: Exit Function
    Select Case Stry.StoryType
        Case wdMainTextStory, wdPrimaryFooterStory, wdPrimaryHeaderStory: StoryInUse = True
        Case wdEvenPagesFooterStory, wdEvenPagesHeaderStory: StoryInUse = Stry.Sections(1).PageSetup.OddAndEvenPagesHeaderFooter = True
        Case wdFirstPageFooterStory, wdFirstPageHeaderStory: StoryInUse = Stry.Sections(1).PageSetup.DifferentFirstPageHeaderFooter = True
        Case wdFootnotesStory, wdFootnoteContinuationSeparatorStory: StoryInUse = dcArgumentDocument.Footnotes.Count > 1
        Case wdFootnoteSeparatorStory, wdFootnoteContinuationNoticeStory: StoryInUse = dcArgumentDocument.Footnotes.Count > 1
        Case wdEndnotesStory, wdEndnoteContinuationSeparatorStory: StoryInUse = dcArgumentDocument.Endnotes.Count > 1
        Case wdEndnoteSeparatorStory, wdEndnoteContinuationNoticeStory: StoryInUse = dcArgumentDocument.Endnotes.Count > 1
        Case wdCommentsStory: StoryInUse = dcArgumentDocument.Comments.Count > 1
        Case wdTextFrameStory: StoryInUse = dcArgumentDocument.Frames.Count > 1
        Case Else: StoryInUse = False ' Must be some new or unknown wdStoryType
    End Select
End Function

Sub DeleteUnusedStyles(dcArgumentDocument As Document)
' DeleteUnusedStyles Sub
    ' Borra todos los estilos no usados, pero hay que dar varias pasadas para que quite todos porque respeta la
        ' jerarquía de estilos y los elimina en orden (para evitar que se borren padres sin uso,
        ' como podrían ser las listas)
' Está en los comentarios de
    ' https://word.tips.net/T001337_Removing_Unused_Styles.html
' ANTES USABA ESTA
    ' https://www.msofficeforums.com/word-vba/37489-macro-deletes-unused-styles.html
' Lo He modificado ligeramente para que no haya que dar varias pasadas a mano y recorra el documento hasta que no
    ' queden estilos que borrar
'
    Dim oStyle As Style
    Dim sCount As Long
    Dim lTotalSCount As Long

    lTotalSCount = 0
    Do
        sCount = 0
        For Each oStyle In dcArgumentDocument.Styles
            'Only check out non-built-in styles
            If oStyle.BuiltIn = False Then
                If StyleInUse(oStyle.NameLocal, dcArgumentDocument) = False Then
                    Application.OrganizerDelete Source:= dcArgumentDocument.FullName, _
                    Name:= oStyle.NameLocal, Object:=wdOrganizerObjectStyles
                    sCount = sCount + 1
                End If
            End If
        Next oStyle
        lTotalSCount = lTotalSCount + sCount
    Loop While sCount > 0

    If lTotalSCount <> 0 Then MsgBox lTotalSCount & " styles deleted"

End Sub






Sub CopiaSeguridad(dcArgumentDocument As Document, _
                            Optional stPrefix As String = "orig-", _
							Optional stSuffix As String)
' CopiaSeguridad Sub
' Copia el archivo original pasado y lo prefija y sufija con
'
    Dim fsFileSystem As Object, stOriginalName As String, stOriginalExtension As String, stNewFullName As String
    With dcArgumentDocument
        stOriginalName = Left(.Name, InStrRev(.Name, ".") - 1)
        stOriginalExtension = Right(.Name, Len(.Name) - InStrRev(.Name, ".") + 1)
        stNewFullName = .Path & Application.PathSeparator & stPrefix & stOriginalName & stSuffix & stOriginalExtension
    End With

    Set fsFileSystem = CreateObject("Scripting.FileSystemObject")
    fsFileSystem.CopyFile dcArgumentDocument.FullName, stNewFullName
End Sub





Function SaveAsNewFile(dcArgumentDocument As Document, _
                                Optional stPrefix As String, _
								Optional stSuffix As String = "noSuffix", _
								Optional bClose As Boolean = True)
' SaveAsNewFile Function
' Guarda una copia del documento pasado como argumento, manteniendo el original abierto
'
    Dim stOriginalName As String, stOriginalExtension As String, stNewFullName As String, dcNewDocument As Document
    With dcArgumentDocument
        stOriginalName = Left(.Name, InStrRev(.Name, ".") - 1)
        stOriginalExtension = Right(.Name, Len(.Name) - InStrRev(.Name, ".") + 1)

        If stSuffix = "noSuffix" Then stSuffix = "-" & RaMacros.GetFormattedDateAndTime(1)
        
        stNewFullName = .Path & Application.PathSeparator & stPrefix & stOriginalName & stSuffix & stOriginalExtension
        Set dcNewDocument = Documents.Add(.FullName)
    End With

    If dcArgumentDocument.CompatibilityMode < 15 Then dcArgumentDocument.Convert
    dcNewDocument.SaveAs2 FileName:=stNewFullName, FileFormat:= wdFormatDocumentDefault

    If bClose = True Then
        dcNewDocument.Close
    Else
        Set SaveAsNewFile = dcNewDocument
    End If

End Function





Function GetFormattedDateAndTime(Optional chosedInfo As Integer = 1) As String
' GetDateAndHour Function
' Devuelve un string con la fecha y hora en formato yymmdd, hhmm o yymmdd_hhmm, según requerido

    Dim formatedInfo As String

    Select Case chosedInfo
        Case 1
            GetFormattedDateAndTime = Format(Date, "yymmdd")
        Case 2
            GetFormattedDateAndTime = Format(Time, "hhnn")
        Case 3
            GetFormattedDateAndTime = Format(Date, "yymmdd") & "_" & Format(Time, "hhnn")
        Case Else
            Err.Raise Number:=513, Description:="Incorrect argument"
    End Select
End Function





Sub RemoveHeadAndFoot(dcArgumentDocument As Document)
' RemoveHeadAndFoot Sub
' Borra todos los pies y encabezados de página

' https://word.tips.net/T001777_Deleting_All_Headers_and_Footers.html

    Dim oSec As Section
    Dim oHead As HeaderFooter
    Dim oFoot As HeaderFooter
    For Each oSec In dcArgumentDocument.Sections
        For Each oHead In oSec.Headers
            If oHead.Exists Then oHead.Range.Delete
        Next oHead

        For Each oFoot In oSec.Footers
            If oFoot.Exists Then oFoot.Range.Delete
        Next oFoot
    Next oSec
End Sub





Sub ListasATexto(dcArgumentDocument As Document)
' ListasATexto Sub
' Convierte todas las viñetas de las listas a texto

' https://wordmvp.com/FAQs/Numbering/ListString.htm
' https://word.tips.net/T001857_Converting_Lists_to_Text.html

	Dim lp As Paragraph
    For Each lp In dcArgumentDocument.ListParagraphs
        lp.Range.ListFormat.ConvertNumbersToText
    Next lp
    
End Sub





Sub LimpiarFindAndReplaceParameters(dcArgumentDocument As Document)
' LimpiarFindAndReplaceParameters Sub
' Limpia los cuadros de búsqueda y reemplazo.
' Útil para llamarla después de automatizar búsquedas

' https://wordmvp.com/FAQs/MacrosVBA/ClearFind.htm
    With dcArgumentDocument.Range.Find
       .ClearFormatting
       .Replacement.ClearFormatting
       .Text = ""
       .Replacement.Text = ""
       .Forward = True
       .Wrap = wdFindStop
       .Format = False
       .MatchCase = False
       .MatchWholeWord = False
       .MatchWildcards = False
       .MatchSoundsLike = False
       .MatchAllWordForms = False
    End With
    
End Sub





Sub LimpiarEspacios(dcArgumentDocument As Document)
' LimpiarEspacios Sub
' Elimina:
    ' Más de 1 espacio seguido
    ' Espacios justo antes de marcas de párrafo, puntos, paréntesis, etc.
    ' Espacios justo después de marcas de párrafo
'
    With dcArgumentDocument.Range.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True

		Do
			.Text = "[^t]"
			.Replacement.Text = " "
			.Execute Replace:=wdReplaceAll
		Loop While .Found

		Do
			.Text = "[ ]{2;}"
			.Replacement.Text = " "
			.Execute Replace:=wdReplaceAll
		Loop While .Found
			
			.Text = "[ ]([^13,.;:\]\)\}])"
			.Replacement.Text = "\1"
			.Execute Replace:=wdReplaceAll
			
			.Text = "(^13)[ ,.;:]"
			.Replacement.Text = "\1"
			.Execute Replace:=wdReplaceAll
    End With
    
    RaMacros.LimpiarFindAndReplaceParameters dcArgumentDocument
    
End Sub





Sub LimpiarParrafosVacios(dcArgumentDocument As Document)
' LimpiarParrafosVacios Sub
' Eliminar párrafos de separación y líneas redundantes
	' Se puede completar con esta macro: https://wordmvp.com/FAQs/MacrosVBA/DeleteEmptyParas.htm
'
    Dim MyRange As Range
    Set MyRange = dcArgumentDocument.Paragraphs(1).Range
    If MyRange.Text = vbCr Then MyRange.Delete
    
    Set MyRange = dcArgumentDocument.Paragraphs.Last.Range
    If MyRange.Text = vbCr Then MyRange.Delete
    
    RaMacros.LimpiarFindAndReplaceParameters dcArgumentDocument
    
    With dcArgumentDocument.Range.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Text = "^13{2;}"
        .Replacement.Text = "^13"
        .Execute Replace:=wdReplaceAll
    End With

    RaMacros.LimpiarFindAndReplaceParameters dcArgumentDocument
    
End Sub





Sub TitulosNoPuntuacionFinal(dcArgumentDocument As Document)
' TitulosNoPuntuacionFinal Sub
' Elimina los puntos finales de los títulos
    ' Se podría hacer con RegEx, pero no parece que valga la pena el consumo de recursos
'
    Dim titulo As Integer, signos(3) As String, signoActual As Integer
    titulo = -2
    signoActual = 0
    signos(0) = "."
    signos(1) = ","
    signos(2) = ";"
    signos(3) = ":"

    With dcArgumentDocument.Range.Find
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
        For signoActual = 0 To 3 Step 1
            For titulo = -2 To -10 Step -1
                .ClearFormatting
                .Style = titulo
                .Text = signos(signoActual) & "^p"
                .Replacement.Text = "^p"
                .Execute Replace:=wdReplaceAll
            Next titulo
            .ClearFormatting
            .Style = wdStyleTitle
            .Text = signos(signoActual) & "^p"
            .Replacement.Text = "^p"
            .Execute Replace:=wdReplaceAll
        Next signoActual
    End With
     
    LimpiarFindAndReplaceParameters dcArgumentDocument
    
End Sub





Sub TitulosQuitarNumeracion(dcArgumentDocument As Document)
' TitulosQuitarNumeracion Sub
' Elimina las numeraciones de los títulos
'
    Dim iTitulo As Integer, stPatron As String, rgexNumeracion As RegExp
    dcArgumentDocument.Activate

    Set rgexNumeracion = New RegExp
    stPatron = "^\d{1,2}\.(\d{1,2}\.?)*[\s]+"

    rgexNumeracion.Pattern = stPatron
    rgexNumeracion.IgnoreCase = True
    rgexNumeracion.Global = False

    RaMacros.LimpiarFindAndReplaceParameters dcArgumentDocument

    For iTitulo = -2 To -10 Step -1
        Selection.HomeKey Unit:=wdStory, Extend:=wdMove
        With Selection.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Forward = True
            .Wrap = wdFindStop
            .MatchWildcards = False
            .Style = iTitulo
            .Replacement.Style = iTitulo
            Do
			.Execute FindText:= ""
                If .Found And rgexNumeracion.Test(Selection.Text) Then
                ' SACAR LA POSICIÓN DEL TEXTO A BORRAR
                ' CAMBIAR LA SELECCION A ESE TROZO DE TEXTO Y BORRARLO
                ' EXPANDIR SELECCIÓN O MOVERLA AL SIGUIENTE PÁRRAFO Y CONTINUAR EL BUCLE DE BÚSQUEDA
                    Selection.End = Selection.End - Len(rgexNumeracion.Replace(Selection.Text, ""))
                    Selection.Delete
                    Selection.Next(Unit:=wdParagraph, Count:=1).Select
                    Selection.Collapse Direction:=wdCollapseStart
                End If
            Loop While .Found
        End With
    Next iTitulo
    
    LimpiarFindAndReplaceParameters dcArgumentDocument

End Sub





Sub HyperlinksOnlyDomain(dcArgumentDocument As Document)
' HyperlinksOnlyDomain Sub
' Limpia los hipervínculos para que no limpien la URL completa y muestren solo el dominio
'

    Dim hlCurrent As Hyperlink, stPatron As String, stResultadoPatron As String, rgexUrlRegEx As RegExp
    stPatron = "(?:https?:(?://)?(?:www\.)?|//|www\.)([a-zA-Z\-]+?\.[a-zA-Z\-\.]+)(?:/[\S]+)?"
        ' Este es más exacto (sin puntos o guiones a principio o final del dominio), pero VBA no permite lookbehinds
    ' (?:https?:(?://)?(?:www\.)?|//|www\.)?((?:[a-zA-Z]|(?<=[a-zA-Z])-(?=[a-zA-Z]))+?\.(?:[a-zA-Z]|(?<=[a-zA-Z])[\.\-](?=[a-zA-Z]))+)(/[\S]+)?
    Set rgexUrlRegEx = New RegExp
    rgexUrlRegEx.Pattern = stPatron
    rgexUrlRegEx.IgnoreCase = True
    rgexUrlRegEx.Global = True

    For Each hlCurrent In dcArgumentDocument.Hyperlinks
        If hlCurrent.Type = 0 And rgexUrlRegEx.Test(hlCurrent.TextToDisplay) = True Then
            hlCurrent.TextToDisplay = rgexUrlRegEx.Replace(hlCurrent.TextToDisplay, "$1")
        End If
    Next hlCurrent

End Sub





Sub HyperlinksFormatting(dcArgumentDocument As Document)
' HyperlinksFormatting Sub
' Aplica el estilo Hipervínculo a todos los hipervínculos
'
    dcArgumentDocument.Activate

    ActiveWindow.View.ShowFieldCodes = Not ActiveWindow.View.ShowFieldCodes
    
    With dcArgumentDocument.Range.Find
        .ClearFormatting
        .Text = "^d HYPERLINK"
        .Replacement.Text = ""
        .Replacement.Style = wdStyleHyperlink
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
        
    RaMacros.LimpiarFindAndReplaceParameters dcArgumentDocument
    
	dcArgumentDocument.Activate
    ActiveWindow.View.ShowFieldCodes = Not ActiveWindow.View.ShowFieldCodes

End Sub





Sub ImagesToCenteredInLine(dcArgumentDocument As Document)
' ImagesToCenteredInLine Sub
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
                If .Width / (.ScaleWidth / 100) > sngRealPageWidth Then .Width = sngRealPageWidth Else .ScaleWidth = 100
                If .Height / (.ScaleHeight / 100) > sngRealPageHeight - 15 Then .Height = sngRealPageHeight - 15
                .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                If .Range.Next(Unit:=wdCharacter, Count:=1).Text <> vbCr Then
                    .Range.InsertAfter vbCr
                End If
                .Range.InsertAfter vbCr
                .Range.Next(Unit:=wdParagraph, Count:=1).Style = wdStyleNormal
                .Range.Next(Unit:=wdParagraph, Count:=1).Font.Size = 5
                If .Range.Previous(Unit:=wdCharacter, Count:=1).Text <> vbCr Then
                    .Range.InsertBefore vbCr
                End If
                .Range.InsertBefore vbCr
                .Range.Previous(Unit:=wdParagraph, Count:=1).Style = wdStyleNormal
                .Range.Previous(Unit:=wdParagraph, Count:=1).Font.Size = 5
                .Range.Style = wdStyleNormal
            End If
        End With
    Next inlShape

End Sub





Sub ComillasRectasAInglesas(dcArgumentDocument As Document)
'
' ComillasRectasAInglesas Sub
' Cambia las comillas problemáticas (" y ') por comillas inglesas
    ' Este método elimina las variables no configurables de Document.Autoformat
'
    Dim bSmtQt As Boolean
    bSmtQt = Options.AutoFormatAsYouTypeReplaceQuotes
    Options.AutoFormatAsYouTypeReplaceQuotes = True
        
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
        .Text = """"
        .Replacement.Text = """"
        .Execute Replace:=wdReplaceAll
        .Text = "'"
        .Replacement.Text = "'"
        .Execute Replace:=wdReplaceAll
    End With
    
    Options.AutoFormatAsYouTypeReplaceQuotes = bSmtQt

End Sub





Sub DirectFormattingToStyles(dcArgumentDocument As Document)
'
' DirectFormattingToStyles Sub
' Convierte los estilos directos de negritas y cursivas a los estilos Strong y Emphasis, respectivamente
'
    Dim iCounter As Integer, arrstStylesToApply(13) As WdBuiltinStyle
    dcArgumentDocument.Activate

    arrstStylesToApply(0) = wdStyleNormal
    arrstStylesToApply(1) = wdStyleCaption
    arrstStylesToApply(2) = wdStyleList
    arrstStylesToApply(3) = wdStyleList2
    arrstStylesToApply(4) = wdStyleList3
    arrstStylesToApply(5) = wdStyleListBullet
    arrstStylesToApply(6) = wdStyleListBullet2
    arrstStylesToApply(7) = wdStyleListBullet3
    arrstStylesToApply(8) = wdStyleListBullet3
    arrstStylesToApply(9) = wdStyleListBullet4
    arrstStylesToApply(10) = wdStyleListBullet5
    arrstStylesToApply(11) = wdStyleListNumber
    arrstStylesToApply(12) = wdStyleListNumber2
    arrstStylesToApply(13) = wdStyleListNumber3

    With dcArgumentDocument.Range.Find
        .ClearFormatting
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False

        .Style = wdStyleQuote
        .Font.Bold = True
        .Replacement.Style = wdStyleStrong
        .Execute Replace:=wdReplaceAll

        For iCounter = 0 To 13
            .Style = arrstStylesToApply(iCounter)
            .Font.Bold = True
            .Font.Italic = True
            .Replacement.Style = wdStyleIntenseEmphasis
            .Execute Replace:=wdReplaceAll
            .Font.Bold = True
            .Font.Italic = False
            .Replacement.Style = wdStyleStrong
            .Execute Replace:=wdReplaceAll
            .Font.Bold = False
            .Font.Italic = True
            .Replacement.Style = wdStyleEmphasis
            .Execute Replace:=wdReplaceAll
        Next iCounter

        .ClearFormatting
        .Text = "^f"
        .Replacement.Style = wdStyleFootnoteReference
        .Execute Replace:=wdReplaceAll

    End With
    
    RaMacros.LimpiarFindAndReplaceParameters dcArgumentDocument
    
End Sub





Sub LimpiezaBasica(dcArgumentDocument As Document)
' LimpiezaBasica Sub
' Ejecuta limpieza de espacios, signos y otros elementos innecesarios:
'
    Application.ScreenUpdating = False
    
    RaMacros.LimpiarEspacios dcArgumentDocument
    RaMacros.LimpiarParrafosVacios dcArgumentDocument
    RaMacros.LimpiarFindAndReplaceParameters dcArgumentDocument
    
    Application.ScreenUpdating = True
    
End Sub





