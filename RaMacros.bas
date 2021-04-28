Attribute VB_Name = "RaMacros"
Option Explicit

Function SaveAsNewFile(Optional stPrefix As String, _
                        Optional stSuffix As String = "noSuffix", _
                        Optional bClose As Boolean = True)
' SaveAsNewFile Function
' Guarda una copia del archivo actual, manteniendo el original abierto y sin guardarlo
'
    Dim stOriginalName As String, stOriginalExtension As String, stNewFullName As String, dcNewDocument As Document
    
    With ActiveDocument
        stOriginalName = Left(.Name, InStrRev(.Name, ".") - 1)
        stOriginalExtension = Right(.Name, Len(.Name) - InStrRev(.Name, ".") + 1)

        If stSuffix = "noSuffix" Then stSuffix = "-" & RaMacros.GetFormattedDateAndTime(1)
        
        stNewFullName = .Path & Application.PathSeparator & stPrefix & stOriginalName & stSuffix & stOriginalExtension
        Set dcNewDocument = Documents.Add(.FullName)
    End With

    dcNewDocument.SaveAs2 FileName:=stNewFullName

    If bClose = True Then
        dcNewDocument.Close
    Else
        SaveAsNewFile = dcNewDocument
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

Sub RemoveHeadAndFoot()

' RemoveHeadAndFoot Function
' Borra todos los pies y encabezados de página

' https://word.tips.net/T001777_Deleting_All_Headers_and_Footers.html

    Dim oSec As Section
    Dim oHead As HeaderFooter
    Dim oFoot As HeaderFooter

    For Each oSec In ActiveDocument.Sections
        For Each oHead In oSec.Headers
            If oHead.Exists Then oHead.Range.Delete
        Next oHead

        For Each oFoot In oSec.Footers
            If oFoot.Exists Then oFoot.Range.Delete
        Next oFoot
    Next oSec
End Sub

Sub ListasATexto()
'
' ListasATexto Macro
' Convierte todas las viñetas de las listas a texto

' https://wordmvp.com/FAQs/Numbering/ListString.htm
' https://word.tips.net/T001857_Converting_Lists_to_Text.html

Dim lp As Paragraph

    For Each lp In ActiveDocument.ListParagraphs
        lp.Range.ListFormat.ConvertNumbersToText
    Next lp
    
End Sub

Function LimpiarFindAndReplaceParameters()
'
' LimpiarFindAndReplaceParameters Macro
' Limpia los cuadros de búsqueda y reemplazo.
' Útil para llamarla después de automatizar búsquedas

' https://wordmvp.com/FAQs/MacrosVBA/ClearFind.htm

    With ActiveDocument.Range.Find
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
    
End Function

Function LimpiarEspacios()
'
' LimpiarEspacios Function
'
' Elimina:
    ' Más de 1 espacio seguido
    ' Espacios justo antes de marcas de párrafo, puntos, paréntesis, etc.
    ' Espacios justo después de marcas de párrafo
'
    Application.Run "RaMacros.LimpiarFindAndReplaceParameters"
    
    With ActiveDocument.Range.Find
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
        .Text = " {2;}"
        .Replacement.Text = " "
        .Execute Replace:=wdReplaceAll
        .Text = " ([^13,.;:\]\)\}])"
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll
        .Text = "(^13) "
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll
    End With
    
    Application.Run "RaMacros.LimpiarFindAndReplaceParameters"
    
End Function

Function NoParrafosVacios()
'
' NoParrafosVacios Macro
'
' Eliminar párrafos de separación y líneas redundantes
'
' Se puede completar con esta macro: https://wordmvp.com/FAQs/MacrosVBA/DeleteEmptyParas.htm

    Dim MyRange As Range
    
    Set MyRange = ActiveDocument.Paragraphs(1).Range
    If MyRange.Text = vbCr Then MyRange.Delete
    
    Set MyRange = ActiveDocument.Paragraphs.Last.Range
    If MyRange.Text = vbCr Then MyRange.Delete
    
    Application.Run "RaMacros.LimpiarFindAndReplaceParameters"
    
    With ActiveDocument.Range.Find
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

    Application.Run "RaMacros.LimpiarFindAndReplaceParameters"
    
End Function

Sub LimpiezaBasica()
'
' LimpiezaBasica Macro
'
' Ejecuta limpieza de espacios, signos y otros elementos innecesarios:
'
    Application.ScreenUpdating = False
    
    Application.Run "RaMacros.LimpiarEspacios"
    Application.Run "RaMacros.NoParrafosVacios"
    Application.Run "RaMacros.LimpiarFindAndReplaceParameters"
    
    Application.ScreenUpdating = True
    
End Sub

Sub TitulosQuitarPuntacionFinal()
'
' TitulosQuitarPuntuacionFinal Macro
'
' Elimina los puntos finales de los títulos
'
    Dim titulo As Integer
    titulo = -2
    
    Dim signoActual As Integer
    signoActual = 0
    
    Dim signos(3) As String
    signos(0) = "."
    signos(1) = ","
    signos(2) = ";"
    signos(3) = ":"

    Application.Run "LimpiarFindAndReplaceParameters"
    
    With ActiveDocument.Range.Find
        For signoActual = 0 To 3 Step 1
            For titulo = -2 To -10 Step -1
                .ClearFormatting
                .Style = ActiveDocument.Styles(titulo)
                .Text = signos(signoActual) & "^p"
                .Replacement.Text = "^p"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = True
                .MatchCase = False
                .MatchWholeWord = False
                .MatchAllWordForms = False
                .MatchSoundsLike = False
                .MatchWildcards = False
                .Execute Replace:=wdReplaceAll
            Next titulo
        Next signoActual
    End With
     
    Application.Run "LimpiarFindAndReplaceParameters"
    
End Sub

Sub TitulosQuitarNumeracion()
'
' TitulosQuitarNumeracion Macro
'
' Elimina las numeraciones de los títulos
'

    Dim titulo As Integer
    titulo = -2
    
    Application.Run "LimpiarFindAndReplaceParameters"
    
    With ActiveDocument.Range.Find
        For titulo = -2 To -10 Step -1
            .ClearFormatting
            .Style = ActiveDocument.Styles(titulo)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True
            .Text = "[0-9].[ ^t]"
            .Replacement.Text = ""
            .Execute Replace:=wdReplaceAll
            .Text = "[0-9].[0-9][ ^t]"
            .Replacement.Text = ""
            .Execute Replace:=wdReplaceAll
            .Text = "[0-9].[0-9].[0-9].[ ^t]"
            .Replacement.Text = ""
            .Execute Replace:=wdReplaceAll
        Next titulo
    End With
    
    Application.Run "LimpiarFindAndReplaceParameters"

End Sub

Sub HipervinculosFormatear()
'
' HipervinculosFormatear Macro
'
' Aplica el estilo Hipervínculo a todos los hipervínculos
'

    ActiveWindow.View.ShowFieldCodes = Not ActiveWindow.View.ShowFieldCodes
    
    Application.Run "RaMacros.LimpiarFindAndReplaceParameters"
    
    With ActiveDocument.Range.Find
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
        
    Application.Run "RaMacros.LimpiarFindAndReplaceParameters"
    
    ActiveWindow.View.ShowFieldCodes = Not ActiveWindow.View.ShowFieldCodes

End Sub

Sub ImagenesFormatoCorrecto()
'
' ImagenesFormatoCorrecto Macro
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

Sub ComillasRectasAInglesas()
'
' ComillasRectasAInglesas Macro
' Cambia las comillas problemáticas (" y ') por comillas inglesas
    ' Este método elimina las variables no configurables de Document.Autoformat
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
        .Text = """"
        .Replacement.Text = """"
        .Execute Replace:=wdReplaceAll
        .Text = "'"
        .Replacement.Text = "'"
        .Execute Replace:=wdReplaceAll
    End With
    
    Options.AutoFormatAsYouTypeReplaceQuotes = bSmtQt

End Sub

Sub DirectFormattingToStyles()
'
' DirectFormattingToStyles Macro
'
' Convierte los estilos directos de negritas y cursivas a los estilos Strong y Emphasis, respectivamente
'
    Dim iCounter As Integer, arrstStylesToApply(13) As WdBuiltinStyle
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

    Application.Run "RaMacros.LimpiarFindAndReplaceParameters"
    
    With ActiveDocument.Range.Find
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
    
    Application.Run "RaMacros.LimpiarFindAndReplaceParameters"
    
End Sub

Sub HyperlinksOnlyDomain()
'
' Sub HyperlinksOnlyDomain Macro
'
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

    For Each hlCurrent In ActiveDocument.Hyperlinks
        If hlCurrent.Type = 0 And rgexUrlRegEx.Test(hlCurrent.TextToDisplay) = True Then
            hlCurrent.TextToDisplay = rgexUrlRegEx.Replace(hlCurrent.TextToDisplay, "$1")
        End If
    Next hlCurrent

End Sub

