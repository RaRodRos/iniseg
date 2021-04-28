Attribute VB_Name = "InisegStory"
Option Explicit

Sub Iniseg4ConversionParrafos()
'
' StoryConversionParrafos Macro
' Conversion de Word impreso a formato para Storyline
'

    ' Eliminar los espaciados verticales entre párrafos (repetición)
    Application.Run "InisegLibro.InisegInterlineado"

    ' Cambio del tamaño de Titulo 1 de 16 a 17
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

' Cambio de tamaño de parrafos de separacion

    ' Listas: 4 a 2
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Font.Size = 4
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 2
    With Selection.Find
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
    End With

    ' Parrafos normales: 5 a 4.
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Font.Size = 5
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 4
    With Selection.Find
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
    End With

    ' Titulos 2, 3 y 4: 8 a 6.
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Font.Size = 8
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 6
    With Selection.Find
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
    End With

' Titulos 1: 11 a 8

    ' Cambiar el primer párrafo tras todos los Headings 1 y marcarlo con la palabra FISTRO
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles(wdStyleHeading1)
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "(*@^13)"
        .Replacement.Text = "\1FISTRO"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With

    ' Seleccionar dos simbolos de parrafos seguidos, de tamaño 11,
        ' anterior a todos Headings 1, y marcarlos con la palabra FISTRO
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Font.Size = 11
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^13^13"
        .Replacement.Text = "^13FISTRO^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With

    ' Seleccionar todos los parrafos con la palabra FISTRO y darles un estilo "Normal"
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles(wdStyleNormal)
    With Selection.Find
        .Text = "FISTRO^13"
        .Replacement.Text = "FISTRO^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With

    ' Seleccionar todos los parrafos con la palabra FISTRO, cambiar su tamaño a 8
        ' y borrar la palabra FISTRO
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 8
    With Selection.Find
        .Text = "FISTRO^13"
        .Replacement.Text = "^13"
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
    
    Application.Run "RaMacros.ListasATexto"
    
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
    
    Selection.HomeKey Unit:=wdStory
    Application.Run "RaMacros.LimpiarFindAndReplaceParameters"
    
    Do While bSeguir = True
    
        With Selection.Find
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
    
    Selection.HomeKey Unit:=wdStory
    Application.Run "RaMacros.LimpiarFindAndReplaceParameters"
    
End Sub
