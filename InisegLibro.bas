Attribute VB_Name = "InisegLibro"
Option Explicit

Sub InisegInterlineado()
'
' InterlineadoSinEspaciado Macro
'
' Interlineado de 1,15 sin espaciado entre párrafos
    ' Eliminar los espaciados verticales entre párrafos y aplica el interlineado correcto

    Selection.WholeStory

    With Selection.ParagraphFormat
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
    
End Sub

Sub Iniseg1Limpieza()
'
' Iniseg1Pre1Limpieza Macro
'
' Ejecuta limpieza de espacios innecesarios:

    InisegInterlineado
    
    Application.Run "RaMacros.LimpiezaBasica"
    
    Application.Run "RaMacros.RemoveHeadAndFoot"
    
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

    Application.Run "HipervinculosFormatear"

    InisegInterlineado
    
End Sub

Sub Iniseg3ParrafosSeparacion()
'
' ParrafosSeparacion Macro
'
'

    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Heading 1")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
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
    End With
Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Heading 2")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Heading 3")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Heading 4")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
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
    End With
    ActiveWindow.ActivePane.VerticalPercentScrolled = 0
    ActiveWindow.DocumentMap = True
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Heading 4")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Normal")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("List")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("List 2")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("List 3")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("List Bullet")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("List Bullet 2")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("List Bullet 3")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("List Bullet 4")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("List Bullet 5")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("List Number")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("List Number 2")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("List Number 3")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 4
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 5
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 6
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 8
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = 11
    With Selection.Find
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
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub


