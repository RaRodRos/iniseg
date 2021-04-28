Attribute VB_Name = "RaMacros"
Option Explicit

Private Sub RemoveHeadAndFoot()

' RemoveHeadAndFoot Macro
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

Sub BorrarEstilosNoUsados()
'
' BorrarEstilosNoUsados Macro
' Borra todos los estilos no usados respetando los
    ' iniciales (si no simplemente volverían al default)

' https://www.msofficeforums.com/word-vba/37489-macro-deletes-unused-styles.html
' ANTES USABA ESTA
    ' https://word.tips.net/T001337_Removing_Unused_Styles.html

' TODO
    ' Informe final (número borrado, número dejado, errores...)

    Dim Doc As Document, Rng As Range, Shp As Shape
    
    Dim StlNm As String, i As Long, bDel As Boolean
    
    Application.ScreenUpdating = False
    
    On Error Resume Next
    
    Set Doc = ActiveDocument
    
    With Doc
      For i = .Styles.Count To 1 Step -1
        With .Styles(i)
          If .BuiltIn = False And .Linked = False Then
            bDel = True: StlNm = .NameLocal
            For Each Rng In Doc.StoryRanges
              With Rng.Find
                .ClearFormatting
                .Format = True
                .Style = StlNm
                .Execute
                If .Found = True Then
                  bDel = False
                  Exit For
                End If
              End With
              For Each Shp In Rng.ShapeRange
                If Not Shp.TextFrame Is Nothing Then
                  With Shp.TextFrame.TextRange.Find
                    .ClearFormatting
                    .Format = True
                    .Style = StlNm
                    .Execute
                    If .Found = True Then
                      bDel = False
                      Exit For
                    End If
                  End With
                End If
              Next
            Next
            If bDel = True Then .Delete
          End If
        End With
      Next
    End With
    
    Application.ScreenUpdating = True

End Sub

Sub ListasATexto()
'
' ListasATexto Macro
' Convierte todas las viñetas de las listas a texto

' https://wordmvp.com/FAQs/Numbering/ListString.htm
' https://word.tips.net/T001857_Converting_Lists_to_Text.html

' TODO
    ' Una macro que destransforme
    ' Que solo se ejecute en los títulos
        ' Que dé a elegir en qué estilos se quiere ejecutar

Dim lp As Paragraph

    For Each lp In ActiveDocument.ListParagraphs
        lp.Range.ListFormat.ConvertNumbersToText
    Next lp
    
End Sub

Private Sub LimpiarFindAndReplaceParameters()
'
' LimpiarFindAndReplaceParameters Macro
' Limpia los cuadros de búsqueda y reemplazo.
' Útil para llamarla después de automatizar búsquedas

' https://wordmvp.com/FAQs/MacrosVBA/ClearFind.htm

    With Selection.Find
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

Private Sub NoDosEspacios()
'
' NoDosEspacios Macro
'
' Eliminar más de 1 espacio seguido
'
    Selection.HomeKey Unit:=wdStory
    
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = " {2;}"
        .Replacement.Text = " "
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
    
    RaMacros.LimpiarFindAndReplaceParameters
    
End Sub

Private Sub NoEspacioAntesParrafo()
'
' NoEspacioAntesParrafo Macro
'
' Eliminar espacios justo antes de marcas de párrafo
'
    Selection.HomeKey Unit:=wdStory

    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = " (^13)"
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
   
    RaMacros.LimpiarFindAndReplaceParameters
    
End Sub

Private Sub NoEspacioDespuesParrafo()
'
' NoEspacioDespuesParrafo Macro
'
' Eliminar espacios justo después de marcas de párrafo
'
    Selection.HomeKey Unit:=wdStory

    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "(^13) "
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

    RaMacros.LimpiarFindAndReplaceParameters
    
End Sub

Private Sub NoParrafosVacios()
'
' NoParrafosVacios Macro
'
' Eliminar párrafos de separación y líneas redundantes
'
' Se puede completar con esta macro: https://wordmvp.com/FAQs/MacrosVBA/DeleteEmptyParas.htm

    Selection.HomeKey Unit:=wdStory

    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^13{2;}"
        .Replacement.Text = "^13"
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

    RaMacros.LimpiarFindAndReplaceParameters
    
End Sub

Sub LimpiezaBasica()
'
' LimpiezaBasica Macro
'
' Ejecuta limpieza de espacios innecesarios:

    Selection.HomeKey Unit:=wdStory

    NoDosEspacios

    NoEspacioAntesParrafo

    NoEspacioDespuesParrafo

    NoParrafosVacios

    LimpiarFindAndReplaceParameters
    
    Selection.HomeKey Unit:=wdStory
    
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
    Selection.HomeKey Unit:=wdStory
    
    With Selection.Find
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
    Selection.HomeKey Unit:=wdStory
    
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
    Selection.HomeKey Unit:=wdStory
    
    With Selection.Find
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
    Selection.HomeKey Unit:=wdStory
End Sub

Sub HipervinculosFormatear()
'
' HipervinculosFormatear Macro
'
' Aplica el estilo Hipervínculo a todos los hipervínculos
'

    ActiveWindow.View.ShowFieldCodes = Not ActiveWindow.View.ShowFieldCodes
    
    Application.Run "RaMacros.LimpiarFindAndReplaceParameters"
    
    With Selection.Find
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
