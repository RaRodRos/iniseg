Option Explicit

' En este módulo se implementarán todas las subrutinas y funciones a las que
	' pueda interesar acceder desde la UI

Sub uiSaveAsNewFileSub()
    RaMacros.SaveAsNewFile ActiveDocument
End Sub

Sub uiCopySecurity()
    RaMacros.CopySecurity ActiveDocument
End Sub

Sub uiHeadersFootersRemove()
    RaMacros.HeadersFootersRemove ActiveDocument
End Sub

Sub uiListsToText()
    RaMacros.ListsToText ActiveDocument
End Sub

Sub uiHeadingsNoPunctuation()
	RaMacros.HeadingsNoPunctuation ActiveDocument
End Sub

Sub uiHeadingsNoNumeration()
    RaMacros.HeadingsNoNumeration ActiveDocument
End Sub

Sub uiHyperlinksAddressToText()
' Chages the texttodisplay property of the selected hyperlinks to their complete address
'
    Dim hlCurrent As Hyperlink

    For Each hlCurrent In Selection.Hyperlinks
        hlCurrent.TextToDisplay = hlCurrent.Address
    Next hlCurrent
End Sub

Sub uiHyperlinksFormatting()
    RaMacros.HyperlinksFormatting ActiveDocument
End Sub

Sub uiHyperlinksOnlyDomain()
	RaMacros.HyperlinksOnlyDomain ActiveDocument
End Sub

Sub uiImagesToCenteredInLine()
    RaMacros.ImagesToCenteredInLine ActiveDocument
End Sub

Sub uiQuotesStraightToCurly()
    RaMacros.QuotesStraightToCurly ActiveDocument
End Sub

Sub uiCleanBasic()
    RaMacros.CleanBasic ActiveDocument
End Sub

Sub uiStylesDeleteUnused()
    RaMacros.StylesDeleteUnused ActiveDocument, True
End Sub

Sub uiStylesNoDirectFormatting()
    RaMacros.StylesNoDirectFormatting ActiveDocument
End Sub

Sub uiSectionBreakBeforeHeading()
    RaMacros.SectionBreakBeforeHeading ActiveDocument
End Sub

Sub uiSectionsFillBlankPages()
    RaMacros.SectionsFillBlankPages ActiveDocument
End Sub

Sub uiSectionsExportEachToFiles()
    RaMacros.SectionsExportEachToFiles ActiveDocument
End Sub
