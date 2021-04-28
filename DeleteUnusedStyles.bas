Attribute VB_Name = "DeleteUnusedStyles"
Option Explicit

Sub DeleteUnusedStyles()
'
' BorrarEstilosNoUsados Macro
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
        For Each oStyle In ActiveDocument.Styles
            'Only check out non-built-in styles
            If oStyle.BuiltIn = False Then
                If StyleInUse(oStyle.NameLocal) = False Then
                    Application.OrganizerDelete Source:=ActiveDocument.FullName, _
                    Name:=oStyle.NameLocal, Object:=wdOrganizerObjectStyles
                    sCount = sCount + 1
                End If
            End If
        Next oStyle
        lTotalSCount = lTotalSCount + sCount
    Loop While sCount > 0

    If lTotalSCount <> 0 Then MsgBox lTotalSCount & " styles deleted"

End Sub

Function StyleInUse(Styname As String) As Boolean
    ' Is Stryname used any of ActiveDocument's story
    Dim Stry As Range
    Dim Shp As Shape
    Dim txtFrame As TextFrame

    If Not ActiveDocument.Styles(Styname).InUse Then StyleInUse = False: Exit Function
    ' check if Currently used in a story?

    For Each Stry In ActiveDocument.StoryRanges
        If StoryInUse(Stry) Then
            If StyleInUseInRangeText(Stry, Styname) Then StyleInUse = True: Exit Function
            For Each Shp In Stry.ShapeRange
                Set txtFrame = Shp.TextFrame
                If Not txtFrame Is Nothing Then
                    If txtFrame.HasText Then
                        If txtFrame.TextRange.Characters.Count > 1 Then
                            If StyleInUseInRangeText(txtFrame.TextRange, Styname) Then StyleInUse = True: Exit Function
                        End If
                    End If
                End If
            Next Shp
        End If
    Next Stry
    StyleInUse = False ' Not currently in use.
End Function

Function StyleInUseInRangeText(rng As Range, Styname As String) As Boolean
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

Function StoryInUse(Stry As Range) As Boolean
    ' Note: this will mark even the always-existing stories as not in use if they're empty
    If Not Stry.StoryLength > 1 Then StoryInUse = False: Exit Function
    Select Case Stry.StoryType
        Case wdMainTextStory, wdPrimaryFooterStory, wdPrimaryHeaderStory: StoryInUse = True
        Case wdEvenPagesFooterStory, wdEvenPagesHeaderStory: StoryInUse = Stry.Sections(1).PageSetup.OddAndEvenPagesHeaderFooter = True
        Case wdFirstPageFooterStory, wdFirstPageHeaderStory: StoryInUse = Stry.Sections(1).PageSetup.DifferentFirstPageHeaderFooter = True
        Case wdFootnotesStory, wdFootnoteContinuationSeparatorStory: StoryInUse = ActiveDocument.Footnotes.Count > 1
        Case wdFootnoteSeparatorStory, wdFootnoteContinuationNoticeStory: StoryInUse = ActiveDocument.Footnotes.Count > 1
        Case wdEndnotesStory, wdEndnoteContinuationSeparatorStory: StoryInUse = ActiveDocument.Endnotes.Count > 1
        Case wdEndnoteSeparatorStory, wdEndnoteContinuationNoticeStory: StoryInUse = ActiveDocument.Endnotes.Count > 1
        Case wdCommentsStory: StoryInUse = ActiveDocument.Comments.Count > 1
        Case wdTextFrameStory: StoryInUse = ActiveDocument.Frames.Count > 1
        Case Else: StoryInUse = False ' Must be some new or unknown wdStoryType
    End Select
End Function


