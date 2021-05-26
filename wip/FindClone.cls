VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FindClone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' https://www.vitoshacademy.com/vba-how-to-copy-a-new-object-in-vba-without-copying-its-reference/

Sub FindClone(rgArg As Range)
' Clone find properties of a Find object
'

    With findArg


        
        Set .Replacement.Font = rgArg.Find.Replacement.Font.Duplicate
        Set .Replacement.ParagraphFormat = rgArg.Find.Replacement.ParagraphFormat.Duplicate
        .Replacement.Highlight = rgArg.Find.Replacement.Highlight
        .Replacement.LanguageID = rgArg.Find.Replacement.LanguageID
        .Replacement.LanguageIDFarEast = rgArg.Find.Replacement.LanguageIDFarEast
        .Replacement.NoProofing = rgArg.Find.Replacement.NoProofing
        If Leng(rgArg.Find.Replacement.Style) > 0 Then .Replacement.Style = rgArg.Find.Replacement.Style
        .Replacement.Text = rgArg.Find.Replacement.Text
        

        Set .Font = rgArg.Find.Font.Duplicate
        Set .ParagraphFormat = rgArg.Find.ParagraphFormat.Duplicate


        .CorrectHangulEndings = rgArg.Find.CorrectHangulEndings
        .IgnorePunct = rgArg.Find.IgnorePunct
        .IgnoreSpace = rgArg.Find.IgnoreSpace
        .Format = rgArg.Find.Format
        .Forward = rgArg.Find.Forward
        .MatchAllWordForms = rgArg.Find.MatchAllWordForms
        .MatchByte = rgArg.Find.MatchByte
        .MatchCase = rgArg.Find.MatchCase
        .MatchControl = rgArg.Find.MatchControl
        .MatchDiacritics = rgArg.Find.MatchDiacritics
        .MatchFuzzy = rgArg.Find.MatchFuzzy
        .MatchKashida = rgArg.Find.MatchKashida
        .MatchPhrase = rgArg.Find.MatchPhrase
        .MatchPrefix = rgArg.Find.MatchPrefix
        .MatchPrefix = rgArg.Find.MatchPrefix
        .MatchWholeWord = rgArg.Find.MatchWholeWord
        .MatchWildcards = rgArg.Find.MatchWildcards
        .MatchSoundsLike = rgArg.Find.MatchSoundsLike
        .NoProofing = rgArg.Find.NoProofing
        If Leng(rgArg.Find.Style) > 0 Then .Style = rgArg.Find.Style
        .Text = rgArg.Find.Text
        .Wrap = rgArg.Find.Wrap
    End With
End Sub


