Sub CleanBullets ( _
					Optional rgArg As Range, _
					Optional dcArg As Document, _
					Optional vStyle As Variant _
)
' Cleans bullets left by bad implemented lists
' Params:
	' rgArg: target range. If nothing, dcArg is mandatory. Supersedes dcArg
	' dcArg: target document. If nothing, rgArg is mandatory
	' styArg: the style that will be targeted
' Symbols and unicode:
	' ? 	->	8270
	' ? 	->	8277
	' ? 	->	9642
	' ? 	->	9656
	' ? 	->	9666
	' ? 	->	9667
	' ? 	->	9655
	' ? 	->	9724
	' ? 	->	9723
	' ? 	->	9679
	' ? 	->	9676
	' ? 	->	9671
	' ? 	->	9670
'
	Dim cllRanges As New Collection
	Dim rgCurrent As Range, rgFind As Range
	Dim pCurrent As Paragraph

	If rgArg Is Nothing Then
		If dcArg Is Nothing Then Err.Raise 516,, "There is no target range"
		For Each rgCurrent In dcArg.StoryRanges
			cllRanges.Add rgCurrent
			If rgCurrent.StoryType >= 5 Then Exit For
		Next rgCurrent
	Else
		Set dcArg = rgArg.Parent
		cllRanges.Add rgArg
	End If

	If Not (TypeName(vStyle) = Style _
		Or TypeName(vStyle) = String _
		Or TypeName(vStyle) = Integer) _
	Then Err.Raise Err.Raise 518,, "vStyle must be a string, integer or style"













	For Each rgCurrent In cllRanges
		For Each pCurrent In rgCurrent.Paragraphs
			Set rgFind = pCurrent.Range
			rgFind.End = rgFind.Start + 4
			With rgFind.Find
				.MatchWildcards = True
				.Text = "[–\-—•" & ChrW$(8270) & ChrW$(8277) & ChrW$(9642) & ChrW$(9656) _
					& ChrW$(9666) & ChrW$(9667) & ChrW$(9655) & ChrW$(9724) & ChrW$(9723) _
					& ChrW$(9679) & ChrW$(9676) & ChrW$(9671) & ChrW$(9670) & "]"
				.Replacement.Text = ""
				.Execute Replace:=wdReplaceAll
			End With
		Next pCurrent
	Next rgCurrent
End Sub
