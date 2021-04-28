Option Explicit

Function StylesDeleteUnused(dcArgument As Document, _
							Optional ByVal bMsgBox As Boolean = False) As Long
' Deletes unused styles using multiple loops to respect their hierarchy 
	' (avoiding the deletion of fathers without use, like in lists)
' Based on:
	' https://word.tips.net/T001337_Removing_Unused_Styles.html
' Modifications:
	' It runs until no unused styles left
	' A message with the number of styles must be turn on by the bMsgBox parameter
	' It's now a function that returns the number of deleted styles
'
	Dim oStyle As Style
	Dim sCount As Long
	Dim lTotalSCount As Long
	Dim sStart As Single, sEnd As Single

	sStart = Timer
	lTotalSCount = 0
	Do
		sCount = 0
		For Each oStyle In dcArgument.Styles
			'Only check out non-built-in styles
			If oStyle.BuiltIn = False Then
				If StyleInUse(oStyle.NameLocal, dcArgument) = False Then
					Application.OrganizerDelete Source:= dcArgument.FullName, _
					Name:= oStyle.NameLocal, Object:=wdOrganizerObjectStyles
					sCount = sCount + 1
				End If
			End If
		Next oStyle
		lTotalSCount = lTotalSCount + sCount
	Loop While sCount > 0

	If bMsgBox Then 
		MsgBox lTotalSCount & " styles deleted"
		StylesDeleteUnused = lTotalSCount
	End If

	sEnd = Timer
	Debug.Print lTotalSCount & " styles erased in " & sEnd - sStart " seconds"
	StylesDeleteUnused = lTotalSCount
End Function

Function StyleInUse(ByVal Styname As String, dcArgument As Document) As Boolean
' Del mismo desarrollador que StylesDeleteUnused
' Is Stryname used any of dcArgument's story
	Dim Stry As Range
	Dim Shp As Shape
	Dim txtFrame As TextFrame

	If Not dcArgument.Styles(Styname).InUse Then StyleInUse = False: Exit Function
	' check if Currently used in a story?

	For Each Stry In dcArgument.StoryRanges
		If StoryInUse(dcArgument, Stry) Then
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
' Del mismo desarrollador que StylesDeleteUnused
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

Function StoryInUse(dcArgument As Document, Stry As Range) As Boolean
' Del mismo desarrollador que StylesDeleteUnused
' Note: this will mark even the always-existing stories as not in use if they're empty
	If Not Stry.StoryLength > 1 Then StoryInUse = False: Exit Function
	Select Case Stry.StoryType
		Case wdMainTextStory, wdPrimaryFooterStory, wdPrimaryHeaderStory: StoryInUse = True
		Case wdEvenPagesFooterStory, wdEvenPagesHeaderStory: StoryInUse = Stry.Sections(1).PageSetup.OddAndEvenPagesHeaderFooter = True
		Case wdFirstPageFooterStory, wdFirstPageHeaderStory: StoryInUse = Stry.Sections(1).PageSetup.DifferentFirstPageHeaderFooter = True
		Case wdFootnotesStory, wdFootnoteContinuationSeparatorStory: StoryInUse = dcArgument.Footnotes.Count > 1
		Case wdFootnoteSeparatorStory, wdFootnoteContinuationNoticeStory: StoryInUse = dcArgument.Footnotes.Count > 1
		Case wdEndnotesStory, wdEndnoteContinuationSeparatorStory: StoryInUse = dcArgument.Endnotes.Count > 1
		Case wdEndnoteSeparatorStory, wdEndnoteContinuationNoticeStory: StoryInUse = dcArgument.Endnotes.Count > 1
		Case wdCommentsStory: StoryInUse = dcArgument.Comments.Count > 1
		Case wdTextFrameStory: StoryInUse = dcArgument.Frames.Count > 1
		Case Else: StoryInUse = False ' Must be some new or unknown wdStoryType
	End Select
End Function






Sub StylesNoDirectFormatting(dcArgument As Document)
' Convierte los estilos directos de negritas y cursivas a los estilos Strong y Emphasis, respectivamente
'
	Dim iCounter As Integer, arrstStylesToApply(13) As WdBuiltinStyle
	dcArgument.Activate

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

	Application.ScreenUpdating = False
	With dcArgument.Range.Find
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
	RaMacros.FindAndReplaceClearParameters
	Application.ScreenUpdating = True
End Sub






Sub CopySecurity(dcArgument As Document, _
				Optional ByVal stPrefix As String, _
				Optional ByVal stSuffix As String)
' Copies dcArgument adding the suffix and/or prefix passed as arguments. In case
' there are none, it appends a number
'
	Dim fsFileSystem As Object
	Dim stOriginalName As String, stOriginalExtension As String, stNewFullName As String
	Dim iCount As Integer

	stOriginalName = Left(dcArgument.Name, InStrRev(dcArgument.Name, ".") - 1)
	stOriginalExtension = Right(dcArgument.Name, Len(dcArgument.Name) - InStrRev(dcArgument.Name, ".") + 1)
	stNewFullName = dcArgument.Path & Application.PathSeparator & stPrefix _
		& stOriginalName & stSuffix & stOriginalExtension

	Do While Dir(stNewFullName) > ""
		stNewFullName = dcArgument.Path & Application.PathSeparator & stPrefix _
			& stOriginalName & stSuffix & "-" & Format(iCount, "00") & stOriginalExtension
		iCount = iCount + 1
	Loop

	Set fsFileSystem = CreateObject("Scripting.FileSystemObject")
	fsFileSystem.CopyFile dcArgument.FullName, stNewFullName
End Sub





Function SaveAsNewFile(dcArgument As Document, _
						Optional ByVal stPrefix As String, _
						Optional ByVal stSuffix As String, _
						Optional ByVal bClose As Boolean = True)
' Guarda una copia del documento pasado como argumento, manteniendo el original abierto y convirtiéndolo al formato actual
' Args:
	' stPrefix: string to prefix the new document's name
	' stSuffix: string to suffix the new document's name. By default it will add the current date
	' bClose: if True the new document is saved AND closed
'
	Dim stOriginalName As String, stOriginalExtension As String, stNewFullName As String, dcNewDocument As Document

	stOriginalName = Left(dcArgument.Name, InStrRev(dcArgument.Name, ".") - 1)
	stOriginalExtension = Right(dcArgument.Name, Len(dcArgument.Name) - InStrRev(dcArgument.Name, ".") + 1)
	If stSuffix = vbNullString Then stSuffix = "-" & Format(Date, "yymmdd")
	stNewFullName = dcArgument.Path & Application.PathSeparator & stPrefix _
		& stOriginalName & stSuffix & stOriginalExtension
	Set dcNewDocument = Documents.Add(dcArgument.FullName)

	' IF THE FILE GETS CONVERTED TO THE LATEST VERSION IT CAN MESS UP SOME FIELDS
		' (INCLUDEPICTURE, particularly), so it's better to do it later
	'If dcArgument.CompatibilityMode < 15 Then dcArgument.Convert

	If Dir(stNewFullName) > "" Then
		stNewFullName = dcArgument.Path & Application.PathSeparator & stPrefix & "_" _
			& Format(Time, "hhnn") & stOriginalName & stSuffix & stOriginalExtension
	End If
	dcNewDocument.SaveAs2 FileName:=stNewFullName, FileFormat:= wdFormatDocumentDefault

	If bClose = True Then
		dcNewDocument.Close
	Else
		Set SaveAsNewFile = dcNewDocument
	End If
End Function





Sub HeadersFootersRemove(dcArgument As Document)
' Borra todos los pies y encabezados de página
'
	Dim scCurrent As Section, hfCurrentHF As HeaderFooter

	Application.ScreenUpdating = False
	For Each scCurrent In dcArgument.Sections
		For Each hfCurrentHF In scCurrent.Headers
			If hfCurrentHF.Exists Then hfCurrentHF.Range.Delete
		Next hfCurrentHF

		For Each hfCurrentHF In scCurrent.Footers
			If hfCurrentHF.Exists Then hfCurrentHF.Range.Delete
		Next hfCurrentHF
	Next scCurrent
	Application.ScreenUpdating = True
End Sub





Sub ListsNoExtraNumeration(dcArgument As Document, Optional ByVal iStory As Integer = 0)
' Deletes lists' manual numerations
'
	Dim iMaxCount As Integer
	Dim stPatron As String
	Dim rgexNumeration As RegExp
	Dim rgStory As Range, rgListRng As Range
	Dim lpList As Paragraph

	If iStory < 0 Or iStory > 5 Then
		Err.Raise Number:=514, Description:="Argument out of range it must be between 0 and 5"
	ElseIf iStory = 0 Then
		iStory = 1
		iMaxCount = 5
	Else
		iMaxCount = iStory
	End If

	Set rgexNumeration = New RegExp
	' stPatron = "^[a-zA-Z0-9]{1,2}[\.\)\-ºª]+(?:[a-zA-Z0-9]{1,2}[\.\)\-ºª]?)*[\s]*"
		' VBA no pilla los caractéres finales del siguiente, tendría que buscar sus códigos
	' stPatron = "^(?:[a-zA-Z0-9]{1,2}[\.\)\-ºª]+(?:[a-zA-Z0-9]{1,2}[\.\)\-ºª]?)*|[–\-—•⁎⁕▪▸◂◃▷◼◻●◌◇◆°])[\s]*"
	stPatron = "^(?:[a-zA-Z0-9]{1,2}[\.\)\-ºª]+(?:[a-zA-Z0-9]{1,2}[\.\)\-ºª]?)*|[–\-—•])[\s]*"
	rgexNumeration.Pattern = stPatron
	rgexNumeration.IgnoreCase = True
	rgexNumeration.Global = False

	Application.ScreenUpdating = False
	For iStory = iStory To iMaxCount Step 1
		On Error Resume Next
		Set rgStory = dcArgument.StoryRanges(iStory)
		If Err.Number = 0 Then
			On Error GoTo 0
			For Each lpList In rgStory.ListParagraphs
				Set rgListRng = lpList.Range
				If rgListRng.Characters.Count > 8 Then
					rgListRng.End = rgListRng.Start + 8
				End If
				If rgexNumeration.Test(rgListRng.Text) Then
					rgListRng.End = rgListRng.End - Len(rgexNumeration.Replace(rgListRng.Text, ""))
					rgListRng.Delete
				End If
			Next lpList
		Else
			On Error GoTo 0
		End If
	Next iStory
	Application.ScreenUpdating = True
End Sub






Sub FindAndReplaceClearParameters(Optional bDummy As Boolean)
' Limpia los cuadros de búsqueda y reemplazo.
' Útil para llamarla después de automatizar búsquedas
	' https://wordmvp.com/FAQs/MacrosVBA/ClearFind.htm
'
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





Sub CleanBasic(dcArgument As Document, Optional ByVal iStory As Integer = 0)
' CleanSpaces + CleanEmptyParagraphs
' It's important to execute the subroutines in the proper order to achieve their optimal effects
' Args:
	' iStory: defines the storyranges that will be cleaned
		' All (1 to 5)		0
		' wdMainTextStory	1
		' wdFootnotesStory	2
		' wdEndnotesStory	3
		' wdCommentsStory	4
		' wdTextFrameStory	5
'
	If iStory < 0 Or iStory > 5 Then
		Err.Raise Number:=514, Description:="iStory out of range it must be between 0 and 5"
	End If
	RaMacros.CleanSpaces dcArgument, iStory
	RaMacros.CleanEmptyParagraphs dcArgument, iStory
	RaMacros.FindAndReplaceClearParameters
End Sub

Sub CleanSpaces(dcArgument As Document, Optional ByVal iStory As Integer = 0)
' Deletes:
	' Tabulations
	' More than 1 consecutive spaces
	' Spaces just before paragraph marks, stops, parenthesis, etc.
	' Spaces just after paragraph marks
' Args:
	' iStory: defines the storyranges that will be cleaned
		' All (1 to 5)		0
		' wdMainTextStory	1
		' wdFootnotesStory	2
		' wdEndnotesStory	3
		' wdCommentsStory	4
		' wdTextFrameStory	5
'
	Dim bFound1 As Boolean, bFound2 As Boolean, iMaxCount As Integer 
	Dim rgFindRange As Range, rgFindRange2 As Range, tbCurrentTable As Table

	If iStory < 0 Or iStory > 5 Then
		Err.Raise Number:=514, Description:="Argument out of range it must be between 0 and 5"
	ElseIf iStory = 0 Then
		iStory = 1
		iMaxCount = 5
	Else
		iMaxCount = iStory
	End If

	bFound1 = False
	bFound2 = False

	Application.ScreenUpdating = False
	For iStory = iStory To iMaxCount Step 1
		On Error Resume Next
		Set rgFindRange = dcArgument.StoryRanges(iStory)
		If Err.Number = 0 Then
			On Error GoTo 0

			' Deletting first and last characters if necessary
			Set rgFindRange2 = rgFindRange.Duplicate
			rgFindRange2.Collapse wdCollapseStart
			Do While rgFindRange.Characters.First.Text = " " _
					Or rgFindRange.Characters.First.Text = vbTab _
					Or rgFindRange.Characters.First.Text = "," _
					Or rgFindRange.Characters.First.Text = "." _
					Or rgFindRange.Characters.First.Text = ";" _
					Or rgFindRange.Characters.First.Text = ":"
				If rgFindRange2.Delete = 0 Then Exit Do
			Loop
			Set rgFindRange2 = rgFindRange.Duplicate
			rgFindRange2.Collapse wdCollapseEnd
			rgFindRange2.MoveStart wdCharacter, -1
			Do While rgFindRange2.Text = " " _
					Or rgFindRange.Characters.Last.Text = vbTab
				rgFindRange2.Collapse wdCollapseStart
				If rgFindRange2.Delete = 0 Then Exit Do
				rgFindRange2.MoveStart wdCharacter, -1
			Loop

			With rgFindRange.Find
				.ClearFormatting
				.Replacement.ClearFormatting
				.Forward = True
				.Wrap = wdFindStop
				.Format = False
				.MatchCase = False
				.MatchWholeWord = False
				.MatchAllWordForms = False
				.MatchSoundsLike = False
				.MatchWildcards = True
			End With
			Do
				With rgFindRange.Find
					.Text = "[^t]"
					.Replacement.Text = " "
					.Execute Replace:=wdReplaceAll
					If .Found Then bFound1 = True

					.Text = " {2;}"
					.Execute Replace:=wdReplaceAll
					If .Found Then bFound1 = True
				End With

				' Deletting spaces before paragraph marks before tables (there is a bug that prevents
					' them to be erased through find and replace)
				For each tbCurrentTable In rgFindRange.Tables
					If tbCurrentTable.Range.Start <> 0 Then
						Set rgFindRange2 = tbCurrentTable.Range.Previous(wdParagraph,1)
						rgFindRange2.MoveEnd wdCharacter, -1
						rgFindRange2.Start = rgFindRange2.End - 1
						If rgFindRange2.Text = " " Then
							bFound2 = False
							Do While rgFindRange2.Previous(wdCharacter, 1).Text = " "
								rgFindRange2.Start = rgFindRange2.Start - 1
								bFound2 = True
							Loop
							If bFound2 Then rgFindRange2.Delete
							rgFindRange2.Collapse wdCollapseStart
							rgFindRange2.Delete
						End If

						Set rgFindRange2 = tbCurrentTable.Range.Next(wdParagraph,1).Characters.First
							rgFindRange2.collapse wdCollapseStart
						Do While rgFindRange2.Text = " "
							If rgFindRange2.Delete = 0 Then Exit Do
						Loop
					End If
				Next tbCurrentTable
				
				bFound1 = False
				With rgFindRange.Find
					If iStory <> 2 Then
						.Text = " @([^13^l,.;:\]\)\}])"
						.Replacement.Text = "\1"
						.Execute Replace:=wdReplaceAll
						If .Found Then bFound1 = True
					Else
						Set rgFindRange2 = rgFindRange.Duplicate
						Do While rgFindRange2.Find.Execute( _
														FindText:=" @[^13^l,.;:\]\)\}]", _
														MatchWildcards:=True, Wrap:=wdFindStop)
							Do While rgFindRange2.Characters.First = " "
								rgFindRange2.Collapse wdCollapseStart
								rgFindRange2.Delete
							Loop
							rgFindRange2.EndOf wdStory, wdExtend
						Loop
					End If

					If iStory <> 2 Then
						.Text = "([^13^l])[ ,.;:]@"
						.Execute Replace:=wdReplaceAll
						If .Found Then bFound1 = True
					Else
						Set rgFindRange2 = rgFindRange.Duplicate
						Do While rgFindRange2.Find.Execute( _
														FindText:="[^13^l][ ,.;:]@", _
														MatchWildcards:=True, Wrap:=wdFindStop)
							rgFindRange2.Collapse wdCollapseStart
							rgFindRange2.Move wdCharacter, 1
							Do While rgFindRange2.Characters.Last = " " _
									Or rgFindRange2.Characters.Last = "," _
									Or rgFindRange2.Characters.Last = "." _
									Or rgFindRange2.Characters.Last = ";" _
									Or rgFindRange2.Characters.Last = ":"
								If rgFindRange2.Delete = 0 Then Exit Do
							Loop
							rgFindRange2.EndOf wdStory, wdExtend
						Loop
					End If
				End With

				If iStory = 5 And Not bFound1 And Not rgFindRange.NextStoryRange Is Nothing Then
					Set rgFindRange = rgFindRange.NextStoryRange
					With rgFindRange.Find
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
					End With
					bFound1 = True
				End If
			Loop While bFound1
		Else
			On Error GoTo 0
		End If
	Next iStory
	Application.ScreenUpdating = True
End Sub

Sub CleanEmptyParagraphs(dcArgument As Document, Optional ByVal iStory As Integer = 0)
' Deletes empty paragraphs
' Args:
	' iStory: defines the storyranges that will be cleaned
		' All (1 to 5)		0
		' wdMainTextStory	1
		' wdFootnotesStory	2
		' wdEndnotesStory	3
		' wdCommentsStory	4
		' wdTextFrameStory	5
'
	Dim rgStory As Range, rgFind As Range
	Dim tbCurrentTable As Table
	Dim cllCurrentCell As Cell
	Dim bAutoFit As Boolean, bFound As Boolean, bWrap As Boolean
	Dim iMaxCount As Integer

	If iStory < 0 Or iStory > 5 Then
		Err.Raise Number:=514, Description:="Argument out of range it must be between 0 and 5"
	ElseIf iStory = 0 Then
		iStory = 1
		iMaxCount = 5
	Else
		iMaxCount = iStory
	End If

	Application.ScreenUpdating = False
	For iStory = iStory To iMaxCount Step 1
		On Error Resume Next
		Set rgStory = dcArgument.StoryRanges(iStory)
		If Err.Number = 0 Then
			On Error GoTo 0
			Do
				bFound = False

				' Deletting first and last paragraphs, if empty
				Do While rgStory.Paragraphs.First.Range.Text = vbCr
					If rgStory.Paragraphs.First.Range.Delete = 0 Then Exit Do
				Loop
				Do While rgStory.Paragraphs.Last.Range.Text = vbCr 
					If iStory = 2 Or iStory = 3 Or iStory = 4 Then
						If rgStory.Paragraphs.Last.Range.Previous(wdCharacter, 1).Delete = 0 Then Exit Do
					Else
						If rgStory.Paragraphs.Last.Range.Delete = 0 Then Exit Do
					End If
				Loop

				' Deletting empty paragraphs related to tables
				For each tbCurrentTable In rgStory.Tables
					bAutoFit = tbCurrentTable.AllowAutoFit
					tbCurrentTable.AllowAutoFit = False
					bWrap = tbCurrentTable.Rows.WrapAroundText
					tbCurrentTable.Rows.WrapAroundText = False
					
					' Deletting empty paragraphs before tables
					Do
						If tbCurrentTable.Range.Start <> 0 Then
							Set rgFind = tbCurrentTable.Range.Previous(wdParagraph,1)
							If rgFind.Text = vbCr Then
								If rgFind.Start = 0 Then
									If rgFind.Delete = 0 Then Exit Do
									Exit Do
								Else
									If rgFind.Previous(wdParagraph, 1).Tables.Count = 0 Then
										If rgFind.Delete = 0 Then Exit Do
										Set rgFind = tbCurrentTable.Range.Previous(wdParagraph,1)
									Else
										Exit Do
									End If
								End If
							Else
								Exit Do
							End If
						Else
							Exit Do
						End If
					Loop

					' Deletting empty paragraphs after tables
					Do
						If tbCurrentTable.Range.End <> rgStory.End Then
							Set rgFind = tbCurrentTable.Range.Next(wdParagraph,1)
							If rgFind.Text = vbCr Then
								If rgFind.End <> rgStory.End Then
									If rgFind.Next(wdParagraph, 1).Tables.Count = 0 Then
										If rgFind.Delete = 0 Then Exit Do
										Set rgFind = tbCurrentTable.Range.Next(wdParagraph,1)
									Else
										Exit Do
									End If
								Else
									If rgFind.Delete = 0 Then Exit Do
									Exit Do
								End If
							Else
								Exit Do
							End If
						Else
							Exit Do
						End If
					Loop

					' Deletting empty paragraphs inside non empty cell tables
					For Each cllCurrentCell In tbCurrentTable.Range.Cells
						If Len(cllCurrentCell.Range.Text) > 2 And _
								cllCurrentCell.Range.Characters(1).Text = vbCr Then
							cllCurrentCell.Range.Characters(1).Delete
						End If

						If Len(cllCurrentCell.Range.Text) > 2 And _
								Asc(Right$(cllCurrentCell.Range.Text, 3)) = 13 Then
							Set rgFind = cllCurrentCell.Range
							rgFind.MoveEnd Unit:=wdCharacter, Count:=-1
							rgFind.Characters.Last.Delete
						End If
					Next cllCurrentCell

					tbCurrentTable.AllowAutoFit = bAutoFit
					tbCurrentTable.Rows.WrapAroundText = bWrap
				Next tbCurrentTable

				With rgStory.Find
					.ClearFormatting
					.Replacement.ClearFormatting
					.Forward = True
					.Wrap = wdFindStop
					.Format = False
					.MatchCase = False
					.MatchWholeWord = False
					.MatchAllWordForms = False
					.MatchSoundsLike = False
					.MatchWildcards = True
				End With

				If iStory = 2 Or iStory = 3 Or iStory = 4 Then
					Do
						bFound = False
						Set rgFind = rgStory.Duplicate
						With rgFind.Find
							.ClearFormatting
							.Replacement.ClearFormatting
							.Forward = True
							.Wrap = wdFindStop
							.Format = False
							.MatchCase = False
							.MatchWholeWord = False
							.MatchAllWordForms = False
							.MatchSoundsLike = False
							.MatchWildcards = True
							' If iStory = 4 Then .Text = "[^13^l]{2;}" Else .Text = "[^13^l]@^13^2"
							.Text = "[^13^l]{2;}"
						End With
                        If rgFind.Find.Execute Then
							' ------------------- 1st version
							' If iStory = 4 Then
							' 	If Len(rgFind) <= 2 Then rgFind.Collapse wdCollapseStart
							' 	If rgFind.Delete <> 0 Then bFound = True
							' ElseIf rgFind.End <> 0 Then
							' 	rgFind.MoveEnd wdCharacter, -2
							' 	If rgFind.Delete <> 0 Then bFound = True
							' End If
							' ------------------- 2nd version (no loops, unlike 3rd version)
                            If Len(rgFind) = 2 Then
                                rgFind.Collapse wdCollapseStart
                                If rgFind.Delete <> 0 Then bFound = True
                            Else
                                If rgFind.Delete <> 0 Then
                                    rgFind.Collapse wdCollapseStart
                                    If rgFind.Delete <> 0 Then bFound = True
                                End If
                            End If
							' ------------------- 3rd version
							' bFound = True
							' rgFind.Collapse wdCollapseStart
							' Do While rgFind.Next(wdCharacter, 1).Text = vbCr
                            '     If rgFind.Delete = 0 Then Exit Do
							' Loop
                        End If
					Loop While bFound
				End If

				bFound = False
				With rgStory.Find
					.Replacement.Text = "\1"

					.Text = "([^13^l]){2;}"
					If .Execute(Replace:=wdReplaceAll) Then bFound = True

					.Text = "(^13)^l"
					If .Execute(Replace:=wdReplaceAll) Then bFound = True

					.Text = "(^l)^13"
					If .Execute(Replace:=wdReplaceAll) Then bFound = True
				End With

				If iStory = 5 And Not bFound And Not rgStory.NextStoryRange Is Nothing Then
					Set rgStory = rgStory.NextStoryRange
					With rgStory.Find
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
					End With
					bFound = True
				End If
			Loop While bFound
		Else
			On Error GoTo 0
		End If
	Next iStory
	Application.ScreenUpdating = True
End Sub





Sub HeadingsNoPunctuation(dcArgument As Document)
' Elimina los puntos finales de los títulos
'
	Dim titulo As Integer, signos(3) As String, signoActual As Integer
	titulo = -2
	signoActual = 0
	signos(0) = "."
	signos(1) = ","
	signos(2) = ";"
	signos(3) = ":"

	Application.ScreenUpdating = False
	With dcArgument.Range.Find
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
	RaMacros.FindAndReplaceClearParameters
	Application.ScreenUpdating = True
End Sub

Sub HeadingsNoNumeration(dcArgument As Document)
' Deletes headings' manual numerations
'
	Dim iTitulo As Integer, stPatron As String, rgexNumeracion As RegExp, rgFindRange As Range, bFound As Boolean

	Set rgexNumeracion = New RegExp
	stPatron = "^[a-zA-Z0-9]{1,2}[\.\)\-ºª]+(?:[a-zA-Z0-9]{1,2}[\.\)\-ºª]?)*[\s]*"
	rgexNumeracion.Pattern = stPatron
	rgexNumeracion.IgnoreCase = True
	rgexNumeracion.Global = False

	Application.ScreenUpdating = False
	RaMacros.FindAndReplaceClearParameters
	For iTitulo = -2 To -10 Step -1
		Set rgFindRange = dcArgument.Content
		Do
			bFound = False
			With rgFindRange.Find
				.ClearFormatting
				.Forward = True
				.Wrap = wdFindStop
				.MatchWildcards = False
				.Style = iTitulo
				.Text = ""
				If .Execute Then
					If rgexNumeracion.Test(rgFindRange.Text) Then
						rgFindRange.End = rgFindRange.End - Len(rgexNumeracion.Replace(rgFindRange.Text, ""))
						rgFindRange.Delete
					End If
					' Continue the find operation using range
					rgFindRange.Expand wdParagraph
					If rgFindRange.End <> dcArgument.Content.End Then
						Set rgFindRange = rgFindRange.Next(Unit:=wdParagraph, Count:=1)
						rgFindRange.EndOf wdStory, wdExtend
						bFound = True
					End If
				End If
			End With
		Loop While bFound
	Next iTitulo
	RaMacros.FindAndReplaceClearParameters
	Application.ScreenUpdating = True
End Sub

Sub HeadingsChangeCase(dcArgument As Document, ByVal iHeading As Integer, ByVal iCase As Integer)
' Changes the case for the heading selected. This subroutine transforms the text, it doesn't change the style option "All caps"
' Args:
	' dcArgument: the document to be changed
	' iHeading: the heading style to be changed. If 0 all headings will be processed
	' iCase: the desired case for the text. It can be one of the WdCharacterCase constants. Options:
		' wdLowerCase: 0
		' wdUpperCase: 1
		' wdTitleWord: 2
		' wdTitleSentence: 4
		' wdToggleCase: 5
'
	Dim iCurrentHeading As Integer, iLowerHeading As Integer, rgFindRange As Range, bFound As Boolean

	If iCase <> 0 And iCase <> 1 And iCase <> 2 And iCase <> 4 And iCase <> 5 Then
		Err.Raise Number:=515, Description:="Incorrect case argument"
	End If

	If iHeading >= 1 And iHeading <= 9 Then
		iHeading = iHeading - iHeading * 2 - 1
		iLowerHeading = iHeading
	ElseIf iHeading = 0 Then
		iHeading = wdstyleheading9
		iLowerHeading = wdstyleheading1
	Else
		Err.Raise Number:=514, Description:="Argument out of range it must be between 0 and 9"
	End If

	Application.ScreenUpdating = False
	For iCurrentHeading = iLowerHeading To iHeading Step -1
		Set rgFindRange = dcArgument.Content
		Do
			bFound = False
			With rgFindRange.Find
				.ClearFormatting
				.Replacement.ClearFormatting
				.Forward = True
				.Wrap = wdFindStop
				.Format = True
				.MatchCase = False
				.MatchWholeWord = False
				.MatchAllWordForms = False
				.MatchSoundsLike = False
				.MatchWildcards = False
                .style = iCurrentHeading
				.Text = ""
				If .Execute Then
					' Sometimes there's a bug that only allow to change from lower case
					' rgFindRange.Case = wdLowerCase
					' If iCase <> 0 Then rgFindRange.Case = iCase
					rgFindRange.Case = iCase

					rgFindRange.Expand wdParagraph
					If rgFindRange.End <> dcArgument.Content.End Then
						Set rgFindRange = rgFindRange.Next(Unit:=wdParagraph, Count:=1)
						rgFindRange.EndOf wdStory, wdExtend
						bFound = True
					End If					
				End If
			End With
		Loop While bFound
	Next
	Application.ScreenUpdating = True
End Sub






Sub HyperlinksFormatting(dcArgument As Document, ByVal iPurpose As Integer, _
						Optional ByVal iStory As Integer = 0)
' It cleans and format hyperlinks
' Args:
	' iPurpose: choose what is the aim of the subroutine:
		' 1: Applies the hyperlink style to all hyperlinks
		' 2: cleans the text showed so only the domain is left
		' 3: both
'
	Dim iMaxCount As Integer

	If iStory < 0 Or iStory > 5 Or iPurpose < 1 Or iPurpose > 3 Then
		Err.Raise Number:=514, Description:="Argument out of range"
	ElseIf iStory = 0 Then
		iStory = 1
		iMaxCount = 5
	Else
		iMaxCount = iStory
	End If

	Dim rgStory As Range
	Dim hlCurrentLink As Hyperlink

	If iPurpose = 2 Or iPurpose = 3 Then
		Dim stPatron As String, stResultadoPatron As String, rgexUrlRegEx As RegExp
		' stPatron = "(?:https?:(?://)?(?:www\.)?|//|www\.)?([a-zA-Z\-]+?\.[a-zA-Z\-\.]+)(?:/[\S]+|/)?"
		stPatron = _
		"(?:https?:(?://)?(?:www\.)?|//|www\.)?((?:[a-zA-Z]|[a-zA-Z][a-zA-Z\-]*?(?:[a-zA-Z]\.))+)(?:/[\S]+|/)?"
		Set rgexUrlRegEx = New RegExp
		rgexUrlRegEx.Pattern = stPatron
		rgexUrlRegEx.IgnoreCase = True
		rgexUrlRegEx.Global = True
	End If

	Application.ScreenUpdating = False
    For iStory = iStory To iMaxCount Step 1
        On Error Resume Next
        Set rgStory = dcArgument.StoryRanges(iStory)
        If Err.Number = 0 Then
            On Error GoTo 0
			For Each hlCurrentLink In rgStory.Hyperlinks
				If hlCurrentLink.Type = 0 Then
					If iPurpose = 1 Or iPurpose = 3 Then
						hlCurrentLink.Range.Style = wdStyleHyperlink
					End If
					If iPurpose = 2 Or iPurpose = 3 Then
						If rgexUrlRegEx.Test(hlCurrentLink.TextToDisplay) = True Then
							hlCurrentLink.TextToDisplay = rgexUrlRegEx.Replace(hlCurrentLink.TextToDisplay, "$1")
						End If
					End If
				End If
			Next hlCurrentLink
		Else
			On Error GoTo 0
		End If
	Next iStory
	Application.ScreenUpdating = True
End Sub





Sub ImagesToCenteredInLine(dcArgument As Document)
' Formatea más cómodamente las imágenes
	' Las convierte de flotantes a inline (de shapes a inlineshapes)
	' Impide que aparezcan deformadas (mismo % relativo al tamaño original en alto y ancho)
	' Las centra
	' Impide que superen el ancho de página
'
	Dim inlShape As InlineShape, shShape As Shape, sngRealPageWidth As Single, sngRealPageHeight As Single, _
		iIndex As Integer
	sngRealPageWidth = dcArgument.PageSetup.PageWidth - dcArgument.PageSetup.Gutter _
		- dcArgument.PageSetup.RightMargin - dcArgument.PageSetup.LeftMargin

	sngRealPageHeight = dcArgument.PageSetup.PageHeight _
		- dcArgument.PageSetup.TopMargin - dcArgument.PageSetup.BottomMargin _
		- dcArgument.PageSetup.FooterDistance - dcArgument.PageSetup.HeaderDistance

	' Se convierten todas de inlineshapes a shapes
	'For Each inlShape In dcArgument.InlineShapes
	'	If inlShape.Type = wdInlineShapePicture Then inlShape.ConvertToShape
	'Next inlShape
'
	'' Se les da el formato correcto
	'For Each shShape In dcArgument.Shapes
	'	With shShape
	'		If .Type = msoPicture Then
	'			shShape.LockAnchor = True
	'			.RelativeVerticalPosition = wdRelativeVerticalPositionParagraph
	'			With .WrapFormat
	'				.AllowOverlap = False
	'				.DistanceTop = 8
	'				.DistanceBottom = 8
	'				.Type = wdWrapTopBottom
	'			End With
	'			.ScaleHeight 1, msoTrue, msoScaleFromBottomRight
	'			.ScaleWidth 1, msoTrue, msoScaleFromBottomRight
	'			.LockAspectRatio = msoTrue
	'			If .Width > sngRealPageWidth Then .Width = sngRealPageWidth
	'			.Left = wdShapeCenter
	'			.Top = 8
	'		End If
	'	End With
	'Next shShape

	' Se convierten todas de shapes a inlineshapes
	' For Each shShape In dcArgument.Shapes
	'	 If shShape.Type = msoPicture Then shShape.ConvertToInlineShape
	' Next shShape


	' Se convierten todas de shapes a inlineshapes

	If dcArgument.Shapes.Count > 0 Then

		For iIndex = 1 To dcArgument.Shapes.Count
			With dcArgument.Shapes(iIndex)
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

				If dcArgument.Shapes.Count = 0 Then Exit For

			End With
		Next iIndex
	End If

	' Se les da el formato correcto
	For Each inlShape In dcArgument.InlineShapes
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






Sub QuotesStraightToCurly(dcArgument As Document)
' Cambia las comillas problemáticas (" y ') por comillas inglesas
	' Este método elimina las variables no configurables de Document.Autoformat
'
	Dim bSmtQt As Boolean
	bSmtQt = Options.AutoFormatAsYouTypeReplaceQuotes
	Options.AutoFormatAsYouTypeReplaceQuotes = True

	With dcArgument.Range.Find
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
	RaMacros.FindAndReplaceClearParameters
End Sub






Sub SectionBreakBeforeHeading(dcArgument As Document, _
							Optional ByVal bRespect = False, _
							Optional ByVal iWdSectionStart As Integer = 2, _
							Optional ByVal iHeading As Integer = 1)
' Inserts section breaks of the type assigned before each heading of the level selected
' Args:
	' dcArgument: the document to be changed
	' bRespect: respect the original section start type before the heading
	' iWdSectionStart: the kind of section break to insert
	' iHeading: heading style that will be found
'
	Dim iPageNumber As Integer, iWdBreakType As Integer, rgFindRange As Range, bFound As Boolean

	If iHeading >= 1 And iHeading <= 9 Then
		iHeading = iHeading - iHeading * 2 - 1
	Else
		Err.Raise Number:=514, Description:="Argument out of range it must be between 1 and 9"
	End If

	Select Case iWdSectionStart
		Case 0
			' wdSectionContinuous
			iWdBreakType = 3
		Case 1
			' wdSectionNewColumn
			iWdBreakType = 8
		Case 2
			' wdSectionNewPage
			iWdBreakType = 2
		Case 3
			' wdSectionEvenPage
			iWdBreakType = 4
		Case 4
			' wdSectionOddPage
			iWdBreakType = 5
		Case Else
			Err.Raise Number:=514, Description:="Argument out of range it must be between 0 and 4"
	End Select					

	Set rgFindRange = dcArgument.Content

	Application.ScreenUpdating = False
	Do
		bFound = False
		With rgFindRange.Find
			.ClearFormatting
			.Forward = True
			.Wrap = wdFindStop
			.Format = True
			.MatchCase = False
			.MatchWholeWord = False
			.MatchWildcards = False
			.MatchSoundsLike = False
			.MatchAllWordForms = False
			.Style = iHeading
			.Text = ""
			If .Execute Then
				If rgFindRange.Start <> rgFindRange.Sections(1).Range.Start Then
					If iWdSectionStart = 1 Or iWdSectionStart = 2 Then
						rgFindRange.Collapse wdCollapseStart
					End If
					rgFindRange.InsertBreak iWdBreakType
					Set rgFindRange = rgFindRange.Next(Unit:=wdParagraph, Count:=1)
					rgFindRange.Collapse Direction:=wdCollapseStart
				ElseIf bRespect = False _
						And rgFindRange.Start = rgFindRange.Sections(1).Range.Start _
						And	rgFindRange.Sections(1).PageSetup.SectionStart <> iWdSectionStart Then
					rgFindRange.Sections(1).PageSetup.SectionStart = iWdSectionStart
				End If
				' Continue the find operation using range
				rgFindRange.Expand wdParagraph
				If rgFindRange.End <> dcArgument.Content.End Then
					Set rgFindRange = rgFindRange.Next(Unit:=wdParagraph, Count:=1)
					rgFindRange.EndOf wdStory, wdExtend
					bFound = True
				End If
			End If
		End With
	Loop While bFound
	RaMacros.FindAndReplaceClearParameters
	Application.ScreenUpdating = True
End Sub

Sub SectionsFillBlankPages(dcArgument As Document, _
							Optional ByVal stFillerText As String = "", _
							Optional styFillStyle As Style)
' Puts a blank page before each even or odd section break
' Args: 
	' dcArgument: the document to be changed
	' stFillerText: an optional dummy string to fill the blank page
	' styFillStyle: style for the dummy text
'
	Dim iEvenOrOdd As Integer, scCurrent As Section, rgLastParagraph As Range

	Application.ScreenUpdating = False
	For Each scCurrent In dcArgument.Sections
		If scCurrent.index > 1 _
				And (scCurrent.PageSetup.SectionStart = 4 Or scCurrent.PageSetup.SectionStart = 3) _
			Then
			Set rgLastParagraph = dcArgument.Sections(scCurrent.index - 1).Range.Paragraphs.Last.Range
			iEvenOrOdd = scCurrent.PageSetup.SectionStart - 3

			' Search and deletion of manual page breaks before the section break
			Do While rgLastParagraph.Previous(wdParagraph, 1).Text = Chr(12)
				rgLastParagraph.Previous(wdParagraph, 1).Delete
			Loop

			' Insertion of blank pages
			If rgLastParagraph.Information(wdActiveEndAdjustedPageNumber) Mod 2 = iEvenOrOdd _
					Or (scCurrent.PageSetup.SectionStart = 3 _
						And rgLastParagraph.Information(wdActiveEndAdjustedPageNumber) = 1) _
				Then
				rgLastParagraph.InsertParagraphBefore
				rgLastParagraph.Paragraphs(1).style = wdStyleNormal
				rgLastParagraph.Collapse wdCollapseStart
				rgLastParagraph.InsertBreak 7
				' Insertion of filler text
				If stFillerText <> "" _
						And rgLastParagraph.Information(wdActiveEndAdjustedPageNumber) Mod 2 <> iEvenOrOdd _
						And Not (scCurrent.PageSetup.SectionStart = 3 _
								And rgLastParagraph.Information(wdActiveEndAdjustedPageNumber) = 1) _
					Then
					rgLastParagraph.InsertParagraph
					rgLastParagraph.Text = stFillerText

					If Not styFillStyle Is Nothing Then
						rgLastParagraph.style = styFillStyle
					End If
				End If
			End If
		End If
	Next scCurrent
	Application.ScreenUpdating = True
End Sub

Sub SectionsExportEachToFiles(dcArgument As Document, _
							Optional ByVal stPrefix As String, _
							Optional ByVal stSuffix As String = "-section_")
' Exports each section of the document to a separate file
'
	Dim iCounter As Integer, lStartingPage As Long, lFirstFootnote As Long, scCurrent As Section, dcNewDocument As Document

	lFirstFootnote = 0

	For Each scCurrent In dcArgument.Sections
		Set dcNewDocument = RaMacros.SaveAsNewFile(dcArgument, , stSuffix & scCurrent.index, False)

		' Delete all sections of new document except the one to be saved
		For iCounter = dcNewDocument.Sections.Count To 1 Step -1
			If iCounter <> scCurrent.index Then
				dcNewDocument.Sections(iCounter).Range.Delete
			End If
		Next iCounter

		' Delete section break and last empty paragraph
		If dcNewDocument.Sections.Count = 2 Then
			dcNewDocument.Sections(1).Range.Characters.Last.Delete
			dcNewDocument.Sections(1).Range.Characters.Last.Delete
		End If

		' Correct footnote starting number
		If scCurrent.Range.Footnotes.Count > 0 And scCurrent.Range.Footnotes.NumberingRule = wdRestartContinuous Then
			lFirstFootnote = scCurrent.Range.Footnotes(1).index
			If scCurrent.Range.Footnotes.StartingNumber = 1 Then
				dcNewDocument.Footnotes.StartingNumber = lFirstFootnote
			ElseIf lFirstFootnote = 1 Then
				dcNewDocument.Footnotes.StartingNumber = scCurrent.Range.Footnotes.StartingNumber
			Else
				lFirstFootnote = lFirstFootnote + scCurrent.Range.Footnotes.StartingNumber - 1
				dcNewDocument.Footnotes.StartingNumber = lFirstFootnote
			End If
			' This remembers the footnote index of the last section, in case the next has none, but BE AWARE
			' that inserting new footnotes in the exported files would require to readjust all following files!!!
			lFirstFootnote = scCurrent.Range.Footnotes.Count
			lFirstFootnote = scCurrent.Range.Footnotes(lFirstFootnote).index + 1
		ElseIf lFirstFootnote <> 0 And scCurrent.Range.Footnotes.NumberingRule = wdRestartContinuous Then
			dcNewDocument.Footnotes.StartingNumber = lFirstFootnote
		End If

		' Correct page starting number
		lStartingPage = scCurrent.Range.Characters(1).Information(wdActiveEndAdjustedPageNumber)
		dcNewDocument.Sections(1).Footers(wdHeaderFooterFirstPage).PageNumbers.RestartNumberingAtSection = True
		dcNewDocument.Sections(1).Footers(wdHeaderFooterFirstPage).PageNumbers.StartingNumber = lStartingPage

		dcNewDocument.Close wdSaveChanges
	Next
End Sub






Sub TablesExportToPdf(dcArgument As Document, Optional ByVal stSuffix As String = "Table ", _
	Optional ByVal bDelete As Boolean = False, Optional ByVal stReplacementText As String = "Link to ", _
	Optional ByVal bLink As Boolean = False, Optional ByVal stAddress As String, _
	Optional ByVal iStyle As Integer = wdStyleNormal, Optional ByVal iSize As Integer = 0)
' Export each table to a PDF file
' Args:
	' stSuffix: the suffix to append to the table title, if it hasn't any
	' bDelete: defines if the table should be replaced
	' stReplacementText: the replacement text before the table title
	' bLink: if true the replacement text will be a hyperlink pointing to the address of the pdf
	' stAddress: the path where the hyperlink will point.
		' The name of the file will be automatically added to the argument, BUT 
		' the last character of the path must be a path separator (\ or /)
		' If empty it will point to the destination of the exported pdf
	' iStyle: the paragraph style of the replacement text
	' iSize: the font size of the replacement text
'
	Dim iCounter As Integer
	Dim rgReplacement As Range
	Dim tbCurrent As Table
	Dim stDocName As String, stTableFullName As String

	stDocName = Left(dcArgument.Name, InStrRev(dcArgument.Name, ".") - 1)
	If bDelete And stAddress = vbNullString Then
		stAddress = dcArgument.Path & Application.PathSeparator
	End If

	Application.ScreenUpdating = False
	For iCounter = dcArgument.Tables.Count To 1 Step -1
		Set tbCurrent = dcArgument.Tables(iCounter)
		If tbCurrent.NestingLevel = 1 Then
			If tbCurrent.Title = vbNullString Then
				tbCurrent.Title = stSuffix & iCounter
			End If
			stTableFullName = stDocName & " " & tbCurrent.Title
			tbCurrent.Range.ExportAsFixedFormat2 _
				OutputFileName:=dcArgument.Path & Application.PathSeparator & stTableFullName, _
				ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, _
				OptimizeFor:=wdExportOptimizeForPrint, ExportCurrentPage:=False, _
				Item:=wdExportDocumentWithMarkup, IncludeDocProps:=True, _
				CreateBookmarks:= wdExportCreateNoBookmarks, DocStructureTags:=True, _
				BitmapMissingFonts:=False, UseISO19005_1:=False, OptimizeForImageQuality:=True
			If bDelete Then
				Set rgReplacement = tbCurrent.Range.Next(wdParagraph, 1)
				rgReplacement.InsertParagraphBefore
				rgReplacement.InsertParagraphBefore
				If bLink Then
					rgReplacement.Hyperlinks.Add Anchor:= rgReplacement.Paragraphs.First.Range, _
						Address:= stAddress & stTableFullName & ".pdf", _
						TextToDisplay:= stReplacementText & tbCurrent.Title
				Else
					rgReplacement.Paragraphs.First.Range.Text = stReplacementText & tbCurrent.Title
				End If
				rgReplacement.Paragraphs.First.Style = iStyle
				If iSize <> 0 Then rgReplacement.Paragraphs.First.Range.Font.Size = iSize
				tbCurrent.Delete
			End If
		End If
	Next iCounter
	Application.ScreenUpdating = True
End Sub








Sub FootnotesHangingIndentation (dcArgument As Document, _
								Optional ByVal sIndentation As Single = 0.5, _
								Optional ByVal iFootnoteStyle As Integer = wdStyleFootnoteText)
' Adds a tab to the beginning of each paragraph and footnote, so their indentation is hanging
' Args:
	' sIndentation: the position of the indented text in centimeters
	' iFootnoteStyle: to indicate a custom footnote style
'
	sIndentation = CentimetersToPoints(sIndentation)

	Application.ScreenUpdating = False
	With dcArgument.Styles(wdStyleFootnoteText).ParagraphFormat
		If .TabStops.Count > 0 Then
			Do Until .TabStops(1).Position >= sIndentation
				.TabStops(1).Clear
			Loop
            If .TabStops(1).Position >= CentimetersToPoints(0.75) Then
				.TabStops.Add sIndentation, 0, 0
			End If
		Else
			.TabStops.Add sIndentation, 0, 0
		End If
        .FirstLineIndent = -sIndentation
	End With

	With dcArgument.StoryRanges(2).Find
		.ClearFormatting
		.Replacement.ClearFormatting
		.Forward = True
		.Format = True
		.MatchCase = False
		.MatchWholeWord = False
		.MatchAllWordForms = False
		.MatchSoundsLike = False
		.MatchWildcards = True
		
		.Text = "(^2)[ ^t]@"
		.Replacement.Text = "\1^t"
		.Execute Replace:=wdReplaceAll
		.Text = "(^13)([!^2])"
		.Replacement.Text = "\1^t\2"
		.Execute Replace:=wdReplaceAll
	End With
	Application.ScreenUpdating = True
End Sub

Sub FootnotesSameNumberingRule (dcArgument As Document, _
								Optional ByVal iNumberingRule As Integer = 3, _
								Optional ByVal iStartingNumber As Integer = -501)
' Set the same footnotes numbering rule in all sections of the document
' Args:
	' iNumberingRule:
		' 3 (default): it gives the numbering rule of the first section to all others
		' 0: wdRestartContinuous
		' 1: wdRestartSection
		' 2: wdRestartPage: wdRestartPage
	' iStartingNumber: starting number of each section
		' -501: doesn't change anything
		' 0: copies the starting number of the first section in al the others
'
	If iNumberingRule < 0 Or iNumberingRule > 3 Then
		Err.Raise Number:=514, Description:="iNumberingRule must be between 0 and 3"
	End If
	If iStartingNumber < 1 And iStartingNumber <> -501 Then
		Err.Raise Number:=514, Description:="iStartingNumber cannot be below 0"
	End If

	Dim scCurrent As Section

	If iNumberingRule = 3 Then
		iNumberingRule = dcArgument.Sections(1).Range.Footnotes.NumberingRule
	End If
	If iStartingNumber = 0 Then
		iStartingNumber = dcArgument.Sections(1).Range.Footnotes.StartingNumber
	End If

	For Each scCurrent In dcArgument.Sections
		scCurrent.Range.Footnotes.NumberingRule = iNumberingRule
		If iStartingNumber <> -501 Then
			scCurrent.Range.Footnotes.StartingNumber = iStartingNumber
		End If
	Next scCurrent
End Sub