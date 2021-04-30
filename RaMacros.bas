' Attribute VB_Name = "RaMacros"
Option Explicit

Function RangeInField(dcArgument As Document, rgArgument As Range) As Boolean
' Returns true if rgArgument is part of a field
'
	Dim fcurrent As Field
	Dim i As Integer

	For Each fCurrent in dcArgument.Fields
		If rgArgument.InRange(fCurrent.Result) Then
			RangeInField = True
			Exit Function
		End If
	Next fCurrent
	RangeInField = False
End Function






Function StylesDeleteUnused(dcArgument As Document, _
							Optional ByVal bMsgBox As Boolean = False) As Long
' Deletes unused styles using multiple loops to respect their hierarchy 
	' (avoiding the deletion of fathers without use, like in lists)
' Based on:
	' https://word.tips.net/T001337_Removing_Unused_Styles.html
' Modifications:
	' Renamed variables
	' It runs until no unused styles left
	' A message with the number of styles must be turn on by the bMsgBox parameter
	' It's now a function that returns the number of deleted styles
	' If the style cannot be found because the NameLocal property is corrupted
		' (eg. because of leading or trailing spaces) it gets automatically deleted
	' It now detects textframes in shapes or inline shapes
'
	Dim stCurrent As Style
	Dim lCount As Long, lTotalCount As Long
	Dim sStart As Single

	sStart = Timer
	lTotalCount = 0
	Do
		lCount = 0
		For Each stCurrent In dcArgument.Styles
			'Only check out non-built-in styles
			If stCurrent.BuiltIn = False Then
				If StyleInUse(stCurrent.NameLocal, dcArgument) = False Then
					Application.OrganizerDelete Source:= dcArgument.FullName, _
					Name:= stCurrent.NameLocal, Object:=wdOrganizerObjectStyles
					lCount = lCount + 1
				End If
			End If
		Next stCurrent
		lTotalCount = lTotalCount + lCount
	Loop While lCount > 0

	If bMsgBox Then 
		MsgBox lTotalCount & " styles deleted"
		StylesDeleteUnused = lTotalCount
	End If

	Debug.Print lTotalCount & " styles erased in " & CInt((Timer-sStart)/60) _
		& " minutes (" & Format(Timer - sStart, "0") & " seconds)"
	StylesDeleteUnused = lTotalCount
End Function

Function StyleInUse(ByVal stStyName As String, dcArgument As Document) As Boolean
' Del mismo desarrollador que StylesDeleteUnused
' Is Stryname used any of dcArgument's story
	Dim rgStory As Range
	Dim Shp As Shape
	Dim txtFrame As TextFrame

	On Error Resume Next
	If Not dcArgument.Styles(stStyName).InUse Then
		StyleInUse = False
		Exit Function
	End If
	On Error GoTo 0
	' check if Currently used in a story
	For Each rgStory In dcArgument.StoryRanges
		If StoryInUse(dcArgument, rgStory) Then
			If StyleInUseInRangeText(rgStory, stStyName) Then 
				StyleInUse = True
				Exit Function
			End If
			For Each Shp In rgStory.ShapeRange
				Set txtFrame = Shp.TextFrame
				If Not txtFrame Is Nothing Then
					If txtFrame.HasText Then
						If txtFrame.TextRange.Characters.Count > 1 Then
							If StyleInUseInRangeText(txtFrame.TextRange, stStyName) Then
								StyleInUse = True
								Exit Function
							End If
						End If
					End If
				End If
			Next Shp
		End If
	Next rgStory
	StyleInUse = False ' Not currently in use.
End Function

Function StyleInUseInRangeText(rng As Range, ByVal stStyName As String) As Boolean
' Del mismo desarrollador que StylesDeleteUnused
' Returns True if "stStyName" is use in rng
	With rng.Find
		.ClearFormatting
		.ClearHitHighlight
		.Style = stStyName
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
'
	Dim sh As Shape, iSh As InlineShape

	If Not Stry.StoryLength > 1 Then
		StoryInUse = False
		Exit Function
	End If
	Select Case Stry.StoryType
		Case wdMainTextStory, wdPrimaryFooterStory, wdPrimaryHeaderStory
			StoryInUse = True
		Case wdEvenPagesFooterStory, wdEvenPagesHeaderStory
			StoryInUse = Stry.Sections(1).PageSetup.OddAndEvenPagesHeaderFooter = True
		Case wdFirstPageFooterStory, wdFirstPageHeaderStory
			StoryInUse = Stry.Sections(1).PageSetup.DifferentFirstPageHeaderFooter = True
		Case wdFootnotesStory, wdFootnoteContinuationSeparatorStory
			StoryInUse = dcArgument.Footnotes.Count > 0
		Case wdFootnoteSeparatorStory, wdFootnoteContinuationNoticeStory
			StoryInUse = dcArgument.Footnotes.Count > 0
		Case wdEndnotesStory, wdEndnoteContinuationSeparatorStory
			StoryInUse = dcArgument.Endnotes.Count > 0
		Case wdEndnoteSeparatorStory, wdEndnoteContinuationNoticeStory
			StoryInUse = dcArgument.Endnotes.Count > 0
		Case wdCommentsStory
			StoryInUse = dcArgument.Comments.Count > 0
		Case wdTextFrameStory
			' StoryInUse = dcArgument.Frames.Count > 0
			If dcArgument.Frames.Count > 0 Then StoryInUse = True: Exit Function
			For Each sh In dcArgument.Shapes
				If sh.Type = msoTextBox Then StoryInUse = True: Exit Function
			Next sh
			For Each iSh In dcArgument.InlineShapes
				If iSh.Type = msoTextBox Then StoryInUse = True: Exit Function
			Next iSh
		Case Else
			StoryInUse = False ' Must be some new or unknown wdStoryType
	End Select
End Function






Function StyleExists(dcArgument As Document, stStyName As String) As Boolean
' Checks if styObjective exists in dcArgument and returns a boolean
'
	On Error GoTo NotExist
	If Not dcArgument.Styles(stStyName) Is Nothing Then StyleExists = True
	Exit Function
NotExist:
	StyleExists = False
End Function






Function StyleSubstitution(dcArgument As Document, _
						stOriginal As String, _
						stSubstitute As String, _
						Optional ByVal bDelete As Boolean _
	) As Integer
' Substitute one style with another one.
' Args:
	' stOriginal: name of the style to be substituted
	' stSubstitute: name of the substitute style 
	' bDelete: if True, stOriginal will be deleted
' Returns:
	' 0: all good
	' 1: stOriginal doesn't exist
	' 2: stSubstitute doesn't exist
	' 3: neither stOriginal nor stSubstitute exists
'

	If Not RaMacros.StyleExists(dcArgument, stOriginal) Then
		StyleSubstitution = 1
	End If
	If Not RaMacros.StyleExists(dcArgument, stSubstitute) Then
		StyleSubstitution = StyleSubstitution + 2
	End If
	If StyleSubstitution > 0 Then Exit Function

	Dim rgStory As Range

	For Each rgStory In dcArgument.StoryRanges
		Do
			With rgStory.Find
				.ClearFormatting
				.Replacement.ClearFormatting
				.Wrap = wdFindStop
				.Forward = True
				.Format = True
				.MatchCase = False
				.MatchWholeWord = False
				.MatchAllWordForms = False
				.MatchSoundsLike = False
				.MatchWildcards = False
				.Text = ""
				.Style = stOriginal
				.Replacement.Style = stSubstitute
				.Execute Replace:=wdReplaceAll
			End With
			Set rgStory = rgStory.NextStoryRange
		Loop Until rgStory Is Nothing
	Next rgStory

	If bDelete Then
		If dcArgument.Styles(stOriginal).BuiltIn Then
			Debug.Print stOriginal & " is a built in style and cannot be deleted"
		Else
			dcArgument.Styles(stOriginal).Delete
		End If
	End If
	RaMacros.FindAndReplaceClearParameters
	StyleSubstitution = 0
End Function






Sub StylesNoDirectFormatting(dcArgument As Document, _
							Optional rgArgument As Range, _
							Optional ByVal bUnderlineDelete As Boolean)
' Converts bold and italic direct style formatting into Strong and Emphasis
' Args:
	' rgArgument: if nothing the sub works over all the story ranges
	' bUnderlineDelete: if true all underlined text reverts to normal
'
	Dim bAllStories As Boolean
	Dim iCounter As Integer
	Dim rgFind As Range
	Dim stStylesToApply(13) As WdBuiltinStyle
	
	stStylesToApply(0) = wdStyleNormal
	stStylesToApply(1) = wdStyleCaption
	stStylesToApply(2) = wdStyleList
	stStylesToApply(3) = wdStyleList2
	stStylesToApply(4) = wdStyleList3
	stStylesToApply(5) = wdStyleListBullet
	stStylesToApply(6) = wdStyleListBullet2
	stStylesToApply(7) = wdStyleListBullet3
	stStylesToApply(8) = wdStyleListBullet3
	stStylesToApply(9) = wdStyleListBullet4
	stStylesToApply(10) = wdStyleListBullet5
	stStylesToApply(11) = wdStyleListNumber
	stStylesToApply(12) = wdStyleListNumber2
	stStylesToApply(13) = wdStyleListNumber3

	For Each rgFind In dcArgument.StoryRanges
		If Not rgArgument Is Nothing Then
			Set rgFind = rgArgument
			bAllStories = True
		End If
		Do
			With rgFind.Find
				.ClearFormatting
				.Text = ""
				.Replacement.Text = ""
				.Forward = True
				.Wrap = wdFindStop
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

				For iCounter = 0 To UBound(stStylesToApply)
					.Style = stStylesToApply(iCounter)
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

					.Font.Bold = False
					.Font.Italic = True
					.Replacement.Style = wdStyleEmphasis
					.Execute Replace:=wdReplaceAll

					If bUnderlineDelete Then
						.Font.Underline = True
						.Replacement.Font.Underline = False
						.Execute Replace:=wdReplaceAll
					End If
				Next iCounter

				.ClearFormatting
				.Text = "^f"
				.Replacement.Style = wdStyleFootnoteReference
				.Execute Replace:=wdReplaceAll
			End With

			If bAllStories Then Exit For
			Set rgFind = rgFind.NextStoryRange
		Loop Until rgFind Is Nothing
	Next rgFind
	RaMacros.FindAndReplaceClearParameters
End Sub






Sub CopySecurity(dcArgument As Document, _
				Optional ByVal stPrefix As String, _
				Optional ByVal stSuffix As String)
' Copies dcArgument adding the suffix and/or prefix passed as arguments. In case
' there are none, it appends a number
'
	Dim fsFileSystem As Object
	Dim stOriginalName As String, stExtension As String, stNewFullName As String
	Dim iCount As Integer

	stOriginalName = Left(dcArgument.Name, InStrRev(dcArgument.Name, ".") - 1)
	stExtension = Right(dcArgument.Name, Len(dcArgument.Name) - InStrRev(dcArgument.Name, ".") + 1)
	stNewFullName = dcArgument.Path & Application.PathSeparator & stPrefix _
		& stOriginalName & stSuffix & stExtension

	Do While Dir(stNewFullName) > ""
		stNewFullName = dcArgument.Path & Application.PathSeparator & stPrefix _
			& stOriginalName & stSuffix & "-" & Format(iCount, "00") & stExtension
		iCount = iCount + 1
	Loop

	Set fsFileSystem = CreateObject("Scripting.FileSystemObject")
	fsFileSystem.CopyFile dcArgument.FullName, stNewFullName
End Sub





Function SaveAsNewFile(dcArgument As Document, _
						Optional ByVal stPrefix As String, _
						Optional ByVal stSuffix As String, _
						Optional ByVal bOpen As Boolean = True, _
						Optional ByVal bCompatibility As Boolean = False)
' Guarda una copia del documento pasado como argumento, manteniendo el original abierto y convirtiéndolo al formato actual
' Args:
	' stPrefix: string to prefix the new document's name
	' stSuffix: string to suffix the new document's name. By default it will add the current date
	' bOpen: if True the new document stays open, if false it's saved AND closed
	' bCompatibility: if True the new document will be converted to the new Word Format
'
	Dim stOriginalName As String, stNewFullName As String, stExtension As String
	Dim dcNewDocument As Document

	stOriginalName = Left(dcArgument.Name, InStrRev(dcArgument.Name, ".") - 1)
	If stSuffix = vbNullString And stPrefix = vbNullString Then stSuffix = "-" & Format(Date, "yymmdd")

	stNewFullName = dcArgument.Path & Application.PathSeparator & stPrefix _
		& stOriginalName & stSuffix

	Set dcNewDocument = Documents.Add(dcArgument.FullName, Visible:=bOpen)

	If bCompatibility Then
		stExtension = ".docx"
		' IF THE FILE GETS CONVERTED TO THE LATEST VERSION THE FIELDS CAN GET MESSED UP
		' (INCLUDEPICTURE and EMBED particularly), so it may be better to closely watch the process
		If dcNewDocument.CompatibilityMode < 15 Then
			RaMacros.FieldsUnlink dcNewDocument
			dcNewDocument.Convert
		End If
	Else
		stExtension = Right(dcArgument.Name, Len(dcArgument.Name) - InStrRev(dcArgument.Name, ".") + 1)
	End If

	If Dir(stNewFullName & stExtension) > "" Then
		stNewFullName = dcArgument.Path & Application.PathSeparator & stPrefix & "_" _
			& Format(Time, "hhnn") & stOriginalName & stSuffix & stExtension
	End If

	If bCompatibility Then
		dcNewDocument.SaveAs2 FileName:=stNewFullName, FileFormat:= wdFormatDocumentDefault
	Else
		dcNewDocument.SaveAs2 FileName:=stNewFullName
	End If

	If bOpen Then
		Set SaveAsNewFile = dcNewDocument
	Else
		dcNewDocument.Close
	End If
End Function






Sub FieldsUnlink(dcArgument As Document)
' Unlinks included and embed fields so the images doesn't corrupt the file when it 
' gets updated from older (or different software) versions
'
	Dim iIndex As Integer
	For iIndex = dcArgument.Content.Fields.Count To 1 Step -1
		If dcArgument.Fields(iIndex).Type = wdFieldIncludePicture _
			Or dcArgument.Fields(iIndex).Type = wdFieldEmbed _
		Then
			dcArgument.Content.Fields(iIndex).Unlink
		End If
	Next iIndex
End Sub
		
		
		
		
		
		
Sub HeadersFootersRemove(dcArgument As Document)
' Borra todos los pies y encabezados de página
'
	Dim scCurrent As Section, hfCurrentHF As HeaderFooter

	For Each scCurrent In dcArgument.Sections
		For Each hfCurrentHF In scCurrent.Headers
			If hfCurrentHF.Exists Then hfCurrentHF.Range.Delete
		Next hfCurrentHF

		For Each hfCurrentHF In scCurrent.Footers
			If hfCurrentHF.Exists Then hfCurrentHF.Range.Delete
		Next hfCurrentHF
	Next scCurrent
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
		Err.Raise Number:=514, Description:="iStory out of range it must be between 0 and 5"
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





Sub CleanBasic(dcArgument As Document, Optional ByVal iStory As Integer = 0, _
	Optional ByVal bTabs As Boolean = True, Optional ByVal bBreakLines As Boolean = False)
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
	RaMacros.CleanSpaces dcArgument, iStory, bTabs
	RaMacros.CleanEmptyParagraphs dcArgument, iStory, bBreakLines
	RaMacros.FindAndReplaceClearParameters
End Sub

Sub CleanSpaces(dcArgument As Document, Optional ByVal iStory As Integer = 0, _
				Optional ByVal bTabs As Boolean = True)
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
	' bTabs: if True Tabs are substituted for a single space
'
	Dim bFound1 As Boolean, bFound2 As Boolean, iMaxCount As Integer 
	Dim rgFind As Range, rgFind2 As Range, tbCurrent As Table

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

	For iStory = iStory To iMaxCount Step 1
		On Error Resume Next
		Set rgFind = dcArgument.StoryRanges(iStory)
		If Err.Number = 0 Then
			On Error GoTo 0

			' Deletting first and last characters if necessary
			Set rgFind2 = rgFind.Duplicate
			rgFind2.Collapse wdCollapseStart
			Do While rgFind.Characters.First.Text = " " _
					Or rgFind.Characters.First.Text = vbTab _
					Or rgFind.Characters.First.Text = "," _
					Or rgFind.Characters.First.Text = "." _
					Or rgFind.Characters.First.Text = ";" _
					Or rgFind.Characters.First.Text = ":"
				If rgFind2.Delete = 0 Then Exit Do
			Loop
			Set rgFind2 = rgFind.Duplicate
			rgFind2.Collapse wdCollapseEnd
			rgFind2.MoveStart wdCharacter, -1
			Do While rgFind2.Text = " " _
					Or rgFind.Characters.Last.Text = vbTab
				rgFind2.Collapse wdCollapseStart
				If rgFind2.Delete = 0 Then Exit Do
				rgFind2.MoveStart wdCharacter, -1
			Loop

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
			End With
			Do
				With rgFind.Find
					.Replacement.Text = " "
					If bTabs Then	
						.Text = "[^t]"
						.Execute Replace:=wdReplaceAll
						If .Found Then bFound1 = True
					End IF

					.Text = " {2;}"
					.Execute Replace:=wdReplaceAll
					If .Found Then bFound1 = True
				End With

				' Deletting spaces before paragraph marks before tables (there is a bug that prevents
					' them to be erased through find and replace)
				For each tbCurrent In rgFind.Tables
					If tbCurrent.Range.Start <> 0 Then
						Set rgFind2 = tbCurrent.Range.Previous(wdParagraph,1)
						rgFind2.MoveEnd wdCharacter, -1
						rgFind2.Start = rgFind2.End - 1
						If rgFind2.Text = " " Then
							bFound2 = False
							Do While rgFind2.Previous(wdCharacter, 1).Text = " "
								rgFind2.Start = rgFind2.Start - 1
								bFound2 = True
							Loop
							If bFound2 Then rgFind2.Delete
							rgFind2.Collapse wdCollapseStart
							rgFind2.Delete
						End If

						Set rgFind2 = tbCurrent.Range.Next(wdParagraph,1).Characters.First
							rgFind2.collapse wdCollapseStart
						Do While rgFind2.Text = " "
							If rgFind2.Delete = 0 Then Exit Do
						Loop
					End If
				Next tbCurrent
				
				bFound1 = False
				With rgFind.Find
					If iStory <> 2 Then
						.Text = " @([^13^l,.;:\]\)\}])"
						.Replacement.Text = "\1"
						.Execute Replace:=wdReplaceAll
						If .Found Then bFound1 = True
					Else
						Set rgFind2 = rgFind.Duplicate
						Do While rgFind2.Find.Execute( _
														FindText:=" @[^13^l,.;:\]\)\}]", _
														MatchWildcards:=True, Wrap:=wdFindStop)
							Do While rgFind2.Characters.First = " "
								rgFind2.Collapse wdCollapseStart
								rgFind2.Delete
							Loop
							rgFind2.EndOf wdStory, wdExtend
						Loop
					End If

					If iStory <> 2 Then
						.Text = "([^13^l])[ ,.;:]@"
						.Execute Replace:=wdReplaceAll
						If .Found Then bFound1 = True
					Else
						Set rgFind2 = rgFind.Duplicate
						Do While rgFind2.Find.Execute( _
														FindText:="[^13^l][ ,.;:]@", _
														MatchWildcards:=True, Wrap:=wdFindStop)
							rgFind2.Collapse wdCollapseStart
							rgFind2.Move wdCharacter, 1
							Do While rgFind2.Characters.Last = " " _
									Or rgFind2.Characters.Last = "," _
									Or rgFind2.Characters.Last = "." _
									Or rgFind2.Characters.Last = ";" _
									Or rgFind2.Characters.Last = ":"
								If rgFind2.Delete = 0 Then Exit Do
							Loop
							rgFind2.EndOf wdStory, wdExtend
						Loop
					End If
				End With

				If iStory = 5 And Not bFound1 And Not rgFind.NextStoryRange Is Nothing Then
					Set rgFind = rgFind.NextStoryRange
					With rgFind.Find
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
End Sub

Sub CleanEmptyParagraphs(dcArgument As Document, Optional ByVal iStory As Integer = 0, _
						Optional ByVal bBreakLines As Boolean = False)
' Deletes empty paragraphs
' Args:
	' iStory: defines the storyranges that will be cleaned
		' All (1 to 5)		0
		' wdMainTextStory	1
		' wdFootnotesStory	2
		' wdEndnotesStory	3
		' wdCommentsStory	4
		' wdTextFrameStory	5
	' bBreakLines: manual break lines get converted to paragraph marks
'
	Dim rgStory As Range, rgFind As Range, rgBibliography As Range
	Dim tbCurrent As Table
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

	For iStory = iStory To iMaxCount Step 1
		On Error Resume Next
		Set rgStory = dcArgument.StoryRanges(iStory)
		If Err.Number = 0 Then
			On Error GoTo 0
			Do
				If bBreakLines Then
					With rgStory.Find
						.ClearFormatting
						.Replacement.ClearFormatting
						.Forward = True
						.Format = False
						.MatchCase = False
						.MatchWholeWord = False
						.MatchAllWordForms = False
						.MatchSoundsLike = False
						.MatchWildcards = False
						.Text = "^l"
						.Replacement.Text = "^p"
						.Execute Replace:=wdReplaceAll
					End With
				End If

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
				For each tbCurrent In rgStory.Tables
					' Check if the table is part of a field (it can get bugged)
					If Not RangeInField(dcArgument, tbCurrent.Range) Then
						bAutoFit = tbCurrent.AllowAutoFit
						tbCurrent.AllowAutoFit = False
						bWrap = tbCurrent.Rows.WrapAroundText
						tbCurrent.Rows.WrapAroundText = False
						
						' Deletting empty paragraphs before tables
						Do
							If tbCurrent.Range.Start <> 0 Then
								Set rgFind = tbCurrent.Range.Previous(wdParagraph,1)
								If rgFind.Text = vbCr Then
									If rgFind.Start = 0 Then
										If rgFind.Delete = 0 Then Exit Do
										Exit Do
									Else
										If rgFind.Previous(wdParagraph, 1).Tables.Count = 0 Then
											If rgFind.Delete = 0 Then Exit Do
											Set rgFind = tbCurrent.Range.Previous(wdParagraph,1)
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
							If tbCurrent.Range.End <> rgStory.End Then
								Set rgFind = tbCurrent.Range.Next(wdParagraph,1)
								If rgFind.Text = vbCr Then
									If rgFind.End <> rgStory.End Then
										If rgFind.Next(wdParagraph, 1).Tables.Count = 0 Then
											If rgFind.Delete = 0 Then Exit Do
											Set rgFind = tbCurrent.Range.Next(wdParagraph,1)
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
						For Each cllCurrentCell In tbCurrent.Range.Cells
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

						tbCurrent.AllowAutoFit = bAutoFit
						tbCurrent.Rows.WrapAroundText = bWrap
					End If
				Next tbCurrent

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
End Sub

Sub HeadingsNoNumeration(dcArgument As Document)
' Deletes headings' manual numerations
'
	Dim iTitulo As Integer, stPatron As String, rgexNumeracion As RegExp, rgFind As Range, bFound As Boolean

	Set rgexNumeracion = New RegExp
	stPatron = "^[a-zA-Z0-9]{1,2}[\.\)\-ºª]+(?:[a-zA-Z0-9]{1,2}[\.\)\-ºª]?)*[\s]*"
	rgexNumeracion.Pattern = stPatron
	rgexNumeracion.IgnoreCase = True
	rgexNumeracion.Global = False

	RaMacros.FindAndReplaceClearParameters
	For iTitulo = -2 To -10 Step -1
		Set rgFind = dcArgument.Content
		Do
			bFound = False
			With rgFind.Find
				.ClearFormatting
				.Forward = True
				.Wrap = wdFindStop
				.MatchWildcards = False
				.Style = iTitulo
				.Text = ""
				If .Execute Then
					If rgexNumeracion.Test(rgFind.Text) Then
						rgFind.End = rgFind.End - Len(rgexNumeracion.Replace(rgFind.Text, ""))
						rgFind.Delete
					End If
					' Continue the find operation using range
					rgFind.Expand wdParagraph
					If rgFind.End <> dcArgument.Content.End Then
						Set rgFind = rgFind.Next(Unit:=wdParagraph, Count:=1)
						rgFind.EndOf wdStory, wdExtend
						bFound = True
					End If
				End If
			End With
		Loop While bFound
	Next iTitulo
	RaMacros.FindAndReplaceClearParameters
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
	Dim iCurrentHeading As Integer, iLowerHeading As Integer, rgFind As Range, bFound As Boolean

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

	For iCurrentHeading = iLowerHeading To iHeading Step -1
		Set rgFind = dcArgument.Content
		Do
			bFound = False
			With rgFind.Find
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
					' so, if it's necessary, it gets converted
					rgFind.Case = wdLowerCase
					If iCase <> 0 Then rgFind.Case = iCase

					rgFind.Expand wdParagraph
					If rgFind.End <> dcArgument.Content.End Then
						Set rgFind = rgFind.Next(Unit:=wdParagraph, Count:=1)
						rgFind.EndOf wdStory, wdExtend
						bFound = True
					End If					
				End If
			End With
		Loop While bFound
	Next
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
	Dim iPageNumber As Integer, iWdBreakType As Integer, rgFind As Range, bFound As Boolean

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

	Set rgFind = dcArgument.Content

	Do
		bFound = False
		With rgFind.Find
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
				If rgFind.Start <> rgFind.Sections(1).Range.Start Then
					If iWdSectionStart = 1 Or iWdSectionStart = 2 Then
						rgFind.Collapse wdCollapseStart
					End If
					rgFind.InsertBreak iWdBreakType
					Set rgFind = rgFind.Next(Unit:=wdParagraph, Count:=1)
					rgFind.Collapse Direction:=wdCollapseStart
				ElseIf bRespect = False _
						And rgFind.Start = rgFind.Sections(1).Range.Start _
						And	rgFind.Sections(1).PageSetup.SectionStart <> iWdSectionStart Then
					rgFind.Sections(1).PageSetup.SectionStart = iWdSectionStart
				End If
				' Continue the find operation using range
				rgFind.Expand wdParagraph
				If rgFind.End <> dcArgument.Content.End Then
					Set rgFind = rgFind.Next(Unit:=wdParagraph, Count:=1)
					rgFind.EndOf wdStory, wdExtend
					bFound = True
				End If
			End If
		End With
	Loop While bFound
	RaMacros.FindAndReplaceClearParameters
End Sub

Function SectionGetFirstFootnoteNumber(dcArgument As Document, lIndex As Long) As Long
' Returns the number of the first footnote of the section or 0 if there is none
' Args:
	' lIndex: the index of the section containing the footnote
'
	Dim scCurrent As Section
	Dim lFirstFootnote As Long

	Set scCurrent = dcArgument.Sections(lIndex)

	If scCurrent.Range.Footnotes.Count > 0 Then
		lFirstFootnote = scCurrent.Range.Footnotes(1).index
		If scCurrent.Range.FootnoteOptions.NumberingRule = wdRestartContinuous Then
			If scCurrent.Range.FootnoteOptions.StartingNumber = 1 Then
				SectionGetFirstFootnoteNumber = lFirstFootnote
			ElseIf lFirstFootnote = 1 Then
				SectionGetFirstFootnoteNumber = scCurrent.Range.FootnoteOptions.StartingNumber
			Else
				SectionGetFirstFootnoteNumber = lFirstFootnote + scCurrent.Range.FootnoteOptions.StartingNumber - 1
			End If
		ElseIf scCurrent.Range.FootnoteOptions.NumberingRule = wdRestartSection Then
			SectionGetFirstFootnoteNumber = scCurrent.Range.FootnoteOptions.StartingNumber
		End If
	Else
		SectionGetFirstFootnoteNumber = 0
	End If
End Function

Function SectionsExportEachToFiles(dcArgument As Document, _
							Optional ByVal bClose As Boolean = True, _
							Optional ByVal bMaintainFootnotesNumeration As Boolean = True, _
							Optional ByVal bMaintainPagesNumeration As Boolean = True, _
							Optional ByVal stPrefix As String, _
							Optional ByVal stSuffix As String)
' Exports each section of the document to a separate file
' ToDo: if bClose false then devolver array con los documentos generados
	Dim iCounter As Integer
	Dim lStartingPage As Long, lFirstFootnote As Long
	Dim scCurrent As Section
	Dim dcNewDocument As Document

	lFirstFootnote = 1

	For Each scCurrent In dcArgument.Sections
		Set dcNewDocument = RaMacros.SaveAsNewFile(dcArgument, stPrefix, _
			stSuffix & scCurrent.index, True, False)

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
		If bMaintainFootnotesNumeration Then
			If scCurrent.Range.Footnotes.Count > 0 Then
				lFirstFootnote = RaMacros.SectionGetFirstFootnoteNumber(dcArgument, scCurrent.Index)
				dcNewDocument.Footnotes.StartingNumber = lFirstFootnote
				' This remembers the footnote index of the last section, in case the
				' next has none, but BE AWARE that inserting new footnotes in the
				' exported files would require to readjust all following files!!!
				lFirstFootnote = lFirstFootnote + dcNewDocument.Footnotes.Count
			Else
				dcNewDocument.Footnotes.StartingNumber = lFirstFootnote
			End If
		End If

		' Correct page starting number
		If bMaintainPagesNumeration Then
			lStartingPage = scCurrent.Range.Characters(1).Information(wdActiveEndAdjustedPageNumber)
			dcNewDocument.Sections(1).Footers(wdHeaderFooterFirstPage).PageNumbers.RestartNumberingAtSection = True
			dcNewDocument.Sections(1).Footers(wdHeaderFooterFirstPage).PageNumbers.StartingNumber = lStartingPage
		End If

		dcNewDocument.Close wdSaveChanges
	Next
End Function

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
End Sub







Sub TablesConvertToImage(dcArgument As Document, _
						Optional ByVal iPlacement As Integer = wdInLine)
' Convert each table to an inline image
' Solution to problems with clipboard (do loop) found in:
	' https://www.mrexcel.com/board/threads/excel-vba-inconsistent-errors-when-trying-to-copy-and-paste-objects-from-excel-to-word.1112368/post-5485704
' Args:
	' iPlacement: WdOLEPlacement enum
		' 0: wdInLine
		' 1: wdFloatOverText
'
	Dim iTable As Integer
	Dim iTry As Integer
	Dim rgTable As Range

	For iTable = dcArgument.Tables.Count To 1 Step -1
		If dcArgument.Tables(iTable).NestingLevel = 1 Then
			iTry = 0
			Do Until iTry = 10
                On Error GoTo 0
				iTry = iTry + 1
				Set rgTable = dcArgument.Tables(iTable).Range
				rgTable.CopyAsPicture
                DoEvents
				Set rgTable = rgTable.Previous(wdCharacter, 1)
				rgTable.Collapse wdCollapseStart
				If iTry < 10 Then On Error Resume Next
				rgTable.PasteSpecial DataType:=wdPasteEnhancedMetafile, Placement:=iPlacement
				If Err.Number = 0 Then
                    On Error GoTo 0
					Exit Do
				End If
			Loop
			dcArgument.Tables(iTable).Delete
		End If
	Next iTable
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
End Sub

Sub TablesKeepTogether(dcArgument As Document, _
						Optional ByVal iPlacement As Integer = wdInLine)
' Convert each table to an inline image
' Args:
	' iPlacement: WdOLEPlacement enum
		' 0: wdInLine
		' 1: wdFloatOverText
'
	Dim iTable As Integer
	Dim rgTable

	For iTable = dcArgument.Tables.Count To 1 Step -1
		If dcArgument.Tables(iTable).NestingLevel = 1 Then
			With dcArgument.Tables(iTable).Range
				.CopyAsPicture
				.Delete
				.PasteSpecial DataType:=wdPasteEnhancedMetafile, Placement:=iPlacement
			End With
		End If
	Next iTable
End Sub







Sub FootnotesFormatting(dcArgument As Document, _
						Optional stFootnotes As String, _
						Optional stFootnoteReferences As String)
' Applies styles to the footnotes story and the footnotes references
' Args:
	' stFootnotes: style for the body text. Default: wdStyleFootnoteText
	' styFootnoteReferences: style for the references. Default: stFootnoteReferences
'
	If stFootnotes = vbNullString Then
		stFootnotes = dcArgument.Styles(wdStyleFootnoteText).NameLocal
	ElseIf Not RaMacros.StyleExists(dcArgument, stFootnotes) Then
		Err.Raise Number:=517, Description:= stFootnotes & _
			" (stFootnotes) doesn't exist in " & dcArgument.Name
	End If
	If stFootnoteReferences = vbNullString Then
		stFootnoteReferences = dcArgument.Styles(wdStyleFootnoteReference).NameLocal
	ElseIf Not RaMacros.StyleExists(dcArgument, stFootnoteReferences) Then
		Err.Raise Number:=517, Description:= stFootnoteReferences & _
			" (stFootnoteReferences) doesn't exist in " & dcArgument.Name
	End If

	Dim i As Integer
	dcArgument.StoryRanges(2).Style = stFootnotes
	For i = 1 To 2
		With dcArgument.StoryRanges(i).Find
			.ClearFormatting
			.Replacement.ClearFormatting
			.Format = False
			.MatchCase = False
			.MatchWholeWord = False
			.MatchWildcards = False
			.MatchSoundsLike = False
			.MatchAllWordForms = False
			.Text = "^f"
			.Replacement.style = stFootnoteReferences
			.Execute Replace:=wdReplaceAll
		End With
	Next i
End Sub

Sub FootnotesHangingIndentation(dcArgument As Document, _
								Optional ByVal sIndentation As Single = 0.5, _
								Optional ByVal iFootnoteStyle As Integer = wdStyleFootnoteText)
' Adds a tab to the beginning of each paragraph and footnote, so their indentation is hanging
' Args:
	' sIndentation: the position of the indented text in centimeters
	' iFootnoteStyle: to indicate a custom footnote style
'
	sIndentation = CentimetersToPoints(sIndentation)

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
		
		Do
			.Text = "(^2)[ ^t]@"
			.Replacement.Text = "\1"
			.Execute Replace:=wdReplaceAll
		Loop While .Found
		.Text = "(^2)"
		.Replacement.Text = "\1^t"
		.Execute Replace:=wdReplaceAll
		.Text = "(^13)([!^2])"
		.Replacement.Text = "\1^t\2"
		.Execute Replace:=wdReplaceAll
	End With
End Sub

Sub FootnotesSameNumberingRule(dcArgument As Document, _
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
		' 0: copies the starting number of the first section in all the others
'
	If iNumberingRule < 0 Or iNumberingRule > 3 Then
		Err.Raise Number:=514, Description:="iNumberingRule must be between 0 and 3"
	End If
	If iStartingNumber < 0 And iStartingNumber <> -501 Then
		Err.Raise Number:=514, Description:="iStartingNumber not in range"
	End If

	Dim scCurrent As Section

	If iNumberingRule = 3 Then
		iNumberingRule = dcArgument.Sections(1).Range.FootnoteOptions.NumberingRule
	End If
	If iStartingNumber = 0 Then
		iStartingNumber = dcArgument.Sections(1).Range.FootnoteOptions.StartingNumber
	End If

	For Each scCurrent In dcArgument.Sections
		If iNumberingRule <> wdRestartContinuous Then
			scCurrent.Range.FootnoteOptions.StartingNumber = 1
		ElseIf iStartingNumber <> -501 Then
			scCurrent.Range.FootnoteOptions.StartingNumber = iStartingNumber
		End If
		scCurrent.Range.FootnoteOptions.NumberingRule = iNumberingRule
	Next scCurrent
End Sub






Function ClearHiddenText(dcArgument As Document, _
						Optional bDelete As Boolean, _
						Optional styWarning As Style, _
						Optional bMaintainHidden As Boolean, _
						Optional bShowHidden As Integer = 0) _
	As Integer()
' Deletes or apply a warning style to all hidden text in the document.
' Returns: array of integers of the story ranges containing hidden text
' Args:
	' bDelete: true deletes all hidden text
	' styWarning: defines the style for the hidden text
	' bMaintainHidden: if true the text maintains its hidden attribute
	' bShowHidden: changes if the hidden text is displayed
		' 0: maintains the current configuration
		' 1: hidden
		' 2: visible
'
	Dim rgStory As Range
	Dim iFound() As Integer, bShowOption As Boolean

	If bShowHidden < 0 Or bShowHidden > 2 Then
		Err.Raise Number:=514, Description:="bShowHidden out of range it must be between 0 and 2"
	ElseIf bShowHidden = 0 Then
		bShowOption = dcArgument.ActiveWindow.View.ShowHiddenText
	ElseIf bShowHidden = 1 Then
		bShowOption = False
	ElseIf bShowHidden = 2 Then
		bShowOption = True
	End If

	If Not bDelete And styWarning Is Nothing Then
		If Not StyleExists(dcArgument, "WarningHiddenText") Then
			Set styWarning = dcArgument.Styles.Add("WarningHiddenText", wdStyleTypeCharacter)
			styWarning.QuickStyle = True
			With styWarning.Font
				.Size = 31
				.ColorIndex = wdYellow
				.Shading.Texture = wdTextureNone
				.Shading.BackgroundPatternColorIndex = wdRed
				.Hidden = bMaintainHidden
			End With
		Else
			Set styWarning = dcArgument.Styles("WarningHiddenText")
		End If
	End If

	dcArgument.ActiveWindow.View.ShowHiddenText = True
	For Each rgStory In dcArgument.StoryRanges
		Do
			With rgStory.Find
				.ClearFormatting
				.Replacement.ClearFormatting
				.Forward = True
				.Format = True
				.MatchCase = False
				.MatchWholeWord = False
				.MatchAllWordForms = False
				.MatchSoundsLike = False
				.MatchWildcards = False
				.Font.Hidden = True
				.Text = ""
				If bDelete Then
					.Replacement.Text = ""
				Else 
					.Replacement.Style = styWarning
					.Replacement.Font.Hidden = bMaintainHidden
				End If
				.Execute Replace:=wdReplaceAll
				If .Found Then
					On Error GoTo EmptyiFound
					ReDim Preserve iFound(UBound(iFound) + 1)
					On Error GoTo 0
					iFound(UBound(iFound)) = rgStory.StoryType
				End If
			End With
			Set rgStory = rgStory.NextStoryRange
		Loop Until rgStory Is Nothing
	Next rgStory
	dcArgument.ActiveWindow.View.ShowHiddenText = bShowOption
	ClearHiddenText = iFound
	
	Exit Function
EmptyiFound:
    On Error GoTo 0
	ReDim iFound(0)
	Resume Next
End Function