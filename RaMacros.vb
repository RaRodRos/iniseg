Option Explicit


Function StylesDeleteUnused(dcArgumentDocument As Document, _
							Optional bMsgBox As Boolean = False) As Long
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

	lTotalSCount = 0
	Do
		sCount = 0
		For Each oStyle In dcArgumentDocument.Styles
			'Only check out non-built-in styles
			If oStyle.BuiltIn = False Then
				If StyleInUse(oStyle.NameLocal, dcArgumentDocument) = False Then
					Application.OrganizerDelete Source:= dcArgumentDocument.FullName, _
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

	StylesDeleteUnused = lTotalSCount
End Function

Function StyleInUse(Styname As String, dcArgumentDocument As Document) As Boolean
' Del mismo desarrollador que StylesDeleteUnused
' Is Stryname used any of dcArgumentDocument's story
	Dim Stry As Range
	Dim Shp As Shape
	Dim txtFrame As TextFrame

	If Not dcArgumentDocument.Styles(Styname).InUse Then StyleInUse = False: Exit Function
	' check if Currently used in a story?

	For Each Stry In dcArgumentDocument.StoryRanges
		If StoryInUse(Stry, dcArgumentDocument) Then
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

Function StoryInUse(Stry As Range, dcArgumentDocument As Document) As Boolean
' Del mismo desarrollador que StylesDeleteUnused
' Note: this will mark even the always-existing stories as not in use if they're empty
	If Not Stry.StoryLength > 1 Then StoryInUse = False: Exit Function
	Select Case Stry.StoryType
		Case wdMainTextStory, wdPrimaryFooterStory, wdPrimaryHeaderStory: StoryInUse = True
		Case wdEvenPagesFooterStory, wdEvenPagesHeaderStory: StoryInUse = Stry.Sections(1).PageSetup.OddAndEvenPagesHeaderFooter = True
		Case wdFirstPageFooterStory, wdFirstPageHeaderStory: StoryInUse = Stry.Sections(1).PageSetup.DifferentFirstPageHeaderFooter = True
		Case wdFootnotesStory, wdFootnoteContinuationSeparatorStory: StoryInUse = dcArgumentDocument.Footnotes.Count > 1
		Case wdFootnoteSeparatorStory, wdFootnoteContinuationNoticeStory: StoryInUse = dcArgumentDocument.Footnotes.Count > 1
		Case wdEndnotesStory, wdEndnoteContinuationSeparatorStory: StoryInUse = dcArgumentDocument.Endnotes.Count > 1
		Case wdEndnoteSeparatorStory, wdEndnoteContinuationNoticeStory: StoryInUse = dcArgumentDocument.Endnotes.Count > 1
		Case wdCommentsStory: StoryInUse = dcArgumentDocument.Comments.Count > 1
		Case wdTextFrameStory: StoryInUse = dcArgumentDocument.Frames.Count > 1
		Case Else: StoryInUse = False ' Must be some new or unknown wdStoryType
	End Select
End Function






Sub StylesNoDirectFormatting(dcArgumentDocument As Document)
' Convierte los estilos directos de negritas y cursivas a los estilos Strong y Emphasis, respectivamente
'
	Dim iCounter As Integer, arrstStylesToApply(13) As WdBuiltinStyle
	dcArgumentDocument.Activate

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

	With dcArgumentDocument.Range.Find
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

End Sub






Sub CopySecurity(dcArgumentDocument As Document, _
							Optional stPrefix As String = "orig-", _
							Optional stSuffix As String)
' Copia el archivo original pasado y lo prefija y sufija con
'
	Dim fsFileSystem As Object, stOriginalName As String, stOriginalExtension As String, stNewFullName As String

	With dcArgumentDocument
		stOriginalName = Left(.Name, InStrRev(.Name, ".") - 1)
		stOriginalExtension = Right(.Name, Len(.Name) - InStrRev(.Name, ".") + 1)
		stNewFullName = .Path & Application.PathSeparator & stPrefix & stOriginalName & stSuffix & stOriginalExtension
	End With

	Set fsFileSystem = CreateObject("Scripting.FileSystemObject")
	fsFileSystem.CopyFile dcArgumentDocument.FullName, stNewFullName
End Sub





Function SaveAsNewFile(dcArgumentDocument As Document, _
								Optional stPrefix As String, _
								Optional stSuffix As String = "noSuffix", _
								Optional bClose As Boolean = True)
' Guarda una copia del documento pasado como argumento, manteniendo el original abierto y convirtiéndolo al formato actual
' Args:
	' stPrefix: string to prefix the new document's name
	' stSuffix: string to suffix the new document's name. By default it will add the current date
	' bClose: if True the new document is saved AND closed
'
	Dim stOriginalName As String, stOriginalExtension As String, stNewFullName As String, dcNewDocument As Document
	With dcArgumentDocument
		stOriginalName = Left(.Name, InStrRev(.Name, ".") - 1)
		stOriginalExtension = Right(.Name, Len(.Name) - InStrRev(.Name, ".") + 1)

		If stSuffix = "noSuffix" Then stSuffix = "-" & RaMacros.GetFormattedDateAndTime(1)

		stNewFullName = .Path & Application.PathSeparator & stPrefix & stOriginalName & stSuffix & stOriginalExtension
		Set dcNewDocument = Documents.Add(.FullName)
	End With

	' IF THE FILE GETS CONVERTED TO THE LATEST VERSION IT MESSES SOME FIELDS
		' (INCLUDEPICTURE, particularly), so it's better to do it later
	'If dcArgumentDocument.CompatibilityMode < 15 Then dcArgumentDocument.Convert

	dcNewDocument.SaveAs2 FileName:=stNewFullName, FileFormat:= wdFormatDocumentDefault

	If bClose = True Then
		dcNewDocument.Close
	Else
		Set SaveAsNewFile = dcNewDocument
	End If

End Function





Function GetFormattedDateAndTime(Optional chosedInfo As Integer = 1) As String
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





Sub HeadersFootersRemove(dcArgumentDocument As Document)
' Borra todos los pies y encabezados de página

' https://word.tips.net/T001777_Deleting_All_Headers_and_Footers.html

	Dim oSec As Section
	Dim oHead As HeaderFooter
	Dim oFoot As HeaderFooter
	For Each oSec In dcArgumentDocument.Sections
		For Each oHead In oSec.Headers
			If oHead.Exists Then oHead.Range.Delete
		Next oHead

		For Each oFoot In oSec.Footers
			If oFoot.Exists Then oFoot.Range.Delete
		Next oFoot
	Next oSec
End Sub





Sub ListsToText(dcArgumentDocument As Document)
' Convierte todas las viñetas de las listas a texto

' https://wordmvp.com/FAQs/Numbering/ListString.htm
' https://word.tips.net/T001857_Converting_Lists_to_Text.html

	Dim lp As Paragraph
	For Each lp In dcArgumentDocument.ListParagraphs
		lp.Range.ListFormat.ConvertNumbersToText
	Next lp

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





Sub CleanBasic(dcArgumentDocument As Document)
' CleanEmptyParagraphs + CleanSpaces
'
	Application.ScreenUpdating = False

	RaMacros.CleanEmptyParagraphs dcArgumentDocument
	RaMacros.CleanSpaces dcArgumentDocument
	RaMacros.FindAndReplaceClearParameters

	Application.ScreenUpdating = True

End Sub

Sub CleanEmptyParagraphs(dcArgumentDocument As Document)
' Deletes empty paragraphs
	' First and last empty paragraphs: https://wordmvp.com/FAQs/MacrosVBA/DeleteEmptyParas.htm
'
	Dim rgActiveRange As Range

	Set rgActiveRange = dcArgumentDocument.Paragraphs(1).Range

	If rgActiveRange.Text = vbCr Then rgActiveRange.Delete

	Set rgActiveRange = dcArgumentDocument.Paragraphs.Last.Range
	If rgActiveRange.Text = vbCr Then rgActiveRange.Delete

	RaMacros.FindAndReplaceClearParameters

	With dcArgumentDocument.Range.Find
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
		.Text = "[^13^l]{2;}"
		.Replacement.Text = "^13"
		.Execute Replace:=wdReplaceAll
	End With

	RaMacros.FindAndReplaceClearParameters
End Sub

Sub CleanSpaces(dcArgumentDocument As Document)
' Deletes:
	' Tabulations
	' More than 1 consecutive spaces
	' Spaces just before paragraph marks, stops, parenthesis, etc.
	' Spaces just after paragraph marks
'
	Dim bFound As Boolean, rgFindRange As Range

	With dcArgumentDocument.Content.Find
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
	Do
		bFound = False

		Set rgFindRange = dcArgumentDocument.Content
		With rgFindRange.Find
			.Text = "[^t]"
			.Replacement.Text = " "
			.Execute Replace:=wdReplaceAll
			If .Found Then
				bFound = True
				Set rgFindRange = dcArgumentDocument.Content
			End If
		End With

		Set rgFindRange = dcArgumentDocument.Content
		With rgFindRange.Find
			.Text = "[ ]{2;}"
			.Execute Replace:=wdReplaceAll
			If .Found Then
				bFound = True
				Set rgFindRange = dcArgumentDocument.Content
			End If
		End With

		Set rgFindRange = dcArgumentDocument.Content
		With rgFindRange.Find
			.Text = "[ ]([^13^l,.;:\]\)\}])"
			.Replacement.Text = "\1"
			.Execute Replace:=wdReplaceAll
			If .Found Then
				bFound = True
				Set rgFindRange = dcArgumentDocument.Content
			End If
		End With

		Set rgFindRange = dcArgumentDocument.Content
		With rgFindRange.Find
			.Text = "([^13^l])[ ,.;:]"
			.Execute Replace:=wdReplaceAll
			If .Found Then
				bFound = True
				Set rgFindRange = dcArgumentDocument.Content
			End If
		End With
	Loop While bFound

	RaMacros.FindAndReplaceClearParameters

End Sub






Sub HeadingsNoPunctuation(dcArgumentDocument As Document)
' Elimina los puntos finales de los títulos
	' Se podría hacer con RegEx, pero no parece que valga la pena el consumo de recursos
'
	Dim titulo As Integer, signos(3) As String, signoActual As Integer
	titulo = -2
	signoActual = 0
	signos(0) = "."
	signos(1) = ","
	signos(2) = ";"
	signos(3) = ":"

	With dcArgumentDocument.Range.Find
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





Sub HeadingsNoNumeration(dcArgumentDocument As Document)
' Deletes headings' manual numerations
'
	Dim iTitulo As Integer, stPatron As String, rgexNumeracion As RegExp, rgFindRange As Range, bFound As Boolean

	Set rgexNumeracion = New RegExp
	stPatron = "^\d{1,2}\.(\d{1,2}\.?)*[\s]+"

	rgexNumeracion.Pattern = stPatron
	rgexNumeracion.IgnoreCase = True
	rgexNumeracion.Global = False

	RaMacros.FindAndReplaceClearParameters

	For iTitulo = -2 To -10 Step -1
		Set rgFindRange = dcArgumentDocument.Content
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
					If rgFindRange.End <> dcArgumentDocument.Content.End Then
						Set rgFindRange = rgFindRange.Next(Unit:=wdParagraph, Count:=1)
						rgFindRange.EndOf wdStory, wdExtend
						bFound = True
					End If
				End If
			End With
		Loop While bFound
	Next iTitulo

	RaMacros.FindAndReplaceClearParameters

End Sub





Sub HeadingsChangeCase(dcArgumentDocument As Document, iHeading As Integer, iCase As Integer)
' Changes the case for the heading selected. This subroutine transforms the text, it doesn't change the style option "All caps"
' Args:
	' dcArgumentDocument: the document to be changed
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

	For iCurrentHeading = iLowerHeading To iHeading
		Set rgFindRange = dcArgumentDocument.Content
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
				.Style = wdstyleheading1
				.Text = ""
				If .Execute Then
					rgFindRange.Case = wdLowerCase
					If iCase <> 0 Then rgFindRange.Case = iCase

					rgFindRange.Expand wdParagraph
					If rgFindRange.End <> dcArgumentDocument.Content.End Then
						Set rgFindRange = rgFindRange.Next(Unit:=wdParagraph, Count:=1)
						rgFindRange.EndOf wdStory, wdExtend
						bFound = True
					End If					
				End If
			End With
		Loop While bFound
	Next
End Sub






Sub HyperlinksOnlyDomain(dcArgumentDocument As Document)
' Limpia los hipervínculos para que limpien la URL completa y muestren solo el dominio
'

	Dim hlCurrent As Hyperlink, stPatron As String, stResultadoPatron As String, rgexUrlRegEx As RegExp
	stPatron = "(?:https?:(?://)?(?:www\.)?|//|www\.)([a-zA-Z\-]+?\.[a-zA-Z\-\.]+)(?:/[\S]+)?"
		' Este es más exacto (sin puntos o guiones a principio o final del dominio), pero VBA no permite lookbehinds
	' (?:https?:(?://)?(?:www\.)?|//|www\.)?((?:[a-zA-Z]|(?<=[a-zA-Z])-(?=[a-zA-Z]))+?\.(?:[a-zA-Z]|(?<=[a-zA-Z])[\.\-](?=[a-zA-Z]))+)(/[\S]+)?
	Set rgexUrlRegEx = New RegExp
	rgexUrlRegEx.Pattern = stPatron
	rgexUrlRegEx.IgnoreCase = True
	rgexUrlRegEx.Global = True

	For Each hlCurrent In dcArgumentDocument.Hyperlinks
		If hlCurrent.Type = 0 And rgexUrlRegEx.Test(hlCurrent.TextToDisplay) = True Then
			hlCurrent.TextToDisplay = rgexUrlRegEx.Replace(hlCurrent.TextToDisplay, "$1")
		End If
	Next hlCurrent

End Sub





Sub HyperlinksFormatting(dcArgumentDocument As Document)
' Aplica el estilo Hipervínculo a todos los hipervínculos
'
	dcArgumentDocument.Activate

	ActiveWindow.View.ShowFieldCodes = Not ActiveWindow.View.ShowFieldCodes

	With dcArgumentDocument.Range.Find
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

	RaMacros.FindAndReplaceClearParameters

	dcArgumentDocument.Activate
	ActiveWindow.View.ShowFieldCodes = Not ActiveWindow.View.ShowFieldCodes

End Sub





Sub ImagesToCenteredInLine(dcArgumentDocument As Document)
' Formatea más cómodamente las imágenes
	' Las convierte de flotantes a inline (de shapes a inlineshapes)
	' Impide que aparezcan deformadas (mismo % relativo al tamaño original en alto y ancho)
	' Las centra
	' Impide que superen el ancho de página
'
	Dim inlShape As InlineShape, shShape As Shape, sngRealPageWidth As Single, sngRealPageHeight As Single, _
		iIndex As Integer
	sngRealPageWidth = dcArgumentDocument.PageSetup.PageWidth - dcArgumentDocument.PageSetup.Gutter _
		- dcArgumentDocument.PageSetup.RightMargin - dcArgumentDocument.PageSetup.LeftMargin

	sngRealPageHeight = dcArgumentDocument.PageSetup.PageHeight _
		- dcArgumentDocument.PageSetup.TopMargin - dcArgumentDocument.PageSetup.BottomMargin _
		- dcArgumentDocument.PageSetup.FooterDistance - dcArgumentDocument.PageSetup.HeaderDistance

	' Se convierten todas de inlineshapes a shapes
	'For Each inlShape In dcArgumentDocument.InlineShapes
	'	If inlShape.Type = wdInlineShapePicture Then inlShape.ConvertToShape
	'Next inlShape
'
	'' Se les da el formato correcto
	'For Each shShape In dcArgumentDocument.Shapes
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
	' For Each shShape In dcArgumentDocument.Shapes
	'	 If shShape.Type = msoPicture Then shShape.ConvertToInlineShape
	' Next shShape


	' Se convierten todas de shapes a inlineshapes

	If dcArgumentDocument.Shapes.Count > 0 Then

		For iIndex = 1 To dcArgumentDocument.Shapes.Count
			With dcArgumentDocument.Shapes(iIndex)
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

				If dcArgumentDocument.Shapes.Count = 0 Then Exit For

			End With
		Next iIndex
	End If

	' Se les da el formato correcto
	For Each inlShape In dcArgumentDocument.InlineShapes
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






Sub QuotesStraightToCurly(dcArgumentDocument As Document)
' Cambia las comillas problemáticas (" y ') por comillas inglesas
	' Este método elimina las variables no configurables de Document.Autoformat
'
	Dim bSmtQt As Boolean
	bSmtQt = Options.AutoFormatAsYouTypeReplaceQuotes
	Options.AutoFormatAsYouTypeReplaceQuotes = True

	With dcArgumentDocument.Range.Find
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






Sub SectionBreakBeforeHeading(dcArgumentDocument As Document, _
						Optional bRespect = False, _
						Optional iWdSectionStart As Integer = 2, _
						Optional iHeading As Integer = 1)
' Inserts section breaks of the type assigned before each heading of the level selected
' Args:
	' dcArgumentDocument: the document to be changed
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

	Set rgFindRange = dcArgumentDocument.Content

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
				If rgFindRange.End <> dcArgumentDocument.Content.End Then
					Set rgFindRange = rgFindRange.Next(Unit:=wdParagraph, Count:=1)
					rgFindRange.EndOf wdStory, wdExtend
					bFound = True
				End If
			End If
		End With
	Loop While bFound
	RaMacros.FindAndReplaceClearParameters
End Sub






Sub SectionsFillBlankPages(dcArgumentDocument As Document, _
							Optional stFillerText As String = "", _
							Optional ByVal styFillStyle As Style)
' Puts a blank page before each even or odd section break
' Args: 
	' dcArgumentDocument: the document to be changed
	' stFillerText: an optional dummy string to fill the blank page
	' styFillStyle: style for the dummy text
'
	Dim iEvenOrOdd As Integer, scCurrentSection As Section, rgLastParagraph As Range

	For Each scCurrentSection In dcArgumentDocument.Sections
		If scCurrentSection.index > 1 _
				And (scCurrentSection.PageSetup.SectionStart = 4 Or scCurrentSection.PageSetup.SectionStart = 3) _
			Then
			Set rgLastParagraph = dcArgumentDocument.Sections(scCurrentSection.index - 1).Range.Paragraphs.Last.Range
			iEvenOrOdd = scCurrentSection.PageSetup.SectionStart - 3

			' Search and deletion of manual page breaks before the section break
			Do While rgLastParagraph.Previous(wdParagraph, 1).Text = Chr(12)
				rgLastParagraph.Previous(wdParagraph, 1).Delete
			Loop

			' Insertion of blank pages
			If rgLastParagraph.Information(wdActiveEndAdjustedPageNumber) Mod 2 = iEvenOrOdd _
					Or (scCurrentSection.PageSetup.SectionStart = 3 _
						And rgLastParagraph.Information(wdActiveEndAdjustedPageNumber) = 1) _
				Then
				rgLastParagraph.InsertParagraphBefore
				rgLastParagraph.Paragraphs(1).style = wdStyleNormal
				rgLastParagraph.Collapse wdCollapseStart
				rgLastParagraph.InsertBreak 7
				' Insertion of filler text
				If stFillerText <> "" _
						And rgLastParagraph.Information(wdActiveEndAdjustedPageNumber) Mod 2 <> iEvenOrOdd _
						And Not (scCurrentSection.PageSetup.SectionStart = 3 _
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
	Next scCurrentSection
End Sub






Sub SectionsExportEachToFiles(dcArgumentDocument As Document, _
								Optional stPrefix As String, _
								Optional stSuffix As String = "-section_")
' Exports each section of the document to a separate file
'
	Dim iCounter As Integer, lStartingPage As Long, lStartingFootnote As Long, scSection As Section, dcNewDocument As Document

	lStartingFootnote = 0

	For Each scSection In dcArgumentDocument.Sections
		Set dcNewDocument = RaMacros.SaveAsNewFile(dcArgumentDocument, , stSuffix & scSection.index, False)

		' Delete all sections of new document except the one to be saved
		For iCounter = dcNewDocument.Sections.Count To 1 Step -1
			If iCounter <> scSection.index Then
				dcNewDocument.Sections(iCounter).Range.Delete
			End If
		Next iCounter

		' Delete section break and last empty paragraph
		If dcNewDocument.Sections.Count = 2 Then
			dcNewDocument.Sections(1).Range.Characters.Last.Delete
			dcNewDocument.Sections(1).Range.Characters.Last.Delete
		End If

		' Correct footnote starting number
		If scSection.Range.Footnotes.Count > 0 Then
			lStartingFootnote = scSection.Range.Footnotes(1).index
			dcNewDocument.Footnotes.StartingNumber = lStartingFootnote
			' This remembers the footnote index of the last section, in case the next has none
			' Be aware that inserting new footnotes in the exported files would require to readjust
				' all the following files!!!
			lStartingFootnote = scSection.Range.Footnotes.Count
			lStartingFootnote = scSection.Range.Footnotes(lStartingFootnote).index + 1
		ElseIf lStartingFootnote <> 0 Then
			dcNewDocument.Footnotes.StartingNumber = lStartingFootnote
		End If

		' Correct page starting number
		lStartingPage = scSection.Range.Characters(1).Information(wdActiveEndAdjustedPageNumber)
		dcNewDocument.Sections(1).Footers(wdHeaderFooterFirstPage).PageNumbers.RestartNumberingAtSection = True
		dcNewDocument.Sections(1).Footers(wdHeaderFooterFirstPage).PageNumbers.StartingNumber = lStartingPage

		dcNewDocument.Close wdSaveChanges
	Next
End Sub






