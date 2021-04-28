Option Explicit

Sub uiInisegConversionLibro()
	ConversionLibro ActiveDocument
End Sub
Sub uiInisegConversionStory()
	ConversionStory ActiveDocument
End Sub
Sub uiInisegBibliografiaExportar()
	BibliografiaExportar ActiveDocument
End Sub






Sub Iniseg1Limpieza()
' Crea copia de seguridad del archivo original.
' Ejecuta limpieza (espacios, estilos innecesarios, etc.)
' Crea y deja abierto el archivo en formato libro para comenzar a darle los estilos
'
	Dim dcOriginalFile As Document, dcLibro As Document
	Dim rgRangoActual As Range
	Dim stFileName As String
	Dim iDeleteAnswer As Integer, lEstilosBorrados As Long, lPrimeraNotaAlPie As Long

	Set dcOriginalFile = ActiveDocument
	Set rgRangoActual = dcOriginalFile.Content
	stFileName = dcOriginalFile.FullName
	lPrimeraNotaAlPie = 0

	' Borrar contenido innecesario
	iDeleteAnswer = MsgBox("¿Borrar contenido hasta el punto seleccionado?", vbYesNoCancel, "Borrar contenido")
	If iDeleteAnswer = vbCancel Then Exit Sub

	rgRangoActual.Start = Selection.Start
	Debug.Print "1.1/13 - Haciendo copia de seguridad (0) del archivo original"
	RaMacros.CopySecurity dcOriginalFile, "0-", ""

	If iDeleteAnswer = vbYes Then
		If rgRangoActual.Footnotes.Count > 0 Then
			lPrimeraNotaAlPie = rgRangoActual.Footnotes(1).Index
		End If
		Debug.Print "1.2/13 - Borrando el texto seleccionado"
		rgRangoActual.End = rgRangoActual.Start
		rgRangoActual.Start = 0
		rgRangoActual.Delete
	End If

	' Actualización del formato del archivo (soluciona problemas de compatibilidad con shapes y campos)
	Debug.Print "1.3/13 - Actualizando formato de archivo"
	Iniseg.ActualizandoVersion dcOriginalFile

	Debug.Print "2/13 - Creando archivo con plantilla Iniseg"
	Set dcLibro = Documents.Add("C:\Users\Ra\Documents\Plantillas personalizadas de Office\iniseg.dotm")

	Debug.Print "3/13 - Copiando encabezados"
	Iniseg.HeaderCopy dcOriginalFile, dcLibro, 1

	Debug.Print "4/13 - Aplicando autoformateo"
	Iniseg.AutoFormateo dcOriginalFile
	Debug.Print "5.1/13 - Limpiando hiperenlaces para que solo figure su dominio"
	RaMacros.HyperlinksFormatting dcOriginalFile, 2, 0
	Debug.Print "5.2/13 - Limpiando espacios"
	RaMacros.CleanSpaces dcOriginalFile, 0

	Debug.Print "6/13 - Borrando encabezados y pies de página"
	RaMacros.HeadersFootersRemove dcOriginalFile

	Debug.Print "7/13 - Dando colores adecuados al texto"
	Iniseg.ColoresCorrectos dcOriginalFile

	Debug.Print "8/13 - Borrando estilos sin uso"
	lEstilosBorrados = RaMacros.StylesDeleteUnused(dcOriginalFile, False)

	' Copia de seguridad limpia
	Debug.Print "9/13 - Creando copia de seguridad limpia (01)"
	RaMacros.SaveAsNewFile dcOriginalFile, "01-", "", True

	' Guarda el archivo con nombre original, preparado para el siguiente paso
	Debug.Print "10.1/13 - Copiando contenido limpio al archivo con plantilla (archivo libro)"
	dcLibro.Content.FormattedText = dcOriginalFile.Content

	If lPrimeraNotaAlPie <> 0 Then
		Debug.Print "10.2/13 - Archivo libro: corrigiendo el número de comienzo de las notas al pie"
		dcLibro.Footnotes.StartingNumber = lPrimeraNotaAlPie
	End If

	Debug.Print "11/13 - Archivo original: cerrando"
	dcOriginalFile.Close wdDoNotSaveChanges
	Debug.Print "12/13 - Archivo libro: guardando"
	dcLibro.SaveAs2 stFileName
	dcLibro.Activate

	Debug.Print "13/13 - Iniseg1Limpieza terminada"
	Beep
	MsgBox lEstilosBorrados & " Estilos borrados" & vbCrLf _
		& "Revisar numeración de notas al pie, aplicar estilos y ejecutar Iniseg2"
End Sub

Sub Iniseg2LibroYStory()
' Llama a las macros de ConversionLibro e ConversionStory y da un aviso para seguir trabajando
	' Organizado de esta forma las macros de libro y story se pueden llamar por separado
'
	Dim dcLibro As Document, dcStory As Document, iExportar As Integer, iNotas As Integer

	If ActiveDocument.Footnotes.Count > 0 Then
		iNotas = MsgBox("¿Exportar notas al pie de página a archivo separado?", vbYesNoCancel, "Opciones exportar")
		If iNotas = vbCancel Then Exit Sub
	Else
		iNotas = vbNo
	End If
	iExportar = MsgBox("¿Exportar cada tema en archivos separados?", vbYesNoCancel, "Opciones exportar")
	If iExportar = vbCancel Then Exit Sub

	Set dcLibro = Iniseg.ConversionLibro(ActiveDocument)
	Debug.Print "A/4 - Archivo libro: salvando"
	dcLibro.Save

	Set dcStory = Iniseg.ConversionStory(dcLibro, iNotas, iExportar)
	Debug.Print "B/4 - Archivo story: salvando"
	dcStory.Save

	Debug.Print "C/4 - Archivo story: cerrando"
	dcStory.Close wdDoNotSaveChanges
	dcLibro.Activate
	dcLibro.Save

	Debug.Print "D/4 - Iniseg2LibroYStory terminada"
	Beep
	MsgBox "Revisar formato libro (viudas/huérfanas, tamaño de imágenes o tablas...), exportar material necesario y ejecutar iniseg 3"
End Sub

Sub Iniseg3PáginasBlancasVisibles()
	' Esta macro es una mala práctica y solo está para evitar confusiones por
		' falta de uniformidad en el uso de plantillas y estilos
	RaMacros.SectionsFillBlankPages ActiveDocument
	Debug.Print "Iniseg3PáginasBlancasVisibles terminada"
End Sub









Function ConversionLibro(dcLibro As Document) As Document
' Realiza la limpieza necesaria y formatea correctamente
'
	Dim iContador As Integer

	Debug.Print "1/17 - Archivo libro: haciendo copia de seguridad (1)"
	RaMacros.SaveAsNewFile dcLibro, "1-", "", True
	Debug.Print "2/17 - Archivo libro: limpieza básica"
	RaMacros.CleanBasic dcLibro

	Debug.Print "3/17 - Archivo libro: títulos sin puntuación"
	RaMacros.HeadingsNoPunctuation dcLibro
	Debug.Print "4.1/17 - Archivo libro: títulos sin numeración repetida"
	RaMacros.HeadingsNoNumeration dcLibro
	Debug.Print "4.2/17 - Archivo libro: listas sin numeración repetida"
	RaMacros.ListsNoNumeration dcLibro

	' Títulos y mayúsculas
	Debug.Print "5/17 - Archivo libro: Títulos sin AllCaps"
	For iContador = -3 To -10 Step -1
		dcLibro.Styles(iContador).Font.AllCaps = False
	Next iContador
	Debug.Print "6/17 - Archivo libro: Título 1 en mayúsculas"
	dcLibro.Styles(wdstyleheading1).Font.AllCaps = True

	Debug.Print "7/17 - Archivo libro: aplicado estilo correcto a hipervínculos"
	RaMacros.HyperlinksFormatting dcLibro, 1, 0
	Debug.Print "8.1/17 - Archivo libro: aplicado estilo correcto a notas al pie"
	If dcLibro.Footnotes.Count > 0 Then
		dcLibro.StoryRanges(2).Style = wdStyleFootnoteText
		With dcLibro.StoryRanges(2).Find
			.ClearFormatting
			.Replacement.ClearFormatting
			.Format = False
			.MatchCase = False
			.MatchWholeWord = False
			.MatchWildcards = False
			.MatchSoundsLike = False
			.MatchAllWordForms = False
			.Text = "^f"
			.Replacement.style = wdStyleFootnoteReference
			.Execute Replace:=wdReplaceAll
		End With
		Debug.Print "8.2/17 - Archivo libro: sangrando notas al pie"
		RaMacros.FootnotesHangingIndentation dcLibro, 0.5, wdStyleFootnoteText
	Else
		Debug.Print "---No hay notas al pie---"
	End If
	Debug.Print "9/17 - Archivo libro: formateando comillas"
	Iniseg.ComillasFormato dcLibro
	Debug.Print "10/17 - Archivo libro: sustituyendo formatos directos por estilos"
	RaMacros.StylesNoDirectFormatting dcLibro

	Debug.Print "11/17 - Archivo libro: formateando imágenes"
	Iniseg.ImagenesLibro dcLibro

	Debug.Print "12/17 - Archivo libro: corrigiendo limpieza e interlineado"
	Iniseg.InterlineadoCorregido dcLibro
	RaMacros.CleanBasic dcLibro

	Debug.Print "13/17 - Archivo libro: añadiendo párrafos de separación"
	Iniseg.ParrafosSeparacionLibro dcLibro
	Debug.Print "14/17 - Archivo libro: añadiendo párrafos de separación antes de tablas"
	Iniseg.TablasParrafosSeparacion dcLibro
	Debug.Print "15/17 - Archivo libro: añadiendo saltos de sección antes de Títulos 1"
	RaMacros.SectionBreakBeforeHeading dcLibro, False, 4, 1
	Debug.Print "16/17 - Archivo libro: añadiendo saltos de página antes de Títulos de bibliografía"
	Iniseg.BibliografiaSaltosDePagina dcLibro

	Do While dcLibro.Paragraphs.Last.Range.Text = vbCr
		If dcLibro.Paragraphs.Last.Range.Delete = 0 Then Exit Do
	Loop

	Debug.Print "17/17 - Conversión a libro terminada"
	Set ConversionLibro = dcLibro
End Function

Function ConversionStory(dcLibro As Document, Optional iExportarNotas As Integer = 0, Optional iExportarSeparados As Integer = 0) As Document
' Da el tamaño correcto a párrafos, imágenes y formatea marcas de pie de página
'
	Dim dcStory As Document, dcBibliografia As Document, iUltima As Integer

	If iExportarNotas = 0 And dcLibro.Footnotes.Count > 0 Then
		iExportarNotas = MsgBox("¿Exportar notas al pie de página a archivo separado?", vbYesNoCancel, "Opciones exportar")
		If iExportarNotas = vbCancel Then Exit Function
	ElseIf iExportarNotas < 6 Or iExportarNotas > 7 Then
		Err.Raise Number:=513, Description:="iExportarNotas out of range"
	End If

	If iExportarSeparados = 0 And dcLibro.Sections.Count > 1 Then
		iExportarSeparados = MsgBox("¿Exportar cada tema en archivos separados?", vbYesNoCancel, "Opciones exportar")
		If iExportarSeparados = vbCancel Then Exit Function
	ElseIf iExportarSeparados < 6 Or iExportarSeparados > 7 Then
		Err.Raise Number:=513, Description:="iExportarSeparados out of range"
	End If

	iUltima = 10
	If iExportarNotas = vbYes Then iUltima = iUltima + 1
	If iExportarNotas = vbYes Then iUltima = iUltima + 1

	Debug.Print "1/" & iUltima & " - Archivo story: creando"
	Set dcStory = RaMacros.SaveAsNewFile(dcLibro, "2-", "", False)

	Debug.Print "2/" & iUltima & " - Archivo story: exportando y borrando bibliografías"
	Iniseg.BibliografiaExportar dcStory

	Debug.Print "3/" & iUltima & " - Archivo story: Títulos 1 sin mayúsculas"
	With dcStory.Content.Find
		.ClearFormatting
		.Replacement.ClearFormatting
		.Forward = True
		.Format = False
		.MatchCase = False
		.MatchWholeWord = False
		.MatchWildcards = True
		.MatchSoundsLike = False
		.MatchAllWordForms = False
		.Style = wdStyleHeading1
		.Text = "([tT][eE][mM][aA] [0-9]{1;2})*^l^l"
		.Replacement.Text = ""
		.Execute Replace:=wdReplaceAll
	End With
	dcStory.Styles(wdstyleheading1).Font.AllCaps = False
	RaMacros.HeadingsChangeCase dcStory, 1, 4

	Debug.Print "4/" & iUltima & " - Archivo story: convirtiendo listas a texto"
	dcStory.ConvertNumbersToText
	Debug.Print "5/" & iUltima & " - Archivo story: adaptando el tamaño de párrafos"
	Iniseg.ParrafosConversionStory dcStory
	Debug.Print "6/" & iUltima & " - Archivo story: títulos con 3 espacios en vez de tabulación"
	Iniseg.TitulosConTresEspacios dcStory
	Debug.Print "7/" & iUltima & " - Archivo story: títulos divididos para no solaparse con el logo en la diapositiva"
	Iniseg.TitulosDivididos dcStory
	Debug.Print "8/" & iUltima & " - Archivo story: formateando imágenes"
	Iniseg.ImagenesStory dcStory
	Debug.Print "9/" & iUltima & " - Archivo story: corrigiendo interlineado"
	Iniseg.InterlineadoCorregido dcStory
	Debug.Print "10/" & iUltima & " - Archivo story: exportando y borrando tablas"
	If dcStory.Tables.Count > 0 Then
		RaMacros.TablesExportToPdf dcStory,, True, wdStyleBlockQuotation, 17
		With dcStory.Content.Find
			.ClearFormatting
			.Replacement.ClearFormatting
			.Forward = True
			.Format = False
			.MatchCase = False
			.MatchWholeWord = False
			.MatchWildcards = False
			.MatchSoundsLike = False
			.MatchAllWordForms = False
			.Text = "Enlace a tabla"
			.Replacement.ParagraphFormat.Alignment = wdAlignParagraphCenter
			.Execute Replace:=wdReplaceAll
		End With
	Else
		Debug.Print "--- No hay tablas ---"
	End If

	If iExportarNotas = vbYes Then
		Debug.Print iUltima - 2 & ".1/" & iUltima & " - exportando notas a archivo externo"
		Iniseg.NotasPieExportar dcLibro
			Debug.Print iUltima - 2 & ".2/" & iUltima & " - Archivo story: formateando notas"
		Iniseg.NotasPieMarcas dcStory, True
	ElseIf dcLibro.Footnotes.Count > 0 Then
		Debug.Print iUltima - 2 & "/" & iUltima & " - Archivo story: formateando notas"
		Iniseg.NotasPieMarcas dcStory, False
	Else
		Debug.Print iUltima - 2 & "/" & iUltima & " - Archivo story: --- no hay notas al pie ---"
	End If

	If iExportarSeparados = vbYes And dcLibro.Sections.Count > 1 Then
		Debug.Print iUltima - 1 & "/" & iUltima & " - Archivo story: exportando en archivos separados"
		RaMacros.SectionsExportEachToFiles dcStory,, "-tema_"
	End If

	Debug.Print iUltima & "/" iUltima & " - Conversión para story terminada"
	Set ConversionStory = dcStory
End Function






Sub ActualizandoVersion(dcArgument As Document)
' Actualiza el formato del archivo a la última versión para solucionar problemas de compatibilidades
'
	Dim iIndex As Integer

	' Conversión de los campos INCLUDEPICTURE a imágenes
	For iIndex = dcArgument.Fields.Count To 1 Step -1
		If dcArgument.Fields(iIndex).Type = wdFieldIncludePicture Then dcArgument.Fields(iIndex).Unlink
	Next iIndex

	' Al convertir el archivo a una versión moderna se les da a las imagenes las propiedades y métodos adecuados para su manipulación
	If dcArgument.CompatibilityMode < 15 Then dcArgument.Convert
End Sub






Sub HeaderCopy(dcOriginalDocument As Document, _
				dcObjectiveDocument As Document, _
				Optional iHeaderOption As Integer = 3)
' Copia los encabezados de un archivo a otro según la opción que se le pase:
	' iHeaderOption = 1 => copia el encabezado de pág. impar en todos los encabezados
	' iHeaderOption = 2 => copia los de pág. impar y par
	' iHeaderOption = 3 => respeta el encabezado diferente de la primera pág.
' TODO
	' GUI para seleccionar qué copiar, cómo y de qué archivo
'
	If iHeaderOption > 3 Or iHeaderOption < 1 Then
		Err.Raise Number:=513, Description:="iHeaderOption out of range"
	End If

	Dim stOriginalHeader As String, iHeader As Integer

	If iHeaderOption = 1 Then
		stOriginalHeader = Replace(dcOriginalDocument.Sections(1).Headers(1).Range.Text, vbLf, "")
		stOriginalHeader = Trim(Replace(stOriginalHeader, vbCr, ""))
	End If

	For iHeader = 1 To 3
		If iHeaderOption > 1 Then
			If iHeaderOption = 2 And iHeader = 2 Then
				stOriginalHeader = Replace(dcOriginalDocument.Sections(1).Headers(1).Range.Text, vbLf, "")
			Else
				stOriginalHeader = Replace(dcOriginalDocument.Sections(1).Headers(iHeader).Range.Text, vbLf, "")
			End If
			stOriginalHeader = Trim(Replace(stOriginalHeader, vbCr, ""))
		End If

		With dcObjectiveDocument.Sections(1).Headers(iHeader).Range.Find
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
			.Text = "Título libro"
			.Replacement.Text = stOriginalHeader
			.Execute Replace:=wdReplaceOne
		End With

		dcObjectiveDocument.Sections(1).Headers(iHeader).Range.Case = wdLowerCase
		dcObjectiveDocument.Sections(1).Headers(iHeader).Range.Case = wdTitleWord
	Next iHeader
End Sub





Sub AutoFormateo(dcArgument As Document)
	' Convierte las URL de texto plano a hiperenlaces
	' Convierte los símbolos en viñeta
	' Da estilo de lista a las listas
	' Hace que los paréntesis tengan principio y cierre
	' Convierte dos guiones seguidos en un guión largo
	'
		' Cambian cosas que no se pueden desactivar:
			' Borra párrafos vacíos
' TODO
	' Recoger y devolver las propiedades con un bucle ForEach, usando ReDim al principio de cada ciclo
'
	Dim optAutoformatValores(14) As Boolean

	With Options

		optAutoformatValores(0) = .AutoFormatApplyBulletedLists
		optAutoformatValores(1) = .AutoFormatApplyFirstIndents
		optAutoformatValores(2) = .AutoFormatApplyHeadings
		optAutoformatValores(3) = .AutoFormatApplyLists
		optAutoformatValores(4) = .AutoFormatApplyOtherParas
		optAutoformatValores(5) = .AutoFormatPlainTextWordMail
		optAutoformatValores(6) = .AutoFormatMatchParentheses
		optAutoformatValores(7) = .AutoFormatPreserveStyles
		optAutoformatValores(8) = .AutoFormatReplaceFarEastDashes
		optAutoformatValores(9) = .AutoFormatReplaceFractions
		optAutoformatValores(10) = .AutoFormatReplaceHyperlinks
		optAutoformatValores(11) = .AutoFormatReplaceOrdinals
		optAutoformatValores(12) = .AutoFormatReplacePlainTextEmphasis
		optAutoformatValores(13) = .AutoFormatReplaceQuotes
		optAutoformatValores(14) = .AutoFormatReplaceSymbols

		.AutoFormatApplyBulletedLists = True
		.AutoFormatApplyFirstIndents = False
		.AutoFormatApplyHeadings = False
		.AutoFormatApplyLists = False
		.AutoFormatApplyOtherParas = False
		.AutoFormatPlainTextWordMail = False
		.AutoFormatMatchParentheses = True
		.AutoFormatPreserveStyles = True
		.AutoFormatReplaceFarEastDashes = False
		.AutoFormatReplaceFractions = False
		.AutoFormatReplaceOrdinals = False
		.AutoFormatReplaceHyperlinks = True
		.AutoFormatReplacePlainTextEmphasis = False
		.AutoFormatReplaceQuotes = False
		.AutoFormatReplaceSymbols = True

		dcArgument.Range.AutoFormat

		.AutoFormatApplyBulletedLists = optAutoformatValores(0)
		.AutoFormatApplyFirstIndents = optAutoformatValores(1)
		.AutoFormatApplyHeadings = optAutoformatValores(2)
		.AutoFormatApplyLists = optAutoformatValores(3)
		.AutoFormatApplyOtherParas = optAutoformatValores(4)
		.AutoFormatPlainTextWordMail = optAutoformatValores(5)
		.AutoFormatMatchParentheses = optAutoformatValores(6)
		.AutoFormatPreserveStyles = optAutoformatValores(7)
		.AutoFormatReplaceFarEastDashes = optAutoformatValores(8)
		.AutoFormatReplaceFractions = optAutoformatValores(9)
		.AutoFormatReplaceHyperlinks = optAutoformatValores(10)
		.AutoFormatReplaceOrdinals = optAutoformatValores(11)
		.AutoFormatReplacePlainTextEmphasis = optAutoformatValores(12)
		.AutoFormatReplaceQuotes = optAutoformatValores(13)
		.AutoFormatReplaceSymbols = optAutoformatValores(14)
	End With

End Sub





Sub ComillasFormato(dcArgument As Document)
' Quita la negrita y cursiva de las comillas y las pasa a curvadas
'
	Dim bSmtQt As Boolean
	bSmtQt = Options.AutoFormatAsYouTypeReplaceQuotes
	Options.AutoFormatAsYouTypeReplaceQuotes = True
	RaMacros.FindAndReplaceClearParameters

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
		.Replacement.Font.Bold = False
		.Replacement.Font.Italic = False
		.Replacement.Font.Underline = wdUnderlineNone
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





Sub ImagenesLibro(dcArgument As Document)
' Formatea más cómodamente las imágenes
	' Las convierte de flotantes a inline (de shapes a inlineshapes)
	' Impide que aparezcan deformadas (mismo % relativo al tamaño original en alto y ancho)
	' Las centra
	' Impide que superen el ancho de página
'
	Dim inlShape As InlineShape, sngRealPageWidth As Single, sngRealPageHeight As Single, iIndex As Integer

	sngRealPageWidth = dcArgument.PageSetup.PageWidth - dcArgument.PageSetup.Gutter _
		- dcArgument.PageSetup.RightMargin - dcArgument.PageSetup.LeftMargin

	sngRealPageHeight = dcArgument.PageSetup.PageHeight _
		- dcArgument.PageSetup.TopMargin - dcArgument.PageSetup.BottomMargin _
		- dcArgument.PageSetup.FooterDistance - dcArgument.PageSetup.HeaderDistance

	' Se convierten los formatos extraños a imágenes 
		' NO FUNCIONA PORQUE CUANDO HAY CAMPOS DE UNA VERSIÓN ANTIGUA SE CORROMPE EL PORTAPAPELES
	' For iIndex = dcArgument.InlineShapes.Count To 1 Step -1
	' 	With dcArgument.Inlineshapes(iIndex)
	' 		If .Type = wdInlineShapeLinkedPicture _ 
	' 			Or .Type = wdInlineShapeEmbeddedOLEObject _
	' 			Or .Type = wdInlineShapeLinkedOLEObject _
	' 		Then
	' 			With .Range
	' 				.CopyAsPicture
	' 				.Delete
	' 				.PasteSpecial DataType:=wdPasteEnhancedMetafile, Placement:=wdInline
	' 			End With
	' 		End If
	' 	End With
	' Next iIndex

	' Se convierten todas de shapes a inlineshapes
	If dcArgument.Shapes.Count > 0 Then
		For iIndex = dcArgument.Shapes.Count To 1 Step -1
			With dcArgument.Shapes(iIndex)
				'If .Type = msoPicture Then
				.LockAnchor = True
				.RelativeVerticalPosition = wdRelativeVerticalPositionParagraph
				With .WrapFormat
					.AllowOverlap = False
					.DistanceTop = 8
					.DistanceBottom = 8
					.Type = wdWrapTopBottom
				End With
				.ConvertToInlineShape
				'End If
			End With
		Next iIndex
	End If

	' Se les da el formato correcto
	For Each inlShape In dcArgument.InlineShapes
		With inlShape
			If .Type = wdInlineShapePicture Then
				.ScaleHeight = .ScaleWidth
				.LockAspectRatio = msoTrue
				.Width = sngRealPageWidth
				' CON ESTO SE LE DA EL ANCHO ORIGINAL DE LA IMAGEN O EL DEL ANCHO DE PÁGINA, SI LO EXCEDE, EN VEZ DE HACER QUE OCUPE TODO EL ANCHO DE PÁGINA
				' If .Width / (.ScaleWidth / 100) > sngRealPageWidth Then .Width = sngRealPageWidth Else .ScaleWidth = 100
				If .Height > .Width And .Height / (.ScaleHeight / 100) > sngRealPageHeight - 15 Then 
					.Height = sngRealPageHeight - 15
				End If

				If .Range.Previous(Unit:=wdCharacter, Count:=1).Text <> vbCr Then
					.Range.InsertBefore vbCr
				End If
				If .Range.Next(Unit:=wdCharacter, Count:=1).Text <> vbCr Then
					.Range.InsertAfter vbCr
				End If

				.Range.Style = wdStyleNormal
				.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
			End If
		End With
	Next inlShape

End Sub





Sub ImagenesStory(dcArgument As Document)
' Hace que todas las imágenes sean enormes, para meterlas en el story
'
	Dim inlShape As InlineShape

	For Each inlShape In dcArgument.InlineShapes
		inlShape.Width = CentimetersToPoints(29)
	Next inlShape
End Sub





Sub InterlineadoCorregido(dcArgument As Document)
' Interlineado de 1,15 sin espaciado entre párrafos
	' Eliminar los espaciados verticales entre párrafos y aplica el interlineado correcto
'
	With dcArgument.Content.ParagraphFormat
		.SpaceBefore = 0
		.SpaceBeforeAuto = False
		.SpaceAfter = 0
		.SpaceAfterAuto = False
		.LineSpacingRule = wdLineSpaceMultiple
		.LineSpacing = LinesToPoints(1.15)
		.LineUnitBefore = 0
		.LineUnitAfter = 0
	End With
	RaMacros.FindAndReplaceClearParameters
End Sub






Sub ParrafosSeparacionLibro(dcArgument As Document)
' Inserta párrafos vacíos de separación
' TODO
	' Refactorizar con variables y recolocando el código
'
	With dcArgument.Range.Find
		.ClearFormatting
		.Replacement.ClearFormatting
		.Forward = True
		.Wrap = wdFindContinue
		.Format = False
		.MatchCase = False
		.MatchWholeWord = False
		.MatchWildcards = True
		.MatchSoundsLike = False
		.MatchAllWordForms = False

		' Elimina saltos manuales de página (innecesarios con los saltos de sección y revisión posteriores)
		.Text = "^m"
		.Replacement.Text = ""
		.Execute Replace:=wdReplaceAll

		' Mete dos saltos de línea manuales en los Heading 1, entre "Tema N" y el nombre del tema
		.Format = True
		.style = wdstyleheading1
		.Text = "([tT][eE][mM][aA] [0-9]{1;2})"
		.Replacement.Text = "\1^l^l"
		.Execute Replace:=wdReplaceAll
	End With

		' Formatea los saltos de línea y les da tamaño 10
	RaMacros.CleanSpaces dcArgument, 0
	With dcArgument.Range.Find
		.ClearFormatting
		.Replacement.ClearFormatting
		.Forward = True
		.Wrap = wdFindContinue
		.Format = True
		.MatchCase = False
		.MatchWholeWord = False
		.MatchWildcards = True
		.MatchSoundsLike = False
		.MatchAllWordForms = False
		.style = wdstyleheading1
		.Replacement.ClearFormatting
		.Replacement.Font.Size = 10
		.Text = "[^13^l]{2;}"
		.Replacement.Text = "^l^l"
		.Execute Replace:=wdReplaceAll
	End With


	' Párrafos de separación generales
	With dcArgument.Range.Find
		.Forward = True
		.Wrap = wdFindContinue
		.Format = True
		.MatchCase = False
		.MatchWholeWord = False
		.MatchAllWordForms = False
		.MatchSoundsLike = False
		.MatchWildcards = True

		' Cambiar saltos de párrafo por saltos de línea en pies de imagen
		.ClearFormatting
		.Replacement.ClearFormatting
		.Text = "(*)^13(*^13)"
		.Style = wdStyleCaption
		.Replacement.Text = "\1^l\2"
		.Execute Replace:=wdReplaceAll

		' Marcas para los títulos
		.Text = "(*^13)"

		.ClearFormatting
		.Replacement.ClearFormatting
		.Style = wdstyleheading1
		.Replacement.Text = "\1SEP_11^13"
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Style = wdstyleheading2
		.Replacement.Text = "SEP_11^13\1SEP_11^13"
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Style = wdstyleheading3
		.Replacement.Text = "SEP_8^13\1SEP_8^13"
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Style = wdstyleheading4
		.Replacement.Text = "SEP_8^13\1SEP_8^13"
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Style = wdStyleHeading5
		.Replacement.Text = "SEP_6^13\1SEP_6^13"
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Style = wdStyleHeading6
		.Replacement.Text = "SEP_6^13\1SEP_6^13"
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Style = wdStyleHeading7
		.Replacement.Text = "SEP_6^13\1SEP_6^13"
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Style = wdStyleHeading8
		.Replacement.Text = "SEP_6^13\1SEP_6^13"
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Style = wdStyleHeading9
		.Replacement.Text = "SEP_6^13\1SEP_6^13"
		.Execute Replace:=wdReplaceAll



		' Marcas del resto de estilos
		.ClearFormatting
		.Replacement.ClearFormatting
		.Style = wdStyleNormal
		.Replacement.Text = "SEP_5^13\1SEP_5^13"
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Style = wdStyleCaption
		.Replacement.Text = "SEP_5^13\1SEP_5^13"
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Style = wdStyleQuote
		.Replacement.Text = "SEP_5^13\1SEP_5^13"
		.Execute Replace:=wdReplaceAll

			' Word tiene un bug y se lía con el último párrafo, si es de lista
		If dcArgument.Paragraphs.Last.Range.ListFormat.ListType <> wdListNoNumbering Then
			dcArgument.Paragraphs.Last.Range.InsertParagraphAfter
			dcArgument.Paragraphs.Last.Style = wdStyleNormal
		End If

		.ClearFormatting
		.Replacement.ClearFormatting
		.Style = wdStyleList
		.Replacement.Text = "SEP_4^13\1SEP_4^13"
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Style = wdStyleList2
		.Replacement.Text = "SEP_4^13\1SEP_4^13"
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Style = wdStyleList3
		.Replacement.Text = "SEP_4^13\1SEP_4^13"
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Style = wdStyleListBullet
		.Replacement.Text = "SEP_4^13\1SEP_4^13"
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Style = wdStyleListBullet2
		.Replacement.Text = "SEP_4^13\1SEP_4^13"
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Style = wdStyleListBullet3
		.Replacement.Text = "SEP_4^13\1SEP_4^13"
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Style = wdStyleListBullet4
		.Replacement.Text = "SEP_4^13\1SEP_4^13"
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Style = wdStyleListBullet5
		.Replacement.Text = "SEP_4^13\1SEP_4^13"
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Style = wdStyleListContinue
		.Replacement.Text = "SEP_4^13\1SEP_4^13"
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Style = wdStyleListContinue2
		.Replacement.Text = "SEP_4^13\1SEP_4^13"
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Style = wdStyleListContinue3
		.Replacement.Text = "SEP_4^13\1SEP_4^13"
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Style = wdStyleListContinue4
		.Replacement.Text = "SEP_4^13\1SEP_4^13"
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Style = wdStyleListContinue5
		.Replacement.Text = "SEP_4^13\1SEP_4^13"
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Style = wdStyleListNumber
		.Replacement.Text = "SEP_4^13\1SEP_4^13"
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Style = wdStyleListNumber2
		.Replacement.Text = "SEP_4^13\1SEP_4^13"
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Style = wdStyleListNumber3
		.Replacement.Text = "SEP_4^13\1SEP_4^13"
		.Execute Replace:=wdReplaceAll



		' Convertir marcas a estilo Normal
		.ClearFormatting
		.Replacement.ClearFormatting
		.Text = "(SEP_[0-9]{1;2}^13)"
		.Replacement.Style = wdStyleNormal
		.Replacement.Text = "\1"
		.Execute Replace:=wdReplaceAll



		' Seleccionar la marca de mayor tamaño, cuando coinciden 2
		.ClearFormatting
		.Replacement.ClearFormatting
		.Format = False

		' Prevalece la segunda
		.Text = "(SEP_4^13)(SEP_4^13)"
		.Replacement.Text = "\2"
		.Execute Replace:=wdReplaceAll

		.Text = "(SEP_4^13)(SEP_5^13)"
		.Replacement.Text = "\2"
		.Execute Replace:=wdReplaceAll

		.Text = "(SEP_4^13)(SEP_6^13)"
		.Replacement.Text = "\2"
		.Execute Replace:=wdReplaceAll

		.Text = "(SEP_4^13)(SEP_8^13)"
		.Replacement.Text = "\2"
		.Execute Replace:=wdReplaceAll

		.Text = "(SEP_4^13)(SEP_11^13)"
		.Replacement.Text = "\2"
		.Execute Replace:=wdReplaceAll

		.Text = "(SEP_5^13)(SEP_5^13)"
		.Replacement.Text = "\2"
		.Execute Replace:=wdReplaceAll

		.Text = "(SEP_5^13)(SEP_6^13)"
		.Replacement.Text = "\2"
		.Execute Replace:=wdReplaceAll

		.Text = "(SEP_5^13)(SEP_8^13)"
		.Replacement.Text = "\2"
		.Execute Replace:=wdReplaceAll

		.Text = "(SEP_5^13)(SEP_11^13)"
		.Replacement.Text = "\2"
		.Execute Replace:=wdReplaceAll

		.Text = "(SEP_6^13)(SEP_6^13)"
		.Replacement.Text = "\2"
		.Execute Replace:=wdReplaceAll

		.Text = "(SEP_6^13)(SEP_8^13)"
		.Replacement.Text = "\2"
		.Execute Replace:=wdReplaceAll

		.Text = "(SEP_6^13)(SEP_11^13)"
		.Replacement.Text = "\2"
		.Execute Replace:=wdReplaceAll

		.Text = "(SEP_8^13)(SEP_8^13)"
		.Replacement.Text = "\2"
		.Execute Replace:=wdReplaceAll

		.Text = "(SEP_8^13)(SEP_11^13)"
		.Replacement.Text = "\2"
		.Execute Replace:=wdReplaceAll

		.Text = "(SEP_11^13)(SEP_11^13)"
		.Replacement.Text = "\2"
		.Execute Replace:=wdReplaceAll


		' Prevalece la primera
		.Text = "(SEP_11^13)(SEP_8^13)"
		.Replacement.Text = "\1"
		.Execute Replace:=wdReplaceAll

		.Text = "(SEP_11^13)(SEP_6^13)"
		.Replacement.Text = "\1"
		.Execute Replace:=wdReplaceAll

		.Text = "(SEP_11^13)(SEP_5^13)"
		.Replacement.Text = "\1"
		.Execute Replace:=wdReplaceAll

		.Text = "(SEP_11^13)(SEP_4^13)"
		.Replacement.Text = "\1"
		.Execute Replace:=wdReplaceAll

		.Text = "(SEP_8^13)(SEP_6^13)"
		.Replacement.Text = "\1"
		.Execute Replace:=wdReplaceAll

		.Text = "(SEP_8^13)(SEP_5^13)"
		.Replacement.Text = "\1"
		.Execute Replace:=wdReplaceAll

		.Text = "(SEP_8^13)(SEP_4^13)"
		.Replacement.Text = "\1"
		.Execute Replace:=wdReplaceAll

		.Text = "(SEP_6^13)(SEP_5^13)"
		.Replacement.Text = "\1"
		.Execute Replace:=wdReplaceAll

		.Text = "(SEP_6^13)(SEP_4^13)"
		.Replacement.Text = "\1"
		.Execute Replace:=wdReplaceAll

		.Text = "(SEP_5^13)(SEP_4^13)"
		.Replacement.Text = "\1"
		.Execute Replace:=wdReplaceAll



		' Redimensionado de párrafo y borrado de marca
		.ClearFormatting
		.Replacement.ClearFormatting
		.Format = True
		.Replacement.Font.Size = 4
		.Text = "SEP_4(^13)"
		.Replacement.Text = "\1"
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Replacement.Font.Size = 5
		.Text = "SEP_5(^13)"
		.Replacement.Text = "\1"
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Replacement.Font.Size = 6
		.Text = "SEP_6(^13)"
		.Replacement.Text = "\1"
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Replacement.Font.Size = 8
		.Text = "SEP_8(^13)"
		.Replacement.Text = "\1"
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Replacement.Font.Size = 11
		.Text = "SEP_11(^13)"
		.Replacement.Text = "\1"
		.Execute Replace:=wdReplaceAll
	End With
	RaMacros.FindAndReplaceClearParameters
End Sub

Sub TablasParrafosSeparacion(dcArgument As Document)
' Inserta un párrafo vacío y marcado antes de cada tabla
'
	Dim iCounter As Integer
	Dim rgTable As Range
	Dim tbCurrent As Table

	For iCounter = 1 To dcArgument.Tables.Count Step 1
		Set tbCurrent = dcArgument.Tables(iCounter)
		If tbCurrent.NestingLevel = 1 Then
			tbCurrent.Title = "Tabla " & iCounter
			tbCurrent.Rows.WrapAroundText = False
			If tbCurrent.Range.Start <> 0 _
				And tbCurrent.Range.Previous(wdParagraph, 1).Text <> vbCr _
			Then
				Set rgTable = tbCurrent.Range.Previous(wdParagraph, 1)
				rgTable.Characters.Last.InsertParagraphBefore
				rgTable.Paragraphs.Last.Style = wdStyleNormal
				rgTable.Paragraphs.Last.Range.Font.Size = 5
			End If
			If tbCurrent.Range.End <> dcArgument.StoryRanges(tbCurrent.Range.StoryType).End _
				And tbCurrent.Range.Next(wdParagraph, 1).Text <> vbCr _
			Then
				Set rgTable = tbCurrent.Range.Next(wdParagraph, 1)
				rgTable.InsertParagraphBefore
				rgTable.Paragraphs.First.Style = wdStyleNormal
				rgTable.Paragraphs.First.Range.Font.Size = 5
			End If
		End If
	Next iCounter
End Sub

Sub ParrafosConversionStory(dcArgument As Document)
' Conversion de Word impreso a formato para Storyline
'
	RaMacros.FindAndReplaceClearParameters

	' Cambio del tamaño de Titulo 2 de 16 a 17
	With dcArgument.Styles(wdstyleheading2).Font
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

	' Eliminar ALLCAPS de los títulos 3 y 4 (por si derivan del título 1 o 2)
	dcArgument.Styles(wdstyleheading3).Font.AllCaps = False
	dcArgument.Styles(wdstyleheading4).Font.AllCaps = False

	' Poner el estilo quote centrado y sin espacio a derecha ni izquierda
	With dcArgument.Styles(wdStyleQuote).ParagraphFormat
		.LeftIndent = 0
		.RightIndent = 0
		.Alignment = wdAlignParagraphCenter
	End With



' Cambio de tamaño de parrafos de separacion
	' Listas: 4 a 2
	With dcArgument.Range.Find
		.ClearFormatting
		.Font.Size = 4
		.Replacement.ClearFormatting
		.Replacement.Font.Size = 2
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
		.Execute Replace:=wdReplaceAll
	End With

	' Parrafos normales: 5 a 4.
	With dcArgument.Range.Find
		.ClearFormatting
		.Font.Size = 5
		.Replacement.ClearFormatting
		.Replacement.Font.Size = 4
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
		.Execute Replace:=wdReplaceAll
	End With

	' Titulos 3, 4 y 5: 8 a 6.
	With dcArgument.Range.Find
		.ClearFormatting
		.Font.Size = 8
		.Replacement.ClearFormatting
		.Replacement.Font.Size = 6
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
		.Execute Replace:=wdReplaceAll
	End With



' Titulos 2: 11 a 8
	' Dar tamaño 8 a todos los párrafos tras los Heading 2
	With dcArgument.Range.Find
		.ClearFormatting
		.Style = wdstyleheading2
		.Replacement.ClearFormatting
		.Forward = True
		.Wrap = wdFindContinue
		.Format = True
		.MatchCase = False
		.MatchWholeWord = False
		.MatchAllWordForms = False
		.MatchSoundsLike = False
		.MatchWildcards = True
		.Text = "(*@^13)"
		.Replacement.Text = "\1FISTRO"
		.Execute Replace:=wdReplaceAll

		.Style = wdStyleNormal
		.Text = "FISTRO^13"
		.Replacement.Font.Size = 8
		.Replacement.Text = "^13"
		.Execute Replace:=wdReplaceAll
	End With

	RaMacros.FindAndReplaceClearParameters
End Sub





Sub TitulosConTresEspacios(dcArgument As Document)
' Sustituye la tabulación en los títulos por 3 espacios
'
	With dcArgument.Range.Find
		.ClearFormatting
		.Replacement.ClearFormatting
		.Forward = True
		.Wrap = wdFindContinue
		.Format = True
		.MatchCase = False
		.MatchWholeWord = False
		.MatchAllWordForms = False
		.MatchSoundsLike = False
		.MatchWildcards = True
		.Text = "([0-9].)^t"
		.Replacement.Text = "\1   "
		.Style = wdstyleheading2
		.Execute Replace:=wdReplaceAll
		.Style = wdstyleheading3
		.Execute Replace:=wdReplaceAll
		.Style = wdstyleheading4
		.Execute Replace:=wdReplaceAll

	End With
	RaMacros.FindAndReplaceClearParameters
End Sub





Sub TitulosDivididos(dcArgument As Document)
' Corta los títulos 2 para que no peguen contra el logo de Iniseg
'	- Título 2: 35 caractéres hasta logo Iniseg
'	- Título 3: 55 caractéres hasta logo Iniseg
'	- Título 4: 65 caractéres hasta logo Iniseg
'	- Título 5: 70 caractéres hasta logo Iniseg
'
	With dcArgument.Range.Find
		.ClearFormatting
		.Replacement.ClearFormatting
		.Forward = True
		.Wrap = wdFindContinue
		.Format = True
		.MatchCase = False
		.MatchWholeWord = False
		.MatchAllWordForms = False
		.MatchSoundsLike = False
		.MatchWildcards = True
		.Text = "(?{30;}) "
		.Replacement.Text = "\1^13      "
		.Style = wdstyleheading2
		.Execute Replace:=wdReplaceAll

		' En principio solo es conveniente hacerlo con los títulos 2, porque los demás no tienen por
			' qué coincidir en la parte de arriba de la diapositiva, pero con el siguiente código
			' se transformarían también los títulos 3 a 5

		' .ClearFormatting
		' .Replacement.ClearFormatting
		' .Text = "(?{50;}) "
		' .Replacement.Text = "\1^13         "
		' .Style = wdstyleheading3
		' .Execute Replace:=wdReplaceAll

		' .ClearFormatting
		' .Replacement.ClearFormatting
		' .Text = "(?{60;}) "
		' .Replacement.Text = "\1^13            "
		' .Style = wdstyleheading4
		' .Execute Replace:=wdReplaceAll

		' .ClearFormatting
		' .Replacement.ClearFormatting
		' .Text = "(?{65;}) "
		' .Replacement.Text = "\1^13"
		' .Style = wdstyleheading5
		' .Execute Replace:=wdReplaceAll

	End With
	RaMacros.FindAndReplaceClearParameters
End Sub





Sub NotasPieMarcas(dcArgument As Document, bExportar As Boolean)
' Convierte las referencias de notas al pie al texto "NOTA_PIE-numNota"
	' para poder automatizar externamente su conversión en el .story
'
	Dim lContadorNotas As Long, lReferencia As Long, rgFootNote As Range, oEstiloNota As Font

	Set oEstiloNota = New Font
	With oEstiloNota
		.Name = "Swis721 Lt BT"
		.Bold = True
		.Color = -738148353
		.Superscript = True
	End With

	For lContadorNotas = ActiveDocument.Footnotes.Count To 1 Step -1
		lReferencia = ActiveDocument.Footnotes.StartingNumber + ActiveDocument.Footnotes(lContadorNotas).Index - 1
		Set rgFootNote = ActiveDocument.Footnotes(lContadorNotas).Reference
		If bExportar Then
			rgFootNote.Text = "NOTA_PIE-" & lReferencia
		Else
			rgFootNote.Previous(wdCharacter, 1).InsertAfter "NOTA_PIE-" & lReferencia
		End If
		rgFootNote.Font = oEstiloNota
	Next lContadorNotas
End Sub






Sub NotasPieExportar(dcArgument As Document)
' Exporta las notas a un archivo separado
' ToDo:
	' Convertir esta subrutina en una función de uso general:
		' Retornar el archivo de notas
'
	Dim dcNotas As Document, stFilename As String, stOriginalName As String, stOriginalExtension As String
	Dim lNotasNuevasInicio As Long, rgNotasNuevas As Range

	With dcArgument
		stOriginalName = Left(.Name, InStrRev(.Name, ".") - 1)
		stOriginalExtension = Right(.Name, Len(.Name) - InStrRev(.Name, ".") + 1)
		stFileName = .Path & Application.PathSeparator & "NOTAS " & stOriginalName & stOriginalExtension
	End With

	If Dir(stFileName) > "" Then
		Set dcNotas = Documents.Open(FileName:=stFileName, ConfirmConversions:=False, ReadOnly:=False, Revert:=False)
		RaMacros.CopySecurity dcNotas, "0-", ""
	Else
		Set dcNotas = RaMacros.SaveAsNewFile(dcArgument, "NOTAS ", "", False)
		dcNotas.Content.Text = "Notas al pie"
		With activedocument.Content.Paragraphs(1)
			.Style = wdStyleTitle
			.Alignment = wdAlignParagraphCenter
		End With
	End If

	dcNotas.Content.InsertParagraphAfter
	Set rgNotasNuevas = dcNotas.Content.Paragraphs.Last.Range
	dcArgument.StoryRanges(wdFootnotesStory).Copy
	dcNotas.Content.Paragraphs.Last.Range.Paste
	rgNotasNuevas.EndOf wdStory, wdExtend

	RaMacros.CleanBasic dcNotas
	rgNotasNuevas.Style = wdStyleListContinue

	With rgNotasNuevas.Find
		.ClearFormatting
		.Replacement.ClearFormatting
		.Forward = True
		.Wrap = wdFindStop
		.Format = True
		.MatchCase = False
		.MatchWholeWord = False
		.MatchWildcards = False
		.MatchSoundsLike = False
		.MatchAllWordForms = False
		.Font.Superscript = True
		.Text = ""
		.Replacement.Style = wdStyleList
		.Replacement.Text = "marca_notas_pie"
		.Execute Replace:=wdReplaceAll

		.Format = False
		.Replacement.ClearFormatting
		.Text = "marca_notas_pie"
		.Replacement.Text = ""
		.Execute Replace:=wdReplaceAll
	End With

	With dcNotas.Styles(wdStyleListContinue)
		.ParagraphFormat.SpaceAfter = 2
		.ParagraphFormat.SpaceBefore = 2
		.ParagraphFormat.Alignment = wdAlignParagraphLeft
		.NoSpaceBetweenParagraphsOfSameStyle = True
	End With
	With dcNotas.Styles(wdStyleList)
		.ParagraphFormat.SpaceAfter = 0
		.ParagraphFormat.SpaceBefore = 5
		.ParagraphFormat.Alignment = wdAlignParagraphLeft
		.NoSpaceBetweenParagraphsOfSameStyle = False
	End With

	RaMacros.CleanBasic dcNotas
	Iniseg.AutoFormateo dcNotas
	RaMacros.HyperlinksFormatting dcNotas, 3, 0
	Do While dcNotas.Paragraphs.Last.Range.Text = vbCr
		dcNotas.Paragraphs.Last.Range.Delete
	Loop

	dcNotas.Save
	stFileName = dcNotas.Path & Application.PathSeparator & "NOTAS " & stOriginalName & ".pdf"
	dcNotas.ExportAsFixedFormat OutputFileName:=stFileName, ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, _
		OptimizeFor:=wdExportOptimizeForPrint, Range:=wdExportAllDocument, Item:=wdExportDocumentWithMarkup, _
		CreateBookmarks:=wdExportCreateHeadingBookmarks
	dcNotas.Close wdSaveChanges
End Sub






Sub BibliografiaSaltosDePagina(dcArgument As Document)
' Inserta un salto de página antes de cada bibliografía
'
	Dim scCurrentSection As Section, rgFindRange As Range

	For Each scCurrentSection In dcArgument.Sections
		Set rgFindRange = scCurrentSection.Range
		With rgFindRange.Find
			.ClearFormatting
			.Replacement.ClearFormatting
			.Forward = True
			.Wrap = wdFindStop
			.Format = True
			.MatchCase = False
			.MatchWholeWord = False
			.MatchWildcards = False
			.MatchSoundsLike = False
			.MatchAllWordForms = False
			.Style = wdStyleHeading2
			.Execute FindText:="bibliografía"
			If Not .Found Then
				.Execute FindText:="bibliografia"
				If Not .Found Then
					.Execute FindText:="referencias"
				End If
			End If
		End With
			If rgFindRange.Find.Found Then
				Set rgFindRange = rgFindRange.Previous(wdParagraph, 1)
				If rgFindRange.Characters(1).Text <> Chr(12) Then
					If rgFindRange.Text = vbCr Then
						rgFindRange.InsertBreak 7
					Else
						rgFindRange.InsertParagraphAfter
						Set rgFindRange = rgFindRange.Paragraphs.Last.Range
						rgFindRange.Select
						rgFindRange.style = wdStyleNormal
						rgFindRange.InsertBreak 7
					End If
				End If
			End If
	Next scCurrentSection
End Sub

Sub BibliografiaExportar(dcArgument As Document)
' Exporta la bibliografía en archivos separados y la borra de dcArgument
'
	Dim dcBibliografia As Document, scCurrentSection As Section
	Dim rgFindRange As Range, rgTitulo As Range, stNombre As String

	For Each scCurrentSection In dcArgument.Sections
		Set rgFindRange = scCurrentSection.Range
		With rgFindRange.Find
			.ClearFormatting
			.Replacement.ClearFormatting
			.Forward = True
			.Wrap = wdFindStop
			.Format = True
			.MatchCase = False
			.MatchWholeWord = False
			.MatchWildcards = False
			.MatchSoundsLike = False
			.MatchAllWordForms = False
			.Style = wdStyleHeading2
			.Execute FindText:="bibliografía"
			If Not .Found Then
				.Execute FindText:="bibliografia"
				If Not .Found Then
					.Execute FindText:="referencias"
				End If
			End If
		End With
		If rgFindRange.Find.Found Then
			' Set dcBibliografia = Documents.Add("C:\Users\Ra\Documents\Plantillas personalizadas de Office\iniseg.dotm")
			' Iniseg.HeaderCopy dcArgument, dcBibliografia, 1
			' rgFindRange.End = scCurrentSection.Range.End
			' dcBibliografia.Content.FormattedText = rgFindRange
			' rgFindRange.End = scCurrentSection.Range.End - 1
			' rgFindRange.Delete

			' Set rgFindRange = scCurrentSection.Range
			' With rgFindRange.Find
			' 	.MatchWildcards = True
			' 	.Execute FindText:="TEMA [0-9][0-9]"
			' 	If Not .Found Then .Execute FindText:="TEMA [0-9]"
			' 	If .Found Then
			' 		stNombre = dcArgument.Path & Application.PathSeparator _
			' 			& "BIBLIOGRAFÍA " & rgFindRange.Text
			' 	Else
			' 		Beep
			' 		stNombre = InputBox("Número de tema no encontrado, completar", "Bibliografía", "TEMA " & scCurrentSection.Index)
			' 		stNombre = dcArgument.Path & Application.PathSeparator _
			' 			& "BIBLIOGRAFÍA " & stNombre
			' 	End If
			' End With

			' dcBibliografia.Close wdSaveChanges

			' Asigna el número de tema
			Set rgTitulo = scCurrentSection.Range
			With rgTitulo.Find
				.MatchWildcards = True
				.Execute FindText:="TEMA [0-9][0-9]"
				If Not .Found Then .Execute FindText:="TEMA [0-9]"
				If .Found Then
					stNombre = dcArgument.Path & Application.PathSeparator _
						& "BIBLIOGRAFÍA " & rgTitulo.Text & ".pdf"
				Else
					Beep
					stNombre = InputBox("Número de tema no encontrado, completar", "Bibliografía", "TEMA " & scCurrentSection.Index)
					stNombre = dcArgument.Path & Application.PathSeparator _
						& "BIBLIOGRAFÍA " & stNombre & ".pdf"
				End If
			End With

			' Exporta el pdf
			rgFindRange.End = scCurrentSection.Range.End - 1
			rgFindRange.ExportAsFixedFormat2 _
				stNombre,wdExportFormatPDF,False,wdExportOptimizeForPrint,True, _
				wdExportDocumentWithMarkup,True,,wdExportCreateNoBookmarks,True,False,False,True

			' Borra la bibliografía de dcStory
			rgFindRange.Delete
		End If
	Next scCurrentSection
End Sub






Sub ConversionAutomaticaLibro(dcArgument As Document)
' Convierte automáticamente los párrafos a los estilos de la plantilla
'
	RaMacros.CleanBasic dcArgument
	With dcArgument.Content.Find
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

		.Text = ""
		.Replacement.Style = wdStyleHeading2
		.Font.Size = 25
		.Execute Replace:=wdReplaceAll
		.Font.Size = 24
		.Execute Replace:=wdReplaceAll
		.Font.Size = 23
		.Execute Replace:=wdReplaceAll
		.Font.Size = 22
		.Execute Replace:=wdReplaceAll
		.Font.Size = 21
		.Execute Replace:=wdReplaceAll
		.Font.Size = 20
		.Execute Replace:=wdReplaceAll
		.Font.Size = 19
		.Execute Replace:=wdReplaceAll
		.Font.Size = 18
		.Execute Replace:=wdReplaceAll
		.Font.Size = 17
		.Execute Replace:=wdReplaceAll
		.Font.Size = 16
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Replacement.Style = wdStyleHeading1
		.Style = wdStyleHeading2
		.MatchWildcards = True
		.Text = "([tT][eE][mM][aA]*[0-9]{1;2}*^13*^13)"
		.Replacement.Text = "\1"
		.Execute Replace:=wdReplaceAll

		.MatchWildcards = False
		.Text = ""
		.Replacement.Text = ""

		.ClearFormatting
		.Replacement.ClearFormatting
		.Replacement.Style = wdStyleHeading4
		.Font.Italic = True
		.Font.Size = 15
		.Execute Replace:=wdReplaceAll
		.Font.Size = 14
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Replacement.Style = wdStyleHeading3
		.Font.Italic = False
		.Font.Size = 15
		.Execute Replace:=wdReplaceAll
		.Font.Size = 14
		.Execute Replace:=wdReplaceAll

		.ClearFormatting
		.Replacement.ClearFormatting
		.Replacement.Style = wdStyleHeading5
		.Font.Size = 13
		.Execute Replace:=wdReplaceAll
		.Font.Size = 12
		.Execute Replace:=wdReplaceAll
End Sub






Sub ColoresCorrectos(dcArgument As Document)
' Quita el subrayado y los colores fuera de plantilla del texto
'
	Dim iMaxCount As Integer
	Dim rgStory As Range
	Dim iStory As Integer

	For iStory = 1 To 5 Step 1
		On Error Resume Next
		Set rgStory = dcArgument.StoryRanges(iStory)
		If Err.Number = 0 Then
			On Error GoTo 0
			rgStory.Font.ColorIndex = wdAuto
			rgStory.Font.Shading.Texture  = wdTextureNone
			rgStory.Font.Shading.BackgroundPatternColor = wdColorAutomatic
			rgStory.Font.Shading.ForegroundPatternColor  = wdColorAutomatic
		Else
			On Error GoTo 0
		End If
	Next iStory
End Sub