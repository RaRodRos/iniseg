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
	Debug.Print "1A/12 - Haciendo copia de seguridad (0) del archivo original"
	RaMacros.CopySecurity dcOriginalFile, "0-", ""

	If iDeleteAnswer = vbYes Then
		If rgRangoActual.Footnotes.Count > 0 Then
			lPrimeraNotaAlPie = rgRangoActual.Footnotes(1).Index
		End If
		Debug.Print "1B/12 - Borrando el texto seleccionado"
		rgRangoActual.End = rgRangoActual.Start
		rgRangoActual.Start = 0
		rgRangoActual.Delete
	End If

	' Actualización del formato del archivo (soluciona problemas de compatibilidad con shapes y campos)
	Debug.Print "1C/12 - Actualizando formato de archivo"
	Iniseg.ActualizandoVersion dcOriginalFile

	Debug.Print "2/12 - Creando archivo con plantilla Iniseg"
	Set dcLibro = Documents.Add("C:\Users\Ra\Documents\Plantillas personalizadas de Office\iniseg.dotm")

	Debug.Print "3/12 - Copiando encabezados"
	Iniseg.HeaderCopy dcOriginalFile, dcLibro, 1

	Debug.Print "4/12 - Aplicando autoformateo"
	Iniseg.AutoFormateo dcOriginalFile
	Debug.Print "5/12 - Limpiando hiperenlaces para que solo figure su dominio"
	RaMacros.HyperlinksOnlyDomain dcOriginalFile
	RaMacros.CleanSpaces dcOriginalFile

	Debug.Print "6/12 - Borrando encabezados y pies de página"
	RaMacros.HeadersFootersRemove dcOriginalFile

	Debug.Print "7/12 - Borrando estilos sin uso"
	lEstilosBorrados = RaMacros.StylesDeleteUnused(dcOriginalFile, False)

	' Copia de seguridad limpia
	Debug.Print "8/12 - Creando copia de seguridad limpia (01)"
	RaMacros.SaveAsNewFile dcOriginalFile, "01-", "", True

	' Guarda el archivo con nombre original, preparado para el siguiente paso
	Debug.Print "9A/12 - Copiando contenido limpio al archivo con plantilla (archivo libro)"
	dcLibro.Content.FormattedText = dcOriginalFile.Content

	If lPrimeraNotaAlPie <> 0 Then
		Debug.Print "9B/12 - Archivo libro: corrigiendo el número de comienzo de las notas al pie"
		dcLibro.Footnotes.StartingNumber = lPrimeraNotaAlPie
	End If

	Debug.Print "10/12 - Archivo original: cerrando"
	dcOriginalFile.Close wdDoNotSaveChanges
	Debug.Print "11/12 - Archivo libro: guardando"
	dcLibro.SaveAs2 stFileName
	dcLibro.Activate

	Debug.Print "12/12 - Iniseg1Limpieza terminada"
	Beep
	MsgBox lEstilosBorrados & " Estilos borrados" & vbCrLf & "Revisar numeración de notas al pie, aplicar estilos y ejecutar Iniseg2"
End Sub





Sub Iniseg2LibroYStory()
' Llama a las macros de ConversionLibro e ConversionStory y da un aviso para seguir trabajando
	' Organizado de esta forma las macros de libro y story se pueden llamar por separado
'
	Dim dcLibro As Document, dcStory As Document, iExportarSeparados As Integer, iNotas As Integer

	iExportarSeparados = MsgBox("¿Exportar cada tema en archivos separados?", vbYesNoCancel, "Opciones exportar")
	If iExportarSeparados = vbCancel Then Exit Sub
	iNotas = MsgBox("¿Exportar notas al pie de página a un archivo?", vbYesNoCancel, "Opciones exportar")
	If iNotas = vbCancel Then Exit Sub

	Set dcLibro = Iniseg.ConversionLibro(ActiveDocument)
	Debug.Print "A/4 - Archivo libro: salvando"
	dcLibro.Save

	Set dcStory = Iniseg.ConversionStory(dcLibro, iNotas)
	Debug.Print "B/4 - Archivo story: salvando"
	dcStory.Save

	If iExportarSeparados = vbYes Then
		Debug.Print "B2/4 - Archivo story: exportando en archivos separados"
		RaMacros.SectionsExportEachToFiles dcStory,, "-tema_"
	End If

	Debug.Print "C/4 - Archivo story: cerrando"
	dcStory.Close wdDoNotSaveChanges
	dcLibro.Activate

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
	Debug.Print "1/15 - Archivo libro: haciendo copia de seguridad (1)"
	RaMacros.SaveAsNewFile dcLibro, "1-", "", True
	Debug.Print "2/15 - Archivo libro: limpieza básica"
	RaMacros.CleanBasic dcLibro

	Debug.Print "3/15 - Archivo libro: títulos sin puntuación"
	RaMacros.HeadingsNoPunctuation dcLibro
	Debug.Print "4/15 - Archivo libro: títulos sin numeración repetida"
	RaMacros.HeadingsNoNumeration dcLibro
	' Títulos con mayúsculas tipo título
	dcLibro.Styles(wdstyleheading1).Font.AllCaps = False
	Debug.Print "5/15 - Archivo libro: Títulos en sentence case"
	RaMacros.HeadingsChangeCase dcLibro, 0, 4
	Debug.Print "6/15 - Archivo libro: Título 1 en mayúsculas"
	dcLibro.Styles(wdstyleheading1).Font.AllCaps = True

	Debug.Print "7/15 - Archivo libro: aplicado estilo correcto a hipervínculos"
	RaMacros.HyperlinksFormatting dcLibro
	Debug.Print "8/15 - Archivo libro: formateando comillas"
	Iniseg.ComillasFormato dcLibro
	Debug.Print "9/15 - Archivo libro: sustituyendo formatos directos por estilos"
	RaMacros.StylesNoDirectFormatting dcLibro

	Debug.Print "10/15 - Archivo libro: formateando imágenes"
	Iniseg.ImagenesLibro dcLibro

	Iniseg.InterlineadoCorregido dcLibro
	Debug.Print "11/15 - Archivo libro: corrigiendo limpieza e interlineado"
	RaMacros.CleanBasic dcLibro

	Debug.Print "12/15 - Archivo libro: añadiendo párrafos de separación"
	Iniseg.ParrafosSeparacionLibro dcLibro
	Debug.Print "13/15 - Archivo libro: añadiendo saltos de sección antes de Títulos 1"
	RaMacros.SectionBreakBeforeHeading dcLibro, False, 4, 1
	Debug.Print "14/15 - Archivo libro: añadiendo saltos de página antes de Títulos de bibliografía"
	Iniseg.BibliografiaSaltosDePagina dcLibro

	Do While dcLibro.Paragraphs.Last.Range.Text = vbCr
		dcLibro.Paragraphs.Last.Range.Delete
	Loop

	Debug.Print "15/15 - Conversión a libro terminada"
	Set ConversionLibro = dcLibro
End Function










Function ConversionStory(dcLibro As Document, Optional iExportarNotas As Integer = 0) As Document
' Da el tamaño correcto a párrafos, imágenes y formatea marcas de pie de página
'
	Dim dcStory As Document, dcBibliografia As Document

	If iExportarNotas = 0 Then
		Beep
		iExportarNotas = MsgBox("¿Exportar notas al pie de página a un archivo?", vbYesNo, "Notas al pie")
	ElseIf iExportarNotas < 6 Or iExportarNotas > 7 Then
		Err.Raise Number:=513, Description:="iExportarNotas out of range"
	End If

	Debug.Print "1/11 - Archivo story: creando"
	Set dcStory = RaMacros.SaveAsNewFile(dcLibro, "2-", "", False)
	Debug.Print "2/11 - Archivo story: exportando y borrando bibliografías"
	Iniseg.BibliografiaExportar dcStory
	Debug.Print "3/11 - Archivo story: Títulos 1 sin mayúsculas"
	dcLibro.Styles(wdstyleheading1).Font.AllCaps = False
	Iniseg.InterlineadoCorregido dcStory
	Debug.Print "4/11 - Archivo story: convirtiendo listas a texto"
	RaMacros.ListsToText dcStory
	Debug.Print "5/11 - Archivo story: adaptando el tamaño de párrafos"
	Iniseg.ParrafosConversionStory dcStory
	Debug.Print "6/11 - Archivo story: títulos con 3 espacios en vez de tabulación"
	Iniseg.TitulosConTresEspacios dcStory
	Debug.Print "7/11 - Archivo story: títulos divididos para no solaparse con el logo en la diapositiva"
	Iniseg.TitulosDivididos dcStory
	Debug.Print "8/11 - Archivo story: formateando imágenes"
	Iniseg.ImagenesStory dcStory
	Debug.Print "9/11 - Archivo story: corrigiendo interlineado"
	Iniseg.InterlineadoCorregido dcStory

	If iExportarNotas = vbYes Then
		Debug.Print "10A/11 - exportando notas a archivo externo"
		Iniseg.NotasPieExportar dcLibro
		Debug.Print "10B/11 - Archivo story: formateando notas"
		Iniseg.NotasPieMarcas dcStory, True
	Else
		Debug.Print "10/11 - Archivo story: formateando notas"
		Iniseg.NotasPieMarcas dcStory, False
	End If

	Debug.Print "11/11 - Conversión para story terminada"
	Set ConversionStory = dcStory
End Function






Sub ActualizandoVersion(dcArgumentDocument As Document)
' Actualiza el formato del archivo a la última versión para solucionar problemas de compatibilidades
'
    Dim iIndex As Integer

    ' Conversión de los campos INCLUDEPICTURE a imágenes
	For iIndex = dcArgumentDocument.Fields.Count To 1 Step -1
		With dcArgumentDocument.Fields(iIndex)
			If .Type = wdFieldIncludePicture Then
				.Unlink
			End If
		End With
	Next iIndex

    ' Al convertir el archivo a una versión moderna se les da a las imagenes las propiedades y métodos adecuados para su manipulación
	If dcArgumentDocument.CompatibilityMode < 15 Then dcArgumentDocument.Convert
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





Sub AutoFormateo(dcArgumentDocument As Document)
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

		dcArgumentDocument.Range.AutoFormat

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





Sub ComillasFormato(dcArgumentDocument As Document)
' Quita la negrita y cursiva de las comillas y las pasa a curvadas
'
	Dim bSmtQt As Boolean
	bSmtQt = Options.AutoFormatAsYouTypeReplaceQuotes
	Options.AutoFormatAsYouTypeReplaceQuotes = True
	RaMacros.FindAndReplaceClearParameters

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





Sub ImagenesLibro(dcArgumentDocument As Document)
' Formatea más cómodamente las imágenes
	' Las convierte de flotantes a inline (de shapes a inlineshapes)
	' Impide que aparezcan deformadas (mismo % relativo al tamaño original en alto y ancho)
	' Las centra
	' Impide que superen el ancho de página
'
	Dim inlShape As InlineShape, sngRealPageWidth As Single, sngRealPageHeight As Single, iIndex As Integer

	sngRealPageWidth = dcArgumentDocument.PageSetup.PageWidth - dcArgumentDocument.PageSetup.Gutter _
		- dcArgumentDocument.PageSetup.RightMargin - dcArgumentDocument.PageSetup.LeftMargin

	sngRealPageHeight = dcArgumentDocument.PageSetup.PageHeight _
		- dcArgumentDocument.PageSetup.TopMargin - dcArgumentDocument.PageSetup.BottomMargin _
		- dcArgumentDocument.PageSetup.FooterDistance - dcArgumentDocument.PageSetup.HeaderDistance

	' Se convierten los formatos extraños a imágenes 
		' NO FUNCIONA PORQUE CUANDO HAY CAMPOS DE UNA VERSIÓN ANTIGUA SE CORROMPE EL PORTAPAPELES
	' For iIndex = dcArgumentDocument.InlineShapes.Count To 1 Step -1
	' 	With dcArgumentDocument.Inlineshapes(iIndex)
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
	If dcArgumentDocument.Shapes.Count > 0 Then
		For iIndex = dcArgumentDocument.Shapes.Count To 1 Step -1
			With dcArgumentDocument.Shapes(iIndex)
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
	For Each inlShape In dcArgumentDocument.InlineShapes
		With inlShape
			If .Type = wdInlineShapePicture Then
				.ScaleHeight = .ScaleWidth
				.LockAspectRatio = msoTrue
				.Width = sngRealPageWidth
				' CON ESTO SE LE DA EL ANCHO ORIGINAL DE LA IMAGEN O EL DEL ANCHO DE PÁGINA, SI LO EXCEDE, EN VEZ DE HACER QUE OCUPE TODO EL ANCHO DE PÁGINA
				' If .Width / (.ScaleWidth / 100) > sngRealPageWidth Then .Width = sngRealPageWidth Else .ScaleWidth = 100
				If .Height / (.ScaleHeight / 100) > sngRealPageHeight - 15 Then .Height = sngRealPageHeight - 15

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





Sub ImagenesStory(dcArgumentDocument As Document)
' Hace que todas las imágenes sean enormes, para meterlas en el story
'
	Dim inlShape As InlineShape

	For Each inlShape In dcArgumentDocument.InlineShapes
		inlShape.Width = CentimetersToPoints(29)
	Next inlShape
End Sub





Sub InterlineadoCorregido(dcArgumentDocument As Document)
' Interlineado de 1,15 sin espaciado entre párrafos
	' Eliminar los espaciados verticales entre párrafos y aplica el interlineado correcto
'
	With dcArgumentDocument.Content.ParagraphFormat
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





Sub ParrafosSeparacionLibro(dcArgumentDocument As Document)
' TODO
	' Refactorizar con variables y recolocando el código
'
	' Mete dos saltos de línea manuales en los Heading 1, entre "Tema N" y el nombre del tema
	With dcArgumentDocument.Range.Find
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

		.Format = True
		.style = wdstyleheading1
		.Text = "([tT][eE][mM][aA] [0-9]{1;2})"
		.Replacement.Text = "\1^l^l"
		.Execute Replace:=wdReplaceAll
	End With

		' Formatea los saltos de línea y les da tamaño 10
	RaMacros.CleanSpaces dcArgumentDocument
	With dcArgumentDocument.Range.Find
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
	With dcArgumentDocument.Range.Find
		.Forward = True
		.Wrap = wdFindContinue
		.Format = True
		.MatchCase = False
		.MatchWholeWord = False
		.MatchAllWordForms = False
		.MatchSoundsLike = False
		.MatchWildcards = True

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





Sub ParrafosConversionStory(dcArgumentDocument As Document)
' Conversion de Word impreso a formato para Storyline
'
	RaMacros.FindAndReplaceClearParameters

	' Títulos 1 en minúsculas (para copiarlos más cómodamente a la primera diapositiva)
	dcArgumentDocument.Styles(wdstyleheading1).Font.AllCaps = False

	' Cambio del tamaño de Titulo 2 de 16 a 17
	With dcArgumentDocument.Styles(wdstyleheading2).Font
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

	' Eliminar ALLCAPS de los títulos 3 y 4
	dcArgumentDocument.Styles(wdstyleheading3).Font.AllCaps = False
	dcArgumentDocument.Styles(wdstyleheading4).Font.AllCaps = False

	' Poner el estilo quote centrado y sin espacio a derecha ni izquierda
	With dcArgumentDocument.Styles(wdStyleQuote).ParagraphFormat
		.LeftIndent = 0
		.RightIndent = 0
		.Alignment = wdAlignParagraphCenter
	End With



' Cambio de tamaño de parrafos de separacion
	' Listas: 4 a 2
	With dcArgumentDocument.Range.Find
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
	With dcArgumentDocument.Range.Find
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
	With dcArgumentDocument.Range.Find
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
	With dcArgumentDocument.Range.Find
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





Sub TitulosConTresEspacios(dcArgumentDocument As Document)
' Sustituye la tabulación en los títulos por 3 espacios
'
	With dcArgumentDocument.Range.Find
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





Sub TitulosDivididos(dcArgumentDocument As Document)
' Corta los títulos 2 para que no peguen contra el logo de Iniseg
'	- Título 2: 35 caractéres hasta logo Iniseg
'	- Título 3: 55 caractéres hasta logo Iniseg
'	- Título 4: 65 caractéres hasta logo Iniseg
'	- Título 5: 70 caractéres hasta logo Iniseg
'
	With dcArgumentDocument.Range.Find
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





Sub NotasPieMarcas(dcArgumentDocument As Document, bExportar As Boolean)
' Convierte las referencias de notas al pie al texto "NOTA_PIE-numNota"
	' para poder automatizar externamente su conversión en el .story
'
	Dim lContadorNotas As Long, bFound As Boolean, rgFindRange As Range, oEstiloNota As Font

	Set oEstiloNota = New Font
	lContadorNotas = dcArgumentDocument.Footnotes.StartingNumber

	RaMacros.FindAndReplaceClearParameters
	If bExportar Then
		With oEstiloNota
			.Name = "Swis721 Lt BT"
			.Bold = True
			.Color = -738148353
			.Superscript = True
		End With
		Do
			With dcArgumentDocument.Range.Find
				.ClearFormatting
				.Text = "^2"
				.Forward = True
				.Wrap = wdFindContinue
				.Format = False
				.MatchCase = False
				.MatchWholeWord = False
				.MatchAllWordForms = False
				.MatchSoundsLike = False
				.MatchWildcards = True
				.Replacement.Font = oEstiloNota
				.Replacement.Text = "NOTA_PIE-" & lContadorNotas
				.Execute Replace:=wdReplaceOne

				If .Found = True Then
					bFound = True
					lContadorNotas = lContadorNotas + 1
				Else
					bFound = False
				End If
			End With
		Loop While bFound = True

	Else
		Set rgFindRange = dcArgumentDocument.Content
		With rgFindRange.Find
			.ClearFormatting
			.Replacement.ClearFormatting
			.Text = "(^2)"
			.Forward = True
			.Wrap = wdFindContinue
			.Format = False
			.MatchCase = False
			.MatchWholeWord = False
			.MatchAllWordForms = False
			.MatchSoundsLike = False
			.MatchWildcards = True
			.Replacement.Text = "NOTA_PIE-\1"
		End With
		Do
			bFound = False
			rgFindRange.find.Execute Replace:=wdReplaceOne
			If rgFindRange.Find.Found Then
				rgFindRange.Expand wdParagraph
				If rgFindRange.End <> dcArgumentDocument.Content.End Then
					Set rgFindRange = rgFindRange.Next(Unit:=wdCharacter, Count:=1)
					rgFindRange.EndOf wdStory, wdExtend
					bFound = True
				End If
			End If
		Loop While bFound = True
	End If

	RaMacros.FindAndReplaceClearParameters

End Sub






Sub NotasPieExportar(dcArgumentDocument As Document)
' Exporta las notas a un archivo separado
' ToDo:
	' Convertir esta subrutina en una función de uso general:
		' Retornar el archivo de notas
'
	Dim dcNotas As Document, stFilename As String, stOriginalName As String, stOriginalExtension As String
	Dim lNotasNuevasInicio As Long, rgNotasNuevas As Range

	With dcArgumentDocument
		stOriginalName = Left(.Name, InStrRev(.Name, ".") - 1)
		stOriginalExtension = Right(.Name, Len(.Name) - InStrRev(.Name, ".") + 1)
		stFileName = .Path & Application.PathSeparator & "NOTAS " & stOriginalName & stOriginalExtension
	End With

	If Dir(stFileName) > "" Then
		Set dcNotas = Documents.Open(FileName:=stFileName, ConfirmConversions:=False, ReadOnly:=False, Revert:=False)
	Else
		Set dcNotas = RaMacros.SaveAsNewFile(dcArgumentDocument, "NOTAS ", "", False)
		dcNotas.Content.Text = "Notas al pie"
		With activedocument.Content.Paragraphs(1)
			.Style = wdStyleHeading1
			.Alignment = wdAlignParagraphCenter
		End With
	End If

	dcNotas.Content.InsertParagraphAfter
	Set rgNotasNuevas = dcNotas.Content.Paragraphs.Last.Range
	dcArgumentDocument.StoryRanges(wdFootnotesStory).Copy
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
		.NoSpaceBetweenParagraphsOfSameStyle = False
	End With
	With dcNotas.Styles(wdStyleList)
		.ParagraphFormat.SpaceAfter = 4
		.ParagraphFormat.SpaceBefore = 4
		.ParagraphFormat.Alignment = wdAlignParagraphLeft
		.NoSpaceBetweenParagraphsOfSameStyle = False
	End With

	RaMacros.CleanBasic dcNotas
	Iniseg.AutoFormateo dcNotas
	RaMacros.HyperlinksFormatting dcNotas
	RaMacros.HyperlinksOnlyDomain dcNotas
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






Sub BibliografiaSaltosDePagina(dcArgumentDocument As Document)
' Inserta un salto de página antes de cada biografía
'
	Dim scCurrentSection As Section, rgFindRange As Range

	For Each scCurrentSection In dcArgumentDocument.Sections
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







Sub BibliografiaExportar(dcArgumentDocument As Document)
' Exporta la bibliografía en archivos separados y la borra de dcArgumentDocument
'
	Dim dcBibliografia As Document, scCurrentSection As Section
	Dim rgFindRange As Range, rgTitulo As Range, stNombre As String

	For Each scCurrentSection In dcArgumentDocument.Sections
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
				' Iniseg.HeaderCopy dcArgumentDocument, dcBibliografia, 1
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
				' 		stNombre = dcArgumentDocument.Path & Application.PathSeparator _
				' 			& "BIBLIOGRAFÍA " & rgFindRange.Text
				' 	Else
				' 		Beep
				' 		stNombre = InputBox("Número de tema no encontrado, completar", "Bibliografía", "TEMA " & scCurrentSection.Index)
				' 		stNombre = dcArgumentDocument.Path & Application.PathSeparator _
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
						stNombre = dcArgumentDocument.Path & Application.PathSeparator _
							& "BIBLIOGRAFÍA " & rgTitulo.Text & ".pdf"
					Else
						Beep
						stNombre = InputBox("Número de tema no encontrado, completar", "Bibliografía", "TEMA " & scCurrentSection.Index)
						stNombre = dcArgumentDocument.Path & Application.PathSeparator _
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






