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
	Dim dEmpiece As Double, dFin As Double

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
	dEmpiece = Timer
	lEstilosBorrados = RaMacros.StylesDeleteUnused(dcOriginalFile, False)
	dFin = Timer
	Debug.Print dFin-dEmpiece & " segundos (" & CInt((dFin-dEmpiece)/60) _
		& " minutos) para borrar " & lEstilosBorrados & " estilos"

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
	' iExportar = MsgBox("¿Exportar cada tema en archivos separados?", vbYesNoCancel, "Opciones exportar")
	' If iExportar = vbCancel Then Exit Sub
	iExportar = vbYes

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

Sub Iniseg3PáginasVaciasVisibles()
	' Esta macro es una mala práctica y solo está para evitar confusiones por
		' falta de uniformidad en el uso de plantillas y estilos
	RaMacros.SectionsFillBlankPages ActiveDocument
	Debug.Print "Iniseg3PáginasVaciasVisibles terminada"
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
	RaMacros.ListsNoExtraNumeration dcLibro

	' Títulos y mayúsculas
	Debug.Print "5/17 - Archivo libro: Títulos sin AllCaps"
	For iContador = -3 To -10 Step -1
		dcLibro.Styles(iContador).Font.AllCaps = False
	Next iContador
	Debug.Print "6/17 - Archivo libro: Título 1 en mayúsculas"
	dcLibro.Styles(wdstyleheading1).Font.AllCaps = True

	Debug.Print "7/17 - Archivo libro: aplicando estilo correcto a hipervínculos"
	RaMacros.HyperlinksFormatting dcLibro, 1, 0
	Debug.Print "8.1/17 - Archivo libro: aplicando estilo correcto a notas al pie"
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
	Debug.Print "15.1/17 - Archivo libro: añadiendo saltos de sección antes de Títulos 1"
	RaMacros.SectionBreakBeforeHeading dcLibro, False, 4, 1
	If dcLibro.Sections.Count > 1 Then
		Debug.Print "15.2/17 - Archivo libro: mismo numbering rule de notas al pie en todas las secciones"
		RaMacros.FootnotesSameNumberingRule dcLibro, 3, -501
	End If
	Debug.Print "16/17 - Archivo libro: añadiendo saltos de página antes de Títulos de bibliografía"
	Iniseg.BibliografiaSaltosDePagina dcLibro

	Do While dcLibro.Paragraphs.Last.Range.Text = vbCr
		If dcLibro.Paragraphs.Last.Range.Delete = 0 Then Exit Do
	Loop

	Debug.Print "17/17 - Conversión a libro terminada"
	Set ConversionLibro = dcLibro
End Function

Function ConversionStory(dcLibro As Document, Optional ByVal iExportarNotas As Integer = 0, _
						Optional ByVal iExportarSeparados As Integer = 0, _
						Optional ByVal iNotasSeparadas As Integer = 0) _
	As Document
' Da el tamaño correcto a párrafos, imágenes y formatea marcas de pie de página
'
	Dim dcStory As Document
	Dim dcBibliografia As Document
	Dim iUltima As Integer
	Dim bNotasSeparadas As Boolean

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

	If iNotasSeparadas = 0 And dcLibro.Sections.Count > 1 Then
		iNotasSeparadas = MsgBox("¿Exportar las notas al pie de cada tema en archivos separados?", _
			vbYesNoCancel, "Opciones notas al pie")
		If iNotasSeparadas = vbCancel Then Exit Function
		If iNotasSeparadas = vbYes Then bNotasSeparadas = True Else bNotasSeparadas = False
	ElseIf iNotasSeparadas < 6 Or iNotasSeparadas > 7 Then
		Err.Raise Number:=513, Description:="iNotasSeparadas out of range"
	End If

	If dcLibro.Sections(1).Range.FootnoteOptions.NumberingRule = wdRestartSection Then
		bNotasSeparadas = True
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

	Debug.Print "4.1/" & iUltima & " - Archivo story: convirtiendo listas y campos LISTNUM a texto"
	dcStory.ConvertNumbersToText
	Debug.Print "4.2/" & iUltima & " - Archivo story: adaptando listas para Storyline"
	Iniseg.ListasParaStory dcStory
	
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
		RaMacros.TablesExportToPdf dcStory, "Tabla ", True, "Enlace a ", True, _
			dcStory.Name, wdStyleBlockQuotation, 17
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
			.Style = wdStyleBlockQuotation
			.Text = "Enlace a tabla"
			.Replacement.ParagraphFormat.Alignment = wdAlignParagraphCenter
			.Execute Replace:=wdReplaceAll
		End With
	Else
		Debug.Print "--- No hay tablas ---"
	End If

	If dcLibro.Footnotes.Count > 0 Then
		If iExportarNotas = vbYes Then
			Debug.Print iUltima - 2 & ".1/" & iUltima & " - exportando notas a archivo externo"
			Iniseg.NotasPieExportar dcLibro, bNotasSeparadas
				Debug.Print iUltima - 2 & ".2/" & iUltima & " - Archivo story: formateando notas"
			Iniseg.NotasPieMarcas dcStory, True
		Else
			Debug.Print iUltima - 2 & "/" & iUltima & " - Archivo story: formateando notas"
			Iniseg.NotasPieMarcas dcStory, False
		End If
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
				Optional ByVal iHeaderOption As Integer = 3)
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

	Application.ScreenUpdating = False
	
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
	Application.ScreenUpdating = True
End Sub





Sub ImagenesStory(dcArgument As Document)
' Hace que todas las imágenes sean enormes, para meterlas en el story
'
	Dim inlShape As InlineShape

	Application.ScreenUpdating = False
	For Each inlShape In dcArgument.InlineShapes
		inlShape.Width = CentimetersToPoints(29)
	Next inlShape
	Application.ScreenUpdating = True
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
'
	Dim lContador As Long
	Dim iStory As Integer, iSize As Integer, iSizeNext As Integer
	Dim rgStory As Range
	Dim pCurrent As Paragraph

	Application.ScreenUpdating = False
    For iStory = 1 To 5 Step 4
        On Error Resume Next
        Set rgStory = dcArgument.StoryRanges(iStory)
        If Err.Number = 0 Then
            On Error GoTo 0
			' El loop es para que pase por todos los textframe
			Do
				With rgStory.Find
					.ClearFormatting
					.Replacement.ClearFormatting
					.Forward = True
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

					' Mete un salto de línea en los títulos 1, entre "Tema N" y el nombre del tema
					.Format = True
					.style = wdstyleheading1
					.Text = "^13(*^13)"
					.Replacement.Text = "^l\1"
					.Execute Replace:=wdReplaceAll
					.Text = "([tT][eE][mM][aA] [0-9]{1;2})"
					.Replacement.Text = "\1 "
					.Execute Replace:=wdReplaceAll
					.Text = "([tT][eE][mM][aA] [0-9]{1;2}) @"
					.Replacement.Text = "\1^l^l"
					.Execute Replace:=wdReplaceAll
					' Formatea los saltos de línea y les da tamaño 10
					.Replacement.ClearFormatting
					.Replacement.Font.Size = 10
					.Text = "[^13^l]{2;}"
					.Replacement.Text = "^l^l"
					.Execute Replace:=wdReplaceAll
				End With
				RaMacros.FindAndReplaceClearParameters

				For lContador = rgStory.Paragraphs.Count - 1 To 1 Step -1
					Set pCurrent = rgStory.Paragraphs(lContador)
					' No se añaden párrafos de separación a los pies de imagen o el interior de tablas
					If pCurrent.Range.Tables.Count = 0 Then
						If pCurrent.Next.Range.Tables.Count = 0 Then
							If Not (pCurrent.style = dcArgument.styles(wdStyleCaption) _
								And pCurrent.Next.style = pCurrent.style) _
							Then
								iSize = GetSeparacionTamaño(pCurrent)
								iSizeNext = GetSeparacionTamaño(pCurrent.Next)
								pCurrent.Range.InsertParagraphAfter
								' Se mantiene el estilo actual si los párrafos adyacentes lo requieren, en caso contrario se asigna estilo Normal
								If (pCurrent.style = dcArgument.styles(wdStyleBlockQuotation) _
									Or pCurrent.style = dcArgument.styles(wdStyleQuote) _
									Or pCurrent.style = dcArgument.styles(wdStyleHeading1)) _
									And pCurrent.Next(2).style = pCurrent.style _
								Then
									pCurrent.Next.style = pCurrent.style
									' Separación de 10 puntos entre "Tema n" y el nombre del tema en Títulos 1
									If pCurrent.style = dcArgument.styles(wdStyleHeading1) Then
										iSize = 10
										iSizeNext = 10
									End If
								Else
									pCurrent.Next.style = wdStyleNormal
								End If
								If iSizeNext > iSize Then
									pCurrent.Next.Range.Font.Size = iSizeNext
								Else
									pCurrent.Next.Range.Font.Size = iSize
								End If
							End If
						End If	
					End If
				Next lContador
				If iStory = 5 And Not rgStory.NextStoryRange Is Nothing Then
					Set rgStory = rgStory.NextStoryRange
				Else
					Exit Do
				End If
			Loop
		Else
			On Error GoTo 0
		End If
	Next iStory
	Application.ScreenUpdating = True
End Sub
Function GetSeparacionTamaño(pArgument As Paragraph) As Integer
' Devuelve el tamaño de separación propio del tipo de párrafo pasado como argumento
'
	Dim dcParent As Document
	Set dcParent = pArgument.Parent
	With dcParent
		Select Case pArgument.style
			Case dcParent.Styles(wdStyleHeading1), dcParent.Styles(wdStyleHeading2)
				GetSeparacionTamaño = 11
			Case dcParent.Styles(wdStyleHeading3), dcParent.Styles(wdStyleHeading4)
				GetSeparacionTamaño = 8
			Case dcParent.Styles(wdStyleHeading5) To dcParent.Styles(wdStyleHeading9)
				GetSeparacionTamaño = 6
			Case dcParent.Styles(wdStyleNormal), dcParent.Styles(wdStyleCaption)
				GetSeparacionTamaño = 5
			Case dcParent.Styles(wdStyleQuote), dcParent.Styles(wdStyleBlockQuotation), _
					dcParent.Styles(wdStyleListParagraph), _
					dcParent.Styles(wdStyleList) To dcParent.Styles(wdStyleList5), _
					dcParent.Styles(wdStyleListBullet) To dcParent.Styles(wdStyleListBullet5), _
					dcParent.Styles(wdStyleListNumber) To dcParent.Styles(wdStyleListNumber5), _
					dcParent.Styles(wdStyleListContinue) To dcParent.Styles(wdStyleListContinue5)
				GetSeparacionTamaño = 4
			' Estilos desconocidos
			Case Else
				GetSeparacionTamaño = 5
		End Select
	End With
End Function

Sub TablasParrafosSeparacion(dcArgument As Document)
' Inserta un párrafo vacío y marcado antes de cada tabla
'
	Dim iCounter As Integer
	Dim rgTable As Range
	Dim tbCurrent As Table

	Application.ScreenUpdating = False
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
	Application.ScreenUpdating = True
End Sub

Sub ParrafosConversionStory(dcArgument As Document)
' Conversion de Word impreso a formato para Storyline
'

	Application.ScreenUpdating = False
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
	Application.ScreenUpdating = True
End Sub





Sub TitulosConTresEspacios(dcArgument As Document)
' Sustituye la tabulación en los títulos por 3 espacios
'
	Dim lstLevel As ListLevel

	' With dcArgument.Range.Find
	' 	.ClearFormatting
	' 	.Replacement.ClearFormatting
	' 	.Forward = True
	' 	.Wrap = wdFindContinue
	' 	.Format = True
	' 	.MatchCase = False
	' 	.MatchWholeWord = False
	' 	.MatchAllWordForms = False
	' 	.MatchSoundsLike = False
	' 	.MatchWildcards = True
	' 	.Text = "([0-9].)^t"
	' 	.Replacement.Text = "\1   "
	' 	.Style = wdstyleheading2
	' 	.Execute Replace:=wdReplaceAll
	' 	.Style = wdstyleheading3
	' 	.Execute Replace:=wdReplaceAll
	' 	.Style = wdstyleheading4
	' 	.Execute Replace:=wdReplaceAll
	' End With
	' RaMacros.FindAndReplaceClearParameters

	For Each lstLevel In dcStory.Styles("iniseg-lista_titulos").ListTemplate.ListLevels
		If lstLevel.NumberStyle <> wdListNumberStyleNone Then
			lstLevel.TrailingCharacter = wdTrailingNone
			lstLevel.NumberFormat = lstLevel.NumberFormat & "   "
		End If
	End With

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





Sub NotasPieMarcas(dcArgument As Document, ByVal bExportar As Boolean)
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

	Application.ScreenUpdating = False
	For lContadorNotas = dcArgument.Footnotes.Count To 1 Step -1
		'lReferencia = dcArgument.Footnotes.StartingNumber + dcArgument.Footnotes(lContadorNotas).Index - 1
		Set rgFootNote = dcArgument.Footnotes(lContadorNotas).Reference
		If bExportar Then
			rgFootNote.Text = "NOTA_PIE-" & lReferencia
		Else
			rgFootNote.Previous(wdCharacter, 1).InsertAfter "NOTA_PIE-" & lReferencia
		End If
		rgFootNote.Font = oEstiloNota
	Next lContadorNotas
	Application.ScreenUpdating = True
End Sub

Sub NotasPieExportar(dcArgument As Document, ByVal bDivide As Boolean, _
					Optional ByVal stSuffix As String = "Footnotes", _
					Optional ByVal stSectionSuffix As String = "Section", _
					Optional ByVal stTitle As String = "Footnotes")
' Exporta las notas a un archivo separado
' Args:
	' dcArgument: file from which the notes need to be extracted from
	' bDivide: if True, the notes of each section get exported to different files
	' Optional stSuffix As String = "Footnotes", _
	' Optional stSectionSuffix As String = "Section"
	' Optional stTitle As String = "Footnotes")

' ToDo:
	' Convertir esta subrutina en una función de uso general:
		' cambiar idioma
		' Retornar el archivo de notas
		' Implementar los argumentos opcionales
'
	Dim dcNotas As Document
	Dim stFilename As String, stOriginalName As String, stOriginalExtension As String
	Dim rgFind As Range
	Dim fnCurrent As Footnote
	Dim scCurrent As Section
	Dim bFirst As Boolean
	Dim lCounter As Long

	stOriginalName = Left(dcArgument.Name, InStrRev(dcArgument.Name, ".") - 1)
	stOriginalExtension = Right(dcArgument.Name, Len(dcArgument.Name) - InStrRev(dcArgument.Name, ".") + 1)
	bFirst = True

	For Each scCurrent In dcArgument.Sections
		If scCurrent.Range.Footnotes.Count > 0 Then
			If bDivide Then
				' Asigna el número de tema
				Set rgFind = scCurrent.Range
				RaMacros.FindAndReplaceClearParameters
				rgFind.Find.Execute FindText:="TEMA [0-9][0-9]", MatchWildcards:= True
				If Not rgFind.Find.Found Then rgFind.Find.Execute FindText:="TEMA [0-9]", MatchWildcards:= True

				If rgFind.Find.Found Then
					stFileName = rgFind.Text & " "
				Else
					Beep
					stFileName = InputBox("Nombre (número) de tema no encontrado, completar", _
										"NOTAS", "TEMA " & scCurrent.Index)
					stFileName = stFileName & " "
				End If
			End If

			If bDivide Or bFirst Then
				stFileName = dcArgument.Path & Application.PathSeparator & "NOTAS " _
							& stFileName & stOriginalName & stOriginalExtension
			End If

			If bDivide = False And bFirst And Dir(stFileName) > "" Then
				Set dcNotas = Documents.Open(FileName:=stFileName, ConfirmConversions:=False, _
											ReadOnly:=False, Revert:=False, Visible:=False)
				RaMacros.CopySecurity dcNotas, "0-", ""
			ElseIf bDivide Or (bFirst And bDivide = False) Then
				Set dcNotas = Documents.Add _
						(Template:= "C:\Users\Ra\Documents\Plantillas personalizadas de Office\iniseg.dotm", _
						Visible:= False)
				dcNotas.SaveAs2 stFilename
				Iniseg.HeaderCopy dcArgument, dcNotas, 1
				dcNotas.Content.Text = "Notas al pie"
				With dcNotas.Content.Paragraphs(1)
					.Style = wdStyleTitle
					.Alignment = wdAlignParagraphCenter
				End With
			End If
				
			If bDivide Or bFirst Then
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
			End If
			
			' Iteraration with a for each bugs out
			For lCounter = 1 To scCurrent.Range.Footnotes.Count
				Set fnCurrent = scCurrent.Range.Footnotes(lCounter)
				dcNotas.Content.InsertParagraphAfter
				Set rgFind = dcNotas.Content.Paragraphs.Last.Range
				rgFind.FormattedText = fnCurrent.Range.FormattedText
				rgFind.Style = wdStyleListContinue
				rgFind.Paragraphs(1).Style = wdStyleList
			Next lCounter
			bFirst = False
		End If
		If Not dcNotas Is Nothing And (bDivide Or scCurrent.Index = dcArgument.Sections.Count) Then
			RaMacros.CleanBasic dcNotas, 1
			Iniseg.AutoFormateo dcNotas
			RaMacros.HyperlinksFormatting dcNotas, 3, 1
			RaMacros.StylesNoDirectFormatting dcNotas
			dcNotas.Content.Select
			Selection.ClearCharacterDirectFormatting
			Selection.ClearParagraphDirectFormatting
			Do While dcNotas.Paragraphs.Last.Range.Text = vbCr
				If dcNotas.Paragraphs.Last.Range.Delete = 0 Then Exit Do
			Loop
			stFileName = Left(stFileName, InStrRev(stFileName, ".") - 1) & ".pdf"
			dcNotas.Save
			dcNotas.ExportAsFixedFormat OutputFileName:=stFileName, ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, _
				OptimizeFor:=wdExportOptimizeForPrint, Range:=wdExportAllDocument, Item:=wdExportDocumentWithMarkup, _
				CreateBookmarks:=wdExportCreateHeadingBookmarks
			dcNotas.Close wdDoNotSaveChanges
		End If
	Next scCurrent

	' THIS METHOD IS FASTER BUT LESS RELIABLE
	' dcNotas.Content.InsertParagraphAfter
	' Set rgFind = dcNotas.Content.Paragraphs.Last.Range
	' rgFind.FormattedText = dcArgument.StoryRanges(wdFootnotesStory).FormattedText
	' RaMacros.CleanBasic dcNotas, 1
	' rgFind.Style = wdStyleListContinue

	' With rgFind.Find
	' 	.ClearFormatting
	' 	.Replacement.ClearFormatting
	' 	.Forward = True
	' 	.Wrap = wdFindStop
	' 	.Format = True
	' 	.MatchCase = False
	' 	.MatchWholeWord = False
	' 	.MatchWildcards = False
	' 	.MatchSoundsLike = False
	' 	.MatchAllWordForms = False
	' 	.Font.Superscript = True
	' 	.Text = ""
	' 	.Replacement.Style = wdStyleList
	' 	.Replacement.Text = "marca_notas_pie"
	' 	.Execute Replace:=wdReplaceAll

	' 	.Format = False
	' 	.Replacement.ClearFormatting
	' 	.Text = "marca_notas_pie"
	' 	.Replacement.Text = ""
	' 	.Execute Replace:=wdReplaceAll
	' End With

	' With dcNotas.Styles(wdStyleListContinue)
	' 	.ParagraphFormat.SpaceAfter = 2
	' 	.ParagraphFormat.SpaceBefore = 2
	' 	.ParagraphFormat.Alignment = wdAlignParagraphLeft
	' 	.NoSpaceBetweenParagraphsOfSameStyle = True
	' End With
	' With dcNotas.Styles(wdStyleList)
	' 	.ParagraphFormat.SpaceAfter = 0
	' 	.ParagraphFormat.SpaceBefore = 5
	' 	.ParagraphFormat.Alignment = wdAlignParagraphLeft
	' 	.NoSpaceBetweenParagraphsOfSameStyle = False
	' End With

	' RaMacros.CleanBasic dcNotas, 1
	' Iniseg.AutoFormateo dcNotas
	' RaMacros.HyperlinksFormatting dcNotas, 3, 1
	' Do While dcNotas.Paragraphs.Last.Range.Text = vbCr
	' 	If dcNotas.Paragraphs.Last.Range.Delete = 0 Then Exit Do
	' Loop

	' dcNotas.Save
	' stFileName = Left(stFileName, InStrRev(stFileName, ".") - 1) & ".pdf"
	' dcNotas.ExportAsFixedFormat OutputFileName:=stFileName, ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, _
	' 	OptimizeFor:=wdExportOptimizeForPrint, Range:=wdExportAllDocument, Item:=wdExportDocumentWithMarkup, _
	' 	CreateBookmarks:=wdExportCreateHeadingBookmarks
	' dcNotas.Close wdSaveChanges
End Sub






Sub BibliografiaSaltosDePagina(dcArgument As Document)
' Inserta un salto de página antes de cada bibliografía
'
	Dim scCurrent As Section, rgFindRange As Range

	Application.ScreenUpdating = False
	For Each scCurrent In dcArgument.Sections
		Set rgFindRange = scCurrent.Range
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
	Next scCurrent
	Application.ScreenUpdating = True
End Sub

Sub BibliografiaExportar(dcArgument As Document)
' Exporta la bibliografía en archivos separados y la borra de dcArgument
'
	Dim dcBibliografia As Document, scCurrent As Section
	Dim rgFindRange As Range, rgTitulo As Range, stNombre As String

	For Each scCurrent In dcArgument.Sections
		Set rgFindRange = scCurrent.Range
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
			' Set dcBibliografia = Documents.Add("C:\Users\Ra\Documents\Plantillas personalizadas de Office\iniseg.dotm", Visible:= False)
			' Iniseg.HeaderCopy dcArgument, dcBibliografia, 1
			' rgFindRange.End = scCurrent.Range.End
			' dcBibliografia.Content.FormattedText = rgFindRange
			' rgFindRange.End = scCurrent.Range.End - 1
			' rgFindRange.Delete

			' Set rgFindRange = scCurrent.Range
			' With rgFindRange.Find
			' 	.MatchWildcards = True
			' 	.Execute FindText:="TEMA [0-9][0-9]"
			' 	If Not .Found Then .Execute FindText:="TEMA [0-9]"
			' 	If .Found Then
			' 		stNombre = dcArgument.Path & Application.PathSeparator _
			' 			& "BIBLIOGRAFÍA " & rgFindRange.Text
			' 	Else
			' 		Beep
			' 		stNombre = InputBox("Número de tema no encontrado, completar", "Bibliografía", "TEMA " & scCurrent.Index)
			' 		stNombre = dcArgument.Path & Application.PathSeparator _
			' 			& "BIBLIOGRAFÍA " & stNombre
			' 	End If
			' End With

			' dcBibliografia.Close wdSaveChanges

			' Asigna el número de tema
			Set rgTitulo = scCurrent.Range
			With rgTitulo.Find
				.MatchWildcards = True
				.Execute FindText:="TEMA [0-9][0-9]"
				If Not .Found Then .Execute FindText:="TEMA [0-9]"
				If .Found Then
					stNombre = dcArgument.Path & Application.PathSeparator _
						& "BIBLIOGRAFÍA " & rgTitulo.Text & ".pdf"
				Else
					Beep
					stNombre = InputBox("Número de tema no encontrado, completar", "Bibliografía", "TEMA " & scCurrent.Index)
					stNombre = dcArgument.Path & Application.PathSeparator _
						& "BIBLIOGRAFÍA " & stNombre & ".pdf"
				End If
			End With

			' Exporta el pdf
			rgFindRange.End = scCurrent.Range.End - 1
			rgFindRange.ExportAsFixedFormat2 _
				stNombre,wdExportFormatPDF,False,wdExportOptimizeForPrint,True, _
				wdExportDocumentWithMarkup,True,,wdExportCreateNoBookmarks,True,False,False,True

			' Borra la bibliografía de dcStory
			rgFindRange.Delete
		End If
	Next scCurrent
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






Sub ListasParaStory(dcArgument As Document)
' Convierte las listas de letras o números romanos a listas de números y les añade 
' una marca para poder cambiarlas externamente y de forma automatizada en el Story
'
	With dcArgument.Styles("iniseg-list_mixta").ListTemplate
		.ListLevels(2).NumberStyle = wdListNumberStyleArabic
		.ListLevels(2).NumberFormat = "A%2."
		.ListLevels(3).NumberStyle = wdListNumberStyleArabic
		.ListLevels(3).NumberFormat = "I%3."
		.ListLevels(4).NumberStyle = wdListNumberStyleArabic
		.ListLevels(4).NumberFormat = "a%4."
		.ListLevels(5).NumberStyle = wdListNumberStyleArabic
		.ListLevels(5).NumberFormat = "i%5."
	End With
End Sub