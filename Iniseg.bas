' Attribute VB_Name = "Iniseg"
Option Explicit

' Private iActual As Integer
' Private iUltima As Integer

Sub uiInisegConversionLibro()
	ConversionLibro ActiveDocument
End Sub
Sub uiInisegConversionStory()
	ConversionStory ActiveDocument
End Sub
Sub uiInisegConversionAutomaticaLibro()
	ConversionAutomaticaLibro ActiveDocument
End Sub
Sub uiInisegBibliografiaExportar()
	BibliografiaExportar ActiveDocument
End Sub






Sub Iniseg1Limpieza()
' Crea copia de seguridad del archivo original.
' Ejecuta limpieza (espacios, estilos innecesarios, etc.)
' Crea y deja abierto el archivo en formato libro para comenzar a darle los estilos
'
	Dim dcOriginal As Document, dcLibro As Document
	Dim rgActual As Range
	Dim stTextoOcultoMsg As String, stFileName As String
	Dim iTextosOcultos() As Integer, Dim i As Integer, iDeleteAnswer As Integer
	Dim lEstilosBorrados As Long, lPrimeraNotaAlPie As Long
	
	Set dcOriginal = ActiveDocument
	Set rgActual = dcOriginal.Content
	stFileName = dcOriginal.FullName
	lPrimeraNotaAlPie = 0

	' Borrar contenido innecesario
	iDeleteAnswer = MsgBox("¿Borrar contenido hasta el punto seleccionado?", vbYesNoCancel, "Borrar contenido")
	If iDeleteAnswer = vbCancel Then Exit Sub

	rgActual.Start = Selection.Start
	Debug.Print "1.1/14 - Copia de seguridad (0) del archivo original"
	RaMacros.FileCopy dcOriginal, "0-", ""

	If iDeleteAnswer = vbYes Then
		If rgActual.Footnotes.Count > 0 Then
			lPrimeraNotaAlPie = rgActual.FootnoteOptions.StartingNumber
			If rgActual.Footnotes(1).Index <> 1 Then
				lPrimeraNotaAlPie = lPrimeraNotaAlPie + rgActual.Footnotes(1).Index
			End If
		End If
		Debug.Print "1.2/14 - Borrando texto seleccionado"
		rgActual.End = rgActual.Start
		rgActual.Start = 0
		rgActual.Delete
	End If

	' Actualización del formato del archivo (soluciona problemas de compatibilidad con shapes y campos)
	If dcOriginal.CompatibilityMode < 15 Then
		Debug.Print "1.3/14 - Actualizando formato de archivo"
		RaMacros.FieldsUnlink dcOriginal
		' Al convertir el archivo a una versión moderna se les da a las imagenes las propiedades y métodos adecuados para su manipulación
		dcOriginal.Convert
	End If
	
	For Each rgActual In dcOriginal.StoryRanges
		Do
			rgActual.Font.Name = "Swis721 Lt BT"
			Set rgActual = rgActual.NextStoryRange
		Loop Until rgActual Is Nothing
	Next rgActual

	Debug.Print "2/14 - Creando archivo con plantilla Iniseg"
	Set dcLibro = Documents.Add("iniseg-wd")

	Debug.Print "3/14 - Copiando encabezados"
	Iniseg.HeaderCopy dcOriginal, dcLibro, 1

	Debug.Print "4/14 - Aplicando autoformateo"
	Iniseg.AutoFormateo dcOriginal
	Debug.Print "5.1/14 - Limpiando hiperenlaces para que solo figure su dominio"
	RaMacros.HyperlinksFormatting dcOriginal, 2, 0
	Debug.Print "5.2/14 - Limpiando espacios"
	RaMacros.CleanSpaces Nothing, True, dcOriginal

	Debug.Print "6/14 - Borrando encabezados y pies de página"
	RaMacros.HeadersFootersRemove dcOriginal

	Debug.Print "7.1/14 - Dando colores adecuados al texto"
	Iniseg.ColoresCorrectos dcOriginal
	Debug.Print "7.2/14 - Borrando texto oculto"
	iTextosOcultos = RaMacros.ClearHiddenText(dcOriginal, True,,,1)
	stTextoOcultoMsg = "No hay texto oculto"
	On Error GoTo NoTextosOcultos
	If iTextosOcultos(0) <> -501 Then
		Dim stStoryRanges(4) As String
		stStoryRanges(0) = "Texto principal"
		stStoryRanges(1) = "Notas a pie de página"
		stStoryRanges(2) = "Notas al final"
		stStoryRanges(3) = "Comentarios"
		stStoryRanges(4) = "Frames de texto"
		stTextoOcultoMsg = "Texto oculto en:"
		For i = 0 To UBound(iTextosOcultos)
			stTextoOcultoMsg = stTextoOcultoMsg & vbCrLf & vbTab & "- "
			If iTextosOcultos(i) - 1 < 4 Then
				stTextoOcultoMsg = stTextoOcultoMsg & stStoryRanges(iTextosOcultos(i) - 1)
			Else
				stTextoOcultoMsg = stTextoOcultoMsg & iTextosOcultos(i)
			End If
		Next i
	End If
	Debug.Print stTextoOcultoMsg

	Debug.Print "8/14 - Borrando estilos sin uso"
	lEstilosBorrados = RaMacros.StylesDeleteUnused(dcOriginal, False)

	Debug.Print "9/14 - Quitando estilos de la galería de estilos rápidos"
	Iniseg.EstilosEsconder dcOriginal

	' Copia de seguridad limpia
	Debug.Print "10/14 - Creando copia de seguridad limpia (01)"
	RaMacros.FileSaveAsNew dcOriginal, "01-", "", False, True

	' Guarda el archivo con nombre original, preparado para el siguiente paso
	Debug.Print "11.1/14 - Copiando contenido limpio al archivo con plantilla (archivo libro)"
	dcLibro.Content.FormattedText = dcOriginal.Content

	If lPrimeraNotaAlPie <> 0 Then
		Debug.Print "11.2/14 - Archivo libro: corrigiendo el número de comienzo de las notas al pie"
		dcLibro.Footnotes.StartingNumber = lPrimeraNotaAlPie
	End If

	Debug.Print "12/14 - Archivo original: cerrando"
	dcOriginal.Close wdDoNotSaveChanges
	Debug.Print "13/14 - Archivo libro: guardando"
	dcLibro.SaveAs2 stFileName
	dcLibro.Activate

	Debug.Print "14/14 - Iniseg1Limpieza terminada"
	Beep
	MsgBox lEstilosBorrados & " Estilos borrados" & vbCrLf _
		& "Revisar numeración de notas al pie, aplicar estilos y ejecutar Iniseg2"

	Exit Sub
NoTextosOcultos:
	On Error GoTo 0
	ReDim iTextosOcultos(0)
	iTextosOcultos(0) = -501
	Resume
End Sub







Sub Iniseg2LibroYStory()
' Llama a las macros de ConversionLibro e ConversionStory y da un aviso para seguir trabajando
	' Organizado de esta forma las macros de libro y story se pueden llamar por separado
'
	Dim dcLibro As Document, dcStory As Document
	Dim iNotasContinuas As Integer
	Dim bNotasExportar As Boolean
	Dim bNotasSeparadas As Boolean

	Set dcLibro = ActiveDocument

	iNotasContinuas = Iniseg.SetNotasOpciones(dcLibro, 0)
	bNotasExportar = Iniseg.SetNotasOpciones(dcLibro, 1)
	If bNotasExportar Then bNotasSeparadas = Iniseg.SetNotasOpciones(dcLibro, 2)

	Set dcLibro = Iniseg.ConversionLibro(dcLibro, iNotasContinuas)
	Debug.Print "A/4 - Archivo libro: salvando"
	dcLibro.Save

	Set dcStory = Iniseg.ConversionStory(dcLibro, bNotasExportar, bNotasSeparadas, False)
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







Sub Iniseg3PaginasVaciasVisibles()
' Exporta cada tema de libro a archivos separados e introduce páginas en blanco 
' tras las secciones que acaban en página impar para que terceros no se confundan
'
	Dim iDocSeparados As Integer
	Dim dcLibro As Document

	Set dcLibro = ActiveDocument
	iDocSeparados = MsgBox("¿Exportar cada tema en archivos separados?", vbYesNoCancel, "Opciones exportar")
	If iDocSeparados = vbCancel Then
		Exit Sub
	ElseIf iDocSeparados = vbYes Then
		Debug.Print "Archivo libro: exportando cada tema a archivos separados"
		RaMacros.SectionsExportEachToFiles dcLibro, False, True, False,, " TEMA "
	End If

	' Esto es una mala práctica y solo está para evitar confusiones por
	' falta de uniformidad en el uso de plantillas y estilos
	Debug.Print "Archivo libro: insertando páginas en blanco en los cambios de sección"
	RaMacros.SectionsFillBlankPages dcLibro
	Debug.Print "Iniseg3PaginasVaciasVisibles terminada"
End Sub










Function ConversionLibro(dcLibro As Document, _
						Optional iNotasContinuas As Integer = -501) _
	As Document
' Realiza la limpieza necesaria y formatea correctamente
' Args:
	' iNotasContinuas:
		' -501: default (se preguntará)
		' 0: wdRestartContinuous
		' 1: wdRestartSection
		' 3: copia el numbering rule de la primera sección en las demás
'
	Dim iContador As Integer, iUltima As Integer

	If Not (iNotasContinuas = 0 Or iNotasContinuas = 1 Or iNotasContinuas = 3) Then
		If iNotasContinuas = -501 Then
			iNotasContinuas = Iniseg.SetNotasOpciones(dcLibro, 0)
		Else
			Err.Raise 513,, "iNotasContinuas fuera de rango"
		End If
	End If

	iUltima = 17

	Debug.Print "1/" & iUltima & " - Archivo libro: haciendo copia de seguridad (1)"
	RaMacros.FileSaveAsNew dcLibro, "1-", "", False, True
	Debug.Print "2/" & iUltima & " - Archivo libro: limpieza básica"
	RaMacros.CleanBasic Nothing, True, True, dcLibro

	Debug.Print "3/" & iUltima & " - Archivo libro: títulos sin puntuación"
	RaMacros.HeadingsNoPunctuation dcLibro
	Debug.Print "4.1/" & iUltima & " - Archivo libro: títulos sin numeración repetida"
	RaMacros.HeadingsNoNumeration dcLibro
	Debug.Print "4.2/" & iUltima & " - Archivo libro: listas sin numeración repetida"
	RaMacros.ListsNoExtraNumeration dcLibro

	' Títulos y mayúsculas
	Debug.Print "5/" & iUltima & " - Archivo libro: Títulos sin AllCaps"
	For iContador = -3 To -10 Step -1
		dcLibro.Styles(iContador).Font.AllCaps = False
	Next iContador
	Debug.Print "6/" & iUltima & " - Archivo libro: Título 1 en mayúsculas"
	dcLibro.Styles(wdstyleheading1).Font.AllCaps = True

	Debug.Print "7/" & iUltima & " - Archivo libro: formateando comillas"
	Iniseg.ComillasFormato dcLibro
	Debug.Print "8/" & iUltima & " - Archivo libro: sustituyendo formatos directos por estilos"
	RaMacros.StylesNoDirectFormatting dcLibro, Nothing, dcLibro.Styles(wdStyleStrong)
	Debug.Print "9/" & iUltima & " - Archivo libro: Aplicando estilo correcto a hipervínculos"
	RaMacros.HyperlinksFormatting dcLibro, 1, 0
	Debug.Print "10.1/" & iUltima & " - Archivo libro: Aplicando estilo correcto a notas al pie"
	If dcLibro.Footnotes.Count > 0 Then
		RaMacros.FootnotesFormatting dcLibro
		Debug.Print "10.2/" & iUltima & " - Archivo libro: sangrando notas al pie"
		RaMacros.FootnotesHangingIndentation dcLibro, 0.5, wdStyleFootnoteText
	Else
		Debug.Print "---No hay notas al pie---"
	End If
	Debug.Print "11/" & iUltima & " - Archivo libro: corrigiendo limpieza e interlineado"
	Iniseg.InterlineadoCorregido dcLibro
	RaMacros.CleanBasic Nothing, False, True, dcLibro
	' Lo siguiente es demasiado agresivo, devuelve numeraciones y cambia cosas sin ton ni son
	' dcLibro.Content.Select
	' Selection.ClearCharacterDirectFormatting
	' Selection.ClearParagraphDirectFormatting

	Debug.Print "12/" & iUltima & " - Archivo libro: formateando imágenes"
	Iniseg.ImagenesLibro dcLibro

	Debug.Print "13/" & iUltima & " - Archivo libro: añadiendo párrafos de separación"
	Iniseg.ParrafosSeparacionLibro dcLibro
	Debug.Print "14/" & iUltima & " - Archivo libro: añadiendo párrafos de separación antes de tablas"
	Iniseg.TablasParrafosSeparacion dcLibro
	Debug.Print "15.1/" & iUltima & " - Archivo libro: añadiendo saltos de sección antes de Títulos 1"
	RaMacros.SectionBreakBeforeHeading dcLibro, False, 4, 1
	If dcLibro.Sections.Count > 1 Then
		Debug.Print "15.2/" & iUltima & " - Archivo libro: mismo numbering rule de notas al pie en todas las secciones"
		RaMacros.FootnotesSameNumberingRule dcLibro, iNotasContinuas, -501
	End If
	Debug.Print "16/" & iUltima & " - Archivo libro: añadiendo saltos de página antes de Títulos de bibliografía"
	Iniseg.BibliografiaSaltosDePagina dcLibro

	' Borrando último párrafo vacío
	Do While dcLibro.Paragraphs.Last.Range.Text = vbCr
		If dcLibro.Paragraphs.Last.Range.Delete = 0 Then Exit Do
	Loop

	Debug.Print iUltima & "/" & iUltima & " - Conversión a libro terminada"
	Set ConversionLibro = dcLibro
End Function






Function ConversionStory(dcLibro As Document, _
						Optional ByVal bNotasExportar As Boolean, _
						Optional ByVal bNotasSeparadas As Boolean, _
						Optional ByVal bPreguntar As Boolean = True _
) As Document
' Da el tamaño correcto a párrafos, imágenes y formatea marcas de pie de página
'
	Dim dcStory As Document
	Dim dcBibliografia As Document
	Dim iUltima As Integer
	Dim stName As String

	If bPreguntar _
		And Not bNotasExportar _
		And Not bNotasSeparadas _
		And dcLibro.FootNotes.Count > 0 _
	Then
		bNotasExportar = Iniseg.SetNotasOpciones(dcLibro, 1)
		If bNotasExportar Then bNotasSeparadas = Iniseg.SetNotasOpciones(dcLibro, 2)
	End If

	iUltima = 9
	If bNotasExportar Then iUltima = iUltima + 1
	If dcLibro.Sections.Count > 1 Then iUltima = iUltima + 1

	Debug.Print "1/" & iUltima & " - Archivo story: creando"
	Set dcStory = RaMacros.FileSaveAsNew(dcLibro, "2-", "", True, True)
	dcStory.ActiveWindow.Visible = True

	Debug.Print "2.1/" & iUltima & " - Archivo story: marcando referencias bibliográficas"
	Iniseg.BibliografiaMarcarReferencias dcStory
	Debug.Print "2.2/" & iUltima & " - Archivo story: exportando y borrando bibliografías"
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

	Debug.Print "4/" & iUltima & " - Archivo story: títulos con 3 espacios en vez de tabulación"
	Iniseg.TitulosConTresEspacios dcStory
	Debug.Print "5.1/" & iUltima & " - Archivo story: adaptando listas para Storyline"
	Iniseg.ListasParaStory dcStory
	Debug.Print "5.2/" & iUltima & " - Archivo story: convirtiendo listas y campos LISTNUM a texto"
	dcStory.ConvertNumbersToText
	Debug.Print "5.3/" & iUltima & " - Archivo story: títulos divididos para no solaparse con el logo en la diapositiva"
	Iniseg.TitulosDivididos dcStory
	
	Debug.Print "6/" & iUltima & " - Archivo story: adaptando el tamaño de párrafos"
	Iniseg.ParrafosConversionStory dcStory
	Iniseg.TitulosConTresEspacios dcStory

	If dcStory.Tables.Count > 0 Then
		Debug.Print "7/" & iUltima & " - Archivo story: transformando/exportando tablas"
		' RaMacros.TablesConvertToImage dcStory
		stName = Left$(dcStory.Name, InStrRev(dcStory.Name, ".") - 1)
		stName = Right$(stName, Len(stName) - 2)
		RaMacros.TablesExportToPdf dcStory, stName, "Tabla ", True, "Enlace a ",,, wdStyleBlockQuotation, 17, True
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
		Debug.Print "7/" & iUltima & "--- No hay tablas ---"
	End If
	Debug.Print "8/" & iUltima & " - Archivo story: formateando imágenes"
	Iniseg.ImagenesStory dcStory
	Debug.Print "9/" & iUltima & " - Archivo story: corrigiendo interlineado"
	Iniseg.InterlineadoCorregido dcStory

	If dcLibro.Footnotes.Count > 0 Then
		If bNotasExportar Then
			Debug.Print "10.1/" & iUltima & " - Exportando notas a archivo/s externo"
			Iniseg.NotasPieExportar dcLibro, bNotasSeparadas
			Debug.Print "10.2/" & iUltima & " - Archivo story: formateando notas"
		Else
			Debug.Print "10/" & iUltima & " - Archivo story: formateando notas"
		End If
		Iniseg.NotasPieMarcas dcStory, bNotasExportar
	End If

	If dcLibro.Sections.Count > 1 Then
		Debug.Print iUltima & "/" & iUltima & " - Archivo story: exportando en archivos separados"
		RaMacros.SectionsExportEachToFiles dcStory, False, True, False,, "-tema_"
	End If

	Debug.Print "Conversión para story terminada"
	Set ConversionStory = dcStory
End Function







Function SetNotasOpciones(dcArg As Document, iPregunta As Integer)
' Pregunta para adjudicar valores a las variables para las notas al pie
' Args:
	' iPregunta: escoge qué pregunta hacer
		' 0: iNotasContinuas (integer)
		' 1: bNotasExportar (boolean)
		' 2: bNotasSeparadas (boolean)
'
	If dcArg.Footnotes.Count < 1 Then
		SetNotasOpciones = False
		If iPregunta = 0 Then SetNotasOpciones = 3
		Exit Function
	End If
	
	If iPregunta < 0 Or iPregunta > 2 Then
		Err.Raise 513,, "iPregunta fuera de rango, debe estar entre 0 y 2"
	End If

	Dim stTextoPregunta(2) As String
	Dim iRespuesta As Integer

	stTextoPregunta(0) = "¿Mantener constante la numeración de notas para cada tema?"
	stTextoPregunta(1) = "¿Exportar notas al pie de página a pdf?"
	stTextoPregunta(2) = "¿Exportar notas al pie de página de cada tema en archivos separados?"

	iRespuesta = MsgBox(stTextoPregunta(iPregunta), vbYesNoCancel, "Opciones notas")
	If iRespuesta = vbCancel Then End

	If iPregunta = 0 Then
		SetNotasOpciones = iRespuesta - 6
	Else
		SetNotasOpciones = False
		If iRespuesta = vbYes Then SetNotasOpciones = True
	End If
End Function







Sub HeaderCopy(dcOriginal As Document, _
				dcTarget As Document, _
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
		stOriginalHeader = Replace(dcOriginal.Sections(1).Headers(1).Range.Text, vbLf, "")
		stOriginalHeader = Trim(Replace(stOriginalHeader, vbCr, ""))
	End If

	For iHeader = 1 To 3
		If iHeaderOption > 1 Then
			If iHeaderOption = 2 And iHeader = 2 Then
				stOriginalHeader = Replace(dcOriginal.Sections(1).Headers(1).Range.Text, vbLf, "")
			Else
				stOriginalHeader = Replace(dcOriginal.Sections(1).Headers(iHeader).Range.Text, vbLf, "")
			End If
			stOriginalHeader = Trim(Replace(stOriginalHeader, vbCr, ""))
		End If

		With dcTarget.Sections(1).Headers(iHeader).Range.Find
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

		dcTarget.Sections(1).Headers(iHeader).Range.Case = wdLowerCase
		dcTarget.Sections(1).Headers(iHeader).Range.Case = wdTitleWord
	Next iHeader
End Sub





Sub AutoFormateo(dcArg As Document)
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

		dcArg.Range.AutoFormat

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





Sub ComillasFormato(dcArg As Document)
' Quita la negrita y cursiva de las comillas y las pasa a curvadas
'
	Dim bSmtQt As Boolean
	bSmtQt = Options.AutoFormatAsYouTypeReplaceQuotes
	Options.AutoFormatAsYouTypeReplaceQuotes = True
	RaMacros.FindAndReplaceClearParameters

	With dcArg.Range.Find
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





Sub ImagenesLibro(dcArg As Document)
' Formatea más cómodamente las imágenes
	' Las convierte de flotantes a inline (de shapes a inlineshapes)
	' Impide que aparezcan deformadas (mismo % relativo al tamaño original en alto y ancho)
	' Las centra
	' Impide que superen el ancho de página
'
	Dim inlShape As InlineShape, sngRealPageWidth As Single, sngRealPageHeight As Single, iIndex As Integer

	
	sngRealPageWidth = dcArg.PageSetup.PageWidth - dcArg.PageSetup.Gutter _
		- dcArg.PageSetup.RightMargin - dcArg.PageSetup.LeftMargin

	sngRealPageHeight = dcArg.PageSetup.PageHeight _
		- dcArg.PageSetup.TopMargin - dcArg.PageSetup.BottomMargin _
		- dcArg.PageSetup.FooterDistance - dcArg.PageSetup.HeaderDistance

	' Se convierten todas de shapes a inlineshapes
	If dcArg.Shapes.Count > 0 Then
		For iIndex = dcArg.Shapes.Count To 1 Step -1
			With dcArg.Shapes(iIndex)
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
	For Each inlShape In dcArg.InlineShapes
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





Sub ImagenesStory(dcArg As Document)
' Hace que todas las imágenes sean enormes, para meterlas en el story
'
	Dim inlShape As InlineShape

	For Each inlShape In dcArg.InlineShapes
		inlShape.Width = CentimetersToPoints(29)
	Next inlShape
End Sub





Sub InterlineadoCorregido(dcArg As Document)
' Interlineado de 1,15 sin espaciado entre párrafos
	' Eliminar los espaciados verticales entre párrafos y aplica el interlineado correcto
'
	With dcArg.Content.ParagraphFormat
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






Sub ParrafosSeparacionLibro(dcArg As Document)
' Inserta párrafos vacíos de separación
'
	Dim bFound As Boolean
	Dim lContador As Long
	Dim iStory As Integer, iSize As Integer, iSizeNext As Integer
	Dim rgStory As Range
	Dim pCurrent As Paragraph

	With dcArg.Content.Find
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
		.Text = "(^13)^m^13"
		.Replacement.Text = "\1"
		.Execute Replace:=wdReplaceAll
		.Text = "^m"
		.Replacement.Text = ""
		.Execute Replace:=wdReplaceAll

		' Mete un salto de línea en los títulos 1, entre "Tema N" y el nombre del tema
		.Format = True
		.style = wdstyleheading1
		.Text = "^13(*^13)"
		.Replacement.Text = "^l\1"
		.Execute Replace:=wdReplaceAll
		.Text = "[tT][eE][mM][aA]([0-9]{1;2})"
		.Replacement.Text = "Tema \1"
		.Execute Replace:=wdReplaceAll
		.Text = "[tT][eE][mM][aA] @([0-9]{1;2})"
		.Replacement.Text = "Tema \1 "
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
	
	' Eliminar signos de puntuación que pudiera haber tras el número de tema
	Set rgStory = dcArg.Content
	With rgStory.Find
		.ClearFormatting
		.Replacement.ClearFormatting
		.Forward = True
		.Format = True
		.style = wdstyleheading1
		.MatchCase = False
		.MatchWholeWord = False
		.MatchWildcards = False
		.MatchSoundsLike = False
		.MatchAllWordForms = False
		.Text = ""
	End With
	Do
		If Not rgStory.Find.Execute Then Exit Do
		RaMacros.CleanSpaces rgStory
		rgStory.Start = rgStory.End
		rgStory.EndOf wdStory, wdExtend
	Loop Until rgStory.Start = dcArg.Content.End - 1

	' Resto de párrafos
	For iStory = 1 To 5 Step 4
		On Error Resume Next
		Set rgStory = dcArg.StoryRanges(iStory)
		If Err.Number = 0 Then
			On Error GoTo 0
			' El loop es para que pase por todos los textframe
			Do
				For lContador = rgStory.Paragraphs.Count - 1 To 1 Step -1
					Set pCurrent = rgStory.Paragraphs(lContador)
					' No se añaden párrafos de separación a los pies de imagen o el interior de tablas
					If pCurrent.Range.Tables.Count = 0 And Not RaMacros.RangeIsField(pCurrent.Range) Then
						If pCurrent.Next.Range.Tables.Count = 0 Then
							If Not (pCurrent.style = dcArg.styles(wdStyleCaption) _
								And pCurrent.Next.style = pCurrent.style) _
							Then
								iSize = GetSeparacionTamaño(pCurrent)
								iSizeNext = GetSeparacionTamaño(pCurrent.Next)
								pCurrent.Range.InsertParagraphAfter
								' Se mantiene el estilo actual si los párrafos adyacentes lo requieren, en caso contrario se asigna el estilo iniseg_separacion pertinente
								If (pCurrent.style = dcArg.styles(wdStyleBlockQuotation) _
									Or pCurrent.style = dcArg.styles(wdStyleQuote)) _
									And pCurrent.Next(2).style = pCurrent.style _
								Then
									pCurrent.Next.style = pCurrent.style
									pCurrent.Next.Range.Font.Size = iSizeNext
								Else
									If iSizeNext > iSize Then iSize = iSizeNext
									If RaMacros.StyleExists(dcArg, _
										"iniseg_separacion" & iSize) _
									Then
										pCurrent.Next.style = "iniseg_separacion" & iSize
									Else
										dcArg.Styles.Add "iniseg_separacion" & iSize, wdStyleTypeParagraph
										dcArg.Styles("iniseg_separacion" & iSize).BaseStyle = wdStyleNormal
										dcArg.Styles("iniseg_separacion" & iSize).Font.Size = iSize
										If iSize > 5 Then
											With dcArg.Styles("iniseg_separacion" & iSize).ParagraphFormat
												.KeepWithNext = True
												.KeepTogether = True
												.WidowControl = True
											End With
										End If
										pCurrent.Next.style = "iniseg_separacion" & iSize
									End If
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
End Sub

Function GetSeparacionTamaño(pArg As Paragraph) As Integer
' Devuelve el tamaño de separación propio del tipo de párrafo pasado como argumento
'
	Dim dcParent As Document
	Set dcParent = pArg.Parent
	With dcParent
		Select Case pArg.style
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

Sub TablasParrafosSeparacion(dcArg As Document)
' Inserta un párrafo vacío y marcado antes de cada tabla
'
	Dim iCounter As Integer
	Dim rgTable As Range
	Dim tbCurrent As Table

	For iCounter = 1 To dcArg.Tables.Count Step 1
		Set tbCurrent = dcArg.Tables(iCounter)
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
			If tbCurrent.Range.End <> dcArg.StoryRanges(tbCurrent.Range.StoryType).End _
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

Sub ParrafosConversionStory(dcArg As Document)
' Conversion de Word impreso a formato para Storyline
'
	Dim iLibro(3) As Integer, iStory(3) As Integer, i As Integer
	Dim rgFind As Range
	Dim stLibro As String, stStory As String

	iLibro(0) = 4
	iLibro(1) = 5
	iLibro(2) = 8
	iLibro(3) = 11

	iStory(0) = 2
	iStory(1) = 4
	iStory(2) = 6
	iStory(3) = 8

	' Cambio del tamaño de Titulo 2 de 16 a 17
	With dcArg.Styles(wdstyleheading2).Font
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
	dcArg.Styles(wdstyleheading3).Font.AllCaps = False
	dcArg.Styles(wdstyleheading4).Font.AllCaps = False

	' Poner el estilo quote centrado y sin espacio a derecha ni izquierda
	With dcArg.Styles(wdStyleQuote).ParagraphFormat
		.LeftIndent = 0
		.RightIndent = 0
		.Alignment = wdAlignParagraphCenter
	End With

	' Cambio de tamaño de parrafos de separacion
	For i = 0 To uBound(iLibro)
		stLibro = "iniseg_separacion" & iLibro(i)
		stStory = "iniseg_separacion" & iStory(i)
		If RaMacros.StyleSubstitution(dcArg, stLibro, stStory, False) = 2 Then
			dcArg.Styles.Add stStory, wdStyleTypeParagraph
			With dcArg.Styles(stStory)
				.BaseStyle = dcArg.Styles(stLibro)
				.Font.Size = iStory(i)
			End With
			dcArg.Styles(stStory).Font.Size = iStory(i)
			If RaMacros.StyleSubstitution(dcArg, stLibro, stStory, False) <> 0 Then
				Debug.Print "Error creando sustituyendo " & stLibro & " por " & stStory
			End If
		End If
	Next i
End Sub





Sub TitulosConTresEspacios(dcArg As Document)
' Sustituye la tabulación en los títulos por 3 espacios
'
	Dim lstLevel As ListLevel

	For Each lstLevel In dcArg.Styles("iniseg-lista_titulos").ListTemplate.ListLevels
		If lstLevel.NumberStyle <> wdListNumberStyleNone Then
			lstLevel.TrailingCharacter = wdTrailingNone
			lstLevel.NumberFormat = lstLevel.NumberFormat & "   "
		End If
	Next lstLevel
End Sub





Sub TitulosDivididos(dcArg As Document)
' Corta los títulos 2 para que no peguen contra el logo de Iniseg
'	- Título 2: 35 caractéres hasta logo Iniseg
'	- Título 3: 55 caractéres hasta logo Iniseg
'	- Título 4: 65 caractéres hasta logo Iniseg
'	- Título 5: 70 caractéres hasta logo Iniseg
'
	With dcArg.Range.Find
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
		If dcArg.Styles(wdstyleheading2).ListLevelNumber = 0 Then
			.Text = "(?{35;}) "
			.Replacement.Text = "\1^13"
		Else
			.Text = "(?{30;}) "
			.Replacement.Text = "\1^13      "
		End If
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





Sub NotasPieMarcas(dcArg As Document, ByVal bExportar As Boolean)
' Convierte las referencias de notas al pie al texto "NOTA_PIE-numNota"
	' para poder automatizar externamente su conversión en el .story
' Args:
	' bExportar: si es true las notas se borran y su referencia se sustituye por el texto
		' si es false la referencia de la nota no se borra, se le aplica el atributo "hidden"
'
	Dim lNota As Long, lReferencia As Long
	Dim iSection As Integer
	Dim rgFootNote As Range
	Dim oEstiloNota As Font
	Dim scCurrent As Section

	Set oEstiloNota = New Font
	With oEstiloNota
		.Name = "Swis721 Lt BT"
		.Bold = True
		.Color = -738148353
		.Superscript = False
	End With


	For iSection = dcArg.Sections.Count To 1 Step -1
		Set scCurrent = dcArg.Sections(iSection)
		lReferencia = RaMacros.SectionGetFirstFootnoteNumber(dcArg, scCurrent.Index) _
			+ scCurrent.Range.Footnotes.Count
		For lNota = scCurrent.Range.Footnotes.Count To 1 Step -1
			lReferencia = lReferencia - 1
			Set rgFootNote = scCurrent.Range.Footnotes(lNota).Reference
			If bExportar Then
				rgFootNote.Text = "NOTA_PIE-" & lReferencia
			Else
                Set rgFootNote = rgFootNote.Previous(wdCharacter, 1)
                rgFootNote.Collapse wdCollapseEnd
                rgFootNote.InsertAfter ("NOTA_PIE-" & lReferencia)
                scCurrent.Range.Footnotes(lNota).Reference.Font.Hidden = True
			End If
			rgFootNote.Font = oEstiloNota
		Next lNota
	Next iSection

End Sub

Sub NotasPieExportar(dcArg As Document, _
					ByVal bDivide As Boolean, _
					Optional ByVal stSuffix As String = "Footnotes", _
					Optional ByVal stSectionSuffix As String = "Section", _
					Optional ByVal stTitle As String = "Footnotes")
' Exporta las notas a un archivo separado
' Args:
	' dcArg: file from which the notes need to be extracted from
	' bDivide: if True, the notes of each section get exported to different files
	' Optional stSuffix As String = "Footnotes", _
	' Optional stSectionSuffix As String = "Section"
	' Optional stTitle As String = "Footnotes")

' ToDo:
	' Convertir esta subrutina en una función de uso general:
		' cambiar idioma
		' Retornar el archivo de notas / array de archivos
		' Implementar los argumentos opcionales
		' implementar un parámetro para elegir si exportar en pdf o no
'
	Dim dcNotas As Document
	Dim stFileName As String, stOriginalName As String
	Dim rgFind As Range
	Dim fnCurrent As Footnote
	Dim scCurrent As Section
	Dim bFirst As Boolean
	Dim lCounter As Long, lStartingFootnote

	stOriginalName = Left$(dcArg.Name, InStrRev(dcArg.Name, ".") - 1)
	bFirst = True
	For Each scCurrent In dcArg.Sections
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
				Set dcNotas = Documents.Add _
						(Template:= "iniseg-wd", _
						Visible:= False)

				Iniseg.HeaderCopy dcArg, dcNotas, 1

				If bDivide Then
					dcNotas.Content.Text = RTrim("NOTAS AL PIE " & stFileName)
				Else
					dcNotas.Content.Text = "Notas al pie"
				End If

				With dcNotas.Content.Paragraphs(1)
					.Style = wdStyleTitle
					.Alignment = wdAlignParagraphCenter
				End With

				stFileName = dcArg.Path & Application.PathSeparator & "NOTAS " _
					& stFileName & stOriginalName
				dcNotas.SaveAs2 stFileName

				lStartingFootnote = _
					RaMacros.SectionGetFirstFootnoteNumber(dcArg, scCurrent.Index)

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
				bFirst = False
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
		End If
		
		If Not dcNotas Is Nothing And (bDivide Or scCurrent.Index = dcArg.Sections.Count) Then
			dcNotas.ListParagraphs(1).Range.ListFormat.ListTemplate.ListLevels(1).StartAt = lStartingFootnote
			RaMacros.CleanBasic dcNotas.Content, True, True, dcNotas
			Iniseg.AutoFormateo dcNotas
			RaMacros.HyperlinksFormatting dcNotas, 3, 1
			RaMacros.StylesNoDirectFormatting dcNotas, Nothing, dcNotas.Styles(wdStyleStrong)
			dcNotas.Content.Select
			Selection.ClearCharacterDirectFormatting
			Selection.ClearParagraphDirectFormatting
			Do While dcNotas.Paragraphs.Last.Range.Text = vbCr
				If dcNotas.Paragraphs.Last.Range.Delete = 0 Then Exit Do
			Loop
			dcNotas.Save
			dcNotas.ExportAsFixedFormat2 _
				OutputFileName:=stFileName, _
				ExportFormat:=wdExportFormatPDF, _
				OpenAfterExport:=False, _
				OptimizeFor:=wdExportOptimizeForPrint, _
				Range:=wdExportAllDocument, _
				Item:=wdExportDocumentWithMarkup, _
				IncludeDocProps:=True, _
				CreateBookmarks:=wdExportCreateWordBookmarks, _
				DocStructureTags:=True, _
				BitmapMissingFonts:=False, _
				UseISO19005_1:=False, _
				OptimizeForImageQuality:=True
			dcNotas.Close wdDoNotSaveChanges
			Set dcNotas = Nothing
		End If
	Next scCurrent
End Sub






Sub BibliografiaMarcarReferencias(dcArg As Document)
' Marcar los campos de referencias bibliograficas con el texto "NOT_BLI-[numNota]"
' para poder automatizar externamente su conversión en el .story
'
	Dim i As Integer

	For i = dcArg.Fields.Count To 1 Step -1
		If ActiveDocument.Fields(i).Type = wdFieldCitation Then
			ActiveDocument.Fields(i).Result.Previous(wdCharacter, 2).InsertAfter "not_bib-"
		End If
	Next i
End Sub







Sub BibliografiaSaltosDePagina(dcArg As Document)
' Inserta un salto de página antes de cada bibliografía
'
	Dim scCurrent As Section, rgFindRange As Range

	For Each scCurrent In dcArg.Sections
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
End Sub

Sub BibliografiaExportar(dcArg As Document)
' Exporta la bibliografía en archivos separados y la borra de dcArg
'
	Dim dcBibliografia As Document, scCurrent As Section
	Dim rgBiblio As Range, rgTitulo As Range, stNombre As String

	For Each scCurrent In dcArg.Sections
		Set rgBiblio = scCurrent.Range
		With rgBiblio.Find
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
				If Not .Found Then .Execute FindText:="referencias"
			End If
		End With
		
		If rgBiblio.Find.Found Then
			' Asigna el número de tema
			Set rgTitulo = scCurrent.Range.Paragraphs.First.Range
			With rgTitulo.Find
				.ClearFormatting
				.Replacement.ClearFormatting
				.Forward = True
				.Wrap = wdFindStop
				.Format = True
				.MatchCase = False
				.MatchWholeWord = False
				.MatchWildcards = True
				.MatchSoundsLike = False
				.MatchAllWordForms = False
				.Style = wdStyleHeading1
				.Execute FindText:="TEMA [0-9][0-9]"
				If Not .Found Then .Execute FindText:="TEMA [0-9]"
				If .Found Then
					stNombre = dcArg.Path & Application.PathSeparator _
						& "BIBLIOGRAFÍA " & rgTitulo.Text & ".pdf"
				Else
					Beep
					stNombre = InputBox("Número de tema no encontrado, completar", "Bibliografía", "TEMA " & scCurrent.Index)
					stNombre = dcArg.Path & Application.PathSeparator _
						& "BIBLIOGRAFÍA " & stNombre & ".pdf"
				End If
			End With

			' Exporta el pdf
			Set rgBiblio = RaMacros.RangeGetCompleteOutlineLevel(rgBiblio.Paragraphs(1))

			dcArg.ExportAsFixedFormat2 _
				OutputFileName:= stNombre, _
				ExportFormat:= wdExportFormatPDF, _
				OpenAfterExport:= False, _
				OptimizeFor:= wdExportOptimizeForPrint, _
				Range:= wdExportFromTo, _
				From:= rgBiblio.Characters(1).Information(wdActiveEndPageNumber), _
				To:= rgBiblio.Information(wdActiveEndPageNumber), _
				Item:= wdExportDocumentWithMarkup, _
				IncludeDocProps:= True, _
				CreateBookmarks:= wdExportCreateWordBookmarks, _
				DocStructureTags:= True, _
				BitmapMissingFonts:= False, _
				UseISO19005_1:= False, _
				OptimizeForImageQuality:= True

			' Borra la bibliografía de dcStory
			rgBiblio.Delete
		End If
	Next scCurrent
End Sub






Sub ConversionAutomaticaLibro(dcArg As Document)
' Convierte automáticamente los párrafos a los estilos de la plantilla
'
	Dim ishCurrent As InlineShape

	RaMacros.CleanBasic Nothing, True, True, dcArg

	With dcArg.Content.Find
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
	End With
	
	For Each ishCurrent In dcArg.InlineShapes
		ishCurrent.Range.Style = wdStyleNormal
	Next ishCurrent
End Sub






Sub ColoresCorrectos(dcArg As Document)
' Quita el subrayado y los colores fuera de plantilla del texto
'
	Dim rgStory As Range

	For Each rgStory In dcArg.StoryRanges
		If rgStory.StoryType > 5 Then Exit For
		Do While Not rgStory Is Nothing
			rgStory.Shading.Texture  = wdTextureNone
			rgStory.Shading.ForegroundPatternColor  = wdColorAutomatic
			rgStory.Shading.BackgroundPatternColor = wdColorAutomatic
			Set rgStory = rgStory.NextStoryRange
		Loop
	Next rgStory
End Sub






Sub ListasParaStory(dcArg As Document)
' Convierte las listas de letras o números romanos a listas de números y les añade 
' una marca para poder cambiarlas externamente y de forma automatizada en el Story
'
	With dcArg.Styles("iniseg-list_mixta").ListTemplate
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






Sub EstilosEsconder(dcArg As Document)
' Esconde todos los estilos de la galería de estilos, para que no se acumulen
	Dim stCurrent As Style
	For Each stCurrent In dcArg.Styles
		On Error Resume Next
		stCurrent.QuickStyle = False
		On Error GoTo 0
	Next stCurrent
End Sub