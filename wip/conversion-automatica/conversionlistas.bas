Sub ConversionListas(dcArgumentDocument As Document)
' Convierte las listas a los estilos de la plantilla
'
	Dim pCurrent As Paragraph
	Dim rgFind As Range
	Dim sIndent As Single
	Dim iCambioNivelLista As Integer
	Dim lCurrent As List
	Dim stPreviousStyle As String, stCurrentStyle As String

	stPreviousStyle = vbNullString


' Se itera por todas las listas, comprobando que los niveles de lista son correlativos (no se pasa del 1 al 3, por ejemplo) y corrigiéndolas, de ser así
' Si no lo tiene, se otorga nivel de lista a cada párrafo comparándolo con la sangría del anterior. De esta forma los párrafos quedan ya metidos en una lista. Después se itera por cada lista y simplemente se busca qué tipo de lista es, si bullet o numerada, y se les otorga el ESTILO de lista necesario.




	' Autoformateo para bulleted
	' Iterar por todas las listas existentes y aplicarles la plantilla adecuada
	' Iterar por cada párrafo:
		' If es de tipo lista:
			' Guardar el estilo en stCurrentStyle
			' If stPreviousStyle <> vbNullString
				' If stCurrentStyle < stPreviousStyle - 1
					' pCurrentStyle = stPreviousStyle - 1
			' Guardar stCurrentStyle en stPreviousStyle
		' Else stPreviousStyle = vbNullString
		' Comprobar si contiene un patrón de lista y, de ser así, guardar el estilo adecuado en otra variable
		' cotejar el leftindent actual y el del párrafo anterior y:
			' si coincide el tipo de lista y la sangría, la lista continúa
			' si coincide el indent y hay patrón de lista, pero no coincide con el anterior: siguiente nivel de lista (respecto al anterior párrafo)
			' si el indent no es 0, coincide con el anterior, pero no hay patrón de lista: listcontinue al nivel equivalente del anterior párrafo
			' Si el anterior párrafo no es de lista, el actual es mayor y contiene patrón de lista
	' ITERAR POR TODAS LAS LISTAS DEL DOCUMENT.CONTENT Y CONVERTIRLAS AL TIPO DE LISTA ADECUADO
	' Arreglar comienzos de lista







	' BUSCAR TODAS LAS VECES QUE SE HAYA PASADO EL AUTOFORMAT POR ALTO UN TIPO DE LISTA 


	For Each lCurrent In dcArgumentDocument.Lists

	Next lCurrent


	For Each pCurrent In dcArgument.Paragraphs
		Select Case pCurrent.Range.ListFormat.ListType
			Case 0
				If pCurrent.range.find.execute findtext=: "patrón lista" Then
					If sLeftIndent = -501 Or pCurrent.Previous.Range.
					sLeftIndent = pCurrent.LeftIndent
				End If
			Case 4
				If pCurrent.style = pCurrent.Parent.Styles(wdStyleList) _
					Or pCurrent.style = pCurrent.Parent.Styles(wdStyleList2) _
					Or pCurrent.style = pCurrent.Parent.Styles(wdStyleList3) _
					Or pCurrent.style = pCurrent.Parent.Styles(wdStyleListNumber) _
					Or pCurrent.style = pCurrent.Parent.Styles(wdStyleListNumber2) _
					Or pCurrent.style = pCurrent.Parent.Styles(wdStyleListNumber3) _
					Or pCurrent.style = pCurrent.Parent.Styles(wdStyleListBullet) _
					Or pCurrent.style = pCurrent.Parent.Styles(wdStyleListBullet2) _
					Or pCurrent.style = pCurrent.Parent.Styles(wdStyleListBullet3) _
					Or pCurrent.style = pCurrent.Parent.Styles(wdStyleListBullet4) _
					Or pCurrent.style = pCurrent.Parent.Styles(wdStyleListBullet5) _
					Or pCurrent.style = pCurrent.Parent.Styles(wdStyleListContinue) _
					Or pCurrent.style = pCurrent.Parent.Styles(wdStyleListContinue2) _
					Or pCurrent.style = pCurrent.Parent.Styles(wdStyleListContinue3) _
					Or pCurrent.style = pCurrent.Parent.Styles(wdStyleListContinue4) _
					Or pCurrent.style = pCurrent.Parent.Styles(wdStyleListContinue5) _
				Then
					pCurrent.Reset
				Else
					Select Case pC
					pCurrent.REBAJARHASTANIVEL3
		End If
	Next pCurrent




	For Each pCurrent In dcArgumentDocument.Content ' no hace falta sobar las notas u otros contenidos, yo creo
		If pCurrent.Range.Words.Count >= 5 Then
			rgFind.SetRange pCurrent.Range.Start, pCurrent.Range.Words(5).Range.End
		Else
			Set rgFind = pCurrent.Range
		End If
		If pCurrent.isTable Or pCurrent.isPicture Or pCurrent.isOle _' Quizá viendo que no sean shapes/inlineshapes
			Or pCurrent.Range.ListFormat.ListLevelNumber = 4 _
		Then
			Next pCurrent
		ElseIf Then
			' titulo 1
			' Hay que asegurar que realmente se convierte a título 1 tanto "Tema xx" como el nombre del tema
		ElseIf rgFind.Find.Execute(FindText:="[tT][eE][mM][aA][ ^t]@[0-9]{1;2}") Then
			' titulo 1
			' Hay que asegurar que realmente se convierte a título 1 tanto "Tema xx" como el nombre del tema
		Else
			Select Case pCurrent.Font.Size
				Case Is >= 15
					' título 2
				Case 13 To 14
					If Font.Italic = False Then
						' título 3
					Else
						' título 4
					End If
				Case 12 To 13
					' títulos 5
				Case Is < 11
					' Comprobar si lo anterior/siguiente es una imagen/tabla (y colocar el pie adecuadamente)
					' si no es una tabla: texto normal
				Case Else
					' COMPROBAR LISTA
						' Búsqueda de bullets fuera de lista: [–\-—•⁎⁕▪▸◂◃▷◼◻●◌◇◆]
						' stPatron = "^(?:[a-zA-Z0-9]{1,2}[\.\)\-ºª]+(?:[a-zA-Z0-9]{1,2}[\.\)\-ºª]?)*|[–\-—•⁎⁕▪▸◂◃▷◼◻●◌◇◆])[\s]*"
					If rgCurrent.Words(1).Range.Find.Execute(FindText:="PATRÓN DE VIÑETAS") Then
						sIndent = pCurrent.LeftIndent
						pCurrent.Style
					ElseIf rgCurrent.Words(1).Range.Find.Execute(FindText:="PATRÓN DE NUMERACIÓN") Then
						If anterior.ListLevelNumber > 1 _
							And rgCurrent.Words(1).Range.Find.Execute(FindText:="PATRÓN DE VIÑETAS O NUMERACIÓN") _
							Or Range.ListFormat.ListType <> 0 _
						Then
							actual.style = anterior.style 
							If sIndent < actual.indent then
								Range.ListFormat.ListIndent
							ElseIf sIndent > actual.indent then
							End If
							sIndent = actual.LeftIndent
						Else
							pCurrent.SeparateList
					Else
						' Comprobar pies de imagen: si lo anterior/siguiente es una imagen/tabla y si el párrafo es centrado
						' Comprobar quote: primer y último caracteres son comillas
					End If
			End Select
			pCurrent.Reset
		End If
	Next pCurrent
End Sub

