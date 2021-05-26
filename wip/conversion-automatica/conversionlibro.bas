Sub ConversionAutomaticaLibro(dcArgument As Document)
' Convierte automáticamente los párrafos a los estilos de la plantilla
' Instrucciones:
	' Es conveniente pasar antes:
		' La conversión de imágenes, para que no anden bailando
		' El autoformat con conversión de listas con viñeta
'
	Dim pCurrent As Paragraph
	Dim rgFind As Range
	Dim sIndent As Single
	Dim iCambioNivelLista As Integer




	Application.ScreenUpdating = False



' PASA PRIMERO LA CONVERSIÓN DE LISTAS









	For Each pCurrent In dcArgument.Content
		If pCurrent.Range.Words.Count > 5 Then
			rgFind.SetRange pCurrent.Range.Start, pCurrent.Range.Words(5).Range.End
		Else
			Set rgFind = pCurrent.Range
		End If
		RaMacros.FindAndReplaceClearParameters
		If pCurrent.Range.Tables.Count = 0 Then
			If rgFind.Find.Execute(FindText:="[tT][eE][mM][aA][0-9]{1;2}", _
									MatchWildcards:=True, Wrap:= wdFindStop) _
				Or rgFind.Find.Execute(FindText:="[tT][eE][mM][aA][ ][0-9]{1;2}", _
										MatchWildcards:=True, Wrap:= wdFindStop) _
			Then
				pCurrent.style = wdStyleHeading1
			ElseIf rgFind.Find.Execute(FindText:="[0-9].[0-9].[0-9].[0-9]", _
										MatchWildcards:=True, Wrap:= wdFindStop) _
				And pCurrent.Range.Font.Size > 11
			Then
				pCurrent.style = wdStyleHeading5
			ElseIf rgFind.Find.Execute(FindText:="[0-9].[0-9].[0-9]", _
										MatchWildcards:=True, Wrap:= wdFindStop) _
				And pCurrent.Range.Font.Size > 11
			Then
				pCurrent.style = wdStyleHeading4
			ElseIf rgFind.Find.Execute(FindText:="[0-9].[0-9]", _
										MatchWildcards:=True, Wrap:= wdFindStop) _
				And pCurrent.Range.Font.Size > 11
			Then
				pCurrent.style = wdStyleHeading3
			ElseIf rgFind.Find.Execute(FindText:="[0-9].", _
										MatchWildcards:=True, Wrap:= wdFindStop) _
				And pCurrent.Range.Font.Size > 11
			Then
				pCurrent.style = wdStyleHeading2
			Else
				Select Case pCurrent.Range.Font.Size
					Case Is >= 15
						pCurrent.style = wdStyleHeading2
					Case 13, 14
						If Font.Italic = False Then
							pCurrent.style = wdStyleHeading3
						Else
							pCurrent.style = wdStyleHeading4
						End If
					Case 12, 13
						pCurrent.style = wdStyleHeading4
					Case Is < 11
						' Comprobar si lo anterior/siguiente es una imagen/tabla (y colocar el pie adecuadamente)
						' si no es una tabla: texto normal
					Case Else
						' COMPROBAR LISTA
							' Búsqueda de bullets fuera de lista: [–\-—•⁎⁕▪▸◂◃▷◼◻●◌◇◆]
							' stPatron = "^(?:[a-zA-Z0-9]{1,2}[\.\)\-ºª]+(?:[a-zA-Z0-9]{1,2}[\.\)\-ºª]?)*|[–\-—•⁎⁕▪▸◂◃▷◼◻●◌◇◆])[\s]*"
						If rgFind.Find.Execute(FindText:="PATRÓN DE VIÑETAS") Then
							sIndent = pCurrent.LeftIndent
							pCurrent.Style
						ElseIf rgFind.Find.Execute(FindText:="[a-zA-Z0-9]{1;2}[\.\)\-ºª]", _
													MatchWildcards:=True, Wrap:= wdFindStop) _
							Or rgFind.Find.Execute(FindText:="[a-zA-Z0-9]{1;2}[\.\)\-ºª]", _
													MatchWildcards:=True, Wrap:= wdFindStop) _
							Or rgFind.Find.Execute(FindText:="[a-zA-Z0-9]{1;2}[\.\)\-ºª]", _
													MatchWildcards:=True, Wrap:= wdFindStop) _
						Then
							If anterior.ListLevelNumber <> 0 _
								And rgFind.Find.Execute(FindText:="PATRÓN DE VIÑETAS O NUMERACIÓN") _
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
		End If
	Next pCurrent





	Application.ScreenUpdating = True



End Sub

