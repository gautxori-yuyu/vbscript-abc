Option Explicit

Class cOp_ValsEcon
	Dim strfold, oDicValoracs
	Dim strFolderOfferNumber
	Private strOferGasFolder
	Private m_ExcelApp
	
	Private Sub Class_Initialize ()
		strOferGasFolder = "C:\Program Files (x86)\Ofertas_Gas\Excel"
		Set oDicValoracs = CreateObject("scripting.dictionary")
	End Sub
	Private Sub Class_Terminate ()
	End Sub

	Function Init(strFold, ExcelApp)
		Set Init = Me
		Set m_ExcelApp = ExcelApp
		Me.strfold = strFold
	End Function
	
	Function identificaValsEconomicas ()
		If Not fso.FolderExists (strFold) Then Exit Function
		
		regex.Pattern = "^" & strQuoteNrPattern
		If Not regex.Test (fso.GetBaseName(fso.GetParentFolderName(strFold))) Then
			MsgBox ("NO SE HA NOMBRADO CORRECTAMENTE LA CARPETA DE 'Oportunidad', DEBE COMENZAR CON EL 'NUMERO DE Oferta / OPORTUNIDAD'")
			Exit Function
		Else
			strFolderOfferNumber = regex.Execute (fso.GetBaseName(fso.GetParentFolderName(strFold))).item(0).Value
		End If

		Dim fich,strNumOferta
		regex.Pattern = "^" & strQuoteNrRevPattern & ".*\.xls[xm]$"
		For Each fich In fso.GetFolder (strFold).Files
			If regex.Test (fich.Name)  And Left (fich.Name,2) <> "~$" Then
				' ESTO NO ES CONDICION SUFICIENTE PARA QUE EL FICHERO DE EXCEL SEA DE OFERGAS... Hay que ABRIR EL EXCEL, y comprobar formato de Hojas.
				strNumOferta = regex.Execute (fich.Name).item(0).submatches(0)
				If Not oDicValoracs.Exists (strNumOferta) Then oDicValoracs.Add strNumOferta, CreateObject("scripting.dictionary")
				oDicValoracs(strNumOferta).Add fich,Empty ' como item se va a almacenar EL CALCULO GAS_VBNET ASOCIADO
			End if
		Next
		
		' en este paso SE REGISTRAN LOS LIBROS en carpeta Ofergas, LUEGO SE RENOMBRAN, junto con los de la carp de valoración económica
		Dim bMoveExcelsMsg
		For Each fich In fso.GetFolder (strOferGasFolder).Files
			regex.Pattern =  "^" & strQuoteNrPattern
			If Not regex.Test (fich.Name) Then
				' fichero no valido
			ElseIf InStr (fich.Name,strFolderOfferNumber) = 0 And strFolderOfferNumber <> "" Then
				' fichero no valido
				MsgBox "Hay un fichero de excel en la carpeta de OferGas, cuyo nombre:" & fich.Name & vbCr _
						& "NO coincide con el número de oferta de esta oportunidad:" & strFolderOfferNumber & vbCr _
						& "DEBES ajustar el nombre de la oferta en OferGas a uno con raíz " & strFolderOfferNumber & _
						" para que se tenga en cuenta en esta oportunidad", vbExclamation
				' ES PARA ASEGURARME DE QUE SOLO PROCESE LOS FICHEROS QUE CASEN CON EL NUMERO DE 'OPORTUNIDAD / OFERTA', si es que está definida
			Else
				Select Case LCase(fso.GetExtensionName(fich.Name))
					Case "xlsx","xlsm"
						If Not bMoveExcelsMsg Then
							If MsgBox ("¿Mover los ficheros de Excel en la carpeta OferGas que correspondan a los cálculos de la oferta, a la carpeta " & fso.GetFileName (strFold) & "?", 4) <> 6 Then
								Exit For
							End If
							bMoveExcelsMsg = true
						End If
						If bMoveExcelsMsg Then
							' en este paso SE REGISTRAN LOS LIBROS en carpeta Ofergas, LUEGO SE RENOMBRAN, junto con los de la carp de valoración económica
							strNumOferta = regex.Execute (fich.Name).item(0).value
							If Not oDicValoracs.Exists (strNumOferta) Then oDicValoracs.Add strNumOferta, CreateObject("scripting.dictionary")
							oDicValoracs(strNumOferta).Add fich,Empty ' como item se va a almacenar EL CALCULO GAS_VBNET ASOCIADO
						End If
					Case Else
				End Select
			End If
		Next
	End Function
	
	Function procesaCarpValEconomica (oDicCalcs)
		Dim fich,strNumOferta,strRevOferta
		' SE PUEDE DAR POR HECHO QUE LA CARPETA EXISTE, SE TRABAJA CON DICCIONARIOS DE FICHEROS UBICADOS EN ELLA...
		'If Not fso.FolderExists (strFold) Then Exit Function
	
		' los XLSX DE VALORACION ECONOMICA, DEBERIA CONVERTIRLOS a XLSM Y BORRARLOS. Son DUPLICADOS, ... y cuesta mantener la duplicidad
		
		Dim bNotProcessed,strOfergas_ABCGasCalcNr,bCalcExists,bOpcExists
		Dim strLog
		' renombra ficheros Excel, TENIENDO EN CUENTA EL CONTENIDO; y añade la plantilla de cálculo de márgenes
		For Each strNumOferta In oDicValoracs
		For Each fich In oDicValoracs(strNumOferta).Keys
			bNotProcessed = m_ExcelApp.IsFileOpenExternally(fich.Path)
			If bNotProcessed Then
				bNotProcessed = (MsgBox ("DEBES CERRAR EL FICHERO '" & fich.Name & "' para poder procesarlo. ¿OMITIR SU PROCESADO?" & vbCr & _
						"(debería cambiar el CODIGO, para que si está abierto, haga los cambios sobre el fichero abierto..." & _
						" pero no es facil controlar en qué instancia de Excel estuviera abierto!!)", 4) = 6)
				If Not bNotProcessed Then bNotProcessed = m_ExcelApp.IsFileOpenExternally(fich.Path)
			End if

			If Not bNotProcessed Then
				' **** REVISAR BIEN COMO PONGO NOMBRES DE LOS FICHEROS:
				' - para que la RELACION ENTRE CALCULOS Y VALORACIONES, Y VALORACIONES-OFERTAS, SEA FACIL!!!
				' - y que contengan FRASE QUE SEA DECRIPTIVA de los que representa el calculo (operativas, diseño, valv escape,...) Y LA VALORACION (modelo del compresor), ...
				' - LOS NOMBRES DE FICHS DE OFERTA NO DEBEN TENER SIMBOLOS , COMAS, NI &, ... para poder subirlos a creatio
				' **** IMPORTANTE RELACIONAR BIEN CALCULOS TECNICOS CON VALORACIONES, Y CON LOS CODIGOS DE OFERTAS CORRESPONDIENTES y con modelos...
				Call renombraFichValEconomica (fich, strFold, oDicValoracs(strNumOferta), strOfergas_ABCGasCalcNr)
				
				regex.Pattern = "([A-Z]{3}\d{5})(?:_(\d{2}))?"
				If regex.Test (strOfergas_ABCGasCalcNr) Then
					bCalcExists = oDicCalcs.Exists(regex.Execute(strOfergas_ABCGasCalcNr).item(0).submatches(0))
					
					If Not bCalcExists Then
						If bLog then
							If Not bCalcExists Then strLog = "NO "
							MsgLog vbTab & "El calculo referido en la hoja, " & strOfergas_ABCGasCalcNr & ", " & strLog & " existe en calculos técnicos"
						Else
							MsgBox ("OJO!!, no existe el fichero de resultados " & strOfergas_ABCGasCalcNr & " para el cálculo de gas_vbnet asociado a la oferta '" & fich.Path & "' --> NO se puede comparar con el cálculo técnico")
						End If
					Else
						' Comprueba que exista la OPCION de calculo... aunque en realidad ME INTERESA EL NUMERO DE CALCULO: cada calculo debería corresponder con UN COMPRESOR A OFERTAR
						bOpcExists = oDicCalcs(regex.Execute(strOfergas_ABCGasCalcNr).item(0).submatches(0)).Exists(UCase(strOfergas_ABCGasCalcNr))
						If Not bOpcExists Then
							If bLog then
								If Not bOpcExists Then strLog = "NO "
								MsgLog vbTab & "La OPCION de calculo referida en la hoja, " & strOfergas_ABCGasCalcNr & ", " & strLog & " existe en calculos técnicos"
							Else
								MsgBox ("OJO!!, el cálculo definido no corresponde a una opción de cálculo de Ofergas, o no existe el fichero de resultados " & strOfergas_ABCGasCalcNr & " para el cálculo de gas_vbnet asociado a la oferta '" & fich.Path & "' --> NO se puede comparar con el cálculo técnico")
							End If
						End If
						
						' *** VINCULA VALORACIONES ECONOMICAS CON CALCULOS DE GAS_VBNET
						'Stop ' TENGO UN ERROR RARO, CON EL TIPO DE strOfergas_ABCGasCalcNr
						'Set v = oDicValoracs(strNumOferta)
						'v = strOfergas_ABCGasCalcNr
						If strOfergas_ABCGasCalcNr <> "" Then oDicValoracs(strNumOferta).Item(fich) = strOfergas_ABCGasCalcNr
					End If
				End if
			
				' por defecto, SOLO AÑADO LA PLANTILLA A FICHEROS XLSX?? (también puedo añadirla a los XLSM... pero en ppio asumo que ya la tuvieran)	
				If (LCase(fso.GetExtensionName(fich)) = "xlsx") Then _
						Call addSheetMargen_FichValEconomica (fich, strNumOferta, oDicValoracs)
			End If
					
			MsgIE ("queda por añadir la VALIDACION DE LA VALORACION ECONOMICA, para asegurarse por ejemplo de que un compresor NACE tiene caldereria INOX, bloque SAS, panel de N2, vastago WC, (empaquetadura AISI-316,...)" & vbcrlf & _
					"(tal vez sea conveniente aqui crear una clase similar a cABCGAS...)")
		Next
		Next
	End Function
	
	Function renombraFichValEconomica (fich, strFold, oDicValoracs_NumOferta, ByRef strOfergas_ABCGasCalcNr)
		Dim strExtValEcon, strDestFN, excelFile, oXlSheet, strFecha, strOfergas_Cliente, strOfergas_Comment, bNoModel
		MsgLog "Procesando: " & fich.Path
		strExtValEcon = fso.GetExtensionName(fich)
		strDestFN = ""
		' extrae la info para el nombre de fichero
		Stop ' ****** POSIBLEMENTE ME INTERESE CAMBIAR ESTO PARA USAR cOferGas, y seguir el mismo patron q cABCGas..
		' ****************
		Set excelFile = m_ExcelApp.OpenFile(fich.Path, True, True) ' Solo lectura
		
		If Not excelFile Is Nothing Then
			If excelFile.HasWorksheet("Hoja2") Then
				Set oXlSheet = excelFile.GetWorksheet("Hoja2")

				strFecha = oXlSheet.range ("L1").Value
				If strFecha <> "" Then
					strFecha = Split (strFecha,"/")(2) & "-" & Split (strFecha,"/")(1) & "-" & Split (strFecha,"/")(0)
				End If
				strOfergas_Cliente = oXlSheet.range ("B1").Value
				strOfergas_ABCGasCalcNr = oXlSheet.range ("F2").Value
				strOfergas_Comment = oXlSheet.range ("A3").value
				strOfergas_Comment = Left (strOfergas_Comment,Instr (strOfergas_Comment,vbcr))
				strOfergas_Comment = Left (strOfergas_Comment,Instr (strOfergas_Comment,vbLf))
				bNoModel = oXlSheet.range ("K22").value = 0

				m_ExcelApp.CloseFile excelFile.FilePath, Empty
			Else
				m_ExcelApp.CloseFile excelFile.FilePath, Empty
				Exit Function
			End If
		Else
			Exit Function
		End If
		
		Dim strAdd
		'strDestFN = Replace(fich.Path,"." & strExtValEcon,"_" & strOfergas_Cliente & "_" & strOfergas_Comment & "_" & strOfergas_ABCGasCalcNr & "." & strExtValEcon)
		If False And strFecha <> "" And InStr (fich.Name, "_" & strFecha) = 0 Then strAdd = "_" & strFecha
		If strOfergas_Cliente <> "" And InStr (fich.Name, "_" & strOfergas_Cliente) = 0 Then strAdd = "_" & strOfergas_Cliente
		If strOfergas_Comment <> "" And InStr (fich.Name, "_" & strOfergas_Comment) = 0 Then strAdd = strAdd & "_" & strOfergas_Comment
		If strOfergas_ABCGasCalcNr <> "" And InStr (fich.Name, "_" & strOfergas_ABCGasCalcNr) = 0 Then strAdd = strAdd & "_" & strOfergas_ABCGasCalcNr
		Stop ' ADEMAS DE EL NUMERO DE CALCULO (O EN SU LUGAR) HAY QUE METER EL MODELO DEL COMPRESOR, para mas claridad
		strDestFN = Replace(fich.Path,"." & strExtValEcon,strAdd & "." & strExtValEcon)
		If Len (strDestFN) > 256 Then
			stop
			strDestFN = Left(Replace(strDestFN,"." & fso.GetExtensionName(strDestFN),""),250) & "." & fso.GetExtensionName(strDestFN)
		End If
	
		regex.Pattern = "_old\(\d+\)"
		if regex.Test(fich.Name) And strDestFN = fich.Path Then
			MsgLog vbTab & "El fichero NO debe renombrarse, es una versión OLD con el nombre correcto"
			 ' no hay nada que renombrar, ni que añadir (doy por hecho)
			 Exit Function
		Elseif strDestFN = "" Or strDestFN = fich.Path Then
			MsgLog vbTab & "El fichero NO debe renombrarse, mantendría el mismo nombre"
			' no hay nada que renombrar
			 Exit Function
		Else
			If InStr (strDestFN,"C:\Program Files (x86)\Ofertas_Gas\Excel") > 0 Then
				' si el fichero está en la carpeta de Excels de Ofergas, la ruta de destino se pone la de "valoracion económica":
				' este fichero se renombra seguro, llevándolo a la carpeta de VALORACIONES
				strDestFN = strFold & "\" & fso.GetFileName (strDestFN)
			ElseIf bNoModel Then
				If MsgBox ("OJO, en la oferta '" & fich.Name & "' NO se ha añadido valoración para el bloque CHASIS - MODEL!! (posiblemente sea un duplicado)." & vbcrlf & _
						"¿Borrarla (habría que reparar la oferta manipulando la DB Ofergas, y volver a generar el Excel)?",4) _
						= 6 Then
					fich.Delete true
					oDicValoracs_NumOferta.Remove(fich)
				End If
			Else
				regex.Pattern = "_" & strFecha & ".*" & strOfergas_ABCGasCalcNr
				If regex.Test(fich.Name) Then If MsgBox ("El nombre del fichero a renombrar YA tiene un patron de haber sido RENOMBRADO. ¿Seguro que quieres cambiar su nombre?",4) _
						<> 6 Then Exit Function
			End If
		End If
	
		Dim strLog
		MsgLog vbTab & "Nombre de destino: " & strDestFN
		If bNoModel Then strLog = "NO "
		MsgLog vbTab & strLog & "existe la sección de MODELO"
		
		Dim strOld, c
		' se renombra el fichero de oferta, a partir de la información en él.
		c = 1
		strOld = strDestFN
		Do While fso.FileExists (strOld)
			strOld = Left (Replace(strDestFN,".xlsx",""),243) & ".xlsx"
			strOld = Replace (strOld,".xlsx","_old(" & c & ").xlsx")
			strOld = Replace(strOld,"_old(1)","")
			c = c + 1
		Loop
		
		Dim bRename
		bRename = InStr (fich.Path,"C:\Program Files (x86)\Ofertas_Gas\Excel") > 0
		If Not bRename Then
		Select Case MsgBox ("¿Mover '" & "\3.VALORACION ECONOMICA\" & fich.Name & "'" & vbCrLf & "a '" & _
				"\3.VALORACION ECONOMICA\" & fso.GetFileName (strDestFN) & "'?" & vbCr & "(añadido al nombre:" & strAdd & ") (CANCELAR detiene el script)",3+48)
			Case 6 : bRename = True
			Case 2
				' Shutdown Excel si cancela
				m_ExcelApp.Shutdown
	 			WScript.Quit
		End Select
		End if
		If bRename Then
			' YA NO HACE FALTA corregir KEYS en oDicValoracs CADA VEZ QUE SE RENOMBRA UN FICHERO, las propias keys cambian...
			If fso.FileExists (strDestFN) Then fso.MoveFile strDestFN,strOld
			If fso.FileExists (Replace(strDestFN,".xlsx",".xlsm")) And Replace(strDestFN,".xlsx",".xlsm") <> Replace(strOld,".xlsx",".xlsm") Then
				Stop ' NO TENGO CLARO QUE SEA MEJOR RENOMBRAR ESTE AQUI!!!
				' - creo que ES MEJOR hacerlo, PARA QUE SE MANTENGAN EMPAREJADOS los nombres de ficheros ... 
				'		sin peligro de que "PIERDA REFERENCIAS": EL "_old" NO me tiene que proporcionar una referencia
				' - posiblemente se renombraria ¿CORRECTAMENTE? cuando toque procesarlo?
				fso.MoveFile Replace(strDestFN,".xlsx",".xlsm"),Replace(strOld,".xlsx",".xlsm")
				MsgLog vbTab & "El nombre de fich de destino YA EXISTE; se ha renombrado a la versión OLD siguiente, " & Replace(strOld,".xlsx",".xlsm")
			End If
	
			fich.move strDestFN
			renombraFichValEconomica = strDestFN
			MsgLog "<b>" & "Fichero movido a destino" & "</b>"
		End if
	End Function
	
	Function addSheetMargen_FichValEconomica (fich, strNumOferta, oDicValoracs)
		Dim oXlWorkBook, oXldestWorkBook
		' se añade la PLANTILLA DE CALCULO DE INGENIERÍA Y MÁRGENES, entendiendo que los .xlsM YA LA TIENEN, y los .xlsx NO la tienen!!!??
		' SE PODRIA COMPROBAR, HOJA POR HOJA, QUE NINGUNA DE ELLAS CUMPLE UN PATRON DE ESA HOJA DE MARGENES...
		' identifica el fichero fuente de la "plantilla de cálculo de ingeniería y márgenes",
		' y AÑADE AL EXCEL de Ofergas la plantilla, guardando el Excel como un xlsm
		If Not m_ExcelApp.HasWorksheet(fich.Path, "Hoja2") Then
			Stop ' EL LIBRO "DE DESTINO" DEBERIA TENER "Hoja2", sino es muy probable que ya esté modificado, e incluya la plantilla....
			' en ppio si no tiene la hoja, NI SE LE METE LA PLANTILLA, ni tendría nada de info que pudiera llevar a un XLSM
			Exit Function	
		End If
		If m_ExcelApp.HasWorksheet(fich.Path, "Margenes Ingenieria etc") Then
			' el libro de destino YA TIENE LA HOJA DE MARGENES, (Y ES UN XLSX... alguien lo ha manipulado y guardado...)
			' ni siquiera se guarda como xlsm
			Exit Function	
		End If
		
		' proceso el caso de que ya exista el XLSM.. (en general NO DEBERIA)
		If fso.FileExists (Replace(fich.Path,".xlsx",".xlsm")) Then
			If m_ExcelApp.HasWorksheet (Replace(fich.Path,".xlsx",".xlsm"), "Margenes Ingenieria etc") Then
				' ya existe un fichero XLSM, y compruebo que YA TIENE LA PLANTILLA DE MARGENES!!!
				' SI EL XLSX ES MAS NUEVO que el XLSM, Y TIENE LA "Hoja2", DOY LA OPCIÓN DE AÑADIR LA HOJA2 al "XLSM"!!!
				Select Case True
					Case fso.GetFile(Replace(fich.Path,".xlsx",".xlsm")).DateLastModified >= fich.DateLastModified
						'Stop
					Case MsgBox ("YA EXISTE un fichero de macros, XLSM, para el fichero de oferta en proceso: ¿quieres añadir a aquel la Hoja2 del XLSX? (*** ES MUY PROBABLE QUE ESA Hoja2 TENGA DATOS ACTUALIZADOS; y podría ser útil para tener todas las Hojas de una misma oferta en un mismo fichero Excel!!)", 4) = 6
						Stop

	                    Dim srcExcelFile, destExcelFile
	                    Set srcExcelFile = m_ExcelApp.OpenFile(fich.Path, True, False) ' Solo lectura, TIENEN QUE SER VISTOS!!, SI NO ERROR AL COPIAR LAS HOJAS.
	                    Set destExcelFile = m_ExcelApp.OpenFile(xlsmPath, False, False) ' Lectura/escritura, TIENEN QUE SER VISTOS!!, SI NO ERROR AL COPIAR LAS HOJAS.

	                    If Not srcExcelFile Is Nothing And Not destExcelFile Is Nothing Then
							Set oXlWorkBook = srcExcelFile.WorkBook
							Set oXldestWorkBook = destExcelFile.WorkBook
							' YA EXISTEN LAS HOJAS EN CADA FICHERO, lo garantizan los pasos anteriores
							oXlWorkBook.WorkSheets("Hoja2").Copy oXldestWorkBook.WorkSheets("Margenes Ingenieria etc")
	                        ' Guardar y cerrar
	                        m_ExcelApp.CloseFile xlsmPath, True ' Guardar cambios
	                        m_ExcelApp.CloseFile fich.Path, False ' No guardar cambios
	                    End If
				End Select
			End If
			Exit Function
		End If
		
		Dim fichsrc,fichsrcPref
		' identifico el fichero desde el que me interesa sacar la plantilla de margenes: EL MAS NUEVO, QUE TENGA UN PATRON DE NOMBRE SIMILAR...
		For Each fichsrc In fich.ParentFolder.Files
			regex.Pattern = "[_ \-]*(?:rev[\._ \-]*\d+\b|old\(\d+\))"
	'		If bLog Then
	'			MsgLog & "Determinando si " & fichsrc.name & " sirve de fuente de plantilla margenes para " & fich.Name & ":"
	'			MsgLog vbTab & vbTab & "¿ " & LCase(regex.Replace (fichsrc.Name,"")) & " == " & LCase (regex.Replace (fich.Name,"")) & " ?"
	'		End If
			Select Case True
				Case fichsrc.Name = fich.Name
				Case LCase(fso.GetExtensionName(fichsrc.Name)) <> "xlsm" ' ESTOS NO VALDRÍAN, no tienen plantilla (válida)
				Case IsEmpty (fichsrcPref) And fichsrc.Name = "plantilla de referencia para calculo ingenieria y margenes.xlsm"
					Set fichsrcPref = fichsrc
					MsgLog vbTab & "Nuevo fichero de base para la plantilla de márgenes: " & fichsrc.name
				Case LCase(regex.Replace (fichsrc.Name,"")) = LCase (regex.Replace (fich.Name,"")) ' fichero OLD o rev
					Select Case True
						Case IsEmpty (fichsrcPref), fichsrcPref.Name = "plantilla de referencia para calculo ingenieria y margenes.xlsm", _
								InStr (fichsrcPref.Name,"_old(") = 0, fichsrcPref.DateLastModified < fichsrc.DateLastModified
							' me aseguro de que el fichero que se elija, TENGA LA PLANTILLA!!
	'						Set oXlWorkBook = objExcel.Application.Workbooks.open(fichsrc.Path, 0, True, 2, "", "",False , 2, "", true, False, ,False, True,0)
	'						oXlWorkBook.Activate
	'						If Not objExcel.Application.Evaluate("ISREF('" & "Margenes Ingenieria etc" & "'!A1)") Then
	'							Set fichsrcPref = fichsrc
	'							oXlWorkBook.Close False
	'						End If
							If m_ExcelApp.HasWorksheet(fichsrc.Path, "Margenes Ingenieria etc") Then
								Set fichsrcPref = fichsrc
								MsgLog vbTab & "Nuevo fichero de base para la plantilla de márgenes: " & fichsrc.name
							End if
							' SE PODRIA HACER LA VERIFICACIÓN DE OTRA MANERA: BUSCAR EN TODAS LAS HOJAS UNA CELDA CON UN VALOR / FORMULA DADOS...
					End Select
				'Case IsEmpty (fichsrcPref) 'para evitar error en el siguiente Case
				Case IsEmpty (fichsrcPref),InStr (fichsrc.Name, strNumOferta) > 0 And InStr (fichsrcPref.Name,"_old(") = 0
					' cualquier otro fichero, DE LA MISMA OFERTA (o incluso de otra) que contuviera la plantilla... podría ser mejor opción que el "estandar"
					regex.Pattern =  "^" & strQuoteNrPattern
					Select Case True
						Case regex.Execute(fichsrc.Name).Item(0).Value <> regex.Execute(fich.Name).Item(0).Value
							' PARA QUE EL FICHERO SEA ACEPTABLE, DEBE COINCIDIR EL NUMERO COMPLETO DE LA OFERTA (la comparación con strNumOferta NO TIENE EN CUENTA LAS 'REVISIONES' \d...)
						Case IsEmpty (fichsrcPref),fichsrcPref.Name = "plantilla de referencia para calculo ingenieria y margenes.xlsm", _
								fichsrcPref.DateLastModified < fichsrc.DateLastModified
							If m_ExcelApp.HasWorksheet(fichsrc.Path, "Margenes Ingenieria etc") Then
								Set fichsrcPref = fichsrc
								MsgLog vbTab & "Nuevo fichero de base para la plantilla de márgenes: " & fichsrc.name
							End if
							' SE PODRIA HACER LA VERIFICACIÓN DE OTRA MANERA: BUSCAR EN TODAS LAS HOJAS UNA CELDA CON UN VALOR / FORMULA DADOS...
					End Select
				Case fichsrc.Name = "plantilla de referencia para calculo ingenieria y margenes.xlsm"
				Case Else
					' simplemente chequear los ficheros, pero SI NO SON DE LA MISMA OFERTA, en ppio, sería mejor usar la plantilla estándar...
					Stop
			End Select
		Next
		If IsEmpty (fichsrcPref) Then Exit Function
	    Dim templateExcelFile, targetExcelFile
	    Set templateExcelFile = m_ExcelApp.OpenFile(fichsrcPref.Path, True, False) ' Solo lectura, TIENEN QUE SER VISTOS!!, SI NO ERROR AL COPIAR LAS HOJAS.
	    Set targetExcelFile = m_ExcelApp.OpenFile(fich.Path, False, False) ' Lectura/escritura, TIENEN QUE SER VISTOS!!
	    If templateExcelFile Is Nothing Or targetExcelFile Is Nothing Then
	        If Not templateExcelFile Is Nothing Then m_ExcelApp.CloseFile fichsrcPref.Path, False
	        If Not targetExcelFile Is Nothing Then m_ExcelApp.CloseFile fich.Path, False
	        Exit Function
	    End If

		Set oXlWorkBook = templateExcelFile.WorkBook
		Set oXldestWorkBook = targetExcelFile.WorkBook
		MsgLog "Añadiendo la plantilla de calculo de MARGENES E INGENIERIA, desde " & oXlWorkBook.Name
		'oXldestWorkBook.Activate
		' YA EXISTEN LAS HOJAS EN CADA FICHERO, lo garantizan los pasos anteriores
		oXlWorkBook.WorkSheets("Margenes Ingenieria etc").Copy oXldestWorkBook.WorkSheets("Hoja2").Next
		If fichsrcPref.Name = "plantilla de referencia para calculo ingenieria y margenes.xlsm" Then
			oXldestWorkBook.WorkSheets("Margenes Ingenieria etc").UsedRange.ClearComments
			oXldestWorkBook.WorkSheets("Margenes Ingenieria etc").UsedRange.ClearNotes
			Stop
			' DE MOMENTO LAS HORAS DE INGENÍERÍA SE ESTIMAN, SEGÚN EL MODELO:
			' MIN 700, HA2 o menos; max 3000, HX6
	'				Select Case
	'					Case
	'				End Select
		End If
		Stop
	    ' Guardar como XLSM y cerrar
	    targetExcelFile.SaveAs Replace(fich.Path,".xlsx",".xlsm"), xlOpenXMLWorkbookMacroEnabled
	    m_ExcelApp.CloseFile Replace(fich.Path,".xlsx",".xlsm"), False ' Ya se guardó con SaveAs
	    m_ExcelApp.CloseFile fichsrcPref.Path, False ' No guardar cambios en plantilla
	End Function
End Class

