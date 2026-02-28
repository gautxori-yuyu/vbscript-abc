Option Explicit
Const CalcFNamePattern = "(([A-Z]{3}\d{5})_\d{2})(?:[_ \-]*rev[\._ \-]*(\d+))?.+?\.(?:txt|rtf|xlsx)$"
Const GasVBNetExportedFNamesPattern = "(?:Antipul_|ABC_(Gas_Cooler|Aircooler|Reducer|Main Motor|Instrumentation|Gas_Filter|Frequency Converter|Cooling Water Pump|Dryer|Piston_rider_ring_selection|Cooling Water Tower|Pressure_Safety_Valve|Valves_selection)\-)?([A-Z]{3}\d{5}_\d{2})(?:[_ \-]*rev[\._ \-]*(\d+))?(_calc(?:_multi)?)?.*?(\.(?:xlsx|rtf|txt))$"

Class cOp_CalcsTecn
	Dim oDicCalcs, oDicCalcsByComprModel, oDicCalcsByProcessCond
	Dim strFold, arrCalculos
	'Private CalcFNamePattern
	Private m_ExcelApp
	Private m_bAPI618
	
	Private Property Get objExcel
		Set objExcel = m_ExcelApp.Application
	End Property
	
	' habría que agruparlos por MODELOS DE COMPRESORES resultantes... Y POR CONDICIONES DE PROCESO:
	' - los calculos para un mismo modelo, deberían estar relacionados
	' - y los calculos con las mismas condiciones de proceso, aunque den distintos modelos, deberían estar relacionados
	Private Sub Class_Initialize
		Set oDicCalcs = CreateObject("scripting.dictionary") ' es un arbol, {strCalcNum,{strCalcOpc,{fich,oABCGas_XLS}}}
		oDicCalcs.CompareMode = 1
		Set oDicCalcsByComprModel = CreateObject("scripting.dictionary") ' es un arbol, {oABCGas_XLS.strModelName,{strCalcOpc,{fich,oABCGas_XLS}}}
		oDicCalcsByComprModel.CompareMode = 1
		'CalcFNamePattern = "(([A-Z]{3}\d{5})_\d{2})(?:[_ \-]*rev[\._ \-]*(\d+))?.+?\.(?:txt|rtf|xlsx)$"
		arrCalculos = Array ("C:\Aire\Calculos")
		If bEthernetConnected Then arrCalculos = Array ("C:\Aire\Calculos","W:\Aplicaciones\Compresores\Calculos")
	End Sub
	
	Function Init (strFold, ExcelApplication, bAPI618)
		Set Init = Me
		Me.strfold = strfold
		Set m_ExcelApp = ExcelApplication
		m_bAPI618 = bAPI618
	End Function
	
	Function identificaCalcTecnicos ()
		If Not fso.FolderExists (strFold) Then Exit Function
		Dim fich,strCalcNum,strCalcOpc,strCalcRev,match,latestModifiedDate
		' strCalcRev ES LA OPCION DE CALCULO QUE REVISARIA ESTE CALCULO, strCalcOpc
		For Each fich In fso.GetFolder (strFold).Files
			strCalcNum = ""
			regex.Pattern = CalcFNamePattern
			If regex.Test (fich.Name) And Left (fich.Name,2) <> "~$" Then
				strCalcNum = regex.Execute (fich.Name).item(0).submatches(1)
				strCalcOpc = regex.Execute (fich.Name).item(0).submatches(0)
				strCalcRev = regex.Execute (fich.Name).item(0).submatches(2) ' AUNQUE LO LEO AQUI, ESTAS ESTRUCTURAS NO SON CAPACES DE ALMACENAR ESTA INFO -->
				' DE MOMENTO LA INFO DE REVISIONES LA LLEVO A oABCGas_XLS, HABRÍA QUE INTERPRETARLA DESDE ALLI.
			ElseIf InStr (fich.Name, "Ayuda_") > 0 And InStr (fich.Name, ".txt") > 0 Then
				regex.Pattern = "Num cálculo : (([A-Z]{3}\d{5})_\d{2})"		
				For Each match In regex.Execute(fich.OpenAstextStream.ReadAll)
					strCalcNum = match.submatches(1)
					strCalcOpc = match.submatches(0)
					strCalcRev = Empty
				Next
				fich.Close
			Else
				MsgLog "Fichero descartado en carpeta de cálculos técnicos: " & fich.Name
			End If
			If strCalcNum <> "" Then
				If Not oDicCalcs.Exists (strCalcNum) Then
					oDicCalcs.Add strCalcNum, CreateObject("scripting.dictionary")
					oDicCalcs(strCalcNum).CompareMode = 1
				End If
				If Not oDicCalcs(strCalcNum).Exists (strCalcOpc) Then
					oDicCalcs(strCalcNum).Add strCalcOpc, CreateObject("scripting.dictionary")
					oDicCalcs(strCalcNum)(strCalcOpc).CompareMode = 1
				End If
				oDicCalcs(strCalcNum)(strCalcOpc).Add fich.Path, strCalcRev
				
				Select Case True
					Case IsEmpty(latestModifiedDate), latestModifiedDate <  fich.DateLastModified
						latestModifiedDate = fich.DateLastModified
				End Select
				
				Call MsgIE_AddRowToCalcContainer (strCalcNum,strCalcOpc,fich.Name)
				'MsgLog "Añadido al calculo " & strCalcNum & ", opción " & strCalcOpc & ", el fichero " & fich.Name
			End If
		Next
		
		Dim bMoveFile,strtmpfold,strFich,strFilter,strDestPath
		For Each strtmpfold In arrCalculos
		If fso.FolderExists (strtmpfold) Then
			' coge la raiz del primer calculo registrado, si no hay ninguno, las 4 primeras letras del ms nuevo en la carpeta
			' lista todos los q tengan esa reiz
			If oDicCalcs.Count > 0 Then strFilter = Left(oDicCalcs.Keys()(0),4)
			For Each strFich In getCalcFilesInFolder (strtmpfold,strFilter)
				' si tienen mismo num de calculo que en destino, se mueven (con confirmación, si ya existen)
				' si no, SI SON MAS NUEVOS que el mas nuevo de destino, tb se mueven.
				If strFich <> "" then
					Set fich = fso.GetFile (strFich)
					regex.Pattern = CalcFNamePattern
					If regex.Test (fich.Name) And Left (fich.Name,2) <> "~$" Then
						strCalcNum = regex.Execute (fich.Name).item(0).submatches(1)
						strCalcOpc = regex.Execute (fich.Name).item(0).submatches(0)
						Select Case True
							Case oDicCalcs.Exists(strCalcNum) And Not fso.FileExists (strFold & "\" & fich.Name)
								' cuando existe una referencia que induzca a moverlo, se mueve...
								bMoveFile = true
							Case Not fso.FileExists (strFold & "\" & fich.Name) And fich.DateLastModified > latestModifiedDate
								' cuando no, sólo si el fichero se ha creado recientemente, Y CON CONFIRMACION
								bMoveFile =(MsgBox ("Hay un fichero de cálculo más reciente en '" & strtmpfold & _
										"'. ¿Moverlo a '" & "\2.CALCULO TECNICO\" & fich.name & "'?",4) = 6)
							Case Not fso.FileExists (strFold & "\" & fich.Name)
								' para asegurarme de que en las condiciones siguientes existe el fichero en destino
							Case fich.DateLastModified > fso.GetFile (strFold & "\" & fich.Name).DateLastModified
								bMoveFile =(MsgBox ("Hay un fichero de cálculo más reciente en '" & strtmpfold & _
										"'. ¿Moverlo a '" & "\2.CALCULO TECNICO\" & fich.name & "'?",4) = 6)
							Case fich.DateLastModified < fso.GetFile (strFold & "\" & fich.Name).DateLastModified
								' el fichero en la carpeta de cálculos es más viejo --> lo borro
								fich.delete True
								bMoveFile = False
							Case Else
								Stop
								bMoveFile = False
						End Select
						If bMoveFile Then
							MsgIE ("Moviendo fichero " & fich.Path & " a la carpeta de oportunidad")
							strDestPath = strFold & "\" & fich.Name
							If m_ExcelApp.RenameWorkbook(fich.Path, strDestPath) Then
								Set fich = fso.getFile (strDestPath)
							End if
							
							If Not oDicCalcs(strCalcNum).Exists (strCalcOpc) Then
								oDicCalcs(strCalcNum).Add strCalcOpc, CreateObject("scripting.dictionary")
								oDicCalcs(strCalcNum)(strCalcOpc).CompareMode = 1
							End If
							If Not oDicCalcs(strCalcNum)(strCalcOpc).Exists (fich.Path) Then
								oDicCalcs(strCalcNum)(strCalcOpc).Add fich.Path, Empty
								Call MsgIE_AddRowToCalcContainer (strCalcNum,strCalcOpc,fich.Name)
								'MsgLog "Añadido al calculo " & strCalcNum & ", opción " & strCalcOpc & ", el fichero " & fich.Name
							End If
						End If
					End If
				End If
			Next
		End If
		Next
		' cierro los spoilers de numeros de calculo, vuelve al de idCalcTecn
		If MsgIE.oCurrContainer.ID <> "idCalcTecn" Then MsgIE.popContainer
	End Function
	
	Private Function MsgIE_getCalcContainerInCalcTecn (strCalcNum)
		' obtiene (o genera si no existe) un "spoiler no desplegable" para contener la informacion de los ficheros relativos a un calculo dado
		If MsgIE.existsContainer ("id" & strCalcNum) Then
			Set MsgIE_getCalcContainerInCalcTecn = MsgIE.setContainer ("id" & strCalcNum)
		Else
			If MsgIE.oCurrContainer.ID <> "idCalcTecn" Then MsgIE.popContainer
			If MsgIE.oCurrContainer.ID <> "idCalcTecn" Then
				Err.Raise 91, "cOp_CalcsTecn", "Out of context call to MsgIE_CreateCalcContainerInCalcTecn, should be in idCalcTecn Spoiler or an inner one"
			End If
			Set MsgIE_getCalcContainerInCalcTecn = MsgIE.Spoiler (True,"",strCalcNum, "id" & strCalcNum,True)
		End If
	End Function
	
	Private Sub MsgIE_AddRowToCalcContainer (strCalcNum,strCalcOpc,fName)
		Dim currContainer
		' Aseguramos que escribimos en el panel DOC
		MsgIE.setContainer "doc"
		Set currContainer = MsgIE_getCalcContainerInCalcTecn(strCalcNum)
		' meto el listado de ficheros, con sus opciones, como tabla
		Call MsgIE.AddTableRow(currContainer,strCalcOpc,fName)
		Call MsgIE.Spoiler (True,"","", "id" & strCalcOpc,True)
		MsgIE.popContainer
	End Sub

	Private Function getCalcFilesInFolder (strFold,strCalcsId)
		' obtiene un array, listado de todos los ficheros en strFold, que pueden considerarse ASOCIADOS A LA MISMA OPORTUNIDAD que el identificado por strCalcsId
		' ....
		Dim strTmpFich
		strTmpFich = fso.GetSpecialFolder(2) & "\" & fso.GetTempName()
		WshShell.Run "cmd /U /C ""dir /s /b /tw /o-d """ & strFold & "\*" & strCalcsId & "*"" > """ & strTmpFich & """""",0,True
		If fso.getfile (strTmpFich).Size > 0 Then
			getCalcFilesInFolder = Split(fso.OpenTextFile (strTmpFich,1,False,-1).ReadAll,vbCrLf)
		Else
			getCalcFilesInFolder = Array()
		End If
		If fso.fileExists(strTmpFich) Then fso.DeleteFile strTmpFich,true
	End Function

	Function RenombrarFichero_FixRefs (strCalcNum,strCalcOpc, oABCGas_XLS, strFPath, strDestFN, indent)
		Dim c, strOld
		regex.Pattern = "_old\(\d+\)"
		'MsgLog String (indent + 1,vbTab) & "Fichero a renombrar," & vbCrLf & vbtab & fso.getFileName(strFPath)
		'MsgLog String (indent + 1,vbTab) & vbtab & "- Se va a renombrar a: " & fso.getFileName(strDestFN)
		If oDicCalcs(strCalcNum)(strCalcOpc).Exists (strFPath) then
			MsgLog String (indent + 1,vbTab) & vbtab & "- tiene referencia en oDicCalcs, asignado como item: " & TypeName (oDicCalcs(strCalcNum)(strCalcOpc)(strFPath))
		Else
			MsgLog String (indent + 1,vbTab) & vbtab & "- NO tiene referencia en oDicCalcs"
		End If
		If Not IsEmpty (oABCGas_XLS) Then
			If oABCGas_XLS.strXLSXPath <> strFPath Then MsgBox ("Datos inconsistentes para el renombrado") : Exit Function
			If LCase(regex.Replace(oABCGas_XLS.strXLSXPath,"")) = LCase(strDestFN) Then Exit Function
			If LCase (oABCGas_XLS.strXLSXPath) = LCase (strDestFN) Then Exit Function
		End if
		If fso.FileExists (strDestFN) Then
			' añade el old secuencial al existente, dejando libre el nombre para el que quiero renombrar
			c = 1
			strOld = strDestFN
			Do While fso.FileExists (strOld)
				strOld = Left (Replace(strDestFN,".xlsx",""),243) & ".xlsx"
				strOld = Replace (strOld,".xlsx","_old(" & c & ").xlsx")
				'strOld = Replace(strOld,"_old(1)","")
				c = c + 1
			Loop
			MsgLog String (indent + 1,vbTab) & vbtab & " - Existe fichero con el nombre de destino, tambien se renombra"
			'Stop : MsgBox ("esto hay que cambiarlo, hay que cambiar tb la referencia  oDicCalcs(strCalcNum)(strCalcOpc).key(fich) DE ESTE FICHERO!!!. Habrá que devolver nombres orig y cambiados en un dicc,... o algo asi" ) 
			Call RenombrarFichero_FixRefs (strCalcNum,strCalcOpc, oDicCalcs(strCalcNum)(strCalcOpc).item(strDestFN), strDestFN, strOld, indent + 1)
'							fso.MoveFile strDestFN,strOld ' POSIBLEMENTE ESTO AFECTE A LA REFERENCIA EN DICCIONARIO... 
'							oDicCalcs(strCalcNum)(strCalcOpc).key(strDestFN) = strOld
			MsgLog String (indent + 1,vbTab) & vbtab & vbtab & "Renombrado el fichero ya existente," & _
					vbCrLf & vbtab & vbtab & Mid(strDestFN,InStrRev(strDestFN,"\")+1) & " a " & vbCrLf & vbtab & vbtab & Mid(strOld,InStrRev(strOld,"\")+1)
		End If
		On Error Resume Next
		fso.MoveFile strFPath,strDestFN ' se intenta renombrar desde el sistema de archivos, si no DESDE EXCEL!!
		If Err Then
			On Error GoTo 0
			' en ppio oABCGas_XLS SOLO ESTA EMPTY CON FICHEROS 'OLD', porque cuando se llama a RenombrarFichero_FixRefs 'la primera vez', el objeto
			' ya estaba creado... y SI NO, aqui se podría crear...
			If IsEmpty (oABCGas_XLS) Then
				Set oABCGas_XLS = (New cABCGas_XLS).Init (m_ExcelApp, strFPath)
				Set oDicCalcs(strCalcNum)(strCalcOpc).item(strFPath) = oABCGas_XLS
			End if
			RenombrarFichero_FixRefs = oABCGas_XLS.RenombrarFichero (strDestFN)
		Else
			RenombrarFichero_FixRefs = True
		End If
		On Error GoTo 0
		' HABRIA QUE ACTUALIZAR TODAS LAS ESTRUCTURAS DE DATOS EN LAS QUE APARECIESE strFPath, Y USAR strDestFN...
		oDicCalcs(strCalcNum)(strCalcOpc).key(strFPath) = strDestFN
		' las modificaciones en oDicCalcs SE GESTIONAN DESDE procesaCalcTecnicos!!:
		' Set oDicCalcs(strCalcNum)(strCalcOpc).item(strDestFN) = oABCGas_XLS
	End Function

	Function procesaCalcTecnicos (bValidate, bCloseExcelFiles)
		Dim strFPath,strFName,strCalcNum,strCalcOpc
		Dim oABCGas_XLS
		Dim regex,match,strFType,strExt
		Set regex = New RegExp
		regex.Global = True : regex.IgnoreCase = True : regex.multiline = False
		' CALCULOS ABCAIRE, que DEBEN ser realizados en TODAS las ofertas
		' El comentario siguiente va en "VALIDACION DE CALCULOS":
'		MsgIE "<font color=blue>" & "Respecto a las OPCIONES DE CALCULO que debería haber: conds operativas; conds de diseño; El calculo ""RATED"" se hace PARA LAS TEMPERATURAS MAS ALTAS, LUEGO:" & vbCrLf & _
'				"<ul><li> asegurarse de hacer EL DIMENSIONAMIENTO DE POTENCIA DE MOTOR, ** PARA LAS TEMPERATURAS MAS BAJAS **, FIJANDO LAS RPM, y aumentando la PRESION DE SALIDA en un 10% respecto a la RATED / de diseño..." & "</li>" & vbCrLf & _
'				"<li> SI ES UN REACTOR, de H2, etc, SE DIMENSIONA PARA LAS PEORES CONDICIONES, y se comprueban TODAS, FIJANDO VUELTAS RPM, las de la peor condicion" & "</li>" & vbCrLf & _
'				"<li> asegurarse de hacer la VERIFICACION DE DISPARO DE VALVULA DE SEGURIDAD, FIJANDO LAS RPM, y aumentando la PRESION DE SALIDA en un 10% respecto a la RATED / de diseño..." & "</li>" & vbCrLf & _
'				"<li> asegurarse de hacer las VERIFICACIONES EN VTPARES: en particular, que haya LOAD REVERSAL (/ ""que NO haya ROD REVERSAL""); si no, se puede corregir AÑADIENDO MASA ""en la BIELA?? (o en la CRUCETA?)""; e incluso SUBIENDO las RPM, a veces" & "</li>" & vbCrLf & _
'				"<li> la TEMPERATURA DE SALIDA DEL GAS, SIEMPRE la podemos dar 10 POR ENCIMA DE LA DEL AGUA DE REFRIGERACION, usando un INTERCOOLER FINAL (pero SIEMPRE hay que respetar los lims de temperatura del gas!)" & "</li>" _
'				 & "</ul></font>"
		For Each strCalcNum In oDicCalcs
		For Each strCalcOpc In oDicCalcs(strCalcNum)
		For Each strFPath In oDicCalcs(strCalcNum)(strCalcOpc).Keys
			strFName = fso.GetFileName (strFPath)
			' Fijamos el contenedor al spoiler especifico de la opcion (creado en AddRowToCalcContainer)
			' Este spoiler esta en el panel DOC
			Call MsgIE.setContainer ("id" & strCalcOpc)
			MsgLog "Procesando fichero " & strFName ' Cambiado a MsgLog para no ensuciar el panel Main
			' SER00034_01.rtf: INTERESANTE, LO CONVERTIRIA A TXT y aporta alguna info:
			' (PARA VERLO EN TXT, se puede sacar del Menu Calculos --> Ver en ventana, y pulsar en LISTAR, se crea como un txt en carp de calculos, y se abre...)
'					Temp. licuación en cond. de aspiración : 
'					Límite max/min RPM = 
'					Transmisión : 
'					Carrera (mm) : 
'					Diámetro del vástago (mm) :
'					Temp. de Licuación(ºC):    (EN CADA ETAPA!)
'					Coefic. compres. entrada :!   (EN CADA ETAPA!)
'					Coefic. compres. salidad :!  (EN CADA ETAPA!)
'			          Distribución de la potencia consumida :
'			            - Aumento de presión :                      
'			            - Perdidas por calentamiento del aire/gas : 
'			            - Perdidas en válvulas :                    
'			            - Perdidas en partes mecánicas :            
'			            NOTAS:
'						MENSAJES DE AVISO/ERROR:
			' Antipul_SER00053_01.txt : SE SACA DEL CALCULO DE ANTIPULSADORES!!, SE LE DA A CALCULAR Y LUEGO A "EXPORTAR", y crea el fich de texto en la carp de calculos!!!
			' ABC_Gas_Cooler-SER00034_01.xlsx NOS DARIA EL MATERIAL DEL REFRIGERADOR
			'ABC_Aircooler-SER00034_01.xlsx nos da la TEMPERATURA DE ENTRADA DEL AGUA EN EL AIRCOOLER (la de salida la damos como dato)
			'ABC_Reducer-SER00034_01.xlsx: NO TIENEN INFORMACION
			'ABC_Main Motor-SER00034_01.xlsx
			'ABC_Instrumentation-SER00034_01.xlsx
			'ABC_Gas_Filter-SER00034_01.xlsx
			'ABC_Frequency Converter-SER00034_01.xlsx
			'ABC_Cooling Water Pump-SER00034_01.xlsx
			'ABC_Dryer-SER00034_01.xlsx
			'ABC_Piston_rider_ring_selection-SER00034_01.xlsx
			'ABC_Cooling Water Tower-SER00034_01.xlsx
			'ABC_Pressure_Safety_Valve-SER00034_01.xlsx
			'ABC_Valves_selection-SER00034_01.xlsx
			' ficheros Ayuda_fecha_hora.txt: son extracciones de la ventana VTPARES, desde el boton LISTADO (OJO, el "listado" de la ventana CON TODO, no funciona; hay qua hacer por trozos...)
			
			' HABRIA QUE SACAR TB LAS HOJAS API!!, pero NO me salen en el LOCAL...
			regex.Pattern = GasVBNetExportedFNamesPattern
			Set match = regex.Execute(strFName).Item(0)
			Select Case True
				Case match.Submatches(3) = "_calc_multi"
					Stop ' este ya debería poder procesarlo... aunque no tiene ningun interés:
					' LA UNICA DIFERENCIA CON LA HOJA DE GASES DE cABCGas_XLS ES QUE EN ESTE FICHERO LA PRIMERA FILA DEL GAS ESTÁ EN BLANCO!
					' de todos modos. como tienen MULTIPLES OPCIONES DEL MISMO CALCULO, se REUBICA en el arbol de diccionarios:
					oDicCalcs(strCalcNum)(strFPath) = oDicCalcs(strCalcNum)(strCalcOpc)(strFPath)
					oDicCalcs(strCalcNum)(strCalcOpc).remove(strFPath)
				Case match.Submatches(3) = "_calc"
					If TypeName (oDicCalcs(strCalcNum)(strCalcOpc).item(strFPath)) = "cABCGas_XLS" Then
						MsgLog vbtab & "<b>EL FICHERO """ & strFName & """ YA HA SIDO PROCESADO!!, es un fallo que DEBERIA RESOLVERSE EN EL PROGRRAMA gestionando bien el bucle...</b>"
					Else
						On Error Resume Next
						Set oABCGas_XLS = (New cABCGas_XLS).Init (m_ExcelApp, strFPath, m_bAPI618)
						'MsgLog vbtab & "Añadiendo referencia a oABCGas_XLS en oDicCalcs(strCalcNum)(strCalcOpc) para " & strFName & " (valor actual: " & oDicCalcs(strCalcNum)(strCalcOpc).Item(fich) & ")"
						If Err.Number <> 0 Then
							MsgLog vbtab & "Error al abrir " & strFName & ": " & Err.Description
							Err.Clear
						ElseIf Not (oABCGas_XLS Is Nothing) Then
							Set oDicCalcs(strCalcNum)(strCalcOpc).item(strFPath) = oABCGas_XLS
							'Stop
							Call oABCGas_XLS.fixCGASING()
							If bValidate Then oABCGas_XLS.ValidarCalculos ' el objeto de esta funcion es COMPROBAR SI HAY QUE REVISAR LOS CALCULOS...
			
							'Stop
							Dim strDestFN
							strDestFN = oABCGas_XLS.getABCFileName (Empty)
							If strDestFN <> strFPath then
								MsgLog vbtab & "Renombrando el fichero a " & fso.getfilename (strDestFN)
								If RenombrarFichero_FixRefs(strCalcNum,strCalcOpc, oABCGas_XLS, strFPath, strDestFN, 0) Then
									strFPath = oABCGas_XLS.strXLSXPath
								End If
							Else
								MsgLog vbtab & "El fichero ya tiene el nombre apropiado"
							End If
						End If
			            ' Cerrar de forma segura
						' TODO: aplicar este patrón de control de errores y cierre de libros de excel al resto de libros que se 
						' van abriendo en la ejecución del script
						m_ExcelApp.CloseFile strFPath, True
						If Err.Number <> 0 Then
							MsgLog vbtab & "Advertencia: No se pudo cerrar " & strFName
							Err.Clear
						End If
					End If
				Case match.Submatches(0) = "ABC_Gas_Cooler-", match.Submatches(0) = "ABC_Aircooler-"
					Stop
				Case match.Submatches(0) = "ABC_Reducer-", match.Submatches(0) = "ABC_Main Motor-", match.Submatches(0) = "ABC_Instrumentation-", _
						match.Submatches(0) = "ABC_Gas_Filter-", match.Submatches(0) = "ABC_Frequency Converter-", match.Submatches(0) = "ABC_Cooling Water Pump-", _
						match.Submatches(0) = "ABC_Dryer-", match.Submatches(0) = "ABC_Piston_rider_ring_selection-", _
						match.Submatches(0) = "ABC_Cooling Water Tower-", match.Submatches(0) = "ABC_Pressure_Safety_Valve-", match.Submatches(0) = "ABC_Valves_selection-"
				Case match.Submatches(0) = "Antipul_" And match.Submatches(3) = ".txt"
					Stop
				Case match.Submatches(4) = ".rtf"
					Stop
				Case InStr (strFName,"Ayuda_") > 0 And InStr (strFName,".txt") > 0 
					Stop
			End Select
			
			MsgIE.popContainer ' "id" & strCalcOpc
		next
		next
		Next
		'Stop ' SEGUIR AQUI!!!
		Call SortCalcs

		MsgIE "<font color=blue><b>" & "PTE: volcar info de gas_vbnet a TXTs para Ofergas (con los calculos YA AGRUPADOS!)" & "</b></font>"
		If False Then
		stop
		Dim strOfergas, oABCGas_XLS_LimPot
		For Each oABCGas_XLS In oDicCalcsByComprModel.keys 
			' SE GENERA UNA OFERTA PARA CADA MODELO DE COMPRESOR DETERMINADO!! (habría que verificar SI YA HUBIERA UNA "VALORACION ECONOMICA" que case con ese modelo de compresor -->
			' --> la "OFERTA" que genero, tendría que casar con esa valoración económica!
			
			' oABCGas_XLS_LimPot deberia estar definido al crear los items del diccionario oDicCalcsByComprModel!!
			oABCGas_XLS_LimPot = Empty
			' escribe la info para ofergas en un TXT, abre los TXT ¡¡Y LOS BORRA!!: hay que moverlo al bucle de strCalcNum, Y CADA CALCULO LLEVARA LAS OPCIONES DE DISPARO VALV SEG, ETC.
			strOfergas = fso.GetParentFolderName (oABCGas_XLS.strXLSXPath) & "\" & fso.GetBaseName (oABCGas_XLS.strXLSXPath) & ".txt"
			fso.CreateTextFile (strOfergas, true).Write (strOfergas & vbCrLf & "===================" & vbCrLf)
			Call oABCGas_XLS.generarInfoOfergas(strOfergas,oABCGas_XLS_LimPot)
			WshShell.Run """" & strOfergas & """",1,False
			Stop : fso.deletefile strOfergas, True ' QUE NO QUEDEN HUELLAS...
		Next
		End if
	End Function
	
	Private Function SortCalcs
		Dim strCalcNum,strCalcOpc,fich,strFPath,oABCGas_XLS,strModelName
		' reclasifica los calculos en distintos diccionarios
		' LAS OFERTAS LAS HAGO POR MODELO DE COMPRESOR, por lo que de momento me basta con agrupar por modelos: cabría la remota posibilidad de que CON GASES DISTINTOS,
		' en un mismo calculo, salieran distintas opciones, y mismo modelo??? (si acaso, hago un chequeo rapido DE FLAGS DE GASES en oABCGas_XLS)
		For Each strCalcNum In oDicCalcs
		For Each strCalcOpc In oDicCalcs(strCalcNum)
			strModelName = Empty
			For Each strFPath In oDicCalcs(strCalcNum)(strCalcOpc).Keys
				If TypeName (oDicCalcs(strCalcNum)(strCalcOpc)(strFPath)) = "cABCGas_XLS" Then
					Set oABCGas_XLS = oDicCalcs(strCalcNum)(strCalcOpc)(strFPath)
					If Not oDicCalcsByComprModel.exists (oABCGas_XLS.strModelName) Then
						oDicCalcsByComprModel.Add oABCGas_XLS.strModelName,CreateObject("scripting.dictionary")
					End if
					If Not oDicCalcsByComprModel(oABCGas_XLS.strModelName).exists (strCalcOpc) Then
						oDicCalcsByComprModel(oABCGas_XLS.strModelName).Add strCalcOpc, CreateObject("scripting.dictionary")
					End If
					Set oDicCalcsByComprModel(oABCGas_XLS.strModelName)(strCalcOpc)(strFPath) = oDicCalcs(strCalcNum)(strCalcOpc)(strFPath)
					' para revisar donde se meten los ficheros que NO son _calc:
					If strModelName = oABCGas_XLS.strModelName Then
					ElseIf IsEmpty (strModelName) Then
						strModelName = oABCGas_XLS.strModelName
					Else
						strModelName = Empty 
						'Stop ' ficheros dentro de una misma opción de calculo, TIENEN DISTINTOS MODELOS DE COMPRESOR... no se a qué modelo asociarlos
						' me pasa bastante: DISEÑADO UN CALCULO, lo calculo y exporto a Excel, Y LUEGO LO MODIFICO (pej por recomend de Arzuaga), Y VUELVO A GUARDARLO Y EXPORTARLO...
						'' DARIA LUGAR A una entrada de UN NUEVO MODELO... PERO **** NO DEBERIAN ESTAR TODOS LOS FICHEROS *** DE LA OPC DE CALCULO DENTRO DE DOS ENTRADAS!!
					End If
				Else
					Stop ' por comprobar que ficheros están en la opción, que NO son _calc, y no definen modelo...
				End if
			Next
			For Each strFPath In oDicCalcs(strCalcNum)(strCalcOpc).Keys
				If TypeName (oDicCalcs(strCalcNum)(strCalcOpc)(strFPath)) <> "cABCGas_XLS" Then
					' para los ficheros de la opción que NO son _calc, sólo sé a dónde moverlos, si SOLO HAY UN MODELO!
					If Not IsEmpty (strModelName) Then
						If Not IsEmpty (oDicCalcs(strCalcNum)(strCalcOpc)(strFPath)) Then _
								Set oDicCalcsByComprModel(strModelName)(strCalcOpc)(strFPath) = oDicCalcs(strCalcNum)(strCalcOpc)(strFPath)
					Else
						Stop ' no se donde meter este fichero... (ver arriba explicación)
					End if
				End if
			Next
		Next
		Next
		dumpCalcsByComprModel
	End Function
	
	Function dumpCalcsByComprModel
		Dim strModelName,strCalcOpc,strFPath
		' Escribir en el panel MAIN
		MsgIE.setContainer "main"
		For Each strModelName In oDicCalcsByComprModel
			MsgIE strModelName
			For Each strCalcOpc In oDicCalcsByComprModel(strModelName)
				MsgIE vbtab & strCalcOpc
				For Each strFPath In oDicCalcsByComprModel(strModelName)(strCalcOpc)
					'Stop ' tengo problemas CUANDO RENOMBRO LOS FICHEROS, PIERDE LA REFERENCIA...
					MsgIE vbtab & vbtab & fso.GetFileName(strFPath) & vbtab & vbtab & TypeName (oDicCalcsByComprModel(strModelName)(strCalcOpc)(strFPath))
				Next
			Next
		Next
		MsgIE.popContainer ' Restaurar contexto
	End Function
End Class
