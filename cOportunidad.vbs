Class cOportunidad
	Dim strOpFolder,oDicCompressors,oDicOthers
	Private MsgIE
	Private oOp_CalcsTecn,oOp_ValsEcon
	Private bAPI618 ' Flag para indicar si es proyecto API 618
	Private Sub Class_Initialize ()
		Set oDicCompressors = CreateObject("scripting.dictionary")
		Set oDicOthers = CreateObject("scripting.dictionary")
	End Sub
	Private Sub Class_Terminate ()
	End Sub
	
	Function Init (strOpFolder, ExcelApp, MsgIE_, folderName)
		Set Init = Me
		Me.strOpFolder = strOpFolder
		Set MsgIE = MsgIE_
		
		If Not MsgIE Is Nothing Then MsgIE_Init  ' Solo si hay UI
		
		' Preguntar por API 618 al inicio del procesado de la oportunidad
		bAPI618 = (MsgBox("¿El proyecto en " & folderName & " requiere cumplimiento de norma API 618?", 4 + 32, "Normativa API 618") = 6)

		Set oOp_CalcsTecn = (New cOp_CalcsTecn).Init(strOpFolder & "\2.CALCULO TECNICO", ExcelApp, bAPI618)
	    Set oOp_ValsEcon = (New cOp_ValsEcon).Init(strOpFolder & "\3.VALORACION ECONOMICA", ExcelApp)
	End Function

	Function MsgIE_Init()
		' Crea la estructura de presentación para una oportunidad
		If MsgIE Is Nothing Then Exit Function
		MsgIE ("<h3>" & strOpFolder & "</h3>")
	End Function

	Function procesaCarp()
		If IsEmpty(strOpFolder) Then
			Err.Raise 91, "cOportunidad", "Opportunity folder not initialized. Call Init() first."
		End If
		' Dim strStdOfferName_Date
		' strStdOfferName_Date = Mid(Year(Now),3) & String(2-Len(Month(Now)),"0") & Month(Now)
		' If InStr (fso.GetBaseName(strOpFolder),strStdOfferName_Date) = 0 Then WshShell.Popup "OJO, el nombre de la carpeta procesada debería COMENZAR por xx[=41], seguido de " & strStdOfferName_Date & "xxx, en lugar de '" & fso.GetBaseName(strOpFolder) & "'", 8
	
		Call MsgIE.Spoiler (True,Empty, "CALCULOS TECNICOS","idCalcTecn",False)
		On Error Resume Next
		Call oOp_CalcsTecn.identificaCalcTecnicos ()
		Call oOp_CalcsTecn.procesaCalcTecnicos (False, True)
		If Err.Number <> 0 Then
			MsgLog "ERROR en CALCULOS TECNICOS: " & Err.Description & " - Continuando..."
			Err.Clear
		End If
		On Error GoTo 0
		MsgIE.popContainer ' idCalcTecn
	
		Call MsgIE.Spoiler (True,Empty, "VALORACIONES ECONOMICAS","idValEcon",False)
		Call oOp_ValsEcon.identificaValsEconomicas ()
		Call oOp_ValsEcon.procesaCarpValEconomica (oOp_CalcsTecn.oDicCalcs)
		MsgIE.popContainer ' idValEcon
	
	    If Err.Number <> 0 Then
	        MsgLog "Error procesando valoraciones económicas: " & Err.Description
	        Err.Clear
	    End If
	
		' a partir de los vinculos entre valoraciones y calculos técnicos, SE GENERAN LOS BUDGETS Y LAS OFERTAS FORMALES
		Call MsgIE.Spoiler (True,Empty, "OFERTA COMERCIAL","idOfertas",False)
	    Call procesaCarpOfertas(strOpFolder & "\4.OFERTA COMERCIAL", oOp_ValsEcon.oDicValoracs, ExcelApp)
		MsgIE.popContainer ' idOfertas
	    
	    If Err.Number <> 0 Then
	        MsgLog "Error procesando ofertas comerciales: " & Err.Description
	        Err.Clear
	    End If

		' Llamada a ToUISummaryNotes en el contenedor MAIN
		MsgIE.setContainer "main"
		ToUISummaryNotes
		MsgIE.popContainer

		' TOCA RENOMBRAR LA CARPETA DE OPORTUNIDAD, ASEGURANDOSE DE QUE TENGA LOS MODELOS DE COMPRESORES... con validación del usuario! (DE TODOS LOS MODELOS CALCULADOS.. .SOLO TENDRÍA QUE SACAR LOS OFERTADOS, y tal vez NO TODOS!!)
		Call AttemptFolderRename()
	    
	    On Error GoTo 0
	End Function

	' --- Data Extraction Methods ---

	Function GetSummaryData()
		' Placeholder: Esta función devolverá un diccionario con los datos clave
		' extraídos de oOp_CalcsTecn, oOp_ValsEcon, etc.
		' Set GetSummaryData = CreateObject("Scripting.Dictionary")
	End Function

	Sub ToUISummaryNotes()
		' Muestra las notas en un spoiler en el contenedor actual
		Call MsgIE.Spoiler (True,"color:blue", "NOTAS","idNotas",True)
		MsgIE ("UNA VEZ GENERADA LA OFERTA COMECIAL, GENERAR TXTS TAMBIEN CON LA INFO PARA CREATIO")
		MsgIE ("<li> en los precios Y VALORES NUMERICOS, SEPARAR LOS DECIMALES CON PUNTOS, y BORRAR LAS COMAS de los miles!!!</li>")
		MsgIE ("<li> tienen que definirse campos para ""APPLICATION"" OLD, y BRING THE OPPORTUNITY, ENGINEERING, BUYER, USER, ... " & _
			"** CUSTOMER ** (normalmente, el ctto ppal),country, destination plant, DESCRIPTION (podria coincidir con los de ofergas / gas_vbnet)," & _
			" OFFERED BUDGET (*** TIENE QUE SER LA SUMA DE LOS COMPONENTES, DE CADA OFERTA), (quotation nr), " & _
			" ** y para cada oferta añadida **: MODELO del compresor, UNIDADES, precio de cada compresor, cuantia total de la oferta (con sus opcionales), " & _
			"y condics de proceso: gas, caudal, presiones, rpm, power,</li>")
		MsgIE ("Y AÑADIR AL NOMBRE DE LA CARPETA, SI NO LO TIENE, LOS MODELOS DE COMPRESORES!!! (poner una SELECCION, de entre todos los identificables EN LAS OFERTAS / VALORACIONES ECONOMICAS...)")
		MsgIE ("<b>Y **crear un Desktop.ini, con propiedades para Everything</b>")
		MsgIE ("<br/>")
		MsgIE ("<b>BUSCAR, mediante Everything, TERMINOS QUE PUEDAN CARACTERIZAR EL COMPRESOR, <u>DESDE LAS ESPECIFICACIONES</u>. como ATEX, o ""area classification"", ""ZONE ..."" y otros</b>")
		MsgIE.popContainer ' idNotas
	End Sub

	Private Sub AttemptFolderRename()
		' 1. Determinar el nuevo nombre (Placeholder)
		Dim newFolderName : newFolderName = fso.GetBaseName(strOpFolder) & "_RENAMED_TEST"
		Dim parentFolder : parentFolder = fso.GetParentFolderName(strOpFolder)
		Dim newFullPath : newFullPath = parentFolder & "\" & newFolderName

		If LCase(strOpFolder) = LCase(newFullPath) Then Exit Sub ' Ya tiene el nombre correcto

		' 2. Pedir confirmación al usuario
		Dim userResponse
		userResponse = MsgBox("¿Desea renombrar la carpeta de oportunidad?" & vbCrLf & _
							"DE: " & strOpFolder & vbCrLf & _
							"A:  " & newFullPath, 4 + 32, "Renombrar Carpeta de Oportunidad") ' vbYesNo + vbQuestion

		If userResponse <> 6 Then Exit Sub ' vbYes

		' 3. Intentar renombrar
		On Error Resume Next
		fso.MoveFolder strOpFolder, newFullPath
		If Err.Number <> 0 Then
			MsgBox "ERROR: No se pudo renombrar la carpeta." & vbCrLf & "Asegúrese de que ningún fichero dentro de la carpeta esté abierto.", 16, "Error al Renombrar"
		Else
			strOpFolder = newFullPath ' Actualizar la ruta interna de la clase
		End If
		On Error GoTo 0
	End Sub

	Function checkXLWBKType (oXlWorkBook, ByRef oXlSheet)
		On Error Resume Next
		' Validar que el workbook sigue siendo válido
		Dim testName
		testName = oXlWorkBook.Name
		If Err.Number <> 0 Then
			MsgLog "Error: El libro Excel fue cerrado externamente"
			checkXLWBKType = ""
			Err.Clear
			Exit Function
		End If
		
		' OJO, oXlSheet no siempre es la hoja a procesar, depende del tipo de libro.
		Dim oExcelApp
		Set oExcelApp = oXlWorkBook.Application
		Select Case True
			Case oExcelApp.Evaluate("ISREF('" & "Hoja 2" & "'!A1)")
				Set oXlSheet = oXlWorkBook.worksheets("Hoja 2")
				If checkXLWBKType = "" And InStr(oXlSheet.range ("A1"),"Cliente") > 0 And InStr(oXlSheet.range ("K1"),"Fecha Oferta") > 0 And InStr(oXlSheet.range ("H3"),"Comercial") > 0  Then
					checkXLWBKType = "OferGas"
				End if
			Case oExcelApp.Evaluate("ISREF('" & "Hoja2" & "'!A1)") 
				Set oXlSheet = oXlWorkBook.worksheets("Hoja2")
				If checkXLWBKType = "" And InStr(oXlSheet.range ("A1"),"Cliente") > 0 And InStr(oXlSheet.range ("K1"),"Fecha Oferta") > 0 And InStr(oXlSheet.range ("H3"),"Comercial") > 0  Then
					checkXLWBKType = "OferGas_OLD"
				End if
			Case oExcelApp.Evaluate("ISREF('" & "BUDGET_QUOTE" & "'!A1)") 
				If checkXLWBKType = "" And oExcelApp.Evaluate("ISREF('" & "C._TEXTS" & "'!A1)") then
					checkXLWBKType = "BUDGET_QUOTE"
				End if
			Case oExcelApp.Evaluate("ISREF('" & "1._SCOPE_OF_SUPPLY" & "'!A1)") 
				If checkXLWBKType = "" And oExcelApp.Evaluate("ISREF('" & "A._DATA_ENTRY" & "'!A1)") then
					checkXLWBKType = "FULL_QUOTE"
				End if
			Case oExcelApp.Evaluate("ISREF('" & "GAS-ING" & "'!A1)") 
				Set oXlSheet = oXlWorkBook.worksheets("GAS-ING")
				If checkXLWBKType = "" And InStr(oXlSheet.range ("B2"),"CALCULATION - GAS") > 0 And InStr(oXlSheet.range ("A10"),"CALCULATION") > 0 And InStr(oXlSheet.range ("G5"),"Author") > 0 Then
					checkXLWBKType = "ABCAire"
				End if
			Case oExcelApp.Evaluate("ISREF('" & "C-GAS-ING" & "'!A1)") 
				Set oXlSheet = oXlWorkBook.worksheets("C-GAS-ING")
				If checkXLWBKType = "" And InStr(oXlSheet.range ("B2"),"CALCULATION - GAS") > 0 And InStr(oXlSheet.range ("A10"),"CALCULATION") > 0 And InStr(oXlSheet.range ("G5"),"Author") > 0 Then
					checkXLWBKType = "ABCAire"
				End if
			Case oExcelApp.Evaluate("ISREF('" & "API1-SI" & "'!A1)") 
				If checkXLWBKType = "" And oExcelApp.Evaluate("ISREF('" & "API2-SI" & "'!A1)") And oExcelApp.Evaluate("ISREF('" & "API3-SI" & "'!A1)") then
					checkXLWBKType = "Hojas_API"
				End if
			Case oExcelApp.Evaluate("ISREF('" & "PRESUPUESTO" & "'!A1)") 
				If checkXLWBKType = "" And oExcelApp.Evaluate("ISREF('" & "Filtros de aceite y carter" & "'!A1)") And _
						oExcelApp.Evaluate("ISREF('" & "Alimentador" & "'!A1)") And _
						oExcelApp.Evaluate("ISREF('" & "AXAPTA" & "'!A1)") And _
						oExcelApp.Evaluate("ISREF('" & "DENOMINACION VALVULAS" & "'!A1)") then
					checkXLWBKType = "8000hSparesTemplate_FULL"
				End if
			Case oExcelApp.Evaluate("ISREF('" & "API 618 5th ED" & "'!A1)") 
				If checkXLWBKType = "" And oExcelApp.Evaluate("ISREF('" & "Project Documentation" & "'!A1)") then
					checkXLWBKType = "TechnicalTabulationMatrix_TR"
				End if
		End Select
	
		If checkXLWBKType = "TechnicalTabulationMatrix_TR" Then
			For Each oXlSheet In oXlWorkBook.Worksheets
				If Not (oXlSheet.Range ("A1:Q50").find ("SCOPE OF SUPPLY") is Nothing) And _
						Not (oXlSheet.Range ("A1:Q50").find ("PROJECT No.:") is Nothing) And _
						Not (oXlSheet.Range ("A1:Q50").find ("DOC No.:") is Nothing) And _
						oExcelApp.WorksheetFunction.CountIf(oXlSheet.Range("H4:H4"), "*TECHNICAL TABULATION MATRIX FOR API 618*") > 0 Then
					checkXLWBKType = "OK"
					Exit For
				End if
			Next
			If checkXLWBKType = "OK" Then checkXLWBKType = "TechnicalTabulationMatrix_TR" Else checkXLWBKType = "" End If
		ElseIf checkXLWBKType = "OferGas" Or checkXLWBKType = "OferGas_OLD" Then
			Stop ' IDENTIFICAR SI TIENE LA PLANTILLA DE MARGENES.
		ElseIf checkXLWBKType = "" Then
			For Each oXlSheet In oXlWorkBook.Worksheets
				If Not (oXlSheet.Range ("A1:Q50").find ("Cliente : ") is Nothing) And Not (oXlSheet.Range ("A1:Q50").find ("Usuario Final : ") is Nothing) And _
						Not (oXlSheet.Range ("A1:Q50").find ("Núm.Oferta : ") is Nothing) And Not (oXlSheet.Range ("A1:Q50").find ("Fecha Oferta : ") is Nothing) And _
						Not (oXlSheet.Range ("A1:Q50").find ("Preparado por : ") is Nothing) And Not (oXlSheet.Range ("A1:Q50").find ("Gas : ") is Nothing) And _
						Not (oXlSheet.Range ("A1:Q50").find ("Comercial : ") is Nothing) And Not (oXlSheet.Range ("A1:Q50").find ("CABEZAL") is Nothing) Then
					checkXLWBKType = "OferGas_renamed"
					Exit For
				End If
				If Not (oXlSheet.Cells.find ("6.2 Bolting") is Nothing) And Not (oXlSheet.Cells.find ("6.8 Compressor cylinders") is Nothing) _
						And Not (oXlSheet.Cells.find ("6.10 Pistons, piston rods, and piston rings") is Nothing) Then '
					checkXLWBKType = "APIDeviations"
					Exit For
				End If
				If (Not (oXlSheet.Range ("A1:Q50").find ("TR's Requisition ref. (number / rev):") is Nothing) Or Not (oXlSheet.Range ("A1:Q50").find ("BUYER's Requisition ref. (number / rev)") is Nothing)) And _
						(Not (oXlSheet.Range ("A1:Q50").find ("VENDOR's Quotation ref. (number, rev, date):") is Nothing) Or Not (oXlSheet.Range ("A1:Q50").find ("VENDOR's Quotation ref. (number, rev, date)") is Nothing)) And _
						(Not (oXlSheet.Range ("A1:Q50").find ("Deviation / Comment Number") is Nothing) Or Not (oXlSheet.Range ("A1:Q50").find ("VENDOR Deviation / Exception") is Nothing)) Then '
					checkXLWBKType = "APIDeviations_TR"
					'Stop ' son muy interesantes, REFLEJAN LO QUE SE NEGOCIA, O LAS EXCLUSIONES Q APLICAMOS, Y POR QUE LO HACEMOS
					Exit For
				End If
				If Not (oXlSheet.Cells.find ("COMMISSIONING SPARE PARTS") is Nothing) And Not (oXlSheet.Cells.find ("SPECIAL TOOLS") is Nothing) And Not (oXlSheet.Cells.find ("QUOTATION/OFERTA:") is Nothing) Then
					checkXLWBKType = "CommissioningSpareParts_SpecialTools"
					Exit For
				End If
				If Not (oXlSheet.Cells.find ("COMMISSIONING SPARE PARTS") is Nothing) And Not (oXlSheet.Cells.find ("QUOTATION/OFERTA:") is Nothing) Then
					checkXLWBKType = "CommissioningSpareParts"
					Exit For
				End If
				If Not (oXlSheet.Cells.find ("List of special tools ") is Nothing) And Not (oXlSheet.Cells.find ("Part Number") is Nothing) _
						And Not (oXlSheet.Cells.find ("Qty") is Nothing) Then
					checkXLWBKType = "SpecialToolsList"
					Exit For
				End If
				If Not (oXlSheet.Cells.find ("Valvulas de Aspiracion") is Nothing) And Not (oXlSheet.Cells.find ("Válvulas de impulsión") is Nothing) And _
						Not (oXlSheet.Cells.find ("Coeficiente general") is Nothing) And Not (oXlSheet.Cells.find ("Coef. Gases Secos") is Nothing) And _
						Not (oXlSheet.Cells.find ("Modelo origen") is Nothing) And Not (oXlSheet.Cells.find ("GAS origen") is Nothing)  Then
					checkXLWBKType = "8000hSparesTemplate"
					Exit For
				End If
				If Not (oXlSheet.Cells.find ("Nº OFERTA / Offer number:") is Nothing) And Not (oXlSheet.Cells.find ("Modelo de Compresor / Compressor model:") is Nothing) Then
					'Stop
					If oExcelApp.WorksheetFunction.CountIf(oXlSheet.Range("C2:C2"), "*CURVA DE PAR RESISTENTE PARA ARRANQUE EN VACIO*") > 0 Then
						checkXLWBKType = "curva_de_par_resistente_para_arranque"
						Exit For
					End If
				End If
				If Not (oXlSheet.Range ("A1:Q50").find ("MANUFACTURING LOCATIONS, COUNTRY OF ORIGIN PREFERRED VENDORS") is Nothing) And Not (oXlSheet.Range ("A1:Q50").find ("2.3 OTHERS") is Nothing) Then 'And _
					'	oExcelApp.WorksheetFunction.CountIf(oXlSheet.Range("A3:A3"), "*3. List of vendors:*") > 0 Then
					checkXLWBKType = "Subvendor_list"
					Exit For
				End If
				regex.Pattern = "\\[^\\]*SG\-CD[^\\]*"
				If (Not (oXlSheet.Range ("A1:Q50").find ("Q&A Technical Clarifications") is Nothing) And Not (oXlSheet.Range ("A1:Q50").find ("Question Date") is Nothing)) Or _
						(regex.Test (oXlWorkBook.FullName) And Not (oXlSheet.Range ("A1:Q50").find ("Sales Guidelines & Comments and Deviation") is Nothing)) Then
						' CONSULTAS DE ABENGOA  // de REPSOL (Sales guidelines & comments & deviations)
					checkXLWBKType = "Technical_clarifications" ' son documentos que recogen consultas de cliente, y respuestas dadas
					Exit For
				End If
				If Not (oXlSheet.Range ("A1:Q50").find ("APPLICABLE TO:") is Nothing) And Not (oXlSheet.Range ("A1:Q50").find ("AS BUILT") is Nothing) And _
						Not (oXlSheet.Range ("A1:Q50").find ("INFORMATION TO BE COMPLETED BY PURCHASER") is Nothing) And Not (oXlSheet.Range ("A1:Q50").find ("COMPR. THROWS: ") is Nothing) And _
						Not (oXlSheet.Range ("A1:Q50").find ("POWER (KW) / RATED SPEED (RPM)") is Nothing) And Not (oXlSheet.Range ("A1:Q50").find ("YES: PURCHASER TO FILL IN ""REQUIRED CAPACITY"" LINES.") is Nothing) Then
					checkXLWBKType = "Hojas_API"
					Exit For
				End If
				'stop
			Next
		End if
		If checkXLWBKType = "" Then
			Stop
		End if
			
		If Err.Number <> 0 Then
			MsgLog "Error al verificar tipo de libro: " & Err.Description
			checkXLWBKType = ""
			Err.Clear
		End If
		On Error GoTo 0
	End Function
End Class
