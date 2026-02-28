Option Explicit

Class cOferGas
	Dim strXLSXPath
	Dim oCellFecha,oCellCliente,oCellABCGasCalcNr,oCellComment,bNoModel
	Dim oCellAutor,oCellComercial,oCellEndUser,oCellOfferNr,oCellGas,oCellCountry
	Dim oOferGas_XLS_Cabezal,oOferGas_XLS_ModeloChasis,oOferGas_XLS_Cooling_Dampeners,oOferGas_XLS_Instrumentacion,oOferGas_XLS_Motor_Accesorios,oOferGas_XLS_Opciones
	Dim oOferGas_XLS_ManoObra,oOferGas_XLS_Extras,oOferGas_XLS_Otros
	Dim oCellCOMPRESOR_TOTAL,oCellCOMPRESOR_TOTALSobrecoste
	Dim oOferGas_XLS_Ingenieria_Gestion_Estruct_Margen
	Private regex
	Private m_ExcelApp
	Private m_ExcelFMFile  ' cExcelFile wrapper del archivo principal

	Private Sub Class_Initialize()
	    Set regex = New RegExp
		regex.Global = True : regex.IgnoreCase = True : regex.multiline = False
		Set oDicOtrosCostes = CreateObject("scripting.dictionary")
		Set oDicStages = CreateObject("scripting.dictionary")
		Set m_ExcelApp = Nothing
		Set m_ExcelFMFile = Nothing
	End Sub
	
	Private Sub Class_Terminate()
		'Stop : Call CloseWorkBook() ' NO DEBERIA NECESITAR ESTA LLAMADA, EN TANTO QUE USE EL EXCELMANAGER...
	End Sub
	
	' =============================================
	' PROPIEDAD DE CONVENIENCIA PARA COMPATIBILIDAD
	' =============================================
	Private Property Get objExcel
		If m_ExcelApp Is Nothing Then
			Err.Raise 91, "cOferGas", "ExcelApp not initialized. Call Init() first."
		End If
		Set objExcel = m_ExcelApp.Application
	End Property
	
	' =============================================
	' INICIALIZACIÓN 
	' =============================================
	
	Public Function Init(ExcelApp, strXLSXPath_)
		Set Init = Me
		strXLSXPath = strXLSXPath_
		Set m_ExcelApp = ExcelApp
		
		Set m_ExcelFMFile = m_ExcelApp.OpenFile(strXLSXPath, False, False)

		If getOferGasSheetInfo () Is Nothing Then
			'Set Init = Nothing
			Stop : Call CloseWorkBook ' pte de comprobar si lo hago aqui, o en el destructor...: SE CIERRA FUERA, O EN EL DESTRUCTOR, "Init" SIEMPRE dejara el fichero abierto
			Exit Function
		End If

		MsgIE ("Cliente: <b>" & oCellCliente.Value & "</b>")
		MsgIE ("Proyecto: <b>" & oCellProyecto.Value & "</b>")
		MsgIE ("Observaciones: <b>" & oCellObservaciones.Value & "</b>")
	End Function
	
	Public Function CloseWorkBook()
		If Not (m_ExcelFMFile Is Nothing) Then
			If bSave Then m_ExcelFMFile.Save
			m_ExcelApp.CloseFile m_ExcelFMFile.FilePath, False
			Set m_ExcelFMFile = Nothing
		End If
	End Function
	
	Public Property Get strOfergas_Comment
		If IsEmpty (oCellComment) Then Exit Property
		strOfergas_Comment = oCellComment.value
		strOfergas_Comment = Left (strOfergas_Comment,Instr (strOfergas_Comment,vbcr))
		strOfergas_Comment = Left (strOfergas_Comment,Instr (strOfergas_Comment,vbLf))
	End Property
	
	Public Property Get strFecha
		If Not IsEmpty (oCellFecha) Then
			strFecha = oCellFecha.Value
			strFecha = Split (strFecha,"/")(2) & "-" & Split (strFecha,"/")(1) & "-" & Split (strFecha,"/")(0)
		End If
	End Property
	
	Public Function iCilindrosMayorados()
		Dim c, oCil
		c = oStage.Count
		For each oCil in oStage
			If oCil.oCellSobrecoste > 0 and oCil.oCellSobrecoste > 0.5 * oCil.oCellCoste then i = i + 1
		Next
		iCilindrosMayorados = i / c
	End Function
	
	Public Function bCilindrosCumplenSerie(StrABCGasSerie)
		Dim oCil
		For each oCil in oStage
			Call oCil.bCumpleSerie (StrABCGasSerie)
		Next
	End Function
	
	Public Function bCilindrosCumplenMayoracion()
		Dim oCil
	' LA MAYORACION QUE TIENEN QUE TENER LOS CILINDROS PUEDE SER POR SERIE, O POR FORJADO / ENCAMISADO, o por ...
		For each oCil in oStage
			If not oCil.bCumpleMayoracion () Then
				Stop
			end if
		Next
	End Function
	
	' HAY QUE ASEGURARSE DE QUE EN LOS HP SE INCLUYA EL PRECIO DE LA **BOMBA DE ACEITE AUXILIAR**, en todos los modelos (NO SE SI ES EL "API 614" en "Grupos de engrase", o si es el "ESTANDAR"... o si es otro...)
	
	Private oOferGasXSLSheet_	
	Private Function getOferGasSheetInfo ()
		' OBTIENE INFORMACION DE LA HOJA 'GAS' DEL FICHERO DE EXCEL
		
		If Not IsEmpty (oOferGasXSLSheet_) Then
			Set getGASSheetInfo =  oOferGasXSLSheet_
			Exit Function ' SOLO SE PROCESA UNA VEZ esta función: la info que lee NO CAMBIA, --> NO tiene sentido hacerlo más veces
		End If
		If m_ExcelFMFile Is Nothing Then
			Set getGASSheetInfo = Nothing
			Exit Function
		End If
		
		' VALIDAR QUE TENGA UNA HOJA CON INFORMACION DE OFERTA
		Dim oXlSheet,strSheet
		For Each strSheet In m_ExcelApp.GetSheetNames(strXLSXPath)
			Set oXlSheet = m_ExcelFMFile.GetWorksheet(strSheet)
			If otmpSheet.range ("H1").Value = "Núm.Oferta : " And otmpSheet.range ("K1").Value = "Fecha Oferta : " And _
					 otmpSheet.range ("A2").Value = "Preparado por : " And otmpSheet.range ("E1").Value = "Usuario Final : " Then
				Set oXlSheet = otmpSheet
				Exit For
			End If
		Next
		If IsEmpty (oXlSheet) Then
			Set getGASSheetInfo = Nothing
			Exit Function
		End If
		
		' genero referencias para todos los valores de la hoja
		Set oCellCliente = oXlSheet.range ("B1") 
		Set oCellAutor = oXlSheet.range ("C2")
		Set oCellComercial = oXlSheet.range ("I3")
		Set oCellEndUser = oXlSheet.range ("F1")
		Set oCellCountry = oXlSheet.range ("H4")
		Set oCellABCGasCalcNr = oXlSheet.range ("F2")
		Set oCellOfferNr = oXlSheet.range ("I1")
		Set oCellFecha = oXlSheet.range ("L1")
		Set oCellGas = oXlSheet.range ("I2")
		Set oCellComment = oXlSheet.range ("A3")
		bNoModel = oXlSheet.range ("K22").value = 0

		Set oOferGas_XLS_Cabezal = New cOferGas_XLS_Cabezal
		With oOferGas_XLS_Cabezal
			'c = 8
			c = getCellRow (oXlSheet,"Cabezal:","A:A")
			If c <> 8 Then Stop ' Juanan ha metido algo por ahi...
			Set .oCellCabezal_TOTAL = oXlSheet.range ("K" & c - 2)
			Set .oCellCabezal_TOTALSobrecoste = oXlSheet.range ("L" & c - 2)

			Set .oCellCabezal = (New ItemCoste).Init (oXlSheet,c,"B","E","K","L")
			Set .oCellCabezal.prop("NumEtapas") = oXlSheet.range ("E" & c)
			'c = c + 1
	
			For c = c + 2 To c + 7 ' PARA CADA ETAPA:
				If oXlSheet.range ("C" & c).Value <> "" Then 'Exit For ' COMPRUEBO SI LA CELDA DE ETAPA NO ESTÁ EN BLANCO -> GENERA UN CILINDRO
					Set oStage = New cOferGas_XLS_Stage
					With oStage
						' DE MOMENTO DEJO LOS CILINDROS COMO ITEMCOSTES, PERO PODRÍAN TENER SU PROPIA CLASE, CON FUNCIONES PARA CHEQUEAR SI SON CORRECTOS ETC.
						'Set .oCil = (New cCil).Init (oXlSheet,c,"C","D","K","L")
						Set .oCil = (New ItemCoste).Init (oXlSheet,c,"C","D","K","L")
						Set .oCil.prop("EmpaqRefr") = oXlSheet.range ("E" & c)
					End With
					.oDicStages.Add .oDicStages.Count, oStage
				End If
			Next
	
			c = getCellRow (oXlSheet,"Segmentadura:","A:A")
			If c <> 17 Then Stop ' Juanan ha metido algo por ahi...
			'c = 17
			If Not IsEmpty (oXlSheet.range ("H" & c)) Then Set .oCabezal_Segmentadura = (New ItemCoste).Init (oXlSheet,c,"C",Empty,"H",Empty) : c = c + 1
	
			'c = 18
			c = getCellRow (oXlSheet,"Bloque SAS:","A:A")
			If c <> 18 Then Stop ' Juanan ha metido algo por ahi...
			Set .oCabezal_BloqueSAS = (New ItemCoste).Init (oXlSheet,c,"C",Empty,"K","L") : c = c + 1
	
			'c = 20
			c = getCellRow (oXlSheet,"Normativa de Compresor:","A:A")
			If c <> 20 Then Stop ' Juanan ha metido algo por ahi...
			Set .oCabezal_Normativa = (New ItemCoste).Init (oXlSheet,c,"C",Empty,"K","L") : c = c + 1
		End With
		
		Set oOferGas_XLS_ModeloChasis = New cOferGas_XLS_ModeloChasis
		With oOferGas_XLS_ModeloChasis
			'c = 24
			c = getCellRow (oXlSheet,"Chásis:","A:A")
			If c <> 24 Then Stop ' Juanan ha metido algo por ahi...
			Set .oCellModelo_TOTAL = oXlSheet.range ("K" & c - 2)
			Set .oCellModelo_TOTALSobrecoste = oXlSheet.range ("L" & c - 2)

			Set .oModelo_Chasis = (New ItemCoste).Init (oXlSheet,c,"C",Empty,"K","L") : c = c + 1
	
			c = getCellRow (oXlSheet,"Transmisión:","A:A")
			If c <> 25 Then Stop ' Juanan ha metido algo por ahi...
			c = c + 1
			For i = asc("C") To Asc("G") step 2 ' PARA CADA ETAPA:
				d = Chr (i)
				If oXlSheet.range (d & c).Value <> "" Then
					Set .oModelo_Transmision = (New ItemCoste).Init (oXlSheet,c,d,Empty,"K","L")
					Set .oModelo_Transmision.prop("Tipo") = oXlSheet.range (d & c - 1)
					Exit For ' COMPRUEBO SI LA CELDA DE ETAPA NO ESTÁ EN BLANCO -> GENERA UN CILINDRO
				End If
			Next
	
			'c = 28
			c = getCellRow (oXlSheet,"ATEX:","A:A")
			If c <> 28 Then Stop ' Juanan ha metido algo por ahi...
			Set .oModelo_Transmision_ATEX = (New ItemCoste).Init (oXlSheet,c,"C",Empty,"K","L") : c = c + 1
	
			'c = 30
			c = getCellRow (oXlSheet,"Tubería Gas:","A:A")
			If c <> 30 Then Stop ' Juanan ha metido algo por ahi...
			Set .oModelo_TubGas = (New ItemCoste).Init (oXlSheet,c,"F","D","K","L")
			Set .oModelo_TubGas.prop("Material") = oXlSheet.range ("F" & c)
			Set .oModelo_TubGas.prop("Diametro") = oXlSheet.range ("I" & c)
			c = c + 1
	
			'c = 32
			c = getCellRow (oXlSheet,"Tubería Agua:","A:A")
			If c <> 32 Then Stop ' Juanan ha metido algo por ahi...
			Set .oModelo_TubAgua = (New ItemCoste).Init (oXlSheet,c,"C",Empty,"K","L") : c = c + 1
		End With

		''''''''' IMPORTANTE, HAY UN CAMPO EN OFERGAS, EN LA SECCION "REFRIGERADORES", QUE DEFINE *** EL MATERIAL DE LOS TUBOS ***, INOX, ETC, 
		''''''''' Y QUE NO APARECE EN LA HOJA DE EXCEL...
		''''''''' --> HAY QUE VER SI LOS PRECIOS SON RAZONABLES, Y SI NO LO SON, PONER UN MENSAJE....
		MsgIE ("PTE DE VER COMO ACCEDER A LA MARCA 'INOX' DE LAS TUBERIAS DE CALDEROS: LOS DE AIRE VAN EN ""COBRE"", deberia APARECER EN EXCEL")
		Set oOferGas_XLS_Cooling_Dampeners = New cOferGas_XLS_Cooling_Dampeners
		With oOferGas_XLS_Cooling_Dampeners
			'c = 38
			c = getCellRow (oXlSheet,"Refrigeradores entre fases:","A:A")
			If c <> 36 Then Stop ' Juanan ha metido algo por ahi...
			c = c + 2
			For i = 0 To 5 ' PARA CADA ETAPA:
				If oXlSheet.range ("C" & c+i).Value = "" And oXlSheet.range ("K" & c+i).Value = "" Then Exit For ' COMPRUEBO SI LA CELDA DE ETAPA NO ESTÁ EN BLANCO -> GENERA UN CILINDRO
				.oDicInterStages.Add .oDicInterStages.Count, New cOferGas_XLS_InterStage
'				If oDicStages.Count < i Then
'					MsgBox ("Se han definido menos etapas para cilindros que en el numero de calderos! (algunos comerciales lo hacen al definir cilindros Tandem...)")
'					oDicStages.Add oDicStages.Count, New cOferGas_XLS_Stage
'				End if
				With .oDicInterStages(i)
					Set .oCoolers = (New ItemCoste).Init (oXlSheet,c+i,"C","I","K","L")
					Set .oCoolers.prop("Refrig") = oXlSheet.range ("C" & c+i)
					Set .oCoolers.prop("P. de diseño") = oXlSheet.range ("F" & c+i)
					Set .oCoolers.prop("Mat refr. Entrada") = oXlSheet.range ("G" & c+i)
					Set .oCoolers.prop("Mat refr. Salida") = oXlSheet.range ("H" & c+i)
					Set .oCoolers.prop("Cantidad") = oXlSheet.range ("I" & c+i)
				End With
			Next
	
			c = getCellRow (oXlSheet,"V.Seguridad:","A:A")
			If c <> 47 Then Stop ' Juanan ha metido algo por ahi...
			c = c - 2
			Set .oCell_StageCoolers_bSobreespCorros = oXlSheet.range ("C" & c).Value = "SI"
			Set .oCell_StageCoolers_bRX = oXlSheet.range ("E" & c).Value = "SI"
			Set .oCell_StageCoolers_bASME = oXlSheet.range ("G" & c).Value = "SI"
			Set .oCell_StageCoolers_Coste = oXlSheet.range ("I" & c)
			c = c + 1
			Set .oCell_StageCoolers_bSelloU = oXlSheet.range ("G" & c).Value = "SI"
			Set .oCell_StageCoolers_SelloU_Coste = oXlSheet.range ("K" & c)
	
			c = c + 1
			Set .oStageCoolers_ValvSeg = (New ItemCoste).Init (oXlSheet,c,"B","E","K","L") : c = c + 1

			c = getCellRow (oXlSheet,"Calderines:","A:A")
			Set .oCell_StageBoilers_Predefined = oXlSheet.range ("B" & c)
	
			For c = c + 4 To c + 9 ' PARA CADA ETAPA:
				If oXlSheet.range ("B" & c).Value = "" Then Exit For ' COMPRUEBO SI LA CELDA DE ETAPA NO ESTÁ EN BLANCO -> GENERA UN CILINDRO
				With .oDicInterStages(c-53)
					' al especificar CANTIDAD, tanto para ASP como para ESC, NO puedo hacer un solo ItemCoste para TODOS!!!
					Set .oCell_StageBoilers = (New ItemCoste).Init (oXlSheet,c,"B",Empty,"K","L")
					Set .oCell_StageBoilers.prop("ASP_Vol") = oXlSheet.range ("B" & c)
					Set .oCell_StageBoilers.prop("ASP_Pnom") = oXlSheet.range ("C" & c)
					Set .oCell_StageBoilers.prop("ASP_Material") = oXlSheet.range ("D" & c)
					Set .oCell_StageBoilers.prop("ASP_Cant") = oXlSheet.range ("E" & c)
					Set .oCell_StageBoilers.prop("ESC_Vol") = oXlSheet.range ("F" & c)
					Set .oCell_StageBoilers.prop("ESC_Pnom") = oXlSheet.range ("G" & c)
					Set .oCell_StageBoilers.prop("ESC_Material") = oXlSheet.range ("H" & c)
					Set .oCell_StageBoilers.prop("ESC_Cant") = oXlSheet.range ("I" & c)
				End With
			Next

			c = 63
			c = getCellRow (oXlSheet,"Depósito Entrada:","A:A")
			c = c - 3	
			Set .oCell_StageBoilers_bSobreespCorros = oXlSheet.range ("C" & c).Value = "SI"
			Set .oCell_StageBoilers_bRX = oXlSheet.range ("E" & c).Value = "SI"
			Set .oCell_StageBoilers_bASME = oXlSheet.range ("G" & c).Value = "SI"
			Set .oCell_StageBoilers_Coste = oXlSheet.range ("I" & c)
			c = c + 1
			Set .oCell_StageBoilers_bSelloU = oXlSheet.range ("G" & c).Value = "SI"
			Set .oCell_StageBoilers_SelloU_Coste = oXlSheet.range ("K" & c)
				
			c = c + 2
			Set .oStageBoilers_DepositoEntrada = (New ItemCoste).Init (oXlSheet,c,"C",Empty,"K","L") : c = c + 1
			Set .oStageBoilers_DepositoSalida = (New ItemCoste).Init (oXlSheet,c,"C",Empty,"K","L") : c = c + 1

			c = getCellRow (oXlSheet,"CALDERERIA","A:A")
			Set .oCellCaldereria_TOTAL = oXlSheet.range ("K" & c)
			Set .oCellCaldereria_TOTALSobrecoste = oXlSheet.range ("L" & c)
		End With

		Set oOferGas_XLS_Instrumentacion = New cOferGas_XLS_Instrumentacion
		With oOferGas_XLS_Instrumentacion
			'c = 68
			c = getCellRow (oXlSheet,"Transmisor Temperatura:","A:A")
			If getCellRow (oXlSheet,"Válv.Termostáticas:","A:A") - c <> 9 Then Stop ' Juanan ha metido algo por ahi...
			Set .oCellInstrumentacion_TOTAL = oXlSheet.range ("K" & c - 2)
			Set .oCellInstrumentacion_TOTALSobrecoste = oXlSheet.range ("L" & c - 2)

			Set .oInstrum_TransTemp = (New ItemCoste).Init (oXlSheet,c,"C","F","K","L") : c = c + 1
			Set .oInstrum_TransPres = (New ItemCoste).Init (oXlSheet,c,"C","F","K","L") : c = c + 1
			Set .oInstrum_Termometros = (New ItemCoste).Init (oXlSheet,c,"C","F","K","L") : c = c + 1
			Set .oInstrum_Manometros = (New ItemCoste).Init (oXlSheet,c,"C","F","K","L") : c = c + 1
			Set .oInstrum_EVRegulac = (New ItemCoste).Init (oXlSheet,c,"C","F","K","L") : c = c + 1
			Set .oInstrum_SensorCaidaVastago = (New ItemCoste).Init (oXlSheet,c,"C","F","K","L") : c = c + 1
			Set .oInstrum_SensorVibracion = (New ItemCoste).Init (oXlSheet,c,"C","F","K","L") : c = c + 1
			Set .oInstrum_SensorNivelAceite = (New ItemCoste).Init (oXlSheet,c,"C","F","K","L") : c = c + 1
			Set .oInstrum_NivelCondensados = (New ItemCoste).Init (oXlSheet,c,"C","F","K","L") : c = c + 1
			Set .oInstrum_ValvTermostaticas = (New ItemCoste).Init (oXlSheet,c,"C","F","K","L") : c = c + 1
		End With

		Set oOferGas_XLS_Motor_Accesorios = New cOferGas_XLS_Motor_Accesorios
		With oOferGas_XLS_Motor_Accesorios
			'c = 81
			c = getCellRow (oXlSheet,"Motor:","A:A")
			If getCellRow (oXlSheet,"Grupo Engrase:","A:A") - c <> 8 Then Stop ' Juanan ha metido algo por ahi...
			Set .oCellAccess_TOTAL = oXlSheet.range ("K" & c - 2)
			Set .oCellAccess_TOTALSobrecoste = oXlSheet.range ("L" & c - 2)

			Set .oAccess_Motor.prop("NormaATEX") = oXlSheet.range ("C" & c)
			Set .oAccess_Motor.prop("Clase") = oXlSheet.range ("E" & c)
			Set .oAccess_Motor.prop("Fabr") = oXlSheet.range ("G" & c)
			Set .oAccess_Motor.prop("Tension") = oXlSheet.range ("I" & c)
			c = c + 1
			Set .oAccess_Motor = (New ItemCoste).Init (oXlSheet,c,"C",Empty,"K","L")
			If Not IsEmpty (oXlSheet.range ("M" & c).Value) Then Set .oAccess_Motor.prop("Comentario") = oXlSheet.range ("M" & c)
			c = c + 2
			Set .oAccess_Arrancador = (New ItemCoste).Init (oXlSheet,c,"C",Empty,"K","L") : c = c + 1
			If Not IsEmpty (oXlSheet.range ("M" & c).Value) Then Set .oAccess_Arrancador.prop("Comentario") = oXlSheet.range ("M" & c)
			Set .oAccess_CajaLocal = (New ItemCoste).Init (oXlSheet,c,"C",Empty,"K","L") : c = c + 1
			If Not IsEmpty (oXlSheet.range ("M" & c).Value) Then Set .oAccess_CajaLocal.prop("Comentario") = oXlSheet.range ("M" & c)
			Set .oAccess_Filtro = (New ItemCoste).Init (oXlSheet,c,"C",Empty,"K","L") : c = c + 1
			If Not IsEmpty (oXlSheet.range ("M" & c).Value) Then Set .oAccess_Filtro.prop("Comentario") = oXlSheet.range ("M" & c)
			Set .oAccess_Aero = (New ItemCoste).Init (oXlSheet,c,"C",Empty,"K","L") : c = c + 1
			If Not IsEmpty (oXlSheet.range ("M" & c).Value) Then Set .oAccess_Aero.prop("Comentario") = oXlSheet.range ("M" & c)
			Set .oAccess_LlaveEntrada = (New ItemCoste).Init (oXlSheet,c,"C",Empty,"K","L") : c = c + 1
			If Not IsEmpty (oXlSheet.range ("M" & c).Value) Then Set .oAccess_LlaveEntrada.prop("Comentario") = oXlSheet.range ("M" & c)
			Set .oAccess_LlaveSalida = (New ItemCoste).Init (oXlSheet,c,"C",Empty,"K","L") : c = c + 1
			If Not IsEmpty (oXlSheet.range ("M" & c).Value) Then Set .oAccess_LlaveSalida.prop("Comentario") = oXlSheet.range ("M" & c)
			Set .oAccess_GrupoEngrase = (New ItemCoste).Init (oXlSheet,c,"C",Empty,"K","L") : c = c + 1
			If Not IsEmpty (oXlSheet.range ("M" & c).Value) Then Set .oAccess_GrupoEngrase.prop("Comentario") = oXlSheet.range ("M" & c)
		End With

		Set oOferGas_XLS_Opciones = New cOferGas_XLS_Opciones
		With oOferGas_XLS_Opciones
			'c = 93
			c = getCellRow (oXlSheet,"OPCIONES","A:A")
			If getCellRow (oXlSheet,"V.Reguladora:","A:A") - c <> 8 Then Stop ' Juanan ha metido algo por ahi...
			Set .oCellOpciones_TOTAL = oXlSheet.range ("K" & c)
			Set .oCellOpciones_TOTALSobrecoste = oXlSheet.range ("L" & c)
			c = c + 2
			Set .oOpciones_ValvRetenc = (New ItemCoste).Init (oXlSheet,c,"C",Empty,"K","L") : c = c + 1
			If Not IsEmpty (oXlSheet.range ("M" & c).Value) Then Set .oOpciones_ValvRetenc.prop("Comentario") = oXlSheet.range ("M" & c)
			Set .oOpciones_EVAgua = (New ItemCoste).Init (oXlSheet,c,"C",Empty,"K","L") : c = c + 1
			If Not IsEmpty (oXlSheet.range ("M" & c).Value) Then Set .oOpciones_EVAgua.prop("Comentario") = oXlSheet.range ("M" & c)
			Set .oOpciones_EngraseCilindros = (New ItemCoste).Init (oXlSheet,c,"C",Empty,"K","L") : c = c + 1
			If Not IsEmpty (oXlSheet.range ("M" & c).Value) Then Set .oOpciones_EngraseCilindros.prop("Comentario") = oXlSheet.range ("M" & c)
			Set .oOpciones_Purgadores = (New ItemCoste).Init (oXlSheet,c,"C","G","K","L") : c = c + 1
			If Not IsEmpty (oXlSheet.range ("M" & c).Value) Then Set .oOpciones_Purgadores.prop("Comentario") = oXlSheet.range ("M" & c)
			Set .oOpciones_ResistCalefacc = (New ItemCoste).Init (oXlSheet,c,"C",Empty,"K","L") : c = c + 1
			If Not IsEmpty (oXlSheet.range ("M" & c).Value) Then Set .oOpciones_ResistCalefacc.prop("Comentario") = oXlSheet.range ("M" & c)
			Set .oOpciones_Bypass = (New ItemCoste).Init (oXlSheet,c,"C",Empty,"K","L") : c = c + 1
			If Not IsEmpty (oXlSheet.range ("M" & c).Value) Then Set .oOpciones_Bypass.prop("Comentario") = oXlSheet.range ("M" & c)
			Set .oOpciones_ValvReguladora = (New ItemCoste).Init (oXlSheet,c,"C",Empty,"K","L") : c = c + 1
			If Not IsEmpty (oXlSheet.range ("M" & c).Value) Then Set .oOpciones_ValvReguladora.prop("Comentario") = oXlSheet.range ("M" & c)
		End With
		
		Set oOferGas_XLS_ManoObra = New cOferGas_XLS_ManoObra
		With oOferGas_XLS_ManoObra
			'c = 103
			c = getCellRow (oXlSheet,"MANO DE OBRA","A:A")
			If getCellRow (oXlSheet,"EXTRA:","A:A") - c <> 10 Then Stop ' Juanan ha metido algo por ahi...
			Set .oCellManoObra_TOTAL = oXlSheet.range ("K" & c)
			Set .oCellManoObra_TOTALSobrecoste = oXlSheet.range ("L" & c)
			c = c + 2
			Set .oCell_ManoObra_Predefined = oXlSheet.range ("B" & c)
			Set .oManoObraFase1 = (New ItemCoste).Init (oXlSheet,c,"D","E","K","L")
			Set .oManoObraFase1.prop("CosteUnitario") = oXlSheet.range ("F" & c)
			If Not IsEmpty (oXlSheet.range ("M" & c).Value) Then Set .oManoObraFase1.prop("Comentario") = oXlSheet.range ("M" & c)
			c = c + 1
			Set .oManoObraFase2 = (New ItemCoste).Init (oXlSheet,c,"D","E","K","L")
			Set .oManoObraFase2.prop("CosteUnitario") = oXlSheet.range ("F" & c)
			If Not IsEmpty (oXlSheet.range ("M" & c).Value) Then Set .oManoObraFase2.prop("Comentario") = oXlSheet.range ("M" & c)
			c = c + 1
			Set .oManoObraSoldadura= (New ItemCoste).Init (oXlSheet,c,"D","E","K","L")
			Set .oManoObraSoldadura.prop("CosteUnitario") = oXlSheet.range ("F" & c)
			If Not IsEmpty (oXlSheet.range ("M" & c).Value) Then Set .oManoObraSoldadura.prop("Comentario") = oXlSheet.range ("M" & c)
			c = c + 1
			Set .oManoObraProbadero= (New ItemCoste).Init (oXlSheet,c,"D","E","K","L")
			Set .oManoObraProbadero.prop("CosteUnitario") = oXlSheet.range ("F" & c)
			If Not IsEmpty (oXlSheet.range ("M" & c).Value) Then Set .oManoObraProbadero.prop("Comentario") = oXlSheet.range ("M" & c)
			c = c + 1
			Set .oManoObraPintura = (New ItemCoste).Init (oXlSheet,c,"D","E","K","L")
			Set .oManoObraPintura.prop("CosteUnitario") = oXlSheet.range ("F" & c)
			If Not IsEmpty (oXlSheet.range ("M" & c).Value) Then Set .oManoObraPintura.prop("Comentario") = oXlSheet.range ("M" & c)
			c = c + 1
			Set .oManoObraElectrica = (New ItemCoste).Init (oXlSheet,c,"D","E","K","L")
			Set .oManoObraElectrica.prop("CosteUnitario") = oXlSheet.range ("F" & c)
			If Not IsEmpty (oXlSheet.range ("M" & c).Value) Then Set .oManoObraElectrica.prop("Comentario") = oXlSheet.range ("M" & c)
			c = c + 1
	
			Set .oManoObraIngenieria = (New ItemCoste).Init (oXlSheet,c,"B","E","K","L")
			Set .oManoObraIngenieria.prop("CosteUnitario") = oXlSheet.range ("F" & c)
			If Not IsEmpty (oXlSheet.range ("M" & c).Value) Then Set .oManoObraIngenieria.prop("Comentario") = oXlSheet.range ("M" & c)
			c = c + 2
			Set .oManoObraExtra = (New ItemCoste).Init (oXlSheet,c,Empty,Empty,"K","L")
			If Not IsEmpty (oXlSheet.range ("M" & c).Value) Then Set .oManoObraExtra.prop("Comentario") = oXlSheet.range ("M" & c)
		End With

		Set oOferGas_XLS_Extras = New cOferGas_XLS_Extras
		With oOferGas_XLS_Extras
			'c = 115
			c = getCellRow (oXlSheet,"Embalaje:","A:A")
			If c > 0 Then
				Set .oCellExtras_TOTAL = oXlSheet.range ("K" & c - 2)
				Set .oCellExtras_TOTALSobrecoste = oXlSheet.range ("L" & c - 2)

				Set .oExtras_Embalaje = (New ItemCoste).Init (oXlSheet,c,"C",Empty,"K","L") : c = c + 1
				Set .oExtras_Transporte = (New ItemCoste).Init (oXlSheet,c,"C",Empty,"K","L") : c = c + 1
				Set .oExtras_PuestaEnMarcha = (New ItemCoste).Init (oXlSheet,c,"C",Empty,"K","L") : c = c + 1		
			End If
		End With

		' *********** ASEGURARSE DE QUE EN ESTA SECCION APARECEN TODAS LAS OPCIONES HABITUALES:
		' - las de API-618...
		' - la de ENCAMISADO TANDEM,...
		' ETC
		Dim oItemCoste
		Set oOferGas_XLS_Otros = New cOferGas_XLS_Otros
		With oOferGas_XLS_Otros
			' JUANAN SUELE METER ALGUNAS LINEAS CON COSTES ADICIONALES, encima de los EXTRAS, cuando la hoja NO hubiera generado los campos "extras" u "otros"...:
			c = getCellRow (oXlSheet,"EXTRAS","A:A")
			If c = 0 Then c = getCellRow (oXlSheet,"OTROS","A:A")
			If c = 0 Then c = getCellRow (oXlSheet,"NOTAS","A:A")
			d = getCellRow (oXlSheet,"Ingeniería:","A:A") + 3
			If objExcel.Application.WorksheetFunction.CountA(oXlSheet.Range("D" & d & ":J" & c - 1)) > 0 Then 
				MsgBox ("En el fichero '" & fso.GetAbsolutePathName(strXLSXPath) & "' HAY INFORMACION ADICIONAL ENTRE LA MANO DE OBRA Y EL PRECIO TOTAL!!!")
				Stop
				For c = d To c -1 
					If oXlSheet.range ("K" & c).Value <> "" Then
						Set oItemCoste = .AddCoste (oXlSheet,c,"DEFGHIJ",Empty,"K","L")
						' en algunos casos Juanan mete comentarios EN LA COLUMNA M; en varios casos, el "comentario" es un cálculo de "cómo quedaría el precio de ese componente, al afectarlo del ratio "precio final del compresor" / "precio del compresor, sin ingeniería etc"
						If Not IsEmpty (oXlSheet.range ("M" & c)) Then Set oItemCoste.prop("Comentario") = oXlSheet.range ("M" & c)
					End If
				Next
			End if

			'c = 121
			c = getCellRow (oXlSheet,"OTROS","A:A")
			If c > 0 Then
				Set .oCellOtrosCostes_TOTAL = oXlSheet.range ("K" & c)
				Set .oCellOtrosCostes_TOTALSobrecoste = oXlSheet.range ("L" & c)
				c = c + 2
				Do While oXlSheet.range ("C" & c).Value <> ""
					Set oItemCoste = .AddCoste (oXlSheet,c,"CDE",Empty,"K","L")
					' en algunos casos Juanan mete comentarios EN LA COLUMNA M; en varios casos, el "comentario" es un cálculo de "cómo quedaría el precio de ese componente, al afectarlo del ratio "precio final del compresor" / "precio del compresor, sin ingeniería etc"
					' (Y EFECTIVAMENTE, LOS RATIOS "% materiales", "%gestion", y "%margen" SE APLICAN TB SOBRE LAS PARTIDAS "OTROS"!!!, lo que de manera EFECTIVA, INCREMENTA SU COSTE!!!)
					' (viene a dar un colchón de un 4% en esas partidas...)
					If Not IsEmpty (oXlSheet.range ("M" & c)) Then Set oItemCoste.prop("Comentario") = oXlSheet.range ("M" & c)
					If Not IsEmpty (oXlSheet.range ("F" & c)) Then Set oItemCoste.prop("Comentario") = oXlSheet.range ("F" & c)
					'.oDicOtrosCostes.Add .oDicOtrosCostes.Count, (New ItemCoste).Init (oXlSheet,c,"C",Empty,"K","L")
					c = c + 1
				Loop
			End If
		End With

		c = c + 3
		c = getCellRow (oXlSheet,"COSTE TOTAL DEL COMPRESOR","A:A")
		If c > 0 Then
			Set oCellCOMPRESOR_TOTAL = oXlSheet.range ("K" & c)
			Set oCellCOMPRESOR_TOTALSobrecoste = oXlSheet.range ("L" & c)
		End If
		
		''' INGENIERIA Y MULTIPPLICADORES: 
		''' OJO!!!, VALORAR ***** SI HAY AGENTE, O NO *****, para el precio final!!!
		Dim oCellPVta,oCellCA, oCellPlanos, oXlTempSheet,strCellTest
		Set oOferGas_XLS_Ingenieria_Gestion_Estruct_Margen = New cOferGas_XLS_Ingenieria_Gestion_Estruct_Margen
		With oOferGas_XLS_Ingenieria_Gestion_Estruct_Margen
			c = oCellCOMPRESOR_TOTAL.row + 1
			Set oCellPlanos = oXlSheet.Range ("A" & c & ":Q" & c + 50).find ("Nº planos/docs")
			If oCellPlanos is Nothing Then
				For Each oXlTempSheet In oXlWorkBook.Worksheets
					Set oCellPlanos = oXlTempSheet.Cells.find ("Nº planos/docs")
					If Not (oCellPlanos is Nothing) Then Exit For
				Next
			End If
			If Not (oCellPlanos Is Nothing) then
				c = oCellPlanos.row - 1
				d = oCellPlanos.column + 1
				Set .oCellNumCompresores = oXlSheet.cells (c, d) : c = c + 1
				Set .oCellNumPlanos = oXlSheet.cells (c, d) : d = d + 1
				Set .oCellHorasPLano = oXlSheet.cells (c, d) : c = c + 1 : d = d - 1
				Set .oCellCosteSeguMaq = oXlSheet.cells (c, d) : c = c + 1
				Set .oCellHorasIngMaq = oXlSheet.cells (c, d) : c = c + 1
			End If

			c = oCellCOMPRESOR_TOTAL.row + 1
			Set .oCellMOD = oXlSheet.Range ("A" & c & ":Q" & c + 50).find ("MOD")
			If .oCellMOD is Nothing Then
				For Each oXlTempSheet In oXlWorkBook.Worksheets
					Set .oCellMOD = oXlTempSheet.Cells.find ("MOD")
					If Not (.oCellMOD is Nothing) Then Exit For
				Next
			End If
			'stop
			If Not (oXlSheet.cells (.oCellMOD.row - 1, .oCellMOD.column). Value = "Material" And oXlSheet.cells (.oCellMOD.row, .oCellMOD.column + 1). Value = "Ad-hoc") Then
				MsgBox ("No se ha encontrado el patrón de cálculo para horas de ingeniería y multiplicadores, Asegúrate de que está en la hoja y tiene el formato habitual")
			Else
				c = .oCellMOD.row - 1
				d = .oCellMOD.column + 3
				Set .oCellMaterial = oXlSheet.cells (c, d)
				If .oCellMaterial <> oCellCOMPRESOR_TOTAL - oOferGas_XLS_ManoObra.oCellManoObra_TOTAL Then
					WScript.Echo "Se ha aplicado algun factor de correccion al coste de materiales (puede que sea la resta de approach, o de test):" & 100 * .oCellMaterial/(oCellCOMPRESOR_TOTAL - oOferGas_XLS_ManoObra.oCellManoObra_TOTAL) & "%"
					If oXlSheet.cells (c, d+1) <> "" Then WScript.Echo vbTab & oXlSheet.cells (c, d+1)
				End If
				c = c + 1
				Set .oCellMOD = oXlSheet.cells (c, d) : c = c + 1
				If oOferGas_XLS_ManoObra.oManoObraPintura.oCellCoste > 0 And _
						InStr(oOferGas_XLS_ManoObra.oManoObraPintura.oCellCoste.Dependents.Address,.oCellMOD.Address) = 0 Then
					Stop ' debe incluir oCellMOD!!
					WScript.Echo "OJO!! en MOD, el coste de PINTURA ¿TIENE QUE ESTAR RESTADO?!!!!"
				End If
				Set .oCellSubcontratac = oXlSheet.cells (c, d) : c = c + 1
				c = c + 1
				Set .oCellFactorMaterial = oXlSheet.cells (c, d-2)
				Set .oCellPlusOnMaterial = oXlSheet.cells (c, d) : c = c + 1
				Set .oCellExpedicion = oXlSheet.cells (c, d) : c = c + 1
				Set .oCellDesignOG_NumHoras = oXlSheet.cells (c, d-1)
				Set .oCellDesignOG = oXlSheet.cells (c, d) : c = c + 1
				If oOferGas_XLS_ManoObra.oManoObraIngenieria.oCellCoste = .oCellDesignOG And _
						InStr(oOferGas_XLS_ManoObra.oManoObraIngenieria.oCellCoste.Dependents.Address,.oCellMOD.Address) = 0 Then
					Stop  ' debe incluir oCellMOD!!
					WScript.Echo "Se han incluido las horas de ingenieria en el OferGas --> en MOD, TIENEN QUE ESTAR RESTADAS!!!!"
				End If
				Set .oCellFactorGestionOG = oXlSheet.cells (c, d-2)
				Set .oCellGestionOG = oXlSheet.cells (c, d) : c = c + 1
				Set .oCellFactorGastosEstruct = oXlSheet.cells (c, d-2)
				Set .oCellGastosEstruct = oXlSheet.cells (c, d) : c = c + 1
				Set .oCellFactorMargen = oXlSheet.cells (c, d-2)
				Set .oCellMargen = oXlSheet.cells (c, d) : c = c + 2
			End If
			'	Stop
			c = oCellCOMPRESOR_TOTAL.row + 1
			Set oCellPVta = oXlSheet.Range ("A" & c & ":Q" & c + 50).find ("Precio Venta") ' es como meten algunos el precio, incluido el margen del comercial
			If oCellPVta is Nothing Then
				For Each oXlTempSheet In oXlWorkBook.Worksheets
					Set oCellPVta = oXlTempSheet.Cells.find ("Precio Venta")
					If Not (oCellPVta is Nothing) Then Exit For
				Next
			End If
			If oCellPVta is Nothing Then ' algunas hojas tienen un campo TOTAL debajo del de margen, en vez del Precio Venta
				stop
				Set oCellPVta = oXlSheet.Range ("A" & c & ":Q" & c + 50).find ("Margen") 
				c = oCellPVta.row + 1
				If objExcel.Application.WorksheetFunction.CountA(oXlSheet.Range("A" & c & ":Q" & c + 50)) > 2 Then 
					MsgBox ("En el fichero '" & fso.GetAbsolutePathName(strXLSXPath) & "' HAY INFORMACION ADICIONAL después del precio total!!!")
					Stop ' EN EL CASO DE ESOS TOTALES, NO SUELE HABER NADA MAS BAJO EL MARGEN... 
				End if
				Set oCellPVta = oXlSheet.Range ("A" & c & ":Q" & c + 50).find ("TOTAL")
				If oCellPVta is Nothing Then Set oCellPVta = oXlSheet.Range ("A" & c & ":Q" & c + 50).find ("total")
			End If
			c = oCellCOMPRESOR_TOTAL.row + 1
			Set oCellCA = oXlSheet.Range ("A" & c & ":Q" & c + 50).find ("Comision agente") ' es como meto yo los precios
			If oCellCA is Nothing Then
				For Each oXlTempSheet In oXlWorkBook.Worksheets
					Set oCellCA = oXlTempSheet.Cells.find ("Comision agente")
					If Not (oCellCA is Nothing) Then Exit For
				Next
			End If
			If Not (oCellPVta Is Nothing) then
				c = oCellPVta.row
				d = oCellPVta.column + 1
				Set .oCellTOTAL = oXlSheet.cells (c, d) : c = c + 1
				If oXlSheet.cells (oCellPVta.row + 1, oCellPVta.column). Value = "MC"  Then
					'stop
					Set .oCellTOTALconComisAgente = oXlSheet.cells (c, d)
				End If
			ElseIf Not (oCellCA Is Nothing) then 
				' mi forma de meter el margen para los comerciales
				stop
				c = oCellPVta.row - 1
				d = oCellPVta.column + 3
				Set .oCellTOTAL = oXlSheet.cells (c, d) : c = c + 1
				Set .oCellFactorComisAgente = oXlSheet.cells (c, d-2) : c = c + 1
				Set .oCellTOTALconComisAgente = oXlSheet.cells (c, d)
			Else
				Stop
			End If
		End With

		Dump

		Set getOferGasSheetInfo = oXlSheet
		Set oOferGasXSLSheet_ =  getOferGasSheetInfo
	End Function
	
	Private strModelName_
	Function strModelName
		If Not IsEmpty (strModelName_) Then strModelName = strModelName_ : Exit Function
		
		' el nombre del modelo se puede deducir del compresor vinculado por el cálculo, o si está definido en "observaciones"
		regex.Pattern = strFullModelPattern
		If regex.Test () Then
			strModelName = regex.Execute ().Item(0).Value
		Else
			If IsEmpty () Then
				SEGUIR AQUI
			End If
		End If
		
		strModelName_ = strModelName
	End Function
	
	Sub Dump
		Wscript.echo "Cliente: " & oCellCliente & vbTab & "EndUser: " & oCellEndUser  & vbTab & "Country: " & oCellCountry 
		Wscript.echo "OfferNr: " & oCellOfferNr & vbTab & "Fecha: " & oCellFecha 
		Wscript.echo "Gas: " & oCellGas & vbTab & "ABCGasCalcNr: " & oCellABCGasCalcNr
		Wscript.echo "Autor: " & oCellAutor & vbTab & "Comercial: " & oCellComercial
		oOferGas_XLS_Cabezal.Dump
		oOferGas_XLS_ModeloChasis.Dump
		oOferGas_XLS_Cooling_Dampeners.Dump
		oOferGas_XLS_Instrumentacion.Dump
		oOferGas_XLS_Motor_Accesorios.Dump
		oOferGas_XLS_Opciones.Dump
		oOferGas_XLS_ManoObra.Dump
		oOferGas_XLS_Extras.Dump
		oOferGas_XLS_Otros.Dump
		WScript.Echo "COMPRESOR, Coste total: " & oCellCOMPRESOR_TOTAL & vbTab & "Sobrecoste: " & oCellCOMPRESOR_TOTALSobrecoste
		
		If Not IsEmpty (oOferGas_XLS_Ingenieria_Gestion_Estruct_Margen) Then oOferGas_XLS_Ingenieria_Gestion_Estruct_Margen.Dump
	End Sub
	
	Function getCellRow (oXlSheet,strContent,strColumnRange)
	    If oXlSheet.Application.WorksheetFunction.CountIf(oXlSheet.Range(strColumnRange), strContent) > 0 Then
	   		getCellRow = oXlSheet.Application.Match (strContent,oXlSheet.Range(strColumnRange), 0)
	    Else
	        'MsgBox strContent & " does not exist in range " & oXlSheet.Range(strColumnRange).Address
	    End If
		' si estuviera buscando una columna, para devolver la letra de columna:
		' getCellRow = oXlSheet.Application.WorksheetFunction.ADDRESS(1,getCellRow,4)
		' getCellRow = replace (getCellRow,"1","")
	End Function
End Class

Class cOferGas_XLS_ModeloChasis
	Dim oModelo_Chasis,oModelo_Transmision,oModelo_Transmision_ATEX,oModelo_TubGas,oModelo_TubAgua,oCellModelo_TOTAL,oCellModelo_TOTALSobrecoste
	Sub Dump
		If Not IsEmpty(oModelo_Chasis) then WScript.Echo "Modelo y Chasis" : oModelo_Chasis.Dump
		If Not IsEmpty(oModelo_Transmision) then WScript.Echo "Transmision" : oModelo_Transmision.Dump
		If Not IsEmpty(oModelo_Transmision_ATEX) then WScript.Echo "Chasis - ATEX" : oModelo_Transmision_ATEX.Dump
		If Not IsEmpty(oModelo_TubGas) then WScript.Echo "Tuberías de gas" : oModelo_TubGas.Dump
		If Not IsEmpty(oModelo_TubAgua) then WScript.Echo "Tuberías de agua" : oModelo_TubAgua.Dump
		WScript.Echo "Coste total: " & oCellModelo_TOTAL
		WScript.Echo "Sobrecoste: " & oCellModelo_TOTALSobrecoste
	End Sub
End Class

Class cOferGas_XLS_Cabezal
	Dim oDicStages,oCellCabezal,oCabezal_Segmentadura,oCabezal_BloqueSAS,oCabezal_Normativa
	Dim oCellCabezal_TOTAL,oCellCabezal_TOTALSobrecoste
	Private Sub Class_Initialize()
		Set oDicStages = CreateObject("scripting.dictionary")
	End Sub
	Sub Dump
		Dim oStage
		WScript.Echo "Cabezal"
		oCellCabezal.Dump
		For Each oStage In oDicStages.Items
			oStage.Dump
		Next
		If Not IsEmpty(oCabezal_Segmentadura) then WScript.Echo "Segmentadura" : oCabezal_Segmentadura.Dump
		If Not IsEmpty(oCabezal_BloqueSAS) then WScript.Echo "Bloque SAS" : oCabezal_BloqueSAS.Dump
		If Not IsEmpty(oCabezal_Normativa) then WScript.Echo "Normativa" : oCabezal_Normativa.Dump
		WScript.Echo "Coste total: " & oCellCabezal_TOTAL
		WScript.Echo "Sobrecoste: " & oCellCabezal_TOTALSobrecoste
	End Sub
End Class

Class cOferGas_XLS_Stage
	Dim oCil
	Sub Dump
		oCil.Dump
	End Sub
End Class

Class cCil
	Private  oItemCoste
	Private Sub Class_Initialize
	End Sub
	Private Sub Class_Terminate
	End Sub
	Public Function Init (oXlSheet, fila, colId, colNr,colCoste,colSobrecoste)
		Set Init = Me
		Set oItemCoste = (New ItemCoste).Init (oXlSheet, fila, colId, colNr,colCoste,colSobrecoste)
	End Function
	Public Function bCumpleSerie (StrABCGasSerie)
		Select Case True
			case StrABCGasSerie = "HA"
				If instr (oItemcoste.strID,StrABCGasSerie) > 0 Or ((instr (oItemcoste.strID,"HG") > 0 Or instr (oItemcoste.strID,"HP") > 0) And oCellSobrecoste < 0) Then
					bCumpleSerie = True
				Else
					Wscript.echo "El cilindro " & oItemcoste.strID & " NO ESTA CORRECTAMENTE SELECCIONADO, debería ser un HA"
				end if
			case StrABCGasSerie = "HG"
				If instr (oItemcoste.strID,StrABCGasSerie) > 0 Or ((instr (oItemcoste.strID,"HP") > 0 Or instr (oItemcoste.strID,"HX") > 0) And oCellSobrecoste < 0) _
						Or ((instr (oItemcoste.strID,"HA") > 0 ) And oCellSobrecoste > 0) Then
					bCumpleSerie = True
				Else
					Wscript.echo "El cilindro " & oItemcoste.strID & " NO ESTA CORRECTAMENTE SELECCIONADO, debería ser un HG, HA mayorado, o HP / HX minorado"
				end if
			case StrABCGasSerie = "HP"
				If instr (oItemcoste.strID,StrABCGasSerie) > 0 Or (instr (oItemcoste.strID,"HX") > 0 And oCellSobrecoste < 0) _
						Or ((instr (oItemcoste.strID,"HP") > 0 Or instr (oItemcoste.strID,"HA") > 0) And oCellSobrecoste > 0) Then
					bCumpleSerie = True
				Else
					Wscript.echo "El cilindro " & oItemcoste.strID & " NO ESTA CORRECTAMENTE SELECCIONADO, debería ser un HP, HA / HG mayorado, o HX minorado"
				end if
			case StrABCGasSerie = "HX"
				If instr (oItemcoste.strID,StrABCGasSerie) > 0 Or ((instr (oItemcoste.strID,"HP") > 0 Or instr (oItemcoste.strID,"HG") > 0 Or instr (oItemcoste.strID,"HA") > 0) And oCellSobrecoste > 0) Then
					bCumpleSerie = True
				Else
					Wscript.echo "El cilindro " & oItemcoste.strID & " NO ESTA CORRECTAMENTE SELECCIONADO, debería ser un HX, o HA / HG / HP mayorado"
				end if
			Case Else
				Stop
		End Select
	End Function
	Sub Dump
		oItemCoste.Dump 
	End Sub
End Class

Class cOferGas_XLS_Cooling_Dampeners
	Dim oDicInterStages, oCell_StageCoolers_bSobreespCorros, oCell_StageCoolers_bRX, oCell_StageCoolers_bASME, oCell_StageCoolers_Coste, oCell_StageCoolers_bSelloU, oCell_StageCoolers_SelloU_Coste
	Dim oStageCoolers_ValvSeg
	Dim oCell_StageBoilers_Predefined
	Dim oCell_StageBoilers_bSobreespCorros, oCell_StageBoilers_bRX, oCell_StageBoilers_bASME, oCell_StageBoilers_Coste, oCell_StageBoilers_bSelloU, oCell_StageBoilers_SelloU_Coste
	Dim oStageBoilers_DepositoEntrada,oStageBoilers_DepositoSalida,oCellCaldereria_TOTAL,oCellCaldereria_TOTALSobrecoste
	
	Private Sub Class_Initialize()
		Set oDicInterStages = CreateObject("scripting.dictionary")
	End Sub
	Sub Dump
		Dim c,oOferGas_XLS_InterStage
		WScript.Echo "Refrigeradores"
		c = 1
		For Each oOferGas_XLS_InterStage In oDicInterStages.Items
			If Not IsEmpty (oOferGas_XLS_InterStage.oCoolers) Then WScript.Echo "Entre-etapa " & c : oOferGas_XLS_InterStage.oCoolers.Dump
			c = c + 1
		Next
		WScript.Echo "Sobreesp Corrosión: " & oCell_StageCoolers_bSobreespCorros
		WScript.Echo "bRX: " & oCell_StageCoolers_bRX
		WScript.Echo "bASME: " & oCell_StageCoolers_bASME
		WScript.Echo "Coste Extra: " & oCell_StageCoolers_Coste
		WScript.Echo "Sello U: " & oCell_StageCoolers_bSelloU
		WScript.Echo "Coste Sello U: " & oCell_StageCoolers_SelloU_Coste
		WScript.Echo "Valvulas de seguridad" 
		oStageCoolers_ValvSeg.Dump
		WScript.Echo "Calderines antipulsadores - basados en plantilla: " & oCell_StageBoilers_Predefined
		c = 1
		For Each oOferGas_XLS_InterStage In oDicInterStages.Items
			WScript.Echo "Entre-etapa " & c
			c = c + 1
			If Not IsEmpty (oOferGas_XLS_InterStage.oCell_StageBoilers) Then oOferGas_XLS_InterStage.oCell_StageBoilers.Dump
		Next
		WScript.Echo "Sobreesp Corros: " & oCell_StageBoilers_bSobreespCorros
		WScript.Echo "bRX: " & oCell_StageBoilers_bRX
		WScript.Echo "bASME: " & oCell_StageBoilers_bASME
		WScript.Echo "Coste Extra: " & oCell_StageBoilers_Coste
		WScript.Echo "Sello U: " & oCell_StageBoilers_bSelloU
		WScript.Echo "Coste Sello U: " & oCell_StageBoilers_SelloU_Coste
	End Sub
End Class

Class cOferGas_XLS_InterStage
	Dim oCoolers,oCell_StageBoilers
	Sub Dump
		If Not IsEmpty (oCoolers) Then oCoolers.Dump
		If Not IsEmpty (oCell_StageBoilers) Then oCell_StageBoilers.Dump
	End Sub
End Class

Class cOferGas_XLS_Instrumentacion
	Dim oInstrum_TransTemp, oInstrum_TransPres, oInstrum_Termometros, oInstrum_Manometros, oInstrum_EVRegulac, oInstrum_SensorCaidaVastago, oInstrum_SensorVibracion
	Dim oInstrum_SensorNivelAceite, oInstrum_NivelCondensados, oInstrum_ValvTermostaticas
	Dim oCellInstrumentacion_TOTAL, oCellInstrumentacion_TOTALSobrecoste
	Sub Dump
		If Not IsEmpty(oInstrum_TransTemp) then Wscript.Echo "Transductor Temp" : oInstrum_TransTemp.Dump
		If Not IsEmpty(oInstrum_TransPres) then Wscript.Echo "Transductor Presion" : oInstrum_TransPres.Dump
		If Not IsEmpty(oInstrum_Termometros) then Wscript.Echo "Termometros" : oInstrum_Termometros.Dump
		If Not IsEmpty(oInstrum_Manometros) then Wscript.Echo "Manometros" : oInstrum_Manometros.Dump
		If Not IsEmpty(oInstrum_EVRegulac) then Wscript.Echo "EV Regulac" : oInstrum_EVRegulac.Dump
		If Not IsEmpty(oInstrum_SensorCaidaVastago) then Wscript.Echo "Sensor Caida Vastago" : oInstrum_SensorCaidaVastago.Dump
		If Not IsEmpty(oInstrum_SensorVibracion) then Wscript.Echo "Sensor Vibracion" : oInstrum_SensorVibracion.Dump
		If Not IsEmpty(oInstrum_SensorNivelAceite) then Wscript.Echo "Sensor Nivel Aceite" : oInstrum_SensorNivelAceite.Dump
		If Not IsEmpty(oInstrum_NivelCondensados) then Wscript.Echo "Nivel Condensados" : oInstrum_NivelCondensados.Dump
		If Not IsEmpty(oInstrum_ValvTermostaticas) then Wscript.Echo "Valv. Termostaticas" : oInstrum_ValvTermostaticas.Dump
		
		Wscript.Echo "TOTAL Instrumentación: " & oCellInstrumentacion_TOTAL
		Wscript.Echo "Sobrecoste Instrumentación: " & oCellInstrumentacion_TOTALSobrecoste
	End Sub
End Class

Class cOferGas_XLS_Motor_Accesorios
	Dim oAccess_Motor, oAccess_Arrancador, oAccess_CajaLocal, oAccess_Filtro, oAccess_Aero, oAccess_LlaveEntrada, oAccess_LlaveSalida
	Dim oAccess_GrupoEngrase
	Dim oCellAccess_TOTAL, oCellAccess_TOTALSobrecoste
	Sub Dump
		If Not IsEmpty (oAccess_Motor) Then Wscript.Echo "Motor" : oAccess_Motor.Dump
		If Not IsEmpty (oAccess_Arrancador) Then Wscript.Echo "Arrancador" : oAccess_Arrancador.Dump
		If Not IsEmpty (oAccess_CajaLocal) Then Wscript.Echo "Caja Local" : oAccess_CajaLocal.Dump
		If Not IsEmpty (oAccess_Filtro) Then Wscript.Echo "Filtro" : oAccess_Filtro.Dump
		If Not IsEmpty (oAccess_Aero) Then Wscript.Echo "Aero" : oAccess_Aero.Dump
		If Not IsEmpty (oAccess_LlaveEntrada) Then Wscript.Echo "Llave Entrada" : oAccess_LlaveEntrada.Dump
		If Not IsEmpty (oAccess_LlaveSalida) Then Wscript.Echo "Llave Salida" : oAccess_LlaveSalida.Dump
		If Not IsEmpty (oAccess_GrupoEngrase) Then Wscript.Echo "Grupo de Engrase" : oAccess_GrupoEngrase.Dump
		Wscript.Echo "TOTAL Accesorios y motor: " & oCellAccess_TOTAL
		Wscript.Echo "Sobrecoste Accesorios y motor: " & oCellAccess_TOTALSobrecoste
	End Sub
End Class

Class cOferGas_XLS_Opciones
	Dim oOpciones_ValvRetenc, oOpciones_EVAgua, oOpciones_EngraseCilindros, oOpciones_Purgadores, oOpciones_ResistCalefacc, oOpciones_Bypass, oOpciones_ValvReguladora
	Dim oCellOpciones_TOTAL, oCellOpciones_TOTALSobrecoste
	Sub Dump
		If Not IsEmpty(oOpciones_ValvRetenc) then Wscript.Echo "ValvRetenc" : oOpciones_ValvRetenc.Dump
		If Not IsEmpty(oOpciones_EVAgua) then Wscript.Echo "EVAgua" : oOpciones_EVAgua.Dump
		If Not IsEmpty(oOpciones_EngraseCilindros) then Wscript.Echo "EngraseCilindros" : oOpciones_EngraseCilindros.Dump
		If Not IsEmpty(oOpciones_Purgadores) then Wscript.Echo "Purgadores" : oOpciones_Purgadores.Dump
		If Not IsEmpty(oOpciones_ResistCalefacc) then Wscript.Echo "ResistCalefacc" : oOpciones_ResistCalefacc.Dump
		If Not IsEmpty(oOpciones_Bypass) then Wscript.Echo "Bypass" : oOpciones_Bypass.Dump
		If Not IsEmpty(oOpciones_ValvReguladora) then Wscript.Echo "ValvReguladora" : oOpciones_ValvReguladora.Dump
		Wscript.Echo "TOTAL Opciones: " & oCellOpciones_TOTAL
		Wscript.Echo "Sobrecoste Opciones: " & oCellOpciones_TOTALSobrecoste
	End Sub
End Class

Class cOferGas_XLS_ManoObra
	Dim oCell_ManoObra_Predefined
	Dim oManoObraFase1, oManoObraFase2, oManoObraSoldadura, oManoObraProbadero, oManoObraPintura, oManoObraElectrica, oManoObraIngenieria, oManoObraExtra
	Dim oCellManoObra_TOTAL, oCellManoObra_TOTALSobrecoste
	Sub Dump
		Wscript.Echo "Mano de obra"
		If Not IsEmpty(oManoObraFase1) then Wscript.Echo "Fase1" : oManoObraFase1.Dump
		If Not IsEmpty(oManoObraFase2) then Wscript.Echo "Fase2" : oManoObraFase2.Dump
		If Not IsEmpty(oManoObraSoldadura) then Wscript.Echo "Soldadura" : oManoObraSoldadura.Dump
		If Not IsEmpty(oManoObraProbadero) then Wscript.Echo "Probadero" : oManoObraProbadero.Dump
		If Not IsEmpty(oManoObraPintura) then Wscript.Echo "Pintura" : oManoObraPintura.Dump
		If Not IsEmpty(oManoObraElectrica) then Wscript.Echo "Electrica" : oManoObraElectrica.Dump
		
		If Not IsEmpty(oManoObraIngenieria) then Wscript.Echo "Ingenieria" : oManoObraIngenieria.Dump
		If Not IsEmpty(oManoObraExtra) then Wscript.Echo "Extra" : oManoObraExtra.Dump
		Wscript.Echo "TOTAL Mano de obra: " & oCellManoObra_TOTAL
		Wscript.Echo "Sobrecoste Mano de obra: " & oCellManoObra_TOTALSobrecoste
	End Sub
End Class

Class cOferGas_XLS_Extras
	Dim oExtras_Embalaje,oExtras_Transporte,oExtras_PuestaEnMarcha,oCellExtras_TOTAL,oCellExtras_TOTALSobrecoste
	Sub Dump
		If Not IsEmpty (oExtras_Embalaje) Or Not IsEmpty (oExtras_Transporte) Or Not IsEmpty (oExtras_PuestaEnMarcha) then
			oExtras_Embalaje.Dump
			oExtras_Transporte.Dump
			oExtras_PuestaEnMarcha.Dump
			WScript.Echo "EXTRAS, Coste total: " & oCellExtras_TOTAL & vbTab & "Sobrecoste: " & oCellExtras_TOTALSobrecoste
		End if
	End Sub
End Class

Class cOferGas_XLS_Otros
	Dim oDicOtrosCostes,oCellOtrosCostes_TOTAL,oCellOtrosCostes_TOTALSobrecoste,bTest_Approach
	Private Sub Class_Initialize()
		Set oDicOtrosCostes = CreateObject("scripting.dictionary")
	End Sub
	Public Function AddCoste (oXlSheet, fila, colId, colNr,colCoste,colSobrecoste)
		Dim oItemCoste
		Set oItemCoste = (New ItemCoste).Init (oXlSheet, fila, colId, colNr,colCoste,colSobrecoste)
		oDicOtrosCostes.Add oDicOtrosCostes.Count, oItemCoste
		If (InStr(LCase(oItemcoste.strID),"test") > 0 Or InStr(LCase(oItemcoste.strID),"approach") > 0) And Not bTest_Approach Then 
			WScript.Echo "HAY QUE RESTAR DEL COSTE DE MATERIALES LAS OPCIONES DE 'test' O 'approach'!!! (los 'approach' son SUBCONTRATACION, y el helium test y extra test irían a 'MOD'... aunque en realidad NO IMPORTA: en los cálculos sucesivos esas partidas se suman entre sí...)"
			'MsgBox "OJO, HAY QUE RESTAR DEL COSTE DE MATERIALES LAS OPCIONES DE 'test' O 'approach'!!! (los 'approach' son SUBCONTRATACION, y el helium test y extra test irían a 'MOD')"
			bTest_Approach = true
		End If
		Set AddCoste = oItemCoste
	End Function
	Sub Dump
		Dim oItemCoste
		If oDicOtrosCostes.Count > 0 Then
			For Each oItemCoste In oDicOtrosCostes.Items
				oItemCoste.Dump
			Next
			WScript.Echo "OTROS COSTES, Coste total: " & oCellOtrosCostes_TOTAL & vbTab & "Sobrecoste: " & oCellOtrosCostes_TOTALSobrecoste
		End if
	End Sub
End Class

Class cOferGas_XLS_Ingenieria_Gestion_Estruct_Margen
	Dim oCellNumCompresores,oCellNumPlanos,oCellHorasPLano,oCellCosteSeguMaq,oCellHorasIngMaq
	Dim oCellMaterial, oCellMOD, oCellSubcontratac, oCellFactorMaterial, oCellPlusOnMaterial, oCellExpedicion, oCellDesignOG_NumHoras, oCellDesignOG, oCellFactorGestionOG, oCellGestionOG
	Dim oCellFactorGastosEstruct, oCellGastosEstruct, oCellFactorMargen, oCellMargen, oCellTOTAL, oCellFactorComisAgente, oCellTOTALconComisAgente
	Sub Dump
		If Not IsEmpty (oCellNumCompresores) Then Wscript.echo "Num Compresores: " & oCellNumCompresores
		If Not IsEmpty (oCellNumPlanos) Then Wscript.echo "Num de planos de máquina: " & oCellNumPlanos
		If Not IsEmpty (oCellHorasPLano) Then Wscript.echo "Horas por cada plano: " & oCellHorasPLano
		If Not IsEmpty (oCellCosteSeguMaq) Then Wscript.echo "Coste segu maquina: " & oCellCosteSeguMaq
		If Not IsEmpty (oCellHorasIngMaq) Then Wscript.echo "Horas de ingenieria por cada máquina: " & oCellHorasIngMaq
		Wscript.echo "Material: " & oCellMaterial
		WScript.echo "MOD: " & oCellMOD
		Wscript.echo "Subcontratac: " & oCellSubcontratac
		Wscript.echo "Factor Material: " & oCellFactorMaterial
		Wscript.echo "Plus On Material: " & oCellPlusOnMaterial
		Wscript.echo "Expedicion: " & oCellExpedicion
		Wscript.echo "DesignOG NumHoras: " & oCellDesignOG_NumHoras
		Wscript.echo "Design OG: " & oCellDesignOG
		Wscript.echo "Factor Gestion OG: " & oCellFactorGestionOG
		Wscript.echo "Gestion OG: " & oCellGestionOG
		Wscript.echo "Factor Gastos Estruct: " & oCellFactorGastosEstruct
		Wscript.echo "Gastos Estruct: " & oCellGastosEstruct
		Wscript.echo "Factor Margen: " & oCellFactorMargen
		Wscript.echo "Margen: " & oCellMargen
		Wscript.echo "TOTAL: " & oCellTOTAL
		Wscript.echo "Factor Comis Agente: " & oCellFactorComisAgente
		Wscript.echo "TOTAL con Comis Agente: " & oCellTOTALconComisAgente
	End Sub
End Class

Class ItemCoste
	' Propiedades genéricas
	Dim oCellId, oCellNr, strID
	Private colId_, colNr_,colCoste_,colSobrecoste_
	Dim oCellCoste, oCellSobrecoste
	Dim oDicProps
	Private Sub Class_Initialize
		Set oDicProps = CreateObject("scripting.dictionary")
	End Sub
	Function Init (oXlSheet, fila, colId, colNr,colCoste,colSobrecoste)
		Dim c
		Set Init = Me
		colId_ = colId
		colNr_ = colNr
		colCoste_ = colCoste
		colSobrecoste_ = colSobrecoste
		If Not IsEmpty (colCoste) Then Set oCellCoste = oXlSheet.range (colCoste & fila) Else Init = Empty : Exit Function : End If
		If Not IsEmpty (colSobrecoste) Then Set oCellSobrecoste = oXlSheet.range (colSobrecoste & fila)
		If Not IsEmpty (colId) Then
			For c = 1 To Len (colId)
				If IsEmpty (oCellId) Then
					Set oCellId = oXlSheet.range (Mid(colId,c,1) & fila)
					strID = oCellId.Value
				Else
					strID = strID & " - " & oXlSheet.range (Mid(colId,c,1) & fila).Value
				End if
			Next
		End if
		If Not IsEmpty (colNr) Then Set oCellNr = oXlSheet.range (colNr & fila)
	End Function
	Public Property Get bDefined ()
		If Not IsEmpty (oCellId) Then If oCellId.Value <> "" Then bDefined = True
	End Property
	Public Default Property Get prop (StrPropName)
		If oDicProps.Exists (StrPropName) Then Set prop = oDicProps(StrPropName)
	End Property
	Public Property Set prop (StrPropName,oPropCell)
		If Not oDicProps.Exists (StrPropName) Then
			oDicProps.Add StrPropName, oPropCell
		Else
			MsgBox ("Ya se había definido una celda con esa propiedad")
		End If
	End Property
	Sub Dump
		Dim StrPropName,strout
		If Not IsEmpty (oCellLabel) Then strout = "##" & oCellLabel & ": " Else strout = vbTab
		strout =  strout & strID
		If Not IsEmpty (oCellNr) Then strout =  strout & " (uds.:" & oCellNr & ")"
		WScript.Echo strout & vbCrLf & vbtab & "Coste: " & oCellCoste & vbtab & "Sobrecoste: " & oCellSobrecoste
		For Each StrPropName In oDicProps
			WScript.Echo vbTab & StrPropName & ": " & prop(StrPropName)
		Next
	End Sub
End Class
