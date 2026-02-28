Option Explicit

Function lanzaPrg (strPrgCmd, strTitleFindPattern, ByRef oExec, ByRef arrPosSize)
	strftmp = fso.GetSpecialFolder(2) & "\" & fso.GetTempName()
	Wshshell.Run "guipropview /stext """ & strftmp & """ /filter Title:""" & strTitleFindPattern & """",0,true
	regex.Pattern = "Handle\s+:\s*(\S+)(?:[\s\S](?!==))+?Title[^\r\n]+?" & Replace(strTitleFindPattern,"*",".*") & "(?:[\s\S](?!==))+?Position\s+:\s*(\d+),\s*(\d+)(?:[\s\S](?!==))+?Size\s+:\s*(\d+),\s*(\d+)[\s\S]+?={3,}"
	Set ExistingWND = Nothing
	If fso.GetFile (strftmp).Size > 4 Then Set ExistingWND = regex.Execute(fso.OpenTextFile (strftmp,1,0,-1).ReadAll)
	fso.DeleteFile strftmp,True
	Set oExec = Wshshell.Exec (strPrgCmd)
	Do
	'	Set oExec = WshShell.Exec ("guipropview /Action SwitchTo Title:""" & strTitleFindPattern & """")
	'	Do While oExec.Status = 0
	'		If Not oExec.StdOut.AtEndOfStream Then
	'			out = oExec.StdOut.Read
	'		End If
	'		WScript.Sleep 100
	'	Loop
		WScript.Sleep 1000
		Wshshell.Run "guipropview /stext """ & strftmp & """ /filter Title:""" & strTitleFindPattern & """",0,true
		regex.Pattern = "Handle\s+:\s*(\S+)(?:[\s\S](?!==))+?Title[^\r\n]+?" & Replace(strTitleFindPattern,"*",".*") & "(?:[\s\S](?!==))+?Position\s+:\s*(\d+),\s*(\d+)(?:[\s\S](?!==))+?Size\s+:\s*(\d+),\s*(\d+)[\s\S]+?={3,}"
		If fso.GetFile (strftmp).Size > 4 Then
			Set NewWND = regex.Execute(fso.OpenTextFile (strftmp,1,0,-1).ReadAll)
			For Each tmpWN In NewWND
				If ExistingWND is Nothing Then
					lanzaPrg = tmpWN.submatches(0)
					arrPosSize = Array (tmpWN.submatches(1),tmpWN.submatches(2),tmpWN.submatches(3),tmpWN.submatches(4))
				else
					For Each prevWN In ExistingWND
						If tmpWN.Value <> prevWN.Value Then
							lanzaPrg = tmpWN.submatches(0)
							arrPosSize = Array (tmpWN.submatches(1),tmpWN.submatches(2),tmpWN.submatches(3),tmpWN.submatches(4))
							Exit For
						End If
					Next
				End if
				If Not IsEmpty (lanzaPrg) Then Exit For
			Next
		End If
		fso.DeleteFile strftmp,True
	Loop While IsEmpty (lanzaPrg)
End Function

Function getAireGasControlIDs (strAireHND, oDicControls)
	strftmp = fso.GetSpecialFolder(2) & "\" & fso.GetTempName()
	Wshshell.Run "guipropview /stext """ & strftmp & """ /ParentWindow " & strAireHND ,0,True
	'Stop
	strControls = fso.OpenTextFile (strftmp,1,0,-1).ReadAll
	Set oDicControls = CreateObject("scripting.dictionary")
	For Each arrControl In Array (Array("Guardar",73,"GuardarEtapas"),Array("Calcular",82,"CalcularEtapas"))
		regex.Pattern = "Handle\s+:\s*(\S+)(?:[\s\S](?!==))+?Text[\s:]*(" & arrControl(0) & ")(?:[\s\S](?!==))+?Z-order\s+:\s*(" & arrControl(1) & ")(?:[\s\S](?!==))+?Position\s+:\s*(\d+),\s*(\d+)(?:[\s\S](?!==))+?Size\s+:\s*(\d+),\s*(\d+)(?:[\s\S](?!==))+?Class Atom\s+:\s*(\d+)[\s\S]+?={3,}"
		Set AireControls = regex.Execute(strControls)
		For Each Control In AireControls
			oDicControls.Add arrControl(2), Array (Control.submatches(0),Control.submatches(1),Control.submatches(2), _
					Control.submatches(3), Control.submatches(4),Control.submatches(5),Control.submatches(6), _
					Control.submatches(7)) ' Handle, Text, Z-order, Positionx, y , Sizex, y, Atom
		Next
	Next
	fso.DeleteFile strftmp,True
End Function

Function getAireGasControlCoords (oDicControls)
	' usar con start /wait guipropview /Action SwitchTo Handle:(Handle de AireGas) & nircmd setcursorwin  x y & nircmd sendmouse left click
	' para las coordenadas x, y, sumar +30, +40 respecto a las coords del control que da guipropview
	oDicControls.Add "TabCalcCompresor", Array (72, 153) ' (22, 113)
		oDicControls.Add "CompresorSerie", Array (88, 205) ' (58, 165)

		oDicControls.Add "SeleccCalc", Array (960, 101) ' (930, 61)
		oDicControls.Add "OpcCalc", Array (1077, 101) ' (1047, 61)
		oDicControls.Add "CalcularEtapas", Array (519, 443) ' (489, 403)
		oDicControls.Add "Errores", Array (995, 440) ' (965, 400)
		oDicControls.Add "TabGas", Array (519, 205) ' (489, 165)
		oDicControls.Add "TabCilindros", Array (569, 205) ' ??
		oDicControls.Add "TabRefrigeradores", Array (639, 205) ' ??
		oDicControls.Add "TabAireBaja", Array (729, 205) ' ??
		oDicControls.Add "TabRegulacionSets", Array (829, 205) ' ??

	oDicControls.Add "TabCalcAhorro", Array (182, 153) 
	oDicControls.Add "TabCalcVTPares", Array (232, 153) 
	oDicControls.Add "TabCalcPotCau", Array (287, 153) 
	oDicControls.Add "TabCalcAntipul", Array (342, 153) 
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' deberia tener distintas clases para distintos tipos de resultados:
' cABCGas_XLS (deberia cambiar de nombre) es PARA UN MODELO CONCRETO DE COMPRESOR, indep de los ficheros de resultados (varios ficheros de resultados, O VARIAS OPCIONES, podrían converger a esta clase)
' para los ficheros MULTI, habría una "CLASE DE COMPARACION DE MODELOS", con UN DICCIONARIO que almacecene cada calculo...
' habría que sacar las funciones comunes...
' (y para otros ficheros pueden aparecer otras clases...)
Class cABCGas_XLS
	Private regex
	Private strNACEcheckFPath, strCylsPreMatcheckFPath
	Dim strXLSXPath,bSave
	Dim oCellAutor, oCellFecha, oCellCalculo, oCellCliente, oCellProyecto, oCellObservaciones
	Dim oCell_Gas_Pcrit, oCell_Gas_Tcrit, oCell_Gas_Zcrit, oCell_Gas_Znorm, oCell_Gas_GammaNorm, oCell_Gas_GammaAsp, oCell_Gas_MW
	Dim oCell_Compressor_Serie, oCell_Compressor_Lubr, oCell_Compressor_Refrig, oCell_Compressor_Tipo, oCell_Compressor_Model
	Dim oCell_INProcess_Flow, oCell_INProcess_FlowAtRefHumid, oCell_INProcess_Pescape, oCell_INProcess_RPM, oCell_INProcess_FullENgine
	Dim oCell_Process_CompRatio_Global, oCell_Process_CompRatio_Mean, oCell_INProcess_Paspirac, oCell_INProcess_Patm, oCell_INProcess_Taspirac, oCell_INProcess_Tamb
	Dim oCell_INProcess_HRamb, oCell_INProcess_Taguarefrig, oCell_INProcess_RefrAceite, oCell_INProcess_DeltaTaguarefrigES, oCell_INProcess_DirtFact_Int, oCell_INProcess_DirtFact_Ext
	Dim oCell_OUTProcess_Flow, oCell_OUTProcess_RPM, oCell_OUTProcess_Paspirac, oCell_OUTProcess_Wabsorb, oCell_OUTProcess_Winst, oCell_OUTProcess_MechLoss, oCell_OUTProcess_NuIsoterm
	Dim oCell_OUTProcess_Pow_perUnitVol, oCell_OUTProcess_CoolingWaterFlow, oCell_OUTProcess_Condens, oCell_OUTProcess_Heat, oCell_OUTProcess_VentAir, oCell_OUTProcess_PistonMeanSpeed
	
	Dim oDicGasComp,bAire,bHumedo,bATEX_Inflamable,bSafeZone, bN2, bH2, bO2, bH2O, bCO, bCO2, bAR, bNH3, bSH2, bC2H4, bHCs
	Dim oDicStages,ncils
	Dim PotenciakW,bCompetitivosPorPotencia,RPM
	Dim strTipoCalculo ' "Rated, para seleccion de etapas y cilindros" | "A vueltas fijas, Design" (puede ser Pescape > PescapeRated (max, disparo valvulas), o Paspirac < PaspiracRated, ...: AQUI NO PUEDO DIFERENCIARLOS!!!
	Dim m_bAPI618
	
	Private m_ExcelApp
	Private m_ExcelFMFile  ' cExcelFile wrapper del archivo principal

	Private Sub Class_Initialize()
		On Error Resume Next
		strNACEcheckFPath = getResource("strNACEcheckFPath") 
		strCylsPreMatcheckFPath = getResource("strCylsPreMatcheckFPath")
		On Error GoTo 0
		If Not fso.fileExists (strNACEcheckFPath) Then _
				strNACEcheckFPath = "C:\abc compressors\INTRANET\OficinaTecnica\Documentacion\Normas\NACE\Herramienta\Herramienta_para_seleccion_de_materiales_v1.4.xlsx"
		If Not fso.fileExists (strCylsPreMatcheckFPath) Then _
				strCylsPreMatcheckFPath = "C:\abc compressors\INTRANET\OilGas\3_OFERTAS\ADJUNTOS OFERTAS\Datos cilindros 2.xlsx"
	    Set regex = New RegExp
		regex.Global = True : regex.IgnoreCase = True : regex.multiline = False
		Set oDicGasComp = CreateObject("scripting.dictionary")
		Set oDicStages = CreateObject("scripting.dictionary")
		Set m_ExcelApp = Nothing
		Set m_ExcelFMFile = Nothing
		Set oLimitsFeatsReqs_ = Nothing
	End Sub
	
	Private Sub Class_Terminate()
		'Stop : Call CloseWorkBook() ' NO DEBERIA NECESITAR ESTA LLAMADA, EN TANTO QUE USE EL EXCELMANAGER...
	End Sub
	
	Private Property Get objExcel
		If m_ExcelApp Is Nothing Then
			Err.Raise 91, "cABCGas", "ExcelApp not initialized. Call Init() first."
		End If
		Set objExcel = m_ExcelApp.Application
	End Property
	
	' =============================================
	' INICIALIZACIÓN 
	' =============================================
	
	Public Function Init(ExcelApp, strXLSXPath_, bAPI618)
		Set Init = Me
		strXLSXPath = strXLSXPath_
		Set m_ExcelApp = ExcelApp
		m_bAPI618 = bAPI618
		
		Set m_ExcelFMFile = m_ExcelApp.OpenFile(strXLSXPath, False, False)
		
		If getGASSheetInfo () Is Nothing Then
			'Set Init = Nothing
			Stop : Call CloseWorkBook ' pte de comprobar si lo hago aqui, o en el destructor...: SE CIERRA FUERA, O EN EL DESTRUCTOR, "Init" SIEMPRE dejara el fichero abierto
			Exit Function
		End If

		' Redirigir salida a DOC
		MsgIE.setContainer "doc"
		MsgIE ("Cliente: <b>" & oCellCliente.Value & "</b>")
		MsgIE ("Proyecto: <b>" & oCellProyecto.Value & "</b>")
		MsgIE ("Observaciones: <b>" & oCellObservaciones.Value & "</b>")
		MsgIE.popContainer
		
		' Inicializo las secciones de volcado de datos para ese calculo: 
		'SEGUIR AQUI

		' DESCARTANDO SOLUCIONES IMPOSIBLES:
		Call MsgIE.Spoiler (True,"background-color:orange;color:black;", "ADVERTENCIAS","id" & oCellCalculo.Value & "Advert",True)
		If ncils mod 2 <> 0 And ncils <> 1 Then
			MsgIE ("- NO SE HA PUESTO UN NUMERO PAR DE CILINDROS, no se puede fabricar como COMPRESOR HORIZONTAL")
			Set Init = Nothing
			'Exit Function
		End If
		If oCell_Compressor_Serie = "HG" Then
			If m_bAPI618 And bH2 Then
				MsgBox ("El compresor es API-618, para H2, HACERLO EN PLATAFORMA HP")
			ElseIf m_bAPI618 Then
				MsgBox ("El compresor es API-618, CONVIENE HACERLO EN PLATAFORMA HP")
			ElseIf bH2 Then
				MsgBox ("Muy posiblemente el compresor es API-618 (H2), CONVIENE HACERLO EN PLATAFORMA HP")
			End If
			MsgIE ("SI EL COMPRESOR ES API-618 (H2, gas natural, etc, o por especificación de cliente). CONVIENE HACERLO <b>MEJOR EN PLATAFORMA HP!!!</b>.")
			'Set Init = Nothing
			'Exit Function
		End If
		If oCell_Compressor_Serie & "-" & ncils = "HG-6" Then
			MsgIE ("ESTE MODELO, " & "<b>HG-6</b>" & ", NO SE FABRICA, debería calcularse como un HP4.")
			Set Init = Nothing
			'Exit Function
		End If
		If Not bCompetitivosPorPotencia Then
			MsgIE ("NO SOMOS COMPETITIVOS en las condiciones de calculo de '" & fso.GetBaseName(strXLSXPath) & "' el compresor es de MUY baja potencia: " & PotenciakW & _
				" kW. *** CONVENDRIA CONSIDERAR LAS PLATAFORMAS ""V"" Y/O ""X"", que son DE SIMPLE EFECTO, aunque más proclives a FUGAS, se está estudiando añadirles un sistema de RECUPERACION DE FUGAS... Si no, DECLINAR OFERTA")
			'Set Init = Nothing
			'Exit Function
		End If
		If ncils = 1 And oDicStages.Count < 2 Then MsgIE vbtab & "el compresor es MANCO, ojo a requisitos"
		If InStr (oCell_Compressor_Serie,"HX") > 0 And ncils > 2 Then _
				MsgIE vbtab & "NO SE HA HECHO NINGUN COMPRESOR " & oCell_Compressor_Serie & "-" & ncils & ", en HX sólo se ha hecho un HX2, para REPSOL..."
		If bAire Then MsgIE vbtab & "Asegurarse de que el compresor " & strModelName & ", que es de aire, NO SE PUEDA OFERTAR COMO PLATAFORMA LP, de máquina estándar (llegan hasta 1000 rpm), o como SYNCRO."
		MsgIE.popContainer ' "id" & oCellCalculo.Value & "Advert"
		'Stop ' puede que aqui tb se justifique CERRAR EL WORKBOOK
		
		' Generar resumen en MAIN
		Call ToUISummary()
	End Function
	
	Public Sub ToUISummary()
		MsgIE.setContainer "main"
		' Placeholder para el resumen en el panel principal
		MsgIE.MsgMain "info", "<b>" & strModelName & "</b> (" & strTipoCalculo & ") - " & PotenciakW & " kW"
		MsgIE.popContainer
	End Sub
	
	Public Function CloseWorkBook()
		If Not (m_ExcelFMFile Is Nothing) Then
			If bSave Then m_ExcelFMFile.Save
			m_ExcelApp.CloseFile m_ExcelFMFile.FilePath, False
			Set m_ExcelFMFile = Nothing
		End If
	End Function
		
	Private Function CellValue (oCell) ' extrae el valor numerico de una celda
		regex.Pattern = "^([\-,\.\d]+)\s+(.+)$"
		CellValue = CDbl(regex.Execute (oCell.Value).Item(0).Submatches(0))
	End Function
	Private Function CellUnits (oCell) ' extrae las unidades del valor de una celda
		regex.Pattern = "^([\-,\.\d]+)\s+(.+)$"
		CellUnits = regex.Execute (oCell.Value).Item(0).Submatches(1)
	End Function
	
	Private Function getGasName (Cell)
		getGasName = Trim(Replace(Cell.Value,":",""))
	End Function

	Private Function getGasComp (Cell)
		getGasComp = Replace(oDicGasComp(Cell)(0).Value,"%","") * 1
	End Function
	
	Public Property Get strCalcOpc ()
		If Not IsEmpty (oCellCalculo) Then
			strCalcOpc = oCellCalculo.Value
		End if
	End Property

	Private oGASXSLSheet_	
	Private Function getGASSheetInfo ()
		' OBTIENE INFORMACION DE LA HOJA 'GAS' DEL FICHERO DE EXCEL
		
		If Not IsEmpty (oGASXSLSheet_) Then
			Set getGASSheetInfo =  oGASXSLSheet_
			Exit Function ' SOLO SE PROCESA UNA VEZ esta función: la info que lee NO CAMBIA, --> NO tiene sentido hacerlo más veces
		End If
		If m_ExcelFMFile Is Nothing Then
			Set getGASSheetInfo = Nothing
			Exit Function
		End If
		If Not m_ExcelFMFile.HasWorksheet("GAS") Then
			Set getGASSheetInfo = Nothing
			Exit Function
		End If
		Dim oXlSheet
		Set oXlSheet = m_ExcelFMFile.GetWorksheet("GAS")
		
		' genero referencias para todos los valores de la hoja
		Set oCellAutor = oXlSheet.range ("H5")
		Set oCellFecha = oXlSheet.range ("H6")
		Dim c,oCellGasName,oCellGasPC,oCellGasDesc
		c = 10
		Set oCellCalculo = oXlSheet.range ("B" & c) : c = c + 1
		Set oCellCliente = oXlSheet.range ("B" & c) : c = c + 1
		Set oCellProyecto = oXlSheet.range ("B" & c) : c = c + 1
		Set oCellObservaciones = oXlSheet.range ("B" & c) : c = c + 1

		c = 20
		Set oCell_Gas_Pcrit = oXlSheet.range ("B" & c) : c = c + 1 ' HAY QUE IDENTIFICAR LAS UNIDADES!!!
		Set oCell_Gas_Tcrit = oXlSheet.range ("B" & c) : c = c + 1 ' HAY QUE IDENTIFICAR LAS UNIDADES!!!
		Set oCell_Gas_Zcrit = oXlSheet.range ("B" & c) : c = c + 1
		Set oCell_Gas_Znorm = oXlSheet.range ("B" & c) : c = c + 1
		' Gamma = peso especifico
		Set oCell_Gas_GammaNorm = oXlSheet.range ("B" & c) : c = c + 1 ' HAY QUE IDENTIFICAR LAS UNIDADES!!!
		Set oCell_Gas_GammaAsp = oXlSheet.range ("B" & c) : c = c + 1 ' HAY QUE IDENTIFICAR LAS UNIDADES!!!
		Set oCell_Gas_MW = oXlSheet.range ("B" & c) : c = c + 1
		
		If CDbl(oCell_Gas_MW.Value) < 12 Then bATEX_Inflamable = True
		
		Dim strGasName, gasComp
		For c = 19 To 28
			If oXlSheet.range ("F" & c).Value = "" Then Exit For
			Set oCellGasName = oXlSheet.range ("F" & c)
			Set oCellGasPC = oXlSheet.range ("G" & c)
			Set oCellGasDesc = oXlSheet.range ("H" & c)
			oDicGasComp.Add oCellGasName, Array (oCellGasPC, oCellGasDesc)
			strGasName = getGasName(oCellGasName)
			gasComp = getGasComp(oCellGasName)
			
			If InStr(strGasName,"Air") > 0 Then bAire = True
			If strGasName="H2O" And gasComp > 1 Then bHumedo = True
			
			regex.Pattern = "C\d*H\d*"
			If (strGasName="H2" Or regex.Test (strGasName)) And gasComp > 1 Then bSafeZone = True
			
			bN2 = bN2 Or (strGasName="N2")
			bH2 = bH2 Or (strGasName="H2")
			bO2 = bO2 Or (strGasName="O2")
			bH2O = bH2O Or (strGasName="H2O")
			bCO = bCO Or (strGasName="CO")
			bCO2 = bCO2 Or (strGasName="CO2")
			bAR = bAR Or (strGasName="AR")
			bNH3 = bNH3 Or (strGasName="NH3")
			bSH2 = bSH2 Or (strGasName="SH2")
			bC2H4 = bC2H4 Or (strGasName="C2H4")
			Select Case strGasName
				Case "CH4","C2H6","C3H8","C4H10","C5H12"
					bHCs = True
			End Select
		Next
		c = 30
		Set oCell_Compressor_Serie = oXlSheet.range ("B" & c) : c = c + 1
		Set oCell_Compressor_Lubr = oXlSheet.range ("B" & c) : c = c + 1
		Set oCell_Compressor_Refrig = oXlSheet.range ("B" & c) : c = c + 1
		Set oCell_Compressor_Tipo = oXlSheet.range ("B" & c) : c = c + 1
		Set oCell_Compressor_Model = oXlSheet.range ("B" & c) : c = c + 1
		Set oCell_INProcess_Flow = oXlSheet.range ("B" & c) : c = c + 1 ' HAY QUE IDENTIFICAR LAS UNIDADES!!!
		Set oCell_INProcess_FlowAtRefHumid = oXlSheet.range ("B" & c) : c = c + 1
		Set oCell_INProcess_Pescape = oXlSheet.range ("B" & c) : c = c + 1 ' HAY QUE IDENTIFICAR LAS UNIDADES!!!
		Set oCell_INProcess_RPM = oXlSheet.range ("B" & c) : c = c + 1 ' FIJARSE QUE NOS ESTÁ DANDO EL LIMITE DE VELOCIDAD!!!
		Set oCell_INProcess_FullEngine = oXlSheet.range ("B" & c) : c = c + 1 ' ES UN COMENTARIO...
		Set oCell_Process_CompRatio_Global = oXlSheet.range ("B" & c) : c = c + 1
		Set oCell_Process_CompRatio_Mean = oXlSheet.range ("B" & c) : c = c + 1
		
		c = 30
		Set oCell_INProcess_Paspirac = oXlSheet.range ("I" & c) : c = c + 1 ' HAY QUE IDENTIFICAR LAS UNIDADES!!!
		Set oCell_INProcess_Patm = oXlSheet.range ("I" & c) : c = c + 1 ' HAY QUE IDENTIFICAR LAS UNIDADES!!!
		Set oCell_INProcess_Taspirac = oXlSheet.range ("I" & c) : c = c + 1 ' HAY QUE IDENTIFICAR LAS UNIDADES!!!
		Set oCell_INProcess_Tamb = oXlSheet.range ("I" & c) : c = c + 1 ' HAY QUE IDENTIFICAR LAS UNIDADES!!!
		Set oCell_INProcess_HRamb = oXlSheet.range ("I" & c) : c = c + 1
		Set oCell_INProcess_Taguarefrig = oXlSheet.range ("I" & c) : c = c + 1 ' HAY QUE IDENTIFICAR LAS UNIDADES!!!

		c = 38
		Set oCell_INProcess_RefrAceite = oXlSheet.range ("I" & c) : c = c + 1
		Set oCell_INProcess_DeltaTaguarefrigES = oXlSheet.range ("I" & c) : c = c + 1
		Set oCell_INProcess_DirtFact_Int = oXlSheet.range ("I" & c) : c = c + 1
		Set oCell_INProcess_DirtFact_Ext = oXlSheet.range ("I" & c) : c = c + 1
		
		c = 46
		Set oCell_OUTProcess_Flow = oXlSheet.range ("B" & c) : c = c + 1 ' HAY QUE IDENTIFICAR LAS UNIDADES!!!
		Set oCell_OUTProcess_RPM = oXlSheet.range ("B" & c) : c = c + 1 ' (APARECE EN FORMATO / num (num): QUE SON???
		Set oCell_OUTProcess_Paspirac = oXlSheet.range ("B" & c) : c = c + 1 ' HAY QUE IDENTIFICAR LAS UNIDADES!!!
		Set oCell_OUTProcess_Wabsorb = oXlSheet.range ("B" & c) : c = c + 1 ' HAY QUE IDENTIFICAR LAS UNIDADES!!! y valores en unidades alternativas
		Set oCell_OUTProcess_Winst = oXlSheet.range ("B" & c) : c = c + 1 ' HAY QUE IDENTIFICAR LAS UNIDADES!!! y valores en unidades alternativas
		Set oCell_OUTProcess_MechLoss = oXlSheet.range ("B" & c) : c = c + 1 ' HAY QUE IDENTIFICAR LAS UNIDADES!!!
		' Rendimiento = Nu
		Set oCell_OUTProcess_NuIsoterm = oXlSheet.range ("B" & c) : c = c + 1
		Set oCell_OUTProcess_Pow_perUnitVol = oXlSheet.range ("B" & c) : c = c + 1 ' HAY QUE IDENTIFICAR LAS UNIDADES!!!
		Set oCell_OUTProcess_CoolingWaterFlow = oXlSheet.range ("B" & c) : c = c + 1 ' HAY QUE IDENTIFICAR LAS UNIDADES!!!
		Set oCell_OUTProcess_Condens = oXlSheet.range ("B" & c) : c = c + 1 ' HAY QUE IDENTIFICAR LAS UNIDADES!!!
		Set oCell_OUTProcess_Heat = oXlSheet.range ("B" & c) : c = c + 1 ' HAY QUE IDENTIFICAR LAS UNIDADES!!!
		Set oCell_OUTProcess_VentAir = oXlSheet.range ("B" & c) : c = c + 1 ' HAY QUE IDENTIFICAR LAS UNIDADES!!!
		Set oCell_OUTProcess_PistonMeanSpeed = oXlSheet.range ("B" & c) : c = c + 1 ' HAY QUE IDENTIFICAR LAS UNIDADES!!!
		
		regex.Pattern = "^([\-,\.\d]+)\s*/\s*([\-,\.\d]+)\s+([^/]+)(?:\s*/\s*(.+))$" '251,81 / 185,20  CV/KW
		'PotenciakW = regex.Execute(Replace(oCell_OUTProcess_Winst,",",".")).Item(0).submatches(1)
		PotenciakW = CDbl(regex.Execute(oCell_OUTProcess_Winst).Item(0).submatches(1))
		' para potencias de menos de 25 kW, NO SOMOS COMPETITIVOS...
		bCompetitivosPorPotencia = PotenciakW > 25
		regex.Pattern = "^([\-,\.\d]+)(?:\s*/\s*([\-,\.\d]+)\s+\(\s+([\-,\.\d]+)\s+\))?\s*$" '251,81 / 0 (0)
		'RPM = regex.Execute(Replace(oCell_OUTProcess_RPM,",",".")).Item(0).submatches(0)
		RPM = CDbl(regex.Execute(oCell_OUTProcess_RPM).Item(0).submatches(0))
		
		Dim i,d,oABCGas_XLS_Stage
		For i = asc("B") To Asc("G") ' PARA CADA ETAPA:
			d = Chr (i)
			c = 63
			If Trim(oXlSheet.range (d & c).Value) = "" Then Exit For ' COMPRUEBO SI LA CELDA DE ETAPA NO ESTÁ EN BLANCO -> GENERA UN CILINDRO
			Set oABCGas_XLS_Stage = New cABCGas_XLS_Stage
			With oABCGas_XLS_Stage
				Set .oCell_Stage_CilsDiam = oXlSheet.range (d & c) : c = c + 1
				Set .oCell_Stage_Flow = oXlSheet.range (d & c) : c = c + 1
				Set .oCell_Stage_Pout = oXlSheet.range (d & c) : c = c + 1
				Set .oCell_Stage_Taspirac = oXlSheet.range (d & c) : c = c + 1
				Set .oCell_Stage_Tescape = oXlSheet.range (d & c) : c = c + 1
				Set .oCell_Stage_TescapeAdiab = oXlSheet.range (d & c) : c = c + 1
				Set .oCell_Stage_CompRatio = oXlSheet.range (d & c) : c = c + 1
				Set .oCell_Stage_ComprStress = oXlSheet.range (d & c) : c = c + 1
				Set .oCell_Stage_TensStress = oXlSheet.range (d & c) : c = c + 1
				Set .oCell_Stage_VolumeGen = oXlSheet.range (d & c) : c = c + 1
				Set .oCell_Stage_NuVolum = oXlSheet.range (d & c) : c = c + 1
				Set .oCell_Stage_FillCoef = oXlSheet.range (d & c) : c = c + 1
				Set .oCell_Stage_DeadVolume = oXlSheet.range (d & c) : c = c + 1
				Set .oCell_Stage_MinVolum = oXlSheet.range (d & c) : c = c + 1
				Set .oCell_Stage_ValveSection = oXlSheet.range (d & c) : c = c + 1
				Set .oCell_Stage_ValveSpeed = oXlSheet.range (d & c) : c = c + 1
				Set .oCell_Stage_NuValve = oXlSheet.range (d & c) : c = c + 1
				Set .oCell_Stage_GammaAdiabIdx = oXlSheet.range (d & c) : c = c + 1
				Set .oCell_Stage_DiagrPower = oXlSheet.range (d & c) : c = c + 1
				Set .oCell_Stage_Regulation = oXlSheet.range (d & c) : c = c + 1
				Set .oCell_Stage_CoolingWater = oXlSheet.range (d & c) : c = c + 1
				
				c = 86
				Set .oCell_Stage_NrCoolers = oXlSheet.range (d & c) : c = c + 1
				Set .oCell_Stage_CoolerSize = oXlSheet.range (d & c) : c = c + 1 ' ES UN TEXTO...
				Set .oCell_Stage_TLR = oXlSheet.range (d & c) : c = c + 1
				Set .oCell_Stage_Pdrop = oXlSheet.range (d & c) : c = c + 1
				Set .oCell_Stage_WaterFlow = oXlSheet.range (d & c) : c = c + 1
				Set .oCell_Stage_CondensateWaterFlow = oXlSheet.range (d & c) : c = c + 1

				c = 93
				Set .oCell_Stage_Gas_Zin = oXlSheet.range (d & c) : c = c + 1
				Set .oCell_Stage_Gas_Zout = oXlSheet.range (d & c) : c = c + 1

				'ncils = ncils + Split (.oCell_Stage_CilsDiam," x ")(0) * 1
				' en el excel se reflejan LAS ETAPAS, no los cilindros: hay DOS etapas en 1 tandem --> 0.5 cils por etapa tandem
				If InStr (UCase(.oCell_Stage_CilsDiam),"T") > 0 Then
					ncils = ncils + 0.5 * .numCils
				Else
					ncils = ncils + .numCils * 1
				End If
			End With
			oDicStages.Add oDicStages.Count, oABCGas_XLS_Stage
		Next
		Set getGASSheetInfo = oXlSheet
		Set oGASXSLSheet_ =  getGASSheetInfo
	End Function
	
	Private strGasType_
	Function strGasType
		If Not IsEmpty (strGasType_) Then strGasType = strGasType_ : Exit Function
		Dim oCellGasName,strGasName
		Dim bPyroGas,bOffGas,bBioGas,bGN,bSynGas,bCoalGas,bCoke,bLPG,bHrich,bHeavyOil, bAcid, bNitrous,bCO2Low,bPlastics_LowTemp,bWet
		Dim bReforming_AirBlown,bCathalHidrogen
		bPyroGas = True : bBioGas = True : bGN = True : bSynGas = True : bCoalGas = True : bCoke = True : bLPG = True : bHrich = True : bHeavyOil = True : bOffGas = True
		bWet = False : bAcid = False : bNitrous = False : bCO2Low = False : bPlastics_LowTemp = False
		For Each oCellGasName In oDicGasComp
			strGasName = getGasName(oCellGasName)
			If InStr(strGasName,"Air") > 0 Then strGasType = "Air"
			If getGasComp(oCellGasName) > 95 Then strGasType = strGasName
			' el syngas, tb se conoce como coal gas, o coke
			'  los gases de combustión de carbón contienen SOx y NOx , mientras que los gases de combustión de gas natural normalmente solo contienen NOx .
			Select Case strGasName
				Case "C3H8"
					bGN = bGN And getGasComp(oCellGasName) < 25
					bSynGas = bSynGas And getGasComp(oCellGasName) < 10
					bLPG = bLPG And getGasComp(oCellGasName) > 30 ' ES EL COMPONENTE CARACTERISTICO DEL LPG
					bHeavyOil = bHeavyOil And getGasComp(oCellGasName) < 5
					bPyroGas = bPyroGas And getGasComp(oCellGasName) < 10
				Case "C2H6"
					bGN = bGN And getGasComp(oCellGasName) > 1 And getGasComp(oCellGasName) < 25
					bSynGas = bSynGas And getGasComp(oCellGasName) < 10
					bLPG = bLPG And getGasComp(oCellGasName) > 30 ' ES EL COMPONENTE CARACTERISTICO DEL LPG
					bHeavyOil = bHeavyOil And getGasComp(oCellGasName) < 8
					bPyroGas = bPyroGas And getGasComp(oCellGasName) > 1 And getGasComp(oCellGasName) < 20 ' ; ES EL COMPONENTE MAS REPRESENTATIVO DE LA PIROLISIS; RARO QUE ALCANCE EL 20%, MAS BIEN DEBERIA SER < 10%...
					bPlastics_LowTemp = getGasComp(oCellGasName) > 10 '; ES EL % CARACTERISTICO PARA EL PROCESADO PIROLITICO DE CIERTOS PLASTICOS (PE) A BAJA TEMPERATURA
				Case "CH4"
					bBioGas = bBioGas And getGasComp(oCellGasName) > 50 And getGasComp(oCellGasName) < 75 ' ES EL COMPONENTE CARACTERISTICO
					bGN = bGN And getGasComp(oCellGasName) > 75
					bSynGas = bSynGas And getGasComp(oCellGasName) < 10
					bCoalGas = bCoalGas And getGasComp(oCellGasName) < 15 ' EN ALGUN SITIO PONE Q < 2%
					bCoke = bCoke And getGasComp(oCellGasName) < 35 ' generalmente incluso un 35%
					bLPG = bLPG And getGasComp(oCellGasName) < 20
					bHeavyOil = bHeavyOil And getGasComp(oCellGasName) < 35 And getGasComp(oCellGasName) > 25
					bPyroGas = bPyroGas And getGasComp(oCellGasName) > 20 And getGasComp(oCellGasName) < 60 ' LO NORMAL ES QUE SEA < 45, SI ES SUPERIOR, PODRÍA CONSIDERARSE "SINTETICO"...
					bOffGas = bOffGas And getGasComp(oCellGasName) < 10
				Case "CO2"
					bBioGas = bBioGas And getGasComp(oCellGasName) > 25 And getGasComp(oCellGasName) < 45 ' ES EL COMPONENTE CARACTERISTICO
					bGN = bGN And getGasComp(oCellGasName) < 20
					'bSynGas = bSynGas And getGasComp(oCellGasName) > 5 And getGasComp(oCellGasName) < 15
					bSynGas = bSynGas And getGasComp(oCellGasName) < 25 And getGasComp(oCellGasName) > 5 ' en gral alrededor del 15%, 
					bLPG = bLPG And getGasComp(oCellGasName) < 5
					bHeavyOil = bHeavyOil And getGasComp(oCellGasName) > 50
					bOffGas = bOffGas And getGasComp(oCellGasName) > 30 And getGasComp(oCellGasName) < 75
					bCO2Low = bCO2Low And getGasComp(oCellGasName) < 5
					bAcid = getGasComp(oCellGasName) > 2.5 ' generalmente un 2% o mas
				Case "H2"
					bBioGas = bBioGas And getGasComp(oCellGasName) < 5 
					bGN = bGN And getGasComp(oCellGasName) < 20
					bSynGas = bSynGas And getGasComp(oCellGasName) > 25 ' generalmente incluso un 25%, o 30%
					bHrich = getGasComp(oCellGasName) > 65
					bReforming_AirBlown = getGasComp(oCellGasName) < 40
					bCoalGas = bCoalGas And getGasComp(oCellGasName) > 40 ' And getGasComp(oCellGasName) < 50 
					bCoke = bCoke And getGasComp(oCellGasName) > 40 ' generalmente incluso un 45%
					bLPG = bLPG And getGasComp(oCellGasName) < 10
					bPyroGas = bPyroGas And getGasComp(oCellGasName) < 40 And getGasComp(oCellGasName) > 20
					bOffGas = bOffGas And getGasComp(oCellGasName) > 5 And getGasComp(oCellGasName) < 40
				Case "CO"
					bGN = bGN And getGasComp(oCellGasName) < 1
					'bSynGas = bSynGas And getGasComp(oCellGasName) > 10 
					bSynGas = bSynGas And getGasComp(oCellGasName) < 60 And getGasComp(oCellGasName) > 15 ' generalmente ALRED DE UN 40%!! incluso un 20%, o 30 a 60%
					bCoalGas = bCoalGas And getGasComp(oCellGasName) < 12 ' generalmente un 8%
					bCoke = bCoke And getGasComp(oCellGasName) < 10 ' generalmente un 8%
					bPyroGas = bPyroGas And getGasComp(oCellGasName) < 10
					bOffGas = bOffGas And getGasComp(oCellGasName) < 10
				Case "N2"
					bCoalGas = bCoalGas And getGasComp(oCellGasName) > 1 And getGasComp(oCellGasName) < 15 ' generalmente un 5-15%
					bCoke = bCoke And getGasComp(oCellGasName) > 1 And getGasComp(oCellGasName) < 10 ' generalmente un 5-8%
					bReforming_AirBlown = getGasComp(oCellGasName) < 10
					bPyroGas = bPyroGas And getGasComp(oCellGasName) < 15
				Case "H2O"
					bWet = True
				Case "CH4O"
					bCathalHidrogen = True
				Case "SH2","SO2","HCl"
					bAcid = getGasComp(oCellGasName) > 0.001 ' generalmente un 0.0004%; 1% = 10000 ppm
				Case "NO","NO2","N2O"
					bNitrous = getGasComp(oCellGasName) > 0.00001 ' NO2 es el más agresivo, corrosivo (y toxico)
			End Select
		Next
		'Stop
		If bBioGas + bGN + bSynGas + bCoalGas + bCoke + bLPG + bHeavyOil + bPyroGas = -1 Or bOffGas Then
			Select Case True
				Case strGasType <> ""
					
				Case bSynGas
					strGasType = "SYNGAS"
					If bOffGas Then
						If InStr (LCase(oCellProyecto.Value),"off-gas") >0 Or InStr (LCase(oCellObservaciones.Value),"off-gas") >0 _
								Or InStr (LCase(oCellProyecto.Value),"offgas") >0 Or InStr (LCase(oCellObservaciones.Value),"offgas") >0 _
								Or InStr (LCase(oCellProyecto.Value),"recircul") >0 Or InStr (LCase(oCellObservaciones.Value),"recircul") >0 _
								Or InStr (LCase(oCellProyecto.Value),"recycle") >0 Or InStr (LCase(oCellObservaciones.Value),"recycle") >0 _
								Or InStr (LCase(oCellProyecto.Value),"flash") >0 Or InStr (LCase(oCellObservaciones.Value),"flash") >0 _
								Then
							strGasType = "OFF-GAS"
						ElseIf MsgBox ("Es compresor PRINCIPAL (SYNGAS; SI) O DE RECIRCULACION (OFF-GAS; NO)?",4+32) <> 6 Then
							strGasType = "OFF-GAS"
						End If
					End If
					If bHrich Then strGasType = strGasType & ", H2 RICH" End If
					If strGasType = "SYNGAS" then
						If bReforming_AirBlown Then strGasType = strGasType & " (REFORMING)" Else strGasType = strGasType & " (AIR BLOWN)" End If
						If bCathalHidrogen Then
							If bReforming_AirBlown Then
								strGasType = Replace (strGasType,"REFORMING","REFORMING, CATHALYTIC HYDROGENATION CO-CO2")
							Else
							End if
						End if
					End if
				Case bBioGas : strGasType = "BIOGAS"
				Case bHeavyOil : strGasType = "Heavy Oil Associated (EOR) / Unconv. Reservoir"
				Case bCoalGas : strGasType = "COAL GAS"
				Case bLPG : strGasType = "LPG"
				Case bGN : strGasType = "NATURAL GAS"
				Case bPyroGas
					strGasType = "PYROGAS"
					If bOffGas Then
						If InStr (LCase(oCellProyecto.Value),"off-gas") >0 Or InStr (LCase(oCellObservaciones.Value),"off-gas") >0 _
								Or InStr (LCase(oCellProyecto.Value),"offgas") >0 Or InStr (LCase(oCellObservaciones.Value),"offgas") >0 _
								Or InStr (LCase(oCellProyecto.Value),"recircul") >0 Or InStr (LCase(oCellObservaciones.Value),"recircul") >0 _
								Or InStr (LCase(oCellProyecto.Value),"recycle") >0 Or InStr (LCase(oCellObservaciones.Value),"recycle") >0 _
								Or InStr (LCase(oCellProyecto.Value),"flash") >0 Or InStr (LCase(oCellObservaciones.Value),"flash") >0 _
								Then
							strGasType = "OFF-GAS"
						ElseIf MsgBox ("Es compresor PRINCIPAL (PYROGAS; SI) O DE RECIRCULACION (OFF-GAS; NO)?",4+32) <> 6 Then
							strGasType = "OFF-GAS"
						End If
					End If
					If bCO2Low Then strGasType = strGasType & " (condensed CO2)"
					If bPlastics_LowTemp Then strGasType = strGasType & " (plastics, low residence, low temp)"
				Case bOffGas
					strGasType = "OFF-GAS"
			End Select	
			If bWet Then strGasType = strGasType & " (wet)"
			If bAcid Then strGasType = strGasType & " (acid gas)"
			If strGasType <> "" Then MsgLog ("El gas a procesar se puede considerar un " & strGasType)
		ElseIf strGasType = "" Then
			' no se sabe
			MsgIE "No se ha identificado el tipo de gas según composicion. (PTE perfeccionar funcion strGasType)"
			strGasType = "gas comp. acc. to data sheet"
		End If
		strGasType_ = strGasType
	End Function
	
	Private bNACE_Corrosivo_
	Public Function bNACE_Corrosivo
		If Not IsEmpty (bNACE_Corrosivo_) Then bNACE_Corrosivo = bNACE_Corrosivo_ : Exit Function
		' sI ES nace HAY QUE PONER: 1. bloque SAS; 2. Panel de N2; 3. Empaquetaduras en AISI 316; 4. Segmentos especiales (T92 / CO2); 5. Vástago con recubrim WC; 6. Calderería en INOX!!
		If fso.FileExists (strNACEcheckFPath) Then
			Dim regex
			Set regex = New RegExp
			regex.Global = True : regex.IgnoreCase = True : regex.multiline = False
			Dim oXlSheet, c, oCellGasName,bExcelScreenUpdating
			
			On Error Resume Next ' Si hay errores en Excel, que continúe
			Dim naceFMFile
			Set naceFMFile = m_ExcelApp.OpenFile(strNACEcheckFPath, True, True)
			If Not (naceFMFile Is Nothing) Then
				bExcelScreenUpdating = objExcel.ScreenUpdating
				If bQuickExcel Then objExcel.ScreenUpdating = False
			
				Set oXlSheet = naceFMFile.Workbook.worksheets("Presión parcial")
				' P absoluta (max), = Psalida
				oXlSheet.range ("D1").Value = oDicStages.Item(oDicStages.Count-1).Stage_Pout / 10
				' Temp gas (max), salida 1A ETAPA!!!
				oXlSheet.range ("F1").Value = oDicStages.Item(0).oCell_Stage_Tescape
				For c = 4 To 15
					oXlSheet.range ("D" & c).Value = 0
					For Each oCellGasName In oDicGasComp
						regex.Pattern ="\b" & Replace(getGasName(oCellGasName),"SH2","H2S") & "$"
						'If regex.Test(oXlSheet.range ("C" & c).Value) Then oXlSheet.range ("D" & c).Value = Replace(getGasComp(oCellGasName),".",",") : Exit For
						If regex.Test(oXlSheet.range ("C" & c).Value) Then oXlSheet.range ("D" & c).Value = CDbl(getGasComp(oCellGasName)) : Exit For
					Next
				Next
				' Ya tenemos el valor de salida:
				MsgIE ("NACE de los gases procesados: " & oXlSheet.range ("J22").Value)
				bNACE_Corrosivo = (oXlSheet.range ("J22").Value > 0)
				m_ExcelApp.CloseFile naceFMFile.FilePath, False
				If bQuickExcel Then objExcel.ScreenUpdating = bExcelScreenUpdating
			End If
			If Err Then Call MsgLog ("No se ha podido verificar NACE!")
			On Error Goto 0
		End If
		Dim PParc, strGasName
		For Each oCellGasName In oDicGasComp
			strGasName = getGasName(oCellGasName)
			If (strGasName="CO" Or strGasName="SH2") And getGasComp(oCellGasName) > 2 Then bNACE_Corrosivo = True : Exit For
			If strGasName="CO2" And getGasComp(oCellGasName) >= 2 Then
				PParc = oDicStages.Item(oDicStages.Count-1).Stage_Pout * getGasComp(oCellGasName) / 100
				If PParc > 1.5 And getGasComp(oCellGasName) > 5 Then bNACE_Corrosivo = True : Exit For
			End if
		Next
		bNACE_Corrosivo_ = bNACE_Corrosivo
	End Function
	
	Private bCylinderMaterialsProcessed_
	Function getCylindersMaterials_Limits()
		If Not fso.fileExists (strCylsPreMatcheckFPath) Or bCylinderMaterialsProcessed_ Then Exit Function

		Dim strfilterRng
		' Presiones POR CILINDRO, desde "C:\abc compressors\INTRANET\OilGas\3_OFERTAS\ADJUNTOS OFERTAS\Datos cilindros 2.xlsx"
		Select Case oCell_Compressor_Serie
			Case "HA": strfilterRng = "A2:G74"
			Case "HG", "HP": strfilterRng = "K5:Q52"
			Case Else : strfilterRng = ""
		End Select
		If strfilterRng = "" Then Exit Function ' El Excel no da información para esta plataforma
		
		On Error Resume Next ' Si hay errores en Excel, que continúe
		Dim cylFMFile
		Set cylFMFile = m_ExcelApp.OpenFile(strCylsPreMatcheckFPath, True, True)
	
		If (cylFMFile Is Nothing) Then Exit Function ' No se ha podido abrir el Excel con informacion de cilindros

		Dim oXlSheet
		Set oXlSheet = cylFMFile.Workbook.worksheets("Hoja1")
	
		Call oXlSheet.Range(strfilterRng).Select
		Call objExcel.Selection.Replace (" Bar", "", 2, 1, False, False, False)
		Stop ' para revisar que se use bien
		Dim oDicMatches, cS, oABCGas_XLS_Stage, rFiltered, cval, cilrow, strDescr
		cS = 0
		For Each oABCGas_XLS_Stage In oDicStages.Items
			' Obtiene el material del cilindro de cada etapa, a partir de la hoja de excel; y el LIMITE DE PRESIÓN QUE RESISTE.
			cS = cS + 1
			Call objExcel.Selection.AutoFilter
			oXlSheet.Range(strfilterRng).AutoFilter 1, "=*Ø " & iEtapaDiam(cS) & "*" , 1
			'oXlSheet.Range(strfilterRng).AutoFilter 4, ">=" & oABCGas_XLS_Stage.Stage_Pout, 1
			Set oDicMatches = CreateObject("scripting.dictionary")
			Set rFiltered = oXlSheet.AutoFilter.Range
			For Each cilrow In rFiltered.Offset(1).Resize(rFiltered.Rows.Count - 1).Columns(1).SpecialCells(12) ' solo celdas visibles
				strDescr = cilrow.Value
				Select Case True
					Case InStr (strDescr,"EN-GJL-") > 0 : strDescr = Replace (strDescr,"EN-GJL-","fund. gris " & "EN-GJL-")
					Case InStr (strDescr,"EN-GJS-") > 0 : strDescr = Replace (strDescr,"EN-GJS-","fund. nodular " & "EN-GJS-")
					Case InStr (strDescr," GGG-") > 0 : strDescr = Replace (strDescr," GGG-"," fund. nodular " & " GGG-")
					Case InStr (strDescr," GG-") > 0 : strDescr = Replace (strDescr," GG-"," fund. nodular " & " GG-")
					Case InStr (strDescr," F-114") > 0 : strDescr = Replace (strDescr," F-114"," forjado " & " F-114")
				End Select
				cval = 0
				If cilrow.Offset (0,3) >= oABCGas_XLS_Stage.Stage_Pout Then
					cval = 3
				ElseIf cilrow.Offset (0,3) < oABCGas_XLS_Stage.Stage_Pout And cilrow.Offset (0,4) > oABCGas_XLS_Stage.Stage_Pout Then
					strDescr = strDescr & " (¡OJO!: a presión de ensayo en probadero)"
					cval = 4
				End if
				If cval = 0 Then
					' cilindro no valido
				ElseIf Not oDicMatches.Exists (strDescr) Then
					oDicMatches.Add strDescr, cilrow.Offset (0,3)
				ElseIf oDicMatches(strDescr).Value < cilrow.Offset (0,3).Value Then
					Set oDicMatches(strDescr) = cilrow.Offset (0,3)
				End if
			Next
			If oDicMatches.Count > 0 Then
				cval = Empty
				For Each strDescr In oDicMatches.Keys
					Select Case True
						Case cval = Empty, oDicMatches(strDescr).Value < oDicMatches(cval).Value : cval = strDescr
					End Select
				Next
				' El cilindro de la etapa CASA CON ALGUNO DE LA TABLA DE EXCEL, y podría fabricarse COMO CILINDRO ESTANDAR
				Stop ' COMPRUEBA QUE SE ASIGNAN BIEN LOS VALORES SIGUIENTES!!!
				oABCGas_XLS_Stage.cilMaterial = cval
				oABCGas_XLS_Stage.cilPressureLimit = oDicMatches(cval).Value
			Else
				' NO se ha encontrado ningún cilindro de dimensiones estándar, que aguante la presión de la etapa 
				' --> habría que fabricar el cilindro A MEDIDA
				oABCGas_XLS_Stage.cilMaterial = "a medida (posiblemente forjado)"
				'oABCGas_XLS_Stage.cilPressureLimit = Empty
			End If
		Next		
		m_ExcelApp.CloseFile cylFMFile.FilePath, False
		If Err Then Call MsgLog ("No se ha podido hacer la selección de cilindros, por presiones / materiales!")
		On Error Goto 0

		bCylinderMaterialsProcessed_ = True
	End Function
	
	Public Function bEtapaIsTandem(iEtapa) ' iEtapa es un indice que COMIENZA EN UNO!!!
		bEtapaIsTandem = oDicStages.Item(iEtapa-1).bTandem
	End Function
	
	Public Function iEtapaDiam(iEtapa) ' iEtapa es un indice que COMIENZA EN UNO!!!
		iEtapaDiam = oDicStages.Item(iEtapa-1).diamCils 
	End Function
	
	Public Function iEtapaNCils(iEtapa) ' iEtapa es un indice que COMIENZA EN UNO!!!
		iEtapaNCils = oDicStages.Item(iEtapa-1).numCils
	End Function
	
	Public Function bEtapaReqForjado(iEtapa) ' iEtapa es un indice que COMIENZA EN UNO!!!
		bEtapaReqForjado = oDicStages.Item(iEtapa-1).Stage_Pout > 80
		' en el caso de H2 / ATEX... SE PONEN FORJADOS INCLUSO DESDE MAS ABAJO:
		bEtapaReqForjado = bEtapaReqForjado Or (bATEX_Inflamable And oDicStages.Item(iEtapa-1).Stage_Pout > 80 * 0.85)
		' ESTO ES FALSO!!!: si la plataforma es HP o HX, SIEMPRE van forjados (los cilindros PUEDE QUE NO!!, si acaso, bielas, etc!!!)
		' bEtapaReqForjado = bEtapaReqForjado Or (oCell_Compressor_Serie = "HP" Or oCell_Compressor_Serie = "HX")
		' si el diam es menor de 75 tb van forjados, NO se pueden hacer en fundic nodular
		bEtapaReqForjado = bEtapaReqForjado Or oDicStages.Item(iEtapa-1).diamCils <= 75
		' compresores HP o HX, DE DIAMETROS GRANDES, Y A ALTAS PRESIONES (ya incluso inferiores a 80 bar)... convendría que fuesen encamisados
		' (se hace en fundicion nodular o acero fundido el cuerpo, y el liner añade proteccion)
		bEtapaReqForjado = bEtapaReqForjado Or (oDicStages.Item(iEtapa-1).diamCils >= 450 _
				And oDicStages.Item(iEtapa-1).Stage_Pout > 65)
	End Function
	
	Public Function bEtapaConvieneCamisa(iEtapa) ' iEtapa es un indice que COMIENZA EN UNO!!!
		bEtapaConvieneCamisa = oDicStages.Item(iEtapa-1).Stage_Pout > 80
		' ESTO ES FALSO!!!: si la plataforma es HP o HX, SIEMPRE van forjados (los cilindros PUEDE QUE NO!!, si acaso, bielas, etc!!!)
		' bEtapaReqForjado = bEtapaReqForjado Or (oCell_Compressor_Serie = "HP" Or oCell_Compressor_Serie = "HX")
		' si el diam es menor de 75 tb van forjados, NO se pueden hacer en fundic nodular
		bEtapaConvieneCamisa = bEtapaConvieneCamisa Or oDicStages.Item(iEtapa-1).diamCils <= 75
	End Function
	
	Private strModelName_
	Function strModelName
		If Not IsEmpty (strModelName_) Then strModelName = strModelName_ : Exit Function
		Dim strCils, oABCGas_XLS_Stage, nEtapa
		On Error Resume Next
		For Each oABCGas_XLS_Stage In oDicStages.Items
			nEtapa = nEtapa + 1
			strCils = strCils & "-" & oABCGas_XLS_Stage.numCils & "x" & oABCGas_XLS_Stage.diamCils
			If oABCGas_XLS_Stage.bTandem Then strCils = strCils & "T"
			If bEtapaReqForjado(nEtapa)	Then
				strCils = strCils & "FC"	
			ElseIf bNACE_Corrosivo Then
				' en gases corrosivos, la camisa protege el cuerpo del cilindro. API 618 practicamente LO EXIGE en caso de NACE
				strCils = strCils & "C"	
			ElseIf bEtapaConvieneCamisa(nEtapa) Then
				strCils = strCils & "(C)"	
			End If
		Next
		strModelName = oDicStages.Count
		If InStr (UCase(strCils),"T") > 0 Then strModelName = strModelName & "T"
		strModelName = strModelName & "E" & oCell_Compressor_Serie & "-" & ncils & "-"
		If bATEX_Inflamable Then strModelName = strModelName & "L"
		If bAire Then
			strModelName = strModelName & "LT" & strCils
		Else
			strModelName = strModelName & "GT" & strCils
		End If
		If bNACE_Corrosivo Then strModelName = strModelName & " NACE"
		If bATEX_Inflamable Then
			strModelName = strModelName & " ATEX" : MsgBox ("poner TODO para ATEX: distanciador tipo C - bloque SAS; panel de purga de N2 + venting; valvula bicera en carter (ATEX), y motor ATEX, etc; zonificando paneles")
		ElseIf bSafeZone Then
			' tiene H2; pero dependiendo del PM de la mezcla... Si < 12, **FULL ATEX**, distanciador, motor, etc; si > 12, (ATEX) : solo MOTOR; y se haria clasificacion de zona
			strModelName = strModelName & " (ATEX)" : MsgBox ("poner valvula bicera en carter (ATEX), y motor ATEX, zonificando paneles")
		End If
		strModelName_ = strModelName
		If Err Then Call MsgLog ("No se ha podido determinar el modelo del compresor!")
		On Error Goto 0
	End Function

	Private Function fixStringFN (str)
		fixStringFN = Replace (str,";","-")
		fixStringFN = Replace (fixStringFN,"|","-")
		regex.Pattern = "\-*[\<\>\*\!\?\/:]"
		fixStringFN = regex.Replace(fixStringFN,"-")
	End Function
	
	Public Function getABCFileName (revisionNr)
		' revisionNr indica LA OPCION DE CALCULO QUE ESTA REVISARÍA!!! (NO es un 'orden' 01-02-03... de numero de revision, OJO!!!)
		Dim regex,match,strFType,strExt
		Set regex = New RegExp
		regex.Global = True : regex.IgnoreCase = True : regex.multiline = False
		' en lo siguiente SIEMPRE debería ser calc, al menos de momento: en ESTA CLASE aún no se procesan el resto de ficheros!!!
		regex.Pattern = "(?:ABC_(Gas_Cooler|Aircooler|Reducer|Main Motor|Instrumentation|Gas_Filter|Frequency Converter|Cooling Water Pump|Dryer|Piston_rider_ring_selection|Cooling Water Tower|" & _
				"Pressure_Safety_Valve|Valves_selection)\-|.*?curv.*?)?([A-Z]{3}\d{5}_\d{2})(_calc(?:_multi)?|.*?curv.*?)?.*?(_old\(\d+\))?(\.(?:xlsx|rtf))$"
		On Error Resume Next
		For Each match in regex.Execute (strXLSXPath)
			If oCellCalculo.Value <> match.submatches (1) Then MsgBox ("El cálculo NO corresponde con el nombre del fichero!!") : Stop
			If match.submatches(0) = "" Then 
				strFType = match.submatches(2) 
			Else 
				strFType = "_" & match.submatches(0) 
			End If
			
			If strFType = "" And InStr (LCase(fso.getfileName(strXLSXPath)),"curv") > 0 Then
				MsgBox ("Fichero pendiente de procesar: curvas de funcionamiento")
				Stop ' INTERESA PROCESAR TB LOS FICHEROS DE CURVAS DE funcionamiento!!!, seria bueno SACAR DE ELLAS EL RENDIMIENTO, Y LAS CURVAS DE SENSIBILIDAD a temperatura, presion, etc...
				strFType = "working curves"
			End If

			strExt = match.submatches (4)
		
			If match.submatches (3) <> "" Then
				'Stop ' tiene en cuenta versiones obsoletas
				'getABCFileName = Left (getABCFileName,Len(getABCFileName)-Len(strExt)) & match.submatches (3) & strExt
				strExt = match.submatches (3) & strExt
			End If
		Next
		If Err Then Call MsgLog ("No se ha podido generar un nombre para el fichero!") : Exit Function
		On Error Goto 0
		
		getABCFileName = fso.GetParentFolderName (strXLSXPath) & "\" & oCellCalculo.Value & strFType & strExt
		
		If False And oCellFecha.Value <> "" Then
			strFecha = Split (oCellFecha.Value,"/")(2) & "-" & Split (oCellFecha.Value,"/")(1) & "-" & Split (oCellFecha.Value,"/")(0)
			getABCFileName = Replace(getABCFileName,strExt,"_" & strFecha & strExt)
		End If
		
		' el MODELO de máquina que resulta del cálculo
		getABCFileName = Replace(getABCFileName,strExt,"_" & strModelName & strExt)
		MsgIE ("<font color=blue>DEBE aparecer en el nombre del calculo, un <u><b>IDENTIFICADOR DEL 'PROYECTO / COMPRESOR (del 'ITEM')'</b></u>, tal y como lo pide el cliente!! (puede haber varios calculos con resultados diferentes para un mismo proyecto...)</font>")
		'getABCFileName = Replace(getABCFileName,strExt,"_" & fixStringFN (oCellCliente.value) & strExt)
		'getABCFileName = Replace(getABCFileName,strExt,"_" & fixStringFN (oCellProyecto.value) & strExt)
		
		On Error Resume Next
		If oCellObservaciones.Value <> "" Then
			' este campo debería tener info SOLO PARA DIFERENCIAR LOS CALCULOS, la razón de los mismos:
			' - CONDICIONES OPERATIVAS
			' - DISPARO VALVULA SEGURIDAD
			' - CONDICIONES DISEÑO
			' ...
			regex.Pattern = "(?:cond[\.\S]*\s+(?:operat|nominal)|disp[\.\S]*\s+valv[\.\S]*\s+seg|cond[\.\S]*\s+(?:dise|operat|nominal)[\.\S]*\s+cond|safety\s+valve\s+trig|des\.?(?:ign)?\s+cond|dim[\.\S]*\s+motor|pow[\.\S]*\s+siz)[\.\S]*"
			If regex.Test (oCellObservaciones.value) Then
				getABCFileName = Replace(getABCFileName,strExt,"_" & regex.Execute(oCellObservaciones.value).Item(0).Value & strExt)
			Else
				MsgIE ("<font color=blue>" & "el campo de OBSERVACIONES en gas_vbnet debería tener info SOLO PARA DIFERENCIAR LOS CALCULOS, la razón de los mismos:" & vbCr & _
						"- CONDICIONES OPERATIVAS" & vbCr & "- DISPARO VALVULA SEGURIDAD" & vbCr & "- CONDICIONES DE DISEÑO" & vbCr & "..." & "</font>")
				If oCellProyecto.value <> "" Then
					getABCFileName = Replace(getABCFileName,strExt,"_" & Left(Trim(fixStringFN (Replace(Replace(Replace(UCase(oCellProyecto.value),"PROJECT",""),"PROYECTO",""),"",""))),30) & strExt)
				ElseIf oCellObservaciones.value <> "" Then
					getABCFileName = Replace(getABCFileName,strExt,"_" & oCellObservaciones.value & strExt)
				End if
			End if
			regex.Pattern = "rev[\._ \-]*(?:isi[oó]n)?(?:\s*de)?(?:[_ \-]*opc[\. ]*(?:i[oó]n)?(?:(?:[\.\s]*de)?\s*c[áa]lc[\.\s]*(?:ulo)?)?)?[\._ \-]*(\d+)"
			If regex.Test (oCellObservaciones.value) Then
				Stop ' ESTO HAY QUE SACARLO A cOp_CalcsTecn
				Set match = regex.Execute(oCellObservaciones.value).item(0)
				If MsgBox ("Detectada una revisión de otra opción de cálculo:" & vbCr & fso.getfilename (strXLSXPath) & vbcr & "¿es la revisión de la opción de cálculo " & _
						match.Submatches(0) & " ?",4) = 6 Then
					revisionNr = match.Submatches(0) : stop
				End If
			End If
			regex.Pattern = "descartado|discarded"
			If regex.Test (oCellObservaciones.value) Then
				MsgBox ("el campo de OBSERVACIONES en gas_vbnet indica que este cálculo, " & fso.getfilename (strXLSXPath) & _
						" se ha DESCARTADO. Se remarca en el nombre de fichero.")
				Stop ' asegurarse de que en el nombre de fichero aparece el término DESCARTADO / DISCARDED
			End If
		End If
		If Err Then Call MsgLog ("No se ha podido generar un nombre para el fichero!") : Exit Function
		On Error Goto 0
		
		If Not IsEmpty (revisionNr) Then
			Stop ' PTE DE IMPLEMENTAR
			getABCFileName = Replace(getABCFileName,oCellCalculo.Value & strFType,oCellCalculo.Value & strFType & "_rev " & revisionNr)
		End If
		
		getABCFileName = fso.GetParentFolderName (getABCFileName) & "\" & fixStringFN (fso.GetFileName(getABCFileName))
		If Len (getABCFileName) > 254 Then
			getABCFileName = Left(Replace(getABCFileName,strExt,""),254-Len(strExt)) & strExt
		End If
	End Function

	Private Function ReplaceInAllCells (Range,strfrom,strto, ByRef bSave)
		If Range is Nothing THen Exit Function
		Dim oCell, strPrevCellAddress
		' Busqueda parcial, xlPart = 2, en xlValues = , -4163
        'Set oCell = Range.Find(strfrom,Range.Application.ActiveCell,-4163,2)
        ' la siguiente es para CASE SENSITIVE, por si acaso
        Set oCell = Range.Find(strfrom,Range.Application.ActiveCell,-4163,2,1,1,True,False)
		If Not oCell is Nothing Then 
			bSave = True
            Do Until oCell Is Nothing 
            	If strPrevCellAddress = oCell.Address Then Exit Do
                oCell.Value = Replace(oCell.Value,strfrom,strto)
                strPrevCellAddress = oCell.Address
                Set oCell = Range.FindNext(oCell)
            Loop
		End if
	End Function
	
	Private oGASINGXSLSheet_Fixed_	
	Public Function fixCGASING()
		If Not IsEmpty (oGASINGXSLSheet_Fixed_) Then
			Set fixCGASING =  oGASINGXSLSheet_Fixed_
			Exit Function ' SOLO SE PROCESA UNA VEZ esta función: la info que lee NO CAMBIA, --> NO tiene sentido hacerlo más veces
		End If
		If m_ExcelFMFile Is Nothing Then Exit Function
		If Not m_ExcelFMFile.HasWorksheet("C-GAS-ING") Then
			Exit Function
		End If
		' cambia algunas cadenas de texto a ingles en C-GAS-ING
		Dim oXlSheet,oCell,c,d,vtmp,bExcelScreenUpdating
		Set oXlSheet = m_ExcelFMFile.GetWorksheet("C-GAS-ING")
		bExcelScreenUpdating = objExcel.ScreenUpdating
	    If bQuickExcel Then objExcel.ScreenUpdating = False
		oXlSheet.Activate
	    objExcel.ActiveWindow.Zoom = 100
		oXlSheet.Range("A1").Select
		Call ReplaceInAllCells (oXlSheet.Cells,"Vapor de agua","Water vapor",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"Agua","Water",bSave)
		' Busqueda parcial, xlPart = 2, en xlValues = , -4163
		Call ReplaceInAllCells (oXlSheet.Cells,"Límite RPM","RPM Limit",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells," / 0 ( 0 )","",bSave)  
		Call ReplaceInAllCells (oXlSheet.Cells,"Seco-LT","Dry-LT",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"o Dry-LT","or Dry-LT",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"Atmosférico (Normal)","Atmospheric (Standard)",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"Metros","Meters",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"Composición del gas en Volumen :","Gas composition by volume :",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"Aire seco","Dry air",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"Aire","Air",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"Monóxido de Carbono","Carbon monoxide",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"Anhídrido Carbónico, Dióxido de Carbono","Carbon dioxide",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"Acido Sulfhídrico, Sulfuro de Hidrógeno","Hydrogen sulfide",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"Nitrógeno","Nitrogen",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"Hidrógeno","Hydrogen",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"Oxígeno","Oxygen",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"Metano","Methane",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"Etano","Ethane",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"Propano","Propane",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells, "propano", "propane", bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"Etileno, Eteno","Ethylene, Ethene",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"Argón","Argon",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"Propileno, Propeno","Propylene, Propene",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"Butano","Buthane",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"butano","buthane",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"Metil","Methyl",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"metil","methyl",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"Amoníaco","Ammonia",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"Pentano","Penthane",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"pentano","penthane",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"Hexano","Hexane",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"Autor :","Author :",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"Fecha :","Date :",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"- Pressure ","- Exhaust pressure ",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"CV/KW","HP/kW",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells," CV"," HP",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,"-Amount and Diameter :","- Amount and diameter:",bSave)
		Call ReplaceInAllCells (oXlSheet.Cells,Wshshell.ExpandEnvironmentStrings("%username%"),objExcel.Application.UserName,bSave)
		Set oCell = oXlSheet.Cells.Find("CH4O       : ",oXlSheet.Cells.Application.ActiveCell,-4163,2,1,1,True,False)
		If Not oCell is Nothing Then if oCell.Offset(0, 2).Value <> "Methanol" Then oCell.Offset(0, 2).Value = "Methanol" : bSave = True
		Set oCell = oXlSheet.Cells.Find("Total mechanical losses : ",oXlSheet.Cells.Application.ActiveCell,-4163,2,1,1,True,False)
		regex.Pattern = "([\d,]+) HP" 
		If Not oCell is Nothing Then
			If InStr (UCase(oCell.Offset(0, 1).Value)," HP/KW") = 0 then
				c = regex.Execute(oCell.Offset(0, 1).Value).Item(0).Submatches(0) * 1
				oCell.Offset(0, 1).Value = Round(c, 2) & " / " & Round(c * 0.7457,2) & " HP/kW"
				bSave = True
			End if
		End If
		regex.Pattern = "\s*:\s*"
		For Each oCell In oXlSheet.Range("F19:F29")
			oCell.value = regex.Replace(oCell.value, "")
		Next
		MsgLog vbtab & "Corregidos errores de idioma y texto en C-GAS-ING"
		Set oCell = oXlSheet.Cells.Find("Compressor model : ",oXlSheet.Cells.Application.ActiveCell,-4163,2,1,1,True,False)
		regex.Pattern = "^(.+?)\-\d+x" 
		If Not oCell is Nothing Then oCell.Offset(0, 1).Value = regex.Execute(strModelName).Item(0).Submatches(0)
	    ' mostrar celdas ocultas, para eliminarlas
		If oXlSheet.Range("A60:A60").Value <> "" then
			oXlSheet.Rows("1:100").Select
			oXlSheet.Application.Selection.EntireRow.Hidden = False
			If oXlSheet.Cells.Find("Motor at max.  : ") Is Nothing Or oXlSheet.Cells.Find("Isothermal efficiency : ") Is Nothing Then
				'Stop
			else
				'xlShiftUp = -4162' CÓMO SE DESPLAZAN LAS CELDAS PARA SUSTITUIR A LAS ELIMINADAS
				oXlSheet.Rows("52:53").Delete
				oXlSheet.Rows("53:55").Delete
				oXlSheet.Rows("63:64").Delete
				oXlSheet.Rows("64:87").Delete
				oXlSheet.Rows("39:39").Delete
				MsgLog vbtab & "Eliminadas filas ocultas en C-GAS-ING"
			End If
			bSave = True
	    End If
	    
		If oXlSheet.Range("E45:E45").Value <> "" Then
			' EL FLOW DRY / WET
		    ' xlDown, -4121 (inserta desplazando filas hacia abajo); xlFormatFromLeftOrAbove = 0 (el formato de las celdas insertadas es el de las de encima)
		    oXlSheet.Rows("46:46").Insert -4121, 1
		    oXlSheet.Range("E45:F45").Cut oXlSheet.Range("B46")
		    oXlSheet.Range("A45").Value = "Actual flow :"
		    If InStr (oXlSheet.Range("B46").Value, "kg") > 0 Then oXlSheet.Range("A46").Value = "Mass flow (dry / wet):" Else oXlSheet.Range("A46").Value = "Normal flow (dry / wet):" End If
			oXlSheet.Range("C45:D45").Copy
		    oXlSheet.Range("C46").PasteSpecial -4122, -4142, False, False
		    oXlSheet.Application.CutCopyMode = False
			bSave = True
		Else
			'Stop ' será que la hoja está en un fichero ya modificado... pero asegurarse
		End If
	    
	    If (oXlSheet.Range("G30").Value <> "Suction pressure :" And oXlSheet.Range("F30").Value <> "Specific weight in normal conditions:") _
	    		Or (oXlSheet.Range("F29").Value <> "") Then	
			' dimensiona la lista de gases, PARA QUE TODAS LAS CELDAS TENGAN EL FORMATO CORRECTO
'			regex.Pattern = "([\d,]+)\%" '13,99%
'			For Each oCell In oXlSheet.Range("G19:G28")
'				If regex.Test (oCell.Value) Then oCell.Value = regex.Execute(oCell.Value).Item(0).submatches(0)*1 & "%"
'			Next

'		    oXlSheet.Range("E19").FormulaR1C1 = "1"
'		    oXlSheet.Range("E19").Cut
'		    oXlSheet.Range("G19:G28").PasteSpecial -4163, 4, False, False ' CONVERSION A NUMEROS, xlPasteValues = -4163, xlMultiply = 4
'		    oXlSheet.Application.CutCopyMode = False

'		    For Each oCell In oXlSheet.Range("G19:G28")
'			    oCell.NumberFormat = "General"
'			    oCell.Value = Replace(Trim(Replace(oCell.Value,"'","")),"%","") / 100
'			    oCell.NumberFormat = "0.00%"
'		    next
	    	' hago una ordenación muy simple de los valores
			regex.Pattern = "([\d,]+)\%" '13,99%
	    	c = 19
	    	Do
		    	Set oCell = oXlSheet.Range("G" & c & ":G" & c)
		    	
				If regex.Test (oCell.Value) Then oCell.Value = regex.Execute(oCell.Value).Item(0).submatches(0)*1 & "%"
			    oCell.NumberFormat = "General"
			    oCell.Value = Replace(Trim(Replace(oCell.Value,"'","")),"%","") / 100
			    oCell.NumberFormat = "0.00%"
			    
                d = 19 - c
                Do
                    If oCell.Value > oCell.Offset(d, 0).Value Then
                    	'Stop
                        vtmp = oCell.Offset(d, 0).Value: oCell.Offset(d, 0).Value = oCell.Value: oCell.Value = vtmp
                        vtmp = oCell.Offset(d, -1).Value: oCell.Offset(d, -1).Value = oCell.Offset(0, -1).Value: oCell.Offset(0, -1).Value = vtmp
                        vtmp = oCell.Offset(d, 1).Value: oCell.Offset(d, 1).Value = oCell.Offset(0, 1).Value: oCell.Offset(0, 1).Value = vtmp
                    End If
                    d = d + 1
                Loop While d <= 0
		    	c = c + 1
		    Loop While oXlSheet.Range("F" & c & ":F" & c).Value <> ""
		    
		 	oXlSheet.Range("F28").Value = "Other     : "
		    oXlSheet.Range("G28").FormulaR1C1 = "=1-SUM(R[-9]C:R[-1]C)"
		    oXlSheet.Range("H28").ClearContents
			' queda corregir las celdas de gases:
			c = 29
			Do While oXlSheet.range ("F" & c).value <> ""
				c = c + 1
			Loop
			oXlSheet.Range("F29:H" & c-1).Clear
			
			oXlSheet.Range("G30").Value = "Suction pressure :"
			oXlSheet.Range("G31").Value = "Atmospheric pressure :"
			oXlSheet.Range("G32").Value = "Suction temperature :"
			oXlSheet.Range("G33").Value = "Ambient temperature :"
			oXlSheet.Range("G34").Value = "Relative humidity :"
			oXlSheet.Range("G35").Value = "Water temperature :"
			bSave = True
			MsgLog vbtab & "Redimensionada la lista de gases en C-GAS-ING"
		End If

	    ' recoloca primera y segunda columnas de INPUT DATA
		If oXlSheet.range ("F29").value = "" And oXlSheet.range ("A24").value = "Specific weight in normal conditions:" _
					And oXlSheet.range ("A30").value = "Compressor series: " And oXlSheet.range ("G30").value = "Suction pressure :" Then
			' SI NO SE CUMPLE oXlSheet.range ("F29").value = ""... LAS CELDAS A MOVER SE HABRIAN REEMPLAZADO POR NOMBRES DE GASES!!!
			'		me aseguro además de que el resto de la hoja no se haya modificado, que sea "la original"; por si acaso
			' PRESENTACION ALTERNATIVA: RECOLOCA LAS FILAS ORDENANDO MEJOR LOS CONCEPTOS DE ENTRADA.. OJO!!, ESTO AFECTA A LAS OFERTAS GENERADAS
			' ASEGURARSE DE CAMBIAR LAS PLANTILLAS DE OFERTAS, LAS QUE HACEN REF A C-GAS-ING, AL HACER ESTE CAMBIO!!
		    'If objExcel.ScreenUpdating = True Then Stop
			oXlSheet.Range("A24:C26").Cut oXlSheet.Range("F37")
			
			oXlSheet.Range("A30:D30").Cut oXlSheet.Range("A17")
			oXlSheet.Range("A34:D34").Cut oXlSheet.Range("A18")
			oXlSheet.Range("A33:D33").Cut oXlSheet.Range("A19")
			oXlSheet.Range("A31:D31").Cut oXlSheet.Range("A20")
			oXlSheet.Range("A35:D36").Cut oXlSheet.Range("A21")
			
			If oXlSheet.Range("F34").Value = "" THen
				oXlSheet.Range("G34").Cut oXlSheet.Range("A23")
			Else
				oXlSheet.Range("A23").Value = "Relative humidity : "
				oXlSheet.Range("A17").Copy
    			oXlSheet.Range("A23").PasteSpecial -4122, -4142, False, False ' PEGA EL FORMATO
    			Application.CutCopyMode = False
			End If
			oXlSheet.Range("I34:J34").Cut oXlSheet.Range("B23")
			oXlSheet.Range("C22:D22").Copy
			oXlSheet.Range("C23").PasteSpecial -4122, -4142, False, False
			oXlSheet.Application.CutCopyMode = False
			
			If oXlSheet.Range("F30").Value = "" THen
				oXlSheet.Range("G30").Cut oXlSheet.Range("A24")
			Else
				oXlSheet.Range("A24").Value = "Suction pressure :"
				oXlSheet.Range("A17").Copy
    			oXlSheet.Range("A24").PasteSpecial -4122, -4142, False, False ' PEGA EL FORMATO
    			Application.CutCopyMode = False
			End If
			oXlSheet.Range("I30:J30").Cut oXlSheet.Range("B24")
			oXlSheet.Range("C23:D23").Copy
			oXlSheet.Range("C24").PasteSpecial -4122, -4142, False, False
			oXlSheet.Application.CutCopyMode = False
			
			If oXlSheet.Range("F32").Value = "" THen
				oXlSheet.Range("G32").Cut oXlSheet.Range("A25")
			Else
				oXlSheet.Range("A25").Value = "Suction temperature : "
				oXlSheet.Range("A17").Copy
    			oXlSheet.Range("A25").PasteSpecial -4122, -4142, False, False ' PEGA EL FORMATO
    			Application.CutCopyMode = False
			End If
			oXlSheet.Range("I32:J32").Cut oXlSheet.Range("B25")
			oXlSheet.Range("C23:D23").Copy
			oXlSheet.Range("C25").PasteSpecial -4122, -4142, False, False
			oXlSheet.Application.CutCopyMode = False
			
			If oXlSheet.Range("F33").Value = "" THen
				oXlSheet.Range("G33").Cut oXlSheet.Range("A26")
			Else
				oXlSheet.Range("A26").Value = "Ambient temperature : "
				oXlSheet.Range("A17").Copy
    			oXlSheet.Range("A26").PasteSpecial -4122, -4142, False, False ' PEGA EL FORMATO
    			Application.CutCopyMode = False
			End If
			oXlSheet.Range("I33:J33").Cut oXlSheet.Range("B26")
			oXlSheet.Range("C24:D24").Copy
			oXlSheet.Range("C26").PasteSpecial -4122, -4142, False, False
			oXlSheet.Application.CutCopyMode = False
			
			If oXlSheet.Range("F31").Value = "" THen
				oXlSheet.Range("G31").Cut oXlSheet.Range("A27")
			Else
				oXlSheet.Range("A27").Value = "Atmospheric pressure :"
				oXlSheet.Range("A17").Copy
    			oXlSheet.Range("A27").PasteSpecial -4122, -4142, False, False ' PEGA EL FORMATO
    			Application.CutCopyMode = False
			End If
			oXlSheet.Range("I31:J31").Cut oXlSheet.Range("B27")
			oXlSheet.Range("C25:D25").Copy
			oXlSheet.Range("C27").PasteSpecial -4122, -4142, False, False
			oXlSheet.Application.CutCopyMode = False
			
			oXlSheet.Range("A37:D37").Cut oXlSheet.Range("A28")
			
			oXlSheet.Range("A32:D32").Cut oXlSheet.Range("A29")
			
			If oXlSheet.Range("F35").Value = "" THen
				oXlSheet.Range("G35").Cut oXlSheet.Range("A30")
			Else
				oXlSheet.Range("A30").Value = "Water temperature : "
				oXlSheet.Range("A17").Copy
    			oXlSheet.Range("A30").PasteSpecial -4122, -4142, False, False ' PEGA EL FORMATO
    			Application.CutCopyMode = False
			End If
			oXlSheet.Range("I35:J35").Cut oXlSheet.Range("B30")
			oXlSheet.Range("C28:D28").Copy
			oXlSheet.Range("C30").PasteSpecial -4122, -4142, False, False
			oXlSheet.Application.CutCopyMode = False
			
			oXlSheet.Range("A38:D40").Cut oXlSheet.Range("A31")
			
			If oXlSheet.Range("F29").Value = "" Then
				oXlSheet.Range("F37:F39").Cut oXlSheet.Range("F30")
			Else
				stop
				' si no son blancos... primero reajustar la lista de gases, luego se actualizaría estas celdas
			End If
			
			oXlSheet.Range("G37:H39").Cut oXlSheet.Range("I30")
			oXlSheet.Rows("35:42").Delete -4162
			bSave = True
		Else
			' LAS CELDAS A MOVER EN LA SEGUNDA COLUMNA PODRIAN HABER REEMPLAZADO POR NOMBRES DE GASES!! --> HAY QUE REVISAR EL FORMATO y el script...
			'Stop
			'MsgBox ("OJO, ESTO HAY QUE REVISARLO, LAS CELDAS QUE QUIERO MOVER en primera y segunda columna PODRIAN NO TENER EL CONTENIDO CORRECTO...") : If Not bLog Then WScript.quit
		End If

		regex.Pattern = "^0\s+" '
		If regex.Test (oXlSheet.Range("B21:B21").Value) Then
			' CORRIGE LA CELDA DE FLOW (CAUDAL), CUANDO SE HAYAN FIJADO LAS RPM (el caudal NO es cero!, será el que aparezca en B38.. pero es un dato DE SALIDA, no INPUT)
			oXlSheet.Range("B21:B21").Value = "'-"
		End If
		If regex.Test (oXlSheet.Range("B31:B31").Value) Then
			' CORRIGE LA CELDA DE RPM, CUANDO SEA EL VALOR A DETERMINAR en funcion del caudal
			oXlSheet.Range("B31:B31").Value = regex.Replace(oXlSheet.Range("B31:B31").Value, "'- ")
		End If

		If oXlSheet.Columns("J:J").ColumnWidth > 10 then
			' ajusta anchos de columnas, para hacer la tabla mas presentable
		    oXlSheet.Columns("A:A").ColumnWidth = 31.6
		    oXlSheet.Columns("H:H").ColumnWidth = 12
		    oXlSheet.Columns("I:I").ColumnWidth = 9
		    oXlSheet.Columns("J:J").ColumnWidth = 6.5
		    oXlSheet.Columns("B:G").ColumnWidth = 9.8
		    bSave = True
	    End If
	    ' Añado unas conversiones de unidades...
	    If InStr (oXlSheet.Range("B25:B25").Value,"ºF") > 0 Or InStr (oXlSheet.Range("B26:B26").Value,"ºF") > 0 _
	    		Or InStr (oXlSheet.Range("B30:B30").Value,"ºF") > 0 Then
	    	For Each vtmp In Array ("B25:B25","B26:B26","B30:B30")
		    	If InStr (oXlSheet.Range(vtmp).Value,"ºC") > 0 Then
		    		oXlSheet.Range(vtmp).Value = Round(Replace (oXlSheet.Range(vtmp).Value," ºC","") * 9 / 5 + 32,1) & " °F"
		    	End If
	    	Next
	    	
	    	oXlSheet.Range("A53:A53").Value = Replace (oXlSheet.Range("A53:A53").Value,"ºC","°F")
	    	oXlSheet.Range("A54:A54").Value = Replace (oXlSheet.Range("A54:A54").Value,"ºC","°F")
			For c = Asc("B") To Asc("G")
				If oXlSheet.Range(Chr(c) & "53:" & Chr(c) & "53").Value <> "" then
	    			oXlSheet.Range(Chr(c) & "53:" & Chr(c) & "53").Value = Round(oXlSheet.Range(Chr(c) & "53:" & Chr(c) & "53").Value * 9 / 5 + 32,1)
	    		End if
				If oXlSheet.Range(Chr(c) & "54:" & Chr(c) & "54").Value <> "" then
	    			oXlSheet.Range(Chr(c) & "54:" & Chr(c) & "54").Value = Round(oXlSheet.Range(Chr(c) & "54:" & Chr(c) & "54").Value * 9 / 5 + 32,1)
	    		End if
	    	Next
	    End If
	    ' Otras pequeñas correcciones de los datos
	    ' Eliminar RPM en los datos de entrada, si se ha puesto caudal > 0
	    regex.Pattern = "^\d+\s*(\(\s*RPM Limit = \d+\s*\))?"
	    If oXlSheet.Range("B21") <> "-" Then If regex.Test(oXlSheet.Range("B31")) Then oXlSheet.Range("B31").Value = regex.Replace(oXlSheet.Range("B31").Value, "--$1") : bSave = True
	    If bSave Then 
		    With oXlSheet
		        .Range("B2").Font.Bold = True        ' Título CALCULATION - GAS
		        .Range("A15;A35;A47;F17").Font.Bold = True  ' INPUT DATA, OUTPUT DATA, STAGES, Coolers
		        '.Range("F18,G18").Font.Bold = True   ' Encabezados Gas/Percentage
		    End With
	    End If
		oXlSheet.Range("A1").Select
	    If bQuickExcel Then objExcel.ScreenUpdating = bExcelScreenUpdating
		If bSave Then oXlSheet.Application.ActiveWorkbook.Save
		Set fixCGASING = oXlSheet
		Set oGASINGXSLSheet_Fixed_ =  fixCGASING
	End Function
	
	Private oLimitsFeatsReqs_
	Public Function oLimitsFeatsReqs
		If Not oLimitsFeatsReqs_ Is Nothing Then Set oLimitsFeatsReqs = oLimitsFeatsReqs_ : Exit Function
		' obtiene, para el compresor, LOS LIMITES DE FUNCIONAMIENTO, y los posibles REQUISITOS IMPUESTOS, Y LAS CARACTS PREDEFINIDAS por la plataforma, 
		' por el numero de cilindros, por su diametro... o por cualquier otro factor limitante
		' SE HACE ANTES DE VALIDAR LOS CALCULOS, porque esa validación depende de esos limites; y ANTES DE GENERAR INFO PARA OFERGAS, porque esa info se apoya
		' en estos limites.
		Stop ' PARA CHEQUEAR ESTO
		Set oLimitsFeatsReqs_ = (New cLimitsFeatsReqs).Init (Me)
		Set oLimitsFeatsReqs = oLimitsFeatsReqs_
	End Function
	
	Function ValidarCalculos
		MsgLog vbtab & "Validando los cálculos en C-GAS"
		' el objeto de esta funcion es COMPROBAR SI HAY QUE REVISAR LOS CALCULOS...
		' PERO SE PUEDE INICIALIZAR UNA HOJA PARA SACAR INFO DE ELLA, SIN NECESIDAD DE VALIDARLOS!
		
		' Aseguramos contexto DOC
		MsgIE.setContainer "doc"

		' validacion POR PLATAFORMAS: esta validación, respecto a oLimitsFeatsReqs, DEBERIA HACERLA GAS_VBNET!!
		Dim oLimitsFeatsReqs
		Set oLimitsFeatsReqs = Me.oLimitsFeatsReqs
		
		' CALCULOS ABCAIRE, que DEBEN ser realizados en TODAS las ofertas
		MsgIE ("<font color=blue>" & "Respecto a las OPCIONES DE CALCULO que debería haber: conds operativas; conds de diseño; " & _
		"(DEBE CHEQUEARSE QUE <b>EN OBSERVACIONES</b> SE INDIQUE ESE DIMENSIONAMIENTO, DEBERIA HABER UN <b>FORMULARIO DE ENTRADA DE DATOS</b> EN EL QUE SE REFLEJE ESA TEMPERATURA MINIMA DE FUNCIONAMIENTO, para que el script valide que se ha hecho el dimensionamiento de potencia de motor)</font>")
		MsgIE "<font color=blue>El calculo ""RATED"" se hace PARA LAS TEMPERATURAS MAS ALTAS, LUEGO:" & vbCrLf & _
				"<ul><li> asegurarse de hacer EL <b>DIMENSIONAMIENTO DE POTENCIA DE MOTOR, ** PARA LAS TEMPERATURAS MAS BAJAS **</b>, FIJANDO LAS RPM, y <b>aumentando la PRESION DE SALIDA en un 10% respecto a la RATED / de diseño</b>..." & "</li>" & vbCrLf & _
				"<li> SI ES UN REACTOR, de H2, etc, SE DIMENSIONA PARA LAS PEORES CONDICIONES, y se comprueban TODAS, FIJANDO VUELTAS RPM, las de la peor condicion" & "</li>" & vbCrLf & _
				"<li> asegurarse de hacer la VERIFICACION DE DISPARO DE VALVULA DE SEGURIDAD, FIJANDO LAS RPM, y aumentando la PRESION DE SALIDA en un 10% respecto a la RATED / de diseño..." & "</li>" & vbCrLf & _
				"<li> asegurarse de hacer las VERIFICACIONES EN VTPARES: en particular, que haya LOAD REVERSAL (/ ""que NO haya ROD REVERSAL""); si no, se puede corregir AÑADIENDO MASA ""en la BIELA?? (o en la CRUCETA?)""; e incluso SUBIENDO las RPM, a veces" & "</li>" & vbCrLf & _
				"<li> la TEMPERATURA DE SALIDA DEL GAS, SIEMPRE la podemos dar 10 POR ENCIMA DE LA DEL AGUA DE REFRIGERACION, usando un INTERCOOLER FINAL (pero SIEMPRE hay que respetar los lims de temperatura del gas!)" & "</li>" _
				 & "</ul></font>"

		Call MsgIE.Spoiler (True,"background-color:orange;color:black;", "ADVERTENCIAS","id" & oCellCalculo.Value & "Advert",True)
		
		Dim c,oABCGas_XLS_Stage
		MsgIE ("<u>RELACION DE COMPRESION</u>")
		On Error Resume Next
		If oCell_Process_CompRatio_Mean < oLimitsFeatsReqs.mincomprel Then _
				MsgIE ("- La relación de compresión media es DEMASIADO BAJA, podría corregirse <b>reduciendo el número de etapas</b>.")
		If oCell_Process_CompRatio_Mean > oLimitsFeatsReqs.maxcomprel Then _
				MsgIE ("- La relación de compresión media es DEMASIADO ALTA, podría corregirse <b>aumentando el número de etapas</b>.")
		For Each c In oDicStages
			Set oABCGas_XLS_Stage = oDicStages(c)
			If oABCGas_XLS_Stage.oCell_Stage_CompRatio > oLimitsFeatsReqs.maxcomprel Or _
					oABCGas_XLS_Stage.oCell_Stage_CompRatio < oLimitsFeatsReqs.mincomprel _
					Then
				MsgIE ("- La relación de compresión en la etapa " & c + 1 & " no tiene un valor adecuado, " & oLimitsFeatsReqs.mincomprel & " < <b>" & _
						Round(oABCGas_XLS_Stage.oCell_Stage_CompRatio,2) & "</b> < " & oLimitsFeatsReqs.maxcomprel & ". Debería ser cercana al valor medio, " & oCell_Process_CompRatio_Mean.Value)
			End If 
		Next
		If Err.Number <> 0 Then
			MsgLog "No se puede acceder a datos del libro de excel: " & Err.Description
			Err.Clear
			Exit Function
		End If
		On Error GoTo 0
	
		Dim strMsg
		MsgIE ("<u>POTENCIA DEMANDADA</u>")
		If PotenciakW > oLimitsFeatsReqs.iOversizekW Then
			MsgIE ("- la potencia que demanda el compresor, " & PotenciakW  & " kW, es DEMASIADO ELEVADA, por encima de la habitual para esta plataforma (" & oLimitsFeatsReqs.iOversizekW & " kW). Sería conveniente SUBIR DE PLATAFORMA") 
		End If

		MsgIE ("<u>VELOCIDAD lineal DE PISTON, Y DE CIGÜEÑAL (rpm)</u>")
		If oLimitsFeatsReqs.iMaxPistonSpeedMS <> 0 And CellValue(oCell_OUTProcess_PistonMeanSpeed) > oLimitsFeatsReqs.iMaxPistonSpeedMS Then
			MsgIE ("- la velocidad media del pistón, " & oCell_OUTProcess_PistonMeanSpeed & ", es DEMASIADO ELEVADA, por encima de la habitual para esta plataforma (" & oLimitsFeatsReqs.iMaxPistonSpeedMS & " m/s). Sería conveniente REDUCIR LAS RPMs desde " & RPM & ", AUMENTANDO DIAMETROS, ...") 
		End If
		If oLimitsFeatsReqs.iMaxRPM <> 0 And RPM > oLimitsFeatsReqs.iMaxRPM Then
			MsgIE ("- la velocidad de giro del cigueñal, " & RPM & " RPM, es DEMASIADO ELEVADA, por encima de la habitual para esta plataforma (" & oLimitsFeatsReqs.iMaxRPM & " rpm). Sería conveniente REDUCIR LAS RPMs, AUMENTANDO DIAMETROS, ...")
			If oCell_Compressor_Serie = "HX" Then
				MsgIE ("- En esta plataforma, HX, la velocidad de cigueñal DEBERIA SER AUN MAS BAJA, <b>CERCANA A LAS 375 RPM</b> (NO puede ser menor, PARA GARANTIZAR LUBRICAC CIGÜEÑAL)")
			End If
		End If

		Call getCylindersMaterials_Limits()

		MsgIE ("<u>ANALISIS DE CADA ETAPA</u>")
		Dim TempPrevStage,cS
		cS = 0
		For Each oABCGas_XLS_Stage In oDicStages.Items
			cS = cS + 1
			MsgIE ("<u>ETAPA <b>" & cS & "</b></u>")
			MsgIE ("<u><i>TEMPERATURAS EN LA SALIDA</i></u>")
			' Validar temperaturas en la salida:
			If Not IsEmpty (TempPrevStage) And TempPrevStage < oABCGas_XLS_Stage.oCell_Stage_Tescape Then
				MsgIE ("- La <b>temperatura de salida de la etapa</b> es MAYOR QUE LA DE LA ETAPA ANTERIOR, y deberían ir DECRECIENDO --> convendria DISMINUIR DIAMETRO de clindros, (o AÑADIR UNA ETAPA)")
			End If
			If bH2 And cS = oDicStages.Count And oABCGas_XLS_Stage.oCell_Stage_Tescape > 135 Then
				MsgBox ("- La <b>temperatura de salida de la etapa</b>, " & oABCGas_XLS_Stage.oCell_Stage_Tescape & ", es MAYOR QUE EL LIMITE ACEPTABLE " & _
						" (PARA <b>H2</b>, 135ºC EN API V5, 120 EN API V6), DEBERIAS <b>INCREMENTAR EL NUMERO DE ETAPAS</b>")
			End If
			If cS = oDicStages.Count And oABCGas_XLS_Stage.oCell_Stage_Tescape > 180 Then
				MsgBox ("- La <b>temperatura de salida de la etapa FINAL</b>, " & oABCGas_XLS_Stage.oCell_Stage_Tescape & ", es MAYOR QUE EL LIMITE ACEPTABLE para cualquier gas, 180ºC, DEBERIAS INCREMENTAR EL NUMERO DE ETAPAS o disminuir diametros")
			End If
			
			TempPrevStage = oABCGas_XLS_Stage.oCell_Stage_Tescape
			
			MsgIE ("<u><i>PRESIONES EN EL CILINDRO</i></u>")
			' presión de diseño de la plataforma
			If cS = oDicStages.Count And oABCGas_XLS_Stage.Stage_Pout > oLimitsFeatsReqs.PMaxDesign Then
				MsgIE ("la <b>PRESIÓN DE SALIDA del compresor</b>, " & oABCGas_XLS_Stage.Stage_Pout & _
						", es DEMASIADO ELEVADA, por encima del limite de diseño para esta plataforma (" & oLimitsFeatsReqs.PMaxDesign & " bar)") 
			End If

			' PRESIONES LIMITE EN LOS CILINDROS: debería hacerlo gas_vbnet, mas preciso que el excel de los cilindros...
			If IsEmpty (oABCGas_XLS_Stage.cilPressureLimit) Then
				Stop ' en la tabla de cilindros NO HAY NINGUNO que alcance la presión de la etapa --> se haría a medida, ¿¿O HARIA LOS CALCULOS CON OTRO CILINDRO??
				MsgIE ("el cilindro de la etapa, como cilindro estándar (" & oABCGas_XLS_Stage.cilMaterial & ") soportaría una PRESIÓN EXCESIVA, " & _
						oABCGas_XLS_Stage.Stage_Pout & ". No hay cilindros para esa presión? --> <b>debería USAR UN CILINDRO MAS PEQUEÑO?</b>")
			End IF
			
			MsgIE ("<u><i>DIAMETROS DEL CILINDRO</i></u>")
			' diametros maximo y minimo
			If iEtapaDiam(cS) < oLimitsFeatsReqs.iMinDiam Then
				MsgIE ("- el diámetro del cilindro de la etapa, " & iEtapaDiam(cS) & ", es MENOR QUE EL LIMITE ACEPTABLE, " & _
						oLimitsFeatsReqs.iMinDiam)
			End If
			If iEtapaDiam(cS) > oLimitsFeatsReqs.iMaxDiam Then
				MsgIE ("- el diámetro del cilindro de la etapa, " & iEtapaDiam(cS) & ", es MAYOR QUE EL LIMITE ACEPTABLE, " & _
						oLimitsFeatsReqs.iMaxDiam & " (podría ser conveniente subir de plataforma)")
			End If

			MsgIE ("<u><i>ESFUERZOS EN EL VASTAGO DEL CILINDRO</i></u>")
			' esfuerzos de compresión (y traccion)
			If oABCGas_XLS_Stage.oCell_Stage_ComprStress > oLimitsFeatsReqs.iMaxLoad * 0.85 Then
				MsgIE ("- LOS ESFUERZOS DE COMPRESION EN VASTAGO DE LA ETAPA, " & oABCGas_XLS_Stage.oCell_Stage_ComprStress & _
					" kg) son SUPERIORES AL LIMITE DE LA PLATAFORMA, " & oLimitsFeatsReqs.iMaxLoad & " * 0.85. conviene BAJAR DIAMETROS DE CILINDROS, O AÑADIR ETAPAS.")
			End If
			If oABCGas_XLS_Stage.oCell_Stage_TensStress > oLimitsFeatsReqs.iMaxLoad * 0.85 Then
				MsgIE ("- LOS ESFUERZOS DE TRACCION EN VASTAGO DE LA ETAPA, " & oABCGas_XLS_Stage.oCell_Stage_TensStress & _
					" kg) son SUPERIORES AL LIMITE DE LA PLATAFORMA, " & oLimitsFeatsReqs.iMaxLoad & " * 0.85. conviene BAJAR DIAMETROS DE CILINDROS, O AÑADIR ETAPAS.")
			End If
			If oABCGas_XLS_Stage.oCell_Stage_TensStress > 0 And oABCGas_XLS_Stage.oCell_Stage_ComprStress > 0 Then
				If Abs (oABCGas_XLS_Stage.oCell_Stage_TensStress - oABCGas_XLS_Stage.oCell_Stage_ComprStress) / oABCGas_XLS_Stage.oCell_Stage_ComprStress > 0.2 Then
					MsgIE ("- LOS ESFUERZOS DE TRACCION - COMPRESION ESTÁN MUY DESCOMPENSADOS en la etapa, puede que SUBIENDO EL DIAMETRO se compensen algo...")
				End If
			End If
		Next
		
		If bN2 And RPM > 500 Then
			MsgIE ("- <b>N2 (gas seco, abrasivo)</b>: la velocidad, " & RPM & " RPM, es DEMASIADO ELEVADA, > 500 rpm, deberia ser de unas 500 RPM o menos") 
		End If

		MsgIE.popContainer ' "id" & oCellCalculo.Value & "Advert"
		
		Call MsgIE.Spoiler (True,"color:blue;", "API618","id" & oCellCalculo.Value & "API618",True)
		MsgIE ("<u>VALIDACION DE CONDICIONES API-618 y otras normativas</u>")
		' si se solicita expresamente que el compresor sea API 618, de manda a ADVERTENCIAS; si no, "en azul"
		If Not m_bAPI618 Then
			'SEGUIR AQUI
			Call MsgIE.Spoiler (True,"color:blue;", "NOTAS","id" & oCellCalculo.Value & "Notes",True)
			MsgIE "<font color=blue>El calculo ""RATED"" se hace PARA LAS TEMPERATURAS MAS ALTAS, LUEGO:" & vbCrLf & _
					"<ul><li> asegurarse de hacer EL <b>DIMENSIONAMIENTO DE POTENCIA DE MOTOR, ** PARA LAS TEMPERATURAS MAS BAJAS **</b>, FIJANDO LAS RPM, y <b>aumentando la PRESION DE SALIDA en un 10% respecto a la RATED / de diseño</b>..." & "</li>" & vbCrLf & _
					"<li> SI ES UN REACTOR, de H2, etc, SE DIMENSIONA PARA LAS PEORES CONDICIONES, y se comprueban TODAS, FIJANDO VUELTAS RPM, las de la peor condicion" & "</li>" & vbCrLf & _
					"<li> asegurarse de hacer la VERIFICACION DE DISPARO DE VALVULA DE SEGURIDAD, FIJANDO LAS RPM, y aumentando la PRESION DE SALIDA en un 10% respecto a la RATED / de diseño..." & "</li>" & vbCrLf & _
					"<li> asegurarse de hacer las VERIFICACIONES EN VTPARES: en particular, que haya LOAD REVERSAL (/ ""que NO haya ROD REVERSAL""); si no, se puede corregir AÑADIENDO MASA ""en la BIELA?? (o en la CRUCETA?)""; e incluso SUBIENDO las RPM, a veces" & "</li>" & vbCrLf & _
					"<li> la TEMPERATURA DE SALIDA DEL GAS, SIEMPRE la podemos dar 10 POR ENCIMA DE LA DEL AGUA DE REFRIGERACION, usando un INTERCOOLER FINAL (pero SIEMPRE hay que respetar los lims de temperatura del gas!)" & "</li>" _
					 & "</ul></font>"
			MsgIE.popContainer ' "id" & oCellCalculo.Value & "Notes"
		End If

		If oCell_Compressor_Serie = "HG" Then Call MsgIE ("important", "- **** LA PLATAFORMA HG NO ESTÁ PENSADA PARA SER 'TODO API-618', SERIA CONVENIENTE HACER EL COMPRESOR EN HP, que SI está pensada para API ****")
		If CellValue(oCell_OUTProcess_PistonMeanSpeed) > 3.5 Then Call MsgIE ("important", "- OJO!!: la velocidad media del pistón, " & oCell_OUTProcess_PistonMeanSpeed & ", es DEMASIADO ELEVADA, SUPERA EL LIMITE DE 3.5 m/s. Sería conveniente REDUCIR LAS RPMs, AUMENTANDO DIAMETROS, ...") 
		Call MsgIE ("important", "- Asegurarse de que EN TODAS LAS ETAPAS de ABCAire, LOS CILINDROS DEBEN IR ENCAMISADOS!!! (tenderá a AUMENTAR LAS RPM!!!)")
		Call MsgIE ("important", "- Superbolt, y recubrimiento de Vastagos en WC, serían opciones tb a incluir, en OFERGAS, en ""OTROS""!!")
		Call MsgIE ("important", "OTRAS NORMAS a tener en cuenta, que NO afectan a ABC Aire: TEMA, en refrigeradores; y ASME, o Merkblatter, en calderines antipulsadores. Y para compresores DE AIRE, en Petroquimica, API 617 y API 672")
		MsgIE.popContainer ' "id" & oCellCalculo.Value & "API618"
		
		' NOTAS sobre los compresores, que DEBERIAN ESTAR CONSIDERADAS POR DEFECTO, INCLUSO EN OFERGAS... Y QUE SIRVEN PARA DEFINIR CORRECTAMENTE EL BUDGET / QUOTATION!!!
		Call MsgIE.Spoiler (True,"color:blue;", "NOTAS","id" & oCellCalculo.Value & "Notes",True)
		If bBombaAceiteAux Then Call MsgIE ("este modelo de compresor, " & strModelName & ", LLEVA POR DEFECTO UNA BOMBA DE ACEITE AUXILIAR (""redundante""), ACCIONADA POR MOTOR ELECTRICO en el sist de lubricación. CON ELLO SE CUMPLE API 610 y la norma ISO 10438-3. DEBERIA considerarse en los precios por defecto (OFERGAS), pero EN EL BUDGET / QUOTATION, posiblemente haya que indicarlo!!")
		If bBielasForjadas Then Call MsgIE ("este modelo de compresor, " & strModelName & ", LLEVA POR DEFECTO las bielas FORJADAS (EN EL BUDGET / QUOTATION, posiblemente haya que indicarlo)")
		If iStroke <> 0 Then Call MsgIE ("POR DEFECTO, este modelo de compresor, " & strModelName & ", tiene una CARRERA MAXIMA DEL PISTÓN, DE " & iStroke & " MM ** pero se podría haber modificado en ABC ** (EN EL BUDGET / QUOTATION, posiblemente haya que indicarlo)")
		If iRodDiam <> 0 Then Call MsgIE ("POR DEFECTO, este modelo de compresor, " & strModelName & ", tiene un diametro de vastago, DE " & iRodDiam & " MM ** pero se podría haber modificado en ABC ** (EN EL BUDGET / QUOTATION, posiblemente haya que indicarlo)")
		' capacidad en litros del CARTER, interesante para los calculos de MTTO 8000 h:
		If iCrankCaseLtr <> 0 Then Call MsgIE ("este modelo de compresor, " & strModelName & ", tiene un CARTER DE " & iCrankCaseLtr & " L (UTIL PARA CALCULAR LITROS DE ACEITE, EN EL INFORME DE MTTO 8000 H)")
		If bAir Then Call MsgIE ("este modelo de compresor, " & strModelName & ", por ser de Aire, NO LLEVA LLAVE DE ENTRADA")
		MsgIE.popContainer ' "id" & oCellCalculo.Value & "Notes"
		
		MsgIE.popContainer ' Salir de DOC

		If MsgBox ("Deseas VALIDAR LOS RESULTADOS DEL CALCULO DE VTPARES? - para hacerlo, es necesario TENER ABIERTA LA VENTANA DE RESULTADOS de VTPares, para este cálculo (es calcular hasta VTPares, y ** EN LA PESTAÑA DE Listados, seleccionar <Todo> **)",4) = 6 Then

			' SEGUIR AQUI
			If fso.FileExists ("W:\Aplicaciones\Compresores\Gas_vbnet.exe") then
				strPrgCmd = "W:\Aplicaciones\Compresores\Gas_vbnet.exe"
			Else
				strPrgCmd = "C:\Aire\Gas_vbnet.exe"
			End If
			strAireHND = lanzaPrg (strPrgCmd, "*ARIZAGA BASTARRICA Y CIA*", oAireExec, arrPosSize)
			Stop
			Wshshell.Run "guipropview /Action SwitchTo  Handle:" & strAireHND & "",0,true
			Wshshell.Run "guipropview /Action Focus  Handle:" & strAireHND & "",0,True
			WScript.Sleep 1000
			
			'strCmd = "guipropview /Action FocusKeyText ""SER00020"" Handle:" & strAireHND & " Child.ZOrder:4"
			'Wshshell.Run strCmd,0,true
			Call shellrun_Error ("guipropview /Action FocusKeyText ""SER00020"" Handle:" & strAireHND & " Child.ZOrder:4")
			WScript.Sleep 5000
			'Wshshell.Run "guipropview /Action FocusKeyText ""S E R 0 0 0 2 2 enter"" Handle:" & strAireHND & " Child.ZOrder:4",0,true
			Call shellrun_Error ("guipropview /Action FocusKeyText ""S E R 0 0 0 2 2 enter"" Handle:" & strAireHND & " Child.ZOrder:4")
			
			Call getAireGasControlIDs (strAireHND, oDicControls)
			If oDicControls.Count > 0 Then
				'Wshshell.Run "guipropview /Action MouseClick Left Handle:" & oDicControls("CalcularEtapas")(0) ,0,true
				Call shellrun_Error ("guipropview /Action MouseClick Left Handle:" & oDicControls("CalcularEtapas")(0))
			End If
			
			' los comandos de GUIPROPVIEW NO DA CONTROL DE LOS CUADROS DE TEXTO, ETC --> se usa NIRCMD, con 
			' nircmd setcursorwin  88 205 & nircmd sendmouse left click,
			' nircmd sendmouse move -300 0 etc
			
			Stop
			oAireExec.Terminate
			
			''' LA CAPTURA DEL CONTENIDO DE LA VENTANA DE VTPARES SE HACE CON LA UTILIDAD DE NIRSOFT, sysexp.exe
			



		End if
	End Function

	Function UpdateTags ()
		Dim oXlWorkBook : Set oXlWorkBook = m_ExcelFMFile.WorkBook
		oXlWorkBook.BuiltinDocumentProperties("Title") = strModelName
		oXlWorkBook.BuiltinDocumentProperties("Subject") = oCellCliente.Value
		If oCellProyecto.Value <> "" Then oXlWorkBook.BuiltinDocumentProperties("Subject") = oXlWorkBook.BuiltinDocumentProperties("Subject") & ". " & oCellProyecto.Value 
		oXlWorkBook.BuiltinDocumentProperties("Manager") = oXlWorkBook.BuiltinDocumentProperties("Author")
		oXlWorkBook.BuiltinDocumentProperties("Author") = Replace (oCellAutor.Value,"srey","Sergio Rey")
		oXlWorkBook.BuiltinDocumentProperties("Keywords") = strGasType
		oXlWorkBook.BuiltinDocumentProperties("Company") = "ABC Compressors"
		oXlWorkBook.BuiltinDocumentProperties("Comments") = oCellObservaciones.Value
		regex.Pattern = "\brev(is\S+|\.\b)\s+(\d+)"
		If regex.Test (oCellObservaciones.Value) Then 
			oXlWorkBook.BuiltinDocumentProperties("Revision Number") = regex.Execute(oCellObservaciones.Value).item(0).submatches(0)
		End If
	End Function

	Function RenombrarFichero (strDestFN)
	' ESTA FUNCION RENOMBRA EL FICHERO DESDE EXCEL, Y REQUIERE ABRIR EL FICHERO, para cambiar ciertas propiedades (tags, titulo, etc)
		' Por si acaso, me aseguro de que abra el fichero, y que oXlWorkBook tenga una referencia a él (permitiría usar la función, aunque se pierda la referencia!)
		If m_ExcelFMFile Is Nothing Then
			Set m_ExcelFMFile = m_ExcelApp.OpenFile(strXLSXPath, False, False)
		End If
		
		UpdateTags
		
		If Not m_ExcelFMFile.SaveAs(strDestFN, xlOpenXMLWorkbook) Then Exit Function
		
		' El SaveAs automáticamente actualiza las referencias internas
		' Cerrar la referencia antigua
		m_ExcelApp.CloseFile strDestFN, False
			
		bSave = False
		RenombrarFichero = True
			
		' Borrar archivo anterior si existe
		If fso.fileExists(strXLSXPath) Then 
			fso.deleteFile strXLSXPath 
		Else 
			Stop 
		End If
			
		' Actualizar referencias
		strXLSXPath = strDestFN
	End Function 

	Function generarInfoOfergas(strOfergas, oABCGas_XLS_LimPot)
		Dim PotenciakWMax, ofout
		PotenciakWMax = PotenciakW
		' TENGO EN CUENTA QUE PARA ESTABLECER LA POTENCIA SE USA OTRO CALCULO!!:
		If Not IsEmpty (oABCGas_XLS_LimPot) Then PotenciakWMax = oABCGas_XLS_LimPot.PotenciakW
		' genera un fichero TXT con la info para ofergas, de momento (la idea es generar directamente una oferta en Ofergas... o donde sea)
		Stop
		Set ofout = fso.OpenTextFile (strOfergas,8)
		Dim oCellGasName
		ofout.Writeline ("Datos Gen:")
		ofout.Writeline (vbTab & "Cliente:" & vbTab & oCellCliente.Value)
		ofout.Writeline (vbTab & "Proyecto (Usuario final/observacs):" & vbTab & oCellProyecto.Value)
		ofout.Writeline (vbTab & "Num Calculo:" & vbTab & oCellCalculo.Value)
		strTmp = ""
		For Each oCellGasName In oDicGasComp
			strTmp = strTmp & getGasName(oCellGasName) & " "
		Next
		ofout.Writeline (vbTab & "Gas(es):" & vbTab & strGasType & " (" & Trim(strTmp) & ")")
		ofout.Writeline (vbTab & "Observaciones (añadir comentarios de precios):" & vbTab & oCellObservaciones.Value)
		
		Stop ' pte de poner: SI EL COMPRESOR LLEVA GAS SO2, SE SUELE HACER CON CILINDROS LUBRICADOS!!!
		
		ofout.WriteBlankLines(1)
		ofout.Writeline (vbcrlf & "Cabezal, Etapas y Cilindros --> modelo:" & strModelName)
		If ncils = 1 And oDicStages.Count < 2 Then ofout.Writeline ("el compresor es MANCO!!, ojo a requisitos" & vbCrLf & "*****************************************")
		If InStr (strModelName,"HG6") > 0 Then ofout.Writeline ("ESTE COMPRESOR, " & strModelName & ", NO SE FABRICA, NO EXISTE MODELO HG6, debería calcularse como un HP4!!!" & vbCrLf & "*****************************************")
		If InStr (strModelName,"HX6") > 0 Then ofout.Writeline ("OJO, NO SE HA HECHO NINGUN COMPRESOR " & strModelName & ", en HX sólo se ha hecho un HX2, para REPSOL" & vbCrLf & "*****************************************")
		If bAire Then ofout.Writeline ("Conviene asegurarse de que el compresor " & strModelName & ", que es de aire, NO SE PUEDA OFERTAR COMO PLATAFORMA LP, de máquina estándar (llegan hasta 1000 rpm), o como SYNCRO." & vbCrLf & "*****************************************")
		If Not bCompetitivosPorPotencia Then ofout.Writeline ("NO SOMOS COMPETITIVOS en las condiciones de calculo de '" & fso.GetBaseName(strXLSXPath) & "', el compresor es de MUY baja potencia: " & PotenciakW & " kW." & _
				"*** CONVENDRIA CONSIDERAR LAS PLATAFORMAS ""V"" Y/O ""X"", que son DE SIMPLE EFECTO, aunque más proclives a FUGAS, e está estudiando añadirles un sistema de RECUPERACION DE FUGAS... Si no, DECLINAR OFERTA" & vbCrLf & "*****************************************")
		strTmp = ""
		
		'=========================================================================
		MsgIE ("<font color=blue><u>CILINDROS, PRESIONES Y MATERIALES</u></font>")
		'SEGUIR AQUI
		' para comprobar las PRESIONES LIMITE EN LOS CILINDROS: me aseguro de que esten calculadas, a partir de la tabla de excel de cilindros.
		' debería hacerlo gas_vbnet, mas preciso que el excel de los cilindros...
		Call getCylindersMaterials_Limits()
		Dim c
		For Each c In oDicStages
			Set oABCGas_XLS_Stage = oDicStages(c)
			MsgIE ("Etapa: " & c + 1 & ". Cilindro de diametro "& iEtapaDiam(c+1) & ", trabaja a presión de " & oABCGas_XLS_Stage.Stage_Pout & " bar.")
			If IsEmpty (oABCGas_XLS_Stage.cilPressureLimit) Then
				MsgIE ("NO se ha encontrado ningún cilindro de dimensiones estándar, que aguante la presión de la etapa,  habría que fabricar el cilindro A MEDIDA (posiblemente forjado)")
			Else
				MsgIE ("podría fabricarse COMO EL SIGUIENTE CILINDRO ESTANDAR: Material del cilindro: " & oABCGas_XLS_Stage.cilMaterial)
				MsgIE ("Limite de presión soportado por el cilindro: " & oABCGas_XLS_Stage.cilPressureLimit)
			End If
			' forjado (y encamisado)
			If bEtapaReqForjado(c + 1)	Then MsgBox ("la etapa " & c+1 & " requiere CILINDROS FORJADOS, que tendrán que ir encamisados (por superarse la presión de 80 bar, o por tener diam < 75, o porque son plataforma HP, o HX, ...)" & vbCrLf & _
					"- el forjado, suele ser acero al carbono o baja aleación (AISI 4130, 4140, 4340), que no es ideal para contacto directo con el émbolo por desgaste y fricción." & vbCrLf & _
					"- El liner permite usar fundición perlítica o acero aleado nitrurado que resiste mucho mejor")	
		Next

		MsgBox ("RESPECTO A MATERIALES DE CILINDROS, procesos de fabricación, y encamisados:" & vbCrLf & _
			"- el forjado, suele ser acero al carbono o baja aleación (AISI 4130, 4140, 4340), que:" & vbCrLf & _
			"   ** es el mas resistente, por encima de - fundicion gris o nodular, - acero fundido, o - acero laminado o mecanizado" & vbCrLf & _
			"   ** aguanta mejor los gases corrosivos: MUY BAJA POROSIDAD" & vbCrLf & _
			"   ** no es ideal para contacto directo con el émbolo por desgaste y fricción." & vbCrLf & _
			"  En el caso de 'GASES ACIDOS', SH2, CO2, ... se usa, por encima de este, el 'acero recubierto', que es forjado, y recubierto de CrNi o inconel" & vbCrLf & _
			"- el componente forjado, NO requiere ir encamisado... pero normalmente lo va, para proteger EL FORJADO, no la camisa!!: en lugar de desgastar el forjado, se hace sufrir a la camisa, que luego se cambia" & vbCrLf & _
			"- El liner permite usar fundición perlítica o acero aleado nitrurado que, frente al acero al carbono, resiste mucho mejor")	
		'=========================================================================

		nEtapa = 0
		For Each oABCGas_XLS_Stage In oDicStages.Items
			nEtapa = nEtapa + 1
			If oABCGas_XLS_Stage.Stage_Pout.value > 42 Then
				If strTmp <> "" Then strTmp = strTmp & ", "
				strTmp = strTmp & nEtapa
			End if
		Next
		If strTmp <> "" Then strTmp = "en etapas " & strTmp Else strTmp = "NO" End If
		ofout.Writeline (vbTab & "Empaquetadura refrigerada:" & vbTab & strTmp)
		strTmp = ""
		If bAire Then strTmp = "PET"
		If bHCs Then strTmp = "T2"
		If bCO2 Then strTmp = "CO2"
		If bH2 Or bN2 Then strTmp = "CPI"
		ofout.Writeline (vbTab & "Segmentos:" & vbTab & strTmp)
		ofout.Writeline (vbTab & "Bloque SAS:" & vbTab & bATEX_Inflamable)
		
		ofout.WriteBlankLines(1)
		ofout.Writeline (vbcrlf & "Modelo:" & strModelName)
		ofout.Writeline (vbTab & "Transmisión - Potencia:" & vbTab & PotenciakWMax & " kW")
		ofout.Writeline (vbTab & "Transmisión - RPMs:" & vbTab & RPM)
		strTmp = RPM
		If strTmp < 500 Then
			strTmp = strTmp & " (permite CORREAS; pueden ser especiales, Predator; o llevar TENSOR DE CORREAS Overly Hautz)"
			If bATEX_Inflamable Then strTmp = Replace(strTmp,"permite CORREAS; ","permite CORREAS; AÑADIR SENSOR DE TEMPERATURA, para ATEX!!; ")
		End if
		If strTmp > 200 And strTmp < 1100 Then
			strTmp = strTmp & " (permite REDUCTORA)"
			If bATEX_Inflamable Then strTmp = Replace(strTmp,"(permite REDUCTORA)","(permite, Y ES MEJOR USAR, REDUCTORA - POR SER ATEX; o acoplamiento magnetico, motor presurizado Ex p, etc)")
		End if
		For np = 4 To 14 Step 2
			If strTmp < 120 * 50 / np And strTmp > 0.95 * (120 * 50 / np) Then strTmp = strTmp & " (permite ACOPLAMIENTO DIRECTO 'en Europa', a 50 Hz)" : Exit For
			If strTmp < 120 * 60 / np And strTmp > 0.95 * (120 * 60 / np) Then strTmp = strTmp & " (permite ACOPLAMIENTO DIRECTO 'en USA', a 60 Hz)" : Exit For
		Next
		ofout.Writeline (vbTab & "Transmisión - RPMs:" & vbTab & strTmp)
		If bATEX_Inflamable Then ofout.Writeline (vbTab & "Asegurarse de MARCAR ATEX en la seleccion de transmisión (además de reductora, o acoplamientos especiales)")
		
		strTmp = ""
		If bH2 Then strTmp = "OJO, HIDROGENO --> asegurarse de que cada refrigerador sale al menos por 20000 euros"
		If oCell_Compressor_Serie.Value = "HX" Then strTmp = "OJO, plataforma HX --> asegurarse de que cada refrigerador sale al menos por 1.5!! el HP; unos 30-50000 por refrigerador"
		If oCell_Compressor_Serie.Value = "HX" And bH2 Then strTmp = "OJO, plataforma HX E HIDROGENO --> asegurarse de que cada refrigerador sale al menos por 1.5!! el HP; unos 50000 por refrigerador"
		ofout.WriteBlankLines(1)
		ofout.Writeline (vbcrlf & "Refrigeradores:" & vbTab & strTmp)
		If bAire Then ofout.Writeline (vbTab & "¡OJO!, poner el material en COBRE, es Aire!")
		nEtapa = 0
		For Each oABCGas_XLS_Stage In oDicStages.Items
			strTmp = ""
			nEtapa = nEtapa + 1
			Select Case oABCGas_XLS_Stage.oCell_Stage_CoolerSize.value
				Case "RH: 1","RH: 0","RH:-1"
					strTmp = strTmp & "RH-85-L"
				Case "RH: 2"
					strTmp = strTmp & "RH-109-L"
				Case "RH: 3"
					strTmp = strTmp & "RH-127-L"
				Case "RH: 4"
					strTmp = strTmp & "RH-187-L"
				Case "RH: 5"
					strTmp = strTmp & "RH-253"
				Case "RH: 6"
					strTmp = strTmp & "RH-309"
				Case "RH: 7"
					strTmp = strTmp & "RH-7"
				Case "RH: 8"
					strTmp = strTmp & "RH-8"
				Case "RH: 9"
					strTmp = strTmp & "RH-9"
			End Select
			ofout.Writeline (vbTab & "Etapa " & nEtapa & ":" & vbTab & strTmp)
			ofout.Writeline (vbTab & vbTab & " PN: " & oABCGas_XLS_Stage.Stage_Pout * 1.1 & " (valor de tarado de valv seg salida).")
			If bNACE_Corrosivo Then ofout.Writeline (vbTab & vbTab & " Material: INOX") Else ofout.Writeline (vbTab & vbTab & " Material: GALVANIZADO") End If
		Next
		ofout.Writeline (vbTab & "Asegurarse de que el refrigerador final salga por 15-20000 euros")
		If bH2 Then
			ofout.Writeline (vbTab & "lleva H2, --> PONER UN REFRIGERADOR FINAL ADICIONAL, ""bypass-cooler"", DUPLICANDO PRECIO, o con 1 etapa más")
		Else
			ofout.Writeline (vbTab & "si se pone bypass (por necesidad de REGULACION, o en el ARRRANQUE), PONER UN REFRIGERADOR FINAL ADICIONAL, ""bypass-cooler"", DUPLICANDO PRECIO, o con 1 etapa más")
		End If
		ofout.Writeline (vbTab & "SI EL CLIENTE PIDE que 'el gas circule por FUERA de los tubos, y el AGUA POR DENTRO' (caso de aguas sucias, agua de planta en circuito abierto, torre de refrigeracion abierta) --> habría que añadir en OTROS, una ""CARCASA""...")
		If bH2O Then
			ofout.Writeline (vbTab & "hay H2O --> potencial formación de CONDENSADOS. Debería haber SEPARADORES, ")
		End If
		If bH2 Then
			ofout.Writeline (vbTab & "Calderería: MARCAR TODO, por ser H2 (sobreespesor, radiografiado, ASME y Sello U?)")
		ElseIf bNACE_Corrosivo Then 
			ofout.Writeline (vbTab & "Calderería: por ser Corrosivo, MARCAR sobreespesor, y radiografiado (OJO; si el CO2 lleva METANOL, no haría falta ninguna, porque el metanol NEUTRALIZA EL EFECTO DEL AGUA, DE FORMA QUE NO SE FORME ELECTROLITO ACUOSO)")
		End if
		ofout.Writeline (vbTab & vbTab & "SI EL COMPRESOR VA A USA, poner tb ASME y Sello U!!!")
		ofout.Writeline (vbTab & "Valvulas de seguridad: Poner " & oDicStages.Count + 1)
		
		ofout.WriteBlankLines(1)
		ofout.Writeline (vbcrlf & "Calderines antipulsadores: predeterminado para " & strModelName & " (incluyen los 'SEPARADORES por defecto')")
		ofout.Writeline (vbTab & "SI EL CLIENTE PIDE expresamente SEPARADORES, uno por etapa, o KO-DRUMs (== separador con brida adicional, para 'esponja'; NORMALMENTE SOLO UNO, A LA ENTRADA) --> hay que dimensionarlos... el KO-DRUM puede salir por 20-30k€ -- ver precios de DEPOSITOS DE ENTRADA / SALIDA")
		For Each oABCGas_XLS_Stage In oDicStages.Items
			nEtapa = nEtapa + 1
			ofout.Writeline (vbTab & "Etapa " & nEtapa & ", ENTRADA:" & vbTab & "(EL VOLUMEN SALE DE gas_vbnet, en ""antipul"")")
			ofout.Writeline (vbTab & vbTab & " PN: " & oABCGas_XLS_Stage.Stage_Pout * 1.1 & " (valor de tarado de valv seg salida).")
			If bNACE_Corrosivo Then ofout.Writeline (vbTab & vbTab & " Material: INOX") Else ofout.Writeline (vbTab & vbTab & " Material: GALVANIZADO") End If

			ofout.Writeline (vbTab & "Etapa " & nEtapa & ", SALIDA:" & vbTab & "(EL VOLUMEN SALE DE gas_vbnet, en ""antipul"")")
			ofout.Writeline (vbTab & vbTab & " PN: " & oABCGas_XLS_Stage.Stage_Pout * 1.1 & " (valor de tarado de valv seg salida).")
			If bNACE_Corrosivo Then ofout.Writeline (vbTab & vbTab & " Material: INOX (opcionalmente, MAS BARATO, GALVANIZADO)") Else ofout.Writeline (vbTab & vbTab & " Material: GALVANIZADO") End If
		Next
		If bH2 Then
			ofout.Writeline (vbTab & "Calderería: MARCAR TODO, por ser H2 (sobreespesor, radiografiado, ASME y Sello U?)")
		ElseIf bNACE_Corrosivo Then 
			ofout.Writeline (vbTab & "Calderería: por ser Corrosivo, MARCAR sobreespesor, y radiografiado (OJO; si el CO2 lleva METANOL, no haría falta ninguna, porque el metanol NEUTRALIZA EL EFECTO DEL AGUA, DE FORMA QUE NO SE FORME ELECTROLITO ACUOSO)")
		End if
		ofout.Writeline (vbTab & vbTab & "SI EL COMPRESOR VA A USA, poner tb ASME y Sello U!!!")
		strTmp = ""
		
		ofout.WriteBlankLines(1)
		ofout.Writeline (vbcrlf & "Instrumentación --> para modelo:" & vbTab & strModelName)
		ofout.Writeline (vbTab & "Como TRANSDUCTOR DE TEMPERATURA poner ALGUNO DE 1200 EUROS… Como TRANSDUCTOR DE PRESIÓN, poner alguno de 1500 EUROS.")
		ofout.Writeline (vbTab & "Poner un SENSOR DE CAIDA DE VASTAGO por cada vastago --> poner " & ncils & " sensores" & vbCrLf)
		' SIEMPRE es posible la regulacion 0-100.
		' strTmp = "0-100"
		For iEtapa = 1 to oDicStages.Count
			' si solo hay un cilindro tandem, SOLO ES POSIBLE la regulación anterior... etc
			If bEtapaIsTandem(iEtapa) And  iEtapaNCils(iEtapa) = 1 Then strTmp = "0-100" : Exit For
			' si hay CILINDROS PARES en cada etapa, es posible 0-25-50-75-100
			If iEtapaNCils(iEtapa) Mod 2 = 0 And (IsEmpty(strTmp) Or strTmp = "0-25-50-75-100") Then strTmp = "0-25-50-75-100"
			If Not IsEmpty(strTmp) And strTmp <> "0-25-50-75-100" Then strTmp = "0-50-100"
		Next
		If IsEmpty (strTmp) then
			ofout.Writeline (vbTab & "SI HAY REGULACION 0-50-100, 0-33-66-100, ... poner una ELECTROVALVULA E.V.REGULACION")
		Else
			ofout.Writeline (vbTab & "Es posible la regulacion " & strTmp & " --> poner una ELECTROVALVULA E.V.REGULACION " & strTmp)
		End If
		
		
		' si hay condensados, ** HAY QUE PONER LOS ELEMENTOS PARA SU RECOGIDA!!! **
		stop
		oCell_OUTProcess_Condens.Value

' COSAS PENDIENTES, CON OFERGAS:
' - PANEL DE n2, IMPLICA "bridas para barrido de gas", que van en budget, pero NO estan en ofergas... etc!!!
' - SI LA TEMPERATURA DEL AGUA ES MENOR A 16ºC, HAY QUE MONTAR VALVULAS TERMOSTATICAS (doy por hecho que MAS CARAS)
		
		' otras consideraciones, a TENER EN CUENTA EN OFERGAS:
		If bC2H4 Then MsgBox ("OJO, etileno: los acumuladores / antipulsadores deben ser GALVANIZADOS; las TUBERIAS DE REFRIGERACION deben ser INOX")
		If bH2 Then MsgBox ("OJO, H2: los componentes deben ser ATEX, y lleva DISTANCIADOR TIPO C. lleva SEGMENTOS ESPECIALES, todo T2  / CPI. Y siempre se suele poner BYPASS, con H2")
		If bCO2 Then MsgBox ("OJO, CO2: lleva SEGMENTOS ESPECIALES, Co2; las TUBERIAS DE REFRIGERACION deben ser INOX")
		If bN2 Then MsgBox ("OJO, N2: lleva SEGMENTOS ESPECIALES, todo T2  / CPI; las TUBERIAS DE REFRIGERACION deben ser INOX ( y debe girar a bajas RPM; minimo 375 RPM para garantizar lubricac cigueñal)")
		If bNACE_Corrosivo Then MsgBox ("ojo!!!, tener en cuenta que EL COMPRESOR TIENE QUE SER NACE: proteger calderería de la corrosión!")

		If oCell_Compressor_Serie = "HX" Then MsgBox ("El compresor es de PLATAFORMA HX, Y 'por defecto' LLEVA DISTANCIADOR TIPO D (que con H2 sería tipo C) / Bloque SAS, ** A MARCAR EN OFERGAS!!")

		' refrigeradores, y caudal a considerar para el aero o torre: QUE NO SEA EXCESIVO...
		'...
		' en funcion de las RPMs, determinar si conviene VARIADOR (solo PARA VUELTAS ALTAS, > 370), CORREAS (PARA POTENCIA < 600-800 kW), REDUCTORA, acople directo (ESPECIFICANDO EL NUMERO DE POLOS, para "tensiones estandar" 480V, 400V, ...)
		If RPM > 370 * 1.3 Then
			MsgBox ("El compresor SE PODRÍA REGULAR MEDIANTE VARIADOR, en un RANGO DE REGULACION desde el 100% de velocidad, hasta el " & Int(370 * 100 / RPM) & "%, correspondiente a la mínima velocidad a que puede funcionar el compresor, 370 RPM. ** EL PRECIO DEL VARIADOR, SE INCLUYE APARTE EN OFERGAS, EN ""OTROS!!!"" (puede ser 'estandar', o 'REGENERATIVO' (que regenera energía del par motor del compresor))")
		Else
			MsgBox ("A las RPM que funciona este compresor, " & RPM & ", NO TIENE SENTIDO USAR UN VARIADOR, el rango de regulación sería MUY BAJO (se regularía MEDIANTE BYPASS, mejor)")
		End if
		If PotenciakWMax < 600 Then
			MsgBox ("Potencia " & oCell_OUTProcess_Winst & " INFERIOR a 600 kW." & vbCrLf & _
				"- El compresor PODRIA ACOPLARSE AL MOTOR CON CORREAS, que para un motor de induccion de 1500 rpm deberían corresponder a una RELACION DE TRANSMISION, i=" & Round (1500 / RPM,2) & vbCrLf & _
				"- para estas potencias, suficientemente BAJAS, PODEMOS OFERTAR TANTO PANEL DE FUERZA COMO DE CONTROL, tenerlo en cuenta EN OFERGAS (alli YA ESTAN INCLUIDOS los precios CON ARRANCADORES, sea soft start o estrella triangulo...), y EN EL BUDGET")
		Else
			MsgBox ("Potencia " & oCell_OUTProcess_Winst & " SUPERIOR a 600 kW." & vbCrLf & _
				"- El compresor DEBERIA ACOPLARSE AL MOTOR MEDIANTE REDUCTORA, que para un motor de induccion de 1500 rpm deberían corresponder a una RELACION DE TRANSMISION, i=" & Round (1500 / RPM,2) & vbCrLf & _
				"- para estas potencias no OFERTAMOS PANEL DE FUERZA, si acaso SOLO EL DE CONTROL, tenerlo en cuenta EN OFERGAS, y EN EL BUDGET. Si el cliente PIDIESE EL ARRANCADOR, PODRÍAMOS OFERTARLO HASTA UN LIMITE DE POTENCIA razonable, NO llegamos a 1 MW!!! - caso de Pequiven")
		End If
		bDone = false
		For Each npolos In Array (4,6,8,10,12,14)
			For Each freq In Array (50,60)
				If 120 * freq / npolos < RPM And 120 * freq / npolos > RPM * 0.96 Then
					MsgBox ("El compresor tambien permite un ACOPLAMIENTO DIRECTO a " & RPM & " rpm, con un DESLIZAMIENTO ACEPTABLE, <= del 4%, ** PARA UNA FRECUENCIA DE RED DE " & freq & " Hz ** (*** ASEGURARSE DE QUE ESA FRECUENCIA SEA DE APLICACION EN EL PAIS EN CUESTION - 60 Hz tipica en EEUU, Latinoamérica, y algún país asiático; y 50 Hz en Europa y casi todo el resto del mundo***)")
					bDone = true
					Exit for
				End If
			Next
			If bDone Then Exit For
		Next

		ofout.Close
	End Function
End Class

Class cABCGas_XLS_Stage
	Dim oCell_Stage_CilsDiam, oCell_Stage_Flow, oCell_Stage_Pout, oCell_Stage_Taspirac, oCell_Stage_Tescape, oCell_Stage_TescapeAdiab, oCell_Stage_CompRatio
	Dim oCell_Stage_ComprStress, oCell_Stage_TensStress, oCell_Stage_VolumeGen, oCell_Stage_NuVolum, oCell_Stage_FillCoef, oCell_Stage_DeadVolume, oCell_Stage_MinVolum
	Dim oCell_Stage_ValveSection, oCell_Stage_ValveSpeed, oCell_Stage_NuValve, oCell_Stage_GammaAdiabIdx, oCell_Stage_DiagrPower, oCell_Stage_Regulation
	Dim oCell_Stage_CoolingWater, oCell_Stage_NrCoolers, oCell_Stage_CoolerSize, oCell_Stage_TLR, oCell_Stage_Pdrop, oCell_Stage_WaterFlow, oCell_Stage_CondensateWaterFlow
	Dim oCell_Stage_Gas_Zin, oCell_Stage_Gas_Zout
	' los dos siguientes atributos se obtienen del fichero de "CILINDROS, limites de presión y materiales" (en la funcion getCylindersMaterials_Limits)
	Dim cilMaterial, cilPressureLimit
	Private regex

	Private Sub Class_Initialize()
	    Set regex = New RegExp
		regex.Global = True : regex.IgnoreCase = True : regex.multiline = False
	End Sub
	Public Function numCils
		If IsEmpty (oCell_Stage_CilsDiam) Then Exit Function
		regex.Pattern = "(\d)\s+x\s+(\d+)([tc]*)"
		numCils = regex.Execute (oCell_Stage_CilsDiam.Value).Item(0).Submatches(0)
	End Function
	Public Function diamCils
		If IsEmpty (oCell_Stage_CilsDiam) Then Exit Function
		regex.Pattern = "(\d)\s+x\s+(\d+)([tc]*)"
		diamCils = regex.Execute (oCell_Stage_CilsDiam.Value).Item(0).Submatches(1)
	End Function
	Public Function bTandem
		If IsEmpty (oCell_Stage_CilsDiam) Then Exit Function
		regex.Pattern = "(\d)\s+x\s+(\d+)([tc]*)"
		bTandem = InStr(LCase(regex.Execute (oCell_Stage_CilsDiam.Value).Item(0).Submatches(2)),"t") > 0
	End Function
	Public Function bCamisa
		If IsEmpty (oCell_Stage_CilsDiam) Then Exit Function
		regex.Pattern = "(\d)\s+x\s+(\d+)([tc]*)"
		bCamisa = InStr(LCase(regex.Execute (oCell_Stage_CilsDiam.Value).Item(0).Submatches(2)),"c") > 0
	End Function
	Public Function Stage_Pout
		If IsEmpty (oCell_Stage_Pout) Then Stop : Exit Function
		Stage_Pout = CDbl (oCell_Stage_Pout)
	End Function
End Class

Class cLimitsFeatsReqs
	' limites
	Dim iOversizekW, iMinDiam, iMaxDiam, iMaxRPM, iMaxPistonSpeedMS, iMaxLoad, PMaxDesign
	Dim maxcomprel,mincomprel
	' requisitos
	Dim bBombaAceiteAux, bBielasForjadas
	' caracts predefinidas (iCrankCaseLtr: litros de aceite que lleva en carter)
	Dim iStroke, iRodDiam, iCrankCaseLtr
	Private Sub Class_Initialize
		iOversizekW = 0
		iMinDiam = 0
		iMaxDiam = 6000
		iMaxRPM = 1000
		bBombaAceiteAux = False
		bBielasForjadas = False
		iStroke = 0
		iRodDiam = 0
		iMaxPistonSpeedMS = 0
		iMaxLoad = 60000
		PMaxDesign = 500 ' POR DEFECTO CREO QUE ES 300 bar, EN TODAS LAS PLATAFORMAS
		iCrankCaseLtr = 0
	End Sub
	Function Init (oABCGas_XLS)
		Set Init = Me
		Select Case oABCGas_XLS.oCell_Compressor_Serie
			Case "HA":
				Select Case oABCGas_XLS.nCils
					Case 1:  iOversizekW = 50 : iCrankCaseLtr = 20
					Case 2:  iOversizekW = 110 : iCrankCaseLtr = 20
					Case 4:  iOversizekW = 250 : iCrankCaseLtr = 40
					Case 6:  iOversizekW = 400 : bBombaAceiteAux = True  : iCrankCaseLtr = 60
					Case Else: Stop ' NO ES POSIBLE UN COMPRESOR CON ESTE NUMERO DE CILINDROS!!
				End Select
				iMinDiam = 45
				iMaxDiam = 310
				iMaxRPM = 800
				bBielasForjadas = True
				iStroke = 150
				iRodDiam = 32
				iMaxPistonSpeedMS = 3.75
				iMaxLoad = 3000
			Case "HG":
				Select Case oABCGas_XLS.nCils
					Case 1,2,4:  iOversizekW = 800
					Case Else: Stop ' NO ES POSIBLE UN COMPRESOR CON ESTE NUMERO DE CILINDROS!!
				End Select
				iMaxLoad = 5000
			Case "HP":
				Select Case oABCGas_XLS.nCils
					Case 1:  iOversizekW = 250 : iCrankCaseLtr = 150
					Case 2:  iOversizekW = 465 : iCrankCaseLtr = 150
					Case 4:  iOversizekW = 930 : iCrankCaseLtr = 300
					Case 6:  iOversizekW = 1400 : iCrankCaseLtr = 450
					Case Else: Stop ' NO ES POSIBLE UN COMPRESOR CON ESTE NUMERO DE CILINDROS!!
				End Select
				iMaxRPM = 650
				bBombaAceiteAux = True
				iStroke = 150
				iRodDiam = 32
				iMaxPistonSpeedMS = 4
				iMaxLoad = 10000
			Case "HX":
				Select Case oABCGas_XLS.nCils
					Case 1,2,4,6:  iOversizekW = 3000
					Case Else: Stop ' NO ES POSIBLE UN COMPRESOR CON ESTE NUMERO DE CILINDROS!!
				End Select
				iMinDiam = 120
				iMaxDiam = 700 ' oficialmente es de 580
				iMaxRPM = 500
				bBombaAceiteAux = True
				iMaxLoad = 24000
			Case "X":
				Stop
			Case Else
		End Select
		maxcomprel = 4
		mincomprel = 2
		With oABCGas_XLS
			Select Case true
				Case .bH2
					If maxcomprel > 3.2 Then maxcomprel = 3.2
				Case .bHCs,.bO2,.bN2,.bCO,.bH2O,.bAR,.bNH3,.bSH2,.bC2H4
				Case .bCO2
					maxcomprel = 10
			End Select
		End With
	End Function
End Class
