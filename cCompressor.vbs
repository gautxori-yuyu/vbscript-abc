Class cCompressor
	Dim oDicCalcs,oDicValoracs,oDicBudgets,oDicQuotations
	
	Private Sub Class_Initialize
		Set oDicCalcs = CreateObject("scripting.dictionary") ' es un arbol, {strCalcNum,{strCalcOpc,{fich,oABCGas_XLS}}}
		
		Set oDicValoracs = CreateObject("scripting.dictionary") 
		oDicValoracs.CompareMode = 1
		Set oDicBudgets = CreateObject("scripting.dictionary") 
		oDicBudgets.CompareMode = 1
		Set oDicQuotations = CreateObject("scripting.dictionary") 
		oDicQuotations.CompareMode = 1
	End Sub
	
	' el modelo de compresor SE DETERMINA UNA SOLA VEZ, al primer item que se añada, que lo defina. Y tiene que ser 'consistente', TODOS LOS ITEMS
	' en el compresor tienen que casar con ese modelo (esa validación de consistencia ¿se hace aqui, o deberia hacerse en un nivel superior?: para chatgpt)
	Private strModel_
	Public Property Get strModel
		If Not IsEmpty (strModel_) Then strModel = strModel_ : Exit Property
		
		' la definición de strModel_ la hacen las funciones con las que se añaden items al compresor
	End Property
	
	Public Function bAddCalc (oABCGas_XLS)
		bAddCalc = False
		If IsEmpty (oABCGas_XLS.strCalcOpc) Then Exit Function ' presupone que NO se ha inicializado el objeto de calculo
		If Not IsEmpty (strModel_) Then
			' comprueba que el modelo que define el calculo, casa con el de este compresor;
			' si no, DEVUELVE FALSE, para indicar que el calculo NO SE PUEDE AÑADIR
			If oABCGas_XLS.strModelName <> strModel_ Then Exit Function
		Else
			strModel_ = oABCGas_XLS.strModelName
		End If
		Dim strNumCalc
		strNumCalc = Left (oABCGas_XLS.strCalcOpc,InStr(oABCGas_XLS.strCalcOpc,"_")-1)
		If Not oDicCalcs.Exists (strNumCalc) Then oDicCalcs.Add strNumCalc, CreateObject("scripting.dictionary") : oDicCalcs(strNumCalc).CompareMode = 1
		If Not oDicCalcs(strNumCalc).Exists (oABCGas_XLS.strCalcOpc) Then oDicCalcs(strNumCalc).Add oABCGas_XLS.strCalcOpc, CreateObject("scripting.dictionary") : oDicCalcs(strNumCalc)(oABCGas_XLS.strCalcOpc).CompareMode = 1
		If Not oDicCalcs(strNumCalc)(oABCGas_XLS.strCalcOpc).Exists (oABCGas_XLS.strXLSXPath) Then
			oDicCalcs(strNumCalc)(oABCGas_XLS.strCalcOpc).Add oABCGas_XLS.strXLSXPath, oABCGas_XLS
		Else
			' ya existía una entrada para ese calculo, en este compresor !!?? (posiblemente sea UNA REVISION, o fichero OLD, para el mismo calculo..: PTE DE DECIDIR COMO PROCESARLA
			MsgLog ("<b>FICHERO DE CALCULO YA ESTABA AÑADIDO!!!</b>")
			Stop
		End If
		bAddCalc = true
	End Function
	
	Public Function bAddValorac (oOferGas_XLS)
		' PTE DE IMPLEMENTAR EL USO DE LOS OBJETOS oOferGas_XLS...
		Stop
	End Function
End Class

