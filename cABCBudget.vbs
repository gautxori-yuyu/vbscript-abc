Class cBudget
	Private regex, WshShell
	Dim strXLSXPath,bSave

	Private m_ExcelApp
	Private m_ExcelFMFile  ' cExcelFile wrapper del archivo principal
	Private oXlEntrySheet

	Private Sub Class_Initialize()
	    Set regex = New RegExp
		regex.Global = True : regex.IgnoreCase = True : regex.multiline = False
		Set WshShell = CreateObject("Wscript.Shell")

		Set m_ExcelApp = Nothing
		Set m_ExcelFMFile = Nothing
	End Sub
	
	Private Sub Class_Terminate()
		'Stop : Call CloseWorkBook() ' NO DEBERIA NECESITAR ESTA LLAMADA, EN TANTO QUE USE EL EXCELMANAGER...
	End Sub
	
	Private Property Get objExcel
		If m_ExcelApp Is Nothing Then
			Err.Raise 91, "cABCBudget", "ExcelApp not initialized. Call Init() first."
		End If
		Set objExcel = m_ExcelApp.Application
	End Property
	
	' =============================================
	' INICIALIZACIÓN Y CIERRE
	' =============================================
	
	Public Function Init(ExcelApp, strXLSXPath_)
		Set Init = Me
		strXLSXPath = strXLSXPath_
		Set m_ExcelApp = ExcelApp
		
		Set m_ExcelFMFile = m_ExcelApp.OpenFile(strXLSXPath, False, False)
		Set oXlEntrySheet = m_ExcelFMFile.GetWorksheet("BUDGET_ENTRY")
	End Function
	
	Public Function CloseWorkBook()
		If Not (m_ExcelFMFile Is Nothing) Then
			If bSave Then m_ExcelFMFile.Save
			m_ExcelApp.CloseFile m_ExcelFMFile.FilePath, False
			Set m_ExcelFMFile = Nothing
		End If
	End Function

	' =============================================
	' FUNCIONES ESPECIFICAS
	' =============================================

	Public Sub setLanguage (strLang)
		Select Case strLang
			Case "ES","EN"
				oXlEntrySheet.range ("E7").Value = strLang
		End Select
	End Sub
	Public Sub InsertCGASING (strXLSCalcTecnPath)
	
	End Sub
	Public Sub overrideGeneralData (strCustomer,strProject,strQuoteNr,strDate,strContactPerson)
		If strCustomer <> "" Then oXlEntrySheet.range ("E15").Value = strCustomer
		If strProject <> "" Then oXlEntrySheet.range ("E16").Value = strProject
		If strQuoteNr <> "" Then oXlEntrySheet.range ("E17").Value = strQuoteNr
		If IsDate (strDate) Then oXlEntrySheet.range ("E19").Value = strDate
		
		' Hay que asegurarse de que cumpla las REGLAS DE VALIDACION...
		If strContactPerson <> "" Then oXlEntrySheet.range ("E20").Value = strContactPerson
	End Sub
	Public Sub overrideCGASINGWorkingConds (strModel,strFlow,strGas,strSuctPres,strDischPres,strLUBED,strPower,strRPM)
		' fija, en BUDGET_ENTRY, las filas 22 a 29, *********** Y TAMBIEN LAS FILAS 31 a 33 **********
		If strModel <> "" Then oXlEntrySheet.range ("E22").Value = strModel
		If strFlow <> "" Then oXlEntrySheet.range ("E23").Value = strFlow
		If strGas <> "" Then oXlEntrySheet.range ("E24").Value = strGas
		If strSuctPres <> "" Then oXlEntrySheet.range ("E25").Value = strSuctPres
		If strDischPres <> "" Then oXlEntrySheet.range ("E26").Value = strDischPres
		If strRPM <> "" Then oXlEntrySheet.range ("E29").Value = strRPM
		
		' Hay que asegurarse de que cumplan las REGLAS DE VALIDACION...
		If strLUBED <> "" Then oXlEntrySheet.range ("E27").Value = strLUBED
		If strPower <> "" Then oXlEntrySheet.range ("E28").Value = strPower
		
		regex.Pattern = strModelPattern
		If Not regex.Test(strModel) Then
			WshShell.Popup "ERROR en la definición del modelo de compresor, " & strModel,8
		Else
			Dim oModelMatch
			Set oModelMatch = regex.Execute(strModel).Item(0)
			If oModelMatch.Submatches(0) <> "" Then oXlEntrySheet.range ("E32").Value = oModelMatch.Submatches(0)
			If oModelMatch.Submatches(2) <> "" Then oXlEntrySheet.range ("E33").Value = oModelMatch.Submatches(2)
			
			' Hay que asegurarse de que cumpla las REGLAS DE VALIDACION...
			If oModelMatch.Submatches(1) <> "" Then oXlEntrySheet.range ("E31").Value = oModelMatch.Submatches(1)
		End If
	End Sub
	Public Sub AddNote (strNoteText)
		' Añade una NOTA, en la hoja BUDGET_QUOTE
	
	End Sub
	Public Sub EnableCover (strPicturePath)
	
	End Sub
End Class