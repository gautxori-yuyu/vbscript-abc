Option Explicit

Function bEthernetConnected
	Dim WshShell,fso,regex, strftmp
	Set WshShell = CreateObject("Wscript.Shell")
	Set fso = CreateObject("Scripting.Filesystemobject")
	Set regex = New RegExp
	regex.Pattern = "Conectado\s+.+?\s+Ethernet"
	strftmp = fso.GetSpecialFolder(2) & "\" & fso.GetTempName
	WshShell.Run "cmd /c ""netsh interface show interface > """ & strftmp & """""",0,true
	If fso.FileExists (strftmp) then
		bEthernetConnected = regex.Test (fso.OpenTextFile(strftmp).readAll)
		fso.DeleteFile strftmp, True
	End If
End Function

Function shellrun_Error (strCmd)
	strftmp = fso.GetSpecialFolder(2) & "\" & fso.GetTempName()
	strCmd = "cmd /c ""start /wait " & strCmd & " & echo %errorlevel% > """ & strftmp & """"""
	Wshshell.Run strCmd,0,True
	If fso.FileExists (strftmp) Then
		WScript.Echo strCmd & vbCrLf & vbTab & fso.OpenTextFile(strftmp).ReadAll
		fso.DeleteFile strftmp,True
	End IF
End Function

Function bFileOpen_FromCmdLine (filename)
	'solo lo hace bien, si el nombre de fichero APARECE EN LA LINEA DE COMANDOS del proceso
    DIM objWMIService, strWMIQuery
	Set objWMIService = GetObject("winmgmts://./root/cimv2")
	' caracteres especiales de LIKE: [] - rango de caracteres; _ cualquier caracter; ^ exluye caracteres, en un rango: [^...]; % cero o más caracteres, como * de regexp
	strWMIQuery = "SELECT * FROM Win32_Process WHERE CommandLine LIKE '%" & Replace(filename,"_","[_]") & "%'"
	
	bFileOpen = objWMIService.ExecQuery(strWMIQuery).Count > 0
End Function

' --- Generic File/Folder Opener ---
Sub OpenFileOrFolder(sPath)
	MsgIE.MsgLog "Abriendo: " & sPath
	On Error Resume Next
	' Usamos explorer.exe para abrir el archivo con su aplicación predeterminada
	' o abrir la carpeta si es un directorio.
	' Las comillas dobles extra son para manejar espacios en la ruta.
	WshShell.Run "explorer.exe """ & sPath & """"
	
	If Err.Number <> 0 Then
		MsgIE.MsgDoc "Error al abrir: " & Err.Description, "error"
	Else
		MsgIE.MsgDoc "success","Solicitud de apertura enviada al sistema."
	End If
	On Error Goto 0
End Sub

Dim bQuit
Class HTMLWindow
	Public objExplorer,objDocument
	Private fLogTick
    Private m_Cuerpo,m_Log  ' [LEGACY] Mantenidas para compatibilidad
    Private m_CurrContainer   ' [LEGACY] Contenedor abierto (div interno)
	Private oDicContainerStack  ' [LEGACY]
	Private unnamedSPCount  ' [LEGACY]
	Private m_Logger  ' Instancia de HtaLogger (refactorización: composición)

	Private Sub Class_Initialize()
		fLogTick = Timer
		Set m_CurrContainer = Nothing
		Set oDicContainerStack = CreateObject("scripting.dictionary")
		unnamedSPCount = 0
		Set m_Logger = Nothing
	End Sub

	Public Function CreaVentanaHTML (strTitulo, Visible, Width, Height)
		Set objExplorer = WScript.GetObject("","InternetExplorer.Application", "IE_")
		objExplorer.Navigate "about:blank"',1 ' ABRE NUEVA PESTAÑA en el navegador:  Use 2048 for new tab, 1 for new browser.
		objExplorer.ToolBar = 0
		objExplorer.StatusBar = 0
        If Not IsEmpty (Width) Then objExplorer.Width = Width
        If Not IsEmpty (Height) Then objExplorer.Height = Height
			Dim objWMI, colItems, objItem
			Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
			Set colItems = objWMI.ExecQuery("Select * from Win32_DesktopMonitor")
			For Each objItem in colItems
				If objItem.Availability  = 3 Then
			        If IsEmpty (Width) Or IsEmpty (Height) Then
					    If IsEmpty (Width) And Not IsEmpty (objItem.ScreenWidth) Then objExplorer.Width = objItem.ScreenWidth - 20
					    If IsEmpty (Height) And Not IsEmpty (objItem.ScreenHeight) Then objExplorer.Height = objItem.ScreenHeight - 60
					End If
					objExplorer.Left = objItem.ScreenWidth - Width - 50
					objExplorer.Top = objItem.ScreenHeight - Height - 80
				    Exit For
			    End If
			Next
		objExplorer.Visible = Visible
		objExplorer.offline = True

		Do While (objExplorer.Busy)
			WScript.Sleep(100)
		Loop
		Do Until objExplorer.ReadyState = 4 : WScript.Sleep 100 : Loop
		Set objDocument = objExplorer.Document

		' REFACTORIZACIÓN: Inyectar template HTML completo (en lugar de Init_Document)
		Call InjectHTMLTemplate(strTitulo)

		' REFACTORIZACIÓN: Instanciar HtaLogger y pasarle el document de IE
		Set m_Logger = New HtaLogger
		m_Logger.SetDocument objDocument, True  ' True = usar diálogos nativos IE (window.confirm/alert)

		' LEGACY: Mantener referencias para compatibilidad (si fuera necesario)
		Set m_CurrContainer = objDocument.body
		Set m_Cuerpo = objDocument.getElementById("main")
		Set m_Log = objDocument.getElementById("log")

		Set CreaVentanaHTML = Me
	End Function

	' =============================================
	' Inyecta el template HTML completo desde archivo externo
	' @param strTitulo - Título de la ventana
	' =============================================
	Private Sub InjectHTMLTemplate(strTitulo)
		Dim fso, templatePath, htmlContent

		Set fso = CreateObject("Scripting.FileSystemObject")

		' Ruta del template (mismo directorio que el script)
		templatePath = fso.GetParentFolderName(WScript.ScriptFullName) & "\template_ie.html"

		If Not fso.FileExists(templatePath) Then
			Err.Raise 53, "HTMLWindow.InjectHTMLTemplate", "Template no encontrado: " & templatePath
		End If

		' Leer template (ANSI, codepage -2 para sistema)
		htmlContent = fso.OpenTextFile(templatePath, 1, False, -2).ReadAll

		' Inyectar en IE usando innerHTML (más robusto que document.open/write/close)
		On Error Resume Next
		objDocument.documentElement.innerHTML = htmlContent

		' Si falla innerHTML, intentar método alternativo
		If Err.Number <> 0 Then
			Err.Clear
			' Método alternativo: open/write/close con parámetros
			objDocument.open "text/html", "replace"
			objDocument.write htmlContent
			objDocument.close
		End If
		On Error GoTo 0

		' Establecer título
		objDocument.title = strTitulo
	End Sub

	' =============================================
	' [LEGACY] Método antiguo - mantenido por compatibilidad pero YA NO SE USA
	' =============================================
	Public Function Init_Document(strTitulo)
		If objExplorer Is Nothing Then
			Err.Raise 91, "HTMLWindow", "InternetExplorer.Application not initialized. Call CreaVentanaHTML() first."
		End If
        objDocument.title = strTitulo

        ' <style> en head
        Dim styleElem
        Set styleElem = objDocument.createElement("style")
        styleElem.type = "text/css"
		styleElem.innerText = "html, body {height:100%; margin:0; padding:0; overflow:hidden;} " & _
					"body {color:green; background:MintCream} " & _
					"body,TD,TR,TH{font-family: Consolas, monospace, Arial;font-size: 11pt;margin: 3px;padding:0;} " & _
					"TD{border:1;} UL{margin:0pt 12pt;padding:2 1;} " & _
					"LI{border-style:dashed;border-width:thin;border-color:dodgerblue;} " & _
					".tight p {margin: 0; line-height: 1.5; } " & _
					".spoilerHead { cursor:pointer; margin:2px 0; background:#eee; padding: 2px; border: 1px solid #999;} " & _
					".spoilerBody { margin-left:10px; padding: 2px; border-left: 2px solid #ccc;}" & _
					".bordered-table { border-collapse: collapse; width: 100%; border: 1px solid #000; }" & _
	                ".bordered-table td { border: 1px solid #000; padding: 2px; padding-left: 20px;}" & _
	                ".fixed-cell { width: 20%; }" & _
	                ".fluid-cell { width: 80%; }" & _
					"#presentacion {height:70%; overflow:auto; border-bottom:2px solid #000; padding:5px;} " & _
					"#log {height:30%; overflow:auto; color:#808000; background:Lavender;font-size: 9pt;} "
        objDocument.getElementsByTagName("head")(0).appendChild styleElem

        ' <script> en head
        Dim scriptElem
        Set scriptElem = objDocument.createElement("script")
        scriptElem.type = "text/javascript"
        scriptElem.text = "function showhide (item) {" & vbCrLf & _
                          " item.style.display = (item.style.display == 'none')? 'block':'none';" & vbCrLf & _
                          "}"
        objDocument.getElementsByTagName("head")(0).appendChild scriptElem
		
		' Crear DIV superior (presentación)
		Set m_Cuerpo = objDocument.createElement("div")
		m_Cuerpo.id = "cuerpo"
		m_Cuerpo.style.cssText = "height:70%; overflow:auto; border-bottom:2px solid #000; padding:5px;"
        m_Cuerpo.className = "tight"
		objDocument.body.appendChild m_Cuerpo
		
		' Crear DIV inferior (log de mensajes)
		Set m_Log = objDocument.createElement("div")
		m_Log.id = "log"
		m_Log.style.cssText = "height:30%; overflow:auto; padding:5px; font-family:Arial;"
        m_Log.className = "tight"
		objDocument.body.appendChild m_Log

		' de momento esta instruccion debe estar despues de crear el m_Log
		Set oCurrContainer = m_Cuerpo
	End Function
		
	' =============================================
	' [REFACTORIZADO] Métodos delegados a m_Logger
	' =============================================

	Public Sub popContainer ()
		If Not m_Logger Is Nothing Then
			m_Logger.popContainer
		Else
			' LEGACY fallback
			If oDicContainerStack.Count = 0 Then Exit Sub
			Set m_CurrContainer = oDicContainerStack(oDicContainerStack.Count-1)
			MsgLog ("fijado contexto: " & m_CurrContainer.ID)
			oDicContainerStack.Remove(oDicContainerStack.Count-1)
		End If
	End Sub

	Public Property Get oCurrContainer ()
		If Not m_Logger Is Nothing Then
			Set oCurrContainer = m_Logger.oCurrContainer
		Else
			' LEGACY fallback
			Set oCurrContainer = m_CurrContainer
		End If
	End Property

	Public Property Set oCurrContainer (oContainer)
		If Not m_Logger Is Nothing Then
			' Delegar a m_Logger usando setContainer
			If Not oContainer Is Nothing Then
				m_Logger.setContainer oContainer.id
			End If
		Else
			' LEGACY fallback
			If oContainer Is m_CurrContainer Then
				MsgLog ("<span style=""background-color:red;color:white;"">se ha intentado fijar el mismo contenedor que estaba activo</span>")
				Exit Property
			End If
			Dim lastContainer
			Set lastContainer = Nothing
			if oDicContainerStack.Count > 0 Then Set lastContainer = oDicContainerStack(oDicContainerStack.Count-1)
			If m_CurrContainer Is Nothing And oDicContainerStack.Count = 0 Then
			ElseIf (m_CurrContainer Is lastContainer) Or (m_CurrContainer Is Nothing) Then
				Stop
			Else
				oDicContainerStack.Add oDicContainerStack.Count,m_CurrContainer
			End If
			Set m_CurrContainer = oContainer
			MsgLog ("fijado contexto: " & oContainer.ID)
		End If
	End Property

	Public Function existsContainer (strSpoilerID)
		If Not m_Logger Is Nothing Then
			existsContainer = m_Logger.existsContainer(strSpoilerID)
		Else
			' LEGACY fallback
			existsContainer = Not (IsNull(objDocument.getElementById(strSpoilerID)) Or IsEmpty(objDocument.getElementById(strSpoilerID)))
		End If
	End Function

	Public Function setContainer (strSpoilerID)
		If Not m_Logger Is Nothing Then
			Set setContainer = m_Logger.setContainer(strSpoilerID)
		Else
			' LEGACY fallback
			Select Case True
				Case IsEmpty (strSpoilerID)
					Set setContainer = m_Cuerpo
				Case IsNull(objDocument.getElementById(strSpoilerID)) Or IsEmpty(objDocument.getElementById(strSpoilerID))
					MsgLog ("<span style=""background-color:red;color:white;"">se ha intentado fijar un contenedor inexistente</span>")
					Set setContainer = Nothing
				Case Else
					Set setContainer = objDocument.getElementById(strSpoilerID)
			End Select
			If Not setContainer Is Nothing Then Set oCurrContainer = setContainer
		End If
	End Function
		
	' =============================================
	' [REFACTORIZADO] Spoiler - Delegado a m_Logger
	' =============================================
	Public Function Spoiler (bOpen,ByVal style,linea_texto, ByVal strSpoilerID,bNoWrap)
		If Not m_Logger Is Nothing Then
			Set Spoiler = m_Logger.Spoiler(bOpen, style, linea_texto, strSpoilerID, bNoWrap)
			Exit Function
		End If

		' LEGACY fallback
		If objExplorer Is Nothing Or m_CurrContainer Is Nothing Then
			Err.Raise 91, "HTMLWindow", "InternetExplorer.Application not initialized. Call CreaVentanaHTML() first."
		End If
        Dim outerDiv, button, innerDiv
        
        If IsEmpty (strSpoilerID) Then strSpoilerID = "id" & unnamedSPCount : unnamedSPCount = unnamedSPCount + 1
        If Not (IsNull(objDocument.getElementById(strSpoilerID)) Or IsEmpty(objDocument.getElementById(strSpoilerID))) Then
        	' quiza seria mas correcto poner un mensaje de error: para ir a un spoiler dado, deberia usar setContainer
        	Set m_CurrContainer = objDocument.getElementById(strSpoilerID)
			MsgLog ("fijado contexto: " & m_CurrContainer.ID)
        	Exit Function
        End If
        
        If bOpen Then
        	If Not bNoWrap then
	            ' Div contenedor clickable
	            Set outerDiv = objDocument.createElement("div")
				outerDiv.className = "spoilerHead"
	            outerDiv.style.display = "block"
	            outerDiv.setAttribute "onclick", "showhide(this.nextSibling);this.firstChild.value = (this.firstChild.value == '+')? '-':'+';"
	
	            ' Botón [+/-]
	            Set button = objDocument.createElement("input")
	            button.Type = "button"
	            button.Value = "-"
	            button.style.cssText = "width=18;border-style:solid;border-width:1px;padding:0;"
	
	            outerDiv.appendChild button
	            outerDiv.appendChild objDocument.createTextNode(linea_texto)

	            ' Insertar 
	            m_CurrContainer.appendChild outerDiv
	        End If

            ' Div plegable
            Set innerDiv = objDocument.createElement("div")
            innerDiv.ID = strSpoilerID
            innerDiv.title = strSpoilerID
            innerDiv.className = "spoilerBody"
			innerDiv.style.csstext = style
			innerDiv.style.display = "block"
            If Not bNoWrap Then innerDiv.setAttribute "ondblclick", "showhide(this); this.previousSibling.firstChild.value = (this.previousSibling.firstChild.value == '+')? '-':'+';"

            Select Case True
            	Case m_CurrContainer Is Nothing, IsNull(m_CurrContainer.getAttribute("ID")) Or IsEmpty(m_CurrContainer.getAttribute("ID"))
            	Case Else
            		On Error Resume Next
					innerDiv.setAttribute "data-parent", m_CurrContainer.getAttribute("ID")
					On Error GoTo 0
            End Select
			
            If bNoWrap Then innerDiv.appendChild objDocument.createTextNode(linea_texto)
            
            ' Insertar 
            m_CurrContainer.appendChild innerDiv

            ' Guardamos referencia para futuros MsgIE
            Set oCurrContainer = innerDiv
        Else
            ' Cerramos contenedor
            closeSpoiler
        End If
            
    	Set Spoiler = m_CurrContainer
	End Function

	' =============================================
	' [REFACTORIZADO] closeSpoiler - Delegado a m_Logger
	' =============================================
	Public Function closeSpoiler ()
		If Not m_Logger Is Nothing Then
			Set closeSpoiler = m_Logger.closeSpoiler()
			Exit Function
		End If

		' LEGACY fallback - resuelve si el cierre de contexto es contra el "parent", o el m_Cuerpo
        Select Case True
        	Case IsEmpty(m_CurrContainer),m_CurrContainer Is Nothing, IsNull(objDocument.getElementById(m_CurrContainer.getAttribute("data-parent")) Or IsEmpty(objDocument.getElementById(m_CurrContainer.getAttribute("data-parent")))),  _
					m_CurrContainer.getAttribute("data-parent") = ""
	            Set oCurrContainer = m_Cuerpo
        	Case Else
            	Set oCurrContainer = objDocument.getElementById(m_CurrContainer.getAttribute("data-parent"))
        End Select
        Set closeSpoiler = m_CurrContainer
	End Function

	' =============================================
	' [REFACTORIZADO] AddTableRow - Delegado a m_Logger
	' =============================================
	Sub AddTableRow(oContainer,cell1,cell2)
		If Not m_Logger Is Nothing Then
			m_Logger.AddTableRow oContainer, cell1, cell2
			Exit Sub
		End If

		' LEGACY fallback
	    Dim tabla,child
	    For Each child In oContainer.getElementsByTagName("TABLE")
            Set tabla = child
	    Next
	    
	    If IsEmpty(tabla) Then
	        oContainer.insertAdjacentHTML "beforeEnd", _
	            "<table class='bordered-table'><tbody></tbody></table>"
	        Set tabla = oContainer.lastChild
	    End If
	    
	    tabla.getElementsByTagName("TBODY")(0).insertAdjacentHTML "beforeEnd", _
	        "<tr><td class='fixed-cell'>" & cell1 & "</td><td class='fluid-cell'>" & cell2 & "</td></tr>"
	End Sub
		
	Private Function fixHTML (ByVal strout)
		Dim regex,bNL
        Set regex = New RegExp : regex.Global = True : regex.IgnoreCase = True : regex.Multiline = False
		regex.Pattern = "<[^>]+>"
		bNL = regex.Replace(strout,"") <> "" ' solo hay cadenas de formato, NO se pone <br/>
		regex.Pattern = "</?(?:li|div|ul|tr|h\d)\b"
		bNL = bNL And Not regex.Test (strout) ' hay tags html de bloque, que introducen su br
		regex.Pattern = "(?:<br/?>)?(?:\r\n)?$"
		If bNL Then
			'strout = strout & "<br/>" & vbCrLf
			' opto por eliminar los br, y poner marcas de parrafo (incluso sin crlf: si quiero formatear, uso npp)
			strout = regex.Replace(strout,"<br/>" & vbCrLf)
			strout = regex.Replace(strout,"")
			strout = "<p>" & strout & "</p>"
		End If
		strout = Replace(strout,vbTab,"&emsp;")
		fixHTML = strout
	End Function

	' =============================================
	' [REFACTORIZADO] MsgIE - Delegado a m_Logger
	' =============================================
	Public Default Sub MsgIE (ByVal strout)
		If Not m_Logger Is Nothing Then
			m_Logger.MsgIE strout
		Else
			' LEGACY fallback
			If objExplorer Is Nothing Then
				Err.Raise 91, "HTMLWindow", "InternetExplorer.Application not initialized. Call CreaVentanaHTML() first."
			End If
			If Not m_CurrContainer Is Nothing Then
				m_CurrContainer.insertAdjacentHTML "beforeEnd", fixHTML(strout) & vbCrLf
			End If
		End If
	End Sub

	' =============================================
	' [REFACTORIZADO] MsgLog - Delegado a m_Logger
	' =============================================
	Public Sub MsgLog (linea_texto)
		If Not m_Logger Is Nothing Then
			m_Logger.MsgLog linea_texto
		Else
			' LEGACY fallback
			m_Log.insertAdjacentHTML "beforeEnd", fixHTML("[t:" & Round(Timer - fLogTick,3) & " seg.] <i>" & linea_texto & "</i>")
			m_Log.scrollTop = m_Log.scrollHeight  - m_Log.clientHeight
		End If
	End Sub

	' =============================================
	' [NUEVO] Métodos específicos de HtaLogger no presentes en HTMLWindow original
	' =============================================
	Public Sub MsgMain(sLevel, sText)
		If Not m_Logger Is Nothing Then m_Logger.MsgMain sLevel, sText
	End Sub

	Public Sub MsgDoc(sLevel, sText)
		If Not m_Logger Is Nothing Then m_Logger.MsgDoc sLevel, sText
	End Sub

	Public Sub LogHeader(sText)
		If Not m_Logger Is Nothing Then m_Logger.LogHeader sText
	End Sub

	Public Sub ResetBodyPanes()
		If Not m_Logger Is Nothing Then m_Logger.ResetBodyPanes
	End Sub

	Public Sub ResetPane(oPane)
		If Not m_Logger Is Nothing Then m_Logger.ResetPane oPane
	End Sub

	Public Function MsgTree(parentID, sText, sKey, sTooltip)
		If Not m_Logger Is Nothing Then
			Set MsgTree = m_Logger.MsgTree(parentID, sText, sKey, sTooltip)
		Else
			Set MsgTree = Nothing
		End If
	End Function

	Public Function HtaConfirm(prompt)
		If Not m_Logger Is Nothing Then
			HtaConfirm = m_Logger.HtaConfirm(prompt)
		Else
			HtaConfirm = MsgBox(prompt, vbYesNo + vbQuestion, "Confirmación")
		End If
	End Function

	Public Sub HtaAccept(prompt)
		If Not m_Logger Is Nothing Then
			m_Logger.HtaAccept prompt
		Else
			MsgBox prompt, vbOKOnly, "Información"
		End If
	End Sub
End Class

Sub IE_onQuit()
	bQuit = True
'   Wscript.Quit
End Sub

Function MsgLog (strTxt)
	If Not bLog Then Exit Function
	If Not MsgIE is Nothing And Not IsEmpty (MsgIE) Then MsgIE.MsgLog (strTxt) : Exit Function
	WScript.Echo strTxt
End Function

Function SeleccCarpetasWSArgs (strInfoMsg,bOnlyParamsCmdLine, strStartFolder)
	' Si no hay carpetas como parámetros de la linea de comandos, presenta
	' un cuadro de selección.
	Dim BIF_USENEWUI,BIF_RETURNONLYFSDIRS,BIF_VALIDATE,CSIDL_DRIVES
	Dim strArg,colCarpetaSelecc,oShellApp, oDicSal
	Set oDicSal = CreateObject("scripting.dictionary")
	Dim fso,ofin,strtmp
	Set fso = CreateObject("scripting.filesystemobject")
	If Not IsEmpty (strInfoMsg) Then MsgIE (strInfoMsg)
	For Each strArg In WScript.Arguments
		If fso.FolderExists (strArg) Then
			If Not oDicSal.Exists (strArg) Then oDicSal.Add strArg,Empty
			' MsgIE ("procesando la siguiente carpeta:<b>" & strArg & "</b>")
		ElseIf LCase(Right (strArg,4)) = ".txt" And fso.FileExists (strArg) Then
			' MsgIE ("procesando fichero de texto con carpetas:<b>" & strArg & "</b>")
			Set ofin = fso.OpenTextFile (strArg)
			While Not ofin.AtEndOfStream
				strtmp = Split(ofin.ReadLine,vbTab)(0)
				If Left (strtmp,1) = """" And Right (strtmp,1) = """" Then strtmp = Mid (strtmp,2,Len(strtmp)-2)
				If fso.FolderExists (strtmp) then
					If Not oDicSal.Exists (strtmp) Then oDicSal.Add strtmp,Empty
				End if
			Wend
			ofin.Close
		End If
	Next
	If Not bOnlyParamsCmdLine Then
		BIF_USENEWUI = &H40
		BIF_RETURNONLYFSDIRS = &H1
		BIF_VALIDATE = &H20
		CSIDL_DRIVES = &H11 ' Carpeta de inicio
		Set oShellApp = CreateObject("Shell.Application")
		If IsEmpty (strStartFolder) Then strStartFolder = CSIDL_DRIVES
		Do
			Set colCarpetaSelecc = oShellApp.BrowseForFolder (0, "Carpeta(s) en las que estan la subcarpetas a procesar" & vbCr & "(CANCELAR PARA FINALIZAR)", BIF_USENEWUI + BIF_RETURNONLYFSDIRS + BIF_VALIDATE + &H0010,strStartFolder)
			If colCarpetaSelecc Is Nothing Then
			Else
				If Not oDicSal.Exists (colCarpetaSelecc.Items.Item.Path) Then oDicSal.Add colCarpetaSelecc.Items.Item.Path,Empty
				'MsgIE (colCarpetaSelecc.Items.Item.Path)
			End If
		Loop Until colCarpetaSelecc Is Nothing
	End If
	SeleccCarpetasWSArgs = oDicSal.Keys
	Set oDicSal = Nothing
	Set fso = Nothing
	Set oShellApp = Nothing 
End Function


' --- cWScript Class (Argument parsing and system functions) ---
Class cWScript
	Private objArgs
	Private Sub Class_Initialize()
		Set objArgs = CreateObject("Scripting.Dictionary")
		On Error Resume Next
		Dim htaObj, fullCmd, regEx, matches, match, i
		Set htaObj = document.getElementById("oHTA")
		If Not htaObj Is Nothing Then
			fullCmd = htaObj.commandLine
			Set regEx = CreateObject("VBScript.RegExp")
			regEx.Global = True
			regEx.Pattern = """([^""]+)""|([^\s""]+)"
			Set matches = regEx.Execute(fullCmd)
			For i = 1 To matches.Count - 1
				Set match = matches.Item(i)
				If match.SubMatches(0) <> "" Then
					objArgs.Add i - 1, match.SubMatches(0)
				Else
					objArgs.Add i - 1, match.SubMatches(1)
				End If
			Next
		End If
		On Error GoTo 0
	End Sub
	Public Property Get Arguments()
		Set Arguments = objArgs
	End Property
	Sub Quit()
		Self.close()
	End Sub
	Sub Echo(str)
		HtaConfirm str ' For compatibility
	End Sub
	Sub Sleep(ms)
		Dim segundos : segundos = Int(ms / 1000)
		If segundos <= 0 Then segundos = 1
		CreateObject("WScript.Shell").Run "ping 127.0.0.1 -n " & (segundos + 1), 0, True
	End Sub
	Sub Sleep_alt(ms)
		Dim i, dtEnd
		dtEnd = DateAdd("s", ms/1000, Now())
		Do While Now() < dtEnd
			' Non-blocking sleep, placeholder. For real sleep, WScript.Shell is needed.
		Loop
	End Sub
End Class
