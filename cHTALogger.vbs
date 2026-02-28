
' --- HtaLogger Class (Restored & Enhanced) ---
' --- Compatibility Class for included VBS files ---
' This class mimics and extends the old cMsgIE so we don't have to rewrite the included files.
Class HtaLogger
	' TODO: devolver a private los atributos, de momento public para Debugging.
	Private m_document  ' Document object del contenedor (HTA, IE, o Nothing para VBScript puro)
	Private m_stack
	Private m_main, m_doc, m_log, m_header
	Private m_ModalOverlay, m_ModalDialog, m_ModalPrompt, m_ModalYesNo, m_ModalOk, m_ModalResponse
	Private unnamedSPCount
	Private m_bUseNativeDialogs  ' True = usar window.confirm/alert (IE), False = usar HTML custom (HTA)
	Private m_HostContext  ' "hta", "ie", "vbscript" (sin UI)

	Private Sub Class_Initialize()
		' Inicialización mínima - el document se inyectará después con SetDocument
		unnamedSPCount = 0
		Set m_stack = CreateObject("System.Collections.Stack")
		m_ModalResponse = -1
		m_bUseNativeDialogs = False
		m_HostContext = "unknown"
		Set m_document = Nothing
	End Sub

	' =============================================
	' Inyecta el document del contenedor (HTA o IE.Application)
	' @param oDocument - Document object del host (o Nothing para VBScript puro)
	' @param bUseNativeDialogs - True para usar window.confirm/alert (IE), False para HTML custom (HTA)
	' =============================================
	Public Sub SetDocument(oDocument, bUseNativeDialogs)
		If IsEmpty(bUseNativeDialogs) Then bUseNativeDialogs = False

		Set m_document = oDocument
		m_bUseNativeDialogs = bUseNativeDialogs

		' Detectar contexto de ejecución
		If m_document Is Nothing Then
			m_HostContext = "vbscript"  ' VBScript puro sin UI
			Exit Sub
		End If

		' Determinar si es HTA o IE basándonos en bUseNativeDialogs
		If bUseNativeDialogs Then
			m_HostContext = "ie"  ' IE.Application con diálogos nativos
		Else
			m_HostContext = "hta"  ' HTA con diálogos HTML custom
		End If

		' Inicializar elementos del DOM
		On Error Resume Next
		Set m_header = m_document.getElementById("header")
		Set m_main = m_document.getElementById("main")
		Set m_doc = m_document.getElementById("doc")
		Set m_log = m_document.getElementById("log")

		' Validar que se obtuvieron los elementos críticos
		If m_main Is Nothing Then
			Err.Raise 91, "HtaLogger", "No se encontró elemento 'main' en el document"
		End If

		' Default root container is Main
		m_stack.Push m_main

		' Asignar elementos de diálogos modales (solo para HTA)
		If m_HostContext = "hta" Then
			Set m_ModalOverlay = m_document.getElementById("modal-overlay")
			Set m_ModalDialog = m_document.getElementById("modal-dialog")
			Set m_ModalPrompt = m_document.getElementById("modal-prompt")
			Set m_ModalYesNo = m_document.getElementById("modal-yesno-buttons")
			Set m_ModalOk = m_document.getElementById("modal-ok-button")
		End If

		On Error GoTo 0
	End Sub

	' =============================================
	' Valida que HtaLogger está inicializado
	' =============================================
	Private Sub ValidateInitialized()
		If m_document Is Nothing And m_HostContext <> "vbscript" Then
			Err.Raise 91, "HtaLogger", "No inicializado. Llamar SetDocument() primero."
		End If
	End Sub

	' =============================================
	' Acceso al document (para clases auxiliares como cOpContainers)
	' =============================================
	Public Property Get Document()
		Set Document = m_document
	End Property

	Public Default Sub MsgIE(linea_texto)
		'MsgDoc "info", linea_texto
		' Writes to the current container (top of stack)
		WriteToContainer oCurrContainer, "<div>" & linea_texto & "</div>"
	End Sub
	
	' Logs a simple activity message to the bottom log panel
	Public Sub MsgLog(linea_texto)
		Log(linea_texto)
	End Sub

		' Spoiler: RESTAURADO a versión original HTMLWindow (DOM nativo)
	Public Function Spoiler(bOpen, ByVal style, linea_texto, ByVal strSpoilerID, bNoWrap)
	    Dim outerDiv, button, innerDiv

	    ' Generar ID único si no se proporciona (patrón original: "id" & N)
	    If IsEmpty(strSpoilerID) Or strSpoilerID = "" Then
	        strSpoilerID = "id" & unnamedSPCount
	        unnamedSPCount = unnamedSPCount + 1
	    End If

	    ' Verificar si el spoiler ya existe
	    Set innerDiv = Nothing
	    On Error Resume Next
	    Set innerDiv = m_document.getElementById(strSpoilerID)
	    On Error GoTo 0

	    If Not innerDiv Is Nothing Then
	        ' El spoiler ya existe, establecerlo como contenedor actual
	        m_stack.Push innerDiv
	        Set Spoiler = innerDiv
	        MsgLog "Reutilizando spoiler existente: " & strSpoilerID
	        Exit Function
	    End If

	    If bOpen Then
	        ' --- ABRIR SPOILER ---
	        If Not bNoWrap Then
	            ' Div contenedor clickable (cabecera)
	            Set outerDiv = m_document.createElement("div")
	            outerDiv.className = "spoilerHead"
	            outerDiv.style.display = "block"
	            outerDiv.setAttribute "onclick", "showhide(this.nextSibling);this.firstChild.value = (this.firstChild.value == '+')? '-':'+';"

	            ' Botón [+/-]
	            Set button = m_document.createElement("input")
	            button.Type = "button"
	            button.Value = "-"
	            button.style.cssText = "width=18;border-style:solid;border-width:1px;padding:0;"

	            outerDiv.appendChild button
	            outerDiv.appendChild m_document.createTextNode(linea_texto)

	            ' Insertar cabecera
	            oCurrContainer.appendChild outerDiv
	        End If

	        ' Div plegable (contenido)
	        Set innerDiv = m_document.createElement("div")
	        innerDiv.id = strSpoilerID
	        innerDiv.title = strSpoilerID
	        innerDiv.className = "spoilerBody"
	        innerDiv.style.cssText = style
	        innerDiv.style.display = "block"
	        If Not bNoWrap Then
	            innerDiv.setAttribute "ondblclick", "showhide(this); this.previousSibling.firstChild.value = (this.previousSibling.firstChild.value == '+')? '-':'+';"
	        End If

	        ' Guardar referencia al padre para poder cerrar después
	        On Error Resume Next
	        If Not oCurrContainer Is Nothing Then
	            If Not (IsNull(oCurrContainer.getAttribute("ID")) Or IsEmpty(oCurrContainer.getAttribute("ID"))) Then
	                innerDiv.setAttribute "data-parent", oCurrContainer.getAttribute("ID")
	            End If
	        End If
	        On Error GoTo 0

	        ' Si bNoWrap, el texto va dentro del div plegable
	        If bNoWrap Then innerDiv.appendChild m_document.createTextNode(linea_texto)

	        ' Insertar contenido
	        oCurrContainer.appendChild innerDiv

	        ' Establecer como contenedor actual
	        m_stack.Push innerDiv
	        Set Spoiler = innerDiv
	    Else
	        ' --- CERRAR SPOILER ---
	        Call closeSpoiler()
	        Set Spoiler = oCurrContainer
	    End If
	End Function
	
	' Cierra el spoiler actual y restaura el contenedor padre.
	' @return Referencia al contenedor padre
	Public Function closeSpoiler()
	    Dim parentId, parentContainer

	    ' Inicializar variables
	    Set parentContainer = Nothing

	    ' Intentar obtener el contenedor padre desde data-parent
	    On Error Resume Next
	    If Not oCurrContainer Is Nothing Then
	        parentId = oCurrContainer.getAttribute("data-parent")
	        If Not IsEmpty(parentId) And parentId <> "" Then
	            Set parentContainer = m_document.getElementById(parentId)
	        End If
	    End If
	    On Error GoTo 0

	    ' Si no hay padre válido, volver al contenedor principal
	    If parentContainer Is Nothing Then
	        Set parentContainer = m_main
	    End If

	    ' Hacer pop del stack si hay más de un elemento
	    If m_stack.Count > 1 Then
	        m_stack.Pop
	    End If
	    
	    ' Establecer el contenedor padre como actual
	    If m_stack.Count > 0 Then
	        Set closeSpoiler = m_stack.Peek()
	    Else
	        m_stack.Push parentContainer
	        Set closeSpoiler = parentContainer
	    End If
	End Function

	' These methods are kept for compatibility but don't need complex logic anymore.
	Public Sub popContainer()
		If m_stack.Count > 1 Then m_stack.Pop
	End Sub
	
	Public Property Get oCurrContainer()
		If m_stack.Count > 0 Then
			Set oCurrContainer = m_stack.Peek()
		Else
			Set oCurrContainer = m_main
		End If
	End Property
	
	Public Function existsContainer(sID)
		existsContainer = Not (m_document.getElementById(sID) Is Nothing)
	End Function

	Public Function setContainer (strSpoilerID)
		Dim el
		Set el = Nothing
		On Error Resume Next
		Set el = m_document.getElementById(strSpoilerID)
		On Error GoTo 0
		If Not el Is Nothing Then
			m_stack.Push el
			Set setContainer = el
		Else
			Set setContainer = oCurrContainer
		End If
	End Function
	
	Public Sub AddTableRow(oContainer, cell1, cell2)
	    Dim tabla, child

	    ' Validar que el contenedor es válido
	    If oContainer Is Nothing Then
	        MsgLog "Error: AddTableRow recibió un contenedor Nothing"
	        Exit Sub
	    End If

	    ' Buscar la última tabla existente en el contenedor
	    For Each child In oContainer.getElementsByTagName("TABLE")
	        Set tabla = child  ' Se queda con la última tabla del bucle
	    Next
	    
	    ' Si no hay tabla, crear una nueva
	    If tabla Is Nothing Then
	        oContainer.insertAdjacentHTML "beforeEnd", _
	            "<table class='bordered-table'><tbody></tbody></table>"
	        Set tabla = oContainer.lastChild
	    End If
	    
	    ' Verificar que obtuvimos la tabla correctamente
	    If tabla Is Nothing Then
	        MsgLog "Error: No se pudo crear/obtener la tabla en AddTableRow"
	        Exit Sub
	    End If
	    
	    ' Insertar la fila en el tbody (usa insertAdjacentHTML para no destruir DOM)
	    On Error Resume Next
	    tabla.getElementsByTagName("TBODY")(0).insertAdjacentHTML "beforeEnd", _
	        "<tr><td class='fixed-cell'>" & cell1 & "</td><td class='fluid-cell'>" & cell2 & "</td></tr>"
	    
	    If Err.Number <> 0 Then
	        MsgLog "Error al insertar fila en tabla: " & Err.Description & " (Container: " & oContainer.id & ")"
	        Err.Clear
	    End If
	    On Error GoTo 0
	End Sub	
	' --- New Specialized Logging Methods ---
	
	Public Sub ResetPane(oPane)
		if TypeName(oPane) = "String" then Set oPane = m_document.getElementById(oPane)
		If Not oPane Is Nothing Then oPane.innerHTML = ""
	End Sub

	Public Sub ResetBodyPanes()
		' No reseteamos el header aqui
		ResetPane m_main 
		ResetPane m_doc
		ResetPane m_log
		If Not m_main Is Nothing Then m_main.style.backgroundColor = "#f8f8f8"
		If Not m_doc Is Nothing Then m_doc.style.backgroundColor = "#f8f8f8"
	End Sub

	' Displays a message in the top header bar
	' REPLACES PREVIOUS CONTENT!!
	Public Sub LogHeader(sText)
		m_header.innerHTML = sText
	End Sub
	
	Function FormatErrorLevel (oDestBlock, sText, sLevel)
		Dim sColor, sBgColor
		Select Case sLevel
			Case "error" : sColor = "#990000" : sBgColor = "#FFDDDD"
			Case "warning" : sColor = "#8B4513" : sBgColor = "#FFF8DC"
			Case "success" : sColor = "#006400" : sBgColor = "#DFF0D8"
			Case "important" : sColor = "#00008B" : sBgColor = "#D9EDF7"
			 ' "info":
			Case Else : sColor = "#333" : sBgColor = "transparent" ' or a light gray like #f9f9f9
		End Select
		
		' Change panel background based on the most severe message
		If sLevel = "error" Then oDestBlock.style.backgroundColor = "#FFF0F0"
		If sLevel = "success" And oDestBlock.style.backgroundColor <> "#FFF0F0" Then oDestBlock.style.backgroundColor = "#F0FFF0"

		FormatErrorLevel = "<div style='color:" & sColor & "; background-color:" & sBgColor & "; border-left: 4px solid " & sColor & _
				"; padding: 5px; margin-bottom: 4px;'>" & sText & "</div>"
	End Function
	
	' Logs a message to the main panel. Level can be "info", "success", "warning", "error", "important"
	Sub MsgMain (sLevel, sText)
		WriteToContainer m_main, FormatErrorLevel (m_main, sText, sLevel)
	End Sub

	' Logs a message to the document-specific panel. Levels are the same as MsgMain.

	Public Sub MsgDoc(sLevel, sText)
		WriteToContainer m_doc, FormatErrorLevel (m_doc, sText, sLevel)
	End Sub

	' Logs a simple activity message to the bottom log panel
	Sub Log(sText)
		Dim sTimestamp
		sTimestamp = "[" & FormatDateTime(Now, 4) & "] "
		WriteToContainer m_log, sTimestamp & sText & "<br>" & vbCrLf
	End Sub
	
	' Appends HTML to a given element and scrolls it to the bottom
	Private Sub WriteToContainer(oElem, sHTML)
		If Not oElem Is Nothing Then
			oElem.innerHTML = oElem.innerHTML & sHTML
			oElem.scrollTop = oElem.scrollHeight ' Auto-scroll
		End If
	End Sub
	
	Public Sub Write(sHTML)
		' Generic write goes to current container
		WriteToContainer oCurrContainer, sHTML
	End Sub

	' Adds a root node to the tree view for a given opportunity folder.
	' Returns the created node object.
	Function MsgTree(parentID, sText, sKey, sTooltip)
		On Error Resume Next
		Dim oNewLi, oNewSpan, oNewIcon, oWrapper, oParentSpan, oParentLi, oParentUl, oRootUl, oParentIcon

		' Inicializar objetos a Nothing
		Set oRootUl = Nothing
		Set oParentSpan = Nothing
		Set oParentLi = Nothing
		Set oParentUl = Nothing
		Set oParentIcon = Nothing

		' Crear los elementos del nodo
		Set oNewLi = m_document.createElement("LI")
		oNewLi.Title = sTooltip ' Añadimos el tooltip al elemento LI

		' Crear Wrapper (para hover y click conjunto)
		Set oWrapper = m_document.createElement("SPAN")
		oWrapper.className = "node-wrapper"

		' Crear Icono (Span)
		Set oNewIcon = m_document.createElement("SPAN")
		oNewIcon.id = "icon_" & sKey
		oNewIcon.className = "icon"
		'oNewIcon.innerText = ChrW(8226) ' Bullet point por defecto
		oWrapper.appendChild oNewIcon

		' Crear Texto (Span)
		Set oNewSpan = m_document.createElement("SPAN")
		oNewSpan.innerText = sText
		oNewSpan.id = sKey
		oNewSpan.className = "node-text"
		oWrapper.appendChild oNewSpan

		' Añadir wrapper al LI
		oNewLi.appendChild oWrapper

		If IsEmpty(parentID) Or parentID = "" Then
			' Añadir a la raíz
			Set oRootUl = m_document.getElementById("OpUL")
			If Not oRootUl Is Nothing Then
				oRootUl.appendChild oNewLi
				' Asignar icono de carpeta cerrada por defecto a los nodos raiz
				oNewIcon.innerText = ChrW(&HD83D) & ChrW(&HDCC1)
			End If
		Else
			' Añadir a un padre existente
			Set oParentSpan = m_document.getElementById(parentID)
			If Not oParentSpan Is Nothing Then
				' CORRECCIÓN: Subimos dos niveles para obtener el LI, no el SPAN wrapper.
				' oParentSpan (node-text) -> parentNode (node-wrapper) -> parentNode (LI)
				Set oParentLi = oParentSpan.parentNode.parentNode
				
				' Buscar si ya tiene una lista anidada (UL)
				Set oParentUl = Nothing
				Dim child
				For Each child In oParentLi.childNodes
					If UCase(child.nodeName) = "UL" Then
						Set oParentUl = child
						Exit For
					End If
				Next
				
				' Si no tiene hijos aun, convertirlo en carpeta
				If oParentUl Is Nothing Then
					Set oParentUl = m_document.createElement("UL")
					oParentUl.className = "nested"
					oParentLi.appendChild oParentUl

					' Cambiar el icono del padre a "Carpeta/Caja"
					Set oParentIcon = m_document.getElementById("icon_" & parentID)
					If Not oParentIcon Is Nothing Then
						oParentIcon.innerText = ChrW(&HD83D) & ChrW(&HDCC1) ' Closed folder
					End If
				End If
				
				oParentUl.appendChild oNewLi
			Else
				MsgLog "Error: No se encontró el nodo padre '" & parentID & "'"
			End If
		End If

		Set MsgTree = oNewLi

		If Err.Number <> 0 Then
			Log "Error al añadir nodo al TreeView: " & Err.Description
			Set MsgTree = Nothing
		End If
		On Error Goto 0
	End Function

	' =============================================
	' Diálogo modal de aceptación (solo OK)
	' Adapta su comportamiento según el contexto de ejecución
	' =============================================
	Sub HtaAccept(prompt)
		Select Case m_HostContext
			Case "hta"
				' HTA: usar diálogos HTML personalizados
				m_ModalResponse = -1
				m_ModalPrompt.innerHTML = prompt
				m_ModalOverlay.style.display = "block"
				m_ModalDialog.style.display = "block"
				m_ModalYesNo.style.display = "none"
				m_ModalOk.style.display = "block"

				Do While m_ModalResponse = -1
					WScript.Sleep 200
				Loop

			Case "ie"
				' IE.Application: usar window.alert nativo
				m_document.parentWindow.alert prompt

			Case "vbscript"
				' VBScript puro: usar MsgBox nativo
				MsgBox prompt, vbOKOnly, "Información"

			Case Else
				' Fallback: intentar MsgBox
				On Error Resume Next
				MsgBox prompt, vbOKOnly, "Información"
				On Error GoTo 0
		End Select
	End Sub

	' =============================================
	' Diálogo modal de confirmación (Sí/No)
	' Adapta su comportamiento según el contexto de ejecución
	' @return 6 (vbYes) o 7 (vbNo)
	' =============================================
	Function HtaConfirm(prompt)
		Select Case m_HostContext
			Case "hta"
				' HTA: usar diálogos HTML personalizados
				m_ModalResponse = -1
				m_ModalPrompt.innerHTML = prompt
				m_ModalOverlay.style.display = "block"
				m_ModalDialog.style.display = "block"
				m_ModalOk.style.display = "none"
				m_ModalYesNo.style.display = "block"

				Do While m_ModalResponse = -1
					WScript.Sleep 200
				Loop
				HtaConfirm = m_ModalResponse

			Case "ie"
				' IE.Application: usar window.confirm nativo
				' Mapear resultado: true -> 6 (vbYes), false -> 7 (vbNo)
				If m_document.parentWindow.confirm(prompt) Then
					HtaConfirm = 6  ' vbYes
				Else
					HtaConfirm = 7  ' vbNo
				End If

			Case "vbscript"
				' VBScript puro: usar MsgBox nativo con vbYesNo
				HtaConfirm = MsgBox(prompt, vbYesNo + vbQuestion, "Confirmación")

			Case Else
				' Fallback: intentar MsgBox
				On Error Resume Next
				HtaConfirm = MsgBox(prompt, vbYesNo + vbQuestion, "Confirmación")
				If Err.Number <> 0 Then HtaConfirm = 7  ' vbNo por defecto
				On Error GoTo 0
		End Select
	End Function
	
	Public Property Let ModalResponse(val)
		m_ModalResponse = val
	End Property
	
	Public Property Get ModalResponse
		ModalResponse = m_ModalResponse
	End Property
	
	Function HtaResetDialog(prompt)
		m_ModalOverlay.style.display = "none"
		m_ModalDialog.style.display = "none"
		m_ModalYesNo.style.display = "none"
		m_ModalOk.style.display = "none"
	End Function
End Class

' --- HTA UI Functions ---
' FUNCIONES GLOBALES DUPLICADAS POR CONVENIENCIA
Function HtaConfirm(prompt)
	HtaConfirm = MsgIE.HtaConfirm (prompt)
	Call MsgIE.MsgDoc("important", prompt)
End Function

Sub HtaAccept(prompt)
	MsgIE.HtaAccept (prompt)
	Call MsgIE.MsgDoc("important", prompt)
End Sub

Sub ModalYes_OnClick()
	MsgIE.ModalResponse = 6 ' vbYes
	MsgIE.HtaResetDialog
End Sub

Sub ModalNo_OnClick()
	MsgIE.ModalResponse = 7 ' vbNo
	MsgIE.HtaResetDialog
End Sub

Sub ModalOk_OnClick()
	MsgIE.ModalResponse = 1 ' vbOK
	MsgIE.HtaResetDialog
End Sub

Sub Log (text)
	MsgIE.Log (text)
End Sub

' --- Class cOpContainers ---
' Helper class to manage Opportunity Containers in the Main Panel
Class cOpContainers
    Private oHTALogger, m_folder, m_folderName, opID
    Private m_MainDiv,m_DocDiv,m_TreeLi
    Private m_dicCalcTecns
    
    Private Sub Class_Initialize
		Set m_dicCalcTecns = CreateObject("scripting.dictionary")
    End Sub
    
    Public Property Get MainDiv ()
    	Set MainDiv = m_MainDiv
    End Property
    
    Public Property Get DocDiv ()
    	Set DocDiv = m_DocDiv
    End Property
    
    Public Property Get TreeLi ()
    	Set TreeLi = m_TreeLi
    End Property
    
    Public Property Get Calc (strCalc)
    	' para acceder al bloque de una seccion dada, se usa Calc(strcalc)("main"), etc
    	Set Calc = m_dicCalcTecns(strCalc)
    End Property
    
    Public Property Get folderName ()
    	folderName = m_folderName
    End Property
    
    Public Property Get folder ()
    	folder = m_folder
    End Property
    
    Public Sub folderUpdate (folderName_, folder_)
    	' cuando se renombre una oportunidad
    	Set m_folder = folder_
    	Set m_folderName = folderName_
    	' TODO: cambiar tambien el tooltip del arbol, y cualquier otra referencia de carpeta
    End Sub
    
    Public Sub AddCalc (strCalc)
    	Dim DivElement
    	m_dicCalcTecns.Add strCalc, CreateObject("scripting.dictionary")
    	' primero el de la sección Main: <div id=main ...><div id=SER00012 class=main-calc-container...>
    	Set DivElement = Create(strCalc, "main","calc-container")
    	m_dicCalcTecns(strCalc).Add "main", DivElement
    	' el de la sección Doc: <div id=doc ...><div id=SER00012 class=doc-calc-container...>
    	' CREO QUE ESTE NO TIENE SENTIDO, estoy mezclando "documentos" (ficheros) con OPORTUNIDAD...
    	Set DivElement = Create(strCalc, "doc","calc-container")
    	m_dicCalcTecns(strCalc).Add "doc", DivElement
    	' TODO: seguir aqui, si hace falta
    End Sub
    
    Public Function Init (oHTALogger_, opID_, folderName_, folder_)
    	Set Init = Me
    	opID = opID_
    	m_folderName = folderName_
    	m_folder = folder_
    	Set oHTALogger = oHTALogger_
    	' bloque general de la seccion main: <div id=main_opID class=main-op-container...>
    	Set m_MainDiv = Create(opID, "main","op-container")
     	' bloque general de la seccion doc: <div id=doc_opID class=doc-op-container...>
	   	Set m_DocDiv = Create(opID, "doc","op-container")
	   	' bloque del tree: PTE: formato??
		Set m_TreeLi = oHTALogger.MsgTree(Empty, folderName, opID, folder) ' Pasamos la ruta completa como tooltip
    End Function

    Private Function Create(itemID, strSection,classEndTag)
    	Dim DivElement, oParent
    	' itemID , se añade a strSection para definir el ID en la sección
    	' strSection = main, doc etc
    	' classEndTag = op-container, calc-container, etc
    	Set oParent = oHTALogger.Document.getElementById(strSection)
        Set DivElement = oHTALogger.Document.createElement("div")
        DivElement.id = strSection & "_" & itemID
        DivElement.className = strSection & "-" & classEndTag
        DivElement.style.display = "none"
        oParent.appendChild DivElement
        Set Create = DivElement
    End Function
    
    Public Sub Show()
        m_MainDiv.style.display = "block"
        m_DocDiv.style.display = "block"
    End Sub
    
    Public Sub Hide()
        m_MainDiv.style.display = "none"
        m_DocDiv.style.display = "none"
    End Sub
    
    Public Sub RemoveContainers()
    	RemoveContainer m_MainDiv
    	RemoveContainer m_DocDiv
    	RemoveContainer m_TreeItem
    End Sub
    
    Private Sub RemoveContainer(DivElement)
        On Error Resume Next
        If Not DivElement Is Nothing Then
            If Not DivElement.parentNode Is Nothing Then
                DivElement.parentNode.removeChild(DivElement)
            End If
            Set DivElement = Nothing
        End If
        On Error GoTo 0
    End Sub
    
    Private Sub Class_Terminate()
        RemoveContainers
    End Sub
End Class
